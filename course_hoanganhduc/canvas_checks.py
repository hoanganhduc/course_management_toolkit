# -*- coding: utf-8 -*-
# Canvas API: https://canvas.instructure.com/doc/api/

"""Canvas PDF checks and quality helpers."""

import base64
import copy
import difflib
import hashlib
import io
import json
import math
import os
import re
import shutil
import statistics
import tempfile
import time
from collections import Counter, OrderedDict, defaultdict
from datetime import datetime, timezone
from itertools import combinations

import cv2
import numpy as np
import pytesseract
import requests
from pdf2image import convert_from_path
from PIL import Image
import PyPDF2
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from tqdm import tqdm
from .canvas_auth import get_canvas_client

from .canvas_people import list_canvas_people
from .data import extract_text_from_scanned_pdf, refine_text_with_ai
from .settings import (
    ALL_AI_METHODS,
    CANVAS_LMS_API_KEY,
    CANVAS_LMS_API_URL,
    CANVAS_LMS_COURSE_ID,
    DEFAULT_AI_METHOD,
    DEFAULT_OCR_METHOD,
)
from .submission_checks import analyze_meaningfulness_in_folder

def compare_texts_from_pdfs_in_folder(
    folder_path,
    ocr_service=DEFAULT_OCR_METHOD,
    lang="auto",
    simple_text=False,
    refine=DEFAULT_AI_METHOD,
    similarity_threshold=0.85,
    db_path=None,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    auto_send=False,
    notify_students=True,
    verbose=False
):
    """
    Extract texts from all PDFs in a folder, compare the extracted texts,
    and output the names of the corresponding PDFs which have high similarity in contents.
    Also saves the comparison results to a TXT file in the same folder.
    If two or more PDFs are highly similar, send a message to all corresponding students
    asking them to resubmit, indicating that this is cheating and not allowed.

    Additionally, save the status of message sent for each pair to the TXT file.
    On later runs, do not send message for the same pair of PDFs again.

    Args:
        folder_path (str): Path to the folder containing PDF files.
        ocr_service (str): OCR service to use ("ocrspace", "tesseract", "paddleocr").
        lang (str): OCR language.
        simple_text (bool): If True, extract simple text.
        refine (str): AI refinement for generated messages ("gemini", "huggingface", or None).
        similarity_threshold (float): Threshold for considering two PDFs as similar (0-1).
        api_url (str): Canvas API base URL.
        api_key (str): Canvas API key.
        course_id (str): Canvas course ID.
        verbose (bool): If True, print more details; otherwise, print only important notice.

    Returns:
        List of tuples: [(pdf1, pdf2, similarity), ...] for pairs above threshold.
    """

    assignment_name_guess = os.path.basename(folder_path).replace("_", " ").strip()
    image_phash_threshold = 0.95
    image_ssim_threshold = 0.9
    layout_threshold = 0.9
    shingle_threshold = 0.6
    embedding_threshold = 0.8

    def _extract_metadata_from_filename(filename):
        # Expected pattern: <name>_<canvas_id>_<assignment_id>_<time>_<status>.pdf
        base = os.path.basename(filename)
        meta = {
            "file": base,
            "student_name": None,
            "canvas_id": None,
            "assignment_id": None,
            "submitted_at": None,
            "status": None,
        }
        match = re.match(r"^(?P<name>.+)_(?P<canvas_id>\\d+)_(?P<assignment_id>\\d+)_(?P<submitted>[^_]+)_(?P<status>[^_]+)\\.pdf$", base)
        if match:
            meta["student_name"] = match.group("name").replace("_", " ").strip()
            meta["canvas_id"] = match.group("canvas_id")
            meta["assignment_id"] = match.group("assignment_id")
            meta["submitted_at"] = match.group("submitted")
            meta["status"] = match.group("status")
            return meta
        # Fallback: try to infer a canvas id from numeric tokens
        parts = base.replace(".pdf", "").split("_")
        numeric = [p for p in parts if p.isdigit()]
        if numeric:
            meta["canvas_id"] = numeric[0]
        meta["student_name"] = parts[0].replace("_", " ").strip() if parts else None
        return meta

    def _file_md5(path, block_size=1 << 20):
        digest = hashlib.md5()
        try:
            with open(path, "rb") as f:
                while True:
                    data = f.read(block_size)
                    if not data:
                        break
                    digest.update(data)
            return digest.hexdigest()
        except OSError:
            return None

    def _phash(gray_image):
        # Perceptual hash via DCT on a 32x32 grayscale image.
        resized = cv2.resize(gray_image, (32, 32))
        dct = cv2.dct(resized.astype(np.float32))
        dct_low = dct[:8, :8].flatten()
        if len(dct_low) <= 1:
            return None
        median = np.median(dct_low[1:])
        bits = dct_low > median
        return "".join("1" if b else "0" for b in bits)

    def _phash_similarity(h1, h2):
        if not h1 or not h2 or len(h1) != len(h2):
            return None
        dist = sum(c1 != c2 for c1, c2 in zip(h1, h2))
        return 1.0 - (dist / float(len(h1)))

    def _ssim(img1, img2):
        # Basic SSIM on full-image statistics (fast, no sliding window).
        img1 = img1.astype(np.float64)
        img2 = img2.astype(np.float64)
        c1 = (0.01 * 255) ** 2
        c2 = (0.03 * 255) ** 2
        mu1 = img1.mean()
        mu2 = img2.mean()
        sigma1 = img1.var()
        sigma2 = img2.var()
        sigma12 = ((img1 - mu1) * (img2 - mu2)).mean()
        numerator = (2 * mu1 * mu2 + c1) * (2 * sigma12 + c2)
        denominator = (mu1 ** 2 + mu2 ** 2 + c1) * (sigma1 + sigma2 + c2)
        if denominator == 0:
            return 0.0
        return float(numerator / denominator)

    def _psnr(img1, img2):
        try:
            return float(cv2.PSNR(img1, img2))
        except Exception:
            mse = np.mean((img1.astype(np.float64) - img2.astype(np.float64)) ** 2)
            if mse == 0:
                return float("inf")
            return 20 * math.log10(255.0 / math.sqrt(mse))

    def _layout_signature(pil_image, grid_size=4):
        # Layout signature based on OCR bounding boxes bucketed into a grid.
        try:
            data = pytesseract.image_to_data(pil_image, output_type=pytesseract.Output.DICT)
        except Exception:
            return None
        w, h = pil_image.size
        if not w or not h:
            return None
        grid = np.zeros((grid_size, grid_size), dtype=np.float64)
        for left, top, width, height, text in zip(
            data.get("left", []),
            data.get("top", []),
            data.get("width", []),
            data.get("height", []),
            data.get("text", []),
        ):
            if not text or not str(text).strip():
                continue
            cx = left + width / 2.0
            cy = top + height / 2.0
            gx = min(grid_size - 1, max(0, int((cx / w) * grid_size)))
            gy = min(grid_size - 1, max(0, int((cy / h) * grid_size)))
            grid[gy, gx] += 1.0
        vec = grid.flatten()
        norm = np.linalg.norm(vec)
        if norm == 0:
            return None
        return vec / norm

    def _cosine_similarity_vector(vec1, vec2):
        if vec1 is None or vec2 is None:
            return None
        denom = (np.linalg.norm(vec1) * np.linalg.norm(vec2))
        if denom == 0:
            return None
        return float(np.dot(vec1, vec2) / denom)

    def _make_shingles(text, size=5):
        tokens = [t for t in text.split() if t]
        if len(tokens) < size:
            return set()
        return {" ".join(tokens[i:i + size]) for i in range(len(tokens) - size + 1)}

    pdf_files = sorted([
        os.path.join(folder_path, f)
        for f in os.listdir(folder_path)
        if f.lower().endswith(".pdf")
    ])
    if not pdf_files:
        if verbose:
            print("[PDFSimilarity] No PDF files found in the folder.")
        else:
            print("No PDF files found in the folder.")
        return []

    # Collect file metadata and extract texts for all PDFs
    file_metadata = {}
    extracted_texts = {}
    for pdf_path in tqdm(pdf_files, desc="Extracting texts from PDFs"):
        meta = _extract_metadata_from_filename(pdf_path)
        meta["size_bytes"] = os.path.getsize(pdf_path) if os.path.exists(pdf_path) else None
        meta["md5"] = _file_md5(pdf_path)
        try:
            reader = PyPDF2.PdfReader(pdf_path)
            meta["page_count"] = len(reader.pages)
            info = getattr(reader, "metadata", None) or {}
            meta["producer"] = info.get("/Producer") or info.get("Producer")
            meta["creator"] = info.get("/Creator") or info.get("Creator")
        except Exception:
            meta.setdefault("page_count", None)
        file_metadata[pdf_path] = meta
        base = os.path.splitext(pdf_path)[0]
        txt_path = base + f"_text_{ocr_service}.txt"
        if not os.path.exists(txt_path):
            txt_path = extract_text_from_scanned_pdf(
                pdf_path,
                txt_output_path=txt_path,
                service=ocr_service,
                lang=lang,
                simple_text=simple_text,
                verbose=verbose
            )
        if txt_path and os.path.exists(txt_path):
            with open(txt_path, "r", encoding="utf-8") as f:
                text = f.read()
            # Normalize text: remove whitespace, lowercase
            norm_text = re.sub(r"\s+", " ", text).strip().lower()
            extracted_texts[pdf_path] = norm_text
        else:
            extracted_texts[pdf_path] = ""

    # Prepare image-based features and layout signatures (first page only)
    image_features = {}
    for pdf_path in tqdm(pdf_files, desc="Extracting image features"):
        features = {}
        try:
            images = convert_from_path(pdf_path, first_page=1, last_page=1, dpi=100)
            if images:
                pil_image = images[0]
                gray = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2GRAY)
                resized = cv2.resize(gray, (256, 256))
                features["phash"] = _phash(gray)
                features["layout"] = _layout_signature(pil_image)
                features["image"] = resized
        except Exception:
            features = {}
        image_features[pdf_path] = features

    # Build shingle sets for text similarity
    shingle_sets = {pdf_path: _make_shingles(text) for pdf_path, text in extracted_texts.items()}

    # Compare all pairs using multiple similarity metrics for better accuracy

    similar_pairs = []
    all_pairs = []
    pdf_list = list(extracted_texts.keys())
    similarity_matrix = {}  # (pdf1, pdf2) -> ratio
    all_pairs_detail = []

    # Prepare TF-IDF vectors for all texts
    texts = [extracted_texts[p] for p in pdf_list]
    tfidf_matrix = None
    if any(texts):
        tfidf_vectorizer = TfidfVectorizer().fit(texts)
        tfidf_matrix = tfidf_vectorizer.transform(texts)

    # Optional sentence-embedding similarity (if sentence-transformers is installed)
    embedding_vectors = None
    embedding_method = None
    try:
        from sentence_transformers import SentenceTransformer

        model = SentenceTransformer("all-MiniLM-L6-v2")
        embedding_vectors = model.encode(texts, convert_to_numpy=True, normalize_embeddings=True)
        embedding_method = "sentence_transformers/all-MiniLM-L6-v2"
    except Exception:
        embedding_vectors = None

    # Helper to compute Jaccard similarity
    def jaccard_similarity(a, b):
        set_a = set(a.split())
        set_b = set(b.split())
        intersection = set_a & set_b
        union = set_a | set_b
        return len(intersection) / len(union) if union else 0.0

    # Helper to compute Euclidean distance similarity (1 / (1 + distance))
    def euclidean_similarity(vec1, vec2):
        dist = np.linalg.norm(vec1 - vec2)
        return 1.0 / (1.0 + dist)

    for i, pdf1 in enumerate(pdf_list):
        for j in range(i + 1, len(pdf_list)):
            pdf2 = pdf_list[j]
            text1 = extracted_texts[pdf1]
            text2 = extracted_texts[pdf2]

            meta1 = file_metadata.get(pdf1, {})
            meta2 = file_metadata.get(pdf2, {})
            exact_hash = meta1.get("md5") and meta1.get("md5") == meta2.get("md5")

            # Cosine similarity (TF-IDF)
            if tfidf_matrix is not None:
                cos_sim = cosine_similarity(tfidf_matrix[i], tfidf_matrix[j])[0, 0]
                euc_sim = euclidean_similarity(tfidf_matrix[i].toarray(), tfidf_matrix[j].toarray())
            else:
                cos_sim = 0.0
                euc_sim = 0.0

            # Jaccard similarity
            jac_sim = jaccard_similarity(text1, text2) if text1 and text2 else 0.0

            # SequenceMatcher similarity
            seq_sim = difflib.SequenceMatcher(None, text1, text2).ratio() if text1 and text2 else 0.0

            # Weighted average (can adjust weights as needed)
            ratio = (0.4 * cos_sim) + (0.25 * jac_sim) + (0.2 * seq_sim) + (0.15 * euc_sim)
            if exact_hash:
                ratio = 1.0

            # Image-based similarity (first page)
            img1 = image_features.get(pdf1, {}).get("image")
            img2 = image_features.get(pdf2, {}).get("image")
            phash_sim = _phash_similarity(
                image_features.get(pdf1, {}).get("phash"),
                image_features.get(pdf2, {}).get("phash")
            )
            ssim_value = _ssim(img1, img2) if img1 is not None and img2 is not None else None
            psnr_value = _psnr(img1, img2) if img1 is not None and img2 is not None else None

            # Layout-aware similarity
            layout_sim = _cosine_similarity_vector(
                image_features.get(pdf1, {}).get("layout"),
                image_features.get(pdf2, {}).get("layout")
            )

            # N-gram shingle similarity
            shingle_sim = 0.0
            shingles1 = shingle_sets.get(pdf1, set())
            shingles2 = shingle_sets.get(pdf2, set())
            if shingles1 and shingles2:
                shingle_sim = len(shingles1 & shingles2) / float(len(shingles1 | shingles2))

            # Embedding similarity (optional)
            embed_sim = None
            if embedding_vectors is not None:
                embed_sim = float(np.dot(embedding_vectors[i], embedding_vectors[j]))

            # Metadata match (does not trigger by itself)
            meta_match = False
            if meta1.get("producer") and meta2.get("producer") and meta1.get("producer") == meta2.get("producer"):
                if meta1.get("creator") == meta2.get("creator") and meta1.get("page_count") == meta2.get("page_count"):
                    meta_match = True

            all_pairs.append((os.path.basename(pdf1), os.path.basename(pdf2), ratio))
            similarity_matrix[(pdf1, pdf2)] = ratio
            similarity_matrix[(pdf2, pdf1)] = ratio
            text_flag = ratio >= similarity_threshold
            image_phash_flag = phash_sim is not None and phash_sim >= image_phash_threshold
            image_ssim_flag = ssim_value is not None and ssim_value >= image_ssim_threshold
            layout_flag = layout_sim is not None and layout_sim >= layout_threshold
            shingle_flag = shingle_sim >= shingle_threshold
            embedding_flag = embed_sim is not None and embed_sim >= embedding_threshold

            # Balanced rule: exact hash OR (text similarity + at least one image/layout signal).
            flagged = exact_hash or (text_flag and (image_phash_flag or image_ssim_flag or layout_flag))

            reasons = []
            if flagged:
                if exact_hash:
                    reasons.append("exact_hash")
                if text_flag:
                    reasons.append("text_similarity")
                if image_phash_flag:
                    reasons.append("image_phash")
                if image_ssim_flag:
                    reasons.append("image_ssim")
                if layout_flag:
                    reasons.append("layout_similarity")
                if shingle_flag:
                    reasons.append("ngram_shingles")
                if embedding_flag:
                    reasons.append("embedding_similarity")
                if meta_match:
                    reasons.append("metadata_match")
                similar_pairs.append((os.path.basename(pdf1), os.path.basename(pdf2), ratio))
            all_pairs_detail.append({
                "pdf1": os.path.basename(pdf1),
                "pdf2": os.path.basename(pdf2),
                "ratio": ratio,
                "metrics": {
                    "cosine": cos_sim,
                    "jaccard": jac_sim,
                    "sequence": seq_sim,
                    "euclidean": euc_sim,
                    "phash_similarity": phash_sim,
                    "ssim": ssim_value,
                    "psnr": psnr_value,
                    "layout_similarity": layout_sim,
                    "shingle_jaccard": shingle_sim,
                    "embedding_cosine": embed_sim,
                },
                "exact_hash": bool(exact_hash),
                "metadata_match": meta_match,
                "reasons": reasons,
            })

    # Save results to file in the same folder
    result_path = os.path.join(folder_path, "pdf_similarity_results.txt")
    status_path = os.path.join(folder_path, "pdf_similarity_status.json")
    report_path = os.path.join(folder_path, "pdf_similarity_report.json")

    # Load previous status if exists
    sent_status = {}
    if os.path.exists(status_path):
        try:
            with open(status_path, "r", encoding="utf-8") as f:
                sent_status = json.load(f)
        except Exception:
            sent_status = {}

    # Helper to create a unique key for a pair (order-independent)
    def pair_key(pdf1, pdf2):
        return "||".join(sorted([pdf1.lower(), pdf2.lower()]))

    # Save results to txt
    with open(result_path, "w", encoding="utf-8") as f:
        f.write("PDF similarity comparison results:\n")
        if all_pairs:
            for pdf1, pdf2, ratio in sorted(all_pairs, key=lambda x: -x[2]):
                mark = " <== HIGH SIMILARITY" if ratio >= similarity_threshold else ""
                key = pair_key(pdf1, pdf2)
                msg_status = sent_status.get(key, "NOT_SENT")
                f.write(f"{pdf1} <-> {pdf2}: similarity = {ratio:.2f}{mark} [Message: {msg_status}]\n")
        else:
            f.write("No PDF pairs to compare.\n")
        if similar_pairs:
            f.write("\nPDF pairs with high similarity (>= {:.2f}):\n".format(similarity_threshold))
            for pdf1, pdf2, ratio in sorted(similar_pairs, key=lambda x: -x[2]):
                key = pair_key(pdf1, pdf2)
                msg_status = sent_status.get(key, "NOT_SENT")
                f.write(f"{pdf1} <-> {pdf2}: similarity = {ratio:.2f} [Message: {msg_status}]\n")
        else:
            f.write("\nNo highly similar PDF pairs found.\n")
    if verbose:
        print(f"[PDFSimilarity] Comparison results saved to {result_path}")
    else:
        print(f"Comparison results saved to {result_path}")

    pair_details_by_key = {
        pair_key(p["pdf1"], p["pdf2"]): p for p in all_pairs_detail
    }

    # Cluster detection for groups of similar submissions
    def _build_clusters(details):
        parent = {}

        def find(x):
            parent.setdefault(x, x)
            if parent[x] != x:
                parent[x] = find(parent[x])
            return parent[x]

        def union(a, b):
            ra, rb = find(a), find(b)
            if ra != rb:
                parent[rb] = ra

        for detail in details:
            if detail.get("reasons"):
                union(detail["pdf1"], detail["pdf2"])

        clusters = {}
        for detail in details:
            for f in (detail["pdf1"], detail["pdf2"]):
                root = find(f)
                clusters.setdefault(root, set()).add(f)
        return [sorted(list(members)) for members in clusters.values() if len(members) > 1]

    clusters = _build_clusters(all_pairs_detail)

    # Save a structured report for downstream processing
    try:
        report_payload = {
            "generated_at": datetime.now().isoformat(),
            "threshold": similarity_threshold,
            "assignment_name": assignment_name_guess,
            "methods": {
                "text_similarity_threshold": similarity_threshold,
                "image_phash_threshold": image_phash_threshold,
                "image_ssim_threshold": image_ssim_threshold,
                "layout_threshold": layout_threshold,
                "shingle_threshold": shingle_threshold,
                "embedding_threshold": embedding_threshold,
                "embedding_method": embedding_method,
                "flag_rule": "exact_hash OR (text_similarity AND (image_phash OR image_ssim OR layout_similarity))",
            },
            "files": {os.path.basename(k): v for k, v in file_metadata.items()},
            "pairs": all_pairs_detail,
            "high_similarity": [
                p for p in all_pairs_detail
                if p["reasons"]
            ],
            "clusters": clusters,
        }
        with open(report_path, "w", encoding="utf-8") as f:
            json.dump(report_payload, f, ensure_ascii=False, indent=2)
        if verbose:
            print(f"[PDFSimilarity] Report saved to {report_path}")
    except Exception as e:
        if verbose:
            print(f"[PDFSimilarity] Failed to save report JSON: {e}")

    if similar_pairs:
        if verbose:
            print("[PDFSimilarity] PDF pairs with high similarity:")
            for pdf1, pdf2, ratio in sorted(similar_pairs, key=lambda x: -x[2]):
                key = pair_key(pdf1, pdf2)
                msg_status = sent_status.get(key, "NOT_SENT")
                print(f"  {pdf1} <-> {pdf2}: similarity = {ratio:.2f} [Message: {msg_status}]")
        else:
            print("PDF pairs with high similarity found.")
    else:
        if verbose:
            print("[PDFSimilarity] No highly similar PDF pairs found.")
        else:
            print("No highly similar PDF pairs found.")

    # Update local database with similarity flags if possible
    if db_path is None:
        db_path = get_default_db_path()
    if db_path and os.path.exists(db_path) and similar_pairs:
        try:
            students = load_database(db_path, verbose=verbose)
            sid_map = {}
            name_map = {}

            def _norm_name(value):
                return re.sub(r"\\s+", " ", str(value or "")).strip().lower()

            for s in students:
                canvas_id = getattr(s, "Canvas ID", None)
                if canvas_id is not None:
                    sid_map[str(canvas_id)] = s
                name = getattr(s, "Name", None)
                if name:
                    name_map[_norm_name(name)] = s

            def _resolve_student(meta):
                sid = str(meta.get("canvas_id") or "")
                if sid and sid in sid_map:
                    return sid_map[sid]
                name_key = _norm_name(meta.get("student_name"))
                return name_map.get(name_key)

            def _pair_key(a, b):
                return "||".join(sorted([a, b]))

            for pdf1, pdf2, ratio in similar_pairs:
                meta1 = file_metadata.get(os.path.join(folder_path, pdf1), file_metadata.get(pdf1, {}))
                meta2 = file_metadata.get(os.path.join(folder_path, pdf2), file_metadata.get(pdf2, {}))
                student1 = _resolve_student(meta1)
                student2 = _resolve_student(meta2)
                if not student1 and not student2:
                    continue
                pair_key = _pair_key(pdf1, pdf2)
                detail = pair_details_by_key.get(pair_key, {})
                reasons = detail.get("reasons", [])

                entry = {
                    "pair_key": pair_key,
                    "other_file": pdf2,
                    "similarity": round(ratio, 4),
                    "reasons": reasons,
                    "report_path": report_path,
                    "assignment_id": meta1.get("assignment_id") or meta2.get("assignment_id"),
                    "assignment_name": assignment_name_guess,
                }
                entry_other = {
                    "pair_key": pair_key,
                    "other_file": pdf1,
                    "similarity": round(ratio, 4),
                    "reasons": reasons,
                    "report_path": report_path,
                    "assignment_id": meta1.get("assignment_id") or meta2.get("assignment_id"),
                    "assignment_name": assignment_name_guess,
                }

                for student, payload in ((student1, entry), (student2, entry_other)):
                    if not student:
                        continue
                    existing = getattr(student, "Plagiarism Matches", [])
                    if not isinstance(existing, list):
                        existing = []
                    if not any(item.get("pair_key") == pair_key for item in existing if isinstance(item, dict)):
                        existing.append(payload)
                        setattr(student, "Plagiarism Matches", existing)

            save_database(students, db_path, verbose=verbose)
            if verbose:
                print(f"[PDFSimilarity] Saved plagiarism flags to database: {db_path}")
        except Exception as e:
            if verbose:
                print(f"[PDFSimilarity] Failed to update database: {e}")

    # Only send messages for pairs that have not been sent before
    pairs_to_notify = []
    for pdf1, pdf2, ratio in similar_pairs:
        key = pair_key(pdf1, pdf2)
        if sent_status.get(key, "NOT_SENT") != "SENT":
            pairs_to_notify.append((pdf1, pdf2, ratio))

    if not notify_students:
        try:
            with open(status_path, "w", encoding="utf-8") as f:
                json.dump(sent_status, f, ensure_ascii=False, indent=2)
            if verbose:
                print(f"[PDFSimilarity] Message status saved to {status_path}")
        except Exception as e:
            if verbose:
                print(f"[PDFSimilarity] Failed to save message status: {e}")
            else:
                print(f"Failed to save message status: {e}")
        return similar_pairs

    # If there are highly similar pairs to notify, send a separate message for each pair
    if pairs_to_notify:
        if verbose:
            print("[PDFSimilarity] Sending messages to students involved in newly detected highly similar submissions (one message per pair)...")
        else:
            print("Sending messages to students involved in highly similar submissions...")
        canvas = get_canvas_client(api_url, api_key)
        course = canvas.get_course(course_id)

        # Helper to extract Canvas ID from filename
        def extract_canvas_id_from_filename(filename):
            parts = filename.split('_')
            for part in parts:
                if part.isdigit():
                    return int(part)
            return None

        # Ask user if they want to refine the message via AI if refine is None
        if refine is None and not auto_send:
            def get_input_with_timeout(prompt, timeout=60, default=None):
                # Use signal.SIGALRM only if available (not on Windows)
                if hasattr(signal, "SIGALRM"):
                    signal.signal(signal.SIGALRM, timeout_handler)
                    signal.alarm(timeout)
                    try:
                        result = input(prompt)
                        signal.alarm(0)
                        if not result and default is not None:
                            return default
                        return result
                    except TimeoutError:
                        signal.alarm(0)
                        if default is not None:
                            return default
                        raise
                    except KeyboardInterrupt:
                        signal.alarm(0)
                        raise
                else:
                    # Fallback for platforms without SIGALRM (e.g., Windows)
                    try:
                        result = input(prompt)
                        if not result and default is not None:
                            return default
                        return result
                    except KeyboardInterrupt:
                        raise

            try:
                refine_choice = get_input_with_timeout(
                    "Do you want to refine the message via AI? (none/gemini/huggingface/local) [none]: ",
                    timeout=60,
                    default="none"
                ).strip().lower()
                if refine_choice in ("none", "gemini", "huggingface"):
                    refine = refine_choice if refine_choice != "none" else None
                else:
                    if verbose:
                        print("[PDFSimilarity] Invalid choice. Using default 'none'.")
                    else:
                        print("Invalid choice. Using default 'none'.")
                    refine = None
            except TimeoutError:
                if verbose:
                    print("[PDFSimilarity] No response after 60 seconds. Using default 'none'.")
                else:
                    print("No response after 60 seconds. Using default 'none'.")
                refine = None
        if refine is None and auto_send:
            refine = None

        for pdf1, pdf2, ratio in pairs_to_notify:
            canvas_id1 = extract_canvas_id_from_filename(pdf1)
            canvas_id2 = extract_canvas_id_from_filename(pdf2)
            recipients = []
            if canvas_id1:
                recipients.append(str(canvas_id1))
            if canvas_id2 and canvas_id2 != canvas_id1:
                recipients.append(str(canvas_id2))
            if not recipients:
                if verbose:
                    print(f"[PDFSimilarity] Could not extract Canvas IDs from {pdf1} and {pdf2}. Skipping message.")
                else:
                    print(f"Could not extract Canvas IDs from {pdf1} and {pdf2}. Skipping message.")
                sent_status[pair_key(pdf1, pdf2)] = "FAILED"
                continue

            detail = pair_details_by_key.get(pair_key(pdf1, pdf2), {})
            reasons = detail.get("reasons", [])
            metrics = detail.get("metrics", {})

            method_descriptions = {
                "exact_hash": "Exact file hash match",
                "text_similarity": "Text similarity (TF-IDF/sequence)",
                "image_phash": "Perceptual image hash match",
                "image_ssim": "Image structural similarity",
                "layout_similarity": "Layout similarity from OCR bounding boxes",
                "ngram_shingles": "N-gram shingle overlap",
                "embedding_similarity": "Embedding similarity (if available)",
                "metadata_match": "PDF metadata match (producer/creator/pages)",
            }
            methods_used = [method_descriptions.get(r, r) for r in reasons]
            if not methods_used:
                methods_used = ["Text similarity (TF-IDF/sequence)"]
            method_block = "\n".join([f"- {m}" for m in methods_used])

            # Generate message for this pair
            similarity_results = f"{pdf1} <-> {pdf2}: similarity = {ratio:.2f}"
            metrics_summary = (
                f"TF-IDF cosine: {metrics.get('cosine')}, "
                f"Jaccard: {metrics.get('jaccard')}, "
                f"Sequence: {metrics.get('sequence')}, "
                f"Image pHash: {metrics.get('phash_similarity')}, "
                f"SSIM: {metrics.get('ssim')}, "
                f"Layout: {metrics.get('layout_similarity')}, "
                f"Shingles: {metrics.get('shingle_jaccard')}, "
                f"Embedding: {metrics.get('embedding_cosine')}"
            )
            if refine in ALL_AI_METHODS:
                if verbose:
                    print(f"[PDFSimilarity] Generating message using AI service: {refine} for pair {pdf1} <-> {pdf2} ...")
                else:
                    print(f"Generating message using AI service: {refine} for pair {pdf1} <-> {pdf2} ...")
                prompt = (
                    "You are an expert assistant. Compose a clear, formal, and professional message in Vietnamese to notify students "
                    "about potential similarity detected in their submissions by the system. The message should include the following points:\n\n"
                    "1. Provide the similarity results, listing the pair of submissions with their similarity score.\n"
                    "2. Explain the methods used (text similarity, image similarity, layout similarity, n-gram overlap, metadata match, optional embedding similarity).\n"
                    "3. Emphasize that automated detection can produce false positives and the case will be reviewed by lecturers and TAs before any final decision.\n"
                    "4. Ask students to wait for review or respond if they believe the detection is incorrect.\n\n"
                    "Ensure the message is complete, concise, and does not require any additional edits or replacements.\n\n"
                    "Similarity result:\n{text}\n"
                    "Methods:\n{methods}\n"
                    "Metrics:\n{metrics}"
                )
                message = refine_text_with_ai(
                    similarity_results,
                    method=refine,
                    user_prompt=prompt.format(text=similarity_results, methods=method_block, metrics=metrics_summary)
                )
            else:
                message = (
                    "Potential similarity detected by automated checks for the following submissions:\n"
                    + similarity_results +
                    "\n\nAssignment: "
                    + assignment_name_guess +
                    "\n\nMethods used:\n"
                    + method_block +
                    "\n\nMetrics:\n"
                    + metrics_summary +
                    "\n\nNote: Automated detection can produce false positives (OCR errors, formatting differences, or similar templates). "
                    "This case will be reviewed by the lecturers and TAs before any final decision. "
                    "If you believe this detection is incorrect, you may respond with clarification."
                )

            subject = "Notice: Potential similarity detected in submissions"

            if verbose:
                print(f"[PDFSimilarity] Subject:\n{subject}")
                print(f"[PDFSimilarity] Message for {pdf1} <-> {pdf2}:\n{message}")
            else:
                print(f"Prepared message for {pdf1} <-> {pdf2}.")

            if not auto_send:
                while True:
                    try:
                        # Only use SIGALRM if available (not on Windows)
                        if hasattr(signal, "SIGALRM"):
                            signal.signal(signal.SIGALRM, timeout_handler)
                            signal.alarm(60)  # 60 second timeout
                            confirm = input(f"\nDo you want to send this message for {pdf1} <-> {pdf2}? (y/n, or 'r' to regenerate, default 'y' in 60s): ").strip().lower()
                            signal.alarm(0)  # Cancel the alarm
                        else:
                            # Fallback for Windows (no timeout)
                            confirm = input(f"\nDo you want to send this message for {pdf1} <-> {pdf2}? (y/n, or 'r' to regenerate): ").strip().lower()
                            if not confirm:
                                confirm = "y"  # Default value

                        if confirm == "y" or confirm == "":
                            break
                        elif confirm == "n":
                            if verbose:
                                print(f"[PDFSimilarity] Message sending canceled for this pair.")
                            else:
                                print("Message sending canceled for this pair.")
                            sent_status[pair_key(pdf1, pdf2)] = "SKIPPED"
                            break
                        elif confirm == "r":
                            if verbose:
                                print(f"[PDFSimilarity] Regenerating message...")
                            else:
                                print("Regenerating message...")
                            if refine in ALL_AI_METHODS:
                                message = refine_text_with_ai(similarity_results, method=refine, user_prompt=prompt)
                            else:
                                message = (
                                    "Potential similarity detected by automated checks for the following submissions:\n"
                                    + similarity_results +
                                    "\n\nAssignment: "
                                    + assignment_name_guess +
                                    "\n\nMethods used:\n"
                                    + method_block +
                                    "\n\nMetrics:\n"
                                    + metrics_summary +
                                    "\n\nNote: Automated detection can produce false positives (OCR errors, formatting differences, or similar templates). "
                                    "This case will be reviewed by the lecturers and TAs before any final decision. "
                                    "If you believe this detection is incorrect, you may respond with clarification."
                                )
                            if verbose:
                                print(f"[PDFSimilarity] Regenerated Message:\n{message}")
                            else:
                                print("Regenerated message.")
                        else:
                            if verbose:
                                print("[PDFSimilarity] Invalid input. Please enter 'y', 'n', or 'r'.")
                            else:
                                print("Invalid input. Please enter 'y', 'n', or 'r'.")
                    except TimeoutError:
                        if verbose:
                            print("[PDFSimilarity] No response after 60 seconds, using default 'y'.")
                        else:
                            print("No response after 60 seconds, using default 'y'.")
                        break  # Use default 'y' option and break the loop

            if sent_status.get(pair_key(pdf1, pdf2)) == "SKIPPED":
                continue

            try:
                canvas.create_conversation(
                    recipients=recipients,
                    subject=subject,
                    body=message,
                    force_new=True
                )
                if verbose:
                    print(f"[PDFSimilarity] Message sent successfully to students for {pdf1} <-> {pdf2}.")
                else:
                    print(f"Message sent for {pdf1} <-> {pdf2}.")
                sent_status[pair_key(pdf1, pdf2)] = "SENT"
            except Exception as e:
                if verbose:
                    print(f"[PDFSimilarity] Failed to send message for {pdf1} <-> {pdf2}: {e}")
                else:
                    print(f"Failed to send message for {pdf1} <-> {pdf2}: {e}")
                sent_status[pair_key(pdf1, pdf2)] = "FAILED"

        # Save updated status to JSON
        try:
            with open(status_path, "w", encoding="utf-8") as f:
                json.dump(sent_status, f, ensure_ascii=False, indent=2)
            if verbose:
                print(f"[PDFSimilarity] Message status saved to {status_path}")
        except Exception as e:
            if verbose:
                print(f"[PDFSimilarity] Failed to save message status: {e}")
            else:
                print(f"Failed to save message status: {e}")

    else:
        # Save status file even if nothing to send, to keep track
        try:
            with open(status_path, "w", encoding="utf-8") as f:
                json.dump(sent_status, f, ensure_ascii=False, indent=2)
            if verbose:
                print(f"[PDFSimilarity] Message status saved to {status_path}")
        except Exception as e:
            if verbose:
                print(f"[PDFSimilarity] Failed to save message status: {e}")
            else:
                print(f"Failed to save message status: {e}")

    return similar_pairs

def detect_meaningful_level_and_notify_students(
    folder_path,
    assignment_id=None,
    ocr_service=DEFAULT_OCR_METHOD,
    lang="auto",
    simple_text=False,
    refine=DEFAULT_AI_METHOD,
    meaningfulness_threshold=0.4,
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    auto_send=False,
    verbose=False
):
    """
    Detect the meaningful level of extracted texts from PDFs in a folder using AI agent.
    If the meaningful level is too low, generate a message via AI and send it to the student
    asking them to reformat and resubmit their submissions.

    Args:
        folder_path (str): Path to the folder containing PDF files.
        assignment_id (str): Canvas assignment ID for sending messages.
        ocr_service (str): OCR service to use ("ocrspace", "tesseract", "paddleocr").
        lang (str): OCR language.
        simple_text (bool): If True, extract simple text.
        refine (str): AI refinement ("gemini", "huggingface", or None).
        meaningfulness_threshold (float): Threshold for considering text as meaningful (0-1).
        api_url, api_key, course_id: Canvas API configuration.
        verbose (bool): If True, print more details; otherwise, print only important notice.

    Returns:
        Dict with results: {filename: {"meaningful_score": score, "message_sent": bool}, ...}
    """

    def get_input_with_timeout_default(prompt, timeout=60, default=None):
        # Use signal.SIGALRM only if available (not on Windows)
        if hasattr(signal, "SIGALRM"):
            try:
                signal.signal(signal.SIGALRM, timeout_handler)
                signal.alarm(timeout)
                result = input(prompt)
                signal.alarm(0)
                if not result and default is not None:
                    if verbose:
                        print(f"[Meaningfulness] Using default: {default}")
                    else:
                        print(f"Using default: {default}")
                    return default
                return result
            except TimeoutError:
                signal.alarm(0)
                if default is not None:
                    if verbose:
                        print(f"\n[Meaningfulness] No response after {timeout} seconds, using default '{default}'")
                    else:
                        print(f"\nNo response after {timeout} seconds, using default '{default}'")
                    return default
                raise
            except KeyboardInterrupt:
                signal.alarm(0)
                if verbose:
                    print("\n[Meaningfulness] Operation cancelled by user.")
                else:
                    print("\nOperation cancelled by user.")
                raise
        else:
            # Fallback for platforms without SIGALRM (e.g., Windows)
            try:
                result = input(prompt)
                if not result and default is not None:
                    if verbose:
                        print(f"[Meaningfulness] Using default: {default}")
                    else:
                        print(f"Using default: {default}")
                    return default
                return result
            except KeyboardInterrupt:
                if verbose:
                    print("\n[Meaningfulness] Operation cancelled by user.")
                else:
                    print("\nOperation cancelled by user.")
                raise

    pdf_files = sorted([
        os.path.join(folder_path, f)
        for f in os.listdir(folder_path)
        if f.lower().endswith(".pdf")
    ])
    
    if not pdf_files:
        if verbose:
            print("[Meaningfulness] No PDF files found in the folder.")
        else:
            print("No PDF files found in the folder.")
        return {}

    status_path = os.path.join(folder_path, "meaningfulness_status.json")
    if os.path.exists(status_path):
        try:
            with open(status_path, "r", encoding="utf-8") as f:
                sent_status = json.load(f)
        except Exception:
            sent_status = {}
    else:
        sent_status = {}

    if refine is None and not auto_send:
        refine = get_input_with_timeout_default(
            "Which AI model do you want to use for meaningfulness analysis? (gemini/huggingface/local, default 'gemini' in 60s): ",
            timeout=60,
            default="gemini"
        ).strip().lower()
        if refine not in ALL_AI_METHODS:
            refine = "gemini"
    elif refine is None and auto_send:
        refine = DEFAULT_AI_METHOD or None

    results, low_quality, extracted_texts, average_length = analyze_meaningfulness_in_folder(
        folder_path,
        ocr_service=ocr_service,
        lang=lang,
        meaningfulness_threshold=meaningfulness_threshold,
        refine_method=refine,
        return_texts=True,
        write_report=True,
        verbose=verbose,
    )
    low_quality_files = []
    for filename, result in results.items():
        already_sent = sent_status.get(filename, {}).get("message_sent", False)
        result["message_sent"] = already_sent
        sent_status[filename] = result
    for filename in low_quality:
        already_sent = results.get(filename, {}).get("message_sent", False)
        score = results.get(filename, {}).get("meaningful_score", 0.0)
        text = extracted_texts.get(filename, "")
        if not already_sent:
            low_quality_files.append((filename, score, text))
    
    result_path = os.path.join(folder_path, "meaningfulness_analysis.txt")
    with open(result_path, "w", encoding="utf-8") as f:
        f.write("PDF meaningfulness analysis results:\n")
        f.write(f"Threshold: {meaningfulness_threshold}\n")
        f.write(f"Average text length: {average_length:.0f} characters\n")
        f.write(f"Analysis date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        total_files = len(results)
        low_quality_count = sum(1 for v in results.values() if v["meaningful_score"] < meaningfulness_threshold)
        acceptable_count = total_files - low_quality_count
        f.write(f"Summary:\n")
        f.write(f"Total files analyzed: {total_files}\n")
        f.write(f"Acceptable quality: {acceptable_count}\n")
        f.write(f"Low quality: {low_quality_count}\n\n")
        f.write("Detailed results:\n")
        f.write("-" * 80 + "\n")
        for filename, result in sorted(results.items()):
            score = result["meaningful_score"]
            status = "LOW QUALITY" if score < meaningfulness_threshold else "ACCEPTABLE"
            error = result.get("error", "")
            text_length = result.get("text_length", 0)
            length_ratio = text_length / average_length if average_length > 0 else 0
            msg_status = "SENT" if result.get("message_sent") else "NOT_SENT"
            issues = result.get("issues", [])
            metrics = result.get("metrics", {})
            issues_text = "; ".join(issues) if issues else "None"
            f.write(
                f"{filename}: score = {score:.2f} ({status}), length = {text_length} chars ({length_ratio:.2f}x avg), message: {msg_status}"
            )
            if error:
                f.write(f" - ERROR: {error}")
            f.write(f"\n  Issues: {issues_text}\n")
            if metrics:
                # Keep metrics on one line to simplify manual scanning.
                f.write("  Metrics: ")
                f.write(
                    f"vn_ratio={metrics.get('vn_char_ratio', 0):.2f}, alnum={metrics.get('alnum_ratio', 0):.2f}, "
                    f"symbol={metrics.get('symbol_ratio', 0):.2f}, unique={metrics.get('unique_char_ratio', 0):.2f}, "
                    f"repeat={metrics.get('repeat_char_ratio', 0):.2f}, empty_lines={metrics.get('line_empty_ratio', 0):.2f}, "
                    f"likely_math={metrics.get('likely_math', False)}\n"
                )
        if low_quality_files:
            f.write(f"\nLow quality files requiring attention (< {meaningfulness_threshold}):\n")
            f.write("-" * 80 + "\n")
            for filename, score, _ in low_quality_files:
                text_length = results.get(filename, {}).get("text_length", 0)
                length_ratio = text_length / average_length if average_length > 0 else 0
                msg_status = "SENT" if sent_status.get(filename, {}).get("message_sent") else "NOT_SENT"
                issues = results.get(filename, {}).get("issues", [])
                issues_text = "; ".join(issues) if issues else "None"
                f.write(
                    f"{filename}: score = {score:.2f}, length = {text_length} chars ({length_ratio:.2f}x avg), message: {msg_status}\n"
                )
                f.write(f"  Issues: {issues_text}\n")
    if verbose:
        print(f"[Meaningfulness] Analysis results saved to {result_path}")
    else:
        print(f"Analysis results saved to {result_path}")
    
    with open(status_path, "w", encoding="utf-8") as f:
        json.dump(sent_status, f, ensure_ascii=False, indent=2)

    if low_quality_files:
        if verbose:
            print(f"[Meaningfulness] Found {len(low_quality_files)} low quality submissions (not yet notified):")
            for filename, score, _ in low_quality_files:
                text_length = results.get(filename, {}).get("text_length", 0)
                length_ratio = text_length / average_length if average_length > 0 else 0
                print(f"  {filename}: score = {score:.2f}, length = {text_length} chars ({length_ratio:.2f}x avg)")
        else:
            print(f"Found {len(low_quality_files)} low quality submissions (not yet notified).")
        
        if not auto_send:
            send_messages = get_input_with_timeout_default(
                "\nDo you want to send messages to students with low quality submissions? (y/n, or 'q' to quit, default 'y' in 60s): ",
                timeout=60,
                default="y"
            ).strip().lower()
            if send_messages in ("q", "quit"):
                with open(status_path, "w", encoding="utf-8") as f:
                    json.dump(sent_status, f, ensure_ascii=False, indent=2)
                return results
            if send_messages not in ("y", "yes", ""):
                if verbose:
                    print("[Meaningfulness] Messages not sent.")
                else:
                    print("Messages not sent.")
                with open(status_path, "w", encoding="utf-8") as f:
                    json.dump(sent_status, f, ensure_ascii=False, indent=2)
                return results
        
        try:
            canvas = get_canvas_client(api_url, api_key)
            course = canvas.get_course(course_id)
            for filename, score, text in low_quality_files:
                canvas_id = extract_canvas_id_from_filename(filename)
                if not canvas_id:
                    if verbose:
                        print(f"[Meaningfulness] Could not extract Canvas ID from {filename}")
                    else:
                        print(f"Could not extract Canvas ID from {filename}")
                    continue
                message = generate_low_quality_message(filename, score, text, refine)
                if verbose:
                    print(f"\n[Meaningfulness] Subject: {subject}")
                    print(f"[Meaningfulness] Message to {filename}:")
                    print("-" * 50)
                    print(message)
                    print("-" * 50)
                else:
                    print(f"\nPrepared message for {filename}.")
                if not auto_send:
                    while True:
                        action = get_input_with_timeout_default(
                            "\nWhat would you like to do? (s)end, (r)egenerate, or (q)uit [default: s in 60s]: ",
                            timeout=60,
                            default="s"
                        ).strip().lower()
                        if action in ('q', 'quit'):
                            if verbose:
                                print("[Meaningfulness] Quitting message sending.")
                            else:
                                print("Quitting message sending.")
                            with open(status_path, "w", encoding="utf-8") as f:
                                json.dump(sent_status, f, ensure_ascii=False, indent=2)
                            return results
                        elif action in ('s', 'send', ''):
                            break
                        elif action in ('r', 'regenerate'):
                            if verbose:
                                print("[Meaningfulness] Regenerating message...")
                            else:
                                print("Regenerating message...")
                            message = generate_low_quality_message(filename, score, text, refine)
                            if verbose:
                                print(f"\n[Meaningfulness] Regenerated message:")
                                print("-" * 50)
                                print(message)
                                print("-" * 50)
                            else:
                                print("Regenerated message.")
                        else:
                            if verbose:
                                print("[Meaningfulness] Please enter 's' to send, 'r' to regenerate, or 'q' to quit.")
                            else:
                                print("Please enter 's' to send, 'r' to regenerate, or 'q' to quit.")
                try:
                    canvas.create_conversation(
                        recipients=[str(canvas_id)],
                        subject="Yêu cầu định dạng lại và nộp lại bài tập",
                        body=message,
                        force_new=True
                    )
                    results[filename]["message_sent"] = True
                    sent_status[filename]["message_sent"] = True
                    if verbose:
                        print(f"[Meaningfulness] Message sent to student {canvas_id} for {filename}")
                    else:
                        print(f"Message sent to student {canvas_id} for {filename}")
                except Exception as e:
                    if verbose:
                        print(f"[Meaningfulness] Failed to send message for {filename}: {e}")
                    else:
                        print(f"Failed to send message for {filename}: {e}")
                    results[filename]["message_sent"] = False
                    sent_status[filename]["message_sent"] = False
        except Exception as e:
            if verbose:
                print(f"[Meaningfulness] Error setting up Canvas connection: {e}")
            else:
                print(f"Error setting up Canvas connection: {e}")
        
        with open(result_path, "a", encoding="utf-8") as f:
            f.write(f"\nMessage sending results:\n")
            f.write("-" * 80 + "\n")
            for filename, result in results.items():
                if result["meaningful_score"] < meaningfulness_threshold:
                    message_status = "SENT" if result["message_sent"] else "FAILED"
                    f.write(f"{filename}: Message {message_status}\n")
        with open(status_path, "w", encoding="utf-8") as f:
            json.dump(sent_status, f, ensure_ascii=False, indent=2)
    else:
        if verbose:
            print("[Meaningfulness] All submissions have acceptable meaningfulness scores.")
        else:
            print("All submissions have acceptable meaningfulness scores.")
        with open(status_path, "w", encoding="utf-8") as f:
            json.dump(sent_status, f, ensure_ascii=False, indent=2)
    
    return results

def extract_canvas_id_from_filename(filename):
    """
    Extract Canvas ID from filename with format: <student name>_<canvas id>_<time>_<status>.<ext>
    
    Args:
        filename (str): The PDF filename.
    
    Returns:
        int or None: Canvas ID if found, None otherwise.
    """
    parts = filename.split('_')
    for part in parts:
        if part.isdigit():
            return int(part)
    return None


