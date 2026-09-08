# -*- coding: utf-8 -*-
# Shared submission checks (meaningfulness + similarity).

import os
from datetime import datetime

from tqdm import tqdm

from .settings import DEFAULT_AI_METHOD, DEFAULT_OCR_METHOD
from .data import (
    extract_text_from_scanned_pdf,
    analyze_text_meaningfulness,
    _compute_text_quality_metrics,
    _summarize_quality_issues,
)


def analyze_meaningfulness_in_folder(
    folder_path,
    ocr_service=DEFAULT_OCR_METHOD,
    lang="auto",
    meaningfulness_threshold=0.4,
    refine_method=DEFAULT_AI_METHOD,
    return_texts=False,
    write_report=True,
    verbose=False,
):
    """
    Analyze meaningfulness for PDFs in a folder. Returns (results, low_quality, texts, average_length).
    """
    refine_method = refine_method or DEFAULT_AI_METHOD
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
        return {}, [], {}, 0.0

    results = {}
    low_quality = []
    text_lengths = []
    extracted_texts = {}

    for pdf_path in tqdm(pdf_files, desc="Extracting texts from PDFs"):
        filename = os.path.basename(pdf_path)
        base = os.path.splitext(pdf_path)[0]
        txt_path = base + f"_text_{ocr_service}.txt"
        if not os.path.exists(txt_path):
            txt_path = extract_text_from_scanned_pdf(
                pdf_path,
                txt_output_path=txt_path,
                service=ocr_service,
                lang=lang,
                simple_text=False,
            )
        if not txt_path or not os.path.exists(txt_path):
            results[filename] = {"meaningful_score": 0.0, "error": "Failed to extract text"}
            continue
        with open(txt_path, "r", encoding="utf-8") as f:
            text = f.read()
        extracted_texts[filename] = text
        text_lengths.append(len(text.strip()))

    average_length = sum(text_lengths) / len(text_lengths) if text_lengths else 0.0
    if verbose:
        print(f"[Meaningfulness] Average text length: {average_length:.0f} characters")
    else:
        print(f"Average text length: {average_length:.0f} characters")

    for filename, text in tqdm(extracted_texts.items(), desc="Analyzing PDF meaningfulness"):
        meaningful_score = analyze_text_meaningfulness(text, refine_method, average_length)
        metrics = _compute_text_quality_metrics(text)
        issues = _summarize_quality_issues(metrics, average_length=average_length)
        results[filename] = {
            "meaningful_score": meaningful_score,
            "text_length": len(text.strip()),
            "issues": issues,
            "metrics": {
                "vn_char_ratio": metrics["vn_char_ratio"],
                "alnum_ratio": metrics["alnum_ratio"],
                "symbol_ratio": metrics["symbol_ratio"],
                "unique_char_ratio": metrics["unique_char_ratio"],
                "repeat_char_ratio": metrics["repeat_char_ratio"],
                "line_empty_ratio": metrics["line_empty_ratio"],
                "likely_math": metrics["likely_math"],
            },
        }
        if meaningful_score < meaningfulness_threshold:
            low_quality.append(filename)

    if write_report:
        result_path = os.path.join(folder_path, "meaningfulness_analysis.txt")
        with open(result_path, "w", encoding="utf-8") as f:
            f.write("PDF meaningfulness analysis results:\n")
            f.write(f"Threshold: {meaningfulness_threshold}\n")
            f.write(f"Average text length: {average_length:.0f} characters\n")
            f.write(f"Analysis date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
            f.write("Detailed results:\n")
            f.write("-" * 80 + "\n")
            for filename, result in sorted(results.items()):
                score = result["meaningful_score"]
                status = "LOW QUALITY" if score < meaningfulness_threshold else "ACCEPTABLE"
                text_length = result.get("text_length", 0)
                issues_text = "; ".join(result.get("issues", [])) or "None"
                f.write(f"{filename}: score = {score:.2f} ({status}), length = {text_length} chars\n")
                f.write(f"  Issues: {issues_text}\n")
        if verbose:
            print(f"[Meaningfulness] Analysis results saved to {result_path}")
        else:
            print(f"Analysis results saved to {result_path}")

    return results, low_quality, (extracted_texts if return_texts else {}), average_length


def compare_texts_from_pdfs_in_folder(*args, **kwargs):
    """
    Wrapper for the Canvas similarity checker to keep shared API for Google Classroom.
    """
    from .canvas import compare_texts_from_pdfs_in_folder as _compare
    return _compare(*args, **kwargs)
