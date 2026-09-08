# -*- coding: utf-8 -*-
# Canvas API: https://canvas.instructure.com/doc/api/

"""Canvas assignment helpers."""

from datetime import datetime

from .canvas_auth import get_canvas_client

from .settings import (
    CANVAS_LMS_API_URL,
    CANVAS_LMS_API_KEY,
    CANVAS_LMS_COURSE_ID,
    CANVAS_DEFAULT_ASSIGNMENT_CATEGORY,
)
from .utils import format_time

def list_canvas_assignments(
    api_url=CANVAS_LMS_API_URL,
    api_key=CANVAS_LMS_API_KEY,
    course_id=CANVAS_LMS_COURSE_ID,
    category=None,
    verbose=False
):
    """
    List all assignments in a Canvas course, grouped by assignment category (assignment group), using the canvasapi library.
    If category is specified (case-insensitive), only list assignments in that category.
    Assignments are sorted by due date (earliest first; undated last).

    Args:
        api_url (str): The base URL for the Canvas instance (e.g., "https://canvas.instructure.com")
        api_key (str): Your Canvas API access token
        course_id (int or str): The Canvas course ID
        category (str, optional): Assignment group name to filter (case-insensitive)
        verbose (bool): If True, print more details; otherwise, print only important notice.

    Returns:
        Dict mapping assignment group names to lists of assignment dicts with id, name, due_at, and points_possible
    """
    try:
        canvas = get_canvas_client(api_url, api_key)
        course = canvas.get_course(course_id)
        if verbose:
            print(f"[CanvasAssignments] Listing assignments for course: \"{course.name} (ID: {course.id})\"")
        else:
            print(f"Listing assignments for course: \"{course.name} (ID: {course.id})\"")
        # Get all assignment groups with assignments included
        assignment_groups = course.get_assignment_groups(include=['assignments'])
        group_assignments = {}
        for group in assignment_groups:
            group_name = group.name
            assignments = []
            for assignment in group.assignments:
                assignments.append({
                    "id": assignment['id'],
                    "name": assignment['name'],
                    "due_at": format_time(assignment.get('due_at')),
                    "due_at_raw": assignment.get('due_at'),
                    "points_possible": assignment.get('points_possible')
                })
            # Sort assignments by due date (None last)
            def due_sort_key(a):
                raw = a.get("due_at_raw")
                if raw:
                    try:
                        return datetime.strptime(raw, "%Y-%m-%dT%H:%M:%SZ")
                    except Exception:
                        return datetime.max
                return datetime.max
            assignments.sort(key=due_sort_key)
            # Remove 'due_at_raw' from output
            for a in assignments:
                a.pop("due_at_raw", None)
            group_assignments[group_name] = assignments
        # If category is specified, filter to only that group (case-insensitive)
        if category:
            matched = None
            for group_name in group_assignments:
                if group_name.lower() == category.lower():
                    matched = group_name
                    break
            if matched:
                group_assignments = {matched: group_assignments[matched]}
            else:
                if verbose:
                    print(f"[CanvasAssignments] No assignment group found matching category '{category}'.")
                else:
                    print(f"No assignment group found matching category '{category}'.")
                return {}
        # Print assignments grouped by category
        for group_name, assignments in group_assignments.items():
            if verbose:
                print(f"[CanvasAssignments] Category: {group_name} ({len(assignments)} assignments)")
            else:
                print(f"\nCategory: {group_name} ({len(assignments)} assignments)")
            for a in assignments:
                if verbose:
                    print(f"[CanvasAssignments]   ID: {a['id']}, Name: {a['name']}, Due: {a['due_at']}, Points: {a['points_possible']}")
                else:
                    print(f"  ID: {a['id']}, Name: {a['name']}, Due: {a['due_at']}, Points: {a['points_possible']}")
        return group_assignments
    except ImportError:
        if verbose:
            print("[CanvasAssignments] canvasapi library is not installed. Please install it with 'pip install canvasapi'.")
        else:
            print("canvasapi library is not installed. Please install it with 'pip install canvasapi'.")
        return {}
    except Exception as e:
        if verbose:
            print(f"[CanvasAssignments] Error listing assignments: {e}")
        else:
            print(f"Error listing assignments: {e}")
        return {}

