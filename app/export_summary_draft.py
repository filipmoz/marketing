"""
Draft: plain-text summary of survey responses for quick checks.
"""
from typing import List
from collections import Counter

def build_summary_text(responses: List) -> str:
    if not responses:
        return "No responses yet.\n"
    lines = [f"Total responses: {len(responses)}"]
    genders = [r.gender for r in responses if getattr(r, "gender", None)]
    if genders:
        for k, v in Counter(genders).items():
            lines.append(f"  {k}: {v}")
    return "\n".join(lines) + "\n"
