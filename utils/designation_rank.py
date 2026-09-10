"""Generic, keyword-based designation ranking used to automatically resolve
an organization's Utility Head from its employees' designations, instead of
relying on a manually-maintained flag.

Usage:
    from utils.designation_rank import resolve_utility_head
    head = resolve_utility_head(employees)   # employees ordered as listed
"""

import re

# Highest seniority first. Each tier lists every full title / abbreviation
# that should resolve to that same rank -- add new titles here, no code
# changes needed elsewhere.
RANKING_TIERS = [
    ["chairman"],
    ["chairman & managing director", "cmd"],
    ["managing director", "md"],
    ["director"],
    ["chief executive officer", "ceo"],
    ["executive director", "ed"],
    ["chief general manager", "cgm"],
    ["additional general manager", "agm"],
    ["general manager", "gm"],
    ["deputy general manager", "dgm"],
    ["senior general manager", "sgm"],
    ["chief engineer", "ce"],
    ["additional chief engineer", "ace"],
    ["superintending engineer", "se"],
    ["executive engineer", "ee"],
    ["deputy chief engineer"],
    ["plant head"],
    ["station head"],
    ["head"],
    ["senior manager"],
    ["manager"],
    ["deputy manager"],
    ["assistant manager"],
    ["senior engineer"],
    ["engineer"],
    ["assistant engineer"],
    ["executive"],
    ["officer"],
    ["supervisor"],
    ["technician"],
    ["operator"],
]

UNRANKED = len(RANKING_TIERS)

_TIER_PATTERNS = [
    [(phrase, re.compile(r"\b" + re.escape(phrase) + r"\b", re.IGNORECASE)) for phrase in tier]
    for tier in RANKING_TIERS
]


def rank_designation(designation):
    """Returns (rank, matched_phrase) for a designation string -- lower rank
    is more senior, UNRANKED means no configured title was found. Matching
    is whole-word and case-insensitive; the LONGEST matching phrase wins
    across all tiers (not the first tier checked), so a more specific title
    like "Additional Chief Engineer" isn't shadowed by "Chief Engineer",
    which it contains."""
    text = (designation or "").strip()
    if not text:
        return UNRANKED, None

    best_rank, best_phrase = UNRANKED, None
    for rank, patterns in enumerate(_TIER_PATTERNS):
        for phrase, pattern in patterns:
            if (best_phrase is None or len(phrase) > len(best_phrase)) and pattern.search(text):
                best_rank, best_phrase = rank, phrase
    return best_rank, best_phrase


def resolve_utility_head(employees):
    """Returns the organization's Utility Head from the given list of its
    employees -- a manually-flagged employee (Employee.is_utility_head)
    always wins over the designation-based auto-resolution below, so an
    admin's explicit choice sticks regardless of designation text. Falls
    back to the most senior designation (per rank_designation) when no one
    is manually flagged -- ties, including all-unranked, keep the first
    employee in the list. Returns None for an empty list."""
    manual = next((e for e in employees if getattr(e, "is_utility_head", False)), None)
    if manual is not None:
        return manual

    best, best_rank = None, None
    for emp in employees:
        rank, _phrase = rank_designation(getattr(emp, "designation", None))
        if best is None or rank < best_rank:
            best, best_rank = emp, rank
    return best


def compute_utility_head_ids():
    """Returns the set of Employee ids that are each organization's
    resolved Utility Head -- a manually-flagged employee (Employee.
    is_utility_head) if one exists for that organization, otherwise the
    designation-based auto-resolution. An organization with
    Organization.utility_head_excluded set is skipped entirely -- it has no
    Utility Head at all, regardless of employees. Computed fresh on every
    call. Used everywhere the app needs to know "is this employee the head"
    without recomputing per row (e.g. a template's star icon)."""
    from collections import defaultdict
    from models.employee import Employee
    from models.organization import Organization

    excluded_org_ids = {
        o.id for o in Organization.query.filter_by(utility_head_excluded=True).all()
    }

    by_org = defaultdict(list)
    # Non-ACTIVE employees are never eligible to resolve as a Utility Head.
    for e in Employee.query.filter(Employee.status == "ACTIVE").order_by(Employee.id).all():
        if e.organization_id is not None and e.organization_id not in excluded_org_ids:
            by_org[e.organization_id].append(e)

    ids = set()
    for employees in by_org.values():
        head = resolve_utility_head(employees)
        if head:
            ids.add(head.id)
    return ids
