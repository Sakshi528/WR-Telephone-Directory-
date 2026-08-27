"""Pins our own organization's sections to the top of the generated
telephone directory (management requirement), ahead of the rest of the
alphabetical ordering everywhere else already uses.

Exact org names, not substring matching -- "WRPC"/"WRLDC" don't appear
literally anywhere in organizations.organization_name, so pinning has to
key off the real names. Verified single, unambiguous matches in the live
DB before hardcoding these:
  - "Western Regional Power Committee" (org id 6291) -- WRPC
  - "Western Region Load Despatch Centre" (org id 619) -- WRLDC
(there are other Regional Load Despatch Centres -- Northern/Eastern/
Southern/North-Eastern -- but none of them are named "Western Region[al]
Load Despatch Centre", so there's no collision risk.)

Usage:
    from utils.section_order import section_sort_key
    grouped = dict(sorted(grouped.items(), key=lambda kv: section_sort_key(kv[0])))
"""

PINNED_SECTIONS = [
    "Western Regional Power Committee",   # WRPC -- first
    "Western Region Load Despatch Centre",  # WRLDC -- second
]


def section_sort_key(org_name):
    """Sort key for a dict of {org_name: ...} sections: pinned sections
    first in PINNED_SECTIONS order, then everything else alphabetically,
    unaffected."""
    if org_name in PINNED_SECTIONS:
        return (0, PINNED_SECTIONS.index(org_name), "")
    return (1, 0, (org_name or "").lower())
