"""ONLY TINC_NAME and SINC_NAME: canonical values and rule matching.

Order follows the business-supplied list, not alphabetical/severity sorting.
Both stock spellings (MIN/MINOR, MED/MEDIUM, HVY/HEAVY) resolve to the
same grade. Values not in the 18 grades raise a clear error, not a silent skip.
"""
import re

import pandas as pd

INCLUSION_COLUMNS = frozenset({'TINC_NAME', 'SINC_NAME'})
INCLUSION_ORDER = [
    'NONE',
    'MIN PINPOINT', 'MED PINPOINT',
    'MIN CRYSTAL', 'MED CRYSTAL',
    'MIN SPREADED', 'MED SPREADED', 'HVY SPREADED',
    'MIN FEATHER', 'MED FEATHER', 'HVY FEATHER',
    'MIN CLOUD', 'MED CLOUD', 'HVY CLOUD',
    'HVY CRYSTAL',
    'CRYSTEAL 1', 'CRYSTEAL 2', 'CRYSTEAL 3',
]


def normalize_inclusion_value(value):
    if pd.isna(value):
        return ''
    result = re.sub(r'\s+', ' ', str(value).strip().upper())
    result = re.sub(r'\bMINOR\b', 'MIN', result)
    result = re.sub(r'\bMEDIUM\b', 'MED', result)
    result = re.sub(r'\bHEAVY\b', 'HVY', result)
    result = re.sub(r'^CRYST(?:EAL|AL)\s*([123])$', r'CRYSTEAL \1', result)
    return 'NONE' if result == 'NON' else result


def _expand_one(term):
    m = re.fullmatch(r'(.+?)\s+TO\s+(.+)', str(term).strip(), flags=re.IGNORECASE)
    if m:
        first, last = (normalize_inclusion_value(x) for x in m.groups())
        if first not in INCLUSION_ORDER or last not in INCLUSION_ORDER:
            raise ValueError(f'Unknown inclusion range: {term!r}')
        a, b = sorted((INCLUSION_ORDER.index(first), INCLUSION_ORDER.index(last)))
        return INCLUSION_ORDER[a:b+1]
    grade = normalize_inclusion_value(term)
    if grade not in INCLUSION_ORDER:
        raise ValueError(f'Unknown inclusion grade: {term!r}')
    return [grade]


def inclusion_rule_groups(rule, tolerance_level=0):
    """Return accepted and blocked canonical grades, with existing +/- tolerance."""
    included, excluded = set(), set()
    for part in str(rule).split(','):
        part = part.strip()
        if not part: continue
        is_exclude = part.startswith('!')
        grades = _expand_one(part[1:].strip() if is_exclude else part)
        target = excluded if is_exclude else included
        target.update(grades)
    tolerance = max(0, int(tolerance_level or 0))
    if tolerance:
        def widen(grades):
            widened = set()
            for grade in grades:
                i = INCLUSION_ORDER.index(grade)
                widened.update(INCLUSION_ORDER[max(0,i-tolerance):i+tolerance+1])
            return widened
        included, excluded = widen(included), widen(excluded)
    return included, excluded


def inclusion_value_matches(stock_value, rule, tolerance_level=0):
    included, excluded = inclusion_rule_groups(rule, tolerance_level)
    value = normalize_inclusion_value(stock_value)
    return (not included or value in included) and value not in excluded


def inclusion_condition(series, rule, tolerance_level=0):
    included, excluded = inclusion_rule_groups(rule, tolerance_level)
    normalized = series.map(normalize_inclusion_value)
    result = pd.Series(True, index=series.index)
    if included:
        result &= normalized.isin(included)
    if excluded:
        result &= ~normalized.isin(excluded)
    return result
