from __future__ import annotations

import random
from collections import OrderedDict
from typing import Dict, Iterable, List, Mapping, MutableMapping, Sequence, Tuple

# DEV-only Random Sim Studio foundation. These are categories backed by the
# scanner's established CAS classifications; weak/unknown/non-CAS rows are not
# eligible for random generation.
ELIGIBLE_BASE_TYPES = {
    "hair", "body", "top", "bottom", "shoes", "accessory", "earrings",
    "glasses_f", "bracelets", "ring_lt", "beard", "lipstick", "eyeshadow",
    "eyeliner", "blush", "makeup", "eyebrow", "glove", "socks", "moles",
    "armband", "tattoo", "pet marking", "skin", "nose preset", "lip preset",
    "eye preset", "face preset",
}

DISPLAY_NAMES = {
    "glasses_f": "Glasses",
    "ring_lt": "Ring",
    "armband": "Arm Band",
}

# Categories where choosing two pieces is especially useful. The UI still lets
# the tester choose 0, 1 or 2 for every category.
DEFAULT_COUNTS = {
    "accessory": 2,
    "earrings": 2,
    "bracelets": 2,
    "ring_lt": 2,
    "makeup": 2,
    "moles": 2,
    "tattoo": 2,
}

BAD_HEALTH = {"broken", "malformed"}
DELETED_STATUSES = {"delete", "deleted"}
RULE_WEIGHT = "weight"
RULE_EXACT = "exact"
UNSPECIFIED_SUBTYPE = "(Unspecified)"


def base_cc_type(value: object) -> str:
    """Return scanner CC Type without the display-only ' — subtype' suffix."""
    text = str(value or "").strip()
    if " — " in text:
        text = text.split(" — ", 1)[0].strip()
    return text


def category_key(record: Mapping[str, object]) -> str:
    """Stable casefolded random-selection category from validated CC Type."""
    return base_cc_type(record.get("CC Type", "")).casefold()


def category_label(key: str) -> str:
    key = str(key or "").strip().casefold()
    if key in DISPLAY_NAMES:
        return DISPLAY_NAMES[key]
    return " ".join(part.capitalize() for part in key.replace("_", " ").split())


def subtype_label(record: Mapping[str, object]) -> str:
    """Return the scanner-recorded CAS subtype, with a stable empty fallback."""
    raw = str(record.get("Subtype", "") or "").strip()
    if not raw:
        cc_type = str(record.get("CC Type", "") or "").strip()
        if " — " in cc_type:
            raw = cc_type.split(" — ", 1)[1].strip()
    return raw or UNSPECIFIED_SUBTYPE


def subtype_key(record: Mapping[str, object]) -> str:
    return subtype_label(record).casefold()


def type_rule_key(category: str, subtype: str) -> str:
    """Persistent identifier for one category/subtype probability rule."""
    return f"{str(category or '').strip().casefold()}\x1f{str(subtype or UNSPECIFIED_SUBTYPE).strip().casefold()}"


def record_key(record: Mapping[str, object]) -> str:
    """Persistent package identity without depending on row order."""
    for field in ("Package Hash", "Normalized Package", "Original Source Path", "Package"):
        value = str(record.get(field, "") or "").strip()
        if value:
            return f"{field}:{value.casefold()}"
    return ""


def is_eligible_record(record: Mapping[str, object], excluded_keys: Iterable[str] = ()) -> bool:
    key = category_key(record)
    if key not in ELIGIBLE_BASE_TYPES:
        return False
    if record_key(record) in set(excluded_keys):
        return False
    if str(record.get("Status", "") or "").strip().casefold() in DELETED_STATUSES:
        return False
    if str(record.get("Health Status", "") or "").strip().casefold() in BAD_HEALTH:
        return False
    return True


def group_eligible_records(records: Sequence[Mapping[str, object]], excluded_keys: Iterable[str] = ()) -> "OrderedDict[str, List[Mapping[str, object]]]":
    excluded = set(excluded_keys)
    grouped: MutableMapping[str, List[Mapping[str, object]]] = {}
    for record in records:
        if not is_eligible_record(record, excluded):
            continue
        key = category_key(record)
        grouped.setdefault(key, []).append(record)

    ordered = OrderedDict()
    for key in sorted(grouped, key=lambda k: category_label(k).casefold()):
        ordered[key] = sorted(grouped[key], key=lambda r: str(r.get("Package", "")).casefold())
    return ordered


def default_count_for_category(key: str) -> int:
    return DEFAULT_COUNTS.get(str(key or "").casefold(), 1)


def _normalise_rule(rule: object) -> Tuple[str, float]:
    if not isinstance(rule, Mapping):
        return RULE_WEIGHT, 1.0
    mode = str(rule.get("mode", RULE_WEIGHT) or RULE_WEIGHT).strip().casefold()
    if mode not in {RULE_WEIGHT, RULE_EXACT}:
        mode = RULE_WEIGHT
    try:
        value = float(rule.get("value", 1.0 if mode == RULE_WEIGHT else 0.0))
    except (TypeError, ValueError):
        value = 1.0 if mode == RULE_WEIGHT else 0.0
    if mode == RULE_EXACT:
        value = max(0.0, min(100.0, value))
    else:
        value = max(0.0, value)
    return mode, value


def _weighted_choice(entries: Sequence[Tuple[object, float]], rng: random.Random):
    positive = [(entry, float(weight)) for entry, weight in entries if float(weight) > 0]
    if not positive:
        return None
    total = sum(weight for _, weight in positive)
    mark = rng.random() * total
    running = 0.0
    for entry, weight in positive:
        running += weight
        if mark < running:
            return entry
    return positive[-1][0]


def _probability_choice(
    entries: Sequence[object],
    rules: Mapping[str, Mapping[str, object]] | None,
    key_func,
    rng: random.Random,
):
    """Choose one entry using exact probability bands plus a weighted remainder.

    Exact % rules consume their literal percentage of a single draw. The remainder
    is distributed among entries in Weight mode (including entries with no saved
    rule, which default to weight 1). If every entry is Exact % and those values sum
    below 100, the unassigned remainder intentionally produces no pick.
    """
    if not entries:
        return None
    rules = rules or {}
    exact: List[Tuple[object, float]] = []
    weighted: List[Tuple[object, float]] = []
    for entry in entries:
        rule_key = str(key_func(entry) or "")
        mode, value = _normalise_rule(rules.get(rule_key))
        if mode == RULE_EXACT:
            exact.append((entry, value))
        else:
            weighted.append((entry, value))

    exact_total = sum(value for _, value in exact)
    if exact_total > 100.000001:
        # UI normally blocks this configuration. Keeping the core conservative
        # means invalid saved data cannot silently be normalised into different %s.
        return None

    mark = rng.random() * 100.0
    running = 0.0
    for entry, value in sorted(exact, key=lambda pair: str(key_func(pair[0]))):
        running += value
        if mark < running:
            return entry

    # Exact bands occupy [0, exact_total). The remainder belongs to Weight mode.
    if mark < 100.0 and weighted:
        return _weighted_choice(weighted, rng)
    return None


def _choose_record(
    category: str,
    candidates: Sequence[Mapping[str, object]],
    type_rules: Mapping[str, Mapping[str, object]] | None,
    item_rules: Mapping[str, Mapping[str, object]] | None,
    rng: random.Random,
):
    by_subtype: MutableMapping[str, List[Mapping[str, object]]] = {}
    labels: Dict[str, str] = {}
    for record in candidates:
        skey = subtype_key(record)
        by_subtype.setdefault(skey, []).append(record)
        labels.setdefault(skey, subtype_label(record))
    subtype_ids = sorted(by_subtype)
    chosen_subtype = _probability_choice(
        subtype_ids,
        type_rules,
        lambda skey: type_rule_key(category, labels.get(str(skey), str(skey))),
        rng,
    )
    if chosen_subtype is None:
        return None
    pool = by_subtype.get(str(chosen_subtype), [])
    return _probability_choice(pool, item_rules, record_key, rng)


def validate_probability_rules(
    grouped: Mapping[str, Sequence[Mapping[str, object]]],
    type_rules: Mapping[str, Mapping[str, object]] | None = None,
    item_rules: Mapping[str, Mapping[str, object]] | None = None,
) -> List[str]:
    """Return human-readable conflicts for literal Exact % rules.

    A group of exact rules may total at most 100. The remaining percentage is the
    Weight-mode pool. Stale rules for packages/types not present in the current MDB
    are ignored rather than blocking generation.
    """
    type_rules = type_rules or {}
    item_rules = item_rules or {}
    errors: List[str] = []

    for category, records in grouped.items():
        subtypes: MutableMapping[str, List[Mapping[str, object]]] = {}
        labels: Dict[str, str] = {}
        for record in records:
            skey = subtype_key(record)
            subtypes.setdefault(skey, []).append(record)
            labels.setdefault(skey, subtype_label(record))

        type_exact = 0.0
        for skey in subtypes:
            mode, value = _normalise_rule(type_rules.get(type_rule_key(category, labels[skey])))
            if mode == RULE_EXACT:
                type_exact += value
        if type_exact > 100.000001:
            errors.append(
                f"{category_label(category)} type/subtype Exact % values total {type_exact:.2f}% (maximum 100%)."
            )

        for skey, subtype_records in subtypes.items():
            item_exact = 0.0
            for record in subtype_records:
                mode, value = _normalise_rule(item_rules.get(record_key(record)))
                if mode == RULE_EXACT:
                    item_exact += value
            if item_exact > 100.000001:
                errors.append(
                    f"{category_label(category)} / {labels[skey]} item Exact % values total {item_exact:.2f}% (maximum 100%)."
                )
    return errors


def generate_selection(
    grouped: Mapping[str, Sequence[Mapping[str, object]]],
    counts: Mapping[str, int],
    locked: Mapping[str, Sequence[str]] | None = None,
    avoid: Mapping[str, Sequence[str]] | None = None,
    rng: random.Random | None = None,
    category_chances: Mapping[str, float] | None = None,
    type_rules: Mapping[str, Mapping[str, object]] | None = None,
    item_rules: Mapping[str, Mapping[str, object]] | None = None,
) -> Dict[str, List[Mapping[str, object]]]:
    """Generate category selections while preserving still-valid locked items.

    Probability layers are DEV-only generator metadata and never edit the MDB:
    category chance -> type/subtype rule -> individual-item rule. Locks win over
    probability; exclusions are already removed from grouped pools.
    """
    rng = rng or random.Random()
    locked = locked or {}
    avoid = avoid or {}
    category_chances = category_chances or {}
    type_rules = type_rules or {}
    item_rules = item_rules or {}
    result: Dict[str, List[Mapping[str, object]]] = {}

    for category, pool in grouped.items():
        wanted = max(0, min(2, int(counts.get(category, 0))))
        if wanted <= 0 or not pool:
            result[category] = []
            continue

        pool_by_key = {record_key(rec): rec for rec in pool if record_key(rec)}
        kept: List[Mapping[str, object]] = []
        seen = set()
        for locked_key in locked.get(category, ()):  # preserve row lock order
            rec = pool_by_key.get(locked_key)
            if rec is None or locked_key in seen:
                continue
            kept.append(rec)
            seen.add(locked_key)
            if len(kept) >= wanted:
                break

        try:
            chance = float(category_chances.get(category, 100.0))
        except (TypeError, ValueError):
            chance = 100.0
        chance = max(0.0, min(100.0, chance))
        category_active = rng.random() * 100.0 < chance
        if not category_active:
            result[category] = kept
            continue

        avoided = set(avoid.get(category, ()))
        candidates = [rec for rec in pool if record_key(rec) not in seen and record_key(rec) not in avoided]
        # A reroll should prefer a different package, but if every available
        # candidate is the one being rerolled, keep the slot populated.
        if len(candidates) < max(0, wanted - len(kept)):
            candidates.extend(
                rec for rec in pool
                if record_key(rec) not in seen and record_key(rec) in avoided and rec not in candidates
            )

        while len(kept) < wanted and candidates:
            chosen = _choose_record(category, candidates, type_rules, item_rules, rng)
            if chosen is None:
                break
            chosen_key = record_key(chosen)
            kept.append(chosen)
            seen.add(chosen_key)
            candidates = [rec for rec in candidates if record_key(rec) != chosen_key]
        result[category] = kept

    return result
