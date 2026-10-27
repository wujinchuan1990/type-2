"""Shared equipment-library and output typography, without changing task snapshots."""

from copy import deepcopy
import re


TEXT_FIELDS = ("name", "model", "performance", "number", "storage_location")
UNCERTAINTY = re.compile(r"(?<![A-Za-z0-9_])U(?:rel|ᵣₑₗ)?(?![A-Za-z0-9_])")
COVERAGE_FACTOR = re.compile(r"(?<![A-Za-z0-9_])k(?![A-Za-z0-9_])")


def fullwidth_parentheses(text):
    return text.replace("(", "（").replace(")", "）")


def equipment_identity(name, model, number):
    return tuple(fullwidth_parentheses(value.strip()) for value in (name, model, number))


def normalize_equipment_typography(item):
    result = deepcopy(item)
    formatting = result.get("formatting")
    for key in TEXT_FIELDS:
        if key not in result:
            continue
        text = result[key]
        runs = formatting.get(key) if formatting is not None else None
        if runs is not None and "".join(run["text"] for run in runs) != text:
            raise ValueError(f"设备 {key} 的格式化文字必须与原文字完全一致，不能自动修正格式。")
        result[key] = fullwidth_parentheses(text)
        if runs is not None:
            for run in runs:
                run["text"] = fullwidth_parentheses(run["text"])
    text = result.get("performance", "")
    runs = (formatting or {}).get("performance")
    symbols = [(match, True) for match in UNCERTAINTY.finditer(text)]
    symbols.extend((match, False) for match in COVERAGE_FACTOR.finditer(text))
    if not symbols:
        return result
    original_runs = runs if runs is not None else [{"text": text, "italic": False, "script": "normal"}]
    characters = [[character, run.get("italic", False), run.get("script", "normal")]
                  for run in original_runs for character in run["text"]]
    original_characters = deepcopy(characters)
    for match, uncertainty in symbols:
        characters[match.start()][1:] = [True, "normal"]
        if uncertainty and len(match[0]) > 1:
            for offset, character in enumerate("rel", match.start() + 1):
                characters[offset] = [character, False, "sub"]
    if characters == original_characters:
        return result
    changed = []
    offset = 0
    for original in original_runs:
        parts = []
        for character, italic, script in characters[offset:offset + len(original["text"])]:
            if parts and parts[-1]["italic"] == italic and parts[-1]["script"] == script:
                parts[-1]["text"] += character
            else:
                parts.append({"text": character, "italic": italic, "script": script})
        changed.extend(parts)
        offset += len(original["text"])
    if changed != original_runs:
        if result.get("formatting") is None:
            result["formatting"] = {}
        result["formatting"]["performance"] = changed
        result["performance"] = "".join(run["text"] for run in changed)
    return result
