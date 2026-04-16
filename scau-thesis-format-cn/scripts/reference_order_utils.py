#!/usr/bin/env python3
from __future__ import annotations

import locale
import re
import unicodedata
from typing import Iterable


CHINESE_LOCALE_CANDIDATES = (
    "Chinese_China.936",
    "Chinese (Simplified)_China.936",
    "zh_CN.UTF-8",
    "zh_CN.GBK",
    "zh_CN.gb2312",
)


def strip_entry_prefix(text: str) -> str:
    return re.sub(r"^\s*(?:[-*]\s+|\[\d+\]\s*|\d+[.)、]\s*)", "", text).strip()


def first_author_token(entry: str) -> str:
    cleaned = strip_entry_prefix(entry)
    token_match = re.match(r"^([^\s,，;；.．:：]+)", cleaned)
    return token_match.group(1) if token_match else cleaned


def is_chinese_token(token: str) -> bool:
    return bool(re.search(r"[\u4e00-\u9fff]", token))


def detect_reference_language(entry: str) -> str:
    return "cn" if is_chinese_token(first_author_token(entry)) else "foreign"


def _pinyin_key(token: str) -> tuple[str, str]:
    try:
        from pypinyin import lazy_pinyin  # type: ignore

        return "".join(lazy_pinyin(token, errors="keep")).casefold(), "pypinyin"
    except Exception:
        previous_locale = None
        for candidate in CHINESE_LOCALE_CANDIDATES:
            try:
                previous_locale = locale.setlocale(locale.LC_COLLATE)
                locale.setlocale(locale.LC_COLLATE, candidate)
                return locale.strxfrm(token), f"locale:{candidate}"
            except locale.Error:
                continue
            finally:
                if previous_locale is not None:
                    try:
                        locale.setlocale(locale.LC_COLLATE, previous_locale)
                    except locale.Error:
                        pass
        return token.casefold(), "fallback:raw"


def _foreign_key(token: str) -> str:
    normalized = unicodedata.normalize("NFKD", token)
    return "".join(char for char in normalized if not unicodedata.combining(char)).casefold()


def reference_sort_key(entry: str) -> tuple[str, str, str]:
    token = first_author_token(entry)
    language = detect_reference_language(entry)
    if language == "cn":
        key, _ = _pinyin_key(token)
    else:
        key = _foreign_key(token)
    return language, key, strip_entry_prefix(entry).casefold()


def sort_reference_entries(entries: Iterable[str]) -> list[str]:
    cleaned_entries = [strip_entry_prefix(entry) for entry in entries if strip_entry_prefix(entry)]
    cn_entries = [entry for entry in cleaned_entries if detect_reference_language(entry) == "cn"]
    foreign_entries = [entry for entry in cleaned_entries if detect_reference_language(entry) != "cn"]
    return sorted(cn_entries, key=reference_sort_key) + sorted(foreign_entries, key=reference_sort_key)


def inspect_reference_sequence(entries: Iterable[str]) -> dict[str, object]:
    normalized_entries = [strip_entry_prefix(entry) for entry in entries if strip_entry_prefix(entry)]
    if not normalized_entries:
        return {
            "status": "manual_confirm",
            "expected": "参考文献应先列中文，再列西文/俄文；中文按第一著者姓氏汉语拼音字母顺序，西文和俄文按第一著者姓氏字母顺序。",
            "note": "未提取到参考文献条目。",
        }

    sequence = [
        {
            "entry": entry,
            "language": detect_reference_language(entry),
            "first_author_token": first_author_token(entry),
        }
        for entry in normalized_entries
    ]

    chinese_backend = "n/a"
    for item in sequence:
        if item["language"] == "cn":
            _, chinese_backend = _pinyin_key(str(item["first_author_token"]))
            break

    issues: list[dict[str, object]] = []
    seen_foreign = False
    for index, item in enumerate(sequence):
        if item["language"] != "cn":
            seen_foreign = True
            continue
        if seen_foreign:
            issues.append(
                {
                    "type": "language_group_order",
                    "index": index,
                    "entry": item["entry"],
                    "message": "中文参考文献出现在西文/俄文文献之后。",
                }
            )

    for language in ("cn", "foreign"):
        previous_key = None
        previous_entry = None
        previous_author = None
        for index, item in enumerate(sequence):
            if item["language"] != language:
                continue
            current_key = reference_sort_key(str(item["entry"]))[1]
            if previous_key is not None and current_key < previous_key:
                issues.append(
                    {
                        "type": "author_order",
                        "language": language,
                        "index": index,
                        "previous_entry": previous_entry,
                        "entry": item["entry"],
                        "previous_author": previous_author,
                        "current_author": item["first_author_token"],
                        "message": "同语言组内未按第一著者姓氏字母顺序排列。",
                    }
                )
            previous_key = current_key
            previous_entry = item["entry"]
            previous_author = item["first_author_token"]

    return {
        "status": "confirmed" if not issues else "suggested",
        "expected": "参考文献应先列中文，再列西文/俄文；中文按第一著者姓氏汉语拼音字母顺序，西文和俄文按第一著者姓氏字母顺序。",
        "entry_count": len(sequence),
        "cn_count": sum(1 for item in sequence if item["language"] == "cn"),
        "foreign_count": sum(1 for item in sequence if item["language"] != "cn"),
        "chinese_collation_backend": chinese_backend,
        "issues": issues[:40],
        "sequence_sample": sequence[:20],
    }
