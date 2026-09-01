from __future__ import annotations

import re


TOKEN_RE = re.compile(r"[a-z0-9]+(?:'[a-z0-9]+)?")

STOP_WORDS = {
    "a",
    "an",
    "and",
    "are",
    "as",
    "at",
    "be",
    "by",
    "for",
    "from",
    "has",
    "in",
    "is",
    "it",
    "of",
    "on",
    "or",
    "that",
    "the",
    "to",
    "was",
    "with",
}


def tokenize(text: str, remove_stop_words: bool = True) -> list[str]:
    tokens = TOKEN_RE.findall(text.lower())
    if remove_stop_words:
        return [token for token in tokens if token not in STOP_WORDS and len(token) > 1]
    return tokens

