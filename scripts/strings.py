"""Translations. English text is the key. Add Hindi (hi) and Gujarati (gu) entries in strings_hi.py / strings_gu.py."""
import os
MISSING = set()
_cache = {}
def tr(s, lang):
    if lang == "en": return s
    if not any(ch.isalpha() for ch in s.replace("Rs ", "")): return s  # numbers only
    if lang not in _cache:
        mod = __import__(f"strings_{lang}"); _cache[lang] = mod.STR
    if s in _cache[lang]: return _cache[lang][s]
    if os.environ.get("COLLECT"):
        MISSING.add(s); return s
    raise KeyError(f"No {lang} translation for: {s!r}")
