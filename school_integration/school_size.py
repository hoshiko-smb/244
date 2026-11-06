"""訪問シートの「最上級生の全体人数」欄などから、管理表に入れる学校規模（1学年○○人程度）を作る。

単純な記述（「約400人」「800名弱」「140〜150名」）は機械的に整形し、
複合的な記述（「大学800名・短大200名」「日本人13名／留学生75名」など）は
size_overrides.json（生の記述 → 表示文字列）で整形する。override が無い複合記述は
「要確認」として生の記述を残し、営業担当の確認対象にする。
"""
import re
import unicodedata

NUM = r"[\d,，]+"
SIMPLE = re.compile(
    rf"^(?:全体|全体で|学年|1学年|最上級生|4年生|毎年)?[:：]?\s*(約|およそ|大体|だいたい)?\s*({NUM})\s*(名|人)?\s*"
    r"(程度|ほど|くらい|ぐらい|位|弱|強|前後|程)?\s*$"
)
RANGE = re.compile(
    rf"^(約|およそ)?\s*({NUM})\s*(?:名|人)?\s*[〜~～\-－ー]\s*({NUM})\s*(名|人)?\s*(程度|ほど|くらい|位)?\s*$"
)
EMPTY_WORDS = {"不明", "未ヒアリング", "把握していない", "ヒアリングできず", "ヒアリングできていません",
               "確認中", "-", "ー", "―", "−", "なし", ""}


def _n(s: str) -> int:
    return int(re.sub(r"[,，]", "", s))


def prep(raw: str) -> str:
    s = unicodedata.normalize("NFKC", raw or "")
    s = s.replace("&#10;", " ").replace("\\~", "~").replace("\\", "")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def fmt(n: int) -> str:
    return f"{n:,}"


def normalize(raw: str, overrides: dict) -> tuple[str | None, str]:
    """(表示文字列 or None, 判定) を返す。判定は auto / override / empty / review"""
    if raw is None:
        return None, "empty"
    if raw in overrides:
        v = overrides[raw]
        return (v or None), ("override" if v else "empty")
    s = prep(raw)
    if s in overrides:
        v = overrides[s]
        return (v or None), ("override" if v else "empty")
    if s in EMPTY_WORDS:
        return None, "empty"
    m = SIMPLE.match(s)
    if m:
        n = _n(m.group(2))
        if n < 5:  # 「400人/7人」のような分母分子の取り違えを避ける
            return None, "review"
        q = m.group(4) or ""
        tail = "人弱" if q == "弱" else "人強" if q == "強" else "人程度"
        return f"1学年{fmt(n)}{tail}", "auto"
    m = RANGE.match(s)
    if m:
        a, b = _n(m.group(2)), _n(m.group(3))
        if a < b:
            return f"1学年{fmt(a)}〜{fmt(b)}人程度", "auto"
    return None, "review"
