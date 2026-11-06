"""競合（AGIC＝エイジェック ほか）の訪問有無をテキストから判定する。

判定は3値: 「あり」「なし」「不明」。
 - 社名の言及があり、直後で否定（来ていない/訪問なし 等）されていなければ「あり」
 - 「人材系の訪問はない」「他社訪問もない」のような全体否定があれば「なし」
 - 「複数社来ているが社名不明」なら その他競合=「あり（社名不明）」
 - いずれの記述も無い／聞けていない場合は「不明」
"""
import re
import unicodedata

AGIC = "エイジェック"
COMPANIES = {
    AGIC: ["エイジェック", "AGEC", "AGIC", "エージェック", "エイジック", "エイジエック"],
    "ブレクサネクスト": ["ブレクサ"],
    "ウィルオブワーク": ["ウィルオブ", "ウィル・オブ", "WILLOF"],
    "エヴァンド": ["エヴァンド", "FIDIA", "フィディア"],
    "ネオキャリア": ["ネオキャリア"],
    "スタッフサービス": ["スタッフサービス", "ミラエール"],
    "マンパワー": ["マンパワー", "エムシャイン", "M-Shine", "Mシャイン"],
    "アウトソーシング": ["アウトソーシング"],
    "アウトソーシングテクノロジー": ["アウトソーシングテクノロジー"],
    "ヒラヤマ": ["ヒラヤマ", "平山グループ", "平山(", "平山、", "平山,", "平山 "],
    "フジアルテ": ["フジアルテ"],
    "日研トータルソーシング": ["日研"],
    "UTグループ": ["UTエイム", "UTグループ", "UTコミュニティ"],
    "パソナ": ["パソナ"],
    "パーソル": ["パーソル", "テンプスタッフ"],
    "ランスタッド": ["ランスタッド"],
    "アデコ": ["アデコ"],
    "リクルートスタッフィング": ["リクルートスタッフィング"],
    "ビューティパートナーズ": ["ビューティパートナーズ"],
    "ワールドインテック": ["ワールドインテック"],
    "nextage": ["ネクステージ"],
}

# 社名の後ろ（同じ節の中）で訪問を否定している
#   例: 「エイジェック 来校実績なし」「(エイジェック、ミラエール、マンパワーはない)」
#       「エイジェック、ネオキャリアなどは担当者の記憶上来ていない」
CLAUSE_END = re.compile(r"[。\n┗⸻|]")
NEG_CLAUSE = re.compile(
    r"(来ていない|来てない|来ない|来ていません|来校(実績)?(は|も)?(ない|なし|無い|無し)"
    r"|訪問(実績)?(は|も|が)?(ない|なし|無い|無し)|(は|も)(ない|なし|無い)\)?\s*$|聞いたことがない|知らない)"
)
# 「アウトソーシング」が一般名詞として使われている
GENERIC_OUTSOURCING = re.compile(r"(BPO|業務|事務)[・/]?アウトソーシング|アウトソーシング(事業|業務|サービス|会社|企業|系|型|業界|センター|の仕事)")
# 人材・派遣・BPO系の訪問自体が無い
NONE_PATTERNS = re.compile(
    r"((人材|派遣|BPO|同業|競合|他社|似た(ような)?)(系|会社|企業|業界)?[^。\n]{0,12}?"
    r"(訪問|来校|来て|来る|きて|くる|引き合い|案内|話)[^。\n]{0,6}?(ない|なし|無い|無し|少ない|ほぼない|あまりない))"
    r"|((ほとんど|ほぼ|全然|特に)(こない|来ない|来ていない|きていない|ない))"
    r"|(他社(の)?訪問(も|は)?(ない|なし))"
    r"|^(なし|無し|特になし|ない|来ていない|0社)$"
)
# 社名不明だが訪問はある
UNKNOWN_NAME_PRESENT = re.compile(
    r"((複数|数社|多数|何社|いくつか|[2-9２-９]社|数十社)[^。\n]{0,14}?(訪問|来て|来校|来る|きて|くる|ある|あり|案内|連絡))"
    r"|((人材|派遣|BPO|同業|アウトソー\S{0,4})(系|会社|企業)?[^。\n]{0,10}?(訪問|来校|来て|きて|来る|くる)[^。\n]{0,4}?(あり|ある|いる|多い|増え))"
    r"|((人材|派遣|BPO|アウトソー\S{0,4})[^。\n]{0,6}?(たまに|時々|ときどき)(来る|くる|来て|きて))"
    r"|((人材|派遣|アウトソー\S{0,4})[^。\n]{0,10}?(求人|案内)[^。\n]{0,6}?(ある|あり|多い|多く|来る|くる))"
    r"|((BPO|アウトソーシング)(系|の)?(企業|会社)?(が|も|は)?[^。\n]{0,4}?(多い|多く来|来て|きて|来る|くる))"
)
NOT_HEARD = re.compile(r"(ヒアリングでき(て)?(い)?な|ヒアリングできず|わからな|わからず|分からな|不明|把握(して|でき)(い)?な|教えてもらえ(ず|な)|確認中|聞けず|聞けていな)")


def norm(s: str) -> str:
    return unicodedata.normalize("NFKC", s or "")


def is_outsourcing_company(t: str, start: int, end: int) -> bool:
    """「アウトソーシング」が社名（株式会社アウトソーシング）として使われているか。
    一般名詞（アウトソーシングの働き方／BPOやアウトソーシングの企業 等）は除外する。"""
    before, after = t[max(0, start - 6): start], t[end: end + 6]
    if after.startswith("テクノロジー") or before.endswith(("フォア", "エイジェック", "エンジニアの")):
        return False
    if t.strip(" ・-•*") == "アウトソーシング":
        return True                                   # 欄の回答が社名だけ
    if re.search(r"(株式会社|㈱|\(株\))\s*$", before):
        return True
    # 社名の列挙の中にある（「競合:アウトソーシング、エイジェック」「アウトソーシング、ひらやま」）
    return bool(re.search(r"[、,/／:：]\s*$", before) or re.match(r"\s*[、,/／]", after))


def find_mentions(text: str):
    """[(社名, 否定されているか, 抜粋)]"""
    t = norm(text)
    out = []
    for canon, aliases in COMPANIES.items():
        for a in aliases:
            for m in re.finditer(re.escape(norm(a)), t):
                if canon == "アウトソーシング" and not is_outsourcing_company(t, m.start(), m.end()):
                    continue
                if canon == "ヒラヤマ" and re.match(r"(様|さん|氏)", t[m.end(): m.end() + 2]):
                    continue
                rest = t[m.end(): m.end() + 40]
                e = CLAUSE_END.search(rest)
                clause = rest[: e.start()] if e else rest
                neg = bool(NEG_CLAUSE.search(clause))
                snip = t[max(0, m.start() - 15): m.end() + 30].replace("\n", " ")
                out.append((canon, neg, snip))
    return out


def classify_text(text: str):
    """1つのテキスト片から {社名: True/False} と、全体否定・社名不明あり・未ヒアリング のフラグを返す"""
    t = norm(text)
    companies = {}
    for canon, neg, _ in find_mentions(t):
        # 1回でも肯定的に言及されていれば「あり」
        companies[canon] = companies.get(canon, False) or (not neg)
    lines = [l.strip(" ・-•*") for l in re.split(r"[\n。]", t) if l.strip()]
    none_all = any(NONE_PATTERNS.search(l) for l in lines)
    unknown_present = bool(UNKNOWN_NAME_PRESENT.search(t))
    not_heard = bool(NOT_HEARD.search(t))
    return companies, none_all, unknown_present, not_heard
