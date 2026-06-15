"""Azure OpenAI (GPT-4o-mini) を使ったマルチモーダル OCR 抽出。

既存の Document Intelligence + 正規表現 (function_app.py の run_zairyu_ocr) を
補完/置換するための「画像 → 構造化 JSON」エンジン。
出力キーは zairyu_submit の confirmedData / parse_zairyu_card_text と互換。

認証は Managed Identity (本番) / az login (ローカル) を DefaultAzureCredential 経由で使う。
API キーは保持しない。
"""
import os
import json
import base64
import logging
from typing import Dict, Any, Optional

_client = None


def _get_client():
    """AzureOpenAI クライアントを Managed Identity (AAD トークン) で取得しキャッシュ。"""
    global _client
    if _client is not None:
        return _client
    from openai import AzureOpenAI
    endpoint = os.environ.get("AZURE_OPENAI_ENDPOINT", "").strip()
    if not endpoint:
        raise RuntimeError("AZURE_OPENAI_ENDPOINT not set")
    api_version = os.environ.get("AZURE_OPENAI_API_VERSION", "2024-10-21")
    api_key = os.environ.get("AZURE_OPENAI_KEY", "").strip()
    if api_key:
        # API キー認証 (設定済みなら優先・ロール付与不要)
        _client = AzureOpenAI(azure_endpoint=endpoint, api_version=api_version, api_key=api_key)
    else:
        # Managed Identity (AAD) フォールバック — 要 "Cognitive Services OpenAI User" ロール
        from azure.identity import DefaultAzureCredential, get_bearer_token_provider
        token_provider = get_bearer_token_provider(
            DefaultAzureCredential(), "https://cognitiveservices.azure.com/.default")
        _client = AzureOpenAI(azure_endpoint=endpoint, api_version=api_version, azure_ad_token_provider=token_provider)
    return _client


def _data_url(image_bytes: bytes) -> str:
    return "data:image/jpeg;base64," + base64.b64encode(image_bytes).decode("ascii")


# 在留カード抽出プロンプト。出力キーは既存スキーマと一致させる。
ZAIRYU_SYSTEM = (
    "あなたは日本の在留カード(Residence Card)を読み取る高精度な情報抽出エンジンです。"
    "画像から記載項目を正確に読み取り、指定された JSON のみを返します。"
    "推測で値を作らず、読み取れない項目は null にします。"
)

ZAIRYU_USER = """この在留カードの画像から以下の項目を抽出し、JSONオブジェクトのみを返してください。
表面・裏面の両方が渡された場合は両方を参照してください。

抽出キーと規則:
- "cardNumber": 在留カード番号。英字2桁+数字8桁+英字2桁 (例 AB12345678CD)。
- "name": 氏名。在留カードはローマ字表記なので大文字ローマ字のフルネームをそのまま (例 "MORI ROBERT YUJI")。
- "birthday": 生年月日。"YYYY-MM-DD" 形式 (西暦)。
- "nationality": 国籍・地域。日本語のカタカナ/漢字表記に正規化 (例 "ブラジル","フィリピン","ペルー","中国")。カード上が英語表記でも日本語名に変換。
- "zairyuShikaku": 在留資格。日本語表記そのまま (例 "永住者","技能","技術・人文知識・国際業務","定住者","特定技能1号","家族滞在")。
- "zairyuKigen": 在留期間(満了日)。"YYYY-MM-DD" 形式 (西暦)。
    ※「永住者」等で在留期間の満了日が空欄/無期限の場合は、カード自体の「有効期間の満了する日」を代わりに入れる。
- "sex": 性別。"M" または "F"。不明なら null。
- "address": 住居地。記載があれば日本語のまま。無ければ null。

規則:
- 和暦(令和/平成/昭和)は西暦に変換する。例: 令和7年7月4日 → 2025-07-04。
- 読み取れない/記載が無い項目は必ず null。
- 余計な説明やマークダウンは付けず、JSON オブジェクトのみを出力する。
"""


def extract_zairyu_fields(front_bytes: bytes, back_bytes: Optional[bytes] = None) -> Dict[str, Any]:
    """在留カード画像(表/任意で裏)から主要項目を GPT-4o-mini で抽出して dict で返す。"""
    client = _get_client()
    deployment = os.environ.get("AZURE_OPENAI_DEPLOYMENT", "gpt-4o-mini")
    content = [{"type": "text", "text": ZAIRYU_USER},
               {"type": "image_url", "image_url": {"url": _data_url(front_bytes), "detail": "high"}}]
    if back_bytes:
        content.append({"type": "image_url", "image_url": {"url": _data_url(back_bytes), "detail": "high"}})
    resp = client.chat.completions.create(
        model=deployment,
        messages=[
            {"role": "system", "content": ZAIRYU_SYSTEM},
            {"role": "user", "content": content},
        ],
        temperature=0,
        max_tokens=800,
        response_format={"type": "json_object"},
    )
    txt = (resp.choices[0].message.content or "").strip()
    try:
        out = json.loads(txt)
    except Exception:
        logging.error("GPT OCR returned non-JSON: %s", txt[:500])
        out = {}
    out["engine"] = "gpt-4o-mini"
    # 使用トークン (コスト把握用)
    try:
        out["_usage"] = {
            "prompt": resp.usage.prompt_tokens,
            "completion": resp.usage.completion_tokens,
        }
    except Exception:
        pass
    return out


# ============================================================
# 汎用 書類抽出 (yukyu-app 社員ファイル一括読み取り用)
# 出力キーは yukyu-app (社員List) のフィールドキーに揃える。
# ============================================================
_COMMON_RULES = """
共通規則:
- 和暦(令和/平成/昭和)は西暦に変換する。例: 令和6年7月4日 → 2024-07-04 / R060704 → 2024-07-04 / 3 580830 → 1983-08-30 (元号コード 1=明治 2=大正 3=昭和 4=平成 5=令和)。
- 日付はすべて "YYYY-MM-DD" 形式。
- 読み取れない/記載が無い項目は必ず null。推測で値を作らない。
- 余計な説明やマークダウンは付けず、JSON オブジェクトのみを出力する。
"""

DOC_PROMPTS: Dict[str, Dict[str, str]] = {
    "license": {
        "system": "あなたは日本の運転免許証を読み取る高精度な情報抽出エンジンです。",
        "user": """この運転免許証の画像から以下を抽出し、JSONのみを返してください。
- "licenseNo": 免許証番号 (第XXXXXXXXXXXX号 の12桁数字のみ)。
- "licenseType": 種類欄で最も上位の種類 (例 "普通","準中型","中型","大型","原付","普自二","大自二")。
- "licenseGetDate": 「二・小・原」「他」「二種」の取得年月日のうち最も古い日付。
- "licenseExpiry": 有効期限 (「〇〇年〇〇月〇〇日まで有効」の日付)。
- "birthday": 生年月日。
- "name": 氏名 (漢字/カナ表記のまま)。
""" + _COMMON_RULES,
    },
    "koyou": {
        "system": "あなたは日本の雇用保険被保険者証・資格取得等確認通知書を読み取る高精度な情報抽出エンジンです。",
        "user": """この雇用保険被保険者証(または資格取得等確認通知書)の画像から以下を抽出し、JSONのみを返してください。
- "koyouNo": 被保険者番号 ("XXXX-XXXXXX-X" のハイフン付き形式に整形)。
- "koyouName": 被保険者氏名 (カナ/ローマ字 記載のまま)。
- "koyouShutoku": 資格取得年月日 (西暦変換。例 R060704 → 2024-07-04)。
""" + _COMMON_RULES,
    },
    "shaken": {
        "system": "あなたは日本の自動車検査証(車検証)を読み取る高精度な情報抽出エンジンです。",
        "user": """この自動車検査証の画像から以下を抽出し、JSONのみを返してください。
- "carNumber": 自動車登録番号または車両番号 (例 "浜松 330 あ 1234" → "浜松330あ1234")。
- "carMaker": 車名欄の値 (例 "トヨタ","ホンダ","スズキ")。
- "carName": 通称名/型式 (分かれば。無ければ null)。
- "carDisplacement": 総排気量 (例 "1.49L" → "1490cc" のようにcc表記。原付・軽二輪も同様)。
- "carFirstReg": 初度登録年月 (西暦 "YYYY-MM")。
- "shakenExpiry": 有効期間の満了する日。
""" + _COMMON_RULES,
    },
    "jibaiseki": {
        "system": "あなたは日本の自動車損害賠償責任保険証明書(自賠責)を読み取る高精度な情報抽出エンジンです。",
        "user": """この自賠責保険証明書の画像から以下を抽出し、JSONのみを返してください。
- "jibaisekiCompany": 保険会社名。
- "jibaisekiNo": 証明書番号。
- "jibaisekiExpiry": 保険期間の末日 (満了日)。
""" + _COMMON_RULES,
    },
    "nini": {
        "system": "あなたは日本の自動車任意保険の証券・契約内容書類を読み取る高精度な情報抽出エンジンです。",
        "user": """この任意保険の証券画像から以下を抽出し、JSONのみを返してください。
- "niniCompany": 保険会社名。
- "niniNo": 証券番号。
- "niniStartDate": 保険期間の開始日。
- "niniExpiry": 保険期間の満了日。
- "taijin": 対人賠償の保険金額 (例 "無制限")。
- "taibutsu": 対物賠償の保険金額 (例 "無制限","1000万円")。
""" + _COMMON_RULES,
    },
    "bank": {
        "system": "あなたは日本の銀行通帳・キャッシュカードを読み取る高精度な情報抽出エンジンです。",
        "user": """この通帳(見開き)またはキャッシュカードの画像から以下を抽出し、JSONのみを返してください。
- "bankName": 金融機関名 (例 "静岡銀行","浜松磐田信用金庫","ゆうちょ銀行")。
- "bankBranch": 支店名 (例 "豊田支店")。
- "bankAccount": 口座番号 (数字のみ7桁)。
- "bankMeigi": 口座名義 (カタカナ。記載のまま)。
""" + _COMMON_RULES,
    },
    "mynumber": {
        "system": "あなたは日本のマイナンバーカード・通知カードを読み取る情報抽出エンジンです。",
        "user": """このマイナンバーカード(または通知カード)の画像から以下を抽出し、JSONのみを返してください。
- "name": 氏名 (漢字)。
- "birthday": 生年月日。
- "address": 住所。
- "yubin": 郵便番号 ("XXX-XXXX" 形式。記載があれば)。
※個人番号(12桁)は抽出しない。
""" + _COMMON_RULES,
    },
}


def extract_doc_fields(doc_type: str, images: list) -> Dict[str, Any]:
    """書類タイプ別の汎用抽出。images は bytes のリスト (表/裏/複数ページ)。"""
    spec = DOC_PROMPTS.get(doc_type)
    if not spec:
        raise ValueError(f"unknown doc_type: {doc_type}")
    client = _get_client()
    deployment = os.environ.get("AZURE_OPENAI_DEPLOYMENT", "gpt-4o-mini")
    content = [{"type": "text", "text": spec["user"]}]
    for b in images[:4]:
        content.append({"type": "image_url", "image_url": {"url": _data_url(b), "detail": "high"}})
    resp = client.chat.completions.create(
        model=deployment,
        messages=[
            {"role": "system", "content": spec["system"]},
            {"role": "user", "content": content},
        ],
        temperature=0,
        max_tokens=800,
        response_format={"type": "json_object"},
    )
    txt = (resp.choices[0].message.content or "").strip()
    try:
        out = json.loads(txt)
    except Exception:
        logging.error("GPT doc OCR returned non-JSON: %s", txt[:500])
        out = {}
    out["engine"] = "gpt-4o-mini"
    try:
        out["_usage"] = {"prompt": resp.usage.prompt_tokens, "completion": resp.usage.completion_tokens}
    except Exception:
        pass
    return out
