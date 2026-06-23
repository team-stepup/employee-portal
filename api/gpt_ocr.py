"""Azure OpenAI (GPT-4o-mini) を使ったマルチモーダル OCR 抽出。

既存の Document Intelligence + 正規表現 (function_app.py の run_zairyu_ocr) を
補完/置換するための「画像 → 構造化 JSON」エンジン。
出力キーは zairyu_submit の confirmedData / parse_zairyu_card_text と互換。

認証は Managed Identity (本番) / az login (ローカル) を DefaultAzureCredential 経由で使う。
API キーは保持しない。
"""
import os
import json
import re
import base64
import logging
from typing import Dict, Any, Optional

_client = None
_claude = None

# Claude (Anthropic) ビジョン OCR — 複雑漢字/数字の読取精度が高い。ANTHROPIC_API_KEY 設定時のみ有効。
CLAUDE_MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-4-6")


def _claude_client():
    """Anthropic クライアント。ANTHROPIC_API_KEY が無ければ None。"""
    global _claude
    if _claude is not None:
        return _claude
    key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if not key:
        return None
    import anthropic
    _claude = anthropic.Anthropic(api_key=key)
    return _claude


def _extract_via_claude(system: str, user_prompt: str, images: list, max_tokens: int = 1000) -> Dict[str, Any]:
    """Claude ビジョンで画像→JSON抽出。ANTHROPIC_API_KEY 必須(無ければ RuntimeError)。"""
    client = _claude_client()
    if client is None:
        raise RuntimeError("ANTHROPIC_API_KEY not set")
    content = [{"type": "text", "text": user_prompt}]
    for b in images[:4]:
        if not b:
            continue
        content.append({"type": "image", "source": {
            "type": "base64", "media_type": "image/jpeg",
            "data": base64.b64encode(b).decode("ascii")}})
    resp = client.messages.create(
        model=CLAUDE_MODEL, max_tokens=max_tokens, temperature=0,
        system=system + " 出力は JSON オブジェクトのみ。前後に説明やコードフェンスを付けない。",
        messages=[{"role": "user", "content": content}],
    )
    txt = "".join(getattr(b, "text", "") for b in resp.content).strip()
    m = re.search(r"\{.*\}", txt, re.S)
    try:
        out = json.loads(m.group(0) if m else txt)
    except Exception:
        logging.error("Claude OCR non-JSON: %s", txt[:500])
        out = {}
    out["engine"] = CLAUDE_MODEL
    try:
        out["_usage"] = {"prompt": resp.usage.input_tokens, "completion": resp.usage.output_tokens}
    except Exception:
        pass
    return out


def _use_claude(model: Optional[str]) -> bool:
    return bool(model) and str(model).startswith("claude")


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
- "zairyuKigen": 在留期間の満了日。"YYYY-MM-DD" 形式 (西暦)。在留資格により読み方が2通り:
   ◆有期の在留資格(定住者・日本人の配偶者等・永住者の配偶者等・技能・技術人文知識国際業務・家族滞在・特定技能・技能実習・留学 等):
      「在留期間(満了日)」欄に実際の満了日が印字されている → その日付を読む。
      この満了日はカード下部「このカードは ○年○月○日 まで有効」(PERIOD OF VALIDITY)と一致するはず(両方確認して同じ値に)。
   ◆無期限の在留資格(永住者・特別永住者・高度専門職2号):
      在留期間欄が「****年**月**日」「無期限」等で満了日が無い → カード下部「このカードは○年○月○日まで有効」を入れる。
      **手がかり: その月日は交付年月日(DATE OF ISSUE)と同じで年だけ7年後**(例 交付2025-09-22→有効期限2032-09-22)。
   ◆共通注意: カードには「許可年月日」(別の日付)も併記され混同しやすい。
      在留期限は必ず「在留期間(満了日)」欄 か カード有効期限 から取り、**許可年月日・交付年月日そのものを在留期限にしてはいけない**。
- "sex": 性別。"M" または "F"。不明なら null。
- "address": 住居地。カード表面/裏面に印字された住所を一字一句そのまま(日本語)。
    **市区町村名を推測・補完・変更しないこと**(例: 読み取りにくくても勝手に別の市名にしない)。読み取れなければ null。

規則:
- 和暦(令和/平成/昭和)は西暦に変換する。例: 令和7年7月4日 → 2025-07-04。
- 読み取れない/記載が無い項目は必ず null。
- 余計な説明やマークダウンは付けず、JSON オブジェクトのみを出力する。
"""


def extract_zairyu_fields(front_bytes: bytes, back_bytes: Optional[bytes] = None, model: Optional[str] = None) -> Dict[str, Any]:
    """在留カード画像(表/任意で裏)から主要項目を抽出して dict で返す。
    model='claude...' なら Claude ビジョン(高精度)、それ以外は Azure OpenAI。省略時は既定(gpt-4o-mini)。"""
    if _use_claude(model):
        try:
            imgs = [front_bytes] + ([back_bytes] if back_bytes else [])
            return _extract_via_claude(ZAIRYU_SYSTEM, ZAIRYU_USER, imgs)
        except Exception as e:
            logging.warning("Claude zairyu OCR failed, fallback to gpt-4o: %s", e)
            model = "gpt-4o"   # Claude不可(キー無し等)→ gpt-4o にフォールバック
    client = _get_client()
    deployment = model if model in ("gpt-4o", "gpt-4o-mini") else os.environ.get("AZURE_OPENAI_DEPLOYMENT", "gpt-4o-mini")
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
    out["engine"] = deployment
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
※1枚に「資格取得等確認通知書(被保険者通知用)」と「雇用保険被保険者証」が並んで写っていることがある。どちらにも同じ被保険者番号が記載されている。
- "koyouNo": 【最重要】被保険者番号。「被保険者番号」欄の数字11桁を必ず読み取り、"XXXX-XXXXXX-X"(4桁-6桁-1桁)のハイフン付きで返す(例 "5107-851970-3")。空欄にしない。
- "koyouName": 被保険者氏名 (カナ/ローマ字 記載のまま)。
- "koyouShutoku": 資格取得年月日 (西暦変換。例 R060704 → 2024-07-04)。
""" + _COMMON_RULES,
    },
    "shaken": {
        "system": "あなたは日本の自動車検査証(車検証)を読み取る高精度な情報抽出エンジンです。",
        "user": """この自動車検査証の画像から以下を抽出し、JSONのみを返してください。
- "carNumber": 自動車登録番号または車両番号 (例 "浜松 330 あ 1234" → "浜松330あ1234")。
- "carMaker": 「車名」欄の値=メーカー名 (例 "トヨタ","ホンダ","スズキ","ダイハツ")。車検証の「車名」欄はメーカー名が入る。
- "carModel": 「型式」欄の値 (例 "DBA-JH1","5BA-MN71S")。「原動機の型式」(例 "S07A") とは別物なので混同しない。
- "carName": 通称名=車種名 (例 "N-WGN","ワゴンR","タント","ヴィッツ")。
  ★重要: 通称名は車検証に記載が無いことが多い。「車名」欄(メーカー名)や「型式」をそのまま carName に入れてはいけない。
  通称名の記載が無い場合は、メーカー名(carMaker)と型式(carModel)から一般に流通している通称名を推測して入れる
  (例: ホンダ+DBA-JH1 → "N-WGN"、スズキ+DBA-MH34S → "ワゴンR"、ダイハツ+LA600S → "タント")。
  メーカー・型式から通称名がどうしても特定できない場合のみ null。
- "carDisplacement": 総排気量 (例 "0.65L" → "650cc"、"1.49L" → "1490cc" のようにcc表記。原付・軽二輪も同様)。
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
- "carName": 契約車両(被保険自動車)欄の「車名・仕様」に書かれた通称名=車種名 (例 "N-WGN","ワゴンR")。
  この欄に型式(例 "JH1")しか記載が無い場合は、その型式から一般に流通している通称名を推測して入れる(例 "JH1" → "N-WGN")。
  どうしても特定できない場合のみ null。
""" + _COMMON_RULES,
    },
    "bank": {
        "system": "あなたは日本の銀行通帳・キャッシュカードを読み取る高精度な情報抽出エンジンです。",
        "user": """この通帳(見開き)またはキャッシュカードの画像から以下を抽出し、JSONのみを返してください。
- "bankName": 金融機関名 (例 "静岡銀行","浜松磐田信用金庫","ゆうちょ銀行")。
- "bankCode": 金融機関コード(4桁の数字)。キャッシュカードや通帳に記載があれば必ず読み取る(例 "0149","1503")。無ければ null。
- "bankBranch": 支店名 (例 "豊田支店")。
- "branchCode": 店番(支店番号。通常3桁の数字)。キャッシュカードや通帳に記載があれば必ず読み取る(例 "001","238")。無ければ null。
- "bankAccount": 口座番号 (数字のみ7桁)。
- "bankMeigi": 口座名義人の氏名(カタカナ。半角/全角どちらでも、姓名のカナのみ)。
  ★重要: 名義カナの末尾や近くにある **「フツウ」(普通)/「トウザ」(当座)/「チョチク」(貯蓄)/「貯蓄」等は預金種目を表す語であり、名義ではない**ので bankMeigi に含めない。氏名のカナだけを返す。
  外国人のカナ名(例 "クルパルサ サヤカ ケルシー カミオカ")も丁寧に読み取る。
※キャッシュカードには「金融機関コード(4桁)-店番(3桁)-科目-口座番号(7桁)」のように数字が並んで刻印(エンボス凸文字)されていることが多い(例 "1503-111-21-5129462")。
  その場合、左から 4桁=bankCode、次の3桁=branchCode、末尾7桁=bankAccount として正確に分解する。
  **エンボス(銀色の凸文字)やキャラクター/イラストで一部が隠れて見えにくくても、数字の並びから店番(4桁コードの直後の3桁)を必ず推定して branchCode に入れる**。ハイフンや空白で区切られていることが多い。
※ゆうちょ銀行の場合: カードに『記号(5桁) 番号(最大8桁)』が記載されている。
  bankAccount には**この「記号-番号」をそのまま**入れる(例 "12340-56980741")。店番・口座番号への変換はしない(アプリ側で変換)。
  bankBranch・bankCode・branchCode は空(null)でよい。
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
    "classify": {
        "system": "あなたは日本の各種証明書類の画像を見て、その書類の種別を判定する分類エンジンです。",
        "user": """この画像が次のどの書類かを判定し、JSONのみを返してください。
{"docType": "<下記のいずれか1つのコード>"}
- "zairyu"     : 在留カード (RESIDENCE CARD / 日本国政府 在留カード)
- "license"    : 運転免許証
- "koyou"      : 雇用保険被保険者証 / 資格取得等確認通知書
- "shaken"     : 自動車検査証 (車検証 / 自動車検査証記録事項)
- "jibaiseki"  : 自動車損害賠償責任保険証明書 (自賠責)
- "nini"       : 任意の自動車保険(任意保険)の証券・契約内容書類
- "bank"       : 預金通帳・キャッシュカード・口座情報が分かる書類
- "mynumber"   : マイナンバーカード / 個人番号通知カード
- "other"      : 上記のいずれにも当てはまらない、または判別できない
最も適切なコードを1つだけ返す。確信が持てない場合は "other"。
""" + _COMMON_RULES,
    },
    "repair": {
        "system": "あなたは日本の自動車整備・修理・車両交換の請求書/整備票/作業伝票を読み取り、車両ごとの明細を構造化する高精度な情報抽出エンジンです。",
        "user": """この自動車の整備・修理・点検・部品交換の請求書/整備票/納品書/作業伝票の画像から、車両ごと(伝票ごと)の明細を漏れなく抽出し、JSONのみを返してください。
1枚に複数台・複数伝票が含まれることがあるので、必ず配列で全件返す。
形式: {"records": [ {明細1}, {明細2}, ... ]}
各明細の項目:
- "date": 実施日/作業日/入庫日/請求日 ("YYYY-MM-DD")。和暦(R6.7.4等)は西暦に変換。
- "plate": 車両のナンバープレート(登録番号)。「浜松 480 あ 12-34」のような地名+分類番号+ひらがな+一連番号を空白・ハイフン無しで連結("浜松480あ1234")。読める範囲で。
- "no2": 一連指定番号(プレート下段の大きい数字。最大4桁。例 "12-34"→"1234"、"・1-23"→"123")。プレートから分かれば必ず入れる。
- "vehicleName": 車名/車種(例 "ハイゼット","N-WGN","タント")。記載があれば。
- "mileage": 走行距離/キロ数(数字のみ、カンマ無し)。記載があれば。
- "detail": 作業内容・品名の要約。1台に複数作業があれば「、」で連結して1つにまとめる(例 "エンジンオイル交換、オイルエレメント交換、12ヶ月点検")。
- "amount": その車両(伝票)の金額。税込合計。数字のみ(カンマ・円記号なし。例 27500)。
重要な注意:
- ヘッダー/フッター(自社宛名・業者の会社名/住所/電話/FAX/登録番号、小計・消費税・総合計の合算行)は明細(records)に含めない。
- 1台に複数作業がある場合は1明細にまとめ、detailを連結・amountはその車両分の合計にする。
- ナンバーが読み取れない明細でも、日付・内容・金額が分かれば records に含める(plate/no2は分かる範囲で、無ければ null)。
- 該当する明細が1件も無ければ {"records": []} を返す。
""" + _COMMON_RULES,
        "max_tokens": 4000,
    },
    "shohyo": {
        "system": "あなたは日本の領収書・レシート・請求書を読み取り、電子帳簿保存法の検索要件(取引年月日・取引金額・取引先)を正確に抽出する高精度な情報抽出エンジンです。",
        "user": """このレシート/領収書/請求書/納品書の画像から以下を抽出し、JSONのみを返してください。
- "torihikiDate": 取引年月日。レシート/領収書は支払日(発行日)、請求書は請求日(発行日)。"YYYY-MM-DD"。
- "amount": 取引金額。税込の支払総額(合計・お会計・ご請求額)。数字のみ、カンマや円記号・マイナスは付けない (例 11000)。割引後の最終支払額を採用する。
- "torihikisaki": 取引先名。発行した店舗名・会社名(屋号や正式名称)。支店名まで分かれば含める。自社(支払者)側ではなく相手先を取る。
- "invoiceNo": 適格請求書発行事業者の登録番号。"T"+13桁の数字 (例 "T1234567890123")。記載が無ければ null。
- "docKind": 書類種別。"レシート" / "領収書" / "請求書" / "納品書" のいずれか。判別できなければ null。
注意:
- 金額は最も大きな「合計/総額/ご請求金額」を優先し、内訳(小計・消費税のみ)や預り金・釣銭は採用しない。
- 取引先は電話番号やロゴ近辺の店名を優先。住所だけの行は取引先名にしない。
""" + _COMMON_RULES,
    },
}


def extract_doc_fields(doc_type: str, images: list, model: Optional[str] = None) -> Dict[str, Any]:
    """書類タイプ別の汎用抽出。images は bytes のリスト。model='claude...' なら Claude、それ以外 Azure。"""
    spec = DOC_PROMPTS.get(doc_type)
    if not spec:
        raise ValueError(f"unknown doc_type: {doc_type}")
    mt = int(spec.get("max_tokens", 0) or 0)
    if _use_claude(model):
        try:
            return _extract_via_claude(spec["system"], spec["user"], images, mt or 1000)
        except Exception as e:
            logging.warning("Claude doc OCR failed (%s), fallback to gpt-4o: %s", doc_type, e)
            model = "gpt-4o"
    client = _get_client()
    deployment = model if model in ("gpt-4o", "gpt-4o-mini") else os.environ.get("AZURE_OPENAI_DEPLOYMENT", "gpt-4o-mini")
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
        max_tokens=mt or 800,
        response_format={"type": "json_object"},
    )
    txt = (resp.choices[0].message.content or "").strip()
    try:
        out = json.loads(txt)
    except Exception:
        logging.error("GPT doc OCR returned non-JSON: %s", txt[:500])
        out = {}
    out["engine"] = deployment
    try:
        out["_usage"] = {"prompt": resp.usage.prompt_tokens, "completion": resp.usage.completion_tokens}
    except Exception:
        pass
    return out
