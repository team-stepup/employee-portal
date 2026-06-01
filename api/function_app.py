# -*- coding: utf-8 -*-
"""employee-portal Backend API (Azure Functions, Python v2 programming model).

エンドポイント:
- POST /api/auth/login                  3要素照合 + JWT 発行
- GET  /api/profile                     本人プロフィール
- GET  /api/maebarai/dates              選択可能な金曜日リスト
- POST /api/maebarai/apply              前払い申請
- GET  /api/maebarai/history            申請履歴

認証: Managed Identity → SharePoint REST API (Sites.Selected on /sites/PowerApps)
JWT: HS256 + 24h 有効。秘密鍵は App Setting JWT_SECRET から取得。
"""
import base64
import datetime as _dt
import hashlib
import hmac
import json
import logging
import os
import re
import time
from typing import Any, Dict, List, Optional, Tuple

import azure.functions as func
import jwt
import requests
from azure.identity import DefaultAzureCredential
from azure.data.tables import TableServiceClient, UpdateMode
from azure.core.exceptions import ResourceNotFoundError

# ====== 設定 ======
SP_HOST = "teamstepupcom774.sharepoint.com"
SITE_URL = f"https://{SP_HOST}/sites/PowerApps"
SITE_TEAMSTEPUP = f"https://{SP_HOST}/sites/TeamStepup"  # 社員ファイル所在
SHAINFILE_ROOT = "/sites/TeamStepup/Shared Documents/社員ファイル"
SP_RESOURCE = f"https://{SP_HOST}"

LIST_SHAIN = "05a5a986-5958-4664-bb0d-7c1cfaedf845"
LIST_HAKENSAKI = "2a2ab839-9dcd-4385-a72b-f20449ceddc4"
LIST_KAISHA_KYUJITSU = "37973404-3c0a-4e02-a64a-6d15b633d072"
LIST_MAEBARAI_SHINSEI = "9e84553c-73ec-405b-862f-6caaae65900f"
LIST_MAEBARAI_OLD = "d34be315-365a-460f-976e-fb8da6977cec"

# 社員 List フィールド (Entity property names で API リクエスト時に使う)
F_SHAIN_NO = "OData__x793e__x54e1__x756a__x53f7_"
F_SHAIN_NAME = "Title"
F_BUKA = "OData__x90e8__x8ab2_"
F_BIRTHDAY = "OData__x751f__x5e74__x6708__x65e5_"
F_TEL = "OData__x643a__x5e2f__x96fb__x8a71_"
F_TAISHA_DATE = "OData__x6d3e__x9063__x5148__x9000__x79"
F_ZAIYOKU = "OData__x63a1__x7528__x7a2e__x5225_syub"
F_GINKO = "OData__x7d66__x4e0e__xff1a__x9280__x88"
F_SHITEN = "OData__x7d66__x4e0e__xff1a__x652f__x5e"
F_KOUZA = "OData__x7d66__x4e0e__xff1a__x53e3__x5e"
F_MEIGI = "OData__x53e3__x5ea7__x6c0f__x540d__x30"
F_ZAIRYU_NAME = "OData__x7279__x8a18__x4e8b__x9805__xff"  # 特記事項１ = 在留カード氏名 (英字フルネーム)
F_KOKUSEKI = "OData__x672c__x0028__x56fd__x0029__x7c"  # 本(国)籍
F_ZAIRYU_SHIKAKU = "OData__x5728__x7559__x8cc7__x683c_"  # 在留資格 (Text)
F_ZAIRYU_KIGEN = "OData__x5728__x7559__x671f__x9650_"    # 在留期限 (DateTime)
F_ZAIRYU_BIKO = "OData__x5728__x7559__xff1a__x5099__x80"  # 在留：備考1 (Text) — 在留カード番号格納用
F_TSUKIN_OLD = "OData__x901a__x52e4__x65b9__x6cd5_tsuk"  # 通勤方法tsukinhouhou (Choice) — 車kuruma/送迎/バイク/自転車徒歩
F_TSUKIN_NEW = "OData__x65b0__x901a__x52e4__x65b9__x6c"  # 新通勤方法 (Text) — 送迎/自通

# 免許証関連フィールド
F_MENKYO_NUMBER = "OData__x514d__x8a31__x8a3c__x756a__x53"   # 免許証番号 (Number)
F_MENKYO_TAIKEN = "OData__x514d__x8a31__x53d6__x5f97__x5e"   # 免許取得年月日 (DateTime)
F_MENKYO_KIGEN = "OData__x514d__x8a31__x8a3c__x6709__x52"    # 免許証有効期限 (DateTime)
F_MENKYO_TYPE = "OData__x514d__x8a31__x306e__x7a2e__x98"     # 免許の種類 (Choice)

# 車検証関連フィールド
F_CAR_NAME = "OData__x8eca__x540d_"                          # 車名 (Text)
F_CAR_MAKER = "OData__x8eca__x4e21__x30e1__x30fc__x30"       # 車両メーカー名 (Choice)
F_CAR_NUMBER = "OData__x8eca__x4e21__x30ca__x30f3__x30"      # 登録番号 (Text)
F_HAIKIRYO = "OData__x6392__x6c17__x91cf_"                   # 排気量 (Text)
F_SHONENDO = "OData__x521d__x5e74__x5ea6__x767b__x93"        # 初年度登録 (Text)
F_SHAKEN_KIGEN = "OData__x8eca__x691c__x6e80__x4e86__x65"    # 車検満了日 (DateTime)
# 自賠責保険関連フィールド
F_JIBAI_KAISHA = "OData__x81ea__x8ce0__x8cac__x4fdd__x96"    # 自賠責保険会社 (Choice)
F_JIBAI_KIGEN = "OData__x81ea__x8ce0__x8cac__x3000__x6e"     # 自賠責満了日 (DateTime)
F_JIBAI_SHOKEN = "OData__x81ea__x8ce0__x8cac__x8a3c__x52"    # 自賠責証券番号 (Text)
# 任意保険関連フィールド
F_NINI_KAISHA = "OData__x81ea__x52d5__x8eca__x4efb__x61"     # 自動車任意保険会社 (Choice)
F_NINI_KAISHI = "OData__x4efb__x610f__x4fdd__x967a__x95"     # 任意保険開始日 (DateTime)
F_NINI_KIGEN = "OData__x4efb__x610f__x4fdd__x967a__x6e"      # 任意保険満了日 (DateTime)
F_NINI_SHOKEN = "OData__x4efb__x610f__x4fdd__x967a__x8a"     # 任意保険証券番号 (Text)

# ポータル PIN (4桁) のハッシュ格納フィールド (Note)
F_PORTAL_PIN = "OData__x30dd__x30fc__x30bf__x30eb_PIN"


def hash_pin(shain_no: int, pin: str) -> str:
    """PIN を社員番号 + サーバ秘密鍵で HMAC-SHA256 ハッシュ化。
    社員番号を salt に含めることで、同じ PIN でも社員ごとに異なるハッシュになる。"""
    secret = (JWT_SECRET or "portal-pin-fallback").encode("utf-8")
    msg = f"{shain_no}:{pin}".encode("utf-8")
    return hmac.new(secret, msg, hashlib.sha256).hexdigest()


def verify_pin(shain_no: int, pin: str, stored_hash: Optional[str]) -> bool:
    if not stored_hash:
        return False
    return hmac.compare_digest(hash_pin(shain_no, pin), stored_hash.strip())


def commutes_by_car(emp: Dict[str, Any]) -> bool:
    """通勤方法フィールドから車/バイク/自通かを判定。"""
    old = str(emp.get(F_TSUKIN_OLD) or "")
    new = str(emp.get(F_TSUKIN_NEW) or "")
    combined = old + " " + new
    for kw in ("車", "kuruma", "バイク", "baiku", "自通"):
        if kw in combined:
            return True
    return False


def guess_lang_from_kokuseki(kokuseki: Optional[str]) -> str:
    """本(国)籍 から推奨言語を推定。
    日本/日本人/中国 → ja
    ブラジル/ペルー/ボリビア/アルゼンチン/パラグアイ → pt
    その他 (フィリピン等) → en
    """
    if not kokuseki:
        return "ja"
    k = str(kokuseki).strip()
    # 日本語ネイティブ + 中国 (中国人は日本語と英語が選択肢、デフォは日本語)
    if "日本" in k or "中国" in k:
        return "ja"
    # ポルトガル/スペイン語圏 (南米)
    pt_countries = ("ブラジル", "ペルー", "ボリビア", "アルゼンチン", "パラグアイ", "ウルグアイ", "コロンビア", "ベネズエラ", "チリ", "エクアドル")
    for c in pt_countries:
        if c in k:
            return "pt"
    # その他外国籍 (フィリピン・東南アジア等) は英語
    return "en"

# 派遣先情報040623 List フィールド
F_HAKEN_BANGO = "OData__x756a__x53f7_"  # 番号
F_HAKEN_TITLE = "Title"  # 部課
F_HAKEN_KAISYAMEI = "OData__x6d3e__x9063__x5148__x4f1a__x790"  # 派遣先会社名
F_HAKEN_HARAIBI = "OData__x652f__x6255__x65e5__xff08__x30"  # 支払日（給料）

# 会社休日 List フィールド
F_KYUJITSU_DATE = "OData__x65e5__x4ed8_"
F_KYUJITSU_NAME = "OData__x540d__x79f0_"
F_KYUJITSU_KIND = "OData__x7a2e__x5225_"

# 前払い申請 List フィールド (POST body には OData_ プレフィックス無しで使う)
P_SHAIN_NO = "_x793e__x54e1__x756a__x53f7_"
P_KINGAKU = "_x91d1__x984d_"
P_KIBOUBI = "_x5e0c__x671b__x65e5_"
P_RIYUU = "_x7533__x8acb__x7406__x7531_"
P_HAKENSAKI = "_x6d3e__x9063__x5148_"
P_TANTOU = "_x62c5__x5f53__x8005_"
P_STATUS = "Status"

# JWT
JWT_SECRET = os.environ.get("JWT_SECRET", "")
JWT_ALG = "HS256"
JWT_EXP_HOURS = 24

# Brute force 防御
LOCKOUT_MAX_ATTEMPTS = 5
LOCKOUT_DURATION_MINUTES = 30
LOCKOUT_TABLE = "loginattempts"

# レート制限 (社員番号 + エンドポイント別)
RATE_LIMIT_WINDOW_SECONDS = 60
RATE_LIMITS = {
    "apply": 10,       # 申請 10/min
    "cancel": 10,      # 取消 10/min
    "zairyu": 5,       # 在留カード提出 5/min (画像アップロード重め)
}
_storage_conn = os.environ.get("AzureWebJobsStorage", "")
_table_client_cache = None


def _get_lockout_table():
    """Login attempt 追跡用 Storage Table を返す。"""
    global _table_client_cache
    if _table_client_cache is not None:
        return _table_client_cache
    if not _storage_conn:
        return None
    try:
        svc = TableServiceClient.from_connection_string(_storage_conn)
        svc.create_table_if_not_exists(table_name=LOCKOUT_TABLE)
        _table_client_cache = svc.get_table_client(LOCKOUT_TABLE)
        return _table_client_cache
    except Exception as e:
        logging.warning(f"lockout table init failed: {e}")
        return None


def check_lockout(shain_no: int) -> Optional[int]:
    """ロックされていれば残り秒数を返す。ロックされていなければ None。"""
    table = _get_lockout_table()
    if not table:
        return None
    try:
        entity = table.get_entity(partition_key="auth", row_key=str(shain_no))
        locked_until = entity.get("lockedUntil")
        if locked_until:
            if isinstance(locked_until, str):
                locked_until = _dt.datetime.fromisoformat(locked_until.replace("Z", "+00:00"))
            now = _dt.datetime.now(_dt.timezone.utc)
            if locked_until > now:
                return int((locked_until - now).total_seconds())
    except ResourceNotFoundError:
        return None
    except Exception as e:
        logging.warning(f"check_lockout failed: {e}")
    return None


def record_failed_attempt(shain_no: int):
    """失敗を記録。閾値に達したらロックアウトを設定。"""
    table = _get_lockout_table()
    if not table:
        return
    now = _dt.datetime.now(_dt.timezone.utc)
    try:
        entity = table.get_entity(partition_key="auth", row_key=str(shain_no))
        count = int(entity.get("count", 0)) + 1
    except ResourceNotFoundError:
        count = 1
    locked_until = None
    if count >= LOCKOUT_MAX_ATTEMPTS:
        locked_until = now + _dt.timedelta(minutes=LOCKOUT_DURATION_MINUTES)
    entity = {
        "PartitionKey": "auth",
        "RowKey": str(shain_no),
        "count": count,
        "lastFailureAt": now.isoformat(),
        "lockedUntil": locked_until.isoformat() if locked_until else "",
    }
    try:
        table.upsert_entity(entity, mode=UpdateMode.REPLACE)
    except Exception as e:
        logging.warning(f"record_failed_attempt failed: {e}")


def clear_attempts(shain_no: int):
    """成功時にロックアウトカウンタをクリア。"""
    table = _get_lockout_table()
    if not table:
        return
    try:
        table.delete_entity(partition_key="auth", row_key=str(shain_no))
    except ResourceNotFoundError:
        pass
    except Exception as e:
        logging.warning(f"clear_attempts failed: {e}")


def check_rate_limit(shain_no: int, endpoint: str) -> Optional[int]:
    """社員番号 + エンドポイント別レート制限。超過なら待機秒数を返す。"""
    limit = RATE_LIMITS.get(endpoint)
    if not limit:
        return None
    table = _get_lockout_table()
    if not table:
        return None
    now = _dt.datetime.now(_dt.timezone.utc)
    rk = f"{endpoint}:{shain_no}"
    try:
        entity = table.get_entity(partition_key="ratelimit", row_key=rk)
        first_at = entity.get("firstAt")
        if isinstance(first_at, str):
            first_at = _dt.datetime.fromisoformat(first_at.replace("Z", "+00:00"))
        count = int(entity.get("count", 0))
        elapsed = (now - first_at).total_seconds()
        if elapsed < RATE_LIMIT_WINDOW_SECONDS:
            # 同じウィンドウ内
            if count >= limit:
                return int(RATE_LIMIT_WINDOW_SECONDS - elapsed)
            entity["count"] = count + 1
            table.upsert_entity(entity, mode=UpdateMode.REPLACE)
            return None
        # 新しいウィンドウ
        entity["firstAt"] = now.isoformat()
        entity["count"] = 1
        table.upsert_entity(entity, mode=UpdateMode.REPLACE)
        return None
    except ResourceNotFoundError:
        table.upsert_entity({
            "PartitionKey": "ratelimit",
            "RowKey": rk,
            "firstAt": now.isoformat(),
            "count": 1,
        }, mode=UpdateMode.REPLACE)
        return None
    except Exception as e:
        logging.warning(f"check_rate_limit failed: {e}")
        return None  # fail-open: 内部エラー時は許可


def _validate_shainfile_path(folder_url: str) -> bool:
    """SP のフォルダ URL が「社員ファイル」配下にあることを検証 (Path Traversal 防止)。"""
    if not folder_url:
        return False
    expected_prefix = SHAINFILE_ROOT + "/"
    if not folder_url.startswith(expected_prefix):
        return False
    # 不正なパス要素を除外
    suffix = folder_url[len(expected_prefix):]
    if ".." in suffix or "//" in suffix:
        return False
    # 制御文字・絶対パス・改行などの除外
    if any(ch in folder_url for ch in ('\x00', '\n', '\r')):
        return False
    return True


def _safe_filename(name: str) -> str:
    """ファイル名から危険文字を除去 (パス区切り・制御文字)。"""
    safe = re.sub(r'[\\/\x00-\x1f<>:"|?*]', '', name or '')
    return safe[:200]

# 金額: ドロップダウン候補 (10,000〜100,000) + 「その他」手動入力 (〜130,000、10,000円刻み)
AMOUNT_MIN = 10000
AMOUNT_MAX = 130000
AMOUNT_STEP = 10000
AMOUNT_DROPDOWN = [10000, 20000, 30000, 40000, 50000, 60000, 70000, 80000, 90000, 100000]


def is_valid_amount(amount) -> bool:
    if not isinstance(amount, int):
        return False
    if amount < AMOUNT_MIN or amount > AMOUNT_MAX:
        return False
    if amount % AMOUNT_STEP != 0:
        return False
    return True

# 部課番号 → 担当者 マッピング (memory ベース)
TANTOU_YOSHIURA = {"114", "083", "086", "105", "043"}
TANTOU_MORI_YUJI = {"077", "078", "071", "052", "040"}
TANTOU_MAKOTO = {"002", "093", "091", "094", "107", "116", "117", "112"}


def get_tantou(buka_no: str) -> str:
    """部課番号 (例 '002' or '002-1') から担当者名を返す。"""
    if not buka_no:
        return "未設定"
    no = buka_no.split("-")[0].strip()
    if no in TANTOU_YOSHIURA:
        return "吉浦マルセロ"
    if no in TANTOU_MORI_YUJI:
        return "森ゆうじ"
    if no in TANTOU_MAKOTO:
        return "森まこと"
    return "未設定"


def strip_buka_prefix(s: str) -> str:
    """'052:和興フィルタテクノロジー' → '和興フィルタテクノロジー'。"""
    if not s:
        return ""
    t = re.sub(r"^\s*\d+(?:[-－]\d+)?\s*[:：]\s*", "", s)
    t = t.replace("ﾕｰｼﾝ", "ユーシン")
    return t


def parse_buka_no(s: str) -> str:
    """'052:和興フィルタテクノロジー' → '052'。"""
    if not s:
        return ""
    m = re.match(r"^\s*(\d+(?:[-－]\d+)?)", s)
    return m.group(1) if m else ""


# ====== SharePoint REST クライアント ======
_token_cache: Dict[str, Tuple[str, float]] = {}


def _get_sp_token() -> str:
    """Managed Identity / az login から SharePoint アクセストークンを取得。
    キャッシュ有効期限は 50 分。"""
    cached = _token_cache.get(SP_RESOURCE)
    now = time.time()
    if cached and cached[1] > now + 60:
        return cached[0]
    cred = DefaultAzureCredential()
    tok = cred.get_token(f"{SP_RESOURCE}/.default")
    _token_cache[SP_RESOURCE] = (tok.token, tok.expires_on)
    return tok.token


def _sp_headers(verbose: bool = False, write: bool = False) -> Dict[str, str]:
    h = {
        "Authorization": f"Bearer {_get_sp_token()}",
        "Accept": "application/json;odata=verbose" if verbose else "application/json;odata=nometadata",
    }
    if write:
        h["Content-Type"] = "application/json;odata=verbose" if verbose else "application/json;odata=nometadata"
    return h


def sp_get_items(list_guid: str, select: Optional[str] = None,
                 filter_: Optional[str] = None, top: int = 5000,
                 orderby: Optional[str] = None) -> List[Dict[str, Any]]:
    """SP List のアイテムを取得。ページング対応。"""
    base = f"{SITE_URL}/_api/web/lists(guid'{list_guid}')/items"
    params = [f"$top={min(top, 5000)}"]
    if select:
        params.append(f"$select={select}")
    if filter_:
        params.append(f"$filter={filter_}")
    if orderby:
        params.append(f"$orderby={orderby}")
    url = base + "?" + "&".join(params)
    items: List[Dict[str, Any]] = []
    while url:
        r = requests.get(url, headers=_sp_headers(), timeout=60)
        r.raise_for_status()
        data = r.json()
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink") or data.get("odata.nextLink")
        if len(items) >= top:
            break
    return items


def sp_patch_item(list_guid: str, item_id: int, fields: Dict[str, Any]) -> None:
    """SP List アイテムを PATCH (MERGE)。verbose POST + X-HTTP-Method=MERGE。"""
    # ListItemEntityTypeFullName を取得
    list_url = f"{SITE_URL}/_api/web/lists(guid'{list_guid}')"
    type_resp = requests.get(f"{list_url}?$select=ListItemEntityTypeFullName",
                             headers=_sp_headers(verbose=True), timeout=30)
    type_resp.raise_for_status()
    entity_type = type_resp.json()["d"]["ListItemEntityTypeFullName"]
    body = {"__metadata": {"type": entity_type}}
    body.update(fields)
    h = {
        "Authorization": f"Bearer {_get_sp_token()}",
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "IF-MATCH": "*",
        "X-HTTP-Method": "MERGE",
    }
    r = requests.post(f"{list_url}/items({item_id})", headers=h,
                      data=json.dumps(body), timeout=60)
    if not r.ok:
        raise RuntimeError(f"PATCH failed: {r.status_code} {r.text[:300]}")


def sp_post_item(list_guid: str, fields: Dict[str, Any]) -> int:
    """SP List に新規アイテムを INSERT。AddValidateUpdateItemUsingPath を使用。
    成功時は新規 ID を返す。"""
    # まず ListItemEntityTypeFullName を取得
    list_url = f"{SITE_URL}/_api/web/lists(guid'{list_guid}')"
    list_path_resp = requests.get(f"{list_url}?$select=RootFolder/ServerRelativeUrl&$expand=RootFolder",
                                  headers=_sp_headers(verbose=True), timeout=30)
    list_path_resp.raise_for_status()
    server_rel = list_path_resp.json()["d"]["RootFolder"]["ServerRelativeUrl"]

    form_values = [{"FieldName": k, "FieldValue": str(v) if v is not None else ""} for k, v in fields.items()]
    body = {
        "listItemCreateInfo": {
            "FolderPath": {"DecodedUrl": server_rel},
            "UnderlyingObjectType": 0
        },
        "formValues": form_values,
        "bNewDocumentUpdate": False
    }
    h = _sp_headers(write=True)
    r = requests.post(f"{list_url}/AddValidateUpdateItemUsingPath",
                      headers=h, data=json.dumps(body), timeout=60)
    r.raise_for_status()
    data = r.json()
    # AddValidateUpdateItemUsingPath は value[] を返す。エラーチェック
    results = data.get("value", [])
    new_id = None
    errors = []
    for f in results:
        if f.get("HasException"):
            errors.append(f"{f.get('FieldName')}: {f.get('ErrorMessage')}")
        if f.get("FieldName") == "Id":
            new_id = int(f.get("FieldValue"))
    if errors:
        raise RuntimeError("SP POST errors: " + "; ".join(errors))
    if new_id is None:
        raise RuntimeError("SP POST: ID not returned")
    return new_id


# ====== JWT ======
def jwt_issue(shain_no: int) -> str:
    payload = {
        "sub": str(shain_no),
        "shainNo": shain_no,
        "iat": int(time.time()),
        "exp": int(time.time()) + JWT_EXP_HOURS * 3600,
    }
    return jwt.encode(payload, JWT_SECRET, algorithm=JWT_ALG)


def jwt_verify(token: str) -> Dict[str, Any]:
    return jwt.decode(token, JWT_SECRET, algorithms=[JWT_ALG])


def require_auth(req: func.HttpRequest) -> Tuple[Optional[Dict[str, Any]], Optional[func.HttpResponse]]:
    """JWT 検証。OK なら (payload, None)、NG なら (None, HttpResponse 401)。"""
    auth = req.headers.get("Authorization", "")
    if not auth.startswith("Bearer "):
        return None, _json_response({"error": "missing_token"}, 401)
    try:
        payload = jwt_verify(auth[7:].strip())
        return payload, None
    except Exception as e:
        return None, _json_response({"error": "invalid_token", "detail": str(e)}, 401)


# ====== HTTP ヘルパ ======
CORS_ORIGIN = os.environ.get("CORS_ORIGIN", "https://portal.team-stepup.com")


def _json_response(data: Any, status: int = 200, extra_headers: Optional[Dict[str, str]] = None) -> func.HttpResponse:
    headers = {
        "Access-Control-Allow-Origin": CORS_ORIGIN,
        "Access-Control-Allow-Headers": "Authorization, Content-Type",
        "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
        # セキュリティヘッダー
        "X-Content-Type-Options": "nosniff",
        "X-Frame-Options": "DENY",
        "Referrer-Policy": "strict-origin-when-cross-origin",
        "Cache-Control": "no-store, no-cache, must-revalidate, private",
        "Strict-Transport-Security": "max-age=31536000; includeSubDomains",
    }
    if extra_headers:
        headers.update(extra_headers)
    return func.HttpResponse(
        json.dumps(data, ensure_ascii=False, default=str),
        status_code=status,
        mimetype="application/json",
        headers=headers,
    )


def _handle_preflight(req: func.HttpRequest) -> Optional[func.HttpResponse]:
    if req.method == "OPTIONS":
        return _json_response({}, 204)
    return None


# ====== 業務ロジック: 社員照合 ======
def _utc_to_jst_date(utc_str: str) -> Optional[_dt.date]:
    if not utc_str:
        return None
    try:
        dt = _dt.datetime.fromisoformat(utc_str.replace("Z", "+00:00"))
    except Exception:
        return None
    jst = dt + _dt.timedelta(hours=9)
    return jst.date()


def find_active_employee(shain_no: int, birthday: str, tel_last4: str) -> Optional[Dict[str, Any]]:
    """社員番号で検索 → 生年月日 + 電話下4桁 で照合 → 在職判定。
    一致して在職中の最新レコードを返す。重複社員番号は複数行ヒット。"""
    items = sp_get_items(
        LIST_SHAIN,
        select=",".join([
            "Id", F_SHAIN_NO, F_SHAIN_NAME, F_BUKA, F_BIRTHDAY, F_TEL,
            F_TAISHA_DATE, F_ZAIYOKU, F_GINKO, F_SHITEN, F_KOUZA, F_MEIGI, F_ZAIRYU_NAME,
            F_KOKUSEKI, F_TSUKIN_OLD, F_TSUKIN_NEW, F_PORTAL_PIN,
        ]),
        filter_=f"{F_SHAIN_NO} eq {shain_no}",
        orderby="Id desc",
    )
    try:
        in_birthday = _dt.date.fromisoformat(birthday)
    except Exception:
        return None
    tel_last4 = (tel_last4 or "").strip()
    today = _dt.date.today()
    for it in items:
        bd = _utc_to_jst_date(it.get(F_BIRTHDAY))
        if bd != in_birthday:
            continue
        tel = (it.get(F_TEL) or "")
        if not tel or not tel.replace("-", "").replace(" ", "").endswith(tel_last4):
            continue
        # 在職判定: 派遣先退社日が None または未来
        taisha = _utc_to_jst_date(it.get(F_TAISHA_DATE))
        if taisha is not None and taisha < today:
            continue
        return it
    return None


# ====== 業務ロジック: 給料日週・休日 ======
def list_kaisha_kyujitsu_dates() -> List[_dt.date]:
    items = sp_get_items(LIST_KAISHA_KYUJITSU,
                         select=f"Id,{F_KYUJITSU_DATE},{F_KYUJITSU_KIND}")
    out = []
    for it in items:
        d = _utc_to_jst_date(it.get(F_KYUJITSU_DATE))
        if d:
            out.append(d)
    return out


def fetch_hakensaki(buka_text: str) -> Optional[Dict[str, Any]]:
    """部課（数字付き）から派遣先情報レコードを取得。"""
    if not buka_text:
        return None
    items = sp_get_items(
        LIST_HAKENSAKI,
        select=f"Id,{F_HAKEN_TITLE},{F_HAKEN_BANGO},{F_HAKEN_KAISYAMEI},{F_HAKEN_HARAIBI}",
    )
    bno = parse_buka_no(buka_text)
    for it in items:
        if it.get(F_HAKEN_BANGO) == bno or it.get(F_HAKEN_TITLE) == buka_text:
            return it
    return None


def parse_payday_rule(haraibi: str, ref_month: _dt.date) -> Optional[_dt.date]:
    """支払日（給料）テキストから ref_month に対応する給料日を計算。
    例: '翌月末日' / '翌月20日' / '翌月10日' / '20日' / '月末'."""
    if not haraibi:
        return None
    s = haraibi.strip()
    target_month = ref_month
    if "翌月" in s:
        # ref_month の翌月
        y, m = ref_month.year, ref_month.month
        m += 1
        if m > 12:
            y += 1
            m = 1
        target_month = _dt.date(y, m, 1)
    if "末日" in s or "月末" in s:
        # 該当月の末日
        y, m = target_month.year, target_month.month
        if m == 12:
            return _dt.date(y, 12, 31)
        return _dt.date(y, m + 1, 1) - _dt.timedelta(days=1)
    # 日付抽出
    m_match = re.search(r"(\d{1,2})\s*日", s)
    if m_match:
        d = int(m_match.group(1))
        try:
            return _dt.date(target_month.year, target_month.month, d)
        except ValueError:
            return None
    return None


def week_range(d: _dt.date) -> Tuple[_dt.date, _dt.date]:
    """月曜〜日曜の週範囲を返す。"""
    mon = d - _dt.timedelta(days=d.weekday())
    sun = mon + _dt.timedelta(days=6)
    return mon, sun


def upcoming_fridays(weeks: int = 6) -> List[_dt.date]:
    today = _dt.date.today()
    days_to_friday = (4 - today.weekday()) % 7
    if days_to_friday == 0:
        days_to_friday = 7
    first = today + _dt.timedelta(days=days_to_friday)
    return [first + _dt.timedelta(weeks=i) for i in range(weeks)]


def is_in_kaisha_kyujitsu_week(d: _dt.date, kyujitsu: List[_dt.date]) -> bool:
    mon, sun = week_range(d)
    return any(mon <= k <= sun for k in kyujitsu)


def is_in_payday_week(friday: _dt.date, payday: _dt.date) -> bool:
    mon, sun = week_range(friday)
    return mon <= payday <= sun


def get_next_payday_for_friday(haraibi: str, friday: _dt.date) -> Optional[_dt.date]:
    """金曜日 f に対する「次回受け取り予定の給料日」(>= f) を返す。
    例: 翌月末日支払いで f = 2026/05/29 なら、5/31 (4月分支払日) を返す。
         f = 2026/07/03 なら、7/31 (6月分支払日) を返す (6/30 は既に過ぎた支払い)。
    給料日週 NG 判定は「次回受け取り給料日が f と同週か」で行う。"""
    for offset in (-2, -1, 0, 1, 2, 3):
        y, m = friday.year, friday.month + offset
        while m > 12:
            y += 1
            m -= 12
        while m < 1:
            y -= 1
            m += 12
        ref = _dt.date(y, m, 1)
        pd = parse_payday_rule(haraibi, ref)
        if pd and pd >= friday:
            return pd
    return None


# ====== Functions ======
app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)


@app.route(route="auth/pin-status", methods=["POST", "OPTIONS"])
def auth_pin_status(req: func.HttpRequest) -> func.HttpResponse:
    """社員番号から PIN 設定済みかを返す (ログイン画面の出し分け用)。
    社員の存在有無は秘匿せず、PIN 設定済みかどうかだけを返す。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    shain_no = body.get("shainNo")
    try:
        shain_no = int(shain_no)
    except (TypeError, ValueError):
        return _json_response({"error": "invalid_shain_no"}, 400)
    try:
        emp = find_active_employee_by_shain(shain_no)
    except Exception:
        emp = None
    has_pin = bool(emp and (emp.get(F_PORTAL_PIN) or "").strip())
    # 在職社員が存在するかも返す (存在しない番号でも pinSet=false を返し列挙対策)
    return _json_response({"pinSet": has_pin})


def _validate_pin(pin: Any) -> Optional[str]:
    """4桁数字 PIN のバリデーション。OK なら文字列、NG なら None。"""
    if pin is None:
        return None
    s = str(pin).strip()
    if re.fullmatch(r"\d{4}", s):
        # 単純すぎる PIN を弾く (0000/1234/1111 等)
        if s in ("0000", "1111", "2222", "3333", "4444", "5555",
                 "6666", "7777", "8888", "9999", "1234", "4321"):
            return None
        return s
    return None


@app.route(route="auth/login", methods=["POST", "OPTIONS"])
def auth_login(req: func.HttpRequest) -> func.HttpResponse:
    """初回ログイン (3要素照合)。PIN 未設定なら pinSetupRequired を返す。
    PIN 設定済みの場合はこのエンドポイントは使わず /auth/pin-login を使う想定だが、
    互換のため 3要素一致時はトークンも発行する。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    shain_no = body.get("shainNo")
    birthday = body.get("birthday")
    tel_last4 = body.get("telLast4")
    if not (shain_no and birthday and tel_last4):
        return _json_response({"error": "missing_fields"}, 400)
    try:
        shain_no = int(shain_no)
    except (TypeError, ValueError):
        return _json_response({"error": "invalid_shain_no"}, 400)

    remaining = check_lockout(shain_no)
    if remaining is not None:
        return _json_response({
            "error": "locked_out",
            "remainingSeconds": remaining,
            "lockoutMinutes": LOCKOUT_DURATION_MINUTES,
        }, 429, extra_headers={"Retry-After": str(remaining)})

    try:
        emp = find_active_employee(shain_no, birthday, str(tel_last4))
    except Exception as e:
        logging.exception("find_active_employee failed")
        return _json_response({"error": "lookup_failed", "detail": str(e)}, 500)
    if not emp:
        record_failed_attempt(shain_no)
        return _json_response({"error": "auth_failed"}, 401)

    clear_attempts(shain_no)
    has_pin = bool((emp.get(F_PORTAL_PIN) or "").strip())
    if not has_pin:
        # PIN 未設定 → 短命の setup トークンを発行し、PIN 設定画面へ誘導
        setup_token = jwt_issue_setup(shain_no)
        return _json_response({
            "pinSetupRequired": True,
            "setupToken": setup_token,
            "profile": _employee_to_profile(emp),
        })
    # PIN 設定済みでも 3要素が合っていればログインさせる (PIN 忘れの救済も兼ねる)
    token = jwt_issue(shain_no)
    profile = _employee_to_profile(emp)
    return _json_response({"token": token, "profile": profile})


def jwt_issue_setup(shain_no: int) -> str:
    """PIN 設定専用の短命トークン (10分)。"""
    payload = {
        "sub": str(shain_no),
        "shainNo": shain_no,
        "purpose": "pin_setup",
        "iat": int(time.time()),
        "exp": int(time.time()) + 600,
    }
    return jwt.encode(payload, JWT_SECRET, algorithm=JWT_ALG)


@app.route(route="auth/set-pin", methods=["POST", "OPTIONS"])
def auth_set_pin(req: func.HttpRequest) -> func.HttpResponse:
    """PIN 設定。setupToken (3要素照合済) + 新しい4桁PIN を受け取りハッシュ保存。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    auth = req.headers.get("Authorization", "")
    if not auth.startswith("Bearer "):
        return _json_response({"error": "missing_token"}, 401)
    try:
        payload = jwt_verify(auth[7:].strip())
    except Exception as e:
        return _json_response({"error": "invalid_token", "detail": str(e)}, 401)
    if payload.get("purpose") != "pin_setup":
        return _json_response({"error": "invalid_setup_token"}, 403)
    shain_no = int(payload["shainNo"])
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    pin = _validate_pin(body.get("pin"))
    if not pin:
        return _json_response({"error": "invalid_pin"}, 400)
    try:
        emp = find_active_employee_by_shain(shain_no)
        if not emp:
            return _json_response({"error": "not_active"}, 403)
        pin_hash = hash_pin(shain_no, pin)
        sp_patch_item(LIST_SHAIN, int(emp.get("Id")), {F_PORTAL_PIN: pin_hash})
        # PIN 設定完了 → 本ログイントークン発行
        token = jwt_issue(shain_no)
        return _json_response({"ok": True, "token": token, "profile": _employee_to_profile(emp)})
    except Exception as e:
        logging.exception("set_pin failed")
        return _json_response({"error": "internal", "detail": str(e)}, 500)


@app.route(route="auth/pin-login", methods=["POST", "OPTIONS"])
def auth_pin_login(req: func.HttpRequest) -> func.HttpResponse:
    """2回目以降のログイン: 社員番号 + 4桁PIN。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    shain_no = body.get("shainNo")
    pin = body.get("pin")
    if not (shain_no and pin):
        return _json_response({"error": "missing_fields"}, 400)
    try:
        shain_no = int(shain_no)
    except (TypeError, ValueError):
        return _json_response({"error": "invalid_shain_no"}, 400)

    remaining = check_lockout(shain_no)
    if remaining is not None:
        return _json_response({
            "error": "locked_out",
            "remainingSeconds": remaining,
            "lockoutMinutes": LOCKOUT_DURATION_MINUTES,
        }, 429, extra_headers={"Retry-After": str(remaining)})

    try:
        emp = find_active_employee_by_shain(shain_no)
    except Exception as e:
        logging.exception("pin-login lookup failed")
        return _json_response({"error": "lookup_failed", "detail": str(e)}, 500)
    stored = (emp.get(F_PORTAL_PIN) if emp else None)
    if not emp or not (stored or "").strip():
        # PIN 未設定 → 初回ログインへ誘導
        return _json_response({"error": "pin_not_set"}, 409)
    if not verify_pin(shain_no, str(pin), stored):
        record_failed_attempt(shain_no)
        return _json_response({"error": "auth_failed"}, 401)
    clear_attempts(shain_no)
    token = jwt_issue(shain_no)
    return _json_response({"token": token, "profile": _employee_to_profile(emp)})


def _employee_to_profile(emp: Dict[str, Any]) -> Dict[str, Any]:
    buka_text = emp.get(F_BUKA) or ""
    kokuseki = emp.get(F_KOKUSEKI) or ""
    return {
        "shainNo": emp.get(F_SHAIN_NO),
        "name": emp.get(F_SHAIN_NAME),
        "zairyuName": emp.get(F_ZAIRYU_NAME) or "",
        "kokuseki": kokuseki,
        "preferredLang": guess_lang_from_kokuseki(kokuseki),
        "hakensaki": strip_buka_prefix(buka_text),
        "bukaRaw": buka_text,
        "bukaNo": parse_buka_no(buka_text),
        "ginko": emp.get(F_GINKO),
        "shiten": emp.get(F_SHITEN),
        "kouza": emp.get(F_KOUZA),
        "meigi": emp.get(F_MEIGI),
        "zaiyokuSyubetu": emp.get(F_ZAIYOKU),
        "commutesByCar": commutes_by_car(emp),
    }


@app.route(route="profile", methods=["GET", "OPTIONS"])
def profile_get(req: func.HttpRequest) -> func.HttpResponse:
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    try:
        emp = find_active_employee_by_shain(payload["shainNo"])
    except Exception as e:
        logging.exception("profile lookup failed")
        return _json_response({"error": "lookup_failed", "detail": str(e)}, 500)
    if not emp:
        return _json_response({"error": "not_active"}, 403)
    return _json_response({"profile": _employee_to_profile(emp)})


def find_active_employee_by_shain(shain_no: int) -> Optional[Dict[str, Any]]:
    items = sp_get_items(
        LIST_SHAIN,
        select=",".join([
            "Id", F_SHAIN_NO, F_SHAIN_NAME, F_BUKA, F_BIRTHDAY, F_TEL,
            F_TAISHA_DATE, F_ZAIYOKU, F_GINKO, F_SHITEN, F_KOUZA, F_MEIGI, F_ZAIRYU_NAME,
            F_KOKUSEKI, F_TSUKIN_OLD, F_TSUKIN_NEW, F_PORTAL_PIN,
        ]),
        filter_=f"{F_SHAIN_NO} eq {shain_no}",
        orderby="Id desc",
    )
    today = _dt.date.today()
    for it in items:
        taisha = _utc_to_jst_date(it.get(F_TAISHA_DATE))
        if taisha is None or taisha >= today:
            return it
    return None


@app.route(route="maebarai/dates", methods=["GET", "OPTIONS"])
def maebarai_dates(req: func.HttpRequest) -> func.HttpResponse:
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    shain_no = payload["shainNo"]
    try:
        emp = find_active_employee_by_shain(shain_no)
        if not emp:
            return _json_response({"error": "not_active"}, 403)
        buka_text = emp.get(F_BUKA) or ""
        hakensaki = fetch_hakensaki(buka_text)
        haraibi = (hakensaki or {}).get(F_HAKEN_HARAIBI) or "翌月末日"
        kyujitsu = list_kaisha_kyujitsu_dates()
        fridays = upcoming_fridays(weeks=6)
        out = []
        for f in fridays:
            next_payday = get_next_payday_for_friday(haraibi, f)
            reasons = []
            if is_in_kaisha_kyujitsu_week(f, kyujitsu):
                reasons.append("会社休日")
            if next_payday and is_in_payday_week(f, next_payday):
                reasons.append("給料日週")
            out.append({
                "date": f.isoformat(),
                "available": len(reasons) == 0,
                "reasons": reasons,
                "nextPayday": next_payday.isoformat() if next_payday else None,
            })
        return _json_response({
            "haraibiRule": haraibi,
            "hakensaki": strip_buka_prefix(buka_text),
            "dates": out,
        })
    except Exception as e:
        logging.exception("maebarai_dates failed")
        return _json_response({"error": "internal", "detail": str(e)}, 500)


@app.route(route="maebarai/apply", methods=["POST", "OPTIONS"])
def maebarai_apply(req: func.HttpRequest) -> func.HttpResponse:
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    shain_no = payload["shainNo"]
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    amount = body.get("amount")
    date = body.get("date")
    reason = (body.get("reason") or "")[:500]
    if not is_valid_amount(amount):
        return _json_response({"error": "invalid_amount", "min": AMOUNT_MIN, "max": AMOUNT_MAX, "step": AMOUNT_STEP}, 400)
    try:
        d = _dt.date.fromisoformat(date)
    except Exception:
        return _json_response({"error": "invalid_date"}, 400)
    if d.weekday() != 4:
        return _json_response({"error": "not_friday"}, 400)

    # レート制限 (10/min)
    wait = check_rate_limit(shain_no, "apply")
    if wait is not None:
        return _json_response({"error": "rate_limited", "retryAfterSeconds": wait},
                              429, extra_headers={"Retry-After": str(wait)})
    try:
        emp = find_active_employee_by_shain(shain_no)
        if not emp:
            return _json_response({"error": "not_active"}, 403)
        buka_text = emp.get(F_BUKA) or ""
        hakensaki = fetch_hakensaki(buka_text)
        haraibi = (hakensaki or {}).get(F_HAKEN_HARAIBI) or "翌月末日"
        kyujitsu = list_kaisha_kyujitsu_dates()
        next_payday = get_next_payday_for_friday(haraibi, d)
        if is_in_kaisha_kyujitsu_week(d, kyujitsu) or (next_payday and is_in_payday_week(d, next_payday)):
            return _json_response({"error": "date_not_available"}, 400)
        # 重複チェック: 同じ社員番号 + 同じ希望日 で status が pending または approved の申請があれば拒否
        dup_filter = (
            f"OData__x793e__x54e1__x756a__x53f7_ eq {shain_no} and "
            f"OData__x5e0c__x671b__x65e5_ eq datetime'{d.isoformat()}T00:00:00' and "
            f"(Status eq 'pending' or Status eq 'approved')"
        )
        try:
            dup_items = sp_get_items(LIST_MAEBARAI_SHINSEI, select="Id,Status", filter_=dup_filter, top=5)
        except Exception:
            dup_items = []
        if dup_items:
            existing_status = dup_items[0].get("Status") or "pending"
            return _json_response({
                "error": "duplicate_application",
                "existingStatus": existing_status,
                "existingId": dup_items[0].get("Id"),
            }, 409)
        # INSERT
        buka_no = parse_buka_no(buka_text)
        tantou = get_tantou(buka_no)
        fields = {
            "Title": f"{shain_no} {d.isoformat()}",
            P_SHAIN_NO: str(shain_no),
            P_KINGAKU: str(amount),
            # AddValidateUpdateItemUsingPath は SP の表示ロケール (ja-JP) でパース。
            # 日本テナントは yyyy/MM/dd 形式で渡す必要がある。
            P_KIBOUBI: d.strftime("%Y/%m/%d"),
            P_RIYUU: reason,
            P_HAKENSAKI: strip_buka_prefix(buka_text),
            P_TANTOU: tantou,
            P_STATUS: "pending",
        }
        new_id = sp_post_item(LIST_MAEBARAI_SHINSEI, fields)
        return _json_response({"ok": True, "id": new_id, "tantou": tantou})
    except Exception as e:
        logging.exception("maebarai_apply failed")
        return _json_response({"error": "internal", "detail": str(e)}, 500)


@app.route(route="maebarai/history", methods=["GET", "OPTIONS"])
def maebarai_history(req: func.HttpRequest) -> func.HttpResponse:
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    shain_no = payload["shainNo"]
    try:
        # SELECT/FILTER は EntityPropertyName (OData_ プレフィックス付き) を使う
        F_M_SHAIN = "OData__x793e__x54e1__x756a__x53f7_"
        F_M_KINGAKU = "OData__x91d1__x984d_"
        F_M_KIBOUBI = "OData__x5e0c__x671b__x65e5_"
        F_M_RIYUU = "OData__x7533__x8acb__x7406__x7531_"
        F_M_HAKEN = "OData__x6d3e__x9063__x5148_"
        F_M_TANTOU = "OData__x62c5__x5f53__x8005_"
        F_M_APPROVED = "OData__x627f__x8a8d__x65e5__x6642_"
        F_M_REJECT = "OData__x5374__x4e0b__x7406__x7531_"
        select_fields = ",".join([
            "Id", F_M_SHAIN, F_M_KINGAKU, F_M_KIBOUBI, F_M_RIYUU,
            F_M_HAKEN, F_M_TANTOU, "Status", F_M_APPROVED, F_M_REJECT, "Created",
        ])
        items = sp_get_items(
            LIST_MAEBARAI_SHINSEI,
            select=select_fields,
            filter_=f"{F_M_SHAIN} eq {shain_no}",
            orderby="Id desc",
        )
        out = []
        for it in items:
            kiboubi = _utc_to_jst_date(it.get(F_M_KIBOUBI))
            created = _utc_to_jst_date(it.get("Created"))
            approved = _utc_to_jst_date(it.get(F_M_APPROVED))
            out.append({
                "id": it.get("Id"),
                "amount": it.get(F_M_KINGAKU),
                "date": kiboubi.isoformat() if kiboubi else None,
                "reason": it.get(F_M_RIYUU),
                "status": it.get("Status"),
                "approvedAt": approved.isoformat() if approved else None,
                "rejectReason": it.get(F_M_REJECT),
                "createdAt": created.isoformat() if created else None,
            })
        return _json_response({"items": out})
    except Exception as e:
        logging.exception("maebarai_history failed")
        return _json_response({"error": "internal", "detail": str(e)}, 500)


@app.route(route="maebarai/cancel", methods=["POST", "OPTIONS"])
def maebarai_cancel(req: func.HttpRequest) -> func.HttpResponse:
    """本人が pending 申請を取り消す。
    Body: { id: <listItemId> }
    成功時: Status を cancelled に PATCH (履歴には残る、cancelled 表示)。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    shain_no = payload["shainNo"]
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    item_id = body.get("id")
    if not item_id:
        return _json_response({"error": "missing_id"}, 400)
    try:
        item_id = int(item_id)
    except (TypeError, ValueError):
        return _json_response({"error": "invalid_id"}, 400)

    # レート制限 (10/min)
    wait = check_rate_limit(shain_no, "cancel")
    if wait is not None:
        return _json_response({"error": "rate_limited", "retryAfterSeconds": wait},
                              429, extra_headers={"Retry-After": str(wait)})
    try:
        # 申請が本人のもの & status が pending か確認
        items = sp_get_items(
            LIST_MAEBARAI_SHINSEI,
            select=f"Id,Status,OData__x793e__x54e1__x756a__x53f7_",
            filter_=f"Id eq {item_id}",
            top=1,
        )
        if not items:
            return _json_response({"error": "not_found"}, 404)
        it = items[0]
        if int(it.get("OData__x793e__x54e1__x756a__x53f7_") or 0) != int(shain_no):
            return _json_response({"error": "forbidden"}, 403)
        if (it.get("Status") or "").lower() != "pending":
            return _json_response({"error": "cannot_cancel", "currentStatus": it.get("Status")}, 400)
        # PATCH Status → cancelled
        sp_patch_item(LIST_MAEBARAI_SHINSEI, item_id, {
            "Status": "cancelled",
            "OData__x627f__x8a8d__x65e5__x6642_": _dt.datetime.utcnow().isoformat() + "Z",
        })
        return _json_response({"ok": True})
    except Exception as e:
        logging.exception("maebarai_cancel failed")
        return _json_response({"error": "internal", "detail": str(e)}, 500)


# ====== 在留カード OCR (Azure Document Intelligence) ======
_di_client_cache = None


def _get_di_client():
    """Document Intelligence クライアントを Managed Identity で取得 (キャッシュ)。"""
    global _di_client_cache
    if _di_client_cache is not None:
        return _di_client_cache
    endpoint = os.environ.get("FORM_RECOGNIZER_ENDPOINT", "").strip()
    if not endpoint:
        raise RuntimeError("FORM_RECOGNIZER_ENDPOINT not set")
    from azure.ai.documentintelligence import DocumentIntelligenceClient
    _di_client_cache = DocumentIntelligenceClient(
        endpoint=endpoint,
        credential=DefaultAzureCredential(),
    )
    return _di_client_cache


def _run_read_ocr(image_bytes: bytes) -> str:
    """prebuilt-read で汎用 OCR を実行し、全行テキストを連結して返す共通ヘルパ。"""
    from azure.ai.documentintelligence.models import AnalyzeDocumentRequest
    client = _get_di_client()
    poller = client.begin_analyze_document(
        "prebuilt-read",
        AnalyzeDocumentRequest(bytes_source=image_bytes),
    )
    result = poller.result()
    full_text = ""
    for page in (result.pages or []):
        for line in (page.lines or []):
            full_text += (line.content or "") + "\n"
    return full_text


def run_zairyu_ocr(image_bytes: bytes) -> Dict[str, Any]:
    """在留カード OCR。prebuilt-read で汎用OCR → 正規表現で抽出。"""
    from azure.ai.documentintelligence.models import AnalyzeDocumentRequest
    client = _get_di_client()
    poller = client.begin_analyze_document(
        "prebuilt-read",
        AnalyzeDocumentRequest(bytes_source=image_bytes),
    )
    result = poller.result()
    full_text = ""
    for page in (result.pages or []):
        for line in (page.lines or []):
            full_text += (line.content or "") + "\n"
    parsed = parse_zairyu_card_text(full_text)
    return parsed


# 国籍の既知値リスト (長い順 — 「アメリカ合衆国」を「アメリカ」より先に)
NATIONALITY_KNOWN = [
    "アメリカ合衆国", "ドミニカ共和国",
    "ブラジル", "フィリピン", "ペルー", "ベトナム", "中国", "韓国",
    "インドネシア", "ネパール", "バングラデシュ", "スリランカ", "パキスタン",
    "ミャンマー", "カンボジア", "ラオス", "モンゴル", "台湾",
    "ボリビア", "アルゼンチン", "パラグアイ", "ウルグアイ",
    "コロンビア", "ベネズエラ", "エクアドル", "メキシコ",
    "ロシア", "ウクライナ", "インド", "マレーシア", "シンガポール",
    "香港", "カナダ", "ドイツ", "フランス", "イタリア", "スペイン",
    "オーストラリア", "ニュージーランド",
    "アメリカ", "イギリス",
    "タイ", "チリ",  # 短い名前は誤マッチ可能性ありなので最後
]

# 在留資格の既知値リスト (長い順 — 「永住者の配偶者等」を「永住者」より先に)
ZAIRYU_SHIKAKU_KNOWN = [
    "高度専門職1号イ", "高度専門職1号ロ", "高度専門職1号ハ", "高度専門職2号", "高度専門職",
    "技術・人文知識・国際業務",
    "永住者の配偶者等", "日本人の配偶者等",
    "特定技能1号", "特定技能2号", "特定技能",
    "技能実習1号イ", "技能実習1号ロ", "技能実習2号イ", "技能実習2号ロ", "技能実習3号イ", "技能実習3号ロ", "技能実習",
    "法律・会計業務",
    "経営・管理",
    "特定活動",
    "短期滞在",
    "家族滞在",
    "永住者", "定住者",
    "教授", "芸術", "宗教", "報道", "教育", "研究", "医療", "介護", "技能",
    "留学", "研修",
    "外交", "公用",
    "Permanent Resident", "Long-Term Resident", "Spouse",
]


def _is_permanent_status(s: Optional[str]) -> bool:
    if not s:
        return False
    return "永住" in s or "Permanent" in s


def parse_zairyu_card_text(text: str) -> Dict[str, Any]:
    """汎用OCRテキストから在留カードの主要項目を抽出 (best-effort)。"""
    out: Dict[str, Any] = {"rawText": text}
    if not text:
        return out

    lines = [ln.strip() for ln in text.split('\n')]

    # 在留カード番号: AB12345678CD 形式 (12桁、頭2/末尾2 が英字)
    m = re.search(r'\b([A-Z]{2}\d{8}[A-Z]{2})\b', text)
    if m:
        out['cardNumber'] = m.group(1)

    # 氏名: 行ベース。「氏名」/「NAME」を含む行の次の英字行を取得
    for i, line in enumerate(lines):
        if '氏名' in line or re.search(r'\bNAME\b', line):
            for j in range(i + 1, min(i + 4, len(lines))):
                cand = lines[j].strip()
                # 英字+空白+記号のみの行 (3文字以上)
                if re.match(r'^[A-Z][A-Z\s,\.\-\']{2,60}$', cand) and 'NAME' not in cand:
                    out['name'] = re.sub(r'\s+', ' ', cand)
                    break
            if 'name' in out:
                break

    # 生年月日 DATE OF BIRTH
    for pat in (
        r'(?:DATE\s+OF\s+BIRTH|生年月日)[\s:：]*\n?\s*(\d{4})[\.\-/年\s]+(\d{1,2})[\.\-/月\s]+(\d{1,2})',
    ):
        m = re.search(pat, text)
        if m:
            out['birthday'] = f"{m.group(1)}-{m.group(2).zfill(2)}-{m.group(3).zfill(2)}"
            break

    # 性別 SEX
    m = re.search(r'SEX[\s:]*\n?\s*([MF])\b', text)
    if m:
        out['sex'] = m.group(1)

    # 国籍 NATIONALITY/REGION — 既知値リストから優先マッチ。
    # OCR が「ブラジル」を「ブラ ジル」のように文字間スペースで分割するため、
    # テキスト・リスト両方からスペースを除去してから比較する。
    def _strip_ws(s: str) -> str:
        return re.sub(r'[\s　]+', '', s or '')
    text_ns = _strip_ws(text)
    # ラベル「国籍」「NATIONALITY」付近を優先
    nat_label_idx = None
    for i, line in enumerate(lines):
        if 'NATIONALITY' in line.upper() or '国籍' in line:
            nat_label_idx = i
            break
    near_text = ''
    if nat_label_idx is not None:
        # ラベル前後3行をくっつけて空白除去
        start = max(0, nat_label_idx - 1)
        end = min(nat_label_idx + 6, len(lines))
        near_text = _strip_ws('\n'.join(lines[start:end]))
    # 第1優先: ラベル近傍
    if near_text:
        for country in NATIONALITY_KNOWN:
            if _strip_ws(country) in near_text:
                out['nationality'] = country
                break
    # フォールバック: 全文 (住所と被るリスクあり)
    if 'nationality' not in out:
        for country in NATIONALITY_KNOWN:
            if _strip_ws(country) in text_ns:
                out['nationality'] = country
                break

    # 在留資格: 既知値リストから優先マッチ (スペース除去マッチ)
    for known in ZAIRYU_SHIKAKU_KNOWN:
        if _strip_ws(known) in text_ns:
            out['zairyuShikaku'] = known
            break

    # 在留期限 — 通常の日付パターン
    # 注意: "**年**月" や "0000年00月" は無効として扱う
    found_kigen = None
    for pat in (
        r'(?:DATE\s+OF\s+EXPIRATION|在留期間.*?満了日|PERIOD\s+OF\s+STAY)[\s\S]{0,80}?(\d{4})[\.\-/年\s]+(\d{1,2})[\.\-/月\s]+(\d{1,2})',
    ):
        m = re.search(pat, text)
        if m:
            y, mo, d = m.group(1), m.group(2).zfill(2), m.group(3).zfill(2)
            if y not in ("0000",) and mo != "00" and d != "00":
                found_kigen = f"{y}-{mo}-{d}"
                break
    if found_kigen:
        out['zairyuKigen'] = found_kigen

    # 永住者 or 在留期限読取失敗 → カード有効期限を代用
    is_permanent = _is_permanent_status(out.get('zairyuShikaku'))
    if is_permanent or 'zairyuKigen' not in out:
        for pat in (
            # 「このカードは 2029年03月09日まで有効」
            r'(\d{4})[年\.\-/]+(\d{1,2})[月\.\-/]+(\d{1,2})\s*日?\s*まで有効',
            # 「PERIOD OF VALIDITY OF THIS CARD ... 2029.03.09」
            r'PERIOD\s+OF\s+VALIDITY[\s\S]{0,120}?(\d{4})[\.\-/年\s]+(\d{1,2})[\.\-/月\s]+(\d{1,2})',
        ):
            m = re.search(pat, text)
            if m:
                y, mo, d = m.group(1), m.group(2).zfill(2), m.group(3).zfill(2)
                if y != "0000":
                    out['zairyuKigen'] = f"{y}-{mo}-{d}"
                    if is_permanent:
                        out['zairyuKigenSource'] = 'cardValidity'  # 永住者 → カード有効期限を代用
                    break

    return out


@app.route(route="zairyu/ocr", methods=["POST", "OPTIONS"])
def zairyu_ocr(req: func.HttpRequest) -> func.HttpResponse:
    """在留カード表面画像から主要項目を OCR で抽出。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    shain_no = int(payload["shainNo"])
    # レート制限 (5/min)
    wait = check_rate_limit(shain_no, "zairyu")
    if wait is not None:
        return _json_response({"error": "rate_limited", "retryAfterSeconds": wait},
                              429, extra_headers={"Retry-After": str(wait)})
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    front_data = body.get("frontImage")
    if not front_data:
        return _json_response({"error": "missing_image"}, 400)
    try:
        front_bytes = base64.b64decode(_strip_data_url(front_data))
    except Exception as e:
        return _json_response({"error": "invalid_base64", "detail": str(e)}, 400)
    if len(front_bytes) > 10 * 1024 * 1024:
        return _json_response({"error": "image_too_large", "maxMB": 10}, 400)

    try:
        ocr_result = run_zairyu_ocr(front_bytes)
    except Exception as e:
        logging.exception("OCR failed")
        return _json_response({"error": "ocr_failed", "detail": str(e)}, 500)

    # 本人の社員データと突合 (生年月日が一致するかを返す)
    emp = find_active_employee_by_shain(shain_no)
    matches: Dict[str, bool] = {}
    if emp:
        emp_birthday = _utc_to_jst_date(emp.get(F_BIRTHDAY))
        ocr_bd = ocr_result.get('birthday')
        if emp_birthday and ocr_bd:
            try:
                matches['birthday'] = (_dt.date.fromisoformat(ocr_bd) == emp_birthday)
            except Exception:
                pass
    return _json_response({"ocr": ocr_result, "matches": matches})


# ====== 運転免許証 OCR ======
def parse_license_card_text(text: str) -> Dict[str, Any]:
    """汎用OCRテキストから日本の運転免許証の主要項目を抽出。"""
    out: Dict[str, Any] = {"rawText": text}
    if not text:
        return out

    def _strip_ws(s: str) -> str:
        return re.sub(r'[\s　]+', '', s or '')
    text_ns = _strip_ws(text)

    # 免許証番号: 12桁の数字
    m = re.search(r'第?\s*(\d{12})\s*号', text)
    if m:
        out['licenseNumber'] = m.group(1)
    else:
        # 「番号」ラベル近くの 12桁を探す
        m = re.search(r'番[\s号]*第?\s*(\d{12})', text)
        if m:
            out['licenseNumber'] = m.group(1)
        else:
            # 全体から 12桁を探す (1つだけある場合)
            all_nums = re.findall(r'\b(\d{12})\b', text)
            if len(all_nums) == 1:
                out['licenseNumber'] = all_nums[0]

    # 生年月日: 「YYYY年MM月DD日生」
    m = re.search(r'(?:昭和\s*(\d+)|平成\s*(\d+)|令和\s*(\d+)|(\d{4}))[\s年]+(\d{1,2})[\s月]+(\d{1,2})[\s日生]', text)
    if m:
        if m.group(1):  # 昭和
            y = 1925 + int(m.group(1))
        elif m.group(2):  # 平成
            y = 1988 + int(m.group(2))
        elif m.group(3):  # 令和
            y = 2018 + int(m.group(3))
        else:
            y = int(m.group(4))
        mo, d = m.group(5).zfill(2), m.group(6).zfill(2)
        out['birthday'] = f"{y}-{mo}-{d}"

    # 氏名: 「氏名」ラベルの後の行 (漢字またはカナ)
    lines = [ln.strip() for ln in text.split('\n')]
    for i, line in enumerate(lines):
        if '氏名' in line or 'NAME' in line.upper():
            for j in range(i + 1, min(i + 4, len(lines))):
                cand = lines[j].strip()
                # 漢字/カナ/ローマ字を含む 2文字以上の名前らしき行
                if cand and len(cand) >= 2 and re.search(r'[一-鿿ァ-ヶー一-鿿A-Z]', cand):
                    # ラベル単語などを除外
                    if not any(skip in cand for skip in ['氏名', '住所', '生年月日', '番号', '本籍']):
                        out['name'] = cand
                        break
            if 'name' in out:
                break

    # 有効期限: 「YYYY年MM月DD日まで有効」 / 「有効期限 YYYY年MM月DD日」
    for pat in (
        r'(\d{4})[年\.\-/]+(\d{1,2})[月\.\-/]+(\d{1,2})\s*日?[\s\S]{0,15}?まで有効',
        r'有効期限?[\s:：]*\n?\s*(\d{4})[年\.\-/]+(\d{1,2})[月\.\-/]+(\d{1,2})',
        r'(?:平成|令和)\s*(\d+)[年\.\-/]+(\d{1,2})[月\.\-/]+(\d{1,2})\s*日?[\s\S]{0,15}?まで有効',
    ):
        m = re.search(pat, text)
        if m:
            groups = m.groups()
            if len(groups) == 3:
                # 西暦パターン
                if groups[0] and len(groups[0]) == 4:
                    y, mo, d = groups[0], groups[1].zfill(2), groups[2].zfill(2)
                    out['licenseExpiry'] = f"{y}-{mo}-{d}"
                    break
                # 元号パターン
                else:
                    y_jp = int(groups[0])
                    y = 2018 + y_jp if y_jp <= 30 else 1988 + y_jp  # 令和か平成かを推定
                    mo, d = groups[1].zfill(2), groups[2].zfill(2)
                    out['licenseExpiry'] = f"{y}-{mo:>02}-{d:>02}"
                    break

    # 免許の種類: 「免許の種類」セクションから (普通/中型/大型/二輪/原付/けん引/大特 等)
    LICENSE_TYPES = ["大型二種", "中型二種", "普通二種", "大特二種", "けん引二種",
                     "大型", "中型", "準中型", "普通", "大特", "けん引",
                     "大型自動二輪", "普通自動二輪", "小型特殊", "原付"]
    found_types = []
    for lt in LICENSE_TYPES:
        if _strip_ws(lt) in text_ns:
            # 重複避け: 「中型」より「準中型」が先にマッチ済みなら追加しない
            if not any(_strip_ws(lt) in _strip_ws(f) for f in found_types):
                found_types.append(lt)
    if found_types:
        out['licenseType'] = ' '.join(found_types[:5])  # 最大5種類まで

    # 取得年月日 (種別ごとの最初の取得日 = 一番古い日付っぽいもの)
    # 免許証下部の「取得年月日」表
    date_matches = re.findall(r'(?:昭和|平成|令和)\s*(\d+)\s*[年\.](\d{1,2})\s*[月\.](\d{1,2})', text)
    if date_matches:
        # 元号変換 → 全部西暦に
        parsed_dates = []
        for jy, mo, d in date_matches:
            jy_int = int(jy)
            # 元号判定: 数字の前のテキスト確認は複雑なので、年数の大きさで推定
            # 30 以下は令和の可能性、それより大きいなら平成、それより大きいなら昭和
            # 簡略: 30以下→令和、31-31→平成可能性、32以上→平成
            # しかし正確には文脈が必要。ここでは保守的に、年が2桁なら西暦推定
            try:
                # 取り合えず 平成/令和 2 パターン試す
                y_reiwa = 2018 + jy_int
                y_heisei = 1988 + jy_int
                # 不可能な日付は除外
                from datetime import date as _date
                today = _date.today()
                if jy_int <= 6 and _date(y_reiwa, int(mo), int(d)) <= today:
                    parsed_dates.append(_date(y_reiwa, int(mo), int(d)))
                elif _date(y_heisei, int(mo), int(d)) <= today:
                    parsed_dates.append(_date(y_heisei, int(mo), int(d)))
            except Exception:
                continue
        if parsed_dates:
            # 一番古い日付 (=最初に取得した免許の日付) を取得年月日として採用
            oldest = min(parsed_dates)
            out['licenseDate'] = oldest.isoformat()

    return out


@app.route(route="license/ocr", methods=["POST", "OPTIONS"])
def license_ocr(req: func.HttpRequest) -> func.HttpResponse:
    """運転免許証 表面画像から主要項目を OCR で抽出。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    shain_no = int(payload["shainNo"])
    wait = check_rate_limit(shain_no, "zairyu")  # 同じレート制限プール使用
    if wait is not None:
        return _json_response({"error": "rate_limited", "retryAfterSeconds": wait},
                              429, extra_headers={"Retry-After": str(wait)})
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    front_data = body.get("frontImage")
    if not front_data:
        return _json_response({"error": "missing_image"}, 400)
    try:
        front_bytes = base64.b64decode(_strip_data_url(front_data))
    except Exception as e:
        return _json_response({"error": "invalid_base64", "detail": str(e)}, 400)
    if len(front_bytes) > 10 * 1024 * 1024:
        return _json_response({"error": "image_too_large", "maxMB": 10}, 400)

    try:
        from azure.ai.documentintelligence.models import AnalyzeDocumentRequest
        client = _get_di_client()
        poller = client.begin_analyze_document(
            "prebuilt-read",
            AnalyzeDocumentRequest(bytes_source=front_bytes),
        )
        result = poller.result()
        full_text = ""
        for page in (result.pages or []):
            for line in (page.lines or []):
                full_text += (line.content or "") + "\n"
        ocr_result = parse_license_card_text(full_text)
    except Exception as e:
        logging.exception("license OCR failed")
        return _json_response({"error": "ocr_failed", "detail": str(e)}, 500)

    # 社員データと突合
    emp = find_active_employee_by_shain(shain_no)
    matches: Dict[str, bool] = {}
    if emp:
        emp_birthday = _utc_to_jst_date(emp.get(F_BIRTHDAY))
        ocr_bd = ocr_result.get('birthday')
        if emp_birthday and ocr_bd:
            try:
                matches['birthday'] = (_dt.date.fromisoformat(ocr_bd) == emp_birthday)
            except Exception:
                pass
    return _json_response({"ocr": ocr_result, "matches": matches})


@app.route(route="license/submit", methods=["POST", "OPTIONS"])
def license_submit(req: func.HttpRequest) -> func.HttpResponse:
    """運転免許証 表/裏画像 + 確認済みデータを社員ファイルに保存 + 社員List PATCH。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    shain_no = int(payload["shainNo"])
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    front_data = body.get("frontImage")
    back_data = body.get("backImage")
    if not (front_data and back_data):
        return _json_response({"error": "missing_images"}, 400)
    try:
        front_bytes = base64.b64decode(_strip_data_url(front_data))
        back_bytes = base64.b64decode(_strip_data_url(back_data))
    except Exception as e:
        return _json_response({"error": "invalid_base64", "detail": str(e)}, 400)
    if len(front_bytes) > 10 * 1024 * 1024 or len(back_bytes) > 10 * 1024 * 1024:
        return _json_response({"error": "image_too_large", "maxMB": 10}, 400)
    wait = check_rate_limit(shain_no, "zairyu")
    if wait is not None:
        return _json_response({"error": "rate_limited", "retryAfterSeconds": wait},
                              429, extra_headers={"Retry-After": str(wait)})
    try:
        emp = find_active_employee_by_shain(shain_no)
        if not emp:
            return _json_response({"error": "not_active"}, 403)
        # 車通勤判定 (車通勤以外には提出を許可しない)
        if not commutes_by_car(emp):
            return _json_response({"error": "not_commute_by_car"}, 403)
        buka_text = emp.get(F_BUKA) or ""
        shain_folder = find_shain_folder_url(shain_no, buka_text)
        if not shain_folder:
            return _json_response({"error": "shain_folder_not_found", "buka": buka_text}, 404)
        if not _validate_shainfile_path(shain_folder):
            return _json_response({"error": "invalid_folder_path"}, 500)
        license_folder = f"{shain_folder}/{int(shain_no)}　免許証"
        if not _validate_shainfile_path(license_folder):
            return _json_response({"error": "invalid_folder_path"}, 500)
        sp_create_folder_if_not_exists(license_folder)
        ts = (_dt.datetime.utcnow() + _dt.timedelta(hours=9)).strftime("%Y%m%d_%H%M%S")
        front_name = _safe_filename(f"{int(shain_no)}_免許証表_{ts}.jpg")
        back_name = _safe_filename(f"{int(shain_no)}_免許証裏_{ts}.jpg")
        sp_upload_file(license_folder, front_name, front_bytes)
        sp_upload_file(license_folder, back_name, back_bytes)
        # 社員 List PATCH (確認済みデータから)
        confirmed = body.get("confirmedData") or {}
        updated_fields: List[str] = []
        if confirmed:
            patch_fields: Dict[str, Any] = {}
            num = (confirmed.get("licenseNumber") or "").strip()
            if num and num.isdigit() and len(num) <= 12:
                patch_fields[F_MENKYO_NUMBER] = int(num)
                updated_fields.append("licenseNumber")
            kigen = (confirmed.get("licenseExpiry") or "").strip()
            if kigen:
                try:
                    _dt.date.fromisoformat(kigen)
                    patch_fields[F_MENKYO_KIGEN] = kigen + "T00:00:00Z"
                    updated_fields.append("licenseExpiry")
                except Exception:
                    pass
            taiken = (confirmed.get("licenseDate") or "").strip()
            if taiken:
                try:
                    _dt.date.fromisoformat(taiken)
                    patch_fields[F_MENKYO_TAIKEN] = taiken + "T00:00:00Z"
                    updated_fields.append("licenseDate")
                except Exception:
                    pass
            lic_type = (confirmed.get("licenseType") or "").strip()
            if lic_type and len(lic_type) <= 100:
                patch_fields[F_MENKYO_TYPE] = lic_type
                updated_fields.append("licenseType")
            if patch_fields:
                target_id = emp.get("Id")
                if target_id:
                    try:
                        sp_patch_item(LIST_SHAIN, int(target_id), patch_fields)
                    except Exception as e:
                        logging.exception("update employee license failed")
        return _json_response({
            "ok": True,
            "folder": license_folder,
            "files": [front_name, back_name],
            "updatedFields": updated_fields,
        })
    except Exception as e:
        logging.exception("license_submit failed")
        return _json_response({"error": "internal", "detail": str(e)}, 500)


# ====== 車検証・自賠責・任意保険 OCR (Phase 3b/3c) ======

# 保険会社の既知リスト (自賠責・任意 共通、スペース無し比較)
HOKEN_KAISHA_KNOWN = [
    "あいおいニッセイ同和損害保険", "あいおいニッセイ同和", "あいおい",
    "三井住友海上火災保険", "三井住友海上", "三井ダイレクト損害保険", "三井ダイレクト",
    "東京海上日動火災保険", "東京海上日動", "東京海上",
    "損害保険ジャパン", "損保ジャパン", "ソニー損害保険", "ソニー損保",
    "チューリッヒ保険", "チューリッヒ", "アクサ損害保険", "アクサダイレクト", "アクサ",
    "イーデザイン損害保険", "イーデザイン損保", "SBI損害保険", "SBI損保",
    "共栄火災海上保険", "共栄火災", "日新火災海上保険", "日新火災",
    "楽天損害保険", "楽天損保", "セコム損害保険", "セコム損保",
    "AIG損害保険", "AIG損保", "大同火災海上保険", "大同火災",
    "富士火災", "朝日火災", "セゾン自動車火災保険", "おとなの自動車保険",
]

# 自動車メーカーの既知リスト
CAR_MAKER_KNOWN = [
    "トヨタ", "ホンダ", "ニッサン", "日産", "マツダ", "スバル", "スズキ", "ダイハツ",
    "三菱", "ミツビシ", "レクサス", "いすゞ", "イスズ", "ヒノ", "日野",
    "メルセデス", "ベンツ", "BMW", "アウディ", "フォルクスワーゲン", "VW",
    "ボルボ", "プジョー", "ルノー", "フィアット", "ポルシェ", "MINI", "ジープ",
]


def _wareki_to_year(era: str, n: int) -> Optional[int]:
    """和暦元号 + 年数 → 西暦年。"""
    if not n:
        return None
    if "令和" in era or era == "R":
        return 2018 + n
    if "平成" in era or era == "H":
        return 1988 + n
    if "昭和" in era or era == "S":
        return 1925 + n
    return None


def _find_wareki_date(text: str, label_patterns: List[str]) -> Optional[str]:
    """ラベル近傍から和暦/西暦の日付 (YYYY-MM-DD) を抽出。"""
    for lbl in label_patterns:
        # ラベル + その後 40 文字以内の日付
        m = re.search(lbl + r'[\s\S]{0,40}?(?:(昭和|平成|令和)\s*(\d{1,2})|(\d{4}))\s*[年\.\-/]\s*(\d{1,2})\s*[月\.\-/]\s*(\d{1,2})', text)
        if m:
            if m.group(3):  # 西暦
                y = int(m.group(3))
            else:
                y = _wareki_to_year(m.group(1), int(m.group(2)))
            if y:
                return f"{y}-{int(m.group(4)):02d}-{int(m.group(5)):02d}"
    return None


def _match_known(text_ns: str, known_list: List[str]) -> Optional[str]:
    for item in known_list:
        if re.sub(r'[\s　]+', '', item) in text_ns:
            return item
    return None


def parse_shaken_jibaiseki_text(shaken_text: str, jibai_text: str) -> Dict[str, Any]:
    """車検証 + 自賠責証明書 の OCR テキストから項目抽出。"""
    out: Dict[str, Any] = {"rawShaken": shaken_text, "rawJibaiseki": jibai_text}
    st = shaken_text or ""
    st_ns = re.sub(r'[\s　]+', '', st)

    # --- 車検証 ---
    # 登録番号 (ナンバー): 「静岡 500 あ 12-34」「浜松300 さ 1234」
    m = re.search(r'([一-龿]{1,4}\s*\d{2,3}\s*[ぁ-んァ-ヶA-Z]\s*[\d\-\s]{2,5}\d)', st)
    if m:
        out['carNumber'] = re.sub(r'\s+', ' ', m.group(1)).strip()

    # メーカー
    maker = _match_known(st_ns, CAR_MAKER_KNOWN)
    if maker:
        out['carMaker'] = maker

    # 車名 (「車名」ラベル後)
    m = re.search(r'車\s*名[\s:：]*([^\n\d]{1,20})', st)
    if m:
        v = m.group(1).strip()
        if v and len(v) <= 20:
            out['carName'] = v

    # 総排気量 (X.XX L or XXXX cc)
    m = re.search(r'(?:総排気量|排気量)[^\d]{0,8}(\d[\.\d]*)\s*(?:L|ℓ|リットル)', st)
    if m:
        out['haikiryo'] = m.group(1) + 'L'
    else:
        m = re.search(r'(\d{3,4})\s*cc', st, re.IGNORECASE)
        if m:
            out['haikiryo'] = m.group(1) + 'cc'

    # 初度登録年月
    init_d = _find_wareki_date(st, [r'初度登録年月', r'初年度登録'])
    if init_d:
        out['shonendo'] = init_d[:7]  # YYYY-MM まで
    else:
        m = re.search(r'初度登録[\s\S]{0,30}?(?:(令和|平成)\s*(\d{1,2}))\s*年\s*(\d{1,2})\s*月', st)
        if m:
            y = _wareki_to_year(m.group(1), int(m.group(2)))
            if y:
                out['shonendo'] = f"{y}-{int(m.group(3)):02d}"

    # 車検満了日 (有効期間の満了する日)
    shaken_kigen = _find_wareki_date(st, [r'有効期間の満了する日', r'有効期間', r'満了'])
    if shaken_kigen:
        out['shakenKigen'] = shaken_kigen

    # --- 自賠責証明書 ---
    jt = jibai_text or ""
    jt_ns = re.sub(r'[\s　]+', '', jt)
    j_kaisha = _match_known(jt_ns, HOKEN_KAISHA_KNOWN)
    if j_kaisha:
        out['jibaiKaisha'] = j_kaisha
    # 証明書番号/証券番号
    m = re.search(r'(?:証明書番号|証券番号|証明書No|No)[\s:：.]*([A-Z0-9\-]{5,20})', jt)
    if m:
        out['jibaiShoken'] = m.group(1).strip()
    # 保険期間 満了 (後の日付 = 満了)
    j_kigen = _find_wareki_date(jt, [r'満了', r'保険期間', r'まで'])
    if j_kigen:
        out['jibaiKigen'] = j_kigen

    return out


def parse_nini_hoken_text(text: str) -> Dict[str, Any]:
    """任意保険証券の OCR テキストから項目抽出。"""
    out: Dict[str, Any] = {"rawText": text}
    t = text or ""
    t_ns = re.sub(r'[\s　]+', '', t)

    kaisha = _match_known(t_ns, HOKEN_KAISHA_KNOWN)
    if kaisha:
        out['niniKaisha'] = kaisha
    # 証券番号
    m = re.search(r'(?:証券番号|証券No|保険証券番号)[\s:：.]*([A-Z0-9\-]{5,25})', t)
    if m:
        out['niniShoken'] = m.group(1).strip()
    # 保険期間: 開始 ～ 満了 (2つの日付。最初=開始、最後=満了)
    dates = []
    for m in re.finditer(r'(?:(令和|平成)\s*(\d{1,2})|(\d{4}))\s*[年\.\-/]\s*(\d{1,2})\s*[月\.\-/]\s*(\d{1,2})', t):
        if m.group(3):
            y = int(m.group(3))
        else:
            y = _wareki_to_year(m.group(1), int(m.group(2)))
        if y and 2000 <= y <= 2100:
            dates.append(f"{y}-{int(m.group(4)):02d}-{int(m.group(5)):02d}")
    if dates:
        out['niniKaishi'] = dates[0]
        if len(dates) >= 2:
            out['niniKigen'] = dates[-1]
    return out


@app.route(route="shaken/ocr", methods=["POST", "OPTIONS"])
def shaken_ocr(req: func.HttpRequest) -> func.HttpResponse:
    """車検証 + 自賠責証明書 画像から項目を OCR 抽出。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    shain_no = int(payload["shainNo"])
    wait = check_rate_limit(shain_no, "zairyu")
    if wait is not None:
        return _json_response({"error": "rate_limited", "retryAfterSeconds": wait}, 429,
                              extra_headers={"Retry-After": str(wait)})
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    shaken_data = body.get("shakenImage")
    jibai_data = body.get("jibaiImage")
    if not (shaken_data and jibai_data):
        return _json_response({"error": "missing_images"}, 400)
    try:
        shaken_bytes = base64.b64decode(_strip_data_url(shaken_data))
        jibai_bytes = base64.b64decode(_strip_data_url(jibai_data))
    except Exception as e:
        return _json_response({"error": "invalid_base64", "detail": str(e)}, 400)
    if len(shaken_bytes) > 10 * 1024 * 1024 or len(jibai_bytes) > 10 * 1024 * 1024:
        return _json_response({"error": "image_too_large", "maxMB": 10}, 400)
    try:
        shaken_txt = _run_read_ocr(shaken_bytes)
        jibai_txt = _run_read_ocr(jibai_bytes)
        result = parse_shaken_jibaiseki_text(shaken_txt, jibai_txt)
    except Exception as e:
        logging.exception("shaken OCR failed")
        return _json_response({"error": "ocr_failed", "detail": str(e)}, 500)
    return _json_response({"ocr": result})


@app.route(route="shaken/submit", methods=["POST", "OPTIONS"])
def shaken_submit(req: func.HttpRequest) -> func.HttpResponse:
    """車検証 + 自賠責 画像を社員ファイルに保存 + 社員List PATCH。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    shain_no = int(payload["shainNo"])
    wait = check_rate_limit(shain_no, "zairyu")
    if wait is not None:
        return _json_response({"error": "rate_limited", "retryAfterSeconds": wait}, 429,
                              extra_headers={"Retry-After": str(wait)})
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    shaken_data = body.get("shakenImage")
    jibai_data = body.get("jibaiImage")
    if not (shaken_data and jibai_data):
        return _json_response({"error": "missing_images"}, 400)
    try:
        shaken_bytes = base64.b64decode(_strip_data_url(shaken_data))
        jibai_bytes = base64.b64decode(_strip_data_url(jibai_data))
    except Exception as e:
        return _json_response({"error": "invalid_base64", "detail": str(e)}, 400)
    try:
        emp = find_active_employee_by_shain(shain_no)
        if not emp:
            return _json_response({"error": "not_active"}, 403)
        if not commutes_by_car(emp):
            return _json_response({"error": "not_commute_by_car"}, 403)
        buka_text = emp.get(F_BUKA) or ""
        shain_folder = find_shain_folder_url(shain_no, buka_text)
        if not shain_folder:
            return _json_response({"error": "shain_folder_not_found", "buka": buka_text}, 404)
        folder = f"{shain_folder}/{int(shain_no)}　車検証・自賠責"
        if not _validate_shainfile_path(folder):
            return _json_response({"error": "invalid_folder_path"}, 500)
        sp_create_folder_if_not_exists(folder)
        ts = (_dt.datetime.utcnow() + _dt.timedelta(hours=9)).strftime("%Y%m%d_%H%M%S")
        n1 = _safe_filename(f"{int(shain_no)}_車検証_{ts}.jpg")
        n2 = _safe_filename(f"{int(shain_no)}_自賠責_{ts}.jpg")
        sp_upload_file(folder, n1, shaken_bytes)
        sp_upload_file(folder, n2, jibai_bytes)
        confirmed = body.get("confirmedData") or {}
        updated: List[str] = []
        if confirmed:
            pf2: Dict[str, Any] = {}
            _set_text(pf2, F_CAR_NAME, confirmed.get("carName"), 60, updated, "carName")
            _set_text(pf2, F_CAR_NUMBER, confirmed.get("carNumber"), 30, updated, "carNumber")
            _set_text(pf2, F_HAIKIRYO, confirmed.get("haikiryo"), 20, updated, "haikiryo")
            _set_text(pf2, F_SHONENDO, confirmed.get("shonendo"), 20, updated, "shonendo")
            _set_text(pf2, F_JIBAI_SHOKEN, confirmed.get("jibaiShoken"), 30, updated, "jibaiShoken")
            _set_date(pf2, F_SHAKEN_KIGEN, confirmed.get("shakenKigen"), updated, "shakenKigen")
            _set_date(pf2, F_JIBAI_KIGEN, confirmed.get("jibaiKigen"), updated, "jibaiKigen")
            # Choice 系 (メーカー/自賠責会社) は値がリストに無いと失敗するので best-effort
            _set_text(pf2, F_CAR_MAKER, confirmed.get("carMaker"), 30, updated, "carMaker")
            _set_text(pf2, F_JIBAI_KAISHA, confirmed.get("jibaiKaisha"), 40, updated, "jibaiKaisha")
            if pf2:
                try:
                    sp_patch_item(LIST_SHAIN, int(emp.get("Id")), pf2)
                except Exception:
                    logging.exception("shaken patch failed (image saved)")
        return _json_response({"ok": True, "folder": folder, "files": [n1, n2], "updatedFields": updated})
    except Exception as e:
        logging.exception("shaken_submit failed")
        return _json_response({"error": "internal", "detail": str(e)}, 500)


@app.route(route="hoken/ocr", methods=["POST", "OPTIONS"])
def hoken_ocr(req: func.HttpRequest) -> func.HttpResponse:
    """任意保険証券 画像から項目を OCR 抽出。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    shain_no = int(payload["shainNo"])
    wait = check_rate_limit(shain_no, "zairyu")
    if wait is not None:
        return _json_response({"error": "rate_limited", "retryAfterSeconds": wait}, 429,
                              extra_headers={"Retry-After": str(wait)})
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    img_data = body.get("image")
    if not img_data:
        return _json_response({"error": "missing_image"}, 400)
    try:
        img_bytes = base64.b64decode(_strip_data_url(img_data))
    except Exception as e:
        return _json_response({"error": "invalid_base64", "detail": str(e)}, 400)
    if len(img_bytes) > 10 * 1024 * 1024:
        return _json_response({"error": "image_too_large", "maxMB": 10}, 400)
    try:
        txt = _run_read_ocr(img_bytes)
        result = parse_nini_hoken_text(txt)
    except Exception as e:
        logging.exception("hoken OCR failed")
        return _json_response({"error": "ocr_failed", "detail": str(e)}, 500)
    return _json_response({"ocr": result})


@app.route(route="hoken/submit", methods=["POST", "OPTIONS"])
def hoken_submit(req: func.HttpRequest) -> func.HttpResponse:
    """任意保険証券 画像を社員ファイルに保存 + 社員List PATCH。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    shain_no = int(payload["shainNo"])
    wait = check_rate_limit(shain_no, "zairyu")
    if wait is not None:
        return _json_response({"error": "rate_limited", "retryAfterSeconds": wait}, 429,
                              extra_headers={"Retry-After": str(wait)})
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    img_data = body.get("image")
    if not img_data:
        return _json_response({"error": "missing_image"}, 400)
    try:
        img_bytes = base64.b64decode(_strip_data_url(img_data))
    except Exception as e:
        return _json_response({"error": "invalid_base64", "detail": str(e)}, 400)
    try:
        emp = find_active_employee_by_shain(shain_no)
        if not emp:
            return _json_response({"error": "not_active"}, 403)
        if not commutes_by_car(emp):
            return _json_response({"error": "not_commute_by_car"}, 403)
        buka_text = emp.get(F_BUKA) or ""
        shain_folder = find_shain_folder_url(shain_no, buka_text)
        if not shain_folder:
            return _json_response({"error": "shain_folder_not_found", "buka": buka_text}, 404)
        folder = f"{shain_folder}/{int(shain_no)}　任意保険"
        if not _validate_shainfile_path(folder):
            return _json_response({"error": "invalid_folder_path"}, 500)
        sp_create_folder_if_not_exists(folder)
        ts = (_dt.datetime.utcnow() + _dt.timedelta(hours=9)).strftime("%Y%m%d_%H%M%S")
        n1 = _safe_filename(f"{int(shain_no)}_任意保険証券_{ts}.jpg")
        sp_upload_file(folder, n1, img_bytes)
        confirmed = body.get("confirmedData") or {}
        updated: List[str] = []
        if confirmed:
            pf2: Dict[str, Any] = {}
            _set_text(pf2, F_NINI_SHOKEN, confirmed.get("niniShoken"), 30, updated, "niniShoken")
            _set_text(pf2, F_NINI_KAISHA, confirmed.get("niniKaisha"), 40, updated, "niniKaisha")
            _set_date(pf2, F_NINI_KAISHI, confirmed.get("niniKaishi"), updated, "niniKaishi")
            _set_date(pf2, F_NINI_KIGEN, confirmed.get("niniKigen"), updated, "niniKigen")
            if pf2:
                try:
                    sp_patch_item(LIST_SHAIN, int(emp.get("Id")), pf2)
                except Exception:
                    logging.exception("hoken patch failed (image saved)")
        return _json_response({"ok": True, "folder": folder, "files": [n1], "updatedFields": updated})
    except Exception as e:
        logging.exception("hoken_submit failed")
        return _json_response({"error": "internal", "detail": str(e)}, 500)


def _set_text(d: Dict[str, Any], key: str, val: Any, maxlen: int, log: List[str], name: str) -> None:
    if val is None:
        return
    s = str(val).strip()
    if s and len(s) <= maxlen:
        d[key] = s
        log.append(name)


def _set_date(d: Dict[str, Any], key: str, val: Any, log: List[str], name: str) -> None:
    if not val:
        return
    s = str(val).strip()
    try:
        _dt.date.fromisoformat(s)
        d[key] = s + "T00:00:00Z"
        log.append(name)
    except Exception:
        pass


# ====== 在留カード提出 ======
def _strip_data_url(data: str) -> str:
    """'data:image/jpeg;base64,XXXX' → 'XXXX'"""
    if not data:
        return ""
    if data.startswith("data:"):
        i = data.find(",")
        return data[i + 1:] if i > 0 else data
    return data


def sp_create_folder_if_not_exists(server_relative_url: str, site_url: str = SITE_TEAMSTEPUP) -> None:
    """SP フォルダを作成 (既存ならスキップ)。"""
    body = {
        "__metadata": {"type": "SP.Folder"},
        "ServerRelativeUrl": server_relative_url,
    }
    h = {
        "Authorization": f"Bearer {_get_sp_token()}",
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
    }
    r = requests.post(f"{site_url}/_api/web/Folders", headers=h, data=json.dumps(body), timeout=60)
    if r.status_code in (200, 201):
        return
    # 既存チェック (GET で確認)
    check = requests.get(
        f"{site_url}/_api/web/GetFolderByServerRelativeUrl('{server_relative_url}')?$select=Name",
        headers={"Authorization": f"Bearer {_get_sp_token()}", "Accept": "application/json;odata=nometadata"},
        timeout=30,
    )
    if check.ok:
        return  # 既存
    raise RuntimeError(f"folder create failed: {r.status_code} {r.text[:300]}")


def sp_upload_file(server_relative_folder_url: str, file_name: str, file_bytes: bytes,
                   site_url: str = SITE_TEAMSTEPUP) -> str:
    """SP の指定フォルダにファイルアップロード (上書き可)。"""
    # URL エンコード対応 (ファイル名に日本語を含む場合)
    from urllib.parse import quote
    fn_encoded = quote(file_name, safe="")
    api = f"{site_url}/_api/web/GetFolderByServerRelativeUrl('{server_relative_folder_url}')/Files/add(url='{fn_encoded}',overwrite=true)"
    h = {
        "Authorization": f"Bearer {_get_sp_token()}",
        "Accept": "application/json;odata=nometadata",
        "Content-Type": "application/octet-stream",
    }
    r = requests.post(api, headers=h, data=file_bytes, timeout=180)
    if not r.ok:
        raise RuntimeError(f"upload failed: {r.status_code} {r.text[:300]}")
    return file_name


def find_shain_folder_url(shain_no: int, buka_text: str) -> Optional[str]:
    """社員ファイル/{派遣先}/[工程]/{shainNo氏名} の構造から個人フォルダの ServerRelativeUrl を探す。"""
    bno = parse_buka_no(buka_text)
    if not bno:
        return None
    sn_str = str(int(shain_no))
    h = {"Authorization": f"Bearer {_get_sp_token()}", "Accept": "application/json;odata=nometadata"}
    # 社員ファイル直下
    try:
        r = requests.get(
            f"{SITE_TEAMSTEPUP}/_api/web/GetFolderByServerRelativeUrl('{SHAINFILE_ROOT}')/Folders?$select=Name,ServerRelativeUrl&$top=500",
            headers=h, timeout=30,
        )
        r.raise_for_status()
    except Exception as e:
        logging.exception("root folder list failed")
        raise
    root_folders = r.json().get("value", [])
    # 派遣先フォルダ (例 077ASTI㈱掛川)
    haken_folder = None
    for f in root_folders:
        nm = f.get("Name", "")
        if nm.startswith(bno):
            haken_folder = f
            break
    if not haken_folder:
        return None
    # 派遣先フォルダ直下
    try:
        r2 = requests.get(
            f"{SITE_TEAMSTEPUP}/_api/web/GetFolderByServerRelativeUrl('{haken_folder['ServerRelativeUrl']}')/Folders?$select=Name,ServerRelativeUrl&$top=500",
            headers=h, timeout=30,
        )
        r2.raise_for_status()
    except Exception:
        return None
    sub_folders = r2.json().get("value", [])
    # 直接ヒット (社員番号で始まる個人フォルダ)
    for sf in sub_folders:
        if sf.get("Name", "").startswith(sn_str):
            return sf["ServerRelativeUrl"]
    # 工程フォルダの可能性: 一段深く探す
    for sf in sub_folders:
        try:
            r3 = requests.get(
                f"{SITE_TEAMSTEPUP}/_api/web/GetFolderByServerRelativeUrl('{sf['ServerRelativeUrl']}')/Folders?$select=Name,ServerRelativeUrl&$top=300",
                headers=h, timeout=20,
            )
            if r3.ok:
                for ssf in r3.json().get("value", []):
                    if ssf.get("Name", "").startswith(sn_str):
                        return ssf["ServerRelativeUrl"]
        except Exception:
            continue
    return None


@app.route(route="zairyu/submit", methods=["POST", "OPTIONS"])
def zairyu_submit(req: func.HttpRequest) -> func.HttpResponse:
    """在留カード表/裏 画像を SP 社員ファイル下の「在留カード」フォルダに保存。
    Body: { frontImage: <base64 or dataURL>, backImage: <base64 or dataURL> }
    """
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    shain_no = int(payload["shainNo"])
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    front_data = body.get("frontImage")
    back_data = body.get("backImage")
    if not (front_data and back_data):
        return _json_response({"error": "missing_images"}, 400)
    # base64 デコード
    try:
        front_bytes = base64.b64decode(_strip_data_url(front_data))
        back_bytes = base64.b64decode(_strip_data_url(back_data))
    except Exception as e:
        return _json_response({"error": "invalid_base64", "detail": str(e)}, 400)
    # サイズ制限 (各 10MB まで)
    if len(front_bytes) > 10 * 1024 * 1024 or len(back_bytes) > 10 * 1024 * 1024:
        return _json_response({"error": "image_too_large", "maxMB": 10}, 400)
    # レート制限チェック (5/min)
    wait = check_rate_limit(shain_no, "zairyu")
    if wait is not None:
        return _json_response({"error": "rate_limited", "retryAfterSeconds": wait},
                              429, extra_headers={"Retry-After": str(wait)})
    try:
        emp = find_active_employee_by_shain(shain_no)
        if not emp:
            return _json_response({"error": "not_active"}, 403)
        buka_text = emp.get(F_BUKA) or ""
        # 社員フォルダ特定
        shain_folder = find_shain_folder_url(shain_no, buka_text)
        if not shain_folder:
            return _json_response({
                "error": "shain_folder_not_found",
                "buka": buka_text,
                "hint": "社員ファイル/<派遣先>/<工程>/<社員番号 氏名> の構造で見つかりませんでした",
            }, 404)
        # ※ Path Traversal 防止: 社員ファイル配下であることを必ず検証
        if not _validate_shainfile_path(shain_folder):
            logging.error(f"Invalid shain folder path detected: {shain_folder!r}")
            return _json_response({"error": "invalid_folder_path"}, 500)
        # 在留カード サブフォルダを作成
        zairyu_folder = f"{shain_folder}/{int(shain_no)}　在留カード"
        if not _validate_shainfile_path(zairyu_folder):
            return _json_response({"error": "invalid_folder_path"}, 500)
        sp_create_folder_if_not_exists(zairyu_folder)
        # タイムスタンプ付きファイル名 (固定パターン、ユーザー入力由来の文字列は使わない)
        ts = (_dt.datetime.utcnow() + _dt.timedelta(hours=9)).strftime("%Y%m%d_%H%M%S")
        front_name = _safe_filename(f"{int(shain_no)}_在留カード表_{ts}.jpg")
        back_name = _safe_filename(f"{int(shain_no)}_在留カード裏_{ts}.jpg")
        sp_upload_file(zairyu_folder, front_name, front_bytes)
        sp_upload_file(zairyu_folder, back_name, back_bytes)
        # 社員 List の在留関連フィールドを更新 (OCR で確認済みの値)
        confirmed = body.get("confirmedData") or {}
        updated_fields: List[str] = []
        if confirmed:
            patch_fields: Dict[str, Any] = {}
            shikaku = (confirmed.get("zairyuShikaku") or "").strip()
            if shikaku and len(shikaku) <= 100:
                patch_fields[F_ZAIRYU_SHIKAKU] = shikaku
                updated_fields.append("zairyuShikaku")
            kigen = (confirmed.get("zairyuKigen") or "").strip()
            if kigen:
                try:
                    _dt.date.fromisoformat(kigen)  # 妥当性
                    patch_fields[F_ZAIRYU_KIGEN] = kigen + "T00:00:00Z"
                    updated_fields.append("zairyuKigen")
                except Exception:
                    pass
            card_no = (confirmed.get("cardNumber") or "").strip()
            if card_no and len(card_no) <= 20:
                # 在留：備考1 に 「AB12345678CD (提出 2026/05/29)」 形式で記録
                ts = (_dt.datetime.utcnow() + _dt.timedelta(hours=9)).strftime("%Y/%m/%d")
                patch_fields[F_ZAIRYU_BIKO] = f"{card_no} (提出 {ts})"
                updated_fields.append("cardNumber")
            if patch_fields:
                target_id = emp.get("Id") if emp else None
                if target_id:
                    try:
                        sp_patch_item(LIST_SHAIN, int(target_id), patch_fields)
                    except Exception as e:
                        logging.exception("update employee zairyu failed")
                        # 画像保存は成功しているので警告のみ
        return _json_response({
            "ok": True,
            "folder": zairyu_folder,
            "files": [front_name, back_name],
            "updatedFields": updated_fields,
        })
    except Exception as e:
        logging.exception("zairyu_submit failed")
        return _json_response({"error": "internal", "detail": str(e)}, 500)


@app.route(route="health", methods=["GET"])
def health(req: func.HttpRequest) -> func.HttpResponse:
    """ヘルスチェック。SP API 疎通も確認。"""
    sp_ok = False
    sp_error = None
    try:
        items = sp_get_items(LIST_SHAIN, select="Id", top=1)
        sp_ok = len(items) >= 0
    except Exception as e:
        sp_error = str(e)
    return _json_response({
        "ok": True,
        "spOk": sp_ok,
        "spError": sp_error,
        "jwtConfigured": bool(JWT_SECRET),
        "timestamp": _dt.datetime.utcnow().isoformat(),
    })
