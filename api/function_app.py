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
import datetime as _dt
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


@app.route(route="auth/login", methods=["POST", "OPTIONS"])
def auth_login(req: func.HttpRequest) -> func.HttpResponse:
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

    # ロックアウトチェック
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
    token = jwt_issue(shain_no)
    profile = _employee_to_profile(emp)
    return _json_response({"token": token, "profile": profile})


def _employee_to_profile(emp: Dict[str, Any]) -> Dict[str, Any]:
    buka_text = emp.get(F_BUKA) or ""
    return {
        "shainNo": emp.get(F_SHAIN_NO),
        "name": emp.get(F_SHAIN_NAME),
        "zairyuName": emp.get(F_ZAIRYU_NAME) or "",
        "hakensaki": strip_buka_prefix(buka_text),
        "bukaRaw": buka_text,
        "bukaNo": parse_buka_no(buka_text),
        "ginko": emp.get(F_GINKO),
        "shiten": emp.get(F_SHITEN),
        "kouza": emp.get(F_KOUZA),
        "meigi": emp.get(F_MEIGI),
        "zaiyokuSyubetu": emp.get(F_ZAIYOKU),
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
