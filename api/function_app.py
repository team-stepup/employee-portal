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
import tempfile
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
LIST_SHORUI = "772de2a2-79b2-4ead-9361-f57a969b8002"  # 必要書類申請 (Teams List監視通知)
LIST_SOUGEI = "60569629-38e0-4dcb-9a73-c5f2904e3036"  # 送迎連絡 (派遣先別・帰り便/終業時間。当日表示+恒久記録)
LIST_KYUJIN = "3b062cf9-73a0-4539-b10c-fd3edcc0e8b3"  # 求人 (HP採用情報・役員がyukyu-appから掲載。公開APIで匿名配信)
LIST_OUBO = "7784b615-c0e1-4587-a10e-3551a2bd47cf"   # 応募 (HP採用ページの応募フォーム→POST /apply で保存。総務へ通知)
LIST_SOUGEI_KIROKU = "ec0a16dc-6bba-49cf-a901-d5b3907553ca"  # 送迎記録 (運転手が1タップで実績記録→月次集計して総務へ)
LIST_JIKO = "56dcb3c2-472b-4472-9d32-904f67cc3afa"           # 事故報告 (運転手の交通事故報告→送迎関係チャネル+Push即時通知)

# 送迎記録 List フィールド (ASCII名なので read/write とも OData_ プレフィックス無し)
SK_SHAINNO = "ShainNo"       # 社員番号 (Number)
SK_NAME = "DriverName"       # 氏名カナ
SK_DATE = "RecordDate"       # 送迎日 (DateOnly)
SK_DEST = "Destination"      # 送迎先 (ASTI浜松/ユーシン/ASTI掛川)
SK_CAT = "Category"          # 区分 (行き/帰り/遅刻/早退/病院対応)
SK_SLOT = "TimeSlot"         # 時刻 (行き/帰りのみ。例 "17:00")
SK_MEMO = "Memo"             # 補足メモ (任意)
SK_STATUS = "Status"         # active / cancelled
# 送迎連絡 List フィールド (ASCII名なので read/write とも OData_ プレフィックス無し)
SG_DATE = "TargetDate"      # 対象日 (DateOnly)
SG_SHAINNO = "ShainNo"      # 社員番号 (人単位配信用)
SG_HAKENSAKI = "Hakensaki"  # 派遣先 (prefix除去済み名)
SG_ENDTIME = "EndTime"      # 終業時間 (例 "17:00")
SG_VEHICLE = "Vehicle"      # 帰りの車両 (例 "1号車")
SG_MEMO = "Memo"            # 補足メモ
SG_SENTBY = "SentBy"        # 送信者 (staff email)

# 必要書類申請 List フィールド (POST は OData_ プレフィックス無し)
SH_SHAIN_NO = "_x793e__x54e1__x756a__x53f7_"
SH_SHAIN_NAME = "_x793e__x54e1__x540d_"
SH_HAKENSAKI = "_x6d3e__x9063__x5148_"
SH_SHURUI = "_x66f8__x985e__x7a2e__x5225_"
SH_NENDO = "_x5bfe__x8c61__x5e74__x5ea6_"
SH_BIKO = "_x5099__x8003_"
SH_STATUS = "Status"

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
F_ADDRESS = "OData__x4f4f__x6240_"                        # 住所 (Text)
F_NYUSHA = "OData__x5165__x793e__x65e5_"                  # 入社日 (DateTime)
F_PORTAL_PUSH = "PortalPushSub"                           # Web Push 購読情報 (Note, JSON) — ASCII名なのでOData_無し
F_RENEW_PLAN = "ZairyuRenewPlan"                          # 在留更新の申請予定日 (DateTime, DateOnly)
F_RENEW_NOTE = "ZairyuRenewNote"                          # 在留更新のメモ (Note)
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
F_CAR_TYPE = "OData__x5bfe__x8c61__x8eca__x4e21_"            # 対象車両 (Choice)
F_CAR_COLOR = "OData__x8eca__x306e__x8272_"                  # 車の色 (Choice)
# 自賠責保険関連フィールド
F_JIBAI_KAISHA = "OData__x81ea__x8ce0__x8cac__x4fdd__x96"    # 自賠責保険会社 (Choice)
F_JIBAI_KIGEN = "OData__x81ea__x8ce0__x8cac__x3000__x6e"     # 自賠責満了日 (DateTime)
F_JIBAI_SHOKEN = "OData__x81ea__x8ce0__x8cac__x8a3c__x52"    # 自賠責証券番号 (Text)
# 任意保険関連フィールド
F_NINI_KAISHA = "OData__x81ea__x52d5__x8eca__x4efb__x61"     # 自動車任意保険会社 (Choice)
F_NINI_KAISHI = "OData__x4efb__x610f__x4fdd__x967a__x95"     # 任意保険開始日 (DateTime)
F_NINI_KIGEN = "OData__x4efb__x610f__x4fdd__x967a__x6e"      # 任意保険満了日 (DateTime)
F_NINI_SHOKEN = "OData__x4efb__x610f__x4fdd__x967a__x8a"     # 任意保険証券番号 (Text)
# 補償内容フィールド
F_HOKEN_TAIJIN = "OData__x4fdd__x967a__xff1a__x5bfe__x4e"    # 保険：対人 (Choice)
F_HOKEN_TAIBUTSU = "OData__x4fdd__x967a__xff1a__x5bfe__x72"  # 保険：対物 (Choice)
F_JINSHIN = "OData__x4fdd__x967a__xff1a__x540c__x4e"         # 人身傷害 (Text)
F_TOJOSHA = "OData__x642d__x4e57__x8005__x50b7__x5b"         # 搭乗者傷害 (Text)
F_SHARYO_HOKEN = "OData__x8eca__x4e21__x4fdd__x967a_"        # 車両保険 (Text)

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


def commutes_by_sougei(emp: Dict[str, Any]) -> bool:
    """通勤方法フィールドから送迎利用かを判定。"""
    combined = str(emp.get(F_TSUKIN_OLD) or "") + " " + str(emp.get(F_TSUKIN_NEW) or "")
    return "送迎" in combined


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
    "kyuyu": 12,       # 給油申請 OCR/送信 12/min (写真2枚×OCR+送信)
    "sougei": 20,      # 送迎記録 20/min (1タップ記録・連続記録あり)
    "jiko": 6,         # 事故報告 6/min (写真3枚アップロードあり)
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
                 orderby: Optional[str] = None, expand: Optional[str] = None) -> List[Dict[str, Any]]:
    """SP List のアイテムを取得。ページング対応。"""
    base = f"{SITE_URL}/_api/web/lists(guid'{list_guid}')/items"
    params = [f"$top={min(top, 5000)}"]
    if select:
        params.append(f"$select={select}")
    if filter_:
        params.append(f"$filter={filter_}")
    if orderby:
        params.append(f"$orderby={orderby}")
    if expand:
        params.append(f"$expand={expand}")
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
            F_KOKUSEKI, F_TSUKIN_OLD, F_TSUKIN_NEW, F_PORTAL_PIN, F_ADDRESS, F_NYUSHA,
            F_ZAIRYU_KIGEN, F_MENKYO_KIGEN, F_SHAKEN_KIGEN, F_JIBAI_KIGEN, F_NINI_KIGEN,
            F_RENEW_PLAN, F_RENEW_NOTE,
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


def _iso_or_none(d: Any) -> Optional[str]:
    return d.isoformat() if d else None


def _today_jst() -> _dt.date:
    return (_dt.datetime.utcnow() + _dt.timedelta(hours=9)).date()


def _norm_shain(v: Any) -> str:
    """社員番号を整数文字列に正規化 (SPはNumber型で 70381.0 と float化される)。"""
    s = str(v if v is not None else "").strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def _fetch_today_sougei(shain_no: Any) -> Optional[Dict[str, Any]]:
    """当日・本人(社員番号一致)の送迎連絡レコード(最新)を返す。無ければ None。"""
    if shain_no in (None, ""):
        return None
    sn = _norm_shain(shain_no)
    today = _today_jst()
    try:
        items = sp_get_items(
            LIST_SOUGEI,
            select=",".join(["Id", SG_DATE, SG_SHAINNO, SG_HAKENSAKI, SG_ENDTIME, SG_VEHICLE, SG_MEMO]),
            filter_=f"{SG_SHAINNO} eq '{sn}'",
            orderby="Id desc", top=10,
        )
    except Exception:
        logging.warning("送迎連絡の取得に失敗")
        return None
    for it in items:
        if _utc_to_jst_date(it.get(SG_DATE)) != today:
            continue
        return {
            "endTime": it.get(SG_ENDTIME) or "",
            "vehicle": it.get(SG_VEHICLE) or "",
            "memo": it.get(SG_MEMO) or "",
            "hakensaki": it.get(SG_HAKENSAKI) or "",
            "date": today.isoformat(),
        }
    return None


def _employee_to_profile(emp: Dict[str, Any]) -> Dict[str, Any]:
    buka_text = emp.get(F_BUKA) or ""
    kokuseki = emp.get(F_KOKUSEKI) or ""
    is_sougei = commutes_by_sougei(emp)
    zairyu_kigen = _utc_to_jst_date(emp.get(F_ZAIRYU_KIGEN))
    menkyo_kigen = _utc_to_jst_date(emp.get(F_MENKYO_KIGEN))
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
        "commutesBySougei": is_sougei,
        # 弁当注文: 注文はユーシン/委託管理のみ。承認は yukyu-app(役員/森まこと/森ゆうじ)で行うため
        # ポータルには承認UIを出さない (派遣社員=注文する側のみ)
        "bentoEligible": (("ユーシン" in strip_buka_prefix(buka_text)) or ("ﾕｰｼﾝ" in strip_buka_prefix(buka_text))),
        "isBentoApprover": False,
        # 給油申請: 送迎運転手(社員番号セット)のみボタン表示
        "kyuyuEligible": _kyuyu_eligible(emp),
        # 送迎記録: 部課=094:本社（送迎者）の運転手のみボタン表示
        "sougeiKirokuEligible": _sougei_kiroku_eligible(emp),
        # 送迎の帰り便連絡 (当日・本人宛のみ。無ければ null)
        "todaySougei": _fetch_today_sougei(emp.get(F_SHAIN_NO)) if is_sougei else None,
        # 期限お知らせ用 (在留カード/免許証/車検/自賠責/任意保険)
        "zairyuKigen": zairyu_kigen.isoformat() if zairyu_kigen else None,
        "menkyoKigen": menkyo_kigen.isoformat() if menkyo_kigen else None,
        "shakenKigen": _iso_or_none(_utc_to_jst_date(emp.get(F_SHAKEN_KIGEN))),
        "jibaiKigen": _iso_or_none(_utc_to_jst_date(emp.get(F_JIBAI_KIGEN))),
        "niniKigen": _iso_or_none(_utc_to_jst_date(emp.get(F_NINI_KIGEN))),
        "zairyuRenewPlan": _iso_or_none(_utc_to_jst_date(emp.get(F_RENEW_PLAN))),
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
            F_KOKUSEKI, F_TSUKIN_OLD, F_TSUKIN_NEW, F_PORTAL_PIN, F_ADDRESS, F_NYUSHA,
            F_ZAIRYU_KIGEN, F_MENKYO_KIGEN, F_SHAKEN_KIGEN, F_JIBAI_KIGEN, F_NINI_KIGEN,
            F_RENEW_PLAN, F_RENEW_NOTE,
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


@app.route(route="jobs", methods=["GET", "OPTIONS"])
def jobs_public(req: func.HttpRequest) -> func.HttpResponse:
    """公開求人をホームページ採用情報へ配信（匿名・誰でも閲覧可）。
    GET /jobs?lang=ja|pt|en
    表示ルール: 対象「日本人向け」は ja のみ、「外国人もOK」は全言語。
    pt/en の訳は yakuJSON から、無ければ日本語列にフォールバック。
    """
    # 公開データなのでオリジンは全許可
    public_cors = {
        "Access-Control-Allow-Origin": "*",
        "Cache-Control": "public, max-age=120",
    }
    if req.method == "OPTIONS":
        return _json_response({}, 204, extra_headers=public_cors)
    lang = (req.params.get("lang") or "ja").lower()
    if lang not in ("ja", "pt", "en"):
        lang = "ja"
    try:
        items = sp_get_items(
            LIST_KYUJIN,
            select="Id,Title,basho,jikyu,jikan,yasumi,naiyo,bikou,taisho,joutai,koushinbi,yakuJSON,Modified",
            filter_="joutai eq '公開'",
            orderby="Modified desc",
            top=200,
        )
    except Exception as e:
        logging.exception("jobs fetch failed")
        return _json_response({"error": "fetch_failed", "detail": str(e)}, 500, extra_headers=public_cors)

    jobs: List[Dict[str, Any]] = []
    for it in items:
        target = it.get("taisho") or "外国人もOK"
        # 日本人向けは日本語ページのみ
        if lang != "ja" and target == "日本人向け":
            continue
        base = {
            "title": it.get("Title") or "",
            "location": it.get("basho") or "",
            "hours": it.get("jikan") or "",
            "off": it.get("yasumi") or "",
            "description": it.get("naiyo") or "",
            "remarks": it.get("bikou") or "",
        }
        if lang in ("pt", "en"):
            try:
                tr = json.loads(it.get("yakuJSON") or "{}").get(lang, {})
            except Exception:
                tr = {}
            for k in ("title", "location", "hours", "off", "description", "remarks"):
                # 訳キーは basho/jikan 等の内部名でも title/location でも拾えるように両対応
                v = tr.get(k) or tr.get({"location": "basho", "hours": "jikan", "off": "yasumi", "description": "naiyo", "remarks": "bikou"}.get(k, k))
                if v:
                    base[k] = v
        updated = it.get("koushinbi") or it.get("Modified") or ""
        jobs.append({
            "id": it.get("Id"),                # 求人ごとの固有ID(Indeed固有URL・ディープリンク用)
            **base,
            "wage": it.get("jikyu") or "",   # 時給は全言語共通
            "target": target,
            "updated": str(updated)[:10],
        })
    return _json_response({"jobs": jobs, "lang": lang}, 200, extra_headers=public_cors)


@app.route(route="jobs/indeed", methods=["GET", "OPTIONS"])
def jobs_indeed_feed(req: func.HttpRequest) -> func.HttpResponse:
    """Indeed求人フィード(XML)。Indeedが定期巡回して自動掲載する。
    GET /jobs/indeed?target=jp|global|all  (既定 jp)
      jp     = 対象「日本人向け」の求人
      global = 対象「外国人もOK」の求人
      all    = 全公開求人
    ※広告文面はSP「求人」の内容をそのまま出力。国籍除外の表現は入れないこと(法令/Indeedポリシー)。"""
    public_cors = {"Access-Control-Allow-Origin": "*", "Cache-Control": "public, max-age=300"}
    if req.method == "OPTIONS":
        return _json_response({}, 204, extra_headers=public_cors)
    target = (req.params.get("target") or "jp").lower()
    if target not in ("jp", "global", "all"):
        target = "jp"

    def xml_esc(s: str) -> str:
        # CDATA内なので基本エスケープ不要。終端記号のみ無害化。
        return str(s or "").replace("]]>", "]]&gt;")

    try:
        items = sp_get_items(
            LIST_KYUJIN,
            select="Id,Title,basho,jikyu,jikan,yasumi,naiyo,bikou,taisho,joutai,koushinbi,Modified",
            filter_="joutai eq '公開'", orderby="Modified desc", top=200)
    except Exception as e:
        logging.exception("indeed feed fetch failed")
        return func.HttpResponse(f"<error>{e}</error>", status_code=500,
                                 mimetype="application/xml", headers=public_cors)

    jobs_xml = []
    for it in items:
        tgt = it.get("taisho") or "外国人もOK"
        if target == "jp" and tgt != "日本人向け":
            continue
        if target == "global" and tgt != "外国人もOK":
            continue
        jid = it.get("Id")
        basho = it.get("basho") or ""
        m_pref = re.search(r"([^\s　]{2,4}?[都道府県])", basho)
        state = m_pref.group(1) if m_pref else "静岡県"
        # 県名を除いた残りから市区町村を抽出(「静岡県磐田市」→ city=磐田市)
        rest = basho[m_pref.end():] if m_pref else basho
        m_city = re.search(r"([^\s　都道府県]{1,6}?[市区町村])", rest)
        city = m_city.group(1) if m_city else ""
        updated = str(it.get("koushinbi") or it.get("Modified") or "")[:10]
        desc_parts = [it.get("naiyo") or ""]
        if it.get("jikyu"):
            desc_parts.append(f"【時給】{it.get('jikyu')}")
        if it.get("jikan"):
            desc_parts.append(f"【勤務時間】{it.get('jikan')}")
        if it.get("yasumi"):
            desc_parts.append(f"【休日】{it.get('yasumi')}")
        if it.get("bikou"):
            desc_parts.append(it.get("bikou"))
        desc_html = "<br>".join(xml_esc(p) for p in desc_parts if p)
        jobs_xml.append(
            "<job>"
            f"<title><![CDATA[{xml_esc(it.get('Title'))}]]></title>"
            f"<date><![CDATA[{updated}]]></date>"
            f"<referencenumber><![CDATA[{jid}]]></referencenumber>"
            f"<url><![CDATA[https://www.team-stepup.com/?job={jid}]]></url>"
            "<company><![CDATA[有限会社ステップ・アップ]]></company>"
            f"<city><![CDATA[{xml_esc(city)}]]></city>"
            f"<state><![CDATA[{xml_esc(state)}]]></state>"
            "<country><![CDATA[JP]]></country>"
            f"<description><![CDATA[{desc_html}]]></description>"
            f"<salary><![CDATA[{xml_esc(it.get('jikyu'))}]]></salary>"
            "<jobtype><![CDATA[temporary]]></jobtype>"
            "<category><![CDATA[製造・軽作業]]></category>"
            "</job>")

    xml = ('<?xml version="1.0" encoding="utf-8"?>\n<source>\n'
           '<publisher>有限会社ステップ・アップ</publisher>\n'
           '<publisherurl>https://www.team-stepup.com</publisherurl>\n'
           + "\n".join(jobs_xml) + "\n</source>\n")
    return func.HttpResponse(xml, status_code=200, mimetype="application/xml", headers=public_cors)


# ====== 応募 (HP採用ページ 公開フォーム) ======
# メール宛先: Teamsチャネル「一般」の投稿用メール + 役員個人(スマホ/PCのOutlook通知用)
APPLY_NOTIFY_EMAILS = [e.strip() for e in os.environ.get(
    "APPLY_NOTIFY_EMAILS",
    "oubosyamennsetu@team-stepup.com,h.yamashita@team-stepup.com").split(",") if e.strip()]
# Web Push宛先(yukyu-app portal 購読端末へロック画面通知)
APPLY_PUSH_EMAILS = [e.strip() for e in os.environ.get(
    "APPLY_PUSH_EMAILS",
    "h.yamashita@team-stepup.com,a.yamashita@team-stepup.com").split(",") if e.strip()]


def _apply_notify(rec: Dict[str, str]) -> None:
    """応募をメール通知(総務/採用担当)。Graph sendMail。best-effort。
    送信元 = SHORUI_MAIL_SENDER (他通知と共用)、宛先 = APPLY_NOTIFY_EMAILS。"""
    sender = os.environ.get("SHORUI_MAIL_SENDER", "").strip()
    if not sender or not APPLY_NOTIFY_EMAILS:
        return
    esc = lambda s: (str(s or "")
                     .replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
    tel = esc(rec.get("tel"))
    rows = [
        ("お名前", esc(rec.get("name"))),
        ("電話", f'<a href="tel:{tel}">{tel}</a>'),
        ("メール", esc(rec.get("mail")) or "―"),
        ("応募求人", esc(rec.get("kyujin")) or "（一般応募）"),
        ("希望勤務地", esc(rec.get("area")) or "―"),
        ("連絡可能時間帯", esc(rec.get("time")) or "―"),
        ("メッセージ", esc(rec.get("message")).replace("\n", "<br>") or "―"),
        ("言語", esc(rec.get("lang"))),
    ]
    tr = "".join(
        f'<tr><td style="padding:6px 12px;background:#f4f6f8;font-weight:600;white-space:nowrap;">{k}</td>'
        f'<td style="padding:6px 12px;">{v}</td></tr>' for k, v in rows)
    html = (
        '<div style="font-family:sans-serif;color:#1a2b3c;">'
        '<h3 style="margin:0 0 4px;">🆕 ホームページから新しい応募が届きました</h3>'
        '<p style="margin:0 0 12px;color:#c0392b;font-weight:600;">'
        '応募者は返信の速さで他社と比較します。できるだけ早くお電話ください。</p>'
        f'<table style="border-collapse:collapse;font-size:14px;">{tr}</table>'
        '<p style="margin:14px 0 0;font-size:12px;color:#888;">'
        'SharePoint「応募」リストにも保存されています。</p></div>')
    subject = f'【HP応募】{rec.get("name","")} 様' + (f'／{rec.get("kyujin")}' if rec.get("kyujin") else "")
    try:
        token = _get_graph_token()
        message = {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html},
            "toRecipients": [{"emailAddress": {"address": a}} for a in APPLY_NOTIFY_EMAILS],
        }
        r = requests.post(
            f"https://graph.microsoft.com/v1.0/users/{sender}/sendMail",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"message": message, "saveToSentItems": False}, timeout=60)
        if not r.ok:
            logging.warning(f"apply notify mail failed: {r.status_code} {r.text[:300]}")
    except Exception as e:
        logging.warning(f"apply notify failed: {e}")


@app.route(route="apply", methods=["POST", "OPTIONS"])
def apply_public(req: func.HttpRequest) -> func.HttpResponse:
    """HP採用ページからの応募を受け付け(匿名・誰でも送信可)。
    body: {name, tel, mail?, kyujin?, area?, time?, message?, lang?, company?(honeypot)}
    → SharePoint「応募」に保存 → 総務/採用担当へメール通知。"""
    public_cors = {"Access-Control-Allow-Origin": "*"}
    if req.method == "OPTIONS":
        return _json_response({}, 204, extra_headers=public_cors)
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400, extra_headers=public_cors)

    # スパム対策: honeypot("company")に値が入っていればボットとみなし静かに成功扱い
    if (body.get("company") or "").strip():
        return _json_response({"ok": True}, 200, extra_headers=public_cors)

    clip = lambda s, n: (str(s or "").strip())[:n]
    name = clip(body.get("name"), 100)
    tel = clip(body.get("tel"), 40)
    if not name or not tel:
        return _json_response({"error": "missing_required", "detail": "name and tel required"},
                              400, extra_headers=public_cors)
    lang = (body.get("lang") or "ja").lower()
    if lang not in ("ja", "pt", "en"):
        lang = "ja"
    rec = {
        "name": name,
        "tel": tel,
        "mail": clip(body.get("mail"), 120),
        "kyujin": clip(body.get("kyujin"), 200),
        "area": clip(body.get("area"), 120),
        "time": clip(body.get("time"), 120),
        "message": clip(body.get("message"), 2000),
        "lang": lang,
    }
    try:
        new_id = sp_post_item(LIST_OUBO, {
            "Title": rec["name"],
            "Denwa": rec["tel"],
            "Mail": rec["mail"],
            "Kyujin": rec["kyujin"],
            "KibouArea": rec["area"],
            "RenrakuTime": rec["time"],
            "Message": rec["message"],
            "Lang": rec["lang"],
            "Source": "HP",
            "Joutai": "未対応",
        })
    except Exception as e:
        logging.exception("apply save failed")
        return _json_response({"error": "save_failed", "detail": str(e)}, 500, extra_headers=public_cors)

    # 通知(メール)。失敗しても応募は成功扱い。
    _apply_notify(rec)
    # 承認者(役員)へロック画面Pushも best-effort
    try:
        _push_to_admins(APPLY_PUSH_EMAILS, {
            "title": "🆕 HPから応募が届きました",
            "body": f'{rec["name"]} 様／{rec["kyujin"] or "一般応募"}',
            "url": YUKYU_APP_URL,
        })
    except Exception:
        logging.warning("apply push failed")

    return _json_response({"ok": True, "id": new_id}, 200, extra_headers=public_cors)


@app.route(route="jobs/admin", methods=["GET", "OPTIONS"])
def jobs_admin_list(req: func.HttpRequest) -> func.HttpResponse:
    """役員専用: 全求人(非公開含む)を管理用に返す。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    email, err = require_staff_auth(req)
    if err:
        return err
    try:
        items = sp_get_items(
            LIST_KYUJIN,
            select="Id,Title,basho,jikyu,jikan,yasumi,naiyo,bikou,taisho,joutai,yakuJSON,Modified",
            orderby="Modified desc", top=500)
    except Exception as e:
        logging.exception("jobs admin list failed")
        return _json_response({"error": "fetch_failed", "detail": str(e)}, 500)
    jobs = [{
        "id": it.get("Id"),
        "title": it.get("Title") or "",
        "location": it.get("basho") or "",
        "wage": it.get("jikyu") or "",
        "hours": it.get("jikan") or "",
        "off": it.get("yasumi") or "",
        "description": it.get("naiyo") or "",
        "remarks": it.get("bikou") or "",
        "target": it.get("taisho") or "外国人もOK",
        "status": it.get("joutai") or "非公開",
        "updated": str(it.get("Modified") or "")[:10],
        "hasTranslation": bool(it.get("yakuJSON")),
    } for it in items]
    return _json_response({"jobs": jobs})


@app.route(route="jobs/translate", methods=["POST", "OPTIONS"])
def jobs_translate(req: func.HttpRequest) -> func.HttpResponse:
    """役員専用: 日本語項目を pt/en に翻訳して返す(保存前のプレビュー用)。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    email, err = require_staff_auth(req)
    if err:
        return err
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    ja = {k: (body.get(k) or "") for k in ("title", "location", "hours", "off", "description", "remarks")}
    try:
        import gpt_ocr
        yaku = gpt_ocr.translate_job(ja)
    except Exception as e:
        logging.exception("translate_job failed")
        return _json_response({"error": "translate_failed", "detail": str(e)}, 500)
    return _json_response({"translations": yaku})


@app.route(route="jobs/indeed-copy", methods=["POST", "OPTIONS"])
def jobs_indeed_copy(req: func.HttpRequest) -> func.HttpResponse:
    """役員専用: 求人データからIndeed直接投稿用の原稿(タイトル+本文)をAI生成。
    id があれば求人Listの indeedgenko 列にも保存する(失敗しても生成結果は返す)。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    email, err = require_staff_auth(req)
    if err:
        return err
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    fields = {k: (body.get(k) or "") for k in ("title", "location", "wage", "hours", "off", "description", "remarks", "target")}
    if not fields["title"]:
        return _json_response({"error": "title_required"}, 400)
    try:
        import gpt_ocr
        copy = gpt_ocr.generate_indeed_copy(fields)
    except Exception as e:
        logging.exception("generate_indeed_copy failed")
        return _json_response({"error": "generate_failed", "detail": str(e)}, 500)
    jid = body.get("id")
    if jid:
        try:
            sp_patch_item(LIST_KYUJIN, int(jid), {"indeedgenko": json.dumps(copy, ensure_ascii=False)})
        except Exception:
            logging.exception("indeedgenko save failed (continue)")
    return _json_response({"ok": True, "copy": copy})


@app.route(route="jobs/save", methods=["POST", "OPTIONS"])
def jobs_save(req: func.HttpRequest) -> func.HttpResponse:
    """役員専用: 求人の新規作成/更新。日本語項目を保存し pt/en 訳を yakuJSON に格納。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    email, err = require_staff_auth(req)
    if err:
        return err
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    title = (body.get("title") or "").strip()
    if not title:
        return _json_response({"error": "title_required"}, 400)
    ja = {
        "title": title,
        "location": (body.get("location") or "").strip(),
        "hours": (body.get("hours") or "").strip(),
        "off": (body.get("off") or "").strip(),
        "description": (body.get("description") or "").strip(),
        "remarks": (body.get("remarks") or "").strip(),
    }
    wage = (body.get("wage") or "").strip()
    target = body.get("target") if body.get("target") in ("日本人向け", "外国人もOK") else "外国人もOK"
    status = body.get("status") if body.get("status") in ("公開", "非公開") else "非公開"
    # 訳: クライアントが確認修正済みの訳を渡せばそれを使う。無ければ翻訳。
    yaku = body.get("translations")
    if not yaku:
        try:
            import gpt_ocr
            yaku = gpt_ocr.translate_job(ja)
        except Exception:
            logging.exception("translate_job failed (continue ja-only)")
            yaku = {}
    fields = {
        "Title": ja["title"], "basho": ja["location"], "jikyu": wage,
        "jikan": ja["hours"], "yasumi": ja["off"], "naiyo": ja["description"],
        "bikou": ja["remarks"], "taisho": target, "joutai": status,
        "yakuJSON": json.dumps(yaku, ensure_ascii=False),
    }
    try:
        jid = body.get("id")
        if jid:
            sp_patch_item(LIST_KYUJIN, int(jid), fields)
            new_id = int(jid)
        else:
            new_id = sp_post_item(LIST_KYUJIN, fields)
    except Exception as e:
        logging.exception("jobs save failed")
        return _json_response({"error": "save_failed", "detail": str(e)}, 500)
    return _json_response({"ok": True, "id": new_id, "translations": yaku})


@app.route(route="jobs/delete", methods=["POST", "OPTIONS"])
def jobs_delete(req: func.HttpRequest) -> func.HttpResponse:
    """役員専用: 求人を削除。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    email, err = require_staff_auth(req)
    if err:
        return err
    try:
        body = req.get_json()
        jid = int(body.get("id"))
    except Exception:
        return _json_response({"error": "invalid_id"}, 400)
    try:
        url = f"{SITE_URL}/_api/web/lists(guid'{LIST_KYUJIN}')/items({jid})"
        r = requests.post(url, headers={**_sp_headers(verbose=True), "IF-MATCH": "*", "X-HTTP-Method": "DELETE"}, timeout=30)
        if not r.ok:
            raise RuntimeError(f"{r.status_code} {r.text[:200]}")
    except Exception as e:
        logging.exception("jobs delete failed")
        return _json_response({"error": "delete_failed", "detail": str(e)}, 500)
    return _json_response({"ok": True})


@app.route(route="jobs/toggle", methods=["POST", "OPTIONS"])
def jobs_toggle(req: func.HttpRequest) -> func.HttpResponse:
    """役員専用: 公開状態のみ切替(翻訳は再実行しない)。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    email, err = require_staff_auth(req)
    if err:
        return err
    try:
        body = req.get_json()
        jid = int(body.get("id"))
        status = body.get("status")
    except Exception:
        return _json_response({"error": "invalid_request"}, 400)
    if status not in ("公開", "非公開"):
        return _json_response({"error": "invalid_status"}, 400)
    try:
        sp_patch_item(LIST_KYUJIN, jid, {"joutai": status})
    except Exception as e:
        logging.exception("jobs toggle failed")
        return _json_response({"error": "toggle_failed", "detail": str(e)}, 500)
    return _json_response({"ok": True, "id": jid, "status": status})


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
        # 担当者(派遣先別)へ ロック画面 Push。担当者未設定なら役員へ。
        try:
            name = (emp.get("Title") or "").strip()
            try:
                amt_disp = f"{int(str(amount).replace(',', '')):,}円"
            except Exception:
                amt_disp = f"{amount}円"
            tgt = TANTOU_EMAIL.get(tantou)
            targets = [tgt] if tgt else list(ADMIN_APPROVER_EMAILS)
            _push_to_admins(targets, {
                "title": "💴 前払い申請 承認待ち",
                "body": f"{name or shain_no} / {amt_disp} / {d.isoformat()}",
                "url": YUKYU_APP_URL, "tag": "maebari-approve", "badge": 1,
            })
        except Exception:
            logging.exception("maebarai push failed")
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


# ====== 弁当注文 (ユーシン/委託管理・当日AM9:00まで・モリマコト承認) ======
LIST_BENTO = "1d0f9e26-ad32-4fc0-8c9e-02f22a87fe35"
BENTO_MENU = {"①": 315, "④": 355, "⑥": 460}
BENTO_ORDER_LABEL = ["①", "④", "⑥"]
BENTO_DEADLINE_HOUR = 9                 # 当日 09:00 まで
MORIMAKOTO_SHAIN = "50500"              # 承認担当(モリマコト)


def _now_jst() -> _dt.datetime:
    return _dt.datetime.utcnow() + _dt.timedelta(hours=9)


def _bento_test_open() -> bool:
    """env BENTO_TEST_OPEN=1 のとき、稼働日/9:00締切のゲートを一時的に解除(テスト用)。"""
    return os.environ.get("BENTO_TEST_OPEN", "").strip() == "1"


def _bento_eligible(emp: Dict[str, Any]) -> bool:
    """ユーシン本体(002) または ユーシン委託管理(002-1) の従業員のみ。"""
    buka = strip_buka_prefix(emp.get(F_BUKA) or "")
    return ("ユーシン" in buka) or ("ﾕｰｼﾝ" in buka)


def _bento_is_working_day(d: _dt.date) -> bool:
    if d.weekday() >= 5:                # 土日
        return False
    try:
        if d in set(list_kaisha_kyujitsu_dates()):
            return False
    except Exception:
        logging.warning("会社休日の取得に失敗(弁当稼働日判定)")
    return True


def _bento_today_order(shain_str: str, d: _dt.date) -> Optional[Dict[str, Any]]:
    try:
        items = sp_get_items(
            LIST_BENTO,
            select="Id,OrderDate,MenuNo,Price,OrderStatus,RejectReason",
            filter_=f"ShainNo eq '{shain_str}' and OrderDate eq datetime'{d.isoformat()}T00:00:00'",
            orderby="Id desc", top=5)
        for it in items:
            if (it.get("OrderStatus") or "") in ("pending", "approved"):
                return it
        return items[0] if items else None
    except Exception:
        logging.exception("弁当: 当日注文の取得に失敗")
        return None


def _bento_push_morimakoto(emp_name: str, menu_no: str, price: int) -> None:
    """新規注文をモリマコトへ Push (購読していれば)。"""
    try:
        recs = sp_get_items(LIST_SHAIN,
                            select=f"Id,{F_SHAIN_NO},{F_PORTAL_PUSH}",
                            filter_=f"{F_SHAIN_NO} eq {MORIMAKOTO_SHAIN}", orderby="Id desc", top=3)
        rec = recs[0] if recs else None
        sub = (rec.get(F_PORTAL_PUSH) or "").strip() if rec else ""
        if not sub:
            return
        res = _send_web_push(sub, {"title": "🍱 弁当注文 承認待ち",
                                   "body": f"{emp_name} さん {menu_no} {price}円",
                                   "url": "/", "tag": "bento-approve", "badge": 1})
        if res == "gone" and rec:
            try:
                sp_patch_item(LIST_SHAIN, int(rec["Id"]), {F_PORTAL_PUSH: ""})
            except Exception:
                pass
    except Exception:
        logging.exception("弁当: モリマコトへのPushに失敗")


@app.route(route="bento/status", methods=["GET", "OPTIONS"])
def bento_status(req: func.HttpRequest) -> func.HttpResponse:
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    shain_str = str(payload["shainNo"])
    emp = find_active_employee_by_shain(payload["shainNo"])
    if not emp:
        return _json_response({"error": "not_active"}, 403)
    eligible = _bento_eligible(emp)
    today = _today_jst()
    working = _bento_is_working_day(today)
    deadline_passed = _now_jst().hour >= BENTO_DEADLINE_HOUR
    if _bento_test_open():                 # テストモード: ゲート解除
        working = True
        deadline_passed = False
    to = _bento_today_order(shain_str, today)
    already = bool(to and (to.get("OrderStatus") in ("pending", "approved")))
    if not eligible:
        reason = "not_eligible"
    elif not working:
        reason = "holiday"
    elif already:
        reason = "already_ordered"
    elif deadline_passed:
        reason = "deadline_passed"
    else:
        reason = "ok"
    return _json_response({
        "ok": True, "eligible": eligible, "today": today.isoformat(),
        "isWorkingDay": working, "deadlineHour": BENTO_DEADLINE_HOUR,
        "deadlinePassed": deadline_passed, "canOrder": (reason == "ok"), "reason": reason,
        "isApprover": (shain_str == MORIMAKOTO_SHAIN),
        "menu": [{"no": k, "price": BENTO_MENU[k]} for k in BENTO_ORDER_LABEL],
        "todayOrder": ({"menuNo": to.get("MenuNo"), "price": to.get("Price"),
                        "status": to.get("OrderStatus")} if to else None),
    })


@app.route(route="bento/order", methods=["POST", "OPTIONS"])
def bento_order(req: func.HttpRequest) -> func.HttpResponse:
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    shain_str = str(payload["shainNo"])
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    menu_no = str(body.get("menuNo") or "").strip()
    if menu_no not in BENTO_MENU:
        return _json_response({"error": "invalid_menu"}, 400)
    wait = check_rate_limit(payload["shainNo"], "bento")
    if wait is not None:
        return _json_response({"error": "rate_limited", "retryAfterSeconds": wait}, 429)
    emp = find_active_employee_by_shain(payload["shainNo"])
    if not emp:
        return _json_response({"error": "not_active"}, 403)
    if not _bento_eligible(emp):
        return _json_response({"error": "not_eligible"}, 403)
    today = _today_jst()
    if not _bento_test_open():             # テストモード時はゲート無効
        if not _bento_is_working_day(today):
            return _json_response({"error": "holiday"}, 400)
        if _now_jst().hour >= BENTO_DEADLINE_HOUR:
            return _json_response({"error": "deadline_passed"}, 400)
    if _bento_today_order(shain_str, today):
        return _json_response({"error": "already_ordered"}, 409)
    name = emp.get(F_SHAIN_NAME) or ""
    price = BENTO_MENU[menu_no]
    new_id = sp_post_item(LIST_BENTO, {
        "Title": f"{today.isoformat()} {shain_str}",
        "OrderDate": today.strftime("%Y/%m/%d"),
        "ShainNo": shain_str, "EmpName": name,
        "Hakensaki": strip_buka_prefix(emp.get(F_BUKA) or ""),
        "MenuNo": menu_no, "Price": str(price), "OrderStatus": "pending",
    })
    # 承認者(yukyu-app: 役員/森まこと/森ゆうじ)へ ロック画面 Push
    try:
        _push_to_admins(BENTO_APPROVER_EMAILS, {
            "title": "🍱 弁当注文 承認待ち",
            "body": f"{name or shain_str} / {menu_no} {price}円",
            "url": YUKYU_APP_URL, "tag": "bento-approve", "badge": 1,
        })
    except Exception:
        logging.exception("bento push failed")
    return _json_response({"ok": True, "id": new_id, "menuNo": menu_no, "price": price})


@app.route(route="bento/history", methods=["GET", "OPTIONS"])
def bento_history(req: func.HttpRequest) -> func.HttpResponse:
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    shain_str = str(payload["shainNo"])
    try:
        items = sp_get_items(
            LIST_BENTO,
            select="Id,OrderDate,MenuNo,Price,OrderStatus,ApprovedAt,RejectReason,Created",
            filter_=f"ShainNo eq '{shain_str}'", orderby="Id desc", top=60)
        out = []
        for it in items:
            od = _utc_to_jst_date(it.get("OrderDate"))
            ap = _utc_to_jst_date(it.get("ApprovedAt"))
            cr = _utc_to_jst_date(it.get("Created"))
            out.append({"id": it.get("Id"), "date": od.isoformat() if od else None,
                        "menuNo": it.get("MenuNo"), "price": it.get("Price"),
                        "status": it.get("OrderStatus"),
                        "approvedAt": ap.isoformat() if ap else None,
                        "rejectReason": it.get("RejectReason"),
                        "createdAt": cr.isoformat() if cr else None})
        return _json_response({"items": out})
    except Exception as e:
        logging.exception("bento_history failed")
        return _json_response({"error": "internal", "detail": str(e)}, 500)


@app.route(route="bento/pending", methods=["GET", "OPTIONS"])
def bento_pending(req: func.HttpRequest) -> func.HttpResponse:
    """承認担当(モリマコト)のみ: 当日の承認待ち/承認済み一覧。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    if str(payload["shainNo"]) != MORIMAKOTO_SHAIN:
        return _json_response({"error": "forbidden"}, 403)
    today = _today_jst()
    try:
        items = sp_get_items(
            LIST_BENTO,
            select="Id,OrderDate,ShainNo,EmpName,Hakensaki,MenuNo,Price,OrderStatus",
            filter_=f"OrderDate eq datetime'{today.isoformat()}T00:00:00'",
            orderby="Id asc", top=500)
        pend, appr = [], []
        for it in items:
            row = {"id": it.get("Id"), "shainNo": it.get("ShainNo"), "name": it.get("EmpName"),
                   "hakensaki": it.get("Hakensaki"), "menuNo": it.get("MenuNo"),
                   "price": it.get("Price"), "status": it.get("OrderStatus")}
            if it.get("OrderStatus") == "pending":
                pend.append(row)
            elif it.get("OrderStatus") == "approved":
                appr.append(row)
        return _json_response({"ok": True, "date": today.isoformat(),
                               "pending": pend, "approved": appr, "pendingCount": len(pend)})
    except Exception as e:
        logging.exception("bento_pending failed")
        return _json_response({"error": "internal", "detail": str(e)}, 500)


@app.route(route="bento/approve", methods=["POST", "OPTIONS"])
def bento_approve(req: func.HttpRequest) -> func.HttpResponse:
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    if str(payload["shainNo"]) != MORIMAKOTO_SHAIN:
        return _json_response({"error": "forbidden"}, 403)
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    oid = body.get("id")
    action = str(body.get("action") or "approve")
    reason = (body.get("reason") or "")[:300]
    if not oid:
        return _json_response({"error": "missing_id"}, 400)
    approved_at = _dt.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")   # MERGEはISO(UTC)
    try:
        if action == "reject":
            sp_patch_item(LIST_BENTO, int(oid),
                          {"OrderStatus": "rejected", "RejectReason": reason,
                           "ApprovedBy": MORIMAKOTO_SHAIN, "ApprovedAt": approved_at})
        else:
            sp_patch_item(LIST_BENTO, int(oid),
                          {"OrderStatus": "approved", "ApprovedBy": MORIMAKOTO_SHAIN,
                           "ApprovedAt": approved_at})
        return _json_response({"ok": True, "id": oid, "action": action})
    except Exception as e:
        logging.exception("bento_approve failed")
        return _json_response({"error": "internal", "detail": str(e)}, 500)


# --- 弁当: 朝の集計メール (モリマコトのアドレスから自動送信) ---
BENTO_WD = ['月', '火', '水', '木', '金', '土', '日']


def _bento_compose_mail(d: _dt.date, orders: List[Dict[str, Any]]):
    cnt = {k: 0 for k in BENTO_ORDER_LABEL}
    total_yen = 0
    total_n = 0
    for o in orders:
        mn = str(o.get("MenuNo") or "")
        try:
            p = int(float(o.get("Price") or BENTO_MENU.get(mn, 0)))
        except Exception:
            p = BENTO_MENU.get(mn, 0)
        if mn in cnt:
            cnt[mn] += 1
        total_yen += p
        total_n += 1
    subj = f"本日 {d.month}/{d.day}（{BENTO_WD[d.weekday()]}）のお弁当注文です"
    if total_n == 0:
        lines = ["本日のお弁当注文はありませんでした。"]
    else:
        lines = [f"{k} {BENTO_MENU[k]:,}円 × {cnt[k]}食" for k in BENTO_ORDER_LABEL if cnt[k]]
        lines.append("─────────────")
        lines.append(f"合計 {total_n}食 ／ {total_yen:,}円")
    body = ("お世話になります。\nステップアップの森です。\n\n"
            "下記の通り、注文をお願いします。\n\n" + "\n".join(lines) +
            "\n\n以上よろしくお願いします。")
    return subj, body


def _bento_send_summary(dry: bool = False) -> Dict[str, Any]:
    """当日の承認済みを集計し、モリマコト(BENTO_MAIL_FROM)から BENTO_MAIL_TO へ送信。"""
    d = _today_jst()
    flt = f"OrderDate eq datetime'{d.isoformat()}T00:00:00' and OrderStatus eq 'approved'"
    orders = sp_get_items(LIST_BENTO, select="Id,MenuNo,Price,EmpName", filter_=flt, top=500)
    subj, body = _bento_compose_mail(d, orders)
    sender = os.environ.get("BENTO_MAIL_FROM", "m.mori@team-stepup.com").strip()
    to_addr = os.environ.get("BENTO_MAIL_TO", "h.yamashita@team-stepup.com").strip()
    cc_addr = os.environ.get("BENTO_MAIL_CC", "").strip()
    as_draft = os.environ.get("BENTO_MAIL_DRAFT", "1").strip() != "0"   # 既定=下書き保存(送信しない)
    res = {"date": d.isoformat(), "count": len(orders), "subject": subj, "body": body,
           "from": sender, "to": to_addr, "cc": cc_addr,
           "mode": ("draft" if as_draft else "send"), "sent": False, "draftSaved": False, "draftId": None}
    if dry:
        return res
    token = _get_graph_token()
    msg = {"subject": subj, "body": {"contentType": "Text", "content": body},
           "toRecipients": [{"emailAddress": {"address": to_addr}}]}
    if cc_addr:
        msg["ccRecipients"] = [{"emailAddress": {"address": a.strip()}} for a in cc_addr.split(",") if a.strip()]
    if as_draft:
        # 森まこと(sender)のメールボックスに【下書き】として保存。送信はしない(本人がOutlookで確認→送信)。Mail.ReadWrite。
        r = requests.post(f"https://graph.microsoft.com/v1.0/users/{sender}/messages",
                          headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
                          json=msg, timeout=30)
        if not r.ok:
            res["error"] = f"{r.status_code} {r.text[:300]}"
            return res
        res["draftSaved"] = True
        res["draftId"] = r.json().get("id")
        return res
    # 送信モード (Mail.Send)
    r = requests.post(f"https://graph.microsoft.com/v1.0/users/{sender}/sendMail",
                      headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
                      json={"message": msg, "saveToSentItems": True}, timeout=30)
    if not r.ok:
        res["error"] = f"{r.status_code} {r.text[:300]}"
        return res
    res["sent"] = True
    return res


@app.timer_trigger(arg_name="timer", schedule="0 5 0 * * 1-5", run_on_startup=False, use_monitor=True)
def bento_morning_mail_timer(timer: func.TimerRequest) -> None:
    """毎朝9:05 JST(=00:05 UTC) 平日。稼働日のみ当日の弁当注文集計をメール送信。"""
    try:
        d = _today_jst()
        if not _bento_is_working_day(d):
            logging.info("弁当朝メール: %s は稼働日でないためスキップ", d)
            return
        res = _bento_send_summary(dry=False)
        logging.info("弁当朝メール: mode=%s draftSaved=%s sent=%s count=%d to=%s err=%s",
                     res.get("mode"), res.get("draftSaved"), res.get("sent"), res.get("count"), res.get("to"), res.get("error"))
    except Exception:
        logging.exception("弁当朝メール(timer)に失敗")


@app.route(route="bento/send-summary", methods=["POST", "OPTIONS"])
def bento_send_summary(req: func.HttpRequest) -> func.HttpResponse:
    """テスト/手動再送用: 当日の集計メールを送信(またはdry=trueでプレビュー)。総務/役員(staff)のみ。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    email, err = require_staff_auth(req)
    if err:
        return err
    try:
        body = req.get_json()
    except Exception:
        body = {}
    dry = bool(body.get("dry"))
    try:
        res = _bento_send_summary(dry=dry)
        res["ok"] = True
        return _json_response(res)
    except Exception as e:
        logging.exception("bento_send_summary failed")
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


def run_zairyu_ocr_gpt(front_bytes: bytes, back_bytes: Optional[bytes] = None) -> Dict[str, Any]:
    """在留カード OCR (Azure OpenAI GPT-4o-mini マルチモーダル抽出)。
    画像を直接 LLM に渡し、構造化 JSON を得る。和暦変換・レイアウト差・ラベル対応に強い。"""
    import gpt_ocr
    return gpt_ocr.extract_zairyu_fields(front_bytes, back_bytes)


# GPT(gpt_ocr) の出力キー → 既存 DI パーサー互換キー の対応
_GPT_KEY_MAP = {
    "license": {"licenseNo": "licenseNumber", "licenseType": "licenseType", "licenseGetDate": "licenseDate",
                "licenseExpiry": "licenseExpiry", "birthday": "birthday", "name": "name"},
    "shaken": {"carNumber": "carNumber", "carMaker": "carMaker", "carName": "carName",
               "carDisplacement": "haikiryo", "carFirstReg": "shonendo", "shakenExpiry": "shakenKigen"},
    "jibaiseki": {"jibaisekiCompany": "jibaiKaisha", "jibaisekiNo": "jibaiShoken", "jibaisekiExpiry": "jibaiKigen"},
    "nini": {"niniCompany": "niniKaisha", "niniNo": "niniShoken", "niniStartDate": "niniKaishi",
             "niniExpiry": "niniKigen", "taijin": "taijin", "taibutsu": "taibutsu"},
}


def _gpt_doc_to_di(doc_type: str, gpt_result: Dict[str, Any]) -> Dict[str, Any]:
    """GPT 抽出結果を DI パーサー互換キーに変換 (非空のみ)。"""
    out: Dict[str, Any] = {}
    for gk, dk in _GPT_KEY_MAP.get(doc_type, {}).items():
        v = gpt_result.get(gk)
        if v not in (None, ""):
            out[dk] = v
    return out


def _gpt_ocr_available() -> bool:
    return bool(os.environ.get("AZURE_OPENAI_ENDPOINT", "").strip())


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
    # 裏面 (任意) — GPT エンジンは表裏両方を参照できる
    back_bytes = None
    back_data = body.get("backImage")
    if back_data:
        try:
            back_bytes = base64.b64decode(_strip_data_url(back_data))
        except Exception:
            back_bytes = None

    # エンジン選択: 既定は "gpt" (Azure OpenAI GPT-4o-mini)。失敗時は DI(prebuilt-read+正規表現) に自動フォールバック
    engine = (body.get("engine") or req.params.get("engine") or "gpt").strip().lower()
    try:
        if engine == "gpt":
            try:
                ocr_result = run_zairyu_ocr_gpt(front_bytes, back_bytes)
            except Exception as ge:
                logging.warning(f"GPT OCR failed, fallback to DI: {ge}")
                ocr_result = run_zairyu_ocr(front_bytes)
                ocr_result["engine"] = "di-fallback"
        else:
            ocr_result = run_zairyu_ocr(front_bytes)
    except Exception as e:
        logging.exception("OCR failed")
        return _json_response({"error": "ocr_failed", "engine": engine, "detail": str(e)}, 500)

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


# ====== 給油申請 (送迎運転手のみ・レシート/メーターOCR自動入力・List54へ書込) ======
# ポータルが SP List「ガソリン代払い戻し」(List54) への INSERT に加え、
# 給油間距離/燃費/間隔(空白期間)の自動計算と Teams「送迎関係」チャネルへの通知も自前で行う。
# (旧 Forms 用 Power Automate フローは Forms 応答トリガーでポータル投入には発火しないため。
#  F1/Forms 廃止後も止まらないようバックエンドに集約。)精算は浩俊が随時。
LIST_GASORIN = "ac84ec0e-6853-463a-9469-d94b52940485"   # SP「ガソリン代払い戻し」(List54)
KYUYU_DRIVERS = {50466, 50565, 50692, 50709, 50736, 50743, 50750, 60057, 111113, 111356}
KYUYU_FOLDER_BASE = "/sites/TeamStepup/Shared Documents/給油申請レシート"
# Teams「送迎関係」チャネルのメールアドレス(チャネルへメール投稿=Forms フロー廃止後も通知継続)。
# 既定値はメモ由来。実際のチャネルメール(Teams→チャネル→メールアドレスを取得)で上書き可。
KYUYU_TEAMS_EMAIL = os.environ.get("KYUYU_TEAMS_EMAIL", "80ade604.team-stepup.com@jp.teams.ms")
# 給油申請のロック画面Push宛(浩俊)。自分メモ(48:notes)は自己発言扱いで通知が出ないため、
# yukyu-app の既存Push基盤(_push_to_admins)で浩俊端末へ別途お知らせを出す。
KYUYU_PUSH_EMAILS = [e.strip() for e in os.environ.get(
    "KYUYU_PUSH_EMAILS", "h.yamashita@team-stepup.com").split(",") if e.strip()]
# List54 業務フィールド (AddValidateUpdateItemUsingPath は InternalName を要求=OData_ プレフィックス無し)
GAS_F_VEHICLE = "_x30ca__x30f3__x30d0__x30fc__x30"   # ナンバー、車両
GAS_F_AMOUNT  = "_x6255__x623b__x91d1__x984d_"        # 払戻金額
GAS_F_PAYTYPE = "_x652f__x6255__x3044__x306e__x7a"    # 支払いの種類
GAS_F_RECEIPT = "_x9818__x53ce__x66f8_NEW"            # 領収書NEW (1枚目)
GAS_F_RECEIPT2 = "_x9818__x53ce__x66f8_2"             # 領収書2 (2枚目)
GAS_F_RECEIPT3 = "_x9818__x53ce__x66f8_3"             # 領収書3 (3枚目)
GAS_F_ODOIMG  = "_x8d70__x884c__x8ddd__x96e2__x75"    # 走行距離画像
GAS_F_ODO     = "_x73fe__x5728__x8d70__x884c__x8d"    # 現在走行距離
GAS_F_LITERS  = "_x7d66__x6cb9__x91cf__x2113_"        # 給油量ℓ
GAS_F_PURPOSE = "_x5229__x7528__x76ee__x7684_"        # 利用目的
GAS_F_DIST    = "_x7d66__x6cb9__x9593__x8ddd__x96"    # 給油間距離 (Note・前回給油からの距離)
GAS_F_NENPI   = "_x71c3__x8cbb_"                       # 燃費 (Note・給油間距離/給油量)
GAS_F_KANKAKU = "_x7a7a__x767d__x671f__x9593_"        # 間隔=空白期間 (Note・前回給油からの経過)
GAS_VEHICLES = ["997　エスクァイア HV", "1967　ハイエース", "3011  ノア", "8504　ハイエース",
                "1724  ノア", "363    ハイエース", "5344  ノア　HV", "5179　ノア HV",
                "4222　ハイエース", "6441  プリウスα　HV", "6803　ノア", "2862　ノア",
                "26　　ヴォクシー", "514     ヴォクシー", "7638   ハイエース", "代車　Carro reserva"]
GAS_PURPOSES = ["ガソリン給油", "オイル交換", "洗車", "その他"]
GAS_PAYTYPES = ["会社カード", "PayPay", "その他"]   # 既定=会社カード(先頭)

_kyuyu_veh_cache = {"vehicles": None, "ts": 0.0}


def _kyuyu_vehicles() -> List[str]:
    """List54「ナンバー、車両」の既存値から車両一覧を動的生成(使用回数多い順)。1時間キャッシュ。失敗時はGAS_VEHICLES。"""
    now = time.time()
    if _kyuyu_veh_cache["vehicles"] and (now - _kyuyu_veh_cache["ts"] < 3600):
        return _kyuyu_veh_cache["vehicles"]
    try:
        items = sp_get_items(LIST_GASORIN, top=5000)
        cnt: Dict[str, int] = {}
        for it in items:
            v = str(it.get("OData_" + GAS_F_VEHICLE) or it.get(GAS_F_VEHICLE) or "").strip()
            if v:
                cnt[v] = cnt.get(v, 0) + 1
        if cnt:
            vs = [v for v, _ in sorted(cnt.items(), key=lambda x: -x[1])]
            _kyuyu_veh_cache["vehicles"] = vs
            _kyuyu_veh_cache["ts"] = now
            return vs
    except Exception:
        logging.exception("kyuyu vehicles fetch failed")
    return GAS_VEHICLES


def _kyuyu_default_vehicle(shain_no: int, vehicles: List[str]) -> str:
    """車両管理(List=/sites/PowerApps/Lists/List)の DriverShainNo==社員番号 の担当車両を探し、
    vehicles(List54の選択肢)から ナンバー一致 する1件を返す。無ければ ''(=選択してください)。"""
    try:
        h = {"Authorization": f"Bearer {_get_sp_token()}", "Accept": "application/json;odata=nometadata"}
        url = f"{SITE_URL}/_api/web/GetList('/sites/PowerApps/Lists/List')/items?$top=500"
        items = requests.get(url, headers=h, timeout=30).json().get("value", [])
        for it in items:
            ds = it.get("DriverShainNo")
            try:
                if ds is None or int(float(ds)) != int(shain_no):
                    continue
            except Exception:
                continue
            num = str(it.get("OData__x004e_O2") or it.get("_x004e_O2") or "").strip()   # ナンバー(例 1724)
            digits = re.sub(r"\D", "", num)
            if not digits:
                continue
            # ①完全一致(車両先頭の数字==ナンバー)を最優先。短い番号(例26)が別車
            #   (2862ノア)に startswith 誤一致するのを防ぐため、先に全車を完全一致で走査。
            for v in vehicles:
                vfirst = (v.split() or [""])[0]
                if re.sub(r"\D", "", vfirst) == digits:
                    return v
            # ②完全一致が無いときだけ前方一致でフォールバック
            for v in vehicles:
                if v.replace("　", "").replace(" ", "").startswith(digits):
                    return v
        return ""
    except Exception:
        logging.exception("kyuyu default vehicle lookup failed")
        return ""


def _kyuyu_eligible(emp: Dict[str, Any]) -> bool:
    """給油申請の対象=送迎運転手(社員番号がドライバーセットに含まれる)。"""
    try:
        return int(float(emp.get(F_SHAIN_NO))) in KYUYU_DRIVERS
    except Exception:
        return False


def _kyuyu_num(v) -> str:
    """金額/給油量/走行距離を数字文字列に正規化(List54はNote型なので文字列で保存)。"""
    try:
        s = re.sub(r"[^\d.]", "", str(v if v is not None else ""))
        return s.strip(".") if s else ""
    except Exception:
        return ""


def _kyuyu_plate_key(vehicle: str) -> str:
    """車両文字列から同一車両判定キー(先頭のナンバー数字)。数字が無ければ空白除去した全体。"""
    m = re.match(r"\s*(\d+)", vehicle or "")
    if m:
        return m.group(1)
    return re.sub(r"[\s　]", "", vehicle or "")


def _kyuyu_parse_created(s: str):
    """SP Created(ISO8601 UTC)を aware datetime に。失敗時 None。"""
    if not s:
        return None
    try:
        t = s.strip().replace("Z", "")
        if "." in t:
            t = t.split(".")[0]
        return _dt.datetime.strptime(t, "%Y-%m-%dT%H:%M:%S").replace(tzinfo=_dt.timezone.utc)
    except Exception:
        return None


def _kyuyu_fmt_kankaku(delta: "_dt.timedelta") -> str:
    """timedelta → 既存フローと同じ「X日Y時間Z分間」。"""
    total = max(0, int(delta.total_seconds()))
    days, rem = divmod(total, 86400)
    hours, rem = divmod(rem, 3600)
    mins = rem // 60
    return f"{days}日{hours}時間{mins}分間"


def _kyuyu_prev_same_vehicle(vehicle: str):
    """同一車両(ナンバー一致)の直近の過去エントリから (前回走行距離float|None, Created datetime|None)。"""
    key = _kyuyu_plate_key(vehicle)
    try:
        items = sp_get_items(LIST_GASORIN, top=5000)
    except Exception:
        logging.exception("kyuyu prev fetch failed")
        return (None, None)
    best = None  # (created_dt, odo_float_or_None)
    for it in items:
        v = str(it.get("OData_" + GAS_F_VEHICLE) or it.get(GAS_F_VEHICLE) or "").strip()
        if not v or _kyuyu_plate_key(v) != key:
            continue
        cdt = _kyuyu_parse_created(it.get("Created"))
        if not cdt:
            continue
        odo_raw = it.get("OData_" + GAS_F_ODO) or it.get(GAS_F_ODO)
        odo = None
        try:
            num = re.sub(r"[^\d.]", "", str(odo_raw if odo_raw is not None else "")).strip(".")
            odo = float(num) if num else None
        except Exception:
            odo = None
        if best is None or cdt > best[0]:
            best = (cdt, odo)
    if best is None:
        return (None, None)
    return (best[1], best[0])


def _kyuyu_calc(vehicle: str, odometer: str, liters: str):
    """同一車両の前回給油から 給油間距離/燃費/間隔(空白期間) を計算。各々文字列(無ければ '')。"""
    dist_s = nenpi_s = kankaku_s = ""
    try:
        prev_odo, prev_created = _kyuyu_prev_same_vehicle(vehicle)
        cur_odo = None
        if odometer:
            n = re.sub(r"[^\d.]", "", odometer).strip(".")
            cur_odo = float(n) if n else None
        if prev_created is not None:
            kankaku_s = _kyuyu_fmt_kankaku(_dt.datetime.now(_dt.timezone.utc) - prev_created)
        # 代車(Carro reserva=ナンバー数字なし)は共用で走行距離が別車のもの→距離/燃費は出さない
        plate_is_num = _kyuyu_plate_key(vehicle).isdigit()
        if plate_is_num and cur_odo is not None and prev_odo is not None and cur_odo > prev_odo:
            dist = cur_odo - prev_odo
            if dist < 100000:   # 桁ミス/車両入替の異常値を除外
                dist_s = str(int(round(dist)))
                n = re.sub(r"[^\d.]", "", liters).strip(".") if liters else ""
                lit = float(n) if n else 0.0
                if lit > 0:
                    nenpi_s = str(round(dist / lit, 1))
    except Exception:
        logging.exception("kyuyu calc failed")
    return dist_s, nenpi_s, kankaku_s


def _sp_download_bytes(abs_url: str, timeout: int = 60):
    """SP の絶対URLからファイル bytes を取得(MIトークン)。失敗時 None。
    REST はファイルが属するサイトWeb(/sites/X)配下で呼ぶ必要がある。"""
    try:
        from urllib.parse import urlparse, unquote, quote
        p = urlparse(abs_url)
        if not p.path:
            return None
        sru = unquote(p.path)   # /sites/TeamStepup/Shared Documents/.../x.jpg
        m = re.match(r"(/(?:sites|teams)/[^/]+)", sru)
        web = m.group(1) if m else ""   # /sites/TeamStepup
        api = f"{p.scheme}://{p.netloc}{web}/_api/web/GetFileByServerRelativeUrl('{quote(sru)}')/$value"
        r = requests.get(api, headers={"Authorization": f"Bearer {_get_sp_token()}"}, timeout=timeout)
        r.raise_for_status()
        return r.content
    except Exception:
        logging.exception("sp download failed: %s", abs_url)
        return None


def _resize_image_jpeg(raw: bytes, max_dim: int = 1500, quality: int = 78):
    """画像bytesを縮小JPEGに(インライン埋め込み用にメール軽量化)。失敗時は十分小さければ原本、大きすぎれば None。"""
    try:
        from PIL import Image, ImageOps
        import io as _io
        im = Image.open(_io.BytesIO(raw))
        im = ImageOps.exif_transpose(im)   # スマホ写真の向きを補正
        im = im.convert("RGB")
        im.thumbnail((max_dim, max_dim))
        buf = _io.BytesIO()
        im.save(buf, format="JPEG", quality=quality, optimize=True)
        return buf.getvalue()
    except Exception:
        logging.exception("image resize failed")
        return raw if raw and len(raw) <= 1_200_000 else None


def _kyuyu_build_inline_images(receipt_urls, odo_url, receipt_bytes=None, odo_bytes_raw=None,
                               total_budget: int = 3_200_000):
    """レシート/メーター画像を縮小して Graph 添付(非インライン)を作る。
    Teams のチャネル投稿は inline CID 画像を「本文」と「添付一覧」の両方に出して重複させるため、
    添付のみ(本文には埋め込まない)にして各画像を1回だけ表示させる。
    receipt_bytes/odo_bytes_raw が在れば再DLせず使用(applyの即時パス用)。戻り (attachments, '')。"""
    sources = []   # (url, raw_or_None, filename)
    for i, u in enumerate(receipt_urls or []):
        raw = receipt_bytes[i] if (receipt_bytes and i < len(receipt_bytes)) else None
        sources.append((u, raw, f"領収書{i + 1}.jpg"))
    if odo_url:
        sources.append((odo_url, odo_bytes_raw, "走行メーター.jpg"))
    attachments, used = [], 0
    for url, raw, fname in sources:
        if raw is None:
            raw = _sp_download_bytes(url)
        small = _resize_image_jpeg(raw) if raw else None
        if not small or (used + len(small) > total_budget):
            continue   # 取得失敗/予算超は付けず(リンクは別途残る)
        used += len(small)
        attachments.append({
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": fname, "contentType": "image/jpeg",
            "contentBytes": base64.b64encode(small).decode("ascii"),
        })
    return attachments, ""


def _kyuyu_notify_teams(shain_no, name, vehicle, day, amount, liters, odometer,
                        dist_s, nenpi_s, kankaku_s, pay_type, purposes,
                        receipt_urls, odo_url, receipt_bytes=None, odo_bytes_raw=None) -> None:
    """Teams「送迎関係」チャネルへ給油申請を通知(チャネルメール宛 sendMail・HTML・best-effort)。
    旧 Forms フロー「⛽給油のお知らせ」カードの体裁に寄せ、URLは名前付きリンクに。
    投稿者は送信元(h.yamashita)になるため『誰が』は本文(申請者)に明記。"""
    def _h(s):
        return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
    veh_short = re.sub(r"[\s　]+", " ", vehicle or "").strip()
    try:
        amt_disp = f"¥{int(float(amount)):,}" if amount else "（未入力）"
    except Exception:
        amt_disp = f"¥{amount}" if amount else "（未入力）"
    # 距離 = 給油間距離（現在走行距離）。代車/メーター逆行で距離が無い時は走行距離のみ明示。
    if dist_s:
        dist_disp = f"{dist_s} km" + (f"（{_h(odometer)} km）" if odometer else "")
    elif odometer:
        dist_disp = f"—（走行 {_h(odometer)} km）"
    else:
        dist_disp = "―"
    rows = [
        ("申請者", _h(name)),   # 社員番号・日付はチャットに出るため省く(メモと統一)
        ("車両", _h(veh_short)),
        ("区分", _h("・".join(purposes))),
        ("前給油", f"{_h(kankaku_s)}前" if kankaku_s else "―"),
        ("距離", dist_disp),
        ("燃費", f"{nenpi_s} km/ℓ" if nenpi_s else "―"),
        ("給油量", f"{_h(liters) if liters else '―'} ℓ"),
        ("払戻金額", amt_disp),
        ("支払い", _h(pay_type) if pay_type else "（未選択）"),
    ]
    tr = "".join(
        f'<tr><td style="padding:1px 14px 1px 0;color:#666;white-space:nowrap">{k}</td>'
        f'<td style="padding:1px 0;font-weight:600">{v}</td></tr>'
        for k, v in rows
    )
    links = []
    if odo_url:
        links.append(f'<a href="{_h(odo_url)}">🚗 走行距離メーターを開く</a>')
    for i, u in enumerate(receipt_urls):
        links.append(f'<a href="{_h(u)}">🧾 領収書{i + 1}を開く</a>')
    links.append(f'<a href="{SITE_URL}/Lists/List54/AllItems.aspx">📊 リストで全体を確認する</a>')
    # レシート/メーター画像を縮小してインライン埋め込み(リンク自動展開は不安定なため確実に表示)
    attachments, img_html = _kyuyu_build_inline_images(receipt_urls, odo_url, receipt_bytes, odo_bytes_raw)
    html = (
        '<div style="font-family:Segoe UI,Meiryo,sans-serif;font-size:14px;color:#222">'
        '<div style="font-size:16px;font-weight:700;margin-bottom:8px">⛽ 給油のお知らせ</div>'
        f'<table style="border-collapse:collapse">{tr}</table>'
        f'<div style="margin-top:12px;line-height:2">{"<br>".join(links)}</div>'
        f'{img_html}</div>'
    )
    subject = f"⛽ 給油のお知らせ {name}（{veh_short}）{amt_disp}"
    _send_notification_mail(subject, html, to_addr=KYUYU_TEAMS_EMAIL, html=True, attachments=attachments)


@app.route(route="kyuyu/status", methods=["GET", "OPTIONS"])
def kyuyu_status(req: func.HttpRequest) -> func.HttpResponse:
    """給油画面の初期データ: 対象可否・車両/利用目的/支払いの選択肢。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    emp = find_active_employee_by_shain(payload["shainNo"])
    if not emp:
        return _json_response({"error": "not_active"}, 403)
    vehicles = _kyuyu_vehicles()
    return _json_response({
        "ok": True, "eligible": _kyuyu_eligible(emp),
        "vehicles": vehicles,
        "defaultVehicle": _kyuyu_default_vehicle(int(payload["shainNo"]), vehicles),
        "purposes": GAS_PURPOSES, "payTypes": GAS_PAYTYPES,
    })


@app.route(route="kyuyu/ocr", methods=["POST", "OPTIONS"])
def kyuyu_ocr(req: func.HttpRequest) -> func.HttpResponse:
    """給油レシート or メーター写真を OCR。body {image:<dataURL>, kind:'receipt'|'odometer'}"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    shain_no = int(payload["shainNo"])
    emp = find_active_employee_by_shain(shain_no)
    if not emp or not _kyuyu_eligible(emp):
        return _json_response({"error": "not_eligible"}, 403)
    wait = check_rate_limit(shain_no, "kyuyu")
    if wait is not None:
        return _json_response({"error": "rate_limited", "retryAfterSeconds": wait}, 429,
                              extra_headers={"Retry-After": str(wait)})
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    kind = (body.get("kind") or "receipt").strip().lower()
    doc_type = "odometer" if kind == "odometer" else "kyuyu"
    img = body.get("image")
    if not img:
        return _json_response({"error": "missing_image"}, 400)
    try:
        img_bytes = base64.b64decode(_strip_data_url(img))
    except Exception as e:
        return _json_response({"error": "invalid_base64", "detail": str(e)}, 400)
    if len(img_bytes) > 10 * 1024 * 1024:
        return _json_response({"error": "image_too_large", "maxMB": 10}, 400)
    try:
        import gpt_ocr
        ocr = gpt_ocr.extract_doc_fields(doc_type, [img_bytes])
    except Exception as e:
        logging.exception("kyuyu OCR failed")
        return _json_response({"error": "ocr_failed", "detail": str(e)}, 500)
    return _json_response({"ok": True, "kind": kind, "ocr": ocr})


@app.route(route="kyuyu/apply", methods=["POST", "OPTIONS"])
def kyuyu_apply(req: func.HttpRequest) -> func.HttpResponse:
    """給油申請を List54 に登録。写真を SP に保存しURLを格納。
    給油間距離/燃費/間隔を計算して書込み、Teams「送迎関係」へ通知(いずれも自前)。
    Body: {vehicle, amount, liters, odometer, payType, purposes:[...],
           receiptImages:[<dataURL>...](最大3) | receiptImage:<dataURL>, odometerImage:<dataURL>(任意)}"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    shain_no = int(payload["shainNo"])
    emp = find_active_employee_by_shain(shain_no)
    if not emp or not _kyuyu_eligible(emp):
        return _json_response({"error": "not_eligible"}, 403)
    wait = check_rate_limit(shain_no, "kyuyu")
    if wait is not None:
        return _json_response({"error": "rate_limited", "retryAfterSeconds": wait}, 429,
                              extra_headers={"Retry-After": str(wait)})
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    vehicle = (body.get("vehicle") or "").strip()
    valid_vehicles = set(_kyuyu_vehicles()) | set(GAS_VEHICLES)
    if vehicle not in valid_vehicles:
        return _json_response({"error": "invalid_vehicle"}, 400)
    # レシート: 最大3枚(receiptImages配列 / 後方互換で receiptImage 単体も可)
    receipt_datas = body.get("receiptImages")
    if not receipt_datas:
        single = body.get("receiptImage")
        receipt_datas = [single] if single else []
    receipt_datas = [r for r in receipt_datas if r][:3]
    if not receipt_datas:
        return _json_response({"error": "missing_receipt"}, 400)
    pay_type = (body.get("payType") or "").strip()
    if pay_type and pay_type not in GAS_PAYTYPES:
        pay_type = "その他"
    purposes = [p for p in (body.get("purposes") or []) if p in GAS_PURPOSES] or ["ガソリン給油"]
    amount = _kyuyu_num(body.get("amount"))
    liters = _kyuyu_num(body.get("liters"))
    odometer = _kyuyu_num(body.get("odometer"))
    odo_data = body.get("odometerImage")
    name = emp.get(F_SHAIN_NAME) or ""
    today = _today_jst()
    # 画像デコード(レシート最大3枚 + メーター)
    try:
        receipt_bytes_list = [base64.b64decode(_strip_data_url(r)) for r in receipt_datas]
        odo_bytes = base64.b64decode(_strip_data_url(odo_data)) if odo_data else None
    except Exception as e:
        return _json_response({"error": "invalid_base64", "detail": str(e)}, 400)
    if any(len(b) > 10 * 1024 * 1024 for b in receipt_bytes_list) or (odo_bytes and len(odo_bytes) > 10 * 1024 * 1024):
        return _json_response({"error": "image_too_large", "maxMB": 10}, 400)
    # 写真を SP に保存
    try:
        year_folder = f"{KYUYU_FOLDER_BASE}/{today.year}"
        sp_create_folder_if_not_exists(KYUYU_FOLDER_BASE)
        sp_create_folder_if_not_exists(year_folder)
        ts = _now_jst().strftime("%Y%m%d_%H%M%S")
        url_base = f"{SITE_TEAMSTEPUP}/Shared Documents/給油申請レシート/{today.year}"
        receipt_urls = []
        for i, rb in enumerate(receipt_bytes_list):
            rname = _safe_filename(f"{shain_no}_{today.isoformat()}_領収書{i+1}_{ts}.jpg")
            sp_upload_file(year_folder, rname, rb)
            receipt_urls.append(f"{url_base}/{rname}")
        odo_url = ""
        if odo_bytes:
            odo_name = _safe_filename(f"{shain_no}_{today.isoformat()}_メーター_{ts}.jpg")
            sp_upload_file(year_folder, odo_name, odo_bytes)
            odo_url = f"{url_base}/{odo_name}"
    except Exception as e:
        logging.exception("kyuyu upload failed")
        return _json_response({"error": "upload_failed", "detail": str(e)}, 500)
    # List54 へ INSERT (既存フローが給油間距離/燃費/間隔の計算 + Teams通知)
    # 申請者(User型)は運転手のM365アカウントが無い前提のため設定せず、Title で識別する。
    fields: Dict[str, Any] = {
        "Title": f"{shain_no} {name} ({today.isoformat()})",
        GAS_F_VEHICLE: vehicle,
        GAS_F_PURPOSE: ",".join(purposes),
    }
    for i, ru in enumerate(receipt_urls):   # レシートURLを 領収書NEW/2/3 に割当
        fields[[GAS_F_RECEIPT, GAS_F_RECEIPT2, GAS_F_RECEIPT3][i]] = ru
    if amount:
        fields[GAS_F_AMOUNT] = amount
    if liters:
        fields[GAS_F_LITERS] = liters
    if odometer:
        fields[GAS_F_ODO] = odometer
    if odo_url:
        fields[GAS_F_ODOIMG] = odo_url
    if pay_type:
        fields[GAS_F_PAYTYPE] = pay_type
    # 同一車両の前回給油から 給油間距離/燃費/間隔 を計算(旧 Forms フローの自動計算を代替)
    dist_s, nenpi_s, kankaku_s = _kyuyu_calc(vehicle, odometer, liters)
    if dist_s:
        fields[GAS_F_DIST] = dist_s
    if nenpi_s:
        fields[GAS_F_NENPI] = nenpi_s
    if kankaku_s:
        fields[GAS_F_KANKAKU] = kankaku_s
    try:
        new_id = sp_post_item(LIST_GASORIN, fields)
    except Exception as e:
        logging.exception("kyuyu apply (List54 insert) failed")
        return _json_response({"error": "insert_failed", "detail": str(e)}, 500)
    # Teams「送迎関係」へ通知(旧 Forms フローの通知を代替・best-effort=失敗しても申請は成功)
    try:
        _kyuyu_notify_teams(shain_no, name, vehicle, today, amount, liters, odometer,
                            dist_s, nenpi_s, kankaku_s, pay_type, purposes,
                            receipt_urls, odo_url,
                            receipt_bytes=receipt_bytes_list, odo_bytes_raw=odo_bytes)
    except Exception:
        logging.exception("kyuyu teams notify failed")
    # 浩俊へロック画面Push (メモ投稿は自己発言で通知が出ないため、Pushで気付けるように・best-effort)
    try:
        veh_short = re.sub(r"[\s　]+", " ", vehicle).strip()
        amt = f"¥{int(float(amount)):,}" if amount else ""
        _push_to_admins(KYUYU_PUSH_EMAILS, {
            "title": "⛽ 給油申請",
            "body": f"{name}（{veh_short}）{amt}".strip(),
            "url": f"{SITE_URL}/Lists/List54/AllItems.aspx",
            "tag": "kyuyu-apply", "badge": 1,
        })
    except Exception:
        logging.exception("kyuyu push failed")
    return _json_response({"ok": True, "id": new_id})


def _kyuyu_item_to_notify(it: Dict[str, Any]) -> None:
    """List54 アイテム(OData_付き) から通知を再構築し、送迎関係チャネル + Push へ送る(再送/テスト用)。"""
    def og(key):
        v = it.get("OData_" + key)
        return str(v) if v not in (None, "") else ""
    title = (it.get("Title") or "").strip()
    name = re.sub(r"\s*\([^)]*\)\s*$", "", re.sub(r"^\s*\d+\s*", "", title)).strip() if title else ""
    m = re.match(r"\s*(\d+)", title)
    shain_no = m.group(1) if m else ""
    vehicle = og(GAS_F_VEHICLE)
    amount, liters, odometer = og(GAS_F_AMOUNT), og(GAS_F_LITERS), og(GAS_F_ODO)
    dist_s, nenpi_s, kankaku_s = og(GAS_F_DIST), og(GAS_F_NENPI), og(GAS_F_KANKAKU)
    pay_type = og(GAS_F_PAYTYPE)
    purposes = [p for p in og(GAS_F_PURPOSE).split(",") if p] or ["ガソリン給油"]
    receipt_urls = [og(k) for k in [GAS_F_RECEIPT, GAS_F_RECEIPT2, GAS_F_RECEIPT3] if og(k)]
    odo_url = og(GAS_F_ODOIMG)
    _kyuyu_notify_teams(shain_no, name, vehicle, _today_jst(), amount, liters, odometer,
                        dist_s, nenpi_s, kankaku_s, pay_type, purposes, receipt_urls, odo_url)
    try:
        veh_short = re.sub(r"[\s　]+", " ", vehicle).strip()
        amt = f"¥{int(float(amount)):,}" if amount else ""
        _push_to_admins(KYUYU_PUSH_EMAILS, {
            "title": "⛽ 給油申請", "body": f"{name}（{veh_short}）{amt}".strip(),
            "url": f"{SITE_URL}/Lists/List54/AllItems.aspx", "tag": "kyuyu-apply", "badge": 1,
        })
    except Exception:
        logging.exception("kyuyu item push failed")


# ====== 送迎記録 (運転手の送迎実績・部課=094:本社（送迎者）) ======
# 送迎先ごとの時刻マスタ。行き/帰りは時刻ボタン、遅刻/早退/病院対応は時刻なしの特別区分。
SOUGEI_DESTS = [
    {"name": "ASTI浜松", "go": ["8:00"], "back": ["17:00", "18:00", "19:00", "20:00"]},
    {"name": "ユーシン", "go": ["9:00"], "back": ["17:40", "19:45"]},
    {"name": "ASTI掛川", "go": ["8:00"], "back": ["17:00", "18:00", "19:00"]},
]
SOUGEI_SPECIALS = ["遅刻", "早退", "病院対応", "その他（休業）"]
SOUGEI_BACKDATE_DAYS = 7   # 記録し忘れの遡り許容日数


def _sougei_kiroku_eligible(emp: Dict[str, Any]) -> bool:
    """送迎記録の対象 = 部課「094:本社（送迎者）」の在職者。"""
    buka = str(emp.get(F_BUKA) or "")
    return buka.startswith("094:") or "本社（送迎" in buka


def _sougei_own_records(shain_no: int, date_from: _dt.date, date_to: _dt.date) -> List[Dict[str, Any]]:
    """本人の送迎記録 (期間内・取消除く) を日付降順で返す。"""
    items = sp_get_items(
        LIST_SOUGEI_KIROKU,
        select=",".join(["Id", SK_SHAINNO, SK_NAME, SK_DATE, SK_DEST, SK_CAT, SK_SLOT, SK_MEMO, SK_STATUS]),
        filter_=f"{SK_SHAINNO} eq {shain_no}",
        orderby="Id desc", top=2000,
    )
    out = []
    for it in items:
        if (it.get(SK_STATUS) or "active") == "cancelled":
            continue
        d = _utc_to_jst_date(it.get(SK_DATE))
        if d is None or d < date_from or d > date_to:
            continue
        out.append({
            "id": it.get("Id"),
            "date": d.isoformat(),
            "destination": it.get(SK_DEST) or "",
            "category": it.get(SK_CAT) or "",
            "timeSlot": it.get(SK_SLOT) or "",
            "memo": it.get(SK_MEMO) or "",
        })
    out.sort(key=lambda r: (r["date"], r["id"]), reverse=True)
    return out


@app.route(route="sougei/config", methods=["GET", "OPTIONS"])
def sougei_config(req: func.HttpRequest) -> func.HttpResponse:
    """送迎記録の画面マスタ (送迎先/時刻/特別区分) + 本日の自分の記録。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    emp = find_active_employee_by_shain(payload["shainNo"])
    if not emp:
        return _json_response({"error": "not_active"}, 403)
    today = _today_jst()
    todays = []
    if _sougei_kiroku_eligible(emp):
        try:
            todays = _sougei_own_records(int(payload["shainNo"]), today, today)
        except Exception:
            logging.exception("sougei today fetch failed")
    return _json_response({
        "ok": True, "eligible": _sougei_kiroku_eligible(emp),
        "destinations": SOUGEI_DESTS, "specials": SOUGEI_SPECIALS,
        "backdateDays": SOUGEI_BACKDATE_DAYS,
        "today": today.isoformat(), "todayRecords": todays,
    })


@app.route(route="sougei/record", methods=["POST", "OPTIONS"])
def sougei_record(req: func.HttpRequest) -> func.HttpResponse:
    """送迎実績を1件記録。Body: {date?, destination, category, timeSlot?, memo?}
    category=行き/帰り は timeSlot 必須(マスタ内)。遅刻/早退/病院対応 は時刻なし。
    行き/帰りの同一 日付+送迎先+区分+時刻 は重複としてエラー(特別区分は複数回OK)。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    shain_no = int(payload["shainNo"])
    emp = find_active_employee_by_shain(shain_no)
    if not emp or not _sougei_kiroku_eligible(emp):
        return _json_response({"error": "not_eligible"}, 403)
    wait = check_rate_limit(shain_no, "sougei")
    if wait is not None:
        return _json_response({"error": "rate_limited", "retryAfterSeconds": wait}, 429,
                              extra_headers={"Retry-After": str(wait)})
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    dest = (body.get("destination") or "").strip()
    dest_conf = next((d for d in SOUGEI_DESTS if d["name"] == dest), None)
    if not dest_conf:
        return _json_response({"error": "invalid_destination"}, 400)
    category = (body.get("category") or "").strip()
    slot = (body.get("timeSlot") or "").strip()
    if category in ("行き", "帰り"):
        valid_slots = dest_conf["go"] if category == "行き" else dest_conf["back"]
        if slot not in valid_slots:
            return _json_response({"error": "invalid_timeslot"}, 400)
    elif category in SOUGEI_SPECIALS:
        slot = ""
    else:
        return _json_response({"error": "invalid_category"}, 400)
    memo = str(body.get("memo") or "").strip()[:200]
    # 日付: 既定=今日。遡りは SOUGEI_BACKDATE_DAYS 日まで、未来は不可
    today = _today_jst()
    date_s = (body.get("date") or "").strip()
    if date_s:
        try:
            rec_date = _dt.date.fromisoformat(date_s)
        except Exception:
            return _json_response({"error": "invalid_date"}, 400)
    else:
        rec_date = today
    if rec_date > today or (today - rec_date).days > SOUGEI_BACKDATE_DAYS:
        return _json_response({"error": "date_out_of_range", "backdateDays": SOUGEI_BACKDATE_DAYS}, 400)
    # 重複チェック (行き/帰りのみ。特別区分は同日複数回あり得る)
    if category in ("行き", "帰り"):
        try:
            existing = _sougei_own_records(shain_no, rec_date, rec_date)
            for r in existing:
                if r["destination"] == dest and r["category"] == category and r["timeSlot"] == slot:
                    return _json_response({"error": "duplicate_record", "existingId": r["id"]}, 409)
        except Exception:
            logging.exception("sougei duplicate check failed")
    name = emp.get(F_SHAIN_NAME) or ""
    fields = {
        "Title": f"{shain_no} {name} ({rec_date.isoformat()})",
        SK_SHAINNO: str(shain_no),
        SK_NAME: name,
        SK_DATE: rec_date.strftime("%Y/%m/%d"),
        SK_DEST: dest,
        SK_CAT: category,
        SK_STATUS: "active",
    }
    if slot:
        fields[SK_SLOT] = slot
    if memo:
        fields[SK_MEMO] = memo
    try:
        new_id = sp_post_item(LIST_SOUGEI_KIROKU, fields)
    except Exception as e:
        logging.exception("sougei record insert failed")
        return _json_response({"error": "insert_failed", "detail": str(e)}, 500)
    return _json_response({"ok": True, "id": new_id, "record": {
        "id": new_id, "date": rec_date.isoformat(), "destination": dest,
        "category": category, "timeSlot": slot, "memo": memo,
    }})


@app.route(route="sougei/records", methods=["GET", "OPTIONS"])
def sougei_records(req: func.HttpRequest) -> func.HttpResponse:
    """本人の送迎記録一覧。?month=YYYY-MM (既定=当月)。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    shain_no = int(payload["shainNo"])
    emp = find_active_employee_by_shain(shain_no)
    if not emp or not _sougei_kiroku_eligible(emp):
        return _json_response({"error": "not_eligible"}, 403)
    month_s = (req.params.get("month") or "").strip()
    today = _today_jst()
    try:
        if month_s:
            y, m = month_s.split("-")
            y, m = int(y), int(m)
        else:
            y, m = today.year, today.month
        first = _dt.date(y, m, 1)
        last = (_dt.date(y + 1, 1, 1) if m == 12 else _dt.date(y, m + 1, 1)) - _dt.timedelta(days=1)
    except Exception:
        return _json_response({"error": "invalid_month"}, 400)
    try:
        recs = _sougei_own_records(shain_no, first, last)
    except Exception as e:
        logging.exception("sougei records fetch failed")
        return _json_response({"error": "fetch_failed", "detail": str(e)}, 500)
    # 送迎先別の件数サマリ
    by_dest: Dict[str, int] = {}
    for r in recs:
        by_dest[r["destination"]] = by_dest.get(r["destination"], 0) + 1
    return _json_response({
        "ok": True, "month": f"{y:04d}-{m:02d}", "records": recs,
        "total": len(recs), "byDestination": by_dest,
    })


@app.route(route="sougei/cancel", methods=["POST", "OPTIONS"])
def sougei_cancel(req: func.HttpRequest) -> func.HttpResponse:
    """自分の送迎記録を取消 (Status=cancelled)。当月分のみ可
    (月を跨いだ分は月次集計対象のため本人は削除不可)。Body: {id}"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    shain_no = int(payload["shainNo"])
    emp = find_active_employee_by_shain(shain_no)
    if not emp or not _sougei_kiroku_eligible(emp):
        return _json_response({"error": "not_eligible"}, 403)
    wait = check_rate_limit(shain_no, "sougei")
    if wait is not None:
        return _json_response({"error": "rate_limited", "retryAfterSeconds": wait}, 429,
                              extra_headers={"Retry-After": str(wait)})
    try:
        body = req.get_json()
        rec_id = int(body.get("id"))
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    try:
        items = sp_get_items(
            LIST_SOUGEI_KIROKU,
            select=",".join(["Id", SK_SHAINNO, SK_DATE, SK_STATUS]),
            filter_=f"Id eq {rec_id}", top=1,
        )
    except Exception as e:
        logging.exception("sougei cancel fetch failed")
        return _json_response({"error": "fetch_failed", "detail": str(e)}, 500)
    if not items:
        return _json_response({"error": "not_found"}, 404)
    it = items[0]
    try:
        owner = int(float(it.get(SK_SHAINNO)))
    except Exception:
        owner = None
    if owner != shain_no:
        return _json_response({"error": "forbidden"}, 403)
    if (it.get(SK_STATUS) or "active") == "cancelled":
        return _json_response({"ok": True, "already": True})
    d = _utc_to_jst_date(it.get(SK_DATE))
    today = _today_jst()
    # 当月分のみ取消可 (月を跨いだら集計対象なので本人は削除不可・修正は総務のExcel「不足・追加」列で)
    same_month = d is not None and d.year == today.year and d.month == today.month
    if not same_month:
        return _json_response({"error": "too_old_to_cancel"}, 400)
    try:
        sp_patch_item(LIST_SOUGEI_KIROKU, rec_id, {SK_STATUS: "cancelled"})
    except Exception as e:
        logging.exception("sougei cancel patch failed")
        return _json_response({"error": "cancel_failed", "detail": str(e)}, 500)
    return _json_response({"ok": True, "id": rec_id})


# ====== 事故報告 (運転手の交通事故報告・部課=094:本社（送迎者）) ======
JIKO_FOLDER_BASE = "/sites/TeamStepup/Shared Documents/事故報告"
JIKO_TEAMS_EMAIL = os.environ.get("JIKO_TEAMS_EMAIL", "80ade604.team-stepup.com@jp.teams.ms")  # 送迎関係チャネル
JIKO_PUSH_EMAILS = [e.strip() for e in os.environ.get(
    "JIKO_PUSH_EMAILS", "h.yamashita@team-stepup.com,a.yamashita@team-stepup.com").split(",") if e.strip()]
JIKO_INJURY = ["なし", "あり"]
JIKO_POLICE = ["届出済", "未届"]


@app.route(route="jiko/config", methods=["GET", "OPTIONS"])
def jiko_config(req: func.HttpRequest) -> func.HttpResponse:
    """事故報告フォームのマスタ (車両リスト・担当車両・選択肢)。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    emp = find_active_employee_by_shain(payload["shainNo"])
    if not emp:
        return _json_response({"error": "not_active"}, 403)
    vehicles = _kyuyu_vehicles()
    return _json_response({
        "ok": True, "eligible": _sougei_kiroku_eligible(emp),
        "vehicles": vehicles,
        "defaultVehicle": _kyuyu_default_vehicle(int(payload["shainNo"]), vehicles),
        "injuryOptions": JIKO_INJURY, "policeOptions": JIKO_POLICE,
        "today": _today_jst().isoformat(),
    })


@app.route(route="jiko/report", methods=["POST", "OPTIONS"])
def jiko_report(req: func.HttpRequest) -> func.HttpResponse:
    """交通事故報告を提出。写真(最大3枚)をSP保存→List INSERT→送迎関係チャネル+役員Push即時通知。
    Body: {date, time, location, vehicle, otherParty?, injury, injuryDetail?,
           police, description, photos:[<dataURL>...](最大3・任意)}"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    shain_no = int(payload["shainNo"])
    emp = find_active_employee_by_shain(shain_no)
    if not emp or not _sougei_kiroku_eligible(emp):
        return _json_response({"error": "not_eligible"}, 403)
    wait = check_rate_limit(shain_no, "jiko")
    if wait is not None:
        return _json_response({"error": "rate_limited", "retryAfterSeconds": wait}, 429,
                              extra_headers={"Retry-After": str(wait)})
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    today = _today_jst()
    try:
        acc_date = _dt.date.fromisoformat((body.get("date") or "").strip())
    except Exception:
        return _json_response({"error": "invalid_date"}, 400)
    if acc_date > today or (today - acc_date).days > 30:
        return _json_response({"error": "date_out_of_range"}, 400)
    acc_time = str(body.get("time") or "").strip()[:10]
    location = str(body.get("location") or "").strip()[:120]
    vehicle = str(body.get("vehicle") or "").strip()[:60]
    description = str(body.get("description") or "").strip()[:2000]
    if not location or not description:
        return _json_response({"error": "missing_fields"}, 400)
    other_party = str(body.get("otherParty") or "").strip()[:1000]
    injury = (body.get("injury") or "なし").strip()
    if injury not in JIKO_INJURY:
        injury = "なし"
    injury_detail = str(body.get("injuryDetail") or "").strip()[:1000]
    police = (body.get("police") or "未届").strip()
    if police not in JIKO_POLICE:
        police = "未届"
    photos = [p for p in (body.get("photos") or []) if p][:3]
    try:
        photo_bytes = [base64.b64decode(_strip_data_url(p)) for p in photos]
    except Exception as e:
        return _json_response({"error": "invalid_base64", "detail": str(e)}, 400)
    if any(len(b) > 10 * 1024 * 1024 for b in photo_bytes):
        return _json_response({"error": "image_too_large", "maxMB": 10}, 400)
    name = emp.get(F_SHAIN_NAME) or ""
    # 写真を SP に保存
    photo_urls: List[str] = []
    try:
        year_folder = f"{JIKO_FOLDER_BASE}/{today.year}"
        sp_create_folder_if_not_exists(JIKO_FOLDER_BASE)
        sp_create_folder_if_not_exists(year_folder)
        ts = _now_jst().strftime("%Y%m%d_%H%M%S")
        url_base = f"{SITE_TEAMSTEPUP}/Shared Documents/事故報告/{today.year}"
        for i, pb in enumerate(photo_bytes):
            fname = _safe_filename(f"{shain_no}_{acc_date.isoformat()}_事故{i+1}_{ts}.jpg")
            sp_upload_file(year_folder, fname, pb)
            photo_urls.append(f"{url_base}/{fname}")
    except Exception as e:
        logging.exception("jiko photo upload failed")
        return _json_response({"error": "upload_failed", "detail": str(e)}, 500)
    fields: Dict[str, Any] = {
        "Title": f"{shain_no} {name} ({acc_date.isoformat()} {acc_time})".strip(),
        "ShainNo": str(shain_no),
        "DriverName": name,
        "AccidentDate": acc_date.strftime("%Y/%m/%d"),
        "AccidentTime": acc_time,
        "Location": location,
        "Vehicle": vehicle,
        "Injury": injury,
        "Police": police,
        "Description": description,
        "Status": "submitted",
    }
    if other_party:
        fields["OtherParty"] = other_party
    if injury_detail:
        fields["InjuryDetail"] = injury_detail
    for i, u in enumerate(photo_urls):
        fields[f"Photo{i+1}"] = u
    try:
        new_id = sp_post_item(LIST_JIKO, fields)
    except Exception as e:
        logging.exception("jiko insert failed")
        return _json_response({"error": "insert_failed", "detail": str(e)}, 500)
    # 通知 (best-effort・失敗しても報告は成功)
    try:
        _jiko_notify(shain_no, name, acc_date, acc_time, location, vehicle,
                     other_party, injury, injury_detail, police, description,
                     photo_urls, photo_bytes)
    except Exception:
        logging.exception("jiko notify failed")
    try:
        _push_to_admins(JIKO_PUSH_EMAILS, {
            "title": "🚨 交通事故報告",
            "body": f"{name}（{vehicle}）{location}".strip(),
            "url": f"{SITE_URL}/Lists/事故報告/AllItems.aspx",
            "tag": "jiko-report", "badge": 1,
        })
    except Exception:
        logging.exception("jiko push failed")
    return _json_response({"ok": True, "id": new_id})


def _jiko_notify(shain_no, name, acc_date, acc_time, location, vehicle,
                 other_party, injury, injury_detail, police, description,
                 photo_urls, photo_bytes) -> None:
    """事故報告を送迎関係チャネルへ HTMLカード + 写真添付で通知。"""
    esc = lambda s: (str(s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
    injury_s = injury + (f"（{injury_detail}）" if injury == "あり" and injury_detail else "")
    rows = [
        ("運転手", f"{esc(name)}（{shain_no}）"),
        ("発生日時", f"{acc_date.year}/{acc_date.month}/{acc_date.day} {esc(acc_time)}"),
        ("場所", esc(location)),
        ("車両", esc(vehicle)),
        ("相手", esc(other_party) or "なし"),
        ("けが人", esc(injury_s)),
        ("警察届出", esc(police)),
        ("状況", esc(description).replace("\n", "<br>")),
    ]
    tr = "".join(
        f'<tr><td style="padding:5px 12px;background:#fdf0ef;font-weight:600;white-space:nowrap;">{k}</td>'
        f'<td style="padding:5px 12px;">{v}</td></tr>' for k, v in rows)
    links = "".join(
        f'<p style="margin:4px 0;"><a href="{u}">📷 事故写真{i+1}を開く</a></p>'
        for i, u in enumerate(photo_urls))
    html = (
        '<div style="font-family:sans-serif;color:#1a2b3c;">'
        '<h3 style="margin:0 0 4px;color:#c0392b;">🚨 交通事故報告が届きました</h3>'
        f'<table style="border-collapse:collapse;font-size:14px;">{tr}</table>'
        f'{links}'
        f'<p style="margin:10px 0 0;"><a href="{SITE_URL}/Lists/%E4%BA%8B%E6%95%85%E5%A0%B1%E5%91%8A/AllItems.aspx">📊 リストで全体を確認する</a></p>'
        '</div>')
    # 写真は縮小して添付 (kyuyu と同じ非インライン方式・重複表示を防ぐ)
    attachments = []
    used = 0
    for i, pb in enumerate(photo_bytes or []):
        small = _resize_image_jpeg(pb)
        if not small or used + len(small) > 3_200_000:
            continue
        used += len(small)
        attachments.append({
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": f"事故写真{i+1}.jpg",
            "contentType": "image/jpeg",
            "contentBytes": base64.b64encode(small).decode("ascii"),
        })
    _send_notification_mail(f"🚨 交通事故報告 {name}", html, to_addr=JIKO_TEAMS_EMAIL,
                            html=True, attachments=attachments or None)


@app.route(route="jiko/reports", methods=["GET", "OPTIONS"])
def jiko_reports(req: func.HttpRequest) -> func.HttpResponse:
    """本人の事故報告一覧 (新しい順・最大50件)。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    shain_no = int(payload["shainNo"])
    emp = find_active_employee_by_shain(shain_no)
    if not emp or not _sougei_kiroku_eligible(emp):
        return _json_response({"error": "not_eligible"}, 403)
    try:
        items = sp_get_items(
            LIST_JIKO,
            select=",".join(["Id", "ShainNo", "AccidentDate", "AccidentTime",
                             "Location", "Vehicle", "Injury", "Police", "Status"]),
            filter_=f"ShainNo eq {shain_no}",
            orderby="Id desc", top=50,
        )
    except Exception as e:
        logging.exception("jiko reports fetch failed")
        return _json_response({"error": "fetch_failed", "detail": str(e)}, 500)
    out = []
    for it in items:
        d = _utc_to_jst_date(it.get("AccidentDate"))
        out.append({
            "id": it.get("Id"),
            "date": d.isoformat() if d else "",
            "time": it.get("AccidentTime") or "",
            "location": it.get("Location") or "",
            "vehicle": it.get("Vehicle") or "",
            "injury": it.get("Injury") or "",
            "police": it.get("Police") or "",
        })
    return _json_response({"ok": True, "reports": out})


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

    ocr_result = None
    # GPT-4o-mini 優先
    if _gpt_ocr_available():
        try:
            import gpt_ocr
            g = gpt_ocr.extract_doc_fields("license", [front_bytes])
            mapped = _gpt_doc_to_di("license", g)
            if mapped.get("licenseNumber") or mapped.get("licenseExpiry") or mapped.get("name"):
                mapped["engine"] = "gpt-4o-mini"
                ocr_result = mapped
        except Exception as ge:
            logging.warning(f"license GPT OCR failed, fallback to DI: {ge}")
    if ocr_result is None:
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
            ocr_result["engine"] = "di"
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

    # 車体の色
    m = re.search(r'(?:車体の?色|色)[\s:：]*([白黒赤青銀灰緑黄茶紫桃金][^\n]{0,6})', st)
    if m:
        out['carColor'] = m.group(1).strip()
    else:
        for col in ("シルバー", "ホワイト", "ブラック", "ホワイトパール", "グレー", "ブルー", "レッド", "ガン"):
            if col in st:
                out['carColor'] = col
                break

    # 対象車両 (車検証は自動車のもの → 既定で「自動車」)
    out['carType'] = '自動車'

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

    # 補償内容: 対人/対物/搭乗者/人身/車両保険
    def _find_amount(labels: List[str]) -> Optional[str]:
        for lbl in labels:
            m = re.search(lbl + r'[\s\S]{0,15}?(無制限|\d[\d,]*\s*(?:万円|億円|万|円))', t)
            if m:
                return m.group(1).replace(' ', '')
        return None
    taijin = _find_amount([r'対人賠償', r'対人'])
    if taijin:
        out['taijin'] = taijin
    taibutsu = _find_amount([r'対物賠償', r'対物'])
    if taibutsu:
        out['taibutsu'] = taibutsu
    tojosha = _find_amount([r'搭乗者傷害', r'搭乗者'])
    if tojosha:
        out['tojoshaShogai'] = tojosha
    jinshin = _find_amount([r'人身傷害'])
    if jinshin:
        out['jinshinShogai'] = jinshin
    # 車両保険: 有/無 or 金額
    m = re.search(r'車両保険[\s\S]{0,15}?(無制限|あり|なし|有|無|\d[\d,]*\s*万円)', t)
    if m:
        out['carHoken'] = m.group(1)
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
    result = None
    if _gpt_ocr_available():
        try:
            import gpt_ocr
            g1 = gpt_ocr.extract_doc_fields("shaken", [shaken_bytes])
            g2 = gpt_ocr.extract_doc_fields("jibaiseki", [jibai_bytes])
            merged = _gpt_doc_to_di("shaken", g1)
            merged.update(_gpt_doc_to_di("jibaiseki", g2))
            if merged.get("shakenKigen") or merged.get("carNumber") or merged.get("jibaiKigen"):
                merged["engine"] = "gpt-4o-mini"
                result = merged
        except Exception as ge:
            logging.warning(f"shaken GPT OCR failed, fallback to DI: {ge}")
    if result is None:
        try:
            shaken_txt = _run_read_ocr(shaken_bytes)
            jibai_txt = _run_read_ocr(jibai_bytes)
            result = parse_shaken_jibaiseki_text(shaken_txt, jibai_txt)
            result["engine"] = "di"
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
            _set_text(pf2, F_CAR_COLOR, confirmed.get("carColor"), 20, updated, "carColor")
            _set_text(pf2, F_CAR_TYPE, confirmed.get("carType"), 20, updated, "carType")
            _set_date(pf2, F_SHAKEN_KIGEN, confirmed.get("shakenKigen"), updated, "shakenKigen")
            _set_date(pf2, F_JIBAI_KIGEN, confirmed.get("jibaiKigen"), updated, "jibaiKigen")
            # Choice 系 (メーカー/自賠責会社/色/対象車両) は値がリストに無いと失敗するので best-effort
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
    result = None
    if _gpt_ocr_available():
        try:
            import gpt_ocr
            g = gpt_ocr.extract_doc_fields("nini", [img_bytes])
            mapped = _gpt_doc_to_di("nini", g)
            if mapped.get("niniKaisha") or mapped.get("niniKigen") or mapped.get("niniShoken"):
                mapped["engine"] = "gpt-4o-mini"
                result = mapped
        except Exception as ge:
            logging.warning(f"hoken GPT OCR failed, fallback to DI: {ge}")
    if result is None:
        try:
            txt = _run_read_ocr(img_bytes)
            result = parse_nini_hoken_text(txt)
            result["engine"] = "di"
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
            # 補償内容 (対人/対物は Choice、他は Text)
            _set_text(pf2, F_HOKEN_TAIJIN, confirmed.get("taijin"), 30, updated, "taijin")
            _set_text(pf2, F_HOKEN_TAIBUTSU, confirmed.get("taibutsu"), 30, updated, "taibutsu")
            _set_text(pf2, F_TOJOSHA, confirmed.get("tojoshaShogai"), 30, updated, "tojoshaShogai")
            _set_text(pf2, F_JINSHIN, confirmed.get("jinshinShogai"), 30, updated, "jinshinShogai")
            _set_text(pf2, F_SHARYO_HOKEN, confirmed.get("carHoken"), 30, updated, "carHoken")
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
            # 国籍 → 本(国)籍フィールド
            nationality = (confirmed.get("nationality") or "").strip()
            if nationality and len(nationality) <= 50:
                patch_fields[F_KOKUSEKI] = nationality
                updated_fields.append("nationality")
            # 在留カードの氏名は英字。カナの「名前(Title)」は上書きせず、
            # 在留カード氏名(特記事項1)フィールドへ反映する。
            romaji_name = (confirmed.get("name") or "").strip()
            if romaji_name and len(romaji_name) <= 100:
                patch_fields[F_ZAIRYU_NAME] = romaji_name
                updated_fields.append("name")
            if patch_fields:
                target_id = emp.get("Id") if emp else None
                if target_id:
                    try:
                        sp_patch_item(LIST_SHAIN, int(target_id), patch_fields)
                    except Exception as e:
                        logging.exception("update employee zairyu failed")
                        # 画像保存は成功しているので警告のみ
        # 提出完了を総務へ通知 + 登録済みの更新予定をクリア (提出されたため)
        try:
            had_plan = bool(emp.get(F_RENEW_PLAN))
            if had_plan:
                try:
                    sp_patch_item(LIST_SHAIN, int(emp.get("Id")), {F_RENEW_PLAN: None, F_RENEW_NOTE: ""})
                except Exception:
                    logging.exception("clear renew plan failed")
            name = emp.get(F_SHAIN_NAME) or ""
            haken = strip_buka_prefix(emp.get(F_BUKA) or "")
            new_kigen = ((body.get("confirmedData") or {}).get("zairyuKigen") or "").strip()
            lines = [
                "在留カードが本人からアプリで提出されました。",
                f"社員番号: {shain_no}",
                f"氏名: {name}",
                f"派遣先: {haken}",
            ]
            if new_kigen:
                lines.append(f"新しい在留期限: {new_kigen}")
            if had_plan:
                lines.append("（登録されていた更新予定はクリアしました）")
            subject = f"【在留カード提出】{shain_no} {name}"
            _send_notification_mail(subject, "\n".join(lines), to_addr=os.environ.get("EXPIRY_NOTIFY_EMAIL") or None)
        except Exception:
            logging.exception("zairyu submit notify failed")
        return _json_response({
            "ok": True,
            "folder": zairyu_folder,
            "files": [front_name, back_name],
            "updatedFields": updated_fields,
        })
    except Exception as e:
        logging.exception("zairyu_submit failed")
        return _json_response({"error": "internal", "detail": str(e)}, 500)


# ====== 必要書類請求 (在職証明書 / 源泉徴収票) ======
SHORUI_TYPES = {"在職証明書", "源泉徴収票"}


def _get_graph_token() -> str:
    """Managed Identity で Microsoft Graph アクセストークンを取得。"""
    cred = DefaultAzureCredential()
    tok = cred.get_token("https://graph.microsoft.com/.default")
    return tok.token


def _send_notification_mail(subject: str, text: str, to_addr: Optional[str] = None,
                            html: bool = False, attachments: Optional[list] = None) -> None:
    """Teams チャネルのメールアドレス宛に Graph sendMail で通知。
    送信元 = SHORUI_MAIL_SENDER (h.yamashita)、宛先 = to_addr or SHORUI_NOTIFY_EMAIL (チャネルメール)。
    html=True なら contentType=HTML。attachments=Graph fileAttachment 配列(インライン画像等)。
    環境変数未設定なら何もしない (申請自体は成功扱い)。"""
    to_addr = (to_addr or os.environ.get("SHORUI_NOTIFY_EMAIL", "")).strip()
    sender = os.environ.get("SHORUI_MAIL_SENDER", "").strip()
    if not to_addr or not sender:
        return
    try:
        token = _get_graph_token()
        message = {
            "subject": subject,
            "body": {"contentType": "HTML" if html else "Text", "content": text},
            "toRecipients": [{"emailAddress": {"address": to_addr}}],
        }
        if attachments:
            message["attachments"] = attachments
        r = requests.post(
            f"https://graph.microsoft.com/v1.0/users/{sender}/sendMail",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"message": message, "saveToSentItems": False}, timeout=60,
        )
        if not r.ok:
            logging.warning(f"sendMail failed: {r.status_code} {r.text[:300]}")
    except Exception as e:
        logging.warning(f"notification mail failed: {e}")


# ====== Web Push (Part C) ======
_VAPID_PEM_PATH = None


def _vapid_pem_path() -> Optional[str]:
    """VAPID 秘密鍵 (base64 PEM) を /tmp に書き出してパスを返す。"""
    global _VAPID_PEM_PATH
    if _VAPID_PEM_PATH:
        return _VAPID_PEM_PATH
    b64 = os.environ.get("VAPID_PRIVATE_KEY_B64", "").strip()
    if not b64:
        return None
    try:
        pem = base64.b64decode(b64).decode()
        path = os.path.join(tempfile.gettempdir(), "vapid_private.pem")
        with open(path, "w") as f:
            f.write(pem)
        _VAPID_PEM_PATH = path
        return path
    except Exception:
        logging.exception("vapid pem write failed")
        return None


def _send_web_push(sub_json: str, payload: Dict[str, Any]) -> str:
    """1件の購読へ Web Push 送信。戻り: 'ok' | 'gone'(失効) | 'skip' | 'error'。"""
    pem = _vapid_pem_path()
    sub = os.environ.get("VAPID_SUBJECT", "").strip()
    if not pem or not sub or not sub_json:
        return "skip"
    try:
        from pywebpush import webpush, WebPushException
        webpush(
            subscription_info=json.loads(sub_json),
            data=json.dumps(payload, ensure_ascii=False),
            vapid_private_key=pem,
            vapid_claims={"sub": sub},
            ttl=86400,  # WNS(Windows PCのEdge/Chrome)はTTL=0を拒否("Ttl value conflicts with X-WNS-Cache-Policy")。明示必須
            timeout=15,
        )
        return "ok"
    except Exception as e:
        # 404/410 = 購読失効
        status = getattr(getattr(e, "response", None), "status_code", None)
        if status in (404, 410):
            return "gone"
        logging.warning(f"web push failed: {e}")
        return "error"


# ====== 承認者(yukyu-app)向け Push: 前払い/弁当の申請が出たら担当者へロック画面通知 ======
# 担当者名 → メール (前払いは派遣先別担当者へ通知)
TANTOU_EMAIL = {
    "吉浦マルセロ": "m.yoshiura@team-stepup.com",
    "森ゆうじ": "y.mori@team-stepup.com",
    "森まこと": "m.mori@team-stepup.com",
}
ADMIN_APPROVER_EMAILS = ["h.yamashita@team-stepup.com", "a.yamashita@team-stepup.com"]
BENTO_APPROVER_EMAILS = ADMIN_APPROVER_EMAILS + ["m.mori@team-stepup.com", "y.mori@team-stepup.com"]
YUKYU_APP_URL = "https://team-stepup.github.io/yukyu-mobile-app/"


def _admin_push_list(email: str):
    """email の購読を全件(list[dict])で返す。1人で複数端末(スマホ+PC等)に対応。
    旧形式(Subが単一オブジェクト)も配列化して扱う(後方互換)。"""
    table = _get_lockout_table()
    if not table or not email:
        return []
    try:
        e = table.get_entity(partition_key="adminpush", row_key=email.lower())
        raw = e.get("Sub")
        if not raw:
            return []
        data = json.loads(raw)
        if isinstance(data, list):
            return [s for s in data if isinstance(s, dict) and s.get("endpoint")]
        if isinstance(data, dict) and data.get("endpoint"):
            return [data]
        return []
    except ResourceNotFoundError:
        return []
    except Exception:
        return []


def _admin_push_write(email: str, subs) -> bool:
    """購読リストを保存。空なら行ごと削除。"""
    table = _get_lockout_table()
    if not table or not email:
        return False
    try:
        if subs:
            table.upsert_entity({"PartitionKey": "adminpush", "RowKey": email.lower(),
                                 "Sub": json.dumps(subs, ensure_ascii=False)},
                                mode=UpdateMode.REPLACE)
        else:
            try:
                table.delete_entity(partition_key="adminpush", row_key=email.lower())
            except Exception:
                pass
        return True
    except Exception as e:
        logging.warning(f"admin push write failed: {e}")
        return False


def _admin_push_save(email: str, sub) -> bool:
    """端末の購読を追加(同一endpointは置換)。複数端末=スマホ+PCに対応。"""
    if not isinstance(sub, dict) or not sub.get("endpoint"):
        return False
    ep = sub.get("endpoint")
    subs = [s for s in _admin_push_list(email) if s.get("endpoint") != ep]   # 同一端末は重複させない
    subs.append(sub)
    if len(subs) > 20:                                                       # 暴走防止の上限
        subs = subs[-20:]
    return _admin_push_write(email, subs)


def _admin_push_delete(email: str, endpoint: str = None) -> None:
    """endpoint 指定でその端末だけ解除。未指定なら全端末解除。"""
    if endpoint:
        subs = [s for s in _admin_push_list(email) if s.get("endpoint") != endpoint]
        _admin_push_write(email, subs)
    else:
        _admin_push_write(email, [])


def _push_to_admins(emails, payload) -> int:
    """承認者(yukyu-app)の全端末へ Push。失効端末(404/410)だけ個別削除。送信成功数を返す。"""
    sent = 0
    seen = set()
    for em in emails:
        if not em:
            continue
        em = em.lower()
        if em in seen:
            continue
        seen.add(em)
        subs = _admin_push_list(em)
        if not subs:
            continue
        survivors = []
        changed = False
        for s in subs:
            r = _send_web_push(json.dumps(s, ensure_ascii=False), payload)
            if r == "ok":
                sent += 1
                survivors.append(s)
            elif r == "gone":
                changed = True           # 失効端末は survivors に入れない=除外
            else:
                survivors.append(s)      # 一時エラーは保持
        if changed:
            _admin_push_write(em, survivors)
    return sent


@app.route(route="push/subscribe-staff", methods=["POST", "OPTIONS"])
def push_subscribe_staff(req: func.HttpRequest) -> func.HttpResponse:
    """yukyu-app 利用者(役員/担当者)の Web Push 購読を保存。
    Body: { subscription: <PushSubscription JSON> } | { unsubscribe: true }"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    email, err = require_staff_auth(req)
    if err:
        return err
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    if body.get("unsubscribe"):
        _admin_push_delete(email, body.get("endpoint"))   # endpoint指定でその端末だけ解除(未指定は全端末)
        return _json_response({"ok": True, "subscribed": False})
    sub = body.get("subscription")
    if not sub or not isinstance(sub, dict) or not sub.get("endpoint"):
        return _json_response({"error": "invalid_subscription"}, 400)
    if len(json.dumps(sub, ensure_ascii=False)) > 4000:
        return _json_response({"error": "subscription_too_large"}, 400)
    ok = _admin_push_save(email, sub)
    return _json_response({"ok": ok, "subscribed": ok})


@app.route(route="push/test", methods=["POST", "OPTIONS"])
def push_test(req: func.HttpRequest) -> func.HttpResponse:
    """承認者が自分宛に前払い/弁当のテスト通知を送って動作確認する。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    email, err = require_staff_auth(req)
    if err:
        return err
    subs = _admin_push_list(email)
    if not subs:
        return _json_response({"ok": False, "error": "no_subscription",
                               "message": "先に「通知をオン」にしてください（ホーム画面に追加したアプリから）"}, 400)
    survivors = []
    n_ok = 0
    for s in subs:                                  # 登録された全端末(スマホ+PC等)へ送る
        sj = json.dumps(s, ensure_ascii=False)
        r1 = _send_web_push(sj, {"title": "💴 前払い申請 承認待ち（テスト）",
                                 "body": "テスト 太郎 / 30,000円 / 金曜日", "url": YUKYU_APP_URL,
                                 "tag": "maebari-test", "badge": 1})
        r2 = _send_web_push(sj, {"title": "🍱 弁当注文 承認待ち（テスト）",
                                 "body": "テスト 花子 / ① 315円", "url": YUKYU_APP_URL,
                                 "tag": "bento-test", "badge": 2})
        if r1 == "gone" or r2 == "gone":
            continue                                # 失効端末は survivors から除外
        survivors.append(s)
        if r1 == "ok" and r2 == "ok":
            n_ok += 1
    if len(survivors) != len(subs):
        _admin_push_write(email, survivors)
    return _json_response({"ok": n_ok > 0, "devices": len(subs), "sent": n_ok,
                           "maebari": "ok" if n_ok > 0 else "error",
                           "bento": "ok" if n_ok > 0 else "error"})


@app.route(route="nippou/notify", methods=["POST", "OPTIONS"])
def nippou_notify(req: func.HttpRequest) -> func.HttpResponse:
    """PC(nippou_auto)が日報同期(役員アプリ用)の完了後に呼ぶ。
    認証ユーザー(=山下)自身の登録端末へ『業務日報を更新しました』Push＋アプリアイコンのバッジを送る。
    Body(任意): { title, body, badge }"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    email, err = require_staff_auth(req)
    if err:
        return err
    try:
        body = req.get_json() or {}
    except Exception:
        body = {}
    title = body.get("title") or "📋 業務日報を更新しました"
    msg = body.get("body") or ""
    try:
        badge = int(body.get("badge", 1))
    except Exception:
        badge = 1
    n = _push_to_admins([email], {"title": title, "body": msg, "url": YUKYU_APP_URL,
                                  "tag": "nikkou-sync", "badge": badge})
    return _json_response({"ok": n > 0, "sent": n})


@app.route(route="nippou/summarize", methods=["POST", "OPTIONS"])
def nippou_summarize(req: func.HttpRequest) -> func.HttpResponse:
    """PC(nippou_auto)が本日の業務日報(整形テキスト)を渡し、Claudeで役員向け要約を返す。
    Body: { text }  ->  { ok, summary }"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    email, err = require_staff_auth(req)
    if err:
        return err
    try:
        body = req.get_json() or {}
    except Exception:
        body = {}
    text = (body.get("text") or "").strip()
    if not text:
        return _json_response({"error": "empty_text"}, 400)
    try:
        import gpt_ocr
        summary = gpt_ocr.summarize_nippou(text)
        return _json_response({"ok": True, "summary": summary})
    except Exception as e:
        logging.exception("nippou_summarize failed")
        return _json_response({"ok": False, "error": str(e)[:200]}, 500)


# ===== 産休育休「注(申請遅れ)」の朝の通知 (毎朝8:30 JST・稼働日のみ・スマホ+PC両方) =====
SANKYU_LIST_PATH = "/sites/PowerApps/Lists/List52"
_SANKYU_SKIP_FIELDS = {'ContentType', 'Modified', 'Created', 'Author', 'Editor', '_UIVersionString',
                       'Attachments', 'Edit', 'LinkTitleNoMenu', 'LinkTitle', 'DocIcon',
                       'ItemChildCount', 'FolderChildCount'}


def _sankyu_overdue_now():
    """List52を読み、産休育休の申請遅れ件数と内訳を返す (yukyu-app の「注N」と同じ計算)。→ (total, details)"""
    import sankyu_overdue as _s   # 遅延import (gpt_ocr と同方式)
    base = f"{SITE_URL}/_api/web/GetList('{SANKYU_LIST_PATH}')"
    h = _sp_headers(verbose=True)
    fr = requests.get(base + "/fields?$filter=Hidden eq false&$select=InternalName,Title,TypeAsString&$top=300",
                      headers=h, timeout=60)
    fr.raise_for_status()
    meta = {}
    for f in fr.json().get("d", {}).get("results", []):
        iname = f.get("InternalName")
        if not iname or iname in _SANKYU_SKIP_FIELDS:
            continue
        if iname in ("Title", "Id", "ID"):
            dk = iname
        elif iname.startswith("_"):
            dk = "OData_" + iname
        else:
            dk = iname
        meta[iname] = {"title": f.get("Title"), "type": f.get("TypeAsString"), "dataKey": dk}
    items = []
    url = base + "/items?$top=500"
    while url:
        ir = requests.get(url, headers=h, timeout=60)
        ir.raise_for_status()
        d = ir.json().get("d", {})
        items.extend(d.get("results", []))
        url = d.get("__next")
    return _s.compute_overdue(items, meta)


def _sankyu_send_overdue_push():
    """遅れがあれば承認者(役員)のスマホ+PC全端末へPush。送信成功数を返す。"""
    total, details = _sankyu_overdue_now()
    if total <= 0:
        return {"ok": True, "overdue": 0, "sent": 0}
    names = "、".join((d.get("name") or "") for d in details[:5])
    if len(details) > 5:
        names += " ほか"
    sent = _push_to_admins(ADMIN_APPROVER_EMAILS, {
        "title": "👶 産休育休 申請遅れ",
        "body": f"申請遅れ {total}件（{names}）。手続きの申請日をご確認ください。",
        "url": YUKYU_APP_URL,
        "tag": "sankyu-overdue",
        "badge": total,
    })
    return {"ok": True, "overdue": total, "sent": sent}


@app.timer_trigger(arg_name="timer", schedule="0 15 0 * * *", run_on_startup=False, use_monitor=True)
def sankyu_overdue_timer(timer: func.TimerRequest) -> None:
    """毎朝9:15 JST(=00:15 UTC)。稼働日のみ、産休育休の申請遅れがあれば承認者へPush(スマホ+PC)。
    休日(土日・会社休日List)はスキップ → 休日明け最初の稼働日に当日の遅れ全件を送る(その時点で再計算するので取りこぼし無し)。"""
    try:
        d = _today_jst()
        if not _bento_is_working_day(d):
            logging.info("産休育休 遅れ通知: %s は稼働日でないためスキップ", d)
            return
        res = _sankyu_send_overdue_push()
        logging.info("産休育休 遅れ通知: overdue=%s sent=%s", res.get("overdue"), res.get("sent"))
        try:
            sres = _soumu_send_push()
            logging.info("社保・離職票 通知: counts=%s sent=%s", sres.get("counts"), sres.get("sent"))
        except Exception:
            logging.exception("社保・離職票 通知に失敗")
    except Exception:
        logging.exception("産休育休 遅れ通知(timer)に失敗")


@app.route(route="sankyu/overdue-test", methods=["POST", "GET", "OPTIONS"])
def sankyu_overdue_test(req: func.HttpRequest) -> func.HttpResponse:
    """手動確認用: 現在の産休育休の遅れ件数・内訳を返す。?push=1 で実際にPushも送る。総務/役員のみ。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    email, err = require_staff_auth(req)
    if err:
        return err
    try:
        total, details = _sankyu_overdue_now()
        sent = 0
        if req.params.get("push") == "1" and total > 0:
            sent = _sankyu_send_overdue_push().get("sent", 0)
        return _json_response({"ok": True, "overdue": total, "sent": sent,
                               "details": [{"name": d.get("name"), "count": d.get("count"),
                                            "procedures": d.get("procedures")} for d in details]})
    except Exception as e:
        logging.exception("sankyu_overdue_test failed")
        return _json_response({"error": "internal", "detail": str(e)}, 500)


# ===== 社保 加入/喪失・離職票 の朝の通知 (産休育休と同じ8:30タイマーで) =====
SOUMU_SHAHO_IN_ROOT = "/Lists/List9"   # ルートサイトの社保加入申請リスト
F_YUKYU_TAISHA = "OData__x9000__x8077__x5e74__x6708__x65"   # 退職年月(有給使用後退社)
F_HAKEN_TAISHA = "OData__x6d3e__x9063__x5148__x9000__x79"   # 派遣先退社
_SOUMU_ALLOWED_KUBUN = {"", "新規加入", "扶養追加", "氏名変更", "その他"}


def _parse_jst_date(raw):
    if not raw:
        return None
    try:
        d = _dt.datetime.fromisoformat(str(raw).replace("Z", "+00:00"))
        if d.tzinfo is not None:
            d = d.astimezone(_dt.timezone(_dt.timedelta(hours=9))).replace(tzinfo=None)
        return d.date()
    except Exception:
        return None


def _emp_field_map():
    """社員Listの {フィールドTitle: dataKey(EntityPropertyName)} を返す。"""
    r = requests.get(f"{SITE_URL}/_api/web/lists(guid'{LIST_SHAIN}')/fields?$select=EntityPropertyName,Title&$top=400",
                     headers=_sp_headers(verbose=True), timeout=60)
    r.raise_for_status()
    m = {}
    for f in r.json().get("d", {}).get("results", []):
        epn = f.get("EntityPropertyName")
        if epn:
            m[f.get("Title")] = epn
    return m


def _norm_bango(v):
    s = str(v if v is not None else "").strip()
    return str(int(s)) if s.isdigit() else s


def _soumu_counts():
    """社保加入/社保喪失/離職票 の件数 (yukyu-app 総務ダッシュボードと同じ計算)。kanyuはroot読取不可ならNone。"""
    today = _today_jst()
    fm = _emp_field_map()
    k_shaho = fm.get("総務-社保")
    k_rishoku = fm.get("総務-離職票")
    k_rhyou = fm.get("離職票risyokuhyou")
    k_zai = fm.get("在職zaisyoku")
    # 社員アイテムは _emp_field_map(verbose) のキーに合わせ verbose(d.results)で読む
    emps = []
    eurl = f"{SITE_URL}/_api/web/lists(guid'{LIST_SHAIN}')/items?$top=5000"
    eh = _sp_headers(verbose=True)
    while eurl:
        er = requests.get(eurl, headers=eh, timeout=90)
        er.raise_for_status()
        ed = er.json().get("d", {})
        emps.extend(ed.get("results", []))
        eurl = ed.get("__next")

    def zai(e):
        if k_zai:
            return str(e.get(k_zai) or "")
        return "退社" if (e.get(F_YUKYU_TAISHA) or e.get(F_HAKEN_TAISHA)) else "在職"

    # 社保喪失: 退社日(遅い方)<=今日 AND 総務-社保が「取得完了」かつ「喪失」でない
    shitsu = 0
    for e in emps:
        if not e.get("Title"):
            continue
        yt = _parse_jst_date(e.get(F_YUKYU_TAISHA))
        ht = _parse_jst_date(e.get(F_HAKEN_TAISHA))
        ds = [d for d in (yt, ht) if d]
        if not ds:
            continue
        if max(ds) > today:
            continue
        v = str((k_shaho and e.get(k_shaho)) or "")
        if "取得完了" in v and "喪失" not in v:
            shitsu += 1

    # 離職票: 退社 AND 離職票=「必要」 AND 総務-離職票 in {不要,空欄}
    rishoku = 0
    if k_rhyou and k_rishoku:
        for e in emps:
            if not e.get("Title"):
                continue
            if zai(e) != "退社":
                continue
            if str(e.get(k_rhyou) or "").strip() != "必要":
                continue
            if str(e.get(k_rishoku) or "").strip() not in ("不要", ""):
                continue
            rishoku += 1

    # 社保加入: ルートサイト List9 ＋ 社員「総務-社保」で重複除外。root未許可なら None
    kanyu = None
    try:
        excl = set()
        for e in emps:
            sv = str((k_shaho and e.get(k_shaho)) or "").strip()
            if sv and ("申請" in sv or "完了" in sv or "喪失" in sv or "移動" in sv):
                bn = _norm_bango(e.get(F_SHAIN_NO))
                if bn:
                    excl.add(bn)
        lbase = f"{SP_RESOURCE}/_api/web/GetList('{SOUMU_SHAHO_IN_ROOT}')"
        h = _sp_headers(verbose=True)
        lf = requests.get(lbase + "/fields?$select=EntityPropertyName,Title&$top=300", headers=h, timeout=60)
        lf.raise_for_status()
        lt2k = {}
        for f in lf.json().get("d", {}).get("results", []):
            epn = f.get("EntityPropertyName")
            if epn:
                lt2k[f.get("Title")] = epn

        def findk(*kw):
            for k in kw:
                if k in lt2k:
                    return lt2k[k]
            for k in kw:
                for t, v in lt2k.items():
                    if k in (t or ""):
                        return v
            return None
        kK = findk("加入日"); kHon = findk("本人加入"); kKub = findk("申請区分"); kBan = findk("社員番号", "社番", "番号")
        items = []
        url = lbase + "/items?$top=2000"
        while url:
            ir = requests.get(url, headers=h, timeout=60)
            ir.raise_for_status()
            dd = ir.json().get("d", {})
            items.extend(dd.get("results", []))
            url = dd.get("__next")
        cnt = 0
        for it in items:
            kd = _parse_jst_date(it.get(kK)) if kK else None
            if not kd or kd >= today:
                continue
            if kHon and str(it.get(kHon) or "").strip() != "":
                continue
            if kKub and str(it.get(kKub) or "").strip() not in _SOUMU_ALLOWED_KUBUN:
                continue
            bn = _norm_bango(it.get(kBan)) if kBan else ""
            if bn and bn in excl:
                continue
            cnt += 1
        kanyu = cnt
    except Exception:
        logging.exception("社保加入カウント(root List9)読取に失敗")
        kanyu = None

    return {"kanyu": kanyu, "shitsu": shitsu, "rishoku": rishoku}


def _soumu_send_push():
    c = _soumu_counts()
    parts = []
    if c.get("kanyu"):
        parts.append(f"社保加入 {c['kanyu']}名")
    if c.get("shitsu"):
        parts.append(f"社保喪失 {c['shitsu']}名")
    if c.get("rishoku"):
        parts.append(f"離職票 {c['rishoku']}名")
    if not parts:
        return {"ok": True, "counts": c, "sent": 0}
    badge = (c.get("kanyu") or 0) + (c.get("shitsu") or 0) + (c.get("rishoku") or 0)
    sent = _push_to_admins(ADMIN_APPROVER_EMAILS, {
        "title": "📋 社保・離職票 手続き",
        "body": "未処理: " + "、".join(parts) + "。総務ダッシュボードをご確認ください。",
        "url": YUKYU_APP_URL,
        "tag": "soumu-tetsuduki",
        "badge": badge,
    })
    return {"ok": True, "counts": c, "sent": sent}


@app.route(route="sankyu/soumu-test", methods=["POST", "GET", "OPTIONS"])
def soumu_counts_test(req: func.HttpRequest) -> func.HttpResponse:
    """手動確認用: 社保加入/喪失・離職票の件数を返す。?push=1 で実際にPushも送る。総務/役員のみ。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    email, err = require_staff_auth(req)
    if err:
        return err
    try:
        c = _soumu_counts()
        sent = 0
        if req.params.get("push") == "1":
            sent = _soumu_send_push().get("sent", 0)
        dbg = None
        if req.params.get("debug") == "1":
            fm = _emp_field_map()
            emps = []
            eurl = f"{SITE_URL}/_api/web/lists(guid'{LIST_SHAIN}')/items?$top=5000"
            eh = _sp_headers(verbose=True)
            while eurl:
                er = requests.get(eurl, headers=eh, timeout=90); er.raise_for_status()
                ed = er.json().get("d", {}); emps.extend(ed.get("results", [])); eurl = ed.get("__next")
            ksh = fm.get("総務-社保")
            today = _today_jst()
            cands = []
            for e in emps:
                v = str(e.get(ksh) or "")
                if "取得完了" not in v:
                    continue
                yt = _parse_jst_date(e.get(F_YUKYU_TAISHA))
                ht = _parse_jst_date(e.get(F_HAKEN_TAISHA))
                ds = [d for d in (yt, ht) if d]
                if not ds:
                    continue
                latest = max(ds)
                cands.append({"name": e.get("Title"), "shaho": v, "latest": str(latest),
                              "le_today": latest <= today, "soshitsu": ("喪失" in v)})
            dbg = {"today": str(today), "emps": len(emps), "k_shaho": ksh,
                   "n_yukyuTaisha": sum(1 for e in emps if e.get(F_YUKYU_TAISHA)),
                   "cand_count": len(cands), "cands": cands[:25]}
        return _json_response({"ok": True, "counts": c, "sent": sent, "debug": dbg})
    except Exception as e:
        logging.exception("soumu_counts_test failed")
        return _json_response({"error": "internal", "detail": str(e)}, 500)


@app.route(route="push/subscribe", methods=["POST", "OPTIONS"])
def push_subscribe(req: func.HttpRequest) -> func.HttpResponse:
    """ブラウザの Push 購読情報を本人の社員レコードに保存。
    Body: { subscription: <PushSubscription JSON> } | { unsubscribe: true }"""
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
    try:
        emp = find_active_employee_by_shain(shain_no)
        if not emp:
            return _json_response({"error": "not_active"}, 403)
        target_id = emp.get("Id")
        if body.get("unsubscribe"):
            sp_patch_item(LIST_SHAIN, int(target_id), {F_PORTAL_PUSH: ""})
            return _json_response({"ok": True, "subscribed": False})
        sub = body.get("subscription")
        if not sub or not isinstance(sub, dict) or not sub.get("endpoint"):
            return _json_response({"error": "invalid_subscription"}, 400)
        sub_str = json.dumps(sub, ensure_ascii=False)
        if len(sub_str) > 4000:
            return _json_response({"error": "subscription_too_large"}, 400)
        sp_patch_item(LIST_SHAIN, int(target_id), {F_PORTAL_PUSH: sub_str})
        return _json_response({"ok": True, "subscribed": True})
    except Exception as e:
        logging.exception("push_subscribe failed")
        return _json_response({"error": "internal", "detail": str(e)}, 500)


@app.route(route="zairyu/renew-plan", methods=["POST", "OPTIONS"])
def zairyu_renew_plan(req: func.HttpRequest) -> func.HttpResponse:
    """本人が「在留(就労ビザ)更新を いつ 申請するか」を伝える。
    Body: { date: "YYYY-MM-DD", note?: "" }
    → 社員Listに予定日/メモを記録 + 総務(期限アラートチャネル)へTeamsメールでN通知。"""
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
    date_str = (body.get("date") or "").strip()
    note = (body.get("note") or "")[:300]
    try:
        plan = _dt.date.fromisoformat(date_str)
    except Exception:
        return _json_response({"error": "invalid_date"}, 400)
    today = (_dt.datetime.utcnow() + _dt.timedelta(hours=9)).date()
    if plan < today or plan > today + _dt.timedelta(days=400):
        return _json_response({"error": "date_out_of_range"}, 400)
    try:
        emp = find_active_employee_by_shain(shain_no)
        if not emp:
            return _json_response({"error": "not_active"}, 403)
        target_id = emp.get("Id")
        sp_patch_item(LIST_SHAIN, int(target_id), {
            F_RENEW_PLAN: plan.isoformat() + "T00:00:00Z",
            F_RENEW_NOTE: note,
        })
        # 総務へ通知 (期限アラートチャネル)
        name = emp.get(F_SHAIN_NAME) or ""
        haken = strip_buka_prefix(emp.get(F_BUKA) or "")
        lines = [
            "在留カード(就労ビザ)更新の予定が登録されました。",
            f"社員番号: {shain_no}",
            f"氏名: {name}",
            f"派遣先: {haken}",
            f"申請予定日: {plan.isoformat()}",
        ]
        if note:
            lines.append(f"メモ: {note}")
        subject = f"【在留更新予定】{shain_no} {name} - {plan.isoformat()}"
        _send_notification_mail(subject, "\n".join(lines), to_addr=os.environ.get("EXPIRY_NOTIFY_EMAIL") or None)
        return _json_response({"ok": True, "date": plan.isoformat()})
    except Exception as e:
        logging.exception("zairyu_renew_plan failed")
        return _json_response({"error": "internal", "detail": str(e)}, 500)


@app.route(route="shorui/apply", methods=["POST", "OPTIONS"])
def shorui_apply(req: func.HttpRequest) -> func.HttpResponse:
    """必要書類を請求。SP List「必要書類申請」に INSERT (Teams が List 監視で通知)。
    Body: { docType: "在職証明書"|"源泉徴収票", years: ["2025","2024"], note: "" }
    """
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    shain_no = int(payload["shainNo"])
    wait = check_rate_limit(shain_no, "apply")
    if wait is not None:
        return _json_response({"error": "rate_limited", "retryAfterSeconds": wait}, 429,
                              extra_headers={"Retry-After": str(wait)})
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    doc_type = (body.get("docType") or "").strip()
    if doc_type not in SHORUI_TYPES:
        return _json_response({"error": "invalid_doc_type", "allowed": list(SHORUI_TYPES)}, 400)
    years = body.get("years") or []
    if not isinstance(years, list):
        years = []
    # 年度は数字4桁のみ許可、最大3件
    years = [str(y).strip() for y in years if re.fullmatch(r"\d{4}", str(y).strip())][:3]
    if doc_type == "源泉徴収票" and not years:
        return _json_response({"error": "year_required"}, 400)
    note = (body.get("note") or "")[:300]
    try:
        emp = find_active_employee_by_shain(shain_no)
        if not emp:
            return _json_response({"error": "not_active"}, 403)
        buka_text = emp.get(F_BUKA) or ""
        fields = {
            "Title": f"{shain_no} {doc_type}",
            SH_SHAIN_NO: str(shain_no),
            SH_SHAIN_NAME: emp.get(F_SHAIN_NAME) or "",
            SH_HAKENSAKI: strip_buka_prefix(buka_text),
            SH_SHURUI: doc_type,
            SH_NENDO: "、".join(years) if years else "",
            SH_BIKO: note,
            SH_STATUS: "pending",
        }
        new_id = sp_post_item(LIST_SHORUI, fields)
        # Teams チャネル(メールアドレス)へ通知 (本文に社員番号・氏名・内容)
        shain_name = emp.get(F_SHAIN_NAME) or ''
        lines = [
            "必要書類の請求がありました。",
            f"社員番号: {shain_no}",
            f"氏名: {shain_name}",
            f"派遣先: {strip_buka_prefix(buka_text)}",
            f"書類: {doc_type}",
        ]
        if years:
            lines.append(f"年度: {'、'.join(years)}")
        if note:
            lines.append(f"備考: {note}")
        subject = f"【必要書類請求】{shain_no} {shain_name} - {doc_type}"
        _send_notification_mail(subject, "\n".join(lines))
        return _json_response({"ok": True, "id": new_id})
    except Exception as e:
        logging.exception("shorui_apply failed")
        return _json_response({"error": "internal", "detail": str(e)}, 500)


@app.route(route="shorui/history", methods=["GET", "OPTIONS"])
def shorui_history(req: func.HttpRequest) -> func.HttpResponse:
    """本人の必要書類請求履歴。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    shain_no = payload["shainNo"]
    try:
        ent_shain = "OData_" + SH_SHAIN_NO
        items = sp_get_items(
            LIST_SHORUI,
            select=f"Id,OData_{SH_SHAIN_NO},OData_{SH_SHURUI},OData_{SH_NENDO},Status,OData_{SH_BIKO},Created",
            filter_=f"{ent_shain} eq {shain_no}",
            orderby="Id desc",
        )
        out = []
        for it in items:
            created = _utc_to_jst_date(it.get("Created"))
            out.append({
                "id": it.get("Id"),
                "docType": it.get("OData_" + SH_SHURUI),
                "years": it.get("OData_" + SH_NENDO),
                "status": it.get("Status"),
                "note": it.get("OData_" + SH_BIKO),
                "createdAt": created.isoformat() if created else None,
            })
        return _json_response({"items": out})
    except Exception as e:
        logging.exception("shorui_history failed")
        return _json_response({"error": "internal", "detail": str(e)}, 500)


# 源泉徴収票 PDF 保存フォルダ (yukyu-app と同じ。年度別サブフォルダ + ファイル名末尾=社員番号)
GENSEN_BASE_PATH = "/sites/TeamStepup/Shared Documents/社員ファイル/※※源泉　gensen"


def _match_gensen_filename(filename: str, shain_no: int) -> bool:
    """ファイル名末尾の数字グループが社員番号と一致するか (yukyu-app と同じロジック)。"""
    if not filename:
        return False
    name = re.sub(r"\.[^.]+$", "", filename)  # 拡張子除去
    emp = str(shain_no)
    emp0 = emp.zfill(5)  # 3桁→00補完 (例 760→00760)
    m = re.search(r"(\d+)\D*$", name)
    if not m:
        return emp in name or emp0 in name
    trailing = m.group(1)
    for e in (emp, emp0):
        if trailing == e or trailing.endswith(e) or (e.endswith(trailing) and len(trailing) >= 3):
            return True
    return False


def _ts_folders(server_relative_url: str) -> List[Dict[str, Any]]:
    h = {"Authorization": f"Bearer {_get_sp_token()}", "Accept": "application/json;odata=nometadata"}
    from urllib.parse import quote
    url = f"{SITE_TEAMSTEPUP}/_api/web/GetFolderByServerRelativeUrl('{quote(server_relative_url)}')/Folders?$select=Name,ServerRelativeUrl&$top=200"
    r = requests.get(url, headers=h, timeout=30)
    r.raise_for_status()
    return r.json().get("value", [])


def _ts_files(server_relative_url: str) -> List[Dict[str, Any]]:
    h = {"Authorization": f"Bearer {_get_sp_token()}", "Accept": "application/json;odata=nometadata"}
    from urllib.parse import quote
    url = f"{SITE_TEAMSTEPUP}/_api/web/GetFolderByServerRelativeUrl('{quote(server_relative_url)}')/Files?$select=Name,ServerRelativeUrl,TimeLastModified,Length&$top=2000"
    r = requests.get(url, headers=h, timeout=30)
    r.raise_for_status()
    return r.json().get("value", [])


def _gensen_collect_folders(base_url: str, max_depth: int = 3) -> List[str]:
    """gensen フォルダ配下の全フォルダ (ネスト含む) を列挙。
    R2 のように年度フォルダの中にさらにサブフォルダがあるケースに対応。"""
    result = [base_url]
    frontier = [(base_url, 0)]
    seen = {base_url}
    while frontier:
        url, depth = frontier.pop()
        if depth >= max_depth:
            continue
        try:
            subs = _ts_folders(url)
        except Exception:
            continue
        for f in subs:
            nm = f.get("Name", "")
            if not nm or nm == "Forms" or nm.startswith("_"):
                continue
            su = f.get("ServerRelativeUrl")
            if su and su not in seen:
                seen.add(su)
                result.append(su)
                frontier.append((su, depth + 1))
    return result


def _gensen_year(folder_name: str, file_name: str) -> int:
    """新しい年度を上に並べるための西暦。ファイル名の 源泉YYYY 優先、無ければフォルダ名 RN(令和)。"""
    m = re.search(r"(20\d{2})", file_name or "")
    if m:
        return int(m.group(1))
    m = re.search(r"R(\d+)", folder_name or "")
    if m:
        return 2018 + int(m.group(1))  # 令和N → 西暦
    return 0


@app.route(route="shorui/gensen-list", methods=["GET", "OPTIONS"])
def shorui_gensen_list(req: func.HttpRequest) -> func.HttpResponse:
    """本人の源泉徴収票 PDF を gensen フォルダから検索して一覧返す (総務不要)。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    shain_no = int(payload["shainNo"])
    try:
        matches: List[Dict[str, Any]] = []
        # 年度サブフォルダ (ネスト含め再帰的に列挙。R2 のように二重フォルダのケースに対応)
        targets = _gensen_collect_folders(GENSEN_BASE_PATH)
        for folder_url in targets:
            try:
                files = _ts_files(folder_url)
            except Exception:
                continue
            folder_name = folder_url.rstrip("/").split("/")[-1]
            for fl in files:
                fname = fl.get("Name", "")
                if fname.lower().endswith(".pdf") and _match_gensen_filename(fname, shain_no):
                    matches.append({
                        "folder": folder_name,
                        "fileName": fname,
                        "serverRelativeUrl": fl.get("ServerRelativeUrl"),
                        "modified": fl.get("TimeLastModified"),
                        "year": _gensen_year(folder_name, fname),
                    })
        # 新しい年度を上に (西暦 → ファイル名 で降順)
        matches.sort(key=lambda m: (m.get("year", 0), m.get("fileName", "")), reverse=True)
        return _json_response({"items": matches})
    except Exception as e:
        logging.exception("gensen-list failed")
        return _json_response({"error": "internal", "detail": str(e)}, 500)


@app.route(route="shorui/gensen-download", methods=["GET", "OPTIONS"])
def shorui_gensen_download(req: func.HttpRequest) -> func.HttpResponse:
    """本人の源泉徴収票 PDF をダウンロード (base64)。
    Query: ?url=<serverRelativeUrl>。本人の社員番号と一致するファイル名のみ許可。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    shain_no = int(payload["shainNo"])
    sru = req.params.get("url") or ""
    # 安全確認: gensen フォルダ配下 + ファイル名が本人の社員番号
    if not sru.startswith(GENSEN_BASE_PATH + "/") or ".." in sru:
        return _json_response({"error": "invalid_path"}, 400)
    fname = sru.rstrip("/").split("/")[-1]
    if not _match_gensen_filename(fname, shain_no):
        return _json_response({"error": "forbidden"}, 403)
    try:
        from urllib.parse import quote
        api = f"{SITE_TEAMSTEPUP}/_api/web/GetFileByServerRelativeUrl('{quote(sru)}')/$value"
        rb = requests.get(api, headers={"Authorization": f"Bearer {_get_sp_token()}"}, timeout=60)
        rb.raise_for_status()
        b64 = base64.b64encode(rb.content).decode("ascii")
        return _json_response({"fileName": fname, "contentType": "application/pdf", "base64": b64})
    except Exception as e:
        logging.exception("gensen-download failed")
        return _json_response({"error": "internal", "detail": str(e)}, 500)


# 会社情報 (在職証明書) — yukyu-app と同一
COMPANY_INFO = {
    "addrJp": "静岡県磐田市上本郷1006-7",
    "nameJp": "有限会社ステップ・アップ",
    "repJp": "代表取締役　山下　浩俊",
    "tel": "0538-36-6968",
}


@app.route(route="zaishoku/data", methods=["GET", "OPTIONS"])
def zaishoku_data(req: func.HttpRequest) -> func.HttpResponse:
    """在職証明書をアプリで本人発行するための確定データを返す。
    本人が編集できない確定値 (社員Listベース) を返し、フロントで PDF を生成する。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    try:
        emp = find_active_employee_by_shain(payload["shainNo"])
    except Exception as e:
        logging.exception("zaishoku_data lookup failed")
        return _json_response({"error": "lookup_failed", "detail": str(e)}, 500)
    if not emp:
        # 退社済みなどで在職中でない → 発行不可
        return _json_response({"error": "not_active"}, 403)

    def _isodate(v: Any) -> str:
        d = _utc_to_jst_date(v)
        return d.isoformat() if d else ""

    address = (emp.get(F_ADDRESS) or "").strip()
    birthday = _isodate(emp.get(F_BIRTHDAY))
    nyusha = _isodate(emp.get(F_NYUSHA))
    name = (emp.get(F_SHAIN_NAME) or "").strip()
    romaji = (emp.get(F_ZAIRYU_NAME) or "").strip()
    nationality = (emp.get(F_KOKUSEKI) or "").strip()
    # JP版で必須なのは 氏名・住所・入社日 (生年月日/国籍は無くても発行可だが揃っている方が望ましい)
    complete = bool(name and address and nyusha)
    return _json_response({
        "shainNo": emp.get(F_SHAIN_NO),
        "name": name,
        "romajiName": romaji,
        "address": address,
        "birthday": birthday,
        "nationality": nationality,
        "nyushaDate": nyusha,
        "complete": complete,
        "company": COMPANY_INFO,
    })


# ====== 日次 期限チェック → 総務へ Teams 通知 (Part B) ======
EXPIRY_WARN_DAYS = 30        # 期限の何日前から通知するか
EXPIRY_OVERDUE_GRACE = 180   # 期限切れ後この日数まで通知 (これ以上前は古い未更新データとみなし除外)


def _notify_lang(kokuseki: str) -> str:
    """通知言語: 日本人=ja / フィリピン人=en / それ以外(ブラジル含む)=pt。"""
    k = kokuseki or ""
    if "日本" in k:
        return "ja"
    if "フィリピン" in k or "philippin" in k.lower():
        return "en"
    return "pt"


# 通知対象の書類: (kind, 満了日フィールド, 条件 'foreign'|'car', アプリの提出画面)
EXPIRY_DOC_SPECS = [
    ("zairyu", F_ZAIRYU_KIGEN, "foreign"),
    ("menkyo", F_MENKYO_KIGEN, "car"),
    ("shaken", F_SHAKEN_KIGEN, "car"),
    ("jibai", F_JIBAI_KIGEN, "car"),
    ("nini", F_NINI_KIGEN, "car"),
]
# 書類名 (言語別)
_DOC_NAMES = {
    "ja": {"zairyu": "在留カード", "menkyo": "免許証", "shaken": "車検", "jibai": "自賠責保険", "nini": "任意保険"},
    "en": {"zairyu": "Residence card", "menkyo": "Driver's license", "shaken": "Vehicle inspection", "jibai": "Compulsory insurance", "nini": "Car insurance"},
    "pt": {"zairyu": "Cartão de permanência", "menkyo": "Carteira de motorista", "shaken": "Inspeção (Shaken)", "jibai": "Seguro obrigatório (Jibaiseki)", "nini": "Seguro do carro"},
}
# 通知文テンプレート (title = 送信元固定 / body = 書類名 + 期限 + 行動)
_NOTIFY_TPL = {
    "ja": {"from": "Step Up からお知らせ", "sep": "・",
           "near": "{docs}の期限まであと{n}日です。アプリから新しい書類を提出してください。",
           "exp": "{docs}の期限が切れています。アプリから新しい書類を提出してください。"},
    "en": {"from": "Notice from Step Up", "sep": ", ",
           "near": "{docs}: {n} days until expiry. Please submit the new document in the app.",
           "exp": "{docs}: expired. Please submit the new document in the app."},
    "pt": {"from": "Aviso da Step Up", "sep": ", ",
           "near": "{docs}: faltam {n} dias para o vencimento. Envie o novo documento no aplicativo.",
           "exp": "{docs}: vencido. Envie o novo documento no aplicativo."},
}


def _build_expiry_push(kinds_days: List[Any], kokuseki: str) -> Dict[str, str]:
    """本人の国籍言語で push の title/body を作る。kinds_days=[(kind,days),...]。title は送信元固定。"""
    lang = _notify_lang(kokuseki)
    names = _DOC_NAMES[lang]
    T = _NOTIFY_TPL[lang]
    kinds = [k for k, _ in kinds_days]
    min_days = min(d for _, d in kinds_days)
    docs = T["sep"].join(dict.fromkeys(names[k] for k in kinds))  # 重複除去・順序保持
    body = (T["exp"] if min_days < 0 else T["near"]).format(docs=docs, n=min_days)
    return {"title": T["from"], "body": body}


def _is_working_day(d: _dt.date) -> bool:
    """就業日か (土日でなく、会社休日カレンダーにもない)。取得失敗時は就業日扱い。"""
    if d.weekday() >= 5:  # 土(5)・日(6)
        return False
    try:
        if d in set(list_kaisha_kyujitsu_dates()):
            return False
    except Exception:
        logging.warning("会社休日の取得に失敗。就業日として続行")
    return True


@app.timer_trigger(schedule="0 0 3 * * *", arg_name="timer", run_on_startup=False)
def daily_expiry_check(timer: func.TimerRequest) -> None:
    """毎日 03:00 UTC (=12:00 JST)。就業日(土日・会社休日を除く)のみ実行。
    在留カード(外国籍)/免許証・車検・自賠責・任意保険(車通勤) の期限が 30日前〜切れ後180日 の
    在職者を集計し、本人へ Web Push + 総務チャネルへ通知 (本人が提出すると翌日に外れる)。"""
    try:
        today = (_dt.datetime.utcnow() + _dt.timedelta(hours=9)).date()
        if not _is_working_day(today):
            logging.info(f"daily_expiry_check: {today} は就業日でないためスキップ")
            return
        items = sp_get_items(
            LIST_SHAIN,
            select=",".join(["Id", F_SHAIN_NO, F_SHAIN_NAME, F_BUKA, F_KOKUSEKI,
                             F_TAISHA_DATE, F_TSUKIN_OLD, F_TSUKIN_NEW, F_PORTAL_PUSH, F_RENEW_PLAN,
                             F_ZAIRYU_KIGEN, F_MENKYO_KIGEN, F_SHAKEN_KIGEN, F_JIBAI_KIGEN, F_NINI_KIGEN]),
            top=6000, orderby="Id desc",
        )
        rows: Dict[str, List[Any]] = {k: [] for k, _, _ in EXPIRY_DOC_SPECS}
        seen: set = set()  # (社員番号, kind)
        push_targets: Dict[Any, Dict[str, Any]] = {}
        zairyu_plans: Dict[Any, _dt.date] = {}  # 社員番号 → 在留更新の申請予定日 (未来)
        for it in items:
            no = it.get(F_SHAIN_NO)
            taisha = _utc_to_jst_date(it.get(F_TAISHA_DATE))
            if taisha is not None and taisha < today:
                continue  # 退社済み
            name = it.get(F_SHAIN_NAME) or ""
            kokuseki = it.get(F_KOKUSEKI) or ""
            haken = strip_buka_prefix(it.get(F_BUKA) or "")
            sub_json = (it.get(F_PORTAL_PUSH) or "").strip()
            emp_id = it.get("Id")
            is_foreign = "日本" not in kokuseki
            is_car = commutes_by_car(it)
            renew_plan = _utc_to_jst_date(it.get(F_RENEW_PLAN))
            plan_active = bool(renew_plan and renew_plan >= today)  # 予定日が未来 → 在留の催促を止める
            for kind, field, cond in EXPIRY_DOC_SPECS:
                if (no, kind) in seen:
                    continue
                if not (is_foreign if cond == "foreign" else is_car):
                    continue
                kig = _utc_to_jst_date(it.get(field))
                if not kig:
                    continue
                days = (kig - today).days
                if -EXPIRY_OVERDUE_GRACE <= days <= EXPIRY_WARN_DAYS:
                    rows[kind].append((days, no, name, kig, haken))
                    seen.add((no, kind))
                    if kind == "zairyu" and plan_active:
                        zairyu_plans[no] = renew_plan  # 総務一覧に「更新予定」併記、本人pushは抑制
                    elif sub_json:
                        t = push_targets.setdefault(no, {"sub": sub_json, "empId": emp_id, "kinds": [], "kokuseki": kokuseki})
                        t["kinds"].append((kind, days))

        total = sum(len(v) for v in rows.values())
        if total == 0:
            logging.info("daily_expiry_check: 対象なし")
            return

        def fmt_rows(rs: List[Any], plans: Optional[Dict[Any, _dt.date]] = None) -> str:
            out = []
            for days, no, name, kig, haken in sorted(rs):
                tag = "【期限切れ】" if days < 0 else (f"あと{days}日")
                extra = ""
                if plans and no in plans:
                    extra = f" → 更新予定: {plans[no].isoformat()}"
                out.append(f"  ・{no} {name}（{haken}） 期限{kig.isoformat()} {tag}{extra}")
            return "\n".join(out)

        lines = [f"📋 期限チェック（{today.isoformat()}）", ""]
        for kind, _, _ in EXPIRY_DOC_SPECS:
            if rows[kind]:
                lines.append(f"■ {_DOC_NAMES['ja'][kind]} 期限間近/切れ（{len(rows[kind])}名）")
                lines.append(fmt_rows(rows[kind], zairyu_plans if kind == "zairyu" else None))
                lines.append("")
        lines.append("※ 本人がポータルで新しい書類を提出すると期限が更新され、翌日以降このリストから外れます。")
        text = "\n".join(lines)
        subject = f"【期限アラート】{total}件 ({today.isoformat()})"
        _send_notification_mail(subject, text, to_addr=os.environ.get("EXPIRY_NOTIFY_EMAIL") or None)
        logging.info("daily_expiry_check sent: " + " ".join(f"{k}={len(rows[k])}" for k, _, _ in EXPIRY_DOC_SPECS))

        # 本人へ Web Push (購読済みのみ・提出するまで毎就業日)
        pushed = gone = 0
        for no, tgt in push_targets.items():
            msg = _build_expiry_push(tgt["kinds"], tgt.get("kokuseki", ""))
            res = _send_web_push(tgt["sub"], {"title": msg["title"], "body": msg["body"],
                                              "url": "/", "badge": len(tgt["kinds"])})
            if res == "ok":
                pushed += 1
            elif res == "gone":
                gone += 1
                try:
                    sp_patch_item(LIST_SHAIN, int(tgt["empId"]), {F_PORTAL_PUSH: ""})
                except Exception:
                    pass
        if push_targets:
            logging.info(f"daily_expiry_check web push: ok={pushed} gone={gone} total={len(push_targets)}")
    except Exception as e:
        logging.exception(f"daily_expiry_check failed: {e}")


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


# ============================================================
# yukyu-app (総務担当者) 向け 社員ファイル AI-OCR 中継
#   認証: 呼び出し元の SharePoint MSAL アクセストークンを検証
#         (SP /_api/web/currentuser に転送し、メールを許可リストと照合)
#   入力: { docType, images: [base64...] } または { docType, paths: [社員ファイル配下のServerRelativeUrl...] }
#   出力: { ocr: {フィールド…}, files: [読み取ったパス…] }
# ============================================================
YUKYU_STAFF_ALLOWED = {
    "h.yamashita@team-stepup.com", "a.yamashita@team-stepup.com",
    "kaori.j@team-stepup.com", "m.mori@team-stepup.com",
    "m.yoshiura@team-stepup.com", "y.mori@team-stepup.com",
}


def require_staff_auth(req: func.HttpRequest):
    """yukyu-app 利用者 (役員/担当者) の認証。呼び出し元の SP トークンで本人確認。"""
    auth = req.headers.get("Authorization", "")
    if not auth.startswith("Bearer "):
        return None, _json_response({"error": "unauthorized"}, 401)
    caller_token = auth[7:].strip()
    try:
        r = requests.get(
            f"{SP_RESOURCE}/_api/web/currentuser",
            headers={"Authorization": f"Bearer {caller_token}",
                     "Accept": "application/json;odata=nometadata"},
            timeout=15,
        )
        if r.status_code != 200:
            return None, _json_response({"error": "unauthorized"}, 401)
        info = r.json()
        email = str(info.get("Email") or info.get("UserPrincipalName") or "").strip().lower()
    except Exception:
        logging.exception("staff auth: SP currentuser failed")
        return None, _json_response({"error": "unauthorized"}, 401)
    if email not in YUKYU_STAFF_ALLOWED:
        logging.warning("staff auth: not allowed: %s", email)
        return None, _json_response({"error": "forbidden"}, 403)
    return email, None


@app.route(route="yukyu/file-ocr", methods=["POST", "OPTIONS"])
def yukyu_file_ocr(req: func.HttpRequest) -> func.HttpResponse:
    """社員ファイル内の書類画像 (またはアップロード前のbase64画像) を GPT-4o-mini で構造化抽出。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    email, err = require_staff_auth(req)
    if err:
        return err
    # レート制限 (メールをキー化して既存プールを流用)
    rl_key = abs(hash(email)) % 1000000
    wait = check_rate_limit(rl_key, "yukyu_ocr")
    if wait is not None:
        return _json_response({"error": "rate_limited", "retryAfterSeconds": wait},
                              429, extra_headers={"Retry-After": str(wait)})
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)

    doc_type = str(body.get("docType") or "").strip()
    import gpt_ocr as _g
    if doc_type != "zairyu" and doc_type not in _g.DOC_PROMPTS:
        return _json_response({"error": "unknown_doc_type", "docType": doc_type}, 400)

    images_bytes = []
    used_files = []
    # a) base64 直接 (登録フォームの撮影直後)
    for b64s in (body.get("images") or [])[:4]:
        try:
            raw = base64.b64decode(_strip_data_url(b64s))
        except Exception:
            return _json_response({"error": "invalid_base64"}, 400)
        if len(raw) > 10 * 1024 * 1024:
            return _json_response({"error": "image_too_large", "maxMB": 10}, 400)
        images_bytes.append(raw)
    # b) 社員ファイル配下のパス (一括読み取り)
    from urllib.parse import quote as _q
    for sru in (body.get("paths") or [])[:4]:
        sru = str(sru or "")
        if not _validate_shainfile_path(sru):
            return _json_response({"error": "invalid_path", "path": sru}, 400)
        try:
            api = f"{SITE_TEAMSTEPUP}/_api/web/GetFileByServerRelativeUrl('{_q(sru)}')/$value"
            rb = requests.get(api, headers={"Authorization": f"Bearer {_get_sp_token()}"}, timeout=60)
            rb.raise_for_status()
        except Exception as e:
            logging.exception("file fetch failed: %s", sru)
            return _json_response({"error": "file_fetch_failed", "path": sru, "detail": str(e)}, 502)
        if len(rb.content) > 10 * 1024 * 1024:
            return _json_response({"error": "image_too_large", "path": sru, "maxMB": 10}, 400)
        images_bytes.append(rb.content)
        used_files.append(sru)

    if not images_bytes:
        return _json_response({"error": "no_images"}, 400)

    # モデル指定(任意): 既定 gpt-4o-mini。gpt-4o / claude(=Claudeビジョン, 要ANTHROPIC_API_KEY) を選べる
    req_model = str(body.get("model") or "").strip()
    if req_model in ("gpt-4o-mini", "gpt-4o") or req_model.startswith("claude"):
        use_model = req_model
    else:
        use_model = None
    try:
        if doc_type == "zairyu":
            result = _g.extract_zairyu_fields(
                images_bytes[0], images_bytes[1] if len(images_bytes) > 1 else None, model=use_model)
        else:
            result = _g.extract_doc_fields(doc_type, images_bytes, model=use_model)
    except Exception as e:
        logging.exception("yukyu file-ocr failed")
        return _json_response({"error": "ocr_failed", "detail": str(e)}, 500)

    logging.info("yukyu file-ocr by %s docType=%s files=%d model=%s", email, doc_type, len(images_bytes), use_model)
    # 一時デバッグ: OCRが実際に抽出した内容をログ出力(原因調査用・内部ログのみ)
    try:
        logging.info("yukyu file-ocr RESULT docType=%s model=%s -> %s",
                     doc_type, use_model, json.dumps(result, ensure_ascii=False))
    except Exception:
        logging.info("yukyu file-ocr RESULT docType=%s -> %r", doc_type, result)
    return _json_response({"ocr": result, "files": used_files, "docType": doc_type})


# ====== 送迎連絡: 当日の帰り便を派遣先別に従業員へ配信 (Push + ホームバナー) ======
_SOUGEI_TPL = {
    "ja": {"title": "🚐 Step Up からお知らせ",
           "tv": "本日の帰り：終業 {time} ／ 車両 {vehicle}",
           "t": "本日の帰り：終業 {time}",
           "v": "本日の帰り：車両 {vehicle}"},
    "en": {"title": "🚐 Notice from Step Up",
           "tv": "Ride home today: ends {time} / vehicle {vehicle}",
           "t": "Ride home today: ends {time}",
           "v": "Ride home today: vehicle {vehicle}"},
    "pt": {"title": "🚐 Aviso da Step Up",
           "tv": "Volta de hoje: término {time} / veículo {vehicle}",
           "t": "Volta de hoje: término {time}",
           "v": "Volta de hoje: veículo {vehicle}"},
}


def _build_sougei_push(end_time: str, vehicle: str, kokuseki: str) -> Dict[str, str]:
    T = _SOUGEI_TPL[_notify_lang(kokuseki)]
    if end_time and vehicle:
        body = T["tv"].format(time=end_time, vehicle=vehicle)
    elif end_time:
        body = T["t"].format(time=end_time)
    elif vehicle:
        body = T["v"].format(vehicle=vehicle)
    else:
        body = T["title"]
    return {"title": T["title"], "body": body}


# 運転手向けマニフェスト(乗車者一覧) テンプレート
_SOUGEI_DRV_TPL = {
    "ja": {"title": "🚐 Step Up 本日の送迎（運転）", "head": "本日のあなたの便：",
           "empty": "本日の送迎担当はなくなりました（変更）", "cnt": "（計{n}名）", "sep": " ／ ", "nm": "・"},
    "en": {"title": "🚐 Step Up — your ride today", "head": "Your passengers today: ",
           "empty": "No passengers assigned today (updated)", "cnt": " ({n})", "sep": " / ", "nm": ", "},
    "pt": {"title": "🚐 Step Up — sua condução hoje", "head": "Seus passageiros hoje: ",
           "empty": "Sem passageiros hoje (atualizado)", "cnt": " ({n})", "sep": " / ", "nm": ", "},
}


def _build_driver_manifest_push(riders: List[Dict[str, str]], kokuseki: str) -> Dict[str, str]:
    """運転手へ送る乗車者マニフェスト。riders=[{name,endTime}]。時間別にまとめる。"""
    T = _SOUGEI_DRV_TPL[_notify_lang(kokuseki)]
    riders = [r for r in (riders or []) if (str(r.get("name") or "").strip() or str(r.get("endTime") or "").strip())]
    if not riders:
        return {"title": T["title"], "body": T["empty"]}
    by_time: Dict[str, List[str]] = {}
    for r in riders:
        t = str(r.get("endTime") or "-").strip() or "-"
        by_time.setdefault(t, []).append(str(r.get("name") or "").strip())
    parts = []
    for t in sorted(by_time.keys()):
        names = T["nm"].join([n for n in by_time[t] if n])
        parts.append(f"{t} {names}".strip())
    body = T["head"] + T["sep"].join(parts) + T["cnt"].format(n=len(riders))
    return {"title": T["title"], "body": body}


def _push_sougei(hakensaki: str, end_time: str, vehicle: str) -> Tuple[int, int]:
    """該当派遣先の送迎社員へ Web Push。戻り (pushed, targets)。"""
    norm = strip_buka_prefix(hakensaki)
    today = _today_jst()
    try:
        emps = sp_get_items(
            LIST_SHAIN,
            select=",".join(["Id", F_SHAIN_NO, F_BUKA, F_KOKUSEKI,
                             F_TSUKIN_OLD, F_TSUKIN_NEW, F_TAISHA_DATE, F_PORTAL_PUSH]),
            top=10000,
        )
    except Exception:
        logging.exception("送迎社員の取得に失敗")
        return 0, 0
    pushed = 0
    targets = 0
    for e in emps:
        taisha = _utc_to_jst_date(e.get(F_TAISHA_DATE))
        if taisha is not None and taisha < today:
            continue
        if strip_buka_prefix(e.get(F_BUKA) or "") != norm:
            continue
        if not commutes_by_sougei(e):
            continue
        targets += 1
        sub = (e.get(F_PORTAL_PUSH) or "").strip()
        if not sub:
            continue
        msg = _build_sougei_push(end_time, vehicle, e.get(F_KOKUSEKI) or "")
        res = _send_web_push(sub, {"title": msg["title"], "body": msg["body"],
                                   "url": "/", "tag": "sougei-notice", "badge": 1})
        if res == "ok":
            pushed += 1
        elif res == "gone":
            try:
                sp_patch_item(LIST_SHAIN, int(e["Id"]), {F_PORTAL_PUSH: ""})
            except Exception:
                pass
    return pushed, targets


@app.route(route="yukyu/sougei-send", methods=["POST", "OPTIONS"])
def yukyu_sougei_send(req: func.HttpRequest) -> func.HttpResponse:
    """送迎連絡: 派遣先別に当日の帰り便/終業時間を保存し、送迎社員へ Push。
    Body: { hakensaki, endTime, vehicle, memo?, date? }。staff 認証必須。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    email, err = require_staff_auth(req)
    if err:
        return err
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    hakensaki = strip_buka_prefix(str(body.get("hakensaki") or "").strip())
    end_time = str(body.get("endTime") or "").strip()
    vehicle = str(body.get("vehicle") or "").strip()
    memo = str(body.get("memo") or "").strip()
    if not hakensaki:
        return _json_response({"error": "missing_hakensaki"}, 400)
    if not end_time and not vehicle:
        return _json_response({"error": "empty_message"}, 400)
    date_str = str(body.get("date") or "").strip()
    today = _today_jst()
    try:
        target = _dt.date.fromisoformat(date_str) if date_str else today
    except Exception:
        target = today
    # 1) 恒久記録 (毎回 INSERT。受信側/プリフィルは最新 Id を採用)
    try:
        rec_id = sp_post_item(LIST_SOUGEI, {
            "Title": f"{target.isoformat()} {hakensaki}",
            SG_DATE: target.strftime("%Y/%m/%d"),
            SG_HAKENSAKI: hakensaki,
            SG_ENDTIME: end_time,
            SG_VEHICLE: vehicle,
            SG_MEMO: memo,
            SG_SENTBY: email,
        })
    except Exception as e:
        logging.exception("送迎連絡の保存に失敗")
        return _json_response({"error": "save_failed", "detail": str(e)}, 500)
    # 2) 送迎社員へ Push (当日のみ)
    pushed, targets = (0, 0)
    if target == today:
        pushed, targets = _push_sougei(hakensaki, end_time, vehicle)
    logging.info("送迎連絡 by %s 派遣先=%s 終業=%s 車両=%s targets=%d pushed=%d",
                 email, hakensaki, end_time, vehicle, targets, pushed)
    return _json_response({"ok": True, "recordId": rec_id, "hakensaki": hakensaki,
                           "date": target.isoformat(), "targets": targets, "pushed": pushed})


# ===== 連絡事項: yukyu-appから派遣社員ポータルへ通知＋お知らせ履歴 =====
LIST_RENRAKU = "05546b2d-4af0-421b-89f9-51fa22dae2e6"  # 連絡事項 (1社員1レコード・既読管理)


@app.route(route="yukyu/renraku-send", methods=["POST", "OPTIONS"])
def yukyu_renraku_send(req: func.HttpRequest) -> func.HttpResponse:
    """連絡事項を選択した派遣社員へ配信(お知らせ記録＋ロック画面Push)。staff認証必須。
    Body: { title, body, targets:[{shainNo,name?,hakensaki?}] }"""
    import uuid
    pf = _handle_preflight(req)
    if pf:
        return pf
    email, err = require_staff_auth(req)
    if err:
        return err
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    title = str(body.get("title") or "").strip()
    honbun = str(body.get("body") or "").strip()
    file_name = str(body.get("fileName") or "").strip()[:255]
    file_url = str(body.get("fileUrl") or "").strip()
    trans = body.get("trans")
    trans_str = ""
    if trans:
        try:
            trans_str = json.dumps(trans, ensure_ascii=False)[:30000]
        except Exception:
            trans_str = ""
    raw_targets = body.get("targets") or []
    if not title and not honbun and not file_url:
        return _json_response({"error": "empty_message"}, 400)
    targets = []
    for t in raw_targets:
        if isinstance(t, dict):
            sn = str(t.get("shainNo") or "").strip()
            info = t
        else:
            sn = str(t or "").strip()
            info = {}
        if sn:
            targets.append((sn, info))
    if not targets:
        return _json_response({"error": "no_targets"}, 400)
    # 社員の push購読＋氏名をまとめて取得
    emp_map = {}
    try:
        emps = sp_get_items(LIST_SHAIN, select=",".join(["Id", F_SHAIN_NO, "Title", F_BUKA, F_PORTAL_PUSH]), top=10000)
        for e in emps:
            sn = _norm_bango(e.get(F_SHAIN_NO))
            if sn and sn not in emp_map:
                emp_map[sn] = e
    except Exception:
        logging.exception("renraku: 社員取得失敗")
    send_id = uuid.uuid4().hex[:12]
    created = 0
    pushed = 0
    for sn, info in targets:
        emp = emp_map.get(_norm_bango(sn))
        name = (str(info.get("name") or "").strip() or (emp.get("Title") if emp else "") or "")
        hakensaki = strip_buka_prefix(str(info.get("hakensaki") or "").strip() or (emp.get(F_BUKA) if emp else "") or "")
        try:
            sp_post_item(LIST_RENRAKU, {
                "Title": title or "連絡",
                "Honbun": honbun,
                "ShainNo": str(sn),
                "EmpName": name,
                "Hakensaki": hakensaki,
                "Sender": email,
                "SendId": send_id,
                "IsRead": False,
                "FileName": file_name,
                "FileUrl": file_url,
                "Trans": trans_str,
            })
            created += 1
        except Exception:
            logging.exception("renraku: record作成失敗 sn=%s", sn)
            continue
        sub = (emp.get(F_PORTAL_PUSH) or "").strip() if emp else ""
        if sub:
            r = _send_web_push(sub, {"title": "📣 " + (title or "連絡事項"),
                                     "body": (honbun[:120] or title), "url": "/",
                                     "tag": "renraku-" + send_id, "badge": 1})
            if r == "ok":
                pushed += 1
            elif r == "gone":
                try:
                    sp_patch_item(LIST_SHAIN, int(emp["Id"]), {F_PORTAL_PUSH: ""})
                except Exception:
                    pass
    logging.info("連絡事項 by %s title=%s targets=%d created=%d pushed=%d", email, title, len(targets), created, pushed)
    return _json_response({"ok": True, "sendId": send_id, "created": created, "pushed": pushed, "targets": len(targets)})


@app.route(route="renraku/list", methods=["GET", "POST", "OPTIONS"])
def renraku_list(req: func.HttpRequest) -> func.HttpResponse:
    """本人(派遣社員)宛の連絡事項(お知らせ)一覧を返す。本人JWT必須。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    sn = str(payload["shainNo"])
    try:
        recs = sp_get_items(LIST_RENRAKU, select="Id,Title,Honbun,Sender,IsRead,Created,FileName,FileUrl,Trans",
                            filter_=f"ShainNo eq '{sn}'", top=100, orderby="Id desc")
        items = []
        for r in recs:
            tr = None
            ts = (r.get("Trans") or "").strip()
            if ts:
                try:
                    tr = json.loads(ts)
                except Exception:
                    tr = None
            items.append({"id": r.get("Id"), "title": r.get("Title"), "body": r.get("Honbun"),
                          "isRead": bool(r.get("IsRead")), "created": r.get("Created"),
                          "fileName": r.get("FileName") or "", "hasFile": bool((r.get("FileUrl") or "").strip()),
                          "trans": tr})
        unread = sum(1 for it in items if not it["isRead"])
        return _json_response({"ok": True, "items": items, "unread": unread})
    except Exception as e:
        logging.exception("renraku_list failed")
        return _json_response({"error": "internal", "detail": str(e)}, 500)


@app.route(route="renraku/read", methods=["POST", "OPTIONS"])
def renraku_read(req: func.HttpRequest) -> func.HttpResponse:
    """本人宛の未読連絡を既読にする。本人JWT必須。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    sn = str(payload["shainNo"])
    try:
        recs = sp_get_items(LIST_RENRAKU, select="Id", filter_=f"ShainNo eq '{sn}' and IsRead eq 0", top=300)
        n = 0
        for r in recs:
            try:
                sp_patch_item(LIST_RENRAKU, int(r["Id"]), {"IsRead": True})
                n += 1
            except Exception:
                pass
        return _json_response({"ok": True, "marked": n})
    except Exception as e:
        logging.exception("renraku_read failed")
        return _json_response({"error": "internal", "detail": str(e)}, 500)


@app.route(route="renraku/file", methods=["GET", "OPTIONS"])
def renraku_file(req: func.HttpRequest) -> func.HttpResponse:
    """本人宛の連絡に添付されたファイルを取得(ダウンロード)。本人JWT必須。?id=レコードId"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    payload, err = require_auth(req)
    if err:
        return err
    sn = str(payload["shainNo"])
    rid = req.params.get("id")
    if not rid or not str(rid).isdigit():
        return _json_response({"error": "missing_id"}, 400)
    try:
        recs = sp_get_items(LIST_RENRAKU, select="Id,ShainNo,FileUrl,FileName", filter_=f"Id eq {int(rid)}", top=1)
        if not recs:
            return _json_response({"error": "not_found"}, 404)
        r = recs[0]
        if str(r.get("ShainNo")) != sn:                 # 本人宛の連絡のみ
            return _json_response({"error": "forbidden"}, 403)
        url = (r.get("FileUrl") or "").strip()
        if not url:
            return _json_response({"error": "no_file"}, 404)
        from urllib.parse import quote as _q
        api = f"{SITE_URL}/_api/web/GetFileByServerRelativeUrl('{_q(url)}')/$value"
        fr = requests.get(api, headers={"Authorization": f"Bearer {_get_sp_token()}"}, timeout=90)
        fr.raise_for_status()
        fname = r.get("FileName") or "file"
        return func.HttpResponse(fr.content, status_code=200, headers={
            "Content-Type": "application/octet-stream",
            "Content-Disposition": "attachment; filename*=UTF-8''" + _q(fname),
            "Access-Control-Allow-Origin": "*",
        })
    except Exception as e:
        logging.exception("renraku_file failed")
        return _json_response({"error": "internal", "detail": str(e)}, 500)


@app.route(route="yukyu/renraku-sent", methods=["GET", "POST", "OPTIONS"])
def yukyu_renraku_sent(req: func.HttpRequest) -> func.HttpResponse:
    """送信履歴＋既読確認: 過去の連絡を一斉送信(SendId)単位でまとめ、既読数・宛先別既読を返す。staff認証。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    email, err = require_staff_auth(req)
    if err:
        return err
    try:
        recs = sp_get_items(LIST_RENRAKU, select="Id,Title,Honbun,Sender,SendId,EmpName,ShainNo,Hakensaki,IsRead,Created",
                            top=5000, orderby="Id desc")
        groups = {}
        order = []
        for r in recs:
            sid = r.get("SendId") or ("_" + str(r.get("Id")))
            g = groups.get(sid)
            if g is None:
                g = {"sendId": sid, "title": r.get("Title"), "body": r.get("Honbun"),
                     "sender": r.get("Sender"), "created": r.get("Created"),
                     "recipients": [], "total": 0, "read": 0}
                groups[sid] = g
                order.append(sid)
            g["total"] += 1
            isread = bool(r.get("IsRead"))
            if isread:
                g["read"] += 1
            g["recipients"].append({"name": r.get("EmpName") or r.get("ShainNo"),
                                    "shainNo": r.get("ShainNo"), "hakensaki": r.get("Hakensaki"),
                                    "isRead": isread})
        items = [groups[s] for s in order][:200]
        return _json_response({"ok": True, "items": items})
    except Exception as e:
        logging.exception("renraku_sent failed")
        return _json_response({"error": "internal", "detail": str(e)}, 500)


@app.route(route="yukyu/translate", methods=["POST", "OPTIONS"])
def yukyu_translate(req: func.HttpRequest) -> func.HttpResponse:
    """連絡事項の件名・本文を やさしい日本語(ふりがな)/英語/ポルトガル語 に翻訳(送信前プレビュー用)。staff認証。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    email, err = require_staff_auth(req)
    if err:
        return err
    try:
        body = req.get_json()
    except Exception:
        body = {}
    title = str((body or {}).get("title") or "").strip()
    text = str((body or {}).get("body") or "").strip()
    if not title and not text:
        return _json_response({"error": "empty"}, 400)
    try:
        import gpt_ocr
        out = gpt_ocr.translate_renraku(title, text)
        return _json_response({"ok": True, "trans": out})
    except Exception as e:
        logging.exception("yukyu_translate failed")
        return _json_response({"error": "translate_failed", "detail": str(e)}, 500)


@app.route(route="yukyu/sougei-send-batch", methods=["POST", "OPTIONS"])
def yukyu_sougei_send_batch(req: func.HttpRequest) -> func.HttpResponse:
    """送迎連絡(人単位): assignments[{shainNo,hakensaki,endTime,vehicle}] を保存し本人へ Push。"""
    pf = _handle_preflight(req)
    if pf:
        return pf
    email, err = require_staff_auth(req)
    if err:
        return err
    try:
        body = req.get_json()
    except Exception:
        return _json_response({"error": "invalid_json"}, 400)
    assignments = body.get("assignments") or []
    if not isinstance(assignments, list) or not assignments:
        return _json_response({"error": "no_assignments"}, 400)
    date_str = str(body.get("date") or "").strip()
    today = _today_jst()
    try:
        target = _dt.date.fromisoformat(date_str) if date_str else today
    except Exception:
        target = today
    # 本人の push購読/国籍を引くため在職社員を1回取得 → 社員番号→最新レコード
    by_no: Dict[str, Dict[str, Any]] = {}
    try:
        emps = sp_get_items(
            LIST_SHAIN,
            select=",".join(["Id", F_SHAIN_NO, F_KOKUSEKI, F_TAISHA_DATE, F_PORTAL_PUSH]),
            top=10000,
        )
        for e in emps:
            taisha = _utc_to_jst_date(e.get(F_TAISHA_DATE))
            if taisha is not None and taisha < today:
                continue
            sn = _norm_shain(e.get(F_SHAIN_NO))
            if sn and sn not in by_no:
                by_no[sn] = e
    except Exception:
        logging.exception("社員一覧の取得に失敗")
    # 当日の前回送信レコードを取得 → 差分送信(終業 or 運転手が変わった人だけ通知)
    prev_today: Dict[str, Dict[str, str]] = {}
    if target == today:
        try:
            recs = sp_get_items(
                LIST_SOUGEI,
                select=",".join(["Id", SG_DATE, SG_SHAINNO, SG_ENDTIME, SG_VEHICLE]),
                orderby="Id desc", top=1000,
            )
            for it in recs:
                if _utc_to_jst_date(it.get(SG_DATE)) != today:
                    continue
                sn0 = _norm_shain(it.get(SG_SHAINNO))
                if sn0 and sn0 not in prev_today:
                    prev_today[sn0] = {"endTime": it.get(SG_ENDTIME) or "", "vehicle": it.get(SG_VEHICLE) or ""}
        except Exception:
            logging.warning("送迎連絡の前回記録取得に失敗")
    # 運転手通知: vehicleラベル→運転手社員番号 / 現在の便(運転手別 乗車者一覧)
    driver_by_vehicle = body.get("driverByVehicle") or {}
    cur_by_driver: Dict[str, List[Dict[str, str]]] = {}
    for a in assignments:
        dsn = _norm_shain(a.get("driverShainNo"))
        if not dsn:
            continue
        cur_by_driver.setdefault(dsn, []).append(
            {"name": str(a.get("name") or ""), "endTime": str(a.get("endTime") or "")})
    affected_drivers: set = set()
    saved = 0
    pushed = 0
    unchanged = 0
    errors = 0
    for a in assignments:
        try:
            sn = _norm_shain(a.get("shainNo"))
            et = str(a.get("endTime") or "").strip()
            veh = str(a.get("vehicle") or "").strip()
            hk = strip_buka_prefix(str(a.get("hakensaki") or "").strip())
            if not sn or (not et and not veh):
                continue
            # 差分: 当日の前回記録と同じ(終業+運転手)なら再通知しない
            prev = prev_today.get(sn)
            if prev is not None and prev.get("endTime") == et and prev.get("vehicle") == veh:
                unchanged += 1
                continue
            sp_post_item(LIST_SOUGEI, {
                "Title": f"{target.isoformat()} {sn}",
                SG_DATE: target.strftime("%Y/%m/%d"),
                SG_SHAINNO: sn,
                SG_HAKENSAKI: hk,
                SG_ENDTIME: et,
                SG_VEHICLE: veh,
                SG_MEMO: "",
                SG_SENTBY: email,
            })
            saved += 1
            # 変わった人の「新しい運転手」と「前の運転手」を通知対象に
            dsn_new = _norm_shain(a.get("driverShainNo"))
            if dsn_new:
                affected_drivers.add(dsn_new)
            if prev is not None:
                old_dsn = _norm_shain(driver_by_vehicle.get(prev.get("vehicle")))
                if old_dsn:
                    affected_drivers.add(old_dsn)
            if target == today:
                emp = by_no.get(sn)
                sub = (emp.get(F_PORTAL_PUSH) or "").strip() if emp else ""
                if sub:
                    msg = _build_sougei_push(et, veh, emp.get(F_KOKUSEKI) or "")
                    res = _send_web_push(sub, {"title": msg["title"], "body": msg["body"],
                                               "url": "/", "tag": "sougei-notice", "badge": 1})
                    if res == "ok":
                        pushed += 1
                    elif res == "gone":
                        try:
                            sp_patch_item(LIST_SHAIN, int(emp["Id"]), {F_PORTAL_PUSH: ""})
                        except Exception:
                            pass
        except Exception:
            errors += 1
            logging.exception("送迎連絡(人単位)の1件処理に失敗")
    # 運転手へ「本日のあなたの便(乗車者一覧)」を通知 (影響のあった運転手のみ・当日のみ)
    driver_targets = 0
    driver_pushed = 0
    if target == today and affected_drivers:
        for dsn in affected_drivers:
            try:
                emp = by_no.get(dsn)
                if not emp:
                    continue   # 運転手が社員Listに無い/退社 → 通知不可
                driver_targets += 1
                sub = (emp.get(F_PORTAL_PUSH) or "").strip()
                if not sub:
                    continue   # 通知OFF(未購読) → 届かない
                msg = _build_driver_manifest_push(cur_by_driver.get(dsn, []), emp.get(F_KOKUSEKI) or "")
                res = _send_web_push(sub, {"title": msg["title"], "body": msg["body"],
                                           "url": "/", "tag": "sougei-driver", "badge": 1})
                if res == "ok":
                    driver_pushed += 1
                elif res == "gone":
                    try:
                        sp_patch_item(LIST_SHAIN, int(emp["Id"]), {F_PORTAL_PUSH: ""})
                    except Exception:
                        pass
            except Exception:
                logging.exception("送迎連絡 運転手通知の1件に失敗")
    logging.info("送迎連絡batch by %s saved=%d pushed=%d unchanged=%d errors=%d drvTargets=%d drvPushed=%d",
                 email, saved, pushed, unchanged, errors, driver_targets, driver_pushed)
    return _json_response({"ok": True, "saved": saved, "pushed": pushed, "unchanged": unchanged,
                           "errors": errors, "date": target.isoformat(),
                           "driverTargets": driver_targets, "driverPushed": driver_pushed})
