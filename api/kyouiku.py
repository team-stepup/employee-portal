# -*- coding: utf-8 -*-
"""入社時教育リンク (kyouiku) — 2026-09-25

yukyu-app の「新規社員登録」画面で、国籍・通勤方法を選んだ時点で「🎓 入社時教育リンク」を発行する。
担当者が契約書を準備している間に、本人が自分のスマホ(QRを読む)で入社時資料(就業規則・社内ルール・安全教育動画)を見る。
ログイン不要の使い捨てリンク(当日23:59 JST まで)。資料は SharePoint の総務部「入社資料」フォルダから配信する。

保存場所: /sites/TeamStepup/Shared Documents/入社時教育/{社員番号 or 0}__{token}.json
  = 発行メタ(氏名・言語・通勤区分・対象資料・視聴/完了の記録)
資料一覧: 入社資料/_サイト作成/kyouiku.json (総務が編集可・5分キャッシュ)
"""
import datetime as _dt
import hashlib
import json
import logging
import os
import re
import secrets
import time
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import quote

import requests
import azure.functions as func

KY_FOLDER = "/sites/TeamStepup/Shared Documents/入社時教育"
SRC_ROOT = "/sites/TeamStepup/Shared Documents/💻総務部‐soumu-/入社資料"
CURRICULUM = SRC_ROOT + "/_サイト作成/kyouiku.json"
TOKEN_RE = re.compile(r"^[A-Za-z0-9_-]{20,64}$")
KEY_RE = re.compile(r"^[a-z0-9_]{1,40}$")
CACHE_DIR = "/tmp/kyouiku_cache"
FILE_CACHE_SEC = 30 * 60
CUR_CACHE_SEC = 5 * 60
REC_CACHE_SEC = 60
LANGS = ("ja", "pt", "en")
MIME = {".pdf": "application/pdf", ".mp4": "video/mp4"}

_cur_cache: Dict[str, Any] = {"t": 0, "v": None}
_rec_cache: Dict[str, Tuple[float, Dict[str, Any]]] = {}


def _fa():
    import function_app as fa  # 遅延 import (循環回避)
    return fa


def _now() -> _dt.datetime:
    return _dt.datetime.now(_dt.timezone.utc)


def _iso(d: _dt.datetime) -> str:
    return d.strftime("%Y-%m-%dT%H:%M:%SZ")


def _parse(s: str) -> Optional[_dt.datetime]:
    try:
        return _dt.datetime.strptime(s, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=_dt.timezone.utc)
    except Exception:
        return None


def _end_of_day_jst(now: _dt.datetime) -> _dt.datetime:
    jst = now + _dt.timedelta(hours=9)
    end_jst = jst.replace(hour=23, minute=59, second=59, microsecond=0)
    return end_jst - _dt.timedelta(hours=9)


def _func_host() -> str:
    return os.environ.get("WEBSITE_HOSTNAME", "func-employee-portal-7833.azurewebsites.net")


def _page_url(token: str) -> str:
    return f"https://{_func_host()}/api/kyouiku/page?t={token}"


def _sn_str(v: Any) -> str:
    s = str(v if v is not None else "").strip()
    return re.sub(r"\.0+$", "", s)


def _commute_kind(raw: str) -> str:
    v = re.sub(r"[\s　]", "", str(raw or ""))
    if not v:
        return "none"
    if re.match(r"^(自転車|徒歩)", v):
        return "bicycle"
    if re.match(r"^(送迎|電車)", v):
        return "sougei"
    if re.search(r"車|バイク|原付|マイカー", v):
        return "car"
    return "none"


def _lang_from_nationality(nat: str) -> str:
    v = str(nat or "")
    if re.search(r"日本|JAPAN", v, re.I):
        return "ja"
    if re.search(r"ブラジル|ﾌﾞﾗｼﾞﾙ|ペルー|ボリビア|アルゼンチン|パラグアイ|ポルトガル|BRAZIL|BRASIL|PERU", v, re.I):
        return "pt"
    return "en" if v.strip() else "ja"


# ---------- SP I/O ----------
def _sp_abs(server_rel: str) -> str:
    return f"https://{_fa().SP_HOST}{server_rel}"


def _list_files(name_filter: str) -> List[Dict[str, Any]]:
    fa = _fa()
    url = (f"{fa.SITE_TEAMSTEPUP}/_api/web/GetFolderByServerRelativeUrl('{quote(KY_FOLDER)}')"
           f"/Files?$select=Name,TimeLastModified&$filter={name_filter}&$top=500")
    r = requests.get(url, headers=fa._sp_headers(), timeout=30)
    if r.status_code == 404:
        return []
    r.raise_for_status()
    return r.json().get("value", []) or []


def _rec_file_name(rec: Dict[str, Any]) -> str:
    return f"{rec.get('syainNo') or '0'}__{rec['token']}.json"


def _load_rec(token: str, use_cache: bool = True) -> Optional[Dict[str, Any]]:
    if not TOKEN_RE.match(token or ""):
        return None
    c = _rec_cache.get(token)
    if use_cache and c and time.time() - c[0] < REC_CACHE_SEC:
        return c[1]
    try:
        files = _list_files(f"substringof('__{token}.json',Name)")
    except Exception:
        logging.exception("kyouiku list failed")
        return None
    for f in files:
        if str(f.get("Name", "")).endswith(f"__{token}.json"):
            raw = _fa()._sp_download_bytes(_sp_abs(f"{KY_FOLDER}/{f['Name']}"))
            if raw:
                try:
                    rec = json.loads(raw.decode("utf-8"))
                    _rec_cache[token] = (time.time(), rec)
                    return rec
                except Exception:
                    logging.exception("kyouiku json parse failed")
    return None


def _save_rec(rec: Dict[str, Any]) -> None:
    _fa().sp_upload_file(KY_FOLDER, _rec_file_name(rec), json.dumps(rec, ensure_ascii=False).encode("utf-8"))
    _rec_cache[rec["token"]] = (time.time(), rec)


def _recycle(server_rel: str) -> None:
    fa = _fa()
    url = f"{fa.SITE_TEAMSTEPUP}/_api/web/GetFileByServerRelativeUrl('{quote(server_rel)}')/recycle()"
    try:
        r = requests.post(url, headers=fa._sp_headers(write=True), timeout=30)
        if r.status_code >= 400 and r.status_code != 404:
            logging.warning("kyouiku recycle %s -> %s %s", server_rel, r.status_code, r.text[:200])
    except Exception:
        logging.exception("kyouiku recycle failed")


def _curriculum() -> Dict[str, Any]:
    if _cur_cache["v"] and time.time() - _cur_cache["t"] < CUR_CACHE_SEC:
        return _cur_cache["v"]
    raw = _fa()._sp_download_bytes(_sp_abs(CURRICULUM))
    if not raw:
        if _cur_cache["v"]:
            return _cur_cache["v"]
        raise RuntimeError("curriculum_not_found")
    v = json.loads(raw.decode("utf-8-sig"))
    _cur_cache.update(t=time.time(), v=v)
    return v


def _cur_items_for(commute: str) -> List[Dict[str, Any]]:
    out = []
    for it in _curriculum().get("items", []):
        f = it.get("for", "all")
        fs = f if isinstance(f, list) else [f]
        if "all" in fs or commute in fs:
            out.append(it)
    return out


def _cur_item(key: str) -> Optional[Dict[str, Any]]:
    for it in _curriculum().get("items", []):
        if it.get("key") == key:
            return it
    return None


def _file_bytes(rel: str) -> Optional[bytes]:
    """入社資料フォルダの相対パス → bytes (インスタンスの /tmp に30分キャッシュ)"""
    if ".." in rel or rel.startswith("/"):
        return None
    os.makedirs(CACHE_DIR, exist_ok=True)
    cp = os.path.join(CACHE_DIR, hashlib.sha1(rel.encode("utf-8")).hexdigest() + ".bin")
    try:
        if os.path.exists(cp) and time.time() - os.path.getmtime(cp) < FILE_CACHE_SEC:
            with open(cp, "rb") as fh:
                return fh.read()
    except Exception:
        pass
    data = _fa()._sp_download_bytes(_sp_abs(f"{SRC_ROOT}/{rel}"), timeout=120)
    if data:
        try:
            tmp = cp + ".part"
            with open(tmp, "wb") as fh:
                fh.write(data)
            os.replace(tmp, cp)
        except Exception:
            logging.exception("kyouiku cache write failed")
    return data


def _pick_file(it: Dict[str, Any], lang: str) -> Tuple[str, str]:
    """(言語, 相対パス)。all=3言語1ファイル。無い言語は ja→pt→en の順で代替。"""
    fs = it.get("files", {})
    if fs.get("all"):
        return "all", fs["all"]
    if fs.get(lang):
        return lang, fs[lang]
    for l in LANGS:
        if fs.get(l):
            return l, fs[l]
    return "", ""


def _kind(rel: str) -> str:
    return "video" if rel.lower().endswith(".mp4") else "pdf"


def _is_expired(rec: Dict[str, Any]) -> bool:
    e = _parse(rec.get("expiresAt") or "")
    return bool(e and _now() > e)


def _progress(rec: Dict[str, Any]) -> Dict[str, Any]:
    keys = rec.get("items", [])
    done = rec.get("done", {}) or {}
    n = sum(1 for k in keys if k in done)
    return {"done": n, "total": len(keys), "complete": bool(keys) and n == len(keys)}


def _public(rec: Dict[str, Any]) -> Dict[str, Any]:
    keys = ("token", "syainNo", "name", "lang", "commute", "createdAt", "expiresAt", "requester",
            "requesterName", "done", "opened", "completedAt", "firstOpenedAt")
    out = {k: rec.get(k) for k in keys}
    out["url"] = _page_url(rec["token"])
    out["expired"] = _is_expired(rec)
    out["progress"] = _progress(rec)
    titles = {}
    for k in rec.get("items", []):
        it = _cur_item(k)
        titles[k] = (it or {}).get("title", {}).get("ja", k)
    out["items"] = [{"key": k, "title": titles.get(k, k)} for k in rec.get("items", [])]
    return out


def _html_response(html: str, status: int = 200) -> func.HttpResponse:
    return func.HttpResponse(html, status_code=status, mimetype="text/html", charset="utf-8",
                             headers={"Cache-Control": "no-store, no-cache, must-revalidate, private",
                                      "X-Content-Type-Options": "nosniff",
                                      "Referrer-Policy": "no-referrer"})


def _simple_page(title: str, msg_jp: str, msg_pt: str) -> str:
    return (
        '<!doctype html><html lang="ja"><head><meta charset="utf-8">'
        '<meta name="viewport" content="width=device-width,initial-scale=1"><meta name="robots" content="noindex">'
        f'<title>{title}</title>'
        '<style>body{font-family:-apple-system,BlinkMacSystemFont,"Hiragino Sans","Yu Gothic",Meiryo,sans-serif;'
        'background:#f4f2ed;margin:0;padding:24px;color:#222}.box{max-width:520px;margin:40px auto;background:#fff;border-radius:12px;'
        'padding:26px 22px;box-shadow:0 2px 14px rgba(0,0,0,.08);text-align:center}.t{font-size:18px;font-weight:700;margin-bottom:10px}'
        '.p{font-size:14px;line-height:1.7;color:#444}.pt{font-size:13px;color:#777;margin-top:8px}</style></head><body>'
        f'<div class="box"><div class="t">{title}</div><div class="p">{msg_jp}</div><div class="pt">{msg_pt}</div></div>'
        '</body></html>'
    )


def _body(req: func.HttpRequest) -> Dict[str, Any]:
    try:
        return req.get_json() or {}
    except Exception:
        return {}


# ---------- handlers (staff) ----------
def handle_request(req: func.HttpRequest, requester_email: str) -> func.HttpResponse:
    """リンク発行。body: name, nationality, commute(通勤方法の生値), syainNo?, lang?, requesterName?"""
    fa = _fa()
    b = _body(req)
    name = str(b.get("name") or "").strip()[:60]
    if not name:
        return fa._json_response({"error": "name_required"}, 400)
    commute = _commute_kind(b.get("commute"))
    lang = str(b.get("lang") or "").strip()
    if lang not in LANGS:
        lang = _lang_from_nationality(b.get("nationality"))
    try:
        items = [it["key"] for it in _cur_items_for(commute) if KEY_RE.match(str(it.get("key", "")))]
    except Exception as e:
        logging.exception("kyouiku curriculum load failed")
        return fa._json_response({"error": "curriculum_failed", "detail": str(e)[:200]}, 500)
    now = _now()
    token = secrets.token_urlsafe(24)
    rec = {
        "token": token, "syainNo": _sn_str(b.get("syainNo")), "name": name,
        "nationality": str(b.get("nationality") or "")[:40], "commuteRaw": str(b.get("commute") or "")[:40],
        "commute": commute, "lang": lang, "items": items,
        "requester": requester_email, "requesterName": str(b.get("requesterName") or "").strip()[:60],
        "createdAt": _iso(now), "expiresAt": _iso(_end_of_day_jst(now)),
        "done": {}, "opened": {}, "firstOpenedAt": "", "completedAt": "", "device": "",
    }
    try:
        fa.sp_create_folder_if_not_exists(KY_FOLDER)
    except Exception:
        logging.exception("kyouiku folder ensure failed")
    try:
        _save_rec(rec)
    except Exception as e:
        logging.exception("kyouiku save failed")
        return fa._json_response({"error": "save_failed", "detail": str(e)[:200]}, 500)
    logging.info("kyouiku link created: %s %s by %s", rec["syainNo"], name, requester_email)
    return fa._json_response({"ok": True, "request": _public(rec)})


def handle_link(req: func.HttpRequest) -> func.HttpResponse:
    """社員登録後に社員番号をひも付ける。body: token, syainNo"""
    fa = _fa()
    b = _body(req)
    token = str(b.get("token") or "")
    sn = _sn_str(b.get("syainNo"))
    if not sn or not re.match(r"^\d{1,10}$", sn):
        return fa._json_response({"error": "invalid_syainNo"}, 400)
    rec = _load_rec(token, use_cache=False)
    if not rec:
        return fa._json_response({"error": "not_found"}, 404)
    old = _rec_file_name(rec)
    if rec.get("syainNo") == sn:
        return fa._json_response({"ok": True, "request": _public(rec)})
    rec["syainNo"] = sn
    try:
        _save_rec(rec)
    except Exception as e:
        return fa._json_response({"error": "save_failed", "detail": str(e)[:200]}, 500)
    if old != _rec_file_name(rec):
        _recycle(f"{KY_FOLDER}/{old}")
    return fa._json_response({"ok": True, "request": _public(rec)})


def handle_status(req: func.HttpRequest) -> func.HttpResponse:
    """進み具合。?t=token または ?syainNo= (新しい順)"""
    fa = _fa()
    token = req.params.get("t") or ""
    if token:
        rec = _load_rec(token, use_cache=False)
        if not rec:
            return fa._json_response({"error": "not_found"}, 404)
        return fa._json_response({"ok": True, "items": [_public(rec)]})
    sn = _sn_str(req.params.get("syainNo"))
    if not sn or not re.match(r"^\d{1,10}$", sn):
        return fa._json_response({"error": "invalid_syainNo"}, 400)
    try:
        files = _list_files(f"startswith(Name,'{sn}__')")
    except Exception as e:
        return fa._json_response({"error": "list_failed", "detail": str(e)[:200]}, 500)
    out = []
    for f in sorted(files, key=lambda x: x.get("TimeLastModified", ""), reverse=True)[:10]:
        m = re.match(r"^\d+__([A-Za-z0-9_-]+)\.json$", f.get("Name", ""))
        if not m:
            continue
        rec = _load_rec(m.group(1), use_cache=False)
        if rec:
            out.append(_public(rec))
    out.sort(key=lambda r: r.get("createdAt") or "", reverse=True)
    return fa._json_response({"ok": True, "items": out})


# ---------- handlers (token のみ・本人) ----------
def _token_rec(req: func.HttpRequest) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
    token = req.params.get("t") or ""
    rec = _load_rec(token)
    if not rec:
        return None, "not_found"
    if _is_expired(rec):
        return rec, "expired"
    return rec, None


def handle_page(req: func.HttpRequest) -> func.HttpResponse:
    rec, err = _token_rec(req)
    if err == "not_found":
        return _html_response(_simple_page("リンクが見つかりません", "リンクが正しくないか、削除されています。担当者に確認してください。",
                                           "Link inválido. Fale com o responsável."), 404)
    if err == "expired":
        return _html_response(_simple_page("リンクの期限が切れました", "このリンクは発行した日のみ使えます。担当者に新しいリンクをもらってください。",
                                           "Este link expirou. Peça um novo link ao responsável."), 410)
    tpl_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "kyouiku_page.html")
    with open(tpl_path, encoding="utf-8") as fh:
        html = fh.read()
    html = html.replace("__TOKEN__", rec["token"]).replace("__API__", f"https://{_func_host()}/api")
    if not rec.get("firstOpenedAt"):
        rec["firstOpenedAt"] = _iso(_now())
        rec["device"] = str(req.headers.get("User-Agent") or "")[:200]
        try:
            _save_rec(rec)
        except Exception:
            logging.exception("kyouiku firstOpened save failed")
    return _html_response(html)


def handle_items(req: func.HttpRequest) -> func.HttpResponse:
    fa = _fa()
    rec, err = _token_rec(req)
    if err:
        return fa._json_response({"error": err}, 410 if err == "expired" else 404)
    out = []
    for k in rec.get("items", []):
        it = _cur_item(k)
        if not it:
            continue
        files = {}
        for l in LANGS:
            fl, rel = _pick_file(it, l)
            if rel:
                files[l] = {"kind": _kind(rel), "lang": fl}
        out.append({"key": k, "icon": it.get("icon", "📄"), "title": it.get("title", {}), "files": files,
                    "done": bool((rec.get("done") or {}).get(k))})
    return fa._json_response({"ok": True, "name": rec.get("name"), "lang": rec.get("lang"),
                              "expiresAt": rec.get("expiresAt"), "items": out, "progress": _progress(rec)},
                             extra_headers={"Cache-Control": "no-store"})


def handle_file(req: func.HttpRequest) -> func.HttpResponse:
    fa = _fa()
    rec, err = _token_rec(req)
    if err:
        return func.HttpResponse("gone", status_code=410 if err == "expired" else 404)
    key = req.params.get("k") or ""
    lang = req.params.get("l") or rec.get("lang") or "ja"
    if key not in rec.get("items", []):
        return func.HttpResponse("not found", status_code=404)
    it = _cur_item(key)
    if not it:
        return func.HttpResponse("not found", status_code=404)
    _, rel = _pick_file(it, lang if lang in LANGS else "ja")
    data = _file_bytes(rel) if rel else None
    if not data:
        return func.HttpResponse("file error", status_code=502)
    mime = MIME.get(os.path.splitext(rel)[1].lower(), "application/octet-stream")
    base = {"Accept-Ranges": "bytes", "Cache-Control": "private, max-age=3600", "X-Content-Type-Options": "nosniff"}
    total = len(data)
    rng = req.headers.get("Range") or req.headers.get("range") or ""
    m = re.match(r"^bytes=(\d*)-(\d*)$", rng.strip())
    if m and (m.group(1) or m.group(2)):
        if m.group(1):
            start = int(m.group(1))
            end = int(m.group(2)) if m.group(2) else total - 1
        else:
            n = int(m.group(2))
            start, end = max(0, total - n), total - 1
        end = min(end, total - 1)
        if start > end or start >= total:
            return func.HttpResponse(status_code=416, headers={**base, "Content-Range": f"bytes */{total}"})
        return func.HttpResponse(body=data[start:end + 1], status_code=206, mimetype=mime,
                                 headers={**base, "Content-Range": f"bytes {start}-{end}/{total}"})
    return func.HttpResponse(body=data, status_code=200, mimetype=mime, headers=base)


def handle_progress(req: func.HttpRequest) -> func.HttpResponse:
    """body: t, k, ev('open'|'done'), lang?"""
    fa = _fa()
    b = _body(req)
    token = str(b.get("t") or "")
    rec = _load_rec(token, use_cache=False)
    if not rec:
        return fa._json_response({"error": "not_found"}, 404)
    if _is_expired(rec):
        return fa._json_response({"error": "expired"}, 410)
    k = str(b.get("k") or "")
    if k not in rec.get("items", []):
        return fa._json_response({"error": "bad_key"}, 400)
    ev = str(b.get("ev") or "")
    now = _iso(_now())
    lang = str(b.get("lang") or "")[:3]
    if ev == "open":
        rec.setdefault("opened", {}).setdefault(k, now)
    elif ev == "done":
        rec.setdefault("done", {})
        if k not in rec["done"]:
            rec["done"][k] = {"at": now, "lang": lang}
        if _progress(rec)["complete"] and not rec.get("completedAt"):
            rec["completedAt"] = now
    else:
        return fa._json_response({"error": "bad_event"}, 400)
    try:
        _save_rec(rec)
    except Exception as e:
        return fa._json_response({"error": "save_failed", "detail": str(e)[:200]}, 500)
    return fa._json_response({"ok": True, "progress": _progress(rec)})
