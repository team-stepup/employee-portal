# -*- coding: utf-8 -*-
"""電子署名リンク基盤 (esign) — 2026-09-03

yukyu-app の書類画面(労働契約書 兼 労働条件通知書 等)から「本人へ署名リンクを送る」を押すと、
書類HTML(自己完結・署名UI付き)を SharePoint に預けて使い捨てリンクを発行する。
本人はスマホ/iPad/PC でリンクを開き(ログイン不要)、内容確認 → 指でサイン → 送信。
署名ページ自身が html2canvas + pdf-lib で署名入りPDFを生成し、/api/esign/submit へ送る。
バックエンド(MI)が社員フォルダ(＋新規雇用契約書 集約フォルダ)へ保存し、発行者へメール通知。

保存場所: /sites/TeamStepup/Shared Documents/電子署名待ち/{社員番号}__{token}.json / .html
  json = 依頼メタ(状態・期限・保存先)、html = 署名ページ本体
"""
import base64
import datetime as _dt
import json
import logging
import os
import re
import secrets
from typing import Any, Dict, List, Optional
from urllib.parse import quote

import requests
import azure.functions as func

ESIGN_FOLDER = "/sites/TeamStepup/Shared Documents/電子署名待ち"
ESIGN_TTL_DAYS = 7          # リンク有効期間
SIGNED_DL_DAYS = 30         # 署名後、本人が控えPDFをダウンロードできる期間
TOKEN_RE = re.compile(r"^[A-Za-z0-9_-]{20,64}$")
MAX_HTML_BYTES = 8 * 1024 * 1024
MAX_PDF_BYTES = 40 * 1024 * 1024


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


def _jst(s: str) -> str:
    d = _parse(s)
    if not d:
        return s or ""
    return (d + _dt.timedelta(hours=9)).strftime("%Y/%m/%d %H:%M")


def _func_host() -> str:
    return os.environ.get("WEBSITE_HOSTNAME", "func-employee-portal-7833.azurewebsites.net")


def _page_url(token: str) -> str:
    return f"https://{_func_host()}/api/esign/page?t={token}"


def _safe_name(s: str, default: str = "file") -> str:
    s = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", str(s or "")).strip()
    return s or default


def _sn_str(v: Any) -> str:
    s = str(v if v is not None else "").strip()
    return re.sub(r"\.0+$", "", s)


# ---------- SP I/O ----------
def _list_files(name_filter: str) -> List[Dict[str, Any]]:
    fa = _fa()
    url = (f"{fa.SITE_TEAMSTEPUP}/_api/web/GetFolderByServerRelativeUrl('{quote(ESIGN_FOLDER)}')"
           f"/Files?$select=Name,TimeLastModified&$filter={name_filter}&$top=500")
    r = requests.get(url, headers=fa._sp_headers(), timeout=30)
    if r.status_code == 404:
        return []
    r.raise_for_status()
    return r.json().get("value", []) or []


def _download(name: str) -> Optional[bytes]:
    fa = _fa()
    return fa._sp_download_bytes(f"https://{fa.SP_HOST}{ESIGN_FOLDER}/{name}")


def _load_rec(token: str) -> Optional[Dict[str, Any]]:
    if not TOKEN_RE.match(token or ""):
        return None
    try:
        files = _list_files(f"substringof('__{token}.json',Name)")
    except Exception:
        logging.exception("esign list failed")
        return None
    for f in files:
        if str(f.get("Name", "")).endswith(f"__{token}.json"):
            raw = _download(f["Name"])
            if raw:
                try:
                    return json.loads(raw.decode("utf-8"))
                except Exception:
                    logging.exception("esign json parse failed")
    return None


def _save_rec(rec: Dict[str, Any]) -> None:
    fa = _fa()
    fa.sp_upload_file(ESIGN_FOLDER, f"{rec['syainNo']}__{rec['token']}.json",
                      json.dumps(rec, ensure_ascii=False).encode("utf-8"))


def _ensure_esign_folder() -> None:
    try:
        _fa().sp_create_folder_if_not_exists(ESIGN_FOLDER)
    except Exception:
        logging.exception("esign folder ensure failed")


def _public(rec: Dict[str, Any]) -> Dict[str, Any]:
    keys = ("token", "syainNo", "name", "docLabel", "fileName", "folderPath", "createdAt", "expiresAt",
            "status", "signedAt", "requester", "empEmail", "empPhone", "savedUrl")
    out = {k: rec.get(k) for k in keys}
    out["url"] = _page_url(rec["token"])
    out["expired"] = _is_expired(rec)
    return out


def _is_expired(rec: Dict[str, Any]) -> bool:
    e = _parse(rec.get("expiresAt") or "")
    return bool(e and _now() > e)


def _html_response(html: str, status: int = 200) -> func.HttpResponse:
    return func.HttpResponse(html, status_code=status, mimetype="text/html", charset="utf-8",
                             headers={"Cache-Control": "no-store, no-cache, must-revalidate, private",
                                      "X-Content-Type-Options": "nosniff",
                                      "X-Frame-Options": "DENY",
                                      "Referrer-Policy": "no-referrer"})


def _simple_page(title: str, msg_jp: str, msg_pt: str, extra_html: str = "") -> str:
    return (
        '<!doctype html><html lang="ja"><head><meta charset="utf-8">'
        '<meta name="viewport" content="width=device-width,initial-scale=1">'
        f'<title>{title}</title>'
        '<style>body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Hiragino Kaku Gothic ProN","Yu Gothic",Meiryo,sans-serif;'
        'background:#f0f2f5;margin:0;padding:24px;color:#222}.box{max-width:560px;margin:40px auto;background:#fff;border-radius:12px;'
        'padding:26px 22px;box-shadow:0 2px 14px rgba(0,0,0,.08);text-align:center}.t{font-size:18px;font-weight:700;margin-bottom:10px}'
        '.p{font-size:14px;line-height:1.7;color:#444}.pt{font-size:13px;color:#777;margin-top:8px}a.btn{display:inline-block;margin-top:16px;'
        'background:#1565c0;color:#fff;text-decoration:none;padding:12px 20px;border-radius:8px;font-weight:700}</style></head><body>'
        f'<div class="box"><div class="t">{title}</div><div class="p">{msg_jp}</div><div class="pt">{msg_pt}</div>{extra_html}</div>'
        '</body></html>'
    )


# ---------- handlers ----------
def handle_request(req: func.HttpRequest, requester_email: str) -> func.HttpResponse:
    """署名リンク発行 (staff)。body: syainNo,name,docLabel,folderPath,fileName,aggFolderPath?,aggFileName?,html,empEmail?,empPhone?"""
    fa = _fa()
    try:
        body = req.get_json()
    except Exception:
        return fa._json_response({"error": "invalid_json"}, 400)
    sn = _sn_str(body.get("syainNo"))
    name = str(body.get("name") or "").strip()
    html = body.get("html") or ""
    folder = str(body.get("folderPath") or "").strip()
    file_name = _safe_name(body.get("fileName"), "署名済.pdf")
    if not sn or not name or not html or not folder:
        return fa._json_response({"error": "missing_fields"}, 400)
    if not fa._validate_shainfile_path(folder):
        return fa._json_response({"error": "invalid_folder"}, 400)
    if len(html.encode("utf-8")) > MAX_HTML_BYTES:
        return fa._json_response({"error": "html_too_large"}, 413)
    agg_folder = str(body.get("aggFolderPath") or "").strip()
    if agg_folder and not fa._validate_shainfile_path(agg_folder):
        agg_folder = ""
    agg_name = _safe_name(body.get("aggFileName"), "") if agg_folder else ""

    token = secrets.token_urlsafe(24)
    now = _now()
    rec = {
        "token": token, "syainNo": sn, "name": name,
        "docLabel": str(body.get("docLabel") or "労働契約書")[:80],
        "folderPath": folder, "fileName": file_name,
        "aggFolderPath": agg_folder, "aggFileName": agg_name,
        "empEmail": str(body.get("empEmail") or "").strip()[:120],
        "empPhone": str(body.get("empPhone") or "").strip()[:40],
        "requester": requester_email,
        "createdAt": _iso(now), "expiresAt": _iso(now + _dt.timedelta(days=ESIGN_TTL_DAYS)),
        "status": "pending", "signedAt": "", "signedIp": "", "savedUrl": "",
    }
    _ensure_esign_folder()
    try:
        fa.sp_upload_file(ESIGN_FOLDER, f"{sn}__{token}.html", html.encode("utf-8"))
        _save_rec(rec)
    except Exception as e:
        logging.exception("esign request save failed")
        return fa._json_response({"error": "save_failed", "detail": str(e)[:200]}, 500)
    logging.info("esign request created: %s %s by %s", sn, name, requester_email)
    return fa._json_response({"ok": True, "request": _public(rec)})


def handle_pending(req: func.HttpRequest) -> func.HttpResponse:
    """社員番号の署名依頼一覧 (staff)。新しい順。"""
    fa = _fa()
    sn = _sn_str(req.params.get("syainNo"))
    if not sn or not re.match(r"^\d{1,8}$", sn):
        return fa._json_response({"error": "invalid_syainNo"}, 400)
    try:
        files = _list_files(f"startswith(Name,'{sn}__') and substringof('.json',Name)")
    except Exception as e:
        logging.exception("esign pending list failed")
        return fa._json_response({"error": "list_failed", "detail": str(e)[:200]}, 500)
    items = []
    for f in files:
        raw = _download(f["Name"])
        if not raw:
            continue
        try:
            items.append(_public(json.loads(raw.decode("utf-8"))))
        except Exception:
            continue
    items.sort(key=lambda x: x.get("createdAt") or "", reverse=True)
    return fa._json_response({"ok": True, "items": items[:20]})


def _mail_body(rec: Dict[str, Any], note: str = "") -> str:
    url = _page_url(rec["token"])
    exp = _jst(rec.get("expiresAt", ""))
    return (
        '<div style="font-family:Segoe UI,Meiryo,sans-serif;font-size:14px;color:#222;line-height:1.7">'
        f'<p>{rec["name"]} 様</p>'
        f'<p>有限会社ステップ・アップです。<b>{rec.get("docLabel") or "労働契約書"}</b> の電子署名をお願いします。<br>'
        f'下のリンクを開き、内容をご確認のうえ、最下部の欄に指でサインして「署名して送信」を押してください。</p>'
        f'<p><a href="{url}" style="display:inline-block;background:#1565c0;color:#fff;padding:12px 20px;border-radius:8px;text-decoration:none;font-weight:700">✍ 書類を開いて署名する</a></p>'
        f'<p style="font-size:12px;color:#666">リンクの有効期限: {exp}（1回限り）<br>{url}</p>'
        '<hr style="border:none;border-top:1px solid #ddd;margin:14px 0">'
        f'<p style="color:#444">Sr(a). {rec["name"]},<br>Aqui é a Step Up. Por favor, abra o link acima, confira o conteúdo do '
        f'<b>{rec.get("docLabel") or "contrato de trabalho"}</b> e assine com o dedo no campo ao final, depois toque em "Assinar e enviar".<br>'
        f'<span style="font-size:12px;color:#666">Válido até: {exp} (uso único)</span></p>'
        f'{note}</div>'
    )


def handle_send(req: func.HttpRequest, requester_email: str) -> func.HttpResponse:
    """リンクをメール送付 (staff)。to: 'ipad' (事務所iPad jimusyo1) | 'employee' (本人のＥメール) | 任意アドレス"""
    fa = _fa()
    try:
        body = req.get_json()
    except Exception:
        return fa._json_response({"error": "invalid_json"}, 400)
    rec = _load_rec(str(body.get("t") or ""))
    if not rec:
        return fa._json_response({"error": "not_found"}, 404)
    if rec.get("status") != "pending" or _is_expired(rec):
        return fa._json_response({"error": "not_pending"}, 409)
    to = str(body.get("to") or "").strip()
    if to == "ipad":
        addr = os.environ.get("ESIGN_IPAD_MAIL", "jimusyo1@team-stepup.com").strip()
        note = (f'<p style="font-size:12px;color:#999">（事務所iPad用: 来所時にこのメールのリンクを開いて署名してもらってください。'
                f'発行者: {requester_email}）</p>')
    elif to == "employee":
        addr = rec.get("empEmail") or ""
        note = ""
        if not addr:
            return fa._json_response({"error": "no_employee_email"}, 400)
    else:
        addr = to
        note = ""
    if not re.match(r"^[^\s@]+@[^\s@]+\.[^\s@]+$", addr):
        return fa._json_response({"error": "invalid_address"}, 400)
    if not os.environ.get("SHORUI_MAIL_SENDER", "").strip():
        return fa._json_response({"error": "mail_sender_not_configured"}, 500)
    subject = f"【電子署名のお願い / Assinatura eletrônica】{rec.get('docLabel') or '労働契約書'} — {rec['name']} 様"
    try:
        fa._send_notification_mail(subject, _mail_body(rec, note), to_addr=addr, html=True)
    except Exception as e:
        logging.exception("esign send mail failed")
        return fa._json_response({"error": "send_failed", "detail": str(e)[:200]}, 500)
    sent = rec.get("sentTo") or []
    sent.append({"to": addr, "at": _iso(_now()), "by": requester_email})
    rec["sentTo"] = sent[-20:]
    try:
        _save_rec(rec)
    except Exception:
        logging.exception("esign sentTo save failed")
    return fa._json_response({"ok": True, "to": addr})


def handle_page(req: func.HttpRequest) -> func.HttpResponse:
    """本人向け署名ページ (認証なし・トークンのみ)。"""
    token = str(req.params.get("t") or "").strip()
    rec = _load_rec(token)
    if not rec:
        return _html_response(_simple_page("リンクが無効です / Link inválido",
                                           "このリンクは無効です。担当者にご確認ください。",
                                           "Este link não é válido. Por favor, fale com o responsável."), 404)
    if rec.get("status") == "signed":
        dl = f'<a class="btn" href="/api/esign/signed?t={token}">📄 署名済PDFを開く / Abrir PDF assinado</a>'
        return _html_response(_simple_page("署名済みです / Já assinado",
                                           f"この書類は {_jst(rec.get('signedAt',''))} に署名済みです。ありがとうございました。",
                                           "Este documento já foi assinado. Obrigado!", dl))
    if _is_expired(rec):
        return _html_response(_simple_page("期限切れ / Expirado",
                                           "このリンクは有効期限が切れています。担当者に再発行を依頼してください。",
                                           "Este link expirou. Peça ao responsável para reenviar."), 410)
    raw = _download(f"{rec['syainNo']}__{token}.html")
    if not raw:
        return _html_response(_simple_page("エラー / Erro", "書類を読み込めませんでした。時間をおいて再度お試しください。",
                                           "Não foi possível carregar o documento. Tente novamente mais tarde."), 500)
    return _html_response(raw.decode("utf-8", errors="replace"))


def handle_submit(req: func.HttpRequest) -> func.HttpResponse:
    """署名済PDF受領 (認証なし・トークンのみ) → 社員フォルダへ保存 → 発行者へ通知。"""
    fa = _fa()
    try:
        body = req.get_json()
    except Exception:
        return fa._json_response({"error": "invalid_json"}, 400)
    token = str(body.get("t") or "").strip()
    rec = _load_rec(token)
    if not rec:
        return fa._json_response({"error": "not_found"}, 404)
    if rec.get("status") != "pending":
        return fa._json_response({"error": "already_signed"}, 409)
    if _is_expired(rec):
        return fa._json_response({"error": "expired"}, 410)
    try:
        pdf = base64.b64decode(str(body.get("pdf") or ""), validate=False)
    except Exception:
        return fa._json_response({"error": "invalid_pdf"}, 400)
    if not pdf or not pdf.startswith(b"%PDF") or len(pdf) > MAX_PDF_BYTES:
        return fa._json_response({"error": "invalid_pdf"}, 400)

    folder = rec["folderPath"]
    file_name = rec["fileName"]
    try:
        try:
            fa.sp_create_folder_if_not_exists(folder)
        except Exception:
            logging.warning("esign: employee folder ensure failed (continue): %s", folder)
        fa.sp_upload_file(folder, file_name, pdf)
    except Exception as e:
        logging.exception("esign upload failed")
        return fa._json_response({"error": "upload_failed", "detail": str(e)[:200]}, 500)
    agg_info = ""
    if rec.get("aggFolderPath") and rec.get("aggFileName"):
        try:
            fa.sp_create_folder_if_not_exists(rec["aggFolderPath"])
        except Exception:
            pass
        try:
            fa.sp_upload_file(rec["aggFolderPath"], rec["aggFileName"], pdf)
            agg_info = f'<br>集約: {rec["aggFolderPath"]}/{rec["aggFileName"]}'
        except Exception:
            logging.exception("esign agg upload failed (ignored)")

    ip = (req.headers.get("X-Forwarded-For") or "").split(",")[0].strip()
    rec["status"] = "signed"
    rec["signedAt"] = _iso(_now())
    rec["signedIp"] = ip[:64]
    rec["savedUrl"] = f"https://{fa.SP_HOST}{folder}/{file_name}"
    try:
        _save_rec(rec)
    except Exception:
        logging.exception("esign rec save failed after upload")

    # 発行者へ完了通知 (best-effort・メール)
    try:
        folder_web = f"https://{fa.SP_HOST}{quote(folder)}"
        html = (
            '<div style="font-family:Segoe UI,Meiryo,sans-serif;font-size:14px;color:#222;line-height:1.7">'
            f'<div style="font-size:16px;font-weight:700">✅ 電子署名が完了しました</div>'
            f'<p><b>{rec["name"]}</b>（No.{rec["syainNo"]}）— {rec.get("docLabel") or ""}<br>'
            f'署名日時: {_jst(rec["signedAt"])}</p>'
            f'<p>保存先: <a href="{folder_web}">{folder}</a><br>ファイル: {file_name}{agg_info}</p></div>'
        )
        if rec.get("requester"):
            fa._send_notification_mail(f"✅ 電子署名完了: {rec['name']}（{rec.get('docLabel') or ''}）", html,
                                       to_addr=rec["requester"], html=True)
    except Exception:
        logging.exception("esign notify failed (ignored)")
    return fa._json_response({"ok": True, "fileName": file_name, "signedAt": rec["signedAt"]})


def handle_signed(req: func.HttpRequest) -> func.HttpResponse:
    """本人の控え: 署名済PDFを返す (署名後 SIGNED_DL_DAYS 日間)。"""
    fa = _fa()
    token = str(req.params.get("t") or "").strip()
    rec = _load_rec(token)
    if not rec or rec.get("status") != "signed" or not rec.get("savedUrl"):
        return _html_response(_simple_page("見つかりません / Não encontrado", "署名済の書類が見つかりません。", "Documento não encontrado."), 404)
    s = _parse(rec.get("signedAt") or "")
    if s and _now() > s + _dt.timedelta(days=SIGNED_DL_DAYS):
        return _html_response(_simple_page("期限切れ / Expirado", "ダウンロード期間が終了しました。担当者にご相談ください。",
                                           "O período de download terminou. Fale com o responsável."), 410)
    raw = fa._sp_download_bytes(rec["savedUrl"])
    if not raw:
        return _html_response(_simple_page("エラー / Erro", "PDFを取得できませんでした。", "Não foi possível obter o PDF."), 500)
    fn = quote(rec.get("fileName") or "signed.pdf")
    return func.HttpResponse(raw, status_code=200, mimetype="application/pdf",
                             headers={"Content-Disposition": f"inline; filename*=UTF-8''{fn}",
                                      "Cache-Control": "no-store, private", "X-Content-Type-Options": "nosniff"})
