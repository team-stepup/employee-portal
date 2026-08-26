# -*- coding: utf-8 -*-
"""面接シート受付 (Forms置き換え・2026-08-26)

応募者がスマホで入力した構造化データを、既存の「面接表」List(ルートサイト)へ
従来のForms→Power Automateと同じ列構成で保存する。ID1は連番を継続。
→ 既存のスキルシート機能(一覧/検索/PDF)が過去データと新データを区別なく扱える。

書き込み: 匿名受付は MI(app-only SPトークン) / 担当者手入力は呼び出し元トークン。
列の解決: フィールド表示名から動的に行う(skillsheet._field_titles / _job_keymap を再利用)。
"""
import datetime
import json
import re

import requests

import skillsheet as _ss

SITE_ROOT = "https://teamstepupcom774.sharepoint.com"
LIST1 = _ss.LIST1
# 新規保存先: 「面接表2000件。」(/sites/PowerApps・MIで書き込み可・列構成は同一系統)
#   ルートサイトのList1はMI(Sites.Selected)の権限外のため書けない(2026-08-26検証)
SAVE_BASE = _ss.SITE_PA
SAVE_LIST = _ss.LIST2


def _hdr(token, write=False):
    h = {"Authorization": f"Bearer {token}",
         "Accept": "application/json;odata=nometadata"}
    if write:
        h["Content-Type"] = "application/json;odata=nometadata"
    return h


def next_id1(token):
    """ID1の連番を継続: List1の最大 と List2直近行の最大 の大きい方+1"""
    cur = 0
    try:
        url = (f"{SITE_ROOT}/_api/web/lists(guid'{LIST1}')/items"
               f"?$select={_ss.K['id1']}&$orderby={_ss.K['id1']} desc&$top=1")
        r = requests.get(url, headers=_hdr(token), timeout=30)
        r.raise_for_status()
        vals = r.json().get("value", [])
        if vals and vals[0].get(_ss.K["id1"]) is not None:
            cur = int(float(vals[0][_ss.K["id1"]]))
    except Exception:
        pass
    try:
        url = (f"{SAVE_BASE}/_api/web/lists(guid'{SAVE_LIST}')/items"
               f"?$select={_ss.K['id1']}&$orderby=Id desc&$top=50")
        r = requests.get(url, headers=_hdr(token), timeout=30)
        r.raise_for_status()
        for it in r.json().get("value", []):
            v = it.get(_ss.K["id1"])
            if v is not None:
                cur = max(cur, int(float(v)))
    except Exception:
        pass
    return cur + 1


def _title_key(token, *words):
    """保存先Listのフィールド表示名に words を全て含む列の itemキー を返す"""
    for key, title in _ss._field_titles(token, SAVE_BASE, SAVE_LIST):
        t = str(title)
        if all(w in t for w in words):
            return key
    return None


def build_row(token, d, id1, tantou=""):
    """入力データ(dict) → List1 itemキー: 値 のペイロード"""
    row = {}

    def put(key, val):
        if key and val not in (None, ""):
            row[key] = val

    K = _ss.K
    put("Title", (f"{tantou}　{id1}" if tantou else str(id1)))
    put(K["id1"], int(id1))
    put(K["done"], datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"))
    put(K["name"], str(d.get("name") or "").strip()[:100])
    put(K["sex"], d.get("sex"))
    put(K["marital"], d.get("marital"))
    put(K["addr"], str(d.get("addr") or "").strip()[:200])
    put(K["commute"], d.get("commute"))
    put(K["zangyo"], d.get("zangyo"))
    put(K["sat"], d.get("sat"))
    put(K["kotai"], d.get("kotai"))
    put(K["uniform"], d.get("uniform"))
    put(K["shoes"], d.get("shoes"))
    put(K["jp_rikai"], d.get("jpRikai"))
    put(K["jp_hanasu"], d.get("jpHanasu"))
    put(K["jp_pct"], d.get("jpPct"))
    put(K["license"], d.get("license"))
    # 生年月日 → 年齢も自動計算
    birth = str(d.get("birth") or "").strip()   # YYYY-MM-DD
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", birth):
        put(_title_key(token, "生年月日"), birth + "T00:00:00Z")
        try:
            b = datetime.date.fromisoformat(birth)
            today = datetime.date.today()
            age = today.year - b.year - ((today.month, today.day) < (b.month, b.day))
            put(K["age"], str(age))
        except Exception:
            pass
    if d.get("age"):
        put(K["age"], str(d.get("age")))
    # 動的解決の列
    put(_title_key(token, "国籍"), d.get("nationality"))
    put(_title_key(token, "電話番号"), str(d.get("tel") or "").strip()[:40])
    put(_title_key(token, "メールアドレス"), str(d.get("mail") or "").strip()[:120])
    put(_title_key(token, "緊急連絡先"), str(d.get("emergency") or "").strip()[:200])
    put(_title_key(token, "在留資格"), d.get("zairyu"))
    put(_title_key(token, "免許証有効期限"), d.get("licenseExp"))
    yomi = d.get("yomeru")
    kaku = d.get("kakeru")
    put(_title_key(token, "読める"), json.dumps(yomi, ensure_ascii=False) if isinstance(yomi, list) and yomi else yomi)
    put(_title_key(token, "書く"), json.dumps(kaku, ensure_ascii=False) if isinstance(kaku, list) and kaku else kaku)
    # 職歴(最大4件・行単位)
    km = _ss._job_keymap(token, SAVE_BASE, SAVE_LIST)
    jobs = d.get("jobs") or []
    # 退職理由列も動的に(①/末尾数字)
    taishoku = {}
    for key, title in _ss._field_titles(token, SAVE_BASE, SAVE_LIST):
        t = str(title)
        if "退職した理由" not in t:
            continue
        if "①" in t:
            i = 0
        else:
            m = re.search(r"([2-4])\s*$", t.strip())
            i = (int(m.group(1)) - 1) if m else 0
        taishoku.setdefault(i, key)
    haken = {}
    for key, title in _ss._field_titles(token, SAVE_BASE, SAVE_LIST):
        t = str(title)
        if "派遣会社" not in t:
            continue
        for mark, idx in (("①", 0), ("②", 1), ("③", 2), ("④", 3)):
            if mark in t:
                haken.setdefault(idx, key)
                break
    for i, jb in enumerate(jobs[:4]):
        if not isinstance(jb, dict):
            continue

        def jkey(cat):
            ks = km.get(cat, {}).get(i, [])
            return ks[0] if ks else None

        put(jkey("company"), str(jb.get("company") or "").strip()[:100])
        put(jkey("shokushu"), str(jb.get("shokushu") or "").strip()[:100])
        put(jkey("naiyou"), str(jb.get("naiyou") or "").strip()[:200])
        put(jkey("kikan"), str(jb.get("kikan") or "").strip()[:50])
        put(haken.get(i), str(jb.get("haken") or "").strip()[:100])
        put(taishoku.get(i), str(jb.get("riyuu") or "").strip()[:200])
    return row


def create_row(token, row):
    """保存先List(面接表2000件。)へ新規行を作成(nometadata POST)。作成アイテムIdを返す。"""
    r = requests.post(f"{SAVE_BASE}/_api/web/lists(guid'{SAVE_LIST}')/items",
                      headers=_hdr(token, write=True),
                      data=json.dumps(row, ensure_ascii=False).encode("utf-8"),
                      timeout=60)
    r.raise_for_status()
    return r.json().get("Id")
