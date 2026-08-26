# -*- coding: utf-8 -*-
"""スキルシートPDF生成 (yukyu-app向け・2026-08-25)

旧PowerApps「面接表mensetsuhyou」のブラウザ印刷様式を再現する。
方式: 空欄テンプレ画像(skillsheet_assets/blank_form.png, 958x1493)に
      Pillowで値を描画し、A4ページのPDFとして書き出す。
座標系: テンプレ画像ピクセル(958x1493基準)。元PDFの文字位置から採取。
データ源: ルート「面接表」List(作業用・優先) + /sites/PowerApps「面接表2000件。」(補完)
"""
import io
import json
import os
import re
import datetime

import requests
from PIL import Image, ImageDraw, ImageFont

SP_HOST = "https://teamstepupcom774.sharepoint.com"
SITE_PA = SP_HOST + "/sites/PowerApps"
LIST1 = "40344733-500d-4d73-902f-dafca65ff24c"   # 面接表(ルート・作業用)
LIST2 = "a81b8fb8-eb1c-4712-be2f-10978bd9e015"   # 面接表2000件。(アーカイブ)

_HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(_HERE, "skillsheet_assets")
BLANK = os.path.join(ASSETS, "blank_form.png")
FONT_CANDIDATES = [
    os.path.join(ASSETS, "font.otf"),
    os.path.join(ASSETS, "font.ttf"),
    r"C:\Windows\Fonts\YuGothM.ttc",   # ローカルテスト用フォールバック
    r"C:\Windows\Fonts\meiryo.ttc",
]

K = {
    "id1": "OData__x0049_D1",
    "name": "OData__x540d__x524d__x3000_Nome",
    "sex": "OData__x6027__x3000__x5225__x0020__x00",
    "age": "OData__x5e74__x9f62__x3000_idade",
    "marital": "OData__x5a5a__x59fb__x72b6__x6cc1__x00",
    "commute": "OData__x901a__x52e4__x65b9__x6cd5__x30",
    "addr": "OData__x4f4f__x6240__x0020_ENDERE_x00c",
    "zangyo": "OData__x8cea__x554f__x6b8b__x696d__x30",
    "sat": "OData__x571f__x66dc__x65e5__x0020_S_x0",
    "kotai": "OData__x4ea4__x66ff__x3000_Turno_x3000",
    "license": "OData__x904b__x8ee2__x514d__x8a31__x8a",
    "uniform": "OData__x4f5c__x696d__x670d__x3000_Unif",
    "shoes": "OData__x9774__x306e__x30b5__x30a4__x30",
    "jp_rikai": "OData__x7406__x89e3__x3059__x308b__x00",
    "jp_hanasu": "OData__x8a71__x3059__x0020_Falar",
    "jp_pct": "OData__x65e5__x672c__x8a9e__x3000__x30",
    "done": "OData__x5b8c__x4e86__x6642__x523b_",
}

# ※$selectは使わない: List1/List2で列構成が異なり、存在しない列指定は400になる


def _sp_get(url, token):
    r = requests.get(url, headers={"Authorization": f"Bearer {token}",
                                   "Accept": "application/json;odata=nometadata"}, timeout=30)
    r.raise_for_status()
    return r.json()


def _norm(s):
    return re.sub(r"[\s　]+", "", str(s or ""))


def _clean(v):
    if v is None:
        return ""
    s = str(v).replace("\t", " ").strip()
    if s.startswith("[") and s.endswith("]"):
        try:
            s = "・".join(json.loads(s))
        except Exception:
            pass
    return s


def _first(v):
    s = _clean(v)
    return re.split(r"[\s　/／−]", s)[0] if s else ""


# ラテン文字塊(ポルトガル語/英語/ローマ字併記)の除去 → 客先向け日本語表記
_LAT_RE = re.compile(
    r"[-‐－]?[0-9]*[A-Za-zÀ-ÖØ-öø-ÿ][A-Za-z0-9À-ÖØ-öø-ÿ'.]*"
    r"(?:[-‐－][A-Za-z0-9À-ÖØ-öø-ÿ'.]+)*")


def jp_clean(v):
    s = _clean(v)
    if not s:
        return ""
    s = _LAT_RE.sub("", s)
    s = re.sub(r"\s*[-‐－]\s*(?=\s|$|[・、。])", " ", s)
    s = re.sub(r"[・]{2,}", "・", s)
    s = re.sub(r"[\s　]+", " ", s)
    s = re.sub(r"(?<=[\s　])\d+(?=[\s　]|$)", "", s)  # 併記除去で残った孤立数字
    s = s.strip(" ・-‐－/／,、~〜？?")
    return s.strip()


# ============================================================
# フリガナ自動生成 (ローマ字/ひらがな → カタカナ。漢字入りは空を返す)
# ============================================================
_R2K = {}
for _row in [
    ("kya","キャ"),("kyu","キュ"),("kyo","キョ"),("gya","ギャ"),("gyu","ギュ"),("gyo","ギョ"),
    ("sha","シャ"),("shu","シュ"),("sho","ショ"),("sya","シャ"),("syu","シュ"),("syo","ショ"),
    ("cha","チャ"),("chu","チュ"),("cho","チョ"),("tya","チャ"),("tyu","チュ"),("tyo","チョ"),
    ("nya","ニャ"),("nyu","ニュ"),("nyo","ニョ"),("hya","ヒャ"),("hyu","ヒュ"),("hyo","ヒョ"),
    ("bya","ビャ"),("byu","ビュ"),("byo","ビョ"),("pya","ピャ"),("pyu","ピュ"),("pyo","ピョ"),
    ("mya","ミャ"),("myu","ミュ"),("myo","ミョ"),("rya","リャ"),("ryu","リュ"),("ryo","リョ"),
    ("jya","ジャ"),("jyu","ジュ"),("jyo","ジョ"),("shi","シ"),("chi","チ"),("tsu","ツ"),
    ("tha","タ"),("thi","ティ"),("thu","トゥ"),("the","テ"),("tho","ト"),
    ("pha","ファ"),("phi","フィ"),("phu","フ"),("phe","フェ"),("pho","フォ"),
    ("qua","クァ"),("qui","キ"),("que","ケ"),("quo","クォ"),
    ("ka","カ"),("ki","キ"),("ku","ク"),("ke","ケ"),("ko","コ"),
    ("ga","ガ"),("gi","ギ"),("gu","グ"),("ge","ゲ"),("go","ゴ"),
    ("sa","サ"),("si","シ"),("su","ス"),("se","セ"),("so","ソ"),
    ("za","ザ"),("zi","ジ"),("zu","ズ"),("ze","ゼ"),("zo","ゾ"),
    ("ja","ジャ"),("ji","ジ"),("ju","ジュ"),("je","ジェ"),("jo","ジョ"),
    ("ta","タ"),("ti","チ"),("tu","ツ"),("te","テ"),("to","ト"),
    ("da","ダ"),("di","ヂ"),("du","ヅ"),("de","デ"),("do","ド"),
    ("na","ナ"),("ni","ニ"),("nu","ヌ"),("ne","ネ"),("no","ノ"),
    ("ha","ハ"),("hi","ヒ"),("hu","フ"),("he","ヘ"),("ho","ホ"),
    ("fa","ファ"),("fi","フィ"),("fu","フ"),("fe","フェ"),("fo","フォ"),
    ("ba","バ"),("bi","ビ"),("bu","ブ"),("be","ベ"),("bo","ボ"),
    ("pa","パ"),("pi","ピ"),("pu","プ"),("pe","ペ"),("po","ポ"),
    ("ma","マ"),("mi","ミ"),("mu","ム"),("me","メ"),("mo","モ"),
    ("ya","ヤ"),("yu","ユ"),("ye","イェ"),("yo","ヨ"),
    ("ra","ラ"),("ri","リ"),("ru","ル"),("re","レ"),("ro","ロ"),
    ("la","ラ"),("li","リ"),("lu","ル"),("le","レ"),("lo","ロ"),
    ("wa","ワ"),("wi","ウィ"),("we","ウェ"),("wo","ヲ"),
    ("va","ヴァ"),("vi","ヴィ"),("vu","ヴ"),("ve","ヴェ"),("vo","ヴォ"),
    ("ca","カ"),("ci","シ"),("cu","ク"),("ce","セ"),("co","コ"),
    ("ph","フ"),
    ("a","ア"),("i","イ"),("u","ウ"),("e","エ"),("o","オ"),
]:
    _R2K[_row[0]] = _row[1]
_LONE = {"b":"ブ","c":"ク","d":"ド","f":"フ","g":"グ","h":"","j":"ジュ","k":"ク","l":"ル",
         "m":"ム","p":"プ","q":"ク","r":"ル","s":"ス","t":"ト","v":"ヴ","w":"ウ",
         "x":"クス","y":"イ","z":"ズ"}


def _romaji_to_kata(w):
    w = w.lower().replace("'", "")
    out = ""
    i = 0
    while i < len(w):
        c = w[i]
        if c == "-":
            out += "ー"
            i += 1
            continue
        if i + 1 < len(w) and c == w[i + 1] and c not in "aiueon":
            out += "ッ"
            i += 1
            continue
        if w.startswith("tch", i):
            out += "ッ"
            i += 1
            continue
        matched = False
        for ln in (3, 2):
            seg = w[i:i + ln]
            if seg in _R2K:
                out += _R2K[seg]
                i += ln
                matched = True
                break
        if matched:
            continue
        if c in _R2K:
            out += _R2K[c]
        elif c == "n":
            out += "ン"
        elif c in _LONE:
            out += _LONE[c]
        i += 1
    return out


def to_katakana(name):
    """氏名のフリガナ自動生成。カタカナ=そのまま/ひらがな=カタカナ化/ローマ字=変換。
    漢字が含まれる場合は自動変換不可として空を返す(手入力運用)。"""
    s = _clean(name)
    if not s:
        return ""
    if re.search(r"[一-鿿々]", s):
        return ""
    out = []
    for token in re.split(r"[\s　]+", s):
        if not token:
            continue
        if re.fullmatch(r"[ぁ-ゖー]+", token):
            out.append("".join(chr(ord(ch) + 0x60) if "ぁ" <= ch <= "ゖ" else ch for ch in token))
        elif re.fullmatch(r"[ァ-ヶー・]+", token):
            out.append(token)
        elif re.fullmatch(r"[A-Za-z'\-]+", token):
            out.append(_romaji_to_kata(token))
        else:
            return ""
    return "　".join(x for x in out if x)


def addr_short(v):
    """住所は町名まで・番地/建物名は載せない(2026-08-26ユーザー指示)。
    例: 浜松市中央区富塚町3602-1グリーンヒルズA101 → 浜松市中央区富塚町"""
    s = jp_clean(v)
    if not s:
        return ""
    s = re.split(r"[0-9０-９]", s)[0]   # 最初の数字(番地)以降を落とす
    return s.strip(" -‐－、,")


# ============================================================
# フィールド定義 (itemキー→表示名, フォーム質問順) と 生回答
# ============================================================
_field_title_cache = {}
_SYS_INTERNALS = {"Title", "ContentType", "Attachments", "Edit", "LinkTitle",
                  "LinkTitleNoMenu", "DocIcon", "ItemChildCount", "FolderChildCount",
                  "AppAuthor", "AppEditor", "ComplianceAssetId", "_UIVersionString"}


def _field_titles(token, base, guid):
    if guid in _field_title_cache:
        return _field_title_cache[guid]
    r = requests.get(
        f"{base}/_api/web/lists(guid'{guid}')/fields"
        "?$filter=Hidden eq false and ReadOnlyField eq false"
        "&$select=InternalName,Title,TypeAsString",
        headers={"Authorization": f"Bearer {token}",
                 "Accept": "application/json;odata=nometadata"}, timeout=30)
    r.raise_for_status()
    out = []
    for f in r.json().get("value", []):
        iname = f.get("InternalName") or ""
        if iname in _SYS_INTERNALS or f.get("TypeAsString") in ("Computed", "Attachments"):
            continue
        key = ("OData_" + iname) if iname.startswith("_") else iname
        title = str(f.get("Title") or iname)
        out.append((key, title))
    _field_title_cache[guid] = out
    return out


def raw_answers(token, r1, r2):
    """応募シートの生回答(原文のまま)を [{label, value}] で返す(フォーム質問順)"""
    for base, guid, row in ((SP_HOST, LIST1, r1), (SITE_PA, LIST2, r2)):
        if not row:
            continue
        out = []
        try:
            pairs = _field_titles(token, base, guid)
        except Exception:
            pairs = []
        for key, title in pairs:
            v = row.get(key)
            if v is None or v is False or str(v).strip() == "":
                continue
            out.append({"label": title, "value": _clean(v)})
        if out:
            return out
    return []


# ============================================================
# 職歴: 列ラベルの①②③④/末尾数字で行単位に組む (位置ズレ防止)
# ============================================================
def _job_keymap(token, base, guid):
    keymap = {"company": {}, "shokushu": {}, "naiyou": {}, "kikan": {}}
    for key, title in _field_titles(token, base, guid):
        t = str(title)
        cat = None
        if "会社名" in t:
            cat = "company"
        elif t.startswith("職種") or "setores" in t:
            cat = "shokushu"
        elif "仕事内容" in t:
            cat = "naiyou"
        elif "勤続期間" in t:
            cat = "kikan"
        if not cat:
            continue
        if "①" in t:
            i = 0
        elif "②" in t:
            i = 1
        elif "③" in t:
            i = 2
        elif "④" in t:
            i = 3
        else:
            m = re.search(r"([2-4])\s*$", t.strip())
            i = (int(m.group(1)) - 1) if m else 0
        keymap[cat].setdefault(i, []).append(key)
    return keymap


def jobs_from_rows(token, r1, r2):
    for base, guid, row in ((SP_HOST, LIST1, r1), (SITE_PA, LIST2, r2)):
        if not row:
            continue
        try:
            km = _job_keymap(token, base, guid)
        except Exception:
            km = {"company": {}, "shokushu": {}, "naiyou": {}, "kikan": {}}

        def val(cat, i):
            for k in km[cat].get(i, []):
                v = _clean(row.get(k))
                if v:
                    return v
            return ""

        rows = []
        for i in range(4):
            comp = val("company", i)
            sh = jp_clean(val("shokushu", i))
            na = jp_clean(val("naiyou", i))
            kk = jp_clean(val("kikan", i))
            if not (comp or sh or na or kk):
                continue
            rows.append({"company": comp,
                         "naiyou": (sh + ("　" if sh and na else "") + na),
                         "kikan": kk})
        if rows:
            while len(rows) < 5:
                rows.append({"company": "", "naiyou": "", "kikan": ""})
            return rows[:5]
    return [{"company": "", "naiyou": "", "kikan": ""} for _ in range(5)]


# ============================================================
# 検索・取得・フォーム初期値
# ============================================================
def _commute_kind(v):
    """通勤手段を 車/送迎/バイク/自転車/徒歩/その他 に分類"""
    s = _clean(v)
    if not s:
        return ""
    if "送迎" in s:
        return "送迎"
    if "自転車" in s:
        return "自転車"
    if "バイク" in s:
        return "バイク"
    if "徒歩" in s:
        return "徒歩"
    if "車" in s:
        return "車"
    return jp_clean(_first(s)) or "その他"


def _jp_min(v):
    """日本語％の下限値(例: '60％~80％'→60)"""
    m = re.search(r"(\d+)", _clean(v))
    return int(m.group(1)) if m else None


def search_candidates(token, query=""):
    """面接表List1から候補【全件】(新しい順・応募IDで重複除去・ページング)。
    絞り込み用に 性別/年齢/通勤/日本語/住所 も返す。検索/フィルタはアプリ側で行う。"""
    sel = ",".join([K["id1"], K["name"], K["done"], K["sex"], K["age"],
                    K["commute"], K["jp_pct"], K["addr"]])
    url = f"{SP_HOST}/_api/web/lists(guid'{LIST1}')/items?$select=Id,{sel}&$orderby=Id desc&$top=1000"
    items = []
    while url:
        js = _sp_get(url, token)
        items.extend(js.get("value", []))
        url = js.get("odata.nextLink") or js.get("@odata.nextLink")
        if url and not url.startswith("http"):
            url = f"{SP_HOST}/_api/" + url
        if len(items) >= 6000:
            break
    def to_cand(it, rid):
        age_s = re.sub(r"[^0-9]", "", _clean(it.get(K["age"])))
        return {
            "id1": rid,
            "name": _clean(it.get(K["name"])),
            "date": _clean(it.get(K["done"]))[:10],
            "sex": jp_clean(_first(it.get(K["sex"]))),
            "age": int(age_s) if age_s else None,
            "commute": _commute_kind(it.get(K["commute"])),
            "jp": _clean(it.get(K["jp_pct"])),
            "jpMin": _jp_min(it.get(K["jp_pct"])),
            "addr": addr_short(it.get(K["addr"])),
        }

    out = []
    seen = set()
    nq = _norm(query)
    for it in items:
        nm = _clean(it.get(K["name"]))
        if not nm:
            continue
        if nq and nq not in _norm(nm):
            continue
        rid = it.get(K["id1"])
        rid = int(float(rid)) if rid is not None else None
        if rid is not None and rid in seen:
            continue   # 再送信等でできた重複行は最新のみ表示
        seen.add(rid)
        out.append(to_cand(it, rid))

    # 新方式(Forms廃止後)の応募・手入力は「面接表2000件。」に保存される → マージ表示
    # 識別: 2026-08-26以降に作成され、かつList1に同じ応募IDが無い行
    try:
        url2 = (f"{SITE_PA}/_api/web/lists(guid'{LIST2}')/items"
                f"?$select=Id,Created,{sel}&$orderby=Id desc&$top=1000")
        for it in _sp_get(url2, token).get("value", []):
            if str(it.get("Created") or "") < "2026-08-26":
                continue
            rid = it.get(K["id1"])
            rid = int(float(rid)) if rid is not None else None
            if rid is None or rid in seen:
                continue
            nm = _clean(it.get(K["name"]))
            if not nm or (nq and nq not in _norm(nm)):
                continue
            seen.add(rid)
            out.append(to_cand(it, rid))
    except Exception:
        pass
    out.sort(key=lambda c: (c["id1"] or 0), reverse=True)
    return out


# ============================================================
# 派遣先マスタ + 住所ジオコーディング (国土地理院API・キャッシュ)
# ============================================================
HAKENSAKI_LIST_PATH = "/sites/PowerApps/Lists/040623"
_HK_NAME = "OData__x6d3e__x9063__x5148__x4f1a__x790"   # 派遣先会社名
_HK_PLACE = "OData__x5c31__x696d__x5834__x6240_"        # 就業場所
_HK_ADDR = "OData__x5c31__x696d__x5834__x6240__xff"     # 就業場所(所在地)


def hakensaki_options(token):
    """派遣先ドロップダウン用 [{name, addr}](会社名+住所で重複除去)。
    住所未登録の会社(契約なし・過去取引先等)も addr="" で末尾に含める(アプリ側で住所入力)。"""
    url = f"{SITE_PA}/_api/web/GetList('{HAKENSAKI_LIST_PATH}')/items?$top=500"
    vals = _sp_get(url, token).get("value", [])
    seen = {}
    noaddr = {}
    for it in vals:
        name = _clean(it.get(_HK_NAME)) or _clean(it.get(_HK_PLACE))
        addr = _clean(it.get(_HK_ADDR))
        if not name:
            continue
        if addr:
            seen.setdefault(name + "|" + addr, {"name": name, "addr": addr})
        else:
            noaddr[name] = {"name": name, "addr": ""}
    have = {v["name"] for v in seen.values()}
    out = sorted(seen.values(), key=lambda x: x["name"])
    out += sorted((v for n, v in noaddr.items() if n not in have), key=lambda x: x["name"])
    return out


def save_hakensaki_addr(token, name, addr):
    """住所未登録の派遣先行(会社名一致・所在地空欄)に住所を書き込む。更新行数を返す。"""
    url = f"{SITE_PA}/_api/web/GetList('{HAKENSAKI_LIST_PATH}')/items?$top=500"
    vals = _sp_get(url, token).get("value", [])
    n = 0
    for it in vals:
        nm = _clean(it.get(_HK_NAME)) or _clean(it.get(_HK_PLACE))
        if nm != name or _clean(it.get(_HK_ADDR)):
            continue
        r = requests.post(
            f"{SITE_PA}/_api/web/GetList('{HAKENSAKI_LIST_PATH}')/items({it['Id']})",
            headers={"Authorization": f"Bearer {token}",
                     "Accept": "application/json;odata=nometadata",
                     "Content-Type": "application/json;odata=nometadata",
                     "If-Match": "*", "X-HTTP-Method": "MERGE"},
            json={_HK_ADDR: addr}, timeout=30)
        if r.ok:
            n += 1
    return n


_geo_cache = {}


def geocode(addresses):
    """住所→[lat, lon] (国土地理院 AddressSearch・モジュール内キャッシュ)"""
    out = {}
    for a in list(addresses)[:100]:
        a = str(a or "").strip()
        if not a:
            continue
        if a not in _geo_cache:
            try:
                r = requests.get("https://msearch.gsi.go.jp/address-search/AddressSearch",
                                 params={"q": a}, timeout=10)
                js = r.json()
                if js:
                    lon, lat = js[0]["geometry"]["coordinates"]
                    _geo_cache[a] = [lat, lon]
                else:
                    _geo_cache[a] = None
            except Exception:
                _geo_cache[a] = None
        out[a] = _geo_cache[a]
    return out


def fetch_person(token, id1=None, name=None):
    def scan(base, guid):
        qs = "$orderby=Id desc&$top=1200"
        vals = _sp_get(f"{base}/_api/web/lists(guid'{guid}')/items?{qs}", token).get("value", [])
        for it in vals:
            if id1 is not None:
                v = it.get(K["id1"])
                if v is not None and int(float(v)) == int(id1):
                    return it
            elif name and _norm(name) in _norm(it.get(K["name"])):
                return it
        return None

    r1 = None
    if id1 is not None:
        # List1(<5000件)は$filterで直接取得(過去データも確実にヒット)
        try:
            qs = f"$filter={K['id1']} eq {int(id1)}&$orderby=Id desc&$top=5"
            vals = _sp_get(f"{SP_HOST}/_api/web/lists(guid'{LIST1}')/items?{qs}", token).get("value", [])
            r1 = vals[0] if vals else None
        except Exception:
            r1 = None
    if r1 is None:
        r1 = scan(SP_HOST, LIST1)
    r2 = scan(SITE_PA, LIST2)   # List2は5000件超で$filter不可→直近スキャンのみ(補完用)
    return r1, r2


def _getv(key, r1, r2):
    for r in (r1, r2):
        if r:
            v = _clean(r.get(K[key]))
            if v:
                return v
    return ""


def person_fields(token, r1, r2):
    """アプリの編集フォーム用: 日本語化済みの初期値一式を返す"""
    jobs = jobs_from_rows(token, r1, r2)
    jp = []
    if _getv("jp_rikai", r1, r2):
        jp.append("理解:" + jp_clean(_first(_getv("jp_rikai", r1, r2))))
    if _getv("jp_hanasu", r1, r2):
        jp.append("会話:" + jp_clean(_first(_getv("jp_hanasu", r1, r2))))
    if _getv("jp_pct", r1, r2):
        jp.append("(" + _getv("jp_pct", r1, r2) + ")")
    age = _getv("age", r1, r2)
    uni = _getv("uniform", r1, r2)
    nm = _getv("name", r1, r2)
    edu = ""
    for r in (r1, r2):
        if r and _clean(r.get("Gakureki")):
            edu = _clean(r.get("Gakureki"))   # 新面接シートの学歴列(旧データは空=手入力)
            break
    return {
        "furigana": to_katakana(nm),
        "name": nm,
        "sex": jp_clean(_first(_getv("sex", r1, r2))),
        "age": re.sub(r"[^0-9]", "", age) or age,
        "marital": jp_clean(_first(_getv("marital", r1, r2))),
        "commute": jp_clean(_getv("commute", r1, r2)),
        "addr": addr_short(_getv("addr", r1, r2)),
        "edu": edu,   # 新面接シートは学歴列あり/旧Forms分は空欄(手入力運用)
        "zangyo": jp_clean(_getv("zangyo", r1, r2)),
        "sat": jp_clean(_first(_getv("sat", r1, r2))),
        "hayade": "",
        "kotai": jp_clean(_first(_getv("kotai", r1, r2))),
        "license": jp_clean(_getv("license", r1, r2)) or "＊＊＊",
        "uwagi": uni,
        "pants": uni,
        "shoes": _getv("shoes", r1, r2),
        "jobs": jobs,
        "jp": "　".join(jp),
        "no": "",
    }


# ============================================================
# 描画
# ============================================================
NAVY = (31, 56, 120)
BASE_W, BASE_H = 958, 1493
DPI = 220
A4_W, A4_H = int(8.2633 * DPI), int(11.6933 * DPI)
PT = DPI / 72.0
IMG_W_PT = 501.97
# 左右は中央配置(旧PowerAppsは左寄り11ptだった→2026-08-25ユーザー指示で中央へ)
OFF_X, OFF_Y = ((8.2633 * 72 - 501.97) / 2.0) * PT, 29.87 * PT
SCALE = IMG_W_PT * PT / BASE_W


def _font(size):
    px = max(10, int(round(size * SCALE)))
    for path in FONT_CANDIDATES:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, px)
            except Exception:
                continue
    return ImageFont.load_default()


def _draw_stamp(d, surname):
    """備考欄右下の作成者印枠(base座標: 中心912,1436)に赤の認印を描く"""
    if not surname:
        return
    RED = (198, 40, 40)
    cx, cy = OFF_X + 912 * SCALE, OFF_Y + 1436 * SCALE
    r = 30 * SCALE
    d.ellipse([cx - r, cy - r, cx + r, cy + r], outline=RED, width=max(2, int(2 * SCALE)))
    chars = list(str(surname)[:3])
    if len(chars) == 1:
        d.text((cx, cy), chars[0], font=_font(34), fill=RED, anchor="mm")
    elif len(chars) == 2:
        d.text((cx, cy - 13 * SCALE), chars[0], font=_font(25), fill=RED, anchor="mm")
        d.text((cx, cy + 13 * SCALE), chars[1], font=_font(25), fill=RED, anchor="mm")
    else:
        for i, ch in enumerate(chars):
            d.text((cx, cy + (i - 1) * 18 * SCALE), ch, font=_font(17), fill=RED, anchor="mm")


def render_pdf_fields(F, stamp=None):
    """編集フォームの値(fields辞書)からPDFを描画。stamp=作成者の苗字(認印)"""
    img = Image.open(BLANK).convert("RGB")
    canvas = Image.new("RGB", (A4_W, A4_H), "white")
    bg = img.resize((int(BASE_W * SCALE), int(BASE_H * SCALE)), Image.LANCZOS)
    canvas.paste(bg, (int(OFF_X), int(OFF_Y)))
    d = ImageDraw.Draw(canvas)

    def put(text, x, y, size=21, anchor="lm", max_w=None, bold=True):
        text = str(text or "").strip()
        if not text:
            return
        f = _font(size)
        while max_w and d.textlength(text, font=f) > max_w * SCALE and size > 11:
            size -= 1
            f = _font(size)
        d.text((OFF_X + x * SCALE, OFF_Y + y * SCALE), text, font=f, fill=NAVY,
               anchor=anchor, stroke_width=(1 if bold else 0), stroke_fill=NAVY)

    g = lambda k: str(F.get(k) or "").strip()
    put(datetime.datetime.utcnow().strftime("%Y/%m/%d"), 934, 103, 24, "rm", bold=False)
    put(g("no"), 115, 96, 20, "mm")
    put(g("furigana"), 154, 214, 17, max_w=430)
    put(g("name"), 154, 262, 21, max_w=430)
    put(g("sex"), 663, 245, 21, "mm")
    put(g("age"), 887, 246, 23, "mm")
    put(g("marital"), 179, 329, 21, max_w=280)
    put(g("commute"), 702, 329, 21, "mm", max_w=420)
    put(g("addr"), 154, 411, 21, max_w=770)
    put(g("edu"), 161, 514, 21, max_w=760)
    # 2026-08-25 行改修: 残業|土曜日(休日出勤)|交代勤務 の3項目(早出は廃止)
    put(g("zangyo"), 154, 610, 20, max_w=155)
    put(g("sat"), 551, 608, 21, "mm", max_w=165)
    put(g("kotai"), 869, 608, 21, "mm", max_w=165)
    put(g("license"), 170, 707, 21, max_w=470)
    put(g("uwagi"), 897, 675, 20, "mm", max_w=110)
    put(g("pants"), 897, 707, 20, "mm", max_w=110)
    put(g("shoes"), 897, 740, 20, "mm", max_w=110)

    jobs = F.get("jobs") or []
    COMP_W = 195   # 会社名列の幅(枠内に収める)
    for i, jb in enumerate(jobs[:5]):
        y = 871 + i * 99
        comp = str(jb.get("company") or "").strip()
        if comp:
            if d.textlength(comp, font=_font(21)) <= COMP_W * SCALE:
                put(comp, 14, y, 21)
            elif d.textlength(comp, font=_font(14)) <= COMP_W * SCALE:
                put(comp, 14, y, 14)
            else:
                mid = len(comp) // 2
                sps = [j for j, ch in enumerate(comp) if ch in " 　・-"]
                cut = min(sps, key=lambda j: abs(j - mid)) if sps else mid
                l1 = comp[:cut].strip(" 　・-")
                l2 = comp[cut:].strip(" 　・-")
                put(l1, 14, y - 12, 15, max_w=COMP_W)
                put(l2, 14, y + 12, 15, max_w=COMP_W)
        put((jb.get("naiyou") or ""), 333, y, 21, max_w=500)
        put((jb.get("kikan") or ""), 895, y, 20, "mm", max_w=95)

    put(g("jp"), 310, 1370, 19, max_w=530)   # 備考1行目「日本語能力：」の右
    _draw_stamp(d, stamp)

    buf = io.BytesIO()
    canvas.save(buf, "PDF", resolution=DPI)
    return buf.getvalue()


def render_pdf(token, r1, r2, no=""):
    """後方互換: 行データから直接PDF"""
    f = person_fields(token, r1, r2)
    if no:
        f["no"] = no
    return render_pdf_fields(f)
