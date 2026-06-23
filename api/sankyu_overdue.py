# -*- coding: utf-8 -*-
"""産休育休「注（申請遅れ）」のサーバ側計算。
yukyu-app(index.html)の fetchSankyuSummary / _sankyuBuildCards / classifySankyuField を
忠実に移植したもの。アプリの「注N」と一致させることが目的。

公開API:
  compute_overdue(items, fields_meta, today=None) -> (total:int, details:list)
    items       : SP verbose の d.results (各レコード dict。値キーは fields_meta[*]['dataKey'])
    fields_meta : { InternalName: {title,type,dataKey} }  (fetchSankyuFields相当)
    返り: total = 全レコード合計の遅れ件数, details = [{name, shainNo, procedures:[...]}]
"""
import re
import datetime as _dt

SUBMIT_RE = re.compile(r'(申請日|実施日|提出日|完了日|記入日|実行日|渡した日|もらった日|受取日|受領日)')
START_RE  = re.compile(r'(開始|始め|はじめ|始|From|from|Start|start|から)')
END_RE    = re.compile(r'(終了|終わり|おわり|終|末|To|to|End|end|まで)')


def _norm_title(title):
    if not title:
        return ''
    t = str(title)
    t = ''.join(chr(ord(c) - 0xFF10 + 0x30) if '０' <= c <= '９' else c for c in t)  # 全角数字→半角
    t = re.sub(r'[\s　]+', '', t)
    t = re.sub(r'[\(（\)）]', '', t)
    return t


def _extract_prefix_num(title):
    if not title:
        return None
    t = _norm_title(title).replace('初回', '1回目')
    m = re.match(r'^(.*?)(\d+)(?:回目)?', t)
    if not m:
        return None
    prefix = m.group(1)
    try:
        num = int(m.group(2))
    except Exception:
        return None
    prefix = re.sub(r'[のの・\-ー、,]+$', '', prefix).strip()
    if not prefix:
        return None
    return (prefix, num)


def _extract_base(title, marker_re):
    if not title:
        return ''
    t = str(title)
    last_idx = -1
    for m in marker_re.finditer(t):
        last_idx = m.start()
    if last_idx <= 0:
        return t
    base = t[:last_idx]
    base = re.sub(r'[\s　のの\(（\)）:：・\-ー、,]+$', '', base)
    return base.strip()


def _title_role(title):
    """末尾に最も近いマーカーで 'start'|'end'|'submit'|'none' を返す。"""
    if not title:
        return 'none'
    t = str(title)

    def last_hit(rx):
        last = -1
        for m in rx.finditer(t):
            last = m.start() + len(m.group(0))
        return last

    s_e = last_hit(SUBMIT_RE)
    e_e = last_hit(END_RE)
    st_e = last_hit(START_RE)
    best, best_idx = 'none', -1
    if s_e > best_idx:
        best_idx = s_e; best = 'submit'
    if e_e > best_idx:
        best_idx = e_e; best = 'end'
    if st_e > best_idx:
        best_idx = st_e; best = 'start'
    return best


_JST = _dt.timezone(_dt.timedelta(hours=9))


def _parse_date(raw):
    if raw is None or raw == '':
        return None
    try:
        s = str(raw).replace('Z', '+00:00')
        d = _dt.datetime.fromisoformat(s)
        if d.tzinfo is not None:        # SPはUTC(T15:00:00Z=JST翌0時)。JST実時刻に変換して日付比較
            d = d.astimezone(_JST).replace(tzinfo=None)
        return d
    except Exception:
        try:
            m = re.match(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', str(raw))
            if m:
                return _dt.datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except Exception:
            pass
    return None


def _is_empty_date(raw):
    if raw is None or raw == '':
        return True
    d = _parse_date(raw)
    if d is None:
        return True
    y, mo = d.year, d.month  # Python month 1-12 (JS getMonth()==11 = 12月)
    if y < 2002 and mo == 12:
        return True
    if y == 1900 or y == 1970:
        return True
    return False


def classify_field(title):
    t = str(title or '')
    if not t:
        return 'misc'
    if re.match(r'^状況$', t):
        return 'hidden'
    if re.match(r'^状況\s*2$', t) or re.match(r'^状況2$', t):
        return 'hidden'
    if re.match(r'^(社員番号|氏名|社員氏名|派遣先)$', t):
        return 'hidden'
    if re.match(r'^(社員フォルダ|社員フォルダー|フォルダ|フォルダー)$', t):
        return 'hidden'
    if re.match(r'^(部課|部署|部課名)$', t):
        return 'hidden'
    if '出産予定日記載母子手帳' in t:
        return 'hidden'
    if re.match(r'^Title$', t, re.I):
        return 'hidden'
    if re.search(r'色タグ|カラータグ|Color\s*Tag', t, re.I):
        return 'hidden'
    if 'コンプライアンス資産' in t:
        return 'hidden'
    if re.search(r'レコードとして登録されているアイテム|Record.*Item|記録アイテム', t):
        return 'hidden'
    if '雇用保険資格取得日' in t:
        return 'hidden'
    if re.match(r'^出産予定日$', t):
        return 'basic'
    if re.match(r'^出産日$', t):
        return 'basic'
    if re.search(r'性別.*[(（]\s*子\s*[)）]', t):
        return 'child'
    if re.search(r'性別.*[(（]\s*親\s*[)）]', t):
        return 'basic'
    if '性別' in t:
        return 'basic'
    if re.search(r'子供|子の', t):
        return 'child'
    if '住所' in t:
        return 'basic'
    if re.search(r'電話|ＴＥＬ|TEL|tel', t):
        return 'basic'
    if '生年月日' in t:
        return 'basic'
    if re.search(r'銀行|支店|口座', t):
        return 'bank'
    if re.search(r'健保|健康保険|被保険者|整理番号|雇用保険', t):
        return 'insurance'
    if re.search(r'母子手帳|説明書|産前産後|出産一時金|出産手当金|社保免除期間変更（出産後）', t):
        return 'sango'
    if re.search(r'育休|社保免除期間変更', t):
        return 'ikukyu'
    if '延長' in t:
        return 'choki'
    if re.search(r'保育園申請依頼.*申請日', t):
        return 'hidden'
    if re.search(r'保育園申請依頼.*担当者.*[１1２2]?$', t) and not re.search(r'ＰＤＦ|PDF', t):
        return 'hidden'
    if re.search(r'保育園|仕事復帰', t):
        return 'fukki'
    if '備考' in t:
        return 'memo'
    return 'basic'


class _F(object):
    __slots__ = ('title', 'raw', 'type')

    def __init__(self, title, raw, ftype):
        self.title = title
        self.raw = raw
        self.type = ftype


def _build_cards(fields):
    """fields: list[_F] → cards: list[dict(type,title,start?,end?,main?,submitted?)]。index.htmlの_sankyuBuildCards移植。"""
    arr = list(fields)
    n = len(arr)
    used = [False] * n
    cards = []

    def find_prefixed_sibling(base_title, wanted_role, exclude):
        n_base = _norm_title(base_title)
        for j in range(n):
            if used[j] or j in exclude:
                continue
            g = arr[j]
            if g.type != 'DateTime' or not g.title:
                continue
            if not _norm_title(g.title).startswith(n_base):
                continue
            if _title_role(g.title) != wanted_role:
                continue
            return j
        return -1

    # 1st pass: prefix-based 開始+終了
    for i in range(n):
        if used[i]:
            continue
        f = arr[i]
        if f.type != 'DateTime' or not f.title:
            continue
        if _title_role(f.title) != 'start':
            continue
        base = _extract_base(f.title, START_RE)
        if not base or len(base) < 2:
            continue
        end_idx = find_prefixed_sibling(base, 'end', [i])
        if end_idx < 0:
            continue
        sub_idx = find_prefixed_sibling(base, 'submit', [i, end_idx])
        cards.append({'type': 'period', 'title': base, 'start': f, 'end': arr[end_idx],
                      'submitted': arr[sub_idx] if sub_idx >= 0 else None, 'main': None})
        used[i] = used[end_idx] = True
        if sub_idx >= 0:
            used[sub_idx] = True

    # 1.5th pass: number-based grouping
    num_groups = {}
    for p in range(n):
        if used[p]:
            continue
        fp = arr[p]
        if fp.type != 'DateTime' or not fp.title:
            continue
        pn = _extract_prefix_num(fp.title)
        if not pn:
            continue
        key = '%s|%d' % (pn[0], pn[1])
        g = num_groups.setdefault(key, {'startIdx': -1, 'endIdx': -1, 'subIdx': -1, 'prefix': pn[0], 'num': pn[1]})
        role = _title_role(fp.title)
        if role == 'start' and g['startIdx'] < 0:
            g['startIdx'] = p
        elif role == 'end' and g['endIdx'] < 0:
            g['endIdx'] = p
        elif role == 'submit' and g['subIdx'] < 0:
            g['subIdx'] = p
    for key, grp in num_groups.items():
        if grp['startIdx'] >= 0 and grp['endIdx'] >= 0:
            if used[grp['startIdx']] or used[grp['endIdx']]:
                continue
            cards.append({'type': 'period', 'title': '%s%d回目' % (grp['prefix'], grp['num']),
                          'start': arr[grp['startIdx']], 'end': arr[grp['endIdx']],
                          'submitted': arr[grp['subIdx']] if (grp['subIdx'] >= 0 and not used[grp['subIdx']]) else None,
                          'main': None})
            used[grp['startIdx']] = used[grp['endIdx']] = True
            if grp['subIdx'] >= 0:
                used[grp['subIdx']] = True

    # 1.7th pass: orphan submit を既存カードへ
    for s in range(n):
        if used[s]:
            continue
        sf = arr[s]
        if sf.type != 'DateTime' or _title_role(sf.title) != 'submit':
            continue
        spn = _extract_prefix_num(sf.title)
        if not spn:
            continue
        for card in cards:
            if card.get('submitted'):
                continue
            rep = card.get('start') or card.get('main')
            if not rep or not rep.title:
                continue
            rpn = _extract_prefix_num(rep.title)
            if not rpn:
                continue
            if rpn[0] == spn[0] and rpn[1] == spn[1]:
                card['submitted'] = sf
                used[s] = True
                break

    # 1.9th pass: 隣接順序ベース
    for adj in range(n):
        if used[adj]:
            continue
        af = arr[adj]
        if af.type != 'DateTime' or not af.title:
            continue
        if _title_role(af.title) != 'start':
            continue
        adj_end = adj_sub = -1
        look = adj + 1
        while look < n:
            if not used[look]:
                lf = arr[look]
                if lf.type == 'DateTime' and lf.title:
                    lrole = _title_role(lf.title)
                    if lrole == 'start':
                        break
                    if lrole == 'end' and adj_end < 0:
                        adj_end = look
                    elif lrole == 'submit' and adj_sub < 0:
                        adj_sub = look
            look += 1
        if adj_end < 0:
            continue
        base = _extract_base(af.title, START_RE) or af.title
        cards.append({'type': 'period', 'title': base, 'start': af, 'end': arr[adj_end],
                      'submitted': arr[adj_sub] if adj_sub >= 0 else None, 'main': None})
        used[adj] = used[adj_end] = True
        if adj_sub >= 0:
            used[adj_sub] = True

    # 2nd pass: 単体カード
    for k in range(n):
        if used[k]:
            continue
        g = arr[k]
        if not g.title:
            cards.append({'type': 'single', 'title': g.title, 'main': g, 'submitted': None})
            used[k] = True
            continue
        g_role = _title_role(g.title)
        is_submit_only = (g.type == 'DateTime' and g_role == 'submit')
        if not is_submit_only and g.type == 'DateTime':
            base2 = re.sub(r'[\s　のの\(（\)）:：・\-ー、,]+$', '', g.title).strip()
            sub2 = -1
            if base2 and len(base2) >= 2:
                sub2 = find_prefixed_sibling(base2, 'submit', [k])
            cards.append({'type': 'single', 'title': g.title, 'main': g,
                          'submitted': arr[sub2] if sub2 >= 0 else None})
            used[k] = True
            if sub2 >= 0:
                used[sub2] = True
            continue
        # それ以外(submitのみ/非Date)は単体カード(申請日扱いの遅れ判定対象外)
        cards.append({'type': 'single', 'title': g.title, 'main': g, 'submitted': None})
        used[k] = True

    return cards


def _card_overdue(card, yesterday):
    """index.html fetchSankyuSummary の遅れ判定と同じ: submitted有り&空欄 かつ 対象日付<=昨日。"""
    sub = card.get('submitted')
    if not sub:
        return False
    if not _is_empty_date(sub.raw):
        return False
    past = False
    if card.get('type') == 'period':
        end = card.get('end')
        if end and not _is_empty_date(end.raw):
            d = _parse_date(end.raw)
            if d and d.date() <= yesterday:
                past = True
        if not past:
            start = card.get('start')
            if start and not _is_empty_date(start.raw):
                d = _parse_date(start.raw)
                if d and d.date() <= yesterday:
                    past = True
    else:
        main = card.get('main')
        if main and main.type == 'DateTime' and not _is_empty_date(main.raw):
            d = _parse_date(main.raw)
            if d and d.date() <= yesterday:
                past = True
    return past


def compute_overdue(items, fields_meta, today=None):
    """全レコード合計の遅れ件数と内訳を返す。"""
    if today is None:
        today = (_dt.datetime.utcnow() + _dt.timedelta(hours=9)).date()
    yesterday = today - _dt.timedelta(days=1)

    # 状況/氏名/社員番号 の dataKey を探す
    status_key = name_key = shain_key = None
    for f in fields_meta.values():
        ttl = f.get('title') or ''
        if ttl == '状況' and not status_key:
            status_key = f.get('dataKey')
        if ttl in ('氏名', '社員氏名') and not name_key:
            name_key = f.get('dataKey')
        if ttl == '社員番号' and not shain_key:
            shain_key = f.get('dataKey')

    total = 0
    details = []
    for it in items:
        s = str((status_key and it.get(status_key)) or '')
        if ('終了' in s) or ('完了' in s):
            continue
        all_fields = []
        for f in fields_meta.values():
            cat = classify_field(f.get('title'))
            if cat not in ('sango', 'ikukyu', 'choki', 'fukki'):
                continue
            all_fields.append(_F(f.get('title') or '', it.get(f.get('dataKey')), f.get('type')))
        cards = _build_cards(all_fields)
        procs = [c.get('title') for c in cards if _card_overdue(c, yesterday)]
        if procs:
            total += len(procs)
            details.append({
                'name': (name_key and it.get(name_key)) or (it.get('Title') or ''),
                'shainNo': (shain_key and it.get(shain_key)) or '',
                'count': len(procs),
                'procedures': procs,
            })
    return total, details


# ===== ローカル検証 (standalone) =====
if __name__ == '__main__':
    import sys, io, json, urllib.request
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.path.insert(0, r'C:/Users/step1/.claude/skills/rishokuhyo-create/scripts')
    from sharepoint_yukyu import get_access_token, SITE_URL
    tok = get_access_token()
    LIST = "/sites/PowerApps/Lists/List52"

    def g(url):
        r = urllib.request.Request(url, headers={'Authorization': 'Bearer ' + tok, 'Accept': 'application/json;odata=verbose'})
        return json.loads(urllib.request.urlopen(r).read().decode('utf-8'))['d']['results']

    flds = g(SITE_URL + "/_api/web/GetList('" + LIST + "')/fields?$filter=Hidden%20eq%20false&$select=InternalName,EntityPropertyName,Title,TypeAsString&$top=300")
    SKIP = {'ContentType', 'Modified', 'Created', 'Author', 'Editor', '_UIVersionString', 'Attachments', 'Edit',
            'LinkTitleNoMenu', 'LinkTitle', 'DocIcon', 'ItemChildCount', 'FolderChildCount'}
    meta = {}
    for f in flds:
        iname = f['InternalName']
        if iname in SKIP:
            continue
        if iname in ('Title', 'Id', 'ID'):
            dk = iname
        elif iname.startswith('_'):
            dk = 'OData_' + iname
        else:
            dk = iname
        meta[iname] = {'title': f.get('Title'), 'type': f.get('TypeAsString'), 'dataKey': dk}

    items = g(SITE_URL + "/_api/web/GetList('" + LIST + "')/items?$top=500")
    total, details = compute_overdue(items, meta)
    print('=== 産休育休 注(申請遅れ) 合計 = %d ===' % total)
    for d in details:
        print('  %s (No.%s): %d件 -> %s' % (d['name'], d['shainNo'], d['count'], ' / '.join(d['procedures'])))
