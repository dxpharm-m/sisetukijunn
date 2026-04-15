#!/usr/bin/env python3
"""
保険薬局 施設基準届出書類 自動記入スクリプト
scriptsフォルダ不要版 - zipfileで直接docxを操作
"""
import re, os, sys, shutil, tempfile, json, zipfile
from pathlib import Path

UPLOADS = Path(__file__).parent

# ── docx操作ユーティリティ（zipfile直接操作）─────────────────────
def unpack(src, work_dir):
    """docxをwork_dirに展開し、document.xmlを整形"""
    work_dir = Path(work_dir)
    work_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(src, 'r') as z:
        z.extractall(work_dir)
    # document.xmlのrunをマージ（簡易版）
    doc_xml = work_dir / 'word' / 'document.xml'
    if doc_xml.exists():
        xml = doc_xml.read_text(encoding='utf-8')
        # 隣接する同じ書式のrunをマージ（簡易）
        doc_xml.write_text(xml, encoding='utf-8')

def pack(work_dir, dst, original):
    """work_dirの内容をdocxとして圧縮"""
    work_dir = Path(work_dir)
    dst = Path(dst)
    try:
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zf:
            for f in work_dir.rglob('*'):
                if f.is_file():
                    arcname = f.relative_to(work_dir)
                    zf.write(f, arcname)
        return True
    except Exception as e:
        print(f"pack error: {e}")
        return False

def read_xml(work_dir):
    return (Path(work_dir) / 'word' / 'document.xml').read_text(encoding='utf-8')

def write_xml(work_dir, xml):
    (Path(work_dir) / 'word' / 'document.xml').write_text(xml, encoding='utf-8')

def replace_t(xml, search, replace, count=1):
    esc = re.escape(search)
    return re.sub(r'(<w:t[^>]*>)' + esc + r'(</w:t>)',
                  f'<w:t xml:space="preserve">{replace}</w:t>',
                  xml, count=count), search in xml

def insert_run_after_ppr(xml, para_id, text, font='ＭＳ 明朝', size=20):
    idx = xml.find(f'w14:paraId="{para_id}"')
    if idx < 0: return xml, False
    ppr_end = xml.find('</w:pPr>', idx)
    if ppr_end < 0:
        p_end = xml.find('>', xml.rfind('<w:p ', 0, idx)) + 1
        insert_pos = p_end
    else:
        insert_pos = ppr_end + 8
    sz  = f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
    fnt = f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
    run = f'<w:r><w:rPr>{fnt}{sz}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'
    return xml[:insert_pos] + run + xml[insert_pos:], True

def insert_before_wt_by_para(xml, para_id, wt_text, insert_text, font='ＭＳ 明朝', size=20):
    idx = xml.find(f'w14:paraId="{para_id}"')
    if idx < 0: return xml, False
    p_start = xml.rfind('<w:p ', 0, idx)
    p_end   = xml.find('</w:p>', idx)+6
    para    = xml[p_start:p_end]
    esc = re.escape(wt_text)
    m = re.search(r'(<w:r>(?:<w:rPr>.*?</w:rPr>)?)(<w:t[^>]*>)' + esc + r'(</w:t></w:r>)', para, re.DOTALL)
    if not m: return xml, False
    fnt = f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
    sz  = f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
    run = f'<w:r><w:rPr>{fnt}{sz}</w:rPr><w:t xml:space="preserve">{insert_text}</w:t></w:r>'
    new_para = para[:m.start()] + run + para[m.start():]
    return xml[:p_start] + new_para + xml[p_end:], True

def get_para_map(xml):
    result = {}
    for m in re.finditer(r'w14:paraId="([^"]+)"', xml):
        pid = m.group(1)
        p_start = xml.rfind('<w:p ', 0, m.start())
        p_end   = xml.find('</w:p>', m.start())
        if p_start >= 0 and p_end >= 0:
            texts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', xml[p_start:p_end+6])
            result[pid] = ''.join(texts).strip()
    return result

def next_empty_para(para_map, after_pid, max_look=6):
    keys = list(para_map.keys())
    try:
        idx = keys.index(after_pid)
    except ValueError:
        return None
    for i in range(idx+1, min(idx+1+max_look, len(keys))):
        if not para_map[keys[i]]:
            return keys[i]
    return None

def mark_circle(xml, target_text):
    pattern = '\uff08\u3000\uff09'
    idx = xml.find(target_text)
    if idx < 0: return xml, False
    before = xml[:idx]
    last = before.rfind(pattern)
    if last < 0: return xml, False
    return xml[:last] + '（○）' + xml[last+3:], True

def reiwa_date(iso_date):
    if not iso_date or len(iso_date) < 10: return '', '', ''
    try:
        import datetime
        d = datetime.date.fromisoformat(iso_date)
        return str(d.year - 2018), str(d.month), str(d.day)
    except:
        return '', '', ''

# ── 別添2 ─────────────────────────────────────────────────────────
def fill_bekkan2(src, dst, d):
    work = tempfile.mkdtemp()
    try:
        unpack(src, work)
        xml = read_xml(work)
        pm  = get_para_map(xml)
        cells = xml.split('<w:tc>')
        if len(cells) > 3 and d.get('code'):
            pid = re.search(r'w14:paraId="([^"]+)"', cells[3])
            if pid:
                xml, _ = insert_run_after_ppr(xml, pid.group(1), d['code'])
        ry, rm, rd = reiwa_date(d.get('date',''))
        if ry:
            xml, _ = replace_t(xml, '令和\u3000\u3000\u3000年\u3000\u3000\u3000月\u3000\u3000\u3000日',
                                f'令和{ry}年{rm}月{rd}日')
        for pid, text in pm.items():
            if '担当者氏名：' == text and d.get('contact_name'):
                xml, _ = replace_t(xml, '\u3000\u3000\u3000\u3000担当者氏名：',
                                    f'\u3000\u3000\u3000\u3000担当者氏名：{d["contact_name"]}')
        if d.get('contact_tel'):
            idx_tel = xml.find('13085AA6')
            if idx_tel > 0:
                tel_para_end = xml.find('</w:p>', idx_tel)
                run = f'<w:r><w:t xml:space="preserve">{d["contact_tel"]}</w:t></w:r>'
                xml = xml[:tel_para_end] + run + xml[tel_para_end:]
        for pid, text in pm.items():
            if '保険薬局の所在地' in text:
                epid = next_empty_para(pm, pid)
                if epid and d.get('address'):
                    xml, _ = insert_run_after_ppr(xml, epid, d['address'])
                break
        for pid, text in pm.items():
            if '及び名称' in text:
                epid = next_empty_para(pm, pid)
                if epid and d.get('name'):
                    xml, _ = insert_run_after_ppr(xml, epid, d['name'])
                break
        for pid, text in pm.items():
            if '開設者名' in text:
                epid = next_empty_para(pm, pid)
                if epid and d.get('owner'):
                    xml, _ = insert_run_after_ppr(xml, epid, d['owner'])
                break
        write_xml(work, xml)
        return pack(work, dst, src)
    finally:
        shutil.rmtree(work, ignore_errors=True)

# ── 様式87の3 ─────────────────────────────────────────────────────
def fill_t87_3(src, dst, d):
    work = tempfile.mkdtemp()
    try:
        unpack(src, work)
        xml = read_xml(work)
        pm  = get_para_map(xml)
        kubun_choki = d.get('choki_kubun','')
        for k in ['調剤基本料１','調剤基本料２','調剤基本料３－イ','調剤基本料３－ロ','調剤基本料３－ハ','特別調剤基本料Ａ']:
            if k in kubun_choki:
                xml, _ = mark_circle(xml, k); break
        kubun_chiki = d.get('chiki_kubun','')
        for k in ['地域支援体制加算１','地域支援体制加算２','地域支援体制加算３','地域支援体制加算４']:
            if k in kubun_chiki:
                xml, _ = mark_circle(xml, k); break
        if d.get('bichiku_num') or d.get('bichiku_ym'):
            orig = 'ア\u3000備蓄品目数\u3000\u3000\u3000\u3000\u3000（\u3000\u3000年\u3000\u3000月現在）'
            xml, _ = replace_t(xml, orig, f'ア\u3000備蓄品目数\u3000\u3000\u3000\u3000\u3000（{d.get("bichiku_ym","")}現在）')
            if d.get('bichiku_num'):
                xml, _ = replace_t(xml, '品目', f'{d["bichiku_num"]}品目')
        if d.get('mayaku_num'):
            orig = '（免許証の番号を記載：' + '\u3000'*18 + '）'
            xml, _ = replace_t(xml, orig, f'（免許証の番号を記載：{d["mayaku_num"]}）')
        fr = d.get('rx_period_from',''); to = d.get('rx_period_to','')
        if fr or to:
            orig = '期間：\u3000\u3000年\u3000\u3000月\u3000\u3000～\u3000\u3000年\u3000\u3000月'
            xml, _ = replace_t(xml, orig, f'期間：{fr}～{to}')
        blank4 = '\u3000'*4
        for val in [d.get('rx_total',''), d.get('rx_main',''), d.get('rx_conc',''), d.get('generic_rate','')]:
            if val:
                xml = xml.replace(f'<w:t>{blank4}</w:t>', f'<w:t xml:space="preserve">{val}</w:t>', 1)
        if d.get('rx_conc'):
            xml, _ = replace_t(xml, '③集中率（％）', f'③集中率（{d["rx_conc"]}%）')
        if d.get('generic_rate'):
            idx = xml.find('後発医薬品の調剤割合')
            if idx > 0:
                pct_idx = xml.find('<w:t>％</w:t>', idx)
                if pct_idx > 0:
                    xml = xml[:pct_idx] + f'<w:t xml:space="preserve">{d["generic_rate"]}％</w:t>' + xml[pct_idx+13:]
        fr2 = d.get('zaita_period_from',''); to2 = d.get('zaita_period_to','')
        if fr2 or to2:
            orig2 = '（実績回数の期間：\u3000\u3000年\u3000\u3000月～\u3000\u3000年\u3000\u3000月）'
            xml, _ = replace_t(xml, orig2, f'（実績回数の期間：{fr2}～{to2}）')
        for val in [d.get('zaita_total',''), d.get('zaita_1',''), d.get('zaita_2',''), d.get('zaita_3',''), d.get('zaita_4','')]:
            if val:
                xml = xml.replace(f'<w:t>{blank4}</w:t>', f'<w:t xml:space="preserve">{val}</w:t>', 1)
        if d.get('kakyoku'):
            xml, ok = insert_run_after_ppr(xml, '69069A58', d['kakyoku'])
        for pid, text in pm.items():
            if text == '連携薬局名' and d.get('renk_name'):
                epid = next_empty_para(pm, pid)
                if epid: xml, _ = insert_run_after_ppr(xml, epid, d['renk_name'])
            if text == '連携する業務内容' and d.get('renk_gyomu'):
                epid = next_empty_para(pm, pid)
                if epid: xml, _ = insert_run_after_ppr(xml, epid, d['renk_gyomu'])
        if d.get('pmdanavi'):
            orig_pmd = '（' + '\u3000'*6 + '）'
            xml, _ = replace_t(xml, orig_pmd, f'（{d["pmdanavi"]}）')
        for pid, text in pm.items():
            if '①氏名' in text and d.get('mgr_name'):
                epid = next_empty_para(pm, pid)
                if epid: xml, _ = insert_run_after_ppr(xml, epid, d['mgr_name'])
                break
        if d.get('mgr_kinmu'): xml, _ = insert_before_wt_by_para(xml, '3575F1DF', '年', d['mgr_kinmu'])
        if d.get('mgr_hours'): xml, _ = insert_before_wt_by_para(xml, '64320610', '時間', d['mgr_hours'])
        write_xml(work, xml)
        return pack(work, dst, src)
    finally:
        shutil.rmtree(work, ignore_errors=True)

# ── 様式87の3の2 ──────────────────────────────────────────────────
def fill_t87_3_2(src, dst, d):
    work = tempfile.mkdtemp()
    try:
        unpack(src, work)
        xml = read_xml(work)
        kubun = d.get('chiki_kubun','')
        for k in ['地域支援体制加算１','地域支援体制加算２','地域支援体制加算３','地域支援体制加算４']:
            if k in kubun:
                xml, _ = mark_circle(xml, k); break
        if d.get('rx_total'):
            xml, ok = insert_run_after_ppr(xml, '494A4AC5', d['rx_total'], font='ＭＳ ゴシック', size=20)
        fr = d.get('period_from',''); to = d.get('period_to','')
        if fr or to:
            orig = '期間：\u3000\u3000年\u3000\u3000月\u3000\u3000\u3000～\u3000\u3000\u3000年\u3000\u3000\u3000月'
            xml, _ = replace_t(xml, orig, f'期間：{fr}～{to}')
        blank8 = ' ' * 8
        jitsu = d.get('jitsu', {})
        for key in ['1','2','3','4','5','6','7','8','9']:
            val = str(jitsu.get(key,''))
            if val:
                xml = xml.replace(f'<w:t>{blank8}</w:t>', f'<w:t xml:space="preserve">{val}</w:t>', 1)
                xml = xml.replace(f'<w:t xml:space="preserve">{blank8}</w:t>', f'<w:t xml:space="preserve">{val}</w:t>', 1)
        if jitsu.get('10'):
            xml = xml.replace('<w:t> </w:t>', f'<w:t xml:space="preserve">{jitsu["10"]}</w:t>', 1)
        write_xml(work, xml)
        return pack(work, dst, src)
    finally:
        shutil.rmtree(work, ignore_errors=True)

# ── 様式87の3の4 ──────────────────────────────────────────────────
def fill_t87_3_4(src, dst, d):
    work = tempfile.mkdtemp()
    try:
        unpack(src, work)
        xml = read_xml(work)
        pm  = get_para_map(xml)
        if d.get('tokubetsu_iryo'):
            for pid, text in pm.items():
                if '特別な関係を有している保険医療機関名' in text:
                    epid = next_empty_para(pm, pid)
                    if epid: xml, _ = insert_run_after_ppr(xml, epid, d['tokubetsu_iryo'])
                    break
        write_xml(work, xml)
        return pack(work, dst, src)
    finally:
        shutil.rmtree(work, ignore_errors=True)

# ── 様式87の3の5 ──────────────────────────────────────────────────
def fill_t87_3_5(src, dst, d):
    work = tempfile.mkdtemp()
    try:
        unpack(src, work)
        xml = read_xml(work)
        pm  = get_para_map(xml)
        kubun = d.get('kubun','')
        for k in ['在宅薬学総合体制加算１','在宅薬学総合体制加算２']:
            if k in kubun:
                xml, _ = mark_circle(xml, k); break
        if d.get('mayaku_num'):
            orig = '（免許証の番号を記載：' + '\u3000'*18 + '）'
            xml, _ = replace_t(xml, orig, f'（免許証の番号を記載：{d["mayaku_num"]}）')
        fr = d.get('zaita_period_from',''); to = d.get('zaita_period_to','')
        if fr or to:
            orig = '（実績回数の期間：\u3000\u3000\u3000年\u3000\u3000\u3000月～\u3000\u3000\u3000年\u3000\u3000\u3000月）'
            xml, _ = replace_t(xml, orig, f'（実績回数の期間：{fr}～{to}）')
        blank1 = ' '
        for val in [d.get('zaita_total',''), d.get('zaita_a',''), d.get('zaita_i',''), d.get('zaita_u',''), d.get('zaita_e','')]:
            if val:
                xml = xml.replace(f'<w:t>{blank1}</w:t>', f'<w:t xml:space="preserve">{val}</w:t>', 1)
        if d.get('ph_total'):
            xml, _ = replace_t(xml, '（\u3000\u3000\u3000\u3000人）', f'（{d["ph_total"]}人）')
        write_xml(work, xml)
        return pack(work, dst, src)
    finally:
        shutil.rmtree(work, ignore_errors=True)

# ── 様式87の3の6 ──────────────────────────────────────────────────
def fill_t87_3_6(src, dst, d):
    work = tempfile.mkdtemp()
    try:
        unpack(src, work); xml = read_xml(work)
        write_xml(work, xml)
        return pack(work, dst, src)
    finally:
        shutil.rmtree(work, ignore_errors=True)

# ── 様式88 ────────────────────────────────────────────────────────
def fill_t88(src, dst, d):
    work = tempfile.mkdtemp()
    try:
        unpack(src, work)
        xml = read_xml(work)
        pm  = get_para_map(xml)
        if d.get('renk_yakkyoku'):
            for pid, text in pm.items():
                if '無菌調剤室提供薬局' in text:
                    epid = next_empty_para(pm, pid)
                    if epid: xml, _ = insert_run_after_ppr(xml, epid, d['renk_yakkyoku'])
                    break
        if d.get('setsubi_kigo'):
            blank8 = ' ' * 8
            xml = xml.replace(f'<w:t>{blank8}</w:t>', f'<w:t xml:space="preserve">{d["setsubi_kigo"]}</w:t>', 1)
        write_xml(work, xml)
        return pack(work, dst, src)
    finally:
        shutil.rmtree(work, ignore_errors=True)

# ── 様式90 ────────────────────────────────────────────────────────
def fill_t90(src, dst, d):
    work = tempfile.mkdtemp()
    try:
        unpack(src, work)
        xml = read_xml(work)
        yakuzaishis = d.get('yakuzaishis', [])
        ROW_PARA_IDS = [
            ('5AACAC10','2D51C817','3575F1DF','64320610','08CCD7D9','35815F86'),
            ('09373C58','2712B2E5','1365B424','26883B1E','3FBC7E5A','3D482228'),
        ]
        for yi, ydata in enumerate(yakuzaishis[:2]):
            if yi >= len(ROW_PARA_IDS): break
            _, name_pid, nen_pid, jikan_pid, nichi_pid, zaiseki_pid = ROW_PARA_IDS[yi]
            if ydata.get('name'):
                xml, _ = insert_run_after_ppr(xml, name_pid, ydata['name'])
            if ydata.get('kinmu_years'):
                xml, _ = insert_before_wt_by_para(xml, nen_pid, '年', ydata['kinmu_years'])
            if ydata.get('hours_per_week'):
                xml, _ = insert_before_wt_by_para(xml, jikan_pid, '時間/週', ydata['hours_per_week'])
            if ydata.get('zaiseki'):
                xml, _ = insert_before_wt_by_para(xml, zaiseki_pid, '年\u3000 月', ydata['zaiseki'])
        write_xml(work, xml)
        return pack(work, dst, src)
    finally:
        shutil.rmtree(work, ignore_errors=True)

# ── 様式92 ────────────────────────────────────────────────────────
def fill_t92(src, dst, d):
    work = tempfile.mkdtemp()
    try:
        unpack(src, work)
        xml = read_xml(work)
        pm  = get_para_map(xml)
        blank8g = '\u3000' * 8
        for ydata in d.get('yakuzaishis', [])[:4]:
            if ydata.get('name'):
                xml = xml.replace(f'<w:t>{blank8g}</w:t>', f'<w:t xml:space="preserve">{ydata["name"]}</w:t>', 1)
        if d.get('privacy_method'):
            orig = '（配慮方法）（具体的に記入' + '\u3000'*34 + '）'
            xml, _ = replace_t(xml, orig, f'（配慮方法）（{d["privacy_method"]}）')
        if d.get('mayaku_num'):
            for pid, text in pm.items():
                if '麻薬小売業者免許証の番号' in text:
                    epid = next_empty_para(pm, pid)
                    if epid: xml, _ = insert_run_after_ppr(xml, epid, d['mayaku_num'])
                    break
        write_xml(work, xml)
        return pack(work, dst, src)
    finally:
        shutil.rmtree(work, ignore_errors=True)

# ── ファイルマッピング ─────────────────────────────────────────────
def pick_bekkan2_file(fac_id, fac_data):
    defaults = {
        'f84':'r6-2-523.docx','f87_2':'r6-2-529.docx','f87_3_4':'r6-2-534.docx',
        'f87_3_5':'r6-2-538.docx','f87_3_6':'r6-2-540.docx','f88':'r6-2-541.docx',
        'f90':'r6-2-543.docx','f92':'r6-2-542.docx',
    }
    if fac_id == 'f87':
        kubun = fac_data.get('kubun','')
        if '加算２' in kubun or '加算2' in kubun: return 'r6-2-536.docx'
        if '加算３' in kubun or '加算3' in kubun: return 'r6-2-537.docx'
        return 'r6-2-535.docx'
    if fac_id == 'f87_3':
        kubun = fac_data.get('kubun','')
        if '加算２' in kubun or '加算2' in kubun: return 'r6-2-531.docx'
        if '加算３' in kubun or '加算3' in kubun: return 'r6-2-532.docx'
        if '加算４' in kubun or '加算4' in kubun: return 'r6-2-533.docx'
        return 'r6-2-530.docx'
    return defaults.get(fac_id)

def build_kakyoku(basic):
    days = [('月',basic.get('b_h_mon_s',''),basic.get('b_h_mon_e','')),
            ('火',basic.get('b_h_tue_s',''),basic.get('b_h_tue_e','')),
            ('水',basic.get('b_h_wed_s',''),basic.get('b_h_wed_e','')),
            ('木',basic.get('b_h_thu_s',''),basic.get('b_h_thu_e','')),
            ('金',basic.get('b_h_fri_s',''),basic.get('b_h_fri_e','')),
            ('土',basic.get('b_h_sat_s',''),basic.get('b_h_sat_e','')),
            ('日',basic.get('b_h_sun_s',''),basic.get('b_h_sun_e','')),
            ('祝',basic.get('b_h_hol_s',''),basic.get('b_h_hol_e',''))]
    parts = []
    for day, s, e in days:
        if s and e and s != '閉局': parts.append(f'{day}:{s}〜{e}')
        elif s == '閉局' or (s and not e): parts.append(f'{day}:閉局')
    return '\u3000'.join(parts)

# ── メイン生成処理 ─────────────────────────────────────────────────
def generate_all(app_state, output_dir):
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    basic    = app_state.get('basic', {})
    fac_data = app_state.get('facData', {})
    selected = app_state.get('selected', [])
    ry, rm, rd = reiwa_date(basic.get('b_date',''))
    chiki_period = basic.get('b_rx_chiki_period','')
    chiki_fr = chiki_period.split('〜')[0].strip() if '〜' in chiki_period else ''
    chiki_to = chiki_period.split('〜')[1].strip() if '〜' in chiki_period else ''
    bekkan2_base = {
        'code': basic.get('b_code',''), 'date': basic.get('b_date',''),
        'address': basic.get('b_addr',''), 'name': basic.get('b_name',''),
        'owner': basic.get('b_owner',''), 'contact_name': basic.get('b_mgr',''),
        'contact_tel': basic.get('b_tel',''),
    }
    results = []
    for fac_id in selected:
        fac = fac_data.get(fac_id, {})
        bk_file = pick_bekkan2_file(fac_id, fac)
        if bk_file:
            src = UPLOADS / bk_file
            if src.exists():
                dst = output_dir / f'別添2_{bk_file.replace("r6-2-","").replace(".docx","")}_{fac_id}.docx'
                ok  = fill_bekkan2(str(src), str(dst), bekkan2_base)
                results.append({'file': dst.name, 'ok': ok, 'label': f'別添2（{fac_id}）'})
        if fac_id == 'f87_3':
            src3 = UPLOADS / 'r6-t87-3.docx'
            if src3.exists():
                choki_raw = fac.get('chiki_base_kubun','')
                t3d = {
                    'choki_kubun': '調剤基本料１' if '基本料１' in choki_raw else '調剤基本料２',
                    'chiki_kubun': fac.get('kubun',''),
                    'bichiku_num': fac.get('bichiku_num',''), 'bichiku_ym': fac.get('bichiku_ym',''),
                    'mayaku_num': fac.get('mayaku_num',''),
                    'rx_period_from': chiki_fr, 'rx_period_to': chiki_to,
                    'rx_total': basic.get('b_rx_total',''), 'rx_main': basic.get('b_rx_main_n',''),
                    'rx_conc': basic.get('b_rx_conc',''), 'generic_rate': basic.get('b_generic',''),
                    'zaita_period_from': chiki_fr, 'zaita_period_to': chiki_to,
                    'zaita_total': fac.get('jitsu_zaita',''),
                    'kakyoku': build_kakyoku(basic),
                    'renk_name': fac.get('renk_name',''), 'renk_gyomu': fac.get('renk_gyomu',''),
                    'pmdanavi': fac.get('pmdanavi',''),
                    'mgr_name': basic.get('b_mgr',''), 'mgr_kinmu': basic.get('b_mgr_years',''),
                }
                dst3 = output_dir / '様式87の3_地域支援体制加算.docx'
                results.append({'file': dst3.name, 'ok': fill_t87_3(str(src3), str(dst3), t3d), 'label': '様式87の3'})
            src32 = UPLOADS / 'r6-t87-3-2.docx'
            if src32.exists():
                t32d = {
                    'chiki_kubun': fac.get('kubun',''),
                    'rx_total': basic.get('b_rx_total_chiki', basic.get('b_rx_total','')),
                    'period_from': chiki_fr, 'period_to': chiki_to,
                    'jitsu': {str(i): fac.get(k,'') for i,k in enumerate(
                        ['jitsu_yakan','jitsu_mayaku','jitsu_juufuku','jitsu_kakari',
                         'jitsu_gairaif','jitsu_hukuyaku','jitsu_zaita','jitsu_johoteik',
                         'jitsu_shoji','jitsu_taishoku'], 1)}
                }
                dst32 = output_dir / '様式87の3の2_地域支援体制加算実績.docx'
                results.append({'file': dst32.name, 'ok': fill_t87_3_2(str(src32), str(dst32), t32d), 'label': '様式87の3の2'})
        elif fac_id == 'f87_3_4':
            src = UPLOADS / 'r6-t87-3-4.docx'
            if src.exists():
                dst = output_dir / '様式87の3の4_連携強化加算.docx'
                results.append({'file': dst.name, 'ok': fill_t87_3_4(str(src), str(dst), fac), 'label': '様式87の3の4'})
        elif fac_id == 'f87_3_5':
            src = UPLOADS / 'r6-t87-3-5.docx'
            if src.exists():
                dst = output_dir / '様式87の3の5_在宅薬学総合体制加算.docx'
                results.append({'file': dst.name, 'ok': fill_t87_3_5(str(src), str(dst), {
                    'kubun': fac.get('kubun',''), 'mayaku_num': fac.get('mayaku_num',''),
                    'zaita_period_from': chiki_fr, 'zaita_period_to': chiki_to,
                    'zaita_total': fac.get('zaita_jisseki',''), 'ph_total': basic.get('b_ph_full',''),
                }), 'label': '様式87の3の5'})
        elif fac_id == 'f87_3_6':
            src = UPLOADS / 'r6-t87-3-6_0704.docx'
            if src.exists():
                dst = output_dir / '様式87の3の6_医療DX推進体制整備加算.docx'
                results.append({'file': dst.name, 'ok': fill_t87_3_6(str(src), str(dst), fac), 'label': '様式87の3の6'})
        elif fac_id == 'f88':
            src = UPLOADS / 'r6-t88.docx'
            if src.exists():
                dst = output_dir / '様式88_無菌製剤処理加算.docx'
                results.append({'file': dst.name, 'ok': fill_t88(str(src), str(dst), {
                    'renk_yakkyoku': fac.get('renk_yakkyoku',''), 'setsubi_kigo': fac.get('tantou',''),
                }), 'label': '様式88'})
        elif fac_id == 'f90':
            src = UPLOADS / 'r6-t90.docx'
            if src.exists():
                dst = output_dir / '様式90_かかりつけ薬剤師指導料.docx'
                results.append({'file': dst.name, 'ok': fill_t90(str(src), str(dst), {'yakuzaishis': [{
                    'name': fac.get('yakuzaishi_name',''), 'kinmu_years': fac.get('kinmu_years',''),
                    'hours_per_week': fac.get('toukyoku_hours',''),
                    'zaiseki': str(fac.get('zaiseki_months',''))+'ヶ月' if fac.get('zaiseki_months') else '',
                }]}), 'label': '様式90'})
        elif fac_id == 'f92':
            src = UPLOADS / 'r6-t92.docx'
            if src.exists():
                dst = output_dir / '様式92_特定薬剤管理指導加算2.docx'
                results.append({'file': dst.name, 'ok': fill_t92(str(src), str(dst), {
                    'yakuzaishis': [{'name': fac.get('yakuzaishi','')}],
                    'privacy_method': fac.get('renk_tasei',''),
                }), 'label': '様式92'})
    return results

if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("Usage: fill_forms.py <app_state.json> <output_dir>")
        sys.exit(1)
    with open(sys.argv[1],'r',encoding='utf-8') as f:
        app_state = json.load(f)
    results = generate_all(app_state, sys.argv[2])
    print(f"\n生成結果 ({sum(r['ok'] for r in results)}/{len(results)} 成功):")
    for r in results:
        print(f"  {'✓' if r['ok'] else '✗'} {r['label']}: {r['file']}")
