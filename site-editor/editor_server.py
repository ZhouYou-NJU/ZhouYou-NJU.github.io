#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
课题组主页编辑器 - 本地服务
仅使用 Python 标准库，无需安装任何依赖。
功能：数据保存 / 照片上传 / Word(docx)转HTML / 一键导出静态网站
"""
import base64
import html as html_mod
import io
import json
import os
import re
import shutil
import sys
import time
import zipfile
import xml.etree.ElementTree as ET
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, unquote

BASE = os.path.dirname(os.path.abspath(__file__))      # site-editor 目录
ROOT = os.path.dirname(BASE)                            # C:\个人课题组主页（网站根目录 = GitHub Pages 仓库根目录）
SITE = ROOT                                             # 新架构：直接部署到主页文件夹本身
DATA_FILE = os.path.join(BASE, "data.json")
UPLOADS = os.path.join(BASE, "uploads")
GH_CFG = os.path.join(BASE, "github_config.json")
PORT = 8765

GITIGNORE = """# Homepage Editor
site-editor/github_config.json
site-editor/server_log.txt
site-editor/data.json.bak
site-editor/data.json.tmp
site-editor/__pycache__/
__pycache__/
备份/
Thumbs.db
desktop.ini
"""

# ---------------------------------------------------------------
# Word (docx) -> HTML  转换（纯标准库实现）
# ---------------------------------------------------------------
WNS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
RNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
ANS = '{http://schemas.openxmlformats.org/drawingml/2006/main}'


def _rpr_flag(r):
    """判断 run 的 rPr/b 或 rPr/i 是否生效"""
    rpr = r.find(WNS + 'rPr')
    if rpr is None:
        return False
    for tag in ('b', 'i', 'u'):
        el = rpr.find(WNS + tag)
        if el is not None:
            v = el.get(WNS + 'val')
            if v in (None, '1', 'true', 'on'):
                return tag
    return False


class DocxConverter:
    def __init__(self, docx_bytes, media_dir, media_url_prefix):
        self.z = zipfile.ZipFile(io.BytesIO(docx_bytes))
        self.media_dir = media_dir
        self.media_prefix = media_url_prefix
        self.rels = {}
        try:
            relxml = ET.fromstring(self.z.read('word/_rels/document.xml.rels'))
            for rel in relxml:
                self.rels[rel.get('Id')] = rel.get('Target') or ''
        except Exception:
            pass
        self.img_count = 0

    def _extract_image(self, rid):
        target = self.rels.get(rid, '')
        if not target:
            return None
        name = os.path.basename(target)
        try:
            data = self.z.read('word/' + target.lstrip('/').replace('media/', 'media/'))
        except KeyError:
            try:
                data = self.z.read(target.lstrip('/'))
            except Exception:
                return None
        self.img_count += 1
        ext = os.path.splitext(name)[1] or '.png'
        out_name = 'img_%d%s' % (self.img_count, ext)
        os.makedirs(self.media_dir, exist_ok=True)
        with open(os.path.join(self.media_dir, out_name), 'wb') as f:
            f.write(data)
        return self.media_prefix + '/' + out_name

    def _runs_to_html(self, el):
        """递归把段落内的 run/hyperlink 转为 HTML，保持顺序"""
        out = []
        for child in el:
            tag = child.tag
            if tag == WNS + 'r':
                bold = _rpr_flag(child) == 'b'
                italic = _rpr_flag(child) == 'i'
                under = _rpr_flag(child) == 'u'
                text = ''
                img_html = ''
                for sub in child:
                    if sub.tag == WNS + 't':
                        text += sub.text or ''
                    elif sub.tag == WNS + 'tab':
                        text += '\u3000'
                    elif sub.tag == WNS + 'br':
                        text += '<br>'
                    elif sub.tag in (WNS + 'drawing', WNS + 'pict'):
                        for blip in sub.iter(ANS + 'blip'):
                            rid = blip.get(RNS + 'embed') or blip.get(RNS + 'link')
                            if rid:
                                url = self._extract_image(rid)
                                if url:
                                    img_html += '<img src="%s" style="max-width:100%%;">' % html_mod.escape(url)
                esc = html_mod.escape(text)
                if bold:
                    esc = '<strong>%s</strong>' % esc
                if italic:
                    esc = '<em>%s</em>' % esc
                if under:
                    esc = '<u>%s</u>' % esc
                out.append(esc + img_html)
            elif tag == WNS + 'hyperlink':
                rid = child.get(RNS + 'id')
                inner = ''.join(t.text or '' for t in child.iter(WNS + 't'))
                href = self.rels.get(rid, '')
                if href:
                    out.append('<a href="%s" target="_blank">%s</a>' % (html_mod.escape(href, True), html_mod.escape(inner)))
                else:
                    out.append(html_mod.escape(inner))
        return ''.join(out)

    def _para_style(self, p):
        ppr = p.find(WNS + 'pPr')
        style = ''
        is_list = False
        if ppr is not None:
            ps = ppr.find(WNS + 'pStyle')
            if ps is not None:
                style = ps.get(WNS + 'val') or ''
            if ppr.find(WNS + 'numPr') is not None:
                is_list = True
        return style, is_list

    def convert(self):
        root = ET.fromstring(self.z.read('word/document.xml'))
        body = root.find(WNS + 'body')
        parts = []
        in_list = False

        def close_list():
            nonlocal in_list
            if in_list:
                parts.append('</ul>')
                in_list = False

        for child in body:
            if child.tag == WNS + 'p':
                text = ''.join(t.text or '' for t in child.iter(WNS + 't')).strip()
                has_img = any(True for _ in child.iter(ANS + 'blip'))
                style, is_list = self._para_style(child)
                if not text and not has_img:
                    continue
                sl = style.lower()
                if is_list:
                    if not in_list:
                        parts.append('<ul>')
                        in_list = True
                    parts.append('<li>%s</li>' % self._runs_to_html(child))
                    continue
                close_list()
                if sl.startswith('heading1') or sl in ('1', 'title'):
                    parts.append('<h2>%s</h2>' % self._runs_to_html(child))
                elif sl.startswith('heading2') or sl == '2':
                    parts.append('<h3>%s</h3>' % self._runs_to_html(child))
                elif sl.startswith('heading3') or sl == '3':
                    parts.append('<h4>%s</h4>' % self._runs_to_html(child))
                else:
                    parts.append('<p>%s</p>' % self._runs_to_html(child))
            elif child.tag == WNS + 'tbl':
                close_list()
                rows_html = []
                for tr in child.iter(WNS + 'tr'):
                    cells = []
                    for tc in tr.findall(WNS + 'tc'):
                        cell_text = '<br>'.join(
                            ''.join(t.text or '' for t in p.iter(WNS + 't')).strip()
                            for p in tc.iter(WNS + 'p'))
                        cells.append('<td style="border:1px solid #ccc;padding:4px 8px;">%s</td>' % cell_text)
                    rows_html.append('<tr>%s</tr>' % ''.join(cells))
                parts.append('<table style="border-collapse:collapse;">%s</table>' % ''.join(rows_html))
        close_list()
        return '\n'.join(parts)


# ---------------------------------------------------------------
# 数据与文件工具
# ---------------------------------------------------------------
def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def save_data(data):
    tmp = DATA_FILE + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    if os.path.exists(DATA_FILE):
        shutil.copyfile(DATA_FILE, DATA_FILE + '.bak')
    os.replace(tmp, DATA_FILE)


def resolve_image(path):
    """把数据里的图片相对路径解析为磁盘绝对路径（只读，不改动原目录）"""
    path = path.replace('\\', '/').lstrip('/')
    if path.startswith('uploads/'):
        return os.path.join(BASE, path.replace('/', os.sep))
    if path.startswith('photos/'):
        return os.path.join(ROOT, path.replace('/', os.sep))
    if path.startswith('YouZhou/'):
        return os.path.join(ROOT, path.replace('/', os.sep))
    if re.match(r'^[A-Za-z]:[\\/]', path):
        return path
    return None


def safe_basename(path):
    return os.path.basename(path.replace('\\', '/'))


def collect_site_images(data):
    """收集所有需要复制进网站的图片 -> {源绝对路径: 目标文件名}"""
    mapping = {}

    def add(rel):
        if not rel:
            return
        src = resolve_image(rel)
        if src and os.path.exists(src):
            mapping[rel] = safe_basename(src)

    add(data.get('profile', {}).get('avatar', ''))
    for p in data.get('publications', []):
        add(p.get('image', ''))
    for group in ('current', 'alumni'):
        for m in data.get('members', {}).get(group, []):
            add(m.get('photo', ''))
            html_content = m.get('homepageHtml', '')
            for src in re.findall(r'src="([^"]+)"', html_content):
                add(src)
    return mapping


def member_slug(name, used):
    slug = re.sub(r'[^a-z0-9]+', '-', (name or '').lower()).strip('-')
    if not slug:
        slug = 'member'
    base, i = slug, 2
    while slug in used:
        slug = '%s-%d' % (base, i)
        i += 1
    used.add(slug)
    return slug


CSS = """
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:Georgia,'Times New Roman','Microsoft YaHei',serif;background:#f4f5f7;color:#222;line-height:1.65}
a{color:#1a5dad;text-decoration:none}
a:hover{text-decoration:underline}
.nav{position:sticky;top:0;background:#fff;box-shadow:0 1px 6px rgba(0,0,0,.08);z-index:10}
.nav-inner{max-width:1060px;margin:0 auto;display:flex;justify-content:space-between;align-items:center;padding:14px 24px}
.brand{font-size:22px;font-weight:bold;color:#1a3c6e}
.brand:hover{text-decoration:none}
.nav-links a{margin-left:22px;font-size:15px;color:#444;font-family:Verdana,Arial,sans-serif}
.nav-links a.active{color:#1a5dad;font-weight:bold;border-bottom:2px solid #1a5dad;padding-bottom:3px}
.wrap{max-width:1060px;margin:28px auto 60px;padding:0 24px}
.card{background:#fff;border-radius:10px;box-shadow:0 1px 4px rgba(0,0,0,.06);padding:32px 38px;margin-bottom:26px}
h1.page-title{font-size:30px;color:#1a3c6e;margin-bottom:6px}
h2.sec{font-size:21px;color:#1a3c6e;border-left:4px solid #1a5dad;padding-left:12px;margin-bottom:16px}
h3.year{font-size:17px;margin:18px 0 8px;color:#333;font-family:Verdana,sans-serif}
.hero{display:flex;gap:34px;align-items:center;flex-wrap:wrap}
.hero img{width:210px;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,.15)}
.hero .info h1{font-size:34px;color:#1a3c6e}
.hero .info .role{font-size:20px;margin:4px 0 10px;color:#333}
.hero .info .meta{font-size:15px;color:#444;font-family:Verdana,sans-serif;line-height:1.9}
.badge{display:inline-block;background:#e8f0fb;color:#1a5dad;border-radius:4px;padding:1px 8px;font-size:12px;font-family:Verdana,sans-serif;margin-left:8px}
.pub{margin:10px 0;padding-left:14px;border-left:2px solid #e3e8ef}
.pub .t{font-weight:bold;color:#1a5dad}
.pub .v{font-style:italic}
.pub .note{color:#c0392b;font-weight:bold;font-size:13px}
.mgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(190px,1fr));gap:22px}
.mcard{text-align:center;background:#fafbfc;border:1px solid #e8ecf1;border-radius:10px;padding:18px 12px;transition:box-shadow .2s}
.mcard:hover{box-shadow:0 3px 12px rgba(0,0,0,.1)}
.mcard img{width:104px;height:130px;object-fit:cover;border-radius:8px}
.mcard .nm{font-weight:bold;font-size:15px;margin-top:10px}
.mcard .dg{font-size:13px;color:#666;margin-top:3px;font-family:Verdana,sans-serif}
.mcard .pd{font-size:12.5px;color:#888;margin-top:2px;font-family:Verdana,sans-serif}
ul.honors{list-style:none}
ul.honors li{padding:9px 6px 9px 30px;border-bottom:1px dashed #e5e9ef;position:relative}
ul.honors li:before{content:'🏆';position:absolute;left:2px;font-size:14px}
.codes li{margin:10px 0}
.edu li{margin:9px 0}
.hlmark{font-size:15px}
.hlgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(310px,1fr));gap:22px;margin-top:14px}
.hlcard{border:1px solid #e3e8ef;border-radius:10px;overflow:hidden;background:#fff;display:flex;flex-direction:column;transition:box-shadow .2s}
.hlcard:hover{box-shadow:0 4px 16px rgba(0,0,0,.12)}
.hlimg{height:185px;background:linear-gradient(135deg,#e8f0fb,#f7f9fc);display:flex;align-items:center;justify-content:center}
.hlimg img{width:100%;height:100%;object-fit:cover;display:block}
.hlnoimg{font-size:36px;color:#b9cbe4}
.hlbody{padding:14px 16px 16px;display:flex;flex-direction:column;gap:7px;flex:1}
.hlvenue{font-size:12px;color:#889;font-family:Verdana,sans-serif}
.hlt{font-weight:bold;color:#1a5dad;font-size:14.5px;line-height:1.45}
.hlsum{font-size:13px;color:#445;line-height:1.65;margin:0}
.hllinks{margin-top:auto;display:flex;gap:12px;flex-wrap:wrap;padding-top:6px}
.hllinks a{font-size:12.5px;font-family:Verdana,sans-serif}
.labline{font-size:16px;color:#1a5dad;font-weight:bold;margin:2px 0 8px}
footer{text-align:center;color:#999;font-size:13px;padding:26px 0;font-family:Verdana,sans-serif}
.avatar-fallback{width:210px;height:245px;border-radius:8px;background:#dde5ef;display:flex;align-items:center;justify-content:center;color:#8aa;font-size:15px}
.memberpage .avatar{width:150px;height:180px;object-fit:cover;border-radius:8px}
.memberpage .head{display:flex;gap:26px;align-items:center;margin-bottom:24px;flex-wrap:wrap}
.memberpage .content p{margin:9px 0}
.backlink{font-size:14px;font-family:Verdana,sans-serif;display:inline-block;margin-bottom:18px}
"""


def nav_html(active, prefix='', brand='You Zhou'):
    def a(href, label, key):
        cls = ' class="active"' if key == active else ''
        return '<a href="%s%s"%s>%s</a>' % (prefix, href, cls, label)
    return ('<nav class="nav"><div class="nav-inner">'
            '<a class="brand" href="%sindex.html">%s</a>'
            '<div class="links nav-links">%s%s%s%s%s</div></div></nav>') % (
        prefix, html_mod.escape(brand),
        a('index.html', 'Home', 'home'),
        a('members.html', 'Group Members', 'members'),
        a('publications.html', 'Publications', 'pubs'),
        a('honors.html', 'Honors & Awards', 'honors'),
        a('codes.html', 'Codes', 'codes'))


def footer_html(brand='You Zhou'):
    return '<footer>© %s %s · Generated by Homepage Editor</footer>' % (time.strftime('%Y'), html_mod.escape(brand))


def img_tag(data, rel, prefix='', cls='', size=''):
    if not rel:
        return '<div class="avatar-fallback%s">No Photo</div>' % (' ' + cls if cls else '')
    base = safe_basename(rel)
    s = ' width="%s"' % size if size else ''
    c = ' class="%s"' % cls if cls else ''
    return '<img src="%sphotos/%s"%s%s alt="">' % (prefix, base, s, c)


def build_publication_html(p, for_editor=False):
    parts = []
    if p.get('featured'):
        parts.append('<span class="hlmark" title="Research Highlight">⭐</span> ')
    link = p.get('link', '')
    title = html_mod.escape(p.get('title', ''))
    if link:
        parts.append('<a href="%s" target="_blank"><span class="t">%s.</span></a>' % (html_mod.escape(link, True), title))
    else:
        parts.append('<span class="t">%s.</span>' % title)
    if p.get('venue'):
        parts.append(' <span class="v">%s</span>,' % html_mod.escape(p['venue']))
    if p.get('meta'):
        parts.append(' %s.' % html_mod.escape(p['meta']))
    if p.get('note'):
        parts.append(' <span class="note">(%s)</span>' % html_mod.escape(p['note']))
    return '<div class="pub">%s</div>' % ''.join(parts)


def rewrite_homepage_imgs(html_str, img_map, prefix=''):
    """把成员主页 html 里引用的 uploads/... 图片替换为 photos/<basename>"""
    def repl(m):
        src = m.group(1)
        if src in img_map:
            return 'src="%sphotos/%s"' % (prefix, img_map[src])
        if src.startswith('photos/'):
            return 'src="%s%s"' % (prefix, src)
        return m.group(0)
    return re.sub(r'src="([^"]+)"', repl, html_str)


def build_site(data):
    prof = data.get('profile', {})
    pubs = data.get('publications', [])
    honors = data.get('honors', [])
    codes = data.get('codes', [])
    members = data.get('members', {})

    os.makedirs(os.path.join(SITE, 'photos'), exist_ok=True)
    memdir = os.path.join(SITE, 'members')
    os.makedirs(memdir, exist_ok=True)
    with open(os.path.join(SITE, '.nojekyll'), 'w') as f:
        f.write('')

    # 0. 成员 slug 映射（先计算，便于清理过期个人主页）
    used = set()
    slug_map = {}
    for group in ('current', 'alumni'):
        for m in members.get(group, []):
            slug_map[m['id']] = member_slug(m.get('name', ''), used)
    keep = {s + '.html' for s in slug_map.values()}
    for fn in os.listdir(memdir):
        if fn.endswith('.html') and fn not in keep:
            try:
                os.remove(os.path.join(memdir, fn))
            except OSError:
                pass

    # 1. 复制图片
    img_map = collect_site_images(data)
    for rel, base in img_map.items():
        src = resolve_image(rel)
        if src:
            shutil.copyfile(src, os.path.join(SITE, 'photos', base))

    # 2. 共用块（课题组名称 Pho-FreeU Lab 等）
    lab = (prof.get('labName') or '').strip()
    tname = html_mod.escape(prof.get('name', '') + ((' · ' + lab) if lab else ''))
    brand = lab or prof.get('name', 'You Zhou')
    foot = footer_html(prof.get('name', 'You Zhou') + ((' · ' + lab) if lab else ''))

    # ---- index.html 个人信息 ----
    offices = (prof.get('offices', '') or '').replace('&', '&amp;').replace('\n', '<br>')
    edu_items = []
    for e in prof.get('education', []):
        seg = '<li><strong>%s.</strong> %s, %s' % (html_mod.escape(e.get('degree', '')),
                                                   html_mod.escape(e.get('period', '')),
                                                   html_mod.escape(e.get('place', '')))
        if e.get('sup'):
            if e.get('supUrl'):
                seg += '. Supervisor: <a href="%s" target="_blank">%s</a>' % (html_mod.escape(e['supUrl'], True), html_mod.escape(e['sup']))
            else:
                seg += '. Supervisor: %s' % html_mod.escape(e['sup'])
        edu_items.append(seg + '.</li>')
    school = ''
    if prof.get('schoolLink'):
        school = ('<div class="meta"><strong>%s: </strong>'
                  '<a href="%s" target="_blank">%s</a></div>') % (
            html_mod.escape(prof.get('schoolLinkLabel', 'Homepage')), html_mod.escape(prof['schoolLink'], True), html_mod.escape(prof['schoolLink']))
    index_html = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{tname}</title><link rel="stylesheet" href="style.css"></head><body>
{nav_html('home', brand=brand)}
<div class="wrap">
<div class="card hero">
  {img_tag(data, prof.get('avatar',''))}
  <div class="info">
    <h1>{html_mod.escape(prof.get('name',''))}</h1>
    <div class="role">{html_mod.escape(prof.get('title',''))}</div>
    {('<div class="labline">' + html_mod.escape(lab) + '</div>') if lab else ''}
    <div class="role" style="font-size:17px">{html_mod.escape(prof.get('affiliation',''))}</div>
    <div class="meta">{html_mod.escape(prof.get('address',''))}</div>
    <div class="meta"><strong>Email: </strong>{html_mod.escape(prof.get('email',''))}</div>
    {school}
    <div class="meta" style="margin-top:8px">{offices}</div>
  </div>
</div>
<div class="card"><h2 class="sec">About Me</h2>{prof.get('aboutHtml','')}</div>
<div class="card"><h2 class="sec">Education</h2><ul class="edu">{''.join(edu_items)}</ul></div>
<div class="card"><h2 class="sec">Services</h2>{prof.get('servicesHtml','')}</div>
{foot}
</div></body></html>"""
    with open(os.path.join(SITE, 'index.html'), 'w', encoding='utf-8') as f:
        f.write(index_html)

    # ---- members.html 团队成员 ----
    def member_card(m, group):
        base_info = ''
        if group == 'current':
            base_info = '%s, %s' % (m.get('degree', ''), m.get('year', ''))
        else:
            base_info = m.get('degree', '')
        period = ''
        if group == 'alumni':
            period = '<div class="pd">%s</div><div class="pd">%s</div>' % (
                html_mod.escape(m.get('period', '')), html_mod.escape(m.get('now', '')))
        name = html_mod.escape(m.get('name', ''))
        if m.get('link'):
            name_html = '<a href="%s" target="_blank">%s</a>' % (html_mod.escape(m['link'], True), name)
        elif m.get('homepageEnabled') and (m.get('homepageHtml') or '').strip():
            name_html = '<a href="members/%s.html">%s</a>' % (slug_map[m['id']], name)
        else:
            name_html = name
        return ('<div class="mcard">%s<div class="nm">%s</div>'
                '<div class="dg">%s</div>%s</div>') % (
            img_tag(data, m.get('photo', '')), name_html,
            html_mod.escape(base_info), period)

    cur_cards = ''.join(member_card(m, 'current') for m in members.get('current', []))
    alm_cards = ''.join(member_card(m, 'alumni') for m in members.get('alumni', []))
    members_html = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Group Members - {tname}</title><link rel="stylesheet" href="style.css"></head><body>
{nav_html('members', brand=brand)}
<div class="wrap">
<div class="card"><h1 class="page-title">Group Members</h1>
<p style="color:#667;font-family:Verdana,sans-serif;font-size:14px">Including co-supervised students. Click a name to visit the personal homepage (if available).</p>
<h2 class="sec" style="margin-top:20px">Current Members</h2>
<div class="mgrid">{cur_cards}</div></div>
<div class="card"><h2 class="sec">Alumni (including Co-supervised Students)</h2>
<div class="mgrid">{alm_cards}</div></div>
{foot}
</div></body></html>"""
    with open(os.path.join(SITE, 'members.html'), 'w', encoding='utf-8') as f:
        f.write(members_html)

    # ---- publications.html ----
    journals = [p for p in pubs if p.get('type') == 'journal']
    confs = [p for p in pubs if p.get('type') == 'conf']
    patents = [p for p in pubs if p.get('type') == 'patent']

    years = sorted({p.get('year', 0) for p in journals}, reverse=True)
    jsec = []
    recent = [y for y in years if y and y >= 2020]
    for y in recent:
        items = ''.join(build_publication_html(p) for p in journals if p.get('year') == y)
        jsec.append('<h3 class="year">%d</h3>%s' % (y, items))
    old = [p for p in journals if not (p.get('year') and p.get('year') >= 2020)]
    if old:
        items = ''.join(build_publication_html(p) for p in old)
        jsec.append('<h3 class="year">Before 2020</h3>%s' % items)

    csec = []
    for track in ('Artificial Intelligence', 'Optics and Photonics'):
        items = ''.join(build_publication_html(p) for p in confs if p.get('track') == track)
        if items:
            csec.append('<h3 class="year">%s</h3>%s' % (track, items))
    others = [p for p in confs if p.get('track') not in ('Artificial Intelligence', 'Optics and Photonics')]
    if others:
        csec.append('<h3 class="year">Others</h3>%s' % ''.join(build_publication_html(p) for p in others))

    psec = ''.join('<li style="margin:8px 0">%s. <em>%s</em>. %s, %s.</li>' % (
        html_mod.escape(p.get('authors', '')), html_mod.escape(p.get('title', '')),
        html_mod.escape(p.get('venue', '')), html_mod.escape(p.get('meta', ''))) for p in patents)

    # Research Highlights（亮点展示卡片）
    featured = [p for p in pubs if p.get('featured')]
    hl_cards = []
    for p in featured:
        rel = p.get('image', '')
        base = img_map.get(rel) if rel else None
        if rel and base and os.path.exists(os.path.join(SITE, 'photos', base)):
            img_html = '<img src="photos/%s" alt="">' % base
        else:
            img_html = '<div class="hlnoimg">⭐</div>'
        links = []
        if p.get('link'):
            links.append('<a href="%s" target="_blank">📄 Paper</a>' % html_mod.escape(p['link'], True))
        for l in (p.get('extraLinks') or []):
            if l.get('url'):
                links.append('<a href="%s" target="_blank">%s</a>' % (
                    html_mod.escape(l['url'], True), html_mod.escape(l.get('label') or 'Link')))
        vline = ' · '.join(x for x in (p.get('venue', ''), str(p.get('year') or '')) if x)
        hl_cards.append(
            ('<div class="hlcard"><div class="hlimg">%s</div><div class="hlbody">'
             '<div class="hlvenue">%s</div>'
             '<a class="hlt" href="%s" target="_blank">%s</a>%s'
             '<div class="hllinks">%s</div></div></div>') % (
            img_html,
            html_mod.escape(vline),
            html_mod.escape(p.get('link') or '#', True) or '#',
            html_mod.escape(p.get('title', '')),
            ('<p class="hlsum">%s</p>' % html_mod.escape(p['summary'])) if p.get('summary') else '',
            ''.join(links)))
    highlights_html = ''
    if featured:
        highlights_html = ('<h2 class="sec" style="margin-top:20px">⭐ Research Highlights</h2>'
                           '<div class="hlgrid">%s</div>' % ''.join(hl_cards))

    pubs_html = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Publications - {tname}</title><link rel="stylesheet" href="style.css"></head><body>
{nav_html('pubs', brand=brand)}
<div class="wrap">
<div class="card">
<h1 class="page-title">Publications</h1>
<p style="color:#667;font-family:Verdana,sans-serif;font-size:14px;margin-bottom:14px"><strong>({html_mod.escape(prof.get('pubNote',''))})</strong></p>
{highlights_html}
<h2 class="sec">1. Journal Papers</h2>{''.join(jsec)}
<h2 class="sec" style="margin-top:26px">2. Conference Papers</h2>{''.join(csec)}
<h2 class="sec" style="margin-top:26px">3. Granted Patents</h2><ul>{psec}</ul>
</div>
{foot}
</div></body></html>"""
    with open(os.path.join(SITE, 'publications.html'), 'w', encoding='utf-8') as f:
        f.write(pubs_html)

    # ---- honors.html ----
    hitems = ''.join('<li>%s</li>' % html_mod.escape(h.get('text', '')) for h in honors)
    honors_html = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Honors and Awards - {tname}</title><link rel="stylesheet" href="style.css"></head><body>
{nav_html('honors', brand=brand)}
<div class="wrap"><div class="card">
<h1 class="page-title">Honors and Awards</h1>
<ul class="honors" style="margin-top:16px">{hitems}</ul>
</div>{foot}</div></body></html>"""
    with open(os.path.join(SITE, 'honors.html'), 'w', encoding='utf-8') as f:
        f.write(honors_html)

    # ---- codes.html ----
    citems = ''.join('<li><a href="%s" target="_blank" style="font-weight:bold">%s</a>%s</li>' % (
        html_mod.escape(c.get('url', ''), True), html_mod.escape(c.get('name', '')),
        (' <span style="color:#888">— %s</span>' % html_mod.escape(c['desc'])) if c.get('desc') else '')
        for c in codes)
    codes_html = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Codes - {tname}</title><link rel="stylesheet" href="style.css"></head><body>
{nav_html('codes', brand=brand)}
<div class="wrap"><div class="card">
<h1 class="page-title">Publication Related Codes</h1>
<ul class="codes" style="margin-top:16px">{citems}</ul>
</div>{foot}</div></body></html>"""
    with open(os.path.join(SITE, 'codes.html'), 'w', encoding='utf-8') as f:
        f.write(codes_html)

    # ---- 成员个人主页 ----
    for group in ('current', 'alumni'):
        for m in members.get(group, []):
            if not (m.get('homepageEnabled') and (m.get('homepageHtml') or '').strip()):
                continue
            content = rewrite_homepage_imgs(m['homepageHtml'], img_map, prefix='../')
            dg = m.get('degree', '') + ((', ' + m['year']) if m.get('year') else '')
            mp = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{html_mod.escape(m.get('name',''))} - {tname}</title><link rel="stylesheet" href="../style.css"></head><body>
{nav_html('members', prefix='../', brand=brand)}
<div class="wrap memberpage">
<a class="backlink" href="../members.html">← Back to Group Members</a>
<div class="card">
<div class="head">
{img_tag(data, m.get('photo',''), cls='avatar')}
<div><h1 class="page-title" style="margin-bottom:4px">{html_mod.escape(m.get('name',''))}</h1>
<div style="color:#555;font-family:Verdana,sans-serif">{html_mod.escape(dg)}</div></div>
</div>
<div class="content">{content}</div>
</div>
{foot}
</div></body></html>"""
            with open(os.path.join(SITE, 'members', slug_map[m['id']] + '.html'), 'w', encoding='utf-8') as f:
                f.write(mp)

    with open(os.path.join(SITE, 'style.css'), 'w', encoding='utf-8') as f:
        f.write(CSS)

    files = []
    for dp, _, fns in os.walk(SITE):
        for fn in fns:
            files.append(os.path.relpath(os.path.join(dp, fn), SITE))
    return sorted(files)


# ---------------------------------------------------------------
# GitHub 授权与一键发布（GitHub Pages）
# ---------------------------------------------------------------
def load_gh_cfg():
    if os.path.exists(GH_CFG):
        try:
            with open(GH_CFG, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_gh_cfg(cfg):
    tmp = GH_CFG + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    os.replace(tmp, GH_CFG)


def gh_api(method, path, token, payload=None):
    """调用 GitHub REST API，返回 (status_code, body_str_or_dict)；网络异常时自动改走探测到的代理重试"""
    import urllib.error
    import urllib.request
    url = 'https://api.github.com' + path
    data = json.dumps(payload).encode('utf-8') if payload is not None else None
    headers = {'Authorization': 'Bearer ' + token, 'Accept': 'application/vnd.github+json',
               'User-Agent': 'homepage-editor'}

    def _try(opener):
        req = urllib.request.Request(url, data=data, method=method)
        for k, v in headers.items():
            req.add_header(k, v)
        with opener.open(req, timeout=30) as resp:
            body = resp.read().decode('utf-8', 'ignore')
            return resp.status, (json.loads(body) if body.strip().startswith(('{', '[')) else body)

    try:
        return _try(urllib.request.build_opener())
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode('utf-8', 'ignore')
    except Exception as e:
        proxy = _detect_proxy()
        if not proxy:
            return 0, 'network error: %s' % e
        try:
            return _try(urllib.request.build_opener(
                urllib.request.ProxyHandler({'https': proxy, 'http': proxy})))
        except Exception as e2:
            return 0, 'network error: %s / via proxy %s: %s' % (e, proxy, e2)


def find_git():
    """按优先级查找可用的 git.exe（覆盖未加入 Windows PATH 的便携版与常见安装位置）"""
    p = shutil.which('git')
    if p:
        return p
    import glob
    home = os.path.expanduser('~')
    candidates = [
        r'C:\Program Files\Git\cmd\git.exe',
        r'C:\Program Files (x86)\Git\cmd\git.exe',
        os.path.join(home, r'AppData\Local\Programs\Git\cmd\git.exe'),
    ]
    candidates += sorted(glob.glob(os.path.join(
        home, 'AppData', 'Local', 'GitHubDesktop', 'app-*', 'resources', 'app', 'git', 'cmd', 'git.exe')), reverse=True)
    # WorkBuddy PortableGit（本机实际存在的位置）
    candidates += sorted(glob.glob(os.path.join(
        home, '.workbuddy', 'binaries', 'PortableGit', 'versions', '*', 'mingw64', 'bin', 'git.exe')), reverse=True)
    candidates += sorted(glob.glob(os.path.join(
        home, '.workbuddy', 'binaries', 'PortableGit', 'versions', '*', 'bin', 'git.exe')), reverse=True)
    for c in candidates:
        if c and os.path.isfile(c):
            return c
    return None


def _detect_proxy():
    """探测可用代理：环境变量/系统代理 → 扫描常见本地代理端口并实际验证连通性"""
    import urllib.request
    import socket
    cands = []
    try:
        px = urllib.request.getproxies()
        for k in ('https', 'http'):
            v = px.get(k)
            if v and v not in cands:
                cands.append(v)
    except Exception:
        pass
    for port in (7897, 7890, 7891, 10809, 10808, 8118, 1080):
        try:
            s = socket.create_connection(('127.0.0.1', port), timeout=0.35)
            s.close()
            c = 'http://127.0.0.1:%d' % port
            if c not in cands:
                cands.append(c)
        except OSError:
            pass
    for c in cands:
        try:
            opener = urllib.request.build_opener(urllib.request.ProxyHandler({'https': c, 'http': c}))
            req = urllib.request.Request('https://api.github.com/zen', headers={'User-Agent': 'homepage-editor'})
            opener.open(req, timeout=6).read(64)
            return c
        except Exception:
            continue
    return ''


def github_publish(user, token, repo, branch, force, proxy=''):
    """构建网站 -> git 提交 -> 创建仓库(如需) -> 推送 -> 开启 Pages，返回 (日志列表, 线上地址)"""
    import subprocess
    log = []
    git = find_git()
    if not git:
        raise RuntimeError('未检测到 git。请安装 Git for Windows（https://git-scm.com/download/win，'
                           '安装时一路默认即可），或重启电脑后重试')

    env = dict(os.environ, GIT_TERMINAL_PROMPT='0', GIT_ASKPASS='echo', GCM_INTERACTIVE='Never')
    # git 不会读取 Windows 系统代理，需要显式注入（-c 配置 + 环境变量双保险）
    cfg = []
    if proxy:
        cfg = ['-c', 'http.proxy=' + proxy, '-c', 'https.proxy=' + proxy]
        env['HTTP_PROXY'] = env['http_proxy'] = proxy
        env['HTTPS_PROXY'] = env['https_proxy'] = proxy
        log.append('✓ 使用代理访问 GitHub: %s' % proxy)

    def run(*args, allow_fail=False):
        r = subprocess.run([git] + cfg + list(args), cwd=ROOT, capture_output=True,
                           text=True, encoding='utf-8', errors='replace', env=env)
        out = (r.stdout or '').strip()
        err = (r.stderr or '').strip()
        line = '$ git ' + ' '.join(args)
        if out:
            line += '\n' + out
        if r.returncode != 0 and err:
            line += '\n[err] ' + err
        log.append(line)
        if r.returncode != 0 and not allow_fail:
            raise RuntimeError('git 命令失败: ' + (err or out))
        return r

    # 1. 构建最新网站
    build_site(load_data())
    log.append('✓ 网站已构建到 %s' % ROOT)

    # 2. 写 .gitignore（排除编辑器密钥与备份）
    with open(os.path.join(ROOT, '.gitignore'), 'w', encoding='utf-8') as f:
        f.write(GITIGNORE)

    # 3. 初始化 / 切换分支
    if not os.path.isdir(os.path.join(ROOT, '.git')):
        run('init')
    r = run('symbolic-ref', 'HEAD', 'refs/heads/' + branch, allow_fail=True)
    if r.returncode != 0:
        run('checkout', '-B', branch)
    if not run('config', 'user.name', allow_fail=True).stdout.strip():
        run('config', 'user.name', user)
    if not run('config', 'user.email', allow_fail=True).stdout.strip():
        run('config', 'user.email', '%s@users.noreply.github.com' % user)

    # 4. 提交
    run('add', '-A')
    st = run('status', '--porcelain')
    if st.stdout.strip():
        run('commit', '-m', 'Publish homepage at %s' % time.strftime('%Y-%m-%d %H:%M:%S'))
    else:
        log.append('✓ 没有新的改动，无需提交')

    # 5. 检查 / 创建远程仓库
    code, body = gh_api('GET', '/repos/%s/%s' % (user, repo), token)
    if code == 200:
        log.append('✓ 仓库已存在: %s/%s' % (user, repo))
    elif code == 404:
        code2, body2 = gh_api('POST', '/user/repos', token, {
            'name': repo, 'description': 'Personal academic homepage',
            'has_issues': False, 'has_wiki': False, 'has_projects': False,
            'auto_init': False, 'private': False})
        if code2 == 201:
            log.append('✓ 已创建 GitHub 仓库: %s/%s' % (user, repo))
        elif code2 == 422:
            log.append('✓ 仓库已存在（跳过创建）')
        else:
            raise RuntimeError('创建仓库失败 (HTTP %s): %s' % (code2, body2))
    elif code == 401:
        raise RuntimeError('GitHub 授权失败（401）：令牌无效或已过期，请点击「更新授权」重新生成令牌')
    else:
        if code == 0:
            raise RuntimeError('无法连接 GitHub API: %s' % body)
        raise RuntimeError('访问 GitHub 失败 (HTTP %s)，请检查网络与用户名' % code)

    # 6. 推送（令牌只用于本次命令，不写入仓库配置）
    push_url = 'https://%s:%s@github.com/%s/%s.git' % (user, token, user, repo)
    args = ['push', push_url, '%s:%s' % (branch, branch)]
    if force:
        args.append('--force')

    def clean(s):
        return (s or '').replace(token, '***')

    # 网络兼容性重试：代理高延迟或部分代理截断 TLS 时，依次降级尝试
    ca = os.path.join(os.path.dirname(os.path.dirname(git)), 'usr', 'ssl', 'certs', 'ca-bundle.crt')
    variants = [([], '默认'), (['-c', 'http.version=HTTP/1.1'], 'HTTP/1.1')]
    if os.path.isfile(ca):
        variants.append((['-c', 'http.version=HTTP/1.1', '-c', 'http.sslBackend=openssl',
                          '-c', 'http.sslCAInfo=' + ca], 'HTTP/1.1+OpenSSL'))
    NET_KW = ('could not connect', 'failed to connect', 'timed out', 'schannel', 'close_notify',
              'ssl', 'tls', 'reset', 'refused', 'proxy', 'curl', 'connect', 'unable to access')
    r = None
    for extra, label in variants:
        r = subprocess.run([git] + cfg + extra + args, cwd=ROOT, capture_output=True,
                           text=True, encoding='utf-8', errors='replace', env=env)
        out = clean(r.stdout) + (('\n' + clean(r.stderr)) if (r.stderr or '').strip() else '')
        log.append('$ git push [%s]%s' % (label, ('\n' + out.strip()) if out.strip() else ''))
        if r.returncode == 0:
            break
        blob = ((r.stderr or '') + (r.stdout or '')).lower()
        if not any(k in blob for k in NET_KW):
            break  # 非网络类错误（如推送被拒绝），无需重试
    msg = clean(r.stderr or r.stdout)
    if r.returncode != 0:
        blob = msg.lower()
        if 'rejected' in blob or 'non-fast-forward' in blob:
            # 远程历史不同（如此前在 GitHub 网页上传的旧版主页）：自动改走 API 直传，
            # 在远程当前历史上追加提交并同步全部内容，不破坏远程已有历史。
            log.append('! 远程仓库已有不同的提交历史，自动改用 GitHub API 直传'
                       '（在远程历史上追加提交，不破坏旧历史）…')
            try:
                _api_publish(user, token, repo, branch, log)
                log.append('✓ 网站已发布（API 直传方式）')
            except Exception as e:
                raise RuntimeError(
                    'git 推送与 API 直传均失败。\ngit 错误: %s\nAPI 错误: %s\n'
                    '提示：也可勾选「强制覆盖远程」用本地历史完全替换远程。' % (msg.splitlines()[0] if msg else '未知', e))
        elif any(k in blob for k in NET_KW):
            # 网络不通（如代理对 github.com 分流异常）：改用 GitHub API 直传兜底
            log.append('! git 推送被网络阻断，自动改用 GitHub API 直传（走 api.github.com）…')
            try:
                _api_publish(user, token, repo, branch, log)
                log.append('✓ 网站已发布（API 直传方式，内容与 git 推送完全一致）')
            except Exception as e:
                raise RuntimeError(
                    'git 推送与 API 直传均失败。\ngit 错误: %s\nAPI 错误: %s\n'
                    '提示：可在发布设置「网络代理」手动填写可用代理后重试。' % (msg.splitlines()[0] if msg else '未知', e))
        else:
            raise RuntimeError('推送失败: ' + msg)
    else:
        log.append('✓ 已推送到 GitHub')

    # 7. 开启 GitHub Pages
    if repo.lower() == (user + '.github.io').lower():
        url = 'https://%s.github.io/' % user
        log.append('✓ 用户主仓库（%s），Pages 将自动生效' % repo)
    else:
        code3, body3 = gh_api('POST', '/repos/%s/%s/pages' % (user, repo), token,
                              {'source': {'branch': branch, 'path': '/'}})
        url = 'https://%s.github.io/%s/' % (user, repo)
        if code3 in (201, 204):
            log.append('✓ 已开启 GitHub Pages')
        elif code3 == 409:
            log.append('✓ GitHub Pages 已处于开启状态')
        else:
            log.append('! 开启 Pages 返回 %s，可稍后在仓库 Settings → Pages 手动开启' % code3)
    log.append('🎉 发布完成！线上地址: %s （首次部署约 1-2 分钟生效）' % url)
    return log, url


def _api_collect_files():
    """按 .gitignore 等价规则收集 ROOT 下待发布文件 -> {相对路径(正斜杠): 本地绝对路径}"""
    ignore_dirs = {'.git', '__pycache__', '备份'}
    ignore_files = {'Thumbs.db', 'desktop.ini'}
    ignore_rel = {'site-editor/github_config.json', 'site-editor/server_log.txt',
                  'site-editor/data.json.bak', 'site-editor/data.json.tmp', 'server_log.txt'}
    files = {}
    for dirpath, dirnames, filenames in os.walk(ROOT):
        dirnames[:] = [d for d in dirnames if d not in ignore_dirs]
        for fn in filenames:
            p = os.path.join(dirpath, fn)
            rel = os.path.relpath(p, ROOT).replace('\\', '/')
            if fn in ignore_files or rel in ignore_rel or rel.endswith('.bak'):
                continue
            files[rel] = p
    return files


def _api_publish(user, token, repo, branch, log):
    """git 推送不可用时的兜底发布：完全通过 api.github.com（Git Data API）创建提交并更新分支。
    基于远程分支当前头构建提交（天然快进），并删除远程已不存在于本地的文件，保持完全同步。"""
    import base64
    files = _api_collect_files()
    log.append('↻ API 直传：共 %d 个文件' % len(files))

    code, ref = gh_api('GET', '/repos/%s/%s/git/ref/heads/%s' % (user, repo, branch), token)
    head_sha = tree_sha = None
    if code == 200 and isinstance(ref, dict) and ref.get('object'):
        head_sha = ref['object']['sha']
        code, cm = gh_api('GET', '/repos/%s/%s/git/commits/%s' % (user, repo, head_sha), token)
        if code == 200 and isinstance(cm, dict):
            tree_sha = cm.get('tree', {}).get('sha')

    remote_paths = set()
    if tree_sha:
        code, tr = gh_api('GET', '/repos/%s/%s/git/trees/%s?recursive=1' % (user, repo, tree_sha), token)
        if code == 200 and isinstance(tr, dict):
            remote_paths = {e['path'] for e in tr.get('tree', []) if e.get('type') == 'blob'}

    entries = []
    for i, rel in enumerate(sorted(files)):
        with open(files[rel], 'rb') as f:
            data = f.read()
        code, b = gh_api('POST', '/repos/%s/%s/git/blobs' % (user, repo), token,
                         {'content': base64.b64encode(data).decode('ascii'), 'encoding': 'base64'})
        if code not in (200, 201):
            raise RuntimeError('上传 %s 失败 (HTTP %s): %s' % (rel, code, str(b)[:200]))
        entries.append({'path': rel, 'mode': '100644', 'type': 'blob', 'sha': b['sha']})
        if (i + 1) % 20 == 0:
            log.append('  …已上传 %d/%d' % (i + 1, len(files)))
    for rel in sorted(remote_paths - set(files)):
        if rel.lower() == 'readme.md':
            continue  # 保留用户在 GitHub 上维护的 README，其余内容完全同步
        entries.append({'path': rel, 'sha': None})  # 删除远程多余文件

    code, t = gh_api('POST', '/repos/%s/%s/git/trees' % (user, repo), token,
                     {'base_tree': tree_sha, 'tree': entries})
    if code not in (200, 201):
        raise RuntimeError('创建 Git 树失败 (HTTP %s): %s' % (code, str(t)[:200]))
    code, cm = gh_api('POST', '/repos/%s/%s/git/commits' % (user, repo), token, {
        'message': 'Publish homepage via API at %s' % time.strftime('%Y-%m-%d %H:%M:%S'),
        'tree': t['sha'], 'parents': [head_sha] if head_sha else []})
    if code not in (200, 201):
        raise RuntimeError('创建提交失败 (HTTP %s): %s' % (code, str(cm)[:200]))
    if head_sha:
        code, rr = gh_api('PATCH', '/repos/%s/%s/git/refs/heads/%s' % (user, repo, branch), token,
                          {'sha': cm['sha']})
    else:
        code, rr = gh_api('POST', '/repos/%s/%s/git/refs' % (user, repo), token,
                          {'ref': 'refs/heads/%s' % branch, 'sha': cm['sha']})
    if code not in (200, 201):
        raise RuntimeError('更新分支 %s 失败 (HTTP %s): %s' % (branch, code, str(rr)[:200]))
    log.append('✓ API 直传完成：已创建提交并更新 %s 分支' % branch)


# ---------------------------------------------------------------
# HTTP 服务
# ---------------------------------------------------------------
class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _send(self, code, body, ctype='application/json; charset=utf-8'):
        if isinstance(body, str):
            body = body.encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        self.end_headers()
        self.wfile.write(body)

    def _send_file(self, path, ctype=None):
        if not os.path.isfile(path):
            self._send(404, json.dumps({'error': 'not found'}))
            return
        ext = os.path.splitext(path)[1].lower()
        ct = {'': None}.get(ext) or {
            '.html': 'text/html; charset=utf-8', '.css': 'text/css; charset=utf-8',
            '.js': 'application/javascript; charset=utf-8', '.json': 'application/json; charset=utf-8',
            '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.png': 'image/png',
            '.gif': 'image/gif', '.svg': 'image/svg+xml', '.webp': 'image/webp'}.get(ext, 'application/octet-stream')
        with open(path, 'rb') as f:
            self._send(200, f.read(), ct)

    def _read_body(self):
        length = int(self.headers.get('Content-Length', 0) or 0)
        return self.rfile.read(length) if length else b''

    def do_GET(self):
        path = urlparse(self.path).path
        if path == '/api/data':
            self._send(200, json.dumps(load_data(), ensure_ascii=False))
        elif path == '/api/export':
            try:
                files = build_site(load_data())
                self._send(200, json.dumps({'ok': True, 'dir': SITE, 'files': files}, ensure_ascii=False))
            except Exception as e:
                self._send(500, json.dumps({'ok': False, 'error': str(e)}, ensure_ascii=False))
        elif path.startswith('/site/'):
            rel = unquote(path[len('/site/'):]).replace('\\', '/')
            first = rel.split('/')[0] if rel else ''
            if '..' in rel or first in ('site-editor', '备份', '.git'):
                self._send(404, json.dumps({'error': 'not found'}))
                return
            self._send_file(os.path.join(SITE, rel))
        elif path == '/api/github/config':
            cfg = load_gh_cfg()
            self._send(200, json.dumps({
                'user': cfg.get('user', ''), 'repo': cfg.get('repo', ''),
                'branch': cfg.get('branch', 'main'), 'remember': cfg.get('remember', True),
                'proxy': cfg.get('proxy', ''),
                'hasToken': bool(cfg.get('token'))}, ensure_ascii=False))
        elif path in ('/', '/index.html', '/editor.html'):
            self._send_file(os.path.join(BASE, 'editor.html'))
        else:
            rel = unquote(path.lstrip('/'))
            self._send_file(os.path.join(BASE, rel))

    def do_POST(self):
        path = urlparse(self.path).path
        try:
            body = json.loads(self._read_body().decode('utf-8'))
        except Exception:
            self._send(400, json.dumps({'ok': False, 'error': 'bad json'}))
            return
        if path == '/api/data':
            save_data(body)
            self._send(200, json.dumps({'ok': True}))
        elif path == '/api/photo':
            raw = base64.b64decode(body.get('b64', ''))
            ext = os.path.splitext(body.get('name', 'photo.jpg'))[1].lower() or '.jpg'
            if ext not in ('.jpg', '.jpeg', '.png', '.gif', '.webp'):
                ext = '.jpg'
            name = 'photo_%d%s' % (int(time.time() * 1000), ext)
            os.makedirs(os.path.join(UPLOADS, 'photos'), exist_ok=True)
            with open(os.path.join(UPLOADS, 'photos', name), 'wb') as f:
                f.write(raw)
            self._send(200, json.dumps({'ok': True, 'path': 'uploads/photos/' + name}))
        elif path == '/api/word2html':
            raw = base64.b64decode(body.get('b64', ''))
            media_dir = os.path.join(UPLOADS, 'word_media')
            conv = DocxConverter(raw, media_dir, 'uploads/word_media')
            try:
                html_str = conv.convert()
                self._send(200, json.dumps({'ok': True, 'html': html_str}, ensure_ascii=False))
            except Exception as e:
                self._send(500, json.dumps({'ok': False, 'error': 'Word 解析失败: %s' % e}, ensure_ascii=False))
        elif path == '/api/github/config':
            saved = load_gh_cfg()
            new_cfg = {
                'user': (body.get('user') or '').strip(),
                'repo': (body.get('repo') or '').strip(),
                'branch': (body.get('branch') or 'main').strip(),
                'remember': bool(body.get('remember', True)),
            }
            # 请求未携带 proxy 字段时保留旧值，避免授权保存时误清除
            new_cfg['proxy'] = (body.get('proxy') or '').strip() if 'proxy' in body else saved.get('proxy', '')
            token = (body.get('token') or '').strip()
            if new_cfg['remember'] and token:
                new_cfg['token'] = token          # 仅在用户勾选记住时写入本地配置
            elif saved.get('token'):
                new_cfg['token'] = saved['token']  # 保留旧令牌
            save_gh_cfg(new_cfg)
            self._send(200, json.dumps({'ok': True}))
        elif path == '/api/github/publish':
            saved = load_gh_cfg()
            user = (body.get('user') or saved.get('user') or '').strip()
            repo = (body.get('repo') or saved.get('repo') or '').strip()
            branch = (body.get('branch') or saved.get('branch') or 'main').strip()
            token = (body.get('token') or '').strip() or saved.get('token', '')
            force = bool(body.get('force'))
            if not user:
                self._send(400, json.dumps({'ok': False, 'error': '尚未授权：请先填写 GitHub 用户名并完成授权'}, ensure_ascii=False))
                return
            if not token:
                self._send(403, json.dumps({'ok': False, 'error': '尚未授权：请先点击「账号授权」生成并粘贴令牌'}, ensure_ascii=False))
                return
            repo = repo or (user + '.github.io')
            proxy = (body.get('proxy') or '').strip()
            if proxy and proxy != (saved.get('proxy') or ''):
                saved['proxy'] = proxy  # 记住手动填写的代理，下次自动带上
                save_gh_cfg(saved)
            try:
                if not proxy:
                    proxy = _detect_proxy()
                log, url = github_publish(user, token, repo, branch, force, proxy)
                self._send(200, json.dumps({'ok': True, 'log': log, 'url': url}, ensure_ascii=False))
            except Exception as e:
                self._send(500, json.dumps({'ok': False, 'error': str(e)}, ensure_ascii=False))
        elif path == '/api/fetchog':
            # 从论文页面抓取配图（og:image / twitter:image）
            import urllib.error
            import urllib.request
            from urllib.parse import urljoin
            url = (body.get('url') or '').strip()
            if not url.lower().startswith(('http://', 'https://')):
                self._send(400, json.dumps({'ok': False, 'error': '请提供有效的论文链接 (http/https)'}, ensure_ascii=False))
                return
            UA = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36',
                  'Accept': 'text/html,application/xhtml+xml,image/*;q=0.9,*/*;q=0.8'}
            try:
                req = urllib.request.Request(url, headers=UA)
                with urllib.request.urlopen(req, timeout=12) as resp:
                    page = resp.read(3000000).decode('utf-8', 'ignore')
            except Exception as e:
                self._send(502, json.dumps({'ok': False, 'error': '无法访问论文页面（部分出版商会屏蔽自动访问）：%s' % e}, ensure_ascii=False))
                return
            m = (re.search(r'<meta[^>]+property=["\']og:image["\'][^>]*content=["\']([^"\']+)', page, re.I)
                 or re.search(r'<meta[^>]+content=["\']([^"\']+)["\'][^>]*property=["\']og:image["\']', page, re.I)
                 or re.search(r'<meta[^>]+name=["\']twitter:image["\'][^>]*content=["\']([^"\']+)', page, re.I)
                 or re.search(r'<link[^>]+rel=["\']image_src["\'][^>]*href=["\']([^"\']+)', page, re.I))
            if not m:
                self._send(404, json.dumps({'ok': False, 'error': '页面未提供配图元数据(og:image)，请手动上传图片'}, ensure_ascii=False))
                return
            img_url = html_mod.unescape(m.group(1)).strip()
            img_url = urljoin(url, img_url)
            try:
                req2 = urllib.request.Request(img_url, headers=UA)
                with urllib.request.urlopen(req2, timeout=20) as resp2:
                    ctype = resp2.headers.get('Content-Type', '')
                    data = resp2.read(10485760)
            except Exception as e:
                self._send(502, json.dumps({'ok': False, 'error': '图片下载失败：%s' % e}, ensure_ascii=False))
                return
            if not ctype.startswith('image/') and not re.search(r'\.(jpe?g|png|gif|webp)(\?|$)', img_url, re.I):
                self._send(415, json.dumps({'ok': False, 'error': '抓取到的不是图片文件，请手动上传'}, ensure_ascii=False))
                return
            ext = {'.jpg': 'image/jpeg', '.png': 'image/png', '.gif': 'image/gif', '.webp': 'image/webp'}
            ext = next((k for k, v in ext.items() if v == ctype.lower().split(';')[0]), None)
            if not ext:
                em = re.search(r'\.(jpe?g|png|gif|webp)(\?|$)', img_url, re.I)
                ext = '.' + em.group(1).lower() if em else '.jpg'
            os.makedirs(os.path.join(UPLOADS, 'fetched'), exist_ok=True)
            name = 'og_%d%s' % (int(time.time() * 1000), ext)
            with open(os.path.join(UPLOADS, 'fetched', name), 'wb') as f:
                f.write(data)
            self._send(200, json.dumps({'ok': True, 'path': 'uploads/fetched/' + name}, ensure_ascii=False))
        else:
            self._send(404, json.dumps({'ok': False, 'error': 'unknown api'}))


def main():
    if '--build' in sys.argv:
        data = load_data()
        if not data:
            print('data.json 不存在或为空')
            sys.exit(1)
        files = build_site(data)
        print('网站已生成: %s (%d 个文件)' % (SITE, len(files)))
        return
    os.makedirs(os.path.join(UPLOADS, 'photos'), exist_ok=True)
    if not os.path.exists(DATA_FILE):
        save_data({})
    # 防止重复启动：先探测端口是否已有编辑器在运行
    import socket
    probe = socket.socket()
    probe.settimeout(1)
    try:
        probe.connect(('127.0.0.1', PORT))
        probe.close()
        print('端口 %d 已被占用：编辑器可能已在运行，请直接访问 http://127.0.0.1:%d' % (PORT, PORT))
        return
    except OSError:
        pass
    finally:
        try:
            probe.close()
        except Exception:
            pass
    server = ThreadingHTTPServer(('127.0.0.1', PORT), Handler)
    print('编辑器已启动: http://127.0.0.1:%d  (Ctrl+C 退出)' % PORT)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\n已退出')


if __name__ == '__main__':
    main()
