# -*- coding: utf-8 -*-
"""页面导出: OneNote 页面 -> Markdown 文件 (复用 onenote_html2md 转换器)。"""

import os
import re

from . import hierarchy
from .onenote_html2md import RenderContext, convert

_TAG_RE = re.compile(r'<[^>]+>')
_SCRIPT_RE = re.compile(r'(?s)<script.*?</script>')
_STYLE_RE = re.compile(r'(?s)<style.*?</style>')
_COMMENT_RE = re.compile(r'(?s)<!--.*?-->')
_BR_RE = re.compile(r'(?i)<br\s*/?>')
_BLOCK_END_RE = re.compile(r'(?i)</(p|div|li|tr|h[1-6]|blockquote)>')


def strip_tags_fallback(html):
    """无第三方依赖的降级提取 (转换器异常时的兜底, 与 PS 版逻辑一致)。"""
    t = _SCRIPT_RE.sub(' ', html)
    t = _STYLE_RE.sub(' ', t)
    t = _COMMENT_RE.sub(' ', t)
    t = _BR_RE.sub('\n', t)
    t = _BLOCK_END_RE.sub('\n', t)
    t = _TAG_RE.sub(' ', t)
    for a, b in (('&nbsp;', ' '), ('&amp;', '&'), ('&lt;', '<'),
                 ('&gt;', '>'), ('&quot;', '"'), ('&#39;', "'")):
        t = t.replace(a, b)
    t = re.sub(r'[ \t]+', ' ', t)
    return '\n'.join(line.strip() for line in t.splitlines() if line.strip())


def html_to_markdown(html, assets_dir=None):
    """HTML -> Markdown; assets_dir 提供时解码 base64 图片为文件。失败时降级为纯文本。

    返回 (markdown, degraded_bool)。
    """
    try:
        ctx = RenderContext(assets_dir=assets_dir)
        md = convert(html, ctx)
        if md and md.strip():
            return md, False
    except Exception:
        pass
    return strip_tags_fallback(html), True


def safe_filename(name, max_len=80):
    s = (name or '').strip()
    for ch in '\\/:*?"<>|':
        s = s.replace(ch, '_')
    s = ''.join(c for c in s if ord(c) >= 32).strip().rstrip('. ')
    return s[:max_len].strip() or 'untitled'


def unique_path(directory, base, ext='.md'):
    """目录内去重: 同名追加 ' (2)'、' (3)'..."""
    path = os.path.join(directory, base + ext)
    n = 2
    while os.path.exists(path):
        path = os.path.join(directory, '%s (%d)%s' % (base, n, ext))
        n += 1
    return path


def write_page_markdown(meta, html, out_root, with_images=True, link=None,
                        degraded_note=None):
    """把单页写入 out_root 下按 笔记本/分区 布局的 .md 文件。

    返回相对 out_root 的路径。
    """
    sec_dir = os.path.join(out_root, *meta.get('section_path', '').split('\\') if meta.get('section_path') else [])
    os.makedirs(sec_dir, exist_ok=True)
    base = safe_filename(meta.get('title', ''))
    md_path = unique_path(sec_dir, base)
    assets_dir = os.path.join(sec_dir, base + '.assets') if with_images else None

    body, degraded = html_to_markdown(html, assets_dir)
    if degraded:
        # 保留原始 HTML 备查
        html_path = unique_path(sec_dir, base, '.html')
        with open(html_path, 'w', encoding='utf-8', newline='') as f:
            f.write(html)
        body += ('\r\n\r\n> 注: Markdown 转换失败, 本文件为降级提取结果; '
                 '完整内容见同目录 %s.html。' % base)
    if degraded_note:
        body += '\r\n\r\n> ' + degraded_note

    fm = hierarchy.build_frontmatter(meta, link)
    with open(md_path, 'w', encoding='utf-8', newline='') as f:
        f.write(fm + body)
    return os.path.relpath(md_path, out_root)


def write_index(out_root, results, failed):
    """导出索引 README.md。results: [(meta, rel_path)]"""
    lines = ['# OneNote 导出索引', '',
             '- 导出时间: ' + __import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
             '- 页面统计: %d 成功, %d 失败' % (len(results), failed),
             '- 来源: OneNote 桌面版 COM API (只读)', '']
    cur_nb = cur_sec = None
    for meta, rel in sorted(results, key=lambda r: (r[0].get('section_path', ''), r[0].get('title', ''))):
        if meta['notebook'] != cur_nb:
            cur_nb, cur_sec = meta['notebook'], None
            lines += ['', '## ' + cur_nb]
        if meta['section'] != cur_sec:
            cur_sec = meta['section']
            lines += ['', '### ' + cur_sec, '']
        lines.append('- [%s](%s)' % (meta['title'], rel.replace('\\', '/')))
    lines.append('')
    path = os.path.join(out_root, 'README.md')
    with open(path, 'w', encoding='utf-8', newline='') as f:
        f.write('\r\n'.join(lines))
    return path
