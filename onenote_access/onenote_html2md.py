#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
onenote_html2md.py - OneNote 页面 HTML 转 Markdown 转换器 (纯 Python 标准库, 无第三方依赖)

用法:
    python onenote_html2md.py input.html -o output.md [--assets-dir DIR]
    python onenote_html2md.py - -o output.md < input.html      # 从 stdin 读取
    python onenote_html2md.py input.html                        # 输出到 stdout

特性:
    - 容错解析 OneNote 导出的 messy HTML (未闭合标签 / 嵌套 / 大量 style 属性)
    - 标题/段落/列表(支持嵌套)/表格/引用/代码块/分割线/链接/加粗/斜体/行内代码
    - base64 data URI 图片自动解码为文件, Markdown 中以相对路径引用
    - Markdown 特殊字符转义, 避免列表/标题误判
    - 输入编码自动探测 (utf-8 / utf-8-sig / gbk)

本脚本可独立于 Extract-OneNote.ps1 使用, 也可嵌入其他流程。
"""

import argparse
import base64
import hashlib
import os
import re
import sys
from html.parser import HTMLParser

# ---------------------------------------------------------------- 常量

VOID_TAGS = {
    'area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input',
    'keygen', 'link', 'meta', 'param', 'source', 'track', 'wbr',
}

# 这些子树整体丢弃 (内容与样式噪音)
DROP_SUBTREE = {'script', 'style', 'title', 'head', 'noscript', 'template'}

# 触发 <p> 自动闭合的块级起始标签
CLOSES_P = {
    'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'div', 'ul', 'ol', 'dl',
    'table', 'blockquote', 'pre', 'hr', 'section', 'article', 'figure',
    'main', 'details', 'fieldset', 'form',
}

HEADING_TAGS = {'h1', 'h2', 'h3', 'h4', 'h5', 'h6'}

# 视为"块级容器"的标签: 其子节点按块渲染
CONTAINER_TAGS = {
    'div', 'section', 'article', 'main', 'figure', 'details', 'fieldset',
    'form', 'body', 'html', 'dl', 'center', 'caption', 'figcaption',
}

BLOCK_TAGS = CONTAINER_TAGS | HEADING_TAGS | {
    'p', 'ul', 'ol', 'li', 'table', 'thead', 'tbody', 'tfoot', 'tr', 'td',
    'th', 'blockquote', 'pre', 'hr',
}

MD_SPECIAL = re.compile(r'([\\`*\[\]])')
LINE_START = re.compile(r'^(\s*)(#{1,6}\s|>\s|[-+*]\s|\d+\.\s)')


# ---------------------------------------------------------------- DOM

class Element(object):
    __slots__ = ('tag', 'attrs', 'children')

    def __init__(self, tag, attrs=None):
        self.tag = tag
        self.attrs = attrs or {}
        self.children = []


class DomBuilder(HTMLParser):
    """容错 DOM 构建器: 处理 OneNote HTML 中常见的未闭合标签。"""

    def __init__(self):
        HTMLParser.__init__(self, convert_charrefs=True)
        self.root = Element('#root')
        self.stack = [self.root]
        self.drop_depth = 0

    # -- 自动闭合规则 -------------------------------------------------
    def _autoclose(self, tag):
        s = self.stack
        if tag == 'li':
            for i in range(len(s) - 1, 0, -1):
                if s[i].tag == 'li':
                    del s[i:]
                    return
                if s[i].tag in ('ul', 'ol', 'menu'):
                    return
        elif tag in ('td', 'th'):
            for i in range(len(s) - 1, 0, -1):
                if s[i].tag in ('td', 'th'):
                    del s[i:]
                    return
                if s[i].tag in ('tr', 'table'):
                    return
        elif tag == 'tr':
            for i in range(len(s) - 1, 0, -1):
                if s[i].tag == 'tr':
                    del s[i:]
                    return
                if s[i].tag in ('table', 'thead', 'tbody', 'tfoot'):
                    return
        elif tag in CLOSES_P:
            for i in range(len(s) - 1, 0, -1):
                if s[i].tag == 'p':
                    del s[i:]
                    return
                if s[i].tag in ('td', 'th', 'li', 'dd', 'dt', 'blockquote'):
                    return

    # -- HTMLParser 回调 ----------------------------------------------
    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        self._autoclose(tag)
        el = Element(tag, dict(attrs))
        if self.drop_depth == 0 and tag not in DROP_SUBTREE:
            self.stack[-1].children.append(el)
        if tag in DROP_SUBTREE:
            self.drop_depth += 1
        if tag not in VOID_TAGS:
            self.stack.append(el)

    def handle_startendtag(self, tag, attrs):
        tag = tag.lower()
        self._autoclose(tag)
        if self.drop_depth == 0 and tag not in DROP_SUBTREE:
            self.stack[-1].children.append(Element(tag, dict(attrs)))

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag in DROP_SUBTREE and self.drop_depth > 0:
            self.drop_depth -= 1
        for i in range(len(self.stack) - 1, 0, -1):
            if self.stack[i].tag == tag:
                del self.stack[i:]
                return

    def handle_data(self, data):
        if self.drop_depth > 0 or not data:
            return
        self.stack[-1].children.append(('#text', data))

    def handle_comment(self, data):
        pass

    def handle_decl(self, decl):
        pass

    def handle_pi(self, data):
        pass

    def unknown_decl(self, data):
        pass


# ---------------------------------------------------------------- 渲染上下文

class RenderContext(object):
    def __init__(self, assets_dir=None, out_path=None):
        self.assets_dir = assets_dir
        self.out_path = out_path
        self.images = 0


def collapse(s):
    return re.sub(r'\s+', ' ', s)


def esc_text(t):
    return MD_SPECIAL.sub(r'\\\1', t)


def fix_line_start(line):
    if LINE_START.match(line):
        return '\\' + line
    return line


def hard_break():
    return '  \n'


# ---------------------------------------------------------------- 行内渲染

def render_inline(node, ctx):
    """把行内节点渲染为字符串 (可包含换行)。"""
    if isinstance(node, tuple):
        return esc_text(collapse(node[1]))
    tag = node.tag
    if tag == 'br':
        return hard_break()
    if tag == 'img':
        return render_img(node, ctx)
    inner = ''.join(render_inline(c, ctx) for c in node.children)
    if not inner.strip():
        return ''
    if tag in ('b', 'strong'):
        return '**' + inner + '**'
    if tag in ('i', 'em'):
        return '*' + inner + '*'
    if tag in ('code', 'tt', 'kbd', 'samp'):
        return '`' + inner.strip() + '`'
    if tag == 'a':
        href = (node.attrs.get('href') or '').strip()
        text = inner.strip()
        if not href or href.startswith(('javascript:', 'onenote:', 'file:', '#')):
            return text
        if not text:
            text = href
        return '[' + text + '](' + href + ')'
    # u/s/strike/del/sub/sup/small/mark/span/font/... 以及未知行内标签: 取内容
    return inner


def render_img(node, ctx):
    src = (node.attrs.get('src') or '').strip()
    alt = collapse(node.attrs.get('alt') or '').strip()
    if not src:
        return ''
    if src.startswith('data:'):
        m = re.match(r'data:(image/[^;,\s]+);base64,(.+)', src, re.S)
        if m:
            mime = m.group(1)
            try:
                data = base64.b64decode(m.group(2), validate=False)
            except Exception:
                data = None
            if data and ctx.assets_dir:
                ext = mime.split('/', 1)[1].replace('jpeg', 'jpg').replace('svg+xml', 'svg')
                digest = hashlib.sha1(data).hexdigest()[:10]
                fname = digest + '.' + ext
                try:
                    if not os.path.isdir(ctx.assets_dir):
                        os.makedirs(ctx.assets_dir)
                    full = os.path.join(ctx.assets_dir, fname)
                    if not os.path.exists(full):
                        with open(full, 'wb') as f:
                            f.write(data)
                    if ctx.out_path:
                        rel = os.path.relpath(full, os.path.dirname(os.path.abspath(ctx.out_path)))
                    else:
                        rel = fname
                    rel = rel.replace(os.sep, '/')
                    ctx.images += 1
                    return '![' + alt + '](' + rel + ')'
                except Exception:
                    pass
            return '![' + (alt or 'image') + ']'
    return '![' + alt + '](' + src + ')'


# ---------------------------------------------------------------- 块级渲染

def iter_child_elements(node):
    for c in node.children:
        if not isinstance(c, tuple):
            yield c


def collect_rows(table_node):
    """收集本表所有 tr (不进入嵌套 table)。"""
    rows = []

    def walk(n):
        for c in iter_child_elements(n):
            if c.tag == 'table':
                continue
            if c.tag == 'tr':
                rows.append(c)
            else:
                walk(c)

    walk(table_node)
    return rows


def render_blocks(node, ctx):
    """容器 -> markdown 块列表 (每块可为多行字符串)。"""
    blocks = []
    inline = []

    def flush():
        text = ''.join(inline).strip()
        del inline[:]
        if not text:
            return
        lines = text.split('\n')
        lines = [fix_line_start(l) if l.strip() else l for l in lines]
        blocks.append('\n'.join(lines))

    for child in node.children:
        if isinstance(child, tuple):
            inline.append(render_inline(child, ctx))
            continue
        tag = child.tag
        if tag in CONTAINER_TAGS or tag == '#root':
            flush()
            blocks.extend(render_blocks(child, ctx))
        elif tag in HEADING_TAGS:
            flush()
            text = render_inline(child, ctx).strip()
            if text:
                level = int(tag[1])
                blocks.append('#' * level + ' ' + fix_line_start(text))
        elif tag == 'p':
            flush()
            text = render_inline(child, ctx).strip()
            if text:
                lines = text.split('\n')
                lines = [fix_line_start(l) if l.strip() else l for l in lines]
                blocks.append('\n'.join(lines))
        elif tag in ('ul', 'ol', 'menu', 'dl'):
            flush()
            lines = render_list(child, ctx, 0)
            if lines:
                blocks.append('\n'.join(lines))
        elif tag == 'table':
            flush()
            t = render_table(child, ctx)
            if t:
                blocks.append(t)
        elif tag == 'blockquote':
            flush()
            inner = render_blocks(child, ctx)
            if inner:
                text = '\n\n'.join(inner)
                prefixed = '\n'.join(
                    ('> ' + l) if l.strip() else '>' for l in text.split('\n')
                )
                blocks.append(prefixed)
        elif tag == 'pre':
            flush()
            blocks.append(render_pre(child, ctx))
        elif tag == 'hr':
            flush()
            blocks.append('---')
        elif tag in DROP_SUBTREE:
            continue
        else:
            inline.append(render_inline(child, ctx))
    flush()
    return blocks


def render_list(node, ctx, indent):
    lines = []
    idx = 0
    ordered = node.tag in ('ol',)
    pad_unit = '  '
    for child in node.children:
        if isinstance(child, tuple):
            t = collapse(child[1]).strip()
            if t:
                lines.append(' ' * indent + '- ' + fix_line_start(t))
            continue
        if child.tag != 'li':
            # ul/ol 内的杂散块级元素: 缩进延续, 不加 marker
            for b in render_blocks(child, ctx):
                for bl in b.split('\n'):
                    lines.append((' ' * (indent + 2) + bl).rstrip() if bl.strip() else '')
            continue
        idx += 1
        marker = (str(idx) + '. ') if ordered else '- '
        pad = ' ' * (indent + len(marker))
        inline_run = []
        sub_blocks = []  # (kind, content)
        for c in child.children:
            if isinstance(c, tuple):
                inline_run.append(render_inline(c, ctx))
            elif c.tag in ('ul', 'ol', 'menu', 'dl'):
                sub_blocks.append(('list', render_list(c, ctx, indent + 2)))
            elif c.tag in BLOCK_TAGS:
                sub_blocks.append(('block', render_blocks(c, ctx)))
            else:
                inline_run.append(render_inline(c, ctx))
        text = ''.join(inline_run).strip()
        if text:
            tlines = text.split('\n')
            lines.append(' ' * indent + marker + tlines[0])
            for extra in tlines[1:]:
                lines.append(pad + extra)
        else:
            lines.append((' ' * indent + marker).rstrip())
        for kind, content in sub_blocks:
            if kind == 'list':
                lines.extend(content)
            else:
                for b in content:
                    for bl in b.split('\n'):
                        lines.append((pad + bl).rstrip() if bl.strip() else '')
    return lines


def render_table(node, ctx):
    rows = collect_rows(node)
    if not rows:
        return ''
    grid = []
    for tr in rows:
        cells = []
        for cell in iter_child_elements(tr):
            if cell.tag not in ('td', 'th'):
                continue
            parts = []
            for c in cell.children:
                if isinstance(c, tuple):
                    parts.append(render_inline(c, ctx))
                elif c.tag == 'br':
                    parts.append('<br>')
                elif c.tag in ('ul', 'ol', 'menu', 'dl'):
                    for l in render_list(c, ctx, 0):
                        parts.append(l.strip() + '<br>')
                elif c.tag in BLOCK_TAGS:
                    for b in render_blocks(c, ctx):
                        parts.append(b.replace('\n', '<br>'))
                else:
                    parts.append(render_inline(c, ctx))
            text = ''.join(parts).strip()
            text = text.replace('|', '\\|')
            text = re.sub(r'[ \t]*\n[ \t]*', '<br>', text)
            text = re.sub(r'[ \t]{2,}', ' ', text)
            cells.append(text)
        if cells:
            grid.append(cells)
    if not grid:
        return ''
    ncols = max(len(r) for r in grid)
    grid = [r + [''] * (ncols - len(r)) for r in grid]
    out = ['| ' + ' | '.join(grid[0]) + ' |',
           '| ' + ' | '.join(['---'] * ncols) + ' |']
    for r in grid[1:]:
        out.append('| ' + ' | '.join(r) + ' |')
    return '\n'.join(out)


def render_pre(node, ctx):
    texts = []

    def walk(n):
        for c in n.children:
            if isinstance(c, tuple):
                texts.append(c[1])
            elif c.tag == 'br':
                texts.append('\n')
            else:
                walk(c)

    walk(node)
    code = ''.join(texts).strip('\n')
    if not code.strip():
        return ''
    longest = 0
    for m in re.finditer(r'`+', code):
        longest = max(longest, len(m.group(0)))
    fence = '`' * max(3, longest + 1)
    lang = ''
    cls = node.attrs.get('class') or ''
    if not cls:
        # 语言标识常写在内层 <code class="language-xxx"> 上
        for c in iter_child_elements(node):
            if c.tag in ('code', 'tt', 'kbd', 'samp') and c.attrs.get('class'):
                cls = c.attrs['class']
                break
    m = re.search(r'language-([\w+#-]+)', cls or '')
    if m:
        lang = m.group(1)
    return fence + lang + '\n' + code + '\n' + fence


# ---------------------------------------------------------------- 入口

def convert(html_text, ctx):
    builder = DomBuilder()
    try:
        builder.feed(html_text)
        builder.close()
    except Exception as e:
        sys.stderr.write('警告: HTML 解析异常 (%s), 已尽量容错输出\n' % e)
    blocks = render_blocks(builder.root, ctx)
    # 去掉纯空块
    blocks = [b for b in blocks if b.strip()]
    return '\n\n'.join(blocks)


def read_text(path):
    with open(path, 'rb') as f:
        raw = f.read()
    for enc in ('utf-8-sig', 'utf-8', 'gbk', 'big5'):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode('utf-8', errors='replace')


def main(argv=None):
    ap = argparse.ArgumentParser(description='OneNote HTML -> Markdown (stdlib only)')
    ap.add_argument('input', help="HTML 文件路径, '-' 表示 stdin")
    ap.add_argument('-o', '--output', help='输出 Markdown 路径 (缺省输出到 stdout)')
    ap.add_argument('--assets-dir', help='图片输出目录 (默认: <输出文件名>.assets)')
    args = ap.parse_args(argv)

    if args.input == '-':
        html_text = sys.stdin.read()
    else:
        if not os.path.isfile(args.input):
            sys.stderr.write('错误: 输入文件不存在: %s\n' % args.input)
            return 2
        html_text = read_text(args.input)

    out_path = args.output
    assets_dir = args.assets_dir
    if out_path and not assets_dir:
        base = os.path.splitext(os.path.basename(out_path))[0]
        assets_dir = os.path.join(os.path.dirname(os.path.abspath(out_path)), base + '.assets')

    ctx = RenderContext(assets_dir=assets_dir, out_path=out_path)
    md = convert(html_text, ctx)

    if out_path:
        with open(out_path, 'w', encoding='utf-8') as f:
            f.write(md)
            if md and not md.endswith('\n'):
                f.write('\n')
    else:
        sys.stdout.write(md + '\n')

    sys.stderr.write('转换完成: 块数=%d 图片=%d 字符=%d\n'
                     % (md.count('\n\n') + 1 if md else 0, ctx.images, len(md)))
    return 0


if __name__ == '__main__':
    sys.exit(main())
