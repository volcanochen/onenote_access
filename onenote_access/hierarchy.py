# -*- coding: utf-8 -*-
"""OneNote 层级 XML 解析与页面选择 (纯标准库, 可离线测试)。

OneNote 各版本默认命名空间不同, 因此全部用 ElementTree 的 {*}tag 通配匹配,
不依赖具体命名空间 URI。
"""

import fnmatch
import xml.etree.ElementTree as ET

KNOWN_TYPES = ('Notebook', 'SectionGroup', 'Section', 'Page')


def _local(tag):
    """'{ns}Name' -> 'Name'。"""
    if isinstance(tag, str) and tag.startswith('{'):
        return tag.rsplit('}', 1)[-1]
    return tag


def parse_hierarchy(xml_text):
    """GetHierarchy XML -> 节点树列表。

    节点: {type, name, id, created, modified, path, children}
    包装元素 (如 Notebooks 根) 透明上钻; 始终返回 list。
    """
    root = ET.fromstring(xml_text)
    return _convert(root, '')


def _convert(el, parent_path):
    t = _local(el.tag)
    if t in KNOWN_TYPES:
        name = el.get('name') or ''
        node = {
            'type': t,
            'name': name,
            'id': el.get('ID') or el.get('id') or '',
            'created': el.get('dateTime') or '',
            'modified': el.get('lastModifiedTime') or '',
            'path': _join_path(parent_path, name),
            'children': [],
        }
        for child in el:
            node['children'].extend(_convert(child, node['path']))
        return [node]
    # 包装元素: 上钻
    nodes = []
    for child in el:
        nodes.extend(_convert(child, parent_path))
    return nodes


def _join_path(parent, name):
    seg = _safe_segment(name)
    if not parent:
        return seg
    return parent + '\\' + seg


def _safe_segment(name, max_len=60):
    s = (name or '').strip()
    for ch in '\\/:*?"<>|':
        s = s.replace(ch, '_')
    s = ''.join(c for c in s if ord(c) >= 32).strip().rstrip('. ')
    return s[:max_len].strip() or 'untitled'


def iter_sections(node, out=None):
    """递归收集容器(笔记本/分区组)下的所有 Section, 自动展开 SectionGroup。"""
    if out is None:
        out = []
    for child in node.get('children', []):
        if child['type'] == 'Section':
            out.append(child)
        elif child['type'] == 'SectionGroup':
            iter_sections(child, out)
    return out


def _match(name, pattern):
    """通配符匹配: pattern 为空/'*' 恒真; 否则不区分大小写的 fnmatch。"""
    if not pattern or pattern == '*':
        return True
    if not name:
        return False
    return fnmatch.fnmatch(name.lower(), pattern.lower())


def find_pages(tree, notebook='*', section='*', page='*'):
    """按 名称过滤条件 (支持 * ? 通配符) 选出页面, 返回页面元数据列表。"""
    result = []
    for nb in tree:
        if nb['type'] != 'Notebook' or not _match(nb['name'], notebook):
            continue
        for sec in iter_sections(nb):
            if not _match(sec['name'], section):
                continue
            for pg in sec.get('children', []):
                if pg['type'] != 'Page' or not _match(pg['name'], page):
                    continue
                title = pg['name'] or _untitled(pg['id'])
                result.append({
                    'notebook': nb['name'],
                    'section': sec['name'],
                    'title': title,
                    'page_id': pg['id'],
                    'created': pg['created'],
                    'modified': pg['modified'],
                    'section_path': sec['path'],
                })
    return result


def find_notebooks(tree):
    return [n for n in tree if n['type'] == 'Notebook']


def _untitled(page_id):
    raw = (page_id or '').strip('{}')
    return '未命名页-' + raw[:8]


# ---------------------------------------------------------- 页面内容 XML

def extract_page_html(page_xml):
    """从 GetPageContent XML 提取页面 HTML 字符串。

    优先取 one:HTML 元素的 CDATA 文本; 无则回退拼接 one:T 文本 (旧式纯文本页)。
    返回 (html, empty_bool)。
    """
    if not page_xml or not page_xml.strip():
        return '', True
    root = ET.fromstring(page_xml)
    fragments = []
    for el in root.iter():
        if _local(el.tag) == 'HTML':
            text = el.text
            if text is None or not text.strip():
                # 非 CDATA: 序列化子元素兜底
                parts = []
                for child in el:
                    parts.append(ET.tostring(child, encoding='unicode'))
                text = ''.join(parts)
            if text and text.strip():
                fragments.append(text)
    if fragments:
        return '\r\n\r\n'.join(fragments), False
    # 旧式纯文本页: 拼 one:T
    parts = [el.text or '' for el in root.iter() if _local(el.tag) == 'T']
    text = '\n'.join(p for p in parts if p.strip())
    return (text, not bool(text.strip()))


def page_title_from_xml(page_xml, fallback=''):
    if not page_xml or not page_xml.strip():
        return fallback
    try:
        root = ET.fromstring(page_xml)
    except ET.ParseError:
        return fallback
    for el in root.iter():
        if _local(el.tag) == 'T' and el.text and el.text.strip():
            return el.text.strip()
    return fallback


def yaml_quote(value):
    v = value or ''
    if v == '' or any(c in v for c in ':#[]{}&*!|>%@`\'"') or \
            v != v.strip() or v.startswith('-') or v.startswith('?'):
        return '"' + v.replace('"', '""') + '"'
    return v


def build_frontmatter(meta, link=None):
    lines = ['---',
             'title: ' + yaml_quote(meta.get('title', '')),
             'notebook: ' + yaml_quote(meta.get('notebook', '')),
             'section: ' + yaml_quote(meta.get('section', '')),
             'created: ' + (meta.get('created') or ''),
             'modified: ' + (meta.get('modified') or ''),
             'page_id: "' + (meta.get('page_id') or '') + '"']
    if link:
        lines.append('onenote_link: "' + link.replace('"', '""') + '"')
    lines.append('source: OneNote desktop (COM API, read-only)')
    lines.append('---')
    lines.append('')
    return '\r\n'.join(lines)
