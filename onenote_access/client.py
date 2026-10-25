# -*- coding: utf-8 -*-
"""OneNoteClient - 面向调用方的高层 API (库入口)。

用法:
    from onenote_access import OneNoteClient
    client = OneNoteClient()
    tree = client.notebooks()                    # 层级树
    pages = client.find_pages(notebook='工作*')   # 页面元数据
    page = client.get_page(pages[0]['page_id'])  # 单页 markdown
    client.export(out_dir='D:\\backup', notebook='*')  # 批量导出
"""

import os

from . import exporter, hierarchy
from .com_backend import HS_PAGES, OneNoteCom, OneNoteComError

# find_pages/search 单次安全上限
MAX_MATCH_DEFAULT = 20
SEARCH_SCAN_CAP = 500


class OneNoteClient(object):
    """OneNote 桌面版只读客户端。

    复用单个 OneNoteCom 连接; 所有方法在 COM 不可用时抛 OneNoteComError
    (get_page 对加密分区返回 None)。
    """

    def __init__(self, com=None):
        self._com = com or OneNoteCom()

    # -------------------------------------------------------------- 层级

    def _tree(self):
        return hierarchy.parse_hierarchy(self._com.get_hierarchy())

    def notebooks(self):
        """返回笔记本/分区组/分区/页面完整层级树 (list[dict])。"""
        return self._tree()

    def find_pages(self, notebook='*', section='*', page='*', limit=None):
        """按通配符条件选择页面, 返回元数据 list[dict]。

        limit: 上限保护 (默认 MAX_MATCH_DEFAULT=20; 传 None 表示不限制,
        服务端会强制限制)。
        """
        pages = hierarchy.find_pages(self._tree(), notebook, section, page)
        if limit is not None:
            pages = pages[:limit]
        return pages

    # -------------------------------------------------------------- 单页

    def get_page(self, page_id, images=True):
        """读取单页并转为 Markdown。

        返回 dict: {title, notebook, section, created, modified, page_id,
                    onenote_link, markdown, degraded}
        加密分区等读取失败返回 None。
        """
        page_xml = self._com.get_page_xml(page_id, with_images=images)
        if not page_xml:
            return None
        html, empty = hierarchy.extract_page_html(page_xml)
        meta = self._meta_for(page_id)
        if meta is None:
            meta = {'notebook': '', 'section': '', 'title': '', 'page_id': page_id,
                    'created': '', 'modified': '', 'section_path': ''}
        title = hierarchy.page_title_from_xml(page_xml, meta['title'])
        meta['title'] = title or meta['title']
        md, degraded = exporter.html_to_markdown(html)
        return {
            'notebook': meta['notebook'],
            'section': meta['section'],
            'title': meta['title'],
            'created': meta['created'],
            'modified': meta['modified'],
            'page_id': page_id,
            'onenote_link': self._com.get_page_hyperlink(page_id),
            'markdown': md,
            'degraded': degraded or empty,
        }

    def _meta_for(self, page_id):
        """在层级树中反查页面元数据 (找不到返回 None)。"""
        pages = hierarchy.find_pages(self._tree())
        for p in pages:
            if p['page_id'] == page_id:
                return p
        return None

    # -------------------------------------------------------------- 搜索

    def search_pages(self, text, notebook='*', section='*', page='*',
                     limit=MAX_MATCH_DEFAULT):
        """全文筛选: 拉取候选页内容 (不含图片) 并做大小写不敏感的字面匹配。"""
        text_l = text.lower()
        if not text_l:
            return []
        candidates = hierarchy.find_pages(self._tree(), notebook, section, page)[:SEARCH_SCAN_CAP]
        hits = []
        for p in candidates:
            xml = self._com.get_page_xml(p['page_id'], with_images=False)
            if xml and text_l in xml.lower():
                hits.append(p)
                if len(hits) >= limit:
                    break
        return hits

    # -------------------------------------------------------------- 导出

    def export(self, out_dir, notebook='*', section='*', page='*',
               search=None, images=True, all_pages=False,
               max_without_confirm=MAX_MATCH_DEFAULT):
        """批量导出为 Markdown 文件, 返回 {out_dir, exported, failed, index, files}。

        安全上限: 匹配数超过 max_without_confirm 且未显式 all_pages=True 时
        抛 ValueError (与 CLI 的 -All 行为一致)。
        """
        selected = hierarchy.find_pages(self._tree(), notebook, section, page)
        if search:
            text_l = search.lower()
            selected = [p for p in selected[:SEARCH_SCAN_CAP]
                        if (xml := self._com.get_page_xml(p['page_id'], with_images=False))
                        and text_l in xml.lower()]
        if len(selected) > max_without_confirm and not all_pages:
            raise ValueError(
                '匹配到 %d 个页面, 超过安全上限 %d。请收窄过滤条件, 或显式传 all_pages=True。'
                % (len(selected), max_without_confirm))
        os.makedirs(out_dir, exist_ok=True)
        results, failed = [], 0
        for p in selected:
            page_xml = self._com.get_page_xml(p['page_id'], with_images=images)
            if not page_xml:
                failed += 1
                continue
            html, empty = hierarchy.extract_page_html(page_xml)
            if empty and not html:
                failed += 1
                continue
            rel = exporter.write_page_markdown(
                p, html, out_dir, with_images=images,
                link=self._com.get_page_hyperlink(p['page_id']))
            results.append((p, rel))
        index = exporter.write_index(out_dir, results, failed)
        return {
            'out_dir': out_dir,
            'exported': len(results),
            'failed': failed,
            'index': index,
            'files': [rel for _, rel in results],
        }

    # -------------------------------------------------------------- 探测

    def health(self):
        return {'com_available': self._com.is_available()}
