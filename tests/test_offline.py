# -*- coding: utf-8 -*-
"""离线自测: 不连 OneNote / 不依赖 pywin32。

覆盖: 层级 XML 解析、通配符过滤、页面内容 XML -> HTML 提取、
HTML -> Markdown 转换、frontmatter 生成、文件名安全化与去重、导出管线、
索引生成、API 路由 (COM 不可用路径)。
"""

import json
import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from onenote_access import exporter, hierarchy  # noqa: E402
from onenote_access.server import ApiState, _handle  # noqa: E402

HIER_XML = '''<?xml version="1.0" encoding="utf-8"?>
<one:Notebooks xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote">
  <one:Notebook name="工作笔记" ID="{NB-0001}">
    <one:Section name="会议记录" ID="{SEC-0001}">
      <one:Page name="周会 2026-10-10" ID="{PG-0001}" dateTime="2026-10-10T09:00:00.000Z" lastModifiedTime="2026-10-10T10:00:00.000Z"/>
      <one:Page name="" ID="{PG-0002}"/>
    </one:Section>
    <one:SectionGroup name="归档">
      <one:Section name="2025" ID="{SEC-0002}">
        <one:Page name="旧页" ID="{PG-0003}"/>
      </one:Section>
    </one:SectionGroup>
  </one:Notebook>
  <one:Notebook name="个人" ID="{NB-0002}">
    <one:Section name="杂项" ID="{SEC-0003}">
      <one:Page name="读书清单" ID="{PG-0004}"/>
    </one:Section>
  </one:Notebook>
</one:Notebooks>'''

PAGE_XML = '''<?xml version="1.0" encoding="utf-8"?>
<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" name="周会" ID="{PG-0001}">
  <one:Title><one:OE><one:T><![CDATA[周会]]></one:T></one:OE></one:Title>
  <one:Outline>
    <one:OEChildren>
      <one:OE><one:HTML><![CDATA[<h1>2026-10-10 项目周会</h1><p>参会人：<b>张三</b>、<i>李四</i></p><ul><li>议题一<ul><li>结论：归入内核驱动子系统</li></ul></li></ul><table><tr><td>需求</td><td>测试用例</td></tr><tr><td>SM9 认证</td><td>TC-AUTH-01</td></tr></table><pre><code class="language-c">int main(void) { return 0; }</code></pre><p><img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==" alt="pixel"/></p>]]></one:HTML></one:OE>
    </one:OEChildren>
  </one:Outline>
</one:Page>'''

_PIXEL_PNG = ('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQ'
              'GAhKmMIQAAAABJRU5ErkJggg==')


class Checker(object):
    def __init__(self):
        self.ok = 0
        self.fail = 0

    def check(self, name, cond, detail=''):
        if cond:
            self.ok += 1
            if self.verbose:
                print('[PASS] %s' % name)
        else:
            self.fail += 1
            print('[FAIL] %s  %s' % (name, detail))

    verbose = True


def run_all(verbose=True):
    c = Checker()
    c.verbose = verbose

    # --- 1. 层级解析 ---
    tree = hierarchy.parse_hierarchy(HIER_XML)
    c.check('层级解析: 2 个笔记本', len(tree) == 2, str(len(tree)))
    c.check('层级解析: 首个笔记本名', tree[0]['name'] == '工作笔记', tree[0]['name'])
    secs = hierarchy.iter_sections(tree[0])
    c.check('SectionGroup 展开: 2 个分区', len(secs) == 2, str(len(secs)))
    c.check('分区组保留在路径中', secs[1]['path'] == '工作笔记\\归档\\2025', secs[1]['path'])

    # --- 2. 过滤 ---
    pages = hierarchy.find_pages(tree)
    c.check('页面收集: 4 页', len(pages) == 4, str(len(pages)))
    c.check('未命名页占位', pages[1]['title'].startswith('未命名页-'), pages[1]['title'])
    c.check('过滤: 精确笔记本', len(hierarchy.find_pages(tree, notebook='个人')) == 1)
    c.check('过滤: 通配符', len(hierarchy.find_pages(tree, section='会议*', page='周会*')) == 1)
    c.check('过滤: 大小写不敏感', len(hierarchy.find_pages(tree, notebook='个*')) == 1)
    c.check('过滤: 无匹配', len(hierarchy.find_pages(tree, notebook='不存在')) == 0)

    # --- 3. 页面内容 XML -> HTML ---
    html, empty = hierarchy.extract_page_html(PAGE_XML)
    c.check('页面 XML: 提取 HTML 非空', bool(html.strip()) and not empty)
    c.check('页面 XML: 标题提取', hierarchy.page_title_from_xml(PAGE_XML) == '周会')
    html2, empty2 = hierarchy.extract_page_html('')
    c.check('页面 XML: 空 XML -> (空, True)', empty2 and html2 == '')

    # --- 4. HTML -> Markdown ---
    tmp = tempfile.mkdtemp(prefix='onenote_access_test_')
    try:
        assets = os.path.join(tmp, 'assets')
        md, degraded = exporter.html_to_markdown(html, assets_dir=assets)
        c.check('HTML->MD: 未降级', not degraded)
        c.check('HTML->MD: 标题', '# 2026-10-10 项目周会' in md, md[:120])
        c.check('HTML->MD: 加粗/斜体', '**张三**' in md and '*李四*' in md)
        c.check('HTML->MD: 嵌套列表', '归入内核驱动子系统' in md)
        c.check('HTML->MD: 表格', 'TC-AUTH-01' in md and '|' in md)
        c.check('HTML->MD: 代码块语言', '```c' in md)
        c.check('HTML->MD: 图片落盘', len(os.listdir(assets)) >= 1)
        # 降级路径
        md2, degraded2 = exporter.html_to_markdown('<p>纯文本 <b>加粗</b></p>', assets_dir=None)
        c.check('降级路径: 文本保留', ('加粗' in md2), md2)

        # --- 5. 导出管线 ---
        out = os.path.join(tmp, 'export')
        meta = pages[0]
        rel = exporter.write_page_markdown(meta, html, out, with_images=True, link=None)
        md_path = os.path.join(out, rel)
        c.check('导出: 文件生成', os.path.exists(md_path), rel)
        with open(md_path, encoding='utf-8', newline='') as f:
            content = f.read()
        c.check('导出: frontmatter', content.startswith('---\r\ntitle: ') and 'page_id: "{PG-0001}"' in content)
        c.check('导出: 文件名清洗', not any(ch in rel for ch in '?*:|"<>'))
        rel2 = exporter.write_page_markdown(meta, '<p>第二页</p>', out, with_images=False)
        c.check('导出: 同名去重', rel != rel2, '%s / %s' % (rel, rel2))
        idx = exporter.write_index(out, [(meta, rel)], 0)
        with open(idx, encoding='utf-8') as f:
            idx_text = f.read()
        c.check('导出: 索引 README', 'OneNote 导出索引' in idx_text and '[周会 2026-10-10]' in idx_text)

        # frontmatter 转义
        fm = hierarchy.build_frontmatter({'title': '周会: 2026', 'notebook': 'nb',
                                          'section': 'sec', 'page_id': '{X}',
                                          'created': '', 'modified': ''})
        c.check('frontmatter: 冒号加引号', 'title: "周会: 2026"' in fm, fm)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    # --- 6. API 路由 (COM 不可用 -> 优雅降级) ---
    from onenote_access.com_backend import OneNoteComError as _ComErr

    class FakeCom(object):
        # 真实 OneNoteCom 只会抛 OneNoteComError, 这里如实模拟
        def get_hierarchy(self):
            raise _ComErr('COM blocked')

        def is_available(self):
            return False

    from onenote_access.client import OneNoteClient
    state = ApiState(token=None, export_root=tempfile.mkdtemp(),
                     client=OneNoteClient(com=FakeCom()))
    status, _, data = _handle(state, 'GET', '/health', {}, None)
    body = json.loads(data.decode('utf-8'))
    c.check('API /health: ok', status == 200 and body['ok'] is True)
    c.check('API /health: com_available=false', body['com_available'] is False)
    status, _, data = _handle(state, 'GET', '/api/v1/notebooks', {}, None)
    c.check('API /notebooks: COM 失败 -> 503', status == 503, str(status))
    status, _, _ = _handle(state, 'GET', '/api/v1/unknown', {}, None)
    c.check('API 未知端点 -> 404', status == 404, str(status))
    status, _, _ = _handle(state, 'GET', '/api/v1/pages/{bad id!}', {}, None)
    c.check('API 非法 page id -> 400', status == 400, str(status))

    # out_dir 越界防护
    try:
        state.resolve_export_dir('../../outside')
        c.check('API export: 目录穿越被拦截', False)
    except ValueError:
        c.check('API export: 目录穿越被拦截', True)

    # --- 7. doctor 诊断工具 ---
    import struct as _struct
    from onenote_access import doctor
    from onenote_access.com_backend import _com_error_hint, OneNoteComError as _OCE

    def _fake_pe(machine):
        data = bytearray(0x100)
        data[0:2] = b'MZ'
        _struct.pack_into('<I', data, 0x3C, 0x80)
        _struct.pack_into('<H', data, 0x84, machine)
        p = os.path.join(tmpdir7, 'fake_onenote.exe')
        with open(p, 'wb') as f:
            f.write(bytes(data))
        return p

    tmpdir7 = tempfile.mkdtemp(prefix='onenote_doctor_')
    try:
        c.check('doctor: 64 位 PE 识别', doctor.exe_bitness(_fake_pe(0x8664)) == 64)
        c.check('doctor: 32 位 PE 识别', doctor.exe_bitness(_fake_pe(0x14C)) == 32)
        c.check('doctor: 非 PE 文件', doctor.exe_bitness(__file__) is None)
        c.check('doctor: python 位数合理', doctor.python_bitness() in (32, 64))
    finally:
        shutil.rmtree(tmpdir7, ignore_errors=True)

    class _FakeComErr(Exception):
        def __init__(self, hr):
            Exception.__init__(self, hr, 'x', None, None)

    c.check('错误映射: 0x80080005 -> 启动提示',
            'OneNote 未在运行' in _com_error_hint(_FakeComErr(-2146959355)))
    c.check('错误映射: 0x80040154 -> 注册提示',
            '未注册' in _com_error_hint(_FakeComErr(-2147221164)))
    c.check('错误映射: 未知 HRESULT -> 空', _com_error_hint(_FakeComErr(-1)) == '')
    try:
        raise _OCE('无法创建 OneNote.Application COM 对象')
    except _OCE:
        c.check('OneNoteComError 可抛可捕', True)

    return c.ok, c.fail


if __name__ == '__main__':
    ok, fail = run_all()
    print('\n自测结果: %d 通过, %d 失败' % (ok, fail))
    raise SystemExit(1 if fail else 0)
