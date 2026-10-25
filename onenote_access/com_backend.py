# -*- coding: utf-8 -*-
"""OneNote 桌面版 COM 后端封装 (只读, 仅 Windows)。

通过 pywin32 (win32com) 调用 OneNote.Application, 只使用三个只读方法:
    GetHierarchy / GetPageContent / GetHyperlinkToObject

不同 OneNote 版本对可选参数/返回值的互操作签名有差异, 因此每个方法都按
多种调用样式逐一尝试, 取第一个返回非空结果的样式 (与 Extract-OneNote.ps1 同思路)。
"""

import logging

log = logging.getLogger('onenote_access.com')

# HierarchyScope
HS_SELF = 0
HS_CHILDREN = 1
HS_NOTEBOOKS = 2
HS_SECTIONS = 3
HS_PAGES = 4

# PageInfo
PI_BASIC = 0
PI_BINARY_DATA = 1


class OneNoteComError(RuntimeError):
    """COM 调用失败 (未安装桌面版 / 未启动 / 权限或环境限制)。"""


class OneNoteCom(object):
    """惰性连接的 OneNote COM 只读封装。进程内复用单个 Application 对象。"""

    def __init__(self):
        self._app = None
        self._connected = False

    # -------------------------------------------------------------- 连接

    def connect(self):
        """连接 OneNote COM。失败抛 OneNoteComError, 携带可操作的排查提示。"""
        if self._connected:
            return self._app
        try:
            import win32com.client  # 延迟导入: 非 Windows / 未装 pywin32 时也能 import 本模块
        except ImportError as e:
            raise OneNoteComError(
                'pywin32 未安装, 无法调用 OneNote COM。请在 Windows 上执行: '
                'pip install pywin32 (原始错误: %s)' % e)
        try:
            self._app = win32com.client.Dispatch('OneNote.Application')
        except Exception as e:
            raise OneNoteComError(
                '无法创建 OneNote.Application COM 对象。请确认: '
                '1) 已安装 OneNote 桌面版 (ONENOTE.EXE); '
                '2) OneNote 已启动并登录; '
                '3) 本进程与 OneNote 位数一致 (均为 64 位或 32 位)。'
                '原始错误: %s' % e)
        self._connected = True
        return self._app

    def is_available(self):
        """探测 COM 是否真正可用 (冒烟测试: 真调一次 GetHierarchy)。"""
        try:
            xml = self.get_hierarchy()
            return bool(xml)
        except Exception as e:
            log.debug('COM 冒烟测试失败: %s', e)
            return False

    # -------------------------------------------------------------- 内部

    @staticmethod
    def _first_str(results, last_error):
        """从多种调用样式的结果中取第一个非空字符串; 全部失败则抛错。"""
        for r, err in results:
            if isinstance(r, tuple) and r:
                r = r[0]
            if isinstance(r, str) and r.strip():
                return r
            if err is not None:
                last_error = err
        raise OneNoteComError('所有调用样式均未返回数据。最后错误: %s' % last_error)

    def _try_variants(self, method, arg_variants):
        """依次尝试多组参数调用 COM 方法, 收集 (结果, 异常)。"""
        results = []
        last_error = None
        for args in arg_variants:
            try:
                r = getattr(self._app, method)(*args)
                results.append((r, None))
            except Exception as e:
                results.append((None, e))
                last_error = e
        return self._first_str(results, last_error)

    # -------------------------------------------------------------- 只读 API

    def get_hierarchy(self):
        """返回完整层级 XML (笔记本/分区组/分区/页面)。"""
        self.connect()
        return self._try_variants('GetHierarchy', [
            ('', HS_PAGES),          # 标准签名: (startNodeID, hsScope) -> retval
            ('', HS_PAGES, 2),       # 部分版本多一个 xmlSchema 参数
        ])

    def get_page_xml(self, page_id, with_images=True):
        """返回页面内容 XML。with_images=True 时内联 base64 图片。"""
        self.connect()
        if with_images:
            variants = [(page_id, PI_BINARY_DATA, 2), (page_id, PI_BINARY_DATA),
                        (page_id, PI_BASIC, 2), (page_id, PI_BASIC), (page_id,)]
        else:
            variants = [(page_id, PI_BASIC, 2), (page_id, PI_BASIC), (page_id,)]
        try:
            return self._try_variants('GetPageContent', variants)
        except OneNoteComError:
            # 加密分区等场景返回空不算致命, 与上层约定: 返回 None
            return None

    def get_page_hyperlink(self, page_id):
        """返回 onenote:// 跳转链接; 失败返回 None (可选能力)。"""
        try:
            self.connect()
            r = getattr(self._app, 'GetHyperlinkToObject')(page_id)
            if isinstance(r, tuple) and r:
                r = r[0]
            if isinstance(r, str) and r.strip():
                return r
        except Exception as e:
            log.debug('GetHyperlinkToObject 失败 (忽略): %s', e)
        return None
