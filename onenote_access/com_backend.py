# -*- coding: utf-8 -*-
"""OneNote 桌面版 COM 后端 (只读, 仅 Windows)。

双后端设计:
    1. comtypes vtable 直调 (首选) —— 绕过 IDispatch::Invoke, 免受
       "Office C2R typelib 注册损坏 -> ITypeInfo 分发失败" 的影响;
       需要 comtypes 库。
    2. pywin32 动态分发 (后备) —— 在 typelib 注册正常的机器上可用。

只使用三个只读方法: GetHierarchy / GetPageContent / GetHyperlinkToObject。
2013 schema 下参数顺序为 [in]..., [out]... 交错 (out 不在末尾), 槽位严格
按 typelib 枚举顺序对齐。
"""

import logging

log = logging.getLogger('onenote_access.com')

# HierarchyScope
HS_PAGES = 4
# PageInfo
PI_BASIC = 0
PI_BINARY_DATA = 1
# XMLSchema (2013)
XS_CURRENT = 2

# HRESULT
HR_LIB_NOT_REGISTERED = -2147319779   # 0x8002801D TYPE_E_LIBNOTREGISTERED
HR_SERVER_EXEC_FAILURE = -2146959355  # 0x80080005 CO_E_SERVER_EXEC_FAILURE

HRESULT_HINTS = {
    HR_LIB_NOT_REGISTERED:
        '库没有注册 —— OneNote typelib 缺少 win64 平台键 (Office Click-to-Run '
        '已知问题)。运行 python -m onenote_access fixregistry 修复。',
    HR_SERVER_EXEC_FAILURE:
        '服务器运行失败 (常为瞬时错误, 本组件已自动重试)。若持续出现: '
        '请先手动打开 OneNote 桌面版并保持运行; 再运行 python -m onenote_access doctor。',
    -2147221164:  # 0x80040154 REGDB_E_CLASSNOTREG
        'COM 类未注册。通常表示未安装 OneNote 桌面版 (UWP 版不支持), '
        '或 Python 与 Office 位数不一致。',
    -2147221005:  # 0x800401F3 CO_E_CLASSSTRING
        '无效的 ProgID。OneNote 桌面版可能未安装或安装损坏, 尝试修复 Office。',
}


def _hint_for(exc):
    """从异常提取 HRESULT 并给出可操作提示 (无法提取返回空串)。"""
    args = getattr(exc, 'args', []) or []
    hr = args[0] if args else None
    if isinstance(hr, int) and hr in HRESULT_HINTS:
        return HRESULT_HINTS[hr]
    return ''


class OneNoteComError(RuntimeError):
    """COM 调用失败 (未安装桌面版 / 未启动 / 权限或环境限制)。"""


# ================================================================ comtypes 后端

_IApplication = None  # 惰性构建的接口类缓存


def _get_IApplication():
    """定义 (缓存) OneNote 2013 IApplication 接口。

    vtable 槽位严格按 typelib 枚举顺序; 未使用的槽用占位方法填充。
    参数顺序为 2013 schema 的实际顺序: out 参数在中间而非末尾。
    """
    global _IApplication
    if _IApplication is not None:
        return _IApplication
    import comtypes
    from comtypes import COMMETHOD, GUID, HRESULT
    from comtypes.automation import IDispatch, BSTR

    dummy = (['in'], ctypes.c_int, 'reserved')

    class IApplication(IDispatch):
        _iid_ = GUID('{452AC71A-B655-4967-A208-A4CC39DD7949}')
        _methods_ = [
            # slot 7: GetHierarchy(id, scope, out xml, xsSchema)
            COMMETHOD([], HRESULT, 'GetHierarchy',
                      (['in'], BSTR, 'bstrStartNodeID'),
                      (['in'], ctypes.c_int, 'hsScope'),
                      (['out'], ctypes.POINTER(BSTR), 'pbstrHierarchyXmlOut'),
                      (['in'], ctypes.c_int, 'xsSchema')),
            # slots 8-13 占位 (UpdateHierarchy/OpenHierarchy/DeleteHierarchy/
            # CreateNewPage/CloseNotebook/GetHierarchyParent)
            COMMETHOD([], HRESULT, 'r8', dummy),
            COMMETHOD([], HRESULT, 'r9', dummy),
            COMMETHOD([], HRESULT, 'r10', dummy),
            COMMETHOD([], HRESULT, 'r11', dummy),
            COMMETHOD([], HRESULT, 'r12', dummy),
            COMMETHOD([], HRESULT, 'r13', dummy),
            # slot 14: GetPageContent(id, out xml, pageInfo, xsSchema)
            COMMETHOD([], HRESULT, 'GetPageContent',
                      (['in'], BSTR, 'bstrPageId'),
                      (['out'], ctypes.POINTER(BSTR), 'pbstrPageXmlOut'),
                      (['in'], ctypes.c_int, 'piPageInfoToGet'),
                      (['in'], ctypes.c_int, 'xsSchema')),
            # slots 15-21 占位 (UpdatePageContent/GetBinaryPageContent/
            # DeletePageContent/NavigateTo/NavigateToUrl/Publish/OpenPackage)
            COMMETHOD([], HRESULT, 'r15', dummy),
            COMMETHOD([], HRESULT, 'r16', dummy),
            COMMETHOD([], HRESULT, 'r17', dummy),
            COMMETHOD([], HRESULT, 'r18', dummy),
            COMMETHOD([], HRESULT, 'r19', dummy),
            COMMETHOD([], HRESULT, 'r20', dummy),
            COMMETHOD([], HRESULT, 'r21', dummy),
            # slot 22: GetHyperlinkToObject(pageId, out link)
            COMMETHOD([], HRESULT, 'GetHyperlinkToObject',
                      (['in'], BSTR, 'bstrPageId'),
                      (['out'], ctypes.POINTER(BSTR), 'pbstrHyperlinkLocation')),
        ]

    _IApplication = IApplication
    return _IApplication


import ctypes  # noqa: E402  (上面 _get_IApplication 引用了它)


class OneNoteCom(object):
    """惰性连接的 OneNote COM 只读封装。进程内复用单个 Application 对象。"""

    def __init__(self):
        self._app = None
        self._mode = None  # 'comtypes' | 'win32com'
        self._connected = False

    # -------------------------------------------------------------- 连接

    def connect(self):
        """连接 OneNote COM。失败抛 OneNoteComError, 携带可操作的排查提示。"""
        if self._connected:
            return self._app
        try:
            self._connect_comtypes()
            self._mode = 'comtypes'
        except ImportError:
            self._connect_win32com()
            self._mode = 'win32com'
        except Exception as e:
            # comtypes 存在但失败 (典型: typelib 注册缺失) -> 尝试 HKCU 自愈后重试,
            # 仍失败再走 pywin32 后备; 都失败则抛出带提示的错误
            last = e
            try:
                from . import fixregistry
                fixregistry.ensure_hkcu_keys()
                self._connect_comtypes()
                self._mode = 'comtypes'
                log.info('HKCU typelib 自愈成功')
            except Exception:
                try:
                    self._connect_win32com()
                    self._mode = 'win32com'
                except Exception as e2:
                    raise OneNoteComError(self._connect_error_text(e2, last))
        self._connected = True
        return self._app

    @staticmethod
    def _connect_error_text(*errors):
        hints = []
        for e in errors:
            h = _hint_for(e)
            if h and h not in hints:
                hints.append(h)
        return ('无法连接 OneNote COM。请确认: '
                '1) 已安装 OneNote 桌面版 (ONENOTE.EXE) 且已启动登录; '
                '2) Python 与 Office 位数一致 (均为 64 位); '
                + ('提示: ' + ' / '.join(hints) if hints else '')
                + '\n运行 python -m onenote_access doctor 可自动诊断。'
                '原始错误: ' + ' | '.join(str(e)[:160] for e in errors))

    def _connect_comtypes(self):
        import comtypes.client
        app = comtypes.client.CreateObject('OneNote.Application',
                                           interface=_get_IApplication())
        self._app = app

    def _connect_win32com(self):
        import time
        import win32com.client
        last = None
        for attempt in range(3):
            try:
                self._app = win32com.client.Dispatch('OneNote.Application')
                return
            except Exception as e:
                last = e
                hr = (getattr(e, 'args', [None]) or [None])[0]
                if hr == HR_SERVER_EXEC_FAILURE and attempt < 2:
                    time.sleep(1.5)  # 瞬态错误重试
        raise last

    def is_available(self):
        """探测 COM 是否真正可用 (冒烟测试: 真调一次 GetHierarchy)。"""
        try:
            return bool(self.get_hierarchy())
        except Exception as e:
            log.debug('COM 冒烟测试失败: %s', e)
            return False

    # -------------------------------------------------------------- 只读 API

    def get_hierarchy(self):
        """返回完整层级 XML (笔记本/分区组/分区/页面)。"""
        self.connect()
        try:
            if self._mode == 'comtypes':
                return self._app.GetHierarchy('', HS_PAGES, XS_CURRENT)
            return self._win32_first('GetHierarchy',
                                     [('', HS_PAGES), ('', HS_PAGES, XS_CURRENT)])
        except Exception as e:
            hint = _hint_for(e)
            raise OneNoteComError('GetHierarchy 失败: %s%s'
                                  % (str(e)[:300],
                                     ('\n提示: ' + hint) if hint else ''))

    def get_page_xml(self, page_id, with_images=True):
        """返回页面内容 XML; 加密分区等不可读场景返回 None。"""
        self.connect()
        pi = PI_BINARY_DATA if with_images else PI_BASIC
        try:
            if self._mode == 'comtypes':
                return self._app.GetPageContent(page_id, pi, XS_CURRENT)
            variants = ([(page_id, pi, XS_CURRENT), (page_id, pi), (page_id,)]
                        if with_images else
                        [(page_id, pi, XS_CURRENT), (page_id, pi), (page_id,)])
            return self._win32_first('GetPageContent', variants)
        except Exception:
            # 加密分区/已移动页面: 服务器返回业务错误, 按约定返回 None
            return None

    def get_page_hyperlink(self, page_id):
        """返回 onenote:// 跳转链接; 失败返回 None (可选能力)。"""
        try:
            self.connect()
            if self._mode == 'comtypes':
                return self._app.GetHyperlinkToObject(page_id)
            r = self._app.GetHyperlinkToObject(page_id)
            if isinstance(r, tuple) and r:
                r = r[0]
            return r if (isinstance(r, str) and r.strip()) else None
        except Exception as e:
            log.debug('GetHyperlinkToObject 失败 (忽略): %s', e)
            return None

    # -------------------------------------------------------------- 内部

    def _win32_first(self, method, arg_variants):
        """pywin32 后备: 依次尝试多组参数, 取第一个非空字符串结果。"""
        results, last_error = [], None
        for args in arg_variants:
            try:
                r = getattr(self._app, method)(*args)
                results.append((r, None))
            except Exception as e:
                if isinstance(e, AttributeError):
                    e = OneNoteComError(
                        '方法 %r 无法解析 (pywin32 动态分发) —— OneNote typelib '
                        '不可加载。运行 python -m onenote_access fixregistry。' % method)
                results.append((None, e))
                last_error = e
        for r, err in results:
            if isinstance(r, tuple) and r:
                r = r[0]
            if isinstance(r, str) and r.strip():
                return r
            if err is not None:
                last_error = err
        hint = _hint_for(last_error)
        raise OneNoteComError('所有调用样式均未返回数据。最后错误: %s%s'
                              % (last_error, ('\n提示: ' + hint) if hint else ''))
