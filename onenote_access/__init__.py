# -*- coding: utf-8 -*-
"""onenote_access - 独立 OneNote 内容读取组件。

三种使用方式:
    1. 库:       from onenote_access import OneNoteClient
    2. REST API: python -m onenote_access serve --port 8765
    3. CLI:      python -m onenote_access list|export|selftest

仅支持 Windows + OneNote 桌面版 (COM 接口, 只读)。
"""

__version__ = '1.0.0'

from .client import OneNoteClient          # noqa: F401
from .com_backend import OneNoteComError   # noqa: F401

__all__ = ['OneNoteClient', 'OneNoteComError', '__version__']
