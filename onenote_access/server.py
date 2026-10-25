# -*- coding: utf-8 -*-
"""REST API 服务 (纯标准库 http.server, 零额外依赖; 其他系统经 HTTP 调用)。

端点 (全部 GET, 除 export 为 POST; 返回 UTF-8 JSON):
    GET  /health                                   探活 + COM 可用性
    GET  /api/v1/notebooks                         完整层级树
    GET  /api/v1/pages?notebook=&section=&page=&search=&limit=
                                                   页面元数据列表 (search 为全文筛选)
    GET  /api/v1/pages/{id}?images=true|false      单页详情 + Markdown
    POST /api/v1/export  body: {notebook,section,page,search,images,out_dir,all}
                                                   批量导出

安全:
    - 默认仅绑定 127.0.0.1; 对外暴露请配合 --token (Bearer) 与反向代理
    - export 的 out_dir 被限制在 --export-root 内 (防目录穿越)
"""

import argparse
import json
import logging
import os
import re
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from . import __version__
from .client import MAX_MATCH_DEFAULT, OneNoteClient
from .com_backend import OneNoteComError

log = logging.getLogger('onenote_access.server')

_PAGE_ID_RE = re.compile(r'^\{?[0-9A-Za-z\-]+\}?$')
_ID_PATH_RE = re.compile(r'^/api/v1/pages/([^/]+)$')


def _json_bytes(obj, status=200):
    data = json.dumps(obj, ensure_ascii=False, indent=2).encode('utf-8')
    return status, 'application/json; charset=utf-8', data


class ApiState(object):
    """服务端共享状态: 客户端 (惰性 COM)、令牌、导出根目录。"""

    def __init__(self, token=None, export_root=None, client=None):
        self.token = token
        self.export_root = os.path.abspath(export_root or os.getcwd())
        self.client = client or OneNoteClient()
        self.lock = threading.Lock()

    def check_auth(self, header_value):
        """token 未配置时跳过鉴权; 配置时要求 Bearer。"""
        if not self.token:
            return True
        return header_value == 'Bearer ' + self.token

    def resolve_export_dir(self, out_dir):
        """把请求里的 out_dir 限制在 export_root 内, 防目录穿越。"""
        if not out_dir:
            import time
            out_dir = 'onenote_export_' + time.strftime('%Y%m%d-%H%M%S')
        full = os.path.realpath(os.path.join(self.export_root, out_dir))
        root = os.path.realpath(self.export_root)
        if full != root and not full.startswith(root + os.sep):
            raise ValueError('out_dir 越界: 必须位于 export-root (%s) 之内' % root)
        return full


def _handle(state, method, path, query, body):
    """路由处理。返回 (status, content_type, bytes)。"""
    if path == '/health':
        com_ok = state.client.health()['com_available']
        return _json_bytes({'ok': True, 'version': __version__,
                            'com_available': com_ok,
                            'hint': None if com_ok else
                            'OneNote COM 不可用: 需 Windows + OneNote 桌面版已启动登录, '
                            '且 pywin32 已安装'})

    if not path.startswith('/api/v1/'):
        return _json_bytes({'error': 'not found', 'path': path}, 404)

    try:
        if path == '/api/v1/notebooks' and method == 'GET':
            return _json_bytes({'notebooks': state.client.notebooks()})

        if path == '/api/v1/pages' and method == 'GET':
            search = query.get('search', [''])[0]
            limit = int(query.get('limit', ['0'])[0] or 0)
            if search:
                pages = state.client.search_pages(
                    search,
                    notebook=query.get('notebook', ['*'])[0],
                    section=query.get('section', ['*'])[0],
                    page=query.get('page', ['*'])[0],
                    limit=limit or MAX_MATCH_DEFAULT)
            else:
                pages = state.client.find_pages(
                    notebook=query.get('notebook', ['*'])[0],
                    section=query.get('section', ['*'])[0],
                    page=query.get('page', ['*'])[0],
                    limit=limit or MAX_MATCH_DEFAULT)
            return _json_bytes({'count': len(pages), 'pages': pages})

        m = _ID_PATH_RE.match(path)
        if m and method == 'GET':
            page_id = m.group(1)
            if not _PAGE_ID_RE.match(page_id):
                return _json_bytes({'error': 'invalid page id'}, 400)
            page = state.client.get_page(
                page_id, images=query.get('images', ['true'])[0] != 'false')
            if page is None:
                return _json_bytes({'error': 'page not readable '
                                    '(不存在 / 加密分区 / 已移动)', 'page_id': page_id}, 404)
            return _json_bytes(page)

        if path == '/api/v1/export' and method == 'POST':
            payload = body or {}
            out_dir = state.resolve_export_dir(payload.get('out_dir', ''))
            result = state.client.export(
                out_dir,
                notebook=payload.get('notebook', '*'),
                section=payload.get('section', '*'),
                page=payload.get('page', '*'),
                search=payload.get('search'),
                images=payload.get('images', True),
                all_pages=bool(payload.get('all', False)))
            return _json_bytes(result)

        return _json_bytes({'error': 'not found', 'path': path}, 404)

    except OneNoteComError as e:
        return _json_bytes({'error': 'OneNote COM 不可用', 'detail': str(e)}, 503)
    except ValueError as e:
        return _json_bytes({'error': str(e)}, 400)
    except Exception as e:  # noqa: BLE001 - 边界处统一兜底
        log.exception('内部错误')
        return _json_bytes({'error': 'internal error', 'detail': str(e)}, 500)


def make_handler(state):
    class Handler(BaseHTTPRequestHandler):
        server_version = 'onenote_access/' + __version__

        def log_message(self, fmt, *args):  # 噪音抑制: 走 logging
            log.debug('%s ' + fmt, self.address_string(), *args)

        def _dispatch(self, method):
            if not state.check_auth(self.headers.get('Authorization', '')):
                body = json.dumps({'error': 'unauthorized'}, ensure_ascii=False).encode('utf-8')
                self.send_response(401)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('WWW-Authenticate', 'Bearer')
                self.send_header('Content-Length', str(len(body)))
                self.end_headers()
                self.wfile.write(body)
                return
            from urllib.parse import parse_qs, urlparse
            parsed = urlparse(self.path)
            query = parse_qs(parsed.query)
            payload = None
            if method == 'POST':
                length = int(self.headers.get('Content-Length') or 0)
                raw = self.rfile.read(length) if length else b''
                if raw:
                    try:
                        payload = json.loads(raw.decode('utf-8'))
                    except Exception:
                        status, ct, data = _json_bytes({'error': 'invalid JSON body'}, 400)
                        self._send(status, ct, data)
                        return
            status, ct, data = _handle(state, method, parsed.path, query, payload)
            self._send(status, ct, data)

        def _send(self, status, content_type, data):
            self.send_response(status)
            self.send_header('Content-Type', content_type)
            self.send_header('Content-Length', str(len(data)))
            self.end_headers()
            self.wfile.write(data)

        def do_GET(self):    # noqa: N802
            self._dispatch('GET')

        def do_POST(self):   # noqa: N802
            self._dispatch('POST')

    return Handler


def serve(host='127.0.0.1', port=8765, token=None, export_root=None):
    state = ApiState(token=token, export_root=export_root)
    httpd = ThreadingHTTPServer((host, port), make_handler(state))
    log.info('onenote_access %s listening on http://%s:%d', __version__, host, port)
    return httpd


def main(argv=None):
    import sys
    p = argparse.ArgumentParser(prog='onenote_access serve',
                                description='OneNote 读取组件 REST API 服务')
    p.add_argument('--host', default='127.0.0.1',
                   help='监听地址 (默认 127.0.0.1; 对外暴露务必配 --token)')
    p.add_argument('--port', type=int, default=8765)
    p.add_argument('--token', default=os.environ.get('ONENOTE_ACCESS_TOKEN', ''),
                   help='Bearer 令牌; 也可用环境变量 ONENOTE_ACCESS_TOKEN')
    p.add_argument('--export-root', default=os.environ.get('ONENOTE_EXPORT_ROOT', os.getcwd()),
                   help='export 接口的允许根目录 (默认当前目录)')
    p.add_argument('-v', '--verbose', action='store_true')
    args = p.parse_args(argv)
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO,
                        format='%(asctime)s %(levelname)s %(name)s: %(message)s')
    if args.host != '127.0.0.1' and not (args.token or os.environ.get('ONENOTE_ACCESS_TOKEN')):
        print('警告: 绑定非回环地址但未配置 token, 任何能访问该端口的主机都可读取你的笔记!',
              file=sys.stderr)
    httpd = serve(args.host, args.port, args.token or None, args.export_root)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        httpd.server_close()
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
