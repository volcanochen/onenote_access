# -*- coding: utf-8 -*-
"""CLI 入口: python -m onenote_access <list|export|serve|selftest>"""

import argparse
import sys

from . import __version__
from .client import OneNoteClient
from .com_backend import OneNoteComError


def cmd_list(args):
    client = OneNoteClient()
    tree = client.notebooks()
    total = 0
    nbs = [n for n in tree if n['type'] == 'Notebook']
    print('共 %d 个笔记本:' % len(nbs))
    for nb in nbs:
        print('\n笔记本: %s' % nb['name'])
        secs = [c for c in nb['children'] if c['type'] == 'Section']
        groups = [c for c in nb['children'] if c['type'] == 'SectionGroup']

        def walk_sectiongroups(g, indent):
            for child in g['children']:
                if child['type'] == 'Section':
                    n = len([p for p in child['children'] if p['type'] == 'Page'])
                    print('%s- %s  (%d 页)' % (indent, child['name'], n))
                elif child['type'] == 'SectionGroup':
                    print('%s+ %s/' % (indent, child['name']))
                    walk_sectiongroups(child, indent + '  ')

        for g in groups:
            print('  + %s/' % g['name'])
            walk_sectiongroups(g, '    ')
        for sec in secs:
            pages = [p for p in sec['children'] if p['type'] == 'Page']
            total += len(pages)
            print('  - %s  (%d 页)' % (sec['name'], len(pages)))
            for pg in pages:
                print('      * %s   [修改: %s]' % (pg['name'] or '(未命名)', pg['modified'] or '-'))
    print('\n合计 %d 页。' % total)


def cmd_export(args):
    client = OneNoteClient()
    try:
        result = client.export(
            args.out_dir, notebook=args.notebook, section=args.section,
            page=args.page, search=args.search, images=not args.no_images,
            all_pages=args.all)
    except ValueError as e:
        print('错误: %s' % e, file=sys.stderr)
        return 2
    print('导出完成: %d 页成功, %d 页失败。' % (result['exported'], result['failed']))
    print('输出目录: %s' % result['out_dir'])
    print('索引文件: %s' % result['index'])
    return 0


def cmd_serve(args):
    from . import server
    return server.main(['--host', args.host, '--port', str(args.port)] +
                       (['--token', args.token] if args.token else []) +
                       (['--export-root', args.export_root] if args.export_root else []))


def cmd_doctor(args):
    from . import doctor
    return doctor.main()


def cmd_selftest(args):
    """离线自测: 不连 OneNote, 验证解析/转换/文件管线。"""
    sys.path.insert(0, 'tests')
    from test_offline import run_all
    ok, fail = run_all(verbose=not args.quiet)
    print('\n自测结果: %d 通过, %d 失败' % (ok, fail))
    return 1 if fail else 0


def build_parser():
    p = argparse.ArgumentParser(
        prog='onenote_access',
        description='OneNote 内容读取组件 (库 / REST API / CLI), 只读, 需 Windows + OneNote 桌面版')
    p.add_argument('--version', action='version', version='onenote_access ' + __version__)
    sub = p.add_subparsers(dest='cmd', required=True)

    sp = sub.add_parser('list', help='列出笔记本/分区/页面层级')
    sp.set_defaults(func=cmd_list)

    sp = sub.add_parser('export', help='导出页面为 Markdown')
    sp.add_argument('--out-dir', required=True, help='输出目录')
    sp.add_argument('--notebook', default='*', help='笔记本名过滤, 支持 * ? 通配符')
    sp.add_argument('--section', default='*', help='分区名过滤')
    sp.add_argument('--page', default='*', help='页面标题过滤')
    sp.add_argument('--search', default=None, help='全文筛选 (内容包含该文本)')
    sp.add_argument('--no-images', action='store_true', help='不解码 base64 图片')
    sp.add_argument('--all', action='store_true', help='匹配超过 20 页时必须显式指定')
    sp.set_defaults(func=cmd_export)

    sp = sub.add_parser('serve', help='启动 REST API 服务')
    sp.add_argument('--host', default='127.0.0.1')
    sp.add_argument('--port', type=int, default=8765)
    sp.add_argument('--token', default='', help='Bearer 令牌 (或环境变量 ONENOTE_ACCESS_TOKEN)')
    sp.add_argument('--export-root', default='', help='export 接口允许的根目录')
    sp.set_defaults(func=cmd_serve)

    sp = sub.add_parser('selftest', help='离线自测 (不连 OneNote)')
    sp.add_argument('-q', '--quiet', action='store_true')
    sp.set_defaults(func=cmd_selftest)

    sp = sub.add_parser('doctor', help='诊断 COM 环境 (位数/进程/pywin32/冒烟测试)')
    sp.set_defaults(func=cmd_doctor)

    return p


def main(argv=None):
    args = build_parser().parse_args(argv)
    try:
        return args.func(args)
    except OneNoteComError as e:
        print('错误: %s' % e, file=sys.stderr)
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
