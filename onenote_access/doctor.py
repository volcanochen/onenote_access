# -*- coding: utf-8 -*-
"""doctor - OneNote COM 环境诊断 (定位"服务器运行失败"等问题)。

检查项: Python 位数 / ONENOTE.EXE 位置与位数 / OneNote 是否在运行 /
pywin32 / COM 冒烟测试。全部纯标准库实现 (PE 头读位数, tasklist 查进程)。
"""

import os
import struct
import subprocess
import sys

_OK = '[OK]  '
_BAD = '[--]  '
_INFO = '[i]   '


def python_bitness():
    return struct.calcsize('P') * 8


def find_onenote_exe():
    pf = os.environ.get('ProgramFiles') or r'C:\Program Files'
    pf86 = os.environ.get('ProgramFiles(x86)') or r'C:\Program Files (x86)'
    candidates = [
        pf + r'\Microsoft Office\root\Office16\ONENOTE.EXE',
        pf86 + r'\Microsoft Office\root\Office16\ONENOTE.EXE',
        pf + r'\Microsoft Office\Office16\ONENOTE.EXE',
        pf86 + r'\Microsoft Office\Office16\ONENOTE.EXE',
        pf + r'\Microsoft Office\root\Office15\ONENOTE.EXE',
        pf86 + r'\Microsoft Office\Office15\ONENOTE.EXE',
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def exe_bitness(path):
    """读 PE 头判断可执行文件位数: 64 / 32 / None(未知)。"""
    try:
        with open(path, 'rb') as f:
            if f.read(2) != b'MZ':
                return None
            f.seek(0x3C)
            (e_lfanew,) = struct.unpack('<I', f.read(4))
            f.seek(e_lfanew + 4)
            (machine,) = struct.unpack('<H', f.read(2))
        return {0x8664: 64, 0x14C: 32}.get(machine)
    except Exception:
        return None


def onenote_running():
    """OneNote 桌面版是否在运行: True/False; None = 无法判断。"""
    try:
        out = subprocess.run(
            ['tasklist', '/FI', 'IMAGENAME eq ONENOTE.EXE'],
            capture_output=True, text=True, timeout=15).stdout
        return 'ONENOTE.EXE' in out
    except Exception:
        return None


def pywin32_ok():
    try:
        import win32com.client  # noqa: F401
        return True
    except Exception:
        return False


def run(verbose=True):
    """执行全部诊断, 返回 (问题数, 修复建议 list)。"""
    issues = []
    advice = []

    def say(line):
        if verbose:
            print(line)

    # 1. Python
    pb = python_bitness()
    say('%sPython: %d 位 (%s)' % (_OK, pb, sys.executable))

    # 2. ONENOTE.EXE
    exe = find_onenote_exe()
    if not exe:
        say(_BAD + '未找到 OneNote 桌面版 (ONENOTE.EXE)')
        issues.append('onenote')
        advice.append('未安装 OneNote 桌面版 (UWP 版不支持 COM)。'
                      '请安装 Office 桌面版 OneNote: https://www.onenote.com/download')
    else:
        ob = exe_bitness(exe)
        say('%sOneNote 桌面版: %s (%s 位)' % (_OK, exe, ob or '?'))
        if ob and ob != pb:
            say(_BAD + '位数不一致! Python %d 位 vs OneNote %d 位' % (pb, ob))
            issues.append('bitness')
            advice.append('位数不一致是"服务器运行失败"的常见原因。'
                          '请安装与 OneNote 位数一致的 Python (当前 Python %d 位, OneNote %d 位)。'
                          '查询已有版本: py -0p' % (pb, ob))

    # 3. OneNote 进程
    running = onenote_running()
    if running is True:
        say(_OK + 'OneNote 正在运行')
    elif running is False:
        say(_BAD + 'OneNote 未在运行')
        issues.append('not_running')
        advice.append('请先手动打开 OneNote 桌面版并保持运行 (可新建空白笔记), '
                      '然后重新执行本命令。COM 首次启动服务失败 (0x80080005) '
                      '多数是 OneNote 未运行导致。')
    else:
        say(_INFO + '无法判断 OneNote 是否在运行 (tasklist 不可用)')

    # 4. pywin32
    if pywin32_ok():
        say(_OK + 'pywin32 已安装')
    else:
        say(_BAD + 'pywin32 未安装')
        issues.append('pywin32')
        advice.append('执行: pip install pywin32')

    # 5. typelib 注册 (Office C2R 缺 win64 键的已知问题)
    try:
        import pythoncom
        pythoncom.LoadRegTypeLib('{0EA692EE-BB50-4E3C-AEF0-356D91732725}', 1, 1, 0)
        say(_OK + 'OneNote typelib 可加载 (LoadRegTypeLib)')
    except Exception as e:
        say(_BAD + 'OneNote typelib 不可加载: %s' % e)
        issues.append('typelib')
        advice.append('typelib 缺少 win64 平台键 (Office C2R 已知问题)。'
                      '运行 python -m onenote_access fixregistry 修复 (弹 UAC)。')

    # 6. COM 冒烟测试
    from .com_backend import OneNoteCom, OneNoteComError
    try:
        xml = OneNoteCom().get_hierarchy()
        say(_OK + 'COM 冒烟测试通过 (GetHierarchy 返回 %d 字符)' % len(xml or ''))
        if not issues:
            advice.append('一切正常, 可直接使用 list / serve / export。')
    except OneNoteComError as e:
        say(_BAD + 'COM 冒烟测试失败: %s' % e)
        issues.append('com')

    if verbose:
        print()
        if advice:
            print('结论 / 建议:')
            for a in advice:
                print('  - ' + a)
    return len(issues), advice


def main():
    print('==== onenote_access 环境诊断 ====\n')
    n, _ = run(verbose=True)
    return 1 if n else 0


if __name__ == '__main__':
    raise SystemExit(main())
