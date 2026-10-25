# -*- coding: utf-8 -*-
"""fixregistry - 修复 Office C2R 的 OneNote typelib 注册缺失 win64 平台键。

症状: COM 调用报 TYPE_E_LIBNOTREGISTERED (-2147319779, "库没有注册"),
GetTypeInfo 失败。根因: Office Click-to-Run 只把 OneNote typelib 注册到
win32 平台键, 64 位进程 (含 64 位 OneNote 自己) 无法加载 typelib 进行
IDispatch 封送/分发。

修复: 在 HKLM 真实 hive 与 C2R 虚拟 hive 补 win64 键 (需要管理员, 自动
弹出 UAC), 然后重启 OneNote 使其重新加载。HKCU 键对本进程免费可用,
但对 OneNote 进程不一定可见, 所以必须写 HKLM。

用法: python -m onenote_access fixregistry
"""

import ctypes
import os
import sys
import time

# (typelib GUID, 版本, 内嵌资源路径后缀)
TYPELIBS = [
    ('{0EA692EE-BB50-4E3C-AEF0-356D91732725}', '1.1',
     r'C:\Program Files\Microsoft Office\Root\Office16\ONENOTE.EXE\3'),
    ('{F2A7EE29-8BF6-4a6d-83F1-098E366C709C}', '1.0',
     r'C:\Program Files\Microsoft Office\Root\Office16\ONENOTE.EXE\2'),
]

C2R_BASE = (r'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE'
            r'\Software\Classes\TypeLib')
HIVE_BASES = [
    r'HKLM:\SOFTWARE\Classes\TypeLib',   # 真实 hive
    C2R_BASE,                            # C2R 虚拟 hive (Office 进程视角)
]

MARKER_DEFAULT = os.path.join(os.environ.get('TEMP', r'C:\Temp'),
                              'onenote_fixregistry_result.txt')


def missing_win64_keys():
    """返回缺失 win64 平台键的 (GUID, 版本) 列表 (查 HKCR 合并视图)。"""
    import winreg
    missing = []
    for guid, ver, _ in TYPELIBS:
        found = False
        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                k = winreg.OpenKey(hive, r'SOFTWARE\Classes\TypeLib\%s\%s\0' % (guid, ver))
                subs = []
                i = 0
                while True:
                    try:
                        subs.append(winreg.EnumKey(k, i).lower())
                        i += 1
                    except OSError:
                        break
                if 'win64' in subs:
                    found = True
                    break
            except OSError:
                continue
        if not found:
            missing.append((guid, ver))
    return missing


def build_ps_script(marker_path):
    """生成提权后执行的 PowerShell 脚本内容 (纯 ASCII 路径, 无编码风险)。"""
    lines = [
        '$ErrorActionPreference = "Continue"',
        '$results = @()',
    ]
    for guid, ver, path in TYPELIBS:
        for base in HIVE_BASES:
            key = r'%s\%s\%s\0\win64' % (base, guid, ver)
            lines.append('try {')
            lines.append('  New-Item -Path "%s" -Force | Out-Null' % key)
            lines.append('  Set-ItemProperty -Path "%s" -Name "(default)" '
                         '-Value "%s" -ErrorAction Stop' % (key, path))
            lines.append('  $results += "OK: %s"' % key)
            lines.append('} catch { $results += "FAIL: %s -> " + '
                         '$_.Exception.Message }' % key)
    # 重启 OneNote 使其重新加载 typelib
    lines.append('$results += "Restarting OneNote..."')
    lines.append('Stop-Process -Name ONENOTE -Force -ErrorAction SilentlyContinue')
    lines.append('Stop-Process -Name ONENOTEM -Force -ErrorAction SilentlyContinue')
    lines.append('$results | Out-File -FilePath "%s" -Encoding utf8' % marker_path)
    return '\r\n'.join(lines)


def run_elevated(marker_path=MARKER_DEFAULT, timeout=120):
    """写临时 ps1 并经 UAC 提权执行; 轮询结果标记文件。返回 (ok, 输出行)。"""
    if os.path.exists(marker_path):
        os.remove(marker_path)
    ps1 = marker_path.replace('.txt', '.ps1')
    with open(ps1, 'w', encoding='ascii') as f:
        f.write(build_ps_script(marker_path))

    ret = ctypes.windll.shell32.ShellExecuteW(
        None, 'runas', 'powershell.exe',
        '-NoProfile -ExecutionPolicy Bypass -File "%s"' % ps1, None, 1)
    if ret <= 32:  # 5 = 拒绝, 1223 = 取消
        return False, ['用户拒绝了 UAC 提权请求 (ShellExecute 返回 %d)' % ret]

    deadline = time.time() + timeout
    while time.time() < deadline:
        if os.path.exists(marker_path):
            time.sleep(0.5)
            with open(marker_path, encoding='utf-8-sig') as f:
                return True, [line.rstrip() for line in f if line.strip()]
        time.sleep(1)
    return False, ['超时: 未等到提权脚本的结果文件 (%s)' % marker_path]


def verify():
    """修复验证: LoadRegTypeLib + COM 冒烟。返回 (ok, 消息)。"""
    try:
        import pythoncom
        pythoncom.LoadRegTypeLib('{0EA692EE-BB50-4E3C-AEF0-356D91732725}', 1, 1, 0)
    except Exception as e:
        return False, 'LoadRegTypeLib 仍失败: %s' % e
    try:
        from .com_backend import OneNoteCom
        xml = OneNoteCom().get_hierarchy()
        return True, 'COM 冒烟测试通过 (层级 %d 字符)' % len(xml or '')
    except Exception as e:
        return False, 'typelib 已可加载, 但 COM 调用仍失败: %s' % e


def main():
    print('==== OneNote typelib 注册修复 ====\n')
    missing = missing_win64_keys()
    if not missing:
        print('[OK] win64 平台键已存在, 无需修复。')
        return 0
    for guid, ver in missing:
        print('[--] 缺失: TypeLib\\%s\\%s\\0\\win64' % (guid, ver))
    print('\n即将以管理员权限补写 HKLM 注册表键并重启 OneNote (弹 UAC, 请点"是")...\n')
    ok, lines = run_elevated()
    for line in lines:
        print('  ' + line)
    if not ok:
        return 1
    # 给 OneNote 自动重启留一点时间
    print('\n等待 OneNote 重新启动...')
    time.sleep(5)
    ok, msg = verify()
    print('[%s] %s' % ('OK' if ok else '--', msg))
    if ok:
        print('\n修复成功! 现在可以正常运行: python -m onenote_access list')
        return 0
    print('\n如仍失败, 尝试手动打开 OneNote 后重跑 python -m onenote_access doctor')
    return 1


if __name__ == '__main__':
    raise SystemExit(main())
