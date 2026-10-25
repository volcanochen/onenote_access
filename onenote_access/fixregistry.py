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

# (typelib GUID, 版本, exe 内嵌 typelib 资源索引)
TYPELIBS = [
    ('{0EA692EE-BB50-4E3C-AEF0-356D91732725}', '1.1', 3),
    ('{F2A7EE29-8BF6-4a6d-83F1-098E366C709C}', '1.0', 2),
]

C2R_BASE = (r'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE'
            r'\Software\Classes\TypeLib')
HIVE_BASES = [
    r'HKLM:\SOFTWARE\Classes\TypeLib',   # 真实 hive
    C2R_BASE,                            # C2R 虚拟 hive (Office 进程视角)
]

MARKER_DEFAULT = os.path.join(os.environ.get('TEMP', r'C:\Temp'),
                              'onenote_fixregistry_result.txt')


def _onenote_exe():
    """定位 ONENOTE.EXE (含 ProgramFiles 环境变量为空的兜底)。"""
    pf = os.environ.get('ProgramFiles') or r'C:\Program Files'
    pf86 = os.environ.get('ProgramFiles(x86)') or r'C:\Program Files (x86)'
    for base in (pf, pf86):
        for rel in (r'Microsoft Office\Root\Office16\ONENOTE.EXE',
                    r'Microsoft Office\Office16\ONENOTE.EXE',
                    r'Microsoft Office\Root\Office15\ONENOTE.EXE'):
            p = os.path.join(base, rel)
            if os.path.exists(p):
                return p
    return None


def _win64_states():
    """检查每个 typelib 在两个 HKLM 视图的 win64 键。

    返回 {(guid, ver): (hklm_ok, c2r_ok, exe_path)}。
    注意: HKCU 的 win64 键对客户端进程有效, 但 OneNote 进程 (C2R 容器内)
    不一定能看到, 所以不能算修复完成 —— 必须两个 HKLM 视图都有。
    """
    import winreg
    exe = _onenote_exe()
    states = {}
    for guid, ver, _res in TYPELIBS:
        def exists(path):
            try:
                winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path)
                return True
            except OSError:
                return False
        hklm_ok = exists(r'SOFTWARE\Classes\TypeLib\%s\%s\0\win64' % (guid, ver))
        c2r_ok = exists(r'SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE'
                        r'\Software\Classes\TypeLib\%s\%s\0\win64' % (guid, ver))
        states[(guid, ver)] = (hklm_ok, c2r_ok, exe)
    return states


def ensure_hkcu_keys():
    """把 win64 typelib 键 + IApplication Interface 键写入 HKCU (免提权)。

    HKCU 键对本进程的 COM 代理构建 (LoadRegTypeLib / universal marshaler)
    即可生效; OneNote 服务器进程不一定可见, 那需要 HKLM 的完整修复。
    comtypes vtable 后端在 connect 失败时调用此函数自愈后重试。
    """
    import winreg
    exe = _onenote_exe()
    if not exe:
        raise OneNoteComErrorUnavailable()
    written = []
    for guid, ver, res in TYPELIBS:
        path = r'SOFTWARE\Classes\TypeLib\%s\%s\0\win64' % (guid, ver)
        k = winreg.CreateKey(winreg.HKEY_CURRENT_USER, path)
        winreg.SetValueEx(k, None, 0, winreg.REG_SZ, '%s\\%d' % (exe, res))
        winreg.CloseKey(k)
        written.append(path)
    iface = r'SOFTWARE\Classes\Interface\{452AC71A-B655-4967-A208-A4CC39DD7949}'
    k = winreg.CreateKey(winreg.HKEY_CURRENT_USER, iface)
    winreg.SetValueEx(k, None, 0, winreg.REG_SZ, 'IApplication')
    winreg.SetValueEx(k, 'ProxyStubClsid32', 0, winreg.REG_SZ,
                      '{00020424-0000-0000-C000-000000000046}')
    winreg.SetValueEx(k, 'TypeLib', 0, winreg.REG_SZ,
                      '{0EA692EE-BB50-4E3C-AEF0-356D91732725}')
    winreg.SetValueEx(k, 'Version', 0, winreg.REG_SZ, '1.1')
    winreg.CloseKey(k)
    written.append(iface)
    return written


class OneNoteComErrorUnavailable(RuntimeError):
    """ONENOTE.EXE 未找到, 无法确定 typelib 资源路径。"""


def missing_win64_keys():
    """返回缺失 win64 平台键的 (GUID, 版本) 列表 (只认 HKLM 两个视图)。"""
    states = _win64_states()
    if not states:
        return []
    out = []
    for (guid, ver), (hklm_ok, c2r_ok, _exe) in states.items():
        if not (hklm_ok and c2r_ok):
            out.append((guid, ver))
    return out


def build_ps_script(marker_path, exe_path):
    """生成提权后执行的 PowerShell 脚本内容 (纯 ASCII 路径, 无编码风险)。"""
    lines = [
        '$ErrorActionPreference = "Continue"',
        '$results = @()',
    ]
    for guid, ver, res in TYPELIBS:
        path = '%s\\%d' % (exe_path, res)
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


def build_deep_repair_script(marker_path):
    """深度修复: 重启 C2R 服务 (刷新注册表过滤视图) + 重启 OneNote。纯 ASCII。"""
    lines = [
        '$results = @()',
        'try {',
        '  Restart-Service ClickToRunSvc -Force -ErrorAction Stop',
        '  $results += "OK: ClickToRunSvc restarted, status=" + '
        '(Get-Service ClickToRunSvc).Status',
        '} catch { $results += "FAIL: " + $_.Exception.Message }',
        'Stop-Process -Name ONENOTE -Force -ErrorAction SilentlyContinue',
        'Stop-Process -Name ONENOTEM -Force -ErrorAction SilentlyContinue',
        '$results += "OK: OneNote restarted"',
        '$results | Out-File -FilePath "%s" -Encoding utf8' % marker_path,
    ]
    return '\r\n'.join(lines)


def run_elevated_script(body, marker_path, timeout=180):
    """把 PowerShell 脚本体写成临时 ps1 并经 UAC 提权执行。返回 (ok, 输出行)。"""
    if os.path.exists(marker_path):
        os.remove(marker_path)
    ps1 = marker_path.replace('.txt', '.ps1')
    with open(ps1, 'w', encoding='ascii') as f:
        f.write(body)

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


def run_elevated(marker_path=MARKER_DEFAULT, timeout=180):
    """补写 win64 键 (提权)。返回 (ok, 输出行)。"""
    exe = _onenote_exe()
    if not exe:
        return False, ['未找到 ONENOTE.EXE, 无法确定 typelib 资源路径']
    return run_elevated_script(build_ps_script(marker_path, exe), marker_path, timeout)


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
    exe = _onenote_exe()
    if not exe:
        print('[--] 未找到 ONENOTE.EXE (OneNote 桌面版未安装?)')
        return 1
    print('[i]  OneNote: %s' % exe)
    states = _win64_states()
    need_fix = False
    for (guid, ver), (hklm_ok, c2r_ok, _e) in states.items():
        mark = lambda ok: '[OK]  ' if ok else '[--]  '
        print('%sHKLM 真实 hive  TypeLib\\%s\\%s\\0\\win64' % (mark(hklm_ok), guid, ver))
        print('%sC2R 虚拟 hive   TypeLib\\%s\\%s\\0\\win64' % (mark(c2r_ok), guid, ver))
        if not (hklm_ok and c2r_ok):
            need_fix = True
    if not need_fix:
        print('\n[OK] 两个 HKLM 视图的 win64 键都已存在。')
    else:
        print('\n即将以管理员权限补写 HKLM 注册表键并重启 OneNote (弹 UAC, 请点"是")...\n')
        ok, lines = run_elevated()
        for line in lines:
            print('  ' + line)
        if not ok:
            return 1
        print('\n等待 OneNote 重新启动...')
        time.sleep(5)

    ok, msg = verify()
    print('[%s] %s' % ('OK' if ok else '--', msg))
    if ok:
        print('\n修复成功! 现在可以正常运行: python -m onenote_access list')
        return 0

    # 键在但调用仍失败 -> 深度修复: 重启 C2R 服务刷新注册表过滤视图
    print('\n尝试深度修复: 重启 ClickToRunSvc (刷新 C2R 注册表过滤视图) '
          '+ 重启 OneNote...')
    ok2, lines = run_elevated_script(
        build_deep_repair_script(MARKER_DEFAULT), MARKER_DEFAULT)
    for line in lines:
        print('  ' + line)
    if not ok2:
        return 1
    print('\n等待 OneNote 就绪...')
    time.sleep(5)
    ok3, msg3 = verify()
    print('[%s] %s' % ('OK' if ok3 else '--', msg3))
    if ok3:
        print('\n深度修复成功! 现在可以正常运行: python -m onenote_access list')
        return 0
    print('\n深度修复后仍失败。建议执行 Office 快速修复 (重建全部 COM 注册):')
    print('  设置 -> 应用 -> 安装的应用 -> Microsoft Office/365 -> 修改 -> 快速修复')
    print('修复完成后重跑 python -m onenote_access doctor')
    return 1


if __name__ == '__main__':
    raise SystemExit(main())
