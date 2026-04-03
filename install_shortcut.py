# -*- coding: utf-8 -*-
"""
快捷方式安装脚本
在 Windows 开始菜单创建程序快捷方式
"""

import os
import subprocess
import sys

def create_shortcut():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    start_bat_path = os.path.join(script_dir, "start.bat")
    
    appdata = os.environ.get("APPDATA", "")
    start_menu_path = os.path.join(
        appdata,
        "Microsoft",
        "Windows",
        "Start Menu",
        "Programs"
    )
    
    shortcut_name = "会议纪要生成助手.lnk"
    shortcut_path = os.path.join(start_menu_path, shortcut_name)
    
    icon_path = ""
    for file in os.listdir(script_dir):
        if file.lower().endswith(".ico"):
            icon_path = os.path.join(script_dir, file)
            break
    
    print(f"项目目录: {script_dir}")
    print(f"开始菜单目录: {start_menu_path}")
    print(f"快捷方式路径: {shortcut_path}")
    print()
    
    ps_script = f'''
try {{
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut('{shortcut_path}')
    $Shortcut.TargetPath = '{start_bat_path}'
    $Shortcut.WorkingDirectory = '{script_dir}'
    $Shortcut.Description = '会议纪要生成助手'
    $Shortcut.Save()
    Write-Host "SUCCESS"
}} catch {{
    Write-Host "ERROR: $($_.Exception.Message)"
}}
'''
    
    result = subprocess.run(
        ["powershell", "-Command", ps_script],
        capture_output=True,
        text=True,
        encoding='utf-8'
    )
    
    print("PowerShell 输出:")
    print(result.stdout)
    if result.stderr:
        print("错误信息:")
        print(result.stderr)
    
    if "SUCCESS" in result.stdout:
        print("=" * 50)
        print("快捷方式创建成功！")
        print("=" * 50)
        print()
        print("您可以在 Windows 搜索栏输入 '会议纪要生成助手' 来启动程序。")
    else:
        print("=" * 50)
        print("快捷方式创建失败，请尝试以管理员身份运行此脚本")
        print("=" * 50)
    
    print()
    input("按回车键退出...")

if __name__ == "__main__":
    create_shortcut()
