from cx_Freeze import setup, Executable
import os

# 收集静态文件并保留目录结构
data_files = []
resources_dir = ''  # 静态资源目录
for root, dirs, files in os.walk(resources_dir):
    for file in files:
        source_path = os.path.join(root, file)
        # 修正目标路径：保留完整路径（含文件名）
        relative_path = os.path.relpath(source_path, resources_dir)  # 获取相对路径（如 css/style.css）
        destination_path = os.path.join('resources', relative_path)  # 目标路径为 resources/css/style.css
        data_files.append((source_path, destination_path))

executables = [
    Executable("gui.py", base='Win32GUI', target_name="CatchExcelImage.exe")  # GUI 应用需指定 base
]

# 设置默认安装目录
build_exe_options = {
    "zip_include_packages": [],
    "include_files": data_files
}

bdist_msi_options = {
    "upgrade_code": "{0a4e2604-67f3-4a21-9076-dbed5699d419}",  # 替换为你生成的 GUID
    "initial_target_dir": r"[ProgramFilesFolder]\CatchExcelImage"  # 默认安装目录
}

setup(
    name="CatchExcelImage",
    version="1.0",
    description="捕获Excel图片",
    options={
        'build_exe': build_exe_options,
        'bdist_msi': bdist_msi_options
    },
    executables=executables,
)