import PyInstaller.__main__
import os
import sys
import platform

# 获取当前脚本所在目录
current_dir = os.path.dirname(os.path.abspath(__file__))

# 设置图标文件路径（如果有的话）
# icon_path = os.path.join(current_dir, 'app_icon.icns')  # macOS 使用 .icns 格式

# 根据操作系统设置数据文件分隔符
separator = ':' if platform.system() == 'Darwin' else ';'

# PyInstaller参数
params = [
    'pdf_splitter_tkinter_new.py',  # 主程序文件
    '--name=PDF分割工具',  # 生成的应用名称
    '--noconsole',  # 不显示控制台窗口
    '--onefile',  # 打包成单个文件
    '--clean',  # 清理临时文件
    '--windowed',  # 不显示命令行
    # f'--icon={icon_path}',  # 设置图标（如果有的话）
    f'--add-data=README.md{separator}.',  # 添加额外文件
]

# 运行PyInstaller
PyInstaller.__main__.run(params) 