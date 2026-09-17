import os
import sys
import shutil
import PyInstaller.__main__
from PyInstaller.utils.hooks import collect_data_files, collect_submodules, copy_metadata

# 1. 自动生成增强版引导脚本 (run_app.py)，加入防崩元数据 Patch
bootstrap_code = """import os
import sys
import importlib.metadata

# --- 防闪退补丁：拦截 PackageNotFoundError，保证 version("streamlit") 永不崩溃 ---
_orig_version = importlib.metadata.version
def _safe_version(pkg_name):
    try:
        return _orig_version(pkg_name)
    except importlib.metadata.PackageNotFoundError:
        if pkg_name.lower() == "streamlit":
            return "1.35.0"
        return "0.0.0"
importlib.metadata.version = _safe_version
# --------------------------------------------------------------------------

import streamlit.web.cli as stcli

if __name__ == "__main__":
    if getattr(sys, 'frozen', False):
        base_dir = sys._MEIPASS
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        
    app_script = os.path.join(base_dir, "app.py")
    
    # 配置生产环境参数：自动打开本地网页、保留黑色控制台、关闭匿名遥测提问
    sys.argv = [
        "streamlit",
        "run",
        app_script,
        "--global.developmentMode=false",
        "--server.headless=false",
        "--server.port=8501",
        "--browser.serverAddress=localhost",
        "--browser.gatherUsageStats=false",
        "--server.enableCORS=false",
        "--server.enableXsrfProtection=false",
    ]
    sys.exit(stcli.main())
"""

with open("run_app.py", "w", encoding="utf-8") as f:
    f.write(bootstrap_code)

print(">>> [1/3] 已生成 Streamlit 启动引导脚本 run_app.py")

# 2. 收集静态文件与 pip 元数据 (核心修复点：copy_metadata)
datas = [
    ("app.py", "."),
]

# 收集静态资源
datas += collect_data_files("streamlit")
datas += collect_data_files("pdf2docx")
datas += collect_data_files("reportlab")
datas += collect_data_files("cryptography")

# 核心修复：收集 streamlit 的 package metadata，解决 PackageNotFoundError
datas += copy_metadata("streamlit")

hidden_imports = [
    "streamlit",
    "streamlit.web.cli",
    "streamlit.runtime.scriptrunner.magic_funcs",
    "pymupdf",
    "pypdf",
    "reportlab",
    "reportlab.pdfgen.canvas",
    "reportlab.platypus",
    "pdf2docx",
    "pdfplumber",
    "pptx",
    "docx",
    "openpyxl",
    "pandas",
    "PIL",
    "cryptography",
    "cryptography.hazmat.primitives.kdf.pbkdf2",
    "cryptography.hazmat.primitives.serialization.pkcs7",
]
hidden_imports += collect_submodules("streamlit")
hidden_imports += collect_submodules("pdf2docx")

# 3. 组装 PyInstaller 参数
icon_param = ["--icon=app.ico"] if os.path.exists("app.ico") else []

pyinstaller_args = [
    "run_app.py",
    "--name=多功能PDF文档智能编辑与处理软件",
    "--onedir",             # 目录模式，启动迅速且稳定
    "--console",            # 保留 CMD 窗口以查看状态
    "--copy-metadata=streamlit",  # 强行注入元数据
    "--noconfirm",
    "--clean",
] + icon_param

# 拼接数据与隐式导入参数
for src, dst in datas:
    pyinstaller_args.extend(["--add-data", f"{src}{os.pathsep}{dst}"])

for imp in set(hidden_imports):
    pyinstaller_args.extend(["--hidden-import", imp])

print(">>> [2/3] 开始打包编译，请稍候...")
PyInstaller.__main__.run(pyinstaller_args)

# 4. 清理临时生成的入口文件
if os.path.exists("run_app.py"):
    os.remove("run_app.py")

print(">>> [3/3] 打包完成！请进入 dist/通用多维PDF智能处理软件 目录运行程序测试。")