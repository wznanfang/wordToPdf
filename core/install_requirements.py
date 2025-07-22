import sys
import subprocess
import importlib.util
import platform
import shutil

# 检查 Python 包依赖
PYTHON_PACKAGES = [
    ("comtypes", "comtypes"),
]
if platform.system().lower() == "windows":
    PYTHON_PACKAGES.append(("pywin32", "win32com"))

# 注意：请在项目根目录下用 python core/install_requirements.py 运行本脚本。

def check_and_install_package(pkg_name, import_name=None):
    import_name = import_name or pkg_name
    spec = importlib.util.find_spec(import_name)
    if spec is None:
        print(f"[未安装] {pkg_name}，正在尝试安装...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", pkg_name])
            print(f"[已安装] {pkg_name}")
        except Exception as e:
            print(f"[安装失败] {pkg_name}：{e}")
            return False
    else:
        print(f"[已安装] {pkg_name}")
    return True

def check_tkinter():
    try:
        import tkinter
        print("[已安装] tkinter")
        return True
    except ImportError:
        print("[未安装] tkinter，请手动安装或检查 Python 安装包")
        return False

def check_word():
    try:
        import comtypes.client
        comtypes.client.CreateObject("Word.Application")
        print("[可用] Microsoft Word (COM)")
        return True
    except Exception:
        print("[不可用] Microsoft Word (COM)")
        return False

def check_wps():
    try:
        import win32com.client
        win32com.client.Dispatch("KWPS.Application")
        print("[可用] WPS Office (COM)")
        return True
    except Exception:
        print("[不可用] WPS Office (COM)")
        return False

def check_libreoffice():
    soffice_path = shutil.which("soffice")
    if soffice_path:
        print(f"[可用] LibreOffice (soffice)：{soffice_path}")
        return True
    else:
        print("[不可用] LibreOffice (soffice)")
        return False

def main():
    print("==== Python 包依赖检测与安装 ====")
    for pkg, import_name in PYTHON_PACKAGES:
        check_and_install_package(pkg, import_name)
    check_tkinter()

    print("\n==== 系统软件依赖检测 ====")
    sys_platform = platform.system().lower()
    if sys_platform == "windows":
        check_word()
        check_wps()
        check_libreoffice()
    else:
        check_libreoffice()

    print("\n==== 安装建议 ====")
    print("如需 Microsoft Word 支持，请确保已安装 Microsoft Office 并激活。")
    print("如需 WPS 支持，请确保已安装 WPS Office 并激活。")
    print("如需 LibreOffice 支持，请访问 https://www.libreoffice.org/download/ 下载并安装。")
    print("\n如遇权限问题，请尝试以管理员身份运行本脚本。")

if __name__ == "__main__":
    main() 


__all__ = ['main']