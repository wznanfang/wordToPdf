from converters import ms_word, wps_word, libre_office
from gui.app import run_app
from core.install_requirements import main as check_requirements
import sys

# 先检查依赖
check_requirements()

# 只要有一个转换器可用就算依赖检测通过
if not (ms_word.is_available() or wps_word.is_available() or libre_office.is_available()):
    print("依赖检测未通过，未检测到可用的 Word 转 PDF 转换器，程序即将退出。")
    sys.exit(1)

if __name__ == '__main__':
    run_app()

