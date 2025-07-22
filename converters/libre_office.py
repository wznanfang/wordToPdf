
import os
import shutil
import subprocess

def is_available():
    return shutil.which("soffice") is not None

def convert(input_path: str, output_path: str) -> str:
    input_path = os.path.abspath(input_path)
    output_dir = os.path.dirname(output_path)

    result = subprocess.run([
        "soffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", output_dir,
        input_path
    ], stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    if result.returncode != 0:
        raise RuntimeError("LibreOffice 转换失败")

    base_name = os.path.splitext(os.path.basename(input_path))[0]
    real_output = os.path.join(output_dir, base_name + ".pdf")

    if not os.path.exists(real_output):
        raise RuntimeError("转换输出文件不存在")

    if real_output != output_path:
        os.rename(real_output, output_path)

    return output_path
