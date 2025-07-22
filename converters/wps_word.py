
import os

def is_available():
    try:
        import win32com.client
        win32com.client.Dispatch("KWPS.Application")
        return True
    except Exception:
        return False

def convert(input_path: str, output_path: str) -> str:
    import win32com.client
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)

    wps = win32com.client.Dispatch("KWPS.Application")
    wps.Visible = False

    try:
        doc = wps.Documents.Open(input_path)
        doc.ExportAsFixedFormat(output_path, 17)
        doc.Close()
    finally:
        wps.Quit()

    if not os.path.exists(output_path):
        raise RuntimeError("WPS 转换失败")
    return output_path
