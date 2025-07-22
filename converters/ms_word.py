
import os

def is_available():
    try:
        import comtypes.client
        comtypes.client.CreateObject("Word.Application")
        return True
    except Exception:
        return False

def convert(input_path: str, output_path: str) -> str:
    import comtypes.client
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)

    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False

    try:
        doc = word.Documents.Open(input_path)
        doc.ExportAsFixedFormat(output_path, ExportFormat=17)
        doc.Close()
    finally:
        word.Quit()

    if not os.path.exists(output_path):
        raise RuntimeError("Word 转换失败")
    return output_path
