from converters import ms_word, wps_word, libre_office

def convert_word_to_pdf(input_path: str, output_path: str) -> str:
    converters = [ms_word, wps_word, libre_office]
    for converter in converters:
        if converter.is_available():
            try:
                return converter.convert(input_path, output_path)
            except Exception as e:
                raise RuntimeError(f"使用 {converter.__name__} 转换失败：{str(e)}")
    raise RuntimeError("未检测到支持的 Office 程序（Word/WPS/LibreOffice）") 