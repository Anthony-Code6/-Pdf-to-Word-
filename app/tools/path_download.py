import sys
from pathlib import Path

def obtener_ruta_descargas():
    sistema = sys.platform

    if sistema == 'win32':  # Windows
        import ctypes.wintypes

        CSIDL_PERSONAL = 0x0005  # Carpeta Documentos
        SHGFP_TYPE_CURRENT = 0  # Obtener el valor actual

        buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
        ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf)

        # Obtener carpeta padre (home) y agregar 'Descargas' en idioma del sistema
        documentos = Path(buf.value)
        descargas = documentos.parent / 'Downloads'
        return descargas

    elif sistema == 'darwin':  # macOS
        return Path.home() / 'Downloads'

    elif sistema.startswith('linux'):  # Linux
        try:
            from subprocess import check_output
            descargas = check_output(['xdg-user-dir', 'DOWNLOAD'], text=True).strip()
            return Path(descargas) if descargas else Path.home() / 'Downloads'
        except Exception:
            return Path.home() / 'Downloads'

    else:
        raise OSError("Sistema operativo no soportado")
