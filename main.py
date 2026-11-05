import os
import sys
import time
import base64
import tkinter as tk
from tkinter import filedialog
import webview
from webview.dom import _dnd_state

# Habilitar captura de rutas completas al arrastrar y soltar en WebView2
_dnd_state['num_listeners'] = 1

VALID_EXTENSIONS = ('.log', '.txt', '.csv')


class Api:
    def __init__(self):
        self._window = None

    def set_window(self, window):
        self._window = window

    def _create_tk_root(self):
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        return root

    def save_excel(self, base64_data, filename):
        try:
            root = self._create_tk_root()

            file_path = filedialog.asksaveasfilename(
                parent=root,
                title='Guardar archivo Excel',
                defaultextension='.xlsx',
                filetypes=[('Archivos de Excel (*.xlsx)', '*.xlsx'), ('Todos los archivos', '*.*')],
                initialfile=filename
            )

            root.destroy()

            if not file_path:
                return False

            excel_bytes = base64.b64decode(base64_data)
            with open(file_path, 'wb') as f:
                f.write(excel_bytes)

            return True
        except Exception as e:
            print(f'Error al guardar archivo: {str(e)}')
            return False

    def save_excel_in_same_folder(self, base64_data, source_path, export_filename):
        try:
            if not source_path:
                return {'success': False, 'error': 'Ruta de origen no disponible'}

            folder = os.path.dirname(os.path.abspath(source_path))
            output_path = os.path.join(folder, export_filename)

            excel_bytes = base64.b64decode(base64_data)
            with open(output_path, 'wb') as f:
                f.write(excel_bytes)

            return {
                'success': True,
                'output_path': output_path,
                'folder': folder
            }
        except Exception as e:
            print(f'Error al guardar en misma carpeta: {str(e)}')
            return {'success': False, 'error': str(e)}

    def save_excel_to_folder(self, base64_data, folder_path, export_filename):
        try:
            if not folder_path:
                return {'success': False, 'error': 'Carpeta de destino no válida'}

            output_path = os.path.join(folder_path, export_filename)
            excel_bytes = base64.b64decode(base64_data)
            with open(output_path, 'wb') as f:
                f.write(excel_bytes)

            return {
                'success': True,
                'output_path': output_path,
                'folder': folder_path
            }
        except Exception as e:
            print(f'Error al guardar en carpeta: {str(e)}')
            return {'success': False, 'error': str(e)}

    def select_files(self):
        try:
            root = self._create_tk_root()
            file_paths = filedialog.askopenfilenames(
                parent=root,
                title='Seleccionar archivos Log para convertir',
                filetypes=[
                    ('Archivos de registro (*.log, *.txt, *.csv)', '*.log *.txt *.csv'),
                    ('Todos los archivos', '*.*')
                ]
            )
            root.destroy()

            if not file_paths:
                return []

            result = []
            for p in file_paths:
                norm_path = os.path.normpath(p)
                if os.path.isfile(norm_path):
                    result.append({
                        'name': os.path.basename(norm_path),
                        'path': norm_path,
                        'size': os.path.getsize(norm_path)
                    })
            return result
        except Exception as e:
            print(f'Error al seleccionar archivos: {str(e)}')
            return []

    def select_folder(self):
        try:
            root = self._create_tk_root()
            folder_path = filedialog.askdirectory(
                parent=root,
                title='Seleccionar carpeta con archivos Log'
            )
            root.destroy()

            if not folder_path:
                return []

            norm_folder = os.path.normpath(folder_path)
            result = []
            for entry in sorted(os.listdir(norm_folder)):
                full_path = os.path.join(norm_folder, entry)
                if os.path.isfile(full_path) and entry.lower().endswith(VALID_EXTENSIONS):
                    result.append({
                        'name': entry,
                        'path': full_path,
                        'size': os.path.getsize(full_path)
                    })
            return result
        except Exception as e:
            print(f'Error al seleccionar carpeta: {str(e)}')
            return []

    def select_output_folder(self):
        try:
            root = self._create_tk_root()
            folder_path = filedialog.askdirectory(
                parent=root,
                title='Seleccionar carpeta donde guardar los archivos Excel'
            )
            root.destroy()
            return os.path.normpath(folder_path) if folder_path else ''
        except Exception as e:
            print(f'Error al seleccionar carpeta de salida: {str(e)}')
            return ''

    def clear_dropped_paths(self):
        _dnd_state['paths'].clear()
        return True

    def get_dropped_paths(self):
        try:
            # Esperar brevemente si el mensaje FilesDropped de WebView2 aún se está procesando
            for _ in range(15):
                if _dnd_state['paths']:
                    break
                time.sleep(0.02)

            raw_items = list(_dnd_state['paths'])
            _dnd_state['paths'].clear()

            result = []
            seen = set()
            for _, full_path in raw_items:
                norm_path = os.path.normpath(full_path)
                if os.path.isdir(norm_path):
                    for entry in sorted(os.listdir(norm_path)):
                        sub_path = os.path.join(norm_path, entry)
                        if os.path.isfile(sub_path) and entry.lower().endswith(VALID_EXTENSIONS):
                            if sub_path not in seen:
                                seen.add(sub_path)
                                result.append({
                                    'name': entry,
                                    'path': sub_path,
                                    'size': os.path.getsize(sub_path)
                                })
                elif os.path.isfile(norm_path):
                    if norm_path not in seen:
                        seen.add(norm_path)
                        result.append({
                            'name': os.path.basename(norm_path),
                            'path': norm_path,
                            'size': os.path.getsize(norm_path)
                        })
            return result
        except Exception as e:
            print(f'Error al obtener rutas arrastradas: {str(e)}')
            return []

    def read_file_content(self, file_path):
        try:
            for encoding in ('utf-8-sig', 'utf-8', 'cp1252', 'latin-1'):
                try:
                    with open(file_path, 'r', encoding=encoding) as f:
                        return {'success': True, 'content': f.read()}
                except UnicodeDecodeError:
                    continue
            with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
                return {'success': True, 'content': f.read()}
        except Exception as e:
            return {'success': False, 'error': str(e)}

    def open_folder(self, folder_path):
        try:
            if folder_path and os.path.exists(folder_path):
                os.startfile(folder_path)
                return True
            return False
        except Exception as e:
            print(f'Error al abrir carpeta: {str(e)}')
            return False


def get_resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath('.')

    return os.path.join(base_path, relative_path)


def main():
    api = Api()
    html_path = get_resource_path('index.html')

    window = webview.create_window(
        title='Log2Excel - Convertidor de Logs a Excel',
        url=html_path,
        js_api=api,
        width=960,
        height=780,
        resizable=True,
        min_size=(640, 480)
    )

    api.set_window(window)

    webview.start(debug=False)


if __name__ == '__main__':
    main()

