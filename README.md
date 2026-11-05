# Log2Excel - Convertidor de Logs a Excel 🚀

**Log2Excel** es una herramienta de escritorio portable para Windows diseñada para convertir archivos de registro (`.log`, `.txt`, `.csv`) a hojas de cálculo de Microsoft Excel (`.xlsx`) de forma rápida, eficiente y 100% local.

Esta aplicación cuenta con una interfaz web moderna, responsiva y con modo oscuro (desarrollada en HTML/CSS/JS) encapsulada en una aplicación nativa de Windows utilizando Python (`pywebview`) y compilada a un archivo ejecutable independiente `.exe`.

---

## 🌟 Características

- **Diseño Moderno y Premium:** Interfaz oscura y translúcida (estilo *glassmorphism*) con micro-animaciones y soporte para arrastrar y soltar (drag & drop).
- **Procesamiento de Alto Rendimiento:** Algoritmo optimizado de un solo paso (*single-pass*) que disminuye en un **66%** el uso de memoria RAM comparado con lectores tradicionales. Soporta archivos de registro de gran tamaño sin ralentizar el sistema.
- **Configuración Dinámica:**
  - Selector de separador de celdas: Punto y coma (`;`), Coma (`,`), Tabulación (`\t`), Barra vertical (`|`) o delimitador personalizado.
  - Opción para activar/desactivar detección automática de fila de encabezados.
- **Conversión Masiva Automática (Batch):** Sube o arrastra múltiples archivos `.log`, `.txt` o `.csv` (o selecciona una carpeta completa) y la herramienta convertirá cada uno a `.xlsx` de forma directa, sin necesidad de mostrar la vista previa de la planilla, guardando automáticamente cada archivo convertido (`_convertido.xlsx`) en su **misma carpeta de origen**.
- **Previsualización Individual Opcional:** Al desmarcar la conversión directa para un único archivo, muestra las primeras 15 filas del archivo parseado antes de guardarlo para verificar que la configuración sea correcta.
- **Privacidad Total:** Todo el procesamiento y la conversión se realizan de forma local en tu computadora. Tus datos nunca se envían a ningún servidor externo.
- **100% Portable:** Un único archivo `.exe` que puedes llevar en un pendrive y usar sin necesidad de instalaciones previas de Python o Node.js.

---

## 📂 Estructura del Repositorio

El repositorio está estructurado tanto para su uso directo como para su desarrollo:

- **[`Log2Excel.exe`](./Log2Excel.exe)**: Ejecutable portable principal para Windows en la raíz del repositorio.
- **[`main.py`](./main.py)**: Backend en Python con `pywebview` y manejo nativo de archivos y carpetas.
- **[`index.html`](./index.html)**, **[`style.css`](./style.css)** y **[`app.js`](./app.js)**: Interfaz gráfica y motor de conversión individual y masiva.
- **`.gitignore`**: Configuración para evitar subir archivos basura o temporales de compilación.

---

## 🛠️ Cómo Usar

1. Descarga el archivo **`Log2Excel.exe`** directamente del repositorio.
2. Abre **`Log2Excel.exe`** en tu máquina con Windows.
3. Elige el delimitador que utilizan tus archivos log (por defecto `;`).
4. **Conversión Masiva Directa (por defecto):**
   - Arrastra uno o múltiples archivos `.log`, `.txt` o `.csv` (o una carpeta entera), o usa los botones **"Seleccionar Archivos"** / **"Seleccionar Carpeta Completa"**.
   - Cada archivo se convertirá automáticamente a Excel (`nombre_convertido.xlsx`) sin mostrar la planilla y quedará guardado en **la misma carpeta de origen** de los logs.
   - Al finalizar puedes hacer clic en **"Abrir Carpeta"** para ver todos los archivos convertidos.
5. **Modo Individual con Vista Previa (opcional):**
   - Desmarca la casilla *"Conversión directa en la misma carpeta (sin vista previa)"* y sube un único archivo para previsualizar sus primeras 15 filas antes de exportarlo.

---

## 💻 Desarrollo y Recompilación

Si deseas descargar el código fuente y modificar la aplicación en el futuro, necesitarás tener instalado **Python 3.10+**.

### 1. Requisitos
Instala las dependencias necesarias:
```powershell
pip install pywebview pyinstaller
```

### 2. Compilación a Ejecutable
Desde la carpeta raíz del proyecto original, ejecuta el siguiente comando en la consola de comandos de Windows o PowerShell para volver a compilar el código fuente a `.exe`:
```powershell
python -m PyInstaller --onefile --noconsole --add-data "index.html;." --add-data "style.css;." --add-data "app.js;." --name "Log2Excel" main.py
```
El nuevo archivo ejecutable se generará en la carpeta `dist/`.
