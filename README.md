# Pdf-ador 📝📊📈

**Pdf-ador** es una aplicación de escritorio moderna y elegante diseñada para convertir archivos de Microsoft Office (Word, Excel, PowerPoint) a PDF de forma rápida y sencilla. Cuenta con una interfaz oscura interactiva basada en **CustomTkinter** y soporta conversión a través de Microsoft Office COM API o LibreOffice (fallback).

## ✨ Características

- **Conversión por lotes:** Arrastra o selecciona múltiples archivos y conviértelos todos a la vez.
- **Soporte multiformato:** Soporta extensiones de Word (`.docx`, `.doc`, `.rtf`), Excel (`.xlsx`, `.xls`) y PowerPoint (`.pptx`, `.ppt`).
- **Doble motor de conversión:**
  - Conversión nativa ultra-rápida y fiel a través de las APIs COM de Microsoft Office (requiere tener instalado Office).
  - Conversión de respaldo usando LibreOffice (si está instalado).
- **Interfaz moderna:** Estética oscura premium con animaciones suaves, barras de progreso e indicadores de estado detallados.
- **Hilos independientes:** La UI nunca se congela durante las conversiones.

## 🚀 Requisitos

- Windows OS
- Python 3.10+
- (Opcional, recomendado) Microsoft Office instalado en el sistema para conversión nativa, o LibreOffice como fallback.

## 🚀 Uso Rápido (Ejecutable)

Si no quieres instalar Python, puedes ejecutar la aplicación directamente utilizando el archivo ejecutable portable precompilado **`Pdf-ador.exe`** que se encuentra en la raíz del repositorio.

## 🛠️ Desarrollo (Instalación y Configuración)

Si deseas modificar el código o correr la aplicación en modo desarrollo:

1. **Clonar el repositorio:**
   ```bash
   git clone https://github.com/KGale77/Pdf-ador.git
   cd pedefeador
   ```

2. **Instalar las dependencias:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Ejecutar la aplicación:**
   ```bash
   python app.py
   ```

## 📄 Licencia

Este proyecto está bajo la Licencia MIT.
