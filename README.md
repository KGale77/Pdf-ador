# Pedefeador 📝📊📈

**Pedefeador** es una aplicación de escritorio moderna y elegante diseñada para convertir archivos de Microsoft Office (Word, Excel, PowerPoint) a PDF de forma rápida y sencilla. Cuenta con una interfaz oscura interactiva basada en **CustomTkinter** y soporta conversión a través de Microsoft Office COM API o LibreOffice (fallback).

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

## 🛠️ Instalación y Configuración

1. **Clonar el repositorio:**
   ```bash
   git clone https://github.com/TU_USUARIO/pedefeador.git
   cd pedefeador
   ```

2. **Instalar las dependencias:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Ejecutar la aplicación en modo desarrollo:**
   ```bash
   python app.py
   ```

## 📦 Compilación (Crear el Ejecutable `.exe`)

Para compilar la aplicación en un único archivo ejecutable portable (`Pedefeador.exe`) que no requiere Python para funcionar, ejecuta el script de compilación provisto:

```bash
python build.py
```

El archivo ejecutable se generará en la carpeta raíz del proyecto.

## 📄 Licencia

Este proyecto está bajo la Licencia MIT.
