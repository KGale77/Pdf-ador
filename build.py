import os
import sys
import PyInstaller.__main__

def run_build():
    print("Iniciando proceso de empaquetado para Pedefeador...")
    
    # Target application file
    script_path = "app.py"
    if not os.path.exists(script_path):
        print(f"Error: No se encontró el archivo {script_path}")
        sys.exit(1)
        
    # PyInstaller options
    args = [
        script_path,
        "--onefile",                    # Single executable file
        "--noconsole",                  # Hide console window (GUI only)
        "--name=Pedefeador",            # Name of the output file
        "--collect-all=customtkinter",  # Package all customtkinter assets
        "--clean",                      # Clean cache before building
    ]
    
    print(f"Ejecutando PyInstaller con los argumentos: {args}")
    try:
        PyInstaller.__main__.run(args)
        print("\n" + "="*50)
        print("¡PROCESO FINALIZADO CON ÉXITO!")
        print("El archivo ejecutable portable ha sido creado en:")
        print(os.path.abspath("dist/Pedefeador.exe"))
        print("="*50 + "\n")
    except Exception as e:
        print(f"Ocurrió un error al compilar con PyInstaller: {e}")
        sys.exit(1)

if __name__ == "__main__":
    run_build()
