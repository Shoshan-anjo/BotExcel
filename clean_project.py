import os
import shutil
import pathlib

def clean_project():
    # Directorios y archivos a eliminar
    targets = [
        "**/__pycache__",
        "**/*.pyc",
        "build",
        "dist",
        "*.spec",
        "logs/screenshots/*",
        # "logs/*.log" # Opcional: descomenta si quieres borrar logs también
    ]

    print("🧹 Iniciando limpieza de caché y archivos temporales...")
    
    deleted_count = 0
    base_path = pathlib.Path(".")

    for target in targets:
        for path in base_path.glob(target):
            try:
                if path.is_file():
                    path.unlink()
                    print(f"  [Eliminado] Archivo: {path}")
                elif path.is_dir():
                    shutil.rmtree(path)
                    print(f"  [Eliminado] Directorio: {path}")
                deleted_count += 1
            except Exception as e:
                print(f"  [Error] No se pudo borrar {path}: {e}")

    print(f"\n✨ Limpieza completada. Se eliminaron {deleted_count} elementos.")

if __name__ == "__main__":
    clean_project()
