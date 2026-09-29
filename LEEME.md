# Transformar reportes de estudiantes

## Crear el ejecutable

En PowerShell, desde esta carpeta, ejecuta:

```powershell
.\construir_exe.ps1
```

El entorno de compilación debe tener Python 3, PyInstaller y openpyxl previamente instalados. El script no instala paquetes ni cambia la configuración del equipo. El resultado queda en `dist\TransformarEstudiantes\TransformarEstudiantes.exe`; conserva toda esa carpeta al trasladarlo.

## Usar la aplicación

1. Abre `TransformarEstudiantes.exe` y elige la carpeta que contiene `ESTUDIANTES.xlsx`.
2. Selecciona cada uno de los cinco reportes por separado, usando el botón correspondiente: **Inscritos** (`.xlsx`), **Admitidos** (`.csv`), **Primer curso** (`.xlsx`), **Matriculados** (`.csv`) y **Graduados** (`.xlsx`). Los archivos pueden estar en la carpeta `reportes` o en otra ubicación accesible; comprueba que cada campo muestre el archivo correcto antes de continuar.
3. Pulsa **Analizar** y revisa las alertas de nombres, documentos, duplicados y valores atípicos.
4. Pulsa **Generar Excel** para crear el libro transformado y revisa el archivo de salida que indique la aplicación.

`ESTUDIANTES.xlsx` se conserva como original; la aplicación genera un archivo nuevo y no sobrescribe los reportes.

El `.exe` no está firmado digitalmente. Si la política institucional impide ejecutarlo, solicite revisión o firma al equipo de TI; no cambie la configuración del antivirus. En un equipo que ya tenga Python y `openpyxl`, también se puede abrir con `python app.py`.
