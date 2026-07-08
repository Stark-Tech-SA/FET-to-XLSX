# FET-to-XLSX

Aplicación de escritorio en Python para convertir archivos de FET Timetabling (`.fet`) a libros Microsoft Excel (`.xlsx`) sin requerir que FET esté instalado.

## Características

- Lectura directa del XML de FET.
- Validación de archivo existente, extensión, XML válido y estructura compatible.
- Modelos de dominio tipados para institución, docentes, materias, grupos, salones, actividades y restricciones.
- Exportación profesional con `openpyxl` a múltiples hojas: información general, docentes, materias, grupos, salones, actividades, restricciones y horarios.
- Interfaz gráfica preferente con PySide6 y fallback Tkinter: selección de archivo, resumen, progreso, exportación y mensajes de error.
- Parser tolerante a variaciones menores entre versiones: las restricciones desconocidas se exportan con todos sus campos en lugar de descartarse.

## Instalación

> Si ve `ModuleNotFoundError: No module named 'PySide6'`, significa que no instaló las dependencias. La aplicación ahora intenta abrir una interfaz alternativa con Tkinter, pero para exportar Excel necesita instalar `openpyxl`.

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Uso

```bash
python main.py
```

1. Pulse **Seleccionar archivo** y elija un `.fet`.
2. Revise el resumen del contenido.
3. Pulse **Exportar a Excel**.
4. Seleccione la ruta de salida `.xlsx`.

## Estructura del proyecto

```text
fet_to_xlsx/
├── parser/       # lectura XML, parser robusto y modelos de dominio
├── exporter/     # generación y formato del Excel
└── ui/           # interfaz gráfica PySide6
assets/           # recursos futuros
tests/            # pruebas unitarias
main.py           # punto de entrada
requirements.txt  # dependencias
```


## Horarios generados

La hoja de horarios se genera cuando el archivo incluye información de colocación. El parser reconoce dos casos:

- Salidas XML que contienen actividades ya ubicadas con día/hora/salón.
- Archivos `.fet` con actividades bloqueadas mediante `ConstraintActivityPreferredStartingTime` y salones bloqueados mediante `ConstraintActivityPreferredRoom`.

Si el `.fet` solo contiene datos de entrada y restricciones generales, pero no una solución ni actividades bloqueadas, no existe suficiente información para reconstruir un horario final; en ese caso se informa que no hay solución incluida.

Además, los horarios se exportan en hojas individuales por docente, grupo y salón.

## Decisiones de diseño

Los archivos `.fet` pueden variar entre versiones de FET. Por eso el parser:

- Usa rutas XML conocidas para entidades principales.
- Trata restricciones de tiempo y espacio como estructuras genéricas.
- Conserva campos desconocidos en la hoja **Restricciones**.
- No falla si faltan listas opcionales; exporta las hojas correspondientes vacías.

## Pruebas

```bash
python -m pytest
```
