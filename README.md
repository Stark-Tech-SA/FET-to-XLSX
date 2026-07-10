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

La hoja de horarios se genera cruzando datos por `Activity_Id`:

- `Activities_List/Activity` aporta `Id`, `Subject`, `Teacher`, `Students` y `Duration`.
- `Time_Constraints_List/ConstraintActivityPreferredStartingTime` aporta `Activity_Id`, `Day`, `Hour` y `Permanently_Locked`; ahí vive la ubicación real del bloque.
- `Space_Constraints_List/ConstraintActivityPreferredRoom` puede aportar el salón por `Activity_Id`.

El exportador agrupa por `Students` para crear hojas individuales de grupos/fichas y por `Teacher` para crear hojas individuales de docentes. Cada bloque ocupa desde la fila de `Hour` hasta `Hour + Duration - 1` y esas celdas se combinan visualmente. Las actividades sin constraint de inicio se listan en la hoja **Sin horario**, y los solapes detectados por grupo/docente se registran en **Conflictos** sin sobrescribir celdas.

Si el `.fet` solo contiene datos de entrada y restricciones generales, y tampoco existen archivos XML de resultado junto al `.fet`, no existe suficiente información para reconstruir un horario final; en ese caso se informa que no hay solución incluida.



## Base Horarios para el Blog

El exportador general agrega al final, justo antes de guardar el workbook, la función `generar_base_horarios_blog(workbook)`. Esta función no vuelve a leer el `.fet`: toma la hoja **Actividades** ya generada en el workbook abierto, elimina solo una hoja previa **Base Horarios para el Blog** si existe, y crea una nueva con las columnas exactas:

`ficha | programa | jornada | dia | hora_inicio | hora_fin | competencia | instructor | aula | fecha_inicio | fecha_fin`

Las filas se crean solo para actividades con **Día** y **Hora**. La hora de inicio se convierte a `datetime.time`, la hora final se calcula como `hora_inicio + Duración` horas, `SIN DOCENTE` se exporta vacío y las fechas quedan vacías para diligenciamiento manual.

## Hoja para base de datos CMM

La función `export_horarios_fet_sheet_from_fet()` permite actualizar un Excel existente llamado `cmm-horarios-base-datos.xlsx` sin recrearlo desde cero. Abre el libro con `openpyxl.load_workbook`, conserva las hojas existentes como **Horarios** e **Instrucciones**, elimina solo una hoja previa **Horarios_FET** si existe y la vuelve a crear con estas columnas exactas:

`ficha | programa | jornada | dia | hora_inicio | hora_fin | competencia | instructor | aula | fecha_inicio | fecha_fin`

La transformación cruza `Activity/Id` con `ConstraintActivityPreferredStartingTime/Activity_Id`, calcula `hora_fin` usando `Duration` y `Hours_List`, deja fechas vacías y omite actividades sin día/hora asignados.

Ejemplo de uso desde Python:

```python
from fet_to_xlsx.exporter.cmm_database_exporter import export_horarios_fet_sheet_from_fet

export_horarios_fet_sheet_from_fet(
    "horario.fet",
    "cmm-horarios-base-datos.xlsx",
)
```

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
