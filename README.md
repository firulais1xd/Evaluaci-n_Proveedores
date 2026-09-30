# Evaluación de Proveedores (app local)

App de Streamlit que toma el archivo **EvaluacionResumenPeriodo** de la plataforma y genera:

- Resumen general por tipo de evaluación (mismo formato de la plataforma) en Excel.
- Informe de gestión en Word con una gráfica por tipo de proveedor, consolidados, calificación por área,
  resumen general, conclusiones y acciones para la mejora.
- Fichas por proveedor en Word: criterios frente al promedio del tipo, puesto, brecha para subir de categoría y evolución.
- Comparación entre periodos (semestres o años) a partir del historial guardado.

La categoría se recalcula desde la columna **Final**: ≥ 98% excelente, 80–98% muy bueno, 70–80% bueno,
60–70% regular, < 60% malo.

## Requisitos

Python 3.10 o superior (https://www.python.org/downloads/). En Windows, marca **"Add Python to PATH"** al instalar.

## Cómo abrirla

- **Windows:** doble clic en `iniciar.bat`.
- **Mac / Linux:** ejecuta `./iniciar.sh` en una terminal dentro de esta carpeta.

La primera vez tarda un par de minutos porque instala las librerías. Luego se abre el navegador en
http://localhost:8501. Para cerrarla, cierra la ventana de la consola.

Instalación manual, si prefieres:

```
python -m venv .venv
.venv\Scripts\activate          (Windows)
source .venv/bin/activate       (Mac/Linux)
pip install -r requirements.txt
streamlit run app.py
```

## Uso cada periodo

1. En la barra lateral, sube el archivo EvaluacionResumenPeriodo y escribe el nombre del periodo (p. ej. `2026-1`).
2. En **Comparar con** elige el periodo anterior (por defecto toma el guardado más reciente).
3. Revisa las pestañas: Consolidado, Por área y tipo, Ficha por proveedor, Comparativo y Revisión de datos.
4. Pulsa **Guardar este periodo en el historial** para poder compararlo más adelante.
5. En **Descargas** edita las conclusiones y descarga el Excel, el informe Word y las fichas.

## Criterios en blanco y cálculo del Final

La plataforma calcula el Final como promedio **ponderado** de los criterios y toma los criterios en blanco como 0.
Por defecto la app recalcula el Final de las filas con criterios en blanco sin tenerlos en cuenta, conservando la
ponderación de los demás (los pesos se deducen de los mismos datos). Las filas completas conservan el Final de la
plataforma. En la barra lateral, **Criterios en blanco** permite volver al Final de la plataforma.

La pestaña **Cálculo del Final** muestra las filas recalculadas y los pesos de cada tipo. Si un peso no se puede
deducir, la app supone el mismo peso para todos los criterios; puedes escribir el peso real y guardarlo
(queda en `config/ponderaciones.json`).

## Comparar contra uno o varios periodos

En la barra lateral, **Comparar con** permite elegir uno o varios periodos guardados. Las variaciones (pp,
mejoraron, bajaron) se calculan frente al periodo anterior más reciente; los demás aparecen como columnas y
series adicionales en Consolidado, Por área y tipo y Comparativo, y como tablas extra en el Excel y el Word.
En la pestaña Comparativo puedes elegir contra cuál de ellos ver los movimientos de proveedores.

## Comparar varios años (pestaña Histórico)

1. En la barra lateral abre **Cargar varios periodos al historial** y sube de una vez los archivos de los años
   anteriores (por ejemplo 2022, 2023 y 2024).
2. Revisa el nombre de cada periodo (la app lo adivina si el nombre del archivo trae el año) y pulsa
   **Guardar todos en el historial**.
3. En la pestaña **Histórico** elige los periodos a comparar. Verás la evolución del promedio general, la
   distribución por categoría, el promedio por área y por tipo, los proveedores que quedaron bajo 80% en varios
   periodos y un buscador por proveedor. El botón **Descargar histórico (.xlsx)** exporta todas esas tablas.

La ficha de cada proveedor también muestra su evolución en todos los periodos guardados, y la descarga
"Datos para Power BI" incluye todos los periodos.

## Historial

Cada periodo guardado es un archivo JSON dentro de la carpeta `historial/` (ya incluye `2025.json`).
Para respaldar el historial copia esa carpeta. Para quitar un periodo usa **Eliminar un periodo guardado**
en la barra lateral o borra su archivo.

## Nombres de tipos de proveedor y áreas

Están en `core/analisis.py`, en `TYPE_NAMES` (nombre corto por código de evaluación) y `PROC` (área según las
dos primeras letras del código: CP, CE, MN, GH, ST, GA, VT, MT, GI, SB, GC, EN). Si la plataforma agrega una
evaluación nueva, la app toma el nombre del título de la hoja; puedes agregarla a `TYPE_NAMES` para acortarlo.
