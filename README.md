# Evaluación de Proveedores

App de Streamlit que toma el archivo **EvaluacionResumenPeriodo** de la plataforma y genera el Resumen General,
el Informe de gestión en Word, las fichas por proveedor, la comparación entre periodos y los datos para Power BI.

Elaborado por **Jersson David Navarro Cáceres**.

## Qué hace

- Lee cada hoja del archivo (un tipo de evaluación por hoja) y reconoce el tipo de proveedor y el área por el código
  y el título.
- Recalcula el Final sin contar los criterios en blanco, conservando la ponderación de la plataforma.
- Asigna la categoría: excelente ≥ 98%, muy bueno 80–98%, bueno 70–80%, regular 60–70%, malo < 60%.
- Pestañas: Consolidado, Por área y tipo, Ficha por proveedor, Comparativo, Histórico, Cálculo del Final,
  Revisión de datos y Descargas.
- Descargas: Resumen General (.xlsx), Informe de gestión (.docx), fichas (.docx), histórico (.xlsx) y datos para
  Power BI (.xlsx).

## Estructura

```
app.py                      interfaz de Streamlit
core/analisis.py            lectura del archivo, recálculo del Final, análisis y textos
core/historico.py           comparación de varios periodos
core/exportar.py            Excel, Word y datos para Power BI
requirements.txt            librerías
.streamlit/config.toml      configuración de Streamlit
.streamlit/secrets.toml.example   ejemplo de clave de acceso
historial/                  periodos guardados (no se suben al repositorio)
config/                     ponderaciones escritas a mano (no se suben)
```

## Publicar en Streamlit Community Cloud

1. **Crear el repositorio en GitHub.** En github.com → *New repository*. Se recomienda marcarlo como **Private**:
   en Streamlit Cloud una app de un repositorio privado también es privada y solo la ven las personas que invites.
2. **Subir los archivos.** En el repositorio → *Add file* → *Upload files* y arrastra el contenido de esta carpeta
   (no la carpeta en sí: `app.py` debe quedar en la raíz). Luego *Commit changes*.
   Las carpetas `.streamlit` y el archivo `.gitignore` son opcionales; si tu sistema oculta los archivos que
   empiezan por punto y no se suben, la app funciona igual.
3. **Crear la app.** Entra a [share.streamlit.io](https://share.streamlit.io) con tu cuenta de GitHub →
   *Create app* → elige el repositorio, la rama `main` y el archivo `app.py`.
4. **Configuración avanzada (opcional).** En *Advanced settings* deja Python 3.12 y, si quieres una clave de acceso,
   pega en *Secrets*:
   ```toml
   clave = "tu-clave-segura"
   ```
   Sin esa línea la app no pide clave.
5. **Deploy.** La primera vez tarda unos minutos mientras instala las librerías.
6. **Compartir.** Con la app abierta → *Share* → escribe los correos de las personas que la van a usar.

Cada vez que cambies un archivo en GitHub, la app se actualiza sola.

## Historial en la nube: descarga el respaldo

Streamlit Cloud **borra los archivos guardados cuando la app se reinicia, se actualiza o se duerme por inactividad**.
Por eso la barra lateral tiene **Respaldo del historial**:

- Después de guardar un periodo, pulsa **Descargar respaldo (.zip)** y guarda el archivo en tu equipo o en una
  carpeta compartida.
- Si al abrir la app el historial está vacío, sube ese .zip en **Restaurar desde respaldo** y pulsa **Restaurar**.

También puedes volver a cargar los archivos originales de cada año con **Cargar varios periodos al historial**.

## Datos de proveedores

El `.gitignore` evita que se suban al repositorio los periodos guardados, los Excel y los Word, porque contienen
nombres y NIT de proveedores. No subas el archivo de la plataforma ni los informes al repositorio.

## Usarla en tu equipo

```
python -m venv .venv
.venv\Scripts\activate          (Windows)
source .venv/bin/activate       (Mac/Linux)
pip install -r requirements.txt
streamlit run app.py
```

Para usar clave en tu equipo, copia `.streamlit/secrets.toml.example` como `.streamlit/secrets.toml` y cambia la clave.

## Uso en cada periodo

1. Sube el archivo EvaluacionResumenPeriodo y escribe el nombre del periodo (p. ej. `2026-1`).
2. En **Comparar con** elige uno o varios periodos guardados.
3. Revisa las pestañas y pulsa **Guardar este periodo en el historial**.
4. Descarga el respaldo del historial.
5. En **Descargas** ajusta las conclusiones y descarga el Excel, el informe y las fichas.

## Criterios en blanco

La plataforma calcula el Final como promedio ponderado de los criterios y toma los criterios en blanco como 0. La app
recalcula el Final de esas filas sin tenerlos en cuenta, conservando los pesos de los demás criterios (se deducen de
los mismos datos). En la barra lateral, **Criterios en blanco** permite volver al Final de la plataforma. La pestaña
**Cálculo del Final** muestra las filas recalculadas y los pesos de cada tipo.
