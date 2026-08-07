# Orden de Producción → Excel

App para leer el PDF de orden de producción (C&A / Conceptos Nancy) y llenar
automáticamente la plantilla de packs/tallas `ETIQUETAC_APACKCH.xlsx`.

## Probarla en tu compu (opcional)

1. Instala Python si no lo tienes.
2. Abre una terminal en esta carpeta y corre:
   ```
   pip install -r requirements.txt
   streamlit run app.py
   ```
3. Se abre solo en tu navegador, normalmente en http://localhost:8501

## Subirla a internet (mismo proceso que ya hiciste con la de Walmart)

1. Crea un repositorio nuevo en GitHub (puede ser privado, o público si
   quieres que Streamlit Cloud lo detecte sin problema como pasó la vez pasada).
2. Sube estos 4 archivos: `app.py`, `requirements.txt`, `LEEME.md`,
   **y también `ETIQUETAC_APACKCH.xlsx`** (la plantilla — sin este archivo
   la app no funciona, es la que trae el formato base).
3. Entra a https://share.streamlit.io con tu cuenta de GitHub.
4. "Create app" → "Deploy a public app from GitHub".
5. Repository: tu nuevo repo · Branch: `main` · Main file path: `app.py`.
6. Deploy. En un par de minutos te da tu URL, igual que la de etiquetas.

## Notas

- La plantilla `ETIQUETAC_APACKCH.xlsx` vive junto al código — si algún día
  cambias el formato de la plantilla (colores, encabezados, columnas),
  reemplaza ese archivo en el repo y la app usa el nuevo automáticamente.
- La app valida que la suma de piezas de los packs cuadre contra
  "Piezas Totales" del PDF antes de dejarte descargar — si no cuadra,
  te avisa el error en pantalla en vez de generar un Excel incorrecto.
- No guarda nada en base de datos: cada PDF se procesa en memoria y se
  descarta al cerrar la pestaña.
