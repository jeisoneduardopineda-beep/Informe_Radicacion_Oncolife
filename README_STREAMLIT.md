# Dashboard de Radicacion en Streamlit Cloud

## Archivos principales

- `streamlit_app.py`: aplicacion para Streamlit Cloud.
- `requirements.txt`: librerias necesarias.
- `.streamlit/secrets.toml.example`: ejemplo de clave de administracion.

## Como desplegar

1. Sube estos archivos a un repositorio de GitHub.
2. Entra a Streamlit Community Cloud.
3. Crea una nueva app desde ese repositorio.
4. En `Main file path`, usa:

```text
streamlit_app.py
```

5. En `Secrets`, agrega:

```toml
ADMIN_PASSWORD = "tu-clave-segura"
```

6. Abre la app, escribe la clave en el panel lateral y carga `BASE_RADICACION.xlsx`.

## Permisos

Quien tenga el link puede:

- filtrar el dashboard,
- navegar las pestañas,
- descargar PDF,
- descargar Excel.

Solo quien tenga la clave puede cargar una nueva base.

## Nota importante

Streamlit Cloud no es una base de datos permanente. El archivo cargado puede perderse si la app se reinicia o se redepliega. Para guardar la base de forma totalmente permanente, conviene conectar almacenamiento externo como Google Drive, OneDrive, Supabase, S3 o una base de datos.
