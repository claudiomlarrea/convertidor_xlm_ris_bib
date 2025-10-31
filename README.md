
# Word Sources.xml → BibTeX / RIS

Pequeña app en Streamlit para convertir el archivo **`Sources.xml`** (del gestor de referencias de Microsoft Word) a **BibTeX (`.bib`)** o **RIS (`.ris`)**, listo para importar en **Mendeley**, **Zotero**, **EndNote**, etc.

## 🚀 Uso local

```bash
# 1) Crear entorno (opcional) e instalar dependencias
pip install -r requirements.txt

# 2) Ejecutar la app
streamlit run streamlit_app.py
```

Abrí el enlace local que aparece (por ejemplo, http://localhost:8501), subí tu `Sources.xml` y descargá el `.bib` o `.ris` generado.

## ☁️ Despliegue en Streamlit Cloud

1. Subí este repo a GitHub.
2. En https://share.streamlit.io crea una nueva app apuntando a `streamlit_app.py`.
3. Listo.

## 📁 ¿Dónde está `Sources.xml` en Word?

- **Mac**: `~/Library/Application Support/Microsoft/Office/Bibliography/Sources.xml`
- **Windows**: `%APPDATA%\Microsoft\Bibliography\Sources.xml`

Desde Word podés exportarlo/administrarlo en **Pestaña Referencias → Administrar fuentes**.

## 🔧 Limitaciones

- Soporta los tipos principales del esquema de Word (JournalArticle, Book, BookSection, ConferenceProceedings, Report, Thesis, MasterThesis, Internet, WebPage).
- Si algún campo viene vacío o con formatos no estándar, la app lo dejará en blanco.
- Sugerencias y PRs bienvenidos.

---

Hecho para investigadores.
