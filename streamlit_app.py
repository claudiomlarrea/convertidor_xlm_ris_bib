
import io
import pandas as pd
import streamlit as st
from converters import parse_sources_xml, to_bibtex, to_ris, to_rows

st.set_page_config(page_title="Word Sources.xml → BibTeX / RIS", page_icon="📚", layout="wide")
st.title("📚 Convertidor: Microsoft Word Sources.xml → BibTeX (.bib) / RIS (.ris)")
st.caption("Subí tu `Sources.xml` (del gestor de referencias de Word) y convertí a archivos compatibles con Mendeley, Zotero, EndNote etc.")

with st.sidebar:
    st.header("⚙️ Opciones")
    out_fmt = st.radio("Formato de salida", ["BibTeX (.bib)", "RIS (.ris)"], index=0)
    show_preview = st.checkbox("Mostrar vista previa del archivo generado", value=True)

uploaded = st.file_uploader("Arrastrá aquí tu `Sources.xml` o hacé clic para seleccionarlo", type=["xml"])

if uploaded is not None:
    xml_bytes = uploaded.read()
    try:
        sources = parse_sources_xml(xml_bytes)
    except Exception as e:
        st.error(f"No pude leer el XML. ¿Es un archivo `Sources.xml` de Word? Detalle: {e}")
        st.stop()

    st.success(f"Referencias detectadas: **{len(sources)}**")
    if len(sources) == 0:
        st.info("No se encontraron referencias. Verificá que el archivo sea el `Sources.xml` exportado por Word (Administrar fuentes).")
    else:
        rows = to_rows(sources)
        df = pd.DataFrame(rows)
        st.dataframe(df, use_container_width=True, hide_index=True)

        if out_fmt.startswith("BibTeX"):
            text = to_bibtex(sources)
            file_ext = "bib"
        else:
            text = to_ris(sources)
            file_ext = "ris"

        col1, col2 = st.columns([2,1])
        with col1:
            if show_preview:
                st.subheader("Vista previa")
                st.code(text[:20000], language="text")
        with col2:
            st.download_button(
                label=f"⬇️ Descargar {file_ext.upper()}",
                data=text.encode("utf-8"),
                file_name=f"sources_converted.{file_ext}",
                mime="text/plain",
            )

st.markdown("---")
st.markdown("Hecho para investigadores: convierte tus fuentes de Word al instante.")

