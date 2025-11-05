import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
import os
import tempfile
import subprocess
import shutil

# --- Configuración de la página ---
st.set_page_config(page_title="Generador de Certificados", page_icon="🎓", layout="centered")

st.title("🎓 Generador de Certificados")
st.write("Subí tus archivos de asistencia y el template PowerPoint para generar certificados automáticamente en formato PPTX y PDF.")

# --- Subida de archivos ---
uploaded_template = st.file_uploader("📄 Subí el template (.pptx)", type=["pptx"])
uploaded_excel1 = st.file_uploader("📘 Subí listado presencial (.xlsx)")
uploaded_excel2 = st.file_uploader("📗 Subí listado virtual (.xlsx)")

# --- Función para convertir PPTX → PDF usando LibreOffice ---
def convert_to_pdf(input_pptx, output_dir):
    """Convierte un archivo .pptx a .pdf usando LibreOffice (modo headless)."""
    try:
        subprocess.run([
            "libreoffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", output_dir,
            input_pptx
        ], check=True)
    except Exception as e:
        print(f"Error al convertir {input_pptx}: {e}")

# --- Lógica principal ---
if uploaded_template and uploaded_excel1 and uploaded_excel2:
    if st.button("🚀 Generar certificados"):
        with st.spinner("Generando certificados..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                # Guardar archivos subidos en el directorio temporal
                template_path = os.path.join(tmpdir, "template.pptx")
                with open(template_path, "wb") as f:
                    f.write(uploaded_template.read())

                # Leer los Excel
                df1 = pd.read_excel(uploaded_excel1)
                df2 = pd.read_excel(uploaded_excel2)

                # Filtrar asistentes presenciales
                df1_asistentes = df1[
                    (df1["Cargo"].str.lower() == "asistente")
                    & (df1["Asistió"].str.upper() == "SI")
                ].copy()

                # Crear columna unificada de nombres
                df1_asistentes["Nombre y apellido"] = (
                    df1_asistentes["Apellido"].str.strip() + " " + df1_asistentes["Nombre"].str.strip()
                ).str.title()

                df2["Nombre y apellido"] = (
                    df2["Apellido"].str.strip() + " " + df2["Nombre"].str.strip()
                ).str.title()

                # Combinar ambos listados
                df_final = pd.concat(
                    [df1_asistentes[["Nombre y apellido"]], df2[["Nombre y apellido"]]],
                    ignore_index=True,
                ).drop_duplicates().reset_index(drop=True)

                # Crear carpeta de salida
                output_dir = os.path.join(tmpdir, "Certificados")
                os.makedirs(output_dir, exist_ok=True)

                # Generar certificados
                for nombre in df_final["Nombre y apellido"]:
                    prs = Presentation(template_path)
                    for slide in prs.slides:
                        for shape in slide.shapes:
                            if shape.has_text_frame:
                                for paragraph in shape.text_frame.paragraphs:
                                    for run in paragraph.runs:
                                        if "Nombre y apellido" in run.text:
                                            run.text = run.text.replace("Nombre y apellido", nombre)
                                            run.font.name = "Quintessential"
                                            run.font.size = Pt(20)
                                            run.font.italic = True
                                            run.font.color.rgb = RGBColor(0, 176, 240)
                                    paragraph.alignment = PP_ALIGN.CENTER

                    # Guardar PPTX
                    output_pptx = os.path.join(output_dir, f"Certificado_{nombre}.pptx")
                    prs.save(output_pptx)

                    # Convertir a PDF
                    convert_to_pdf(output_pptx, output_dir)

                # Crear archivo ZIP con todo
                zip_path = os.path.join(tmpdir, "certificados.zip")
                shutil.make_archive(zip_path.replace(".zip", ""), "zip", output_dir)

                # Mostrar mensaje de éxito
                st.success("✅ Certificados generados correctamente.")
                st.write("Podés descargar todos los archivos a continuación:")

                # Botón de descarga
                with open(zip_path, "rb") as f:
                    st.download_button("⬇️ Descargar ZIP", f, "certificados.zip", "application/zip")
