import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from pptxtopdf import convert
import os
import tempfile

st.set_page_config(page_title="Generador de Certificados", page_icon="🎓", layout="centered")

st.title("🎓 Generador de Certificados")
st.write("Subí tus archivos de asistencia y el template para generar certificados automáticamente.")

uploaded_template = st.file_uploader("📄 Subí el template (.pptx)", type=["pptx"])
uploaded_excel1 = st.file_uploader("📘 Subí listado presencial (.xlsx)")
uploaded_excel2 = st.file_uploader("📗 Subí listado virtual (.xlsx)")

if uploaded_template and uploaded_excel1 and uploaded_excel2:
    if st.button("🚀 Generar certificados"):
        with tempfile.TemporaryDirectory() as tmpdir:
            template_path = os.path.join(tmpdir, "template.pptx")
            with open(template_path, "wb") as f:
                f.write(uploaded_template.read())

            df1 = pd.read_excel(uploaded_excel1)
            df1_asistentes = df1[
                (df1["Cargo"].str.lower() == "asistente")
                & (df1["Asistió"].str.upper() == "SI")
            ].copy()

            df1_asistentes["Nombre y apellido"] = (
                df1_asistentes["Apellido"].str.strip() + " " + df1_asistentes["Nombre"].str.strip()
            ).str.title()

            df2 = pd.read_excel(uploaded_excel2)
            df2["Nombre y apellido"] = (
                df2["Apellido"].str.strip() + " " + df2["Nombre"].str.strip()
            ).str.title()

            df_final = pd.concat(
                [df1_asistentes[["Nombre y apellido"]], df2[["Nombre y apellido"]]],
                ignore_index=True,
            ).drop_duplicates().reset_index(drop=True)

            output_dir = os.path.join(tmpdir, "Certificados")
            os.makedirs(output_dir, exist_ok=True)

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

                output_pptx = os.path.join(output_dir, f"Certificado_{nombre}.pptx")
                prs.save(output_pptx)
                convert(output_pptx, output_dir)

            st.success("✅ Certificados generados correctamente.")
            st.write("Podés descargar la carpeta completa a continuación:")

            import shutil
            zip_path = os.path.join(tmpdir, "certificados.zip")
            shutil.make_archive(zip_path.replace(".zip", ""), "zip", output_dir)
            with open(zip_path, "rb") as f:
                st.download_button("⬇️ Descargar ZIP", f, "certificados.zip", "application/zip")
