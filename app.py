
import importlib.metadata
import streamlit as st
import pandas as pd
from pptx import Presentation
from io import BytesIO
import tempfile
from copy import deepcopy
import re

# Evita errore su importlib.metadata.version
importlib.metadata.version = lambda name: "1.48.0" if name == "streamlit" else importlib.metadata.version(name)

def replace_text_in_shapes(slide, data_dict):
    pattern = r"<(.*?)>"
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    matches = re.findall(pattern, run.text)
                    for key in matches:
                        value = data_dict.get(key.strip(), "")
                        run.text = run.text.replace(f"<{key}>", str(value))
    return slide

def duplicate_slide(prs, source_slide):
    slide_layout = prs.slide_layouts[0]
    new_slide = prs.slides.add_slide(slide_layout)
    for shape in source_slide.shapes:
        try:
            # Per immagini
            if shape.shape_type == 13:
                image_stream = BytesIO(shape.image.blob)
                new_slide.shapes.add_picture(
                    image_stream, shape.left, shape.top, shape.width, shape.height
                )
            else:
                # Copia XML di forme (linee, box, rettangoli, ecc.)
                el = shape.element
                new_el = deepcopy(el)
                new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
        except Exception:
            pass
    return new_slide

st.set_page_config(page_title="Generatore Slide Bennet", layout="centered")
st.title("📊 Generatore Slide Bennet")

ppt_file = st.file_uploader("📥 Carica il template PowerPoint (.pptx)", type=["pptx"])
excel_file = st.file_uploader("📊 Carica il file Excel (.xlsx o .xls)", type=["xlsx", "xls"])

if ppt_file and excel_file:
    df = pd.read_excel(excel_file)
    st.success(f"{len(df)} righe caricate dal file Excel.")

    if st.button("🚀 Genera PowerPoint"):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
            tmp.write(ppt_file.read())
            template_path = tmp.name

        template_ppt = Presentation(template_path)
        template_slide = template_ppt.slides[0]

        final_ppt = Presentation()
        final_ppt.slides._sldIdLst.clear()
        final_ppt.slide_width = template_ppt.slide_width
        final_ppt.slide_height = template_ppt.slide_height

        for _, row in df.iterrows():
            data_dict = {
                str(k).strip(): "" if pd.isna(v) else str(v).strip()
                for k, v in row.items()
            }
            new_slide = duplicate_slide(final_ppt, template_slide)
            replace_text_in_shapes(new_slide, data_dict)

        with BytesIO() as result:
            final_ppt.save(result)
            result.seek(0)
            st.download_button(
                "📥 Scarica PowerPoint Compilato",
                data=result,
                file_name="output_bennet_finale.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
