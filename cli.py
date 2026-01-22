import pandas as pd
from docx2pdf import convert
from docxtpl import DocxTemplate
from datetime import datetime
import docx2pdf


def funcionConDocxtpl(df):

    # Pasar a dict para poder reccorer
    socios = df.to_dict(orient="records")

    # Cargar plantilla
    doc = DocxTemplate("files/Ziurtagiria1.docx")
    # doc = DocxTemplate("files/Prueba1.docx")

    # Renderizar
    doc.render({
        "socios": socios,
        "salto": "\n\n\n\n\n\n\n\n",
        "data": datetime.now().strftime("%Y-%m-%d")
    })

    # guardar resultado
    doc.save("files/Ziurtagiriak_MECNA.docx")

    return doc

if __name__ == '__main__':
    mecna_names = pd.read_excel("files/MECNA.xlsx", usecols="A:D", header=0, index_col=None,
                                sheet_name="2025 (uztailetik)")

    mecna_names = mecna_names.rename(columns={"NAN zkia": "NAN_zkia"})

    doc = funcionConDocxtpl(mecna_names)
    # Convertir a pdf
    convert("files/Ziurtagiriak_MECNA.docx", "files/Ziurtagiria_MECNA.pdf")

    exit()
