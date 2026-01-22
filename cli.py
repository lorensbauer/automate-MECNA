import pandas as pd
from docx import Document
from docxtpl import DocxTemplate
from datetime import date, datetime


def funcionConDocxtpl(df):
    # pasar a lista de dicts
    print(df["Izena"])
    socios = df.to_dict(orient="records")

    # cargar plantilla
    doc = DocxTemplate("files/Ziurtagiria1.docx")
    # doc = DocxTemplate("files/Prueba1.docx")

    # renderizar
    doc.render({
        "socios": socios,
        "salto": "\n\n\n\n\n\n\n\n",
        "data": datetime.now().strftime("%Y-%m-%d")
    })

    # guardar resultado
    doc.save("files/Ziurtagiriak_MECNA.docx")



if __name__ == '__main__':
    mecna_names = pd.read_excel("files/MECNA.xlsx", usecols="A:D", header=0, index_col=None,
                                sheet_name="2025 (uztailetik)")

    mecna_names = mecna_names.rename(columns={"NAN zkia": "NAN_zkia"})

    funcionConDocxtpl(mecna_names)

    exit()
