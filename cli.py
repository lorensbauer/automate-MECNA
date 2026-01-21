import pandas as pd
from docx import Document

def reemplazar_texto(doc, reemplazos: dict):
    for p in doc.paragraphs:
        for clave, valor in reemplazos.items():
            if clave in p.text:
                p.text = p.text.replace(clave, str(valor))

def copiar_plantilla(doc_destino, doc_origen):
    for p in doc_origen.paragraphs:
        nuevo_p = doc_destino.add_paragraph(p.text)
        nuevo_p.style = p.style


if __name__ == '__main__':
    mecna_names = pd.read_excel("files/MECNA.xlsx", usecols="A:D", header=0, index_col=None,
                                sheet_name="2025 (uztailetik)")

    plantilla = Document("files/Ziurtagiria.docx")
    doc_final = Document()

    for i, socio in mecna_names.iterrows():
        doc_temp = plantilla

        reemplazos = {
            "{{Izena}}": socio["Izena"],
            "{{Abizenak}}": socio["Abizenak"],
            "{{NAN_zkia}}": socio["NAN zkia"],
            "{{Zenbatekoa}}": socio["Zenbatekoa"],
        }

        reemplazar_texto(doc_temp, reemplazos)

        copiar_plantilla(doc_final, doc_temp)

        if i == 2:
            break
        if i != len(mecna_names) - 1:
            doc_final.add_page_break()

    doc_final.save("documento_mecna.docx")
