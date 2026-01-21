import pandas as pd
from docx import Document
from docxtpl import DocxTemplate
from datetime import date

def funcionConDocxtpl(df):
    # pasar a lista de dicts
    socios = df.to_dict(orient="records")

    # cargar plantilla
    doc = DocxTemplate("files/Ziurtagiria.docx")

    # renderizar
    doc.render({
        "socios": socios
    })

    # guardar resultado
    doc.save("Ziurtagiriak_MECNA.docx")


def funcionConDocx(lista_socios):
    doc = Document('files/Ziurtagiria2.docx')

    for idx, socio in enumerate(lista_socios):
        # list of all words to be replaced, with its new word
        replace_word = {'{{Izena}}':'Loren', '{{Abizenak}}':'Otermin Motos','{{NAN_zkia}}':'4564654','{{Zenbatekoa}}':'90'}

        # to replace words within paragraphs
        for word in replace_word:
            for p in doc.paragraphs:
                if p.text.find(word) >= 0:
                    p.text = p.text.replace(word, replace_word[word])

        # to replace words within tables
        for word in replace_word:
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            if p.text.find(word) >= 0:
                                p.text = p.text.replace(word, replace_word[word])

        # to replace words within headers
        for word in replace_word:
            for section in doc.sections:
                header = section.header
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for p in cell.paragraphs:
                                if p.text.find(word) >= 0:
                                    p.text = p.text.replace(word, replace_word[word])

        # to replace words within footers
        for word in replace_word:
            for footer in doc.sections:
                footer = section.footer
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for p in cell.paragraphs:
                                if p.text.find(word) >= 0:
                                    p.text = p.text.replace(word, replace_word[word])
    doc.save('Output.docx')


if __name__ == '__main__':
    mecna_names = pd.read_excel("files/MECNA.xlsx", usecols="A:D", header=0, index_col=None,
                                sheet_name="2025 (uztailetik)")

    mecna_names = mecna_names.rename(columns={"NAN zkia": "NAN_zkia"})
    mecna_names = mecna_names.iloc[:2,:]

    # funcionConDocxtpl(mecna_names)
    funcionConDocx(mecna_names)

    exit()