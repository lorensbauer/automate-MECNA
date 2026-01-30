from tempfile import TemporaryDirectory

from PyPDF2 import PdfReader, PdfWriter
import pandas as pd
from docx2pdf import convert
from pathlib import Path
import os
from dotenv import load_dotenv

from core.certificatePDF import sign_pdf
from core.fillTemplate import funcionConDocxtpl
from core.sendEmail import send_email

load_dotenv()
if __name__ == "__main__":
    plantilla_path = "input/Ziurtagiria2.docx"
    mecna_names = pd.read_excel(
        io="input/MECNA _korreoekin.xlsx",
        usecols="A:E",
        header=0,
        index_col=None,
        sheet_name="2025 (uztailetik)",
    )
    mecna_names = mecna_names.rename(columns={"NAN zkia": "NAN_zkia"})
    with TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        # 1. Rellenar plantilla usando docxtpl
        doc = funcionConDocxtpl(mecna_names)
        docx_path = tmpdir / "plantillaRellenada.docx"
        doc.save(docx_path)

        # 2. Convertir a pdf
        pdf_path = tmpdir / "doc2pdf.pdf"
        convert(docx_path, pdf_path)

        # 3. Dividir pdf en tantos pdfs como paginas

        reader = PdfReader(pdf_path)

        for i, page in enumerate(reader.pages):
            if i != len(reader.pages) - 1:
                writer = PdfWriter()
                writer.add_page(page)
                with open(tmpdir / "doc.pdf", "wb") as out:
                    writer.write(out)
                # Firmar documento
                pdf = sign_pdf(tmpdir / "doc.pdf")
                # Enviar por correo
                subject = "MECNA ziurtagiria"
                sender = "99lotermin@gmail.com"
                recipients = [mecna_names.loc[i, "Email"]]
                recipients_name = f'{mecna_names.loc[i, "Izena"]} {mecna_names.loc[i, "Abizenak"]}'
                body = (
                    f"Kaixo {recipients_name}\n"
                    "Hemen duzu MECNA-ren ziurtagiria sinatuta. "
                    "Mezu hau automatikoki sortu da, arazorik egotekotan idatz iezagouzu korreo bat mesedez\n"
                    "Agur bero bat"
                )
                password = os.getenv("PASSWORD")

                send_email(subject, body, sender, recipients, password, pdf, recipients_name)
