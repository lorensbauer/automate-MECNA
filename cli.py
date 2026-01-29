import os
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd
from docx2pdf import convert
from dotenv import load_dotenv

from core.certificatePDF import sign_pdf
from core.fillTemplate import fillTemplate
from core.sendEmail import send_email

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

    # Para hacer pruebas probamos con la dos primeras personas
    mecna_names = mecna_names.loc[:2, :]

    lista_socios = mecna_names.to_dict(orient="records")
    with TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        for socio in lista_socios:
            # 1. Rellenar DOCX
            doc = fillTemplate(socio, plantilla_path)
            docx_path = tmpdir / "plantillaRellenada.docx"
            doc.save(docx_path)

            # 2. DOCX → PDF
            pdf_path = tmpdir / "doc2pdf.pdf"
            convert(docx_path, pdf_path)

            # 3. Firmar PDF (en memoria)
            pdf_signed_bytes = sign_pdf(pdf_path)
            # with open(tmpdir / "pdf_signed.pdf", "wb") as f:
            #     f.write(pdf_signed_bytes)

            # 4. Enviar por correo
            subject = "MECNA ziurtagiria"
            sender = "99lotermin@gmail.com"
            recipients = [socio["Email"]]
            recipients_name = f'{socio["Izena"]} {socio["Abizenak"]}'
            body = (
                f"Kaixo {recipients_name}\n"
                "Hemen duzue MECNA-ren ziurtagiria sinatuta. "
                "Mezu hau automatikoki sortu da, arazorik egotekotan idatzi hurrengo posta el"
            )
            password = os.getenv("PASSWORD")

            send_email(subject, body, sender, recipients, password, pdf_signed_bytes, recipients_name)

            break
