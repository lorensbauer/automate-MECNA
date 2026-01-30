import datetime
import os
from io import BytesIO

from pyhanko.sign import signers, PdfSigner, fields
from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
from dotenv import load_dotenv
from pyhanko.sign.fields import SigFieldSpec
from pyhanko.stamp import TextStampStyle

load_dotenv()


def sign_pdf(path_pdf_to_sign):
    signer = signers.SimpleSigner.load_pkcs12(
        "certIT.p12", passphrase=os.getenv("PASSWORD_DEL_CERTIFICADO").encode()
    )

    with open(path_pdf_to_sign, "rb") as pdf:
        writer = IncrementalPdfFileWriter(pdf, strict=False)
        outputBuffer = BytesIO()
        signers.sign_pdf(
            writer,
            signature_meta=signers.PdfSignatureMetadata(
                field_name="Firma_FNMT", reason="Certificacion MECNA", location="Pamplona"
            ),
            signer=signer,
            output=outputBuffer,
        )
    outputBuffer.seek(0)
    return outputBuffer.getvalue()


def sign_pdf_visible(path_pdf_to_sign):
    # Cargar certificado FNMT
    signer = signers.SimpleSigner.load_pkcs12(
        "certIT.p12", passphrase=os.getenv("PASSWORD_DEL_CERTIFICADO").encode()
    )

    with open(path_pdf_to_sign, "rb") as f:
        writer = IncrementalPdfFileWriter(f, strict=False)

        # (Opcional pero recomendado) definir posición de la firma
        fields.append_signature_field(
            writer,
            fields.SigFieldSpec(
                sig_field_name="Firma_FNMT",
                box=(50, 50, 300, 150),  # coordenadas
                on_page=0,  # primera página
            ),
        )

        meta = signers.PdfSignatureMetadata(
            field_name="Firma_FNMT", reason="Certificación MECNA", location="Pamplona"
        )

        stamp = TextStampStyle(
            stamp_text=(
                "Iruña taldeko lehendakariak digitalki sinatutako dokumentua\n"
                f"Data: {datetime.now().strftime("%Y-%m-%d")}\n"
                "MECNA Zurtagiria"
            )
        )

        pdf_signer = PdfSigner(signature_meta=meta, signer=signer, stamp_style=stamp)

        output = BytesIO()
        pdf_signer.sign_pdf(writer, output=output)

    output.seek(0)
    return output.getvalue()


def firmar_pdf():
    signer = signers.SimpleSigner.load_pkcs12(
        "../certIT.p12", passphrase=os.getenv("PASSWORD_DEL_CERTIFICADO").encode()
    )

    # Firmar PDF
    with open("../input/Gmail - Kenkaria errenta aitorpenean.pdf", "rb") as inf:
        writer = IncrementalPdfFileWriter(inf)

        with open("../output/documento_firmado.pdf", "wb") as outf:
            signers.sign_pdf(
                writer,
                signature_meta=signers.PdfSignatureMetadata(
                    field_name="Firma_FNMT", reason="Certificación MECNA", location="Pamplona"
                ),
                signer=signer,
                output=outf,
            )


if __name__ == "__main__":
    firmar_pdf()
    print(os.getenv("PASSWORD_DEL_CERTIFICADO"))
