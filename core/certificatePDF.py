import os
from io import BytesIO

from pyhanko.sign import signers
from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
from dotenv import load_dotenv

load_dotenv()

def sign_pdf(path_pdf_to_sign):
    signer = signers.SimpleSigner.load_pkcs12(
        'certIT.p12',
        passphrase=os.getenv("PASSWORD_DEL_CERTIFICADO").encode()
    )

    with open(path_pdf_to_sign, "rb") as pdf:
        writer = IncrementalPdfFileWriter(pdf, strict=False)
        outputBuffer = BytesIO()
        signers.sign_pdf(
            writer,
            signature_meta=signers.PdfSignatureMetadata(
                field_name="Firma_FNMT",
                reason="Certificacion MECNA",
                location="Pamplona"
            ),
            signer=signer,
            output=outputBuffer
        )
    outputBuffer.seek(0)
    return outputBuffer.getvalue()


def firmar_pdf():
    signer = signers.SimpleSigner.load_pkcs12(
        "../certIT.p12",
        passphrase=os.getenv("PASSWORD_DEL_CERTIFICADO").encode()
    )

    # Firmar PDF
    with open("../input/Gmail - Kenkaria errenta aitorpenean.pdf", "rb") as inf:
        writer = IncrementalPdfFileWriter(inf)

        with open("../output/documento_firmado.pdf", "wb") as outf:
            signers.sign_pdf(
                writer,
                signature_meta=signers.PdfSignatureMetadata(
                    field_name="Firma_FNMT",
                    reason="Certificación MECNA",
                    location="Pamplona"
                ),
                signer=signer,
                output=outf
            )

if __name__ == '__main__':
    firmar_pdf()
    print(os.getenv("PASSWORD_DEL_CERTIFICADO"))