import pdfplumber
import re
import pandas as pd
if __name__ == "__main__":
    texto = ""
    with pdfplumber.open("input/Gmail - Kenkaria errenta aitorpenean.pdf") as pdf:
        for page in pdf.pages[1:]:
            texto += page.extract_text() + "\n"
    emails = re.findall(r'<([^<>@\s]+@[^<>@\s]+)>', texto)

    emails = [
        e for e in emails
        if e.lower() != "irunataldeaelkartea@gmail.com"
    ]

    df = pd.DataFrame(emails, columns=["Email"])
    df.to_excel("output/email.xlsx", index=False)
    print(emails)