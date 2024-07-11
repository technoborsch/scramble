import docx

if __name__ == "__main__":
    doc = docx.Document(r"materials\AKU.0120.10UKC.SGK.PT.TB0002-MAB0001.docx")

    for table in doc.tables:
        for row in table.rows[1:-2]:
            for cell in row.cells:
                print(cell.text)
