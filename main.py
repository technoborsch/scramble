# import io
import os

from pypdf import PdfWriter, PdfReader
# from reportlab.pdfgen import canvas
# from reportlab.lib.pagesizes import A4

# packet = io.BytesIO()
# can = canvas.Canvas(packet, pagesize=A4)
# can.drawString(10, 100, "Тест")
# can.save()

directory = r"C:\Users\n.nikiforov\PycharmProjects\IIhelper\3_ИИ"
pdf_list = []

for file in os.listdir(directory):
    filename = os.fsdecode(file)
    if filename.endswith(".pdf"):
        pdf_list.append(file)
    print(file)
print(pdf_list)

# move to the beginning of the StringIO buffer
# packet.seek(0)

# create a new PDF with Reportlab
# new_pdf = PdfReader(packet)
stamp = PdfReader(open("stampA4.pdf", "rb"))
# read your existing PDF
output = PdfWriter()
# add the "watermark" (which is the new pdf) on the existing page
for pdf_name in pdf_list:
    pdf_path = os.path.join(directory, pdf_name)
    existing_pdf = PdfReader(open(pdf_path, "rb"))
    for page in existing_pdf.pages:
        if not pdf_name.startswith("!"):
            page.merge_translated_page(stamp.pages[0], 0, 30)
        # page.merge_page(new_pdf.pages[0])
        output.add_page(page)
# finally, write "output" to a real file
output_stream = open("destination.pdf", "wb")
output.write(output_stream)
output_stream.close()
