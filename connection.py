from pdfminer.high_level import extract_text
import core.pdfnormal as pdfnormal

from core.pdfnormal import extract_pdf_edges
import core.pdfscanned as pdfscanned
import core.autocorrect as autocorrect


pdf_path = "test.pdf"


text = extract_text(pdf_path)
h_edges, v_edges = extract_pdf_edges(pdf_path)
text = text.strip()

pdf_path = autocorrect.auto_correct_pdf_per_page(pdf_path)



if text and len(h_edges) > 2 and len(v_edges) > 2:
    
    wb = pdfnormal.run(pdf_path,None)
    wb.save("final.xlsx")
else:
   wb = pdfscanned.run(pdf_path,None)
   wb.save("final.xlsx")
