from docx import Document
from docx.oxml.ns import qn

doc = Document(r'C:\Users\houping\Desktop\配方表助手\test_output.docx')

for i, para in enumerate(doc.paragraphs):
    text = para.text.strip()
    if not text:
        continue
    print(f'P{i} style="{para.style.name}"')
    for j, run in enumerate(para.runs):
        f = run.font
        rPr = run._element.rPr
        rFonts = rPr.find(qn('w:rFonts')) if rPr is not None else None
        ascii_name = rFonts.get(qn('w:ascii')) if rFonts is not None else None
        ea_name = rFonts.get(qn('w:eastAsia')) if rFonts is not None else None
        hAnsi = rFonts.get(qn('w:hAnsi')) if rFonts is not None else None
        print(f'  R{j}: bold={f.bold} sz={f.size} ascii="{ascii_name}" ea="{ea_name}" hAnsi="{hAnsi}" txt="{run.text[:60]}"')
