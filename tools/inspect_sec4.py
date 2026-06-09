"""Inspect section 4 output in generated report."""
from docx import Document

doc = Document(r'C:\Users\houping\Desktop\配方表助手\test_output.docx')
in_sec4 = False
for i, p in enumerate(doc.paragraphs):
    txt = p.text.strip()
    if '四、' in txt and '配方' in txt:
        in_sec4 = True
        print(f'=== Section 4 heading P{i} ===')
        continue
    if in_sec4:
        if '五、' in txt:
            break
        if txt:
            runs_info = []
            for ri, r in enumerate(p.runs):
                runs_info.append(f'R{ri}: bold={r.font.bold} text="{r.text[:50]}"')
            print(f'P{i}: {" | ".join(runs_info)}')
