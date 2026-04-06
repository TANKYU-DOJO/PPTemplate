import pypdf as pdf
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import os

root = tk.Tk()
root.withdraw()  

pdf_path = filedialog.askopenfilename(
    title="名札PDFファイルを選んでください",
    filetypes=[("PDFファイル", "*.pdf")]
)

if not pdf_path:
    messagebox.showerror("エラー", "PDFファイルが選ばれていません。")
    exit()

output_path = os.path.splitext(pdf_path)[0] + "_入稿用.pdf"

reader = pdf.PdfReader(pdf_path)
writer = pdf.PdfWriter()

width = reader.pages[0].mediabox.width
height = reader.pages[0].mediabox.height

for i in range(0, len(reader.pages), 8):
    page = pdf.PageObject.create_blank_page(width=width*2, height=height*4)
    for x in range(2):
        for y in range(4):
            j = i + x + y*2
            if j < len(reader.pages):
                page.merge_translated_page(reader.pages[j], width*x, height*y)
    writer.add_page(page)

with open(output_path, "wb") as f:
    writer.write(f)

messagebox.showinfo("完了", f"入稿用PDFを作成しました！\n\n{output_path}")
