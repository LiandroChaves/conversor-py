import os
import ffmpeg
import tkinter as tk
from tkinter import filedialog, simpledialog
from pdf2docx import Converter
from docx import Document
from reportlab.pdfgen import canvas
from PyPDF2 import PdfReader
import comtypes.client
from PIL import Image

CONVERSION_MAP = {
    ".mp3": ["avi", "mp4", "wav"],
    ".wav": ["mp3", "avi"],
    ".mp4": ["avi", "mkv"],
    ".avi": ["mp4", "mkv"],
    ".jpg": ["png", "bmp", "ico"],
    ".png": ["jpg", "bmp", "ico"],
    ".txt": ["pdf", "docx"],
    ".pdf": ["txt", "docx"],
    ".docx": ["txt", "pdf"]
}

def convert_docx_to_pdf(input_file, output_file):
    """ Converte um arquivo DOCX para PDF usando o Microsoft Word (requer Word instalado). """
    try:
        word = comtypes.client.CreateObject("Word.Application")
        doc = word.Documents.Open(input_file)
        doc.SaveAs(output_file, FileFormat=17) 
        doc.Close()
        word.Quit()
        return True
    except Exception as e:
        print(f"❌ Erro ao converter {input_file} para PDF: {e}")
        return False

def convert_image_to_ico(input_file, output_file):
    """ Converte uma imagem para o formato .ico usando Pillow. """
    try:
        img = Image.open(input_file)
        img.save(output_file, format="ICO")
        print(f"✅ Convertido com sucesso para .ico: {output_file}")
    except Exception as e:
        print(f"❌ Erro ao converter {input_file} para ICO: {e}")

def convert_file(input_files, output_format):
    """ Converte arquivos um por um. """
    output_dir = os.path.join(os.path.expanduser("~"), "converted_files")
    os.makedirs(output_dir, exist_ok=True)

    for file in input_files:
        base_name, ext = os.path.splitext(os.path.basename(file))
        output_file = os.path.join(output_dir, f"{base_name}.{output_format}")

        try:
            if ext in [".mp3", ".wav", ".mp4", ".avi"]:
                ffmpeg.input(file).output(output_file).run(overwrite_output=True)

            elif ext == ".pdf" and output_format == "docx":
                cv = Converter(file)
                cv.convert(output_file, start=0, end=None)
                cv.close()

            elif ext == ".pdf" and output_format == "txt":
                reader = PdfReader(file)
                text = "\n".join([page.extract_text() for page in reader.pages if page.extract_text()])
                with open(output_file, "w", encoding="utf-8") as f:
                    f.write(text)

            elif ext == ".docx" and output_format == "txt":
                doc = Document(file)
                text = "\n".join([para.text for para in doc.paragraphs])
                with open(output_file, "w", encoding="utf-8") as f:
                    f.write(text)

            elif ext == ".docx" and output_format == "pdf":
                if convert_docx_to_pdf(file, output_file):
                    print(f"✅ Convertido com sucesso: {output_file}")
                else:
                    print(f"❌ Erro: O arquivo {output_file} não foi gerado!")

            elif ext == ".txt" and output_format == "docx":
                doc = Document()
                with open(file, "r", encoding="utf-8") as f:
                    for line in f:
                        doc.add_paragraph(line)
                doc.save(output_file)

            elif ext == ".txt" and output_format == "pdf":
                c = canvas.Canvas(output_file)
                with open(file, "r", encoding="utf-8") as f:
                    text = f.readlines()
                y = 800
                for line in text:
                    c.drawString(100, y, line.strip())
                    y -= 15
                c.save()

            elif ext in [".jpg", ".png", ".bmp"] and output_format == "ico":
                convert_image_to_ico(file, output_file)

            if os.path.exists(output_file):
                print(f"✅ Convertido com sucesso: {output_file}")
            else:
                print(f"❌ Erro: O arquivo {output_file} não foi gerado!")

        except Exception as e:
            print(f"❌ Erro ao converter {file}: {e}")

def select_files():
    """ Abre uma janela para o usuário selecionar arquivos. """
    files = filedialog.askopenfilenames(title="Selecione arquivos")
    if not files:
        return
    
    ext = os.path.splitext(files[0])[1].lower()
    options = CONVERSION_MAP.get(ext, [])

    if not options:
        print(f"⚠️ Não há opções de conversão disponíveis para {ext}")
        return
    
    chosen_format = simpledialog.askstring("Escolha o formato", f"Opções para {ext}: {', '.join(options)}\nDigite o formato desejado:")
    
    if chosen_format in options:
        convert_file(files, chosen_format)
    else:
        print(f"⚠️ Formato inválido: {chosen_format}")

root = tk.Tk()
root.withdraw()

select_files()
