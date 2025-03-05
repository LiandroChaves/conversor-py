import os
import re
import ffmpeg
import tkinter as tk
from tkinter import filedialog, simpledialog
from pdf2docx import Converter
from docx import Document
from reportlab.pdfgen import canvas
from PyPDF2 import PdfReader
import comtypes.client
from PIL import Image
from unidecode import unidecode 
import subprocess

CONVERSION_MAP = {
    ".mp3": ["avi", "mp4", "wav"],
    ".wav": ["mp3", "avi"],
    ".mp4": ["avi", "mkv", "webm"],
    ".avi": ["mp4", "mkv", "webm"],
    ".jpg": ["png", "bmp", "ico", "webp", "jpeg"],
    ".png": ["jpg", "bmp", "ico", "webp", "jpeg"],
    ".txt": ["pdf", "docx"],
    ".pdf": ["txt", "docx"],
    ".docx": ["txt", "pdf"],
    ".webp": ["jpg", "png", "bmp", "jpeg"],
    ".bmp": ["jpg", "png", "webp", "jpeg"]
}

def sanitize_filename(filename):
    """
    Remove caracteres especiais, acentos e substitui espaços por '-'.
    """
    filename = unidecode(filename)
    filename = re.sub(r'[^\w\-\./]', '', filename)
    filename = re.sub(r'-+/', '', filename)
    filename = filename.strip('')
    print(filename)
    return filename

def convert_docx_to_pdf(input_file, output_file):
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

def convert_image(input_file, output_file, output_format):
    try:
        img = Image.open(input_file)
        img = img.convert("RGB") if output_format in ["jpeg", "jpg"] else img
        img.save(output_file, format=output_format.upper())
        print(f"✅ Convertido com sucesso para {output_format}: {output_file}")
    except Exception as e:
        print(f"❌ Erro ao converter {input_file} para {output_format}: {e}")


def convert_video(input_file, output_file):
    try:
        print(f"Convertendo arquivo: {input_file} para {output_file}")
        
        # Usando subprocess para chamar o FFmpeg diretamente
        command = [
            'ffmpeg', '-i', input_file, '-vcodec', 'libx264', '-acodec', 'aac',
            '-b:v', '2000k', '-f', 'mp4', output_file
        ]
        
        result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        
        if result.returncode == 0:
            print(f"✅ Vídeo convertido com sucesso: {output_file}")
        else:
            print(f"❌ Erro ao converter {input_file}: {result.stderr}")
    
    except Exception as e:
        print(f"❌ Erro ao converter {input_file}: {e}")

def convert_file(input_files, output_format):
    output_dir = os.path.join(os.path.expanduser("~"), "converted_files")
    os.makedirs(output_dir, exist_ok=True)

    for file in input_files:
        base_name, ext = os.path.splitext(os.path.basename(file))
        base_name = sanitize_filename(base_name)
        output_file = os.path.join(output_dir, f"{base_name}.{output_format}")

        try:
            if ext in [".mp3", ".wav", ".mp4", ".avi"]:
                convert_video(file, output_file)
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
            elif ext in [".jpg", ".png", ".bmp", ".webp"]:
                convert_image(file, output_file, output_format)

            if os.path.exists(output_file):
                print(f"✅ Convertido com sucesso: {output_file}")
            else:
                print(f"❌ Erro: O arquivo {output_file} não foi gerado!")
        except Exception as e:
            print(f"❌ Erro ao converter {file}: {e}")

def select_files():
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