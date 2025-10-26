import os, re
from docx import Document
import PyPDF2

# Danh sách định dạng hợp lệ
ALLOWED_EXTENSIONS = {'.docx', '.pdf', '.png', '.jpg', '.jpeg', '.webp'}

def allowed(filename: str) -> bool:
    """Kiểm tra file hợp lệ theo phần mở rộng"""
    return os.path.splitext(filename)[1].lower() in ALLOWED_EXTENSIONS

def extract_text_from_docx(path):
    """Đọc nội dung từ file Word"""
    try:
        doc = Document(path)
        return "\n".join([p.text for p in doc.paragraphs])
    except Exception as e:
        return f"Lỗi đọc DOCX: {e}"

def extract_text_from_pdf(path):
    """Đọc nội dung từ file PDF"""
    text = ""
    try:
        with open(path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            for page in reader.pages:
                text += page.extract_text() + "\n"
        return text
    except Exception as e:
        return f"Lỗi đọc PDF: {e}"

def parse_problem_file(content: str):
    """
    Phân tích file bài toán dạng:
    [BÀI TOÁN] ... [ĐÁP ÁN MẪU] ... [ĐÁP SỐ] ...
    """
    pattern = r"\[BÀI TOÁN\](.*?)\[ĐÁP ÁN MẪU\](.*?)\[ĐÁP SỐ\](.*)"
    matches = re.findall(pattern, content, flags=re.S)
    problems = []
    for m in matches:
        text, model, answer = [x.strip() for x in m]
        value = None
        mnum = re.search(r"(\d[\d\.]*)", answer)
        if mnum:
            try: value = float(mnum.group(1).replace(".", ""))
            except: value = None
        problems.append({
            "text": text,
            "model_answer": model,
            "correct_value": value,
            "topic": "Rút về đơn vị"
        })
    return problems
