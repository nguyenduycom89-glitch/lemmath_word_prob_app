from flask import Flask, render_template, request, jsonify, redirect, url_for, flash
import re
import time
import os
from werkzeug.utils import secure_filename
from docx import Document
import PyPDF2

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'  # Thay bằng chuỗi ngẫu nhiên

# ==============================
# CẤU HÌNH BẢO MẬT GIÁO VIÊN
# ==============================
TEACHER_EMAIL = "nguyenduycom89@gmail.com"  # ← SỬA THÀNH EMAIL THẬT CỦA BẠN
TEACHER_PASSWORD = "nguyenmocgiao"  # ← MẬT KHẨU BẠN TỰ ĐẶT

# ==============================
# DỮ LIỆU
# ==============================
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Danh sách bài toán (ngân hàng)
PROBLEMS = [
    {
        "id": 1,
        "text": "Một cửa hàng bán được 5 bao gạo với số tiền là 400000 đồng. Hỏi cửa hàng đó bán 8 bao gạo được bao nhiêu tiền?",
        "model_answer": "Giá 1 bao gạo là: 400000 : 5 = 80000 (đồng). Giá 8 bao gạo là: 80000 × 8 = 640000 (đồng). Đáp số: 640000 đồng.",
        "correct_value": 640000,
        "topic": "Rút về đơn vị"
    }
]

STUDENT_SUBMISSIONS = []

# ==============================
# HÀM HỖ TRỢ
# ==============================

def extract_text_from_docx(filepath):
    doc = Document(filepath)
    return '\n'.join([p.text for p in doc.paragraphs])

def extract_text_from_pdf(filepath):
    text = ""
    with open(filepath, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        for page in reader.pages:
            text += page.extract_text() or ""
    return text

def parse_problem_file(content):
    """Phân tích nội dung file theo định dạng [BÀI TOÁN], [ĐÁP ÁN MẪU], [ĐÁP SỐ]"""
    problem = ""
    model_answer = ""
    correct_value = 0

    # Tách theo tiêu đề
    if "[BÀI TOÁN]" in content:
        start = content.find("[BÀI TOÁN]") + len("[BÀI TOÁN]")
        end = content.find("[ĐÁP ÁN MẪU]") if "[ĐÁP ÁN MẪU]" in content else len(content)
        problem = content[start:end].strip()

    if "[ĐÁP ÁN MẪU]" in content:
        start = content.find("[ĐÁP ÁN MẪU]") + len("[ĐÁP ÁN MẪU]")
        end = content.find("[ĐÁP SỐ]") if "[ĐÁP SỐ]" in content else len(content)
        model_answer = content[start:end].strip()

    if "[ĐÁP SỐ]" in content:
        start = content.find("[ĐÁP SỐ]") + len("[ĐÁP SỐ]")
        num_str = re.search(r'\d+', content[start:])
        if num_str:
            correct_value = int(num_str.group())

    return problem, model_answer, correct_value

# ==============================
# AI ĐÁNH GIÁ THEO THÔNG TƯ 27/2020 & 29/2022
# ==============================

def evaluate_solution_v2(student_text, model_answer, correct_value):
    """
    Đánh giá theo 3 tiêu chí Thông tư:
    1. Nhận biết & thông hiểu
    2. Vận dụng
    3. Vận dụng cao (diễn đạt, logic)
    """
    text = student_text.lower()
    numbers = [int(x) for x in re.findall(r'\d+', text)] if re.findall(r'\d+', text) else []

    score = 0
    feedback = []

    # Tiêu chí 1: Nhận biết dạng toán (rút về đơn vị)
    if any(kw in text for kw in ["1 bao", "mỗi bao", "một bao", "giá mỗi", "tiền mỗi"]):
        score += 3
        feedback.append("✅ Em nhận diện đúng dạng toán 'rút về đơn vị' – rất tốt!")
    else:
        feedback.append("💡 Gợi ý: Em hãy tìm giá trị của 1 đơn vị trước.")

    # Tiêu chí 2: Vận dụng tính toán
    if correct_value in numbers:
        score += 4
        feedback.append("✅ Em tính đúng đáp số – tuyệt vời!")
    elif len(numbers) >= 2:
        # Kiểm tra có phép chia và nhân hợp lý
        if any(n == 400000 for n in numbers) and any(n == 5 for n in numbers):
            feedback.append("📝 Em đã ghi đúng dữ kiện, nhớ chia để tìm 1 bao.")
        if any(n == 8 for n in numbers) and any(n == 80000 for n in numbers):
            feedback.append("✏️ Em nhớ nhân giá 1 bao với 8 để ra kết quả.")

    # Tiêu chí 3: Diễn đạt lời giải (theo Thông tư: trình bày rõ ràng)
    sentences = student_text.split('\n')
    if len(sentences) >= 3 or ('-' in student_text) or ('bước' in text):
        score += 3
        feedback.append("🌟 Em trình bày lời giải rõ ràng, đủ bước – rất đáng khen!")
    else:
        feedback.append("📄 Gợi ý: Em nên viết lời giải theo từng bước (Bước 1, Bước 2...)")

    score = min(score, 10)

    # Ánh xạ sang 4 mức theo Thông tư 29/2022
    if score >= 9:
        level = "Hoàn thành xuất sắc"
        level_class = "excellent"
    elif score >= 7:
        level = "Hoàn thành tốt"
        level_class = "good"
    elif score >= 5:
        level = "Hoàn thành"
        level_class = "ok"
    else:
        level = "Chưa hoàn thành"
        level_class = "fail"

    return {
        "score": score,
        "level": level,
        "level_class": level_class,
        "feedback": " ".join(feedback),
        "criteria": {
            "nhan_biet": score >= 3,
            "van_dung": score >= 7,
            "trinh_bay": "trình bày" in " ".join(feedback).lower()
        }
    }

# ==============================
# ROUTES
# ==============================

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/student')
def student():
    # Lấy bài toán đầu tiên (hoặc ngẫu nhiên)
    problem = PROBLEMS[0] if PROBLEMS else {"text": "Chưa có bài tập."}
    return render_template('student.html', problem=problem["text"])

@app.route('/teacher-login', methods=['GET', 'POST'])
def teacher_login():
    if request.method == 'POST':
        email = request.form.get('email')
        password = request.form.get('password')
        if email == TEACHER_EMAIL and password == TEACHER_PASSWORD:
            return redirect(url_for('teacher_dashboard'))
        else:
            flash("❌ Sai email hoặc mật khẩu!", "error")
    return render_template('teacher_login.html')

@app.route('/teacher')
def teacher_dashboard():
    # Kiểm tra đã đăng nhập chưa? (đơn giản: kiểm tra session hoặc redirect)
    # Trong phiên bản này: chỉ vào được qua /teacher-login
    return render_template(
        'teacher.html',
        problems=PROBLEMS,
        submissions=STUDENT_SUBMISSIONS
    )

@app.route('/upload-problem', methods=['POST'])
def upload_problem():
    if 'file' not in request.files:
        return jsonify({"status": "error", "message": "Chưa chọn file"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"status": "error", "message": "File rỗng"}), 400

    if file:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        # Đọc nội dung
        try:
            if filename.endswith('.docx'):
                content = extract_text_from_docx(filepath)
            elif filename.endswith('.pdf'):
                content = extract_text_from_pdf(filepath)
            else:
                return jsonify({"status": "error", "message": "Chỉ hỗ trợ .docx và .pdf"}), 400

            problem, model_answer, correct_value = parse_problem_file(content)

            if not problem:
                return jsonify({"status": "error", "message": "Không tìm thấy [BÀI TOÁN] trong file"}), 400

            # Thêm vào ngân hàng
            new_id = max([p["id"] for p in PROBLEMS], default=0) + 1
            PROBLEMS.append({
                "id": new_id,
                "text": problem,
                "model_answer": model_answer,
                "correct_value": correct_value,
                "topic": "Tự động"
            })

            return jsonify({"status": "success", "message": f"Đã thêm bài tập mới! Tổng: {len(PROBLEMS)} bài."})

        except Exception as e:
            return jsonify({"status": "error", "message": f"Lỗi xử lý file: {str(e)}"}), 500

@app.route('/submit', methods=['POST'])
def submit():
    try:
        data = request.get_json()
        answer = data.get("answer", "").strip()
        start_time = data.get("start_time", time.time())

        if not answer:
            return jsonify({"error": "Lời giải không được để trống"}), 400

        # Dùng bài toán đầu tiên để chấm
        current = PROBLEMS[0] if PROBLEMS else {"model_answer": "", "correct_value": 0}
        result = evaluate_solution_v2(answer, current["model_answer"], current["correct_value"])

        duration_sec = time.time() - start_time
        minutes = int(duration_sec // 60)
        seconds = int(duration_sec % 60)
        result["duration"] = f"{minutes} phút {seconds} giây"

        STUDENT_SUBMISSIONS.append({
            "answer": answer,
            "result": result,
            "time_submitted": time.strftime("%H:%M - %d/%m/%Y", time.localtime())
        })

        return jsonify(result)

    except Exception as e:
        return jsonify({"error": "Có lỗi khi xử lý bài làm"}), 500

@app.route('/ask-teacher', methods=['POST'])
def ask_teacher():
    return jsonify({
        "status": "success",
        "message": "Tin nhắn của em đã được gửi đến thầy/cô! Thầy/cô sẽ trả lời sớm nhất nhé."
    })

# ==============================
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5001))  # Render cung cấp biến PORT
    app.run(host='0.0.0.0', port=port)