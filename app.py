from flask import Flask, render_template, request, jsonify, redirect, url_for, flash, session, send_from_directory
from flask_sqlalchemy import SQLAlchemy
import re, os, time, random, requests
from werkzeug.utils import secure_filename
from docx import Document
import PyPDF2
from datetime import datetime, timedelta
import threading
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ==============================
# CẤU HÌNH ỨNG DỤNG
# ==============================
app = Flask(__name__, static_url_path='/static', static_folder='static', template_folder='templates')
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "secret-key-dev")

# ==============================
# CƠ SỞ DỮ LIỆU (SQLite)
# ==============================
app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get("DATABASE_URL", "sqlite:///math_system.db")
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

# ==============================
# MODEL CƠ SỞ DỮ LIỆU
# ==============================
class Problem(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    text = db.Column(db.Text, nullable=False)
    model_answer = db.Column(db.Text, nullable=False)
    correct_value = db.Column(db.Integer, nullable=False)
    topic = db.Column(db.String(100), default="Thủ công")
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class Submission(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    student_name = db.Column(db.String(50), nullable=False)
    answer = db.Column(db.Text, nullable=False)
    image_path = db.Column(db.String(255), nullable=True)
    ai_score = db.Column(db.Integer, nullable=False)
    ai_level = db.Column(db.String(50), nullable=False)
    ai_level_class = db.Column(db.String(20), nullable=False)
    ai_feedback = db.Column(db.Text, nullable=False)
    ai_duration = db.Column(db.String(50), nullable=True)
    manual_score = db.Column(db.Integer, nullable=True)
    manual_level = db.Column(db.String(50), nullable=True)
    manual_level_class = db.Column(db.String(20), nullable=True)
    manual_feedback = db.Column(db.Text, nullable=True)
    manual_override = db.Column(db.Boolean, default=False)
    time_submitted = db.Column(db.DateTime, default=datetime.utcnow)

# ==============================
# CẤU HÌNH EMAIL (giữ nguyên)
# ==============================
app.config['MAIL_SERVER'] = os.environ.get("MAIL_SERVER", "smtp.gmail.com")
app.config['MAIL_PORT'] = int(os.environ.get("MAIL_PORT", 587))
app.config['MAIL_USE_TLS'] = os.environ.get("MAIL_USE_TLS", "true").lower() == "true"
app.config['MAIL_USERNAME'] = os.environ.get("MAIL_USERNAME", "nguyenduycom89@gmail.com")
app.config['MAIL_PASSWORD'] = os.environ.get("MAIL_PASSWORD", "")
TEACHER_EMAIL = os.environ.get("TEACHER_EMAIL", "nguyenduycom89@gmail.com")
TEACHER_PASSWORD = os.environ.get("TEACHER_PASSWORD", "nguyenmocgiao")

# ==============================
# DỮ LIỆU / UPLOAD
# ==============================
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024
ALLOWED_EXT = {'.docx', '.pdf', '.png', '.jpg', '.jpeg', '.webp'}

# ==============================
# TIỆN ÍCH (giữ nguyên)
# ==============================
def allowed(filename: str) -> bool:
    _, ext = os.path.splitext(filename.lower())
    return ext in ALLOWED_EXT

def extract_text_from_docx(path: str) -> str:
    doc = Document(path)
    return "\n".join([p.text for p in doc.paragraphs])

def extract_text_from_pdf(path: str) -> str:
    text = ""
    with open(path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        for page in reader.pages:
            text += page.extract_text() or ""
    return text

def parse_problem_file(content):
    pattern = r"\[BÀI TOÁN\](.*?)\[ĐÁP ÁN MẪU\](.*?)\[ĐÁP SỐ\](\s*\d+)"
    matches = re.findall(pattern, content, re.S)
    problems = []
    for prob, model, ans in matches:
        try:
            correct = int(re.search(r"\d+", ans).group())
        except:
            correct = 0
        problems.append({
            "text": prob.strip(),
            "model_answer": model.strip(),
            "correct_value": correct,
            "topic": "Tự động"
        })
    return problems

# ==============================
# AI ĐÁNH GIÁ (giữ nguyên)
# ==============================
def evaluate_solution_v2(student_text, model_answer, correct_value):
    text = student_text.lower()
    nums = [int(x) for x in re.findall(r'\d+', text)] if re.findall(r'\d+', text) else []
    score, fb = 0, []
    if any(kw in text for kw in ["1 bao", "mỗi bao", "1 đơn vị", "giá mỗi"]):
        score += 3
        fb.append("✅ Em nhận diện đúng dạng toán 'rút về đơn vị'.")
    else:
        fb.append("💡 Gợi ý: Em hãy tìm giá trị của 1 đơn vị trước.")
    if correct_value in nums:
        score += 4
        fb.append("✅ Em tính đúng đáp số – rất tốt!")
    elif len(nums) >= 2:
        fb.append("✏️ Em cần chia tổng cho số lượng để tìm giá 1 đơn vị, rồi nhân với số cần hỏi.")
    lines = [s for s in student_text.split('\n') if s.strip()]
    if len(lines) >= 3 or "bước" in text:
        score += 3
        fb.append("🌟 Lời giải rõ ràng, đủ bước – rất đáng khen!")
    else:
        fb.append("📄 Gợi ý: Em nên viết lời giải theo từng bước (Bước 1, Bước 2...).")
    score = min(score, 10)
    if score >= 9:
        level, cls = "Hoàn thành xuất sắc", "excellent"
    elif score >= 7:
        level, cls = "Hoàn thành tốt", "good"
    elif score >= 5:
        level, cls = "Hoàn thành", "ok"
    else:
        level, cls = "Chưa hoàn thành", "fail"
    return {"score": score, "level": level, "level_class": cls, "feedback": " ".join(fb)}

# ==============================
# GỬI EMAIL THÔNG BÁO (giữ nguyên)
# ==============================
def send_teacher_notification(student_name, submission_id, url_root):
    try:
        msg = MIMEMultipart()
        msg['From'] = app.config['MAIL_USERNAME']
        msg['To'] = TEACHER_EMAIL
        msg['Subject'] = f"📝 Có bài làm mới từ học sinh: {student_name}"
        body = f"""Kính gửi Thầy Nguyễn Đức Duy,
Học sinh **{student_name}** vừa nộp bài làm (ID: #{submission_id}) trên hệ thống.
Vui lòng đăng nhập vào trang quản trị để xem và chấm điểm:
{url_root}teacher
Trân trọng!"""
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        with smtplib.SMTP(app.config['MAIL_SERVER'], app.config['MAIL_PORT']) as server:
            server.starttls()
            server.login(app.config['MAIL_USERNAME'], app.config['MAIL_PASSWORD'])
            server.send_message(msg)
    except Exception as e:
        print(f"[ERROR] Gửi email thất bại: {e}")

# ==============================
# CONTEXT PROCESSOR
# ==============================
@app.context_processor
def inject_time():
    return dict(time=time)

# ==============================
# KHỞI TẠO CSDL
# ==============================
with app.app_context():
    db.create_all()
    # Thêm bài mặc định nếu chưa có
    if Problem.query.count() == 0:
        default_prob = Problem(
            text="Một cửa hàng bán được 5 bao gạo với số tiền là 400000 đồng. Hỏi cửa hàng đó bán 8 bao gạo được bao nhiêu tiền?",
            model_answer="Giá 1 bao gạo là: 400000 : 5 = 80000 (đồng). Giá 8 bao gạo là: 80000 × 8 = 640000 (đồng). Đáp số: 640000 đồng.",
            correct_value=640000,
            topic="Rút về đơn vị"
        )
        db.session.add(default_prob)
        db.session.commit()

# ==============================
# ROUTES CÔNG KHAI (giữ nguyên logic, chỉ đổi cách lấy dữ liệu)
# ==============================
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/student-login', methods=['GET', 'POST'])
def student_login():
    if request.method == 'POST':
        name = request.form.get('name', '').strip()
        if name and 2 <= len(name) <= 30:
            session['student_name'] = name
            return redirect(url_for('student'))
        flash("❌ Tên học sinh không hợp lệ (2–30 ký tự).", "error")
    return render_template('student_login.html')

@app.route('/student')
def student():
    if not session.get('student_name'):
        return redirect(url_for('student_login'))
    problem = Problem.query.first()  # Lấy bài đầu tiên
    problem_text = problem.text if problem else "Chưa có bài tập."
    return render_template('student.html', problem=problem_text)

# ==============================
# ROUTES GIÁO VIÊN
# ==============================
@app.route('/teacher-login', methods=['GET', 'POST'])
def teacher_login():
    if request.method == 'POST':
        email = request.form.get('email')
        pw = request.form.get('password')
        if email == TEACHER_EMAIL and pw == TEACHER_PASSWORD:
            session['is_teacher'] = True
            flash("✅ Đăng nhập thành công!", "success")
            return redirect(url_for('teacher_dashboard'))
        flash("❌ Sai email hoặc mật khẩu!", "error")
    return render_template('teacher_login.html')

@app.route('/teacher-logout')
def teacher_logout():
    session.pop('is_teacher', None)
    flash("🔒 Đã đăng xuất!", "info")
    return redirect('/')

@app.route('/teacher')
def teacher_dashboard():
    if not session.get('is_teacher'):
        return redirect(url_for('teacher_login'))
    problems = Problem.query.all()
    submissions = Submission.query.order_by(Submission.time_submitted.desc()).all()
    # Định dạng thời gian giống trước
    for sub in submissions:
        sub.time_submitted_formatted = sub.time_submitted.strftime("%H:%M - %d/%m/%Y")
        sub.time_submitted_raw = sub.time_submitted.strftime("%Y-%m-%d %H:%M:%S")
        sub.image_url = url_for('uploaded_file', filename=os.path.basename(sub.image_path)) if sub.image_path else None
        # Tái tạo `result` và `final_result` để tương thích template
        sub.result = {
            "score": sub.ai_score,
            "level": sub.ai_level,
            "level_class": sub.ai_level_class,
            "feedback": sub.ai_feedback,
            "duration": sub.ai_duration
        }
        if sub.manual_override:
            sub.final_result = {
                "score": sub.manual_score,
                "level": sub.manual_level,
                "level_class": sub.manual_level_class,
                "feedback": sub.manual_feedback,
                "duration": sub.ai_duration
            }
        else:
            sub.final_result = sub.result
    return render_template('teacher.html', problems=problems, submissions=submissions, is_teacher=True)

# ==============================
# NỘP BÀI HỌC SINH
# ==============================
@app.route('/submit', methods=['POST'])
def submit():
    if not session.get('student_name'):
        return jsonify({"error": "Bạn chưa đăng nhập. Vui lòng nhập tên trước khi nộp bài."}), 403

    student_name = session['student_name']  # Đảm bảo có giá trị
    start_time = time.time()
    answer_text = ""

    if request.is_json:
        data = request.get_json()
        answer_text = (data.get("answer") or "").strip()
        start_time = float(data.get("start_time", time.time()))
    else:
        answer_text = (request.form.get("answer") or "").strip()
        try:
            start_time = float(request.form.get("start_time", time.time()))
        except (ValueError, TypeError):
            pass

    if not answer_text:
        return jsonify({"error": "Lời giải không được để trống."}), 400

    image_path = None
    if 'image' in request.files:
        file = request.files['image']
        if file and file.filename != '':
            filename = secure_filename(f"{int(time.time())}_{file.filename}")
            ext = os.path.splitext(filename)[1].lower()
            if ext in {'.png', '.jpg', '.jpeg', '.webp'}:
                image_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(image_path)

    current = Problem.query.first()
    if not current:
        return jsonify({"error": "Chưa có bài tập để làm."}), 400

    result = evaluate_solution_v2(answer_text, current.model_answer, current.correct_value)
    dur = int(time.time() - start_time)
    result["duration"] = f"{dur // 60} phút {dur % 60} giây"

    submission = Submission(
        student_name=student_name,
        answer=answer_text,
        image_path=image_path,
        ai_score=result["score"],
        ai_level=result["level"],
        ai_level_class=result["level_class"],
        ai_feedback=result["feedback"],
        ai_duration=result["duration"],
        manual_override=False
    )
    db.session.add(submission)
    db.session.commit()

    threading.Thread(
        target=send_teacher_notification,
        args=(student_name, submission.id, request.url_root),
        daemon=True
    ).start()

    return jsonify(result)

# ==============================
# CHẤM THỦ CÔNG
# ==============================
@app.route('/grade-manual', methods=['POST'])
def grade_manual():
    if not session.get('is_teacher'):
        return jsonify({"status": "error", "message": "Bạn chưa đăng nhập."}), 403
    data = request.get_json(force=True)
    try:
        sid = int(data.get("submission_id"))
    except:
        return jsonify({"status": "error", "message": "submission_id không hợp lệ."}), 400
    sub = Submission.query.get(sid)
    if not sub:
        return jsonify({"status": "error", "message": "Không tìm thấy bài nộp."}), 404

    sub.manual_score = int(data.get("score"))
    sub.manual_level = data.get("level")
    sub.manual_level_class = data.get("level_class")
    sub.manual_feedback = data.get("feedback")
    sub.manual_override = True
    db.session.commit()
    return jsonify({"status": "success", "message": f"✅ Đã chấm thủ công bài #{sid}."})

# ==============================
# BÀI TIẾP THEO & AI
# ==============================
@app.route('/api/next-problem')
def api_next_problem():
    problem = Problem.query.order_by(db.func.random()).first()
    if not problem:
        return jsonify({"problem": "Chưa có bài tập."})
    return jsonify({"problem": problem.text})

@app.route('/generate-ai', methods=['POST'])
def generate_ai():
    # Giữ nguyên logic AI (không thay đổi)
    try:
        data = request.get_json(force=True)
        base_problem = (data.get("problem") or "").strip()
        if not base_problem:
            return jsonify({"status": "error", "message": "Chưa nhập bài mẫu để tạo tương tự."}), 400
        GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "")
        if not GROQ_API_KEY:
            sample = """[BÀI TOÁN]
Một người thợ đóng 6 cái kệ hết 540000 đồng. Hỏi 9 cái kệ hết bao nhiêu tiền?
[ĐÁP ÁN MẪU]
Giá 1 cái kệ là: 540000 : 6 = 90000 (đồng).
Giá 9 cái kệ là: 90000 × 9 = 810000 (đồng).
[ĐÁP SỐ]
810000
"""
            return jsonify({"status": "success", "generated": sample})
        prompt = f"""
Hãy tạo một bài toán có lời văn tương tự dạng toán lớp 4 sau,
nhưng thay đổi dữ kiện số học (tổng, số lượng, đơn vị, vật thể, giá trị)
sao cho hợp lý, giữ nguyên cấu trúc dạng toán và cách giải.
Trả kết quả theo đúng định dạng:
[BÀI TOÁN]
...
[ĐÁP ÁN MẪU]
...
[ĐÁP SỐ]
...
Bài gốc:
{base_problem}
"""
        headers = {"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"}
        body = {
            "model": os.environ.get("GROQ_MODEL", "llama-3.3-70b-versatile"),
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.8,
            "max_tokens": 400
        }
        resp = requests.post("https://api.groq.com/openai/v1/chat/completions", headers=headers, json=body, timeout=60)
        if resp.status_code != 200:
            if resp.status_code == 429 or "insufficient_quota" in resp.text.lower():
                sample = """[BÀI TOÁN]
Một người thợ đóng 6 cái kệ hết 540000 đồng. Hỏi 9 cái kệ hết bao nhiêu tiền?
[ĐÁP ÁN MẪU]
Giá 1 cái kệ là: 540000 : 6 = 90000 (đồng).
Giá 9 cái kệ là: 90000 × 9 = 810000 (đồng).
[ĐÁP SỐ]
810000
"""
                return jsonify({"status": "success", "generated": sample})
            return jsonify({"status": "error", "message": f"Lỗi Groq {resp.status_code}: {resp.text}"}), resp.status_code
        generated = resp.json()["choices"][0]["message"]["content"].strip()
        return jsonify({"status": "success", "generated": generated})
    except Exception as e:
        return jsonify({"status": "error", "message": f"Lỗi AI/Groq: {str(e)}"}), 500

@app.route('/save-teacher', methods=['POST'])
def save_teacher():
    if not session.get('is_teacher'):
        return jsonify({"status": "error", "message": "Không có quyền."}), 403
    data = request.get_json(force=True)
    prob = (data.get("problem") or "").strip()
    model = (data.get("model_answer") or "").strip()
    corr = int(data.get("correct_value") or 0)
    if not prob or not model:
        return jsonify({"status": "error", "message": "Thiếu [BÀI TOÁN] hoặc [ĐÁP ÁN MẪU]."}), 400
    new_prob = Problem(text=prob, model_answer=model, correct_value=corr, topic="Thủ công")
    db.session.add(new_prob)
    db.session.commit()
    return jsonify({"status": "success", "message": f"Đã thêm bài mới (ID {new_prob.id})."})

@app.route('/upload-problem', methods=['POST'])
def upload_problem():
    if not session.get('is_teacher'):
        return jsonify({"status": "error", "message": "Bạn chưa đăng nhập."}), 403
    file = request.files.get('file')
    if not file:
        return jsonify({"status": "error", "message": "Chưa chọn file."}), 400
    filename = secure_filename(file.filename)
    if not allowed(filename):
        return jsonify({"status": "error", "message": "Chỉ hỗ trợ .docx, .pdf"}), 400
    path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(path)
    try:
        if filename.endswith('.docx'):
            content = extract_text_from_docx(path)
        elif filename.endswith('.pdf'):
            content = extract_text_from_pdf(path)
        else:
            return jsonify({"status": "error", "message": "File không phải là tài liệu đề bài."}), 400
    except Exception as e:
        return jsonify({"status": "error", "message": f"Lỗi đọc file: {str(e)}"}), 500
    parsed_problems = parse_problem_file(content)
    if not parsed_problems:
        return jsonify({"status": "error", "message": "Không tìm thấy bài toán nào trong file."}), 400
    for p in parsed_problems:
        prob = Problem(text=p["text"], model_answer=p["model_answer"], correct_value=p["correct_value"], topic=p["topic"])
        db.session.add(prob)
    db.session.commit()
    return jsonify({
        "status": "success",
        "message": f"✅ Đã thêm {len(parsed_problems)} bài mới! Tổng: {Problem.query.count()}"
    })

# ==============================
# PHỤC VỤ ẢNH TỪ THƯ MỤC UPLOADS
# ==============================
@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)
# ==============================
# MAIN
# ==============================
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5001))
    app.run(host='0.0.0.0', port=port, debug=False)