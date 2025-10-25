from flask import Flask, render_template, request, jsonify, redirect, url_for, flash, session
import re
import time
import os
from werkzeug.utils import secure_filename
from docx import Document
import PyPDF2

# =====================================================
# KHỞI TẠO ỨNG DỤNG FLASK CHUẨN CHO DEPLOY
# =====================================================
app = Flask(
    __name__,
    static_url_path='/static',
    static_folder='static',
    template_folder='templates'
)

app.secret_key = os.environ.get("FLASK_SECRET_KEY", "your-secret-key-here")

# =====================================================
# CẤU HÌNH GIÁO VIÊN (TÀI KHOẢN ADMIN)
# =====================================================
TEACHER_EMAIL = os.environ.get("TEACHER_EMAIL", "nguyenduycom89@gmail.com")
TEACHER_PASSWORD = os.environ.get("TEACHER_PASSWORD", "nguyenmocgiao")

# =====================================================
# CẤU HÌNH UPLOAD & DỮ LIỆU
# =====================================================
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'.docx', '.pdf'}
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10MB

PROBLEMS = [
    {
        "id": 1,
        "text": "Một cửa hàng bán được 5 bao gạo với số tiền là 400000 đồng. "
                "Hỏi cửa hàng đó bán 8 bao gạo được bao nhiêu tiền?",
        "model_answer": "Giá 1 bao gạo là: 400000 : 5 = 80000 (đồng). "
                        "Giá 8 bao gạo là: 80000 × 8 = 640000 (đồng). "
                        "Đáp số: 640000 đồng.",
        "correct_value": 640000,
        "topic": "Rút về đơn vị"
    }
]
STUDENT_SUBMISSIONS = []  # mỗi item: {id, answer, result(ai), final_result, manual_override, time_submitted}
SUBMISSION_SEQ = 1

# =====================================================
# TIỆN ÍCH
# =====================================================
def allowed_file_extension(filename: str) -> bool:
    _, ext = os.path.splitext(filename.lower())
    return ext in ALLOWED_EXTENSIONS

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
            try:
                correct_value = int(num_str.group())
            except Exception:
                correct_value = 0

    return problem, model_answer, correct_value

# =====================================================
# AI: ĐÁNH GIÁ THEO TT 27/2020 & 29/2022
# =====================================================
def evaluate_solution_v2(student_text, model_answer, correct_value):
    text = student_text.lower()
    numbers = [int(x) for x in re.findall(r'\d+', text)] if re.findall(r'\d+', text) else []

    score = 0
    feedback = []

    # 1) Nhận biết dạng toán
    if any(kw in text for kw in ["1 bao", "mỗi bao", "một bao", "giá mỗi", "tiền mỗi", "1 đơn vị"]):
        score += 3
        feedback.append("✅ Em nhận diện đúng dạng toán 'rút về đơn vị' – rất tốt!")
    else:
        feedback.append("💡 Gợi ý: Em hãy tìm giá trị của 1 đơn vị trước.")

    # 2) Vận dụng tính toán
    if correct_value in numbers:
        score += 4
        feedback.append("✅ Em tính đúng đáp số – tuyệt vời!")
    elif len(numbers) >= 2:
        if any(n == 400000 for n in numbers) and any(n == 5 for n in numbers):
            feedback.append("📝 Em đã ghi đúng dữ kiện, nhớ chia để tìm 1 bao.")
        if any(n == 8 for n in numbers) and any(n == 80000 for n in numbers):
            feedback.append("✏️ Em nhớ nhân giá 1 bao với 8 để ra kết quả.")

    # 3) Trình bày
    sentences = [s for s in student_text.split('\n') if s.strip()]
    if len(sentences) >= 3 or ('-' in student_text) or ('bước' in text):
        score += 3
        feedback.append("🌟 Em trình bày lời giải rõ ràng, đủ bước – rất đáng khen!")
    else:
        feedback.append("📄 Gợi ý: Em nên viết lời giải theo từng bước (Bước 1, Bước 2...).")

    score = min(score, 10)

    if score >= 9:
        level, level_class = "Hoàn thành xuất sắc", "excellent"
    elif score >= 7:
        level, level_class = "Hoàn thành tốt", "good"
    elif score >= 5:
        level, level_class = "Hoàn thành", "ok"
    else:
        level, level_class = "Chưa hoàn thành", "fail"

    return {
        "score": score,
        "level": level,
        "level_class": level_class,
        "feedback": " ".join(feedback)
    }

# =====================================================
# AI: HƯỚNG DẪN TỪNG BƯỚC (RULE-BASED)
# =====================================================
def ai_step_by_step(problem_text, correct_value=None):
    """
    Trả về hướng dẫn từng bước dạng rút về đơn vị (nếu nhận diện được).
    """
    txt = problem_text.lower()
    nums = [int(x) for x in re.findall(r'\d+', txt)]
    steps = []
    tips = []
    detected = "unknown"

    # Heuristic: 2 số đầu là (tổng tiền, số đơn vị); số hỏi là số thứ 3 (nếu có)
    # Ví dụ: 400000, 5, 8
    if len(nums) >= 2 and ("bao" in txt or "mỗi" in txt or "1 bao" in txt or "1 đơn vị" in txt):
        detected = "rut_ve_don_vi"
        total = nums[0]
        count = nums[1]
        ask = nums[2] if len(nums) >= 3 else None

        # Bước 1
        steps.append(f"Bước 1 (Rút về đơn vị): Tính giá trị **1 đơn vị**: {total} : {count} = {total // count if count else '...'}")
        # Bước 2
        if ask:
            unit = total // count if count else None
            if unit is not None:
                steps.append(f"Bước 2 (Nhân lên): Tính giá trị {ask} đơn vị: {unit} × {ask} = {unit * ask}")
            else:
                steps.append("Bước 2 (Nhân lên): Lấy giá trị 1 đơn vị nhân với số đơn vị cần tìm.")
        else:
            steps.append("Bước 2 (Nhân lên): Lấy giá trị 1 đơn vị nhân với số đơn vị cần tìm (chưa xác định cụ thể trong đề).")

        # Bước 3
        ans = None
        if correct_value is not None and isinstance(correct_value, int) and correct_value > 0:
            ans = correct_value
        elif len(nums) >= 3 and count:
            ans = (total // count) * nums[2]
        steps.append(f"Bước 3 (Kết luận): Viết đáp số kèm đơn vị.{' Đáp số: ' + str(ans) if ans is not None else ''}")

        tips = [
            "Nếu đề cho **tổng** và **số lượng**, ta **chia** để tìm 1 đơn vị.",
            "Muốn biết nhiều đơn vị → **nhân** giá trị 1 đơn vị với số đơn vị cần hỏi.",
            "Luôn ghi **đơn vị** (đồng, kg, lít...)."
        ]
    else:
        steps = [
            "Bước 1: Xác định 1 đơn vị (chia tổng cho số lượng).",
            "Bước 2: Nhân giá 1 đơn vị với số lượng cần hỏi.",
            "Bước 3: Viết đáp số và đơn vị rõ ràng."
        ]
        tips = ["Nếu đề bài không theo mẫu rút về đơn vị, hãy viết lại dữ kiện theo dạng 'tổng' và 'số lượng' trước."]

    return {
        "method": detected,
        "steps": steps,
        "tips": tips
    }

# =====================================================
# CONTEXT / CACHE-BUSTING
# =====================================================
@app.context_processor
def inject_time():
    return dict(time=time)

@app.after_request
def add_header(response):
    response.cache_control.no_store = True
    response.cache_control.no_cache = True
    response.headers['Pragma'] = 'no-cache'
    return response

# =====================================================
# ROUTES: PUBLIC
# =====================================================
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/student')
def student():
    problem = PROBLEMS[0] if PROBLEMS else {"text": "Chưa có bài tập."}
    return render_template('student.html', problem=problem["text"])

# =====================================================
# AUTH (GIÁO VIÊN)
# =====================================================
@app.route('/teacher-login', methods=['GET', 'POST'])
def teacher_login():
    if request.method == 'POST':
        email = request.form.get('email', '').strip()
        password = request.form.get('password', '').strip()
        if email == TEACHER_EMAIL and password == TEACHER_PASSWORD:
            session['is_teacher'] = True
            flash("✅ Đăng nhập thành công!", "success")
            return redirect(url_for('teacher_dashboard'))
        else:
            flash("❌ Sai email hoặc mật khẩu!", "error")
    return render_template('teacher_login.html')

@app.route('/teacher-logout')
def teacher_logout():
    session.pop('is_teacher', None)
    flash("🔒 Đã đăng xuất!", "info")
    return redirect(url_for('index'))

# =====================================================
# TEACHER DASHBOARD
# =====================================================
@app.route('/teacher')
def teacher_dashboard():
    return render_template(
        'teacher.html',
        problems=PROBLEMS,
        submissions=STUDENT_SUBMISSIONS,
        is_teacher=session.get('is_teacher', False)
    )

# =====================================================
# WRITE ACTIONS (GIÁO VIÊN)
# =====================================================
@app.route('/save-teacher', methods=['POST'])
def save_teacher():
    if not session.get('is_teacher'):
        return jsonify({"status": "error", "message": "Bạn không có quyền lưu bài tập!"}), 403
    try:
        data = request.get_json(force=True) or {}
        problem = (data.get("problem") or "").strip()
        model_answer = (data.get("model_answer") or "").strip()
        correct_value_raw = (data.get("correct_value") or "").strip()

        if not problem or not model_answer:
            return jsonify({"status": "error", "message": "Thiếu [BÀI TOÁN] hoặc [ĐÁP ÁN MẪU]."}), 400

        try:
            correct_value = int(correct_value_raw)
        except Exception:
            correct_value = 0

        new_id = max([p["id"] for p in PROBLEMS], default=0) + 1
        PROBLEMS.append({
            "id": new_id,
            "text": problem,
            "model_answer": model_answer,
            "correct_value": correct_value,
            "topic": "Thủ công"
        })
        return jsonify({"status": "success", "message": f"Đã lưu bài tập mới (ID {new_id}). Tổng: {len(PROBLEMS)} bài."})
    except Exception as e:
        return jsonify({"status": "error", "message": f"Lỗi xử lý dữ liệu: {str(e)}"}), 500

@app.route('/upload-problem', methods=['POST'])
def upload_problem():
    if not session.get('is_teacher'):
        return jsonify({"status": "error", "message": "Bạn chưa đăng nhập!"}), 403

    if 'file' not in request.files:
        return jsonify({"status": "error", "message": "Chưa chọn file"}), 400

    file = request.files['file']
    filename = (file.filename or '').strip()
    if not filename:
        return jsonify({"status": "error", "message": "File rỗng"}), 400
    if not allowed_file_extension(filename):
        return jsonify({"status": "error", "message": "Chỉ hỗ trợ .docx và .pdf"}), 400

    safe_name = secure_filename(filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], safe_name)
    file.save(filepath)

    try:
        content = extract_text_from_docx(filepath) if safe_name.endswith('.docx') else extract_text_from_pdf(filepath)
        problem, model_answer, correct_value = parse_problem_file(content)
        if not problem:
            return jsonify({"status": "error", "message": "Không tìm thấy [BÀI TOÁN] trong file"}), 400

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

# =====================================================
# AI HƯỚNG DẪN TỪNG BƯỚC (HỌC SINH GỌI)
# =====================================================
@app.route('/ai-guide', methods=['POST'])
def ai_guide():
    """
    Body JSON: { "problem": "...", "answer": "..."(optional) }
    Trả về: { method, steps:[], tips:[] }
    """
    try:
        data = request.get_json(force=True) or {}
        problem = (data.get("problem") or "").strip()
        if not problem and PROBLEMS:
            problem = PROBLEMS[0]["text"]
        correct = PROBLEMS[0]["correct_value"] if PROBLEMS else None

        guide = ai_step_by_step(problem, correct_value=correct)
        return jsonify(guide)
    except Exception as e:
        return jsonify({"error": f"Lỗi tạo hướng dẫn: {str(e)}"}), 500

# =====================================================
# STUDENT SUBMIT
# =====================================================
@app.route('/submit', methods=['POST'])
def submit():
    global SUBMISSION_SEQ
    try:
        data = request.get_json(force=True) or {}
        answer = (data.get("answer") or "").strip()
        start_time = float(data.get("start_time", time.time()))

        if not answer:
            return jsonify({"error": "Lời giải không được để trống"}), 400

        current = PROBLEMS[0] if PROBLEMS else {"model_answer": "", "correct_value": 0}
        ai_result = evaluate_solution_v2(answer, current["model_answer"], current["correct_value"])

        duration_sec = time.time() - start_time
        minutes = int(duration_sec // 60)
        seconds = int(duration_sec % 60)
        ai_result["duration"] = f"{minutes} phút {seconds} giây"

        submission = {
            "id": SUBMISSION_SEQ,
            "answer": answer,
            "result": ai_result,         # kết quả AI
            "final_result": ai_result,   # mặc định final = AI (có thể bị ghi đè)
            "manual_override": False,
            "time_submitted": time.strftime("%H:%M - %d/%m/%Y", time.localtime())
        }
        STUDENT_SUBMISSIONS.append(submission)
        SUBMISSION_SEQ += 1

        return jsonify(ai_result | {"submission_id": submission["id"]})
    except Exception as e:
        return jsonify({"error": f"Có lỗi khi xử lý bài làm: {str(e)}"}), 500

# =====================================================
# TEACHER MANUAL GRADING (GHI ĐÈ)
# =====================================================
@app.route('/grade-manual', methods=['POST'])
def grade_manual():
    """
    Body JSON: {
      "submission_id": 3,
      "score": 8,
      "level": "Hoàn thành tốt",
      "level_class": "good",
      "feedback": "Giá 1 bao đúng, diễn đạt rõ."
    }
    """
    if not session.get('is_teacher'):
        return jsonify({"status": "error", "message": "Bạn chưa đăng nhập!"}), 403
    try:
        data = request.get_json(force=True) or {}
        sid = int(data.get("submission_id"))
        score = int(data.get("score"))
        level = (data.get("level") or "").strip()
        level_class = (data.get("level_class") or "").strip()
        feedback = (data.get("feedback") or "").strip()

        # tìm submission
        sub = next((s for s in STUDENT_SUBMISSIONS if s["id"] == sid), None)
        if not sub:
            return jsonify({"status": "error", "message": "Không tìm thấy bài nộp."}), 404

        final = {
            "score": score,
            "level": level,
            "level_class": level_class,
            "feedback": feedback,
            "duration": sub["result"].get("duration", "")
        }
        sub["final_result"] = final
        sub["manual_override"] = True

        return jsonify({"status": "success", "message": f"✅ Đã chấm thủ công bài #{sid}.", "final_result": final})
    except Exception as e:
        return jsonify({"status": "error", "message": f"Lỗi chấm thủ công: {str(e)}"}), 500

# =====================================================
# MAIN
# =====================================================
if __name__ == '__main__':
    # Cho phép Render đặt PORT; nếu local trùng cổng, đổi PORT env để chạy song song
    port = int(os.environ.get("PORT", 5001))
    app.run(host='0.0.0.0', port=port)
