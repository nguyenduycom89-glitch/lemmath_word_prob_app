from flask import Flask, render_template, request, jsonify, redirect, url_for, flash, session
import re, os, time, random, requests
from werkzeug.utils import secure_filename
from docx import Document
import PyPDF2

# ==============================
# CẤU HÌNH ỨNG DỤNG
# ==============================
app = Flask(__name__, static_url_path='/static', static_folder='static', template_folder='templates')
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "secret-key-dev")

# ==============================
# CẤU HÌNH GIÁO VIÊN
# ==============================
TEACHER_EMAIL = os.environ.get("TEACHER_EMAIL", "nguyenduycom89@gmail.com")
TEACHER_PASSWORD = os.environ.get("TEACHER_PASSWORD", "nguyenmocgiao")

# ==============================
# DỮ LIỆU / UPLOAD
# ==============================
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10MB
ALLOWED_EXT = {'.docx', '.pdf'}

PROBLEMS = [
    {
        "id": 1,
        "text": "Một cửa hàng bán được 5 bao gạo với số tiền là 400000 đồng. "
                "Hỏi cửa hàng đó bán 8 bao gạo được bao nhiêu tiền?",
        "model_answer": "Giá 1 bao gạo là: 400000 : 5 = 80000 (đồng). "
                        "Giá 8 bao gạo là: 80000 × 8 = 640000 (đồng). Đáp số: 640000 đồng.",
        "correct_value": 640000,
        "topic": "Rút về đơn vị"
    }
]

STUDENT_SUBMISSIONS = []
SUBMISSION_SEQ = 1

# ==============================
# CẤU HÌNH GROQ
# ==============================
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "")
GROQ_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_MODEL = os.environ.get("GROQ_MODEL", "llama-3.3-70b-versatile")  # có thể đổi sang mixtral-8x7b-32768


# ==============================
# TIỆN ÍCH
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

def parse_problem_file(content: str):
    """Phân tích nội dung file có [BÀI TOÁN] / [ĐÁP ÁN MẪU] / [ĐÁP SỐ]"""
    problem, model, correct = "", "", 0
    if "[BÀI TOÁN]" in content:
        start = content.find("[BÀI TOÁN]") + len("[BÀI TOÁN]")
        end = content.find("[ĐÁP ÁN MẪU]") if "[ĐÁP ÁN MẪU]" in content else len(content)
        problem = content[start:end].strip()
    if "[ĐÁP ÁN MẪU]" in content:
        start = content.find("[ĐÁP ÁN MẪU]") + len("[ĐÁP ÁN MẪU]")
        end = content.find("[ĐÁP SỐ]") if "[ĐÁP SỐ]" in content else len(content)
        model = content[start:end].strip()
    if "[ĐÁP SỐ]" in content:
        m = re.search(r"\d+", content[content.find("[ĐÁP SỐ]"):])
        if m:
            correct = int(m.group())
    return problem, model, correct


# ==============================
# AI ĐÁNH GIÁ (rule-based)
# ==============================
def evaluate_solution_v2(student_text, model_answer, correct_value):
    text = student_text.lower()
    nums = [int(x) for x in re.findall(r'\d+', text)] if re.findall(r'\d+', text) else []
    score, fb = 0, []

    # Nhận biết dạng toán
    if any(kw in text for kw in ["1 bao", "mỗi bao", "1 đơn vị", "giá mỗi"]):
        score += 3
        fb.append("✅ Em nhận diện đúng dạng toán 'rút về đơn vị'.")
    else:
        fb.append("💡 Gợi ý: Em hãy tìm giá trị của 1 đơn vị trước.")

    # Vận dụng tính toán
    if correct_value in nums:
        score += 4
        fb.append("✅ Em tính đúng đáp số – rất tốt!")
    elif len(nums) >= 2:
        fb.append("✏️ Em cần chia tổng cho số lượng để tìm giá 1 đơn vị, rồi nhân với số cần hỏi.")

    # Trình bày
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
# AI HƯỚNG DẪN TỪNG BƯỚC
# ==============================
def ai_step_by_step(problem_text, correct_value=None):
    txt = (problem_text or "").lower()
    nums = [int(x) for x in re.findall(r'\d+', txt)]
    steps, tips = [], []
    detected = "unknown"

    if len(nums) >= 2 and ("bao" in txt or "mỗi" in txt or "1 đơn vị" in txt):
        detected = "rut_ve_don_vi"
        total, count = nums[0], nums[1]
        ask = nums[2] if len(nums) >= 3 else None

        unit = total // count if count else None
        steps.append(f"Bước 1 (Rút về đơn vị): {total} : {count} = {unit if unit else '...'} (giá 1 đơn vị).")
        if ask and unit:
            steps.append(f"Bước 2 (Nhân lên): {unit} × {ask} = {unit * ask} (giá {ask} đơn vị).")
        else:
            steps.append("Bước 2 (Nhân lên): Nhân giá trị 1 đơn vị với số đơn vị cần hỏi.")
        ans = correct_value or (unit * ask if unit and ask else None)
        steps.append(f"Bước 3 (Kết luận): Viết đáp số kèm đơn vị.{f' Đáp số: {ans}' if ans else ''}")

        tips = [
            "Nếu đề cho tổng và số lượng → chia để tìm 1 đơn vị.",
            "Sau đó nhân với số cần hỏi để ra kết quả.",
            "Luôn ghi rõ đơn vị (đồng, kg, lít...)."
        ]
    else:
        steps = [
            "Bước 1: Xác định 1 đơn vị bằng phép chia.",
            "Bước 2: Nhân giá 1 đơn vị với số cần hỏi.",
            "Bước 3: Ghi đáp số có đơn vị."
        ]
        tips = ["Nếu chưa rõ dạng toán, hãy tách dữ kiện 'tổng', 'số lượng', 'số cần hỏi' trước."]

    return {"method": detected, "steps": steps, "tips": tips}


# ==============================
# CONTEXT
# ==============================
@app.context_processor
def inject_time():
    return dict(time=time)


# ==============================
# ROUTES: PUBLIC
# ==============================
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/student')
def student():
    problem = PROBLEMS[0] if PROBLEMS else {"text": "Chưa có bài tập."}
    return render_template('student.html', problem=problem["text"])


# ==============================
# ROUTES: AUTH (GIÁO VIÊN)
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


# ==============================
# ROUTES: TEACHER DASHBOARD
# ==============================
@app.route('/teacher')
def teacher_dashboard():
    return render_template(
        'teacher.html',
        problems=PROBLEMS,
        submissions=STUDENT_SUBMISSIONS,
        is_teacher=session.get('is_teacher', False)
    )


# ==============================
# ROUTES: WRITE ACTIONS (CHỈ GIÁO VIÊN)
# ==============================
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

    new_id = max([p["id"] for p in PROBLEMS], default=0) + 1
    PROBLEMS.append({"id": new_id, "text": prob, "model_answer": model, "correct_value": corr, "topic": "Thủ công"})
    return jsonify({"status": "success", "message": f"Đã thêm bài mới (ID {new_id})."})


@app.route('/upload-problem', methods=['POST'])
def upload_problem():
    if not session.get('is_teacher'):
        return jsonify({"status": "error", "message": "Bạn chưa đăng nhập."}), 403

    if 'file' not in request.files:
        return jsonify({"status": "error", "message": "Chưa chọn file."}), 400

    file = request.files['file']
    filename = (file.filename or "").strip()
    if filename == "":
        return jsonify({"status": "error", "message": "File rỗng."}), 400

    if not allowed(filename):
        return jsonify({"status": "error", "message": "Chỉ hỗ trợ .docx và .pdf"}), 400

    safe_name = secure_filename(filename)
    path = os.path.join(app.config['UPLOAD_FOLDER'], safe_name)
    file.save(path)

    # Đọc nội dung
    try:
        if safe_name.endswith('.docx'):
            content = extract_text_from_docx(path)
        else:
            content = extract_text_from_pdf(path)
    except Exception as e:
        return jsonify({"status": "error", "message": f"Lỗi đọc file: {str(e)}"}), 500

    # Phân tích
    prob, model, corr = parse_problem_file(content)
    if not prob:
        return jsonify({"status": "error", "message": "Không tìm thấy [BÀI TOÁN] trong file."}), 400

    new_id = max([p["id"] for p in PROBLEMS], default=0) + 1
    PROBLEMS.append({"id": new_id, "text": prob, "model_answer": model, "correct_value": corr, "topic": "Tự động"})
    return jsonify({"status": "success", "message": f"Đã thêm bài tập mới! Tổng: {len(PROBLEMS)}"})


# ==============================
# ROUTES: STUDENT SUBMIT
# ==============================
@app.route('/submit', methods=['POST'])
def submit():
    global SUBMISSION_SEQ
    data = request.get_json(force=True)
    ans = (data.get("answer") or "").strip()
    start = float(data.get("start_time", time.time()))

    if not ans:
        return jsonify({"error": "Lời giải không được để trống."}), 400

    current = PROBLEMS[0] if PROBLEMS else {"text": "", "model_answer": "", "correct_value": 0}
    result = evaluate_solution_v2(ans, current["model_answer"], current["correct_value"])
    dur = int(time.time() - start)
    result["duration"] = f"{dur//60} phút {dur%60} giây"
    result["guide"] = ai_step_by_step(current["text"], current["correct_value"])

    STUDENT_SUBMISSIONS.append({
        "id": SUBMISSION_SEQ,
        "answer": ans,
        "result": result,          # kết quả AI
        "final_result": result,    # mặc định trùng AI; có thể bị override
        "manual_override": False,
        "time_submitted": time.strftime("%H:%M - %d/%m/%Y", time.localtime())
    })
    SUBMISSION_SEQ += 1
    return jsonify(result)


# ==============================
# ROUTES: TEACHER MANUAL GRADE
# ==============================
@app.route('/grade-manual', methods=['POST'])
def grade_manual():
    if not session.get('is_teacher'):
        return jsonify({"status": "error", "message": "Bạn chưa đăng nhập."}), 403

    data = request.get_json(force=True)
    try:
        sid = int(data.get("submission_id"))
    except Exception:
        return jsonify({"status": "error", "message": "submission_id không hợp lệ."}), 400

    sub = next((s for s in STUDENT_SUBMISSIONS if s["id"] == sid), None)
    if not sub:
        return jsonify({"status": "error", "message": "Không tìm thấy bài nộp."}), 404

    final = {
        "score": int(data.get("score")),
        "level": data.get("level"),
        "level_class": data.get("level_class"),
        "feedback": data.get("feedback"),
        "duration": sub["result"].get("duration", "")
    }
    sub["final_result"] = final
    sub["manual_override"] = True
    return jsonify({"status": "success", "message": f"✅ Đã chấm thủ công bài #{sid}."})


# ==============================
# ROUTES: NEXT PROBLEM (random)
# ==============================
@app.route('/api/next-problem')
def api_next_problem():
    if not PROBLEMS:
        return jsonify({"problem": "Chưa có bài tập."})
    return jsonify({"problem": random.choice(PROBLEMS)["text"]})


# ==============================
# ROUTE: AI TẠO BÀI TOÁN TƯƠNG TỰ (GROQ)
# ==============================
@app.route('/generate-ai', methods=['POST'])
def generate_ai():
    """
    Tạo bài toán tương tự bằng AI (dựa trên bài mẫu) — dùng Groq API (OpenAI-compatible endpoint).
    Cần biến môi trường: GROQ_API_KEY
    """
    try:
        data = request.get_json(force=True)
        base_problem = (data.get("problem") or "").strip()
        if not base_problem:
            return jsonify({"status": "error", "message": "Chưa nhập bài mẫu để tạo tương tự."}), 400

        if not GROQ_API_KEY:
            return jsonify({"status": "error", "message": "Thiếu GROQ_API_KEY trên môi trường."}), 500

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

        headers = {
            "Authorization": f"Bearer {GROQ_API_KEY}",
            "Content-Type": "application/json"
        }
        body = {
            "model": GROQ_MODEL,
            "messages": [
                {"role": "user", "content": prompt}
            ],
            "temperature": 0.8,
            "max_tokens": 400
        }

        resp = requests.post(GROQ_URL, headers=headers, json=body, timeout=60)
        if resp.status_code != 200:
            # Fallback khi hết quota / lỗi model
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

        data_json = resp.json()
        generated = data_json["choices"][0]["message"]["content"].strip()
        return jsonify({"status": "success", "generated": generated})

    except Exception as e:
        return jsonify({"status": "error", "message": f"Lỗi AI/Groq: {str(e)}"}), 500


# ==============================
# MAIN
# ==============================
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5001))
    app.run(host='0.0.0.0', port=port)
