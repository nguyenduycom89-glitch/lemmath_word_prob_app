import re

def evaluate_solution_v2(answer, model_answer, correct_value):
    """Chấm điểm theo logic cơ bản (từ khóa & giá trị đúng)"""
    score, feedback, level, level_class = 0, "", "", ""
    normalized = answer.lower().replace(".", "")
    if str(int(correct_value)) in normalized:
        score, feedback = 10, "✅ Kết quả chính xác."
    elif re.search(r"\d", normalized):
        score, feedback = 7, "🟡 Có phép tính nhưng sai kết quả."
    else:
        score, feedback = 5, "⚪ Chưa có phép tính cụ thể."
    if score >= 9:
        level, level_class = "Hoàn thành tốt", "success"
    elif score >= 7:
        level, level_class = "Hoàn thành", "warning"
    else:
        level, level_class = "Chưa hoàn thành", "danger"
    return {
        "score": score,
        "level": level,
        "level_class": level_class,
        "feedback": feedback
    }

def ai_step_by_step(problem, correct_value):
    """Sinh hướng dẫn từng bước (mẫu cố định, gọn gàng)"""
    return (
        f"1️⃣ Đọc kĩ đề bài: {problem}\n"
        "2️⃣ Tìm giá trị của 1 đơn vị.\n"
        "3️⃣ Nhân số đơn vị để tìm đáp số cuối cùng.\n"
        f"✅ Đáp số đúng: {correct_value:,} đồng."
    )
