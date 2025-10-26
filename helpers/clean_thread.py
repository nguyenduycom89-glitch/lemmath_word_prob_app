import threading, time, os

def cleanup_old_submissions(submissions):
    """Tự động xóa bài cũ sau 72h"""
    while True:
        now = time.time()
        keep = []
        for s in submissions:
            try:
                created = time.mktime(time.strptime(s["time_submitted_raw"], "%Y-%m-%d %H:%M:%S"))
                if now - created <= 72*3600:
                    keep.append(s)
                else:
                    if s.get("image_path") and os.path.exists(s["image_path"]):
                        os.remove(s["image_path"])
            except:  # bỏ qua lỗi thời gian
                keep.append(s)
        submissions[:] = keep
        time.sleep(3600)  # kiểm tra mỗi giờ

def start_cleanup_thread(submissions):
    """Khởi động luồng nền dọn dẹp"""
    t = threading.Thread(target=cleanup_old_submissions, args=(submissions,), daemon=True)
    t.start()
