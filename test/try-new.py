import json, os, subprocess, sys, tempfile

ROOT = r"D:/eepseek Harness Workspace/teacher-kit"
CLI = os.path.join(ROOT, "engine", "cli.py")
TMP = tempfile.mkdtemp(prefix="tk15-")
os.environ["TEACHER_KIT_PROFILE"] = os.path.join(TMP, "profile.json")

def call(command, payload):
    proc = subprocess.run([sys.executable, CLI, command, "--args", json.dumps(payload, ensure_ascii=False)],
                          capture_output=True, text=True, encoding="utf-8")
    out = json.loads(proc.stdout.strip().splitlines()[-1])
    return proc.returncode, out

# ① 画像：先记住这所学校
rc, out = call("profile:set", {"patch": {
    "school": {"name": "临河区第三中学", "stage": "初中", "subject": "英语", "grade": "七年级",
               "teacher": "王老师", "classNames": "七年(3)班"},
    "paper": {"size": "8K-2", "footer": "第X页 共Y页　使用班级：{班级}　出卷老师：{老师}"},
    "plan": {"layout": "表格式"},
    "wechat": {"tone": "亲切", "signature": "临河三中 英语组"},
    "docflow": {"unit": "临河区第三中学", "dept": "教务处"},
}, "note": "第一次用，先把学校情况记下", "learn": True})
print("profile:set", rc, out["ok"], out["data"]["filled"], "saved=", out["data"].get("saved"))
print(out["data"]["text"][:400])

# ② 公众号：一篇家长会通知
rc, out = call("wechat:render", {"title": "本周五家长会，聊聊孩子们的英语朗读",
    "digest": "七年级英语组这学期的第一次家长会，重点说三件事。",
    "blocks": [
        {"type": "lead", "text": "各位家长：这周五下午四点，我们想和您聊聊孩子这学期怎么读英语。"},
        {"type": "heading", "text": "一、时间地点"},
        {"type": "text", "text": "本周五 16:00—17:30，七年(3)班教室。请您提前十分钟到。"},
        {"type": "list", "items": ["带上孩子的朗读本", "手机调静音", "有事的家长可以会后单独留一下"]},
        {"type": "heading", "text": "二、聊什么"},
        {"type": "text", "text": "一是这学期朗读怎么练，二是作业量会不会太重，三是家校怎么配合。"},
        {"type": "note", "text": "如果您来不了，微信上和我说一声，我把要点单独发您。"},
    ],
    "author": "七年级英语组",
    "options": {"theme": "school", "outputPath": os.path.join(TMP, "wechat.html")}})
print("wechat:render", rc, out["ok"], json.dumps(out["data"].get("stats"), ensure_ascii=False)[:200])
print("warnings:", [w["code"] for w in out["data"].get("warnings", [])])
print("paths:", out["data"].get("htmlPath"), out["data"].get("outputPath"))

# ③ 公文：教育局关于开展校园安全排查的通知
rc, out = call("docflow:build", {"title": "关于开展2026年秋季校园安全隐患排查整治工作的通知",
    "from": "临河区教育局", "docNo": "临教发〔2026〕41号", "receivedAt": "2026-10-24",
    "copies": "1", "secret": "内部", "urgency": "急件", "dept": "总务处",
    "deadline": "2026-12-10", "handler": "李主任", "registryNo": "临三中收〔2026〕87",
    "unit": "临河区第三中学",
    "summary": "区教育局要求各校在 12 月 10 日前完成校舍、消防、食品、交通四项安全隐患排查，并报送整改台账。",
    "items": ["成立由校长任组长的排查小组，明确分工", "按四项内容逐项自查，建立问题清单", "12 月 10 日前将整改台账报教育局安全科"],
    "opinion": "拟请教务处牵头，总务处、德育处配合，11 月 30 日前完成自查，12 月 5 日前形成台账报教育局安全科。",
    "steps": [
        {"name": "登记", "dept": "办公室", "owner": "李主任", "due": "10-25"},
        {"name": "拟办", "dept": "教务处", "owner": "张主任", "due": "10-27"},
        {"name": "批示", "dept": "校长室", "owner": "王校长", "due": "10-28"},
        {"name": "承办", "dept": "总务处", "owner": "刘老师", "due": "11-30"},
        {"name": "报送归档", "dept": "办公室", "owner": "李主任", "due": "12-05"}],
    "reply": {"to": "临河区教育局：", "blocks": [
        "贵局《关于开展2026年秋季校园安全隐患排查整治工作的通知》（临教发〔2026〕41号）收悉。",
        "我校已成立排查小组，对校舍、消防、食品、交通四项内容逐项自查，共排查出隐患 7 处，其中 5 处已即时整改，2 处（教学楼三楼护栏、食堂排烟管道）已列入维修计划，11 月 30 日前完成。",
        "现将整改台账报上，请审阅。"], "signature": "临河区第三中学", "date": "2026年12月4日"},
    "options": {"redHead": False, "sealNote": True},
    "outputPath": os.path.join(TMP, "docflow.docx")})
print("docflow:build", rc, out["ok"], out["data"].get("sections"), out["data"].get("bytes"))
print("warnings:", [(w["code"], w["level"]) for w in out["data"].get("warnings", [])])
print("TMP=", TMP)
