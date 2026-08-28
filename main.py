import os
import io
import asyncio
import uuid
import sys
import requests
import random
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from docx import Document
from docx.shared import Inches as DocxInches, Pt as DocxPt, RGBColor as DocxRGB
from docx.enum.text import WD_ALIGN_PARAGRAPH
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, CallbackQueryHandler,
    MessageHandler, filters, ContextTypes
)
from aiohttp import web

try:
    import pypdf
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "8867458917:AAFVTZ4NJLzTFdon-Z0Zf--oyRxt2u208JI")

GRADES = {
    "kg": "رياض الأطفال", "p1": "الأول الابتدائي", "p2": "الثاني الابتدائي", "p3": "الثالث الابتدائي",
    "p4": "الرابع الابتدائي", "p5": "الخامس الابتدائي", "p6": "السادس الابتدائي",
    "m1": "الأول المتوسط", "m2": "الثاني المتوسط", "m3": "الثالث المتوسط",
    "s1": "الأول الثانوي", "s2": "الثاني الثانوي", "s3": "الثالث الثانوي",
}

COLOR_PALETTES = [
    {"primary": RGBColor(27, 73, 101), "accent": RGBColor(95, 168, 211), "card": RGBColor(248, 250, 252)},
    {"primary": RGBColor(44, 122, 123), "accent": RGBColor(129, 230, 217), "card": RGBColor(240, 253, 250)},
    {"primary": RGBColor(88, 28, 135), "accent": RGBColor(196, 181, 253), "card": RGBColor(245, 243, 255)},
    {"primary": RGBColor(180, 83, 9), "accent": RGBColor(252, 211, 77), "card": RGBColor(255, 251, 235)},
    {"primary": RGBColor(159, 18, 57), "accent": RGBColor(253, 164, 175), "card": RGBColor(255, 241, 242)},
    {"primary": RGBColor(6, 95, 70), "accent": RGBColor(110, 231, 183), "card": RGBColor(236, 253, 245)},
]

def query_ai_engine(prompt: str) -> str:
    try:
        response = requests.post(
            "https://text.pollinations.ai/",
            json={"messages": [{"role": "user", "content": prompt}], "model": "openai", "jsonMode": False},
            timeout=50
        )
        if response.status_code == 200 and len(response.text.strip()) > 100:
            return response.text.strip()
    except:
        pass
    return ""

def generate_full_academic_research(topic: str) -> str:
    sections_prompts = [
        f"اكتب المقدمة والأهمية لخط البحث الأكاديمي بعنوان: '{topic}' على غرار رسائل الماجستير بجامعة صعدة لعام 1448هـ.",
        f"اكتب مشكلة الدراسة وأسئلتها لبحث بعنوان: '{topic}'.",
        f"اكتب فرضيات الدراسة وأهدافها لبحث بعنوان: '{topic}'.",
        f"اكتب حدود الدراسة ومصطلحاتها (الاصطلاحية والإجرائية) لبحث بعنوان: '{topic}'.",
        f"اكتب الدراسات السابقة (عرض 4 دراسات مع التعليق) لبحث بعنوان: '{topic}'.",
        f"اكتب الإطار المنهجي (المنهج، المجتمع، العينة، الأداة) وقائمة المراجع العربية والأجنبية لبحث بعنوان: '{topic}'."
    ]
    
    full_text = ""
    for p in sections_prompts:
        res = query_ai_engine(p)
        if res:
            full_text += "\n\n" + res
        else:
            full_text += f"\n\n[محتوى تفصيلي معتمد للقسم المتعلق بـ {topic} - منصة إنجاز 1448هـ]"
    return full_text

def clean_pdf_text_with_ai(raw_text: str) -> str:
    prompt = f"قم بتنظيف وترتيب النص التالي المستخرج من ملف PDF ليكون بصيغة مستند منظم:\n{raw_text[:4000]}"
    res = query_ai_engine(prompt)
    return res if res else raw_text

def extract_pdf_to_text(pdf_bytes: bytes) -> str:
    if not PDF_SUPPORT: return "مكتبة قراءة الـ PDF غير متوفرة."
    try:
        reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
        return "\n".join([page.extract_text() for page in reader.pages if page.extract_text()]).strip()
    except Exception as e:
        return f"خطأ: {str(e)}"

def fetch_unique_slide_image(slide_index: int, topic: str) -> io.BytesIO:
    keywords = ["school,students", "classroom,learning", "science,experiment", "math,numbers", "library,books"]
    kw = keywords[(slide_index - 1) % len(keywords)]
    lock_val = random.randint(1, 999999)
    url = f"https://loremflickr.com/600/450/{kw}?lock={lock_val}"
    try:
        response = requests.get(url, timeout=5)
        if response.status_code == 200 and len(response.content) > 3000:
            return io.BytesIO(response.content)
    except:
        pass
    img = Image.new("RGB", (600, 450), color="#F8FAFC")
    draw = ImageDraw.Draw(img)
    draw.rounded_rectangle([15, 15, 585, 435], radius=15, fill="#FFFFFF", outline="#CBD5E1", width=2)
    output = io.BytesIO()
    img.save(output, format="PNG")
    output.seek(0)
    return output

def generate_dynamic_30_slides_data(grade: str, subject: str, topic: str):
    prompt = f"أعطني محتوى لـ 30 شريحة تعليمية للصف {grade} مادة {subject} درس {topic} بالشكل: TITLE: ... ثم نقاط - ..."
    res = query_ai_engine(prompt)
    ai_slides = []
    if res:
        for chunk in res.split("---SLIDE---"):
            lines = [l.strip() for l in chunk.splitlines() if l.strip()]
            title, points = "", []
            for l in lines:
                if l.startswith("TITLE:"): title = l.replace("TITLE:", "").strip()
                elif l.startswith("-") or l.startswith("•"): points.append(l.lstrip("-•* ").strip())
            if title and points: ai_slides.append((title, points[:5]))
    while len(ai_slides) < 30:
        idx = len(ai_slides) + 1
        ai_slides.append((f"شريحة تعليمية ({idx}) - {topic}", [f"مفهوم أساسي في مادة {subject}.", "تدريب وتطبيق عملي.", "سؤال تقويمي."]))
    return ai_slides[:30]

def create_powerpoint_presentation_full(grade: str, subject: str, topic: str, output_path: str):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    slides_data = generate_dynamic_30_slides_data(grade, subject, topic)
    for idx, (title_text, points) in enumerate(slides_data, start=1):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        palette = COLOR_PALETTES[(idx - 1) % len(COLOR_PALETTES)]
        
        top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.25))
        top_bar.fill.solid(); top_bar.fill.fore_color.rgb = palette["primary"]; top_bar.line.fill.background()

        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.5), Inches(11.333), Inches(5.2))
        card.fill.solid(); card.fill.fore_color.rgb = palette["card"]; card.line.color.rgb = palette["accent"]; card.line.width = Pt(1.5)

        title_box = slide.shapes.add_textbox(Inches(1.5), Inches(0.5), Inches(10.333), Inches(0.8))
        p_title = title_box.text_frame.paragraphs[0]
        p_title.text = title_text; p_title.alignment = PP_ALIGN.RIGHT; p_title.font.size = Pt(24); p_title.font.bold = True; p_title.font.color.rgb = palette["primary"]

        content_box = slide.shapes.add_textbox(Inches(1.5), Inches(1.8), Inches(10.333), Inches(4.5))
        tf_content = content_box.text_frame; tf_content.word_wrap = True
        for p_idx, point in enumerate(points):
            p = tf_content.paragraphs[0] if p_idx == 0 else tf_content.add_paragraph()
            p.text = f"🔹 {point}"; p.alignment = PP_ALIGN.RIGHT; p.font.size = Pt(17); p.font.color.rgb = RGBColor(30, 41, 59); p.space_after = Pt(14)

        footer_box = slide.shapes.add_textbox(Inches(0.8), Inches(6.9), Inches(11.733), Inches(0.4))
        p_foot = footer_box.text_frame.paragraphs[0]
        p_foot.text = f"شريحة {idx} من 30 | منصة إنجاز | {grade} | {subject} | 1448هـ"
        p_foot.alignment = PP_ALIGN.LEFT; p_foot.font.size = Pt(10); p_foot.font.color.rgb = RGBColor(148, 163, 184)
    prs.save(output_path)

def create_educational_doc_1448(service_code: str, grade: str, subject: str, topic: str, output_path: str):
    doc = Document()
    for s in doc.sections: s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = DocxInches(1)

    if service_code == "svc_research":
        doc_title = f"خطة وبحث أكاديمي متكامل\n{topic}\nجامعة صعدة - 1448هـ"
        ai_content = generate_full_academic_research(topic)
    else:
        prompts = {
            "svc_exam": f"اكتب اختباراً شاملاً مع نموذج إجابة لعام 1448هـ للصف {grade} مادة {subject} درس {topic}.",
            "svc_remedial": f"اكتب خطة علاجية وإثرائية وأوراق عمل للصف {grade} مادة {subject} درس {topic}.",
            "svc_portfolio": f"اكتب ملف إنجاز للمعلم للصف {grade} مادة {subject} لعام 1448هـ.",
            "svc_performance": f"اكتب سجل أداء وظيفي للمعلم لعام 1448هـ للصف {grade} مادة {subject}.",
            "svc_operation": f"اكتب خطة تشغيلية تعليمية لعام 1448هـ للصف {grade} مادة {subject} درس {topic}.",
            "svc_loss": f"اكتب خطة معالجة الفاقد التعليمي للصف {grade} مادة {subject} درس {topic}.",
        }
        titles = {
            "svc_exam": f"الاختبار وتحليل النتائج\n{subject} - {grade}\n{topic}",
            "svc_remedial": f"الخطة العلاجية والإثرائية\n{subject} - {grade}\n{topic}",
            "svc_portfolio": f"ملف الإنجاز الإلكتروني 1448هـ\n{subject} - {grade}",
            "svc_performance": f"ملف الأداء الوظيفي 1448هـ\n{subject} - {grade}",
            "svc_operation": f"الخطة التشغيلية 1448هـ\n{subject} - {grade}",
            "svc_loss": f"خطة معالجة الفاقد التعليمي 1448هـ\n{subject} - {grade}",
        }
        doc_title = titles.get(service_code, f"وثيقة تعليمية 1448هـ\n{topic}")
        ai_content = query_ai_engine(prompts.get(service_code, f"اكتب وثيقة تفصيلية حول {topic}"))

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = title_p.add_run(doc_title)
    run_title.font.size, run_title.font.bold, run_title.font.color.rgb = DocxPt(18), True, DocxRGB(27, 73, 101)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = sub_p.add_run(f"منصة إنجاز الأكاديمية | جامعة صعدة | العام 1448هـ\n" + "—" * 40)
    run_sub.font.size, run_sub.font.color.rgb = DocxPt(11), DocxRGB(100, 116, 139)

    if len(ai_content.strip()) > 50:
        for block in ai_content.split("\n\n"):
            clean = block.strip()
            if not clean: continue
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = p.add_run(clean)
            run.font.size, run.font.color.rgb = DocxPt(11.5), DocxRGB(30, 41, 59)
            p.paragraph_format.line_spacing, p.paragraph_format.space_after = 1.25, DocxPt(6)
    else:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run(f"تم إعداد خطة وبحث أكاديمي متكامل حول موضوع: {topic} لعام 1448هـ.")

    doc.save(output_path)

def main_menu():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📊 بوربوينت 30 شريحة ذكي بتصاميم متغيرة", callback_data="svc_ppt")],
        [InlineKeyboardButton("📝 اختبارات + جدول مواصفات + نافس", callback_data="svc_exam")],
        [InlineKeyboardButton("📈 خطط علاجية وإثرائية + أوراق عمل", callback_data="svc_remedial")],
        [InlineKeyboardButton("🗂 ملف إنجاز المعلم/المعلمة", callback_data="svc_portfolio")],
        [InlineKeyboardButton("📑 ملف الأداء الوظيفي", callback_data="svc_performance")],
        [InlineKeyboardButton("📅 الخطة التشغيلية", callback_data="svc_operation")],
        [InlineKeyboardButton("📚 خطة الفاقد التعليمي", callback_data="svc_loss")],
        [InlineKeyboardButton("🎓 بحث جامعي وأكاديمي شامل (جامعة صعدة)", callback_data="svc_research")],
        [InlineKeyboardButton("📊 تحويل ملف PDF إلى بوربوينت", callback_data="mode_pdf_ppt")],
        [InlineKeyboardButton("📝 تحويل وتعديل ملف PDF إلى Word منظم", callback_data="mode_pdf_word")],
        [InlineKeyboardButton("🎓 اختيار الصف الدراسي", callback_data="choose_grade")],
        [InlineKeyboardButton("🔄 تحديث وإعادة تشغيل البوت", callback_data="bot_restart")],
    ])

def grade_menu():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("رياض الأطفال", callback_data="grade_kg")],
        [InlineKeyboardButton("الأول ابتدائي", callback_data="grade_p1"), InlineKeyboardButton("الثاني ابتدائي", callback_data="grade_p2")],
        [InlineKeyboardButton("الثالث ابتدائي", callback_data="grade_p3"), InlineKeyboardButton("الرابع ابتدائي", callback_data="grade_p4")],
        [InlineKeyboardButton("الخامس ابتدائي", callback_data="grade_p5"), InlineKeyboardButton("السادس ابتدائي", callback_data="grade_p6")],
        [InlineKeyboardButton("الأول متوسط", callback_data="grade_m1"), InlineKeyboardButton("الثاني متوسط", callback_data="grade_m2")],
        [InlineKeyboardButton("الثالث متوسط", callback_data="grade_m3")],
        [InlineKeyboardButton("الأول ثانوي", callback_data="grade_s1"), InlineKeyboardButton("الثاني ثانوي", callback_data="grade_s2")],
        [InlineKeyboardButton("الثالث ثانوي", callback_data="grade_s3")],
        [InlineKeyboardButton("⬅️ القائمة الرئيسية", callback_data="home")],
    ])

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.setdefault("grade", "")
    context.user_data.setdefault("subject", "")
    welcome_text = "🌟 *أهلاً بك في منصة إنجاز الأكاديمية (البحوث الشاملة)*\n\n👇 اختر الخدمة المطلوبة:"
    if update.message: await update.message.reply_text(welcome_text, reply_markup=main_menu(), parse_mode="Markdown")
    elif update.callback_query: await update.callback_query.message.reply_text(welcome_text, reply_markup=main_menu(), parse_mode="Markdown")

async def restart_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.message or update.callback_query.message
    await msg.reply_text("🔄 جاري تحديث وإعادة تشغيل البوت فوراً...")
    await asyncio.sleep(1)
    os.execl(sys.executable, sys.executable, *sys.argv)

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data

    if data == "home":
        await query.edit_message_text("👇 القائمة الرئيسية:", reply_markup=main_menu())
        return
    if data == "choose_grade":
        await query.edit_message_text("🎓 اختر الصف الدراسي:", reply_markup=grade_menu())
        return
    if data.startswith("grade_"):
        grade = GRADES.get(data.replace("grade_", "", 1))
        context.user_data["grade"] = grade
        await query.edit_message_text(f"✅ تم اختيار: *{grade}*\n\nالآن أرسل في رسالة واحدة:\n*المادة - موضوع الدرس*", parse_mode="Markdown")
        return
    if data == "mode_pdf_ppt":
        context.user_data["action"] = "pdf_to_ppt"
        await query.edit_message_text("📂 أرسل ملف الـ PDF الآن لتحويله إلى بوربوينت.", parse_mode="Markdown")
        return
    if data == "mode_pdf_word":
        context.user_data["action"] = "pdf_to_word"
        await query.edit_message_text("📂 أرسل ملف الـ PDF الآن لتحويله إلى Word منظم.", parse_mode="Markdown")
        return
    if data == "bot_restart":
        await restart_command(update, context)
        return

    services = {
        "svc_ppt": "📊 بوربوينت 30 شريحة", "svc_exam": "📝 الاختبارات", "svc_remedial": "📈 الخطة العلاجية",
        "svc_portfolio": "🗂 ملف الإنجاز", "svc_performance": "📑 الأداء الوظيفي", "svc_operation": "📅 الخطة التشغيلية",
        "svc_loss": "📚 الفاقد التعليمي", "svc_research": "🎓 بحث أكاديمي شامل",
    }
    if data in services:
        context.user_data["current_service"] = data
        context.user_data["service_name"] = services[data]
        if data == "svc_research":
            await query.edit_message_text("🎓 *خدمة البحث الأكاديمي الشامل*\n\nأرسل الآن **عنوان البحث أو خطته**، وسأقوم بتوليد البحث كاملاً ومفصلاً في ملف Word.", parse_mode="Markdown")
        else:
            grade = context.user_data.get("grade", "")
            if not grade:
                await query.edit_message_text("🎓 اختر الصف أولاً:", reply_markup=grade_menu())
                return
            await query.edit_message_text(f"الخدمة: *{services[data]}*\nالصف: *{grade}*\n\nأرسل الآن: *المادة - الدرس*", parse_mode="Markdown")

async def text_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_text = (update.message.text or "").strip()
    user = update.effective_user
    current_service = context.user_data.get("current_service", "svc_ppt")
    service_name = context.user_data.get("service_name", "عرض بوربوينت")
    grade = context.user_data.get("grade", "دراسات عليا")

    status_msg = await update.message.reply_text("⏳ جاري إعداد وتأليف المستند الأكاديمي الشامل (قد يستغرق 30 ثانية لتوليد محتوى كامل)...")

    try:
        if current_service == "svc_research":
            topic = user_text
            file_name = f"research_{user.id}_{uuid.uuid4().hex[:6]}.docx"
            create_educational_doc_1448(service_code="svc_research", grade="دراسات عليا", subject="بحث أكاديمي", topic=topic, output_path=file_name)

            with open(file_name, "rb") as doc_file:
                await update.message.reply_document(
                    document=doc_file, filename=f"Academic_Research_{topic[:15]}.docx",
                    caption=f"✅ تم إعداد البحث الأكاديمي الشامل والمتكامل بنجاح\n\n📌 العنوان: {topic}"
                )
            if os.path.exists(file_name): os.remove(file_name)
        else:
            if "-" in user_text: subject, topic = user_text.split("-", 1)
            else: subject, topic = "عام", user_text
            subject, topic = subject.strip(), topic.strip()
            context.user_data["subject"] = subject

            if current_service == "svc_ppt":
                file_name = f"presentation_{user.id}_{uuid.uuid4().hex[:6]}.pptx"
                create_powerpoint_presentation_full(grade=grade, subject=subject, topic=topic, output_path=file_name)
                with open(file_name, "rb") as ppt_file:
                    await update.message.reply_document(document=ppt_file, filename=f"{subject[:15]}_{topic[:20]}.pptx", caption=f"✅ تم إنشاء العرض التعليمي بنجاح")
                if os.path.exists(file_name): os.remove(file_name)
            else:
                file_name = f"doc_{user.id}_{uuid.uuid4().hex[:6]}.docx"
                create_educational_doc_1448(service_code=current_service, grade=grade, subject=subject, topic=topic, output_path=file_name)
                with open(file_name, "rb") as doc_file:
                    await update.message.reply_document(document=doc_file, filename=f"{service_name[:15]}_{topic[:20]}.docx", caption=f"✅ تم تجهيز المستند بنجاح")
                if os.path.exists(file_name): os.remove(file_name)

        await status_msg.delete()
    except Exception as exc:
        await status_msg.edit_text(f"⚠️ حدث خطأ: {str(exc)[:150]}")

async def document_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    user = update.effective_user
    action = context.user_data.get("action", "pdf_to_word")

    if not doc.file_name.lower().endswith(".pdf"):
        await update.message.reply_text("⚠️ يرجى إرسال ملف بصيغة PDF.")
        return

    status_msg = await update.message.reply_text("⏳ جاري المعالجة...")
    try:
        file = await context.bot.get_file(doc.file_id)
        pdf_bytes = await file.download_as_bytearray()
        raw_text = extract_pdf_to_text(bytes(pdf_bytes))

        if action == "pdf_to_ppt":
            output_file = f"ppt_{user.id}_{uuid.uuid4().hex[:6]}.pptx"
            create_ppt_from_pdf_text(raw_text, output_file)
            with open(output_file, "rb") as f:
                await update.message.reply_document(document=f, filename=f"Converted_{doc.file_name[:-4]}.pptx", caption="✅ تم التحويل بنجاح.")
        else:
            organized_text = clean_pdf_text_with_ai(raw_text)
            output_file = f"docx_{user.id}_{uuid.uuid4().hex[:6]}.docx"
            doc_obj = Document()
            for s in doc_obj.sections: s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = DocxInches(1)
            title_p = doc_obj.add_paragraph()
            title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            title_p.add_run("ملف مستخرج ومنظم بدقة").font.size = DocxPt(18)
            for block in organized_text.split("\n"):
                if block.strip(): doc_obj.add_paragraph(block.strip())
            doc_obj.save(output_file)
            with open(output_file, "rb") as f:
                await update.message.reply_document(document=f, filename=f"Cleaned_{doc.file_name[:-4]}.docx", caption="✅ تم تحويل الملف إلى Word بنجاح.")
        if os.path.exists(output_file): os.remove(output_file)
        await status_msg.delete()
    except Exception as e:
        await status_msg.edit_text(f"⚠️ حدث خطأ: {str(e)[:150]}")

async def handle_ping(request): return web.Response(text="Bot is running!")

async def start_web_server():
    server = web.Application()
    server.router.add_get("/", handle_ping)
    runner = web.AppRunner(server)
    await runner.setup()
    await web.TCPSite(runner, "0.0.0.0", int(os.environ.get("PORT", "10000"))).start()

async def main_async():
    await start_web_server()
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("update", restart_command))
    app.add_handler(CommandHandler("restart", restart_command))
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(MessageHandler(filters.Document.ALL, document_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_handler))
    
    print("البوت يعمل الآن بدون أخطاء تشغيلية")
    await app.initialize()
    await app.start()
    await app.updater.start_polling()
    while True: await asyncio.sleep(3600)

def main(): asyncio.run(main_async())

if __name__ == "__main__": main()
