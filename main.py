import os
import io
import asyncio
import re
import requests
import uuid
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

# ============================================================
# منصة إنجاز - بوت 1448هـ (النسخة المحسنة - مانع التكرار)
# ============================================================

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
if not TELEGRAM_TOKEN:
    raise ValueError("⚠️ التوكن غير موجود! يرجى إضافة TELEGRAM_TOKEN في متغيرات البيئة.")

GRADES = {
    "kg": "رياض الأطفال",
    "p1": "الأول الابتدائي",
    "p2": "الثاني الابتدائي",
    "p3": "الثالث الابتدائي",
    "p4": "الرابع الابتدائي",
    "p5": "الخامس الابتدائي",
    "p6": "السادس الابتدائي",
    "m1": "الأول المتوسط",
    "m2": "الثاني المتوسط",
    "m3": "الثالث المتوسط",
    "s1": "الأول الثانوي",
    "s2": "الثاني الثانوي",
    "s3": "الثالث الثانوي",
}

SYSTEM_PROMPT = """
أنت معلم متخصص ومصمم مناهج تعليمية خبير بالمناهج السعودية للمدارس.
مهمتك كتابة محتوى مدرسي للطلاب. 
تنبيه صارم: ممنوع منعاً باتاً استخدام لغة الشركات والمؤسسات (مثل: الأداء المؤسسي، العمليات التشغيلية، حجر الزاوية، التطوير الاستراتيجي).
اكتب محتوى علمي وأكاديمي (أمثلة، تمارين، شرح دروس حقيقية) يتناسب مع عمر الطالب.
"""

# ============================================================
# استدعاء الذكاء الاصطناعي مع كسر الكاش (Cache Busting)
# ============================================================
def query_ai_engine(prompt: str) -> str:
    # إضافة نص عشوائي لكسر ذاكرة التخزين المؤقت (Cache)
    unique_id = str(uuid.uuid4())
    cache_buster_prompt = f"{prompt}\n\n[System Note: Unique ID {unique_id} - Generate completely fresh and specific content.]"
    
    payloads = [
        {
            "url": f"https://text.pollinations.ai/?seed={random.randint(1, 999999)}",
            "data": {
                "messages": [
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": cache_buster_prompt},
                ],
                "model": "openai",
                "jsonMode": False
            },
        },
        {
            "url": "https://api.airforce/v1/chat/completions",
            "data": {
                "model": "gpt-4o-mini",
                "messages": [
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": cache_buster_prompt},
                ],
            },
        },
    ]

    for payload in payloads:
        try:
            response = requests.post(
                payload["url"],
                json=payload["data"],
                timeout=60,
            )
            if response.status_code != 200:
                continue

            text = response.text.strip()
            if not text:
                continue

            try:
                data = response.json()
                if "choices" in data:
                    return data["choices"][0]["message"]["content"].strip()
            except Exception:
                pass

            if len(text) > 100:
                return text

        except Exception:
            continue

    return ""


# ============================================================
# صور الشرائح (تم تحسين الكلمات لتكون تعليمية فقط)
# ============================================================
def fetch_unique_slide_image(slide_index: int, topic: str) -> io.BytesIO:
    keywords = [
        "school,students", "classroom,learning", "science,experiment", 
        "math,numbers", "library,books", "education,study", 
        "teacher,blackboard", "reading,desk"
    ]

    kw = keywords[(slide_index - 1) % len(keywords)]
    stable_seed = sum(ord(c) for c in topic) % 5000
    lock = (slide_index * 37 + stable_seed) % 5000

    url = f"https://loremflickr.com/600/450/{kw}?lock={lock}"

    try:
        response = requests.get(url, timeout=8)
        if response.status_code == 200 and len(response.content) > 3000:
            stream = io.BytesIO(response.content)
            stream.seek(0)
            return stream
    except Exception:
        pass

    # صورة بديلة مرسومة في حال فشل جلب الصورة
    img = Image.new("RGB", (600, 450), color="#F8FAFC")
    draw = ImageDraw.Draw(img)
    draw.rounded_rectangle([15, 15, 585, 435], radius=15, fill="#FFFFFF", outline="#CBD5E1", width=2)
    draw.rectangle([15, 15, 585, 70], fill="#1B4965")
    bars = [120, 180, 150, 230, 200]
    for idx, bar in enumerate(bars):
        x0 = 70 + idx * 90
        factor = ((slide_index + idx) % 5 + 6) / 10
        y0 = 400 - int(bar * factor)
        draw.rounded_rectangle([x0, y0, x0 + 55, 400], radius=6, fill="#5FA8D3" if idx % 2 == 0 else "#62B6CB")
    output = io.BytesIO()
    img.save(output, format="PNG")
    output.seek(0)
    return output


# ============================================================
# توليد 30 شريحة مخصصة للدرس
# ============================================================
def generate_dynamic_30_slides_data(grade: str, subject: str, topic: str):
    grade = (grade or "غير محدد").strip()
    subject = (subject or "غير محددة").strip()
    topic = (topic or "الدرس").strip()

    ai_prompt = f"""
أنت مدرس تصمم عرض PowerPoint مكون من 30 شريحة.
الصف: {grade}
المادة: {subject}
الموضوع/الدرس: {topic}

تنبيهات هامة جداً (يمنع مخالفتها):
1. هذا العرض موجه لطلاب مدرسة، تجنب تماماً العبارات الإدارية والمؤسسية (مثل: الأداء المؤسسي، العمليات التشغيلية، حجر الزاوية، التطوير الاستراتيجي).
2. لا تستخدم قوالب عامة. إذا كان الموضوع عاماً جداً (مثل: ترم كامل أو منهج كامل)، قم باختيار دروس فعلية محددة من المادة واشرحها بالأمثلة والقواعد.
3. أريد محتوى علمياً حقيقياً في الشرائح: معادلات، أمثلة وتطبيقات، نصوص، تجارب، وقواعد، وليس فقط ديباجات ومقدمات فارغة.

التنسيق الإلزامي لكل شريحة:
---SLIDE---
TITLE: عنوان خاص بالدرس أو الفقرة
CONTENT:
- معلومة دقيقة أو شرح فعلي للدرس
- مثال تطبيقي أو قاعدة علمية أو نصية
- سؤال أو نشاط للطلاب حول هذه الجزئية

كرر هذا التنسيق بالضبط لإنتاج 30 شريحة علمية مختلفة.
"""

    ai_raw = query_ai_engine(ai_prompt)
    
    # فحص المحتوى: إذا أنتج الذكاء الاصطناعي قالب الشركات الخاطئ، نتجاهله فوراً
    bad_keywords = ["الأداء المؤسسي", "العمليات التشغيلية", "البيئة المؤسسية", "حجر الزاوية"]
    if any(word in ai_raw for word in bad_keywords):
        ai_raw = "" # تجاهل النتيجة للانتقال للمحتوى المحلي الموثوق

    ai_slides = []

    if ai_raw:
        chunks = re.split(r"---\s*SLIDE\s*---", ai_raw, flags=re.IGNORECASE)
        for chunk in chunks:
            if not chunk.strip():
                continue
            title_match = re.search(r"TITLE\s*:\s*(.+)", chunk, flags=re.IGNORECASE)
            title = title_match.group(1).strip() if title_match else ""

            points = []
            for line in chunk.splitlines():
                line = line.strip()
                if re.match(r"^[-•*]\s+", line):
                    point = re.sub(r"^[-•*]\s+", "", line).strip()
                    if point and len(point) > 8:
                        points.append(point)

            if title and points:
                ai_slides.append((title, points[:5]))

    # المولد المحلي الآمن (يعمل إذا فشل الذكاء الاصطناعي في إعطاء محتوى تعليمي حقيقي)
    local_templates = [
        (f"تمهيد: أهمية {topic}", [f"موقف تمهيدي من الحياة يوضح أهمية {topic} لطلاب {grade}.", f"سؤال استثارة: ماذا تعرفون عن {topic}؟", "لنفكر معاً في حل للمشكلة المعروضة."]),
        (f"المفهوم الأساسي: {topic}", [f"تعريف دقيق ومبسط لمفهوم {topic}.", "المصطلحات العلمية المرتبطة بالدرس.", f"توضيح الفكرة الرئيسة في مادة {subject}."]),
        (f"قاعدة أو قانون: {topic}", [f"استعراض القاعدة الأساسية لفهم {topic}.", "خطوات تطبيق القاعدة بشكل متسلسل.", "مثال مباشر لتوضيح القاعدة."]),
        (f"مثال تطبيقي (1)", [f"تطبيق واقعي للمفهوم {topic}.", "طريقة الحل أو التفسير العلمي خطوة بخطوة.", "ما هي النتيجة المستخلصة؟"]),
        (f"مثال تطبيقي (2)", ["تطبيق آخر بأرقام أو معطيات مختلفة.", "تحليل المثال وتحديد الفرق عن المثال السابق.", "تدريب سريع للطلاب لحله في دقيقة."]),
        (f"تنبيه هام وأخطاء شائعة", [f"أبرز الأخطاء التي يقع فيها الطلاب عند دراسة {topic}.", "السبب العلمي أو المنطقي وراء هذا الخطأ.", "الطريقة الصحيحة لتجنبه والتأكد من الحل."]),
        (f"نشاط فردي: اختبر نفسك", [f"حل هذه المسألة/السؤال المتعلق بـ {topic}.", "دون خطوات الحل في الدفتر.", "سنقوم بمراجعة الإجابات سوياً."]),
        (f"نشاط ثنائي: فكر وشارك", [f"ناقش مع زميلك حلاً للمشكلة المرتبطة بـ {topic}.", "استخرجوا المعطيات وتوصلوا لنتيجة مشتركة.", "تبادلوا الأوراق لتصحيح الأخطاء."]),
        (f"ربط الدرس بالواقع", [f"أين نرى {topic} في حياتنا اليومية؟", "أهمية هذه المعرفة في المهن المستقبلية.", "تطبيق عملي نستخدمه باستمرار."]),
        (f"تدريب مهاري (أسئلة نافس)", [f"سؤال يعتمد على الفهم والاستنتاج في {topic}.", "استبعاد الخيارات الخاطئة للوصول للإجابة.", "تفسير سبب اختيار الإجابة الصحيحة."]),
        (f"تقويم ختامي للدرس", [f"تلخيص لأهم النقاط التي تعلمناها في {topic}.", "سؤال ختامي شامل لقياس مدى الفهم.", "واجب منزلي لتثبيت المعلومة."]),
    ]

    final_slides = []
    used_titles = set()

    # أخذ شرائح الذكاء الاصطناعي
    for title, points in ai_slides:
        normalized = re.sub(r"\s+", " ", title).strip().lower()
        if normalized in used_titles:
            continue
        final_slides.append((title, points))
        used_titles.add(normalized)
        if len(final_slides) >= 30:
            break

    # تكميل الشرائح من القوالب المحلية التعليمية إذا كان العدد أقل من 30
    for title, points in local_templates:
        normalized = re.sub(r"\s+", " ", title).strip().lower()
        if normalized in used_titles:
            continue
        final_slides.append((title, points))
        used_titles.add(normalized)
        if len(final_slides) >= 30:
            break

    # استكمال العدد حتى 30 شريحة بتمارين وتدريبات
    while len(final_slides) < 30:
        number = len(final_slides) + 1
        final_slides.append(
            (f"تمرين إضافي وتطبيق ({number})", [
                f"مسألة أو سؤال تطبيقي إضافي حول {topic}.",
                f"تحدي للطلاب المتفوقين في {subject}.",
                "مراجعة الحلول والمناقشة الجماعية.",
            ])
        )

    # ترقيم العناوين
    numbered = []
    for i, (title, points) in enumerate(final_slides[:30], start=1):
        clean_title = re.sub(r"^\s*\d+\s*[\.\-:)]\s*", "", title).strip()
        numbered.append((f"الشريحة {i}: {clean_title}", points))

    return numbered


# ============================================================
# إنشاء PowerPoint
# ============================================================
def create_powerpoint_presentation_full(grade: str, subject: str, topic: str, output_path: str):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slides_data = generate_dynamic_30_slides_data(grade, subject, topic)

    for idx, (title_text, points) in enumerate(slides_data, start=1):
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.2))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = RGBColor(27, 73, 101)
        top_bar.line.fill.background()

        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.4), Inches(1.5), Inches(7.2), Inches(5.1))
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(248, 250, 252)
        card.line.color.rgb = RGBColor(203, 213, 225)
        card.line.width = Pt(1.5)

        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(1.0))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True

        p_title = tf_title.paragraphs[0]
        p_title.text = title_text
        p_title.alignment = PP_ALIGN.RIGHT
        p_title.font.size = Pt(26)
        p_title.font.bold = True
        p_title.font.color.rgb = RGBColor(27, 73, 101)

        content_box = slide.shapes.add_textbox(Inches(5.6), Inches(1.7), Inches(6.8), Inches(4.7))
        tf_content = content_box.text_frame
        tf_content.word_wrap = True

        for p_idx, point in enumerate(points):
            p = tf_content.paragraphs[0] if p_idx == 0 else tf_content.add_paragraph()
            p.text = f"◀ {point}"
            p.alignment = PP_ALIGN.RIGHT
            p.font.size = Pt(18)
            p.font.color.rgb = RGBColor(30, 41, 59)
            p.space_after = Pt(14)

        img_stream = fetch_unique_slide_image(idx, topic)
        if img_stream:
            slide.shapes.add_picture(img_stream, Inches(0.8), Inches(1.5), Inches(4.3), Inches(5.1))

        footer_box = slide.shapes.add_textbox(Inches(0.8), Inches(6.9), Inches(11.733), Inches(0.4))
        p_foot = footer_box.text_frame.paragraphs[0]
        p_foot.text = f"شريحة {idx} من 30 | منصة إنجاز | {grade} | {subject} | {topic} | 1448هـ"
        p_foot.alignment = PP_ALIGN.LEFT
        p_foot.font.size = Pt(11)
        p_foot.font.color.rgb = RGBColor(148, 163, 184)

    prs.save(output_path)

# ============================================================
# مستندات Word
# ============================================================
def create_educational_doc_1448(service_code: str, grade: str, subject: str, topic: str, output_path: str):
    doc = Document()
    for section in doc.sections:
        section.top_margin, section.bottom_margin = DocxInches(1), DocxInches(1)
        section.left_margin, section.right_margin = DocxInches(1), DocxInches(1)

    prompts = {
        "svc_exam": f"اكتب اختباراً شاملاً لعام 1448هـ للصف {grade} مادة {subject} حول {topic}. مع جدول مواصفات ونموذج إجابة.",
        "svc_remedial": f"اكتب خطة علاجية وإثرائية وأوراق عمل للصف {grade} مادة {subject} حول {topic}.",
        "svc_portfolio": f"اكتب ملف إنجاز إلكترونياً للمعلم للصف {grade} ومادة {subject} لعام 1448هـ.",
        "svc_performance": f"اكتب سجل ملف أداء وظيفي منظم للمعلم لعام 1448هـ للصف {grade} ومادة {subject}.",
        "svc_operation": f"اكتب خطة تشغيلية تعليمية لعام 1448هـ للصف {grade} ومادة {subject} حول {topic}.",
        "svc_loss": f"اكتب خطة معالجة فاقد تعليمي للصف {grade} في مادة {subject} حول {topic}.",
        "svc_research": f"اكتب بحثاً أكاديمياً جامعياً مفصلاً حول {topic}.",
    }

    titles = {
        "svc_exam": f"الاختبار وتحليل النتائج\n{subject} - {grade}\n{topic}",
        "svc_remedial": f"الخطة العلاجية والإثرائية\n{subject} - {grade}\n{topic}",
        "svc_portfolio": f"ملف الإنجاز الإلكتروني 1448هـ\n{subject} - {grade}",
        "svc_performance": f"ملف الأداء الوظيفي 1448هـ\n{subject} - {grade}",
        "svc_operation": f"الخطة التشغيلية 1448هـ\n{subject} - {grade}",
        "svc_loss": f"خطة معالجة الفاقد التعليمي 1448هـ\n{subject} - {grade}",
        "svc_research": f"بحث أكاديمي متكامل\n{topic}",
    }

    prompt = prompts.get(service_code, f"اكتب وثيقة تعليمية لعام 1448هـ حول {topic}.")
    doc_title = titles.get(service_code, f"وثيقة تعليمية 1448هـ\n{topic}")
    ai_content = query_ai_engine(prompt)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = title_p.add_run(doc_title)
    run_title.font.size, run_title.font.bold, run_title.font.color.rgb = DocxPt(20), True, DocxRGB(27, 73, 101)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = sub_p.add_run(f"منصة إنجاز | الصف: {grade} | المادة: {subject} | العام 1448هـ\n" + "—" * 35)
    run_sub.font.size, run_sub.font.color.rgb = DocxPt(11), DocxRGB(100, 116, 139)

    if len(ai_content) > 300:
        for block in ai_content.split("\n\n"):
            clean = block.strip()
            if not clean: continue
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

            if clean.startswith("#") or any(clean.startswith(f"{i}.") for i in range(1, 20)) or "المحور" in clean or "الهدف" in clean:
                run = p.add_run(clean.replace("#", "").strip())
                run.font.size, run.font.bold, run.font.color.rgb = DocxPt(14), True, DocxRGB(27, 73, 101)
                p.paragraph_format.space_before, p.paragraph_format.space_after = DocxPt(12), DocxPt(4)
            else:
                run = p.add_run(clean)
                run.font.size, run.font.color.rgb = DocxPt(11.5), DocxRGB(30, 41, 59)
                p.paragraph_format.line_spacing, p.paragraph_format.space_after = 1.25, DocxPt(6)
    else:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run(f"تم إعداد المستند للصف {grade} في مادة {subject} حول {topic} لعام 1448هـ.")

    doc.save(output_path)


# ============================================================
# القوائم
# ============================================================
def main_menu():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📊 بوربوينت 30 شريحة + جميع الإضافات", callback_data="svc_ppt")],
        [InlineKeyboardButton("📝 اختبارات + جدول مواصفات + نافس", callback_data="svc_exam")],
        [InlineKeyboardButton("📈 خطط علاجية وإثرائية + أوراق عمل", callback_data="svc_remedial")],
        [InlineKeyboardButton("🗂 ملف إنجاز المعلم/المعلمة", callback_data="svc_portfolio")],
        [InlineKeyboardButton("📑 ملف الأداء الوظيفي", callback_data="svc_performance")],
        [InlineKeyboardButton("📅 الخطة التشغيلية", callback_data="svc_operation")],
        [InlineKeyboardButton("📚 خطة الفاقد التعليمي", callback_data="svc_loss")],
        [InlineKeyboardButton("🎓 بحث جامعي وأكاديمي Word", callback_data="svc_research")],
        [InlineKeyboardButton("🎓 اختيار الصف الدراسي", callback_data="choose_grade")],
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

# ============================================================
# الأوامر واستقبال النصوص
# ============================================================
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.setdefault("grade", "")
    context.user_data.setdefault("subject", "")
    welcome_text = (
        "🌟 *أهلاً بك في منصة إنجاز للخدمات التعليمية والأكاديمية 1448هـ*\n\n"
        "👇 اختر الخدمة من القائمة:"
    )
    if update.message:
        await update.message.reply_text(welcome_text, reply_markup=main_menu(), parse_mode="Markdown")
    elif update.callback_query:
        await update.callback_query.message.reply_text(welcome_text, reply_markup=main_menu(), parse_mode="Markdown")

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
        grade_key = data.replace("grade_", "", 1)
        grade = GRADES.get(grade_key)
        context.user_data["grade"] = grade
        await query.edit_message_text(f"✅ تم اختيار: *{grade}*\n\nالآن أرسل في رسالة واحدة:\n*المادة - موضوع الدرس*\nمثال: رياضيات - القسمة", parse_mode="Markdown")
        return

    services = {
        "svc_ppt": "📊 بوربوينت 30 شريحة", "svc_exam": "📝 الاختبارات", "svc_remedial": "📈 الخطة العلاجية",
        "svc_portfolio": "🗂 ملف الإنجاز", "svc_performance": "📑 الأداء الوظيفي", "svc_operation": "📅 الخطة التشغيلية",
        "svc_loss": "📚 الفاقد التعليمي", "svc_research": "🎓 البحث الأكاديمي",
    }

    if data in services:
        context.user_data["current_service"] = data
        context.user_data["service_name"] = services[data]
        grade = context.user_data.get("grade", "")

        if not grade:
            await query.edit_message_text("🎓 اختر الصف أولاً لتخصيص المحتوى للمرحلة:", reply_markup=grade_menu())
            return
        await query.edit_message_text(f"الخدمة: *{services[data]}*\nالصف: *{grade}*\n\nأرسل الآن: *المادة - الدرس*", parse_mode="Markdown")

async def text_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_text = (update.message.text or "").strip()
    user = update.effective_user
    current_service = context.user_data.get("current_service", "svc_ppt")
    service_name = context.user_data.get("service_name", "عرض بوربوينت")
    grade = context.user_data.get("grade", "")

    if not grade:
        for key, grade_name in GRADES.items():
            if grade_name in user_text:
                grade = grade_name
                context.user_data["grade"] = grade
                break
        if not grade:
            await update.message.reply_text("🎓 اختر الصف أولاً من الزر، ثم أرسل المادة والدرس.", reply_markup=grade_menu())
            return

    status_msg = await update.message.reply_text("⏳ جارٍ إعداد وتأليف المحتوى الأكاديمي، يرجى الانتظار...")

    try:
        if "-" in user_text:
            subject, topic = user_text.split("-", 1)
        else:
            subject, topic = "عام", user_text

        subject, topic = subject.strip(), topic.strip()
        context.user_data["subject"] = subject

        if current_service == "svc_ppt":
            file_name = f"presentation_{user.id}_{uuid.uuid4().hex[:4]}.pptx"
            create_powerpoint_presentation_full(grade=grade, subject=subject, topic=topic, output_path=file_name)

            with open(file_name, "rb") as ppt_file:
                await update.message.reply_document(
                    document=ppt_file, filename=f"{subject[:20]}_{topic[:25]}_1448H.pptx",
                    caption=f"✅ تم إنشاء العرض التعليمي بنجاح\n\n🎓 الصف: {grade}\n📚 المادة: {subject}\n📌 الدرس: {topic}"
                )
            if os.path.exists(file_name): os.remove(file_name)

        else:
            file_name = f"doc_{user.id}_{uuid.uuid4().hex[:4]}.docx"
            create_educational_doc_1448(service_code=current_service, grade=grade, subject=subject, topic=topic, output_path=file_name)

            with open(file_name, "rb") as doc_file:
                await update.message.reply_document(
                    document=doc_file, filename=f"{service_name[:20]}_{topic[:25]}_1448H.docx",
                    caption=f"✅ تم تجهيز مستند Word\n\nالخدمة: {service_name}\nالصف: {grade}\nالمادة: {subject}"
                )
            if os.path.exists(file_name): os.remove(file_name)

        await status_msg.delete()

    except Exception as exc:
        await status_msg.edit_text(f"⚠️ حدث خطأ. حاول صياغة عنوان الدرس بشكل أوضح. \nالتفاصيل: {str(exc)[:200]}")

# ============================================================
# خادم الويب للتشغيل على الاستضافة
# ============================================================
async def handle_ping(request): return web.Response(text="Bot is running!")

async def start_web_server():
    server = web.Application()
    server.router.add_get("/", handle_ping)
    runner = web.AppRunner(server)
    await runner.setup()
    port = int(os.environ.get("PORT", "10000"))
    site = web.TCPSite(runner, "0.0.0.0", port)
    await site.start()

async def main_async():
    await start_web_server()
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_handler))
    
    print("البوت يعمل الآن (النسخة الخالية من التكرار) - 1448هـ")
    await app.initialize()
    await app.start()
    await app.updater.start_polling()
    while True: await asyncio.sleep(3600)

def main(): asyncio.run(main_async())

if __name__ == "__main__": main()
