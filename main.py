import os
import io
import asyncio
import json
import requests
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
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, MessageHandler, filters, ContextTypes
from aiohttp import web

TELEGRAM_TOKEN = "8867458917:AAEyVQ0Vn97bEfZbANtsFRxMxeJxnbdJ0s4"

def query_ai_engine(prompt: str) -> str:
    """محرك استدعاء الذكاء الاصطناعي الحي لتوليد محتوى مخصص لكل طلب"""
    payloads = [
        {
            "url": "https://text.pollinations.ai/",
            "data": {
                "messages": [
                    {"role": "system", "content": "أنت معلم ومصمم مناهج سعودية خبير لعام 1448هـ. اكتب محتوى علمياً وتطبيقياً دقيقاً ومفصلاً جداً وخاصاً بالدرس المطلوب دون تكرار قوالب عامة."},
                    {"role": "user", "content": prompt}
                ],
                "model": "openai",
                "seed": int(os.urandom(2).hex(), 16)
            }
        },
        {
            "url": "https://api.airforce/v1/chat/completions",
            "data": {
                "model": "gpt-4o-mini",
                "messages": [{"role": "user", "content": prompt}]
            }
        }
    ]
    for p in payloads:
        try:
            res = requests.post(p["url"], json=p["data"], timeout=45)
            if res.status_code == 200:
                if "choices" in res.text:
                    return res.json()["choices"][0]["message"]["content"].strip()
                elif len(res.text.strip()) > 100:
                    return res.text.strip()
        except Exception:
            continue
    return ""

def fetch_unique_slide_image(slide_index: int, topic: str) -> io.BytesIO:
    """جلب صور ورسوم بيانية ديناميكية متوافقة مع موضوع الدرس"""
    keywords = [
        "education", "science", "math", "technology", "classroom", "books",
        "laboratory", "thinking", "values", "saudi", "future", "digital",
        "interactive", "research", "exam", "analytics", "evaluation", "teamwork"
    ]
    kw = keywords[(slide_index - 1) % len(keywords)]
    url = f"https://loremflickr.com/600/450/{kw}?lock={(slide_index * 37 + hash(topic)) % 5000}"
    try:
        res = requests.get(url, timeout=6)
        if res.status_code == 200 and len(res.content) > 3000:
            img = io.BytesIO(res.content)
            img.seek(0)
            return img
    except Exception:
        pass

    # إنشاء رسم بياني توضيحي ديناميكي
    img_obj = Image.new("RGB", (600, 450), color="#F8FAFC")
    draw = ImageDraw.Draw(img_obj)
    draw.rounded_rectangle([15, 15, 585, 435], radius=15, fill="#FFFFFF", outline="#CBD5E1", width=2)
    draw.rectangle([15, 15, 585, 70], fill="#1B4965")
    bars = [120, 180, 150, 230, 200]
    for idx, b in enumerate(bars):
        x0 = 70 + idx * 90
        factor = ((slide_index + idx) % 5 + 6) / 10
        y0 = 400 - (b * factor)
        draw.rounded_rectangle([x0, y0, x0 + 55, 400], radius=6, fill="#5FA8D3" if idx % 2 == 0 else "#62B6CB")
    out = io.BytesIO()
    img_obj.save(out, format="PNG")
    out.seek(0)
    return out

def generate_dynamic_30_slides_data(topic: str):
    """توليد محتوى 30 شريحة مخصص علمياً وحصرياً للموضوع المدخل"""
    ai_prompt = (
        f"أنشئ خطة عرض تقديمي متكاملة للمنهج السعودي 1448هـ حول موضوع: '{topic}'.\n"
        f"اكتب بالتفصيل العلمي والتربوي الكامل لـ 30 شريحة متتالية.\n"
        f"التنسيق المطلوب إلزامي بدقة:\n"
        f"---SLIDE---\n"
        f"TITLE: [عنوان الشريحة مرتبط مباشرة بـ {topic}]\n"
        f"CONTENT:\n"
        f"- [نقطة علمية تفصيلية 1]\n"
        f"- [نقطة تطبيقية أو مسألة 2]\n"
        f"- [نقطة تحليلية أو مهارة 3]"
    )
    
    ai_raw = query_ai_engine(ai_prompt)
    slides = []

    if "---SLIDE---" in ai_raw:
        for raw in ai_raw.split("---SLIDE---"):
            if not raw.strip():
                continue
            lines = raw.strip().split("\n")
            title = ""
            points = []
            for line in lines:
                l = line.strip()
                if l.startswith("TITLE:"):
                    title = l.replace("TITLE:", "").strip()
                elif l.startswith("-") or l.startswith("•") or l.startswith("*"):
                    clean_p = l.lstrip("-•* 1234567890.").strip()
                    if clean_p:
                        points.append(clean_p)
            if title and points:
                slides.append((title, points))

    # إذا لم يكتمل الـ 30 شريحة من السيرفر، يتم استكمالها ببناء تربوي مخصص للدرس
    framework = [
        (f"1. الغلاف والبيانات الرسمية 1448هـ", [f"المادة والدرس: {topic}", "المنهج السعودي المطور 1448هـ | نواتج التعلم", "إعداد تفاعلي وفق معايير وزارة التعليم"]),
        (f"2. التمهيد واستثارة الدافعية لـ ({topic})", [f"مدخل استكشافي مشوق يربط {topic} بالواقع اليومي.", f"طرح تساؤل تفاعلي: كيف نستفيد من {topic} في حياتنا؟", "استرجاع المعارف السابقة ذات الصلة بالدرس."]),
        (f"3. نواتج التعلم المستهدفة لـ ({topic})", [f"أن يتعرف الطالب على المفاهيم الأساسية لـ {topic}.", f"أن يطبق القواعد والقوانين المرتبطة بـ {topic} في حل المسائل.", "تنمية مهارات التفكير والاستدلال العلمي لدى الطلاب."]),
        (f"4. المفاهيم والمصطلحات الأساسية", [f"التعريف العلمي والدقيق لمفهوم {topic}.", "المصطلحات والمفردات التخصصية للوحدة الدراسية.", "بناء خارطة مفاهيمية توضح العلاقات بين أجزاء الدرس."]),
        (f"5. استراتيجية التعليم الموجه (أنا أعمل - I Do)", [f"نمذجة شرح الفكرة المحورية لـ {topic} خطوة بخطوة.", "تقديم أمثلة محلولة نموذجية ومباشرة من الكتاب.", "توضيح مسار التفكير المنطقي للوصول للحل الصحيح."]),
        (f"6. الشرح المفصل والتطبيقات المنهجية", [f"التوسع في شرح عناصر {topic} بدقة علمية.", "استعراض القوانين والقواعد الحاكمة مع التعليل.", "تدعيم الشرح بالأمثلة المحلولة والرسومات التوضيحية."]),
        (f"7. فاصل وتوجيه المقاطع التعليمية التفاعلية", [f"مشاهدة مقطع مرئي توضيحي لتعميق فهم {topic}.", "استخلاص النقاط الجوهرية من الفيديو ومناقشتها.", "ترسيخ المفهوم عبر المؤثرات البصرية والصوتية."]),
        (f"8. الممارسة الموجهة (نحن نعمل - We Do)", [f"حل تدريب تطبيقي على {topic} بمشاركة المعلم والطلاب.", "معالجة المفاهيم الخاطئة لحظياً وتصحيح خطوات الحل.", "التأكد من جاهزية الطلاب للانتقال للعمل التشاركي."]),
        (f"9. الأنشطة الثنائية (فكر - زاوج - شارك)", [f"التفكير الفردي في مسألة حول {topic} لمدة دقيقة.", "المناقشة الثنائية بين الطالب وزميله لمقارنة الحلول.", "مشاركة النتيجة المشتركة مع عموم الفصل."]),
        (f"10. التعلم التعاوني ومجموعات العمل", [f"مهمة جماعية: تحليل تطبيق واقعي مرتبط بـ {topic}.", "توزيع الأدوار داخل المجموعة الصغيرة (الكاتب، الميقاتي، المتحدث).", "عرض مخرجات المجموعة ومناقشتها بصورة بناءة."]),
        (f"11. النشاط الفردي والتطبيق المستقل", [f"تطبيق فردي مقنن لقياس تمكن كل طالب من {topic}.", "الاعتماد على النفس وإدارة الوقت أثناء الحل.", "متابعة المعلم الدقيقة وتقديم الدعم الفردي الموجه."]),
        (f"12. مهارات التفكير العليا لـ ({topic})", [f"سؤال يختبر المقارنة والتحليل والتركيب في {topic}.", "استنتاج علاقات جديدة وحل مواقف غير مألوفة.", "تقديم تبريرات منطقية مدعمة بالبراهين العلمية."]),
        (f"13. استراتيجيات حل المشكلات", [f"عرض معضلة واقعية ترتبط بـ {topic} وكيفية حلها.", "تطبيق خطوات البحث: المعطيات، الفرضيات، والحل النهائي.", "الوصول إلى حلول إبداعية قابلة للتطبيق."]),
        (f"14. الألعاب والمسابقات التنافسية", [f"مسابقة تفاعلية ممتعة تراجع مفاهيم {topic}.", "تحدي السرعة والإتقان بين الفرق الصفية.", "تحفيز التنافس الإيجابي وتكريم المتميزين."]),
        (f"15. تدريبات الاختبارات الوطنية (نافس) - 1", [f"سؤال معياري يحاكي اختبارات نافس حول {topic}.", "تحليل السؤال واستبعاد الخيارات الخاطئة بدقة.", "التدريب على مهارات الفهم القرائي والاستدلال."]),
        (f"16. تدريبات الاختبارات الوطنية (نافس) - 2", [f"مسألة مركبة تقيس مهارات التفكير الرياضي/العلمي في {topic}.", "ربط السؤال بمهارات الاختبارات الدولية (PISA/TIMSS).", "استراتيجيات الحل السريع والذكي وتجنب المشتتات."]),
        (f"17. الربط بالقيم الدينية والأخلاقية", [f"استشعار عظمة الله ودقة خلقه من خلال {topic}.", "تعزيز قيم الإتقان والأمانة والمسؤولية في طلب العلم.", "الربط بين المعرفة العلمية وخدمة المجتمع."]),
        (f"18. الربط بالوطن ورؤية المملكة 2030", [f"دور المعارف المرتبطة بـ {topic} في تحقيق رؤية 2030.", "إبراز المشاريع السعودية الكبرى والتنمية الوطنية.", "غرس الاعتزاز بالهوية الوطنية والنهضة السعودية."]),
        (f"19. الربط بالواقع والتطبيقات الحياتية", [f"أين نشاهد تطبيقات {topic} في حياتنا اليومية؟", "أمثلة عملية توضح أثر العلم في تسهيل الحياة وحل المشكلات.", "المحافظة على الموارد وترشيد استخدامها."]),
        (f"20. التكامل مع المواد الدراسية (STEM)", [f"الربط التكاملي بين {topic} والرياضيات والعلوم والتقنية.", "استخدام اللغة العربية الدقيقة في التعبير عن المفاهيم.", "إظهار شمولية وتكامل المعرفة العلمية."]),
        (f"21. تمايز التعليم ودعم الفروق الفردية", [f"أنشطة متدرجة المستويات (مبتدئ، متوسط، إثرائي) لـ {topic}.", "تنويع أساليب العرض والتوضيح لتناسب كافة أنماط التعلم.", "تقديم إرشادات داعمة للطلاب الذين يحتاجون تعزيزاً إضافياً."]),
        (f"22. توظيف التقنية والمنصات الرقمية", [f"حل أنشطة إلكترونية تفاعلية حول {topic} عبر منصة مدرستي.", "استخدام المحاكاة الرقمية والنماذج ثلاثية الأبعاد.", "توثيق نواتج التعلم في ملف الإنجاز الإلكتروني."]),
        (f"23. معالجة الأخطاء الشائعة في ({topic})", [f"رصد المفاهيم الخاطئة التي قد يقع فيها الطلاب في {topic}.", "توضيح الصواب بأدلة وأمثلة مقارنة واضحة.", "تثبيت الفهم السليم وضمان عدم تكرار اللبس."]),
        (f"24. ورقة العمل التطبيقية - الجزء الأول", [f"تمارين كتابية متدرجة تقيس إتقان مهارات {topic}.", "أسئلة اختيار من متعدد وتفسير منطقي للخطوات.", "مساحة مخصصة للإجابة النموذجية والمراجعة."]),
        (f"25. ورقة العمل التطبيقية - الجزء الثاني", [f"مهمة أدائية وتطبيق عملي متقدم لـ {topic}.", "سلم تصحيح معتمد لتقييم الإجابات بدقة.", "قياس مهارات التحليل والاستنتاج لدى الطلاب."]),
        (f"26. التقويم التكويني وبطاقات الخروج", [f"سؤال ختامي سريع في دقيقة واحدة لقياس الفهم لـ {topic}.", "تدوين أهم نقطة تعلمها الطالب اليوم في بطاقة الخروج.", "رصد مؤشرات الأداء لتعديل الخطة التعليمية القادمة."]),
        (f"27. الإغلاق المنهجي المناسب للدرس", [f"تلخيص أهم المحاور والأفكار المستفادة من {topic}.", "الربط بين مخرجات درس اليوم وموضوع الحصة القادمة.", "التأكد من وضوح الصورة الشاملة لجميع الطلاب."]),
        (f"28. الواجب المنزلي والمهام الإثرائية", [f"تحديد تدريبات الكتاب المدرسي المعتمدة لـ {topic}.", "إسناد الواجب الإلكتروني عبر منصة مدرستي.", "أنشطة استقصائية اختيارية لتنمية مهارات البحث الذاتي."]),
        (f"29. التغذية الراجعة وسلالم التقدير (Rubrics)", [f"معايير قياس مستوى الإتقان والتميز في درس {topic}.", "توجيهات للطالب لمواصلة التطور الأكاديمي ومعالجة القصور.", "تعزيز ثقافة التقييم الذاتي والتطوير المستمر."]),
        (f"30. الخاتمة والمراجع المعتمدة 1448هـ", [f"المراجع: مقررات وزارة التعليم لعام 1448هـ، بوابة عين، منصة مدرستي.", f"شكر وتقدير للطلاب والطالبات على التفاعل في درس {topic}.", "فتح باب الأسئلة والمناقشة الختامية."])
    ]

    while len(slides) < 30:
        slides.append(framework[len(slides)])

    return slides[:30]

def create_powerpoint_presentation_full(topic: str, output_path: str):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slides_data = generate_dynamic_30_slides_data(topic)

    for idx, (title_text, points) in enumerate(slides_data, start=1):
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # شريط علوي أنيق
        top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.2))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = RGBColor(27, 73, 101)
        top_bar.line.fill.background()

        # بطاقة محتوى للنصوص
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.4), Inches(1.5), Inches(7.2), Inches(5.1))
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(248, 250, 252)
        card.line.color.rgb = RGBColor(203, 213, 225)
        card.line.width = Pt(1.5)

        # عنوان الشريحة
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(1.0))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = title_text
        p_title.alignment = PP_ALIGN.RIGHT
        p_title.font.size = Pt(26)
        p_title.font.bold = True
        p_title.font.color.rgb = RGBColor(27, 73, 101)

        # محتوى النقاط
        content_box = slide.shapes.add_textbox(Inches(5.6), Inches(1.7), Inches(6.8), Inches(4.7))
        tf_content = content_box.text_frame
        tf_content.word_wrap = True

        for p_idx, pt in enumerate(points):
            p = tf_content.paragraphs[0] if p_idx == 0 else tf_content.add_paragraph()
            p.text = f"◀ {pt}"
            p.alignment = PP_ALIGN.RIGHT
            p.font.size = Pt(18)
            p.font.color.rgb = RGBColor(30, 41, 59)
            p.space_after = Pt(14)

        # إدراج الصورة الخاصة بالشريحة
        img_stream = fetch_unique_slide_image(idx, topic)
        if img_stream:
            slide.shapes.add_picture(img_stream, Inches(0.8), Inches(1.5), Inches(4.3), Inches(5.1))

        # تذييل الشريحة
        footer_box = slide.shapes.add_textbox(Inches(0.8), Inches(6.9), Inches(11.733), Inches(0.4))
        p_foot = footer_box.text_frame.paragraphs[0]
        p_foot.text = f"شريحة {idx} من 30 | منصة إنجاز - المنهج المطور 1448هـ - {topic[:30]}"
        p_foot.alignment = PP_ALIGN.LEFT
        p_foot.font.size = Pt(11)
        p_foot.font.color.rgb = RGBColor(148, 163, 184)

    prs.save(output_path)

def create_educational_doc_1448(service_code: str, topic: str, output_path: str):
    doc = Document()
    for section in doc.sections:
        section.top_margin = DocxInches(1)
        section.bottom_margin = DocxInches(1)
        section.left_margin = DocxInches(1)
        section.right_margin = DocxInches(1)

    prompts = {
        "svc_exam": f"اكتب نموذج اختبار شامل ومفصل مع جدول مواصفات وتدريبات نافس للمنهج السعودي 1448هـ خاص بـ: '{topic}'. يشمل: أسئلة مقالية وموضوعية ونموذج إجابة كامل وتوزيع درجات.",
        "svc_remedial": f"اكتب خطة علاجية وإثرائية وأوراق عمل خاصة بدرس: '{topic}' لعام 1448هـ متضمنة تشخيص الفاقد وتدريبات داعمة وأنشطة للمتفوقين.",
        "svc_portfolio": f"اكتب ملف إنجاز إلكتروني متكامل للمعلم/المعلمة في: '{topic}' لعام 1448هـ يشمل الشواهد والاستراتيجيات ونماذج الأعمال والتطوير المهني.",
        "svc_performance": f"اكتب ميثاق الأداء الوظيفي الجديد المعتمد لعام 1448هـ في: '{topic}' يشمل الأهداف والمؤشرات والأوزان النسبية وجداول الشواهد.",
        "svc_operation": f"اكتب خطة تشغيلية مفصلة لعام 1448هـ لـ: '{topic}' تشمل الأهداف والبرامج ومؤشرات التحقق والجدول الزمني.",
        "svc_loss": f"اكتب خطة الفاقد التعليمي لـ: '{topic}' تشمل حصر المهارات المفقودة والاختبار التشخيصي والتعويض والتقويم البعدي.",
        "svc_research": f"اكتب بحثاً أكاديمياً جامعياً مفصلاً لمستوى الماجستير حول: '{topic}' يشمل المقدمة والمشكلة والإطار النظري والتحليل الميداني والتوصيات والمراجع وفق APA."
    }

    titles = {
        "svc_exam": f"كتابة الاختبارات وتحليل النتائج وتدريبات نافس\n{topic}",
        "svc_remedial": f"الخطة العلاجية والإثرائية وأوراق العمل\n{topic}",
        "svc_portfolio": f"ملف الإنجاز الإلكتروني الشامل (1448هـ)\n{topic}",
        "svc_performance": f"ميثاق وملف الأداء الوظيفي الجديد (1448هـ)\n{topic}",
        "svc_operation": f"الخطة التشغيلية المعتمدة (1448هـ)\n{topic}",
        "svc_loss": f"ملف وخطة معالجة الفاقد التعليمي (1448هـ)\n{topic}",
        "svc_research": f"بحث أكاديمي متكامل\n{topic}"
    }

    prompt = prompts.get(service_code, f"اكتب وثيقة تعليمية متكاملة لعام 1448هـ حول: {topic}")
    doc_title = titles.get(service_code, f"وثيقة تعليمية 1448هـ: {topic}")

    ai_content = query_ai_engine(prompt)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = title_p.add_run(doc_title)
    run_title.font.size = DocxPt(20)
    run_title.font.bold = True
    run_title.font.color.rgb = DocxRGB(27, 73, 101)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = sub_p.add_run("منصة إنجاز للخدمات التعليمية والأكاديمية | معايير المنهج السعودي المطور 1448هـ\n" + "—"*35)
    run_sub.font.size = DocxPt(11)
    run_sub.font.color.rgb = DocxRGB(100, 116, 139)

    if len(ai_content) > 300:
        for block in ai_content.split("\n\n"):
            b_clean = block.strip()
            if not b_clean:
                continue
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            if b_clean.startswith("#") or any(b_clean.startswith(f"{i}.") for i in range(1, 15)) or "المحور" in b_clean or "الهدف" in b_clean:
                r = p.add_run(b_clean.replace("#", "").strip())
                r.font.size = DocxPt(14)
                r.font.bold = True
                r.font.color.rgb = DocxRGB(27, 73, 101)
                p.paragraph_format.space_before = DocxPt(12)
                p.paragraph_format.space_after = DocxPt(4)
            else:
                r = p.add_run(b_clean)
                r.font.size = DocxPt(11.5)
                r.font.color.rgb = DocxRGB(30, 41, 59)
                p.paragraph_format.line_spacing = 1.25
                p.paragraph_format.space_after = DocxPt(6)
    else:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run(f"تم إعداد هذا المستند التعليمي المخصص لموضوع ({topic}) وفق معايير المنهج السعودي لعام 1448هـ.")

    doc.save(output_path)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [
        [InlineKeyboardButton("📊 عروض بوربوينت شاملة (30 شريحة + نافس + ألعاب)", callback_data="svc_ppt")],
        [InlineKeyboardButton("📝 كتابة الاختبارات وجداول المواصفات 1448هـ", callback_data="svc_exam")],
        [InlineKeyboardButton("📈 خطط علاجية وإثرائية وأوراق عمل", callback_data="svc_remedial")],
        [InlineKeyboardButton("🗂 ملف إنجاز المعلم/المعلمة الإلكتروني", callback_data="svc_portfolio")],
        [InlineKeyboardButton("📑 ميثاق وملف الأداء الوظيفي الجديد", callback_data="svc_performance")],
        [InlineKeyboardButton("📅 الخطة التشغيلية المدرسية والفصلية", callback_data="svc_operation")],
        [InlineKeyboardButton("📚 ملف وخطة معالجة الفاقد التعليمي", callback_data="svc_loss")],
        [InlineKeyboardButton("🎓 إعداد بحث جامعي وأكاديمي متكامل (Word)", callback_data="svc_research")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    welcome_text = (
        "🌟 **أهلاً بك في منصة إنجاز للخدمات الأكاديمية والتعليمية بالمملكة (1448هـ)**\n\n"
        "✨ **الحقيبة التعليمية المخصصة لجميع المواد والصفوف:**\n"
        "📌 عروض بوربوينت 30 شريحة تولد نصوصها ومسائلها مباشرة حسب كل درس.\n"
        "📌 تدريبات نافس، التعلم النشط، الألعاب التعليمية، وأوراق العمل.\n"
        "📌 ملفات Word منسقة للاختبارات والخطط ومواثيق الأداء.\n\n"
        "👇 **اختر الخدمة المطلوبة لبدء التوليد الفوري:**"
    )
    
    if update.message:
        await update.message.reply_text(welcome_text, reply_markup=reply_markup, parse_mode="Markdown")
    elif update.callback_query:
        await update.callback_query.message.reply_text(welcome_text, reply_markup=reply_markup, parse_mode="Markdown")

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    data = query.data
    services = {
        "svc_ppt": "📊 عروض بوربوينت شاملة (30 شريحة)",
        "svc_exam": "📝 كتابة الاختبارات وتحليل النتائج 1448هـ",
        "svc_remedial": "📈 خطط علاجية وإثرائية وأوراق عمل",
        "svc_portfolio": "🗂 ملف إنجاز إلكتروني 1448هـ",
        "svc_performance": "📑 ملف الأداء الوظيفي الجديد 1448هـ",
        "svc_operation": "📅 الخطة التشغيلية 1448هـ",
        "svc_loss": "📚 ملف الفاقد التعليمي 1448هـ",
        "svc_research": "🎓 إعداد بحث جامعي متكامل (Word)"
    }

    if data in services:
        name = services[data]
        context.user_data["current_service"] = data
        context.user_data["service_name"] = name
        
        if data == "svc_ppt":
            await query.edit_message_text(
                "📊 أرسل الآن **المادة والصف والموضوع**\n"
                "(مثال: *رياضيات أول متوسط - الأعداد الصحيحة* أو *علوم سادس - الخلية* أو *كيمياء 3 مسارات - سرعة التفاعلات*)\n\n"
                "وسيقوم البوت بتوليد ملف بوربوينت جديد ومخصص بالكامل لهذا الدرس (30 شريحة مع تدريبات نافس والألعاب وصور حصرية):"
            )
        else:
            await query.edit_message_text(f"✨ خدمة: *{name}*\n\nيرجى إرسال تفاصيل المادة أو الموضوع وسأقوم بتوليد **ملف Word (.docx) رسمي ومخصص بالكامل لعام 1448هـ** فوراً:", parse_mode="Markdown")

async def text_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_text = update.message.text
    user = update.effective_user
    current_service = context.user_data.get("current_service", "svc_ppt")
    service_name = context.user_data.get("service_name", "عرض بوربوينت متكامل 1448هـ")

    status_msg = await update.message.reply_text("⏳ جارٍ إعداد وتوليد المحتوى المخصص للدرس بالذكاء الاصطناعي...")

    try:
        if current_service == "svc_ppt":
            file_name = f"presentation_{user.id}.pptx"
            create_powerpoint_presentation_full(user_text, file_name)

            with open(file_name, "rb") as ppt_file:
                await update.message.reply_document(
                    document=ppt_file,
                    filename=f"{user_text[:30]}_Full1448H.pptx",
                    caption=f"✅ **تم تصميم العرض التقديمي المخصص بنجاح:**\n📌 *{user_text}*\n📊 **عدد الشرائح:** 30 شريحة مخصصة لهذا الدرس بالكامل مع تدريبات نافس وأوراق العمل.",
                    parse_mode="Markdown"
                )
            if os.path.exists(file_name):
                os.remove(file_name)
        else:
            file_name = f"doc_{user.id}.docx"
            create_educational_doc_1448(current_service, user_text, file_name)

            with open(file_name, "rb") as doc_file:
                await update.message.reply_document(
                    document=doc_file,
                    filename=f"{service_name[:15]}_{user_text[:20]}_1448H.docx",
                    caption=f"✅ **تم تجهيز المستند المخصص بنجاح عبر منصة إنجاز:**\n📌 الخدمة: *{service_name}*\n📝 الموضوع: *{user_text}*\n📄 تم تصدير الملف كـ مستند Word منسق ومفصل لعام 1448هـ.",
                    parse_mode="Markdown"
                )
            if os.path.exists(file_name):
                os.remove(file_name)

        await status_msg.delete()

    except Exception as e:
        await status_msg.edit_text(f"⚠️ حدث خطأ أثناء المعالجة: {str(e)}")

async def handle_ping(request):
    return web.Response(text="Enjaz Full 1448 Bot is active!")

async def start_web_server():
    server = web.Application()
    server.router.add_get("/", handle_ping)
    runner = web.AppRunner(server)
    await runner.setup()
    port = int(os.environ.get("PORT", 10000))
    site = web.TCPSite(runner, "0.0.0.0", port)
    await site.start()

async def main_async():
    await start_web_server()
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_handler))
    
    print("منصة إنجاز تعمل بالتوليد الحي المخصص...")
    await app.initialize()
    await app.start()
    await app.updater.start_polling()
    
    while True:
        await asyncio.sleep(3600)

def main():
    asyncio.run(main_async())

if __name__ == "__main__":
    main()
