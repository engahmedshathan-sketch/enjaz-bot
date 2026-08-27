import os
import io
import asyncio
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
    payloads = [
        {
            "url": "https://text.pollinations.ai/",
            "data": {
                "messages": [
                    {"role": "system", "content": "أنت خبير معتمد في المناهج السعودية والخطط الدراسية الحديثة لعام 1448هـ بنظام المسارات ومنصة مدرستي. اكتب محتوى تعليمياً مفصلاً ورصيناً وفق أحدث أدلة وزارة التعليم وهيئة تقويم التعليم والتدريب."},
                    {"role": "user", "content": prompt}
                ],
                "model": "openai",
                "seed": 1448
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
                elif len(res.text.strip()) > 200:
                    return res.text.strip()
        except Exception:
            continue
    return ""

def fetch_unique_slide_image(slide_index: int) -> io.BytesIO:
    keywords = [
        "education", "saudi", "ai", "technology", "cybersecurity", "space",
        "engineering", "health", "business", "data", "classroom", "learning",
        "robotics", "digital", "research", "exam", "excellence", "vision2030",
        "analytics", "innovation", "stem", "logic", "strategy", "evaluation",
        "sustainability", "coding", "physics", "biology", "presentation", "success"
    ]
    kw = keywords[(slide_index - 1) % len(keywords)]
    url = f"https://loremflickr.com/600/450/{kw}?lock={slide_index * 23 + 1448}"
    try:
        res = requests.get(url, timeout=8)
        if res.status_code == 200 and len(res.content) > 3000:
            img = io.BytesIO(res.content)
            img.seek(0)
            return img
    except Exception:
        pass

    img_obj = Image.new("RGB", (600, 450), color="#F8FAFC")
    draw = ImageDraw.Draw(img_obj)
    draw.rounded_rectangle([15, 15, 585, 435], radius=15, fill="#FFFFFF", outline="#CBD5E1", width=2)
    draw.rectangle([15, 15, 585, 70], fill="#1B4965")
    bars = [130, 190, 160, 240, 210]
    for idx, b in enumerate(bars):
        x0 = 70 + idx * 90
        factor = ((slide_index + idx) % 5 + 6) / 10
        y0 = 400 - (b * factor)
        draw.rounded_rectangle([x0, y0, x0 + 55, 400], radius=6, fill="#5FA8D3" if idx % 2 == 0 else "#62B6CB")
    out = io.BytesIO()
    img_obj.save(out, format="PNG")
    out.seek(0)
    return out

def generate_curriculum_1448_slides(topic: str):
    return [
        ("1. الغلاف والبيانات الرسمية 1448هـ", [f"المادة والدرس: {topic}", "المنهج السعودي المطور 1448هـ | نظام المسارات", "إعداد تفاعلي معتمد وفق نواتج التعلم"]),
        ("2. التهيئة واستثارة الدافعية", [f"مدخل تفاعلي لموضوع ({topic}) عبر ربطه بالواقع الرقمي.", "ربط معارف الدرس بالخبرات السابقة والتطبيقات الحياتية.", "طرح تساؤل محوري يستثير التفكير الإبداعي والاستكشاف."]),
        ("3. نواتج التعلم المستهدفة (معايير ETEC)", ["صياغة الأهداف السلوكية والمعرفية وفق مستويات بلوم المتقدمة.", "تنمية المهارات الأدائية والمهنية المطلوبة لعام 1448هـ.", "تعزيز القيم والاتجاهات والمواطنة الرقمية الصالحة."]),
        ("4. المفاهيم والمصطلحات التخصصية", ["حصر وتدقيق المصطلحات والمفاهيم العلمية للوحدة الدراسية.", "التفرقة الدقيقة بين المفاهيم المتقاربة بأمثلة عملية.", "بناء خارطة مفاهيمية متكاملة تثبت في الذاكرة طويلة المدى."]),
        ("5. استراتيجية التعليم الموجه (أنا أعمل - I Do)", ["النمذجة المباشرة وتفكيك المهارة إلى خطوات محددة.", "شرح الإجراءات والحلول النموذجية بصوت مسموع ومرئي.", "إبراز أساليب التفكير العلمي السليم في حل المسائل."]),
        ("6. الشرح المفصل والتطبيقات المنهجية", ["التوسع في محاور الدرس وفق الطبعة الحديثة للكتاب المدرسي.", "استعراض القوانين والقواعد العلمية وتفسيراتها الدقيقة.", "تدعيم الشرح بمخططات وأشكال بيانية ورقمية تفاعلية."]),
        ("7. الممارسة الموجهة (نحن نعمل - We Do)", ["حل تطبيقات وأنشطة صفية بمشاركة المعلم والطلاب معاً.", "تصحيح المفاهيم الخاطئة لحظياً وتقديم التغذية الراجعة.", "التأكد من جاهزية الطلاب للانتقال للعمل التشاركي."]),
        ("8. التعلم التعاوني ومجموعات العمل", ["توزيع الأدوار داخل المجموعات الصغيرة لإنجاز مهمة محددة.", "تحفيز العصف الذهني وتبادل الأفكار لتعميق الفهم.", "عرض نتائج المجموعات ومناقشتها بشكل جماعي بناء."]),
        ("9. تمايز التعليم ومراعاة الفروق الفردية", ["تصميم أنشطة متدرجة المستويات (تأسيسي، متوسط، إثرائي).", "تنويع أساليب التعلم (البصري، السمعي، الرقمي، الحركي).", "توفير بطاقات دعم مساندة للطلبة ذوي الاحتياج الإضافي."]),
        ("10. النشاط والتطبيق المستقل (أنت تعمل)", ["تكليف الطالب بمهمة فردية لقياس مدى تمكنه من المهارة.", "تنمية مهارات التفكير الناقد والاستقلالية في الحل.", "متابعة المعلم الدقيقة لتقديم الدعم الفردي الموجه."]),
        ("11. التفكير الناقد واستراتيجيات حل المشكلات", ["طرح سيناريو واقعي يتطلب اتخاذ قرار علمي منطقي.", "تحليل المعطيات وفرض الفروض وتقييم البدائل المتاحة.", "تدريب الطلاب على الدفاع عن استنتاجاتهم بالبراهين."]),
        ("12. مهارات المستقبل والتحول الرقمي 1448هـ", ["ربط مخرجات الدرس بتقنيات الذكاء الاصطناعي والأتمتة.", "توظيف البرمجيات التفاعلية والمحاكاة الافتراضية للعلوم.", "تطوير مهارات التعامل مع البيانات والتحليل الرقمي."]),
        ("13. المواءمة مع رؤية 2030 ومشاريع الوطن", ["إبراز دور الموضوع في مسارات التنمية الوطنية المستدامة.", "ربط الدرس بقطاعات الطاقة المتجددة والاقتصاد المعرفي.", "غرس الاعتزاز بالمشاريع السعودية الكبرى ومستقبل الوطن."]),
        ("14. الأنشطة العملية والمعامل الافتراضية", ["خطوات تنفيذ التجارب العلمية والاستقصاء الصفي بأمان.", "تدوين الملاحظات والبيانات وتفسير النتائج موضوعياً.", "الربط بين التجربة الميدانية والقوانين النظرية المعتمدة."]),
        ("15. التكامل مع منصة مدرستي ومايكروسوفت 365", ["تفعيل الأنشطة والواجبات الإلكترونية عبر منصة مدرستي.", "استخدام النماذج الرقمية (Forms) في التقويم السريع.", "أرشفة نواتج التعلم في ملف الإنجاز الإلكتروني الموحد."]),
        ("16. خطة معالجة الفاقد التعليمي في الدرس", ["تشخيص المهارات السابقة غير المتقنة ومعالجتها فوراً.", "تقديم تدريبات تركيزية لسد الفجوات المعرفية التراكمية.", "ضمان استيفاء المتطلبات القبلية لمواصلة بناء الدرس."]),
        ("17. الأنشطة الإثرائية ورعاية الموهوبين", ["تكليف الطلبة المتميزين ببحوث وتحديات ابتكارية موسعة.", "توجيه الموهوبين للمشاركة في المسابقات والمعارض العلمية.", "تنمية مهارات القيادة والبحث المستقل المتقدم."]),
        ("18. الألعاب التعليمية والتحفيز التنافسي", ["مسابقات سريعة وتحديات تفاعلية تنشط التفكير الصفي.", "استخدام استراتيجيات (فكر - زاوج - شارك) والمسابقات الرقمية.", "تعزيز روح التنافس الإيجابي والشغف بالمعرفة."]),
        ("19. السلامة المهنية والمواطنة الرقمية", ["إرشادات الأمن والسلامة في البيئات المدرسية والمعامل.", "ضوابط الاستخدام الآمن للإنترنت والذكاء الاصطناعي.", "تعزيز الانضباط الذاتي والأخلاقيات الرقمية السليمة."]),
        ("20. المهارات الحياتية والتطبيق الواقعي", ["تحويل المعرفة النظرية إلى ممارسات وسلوكيات يومية.", "تنمية مهارات التفاوض، العمل الجماعي، وإدارة الوقت.", "ترشيد استهلاك الموارد وحماية البيئة الطبيعية."]),
        ("21. التقويم التكويني السريع (بطاقات الخروج)", ["قياس استيعاب المفاهيم الأساسية عبر أسئلة ختامية موجزة.", "تدوين أهم فكرة تم اكتسابها في بطاقة الخروج الرقمية.", "رصد المؤشرات الفورية لتعديل خطة التدريس للحصة القادمة."]),
        ("22. جدول المواصفات ونماذج الاختبارات", ["توزيع الأسئلة وفق الوزن النسبي لدروس المنهج المطور 1448هـ.", "نماذج لأسئلة قياس التفكير العليا والأسئلة المقننة.", "معايير تصحيح وسلالم تقدير (Rubrics) دقيقة وواضحة."]),
        ("23. معالجة الأخطاء الشائعة وسوء الفهم", ["رصد المفاهيم المغلوطة الشائعة وتصحيحها علمياً.", "تقديم أدلة وأمثلة توضح الفرق بين الصواب والخطأ.", "تثبيت الفهم السليم ومنع تكرار اللبس المفاهيمي."]),
        ("24. ملخص الوحدة والخارطة الذهنية الشاملة", ["تجميع محاور الدرس الرئيسية في خارطة ذهنية مركزة.", "إبراز العلاقات المنطقية بين جميع عناصر الحصة.", "تسهيل المراجعة السريعة استعداداً للاختبارات الدورية."]),
        ("25. المهام الأدائية المعتمدة 1448هـ", ["تحديد تفاصيل المهمة الأدائية ومعايير تسليمها بدقة.", "ربط المهمة بنواتج التعلم المستهدفة وسوق العمل.", "استمارة تقييم معيارية لمخرجات الطلاب الفردية والجماعية."]),
        ("26. الواجبات المنزلية والأنشطة الذاتية", ["تحديد تمارين الكتاب المدرسي المعتمدة بدقة.", "إسناد الواجب الإلكتروني عبر منصة مدرستي.", "أنشطة استقصائية اختيارية لتعزيز التعلم الذاتي."]),
        ("27. التغذية الراجعة وخطة التطوير", ["إرشادات تفصيلية موجهة لكل طالب لتحسين مستواه.", "تعزيز نقاط القوة ووضع خطة علاجية لنقاط الضعف.", "تمكين الطالب من مهارات التقييم الذاتي والتأمل المعرفي."]),
        ("28. مؤشرات تحقق نواتج التعلم 1448هـ", ["قياس نسبة إتقان الطلاب للأهداف مقارنة بالمستهدف.", "تحليل نتائج الاختبارات القصيرة والأنشطة الصفية.", "توثيق مؤشرات الأداء لدعم ملف تقييم الأداء الوظيفي."]),
        ("29. التوصيات والإرشادات للطلاب والمعلمين", ["إرشادات للمذاكرة الفعالة والتعامل مع الاختبارات الوطنية.", "توجيهات للمعلم لتطوير استراتيجيات التدريس الحديثة.", "تفعيل قنوات الشراكة الإيجابية مع الأسرة وأولياء الأمور."]),
        ("30. الخاتمة والمراجع المعتمدة وشكر وتقدير", ["المراجع: المقررات الرسمية لوزارة التعليم 1448هـ، بوابة عين، منصة مدرستي.", "شكر وتقدير لكافة الطلاب والطالبات والمعلمين المتميزين.", "فتح المجال للمناقشة وتلقي الأسئلة الختامية."])
    ]

def create_powerpoint_presentation_1448(topic: str, output_path: str):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slides_data = generate_curriculum_1448_slides(topic)

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
        p_title.font.size = Pt(28)
        p_title.font.bold = True
        p_title.font.color.rgb = RGBColor(27, 73, 101)

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

        img_stream = fetch_unique_slide_image(idx)
        if img_stream:
            slide.shapes.add_picture(img_stream, Inches(0.8), Inches(1.5), Inches(4.3), Inches(5.1))

        footer_box = slide.shapes.add_textbox(Inches(0.8), Inches(6.9), Inches(11.733), Inches(0.4))
        p_foot = footer_box.text_frame.paragraphs[0]
        p_foot.text = f"شريحة {idx} من 30 | المنهج السعودي 1448هـ - {topic[:30]}"
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
        "svc_exam": f"اكتب نموذجاً لاختبار شامل ومفصل مع جدول المواصفات وتحليل النتائج للمنهج السعودي 1448هـ لـ: '{topic}'. يشمل: أسئلة مقالية وموضوعية، نموذج الإجابة، توزيع الدرجات، وآلية تحليل النتائج واستخراج نسب النجاح.",
        "svc_remedial": f"اكتب خطة علاجية وإثرائية متكاملة لعام 1448هـ لـ: '{topic}'. تشمل: تشخيص الفاقد والضعف، أوراق عمل وأنشطة داعمة، خطة المتفوقين والإثرائيات، واستمارة متابعة دورية.",
        "svc_portfolio": f"اكتب هيكل ومحتوى ملف الإنجاز الإلكتروني للمعلم/المعلمة لعام 1448هـ في: '{topic}'. يشمل: الفلسفة التعليمية، ميثاق الأخلاقيات، شواهد التطوير المهني، استراتيجيات التدريس الموجه، ونماذج أعمال الطلاب.",
        "svc_performance": f"اكتب ميثاق وخطة ملف الأداء الوظيفي الجديد المعتمد لعام 1448هـ لـ: '{topic}'. يشمل: الأهداف الوظيفية المحددة، معايير القياس (مؤشرات الأداء)، الوزن النسبي، وجدول الشواهد والإثباتات.",
        "svc_operation": f"اكتب الخطة التشغيلية السنوية/الفصلية لعام 1448هـ لـ: '{topic}'. تشمل: الأهداف العامة والتفصيلية، البرامج والمبادرات، المسؤول عن التنفيذ، مؤشرات التحقق، والجدول الزمني.",
        "svc_loss": f"اكتب خطة وملف الفاقد التعليمي لعام 1448هـ لـ: '{topic}'. تشمل: حصر المهارات الأساسية المفقودة، الاختبار التشخيصي القبلي، استراتيجيات التعويض المكثف، والتقويم البعدي لقياس الأثر.",
        "svc_research": f"اكتب بحثاً أكاديمياً جامعياً مفصلاً لمستوى الماجستير حول: '{topic}' يشمل: المقدمة، مشكلة البحث، الأهداف، الإطار النظري، التحليل الميداني، التوصيات، والمراجع وفق APA."
    }

    titles = {
        "svc_exam": f"كتابة الاختبارات وتحليل النتائج (1448هـ)\n{topic}",
        "svc_remedial": f"الخطة العلاجية والإثرائية المتكاملة (1448هـ)\n{topic}",
        "svc_portfolio": f"ملف الإنجاز الإلكتروني الشامل (1448هـ)\n{topic}",
        "svc_performance": f"ميثاق وملف الأداء الوظيفي الجديد (1448هـ)\n{topic}",
        "svc_operation": f"الخطة التشغيلية المعتمدة (1448هـ)\n{topic}",
        "svc_loss": f"ملف وخطة معالجة الفاقد التعليمي (1448هـ)\n{topic}",
        "svc_research": f"بحث أكاديمي متكامل\n{topic}"
    }

    prompt = prompts.get(service_code, f"اكتب وثيقة تعليمية وإدارية مفصلة لعام 1448هـ حول: {topic}")
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
            if b_clean.startswith("#") or any(b_clean.startswith(f"{i}.") for i in range(1, 15)) or "المحور" in b_clean or "الهدف" in b_clean or "الخاتمة" in b_clean:
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
        p.add_run(f"تم إعداد هذا المستند التعليمي المخصص لموضوع ({topic}) وفق الخطة الدراسية المعتمدة لعام 1448هـ.")

    doc.save(output_path)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [
        [InlineKeyboardButton("📊 عروض بوربوينت المناهج 1448هـ (30 شريحة + صور)", callback_data="svc_ppt")],
        [InlineKeyboardButton("📝 كتابة الاختبارات وجداول المواصفات 1448هـ", callback_data="svc_exam")],
        [InlineKeyboardButton("📈 خطط علاجية وإثرائية 1448هـ", callback_data="svc_remedial")],
        [InlineKeyboardButton("🗂 ملف إنجاز المعلم/المعلمة الإلكتروني 1448هـ", callback_data="svc_portfolio")],
        [InlineKeyboardButton("📑 ميثاق وملف الأداء الوظيفي الجديد 1448هـ", callback_data="svc_performance")],
        [InlineKeyboardButton("📅 الخطة التشغيلية المدرسية والفصلية 1448هـ", callback_data="svc_operation")],
        [InlineKeyboardButton("📚 ملف وخطة معالجة الفاقد التعليمي 1448هـ", callback_data="svc_loss")],
        [InlineKeyboardButton("🎓 إعداد بحث جامعي وأكاديمي متكامل (Word)", callback_data="svc_research")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    welcome_text = (
        "🌟 **أهلاً بك في منصة إنجاز للخدمات التعليمية والأكاديمية - المنهج السعودي 1448هـ**\n\n"
        "✨ نقدم لكم الباقة الشاملة المعتمدة للعام الدراسي 1448هـ لجميع المراحل (ابتدائي، متوسط، مسارات الثانوي، والجامعة):\n\n"
        "🔹 **عروض بوربوينت 1448هـ:** 30 شريحة تفصيلية مع صور حصرية غير مكررة ومطابقة لنواتج التعلم.\n"
        "🔹 **مستندات Word منسقة:** ميثاق الأداء الوظيفي الجديد، خطط علاجية، فاقد تعليمي، واختبارات معيارية.\n\n"
        "👇 **اختر الخدمة المطلوبة من القائمة لبدء التوليد فوراً:**"
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
        "svc_ppt": "📊 عروض بوربوينت المناهج 1448هـ (30 شريحة)",
        "svc_exam": "📝 كتابة الاختبارات وتحليل النتائج 1448هـ",
        "svc_remedial": "📈 خطط علاجية وإثرائية 1448هـ",
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
                "📊 أرسل الآن **المادة والصف والموضوع لمنهج 1448هـ**\n"
                "(مثال: *تقنية رقمية 1-1 مسارات - الذكاء الاصطناعي* أو *علوم سادس ابتدائي - الخلايا* أو *فيزياء 3 مسارات - الكهرومغناطيسية*)\n\n"
                "وسيقوم البوت بتوليد ملف بوربوينت احترافي يتكون من **30 شريحة كاملة مع صور ورسومات بيانية حصرية**:"
            )
        else:
            await query.edit_message_text(f"✨ خدمة: *{name}*\n\nيرجى إرسال تفاصيل المادة أو الموضوع وسأقوم بتوليد **ملف Word (.docx) رسمي ومنسق بالكامل لعام 1448هـ** فوراً:", parse_mode="Markdown")

async def text_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_text = update.message.text
    user = update.effective_user
    current_service = context.user_data.get("current_service", "svc_ppt")
    service_name = context.user_data.get("service_name", "عرض بوربوينت 1448هـ")

    status_msg = await update.message.reply_text("⏳ جارٍ إعداد وتنسيق المحتوى وفق المنهج السعودي 1448هـ...")

    try:
        if current_service == "svc_ppt":
            file_name = f"presentation_{user.id}.pptx"
            create_powerpoint_presentation_1448(user_text, file_name)

            with open(file_name, "rb") as ppt_file:
                await update.message.reply_document(
                    document=ppt_file,
                    filename=f"{user_text[:30]}_1448H_30Slides.pptx",
                    caption=f"✅ **تم إعداد العرض التقديمي بنجاح:**\n📌 *{user_text}*\n📊 **عدد الشرائح:** 30 شريحة كاملة وفق منهج 1448هـ مع صور حصرية غير مكررة.",
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
                    caption=f"✅ **تم تجهيز المستند بنجاح عبر منصة إنجاز:**\n📌 الخدمة: *{service_name}*\n📝 الموضوع: *{user_text}*\n📄 تم تصدير الملف كـ مستند Word منسق وفق متطلبات عام 1448هـ.",
                    parse_mode="Markdown"
                )
            if os.path.exists(file_name):
                os.remove(file_name)

        await status_msg.delete()

    except Exception as e:
        await status_msg.edit_text(f"⚠️ حدث خطأ أثناء المعالجة: {str(e)}")

async def handle_ping(request):
    return web.Response(text="Enjaz 1448 Bot is active!")

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
    
    print("منصة إنجاز 1448هـ تعمل بنجاح...")
    await app.initialize()
    await app.start()
    await app.updater.start_polling()
    
    while True:
        await asyncio.sleep(3600)

def main():
    asyncio.run(main_async())

if __name__ == "__main__":
    main()
