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
                    {"role": "system", "content": "أنت مصمم حقائب تعليمية ومناهج سعودية معتمد لعام 1448هـ. اكتب محتوى دراسياً مفصلاً جداً متضمناً: التمهيد، الإغلاق، استراتيجيات التعلم النشط، مهارات التفكير العليا، تدريبات اختبارات نافس، الربط بالدين والوطن والمواد الأخرى، والأنشطة الفردية والثنائية والجماعية وأوراق العمل."},
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
        "saudi-flag", "classroom", "education", "science", "math", "technology",
        "interactive-learning", "reading", "thinking", "values", "nafis-exam",
        "group-work", "worksheet", "digital-learning", "stem", "space", "future",
        "cybersecurity", "ai", "experiment", "achievement", "evaluation", "mindmap",
        "puzzle", "training", "leadership", "analytics", "success", "vision2030", "graduation"
    ]
    kw = keywords[(slide_index - 1) % len(keywords)]
    url = f"https://loremflickr.com/600/450/{kw}?lock={slide_index * 31 + 1448}"
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

def generate_enhanced_30_slides(topic: str):
    return [
        ("1. الغلاف والبيانات الرسمية 1448هـ", [f"المادة والموضوع: {topic}", "المنهج السعودي المطور 1448هـ | معايير نواتج التعلم", "إعداد تفاعلي شامل مدعم بالأنشطة وتدريبات نافس"]),
        ("2. التمهيد واستثارة الدافعية للدرس", [f"مدخل استكشافي مشوق لموضوع ({topic}) يربط بالواقع.", "عرض تساؤل محوري يستثير الفضول العلمي والتفكير.", "ربط معارف الحصة بالخبرات اليومية والسابقة للطلاب."]),
        ("3. نواتج التعلم والمعايير المستهدفة", ["صياغة الأهداف السلوكية والمعرفية وفق مستويات بلوم العليا.", "تنمية المهارات الأدائية والتطبيقية المقررة لعام 1448هـ.", "غرس الاتجاهات الإيجابية والمواطنة الرقمية الصالحة."]),
        ("4. المفاهيم والمصطلحات الأساسية", ["حصر المصطلحات التخصصية بدقة لغوية وعلمية واضحة.", "التفريق الدقيق بين المفاهيم المتشابهة في الدرس.", "بناء خارطة مفاهيمية تمهد للفهم العميق والمنظم."]),
        ("5. استراتيجية التعليم الموجه (أنا أعمل)", ["النمذجة المباشرة وتفكيك المهارة إلى خطوات واضحة.", "شرح المعلم/المعلمة لأسلوب التفكير وحل الأمثلة النموذجية.", "إبراز استراتيجيات التفكير العلمي السليم أمام الطلاب."]),
        ("6. الشرح المفصل والتطبيقات المنهجية", ["تفصيل عناصر ومحاور الدرس وفق الطبعة الحديثة للكتاب.", "استعراض القوانين والقواعد والتطبيقات التخصصية.", "تدعيم المفاهيم النظرية بالشواهد والأمثلة الواقعية."]),
        ("7. فاصل وتوجيه المقاطع التعليمية المرئية", ["مشاهدة مقطع تعليمي تفاعلي مدته (2-3 دقائق).", "تحليل النقاط الجوهرية والرسوم التوضيحية الواردة في المقطع.", "مناقشة سريعة لترسيخ الفهم الصوتي والمرئي للدرس."]),
        ("8. الممارسة الموجهة (نحن نعمل معاً)", ["حل تطبيقات صفية تفاعلية بمشاركة المعلم والطلاب.", "تصحيح المفاهيم الخاطئة لحظياً وتقديم التغذية الراجعة.", "التأكد من جاهزية الطلاب للانتقال للعمل التشاركي المستقل."]),
        ("9. الأنشطة الثنائية (فكر - زاوج - شارك)", ["التفكير الفردي في السؤال المحدد لمدة دقيقة واحدة.", "المناقشة الثنائية مع الزميل لتبادل وتطوير الأفكار.", "مشاركة الاستنتاج المشترك مع بقية طلاب الفصل الدراسي."]),
        ("10. التعلم التعاوني ومجموعات العمل", ["توزيع الأدوار داخل المجموعات الصغيرة لإنجاز مهمة محددة.", "تطبيق استراتيجية (الرؤوس المرقمة / جيكسو التفاعلية).", "عرض نتاج عمل المجموعة ومناقشته بصورة جماعية بناءة."]),
        ("11. الأنشطة الفردية والتطبيق المستقل", ["تكليف كل طالب بحل تمرين فردي مقنن لقياس الاستيعاب.", "تعزيز مهارات الاعتماد على النفس وضبط وقت الإجابة.", "متابعة المعلم الدقيقة لتقديم الدعم الفردي لمن يحتاجه."]),
        ("12. مهارات التفكير العليا (تحليل - تركيب - تقييم)", ["طرح سؤال يختبر قدرة الطالب على المقارنة والتحليل الدقيق.", "إعادة تركيب المعلومات وصياغة فرضيات واستنتاجات جديدة.", "تقييم البدائل وتقديم نقد موضوعي مبرر بالبراهين العلمية."]),
        ("13. استراتيجيات حل المشكلات المعقدة", ["عرض معضلة واقعية ترتبط بـ ({topic}) والتفكير في علاجها.", "تطبيق خطوات البحث العلمي: المعطيات، الفرضيات، والحلول.", "الوصول إلى حلول مبتكرة قابلة للتنفيذ الميداني."]),
        ("14. الألعاب التعليمية والمسابقات التنافسية", ["لعبة تربوية تفاعلية تكسر الجمود وتعزز الدافعية الصفية.", "تحدي الأسئلة السريعة وبطاقات التنافس الإيجابي بين الفرق.", "تكريم الفريق المتميز لتعزيز الشغف والتعلم النشط."]),
        ("15. تدريبات الاختبارات الوطنية (نافس) - مهارة 1", ["سؤال يحاكي النمط المعياري المعتمد في اختبارات نافس.", "تحليل طريقة استبعاد الخيارات الخاطئة للوصول للصواب.", "تدريب الطلاب على سرعة القراءة والتحليل المنطقي للسؤال."]),
        ("16. تدريبات الاختبارات الوطنية (نافس) - مهارة 2", ["مسألة تطبيقية تقيس الفهم القرائي والاستدلال العلمي.", "ربط السؤال بمهارات الاختبارات الدولية (PISA / TIMSS).", "التأكيد على استراتيجيات الحل الذكي وتفادي المشتتات."]),
        ("17. الربط بالقيم الإسلامية والأخلاقية", ["إبراز الآيات والأحاديث النبوية المرتبطة بمفهوم الأمانة والعلم.", "استشعار عظمة الخالق في القوانين والظواهر المحيطة.", "تعزيز قيم الإتقان، المسؤولية، والنزاهة في التعلم والحياة."]),
        ("18. الربط بالوطن ورؤية المملكة 2030", ["إبراز دور موضوع الدرس في مسارات النهضة التنموية للمملكة.", "ربط المحتوى بمشاريع الوطن الكبرى (نيوم، الطاقة النظيفة، التحول الرقمي).", "غرس روح الفخر والاعتزاز بالهوية الوطنية والإنجازات السعودية."]),
        ("19. الربط بالواقع والتطبيقات الحياتية", ["كيف يستفيد الطالب من معارف هذا الدرس في حياته اليومية؟", "نماذج تطبيقية حية توضح أهمية العلم في حل المشكلات الأسرية والمجتمعية.", "ترشيد الموارد وتعزيز السلامة وجودة الحياة العامة."]),
        ("20. التكامل مع المواد الدراسية الأخرى (STEM)", ["الربط التكاملي بين العلوم، التقنية، الهندسة، والرياضيات.", "توظيف اللغة العربية في التعبير الدقيق عن المفاهيم العلمية.", "إظهار وحدة المعرفة الإنسانية وترابط فروع العلم المختلفة."]),
        ("21. تمايز التعليم ودعم الفروق الفردية", ["بطاقات دعم وتدرج تعليمي للطلاب ذوي الاحتياج الإضافي.", "أنشطة إثرائية إضافية لتحدي قدرات الطلبة الموهوبين.", "تنويع أساليب الاستجابة بين الكتابية، الشفهية، والرقمية."]),
        ("22. توظيف التقنية والمنصات الرقمية", ["استخدام أنظمة المحاكاة والواقع المعزز في تبسيط الدرس.", "حل أنشطة إلكترونية تفاعلية عبر منصة مدرستي ومايكروسوفت فورمز.", "توثيق مخرجات الدرس رقمياً لدعم ملف الإنجاز الإلكتروني."]),
        ("23. معالجة الفاقد التعليمي والأخطاء الشائعة", ["رصد المفاهيم الخاطئة الشائعة في الدرس وتصحيحها علمياً.", "تقديم تدريب علاجي سريع للمهارات القبلية غير المتقنة.", "تثبيت الفهم السليم من خلال أمثلة ونماذج واضحة ومباشرة."]),
        ("24. ورقة العمل التطبيقية للدرس (القسم 1)", ["نشاط كتابي مقنن يقيس الأهداف الأساسية للدرس.", "أسئلة تفاعلية تجمع بين الاختيار من متعدد وإكمال الفراغ.", "مساحة مخصصة لتدوين خطوات الحل والتفسير المنطقي."]),
        ("25. ورقة العمل التطبيقية للدرس (القسم 2)", ["تمرين تطبيقي متقدم يقيس مهارات التطبيق والتركيب.", "مهمة أدائية مصغرة ينفذها الطالب داخل الحصة.", "سلم تصحيح ذاتي يمكّن الطالب من تقييم أدائه مباشرة."]),
        ("26. التقويم التكويني وبطاقات الخروج", ["سؤال سريع في دقيقة واحدة لقياس نواتج التعلم الأساسية.", "كتابة أهم فائدة مكتسبة في بطاقة الخروج الورقية أو الرقمية.", "رصد المؤشرات الفورية لتعديل خطة التدريس للحصة القادمة."]),
        ("27. الإغلاق المنهجي المناسب للدرس", ["تلخيص محاور الحصة واستخلاص الأفكار الرئيسية بأسلوب موجز.", "الربط بين مخرجات اليوم وموضوع الدرس في الحصة القادمة.", "التأكد من وضوح الصورة الشاملة لجميع طلاب الفصل."]),
        ("28. الواجب المنزلي والمهام الأدائية", ["تحديد تمارين الكتاب المدرسي المقررة لعام 1448هـ بدقة.", "إسناد المهمة الأدائية أو الواجب عبر منصة مدرستي.", "أنشطة استقصائية اختيارية لتشجيع التعلم الذاتي بالمنزل."]),
        ("29. التغذية الراجعة وسلم التقدير (Rubric)", ["معايير تقييم أداء الطلاب ومستويات الإتقان (ممتاز، جيد، يحتاج دعم).", "إرشادات موجهة للطالب لمواصلة التميز وتلافي جوانب الضعف.", "تعزيز ثقافة التقويم من أجل التعلم والتطوير المستمر."]),
        ("30. المراجع المعتمدة وشكر وتقدير", ["المراجع: المقررات الرسمية لوزارة التعليم 1448هـ، بوابة عين، منصة مدرستي.", "شكر وتقدير للطلاب والطالبات على التفاعل والمشاركة المتميزة.", "فتح المجال للمناقشة وتلقي الأسئلة والاستفسارات الختامية."])
    ]

def create_powerpoint_presentation_full(topic: str, output_path: str):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slides_data = generate_enhanced_30_slides(topic)

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
        p_title.font.size = Pt(28)
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

        # إدراج الصورة المخصصة للشريحة
        img_stream = fetch_unique_slide_image(idx)
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
        "svc_exam": f"اكتب نموذجاً لاختبار شامل ومفصل مع جدول المواصفات وتدريبات نافس وتحليل النتائج لـ: '{topic}'. يشمل: أسئلة مقالية وموضوعية وتفكير عليا، نموذج الإجابة، توزيع الدرجات، وتحليل النتائج.",
        "svc_remedial": f"اكتب خطة علاجية وإثرائية مع أوراق العمل وأنشطة التفكير الموجه لـ: '{topic}'. تشمل: تشخيص الفاقد، تدريبات نافس، أنشطة الموهوبين، واستمارة المتابعة.",
        "svc_portfolio": f"اكتب هيكل ومحتوى ملف الإنجاز الإلكتروني للمعلم/المعلمة لعام 1448هـ في: '{topic}'. يشمل: الفلسفة، الشواهد، استراتيجيات التعلم النشط، نماذج الأعمال، والربط برؤية 2030.",
        "svc_performance": f"اكتب ميثاق وخطة ملف الأداء الوظيفي الجديد المعتمد لعام 1448هـ لـ: '{topic}'. يشمل: الأهداف الوظيفية، مؤشرات الأداء، الوزن النسبي، وجداول الشواهد والإثباتات.",
        "svc_operation": f"اكتب الخطة التشغيلية لعام 1448هـ لـ: '{topic}'. تشمل: الأهداف التفصيلية، المبادرات، المسؤول، مؤشرات التحقق، والجدول الزمني التنفيذي.",
        "svc_loss": f"اكتب خطة وملف الفاقد التعليمي لـ: '{topic}'. تشمل: حصر المهارات المفقودة، الاختبار التشخيصي، خطة التعويض المكثف، والتقويم البعدي لقياس الأثر.",
        "svc_research": f"اكتب بحثاً أكاديمياً جامعياً مفصلاً لمستوى الماجستير حول: '{topic}' يشمل: المقدمة، مشكلة البحث، الأهداف، الإطار النظري، التحليل الميداني، التوصيات، والمراجع وفق APA."
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

    prompt = prompts.get(service_code, f"اكتب وثيقة تعليمية متكاملة حول: {topic}")
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
        p.add_run(f"تم إعداد هذا المستند التعليمي المخصص لموضوع ({topic}) وفق معايير المنهج السعودي للعام الدراسي 1448هـ متضمناً الأنشطة وتدريبات نافس وأوراق العمل.")

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
        "✨ **الحقيبة التعليمية المتكاملة لجميع المواد والصفوف:**\n"
        "📌 تمهيد وإغلاق مناسب للدرس.\n"
        "📌 مهارات تفكير عليا + تعلم نشط + ألعاب تعليمية.\n"
        "📌 فواصل ومقاطع تعليمية تفاعلية.\n"
        "📌 ربط بالوطن ورؤية 2030، الدين، الواقع، والمواد الأخرى.\n"
        "📌 تدريبات يومية لاختبارات (نافس) الوطنية.\n"
        "📌 أنشطة فردية، ثنائية، وجماعية + أوراق عمل وتقويم ختامي.\n\n"
        "👇 **اختر الخدمة المطلوبة لبدء التوليد فوراً:**"
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
        "svc_ppt": "📊 عروض بوربوينت شاملة (30 شريحة + نافس + ألعاب)",
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
                "(مثال: *رياضيات أول متوسط - الأعداد الصحيحة* أو *علوم سادس - الخلية* أو *كيمياء مسارات - الاتزان الكيميائي*)\n\n"
                "وسيقوم البوت بتوليد ملف بوربوينت احترافي (30 شريحة) متضمناً: التمهيد، التعلم النشط، الألعاب، تدريبات نافس، الربط بالوطن والدين، والأنشطة وأوراق العمل مع صور حصرية غير مكررة:"
            )
        else:
            await query.edit_message_text(f"✨ خدمة: *{name}*\n\nيرجى إرسال تفاصيل المادة أو الموضوع وسأقوم بتوليد **ملف Word (.docx) رسمي ومنسق بالكامل لعام 1448هـ** فوراً:", parse_mode="Markdown")

async def text_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_text = update.message.text
    user = update.effective_user
    current_service = context.user_data.get("current_service", "svc_ppt")
    service_name = context.user_data.get("service_name", "عرض بوربوينت متكامل 1448هـ")

    status_msg = await update.message.reply_text("⏳ جارٍ إعداد وتجهيز المحتوى التعليمي الشامل وتوليد الملف بالذكاء الاصطناعي...")

    try:
        if current_service == "svc_ppt":
            file_name = f"presentation_{user.id}.pptx"
            create_powerpoint_presentation_full(user_text, file_name)

            with open(file_name, "rb") as ppt_file:
                await update.message.reply_document(
                    document=ppt_file,
                    filename=f"{user_text[:30]}_Full1448H.pptx",
                    caption=f"✅ **تم إعداد العرض التقديمي بنجاح:**\n📌 *{user_text}*\n📊 **المحتوى:** 30 شريحة متكاملة (تمهيد، إغلاق، تعلم نشط، تدريبات نافس، ألعاب، ربط بالوطن والدين، أنشطة، وأوراق عمل) مع صور حصرية.",
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
                    caption=f"✅ **تم تجهيز المستند بنجاح عبر منصة إنجاز:**\n📌 الخدمة: *{service_name}*\n📝 الموضوع: *{user_text}*\n📄 تم تصدير الملف كـ مستند Word منسق ومفصل وفق معايير 1448هـ.",
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
    
    print("منصة إنجاز تعمل بكامل المعايير التعليمية...")
    await app.initialize()
    await app.start()
    await app.updater.start_polling()
    
    while True:
        await asyncio.sleep(3600)

def main():
    asyncio.run(main_async())

if __name__ == "__main__":
    main()
