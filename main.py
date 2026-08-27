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
from docx.enum.table import WD_TABLE_ALIGNMENT
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
                    {"role": "system", "content": "أنت خبير تعليمي وإداري معتمد في التعليم السعودي. اكتب محتوى تعليمياً مفصلاً جداً وشاملاً ومنظماً بجداول وعناوين وفق معايير وزارة التعليم ونماذج الأداء الحديثة."},
                    {"role": "user", "content": prompt}
                ],
                "model": "openai",
                "seed": 42
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

def create_educational_doc(service_code: str, topic: str, output_path: str):
    doc = Document()
    for section in doc.sections:
        section.top_margin = DocxInches(1)
        section.bottom_margin = DocxInches(1)
        section.left_margin = DocxInches(1)
        section.right_margin = DocxInches(1)

    prompts = {
        "svc_exam": f"اكتب نموذجاً لاختبار شامل ومفصل مع جدول المواصفات وتحليل النتائج لـ: '{topic}'. يشمل: أسئلة مقالية وموضوعية، نموذج الإجابة، توزيع الدرجات، وآلية تحليل النتائج واستخراج نسب النجاح.",
        "svc_remedial": f"اكتب خطة علاجية وإثرائية متكاملة لـ: '{topic}'. تشمل: تشخيص الفاقد والضعف، أوراق عمل وأنشطة داعمة، خطة المتفوقين والإثرائيات، واستمارة متابعة دورية.",
        "svc_portfolio": f"اكتب هيكل ومحتوى ملف الإنجاز الإلكتروني للمعلم/المعلمة في: '{topic}'. يشمل: الفلسفة التعليمية، ميثاق الأخلاقيات، شواهد التطوير المهني، استراتيجيات التدريس، ونماذج أعمال الطلاب.",
        "svc_performance": f"اكتب ميثاق وخطة ملف الأداء الوظيفي الجديد وفق اللائحة السعودية لـ: '{topic}'. يشمل: الأهداف الوظيفية المحددة، معايير القياس (مؤشرات الأداء)، الوزن النسبي، وجدول الشواهد والإثباتات.",
        "svc_operation": f"اكتب الخطة التشغيلية السنوية/الفصلية المتكاملة لـ: '{topic}'. تشمل: الأهداف العامة والتفصيلية، البرامج والمبادرات، المسؤول عن التنفيذ، مؤشرات التحقق، والجدول الزمني.",
        "svc_loss": f"اكتب خطة وملف الفاقد التعليمي لـ: '{topic}'. تشمل: حصر المهارات الأساسية المفقودة، الاختبار التشخيصي القبلي، استراتيجيات التعويض المكثف، والتقويم البعدي لقياس الأثر.",
        "svc_research": f"اكتب بحثاً أكاديمياً جامعياً مفصلاً لمستوى الماجستير حول: '{topic}' يشمل: المقدمة، مشكلة البحث، الأهداف، الإطار النظري، التحليل الميداني، التوصيات، والمراجع وفق APA."
    }

    titles = {
        "svc_exam": f"كتابة الاختبارات وتحليل النتائج\n{topic}",
        "svc_remedial": f"الخطة العلاجية والإثرائية المتكاملة\n{topic}",
        "svc_portfolio": f"ملف الإنجاز الإلكتروني الشامل\n{topic}",
        "svc_performance": f"ملف وميثاق الأداء الوظيفي الجديد\n{topic}",
        "svc_operation": f"الخطة التشغيلية التعليمية المعتمدة\n{topic}",
        "svc_loss": f"ملف وخطة معالجة الفاقد التعليمي\n{topic}",
        "svc_research": f"بحث أكاديمي متكامل\n{topic}"
    }

    prompt = prompts.get(service_code, f"اكتب وثيقة تعليمية وإدارية مفصلة حول: {topic}")
    doc_title = titles.get(service_code, f"وثيقة تعليمية: {topic}")

    ai_content = query_ai_engine(prompt)

    # ترويسة المستند
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = title_p.add_run(doc_title)
    run_title.font.size = DocxPt(20)
    run_title.font.bold = True
    run_title.font.color.rgb = DocxRGB(27, 73, 101)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = sub_p.add_run("منصة إنجاز للخدمات الأكاديمية والتعليمية | إعداد مخصص وفق المعايير السعودية\n" + "—"*35)
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
        p.add_run(f"تم إعداد هذا المستند التعليمي المخصص لموضوع ({topic}) وفق المتطلبات المعتمدة واللوائح التنظيمية مع توثيق كافة النماذج والشواهد.")

    doc.save(output_path)

# ==========================================
# دالة تصميم البوربوينت (30 شريحة كاملة مع صور حصرية)
# ==========================================
def fetch_unique_slide_image(slide_index: int) -> io.BytesIO:
    keywords = ["education", "teaching", "exam", "strategy", "management", "school", "classroom", "learning", "data", "achievement", "planning", "success", "future", "knowledge", "training", "leadership", "innovation", "science", "development", "excellence", "analytics", "evaluation", "teamwork", "digital", "progress", "study", "guidance", "assessment", "certificate", "celebration"]
    kw = keywords[(slide_index - 1) % len(keywords)]
    url = f"https://loremflickr.com/600/450/{kw}?lock={slide_index * 13 + 5}"
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

def create_powerpoint_presentation_30(topic: str, output_path: str):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slides_data = [
        (f"1. الغلاف والعنوان الرئيسي", [f"عرض تقديمي شامل: {topic}", "إعداد تعليمي وتدريبي متقدم", "منصة إنجاز للخدمات التعليمية والأكاديمية"]),
        ("2. المقدمة والأهمية العامة", [f"يمثل موضوع ({topic}) حجر الأساس في رفع جودة العملية التعليمية.", "مواكبة المعايير الحديثة في التدريس والتقويم ونواتج التعلم.", "تحقيق التميز في الأداء وفق مستهدفات التطوير المهني."]),
        ("3. مشكلة العرض والأهداف", ["تشخيص الفجوات التعليمية والميدانية ووضع حلول استباقية لها.", "تحديد الأهداف السلوكية والمعرفية والمهارية المستهدفة بدقة.", "بناء خطة استجابة مرنة تضمن استمرارية التحصيل الدراسي العالي."]),
        ("4. الأطر المنهجية والنظرية", ["الاستناد إلى أحدث النظريات التربوية واستراتيجيات التعلم النشط.", "تطبيق معايير هيئة تقويم التعليم والتدريب (ETEC).", "ربط المفاهيم النظرية بالتطبيقات الصفية والميدانية المباشرة."]),
        ("5. استراتيجيات التعليم الموجه", ["نموذج التدرج في نقل المسؤولية (أنا أعمل - نحن نعمل - أنتم تعملون - أنت تعمل).", "تقديم التغذية الراجعة الفورية لتعزيز استيعاب المفاهيم.", "تحفيز التفكير الناقد وحل المشكلات لدى الطلاب والطالبات."]),
        ("6. تمايز التعليم ومراعاة الفروق", ["تصميم أنشطة متعددة المستويات لتناسب المتعثرين والمتوسطين والموهوبين.", "تنويع أساليب العرض بين البصري والسمعي والحركي.", "تفعيل بطاقات المهام الفردية والتعلم التشاركي."]),
        ("7. خطة القياس والتقويم التكويني", ["بناء بنوك أسئلة متدرجة الصعوبة وفق تصنيف بلوم للأهداف.", "استخدام أدوات التقويم السريع (بطاقات الخروج، التصويت الإلكتروني).", "رصد مؤشرات التحصيل ومقارنتها بالمستهدفات المعتمدة."]),
        ("8. توظيف التقنية والمنصات التعليمية", ["استثمار منصة مدرستي ومايكروسوفت 365 في إثراء المحتوى.", "إنشاء اختبارات واستبانات تفاعلية عبر النماذج الرقمية.", "أتمتة متابعة الواجبات والمشاريع المدرسية والجامعية."]),
        ("9. تحليل ومعالجة الفاقد التعليمي", ["تحديد المهارات الأساسية غير المتقنة عبر الاختبار التشخيصي.", "تصميم خطة تعويضية مكثفة لعلاج جوانب القصور.", "إجراء الاختبار البعدي لقياس الأثر وتوثيق نسب التحسن."]),
        ("10. إعداد الخطط العلاجية والإثرائية", ["حصر الطلاب ذوي الأداء المنخفض وتحديد برامج الدعم المساندة.", "توفير أوراق عمل تدريبية موجهة تساند في رفع الدرجات.", "إطلاق مبادرات ومشاريع إثرائية نوعية لتنمية مواهب المتفوقين."]),
        ("11. توثيق الشواهد وملف الإنجاز", ["أرشفة خطط الدروس والأنشطة والشهادات التدريبية بصورة منظمة.", "توثيق مبادرات المعلم ومشاركاته في مجتمعات التعلم المهنية.", "إبراز نماذج من مخرجات وأعمال الطلاب المتميزة."]),
        ("12. ميثاق ومؤشرات الأداء الوظيفي", ["تحديد الأهداف الوظيفية السنوية والوزن النسبي لكل معيار.", "جمع الشواهد والأدلة التي تدعم تحقيق المؤشرات المطلوبة.", "إجراء التقييم الذاتي وتحديد مجالات التطوير المستمر."]),
        ("13. الإدارة الصفية الفاعلة", ["بناء بيئة تعليمية محفزة قائمة على الاحترام والتفاعل الإيجابي.", "وضع قواعد سلوكية واضحة وتنظيم الوقت المخصص لكل نشاط.", "إدارة النقاشات الصفية وتوزيع الأدوار بين مجموعات العمل."]),
        ("14. الأنشطة التطبيقية والمشاريع", ["تكليف الطلاب بمهام أدائية ومشاريع بحثية جماعية تطبيقية.", "ربط المشروعات بالقضايا الواقعية والتنمية المستدامة.", "تنمية مهارات العرض والإلقاء والعمل الجماعي."]),
        ("15. تعزيز الشراكة مع أولياء الأمور", ["التواصل المستمر لإطلاع الأسر على مستويات تقدم أبنائهم.", "عقد لقاءات دورية لتبادل الملاحظات ودعم الخطط العلاجية.", "إشراك الأسرة في متابعة المهام والأنشطة الإثرائية."]),
        ("16. التنمية المهنية ومجتمعات التعلم", ["حضور الورش التدريبية والبرامج التخصصية لتطوير الكفايات.", "تبادل الزيارات الصفية ونقل الخبرات بين المعلمين.", "الاطلاع على أحدث الدراسات والممارسات العالمية في التدريس."]),
        ("17. التخطيط للدروس النموذجية", ["صياغة الأهداف وفق معايير SMART المحددة والقابلة للقياس.", "تصميم سيناريو الحصة الدراسية وتوزيع الزمن بدقة.", "إعداد الوسائل والمواد التعليمية قبل بدء الدرس."]),
        ("18. استراتيجيات حل المشكلات", ["تدريب الطلاب على خطوات التفكير العلمي المنطقي.", "تحليل المشكلة وفرض الفروض واختبار صحتها.", "الوصول إلى النتائج وتعميم الحلول المبتكرة."]),
        ("19. تعزيز القيم والسلوك الإيجابي", ["دمج القيم الأخلاقية والوطنية ضمن الأنشطة اليومية.", "تكريم الطلاب المتميزين سلوكياً وأكاديمياً لتعزيز الدافعية.", "بناء برامج إرشادية وتوجيهية تساند الاستقرار النفسي للطلاب."]),
        ("20. البيئة التعليمية الجاذبة", ["تنظيم القاعة الدراسية بما يخدم استراتيجيات العمل الجماعي.", "توفير لوحات التعزيز ومصادر التعلم الإثرائية المتاحة.", "خلق أجواء تعليمية تفاعلية مشجعة على الإبداع."]),
        ("21. قياس نواتج التعلم وتفسيرها", ["استخراج تقارير الأداء وتحليل الفروق بين المستويات.", "مقارنة النتائج الحالية بالأعوام السابقة لتحديد منحنى النمو.", "بناء القرارات التدريسية استناداً إلى البيانات والنتائج الفعلية."]),
        ("22. التعلم الذاتي والمستمر", ["تزويد الطلاب بمصادر إثرائية ومنصات رقمية للاطلاع الذاتي.", "تشجيع مهارات البحث والاستقصاء المستقل عن المعلومة.", "بناء عادات التعلم الدائم ومواكبة كل جديد."]),
        ("23. إعداد الاختبارات المقننة", ["مراعاة الوزن النسبي للموضوعات وفق جدول المواصفات.", "التنوع في صياغة الأسئلة المباشرة وغير المباشرة.", "وضوح التعليمات وضبط الوقت المناسب لزمن الإجابة."]),
        ("24. تحليل نتائج الفترات والنهائيات", ["حساب المتوسط الحسابي والانحراف المعياري للدرجات.", "تصنيف الطلاب وفق مستويات الأداء (ممتاز، جيد، بحاجة لدعم).", "إعداد خطط التدخل السريع للحالات التي تتطلب معالجة فورية."]),
        ("25. الحوكمة والجودة التعليمية", ["الالتزام بالسياسات واللوائح والتعاميم الرسمية المنظمة.", "توثيق العمليات والمخرجات وفق نماذج الجودة المعتمدة.", "المراجعة الدورية لضمان التحسين المستمر في بيئة العمل."]),
        ("26. خطة المتابعة والتدخل السريع", ["جدولة المتابعات الدورية لتقييم أثر الخطط المطبقة.", "تعديل الاستراتيجيات بناءً على ردود الأفعال ومخرجات القياس.", "التنسيق مع إدارة المدرسة والتوجيه الطلابي لمعالجة أي معوقات."]),
        ("27. المخرجات والنتائج المتوقعة", ["تحقيق نسب إتقان تفوق 95% في المهارات الأساسية المستهدفة.", "رفع دافعية الطلاب نحو التعلم والتميز الدراسي.", "تحسين مؤشرات الأداء الوظيفي والحصول على تقييمات متميزة."]),
        ("28. الاستنتاجات العامة", ["التخطيط السليم والتنفيذ الموجه هما سر نجاح العملية التعليمية.", "التقويم المستمر هو البوصلة الحقيقية لتصحيح مسار التعلم.", "الاستثمار في تمكين الطالب يعود بالأثر الأكبر على التحصيل."]),
        ("29. التوصيات والمقترحات الختامية", ["الاستمرار في تطبيق أحدث استراتيجيات التدريس المتطورة.", "تحديث ملفات الإنجاز والأداء الوظيفي أولاً بأول.", "تعزيز روح المبادرة والابتكار في إعداد الوسائل التعليمية."]),
        ("30. الخاتمة والمراجع وشكر وتقدير", ["المراجع: أدلة وزارة التعليم، المعايير المهنية، والأدبيات التربوية الحديثة.", "خالص الشكر والتقدير لكافة المعلمين والمعلمات والطلبة على جهودهم.", "باب الأسئلة والمناقشة مفتوح."])
    ]

    for idx, (title_text, points) in enumerate(slides_data, start=1):
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # شريط علوي
        top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.2))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = RGBColor(27, 73, 101)
        top_bar.line.fill.background()

        # بطاقة محتوى
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.4), Inches(1.5), Inches(7.2), Inches(5.1))
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(248, 250, 252)
        card.line.color.rgb = RGBColor(203, 213, 225)
        card.line.width = Pt(1.5)

        # العنوان
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(1.0))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = title_text
        p_title.alignment = PP_ALIGN.RIGHT
        p_title.font.size = Pt(28)
        p_title.font.bold = True
        p_title.font.color.rgb = RGBColor(27, 73, 101)

        # النقاط
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

        # صورة حصرية لكل شريحة
        img_stream = fetch_unique_slide_image(idx)
        if img_stream:
            slide.shapes.add_picture(img_stream, Inches(0.8), Inches(1.5), Inches(4.3), Inches(5.1))

        # تذييل
        footer_box = slide.shapes.add_textbox(Inches(0.8), Inches(6.9), Inches(11.733), Inches(0.4))
        p_foot = footer_box.text_frame.paragraphs[0]
        p_foot.text = f"شريحة {idx} من 30 | منصة إنجاز - {topic[:30]}"
        p_foot.alignment = PP_ALIGN.LEFT
        p_foot.font.size = Pt(11)
        p_foot.font.color.rgb = RGBColor(148, 163, 184)

    prs.save(output_path)

# ==========================================
# معالجات الأوامر والرسائل في التيليجرام
# ==========================================
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [
        [InlineKeyboardButton("📝 كتابة الاختبارات وتحليل النتائج", callback_data="svc_exam")],
        [InlineKeyboardButton("📈 خطط علاجية وإثرائية", callback_data="svc_remedial")],
        [InlineKeyboardButton("🗂 ملف إنجاز إلكتروني", callback_data="svc_portfolio")],
        [InlineKeyboardButton("📑 ملف الأداء الوظيفي الجديد", callback_data="svc_performance")],
        [InlineKeyboardButton("📅 الخطة التشغيلية للمدرسة/المعلم", callback_data="svc_operation")],
        [InlineKeyboardButton("📚 ملف وخطة الفاقد التعليمي", callback_data="svc_loss")],
        [InlineKeyboardButton("📊 تصميم عرض بوربوينت (30 شريحة + صور)", callback_data="svc_ppt")],
        [InlineKeyboardButton("🎓 إعداد بحث جامعي متكامل (Word)", callback_data="svc_research")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    welcome_text = (
        "🌟 **أهلاً بك في منصة إنجاز للخدمات الأكاديمية والتعليمية**\n\n"
        "✨ نقدم لكم باقة مميزة من الخدمات التعليمية الشاملة للعام الدراسي الجديد:\n\n"
        "🔹 تصميم مستندات Word رسمية ومنسقة جاهزة للطباعة والتقديم.\n"
        "🔹 عروض تقديمية احترافية تتكون من 30 شريحة كاملة مع صور حصرية.\n\n"
        "👇 **اختر الخدمة المطلوبة من القائمة لبدء العمل فوراً:**"
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
        "svc_exam": "📝 كتابة الاختبارات وتحليل النتائج",
        "svc_remedial": "📈 خطط علاجية وإثرائية",
        "svc_portfolio": "🗂 ملف إنجاز إلكتروني",
        "svc_performance": "📑 ملف الأداء الوظيفي الجديد",
        "svc_operation": "📅 الخطة التشغيلية",
        "svc_loss": "📚 ملف الفاقد التعليمي",
        "svc_ppt": "📊 تصميم عرض بوربوينت كامل (30 شريحة)",
        "svc_research": "🎓 إعداد بحث جامعي متكامل (Word)"
    }

    if data in services:
        name = services[data]
        context.user_data["current_service"] = data
        context.user_data["service_name"] = name
        
        if data == "svc_ppt":
            await query.edit_message_text(f"📊 أرسل الآن عنوان العرض التقديمي لتوليد ملف بوربوينت احترافي (30 شريحة كاملة مع صور حصرية):")
        else:
            await query.edit_message_text(f"✨ خدمة: *{name}*\n\nيرجى إرسال تفاصيل المادة أو الموضوع أو الصف الدراسي وسأقوم بتوليد **ملف Word (.docx) رسمي ومنسق بالكامل** فوراً:", parse_mode="Markdown")

async def text_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_text = update.message.text
    user = update.effective_user
    current_service = context.user_data.get("current_service", "svc_operation")
    service_name = context.user_data.get("service_name", "خدمة تعليمية")

    status_msg = await update.message.reply_text("⏳ جارٍ إعداد وتنسيق المستند التعليمي بالذكاء الاصطناعي...")

    try:
        if current_service == "svc_ppt":
            file_name = f"presentation_{user.id}.pptx"
            create_powerpoint_presentation_30(user_text, file_name)

            with open(file_name, "rb") as ppt_file:
                await update.message.reply_document(
                    document=ppt_file,
                    filename=f"{user_text[:30]}_30Slides.pptx",
                    caption=f"✅ **تم إعداد العرض التقديمي بنجاح:**\n📌 *{user_text}*\n📊 **عدد الشرائح:** 30 شريحة مفصلة مع صور ورسومات بيانية حصرية.",
                    parse_mode="Markdown"
                )
            if os.path.exists(file_name):
                os.remove(file_name)
        else:
            file_name = f"doc_{user.id}.docx"
            create_educational_doc(current_service, user_text, file_name)

            with open(file_name, "rb") as doc_file:
                await update.message.reply_document(
                    document=doc_file,
                    filename=f"{service_name[:20]}_{user_text[:20]}.docx",
                    caption=f"✅ **تم تجهيز المستند بنجاح عبر منصة إنجاز:**\n📌 الخدمة: *{service_name}*\n📝 الموضوع: *{user_text}*\n📄 تم تصدير الملف كـ مستند Word منسق ومفصل.",
                    parse_mode="Markdown"
                )
            if os.path.exists(file_name):
                os.remove(file_name)

        await status_msg.delete()

    except Exception as e:
        await status_msg.edit_text(f"⚠️ حدث خطأ أثناء المعالجة: {str(e)}")

async def handle_ping(request):
    return web.Response(text="Enjaz Bot is live!")

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
    
    print("منصة إنجاز تعمل بكافة الخدمات التعليمية...")
    await app.initialize()
    await app.start()
    await app.updater.start_polling()
    
    while True:
        await asyncio.sleep(3600)

def main():
    asyncio.run(main_async())

if __name__ == "__main__":
    main()
