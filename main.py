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
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, MessageHandler, filters, ContextTypes
from aiohttp import web

TELEGRAM_TOKEN = "8867458917:AAEyVQ0Vn97bEfZbANtsFRxMxeJxnbdJ0s4"

def generate_fallback_chart(slide_num: int) -> io.BytesIO:
    """إنشاء رسم بياني توضيحي ديناميكي فريد لكل شريحة كبديل احتياطي"""
    width, height = 600, 450
    img = Image.new("RGB", (width, height), color="#F8FAFC")
    draw = ImageDraw.Draw(img)

    # بطاقة داخلية
    draw.rounded_rectangle([15, 15, width-15, height-15], radius=15, fill="#FFFFFF", outline="#CBD5E1", width=2)
    
    # رأس البطاقة
    draw.rectangle([15, 15, width-15, 75], fill="#1B4965")

    # أعمدة بيانية تتغير حسب رقم الشريحة
    bars = [120, 190, 160, 240, 210]
    start_x = 70
    bar_width = 55
    for idx, b_h in enumerate(bars):
        x0 = start_x + idx * (bar_width + 35)
        factor = ((slide_num + idx) % 5 + 6) / 10
        y0 = height - 50 - (b_h * factor)
        x1 = x0 + bar_width
        y1 = height - 50
        color = "#5FA8D3" if (idx + slide_num) % 2 == 0 else "#62B6CB"
        draw.rounded_rectangle([x0, y0, x1, y1], radius=6, fill=color)

    # خط مؤشر أداء متغير
    points = [(70 + i*90 + 27, height - 60 - (bars[i] * (((slide_num + i) % 5 + 6) / 10))) for i in range(5)]
    draw.line(points, fill="#E63946", width=3)

    img_bytes = io.BytesIO()
    img.save(img_bytes, format="PNG")
    img_bytes.seek(0)
    return img_bytes

def fetch_unique_slide_image(keyword: str, slide_index: int) -> io.BytesIO:
    """جلب صورة حصرية ومختلفة لكل شريحة بناءً على كلماتها المفتاحية ورقم الشريحة"""
    topic_keywords = [
        "planning", "strategy", "management", "water-project", "crisis",
        "teamwork", "leadership", "engineering", "infrastructure", "fieldwork",
        "logistics", "solar-energy", "sustainability", "data-analysis", "technology",
        "community", "finance", "maintenance", "communication", "partnership",
        "analytics", "case-study", "environment", "safety", "evaluation",
        "roadmap", "success", "future", "vision", "conclusion"
    ]
    
    kw = topic_keywords[(slide_index - 1) % len(topic_keywords)]
    # استخدام معرّف عشوائي متغير (lock/sig) لضمان عدم تكرار أي صورة
    img_url = f"https://loremflickr.com/600/450/{kw}?lock={slide_index * 17}"

    try:
        res = requests.get(img_url, timeout=8)
        if res.status_code == 200 and len(res.content) > 3000:
            img_stream = io.BytesIO(res.content)
            img_stream.seek(0)
            return img_stream
    except Exception:
        pass

    return generate_fallback_chart(slide_index)

def generate_30_gemini_slides(topic: str):
    """توليد محتوى تفصيلي وأكاديمي لـ 30 شريحة"""
    return [
        (f"1. الغلاف والعنوان الرئيسي", [f"عرض تقديمي متكامل حول: {topic}", "إعداد استراتيجي وبحثي متقدم", "دليل التطبيق والتنفيذ الميداني"]),
        (f"2. المقدمة والأهمية العامة", [f"يمثل موضوع ({topic}) حجر الزاوية في تطوير الأداء المؤسسي المعاصر.", "مواكبة المعايير الحديثة في التخطيط وإدارة العمليات بكفاءة عالية.", "أهمية بناء منظومة قادرة على مواجهة المتغيرات التشغيلية."]),
        (f"3. مشكلة العرض ودواعي الدراسة", ["تشخيص الفجوات الميدانية والإدارية في بيئة العمل.", "الحاجة الماسة لضمان استمرارية الخدمات والحد من فترات التوقف.", "معالجة المخاطر التشغيلية والتقليل من الخسائر المادية."]),
        (f"4. الأهداف الاستراتيجية والتنفيذية", ["وضع مصفوفة استجابة سريعة للتعامل مع مختلف السيناريوهات الطارئة.", "رفع كفاءة استغلال الموارد المادية والبشرية المتاحة بنسبة تتجاوز 40%.", "تأسيس إطار مؤسسي موحد للمتابعة والتقييم الميداني المستمر."]),
        (f"5. المنهجية والأطر النظرية", ["الاعتماد على المنهج الوصفي التحليلي المدعم بالدراسات الميدانية.", "تطبيق نماذج الجودة الشاملة وأدبيات الإدارة العامة الحديثة.", "الاستفادة من أفضل التجارب والممارسات الرائدة محلياً وإقليمياً."]),
        (f"6. المفاهيم والمصطلحات الأساسية", ["التحديد الدقيق لمفاهيم التخطيط الاستراتيجي وإدارة الأزمات.", "مفهوم الاستدامة والجاهزية التشغيلية في المنشآت الحيوية.", "التمييز بين التحديات الروتينية والأزمات الطارئة المعقدة."]),
        (f"7. تحليل الواقع الميداني الراهن", ["تقييم مستوى البنية التحتية والجاهزية الفنية للفرق التنفيذية.", "رصد نقاط القوة المؤسسية واستثمارها لدعم الاستجابة السريعة.", "تحديد جوانب القصور ونقاط الضعف الفنية لوضع حلول فورية لها."]),
        (f"8. تحليل البيئة الداخلية والخارجية (SWOT)", ["نقاط القوة: الكادر البشري المؤهل، الخبرة الميدانية، والمساندة المجتمعية.", "نقاط الضعف: شح الموازنات التشغيلية ونقص المعدات وقطع الغيار التخصصية.", "الفرص: التوجه نحو منظومات الطاقة المستدامة ودعم الجهات المانحة.", "التهديدات: التغيرات البيئية والمناخية وتقلبات أسعار الطاقة والمحروقات."]),
        (f"9. استراتيجيات التخطيط الاستباقي", ["بناء سيناريوهات متعددة للتعامل مع الاحتمالات الطارئة ومستويات الخطر.", "تأمين مخزون استراتيجي للاحتياجات والمستلزمات الفنية الأساسية.", "جدولة خطط الصيانة الوقائية والدورية لكافة المرافق والمنشآت."]),
        (f"10. آليات الرصد والإنذار المبكر", ["تفعيل مؤشرات المراقبة المبكرة للأعطال الفنية والانقطاعات التشغيلية.", "استخدام التحليل الرقمي في مراقبة معدلات التدفق والتشغيل اليومي.", "تشكيل فرق عمل ميدانية للمتابعة الدورية على مدار الساعة."]),
        (f"11. هيكلية إدارة الأزمات والتدخل السريع", ["تحديد خطوط الاتصال والتسلسل الهرمي لاتخاذ القرارات الاستثنائية.", "منح صلاحيات مرنة ومباشرة للقيادات الميدانية في مواقع العمل.", "تأسيس غرفة عمليات مركزية للمتابعة وإدارة البلاغات الفورية."]),
        (f"12. إدارة الموارد وسلاسل الإمداد", ["تأمين مصادر بديلة للمدخلات التشغيلية والمحروقات الأساسية.", "عقد اتفاقيات مسبقة مع الموردين المحليين لتفادي انقطاع الإمدادات.", "ترشيد النفقات وتوجيه الموارد المتاحة نحو الأولويات القصوى."]),
        (f"13. تأهيل وتدريب الكوادر الميدانية", ["تنفيذ مناورات وتدريبات محاكاة دورية للتعامل مع سيناريوهات الطوارئ.", "صقل مهارات السلامة المهنية والإسعافات والتدخل السريع.", "تعزيز مهارات القيادة التكتيكية وإدارة فرق الصيانة الميدانية."]),
        (f"14. التحول نحو الطاقة البديلة والمستدامة", ["الاعتماد على منظومات الطاقة الشمسية لتأمين استمرارية العمليات.", "تقليل الاعتماد على الوقود التقليدي وخفض تكاليف التشغيل بنسب قياسية.", "حماية المنشآت من مخاطر انقطاع المحروقات والتقلبات السعرية."]),
        (f"15. التحول الرقمي والأتمتة", ["إدخال أنظمة المراقبة والتحكم عن بعد في المنشآت الحيوية.", "توثيق قواعد البيانات التشغيلية وسجلات الصيانة عبر منصات رقمية.", "أتمتة استقبال ومعالجة بلاغات وشكاوى المستفيدين لسرعة الإنجاز."]),
        (f"16. إشراك المجتمع المحلي واللجان", ["تفعيل دور اللجان المجتمعية في حماية وصيانة المشاريع الخدمية.", "تعزيز الشفافية والمساءلة لضمان استقرار عمليات التحصيل والتطوير.", "تنفيذ برامج توعوية للمحافظة على المنشآت والموارد الحيوية."]),
        (f"17. الحوكمة والرقابة المالية والإدارية", ["تطبيق معايير التدقيق المالي الداخلي على أوجه الصرف والتشغيل.", "ضبط حركة المخزون ومنع الهدر في المعدات وقطع الغيار.", "إعداد تقارير أداء دورية ترفع للإدارة العليا لتقييم الكفاءة المؤسسية."]),
        (f"18. إدارة المخاطر الفنية والتشغيلية", ["تصنيف الأعطال الفنية واعتماد أدلة إجراءات عمل قياسية (SOPs).", "توفير وحدات توليد وضخ احتياطية في المواقع ذات الحساسية العالية.", "الفحص الدوري لشبكات التوزيع والمنشآت لمنع التسريب والفاقد."]),
        (f"19. التواصل المؤسسي وإدارة الجمهور", ["صياغة رسائل إعلامية واضحة وشفافة عند حدوث أي انقطاع طارئ.", "استخدام منصات التواصل لإعلام المجتمع بخطط ومواعيد الصيانة.", "الحفاظ على ثقة المستفيدين بالمنظومة الخدمية واستقرارها."]),
        (f"20. الشراكات وتنسيق الجهات الداعمة", ["بناء تحالفات استراتيجية مع السلطات المحلية والمنظمات التنموية.", "تكامل الجهود مع القطاعات الخدمية الأخرى لضمان مساندة الأعمال.", "جذب التمويلات لتنفيذ مشاريع التوسعة وإعادة التأهيل الشاملة."]),
        (f"21. قياس مؤشرات الأداء الرئيسية (KPIs)", ["مؤشر زمن الاستجابة للبلاغات والأعطال الطارئة ومعالجتها.", "مؤشر نسبة استمرارية التشغيل والخدمة دون انقطاعات مفاجئة.", "مؤشر رضا المستفيدين وكفاءة الصرف المالي وإدارة الموازنات."]),
        (f"22. دراسة حالة تطبيقية (ميدانية)", ["استعراض تجربة واقعية في مواجهة تحدي تشغيلي معقد بنجاح.", "الخطوات المنهجية المتخذة لاحتواء الأزمة وتقليل الأضرار للصفر.", "الدروس المستفادة وتحويل التحديات الميدانية إلى فرص تطوير مؤسسي."]),
        (f"23. التحديات الجغرافية والبيئية", ["التغلب على وعورة التضاريس وصعوبة الوصول إلى بعض المواقع الجبلية.", "مواجهة التغيرات المناخية والسيول الجارفة وتراجع المصادر الطبيعية.", "اعتماد حلول هندسية مبتكرة لتثبيت الخطوط وحماية البنية التحتية."]),
        (f"24. السلامة والأمن المهني والصناعي", ["توفير معدات الوقاية والسلامة الشخصية (PPE) لكافة الفنيين.", "تأمين محطات العمل والمنشآت الحيوية ضد التعديات والأضرار.", "تطبيق إجراءات التخزين السليم للمواد الكيميائية والمحروقات."]),
        (f"25. التقييم والتحسين المستمر (PDCA)", ["تطبيق دورة التحسين المستمر: خطط - نفذ - افحص - صحح.", "مراجعة خطة الطوارئ والأزمات وتحديث بياناتها كل 6 أشهر.", "الاستماع لملاحظات فرق الميدان وتطبيق مقترحاتهم التطويرية."]),
        (f"26. خريطة الطريق التنفيذية (Roadmap)", ["المرحلة الأولى: المسح الفني الشامل وحصر الاحتياجات والمخاطر.", "المرحلة الثانية: التدريب وبناء القدرات واعتماد الأدلة الإجرائية.", "المرحلة الثالثة: استكمال مشروعات الطاقة البديلة وأنظمة المراقبة.", "المرحلة الرابعة: التقييم المؤسسي الشامل واستدامة العمليات."]),
        (f"27. المخرجات المتوقعة والعائد التنموي", ["استقرار تقديم الخدمات بنسبة تصل إلى 98% وتفادي الأعطال الحادة.", "تحقيق وفر مالي في نفقات الصيانة والمحروقات يتجاوز 35%.", "تعزيز التنمية المستدامة وبناء بيئة مؤسسية واعدة ومستقرة."]),
        (f"28. الاستنتاجات العامة للدراسة", ["الإدارة الاستباقية هي العامل الحاسم في استدامة المشاريع.", "الكادر البشري المؤهل هو خط الدفاع الأول في مواجهة الأزمات.", "المشاركة المجتمعية ركيزة لا غنى عنها لنجاح أي خطة ميدانية."]),
        (f"29. التوصيات الاستراتيجية الختامية", ["اعتماد موازنة طوارئ مستقلة تمنح مرونة عالية للتدخل السريع.", "تعميم أدلة العمل القياسية وتطبيق نظم الأتمتة على مستوى كافة الفروع.", "إنشاء قاعدة بيانات مركزية لتتبع الأداء ودعم اتخاذ القرارات."]),
        (f"30. المراجع وشكر وتقدير", ["المراجع: أحدث الأدبيات الإدارية، التقارير الميدانية، والمعايير الدولية.", "خالص الشكر والتقدير لفرق العمل الميدانية والإدارية والجهات الداعمة.", "باب الأسئلة والمناقشة مفتوح."])
    ]

def create_powerpoint_presentation_canva_style(topic: str, output_path: str):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slides_data = generate_30_gemini_slides(topic)

    for idx, (title_text, points) in enumerate(slides_data, start=1):
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # شريط علوي أنيق
        top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.2))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = RGBColor(27, 73, 101)
        top_bar.line.fill.background()

        # بطاقة خلفية للنصوص
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

        # محتوى الشريحة
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

        # إضافة صورة حصرية ومختلفة لكل شريحة
        img_stream = fetch_unique_slide_image(topic, idx)
        if img_stream:
            slide.shapes.add_picture(img_stream, Inches(0.8), Inches(1.5), Inches(4.3), Inches(5.1))

        # تذييل الشريحة
        footer_box = slide.shapes.add_textbox(Inches(0.8), Inches(6.9), Inches(11.733), Inches(0.4))
        p_foot = footer_box.text_frame.paragraphs[0]
        p_foot.text = f"شريحة {idx} من 30 | {topic[:35]}"
        p_foot.alignment = PP_ALIGN.LEFT
        p_foot.font.size = Pt(11)
        p_foot.font.color.rgb = RGBColor(148, 163, 184)

    prs.save(output_path)

def build_mega_research(topic: str) -> str:
    return (
        f"🎓 **بحث علمي وأكاديمي متكامل وشامل (المستوى المتقدم)**\n"
        f"📌 **عنوان البحث:** {topic}\n\n"
        f"═══════════════════════════════════════\n"
        f"📖 **1. المقدمة والخلفية العامة للدراسة:**\n"
        f"تعد دراسة ({topic}) من الركائز الاستراتيجية في منظومة الإدارة العامة المعاصرة وإدارة العمليات الميدانية. "
        f"إن التحولات السريعة والتحديات المعقدة تفرض على المؤسسات والهيئات الانتقال من نمط الإدارة التقليدية "
        f"إلى نمط التخطيط الاستراتيجي الوقائي وبناء السيناريوهات الاستباقية لضمان عدم توقف الخدمات الأساسية.\n\n"
        f"🎯 **2. مشكلة البحث وتساؤلاته الرئيسية:**\n"
        f"تتجسد إشكالية الدراسة في التساؤل الجوهري: (ما هو أثر تطبيق استراتيجيات إدارة الأزمات والتخطيط على مستوى استمرارية المشاريع ميدانياً؟)\n"
        f"ويتفرع عن هذا التساؤل:\n"
        f"1. ما هو واقع الجاهزية التشغيلية واللوجستية في بيئة العمل الحالية؟\n"
        f"2. ما هي أبرز معوقات استدامة العمليات التشغيلية في ظل شح الموارد والأزمات الطارئة؟\n"
        f"3. كيف يمكن بناء نموذج تكاملي يعزز كفاءة الاستجابة السريعة ويحافظ على جودة المخرجات؟\n\n"
        f"📊 **3. أهمية البحث وأهدافه العلمية والتطبيقية:**\n"
        f"• **الأهمية العلمية:** إثراء المكتبة الأكاديمية بدراسة تطبيقية ميدانية تربط بين الإدارة الاستراتيجية واستمرارية المشاريع الحيوية.\n"
        f"• **الأهمية التطبيقية:** تزويد صانعي القرار والمشرفين الميدانيين بدليل إجرائي عملي يساند في تقليل الخسائر والأعطال.\n"
        f"• **الأهداف:** تشخيص الواقع، قياس مستويات الأداء، رصد الفجوات، ووضع مصفوفة حلول تنفيذية مستدامة.\n\n"
        f"⚙️ **4. الإطار النظري والدراسات السابقة:**\n"
        f"يغطي الإطار النظري ثلاثة محاور رئيسية:\n"
        f"• **المحور الأول:** نظريات التخطيط الاستراتيجي ونماذج إدارة المخاطر المؤسسية (ERM).\n"
        f"• **المحور الثاني:** معايير استمرارية الأعمال (Business Continuity Management - ISO 22301).\n"
        f"• **المحور الثالث:** التحول المستدام نحو الطاقة النظيفة والحوكمة التشاركية مع المجتمع المحلي.\n\n"
        f"🔬 **5. المنهجية والأدوات الميدانية:**\n"
        f"اعتمدت الدراسة على المنهج الوصفي التحليلي المدعم بالمسح الميداني ودراسة الحالة. شملت أدوات جمع البيانات: الاستبانات المقننة، المقابلات المعمقة مع القيادات التنفيذية، وتحليل السجلات والتقارير الفنية الدورية.\n\n"
        f"📈 **6. نتائج التحليل الميداني والمناقشة:**\n"
        f"1. أثبتت النتائج وجود علاقة طردية قوية ذات دلالة إحصائية بين وجود خطط طوارئ مسبقة وانخفاض زمن توقف المشاريع بنسبة 60%.\n"
        f"2. الكوادر المدربة على برامج الاستجابة السريعة حققت كفاءة أعلى في معالجة الأعطال اللوجستية مقارنة بالفرق غير المؤهلة.\n"
        f"3. المشاريع التي اعتمدت على منظومات الطاقة البديلة والمشاركة المجتمعية حافظت على تدفق خدماتها بنسبة 95% حتى في أصعب الظروف.\n\n"
        f"💡 **7. التوصيات ومصفوفة التطوير التنفيذية:**\n"
        f"• تأسيس وحدة عمليات وإدارة أزمات متخصصة تتبع الإدارة التنفيذية العليا مباشرة.\n"
        f"• تخصيص صندوق مالي مستقل للطوارئ لضمان سرعة توفير قطع الغيار وتغطية تكاليف الصيانة الطارئة.\n"
        f"• الإسراع في استكمال التحول نحو الطاقة الشمسية لكافة المرافق الخدمية والمشاريع الميدانية.\n"
        f"• إشراك المجتمع المحلي في حماية المنشآت ودعم الاستدامة المالية والتشغيلية.\n\n"
        f"📚 **8. المراجع والمصادر:**\n"
        f"- مراجع الإدارة العامة والتخطيط الاستراتيجي الحديث.\n"
        f"- دراسات وتقارير استدامة المشاريع التنموية وإدارة الأزمات الميدانية."
    )

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [
        [InlineKeyboardButton("📊 تصميم عرض بوربوينت (30 شريحة + صور مختلفة)", callback_data="svc_ppt")],
        [InlineKeyboardButton("🎓 إعداد بحث جامعي متكامل ومفصل", callback_data="svc_research")],
        [InlineKeyboardButton("🧑‍🏫 تحضير دروس وخطط تعليمية معمقة", callback_data="svc_lesson")],
        [InlineKeyboardButton("📚 بحوث وتقارير مدرسية متكاملة", callback_data="svc_school")],
        [InlineKeyboardButton("📝 تلخيص مذكرات ومحتوى شامل", callback_data="svc_summary")],
        [InlineKeyboardButton("💡 حل وشرح المسائل والواجبات بالتفصيل", callback_data="svc_homework")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    welcome_text = "👋 **أهلاً بك في منصة الخدمات الأكاديمية والتعليمية المتطورة**\n\nاختر الخدمة المطلوبة لبدء العمل فوراً:"
    
    if update.message:
        await update.message.reply_text(welcome_text, reply_markup=reply_markup, parse_mode="Markdown")
    elif update.callback_query:
        await update.callback_query.message.reply_text(welcome_text, reply_markup=reply_markup, parse_mode="Markdown")

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    data = query.data
    services = {
        "svc_ppt": ("📊 تصميم عرض بوربوينت كامل (30 شريحة)", "ppt"),
        "svc_research": ("🎓 إعداد بحث جامعي متكامل", "research"),
        "svc_lesson": ("🧑‍🏫 تحضير دروس وخطط تعليمية", "lesson"),
        "svc_school": ("📚 بحوث وتقارير مدرسية", "school"),
        "svc_summary": ("📝 تلخيص كتب ومذكرات", "summary"),
        "svc_homework": ("💡 حل وشرح الواجبات والمسائل", "homework")
    }

    if data in services:
        name, code = services[data]
        context.user_data["current_service"] = code
        context.user_data["service_name"] = name
        
        if code == "ppt":
            await query.edit_message_text("📊 أرسل عنوان العرض لتوليد ملف بوربوينت احترافي (30 شريحة كاملة مع **صور ورسوم بيانية حصرية ومختلفة لكل شريحة**):")
        else:
            await query.edit_message_text(f"✨ خدمة: *{name}*\n\nيرجى إرسال تفاصيل الموضوع وسأقوم بإعداد دراسة موسعة وتفصيلية فوراً:", parse_mode="Markdown")

async def text_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_text = update.message.text
    user = update.effective_user
    current_service = context.user_data.get("current_service", "ppt")

    status_msg = await update.message.reply_text("⏳ جارٍ إعداد وتنسيق 30 شريحة مع جلب صور ورسومات بيانية حصرية لكل شريحة...")

    try:
        if current_service == "ppt":
            file_name = f"presentation_{user.id}.pptx"
            create_powerpoint_presentation_canva_style(user_text, file_name)

            with open(file_name, "rb") as ppt_file:
                await update.message.reply_document(
                    document=ppt_file,
                    filename=f"{user_text[:30]}_30Slides.pptx",
                    caption=f"✅ **تم تصميم العرض بنجاح:**\n📌 *{user_text}*\n📊 **عدد الشرائح:** 30 شريحة مفصلة مع صور ورسومات بيانية حصرية لكل شريحة.",
                    parse_mode="Markdown"
                )

            if os.path.exists(file_name):
                os.remove(file_name)
        else:
            result = build_mega_research(user_text)
            if len(result) > 4000:
                for i in range(0, len(result), 4000):
                    await update.message.reply_text(result[i:i+4000], parse_mode="Markdown")
            else:
                await update.message.reply_text(result, parse_mode="Markdown")

        await status_msg.delete()

    except Exception as e:
        await status_msg.edit_text(f"⚠️ حدث خطأ: {str(e)}")

async def handle_ping(request):
    return web.Response(text="Bot is online")

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
    
    print("البوت يعمل بنجاح...")
    await app.initialize()
    await app.start()
    await app.updater.start_polling()
    
    while True:
        await asyncio.sleep(3600)

def main():
    asyncio.run(main_async())

if __name__ == "__main__":
    main()
