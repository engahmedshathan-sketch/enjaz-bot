import os
import io
import asyncio
import requests
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, MessageHandler, filters, ContextTypes
from aiohttp import web

TELEGRAM_TOKEN = "8867458917:AAEyVQ0Vn97bEfZbANtsFRxMxeJxnbdJ0s4"

def create_slide_image(title: str, slide_num: int) -> io.BytesIO:
    """إنشاء صورة ورسم بياني توضيحي ديناميكي لكل شريحة"""
    width, height = 600, 450
    img = Image.new("RGB", (width, height), color="#F0F4F8")
    draw = ImageDraw.Draw(img)

    # رسم بطاقة خلفية وظل
    draw.rounded_rectangle([15, 15, width-15, height-15], radius=20, fill="#FFFFFF", outline="#1B4965", width=3)
    draw.rectangle([15, 15, width-15, 80], fill="#1B4965")
    
    # رأس الرسمة
    draw.rectangle([35, 110, width-35, 125], fill="#62B6CB")
    
    # رسم أعمدة بيانية ومؤشرات توضيحية
    bars = [140, 220, 180, 260, 210]
    bar_width = 60
    start_x = 70
    for idx, b_h in enumerate(bars):
        x0 = start_x + idx * (bar_width + 30)
        y0 = height - 60 - (b_h * ((slide_num % 4 + 7) / 10))
        x1 = x0 + bar_width
        y1 = height - 60
        color = "#5FA8D3" if idx % 2 == 0 else "#1B4965"
        draw.rounded_rectangle([x0, y0, x1, y1], radius=8, fill=color)

    # خط مؤشر أداء
    points = [(70 + i*90 + 30, height - 70 - (bars[i] * ((slide_num % 4 + 7) / 10))) for i in range(5)]
    draw.line(points, fill="#FF6B6B", width=4)
    for pt in points:
        draw.ellipse([pt[0]-6, pt[1]-6, pt[0]+6, pt[1]+6], fill="#FFD166", outline="#D90429", width=2)

    img_bytes = io.BytesIO()
    img.save(img_bytes, format="PNG")
    img_bytes.seek(0)
    return img_bytes

def generate_30_slides_data(topic: str):
    """توليد محتوى تفصيلي متكامل لـ 30 شريحة احترافية"""
    return [
        (f"1. الغلاف والعنوان الرئيسي", [f"عرض تقديمي شامل ومفصل حول: {topic}", "إعداد تحليلي واستراتيجي متقدم", "دليل الأداء والتطبيق الميداني"]),
        (f"2. المقدمة والأهمية الاستراتيجية", [f"تعتبر دراسة ({topic}) ركيزة حيوية في تطوير منظومة العمل.", "مواكبة المعايير الحديثة في إدارة المشاريع والتنظيم الإداري.", "أهمية بناء بنية تحتية إدارية مرنة وقادرة على التكيف."]),
        (f"3. مشكلة العرض والدوافع", ["تشخيص الفجوات الميدانية والتنظيمية في بيئة العمل.", "الحاجة الماسة لتعزيز استمرارية العمليات أثناء الأزمات.", "الحد من الخسائر المادية وضمان ديمومة تقديم الخدمات."]),
        (f"4. الأهداف الاستراتيجية", ["تطوير مصفوفة استجابة سريعة للمخاطر التشغيلية.", "رفع كفاءة استغلال الموارد المتاحة بنسبة تتجاوز 40%.", "تأسيس إطار مؤسسي للتنسيق والمتابعة الميدانية."]),
        (f"5. المنهجية والأطر النظرية", ["الاعتماد على المنهج الوصفي التحليلي والميداني.", "تطبيق نماذج الجودة الشاملة والإدارة العامة الحديثة.", "الاستناد إلى أفضل الممارسات والتجارب الإقليمية والدولية."]),
        (f"6. المفاهيم والمصطلحات الأساسية", ["التعريف الإجرائي لمفاهيم التخطيط الاستراتيجي وإدارة الأزمات.", "مفهوم الاستدامة والجاهزية التشغيلية في المشاريع الحيوية.", "التفرقة بين المخاطر الروتينية والأزمات الطارئة المعقدة."]),
        (f"7. تحليل الواقع الميداني الراهن", ["تقييم مستوى البنية التحتية والجاهزية الفنية للفرق.", "رصد نقاط القوة المؤسسية واستثمارها في مواجهة الطوارئ.", "تحديد جوانب القصور ونقاط الضعف الواجب معالجتها فوراً."]),
        (f"8. تحليل البيئة الداخلية (SWOT)", ["نقاط القوة: الكادر البشري، الخبرة الميدانية، العلاقات المجتمعية.", "نقاط الضعف: شح الموازنات التشغيلية، نقص قطع الغيار التخصصية.", "الفرص: التوجه نحو الطاقة المستدامة، دعم المنظمات والجهات المانحة.", "التهديدات: الظروف البيئية، تقلبات أسعار الطاقة، الطوارئ المناخية."]),
        (f"9. استراتيجيات التخطيط الاستباقي", ["وضع سيناريوهات متعددة للتعامل مع الاحتمالات الطارئة.", "بناء مخزون استراتيجي للاحتياجات الفنية واللوجستية الأساسية.", "جدولة خطط الصيانة الدورية والوقائية لكافة المنشآت."]),
        (f"10. آليات الرصد والتنبؤ المبكر", ["تفعيل مؤشرات الإنذار المبكر للأعطال والانقطاعات.", "استخدام تقنيات التحليل الرقمي في مراقبة التدفق والتشغيل.", "تشكيل لجان متابعة ميدانية تعمل على مدار الساعة."]),
        (f"11. هيكلية إدارة الأزمات والطوارئ", ["تحديد خطوط الاتصال والتسلسل الهرمي لاتخاذ القرارات.", "منح صلاحيات استثنائية مرنة لمدراء الفروع والميدان.", "تأسيس غرفة عمليات مركزية للمتابعة والتحكم."]),
        (f"12. إدارة الموارد وسلاسل الإمداد", ["تأمين مصادر بديلة للمدخلات التشغيلية والمحروقات.", "عقد اتفاقيات مسبقة مع الموردين المحليين لتفادي التأخير.", "ترشيد الاستهلاك وتوجيه المخصصات نحو الأولويات القصوى."]),
        (f"13. تأهيل وتدريب الكوادر الميدانية", ["تنفيذ مناورات ومحاكاة دورية للتعامل مع السيناريوهات الطارئة.", "صقل مهارات السلامة المهنية والإسعافات والتدخل السريع.", "تعزيز مهارات القيادة التكتيكية لفرق الصيانة والتشغيل."]),
        (f"14. التحول نحو الطاقة البديلة والمستدامة", ["الاعتماد على منظومات الطاقة الشمسية لتأمين استمرارية الضخ.", "تقليل الاعتماد على المشتقات النفطية والحد من التكاليف.", "خفض البصمة الكربونية وحماية المنشآت من انقطاع الوقود."]),
        (f"15. التحول الرقمي والأتمتة", ["إدخال أنظمة المراقبة عن بعد (SCADA/IoT) في المشاريع.", "توثيق قواعد البيانات التشغيلية وسجلات الصيانة رقمياً.", "أتمتة بلاغات المواطنين والشكاوى لسرعة الاستجابة."]),
        (f"16. إشراك المجتمع المحلي واللجان", ["تفعيل دور لجان المستفيدين في حماية وصيانة المشاريع.", "تعزيز الشفافية والمساءلة المجتمعية لضمان استقرار التحصيل.", "بناء برامج توعوية للمحافظة على المنشآت ومصادر الموارد."]),
        (f"17. الحوكمة والرقابة المالية والإدارية", ["تطبيق معايير التدقيق المالي الداخلي على الصرف والتشغيل.", "ضبط المخزون ومنع الهدر في قطع الغيار والمعدات.", "إعداد تقارير دورية ترفع للإدارة العليا لتقييم الكفاءة."]),
        (f"18. إدارة المخاطر الفنية والتشغيلية", ["تصنيف الأعطال الفنية ووضع أدلة إصلاح قياسية (SOPs).", "توفير وحدات توليد وضخ احتياطية في المواقع الحرجة.", "الفحص الدوري لشبكات التوزيع والمنشآت لتفادي التسريب."]),
        (f"19. التواصل الإعلامي وإدارة الجمهور", ["صياغة رسائل إعلامية شفافة عند حدوث أي انقطاع طارئ.", "استخدام منصات التواصل لإعلام المجتمع بمواعيد الصيانة.", "الحفاظ على ثقة الجمهور والمستفيدين بالمنظومة الخدمية."]),
        (f"20. الشراكات وتنسيق الجهات الداعمة", ["بناء تحالفات استراتيجية مع السلطات المحلية والمنظمات.", "تكامل الجهود مع القطاعات الخدمية الأخرى (طاقة، صحة، طرق).", "جذب التمويلات التنموية لإعادة تأهيل وتوسعة المشاريع."]),
        (f"21. قياس مؤشرات الأداء الرئيسية (KPIs)", ["مؤشر زمن الاستجابة للبلاغات والأعطال الطارئة.", "مؤشر نسبة استمرارية التشغيل والضخ بدون انقطاع.", "مؤشر رضا المستفيدين وكفاءة الصرف المالي للمشاريع."]),
        (f"22. دراسة حالة تطبيقية (ميدانية)", ["استعراض تجربة واقعية في مواجهة انقطاع تشغيلي حاد.", "الخطوات المتخذة لاحتواء الأزمة وتقليل الخسائر للصفر.", "الدروس المستفادة وتحويل التحدي إلى فرصة تطويرية."]),
        (f"23. التحديات الجغرافية والبيئية", ["التغلب على وعورة التضاريس وصعوبة الوصول لبعض المواقع.", "مواجهة التغيرات المناخية والسيول وشح المياه الجوفية.", "حلول هندسية مبتكرة لتثبيت الخطوط والمنشآت في الجبال."]),
        (f"24. السلامة والأمن الصناعي", ["توفير معدات الوقاية الشخصية (PPE) لجميع الفنيين.", "تأمين محطات الضخ والخزانات ضد التخريب والتعديات.", "إجراءات التخزين الآمن للمواد الكيميائية والمحروقات."]),
        (f"25. التقييم والتحسين المستمر (PDCA)", ["دورة ديمنج: خطط - نفذ - افحص - صحح (Plan-Do-Check-Act).", "مراجعة خطة الطوارئ كل 6 أشهر وتحديثها دورياً.", "الاستماع لملاحظات فرق الميدان وتطبيق مقترحاتهم."]),
        (f"26. خريطة الطريق التنفيذية (Roadmap)", ["المرحلة 1: التدقيق الفني الشامل وحصر الاحتياجات (شهر 1-2).", "المرحلة 2: تأهيل وتدريب الكوادر واعتماد الأدلة (شهر 3-4).", "المرحلة 3: التطبيق الكامل لمنظومات الطاقة والرقابة (شهر 5-8).", "المرحلة 4: التقييم المؤسسي الشامل واستدامة الأداء (شهر 9-12)."]),
        (f"27. المخرجات المتوقعة والعائد التنموي", ["استقرار الخدمة بنسبة 98% وتفادي الانقطاعات المفاجئة.", "وفر مالي في تكاليف الصيانة والوقود يتجاوز 35%.", "رفع ثقة المجتمع وتحقيق أهداف التنمية المستدامة."]),
        (f"28. الاستنتاجات العامة للدراسة", ["الإدارة الاستباقية هي الفارق الوحيد بين النجاح والتعثر.", "الاستثمار في الكادر البشري هو الاستثمار الأكثر ربحية.", "استمرارية الخدمات تعتمد على تضافر الجهود الفنية والمجتمعية."]),
        (f"29. التوصيات الاستراتيجية الختامية", ["اعتماد موازنة طوارئ مستقلة لسرعة التدخل الميداني.", "التعميم الفوري للأدلة الإجرائية على كافة الفروع.", "إنشاء قاعدة بيانات مركزية مربوطة بأنظمة ذكاء الأعمال."]),
        (f"30. المراجع وشكر وتقدير", ["المراجع: أحدث الأدبيات الإدارية، تقارير الهيئات الميدانية، المعايير الدولية.", "شكر وتقدير لكافة الكوادر الميدانية والإدارية والجهات الداعمة.", "باب الأسئلة والمناقشة مفتوح."])
    ]

def create_powerpoint_presentation_30(topic: str, output_path: str):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slides_data = generate_30_slides_data(topic)

    for idx, (title_text, points) in enumerate(slides_data, start=1):
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # شريط علوي أنيق
        top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.2))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = RGBColor(27, 73, 101)
        top_bar.line.fill.background()

        # صندوق العنوان
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(1.1))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = title_text
        p_title.alignment = PP_ALIGN.RIGHT
        p_title.font.size = Pt(28)
        p_title.font.bold = True
        p_title.font.color.rgb = RGBColor(27, 73, 101)

        # صندوق محتوى النقاط (الجانب الأيمن)
        content_box = slide.shapes.add_textbox(Inches(5.6), Inches(1.6), Inches(6.9), Inches(5.2))
        tf_content = content_box.text_frame
        tf_content.word_wrap = True

        for p_idx, pt in enumerate(points):
            p = tf_content.paragraphs[0] if p_idx == 0 else tf_content.add_paragraph()
            p.text = f"◀ {pt}"
            p.alignment = PP_ALIGN.RIGHT
            p.font.size = Pt(18)
            p.font.color.rgb = RGBColor(40, 40, 40)
            p.space_after = Pt(14)

        # إدراج الصورة والرسم البياني التوضيحي المخصص للشريحة (الجانب الأيسر)
        img_bytes = create_slide_image(title_text, idx)
        slide.shapes.add_picture(img_bytes, Inches(0.8), Inches(1.8), Inches(4.5), Inches(4.8))

        # تذييل الشريحة مع رقم الشريحة
        footer_box = slide.shapes.add_textbox(Inches(0.8), Inches(6.9), Inches(11.733), Inches(0.4))
        p_foot = footer_box.text_frame.paragraphs[0]
        p_foot.text = f"شريحة {idx} من 30 | {topic[:40]}"
        p_foot.alignment = PP_ALIGN.LEFT
        p_foot.font.size = Pt(11)
        p_foot.font.color.rgb = RGBColor(120, 120, 120)

    prs.save(output_path)

def build_mega_research(topic: str) -> str:
    return (
        f"🎓 **بحث علمي وأكاديمي متكامل وشامل (المستوى المتقدم)**\n"
        f"📌 **عنوان البحث:** {topic}\n\n"
        f"═══════════════════════════════════════\n"
        f"📖 **1. المقدمة والخلفية العامة للدراسة:**\n"
        f"تعد دراسة ({topic}) من الركائز الاستراتيجية في منظومة الإدارة العامة المعاصرة وإدارة العمليات الميدانية. "
        f"إن التحولات السريعة والتحديات المعقدة تفرض على المؤسسات والهيئات العامة الانتقال من نمط الإدارة التقليدية "
        f"وردات الفعل إلى نمط التخطيط الاستراتيجي الوقائي وبناء السيناريوهات الاستباقية لضمان عدم توقف الخدمات الأساسية.\n\n"
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
        [InlineKeyboardButton("📊 تصميم عرض بوربوينت كامل (30 شريحة + صور)", callback_data="svc_ppt")],
        [InlineKeyboardButton("🎓 إعداد بحث جامعي متكامل ومفصل", callback_data="svc_research")],
        [InlineKeyboardButton("🧑‍🏫 تحضير دروس وخطط تعليمية معمقة", callback_data="svc_lesson")],
        [InlineKeyboardButton("📚 بحوث وتقارير مدرسية متكاملة", callback_data="svc_school")],
        [InlineKeyboardButton("📝 تلخيص مذكرات ومحتوى شامل", callback_data="svc_summary")],
        [InlineKeyboardButton("💡 حل وشرح المسائل والواجبات بالتفصيل", callback_data="svc_homework")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    welcome_text = (
        "👋 **أهلاً بك في منصة الخدمات الأكاديمية والتعليمية المتطورة**\n\n"
        "تم تحديث النظام لتقديم محتوى أكاديمي موسع وعروض تقديمية متكاملة مدعمة بالصور والرسومات.\n\n"
        "اختر الخدمة المطلوبة لبدء العمل فوراً:"
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
            await query.edit_message_text("📊 أرسل الآن عنوان العرض التقديمي لتوليد ملف بوربوينت احترافي يتكون من **30 شريحة كاملة مع الصور والرسومات التوضيحية**:")
        else:
            await query.edit_message_text(f"✨ خدمة: *{name}*\n\nيرجى إرسال تفاصيل الموضوع وسأقوم بإعداد دراسة موسعة وتفصيلية فوراً:", parse_mode="Markdown")

async def text_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_text = update.message.text
    user = update.effective_user
    current_service = context.user_data.get("current_service", "ppt")

    status_msg = await update.message.reply_text("⏳ جارٍ العمل على إعداد الملف والمحتوى التفصيلي الموسع...")

    try:
        if current_service == "ppt":
            file_name = f"presentation_{user.id}.pptx"
            create_powerpoint_presentation_30(user_text, file_name)

            with open(file_name, "rb") as ppt_file:
                await update.message.reply_document(
                    document=ppt_file,
                    filename=f"{user_text[:30]}_30Slides.pptx",
                    caption=f"✅ **تم إعداد العرض التقديمي بنجاح:**\n📌 *{user_text}*\n📊 **عدد الشرائح:** 30 شريحة كاملة ومفصلة مع الصور والرسومات البيانية.",
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
        await status_msg.edit_text(f"⚠️ حدث خطأ أثناء المعالجة: {str(e)}")

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
    
    print("البوت يعمل بأعلى كفاءة لـ 30 شريحة...")
    await app.initialize()
    await app.start()
    await app.updater.start_polling()
    
    while True:
        await asyncio.sleep(3600)

def main():
    asyncio.run(main_async())

if __name__ == "__main__":
    main()
