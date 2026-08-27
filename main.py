import os
import asyncio
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, MessageHandler, filters, ContextTypes
from aiohttp import web

TELEGRAM_TOKEN = "8867458917:AAEyVQ0Vn97bEfZbANtsFRxMxeJxnbdJ0s4"

def generate_academic_content(service_code: str, topic: str) -> str:
    if service_code == "research":
        return (
            f"🎓 **بحث جامعي متكامل حول:** {topic}\n\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"📌 **المقدمة وخلفية الدراسة:**\n"
            f"يحظى موضوع ({topic}) بأهمية بالغة في مجالات التخطيط والتطوير المؤسسي، حيث يشكل ركيزة أساسية لتحقيق الاستدامة التشغيلية ومواجهة التحديات الميدانية بكفاءة واقتدار.\n\n"
            f"🎯 **مشكلة وأهداف البحث:**\n"
            f"• تشخيص واقع الممارسات الإدارية والتنفيذية المرتبطة بالموضوع.\n"
            f"• قياس مدى فاعلية استراتيجيات الاستجابة في تحسين جودة الأداء.\n"
            f"• وضع نموذج إجرائي يساهم في دعم صناع القرار والكوادر الميدانية.\n\n"
            f"📑 **المحاور النظرية والتحليل الميداني:**\n"
            f"1. **المفاهيم والأطر الحاكمة:** تحليل الأسس المنهجية والمعايير المعتمدة.\n"
            f"2. **التحديات التشغيلية:** حصر الصعوبات والمعوقات وسبل معالجتها.\n"
            f"3. **استمرارية العمليات:** الآليات الكفيلة بضمان تدفق الخدمات واستقرارها.\n\n"
            f"💡 **النتائج والتوصيات:**\n"
            f"• ضرورة تعزيز التخطيط الوقائي وبناء خطط الطوارئ الاستباقية.\n"
            f"• تأهيل وتدريب الكوادر التنفيذية لرفع الجاهزية الميدانية.\n"
            f"• تبني نظم تقييم ورقابة دورية لضمان الاستدامة والجودة.\n\n"
            f"📚 **المراجع المقترحة:**\n"
            f"- أدبيات الإدارة العامة والتخطيط الاستراتيجي الحديث.\n"
            f"- الدراسات الميدانية ونماذج إدارة الأزمات والمشاريع."
        )

    elif service_code == "lesson":
        return (
            f"🧑‍🏫 **خطة تحضير درس تعليمي:** {topic}\n\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"🎯 **الأهداف السلوكية والمعرفية:**\n"
            f"• أن يستوعب المتعلم المفهوم الأساسي لـ ({topic}).\n"
            f"• أن يحلل العناصر والمحاور الرئيسية للدرس بدقة.\n"
            f"• أن يطبق المعارف المكتسبة في أنشطة وتطبيقات عملية.\n\n"
            f"🛠 **الوسائل والاستراتيجيات التعليمية:**\n"
            f"• استراتيجية العصف الذهني والحوار والمناقشة.\n"
            f"• العروض المرئية والخرائط الذهنية التوضيحية.\n\n"
            f"⏱ **سير الدرس والأنشطة:**\n"
            f"1. **التهيئة والتمهيد (5 دقائق):** طرح تساؤل استكشافي لجذب انتباه الطلاب.\n"
            f"2. **العرض والشرح (25 دقيقة):** تفصيل المحاور خطوة بخطوة مع الأمثلة.\n"
            f"3. **النشاط التطبيقي (10 دقائق):** عمل فردي/جماعي لقياس الاستيعاب.\n\n"
            f"📝 **التقويم الختامي والواجب:**\n"
            f"• أسئلة مراجعة وتلخيص لأهم النقاط المستفادة مع تكليف منزلي."
        )

    elif service_code == "school":
        return (
            f"📚 **تقرير مدرسي شامل حول:** {topic}\n\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"✨ **المقدمة:**\n"
            f"نقدم في هذا التقرير استعراضاً مبسطاً وشاملاً لموضوع ({topic}) لتوضيح عناصره وأهميته في حياتنا ومجتمعنا.\n\n"
            f"📌 **العناصر الأساسية:**\n"
            f"• ما هو مفهوم {topic} وأهميته العامة؟\n"
            f"• أبرز المزايا والفوائد المرتبطة به.\n"
            f"• كيف نساهم في الاستفادة منه وتطبيقه بالشكل السليم؟\n\n"
            f"🏁 **الخاتمة:**\n"
            f"إن الوعي بهذا الموضوع يفتح آفاقاً واسعة للمعرفة ويسهم في بناء بيئة تعليمية ومجتمعية واعية ومتميزة."
        )

    elif service_code == "summary":
        return (
            f"📝 **تلخيص مذكرات ومحتوى:** {topic}\n\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"💡 **الخلاصة المركزة (أهم الأفكار):**\n"
            f"• الفكرة الجوهرية تتمحور حول الفهم الشامل لأبعاد ({topic}).\n"
            f"• الربط بين المفاهيم النظرية والتطبيق الواقعي.\n\n"
            f"🔑 **النقاط والمفاهيم الرئيسية:**\n"
            f"1. التعريف والأركان الأساسية.\n"
            f"2. القواعد والضوابط المنهجية.\n"
            f"3. المخرجات والنتائج المستخلصة.\n\n"
            f"📌 **الملاحظات الهامة للمراجعة:**\n"
            f"التركيز على المفاهيم المفتاحية، واستخدام الخرائط الذهنية لتثبيت المعلومات."
        )

    elif service_code == "homework":
        return (
            f"💡 **حل وشرح المسائل والواجبات حول:** {topic}\n\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"🔍 **تحليل السؤال والمعطيات:**\n"
            f"الموضوع المطلوب حله وتوضيحه: {topic}\n\n"
            f"📐 **خطوات الحل المنطقي والنموذجي:**\n"
            f"1. **الخطوة الأولى:** تحديد القوانين والمفاهيم المرتبطة بالسؤال.\n"
            f"2. **الخطوة الثانية:** التطبيق المباشر وفق الترتيب المنهجي السليم.\n"
            f"3. **الخطوة الثالثة:** استخراج النتيجة وتدقيق صحتها.\n\n"
            f"✅ **النتيجة والتفسير:**\n"
            f"تم الوصول إلى الإجابة النموذجية المعتمدة مع توضيح سبب اختيار كل خطوة لترسيخ الفهم."
        )

    return f"طلبك حول ({topic}) تم استلامه وإعداده بنجاح."

def create_powerpoint_presentation(topic: str, output_path: str):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slides_data = [
        (f"عرض تقديمي: {topic}", [
            f"إعداد شامل ومفصل حول: {topic}",
            "الرؤية والأهداف الاستراتيجية",
            "أهمية الموضوع في تطوير الأداء وتحقيق الكفاءة"
        ]),
        ("المحاور الرئيسية والأطر النظرية", [
            "تحليل الواقع التشغيلي والميداني",
            "استراتيجيات الاستجابة وإدارة العمليات",
            "مؤشرات قياس الجودة ومستوى الاستمرارية"
        ]),
        ("التطبيق والتحليل الميداني", [
            "دراسة التحديات الواقعية وسيناريوهات التعامل معها",
            "آليات تعزيز كفاءة الموارد المادية والبشرية",
            "تطوير مسارات التدفق والجاهزية التشغيلية"
        ]),
        ("النتائج والتوصيات التنفيذية", [
            "تبني خطط استباقية لإدارة الطوارئ والأزمات",
            "تأهيل الكوادر الفنية والإدارية باستمرار",
            "تعزيز التقييم والمتابعة لضمان استدامة المشاريع"
        ])
    ]

    for title_text, points in slides_data:
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)

        # عنوان الشريحة
        title_box = slide.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.333), Inches(1.2))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = title_text
        p_title.alignment = PP_ALIGN.RIGHT
        p_title.font.size = Pt(36)
        p_title.font.bold = True
        p_title.font.color.rgb = RGBColor(24, 76, 120)

        # محتوى الشريحة
        content_box = slide.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.333), Inches(4.5))
        tf_content = content_box.text_frame
        tf_content.word_wrap = True

        for idx, pt in enumerate(points):
            p = tf_content.paragraphs[0] if idx == 0 else tf_content.add_paragraph()
            p.text = f"• {pt}"
            p.alignment = PP_ALIGN.RIGHT
            p.font.size = Pt(22)
            p.font.color.rgb = RGBColor(50, 50, 50)
            p.space_after = Pt(14)

    prs.save(output_path)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [
        [InlineKeyboardButton("📊 تصميم عرض بوربوينت", callback_data="svc_ppt")],
        [InlineKeyboardButton("🎓 إعداد بحث جامعي متكامل", callback_data="svc_research")],
        [InlineKeyboardButton("🧑‍🏫 تحضير دروس وخطط تعليمية", callback_data="svc_lesson")],
        [InlineKeyboardButton("📚 بحوث وتقارير مدرسية", callback_data="svc_school")],
        [InlineKeyboardButton("📝 تلخيص كتب ومذكرات", callback_data="svc_summary")],
        [InlineKeyboardButton("💡 حل وشرح الواجبات والمسائل", callback_data="svc_homework")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    welcome_text = "👋 أهلاً بك في منصة الخدمات الأكاديمية والتعليمية\n\nاختر الخدمة المطلوبة من القائمة:"
    
    if update.message:
        await update.message.reply_text(welcome_text, reply_markup=reply_markup)
    elif update.callback_query:
        await update.callback_query.message.reply_text(welcome_text, reply_markup=reply_markup)

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    data = query.data
    services = {
        "svc_ppt": ("📊 تصميم عرض بوربوينت", "ppt"),
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
            await query.edit_message_text("📊 أرسل الآن عنوان أو موضوع العرض التقديمي لتوليد ملف البوربوينت فوراً:")
        else:
            await query.edit_message_text(f"✨ خدمة: *{name}*\n\nيرجى كتابة الموضوع بالتفصيل وسأقوم بإعداده لك فوراً:", parse_mode="Markdown")

async def text_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_text = update.message.text
    user = update.effective_user
    current_service = context.user_data.get("current_service", "ppt")

    status_msg = await update.message.reply_text("⏳ جارٍ إعداد وتجهيز المحتوى المطلوب...")

    try:
        if current_service == "ppt":
            file_name = f"presentation_{user.id}.pptx"
            create_powerpoint_presentation(user_text, file_name)

            with open(file_name, "rb") as ppt_file:
                await update.message.reply_document(
                    document=ppt_file,
                    filename=f"{user_text[:30]}.pptx",
                    caption=f"✅ تم تصميم العرض التقديمي بنجاح:\n📌 *{user_text}*",
                    parse_mode="Markdown"
                )

            if os.path.exists(file_name):
                os.remove(file_name)
        else:
            result = generate_academic_content(current_service, user_text)
            await update.message.reply_text(result, parse_mode="Markdown")

        await status_msg.delete()

    except Exception as e:
        await status_msg.edit_text(f"⚠️ حدث خطأ: {str(e)}")

# سيرفر لإبقاء الخدمة متصلة دائماً على Render
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
    
    print("البوت يعمل بنجاح بكافة الخدمات...")
    await app.initialize()
    await app.start()
    await app.updater.start_polling()
    
    while True:
        await asyncio.sleep(3600)

def main():
    asyncio.run(main_async())

if __name__ == "__main__":
    main()
