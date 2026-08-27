import os
import asyncio
import requests
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, MessageHandler, filters, ContextTypes
from aiohttp import web

TELEGRAM_TOKEN = "8867458917:AAEyVQ0Vn97bEfZbANtsFRxMxeJxnbdJ0s4"

def generate_ai_content(prompt: str) -> str:
    # استخدام سيرفر Blackbox AI المجاني (لا يحتاج مفاتيح ولا يحظر السيرفرات)
    url = "https://api.blackbox.ai/api/chat"
    headers = {"Content-Type": "application/json"}
    payload = {
        "messages": [
            {"role": "user", "content": prompt}
        ],
        "model": "deepseek-coder-v2"
    }
    
    try:
        response = requests.post(url, headers=headers, json=payload, timeout=60)
        if response.status_code == 200:
            text = response.text
            # تنظيف أي نصوص إضافية يرسلها السيرفر أحياناً
            if "$$$" in text:
                text = text.split("$$$")[-1]
            return text.strip()
        else:
            raise Exception(f"API Error {response.status_code}")
    except Exception:
        # رابط احتياطي في حال الضغط على السيرفر الأول
        try:
            fallback_url = "https://backend.buildpicoapps.com/aero/run/llm-api?pk=v1-Z0FBQUFBQm1fM1M2S2s1N2xVb19yX18tOFRHcjNpd1ZSNi1zT0ZtMmlVd0F6RnRKZnNqdWRtSGxjTklNTE5Mbnc5cEhrbC1hQk4xX3dJRG1oanFZa0ZJLVpEWE51U3o1Umc9PQ=="
            res = requests.post(fallback_url, json={"prompt": prompt}, timeout=60)
            if res.status_code == 200:
                return res.json().get("text", "عذراً، لم أتمكن من توليد النص.")
            raise Exception()
        except Exception:
            raise Exception("سيرفرات الذكاء الاصطناعي المجانية تواجه ضغطاً حالياً، يرجى المحاولة بعد دقيقة.")

def create_powerpoint_presentation(topic: str, output_path: str):
    prompt = (
        f"أنشئ محتوى عرض تقديمي احترافي باللغة العربية حول: '{topic}'.\n"
        f"يجب أن يتكون العرض من 4 شرائح على النحو التالي:\n"
        f"شريحة 1: المقدمة والعنوان\n"
        f"شريحة 2: المحاور الرئيسية والأهمية\n"
        f"شريحة 3: التفاصيل والتطبيق العملي\n"
        f"شريحة 4: التوصيات والخاتمة\n\n"
        f"التنسيق المطلوب إلزامي ودقيق جداً:\n"
        f"---SLIDE---\n"
        f"TITLE: [عنوان الشريحة]\n"
        f"CONTENT:\n"
        f"- [نقطة 1]\n"
        f"- [نقطة 2]\n"
        f"- [نقطة 3]"
    )

    ai_text = generate_ai_content(prompt)

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slides_raw = ai_text.split("---SLIDE---")
    
    for raw in slides_raw:
        if not raw.strip():
            continue
            
        title_text = ""
        points = []
        
        lines = raw.strip().split("\n")
        content_started = False
        
        for line in lines:
            line_str = line.strip()
            if line_str.startswith("TITLE:"):
                title_text = line_str.replace("TITLE:", "").strip()
            elif "CONTENT:" in line_str:
                content_started = True
            elif content_started and line_str:
                clean_point = line_str.lstrip("-*• 1234567890.").strip()
                if clean_point:
                    points.append(clean_point)
                    
        blank_slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_slide_layout)

        # عنوان الشريحة
        title_box = slide.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.333), Inches(1.2))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = title_text if title_text else topic
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
            await query.edit_message_text("📊 أرسل الآن عنوان أو موضوع العرض التقديمي لتوليد الملف فوراً:")
        else:
            await query.edit_message_text(f"✨ خدمة: *{name}*\n\nيرجى كتابة الموضوع بالتفصيل وسأقوم بإعداده لك فوراً:", parse_mode="Markdown")

async def text_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_text = update.message.text
    user = update.effective_user
    current_service = context.user_data.get("current_service", "general")
    service_name = context.user_data.get("service_name", "طلب عام")

    status_msg = await update.message.reply_text("⏳ جارٍ العمل على طلبك ومعالجة البيانات بالذكاء الاصطناعي...")

    try:
        if current_service == "ppt":
            file_name = f"presentation_{user.id}.pptx"
            create_powerpoint_presentation(user_text, file_name)

            with open(file_name, "rb") as ppt_file:
                await update.message.reply_document(
                    document=ppt_file,
                    filename=f"{user_text[:30]}.pptx",
                    caption=f"✅ تم إعداد العرض التقديمي بنجاح:\n📌 *{user_text}*",
                    parse_mode="Markdown"
                )

            if os.path.exists(file_name):
                os.remove(file_name)
        else:
            prompt = (
                f"أنت خبير أكاديمي وباحث محترف. قم بإعداد: {service_name}.\n"
                f"الموضوع المطلوب: {user_text}\n\n"
                f"يرجى كتابة محتوى شامل ومتكامل، منظم بعناوين وفقرات واضحة، مع مقدمة ومحاور رئيسية وخاتمة باللغة العربية الفصحى."
            )
            result = generate_ai_content(prompt)
            
            if len(result) > 4000:
                for i in range(0, len(result), 4000):
                    await update.message.reply_text(result[i:i+4000])
            else:
                await update.message.reply_text(result)

        await status_msg.delete()

    except Exception as e:
        await status_msg.edit_text(f"⚠️ حدث خطأ: {str(e)}")

# سيرفر لإبقاء الخدمة تعمل على Render
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
