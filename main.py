import json
import logging
import os
from pptx import Presentation
from pptx.util import Pt
import requests
from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import (
    ApplicationBuilder,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

# ----------------- الإعدادات -----------------
TELEGRAM_BOT_TOKEN = "8867458917:AAEyVQ0Vn97bEfZbANtsFRxMxeJxnbdJ0s4"
GEMINI_API_KEY = "AQ.Ab8RN6LOZVx1Re_-xU0Uo_cLpfg1pDcg1muqRIEERFEZX4p8WQ"
ADMIN_ID = 578187098

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)

USER_MODES = {}


# ----------------- استدعاء Gemini -----------------
def call_gemini(
    prompt: str, model_name: str = "gemini-1.5-flash", json_mode: bool = False
) -> str:
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={GEMINI_API_KEY}"
    headers = {"Content-Type": "application/json"}
    payload = {"contents": [{"parts": [{"text": prompt}]}]}

    if json_mode:
        payload["generationConfig"] = {
            "response_mime_type": "application/json"
        }

    response = requests.post(url, headers=headers, json=payload, timeout=90)
    response_data = response.json()

    if "error" in response_data:
        raise Exception(response_data["error"].get("message", "خطأ في API"))

    return response_data["candidates"][0]["content"]["parts"][0]["text"]


# ----------------- توليد ملف البوربوينت -----------------
def generate_ppt_file(topic: str, filename: str = "presentation.pptx") -> str:
    prompt = f"""
    قم بإنشاء محتوى عرض تقديمي تفصيلي واحترافي باللغة العربية عن: "{topic}".
    يجب أن يكون الإخراج بصيغة JSON فقط كقائمة شرائح:
    [
      {{"title": "عنوان الشريحة", "points": ["نقطة 1", "نقطة 2", "نقطة 3"]}}
    ]
    """
    raw_text = call_gemini(
        prompt, model_name="gemini-1.5-flash", json_mode=True
    )
    slides_data = json.loads(raw_text)

    prs = Presentation()
    for slide_data in slides_data:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = slide_data.get("title", "")
        tf = slide.shapes.placeholders[1].text_frame
        tf.clear()
        for i, point in enumerate(slide_data.get("points", [])):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = str(point)
            p.font.size = Pt(16)

    prs.save(filename)
    return filename


# ----------------- القوائم -----------------
def get_main_menu():
    keyboard = [
        [
            InlineKeyboardButton(
                "📊 تصميم عرض بوربوينت", callback_data="mode_ppt"
            )
        ],
        [
            InlineKeyboardButton(
                "🎓 إعداد بحث جامعي متكامل", callback_data="mode_research"
            )
        ],
        [
            InlineKeyboardButton(
                "👨‍🏫 تحضير دروس وخطط تعليمية", callback_data="mode_teacher"
            )
        ],
        [
            InlineKeyboardButton(
                "📚 بحوث وتقارير مدرسية", callback_data="mode_school"
            )
        ],
        [
            InlineKeyboardButton(
                "📝 تلخيص كتب ومذكرات", callback_data="mode_summary"
            )
        ],
        [
            InlineKeyboardButton(
                "💡 حل وشرح الواجبات والمسائل", callback_data="mode_homework"
            )
        ],
    ]
    return InlineKeyboardMarkup(keyboard)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    welcome_text = (
        "👋 **أهلاً بك في منصة الخدمات الأكاديمية والتعليمية**\n\n"
        "اختر الخدمة المطلوبة من القائمة:"
    )
    await update.message.reply_text(
        welcome_text, reply_markup=get_main_menu(), parse_mode="Markdown"
    )


async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    mode = query.data
    USER_MODES[user_id] = mode

    prompts = {
        "mode_ppt": "📊 أرسل الآن عنوان أو موضوع العرض التقديمي لتوليد الملف فوراً:",
        "mode_research": "🎓 أرسل موضوع البحث الجامعي بالتفصيل لإعداده منهجياً:",
        "mode_teacher": "👨‍🏫 أرسل عنوان الدرس والمرحلة الدراسية لتجهيز الخطة:",
        "mode_school": "📚 أرسل عنوان التقرير أو البحث المدرسي المطلوب:",
        "mode_summary": "📝 أرسل النص أو الموضوع المراد تلخيصه بنقاط مركزة:",
        "mode_homework": "💡 أرسل المسألة أو الواجب المطلوب حله خطوة بخطوة:",
    }
    await query.edit_message_text(prompts.get(mode, "أرسل طلبك الآن:"))


# ----------------- معالجة الرسائل -----------------
async def handle_requests(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_name = update.effective_user.full_name
    user_text = update.message.text
    mode = USER_MODES.get(user_id, "mode_ppt")

    try:
        if user_id != ADMIN_ID:
            await context.bot.send_message(
                chat_id=ADMIN_ID,
                text=f"🔔 **طلب جديد**\n👤 المستخدم: {user_name} (`{user_id}`)\n📌 الخدمة: `{mode}`\n📝 الطلب:\n{user_text}",
                parse_mode="Markdown",
            )
    except Exception:
        pass

    status_msg = await update.message.reply_text(
        "⏳ جاري معالجة طلبك بواسطة الذكاء الاصطناعي..."
    )

    try:
        if mode == "mode_ppt":
            ppt_file = generate_ppt_file(user_text, f"ppt_{user_id}.pptx")
            await status_msg.edit_text("✅ تم التجهيز! جاري إرسال الملف...")
            with open(ppt_file, "rb") as doc:
                await context.bot.send_document(
                    chat_id=user_id,
                    document=doc,
                    caption=f"📊 عرض بوربوينت: {user_text}",
                    filename="عرض_تقديمي.pptx",
                )
        else:
            prompts_map = {
                "mode_research": f"أنت باحث أكاديمي متخصص. قم بإعداد بحث جامعي رصين وشامل عن: '{user_text}'. قسّم البحث إلى: مقدمة، مباحث وفروع، خاتمة وتوصيات، ومراجع.",
                "mode_teacher": f"أنت خبير تربوي. قم بإعداد تحضير درس نموذجي عن: '{user_text}' يشمل الأهداف واستراتيجيات التدريس وخطة الحصة والتقويم.",
                "mode_school": f"قم بإعداد بحث مدرسي منظم ومبسط عن: '{user_text}'.",
                "mode_summary": f"قم بتلخيص المحتوى التالي في نقاط واضحة: '{user_text}'",
                "mode_homework": f"قم بحل المسألة أو الواجب التالي خطوة بخطوة مع الشرح: '{user_text}'",
            }
            prompt = prompts_map.get(mode, user_text)
            result = call_gemini(prompt, model_name="gemini-1.5-flash")

            if len(result) > 4000:
                for chunk in [
                    result[i : i + 4000] for i in range(0, len(result), 4000)
                ]:
                    await update.message.reply_text(chunk)
            else:
                await update.message.reply_text(result)

            await status_msg.delete()

        await update.message.reply_text(
            "اختر خدمة أخرى للمتابعة:", reply_markup=get_main_menu()
        )

    except Exception as e:
        await status_msg.edit_text(f"⚠️ حدث خطأ: {e}")


def main():
    app = (
        ApplicationBuilder()
        .token(TELEGRAM_BOT_TOKEN)
        .connect_timeout(60.0)
        .read_timeout(60.0)
        .write_timeout(60.0)
        .get_updates_connect_timeout(60.0)
        .get_updates_read_timeout(60.0)
        .build()
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(
        MessageHandler(filters.TEXT & ~filters.COMMAND, handle_requests)
    )

    print("البوت يعمل بنجاح...")
    app.run_polling()


if __name__ == "__main__":
    main()
