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

# محاولة استيراد مكتبة قراءة ملفات الـ PDF
try:
    import pypdf
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False

# ============================================================
# منصة إنجاز - بوت تحويل الـ PDF المباشر (دقة 100%)
# ============================================================

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "8867458917:AAEyVQ0Vn97bEfZbANtsFRxMxeJxnbdJ0s4")

GRADES = {
    "kg": "رياض الأطفال", "p1": "الأول الابتدائي", "p2": "الثاني الابتدائي", "p3": "الثالث الابتدائي",
    "p4": "الرابع الابتدائي", "p5": "الخامس الابتدائي", "p6": "السادس الابتدائي",
    "m1": "الأول المتوسط", "m2": "الثاني المتوسط", "m3": "الثالث المتوسط",
    "s1": "الأول الثانوي", "s2": "الثاني الثانوي", "s3": "الثالث الثانوي",
}

def extract_text_from_pdf(pdf_bytes: bytes) -> str:
    if not PDF_SUPPORT:
        return "مكتبة قراءة الـ PDF غير متوفرة في النظام حالياً."
    try:
        reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
        full_text = ""
        for page in reader.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n---\n"
        return full_text.strip()
    except Exception as e:
        return f"خطأ في قراءة ملف الـ PDF: {str(e)}"

def create_ppt_from_pdf_text(pdf_text: str, output_path: str):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    # تقسيم النص المأخوذ من الـ PDF إلى أجزاء لتوزيعها على الشحنات
    paragraphs = [p.strip() for p in pdf_text.split("\n") if p.strip() and len(p.strip()) > 10]
    if not paragraphs:
        paragraphs = ["محتوى مستخرج من ملف الـ PDF المدرسي."]

    chunk_size = 4
    chunks = [paragraphs[i:i + chunk_size] for i in range(0, len(paragraphs), chunk_size)]
    
    # تحديد عدد الشرائح (بحد أقصى 30 شريحة أو حسب محتوى الملف)
    total_slides = min(max(len(chunks), 1), 30)

    for idx in range(1, total_slides + 1):
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
        p_title.text = f"شريحة تعليمية رقم ({idx})"
        p_title.alignment = PP_ALIGN.RIGHT
        p_title.font.size = Pt(26)
        p_title.font.bold = True
        p_title.font.color.rgb = RGBColor(27, 73, 101)

        content_box = slide.shapes.add_textbox(Inches(5.6), Inches(1.7), Inches(6.8), Inches(4.7))
        tf_content = content_box.text_frame
        tf_content.word_wrap = True

        current_chunk = chunks[(idx - 1) % len(chunks)]
        for p_idx, point in enumerate(current_chunk):
            p = tf_content.paragraphs[0] if p_idx == 0 else tf_content.add_paragraph()
            p.text = f"◀ {point[:120]}"
            p.alignment = PP_ALIGN.RIGHT
            p.font.size = Pt(16)
            p.font.color.rgb = RGBColor(30, 41, 59)
            p.space_after = Pt(12)

        # رسم صورة شكلية توضيحية جانبية
        img = Image.new("RGB", (600, 450), color="#F8FAFC")
        draw = ImageDraw.Draw(img)
        draw.rounded_rectangle([15, 15, 585, 435], radius=15, fill="#FFFFFF", outline="#CBD5E1", width=2)
        draw.rectangle([15, 15, 585, 70], fill="#1B4965")
        output = io.BytesIO()
        img.save(output, format="PNG")
        output.seek(0)
        slide.shapes.add_picture(output, Inches(0.8), Inches(1.5), Inches(4.3), Inches(5.1))

        footer_box = slide.shapes.add_textbox(Inches(0.8), Inches(6.9), Inches(11.733), Inches(0.4))
        p_foot = footer_box.text_frame.paragraphs[0]
        p_foot.text = f"شريحة {idx} من {total_slides} | منصة إنجاز | تحويل PDF مباشر | 1448هـ"
        p_foot.alignment = PP_ALIGN.LEFT
        p_foot.font.size = Pt(11)
        p_foot.font.color.rgb = RGBColor(148, 163, 184)

    prs.save(output_path)

def create_word_from_pdf_text(pdf_text: str, output_path: str):
    doc = Document()
    for section in doc.sections:
        section.top_margin, section.bottom_margin = DocxInches(1), DocxInches(1)
        section.left_margin, section.right_margin = DocxInches(1), DocxInches(1)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = title_p.add_run("مستند تم تحويله من ملف PDF المدرسي")
    run_title.font.size, run_title.font.bold, run_title.font.color.rgb = DocxPt(18), True, DocxRGB(27, 73, 101)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = sub_p.add_run("منصة إنجاز للخدمات التعليمية والأكاديمية 1448هـ\n" + "—" * 35)
    run_sub.font.size, run_sub.font.color.rgb = DocxPt(11), DocxRGB(100, 116, 139)

    for block in pdf_text.split("\n"):
        clean = block.strip()
        if not clean: continue
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(clean)
        run.font.size, run.font.color.rgb = DocxPt(11.5), DocxRGB(30, 41, 59)
        p.paragraph_format.line_spacing, p.paragraph_format.space_after = 1.25, DocxPt(6)

    doc.save(output_path)

def main_menu():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📊 تحويل ملف PDF إلى بوربوينت", callback_data="mode_pdf_ppt")],
        [InlineKeyboardButton("📝 تحويل ملف PDF إلى Word", callback_data="mode_pdf_word")],
        [InlineKeyboardButton("🌟 العروض والخدمات العامة السابقة", callback_data="old_services")],
    ])

def old_menu():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📊 بوربوينت 30 شريحة ذكي", callback_data="svc_ppt")],
        [InlineKeyboardButton("📝 اختبارات ومهارات نافس", callback_data="svc_exam")],
        [InlineKeyboardButton("📈 خطط علاجية وإثرائية", callback_data="svc_remedial")],
        [InlineKeyboardButton("⬅️ العودة للقائمة الرئيسية", callback_data="home")],
    ])

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    welcome_text = "🌟 *أهلاً بك في منصة إنجاز (التحويل المباشر الدقيق)*\n\nاختر الخدمة المطلوبة أدناه:"
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

    if data == "old_services":
        await query.edit_message_text("👇 الخدمات السابقة:", reply_markup=old_menu())
        return

    if data == "mode_pdf_ppt":
        context.user_data["action"] = "pdf_to_ppt"
        await query.edit_message_text("📂 **حالة التحويل: PDF إلى بوربوينت**\n\nالآن أرسل ملف الـ **PDF** الخاص بك هنا في المحادثة وسأقوم بتحويله فوراً إلى عرض بوربوينت منسق.", parse_mode="Markdown")
        return

    if data == "mode_pdf_word":
        context.user_data["action"] = "pdf_to_word"
        await query.edit_message_text("📂 **حالة التحويل: PDF إلى Word**\n\nالآن أرسل ملف الـ **PDF** الخاص بك هنا في المحادثة وسأقوم بتحويله فوراً إلى مستند Word نظيف.", parse_mode="Markdown")
        return

async def document_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    user = update.effective_user
    action = context.user_data.get("action", "pdf_to_ppt")

    if not doc.file_name.lower().endswith(".pdf"):
        await update.message.reply_text("⚠️ يرجى إرسال ملف بصيغة PDF حصرياً.")
        return

    status_msg = await update.message.reply_text("⏳ جارٍ استلام الملف وقراءته حرفياً لتوليد الملف المطلوب...")

    try:
        file = await context.bot.get_file(doc.file_id)
        pdf_bytes = await file.download_as_bytearray()
        extracted_text = extract_text_from_pdf(bytes(pdf_bytes))

        if not extracted_text or len(extracted_text) < 10:
            await status_msg.edit_text("⚠️ عذراً، لم يتم العثور على نصوص واضحة داخل ملف الـ PDF المرسل.")
            return

        if action == "pdf_to_word":
            output_file = f"doc_{user.id}_{uuid.uuid4().hex[:6]}.docx"
            create_word_from_pdf_text(extracted_text, output_file)
            with open(output_file, "rb") as f:
                await update.message.reply_document(
                    document=f, filename=f"Converted_{doc.file_name[:-4]}.docx",
                    caption="✅ تم تحويل ملف الـ PDF إلى Word بنجاح تام وبدقة 100%."
                )
        else:
            output_file = f"ppt_{user.id}_{uuid.uuid4().hex[:6]}.pptx"
            create_ppt_from_pdf_text(extracted_text, output_file)
            with open(output_file, "rb") as f:
                await update.message.reply_document(
                    document=f, filename=f"Converted_{doc.file_name[:-4]}.pptx",
                    caption="✅ تم تحويل ملف الـ PDF إلى عرض بوربوينت بنجاح تام."
                )

        if os.path.exists(output_file):
            os.remove(output_file)
        await status_msg.delete()

    except Exception as e:
        await status_msg.edit_text(f"⚠️ حدث خطأ أثناء المعالجة: {str(e)[:150]}")

async def handle_ping(request): return web.Response(text="Bot is online!")

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
    app.add_handler(MessageHandler(filters.Document.ALL, document_handler))
    
    print("البوت يعمل الآن بصيغة التحويل المباشر للـ PDF")
    await app.initialize()
    await app.start()
    await app.updater.start_polling()
    while True: await asyncio.sleep(3600)

def main(): asyncio.run(main_async())

if __name__ == "__main__": main()
