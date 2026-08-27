import os
import io
import asyncio
import uuid
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

# استيراد مكتبة قراءة الـ PDF بشكل مؤكد
try:
    import pypdf
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "8867458917:AAEyVQ0Vn97bEfZbANtsFRxMxeJxnbdJ0s4")

def extract_clean_paragraphs(pdf_bytes: bytes):
    if not PDF_SUPPORT:
        return ["خطأ: مكتبة pypdf غير مثبتة في متطلبات النظام."]
    try:
        reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
        paragraphs = []
        for page in reader.pages:
            text = page.extract_text()
            if text:
                for line in text.split("\n"):
                    clean_line = line.strip()
                    if clean_line and len(clean_line) > 2:
                        paragraphs.append(clean_line)
        return paragraphs if paragraphs else ["لم يتم العثور على نصوص واضحة داخل ملف الـ PDF."]
    except Exception as e:
        return [f"خطأ في قراءة الملف: {str(e)}"]

def create_clean_ppt_from_pdf(paragraphs, output_path: str):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    chunk_size = 4
    chunks = [paragraphs[i:i + chunk_size] for i in range(0, len(paragraphs), chunk_size)]
    total_slides = min(len(chunks), 35)

    if total_slides == 0:
        total_slides = 1
        chunks = [["محتوى الملف المستخرج"]]

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
        p_title.text = f"محور العمل - شريحة ({idx})"
        p_title.alignment = PP_ALIGN.RIGHT
        p_title.font.size = Pt(24)
        p_title.font.bold = True
        p_title.font.color.rgb = RGBColor(27, 73, 101)

        content_box = slide.shapes.add_textbox(Inches(5.6), Inches(1.7), Inches(6.8), Inches(4.7))
        tf_content = content_box.text_frame
        tf_content.word_wrap = True

        current_chunk = chunks[idx - 1]
        for p_idx, point in enumerate(current_chunk):
            p = tf_content.paragraphs[0] if p_idx == 0 else tf_content.add_paragraph()
            p.text = f"• {point[:110]}"
            p.alignment = PP_ALIGN.RIGHT
            p.font.size = Pt(16)
            p.font.color.rgb = RGBColor(30, 41, 59)
            p.space_after = Pt(12)

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
        p_foot.text = f"شريحة {idx} من {total_slides} | منصة إنجاز 1448هـ"
        p_foot.alignment = PP_ALIGN.LEFT
        p_foot.font.size = Pt(11)
        p_foot.font.color.rgb = RGBColor(148, 163, 184)

    prs.save(output_path)

def create_word_from_pdf_paras(paragraphs, output_path: str):
    doc = Document()
    for section in doc.sections:
        section.top_margin, section.bottom_margin = DocxInches(1), DocxInches(1)
        section.left_margin, section.right_margin = DocxInches(1), DocxInches(1)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = title_p.add_run("المستند المستخرج من ملف PDF")
    run_title.font.size, run_title.font.bold, run_title.font.color.rgb = DocxPt(18), True, DocxRGB(27, 73, 101)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = sub_p.add_run("منصة إنجاز للخدمات التعليمية 1448هـ\n" + "—" * 35)
    run_sub.font.size, run_sub.font.color.rgb = DocxPt(11), DocxRGB(100, 116, 139)

    for clean in paragraphs:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(clean)
        run.font.size, run.font.color.rgb = DocxPt(11.5), DocxRGB(30, 41, 59)
        p.paragraph_format.line_spacing, p.paragraph_format.space_after = 1.25, DocxPt(6)

    doc.save(output_path)

def main_menu():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📊 تحويل ملف PDF إلى بوربوينت مرتب", callback_data="mode_pdf_ppt")],
        [InlineKeyboardButton("📝 تحويل ملف PDF إلى Word منظم", callback_data="mode_pdf_word")],
    ])

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    text = "🌟 *أهلاً بك في منصة إنجاز للتحويل الدقيق*\n\nاختر نوع التحويل المطلوب:"
    if update.message:
        await update.message.reply_text(text, reply_markup=main_menu(), parse_mode="Markdown")
    elif update.callback_query:
        await update.callback_query.message.reply_text(text, reply_markup=main_menu(), parse_mode="Markdown")

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    if data == "mode_pdf_ppt":
        context.user_data["action"] = "pdf_to_ppt"
        await query.edit_message_text("📂 أرسل ملف الـ **PDF** الآن وسأحوله إلى بوربوينت مقسم لعدة شرائح مرتبة.", parse_mode="Markdown")
    elif data == "mode_pdf_word":
        context.user_data["action"] = "pdf_to_word"
        await query.edit_message_text("📂 أرسل ملف الـ **PDF** الآن وسأحوله إلى مستند Word منظم.", parse_mode="Markdown")

async def document_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    user = update.effective_user
    action = context.user_data.get("action", "pdf_to_ppt")

    if not doc.file_name.lower().endswith(".pdf"):
        await update.message.reply_text("⚠️ يرجى إرسال ملف PDF فقط.")
        return

    status_msg = await update.message.reply_text("⏳ جارٍ قراءة وترتيب محتوى الملف لإنشاء شرائح منظمة...")

    try:
        file = await context.bot.get_file(doc.file_id)
        pdf_bytes = await file.download_as_bytearray()
        paragraphs = extract_clean_paragraphs(bytes(pdf_bytes))

        if action == "pdf_to_word":
            output_file = f"doc_{user.id}_{uuid.uuid4().hex[:6]}.docx"
            create_word_from_pdf_paras(paragraphs, output_file)
            with open(output_file, "rb") as f:
                await update.message.reply_document(
                    document=f, filename=f"Organized_{doc.file_name[:-4]}.docx",
                    caption="✅ تم تحويل الملف إلى Word منظم بدقة."
                )
        else:
            output_file = f"ppt_{user.id}_{uuid.uuid4().hex[:6]}.pptx"
            create_clean_ppt_from_pdf(paragraphs, output_file)
            with open(output_file, "rb") as f:
                await update.message.reply_document(
                    document=f, filename=f"Organized_{doc.file_name[:-4]}.pptx",
                    caption=f"✅ تم تحويل الملف إلى بوربوينت مقسم لعدة شرائح مرتبة بنجاح."
                )

        if os.path.exists(output_file):
            os.remove(output_file)
        await status_msg.delete()
    except Exception as e:
        await status_msg.edit_text(f"⚠️ حدث خطأ: {str(e)[:150]}")

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
    
    print("البوت يعمل الآن بصيغة التقسيم الاحترافي للشرائح")
    await app.initialize()
    await app.start()
    await app.updater.start_polling()
    while True: await asyncio.sleep(3600)

def main(): asyncio.run(main_async())

if __name__ == "__main__": main()
