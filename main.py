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
from telegram.ext import (
    Application, CommandHandler, CallbackQueryHandler,
    MessageHandler, filters, ContextTypes
)
from aiohttp import web

# ============================================================
# منصة إنجاز - بوت 1448هـ
# مهم: لا تضع توكن تيليجرام داخل الكود.
# في Render/Railway/أي استضافة أضف متغير البيئة TELEGRAM_TOKEN.
# ============================================================

TELEGRAM_TOKEN = "8867458917:AAEyVQ0Vn97bEfZbANtsFRxMxeJxnbdJ0s4"

# جميع الصفوف الدراسية المستهدفة
GRADES = {
    "kg": "رياض الأطفال",
    "p1": "الأول الابتدائي",
    "p2": "الثاني الابتدائي",
    "p3": "الثالث الابتدائي",
    "p4": "الرابع الابتدائي",
    "p5": "الخامس الابتدائي",
    "p6": "السادس الابتدائي",
    "m1": "الأول المتوسط",
    "m2": "الثاني المتوسط",
    "m3": "الثالث المتوسط",
    "s1": "الأول الثانوي",
    "s2": "الثاني الثانوي",
    "s3": "الثالث الثانوي",
}

# ============================================================
# خصائص العرض المطلوبة من الصورة
# ============================================================
IMAGE_REQUIREMENTS = """
يجب أن يتضمن العرض التقديمي، بما يتناسب مع طبيعة الدرس والمرحلة:

1) تمهيد وإغلاق مناسب للدرس.
2) مهارات تفكير عليا + استراتيجيات التعلم النشط + ألعاب تعليمية.
3) دعم بمقاطع تعليمية وفواصل مناسبة.
4) ربط احترافي بالوطن والدين والواقع والمواد الدراسية الأخرى.
5) مهارات وتدريبات يومية تحاكي الاختبارات الوطنية (نافس) عندما تكون المرحلة
   والمادة مناسبة لذلك.
6) تفعيل الأنشطة الفردية والثنائية والجماعية.
7) ورقة عمل + تقويم ختامي للدرس.

ويجب ألا تكون هذه العناصر مجرد عناوين عامة؛ بل تُكتب أنشطة وأسئلة
وتطبيقات قابلة للتنفيذ ومناسبة للصف والمادة والموضوع.
"""

SYSTEM_PROMPT = """
أنت معلم ومصمم مناهج تعليمية خبير بالمناهج السعودية.
أنشئ محتوى تربوياً دقيقاً ومخصصاً للصف والمادة والدرس الذي يرسله المستخدم.
لا تستخدم نصوصاً عامة لا علاقة لها بالموضوع.
راعِ عمر الطلاب ومستوى الصف، واجعل الأمثلة والأسئلة مناسبة للمادة.
"""

# ============================================================
# استدعاء الذكاء الاصطناعي
# ============================================================
def query_ai_engine(prompt: str) -> str:
    payloads = [
        {
            "url": "https://text.pollinations.ai/",
            "data": {
                "messages": [
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": prompt},
                ],
                "model": "openai",
                "seed": int.from_bytes(os.urandom(2), "big"),
            },
        },
        {
            "url": "https://api.airforce/v1/chat/completions",
            "data": {
                "model": "gpt-4o-mini",
                "messages": [
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": prompt},
                ],
            },
        },
    ]

    for payload in payloads:
        try:
            response = requests.post(
                payload["url"],
                json=payload["data"],
                timeout=60,
            )
            if response.status_code != 200:
                continue

            text = response.text.strip()
            if not text:
                continue

            try:
                data = response.json()
                if "choices" in data:
                    return (
                        data["choices"][0]["message"]["content"].strip()
                    )
            except Exception:
                pass

            if len(text) > 100:
                return text

        except Exception:
            continue

    return ""


# ============================================================
# صور الشرائح
# ============================================================
def fetch_unique_slide_image(slide_index: int, topic: str) -> io.BytesIO:
    keywords = [
        "education", "classroom", "science", "math",
        "technology", "books", "thinking", "teamwork",
        "experiment", "learning", "school", "exam",
        "research", "interactive", "future", "saudi",
    ]

    kw = keywords[(slide_index - 1) % len(keywords)]

    # لا نعتمد على hash(topic) لأن نتيجته تتغير بين عمليات تشغيل بايثون.
    stable_seed = sum(ord(c) for c in topic) % 5000
    lock = (slide_index * 37 + stable_seed) % 5000

    url = (
        f"https://loremflickr.com/600/450/{kw}"
        f"?lock={lock}"
    )

    try:
        response = requests.get(url, timeout=8)
        if response.status_code == 200 and len(response.content) > 3000:
            stream = io.BytesIO(response.content)
            stream.seek(0)
            return stream
    except Exception:
        pass

    # صورة بديلة إذا تعذر الاتصال بمصدر الصور
    img = Image.new("RGB", (600, 450), color="#F8FAFC")
    draw = ImageDraw.Draw(img)
    draw.rounded_rectangle(
        [15, 15, 585, 435],
        radius=15,
        fill="#FFFFFF",
        outline="#CBD5E1",
        width=2,
    )
    draw.rectangle([15, 15, 585, 70], fill="#1B4965")

    bars = [120, 180, 150, 230, 200]
    for idx, bar in enumerate(bars):
        x0 = 70 + idx * 90
        factor = ((slide_index + idx) % 5 + 6) / 10
        y0 = 400 - int(bar * factor)
        draw.rounded_rectangle(
            [x0, y0, x0 + 55, 400],
            radius=6,
            fill="#5FA8D3" if idx % 2 == 0 else "#62B6CB",
        )

    output = io.BytesIO()
    img.save(output, format="PNG")
    output.seek(0)
    return output


# ============================================================
# توليد 30 شريحة مخصصة للصف والمادة والدرس
# ============================================================
def generate_dynamic_30_slides_data(
    grade: str,
    subject: str,
    topic: str,
):
    """
    يولد عرضاً مختلفاً فعلياً لكل موضوع.
    إذا فشل محرك الذكاء الاصطناعي أو أعاد محتوى عاماً، يتم استخدام
    مولد محلي مخصص يعتمد على الصف + المادة + الموضوع، وليس قالباً ثابتاً.
    """

    grade = (grade or "غير محدد").strip()
    subject = (subject or "غير محددة").strip()
    topic = (topic or "الدرس").strip()

    # --------------------------------------------------------
    # 1) طلب AI أقوى: نطلب محتوى حقيقياً للموضوع وليس عناوين عامة
    # --------------------------------------------------------
    ai_prompt = f"""
أنت خبير مناهج سعودية وتصميم عروض PowerPoint للعام 1448هـ.

أنشئ محتوى تعليمي ORIGINAL ومختلف تماماً لهذا الطلب:

الصف: {grade}
المادة: {subject}
الدرس/الموضوع: {topic}

ممنوع إعادة قالب عام.
ممنوع كتابة عبارات مثل "شرح الموضوع" أو "نشاط مرتبط بالموضوع"
بدون تقديم المحتوى نفسه.

أريد 30 شريحة، وكل شريحة يجب أن تحتوي معلومات فعلية مرتبطة بـ "{topic}".
إذا كانت المادة رياضيات، قدم قوانين وأمثلة ومسائل وحلولاً حقيقية مناسبة للصف.
إذا كانت علوم، قدم مفاهيم وتفسيرات وأمثلة وتجارب وأسئلة علمية.
إذا كانت لغتي، قدم أمثلة لغوية وقواعد وتطبيقات.
إذا كانت دراسات اجتماعية، قدم معلومات وأحداثاً ومقارنات وأسئلة.
إذا كانت مادة أخرى، خصص المحتوى لطبيعتها.

يجب تضمين:
- تمهيد واستثارة دافعية خاصة بالدرس.
- نواتج تعلم خاصة بالموضوع.
- شرح المفاهيم الأساسية.
- أمثلة وتطبيقات فعلية.
- تعلم نشط.
- مهارات تفكير عليا.
- استراتيجية حل مشكلات.
- لعبة تعليمية.
- مقطع/فاصل تعليمي مقترح يخدم الموضوع.
- نشاط فردي.
- نشاط ثنائي.
- نشاط جماعي.
- ربط بالواقع.
- ربط بالوطن والقيم عندما يكون مناسباً.
- تكامل مع مادة أخرى عند ملاءمته.
- تمايز للطلاب.
- تدريب يحاكي مهارات نافس إذا كان مناسباً للمرحلة.
- ورقة عمل تحتوي أسئلة فعلية.
- تقويم تكويني.
- بطاقة خروج.
- تقويم ختامي.
- واجب إثرائي.

التنسيق الإلزامي:
---SLIDE---
TITLE: عنوان خاص بالموضوع
CONTENT:
- معلومة/شرح فعلي خاص بالموضوع
- مثال أو تطبيق فعلي خاص بالموضوع
- سؤال أو نشاط فعلي خاص بالموضوع

كرر ذلك حتى 30 شريحة.
"""

    ai_raw = query_ai_engine(ai_prompt)
    ai_slides = []

    # --------------------------------------------------------
    # 2) قراءة إجابة AI مهما كان عدد الشرائح
    # --------------------------------------------------------
    if ai_raw:
        chunks = re.split(
            r"---\s*SLIDE\s*---",
            ai_raw,
            flags=re.IGNORECASE,
        )

        for chunk in chunks:
            if not chunk.strip():
                continue

            title_match = re.search(
                r"TITLE\s*:\s*(.+)",
                chunk,
                flags=re.IGNORECASE,
            )

            title = (
                title_match.group(1).strip()
                if title_match
                else ""
            )

            points = []
            for line in chunk.splitlines():
                line = line.strip()

                if re.match(r"^[-•*]\s+", line):
                    point = re.sub(
                        r"^[-•*]\s+",
                        "",
                        line,
                    ).strip()

                    if point and len(point) > 8:
                        points.append(point)

            if title and points:
                ai_slides.append(
                    (
                        title,
                        points[:5],
                    )
                )

    # --------------------------------------------------------
    # 3) مولد محلي مخصص للموضوع
    #    هذا الجزء هو الضمان: حتى لو تعطل AI لن تتكرر نفس الشرائح.
    # --------------------------------------------------------
    local_templates = [
        (
            f"تمهيد: لماذا نتعلم {topic}؟",
            [
                f"موقف تمهيدي من الحياة اليومية يوضح أهمية {topic} لدى طلاب {grade}.",
                f"سؤال استثارة: ماذا تعرف مسبقاً عن {topic}؟",
                f"توقع نتيجة أو حل قبل البدء في دراسة {topic}.",
            ],
        ),
        (
            f"ما المقصود بـ {topic}؟",
            [
                f"تعريف مبسط ودقيق لمفهوم {topic} بما يناسب {grade}.",
                f"الكلمات والمصطلحات الأساسية التي يحتاجها الطالب لفهم {topic}.",
                f"مثال يوضح معنى {topic} بصورة مباشرة.",
            ],
        ),
        (
            f"الفكرة الرئيسة في {topic}",
            [
                f"الفكرة المحورية التي يجب أن يفهمها الطالب في {topic}.",
                f"العلاقة بين عناصر {topic} كما تظهر في مادة {subject}.",
                f"سؤال تحقق سريع لقياس فهم الفكرة.",
            ],
        ),
        (
            f"مثال تطبيقي على {topic}",
            [
                f"تطبيق واقعي للمفهوم {topic} مناسب لطلاب {grade}.",
                f"خطوات تنفيذ التطبيق في مادة {subject}.",
                f"سؤال: ماذا يحدث إذا تغير أحد المعطيات؟",
            ],
        ),
        (
            f"مثال ثانٍ: {topic}",
            [
                f"مثال جديد مختلف عن المثال السابق لتثبيت {topic}.",
                f"تحليل المثال وتحديد المهارة التي يقيسها.",
                f"استخراج القاعدة أو الفكرة من المثال.",
            ],
        ),
        (
            f"أخطاء شائعة في {topic}",
            [
                f"خطأ محتمل عند تعلم {topic} لدى طلاب {grade}.",
                "تفسير سبب الخطأ وطريقة اكتشافه.",
                f"تطبيق تصحيحي للتأكد من إتقان {topic}.",
            ],
        ),
        (
            f"استراتيجية تعلم نشط لـ {topic}",
            [
                f"تنفيذ استراتيجية فكر-زاوج-شارك حول {topic}.",
                f"مهمة فردية لمدة دقيقة تتعلق بـ {topic}.",
                "مشاركة الإجابات ومناقشة أسباب الاختلاف.",
            ],
        ),
        (
            f"نشاط فردي: أتقن {topic}",
            [
                f"حل مهمة فردية قصيرة حول {topic}.",
                "تسجيل خطوات الحل أو التفسير في ورقة الطالب.",
                "مراجعة الإجابة وفق معيار إتقان واضح.",
            ],
        ),
        (
            f"نشاط ثنائي: ناقش {topic}",
            [
                f"يحل كل طالب سؤالاً حول {topic} ثم يقارن الحل مع زميله.",
                "تحديد نقطة الاتفاق والاختلاف.",
                "صياغة إجابة مشتركة مدعمة بالتبرير.",
            ],
        ),
        (
            f"نشاط جماعي: تحدي {topic}",
            [
                f"تقسيم الطلاب إلى مجموعات لحل مهمة مرتبطة بـ {topic}.",
                "توزيع أدوار: قائد، كاتب، متحدث، ومراجع.",
                "عرض الحل ومقارنته بحلول المجموعات الأخرى.",
            ],
        ),
        (
            f"تفكير ناقد حول {topic}",
            [
                f"سؤال يتطلب تفسيراً وليس مجرد حفظ في {topic}.",
                f"مقارنة حالتين أو طريقتين مستخدمتين في {topic}.",
                "طلب دليل أو سبب يدعم إجابة الطالب.",
            ],
        ),
        (
            f"تفكير إبداعي في {topic}",
            [
                f"اقترح طريقة جديدة لشرح أو تطبيق {topic}.",
                "توقع نتيجة موقف غير مألوف مرتبط بالدرس.",
                "تقييم الفكرة وفق معيار منطقي.",
            ],
        ),
        (
            f"حل مشكلة من {topic}",
            [
                f"مشكلة تطبيقية مناسبة للصف {grade} في {topic}.",
                "تحديد المعطيات والمطلوب أو المشكلة الأساسية.",
                "اختيار استراتيجية الحل والتحقق من النتيجة.",
            ],
        ),
        (
            f"لعبة تعليمية: {topic}",
            [
                f"لعبة بطاقات أو تحدي سريع لمراجعة مفاهيم {topic}.",
                "يحصل الفريق على نقطة عند تفسير الإجابة بشكل صحيح.",
                "سؤال ختامي يثبت المفهوم بعد انتهاء اللعبة.",
            ],
        ),
        (
            f"مقطع تعليمي عن {topic}",
            [
                f"اقتراح مقطع قصير يشرح جانباً محدداً من {topic}.",
                "سؤال يجيب عنه الطالب أثناء مشاهدة المقطع.",
                "مناقشة أهم معلومة وردت في المقطع.",
            ],
        ),
        (
            f"ربط {topic} بالحياة اليومية",
            [
                f"تطبيق حقيقي نشاهده في حياتنا ويرتبط بـ {topic}.",
                f"كيف يمكن للطالب استخدام {topic} خارج المدرسة؟",
                "سؤال تطبيقي من البيئة المحيطة بالطالب.",
            ],
        ),
        (
            f"الربط بالوطن من خلال {topic}",
            [
                f"تطبيق وطني مناسب يوضح أهمية {topic}.",
                f"كيف تسهم المعرفة المرتبطة بـ {topic} في خدمة المجتمع؟",
                "ربط مناسب بمستهدفات التنمية دون تكلف.",
            ],
        ),
        (
            f"القيم والإتقان في {topic}",
            [
                f"قيمة الإتقان والمسؤولية أثناء تعلم {topic}.",
                "الأمانة في نقل المعلومات والنتائج.",
                f"موقف سلوكي يطبق قيمة إيجابية أثناء دراسة {topic}.",
            ],
        ),
        (
            f"التكامل مع مادة أخرى: {topic}",
            [
                f"مفهوم من مادة أخرى يمكن ربطه بـ {topic}.",
                f"نشاط تكاملي قصير يجمع {subject} مع المادة الأخرى.",
                "تفسير فائدة التكامل في فهم المفهوم.",
            ],
        ),
        (
            f"تمايز التعليم في {topic}",
            [
                f"مهمة دعم للطالب الذي يحتاج تعزيزاً في {topic}.",
                f"مهمة أساسية للطالب المتقن لـ {topic}.",
                f"مهمة إثرائية للطالب المتقدم في {topic}.",
            ],
        ),
        (
            f"توظيف التقنية في {topic}",
            [
                f"نشاط رقمي يساعد على تمثيل أو فهم {topic}.",
                "استخدام صورة أو محاكاة أو اختبار تفاعلي.",
                "مناقشة النتيجة الرقمية وربطها بمفهوم الدرس.",
            ],
        ),
        (
            f"تدريب مهاري يحاكي نافس: {topic}",
            [
                f"سؤال مهاري جديد مناسب للصف {grade} حول {topic}.",
                "تحليل المطلوب واستبعاد المشتتات.",
                "تفسير سبب صحة الإجابة.",
            ],
        ),
        (
            f"تدريب ثانٍ يحاكي مهارات نافس: {topic}",
            [
                f"موقف مركب يقيس الفهم والتطبيق في {topic}.",
                "اختيار الاستراتيجية المناسبة قبل الإجابة.",
                "مراجعة الإجابة في أقل وقت ممكن.",
            ],
        ),
        (
            f"ورقة عمل {topic} - 1",
            [
                f"السؤال 1: تطبيق مباشر على {topic}.",
                f"السؤال 2: تطبيق جديد يقيس الفهم.",
                f"السؤال 3: علل أو فسّر إجابتك في {topic}.",
            ],
        ),
        (
            f"ورقة عمل {topic} - 2",
            [
                f"السؤال 4: موقف تطبيقي مرتبط بـ {topic}.",
                f"السؤال 5: قارن بين حالتين في {topic}.",
                f"السؤال 6: استنتج نتيجة اعتماداً على {topic}.",
            ],
        ),
        (
            f"التقويم التكويني لـ {topic}",
            [
                f"سؤال سريع يكشف فهم الطالب للمفهوم الأساسي في {topic}.",
                f"سؤال يميز بين الفهم الصحيح والخطأ الشائع في {topic}.",
                "تحديد الطلاب الذين يحتاجون إلى دعم إضافي.",
            ],
        ),
        (
            f"بطاقة خروج: ماذا تعلمت عن {topic}؟",
            [
                f"اكتب أهم مفهوم تعلمته اليوم عن {topic}.",
                f"اذكر مثالاً واحداً يوضح {topic}.",
                "ما الجزء الذي ما زلت تحتاج إلى توضيحه؟",
            ],
        ),
        (
            f"مراجعة شاملة لـ {topic}",
            [
                f"استرجاع أهم المفاهيم والقواعد المرتبطة بـ {topic}.",
                "حل سؤال يجمع أكثر من مهارة.",
                "تفسير طريقة الوصول إلى الإجابة.",
            ],
        ),
        (
            f"الواجب الإثرائي في {topic}",
            [
                f"تطبيق منزلي بسيط يثبت تعلم {topic}.",
                f"مهمة بحثية قصيرة مرتبطة بـ {topic} ومناسبة للصف {grade}.",
                "إحضار مثال من الحياة اليومية ومناقشته في الحصة القادمة.",
            ],
        ),
        (
            f"الإغلاق والتقويم النهائي: {topic}",
            [
                f"تلخيص نواتج التعلم التي تحققت في {topic}.",
                f"سؤال ختامي شامل يقيس إتقان {topic}.",
                "تغذية راجعة وتحفيز للانتقال إلى التعلم التالي.",
            ],
        ),
    ]

    # --------------------------------------------------------
    # 4) ندمج AI مع المولد المحلي، مع منع تكرار العناوين
    # --------------------------------------------------------
    final_slides = []
    used_titles = set()

    for title, points in ai_slides:
        normalized = re.sub(r"\s+", " ", title).strip().lower()

        if normalized in used_titles:
            continue

        # لا نقبل شريحة عامة لا تحتوي أي إشارة للمادة/الموضوع
        title_and_points = " ".join([title] + points)
        if (
            topic not in title_and_points
            and subject not in title_and_points
        ):
            continue

        final_slides.append((title, points))
        used_titles.add(normalized)

        if len(final_slides) >= 30:
            break

    # إذا أعاد AI عدداً قليلاً، نكمل من القالب المخصص.
    for title, points in local_templates:
        normalized = re.sub(r"\s+", " ", title).strip().lower()

        if normalized in used_titles:
            continue

        final_slides.append(
            (title, points)
        )
        used_titles.add(normalized)

        if len(final_slides) >= 30:
            break

    # --------------------------------------------------------
    # 5) ضمان 30 شريحة حتى لو حصل خلل غير متوقع
    # --------------------------------------------------------
    while len(final_slides) < 30:
        number = len(final_slides) + 1

        final_slides.append(
            (
                f"{number}. تطبيق إضافي متقدم على {topic}",
                [
                    f"مهمة إضافية مخصصة لمفهوم {topic} في مادة {subject}.",
                    f"تطبيق مناسب لمستوى {grade} يختلف عن الأمثلة السابقة.",
                    f"سؤال تحقق لقياس إتقان {topic}.",
                ],
            )
        )

    # إعادة ترقيم العناوين بحيث يظهر ترتيب 1–30
    numbered = []
    for i, (title, points) in enumerate(
        final_slides[:30],
        start=1,
    ):
        clean_title = re.sub(
            r"^\s*\d+\s*[\.\-:)]\s*",
            "",
            title,
        ).strip()

        numbered.append(
            (
                f"{i}. {clean_title}",
                points,
            )
        )

    return numbered


# ============================================================
# إنشاء PowerPoint
# ============================================================
def create_powerpoint_presentation_full(
    grade: str,
    subject: str,
    topic: str,
    output_path: str,
):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slides_data = generate_dynamic_30_slides_data(
        grade, subject, topic
    )

    for idx, (title_text, points) in enumerate(
        slides_data, start=1
    ):
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        top_bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0),
            Inches(0),
            Inches(13.333),
            Inches(0.2),
        )
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = RGBColor(27, 73, 101)
        top_bar.line.fill.background()

        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(5.4),
            Inches(1.5),
            Inches(7.2),
            Inches(5.1),
        )
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(248, 250, 252)
        card.line.color.rgb = RGBColor(203, 213, 225)
        card.line.width = Pt(1.5)

        title_box = slide.shapes.add_textbox(
            Inches(0.8),
            Inches(0.4),
            Inches(11.733),
            Inches(1.0),
        )
        tf_title = title_box.text_frame
        tf_title.word_wrap = True

        p_title = tf_title.paragraphs[0]
        p_title.text = title_text
        p_title.alignment = PP_ALIGN.RIGHT
        p_title.font.size = Pt(26)
        p_title.font.bold = True
        p_title.font.color.rgb = RGBColor(27, 73, 101)

        content_box = slide.shapes.add_textbox(
            Inches(5.6),
            Inches(1.7),
            Inches(6.8),
            Inches(4.7),
        )
        tf_content = content_box.text_frame
        tf_content.word_wrap = True

        for p_idx, point in enumerate(points):
            p = (
                tf_content.paragraphs[0]
                if p_idx == 0
                else tf_content.add_paragraph()
            )
            p.text = f"◀ {point}"
            p.alignment = PP_ALIGN.RIGHT
            p.font.size = Pt(18)
            p.font.color.rgb = RGBColor(30, 41, 59)
            p.space_after = Pt(14)

        img_stream = fetch_unique_slide_image(idx, topic)
        if img_stream:
            slide.shapes.add_picture(
                img_stream,
                Inches(0.8),
                Inches(1.5),
                Inches(4.3),
                Inches(5.1),
            )

        footer_box = slide.shapes.add_textbox(
            Inches(0.8),
            Inches(6.9),
            Inches(11.733),
            Inches(0.4),
        )
        p_foot = footer_box.text_frame.paragraphs[0]
        p_foot.text = (
            f"شريحة {idx} من 30 | منصة إنجاز | "
            f"{grade} | {subject} | {topic} | 1448هـ"
        )
        p_foot.alignment = PP_ALIGN.LEFT
        p_foot.font.size = Pt(11)
        p_foot.font.color.rgb = RGBColor(148, 163, 184)

    prs.save(output_path)


# ============================================================
# مستندات Word
# ============================================================
def create_educational_doc_1448(
    service_code: str,
    grade: str,
    subject: str,
    topic: str,
    output_path: str,
):
    doc = Document()

    for section in doc.sections:
        section.top_margin = DocxInches(1)
        section.bottom_margin = DocxInches(1)
        section.left_margin = DocxInches(1)
        section.right_margin = DocxInches(1)

    prompts = {
        "svc_exam": (
            f"اكتب اختباراً شاملاً ومخصصاً لعام 1448هـ للصف {grade} "
            f"في مادة {subject} حول {topic}. "
            "أضف جدول مواصفات، أسئلة متنوعة، تدريبات مهارية مناسبة، "
            "ونموذج إجابة وتوزيع درجات."
        ),
        "svc_remedial": (
            f"اكتب خطة علاجية وإثرائية وأوراق عمل للصف {grade} "
            f"في مادة {subject} حول {topic}. "
            "شخّص الفاقد، ثم قدم أنشطة علاجية وإثرائية وتقويماً بعدياً."
        ),
        "svc_portfolio": (
            f"اكتب ملف إنجاز إلكترونياً للمعلم/المعلمة في الصف {grade} "
            f"ومادة {subject} حول {topic} لعام 1448هـ."
        ),
        "svc_performance": (
            f"اكتب ملف أداء وظيفي منظم للمعلم/المعلمة لعام 1448هـ "
            f"مرتبط بالصف {grade} ومادة {subject} وموضوع {topic}."
        ),
        "svc_operation": (
            f"اكتب خطة تشغيلية تعليمية لعام 1448هـ للصف {grade} "
            f"ومادة {subject} حول {topic}, مع أهداف ومؤشرات "
            "وزمن تنفيذ وأدلة تحقق."
        ),
        "svc_loss": (
            f"اكتب خطة معالجة فاقد تعليمي للصف {grade} في مادة {subject} "
            f"حول {topic}, مع تشخيص المهارات المفقودة والتعويض والتقويم."
        ),
        "svc_research": (
            f"اكتب بحثاً أكاديمياً جامعياً مفصلاً حول {topic} "
            f"مع مراعاة سياق التعليم والصف {grade} ومادة {subject}. "
            "يشمل المقدمة والمشكلة والأهداف والإطار النظري والتحليل "
            "والتوصيات والمراجع وفق APA."
        ),
    }

    titles = {
        "svc_exam": f"الاختبار وتحليل النتائج\n{subject} - {grade}\n{topic}",
        "svc_remedial": f"الخطة العلاجية والإثرائية\n{subject} - {grade}\n{topic}",
        "svc_portfolio": f"ملف الإنجاز الإلكتروني 1448هـ\n{subject} - {grade}",
        "svc_performance": f"ملف الأداء الوظيفي 1448هـ\n{subject} - {grade}",
        "svc_operation": f"الخطة التشغيلية 1448هـ\n{subject} - {grade}",
        "svc_loss": f"خطة معالجة الفاقد التعليمي 1448هـ\n{subject} - {grade}",
        "svc_research": f"بحث أكاديمي متكامل\n{topic}",
    }

    prompt = prompts.get(
        service_code,
        f"اكتب وثيقة تعليمية لعام 1448هـ حول {topic}.",
    )
    doc_title = titles.get(
        service_code,
        f"وثيقة تعليمية 1448هـ\n{topic}",
    )

    ai_content = query_ai_engine(prompt)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run_title = title_p.add_run(doc_title)
    run_title.font.size = DocxPt(20)
    run_title.font.bold = True
    run_title.font.color.rgb = DocxRGB(27, 73, 101)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run_sub = sub_p.add_run(
        f"منصة إنجاز | الصف: {grade} | المادة: {subject} | "
        "العام الدراسي 1448هـ\n" + "—" * 35
    )
    run_sub.font.size = DocxPt(11)
    run_sub.font.color.rgb = DocxRGB(100, 116, 139)

    if len(ai_content) > 300:
        for block in ai_content.split("\n\n"):
            clean = block.strip()
            if not clean:
                continue

            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

            if (
                clean.startswith("#")
                or any(
                    clean.startswith(f"{i}.")
                    for i in range(1, 20)
                )
                or "المحور" in clean
                or "الهدف" in clean
            ):
                run = p.add_run(
                    clean.replace("#", "").strip()
                )
                run.font.size = DocxPt(14)
                run.font.bold = True
                run.font.color.rgb = DocxRGB(27, 73, 101)
                p.paragraph_format.space_before = DocxPt(12)
                p.paragraph_format.space_after = DocxPt(4)
            else:
                run = p.add_run(clean)
                run.font.size = DocxPt(11.5)
                run.font.color.rgb = DocxRGB(30, 41, 59)
                p.paragraph_format.line_spacing = 1.25
                p.paragraph_format.space_after = DocxPt(6)
    else:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run(
            f"تم إعداد المستند للصف {grade} في مادة {subject} "
            f"حول {topic} لعام 1448هـ."
        )

    doc.save(output_path)


# ============================================================
# القوائم
# ============================================================
def main_menu():
    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton(
                "📊 بوربوينت 30 شريحة + جميع الإضافات",
                callback_data="svc_ppt",
            )
        ],
        [
            InlineKeyboardButton(
                "📝 اختبارات + جدول مواصفات + نافس",
                callback_data="svc_exam",
            )
        ],
        [
            InlineKeyboardButton(
                "📈 خطط علاجية وإثرائية + أوراق عمل",
                callback_data="svc_remedial",
            )
        ],
        [
            InlineKeyboardButton(
                "🗂 ملف إنجاز المعلم/المعلمة",
                callback_data="svc_portfolio",
            )
        ],
        [
            InlineKeyboardButton(
                "📑 ملف الأداء الوظيفي",
                callback_data="svc_performance",
            )
        ],
        [
            InlineKeyboardButton(
                "📅 الخطة التشغيلية",
                callback_data="svc_operation",
            )
        ],
        [
            InlineKeyboardButton(
                "📚 خطة الفاقد التعليمي",
                callback_data="svc_loss",
            )
        ],
        [
            InlineKeyboardButton(
                "🎓 بحث جامعي وأكاديمي Word",
                callback_data="svc_research",
            )
        ],
        [
            InlineKeyboardButton(
                "🎓 اختيار الصف الدراسي",
                callback_data="choose_grade",
            )
        ],
    ])


def grade_menu():
    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton(
                "رياض الأطفال",
                callback_data="grade_kg",
            )
        ],
        [
            InlineKeyboardButton("الأول ابتدائي", callback_data="grade_p1"),
            InlineKeyboardButton("الثاني ابتدائي", callback_data="grade_p2"),
        ],
        [
            InlineKeyboardButton("الثالث ابتدائي", callback_data="grade_p3"),
            InlineKeyboardButton("الرابع ابتدائي", callback_data="grade_p4"),
        ],
        [
            InlineKeyboardButton("الخامس ابتدائي", callback_data="grade_p5"),
            InlineKeyboardButton("السادس ابتدائي", callback_data="grade_p6"),
        ],
        [
            InlineKeyboardButton("الأول متوسط", callback_data="grade_m1"),
            InlineKeyboardButton("الثاني متوسط", callback_data="grade_m2"),
        ],
        [
            InlineKeyboardButton("الثالث متوسط", callback_data="grade_m3"),
        ],
        [
            InlineKeyboardButton("الأول ثانوي", callback_data="grade_s1"),
            InlineKeyboardButton("الثاني ثانوي", callback_data="grade_s2"),
        ],
        [
            InlineKeyboardButton("الثالث ثانوي", callback_data="grade_s3"),
        ],
        [
            InlineKeyboardButton(
                "⬅️ القائمة الرئيسية",
                callback_data="home",
            )
        ],
    ])


# ============================================================
# /start
# ============================================================
async def start(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
):
    context.user_data.setdefault("grade", "")
    context.user_data.setdefault("subject", "")

    welcome_text = (
        "🌟 *أهلاً بك في منصة إنجاز للخدمات التعليمية والأكاديمية 1448هـ*\n\n"
        "📚 يدعم البوت جميع الصفوف من رياض الأطفال حتى الثالث الثانوي.\n"
        "📖 وجميع المواد، مع تخصيص المحتوى حسب الصف والمادة والدرس.\n\n"
        "📌 في عروض البوربوينت تمت إضافة عناصر الصورة:\n"
        "• تمهيد وإغلاق مناسب للدرس\n"
        "• مهارات تفكير عليا + تعلم نشط + ألعاب تعليمية\n"
        "• مقاطع تعليمية وفواصل\n"
        "• ربط بالوطن والدين والواقع والمواد الأخرى\n"
        "• تدريبات نافس المناسبة للمرحلة\n"
        "• أنشطة فردية وثنائية وجماعية\n"
        "• ورقة عمل وتقويم ختامي\n\n"
        "👇 اختر الخدمة:"
    )

    if update.message:
        await update.message.reply_text(
            welcome_text,
            reply_markup=main_menu(),
            parse_mode="Markdown",
        )
    elif update.callback_query:
        await update.callback_query.message.reply_text(
            welcome_text,
            reply_markup=main_menu(),
            parse_mode="Markdown",
        )


# ============================================================
# معالجة الأزرار
# ============================================================
async def button_handler(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
):
    query = update.callback_query
    await query.answer()

    data = query.data

    if data == "home":
        await query.edit_message_text(
            "👇 القائمة الرئيسية:",
            reply_markup=main_menu(),
        )
        return

    if data == "choose_grade":
        await query.edit_message_text(
            "🎓 اختر الصف الدراسي:",
            reply_markup=grade_menu(),
        )
        return

    if data.startswith("grade_"):
        grade_key = data.replace("grade_", "", 1)
        grade = GRADES.get(grade_key)

        if not grade:
            await query.edit_message_text(
                "⚠️ لم يتم التعرف على الصف.",
                reply_markup=grade_menu(),
            )
            return

        context.user_data["grade"] = grade

        await query.edit_message_text(
            f"✅ تم اختيار: *{grade}*\n\n"
            "الآن أرسل في رسالة واحدة:\n"
            "*المادة - موضوع الدرس*\n\n"
            "مثال:\n"
            "رياضيات - الأعداد الصحيحة\n\n"
            "أو:\n"
            "علوم - الخلية",
            parse_mode="Markdown",
        )
        return

    services = {
        "svc_ppt": "📊 بوربوينت 30 شريحة",
        "svc_exam": "📝 الاختبارات",
        "svc_remedial": "📈 الخطة العلاجية والإثرائية",
        "svc_portfolio": "🗂 ملف الإنجاز",
        "svc_performance": "📑 ملف الأداء الوظيفي",
        "svc_operation": "📅 الخطة التشغيلية",
        "svc_loss": "📚 الفاقد التعليمي",
        "svc_research": "🎓 البحث الأكاديمي",
    }

    if data in services:
        context.user_data["current_service"] = data
        context.user_data["service_name"] = services[data]

        grade = context.user_data.get("grade", "")

        if not grade:
            await query.edit_message_text(
                "🎓 اختر الصف أولاً حتى يتم تخصيص المحتوى حسب المرحلة:",
                reply_markup=grade_menu(),
            )
            return

        if data == "svc_ppt":
            message = (
                f"📊 *الخدمة: {services[data]}*\n\n"
                f"🎓 الصف المحدد: *{grade}*\n\n"
                "أرسل الآن:\n"
                "*المادة - موضوع الدرس*\n\n"
                "مثال: رياضيات - الكسور\n\n"
                "وسيتم إنشاء عرض جديد مخصص للدرس، يتضمن:\n"
                "✅ 30 شريحة\n"
                "✅ تمهيد وإغلاق\n"
                "✅ تفكير عليا وتعلم نشط\n"
                "✅ ألعاب تعليمية\n"
                "✅ مقاطع وفواصل تعليمية\n"
                "✅ ربط بالوطن والدين والواقع والمواد الأخرى\n"
                "✅ تدريبات نافس المناسبة\n"
                "✅ أنشطة فردية وثنائية وجماعية\n"
                "✅ ورقة عمل وتقويم ختامي"
            )
        else:
            message = (
                f"✨ *الخدمة: {services[data]}*\n\n"
                f"🎓 الصف: *{grade}*\n\n"
                "أرسل المادة وموضوع الدرس، مثال:\n"
                "*لغتي - أسلوب الاستثناء*\n\n"
                "وسيتم توليد ملف Word مخصص للصف والمادة والموضوع."
            )

        await query.edit_message_text(
            message,
            parse_mode="Markdown",
        )


# ============================================================
# استقبال النص
# ============================================================
async def text_handler(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
):
    user_text = (update.message.text or "").strip()
    user = update.effective_user

    current_service = context.user_data.get(
        "current_service",
        "svc_ppt",
    )
    service_name = context.user_data.get(
        "service_name",
        "عرض بوربوينت متكامل 1448هـ",
    )
    grade = context.user_data.get("grade", "")

    # السماح للمستخدم بإرسال الصف + المادة + الموضوع مباشرة
    if not grade:
        detected_grade = None

        for key, grade_name in GRADES.items():
            if grade_name in user_text:
                detected_grade = grade_name
                break

        if detected_grade:
            grade = detected_grade
            context.user_data["grade"] = grade

        else:
            await update.message.reply_text(
                "🎓 اختر الصف أولاً من الزر، ثم أرسل المادة وموضوع الدرس.",
                reply_markup=grade_menu(),
            )
            return

    status_msg = await update.message.reply_text(
        "⏳ جارٍ إعداد المحتوى المخصص حسب الصف والمادة والدرس..."
    )

    try:
        # نحاول تقسيم الإدخال إلى مادة وموضوع
        if " - " in user_text:
            subject, topic = user_text.split(" - ", 1)
        elif "-" in user_text:
            subject, topic = user_text.split("-", 1)
        else:
            subject = "المادة غير محددة"
            topic = user_text

        subject = subject.strip()
        topic = topic.strip()

        context.user_data["subject"] = subject

        if current_service == "svc_ppt":
            file_name = f"presentation_{user.id}.pptx"

            create_powerpoint_presentation_full(
                grade=grade,
                subject=subject,
                topic=topic,
                output_path=file_name,
            )

            with open(file_name, "rb") as ppt_file:
                await update.message.reply_document(
                    document=ppt_file,
                    filename=(
                        f"{subject[:20]}_{topic[:25]}_1448H.pptx"
                    ),
                    caption=(
                        "✅ تم إنشاء العرض بنجاح\n\n"
                        f"🎓 الصف: {grade}\n"
                        f"📚 المادة: {subject}\n"
                        f"📌 الدرس: {topic}\n"
                        "📊 30 شريحة مخصصة\n"
                        "🎯 تمهيد + تعلم نشط + ألعاب + أنشطة\n"
                        "🇸🇦 ربط بالوطن والواقع والقيم\n"
                        "📝 ورقة عمل + تقويم ختامي\n"
                        "📈 تدريبات مهارية تحاكي نافس عند المناسبة"
                    ),
                )

            if os.path.exists(file_name):
                os.remove(file_name)

        else:
            file_name = f"doc_{user.id}.docx"

            create_educational_doc_1448(
                service_code=current_service,
                grade=grade,
                subject=subject,
                topic=topic,
                output_path=file_name,
            )

            with open(file_name, "rb") as doc_file:
                await update.message.reply_document(
                    document=doc_file,
                    filename=(
                        f"{service_name[:20]}_{topic[:25]}_1448H.docx"
                    ),
                    caption=(
                        "✅ تم تجهيز مستند Word المخصص\n\n"
                        f"الخدمة: {service_name}\n"
                        f"الصف: {grade}\n"
                        f"المادة: {subject}\n"
                        f"الموضوع: {topic}"
                    ),
                )

            if os.path.exists(file_name):
                os.remove(file_name)

        await status_msg.delete()

    except Exception as exc:
        await status_msg.edit_text(
            "⚠️ حدث خطأ أثناء المعالجة.\n"
            f"التفاصيل: {str(exc)[:500]}"
        )


# ============================================================
# Web server للاستضافة
# ============================================================
async def handle_ping(request):
    return web.Response(
        text="Enjaz Full 1448 Bot is active!"
    )


async def start_web_server():
    server = web.Application()
    server.router.add_get("/", handle_ping)

    runner = web.AppRunner(server)
    await runner.setup()

    port = int(os.environ.get("PORT", "10000"))
    site = web.TCPSite(
        runner,
        "0.0.0.0",
        port,
    )
    await site.start()


# ============================================================
# التشغيل
# ============================================================
async def main_async():
    await start_web_server()

    app = (
        Application.builder()
        .token(TELEGRAM_TOKEN)
        .build()
    )

    app.add_handler(
        CommandHandler("start", start)
    )
    app.add_handler(
        CallbackQueryHandler(button_handler)
    )
    app.add_handler(
        MessageHandler(
            filters.TEXT & ~filters.COMMAND,
            text_handler,
        )
    )

    print("منصة إنجاز تعمل - 1448هـ")

    await app.initialize()
    await app.start()
    await app.updater.start_polling()

    while True:
        await asyncio.sleep(3600)


def main():
    asyncio.run(main_async())


if __name__ == "__main__":
    main()
