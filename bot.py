import os
import re
import logging
import logging.handlers
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler, CallbackQueryHandler,
    ConversationHandler, ContextTypes, filters, PicklePersistence
)
from cv_generator import generate_cv, convert_to_pdf
from dotenv import load_dotenv

load_dotenv()
TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")

# ── Logging Setup ──────────────────────
os.makedirs("logs", exist_ok=True)
os.makedirs("bot_data", exist_ok=True)

logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO,
    handlers=[
        logging.StreamHandler(),
        logging.handlers.RotatingFileHandler(
            "logs/bot.log", maxBytes=5*1024*1024, backupCount=3, encoding='utf-8'
        )
    ]
)
logger = logging.getLogger(__name__)

# ── States ─────────────────────────────
(
    CV_LANG,
    NAME, JOB_TITLE, EMAIL, PHONE, LOCATION, LINKEDIN, GITHUB,
    SUMMARY,
    EXP_TITLE, EXP_COMPANY, EXP_DATE, EXP_BULLETS, EXP_MORE,
    EDU_DEGREE, EDU_UNIVERSITY, EDU_DATE,
    SKILLS,
    CERT_ADD, CERT_NAME, CERT_ISSUER, CERT_DATE, CERT_MORE,
    LANGUAGES,
    EXPORT_FORMAT
) = range(25)

# ── Keyboards ──────────────────────────
def lang_kb():
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("🇸🇦 عربي", callback_data="lang_ar"),
        InlineKeyboardButton("🇬🇧 English", callback_data="lang_en")
    ]])

def yes_no_kb():
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ نعم", callback_data="yes"),
        InlineKeyboardButton("❌ لا", callback_data="no")
    ]])

def export_kb():
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("📝 Word (.docx)", callback_data="word"),
        InlineKeyboardButton("📕 PDF", callback_data="pdf")
    ]])

def skip_kb():
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("⏭️ تخطي / Skip", callback_data="skip")
    ]])

# ── Validation ─────────────────────────
EMAIL_REGEX = re.compile(r'^[\w\.\+\-]+@[\w\-]+\.[a-zA-Z]{2,}$')
PHONE_REGEX = re.compile(r'^\+?[\d\s\-\(\)]{7,20}$')

def is_valid_email(text: str) -> bool:
    return bool(EMAIL_REGEX.match(text.strip()))

def is_valid_phone(text: str) -> bool:
    return bool(PHONE_REGEX.match(text.strip()))

# ════════════════════════════════════════
#         COMMANDS
# ════════════════════════════════════════

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    context.user_data['experiences'] = []
    context.user_data['certificates'] = []
    user = update.effective_user
    logger.info(f"User {user.id} ({user.username}) started new CV")
    await update.message.reply_text(
        "👋 *Welcome / مرحباً!*\n\n"
        "🌐 *اختر لغة الـ CV / Choose CV language:*",
        reply_markup=lang_kb(),
        parse_mode='Markdown'
    )
    return CV_LANG


async def get_cv_lang(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    lang = 'ar' if update.callback_query.data == 'lang_ar' else 'en'
    context.user_data['lang'] = lang

    if lang == 'ar':
        prompt = (
            "✅ *تم اختيار العربية*\n\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            "📝 *ما اسمك الكامل؟*"
        )
    else:
        prompt = (
            "✅ *English selected*\n\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            "📝 *What is your full name?*"
        )

    await update.callback_query.message.reply_text(prompt, parse_mode='Markdown')
    return NAME


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📖 *مساعدة - بوت إنشاء CV*\n\n"
        "━━━━━━━━━━━━━━━━━━━━\n"
        "🤖 *ما الذي يفعله هذا البوت؟*\n"
        "يساعدك في إنشاء سيرة ذاتية احترافية\n"
        "بصيغة Word أو PDF خلال دقائق.\n\n"
        "━━━━━━━━━━━━━━━━━━━━\n"
        "⌨️ *الأوامر المتاحة:*\n"
        "/start — بدء إنشاء CV جديد\n"
        "/cancel — إلغاء العملية الحالية\n"
        "/help — عرض هذه المساعدة\n\n"
        "━━━━━━━━━━━━━━━━━━━━\n"
        "📋 *ما الذي ستحتاجه؟*\n"
        "• اسمك الكامل والعنوان الوظيفي\n"
        "• بريدك الإلكتروني ورقم هاتفك\n"
        "• خبراتك العملية وتعليمك\n"
        "• مهاراتك ولغاتك\n"
        "• شهاداتك (اختياري)\n",
        parse_mode='Markdown'
    )


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text(
        "❌ تم إلغاء العملية.\n"
        "يمكنك البدء من جديد بكتابة /start"
    )
    return ConversationHandler.END


# ════════════════════════════════════════
#         PERSONAL INFO
# ════════════════════════════════════════

async def get_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    name = update.message.text.strip()
    if len(name) < 2:
        await update.message.reply_text("⚠️ الرجاء إدخال اسم صحيح.")
        return NAME
    context.user_data['name'] = name
    await update.message.reply_text(
        "💼 *ما عنوانك الوظيفي؟*\n"
        "_(مثال: Data Analyst | Project Manager)_",
        parse_mode='Markdown'
    )
    return JOB_TITLE


async def get_job_title(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['job_title'] = update.message.text.strip()
    await update.message.reply_text("📧 *ما بريدك الإلكتروني؟*", parse_mode='Markdown')
    return EMAIL


async def get_email(update: Update, context: ContextTypes.DEFAULT_TYPE):
    email = update.message.text.strip()
    if not is_valid_email(email):
        await update.message.reply_text(
            "⚠️ البريد الإلكتروني غير صحيح.\n"
            "مثال: name@gmail.com"
        )
        return EMAIL
    context.user_data['email'] = email
    await update.message.reply_text("📞 *ما رقم هاتفك؟*", parse_mode='Markdown')
    return PHONE


async def get_phone(update: Update, context: ContextTypes.DEFAULT_TYPE):
    phone = update.message.text.strip()
    if not is_valid_phone(phone):
        await update.message.reply_text(
            "⚠️ رقم الهاتف غير صحيح.\n"
            "مثال: +963912345678"
        )
        return PHONE
    context.user_data['phone'] = phone
    await update.message.reply_text(
        "📍 *ما موقعك الجغرافي؟*\n_(مثال: دمشق، سوريا)_",
        parse_mode='Markdown'
    )
    return LOCATION


async def get_location(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['location'] = update.message.text.strip()
    await update.message.reply_text(
        "🔗 *ما رابط LinkedIn الخاص بك؟*\n"
        "_(اضغط تخطي إن لم يكن لديك)_",
        reply_markup=skip_kb(),
        parse_mode='Markdown'
    )
    return LINKEDIN


async def get_linkedin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.callback_query:
        await update.callback_query.answer()
        context.user_data['linkedin'] = ''
        msg = update.callback_query.message
    else:
        context.user_data['linkedin'] = update.message.text.strip()
        msg = update.message
    await msg.reply_text(
        "💻 *ما رابط GitHub الخاص بك؟*\n"
        "_(اضغط تخطي إن لم يكن لديك)_",
        reply_markup=skip_kb(),
        parse_mode='Markdown'
    )
    return GITHUB


async def get_github(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.callback_query:
        await update.callback_query.answer()
        context.user_data['github'] = ''
        msg = update.callback_query.message
    else:
        context.user_data['github'] = update.message.text.strip()
        msg = update.message
    await msg.reply_text(
        "📄 *اكتب ملخصك المهني (Professional Summary)*\n\n"
        "_جملتان تصفان خبرتك وأهدافك المهنية_",
        parse_mode='Markdown'
    )
    return SUMMARY


# ════════════════════════════════════════
#         SUMMARY
# ════════════════════════════════════════

async def get_summary(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['summary'] = update.message.text.strip()
    await update.message.reply_text(
        "💼 *الآن سنضيف خبراتك العملية*\n\n"
        "━━━━━━━━━━━━━━━━━━━━\n"
        "🏢 *ما المسمى الوظيفي في أول وظيفة؟*\n"
        "_(مثال: Project Manager)_",
        parse_mode='Markdown'
    )
    return EXP_TITLE


# ════════════════════════════════════════
#         EXPERIENCE
# ════════════════════════════════════════

async def get_exp_title(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['current_exp'] = {'title': update.message.text.strip()}
    await update.message.reply_text("🏛️ *ما اسم الشركة أو المؤسسة؟*", parse_mode='Markdown')
    return EXP_COMPANY


async def get_exp_company(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['current_exp']['company'] = update.message.text.strip()
    await update.message.reply_text(
        "📅 *ما فترة العمل؟*\n_(مثال: 01/2024 – Present)_",
        parse_mode='Markdown'
    )
    return EXP_DATE


async def get_exp_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['current_exp']['date'] = update.message.text.strip()
    await update.message.reply_text(
        "📋 *اكتب مهامك ومسؤولياتك في هذه الوظيفة*\n\n"
        "_كل مهمة في سطر جديد:_\n\n"
        "مثال:\n"
        "إدارة ومتابعة المشاريع التنموية\n"
        "إعداد التقارير الدورية\n"
        "تحليل بيانات المستفيدين",
        parse_mode='Markdown'
    )
    return EXP_BULLETS


async def get_exp_bullets(update: Update, context: ContextTypes.DEFAULT_TYPE):
    bullets = [b.strip() for b in update.message.text.strip().split('\n') if b.strip()]
    context.user_data['current_exp']['bullets'] = bullets
    context.user_data['experiences'].append(context.user_data.pop('current_exp'))
    count = len(context.user_data['experiences'])
    await update.message.reply_text(
        f"✅ *تم إضافة الخبرة بنجاح!* ({count} خبرة)\n\n"
        "هل تريد إضافة خبرة عملية أخرى؟",
        reply_markup=yes_no_kb(),
        parse_mode='Markdown'
    )
    return EXP_MORE


async def get_exp_more(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    if update.callback_query.data == "yes":
        await update.callback_query.message.reply_text(
            "🏢 *المسمى الوظيفي للخبرة التالية؟*", parse_mode='Markdown'
        )
        return EXP_TITLE
    else:
        await update.callback_query.message.reply_text(
            "🎓 *التعليم*\n\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            "📚 *ما اسم شهادتك أو درجتك العلمية؟*\n"
            "_(مثال: Bachelor's Degree — Computer Engineering)_",
            parse_mode='Markdown'
        )
        return EDU_DEGREE


# ════════════════════════════════════════
#         EDUCATION
# ════════════════════════════════════════

async def get_edu_degree(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['edu_degree'] = update.message.text.strip()
    await update.message.reply_text("🏛️ *ما اسم الجامعة أو المعهد؟*", parse_mode='Markdown')
    return EDU_UNIVERSITY


async def get_edu_university(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['edu_university'] = update.message.text.strip()
    await update.message.reply_text(
        "📅 *ما سنوات الدراسة؟*\n_(مثال: 2018 – 2024)_", parse_mode='Markdown'
    )
    return EDU_DATE


async def get_edu_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['edu_date'] = update.message.text.strip()
    await update.message.reply_text(
        "🛠️ *المهارات*\n\n"
        "━━━━━━━━━━━━━━━━━━━━\n"
        "اكتب مهاراتك مفصولة بفاصلة:\n\n"
        "_مثال: Python, Excel, Data Analysis, Problem Solving_",
        parse_mode='Markdown'
    )
    return SKILLS


# ════════════════════════════════════════
#         SKILLS
# ════════════════════════════════════════

async def get_skills(update: Update, context: ContextTypes.DEFAULT_TYPE):
    skills = [s.strip() for s in update.message.text.strip().split(',') if s.strip()]
    if not skills:
        await update.message.reply_text("⚠️ الرجاء إدخال مهارة واحدة على الأقل.")
        return SKILLS
    context.user_data['skills'] = skills
    await update.message.reply_text(
        "📜 *الشهادات والدورات التدريبية*\n\n"
        "هل تريد إضافة شهادات أو دورات؟",
        reply_markup=yes_no_kb(),
        parse_mode='Markdown'
    )
    return CERT_ADD


# ════════════════════════════════════════
#         CERTIFICATES
# ════════════════════════════════════════

async def get_cert_add(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    if update.callback_query.data == "yes":
        await update.callback_query.message.reply_text(
            "📜 *ما اسم الشهادة أو الدورة؟*", parse_mode='Markdown'
        )
        return CERT_NAME
    else:
        await update.callback_query.message.reply_text(
            "🌍 *اللغات*\n\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            "اكتب لغاتك ومستوياتها مفصولة بفاصلة:\n\n"
            "_مثال: Arabic – Native, English – Intermediate_",
            parse_mode='Markdown'
        )
        return LANGUAGES


async def get_cert_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['current_cert'] = {'name': update.message.text.strip()}
    await update.message.reply_text(
        "🏛️ *من أصدر هذه الشهادة؟*\n_(مثال: Coursera, Damascus University)_",
        parse_mode='Markdown'
    )
    return CERT_ISSUER


async def get_cert_issuer(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['current_cert']['issuer'] = update.message.text.strip()
    await update.message.reply_text(
        "📅 *ما تاريخ الحصول عليها؟*\n_(مثال: 09/2025)_", parse_mode='Markdown'
    )
    return CERT_DATE


async def get_cert_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['current_cert']['date'] = update.message.text.strip()
    context.user_data['certificates'].append(context.user_data.pop('current_cert'))
    count = len(context.user_data['certificates'])
    await update.message.reply_text(
        f"✅ *تم إضافة الشهادة!* ({count} شهادة)\n\n"
        "هل تريد إضافة شهادة أخرى؟",
        reply_markup=yes_no_kb(),
        parse_mode='Markdown'
    )
    return CERT_MORE


async def get_cert_more(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    if update.callback_query.data == "yes":
        await update.callback_query.message.reply_text(
            "📜 *ما اسم الشهادة التالية؟*", parse_mode='Markdown'
        )
        return CERT_NAME
    else:
        await update.callback_query.message.reply_text(
            "🌍 *اللغات*\n\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            "اكتب لغاتك ومستوياتها مفصولة بفاصلة:\n\n"
            "_مثال: Arabic – Native, English – Intermediate_",
            parse_mode='Markdown'
        )
        return LANGUAGES


# ════════════════════════════════════════
#         LANGUAGES & EXPORT
# ════════════════════════════════════════

async def get_languages(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['languages'] = update.message.text.strip()

    # Show data summary before generating
    d = context.user_data
    exp_count  = len(d.get('experiences', []))
    cert_count = len(d.get('certificates', []))
    skill_count= len(d.get('skills', []))

    await update.message.reply_text(
        "🎉 *ممتاز! تم جمع جميع المعلومات.*\n\n"
        "━━━━━━━━━━━━━━━━━━━━\n"
        f"👤 *الاسم:* {d.get('name')}\n"
        f"💼 *الوظيفة:* {d.get('job_title')}\n"
        f"📧 *الإيميل:* {d.get('email')}\n"
        f"📞 *الهاتف:* {d.get('phone')}\n"
        f"📍 *الموقع:* {d.get('location')}\n"
        f"💼 *الخبرات:* {exp_count} خبرة\n"
        f"🛠️ *المهارات:* {skill_count} مهارة\n"
        f"📜 *الشهادات:* {cert_count} شهادة\n"
        "━━━━━━━━━━━━━━━━━━━━\n\n"
        "بأي صيغة تريد استلام CV الخاص بك؟",
        reply_markup=export_kb(),
        parse_mode='Markdown'
    )
    return EXPORT_FORMAT


async def get_export_format(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.callback_query.answer()
    fmt = update.callback_query.data  # 'word' or 'pdf'
    user = update.effective_user
    logger.info(f"User {user.id} requesting CV in {fmt} format")

    await update.callback_query.message.reply_text(
        "⏳ *جاري إعداد CV الخاص بك...*", parse_mode='Markdown'
    )

    docx_path = None
    pdf_path  = None
    try:
        lang = context.user_data.get('lang', 'en')
        docx_path = generate_cv(context.user_data, lang=lang)
        name_clean = context.user_data['name'].replace(' ', '_')

        if fmt == 'pdf':
            pdf_path = convert_to_pdf(docx_path)
            with open(pdf_path, 'rb') as f:
                await update.callback_query.message.reply_document(
                    document=f,
                    filename=f"CV_{name_clean}.pdf",
                    caption="✅ هذا هو CV الخاص بك بصيغة PDF!"
                )
        else:
            with open(docx_path, 'rb') as f:
                await update.callback_query.message.reply_document(
                    document=f,
                    filename=f"CV_{name_clean}.docx",
                    caption="✅ هذا هو CV الخاص بك بصيغة Word!"
                )

        await update.callback_query.message.reply_text(
            "🌟 شكراً لاستخدام الخدمة!\n"
            "يمكنك إنشاء CV جديد بكتابة /start"
        )
        logger.info(f"CV sent successfully to user {user.id}")

    except FileNotFoundError as e:
        logger.error(f"LibreOffice not found for user {user.id}: {e}")
        await update.callback_query.message.reply_text(
            "⚠️ تعذّر تحويل الملف إلى PDF.\n"
            "سيتم إرسال الملف بصيغة Word بدلاً من ذلك."
        )
        if docx_path and os.path.exists(docx_path):
            name_clean = context.user_data['name'].replace(' ', '_')
            with open(docx_path, 'rb') as f:
                await update.callback_query.message.reply_document(
                    document=f,
                    filename=f"CV_{name_clean}.docx",
                    caption="✅ هذا هو CV الخاص بك بصيغة Word!"
                )
    except Exception as e:
        logger.error(f"Error generating CV for user {user.id}: {e}")
        await update.callback_query.message.reply_text(
            "❌ حدث خطأ أثناء إنشاء CV.\n"
            "يرجى المحاولة مجدداً بكتابة /start"
        )
    finally:
        # Always clean up temp files
        if docx_path and os.path.exists(docx_path):
            os.remove(docx_path)
        if pdf_path and os.path.exists(pdf_path):
            os.remove(pdf_path)

    return ConversationHandler.END


# ════════════════════════════════════════
#               MAIN
# ════════════════════════════════════════

def main():
    if not TOKEN:
        logger.critical("TELEGRAM_BOT_TOKEN is not set! Check your .env file.")
        raise SystemExit(1)

    # حذف ملف الـ pickle القديم لتجنب تعارض الحالات بعد التحديثات
    pickle_path = "bot_data/conversations.pickle"
    if os.path.exists(pickle_path):
        os.remove(pickle_path)
        logger.info("Old persistence file removed — starting fresh.")

    persistence = PicklePersistence(filepath=pickle_path)

    app = (
        Application.builder()
        .token(TOKEN)
        .persistence(persistence)
        .build()
    )

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler('start', start)],
        states={
            CV_LANG: [CallbackQueryHandler(get_cv_lang, pattern='^lang_(ar|en)$')],
            NAME:          [MessageHandler(filters.TEXT & ~filters.COMMAND, get_name)],
            JOB_TITLE:     [MessageHandler(filters.TEXT & ~filters.COMMAND, get_job_title)],
            EMAIL:         [MessageHandler(filters.TEXT & ~filters.COMMAND, get_email)],
            PHONE:         [MessageHandler(filters.TEXT & ~filters.COMMAND, get_phone)],
            LOCATION:      [MessageHandler(filters.TEXT & ~filters.COMMAND, get_location)],
            LINKEDIN: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, get_linkedin),
                CallbackQueryHandler(get_linkedin, pattern='^skip$')
            ],
            GITHUB: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, get_github),
                CallbackQueryHandler(get_github, pattern='^skip$')
            ],
            SUMMARY:       [MessageHandler(filters.TEXT & ~filters.COMMAND, get_summary)],
            EXP_TITLE:     [MessageHandler(filters.TEXT & ~filters.COMMAND, get_exp_title)],
            EXP_COMPANY:   [MessageHandler(filters.TEXT & ~filters.COMMAND, get_exp_company)],
            EXP_DATE:      [MessageHandler(filters.TEXT & ~filters.COMMAND, get_exp_date)],
            EXP_BULLETS:   [MessageHandler(filters.TEXT & ~filters.COMMAND, get_exp_bullets)],
            EXP_MORE:      [CallbackQueryHandler(get_exp_more, pattern='^(yes|no)$')],
            EDU_DEGREE:    [MessageHandler(filters.TEXT & ~filters.COMMAND, get_edu_degree)],
            EDU_UNIVERSITY:[MessageHandler(filters.TEXT & ~filters.COMMAND, get_edu_university)],
            EDU_DATE:      [MessageHandler(filters.TEXT & ~filters.COMMAND, get_edu_date)],
            SKILLS:        [MessageHandler(filters.TEXT & ~filters.COMMAND, get_skills)],
            CERT_ADD:      [CallbackQueryHandler(get_cert_add, pattern='^(yes|no)$')],
            CERT_NAME:     [MessageHandler(filters.TEXT & ~filters.COMMAND, get_cert_name)],
            CERT_ISSUER:   [MessageHandler(filters.TEXT & ~filters.COMMAND, get_cert_issuer)],
            CERT_DATE:     [MessageHandler(filters.TEXT & ~filters.COMMAND, get_cert_date)],
            CERT_MORE:     [CallbackQueryHandler(get_cert_more, pattern='^(yes|no)$')],
            LANGUAGES:     [MessageHandler(filters.TEXT & ~filters.COMMAND, get_languages)],
            EXPORT_FORMAT: [CallbackQueryHandler(get_export_format, pattern='^(word|pdf)$')],
        },
        fallbacks=[CommandHandler('cancel', cancel)],
        allow_reentry=True,
        persistent=True,
        name="cv_conversation"
    )

    app.add_handler(conv_handler)
    app.add_handler(CommandHandler('help', help_command))

    logger.info("✅ Bot is running...")
    app.run_polling(drop_pending_updates=True)


if __name__ == '__main__':
    main()
