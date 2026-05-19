# 🤖 CV Builder Telegram Bot

بوت تلغرام يُنشئ CV احترافياً عبر محادثة تفاعلية، مع تصدير Word أو PDF.

---

## 📁 هيكل المشروع

```
cv_bot/
├── bot.py              ← البوت الرئيسي ومنطق المحادثة
├── cv_generator.py     ← توليد DOCX وتحويله إلى PDF
├── requirements.txt    ← المكتبات المطلوبة
├── Dockerfile          ← صورة Docker للنشر
├── docker-compose.yml  ← للتشغيل المحلي بـ Docker
├── Procfile            ← للنشر على Railway/Render
├── railway.toml        ← إعدادات Railway
├── .env                ← متغيرات البيئة (لا ترفعها!)
└── .env.example        ← مثال على ملف .env
```

---

## ⚙️ الإعداد المحلي

### 1. إنشاء البوت
- افتح [@BotFather](https://t.me/BotFather) → `/newbot`
- احفظ الـ Token

### 2. إعداد البيئة
```bash
cp .env.example .env
# ثم ضع التوكن في ملف .env
```

### 3. تثبيت المتطلبات
```bash
pip install -r requirements.txt
```

### 4. تثبيت LibreOffice (لـ PDF)
- **Linux:** `sudo apt install libreoffice-writer -y`
- **Windows:** حمّل من https://www.libreoffice.org/download/
- **Docker:** مثبّت تلقائياً ✅

### 5. تشغيل البوت
```bash
python bot.py
```

---

## ☁️ النشر المجاني على Railway

### الطريقة الأسرع (5 دقائق):

**1. رفع الكود على GitHub:**
```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/USERNAME/cv_bot.git
git push -u origin main
```

**2. إنشاء مشروع على Railway:**
- اذهب إلى [railway.app](https://railway.app)
- سجّل دخول بحساب GitHub
- اضغط **"New Project"** → **"Deploy from GitHub repo"**
- اختر مستودع `cv_bot`

**3. إضافة متغير البيئة:**
- في لوحة Railway: **Variables** → **Add Variable**
- أضف: `TELEGRAM_BOT_TOKEN` = `توكنك هنا`

**4. النشر تلقائي ✅**
- Railway سيكتشف الـ `Dockerfile` ويبني الصورة تلقائياً
- البوت سيعمل خلال 2-3 دقائق

> [!NOTE]
> Railway يعطيك **$5 رصيد مجاني شهرياً** — كافٍ لتشغيل البوت 24/7.

---

## 🐳 تشغيل بـ Docker محلياً

```bash
docker-compose up -d
```

لمتابعة الـ logs:
```bash
docker-compose logs -f
```

لإيقاف البوت:
```bash
docker-compose down
```

---

## ⌨️ الأوامر

| الأمر | الوظيفة |
|-------|---------|
| `/start` | بدء إنشاء CV جديد |
| `/cancel` | إلغاء العملية الحالية |
| `/help` | عرض المساعدة |

---

## 🔄 سير المحادثة

```
/start → الاسم → الوظيفة → الإيميل → الهاتف → الموقع
→ LinkedIn (اختياري) → GitHub (اختياري) → الملخص المهني
→ الخبرات (متعددة) → التعليم → المهارات
→ الشهادات (اختيارية) → اللغات
→ معاينة البيانات → اختيار الصيغة (Word/PDF)
→ إرسال الملف ✅
```

---

## 📦 المكتبات

- `python-telegram-bot` — إدارة البوت
- `python-docx` — توليد Word
- `LibreOffice` — تحويل PDF (في Docker)
- `python-dotenv` — متغيرات البيئة
