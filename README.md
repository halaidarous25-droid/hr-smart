# نظام توزيع المسؤوليات الذكي — Python Flask

## ⚡ التشغيل في 3 خطوات

```bash
# 1. تثبيت المكتبات (مرة واحدة فقط)
pip install flask anthropic openpyxl

# 2. مفتاح Anthropic API
export ANTHROPIC_API_KEY="sk-ant-..."    # Linux/Mac
set ANTHROPIC_API_KEY=sk-ant-...         # Windows

# 3. تشغيل الخادم
cd hr_app
python app.py

# 4. افتح المتصفح على
http://localhost:5000
```

## لماذا Flask وليس HTML مباشرة؟

الملف HTML وحده يرفض الاتصال بـ Anthropic API بسبب **CORS**:
- المتصفح يمنع الاتصال بـ `api.anthropic.com` من ملف محلي
- Flask يحل هذا: الـ HTML يتصل بـ `/api/ai/distribute` (نفس الخادم)، وFlask يتصل بـ Anthropic من Python (لا CORS)

## البنية
```
hr_app/
├── app.py              ← Flask + كل منطق AI (Python)
├── data.json           ← قاعدة البيانات (تُنشأ تلقائياً)
├── templates/
│   └── index.html      ← الواجهة الكاملة
└── README.md
```

## API Endpoints

| Endpoint | الوظيفة |
|----------|---------|
| `POST /api/ai/distribute` | التوزيع الذكي ← Anthropic API |
| `POST /api/ai/analyze-employee` | تحليل موظف ← Anthropic API |
| `POST /api/ai/analyze-responsibility` | تحليل مسؤولية ← Anthropic API |
| `GET/POST /api/employees` | CRUD الموظفين |
| `GET/POST /api/responsibilities` | CRUD المسؤوليات |
| `POST /api/import/employees` | استيراد Excel |
| `POST /api/import/responsibilities` | استيراد Excel |
| `GET /api/template/employees` | تحميل قالب Excel |
| `GET /api/template/responsibilities` | تحميل قالب Excel |
| `GET /api/export/report` | تصدير تقرير Excel |
