# مترجم الملفات (إنجليزي → عربي)

## التشغيل
1. python -m venv venv && source venv/bin/activate   (ويندوز: venv\Scripts\activate)
2. pip install -r requirements.txt
3. cp .env.example .env  ثم عدّل SECRET_KEY (أول تشغيل يحتاج إنترنت لتنزيل نموذج الترجمة مرة واحدة ~100MB)
4. uvicorn app:app --host 0.0.0.0 --port 8000
5. افتح http://localhost:8000 — أول حساب تسجله يصبح المشرف تلقائياً.

## الإنتاج
- ضع الموقع خلف Nginx + شهادة HTTPS (Let's Encrypt) ثم فعّل https_only=True في SessionMiddleware.
- شغّل الخدمة عبر systemd أو Docker، وخذ نسخاً احتياطية من مجلد data/.
