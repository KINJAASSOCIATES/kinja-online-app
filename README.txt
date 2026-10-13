KINJA ASSOCIATES APP - ONLINE DEPLOYMENT FILES

Files:
- server.py
- index.html
- requirements.txt
- Procfile

Important:
1. Set environment variable APP_USER to your desired login username.
2. Set APP_PASSWORD to a strong, unique password before making the app public.
3. Set DATA_DIR to a persistent writable disk/folder. If the host uses temporary/ephemeral storage, app data may be lost after restart/redeploy.
4. This app stores client and payment records. Use HTTPS and persistent storage, keep backups, and do not deploy without authentication.
5. Start command: python server.py
6. Python dependency: openpyxl

For Render, connect this folder/repository as a Python web service, build command: pip install -r requirements.txt, start command: python server.py. Configure a persistent disk and set DATA_DIR to its mount path (for example /var/data) before entering real client data.
