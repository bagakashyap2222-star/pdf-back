# PDFiyo Conversion API Backend

High-fidelity document conversion microservice for PDFiyo (PDF to Word, Word to PDF, PDF to Excel, PPT, etc.).

Powered by:
- **FastAPI** (Python 3.11)
- **pdf2docx** (State-of-the-art layout & table reconstruction for Word)
- **LibreOffice Headless** (Enterprise-grade Office document to PDF engine)
- **pdfplumber & openpyxl** (Accurate spreadsheet grid extraction)

---

## 🚀 How to Push to your NEW GitHub Repository

Open PowerShell or Terminal inside this `pdfiyo-backend` folder:

```bash
# 1. Initialize git (if not already initialized)
git init

# 2. Add all files
git add .

# 3. Create initial commit
git commit -m "feat: initial commit for pdfiyo backend conversion engine"

# 4. Link to your new GitHub repository (replace with your repo URL)
git branch -M main
git remote add origin https://github.com/<YOUR-USERNAME>/pdfiyo-backend.git

# 5. Push code
git push -u origin main
```

---

## 🌐 How to Deploy to Render.com (100% Free Lifetime)

1. Go to [render.com](https://render.com) and Sign In (via GitHub).
2. Click **New +** ➔ Select **Web Service**.
3. Choose your newly created repository: **`pdfiyo-backend`**.
4. Configure these simple settings:
   - **Name:** `pdfiyo-backend` (or any name you prefer)
   - **Runtime:** `Docker` (Render auto-detects `Dockerfile`)
   - **Instance Type:** `Free`
5. Click **Deploy Web Service**!
6. Once deployed, Render gives you a free HTTPS live URL (e.g. `https://pdfiyo-backend.onrender.com`).

---

## 🔗 Connect to PDFiyo Frontend

In your PDFiyo Next.js project's `.env.local`:

```env
NEXT_PUBLIC_CONVERTER_API_URL=https://pdfiyo-backend.onrender.com
```

All conversions will now run with 100% precision on your own free backend!
