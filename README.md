# 📄 Resume Reformat Studio (Next.js + Gemini Flash + Vercel)

A lightweight, production-ready Next.js web application built for HR teams to batch-convert **20–30 candidate resumes** at a time into a standardized corporate Word document (`.docx`) template using **Google Gemini 2.0 / 1.5 Flash (100% Free Tier)**.

---

## 🚀 Key Features

- **Batch Upload**: Select and convert 20–30 candidate resumes (`.pdf`, `.docx`, `.txt`) simultaneously.
- **Strict Factual Accuracy**: AI is prompted with strict guardrails: **zero hallucinations**—it restructures and formats candidate facts without inventing skills, employers, or dates.
- **100% Free AI Tier**: Uses Google Gemini 2.0 / 1.5 Flash via Google AI Studio (**1,500 free requests per day**; 20–30 daily files uses only ~2% of the free quota).
- **Executive Word Output (`.docx`)**: Generates styled Word documents adhering to the target template's sections, typography, contact layout, and bullet structure.
- **Client-Side ZIP Generation**: HR can download all converted resumes in a single `.zip` file with one click, or download individual `.docx` files.
- **Vercel Optimized**: Client-orchestrated batch queue prevents serverless timeouts. Compatible with Vercel's free Hobby plan.

---

## 🛠️ Getting Started (Local Development)

### 1. Clone & Install
```bash
git clone <your-repo-url>
cd ResmueConverter
npm install
```

### 2. Configure Your Free Gemini API Key
Get a free API key from [Google AI Studio](https://aistudio.google.com/app/apikey).

Create a `.env.local` file in the project root:
```env
GEMINI_API_KEY=your_gemini_api_key_here
```
*(Alternatively, HR users can paste the key directly into the web application's top navigation bar).*

### 3. Run the Development Server
```bash
npm run dev
```

Open [http://localhost:3000](http://localhost:3000) in your browser.

---

## ☁️ Deploying to Vercel

### Method 1: Push to GitHub & Import in Vercel
1. Push this repository to your GitHub/GitLab account.
2. Go to [Vercel Dashboard](https://vercel.com/new).
3. Import this repository.
4. Under **Environment Variables**, add:
   - `GEMINI_API_KEY` = `<your-free-gemini-key>`
5. Click **Deploy**.

### Method 2: Deploy using Vercel CLI
```bash
npx vercel
```

---

## 📋 How HR Uses the Application

1. **Step 1: Target Template**
   - Upload your company's benchmark resume (`.docx`, `.pdf`, or `.txt`), or use the built-in corporate template.
2. **Step 2: Upload Resumes**
   - Click the candidate upload box and select 20–30 resumes (`.pdf` or `.docx`).
3. **Step 3: Convert**
   - Click **"Convert Resumes"**. Watch the real-time progress bar.
4. **Step 4: Download**
   - Click **"Download Converted Resumes (.ZIP)"** to save all formatted `.docx` files at once.

---

## 🔒 HR Safeguards & Data Integrity

- The transformation engine maps source facts into standardized sections (Summary, Skills, Experience, Education, Certifications, Projects).
- It preserves source companies, job titles, dates of employment, and key achievements.
- Rate-limiting intervals prevent hitting Google's Free Tier RPM thresholds.
