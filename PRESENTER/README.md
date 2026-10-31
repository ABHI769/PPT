# Smart PPT Presentation Controller

A modern web application built in **Python, HTML, CSS, JavaScript** featuring a sleek **gradient black aesthetic**. It allows presenters to upload PowerPoint (`.pptx` / `.ppt`) presentations, label pages with custom voice keywords (such as **"Introduction"**, **"Technology Used"**, **"Conclusion"**), and navigate directly to any slide using spoken voice commands or typed keywords.

---

## 🌟 Key Capabilities

1. **PowerPoint Upload & 100% Visual Fidelity Preservation Engine**:
   - Drag & Drop or click to upload `.pptx` and `.ppt` presentation files.
   - **Full Original Visual Preservation**: Automatically exports each slide to high-definition 1080p slide images, preserving 100% of the original layouts, diagrams, custom fonts, vector graphics, embedded photos, and color palettes without altering or discarding any content.
   - Full text extraction for real-time speech matching and speaker notes.
   - Built-in sample pitch deck preloaded with common voice labels.

2. **Custom Slide Labeling System**:
   - Assign multiple custom voice labels to each slide (e.g. Slide 1: *"Introduction"*, Slide 3: *"Technology Used"*, Slide 5: *"Conclusion"*).
   - Instant label addition with `Enter` key or click, and 1-click removal chips.
   - Persistent storage in JSON — labels are restored automatically across sessions.
   - **Pre-Presentation Review Table**: Inspect all slides and assigned voice labels before presenting.

3. **Voice-Controlled Slide Navigation (Web Speech API)**:
   - Natural language variations: *"Introduction"*, *"Show introduction"*, *"Open technology used"*, *"Go to conclusion"*, *"Next slide"*, *"Previous slide"*, *"Slide 4"*.
   - Typed fallback command bar in the Presenter HUD for noisy rooms or mic-less devices.
   - Ambiguous match protection: Displays an interactive choice selector if a label matches multiple slides.

4. **100% Isolated Audience Display (Highest Priority)**:
   - The audience screen displays **only** the PowerPoint slide on a pitch-black background.
   - **Zero Website UI**: No headers, footers, logos, navigation bars, microphone badges, slide numbers outside content, search bars, or overlays.
   - Aspect-ratio preservation (16:9 / 4:3) with Fullscreen API support.

5. **Multi-Window Synchronization**:
   - **Presenter HUD** (`/presenter`): Control cockpit with timer, speech transcript, next-slide preview, notes, and mic mute toggle.
   - **Audience View** (`/audience`): Clean display window meant for a projector or secondary monitor.
   - Real-time zero-latency sync across browser windows using `BroadcastChannel`.

6. **Gradient Black Design System**:
   - Deep obsidian and metallic black gradients, subtle glassmorphism, and readable typography (`Outfit`, `Plus Jakarta Sans`, and `JetBrains Mono`).

---

## 🚀 Running the Application (Python + HTML + CSS + JS)

### 1. Prerequisites
- Python 3.10+
- Dependencies in `requirements.txt`:
  ```bash
  pip install -r requirements.txt
  ```

### 2. Launch Server
```bash
python app.py
```

### 3. Open in Browser
- **Main Dashboard & Upload**: [http://localhost:5000/](http://localhost:5000/)
- **Presenter Command HUD**: [http://localhost:5000/presenter](http://localhost:5000/presenter)
- **Isolated Audience View**: [http://localhost:5000/audience](http://localhost:5000/audience)

---

## ☁️ Deploying to Vercel

The repository is pre-configured for one-click deployment on **Vercel**:

1. **Push to GitHub**:
   ```bash
   git init
   git add .
   git commit -m "Initial commit for Vercel deployment"
   git branch -M main
   git remote add origin https://github.com/<your-username>/<your-repo-name>.git
   git push -u origin main
   ```

2. **Deploy on Vercel**:
   - Go to [vercel.com](https://vercel.com/) and click **Add New -> Project**.
   - Import your GitHub repository.
   - Leave the Framework preset as **Other** (Vercel automatically detects the Python runtime from `vercel.json` and `api/index.py`).
   - Click **Deploy**.

---

## 🎤 Spoken Voice Commands

| Spoken Phrase | Action |
|---|---|
| `"Introduction"` / `"Show introduction"` | Navigates immediately to Slide 1 |
| `"Problem Statement"` / `"Challenges"` | Navigates immediately to Slide 2 |
| `"Technology Used"` / `"Tech Stack"` / `"Tools"` | Navigates immediately to Slide 3 |
| `"Architecture"` / `"System Design"` | Navigates immediately to Slide 4 |
| `"Conclusion"` / `"Final Summary"` | Navigates immediately to Slide 5 |
| `"Next slide"` / `"Next"` / `"Advance"` | Goes to next slide |
| `"Previous slide"` / `"Go back"` | Goes to previous slide |
| `"Slide 3"` / `"Page 2"` | Jumps to specific slide number |

---

## 📂 Project Architecture

```
PRESENTER/
├── api/
│   └── index.py                # Vercel Serverless Function entry point
├── app.py                      # Python Flask backend & PPTX parser
├── requirements.txt            # Python dependencies (cross-platform)
├── vercel.json                 # Vercel routing configuration
├── .vercelignore               # Ignored build artifacts for Vercel
├── templates/
│   ├── index.html              # Upload section, slide editor & review modal
│   ├── presenter.html          # Presenter HUD with voice recognition
│   └── audience.html           # 100% clean, isolated audience view
├── static/
│   ├── css/style.css           # Gradient Black modern design system
│   ├── js/main.js              # Upload handling, labeling & sync
│   └── slides/                 # Bundled & preloaded slide images
├── uploads/                    # Presentation upload directory
├── data/
│   └── presentations.json      # Persistent presentation & label storage
└── README.md
```
