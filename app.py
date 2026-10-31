import os
import json
import uuid
import shutil
import textwrap
import tempfile
from flask import Flask, request, jsonify, send_file
from pptx import Presentation

try:
    import pythoncom
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IS_VERCEL = bool(os.environ.get('VERCEL') or os.environ.get('AWS_LAMBDA_FUNCTION_NAME'))

app = Flask(
    __name__,
    template_folder=BASE_DIR,
    static_folder=BASE_DIR,
    static_url_path=''
)

# Use system temp directory for uploads and dynamic slide generation
TEMP_BASE = tempfile.gettempdir()
app.config['UPLOAD_FOLDER'] = os.path.join(TEMP_BASE, 'uploads')
app.config['SLIDES_FOLDER'] = os.path.join(TEMP_BASE, 'slides')

# Presentations metadata file: in root directory locally, or in temp on Vercel
ROOT_DATA_FILE = os.path.join(BASE_DIR, 'presentations.json')
app.config['DATA_FILE'] = os.path.join(TEMP_BASE, 'presentations.json') if IS_VERCEL else ROOT_DATA_FILE

for folder in [app.config['UPLOAD_FOLDER'], app.config['SLIDES_FOLDER']]:
    try:
        os.makedirs(folder, exist_ok=True)
    except Exception:
        pass

def load_data():
    target = app.config['DATA_FILE']
    if not os.path.exists(target) and os.path.exists(ROOT_DATA_FILE):
        try:
            with open(ROOT_DATA_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            save_data(data)
            return data
        except Exception:
            pass

    if os.path.exists(target):
        try:
            with open(target, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}

    if os.path.exists(ROOT_DATA_FILE):
        try:
            with open(ROOT_DATA_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}

    return {}

def save_data(data):
    try:
        target_dir = os.path.dirname(app.config['DATA_FILE'])
        if target_dir:
            os.makedirs(target_dir, exist_ok=True)
        with open(app.config['DATA_FILE'], 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"Error saving data: {e}")

def render_slide_fallback_image(slide_num, title, points, notes, out_file):
    """
    Renders a high-definition 1080p slide image in the gradient black aesthetic
    as a resilient fallback when PowerPoint COM export is not available.
    """
    try:
        from PIL import Image, ImageDraw, ImageFont

        width, height = 1920, 1080
        img = Image.new('RGB', (width, height), color=(8, 3, 16))
        draw = ImageDraw.Draw(img)

        # Outer card frame with subtle purple border
        draw.rounded_rectangle([50, 50, width - 50, height - 50], radius=24, outline=(76, 29, 149), width=2)

        # Attempt to load system fonts
        font_badge = font_title = font_body = font_notes = None
        for font_name in [
            'DejaVuSans-Bold.ttf', 'LiberationSans-Bold.ttf', 'FreeSansBold.ttf',
            'segoeuib.ttf', 'arialbd.ttf', 'calibrib.ttf'
        ]:
            try:
                font_badge = ImageFont.truetype(font_name, 22)
                font_title = ImageFont.truetype(font_name, 54)
                break
            except Exception:
                continue

        for font_name in [
            'DejaVuSans.ttf', 'LiberationSans-Regular.ttf', 'FreeSans.ttf',
            'segoeui.ttf', 'arial.ttf', 'calibri.ttf'
        ]:
            try:
                font_body = ImageFont.truetype(font_name, 32)
                font_notes = ImageFont.truetype(font_name, 22)
                break
            except Exception:
                continue

        if not font_title:
            font_badge = font_title = font_body = font_notes = ImageFont.load_default()

        # Slide Number Badge in vibrant violet
        badge_text = f"SLIDE {slide_num:02d}"
        draw.rounded_rectangle([120, 110, 275, 160], radius=16, fill=(124, 58, 237), outline=(168, 85, 247), width=1)
        draw.text((142, 122), badge_text, fill=(255, 255, 255), font=font_badge)

        # Slide Title
        wrapped_title = textwrap.fill(title or f"Slide {slide_num}", width=50)
        draw.text((120, 195), wrapped_title, fill=(248, 250, 252), font=font_title)

        # Divider line
        title_lines = wrapped_title.count('\n') + 1
        divider_y = 195 + (title_lines * 65) + 20
        draw.line([(120, divider_y), (width - 120, divider_y)], fill=(50, 20, 90), width=2)

        # Bullet Points with violet bullet markers
        y = divider_y + 40
        if points:
            for pt in points[:7]:
                clean_pt = pt.strip()
                if not clean_pt:
                    continue
                wrapped_pt = textwrap.fill(clean_pt, width=70)
                draw.text((120, y), "•", fill=(192, 132, 252), font=font_body)
                draw.text((160, y), wrapped_pt, fill=(226, 232, 240), font=font_body)
                line_count = wrapped_pt.count('\n') + 1
                y += (line_count * 46) + 16
                if y > 920:
                    break
        else:
            draw.text((120, y), "• Presentation slide content", fill=(196, 181, 253), font=font_body)

        # Speaker notes hint at bottom
        if notes:
            clean_notes = textwrap.fill(f"Speaker Note: {notes}", width=95).split('\n')[0]
            draw.text((120, height - 90), clean_notes, fill=(100, 116, 139), font=font_notes)

        os.makedirs(os.path.dirname(out_file), exist_ok=True)
        img.save(out_file, 'PNG')
        return True
    except Exception as e:
        print(f"Fallback slide image rendering error: {e}")
        return False

def ensure_sample_slide(slide_num, out_file):
    sample_info = {
        1: ("Smart PPT Presentation Controller", ["Voice-controlled slide navigation with customizable tags.", "100% clean presentation mode for audience display."], "Introduce the application and how voice navigation streamlines Q&A."),
        2: ("Problem Statement & Challenges", ["Overcoming Presentation Friction.", "Presenters waste valuable time searching through slides during live Q&A."], "Discuss how speakers struggle when flipping through 40+ slides during Q&A."),
        3: ("Technologies Used & Architecture", ["Backend: Python (Flask) with Native PowerPoint Engine.", "Frontend: Gradient Black Vanilla CSS & Modern JavaScript.", "Web Speech API with zero-latency channel sync."], "Explain Python backend with python-pptx and Web Speech API in frontend."),
        4: ("Presentation Mode & Audience View", ["Audience screen displays ONLY the slide on black background.", "Zero website controls, browser chrome, headers, or watermarks.", "Preserves 100% of shapes, fonts, and graphics."], "Highlight that the audience display is pure and contains zero website elements."),
        5: ("Conclusion & Future Scope", ["The Future of AI-Powered Presentations.", "Say 'Technology Used' to test instant voice jump!"], "Conclude and invite questions to test voice jumps.")
    }
    info = sample_info.get(slide_num, (f"Slide {slide_num}", ["Presentation slide content"], ""))
    render_slide_fallback_image(slide_num, info[0], info[1], info[2], out_file)

def export_slides_to_images(filepath, pres_id):
    """
    Exports each slide of the PowerPoint file into high-definition 1080p PNG images
    using native PowerPoint COM engine if available. Safely closes all COM objects.
    """
    if not HAS_WIN32COM:
        return []

    out_dir = os.path.join(app.config['SLIDES_FOLDER'], pres_id)
    os.makedirs(out_dir, exist_ok=True)
    abs_filepath = os.path.abspath(filepath)
    abs_out_dir = os.path.abspath(out_dir)

    exported_images = []
    ppt_app = None
    pres = None
    co_initialized = False

    try:
        pythoncom.CoInitialize()
        co_initialized = True
        ppt_app = win32com.client.Dispatch('PowerPoint.Application')
        pres = ppt_app.Presentations.Open(abs_filepath, ReadOnly=1, Untitled=0, WithWindow=0)
        slide_count = pres.Slides.Count
        for i in range(1, slide_count + 1):
            img_file = os.path.join(abs_out_dir, f'slide_{i}.png')
            pres.Slides(i).Export(img_file, 'PNG', 1920, 1080)
            exported_images.append(f'/static/slides/{pres_id}/slide_{i}.png')
    except Exception as e:
        print(f"PowerPoint COM export notice: {e}")
    finally:
        if pres:
            try:
                pres.Close()
            except Exception:
                pass
        if ppt_app:
            try:
                if ppt_app.Presentations.Count == 0:
                    ppt_app.Quit()
            except Exception:
                try:
                    ppt_app.Quit()
                except Exception:
                    pass
        if co_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    return exported_images

def parse_pptx_file(filepath, pres_id):
    """
    Extracts slide structure, text, speaker notes, and ensures high-definition slide images.
    Falls back to high-res Pillow rendering if COM export is unavailable.
    """
    exported_images = export_slides_to_images(filepath, pres_id)

    prs = Presentation(filepath)
    slides_data = []

    out_dir = os.path.join(app.config['SLIDES_FOLDER'], pres_id)
    os.makedirs(out_dir, exist_ok=True)

    for idx, slide in enumerate(prs.slides):
        slide_num = idx + 1
        title = ""
        notes = ""
        text_snippets = []

        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            notes = slide.notes_slide.notes_text_frame.text.strip()

        for shape in slide.shapes:
            if shape.has_text_frame:
                txt = shape.text_frame.text.strip()
                if txt:
                    text_snippets.append(txt)
                    if not title and (shape == getattr(slide.shapes, 'title', None) or len(txt) < 90):
                        title = txt.split('\n')[0].strip()

        if not title:
            title = f"Slide {slide_num}"

        raw_text = " ".join(text_snippets)
        points = [snippet for snippet in text_snippets if snippet != title]

        initial_labels = []
        lower_title = title.lower()
        if idx == 0 or 'intro' in lower_title or 'overview' in lower_title:
            initial_labels.append("Introduction")
        if 'problem' in lower_title or 'challenge' in lower_title:
            initial_labels.append("Problem Statement")
        if any(w in lower_title for w in ['tech', 'hardware', 'software', 'stack', 'architecture', 'tool']):
            initial_labels.append("Technology Used")
        if any(w in lower_title for w in ['work', 'implementation', 'method', 'flow']):
            initial_labels.append("Implementation")
        if any(w in lower_title for w in ['result', 'application', 'feature', 'demo', 'metric']):
            initial_labels.append("Results")
        if any(w in lower_title for w in ['conclusion', 'future', 'summary', 'end', 'thank']):
            initial_labels.append("Conclusion")

        img_rel_path = f"/static/slides/{pres_id}/slide_{slide_num}.png"
        img_abs_path = os.path.join(out_dir, f"slide_{slide_num}.png")

        if idx < len(exported_images) and os.path.exists(img_abs_path):
            image_url = exported_images[idx]
        elif not os.path.exists(img_abs_path):
            render_slide_fallback_image(slide_num, title, points, notes, img_abs_path)
            image_url = img_rel_path
        else:
            image_url = img_rel_path

        slides_data.append({
            "number": slide_num,
            "index": idx,
            "title": title,
            "image_url": image_url,
            "notes": notes,
            "raw_text": raw_text,
            "points": points,
            "labels": initial_labels
        })

    return slides_data

# HTML Pages (served directly from the root directory)
@app.route('/')
def index():
    return send_file(os.path.join(BASE_DIR, 'index.html'))

@app.route('/presenter')
def presenter_view():
    return send_file(os.path.join(BASE_DIR, 'presenter.html'))

@app.route('/audience')
def audience_view():
    return send_file(os.path.join(BASE_DIR, 'audience.html'))

# Static Assets (served directly from root)
@app.route('/style.css')
@app.route('/static/css/style.css')
def serve_css():
    return send_file(os.path.join(BASE_DIR, 'style.css'), mimetype='text/css')

@app.route('/main.js')
@app.route('/static/js/main.js')
def serve_js():
    return send_file(os.path.join(BASE_DIR, 'main.js'), mimetype='application/javascript')

# Slide Image Serving
@app.route('/static/slides/<path:filename>')
@app.route('/slides/<path:filename>')
def serve_slides(filename):
    out_path = os.path.join(app.config['SLIDES_FOLDER'], filename)
    if os.path.isfile(out_path):
        return send_file(out_path, mimetype='image/png')

    # If it's a sample deck slide, generate it on demand
    if 'sample_deck/slide_' in filename or 'sample_deck\\slide_' in filename:
        try:
            slide_str = filename.replace('\\', '/').split('slide_')[1].split('.')[0]
            slide_num = int(slide_str)
            os.makedirs(os.path.dirname(out_path), exist_ok=True)
            ensure_sample_slide(slide_num, out_path)
            if os.path.isfile(out_path):
                return send_file(out_path, mimetype='image/png')
        except Exception as e:
            print(f"Error generating sample slide: {e}")

    return jsonify({"error": "Slide image not found"}), 404

# API Endpoints
@app.route('/api/presentations', methods=['GET'])
def get_presentations():
    data = load_data()
    return jsonify(list(data.values()))

@app.route('/api/presentation/<pres_id>', methods=['GET'])
def get_presentation(pres_id):
    data = load_data()
    if pres_id in data:
        return jsonify(data[pres_id])
    return jsonify({"error": "Presentation not found"}), 404

@app.route('/api/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']
    if not file.filename:
        return jsonify({"error": "Empty filename"}), 400

    lower_name = file.filename.lower()
    if not (lower_name.endswith('.pptx') or lower_name.endswith('.ppt')):
        return jsonify({"error": "Only .pptx or .ppt files supported"}), 400

    pres_id = f"pres_{uuid.uuid4().hex[:8]}"
    saved_filename = f"{pres_id}_{file.filename}"
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], saved_filename)
    file.save(filepath)

    parse_path = filepath
    if lower_name.endswith('.ppt') and not lower_name.endswith('.pptx'):
        if HAS_WIN32COM:
            try:
                pythoncom.CoInitialize()
                ppt_app = win32com.client.Dispatch('PowerPoint.Application')
                abs_ppt = os.path.abspath(filepath)
                pres = ppt_app.Presentations.Open(abs_ppt, ReadOnly=1, Untitled=0, WithWindow=0)
                pptx_path = filepath + 'x'
                pres.SaveCopyAs(os.path.abspath(pptx_path), 24)
                pres.Close()
                ppt_app.Quit()
                pythoncom.CoUninitialize()
                parse_path = pptx_path
            except Exception as e:
                return jsonify({"error": f"Binary .ppt conversion failed: {str(e)}. Please save as .pptx and upload."}), 400
        else:
            return jsonify({"error": "Binary .ppt format requires PowerPoint. Please save file as .pptx and upload."}), 400

    try:
        slides = parse_pptx_file(parse_path, pres_id)
    except Exception as e:
        return jsonify({"error": f"Failed to parse presentation: {str(e)}"}), 500

    pres_info = {
        "id": pres_id,
        "title": file.filename.rsplit('.', 1)[0].replace('_', ' '),
        "filename": file.filename,
        "slide_count": len(slides),
        "slides": slides
    }

    data = load_data()
    data[pres_id] = pres_info
    save_data(data)

    return jsonify(pres_info)

@app.route('/api/presentation/<pres_id>/labels', methods=['POST'])
def update_labels(pres_id):
    payload = request.get_json() or {}
    slide_index = payload.get('slide_index')
    labels = payload.get('labels', [])

    data = load_data()
    if pres_id not in data:
        return jsonify({"error": "Presentation not found"}), 404

    slides = data[pres_id].get('slides', [])
    if isinstance(slide_index, int) and 0 <= slide_index < len(slides):
        slides[slide_index]['labels'] = [l.strip() for l in labels if l.strip()]
        data[pres_id]['slides'] = slides
        save_data(data)
        return jsonify({"success": True, "slides": slides})

    return jsonify({"error": "Slide index out of range"}), 400

@app.route('/api/presentation/<pres_id>', methods=['DELETE'])
def delete_presentation(pres_id):
    data = load_data()
    if pres_id in data:
        slides_dir = os.path.join(app.config['SLIDES_FOLDER'], pres_id)
        if os.path.exists(slides_dir):
            try:
                shutil.rmtree(slides_dir)
            except Exception:
                pass

        filename = data[pres_id].get('filename')
        if filename:
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], f"{pres_id}_{filename}")
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except Exception:
                    pass

        del data[pres_id]
        save_data(data)
        return jsonify({"success": True})
    return jsonify({"error": "Not found"}), 404

if __name__ == '__main__':
    print("Starting Smart PPT Controller on http://localhost:5000 ...")
    app.run(host='0.0.0.0', port=5000, debug=False)
