"""
SmartSlide AI - AI PPT Generator (single file: backend + frontend)

Run:   uvicorn app:app --reload --port 8000
Open:  http://127.0.0.1:8000
"""
import base64
import datetime
import importlib.util
import io
import json
import os
import re
import subprocess
import sys
import urllib.parse
import urllib.request

# Install any missing package automatically
for _mod, _pkg in (
    ("fastapi", "fastapi"),
    ("uvicorn", "uvicorn[standard]"),
    ("pptx", "python-pptx"),
    ("requests", "requests"),
    ("PIL", "Pillow"),
    ("bs4", "beautifulsoup4"),
    ("youtube_transcript_api", "youtube-transcript-api"),
    ("pypdf", "pypdf"),
    ("docx", "python-docx"),
):
    if importlib.util.find_spec(_mod) is None:
        try:
            print(f"Installing {_pkg} ...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", _pkg])
        except Exception as e:
            print(f"Auto-install warning ({_pkg}): {e}")

import fastapi  # noqa: E402
import requests  # noqa: E402
import uvicorn  # noqa: E402
from fastapi import Body, FastAPI, HTTPException, UploadFile, File, Form  # noqa: E402
from fastapi.responses import HTMLResponse, Response  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.dml.color import RGBColor  # noqa: E402
from pptx.enum.text import PP_ALIGN  # noqa: E402
from pptx.enum.shapes import MSO_SHAPE  # noqa: E402
from pptx.util import Inches, Pt  # noqa: E402

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "settings.json")
PROVIDERS = ("gemini", "groq", "openai", "deepseek", "openrouter", "anthropic", "grok", "custom", "ollama")

CFG = {
    "provider": os.getenv("SMARTSLIDE_PROVIDER", "openai"),
    "base_url": os.getenv("SMARTSLIDE_BASE_URL", "https://api.groq.com/openai/v1"),
    "model": os.getenv("SMARTSLIDE_MODEL", "openai/gpt-oss-120b"),
    "api_key": os.getenv("SMARTSLIDE_API_KEY", ""),
}
if os.path.exists(CONFIG_FILE):
    try:
        with open(CONFIG_FILE, encoding="utf-8") as _f:
            CFG.update({k: v for k, v in json.load(_f).items() if k in CFG})
    except Exception:
        pass

app = fastapi.FastAPI(title="SmartSlide AI")

# ----------------------------------------------------------------------------
# AI helpers
# ----------------------------------------------------------------------------

def parse_json(text: str):
    text = str(text or "").strip()
    text = re.sub(r"<think>[\s\S]*?</think>", "", text, flags=re.IGNORECASE).strip()
    try:
        return json.loads(text)
    except Exception:
        pass
    fence = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", text, re.IGNORECASE)
    if fence:
        try:
            return json.loads(fence.group(1).strip())
        except Exception:
            text = fence.group(1).strip()
    m = re.search(r"([\[{][\s\S]*[\]}])", text)
    if m:
        candidate = m.group(1).strip()
        try:
            return json.loads(candidate)
        except Exception:
            candidate_clean = re.sub(r",\s*([\]}])", r"\1", candidate)
            try:
                return json.loads(candidate_clean)
            except Exception:
                pass
    raise HTTPException(502, "Model did not return valid JSON. Please try again.")


def http_error(r, provider: str, model: str = ""):
    if r.status_code in (401, 403):
        raise HTTPException(502, f"{provider}: API key is invalid or unauthorized (HTTP {r.status_code}).")
    if r.status_code == 404:
        raise HTTPException(502, f"{provider}: Model name '{model}' or Base URL is incorrect (HTTP 404).")
    if r.status_code == 429:
        raise HTTPException(502, f"{provider}: Rate limit or quota exhausted.")
    raise HTTPException(502, f"{provider} error {r.status_code}: {r.text[:200]}")


def ask(prompt: str, ai_config: dict = None):
    cfg = dict(CFG)
    if ai_config and isinstance(ai_config, dict):
        for k in ("provider", "base_url", "model", "api_key"):
            if ai_config.get(k):
                cfg[k] = str(ai_config[k]).strip()

    p, key, model = cfg["provider"], cfg["api_key"], cfg["model"]
    base = (cfg.get("base_url") or "").rstrip("/")

    if p != "ollama" and not key:
        raise HTTPException(400, "API key is not configured. Please enter your API key in AI Settings.")
    if not model:
        raise HTTPException(400, "Model name is not selected. Please select a model in AI Settings.")

    try:
        if p == "ollama":
            r = requests.post(
                f"{base or 'http://localhost:11434'}/api/generate",
                json={"model": model, "prompt": prompt, "stream": False, "format": "json",
                      "options": {"temperature": 0.6}},
                timeout=120,
            )
            if not r.ok:
                http_error(r, "Ollama", model)
            text = r.json().get("response", "")
        elif p == "gemini":
            r = requests.post(
                f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={key}",
                json={"contents": [{"parts": [{"text": prompt}]}],
                      "generationConfig": {"responseMimeType": "application/json", "temperature": 0.6}},
                timeout=120,
            )
            if not r.ok:
                http_error(r, "Gemini", model)
            text = r.json()["candidates"][0]["content"]["parts"][0]["text"]
        elif p == "anthropic":
            r = requests.post(
                (base or "https://api.anthropic.com/v1") + "/messages",
                headers={"x-api-key": key, "anthropic-version": "2023-06-01", "Content-Type": "application/json"},
                json={"model": model, "max_tokens": 2048, "temperature": 0.6,
                      "messages": [{"role": "user", "content": prompt + "\n\nReply with valid JSON only."}]},
                timeout=120,
            )
            if not r.ok:
                http_error(r, "Anthropic", model)
            text = r.json()["content"][0]["text"]
        else:
            url = (base or "https://api.openai.com/v1") + "/chat/completions"
            headers = {"Authorization": f"Bearer {key}", "Content-Type": "application/json"}
            body = {"model": model, "temperature": 0.6, "messages": [{"role": "user", "content": prompt}],
                    "response_format": {"type": "json_object"}}
            r = requests.post(url, headers=headers, json=body, timeout=120)
            if r.status_code == 400:
                body.pop("response_format", None)
                r = requests.post(url, headers=headers, json=body, timeout=120)
            if not r.ok:
                http_error(r, p.upper(), model)
            text = r.json()["choices"][0]["message"]["content"]
    except HTTPException:
        raise
    except requests.exceptions.ConnectionError:
        raise HTTPException(502, f"Connection failed to {p} service. Check network or Base URL.")
    except (KeyError, IndexError, ValueError, TypeError):
        raise HTTPException(502, "Failed to parse API response. Please verify model name and provider.")
    except requests.RequestException as e:
        raise HTTPException(502, f"Network error: {e}")
    return parse_json(text)

@app.get("/api/config")
def get_config():
    return {"provider": CFG["provider"], "base_url": CFG["base_url"], "model": CFG["model"],
            "has_key": bool(CFG["api_key"])}


@app.post("/api/config")
def set_config(p: dict = Body(...)):
    provider = str(p.get("provider", "ollama")).lower()
    if provider not in PROVIDERS:
        raise HTTPException(400, "Unknown provider")
    CFG["provider"] = provider
    CFG["base_url"] = str(p.get("base_url", "")).strip()
    CFG["model"] = str(p.get("model", "")).strip()
    if p.get("clear_key"):
        CFG["api_key"] = ""
    elif str(p.get("api_key", "")).strip():
        CFG["api_key"] = str(p["api_key"]).strip()
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(CFG, f, indent=2)
        try:
            os.chmod(CONFIG_FILE, 0o600)
        except OSError:
            pass
    except OSError as e:
        raise HTTPException(500, f"settings.json could not be saved: {e}")
    return get_config()


@app.post("/api/config/test")
def test_config():
    ask('Return JSON exactly like {"ok": true}')
    return {"ok": True}


def ctx(s: dict) -> str:
    src_text = s.get("source_material", "")
    src_part = f"\nSource context to base presentation on:\n{src_text[:3500]}\n" if src_text else ""
    return (
        f"Topic: {s.get('topic', '')}. Language: {s.get('language', 'English')}. "
        f"Presentation type: {s.get('type', 'Educational')}. "
        f"Audience: {s.get('audience', 'General Audience')}. "
        f"Difficulty: {s.get('difficulty', 'Beginner')}. "
        f"Duration: {s.get('duration', '10 minutes')}. "
        f"Extra instructions: {s.get('instructions', '') or 'none'}."
        f"{src_part}"
    )


def as_list(v):
    if isinstance(v, list):
        return [str(x.get("title", x) if isinstance(x, dict) else x).strip() for x in v if x]
    return []


@app.post("/api/design")
def generate_design(p: dict = Body(...)):
    topic = str(p.get("topic", "")).strip() or "General Presentation"
    prompt = (
        f"You are a world-class presentation art director (like Gamma AI / Pitch).\n"
        f"Analyze this topic: \"{topic}\" and choose an aesthetic, topic-tailored presentation theme.\n"
        f"Requirements:\n"
        f"1. Choose a background style: light pastel gradient or modern soft backdrop.\n"
        f"2. Ensure very high contrast so text is effortlessly readable.\n"
        f"3. Select a decorative watermark theme: 'space' (stars/planets), 'nature' (leaves/vines), "
        f"'tech' (nodes/circuits), 'business' (geometric grid), 'health' (organic curves), or 'abstract' (soft shapes).\n"
        f"4. Provide 4 preset color palettes.\n"
        f"Return ONLY JSON formatted as:\n"
        f"{{\n"
        f'  "mood": "tech",\n'
        f'  "bg_start": "#F8FAFC",\n'
        f'  "bg_end": "#EEF2F6",\n'
        f'  "card_bg": "rgba(255,255,255,0.85)",\n'
        f'  "accent": "#4F46E5",\n'
        f'  "accent_light": "#EEF2FF",\n'
        f'  "text_primary": "#0F172A",\n'
        f'  "text_secondary": "#475569",\n'
        f'  "gradient_deg": 135,\n'
        f'  "decorative": "tech",\n'
        f'  "font_title": "Inter, system-ui, sans-serif",\n'
        f'  "font_body": "Inter, system-ui, sans-serif",\n'
        f'  "presets": [\n'
        f'    {{"name": "Topic Tailored", "bg_start": "#F8FAFC", "bg_end": "#EEF2FF", "accent": "#4F46E5", "card_bg": "rgba(255,255,255,0.9)", "text": "#0F172A"}},\n'
        f'    {{"name": "Warm Paper", "bg_start": "#FFFDF9", "bg_end": "#F7EFE5", "accent": "#D97706", "card_bg": "rgba(255,255,255,0.95)", "text": "#292524"}},\n'
        f'    {{"name": "Calm Mint", "bg_start": "#F0FDF4", "bg_end": "#DCFCE7", "accent": "#059669", "card_bg": "rgba(255,255,255,0.9)", "text": "#064E3B"}},\n'
        f'    {{"name": "Midnight Glow", "bg_start": "#0F172A", "bg_end": "#1E1B4B", "accent": "#818CF8", "card_bg": "rgba(30,41,59,0.8)", "text": "#F8FAFC"}}\n'
        f'  ]\n'
        f"}}"
    )
    try:
        res = ask(prompt)
        return res
    except Exception:
        return {
            "mood": "tech",
            "bg_start": "#F8FAFC",
            "bg_end": "#EEF2FF",
            "card_bg": "rgba(255,255,255,0.88)",
            "accent": "#4F46E5",
            "accent_light": "#EEF2FF",
            "text_primary": "#0F172A",
            "text_secondary": "#475569",
            "gradient_deg": 135,
            "decorative": "tech",
            "font_title": "Inter, system-ui, sans-serif",
            "font_body": "Inter, system-ui, sans-serif",
            "presets": [
                {"name": "Modern Slate", "bg_start": "#F8FAFC", "bg_end": "#EEF2FF", "accent": "#4F46E5", "card_bg": "rgba(255,255,255,0.9)", "text": "#0F172A"},
                {"name": "Warm Paper", "bg_start": "#FFFDF9", "bg_end": "#F7EFE5", "accent": "#D97706", "card_bg": "rgba(255,255,255,0.95)", "text": "#292524"},
                {"name": "Calm Mint", "bg_start": "#F0FDF4", "bg_end": "#DCFCE7", "accent": "#059669", "card_bg": "rgba(255,255,255,0.9)", "text": "#064E3B"},
                {"name": "Midnight Glow", "bg_start": "#0F172A", "bg_end": "#1E1B4B", "accent": "#818CF8", "card_bg": "rgba(30,41,59,0.8)", "text": "#F8FAFC"}
            ]
        }

@app.post("/api/extract_source")
async def extract_source(
    url: str = Form(None),
    file: UploadFile = File(None)
):
    extracted_text = ""
    source_name = ""

    # 1. Process uploaded file
    if file and file.filename:
        filename = file.filename.lower()
        source_name = file.filename
        content = await file.read()
        try:
            if filename.endswith(".pdf"):
                import pypdf
                reader = pypdf.PdfReader(io.BytesIO(content))
                pages = [p.extract_text() or "" for p in reader.pages]
                extracted_text = "\n\n".join(pages)
            elif filename.endswith(".docx"):
                import docx
                doc = docx.Document(io.BytesIO(content))
                extracted_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
            else:  # plain text or markdown
                extracted_text = content.decode("utf-8", errors="ignore")
        except Exception as e:
            raise HTTPException(400, f"File reading error: {str(e)}")

    # 2. Process URL (YouTube or Web Page)
    elif url and url.strip():
        u = url.strip()
        parsed = urllib.parse.urlparse(u)
        source_name = u

        # Basic SSRF security check: forbid private IPs, localhost, metadata endpoints
        hostname = (parsed.hostname or "").lower()
        if not hostname or hostname in ("localhost", "127.0.0.1", "0.0.0.0", "::1") or hostname.startswith("192.168.") or hostname.startswith("10.") or hostname.endswith(".local") or "169.254" in hostname:
            raise HTTPException(400, "Invalid or disallowed local network URL.")

        # Check for YouTube URL
        yt_id = None
        if "youtube.com" in hostname and "v=" in parsed.query:
            qs = urllib.parse.parse_qs(parsed.query)
            yt_id = qs.get("v", [None])[0]
        elif "youtu.be" in hostname:
            yt_id = parsed.path.lstrip("/").split("?")[0]

        if yt_id:
            try:
                from youtube_transcript_api import YouTubeTranscriptApi
                transcript = YouTubeTranscriptApi.get_transcript(yt_id)
                extracted_text = " ".join([t.get("text", "") for t in transcript])
                source_name = f"YouTube ({u})"
            except Exception as e:
                # Fallback to page scraping if transcript disabled or unavailable
                try:
                    req = urllib.request.Request(u, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})
                    with urllib.request.urlopen(req, timeout=10) as resp:
                        html = resp.read().decode("utf-8", errors="ignore")
                        from bs4 import BeautifulSoup
                        soup = BeautifulSoup(html, "html.parser")
                        title = soup.title.string if soup.title else u
                        extracted_text = f"YouTube video: {title} (Captions not available). Link: {u}"
                except Exception:
                    raise HTTPException(400, "Could not fetch YouTube transcript. Video may have captions disabled.")
        else:
            # Standard webpage scraping
            try:
                req = urllib.request.Request(u, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})
                with urllib.request.urlopen(req, timeout=12) as resp:
                    html = resp.read().decode("utf-8", errors="ignore")
                    from bs4 import BeautifulSoup
                    soup = BeautifulSoup(html, "html.parser")
                    for tag in soup(["script", "style", "nav", "footer", "header", "aside", "noscript", "svg"]):
                        tag.decompose()
                    extracted_text = " ".join(soup.stripped_strings)
            except Exception as e:
                raise HTTPException(400, f"Could not fetch webpage: {str(e)}")

    clean_text = re.sub(r"\s+", " ", extracted_text).strip()
    if not clean_text:
        raise HTTPException(400, "No readable text content could be extracted.")

    # Limit to first 25,000 characters to fit token limits gracefully
    return {
        "ok": True,
        "source": source_name,
        "text": clean_text[:25000],
        "word_count": len(clean_text.split()),
        "summary": clean_text[:300] + ("..." if len(clean_text) > 300 else "")
    }

SUPPORTED_LAYOUTS = (
    "title", "bullets", "cards", "image", "pennant",
    "hub_arrows", "notebook", "sticky_quad", "pushpin",
    "stats", "flowchart", "sources"
)

def norm_slide(d: dict, fallback_title: str, fallback_slide: dict = None, instruction: str = "") -> dict:
    if not isinstance(d, dict):
        d = {}
    for wrapper in ("slide", "updated_slide", "result", "data", "output"):
        if isinstance(d.get(wrapper), dict):
            d = d[wrapper]
            break

    orig = fallback_slide or {}
    fb_title = str(fallback_title or orig.get("title") or "").strip()
    fb_bullets = list(orig.get("bullets") or [])
    fb_notes = str(orig.get("notes") or "").strip()
    fb_layout = str(orig.get("layout") or "bullets").strip().lower()
    fb_iq = str(orig.get("image_query") or fb_title).strip()

    title = d.get("title") or d.get("heading") or d.get("name")
    if not title or not str(title).strip():
        title = fb_title

    bullets_raw = (
        d.get("bullets")
        or d.get("points")
        or d.get("bullet_points")
        or d.get("items")
        or d.get("content")
        or d.get("subtopics")
        or d.get("key_points")
    )
    if not bullets_raw and fb_layout == "title" and d.get("subtitle"):
        bullets_raw = [d["subtitle"]]

    if isinstance(bullets_raw, str):
        bullets = [b.strip("-•* \t").strip() for b in bullets_raw.split("\n") if b.strip()]
    elif isinstance(bullets_raw, list):
        bullets = []
        for item in bullets_raw:
            if isinstance(item, dict):
                bullets.append(str(item.get("text") or item.get("point") or item.get("title") or item).strip())
            elif item is not None and str(item).strip():
                bullets.append(str(item).strip("-•* \t").strip())
    else:
        bullets = []

    instr_lower = instruction.lower()
    wants_empty = any(w in instr_lower for w in ("remove bullet", "delete bullet", "clear bullet", "no bullet", "remove all points"))
    if not bullets and not wants_empty and fb_bullets:
        bullets = list(fb_bullets)

    notes_raw = d.get("notes") or d.get("speaker_notes") or d.get("script")
    if isinstance(notes_raw, list):
        notes = " ".join(str(n).strip() for n in notes_raw if str(n).strip())
    elif notes_raw is not None and str(notes_raw).strip():
        notes = str(notes_raw).strip()
    else:
        notes = fb_notes

    layout = str(d.get("layout") or "").strip().lower()
    if layout not in SUPPORTED_LAYOUTS:
        if fb_layout == "title":
            layout = "title"
        elif any(k in instr_lower for k in ("card", "quadrant", "swot", "grid")):
            layout = "cards"
        elif any(k in instr_lower for k in ("process", "step", "flow")):
            layout = "flowchart"
        elif any(k in instr_lower for k in ("stat", "number", "metric")):
            layout = "stats"
        else:
            layout = fb_layout or "bullets"

    image_query = str(d.get("image_query") or d.get("image") or "").strip()
    if not image_query:
        image_query = str(d.get("title") or fb_iq or fb_title).strip()

    return {
        "title": str(title).strip(),
        "bullets": bullets,
        "notes": str(notes).strip(),
        "layout": layout,
        "image_query": str(image_query).strip(),
        "stat_number": str(d.get("stat_number", "85%")).strip(),
        "stat_label": str(d.get("stat_label", "Key Benchmark")).strip(),
    }


@app.post("/api/titles")
def titles(p: dict = Body(...)):
    s = p.get("settings", {})
    existing = p.get("existing", [])
    prompt = (
        f"{ctx(s)}\nSuggest 4 different, catchy, modern presentation titles for this topic."
        + (f" Do NOT repeat these: {existing}." if existing else "")
        + ' Return JSON exactly like {"titles": ["...", "...", "...", "..."]}'
    )
    return {"titles": as_list(ask(prompt).get("titles"))[:4]}


@app.post("/api/suggestions")
def suggestions(p: dict = Body(...)):
    s = p.get("settings", {})
    prompt = (
        f"{ctx(s)}\nSuggest 8 important subtopics that a presentation on this topic "
        'should cover. Return JSON exactly like {"subtopics": ["...", "..."]}'
    )
    return {"subtopics": as_list(ask(prompt).get("subtopics"))[:10]}


@app.post("/api/outline")
def outline(p: dict = Body(...)):
    s = p.get("settings", {})
    n = int(p.get("count", 8))
    title = p.get("title", s.get("topic", ""))
    subs = p.get("subtopics", [])
    prompt = (
        f"{ctx(s)}\nPresentation title: {title}.\n"
        f"Create a slide-by-slide outline with EXACTLY {n} slide headings.\n"
        "Slide 1 is the Title Slide; the last slide is a Conclusion / Sources closing slide.\n"
        f"Incorporate these subtopics where relevant: {subs}.\n"
        'Return JSON exactly like {"outline": ["Title Slide", "Overview", "...", "Conclusion & Sources"]}'
    )
    items = as_list(ask(prompt).get("outline"))
    items = items[:n]
    while len(items) < n:
        items.append(f"Topic {len(items) + 1}")
    return {"outline": items}


@app.post("/api/slide")
def make_slide(p: dict = Body(...)):
    s = p.get("settings", {})
    heading = p.get("heading", "")
    idx = int(p.get("index", 1))
    total = int(p.get("total", 1))
    length = p.get("length", "Balanced")
    count = {"Short": "3", "Balanced": "4", "Detailed": "5"}.get(length, "4")
    words = {"Short": "8", "Balanced": "12", "Detailed": "16"}.get(length, "12")

    # Smart layout suggestion based on slide position & title keywords
    is_title = (idx == 1)
    is_last = (idx == total and total > 2)
    h_lower = heading.lower()

    suggested_layout = "bullets"
    if is_title:
        suggested_layout = "title"
    elif is_last:
        suggested_layout = "sources" if any(w in h_lower for w in ("source", "ref", "conclu")) else "cards"
    elif any(w in h_lower for w in ("step", "process", "flow", "timeline", "phase")):
        suggested_layout = "flowchart"
    elif any(w in h_lower for w in ("stat", "metric", "impact", "growth", "data", "result", "number")):
        suggested_layout = "stats"
    elif any(w in h_lower for w in ("vs", "compare", "quadrant", "swot", "matrix", "pros")):
        suggested_layout = "sticky_quad"
    elif any(w in h_lower for w in ("feature", "pillar", "benefit", "team", "part", "role", "core")):
        suggested_layout = "cards"
    elif any(w in h_lower for w in ("hub", "central", "model", "architecture")):
        suggested_layout = "hub_arrows"
    elif any(w in h_lower for w in ("persona", "notebook", "profile", "need", "goal")):
        suggested_layout = "notebook"
    elif any(w in h_lower for w in ("quote", "key take", "principle", "rule")):
        suggested_layout = "pushpin"

    prompt = (
        f"{ctx(s)}\nPresentation title: {p.get('title', '')}.\n"
        f"Write slide {idx} of {total}: heading \"{heading}\".\n"
        f"Outline: {p.get('outline', [])}.\n"
        f"Provide {count} crisp bullet points (max {words} words each), engaging and highly informative for the audience.\n"
        f"Choose the best visual layout from: ['title', 'bullets', 'cards', 'image', 'pennant', 'hub_arrows', 'notebook', 'sticky_quad', 'pushpin', 'stats', 'flowchart', 'sources'].\n"
        f"Suggested layout: '{suggested_layout}'.\n"
        f"Also provide 2-3 sentences of speaker notes, a 2-4 word image search query, and if stat-related, a stat_number (e.g. '78%') and stat_label.\n"
        'Return JSON exactly like: {"title": "...", "bullets": ["...", "..."], "layout": "cards", "notes": "...", "image_query": "...", "stat_number": "85%", "stat_label": "Accuracy Rate"}'
    )
    return norm_slide(ask(prompt), heading)


@app.post("/api/assist")
def assist(p: dict = Body(...)):
    s = p.get("settings", {})
    slide = p.get("slide", {})
    instr = str(p.get("instruction", "")).strip()
    layout = str(slide.get("layout") or "bullets").strip().lower()

    prompt = (
        f"{ctx(s)}\n"
        "You are an expert presentation editor assisting with ONE slide of a presentation.\n"
        f"Current slide content:\n"
        f"{json.dumps({'title': slide.get('title', ''), 'bullets': slide.get('bullets', []), 'notes': slide.get('notes', ''), 'layout': layout}, ensure_ascii=False)}\n\n"
        f"User Instruction: \"{instr}\"\n\n"
        "Update the slide according to the user instruction.\n"
        "Return ONLY a valid JSON object matching:\n"
        '{"title": "...", "bullets": ["..."], "notes": "...", "layout": "...", "image_query": "..."}'
    )
    raw = ask(prompt)
    return norm_slide(raw, fallback_title=slide.get("title", ""), fallback_slide=slide, instruction=instr)


@app.post("/api/extra")
def extra(p: dict = Body(...)):
    s = p.get("settings", {})
    deck = "\n".join(
        f"Slide {i + 1}: {x.get('title', '')} - " + "; ".join(x.get("bullets", []))
        for i, x in enumerate(p.get("slides", []))
    )
    if p.get("kind") == "quiz":
        task = (
            "Create 5 multiple-choice quiz questions based on this presentation. Return JSON like "
            '{"items": [{"q": "...", "options": ["A","B","C","D"], "answer": "A"}]}'
        )
    else:
        task = 'Write a short 4-5 sentence summary. Return JSON like {"summary": "..."}'
    return ask(f"{ctx(s)}\n{task}\nPresentation:\n{deck}")


@app.get("/api/images")
def images(q: str):
    try:
        r = requests.get(
            "https://en.wikipedia.org/w/api.php",
            params={
                "action": "query",
                "generator": "search",
                "gsrsearch": q,
                "gsrlimit": 12,
                "prop": "pageimages",
                "piprop": "thumbnail",
                "pithumbsize": 800,
                "format": "json",
            },
            headers={"User-Agent": "SmartSlideAI/1.0"},
            timeout=15,
        )
        pages = r.json().get("query", {}).get("pages", {})
    except Exception:
        return {"urls": []}
    urls = []
    for pg in sorted(pages.values(), key=lambda x: x.get("index", 99)):
        u = pg.get("thumbnail", {}).get("source")
        if u and u.lower().split("?")[0].endswith((".jpg", ".jpeg", ".png")):
            urls.append(u)
    return {"urls": urls[:6]}

THEMES = {
    "modern": ("0F172A", "F8FAFC", "38BDF8"),
    "light": ("FFFFFF", "0F172A", "4F46E5"),
    "corporate": ("0B2545", "FFFFFF", "F5A623"),
    "minimalist": ("F8FAFC", "1E293B", "475569"),
    "creative": ("3B0764", "FDF4FF", "F472B6"),
    "academic": ("FFF8E7", "2B2118", "9A3412"),
}

def rgb(h: str) -> RGBColor:
    c = h.lstrip("#").upper()
    if len(c) == 3:
        c = "".join([x * 2 for x in c])
    return RGBColor.from_string(c[:6])


def luminance(h: str) -> float:
    h = h.lstrip("#")
    if len(h) < 6:
        return 0.9
    r, g, b = (int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def load_image(src: str):
    if not src:
        return None
    try:
        if src.startswith("data:"):
            return base64.b64decode(src.split(",", 1)[1])
        r = requests.get(src, timeout=15, headers={"User-Agent": "SmartSlideAI/1.0"})
        r.raise_for_status()
        return r.content
    except Exception:
        return None


def textbox(slide, x, y, w, h, color, anchor=None):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    return tf


def add_run(p, text, size, color, bold=False):
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = rgb(color)
    return r


@app.post("/api/export")
def export(p: dict = Body(...)):
    theme_key = p.get("theme", "light")
    design = p.get("design") or {}
    
    # Theme & design colors
    bg_hex = (design.get("bg_start") or "").lstrip("#")
    if not bg_hex or len(bg_hex) < 6:
        bg_hex, fg_hex, accent_hex = THEMES.get(theme_key, THEMES["light"])
    else:
        accent_hex = (design.get("accent") or "4F46E5").lstrip("#")
        fg_hex = "0F172A" if luminance(bg_hex) > 0.55 else "F8FAFC"

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)

    for idx, s in enumerate(p.get("slides", [])):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide_bg = (s.get("bg") or "").lstrip("#") or bg_hex
        slide_fg = (("0F172A" if luminance(slide_bg) > 0.55 else "F8FAFC") if s.get("bg") else fg_hex)

        # 1. Slide Solid Background
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = rgb(slide_bg)

        # 2. Subtle decorative top accent bar
        top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.08))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = rgb(accent_hex)
        top_bar.line.fill.background()

        layout = s.get("layout", "bullets")
        size = int(s.get("fontSize") or 24)
        bullets = s.get("bullets", [])
        title = s.get("title", "")

        # 3. Layout Rendering
        if layout == "title":
            # Decorative Pill Tag
            tag = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), Inches(2.2), Inches(1.8), Inches(0.4))
            tag.fill.solid()
            tag.fill.fore_color.rgb = rgb(accent_hex)
            tag.line.fill.background()
            add_run(tag.text_frame.paragraphs[0], "PRESENTATION", 12, "FFFFFF", True)
            tag.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

            # Main Presentation Title
            tf = textbox(slide, 1.2, 2.8, 10.8, 2.2, slide_fg)
            add_run(tf.paragraphs[0], title, 48, slide_fg, True)
            if bullets:
                tf2 = textbox(slide, 1.2, 4.8, 10.8, 1.4, slide_fg)
                add_run(tf2.paragraphs[0], bullets[0], 22, slide_fg)

        elif layout in ("cards", "sticky_quad", "notebook", "hub_arrows"):
            # Header
            tf = textbox(slide, 0.8, 0.45, 11.7, 1.0, slide_fg)
            add_run(tf.paragraphs[0], title, 34, slide_fg, True)
            
            # 2x2 or 3 Column Card Grid
            items = bullets[:6]
            cols = 3 if len(items) == 3 else 2
            card_w = Inches(3.6 if cols == 3 else 5.6)
            card_h = Inches(2.4 if len(items) > 2 else 4.0)

            for i, b in enumerate(items):
                col_i = i % cols
                row_i = i // cols
                cx = Inches(0.8) + col_i * (card_w + Inches(0.4))
                cy = Inches(1.8) + row_i * (card_h + Inches(0.3))

                card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, cy, card_w, card_h)
                card.fill.solid()
                card.fill.fore_color.rgb = RGBColor(255, 255, 255) if luminance(slide_bg) > 0.55 else RGBColor(30, 41, 59)
                card.line.color.rgb = rgb(accent_hex)
                card.line.width = Pt(1.5)

                ctf = card.text_frame
                ctf.word_wrap = True
                ctf.margin_top = Inches(0.2)
                ctf.margin_left = Inches(0.25)
                ctf.margin_right = Inches(0.25)
                
                # Badge number
                p0 = ctf.paragraphs[0]
                add_run(p0, f"0{i+1}. ", 14, accent_hex, True)
                add_run(p0, b, size - 4, slide_fg, False)

        elif layout == "stats":
            # Header
            tf = textbox(slide, 0.8, 0.45, 11.7, 1.0, slide_fg)
            add_run(tf.paragraphs[0], title, 34, slide_fg, True)

            # Left Big Stat Card
            scard = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.8), Inches(4.5), Inches(4.8))
            scard.fill.solid()
            scard.fill.fore_color.rgb = rgb(accent_hex)
            scard.line.fill.background()

            stf = scard.text_frame
            stf.word_wrap = True
            p_stat = stf.paragraphs[0]
            p_stat.alignment = PP_ALIGN.CENTER
            add_run(p_stat, s.get("stat_number", "85%") + "\n", 56, "FFFFFF", True)
            p_lbl = stf.add_paragraph()
            p_lbl.alignment = PP_ALIGN.CENTER
            add_run(p_lbl, s.get("stat_label", "Key Benchmark Impact"), 18, "FFFFFF", False)

            # Right content
            rtf = textbox(slide, 5.8, 1.9, 6.7, 4.6, slide_fg)
            for b in bullets:
                p = rtf.add_paragraph()
                p.space_after = Pt(14)
                add_run(p, "• " + b, size, slide_fg)

        elif layout == "flowchart":
            tf = textbox(slide, 0.8, 0.45, 11.7, 1.0, slide_fg)
            add_run(tf.paragraphs[0], title, 34, slide_fg, True)

            items = bullets[:4]
            step_w = Inches(2.5)
            step_h = Inches(3.6)
            for i, b in enumerate(items):
                sx = Inches(0.8) + i * (step_w + Inches(0.6))
                step_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, sx, Inches(2.2), step_w, step_h)
                step_box.fill.solid()
                step_box.fill.fore_color.rgb = RGBColor(255, 255, 255) if luminance(slide_bg) > 0.55 else RGBColor(30, 41, 59)
                step_box.line.color.rgb = rgb(accent_hex)

                stf = step_box.text_frame
                stf.word_wrap = True
                p0 = stf.paragraphs[0]
                add_run(p0, f"STEP {i+1}\n\n", 13, accent_hex, True)
                add_run(p0, b, size - 4, slide_fg, False)

                # Add arrow connector between steps
                if i < len(items) - 1:
                    ax = sx + step_w + Inches(0.1)
                    arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, ax, Inches(3.6), Inches(0.4), Inches(0.3))
                    arrow.fill.solid()
                    arrow.fill.fore_color.rgb = rgb(accent_hex)
                    arrow.line.fill.background()

        else:  # bullets, image, sources
            tf = textbox(slide, 0.8, 0.45, 11.7, 1.0, slide_fg)
            add_run(tf.paragraphs[0], title, 34, slide_fg, True)
            
            # Subtle accent underline
            bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(1.2), Inches(0.06))
            bar.fill.solid()
            bar.fill.fore_color.rgb = rgb(accent_hex)
            bar.line.fill.background()

            img_bytes = load_image(s.get("image_url", "")) if layout == "image" else None
            width = 6.4 if img_bytes else 11.7
            tf_body = textbox(slide, 0.8, 1.8, width, 5.0, slide_fg)
            first = True
            for b in bullets:
                para = tf_body.paragraphs[0] if first else tf_body.add_paragraph()
                first = False
                para.space_after = Pt(12)
                add_run(para, "• " + b, size, slide_fg)

            if img_bytes:
                try:
                    from PIL import Image
                    iw, ih = Image.open(io.BytesIO(img_bytes)).size
                    sc = min(5.0 / iw, 4.6 / ih)
                    pw, ph = iw * sc, ih * sc
                    slide.shapes.add_picture(
                        io.BytesIO(img_bytes),
                        Inches(7.5 + (5.0 - pw) / 2), Inches(1.9), Inches(pw), Inches(ph),
                    )
                except Exception:
                    pass

        # Footer with Slide Number
        ftf = textbox(slide, 0.8, 6.9, 11.7, 0.4, slide_fg)
        fp = ftf.paragraphs[0]
        deck_name = p.get("title", "SmartSlide AI")
        add_run(fp, f"{deck_name}  |  Slide {idx + 1} of {len(p.get('slides', []))}", 10, "94A3B8")

        # Speaker notes
        if s.get("notes"):
            slide.notes_slide.notes_text_frame.text = s["notes"]

    buf = io.BytesIO()
    prs.save(buf)
    name = re.sub(r"[^A-Za-z0-9_-]+", "_", p.get("title", "presentation"))[:50] or "presentation"
    return Response(
        buf.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f'attachment; filename="{name}.pptx"'},
    )


# ----------------------------------------------------------------------------
# Factcheck & Feedback endpoints
# ----------------------------------------------------------------------------

UA = {"User-Agent": "SmartSlideAI/1.0"}

def wiki_context(query: str):
    try:
        r = requests.get(
            "https://en.wikipedia.org/w/api.php",
            params={
                "action": "query", "generator": "search", "gsrsearch": query, "gsrlimit": 3,
                "prop": "extracts|info", "exintro": 1, "explaintext": 1, "exchars": 1200,
                "inprop": "url", "format": "json",
            },
            headers=UA, timeout=15,
        )
        pages = sorted(r.json().get("query", {}).get("pages", {}).values(), key=lambda x: x.get("index", 99))
    except Exception:
        return []
    return [
        {"title": p.get("title", ""), "url": p.get("fullurl", ""), "extract": p.get("extract", "")}
        for p in pages if p.get("extract")
    ]


@app.post("/api/factcheck")
def factcheck(p: dict = Body(...)):
    bullets = p.get("bullets", [])
    if not bullets:
        return {"checks": []}
    sources = wiki_context(p.get("query", "") or p.get("title", ""))
    src_text = "\n\n".join(f"[{i + 1}] {s['title']}: {s['extract']}" for i, s in enumerate(sources))
    prompt = (
        "You are an objective fact-checker for presentation slides.\n"
        f"Slide Title: {p.get('title', '')}\n"
        f"Slide Content:\n"
        + "\n".join(f"- {b}" for b in bullets)
        + "\n\nReference Wikipedia summaries:\n"
        + (src_text or "No specific Wikipedia summary available.")
        + "\n\nFor EACH bullet point, verify its factual accuracy and return ONLY JSON like:\n"
        '{"checks": [{"verdict": "correct|unclear|wrong", "confidence": "high|medium|low", "issue": "brief reason", "correction": "how to fix"}]}'
    )
    raw = ask(prompt).get("checks", [])
    checks = []
    for i, b in enumerate(bullets):
        c = raw[i] if i < len(raw) and isinstance(raw[i], dict) else {}
        checks.append({
            "bullet": b,
            "verdict": c.get("verdict", "unclear"),
            "confidence": c.get("confidence", "low"),
            "issue": c.get("issue", ""),
            "correction": c.get("correction", ""),
        })
    return {"checks": checks, "sources": sources}


FEEDBACK_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "feedback.jsonl")

@app.post("/api/feedback")
def feedback(p: dict = Body(...)):
    msg = str(p.get("message", "")).strip()
    if not msg:
        raise HTTPException(400, "Feedback message cannot be empty.")
    entry = {
        "time": datetime.datetime.now().isoformat(timespec="seconds"),
        "rating": max(1, min(5, int(p.get("rating") or 5))),
        "category": str(p.get("category", "General"))[:40],
        "email": str(p.get("email", ""))[:120],
        "message": msg[:3000],
    }
    with open(FEEDBACK_FILE, "a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False) + "\n")
    return {"ok": True}


@app.get("/api/feedback")
def feedback_list():
    if not os.path.exists(FEEDBACK_FILE):
        return []
    with open(FEEDBACK_FILE, encoding="utf-8") as f:
        return [json.loads(line) for line in f if line.strip()]


@app.get("/api/health")
def health():
    p = CFG["provider"]
    if p == "ollama":
        try:
            tags = requests.get(f"{CFG['base_url'].rstrip('/')}/api/tags", timeout=3).json()
            names = [m["name"] for m in tags.get("models", [])]
        except Exception:
            return {"ready": False, "message": 'Ollama is not running. Start "ollama serve" or enter your API key in AI Settings.'}
        if not any(n.startswith(CFG["model"]) for n in names):
            return {"ready": False, "message": f'Model "{CFG["model"]}" not installed. Run: ollama pull {CFG["model"]}'}
        return {"ready": True, "message": ""}
    if not CFG["api_key"] or not CFG["model"]:
        return {"ready": False, "message": "API key or model not configured. Click the AI Settings button."}
    return {"ready": True, "message": ""}


@app.get("/", response_class=HTMLResponse)
def index():
    return INDEX_HTML

INDEX_HTML = r'''<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>SmartSlide AI - Intelligent Presentation Studio</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Noto+Sans+Devanagari:wght@400;600;700&display=swap" rel="stylesheet">
<script src="https://cdn.tailwindcss.com"></script>
<style>
  :root {
    --accent: #4f46e5;
    --accent-light: #eef2ff;
  }
  body {
    font-family: 'Plus Jakarta Sans', 'Noto Sans Devanagari', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    background-color: #f8fafc;
    color: #0f172a;
    -webkit-font-smoothing: antialiased;
  }
  .btn {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    gap: 0.5rem;
    padding: 0.55rem 1rem;
    border-radius: 0.85rem;
    font-weight: 600;
    font-size: 0.875rem;
    transition: all 0.2s cubic-bezier(0.16, 1, 0.3, 1);
    white-space: nowrap;
    cursor: pointer;
  }
  .btn:hover {
    transform: translateY(-1px);
  }
  .btn:active {
    transform: translateY(0);
  }
  .btn-p {
    background: linear-gradient(135deg, #4f46e5 0%, #6366f1 100%);
    color: #ffffff;
    box-shadow: 0 4px 12px rgba(79, 70, 229, 0.25);
  }
  .btn-p:hover {
    box-shadow: 0 6px 18px rgba(79, 70, 229, 0.35);
  }
  .btn-s {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    color: #334155;
    box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04);
  }
  .btn-s:hover {
    background: #f8fafc;
    border-color: #cbd5e1;
    color: #0f172a;
  }
  .btn:disabled {
    opacity: 0.45;
    cursor: not-allowed;
    transform: none !important;
    box-shadow: none !important;
  }
  .inp {
    width: 100%;
    border: 1px solid #e2e8f0;
    border-radius: 0.85rem;
    padding: 0.6rem 0.9rem;
    font-size: 0.9rem;
    background: #ffffff;
    color: #0f172a;
    transition: all 0.2s ease;
  }
  .inp:focus {
    outline: none;
    border-color: #6366f1;
    box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.15);
  }
  .card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 1.25rem;
    box-shadow: 0 4px 20px -2px rgba(0, 0, 0, 0.05);
  }
  .glass-card {
    background: rgba(255, 255, 255, 0.85);
    backdrop-filter: blur(12px);
    border: 1px solid rgba(226, 232, 240, 0.8);
  }
  .thumb.active {
    outline: 3px solid #6366f1;
    outline-offset: 2px;
  }
  .chip {
    display: inline-flex;
    align-items: center;
    gap: 0.35rem;
    padding: 0.35rem 0.75rem;
    border-radius: 9999px;
    font-size: 0.8rem;
    font-weight: 500;
    background: #ffffff;
    border: 1px solid #e2e8f0;
    color: #475569;
    transition: all 0.15s ease;
    cursor: pointer;
  }
  .chip:hover {
    border-color: #6366f1;
    color: #4f46e5;
    background: #eef2ff;
    transform: translateY(-1px);
  }
  .pulse-mic {
    animation: pulse-ring 1.5s cubic-bezier(0.215, 0.61, 0.355, 1) infinite;
  }
  @keyframes pulse-ring {
    0% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(239, 68, 68, 0.7); }
    70% { transform: scale(1); box-shadow: 0 0 0 10px rgba(239, 68, 68, 0); }
    100% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(239, 68, 68, 0); }
  }
  .hidden { display: none !important; }
</style>
</head>
<body class="bg-slate-50 text-slate-900 min-h-screen flex flex-col selection:bg-indigo-100 selection:text-indigo-900">

<!-- Global Toast & Modals -->
<div id="toast" class="hidden fixed top-5 right-5 z-[100] max-w-sm px-4 py-3 rounded-2xl text-white text-sm shadow-xl flex items-center gap-2 transition-all"></div>

<div id="busy" class="hidden fixed inset-0 z-[90] bg-slate-900/40 backdrop-blur-sm flex items-center justify-center">
  <div class="bg-white rounded-3xl px-8 py-6 shadow-2xl flex flex-col items-center gap-4 text-center max-w-xs border border-slate-100">
    <div class="relative w-12 h-12">
      <div class="w-12 h-12 border-4 border-indigo-100 border-t-indigo-600 rounded-full animate-spin"></div>
    </div>
    <div>
      <h4 id="busyMsg" class="font-bold text-slate-800 text-base">Crafting slides...</h4>
      <p class="text-xs text-slate-400 mt-1">SmartSlide AI is generating structured insights</p>
    </div>
  </div>
</div>

<div id="modal" class="hidden fixed inset-0 z-[80] bg-slate-900/50 backdrop-blur-sm flex items-center justify-center p-4" onclick="if(event.target===this)closeModal()">
  <div class="bg-white rounded-3xl shadow-2xl max-w-2xl w-full max-h-[88vh] overflow-y-auto p-6 sm:p-8 border border-slate-100" id="modalBody"></div>
</div>

<!-- Fullscreen Presenter Mode -->
<div id="present" class="hidden fixed inset-0 z-[99] bg-slate-950 flex flex-col items-center justify-center p-4">
  <div class="absolute top-4 right-4 flex items-center gap-3 z-10">
    <span id="presTimer" class="px-3 py-1.5 rounded-xl bg-slate-800/80 text-indigo-300 font-mono text-sm border border-slate-700">00:00</span>
    <button class="btn btn-s !bg-slate-800 !text-white !border-slate-700 !py-1 text-xs" onclick="endPresent()">Exit (ESC)</button>
  </div>
  <div id="presCanvasWrap" class="w-full flex-1 flex items-center justify-center overflow-hidden">
    <div id="presCanvas" class="shadow-2xl rounded-2xl overflow-hidden"></div>
  </div>
  <div class="w-full max-w-4xl mt-3 flex items-center justify-between text-slate-400 text-sm px-4">
    <button class="btn btn-s !bg-slate-800 !text-slate-200 !border-slate-700 !py-1 text-xs" onclick="presPrev()">&larr; Previous</button>
    <span id="presCounter" class="font-medium text-slate-300">Slide 1 / 1</span>
    <button class="btn btn-s !bg-slate-800 !text-slate-200 !border-slate-700 !py-1 text-xs" onclick="presNext()">Next &rarr;</button>
  </div>
  <div id="presNotesBox" class="w-full max-w-4xl mt-2 p-3 rounded-xl bg-slate-900/80 border border-slate-800 text-xs text-slate-300 max-h-20 overflow-y-auto text-left">
    <span class="font-bold text-indigo-400 uppercase tracking-wider text-[10px] block mb-1">Speaker Notes:</span>
    <span id="presNotesText">No notes for this slide.</span>
  </div>
</div>

<!-- ======================================================== -->
<!-- SCREEN: HOME / HERO                                      -->
<!-- ======================================================== -->
<section id="home" class="max-w-6xl mx-auto w-full p-4 sm:p-8 flex-1 flex flex-col">
  <!-- Top Navigation -->
  <nav class="flex items-center justify-between mb-8 pb-4 border-b border-slate-200/80">
    <div class="flex items-center gap-3">
      <div class="w-10 h-10 rounded-2xl bg-gradient-to-tr from-indigo-600 to-violet-500 flex items-center justify-center text-white font-extrabold text-xl shadow-lg shadow-indigo-500/25">
        ✦
      </div>
      <div>
        <h1 class="text-xl font-extrabold tracking-tight text-slate-900">SmartSlide <span class="text-indigo-600">AI</span></h1>
        <p class="text-[11px] font-medium text-slate-500">Aesthetic Presentation Studio</p>
      </div>
    </div>
    <div class="flex items-center gap-2">
      <button class="btn btn-s text-xs" onclick="openSettings()">⚙ AI Settings</button>
      <button class="btn btn-s text-xs" onclick="openFeedback()">💬 Feedback</button>
    </div>
  </nav>

  <!-- Hero Section -->
  <div class="text-center max-w-3xl mx-auto py-6 sm:py-10">
    <span class="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-semibold bg-indigo-50 text-indigo-700 border border-indigo-100 mb-4 shadow-sm">
      ✨ Topic-Based Aesthetic Backgrounds &amp; Multi-Source Input
    </span>
    <h2 class="text-3xl sm:text-5xl font-extrabold text-slate-900 tracking-tight leading-tight mb-4">
      Turn any thought into <span class="bg-gradient-to-r from-indigo-600 to-violet-600 bg-clip-text text-transparent">stunning presentations</span> in seconds.
    </h2>
    <p class="text-slate-600 text-sm sm:text-base mb-8 max-w-xl mx-auto">
      Type a topic, upload a document, paste a YouTube link, or speak with your voice. SmartSlide AI crafts customized slides with Gamma-style visual designs.
    </p>

    <!-- Big Prompt Box -->
    <div class="card p-3 sm:p-4 max-w-2xl mx-auto shadow-2xl shadow-indigo-100/60 border-indigo-100/80 relative">
      <div class="flex items-center gap-2 mb-2">
        <input id="hero_topic" class="inp !border-0 !p-2 !text-base sm:!text-lg font-medium placeholder:text-slate-400 focus:!ring-0" placeholder="What would you like to present? e.g. Quantum Computing..." onkeydown="if(event.key==='Enter')startFromHero()">
        <button id="hero_mic" class="btn btn-s !p-2.5 rounded-full text-slate-600 hover:text-red-600 hover:border-red-200" title="Voice Input (English/Hindi)" onclick="toggleVoice('hero_topic', 'hero_mic')">
          🎤
        </button>
      </div>
      <div class="flex flex-col sm:flex-row items-center justify-between gap-3 pt-2 border-t border-slate-100">
        <div class="flex items-center gap-2 w-full sm:w-auto">
          <button class="btn btn-s !py-1.5 !px-3 text-xs w-full sm:w-auto" onclick="openSourceModal()">
            📎 Add Source (PDF/Doc/YouTube/URL)
          </button>
          <span id="heroSourceTag" class="hidden text-xs px-2.5 py-1 rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200 font-medium truncate max-w-[160px]"></span>
        </div>
        <button class="btn btn-p !py-2.5 !px-6 text-sm w-full sm:w-auto" onclick="startFromHero()">
          Generate Deck ✦
        </button>
      </div>
    </div>

    <!-- Suggested Topics Chips -->
    <div class="flex items-center justify-center gap-2 flex-wrap mt-5">
      <span class="text-xs font-semibold text-slate-400 mr-1">Popular Ideas:</span>
      <span class="chip" onclick="pickHeroIdea('Artificial Intelligence &amp; Robotics')">🤖 AI &amp; Robotics</span>
      <span class="chip" onclick="pickHeroIdea('Space Exploration &amp; Mars Colonization')">🚀 Space Exploration</span>
      <span class="chip" onclick="pickHeroIdea('Sustainable Green Architecture')">🌿 Green Architecture</span>
      <span class="chip" onclick="pickHeroIdea('Psychology of Decision Making')">🧠 Psychology</span>
      <span class="chip" onclick="pickHeroIdea('The Future of Renewable Energy')">⚡ Renewable Energy</span>
    </div>
  </div>

  <div id="health" class="hidden my-4 p-3.5 rounded-2xl bg-amber-50 border border-amber-200 text-amber-900 text-xs sm:text-sm font-medium text-center"></div>

  <!-- Saved Presentations List -->
  <div class="mt-8 pt-6 border-t border-slate-200/80">
    <div class="flex items-center justify-between gap-4 mb-4 flex-wrap">
      <h3 class="text-lg font-bold text-slate-900">Your Presentations</h3>
      <div class="flex items-center gap-2">
        <label class="btn btn-s text-xs cursor-pointer">
          📂 Import JSON
          <input type="file" accept=".json" class="hidden" onchange="importDeck(this)">
        </label>
        <input id="search" class="inp !py-1.5 text-xs w-44 sm:w-60" placeholder="Search decks..." oninput="renderHome()">
      </div>
    </div>
    <div id="deckList" class="grid sm:grid-cols-2 lg:grid-cols-3 gap-5"></div>
  </div>
</section>

<!-- ======================================================== -->
<!-- SCREEN: SETUP WIZARD                                     -->
<!-- ======================================================== -->
<section id="setup" class="hidden max-w-3xl mx-auto w-full p-4 sm:p-6 flex-1 flex flex-col">
  <button class="btn btn-s mb-4 text-xs self-start" onclick="show('home')">&larr; Back to Home</button>
  <div class="card p-6 sm:p-8 space-y-6">
    <div class="border-b pb-4">
      <h2 class="text-2xl font-extrabold text-slate-900">Configure Presentation</h2>
      <p class="text-xs text-slate-500 mt-1">Customize audience, visual depth, and parameters before generating.</p>
    </div>

    <!-- Topic & Voice -->
    <div>
      <label class="text-xs font-bold text-slate-700 uppercase tracking-wider block mb-1.5">Topic or Theme *</label>
      <div class="flex items-center gap-2">
        <input id="f_topic" class="inp font-medium text-base" placeholder="e.g. History of Renaissance Art">
        <button id="f_mic" class="btn btn-s !p-2.5 rounded-xl text-slate-600 hover:text-red-600 hover:border-red-200 shrink-0" title="Voice Input" onclick="toggleVoice('f_topic', 'f_mic')">
          🎤
        </button>
      </div>
    </div>

    <!-- Attached Source Notification -->
    <div id="setupSourceCard" class="hidden p-3.5 rounded-2xl bg-indigo-50/80 border border-indigo-200 flex items-center justify-between text-xs">
      <div class="flex items-center gap-2 text-indigo-900 font-medium truncate">
        <span>📎 Source Attached:</span>
        <span id="setupSourceName" class="font-bold truncate max-w-sm"></span>
      </div>
      <button class="text-indigo-600 hover:text-indigo-800 font-bold ml-2 underline shrink-0" onclick="clearAttachedSource()">Remove</button>
    </div>

    <div class="grid sm:grid-cols-2 gap-4">
      <div>
        <label class="text-xs font-bold text-slate-700 uppercase tracking-wider block mb-1.5">Number of Slides</label>
        <select id="f_count" class="inp font-medium" onchange="$('f_custom').classList.toggle('hidden',this.value!=='custom')">
          <option>5</option><option selected>8</option><option>10</option><option>12</option><option>15</option><option value="custom">Custom...</option>
        </select>
        <input id="f_custom" type="number" min="2" max="30" value="12" class="inp mt-2 hidden">
      </div>
      <div>
        <label class="text-xs font-bold text-slate-700 uppercase tracking-wider block mb-1.5">Language</label>
        <select id="f_language" class="inp font-medium">
          <option>English</option><option>Hindi</option><option>Spanish</option><option>French</option><option>German</option><option>Marathi</option><option>Bengali</option><option>Tamil</option>
        </select>
      </div>
      <div>
        <label class="text-xs font-bold text-slate-700 uppercase tracking-wider block mb-1.5">Presentation Type</label>
        <select id="f_type" class="inp font-medium">
          <option>Educational</option><option>Business / Startup Pitch</option><option>Academic Research</option><option>Technical Deep Dive</option><option>Marketing &amp; Strategy</option><option>Storytelling</option>
        </select>
      </div>
      <div>
        <label class="text-xs font-bold text-slate-700 uppercase tracking-wider block mb-1.5">Target Audience</label>
        <select id="f_audience" class="inp font-medium">
          <option>General Audience</option><option>Students &amp; Learners</option><option>Executives &amp; Investors</option><option>Engineers &amp; Developers</option><option>Clients / Customers</option>
        </select>
      </div>
      <div>
        <label class="text-xs font-bold text-slate-700 uppercase tracking-wider block mb-1.5">Aesthetic Background Mood</label>
        <select id="f_theme" class="inp font-medium">
          <option value="auto">✨ Auto Gamma AI Topic Background</option>
          <option value="modern">Modern (Dark Slate)</option>
          <option value="light">Light Professional</option>
          <option value="corporate">Corporate Navy</option>
          <option value="minimalist">Minimalist Paper</option>
          <option value="creative">Creative Violet</option>
          <option value="academic">Academic Warm</option>
        </select>
      </div>
      <div>
        <label class="text-xs font-bold text-slate-700 uppercase tracking-wider block mb-1.5">Content Detail</label>
        <select id="f_length" class="inp font-medium">
          <option>Short &amp; Punchy</option><option selected>Balanced</option><option>Detailed &amp; In-depth</option>
        </select>
      </div>
    </div>

    <!-- Feature Switches -->
    <div class="space-y-2 pt-2 border-t border-slate-100">
      <label class="flex items-center gap-2.5 text-xs text-slate-700 cursor-pointer select-none">
        <input type="checkbox" id="f_images" checked class="w-4 h-4 rounded text-indigo-600 focus:ring-indigo-500">
        <span>Automatically fetch relevant Wikipedia images for slides</span>
      </label>
      <label class="flex items-center gap-2.5 text-xs text-slate-700 cursor-pointer select-none">
        <input type="checkbox" id="f_sourceslide" checked class="w-4 h-4 rounded text-indigo-600 focus:ring-indigo-500">
        <span>Include a dedicated "Sources &amp; References" closing slide</span>
      </label>
    </div>

    <div>
      <label class="text-xs font-bold text-slate-700 uppercase tracking-wider block mb-1.5">Custom Instructions (Optional)</label>
      <textarea id="f_instr" rows="2" class="inp" placeholder="e.g. Focus specifically on recent 2026 breakthroughs, include real statistics..."></textarea>
    </div>

    <button class="btn btn-p w-full !py-3.5 text-base font-bold shadow-xl shadow-indigo-500/20" onclick="goTitles()">
      Next: Suggest Titles &amp; Topics &rarr;
    </button>
  </div>
</section>

<!-- ======================================================== -->
<!-- SCREEN: TITLES & TOPICS                                  -->
<!-- ======================================================== -->
<section id="titles" class="hidden max-w-3xl mx-auto w-full p-4 sm:p-6 flex-1 flex flex-col">
  <button class="btn btn-s mb-4 text-xs self-start" onclick="show('setup')">&larr; Back to Setup</button>
  <div class="card p-6 sm:p-8 space-y-6">
    <div class="border-b pb-4">
      <h2 class="text-2xl font-extrabold text-slate-900">Choose Presentation Title</h2>
      <p class="text-xs text-slate-500 mt-1">Select the most engaging title and verify core subtopics.</p>
    </div>

    <div id="titleList" class="space-y-2.5"></div>
    <button class="btn btn-s text-xs" onclick="moreTitles()">🔄 Generate More Titles</button>

    <div>
      <label class="text-xs font-bold text-slate-700 uppercase tracking-wider block mb-1.5">Subtitle (Optional)</label>
      <input id="subtitle" class="inp" placeholder="e.g. A comprehensive overview for 2026">
    </div>

    <div>
      <h3 class="text-xs font-bold text-slate-700 uppercase tracking-wider block mb-2">
        Subtopics to Cover <span class="text-slate-400 font-normal lowercase">(uncheck to omit)</span>
      </h3>
      <div id="subList" class="grid sm:grid-cols-2 gap-2.5"></div>
    </div>

    <button class="btn btn-p w-full !py-3.5 text-base font-bold shadow-xl shadow-indigo-500/20" onclick="goOutline()">
      Next: Generate Slide Outline &rarr;
    </button>
  </div>
</section>

<!-- ======================================================== -->
<!-- SCREEN: OUTLINE                                          -->
<!-- ======================================================== -->
<section id="outline" class="hidden max-w-3xl mx-auto w-full p-4 sm:p-6 flex-1 flex flex-col">
  <button class="btn btn-s mb-4 text-xs self-start" onclick="show('titles')">&larr; Back to Titles</button>
  <div class="card p-6 sm:p-8 space-y-6">
    <div class="border-b pb-4">
      <h2 class="text-2xl font-extrabold text-slate-900">Review &amp; Refine Outline</h2>
      <p class="text-xs text-slate-500 mt-1">Reorder, edit headings, or add slides before generating presentation content.</p>
    </div>

    <div id="outlineList" class="space-y-2.5"></div>

    <div class="flex gap-2 flex-wrap">
      <button class="btn btn-s text-xs" onclick="addOutline()">+ Add Slide</button>
      <button class="btn btn-s text-xs" onclick="regenOutline()">🔄 Regenerate Outline</button>
    </div>

    <button class="btn btn-p w-full !py-3.5 text-base font-bold shadow-xl shadow-indigo-500/20" onclick="generateDeck()">
      Approve &amp; Generate Full Presentation ✦
    </button>
  </div>
</section>

<!-- ======================================================== -->
<!-- SCREEN: GENERATING (SKELETON & PROGRESS)                 -->
<!-- ======================================================== -->
<section id="generating" class="hidden max-w-lg mx-auto w-full p-6 mt-16 text-center">
  <div class="card p-8 sm:p-10 shadow-2xl space-y-6">
    <div class="relative w-16 h-16 mx-auto">
      <div class="w-16 h-16 border-4 border-indigo-100 border-t-indigo-600 rounded-full animate-spin"></div>
      <div class="absolute inset-0 flex items-center justify-center font-bold text-xs text-indigo-600" id="genPercent">0%</div>
    </div>
    <div>
      <h3 class="text-xl font-extrabold text-slate-900 mb-1">Generating Presentation...</h3>
      <p id="genMsg" class="text-xs text-slate-500 font-medium">Writing slides and tailoring aesthetic layouts</p>
    </div>
    <div class="w-full bg-slate-100 rounded-full h-3 overflow-hidden p-0.5 border border-slate-200">
      <div id="genBar" class="bg-gradient-to-r from-indigo-600 to-violet-600 h-2 rounded-full transition-all duration-300" style="width:0%"></div>
    </div>
    <div class="grid grid-cols-3 gap-2 pt-2">
      <div class="p-2.5 rounded-xl bg-slate-50 border border-slate-100 text-[11px] font-medium text-slate-500 animate-pulse">Design System</div>
      <div class="p-2.5 rounded-xl bg-slate-50 border border-slate-100 text-[11px] font-medium text-slate-500 animate-pulse">Content Writing</div>
      <div class="p-2.5 rounded-xl bg-slate-50 border border-slate-100 text-[11px] font-medium text-slate-500 animate-pulse">Wikipedia Images</div>
    </div>
  </div>
</section>

<!-- ======================================================== -->
<!-- SCREEN: FULL SLIDE STUDIO EDITOR                         -->
<!-- ======================================================== -->
<section id="editor" class="hidden min-h-screen md:h-screen flex flex-col bg-slate-100">
  <!-- Top Navigation & Controls Bar -->
  <header class="bg-white border-b border-slate-200/90 px-3 sm:px-5 py-2.5 flex items-center justify-between gap-3 shrink-0 shadow-sm z-20 flex-wrap">
    <div class="flex items-center gap-2.5 flex-1 min-w-[240px]">
      <button class="btn btn-s !py-1.5 !px-3 text-xs" onclick="goHome()">&larr; Home</button>
      <input id="deckName" class="inp !w-44 sm:!w-64 font-bold text-xs sm:text-sm !py-1.5" oninput="S.title=this.value;scheduleSave()" placeholder="Presentation Title">
      <button class="btn btn-s !p-1.5 text-xs hidden sm:inline-flex" id="undoBtn" onclick="undo()" title="Undo (Ctrl+Z)">↶</button>
      <button class="btn btn-s !p-1.5 text-xs hidden sm:inline-flex" id="redoBtn" onclick="redo()" title="Redo (Ctrl+Y)">↷</button>
      <span id="saveStatus" class="text-[11px] font-medium text-slate-400">All changes saved</span>
    </div>

    <div class="flex items-center gap-2 flex-wrap">
      <!-- Background / Theme Selector -->
      <select id="themeSel" class="inp !w-auto !py-1.5 text-xs font-semibold" onchange="changeThemeOrDesign(this.value)">
        <option value="topic">✨ Topic Background (Gamma AI)</option>
        <option value="light">Light Slate</option>
        <option value="modern">Dark Modern</option>
        <option value="corporate">Corporate Navy</option>
        <option value="minimalist">Minimalist</option>
        <option value="creative">Creative Violet</option>
        <option value="academic">Academic Warm</option>
      </select>

      <button class="btn btn-s !py-1.5 text-xs hidden md:inline-flex" onclick="regenDesignStyle()" title="Regenerate Gamma AI topic background colors">
        🎨 New Palette
      </button>

      <button class="btn btn-s !py-1.5 text-xs" onclick="startPresent()" title="Presenter Fullscreen Mode">
        ▶ Present
      </button>

      <button class="btn btn-s !py-1.5 text-xs" onclick="exportDeckJSON()" title="Download Presentation Backup JSON">
        💾 Save JSON
      </button>

      <button class="btn btn-p !py-1.5 !px-4 text-xs font-bold" onclick="downloadPPTX()">
        Download .pptx
      </button>
    </div>
  </header>

  <!-- Workspace Body -->
  <div class="flex-1 flex flex-col md:flex-row min-h-0 overflow-hidden">
    <!-- Left: Slide Thumbnails Panel -->
    <aside class="w-full md:w-56 bg-slate-50 border-b md:border-b-0 md:border-r border-slate-200/80 flex flex-row md:flex-col shrink-0 z-10">
      <div class="p-2.5 flex items-center gap-1.5 border-b border-slate-200/60 shrink-0">
        <button class="btn btn-s flex-1 !py-1 text-xs" onclick="addSlide()">+ Add</button>
        <button class="btn btn-s !px-2.5 !py-1 text-xs" title="Duplicate Slide" onclick="dupSlide()">⧉</button>
        <button class="btn btn-s !px-2.5 !py-1 text-xs text-red-600 hover:border-red-200" title="Delete Slide" onclick="delSlide()">🗑</button>
      </div>
      <div id="thumbs" class="flex-1 overflow-x-auto md:overflow-y-auto p-2 flex flex-row md:flex-col gap-2.5 items-center md:items-stretch"></div>
    </aside>

    <!-- Center: Main Slide Canvas & Inspector -->
    <main class="flex-1 overflow-y-auto p-3 sm:p-6 flex flex-col items-center">
      <!-- Slide Canvas Wrap -->
      <div id="canvasWrap" class="w-full max-w-4xl flex justify-center">
        <div id="canvas" class="shadow-2xl rounded-2xl overflow-hidden transition-all duration-200 border border-slate-200/60"></div>
      </div>

      <!-- Navigation & Zoom Bar -->
      <div class="w-full max-w-4xl mt-3.5 flex items-center justify-between text-xs text-slate-500 px-1 flex-wrap gap-2">
        <div class="flex items-center gap-2">
          <button class="btn btn-s !py-1 text-xs" onclick="go(-1)">&larr; Prev</button>
          <span id="slideNo" class="font-bold text-slate-700 px-1">Slide 1 / 1</span>
          <button class="btn btn-s !py-1 text-xs" onclick="go(1)">Next &rarr;</button>
        </div>
        <div class="flex items-center gap-2">
          <span>Zoom</span>
          <input type="range" min="50" max="100" value="100" id="zoom" oninput="renderCanvas()" class="w-24 accent-indigo-600">
        </div>
      </div>

      <!-- Slide Editing Panel -->
      <div class="w-full max-w-4xl mt-4 card p-5 sm:p-6 space-y-4 shadow-sm border-slate-200">
        <div class="flex items-center justify-between border-b pb-3">
          <h4 class="text-sm font-bold text-slate-900 flex items-center gap-2">
            ✏️ Slide Content &amp; Visual Layout
          </h4>
          <span class="text-[11px] text-slate-400">All changes reflect live in canvas &amp; PPTX</span>
        </div>

        <div class="grid md:grid-cols-2 gap-4">
          <!-- Slide Title -->
          <div class="md:col-span-2">
            <label class="text-xs font-bold text-slate-700 uppercase tracking-wider block mb-1">Slide Heading</label>
            <input id="e_title" class="inp font-semibold text-sm" oninput="edit('title',this.value)">
          </div>

          <!-- Bullet Points / Cards Content -->
          <div class="md:col-span-2">
            <div class="flex items-center justify-between mb-1">
              <label class="text-xs font-bold text-slate-700 uppercase tracking-wider">Content Points / Items (one per line)</label>
              <span class="text-[11px] text-slate-400">Items adapt to chosen layout</span>
            </div>
            <textarea id="e_bullets" rows="5" class="inp font-normal text-xs sm:text-sm leading-relaxed" oninput="edit('bullets',this.value)"></textarea>
          </div>

          <!-- Layout Picker -->
          <div>
            <label class="text-xs font-bold text-slate-700 uppercase tracking-wider block mb-1">Visual Layout Library</label>
            <select id="e_layout" class="inp font-semibold text-xs" onchange="edit('layout',this.value)">
              <option value="title">Title Slide</option>
              <option value="bullets">Clean Bullets with Accent</option>
              <option value="cards">Aesthetic Pastel Cards Grid</option>
              <option value="image">Left Bullets + Right Image</option>
              <option value="stats">High-Impact Key Statistic</option>
              <option value="flowchart">Process &amp; Steps Flowchart</option>
              <option value="hub_arrows">Hub &amp; Connected Cards</option>
              <option value="notebook">Spiral Notebook Cards</option>
              <option value="sticky_quad">Sticky Note Quadrant Board</option>
              <option value="pushpin">Pushpin &amp; Quote Note</option>
              <option value="pennant">Pennant / Ribbon Columns</option>
              <option value="sources">Sources &amp; References Slide</option>
            </select>
          </div>

          <!-- Font Size Slider -->
          <div>
            <label class="text-xs font-bold text-slate-700 uppercase tracking-wider block mb-1">
              Font Size: <span id="fsv" class="text-indigo-600 font-bold">24px</span>
            </label>
            <input id="e_fs" type="range" min="16" max="36" value="24" class="w-full accent-indigo-600 mt-2" oninput="edit('fontSize',+this.value)">
          </div>

          <!-- Image Search & Upload -->
          <div class="md:col-span-2">
            <label class="text-xs font-bold text-slate-700 uppercase tracking-wider block mb-1">Slide Image (Wikipedia / Upload)</label>
            <div class="flex gap-2">
              <input id="e_iq" class="inp text-xs" placeholder="Search keywords on Wikipedia...">
              <button class="btn btn-s text-xs" onclick="searchImg()">🔍 Search</button>
              <label class="btn btn-s text-xs cursor-pointer">
                📤 Upload
                <input type="file" accept="image/*" class="hidden" onchange="uploadImg(this)">
              </label>
              <button class="btn btn-s text-xs text-red-600" onclick="edit('image_url','');edit('layout','bullets');renderPanel()">✕ Remove</button>
            </div>
            <div id="imgResults" class="flex gap-2 flex-wrap mt-2.5"></div>
          </div>

          <!-- Speaker Notes -->
          <div class="md:col-span-2">
            <label class="text-xs font-bold text-slate-700 uppercase tracking-wider block mb-1">Speaker Notes</label>
            <textarea id="e_notes" rows="2" class="inp text-xs" placeholder="Talking points for this slide..." oninput="edit('notes',this.value)"></textarea>
          </div>
        </div>
      </div>
    </main>

    <!-- Right: AI Assistant & Productivity Tools -->
    <aside class="w-full md:w-80 bg-white border-t md:border-t-0 md:border-l border-slate-200/80 overflow-y-auto p-4 sm:p-5 space-y-6 shrink-0 z-10">
      <!-- Per-Slide AI Assistant -->
      <div>
        <div class="flex items-center justify-between mb-2">
          <h4 class="text-xs font-extrabold text-slate-900 uppercase tracking-wider flex items-center gap-1.5">
            ✨ AI Assistant
          </h4>
          <span class="text-[10px] px-2 py-0.5 rounded-full bg-indigo-50 text-indigo-700 font-bold">Current Slide</span>
        </div>

        <div class="flex flex-wrap gap-1.5 mb-2.5" id="quick"></div>

        <div class="relative">
          <textarea id="aiCmd" rows="2" class="inp !pr-9 text-xs" placeholder="Ask AI: e.g. Make concise, rewrite heading, add metrics..." onkeydown="if(event.key==='Enter'&&!event.shiftKey){event.preventDefault();runAssist(this.value)}"></textarea>
          <button id="ai_mic" class="absolute top-2.5 right-2 text-slate-400 hover:text-red-600" title="Voice Input" onclick="toggleVoice('aiCmd', 'ai_mic')">🎤</button>
        </div>

        <button class="btn btn-p w-full !py-2 text-xs font-bold mt-2" onclick="runAssist($('aiCmd').value)">
          Transform This Slide ✦
        </button>

        <button class="btn btn-s w-full !py-1.5 text-xs mt-2" onclick="factCheckUI(S.cur)">
          ✅ Fact-Check Slide
        </button>

        <div id="aiPreview" class="mt-3"></div>
      </div>

      <!-- Quick Deck Operations -->
      <div class="border-t border-slate-100 pt-4">
        <h4 class="text-xs font-extrabold text-slate-900 uppercase tracking-wider mb-2.5">
          🛠 Presentation Tools
        </h4>
        <div class="grid grid-cols-2 gap-2">
          <button class="btn btn-s !py-1.5 text-xs" onclick="factCheckAll()">Check All Facts</button>
          <button class="btn btn-s !py-1.5 text-xs" onclick="allNotes()">View Notes</button>
          <button class="btn btn-s !py-1.5 text-xs" onclick="extra('summary')">Deck Summary</button>
          <button class="btn btn-s !py-1.5 text-xs" onclick="extra('quiz')">Generate Quiz</button>
        </div>
        <p id="stats" class="text-[11px] text-slate-400 mt-3 font-medium text-center"></p>
      </div>
    </aside>
  </div>
</section>

<script>
const $ = id => document.getElementById(id);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));

const THEMES = {
  modern:{bg:'#0f172a',fg:'#f8fafc',accent:'#38bdf8',font:'Plus Jakarta Sans, sans-serif'},
  light:{bg:'#ffffff',fg:'#0f172a',accent:'#4f46e5',font:'Plus Jakarta Sans, sans-serif'},
  corporate:{bg:'#0b2545',fg:'#ffffff',accent:'#f5a623',font:'Plus Jakarta Sans, sans-serif'},
  minimalist:{bg:'#f8fafc',fg:'#1e293b',accent:'#475569',font:'Plus Jakarta Sans, sans-serif'},
  creative:{bg:'#3b0764',fg:'#fdf4ff',accent:'#f472b6',font:'Plus Jakarta Sans, sans-serif'},
  academic:{bg:'#fff8e7',fg:'#2b2118',accent:'#9a3412',font:'Georgia, serif'}
};

const QUICK = [
  'Make more professional',
  'Add 4 impactful points',
  'Explain in simple language',
  'Convert to cards layout',
  'Make punchy & concise',
  'Add real-world statistics'
];

let S = {
  id: null,
  title: '',
  subtitle: '',
  settings: {},
  design: null,
  outline: [],
  slides: [],
  cur: 0,
  theme: 'topic',
  versions: [],
  titleOptions: [],
  subtopics: []
};

let hist = [], hpos = -1, saveTimer = null, commitTimer = null, pending = null;
let currentSource = null;
let recognition = null;
let isRecording = false;

/* ---------- Utilities ---------- */
function toast(msg, err){
  const t = $('toast');
  t.textContent = msg;
  t.className = 'fixed top-5 right-5 z-[100] max-w-sm px-4 py-3 rounded-2xl text-white text-xs sm:text-sm font-semibold shadow-2xl flex items-center gap-2 ' + (err ? 'bg-rose-600' : 'bg-indigo-600');
  clearTimeout(t._t);
  t._t = setTimeout(() => t.classList.add('hidden'), 5000);
}
function busy(msg){ $('busyMsg').textContent = msg || 'Working...'; $('busy').classList.remove('hidden'); }
function unbusy(){ $('busy').classList.add('hidden'); }
function openModal(html){ $('modalBody').innerHTML = html; $('modal').classList.remove('hidden'); }
function closeModal(){ $('modal').classList.add('hidden'); }

async function post(path, body){
  const r = await fetch(path, {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(body)});
  if(!r.ok){
    let m = r.statusText;
    try{ m = (await r.json()).detail || m; }catch(e){}
    throw new Error(m);
  }
  return r;
}
async function J(path, body){ return (await post(path, body)).json(); }

function show(v){
  ['home','setup','titles','outline','generating','editor'].forEach(x => $(x).classList.toggle('hidden', x !== v));
  if(v === 'home') renderHome();
}

const loadAll = () => { try{ return JSON.parse(localStorage.getItem('ss_decks') || '[]'); }catch(e){ return []; } };
const saveAll = a => localStorage.setItem('ss_decks', JSON.stringify(a));
const lum = h => {
  if(!h) return 0.9;
  h = h.replace('#','');
  if(h.length < 6) return 0.9;
  const [r,g,b] = [0,2,4].map(i => parseInt(h.substr(i,2),16)/255);
  return .2126*r + .7152*g + .0722*b;
};

/* ---------- Speech / Microphone Input ---------- */
function toggleVoice(targetInputId, micButtonId){
  const SpeechRec = window.SpeechRecognition || window.webkitSpeechRecognition;
  if(!SpeechRec){
    return alert('Speech recognition is not supported in this browser. Please use Chrome, Edge, or Safari.');
  }
  const btn = $(micButtonId);
  const inp = $(targetInputId);

  if(isRecording){
    recognition && recognition.stop();
    isRecording = false;
    btn.classList.remove('pulse-mic', 'text-red-600');
    return;
  }

  recognition = new SpeechRec();
  recognition.continuous = false;
  recognition.interimResults = true;
  // Support both English and Hindi voice input
  const langSel = ($('f_language') && $('f_language').value === 'Hindi') ? 'hi-IN' : 'en-US';
  recognition.lang = langSel;

  recognition.onstart = () => {
    isRecording = true;
    btn.classList.add('pulse-mic', 'text-red-600');
    toast('Listening... Speak now (' + langSel + ')');
  };

  recognition.onresult = event => {
    let transcript = '';
    for(let i = 0; i < event.results.length; i++){
      transcript += event.results[i][0].transcript;
    }
    inp.value = transcript;
    if(targetInputId === 'hero_topic' && $('f_topic')) $('f_topic').value = transcript;
  };

  recognition.onerror = e => {
    isRecording = false;
    btn.classList.remove('pulse-mic', 'text-red-600');
    toast('Voice recognition error: ' + e.error, true);
  };

  recognition.onend = () => {
    isRecording = false;
    btn.classList.remove('pulse-mic', 'text-red-600');
  };

  recognition.start();
}

/* ---------- Source Ingestion Modal ---------- */
function openSourceModal(){
  openModal(`
    <div class="space-y-5">
      <div class="border-b pb-3">
        <h3 class="text-xl font-extrabold text-slate-900">Attach Knowledge Source</h3>
        <p class="text-xs text-slate-500 mt-1">SmartSlide AI extracts text from files, YouTube videos, or web articles to base slides on.</p>
      </div>

      <!-- URL Input -->
      <div>
        <label class="text-xs font-bold text-slate-700 uppercase tracking-wider block mb-1">Web Page URL or YouTube Link</label>
        <div class="flex gap-2">
          <input id="src_url" class="inp text-xs" placeholder="https://en.wikipedia.org/... or https://youtube.com/watch?v=...">
          <button class="btn btn-p text-xs" onclick="fetchSourceUrl()">Fetch Text</button>
        </div>
      </div>

      <div class="relative flex py-1 items-center">
        <div class="flex-grow border-t border-slate-200"></div>
        <span class="flex-shrink mx-3 text-xs font-bold text-slate-400 uppercase">OR Upload Document</span>
        <div class="flex-grow border-t border-slate-200"></div>
      </div>

      <!-- File Input -->
      <div>
        <label class="text-xs font-bold text-slate-700 uppercase tracking-wider block mb-1">Upload PDF, Word (.docx) or Text file</label>
        <input type="file" id="src_file" accept=".pdf,.docx,.txt,.md" class="inp text-xs" onchange="uploadSourceFile(this)">
      </div>

      <div id="sourcePreview" class="hidden p-3.5 rounded-2xl bg-slate-50 border border-slate-200 text-xs">
        <div class="flex justify-between items-center font-bold text-slate-800 mb-1">
          <span id="srcTitle"></span>
          <span id="srcWords" class="text-indigo-600"></span>
        </div>
        <p id="srcSnippet" class="text-slate-500 italic text-[11px]"></p>
      </div>

      <div class="flex justify-end gap-2 pt-2 border-t">
        <button class="btn btn-s text-xs" onclick="closeModal()">Close</button>
        <button id="srcApplyBtn" class="btn btn-p text-xs hidden" onclick="applySource()">Use this Source</button>
      </div>
    </div>
  `);
}

async function fetchSourceUrl(){
  const url = $('src_url').value.trim();
  if(!url) return toast('Please enter a URL', true);
  busy('Extracting article/transcript...');
  try {
    const fd = new FormData();
    fd.append('url', url);
    const r = await fetch('/api/extract_source', {method:'POST', body:fd});
    if(!r.ok) throw new Error((await r.json()).detail || 'Failed to extract');
    const data = await r.json();
    displayExtractedSource(data);
  } catch(e) {
    toast(e.message, true);
  } finally {
    unbusy();
  }
}

async function uploadSourceFile(input){
  const file = input.files[0];
  if(!file) return;
  busy('Reading and extracting ' + file.name + '...');
  try {
    const fd = new FormData();
    fd.append('file', file);
    const r = await fetch('/api/extract_source', {method:'POST', body:fd});
    if(!r.ok) throw new Error((await r.json()).detail || 'Failed to extract');
    const data = await r.json();
    displayExtractedSource(data);
  } catch(e) {
    toast(e.message, true);
  } finally {
    unbusy();
  }
}

function displayExtractedSource(data){
  currentSource = data;
  $('srcTitle').textContent = '📎 ' + data.source;
  $('srcWords').textContent = data.word_count + ' words extracted';
  $('srcSnippet').textContent = '"' + data.summary + '"';
  $('sourcePreview').classList.remove('hidden');
  $('srcApplyBtn').classList.remove('hidden');
  toast('Source content successfully extracted!');
}

function applySource(){
  if(!currentSource) return;
  // Automatically fill topic if empty
  if(!$('hero_topic').value.trim() && currentSource.source){
    $('hero_topic').value = currentSource.source.replace(/^https?:\/\//, '').substring(0, 50);
  }
  if($('f_topic') && !$('f_topic').value.trim() && currentSource.source){
    $('f_topic').value = currentSource.source.replace(/^https?:\/\//, '').substring(0, 50);
  }
  $('heroSourceTag').textContent = '📎 ' + currentSource.source;
  $('heroSourceTag').classList.remove('hidden');
  if($('setupSourceCard')){
    $('setupSourceName').textContent = currentSource.source;
    $('setupSourceCard').classList.remove('hidden');
  }
  closeModal();
  toast('Source attached to presentation!');
}

function clearAttachedSource(){
  currentSource = null;
  if($('heroSourceTag')) $('heroSourceTag').classList.add('hidden');
  if($('setupSourceCard')) $('setupSourceCard').classList.add('hidden');
  toast('Attached source removed');
}

/* ---------- Topic Aesthetic Background Generator ---------- */
function getDecorativeSVG(type, accent){
  const col = accent || '#4f46e5';
  if(type === 'space'){
    return `
      <svg class="absolute inset-0 w-full h-full pointer-events-none opacity-[0.14]" xmlns="http://www.w3.org/2000/svg">
        <circle cx="90" cy="80" r="2.5" fill="${col}"/>
        <circle cx="850" cy="110" r="3" fill="${col}"/>
        <circle cx="780" cy="460" r="2" fill="${col}"/>
        <circle cx="200" cy="490" r="2.5" fill="${col}"/>
        <circle cx="900" cy="70" r="38" fill="none" stroke="${col}" stroke-width="1.5" stroke-dasharray="4 4"/>
        <circle cx="900" cy="70" r="26" fill="${col}" opacity="0.3"/>
        <path d="M40,20 Q60,60 120,40" fill="none" stroke="${col}" stroke-width="1"/>
      </svg>
    `;
  } else if(type === 'nature'){
    return `
      <svg class="absolute inset-0 w-full h-full pointer-events-none opacity-[0.15]" xmlns="http://www.w3.org/2000/svg">
        <path d="M-20,400 Q140,360 160,540" fill="none" stroke="${col}" stroke-width="2"/>
        <path d="M960,100 Q840,160 880,320" fill="none" stroke="${col}" stroke-width="2"/>
        <circle cx="160" cy="420" r="14" fill="${col}" opacity="0.4"/>
        <circle cx="860" cy="180" r="18" fill="${col}" opacity="0.4"/>
      </svg>
    `;
  } else if(type === 'business'){
    return `
      <svg class="absolute inset-0 w-full h-full pointer-events-none opacity-[0.12]" xmlns="http://www.w3.org/2000/svg">
        <defs>
          <pattern id="grid" width="40" height="40" patternUnits="userSpaceOnUse">
            <path d="M 40 0 L 0 0 0 40" fill="none" stroke="${col}" stroke-width="0.75"/>
          </pattern>
        </defs>
        <rect width="100%" height="100%" fill="url(#grid)" />
      </svg>
    `;
  } else {
    // Modern tech / abstract nodes
    return `
      <svg class="absolute inset-0 w-full h-full pointer-events-none opacity-[0.14]" xmlns="http://www.w3.org/2000/svg">
        <circle cx="880" cy="90" r="6" fill="${col}"/>
        <circle cx="820" cy="140" r="4" fill="${col}"/>
        <line x1="880" y1="90" x2="820" y2="140" stroke="${col}" stroke-width="1.5" stroke-dasharray="3 3"/>
        <circle cx="90" cy="480" r="6" fill="${col}"/>
        <circle cx="150" cy="450" r="4" fill="${col}"/>
        <line x1="90" y1="480" x2="150" y2="450" stroke="${col}" stroke-width="1.5" stroke-dasharray="3 3"/>
        <circle cx="480" cy="20" r="80" fill="${col}" opacity="0.04"/>
      </svg>
    `;
  }
}

/* ---------- Comprehensive Slide Layout Renderer ---------- */
function slideHTML(s, tn){
  const d = S.design || {
    bg_start: '#F8FAFC',
    bg_end: '#EEF2FF',
    accent: '#4F46E5',
    card_bg: 'rgba(255,255,255,0.92)',
    text_primary: '#0F172A',
    text_secondary: '#475569',
    decorative: 'tech'
  };

  let bgCss, fg, accent, cardBg;
  if(tn === 'topic'){
    bgCss = `background: linear-gradient(135deg, ${d.bg_start} 0%, ${d.bg_end} 100%);`;
    fg = d.text_primary;
    accent = d.accent;
    cardBg = d.card_bg;
  } else {
    const t = THEMES[tn] || THEMES.light;
    bgCss = `background: ${t.bg};`;
    fg = t.fg;
    accent = t.accent;
    cardBg = (lum(t.bg) > 0.55) ? 'rgba(255,255,255,0.95)' : 'rgba(30,41,59,0.85)';
  }

  // Override per slide background if manually chosen
  if(s.bg){
    bgCss = `background: ${s.bg};`;
    fg = lum(s.bg) > 0.55 ? '#0F172A' : '#F8FAFC';
    cardBg = lum(s.bg) > 0.55 ? 'rgba(255,255,255,0.95)' : 'rgba(30,41,59,0.85)';
  }

  const fs = s.fontSize || 24;
  const lay = s.layout || 'bullets';
  const bl = s.bullets || [];
  const svgWatermark = (tn === 'topic') ? getDecorativeSVG(d.decorative, accent) : '';

  const baseStyle = `width:960px;height:540px;${bgCss}color:${fg};font-family:Plus Jakarta Sans, sans-serif;box-sizing:border-box;position:relative;overflow:hidden;`;

  // 1. Title Slide Layout
  if(lay === 'title'){
    return `
      <div style="${baseStyle}display:flex;flex-direction:column;justify-content:center;padding:70px 80px">
        ${svgWatermark}
        <div style="z-index:2;display:flex;align-items:center;gap:12px;margin-bottom:24px">
          <span style="background:${accent};color:#ffffff;font-size:12px;font-weight:800;letter-spacing:1px;padding:6px 16px;border-radius:20px;text-transform:uppercase">Presentation</span>
          <span style="height:2px;width:40px;background:${accent};opacity:0.6"></span>
        </div>
        <div style="z-index:2;font-size:50px;font-weight:800;line-height:1.15;letter-spacing:-0.5px">${esc(s.title)}</div>
        <div style="z-index:2;font-size:22px;opacity:0.85;margin-top:20px;font-weight:500">${esc(bl[0] || S.subtitle || '')}</div>
      </div>
    `;
  }

  // Common Header for content slides
  const head = `
    <div style="z-index:2;margin-bottom:20px">
      <div style="display:flex;align-items:center;gap:8px">
        <span style="width:8px;height:8px;border-radius:50%;background:${accent}"></span>
        <span style="font-size:12px;font-weight:700;color:${accent};text-transform:uppercase;letter-spacing:0.8px">${esc(S.title || 'SmartSlide')}</span>
      </div>
      <div style="font-size:34px;font-weight:800;line-height:1.2;margin-top:4px;letter-spacing:-0.4px">${esc(s.title)}</div>
    </div>
  `;

  // 2. Cards Grid Layout (2x2 or 3 columns)
  if(lay === 'cards'){
    const cols = (bl.length === 3) ? 'repeat(3, 1fr)' : 'repeat(2, 1fr)';
    const cardItems = bl.slice(0, 6).map((b, i) => `
      <div style="background:${cardBg};border:1.5px solid ${accent}33;border-radius:18px;padding:18px 20px;box-shadow:0 4px 14px rgba(0,0,0,0.03);display:flex;flex-direction:column;justify-content:space-between">
        <div style="display:flex;align-items:center;gap:10px;margin-bottom:8px">
          <span style="width:28px;height:28px;border-radius:50%;background:${accent};color:#ffffff;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:800">${i + 1}</span>
          <span style="height:1px;flex:1;background:${accent}22"></span>
        </div>
        <div style="font-size:${Math.min(fs, 20)}px;line-height:1.4;font-weight:500">${esc(b)}</div>
      </div>
    `).join('');

    return `
      <div style="${baseStyle}display:flex;flex-direction:column;padding:40px 50px">
        ${svgWatermark}
        ${head}
        <div style="z-index:2;display:grid;grid-template-columns:${cols};gap:16px;flex:1;min-height:0">${cardItems}</div>
      </div>
    `;
  }

  // 3. High-Impact Statistic Layout
  if(lay === 'stats'){
    const statNum = esc(s.stat_number || '85%');
    const statLbl = esc(s.stat_label || 'Key Benchmark Impact');
    const points = bl.map(b => `<li style="margin-bottom:12px">${esc(b)}</li>`).join('');

    return `
      <div style="${baseStyle}display:flex;flex-direction:column;padding:40px 50px">
        ${svgWatermark}
        ${head}
        <div style="z-index:2;display:flex;gap:30px;flex:1;align-items:center">
          <div style="flex:1;background:linear-gradient(135deg, ${accent} 0%, ${accent}dd 100%);color:#ffffff;border-radius:24px;padding:36px;box-shadow:0 10px 25px ${accent}40;text-align:center">
            <div style="font-size:68px;font-weight:800;line-height:1">${statNum}</div>
            <div style="font-size:18px;font-weight:600;margin-top:12px;opacity:0.95">${statLbl}</div>
          </div>
          <div style="flex:1.3;background:${cardBg};border:1px solid ${accent}22;border-radius:24px;padding:28px;box-shadow:0 4px 14px rgba(0,0,0,0.03)">
            <ul style="margin:0;padding-left:22px;list-style:disc;font-size:${fs}px;line-height:1.4">${points}</ul>
          </div>
        </div>
      </div>
    `;
  }

  // 4. Process Flowchart Layout
  if(lay === 'flowchart'){
    const steps = bl.slice(0, 4).map((b, i) => `
      <div style="flex:1;background:${cardBg};border:1.5px solid ${accent}33;border-radius:18px;padding:20px 16px;box-shadow:0 4px 14px rgba(0,0,0,0.03);display:flex;flex-direction:column;position:relative">
        <div style="background:${accent};color:#ffffff;font-size:11px;font-weight:800;padding:4px 10px;border-radius:12px;align-self:flex-start;margin-bottom:12px">PHASE ${i + 1}</div>
        <div style="font-size:${Math.min(fs, 18)}px;line-height:1.35;font-weight:500;flex:1">${esc(b)}</div>
      </div>
    `).join('<div style="display:flex;align-items:center;color:' + accent + ';font-size:22px;font-weight:bold">➔</div>');

    return `
      <div style="${baseStyle}display:flex;flex-direction:column;padding:40px 50px">
        ${svgWatermark}
        ${head}
        <div style="z-index:2;display:flex;gap:12px;flex:1;align-items:stretch;margin-top:10px">${steps}</div>
      </div>
    `;
  }

  // 5. Sticky Note Quadrant Layout
  if(lay === 'sticky_quad'){
    const colors = ['#fef3c7', '#dcfce7', '#e0e7ff', '#fce7f3'];
    const textCols = ['#78350f', '#064e3b', '#1e1b4b', '#831843'];
    const quads = bl.slice(0, 4).map((b, i) => `
      <div style="background:${colors[i % 4]};color:${textCols[i % 4]};border-radius:16px;padding:20px;box-shadow:0 4px 12px rgba(0,0,0,0.04);position:relative">
        <div style="width:40px;height:10px;background:rgba(255,255,255,0.6);border-radius:4px;margin-bottom:8px"></div>
        <div style="font-size:12px;font-weight:800;text-transform:uppercase;margin-bottom:6px">Quadrant 0${i + 1}</div>
        <div style="font-size:${Math.min(fs, 19)}px;line-height:1.3">${esc(b)}</div>
      </div>
    `).join('');

    return `
      <div style="${baseStyle}display:flex;flex-direction:column;padding:40px 50px">
        ${svgWatermark}
        ${head}
        <div style="z-index:2;display:grid;grid-template-columns:1fr 1fr;gap:16px;flex:1">${quads}</div>
      </div>
    `;
  }

  // 6. Hub & Arrows Layout
  if(lay === 'hub_arrows'){
    const cards = bl.slice(0, 3).map((b, i) => `
      <div style="background:${cardBg};border:1.5px solid ${accent}44;border-radius:18px;padding:16px 20px;box-shadow:0 4px 14px rgba(0,0,0,0.03);display:flex;align-items:center;gap:14px">
        <span style="font-size:20px;color:${accent}">➔</span>
        <div style="font-size:${Math.min(fs, 20)}px;line-height:1.35;font-weight:500">${esc(b)}</div>
      </div>
    `).join('');

    return `
      <div style="${baseStyle}display:flex;flex-direction:column;padding:40px 50px">
        ${svgWatermark}
        ${head}
        <div style="z-index:2;display:flex;gap:24px;flex:1;align-items:center">
          <div style="flex:1;background:${cardBg};border:2px dashed ${accent};border-radius:24px;padding:30px;text-align:center;box-shadow:0 4px 14px rgba(0,0,0,0.03)">
            <div style="width:50px;height:50px;border-radius:50%;background:${accent}22;color:${accent};display:flex;align-items:center;justify-content:center;font-size:24px;margin:0 auto 12px">✦</div>
            <div style="font-size:22px;font-weight:800">${esc(s.title)}</div>
            <div style="font-size:12px;color:${accent};font-weight:700;margin-top:6px;text-transform:uppercase">Central Model</div>
          </div>
          <div style="flex:1.4;display:flex;flex-direction:column;gap:12px">${cards}</div>
        </div>
      </div>
    `;
  }

  // 7. Image + Bullets Layout
  if(lay === 'image' && s.image_url){
    return `
      <div style="${baseStyle}display:flex;flex-direction:column;padding:40px 50px">
        ${svgWatermark}
        ${head}
        <div style="z-index:2;display:flex;gap:30px;flex:1;align-items:stretch">
          <div style="flex:1.3;background:${cardBg};border-radius:20px;padding:24px;box-shadow:0 4px 14px rgba(0,0,0,0.03)">
            <ul style="margin:0;padding-left:20px;list-style:disc;font-size:${fs}px;line-height:1.4">
              ${bl.map(b => `<li style="margin-bottom:12px">${esc(b)}</li>`).join('')}
            </ul>
          </div>
          <div style="flex:1;border-radius:20px;overflow:hidden;box-shadow:0 8px 24px rgba(0,0,0,0.08);background:#000">
            <img src="${esc(s.image_url)}" style="width:100%;height:100%;object-fit:cover">
          </div>
        </div>
      </div>
    `;
  }

  // 8. Sources & References Slide
  if(lay === 'sources'){
    const refItems = bl.map((b, i) => `
      <div style="background:${cardBg};border-radius:14px;padding:12px 18px;border-left:4px solid ${accent};box-shadow:0 2px 8px rgba(0,0,0,0.03);display:flex;align-items:center;gap:10px">
        <span style="font-size:14px;color:${accent};font-weight:bold">[${i + 1}]</span>
        <span style="font-size:${Math.min(fs, 18)}px;font-weight:500">${esc(b)}</span>
      </div>
    `).join('');

    return `
      <div style="${baseStyle}display:flex;flex-direction:column;padding:40px 50px">
        ${svgWatermark}
        ${head}
        <div style="z-index:2;display:flex;flex-direction:column;gap:10px;flex:1">${refItems}</div>
      </div>
    `;
  }

  // Default: Clean Bullets with Card Wrap
  return `
    <div style="${baseStyle}display:flex;flex-direction:column;padding:40px 50px">
      ${svgWatermark}
      ${head}
      <div style="z-index:2;flex:1;background:${cardBg};border-radius:24px;padding:28px 36px;box-shadow:0 4px 18px rgba(0,0,0,0.03);border:1px solid ${accent}15">
        <ul style="margin:0;padding-left:24px;list-style:disc;font-size:${fs}px;line-height:1.45">
          ${bl.map(b => `<li style="margin-bottom:14px">${esc(b)}</li>`).join('')}
        </ul>
      </div>
    </div>
  `;
}

function scaled(html, w){
  return `<div style="width:${w}px;height:${w * 9 / 16}px;overflow:hidden"><div style="width:960px;height:540px;transform:scale(${w / 960});transform-origin:top left">${html}</div></div>`;
}

/* ---------- Editor Actions ---------- */
const cur = () => S.slides[S.cur];

function renderAll(){
  $('deckName').value = S.title;
  $('themeSel').value = S.theme || 'topic';
  renderThumbs();
  renderCanvas();
  renderPanel();
  updateUndo();
  updateStats();
  renderQuick();
}

function renderThumbs(){
  const tw = (window.innerWidth < 768) ? 110 : 190;
  $('thumbs').innerHTML = S.slides.map((s, i) => `
    <div class="thumb rounded-xl overflow-hidden cursor-pointer relative shrink-0 shadow-sm border border-slate-200/80 transition-all ${i === S.cur ? 'active' : ''}" onclick="selectSlide(${i})">
      ${scaled(slideHTML(s, S.theme), tw)}
      <span class="absolute top-1.5 left-1.5 bg-slate-900/80 text-white text-[10px] font-bold px-1.5 py-0.5 rounded-md shadow">${i + 1}</span>
      ${factBadge(s)}
    </div>
  `).join('');
}

function renderCanvas(){
  const wrapW = ($('canvasWrap') && $('canvasWrap').clientWidth) ? $('canvasWrap').clientWidth : (window.innerWidth - 32);
  const maxW = Math.min(880, Math.max(280, wrapW - 10));
  const w = maxW * (+$('zoom').value / 100);
  $('canvas').style.width = w + 'px';
  $('canvas').innerHTML = scaled(slideHTML(cur(), S.theme), w);
  $('slideNo').textContent = `Slide ${S.cur + 1} / ${S.slides.length}`;
}

function renderPanel(){
  const s = cur();
  if(!s) return;
  $('e_title').value = s.title;
  $('e_bullets').value = (s.bullets || []).join('\n');
  $('e_layout').value = s.layout || 'bullets';
  $('e_fs').value = s.fontSize || 24;
  $('fsv').textContent = (s.fontSize || 24) + 'px';
  $('e_iq').value = s.image_query || s.title;
  $('e_notes').value = s.notes || '';
  $('imgResults').innerHTML = '';
  $('aiPreview').innerHTML = '';
}

function renderQuick(){
  $('quick').innerHTML = QUICK.map(q => `
    <button class="text-[11px] px-2.5 py-1 rounded-full bg-indigo-50 text-indigo-700 hover:bg-indigo-100 font-medium transition-all" onclick="runAssist('${q}')">
      ${esc(q)}
    </button>
  `).join('');
}

function updateStats(){
  const words = S.slides.reduce((n, s) => n + (s.notes || '').split(/\s+/).filter(Boolean).length, 0) || S.slides.reduce((n, s) => n + (s.bullets || []).join(' ').split(/\s+/).length, 0);
  $('stats').textContent = `${S.slides.length} slides • ~${Math.max(1, Math.round(words / 130))} min speaking time`;
}

function selectSlide(i){
  S.cur = i;
  pending = null;
  renderAll();
}

function go(d){
  const j = S.cur + d;
  if(j >= 0 && j < S.slides.length) selectSlide(j);
}

function edit(field, val){
  const s = cur();
  if(!s) return;
  if(field === 'bullets'){
    s.bullets = val.split('\n').map(x => x.replace(/^[-•*\s]+/, '').trim()).filter(Boolean);
  } else {
    s[field] = val;
  }
  if(field === 'fontSize') $('fsv').textContent = val + 'px';
  renderCanvas();
  renderThumbs();
  updateStats();
  clearTimeout(commitTimer);
  commitTimer = setTimeout(commit, 400);
}

function addSlide(){
  S.slides.splice(S.cur + 1, 0, {
    title: 'New Slide',
    bullets: ['First key insight', 'Second key supporting fact', 'Third takeaway'],
    notes: '',
    layout: 'cards',
    image_query: '',
    image_url: ''
  });
  S.cur++;
  commit();
  renderAll();
}

function dupSlide(){
  S.slides.splice(S.cur + 1, 0, JSON.parse(JSON.stringify(cur())));
  S.cur++;
  commit();
  renderAll();
}

function delSlide(){
  if(S.slides.length < 2) return toast('A presentation must have at least one slide', true);
  S.slides.splice(S.cur, 1);
  S.cur = Math.min(S.cur, S.slides.length - 1);
  commit();
  renderAll();
}

function changeThemeOrDesign(val){
  S.theme = val;
  commit();
  renderAll();
}

async function regenDesignStyle(){
  busy('Generating new aesthetic background palette...');
  try {
    const res = await J('/api/design', {topic: S.title || S.settings.topic || 'Innovation'});
    S.design = res;
    S.theme = 'topic';
    commit();
    renderAll();
    toast('New aesthetic background palette applied!');
  } catch(e) {
    toast(e.message, true);
  } finally {
    unbusy();
  }
}

/* ---------- Generation Wizard & Lifecycle ---------- */
function startFromHero(){
  const t = $('hero_topic').value.trim();
  if(!t && !currentSource) return toast('Please enter a topic or attach a source', true);
  $('f_topic').value = t;
  show('setup');
}

function pickHeroIdea(idea){
  $('hero_topic').value = idea;
  if($('f_topic')) $('f_topic').value = idea;
  startFromHero();
}

function getSettings(){
  const c = $('f_count').value;
  return {
    topic: $('f_topic').value.trim(),
    count: c === 'custom' ? Math.max(2, Math.min(30, +$('f_custom').value || 8)) : +c,
    language: $('f_language').value,
    type: $('f_type').value,
    audience: $('f_audience').value,
    theme: $('f_theme').value,
    length: $('f_length').value,
    images: $('f_images').checked,
    source_material: currentSource ? currentSource.text : '',
    instructions: $('f_instr').value.trim()
  };
}

async function goTitles(){
  S.settings = getSettings();
  if(!S.settings.topic && !S.settings.source_material) return toast('Please provide a topic or source', true);
  busy('Suggesting creative titles and topics...');
  try {
    const [t, s] = await Promise.all([
      J('/api/titles', {settings:S.settings, existing:[]}),
      J('/api/suggestions', {settings:S.settings})
    ]);
    S.titleOptions = t.titles;
    S.subtopics = s.subtopics;
    renderTitles();
    show('titles');
  } catch(e) {
    toast(e.message, true);
  } finally {
    unbusy();
  }
}

function renderTitles(){
  $('titleList').innerHTML = S.titleOptions.map((t, i) => `
    <label class="flex items-center gap-3 p-3.5 rounded-2xl border border-slate-200 bg-white hover:border-indigo-500 cursor-pointer transition-all">
      <input type="radio" name="tsel" value="${i}" ${i === 0 ? 'checked' : ''} class="w-4 h-4 text-indigo-600">
      <span class="font-semibold text-sm text-slate-800">${esc(t)}</span>
    </label>
  `).join('');

  $('subList').innerHTML = S.subtopics.map((s, i) => `
    <label class="flex items-center gap-2.5 p-2.5 rounded-xl bg-slate-50 border border-slate-200 text-xs font-medium cursor-pointer">
      <input type="checkbox" checked value="${i}" class="w-3.5 h-3.5 rounded text-indigo-600 subcheck">
      <span class="truncate">${esc(s)}</span>
    </label>
  `).join('');
}

async function moreTitles(){
  busy('Suggesting more titles...');
  try {
    const t = await J('/api/titles', {settings:S.settings, existing:S.titleOptions});
    S.titleOptions = S.titleOptions.concat(t.titles);
    renderTitles();
  } catch(e) {
    toast(e.message, true);
  } finally {
    unbusy();
  }
}

async function goOutline(){
  const selIdx = (document.querySelector('input[name="tsel"]:checked') || {}).value || 0;
  S.title = S.titleOptions[selIdx] || S.settings.topic;
  S.subtitle = $('subtitle').value.trim();

  const chosenSubs = Array.from(document.querySelectorAll('.subcheck:checked')).map(el => S.subtopics[+el.value]);
  busy('Drafting slide outline...');
  try {
    const o = await J('/api/outline', {settings:S.settings, title:S.title, subtopics:chosenSubs, count:S.settings.count});
    S.outline = o.outline;
    renderOutline();
    show('outline');
  } catch(e) {
    toast(e.message, true);
  } finally {
    unbusy();
  }
}

function renderOutline(){
  $('outlineList').innerHTML = S.outline.map((o, i) => `
    <div class="flex items-center gap-2 p-2.5 rounded-xl bg-white border border-slate-200 shadow-sm">
      <span class="w-6 text-xs font-bold text-slate-400 text-center">${i + 1}</span>
      <input class="inp !py-1 text-xs font-medium flex-1 outline-inp" value="${esc(o)}">
      <button class="text-slate-400 hover:text-red-600 px-1 text-sm" onclick="delOutline(${i})">✕</button>
    </div>
  `).join('');
}

function addOutline(){
  S.outline.push('New Section');
  renderOutline();
}
function delOutline(i){
  if(S.outline.length < 2) return toast('At least one slide required', true);
  S.outline.splice(i, 1);
  renderOutline();
}
function regenOutline(){
  goOutline();
}

async function generateDeck(){
  S.outline = Array.from(document.querySelectorAll('.outline-inp')).map(el => el.value.trim()).filter(Boolean);
  show('generating');

  // Step 1: Generate Gamma AI topic background design
  $('genMsg').textContent = 'Designing topic-tailored background...';
  $('genPercent').textContent = '10%';
  $('genBar').style.width = '10%';
  try {
    const design = await J('/api/design', {topic: S.title || S.settings.topic});
    S.design = design;
    S.theme = (S.settings.theme === 'auto' || !S.settings.theme) ? 'topic' : S.settings.theme;
  } catch(e) {
    console.error('Design fallback:', e);
  }

  S.slides = [];
  const total = S.outline.length;

  for(let i = 0; i < total; i++){
    const pct = Math.round(15 + ((i + 1) / total) * 80);
    $('genPercent').textContent = pct + '%';
    $('genBar').style.width = pct + '%';
    $('genMsg').textContent = `Writing slide ${i + 1} of ${total}: "${S.outline[i]}"...`;

    try {
      const slide = await J('/api/slide', {
        settings: S.settings,
        title: S.title,
        heading: S.outline[i],
        index: i + 1,
        total: total,
        outline: S.outline,
        length: S.settings.length
      });

      // Optionally fetch Wikipedia image
      if(S.settings.images && slide.image_query && slide.layout !== 'title'){
        try {
          const imgRes = await J('/api/images?q=' + encodeURIComponent(slide.image_query));
          if(imgRes.urls && imgRes.urls.length > 0){
            slide.image_url = imgRes.urls[0];
            if(slide.layout === 'bullets') slide.layout = 'image';
          }
        } catch(imgErr){}
      }

      S.slides.push(slide);
    } catch(err) {
      toast('Slide ' + (i + 1) + ' generation warning: ' + err.message, true);
      S.slides.push({
        title: S.outline[i],
        bullets: ['Overview of ' + S.outline[i], 'Key analytical perspective'],
        notes: '',
        layout: 'cards'
      });
    }
  }

  S.cur = 0;
  S.id = 'd' + Date.now();
  hist = [];
  hpos = -1;
  commit();
  flushSave();
  show('editor');
  renderAll();
  toast('Presentation ready! Explore and edit your slides.');
}

/* ---------- Per-Slide AI Assistant ---------- */
async function runAssist(instr){
  if(!instr || !instr.trim()) return toast('Please enter an instruction', true);
  busy('Applying AI transformation...');
  try {
    const updated = await J('/api/assist', {settings:S.settings, slide:cur(), instruction:instr.trim()});
    Object.assign(cur(), updated);
    commit();
    renderAll();
    $('aiCmd').value = '';
    toast('Slide updated successfully!');
  } catch(e) {
    toast(e.message, true);
  } finally {
    unbusy();
  }
}

/* ---------- Wikipedia Image Search & Upload ---------- */
async function searchImg(){
  const q = $('e_iq').value.trim();
  if(!q) return toast('Please enter search keywords', true);
  busy('Searching images on Wikipedia...');
  try {
    const r = await (await fetch('/api/images?q=' + encodeURIComponent(q))).json();
    $('imgResults').innerHTML = (r.urls && r.urls.length) ? r.urls.map(u => `
      <img src="${esc(u)}" class="h-16 w-20 object-cover rounded-xl cursor-pointer border-2 border-transparent hover:border-indigo-600 transition-all shadow-sm" onclick="pickImg(this.src)">
    `).join('') : '<span class="text-xs text-slate-400">No images found. Try other keywords or upload a photo.</span>';
  } catch(e) {
    toast('Image search failed', true);
  } finally {
    unbusy();
  }
}

function pickImg(url){
  edit('image_url', url);
  edit('layout', 'image');
  renderPanel();
  toast('Image added to slide!');
}

function uploadImg(input){
  const file = input.files[0];
  if(!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    pickImg(e.target.result);
  };
  reader.readAsDataURL(file);
}

/* ---------- Fact-Checking System ---------- */
const factStatus = s => (s._facts || {}).status || 'none';
const factBadge = s => {
  const st = factStatus(s);
  if(st === 'correct') return '<span class="absolute top-1.5 right-1.5 bg-emerald-500 text-white text-[9px] font-bold px-1.5 py-0.5 rounded shadow">✓ verified</span>';
  if(st === 'wrong') return '<span class="absolute top-1.5 right-1.5 bg-rose-500 text-white text-[9px] font-bold px-1.5 py-0.5 rounded shadow">! warning</span>';
  return '';
};

async function factCheckUI(idx){
  const s = S.slides[idx];
  if(!s) return;
  busy('Verifying facts against Wikipedia sources...');
  try {
    const res = await J('/api/factcheck', {title:s.title, bullets:s.bullets});
    s._facts = {
      status: res.checks.some(c => c.verdict === 'wrong') ? 'wrong' : 'correct',
      checks: res.checks,
      sources: res.sources
    };
    renderThumbs();
    openFactModal(s);
  } catch(e) {
    toast(e.message, true);
  } finally {
    unbusy();
  }
}

function openFactModal(s){
  const f = s._facts || {};
  const items = (f.checks || []).map(c => `
    <div class="p-3 rounded-2xl border ${c.verdict==='wrong'?'bg-rose-50 border-rose-200 text-rose-950':'bg-emerald-50 border-emerald-200 text-emerald-950'} text-xs space-y-1">
      <div class="font-bold flex items-center justify-between">
        <span>${esc(c.bullet)}</span>
        <span class="uppercase text-[10px] px-2 py-0.5 rounded-full font-extrabold ${c.verdict==='wrong'?'bg-rose-200 text-rose-800':'bg-emerald-200 text-emerald-800'}">${c.verdict}</span>
      </div>
      ${c.issue ? `<div class="text-[11px] text-slate-600"><strong>Note:</strong> ${esc(c.issue)}</div>` : ''}
      ${c.correction ? `<div class="text-[11px] text-indigo-700"><strong>Suggested Fix:</strong> ${esc(c.correction)}</div>` : ''}
    </div>
  `).join('');

  openModal(`
    <div class="space-y-4">
      <div class="border-b pb-3">
        <h3 class="text-lg font-bold text-slate-900">Fact-Check Report</h3>
        <p class="text-xs text-slate-500">Slide: "${esc(s.title)}"</p>
      </div>
      <div class="space-y-2 max-h-96 overflow-y-auto">${items || '<p class="text-xs text-slate-400">No bullets to check.</p>'}</div>
      <div class="flex justify-end pt-2 border-t">
        <button class="btn btn-s text-xs" onclick="closeModal()">Close</button>
      </div>
    </div>
  `);
}

async function factCheckAll(){
  busy('Verifying all presentation slides...');
  for(let i = 0; i < S.slides.length; i++){
    try {
      const res = await J('/api/factcheck', {title:S.slides[i].title, bullets:S.slides[i].bullets});
      S.slides[i]._facts = {
        status: res.checks.some(c => c.verdict === 'wrong') ? 'wrong' : 'correct',
        checks: res.checks
      };
    } catch(e){}
  }
  unbusy();
  renderThumbs();
  toast('Fact check completed across all slides!');
}

/* ---------- Presentation Mode ---------- */
let pIdx = 0;
let presTimerInterval = null;
let presStartTime = 0;

function startPresent(){
  pIdx = S.cur;
  drawPresent();
  $('present').classList.remove('hidden');
  presStartTime = Date.now();
  clearInterval(presTimerInterval);
  presTimerInterval = setInterval(updatePresTimer, 1000);
  document.documentElement.requestFullscreen && document.documentElement.requestFullscreen().catch(() => {});
}

function updatePresTimer(){
  const sec = Math.floor((Date.now() - presStartTime) / 1000);
  const m = String(Math.floor(sec / 60)).padStart(2, '0');
  const s = String(sec % 60).padStart(2, '0');
  if($('presTimer')) $('presTimer').textContent = `${m}:${s}`;
}

function drawPresent(){
  const s = S.slides[pIdx];
  const sc = Math.min((window.innerWidth - 40) / 960, (window.innerHeight - 150) / 540);
  $('presCanvas').innerHTML = scaled(slideHTML(s, S.theme), 960 * sc);
  $('presCounter').textContent = `Slide ${pIdx + 1} of ${S.slides.length}`;
  $('presNotesText').textContent = s.notes || 'No speaker notes for this slide.';
}

function presNext(){
  if(pIdx < S.slides.length - 1){
    pIdx++;
    drawPresent();
  } else {
    endPresent();
  }
}

function presPrev(){
  if(pIdx > 0){
    pIdx--;
    drawPresent();
  }
}

function endPresent(){
  $('present').classList.add('hidden');
  clearInterval(presTimerInterval);
  if(document.fullscreenElement) document.exitFullscreen().catch(() => {});
}

document.addEventListener('keydown', e => {
  if($('present').classList.contains('hidden')) return;
  if(e.key === 'ArrowRight' || e.key === ' ') presNext();
  else if(e.key === 'ArrowLeft') presPrev();
  else if(e.key === 'Escape') endPresent();
});

/* ---------- Storage, History & Export ---------- */
function commit(){
  hist = hist.slice(0, hpos + 1);
  hist.push(JSON.stringify({slides:S.slides, theme:S.theme, title:S.title, design:S.design}));
  if(hist.length > 30) hist.shift();
  hpos = hist.length - 1;
  updateUndo();
  scheduleSave();
}

function undo(){
  if(hpos > 0){
    hpos--;
    const st = JSON.parse(hist[hpos]);
    S.slides = st.slides;
    S.theme = st.theme;
    S.title = st.title;
    S.design = st.design;
    renderAll();
  }
}

function redo(){
  if(hpos < hist.length - 1){
    hpos++;
    const st = JSON.parse(hist[hpos]);
    S.slides = st.slides;
    S.theme = st.theme;
    S.title = st.title;
    S.design = st.design;
    renderAll();
  }
}

function updateUndo(){
  if($('undoBtn')) $('undoBtn').disabled = (hpos <= 0);
  if($('redoBtn')) $('redoBtn').disabled = (hpos >= hist.length - 1);
}

function scheduleSave(){
  $('saveStatus').textContent = 'Saving...';
  clearTimeout(saveTimer);
  saveTimer = setTimeout(flushSave, 800);
}

function flushSave(){
  const all = loadAll();
  const deckData = {
    id: S.id || ('d' + Date.now()),
    title: S.title,
    subtitle: S.subtitle,
    topic: S.settings.topic,
    slides: S.slides,
    theme: S.theme,
    design: S.design,
    settings: S.settings,
    updated: Date.now()
  };
  const idx = all.findIndex(x => x.id === S.id);
  if(idx >= 0) all[idx] = deckData;
  else all.unshift(deckData);
  try {
    saveAll(all);
    $('saveStatus').textContent = 'All changes saved';
  } catch(e) {
    $('saveStatus').textContent = 'Storage full';
  }
}

function exportDeckJSON(){
  const blob = new Blob([JSON.stringify(S, null, 2)], {type:'application/json'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = (S.title || 'presentation').replace(/[^\w-]+/g, '_') + '.json';
  a.click();
  toast('Presentation JSON project downloaded!');
}

function importDeck(input){
  const file = input.files[0];
  if(!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = JSON.parse(e.target.result);
      S = Object.assign({cur:0, versions:[]}, data);
      commit();
      flushSave();
      show('editor');
      renderAll();
      toast('Presentation project imported!');
    } catch(err) {
      toast('Invalid project file: ' + err.message, true);
    }
  };
  reader.readAsText(file);
}

async function downloadPPTX(){
  busy('Building native PowerPoint file (.pptx)...');
  try {
    const res = await post('/api/export', {
      title: S.title,
      theme: S.theme,
      design: S.design,
      slides: S.slides
    });
    const blob = await res.blob();
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = (S.title || 'presentation').replace(/[^\w-]+/g, '_') + '.pptx';
    a.click();
    toast('PowerPoint downloaded successfully!');
  } catch(e) {
    toast(e.message, true);
  } finally {
    unbusy();
  }
}

function goHome(){
  flushSave();
  show('home');
}

function renderHome(){
  const q = ($('search').value || '').toLowerCase();
  const decks = loadAll().filter(d => !q || (d.title||'').toLowerCase().includes(q) || (d.topic||'').toLowerCase().includes(q));
  $('deckList').innerHTML = decks.length ? decks.map(d => `
    <div class="card overflow-hidden hover:shadow-xl transition-all border border-slate-200/80 flex flex-col justify-between">
      <div class="cursor-pointer" onclick="openDeck('${d.id}')">
        ${scaled(slideHTML(d.slides[0] || {title:d.title,bullets:[],layout:'title'}, d.theme || 'topic'), 340)}
      </div>
      <div class="p-4 border-t border-slate-100 bg-white">
        <div class="font-bold text-sm text-slate-900 truncate">${esc(d.title)}</div>
        <div class="text-[11px] text-slate-400 mt-0.5">${d.slides.length} slides • ${new Date(d.updated).toLocaleDateString()}</div>
        <div class="flex items-center gap-1.5 mt-3 flex-wrap">
          <button class="btn btn-p !py-1 !px-3 text-xs" onclick="openDeck('${d.id}')">Open</button>
          <button class="btn btn-s !py-1 text-xs text-red-600" onclick="delDeck('${d.id}')">Delete</button>
        </div>
      </div>
    </div>
  `).join('') : '<p class="text-xs text-slate-400 col-span-full py-8 text-center font-medium">No saved presentations yet. Enter a topic above to create your first presentation.</p>';
}

function openDeck(id){
  const d = loadAll().find(x => x.id === id);
  if(!d) return;
  S = Object.assign({cur:0, versions:[], design:null}, d);
  hist = [];
  hpos = -1;
  commit();
  show('editor');
  renderAll();
}

function delDeck(id){
  if(confirm('Delete this presentation?')){
    saveAll(loadAll().filter(x => x.id !== id));
    renderHome();
  }
}

/* ---------- AI Settings & Feedback Modals ---------- */
async function openSettings(){
  let c;
  try {
    c = await (await fetch('/api/config')).json();
  } catch(e) {
    return toast('Could not load AI settings', true);
  }

  openModal(`
    <div class="space-y-4">
      <div class="border-b pb-3">
        <h3 class="text-lg font-bold text-slate-900">AI Provider Settings</h3>
        <p class="text-xs text-slate-500">Your API key is saved locally in <code>settings.json</code> and never exposed.</p>
      </div>

      <div>
        <label class="text-xs font-bold text-slate-700 uppercase tracking-wider block mb-1">Provider</label>
        <select id="cf_prov" class="inp text-xs font-semibold">
          <option value="openai">Groq / OpenAI Compatible</option>
          <option value="gemini">Google Gemini</option>
          <option value="anthropic">Anthropic (Claude)</option>
          <option value="ollama">Ollama (Local)</option>
        </select>
      </div>

      <div>
        <label class="text-xs font-bold text-slate-700 uppercase tracking-wider block mb-1">Base URL</label>
        <input id="cf_url" class="inp text-xs font-mono" value="${c.base_url || 'https://api.groq.com/openai/v1'}">
      </div>

      <div>
        <label class="text-xs font-bold text-slate-700 uppercase tracking-wider block mb-1">Model Name</label>
        <input id="cf_model" class="inp text-xs font-mono" value="${c.model || 'openai/gpt-oss-120b'}">
      </div>

      <div>
        <label class="text-xs font-bold text-slate-700 uppercase tracking-wider block mb-1">API Key</label>
        <input id="cf_key" type="password" class="inp text-xs" placeholder="${c.has_key ? 'Key saved — leave blank to keep existing' : 'Enter API key'}">
      </div>

      <div class="flex justify-end gap-2 pt-2 border-t">
        <button class="btn btn-s text-xs" onclick="closeModal()">Cancel</button>
        <button class="btn btn-p text-xs" onclick="saveSettings()">Save &amp; Test Connection</button>
      </div>
    </div>
  `);
}

async function saveSettings(){
  busy('Saving settings...');
  try {
    await J('/api/config', {
      provider: $('cf_prov').value,
      base_url: $('cf_url').value.trim(),
      model: $('cf_model').value.trim(),
      api_key: $('cf_key').value.trim()
    });
    await J('/api/config/test', {});
    closeModal();
    toast('Connected successfully! Settings saved.');
  } catch(e) {
    toast(e.message, true);
  } finally {
    unbusy();
  }
}

function openFeedback(){
  openModal(`
    <div class="space-y-4">
      <div class="border-b pb-3">
        <h3 class="text-lg font-bold text-slate-900">Share Feedback</h3>
        <p class="text-xs text-slate-500">Your feedback helps improve SmartSlide AI.</p>
      </div>
      <div>
        <label class="text-xs font-bold text-slate-700 uppercase tracking-wider block mb-1">Feedback Category</label>
        <select id="fb_cat" class="inp text-xs"><option>General Suggestion</option><option>Bug Report</option><option>Design &amp; Layouts</option><option>AI Content Quality</option></select>
      </div>
      <div>
        <label class="text-xs font-bold text-slate-700 uppercase tracking-wider block mb-1">Your Message</label>
        <textarea id="fb_msg" rows="4" class="inp text-xs" placeholder="Tell us what you liked or want improved..."></textarea>
      </div>
      <div class="flex justify-end gap-2 pt-2 border-t">
        <button class="btn btn-s text-xs" onclick="closeModal()">Cancel</button>
        <button class="btn btn-p text-xs" onclick="sendFeedback()">Submit Feedback</button>
      </div>
    </div>
  `);
}

async function sendFeedback(){
  const msg = $('fb_msg').value.trim();
  if(!msg) return toast('Please enter a message', true);
  try {
    await post('/api/feedback', {message:msg, category:$('fb_cat').value, rating:5});
    closeModal();
    toast('Thank you for your feedback!');
  } catch(e) {
    toast(e.message, true);
  }
}

async function extra(kind){
  busy('Generating ' + kind + '...');
  try {
    const res = await J('/api/extra', {settings:S.settings, slides:S.slides, kind:kind});
    if(kind === 'quiz'){
      const qHtml = (res.items || []).map((it, i) => `
        <div class="p-3.5 rounded-2xl bg-slate-50 border border-slate-200 text-xs space-y-1.5">
          <div class="font-bold text-slate-900">${i + 1}. ${esc(it.q)}</div>
          <div class="grid grid-cols-2 gap-1.5 pt-1">
            ${(it.options || []).map(o => `<div class="p-2 rounded-xl bg-white border border-slate-200">${esc(o)}</div>`).join('')}
          </div>
          <div class="text-[11px] font-bold text-emerald-700 pt-1">Correct Answer: ${esc(it.answer)}</div>
        </div>
      `).join('');
      openModal(`<div class="space-y-3"><h3 class="text-lg font-bold">Interactive Deck Quiz</h3><div class="space-y-2.5 max-h-96 overflow-y-auto">${qHtml}</div><button class="btn btn-s text-xs mt-2" onclick="closeModal()">Close</button></div>`);
    } else {
      openModal(`<div class="space-y-3"><h3 class="text-lg font-bold">Presentation Summary</h3><p class="text-xs leading-relaxed text-slate-700 bg-slate-50 p-4 rounded-2xl border border-slate-200">${esc(res.summary)}</p><button class="btn btn-s text-xs mt-2" onclick="closeModal()">Close</button></div>`);
    }
  } catch(e) {
    toast(e.message, true);
  } finally {
    unbusy();
  }
}

function allNotes(){
  const html = S.slides.map((s, i) => `
    <div class="p-3 rounded-2xl bg-slate-50 border border-slate-200 text-xs">
      <div class="font-bold text-indigo-700 mb-1">Slide ${i + 1}: ${esc(s.title)}</div>
      <div class="text-slate-600">${esc(s.notes || 'No notes.')}</div>
    </div>
  `).join('');
  openModal(`<div class="space-y-3"><h3 class="text-lg font-bold">All Speaker Notes</h3><div class="space-y-2 max-h-96 overflow-y-auto">${html}</div><button class="btn btn-s text-xs mt-2" onclick="closeModal()">Close</button></div>`);
}

async function checkHealth(){
  try {
    const h = await (await fetch('/api/health')).json();
    if(!h.ready && $('health')){
      $('health').textContent = h.message;
      $('health').classList.remove('hidden');
    }
  } catch(e){}
}

addEventListener('resize', () => { if(!$('editor').classList.contains('hidden')) renderCanvas(); });
addEventListener('beforeunload', flushSave);

checkHealth();
show('home');
</script>
</body>
</html>
'''

if __name__ == "__main__":
    print("SmartSlide AI running at http://127.0.0.1:8000 (Ctrl+C to stop)")
    uvicorn.run(app, host="127.0.0.1", port=8000)
