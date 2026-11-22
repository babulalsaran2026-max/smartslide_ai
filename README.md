# SmartSlide AI

AI PPT generator: FastAPI backend + HTML/Tailwind frontend + Ollama (Llama 3.2).

## Run (VS Code terminal)

```bash
# 1. Ollama
ollama pull llama3.2
ollama serve          # skip if Ollama is already running

# 2. Project (folder with app.py and requirements.txt)
python -m venv venv
venv\Scripts\activate          # Windows
# source venv/bin/activate     # macOS / Linux
pip install -r requirements.txt
uvicorn app:app --reload --port 8000
```

Open http://127.0.0.1:8000

- Feedback submitted in the app is saved to `feedback.jsonl`; view all at http://127.0.0.1:8000/api/feedback
- Change model: `set OLLAMA_MODEL=llama3.2:3b` (Windows) / `export OLLAMA_MODEL=...`

## Use your own API (OpenAI / Gemini / Claude / any OpenAI-compatible)

1. Open the app, click **AI Settings** (home screen, top right).
2. Choose a provider, type the **model name**, paste your **API key**, click **Save & Test**.
3. The key is saved only in `settings.json` next to `app.py`. Never share that file or upload it to GitHub.

Alternative: set environment variables `SMARTSLIDE_PROVIDER` (ollama / openai / gemini / anthropic), `SMARTSLIDE_API_KEY`, `OLLAMA_MODEL`.
