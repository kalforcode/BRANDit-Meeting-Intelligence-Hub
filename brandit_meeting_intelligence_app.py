"""
BRANDit Meeting Intelligence
Groq API transcription, pyannoteAI API speaker diarization, Groq API minutes, action tracking, and Excel reporting

Run:
    streamlit run app.py --server.fileWatcherType none

Streamlit secrets / .env:
    GROQ_API_KEY=...
    PYANNOTE_API_KEY=...
    GROQ_MODEL=llama-3.3-70b-versatile
    WHISPER_API_MODEL=whisper-large-v3-turbo

Notes:
    - Transcription uses Groq Whisper API, not local Faster-Whisper.
    - Speaker diarization uses pyannoteAI hosted API, not local pyannote.audio.
    - Groq LLM creates minutes/action points.
    - FFmpeg is still required only for light audio conversion/compression before API upload.
"""

# ── Windows / OpenMP / Streamlit watcher fixes ───────────────────────────────
import os
os.environ.setdefault("KMP_DUPLICATE_LIB_OK", "TRUE")
os.environ.setdefault("STREAMLIT_SERVER_FILE_WATCHER_TYPE", "none")

import sys
if sys.platform == "win32":
    import asyncio
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

import warnings
warnings.filterwarnings("ignore", message=".*torch.classes.*")
warnings.filterwarnings("ignore", category=UserWarning)

# ── Imports ──────────────────────────────────────────────────────────────────
import html
import json
import re
import shutil
import subprocess
import tempfile
import time
import uuid
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import requests
import streamlit as st
from dotenv import load_dotenv

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

load_dotenv()

# ── Page Config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="BRANDit · Meeting Intelligence",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Styling ──────────────────────────────────────────────────────────────────
st.markdown("""
<style>
:root{
    --bg:#050814;
    --panel:#0b1020;
    --panel-2:#11182a;
    --line:rgba(255,255,255,.10);
    --text:#e8edf7;
    --muted:#94a3b8;
    --teal:#20e3c2;
    --blue:#38bdf8;
    --navy:#07111f;
}
html,body,[class*="css"]{font-family:Inter,Segoe UI,Arial,sans-serif;}
.stApp{background:radial-gradient(circle at top right,rgba(32,227,194,.09),transparent 34%),linear-gradient(180deg,#060914 0%,#050814 100%);color:var(--text);}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#07111f 0%,#090d1a 100%)!important;border-right:1px solid rgba(32,227,194,.13);}
.block-container{padding-top:1.3rem!important;max-width:1380px;}
.brand-card{background:linear-gradient(135deg,rgba(32,227,194,.16),rgba(56,189,248,.08));border:1px solid rgba(32,227,194,.22);border-radius:20px;padding:18px;margin:8px 0 20px;box-shadow:0 16px 38px rgba(0,0,0,.24);}
.brand-card h2{font-size:1.15rem;margin:0 0 8px;color:#ffffff;font-weight:850;letter-spacing:.01em;}
.brand-card p{font-size:.84rem;color:#cbd5e1;line-height:1.55;margin:0;}
.contact-strip{font-size:.75rem;color:#9fb7c8;border-top:1px solid rgba(255,255,255,.08);padding-top:12px;margin-top:12px;line-height:1.55;}
.hero{position:relative;overflow:hidden;background:linear-gradient(135deg,rgba(10,20,35,.98),rgba(9,22,38,.94) 56%,rgba(5,38,44,.90));border:1px solid rgba(32,227,194,.16);border-radius:22px;padding:34px 38px;margin-bottom:24px;box-shadow:0 18px 45px rgba(0,0,0,.24);}
.hero:before{content:'';position:absolute;inset:0;background:linear-gradient(90deg,rgba(255,255,255,.04) 1px,transparent 1px),linear-gradient(180deg,rgba(255,255,255,.025) 1px,transparent 1px);background-size:84px 84px;opacity:.18;pointer-events:none;}
.hero>*{position:relative;z-index:1;}
.eyebrow{display:inline-flex;align-items:center;gap:8px;background:rgba(32,227,194,.10);border:1px solid rgba(32,227,194,.18);color:var(--teal);border-radius:7px;padding:7px 12px;font-size:.72rem;font-weight:850;letter-spacing:.14em;text-transform:uppercase;margin-bottom:18px;}
.hero h1{font-size:2.45rem;margin:0 0 12px;color:var(--teal);font-weight:900;letter-spacing:-.04em;}
.hero p{color:#d9e4ef;margin:0;font-size:1.02rem;line-height:1.65;max-width:850px;}
.workflow{display:flex;gap:10px;align-items:center;margin:0 0 22px;color:#64748b;font-size:.72rem;text-transform:uppercase;letter-spacing:.12em;flex-wrap:wrap;}
.workflow span{border:1px solid rgba(255,255,255,.08);background:rgba(255,255,255,.035);color:#a7b6c9;border-radius:9px;padding:8px 12px;text-transform:none;letter-spacing:0;font-size:.78rem;}
.workflow .active{background:rgba(32,227,194,.16);border-color:rgba(32,227,194,.32);color:var(--teal);font-weight:800;}
.card{background:rgba(255,255,255,.035);border:1px solid rgba(255,255,255,.09);border-radius:16px;padding:20px;margin-bottom:18px;box-shadow:0 12px 32px rgba(0,0,0,.16);}
.card-title{font-size:.76rem;font-weight:850;color:var(--teal);text-transform:uppercase;letter-spacing:.12em;margin-bottom:12px;}
.feature-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin:8px 0 24px;}
.feature-card{background:rgba(255,255,255,.035);border:1px solid rgba(255,255,255,.09);border-radius:16px;padding:18px;min-height:118px;}
.feature-card h3{font-size:1rem;margin:0 0 8px;color:#f8fafc;}
.feature-card p{font-size:.9rem;color:#aab8c9;line-height:1.55;margin:0;}
.deliverable{display:flex;gap:12px;align-items:flex-start;margin-bottom:13px;padding-bottom:13px;border-bottom:1px solid rgba(255,255,255,.06);}
.deliverable:last-child{border-bottom:0;margin-bottom:0;padding-bottom:0;}
.deliverable-number{min-width:28px;height:28px;border-radius:9px;display:flex;align-items:center;justify-content:center;background:rgba(32,227,194,.12);border:1px solid rgba(32,227,194,.22);color:var(--teal);font-weight:900;font-size:.8rem;}
.deliverable b{color:#eaf3fb;}
.deliverable div:last-child{color:#97a6ba;font-size:.88rem;line-height:1.5;}
.metric-strip{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin:16px 0 24px;}
.metric-box{background:rgba(255,255,255,.035);border:1px solid rgba(255,255,255,.08);border-radius:14px;padding:16px;text-align:center;}
.metric-num{font-size:1.65rem;font-weight:850;color:var(--teal);line-height:1;}
.metric-label{font-size:.72rem;color:#8796ab;margin-top:6px;text-transform:uppercase;letter-spacing:.07em;}
.section-header{font-size:1.05rem;font-weight:850;color:#f1f5f9;margin:22px 0 12px;display:flex;align-items:center;gap:10px;}
.section-header:after{content:'';height:1px;flex:1;background:linear-gradient(90deg,rgba(32,227,194,.45),transparent);}
.transcript-box{background:#07101f;border:1px solid rgba(255,255,255,.09);border-radius:12px;padding:14px;max-height:430px;overflow-y:auto;}
.seg-row{display:grid;grid-template-columns:70px 110px 1fr;gap:10px;align-items:flex-start;padding:8px;border-bottom:1px solid rgba(255,255,255,.055);}
.seg-time{color:#7dd3fc;font-family:Consolas,monospace;font-size:.8rem;}
.seg-spk{color:var(--teal);font-weight:850;font-size:.82rem;white-space:nowrap;}
.seg-text{color:#dbe5ef;font-size:.9rem;line-height:1.55;}
.small-muted{color:#8ca0b8;font-size:.82rem;}
.status-ok{color:#34d399;font-weight:700;}
.status-bad{color:#fb7185;font-weight:700;}
.stButton>button{background:linear-gradient(135deg,#20e3c2,#12b5cb)!important;color:#031018!important;border:0!important;border-radius:10px!important;font-weight:850!important;width:100%;}
.stDownloadButton>button{background:rgba(255,255,255,.04)!important;color:#d9f7f2!important;border:1px solid rgba(32,227,194,.18)!important;border-radius:10px!important;font-weight:750!important;width:100%;}
[data-testid="stFileUploader"]{background:rgba(255,255,255,.025)!important;border:2px dashed rgba(32,227,194,.28)!important;border-radius:14px!important;padding:8px;}
.footer-bar{text-align:center;color:#7d90a8;font-size:.76rem;padding:14px 16px;margin:26px 0 8px;border:1px solid rgba(255,255,255,.08);background:rgba(255,255,255,.025);border-radius:14px;}
.footer-bar b{color:var(--teal);}
#MainMenu, footer{visibility:hidden;}
@media(max-width:900px){.metric-strip{grid-template-columns:repeat(2,1fr)}.seg-row{grid-template-columns:60px 95px 1fr}.hero h1{font-size:1.8rem}}
</style>
""", unsafe_allow_html=True)

# ── Helpers ──────────────────────────────────────────────────────────────────
def get_secret(name: str, default: Optional[str] = None) -> Optional[str]:
    """Read from Streamlit secrets first, then .env / environment variables."""
    try:
        value = st.secrets.get(name)  # type: ignore[attr-defined]
        if value:
            return str(value)
    except Exception:
        pass
    return os.getenv(name, default)


def fmt_time(seconds: float) -> str:
    seconds = int(seconds or 0)
    m, s = divmod(seconds, 60)
    h, m = divmod(m, 60)
    return f"{h:02d}:{m:02d}:{s:02d}" if h else f"{m:02d}:{s:02d}"


def check_ffmpeg() -> bool:
    return shutil.which("ffmpeg") is not None


def prepare_audio_for_apis(input_path: str, output_path: str) -> None:
    """
    Convert uploaded audio/video to 16 kHz mono FLAC.

    This keeps the Streamlit app light: FFmpeg only prepares/compresses audio;
    transcription and diarization happen through hosted APIs.
    """
    if not check_ffmpeg():
        raise RuntimeError("FFmpeg not found. Add ffmpeg to packages.txt on Streamlit Cloud.")

    cmd = [
        "ffmpeg", "-y",
        "-i", input_path,
        "-vn",
        "-map", "0:a:0",
        "-ac", "1",
        "-ar", "16000",
        "-c:a", "flac",
        output_path,
    ]
    result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if result.returncode != 0:
        raise RuntimeError("FFmpeg failed to prepare audio:\n" + result.stderr[-2000:])


# ── Groq Whisper API Transcription ───────────────────────────────────────────
GROQ_TRANSCRIPTION_URL = "https://api.groq.com/openai/v1/audio/transcriptions"


def _raise_groq_audio_error(response: requests.Response) -> None:
    if response.ok:
        return
    try:
        detail = response.json()
    except Exception:
        detail = response.text
    raise RuntimeError(f"Groq Whisper API failed ({response.status_code}): {detail}")


def _to_plain_dict(value):
    """Convert Groq/OpenAI SDK-like objects to plain dicts when needed."""
    if isinstance(value, dict):
        return value
    if hasattr(value, "model_dump"):
        return value.model_dump()
    if hasattr(value, "dict"):
        return value.dict()
    return value


def transcribe_with_groq_api(
    audio_path: str,
    groq_key: str,
    model_name: str = "whisper-large-v3-turbo",
    language: str = "en",
    prompt: str = "Meeting discussion with multiple speakers. Use clear punctuation.",
) -> dict:
    """
    Transcribe audio using Groq's OpenAI-compatible Whisper endpoint.

    Returns the same structure the rest of the app expects:
    {
        full_text,
        segments: [{start, end, text}],
        words,
        duration,
        language,
        lang_prob
    }
    """
    if not groq_key:
        raise RuntimeError("GROQ_API_KEY is missing.")

    headers = {"Authorization": f"Bearer {groq_key}"}

    data = {
        "model": model_name,
        "response_format": "verbose_json",
        "temperature": "0",
        "timestamp_granularities[]": "segment",
    }
    if prompt:
        data["prompt"] = prompt[:900]
    if language != "auto":
        data["language"] = language

    with open(audio_path, "rb") as audio_file:
        files = {
            "file": (Path(audio_path).name, audio_file, "audio/flac")
        }
        response = requests.post(
            GROQ_TRANSCRIPTION_URL,
            headers=headers,
            data=data,
            files=files,
            timeout=900,
        )

    _raise_groq_audio_error(response)
    result = _to_plain_dict(response.json())

    raw_segments = result.get("segments") or []
    segments = []
    full_parts = []

    for seg in raw_segments:
        seg = _to_plain_dict(seg)
        if not isinstance(seg, dict):
            continue
        text = (seg.get("text") or "").strip()
        if not text:
            continue
        start = float(seg.get("start", 0) or 0)
        end = float(seg.get("end", start) or start)
        if end <= start:
            # Keep the segment usable even if the provider omits/duplicates end time.
            end = start + 0.01
        segments.append({
            "start": start,
            "end": end,
            "text": text,
        })
        full_parts.append(text)

    text = (result.get("text") or "").strip()
    if not segments and text:
        segments.append({"start": 0.0, "end": float(result.get("duration", 0) or 0), "text": text})
        full_parts.append(text)

    duration = float(result.get("duration", 0) or 0)
    if not duration and segments:
        duration = float(segments[-1]["end"])

    return {
        "full_text": " ".join(full_parts).strip() or text,
        "segments": segments,
        "words": [],
        "duration": duration,
        "language": result.get("language", language),
        "lang_prob": 0,
    }


# ── pyannoteAI API Diarization ───────────────────────────────────────────────
PYANNOTE_BASE_URL = "https://api.pyannote.ai/v1"


def _pyannote_headers(api_key: str, json_content: bool = True) -> dict:
    headers = {"Authorization": f"Bearer {api_key}"}
    if json_content:
        headers["Content-Type"] = "application/json"
    return headers


def _raise_pyannote_error(response: requests.Response, action: str) -> None:
    if response.ok:
        return
    try:
        detail = response.json()
    except Exception:
        detail = response.text
    raise RuntimeError(f"pyannoteAI {action} failed ({response.status_code}): {detail}")


def upload_audio_to_pyannote_media(audio_path: str, api_key: str) -> str:
    """
    Upload local audio to pyannoteAI temporary media storage.
    Returns a media:// URL that can be used in the diarize endpoint.
    """
    extension = Path(audio_path).suffix.lower().lstrip(".") or "flac"
    object_key = f"brandit-meetings/{datetime.now().strftime('%Y%m%d')}/{uuid.uuid4().hex}.{extension}"
    media_url = f"media://{object_key}"

    create_response = requests.post(
        f"{PYANNOTE_BASE_URL}/media/input",
        headers=_pyannote_headers(api_key),
        json={"url": media_url},
        timeout=60,
    )
    _raise_pyannote_error(create_response, "media URL creation")

    upload_url = create_response.json().get("url")
    if not upload_url:
        raise RuntimeError("pyannoteAI did not return a pre-signed upload URL.")

    with open(audio_path, "rb") as audio_file:
        upload_response = requests.put(
            upload_url,
            data=audio_file,
            headers={"Content-Type": "application/octet-stream"},
            timeout=600,
        )
    _raise_pyannote_error(upload_response, "audio upload")

    return media_url


def submit_pyannote_diarization_job(
    media_url: str,
    api_key: str,
    exact_speakers: Optional[int] = None,
    min_speakers: int = 1,
    max_speakers: int = 6,
    model: str = "precision-2",
) -> str:
    """Create a hosted diarization job and return jobId."""
    payload = {
        "url": media_url,
        "model": model,
        "exclusive": True,
        "turnLevelConfidence": False,
        "confidence": False,
        "transcription": False,
    }

    if exact_speakers:
        payload["numSpeakers"] = int(exact_speakers)
    else:
        payload["minSpeakers"] = int(min_speakers)
        payload["maxSpeakers"] = int(max_speakers)

    response = requests.post(
        f"{PYANNOTE_BASE_URL}/diarize",
        headers=_pyannote_headers(api_key),
        json=payload,
        timeout=60,
    )
    _raise_pyannote_error(response, "diarization job creation")

    data = response.json()
    job_id = data.get("jobId")
    if not job_id:
        raise RuntimeError(f"pyannoteAI did not return jobId: {data}")
    return job_id


def poll_pyannote_job(
    job_id: str,
    api_key: str,
    timeout_seconds: int = 1800,
    sleep_seconds: int = 8,
    progress_callback=None,
) -> dict:
    """Poll pyannoteAI job until terminal status and return full job payload."""
    start_time = time.time()

    while True:
        response = requests.get(
            f"{PYANNOTE_BASE_URL}/jobs/{job_id}",
            headers=_pyannote_headers(api_key, json_content=False),
            timeout=60,
        )
        _raise_pyannote_error(response, "job polling")

        data = response.json()
        status = data.get("status", "unknown")

        if progress_callback:
            progress_callback(status)

        if status == "succeeded":
            return data
        if status in {"failed", "canceled"}:
            raise RuntimeError(f"pyannoteAI diarization job {status}: {data}")

        if time.time() - start_time > timeout_seconds:
            raise RuntimeError("pyannoteAI diarization timed out. Try a shorter audio file or rerun.")

        time.sleep(sleep_seconds)


def _read_time_value(value) -> float:
    try:
        return float(value)
    except Exception:
        return 0.0


def parse_pyannote_turns(job_payload: dict) -> List[dict]:
    """
    Convert pyannoteAI output to the app's format:
    [{start, end, speaker_raw}, ...]
    Handles both diarization and exclusiveDiarization shapes.
    """
    output = job_payload.get("output", {}) or {}
    diarization = (
        output.get("exclusiveDiarization")
        or output.get("exclusive_diarization")
        or output.get("diarization")
        or []
    )

    turns: List[dict] = []
    for item in diarization:
        if not isinstance(item, dict):
            continue

        segment = item.get("segment") if isinstance(item.get("segment"), dict) else {}
        start = item.get("start", segment.get("start"))
        end = item.get("end", segment.get("end"))
        speaker = (
            item.get("speaker")
            or item.get("label")
            or item.get("speakerLabel")
            or item.get("speaker_label")
            or "Unknown"
        )

        turns.append({
            "start": _read_time_value(start),
            "end": _read_time_value(end),
            "speaker_raw": str(speaker),
        })

    turns = [t for t in turns if t["end"] > t["start"]]
    turns.sort(key=lambda x: (x["start"], x["end"]))
    return turns


def diarize_audio_with_pyannote_api(
    audio_path: str,
    api_key: str,
    exact_speakers: Optional[int] = None,
    min_speakers: int = 1,
    max_speakers: int = 6,
    model: str = "precision-2",
    progress_callback=None,
) -> Tuple[List[dict], dict]:
    """Upload audio, create diarization job, poll result, parse turns."""
    media_url = upload_audio_to_pyannote_media(audio_path, api_key)
    job_id = submit_pyannote_diarization_job(
        media_url=media_url,
        api_key=api_key,
        exact_speakers=exact_speakers,
        min_speakers=min_speakers,
        max_speakers=max_speakers,
        model=model,
    )
    job_payload = poll_pyannote_job(
        job_id=job_id,
        api_key=api_key,
        progress_callback=progress_callback,
    )
    turns = parse_pyannote_turns(job_payload)
    if not turns:
        raise RuntimeError("pyannoteAI completed, but no speaker turns were returned.")
    return turns, job_payload


def best_speaker_for_segment(start: float, end: float, turns: List[dict]) -> str:
    best = "Unknown"
    best_overlap = 0.0

    for t in turns:
        overlap_start = max(start, t["start"])
        overlap_end = min(end, t["end"])
        overlap = max(0.0, overlap_end - overlap_start)
        if overlap > best_overlap:
            best_overlap = overlap
            best = t["speaker_raw"]

    return best


def attach_speakers_to_segments(td: dict, turns: List[dict]) -> dict:
    """Attach pyannoteAI labels to Groq Whisper segments by maximum time overlap."""
    for seg in td.get("segments", []):
        seg["speaker_raw"] = best_speaker_for_segment(seg["start"], seg["end"], turns)

    td["speaker_turns"] = turns
    return normalize_speaker_labels(td)


def normalize_speaker_labels(td: dict) -> dict:
    """
    pyannote cluster IDs are arbitrary, so remap by first appearance:
        SPEAKER_04 -> Speaker 01
        SPEAKER_01 -> Speaker 02
    """
    mapping: Dict[str, str] = {}

    for seg in td.get("segments", []):
        raw = seg.get("speaker_raw", "Unknown") or "Unknown"
        if raw == "Unknown":
            seg["speaker"] = "Unknown"
            continue
        if raw not in mapping:
            mapping[raw] = f"Speaker {len(mapping) + 1:02d}"
        seg["speaker"] = mapping[raw]

    for turn in td.get("speaker_turns", []):
        raw = turn.get("speaker_raw", "Unknown")
        turn["speaker"] = mapping.get(raw, "Unknown")

    td["speaker_map_raw_to_clean"] = mapping
    td["full_text"] = build_speaker_transcript(td, use_name=False)
    return td


def build_speaker_transcript(td: dict, speaker_names: Optional[Dict[str, str]] = None, use_name: bool = True) -> str:
    lines = []
    speaker_names = speaker_names or {}

    for seg in td.get("segments", []):
        speaker = seg.get("speaker", "Unknown")
        display_name = speaker_names.get(speaker, speaker) if use_name else speaker
        lines.append(f"[{fmt_time(seg.get('start', 0))}] {display_name}: {seg.get('text', '')}")

    return "\n".join(lines)


# ── Groq Meeting Analysis ────────────────────────────────────────────────────
ANALYSIS_SCHEMA = """\
{
  "meeting_title": "<inferred title>",
  "date": "<dd MMM YYYY>",
  "duration_estimate": "<e.g. 45 minutes>",
  "attendees_estimate": "<count or range>",
  "executive_summary": "<3-5 sentences>",
  "key_topics": [{"topic": "...", "summary": "..."}],
  "decisions_made": ["..."],
  "action_items": [
    {
      "id": 1,
      "action": "<clear task>",
      "owner": "<speaker/person/team or TBD>",
      "priority": "<High|Medium|Low>",
      "due_days": <integer>,
      "context": "<brief note>"
    }
  ],
  "risks_flagged": ["..."],
  "follow_up_questions": ["..."],
  "sentiment": "<Positive|Neutral|Mixed|Tense>",
  "meeting_effectiveness": "<N/10 — one sentence rationale>"
}"""

ESCAPED_SCHEMA = ANALYSIS_SCHEMA.replace("{", "{{").replace("}", "}}")

SYSTEM_PROMPT = """You are an expert meeting analyst and chief-of-staff AI.
Analyze the speaker-labelled transcript and reply ONLY with valid JSON matching this schema exactly.
No markdown fences, no preamble, no trailing text — pure JSON only.

Schema:
{schema}

Today: {{today}}
Rules:
- Use speaker labels/names when assigning action owners.
- If the actual owner is unclear, use "TBD".
- due_days must be an integer.
- Extract explicit and strongly implied decisions.
- Flag blockers, risks, open questions, and follow-ups.
- Be precise and professional.""".format(schema=ESCAPED_SCHEMA)


GROQ_CHAT_COMPLETIONS_URL = "https://api.groq.com/openai/v1/chat/completions"


def _raise_groq_chat_error(response: requests.Response) -> None:
    if response.ok:
        return
    try:
        detail = response.json()
    except Exception:
        detail = response.text
    raise RuntimeError(f"Groq chat completion failed ({response.status_code}): {detail}")


def analyse_with_groq(transcript: str, filename: str, groq_key: str, model_name: str) -> dict:
    """Create meeting minutes/action tracker using Groq Chat Completions API."""
    if not groq_key:
        raise RuntimeError("GROQ_API_KEY is missing.")

    today = datetime.now().strftime("%d %B %Y")
    system_prompt = SYSTEM_PROMPT.replace("{{today}}", today).replace("{today}", today)
    user_prompt = f"Filename: {filename}\n\nSpeaker-labelled transcript:\n{transcript[:18000]}"

    payload = {
        "model": model_name,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        "temperature": 0.1,
        "max_completion_tokens": 4096,
        "response_format": {"type": "json_object"},
    }

    response = requests.post(
        GROQ_CHAT_COMPLETIONS_URL,
        headers={
            "Authorization": f"Bearer {groq_key}",
            "Content-Type": "application/json",
        },
        json=payload,
        timeout=900,
    )
    _raise_groq_chat_error(response)

    data = response.json()
    raw = (
        data.get("choices", [{}])[0]
        .get("message", {})
        .get("content", "")
        .strip()
    )

    raw = re.sub(r"^```(?:json)?\s*", "", raw, flags=re.IGNORECASE)
    raw = re.sub(r"\s*```$", "", raw)

    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", raw, flags=re.DOTALL)
        if match:
            return json.loads(match.group(0))
        raise


# ── Excel Export ─────────────────────────────────────────────────────────────
def build_excel(analysis: dict, td: dict, filename: str, speaker_names: Optional[Dict[str, str]] = None) -> bytes:
    speaker_names = speaker_names or {}

    wb = Workbook()
    P = {
        "bg": "070712", "slate": "1A1A30", "sub": "111127", "text": "E0E0FF",
        "muted": "707090", "indigo": "6366F1", "high": "F43F5E", "med": "FBBF24",
        "low": "34D399", "row_a": "0A0A18", "row_b": "0D0D20", "white": "FFFFFF",
    }

    def fill(c): return PatternFill("solid", fgColor=c)
    def font(size=10, bold=False, color=None): return Font(name="Arial", size=size, bold=bold, color=color or P["text"])
    def center(): return Alignment(horizontal="center", vertical="center", wrap_text=True)
    def left(): return Alignment(horizontal="left", vertical="center", wrap_text=True)
    border = Border(
        left=Side(style="thin", color="222244"),
        right=Side(style="thin", color="222244"),
        top=Side(style="thin", color="222244"),
        bottom=Side(style="thin", color="222244"),
    )

    def banner(ws, title, end_col="H"):
        ws.sheet_view.showGridLines = False
        ws.merge_cells(f"A1:{end_col}2")
        c = ws["A1"]
        c.value = title
        c.fill = fill(P["indigo"])
        c.font = font(14, True, P["white"])
        c.alignment = center()
        ws.row_dimensions[1].height = 24
        ws.row_dimensions[2].height = 8

    def section(ws, row, title, end_col="H"):
        ws.merge_cells(f"A{row}:{end_col}{row}")
        c = ws[f"A{row}"]
        c.value = title
        c.fill = fill(P["sub"])
        c.font = font(10, True, "93C5FD")
        c.alignment = left()

    # Summary
    ws = wb.active
    ws.title = "Summary"
    banner(ws, "BRANDit · MEETING INTELLIGENCE REPORT")
    r = 4
    section(ws, r, "EXECUTIVE SUMMARY")
    r += 1
    ws.merge_cells(f"A{r}:H{r+3}")
    ws[f"A{r}"].value = analysis.get("executive_summary", "")
    ws[f"A{r}"].fill = fill(P["bg"])
    ws[f"A{r}"].font = font()
    ws[f"A{r}"].alignment = left()

    r += 5
    section(ws, r, "KEY TOPICS")
    for topic in analysis.get("key_topics", []):
        r += 1
        ws.merge_cells(f"A{r}:H{r}")
        ws[f"A{r}"].value = f"- {topic.get('topic','')}: {topic.get('summary','')}"
        ws[f"A{r}"].fill = fill(P["row_a"])
        ws[f"A{r}"].font = font()
        ws[f"A{r}"].alignment = left()
        ws[f"A{r}"].border = border

    r += 2
    section(ws, r, "DECISIONS MADE")
    for decision in analysis.get("decisions_made", []):
        r += 1
        ws.merge_cells(f"A{r}:H{r}")
        ws[f"A{r}"].value = f"Decision: {decision}"
        ws[f"A{r}"].fill = fill(P["row_b"])
        ws[f"A{r}"].font = font(color="A7F3D0")
        ws[f"A{r}"].alignment = left()
        ws[f"A{r}"].border = border

    for col in range(1, 9):
        ws.column_dimensions[get_column_letter(col)].width = 22

    # Action Tracker
    wa = wb.create_sheet("Action Tracker")
    banner(wa, "ACTION POINTS TRACKER", "I")
    headers = ["#", "Action", "Owner", "Priority", "Deadline", "Reminder", "Status", "Context", "Remarks"]
    widths = [6, 46, 22, 14, 16, 16, 14, 34, 28]
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        c = wa.cell(3, i, h)
        c.fill = fill(P["slate"])
        c.font = font(10, True, P["white"])
        c.alignment = center()
        c.border = border
        wa.column_dimensions[get_column_letter(i)].width = w

    prio_color = {"High": P["high"], "Medium": P["med"], "Low": P["low"]}
    for idx, a in enumerate(analysis.get("action_items", []), 1):
        row = 3 + idx
        due_days = int(a.get("due_days", 5) or 5)
        due = datetime.now() + timedelta(days=due_days)
        vals = [
            idx,
            a.get("action", ""),
            a.get("owner", "TBD"),
            a.get("priority", "Medium"),
            due.strftime("%d %b %Y"),
            (due - timedelta(days=1)).strftime("%d %b %Y"),
            "Pending",
            a.get("context", ""),
            "",
        ]
        bg = P["row_a"] if idx % 2 else P["row_b"]
        for col, val in enumerate(vals, 1):
            c = wa.cell(row, col, val)
            c.fill = fill(bg)
            c.border = border
            c.alignment = center() if col in (1, 4, 5, 6, 7) else left()
            c.font = font(color=prio_color.get(val, P["text"]) if col == 4 else P["text"])

    # Transcript
    wt = wb.create_sheet("Transcript")
    banner(wt, "SPEAKER-LABELLED TRANSCRIPT", "D")
    headers = ["Timestamp", "Speaker", "Text", "System Label"]
    widths = [16, 22, 110, 24]
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        c = wt.cell(3, i, h)
        c.fill = fill(P["slate"])
        c.font = font(10, True, P["white"])
        c.alignment = center()
        c.border = border
        wt.column_dimensions[get_column_letter(i)].width = w

    for idx, seg in enumerate(td.get("segments", []), 1):
        row = 3 + idx
        speaker = seg.get("speaker", "Unknown")
        display = speaker_names.get(speaker, speaker)
        vals = [fmt_time(seg.get("start", 0)), display, seg.get("text", ""), seg.get("speaker_raw", "")]
        bg = P["row_a"] if idx % 2 else P["row_b"]
        for col, val in enumerate(vals, 1):
            c = wt.cell(row, col, val)
            c.fill = fill(bg)
            c.border = border
            c.font = font()
            c.alignment = left() if col == 3 else center()

    # Speaker Turns
    wspt = wb.create_sheet("Speaker Turns")
    banner(wspt, "SPEAKER TIMELINE", "D")
    headers = ["Start", "End", "Speaker", "System Label"]
    for i, h in enumerate(headers, 1):
        c = wspt.cell(3, i, h)
        c.fill = fill(P["slate"])
        c.font = font(10, True, P["white"])
        c.alignment = center()
        c.border = border
        wspt.column_dimensions[get_column_letter(i)].width = 24
    for idx, turn in enumerate(td.get("speaker_turns", []), 1):
        row = 3 + idx
        vals = [fmt_time(turn.get("start", 0)), fmt_time(turn.get("end", 0)), turn.get("speaker", ""), turn.get("speaker_raw", "")]
        for col, val in enumerate(vals, 1):
            c = wspt.cell(row, col, val)
            c.fill = fill(P["row_a"] if idx % 2 else P["row_b"])
            c.border = border
            c.font = font()
            c.alignment = center()

    # Risks
    wr = wb.create_sheet("Risks & Follow-ups")
    banner(wr, "RISKS, BLOCKERS & OPEN QUESTIONS", "D")
    r = 4
    section(wr, r, "RISKS", "D")
    for risk in analysis.get("risks_flagged", []):
        r += 1
        wr.merge_cells(f"A{r}:D{r}")
        wr[f"A{r}"].value = f"Risk: {risk}"
        wr[f"A{r}"].fill = fill(P["row_a"])
        wr[f"A{r}"].font = font(color="FCA5A5")
        wr[f"A{r}"].alignment = left()
        wr[f"A{r}"].border = border
    r += 2
    section(wr, r, "FOLLOW-UP QUESTIONS", "D")
    for q in analysis.get("follow_up_questions", []):
        r += 1
        wr.merge_cells(f"A{r}:D{r}")
        wr[f"A{r}"].value = f"? {q}"
        wr[f"A{r}"].fill = fill(P["row_b"])
        wr[f"A{r}"].font = font(color="FDE68A")
        wr[f"A{r}"].alignment = left()
        wr[f"A{r}"].border = border
    for col in range(1, 5):
        wr.column_dimensions[get_column_letter(col)].width = 42

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


# ── Session State ────────────────────────────────────────────────────────────
for key, default in {
    "history": [],
    "current_session": None,
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

# ── Secrets / Environment ────────────────────────────────────────────────────
GROQ_API_KEY = get_secret("GROQ_API_KEY")
PYANNOTE_API_KEY = (
    get_secret("PYANNOTE_API_KEY")
    or get_secret("PYANNOTEAI_API_KEY")
    or get_secret("PYANNOTE_API_TOKEN")
)
DEFAULT_GROQ_MODEL = get_secret("GROQ_MODEL", "llama-3.3-70b-versatile")
DEFAULT_WHISPER_API_MODEL = get_secret("WHISPER_API_MODEL", "whisper-large-v3-turbo")

# ── Sidebar ──────────────────────────────────────────────────────────────────
VALID_GROQ_MODELS = ["llama-3.3-70b-versatile", "llama-3.1-8b-instant", "gemma2-9b-it"]
VALID_WHISPER_API_MODELS = ["whisper-large-v3-turbo", "whisper-large-v3"]
VALID_PYANNOTE_MODELS = ["precision-2", "community-1"]

groq_model = DEFAULT_GROQ_MODEL if DEFAULT_GROQ_MODEL in VALID_GROQ_MODELS else "llama-3.3-70b-versatile"
base_whisper_api_model = DEFAULT_WHISPER_API_MODEL if DEFAULT_WHISPER_API_MODEL in VALID_WHISPER_API_MODELS else "whisper-large-v3-turbo"

with st.sidebar:
    st.markdown("""
    <div class="brand-card">
        <h2>BRANDit Meeting Intelligence</h2>
        <p>Upload a meeting recording and convert it into minutes, decisions, speaker-wise transcript and action tracker.</p>
        <div class="contact-strip">
            Prepared by <b>Kalpanasingh Chauhan</b><br>
            +91 8850159663<br>
            chauhankalpana2020@gmail.com
        </div>
    </div>
    """, unsafe_allow_html=True)

    page = st.radio("Navigation", ["Process Meeting", "History"], label_visibility="collapsed")
    st.divider()

    processing_mode = st.selectbox(
        "Processing Mode",
        ["Fast API", "Balanced API", "Detailed API"],
        index=1,
        help="Fast/Balanced use Groq whisper-large-v3-turbo. Detailed uses whisper-large-v3.",
    )
    if processing_mode == "Detailed API":
        whisper_api_model = "whisper-large-v3"
    else:
        whisper_api_model = base_whisper_api_model

    whisper_api_model = st.selectbox(
        "Transcription Model",
        VALID_WHISPER_API_MODELS,
        index=VALID_WHISPER_API_MODELS.index(whisper_api_model) if whisper_api_model in VALID_WHISPER_API_MODELS else 0,
        help="No local Whisper model is loaded. Transcription runs on Groq API.",
    )

    language_label = st.selectbox("Meeting Language", ["English", "Auto detect"], index=0)
    language = "en" if language_label == "English" else "auto"

    st.divider()
    enable_diarization = st.checkbox("Create speaker-wise transcript", value=True)
    pyannote_model = st.selectbox("Speaker diarization model", VALID_PYANNOTE_MODELS, index=0)
    known_count = st.checkbox("I know the number of speakers", value=True)

    if known_count:
        exact_speakers = st.number_input("Number of speakers", min_value=1, max_value=12, value=2, step=1)
        min_speakers = max_speakers = int(exact_speakers)
    else:
        exact_speakers = None
        min_speakers = st.number_input("Minimum speakers", min_value=1, max_value=12, value=1, step=1)
        max_speakers = st.number_input("Maximum speakers", min_value=int(min_speakers), max_value=12, value=max(2, int(min_speakers)), step=1)

    st.divider()
    st.markdown("**Setup Status**")
    st.markdown(f"Groq transcription: {'<span class=status-ok>Ready</span>' if GROQ_API_KEY else '<span class=status-bad>Needs setup</span>'}", unsafe_allow_html=True)
    st.markdown(f"Groq AI minutes: {'<span class=status-ok>Ready</span>' if GROQ_API_KEY else '<span class=status-bad>Needs setup</span>'}", unsafe_allow_html=True)
    st.markdown(f"Speaker tracking API: {'<span class=status-ok>Ready</span>' if PYANNOTE_API_KEY else '<span class=status-bad>Needs setup</span>'}", unsafe_allow_html=True)
    st.markdown(f"Media upload support: {'<span class=status-ok>Ready</span>' if check_ffmpeg() else '<span class=status-bad>Needs setup</span>'}", unsafe_allow_html=True)

    st.divider()
    st.caption("Supported files: MP3, MP4, WAV, M4A, OGG, WEBM, FLAC, MKV, MOV, AVI")

# ── Process Page ─────────────────────────────────────────────────────────────
if "Process" in page:
    st.markdown("""
    <div class="hero">
        <div class="eyebrow">BRANDit AI Prototype · Meeting Intelligence</div>
        <h1>BRANDit Meeting Intelligence Hub</h1>
        <p>Convert client calls, internal meetings and campaign discussions into clean minutes, decisions, action points, timelines and export-ready reports.</p>
    </div>
    <div class="workflow">
        Workflow
        <span class="active">01 Upload</span>
        <span>02 Transcribe</span>
        <span>03 Identify Speakers</span>
        <span>04 Summarise</span>
        <span>05 Export</span>
    </div>
    """, unsafe_allow_html=True)

    feature_cols = st.columns(3, gap="medium")
    with feature_cols[0]:
        st.markdown("""
        <div class="feature-card">
            <h3>Meeting Minutes</h3>
            <p>Creates a concise executive summary, key discussion points and important decisions from the recording.</p>
        </div>
        """, unsafe_allow_html=True)
    with feature_cols[1]:
        st.markdown("""
        <div class="feature-card">
            <h3>Speaker-wise Transcript</h3>
            <p>Uses hosted speaker diarization to review who said what without manual note-taking.</p>
        </div>
        """, unsafe_allow_html=True)
    with feature_cols[2]:
        st.markdown("""
        <div class="feature-card">
            <h3>Action Tracker</h3>
            <p>Turns follow-ups into owner-wise tasks with priority, deadlines, reminders and Excel-ready status tracking.</p>
        </div>
        """, unsafe_allow_html=True)

    left, right = st.columns([1.08, 0.92], gap="large")

    with left:
        st.markdown('<div class="card"><div class="card-title">Upload Meeting Recording</div>', unsafe_allow_html=True)
        uploaded = st.file_uploader(
            "Meeting audio/video",
            type=["mp3", "mp4", "wav", "m4a", "ogg", "webm", "flac", "mkv", "avi", "mov"],
            label_visibility="collapsed",
        )
        if uploaded:
            st.success(f"Ready for analysis: {uploaded.name} ({uploaded.size / 1024 / 1024:.1f} MB)")
        st.markdown('</div>', unsafe_allow_html=True)

        missing = []
        if not GROQ_API_KEY:
            missing.append("AI minutes setup: GROQ_API_KEY")
        if enable_diarization and not PYANNOTE_API_KEY:
            missing.append("speaker tracking setup: PYANNOTE_API_KEY")
        if not check_ffmpeg():
            missing.append("media processing setup: ffmpeg")

        run = st.button("Analyze Meeting", disabled=(uploaded is None or bool(missing)))
        if missing:
            st.warning("Setup required: " + ", ".join(missing))

    with right:
        st.markdown('<div class="card"><div class="card-title">What This Prototype Delivers</div>', unsafe_allow_html=True)
        deliverables = [
            ("01", "Clear meeting summary", "Short executive-level overview for quick review."),
            ("02", "Discussion points and decisions", "Captures what was discussed and what was agreed."),
            ("03", "Owner-wise action points", "Creates tasks with priority, timelines and reminders."),
            ("04", "Speaker-labelled transcript", "Hosted diarization keeps the meeting record easy to audit and share."),
            ("05", "Excel-ready tracker", "Exports minutes, action items, transcript and follow-ups."),
        ]
        for number, title, desc in deliverables:
            st.markdown(f"""
            <div class="deliverable">
                <div class="deliverable-number">{number}</div>
                <div><b>{title}</b><br>{desc}</div>
            </div>
            """, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    if uploaded and run:
        suffix = Path(uploaded.name).suffix or ".tmp"
        tmp_input = None
        tmp_audio = None

        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as f:
                f.write(uploaded.read())
                tmp_input = f.name

            with tempfile.NamedTemporaryFile(delete=False, suffix=".flac") as f:
                tmp_audio = f.name

            progress = st.progress(0, "Preparing the recording...")
            prepare_audio_for_apis(tmp_input, tmp_audio)

            progress.progress(15, "Creating the transcript...")
            td = transcribe_with_groq_api(tmp_audio, GROQ_API_KEY or "", whisper_api_model, language)

            pyannote_payload = None
            if enable_diarization:
                progress.progress(45, "Uploading audio to speaker tracking API...")

                def api_status(status: str):
                    progress.progress(55, f"Speaker tracking API status: {status}...")

                turns, pyannote_payload = diarize_audio_with_pyannote_api(
                    audio_path=tmp_audio,
                    api_key=PYANNOTE_API_KEY or "",
                    exact_speakers=int(exact_speakers) if exact_speakers else None,
                    min_speakers=int(min_speakers),
                    max_speakers=int(max_speakers),
                    model=pyannote_model,
                    progress_callback=api_status,
                )
                td = attach_speakers_to_segments(td, turns)
            else:
                for seg in td.get("segments", []):
                    seg["speaker_raw"] = "Unknown"
                    seg["speaker"] = "Speaker 01"
                td["speaker_turns"] = []
                td["speaker_map_raw_to_clean"] = {}
                td["full_text"] = build_speaker_transcript(td, use_name=False)

            progress.progress(75, "Generating meeting minutes and action points...")
            transcript_for_llm = build_speaker_transcript(td, use_name=False)
            analysis = analyse_with_groq(transcript_for_llm, uploaded.name, GROQ_API_KEY or "", groq_model)

            progress.progress(92, "Preparing the Excel report...")
            xlsx = build_excel(analysis, td, uploaded.name)

            progress.progress(100, "Analysis complete.")

            session = {
                "id": len(st.session_state.history) + 1,
                "filename": uploaded.name,
                "timestamp": datetime.now().strftime("%d %b %Y · %H:%M"),
                "td": td,
                "analysis": analysis,
                "xlsx": xlsx,
                "model": groq_model,
                "whisper_api_model": whisper_api_model,
                "pyannote_model": pyannote_model,
                "pyannote_payload": pyannote_payload,
            }
            st.session_state.history.append(session)
            st.session_state.current_session = session
            st.success("Meeting analysis completed successfully.")

        except json.JSONDecodeError as e:
            st.error(f"The analysis response was incomplete: {e}. Please rerun the meeting analysis.")
        except Exception as e:
            st.error(f"Error: {e}")
        finally:
            for p in [tmp_input, tmp_audio]:
                if p and os.path.exists(p):
                    try:
                        os.remove(p)
                    except Exception:
                        pass

    # Results
    session = st.session_state.current_session
    if session:
        td = session["td"]
        analysis = session["analysis"]
        segments = td.get("segments", [])
        speaker_labels = sorted({s.get("speaker", "Unknown") for s in segments})

        st.markdown('<div class="metric-strip">', unsafe_allow_html=True)
        metrics = [
            (len(segments), "Transcript Lines"),
            (fmt_time(td.get("duration", 0)), "Duration"),
            (len(speaker_labels), "Speakers"),
            (len(analysis.get("action_items", [])), "Action Points"),
            (len(analysis.get("risks_flagged", [])), "Risks"),
        ]
        for value, label in metrics:
            st.markdown(f'<div class="metric-box"><div class="metric-num">{value}</div><div class="metric-label">{label}</div></div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        speaker_names: Dict[str, str] = {}
        if speaker_labels:
            st.markdown('<div class="section-header">Speaker Names</div>', unsafe_allow_html=True)
            cols = st.columns(min(4, max(1, len(speaker_labels))))
            for i, sp in enumerate(speaker_labels):
                with cols[i % len(cols)]:
                    default = sp if sp != "Unknown" else "Unknown"
                    speaker_names[sp] = st.text_input(f"Name for {sp}", value=default, key=f"name_{session['id']}_{sp}")

        named_transcript = build_speaker_transcript(td, speaker_names=speaker_names, use_name=True)

        tabs = st.tabs(["Minutes", "Transcript", "Action Tracker", "Risks & Follow-ups", "Export"])

        with tabs[0]:
            st.markdown('<div class="card"><div class="card-title">Executive Summary</div>', unsafe_allow_html=True)
            st.write(analysis.get("executive_summary", ""))
            st.markdown('</div>', unsafe_allow_html=True)

            topics = analysis.get("key_topics", [])
            if topics:
                st.markdown('<div class="section-header">Key Topics</div>', unsafe_allow_html=True)
                for t in topics:
                    st.markdown(f"**{t.get('topic','')}**")
                    st.write(t.get("summary", ""))

            decisions = analysis.get("decisions_made", [])
            if decisions:
                st.markdown('<div class="section-header">Decisions Made</div>', unsafe_allow_html=True)
                for d in decisions:
                    st.success(d)

            effectiveness = analysis.get("meeting_effectiveness", "")
            if effectiveness:
                st.info(f"Effectiveness: {effectiveness}")

        with tabs[1]:
            st.markdown(f"<div class='small-muted'>Language: <b>{html.escape(str(td.get('language',''))).upper()}</b> · Duration: <b>{fmt_time(td.get('duration',0))}</b> · Speakers: <b>{len(speaker_labels)}</b></div>", unsafe_allow_html=True)
            rows = []
            for seg in segments[:300]:
                speaker = seg.get("speaker", "Unknown")
                display = speaker_names.get(speaker, speaker)
                rows.append(
                    '<div class="seg-row">'
                    f'<span class="seg-time">{html.escape(fmt_time(seg.get("start", 0)))}</span>'
                    f'<span class="seg-spk">{html.escape(display)}</span>'
                    f'<span class="seg-text">{html.escape(seg.get("text", ""))}</span>'
                    '</div>'
                )
            if len(segments) > 300:
                rows.append(f"<p class='small-muted' style='text-align:center;'>… {len(segments)-300} more segments in export</p>")
            st.markdown('<div class="transcript-box">' + ''.join(rows) + '</div>', unsafe_allow_html=True)

            st.text_area("Plain transcript", named_transcript, height=240, label_visibility="collapsed")

        with tabs[2]:
            actions = analysis.get("action_items", [])
            if actions:
                table = []
                for a in actions:
                    due_days = int(a.get("due_days", 5) or 5)
                    due = datetime.now() + timedelta(days=due_days)
                    table.append({
                        "#": a.get("id", ""),
                        "Action": a.get("action", ""),
                        "Owner": a.get("owner", "TBD"),
                        "Priority": a.get("priority", "Medium"),
                        "Deadline": due.strftime("%d %b %Y"),
                        "Reminder": (due - timedelta(days=1)).strftime("%d %b %Y"),
                        "Status": "Pending",
                        "Context": a.get("context", ""),
                    })
                st.dataframe(pd.DataFrame(table), use_container_width=True, hide_index=True)
            else:
                st.info("No action items detected.")

        with tabs[3]:
            risks = analysis.get("risks_flagged", [])
            followups = analysis.get("follow_up_questions", [])
            if risks:
                st.markdown('<div class="section-header">Risks</div>', unsafe_allow_html=True)
                for r in risks:
                    st.error(r)
            if followups:
                st.markdown('<div class="section-header">Follow-up Questions</div>', unsafe_allow_html=True)
                for q in followups:
                    st.warning(q)
            if not risks and not followups:
                st.info("No major risks or follow-up questions detected.")

        with tabs[4]:
            base = Path(session["filename"]).stem
            c1, c2, c3 = st.columns(3)

            with c1:
                xlsx_named = build_excel(analysis, td, session["filename"], speaker_names=speaker_names)
                st.download_button(
                    "Excel Report",
                    data=xlsx_named,
                    file_name=f"brandit_meeting_report_{base}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            with c2:
                minutes = (
                    f"MEETING MINUTES\n{'='*60}\n"
                    f"{analysis.get('meeting_title','Meeting')}\n"
                    f"Date: {analysis.get('date','')}\n"
                    f"Duration: {analysis.get('duration_estimate','')}\n\n"
                    f"EXECUTIVE SUMMARY\n{'-'*30}\n{analysis.get('executive_summary','')}\n\n"
                    f"DECISIONS\n{'-'*30}\n" + "\n".join(f"- {d}" for d in analysis.get("decisions_made", [])) + "\n\n"
                    f"ACTION ITEMS\n{'-'*30}\n" + "\n".join(f"[{a.get('priority','')}] {a.get('action','')} — {a.get('owner','TBD')}" for a in analysis.get("action_items", [])) + "\n\n"
                    f"RISKS\n{'-'*30}\n" + "\n".join(f"Risk: {r}" for r in analysis.get("risks_flagged", []))
                )
                st.download_button("Minutes TXT", data=minutes, file_name=f"brandit_minutes_{base}.txt", mime="text/plain")

            with c3:
                st.download_button("Speaker Transcript", data=named_transcript, file_name=f"brandit_speaker_transcript_{base}.txt", mime="text/plain")


# ── History Page ─────────────────────────────────────────────────────────────
else:
    st.markdown("""
    <div class="hero">
        <div class="eyebrow">BRANDit AI Prototype · Archive</div>
        <h1>Meeting History</h1>
        <p>Review meeting reports processed during this Streamlit session.</p>
    </div>
    """, unsafe_allow_html=True)

    if not st.session_state.history:
        st.info("No meetings processed yet.")
    else:
        for sess in reversed(st.session_state.history):
            with st.expander(f"{sess['filename']} · {sess['timestamp']}"):
                st.write(sess["analysis"].get("executive_summary", ""))
                st.download_button(
                    "Download Excel Report",
                    data=sess["xlsx"],
                    file_name=f"brandit_meeting_report_{Path(sess['filename']).stem}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"hist_xlsx_{sess['id']}",
                )

st.markdown("""
<div class="footer-bar">
    Prepared by <b>Kalpanasingh Chauhan</b> &nbsp; · &nbsp; +91 8850159663 &nbsp; · &nbsp; chauhankalpana2020@gmail.com
</div>
""", unsafe_allow_html=True)
