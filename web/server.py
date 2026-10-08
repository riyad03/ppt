"""Web interface for the regulatory-note generator.

    python -m uvicorn web.server:app --port 8000

A generation run takes fifteen to twenty-five minutes on CPU, so nothing here
can be a plain request/response: a job is started, the browser polls it, and
the finished `.pptx` is downloaded afterwards.

The pipeline runs as a **subprocess** rather than inside this process, for two
reasons. A failure in generation — and a small local model does fail, sometimes
by raising after three invalid attempts — then kills one job instead of the
server. And the subprocess is the same `agent.py` the command line uses, so the
interface cannot drift from the documented entry point.
"""

from __future__ import annotations

import posixpath
import subprocess
import sys
import threading
import time
import uuid
import zipfile
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path
from xml.etree import ElementTree

from fastapi import FastAPI, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, Response

ROOT = Path(__file__).resolve().parent.parent
TEMPLATES_DIR = ROOT / "templates"
UPLOADS_DIR = TEMPLATES_DIR / "uploads"
OUTPUT_DIR = ROOT / "sorties" / "web"
STATIC = Path(__file__).resolve().parent / "static"

# Phases the browser shows as a progress track. Matched against the lines the
# pipeline prints, so the two stay in step without a second channel.
_PHASES = ("mesure", "plan", "slides", "rendu", "fini")

app = FastAPI(title="Atlas Guardian")


@dataclass
class Job:
    id: str
    status: str = "running"          # running | done | failed
    phase: str = "mesure"
    started: float = field(default_factory=time.time)
    finished: float | None = None
    log: list[str] = field(default_factory=list)
    slides_done: int = 0
    slides_total: int = 0
    measured: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    result: str | None = None
    error: str | None = None

    def elapsed(self) -> float:
        return (self.finished or time.time()) - self.started


JOBS: dict[str, Job] = {}
_LOCK = threading.Lock()


def _read_output(job: Job, process: subprocess.Popen) -> None:
    """Follow the run's stdout, keeping just enough state for the UI.

    The lines are parsed rather than merely displayed, because a wall of log is
    not progress: the browser needs to know which slide is being written and
    what the run finally produced.
    """
    collecting_warnings = False
    for raw in process.stdout:  # type: ignore[union-attr]
        line = raw.rstrip()
        if not line:
            continue
        with _LOCK:
            job.log.append(line)
            if len(job.log) > 400:
                del job.log[:100]

            if line.startswith("Measured "):
                job.phase = "mesure"
            elif line.startswith("  - ") and job.phase == "mesure" and not collecting_warnings:
                job.measured.append(line[4:])
            elif line.startswith("[slide "):
                job.phase = "slides"
                counter = line.split("]")[0].removeprefix("[slide ").strip()
                done, _, total = counter.partition("/")
                if done.isdigit() and total.isdigit():
                    job.slides_done, job.slides_total = int(done), int(total)
            elif line.startswith("[render]"):
                job.phase = "rendu"
            elif line.startswith("[qa]"):
                job.phase = "slides"
                job.slides_done = 0
            elif line.startswith("File: "):
                job.result = line.removeprefix("File: ").strip()
                job.phase = "fini"
            elif line.startswith("Layout warnings"):
                collecting_warnings = True
            elif collecting_warnings and line.startswith("  - "):
                job.warnings.append(line[4:])

    process.wait()
    with _LOCK:
        job.finished = time.time()
        if process.returncode == 0 and job.result:
            job.status = "done"
            job.phase = "fini"
        else:
            job.status = "failed"
            # The useful part of a Python traceback is its last line.
            tail = [l for l in job.log if l.strip()][-1:] if job.log else []
            job.error = tail[0] if tail else f"code de sortie {process.returncode}"


def _start(job: Job, command: list[str]) -> None:
    process = subprocess.Popen(
        command,
        cwd=str(ROOT),
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,
    )
    threading.Thread(target=_read_output, args=(job, process), daemon=True).start()


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    return (STATIC / "index.html").read_text(encoding="utf-8")


@app.get("/api/templates")
def list_templates() -> dict:
    found = []
    for directory in (TEMPLATES_DIR, UPLOADS_DIR):
        if not directory.exists():
            continue
        for path in sorted(directory.glob("*.pptx")):
            if path.name.startswith("~$"):  # a file Word/PowerPoint has open
                continue
            found.append(
                {
                    "name": path.name,
                    "path": str(path.relative_to(ROOT)).replace("\\", "/"),
                    "size": path.stat().st_size,
                    "uploaded": directory == UPLOADS_DIR,
                }
            )
    return {"templates": found}


_IMAGE_TYPES = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".webp": "image/webp",
}
_REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
_P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
_R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def _rels(z: zipfile.ZipFile, part: str) -> dict[str, tuple[str, str]]:
    """Relationship id -> (type, absolute part name) for one package part."""
    rels_name = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
    try:
        root = ElementTree.fromstring(z.read(rels_name))
    except KeyError:
        return {}
    out = {}
    for rel in root.iter(f"{_REL_NS}Relationship"):
        if rel.get("TargetMode") == "External":
            continue
        target = posixpath.normpath(posixpath.join(posixpath.dirname(part), rel.get("Target", "")))
        out[rel.get("Id")] = (rel.get("Type", ""), target.lstrip("/"))
    return out


@lru_cache(maxsize=64)
def _cover(path: str, mtime: float) -> tuple[bytes, str] | None:
    """A picture that stands for the deck in the template gallery.

    Nothing here can render a slide, so the cover is the largest browser-ready
    picture on the first slide — falling back to its layout, then its master,
    which is where a template's background artwork usually lives — and then to
    the small thumbnail PowerPoint saves in docProps. `mtime` is only there to
    invalidate the cache when a file is replaced.
    """
    try:
        with zipfile.ZipFile(path) as z:
            names = set(z.namelist())
            pres = "ppt/presentation.xml"
            first_slide = None
            if pres in names:
                pres_rels = _rels(z, pres)
                lst = ElementTree.fromstring(z.read(pres)).find(f"{_P_NS}sldIdLst")
                if lst is not None and len(lst):
                    first_slide = pres_rels.get(lst[0].get(_R_ID), ("", None))[1]

            part = first_slide
            for _ in range(3):  # slide -> layout -> master
                if not part or part not in names:
                    break
                rels = _rels(z, part)
                images = [
                    t for kind, t in rels.values()
                    if kind.endswith("/image") and t in names
                    and posixpath.splitext(t)[1].lower() in _IMAGE_TYPES
                    # A few hundred bytes is a bullet glyph or a 1-px fill, not artwork.
                    and z.getinfo(t).file_size >= 4096
                ]
                if images:
                    best = max(images, key=lambda t: z.getinfo(t).file_size)
                    return z.read(best), _IMAGE_TYPES[posixpath.splitext(best)[1].lower()]
                part = next(
                    (t for kind, t in rels.values()
                     if kind.endswith("/slideLayout") or kind.endswith("/slideMaster")),
                    None,
                )

            if "docProps/thumbnail.jpeg" in names:
                return z.read("docProps/thumbnail.jpeg"), "image/jpeg"
    except (zipfile.BadZipFile, ElementTree.ParseError, OSError):
        pass
    return None


@app.get("/api/templates/preview")
def template_preview(path: str) -> Response:
    candidate = (ROOT / path).resolve()
    if not candidate.is_file() or TEMPLATES_DIR.resolve() not in candidate.parents:
        raise HTTPException(404, "Gabarit introuvable.")
    cover = _cover(str(candidate), candidate.stat().st_mtime)
    if cover is None:
        raise HTTPException(404, "Aucun aperçu dans ce fichier.")
    data, media_type = cover
    return Response(data, media_type=media_type, headers={"Cache-Control": "max-age=3600"})


@app.post("/api/jobs")
async def create_job(
    request: str = Form(...),
    context: str = Form(""),
    template: str = Form(""),
    upload: UploadFile | None = None,
) -> dict:
    if not request.strip():
        raise HTTPException(400, "La demande est vide.")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    job = Job(id=uuid.uuid4().hex[:12])

    template_path: Path | None = None
    if upload is not None and upload.filename:
        if not upload.filename.lower().endswith(".pptx"):
            raise HTTPException(400, "Le gabarit doit être un fichier .pptx.")
        UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
        template_path = UPLOADS_DIR / Path(upload.filename).name
        template_path.write_bytes(await upload.read())
    elif template:
        candidate = (ROOT / template).resolve()
        # The name comes from the browser, so it is checked against the one
        # place templates may live rather than trusted as a path.
        if not candidate.is_file() or TEMPLATES_DIR.resolve() not in candidate.parents:
            raise HTTPException(400, "Gabarit introuvable.")
        template_path = candidate

    run_dir = OUTPUT_DIR / job.id
    run_dir.mkdir(parents=True, exist_ok=True)

    command = [sys.executable, "-u", "agent.py", "--request", request, "--out", str(run_dir)]
    if context.strip():
        context_file = run_dir / "contexte.txt"
        context_file.write_text(context, encoding="utf-8")
        command += ["--context-file", str(context_file)]
    if template_path is not None:
        command += ["--template", str(template_path)]

    JOBS[job.id] = job
    _start(job, command)
    return {"id": job.id}


@app.get("/api/jobs/{job_id}")
def job_status(job_id: str) -> dict:
    job = JOBS.get(job_id)
    if job is None:
        raise HTTPException(404, "Tâche inconnue.")
    with _LOCK:
        return {
            "id": job.id,
            "status": job.status,
            "phase": job.phase,
            "phases": list(_PHASES),
            "elapsed": round(job.elapsed()),
            "slides_done": job.slides_done,
            "slides_total": job.slides_total,
            "measured": job.measured,
            "warnings": job.warnings,
            "error": job.error,
            "log": job.log[-60:],
            "filename": Path(job.result).name if job.result else None,
        }


@app.get("/api/jobs/{job_id}/file")
def download(job_id: str) -> FileResponse:
    job = JOBS.get(job_id)
    if job is None or not job.result:
        raise HTTPException(404, "Aucun fichier pour cette tâche.")
    path = Path(job.result)
    if not path.is_absolute():
        path = ROOT / path
    if not path.is_file():
        raise HTTPException(404, "Le fichier a disparu du disque.")
    return FileResponse(
        path,
        filename=path.name,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
