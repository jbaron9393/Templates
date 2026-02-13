#!/usr/bin/env python3
"""
build_gross_viewer.py

Drop .docx grossing template files into a folder -> generate a single-file HTML
with:
- left nav categories + search
- right preview + copy
- startup category chooser (persisted in localStorage)

Requires:
    pip install python-docx

Usage:
    python build_gross_viewer.py --in "./docs" --out "./grossing_templates_viewer.html"

Optional:
    --cap "./cap_synoptic_generator.html"   # pulls embedded base64 logo if present
    --prefix-mode filename                  # categories prefixed by filename stem (default)
    --prefix-mode none                      # do not prefix categories
    --keep-heading-levels 2                 # use Heading 1..N for display label selection

Notes:
- The parser looks for paragraphs starting with "Dictation Example" and captures following paragraphs
  until the next heading or a known stop label.
- It uses Word paragraph styles (Heading 1/2/3) when available.
"""

from __future__ import annotations

import argparse
import json
import re
from collections import defaultdict
from pathlib import Path

from docx import Document


STOP_PREFIXES = [
    "Dictation Template",
    "Dragon Template",
    "Sections for Histology",
    "Procedure",
    "Description",
    "Example Header",
    "Header Example",
    "Sample Header",
    "Triage Needed",
    "Notes",
    "Header Notes",
    "Orientation",
    "Tips for opening",
    "Other parts",
]

HEADING_STYLES = {"Heading 1": 1, "Heading 2": 2, "Heading 3": 3, "Heading 4": 4, "Heading 5": 5}


def is_heading(paragraph, max_level: int) -> bool:
    name = getattr(paragraph.style, "name", "")
    lvl = HEADING_STYLES.get(name)
    return bool(lvl and lvl <= max_level)


def heading_level(paragraph) -> int | None:
    name = getattr(paragraph.style, "name", "")
    return HEADING_STYLES.get(name)


def extract_logo_tag_from_cap(cap_html_path: Path | None) -> str:
    if not cap_html_path:
        return ""
    try:
        cap_html = cap_html_path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return ""
    m = re.search(r'<img\s+src="(data:image/[^"]+)"', cap_html)
    if not m:
        return ""
    src = m.group(1)
    return f'<img src="{src}" alt="Logo" />'


def collect_dictation_example(paras, start_idx: int, max_heading_level: int) -> str:
    """Collect lines after a 'Dictation Example' paragraph."""
    collected = []
    t0 = paras[start_idx].text.strip()

    # Allow inline content: "Dictation Example: blah..."
    if ":" in t0 and t0.lower().startswith("dictation example"):
        after = t0.split(":", 1)[1].strip()
        if after:
            collected.append(after)

    j = start_idx + 1
    while j < len(paras):
        p = paras[j]
        txt = p.text.strip()

        if not txt:
            # Keep intentional blank lines unless we're immediately before a heading/stop label
            k = j + 1
            while k < len(paras) and not paras[k].text.strip():
                k += 1
            if k >= len(paras):
                break
            nxt = paras[k]
            nxtt = nxt.text.strip()
            if is_heading(nxt, max_heading_level) or any(nxtt.startswith(lbl) for lbl in STOP_PREFIXES):
                break
            collected.append("")
            j = k
            continue

        if is_heading(p, max_heading_level) or any(txt.startswith(lbl) for lbl in STOP_PREFIXES):
            break

        collected.append(txt)
        j += 1

    return "\n".join(collected).strip()


def parse_docx(docx_path: Path, *, max_heading_level: int = 3) -> tuple[list[dict], list[str]]:
    """
    Returns:
      entries: [{category, display, text}, ...]
      ordered_categories: [category, ...] based on appearance in doc
    """
    doc = Document(str(docx_path))
    paras = doc.paragraphs

    entries = []
    # Track headings
    h = {1: None, 2: None, 3: None, 4: None, 5: None}

    for i, p in enumerate(paras):
        txt = p.text.strip()
        if not txt:
            continue

        lvl = heading_level(p)
        if lvl and lvl <= max_heading_level:
            h[lvl] = txt
            # Clear deeper headings
            for d in range(lvl + 1, 6):
                h[d] = None
            continue

        if txt.startswith("Dictation Example"):
            example = collect_dictation_example(paras, i, max_heading_level)
            if not example:
                continue

            # Category: prefer Heading 1, fallback to nearest available
            category = h.get(1) or h.get(2) or "General"

            # Display label: prefer most specific heading available (max -> 1)
            display = None
            for d in range(max_heading_level, 0, -1):
                if h.get(d):
                    display = h[d]
                    break
            if not display:
                display = category

            entries.append({"category": category, "display": display, "text": example})

    # Preserve category order based on headings in doc
    ordered = []
    seen = set()
    for p in paras:
        if getattr(p.style, "name", "") == "Heading 1":
            nm = p.text.strip()
            if nm and nm not in seen and any(e["category"] == nm for e in entries):
                ordered.append(nm)
                seen.add(nm)

    # Append remaining categories
    for e in entries:
        if e["category"] not in seen:
            ordered.append(e["category"])
            seen.add(e["category"])

    return entries, ordered


def build_html(data: list[dict], logo_tag: str) -> str:
    data_json = json.dumps(data, ensure_ascii=False)

    html = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Grossing Template Viewer</title>
  <style>
    * { margin:0; padding:0; box-sizing:border-box; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, sans-serif;
      background-color: #f5f5f5;
      color: #333;
      line-height: 1.6;
    }
    .header {
      background: linear-gradient(135deg, hsl(157, 100%, 22%) 0%, #1e293b 100%);
      color: white;
      padding: 15px 30px;
      display: flex;
      align-items: center;
      gap: 20px;
      border-bottom: 1px solid #334155;
    }
    .header img {
      height: 80px;
      width: auto;
      border-radius: 12px;
      background: white;
      padding: 6px;
      box-shadow: 0 4px 12px rgba(0,0,0,0.25);
    }
    .header-text h1 { color:#f8fafc; font-size: 1.4rem; }
    .header-text p { font-size: 0.95rem; color:#cbd5e1; }

    .main-container {
      display: flex;
      gap: 20px;
      padding: 20px;
      max-width: 1600px;
      margin: 0 auto;
      min-height: calc(100vh - 150px);
    }

    .nav-panel {
      flex: 0 0 36%;
      background: white;
      border-radius: 8px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.1);
      padding: 18px;
      overflow: hidden;
      max-height: calc(100vh - 170px);
      display: flex;
      flex-direction: column;
    }

    .preview-panel {
      flex: 1;
      background: white;
      border-radius: 8px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.1);
      padding: 18px;
      display: flex;
      flex-direction: column;
      max-height: calc(100vh - 170px);
    }

    .panel-header {
      font-size: 1.1rem;
      font-weight: 700;
      color: #000;
      padding-bottom: 12px;
      border-bottom: 2px solid #CFB87C;
      margin-bottom: 12px;
      display:flex;
      align-items:center;
      justify-content: space-between;
      gap: 10px;
    }

    .search {
      width: 100%;
      padding: 10px 12px;
      border: 2px solid #ddd;
      border-radius: 8px;
      font-size: 0.95rem;
      margin-bottom: 12px;
    }
    .search:focus { outline:none; border-color:#CFB87C; }

    .nav {
      overflow-y: auto;
      padding-right: 6px;
      flex: 1;
    }

    .cat {
      margin-bottom: 10px;
      border: 1px solid #eee;
      border-radius: 8px;
      overflow: hidden;
    }
    .cat > button {
      width: 100%;
      text-align: left;
      padding: 12px 12px;
      font-weight: 700;
      border: none;
      background: #fafafa;
      cursor: pointer;
      display:flex;
      align-items:center;
      justify-content: space-between;
      gap: 10px;
    }
    .cat > button:hover { background:#f2f2f2; }
    .cat.open > button { background:#fff8e7; border-bottom:1px solid #eee; }

    .items {
      display: none;
      padding: 10px 10px 12px 10px;
      background: white;
    }
    .cat.open .items { display:block; }

    .item {
      width: 100%;
      text-align: left;
      padding: 10px 10px;
      border-radius: 8px;
      border: 1px solid #e5e7eb;
      background: white;
      cursor: pointer;
      margin-top: 8px;
      font-size: 0.95rem;
    }
    .item:hover { border-color:#CFB87C; background:#fffdf6; }
    .item.active { border-color:#CFB87C; background:#fff8e7; }

    .meta {
      font-size: 0.85rem;
      color: #6b7280;
      margin-top: 6px;
    }

    .preview-title {
      font-size: 1.0rem;
      font-weight: 800;
      margin-bottom: 10px;
    }

    .preview-content {
      flex: 1;
      overflow-y: auto;
      font-family: 'Courier New', Courier, monospace;
      font-size: 0.92rem;
      line-height: 1.5;
      padding: 15px;
      background: #fafafa;
      border: 1px solid #eee;
      border-radius: 8px;
      white-space: pre-wrap;
      margin-bottom: 12px;
    }

    .button-row {
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
    }

    .btn {
      padding: 12px 18px;
      border-radius: 10px;
      font-size: 0.95rem;
      font-weight: 800;
      cursor: pointer;
      border: 2px solid #CFB87C;
      background: #000;
      color: #CFB87C;
      transition: all 0.2s;
    }
    .btn:hover { background:#CFB87C; color:#000; }
    .btn.secondary {
      border-color: #475569;
      background:#111827;
      color:#e2e8f0;
    }
    .btn.secondary:hover { background:#334155; color:#e2e8f0; }

    .btn.copied {
      background: #28a745;
      border-color: #28a745;
      color: white;
    }

    .footer {
      text-align: center;
      padding: 15px;
      color: #888;
      font-size: 0.8rem;
      border-top: 1px solid #eee;
    }

    /* Modal */
    .overlay {
      position: fixed;
      inset: 0;
      background: rgba(0,0,0,0.55);
      display: none;
      align-items: center;
      justify-content: center;
      padding: 18px;
      z-index: 50;
    }
    .overlay.open { display: flex; }
    .modal {
      width: min(900px, 100%);
      background: white;
      border-radius: 14px;
      box-shadow: 0 20px 60px rgba(0,0,0,0.35);
      overflow: hidden;
      border: 1px solid #e5e7eb;
    }
    .modal-header {
      padding: 16px 18px;
      background: #111827;
      color: #f8fafc;
      display:flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
    }
    .modal-header h2 { font-size: 1.05rem; }
    .modal-body { padding: 16px 18px; }
    .pillrow { display:flex; gap:10px; flex-wrap: wrap; margin-bottom: 12px; }
    .pill {
      border: 1px solid #e5e7eb;
      background: #f9fafb;
      padding: 8px 10px;
      border-radius: 999px;
      cursor: pointer;
      font-weight: 700;
      font-size: 0.9rem;
    }
    .pill:hover { border-color:#CFB87C; background:#fffdf6; }
    .catgrid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 10px;
      max-height: 55vh;
      overflow: auto;
      padding-right: 6px;
    }
    .check {
      display:flex;
      gap: 10px;
      align-items:flex-start;
      padding: 10px 10px;
      border: 1px solid #e5e7eb;
      border-radius: 12px;
      background: white;
    }
    .check:hover { border-color:#CFB87C; background:#fffdf6; }
    .check input { margin-top: 3px; }
    .check .label { font-weight: 800; }
    .check .small { font-size: 0.85rem; color:#6b7280; margin-top: 2px; }

    .modal-footer {
      padding: 14px 18px;
      display:flex;
      align-items:center;
      justify-content: space-between;
      gap: 12px;
      border-top: 1px solid #e5e7eb;
      background: #fafafa;
      flex-wrap: wrap;
    }

    @media (max-width: 800px) {
      .catgrid { grid-template-columns: 1fr; }
      .main-container { flex-direction: column; }
      .nav-panel { flex: 1 1 auto; }
    }

    @media (prefers-color-scheme: dark) {
      body { background-color:#0f172a; color:#e2e8f0; }
      .nav-panel, .preview-panel { background:#111827; border: 1px solid #334155; box-shadow: 0 2px 10px rgba(0,0,0,0.45); }
      .panel-header { color:#e2e8f0; border-bottom-color:#CFB87C; }
      .search { background:#0f172a; color:#e2e8f0; border-color:#475569; }
      .cat { border-color:#334155; }
      .cat > button { background:#0b1220; color:#e2e8f0; }
      .cat > button:hover { background:#111b31; }
      .cat.open > button { background:#3a2f18; border-bottom-color:#334155; }
      .items { background:#111827; }
      .item { background:#0f172a; color:#e2e8f0; border-color:#475569; }
      .item:hover { background:#1a2646; border-color:#CFB87C; }
      .item.active { background:#3a2f18; border-color:#CFB87C; }
      .preview-content { background:#0b1220; border-color:#334155; color:#e2e8f0; }
      .meta, .footer { color:#cbd5e1; border-top-color:#334155; }

      .modal { background:#0b1220; border-color:#334155; }
      .modal-body { color:#e2e8f0; }
      .pill { border-color:#334155; background:#111827; color:#e2e8f0; }
      .check { border-color:#334155; background:#111827; }
      .check:hover { background:#1a2646; border-color:#CFB87C; }
      .check .small { color:#cbd5e1; }
      .modal-footer { border-top-color:#334155; background:#0f172a; }
    }
  </style>
</head>
<body>
  <div class="header">
    {LOGO_TAG}
    <div class="header-text">
      <h1>Grossing Template Viewer</h1>
      <p>Pick sections once → browse + copy examples fast</p>
    </div>
  </div>

  <div class="main-container">
    <div class="nav-panel">
      <div class="panel-header">
        <span>Templates</span>
        <span class="meta" id="countMeta"></span>
      </div>

      <div class="button-row" style="margin-bottom:12px;">
        <button class="btn secondary" id="chooseBtn">Choose categories</button>
        <button class="btn secondary" id="resetBtn">Reset selection</button>
      </div>

      <input id="search" class="search" type="text" placeholder="Search (e.g., gallbladder, TURBT, ureter)..." />
      <div id="nav" class="nav"></div>

      <div class="meta" style="margin-top:10px;">
        Local, single-file HTML. No data leaves your computer.
      </div>
    </div>

    <div class="preview-panel">
      <div class="panel-header">
        <span>Preview</span>
      </div>

      <div class="preview-title" id="previewTitle">Select a template</div>
      <div class="preview-content" id="preview">Choose something on the left to see the example gross dictation here.</div>

      <div class="button-row">
        <button class="btn" id="copyBtn">Copy</button>
        <button class="btn secondary" id="copyHeaderBtn">Copy header only</button>
        <button class="btn secondary" id="clearBtn">Clear</button>
      </div>

      <div class="meta" id="status" style="margin-top:10px;"></div>
    </div>
  </div>

  <div class="footer">
    Built from Word templates
  </div>

  <!-- Category chooser modal -->
  <div class="overlay" id="overlay" aria-hidden="true">
    <div class="modal" role="dialog" aria-modal="true" aria-labelledby="modalTitle">
      <div class="modal-header">
        <h2 id="modalTitle">Choose categories to show</h2>
        <div class="meta" id="selMeta"></div>
      </div>
      <div class="modal-body">
        <div class="pillrow">
          <button class="pill" id="selectAllBtn" type="button">Select all</button>
          <button class="pill" id="selectNoneBtn" type="button">Select none</button>
        </div>
        <div class="catgrid" id="catGrid"></div>
      </div>
      <div class="modal-footer">
        <div class="meta">Tip: This saves locally in your browser (localStorage).</div>
        <div class="button-row">
          <button class="btn secondary" id="cancelBtn" type="button">Cancel</button>
          <button class="btn" id="applyBtn" type="button">Apply</button>
        </div>
      </div>
    </div>
  </div>

<script>
const DATA = {DATA_JSON};
const STORAGE_KEY = "gross_viewer_selected_categories_v1";

const navEl = document.getElementById('nav');
const searchEl = document.getElementById('search');
const previewEl = document.getElementById('preview');
const previewTitleEl = document.getElementById('previewTitle');
const countMetaEl = document.getElementById('countMeta');
const statusEl = document.getElementById('status');

const copyBtn = document.getElementById('copyBtn');
const copyHeaderBtn = document.getElementById('copyHeaderBtn');
const clearBtn = document.getElementById('clearBtn');

const chooseBtn = document.getElementById('chooseBtn');
const resetBtn = document.getElementById('resetBtn');

const overlay = document.getElementById('overlay');
const catGrid = document.getElementById('catGrid');
const selMeta = document.getElementById('selMeta');

const selectAllBtn = document.getElementById('selectAllBtn');
const selectNoneBtn = document.getElementById('selectNoneBtn');

const cancelBtn = document.getElementById('cancelBtn');
const applyBtn = document.getElementById('applyBtn');

function esc(s) {
  return (s || '').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
}
function normalize(s) {
  return (s || '').toLowerCase().replace(/\\s+/g, ' ').trim();
}
function setStatus(msg) { statusEl.textContent = msg || ''; }
function setPreview(title, text) {
  previewTitleEl.textContent = title || 'Template';
  previewEl.textContent = text || '';
  setStatus('');
}

function allCategoryNames() {
  return DATA.map(d => d.category);
}
function defaultSelection() {
  return new Set(allCategoryNames());
}
function loadSelection() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    const arr = JSON.parse(raw);
    if (!Array.isArray(arr)) return null;
    return new Set(arr.filter(x => typeof x === "string"));
  } catch (e) {
    return null;
  }
}
function saveSelection(set) {
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(Array.from(set))); }
  catch (e) {}
}

let selected = loadSelection();
if (!selected || selected.size === 0) selected = defaultSelection();

function selectedDATA() {
  const allowed = selected;
  return DATA.filter(d => allowed.has(d.category));
}

function render(filterText = '') {
  const q = normalize(filterText);
  navEl.innerHTML = '';
  let visibleItemCount = 0;

  const DATA2 = selectedDATA();

  DATA2.forEach((catObj) => {
    const catName = catObj.category;
    const items = catObj.items || [];

    const filtered = items.filter(it => {
      const hay = normalize(catName + ' ' + (it.display || '') + ' ' + (it.text || ''));
      return !q || hay.includes(q);
    });
    if (!filtered.length) return;

    visibleItemCount += filtered.length;

    const catDiv = document.createElement('div');
    catDiv.className = 'cat';

    const catBtn = document.createElement('button');
    catBtn.type = 'button';
    catBtn.innerHTML = `<span>${esc(catName)}</span><span class="meta">${filtered.length}</span>`;
    catBtn.addEventListener('click', () => catDiv.classList.toggle('open'));
    catDiv.appendChild(catBtn);

    const itemsDiv = document.createElement('div');
    itemsDiv.className = 'items';

    filtered.forEach((it) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'item';
      btn.textContent = it.display || catName;

      btn.addEventListener('click', () => {
        document.querySelectorAll('.item.active').forEach(el => el.classList.remove('active'));
        btn.classList.add('active');
        setPreview(it.display || catName, it.text || '');
      });

      itemsDiv.appendChild(btn);
    });

    catDiv.appendChild(itemsDiv);
    navEl.appendChild(catDiv);
  });

  countMetaEl.textContent = `${visibleItemCount} example(s) • ${selected.size}/${allCategoryNames().length} categories`;

  const first = navEl.querySelector('.cat');
  if (first && !filterText) first.classList.add('open');
}

async function copyText(t) {
  try { await navigator.clipboard.writeText(t); return true; }
  catch (e) { return false; }
}

copyBtn.addEventListener('click', async () => {
  const t = previewEl.textContent || '';
  if (!t.trim()) return setStatus('Nothing to copy.');
  const ok = await copyText(t);
  if (!ok) return setStatus('Copy failed (clipboard permission).');
  copyBtn.classList.add('copied');
  copyBtn.textContent = 'Copied!';
  setStatus('Copied preview to clipboard.');
  setTimeout(() => { copyBtn.classList.remove('copied'); copyBtn.textContent = 'Copy'; }, 900);
});

copyHeaderBtn.addEventListener('click', async () => {
  const title = (previewTitleEl.textContent || '').trim();
  if (!title || title === 'Select a template') return setStatus('Select a template first.');
  const ok = await copyText(title + ':');
  if (!ok) return setStatus('Copy failed (clipboard permission).');
  copyHeaderBtn.classList.add('copied');
  copyHeaderBtn.textContent = 'Copied!';
  setStatus('Copied header to clipboard.');
  setTimeout(() => { copyHeaderBtn.classList.remove('copied'); copyHeaderBtn.textContent = 'Copy header only'; }, 900);
});

clearBtn.addEventListener('click', () => {
  setPreview('Select a template', 'Choose something on the left to see the example gross dictation here.');
  setStatus('Cleared.');
});

searchEl.addEventListener('input', (e) => render(e.target.value || ''));

/* Modal logic */
let pending = new Set(selected);

function updateSelMeta() {
  selMeta.textContent = `${pending.size}/${allCategoryNames().length} selected`;
}
function openModal() {
  pending = new Set(selected);
  catGrid.innerHTML = "";

  allCategoryNames().forEach((cat) => {
    const div = document.createElement("label");
    div.className = "check";
    const checked = pending.has(cat);

    div.innerHTML = `
      <input type="checkbox" ${checked ? "checked" : ""} />
      <div>
        <div class="label">${esc(cat)}</div>
        <div class="small">Category</div>
      </div>
    `;

    const cb = div.querySelector("input");
    cb.addEventListener("change", () => {
      if (cb.checked) pending.add(cat);
      else pending.delete(cat);
      updateSelMeta();
    });

    catGrid.appendChild(div);
  });

  updateSelMeta();
  overlay.classList.add("open");
  overlay.setAttribute("aria-hidden", "false");
}
function closeModal() {
  overlay.classList.remove("open");
  overlay.setAttribute("aria-hidden", "true");
}

chooseBtn.addEventListener("click", openModal);
resetBtn.addEventListener("click", () => {
  selected = defaultSelection();
  saveSelection(selected);
  render(searchEl.value || "");
  setStatus("Reset category selection.");
});

selectAllBtn.addEventListener("click", () => {
  pending = defaultSelection();
  catGrid.querySelectorAll("input[type=checkbox]").forEach(cb => cb.checked = true);
  updateSelMeta();
});
selectNoneBtn.addEventListener("click", () => {
  pending = new Set();
  catGrid.querySelectorAll("input[type=checkbox]").forEach(cb => cb.checked = false);
  updateSelMeta();
});

cancelBtn.addEventListener("click", () => closeModal());
applyBtn.addEventListener("click", () => {
  if (pending.size === 0) {
    alert("Select at least one category.");
    return;
  }
  selected = new Set(pending);
  saveSelection(selected);
  closeModal();
  render(searchEl.value || "");
  setStatus("Applied category selection.");
});

overlay.addEventListener("click", (e) => {
  if (e.target === overlay) closeModal();
});
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && overlay.classList.contains("open")) closeModal();
});

/* Initial: if no prior selection, force modal once */
const hadStored = !!localStorage.getItem(STORAGE_KEY);
render();
if (!hadStored) openModal();
</script>
</body>
</html>
"""
    return html.replace("{LOGO_TAG}", logo_tag).replace("{DATA_JSON}", data_json)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", required=True, help="Input folder containing .docx files")
    ap.add_argument("--out", dest="out", required=True, help="Output HTML path")
    ap.add_argument("--cap", dest="cap", default=None, help="Optional CAP html to copy base64 logo from")
    ap.add_argument("--prefix-mode", choices=["filename", "none"], default="filename",
                    help="How to separate files into categories")
    ap.add_argument("--keep-heading-levels", type=int, default=3, help="Use Heading 1..N for display/category logic")
    args = ap.parse_args()

    in_dir = Path(args.inp).expanduser().resolve()
    out_path = Path(args.out).expanduser().resolve()
    cap_path = Path(args.cap).expanduser().resolve() if args.cap else None

    if not in_dir.exists() or not in_dir.is_dir():
        raise SystemExit(f"Input folder not found: {in_dir}")

    docx_files = sorted([p for p in in_dir.glob("*.docx") if p.is_file()])
    if not docx_files:
        raise SystemExit(f"No .docx files found in: {in_dir}")

    all_entries = []
    ordered_categories = []
    seen_cat = set()

    for docx in docx_files:
        entries, order = parse_docx(docx, max_heading_level=max(1, args.keep_heading_levels))

        prefix = ""
        if args.prefix_mode == "filename":
            prefix = docx.stem.strip()
            # Keep it shortish
            if len(prefix) > 40:
                prefix = prefix[:40].rstrip() + "…"

        # Prefix categories to keep different docs separate
        for e in entries:
            if prefix:
                e["category"] = f"{prefix} — {e['category']}"
        for c in order:
            c2 = f"{prefix} — {c}" if prefix else c
            if c2 not in seen_cat:
                ordered_categories.append(c2)
                seen_cat.add(c2)

        all_entries.extend(entries)

    # Build data structure for HTML
    tree = defaultdict(list)
    for e in all_entries:
        tree[e["category"]].append({"display": e["display"], "text": e["text"]})

    data = [{"category": c, "items": tree[c]} for c in ordered_categories if c in tree]

    logo_tag = extract_logo_tag_from_cap(cap_path)

    html = build_html(data, logo_tag)
    out_path.write_text(html, encoding="utf-8")

    print(f"Wrote: {out_path}")
    print(f"Docs: {len(docx_files)} | Categories: {len(data)} | Examples: {sum(len(d['items']) for d in data)}")


if __name__ == "__main__":
    main()
