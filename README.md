# Template Viewer: easy editing workflow

Yes—there are two easy ways to tweak formatting and see results quickly.

## 1) Use VS Code (best for permanent changes)

1. Open this folder in VS Code.
2. Install recommended extensions when prompted (Live Server + Prettier).
3. Open `index.html`.
4. Right-click and choose **Open with Live Server**.
5. Edit HTML/CSS/JS and save. The browser auto-reloads.

This repo includes workspace settings so format-on-save is enabled for HTML/CSS/JS.

## 2) Use browser DevTools (best for quick experiments)

1. Run the page in your browser (for example from Live Server).
2. Press `F12` / **Inspect**.
3. In the **Elements** panel, edit styles live to experiment.
4. Once it looks right, copy those changes back into the files in VS Code.

> Tip: DevTools edits are temporary unless you copy them into files.

## Quick command-line option

If you prefer terminal only, you can serve this locally with:

```bash
python3 -m http.server 5500
```

Then open `http://localhost:5500`.
