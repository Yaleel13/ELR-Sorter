# ELR Sorter

A production-clean, browser-based local tool for organizing and assigning Elaria visual assets — **ELR 001 through ELR 077**.

All processing happens entirely in your browser. No server, no build step, no data leaves your machine.

---

## What is ELR Sorter?

ELR Sorter lets you:

1. Load your **ELR Master Registry** Excel workbook.
2. Load a **batch of mixed images** (drag-and-drop or file picker).
3. **Smart auto-match** images using visual similarity plus registry metadata.
4. **Manually assign** remaining images by clicking a registry row, then an image card.
5. **Export** clean CSV files and a patched workbook for direct use in your pipeline.

---

## Files

| File | Purpose |
|---|---|
| `index.html` | App shell and layout |
| `styles.css` | Dark UI styling |
| `app.js` | All application logic |
| `README.md` | This file |

---

## Running Locally

### Option 1 — Open directly (simplest)

1. Download or clone this repository.
2. Double-click `index.html` to open it in your browser.

> **Note:** Some browsers restrict local file access for drag-and-drop. If you experience issues, use Live Server (below).

### Option 2 — VS Code + Live Server (recommended)

1. Install [Visual Studio Code](https://code.visualstudio.com/).
2. Install the [Live Server extension](https://marketplace.visualstudio.com/items?itemName=ritwickdey.LiveServer).
3. Open the project folder in VS Code.
4. Right-click `index.html` → **Open with Live Server**.
5. The app opens at `http://127.0.0.1:5500` (or similar).

---

## How to Use

### 1. Load Your Workbook

Click **📂 Load Workbook** and select your `.xlsx` file.

- The app reads the sheet named exactly **`ELR Master Registry`**.
- If that sheet is missing, you'll see a clear error.
- Only rows with valid ELR IDs (ELR 001–ELR 077) are loaded.
- Rows are sorted by the `Seq` column (falls back to ELR number order).

### 2. Load Images

Click **🖼️ Load Images** or drag-and-drop files onto the drop zone.

- Supported formats: PNG, JPG/JPEG, WEBP, GIF.
- Duplicate files (same name + size) are skipped automatically.
- Each image shows its dimensions, orientation, and any detected ELR ID.

### 3. Smart Auto-Assign

Click **⚡ Smart Auto-Assign**.

- The app first attempts a quick filename pass (`ELR_001`, `ELR-001`, `ELR 001`, etc.).
- Then it runs visual similarity matching in the browser using a CLIP model.
- Visual matching compares each image against ELR row context (Asset Title, App Section, Collection, Usage Objective, Visual Direction, Art Notes).
- Use the **Visual Confidence** slider to control strictness (higher = stricter, lower = more aggressive matching).
- A status message shows progress and final counts for visual vs filename matches.
- Unmatched images are left for manual review.

> First run note: the visual model is downloaded once by your browser cache, so first run can take longer.

### 4. Manual Assignment

For images without ELR IDs in their names:

1. Click an **ELR row** in the registry table on the left — it highlights and a detail panel appears.
2. Find the correct image on the right.
3. Click **Assign to ELR XXX** on that image card.

The row is now mapped. To change an assignment, select the row again and click a different image.

### 5. Searching

Use the **search box** (top right of the stats bar) to filter registry rows by:

- ELR ID
- Asset Title
- Collection
- App Section
- Storage Folder
- Workbook Filename
- Currently assigned image name

---

## Exports

### 📊 Export Mapping CSV (`elr_mapping.csv`)

One row per ELR registry entry. Fields:

| Field | Description |
|---|---|
| `elr_id` | ELR ID (e.g. ELR 001) |
| `seq` | Sequence number from workbook |
| `asset_title` | Asset title from workbook |
| `collection` | Collection name |
| `app_section` | App section |
| `storage_folder` | Target storage folder |
| `workbook_filename` | Filename column from workbook |
| `source_image_name` | Assigned image filename |
| `assigned` | `yes` or `no` |

### ✏️ Export Rename CSV (`elr_rename.csv`)

Instructions for renaming and placing files. Fields:

| Field | Description |
|---|---|
| `elr_id` | ELR ID |
| `source_image_name` | Original uploaded filename |
| `rename_to` | Target filename (from workbook, or `ELR_001.png` fallback) |
| `target_folder` | Storage folder from workbook |
| `final_path` | `/{storage_folder}/{rename_to}` |

Rows without an assigned image have blank `source_image_name` and `final_path`.

### 📥 Export Patched Workbook (`{original}_patched.xlsx`)

A copy of your workbook with these columns updated for mapped rows:

- **Filename** — set to the rename target
- **Upload Status** — set to `Mapped`
- **Final URL** — set to `/{storage_folder}/{rename_to}`
- **Notes** — appended with `mapped by ELR Sorter from source image {original filename}`

All other data is preserved exactly.

---

## GitHub Pages Deployment

1. Push the repository to GitHub.
2. Go to **Settings → Pages**.
3. Set Source to **Deploy from a branch** → `main` (or `master`) → `/ (root)`.
4. Click **Save**.
5. Your app will be live at `https://{username}.github.io/{repo-name}/`.

No server configuration needed — the app is fully static.

---

## Privacy

All files (workbook and images) are processed entirely in your browser using the [File API](https://developer.mozilla.org/en-US/docs/Web/API/File_API). Nothing is uploaded to any server.

