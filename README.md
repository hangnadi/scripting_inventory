
# Inventory Excel Generator from Images

This tool scans a folder of images and generates an Excel workbook with:
- A thumbnail preview of each image (embedded in the sheet)
- A temporary name derived from the file name (without extension)

## Features
- Supports `.jpg`, `.jpeg`, `.png`, `.webp`
- Adjustable thumbnail size (default 100px)
- Optional recursive scanning of subfolders
- Optional timestamp appended to output filename
- Sensible column widths and frozen header row

## Quick Start

1. **Prepare your images**
   - Put all item photos into a folder, e.g. `images/`

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the script**
   ```bash
   python inventory_from_images.py --input ./images --output ./inventory_audit.xlsx --timestamped --recursive
   ```

### Minimal command
```bash
python inventory_from_images.py
```
(defaults to `--input ./images` and `--output ./inventory_audit.xlsx`)

## Arguments
- `--input` (or `-i`): Path to the folder containing images (default: `./images`)
- `--output` (or `-o`): Path to the output Excel file (default: `./inventory_audit.xlsx`)
- `--thumb-size`: Thumbnail size in pixels for width & height (default: `160`)
- `--timestamped`: Append a timestamp to the output file name
- `--recursive`: Search images inside subfolders too

## Tips
- **Naming:** Keep photo file names meaningful. The script uses the file name as the temporary item name.
- **Google Drive/CDN later:** Start locally to get the Excel correct. Later, we can:
  - Point `--input` to a mounted/synced folder (e.g., Google Drive for desktop).
  - Or extend the script to list Drive files via the Google Drive API and download them locally before processing.

## Troubleshooting
- If you see `ERROR: Pillow is not installed`, run `pip install Pillow`.
- If thumbnails look squished in Excel, increase `--thumb-size` and reopen the file.
- If some images fail, check the console output; the script will continue and mark the row.

---

Built to support quick warehouse/crafting inventory audits.
