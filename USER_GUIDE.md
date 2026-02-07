# User Guide

## Overview

Print Manager helps you prepare and print paired images (front/back) on A4.

## Screenshot

![UI Screenshot](doc/userui.png)

## Getting Started

1. Start the app using `start.bat` (Windows) or `start.sh` (Linux/Mac).
2. Drag images into the left panel or click "Select images".
3. Images are paired in order: 1+2, 3+4, etc.

## Main Views

- **Preview**: Shows the selected pair (front and back).
- **Tiles**: Shows all pairs as thumbnails.

## Reordering Pairs

- Drag the blue handle ("Drag to reorder") to change the order of pairs.

## Swapping Images

- Click "Swap" in a tile, or "Swap front/back" in the preview, to swap images inside a pair.

## Swapping Between Pairs

- Drag a thumbnail from one pair onto another thumbnail to swap those two images.

## Deleting Pairs

- Right-click a pair tile and choose "Delete pair".

## Mirroring

- Right-click an image to open the mirroring menu:
  - No mirroring
  - Mirror horizontally
  - Mirror vertically
  - Mirror both

## Settings

- **Margins (cm)**: Page margins used for PDF/Word export.
- **Mirror back side automatically**: Global mirroring for back side.
- **Scale to A4 width (29.7 cm)**: Scale to full page width (minus margins).
- **Auto trim white borders**: Remove white/transparent borders before scaling.
- **PDF landscape (A4)**: Optional landscape PDF export (default off).
- **Enable debug output**: Show debug log panel.
- **Auto open exported files**: Open PDF/Word after export.

## Export

- **Save as PDF**: Exports all pairs to a PDF (one side per page).
- **Save as Word**: Exports all pairs to a Word document (one side per page).

## Printing

- Click **Print** to generate a PDF and open it in your default PDF viewer.
- Use the viewer's print dialog and enable duplex printing.

## Troubleshooting

- If images do not fill the width, try enabling "Auto trim white borders".
- If a file does not open after export, check that file associations exist for PDF/DOCX.
