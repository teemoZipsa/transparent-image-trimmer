# Transparent Image Trimmer

Drag and drop one or more transparent images, then export versions cropped to their visible pixels.

Use the hosted app: [Transparent Image Trimmer](https://teemozipsa.github.io/transparent-image-trimmer/)

## Features

- Batch image upload by file picker or drag and drop
- Automatic transparent edge trimming by alpha threshold
- Optional padding around the trimmed result
- Per-image PNG download and copy
- One-click PNG download for a single result
- ZIP download for multiple processed images
- Saved threshold, padding, and filename preferences
- Ctrl+Enter download shortcut and per-image removal

## Use

Open `index.html` in a browser, or publish the folder with any static web host.

## PowerPoint add-in prototype

This repo also includes a local PowerPoint task pane add-in manifest.

1. Install dependencies.

   ```powershell
   npm install
   ```

2. Install the local HTTPS certificate used by Office add-ins.

   ```powershell
   npm run certs
   ```

3. Start the HTTPS dev server.

   ```powershell
   npm run dev
   ```

4. Sideload `manifest.xml` into PowerPoint.

The add-in opens a compact task pane UI for PowerPoint. Copy a selected image in PowerPoint, paste it into the pane with Ctrl+V, then insert every trimmed result into the current slide as separate images. File selection is still available for images outside the presentation.
