# Transparent Image Trimmer

Drag and drop one or more transparent images, then export versions cropped to their visible pixels.

## Features

- Batch image upload by file picker or drag and drop
- Automatic transparent edge trimming by alpha threshold
- Optional padding around the trimmed result
- Per-image PNG download and copy
- Download all processed images

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

The add-in opens the same trimmer UI in a task pane. Drop an image into the pane, or copy a selected image in PowerPoint and paste it into the pane with Ctrl+V. Then use `PPT에 넣기` to insert the trimmed PNG into the current slide.
