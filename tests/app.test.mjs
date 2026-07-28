import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";
import vm from "node:vm";

function createElementStub() {
  return {
    __clickCount: 0,
    __listeners: new Map(),
    addEventListener(type, listener) {
      const listeners = this.__listeners.get(type) || [];
      listeners.push(listener);
      this.__listeners.set(type, listeners);
    },
    append() {},
    checked: true,
    classList: {
      add() {},
      remove() {},
      toggle() {}
    },
    disabled: false,
    hidden: false,
    innerHTML: "",
    click() {
      this.__clickCount += 1;
    },
    remove() {},
    textContent: "",
    value: ""
  };
}

function loadAppFunctions() {
  const html = readFileSync(new URL("../index.html", import.meta.url), "utf8");
  const script = [...html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g)]
    .map((match) => match[1])
    .find((source) => source.includes("const fileInput"));
  assert.ok(script, "inline app script should exist");

  const elements = new Map();
  const document = {
    addEventListener() {},
    body: createElementStub(),
    createElement: createElementStub,
    querySelector(selector) {
      if (!elements.has(selector)) {
        elements.set(selector, createElementStub());
      }
      return elements.get(selector);
    },
    querySelectorAll() {
      return [];
    }
  };

  document.querySelector("#threshold").value = "8";
  document.querySelector("#padding").value = "0";
  document.querySelector("#downloadOriginalName").checked = true;

  const context = vm.createContext({
    Blob,
    DataView,
    Date,
    Image: class {},
    Int32Array,
    JSON,
    Math,
    Set,
    TextEncoder,
    Uint8Array,
    Uint32Array,
    URL: {
      createObjectURL() { return "blob:test"; },
      revokeObjectURL() {}
    },
    clearTimeout,
    console,
    document,
    localStorage: {
      getItem() { return null; },
      setItem() {}
    },
    navigator: { clipboard: {} },
    setTimeout,
    window: {
      ClipboardItem: class {},
      Office: undefined,
      setTimeout
    }
  });

  vm.runInContext(script, context, { filename: "index.html" });
  return context;
}

const app = loadAppFunctions();

test("alpha bounds are cached for every threshold", () => {
  const pixels = new Uint8ClampedArray([
    0, 0, 0, 0,    0, 0, 0, 10,   0, 0, 0, 255,
    0, 0, 0, 0,    0, 0, 0, 0,    0, 0, 0, 0
  ]);
  const bounds = app.buildTrimBoundsByThreshold({
    getImageData() { return { data: pixels }; }
  }, 3, 2);

  assert.deepEqual(JSON.parse(JSON.stringify(bounds[0])), {
    minX: 1,
    minY: 0,
    maxX: 2,
    maxY: 0
  });
  assert.deepEqual(JSON.parse(JSON.stringify(bounds[10])), {
    minX: 2,
    minY: 0,
    maxX: 2,
    maxY: 0
  });
  assert.equal(bounds[254].minX, 2);
});

test("strong trim ignores an isolated edge pixel", () => {
  const width = 5;
  const height = 5;
  const pixels = new Uint8ClampedArray(width * height * 4);
  const setAlpha = (x, y, alpha) => {
    pixels[(y * width + x) * 4 + 3] = alpha;
  };

  setAlpha(0, 0, 255);
  for (let y = 1; y <= 3; y += 1) {
    for (let x = 1; x <= 3; x += 1) {
      setAlpha(x, y, 255);
    }
  }

  const context = {
    getImageData() {
      return { data: pixels };
    }
  };
  const exactBounds = app.buildTrimBoundsByThreshold(context, width, height)[8];
  const strongBounds = app.buildStrongTrimBounds(context, width, height, 8, exactBounds);

  assert.deepEqual(JSON.parse(JSON.stringify(exactBounds)), {
    minX: 0,
    minY: 0,
    maxX: 3,
    maxY: 3
  });
  assert.deepEqual(JSON.parse(JSON.stringify(strongBounds)), {
    minX: 1,
    minY: 1,
    maxX: 3,
    maxY: 3
  });
});

test("eight-image batch processes every item and reports progress", async () => {
  const items = Array.from({ length: 8 }, (_, index) => ({ index }));
  const processed = [];
  const progress = [];

  const completed = await app.processItemsSequentially(
    items,
    (item) => processed.push(item.index),
    (done, total) => progress.push([done, total]),
    () => true
  );

  assert.equal(completed, true);
  assert.deepEqual(processed, [0, 1, 2, 3, 4, 5, 6, 7]);
  assert.deepEqual(progress.at(-1), [8, 8]);
});

test("duplicate PNG names receive stable suffixes", () => {
  assert.deepEqual(Array.from(app.makeUniqueFileNames([
    "image-trimmed.png",
    "IMAGE-trimmed.png",
    "image-trimmed.png"
  ])), [
    "image-trimmed.png",
    "IMAGE-trimmed-2.png",
    "image-trimmed-3.png"
  ]);
});

test("multiple results are downloaded as individual PNG files", async () => {
  app.__downloadedFiles = [];
  vm.runInContext(`
    imageItems = [
      {
        valid: true,
        fileName: "image.png",
        trimmedCanvas: { toBlob(callback) { callback(new Blob(["one"], { type: "image/png" })); } },
        insertButton: { disabled: false }
      },
      {
        valid: true,
        fileName: "IMAGE.png",
        trimmedCanvas: { toBlob(callback) { callback(new Blob(["two"], { type: "image/png" })); } },
        insertButton: { disabled: false }
      }
    ];
    downloadBlob = (blob, fileName) => __downloadedFiles.push({ blob, fileName });
  `, app);

  await app.downloadAllImages();

  assert.deepEqual(
    Array.from(app.__downloadedFiles, (entry) => entry.fileName),
    ["image-trimmed.png", "IMAGE-trimmed-2.png"]
  );
  assert.ok(app.__downloadedFiles.every((entry) => entry.blob.type === "image/png"));
});

test("ZIP output contains all named PNG entries", async () => {
  const zip = await app.createZipBlob([
    { name: "첫 이미지.png", blob: new Blob(["one"], { type: "image/png" }) },
    { name: "second.png", blob: new Blob(["two"], { type: "image/png" }) }
  ]);
  const bytes = new Uint8Array(await zip.arrayBuffer());
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const decoded = new TextDecoder().decode(bytes);

  assert.equal(zip.type, "application/zip");
  assert.equal(view.getUint32(0, true), 0x04034b50);
  assert.equal(view.getUint32(bytes.length - 22, true), 0x06054b50);
  assert.equal(view.getUint16(bytes.length - 12, true), 2);
  assert.match(decoded, /첫 이미지\.png/);
  assert.match(decoded, /second\.png/);
});

test("multiple results can be downloaded as one ZIP", async () => {
  app.__downloadedFiles = [];
  vm.runInContext(`
    isDownloading = false;
    isTrimming = false;
    imageItems = [
      {
        valid: true,
        fileName: "one.png",
        trimmedCanvas: { toBlob(callback) { callback(new Blob(["one"], { type: "image/png" })); } }
      },
      {
        valid: true,
        fileName: "two.png",
        trimmedCanvas: { toBlob(callback) { callback(new Blob(["two"], { type: "image/png" })); } }
      }
    ];
    downloadBlob = (blob, fileName) => __downloadedFiles.push({ blob, fileName });
  `, app);

  await app.downloadAllImagesAsZip();

  assert.equal(app.__downloadedFiles.length, 1);
  assert.equal(app.__downloadedFiles[0].fileName, "trimmed-images.zip");
  assert.equal(app.__downloadedFiles[0].blob.type, "application/zip");
});

test("ZIP download is enabled only for multiple valid results", () => {
  vm.runInContext(`
    isDownloading = false;
    isTrimming = false;
    imageItems = [{ valid: true }, { valid: true }];
    updateBatchButtons();
    __zipEnabledForBatch = !zipButton.disabled;

    imageItems = [{ valid: true }];
    updateBatchButtons();
    __zipDisabledForSingle = zipButton.disabled;
  `, app);

  assert.equal(app.__zipEnabledForBatch, true);
  assert.equal(app.__zipDisabledForSingle, true);
});

test("drop zone keyboard activation opens the file picker", () => {
  const dropZone = app.document.querySelector("#dropZone");
  const fileInput = app.document.querySelector("#fileInput");
  const keydown = dropZone.__listeners.get("keydown")[0];
  let prevented = false;

  keydown({
    key: "Enter",
    preventDefault() {
      prevented = true;
    }
  });

  assert.equal(prevented, true);
  assert.equal(fileInput.__clickCount, 1);
});

test("PowerPoint placements stay inside the slide and do not overlap", () => {
  const items = Array.from({ length: 4 }, () => ({
    trimmedCanvas: { width: 800, height: 600 }
  }));
  const placements = Array.from(app.createPowerPointPlacements(items));

  assert.equal(placements.length, 4);
  for (const placement of placements) {
    assert.ok(placement.left >= 36);
    assert.ok(placement.top >= 36);
    assert.ok(placement.left + placement.width <= 924);
    assert.ok(placement.top + placement.height <= 504);
  }
  assert.notEqual(placements[0].left, placements[1].left);
  assert.notEqual(placements[0].top, placements[2].top);
});
