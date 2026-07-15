import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";
import vm from "node:vm";

function createElementStub() {
  return {
    addEventListener() {},
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

test("ZIP output contains two UTF-8 entries and a central directory", async () => {
  const zip = await app.createZipBlob([
    { name: "one.png", blob: new Blob([new Uint8Array([1, 2, 3])]) },
    { name: "둘.png", blob: new Blob([new Uint8Array([4, 5])]) }
  ]);
  const bytes = new Uint8Array(await zip.arrayBuffer());
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const endOffset = bytes.byteLength - 22;

  assert.equal(view.getUint32(0, true), 0x04034b50);
  assert.equal(view.getUint32(endOffset, true), 0x06054b50);
  assert.equal(view.getUint16(endOffset + 10, true), 2);
  assert.equal(view.getUint32(view.getUint32(endOffset + 16, true), true), 0x02014b50);
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
