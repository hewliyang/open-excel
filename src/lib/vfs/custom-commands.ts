import type { Command, CustomCommand } from "just-bash/browser";
import { defineCommand } from "just-bash/browser";
import type { CellInput } from "../excel/api";
import { getRangeAsCsv, getWorksheetById, setCellRange } from "../excel/api";
import { loadSavedConfig } from "../provider-config";
import { loadWebConfig } from "../web/config";
import { fetchWeb } from "../web/fetch";
import { searchWeb } from "../web/search";

function columnIndexToLetter(index: number): string {
  let letter = "";
  let temp = index;
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let current = "";
  let inQuotes = false;
  let row: string[] = [];

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    const next = text[i + 1];

    if (inQuotes) {
      if (ch === '"') {
        if (next === '"') {
          current += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        current += ch;
      }
    } else {
      if (ch === '"') {
        inQuotes = true;
      } else if (ch === ",") {
        row.push(current);
        current = "";
      } else if (ch === "\n") {
        row.push(current);
        current = "";
        if (row.length > 0) rows.push(row);
        row = [];
      } else if (ch === "\r") {
        // skip, \n will handle the row break
      } else {
        current += ch;
      }
    }
  }

  // Final field/row
  row.push(current);
  if (row.some((cell) => cell !== "")) rows.push(row);

  return rows;
}

function parseStartCell(startCell: string): { col: number; row: number } {
  const match = startCell.match(/^([A-Z]+)(\d+)$/i);
  if (!match) return { col: 0, row: 0 };
  const col =
    match[1]
      .toUpperCase()
      .split("")
      .reduce((acc, c) => acc * 26 + c.charCodeAt(0) - 64, 0) - 1;
  const row = Number.parseInt(match[2], 10) - 1;
  return { col, row };
}

function buildRangeAddress(
  startCell: string,
  rows: number,
  cols: number,
): string {
  const { col, row } = parseStartCell(startCell);
  const endCol = columnIndexToLetter(col + cols - 1);
  const endRow = row + rows;
  return `${startCell}:${endCol}${endRow}`;
}

function coerceValue(raw: string): string | number | boolean {
  if (raw === "") return "";
  if (raw.toLowerCase() === "true") return true;
  if (raw.toLowerCase() === "false") return false;
  const num = Number(raw);
  if (!Number.isNaN(num) && raw.trim() !== "") return num;
  return raw;
}

const csvToSheet: Command = defineCommand("csv-to-sheet", async (args, ctx) => {
  // Extract flags
  const force = args.includes("--force") || args.includes("-f");
  const positional = args.filter((a) => a !== "--force" && a !== "-f");

  if (positional.length < 2) {
    return {
      stdout: "",
      stderr:
        "Usage: csv-to-sheet <file> <sheetId> [startCell] [--force]\n  file      - Path to CSV file in VFS\n  sheetId   - Target sheet ID (number)\n  startCell - Top-left cell, default A1\n  --force   - Overwrite existing cell data\n",
      exitCode: 1,
    };
  }

  const [filePath, sheetIdStr, startCell = "A1"] = positional;
  const sheetId = Number.parseInt(sheetIdStr, 10);
  if (Number.isNaN(sheetId)) {
    return {
      stdout: "",
      stderr: `Invalid sheetId: ${sheetIdStr}`,
      exitCode: 1,
    };
  }

  const upperStartCell = startCell.toUpperCase();
  if (!/^[A-Z]+\d+$/.test(upperStartCell)) {
    return {
      stdout: "",
      stderr: `Invalid start cell: ${startCell}`,
      exitCode: 1,
    };
  }

  try {
    const resolvedPath = filePath.startsWith("/")
      ? filePath
      : `${ctx.cwd}/${filePath}`;
    const content = await ctx.fs.readFile(resolvedPath);
    const rows = parseCsv(content);

    if (rows.length === 0) {
      return { stdout: "", stderr: "CSV file is empty", exitCode: 1 };
    }

    // Normalize column count (pad shorter rows)
    const maxCols = Math.max(...rows.map((r) => r.length));
    const cells: CellInput[][] = rows.map((row) => {
      const padded = [...row];
      while (padded.length < maxCols) padded.push("");
      return padded.map((raw) => ({ value: coerceValue(raw) }));
    });

    const rangeAddr = buildRangeAddress(upperStartCell, rows.length, maxCols);
    const result = await setCellRange(sheetId, rangeAddr, cells, {
      allowOverwrite: force,
    });

    return {
      stdout: `Imported ${rows.length} rows × ${maxCols} columns into sheet ${sheetId} at ${upperStartCell} (${rangeAddr}). ${result.cellsWritten} cells written.`,
      stderr: "",
      exitCode: 0,
    };
  } catch (error) {
    const msg = error instanceof Error ? error.message : String(error);
    return { stdout: "", stderr: msg, exitCode: 1 };
  }
});

function looksLikeRange(s: string): boolean {
  return /^[A-Z]+\d+(:[A-Z]+\d+)?$/i.test(s);
}

async function getUsedRangeAddress(sheetId: number): Promise<string | null> {
  return Excel.run(async (context) => {
    const sheet = await getWorksheetById(context, sheetId);
    if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
    const usedRange = sheet.getUsedRangeOrNullObject();
    usedRange.load("address");
    await context.sync();
    if (usedRange.isNullObject) return null;
    return usedRange.address.split("!")[1] || usedRange.address;
  });
}

const sheetToCsv: Command = defineCommand("sheet-to-csv", async (args, ctx) => {
  if (args.length < 1) {
    return {
      stdout: "",
      stderr:
        "Usage: sheet-to-csv <sheetId> [range] [file]\n  sheetId - Source sheet ID (number)\n  range   - Cell range, e.g. A1:D100 (optional, defaults to used range)\n  file    - Output file path (optional, prints to stdout if omitted)\n",
      exitCode: 1,
    };
  }

  // Parse args: sheetId is always first, then optionally a range, then optionally a file
  const sheetIdStr = args[0];
  const sheetId = Number.parseInt(sheetIdStr, 10);
  if (Number.isNaN(sheetId)) {
    return {
      stdout: "",
      stderr: `Invalid sheetId: ${sheetIdStr}`,
      exitCode: 1,
    };
  }

  let rangeAddr: string | undefined;
  let outFile: string | undefined;

  if (args.length === 2) {
    // Could be range or file
    if (looksLikeRange(args[1])) {
      rangeAddr = args[1];
    } else {
      outFile = args[1];
    }
  } else if (args.length >= 3) {
    rangeAddr = args[1];
    outFile = args[2];
  }

  try {
    // Auto-detect used range if none specified
    if (!rangeAddr) {
      const usedAddr = await getUsedRangeAddress(sheetId);
      if (!usedAddr) {
        return {
          stdout: "",
          stderr: "Sheet is empty (no used range)",
          exitCode: 1,
        };
      }
      rangeAddr = usedAddr;
    }

    const result = await getRangeAsCsv(sheetId, rangeAddr, { maxRows: 50000 });

    if (outFile) {
      const resolvedPath = outFile.startsWith("/")
        ? outFile
        : `${ctx.cwd}/${outFile}`;
      // Ensure parent directory exists
      const dir = resolvedPath.substring(0, resolvedPath.lastIndexOf("/"));
      if (dir && dir !== "/") {
        try {
          await ctx.fs.mkdir(dir, { recursive: true });
        } catch {
          // directory may already exist
        }
      }
      await ctx.fs.writeFile(resolvedPath, result.csv);
      const moreNote = result.hasMore
        ? " (truncated, more rows available)"
        : "";
      return {
        stdout: `Exported ${result.rowCount} rows × ${result.columnCount} columns from "${result.sheetName}" to ${outFile}${moreNote}`,
        stderr: "",
        exitCode: 0,
      };
    }

    // No file → stdout (pipeable)
    return {
      stdout: result.csv,
      stderr: "",
      exitCode: 0,
    };
  } catch (error) {
    const msg = error instanceof Error ? error.message : String(error);
    return { stdout: "", stderr: msg, exitCode: 1 };
  }
});

async function resolveVfsPath(
  ctx: { cwd: string; fs: { readFileBuffer(p: string): Promise<Uint8Array> } },
  filePath: string,
): Promise<{ path: string; data: Uint8Array }> {
  const resolved = filePath.startsWith("/")
    ? filePath
    : `${ctx.cwd}/${filePath}`;
  const data = await ctx.fs.readFileBuffer(resolved);
  return { path: resolved, data };
}

async function writeVfsOutput(
  ctx: {
    cwd: string;
    fs: {
      mkdir(p: string, o: { recursive: boolean }): Promise<void>;
      writeFile(p: string, c: string): Promise<void>;
    };
  },
  outFile: string,
  content: string,
): Promise<string> {
  const resolved = outFile.startsWith("/") ? outFile : `${ctx.cwd}/${outFile}`;
  const dir = resolved.substring(0, resolved.lastIndexOf("/"));
  if (dir && dir !== "/") {
    try {
      await ctx.fs.mkdir(dir, { recursive: true });
    } catch {
      // directory may already exist
    }
  }
  await ctx.fs.writeFile(resolved, content);
  return resolved;
}

const pdfToText: CustomCommand = {
  name: "pdf-to-text",
  load: async () =>
    defineCommand("pdf-to-text", async (args, ctx) => {
      if (args.length < 2) {
        return {
          stdout: "",
          stderr:
            "Usage: pdf-to-text <file> <outfile>\n  file    - Path to PDF file in VFS\n  outfile - Output text file\n",
          exitCode: 1,
        };
      }

      const [filePath, outFile] = args;

      try {
        const { data } = await resolveVfsPath(ctx, filePath);
        await import("pdfjs-dist/build/pdf.worker.mjs");
        const pdfjsLib = await import("pdfjs-dist");

        const doc = await pdfjsLib.getDocument({
          data,
          useWorkerFetch: false,
          isEvalSupported: false,
          useSystemFonts: true,
        }).promise;
        const pages: string[] = [];

        for (let i = 1; i <= doc.numPages; i++) {
          const page = await doc.getPage(i);
          const content = await page.getTextContent();
          const text = content.items
            .filter((item) => "str" in item)
            .map((item) => (item as { str: string }).str)
            .join(" ");
          if (text.trim()) pages.push(text);
        }

        const fullText = pages.join("\n\n");
        await writeVfsOutput(ctx, outFile, fullText);

        return {
          stdout: `Extracted text from ${doc.numPages} page(s) to ${outFile} (${fullText.length} chars)`,
          stderr: "",
          exitCode: 0,
        };
      } catch (error) {
        const msg = error instanceof Error ? error.message : String(error);
        return { stdout: "", stderr: msg, exitCode: 1 };
      }
    }),
};

function parsePageRanges(spec: string, maxPage: number): Set<number> {
  const pages = new Set<number>();
  for (const part of spec.split(",")) {
    const trimmed = part.trim();
    const rangeParts = trimmed.split("-");
    if (rangeParts.length === 2) {
      const start = Math.max(1, Number.parseInt(rangeParts[0], 10));
      const end = Math.min(maxPage, Number.parseInt(rangeParts[1], 10));
      if (!Number.isNaN(start) && !Number.isNaN(end)) {
        for (let i = start; i <= end; i++) pages.add(i);
      }
    } else {
      const p = Number.parseInt(trimmed, 10);
      if (!Number.isNaN(p) && p >= 1 && p <= maxPage) pages.add(p);
    }
  }
  return pages;
}

const pdfToImages: CustomCommand = {
  name: "pdf-to-images",
  load: async () =>
    defineCommand("pdf-to-images", async (args, ctx) => {
      const positional = args.filter((a) => !a.startsWith("--"));
      const scaleArg = args.find((a) => a.startsWith("--scale="));
      const pagesArg = args.find((a) => a.startsWith("--pages="));

      if (positional.length < 2) {
        return {
          stdout: "",
          stderr:
            "Usage: pdf-to-images <file> <outdir> [--scale=N] [--pages=1,3,5-8]\n  file    - Path to PDF file in VFS\n  outdir  - Output directory for PNG images\n  --scale - Render scale factor (default: 2)\n  --pages - Page selection (e.g. 1,3,5-8). Default: all\n",
          exitCode: 1,
        };
      }

      const [filePath, outDir] = positional;
      const scale = scaleArg ? Number.parseFloat(scaleArg.split("=")[1]) : 2;

      if (Number.isNaN(scale) || scale <= 0 || scale > 5) {
        return {
          stdout: "",
          stderr: "Scale must be between 0 and 5",
          exitCode: 1,
        };
      }

      try {
        const { data } = await resolveVfsPath(ctx, filePath);
        await import("pdfjs-dist/build/pdf.worker.mjs");
        const pdfjsLib = await import("pdfjs-dist");

        const doc = await pdfjsLib.getDocument({
          data,
          useWorkerFetch: false,
          isEvalSupported: false,
          useSystemFonts: true,
        }).promise;

        const selectedPages = pagesArg
          ? parsePageRanges(pagesArg.split("=")[1], doc.numPages)
          : new Set(Array.from({ length: doc.numPages }, (_, i) => i + 1));

        if (selectedPages.size === 0) {
          return {
            stdout: "",
            stderr: "No valid pages in selection",
            exitCode: 1,
          };
        }

        const resolvedDir = outDir.startsWith("/")
          ? outDir
          : `${ctx.cwd}/${outDir}`;
        try {
          await ctx.fs.mkdir(resolvedDir, { recursive: true });
        } catch {
          // may exist
        }

        const outputs: string[] = [];
        const sortedPages = [...selectedPages].sort((a, b) => a - b);

        for (const pageNum of sortedPages) {
          const page = await doc.getPage(pageNum);
          const viewport = page.getViewport({ scale });

          const canvas = document.createElement("canvas");
          canvas.width = Math.floor(viewport.width);
          canvas.height = Math.floor(viewport.height);
          const canvasCtx = canvas.getContext("2d");
          if (!canvasCtx) throw new Error("Failed to create canvas 2D context");

          await page.render({ canvasContext: canvasCtx, canvas, viewport })
            .promise;

          const pngData = await new Promise<Uint8Array>((resolve, reject) => {
            canvas.toBlob((blob) => {
              if (!blob) return reject(new Error("Canvas toBlob failed"));
              blob.arrayBuffer().then((buf) => resolve(new Uint8Array(buf)));
            }, "image/png");
          });

          const pagePath = `${resolvedDir}/page-${pageNum}.png`;
          await ctx.fs.writeFile(pagePath, pngData);
          outputs.push(
            `page-${pageNum}.png (${Math.round(pngData.length / 1024)}KB, ${canvas.width}×${canvas.height})`,
          );

          // Help GC
          canvas.width = 0;
          canvas.height = 0;
        }

        return {
          stdout: `Converted ${outputs.length} page(s) from ${doc.numPages} total to ${outDir}/:\n${outputs.map((o) => `  ${o}`).join("\n")}`,
          stderr: "",
          exitCode: 0,
        };
      } catch (error) {
        const msg = error instanceof Error ? error.message : String(error);
        return { stdout: "", stderr: msg, exitCode: 1 };
      }
    }),
};

const docxToText: CustomCommand = {
  name: "docx-to-text",
  load: async () =>
    defineCommand("docx-to-text", async (args, ctx) => {
      if (args.length < 2) {
        return {
          stdout: "",
          stderr:
            "Usage: docx-to-text <file> <outfile>\n  file    - Path to DOCX file in VFS\n  outfile - Output text file\n",
          exitCode: 1,
        };
      }

      const [filePath, outFile] = args;

      try {
        const { data } = await resolveVfsPath(ctx, filePath);
        const mammoth = await import("mammoth");
        const result = await mammoth.extractRawText({
          arrayBuffer: data.buffer as ArrayBuffer,
        });

        await writeVfsOutput(ctx, outFile, result.value);

        return {
          stdout: `Extracted text from DOCX to ${outFile} (${result.value.length} chars)`,
          stderr: "",
          exitCode: 0,
        };
      } catch (error) {
        const msg = error instanceof Error ? error.message : String(error);
        return { stdout: "", stderr: msg, exitCode: 1 };
      }
    }),
};

const xlsxToCsv: CustomCommand = {
  name: "xlsx-to-csv",
  load: async () =>
    defineCommand("xlsx-to-csv", async (args, ctx) => {
      if (args.length < 2) {
        return {
          stdout: "",
          stderr:
            "Usage: xlsx-to-csv <file> <outfile> [sheet]\n  file    - Path to XLSX/XLS/ODS file in VFS\n  outfile - Output CSV file (for multiple sheets: <name>.<sheet>.csv)\n  sheet   - Sheet name or 0-based index (optional, exports all sheets if omitted)\n",
          exitCode: 1,
        };
      }

      const [filePath, outFile, sheetArg] = args;

      try {
        const { data } = await resolveVfsPath(ctx, filePath);
        const XLSX = await import("xlsx");
        const workbook = XLSX.read(data, { type: "array" });

        // Specific sheet requested
        if (sheetArg) {
          let sheetName: string;
          if (workbook.SheetNames.includes(sheetArg)) {
            sheetName = sheetArg;
          } else {
            const idx = Number.parseInt(sheetArg, 10);
            if (
              !Number.isNaN(idx) &&
              idx >= 0 &&
              idx < workbook.SheetNames.length
            ) {
              sheetName = workbook.SheetNames[idx];
            } else {
              return {
                stdout: "",
                stderr: `Sheet not found: ${sheetArg}. Available: ${workbook.SheetNames.join(", ")}`,
                exitCode: 1,
              };
            }
          }

          const sheet = workbook.Sheets[sheetName];
          if (!sheet) {
            return {
              stdout: "",
              stderr: `Sheet "${sheetName}" not found`,
              exitCode: 1,
            };
          }

          const csv = XLSX.utils.sheet_to_csv(sheet);
          await writeVfsOutput(ctx, outFile, csv);

          return {
            stdout: `Converted sheet "${sheetName}" → ${outFile}`,
            stderr: "",
            exitCode: 0,
          };
        }

        // No sheet specified: export all
        const names = workbook.SheetNames;

        if (names.length === 1) {
          const csv = XLSX.utils.sheet_to_csv(workbook.Sheets[names[0]]);
          await writeVfsOutput(ctx, outFile, csv);
          return {
            stdout: `Converted sheet "${names[0]}" → ${outFile}`,
            stderr: "",
            exitCode: 0,
          };
        }

        // Multiple sheets: <base>.<sheetName>.csv
        const dotIdx = outFile.lastIndexOf(".");
        const base = dotIdx > 0 ? outFile.substring(0, dotIdx) : outFile;
        const ext = dotIdx > 0 ? outFile.substring(dotIdx) : ".csv";
        const outputs: string[] = [];

        for (const name of names) {
          const sheet = workbook.Sheets[name];
          if (!sheet) continue;
          const csv = XLSX.utils.sheet_to_csv(sheet);
          const safeName = name.replace(/[/\\?*[\]]/g, "_");
          const path = `${base}.${safeName}${ext}`;
          await writeVfsOutput(ctx, path, csv);
          outputs.push(`  "${name}" → ${path}`);
        }

        return {
          stdout: `Converted ${names.length} sheets:\n${outputs.join("\n")}`,
          stderr: "",
          exitCode: 0,
        };
      } catch (error) {
        const msg = error instanceof Error ? error.message : String(error);
        return { stdout: "", stderr: msg, exitCode: 1 };
      }
    }),
};

async function decodeImagePixels(
  data: Uint8Array,
  targetW: number,
  targetH: number,
): Promise<Uint8ClampedArray> {
  const blob = new Blob([data as BlobPart]);
  const url = URL.createObjectURL(blob);
  try {
    const img = new Image();
    img.src = url;
    await img.decode();

    const canvas = new OffscreenCanvas(targetW, targetH);
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("Failed to create canvas 2D context");

    ctx.drawImage(img, 0, 0, targetW, targetH);
    return ctx.getImageData(0, 0, targetW, targetH).data;
  } finally {
    URL.revokeObjectURL(url);
  }
}

function rgbToHex(r: number, g: number, b: number): string {
  return `#${r.toString(16).padStart(2, "0")}${g.toString(16).padStart(2, "0")}${b.toString(16).padStart(2, "0")}`;
}

async function paintPixelsToSheet(
  sheetId: number,
  startCell: string,
  width: number,
  height: number,
  pixels: Uint8ClampedArray,
  cellSize: number,
): Promise<void> {
  const { col: startCol, row: startRow } = parseStartCell(startCell);

  // RLE: merge adjacent same-color pixels into range spans, grouped by color
  const colorRanges = new Map<string, string[]>();
  for (let y = 0; y < height; y++) {
    const rowNum = startRow + y + 1;
    let x = 0;
    while (x < width) {
      const i = (y * width + x) * 4;
      const hex = rgbToHex(pixels[i], pixels[i + 1], pixels[i + 2]);
      const runStart = x;
      x++;
      while (x < width) {
        const j = (y * width + x) * 4;
        if (
          pixels[j] !== pixels[i] ||
          pixels[j + 1] !== pixels[i + 1] ||
          pixels[j + 2] !== pixels[i + 2]
        )
          break;
        x++;
      }
      const rangeAddr =
        runStart === x - 1
          ? `${columnIndexToLetter(startCol + runStart)}${rowNum}`
          : `${columnIndexToLetter(startCol + runStart)}${rowNum}:${columnIndexToLetter(startCol + x - 1)}${rowNum}`;
      let ranges = colorRanges.get(hex);
      if (!ranges) {
        ranges = [];
        colorRanges.set(hex, ranges);
      }
      ranges.push(rangeAddr);
    }
  }

  await Excel.run(async (context) => {
    const sheet = await getWorksheetById(context, sheetId);
    if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);

    const endCol = columnIndexToLetter(startCol + width - 1);
    const endRow = startRow + height;
    const fullRange = sheet.getRange(`${startCell}:${endCol}${endRow}`);
    fullRange.format.columnWidth = cellSize;
    fullRange.format.rowHeight = cellSize;
    const emptyValues: string[][] = Array.from({ length: height }, () =>
      Array.from({ length: width }, () => ""),
    );
    fullRange.values = emptyValues;
    await context.sync();

    const RANGES_PER_BATCH = 1000;
    const entries = [...colorRanges.entries()];
    let queued = 0;

    for (const [color, ranges] of entries) {
      for (let i = 0; i < ranges.length; i += RANGES_PER_BATCH) {
        const batch = ranges.slice(i, i + RANGES_PER_BATCH);
        const areas = sheet.getRanges(batch.join(","));
        areas.format.fill.color = color;
        queued += batch.length;

        if (queued >= RANGES_PER_BATCH) {
          await context.sync();
          queued = 0;
        }
      }
    }

    await context.sync();
  });
}

const imageToSheet: Command = defineCommand(
  "image-to-sheet",
  async (args, ctx) => {
    const positional = args.filter((a) => !a.startsWith("--"));
    const cellSizeArg = args.find((a) => a.startsWith("--cell-size="));
    const cellSize = cellSizeArg
      ? Number.parseFloat(cellSizeArg.split("=")[1])
      : 6;

    if (positional.length < 4) {
      return {
        stdout: "",
        stderr:
          "Usage: image-to-sheet <file> <width> <height> <sheetId> [startCell] [--cell-size=N]\n" +
          "  file       - Path to image file in VFS (png, jpg, gif, webp)\n" +
          "  width      - Target width in pixels (columns)\n" +
          "  height     - Target height in pixels (rows)\n" +
          "  sheetId    - Target worksheet ID\n" +
          "  startCell  - Top-left cell, default A1\n" +
          "  --cell-size - Cell width/height in points (default: 6)\n\n" +
          "Decodes an image, downsamples to target size, and paints it as pixel art\n" +
          "in Excel by setting cell background colors. Cells are resized to squares.\n" +
          "Example: image-to-sheet uploads/logo.png 64 64 1 A1 --cell-size=4\n",
        exitCode: 1,
      };
    }

    const [filePath, widthStr, heightStr, sheetIdStr, startCell = "A1"] =
      positional;
    const targetW = Number.parseInt(widthStr, 10);
    const targetH = Number.parseInt(heightStr, 10);
    const sheetId = Number.parseInt(sheetIdStr, 10);

    if (
      Number.isNaN(targetW) ||
      Number.isNaN(targetH) ||
      targetW < 1 ||
      targetH < 1
    ) {
      return {
        stdout: "",
        stderr: "Width and height must be positive integers",
        exitCode: 1,
      };
    }

    if (targetW > 200 || targetH > 200) {
      return {
        stdout: "",
        stderr:
          "Maximum dimensions: 200×200. Use smaller values for Excel pixel art.",
        exitCode: 1,
      };
    }

    if (Number.isNaN(sheetId)) {
      return {
        stdout: "",
        stderr: `Invalid sheetId: ${sheetIdStr}`,
        exitCode: 1,
      };
    }

    const upperStartCell = startCell.toUpperCase();
    if (!/^[A-Z]+\d+$/.test(upperStartCell)) {
      return {
        stdout: "",
        stderr: `Invalid start cell: ${startCell}`,
        exitCode: 1,
      };
    }

    if (Number.isNaN(cellSize) || cellSize < 1 || cellSize > 50) {
      return {
        stdout: "",
        stderr: "Cell size must be between 1 and 50 points",
        exitCode: 1,
      };
    }

    try {
      const { data } = await resolveVfsPath(ctx, filePath);
      const pixels = await decodeImagePixels(data, targetW, targetH);

      await paintPixelsToSheet(
        sheetId,
        upperStartCell,
        targetW,
        targetH,
        pixels,
        cellSize,
      );

      const uniqueColors = new Set<string>();
      for (let i = 0; i < targetW * targetH; i++) {
        const idx = i * 4;
        uniqueColors.add(
          rgbToHex(pixels[idx], pixels[idx + 1], pixels[idx + 2]),
        );
      }

      return {
        stdout:
          `Painted ${targetW}×${targetH} pixel art (${targetW * targetH} cells, ${uniqueColors.size} colors) ` +
          `from ${filePath} into sheet ${sheetId} at ${upperStartCell} (cell size: ${cellSize}pt)`,
        stderr: "",
        exitCode: 0,
      };
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      return { stdout: "", stderr: msg, exitCode: 1 };
    }
  },
);

const webSearchCmd: Command = defineCommand("web-search", async (args) => {
  const flags: Record<string, string> = {};
  const positional: string[] = [];
  for (const arg of args) {
    const match = arg.match(/^--(\w+)=(.+)$/);
    if (match) {
      flags[match[1]] = match[2];
    } else if (arg === "--json") {
      flags.json = "true";
    } else {
      positional.push(arg);
    }
  }

  const query = positional.join(" ");
  if (!query) {
    return {
      stdout: "",
      stderr:
        "Usage: web-search <query> [--max=N] [--region=REGION] [--time=d|w|m|y] [--page=N] [--json]\n  query    - Search query\n  --max    - Max results (default: 10)\n  --region - Region code, e.g. us-en, uk-en (default: us-en)\n  --time   - Time filter: d(ay), w(eek), m(onth), y(ear)\n  --page   - Page number (default: 1)\n  --json   - Output as JSON\n",
      exitCode: 1,
    };
  }

  try {
    const webConfig = loadWebConfig();
    const results = await searchWeb(
      query,
      {
        maxResults: flags.max ? Number.parseInt(flags.max, 10) : 10,
        region: flags.region,
        timelimit: flags.time as "d" | "w" | "m" | "y" | undefined,
        page: flags.page ? Number.parseInt(flags.page, 10) : undefined,
      },
      {
        proxyUrl: getProxyUrl(),
        apiKeys: webConfig.apiKeys,
      },
      webConfig.searchProvider,
    );

    if (results.length === 0) {
      return { stdout: "No results found.", stderr: "", exitCode: 0 };
    }

    if (flags.json === "true") {
      return {
        stdout: JSON.stringify(results, null, 2),
        stderr: "",
        exitCode: 0,
      };
    }

    const lines = results.map(
      (r, i) => `${i + 1}. ${r.title}\n   ${r.href}\n   ${r.body}`,
    );
    return {
      stdout: lines.join("\n\n"),
      stderr: "",
      exitCode: 0,
    };
  } catch (error) {
    const msg = error instanceof Error ? error.message : String(error);
    return { stdout: "", stderr: msg, exitCode: 1 };
  }
});

function getProxyUrl(): string | undefined {
  const config = loadSavedConfig();
  return config?.useProxy && config?.proxyUrl ? config.proxyUrl : undefined;
}

const webFetchCmd: Command = defineCommand("web-fetch", async (args, ctx) => {
  const url = args[0];
  const outFile = args[1];
  if (!url || !outFile) {
    return {
      stdout: "",
      stderr:
        "Usage: web-fetch <url> <outfile>\n  url      - URL to fetch\n  outfile  - Output file path\n\nFetches a URL and saves to a file.\n  - HTML pages: extracts readable content (Markdown)\n  - Binary files (PDF, DOCX, XLSX, etc.): downloads raw file\n  - Text/JSON/XML: saves as-is\n",
      exitCode: 1,
    };
  }

  try {
    const webConfig = loadWebConfig();
    const result = await fetchWeb(
      url,
      {
        proxyUrl: getProxyUrl(),
        apiKeys: webConfig.apiKeys,
      },
      webConfig.fetchProvider,
    );

    if (result.kind === "text") {
      const header = [
        result.title ? `Title: ${result.title}` : "",
        ...Object.entries(result.metadata || {}).map(([k, v]) => `${k}: ${v}`),
      ]
        .filter(Boolean)
        .join("\n");
      const output = header ? `${header}\n\n${result.text}` : result.text;

      await writeVfsOutput(ctx, outFile, output);
      return {
        stdout: `Fetched text → ${outFile} (${result.text.length} chars, ${result.contentType})`,
        stderr: "",
        exitCode: 0,
      };
    }

    const resolvedPath = outFile.startsWith("/")
      ? outFile
      : `${ctx.cwd}/${outFile}`;
    const dir = resolvedPath.substring(0, resolvedPath.lastIndexOf("/"));
    if (dir && dir !== "/") {
      try {
        await ctx.fs.mkdir(dir, { recursive: true });
      } catch {
        // directory may already exist
      }
    }
    await ctx.fs.writeFile(resolvedPath, result.data);

    const size =
      result.data.length >= 1024
        ? `${Math.round(result.data.length / 1024)}KB`
        : `${result.data.length}B`;

    return {
      stdout: `Downloaded → ${outFile} (${size}, ${result.contentType || "unknown type"})`,
      stderr: "",
      exitCode: 0,
    };
  } catch (error) {
    const msg = error instanceof Error ? error.message : String(error);
    return { stdout: "", stderr: msg, exitCode: 1 };
  }
});

export function getCustomCommands(): CustomCommand[] {
  return [
    csvToSheet,
    sheetToCsv,
    pdfToText,
    pdfToImages,
    docxToText,
    xlsxToCsv,
    imageToSheet,
    webSearchCmd,
    webFetchCmd,
  ];
}
