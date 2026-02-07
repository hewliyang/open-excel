import type { Command, CustomCommand } from "just-bash/browser";
import { defineCommand } from "just-bash/browser";
import type { CellInput } from "../excel/api";
import { getRangeAsCsv, getWorksheetById, setCellRange } from "../excel/api";

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

function buildRangeAddress(startCell: string, rows: number, cols: number): string {
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
    return { stdout: "", stderr: `Invalid sheetId: ${sheetIdStr}`, exitCode: 1 };
  }

  const upperStartCell = startCell.toUpperCase();
  if (!/^[A-Z]+\d+$/.test(upperStartCell)) {
    return { stdout: "", stderr: `Invalid start cell: ${startCell}`, exitCode: 1 };
  }

  try {
    const resolvedPath = filePath.startsWith("/") ? filePath : `${ctx.cwd}/${filePath}`;
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
    const result = await setCellRange(sheetId, rangeAddr, cells, { allowOverwrite: force });

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
    return { stdout: "", stderr: `Invalid sheetId: ${sheetIdStr}`, exitCode: 1 };
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
        return { stdout: "", stderr: "Sheet is empty (no used range)", exitCode: 1 };
      }
      rangeAddr = usedAddr;
    }

    const result = await getRangeAsCsv(sheetId, rangeAddr, { maxRows: 50000 });

    if (outFile) {
      const resolvedPath = outFile.startsWith("/") ? outFile : `${ctx.cwd}/${outFile}`;
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
      const moreNote = result.hasMore ? " (truncated, more rows available)" : "";
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
  const resolved = filePath.startsWith("/") ? filePath : `${ctx.cwd}/${filePath}`;
  const data = await ctx.fs.readFileBuffer(resolved);
  return { path: resolved, data };
}

async function writeVfsOutput(
  ctx: {
    cwd: string;
    fs: { mkdir(p: string, o: { recursive: boolean }): Promise<void>; writeFile(p: string, c: string): Promise<void> };
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
        const result = await mammoth.extractRawText({ arrayBuffer: data.buffer as ArrayBuffer });

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
            if (!Number.isNaN(idx) && idx >= 0 && idx < workbook.SheetNames.length) {
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
            return { stdout: "", stderr: `Sheet "${sheetName}" not found`, exitCode: 1 };
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

export function getCustomCommands(): CustomCommand[] {
  return [csvToSheet, sheetToCsv, pdfToText, docxToText, xlsxToCsv];
}
