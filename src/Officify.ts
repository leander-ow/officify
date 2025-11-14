import JSZip from "jszip";
import { spawn } from "child_process";
import fs from "fs/promises";
import fsSync from "fs";
import os from "os";
import path from "path";
import crypto from "crypto";
import { escapeRegExp } from "./utils.js";
import type { PlaceholderMap, ImageMap, OfficifyOptions } from "./types.js";
import { resolveSofficePath } from "./sofficeResolver.js";

async function guessExtensionFromManifest(
  zip: JSZip
): Promise<string | undefined> {
  const manifestFile = zip.file("META-INF/manifest.xml");
  if (!manifestFile) return undefined;
  try {
    const content = await manifestFile.async("string");
    if (content.includes("vnd.oasis.opendocument.text")) return ".odt";
    if (content.includes("vnd.oasis.opendocument.spreadsheet")) return ".ods";
    if (content.includes("vnd.oasis.opendocument.presentation")) return ".odp";
  } catch {
    // ignore
  }
  return undefined;
}

/**
 * Officify class
 */
export class Officify {
  private zip!: JSZip;
  private sofficePath: string;
  constructor(
    private inputBuffer: Buffer,
    private options: OfficifyOptions = {}
  ) {
    this.zip = new JSZip();
    this.sofficePath = resolveSofficePath(options.sofficePath);
  }

  async load(): Promise<void> {
    this.zip = await JSZip.loadAsync(this.inputBuffer);
  }

  private ensureLoaded() {
    if (!this.zip) throw new Error("Archive not loaded. Call load() first.");
  }

  async replacePlaceholders(placeholders: PlaceholderMap): Promise<void> {
    this.ensureLoaded();
    const keys = Object.keys(placeholders);
    if (keys.length === 0) return;

    // Fix: ensure value is a string (not string | undefined) by defaulting to empty string
    const map: Array<{ re: RegExp; value: string }> = keys.map((k) => {
      const raw = placeholders[k as keyof PlaceholderMap];
      const value = raw === undefined ? "" : raw;
      return { re: new RegExp(escapeRegExp(k), "g"), value };
    });

    const fileNames = Object.keys(this.zip.files);
    for (const fileName of fileNames) {
      if (!fileName.match(/\.(xml|opf|rels|html|htm|xhtml|svg|css|txt)$/i))
        continue;
      const file = this.zip.file(fileName);
      if (!file) continue;
      let content = await file.async("string");
      for (const { re, value } of map) content = content.replace(re, value);
      this.zip.file(fileName, content);
    }
  }

  replaceImages(images: ImageMap): void {
    this.ensureLoaded();
    for (const [name, buffer] of Object.entries(images)) {
      const zipPath = `Pictures/${name}`;
      this.zip.file(zipPath, buffer);
    }
  }

  async getBuffer(): Promise<Buffer> {
    this.ensureLoaded();
    return Buffer.from(
      await this.zip.generateAsync({
        type: "nodebuffer",
        compression: "DEFLATE",
      })
    );
  }

  /**
   * Export to PDF. Determines a sensible input extension from:
   * 1. options.inputExtensionHint
   * 2. manifest mime-type inside the archive
   * 3. default to .odt
   */
  async exportPDF(pages?: string): Promise<Buffer> {
    this.ensureLoaded();
    const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "officify-"));
    const randomPart = crypto.randomBytes(8).toString("hex");
    let ext = this.options.inputExtensionHint;
    if (ext && !ext.startsWith(".")) ext = "." + ext;
    if (!ext) {
      const guessed = await guessExtensionFromManifest(this.zip);
      if (guessed) ext = guessed;
    }
    if (!ext) ext = ".odt";

    const inputFileName = `input-${Date.now()}-${randomPart}${ext}`;
    const inputFile = path.join(tmpDir, inputFileName);

    try {
      const buffer = await this.getBuffer();
      await fs.writeFile(inputFile, buffer, { mode: 0o600 });

      const outputFileName = inputFileName.replace(/\.[^/.]+$/, ".pdf");
      const outputFile = path.join(tmpDir, outputFileName);

      await new Promise<void>((resolve, reject) => {
        const soffice = spawn(this.sofficePath, [
          "--headless",
          "--convert-to",
          pages
            ? `pdf:writer_pdf_Export:{"PageRange":{"type":"string","value":"${pages}"}}`
            : "pdf",
          "--outdir",
          tmpDir,
          inputFile,
        ]);

        let stdout = "";
        let stderr = "";

        soffice.stdout.on("data", (chunk) => (stdout += chunk.toString()));
        soffice.stderr.on("data", (chunk) => (stderr += chunk.toString()));

        soffice.on("error", (err) => {
          reject(
            new Error(
              `Failed to start soffice (${this.sofficePath}): ${err.message}`
            )
          );
        });

        soffice.on("close", (code, signal) => {
          if (code === 0) {
            if (fsSync.existsSync(outputFile)) resolve();
            else
              reject(
                new Error(
                  `LibreOffice finished with code 0 but no output. stdout: ${stdout} stderr: ${stderr}`
                )
              );
          } else {
            reject(
              new Error(
                `LibreOffice conversion failed (code:${code} signal:${signal}). stdout: ${stdout} stderr: ${stderr}`
              )
            );
          }
        });
      });

      const pdfBuffer = await fs.readFile(path.join(tmpDir, outputFileName));
      return pdfBuffer;
    } finally {
      try {
        await fs.rm(tmpDir, { recursive: true, force: true });
      } catch {
        // ignore cleanup errors
      }
    }
  }
}
