import fs from "fs";
import path from "path";
import os from "os";

/**
 * Resolve a soffice executable path.
 * Priority:
 * 1. explicitPath argument
 * 2. process.env.OFFICIFY_SOFFICE_PATH
 * 3. common platform-specific locations
 * 4. fallback to "soffice" (rely on PATH)
 */
export function resolveSofficePath(explicitPath?: string): string {
  const candidates: string[] = [];

  if (explicitPath) candidates.push(explicitPath);
  if (process.env.OFFICIFY_SOFFICE_PATH) candidates.push(process.env.OFFICIFY_SOFFICE_PATH);

  const platform = os.platform();
  if (platform === "win32") {
    // common Windows install locations
    candidates.push(
      "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
      "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe"
    );
  } else if (platform === "darwin") {
    // macOS app bundle
    candidates.push(
      "/Applications/LibreOffice.app/Contents/MacOS/soffice",
      "/usr/local/bin/soffice"
    );
  } else {
    // linux/unix
    candidates.push("/usr/bin/soffice", "/usr/local/bin/soffice", "/snap/bin/soffice");
  }

  // Finally rely on PATH
  candidates.push("soffice");

  // return the first candidate that exists and is executable (except "soffice" which we assume in PATH)
  for (const c of candidates) {
    if (c === "soffice") return c;
    try {
      const stat = fs.statSync(c);
      if (stat && (stat.mode & 0o111) !== 0) return c; // executable bit set
    } catch {
      // ignore
    }
  }
  // fallback
  return "soffice";
}