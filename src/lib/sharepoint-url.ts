/**
 * Normalize SharePoint library view URLs (AllItems.aspx?id=…) into a folder web URL,
 * then encode for Microsoft Graph GET /shares/{shareId}.
 *
 * Sample:
 *   https://dapulse.sharepoint.com/sites/.../Forms/AllItems.aspx?...&id=%2Fsites%2F...
 */

export function normalizeSharePointFolderUrl(raw: string): string | null {
  const trimmed = (raw || "").trim();
  if (!trimmed) return null;

  let url: URL;
  try {
    url = new URL(trimmed);
  } catch {
    return null;
  }

  const id = url.searchParams.get("id");
  if (id) {
    // Prefer the server-relative folder path from id= over AllItems.aspx chrome.
    const path = decodeURIComponent(id);
    const origin = `${url.protocol}//${url.host}`;
    const normalizedPath = path.startsWith("/") ? path : `/${path}`;
    return `${origin}${normalizedPath}`;
  }

  // Already a folder/file webUrl — strip hash/noise.
  url.hash = "";
  return url.toString();
}

/** Graph sharing token: base64url(UTF-8 URL) prefixed with u! */
export function toGraphShareId(absoluteFolderUrl: string): string {
  const bytes = new TextEncoder().encode(absoluteFolderUrl);
  let binary = "";
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]!);
  }
  const b64 =
    typeof btoa === "function"
      ? btoa(binary)
      : Buffer.from(bytes).toString("base64");
  const base64url = b64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
  return `u!${base64url}`;
}

export function folderUrlToShareId(rawColumnValue: string): string | null {
  const folderUrl = normalizeSharePointFolderUrl(rawColumnValue);
  if (!folderUrl) return null;
  return toGraphShareId(folderUrl);
}
