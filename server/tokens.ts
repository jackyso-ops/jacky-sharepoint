/**
 * Local-first token store (JSON file). Production later: monday-code Secure Storage.
 */

import fs from "node:fs";
import path from "node:path";

export type MicrosoftTokenRecord = {
  accessToken: string;
  refreshToken?: string;
  expiresAt: number;
  scope?: string;
};

function storePath(): string {
  return (
    process.env.TOKEN_STORE_PATH ||
    path.join(process.cwd(), ".tokens", "microsoft-tokens.json")
  );
}

function ensureDir(filePath: string): void {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
}

function readAll(): Record<string, MicrosoftTokenRecord> {
  const file = storePath();
  if (!fs.existsSync(file)) return {};
  try {
    const raw = fs.readFileSync(file, "utf8");
    return JSON.parse(raw) as Record<string, MicrosoftTokenRecord>;
  } catch {
    return {};
  }
}

function writeAll(data: Record<string, MicrosoftTokenRecord>): void {
  const file = storePath();
  ensureDir(file);
  fs.writeFileSync(file, JSON.stringify(data, null, 2), "utf8");
}

export function tokenKey(accountId: string, userId: string): string {
  return `ms-graph:${accountId}:${userId}`;
}

export async function saveMicrosoftToken(
  accountId: string,
  userId: string,
  record: MicrosoftTokenRecord
): Promise<void> {
  const all = readAll();
  all[tokenKey(accountId, userId)] = record;
  writeAll(all);
}

export async function getMicrosoftToken(
  accountId: string,
  userId: string
): Promise<MicrosoftTokenRecord | null> {
  const all = readAll();
  return all[tokenKey(accountId, userId)] ?? null;
}

export async function deleteMicrosoftToken(
  accountId: string,
  userId: string
): Promise<void> {
  const all = readAll();
  delete all[tokenKey(accountId, userId)];
  writeAll(all);
}
