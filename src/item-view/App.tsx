/**
 * monday item-view — SharePoint folder browser (view-only).
 *
 * Board: HCA (5032301117)
 * Column: text_mm90r6dr
 *
 * Renders a custom list UI. Never iframes AllItems.aspx / SharePoint pages.
 */

import mondaySdk from "monday-sdk-js";
import { useCallback, useEffect, useState } from "react";
import {
  folderUrlToShareId,
  normalizeSharePointFolderUrl,
} from "../lib/sharepoint-url";

const URL_COLUMN_ID = "text_mm90r6dr";
const BOARD_ID = "5032301117";

const monday = mondaySdk();

type ChildItem = {
  id: string;
  name: string;
  webUrl?: string;
  folder?: object;
  file?: object;
  size?: number;
  lastModifiedDateTime?: string;
  parentReference?: { driveId?: string };
};

type Crumb = { id: string; name: string; driveId?: string };

type FolderState =
  | { status: "loading" }
  | { status: "empty"; message: string }
  | { status: "need_auth" }
  | { status: "error"; message: string }
  | {
      status: "ready";
      name: string;
      driveId?: string;
      children: ChildItem[];
      breadcrumb: Crumb[];
    };

type MondayContext = {
  accountId?: string | number;
  user?: { id?: string | number };
  userId?: string | number;
  itemId?: string | number;
  boardId?: string | number;
};

async function apiFetch(path: string, init?: RequestInit): Promise<Response> {
  return fetch(path, {
    credentials: "include",
    ...init,
    headers: {
      Accept: "application/json",
      ...(init?.headers || {}),
    },
  });
}

export default function App() {
  const [folderUrl, setFolderUrl] = useState<string | null>(null);
  const [ctx, setCtx] = useState<{ accountId: string; userId: string } | null>(
    null
  );
  const [connected, setConnected] = useState(false);
  const [state, setState] = useState<FolderState>({ status: "loading" });
  const [driveId, setDriveId] = useState<string | undefined>();

  const persistContext = useCallback(async (accountId: string, userId: string) => {
    setCtx({ accountId, userId });
    await apiFetch("/api/auth/context", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ accountId, userId }),
    });
  }, []);

  const refreshAuthStatus = useCallback(async () => {
    const res = await apiFetch("/api/auth/status");
    if (!res.ok) {
      setConnected(false);
      return false;
    }
    const data = (await res.json()) as { connected?: boolean };
    setConnected(Boolean(data.connected));
    return Boolean(data.connected);
  }, []);

  const loadColumnUrl = useCallback(async () => {
    setState({ status: "loading" });
    try {
      const contextRes = await monday.get("context");
      const context = (contextRes?.data || contextRes) as MondayContext;
      const accountId = String(context.accountId ?? "").trim();
      const userId = String(context.user?.id ?? context.userId ?? "").trim();
      const itemId = String(context.itemId ?? "").trim();

      if (accountId && userId) {
        await persistContext(accountId, userId);
      }

      await refreshAuthStatus();

      if (!itemId) {
        // Outside monday — allow empty state for local smoke
        setFolderUrl(null);
        setState({
          status: "empty",
          message:
            "No SharePoint folder linked (open this view on an HCA item, or paste a URL in monday column text_mm90r6dr).",
        });
        return;
      }

      const query = `query ($itemIds: [ID!]) {
        items(ids: $itemIds) {
          id
          name
          column_values(ids: ["${URL_COLUMN_ID}"]) {
            id
            text
            value
          }
        }
      }`;
      const apiRes = await monday.api(query, { variables: { itemIds: [itemId] } });
      const item = apiRes?.data?.items?.[0];
      const col = item?.column_values?.find(
        (c: { id: string }) => c.id === URL_COLUMN_ID
      );
      const raw = (col?.text || "").trim();
      if (!raw) {
        setFolderUrl(null);
        setState({ status: "empty", message: "No SharePoint folder linked" });
        return;
      }
      const clean = normalizeSharePointFolderUrl(raw) || raw;
      setFolderUrl(clean);
      void BOARD_ID;
    } catch (e) {
      setState({
        status: "error",
        message:
          e instanceof Error
            ? e.message
            : "Failed to read monday context / column",
      });
    }
  }, [persistContext, refreshAuthStatus]);

  const openRoot = useCallback(async () => {
    if (!folderUrl) {
      setState({ status: "empty", message: "No SharePoint folder linked" });
      return;
    }
    const shareId = folderUrlToShareId(folderUrl);
    if (!shareId) {
      setState({ status: "error", message: "Could not normalize SharePoint URL" });
      return;
    }
    setState({ status: "loading" });
    try {
      const res = await apiFetch(
        `/api/folder?shareId=${encodeURIComponent(shareId)}`
      );
      if (res.status === 401) {
        setConnected(false);
        setState({ status: "need_auth" });
        return;
      }
      if (!res.ok) {
        const body = await res.json().catch(async () => ({
          message: await res.text(),
        }));
        setState({
          status: "error",
          message: body.message || `Folder load failed (${res.status})`,
        });
        return;
      }
      const data = await res.json();
      setDriveId(data.driveId);
      setState({
        status: "ready",
        name: data.name ?? "Folder",
        driveId: data.driveId,
        children: sortChildren(data.children ?? []),
        breadcrumb: [
          { id: data.id, name: data.name ?? "Folder", driveId: data.driveId },
        ],
      });
    } catch (e) {
      setState({
        status: "error",
        message: e instanceof Error ? e.message : "Request failed",
      });
    }
  }, [folderUrl]);

  useEffect(() => {
    void loadColumnUrl();
  }, [loadColumnUrl]);

  useEffect(() => {
    if (folderUrl) {
      void openRoot();
    }
  }, [folderUrl, openRoot]);

  const connectMicrosoft = () => {
    const q = new URLSearchParams();
    if (ctx?.accountId) q.set("accountId", ctx.accountId);
    if (ctx?.userId) q.set("userId", ctx.userId);
    const qs = q.toString();
    window.open(
      `/auth/microsoft/start${qs ? `?${qs}` : ""}`,
      "_blank",
      "noopener,noreferrer"
    );
  };

  const disconnectMicrosoft = async () => {
    await apiFetch("/auth/microsoft/disconnect", { method: "POST" });
    setConnected(false);
    setState({ status: "need_auth" });
  };

  const onFileClick = (item: ChildItem) => {
    if (item.webUrl) window.open(item.webUrl, "_blank", "noopener,noreferrer");
  };

  const onFolderClick = async (item: ChildItem) => {
    const dId = driveId || item.parentReference?.driveId;
    if (!dId) {
      setState({
        status: "error",
        message: "Missing driveId for drill-down",
      });
      return;
    }
    setState({ status: "loading" });
    try {
      const res = await apiFetch(
        `/api/folder/children?driveId=${encodeURIComponent(dId)}&driveItemId=${encodeURIComponent(item.id)}`
      );
      if (res.status === 401) {
        setConnected(false);
        setState({ status: "need_auth" });
        return;
      }
      if (!res.ok) {
        const body = await res.json().catch(async () => ({
          message: await res.text(),
        }));
        setState({
          status: "error",
          message: body.message || `Children load failed (${res.status})`,
        });
        return;
      }
      const data = await res.json();
      setState((prev) => {
        const crumb =
          prev.status === "ready"
            ? [...prev.breadcrumb, { id: item.id, name: item.name, driveId: dId }]
            : [{ id: item.id, name: item.name, driveId: dId }];
        return {
          status: "ready",
          name: item.name,
          driveId: dId,
          children: sortChildren(data.children ?? []),
          breadcrumb: crumb,
        };
      });
    } catch (e) {
      setState({
        status: "error",
        message: e instanceof Error ? e.message : "Request failed",
      });
    }
  };

  const onBreadcrumbClick = async (index: number) => {
    if (state.status !== "ready") return;
    if (index === 0) {
      await openRoot();
      return;
    }
    const crumbs = state.breadcrumb;
    const target = crumbs[index];
    if (!target) return;
    const dId = target.driveId || driveId;
    if (!dId) {
      await openRoot();
      return;
    }
    setState({ status: "loading" });
    try {
      const res = await apiFetch(
        `/api/folder/children?driveId=${encodeURIComponent(dId)}&driveItemId=${encodeURIComponent(target.id)}`
      );
      if (res.status === 401) {
        setState({ status: "need_auth" });
        return;
      }
      if (!res.ok) {
        const body = await res.json().catch(async () => ({
          message: await res.text(),
        }));
        setState({
          status: "error",
          message: body.message || `Breadcrumb load failed (${res.status})`,
        });
        return;
      }
      const data = await res.json();
      setState({
        status: "ready",
        name: target.name,
        driveId: dId,
        children: sortChildren(data.children ?? []),
        breadcrumb: crumbs.slice(0, index + 1),
      });
    } catch (e) {
      setState({
        status: "error",
        message: e instanceof Error ? e.message : "Request failed",
      });
    }
  };

  return (
    <div className="sp-folder-browser">
      <header>
        <h1>SharePoint folder</h1>
        {folderUrl ? (
          <a href={folderUrl} target="_blank" rel="noreferrer">
            Open in SharePoint
          </a>
        ) : null}
      </header>

      <div className="sp-toolbar">
        {!connected ? (
          <button type="button" className="primary" onClick={connectMicrosoft}>
            Connect Microsoft
          </button>
        ) : (
          <button type="button" className="danger" onClick={() => void disconnectMicrosoft()}>
            Disconnect
          </button>
        )}
        <button type="button" onClick={() => void loadColumnUrl()}>
          Refresh URL
        </button>
        {folderUrl ? (
          <button type="button" onClick={() => void openRoot()}>
            Reload root
          </button>
        ) : null}
        <button type="button" onClick={() => void refreshAuthStatus()}>
          Check auth
        </button>
      </div>

      {state.status === "need_auth" ? (
        <div className="sp-status">
          Connect Microsoft to browse this folder.{" "}
          <button type="button" className="primary" onClick={connectMicrosoft}>
            Connect Microsoft
          </button>
        </div>
      ) : null}

      {state.status === "empty" ? (
        <div className="sp-status">{state.message}</div>
      ) : null}
      {state.status === "error" ? (
        <div className="sp-status" role="alert">
          {state.message}
        </div>
      ) : null}
      {state.status === "loading" ? (
        <div className="sp-status">Loading…</div>
      ) : null}

      {state.status === "ready" ? (
        <>
          <nav className="sp-breadcrumb" aria-label="Breadcrumb">
            {state.breadcrumb.map((c, i) => (
              <span key={`${c.id}-${i}`}>
                {i > 0 ? " / " : null}
                <button type="button" onClick={() => void onBreadcrumbClick(i)}>
                  {c.name}
                </button>
              </span>
            ))}
          </nav>
          {state.children.length === 0 ? (
            <div className="sp-status">This folder is empty.</div>
          ) : (
            <ul className="sp-list">
              {state.children.map((child) => (
                <li key={child.id}>
                  {child.folder ? (
                    <button
                      type="button"
                      className="row"
                      onClick={() => void onFolderClick(child)}
                    >
                      <span className="sp-icon folder" aria-hidden>
                        ▣
                      </span>
                      <span>{child.name}</span>
                    </button>
                  ) : (
                    <button
                      type="button"
                      className="row"
                      onClick={() => onFileClick(child)}
                    >
                      <span className="sp-icon" aria-hidden>
                        ▭
                      </span>
                      <span>{child.name}</span>
                      {child.size != null ? (
                        <span className="sp-meta">{formatSize(child.size)}</span>
                      ) : null}
                    </button>
                  )}
                </li>
              ))}
            </ul>
          )}
        </>
      ) : null}

      <p className="sp-hint">
        View-only · HCA column <code>{URL_COLUMN_ID}</code> · custom UI (no
        SharePoint iframe)
      </p>
    </div>
  );
}

function sortChildren(items: ChildItem[]): ChildItem[] {
  return [...items].sort((a, b) => {
    const af = a.folder ? 0 : 1;
    const bf = b.folder ? 0 : 1;
    if (af !== bf) return af - bf;
    return a.name.localeCompare(b.name);
  });
}

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}
