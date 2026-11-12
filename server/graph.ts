/**
 * Microsoft Graph — view-only folder browsing.
 *
 * Primary: GET /shares/{shareId}/driveItem?$expand=children
 * Drill-down: GET /drives/{driveId}/items/{itemId}/children
 */

const GRAPH = "https://graph.microsoft.com/v1.0";
const CHILD_SELECT =
  "id,name,folder,file,webUrl,size,lastModifiedDateTime,parentReference";

export type GraphDriveItem = {
  id: string;
  name: string;
  webUrl?: string;
  folder?: { childCount?: number };
  file?: { mimeType?: string };
  size?: number;
  lastModifiedDateTime?: string;
  parentReference?: { driveId?: string };
  children?: GraphDriveItem[];
};

async function graphGet<T>(accessToken: string, path: string): Promise<T> {
  const res = await fetch(`${GRAPH}${path}`, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: "application/json",
    },
  });
  if (!res.ok) {
    const text = await res.text();
    const err = new Error(`Graph ${res.status}: ${text}`);
    (err as Error & { status: number }).status = res.status;
    throw err;
  }
  return (await res.json()) as T;
}

/** Resolve a Graph share id (u!…) to driveItem + children. */
export async function getShareRootWithChildren(
  accessToken: string,
  shareId: string
): Promise<GraphDriveItem> {
  const encoded = encodeURIComponent(shareId);
  return graphGet<GraphDriveItem>(
    accessToken,
    `/shares/${encoded}/driveItem?$expand=children($select=${CHILD_SELECT})&$select=id,name,webUrl,folder,file,parentReference`
  );
}

/** List children of a drive item (subfolder drill-down). */
export async function listChildren(
  accessToken: string,
  driveId: string,
  itemId: string
): Promise<GraphDriveItem[]> {
  const data = await graphGet<{ value: GraphDriveItem[] }>(
    accessToken,
    `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/children?$select=${CHILD_SELECT}`
  );
  return data.value ?? [];
}
