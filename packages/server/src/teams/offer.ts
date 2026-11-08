// Files an agent sends with teams_reply in a direct chat, by Teams' bot file consent: a consent
// card per file; the person's Accept arrives as a fileConsent/invoke carrying an upload address
// in their own OneDrive; the bot puts the file there and posts a file card. docs/teams.md 6.7.

export const OFFER_HOURS = 24;
export const MAX_OFFER_FILES = 10;
// All the files of one reply together; each waits in the database until answered.
export const MAX_OFFER_BYTES = 25 * 1024 * 1024;

/** A file the agent named that cannot be sent; the reason is for the agent. */
export class FileRefused extends Error {}

// SharePoint and OneDrive for work and school: where a file's short-lived download link and a
// consent card's upload address both are. Nothing else is fetched or written to.
const FILE_HOST = /^[a-z0-9-]+\.sharepoint\.com$/i;

export function fileStorageUrl(value: unknown): string | null {
  if (typeof value !== "string") return null;
  try {
    const u = new URL(value);
    const ok =
      u.protocol === "https:" && !u.username && !u.password && !u.port && FILE_HOST.test(u.hostname);
    return ok ? u.toString() : null;
  } catch {
    return null;
  }
}

/** A name OneDrive takes as is. */
export function offerName(name: string): string {
  return (
    name
      .replace(/[\\/:*?"<>|\u0000-\u001f]/g, "_")
      .trim()
      .slice(0, 255) || "file"
  );
}

export const offerAad = (id: string) => `agent-connect:teams-file-offer:${id}`;

export function consentCard(offerId: string, name: string, size: number) {
  return {
    type: "message",
    attachments: [
      {
        contentType: "application/vnd.microsoft.teams.card.file.consent",
        name,
        content: {
          description: "Accept to save this file to your OneDrive.",
          sizeInBytes: size,
          acceptContext: { offer: offerId },
          declineContext: { offer: offerId },
        },
      },
    ],
  };
}

export type UploadInfo = {
  name: string;
  uploadUrl: string;
  contentUrl: string;
  uniqueId: string;
  fileType: string;
};

/** The person's answer to a consent card, from the invoke's value. */
export function consentAnswer(value: unknown): {
  action: "accept" | "decline";
  offer: string | null;
  upload: UploadInfo | null;
} | null {
  const v = (value ?? {}) as {
    action?: unknown;
    context?: { offer?: unknown };
    uploadInfo?: Record<string, unknown>;
  };
  if (v.action !== "accept" && v.action !== "decline") return null;
  const offer = typeof v.context?.offer === "string" ? v.context.offer : null;
  const u = v.uploadInfo ?? {};
  const str = (k: string) => (typeof u[k] === "string" ? (u[k] as string) : "");
  const upload =
    str("uploadUrl") && str("contentUrl") && str("uniqueId")
      ? {
          name: str("name"),
          uploadUrl: str("uploadUrl"),
          contentUrl: str("contentUrl"),
          uniqueId: str("uniqueId"),
          fileType: str("fileType"),
        }
      : null;
  return { action: v.action, offer, upload };
}

export function fileInfoCard(upload: UploadInfo, name: string) {
  return {
    type: "message",
    attachments: [
      {
        contentType: "application/vnd.microsoft.teams.card.file.info",
        contentUrl: upload.contentUrl,
        name: upload.name || name,
        content: { uniqueId: upload.uniqueId, fileType: upload.fileType },
      },
    ],
  };
}

/** A file's bytes from its download link, refused past `max` bytes without reading further. */
export async function downloadFile(
  fetchImpl: typeof fetch,
  url: string,
  max: number,
): Promise<Buffer> {
  const at = fileStorageUrl(url);
  if (!at)
    throw new FileRefused("Give a download link to a file in SharePoint or OneDrive");
  const res = await fetchImpl(at, { redirect: "error" });
  if (!res.ok) {
    await res.body?.cancel();
    throw new FileRefused(
      `The download link answered ${res.status}; links work for about an hour, so get a new one`,
    );
  }
  const tooBig = () =>
    new FileRefused(
      `The files of one reply are at most ${MAX_OFFER_BYTES / 1024 / 1024} MB together`,
    );
  if (Number(res.headers.get("content-length") ?? 0) > max) {
    await res.body?.cancel();
    throw tooBig();
  }
  const chunks: Uint8Array[] = [];
  let size = 0;
  for await (const chunk of (res.body ?? []) as unknown as AsyncIterable<Uint8Array>) {
    size += chunk.length;
    if (size > max) throw tooBig();
    chunks.push(chunk);
  }
  if (!size) throw new FileRefused("That file is empty");
  return Buffer.concat(chunks);
}
