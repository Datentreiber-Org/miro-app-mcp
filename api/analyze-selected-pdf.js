export default async function handler(req, res) {
  // --- Quick & Dirty CORS ---
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") {
    res.status(204).end();
    return;
  }

  if (req.method !== "POST") {
    res.status(405).send("Use POST.");
    return;
  }

  const MIRO_ACCESS_TOKEN = process.env.MIRO_ACCESS_TOKEN || "";
  if (!MIRO_ACCESS_TOKEN) {
    res.status(500).send("Server misconfigured: MIRO_ACCESS_TOKEN is missing.");
    return;
  }

  // Robust JSON body parse (Vercel Node Functions sind nicht immer automatisch geparst)
  const body = await readJson(req).catch(() => null);
  if (!body) {
    res.status(400).send("Invalid JSON body.");
    return;
  }

  const { boardId, itemId, openaiKey, model, prompt } = body;

  if (!boardId || !itemId) {
    res.status(400).send("boardId or itemId missing.");
    return;
  }
  if (!openaiKey) {
    res.status(400).send("openaiKey missing.");
    return;
  }
  if (!prompt) {
    res.status(400).send("prompt missing.");
    return;
  }

  try {
    // 1) Item typisieren: /v2/boards/{board_id}/items/{item_id}
    // (Damit wir sicher wissen, dass es ein Document Item ist.)
    const item = await miroGetJson(
      `https://api.miro.com/v2/boards/${encodeURIComponent(boardId)}/items/${encodeURIComponent(itemId)}`,
      MIRO_ACCESS_TOKEN
    );

    const itemType = item && item.type ? String(item.type) : "";
    if (itemType !== "document") {
      res.status(400).send(`Selected item is not a document item. REST type=${itemType}`);
      return;
    }

    // 2) Document Details: /v2/boards/{board_id}/documents/{item_id}
    const doc = await miroGetJson(
      `https://api.miro.com/v2/boards/${encodeURIComponent(boardId)}/documents/${encodeURIComponent(itemId)}`,
      MIRO_ACCESS_TOKEN
    );

    // Miro liefert in der Praxis häufig doc.data.documentUrl (z.B. /resources/documents/...).
    // Diese URL liefert NICHT zwingend direkt PDF-Binary; sie kann JSON-Metadaten liefern (z.B. redirect=false),
    // die wiederum einen presigned Download-Link enthalten. Deshalb: robustes Resolve + PDF-Magic-Check.
    const docData = doc && doc.data ? doc.data : {};
    const downloadUrl =
      docData.documentUrl ||
      docData.downloadUrl ||
      docData.download_url ||
      null;

    if (!downloadUrl) {
      res.status(500).json({
        error: "No download URL found in document item response.",
        doc
      });
      return;
    }

    // 3) PDF binary laden (robust: JSON-Metadaten/Redirects/Presigned URLs/PDF-Header prüfen)
    const pdfBytes = await miroDownloadBinary(downloadUrl, MIRO_ACCESS_TOKEN);

    // 4) OpenAI: PDF upload → responses with input_file
    const fileMeta = await openaiUploadPdf(openaiKey, `miro-${itemId}.pdf`, pdfBytes);
    const answer = await openaiAnalyzePdf(openaiKey, model || "gpt-5.2", prompt, fileMeta.id);

    // 5) Optional: Versuch, ein Doc Format auf dem Board anzulegen
    // Doc Formats sind native Miro Docs (nicht PDFs) und haben eigene Endpoints.
    // Falls das Payload nicht passt/Account das blockt: wir geben answer trotzdem zurück.
    let createdDocId = null;
    try {
      const title = `OKR-Analyse – ${new Date().toISOString()}`;
      const content = answer; // Quick & Dirty: plain text

      const created = await miroPostJson(
        `https://api.miro.com/v2/boards/${encodeURIComponent(boardId)}/docs`,
        MIRO_ACCESS_TOKEN,
        {
          data: { title, content },
          position: { x: 0, y: 0, origin: "center" }
        }
      );
      createdDocId = created && created.id ? String(created.id) : null;
    } catch (eDoc) {
      // ignore → fallback im Frontend möglich
    }

    res.status(200).json({
      ok: true,
      boardId,
      itemId,
      openaiFileId: fileMeta.id,
      createdDocId,
      answer
    });
  } catch (e) {
    res.status(500).send(e && e.message ? e.message : String(e));
  }
}

async function readJson(req) {
  if (req.body) {
    if (typeof req.body === "object") return req.body;
    if (typeof req.body === "string") return JSON.parse(req.body);
  }
  const chunks = [];
  for await (const chunk of req) chunks.push(chunk);
  const raw = Buffer.concat(chunks).toString("utf8");
  return raw ? JSON.parse(raw) : {};
}

async function miroGetJson(url, token) {
  const res = await fetch(url, {
    method: "GET",
    headers: { Authorization: `Bearer ${token}` }
  });
  const text = await res.text();
  if (!res.ok) throw new Error(`Miro GET ${url} → ${res.status}: ${text}`);
  return text ? JSON.parse(text) : null;
}

async function miroPostJson(url, token, payload) {
  const res = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  });
  const text = await res.text();
  if (!res.ok) throw new Error(`Miro POST ${url} → ${res.status}: ${text}`);
  return text ? JSON.parse(text) : null;
}

function isPdfMagic(bytes) {
  // "%PDF" = 0x25 0x50 0x44 0x46
  return (
    bytes &&
    bytes.length >= 4 &&
    bytes[0] === 0x25 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x44 &&
    bytes[3] === 0x46
  );
}

function getHeaderLower(res, name) {
  try {
    const v = res.headers.get(name);
    return (v || "").toLowerCase();
  } catch {
    return "";
  }
}

async function readBytes(res) {
  const buf = await res.arrayBuffer();
  return new Uint8Array(buf);
}

function extractUrlCandidatesFromJson(obj) {
  const urls = [];

  function pushIfString(v) {
    if (typeof v === "string" && v.trim()) {
      urls.push(v.trim());
    }
  }

  if (!obj || typeof obj !== "object") {
    return urls;
  }

  // Top-level common fields
  pushIfString(obj.url);
  pushIfString(obj.downloadUrl);
  pushIfString(obj.download_url);
  pushIfString(obj.documentUrl);
  pushIfString(obj.document_url);
  pushIfString(obj.href);
  pushIfString(obj.location);

  // Nested common fields
  if (obj.data && typeof obj.data === "object") {
    pushIfString(obj.data.url);
    pushIfString(obj.data.downloadUrl);
    pushIfString(obj.data.download_url);
    pushIfString(obj.data.documentUrl);
    pushIfString(obj.data.document_url);
    pushIfString(obj.data.href);
  }

  if (obj.links && typeof obj.links === "object") {
    pushIfString(obj.links.download);
    pushIfString(obj.links.file);
    pushIfString(obj.links.self);
  }

  if (obj._links && typeof obj._links === "object") {
    pushIfString(obj._links.download);
    pushIfString(obj._links.file);
    pushIfString(obj._links.self);
  }

  // Generic: search one level deep for any string that looks like an https? URL
  for (const k of Object.keys(obj)) {
    const v = obj[k];
    if (typeof v === "string") {
      if (v.startsWith("http://") || v.startsWith("https://")) {
        urls.push(v);
      }
    } else if (v && typeof v === "object" && !Array.isArray(v)) {
      for (const k2 of Object.keys(v)) {
        const v2 = v[k2];
        if (typeof v2 === "string" && (v2.startsWith("http://") || v2.startsWith("https://"))) {
          urls.push(v2);
        }
      }
    }
  }

  // De-dup
  return Array.from(new Set(urls));
}

async function fetchWithOptionalAuth(url, token, withAuth) {
  const headers = {};
  if (withAuth) {
    headers.Authorization = `Bearer ${token}`;
  }
  return fetch(url, {
    method: "GET",
    headers,
    redirect: "follow"
  });
}

async function miroDownloadBinary(downloadUrl, token) {
  // Robust strategy:
  // 1) Try GET with Authorization → if PDF (by content-type or magic) return bytes.
  // 2) If JSON, parse it and extract a real download URL (presigned) and fetch again.
  // 3) If not JSON and not PDF by content-type, still read bytes and check PDF magic.
  // 4) Retry the second fetch WITHOUT Authorization as presigned URLs often don't need it.
  // 5) Provide actionable error details (content-type + first bytes as hex snippet).

  // ---- First request (with auth) ----
  const res1 = await fetchWithOptionalAuth(downloadUrl, token, true);

  if (!res1.ok) {
    const t = await res1.text().catch(() => "");
    throw new Error(`Download (step1) failed ${res1.status}: ${t}`);
  }

  const ct1 = getHeaderLower(res1, "content-type");

  // If content-type claims PDF, verify magic and return
  if (ct1.includes("application/pdf")) {
    const bytes = await readBytes(res1);
    if (!isPdfMagic(bytes)) {
      const head = Array.from(bytes.slice(0, 16)).map((b) => b.toString(16).padStart(2, "0")).join("");
      throw new Error(`Download (step1) content-type=application/pdf but missing %PDF magic. headHex=${head}`);
    }
    return bytes;
  }

  // If JSON, parse and resolve
  if (ct1.includes("application/json")) {
    const meta = await res1.json().catch(() => null);
    const candidates = extractUrlCandidatesFromJson(meta);

    if (!candidates.length) {
      throw new Error(
        `Download (step1) returned JSON but no download URL candidates found. content-type=${ct1}`
      );
    }

    // Prefer non-api.miro.com URLs if present (often presigned file/CDN)
    const sorted = candidates.slice().sort((a, b) => {
      const aIsMiroApi = a.includes("api.miro.com");
      const bIsMiroApi = b.includes("api.miro.com");
      if (aIsMiroApi === bIsMiroApi) return 0;
      return aIsMiroApi ? 1 : -1;
    });

    let lastErr = null;

    for (const u of sorted) {
      // Try with auth first
      try {
        const res2 = await fetchWithOptionalAuth(u, token, true);
        if (!res2.ok) {
          const t2 = await res2.text().catch(() => "");
          throw new Error(`Download (step2/auth) failed ${res2.status}: ${t2}`);
        }

        const ct2 = getHeaderLower(res2, "content-type");
        const bytes2 = await readBytes(res2);

        if (ct2.includes("application/pdf") || isPdfMagic(bytes2)) {
          if (!isPdfMagic(bytes2)) {
            const head = Array.from(bytes2.slice(0, 16)).map((b) => b.toString(16).padStart(2, "0")).join("");
            throw new Error(`Download (step2/auth) ct=${ct2} but missing %PDF magic. headHex=${head}`);
          }
          return bytes2;
        }

        // Not PDF, keep trying candidates
        const head2 = Array.from(bytes2.slice(0, 16)).map((b) => b.toString(16).padStart(2, "0")).join("");
        lastErr = new Error(`Download (step2/auth) not PDF. ct=${ct2 || "(none)"} headHex=${head2}`);
      } catch (e) {
        lastErr = e;
      }

      // Try without auth (presigned URLs often require no auth; some reject extra headers)
      try {
        const res3 = await fetchWithOptionalAuth(u, token, false);
        if (!res3.ok) {
          const t3 = await res3.text().catch(() => "");
          throw new Error(`Download (step2/noauth) failed ${res3.status}: ${t3}`);
        }

        const ct3 = getHeaderLower(res3, "content-type");
        const bytes3 = await readBytes(res3);

        if (ct3.includes("application/pdf") || isPdfMagic(bytes3)) {
          if (!isPdfMagic(bytes3)) {
            const head = Array.from(bytes3.slice(0, 16)).map((b) => b.toString(16).padStart(2, "0")).join("");
            throw new Error(`Download (step2/noauth) ct=${ct3} but missing %PDF magic. headHex=${head}`);
          }
          return bytes3;
        }

        const head3 = Array.from(bytes3.slice(0, 16)).map((b) => b.toString(16).padStart(2, "0")).join("");
        lastErr = new Error(`Download (step2/noauth) not PDF. ct=${ct3 || "(none)"} headHex=${head3}`);
      } catch (e) {
        lastErr = e;
      }
    }

    throw new Error(
      `Could not resolve PDF from documentUrl JSON. Tried ${sorted.length} candidate(s). Last error: ${lastErr && lastErr.message ? lastErr.message : String(lastErr)}`
    );
  }

  // Not JSON and not declared PDF: read bytes and validate magic
  const bytes1 = await readBytes(res1);
  if (isPdfMagic(bytes1)) {
    return bytes1;
  }

  // As a last resort: try the same URL without auth (some endpoints might behave differently)
  const res4 = await fetchWithOptionalAuth(downloadUrl, token, false);
  if (res4.ok) {
    const ct4 = getHeaderLower(res4, "content-type");
    const bytes4 = await readBytes(res4);
    if (ct4.includes("application/pdf") || isPdfMagic(bytes4)) {
      if (!isPdfMagic(bytes4)) {
        const head = Array.from(bytes4.slice(0, 16)).map((b) => b.toString(16).padStart(2, "0")).join("");
        throw new Error(`Download (step3/noauth) ct=${ct4} but missing %PDF magic. headHex=${head}`);
      }
      return bytes4;
    }
  }

  const head1 = Array.from(bytes1.slice(0, 32)).map((b) => b.toString(16).padStart(2, "0")).join("");
  throw new Error(`Download did not yield a PDF. step1 content-type=${ct1 || "(none)"} headHex=${head1}`);
}

async function openaiUploadPdf(openaiKey, filename, bytes) {
  const form = new FormData();
  form.append("purpose", "user_data");

  // Ensure filename ends with .pdf (helps downstream heuristics)
  const safeName = (typeof filename === "string" && filename.toLowerCase().endsWith(".pdf"))
    ? filename
    : "document.pdf";

  // Double-check magic before upload to fail fast with a clear server-side error
  if (!isPdfMagic(bytes)) {
    const head = Array.from((bytes || []).slice(0, 32)).map((b) => b.toString(16).padStart(2, "0")).join("");
    throw new Error(`Refusing to upload non-PDF bytes to OpenAI. headHex=${head}`);
  }

  form.append("file", new Blob([bytes], { type: "application/pdf" }), safeName);

  const res = await fetch("https://api.openai.com/v1/files", {
    method: "POST",
    headers: { Authorization: `Bearer ${openaiKey}` },
    body: form
  });

  const text = await res.text();
  if (!res.ok) throw new Error(`OpenAI /v1/files → ${res.status}: ${text}`);
  return JSON.parse(text);
}

async function openaiAnalyzePdf(openaiKey, model, prompt, fileId) {
  const body = {
    model,
    input: [
      {
        role: "user",
        content: [
          { type: "input_file", file_id: fileId },
          { type: "input_text", text: prompt }
        ]
      }
    ],
    max_output_tokens: 4000
  };

  const res = await fetch("https://api.openai.com/v1/responses", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${openaiKey}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(body)
  });

  const text = await res.text();
  if (!res.ok) throw new Error(`OpenAI /v1/responses → ${res.status}: ${text}`);

  const data = JSON.parse(text);
  const first = data.output && data.output[0];
  const content = first && first.content;
  const textPart = content && content.find((c) => c.type === "output_text");
  return textPart ? textPart.text : "";
}
