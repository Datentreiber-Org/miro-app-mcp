const DEFAULT_OKR_PROMPT = [
  "ROLE",
  "You are a senior strategy-to-execution consultant and OKR architect.",
  "",
  "TASK",
  "Given the attached Corporate Strategy document (PDF with text + charts), produce a corporate OKR catalog that is traceable to the document.",
  "",
  "NON-NEGOTIABLE RULES",
  "- Do NOT invent facts, numbers, dates, or commitments that are not supported by the document.",
  "- Every Objective and every Key Result MUST include:",
  "  a) Source page(s)",
  "  b) Evidence snippet (<= 20 words; paraphrase preferred; short quote ok)",
  "  c) A label: {EXPLICIT} if directly stated, {INFERRED} if you created a measurable proxy.",
  "- If a critical metric is missing, create a measurable proxy KR but mark it {INFERRED}.",
  "- Do not ask questions; proceed best-effort and list assumptions briefly if needed.",
  "",
  "OUTPUT FORMAT (IMPORTANT)",
  "Return ONLY the final OKR catalog in MARKDOWN (no JSON, no extra sections, no analysis, no code fences).",
  "",
  "MARKDOWN STRUCTURE",
  "# OKR Catalog",
  "- Company: <if stated, else 'not stated'>",
  "- Strategy name: <if stated, else 'not stated'>",
  "- Publication date: <if stated, else 'not stated'>",
  "- Time horizon: <if stated, else 'not stated'>",
  "",
  "## Objective O1: <title>",
  "**Intent:** <1–2 sentences>",
  "**Key Results:**",
  "- **KR1:** <outcome> | Baseline: <.../n/a> | Target: <...> | Due: <...> | Evidence: p.<n> <snippet> | Tag: {EXPLICIT|INFERRED}",
  "- **KR2:** ...",
  "- **KR3:** ...",
  "",
  "## Objective O2: <title>",
  "**Intent:** <1–2 sentences>",
  "**Key Results:**",
  "- **KR1:** ...",
  "- **KR2:** ...",
  "- **KR3:** ...",
  "",
  "Continue up to O9 max (5–9 objectives total; 3–5 KRs per objective)."
].join("\n");

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

  const MIRO_ACCESS_TOKEN = (process.env.MIRO_ACCESS_TOKEN || "").trim();
  if (!MIRO_ACCESS_TOKEN) {
    res.status(500).send("Server misconfigured: MIRO_ACCESS_TOKEN is missing.");
    return;
  }

  const OPENAI_API_KEY = (process.env.OPENAI_API_KEY || "").trim();

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

  const effectiveOpenaiKey =
    (typeof openaiKey === "string" && openaiKey.trim()) ? openaiKey.trim() : OPENAI_API_KEY;

  if (!effectiveOpenaiKey) {
    res.status(400).send("openaiKey missing (set OPENAI_API_KEY in Vercel env or pass openaiKey).");
    return;
  }

  const effectiveModel =
    (typeof model === "string" && model.trim()) ? model.trim() : "gpt-5.2";

  const effectivePrompt =
    (typeof prompt === "string" && prompt.trim()) ? prompt.trim() : DEFAULT_OKR_PROMPT;

  try {
    // 1) Item typisieren
    const item = await miroGetJson(
      `https://api.miro.com/v2/boards/${encodeURIComponent(boardId)}/items/${encodeURIComponent(itemId)}`,
      MIRO_ACCESS_TOKEN
    );

    const itemType = item && item.type ? String(item.type) : "";
    if (itemType !== "document") {
      res.status(400).send(`Selected item is not a document item. REST type=${itemType}`);
      return;
    }

    // 2) Document Details
    const doc = await miroGetJson(
      `https://api.miro.com/v2/boards/${encodeURIComponent(boardId)}/documents/${encodeURIComponent(itemId)}`,
      MIRO_ACCESS_TOKEN
    );

    const docData = doc && doc.data ? doc.data : {};
    const pdfTitleRaw =
      (docData && typeof docData.title === "string" && docData.title.trim())
        ? docData.title.trim()
        : "Strategy PDF";

    const downloadUrl =
      docData.documentUrl ||
      docData.downloadUrl ||
      docData.download_url ||
      null;

    if (!downloadUrl) {
      res.status(500).json({ error: "No download URL found in document item response.", doc });
      return;
    }

    // Position near the selected PDF (best-effort)
    const srcPos =
      (doc && doc.position && typeof doc.position.x === "number" && typeof doc.position.y === "number")
        ? doc.position
        : (item && item.position && typeof item.position.x === "number" && typeof item.position.y === "number")
          ? item.position
          : { x: 0, y: 0 };

    const srcGeom =
      (doc && doc.geometry && typeof doc.geometry.width === "number" && typeof doc.geometry.height === "number")
        ? doc.geometry
        : (item && item.geometry && typeof item.geometry.width === "number" && typeof item.geometry.height === "number")
          ? item.geometry
          : { width: 800, height: 600 };

    const outX = srcPos.x + (srcGeom.width / 2) + 600;
    const outY = srcPos.y;

    // 3) PDF binary laden
    const pdfBytes = await miroDownloadBinary(downloadUrl, MIRO_ACCESS_TOKEN);

    // 4) OpenAI: PDF upload → responses with input_file
    const fileMeta = await openaiUploadPdf(effectiveOpenaiKey, `miro-${itemId}.pdf`, pdfBytes);

    // IMPORTANT: no artificial max_output_tokens cap here (leave unset)
    let answer = await openaiAnalyzePdf(effectiveOpenaiKey, effectiveModel, effectivePrompt, fileMeta.id);
    answer = normalizeMarkdown(answer);

    // 5) Create Miro Doc Format Item correctly (Doc Formats OpenAPI: data.contentType + data.content)
    // No title field exists; put the title into markdown content.
    const docMarkdown =
      `# OKR Catalog — ${escapeMdInline(pdfTitleRaw)}\n\n` +
      answer;

    let createdDocId = null;
    let createdTextId = null;
    const docCreateErrors = [];

    // Variant A: markdown
    try {
      const created = await miroPostJson(
        `https://api.miro.com/v2/boards/${encodeURIComponent(boardId)}/docs`,
        MIRO_ACCESS_TOKEN,
        {
          data: {
            contentType: "markdown",
            content: docMarkdown
          },
          position: { x: outX, y: outY, origin: "center" }
        }
      );
      createdDocId = created && created.id ? String(created.id) : null;
    } catch (e) {
      docCreateErrors.push(e && e.message ? e.message : String(e));
    }

    // Variant B: html (fallback)
    if (!createdDocId) {
      try {
        const created = await miroPostJson(
          `https://api.miro.com/v2/boards/${encodeURIComponent(boardId)}/docs`,
          MIRO_ACCESS_TOKEN,
          {
            data: {
              contentType: "html",
              content: toSimpleHtml(docMarkdown)
            },
            position: { x: outX, y: outY, origin: "center" }
          }
        );
        createdDocId = created && created.id ? String(created.id) : null;
      } catch (e) {
        docCreateErrors.push(e && e.message ? e.message : String(e));
      }
    }

    // Text fallback only if doc creation failed completely
    if (!createdDocId) {
      try {
        const createdText = await miroPostJson(
          `https://api.miro.com/v2/boards/${encodeURIComponent(boardId)}/texts`,
          MIRO_ACCESS_TOKEN,
          { data: { content: answer }, position: { x: outX, y: outY, origin: "center" } }
        );
        createdTextId = createdText && createdText.id ? String(createdText.id) : null;
      } catch {
        // ignore
      }
    }

    res.status(200).json({
      ok: true,
      boardId,
      itemId,
      openaiFileId: fileMeta.id,
      createdDocId,
      createdTextId,
      docCreateErrors,
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
  const res = await fetch(url, { method: "GET", headers: { Authorization: `Bearer ${token}` } });
  const text = await res.text();
  if (!res.ok) throw new Error(`Miro GET ${url} → ${res.status}: ${text}`);
  return text ? JSON.parse(text) : null;
}

async function miroPostJson(url, token, payload) {
  const res = await fetch(url, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  const text = await res.text();
  if (!res.ok) throw new Error(`Miro POST ${url} → ${res.status}: ${text}`);
  return text ? JSON.parse(text) : null;
}

function normalizeMarkdown(s) {
  if (typeof s !== "string") return "";
  let t = s.trim();

  // remove code fences if model adds them
  if (t.startsWith("```")) {
    t = t.replace(/^```[a-zA-Z0-9_-]*\s*\r?\n?/, "");
    t = t.replace(/\r?\n```$/, "");
    t = t.replace(/```$/, "");
    t = t.trim();
  }
  return t;
}

function escapeMdInline(s) {
  return String(s || "").replaceAll("\n", " ").replaceAll("|", "\\|").trim();
}

function escapeHtml(s) {
  return String(s || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function toSimpleHtml(markdownish) {
  const esc = escapeHtml(markdownish);
  const paragraphs = esc.split(/\n\s*\n/);
  const htmlParas = paragraphs.map((p) => `<p>${p.replaceAll("\n", "<br/>")}</p>`);
  return htmlParas.join("");
}

function isPdfMagic(bytes) {
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
    if (typeof v === "string" && v.trim()) urls.push(v.trim());
  }
  if (!obj || typeof obj !== "object") return urls;

  pushIfString(obj.url);
  pushIfString(obj.downloadUrl);
  pushIfString(obj.download_url);
  pushIfString(obj.documentUrl);
  pushIfString(obj.document_url);
  pushIfString(obj.href);
  pushIfString(obj.location);

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

  for (const k of Object.keys(obj)) {
    const v = obj[k];
    if (typeof v === "string") {
      if (v.startsWith("http://") || v.startsWith("https://")) urls.push(v);
    } else if (v && typeof v === "object" && !Array.isArray(v)) {
      for (const k2 of Object.keys(v)) {
        const v2 = v[k2];
        if (typeof v2 === "string" && (v2.startsWith("http://") || v2.startsWith("https://"))) urls.push(v2);
      }
    }
  }
  return Array.from(new Set(urls));
}

async function fetchWithOptionalAuth(url, token, withAuth) {
  const headers = {};
  if (withAuth) headers.Authorization = `Bearer ${token}`;
  return fetch(url, { method: "GET", headers, redirect: "follow" });
}

async function miroDownloadBinary(downloadUrl, token) {
  const res1 = await fetchWithOptionalAuth(downloadUrl, token, true);
  if (!res1.ok) {
    const t = await res1.text().catch(() => "");
    throw new Error(`Download (step1) failed ${res1.status}: ${t}`);
  }

  const ct1 = getHeaderLower(res1, "content-type");

  if (ct1.includes("application/pdf")) {
    const bytes = await readBytes(res1);
    if (!isPdfMagic(bytes)) {
      const head = Array.from(bytes.slice(0, 16)).map((b) => b.toString(16).padStart(2, "0")).join("");
      throw new Error(`Download (step1) content-type=application/pdf but missing %PDF magic. headHex=${head}`);
    }
    return bytes;
  }

  if (ct1.includes("application/json")) {
    const meta = await res1.json().catch(() => null);
    const candidates = extractUrlCandidatesFromJson(meta);
    if (!candidates.length) {
      throw new Error(`Download (step1) returned JSON but no URL candidates found. content-type=${ct1}`);
    }

    const sorted = candidates.slice().sort((a, b) => {
      const aIsMiroApi = a.includes("api.miro.com");
      const bIsMiroApi = b.includes("api.miro.com");
      if (aIsMiroApi === bIsMiroApi) return 0;
      return aIsMiroApi ? 1 : -1;
    });

    let lastErr = null;

    for (const u of sorted) {
      try {
        const res2 = await fetchWithOptionalAuth(u, token, true);
        if (!res2.ok) {
          const t2 = await res2.text().catch(() => "");
          throw new Error(`Download (step2/auth) failed ${res2.status}: ${t2}`);
        }
        const ct2 = getHeaderLower(res2, "content-type");
        const bytes2 = await readBytes(res2);
        if (ct2.includes("application/pdf") || isPdfMagic(bytes2)) return bytes2;
        lastErr = new Error(`Download (step2/auth) not PDF. ct=${ct2 || "(none)"}`);
      } catch (e) {
        lastErr = e;
      }

      try {
        const res3 = await fetchWithOptionalAuth(u, token, false);
        if (!res3.ok) {
          const t3 = await res3.text().catch(() => "");
          throw new Error(`Download (step2/noauth) failed ${res3.status}: ${t3}`);
        }
        const ct3 = getHeaderLower(res3, "content-type");
        const bytes3 = await readBytes(res3);
        if (ct3.includes("application/pdf") || isPdfMagic(bytes3)) return bytes3;
        lastErr = new Error(`Download (step2/noauth) not PDF. ct=${ct3 || "(none)"}`);
      } catch (e) {
        lastErr = e;
      }
    }

    throw new Error(`Could not resolve PDF from JSON. Last error: ${lastErr && lastErr.message ? lastErr.message : String(lastErr)}`);
  }

  const bytes1 = await readBytes(res1);
  if (isPdfMagic(bytes1)) return bytes1;

  throw new Error(`Download did not yield a PDF. step1 content-type=${ct1 || "(none)"}`);
}

async function openaiUploadPdf(openaiKey, filename, bytes) {
  const form = new FormData();
  form.append("purpose", "user_data");

  const safeName =
    (typeof filename === "string" && filename.toLowerCase().endsWith(".pdf"))
      ? filename
      : "document.pdf";

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
    ]
    // IMPORTANT: do NOT set max_output_tokens here (no artificial cap). :contentReference[oaicite:2]{index=2}
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
