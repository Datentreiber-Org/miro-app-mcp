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

// The OKR output must be written into the existing doc format item on the board.
// The doc must be found by this exact title.
const TARGET_DOC_TITLE = "Objectives and Key Results Catalog";

const PROMPT_DOC_TITLE = "OKR Extraction Prompt";

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

  let effectivePrompt =
    (typeof prompt === "string" && prompt.trim()) ? prompt.trim() : "";

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

    // If no prompt was provided via API, try to load it from an on-board Doc ("OKR Extraction Prompt").
    // GitHub hardcoded DEFAULT_OKR_PROMPT remains the fallback.
    if (!effectivePrompt) {
      try {
        const boardPrompt = await loadPromptFromBoardDoc(boardId, MIRO_ACCESS_TOKEN, PROMPT_DOC_TITLE);
        if (boardPrompt) effectivePrompt = boardPrompt;
      } catch {
        // ignore -> fallback below
      }
    }
    if (!effectivePrompt) effectivePrompt = DEFAULT_OKR_PROMPT;

    // 4) OpenAI: PDF upload → responses with input_file
    const fileMeta = await openaiUploadPdf(effectiveOpenaiKey, `miro-${itemId}.pdf`, pdfBytes);

    // IMPORTANT: no artificial max_output_tokens cap here (leave unset)
    let answer = await openaiAnalyzePdf(effectiveOpenaiKey, effectiveModel, effectivePrompt, fileMeta.id);
    answer = normalizeMarkdown(answer);

    // 5) Write result into an existing, pre-positioned Miro Doc Format item.
    // As there is no REST "update doc format content" endpoint, we recreate the doc at the same
    // position (and best-effort geometry/parent), then delete the old placeholder.
    const docMarkdown =
      `# OKR Catalog — ${escapeMdInline(pdfTitleRaw)}\n\n` +
      answer;

    const targetDoc = await findDocFormatByTitle(boardId, MIRO_ACCESS_TOKEN, TARGET_DOC_TITLE);
    if (!targetDoc || !targetDoc.id) {
      res.status(404).send(`Target doc format item not found. Expected an existing doc titled "${TARGET_DOC_TITLE}".`);
      return;
    }

    const targetPos =
      (targetDoc.position && typeof targetDoc.position.x === "number" && typeof targetDoc.position.y === "number")
        ? targetDoc.position
        : { x: outX, y: outY, origin: "center" };

    const targetGeom =
      (targetDoc.geometry && typeof targetDoc.geometry === "object")
        ? targetDoc.geometry
        : null;

    const targetParentId =
      (targetDoc.parent && typeof targetDoc.parent.id === "string" && targetDoc.parent.id)
        ? targetDoc.parent.id
        : (targetDoc.parent && typeof targetDoc.parent.id === "number")
          ? String(targetDoc.parent.id)
          : null;

    const SCALE_FACTOR = 3;

    const targetGeomWidthBase =
      (targetGeom && typeof targetGeom.width === "number" && Number.isFinite(targetGeom.width))
        ? targetGeom.width
        : null;

    const targetGeomHeightBase =
      (targetGeom && typeof targetGeom.height === "number" && Number.isFinite(targetGeom.height))
        ? targetGeom.height
        : null;

    const targetGeomWidthScaled =
      (targetGeomWidthBase !== null) ? (targetGeomWidthBase * SCALE_FACTOR) : null;

    const targetGeomHeightScaled =
      (targetGeomHeightBase !== null) ? (targetGeomHeightBase * SCALE_FACTOR) : null;

    const createPos = {
      x: targetPos.x,
      y: targetPos.y,
      origin: (targetPos && typeof targetPos.origin === "string" && targetPos.origin) ? targetPos.origin : "center"
    };

    // Keep the placeholder's top-left corner fixed while scaling.
    // For origin=center: shifting center by (Δwidth/2, Δheight/2) keeps top-left constant.
    const createPosScaled = {
      x: createPos.x,
      y: createPos.y,
      origin: createPos.origin
    };

    const originNorm = String(createPos.origin || "center").toLowerCase().replaceAll("_", "-");
    if (originNorm === "center") {
      if (targetGeomWidthBase !== null && targetGeomWidthScaled !== null) {
        createPosScaled.x = createPos.x + ((targetGeomWidthScaled - targetGeomWidthBase) / 2);
      }
      if (targetGeomHeightBase !== null && targetGeomHeightScaled !== null) {
        createPosScaled.y = createPos.y + ((targetGeomHeightScaled - targetGeomHeightBase) / 2);
      }
    }


    const docMarkdownWithTitleHeading = `# ${TARGET_DOC_TITLE}\n\n${docMarkdown}`;

    let createdDocId = null;
    let createdTextId = null;
    const docCreateErrors = [];
    let replacedDocId = null;

    const createVariants = [
      // Prefer setting the doc title explicitly (if supported by the endpoint).
      { contentType: "markdown", content: docMarkdown, includeTitle: true, includeGeometry: true },
      { contentType: "markdown", content: docMarkdown, includeTitle: true, includeGeometry: false },

      // Fallback: if title is not supported, force the title via the first heading.
      { contentType: "markdown", content: docMarkdownWithTitleHeading, includeTitle: false, includeGeometry: true },
      { contentType: "markdown", content: docMarkdownWithTitleHeading, includeTitle: false, includeGeometry: false },

      // HTML fallbacks.
      { contentType: "html", content: toSimpleHtml(docMarkdown), includeTitle: true, includeGeometry: true },
      { contentType: "html", content: toSimpleHtml(docMarkdown), includeTitle: true, includeGeometry: false },
      { contentType: "html", content: toSimpleHtml(docMarkdownWithTitleHeading), includeTitle: false, includeGeometry: true },
      { contentType: "html", content: toSimpleHtml(docMarkdownWithTitleHeading), includeTitle: false, includeGeometry: false }
    ];

    for (const v of createVariants) {
      if (createdDocId) break;

      const payload = {
        data: {
          contentType: v.contentType,
          content: v.content
        },
        position: (v.includeGeometry ? createPosScaled : createPos)
      };


      if (v.includeTitle) {
        payload.data.title = TARGET_DOC_TITLE;
      }

      if (v.includeGeometry && targetGeom) {
        const w = targetGeomWidthScaled;
        const h = targetGeomHeightScaled;
        if (w !== null || h !== null) {
          payload.geometry = {};
          // IMPORTANT: set only ONE dimension (width OR height), not both.
          // This prevents the geometry-variant from failing and falling back to a non-geometry variant.
          if (w !== null) {
            payload.geometry.width = w;
          } else if (h !== null) {
            payload.geometry.height = h;
          }
        }
      }


      try {
        const created = await miroPostJson(
          `https://api.miro.com/v2/boards/${encodeURIComponent(boardId)}/docs`,
          MIRO_ACCESS_TOKEN,
          payload
        );
        createdDocId = created && created.id ? String(created.id) : null;
      } catch (e) {
        docCreateErrors.push(e && e.message ? e.message : String(e));
      }
    }

    // Best-effort: attach the new doc to the same parent (e.g., frame) as the placeholder.
    if (createdDocId && targetParentId) {
      try {
        await miroPatchJson(
          `https://api.miro.com/v2/boards/${encodeURIComponent(boardId)}/items/${encodeURIComponent(createdDocId)}`,
          MIRO_ACCESS_TOKEN,
          { parent: { id: targetParentId } }
        );
      } catch (e) {
        docCreateErrors.push(e && e.message ? e.message : String(e));
      }
    }

    // Delete the placeholder only after the new doc has been created.
    if (createdDocId) {
      replacedDocId = String(targetDoc.id);
      try {
        await miroDelete(
          `https://api.miro.com/v2/boards/${encodeURIComponent(boardId)}/docs/${encodeURIComponent(replacedDocId)}`,
          MIRO_ACCESS_TOKEN
        );
      } catch (e) {
        // Fallback: some boards/items behave more consistently via the generic delete-item endpoint.
        try {
          await miroDelete(
            `https://api.miro.com/v2/boards/${encodeURIComponent(boardId)}/items/${encodeURIComponent(replacedDocId)}`,
            MIRO_ACCESS_TOKEN
          );
        } catch (e2) {
          docCreateErrors.push(e && e.message ? e.message : String(e));
          docCreateErrors.push(e2 && e2.message ? e2.message : String(e2));
        }
      }
    }


    // Text fallback only if doc creation failed completely (we keep the placeholder in that case).
    if (!createdDocId) {
      try {
        const createdText = await miroPostJson(
          `https://api.miro.com/v2/boards/${encodeURIComponent(boardId)}/texts`,
          MIRO_ACCESS_TOKEN,
          { data: { content: answer }, position: createPos }
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
      targetDocTitle: TARGET_DOC_TITLE,
      replacedDocId,
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

async function miroPatchJson(url, token, payload) {
  const res = await fetch(url, {
    method: "PATCH",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  const text = await res.text();
  if (!res.ok) throw new Error(`Miro PATCH ${url} → ${res.status}: ${text}`);
  return text ? JSON.parse(text) : null;
}

async function miroDelete(url, token) {
  const res = await fetch(url, {
    method: "DELETE",
    headers: { Authorization: `Bearer ${token}` }
  });
  const text = await res.text().catch(() => "");
  if (!res.ok) throw new Error(`Miro DELETE ${url} → ${res.status}: ${text}`);
  return true;
}

function extractItemTitle(item) {
  if (!item || typeof item !== "object") return "";
  const candidates = [
    item.title,
    item.name,
    item.data && item.data.title,
    item.data && item.data.name
  ];
  for (const c of candidates) {
    if (typeof c === "string" && c.trim()) return c.trim();
  }
  return "";
}

function extractNextCursor(listResponse) {
  if (!listResponse || typeof listResponse !== "object") return null;

  // Common patterns in Miro cursor-based pagination responses.
  if (typeof listResponse.cursor === "string" && listResponse.cursor) return listResponse.cursor;
  if (listResponse.cursor && typeof listResponse.cursor === "object") {
    if (typeof listResponse.cursor.next === "string" && listResponse.cursor.next) return listResponse.cursor.next;
    if (typeof listResponse.cursor.cursor === "string" && listResponse.cursor.cursor) return listResponse.cursor.cursor;
  }
  if (typeof listResponse.nextCursor === "string" && listResponse.nextCursor) return listResponse.nextCursor;
  if (typeof listResponse.next_page_token === "string" && listResponse.next_page_token) return listResponse.next_page_token;

  return null;
}

async function findDocFormatByTitle(boardId, token, exactTitle) {
  const wanted = String(exactTitle || "").trim();
  if (!wanted) return null;

  function stripHtmlTags(s) {
    return String(s || "").replace(/<[^>]*>/g, "");
  }

  function extractDocContent(docObj) {
    if (!docObj || typeof docObj !== "object") return "";
    const data = (docObj.data && typeof docObj.data === "object") ? docObj.data : {};
    if (typeof data.content === "string" && data.content.trim()) return data.content;
    if (typeof docObj.content === "string" && docObj.content.trim()) return docObj.content;
    return "";
  }

  function isEffectivelyEmptyDoc(detailsObj) {
    const content = extractDocContent(detailsObj);
    if (!content) return false;

    const stripped = stripHtmlTags(content).replaceAll("\u00a0", " ").trim();
    if (!stripped) return true;

    // Exactly the title and nothing else.
    if (stripped === wanted) return true;

    // Single-line markdown title only: "# <wanted>"
    const lines = String(content || "")
      .split(/\r?\n/)
      .map((l) => stripHtmlTags(l).replaceAll("\u00a0", " ").trim())
      .filter(Boolean);

    if (lines.length === 1) {
      const first = lines[0].replace(/^#{1,6}\s+/, "").trim();
      if (first === wanted) return true;
    }

    return false;
  }

  function mergeListItemAndDetails(listItem, details) {
    // Keep position/geometry/parent from the items list response,
    // but keep content from /docs/{id}.
    const merged = {};

    if (details && typeof details === "object") {
      for (const k of Object.keys(details)) merged[k] = details[k];
    }
    if (listItem && typeof listItem === "object") {
      for (const k of Object.keys(listItem)) merged[k] = listItem[k];
    }

    // Ensure the id is the board item id from /items listing.
    if (listItem && listItem.id) merged.id = listItem.id;

    // Deep-merge data (prefer details.data for content).
    const data = {};
    if (listItem && listItem.data && typeof listItem.data === "object") Object.assign(data, listItem.data);
    if (details && details.data && typeof details.data === "object") Object.assign(data, details.data);
    if (Object.keys(data).length) merged.data = data;

    return merged;
  }

  function extractDocFormatTitleFromContent(docObj) {
    if (!docObj || typeof docObj !== "object") return "";

    // If Miro ever adds a real title field here, prefer it.
    const t = extractItemTitle(docObj);
    if (t) return t;

    const content = extractDocContent(docObj);
    if (!content) return "";

    const lines = content.split(/\r?\n/);
    for (const line of lines) {
      const l = String(line || "").trim();
      if (!l) continue;

      // Markdown title pattern: "# Title"
      const m = l.match(/^#{1,6}\s+(.*)$/);
      if (m && m[1] && m[1].trim()) return m[1].trim();

      // HTML title pattern: <h1>Title</h1>
      const h1 = l.match(/<h1[^>]*>(.*?)<\/h1>/i);
      if (h1 && h1[1]) {
        const txt = stripHtmlTags(h1[1]).trim();
        if (txt) return txt;
      }

      // Fallback: first non-empty line, stripped
      return stripHtmlTags(l).trim();
    }

    return "";
  }

  function geometryArea(listItem) {
    const g = (listItem && listItem.geometry && typeof listItem.geometry === "object") ? listItem.geometry : {};
    const w = (typeof g.width === "number" && Number.isFinite(g.width)) ? g.width : 0;
    const h = (typeof g.height === "number" && Number.isFinite(g.height)) ? g.height : 0;
    return w * h;
  }

  let cursor = null;
  let best = null;
  let bestArea = -1;

  // Hard cap to avoid accidental infinite loops in case pagination behaves unexpectedly.
  for (let i = 0; i < 100; i++) {
    const url = new URL(`https://api.miro.com/v2/boards/${encodeURIComponent(boardId)}/items`);
    url.searchParams.set("type", "doc_format");
    url.searchParams.set("limit", "50");
    if (cursor) url.searchParams.set("cursor", cursor);

    const page = await miroGetJson(url.toString(), token);
    const items = page && Array.isArray(page.data) ? page.data : [];

    for (const it of items) {
      if (!it || typeof it !== "object") continue;
      if (String(it.type || "") !== "doc_format") continue;
      if (!it.id) continue;

      const listTitle = extractItemTitle(it);

      let details = null;
      let isMatch = false;

      if (listTitle === wanted) {
        isMatch = true;
        try {
          details = await miroGetJson(
            `https://api.miro.com/v2/boards/${encodeURIComponent(boardId)}/docs/${encodeURIComponent(it.id)}`,
            token
          );
        } catch {
          details = null;
        }
      } else {
        try {
          details = await miroGetJson(
            `https://api.miro.com/v2/boards/${encodeURIComponent(boardId)}/docs/${encodeURIComponent(it.id)}`,
            token
          );
          const derived = extractDocFormatTitleFromContent(details);
          if (derived === wanted) isMatch = true;
        } catch {
          details = null;
        }
      }

      if (!isMatch) continue;

      // Prefer the pre-positioned placeholder doc (effectively empty).
      if (details && isEffectivelyEmptyDoc(details)) {
        return mergeListItemAndDetails(it, details);
      }

      // Otherwise, prefer the visually "intended" one (largest geometry),
      // so duplicates (typically smaller, created by the buggy run) don't get picked.
      const area = geometryArea(it);
      const merged = mergeListItemAndDetails(it, details);

      if (!best || area > bestArea) {
        best = merged;
        bestArea = area;
      }
    }

    const nextCursor = extractNextCursor(page);
    if (!nextCursor) break;
    cursor = nextCursor;
  }

  return best;
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
    // IMPORTANT: do NOT set max_output_tokens here (no artificial cap).
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

function extractDocFormatContent(docObj) {
  if (!docObj || typeof docObj !== "object") return "";
  const data = (docObj.data && typeof docObj.data === "object") ? docObj.data : {};
  if (typeof data.content === "string" && data.content.trim()) return data.content;
  if (typeof docObj.content === "string" && docObj.content.trim()) return docObj.content;
  return "";
}

function looksLikeHtmlDocContent(s) {
  const t = String(s || "");
  // Heuristic: only treat as HTML when it contains common HTML tags that Miro docs often return.
  return /<\/?(p|br|div|span|h[1-6]|strong|em|ul|ol|li|a)\b/i.test(t);
}

function decodeHtmlEntities(s) {
  return String(s || "")
    .replaceAll("&nbsp;", " ")
    .replaceAll("&#160;", " ")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&amp;", "&")
    .replaceAll("&quot;", "\"")
    .replaceAll("&#39;", "'")
    .replaceAll("&#x27;", "'");
}

function htmlToTextPreserveAngles(html) {
  let t = String(html || "");

  t = t.replace(/<\s*br\s*\/?>/gi, "\n");
  t = t.replace(/<\/\s*p\s*>/gi, "\n");
  t = t.replace(/<\/\s*div\s*>/gi, "\n");
  t = t.replace(/<\/\s*h[1-6]\s*>/gi, "\n");
  t = t.replace(/<\/\s*li\s*>/gi, "\n");

  // Remove only common formatting tags; do NOT remove unknown <...> sequences
  // because prompts may include placeholders like "<if stated, else ...>".
  t = t.replace(/<\/?\s*(p|div|span|h[1-6]|strong|em|ul|ol|li|a)\b[^>]*>/gi, "");

  // Decode entities AFTER stripping tags so &lt;...&gt; becomes <...> (your prompt placeholders)
  t = decodeHtmlEntities(t);

  t = t.replace(/\r\n/g, "\n");
  t = t.replace(/\n{3,}/g, "\n\n");
  return t.trim();
}

function removeLeadingTitleIfPresent(text, title) {
  const wanted = String(title || "").trim();
  if (!wanted) return String(text || "").trim();

  const lines = String(text || "").replace(/\r\n/g, "\n").split("\n");

  // drop leading empty lines
  let i = 0;
  while (i < lines.length && !String(lines[i]).trim()) i++;

  if (i >= lines.length) return "";

  const first = String(lines[i]).trim();
  const firstNoHashes = first.replace(/^#{1,6}\s+/, "").trim();

  if (first === wanted || firstNoHashes === wanted) {
    i++;
    // also drop blank line(s) after title
    while (i < lines.length && !String(lines[i]).trim()) i++;
    return lines.slice(i).join("\n").trim();
  }

  return lines.join("\n").trim();
}

async function loadPromptFromBoardDoc(boardId, token, promptDocTitle) {
  const doc = await findDocFormatByTitle(boardId, token, promptDocTitle);
  if (!doc || !doc.id) return "";

  const contentRaw = extractDocFormatContent(doc);
  if (!contentRaw) return "";

  let text = contentRaw;

  if (looksLikeHtmlDocContent(contentRaw)) {
    text = htmlToTextPreserveAngles(contentRaw);
  } else {
    // Treat as markdown/plain text; DO NOT strip <...> because prompt contains placeholders.
    text = String(contentRaw || "").replaceAll("\u00a0", " ").trim();
  }

  text = removeLeadingTitleIfPresent(text, promptDocTitle);

  // “Empty” guard
  if (!text) return "";

  // If user accidentally only wrote the title, treat as empty.
  if (text === promptDocTitle || text === `# ${promptDocTitle}`) return "";

  return text;
}
