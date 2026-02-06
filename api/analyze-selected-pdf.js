const DEFAULT_OKR_PROMPT = `ROLE
You are a senior strategy-to-execution consultant and OKR architect.

TASK
Read the attached corporate strategy PDF (text + charts) and produce a corporate OKR catalog.

STRICT RULES
- Do NOT invent facts, numbers, dates, or commitments that are not supported by the document.
- If a key metric is missing, create a measurable proxy KR but mark it as {INFERRED}.
- Keep it concise and usable as a final OKR catalog.

OUTPUT (IMPORTANT)
Return ONLY the final OKR CATALOG as plain text (no Markdown, no JSON, no extra sections, no analysis, no “Step 1/2/3/4”).
Do NOT include any machine-readable formats.

FORMAT
OKR CATALOG
Time horizon: <if stated, else "not stated">

Objective O1: <title>
Intent: <1–2 sentences>
Key Results:
- KR1: <measurable outcome> | Target: <...> | Due: <...> | Evidence: p.<n> <short snippet> | Tag: {EXPLICIT|INFERRED}
- KR2: ...
- KR3: ...

Objective O2: ...
...

CONSTRAINTS
- 5–9 objectives max
- 3–5 key results per objective
- Evidence snippet <= 20 words per KR (paraphrase preferred; if quoting keep short)`;

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
    // 1) Item typisieren: /v2/boards/{board_id}/items/{item_id}
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

    const docData = doc && doc.data ? doc.data : {};
    const pdfTitleRaw = (docData && typeof docData.title === "string" && docData.title.trim())
      ? docData.title.trim()
      : "Strategy PDF";

    const d
