const DEFAULT_OKR_PROMPT = [
  "ROLE",
  "You are a senior strategy-to-execution consultant and OKR architect.",
  "",
  "TASK",
  "Given the attached Corporate Strategy document (usually a PDF with text + charts), produce a corporate OKR catalog (Objectives & Key Results) that is traceable to the document.",
  "",
  "NON-NEGOTIABLE RULES",
  "- Do NOT invent facts, numbers, dates, or commitments that are not supported by the document.",
  "- Every Objective and every Key Result MUST include:",
  "  a) Source page(s)",
  "  b) A short evidence snippet (<= 20 words, paraphrase preferred; if quoting, keep it short)",
  "  c) A label: {EXPLICIT} if directly stated, {INFERRED} if you created a measurable proxy.",
  "- If a critical metric is missing, create a measurable proxy KR but mark it {INFERRED}.",
  "- Do not ask the user questions; proceed with best-effort assumptions and clearly list assumptions.",
  "",
  "OKR DESIGN PRINCIPLES",
  "- 5–9 Corporate-level Objectives max, each with 3–5 Key Results.",
  "- Objectives: qualitative, outcome-oriented, direction-setting (not a metric).",
  "- Key Results: measurable outcomes (numbers or verifiable states), time-bound to the strategy horizon.",
  "- Ensure coverage across these themes IF they exist in the document:",
  "  (1) Growth outcomes,",
  "  (2) Efficiency / capital productivity,",
  "  (3) Portfolio / capital allocation,",
  "  (4) Core-business strengthening,",
  "  (5) New growth creation,",
  "  (6) Capabilities (talent/AI/ops model),",
  "  (7) Shareholder returns and financial guardrails.",
  "- Avoid duplicates: each KR should measure a distinct outcome.",
  "",
  "OUTPUT (IMPORTANT)",
  "Return ONLY the final OKR CATALOG as plain text (no Markdown, no JSON, no extra sections, no analysis, no internal reasoning).",
  "Do NOT output strategy extraction, design principles, quality checks, or machine-readable formats.",
  "",
  "FORMAT (PLAIN TEXT)",
  "OKR CATALOG",
  "Company: <if stated, else 'not stated'>",
  "Strategy name: <if stated, else 'not stated'>",
  "Publication date: <if stated, else 'not stated'>",
  "Time horizon: <if stated, else 'not stated'>",
  "",
  "Objective O1: <objective title>",
  "Intent: <1–2 sentences>",
  "Key Results:",
  "- KR1: <measurable outcome> | Baseline: <if stated else 'n/a'> | Target: <...> | Due: <...> | Evidence: p.<n> <snippet> | Tag: {EXPLICIT|INFERRED}",
  "- KR2: <...> | Baseline: <...> | Target: <...> | Due: <...> | Evidence: p.<n> <snippet> | Tag: {EXPLICIT|INFERRED}",
  "- KR3: <...> | Baseline: <...> | Target: <...> | Due: <...> | Evidence: p.<n> <snippet> | Tag: {EXPLICIT|INFERRED}",
  "",
  "Objective O2: <objective title>",
  "Intent: <1–2 sentences>",
  "Key Results:",
  "- KR1: <...> | Baseline: <...> | Target: <...> | Due: <...> | Evidence: p.<n> <snippet> | Tag: {EXPLICIT|INFERRED}",
  "- KR2: <...> | Baseline: <...> | Target: <...> | Due: <...> | Evidence: p.<n> <snippet> | Tag: {EXPLICIT|INFERRED}",
  "- KR3: <...> | Baseline: <...> | Target: <...> | Due: <...> | Evidence: p.<n> <snippet> | Tag: {EXPLICIT|INFERRED}",
  "",
  "Continue O3..O9 as needed (max 9 objectives)."
].join("\\n");

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
