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
  const body = await readJson(req).catch((e) => null);
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

    // Doku ist im UI dynamisch; in der Praxis enthält doc.data typischerweise eine Download-URL,
    // die mit Access Token abgerufen werden muss.
    const docData = doc && doc.data ? doc.data : {};
    const downloadUrl =
      docData.downloadUrl ||
      docData.download_url ||
      docData.url ||
      docData.downloadLink ||
      docData.download_link ||
      null;

    if (!downloadUrl) {
      res.status(500).json({
        error: "No download URL found in document item response.",
        doc
      });
      return;
    }

    // 3) PDF binary laden (Bearer Token, Redirects folgen)
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

async function miroDownloadBinary(downloadUrl, token) {
  const res = await fetch(downloadUrl, {
    method: "GET",
    headers: { Authorization: `Bearer ${token}` },
    redirect: "follow"
  });
  if (!res.ok) {
    const t = await res.text().catch(() => "");
    throw new Error(`Download failed ${res.status}: ${t}`);
  }
  const buf = await res.arrayBuffer();
  return new Uint8Array(buf);
}

async function openaiUploadPdf(openaiKey, filename, bytes) {
  const form = new FormData();
  form.append("purpose", "user_data");
  form.append("file", new Blob([bytes], { type: "application/pdf" }), filename);

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
