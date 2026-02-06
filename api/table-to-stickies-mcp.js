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

  // MCP token optional: if not set, we try MIRO_ACCESS_TOKEN.
  // If MCP server rejects it, set MIRO_MCP_ACCESS_TOKEN in Vercel env.
  const MIRO_MCP_ACCESS_TOKEN = (process.env.MIRO_MCP_ACCESS_TOKEN || "").trim();
  const MCP_TOKEN = MIRO_MCP_ACCESS_TOKEN || MIRO_ACCESS_TOKEN;

  const body = await readJson(req).catch(() => null);
  if (!body) {
    res.status(400).send("Invalid JSON body.");
    return;
  }

  const { boardId, tableItemId } = body;
  if (!boardId || !tableItemId) {
    res.status(400).send("boardId or tableItemId missing.");
    return;
  }

  try {
    // 1) Get table item geometry/position via REST (for placement)
    const tableItem = await miroGetJson(
      `https://api.miro.com/v2/boards/${encodeURIComponent(boardId)}/items/${encodeURIComponent(tableItemId)}`,
      MIRO_ACCESS_TOKEN
    );

    const pos = (tableItem && tableItem.position) ? tableItem.position : {};
    const geom = (tableItem && tableItem.geometry) ? tableItem.geometry : {};

    const tableX = num(pos.x, 0);
    const tableY = num(pos.y, 0);
    const tableW = num(geom.width, 900);
    const tableH = num(geom.height, 500);

    // 2) Read table rows via MCP tool table_list_rows
    // Tools are documented; table_list_rows exists. :contentReference[oaicite:3]{index=3}
    const mcp = await mcpStartSession(MCP_TOKEN);

    const toolsList = await mcpRequestJson(mcp, {
      jsonrpc: "2.0",
      id: 2,
      method: "tools/list",
      params: {}
    });

    const tools = toolsList && toolsList.result && Array.isArray(toolsList.result.tools) ? toolsList.result.tools : [];
    const tableTool = tools.find((t) => t && t.name === "table_list_rows");

    // If tool schema not available, still try a reasonable default.
    const inputSchema = tableTool && tableTool.inputSchema ? tableTool.inputSchema : null;

    const args = buildTableListArgs(inputSchema, boardId, tableItemId);

    const call = await mcpRequestJson(mcp, {
      jsonrpc: "2.0",
      id: 3,
      method: "tools/call",
      params: {
        name: "table_list_rows",
        arguments: args
      }
    });

    const tableData = parseTableListRowsCall(call);
    const columns = tableData.columns;
    const rows = tableData.rows;

    if (!columns.length || !rows.length) {
      res.status(500).json({
        error: "MCP table_list_rows returned no rows/columns.",
        debug: { columnsCount: columns.length, rowsCount: rows.length, args, call }
      });
      return;
    }

    // We include header row as stickies + all data rows.
    const headerRow = columns.map((c) => c.title);
    const allRows = [headerRow, ...rows];

    // 3) Create stickies (each cell -> sticky) via REST
    // Sticky endpoint exists. :contentReference[oaicite:4]{index=4}
    const stickyW = 280;
    const stickyH = 170;
    const gapX = 40;
    const gapY = 30;

    const stepX = stickyW + gapX;
    const stepY = stickyH + gapY;

    const startX = tableX + (tableW / 2) + 200 + (stickyW / 2);

    const totalRows = allRows.length;
    const gridHeight = (totalRows > 1) ? ((totalRows - 1) * stepY) : 0;
    const startY = tableY - (gridHeight / 2);

    const createdIds = [];

    for (let r = 0; r < allRows.length; r++) {
      const row = allRows[r];
      for (let c = 0; c < columns.length; c++) {
        const text = (row && typeof row[c] === "string") ? row[c] : String(row[c] ?? "").trim();
        const x = startX + c * stepX;
        const y = startY + r * stepY;

        const stickyId = await createStickyNote(boardId, MIRO_ACCESS_TOKEN, text, x, y);
        if (stickyId) {
          createdIds.push(stickyId);
        }
      }
    }

    res.status(200).json({
      ok: true,
      boardId,
      tableItemId,
      columns: columns.map((c) => c.title),
      rowCount: rows.length,
      createdCount: createdIds.length,
      createdIds
    });
  } catch (e) {
    res.status(500).send(e && e.message ? e.message : String(e));
  }
}

function num(v, fallback) {
  return (typeof v === "number" && Number.isFinite(v)) ? v : fallback;
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

async function createStickyNote(boardId, token, content, x, y) {
  const text = String(content || "").trim() || " ";
  const payloadA = {
    data: { content: text, shape: "square" },
    position: { x, y, origin: "center" }
  };

  try {
    const created = await miroPostJson(
      `https://api.miro.com/v2/boards/${encodeURIComponent(boardId)}/sticky_notes`,
      token,
      payloadA
    );
    return created && created.id ? String(created.id) : null;
  } catch (e1) {
    // Fallback payload shape (some environments accept root-level fields)
    const payloadB = {
      content: text,
      shape: "square",
      position: { x, y, origin: "center" }
    };
    try {
      const created = await miroPostJson(
        `https://api.miro.com/v2/boards/${encodeURIComponent(boardId)}/sticky_notes`,
        token,
        payloadB
      );
      return created && created.id ? String(created.id) : null;
    } catch (e2) {
      return null;
    }
  }
}

// --------------------
// MCP CLIENT (Streamable HTTP JSON-RPC)
// --------------------
// MCP endpoint URL per Miro docs. :contentReference[oaicite:5]{index=5}
// Transport spec: POST JSON-RPC to endpoint; server may return JSON or SSE. :contentReference[oaicite:6]{index=6}

async function mcpStartSession(token) {
  const endpoint = "https://mcp.miro.com/";
  const initReq = {
    jsonrpc: "2.0",
    id: 1,
    method: "initialize",
    params: {
      protocolVersion: "2025-06-18",
      capabilities: {},
      clientInfo: { name: "datentreiber-miro-app", version: "1.0" }
    }
  };

  const { json, sessionId } = await mcpPost(endpoint, token, null, initReq);
  return { endpoint, token, sessionId };
}

async function mcpRequestJson(mcp, reqObj) {
  const { json } = await mcpPost(mcp.endpoint, mcp.token, mcp.sessionId, reqObj);
  return json;
}

async function mcpPost(endpoint, token, sessionId, reqObj) {
  const headers = {
    "Content-Type": "application/json",
    "Accept": "application/json, text/event-stream",
    "Authorization": `Bearer ${token}`,
    // helps some servers that validate Origin
    "Origin": "https://miro-app-mcp.vercel.app"
  };
  if (sessionId) {
    headers["Mcp-Session-Id"] = sessionId;
  }

  const res = await fetch(endpoint, {
    method: "POST",
    headers,
    body: JSON.stringify(reqObj),
    redirect: "follow"
  });

  const newSessionId = res.headers.get("Mcp-Session-Id") || sessionId || null;

  if (!res.ok) {
    const t = await res.text().catch(() => "");
    throw new Error(`MCP POST ${endpoint} → ${res.status}: ${t}`);
  }

  const ct = (res.headers.get("content-type") || "").toLowerCase();

  if (ct.includes("application/json")) {
    const json = await res.json();
    return { json, sessionId: newSessionId };
  }

  if (ct.includes("text/event-stream")) {
    const text = await res.text();
    const json = parseSseForJsonRpc(text, reqObj.id);
    if (!json) {
      throw new Error("MCP SSE response did not contain a JSON-RPC message.");
    }
    return { json, sessionId: newSessionId };
  }

  // fallback: try JSON parse
  const raw = await res.text();
  try {
    return { json: JSON.parse(raw), sessionId: newSessionId };
  } catch {
    throw new Error(`MCP unexpected content-type=${ct}, body=${raw.slice(0, 200)}`);
  }
}

function parseSseForJsonRpc(sseText, desiredId) {
  const lines = String(sseText || "").split("\n");
  let lastJson = null;

  for (const line of lines) {
    const l = line.trim();
    if (!l.startsWith("data:")) continue;
    const payload = l.slice(5).trim();
    if (!payload) continue;

    try {
      const obj = JSON.parse(payload);
      lastJson = obj;
      if (typeof desiredId !== "undefined" && obj && obj.id === desiredId) {
        return obj;
      }
    } catch {
      // ignore
    }
  }
  return lastJson;
}

// --------------------
// MCP table_list_rows helpers
// --------------------

function buildTableListArgs(inputSchema, boardId, tableId) {
  const args = {};

  const props = (inputSchema && inputSchema.properties && typeof inputSchema.properties === "object")
    ? Object.keys(inputSchema.properties)
    : [];

  const has = (k) => props.includes(k);

  // Board
  if (has("board_id")) args.board_id = boardId;
  else if (has("boardId")) args.boardId = boardId;
  else if (has("board")) args.board = boardId;

  // Table
  if (has("table_id")) args.table_id = tableId;
  else if (has("tableId")) args.tableId = tableId;
  else if (has("table")) args.table = tableId;
  else if (has("id")) args.id = tableId;

  // Pagination size
  if (has("limit")) args.limit = 100;
  else if (has("pageSize")) args.pageSize = 100;
  else if (has("page_size")) args.page_size = 100;

  // If schema is unknown, use common defaults
  if (props.length === 0) {
    args.board_id = boardId;
    args.table_id = tableId;
    args.limit = 100;
  }

  return args;
}

function parseTableListRowsCall(callResp) {
  // Expected: JSON-RPC response with result = CallToolResult
  if (!callResp || typeof callResp !== "object") {
    return { columns: [], rows: [] };
  }
  if (callResp.error) {
    const msg = callResp.error.message || JSON.stringify(callResp.error);
    throw new Error(`MCP tools/call error: ${msg}`);
  }

  const result = callResp.result || {};
  const structured = result.structuredContent || null;

  // Preferred: structuredContent
  if (structured && typeof structured === "object") {
    const columns = normalizeColumns(structured.columns || structured.columnMetadata || structured.cols || []);
    const rows = normalizeRows(structured.rows || structured.data || structured.items || [], columns);
    return { columns, rows };
  }

  // Fallback: sometimes content[0].text contains JSON
  const content = Array.isArray(result.content) ? result.content : [];
  const firstText = content.find((c) => c && c.type === "text" && typeof c.text === "string");
  if (firstText && firstText.text) {
    try {
      const obj = JSON.parse(firstText.text);
      const columns = normalizeColumns(obj.columns || obj.columnMetadata || obj.cols || []);
      const rows = normalizeRows(obj.rows || obj.data || obj.items || [], columns);
      return { columns, rows };
    } catch {
      // ignore
    }
  }

  return { columns: [], rows: [] };
}

function normalizeColumns(cols) {
  const columns = [];
  if (Array.isArray(cols)) {
    for (const c of cols) {
      if (!c) continue;
      if (typeof c === "string") {
        columns.push({ id: c, title: c });
      } else if (typeof c === "object") {
        const id = (c.id || c.columnId || c.key || c.name || c.title || "").toString();
        const title = (c.title || c.name || c.label || id || "Column").toString();
        columns.push({ id: id || title, title });
      }
    }
  }
  // Ensure at least 1
  return columns;
}

function normalizeRows(rowsRaw, columns) {
  const rows = [];
  const colIds = columns.map((c) => c.id);

  if (!Array.isArray(rowsRaw)) return rows;

  for (const r of rowsRaw) {
    const row = [];

    // Case A: array row
    if (Array.isArray(r)) {
      for (let i = 0; i < columns.length; i++) {
        row.push(cellToText(r[i]));
      }
      rows.push(row);
      continue;
    }

    // Case B: row.cells array
    if (r && typeof r === "object" && Array.isArray(r.cells)) {
      for (let i = 0; i < columns.length; i++) {
        row.push(cellToText(r.cells[i]));
      }
      rows.push(row);
      continue;
    }

    // Case C: row.values object keyed by column id
    if (r && typeof r === "object" && r.values && typeof r.values === "object") {
      for (const id of colIds) {
        row.push(cellToText(r.values[id]));
      }
      rows.push(row);
      continue;
    }

    // Case D: row object keyed directly by column id/title
    if (r && typeof r === "object") {
      for (const id of colIds) {
        row.push(cellToText(r[id]));
      }
      rows.push(row);
      continue;
    }
  }

  return rows;
}

function cellToText(v) {
  if (v === null || typeof v === "undefined") return "";
  if (typeof v === "string") return v.trim();
  if (typeof v === "number" || typeof v === "boolean") return String(v);

  if (typeof v === "object") {
    if (typeof v.text === "string") return v.text.trim();
    if (typeof v.value === "string") return v.value.trim();
    if (typeof v.label === "string") return v.label.trim();
    if (typeof v.name === "string") return v.name.trim();
    if (typeof v.title === "string") return v.title.trim();
    if (typeof v.displayValue === "string") return v.displayValue.trim();
    try { return JSON.stringify(v); } catch { return ""; }
  }
  return String(v);
}
