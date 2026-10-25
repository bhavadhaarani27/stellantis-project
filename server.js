const http = require("node:http");
const fs = require("node:fs");
const path = require("node:path");
const crypto = require("node:crypto");
const zlib = require("node:zlib");
const store = require("./lib/viewer");
const { listSections } = require("./lib/pptx-render");
const { buildTeamList, CATALOGUE, teamKey } = require("./lib/teams");
const { URL } = require("node:url");

const root = __dirname;
const dataRoot = path.join(root, "data");
const uploadRoot = path.join(dataRoot, "uploads");
const reviewRoot = path.join(dataRoot, "reviews");
const simpleInputRoot = path.join(dataRoot, "simple-input");
const port = Number(process.env.PORT || 8080);
const host = process.env.HOST || "0.0.0.0";
for (const folder of [dataRoot, uploadRoot, reviewRoot, simpleInputRoot]) fs.mkdirSync(folder, { recursive: true });

const contentTypes = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".svg": "image/svg+xml"
};

function sendJson(response, status, body) {
  response.writeHead(status, { "Content-Type": "application/json; charset=utf-8" });
  response.end(JSON.stringify(body));
}

function safeName(value) {
  return String(value || "file").replace(/[^a-zA-Z0-9._-]+/g, "_");
}

// ---- Synced SharePoint/OneDrive folder (a normal local folder) ----
const syncSettingsPath = path.join(dataRoot, "sync-folder.json");

function readSyncFolder() {
  try { return JSON.parse(fs.readFileSync(syncSettingsPath, "utf8")).folder || ""; } catch { return ""; }
}

// Lists .pptx files up to 5 levels deep, newest first.
function listSyncFiles(folder) {
  const found = [];
  const walk = (dir, depth) => {
    let entries = [];
    try { entries = fs.readdirSync(dir, { withFileTypes: true }); } catch { return; }
    for (const entry of entries) {
      const full = path.join(dir, entry.name);
      if (entry.isDirectory() && depth < 5) walk(full, depth + 1);
      else if (entry.isFile() && /\.pptx$/i.test(entry.name) && !entry.name.startsWith("~$")) {
        found.push({ path: path.relative(folder, full), modified: fs.statSync(full).mtimeMs });
      }
    }
  };
  walk(folder, 0);
  return found.sort((a, b) => b.modified - a.modified).slice(0, 300);
}

// Only files inside the configured synced folder may be read.
function resolveSyncFile(relative) {
  const folder = readSyncFolder();
  if (!folder) throw new Error("Set the synced folder first.");
  const full = path.resolve(folder, relative);
  if (!full.toLowerCase().startsWith(path.resolve(folder).toLowerCase() + path.sep) || !/\.pptx$/i.test(full) || !fs.existsSync(full)) {
    throw new Error("That file is not available in the synced folder.");
  }
  return full;
}

function readBody(request) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    request.on("data", (chunk) => {
      size += chunk.length;
      if (size > 250 * 1024 * 1024) {
        reject(new Error("Upload exceeds the 250 MB limit."));
        request.destroy();
        return;
      }
      chunks.push(chunk);
    });
    request.on("end", () => resolve(Buffer.concat(chunks)));
    request.on("error", reject);
  });
}

function parseMultipart(body, contentType) {
  const match = /boundary=(?:"([^"]+)"|([^;]+))/i.exec(contentType || "");
  if (!match) throw new Error("Multipart boundary is missing.");
  const boundary = Buffer.from(`--${match[1] || match[2]}`);
  const fields = {};
  let file = null;
  let offset = 0;

  while (offset < body.length) {
    const start = body.indexOf(boundary, offset);
    if (start < 0) break;
    const partStart = start + boundary.length;
    if (body.subarray(partStart, partStart + 2).toString() === "--") break;
    const headerStart = partStart + 2;
    const headerEnd = body.indexOf(Buffer.from("\r\n\r\n"), headerStart);
    if (headerEnd < 0) break;
    const headers = body.subarray(headerStart, headerEnd).toString("utf8");
    const nextBoundary = body.indexOf(boundary, headerEnd + 4);
    if (nextBoundary < 0) break;
    const value = body.subarray(headerEnd + 4, nextBoundary - 2);
    const disposition = /name="([^"]+)"(?:;\s*filename="([^"]*)")?/i.exec(headers);
    if (!disposition) {
      offset = nextBoundary;
      continue;
    }
    const fieldName = disposition[1];
    if (disposition[2] !== undefined) {
      file = { fieldName, name: disposition[2], data: value };
    } else {
      fields[fieldName] = value.toString("utf8");
    }
    offset = nextBoundary;
  }
  return { fields, file };
}

function decodeXml(value) {
  return String(value || "")
    .replace(/&#x([0-9a-f]+);/gi, (_, code) => String.fromCodePoint(parseInt(code, 16)))
    .replace(/&#([0-9]+);/g, (_, code) => String.fromCodePoint(Number(code)))
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function readZipEntries(buffer) {
  const endSignature = Buffer.from([0x50, 0x4b, 0x05, 0x06]);
  const centralSignature = 0x02014b50;
  const localSignature = 0x04034b50;
  const end = buffer.lastIndexOf(endSignature);
  if (end < 0) throw new Error("The uploaded file is not a readable .pptx archive.");
  const centralSize = buffer.readUInt32LE(end + 12);
  const centralOffset = buffer.readUInt32LE(end + 16);
  const entries = new Map();
  let cursor = centralOffset;
  const limit = centralOffset + centralSize;
  while (cursor < limit && buffer.readUInt32LE(cursor) === centralSignature) {
    const compression = buffer.readUInt16LE(cursor + 10);
    const compressedSize = buffer.readUInt32LE(cursor + 20);
    const nameLength = buffer.readUInt16LE(cursor + 28);
    const extraLength = buffer.readUInt16LE(cursor + 30);
    const commentLength = buffer.readUInt16LE(cursor + 32);
    const localOffset = buffer.readUInt32LE(cursor + 42);
    const name = buffer.subarray(cursor + 46, cursor + 46 + nameLength).toString("utf8");
    entries.set(name, { compression, compressedSize, localOffset });
    cursor += 46 + nameLength + extraLength + commentLength;
  }
  function read(name) {
    const entry = entries.get(name);
    if (!entry) return null;
    const local = entry.localOffset;
    if (buffer.readUInt32LE(local) !== localSignature) throw new Error(`Invalid archive entry: ${name}`);
    const nameLength = buffer.readUInt16LE(local + 26);
    const extraLength = buffer.readUInt16LE(local + 28);
    const start = local + 30 + nameLength + extraLength;
    const compressed = buffer.subarray(start, start + entry.compressedSize);
    if (entry.compression === 0) return compressed;
    if (entry.compression === 8) return zlib.inflateRawSync(compressed);
    throw new Error(`Unsupported archive compression in ${name}.`);
  }
  return { entries, read };
}

function xmlText(xml) {
  return [...String(xml || "").matchAll(/<a:t\b[^>]*>([\s\S]*?)<\/a:t>/g)]
    .map((match) => decodeXml(match[1]).replace(/\s+/g, " ").trim())
    .filter(Boolean);
}

function relationshipTargets(xml) {
  const targets = new Map();
  for (const match of String(xml || "").matchAll(/<Relationship\b([^>]+?)\/?>/g)) {
    const attributes = {};
    for (const attribute of match[1].matchAll(/(\w+):?(\w*)="([^"]*)"/g)) {
      attributes[attribute[1] + attribute[2]] = decodeXml(attribute[3]);
    }
    if (attributes.Id && attributes.Target) targets.set(attributes.Id, attributes.Target);
  }
  return targets;
}

function mimeTypeForImage(name) {
  const extension = path.extname(name).toLowerCase();
  return { ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".gif": "image/gif", ".svg": "image/svg+xml" }[extension] || "application/octet-stream";
}

function classifySlide(text, index) {
  const value = text.toLowerCase();
  if (index === 1) return ["home", "Home"];
  if (index === 2 || value.includes("agenda")) return ["agenda", "Agenda"];
  if (value.includes("fqm")) return ["fqm", "FQM"];
  if (value.includes("afme")) return ["afme", "AFME"];
  if (value.includes("top chart") || value.includes("top quality")) return ["top-chart", "Top Chart"];
  if (value.includes("infield")) return ["infield", "Infield Quality"];
  if (value.includes("eeq")) return ["eeq", "EEQ Dashboard"];
  return ["additional", "Review Slides"];
}

function htmlEscape(value) {
  return String(value || "").replace(/[&<>"']/g, (character) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;"
  }[character]));
}

function reviewFile(month, year) {
  return path.join(reviewRoot, `${safeName(month)}-${safeName(year)}.json`);
}

function readReview(month, year) {
  const target = reviewFile(month, year);
  if (!fs.existsSync(target)) return null;
  return JSON.parse(fs.readFileSync(target, "utf8"));
}

function writeReview(review) {
  fs.writeFileSync(reviewFile(review.month, review.year), JSON.stringify(review, null, 2), "utf8");
}

// Team status for the "late teams" panel: every catalogue team plus any extra team already in the review.
function sessionSummary(state, folder, period) {
  const entries = state.order.map((key) => ({ key, ...state.teams[key] }));
  const extras = CATALOGUE.filter((c) => !state.teams[c.key]).map((c) => ({ key: c.key, label: c.label, source: "missing" }));
  return {
    month: state.month, year: state.year, session: state.session, meetingDate: state.meetingDate, agenda: state.agenda || "",
    url: `/monthly-output/simple/${period}/${folder}/index.html`,
    teams: [...entries, ...extras].map((t) => ({ key: t.key, label: t.label, source: t.source, link: t.link || "", fileName: t.fileName || "", slides: t.slides ? t.slides.length : 0, updatedAt: t.updatedAt || "" }))
  };
}
async function handleApi(request, response, url) {
  if (request.method === "GET" && url.pathname === "/api/health") {
    return sendJson(response, 200, { ok: true, service: "EHEQ Monthly Review", time: new Date().toISOString() });
  }
  if (request.method === "GET" && url.pathname === "/api/review") {
    const review = readReview(url.searchParams.get("month"), url.searchParams.get("year"));
    return sendJson(response, 200, { review });
  }
  if (request.method === "POST" && url.pathname === "/api/reviews") {
    const body = JSON.parse((await readBody(request)).toString("utf8"));
    if (!body.month || !body.year) return sendJson(response, 400, { error: "Month and year are required." });
    writeReview(body);
    return sendJson(response, 201, { ok: true, review: body });
  }
  if (request.method === "POST" && url.pathname === "/api/uploads") {
    const body = await readBody(request);
    const parsed = parseMultipart(body, request.headers["content-type"]);
    if (!parsed.file || !/\.(ppt|pptx)$/i.test(parsed.file.name)) {
      return sendJson(response, 400, { error: "Only .ppt and .pptx files are accepted." });
    }
    const id = `${Date.now()}-${crypto.randomBytes(4).toString("hex")}`;
    const fileName = `${id}-${safeName(parsed.file.name)}`;
    const destination = path.join(uploadRoot, fileName);
    fs.writeFileSync(destination, parsed.file.data);
    const record = {
      id,
      fileName: parsed.file.name,
      storedFile: path.relative(root, destination),
      month: parsed.fields.month || "",
      year: parsed.fields.year || "",
      session: parsed.fields.session || "",
      team: parsed.fields.team || "",
      status: "Uploaded",
      uploadedAt: new Date().toISOString(),
      converter: "Not configured"
    };
    fs.writeFileSync(path.join(uploadRoot, `${id}.json`), JSON.stringify(record, null, 2), "utf8");
    return sendJson(response, 201, { ok: true, record });
  }
  if (request.method === "DELETE" && url.pathname.startsWith("/api/uploads/")) {
    const id = decodeURIComponent(url.pathname.slice("/api/uploads/".length));
    if (!/^[0-9a-f-]+$/i.test(id)) return sendJson(response, 400, { error: "Invalid upload identifier." });
    const metadataPath = path.join(uploadRoot, `${id}.json`);
    if (!fs.existsSync(metadataPath)) return sendJson(response, 404, { error: "Uploaded presentation was not found." });
    const record = JSON.parse(fs.readFileSync(metadataPath, "utf8"));
    const storedPath = record.storedFile ? path.resolve(root, record.storedFile) : null;
    if (storedPath && storedPath.startsWith(uploadRoot) && fs.existsSync(storedPath)) fs.unlinkSync(storedPath);
    fs.unlinkSync(metadataPath);
    return sendJson(response, 200, { ok: true, id });
  }
  if (request.method === "GET" && url.pathname === "/api/simple-list") {
    const base = path.join(root, "monthly-output", "simple");
    const items = [];
    for (const period of fs.existsSync(base) ? fs.readdirSync(base) : []) {
      for (const folder of ["Session1", "Session2"]) {
        const metaPath = path.join(base, period, folder, "meta.json");
        if (!fs.existsSync(metaPath)) continue;
        const meta = JSON.parse(fs.readFileSync(metaPath, "utf8"));
        items.push({ period, session: meta.session, month: meta.month, year: meta.year, slideCount: meta.slideCount, generatedAt: meta.generatedAt, url: `/monthly-output/simple/${period}/${folder}/index.html` });
      }
    }
    items.sort((a, b) => String(b.generatedAt).localeCompare(String(a.generatedAt)));
    return sendJson(response, 200, { items });
  }
  if (request.method === "GET" && url.pathname === "/api/sync-files") {
    const folder = readSyncFolder();
    if (!folder) return sendJson(response, 200, { folder: "", files: [] });
    if (!fs.existsSync(folder)) return sendJson(response, 200, { folder, files: [], error: "Folder not found on this PC." });
    return sendJson(response, 200, { folder, files: listSyncFiles(folder) });
  }
  if (request.method === "POST" && url.pathname === "/api/sync-folder") {
    const body = JSON.parse((await readBody(request)).toString("utf8") || "{}");
    const folder = String(body.folder || "").trim().replace(/^"|"$/g, "");
    if (folder && !fs.existsSync(folder)) return sendJson(response, 400, { error: "That folder does not exist on this PC." });
    fs.mkdirSync(dataRoot, { recursive: true });
    fs.writeFileSync(syncSettingsPath, JSON.stringify({ folder }));
    return sendJson(response, 200, { folder, files: folder ? listSyncFiles(folder) : [] });
  }
  if (request.method === "GET" && url.pathname === "/api/simple-session") {
    const folder = url.searchParams.get("session") === "Session 2" ? "Session2" : "Session1";
    const dir = path.join(root, "monthly-output", "simple", safeName(url.searchParams.get("period") || ""), folder);
    const state = store.loadState(dir);
    if (!state) return sendJson(response, 404, { error: "Review not found." });
    return sendJson(response, 200, sessionSummary(state, folder, safeName(url.searchParams.get("period"))));
  }
  if (request.method === "POST" && url.pathname === "/api/simple-update") {
    const parsed = parseMultipart(await readBody(request), request.headers["content-type"]);
    const f = parsed.fields;
    const folder = f.session === "Session 2" ? "Session2" : "Session1";
    const period = `${safeName(f.month || "")}-${safeName(f.year || "")}`;
    const dir = path.join(root, "monthly-output", "simple", period, folder);
    const key = teamKey(f.key || "");
    try {
      if (f.action === "setPptx") {
        let data, name;
        if (f.syncPath) {
          const file = resolveSyncFile(f.syncPath);
          data = fs.readFileSync(file);
          name = path.basename(file);
        } else {
          if (!parsed.file || !/\.pptx$/i.test(parsed.file.name)) throw new Error("Please choose a .pptx file.");
          ({ data, name } = { data: parsed.file.data, name: parsed.file.name });
        }
        store.setTeamPptx(dir, key, data, name);
      } else if (f.action === "setLink") store.setTeamLink(dir, key, (f.link || "").trim());
      else if (f.action === "setPending") store.setTeamPending(dir, key);
      else if (f.action === "remove") store.removeTeam(dir, key);
      else if (f.action === "order") store.setOrder(dir, JSON.parse(f.order || "[]"));
      else if (f.action === "date") store.setMeetingDate(dir, f.date || "");
      else if (f.action === "agenda") store.setAgenda(dir, f.agenda || "");
      else throw new Error("Unknown action.");
    } catch (error) {
      return sendJson(response, 400, { error: error.message });
    }
    return sendJson(response, 200, sessionSummary(store.loadState(dir), folder, period));
  }  if (request.method === "POST" && url.pathname === "/api/simple-sections") {
    const parsed = parseMultipart(await readBody(request), request.headers["content-type"]);
    if (!parsed.file || !/\.pptx$/i.test(parsed.file.name)) return sendJson(response, 400, { error: "Please choose a .pptx file." });
    try { return sendJson(response, 200, { teams: buildTeamList(listSections(parsed.file.data)) }); }
    catch (error) { return sendJson(response, 400, { error: error.message }); }
  }
  if (request.method === "POST" && url.pathname === "/api/simple-generate") {
    const body = await readBody(request);
    const parsed = parseMultipart(body, request.headers["content-type"]);
    const month = parsed.fields.month;
    const year = parsed.fields.year;
    const session = parsed.fields.session;
    if (!month || !year || !/^Session [12]$/.test(session || "")) {
      return sendJson(response, 400, { error: "Month, year, and session are required." });
    }
    if (!parsed.file || !/\.(ppt|pptx)$/i.test(parsed.file.name)) {
      return sendJson(response, 400, { error: "Only .ppt and .pptx files are accepted." });
    }
    const period = `${safeName(month)}-${safeName(year)}`;
    const sessionFolder = session === "Session 2" ? "Session2" : "Session1";
    const inputFolder = path.join(simpleInputRoot, period, sessionFolder);
    const outputRoot = path.join(root, "monthly-output", "simple", period, sessionFolder);
    fs.mkdirSync(inputFolder, { recursive: true });
    if (fs.existsSync(outputRoot)) fs.rmSync(outputRoot, { recursive: true, force: true });
    fs.mkdirSync(path.join(outputRoot, "slides"), { recursive: true });
    const inputPath = path.join(inputFolder, `${Date.now()}-${safeName(parsed.file.name)}`);
    fs.writeFileSync(inputPath, parsed.file.data);
    if (!/\.pptx$/i.test(parsed.file.name)) {
      return sendJson(response, 400, { error: "This self-contained generator accepts .pptx files. Save older .ppt files as .pptx and upload again." });
    }
    try {
      store.createSession(parsed.file.data, parsed.file.name, month, year, session, outputRoot, { teams: JSON.parse(parsed.fields.teams || "[]"), meetingDate: parsed.fields.meetingDate || "", agenda: parsed.fields.agenda || "" });
    } catch (error) {
      return sendJson(response, 500, { error: error.message || "The PowerPoint file could not be read." });
    }
    const generatedIndex = path.join(outputRoot, "index.html");
    if (!fs.existsSync(generatedIndex)) return sendJson(response, 500, { error: "The presentation was processed but no webpage was generated." });
    const relativeUrl = `/${path.relative(root, generatedIndex).replace(/\\/g, "/")}`;
    return sendJson(response, 201, { ok: true, session, url: relativeUrl });
  }
  return sendJson(response, 404, { error: "API route not found." });
}

function serveStatic(request, response, url) {
  let requested = decodeURIComponent(url.pathname === "/" ? "/simple.html" : url.pathname);
  if (requested === "/workspace") requested = "/workspace.html";
  if (requested === "/simple") requested = "/simple.html";
  const target = path.resolve(root, `.${requested}`);
  if (!target.startsWith(root) || !fs.existsSync(target) || fs.statSync(target).isDirectory()) {
    response.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    response.end("Not found");
    return;
  }
  response.writeHead(200, { "Content-Type": contentTypes[path.extname(target).toLowerCase()] || "application/octet-stream" });
  fs.createReadStream(target).pipe(response);
}

const server = http.createServer(async (request, response) => {
  const url = new URL(request.url, `http://${request.headers.host || "localhost"}`);
  try {
    if (url.pathname.startsWith("/api/")) await handleApi(request, response, url);
    else serveStatic(request, response, url);
  } catch (error) {
    sendJson(response, 500, { error: error.message });
  }
});

server.listen(port, host, () => {
  console.log(`EHEQ Monthly Review running at http://localhost:${port}`);
  console.log(`Shared network URL: http://<this-computer-name>:${port}`);
});


