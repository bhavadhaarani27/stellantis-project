"use strict";
/*
 * Self-contained PPTX reader and slide renderer.
 * Converts a .pptx file into positioned HTML/SVG slides without PowerPoint,
 * Office, LibreOffice or any external service.
 */
const fs = require("node:fs");
const path = require("node:path");
const zlib = require("node:zlib");

/* ---------------------------------------------------------------- ZIP --- */
function openZip(buffer) {
  const end = buffer.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
  if (end < 0) throw new Error("The uploaded file is not a valid .pptx (PowerPoint) file.");
  const total = buffer.readUInt16LE(end + 10);
  let cursor = buffer.readUInt32LE(end + 16);
  const entries = new Map();
  for (let i = 0; i < total && buffer.readUInt32LE(cursor) === 0x02014b50; i++) {
    const nameLength = buffer.readUInt16LE(cursor + 28);
    const extraLength = buffer.readUInt16LE(cursor + 30);
    const commentLength = buffer.readUInt16LE(cursor + 32);
    const name = buffer.subarray(cursor + 46, cursor + 46 + nameLength).toString("utf8");
    entries.set(name, {
      method: buffer.readUInt16LE(cursor + 10),
      size: buffer.readUInt32LE(cursor + 20),
      offset: buffer.readUInt32LE(cursor + 42)
    });
    cursor += 46 + nameLength + extraLength + commentLength;
  }
  function read(name) {
    const entry = entries.get(name);
    if (!entry) return null;
    const nameLength = buffer.readUInt16LE(entry.offset + 26);
    const extraLength = buffer.readUInt16LE(entry.offset + 28);
    const start = entry.offset + 30 + nameLength + extraLength;
    const raw = buffer.subarray(start, start + entry.size);
    if (entry.method === 0) return raw;
    if (entry.method === 8) return zlib.inflateRawSync(raw);
    throw new Error(`Unsupported compression in ${name}`);
  }
  return { entries, read };
}

/* ---------------------------------------------------------------- XML --- */
const ENTITIES = { lt: "<", gt: ">", quot: '"', apos: "'", amp: "&" };
function decodeXml(text) {
  return text.replace(/&(#x[0-9a-f]+|#\d+|lt|gt|quot|apos|amp);/gi, (_, code) => {
    if (code[0] === "#") return String.fromCodePoint(code[1].toLowerCase() === "x" ? parseInt(code.slice(2), 16) : Number(code.slice(1)));
    return ENTITIES[code.toLowerCase()];
  });
}

// Minimal XML parser: returns {n: tagName, a: attributes, c: children}. Text is kept only for a:t and c:v.
function parseXml(xml) {
  const root = { n: "#root", a: {}, c: [] };
  const stack = [root];
  const pattern = /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<(\/?)([^\s>\/]+)([^>]*?)(\/?)>|([^<]+)/g;
  let match;
  while ((match = pattern.exec(xml))) {
    const top = stack[stack.length - 1];
    if (match[5] !== undefined) {
      if (top.n === "a:t" || top.n === "c:v") top.c.push({ t: decodeXml(match[5]) });
    } else if (match[2]) {
      if (match[1]) { if (stack.length > 1) stack.pop(); continue; }
      const node = { n: match[2], a: {}, c: [] };
      for (const attr of match[3].matchAll(/([^\s=]+)="([^"]*)"/g)) node.a[attr[1]] = decodeXml(attr[2]);
      top.c.push(node);
      if (!match[4]) stack.push(node);
    }
  }
  return root.c[0];
}
const kids = (node, name) => (node && node.c ? node.c.filter((child) => child.n === name) : []);
const kid = (node, name) => kids(node, name)[0];
const textOf = (node) => (node && node.c ? node.c.map((child) => child.t || textOf(child)).join("") : "");
function findAll(node, name, out = []) {
  if (!node || !node.c) return out;
  for (const child of node.c) {
    if (child.n === name) out.push(child);
    findAll(child, name, out);
  }
  return out;
}
function find(node, name) {
  if (!node || !node.c) return null;
  for (const child of node.c) {
    if (child.n === name) return child;
    const deeper = find(child, name);
    if (deeper) return deeper;
  }
  return null;
}
const num = (value, fallback = 0) => (value === undefined || value === "" || value === null || Number.isNaN(Number(value)) ? fallback : Number(value));
const esc = (value) => String(value ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const round = (value, digits = 3) => Number(value.toFixed(digits));

/* ------------------------------------------------------------- COLOURS --- */
const PRESET_COLORS = { black: "000000", white: "ffffff", red: "ff0000", green: "008000", blue: "0000ff", yellow: "ffff00", gray: "808080", grey: "808080", orange: "ffa500" };

function rgbToHsl([r, g, b]) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = 0, s = 0;
  const l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
    h /= 6;
  }
  return [h, s, l];
}
function hslToRgb([h, s, l]) {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const f = (t) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return [f(h + 1 / 3) * 255, f(h) * 255, f(h - 1 / 3) * 255];
}

// Resolves a colour element (srgbClr / schemeClr / sysClr ...) inside "holder" to a CSS colour.
function colorFrom(holder, ctx) {
  if (!holder || !holder.c) return null;
  const node = holder.c.find((child) => /^a:(srgbClr|schemeClr|sysClr|prstClr)$/.test(child.n));
  if (!node) return null;
  let hex;
  if (node.n === "a:srgbClr") hex = node.a.val;
  else if (node.n === "a:sysClr") hex = node.a.lastClr || "000000";
  else if (node.n === "a:prstClr") hex = PRESET_COLORS[node.a.val] || "000000";
  else {
    let name = node.a.val;
    if (name === "phClr") name = ctx.phClr || "tx1";
    name = ctx.clrMap[name] || name;
    hex = ctx.theme[name] || "000000";
  }
  let rgb = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
  let alpha = 1;
  let hsl = null;
  const toHsl = () => { if (!hsl) hsl = rgbToHsl(rgb); };
  const fromHsl = () => { if (hsl) { rgb = hslToRgb(hsl); hsl = null; } };
  for (const mod of node.c) {
    const value = num(mod.a && mod.a.val) / 100000;
    if (mod.n === "a:lumMod") { toHsl(); hsl[2] = Math.min(1, hsl[2] * value); }
    else if (mod.n === "a:lumOff") { toHsl(); hsl[2] = Math.min(1, Math.max(0, hsl[2] + value)); }
    else if (mod.n === "a:satMod") { toHsl(); hsl[1] = Math.min(1, hsl[1] * value); }
    else if (mod.n === "a:tint") { fromHsl(); rgb = rgb.map((v) => v * value + 255 * (1 - value)); }
    else if (mod.n === "a:shade") { fromHsl(); rgb = rgb.map((v) => v * value); }
    else if (mod.n === "a:alpha") alpha = value;
  }
  fromHsl();
  const [r, g, b] = rgb.map((v) => Math.max(0, Math.min(255, Math.round(v))));
  return alpha < 1 ? `rgba(${r},${g},${b},${round(alpha, 2)})` : `rgb(${r},${g},${b})`;
}
const schemeHolder = (name, mods = []) => ({ c: [{ n: "a:schemeClr", a: { val: name }, c: mods }] });

/* --------------------------------------------------------------- DECK ---- */
class Deck {
  constructor(buffer, mediaDir) {
    this.zip = openZip(buffer);
    this.mediaDir = mediaDir;
    this.xmlCache = new Map();
    this.mediaMap = new Map();
    this.mediaCount = 0;
    fs.mkdirSync(mediaDir, { recursive: true });
  }
  xml(part) {
    if (!this.xmlCache.has(part)) {
      const data = this.zip.read(part);
      this.xmlCache.set(part, data ? parseXml(data.toString("utf8")) : null);
    }
    return this.xmlCache.get(part);
  }
  rels(part) {
    const dir = path.posix.dirname(part);
    const root = this.xml(`${dir}/_rels/${path.posix.basename(part)}.rels`);
    const map = new Map();
    for (const rel of kids(root, "Relationship")) {
      const target = rel.a.Target || "";
      const resolved = target.startsWith("/") ? target.slice(1) : path.posix.normalize(path.posix.join(dir, target));
      map.set(rel.a.Id, { type: rel.a.Type || "", target: resolved, external: rel.a.TargetMode === "External" });
    }
    return map;
  }
  // Copies an embedded image to the output folder once and returns its relative URL.
  media(partInfo, rid) {
    const rel = partInfo.rels.get(rid);
    if (!rel || rel.external) return null;
    const ext = path.extname(rel.target).toLowerCase();
    if (![".png", ".jpg", ".jpeg", ".gif", ".svg", ".bmp", ".webp"].includes(ext)) return null;
    if (!this.mediaMap.has(rel.target)) {
      const data = this.zip.read(rel.target);
      if (!data) return null;
      const name = `m${++this.mediaCount}${ext}`;
      fs.writeFileSync(path.join(this.mediaDir, name), data);
      this.mediaMap.set(rel.target, `media/${name}`);
    }
    return this.mediaMap.get(rel.target);
  }
}

function loadTheme(deck, masterPath) {
  const masterRels = deck.rels(masterPath);
  const themeRel = [...masterRels.values()].find((rel) => rel.type.endsWith("/theme"));
  const theme = {};
  const root = themeRel ? deck.xml(themeRel.target) : null;
  const scheme = find(root, "a:clrScheme");
  for (const entry of scheme ? scheme.c : []) {
    const color = entry.c[0];
    if (color) theme[entry.n.replace("a:", "")] = color.a.val === undefined || color.n === "a:sysClr" ? color.a.lastClr : color.a.val;
  }
  const defaults = { dk1: "000000", lt1: "ffffff", dk2: "1f2937", lt2: "e5e7eb", accent1: "4472c4", accent2: "ed7d31", accent3: "a5a5a5", accent4: "ffc000", accent5: "5b9bd5", accent6: "70ad47", hlink: "0563c1", folHlink: "954f72" };
  for (const key of Object.keys(defaults)) if (!theme[key]) theme[key] = defaults[key];
  return theme;
}

/* ------------------------------------------------------ TEXT / STYLES ---- */
function xfrmOf(node, wrapperName = "a:xfrm") {
  const xf = kid(node, wrapperName);
  if (!xf) return null;
  const off = kid(xf, "a:off"), ext = kid(xf, "a:ext");
  if (!off || !ext) return null;
  return { x: num(off.a.x), y: num(off.a.y), cx: num(ext.a.cx), cy: num(ext.a.cy), rot: num(xf.a.rot), flipH: xf.a.flipH === "1", flipV: xf.a.flipV === "1" };
}
function placeholderOf(node) {
  return kid(kid(kid(node, "p:nvSpPr") || kid(node, "p:nvPicPr") || kid(node, "p:nvGraphicFramePr"), "p:nvPr"), "p:ph");
}
function findPlaceholder(spTree, ph) {
  if (!spTree || !ph) return null;
  for (const shape of spTree.c) {
    if (shape.n !== "p:sp") continue;
    const other = placeholderOf(shape);
    if (other && ph.a.idx !== undefined && other.a.idx === ph.a.idx) return shape;
  }
  const alias = (t) => ({ ctrTitle: "title", subTitle: "body", obj: "body" }[t] || t || "body");
  for (const shape of spTree.c) {
    if (shape.n !== "p:sp") continue;
    const other = placeholderOf(shape);
    if (other && alias(other.a.type) === alias(ph.a.type)) return shape;
  }
  return null;
}

// Merges paragraph properties from lowest to highest priority.
function mergeParagraph(pPrList) {
  const out = { attrs: {}, bu: null, lnSpc: null, spcBef: null, defRPr: { attrs: {}, fill: null, latin: null } };
  for (const pPr of pPrList) {
    if (!pPr) continue;
    Object.assign(out.attrs, pPr.a);
    for (const child of pPr.c) {
      if (child.n === "a:lnSpc" && kid(child, "a:spcPct")) out.lnSpc = num(kid(child, "a:spcPct").a.val) / 100000;
      else if (child.n === "a:spcBef") {
        if (kid(child, "a:spcPts")) out.spcBef = { pts: num(kid(child, "a:spcPts").a.val) / 100 };
        else if (kid(child, "a:spcPct")) out.spcBef = { pct: num(kid(child, "a:spcPct").a.val) / 100000 };
      } else if (child.n === "a:buNone") out.bu = { none: true };
      else if (child.n === "a:buChar") out.bu = { char: child.a.char };
      else if (child.n === "a:buAutoNum") out.bu = { auto: true };
      else if (child.n === "a:defRPr") mergeRun(out.defRPr, child);
    }
  }
  return out;
}
function mergeRun(target, rPr) {
  if (!rPr) return target;
  Object.assign(target.attrs, rPr.a);
  const fill = kid(rPr, "a:solidFill");
  if (fill) target.fill = fill;
  const latin = kid(rPr, "a:latin");
  if (latin && latin.a.typeface && !latin.a.typeface.startsWith("+")) target.latin = latin.a.typeface;
  return target;
}

function cqw(emu, ctx) { return `${round((emu / ctx.slideW) * 100, 4)}cqw`; }

function renderText(txBody, ctx, shapeCtx) {
  if (!txBody) return { html: "", hasText: false };
  const bodyPr = kid(txBody, "a:bodyPr") || { a: {}, c: [] };
  const fontScale = kid(bodyPr, "a:normAutofit") ? num(kid(bodyPr, "a:normAutofit").a.fontScale, 100000) / 100000 : 1;
  const shapeList = kid(txBody, "a:lstStyle");
  const chain = shapeCtx.styleChain;
  let hasText = false;
  const paragraphs = [];
  let autoNumber = 0;
  for (const p of kids(txBody, "a:p")) {
    const pPr = kid(p, "a:pPr");
    const level = num(pPr && pPr.a.lvl);
    const levelName = `a:lvl${level + 1}pPr`;
    const merged = mergeParagraph([...chain.map((lst) => kid(lst, levelName)), kid(shapeList, levelName), pPr]);
    const runs = [];
    let paraText = "";
    for (const child of p.c) {
      if (child.n === "a:br") { runs.push("<br>"); continue; }
      if (child.n !== "a:r" && child.n !== "a:fld") continue;
      const value = textOf(kid(child, "a:t"));
      if (!value) continue;
      paraText += value;
      const style = mergeRun({ attrs: { ...merged.defRPr.attrs }, fill: merged.defRPr.fill, latin: merged.defRPr.latin }, kid(child, "a:rPr"));
      runs.push(runHtml(value, style, fontScale, ctx, shapeCtx));
    }
    if (paraText.trim()) hasText = true;
    const endRPr = kid(p, "a:endParaRPr");
    const sizeEmu = num((endRPr && endRPr.a.sz) || merged.defRPr.attrs.sz, 1800) * fontScale * 127;
    const marL = num(merged.attrs.marL), indent = num(merged.attrs.indent);
    const align = { l: "left", ctr: "center", r: "right", just: "justify" }[merged.attrs.algn] || "left";
    let bullet = "";
    if (paraText.trim() && merged.bu && !merged.bu.none) {
      if (merged.bu.char) bullet = `<span class="bu">${esc(merged.bu.char)}&nbsp;</span>`;
      else if (merged.bu.auto) bullet = `<span class="bu">${++autoNumber}.&nbsp;</span>`;
    }
    const spaceBefore = merged.spcBef ? (merged.spcBef.pts ? merged.spcBef.pts * 12700 : (merged.spcBef.pct || 0) * sizeEmu) : 0;
    const lineHeight = (merged.lnSpc || 1) * 1.17;
    paragraphs.push(`<p style="text-align:${align};padding-left:${cqw(Math.max(0, marL), ctx)};text-indent:${cqw(indent, ctx)};margin:${cqw(spaceBefore, ctx)} 0 0;line-height:${round(lineHeight, 3)};min-height:1em;font-size:${cqw(sizeEmu, ctx)}">${bullet}${runs.join("") || "&nbsp;"}</p>`);
  }
  return { html: paragraphs.join(""), hasText, bodyPr };
}

function runHtml(value, style, fontScale, ctx, shapeCtx) {
  const attrs = style.attrs;
  const css = [];
  if (attrs.sz) css.push(`font-size:${cqw(num(attrs.sz) * fontScale * 127 * (num(attrs.baseline) ? 0.7 : 1), ctx)}`);
  if (attrs.b === "1") css.push("font-weight:700");
  if (attrs.i === "1") css.push("font-style:italic");
  if (attrs.u && attrs.u !== "none") css.push("text-decoration:underline");
  if (attrs.cap === "all") css.push("text-transform:uppercase");
  if (num(attrs.baseline) > 0) css.push("vertical-align:super");
  else if (num(attrs.baseline) < 0) css.push("vertical-align:sub");
  const color = colorFrom(style.fill, ctx) || shapeCtx.defaultColor;
  if (color) css.push(`color:${color}`);
  if (style.latin) css.push(`font-family:'${style.latin.replace(/'/g, "")}',Segoe UI,Arial,sans-serif`);
  return `<span style="${css.join(";")}">${esc(value)}</span>`;
}

/* ------------------------------------------------------------ SHAPES ---- */
function fillCss(spPr, style, ctx, part) {
  if (!spPr) spPr = { c: [] };
  if (kid(spPr, "a:noFill")) return "";
  const solid = kid(spPr, "a:solidFill");
  if (solid) return `background:${colorFrom(solid, ctx)}`;
  const grad = kid(spPr, "a:gradFill");
  if (grad) {
    const stops = kids(kid(grad, "a:gsLst"), "a:gs").map((gs) => `${colorFrom(gs, ctx)} ${num(gs.a.pos) / 1000}%`);
    const angle = num(kid(grad, "a:lin") && kid(grad, "a:lin").a.ang) / 60000;
    if (stops.length > 1) return `background:linear-gradient(${round(angle + 90, 1)}deg,${stops.join(",")})`;
    if (stops.length === 1) return `background:${stops[0].split(" ")[0]}`;
  }
  const blip = find(kid(spPr, "a:blipFill"), "a:blip");
  if (blip && blip.a["r:embed"]) {
    const url = ctx.deck.media(part, blip.a["r:embed"]);
    if (url) return `background:url('${url}') center/100% 100% no-repeat`;
  }
  const ref = style && kid(style, "a:fillRef");
  if (ref && num(ref.a.idx) > 0) {
    const color = colorFrom(ref, ctx);
    if (color) return `background:${color}`;
  }
  return "";
}
function lineCss(spPr, style, ctx) {
  const ln = kid(spPr, "a:ln");
  if (ln && kid(ln, "a:noFill")) return "";
  let color = ln && colorFrom(kid(ln, "a:solidFill"), ctx);
  let width = ln && ln.a.w ? num(ln.a.w) : null;
  if (!color) {
    const ref = style && kid(style, "a:lnRef");
    if (ref && num(ref.a.idx) > 0) { color = colorFrom(ref, ctx); width = width || 12700; }
  }
  if (!color) return "";
  const dash = ln && kid(ln, "a:prstDash") && /dash|dot/i.test(kid(ln, "a:prstDash").a.val) ? "dashed" : "solid";
  return `border:${cqw(Math.max(width || 9525, 6000), ctx)} ${dash} ${color}`;
}
function boxCss(xf, ctx, tf, extra = []) {
  const x = tf.ox + tf.sx * xf.x, y = tf.oy + tf.sy * xf.y, w = tf.sx * xf.cx, h = tf.sy * xf.cy;
  const css = [`left:${round((x / ctx.slideW) * 100, 4)}%`, `top:${round((y / ctx.slideH) * 100, 4)}%`, `width:${round((w / ctx.slideW) * 100, 4)}%`, `height:${round((h / ctx.slideH) * 100, 4)}%`];
  const transforms = [];
  if (xf.rot) transforms.push(`rotate(${round(xf.rot / 60000, 2)}deg)`);
  if (xf.flipH) transforms.push("scaleX(-1)");
  if (xf.flipV) transforms.push("scaleY(-1)");
  if (transforms.length) css.push(`transform:${transforms.join(" ")}`);
  return css.concat(extra).filter(Boolean).join(";");
}
const identity = { ox: 0, oy: 0, sx: 1, sy: 1 };

function renderSp(node, ctx, part, tf, layers) {
  const spPr = kid(node, "p:spPr");
  const ph = placeholderOf(node);
  let xf = xfrmOf(spPr);
  const layoutPh = ph ? findPlaceholder(layers.layoutTree, ph) : null;
  const masterPh = ph ? findPlaceholder(layers.masterTree, layoutPh ? placeholderOf(layoutPh) || ph : ph) : null;
  if (!xf && ph) xf = xfrmOf(kid(layoutPh, "p:spPr")) || xfrmOf(kid(masterPh, "p:spPr"));
  if (!xf) return "";
  const style = kid(node, "p:style");
  const geom = kid(spPr, "a:prstGeom");
  const preset = geom ? geom.a.prst : "rect";
  const fill = fillCss(spPr, style, ctx, part);
  const line = lineCss(spPr, style, ctx);
  const phType = ph ? ph.a.type || "body" : null;
  const txStyles = kid(ctx.masterRoot, "p:txStyles");
  const styleName = phType === "title" || phType === "ctrTitle" ? "p:titleStyle" : ph ? "p:bodyStyle" : "p:otherStyle";
  const styleChain = [kid(txStyles, styleName), kid(kid(masterPh, "p:txBody"), "a:lstStyle"), kid(kid(layoutPh, "p:txBody"), "a:lstStyle")].filter(Boolean);
  const fontRef = style && kid(style, "a:fontRef");
  const defaultColor = (fontRef && colorFrom(fontRef, ctx)) || null;
  const text = renderText(kid(node, "p:txBody"), ctx, { styleChain, defaultColor });
  if (!fill && !line && !text.hasText) return "";
  const bodyPr = text.bodyPr || { a: {}, c: [] };
  const inheritedBodyPr = kid(kid(layoutPh, "p:txBody"), "a:bodyPr") || kid(kid(masterPh, "p:txBody"), "a:bodyPr");
  const inset = (key, fallback) => cqw(num(bodyPr.a[key], num(inheritedBodyPr && inheritedBodyPr.a[key], fallback)), ctx);
  const anchorKey = bodyPr.a.anchor || (inheritedBodyPr && inheritedBodyPr.a.anchor);
  const anchor = { ctr: "center", b: "flex-end", t: "flex-start" }[anchorKey] || "flex-start";
  const radius = preset === "ellipse" ? "border-radius:50%" : preset === "roundRect" ? "border-radius:8%" : "";
  const clip = preset === "triangle" ? "clip-path:polygon(50% 0,100% 100%,0 100%)" : "";
  const wrap = bodyPr.a.wrap === "none" ? "white-space:nowrap" : "";
  const css = boxCss(xf, ctx, tf, [fill, line, radius, clip, `padding:${inset("tIns", 45720)} ${inset("rIns", 91440)} ${inset("bIns", 45720)} ${inset("lIns", 91440)}`, `justify-content:${anchor}`, wrap]);
  return `<div class="sh" style="${css}">${text.html}</div>`;
}

function renderPic(node, ctx, part, tf, overrideXfrm) {
  const xf = overrideXfrm || xfrmOf(kid(node, "p:spPr"));
  const blipFill = kid(node, "p:blipFill");
  const blip = find(blipFill, "a:blip");
  if (!xf || !blip || !blip.a["r:embed"]) return "";
  const url = ctx.deck.media(part, blip.a["r:embed"]);
  if (!url) return "";
  const src = kid(blipFill, "a:srcRect");
  const l = num(src && src.a.l) / 100000, t = num(src && src.a.t) / 100000, r = num(src && src.a.r) / 100000, b = num(src && src.a.b) / 100000;
  const w = 1 - l - r || 1, h = 1 - t - b || 1;
  const imgCss = `position:absolute;left:${round((-l / w) * 100, 3)}%;top:${round((-t / h) * 100, 3)}%;width:${round(100 / w, 3)}%;height:${round(100 / h, 3)}%`;
  return `<div class="sh pic" style="${boxCss(xf, ctx, tf, ["overflow:hidden"])}"><img src="${url}" alt="" style="${imgCss}"></div>`;
}

function renderLine(node, ctx, part, tf) {
  const spPr = kid(node, "p:spPr");
  const xf = xfrmOf(spPr);
  if (!xf) return "";
  const style = kid(node, "p:style");
  const ln = kid(spPr, "a:ln");
  const color = (ln && colorFrom(kid(ln, "a:solidFill"), ctx)) || colorFrom(kid(style, "a:lnRef"), ctx);
  if (!color) return "";
  const width = ((ln && ln.a.w ? num(ln.a.w) : 12700) / ctx.slideW) * 100;
  const css = boxCss({ ...xf, flipH: false, flipV: false }, ctx, tf);
  const x1 = xf.flipH ? "100%" : "0", x2 = xf.flipH ? "0" : "100%", y1 = xf.flipV ? "100%" : "0", y2 = xf.flipV ? "0" : "100%";
  return `<svg class="sh" style="${css};overflow:visible"><line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${color}" stroke-width="${round(width, 3)}cqw"${ln && kid(ln, "a:prstDash") ? ' stroke-dasharray="6 4"' : ""}/></svg>`;
}

/* ------------------------------------------------------------- TABLES ---- */
function renderTable(tbl, xf, ctx, part, tf) {
  const grid = kids(kid(tbl, "a:tblGrid"), "a:gridCol").map((col) => num(col.a.w));
  const totalWidth = grid.reduce((a, b) => a + b, 0) || xf.cx;
  const tblPr = kid(tbl, "a:tblPr") || { a: {}, c: [] };
  const styleId = textOf(kid(tblPr, "a:tableStyleId"));
  const mediumAccent1 = styleId === "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";
  const gridStyle = styleId === "{5940675A-B579-460E-94D1-54222C63F5DA}";
  const rows = kids(tbl, "a:tr").map((row, rowIndex) => {
    const cells = [];
    kids(row, "a:tc").forEach((cell) => {
      if (cell.a.hMerge === "1" || cell.a.vMerge === "1") return;
      const tcPr = kid(cell, "a:tcPr") || { a: {}, c: [] };
      let fill = fillCss(tcPr, null, ctx, part);
      let color = null;
      let bold = false;
      if (!fill && mediumAccent1) {
        if (tblPr.a.firstRow === "1" && rowIndex === 0) { fill = `background:${colorFrom(schemeHolder("accent1"), ctx)}`; color = "#fff"; bold = true; }
        else if (tblPr.a.bandRow === "1") fill = `background:${colorFrom(schemeHolder("accent1", [{ n: "a:tint", a: { val: rowIndex % 2 ? "40000" : "20000" }, c: [] }]), ctx)}`;
      }
      const borders = ["L", "R", "T", "B"].map((side) => {
        const ln = kid(tcPr, `a:ln${side}`);
        const lineColor = ln && colorFrom(kid(ln, "a:solidFill"), ctx);
        const name = { L: "left", R: "right", T: "top", B: "bottom" }[side];
        if (lineColor) return `border-${name}:${cqw(Math.max(num(ln.a.w, 6350), 6350), ctx)} solid ${lineColor}`;
        return gridStyle || mediumAccent1 ? `border-${name}:${cqw(6350, ctx)} solid rgba(120,120,120,.55)` : "";
      });
      const text = renderText(kid(cell, "a:txBody"), ctx, { styleChain: [], defaultColor: color || "#16243a" });
      const anchor = { ctr: "middle", b: "bottom" }[tcPr.a.anchor] || "top";
      const pad = `padding:${cqw(num(tcPr.a.marT, 45720), ctx)} ${cqw(num(tcPr.a.marR, 91440), ctx)} ${cqw(num(tcPr.a.marB, 45720), ctx)} ${cqw(num(tcPr.a.marL, 91440), ctx)}`;
      cells.push(`<td colspan="${num(cell.a.gridSpan, 1)}" rowspan="${num(cell.a.rowSpan, 1)}" style="${[fill, ...borders, pad, `vertical-align:${anchor}`, bold ? "font-weight:700" : ""].filter(Boolean).join(";")}">${text.html}</td>`);
    });
    return `<tr style="height:${cqw(num(row.a.h), ctx)}">${cells.join("")}</tr>`;
  });
  const cols = grid.map((w) => `<col style="width:${round((w / totalWidth) * 100, 3)}%">`).join("");
  return `<div class="sh" style="${boxCss(xf, ctx, tf)}"><table class="tbl"><colgroup>${cols}</colgroup>${rows.join("")}</table></div>`;
}

/* ------------------------------------------------------------- CHARTS ---- */
function ptList(node) {
  if (!node) return [];
  const cache = find(node, "c:strCache") || find(node, "c:numCache") || find(node, "c:lvl");
  if (!cache) return [];
  const count = num(kid(cache, "c:ptCount") && kid(cache, "c:ptCount").a.val);
  const values = new Array(count).fill("");
  for (const pt of kids(cache, "c:pt")) values[num(pt.a.idx)] = textOf(kid(pt, "c:v"));
  return values;
}
function niceScale(min, max) {
  if (min === max) max = min + 1;
  const rough = (max - min) / 5;
  const power = Math.pow(10, Math.floor(Math.log10(rough)));
  const step = [1, 2, 2.5, 5, 10].map((m) => m * power).find((s) => s >= rough) || power * 10;
  const lo = Math.floor(min / step) * step, hi = Math.ceil(max / step) * step;
  const ticks = [];
  for (let v = lo; v <= hi + step / 1000; v += step) ticks.push(round(v, 6));
  return { min: lo, max: hi, ticks };
}
function fmtNumber(value, code) {
  if (value === null || value === undefined || Number.isNaN(value)) return "";
  if (code && code.includes("%")) return `${round(value * 100, 1)}%`;
  const abs = Math.abs(value);
  if (abs >= 10000) return `${round(value / 1000, 1)}k`;
  if (abs >= 100) return String(Math.round(value));
  return String(round(value, abs < 1 ? 2 : 1));
}
const optionalNumber = (node) => (node ? num(node.a.val, null) : null);

function renderChart(chartRoot, wEmu, hEmu, ctx) {
  const W = Math.max(60, Math.round(wEmu / 12700)), H = Math.max(40, Math.round(hEmu / 12700));
  const chart = kid(chartRoot, "c:chart");
  const plot = kid(chart, "c:plotArea");
  const palette = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"].map((n) => colorFrom(schemeHolder(n), ctx));
  const groups = plot ? plot.c.filter((n) => /^c:(bar|line|area|pie|doughnut|scatter|bar3D|pie3D|line3D|area3D)Chart$/.test(n.n)) : [];
  const out = [];
  const titleText = findAll(kid(chart, "c:title"), "a:t").map(textOf).join("");
  const hasLegend = !!kid(chart, "c:legend");
  const T = titleText ? 22 : 8;
  if (titleText) out.push(`<text x="${W / 2}" y="15" text-anchor="middle" font-size="11" font-weight="700" fill="#16243a">${esc(titleText)}</text>`);
  let seriesIndex = 0;
  const allGroups = groups.map((g) => {
    const type = g.n.replace("c:", "").replace("Chart", "").replace("3D", "");
    const series = kids(g, "c:ser").map((s) => {
      const spPr = kid(s, "c:spPr");
      const color = colorFrom(kid(spPr, "a:solidFill"), ctx) || colorFrom(kid(kid(spPr, "a:ln"), "a:solidFill"), ctx) || palette[seriesIndex % palette.length];
      seriesIndex++;
      const points = {};
      for (const dPt of kids(s, "c:dPt")) {
        const c = colorFrom(kid(kid(dPt, "c:spPr"), "a:solidFill"), ctx);
        if (c) points[num(dPt.a.idx)] = c;
      }
      const dLbls = kid(s, "c:dLbls") || kid(g, "c:dLbls");
      const showVal = !!(dLbls && kid(dLbls, "c:showVal") && kid(dLbls, "c:showVal").a.val !== "0");
      const showPct = !!(dLbls && kid(dLbls, "c:showPercent") && kid(dLbls, "c:showPercent").a.val === "1");
      return {
        name: textOf(find(kid(s, "c:tx"), "c:v")), color, points, showVal, showPct,
        cats: ptList(kid(s, "c:cat")), vals: ptList(kid(s, "c:val")).map((v) => (v === "" ? null : Number(v))),
        code: textOf(find(kid(s, "c:val"), "c:formatCode"))
      };
    });
    return {
      type, series, grouping: kid(g, "c:grouping") ? kid(g, "c:grouping").a.val : "clustered",
      horizontal: !!(kid(g, "c:barDir") && kid(g, "c:barDir").a.val === "bar"), axIds: kids(g, "c:axId").map((a) => a.a.val)
    };
  }).filter((g) => g.series.length);
  if (!allGroups.length) {
    return `<svg viewBox="0 0 ${W} ${H}" style="width:100%;height:100%"><text x="${W / 2}" y="${H / 2}" text-anchor="middle" font-size="11" fill="#777">Chart type not supported - see source deck</text></svg>`;
  }
  const legendItems = [];
  const pieGroup = allGroups.find((g) => g.type === "pie" || g.type === "doughnut");
  if (pieGroup) {
    const s = pieGroup.series[0];
    const total = s.vals.reduce((a, v) => a + Math.max(0, v || 0), 0) || 1;
    const R = Math.min(W * 0.5, H - T - (hasLegend ? 24 : 8)) / 2 - 4;
    const cx = W / 2, cy = T + (H - T - (hasLegend ? 18 : 4)) / 2;
    let angle = -Math.PI / 2;
    s.vals.forEach((v, i) => {
      const share = Math.max(0, v || 0) / total;
      if (!share) return;
      const color = s.points[i] || palette[i % palette.length];
      legendItems.push({ name: s.cats[i] || `#${i + 1}`, color });
      const next = angle + share * Math.PI * 2;
      const p = (a, r) => `${round(cx + Math.cos(a) * r, 2)} ${round(cy + Math.sin(a) * r, 2)}`;
      out.push(share >= 0.9999 ? `<circle cx="${cx}" cy="${cy}" r="${R}" fill="${color}"/>` : `<path d="M ${cx} ${cy} L ${p(angle, R)} A ${R} ${R} 0 ${share > 0.5 ? 1 : 0} 1 ${p(next, R)} Z" fill="${color}" stroke="#fff" stroke-width="0.8"/>`);
      if (s.showVal || s.showPct) {
        const mid = (angle + next) / 2;
        out.push(`<text x="${round(cx + Math.cos(mid) * R * 0.65, 2)}" y="${round(cy + Math.sin(mid) * R * 0.65 + 3, 2)}" text-anchor="middle" font-size="8" fill="#fff" font-weight="700">${esc(s.showPct ? `${Math.round(share * 100)}%` : fmtNumber(v, s.code))}</text>`);
      }
      angle = next;
    });
    if (pieGroup.type === "doughnut") out.push(`<circle cx="${cx}" cy="${cy}" r="${R * 0.5}" fill="#fff"/>`);
  } else {
    // Cartesian charts (bar / line / area) with up to two value scales.
    const valAxes = kids(plot, "c:valAx");
    const axisOf = (id) => valAxes.find((ax) => kid(ax, "c:axId").a.val === id);
    const scaleFor = (g) => g.axIds.find((id) => axisOf(id)) || "primary";
    const scales = new Map();
    for (const g of allGroups) {
      const id = scaleFor(g);
      const entry = scales.get(id) || { min: 0, max: 0, stackedPct: false };
      const stacked = (g.type === "bar" && g.grouping !== "clustered") || (g.type === "area" && g.grouping !== "standard");
      if (g.grouping === "percentStacked") entry.stackedPct = true;
      const count = Math.max(...g.series.map((s) => s.vals.length));
      for (let i = 0; i < count; i++) {
        let pos = 0, neg = 0;
        for (const s of g.series) {
          const v = s.vals[i] || 0;
          if (stacked) { if (v >= 0) pos += v; else neg += v; entry.max = Math.max(entry.max, pos); entry.min = Math.min(entry.min, neg); }
          else { entry.max = Math.max(entry.max, v); entry.min = Math.min(entry.min, v); }
        }
      }
      scales.set(id, entry);
    }
    const scaleIds = [...scales.keys()];
    for (const id of scaleIds) {
      const entry = scales.get(id);
      const scaling = kid(axisOf(id), "c:scaling");
      const explicitMin = optionalNumber(kid(scaling, "c:min")), explicitMax = optionalNumber(kid(scaling, "c:max"));
      Object.assign(entry, entry.stackedPct ? niceScale(0, 1) : niceScale(explicitMin !== null ? explicitMin : entry.min, explicitMax !== null ? explicitMax : entry.max));
    }
    const horizontal = allGroups.some((g) => g.type === "bar" && g.horizontal);
    const categories = allGroups.reduce((best, g) => g.series.reduce((b, s) => (s.cats.length > b.length ? s.cats : b), best), []);
    const n = Math.max(categories.length, ...allGroups.map((g) => Math.max(...g.series.map((s) => s.vals.length))));
    const longLabels = !horizontal && (n > 8 || categories.some((c) => c.length > 9));
    const left = horizontal ? Math.min(90, 14 + Math.max(...categories.map((c) => c.length), 4) * 4.2) : 34;
    const right = scaleIds.length > 1 ? 34 : 12;
    const bottom = (horizontal ? 18 : longLabels ? 44 : 22) + (hasLegend ? 16 : 0);
    const x0 = left, x1 = W - right, y0 = T + 4, y1 = H - bottom;
    const primary = scales.get(scaleIds[0]);
    const mapVal = (entry, v) => (horizontal ? x0 + ((v - entry.min) / (entry.max - entry.min)) * (x1 - x0) : y1 - ((v - entry.min) / (entry.max - entry.min)) * (y1 - y0));
    primary.ticks.forEach((tick) => {
      const pos = mapVal(primary, tick);
      const label = fmtNumber(tick, primary.stackedPct ? "%" : "");
      out.push(horizontal
        ? `<line x1="${pos}" y1="${y0}" x2="${pos}" y2="${y1}" stroke="#e1e6ee"/><text x="${pos}" y="${y1 + 11}" text-anchor="middle" font-size="8" fill="#555">${label}</text>`
        : `<line x1="${x0}" y1="${pos}" x2="${x1}" y2="${pos}" stroke="#e1e6ee"/><text x="${x0 - 4}" y="${pos + 3}" text-anchor="end" font-size="8" fill="#555">${label}</text>`);
    });
    if (scaleIds.length > 1) {
      const secondary = scales.get(scaleIds[1]);
      secondary.ticks.forEach((tick) => out.push(`<text x="${x1 + 4}" y="${mapVal(secondary, tick) + 3}" font-size="8" fill="#555">${fmtNumber(tick, "")}</text>`));
    }
    const slot = (horizontal ? y1 - y0 : x1 - x0) / Math.max(n, 1);
    const centerOf = (i) => (horizontal ? y0 : x0) + slot * (i + 0.5);
    categories.forEach((label, i) => {
      const short = label.length > 14 ? `${label.slice(0, 13)}...` : label;
      if (horizontal) out.push(`<text x="${x0 - 4}" y="${centerOf(i) + 3}" text-anchor="end" font-size="8" fill="#555">${esc(short)}</text>`);
      else if (longLabels) out.push(`<text transform="translate(${centerOf(i)} ${y1 + 8}) rotate(-35)" text-anchor="end" font-size="8" fill="#555">${esc(short)}</text>`);
      else out.push(`<text x="${centerOf(i)}" y="${y1 + 12}" text-anchor="middle" font-size="8" fill="#555">${esc(short)}</text>`);
    });
    const zero = mapVal(primary, 0);
    out.push(horizontal ? `<line x1="${zero}" y1="${y0}" x2="${zero}" y2="${y1}" stroke="#889"/>` : `<line x1="${x0}" y1="${zero}" x2="${x1}" y2="${zero}" stroke="#889"/>`);

    const barGroups = allGroups.filter((g) => g.type === "bar");
    const barSlots = barGroups.reduce((a, g) => a + (g.grouping === "clustered" ? g.series.length : 1), 0) || 1;
    const barWidth = (slot * 0.72) / barSlots;
    let barOffset = 0;
    for (const g of barGroups) {
      const entry = scales.get(scaleFor(g));
      const clustered = g.grouping === "clustered";
      for (let i = 0; i < n; i++) {
        let pos = 0, neg = 0;
        const total = g.grouping === "percentStacked" ? g.series.reduce((a, x) => a + Math.abs(x.vals[i] || 0), 0) || 1 : 1;
        g.series.forEach((s, si) => {
          const raw = s.vals[i];
          if (raw === null || raw === undefined) return;
          const v = g.grouping === "percentStacked" ? raw / total : raw;
          const start = clustered ? 0 : v >= 0 ? pos : neg;
          const end = start + v;
          if (!clustered) { if (v >= 0) pos = end; else neg = end; }
          const a = mapVal(entry, start), b = mapVal(entry, end);
          const color = s.points[i] || s.color;
          const begin = centerOf(i) - (slot * 0.72) / 2 + (barOffset + (clustered ? si : 0)) * barWidth;
          if (horizontal) out.push(`<rect x="${round(Math.min(a, b), 2)}" y="${round(begin, 2)}" width="${round(Math.abs(b - a), 2)}" height="${round(barWidth, 2)}" fill="${color}"/>`);
          else out.push(`<rect x="${round(begin, 2)}" y="${round(Math.min(a, b), 2)}" width="${round(barWidth, 2)}" height="${round(Math.abs(b - a), 2)}" fill="${color}"/>`);
          if (s.showVal) {
            const label = fmtNumber(raw, s.code);
            if (horizontal) out.push(`<text x="${round(Math.max(a, b) + 2, 2)}" y="${round(begin + barWidth / 2 + 3, 2)}" font-size="7" fill="#333">${label}</text>`);
            else out.push(`<text x="${round(begin + barWidth / 2, 2)}" y="${round(clustered ? Math.min(a, b) - 2 : (a + b) / 2 + 3, 2)}" text-anchor="middle" font-size="7" fill="${clustered ? "#333" : "#fff"}">${label}</text>`);
          }
        });
      }
      barOffset += clustered ? g.series.length : 1;
    }
    for (const g of allGroups.filter((x) => x.type === "area" || x.type === "line")) {
      const entry = scales.get(scaleFor(g));
      for (const s of g.series) {
        const pts = s.vals.map((v, i) => (v === null || v === undefined ? null : [centerOf(i), mapVal(entry, v), v])).filter(Boolean);
        if (!pts.length) continue;
        const poly = pts.map((p) => `${round(p[0], 2)},${round(p[1], 2)}`).join(" ");
        if (g.type === "area") {
          const base = round(mapVal(entry, 0), 2);
          out.push(`<polygon points="${round(pts[0][0], 2)},${base} ${poly} ${round(pts[pts.length - 1][0], 2)},${base}" fill="${s.color}" fill-opacity=".55"/>`);
        } else {
          out.push(`<polyline points="${poly}" fill="none" stroke="${s.color}" stroke-width="2"/>`);
          pts.forEach((p) => out.push(`<circle cx="${round(p[0], 2)}" cy="${round(p[1], 2)}" r="2.4" fill="${s.color}"/>`));
        }
        if (s.showVal) pts.forEach((p) => out.push(`<text x="${round(p[0], 2)}" y="${round(p[1] - 5, 2)}" text-anchor="middle" font-size="7" fill="#333">${fmtNumber(p[2], s.code)}</text>`));
      }
    }
    for (const g of allGroups) for (const s of g.series) if (s.name) legendItems.push({ name: s.name, color: s.color });
  }
  if (hasLegend && legendItems.length) {
    const items = legendItems.slice(0, 12);
    const widths = items.map((i) => Math.min(i.name.length, 22) * 4.3 + 18);
    let x = Math.max(6, (W - widths.reduce((a, b) => a + b, 0)) / 2);
    items.forEach((item, i) => {
      out.push(`<rect x="${round(x, 1)}" y="${H - 12}" width="8" height="8" fill="${item.color}"/><text x="${round(x + 11, 1)}" y="${H - 5}" font-size="8" fill="#444">${esc(item.name.slice(0, 22))}</text>`);
      x += widths[i];
    });
  }
  return `<svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="xMidYMid meet" style="width:100%;height:100%;display:block" font-family="Segoe UI,Arial,sans-serif">${out.join("")}</svg>`;
}

/* ---------------------------------------------------------- FRAMES ------- */
function renderFrame(node, ctx, part, tf) {
  const xf = xfrmOf(node, "p:xfrm");
  if (!xf) return "";
  const data = find(node, "a:graphicData");
  const tbl = find(data, "a:tbl");
  if (tbl) return renderTable(tbl, xf, ctx, part, tf);
  const chartRef = find(data, "c:chart");
  if (chartRef && chartRef.a["r:id"]) {
    const rel = part.rels.get(chartRef.a["r:id"]);
    const chartRoot = rel ? ctx.deck.xml(rel.target) : null;
    if (chartRoot) return `<div class="sh chart" style="${boxCss(xf, ctx, tf)}">${renderChart(chartRoot, xf.cx, xf.cy, ctx)}</div>`;
  }
  const pic = find(data, "p:pic");
  if (pic) return renderPic(pic, ctx, part, tf, xf);
  return "";
}

function renderTree(spTree, ctx, part, tf, layers) {
  const out = [];
  for (const node of spTree ? spTree.c : []) {
    if (node.n === "p:sp") out.push(renderSp(node, ctx, part, tf, layers));
    else if (node.n === "p:pic") out.push(renderPic(node, ctx, part, tf));
    else if (node.n === "p:cxnSp") out.push(renderLine(node, ctx, part, tf));
    else if (node.n === "p:graphicFrame") out.push(renderFrame(node, ctx, part, tf));
    else if (node.n === "p:grpSp") {
      const gx = kid(kid(node, "p:grpSpPr"), "a:xfrm");
      const off = kid(gx, "a:off"), ext = kid(gx, "a:ext"), chOff = kid(gx, "a:chOff"), chExt = kid(gx, "a:chExt");
      if (!off || !ext || !chOff || !chExt) continue;
      const kx = num(chExt.a.cx) ? num(ext.a.cx) / num(chExt.a.cx) : 1, ky = num(chExt.a.cy) ? num(ext.a.cy) / num(chExt.a.cy) : 1;
      const mapped = { ox: num(off.a.x) - num(chOff.a.x) * kx, oy: num(off.a.y) - num(chOff.a.y) * ky, sx: kx, sy: ky };
      const next = { ox: tf.ox + tf.sx * mapped.ox, oy: tf.oy + tf.sy * mapped.oy, sx: tf.sx * mapped.sx, sy: tf.sy * mapped.sy };
      out.push(renderTree(node, ctx, part, next, layers));
    } else if (node.n === "mc:AlternateContent") {
      const choice = kid(node, "mc:Fallback") || kid(node, "mc:Choice");
      if (choice) out.push(renderTree(choice, ctx, part, tf, layers));
    }
  }
  return out.join("");
}

function backgroundCss(cSld, ctx, part) {
  const bg = kid(cSld, "p:bg");
  if (!bg) return "";
  const props = kid(bg, "p:bgPr");
  if (props) return fillCss(props, null, ctx, part);
  const color = colorFrom(kid(bg, "p:bgRef"), ctx);
  return color ? `background:${color}` : "";
}

/* ---------------------------------------------------------- PUBLIC ------- */
// Reads the deck and returns { width, height, slides, sections } with rendered HTML for every slide.
function renderDeck(buffer, mediaDir) {
  const deck = new Deck(buffer, mediaDir);
  const presentation = deck.xml("ppt/presentation.xml");
  if (!presentation) throw new Error("This file does not contain PowerPoint presentation data.");
  const size = kid(presentation, "p:sldSz");
  const slideW = num(size && size.a.cx, 12192000), slideH = num(size && size.a.cy, 6858000);
  const presRels = deck.rels("ppt/presentation.xml");
  const order = kids(kid(presentation, "p:sldIdLst"), "p:sldId").map((entry) => ({ id: entry.a.id, rel: presRels.get(entry.a["r:id"]) })).filter((entry) => entry.rel);
  if (!order.length) throw new Error("No slides were found in the uploaded .pptx file.");
  const sectionList = find(presentation, "p14:sectionLst");
  const sections = (sectionList ? kids(sectionList, "p14:section") : []).map((section) => ({
    name: section.a.name, ids: kids(kid(section, "p14:sldIdLst"), "p14:sldId").map((entry) => entry.a.id)
  }));

  const fallbackTheme = loadTheme(deck, "ppt/slideMasters/slideMaster1.xml");
  const layerCache = new Map();
  const slides = order.map((entry, index) => {
    const slidePath = entry.rel.target;
    const slideRoot = deck.xml(slidePath);
    const part = { path: slidePath, rels: deck.rels(slidePath) };
    const layoutRel = [...part.rels.values()].find((rel) => rel.type.endsWith("/slideLayout"));
    const layoutPath = layoutRel ? layoutRel.target : null;
    if (layoutPath && !layerCache.has(layoutPath)) {
      const layoutPart = { path: layoutPath, rels: deck.rels(layoutPath) };
      const masterRel = [...layoutPart.rels.values()].find((rel) => rel.type.endsWith("/slideMaster"));
      const masterPath = masterRel ? masterRel.target : null;
      const masterRoot = masterPath ? deck.xml(masterPath) : null;
      const masterPart = masterPath ? { path: masterPath, rels: deck.rels(masterPath) } : layoutPart;
      const clrMapNode = kid(masterRoot, "p:clrMap");
      const ctx = { deck, theme: masterPath ? loadTheme(deck, masterPath) : fallbackTheme, clrMap: clrMapNode ? clrMapNode.a : { bg1: "lt1", tx1: "dk1", bg2: "lt2", tx2: "dk2" }, slideW, slideH, masterRoot };
      const layoutRoot = deck.xml(layoutPath);
      const layoutTree = kid(kid(layoutRoot, "p:cSld"), "p:spTree");
      const masterTree = kid(kid(masterRoot, "p:cSld"), "p:spTree");
      const layers = { layoutTree, masterTree };
      const decorOnly = (tree) => ({ c: (tree ? tree.c : []).filter((n) => !placeholderOf(n)) });
      const showMaster = layoutRoot && layoutRoot.a.showMasterSp !== "0";
      const decor = (showMaster && masterTree ? renderTree(decorOnly(masterTree), ctx, masterPart, identity, layers) : "") + renderTree(decorOnly(layoutTree), ctx, layoutPart, identity, layers);
      const background = backgroundCss(kid(layoutRoot, "p:cSld"), ctx, layoutPart) || backgroundCss(kid(masterRoot, "p:cSld"), ctx, masterPart);
      layerCache.set(layoutPath, { ctx, layers, decor, background });
    }
    const layer = layoutPath ? layerCache.get(layoutPath) : null;
    const ctx = layer ? layer.ctx : { deck, theme: fallbackTheme, clrMap: { bg1: "lt1", tx1: "dk1", bg2: "lt2", tx2: "dk2" }, slideW, slideH, masterRoot: null };
    const layers = layer ? layer.layers : { layoutTree: null, masterTree: null };
    const cSld = kid(slideRoot, "p:cSld");
    const spTree = kid(cSld, "p:spTree");
    const background = backgroundCss(cSld, ctx, part) || (layer ? layer.background : "");
    const body = renderTree(spTree, ctx, part, identity, layers);
    const texts = [];
    let title = "";
    for (const node of spTree ? spTree.c : []) {
      if (node.n !== "p:sp") continue;
      const value = findAll(node, "a:t").map(textOf).join(" ").replace(/\s+/g, " ").trim();
      if (!value) continue;
      const ph = placeholderOf(node);
      if (!title && ph && /^(title|ctrTitle)$/.test(ph.a.type || "")) title = value;
      texts.push(value);
    }
    return {
      index: index + 1, id: entry.id, hidden: slideRoot.a.show === "0",
      title: title || texts[0] || `Slide ${index + 1}`, text: texts,
      html: `<div class="sl" style="aspect-ratio:${slideW}/${slideH};${background}">${slideRoot.a.showMasterSp !== "0" && layer ? layer.decor : ""}${body}</div>`
    };
  });
  return { width: slideW, height: slideH, slides, sections };
}

// Returns the sections (teams) stored in the deck with their slide counts, without rendering any slide.
function listSections(buffer) {
  const zip = openZip(buffer);
  const data = zip.read("ppt/presentation.xml");
  if (!data) throw new Error("This file does not contain PowerPoint presentation data.");
  const list = find(parseXml(data.toString("utf8")), "p14:sectionLst");
  return (list ? kids(list, "p14:section") : []).map((s) => ({ name: s.a.name, count: kids(kid(s, "p14:sldIdLst"), "p14:sldId").length }));
}

module.exports = { renderDeck, listSections, esc };

