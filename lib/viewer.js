"use strict";
/*
 * Builds the manager-style presentation page (index.html) from a rendered deck.
 * Layout: dark navy sidebar with the deck's own PowerPoint sections, one slide at a time.
 */
const fs = require("node:fs");
const path = require("node:path");
const { renderDeck, esc } = require("./pptx-render");
const { teamKey, prettyName } = require("./teams");

function formatDate(iso) {
  const d = new Date(`${iso}T00:00:00`);
  if (Number.isNaN(d.getTime())) return iso;
  return `${String(d.getDate()).padStart(2, "0")}-${["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][d.getMonth()]}-${d.getFullYear()}`;
}

const CSS = `
:root{--navy:#0b1638;--blue:#16367f;--cyan:#57d6ff;--paper:#e9eef5;--side:230px}
*{box-sizing:border-box}html,body{height:100%}
body{margin:0;background:var(--paper);color:#16243a;font-family:"Segoe UI",Arial,sans-serif}
.sidebar{position:fixed;inset:0 auto 0 0;width:var(--side);padding:14px 8px 20px;background:linear-gradient(180deg,#091332,#102d63);color:#fff;overflow:auto;z-index:5}
.brand{padding:6px 8px 14px}.brand b{display:block;font-size:14px;letter-spacing:.08em}.brand span{display:block;color:#bfe7f4;font-size:11px;margin-top:4px}.brand .dt{color:#57d6ff;font-weight:600}
.back{display:block;margin:0 8px 10px;color:var(--cyan);font-size:12px;text-decoration:none}
.nav{display:flex;justify-content:space-between;gap:6px;width:100%;padding:9px;margin:3px 0;border:0;border-radius:7px;background:none;color:#fff;font:inherit;font-size:12px;text-align:left;cursor:pointer}
.nav:hover,.nav.active{background:#1d4f8e}.nav small{color:#a9dbe8;white-space:nowrap}
.opt{margin:14px 8px 0;font-size:12px;color:#cfe}.opt label{display:flex;gap:6px;align-items:center;cursor:pointer}
.main{margin-left:var(--side);min-height:100vh;padding:14px 22px 70px}
.bar{display:flex;justify-content:space-between;align-items:center;gap:10px;margin-bottom:10px}
.bar h1{margin:0;font-size:18px;color:#073b63}.bar span{font-size:12px;color:#668096}
.stage{background:#fff;border:1px solid #cfdbe6;box-shadow:0 6px 20px #16324d1f;max-width:1400px;margin:0 auto}
.sl{position:relative;width:100%;container-type:inline-size;overflow:hidden;background:#fff}
.sh{position:absolute;box-sizing:border-box;display:flex;flex-direction:column;overflow:hidden}
.sh p{margin:0;word-wrap:break-word}.sh.chart,.sh.pic{display:block}
.tbl{width:100%;height:100%;table-layout:fixed;border-collapse:collapse}.tbl td{overflow:hidden}
.bu{display:inline-block}
.ctl{position:fixed;right:22px;bottom:14px;z-index:6;display:flex;gap:8px;align-items:center}
.ctl button{min-width:38px;height:34px;border:0;border-radius:7px;background:var(--blue);color:#fff;font-size:18px;cursor:pointer;padding:0 10px}
.cnt{padding:8px 12px;border-radius:7px;background:#fff;font-size:12px;font-weight:700;box-shadow:0 3px 12px #16324d22}
.menu{display:none}
@media(max-width:820px){:root{--side:0px}.sidebar{transform:translateX(-100%);width:240px;transition:transform .2s}.sidebar.open{transform:none}.main{padding:10px}.menu{display:inline-block;border:0;background:var(--blue);color:#fff;border-radius:6px;padding:6px 10px}}/* corporate polish */
.sidebar{background:linear-gradient(180deg,#061433,#0a2f6b);border-right:1px solid #ffffff14;box-shadow:2px 0 12px #00000033}
.brand{border-bottom:1px solid #ffffff1f;margin-bottom:10px;padding-bottom:14px}
.brand b{font-size:15px;font-weight:700}.brand .dt{color:#7db4ff}
.back{color:#9cc4ff}.back:hover{color:#fff}
.nav{border-radius:6px;padding:10px 12px;color:#dbe6f7;transition:background .15s}
.nav:hover{background:#ffffff14}
.nav.active{background:#ffffff1f;color:#fff;font-weight:600;box-shadow:inset 3px 0 0 #4f9bff}
.nav small{color:#8fb3e6;font-variant-numeric:tabular-nums}
.bar{padding:6px 2px 10px;border-bottom:2px solid #d3deec;margin-bottom:14px}
.bar h1{font-size:20px;font-weight:600;color:#0a2a63}.bar span{font-weight:600;color:#5a6f8f}
.stage{border:0;border-radius:8px;overflow:hidden;box-shadow:0 8px 28px #0b1f4d26}
.ctl{background:#ffffffee;padding:6px;border-radius:10px;box-shadow:0 4px 18px #0b1f4d33}
.ctl button{background:#0a3d91;border-radius:6px;transition:background .15s}.ctl button:hover{background:#08306f}
.cnt{box-shadow:none;background:#eef2f8;color:#0a2a63}
.nav:focus-visible,.ctl button:focus-visible{outline:2px solid #4f9bff;outline-offset:2px}

`;

const SCRIPT = `
(function(){
var slides=[].slice.call(document.querySelectorAll(".slide"));var navs=[].slice.call(document.querySelectorAll(".nav"));
var cnt=document.getElementById("cnt"),title=document.getElementById("ttl"),hid=document.getElementById("hid"),cur=0;
function visible(){return slides.filter(function(s){return hid.checked||s.dataset.hidden!=="1"});}
function show(i){var v=visible();if(!v.length)return;cur=(i+v.length)%v.length;slides.forEach(function(s){s.style.display="none"});
 var s=v[cur];s.style.display="block";cnt.textContent=(cur+1)+" / "+v.length;title.textContent=s.dataset.title;
 navs.forEach(function(n){var a=+n.dataset.a,b=+n.dataset.b,k=+s.dataset.i;n.classList.toggle("active",k>=a&&k<=b)});
 var sb=document.querySelector(".sidebar");sb.classList.remove("open");history.replaceState(null,"","#"+s.dataset.i);}
function go(k){var v=visible();for(var j=0;j<v.length;j++){if(+v[j].dataset.i>=k){show(j);return}}show(v.length-1);}
function labels(){var v=visible();navs.forEach(function(n){var a=+n.dataset.a,b=+n.dataset.b,f=-1,l=-1;v.forEach(function(s,j){var k=+s.dataset.i;if(k>=a&&k<=b){if(f<0)f=j+1;l=j+1}});n.style.display=f<0?"none":"";n.querySelector("small").textContent=f<0?"":(f===l?f:f+"\u2013"+l)})}
navs.forEach(function(n){n.addEventListener("click",function(){go(+n.dataset.a)})});
document.getElementById("prev").onclick=function(){show(cur-1)};document.getElementById("next").onclick=function(){show(cur+1)};
document.getElementById("fs").onclick=function(){document.fullscreenElement?document.exitFullscreen():document.documentElement.requestFullscreen()};
document.getElementById("menu").onclick=function(){document.querySelector(".sidebar").classList.toggle("open")};
hid.onchange=function(){labels();var k=+(slides[0]&&visible()[cur]?visible()[cur].dataset.i:1);show(0);go(k)};
document.addEventListener("keydown",function(e){if(e.key==="ArrowLeft"||e.key==="PageUp")show(cur-1);if(e.key==="ArrowRight"||e.key==="PageDown"||e.key===" ")show(cur+1);if(e.key==="Home")show(0);if(e.key==="End")show(-1)});
var h=parseInt(location.hash.slice(1),10);labels();show(0);if(h)go(h);
})();`;

/* ------------------------------------------------------------------ STATE --- */
// A session folder holds state.json (all teams, their slides or links), media/ and the generated index.html.
const stateFile = (dir) => path.join(dir, "state.json");
function loadState(dir) {
  return fs.existsSync(stateFile(dir)) ? JSON.parse(fs.readFileSync(stateFile(dir), "utf8")) : null;
}
function saveState(dir, state) {
  state.updatedAt = new Date().toISOString();
  fs.writeFileSync(stateFile(dir), JSON.stringify(state), "utf8");
}

// Renders a deck and returns its sections as teams. Media goes to media/<tag>/ so later uploads never collide.
function splitDeck(fileData, dir, tag) {
  const deck = renderDeck(fileData, path.join(dir, "media", tag));
  const fix = (html) => html.replace(/(src="|url\(')media\//g, `$1media/${tag}/`);
  const byId = new Map(deck.slides.map((slide) => [slide.id, slide]));
  let sections = deck.sections
    .map((s) => ({ name: s.name, key: teamKey(s.name), slides: s.ids.map((id) => byId.get(id)).filter(Boolean) }))
    .filter((s) => s.slides.length);
  if (!sections.length) sections = [{ name: "Slides", key: "SLIDES", slides: deck.slides }];
  return sections.map((s) => ({ name: s.name, key: s.key, slides: s.slides.map((x) => ({ title: x.title, hidden: x.hidden, html: fix(x.html) })) }));
}

const labelFor = (key, fallback) => fallback ? String(fallback).replace(/\s+/g, " ").trim() : prettyName(key);

// Creates a session from a full deck: only the ticked teams are kept, in the manager's order.
function createSession(fileData, fileName, month, year, session, dir, options = {}) {
  fs.mkdirSync(dir, { recursive: true });
  const parts = splitDeck(fileData, dir, `b${Date.now()}`);
  const state = { month, year, session, meetingDate: options.meetingDate || "", agenda: options.agenda || "", sourceFile: fileName, generatedAt: new Date().toISOString(), home: null, order: [], teams: {} };
  const homePart = parts.find((p) => p.key === "EHEQ");
  if (homePart) state.home = homePart.slides;
  else if (parts[0]) state.home = [parts[0].slides.shift()];
  const wanted = Array.isArray(options.teams) && options.teams.length ? options.teams : parts.map((p) => p.key).filter((k) => k !== "EHEQ");
  for (const key of wanted) {
    const part = parts.find((p) => p.key === key);
    if (!part || !part.slides.length) continue;
    state.teams[key] = { label: labelFor(key, part.name), source: "pptx", fileName, updatedAt: new Date().toISOString(), slides: part.slides };
    state.order.push(key);
  }
  if (!state.order.length) throw new Error("None of the selected teams were found in this file.");
  saveState(dir, state);
  return buildPage(dir, state);
}

// Adds or replaces one team from a late PPTX (only that team's section is used when the file has sections).
function setTeamPptx(dir, key, fileData, fileName) {
  const state = requireState(dir);
  const parts = splitDeck(fileData, dir, `u${Date.now()}`);
  const match = parts.find((p) => p.key === key);
  const slides = match ? match.slides : parts.filter((p) => p.key !== "EHEQ").flatMap((p) => p.slides);
  if (!slides.length) throw new Error("No slides were found in that file for this team.");
  const old = state.teams[key];
  state.teams[key] = { label: match ? labelFor(key, match.name) : (old ? old.label : labelFor(key)), source: "pptx", fileName, updatedAt: new Date().toISOString(), slides };
  if (!state.order.includes(key)) state.order.push(key);
  saveState(dir, state);
  return buildPage(dir, state);
}

// A team that only supplied a SharePoint (or other web) link.
function setTeamLink(dir, key, link) {
  if (!/^https?:\/\/\S+$/i.test(link || "")) throw new Error("Please paste a full link starting with http:// or https://");
  const state = requireState(dir);
  const old = state.teams[key];
  state.teams[key] = { label: old ? old.label : labelFor(key), source: "link", link, updatedAt: new Date().toISOString(), slides: [] };
  if (!state.order.includes(key)) state.order.push(key);
  saveState(dir, state);
  return buildPage(dir, state);
}

// A team that is on the agenda but has not delivered yet.
function setTeamPending(dir, key) {
  const state = requireState(dir);
  const old = state.teams[key];
  state.teams[key] = { label: old ? old.label : labelFor(key), source: "pending", updatedAt: new Date().toISOString(), slides: [] };
  if (!state.order.includes(key)) state.order.push(key);
  saveState(dir, state);
  return buildPage(dir, state);
}

function removeTeam(dir, key) {
  const state = requireState(dir);
  delete state.teams[key];
  state.order = state.order.filter((k) => k !== key);
  if (!state.order.length) throw new Error("A review needs at least one team.");
  saveState(dir, state);
  return buildPage(dir, state);
}

function setOrder(dir, order) {
  const state = requireState(dir);
  const known = order.filter((k) => state.teams[k]);
  state.order = [...known, ...state.order.filter((k) => !known.includes(k))];
  saveState(dir, state);
  return buildPage(dir, state);
}

function setAgenda(dir, text) {
  const state = requireState(dir);
  state.agenda = String(text || "").slice(0, 3000);
  saveState(dir, state);
  return buildPage(dir, state);
}

function setMeetingDate(dir, date) {
  const state = requireState(dir);
  state.meetingDate = date || "";
  saveState(dir, state);
  return buildPage(dir, state);
}

function requireState(dir) {
  const state = loadState(dir);
  if (!state) throw new Error("This review has not been generated yet. Upload a Session file first.");
  return state;
}

/* ------------------------------------------------------------------- PAGE --- */
// A simple 16:9 slide shown for teams that sent a link or have not delivered yet.
function placeholderSlide(team) {
  const link = team.source === "link";
  const body = link
    ? `<p style="font-size:2.2cqw;margin-top:2cqw">This team shared its update as a link.</p><a href="${esc(team.link)}" target="_blank" rel="noopener" style="display:inline-block;margin-top:3cqw;padding:1.2cqw 3cqw;border-radius:.8cqw;background:#57d6ff;color:#0b1638;font-size:2.4cqw;font-weight:700;text-decoration:none">Open in SharePoint</a>`
    : `<p style="font-size:2.2cqw;margin-top:2cqw;color:#bcd">Awaiting slides from this team.</p>`;
  return `<div class="sl" style="aspect-ratio:16/9;background:linear-gradient(120deg,#0b1638,#16367f)"><div class="sh" style="left:8%;top:25%;width:84%;height:50%;align-items:center;justify-content:center;text-align:center;color:#fff;overflow:visible"><p style="font-size:4.4cqw;font-weight:700">${esc(team.label)}</p>${body}</div></div>`;
}

// Applies month, year and meeting date to the cover slide text.
function updateCover(html, state) {
  const sessionNo = /2/.test(state.session) ? 2 : 1;
  const parsed = state.meetingDate ? new Date(`${state.meetingDate}T00:00:00`) : null;
  const longDate = parsed && !Number.isNaN(parsed.getTime()) ? parsed.toLocaleDateString("en-GB", { day: "numeric", month: "long", year: "numeric" }) : "";
  return html
    .replace(/(January|February|March|April|May|June|July|August|September|October|November|December)\s+20\d\d/gi, (m) => (m === m.toUpperCase() ? `${state.month} ${state.year}`.toUpperCase() : `${state.month} ${state.year}`))
    .replace(new RegExp(`(Session ${sessionNo}:\\s*)[^<]*`), `$1${longDate || "DD Month YYYY"}`)
    .replace(new RegExp(`(Session ${3 - sessionNo}:\\s*)[^<]*`), "$1DD Month YYYY");
}

// Agenda page built from the lines the manager typed (e.g. "INFQ - 30 min").
function agendaSlide(state, dateText) {
  const lines = String(state.agenda || "").split(/\r?\n/).map((l) => l.trim()).filter(Boolean).slice(0, 12);
  const rows = lines.map((line, i) => `<div style="display:flex;align-items:center;gap:1.6cqw;padding:.9cqw 0;border-bottom:1px solid #ffffff30"><span style="flex:none;width:3.4cqw;height:3.4cqw;border-radius:50%;background:#57d6ff;color:#0b1638;font-weight:700;font-size:1.8cqw;display:grid;place-items:center">${i + 1}</span><span style="font-size:2.3cqw;font-weight:600">${esc(line)}</span></div>`).join("");
  return `<div class="sl" style="aspect-ratio:16/9;background:linear-gradient(120deg,#0b1638,#16367f)"><div class="sh" style="left:7%;top:8%;width:86%;height:86%;color:#fff;overflow:hidden"><p style="font-size:4cqw;font-weight:700">Agenda</p><p style="font-size:1.9cqw;color:#9cc4ff;margin:.6cqw 0 2cqw">${esc(state.session)} &middot; ${esc(dateText)}</p>${rows}</div></div>`;
}

function closingSlide(state, dateText) {
  return `<div class="sl" style="aspect-ratio:16/9;background:linear-gradient(120deg,#06163e,#0a3d91)"><div class="sh" style="left:8%;top:22%;width:84%;height:56%;display:flex;align-items:center;justify-content:center;text-align:center;color:#fff;overflow:hidden"><div><p style="font-size:5cqw;font-weight:700;margin:0 0 1.5cqw">Thank You</p><p style="font-size:2.1cqw;color:#bcd8ff;margin:0">${esc(state.session)} &middot; ${esc(dateText)}</p></div></div></div>`;
}

function arrangeEndingSections(sections, state, dateText) {
  const isClosing = (section) => /\bclosing\b|\bthank\s*you\b/i.test(section.name)
    || section.slides.some((slide) => /\bclosing\b|\bthank\s*you\b/i.test(String(slide.title || "")));
  const closing = sections.filter(isClosing);
  const targets = sections.filter((section) => /\btarget\b|\btgt\b/i.test(section.name) && !closing.includes(section));
  const body = sections.filter((section) => !closing.includes(section) && !targets.includes(section));
  const ending = targets.concat(closing);
  if (!closing.length) ending.push({ name: "Closing", slides: [{ title: "Thank You", hidden: false, html: closingSlide(state, dateText) }] });
  return body.concat(ending);
}

// Writes index.html and meta.json from the saved state.
function buildPage(dir, state) {
  const { month, year, session } = state;
  const sections = [];
  if (state.home && state.home.length) {
    sections.push({ name: "Home", slides: [state.home[0]] });
    const typedAgenda = String(state.agenda || "").trim();
    if (typedAgenda) sections.push({ name: "Agenda", slides: [{ title: "Agenda", hidden: false, html: agendaSlide(state, state.meetingDate ? formatDate(state.meetingDate) : `${month} ${year}`) }] });
    else if (state.home.length > 1) sections.push({ name: "Agenda", slides: state.home.slice(1) });
  }
  for (const key of state.order) {
    const team = state.teams[key];
    if (!team) continue;
    const slides = team.source === "pptx" ? team.slides : [{ title: team.label, hidden: false, html: placeholderSlide(team) }];
    if (slides.length) sections.push({ name: team.label, slides, source: team.source });
  }
  const dateText = state.meetingDate ? formatDate(state.meetingDate) : `${month} ${year}`;
  const arrangedSections = arrangeEndingSections(sections, state, dateText);
  let counter = 0;
  const all = [];
  for (const section of arrangedSections) for (const slide of section.slides) all.push({ ...slide, index: ++counter, section });
  const cover = all[0];
  if (cover && state.home) cover.html = updateCover(cover.html, state);
  const nav = arrangedSections.map((section) => {
    const own = all.filter((s) => s.section === section);
    return `<button class="nav" data-a="${own[0].index}" data-b="${own[own.length - 1].index}"><span>${esc(section.name)}${section.source === "link" ? " &#128279;" : ""}${section.source === "pending" ? " (pending)" : ""}</span><small></small></button>`;
  }).join("");
  const slides = all.map((slide) =>
    `<section class="slide" style="display:none" data-i="${slide.index}" data-hidden="${slide.hidden ? 1 : 0}" data-title="${esc(String(slide.title).slice(0, 140))}">${slide.html}</section>`).join("");
  const displayDateText = state.meetingDate ? formatDate(state.meetingDate) : `${month} ${year} (meeting date not set)`;
  const hiddenCount = all.filter((s) => s.hidden).length;
  const html = `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>EHEQ Monthly Review | ${esc(month)} ${esc(year)} | ${esc(session)}</title><style>${CSS}</style></head><body>
<aside class="sidebar"><div class="brand"><b>EHE QUALITY REVIEW</b><span>Interactive presentation webpage</span><span class="dt">${esc(session)} &middot; ${esc(displayDateText)}</span></div><a class="back" href="/">&larr; Back to generator</a>${nav}
<div class="opt"><label><input type="checkbox" id="hid"> Show hidden slides (${hiddenCount})</label></div></aside>
<main class="main"><div class="bar"><button class="menu" id="menu">&#9776; Sections</button><h1 id="ttl"></h1><span>${esc(session)} &middot; ${esc(displayDateText)}</span></div><div class="stage">${slides}</div></main>
<div class="ctl"><button id="prev" aria-label="Previous">&#8249;</button><div class="cnt" id="cnt"></div><button id="next" aria-label="Next">&#8250;</button><button id="fs" aria-label="Fullscreen">&#x26F6;</button></div>
<script>${SCRIPT}</script></body></html>`;
  fs.writeFileSync(path.join(dir, "index.html"), html, "utf8");
  const meta = { sourceFile: state.sourceFile, month, year, session, meetingDate: state.meetingDate, generatedAt: state.generatedAt, updatedAt: new Date().toISOString(), slideCount: all.length, hiddenSlides: hiddenCount, sections: arrangedSections.map((s) => ({ name: s.name, slides: s.slides.length })) };
  fs.writeFileSync(path.join(dir, "meta.json"), JSON.stringify(meta, null, 2), "utf8");
  return meta;
}

module.exports = { createSession, setTeamPptx, setTeamLink, setTeamPending, removeTeam, setOrder, setMeetingDate, setAgenda, loadState };