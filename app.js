const months = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const now = new Date();
const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => [...document.querySelectorAll(selector)];
let review = createReview(months[now.getMonth()], String(now.getFullYear()));
let importedSlides = [];
let currentSlide = 0;

function createReview(month, year) {
  return { month, year, session1: { date: "", agenda: "", teams: [] }, session2: { enabled: true, date: "", agenda: "", teams: [] }, teams: [], files: [], importedFolder: null, history: [], updatedAt: "" };
}
function key() { return `eheq-unified-review-${review.month}-${review.year}`; }
function showMessage(text, error = false) { const node = $("#message"); node.textContent = text; node.className = error ? "message error" : "message"; if (text) setTimeout(() => { node.textContent = ""; node.className = "message"; }, 5000); }
function escapeHtml(value) { return String(value || "").replace(/[&<>"']/g, (c) => ({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;" }[c])); }
function clearImportedSlides() { importedSlides.forEach((slide) => { if (slide.url.startsWith("blob:")) URL.revokeObjectURL(slide.url); }); importedSlides = []; currentSlide = 0; }
async function uploadToServer(file, session) {
  if (location.protocol === "file:") return null;
  const form = new FormData();
  form.append("file", file);
  form.append("month", review.month);
  form.append("year", review.year);
  form.append("session", session);
  const response = await fetch("/api/uploads", { method: "POST", body: form });
  const result = await response.json();
  if (!response.ok) throw new Error(result.error || "The server rejected the upload.");
  return result.record;
}
function save() { review.updatedAt = new Date().toISOString(); review.history.unshift({ at: review.updatedAt, change: "Monthly review updated" }); localStorage.setItem(key(), JSON.stringify(review)); showMessage(`Saved ${review.month} ${review.year}.`); render(); }
function loadPeriod() { clearImportedSlides(); const stored = localStorage.getItem(key()); review = stored ? JSON.parse(stored) : createReview($("#monthSelect").value, $("#yearSelect").value); render(); }
function fillSelectors() { $("#monthSelect").innerHTML = months.map((m) => `<option>${m}</option>`).join(""); $("#yearSelect").innerHTML = Array.from({length:11}, (_, i) => now.getFullYear()-5+i).map((y) => `<option>${y}</option>`).join(""); $("#monthSelect").value=review.month; $("#yearSelect").value=review.year; }
function render() { fillSelectors(); $("#dashboardTitle").textContent=`${review.month} ${review.year}`; $("#openPresentationLink").href=`/?month=${encodeURIComponent(review.month)}&year=${encodeURIComponent(review.year)}&session=1`; $("#openSession2PresentationLink").href=`/?month=${encodeURIComponent(review.month)}&year=${encodeURIComponent(review.year)}&session=2`; renderDashboard(); renderForms(); renderTeams(); renderFiles(); renderPreview(); renderHistory(); }
function renderDashboard() {
  const session2Ready = review.session2.date && review.session2.agenda.trim();
  $("#dashboardCards").innerHTML = [`<div class="card"><span>Session 1</span><strong class="${review.session1.date ? "ok":""}">${review.session1.date ? "Ready":"Pending"}</strong></div>`,`<div class="card"><span>Session 2</span><strong class="${session2Ready ? "ok":""}">${session2Ready ? "Ready":"Pending"}</strong></div>`,`<div class="card"><span>Selected teams</span><strong>${review.session1.teams.length + review.session2.teams.length}</strong></div>`,`<div class="card"><span>Imported files</span><strong>${review.files.length}</strong></div>`].join("");
  $("#readiness").innerHTML = [["Session 1 date",!!review.session1.date],["Session 1 agenda",!!review.session1.agenda.trim()],["Session 1 teams",review.session1.teams.length > 0],["Session 2 date",!!review.session2.date],["Session 2 agenda",!!review.session2.agenda.trim()],["Presentation slides",importedSlides.length > 0]].map(([label,ok])=>`<div class="readiness-row"><span>${label}</span><b class="${ok ? "ok":"pending"}">${ok ? "Complete":"Pending"}</b></div>`).join("");
  const missing = !review.session1.date ? "Session 1 date is still required." : !review.session1.agenda.trim() ? "Session 1 agenda is still required." : !review.session1.teams.length ? "Select at least one Session 1 team." : !review.session2.date ? "Session 2 date is still required." : !review.session2.agenda.trim() ? "Session 2 agenda is still required." : "The review is ready for the next step.";
  $("#assistantMessage").textContent = missing;
}
function renderForms() {
  $("#session1Date").value=review.session1.date; $("#session1Agenda").value=review.session1.agenda; $("#session2Date").value=review.session2.date; $("#session2Agenda").value=review.session2.agenda;
  const renderChecks=(id,selected)=>{ const node=$(id); node.innerHTML=review.teams.map((team)=>(`<label><input type="checkbox" value="${escapeHtml(team)}" ${selected.includes(team) ? "checked":""}>${escapeHtml(team)}</label>`)).join(""); };
  renderChecks("#session1Teams",review.session1.teams); renderChecks("#session2Teams",review.session2.teams); $("#noTeamsHint").style.display=review.teams.length ? "none":"block";
}
function renderTeams() { $("#teamList").innerHTML=review.teams.length ? review.teams.map((team)=>`<div class="item-row"><span>${escapeHtml(team)}</span><button class="button secondary" data-remove-team="${escapeHtml(team)}">Remove</button></div>`).join("") : `<p class="muted">No teams configured yet.</p>`; }
function renderFiles() {
  $("#pptFiles").innerHTML=review.files.length ? review.files.map((file,index)=>`<div class="file-row"><div><b>${escapeHtml(file.name || file.fileName)}</b><small>${escapeHtml(file.status || "Recorded")} · ${file.updatedAt ? `Updated ${new Date(file.updatedAt).toLocaleString()}` : file.uploadedAt ? `Uploaded ${new Date(file.uploadedAt).toLocaleString()}` : "Recorded file reference"}</small></div><div class="file-placement"><select data-file-session="${index}"><option value="">Select session</option><option value="Session 1" ${file.session==="Session 1"?"selected":""}>Session 1</option><option value="Session 2" ${file.session==="Session 2"?"selected":""}>Session 2</option></select><select data-file-team="${index}"><option value="">Select team</option>${review.teams.map((team)=>`<option value="${escapeHtml(team)}" ${file.team===team?"selected":""}>${escapeHtml(team)}</option>`).join("")}</select><button class="button secondary" type="button" data-remove-file="${index}">Remove</button></div></div>`).join("") : `<p class="muted">No team files recorded.</p>`;
}
async function removeFile(index) {
  const file = review.files[index];
  if (!file || !window.confirm(`Remove ${file.name || file.fileName}?`)) return;
  if (file.id && location.protocol !== "file:") {
    const response = await fetch(`/api/uploads/${encodeURIComponent(file.id)}`, { method: "DELETE" });
    const result = await response.json();
    if (!response.ok) return showMessage(result.error || "The uploaded presentation could not be removed.", true);
  }
  review.files.splice(index, 1);
  save();
  showMessage("Uploaded presentation removed.");
}
function renderPreview() {
  $("#previewSummary").innerHTML=`<div class="card"><span>Slides imported</span><strong>${importedSlides.length}</strong></div><div class="card"><span>Updated</span><strong>${review.updatedAt ? new Date(review.updatedAt).toLocaleDateString():"Not saved"}</strong></div>`;
  $("#previewPanel").innerHTML=importedSlides.length ? `<div class="preview-grid">${importedSlides.slice(0,12).map((s,i)=>`<img src="${s.url}" alt="Slide ${i+1}">`).join("")}</div><p class="muted">Showing the first 12 slides from the manager-provided presentation. Open fullscreen preview for presentation mode.</p>` : `<div class="preview-empty">Upload a monthly PowerPoint to preview its slides here.</div>`;
}
function renderHistory() { $("#historyList").innerHTML=review.history.length ? review.history.slice(0,30).map((item)=>`<div class="history-row"><span>${escapeHtml(item.change)}</span><small>${new Date(item.at).toLocaleString()}</small></div>`).join("") : `<p class="muted">No saved updates yet.</p>`; }
function loadManagerData() {
  review=createReview("May","2026");
  review.session1={date:"2026-05-13",agenda:"AGENDA from manager-provided EHEQ Monthly Review_2026_May_S1 13May.pptx",teams:["EEQ","INFIELD QUALITY_SA","INFIELD QUALITY_IAP","INFIELD QUALITY_CHN","INFIELD QUALITY_NA","INFIELD QUALITY_EMEA"]};
  review.session2={enabled:true,date:"2026-05-27",agenda:"Session 2 from manager-provided presentation",teams:[]};
  review.teams=["EEQ","INFIELD QUALITY_SA","INFIELD QUALITY_IAP","INFIELD QUALITY_CHN","INFIELD QUALITY_NA","INFIELD QUALITY_EMEA"];
  review.files=[{name:"EHEQ Monthly Review_2026_May_S1 13May.pptx",session:"Session 1",team:"Manager-provided monthly presentation",version:0,status:"Manager-provided source · 82 slides",updatedAt:new Date().toISOString()},{name:"EHEQ Monthly Review_2026_27May_S2.pptx",session:"Session 2",team:"Manager-provided monthly presentation",version:0,status:"Manager-provided source · 88 slides",updatedAt:new Date().toISOString()}];
  review.importedFolder="monthly-output/May-2026";
  review.history=[{at:new Date().toISOString(),change:"Loaded manager-provided May 2026 presentation"}];
  importedSlides=Array.from({length:82},(_,i)=>({name:`slide-${String(i+1).padStart(3,"0")}.png`,url:`monthly-output/May-2026/slides/slide-${String(i+1).padStart(3,"0")}.png`}));
  save();
  showMessage("Manager-provided May 2026 data loaded.");
}
function switchView(view) { $$(".view").forEach((node)=>node.classList.toggle("is-visible",node.dataset.viewPanel===view)); $$(".nav-link").forEach((node)=>node.classList.toggle("is-active",node.dataset.view===view)); }
function selectedTeams(id) { return $$(id+" input:checked").map((input)=>input.value); }
function bind() {
  $("#monthSelect").addEventListener("change",()=>{review.month=$("#monthSelect").value;loadPeriod()}); $("#yearSelect").addEventListener("change",()=>{review.year=$("#yearSelect").value;loadPeriod()});
  $("#saveButton").addEventListener("click",save); $("#managerDataButton").addEventListener("click",loadManagerData); $("#newReviewButton").addEventListener("click",()=>{clearImportedSlides();review=createReview($("#monthSelect").value,$("#yearSelect").value);render();showMessage("Started a new review. Upload Session 1 and Session 2 PowerPoint files separately.")});
  $("#session1Form").addEventListener("submit",(e)=>{e.preventDefault();review.session1={date:$("#session1Date").value,agenda:$("#session1Agenda").value,teams:selectedTeams("#session1Teams")};save()});
  $("#session2Form").addEventListener("submit",(e)=>{e.preventDefault();review.session2={enabled:true,date:$("#session2Date").value,agenda:$("#session2Agenda").value,teams:selectedTeams("#session2Teams")};save()});
  $("#teamForm").addEventListener("submit",(e)=>{e.preventDefault();const name=$("#teamName").value.trim();if(!name||review.teams.includes(name))return showMessage("Enter a unique team name.",true);review.teams.push(name);$("#teamName").value="";save()});
  document.addEventListener("click",(e)=>{if(e.target.dataset.view||e.target.dataset.viewAction)switchView(e.target.dataset.view||e.target.dataset.viewAction);if(e.target.dataset.removeTeam){const team=e.target.dataset.removeTeam;review.teams=review.teams.filter((t)=>t!==team);review.session1.teams=review.session1.teams.filter((t)=>t!==team);review.session2.teams=review.session2.teams.filter((t)=>t!==team);save()}if(e.target.dataset.removeFile!==undefined)removeFile(Number(e.target.dataset.removeFile));});
  $("#pptImportButton").addEventListener("click",()=>$("#pptInput").click()); $("#pptInput").addEventListener("change",async(e)=>{const file=e.target.files[0];if(!file)return;const session=$("#uploadSession").value;const duplicate=review.files.find((item)=>(item.name||item.fileName)===file.name&&item.session===session);if(duplicate){duplicate.version=file.lastModified;duplicate.updatedAt=new Date().toISOString();duplicate.status=location.protocol==="file:"?"Recorded locally":"Uploaded";showMessage(`${session} presentation updated.`)}else{try{const serverRecord=await uploadToServer(file,session);review.files.push(serverRecord||{name:file.name,session,team:"",version:file.lastModified,updatedAt:new Date().toISOString(),status:"Recorded locally"});showMessage(serverRecord?`${session} PowerPoint uploaded to the shared service.`:`${session} PowerPoint recorded locally.`)}catch(error){showMessage(error.message,true);return}}e.target.value="";save()});
  $("#pptFiles").addEventListener("change",(e)=>{const sessionIndex=e.target.dataset.fileSession;const teamIndex=e.target.dataset.fileTeam;const index=sessionIndex ?? teamIndex;if(index===undefined)return;if(sessionIndex!==undefined)review.files[Number(index)].session=e.target.value;if(teamIndex!==undefined)review.files[Number(index)].team=e.target.value;review.files[Number(index)].updatedAt=new Date().toISOString();save()});
  $("#openPresentationButton").addEventListener("click",openFullscreen); $("#closeFullscreen").addEventListener("click",closeFullscreen); $("#prevSlide").addEventListener("click",()=>showSlide(currentSlide-1)); $("#nextSlide").addEventListener("click",()=>showSlide(currentSlide+1));
  document.addEventListener("keydown",(e)=>{if(!$("#fullscreen").classList.contains("is-open"))return;if(e.key==="Escape")closeFullscreen();if(e.key==="ArrowLeft")showSlide(currentSlide-1);if(e.key==="ArrowRight")showSlide(currentSlide+1)});
}
function importFolder(e) {
  const files=[...e.target.files]; const images=files.filter((f)=>/slides[\\/].*\.png$/i.test(f.webkitRelativePath)||/slides[\\/].*\.jpe?g$/i.test(f.webkitRelativePath)).sort((a,b)=>a.name.localeCompare(b.name,undefined,{numeric:true}));
  importedSlides.forEach((s)=>URL.revokeObjectURL(s.url)); importedSlides=images.map((file)=>({name:file.name,url:URL.createObjectURL(file)})); review.importedFolder=files[0]?.webkitRelativePath.split(/[\\/]/)[0]||null; review.history.unshift({at:new Date().toISOString(),change:`Imported ${importedSlides.length} generated slides`}); save(); showMessage(`${importedSlides.length} generated slides imported.`,false);
}
function openFullscreen(){if(!importedSlides.length)return showMessage("Import a generated presentation folder first.",true);$("#fullscreen").classList.add("is-open");$("#fullscreen").setAttribute("aria-hidden","false");showSlide(0)}
function closeFullscreen(){$("#fullscreen").classList.remove("is-open");$("#fullscreen").setAttribute("aria-hidden","true")}
function showSlide(index){if(!importedSlides.length)return;currentSlide=Math.max(0,Math.min(index,importedSlides.length-1));$("#slideCanvas").innerHTML=`<img src="${importedSlides[currentSlide].url}" alt="Slide ${currentSlide+1}">`;$("#slideCounter").textContent=`${currentSlide+1} / ${importedSlides.length}`;$("#slideNav").innerHTML=importedSlides.map((slide,i)=>`<button class="${i===currentSlide?"is-active":""}" data-index="${i}">Slide ${i+1}<small>${escapeHtml(slide.name)}</small></button>`).join("");$$("[data-index]").forEach((button)=>button.addEventListener("click",()=>showSlide(Number(button.dataset.index))))}
fillSelectors(); bind(); loadPeriod();
