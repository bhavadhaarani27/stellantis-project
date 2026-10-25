const months=["January","February","March","April","May","June","July","August","September","October","November","December"];
const now=new Date();
const $=(selector)=>document.querySelector(selector);
const state={files:{1:null,2:null},results:{1:null,2:null}};

function fillSelectors(){
  $("#monthSelect").innerHTML=months.map((month)=>`<option>${month}</option>`).join("");
  $("#yearSelect").innerHTML=Array.from({length:11},(_,index)=>now.getFullYear()-5+index).map((year)=>`<option>${year}</option>`).join("");
  $("#monthSelect").value=months[now.getMonth()];$("#yearSelect").value=String(now.getFullYear());
}
function escapeHtml(value){return String(value||"").replace(/[&<>"']/g,(char)=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[char]));}
function setMessage(text,error=false){$("#message").textContent=text;$("#message").className=error?"message error":"message";}
function renderFile(session){
  const file=state.files[session];
  $(`#session${session}File`).textContent=file?file.name:"No Session "+session+" file selected.";
}
function renderResult(session){
  const result=state.results[session];
  const node=$(`#session${session}Result`);
  if(!result){node.className="result dummy";node.textContent=`Session ${session} presentation will appear here.`;return;}
  if(result.dummy){node.className="result dummy";node.textContent=`Session ${session} is a placeholder because no PowerPoint was supplied.`;return;}
  node.className="result ready";
  node.innerHTML=`Session ${session} presentation is ready.<br><a href="${result.url}">Open Session ${session} presentation</a>`;
}
function clearSelection(){
  state.files={1:null,2:null};state.results={1:null,2:null};
  $("#monthSelect").value=months[now.getMonth()];$("#yearSelect").value=String(now.getFullYear());
  $("#session1Input").value="";$("#session2Input").value="";$("#session1Date").value="";$("#session1Agenda").value="";$("#session2Agenda").value="";$("#session2Date").value="";$("#session1Teams").innerHTML="";$("#session2Teams").innerHTML="";
  renderFile(1);renderFile(2);renderResult(1);renderResult(2);setMessage("Selection cleared. Choose a month and upload the required session files.");
}
function selectedTeams(session){return [...document.querySelectorAll(`#session${session}Teams input:checked`)].map((box)=>box.value);}
function renderTeams(session,teams){
  const node=$(`#session${session}Teams`);
  if(!teams.length){node.innerHTML="";return;}
  node.innerHTML=`<h3>Choose teams for Session ${session}</h3><div class="tools"><button type="button" data-all="1">All</button><button type="button" data-all="0">None</button></div>`+teams.map((team)=>`<label class="${team.slides?"":"off"}"><span class="mv"><button type="button" data-mv="-1" title="Move up">&#9650;</button><button type="button" data-mv="1" title="Move down">&#9660;</button></span><input type="checkbox" ${team.slides?"checked":"disabled"} value="${escapeHtml(team.key)}"> ${escapeHtml(team.label)} <small>${team.slides?team.slides+" slides":"not in this file"}</small></label>`).join("");
  node.querySelectorAll("[data-mv]").forEach((button)=>button.addEventListener("click",(event)=>{event.preventDefault();const row=button.closest("label");const other=button.dataset.mv==="-1"?row.previousElementSibling:row.nextElementSibling;if(!other||other.tagName!=="LABEL")return;button.dataset.mv==="-1"?other.before(row):other.after(row);}));
  node.querySelectorAll("[data-all]").forEach((button)=>button.addEventListener("click",()=>node.querySelectorAll("input:not(:disabled)").forEach((box)=>{box.checked=button.dataset.all==="1";})));
}
async function selectFile(session,event){
  state.files[session]=event.target.files[0]||null;state.results[session]=null;renderFile(session);renderResult(session);setMessage("");
  $(`#session${session}Teams`).innerHTML="";
  if(!state.files[session])return;
  const form=new FormData();form.append("file",state.files[session]);
  try{
    const response=await fetch("/api/simple-sections",{method:"POST",body:form});
    const data=await response.json();
    if(!response.ok)throw new Error(data.error);
    renderTeams(session,data.teams);
  }catch(error){setMessage(error.message||"Could not read the teams from this file.",true);}
}
async function generateSession(session){
  const file=state.files[session];
  if(!file)return {dummy:true};
  const form=new FormData();
  form.append("file",file);
  form.append("month",$("#monthSelect").value);
  form.append("year",$("#yearSelect").value);
  form.append("session",`Session ${session}`);
  form.append("meetingDate",$(`#session${session}Date`).value);form.append("agenda",$(`#session${session}Agenda`).value);
  const teams=document.querySelectorAll(`#session${session}Teams input`).length?selectedTeams(session):[];
  if(document.querySelectorAll(`#session${session}Teams input`).length&&!teams.length)throw new Error(`Select at least one team for Session ${session}.`);
  form.append("teams",JSON.stringify(teams));
  const response=await fetch("/api/simple-generate",{method:"POST",body:form});
  const result=await response.json();
  if(!response.ok)throw new Error(result.error||`Session ${session} could not be generated.`);
  return {url:result.url};
}
async function generate(){
  const month=$("#monthSelect").value,year=$("#yearSelect").value;
  if(!state.files[1]&&!state.files[2])return setMessage("Select at least one Session 1 or Session 2 PowerPoint file.",true);
  $("#generateButton").disabled=true;setMessage("Generating the selected presentations...");
  try{
    state.results[1]=await generateSession(1);
    state.results[2]=await generateSession(2);
    renderResult(1);renderResult(2);
    loadReviewList();
    setMessage(`${month} ${year} monthly review is ready. Each session uses only its own uploaded file.`);
  }catch(error){setMessage(error.message,true);}
  finally{$("#generateButton").disabled=false;}
}
fillSelectors();
$("#session1Input").addEventListener("change",(event)=>selectFile(1,event));
$("#session2Input").addEventListener("change",(event)=>selectFile(2,event));
$("#clearButton").addEventListener("click",clearSelection);
$("#generateButton").addEventListener("click",generate);
renderFile(1);renderFile(2);renderResult(1);renderResult(2);


/* ---- Late teams / SharePoint links for an existing review ---- */
const manage={current:null};
const manageMessage=(text,error=false)=>{$("#manageMessage").textContent=text;$("#manageMessage").className=error?"message error":"message";};
async function loadReviewList(){
  const data=await (await fetch("/api/simple-list")).json();
  $("#reviewSelect").innerHTML='<option value="">Select a generated review</option>'+data.items.map((item)=>`<option value="${escapeHtml(item.period)}|${escapeHtml(item.session)}">${escapeHtml(item.month)} ${escapeHtml(item.year)} - ${escapeHtml(item.session)}</option>`).join("");
}
async function openReview(value){
  if(!value){manage.current=null;$("#teamTable").innerHTML="";return;}
  const [period,session]=value.split("|");
  const response=await fetch(`/api/simple-session?period=${encodeURIComponent(period)}&session=${encodeURIComponent(session)}`);
  const data=await response.json();
  if(!response.ok)return manageMessage(data.error,true);
  manage.current={period,session,...data};
  renderTeamTable();
}
async function updateTeam(fields,file){
  const [month,year]=[manage.current.month,manage.current.year];
  const form=new FormData();
  form.append("month",month);form.append("year",year);form.append("session",manage.current.session);
  Object.entries(fields).forEach(([name,value])=>form.append(name,value));
  if(file)form.append("file",file);
  manageMessage("Updating the web page...");
  const response=await fetch("/api/simple-update",{method:"POST",body:form});
  const data=await response.json();
  if(!response.ok)return manageMessage(data.error||"Update failed.",true);
  manage.current={period:manage.current.period,session:manage.current.session,...data};
  renderTeamTable();manageMessage("Updated. The web page now includes this change.");
}
const statusLabel={pptx:"Slides added",link:"Link provided",pending:"Awaiting slides",missing:"Not in review"};
function renderTeamTable(){
  const review=manage.current;
  $("#manageDate").value=review.meetingDate||"";$("#manageAgenda").value=review.agenda||"";
  const present=review.teams.filter((team)=>team.source!=="missing");
  $("#teamTable").innerHTML=`<p><a href="${review.url}" target="_blank" rel="noopener">Open ${escapeHtml(review.month)} ${escapeHtml(review.year)} ${escapeHtml(review.session)} page</a></p><table class="teams-table"><thead><tr><th>Team</th><th>Status</th><th>Late PowerPoint</th><th>SharePoint link</th><th></th></tr></thead><tbody>`+review.teams.map((team)=>{
    const index=present.indexOf(team);
    return `<tr data-key="${escapeHtml(team.key)}"><td><b>${escapeHtml(team.label)}</b>${team.source!=="missing"?` <span class="mv"><button data-act="up" ${index<=0?"disabled":""}>&#9650;</button><button data-act="down" ${index===present.length-1?"disabled":""}>&#9660;</button></span>`:""}</td><td><span class="badge ${team.source}">${statusLabel[team.source]}${team.source==="pptx"?` (${team.slides})`:""}</span></td><td><input type="file" accept=".pptx" data-act="file">${syncOptions()?`<br><select data-act="syncSel"><option value="">From synced folder...</option>${syncOptions()}</select> <button data-act="syncUse">Use</button>`:""}</td><td><input type="url" placeholder="https://...sharepoint.com/..." value="${escapeHtml(team.link)}" data-act="linkText"> <button data-act="link">Save</button></td><td>${team.source==="missing"?`<button data-act="pending">Mark pending</button>`:`<button data-act="remove">Remove</button>`}</td></tr>`;
  }).join("")+"</tbody></table>";
}
$("#teamTable").addEventListener("click",(event)=>{
  const button=event.target.closest("button[data-act]");if(!button||!manage.current)return;
  const row=button.closest("tr");const key=row.dataset.key;const act=button.dataset.act;
  if(act==="link")return updateTeam({action:"setLink",key,link:row.querySelector("[data-act=linkText]").value});
  if(act==="syncUse"){const syncPath=row.querySelector("[data-act=syncSel]").value;return syncPath?updateTeam({action:"setPptx",key,syncPath}):manageMessage("Choose a file from the synced folder first.",true);}
  if(act==="pending")return updateTeam({action:"setPending",key});
  if(act==="remove"){if(confirm("Remove this team from the review?"))updateTeam({action:"remove",key});return;}
  const order=manage.current.teams.filter((team)=>team.source!=="missing").map((team)=>team.key);
  const at=order.indexOf(key);const to=act==="up"?at-1:at+1;
  [order[at],order[to]]=[order[to],order[at]];
  updateTeam({action:"order",order:JSON.stringify(order)});
});
$("#teamTable").addEventListener("change",(event)=>{
  if(event.target.dataset.act!=="file"||!event.target.files[0])return;
  updateTeam({action:"setPptx",key:event.target.closest("tr").dataset.key},event.target.files[0]);
});
$("#reviewSelect").addEventListener("change",(event)=>openReview(event.target.value));
$("#manageDate").addEventListener("change",(event)=>{if(manage.current)updateTeam({action:"date",date:event.target.value});});
loadReviewList();
/* ---- Synced folder (SharePoint/OneDrive sync) ---- */
let syncFiles=[];
const syncOptions=()=>syncFiles.map((file)=>`<option value="${escapeHtml(file.path)}">${escapeHtml(file.path)}</option>`).join("");
async function loadSync(){
  const data=await (await fetch("/api/sync-files")).json();
  syncFiles=data.files||[];$("#syncFolder").value=data.folder||"";
  if(data.error)manageMessage(data.error,true);
  if(manage.current)renderTeamTable();
}
$("#syncSave").addEventListener("click",async()=>{
  const response=await fetch("/api/sync-folder",{method:"POST",body:JSON.stringify({folder:$("#syncFolder").value})});
  const data=await response.json();
  if(!response.ok)return manageMessage(data.error,true);
  syncFiles=data.files;manageMessage(`Synced folder saved. ${syncFiles.length} PowerPoint file(s) found.`);
  if(manage.current)renderTeamTable();
});
loadSync();

$("#agendaSave").addEventListener("click",()=>{if(manage.current)updateTeam({action:"agenda",agenda:$("#manageAgenda").value});else manageMessage("Select an existing review first.",true);});
