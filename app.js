const TEAM = ["Sher", "Leila", "Jen", "Shona"];
const STATUSES = ["Not started", "In preparation", "Waiting for client", "Ready for review", "Review notes", "Approved", "Filed"];
const storageKey = "tax-desk-returns-v1";
const sample = [
  {id:"t1-1",client:"Alex Martin",year:2025,preparer:"Sher",reviewer:"Leila",status:"In preparation",deadline:"2026-04-30",notes:"T4 received",updated:"2026-10-06"},
  {id:"t1-2",client:"Samira Rahman",year:2025,preparer:"Jen",reviewer:"Shona",status:"Ready for review",deadline:"2026-04-30",notes:"Review foreign income slip",updated:"2026-10-05"},
  {id:"t1-3",client:"David Wong",year:2025,preparer:"Sher",reviewer:"Leila",status:"Waiting for client",deadline:"2026-04-30",notes:"Awaiting RRSP receipt",updated:"2026-10-04"}
];
let returns = JSON.parse(localStorage.getItem(storageKey) || "null") || sample;
let currentView = "all";
const $ = (s) => document.querySelector(s);
function save(){localStorage.setItem(storageKey,JSON.stringify(returns));}
function slug(status){return status.toLowerCase().replaceAll(" ","-");}
function showToast(message){const toast=$("#toast");toast.textContent=message;toast.classList.add("show");setTimeout(()=>toast.classList.remove("show"),2800);}
function render(){
  const q=$("#search").value.toLowerCase().trim();
  const selected=returns.filter(r=>{
    const matches={all:true,prepare:["Not started","In preparation","Review notes"].includes(r.status),review:r.status==="Ready for review",waiting:r.status==="Waiting for client"}[currentView];
    return matches && [r.client,r.preparer,r.reviewer,r.status].join(" ").toLowerCase().includes(q);
  });
  $("#return-list").innerHTML=selected.map(r=>`<tr><td><span class="client">${escapeHtml(r.client)}</span><span class="small">${r.year} T1</span></td><td>${escapeHtml(r.preparer)}</td><td>${escapeHtml(r.reviewer)}</td><td><span class="pill ${slug(r.status)}">${r.status}</span></td><td>${formatDate(r.deadline)}</td><td>${formatDate(r.updated)}</td><td><button class="row-action" data-id="${r.id}">Open</button></td></tr>`).join("");
  $("#empty-state").hidden=selected.length>0;
  const counts=[ ["To prepare",returns.filter(r=>["Not started","In preparation","Review notes"].includes(r.status)).length], ["Ready for review",returns.filter(r=>r.status==="Ready for review").length], ["Waiting for client",returns.filter(r=>r.status==="Waiting for client").length], ["Filed",returns.filter(r=>r.status==="Filed").length] ];
  $("#metrics").innerHTML=counts.map(([label,value])=>`<div class="metric"><div class="metric-value">${value}</div><div class="metric-label">${label}</div></div>`).join("");
}
function formatDate(value){return value?new Intl.DateTimeFormat("en-CA",{month:"short",day:"numeric",year:"numeric"}).format(new Date(`${value}T12:00:00`)):"—";}
function escapeHtml(v){return String(v||"").replace(/[&<>'"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[c]));}
function populateSelect(id,values){$(id).innerHTML=values.map(v=>`<option>${v}</option>`).join("");}
function openReturn(record){
  populateSelect("#preparer",TEAM);populateSelect("#reviewer",TEAM);populateSelect("#status",STATUSES);
  $("#dialog-title").textContent=record?"T1 return details":"Add a T1 return";
  $("#record-id").value=record?.id||"";$("#client-name").value=record?.client||"";$("#tax-year").value=record?.year||2025;$("#preparer").value=record?.preparer||TEAM[0];$("#reviewer").value=record?.reviewer||TEAM[1];$("#status").value=record?.status||"Not started";$("#deadline").value=record?.deadline||"2026-04-30";$("#notes").value=record?.notes||"";
  $("#return-dialog").showModal();
}
$("#add-return").addEventListener("click",()=>openReturn());
$("#return-form").addEventListener("submit",(e)=>{e.preventDefault();const id=$("#record-id").value;const record={id:id||crypto.randomUUID(),client:$("#client-name").value.trim(),year:+$("#tax-year").value,preparer:$("#preparer").value,reviewer:$("#reviewer").value,status:$("#status").value,deadline:$("#deadline").value,notes:$("#notes").value.trim(),updated:new Date().toISOString().slice(0,10)};if(id)returns=returns.map(r=>r.id===id?record:r);else returns=[record,...returns];save();$("#return-dialog").close();render();showToast(id?"T1 return updated":"T1 return added");});
$("#return-list").addEventListener("click",e=>{const id=e.target.dataset.id;if(id)openReturn(returns.find(r=>r.id===id));});
document.querySelectorAll(".tab").forEach(tab=>tab.addEventListener("click",()=>{document.querySelector(".tab.active").classList.remove("active");tab.classList.add("active");currentView=tab.dataset.view;render();}));
$("#search").addEventListener("input",render);
$("#file-upload").addEventListener("change",async e=>{const file=e.target.files[0];if(!file)return;if(!window.XLSX){showToast("Excel importer is not available. Please try again online.");return;}try{const workbook=XLSX.read(await file.arrayBuffer(),{type:"array"});const rows=XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]],{defval:""});const added=rows.map(row=>({id:crypto.randomUUID(),client:String(row["Client name"]||row["Client"]||"").trim(),year:+(row["Tax year"]||row["Year"]||2025),preparer:String(row["Preparer"]||TEAM[0]),reviewer:String(row["Reviewer"]||TEAM[1]),status:STATUSES.includes(row["Current status"]||row["Status"])?(row["Current status"]||row["Status"]):"Not started",deadline:toIso(row["Due date"]||row["Deadline"])||"2026-04-30",notes:String(row["Notes"]||""),updated:new Date().toISOString().slice(0,10)})).filter(r=>r.client);if(!added.length)throw new Error("No client names found");returns=[...added,...returns];save();render();showToast(`${added.length} T1 return${added.length===1?"":"s"} imported`);}catch(error){showToast("Could not import. Use the column headings in the template.");}e.target.value="";});
function toIso(value){if(value instanceof Date)return value.toISOString().slice(0,10);if(typeof value==="number"&&window.XLSX)return XLSX.SSF.format("yyyy-mm-dd",value);const parsed=new Date(value);return Number.isNaN(parsed.getTime())?"":parsed.toISOString().slice(0,10);}
render();
