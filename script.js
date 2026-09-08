let rawRows=[], charts={};

const $=id=>document.getElementById(id);
const num=v=>{
  if(v===null||v===undefined||v==="") return 0;
  if(typeof v==="number") return isFinite(v)?v:0;
  const s=String(v).replace(/,/g,"").replace(/%/g,"").trim();
  const n=parseFloat(s);
  return isFinite(n)?n:0;
};
const pct=v=>`${num(v).toFixed(1)}%`;
const sum=(arr,key)=>arr.reduce((a,r)=>a+num(r[key]),0);
const avg=(arr,key)=>{
  const vals=arr.map(r=>num(r[key])).filter(v=>isFinite(v));
  return vals.length?vals.reduce((a,b)=>a+b,0)/vals.length:0;
};
const weightedRate=(rows,numerator,denominator)=>{
  const d=sum(rows,denominator);
  return d?sum(rows,numerator)/d*100:0;
};
function cleanHeader(s){return String(s||"").trim().toLowerCase().replace(/\s+/g," ").replace(/[%]/g,"");}
function canonical(row){
  const out={};
  Object.keys(row).forEach(k=>out[cleanHeader(k)]=row[k]);
  return out;
}
function get(row, names){
  for(const n of names){ if(row[n]!==undefined) return row[n]; }
  return "";
}
function normalizeRows(rows){
  let currentMonth="";
  return rows.map(r=>{
    const x=canonical(r);
    let month=String(get(x,["month"])||"").trim();
    if(month) currentMonth=month;
    const location=String(get(x,["location"])||"").trim();
    const totalBatch=String(get(x,["total batch conducted","batch","total batches"])||"").trim();
    const marker=(month+" "+location+" "+totalBatch).toLowerCase();
    const isTotal=/total|grand total|overall/.test(marker);
    return {
      month:currentMonth||"Unknown",
      location:location||"Unknown",
      batch:num(get(x,["total batch conducted","batch","total batches"])),
      inflow:num(get(x,["total inflow","inflow"])),
      outflow:num(get(x,["total outflow","outflow"])),
      hr:num(get(x,["hr attrition","hr attrition "])),
      training:num(get(x,["training attrition"])),
      throughput:num(get(x,["throughput"])),
      joined:num(get(x,["total joined","joined"])),
      joiningThroughput:num(get(x,["joining throughput"])),
      isTotal
    };
  }).filter(r=>r.location||r.month);
}
function parseFile(file){
  $("fileStatus").textContent=`Reading ${file.name}…`;
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const wb=XLSX.read(new Uint8Array(e.target.result),{type:"array"});
      const sheet=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(sheet,{defval:""});
      rawRows=normalizeRows(rows);
      populateFilters();
      render();
      $("fileStatus").textContent=`✓ ${file.name} loaded`;
      $("updatedAt").textContent=`Last upload: ${new Date().toLocaleString()}`;
    }catch(err){
      console.error(err);
      $("fileStatus").textContent="Could not read this workbook.";
      alert("The Excel file could not be read. Please check the workbook format.");
    }
  };
  reader.readAsArrayBuffer(file);
}
function populateFilters(){
  const months=[...new Set(rawRows.map(r=>r.month).filter(Boolean))];
  const locations=[...new Set(rawRows.map(r=>r.location).filter(Boolean).filter(x=>!/^total|grand total|overall/i.test(x)))];
  $("monthFilter").innerHTML='<option value="ALL">All months</option>'+months.map(x=>`<option>${escapeHtml(x)}</option>`).join("");
  $("locationFilter").innerHTML='<option value="ALL">All locations</option>'+locations.map(x=>`<option>${escapeHtml(x)}</option>`).join("");
}
function escapeHtml(s){return String(s).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[c]));}
function filteredRows(){
  const m=$("monthFilter").value,l=$("locationFilter").value;
  return rawRows.filter(r=>(m==="ALL"||r.month===m)&&(l==="ALL"||r.location===l)&&!r.isTotal);
}
function monthlyRows(){
  const groups={};
  filteredRows().forEach(r=>{
    const k=r.month;
    if(!groups[k]) groups[k]={month:k,inflow:0,outflow:0,hr:0,training:0,joined:0};
    groups[k].inflow+=r.inflow; groups[k].outflow+=r.outflow; groups[k].hr+=r.hr;
    groups[k].training+=r.training; groups[k].joined+=r.joined;
  });
  return Object.values(groups);
}
function render(){
  const rows=filteredRows(), months=monthlyRows();
  const batches=rows.reduce((a,r)=>a+r.batch,0);
  const inflow=sum(rows,"inflow"), outflow=sum(rows,"outflow"), joined=sum(rows,"joined");
  const hr=sum(rows,"hr"), training=sum(rows,"training");
  const throughput=weightedRate(rows,"joined","inflow");
  const joining=weightedRate(rows,"joined","inflow"); // source summary's joining throughput is not assumed to have a separate denominator
  $("kpiBatches").textContent=batches?format(batches):"0";
  $("kpiInflow").textContent=format(inflow);
  $("kpiOutflow").textContent=format(outflow);
  $("kpiJoined").textContent=format(joined);
  $("kpiHr").textContent=format(hr);
  $("kpiTraining").textContent=format(training);
  $("kpiThroughput").textContent=pct(throughput);
  $("kpiJoiningThroughput").textContent=pct(joining);

  const m=$("monthFilter").value,l=$("locationFilter").value;
  $("viewTitle").textContent=(m==="ALL"?"All months":m)+(l==="ALL"?"":" • "+l);
  const best=locationGroups(rows).sort((a,b)=>b.rate-a.rate)[0];
  $("insightText").textContent=rows.length
    ? `${rows.length} location-level records are in view. ${best?best.location+" currently has the highest calculated throughput at "+pct(best.rate)+".":""}`
    : "No data matches the current filters.";

  drawThroughput(months);
  drawLocation(locationGroups(rows));
  drawMovement(months);
  drawAttrition(months);
  drawTable(locationGroups(rows));
}
function format(n){return Math.round(n).toLocaleString("en-IN");}
function locationGroups(rows){
  const groups={};
  rows.forEach(r=>{
    if(!groups[r.location]) groups[r.location]={location:r.location,inflow:0,outflow:0,hr:0,training:0,joined:0};
    const g=groups[r.location];
    g.inflow+=r.inflow;g.outflow+=r.outflow;g.hr+=r.hr;g.training+=r.training;g.joined+=r.joined;
  });
  return Object.values(groups).map(g=>({...g,rate:g.inflow?g.joined/g.inflow*100:0,joinRate:g.inflow?g.joined/g.inflow*100:0}));
}
function destroy(name){if(charts[name]) charts[name].destroy();}
function commonScales(){return {responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{usePointStyle:true,boxWidth:8,font:{size:11}}}},scales:{x:{grid:{display:false},ticks:{font:{size:10}}},y:{beginAtZero:true,ticks:{font:{size:10}}}}}}
function drawThroughput(rows){
  destroy("throughput");
  charts.throughput=new Chart($("throughputChart"),{type:"line",data:{labels:rows.map(r=>r.month),datasets:[
    {label:"Throughput",data:rows.map(r=>r.inflow?r.joined/r.inflow*100:0),tension:.35,borderWidth:3,pointRadius:3},
    {label:"Joining Throughput",data:rows.map(r=>r.inflow?r.joined/r.inflow*100:0),tension:.35,borderWidth:2,pointRadius:3,borderDash:[6,5]}
  ]},options:{...commonScales(),scales:{...commonScales().scales,y:{beginAtZero:true,max:100,ticks:{callback:v=>v+"%"}}}}});
}
function drawLocation(rows){
  destroy("location");
  const sorted=[...rows].sort((a,b)=>b.rate-a.rate);
  charts.location=new Chart($("locationChart"),{type:"bar",data:{labels:sorted.map(r=>r.location),datasets:[{label:"Throughput",data:sorted.map(r=>r.rate),borderRadius:6}]},options:{indexAxis:"y",responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{beginAtZero:true,max:100,ticks:{callback:v=>v+"%"},grid:{display:false}},y:{grid:{display:false},ticks:{font:{size:10}}}}}});
}
function drawMovement(rows){
  destroy("movement");
  charts.movement=new Chart($("movementChart"),{type:"bar",data:{labels:rows.map(r=>r.month),datasets:[
    {label:"Inflow",data:rows.map(r=>r.inflow),borderRadius:5},
    {label:"Outflow",data:rows.map(r=>r.outflow),borderRadius:5}
  ]},options:commonScales()});
}
function drawAttrition(rows){
  destroy("attrition");
  charts.attrition=new Chart($("attritionChart"),{type:"line",data:{labels:rows.map(r=>r.month),datasets:[
    {label:"HR Attrition",data:rows.map(r=>r.hr),tension:.3,borderWidth:3,pointRadius:3},
    {label:"Training Attrition",data:rows.map(r=>r.training),tension:.3,borderWidth:3,pointRadius:3}
  ]},options:commonScales()});
}
function drawTable(rows){
  const body=$("summaryTable");
  if(!rows.length){body.innerHTML='<tr><td colspan="8" class="empty">No records match the selected filters.</td></tr>';return;}
  body.innerHTML=rows.sort((a,b)=>b.rate-a.rate).map(r=>`<tr>
    <td>${escapeHtml(r.location)}</td><td>${format(r.inflow)}</td><td>${format(r.outflow)}</td>
    <td>${format(r.hr)}</td><td>${format(r.training)}</td><td>${format(r.joined)}</td>
    <td>${pct(r.rate)}</td><td>${pct(r.joinRate)}</td>
  </tr>`).join("");
}
$("excelFile").addEventListener("change",e=>{if(e.target.files[0])parseFile(e.target.files[0]);});
$("monthFilter").addEventListener("change",render);
$("locationFilter").addEventListener("change",render);
