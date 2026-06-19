import { useState, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./supabase";
const STORAGE_KEY = "ipc_project_v2";

const C = {
  blue:"#1F4E79", blueMid:"#2E75B6", blueLight:"#D6E4F0", navy:"#0D2B45",
  white:"#FFFFFF", offwhite:"#F5F7FA", gray:"#6B7280", grayLight:"#E5E7EB",
  done:"#166534", doneBg:"#DCFCE7", progress:"#92400E", progressBg:"#FEF3C7",
  open:"#1E3A5F", openBg:"#DBEAFE", hold:"#7C3AED", holdBg:"#EDE9FE",
  red:"#991B1B", redBg:"#FEE2E2",
};

const STATUS_CONFIG = {
  "Done":        { color:C.done,     bg:C.doneBg,      label:"مكتمل" },
  "In Progress": { color:C.progress, bg:C.progressBg,  label:"جاري"  },
  "Open":        { color:C.open,     bg:C.openBg,      label:"مفتوح" },
  "On Hold":     { color:C.hold,     bg:C.holdBg,      label:"معلق"  },
};

// ── البيانات بتاعتك كما هي بالضبط ────────────────────────────────────────────
const DEFAULT_DATA = [
  { id:"P1", title:"Production", titleAr:"الإنتاج", weight:30, owner:"Eng. Ehab Eweis", tasks:[
    { id:"P1-1", title:"توقيع عقد المصنع",           weight:25, status:"Open",        owner:"Eng. Ahmed Elmohamady", due:"11-06-2026" },
    { id:"P1-2", title:"زيارة تحقق للموقع",            weight:10, status:"Open",        owner:"Projects + Technical",  due:"بعد العقد"  },
    { id:"P1-3", title:"إعداد المشتريات (BOM Final)", weight:15, status:"Open",        owner:"Eng. Ehab Eweis",       due:"TBD"        },
    { id:"P1-4", title:"شراء وتوريد المعدات",          weight:20, status:"Open",        owner:"Procurement",           due:"TBD"        },
    { id:"P1-5", title:"تركيب المعدات",                weight:15, status:"Open",        owner:"Projects",              due:"TBD"        },
    { id:"P1-6", title:"شراء المواد الخام",             weight:8,  status:"Open",        owner:"Procurement",           due:"TBD"        },
    { id:"P1-7", title:"تجربة تشغيلية أولى",           weight:7,  status:"Open",        owner:"Technical + Projects",  due:"TBD"        },
  ]},
  { id:"P2", title:"Product", titleAr:"المنتج", weight:20, owner:"Dr. Hanan Sayed", tasks:[
    { id:"P2-1", title:"اعتماد اللون النهائي",          weight:20, status:"On Hold", owner:"Dr. Mohamed Eid",  due:"TBD" },
    { id:"P2-2", title:"تحديد موضع CAS Number",         weight:15, status:"On Hold", owner:"Technical / R&D",  due:"TBD" },
    { id:"P2-3", title:"إعداد Toxicity Assessment",    weight:15, status:"On Hold", owner:"Technical / R&D",  due:"TBD" },
    { id:"P2-4", title:"إصدار Datasheet نهائي",         weight:25, status:"On Hold", owner:"Dr. Hanan Sayed",  due:"TBD" },
    { id:"P2-5", title:"إصدار MSDS نهائي",              weight:25, status:"On Hold", owner:"Dr. Hanan Sayed",  due:"TBD" },
  ]},
  { id:"P3", title:"Brand & Marketing", titleAr:"الهوية والتسويق", weight:15, owner:"Eng. Ahmed Elmohamady", tasks:[
    { id:"P3-1", title:"حسم اسم المنتج النهائي",  weight:25, status:"Open", owner:"Eng. Tamer Ibrahim", due:"TBD" },
    { id:"P3-2", title:"تسجيل الدومين",            weight:15, status:"Open", owner:"Eng. Tamer Ibrahim", due:"TBD" },
    { id:"P3-3", title:"Packaging & Label Brief", weight:25, status:"Open", owner:"PM – Eng. Ahmed A.", due:"TBD" },
    { id:"P3-4", title:"تصميم الهوية البصرية",      weight:20, status:"Open", owner:"Marketing",          due:"TBD" },
    { id:"P3-5", title:"Campaign Calendar",        weight:15, status:"Open", owner:"BD / Marketing",     due:"TBD" },
  ]},
  { id:"P4", title:"Sales & Pipeline", titleAr:"المبيعات والفرص", weight:15, owner:"Eng. Ahmed Elmohamady", tasks:[
    { id:"P4-1", title:"إعداد Pipeline Sheet رسمي", weight:25, status:"Open",    owner:"Eng. Ahmed Elmohamady", due:"11-06-2026" },
    { id:"P4-2", title:"تجهيز 10 عينات IPC",         weight:20, status:"Open",    owner:"Dr. Hanan + PM",        due:"TBD"        },
    { id:"P4-3", title:"زيارات عملاء وجمع عينات",    weight:20, status:"Open",    owner:"Dr. Hanan + PM",        due:"TBD"        },
    { id:"P4-4", title:"Pricing Matrix v0",          weight:20, status:"On Hold", owner:"BD + Finance",          due:"TBD"        },
    { id:"P4-5", title:"Distribution Channel Map",  weight:15, status:"Open",    owner:"Eng. Ahmed Elmohamady", due:"TBD"        },
  ]},
  { id:"P5", title:"Legal & Corporate", titleAr:"القانوني والمؤسسي", weight:10, owner:"Eng. Tamer Ibrahim", tasks:[
    { id:"P5-1", title:"تأسيس الكيان القانوني",     weight:40, status:"In Progress", owner:"Eng. Tamer Ibrahim",  due:"TBD" },
    { id:"P5-2", title:"التراخيص الصناعية والبيئية",weight:35, status:"Open",        owner:"Legal / Management",  due:"TBD" },
    { id:"P5-3", title:"عقود الموردين والشركاء",    weight:25, status:"Open",        owner:"Management",          due:"TBD" },
  ]},
  { id:"P6", title:"R&D", titleAr:"البحث والتطوير", weight:10, owner:"Dr. Mohamed Eid", tasks:[
    { id:"P6-1", title:"R&D Brief – Sewage Module",      weight:30, status:"Open", owner:"Dr. Mohamed Eid", due:"TBD" },
    { id:"P6-2", title:"R&D Brief – IWWT Module",        weight:30, status:"Open", owner:"Dr. Mohamed Eid", due:"TBD" },
    { id:"P6-3", title:"تجارب تحسين اللون",               weight:20, status:"Open", owner:"R&D / Technical", due:"TBD" },
    { id:"P6-4", title:"تطوير منتجات كيماوية مستقبلية",   weight:20, status:"Open", owner:"R&D",             due:"TBD" },
  ]},
];

// ── Helpers ───────────────────────────────────────────────────────────────────
function taskWeightSum(tasks) { return tasks.reduce((s,t)=>s+Number(t.weight),0); }

function calcProgress(tasks) {
  const total = taskWeightSum(tasks);
  if (!total) return 0;
  const done = tasks.filter(t=>t.status==="Done").reduce((s,t)=>s+Number(t.weight),0);
  return Math.round((done/total)*100);
}

function calcOverall(data) {
  const totalW = data.reduce((s,p)=>s+p.weight,0);
  const contrib = data.reduce((s,p)=>s+(calcProgress(p.tasks)/100)*p.weight,0);
  return totalW ? Math.round((contrib/totalW)*100) : 0;
}

function ProgressBar({ pct, height=8, color=C.blueMid }) {
  return (
    <div style={{background:C.grayLight,borderRadius:99,height,overflow:"hidden",width:"100%"}}>
      <div style={{width:`${pct}%`,height:"100%",borderRadius:99,
        background:pct===100?C.done:pct>0?color:C.grayLight,transition:"width 0.4s ease"}}/>
    </div>
  );
}

function StatusBadge({status}) {
  const cfg = STATUS_CONFIG[status]||STATUS_CONFIG["Open"];
  return <span style={{fontSize:11,fontWeight:700,padding:"2px 8px",borderRadius:99,color:cfg.color,background:cfg.bg,whiteSpace:"nowrap"}}>{cfg.label}</span>;
}

function WeightWarning({ tasks }) {
  const sum = taskWeightSum(tasks);
  if (sum === 100) return null;
  return (
    <span style={{fontSize:10,color:sum>100?C.red:C.progress,background:sum>100?C.redBg:C.progressBg,padding:"1px 7px",borderRadius:99,fontWeight:700}}>
      مجموع الأوزان = {sum}% {sum>100?"(زيادة)":"(ناقص)"}
    </span>
  );
}

// ── Export Excel ──────────────────────────────────────────────────────────────
function exportExcel(data) {
  const overall = calcOverall(data);
  const wb = XLSX.utils.book_new();

  // Sheet 1: ملخص
  const s1 = [
    ["IPC Project – ملخص تقدم المشروع","","","","","",""],
    ["التاريخ:", new Date().toLocaleDateString("ar-EG"),"","التقدم الكلي:",`${overall}%`,"",""],
    ["PM:","Eng. Ahmed Aboelhasan","","","","",""],
    [""],
    ["العنوان الرئيسي","الوزن %","المسؤول","عدد المهام","مكتمل","جاري","مفتوح","معلق","التقدم %"],
  ];
  data.forEach(p=>{
    const prog=calcProgress(p.tasks);
    s1.push([p.titleAr,p.weight,p.owner,p.tasks.length,
      p.tasks.filter(t=>t.status==="Done").length,
      p.tasks.filter(t=>t.status==="In Progress").length,
      p.tasks.filter(t=>t.status==="Open").length,
      p.tasks.filter(t=>t.status==="On Hold").length,
      `${prog}%`]);
  });
  const ws1=XLSX.utils.aoa_to_sheet(s1);
  ws1["!cols"]=[{wch:28},{wch:10},{wch:24},{wch:12},{wch:10},{wch:10},{wch:10},{wch:10},{wch:12}];
  XLSX.utils.book_append_sheet(wb,ws1,"ملخص");

  // Sheet 2: كل المهام
  const s2=[["العنوان الرئيسي","المهمة","المسؤول","الموعد","الوزن %","الحالة"]];
  data.forEach(p=>p.tasks.forEach(t=>{
    s2.push([p.titleAr,t.title,t.owner,t.due,t.weight,STATUS_CONFIG[t.status]?.label||t.status]);
  }));
  const ws2=XLSX.utils.aoa_to_sheet(s2);
  ws2["!cols"]=[{wch:22},{wch:34},{wch:26},{wch:14},{wch:10},{wch:12}];
  XLSX.utils.book_append_sheet(wb,ws2,"كل المهام");

  // Sheet per pillar
  data.forEach(p=>{
    const rows=[
      [`${p.titleAr} – ${p.title}`,"","","",""],
      ["المسؤول:",p.owner,"","وزن العنوان:",`${p.weight}%`],
      ["التقدم:",`${calcProgress(p.tasks)}%`,"","مجموع أوزان المهام:",`${taskWeightSum(p.tasks)}%`],
      [""],
      ["المهمة","المسؤول","الموعد","الوزن %","الحالة"],
    ];
    p.tasks.forEach(t=>rows.push([t.title,t.owner,t.due,t.weight,STATUS_CONFIG[t.status]?.label||t.status]));
    const ws=XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"]=[{wch:34},{wch:26},{wch:14},{wch:10},{wch:12}];
    XLSX.utils.book_append_sheet(wb,ws,p.title.substring(0,31));
  });

  XLSX.writeFile(wb,`IPC_WBS_${new Date().toISOString().slice(0,10)}.xlsx`);
}

function backupData(data){
  const blob = new Blob(
    [JSON.stringify(data, null, 2)],
    { type: "application/json" }
  );

  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = `IPC_Backup_${new Date().toISOString().slice(0,10)}.json`;

  document.body.appendChild(a);
  a.click();

  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function restoreBackup(setData){
  const input = document.createElement("input");
  input.type = "file";
  input.accept = ".json";

  input.onchange = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    try {
      const text = await file.text();
      const importedData = JSON.parse(text);

      if (window.confirm("هل تريد استرجاع النسخة الاحتياطية؟")) {
        setData(importedData);
        alert("تم استرجاع البيانات بنجاح");
      }

    } catch (err) {
      console.error(err);
      alert("الملف غير صالح أو تالف");
    }
  };

  input.click();
}

// ── Export Word (HTML→.doc) ───────────────────────────────────────────────────
async function exportWord(data, setWordStatus) {
  setWordStatus("generating");
  const overall = calcOverall(data);
  const today = new Date().toLocaleDateString("ar-EG");

  try {
    const res = await fetch("https://api.anthropic.com/v1/messages",{
      method:"POST",
      headers:{"Content-Type":"application/json"},
      body:JSON.stringify({
        model:"claude-sonnet-4-6", max_tokens:1000,
        messages:[{role:"user",content:`اكتب ملخصاً تنفيذياً باللغة العربية في 4-6 جمل مناسب لتقرير بورد الإدارة عن هذا المشروع. لا تستخدم markdown أو نجوم أو شرطات:

IPC Project – تقدم المشروع الكلي: ${overall}%
${data.map(p=>`${p.titleAr}: تقدم ${calcProgress(p.tasks)}% (وزن ${p.weight}%) – مكتمل ${p.tasks.filter(t=>t.status==="Done").length} / جاري ${p.tasks.filter(t=>t.status==="In Progress").length} / مفتوح ${p.tasks.filter(t=>t.status==="Open").length} / معلق ${p.tasks.filter(t=>t.status==="On Hold").length}`).join("\n")}`}]
      })
    });
    const d = await res.json();
    const summary = d.content?.[0]?.text || "لا يوجد ملخص متاح.";

    const statusColor = {Done:"#166534","In Progress":"#92400E",Open:"#1E3A5F","On Hold":"#7C3AED"};
    const statusLabel = {Done:"مكتمل","In Progress":"جاري",Open:"مفتوح","On Hold":"معلق"};

    const html=`<!DOCTYPE html>
<html dir="rtl" lang="ar">
<head><meta charset="UTF-8">
<style>
body{font-family:'Sakkal Majalla',Arial;margin:2cm;color:#333;direction:rtl;font-size:12pt}
h1{color:#1F4E79;font-size:20pt;border-bottom:3px solid #1F4E79;padding-bottom:6px;margin-bottom:4px}
h2{color:#1F4E79;font-size:14pt;margin-top:24px;margin-bottom:4px}
.meta{color:#595959;font-size:10pt;margin-bottom:16px}
.overall{background:#D6E4F0;padding:10px 16px;border-radius:6px;font-size:14pt;font-weight:bold;color:#1F4E79;margin-bottom:18px}
.summary{background:#F5F7FA;border-right:4px solid #1F4E79;padding:12px 16px;margin-bottom:20px;line-height:1.9;font-size:11pt}
.prog-bg{background:#E5E7EB;border-radius:99px;height:10px;margin:6px 0 10px}
.prog-fill{background:#2E75B6;border-radius:99px;height:10px}
.pillar-head{display:flex;justify-content:space-between;align-items:center}
.pct{font-size:18pt;font-weight:900;color:#1F4E79}
table{width:100%;border-collapse:collapse;font-size:10pt;margin-bottom:6px}
th{background:#1F4E79;color:white;padding:7px 10px;text-align:right;font-size:10pt}
td{padding:6px 10px;border:1px solid #E5E7EB;vertical-align:top}
tr:nth-child(even) td{background:#F5F7FA}
.pillar-info{font-size:10pt;color:#6B7280;margin-bottom:6px}
.badges{display:flex;gap:8px;flex-wrap:wrap;margin-top:6px;font-size:10pt}
.badge{padding:2px 10px;border-radius:99px;font-weight:bold}
footer{margin-top:30px;font-size:9pt;color:#9CA3AF;text-align:center;border-top:1px solid #E5E7EB;padding-top:10px}
@media print{body{margin:1.5cm}}
</style>
</head>
<body>
<h1>IPC Project — تقرير تنفيذي للبورد</h1>
<div class="meta">تاريخ: ${today} &nbsp;|&nbsp; PM: Eng. Ahmed Aboelhasan &nbsp;|&nbsp; داخلي / Internal</div>
<div class="overall">📊 التقدم الكلي للمشروع: ${overall}%</div>
<h2>الملخص التنفيذي</h2>
<div class="summary">${summary.replace(/\n/g,"<br>")}</div>
${data.map(p=>{
  const prog=calcProgress(p.tasks);
  const done=p.tasks.filter(t=>t.status==="Done").length;
  const ip=p.tasks.filter(t=>t.status==="In Progress").length;
  const open=p.tasks.filter(t=>t.status==="Open").length;
  const hold=p.tasks.filter(t=>t.status==="On Hold").length;
  return `
<h2><div class="pillar-head"><span>${p.titleAr} | ${p.title}</span><span class="pct">${prog}%</span></div></h2>
<div class="prog-bg"><div class="prog-fill" style="width:${prog}%"></div></div>
<div class="pillar-info">المسؤول: ${p.owner} &nbsp;|&nbsp; وزن العنوان: ${p.weight}% من المشروع &nbsp;|&nbsp; إجمالي المهام: ${p.tasks.length}</div>
<table>
<tr><th>المهمة</th><th>المسؤول</th><th>الموعد</th><th>الوزن</th><th>الحالة</th></tr>
${p.tasks.map(t=>`<tr><td>${t.title}</td><td>${t.owner}</td><td>${t.due}</td><td style="text-align:center">${t.weight}%</td><td style="color:${statusColor[t.status]||"#333"};font-weight:bold">${statusLabel[t.status]||t.status}</td></tr>`).join("")}
</table>
<div class="badges">
${done>0?`<span class="badge" style="color:#166534;background:#DCFCE7">✓ مكتمل ${done}</span>`:""}
${ip>0?`<span class="badge" style="color:#92400E;background:#FEF3C7">→ جاري ${ip}</span>`:""}
${open>0?`<span class="badge" style="color:#1E3A5F;background:#DBEAFE">○ مفتوح ${open}</span>`:""}
${hold>0?`<span class="badge" style="color:#7C3AED;background:#EDE9FE">⏸ معلق ${hold}</span>`:""}
</div>`;
}).join("")}
<footer>IPC Project · داخلي / Internal · ${today} · PM: Eng. Ahmed Aboelhasan</footer>
</body></html>`;

    const blob=new Blob(["\ufeff"+html],{type:"application/msword;charset=utf-8"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");
    a.href=url; a.download=`IPC_Board_Report_${new Date().toISOString().slice(0,10)}.doc`;
    a.click(); URL.revokeObjectURL(url);
    setWordStatus("done");
    setTimeout(()=>setWordStatus("idle"),3000);
  } catch(e){
    console.error(e);
    setWordStatus("error");
    setTimeout(()=>setWordStatus("idle"),3000);
  }
}

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  // حفظ تلقائي: لو في بيانات محفوظة نستخدمها، غير كده DEFAULT_DATA
  const [data, setData] = useState(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      return saved ? JSON.parse(saved) : DEFAULT_DATA;
    } catch { return DEFAULT_DATA; }
  });

  const [view, setView]             = useState("dashboard");
  const [addingTask, setAddingTask] = useState(null);
  const [newTask, setNewTask]       = useState({title:"",weight:10,status:"Open",owner:"",due:"TBD"});
  const [editingTask, setEditTask]  = useState(null);
  const [editingWeights, setEditW]  = useState(false);
  const [draftWeights, setDraftW]   = useState({});
  const [saveStatus, setSaveStatus] = useState("saved");
  const [wordStatus, setWordStatus] = useState("idle");

  const overall = useMemo(()=>calcOverall(data),[data]);
useEffect(() => {
  testConnection();
}, []);
useEffect(() => {
  loadData();
}, []);

async function loadData() {
  const { data: rows, error } = await supabase
    .from("project_data")
    .select("*")
    .limit(1);

  if (error) {
    console.log(error);
    return;
  }

  console.log("LOADED:", rows);

console.log("COPY THIS:");
console.log(JSON.stringify(DEFAULT_DATA));

if (
  rows.length > 0 &&
  Array.isArray(rows[0].data) &&
  rows[0].data.length > 0
) {
  setData(rows[0].data);
} else {
  setData(DEFAULT_DATA);
}}
async function testConnection() {
  const { data, error } = await supabase
    .from("project_data")
    .select("*");

  console.log("SUPABASE DATA:", data);
  console.log("SUPABASE ERROR:", error);
}
  // ── حفظ تلقائي في المتصفح ────────────────────────────────────────────────
useEffect(()=>{
  const t=setTimeout(async ()=>{
    try{

      // حفظ محلي
      localStorage.setItem(STORAGE_KEY,JSON.stringify(data));

      // حفظ في Supabase
      const { error } = await supabase
        .from("project_data")
        .update({
          data: data
        })
        .eq("id",1);

      if(error){
        console.log(error);
        setSaveStatus("error");
      }else{
        setSaveStatus("saved");
      }

    }catch(e){
      console.log(e);
      setSaveStatus("error");
    }
  },700);

  return ()=>clearTimeout(t);
},[data]);
  // ── Mutations ─────────────────────────────────────────────────────────────
  function updateTask(pillarId,taskId,field,value){
    setData(d=>d.map(p=>p.id!==pillarId?p:{
      ...p,tasks:p.tasks.map(t=>t.id!==taskId?t:{...t,[field]:field==="weight"?Number(value):value})
    }));
  }
  function deleteTask(pillarId,taskId){
    setData(d=>d.map(p=>p.id!==pillarId?p:{...p,tasks:p.tasks.filter(t=>t.id!==taskId)}));
    if(editingTask?.taskId===taskId) setEditTask(null);
  }
  function addTask(pillarId){
    if(!newTask.title.trim()) return;
    setData(d=>d.map(p=>p.id!==pillarId?p:{...p,tasks:[...p.tasks,{...newTask,id:`${pillarId}-${Date.now()}`,weight:Number(newTask.weight)}]}));
    setNewTask({title:"",weight:10,status:"Open",owner:"",due:"TBD"});
    setAddingTask(null);
  }
  function startEditWeights(){
    const w={};data.forEach(p=>{w[p.id]=p.weight;});setDraftW(w);setEditW(true);
  }
  function saveWeights(){
    const total=Object.values(draftWeights).reduce((s,v)=>s+Number(v),0);
    if(total!==100){alert(`مجموع الأوزان = ${total}%، لازم يكون 100%`);return;}
    setData(d=>d.map(p=>({...p,weight:Number(draftWeights[p.id])})));
    setEditW(false);
  }
  function resetData(){
    if(window.confirm("هتمسح كل التعديلات وترجع للبيانات الأصلية؟")){
      localStorage.removeItem(STORAGE_KEY);
      setData(DEFAULT_DATA);
    }
  }

  // ── Task Row (mobile-first) ───────────────────────────────────────────────
  function TaskRow({pillar,task}){
    const isEditing=editingTask?.pillarId===pillar.id&&editingTask?.taskId===task.id;
    if(isEditing) return(
      <div style={{padding:"10px 14px",borderBottom:`1px solid ${C.grayLight}`,background:"#F0F5FF"}}>
        <input value={task.title} onChange={e=>updateTask(pillar.id,task.id,"title",e.target.value)}
          style={{width:"100%",padding:"7px 8px",border:`1px solid ${C.blueMid}`,borderRadius:6,fontSize:13,boxSizing:"border-box",marginBottom:6}}/>
        <div style={{display:"flex",gap:6,marginBottom:6}}>
          <input value={task.owner} onChange={e=>updateTask(pillar.id,task.id,"owner",e.target.value)}
            placeholder="المسؤول" style={{flex:2,padding:"6px 8px",border:`1px solid ${C.grayLight}`,borderRadius:6,fontSize:12}}/>
          <input value={task.due} onChange={e=>updateTask(pillar.id,task.id,"due",e.target.value)}
            placeholder="الموعد" style={{flex:1,padding:"6px 8px",border:`1px solid ${C.grayLight}`,borderRadius:6,fontSize:12}}/>
        </div>
        <div style={{display:"flex",gap:6,alignItems:"center",marginBottom:8}}>
          <span style={{fontSize:12,color:C.gray,whiteSpace:"nowrap"}}>الوزن %</span>
          <input type="number" min="1" max="100" value={task.weight}
            onChange={e=>updateTask(pillar.id,task.id,"weight",e.target.value)}
            style={{width:60,padding:"6px 8px",border:`1px solid ${C.grayLight}`,borderRadius:6,fontSize:13,textAlign:"center"}}/>
          <select value={task.status} onChange={e=>updateTask(pillar.id,task.id,"status",e.target.value)}
            style={{flex:1,fontSize:12,border:`1px solid ${C.grayLight}`,borderRadius:6,padding:"6px 4px",background:C.white}}>
            {Object.keys(STATUS_CONFIG).map(s=><option key={s} value={s}>{STATUS_CONFIG[s].label}</option>)}
          </select>
        </div>
        <div style={{display:"flex",gap:8}}>
          <button onClick={()=>setEditTask(null)} style={{flex:2,padding:"9px",borderRadius:7,border:"none",background:C.blue,color:C.white,cursor:"pointer",fontSize:13,fontWeight:700}}>✓ حفظ</button>
          <button onClick={()=>deleteTask(pillar.id,task.id)} style={{flex:1,padding:"9px",borderRadius:7,border:"none",background:C.redBg,color:C.red,cursor:"pointer",fontSize:13,fontWeight:700}}>🗑 حذف</button>
        </div>
      </div>
    );
    return(
      <div style={{padding:"10px 14px",borderBottom:`1px solid ${C.grayLight}`,display:"flex",alignItems:"center",gap:8}}>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontSize:13,fontWeight:600,color:task.status==="Done"?C.gray:C.navy,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{task.title}</div>
          <div style={{fontSize:11,color:C.gray,marginTop:2}}>{task.owner} · {task.due}</div>
        </div>
        <span style={{fontSize:11,fontWeight:700,color:C.blueMid,background:C.blueLight,padding:"2px 8px",borderRadius:99,whiteSpace:"nowrap"}}>{task.weight}%</span>
        <select value={task.status} onChange={e=>updateTask(pillar.id,task.id,"status",e.target.value)}
          style={{fontSize:11,border:`1px solid ${C.grayLight}`,borderRadius:6,padding:"4px 3px",background:C.white,cursor:"pointer"}}>
          {Object.keys(STATUS_CONFIG).map(s=><option key={s} value={s}>{STATUS_CONFIG[s].label}</option>)}
        </select>
        <button onClick={()=>setEditTask({pillarId:pillar.id,taskId:task.id})} style={{
          minWidth:36,minHeight:36,padding:"4px",borderRadius:6,border:`1px solid ${C.grayLight}`,
          background:C.white,cursor:"pointer",fontSize:16,display:"flex",alignItems:"center",justifyContent:"center"}}>✎</button>
        <button onClick={()=>deleteTask(pillar.id,task.id)} style={{
          minWidth:36,minHeight:36,padding:"4px",borderRadius:6,border:"none",
          background:C.redBg,cursor:"pointer",fontSize:16,color:C.red,
          display:"flex",alignItems:"center",justifyContent:"center"}}>🗑</button>
      </div>
    );
  }

  const saveLabel=saveStatus==="saving"?"⏳ جاري الحفظ...":saveStatus==="saved"?"✓ محفوظ":"⚠ خطأ في الحفظ";
  const saveColor=saveStatus==="saved"?C.done:saveStatus==="saving"?C.progress:C.red;

  // ════════════════════════════════════════════════════════════════════════════
  return(
    <div style={{fontFamily:"'Segoe UI',Arial,sans-serif",background:C.offwhite,minHeight:"100vh",direction:"rtl"}}>

      {/* Header */}
      <div style={{background:C.blue,color:C.white,padding:"12px 18px",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
        <div>
          <div style={{fontSize:17,fontWeight:800}}>IPC Project — لوحة إدارة المشروع</div>
          <div style={{fontSize:11,opacity:0.65,marginTop:1}}>PM: Eng. Ahmed Aboelhasan</div>
        </div>
        <div style={{display:"flex",gap:6,flexWrap:"wrap",alignItems:"center"}}>
          {[["dashboard","الرئيسية"],["board","Kanban"],["report","تقرير"]].map(([v,l])=>(
            <button key={v} onClick={()=>setView(v)} style={{
              padding:"5px 12px",borderRadius:6,border:"none",cursor:"pointer",fontSize:12,fontWeight:600,
              background:view===v?C.white:"rgba(255,255,255,0.18)",color:view===v?C.blue:C.white
            }}>{l}</button>
          ))}
          <button onClick={()=>exportExcel(data)} style={{padding:"5px 12px",borderRadius:6,border:"1px solid rgba(255,255,255,0.35)",background:"rgba(255,255,255,0.12)",color:C.white,cursor:"pointer",fontSize:12,fontWeight:600}}>
            📊 Excel
          </button>

<button
  onClick={()=>backupData(data)}
  style={{
    padding:"5px 12px",
    borderRadius:6,
    border:"1px solid rgba(255,255,255,0.35)",
    background:"rgba(255,255,255,0.12)",
    color:C.white,
    cursor:"pointer",
    fontSize:12,
    fontWeight:600
  }}
>
  💾 Backup
</button>

<button
  onClick={()=>restoreBackup(setData)}
  style={{
    padding:"5px 12px",
    borderRadius:6,
    border:"1px solid rgba(255,255,255,0.35)",
    background:"rgba(255,255,255,0.12)",
    color:C.white,
    cursor:"pointer",
    fontSize:12,
    fontWeight:600
  }}
>
  🔄 Restore
</button>
                 </div>
      </div>

      {/* Overall + save indicator */}
      <div style={{background:C.navy,padding:"10px 18px",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
        <span style={{color:C.white,fontWeight:700,fontSize:13,whiteSpace:"nowrap"}}>تقدم المشروع الكلي</span>
        <div style={{flex:1,minWidth:100}}><ProgressBar pct={overall} height={12}/></div>
        <span style={{color:C.white,fontWeight:900,fontSize:20,minWidth:44,textAlign:"center"}}>{overall}%</span>
        <button onClick={startEditWeights} style={{padding:"4px 10px",borderRadius:5,border:"1px solid rgba(255,255,255,0.3)",background:"transparent",color:C.white,cursor:"pointer",fontSize:12}}>⚙ الأوزان</button>
        <span style={{fontSize:11,fontWeight:600,color:saveColor,background:"rgba(255,255,255,0.08)",padding:"3px 9px",borderRadius:5,whiteSpace:"nowrap"}}>{saveLabel}</span>
        <button onClick={resetData} title="إعادة تعيين البيانات" style={{padding:"3px 8px",borderRadius:5,border:"1px solid rgba(255,120,120,0.4)",background:"transparent",color:"rgba(255,160,160,0.9)",cursor:"pointer",fontSize:11}}>↺</button>
      </div>

      {/* Weights modal */}
      {editingWeights&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
          <div style={{background:C.white,borderRadius:12,padding:24,width:"100%",maxWidth:360,boxShadow:"0 20px 60px rgba(0,0,0,0.3)"}}>
            <div style={{fontSize:16,fontWeight:800,marginBottom:4,color:C.blue}}>أوزان العناوين الرئيسية</div>
            <div style={{fontSize:12,color:C.gray,marginBottom:14}}>المجموع لازم يساوي 100%</div>
            {data.map(p=>(
              <div key={p.id} style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}>
                <span style={{flex:1,fontSize:13,fontWeight:600}}>{p.titleAr}</span>
                <input type="number" min="0" max="100" value={draftWeights[p.id]??p.weight}
                  onChange={e=>setDraftW(w=>({...w,[p.id]:e.target.value}))}
                  style={{width:60,padding:"5px 8px",border:`1px solid ${C.grayLight}`,borderRadius:5,textAlign:"center",fontSize:14}}/>
                <span style={{color:C.gray,fontSize:12}}>%</span>
              </div>
            ))}
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:12}}>
              <span style={{fontSize:13,fontWeight:700,color:Object.values(draftWeights).reduce((s,v)=>s+Number(v),0)===100?C.done:C.red}}>
                المجموع: {Object.values(draftWeights).reduce((s,v)=>s+Number(v),0)}%
              </span>
              <div style={{display:"flex",gap:8}}>
                <button onClick={()=>setEditW(false)} style={{padding:"6px 14px",borderRadius:6,border:`1px solid ${C.grayLight}`,background:C.white,cursor:"pointer"}}>إلغاء</button>
                <button onClick={saveWeights} style={{padding:"6px 14px",borderRadius:6,border:"none",background:C.blue,color:C.white,cursor:"pointer",fontWeight:700}}>حفظ</button>
              </div>
            </div>
          </div>
        </div>
      )}

      <div style={{padding:"16px 18px"}}>

        {/* ══ DASHBOARD ════════════════════════════════════════════════════ */}
        {view==="dashboard"&&(
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(360px,1fr))",gap:14}}>
            {data.map(pillar=>{
              const prog=calcProgress(pillar.tasks);
              return(
                <div key={pillar.id} style={{background:C.white,borderRadius:10,overflow:"hidden",boxShadow:"0 1px 4px rgba(0,0,0,0.07)",border:`1px solid ${C.grayLight}`}}>
                  <div style={{background:C.blue,padding:"11px 14px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                    <div>
                      <div style={{color:C.white,fontWeight:800,fontSize:15}}>{pillar.titleAr}</div>
                      <div style={{color:"rgba(255,255,255,0.6)",fontSize:11,marginTop:1}}>{pillar.title} · {pillar.owner}</div>
                    </div>
                    <div style={{textAlign:"center"}}>
                      <div style={{color:C.white,fontWeight:900,fontSize:22}}>{prog}%</div>
                      <div style={{color:"rgba(255,255,255,0.55)",fontSize:10}}>وزن {pillar.weight}%</div>
                    </div>
                  </div>
                  <div style={{padding:"8px 14px 4px",display:"flex",alignItems:"center",gap:8}}>
                    <div style={{flex:1}}><ProgressBar pct={prog} height={5}/></div>
                    <WeightWarning tasks={pillar.tasks}/>
                  </div>
                  <div style={{borderTop:`1px solid ${C.grayLight}`,marginTop:4}}>
                    {pillar.tasks.map(task=><TaskRow key={task.id} pillar={pillar} task={task}/>)}
                    {addingTask===pillar.id?(
                      <div style={{padding:"10px 14px",background:"#F0F5FF",borderTop:`1px solid ${C.blueLight}`}}>
                        <div style={{fontSize:11,fontWeight:700,color:C.blue,marginBottom:6}}>مهمة جديدة</div>
                        <input placeholder="عنوان المهمة *" value={newTask.title} onChange={e=>setNewTask(n=>({...n,title:e.target.value}))}
                          style={{width:"100%",padding:"5px 8px",borderRadius:5,border:`1px solid ${C.grayLight}`,marginBottom:6,fontSize:12,boxSizing:"border-box"}}/>
                        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6,marginBottom:6}}>
                          <input placeholder="المسؤول" value={newTask.owner} onChange={e=>setNewTask(n=>({...n,owner:e.target.value}))}
                            style={{padding:"5px 6px",borderRadius:5,border:`1px solid ${C.grayLight}`,fontSize:11}}/>
                          <input placeholder="الموعد" value={newTask.due} onChange={e=>setNewTask(n=>({...n,due:e.target.value}))}
                            style={{padding:"5px 6px",borderRadius:5,border:`1px solid ${C.grayLight}`,fontSize:11}}/>
                        </div>
                        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                          <label style={{fontSize:11,color:C.gray,whiteSpace:"nowrap"}}>الوزن %</label>
                          <input type="number" min="1" max="100" value={newTask.weight} onChange={e=>setNewTask(n=>({...n,weight:e.target.value}))}
                            style={{width:58,padding:"4px 6px",borderRadius:5,border:`1px solid ${C.grayLight}`,fontSize:12,textAlign:"center"}}/>
                          <span style={{fontSize:11,color:taskWeightSum(pillar.tasks)+Number(newTask.weight)>100?C.red:C.done}}>
                            المجموع: {taskWeightSum(pillar.tasks)+Number(newTask.weight)}%
                          </span>
                          <select value={newTask.status} onChange={e=>setNewTask(n=>({...n,status:e.target.value}))}
                            style={{fontSize:11,border:`1px solid ${C.grayLight}`,borderRadius:5,padding:"4px 3px"}}>
                            {Object.keys(STATUS_CONFIG).map(s=><option key={s} value={s}>{STATUS_CONFIG[s].label}</option>)}
                          </select>
                        </div>
                        <div style={{display:"flex",gap:6}}>
                          <button onClick={()=>addTask(pillar.id)} style={{flex:1,padding:"7px",borderRadius:5,border:"none",background:C.blue,color:C.white,cursor:"pointer",fontSize:12,fontWeight:700}}>+ إضافة</button>
                          <button onClick={()=>setAddingTask(null)} style={{padding:"7px 12px",borderRadius:5,border:`1px solid ${C.grayLight}`,background:C.white,cursor:"pointer",fontSize:12}}>إلغاء</button>
                        </div>
                      </div>
                    ):(
                      <button onClick={()=>setAddingTask(pillar.id)} style={{
                        width:"100%",padding:"8px",border:"none",background:"transparent",
                        color:C.blueMid,cursor:"pointer",fontSize:12,fontWeight:600,
                        borderTop:`1px dashed ${C.blueLight}`}}>+ إضافة مهمة</button>
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {/* ══ KANBAN ═══════════════════════════════════════════════════════ */}
        {view==="board"&&(
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:12}}>
            {Object.entries(STATUS_CONFIG).map(([status,cfg])=>{
              const all=data.flatMap(p=>p.tasks.filter(t=>t.status===status).map(t=>({...t,pillarId:p.id,pillarAr:p.titleAr})));
              return(
                <div key={status}>
                  <div style={{padding:"7px 12px",borderRadius:"7px 7px 0 0",background:cfg.bg,display:"flex",justifyContent:"space-between",marginBottom:8}}>
                    <span style={{fontWeight:800,color:cfg.color,fontSize:13}}>{cfg.label}</span>
                    <span style={{fontWeight:700,color:cfg.color,fontSize:12,background:C.white,borderRadius:99,padding:"1px 8px"}}>{all.length}</span>
                  </div>
                  <div style={{display:"flex",flexDirection:"column",gap:8}}>
                    {all.map(task=>(
                      <div key={task.id} style={{background:C.white,borderRadius:7,padding:"10px 12px",boxShadow:"0 1px 3px rgba(0,0,0,0.06)",border:`1px solid ${C.grayLight}`}}>
                        <div style={{fontSize:12,fontWeight:700,color:C.navy,marginBottom:3}}>{task.title}</div>
                        <div style={{fontSize:10,color:C.gray}}>{task.pillarAr}</div>
                        <div style={{fontSize:10,color:C.gray,marginTop:1}}>{task.owner}</div>
                        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:8}}>
                          <span style={{fontSize:10,color:C.blueMid,fontWeight:700,background:C.blueLight,padding:"1px 6px",borderRadius:99}}>{task.weight}%</span>
                          <select value={task.status} onChange={e=>updateTask(task.pillarId,task.id,"status",e.target.value)}
                            style={{fontSize:10,border:`1px solid ${C.grayLight}`,borderRadius:4,padding:"2px 3px",background:C.white,cursor:"pointer"}}>
                            {Object.keys(STATUS_CONFIG).map(s=><option key={s} value={s}>{STATUS_CONFIG[s].label}</option>)}
                          </select>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {/* ══ REPORT ═══════════════════════════════════════════════════════ */}
        {view==="report"&&(
          <div style={{maxWidth:860,margin:"0 auto"}}>
            <div style={{background:C.blue,color:C.white,borderRadius:"10px 10px 0 0",padding:"18px 24px"}}>
              <div style={{fontSize:11,opacity:0.6,marginBottom:3}}>تقرير تقدم المشروع – للشركاء · داخلي</div>
              <div style={{fontSize:20,fontWeight:900}}>IPC Project — تقرير تنفيذي</div>
              <div style={{fontSize:12,opacity:0.65,marginTop:3}}>{new Date().toLocaleDateString("ar-EG")} · PM: Eng. Ahmed Aboelhasan</div>
            </div>
            <div style={{background:C.navy,padding:"12px 24px",display:"flex",alignItems:"center",gap:14}}>
              <div style={{flex:1}}><div style={{color:"rgba(255,255,255,0.6)",fontSize:11,marginBottom:4}}>التقدم الكلي</div><ProgressBar pct={overall} height={14}/></div>
              <div style={{color:C.white,fontWeight:900,fontSize:30}}>{overall}%</div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",background:C.white,borderBottom:`1px solid ${C.grayLight}`}}>
              {[
                ["إجمالي المهام",data.flatMap(p=>p.tasks).length,C.blue],
                ["مكتملة",data.flatMap(p=>p.tasks).filter(t=>t.status==="Done").length,C.done],
                ["جارية",data.flatMap(p=>p.tasks).filter(t=>t.status==="In Progress").length,C.progress],
                ["معلقة",data.flatMap(p=>p.tasks).filter(t=>t.status==="On Hold").length,C.hold],
              ].map(([l,v,col])=>(
                <div key={l} style={{padding:"14px 16px",borderLeft:`1px solid ${C.grayLight}`,textAlign:"center"}}>
                  <div style={{fontSize:24,fontWeight:900,color:col}}>{v}</div>
                  <div style={{fontSize:11,color:C.gray,marginTop:2}}>{l}</div>
                </div>
              ))}
            </div>
            {data.map(pillar=>{
              const prog=calcProgress(pillar.tasks);
              return(
                <div key={pillar.id} style={{background:C.white,borderTop:`2px solid ${C.blue}`,marginTop:14,borderRadius:8,overflow:"hidden",boxShadow:"0 1px 4px rgba(0,0,0,0.05)"}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"12px 20px",background:C.blueLight}}>
                    <div>
                      <span style={{fontWeight:800,color:C.blue,fontSize:14}}>{pillar.titleAr}</span>
                      <span style={{color:C.gray,fontSize:12,marginRight:8}}>{pillar.title}</span>
                    </div>
                    <div style={{display:"flex",alignItems:"center",gap:12}}>
                      <span style={{fontSize:11,color:C.gray}}>وزن {pillar.weight}%</span>
                      <span style={{fontWeight:900,fontSize:20,color:prog===100?C.done:C.blue}}>{prog}%</span>
                    </div>
                  </div>
                  <div style={{padding:"6px 20px 3px"}}><ProgressBar pct={prog} height={5}/></div>
                  <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                    <thead>
                      <tr style={{background:C.offwhite}}>
                        {["المهمة","المسؤول","الموعد","الوزن","الحالة"].map(h=>(
                          <th key={h} style={{padding:"6px 12px",color:C.gray,fontWeight:700,textAlign:"right",borderBottom:`1px solid ${C.grayLight}`,fontSize:11}}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {pillar.tasks.map((task,i)=>(
                        <tr key={task.id} style={{background:i%2===0?C.white:C.offwhite}}>
                          <td style={{padding:"7px 12px",color:C.navy,fontWeight:600}}>{task.title}</td>
                          <td style={{padding:"7px 12px",color:C.gray}}>{task.owner}</td>
                          <td style={{padding:"7px 12px",color:C.gray}}>{task.due}</td>
                          <td style={{padding:"7px 12px",textAlign:"center"}}>
                            <span style={{fontSize:11,fontWeight:700,color:C.blueMid,background:C.blueLight,padding:"1px 7px",borderRadius:99}}>{task.weight}%</span>
                          </td>
                          <td style={{padding:"7px 12px"}}><StatusBadge status={task.status}/></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  <div style={{padding:"7px 20px 10px",display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
                    {Object.entries(STATUS_CONFIG).map(([s,cfg])=>{
                      const cnt=pillar.tasks.filter(t=>t.status===s).length;
                      return cnt>0?<span key={s} style={{fontSize:11,color:cfg.color,background:cfg.bg,padding:"2px 8px",borderRadius:99}}>{cfg.label} {cnt}</span>:null;
                    })}
                    <WeightWarning tasks={pillar.tasks}/>
                  </div>
                </div>
              );
            })}
            <div style={{marginTop:14,padding:"12px 18px",background:C.white,borderRadius:8,border:`1px solid ${C.grayLight}`,fontSize:11,color:C.gray,textAlign:"center"}}>
              IPC Project · داخلي / Internal · {new Date().toLocaleDateString("ar-EG")}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
