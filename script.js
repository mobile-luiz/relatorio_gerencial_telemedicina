let allData=[];
let filteredData=[];
let showAllRows=false;

const $=id=>document.getElementById(id);
const norm=s=>String(s??"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().trim();

function keyMap(row){
  const map={};
  Object.keys(row).forEach(k=>map[norm(k)]=row[k]);
  return map;
}
function value(row, names){
  const m=keyMap(row);
  // 1) correspondência exata
  for(const n of names){
    const wanted=norm(n);
    if(m[wanted]!==undefined && m[wanted]!=="" && m[wanted]!==null) return m[wanted];
  }
  // 2) correspondência parcial, ignorando acentos
  const keys=Object.keys(m);
  for(const n of names){
    const wanted=norm(n);
    const k=keys.find(k=>k===wanted || k.includes(wanted) || wanted.includes(k));
    if(k && m[k]!=="" && m[k]!==null) return m[k];
  }
  return "";
}
function findTimeColumn(row, kind){
  const keys=Object.keys(row);
  const preferred = kind==="inicio" ? [
    "horario inicio","horário inicio","hora inicio","hora de inicio",
    "horario entrada","horário entrada","hora entrada","entrada",
    "inicio atendimento","início atendimento","inicio","início",
    "hora inicial","horario inicial","horário inicial"
  ] : [
    "horario fim","horário fim","hora fim","hora de fim",
    "horario saida","horário saída","hora saida","saída","saida",
    "fim atendimento","fim","hora final","horario final","horário final"
  ];
  const v=value(row,preferred);
  if(v!=="" && v!==null && v!==undefined)return v;

  // Fallback: procura cabeçalhos que contenham início/entrada ou fim/saída.
  const candidates=keys.filter(k=>{
    const x=norm(k);
    if(kind==="inicio") return /(inicio|entrada|inicial)/.test(x) && /(hora|horario|atendimento|data)?/.test(x);
    return /(fim|saida|final)/.test(x) && /(hora|horario|atendimento|data)?/.test(x);
  });
  return candidates.length ? row[candidates[0]] : "";
}
function normalizeRow(r,i){
  // Este CSV tem campos específicos do sistema de telemedicina.
  // Para o relatório gerencial:
  // Data = data_fim (quando existe), depois finalização, data_turno e convite.
  // Início = data_ultimo_entrada_medico_vc.
  // Fim = data_ultima_finalização_medico.
  // Tempo = t4_duração_consulta.
  const dataFim=value(r,["data_fim"]);
  const dataFinal=value(r,["data_ultima_finalização_medico"]);
  const dataTurno=value(r,["data_turno"]);
  const dataConvite=value(r,["data_convite"]);
  const data=dataFim||dataFinal||dataTurno||dataConvite;

  const inicio=value(r,["data_ultimo_entrada_medico_vc"]);
  const fim=value(r,["data_ultima_finalização_medico"]);
  const duracao=value(r,["t4_duração_consulta"]);
  const paciente=value(r,["nome","paciente","nome paciente"]);
  const especialidade=value(r,["especialidade","especialidade medica","especialidade médica"]);
  const situacao=value(r,["situação_consulta","situacao_consulta","status","situacao","situação"]);
  const convite=value(r,["usuario_convite"]);
  const d=parseDate(data);

  if(!d && !paciente && !especialidade)return null;

  // F e C são códigos do campo situação_consulta.
  // Quando existe t4_duração_consulta, temos uma consulta efetivamente realizada.
  let status;
  if(duracao) status="Concluído";
  else if(situacao) status=String(situacao).trim();
  else if(dataFim) status="Encerrado";
  else status="Agendado";

  return {
    data:d,
    inicio,
    fim,
    duracao,
    paciente:String(paciente||"Não informado").trim(),
    especialidade:String(especialidade||"Não informado").trim(),
    status,
    canal:String(convite||"").trim(),
    situacao:String(situacao||"").trim()
  };
}
function parseDate(v){
  if(v===null||v===undefined||v==="") return null;
  if(v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(),v.getMonth(),v.getDate());
  if(typeof v==="number" && window.XLSX){
    const d=XLSX.SSF.parse_date_code(v);
    if(d) return new Date(d.y,d.m-1,d.d);
  }
  const s=String(v).trim();
  if(!s)return null;
  let m=s.match(/^(\d{4})[-\/]([0-9]{1,2})[-\/]([0-9]{1,2})/);
  if(m)return new Date(+m[1],+m[2]-1,+m[3]);
  m=s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})/);
  if(m){let y=+m[3];if(y<100)y+=2000;return new Date(y,+m[2]-1,+m[1]);}
  const d=new Date(s);
  return isNaN(d)?null:new Date(d.getFullYear(),d.getMonth(),d.getDate());
}
function dateISO(d){return d?`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`:"";}
function dateBR(d){return d?`${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}/${d.getFullYear()}`:"—";}
function timeMinutes(v){
  if(v===null||v===undefined||v==="")return null;
  if(v instanceof Date&&!isNaN(v))return v.getHours()*60+v.getMinutes();
  if(typeof v==="number"){
    if(v>=0&&v<1)return Math.round(v*1440);
    if(v>=0&&v<1440)return Math.round(v);
  }
  let s=String(v).trim().toLowerCase();
  let m=s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if(!m)m=s.match(/\b(\d{1,2})h(\d{2})?/);
  if(m){let h=+m[1],mi=+(m[2]||0);if(h<24&&mi<60)return h*60+mi;}
  return null;
}
function timeBR(v){const m=timeMinutes(v);return m===null?"":`${String(Math.floor(m/60)%24).padStart(2,"0")}:${String(m%60).padStart(2,"0")}`;}
function duration(row){
  // O campo t4 é a duração oficial da consulta e deve ser usado quando disponível.
  const explicit=timeDurationToSeconds(row.duracao);
  if(explicit!==null) return Math.round(explicit/60);

  const a=timeMinutes(row.inicio),b=timeMinutes(row.fim);
  if(a===null||b===null)return null;
  let d=b-a;if(d<0)d+=1440;
  return d;
}
function timeDurationToSeconds(v){
  if(v===null||v===undefined||v==="")return null;
  const s=String(v).trim();
  const m=s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if(!m)return null;
  return +m[1]*3600 + +m[2]*60 + +(m[3]||0);
}
function formatDuration(v){
  if(v===null || v===undefined || v<0) return "—";
  if(v<60) return `${v} min`;
  const h=Math.floor(v/60), m=v%60;
  return h>0 ? `${String(h).padStart(2,"0")}:${String(m).padStart(2,"0")}` : `${m} min`;
}
async function readFile(file){
  const name=file.name.toLowerCase();
  if(name.endsWith(".csv")){
    const buf=await file.arrayBuffer();
    let text=new TextDecoder("utf-8",{fatal:false}).decode(buf);
    // Corrige CSV Windows-1252 quando UTF-8 produzir caracteres inválidos
    if(text.includes("�")) text=new TextDecoder("windows-1252").decode(buf);
    const wb=XLSX.read(text,{type:"string",raw:true,codepage:1252});
    return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:"",raw:true});
  }
  const buf=await file.arrayBuffer();
  const wb=XLSX.read(buf,{type:"array",raw:true,cellDates:true});
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:""});
}
$("fileInput").addEventListener("change",async e=>{
  const file=e.target.files[0]; if(!file)return;
  $("fileStatus").textContent="Lendo arquivo...";
  try{
    const raw=await readFile(file);
    allData=raw.map(normalizeRow).filter(Boolean).filter(r=>r.data);
    if(!allData.length)throw new Error("Nenhum registro com data foi encontrado.");
    setupFilters();
    setDateLimits();
    applyFilters();
    $("emptyState").classList.add("hidden");
    $("report").classList.remove("hidden");
    $("fileStatus").textContent=`✓ ${allData.length.toLocaleString("pt-BR")} registros importados`;
  }catch(err){
    console.error(err); $("fileStatus").textContent="Erro na importação";
    alert("Não foi possível ler os dados. Verifique as colunas de data, paciente e atendimento.");
  }
});
function setupFilters(){
  const statuses=[...new Set(allData.map(r=>r.status).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
  $("statusFilter").innerHTML='<option value="Todos">Todos</option>'+statuses.map(x=>`<option>${escapeHTML(x)}</option>`).join("");
  const specs=[...new Set(allData.map(r=>r.especialidade).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
  $("specialtyFilter").innerHTML='<option value="Todas">Todas</option>'+specs.map(x=>`<option>${escapeHTML(x)}</option>`).join("");
}
function setDateLimits(){
  const dates=allData.map(r=>r.data).sort((a,b)=>a-b), min=dateISO(dates[0]),max=dateISO(dates.at(-1));
  $("startDate").min=min;$("startDate").max=max;$("endDate").min=min;$("endDate").max=max;
  $("startDate").value=min;$("endDate").value=max;
}
$("applyBtn").addEventListener("click",applyFilters);
$("statusFilter").addEventListener("change",applyFilters);
$("specialtyFilter").addEventListener("change",applyFilters);
$("startDate").addEventListener("change",applyFilters);
$("endDate").addEventListener("change",applyFilters);
$("clearBtn").addEventListener("click",()=>{if(!allData.length)return;setDateLimits();$("statusFilter").value="Todos";$("specialtyFilter").value="Todas";applyFilters();});
$("chartMode").addEventListener("change",renderCharts);
$("timeMode").addEventListener("change",renderCharts);
function applyFilters(){
  if(!allData.length)return;
  const a=$("startDate").value?new Date($("startDate").value+"T00:00:00"):null;
  const b=$("endDate").value?new Date($("endDate").value+"T23:59:59"):null;
  const st=$("statusFilter").value, sp=$("specialtyFilter").value;
  filteredData=allData.filter(r=>(!a||r.data>=a)&&(!b||r.data<=b)&&(st==="Todos"||r.status===st)&&(sp==="Todas"||r.especialidade===sp));
  showAllRows=false; updateReport();
}
function updateReport(){
  const n=filteredData.length;
  $("total").textContent=n.toLocaleString("pt-BR");
  const patients=new Set(filteredData.map(r=>norm(r.paciente)).filter(Boolean));
  $("patients").textContent=patients.size.toLocaleString("pt-BR");
  const completed=filteredData.filter(r=>norm(r.status).includes("concl")).length;
  $("completion").textContent=n?Math.round(completed/n*100)+"%":"0%";
  $("statusInfo").textContent=n?`${completed} concluídos`:"—";
  const durations=filteredData.map(duration).filter(x=>x!==null&&x>=0);
  const avg=durations.length?Math.round(durations.reduce((a,b)=>a+b,0)/durations.length):0;
  $("avg").textContent=`${String(Math.floor(avg/60)).padStart(2,"0")}:${String(avg%60).padStart(2,"0")}`;
  $("avgInfo").textContent=durations.length?`${durations.length} atendimentos com horários`:"Sem horários válidos";
  $("patientInfo").textContent=n?`${patients.size} pacientes distintos`:"—";
  const dates=filteredData.map(r=>r.data).sort((a,b)=>a-b);
  $("periodLabel").textContent=dates.length?`${dateBR(dates[0])} - ${dateBR(dates.at(-1))}`:"Nenhum dado";
  $("footerInfo").textContent=`${n.toLocaleString("pt-BR")} registros | Atualizado agora`;
  renderTable();renderCharts();
}
function renderTable(){
  // Mostra os registros que correspondem aos filtros.
  // Quando início/fim não existem, o tempo é exibido como "—".
  const rowsSorted=filteredData.slice().sort((a,b)=>{
    const dateDiff=b.data-a.data;
    if(dateDiff!==0)return dateDiff;
    const ta=timeMinutes(b.inicio), tb=timeMinutes(a.inicio);
    return (tb??-1)-(ta??-1);
  });
  const rows=showAllRows?rowsSorted:rowsSorted.slice(0,5);

  const hasFilter=$("statusFilter").value!=="Todos" ||
                  $("specialtyFilter").value!=="Todas" ||
                  $("startDate").value!==$("startDate").min ||
                  $("endDate").value!==$("endDate").max;

  $("tableTitle").textContent=hasFilter?"Atendimentos Filtrados":"Últimos Atendimentos";
  $("showAll").textContent=showAllRows?"Mostrar menos":"Ver todos";

  $("tableBody").innerHTML=rows.length?rows.map(r=>`<tr>
    <td>${dateBR(r.data)}</td>
    <td>${timeBR(r.inicio)||"—"}</td>
    <td>${timeBR(r.fim)||"—"}</td>
    <td>${escapeHTML(r.paciente)}</td>
    <td>${escapeHTML(r.especialidade)}</td>
    <td><span class="duration-badge">${formatDuration(duration(r))}</span></td>
  </tr>`).join(""):`<tr><td colspan="6" style="text-align:center;padding:25px;color:#8195ac">Nenhum atendimento encontrado para os filtros selecionados.</td></tr>`;
}
$("showAll").addEventListener("click",()=>{showAllRows=!showAllRows;$("showAll").textContent=showAllRows?"Mostrar menos":"Ver todos";renderTable();});
function badgeClass(s){const x=norm(s);return x.includes("concl")?"ok":(x.includes("cancel")?"cancel":"warn")}
function groupCount(data,field){const o={};data.forEach(r=>{const k=r[field]||"Não informado";o[k]=(o[k]||0)+1});return Object.entries(o).sort((a,b)=>b[1]-a[1]);}
function renderCharts(){renderBars();renderSpecialties();renderStatus();renderTime();}
function renderBars(){
  const mode=$("chartMode").value;
  const agendadoMap={}, realizadoMap={};

  // Agendado: usa a data_turno/data agendada.
  // Realizado: usa a data_fim e considera somente registros efetivamente encerrados/finalizados.
  filteredData.forEach(r=>{
    const da=r.dataAgendada || r.data;
    if(da){
      const key=mode==="weekly"?weekKey(da):dateISO(da);
      agendadoMap[key]=(agendadoMap[key]||0)+1;
    }
    const x=norm(r.status);
    const realizado = r.data && (
      x.includes("encerr") || x.includes("conclu") || x.includes("realiz")
      || (timeMinutes(r.inicio)!==null && timeMinutes(r.fim)!==null)
    );
    if(realizado && r.data){
      const key=mode==="weekly"?weekKey(r.data):dateISO(r.data);
      realizadoMap[key]=(realizadoMap[key]||0)+1;
    }
  });

  const keys=[...new Set([...Object.keys(agendadoMap),...Object.keys(realizadoMap)])].sort();
  const el=$("barChart"),axis=$("barAxis");
  el.innerHTML="";axis.innerHTML="";
  if(!keys.length){
    el.innerHTML='<div class="chart-empty">Sem dados para o período</div>';
    return;
  }

  const max=Math.max(...keys.flatMap(k=>[agendadoMap[k]||0,realizadoMap[k]||0]),1);
  keys.forEach(k=>{
    const group=document.createElement("div");
    group.className="bar-group";

    const realizado=document.createElement("div");
    realizado.className="bar bar-realizado";
    realizado.style.height=Math.max((realizadoMap[k]||0)/max*100,2)+"%";
    realizado.title=`Realizado: ${realizadoMap[k]||0}`;

    const agendado=document.createElement("div");
    agendado.className="bar bar-agendado";
    agendado.style.height=Math.max((agendadoMap[k]||0)/max*100,2)+"%";
    agendado.title=`Agendado: ${agendadoMap[k]||0}`;

    group.append(realizado,agendado);
    el.appendChild(group);
  });

  const labels=keys.length<=7?keys:keys.filter((_,i)=>i===0||i===keys.length-1||i%Math.ceil(keys.length/6)===0);
  labels.forEach(k=>{
    const s=document.createElement("span");
    const d=new Date(k+"T00:00:00");
    s.textContent=dateBR(d).slice(0,5);
    axis.appendChild(s);
  });
}
function weekKey(d){
  const x=new Date(d),day=(x.getDay()+6)%7;
  x.setDate(x.getDate()-day);
  return dateISO(x);
}
function donutSVG(groups,total,colors,centerLabel){
  const size=220, cx=110, cy=110, radius=78, stroke=34;
  if(!groups.length || !total) return `<div class="donut-fallback"><strong>0</strong><span>${centerLabel}</span></div>`;
  let cumulative=0;
  const circumference=2*Math.PI*radius;
  const circles=groups.map(([name,value],i)=>{
    const pct=value/total;
    const dash=pct*circumference;
    const gap=Math.max(circumference-dash,0);
    const offset=-cumulative*circumference;
    cumulative+=pct;
    return `<circle cx="${cx}" cy="${cy}" r="${radius}" fill="none" stroke="${colors[i%colors.length]}" stroke-width="${stroke}" stroke-dasharray="${dash} ${gap}" stroke-dashoffset="${offset}" transform="rotate(-90 ${cx} ${cy})"></circle>`;
  }).join("");
  return `<svg class="donut-svg" viewBox="0 0 ${size} ${size}" aria-hidden="true">
    <circle cx="${cx}" cy="${cy}" r="${radius}" fill="none" stroke="#edf4fb" stroke-width="${stroke}"></circle>
    ${circles}
    <circle cx="${cx}" cy="${cy}" r="52" fill="#fff"></circle>
    <text x="${cx}" y="106" text-anchor="middle" class="donut-number">${total.toLocaleString("pt-BR")}</text>
    <text x="${cx}" y="128" text-anchor="middle" class="donut-label">${centerLabel}</text>
  </svg>`;
}
function renderSpecialties(){
  const groups=groupCount(filteredData,"especialidade"),total=filteredData.length;
  $("donutTotal").textContent=total.toLocaleString("pt-BR");
  const colors=["#1269c7","#2888e8","#4d9ee9","#6eb1ed","#83bdf0","#9ccbf3","#b9daf5"];
  const donut=$("specialtyDonut");
  donut.innerHTML=donutSVG(groups,total,colors,"Atendimentos");
  donut.style.background="none";
  $("specialtyLegend").innerHTML=groups.slice(0,7).map(([k,v],i)=>`<li><i style="background:${colors[i%colors.length]}"></i>${escapeHTML(k)} <b>${Math.round(v/total*100)}%</b></li>`).join("");
}
function renderStatus(){
  const groups=groupCount(filteredData,"status"),total=filteredData.length;
  $("statusTotal").textContent=total.toLocaleString("pt-BR");
  const palette=["#18b875","#2987db","#ff4c32","#f2a93b","#8c6be8"];
  const donut=$("statusDonut");
  donut.innerHTML=donutSVG(groups,total,palette,"Atendimentos");
  donut.style.background="none";
  $("statusLegend").innerHTML=groups.slice(0,6).map(([k,v],i)=>`<li><i style="background:${palette[i%palette.length]}"></i>${escapeHTML(k)} <b>${Math.round(v/total*100)}%</b></li>`).join("");
}
function renderTime(){
  const mode=$("timeMode").value, map={};
  filteredData.forEach(r=>{const d=duration(r);if(d===null)return;let k;if(mode==="weekly"){const x=new Date(r.data),day=(x.getDay()+6)%7;x.setDate(x.getDate()-day);k=dateISO(x)}else k=dateISO(r.data);if(!map[k])map[k]=[];map[k].push(d)});
  const entries=Object.entries(map).sort((a,b)=>a[0].localeCompare(b[0])).map(([k,v])=>[new Date(k+"T00:00:00"),v.reduce((a,b)=>a+b,0)/v.length]);
  const el=$("lineChart");el.innerHTML="";
  if(!entries.length){el.innerHTML='<div class="chart-empty">Sem dados de horários válidos</div>';return}
  const w=700,h=215,pad=5,max=Math.max(...entries.map(x=>x[1]),1),min=Math.min(...entries.map(x=>x[1]),0),range=Math.max(max-min,1);
  const pts=entries.map(([,v],i)=>`${i*(w-pad*2)/Math.max(entries.length-1,1)+pad},${h-pad-(v-min)/range*(h-pad*2)}`).join(" ");
  const area=pts+` ${w-pad},${h} ${pad},${h}`;
  const svg=`<svg viewBox="0 0 ${w} ${h}" preserveAspectRatio="none"><polygon class="area" points="${area}"></polygon><polyline points="${pts}"></polyline>${entries.map(([,v],i)=>{const [x,y]=pts.split(" ")[i].split(",");return `<circle cx="${x}" cy="${y}" r="3"></circle>`}).join("")}</svg><div class="time-axis">${entries.slice(0,7).map(([d])=>`<span>${dateBR(d).slice(0,5)}</span>`).join("")}</div>`;
  el.innerHTML=svg;
}
function escapeHTML(s){return String(s).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[c]));}

// Salvar PDF: captura visual do painel como ele aparece na tela.
$("pdfBtn").addEventListener("click", async ()=>{
  if(!allData.length){
    alert("Importe os dados antes de salvar o relatório em PDF.");
    return;
  }

  const btn=$("pdfBtn");
  const original=btn.textContent;
  btn.disabled=true;
  btn.textContent="Gerando PDF...";

  try{
    // Garante que o relatório esteja visível e que os gráficos tenham sido atualizados.
    $("report").classList.remove("hidden");
    await new Promise(resolve=>setTimeout(resolve,250));

    const canvas=await html2canvas(document.querySelector(".dashboard"),{
      scale:2,
      useCORS:true,
      backgroundColor:"#eef5fc",
      logging:false,
      windowWidth:document.querySelector(".dashboard").scrollWidth,
      windowHeight:document.querySelector(".dashboard").scrollHeight
    });

    const {jsPDF}=window.jspdf;
    const pdf=new jsPDF({
      orientation:"landscape",
      unit:"mm",
      format:"a4",
      compress:true
    });

    const pageW=297, pageH=210;
    const margin=6;
    const availableW=pageW-margin*2;
    const availableH=pageH-margin*2;
    const imgRatio=canvas.width/canvas.height;
    let imgW=availableW;
    let imgH=imgW/imgRatio;

    // Se o painel for muito alto, mantém a captura inteira proporcionalmente.
    if(imgH>availableH){
      imgH=availableH;
      imgW=imgH*imgRatio;
    }

    const x=(pageW-imgW)/2;
    const y=(pageH-imgH)/2;

    pdf.addImage(canvas.toDataURL("image/png"),"PNG",x,y,imgW,imgH,"FAST");
    const date=new Date();
    const stamp=`${date.getFullYear()}${String(date.getMonth()+1).padStart(2,"0")}${String(date.getDate()).padStart(2,"0")}_${String(date.getHours()).padStart(2,"0")}${String(date.getMinutes()).padStart(2,"0")}`;
    pdf.save(`relatorio_gerencial_pad_saude_${stamp}.pdf`);
  }catch(err){
    console.error(err);
    alert("Não foi possível gerar o PDF. Verifique a conexão com a internet para carregar o recurso de PDF.");
  }finally{
    btn.disabled=false;
    btn.textContent=original;
  }
});
