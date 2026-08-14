import { useMemo, useState } from "react";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import ExcelJS from "exceljs";
import type { Recharge } from "./data";

type Occurrence = { id:number; employeeId:number; date:string; endDate?:string; type:"Falta"|"Atestado"|"Atraso"|"Aviso"; hours?:number; minutes?:number; days?:number; note?:string };
type ReportType = "Completo"|"Atrasos"|"Faltas"|"Atestados"|"Período de experiência"|"Cumprindo aviso"|"Funcionários desligados"|"Dados dos funcionários";
type ReportResult = { type: ReportType; headers:string[]; rows:{values:(string|number)[]}[] };

const types:ReportType[]=["Completo","Atrasos","Faltas","Atestados","Período de experiência","Cumprindo aviso","Funcionários desligados","Dados dos funcionários"];
const format=(value?:string)=>value?new Date(value+"T12:00:00").toLocaleDateString("pt-BR"):"-";
const addDays=(value:string,days:number)=>{const date=new Date(value+"T12:00:00");date.setDate(date.getDate()+days);return date};
const monthLabel=(value:string)=>{const label=new Date(`${value}-01T12:00:00`).toLocaleDateString("pt-BR",{month:"long",year:"numeric"});return label.charAt(0).toUpperCase()+label.slice(1)};

export default function HRReports({employees,occurrences}:{employees:Recharge[];occurrences:Occurrence[]}){
  const now=new Date(),
    [month,setMonth]=useState(`${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,"0")}`),
    stores=[...new Set(employees.map(employee=>employee.store))].sort(),
    [selectedStores,setSelectedStores]=useState<string[]>(stores),
    [selectedTypes,setSelectedTypes]=useState<ReportType[]>(["Completo"]);
  const results=useMemo<ReportResult[]>(()=>{
    const selectedEmployees=employees.filter(employee=>selectedStores.includes(employee.store)),
      active=selectedEmployees.filter(employee=>employee.terminationDate?new Date(employee.terminationDate+"T12:00:00")>now:employee.active!==false&&employee.employmentStatus!=="Desligado"),
      activeIds=new Set(active.map(employee=>employee.id)),
      map=new Map(employees.map(employee=>[employee.id,employee])),
      monthOccurrences=occurrences.filter(item=>item.date.startsWith(month)&&activeIds.has(item.employeeId));
    const build=(type:ReportType):ReportResult=>{
      if(type==="Dados dos funcionários")return {type,headers:["Nome completo","CPF","Nascimento","Admissão","Cargo / setor","Loja","Situação","Carteira","Benefício"],rows:active.map(employee=>({values:[employee.employee,employee.cpf||"-",format(employee.birthDate),format(employee.hiredAt),employee.role,employee.store,employee.employmentStatus||"Ativo",employee.formalEmployment===false?"Sem carteira":"Assinada",employee.receivesCostAssistance?"Ajuda de custo":employee.receivesTransit===false?"Sem vale-transporte":"Vale-transporte"]}))};
      if(type==="Funcionários desligados"){const list=selectedEmployees.filter(employee=>employee.terminationDate?.startsWith(month)||((employee.active===false||employee.employmentStatus==="Desligado")&&!employee.terminationDate));return {type,headers:["Nome completo","CPF","Cargo / setor","Loja","Admissão","Desligamento","Acerto até"],rows:list.map(employee=>({values:[employee.employee,employee.cpf||"-",employee.role,employee.store,format(employee.hiredAt),format(employee.terminationDate),employee.terminationDate?addDays(employee.terminationDate,10).toLocaleDateString("pt-BR"):"Não informado"]}))}}
      if(type==="Período de experiência"){const start=new Date(`${month}-01T00:00:00`),end=new Date(start.getFullYear(),start.getMonth()+1,0,23,59,59),list=active.filter(employee=>{if(!employee.hiredAt||employee.formalEmployment===false)return false;const admission=new Date(employee.hiredAt+"T12:00:00"),finish=addDays(employee.hiredAt,89);return admission<=end&&finish>=start});return {type,headers:["Funcionário","CPF","Cargo / setor","Loja","Admissão","Fim da experiência","Dias restantes"],rows:list.map(employee=>{const finish=addDays(employee.hiredAt!,89),days=Math.ceil((finish.getTime()-now.getTime())/86400000);return {values:[employee.employee,employee.cpf||"-",employee.role,employee.store,format(employee.hiredAt),finish.toLocaleDateString("pt-BR"),days>=0?days:"Encerrada"]}})}}
      if(type==="Cumprindo aviso"){
        const start=new Date(`${month}-01T00:00:00`),
          end=new Date(start.getFullYear(),start.getMonth()+1,0,23,59,59),
          selectedIds=new Set(selectedEmployees.map(employee=>employee.id)),
          employeeNotices=selectedEmployees
            .filter(employee=>employee.noticeStart&&employee.noticeEnd&&new Date(employee.noticeStart+"T12:00:00")<=end&&new Date(employee.noticeEnd+"T12:00:00")>=start)
            .map(employee=>({employee,start:employee.noticeStart!,end:employee.noticeEnd!})),
          occurrenceNotices=occurrences
            .filter(item=>item.type==="Aviso"&&item.endDate&&selectedIds.has(item.employeeId)&&new Date(item.date+"T12:00:00")<=end&&new Date(item.endDate+"T12:00:00")>=start)
            .filter(item=>!employeeNotices.some(notice=>notice.employee.id===item.employeeId))
            .map(item=>({employee:map.get(item.employeeId)!,start:item.date,end:item.endDate!}));
        return {type,headers:["Funcionário","CPF","Cargo / setor","Loja","Início do aviso","Término do aviso","Dias restantes"],rows:[...employeeNotices,...occurrenceNotices].map(notice=>{const finish=new Date(notice.end+"T12:00:00"),days=Math.ceil((finish.getTime()-now.getTime())/86400000);return {values:[notice.employee.employee,notice.employee.cpf||"-",notice.employee.role,notice.employee.store,format(notice.start),format(notice.end),days>=0?days:"Encerrado"]}})}
      }
      const wanted=type==="Atrasos"?"Atraso":type==="Faltas"?"Falta":type==="Atestados"?"Atestado":null,list=wanted?monthOccurrences.filter(item=>item.type===wanted):monthOccurrences;
      return {type,headers:["Data","Funcionário","Loja","Tipo","Detalhe"],rows:list.map(item=>{const employee=map.get(item.employeeId),detail=item.type==="Atestado"?`${item.days||1} dia(s)`:item.type==="Aviso"?`Até ${format(item.endDate)}`:item.type==="Falta"?"Falta registrada nesta data":`${item.hours||0}h ${item.minutes||0}min`;return {values:[format(item.date),employee?.employee||"Funcionário removido",employee?.store||"-",item.type,detail]}})};
    };
    return selectedTypes.map(build);
  },[employees,occurrences,month,selectedStores,selectedTypes]);
  const totalRows=results.reduce((sum,result)=>sum+result.rows.length,0),
    storesLabel=selectedStores.length===stores.length?"Todas as lojas":selectedStores.join(", "),
    toggleStore=(store:string)=>setSelectedStores(current=>current.includes(store)?current.filter(item=>item!==store):[...current,store]),
    toggleType=(type:ReportType)=>setSelectedTypes(current=>current.includes(type)?current.filter(item=>item!==type):[...current,type]);
  const exportPdf=async()=>{
    const doc=new jsPDF({orientation:"landscape"}),logo=await fetch("/sacolao-abc-logo.png?v=4").then(r=>r.blob()).then(blob=>new Promise<string>(resolve=>{const reader=new FileReader();reader.onload=()=>resolve(String(reader.result));reader.readAsDataURL(blob)}));
    doc.setFillColor(38,38,38);doc.rect(0,0,297,34,"F");doc.addImage(logo,"PNG",12,5,27,23);doc.setTextColor(255,255,255);doc.setFontSize(18);doc.text("Relatórios de Recursos Humanos",45,15);doc.setFontSize(9);doc.text(`${monthLabel(month)} | ${storesLabel} | ${totalRows} registro(s)`,45,23);
    let startY=41;
    results.forEach((result,index)=>{if(index&&startY>165){doc.addPage();startY=18}doc.setTextColor(15,23,42);doc.setFontSize(12);doc.setFont("helvetica","bold");doc.text(result.type,14,startY);autoTable(doc,{startY:startY+4,head:[result.headers],body:result.rows.map(row=>row.values),headStyles:{fillColor:[38,38,38]},alternateRowStyles:{fillColor:[244,244,244]},styles:{fontSize:8,cellPadding:2.5}});startY=((doc as any).lastAutoTable?.finalY||startY)+10});
    doc.save(`rh-relatorios-${month}.pdf`)
  };
  const exportExcel=async()=>{
    const wb=new ExcelJS.Workbook(),ws=wb.addWorksheet("Relatórios RH",{views:[{showGridLines:false}]});
    const logo=await fetch("/sacolao-abc-logo.png?v=4").then(r=>r.arrayBuffer()),id=wb.addImage({buffer:logo as never,extension:"png"});ws.addImage(id,{tl:{col:.2,row:.1},ext:{width:145,height:65}});ws.mergeCells("C1:H2");ws.getCell("C1").value="RELATÓRIOS DE RECURSOS HUMANOS";ws.getCell("C1").font={size:16,bold:true,color:{argb:"FF262626"}};ws.mergeCells("C3:H3");ws.getCell("C3").value=`${monthLabel(month)} | ${storesLabel}`;ws.addRow([]);
    results.forEach(result=>{const title=ws.addRow([result.type.toUpperCase()]);ws.mergeCells(title.number,1,title.number,Math.max(3,result.headers.length));title.getCell(1).font={bold:true,color:{argb:"FFFFFFFF"}};title.getCell(1).fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF262626"}};title.getCell(1).alignment={horizontal:"center"};const header=ws.addRow(result.headers);header.eachCell(cell=>{cell.font={bold:true,color:{argb:"FF111827"}};cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFE5E7EB"}}});result.rows.forEach(row=>ws.addRow(row.values));ws.addRow([])});
    for(let index=1;index<=Math.max(...results.map(result=>result.headers.length),3);index++)ws.getColumn(index).width=index===1?28:20;
    const bytes=await wb.xlsx.writeBuffer(),url=URL.createObjectURL(new Blob([bytes])),anchor=document.createElement("a");anchor.href=url;anchor.download=`rh-relatorios-${month}.xlsx`;anchor.click();URL.revokeObjectURL(url)
  };
  return <main className="fade-in p-4 sm:p-7">
    <div><h2 className="text-2xl font-bold text-slate-900">Relatórios de Recursos Humanos</h2><p className="mt-1 text-sm text-slate-500">Selecione um ou mais relatórios e lojas para visualizar e exportar.</p></div>
    <div className="mt-6 rounded-2xl border border-slate-200 bg-white p-5 shadow-sm"><div className="grid gap-3 md:grid-cols-3 xl:grid-cols-5">
      <details className="relative self-end"><summary className="cursor-pointer list-none rounded-xl border border-slate-200 px-3 py-2.5 text-sm font-semibold">Tipos: {selectedTypes.length===types.length?"Todos":`${selectedTypes.length} selecionado(s)`}</summary><div className="absolute z-20 mt-2 w-full min-w-[280px] rounded-xl border border-slate-200 bg-white p-3 shadow-xl"><label className="flex gap-2 border-b py-2 font-bold"><input type="checkbox" checked={selectedTypes.length===types.length} onChange={event=>setSelectedTypes(event.target.checked?types:[])}/> Todos</label>{types.map(item=><label key={item} className="flex gap-2 py-2 text-sm"><input type="checkbox" checked={selectedTypes.includes(item)} onChange={()=>toggleType(item)}/>{item}</label>)}</div></details>
      <label className="text-xs font-semibold text-slate-600">Mês<input type="month" value={month} onChange={event=>setMonth(event.target.value)} className="mt-2 w-full rounded-xl border border-slate-200 px-3 py-2.5"/></label>
      <details className="relative self-end"><summary className="cursor-pointer list-none rounded-xl border border-slate-200 px-3 py-2.5 text-sm font-semibold">Lojas: {selectedStores.length===stores.length?"Todas":`${selectedStores.length} selecionada(s)`}</summary><div className="absolute z-20 mt-2 w-full min-w-[260px] rounded-xl border border-slate-200 bg-white p-3 shadow-xl"><label className="flex gap-2 border-b py-2 font-bold"><input type="checkbox" checked={selectedStores.length===stores.length} onChange={event=>setSelectedStores(event.target.checked?stores:[])}/> Todas</label>{stores.map(item=><label key={item} className="flex gap-2 py-2 text-sm"><input type="checkbox" checked={selectedStores.includes(item)} onChange={()=>toggleStore(item)}/>{item}</label>)}</div></details>
      <button onClick={exportPdf} disabled={!results.length} className="self-end rounded-xl border border-red-300 px-4 py-2.5 text-sm font-semibold text-red-600 disabled:opacity-40">Gerar PDF</button><button onClick={exportExcel} disabled={!results.length} className="self-end rounded-xl bg-[#262626] px-4 py-2.5 text-sm font-semibold text-white disabled:opacity-40">Gerar Excel</button>
    </div></div>
    <div className="mt-5 space-y-5">{results.map(result=><div key={result.type} className="overflow-x-auto rounded-2xl border border-slate-200 bg-white shadow-sm"><div className="border-b border-slate-100 px-5 py-4"><h3 className="font-bold">{result.type}</h3><p className="text-xs text-slate-400">{result.rows.length} registro(s) em {monthLabel(month)}</p></div><table className="w-full min-w-[900px] text-left text-sm"><thead className="bg-slate-50 text-xs uppercase text-slate-400"><tr>{result.headers.map(header=><th key={header} className="px-4 py-3">{header}</th>)}</tr></thead><tbody className="divide-y divide-slate-100">{result.rows.map((row,index)=><tr key={index}>{row.values.map((value,column)=><td key={column} className="px-4 py-3">{value}</td>)}</tr>)}{!result.rows.length&&<tr><td colSpan={result.headers.length} className="py-12 text-center text-slate-400">Nenhum registro para os filtros selecionados.</td></tr>}</tbody></table></div>)}</div>
  </main>
}
