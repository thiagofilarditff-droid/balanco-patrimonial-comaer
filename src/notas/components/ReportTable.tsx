import React, {useState} from 'react';
import {Edit2} from 'lucide-react';
import {TableData, TableRow} from '../types';
import {formatDateDDMMYYYY} from '../utils/parser';

const currency=(n:number)=>{const v=Math.round(n/1000);const text=Math.abs(v).toLocaleString('pt-BR');return v<0?`(${text})`:text;};
const percent=(n:number,total=false)=>{if(total&&Math.abs(n-100)<.01)return '100';const text=Math.abs(n||0).toFixed(2).replace('.',',');return n<0?`(${text})`:text;};
export function ReportTable({data,onUpdate}:{data:TableData;onUpdate?:(updates:Partial<TableData>)=>void}){
 const [editing,setEditing]=useState(false);
 const [draft,setDraft]=useState({noteTitle:'',preCommentary:'',commentary:''});
 const edit=()=>{setDraft({noteTitle:data.noteTitle||'',preCommentary:data.preCommentary||'',commentary:data.commentary||''});setEditing(true);};
 const widths=data.onlyAV?[60,20,20]:data.hideAV&&data.hideAH?[50,25,25]:data.hideAV?[52,16,16,16]:[44,14,14,14,14];
 const row=(r:TableRow,key:string,total=false)=><tr key={key} className={total?'word-total':r.isSubtotal?'word-subtotal':''}>
  <td className={r.rubrica.toLowerCase().startsWith('ajuste')?'word-adjustment':''}>{r.rubrica}</td>
  <td>{currency(r.atual)}</td>{!data.onlyAV&&<td>{currency(r.anterior)}</td>}
  {!data.onlyAV&&!data.hideAH&&<td>{percent(r.variacao)}</td>}{!data.hideAV&&<td>{percent(r.av,total)}</td>}
 </tr>;
 return <article className="word-section">
  {onUpdate&&!editing&&<button className="word-edit no-print" onClick={edit} title="Editar textos desta seção"><Edit2 size={13}/>Editar textos</button>}
  {editing?<div className="word-editor no-print">{([['noteTitle','Título da nota'],['preCommentary','Texto introdutório'],['commentary','Comentário']] as const).map(([key,label])=><label key={key}>{label}<textarea value={draft[key]} rows={key==='noteTitle'?2:5} onChange={e=>setDraft({...draft,[key]:e.target.value})}/></label>)}<div className="flex gap-2"><button className="action" onClick={()=>{onUpdate?.(draft);setEditing(false);}}>Salvar</button><button className="action secondary" onClick={()=>setEditing(false)}>Cancelar</button></div></div>:<>{data.noteTitle&&<h3 className="word-note-title">{data.noteTitle}</h3>}{data.preCommentary&&<p className="word-prose">{data.preCommentary}</p>}</>}
  <h4 className="word-table-title">{data.title}</h4>
  <p className="word-unit">R$ milhares</p>
  <table className="word-table"><colgroup>{widths.map((w,i)=><col key={i} style={{width:`${w}%`}}/>)}</colgroup><thead><tr><th>Rubrica</th><th>{formatDateDDMMYYYY(data.currentDate||data.currentPeriod)}</th>{!data.onlyAV&&<th>{formatDateDDMMYYYY(data.previousDate||data.previousPeriod)}</th>}{!data.onlyAV&&!data.hideAH&&<th>AH (%)</th>}{!data.hideAV&&<th>AV (%)</th>}</tr></thead><tbody>{data.rows.map((r,i)=>row(r,String(i)))}{data.subtotals?.map((r,i)=>row({...r,isSubtotal:true},`sub-${i}`))}{row(data.total,'total',true)}</tbody></table>
  <p className="word-source">Fonte: SIAFI.</p>
  {!editing&&data.commentary&&<p className="word-prose">{data.commentary}</p>}
 </article>;
}
