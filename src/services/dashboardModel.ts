import {BalancoData,StatementItem} from '../types/balanco';
// Blank/dash mapped balances are neutral in the source; unmapped rubrics remain visible.
export const visibleBalance=(item:StatementItem)=>item.isUnmapped||(item.valCurrent??0)!==0||(item.valPrevious??0)!==0;
function children(items:StatementItem[],start:string,end:string){
 const a=items.findIndex(x=>x.id===start),b=items.findIndex(x=>x.id===end);
 return a<0?[]:items.slice(a+1,b>a?b:undefined).filter(x=>x.level===1);
}
export function dashboardModel(data:BalancoData){
 const lp=children(data.passivoPL,'pass_nao_circ','pass_exig_total');
 const others=lp.filter(x=>!['pass_emp_lp','pass_prov_lp'].includes(x.id));
 const sum=(field:'valCurrent'|'valPrevious')=>others.some(x=>x.isUnmapped)?null:others.reduce((n,x)=>n+(x[field]??0),0);
 const demais:StatementItem={id:'dashboard_demais_lp',label:'Demais',level:1,valCurrent:sum('valCurrent'),valPrevious:sum('valPrevious'),isUnmapped:others.some(x=>x.isUnmapped)};
 return {ativoCirculante:children(data.ativo,'at_circ','at_nao_circ').filter(visibleBalance),ativoNaoCirculante:children(data.ativo,'at_nao_circ','at_total').filter(visibleBalance),passivoCirculante:children(data.passivoPL,'pass_circ','pass_nao_circ').filter(visibleBalance),passivoNaoCirculante:[...lp.filter(x=>['pass_emp_lp','pass_prov_lp'].includes(x.id)),demais].filter(visibleBalance)};
}
