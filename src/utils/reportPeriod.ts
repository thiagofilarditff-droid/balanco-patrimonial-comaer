const MONTHS=['JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ'];
const monthPattern=/\b(JAN(?:EIRO)?|FEV(?:EREIRO)?|MAR(?:CO)?|ABR(?:IL)?|MAI(?:O)?|JUN(?:HO)?|JUL(?:HO)?|AGO(?:STO)?|SET(?:EMBRO)?|OUT(?:UBRO)?|NOV(?:EMBRO)?|DEZ(?:EMBRO)?)\b/i;
const normalize=(v:unknown)=>String(v??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().replace(/\s+/g,' ').trim();
export function extractReportPeriod(rows:unknown[][]):string {
 const read=(text:string)=>{
  const quarter=text.match(/\b([1-4])\s*[º°o]?\s*TRIMESTRE\b/) || text.match(/\bTRIMESTRE\s*([1-4])\b/);
  if(quarter){const year=text.match(/\b(20\d{2})\b/)?.[1];return `${quarter[1]}° trimestre${year?'/'+year:''}${/FECHADO/.test(text)?' (fechado)':/ABERTO/.test(text)?' (aberto)':''}`;}
  const fullDate=text.match(/\b([0-3]?\d)[/.-](0?[1-9]|1[0-2])[/.-](20\d{2})\b/);
  if(fullDate){const [,d,m,y]=fullDate;const date=new Date(Date.UTC(+y,+m-1,+d));if(date.getUTCMonth()===+m-1&&date.getUTCDate()===+d)return `${d.padStart(2,'0')}/${m.padStart(2,'0')}/${y}`;}
  const match=text.match(monthPattern);
  const numeric=text.match(/^(?:(?:\d{1,2})\/)?(0?[1-9]|1[0-2])\/(20\d{2}|\d{2})\b/);
  const month=match?match[1].slice(0,3):numeric?MONTHS[Number(numeric[1])-1]:null;
  if(!month)return '';
  const year=numeric?.[2]||text.match(/(?:\/|\bDE\s+)(20\d{2}|\d{2})\b/)?.[1];
  return month+(year?`/${year.length===2?'20'+year:year}`:'')+(/FECHADO/.test(text)?' (FECHADO)':/ABERTO/.test(text)?' (ABERTO)':'');
 };
 const header=rows.slice(0,30).map(row=>row.map(normalize));
 // Only period-labelled fields, never an issue/emission date.
 for(let r=0;r<header.length;r++){
  const joined=header[r].join(' ');
  const label=joined.match(/(?:DATA[\s-]*(?:DE\s*)?(?:REFERENCIA|BASE|DO\s+BALANCO)|POSICAO(?:\s+EM)?|PERIODO(?:\s+(?:DE\s+)?REFERENCIA)?|MES(?:\s+(?:DE\s+)?(?:REFERENCIA|LANCAMENTO))?)\s*[:=\-]?\s*/);
  if(!label||label.index===undefined)continue;
  const after=joined.slice(label.index+label[0].length).split(/DATA[\s-]*(?:DE\s*)?(?:EMISSAO|EXECUCAO)|EXERCICIO/)[0].trim();
  const value=read(after);if(value)return value;
  if(!after){const next=read(header[r+1]?.join(' ')||'');if(next)return next;}
 }
 // SIAFI can omit the label and expose a standalone month/status cell.
 for(const row of header)for(const cell of row){if(/FECHADO|ABERTO/.test(cell)){const value=read(cell);if(value)return value;}}
 return '';
}
export function displayReportPeriod(period:string,year:string):string {
 if(!period)return 'Período não identificado';
 const text=normalize(period), match=text.match(monthPattern);
 const quarter=text.match(/\b([1-4])\s*[º°o]?\s*TRIMESTRE\b/i);
 if(quarter){const ownYear=text.match(/\b(20\d{2})\b/)?.[1]||year;return `${quarter[1]}° trimestre/${ownYear}${/FECHADO/.test(text)?' (fechado)':/ABERTO/.test(text)?' (aberto)':''}`;}
 if(/^\d{2}\/\d{2}\/20\d{2}$/.test(text))return text;
 if(!match)return period;
 const names=['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
 const ownYear=text.match(/\/(20\d{2})/)?.[1]||year;
 const month=MONTHS.indexOf(match[1].slice(0,3));
 const day=new Date(Date.UTC(Number(ownYear),month+1,0)).getUTCDate();
 return `${String(day).padStart(2,'0')}/${String(month+1).padStart(2,'0')}/${ownYear}`;
}
