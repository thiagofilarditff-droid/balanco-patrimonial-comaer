import assert from 'node:assert/strict';
import {extractReportPeriod,displayReportPeriod} from '../src/utils/reportPeriod';
import {dashboardModel,visibleBalance} from '../src/services/dashboardModel';
import {getSampleCOMAERData} from '../src/services/sampleData';
import {parseBalancoExcel} from '../src/services/excelParser';
import {generateSampleExcelWorkbook} from '../src/services/excelGenerator';
import * as XLSX from 'xlsx';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
assert.equal(extractReportPeriod([['Período','SET (FECHADO)']]),'SET (FECHADO)');
assert.equal(extractReportPeriod([['Mês de referência: set/26']]),'SET/2026');
assert.equal(extractReportPeriod([['Período:'],['SETEMBRO (FECHADO)']]),'SET (FECHADO)');
assert.equal(extractReportPeriod([['Data de Emissão: 22/04/2026']]),'');
assert.equal(extractReportPeriod([['Período: 30/09/2026']]),'30/09/2026');
assert.equal(displayReportPeriod('SET (FECHADO)','2026'),'30/09/2026');
const d=getSampleCOMAERData();
for(const item of d.passivoPL){if(item.id==='pass_trab_lp'){item.valCurrent=100;item.valPrevious=50;} if(item.id==='pass_fisc_lp'){item.valCurrent=200;item.valPrevious=25;}}
const m=dashboardModel(d), rest=d.passivoPL.filter(x=>['pass_trab_lp','pass_fornec_lp','pass_fisc_lp','pass_transf_fisc_lp','pass_demais_lp','pass_res_dif'].includes(x.id));
assert.equal(m.passivoNaoCirculante.find(x=>x.id==='dashboard_demais_lp')!.valCurrent,rest.reduce((a,x)=>a+(x.valCurrent??0),0));
assert(m.passivoNaoCirculante.every(x=>['pass_emp_lp','pass_prov_lp','dashboard_demais_lp'].includes(x.id)));
assert(!visibleBalance({id:'z',label:'zero',level:1,valCurrent:0,valPrevious:0}));
assert(visibleBalance({id:'z',label:'zero',level:1,valCurrent:0,valPrevious:10}));
assert(visibleBalance({id:'z',label:'missing',level:1,valCurrent:null,valPrevious:null,isUnmapped:true}));
const bytes=generateSampleExcelWorkbook();const wb=XLSX.read(bytes,{type:'array'});const ws=wb.Sheets[wb.SheetNames[0]];
for(const [address,cell] of Object.entries(ws)){if(!address.startsWith('!')&&typeof cell.v==='string')cell.v=cell.v.replace(/MAR \(FECHADO\)/g,'SET (FECHADO)');}
const out=XLSX.write(wb,{type:'buffer',bookType:'xlsx'});const parsed=await parseBalancoExcel(out.buffer.slice(out.byteOffset,out.byteOffset+out.byteLength));
assert.equal(parsed.data.metadata.periodo,'SET (FECHADO)');
fs.writeFileSync(path.join(os.tmpdir(),'dashboard-test-setembro.xlsx'),out);
console.log('OK: período setembro, rótulos separados, ausência de fallback, agrupamento LP e saldos zero.');

assert.equal(displayReportPeriod('FEV/2024','2026'),'29/02/2024');
assert.equal(extractReportPeriod([['Data-base', '15/09/2026']]),'15/09/2026');

assert.equal(extractReportPeriod([['EXERCÍCIO','','2026'],['PERÍODO','','3º TRIMESTRE (FECHADO)']]),'3° trimestre (fechado)');
assert.equal(displayReportPeriod(extractReportPeriod([['PERÍODO','3º TRIMESTRE (FECHADO)']]),'2026'),'3° trimestre/2026 (fechado)');
assert.equal(displayReportPeriod('1º TRIMESTRE (ABERTO)','2027'),'1° trimestre/2027 (aberto)');
