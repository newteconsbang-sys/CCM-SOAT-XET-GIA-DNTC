import * as excel from './excel';
import {autoMatch,items,newRate,suggest,Review} from './core';
// Runs only in the dedicated, explicitly labelled test workbook.
export async function acceptance(api:(path:string,body?:any)=>Promise<any>){
 const p=await Excel.run(async c=>{const p=c.workbook.properties;p.load('title');await c.sync();return p.title;});
 if(p!=='CCM-AI-ACCEPTANCE-ONLY')return;
 const tests:{name:string;passed:boolean;detail?:string}[]=[];
 const check=(name:string,ok:boolean,detail?:string)=>{tests.push({name,passed:ok,detail});if(!ok)throw Error(name+': '+detail);};
 try{
  await Excel.run(async c=>{let a=c.workbook.worksheets.getItemOrNullObject('TEST Báo giá');a.load('isNullObject');await c.sync();if(a.isNullObject)a=c.workbook.worksheets.add('TEST Báo giá');let b=c.workbook.worksheets.getItemOrNullObject('TEST Soát');b.load('isNullObject');await c.sync();if(b.isNullObject)b=c.workbook.worksheets.add('TEST Soát');a.getRange('A1:G3').values=[['STT','Hạng mục','ĐVT','Khối lượng','Đơn giá vật tư','Đơn giá nhân công','Thành tiền'],['1.10','Ống đồng Ø20','m',10,100,20,1200],['1.1','Ống đồng Ø15','m',5,80,15,475]];a.getRange('A2:A3').numberFormat=[['@'],['@']];b.getRange('A1:G3').values=[['STT','Hạng mục','ĐVT','Khối lượng','Đơn giá vật tư','Đơn giá nhân công','Thành tiền'],['1.1','Ống đồng Ø15','m',5,75,12,435],['1.10','Ống đồng Ø20','m',10,90,18,1080]];b.getRange('A2:A3').numberFormat=[['@'],['@']];b.getRange('A1:G3').select();await c.sync();});
  const r=await excel.capture();check('Office.js đọc vùng giá trị/công thức',r.values[2][4]===90);
  await Excel.run(async c=>{c.workbook.worksheets.getItem('TEST Báo giá').getRange('A1:G3').select();await c.sync();});const q=await excel.capture();
  const rs=items(r,suggest(r));const qs=items(q,suggest(q));check('Ghép hai sheet khác thứ tự và giữ STT 1.10',autoMatch(qs[0],rs)?.material===90&&qs[0].id==='1.10');
  const rows:Review[]=qs.map(x=>{const ref=autoMatch(x,rs)!;return{quote:x,reference:ref,confirmed:true,material:newRate(ref.material,'Nguồn kiểm thử'),labor:newRate(ref.labor,'Nguồn kiểm thử'),reason:'DỮ LIỆU MINH HỌA',status:'yellow',amountMode:'qtySum',evidence:''};});
  const name='TEST Kết quả '+Date.now().toString().slice(-6);const plan=await excel.makePlan(rows,[q,r],name,'A1',true,true,false);await excel.apply(plan,api);
  const result=await Excel.run(async c=>{const z=c.workbook.worksheets.getItem(name).getRange('A2:O3');z.load('values,formulas');await c.sync();return{values:z.values,formulas:z.formulas};});
  check('Office.js ghi công thức Excel thật và Excel tính lại',result.formulas[0][6].startsWith('=')&&result.values[0][10]===1080&&result.values[0][11]===120,JSON.stringify(result.values[0]));
  await excel.restore(plan);const exists=await Excel.run(async c=>{const s=c.workbook.worksheets.getItemOrNullObject(name);s.load('isNullObject');await c.sync();return!s.isNullObject;});check('Khôi phục sheet mới từ bản sao',!exists);
  await Excel.run(async c=>{const s=c.workbook.worksheets.getItem('TEST Soát');s.getRange('J1').values=[['Giữ trước ghi']];await c.sync();});
  let occupied=false;try{await excel.makePlan(rows,[q,r],'TEST Soát','J1',false,true,false);}catch{occupied=true;}check('Chặn ghi đè vùng có dữ liệu',occupied);
  const p2=await excel.makePlan(rows,[q,r],'TEST Soát','J1',false,true,true);await excel.apply(p2,api);await excel.restore(p2);const restored=await Excel.run(async c=>{const x=c.workbook.worksheets.getItem('TEST Soát').getRange('J1');x.load('values');await c.sync();return x.values[0][0];});check('Ghi đè đã xác nhận và khôi phục giá trị cũ',restored==='Giữ trước ghi');
 }catch(e){tests.push({name:'Hoàn tất kiểm thử tích hợp',passed:false,detail:String(e)});}
 await api('/api/acceptance',{host:'Excel Desktop',excelApi19:Office.context.requirements.isSetSupported('ExcelApi','1.9'),tests,finished:new Date().toISOString()});
}
