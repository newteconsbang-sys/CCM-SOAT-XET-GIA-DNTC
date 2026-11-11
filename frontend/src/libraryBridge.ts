import {LibraryRow} from './directCore';

export type LibrarySnapshot={schema:'ccm-ai-library-v1';business:'DNTC'|'EW'|'NCC'|'NTP';version:string;sha256:string;exportedAt:string;count:number;entries:(LibraryRow&{date?:string})[]};
export const bridgeHeaders=['description','brand','unit','material','labor','source','conditions','date'];
export function bridgeMatrix(s:LibrarySnapshot):(string|number)[][]{
 if(s.schema!=='ccm-ai-library-v1'||!['DNTC','EW','NCC','NTP'].includes(s.business)||!s.version||!/^[a-f0-9]{64}$/.test(s.sha256)||!Number.isFinite(Date.parse(s.exportedAt))||!Array.isArray(s.entries)||s.count!==s.entries.length||s.count<1||s.count>5000)throw Error('Bản thư viện cho ChatGPT không hợp lệ.');
 const metadata=[['CCM_AI_LIBRARY_V1','CCM-AI'],['business',s.business],['version',s.version],['sha256',s.sha256],['exportedAt',s.exportedAt],['count',s.count]];
 const data=s.entries.map(p=>{
  const expressions=[p.materialExpression,p.laborExpression].filter(Boolean);
  for(const key of ['description','source'] as const)if(typeof p[key]!=='string'||!p[key].trim())throw Error('Thư viện thiếu mô tả hoặc nguồn.');
  if(typeof p.unit!=='string'||!p.unit.trim()&&!expressions.length)throw Error('Thư viện thiếu ĐVT.');
  for(const role of ['material','labor'] as const){const expr=p[`${role}Expression`];if(expr!==undefined&&(typeof expr!=='string'||expr.length>2000||expr&&!/^(?:=?\s*\d+(?:[.,]\d+)?\s*%|(?:VT|VTP|NC)\b.*[+*/-])/i.test(expr)||expr&&p[role]!==null))throw Error('Tỷ lệ thư viện không hợp lệ.');}
  for(const key of ['material','labor'] as const)if(p[key]!==null&&(typeof p[key]!=='number'||!Number.isFinite(p[key])||p[key]!<0))throw Error('Giá thư viện không hợp lệ.');
  if(p.material===null&&p.labor===null&&!expressions.length)throw Error('Dòng thư viện thiếu cả hai giá.');
  let conditions=p.conditions??'';
  for(const [expr,label] of [[p.materialExpression,'Vật tư'],[p.laborExpression,'Nhân công']])if(expr){const note=`${label} theo tỷ lệ / biểu thức: ${expr}`;if(!conditions.includes(note))conditions=[conditions,note].filter(Boolean).join('; ');}
  return bridgeHeaders.map(key=>{const value=key==='conditions'?conditions:p[key as keyof typeof p]??'';if(typeof value==='string'&&value.length>32767)throw Error('Nội dung vượt giới hạn ô Excel.');return value;});
 });
 return [...metadata.map(r=>[...r,...Array(8-r.length).fill('')]),Array(8).fill(''),bridgeHeaders,...data];
}
export function bridgeSheetName(s:LibrarySnapshot,id:string){return `CCM_AI_${s.business}_${id.replace(/-/g,'').slice(0,12)}`;}
