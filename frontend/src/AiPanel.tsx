import React from 'react';
import {Item,Review,newRate} from './core';
type Props={api:(path:string,body?:any)=>Promise<any>;business:string;current?:Review;references:Item[];documents:any[];library:any[];accept:(p:Partial<Review>)=>void;message:(s:string)=>void};
export function AiPanel({api,business,current,references,documents,library,accept,message}:Props){
 const [status,setStatus]=React.useState<any>(null),[model,setModel]=React.useState('qwen3:4b'),[busy,setBusy]=React.useState(false),[docs,setDocs]=React.useState<number[]>([]),[prices,setPrices]=React.useState<number[]>([]),[preview,setPreview]=React.useState<any>(null),[proposal,setProposal]=React.useState<any>(null),[consent,setConsent]=React.useState(false),[note,setNote]=React.useState('');
 const fingerprint=JSON.stringify({business,current,references,docs,prices,documents,library});
 const latest=React.useRef(fingerprint);latest.current=fingerprint;
 React.useEffect(()=>{setPreview(null);setProposal(null);setConsent(false);},[fingerprint]);
 async function run(fn:()=>Promise<void>){setBusy(true);try{await fn();}catch(e){message(e instanceof Error?e.message:String(e));}finally{setBusy(false);}}
 async function refresh(){const s=await api('/api/ai/status');setStatus(s);setModel(s.model);}
 React.useEffect(()=>{run(refresh);},[]);
 function toggle(id:number,list:number[],set:(x:number[])=>void){set(list.includes(id)?list.filter(x=>x!==id):[...list,id]);}
 async function decide(accepted:boolean){
  if(!proposal||!current)return;
  const p:Partial<Review>={confirmed:false,status:'yellow',ruleChecks:[],amountMode:'none',evidence:'',reason:`AI ${proposal.model}: ${proposal.explanation}\n${proposal.warnings.join('\n')}\nNguồn đã dùng: ${proposal.evidenceIds.join(', ')}; cần xác minh phạm vi.`,reference:proposal.reference};
  const source=(x:any)=>`Thư viện #${x.id} v${x.version}: ${x.source}; ${x.date??''}; ${JSON.stringify(x.conditions??{})}`;
  const ref=proposal.reference;
  p.material=proposal.materialPrice?newRate(proposal.materialPrice.price,source(proposal.materialPrice)):newRate(ref?.material??null,ref?Object.values(ref.refs??{}).join('; '):'');
  p.labor=proposal.laborPrice?newRate(proposal.laborPrice.price,source(proposal.laborPrice)):newRate(ref?.labor??null,ref?Object.values(ref.refs??{}).join('; '):'');
  const stamp=fingerprint;
  await api('/api/ai/feedback',{resultId:proposal.resultId,decision:accepted?'accepted':'rejected',note});
  if(latest.current!==stamp)throw Error('Dữ liệu đã thay đổi; phản hồi đã lưu nhưng chưa áp dụng gợi ý. Phân tích lại dòng hiện tại.');
  if(accepted)accept(p);
  setProposal(null);setPreview(null);setConsent(false);
  message(accepted?'Đã đưa gợi ý AI vào bảng kiểm tra. Kiểm tra loại giá, nguồn, ĐVT và điều kiện; xác nhận từng dòng rồi xem trước để ghi Excel.':'Đã từ chối gợi ý và lưu phản hồi cục bộ.');
 }
 return <section><h2>AI cục bộ · Ollama</h2><fieldset disabled={busy}>
 <small>{status?.online?'Ollama đang chạy':'Ollama chưa kết nối'} · Hồ sơ xử lý trên máy · Không dùng cloud</small>
 <details><summary>Cài đặt & kiểm tra AI</summary><p>Chạy “Cài AI cục bộ.cmd” nếu chưa cài mô hình. AI dùng CPU; lần đầu có thể mất vài phút.</p>
 <label>Mô hình trên máy<select value={model} onChange={e=>{setModel(e.target.value);setPreview(null);setProposal(null);}}><option value="qwen3:4b">qwen3:4b</option>{(status?.models??[]).filter((m:string)=>m!=='qwen3:4b').map((m:string)=><option key={m}>{m}</option>)}</select></label>
 <button onClick={()=>run(refresh)}>Kiểm tra dịch vụ / tải lại danh sách</button>
 <button onClick={()=>run(async()=>{const r=await api('/api/ai/test',{model});message(`AI trả lời thực tế bằng ${r.model}, ${r.seconds} giây (dữ liệu kiểm thử): ${r.response.explanation}`);})}>Thử AI bằng dữ liệu minh họa</button>
 <button onClick={()=>run(async()=>{await api('/api/ai/settings',{model});await refresh();message('Đã lưu cấu hình Ollama cục bộ.');})}>Lưu cấu hình AI</button></details>
 {current?<><p>Phân tích dòng: {current.quote.id} · {current.quote.description}</p><small>AI nhận dòng hiện tại, tối đa 8 ứng viên từ vùng giá đã chọn và quy tắc nghiệp vụ. Chọn thêm nguồn dưới đây nếu cần.</small>
 <details><summary>Chọn hồ sơ bổ sung (tối đa 3)</summary>{documents.map(d=><label className="check" key={d.id}><input type="checkbox" checked={docs.includes(d.id)} onChange={()=>toggle(d.id,docs,setDocs)}/>{d.name} · v{d.version}</label>)}</details>
 <details><summary>Chọn giá thư viện (tối đa 15)</summary>{library.map(p=><label className="check" key={p.id}><input type="checkbox" checked={prices.includes(p.id)} onChange={()=>toggle(p.id,prices,setPrices)}/>{p.name} · {p.unit} · {p.price.toLocaleString('vi-VN')} · {p.source}</label>)}</details>
 <button disabled={!status?.enabled} onClick={()=>run(async()=>{setProposal(null);setConsent(false);const stamp=fingerprint;const p=await api('/api/ai/prepare',{business,quote:current.quote,references,documentIds:docs,priceIds:prices});if(latest.current!==stamp)throw Error('Dữ liệu đã thay đổi; chuẩn bị lại nội dung.');setPreview(p);})}>Chuẩn bị nội dung cho AI</button>
 {preview&&<><details open><summary>Nội dung AI sẽ xử lý · {preview.model}</summary><pre style={{maxHeight:300,overflow:'auto',whiteSpace:'pre-wrap'}}>{JSON.stringify(preview.context,null,2)}</pre></details>
 <label className="check"><input type="checkbox" checked={consent} onChange={e=>setConsent(e.target.checked)}/>Tôi xác nhận xử lý nội dung này bằng AI trên máy</label>
 <button className="primary" disabled={!consent} onClick={()=>run(async()=>{setProposal(null);const stamp=fingerprint;const p=await api('/api/ai/analyze',{planId:preview.planId,digest:preview.digest,confirmed:true});if(latest.current!==stamp)throw Error('Dòng hoặc nguồn đã thay đổi trong khi AI xử lý. Chuẩn bị lại nội dung.');setProposal(p);setPreview(null);message('AI đã trả gợi ý. Chưa thay đổi bảng kiểm tra hoặc workbook.');})}>Phân tích bằng AI cục bộ</button></>}
 {proposal&&<div className="yellow"><p>{proposal.explanation}</p>{proposal.warnings.map((w:string,i:number)=><p key={i}>{w}</p>)}<p>Nguồn: {proposal.evidenceIds.join(', ')||'Chưa đủ dữ liệu'}</p><p>Ghép dòng: {proposal.reference?.description??'Chưa đủ dữ liệu'}</p><p>Vật tư thư viện: {proposal.materialPrice?.price?.toLocaleString('vi-VN')??'Không chọn'} · Nhân công thư viện: {proposal.laborPrice?.price?.toLocaleString('vi-VN')??'Không chọn'}</p><label>Phản hồi<textarea value={note} onChange={e=>setNote(e.target.value)}/></label><button onClick={()=>run(()=>decide(true))}>Dùng gợi ý trong bảng kiểm tra</button><button onClick={()=>run(()=>decide(false))}>Từ chối và lưu phản hồi</button></div>}</>:<small>Lấy hai vùng Excel, xác nhận ánh xạ rồi bấm Bắt đầu soát xét để phân tích từng dòng với AI.</small>}
 <small>Phản hồi được lưu trên máy; chưa tự huấn luyện mô hình. Tài liệu chỉ là căn cứ. AI chưa tự phân rã giá gộp hoặc chọn cơ sở tỷ lệ.</small>
 </fieldset>{busy&&<p role="status">AI đang xử lý trên CPU… Có thể mất vài phút.</p>}</section>;
}
