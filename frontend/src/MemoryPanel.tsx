import React from 'react';
import {Review,validateRows} from './core';
export function MemoryPanel({api,business,current,message,changed}:{api:(p:string,b?:any)=>Promise<any>;business:string;current?:Review;message:(s:string)=>void;changed:()=>void}){
 const [memories,setMemories]=React.useState<any[]>([]),[lesson,setLesson]=React.useState(''),[approved,setApproved]=React.useState(false),[busy,setBusy]=React.useState(false);
 async function refresh(){setMemories(await api(`/api/ai/memories?business=${encodeURIComponent(business)}`));}
 async function run(fn:()=>Promise<void>){setBusy(true);try{await fn();}catch(e){message(e instanceof Error?e.message:String(e));}finally{setBusy(false);}}
 React.useEffect(()=>{setLesson('');setApproved(false);run(refresh);},[business]);
 React.useEffect(()=>{setApproved(false);},[JSON.stringify(current)]);
 return <details><summary>Ghi nhớ nghiệp vụ đã duyệt ({memories.length})</summary><fieldset disabled={busy}>
 <p>Ví dụ lưu trên máy, còn sau khi đóng ứng dụng. AI tự tìm tối đa 3 ví dụ cùng nghiệp vụ và ĐVT, có mô tả tương tự. Giá cũ không được dùng làm giá mới.</p>
 {current&&<><label>Bài học cần nhớ<textarea value={lesson} onChange={e=>{setLesson(e.target.value);setApproved(false);}} placeholder="Ví dụ: Khi so sánh ống và phụ kiện, phải kiểm tra phụ kiện đã gồm trong giá hay tách riêng."/></label>
 <label className="check"><input type="checkbox" checked={approved} onChange={e=>setApproved(e.target.checked)}/>Tôi đã kiểm tra dòng hiện tại và duyệt bài học này thành ví dụ chuẩn</label>
 <button disabled={!approved||!current.confirmed} onClick={()=>run(async()=>{validateRows([current]);await api('/api/ai/memories',{business,review:current,lesson,approved:true});setApproved(false);setLesson('');await refresh();changed();message('Đã ghi nhớ ví dụ đã duyệt trên máy. Lần phân tích sau sẽ tìm ví dụ phù hợp và hiển thị trong nội dung xem trước.');})}>Lưu ví dụ chuẩn đã duyệt</button><small>Cần xác nhận ghép dòng và nguồn giá trong bảng kiểm tra trước khi lưu.</small></>}
 {memories.map(m=><details key={m.memoryId}><summary>{m.review.quote.description} · {m.review.quote.unit}</summary><p>{m.lesson}</p><small>Duyệt: {m.updated} · Quy tắc v{m.ruleVersion}</small><button onClick={()=>run(async()=>{await api('/api/ai/memories/revoke',{business,memoryId:m.memoryId});await refresh();changed();message('Đã thu hồi ví dụ; AI không dùng cho lần phân tích mới. Lịch sử vẫn được giữ trên máy.');})}>Thu hồi ví dụ này</button></details>)}
 </fieldset></details>;
}
