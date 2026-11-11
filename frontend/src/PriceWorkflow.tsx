import React from 'react';

export function PriceIcon({kind='library'}:{kind?:string}){
 const paths:Record<string,React.ReactNode>={
  library:<><path d="M4 5c0-2 16-2 16 0v14c0 2-16 2-16 0Z"/><path d="M4 5c0 2 16 2 16 0M4 10c0 2 16 2 16 0M4 15c0 2 16 2 16 0"/></>,
  upload:<><path d="M12 16V3m-5 5 5-5 5 5M4 15v5h16v-5"/></>,
  review:<><path d="M14 3H5v18h14V8Z M14 3v5h5M8 12h8M8 16h5"/><path d="m16 18 2 2 4-5"/></>,
  book:<><path d="M12 5c-4-3-9-1-9-1v15s5-2 9 1c4-3 9-1 9-1V4s-5-2-9 1Zm0 0v15"/></>,
  arrow:<path d="m9 5 7 7-7 7"/>,
 };
 return <svg width="21" height="21" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">{paths[kind]??paths.library}</svg>;
}

export function ReviewSetup({business,version,count,onCopy}:{business:string;version:string;count:number;onCopy:(text:string)=>void}){
 const [sheet,setSheet]=React.useState(''),[description,setDescription]=React.useState(''),[unit,setUnit]=React.useState(''),[material,setMaterial]=React.useState(''),[labor,setLabor]=React.useState('');
 const fields=[description,unit,material,labor];
 const columnValid=(c:string)=>{if(!/^[A-Z]{1,3}$/.test(c))return false;return [...c].reduce((n,ch)=>n*26+ch.charCodeAt(0)-64,0)<=16384;};
 const valid=!!sheet.trim()&&fields.every(columnValid)&&new Set(fields).size===4;
 const prompt=`Dùng skill soatxetgia-ccm soát giá trong workbook Excel đang mở. Thư viện ${business}, phiên bản ${version}, ${count} đầu mục đã lưu trên web local CCM-AI. Kiểm tra kết nối Excel và đọc đủ thư viện trước khi xử lý. Sheet hồ sơ: ${sheet.trim()}; cột mô tả: ${description}; đơn vị: ${unit}; vật tư: ${material}; nhân công: ${labor}. Nếu workbook chưa có sheet thư viện, đưa bản thư viện này vào sheet nguồn riêng và đọc lại để kiểm tra. Dò theo mô tả, đơn vị và thông số; chỉ dùng giá có nguồn phù hợp duy nhất. Xem trước ô đích, giá cũ, giá đề xuất và nguồn; giữ nguyên dòng Lô, tỷ lệ/biểu thức chờ thiết lập, ô công thức và dòng chưa khớp. Chỉ ghi sau khi đã kiểm tra ô đích và có bản sao giá trị cũ. Nếu chưa kết nối được thì báo rõ, không báo đã điền.`;
 return <div className="review-setup"><div className="section-kicker">THIẾT LẬP HỒ SƠ EXCEL</div><h2>Chọn nơi điền đơn giá</h2><p>Nhập tên sheet và các cột trong file Excel của bạn để tạo lệnh soát giá cho Codex.</p><label>Sheet hồ sơ<input value={sheet} onChange={e=>setSheet(e.target.value)} placeholder="Ví dụ: BoQ" maxLength={31}/></label><div className="web-grid">{([['Cột mô tả',description,setDescription],['Cột đơn vị',unit,setUnit],['Cột đơn giá vật tư',material,setMaterial],['Cột đơn giá nhân công',labor,setLabor]] as const).map(([label,value,setValue])=><label key={label}>{label}<input value={value} onChange={e=>setValue(e.target.value.toUpperCase().trim())} placeholder="Chữ cột Excel" maxLength={3}/></label>)}</div><small>Bốn cột phải khác nhau, trong khoảng A–XFD. Không cố định cột vật tư hoặc nhân công.</small><button className="primary" disabled={!valid||!count} onClick={()=>onCopy(prompt)}>Sao chép lệnh soát giá</button><p className="review-connection">Web chưa xác minh kết nối với workbook đang mở. Dán lệnh vào Codex trên máy để kiểm tra và thực hiện qua phiên Excel được kết nối.</p></div>;
}
