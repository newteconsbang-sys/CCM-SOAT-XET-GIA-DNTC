let sessionToken='';
export async function localApi(path:string,body?:unknown,responseType:'json'|'blob'='json'){
 async function session(){const r=await fetch('/api/session');if(!r.ok)throw Error('Dịch vụ local chưa sẵn sàng. Mở Khởi động CCM-AI.');sessionToken=(await r.json()).token;}
 if(!sessionToken)await session();
 const request=()=>fetch(path,{method:body===undefined?'GET':'POST',headers:{'X-CCM-Token':sessionToken,...(body instanceof FormData?{}:{'Content-Type':'application/json'})},body:body===undefined?undefined:body instanceof FormData?body:JSON.stringify(body)});
 let response=await request();if(response.status===401){await session();response=await request();}
 if(response.ok&&responseType==='blob')return response.blob();
 const result=await response.json().catch(()=>({detail:response.statusText}));
 if(!response.ok)throw Error(typeof result.detail==='string'?result.detail:JSON.stringify(result.detail));return result;
}
