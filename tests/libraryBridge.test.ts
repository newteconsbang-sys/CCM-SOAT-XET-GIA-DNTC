import {test} from 'node:test';
import assert from 'node:assert/strict';
import {bridgeMatrix,bridgeSheetName,LibrarySnapshot} from '../frontend/src/libraryBridge';
const snapshot:LibrarySnapshot={schema:'ccm-ai-library-v1',business:'NTP',version:'app-1-v1',sha256:'a'.repeat(64),exportedAt:'2026-10-01T09:00:00Z',count:1,entries:[{description:'Ống DN25',brand:'',unit:'m',material:0,labor:null,source:'Giá duyệt',conditions:''}]};
test('bridge preserves typed zero and missing price, metadata and all source rows',()=>{
 const matrix=bridgeMatrix(snapshot);assert.equal(matrix[1][1],'NTP');assert.equal(matrix[8][3],0);assert.equal(matrix[8][4],'');assert.equal(matrix.length,9);assert.equal(matrix[8][5],'Giá duyệt');assert(matrix.every(r=>r.length===8));
 assert(bridgeSheetName(snapshot,'12345678-1234-1234-abcd-123456789abc').length<=31);
});
test('bridge rejects partial, empty and malformed price snapshots',()=>{
 assert.throws(()=>bridgeMatrix({...snapshot,count:2}));assert.throws(()=>bridgeMatrix({...snapshot,count:0,entries:[]}));assert.throws(()=>bridgeMatrix({...snapshot,entries:[{...snapshot.entries[0],labor:-1}]}));assert.throws(()=>bridgeMatrix({...snapshot,sha256:'invalid'}));
});
test('bridge keeps source rates in conditions, with blank monetary cells',()=>{
 const s={...snapshot,entries:[{...snapshot.entries[0],unit:'',material:null,labor:null,materialExpression:'30% (VT+NC) phần ống',laborExpression:'VT VTP/3',conditions:'Không tính thiết bị'}]};
 const row=bridgeMatrix(s)[8];assert.equal(row[3],'');assert.equal(row[4],'');assert.equal(row[2],'');
 assert.match(String(row[6]),/30% \(VT\+NC\) phần ống/);assert.match(String(row[6]),/VT VTP\/3/);assert.match(String(row[6]),/Không tính thiết bị/);
 assert.throws(()=>bridgeMatrix({...s,entries:[{...s.entries[0],material:123}]}));
});
