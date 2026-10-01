import {test} from 'node:test';
import assert from 'node:assert/strict';
import {column,exactMatches,LibraryRow} from '../frontend/src/directCore';
const entry:LibraryRow={description:'Ống thép DN25',unit:'m',brand:'',material:0,labor:18000,source:'Sheet1!8'};
test('user columns are arbitrary and validated',()=>{assert.equal(column('P'),15);assert.equal(column('Q'),16);assert.throws(()=>column(''));assert.throws(()=>column('XFE'));});
test('match ignores numbering and retains dimensions, units and brands',()=>{const row={row:17,description:entry.description,unit:'m',brand:''};assert.equal(exactMatches(row,[entry]).length,1);assert.equal(exactMatches({...row,description:'Ống thép DN32'},[entry]).length,0);assert.equal(exactMatches({...row,unit:'cái'},[entry]).length,0);assert.equal(exactMatches({...row,brand:'ABC'},[entry]).length,0);assert.equal(exactMatches(row,[entry,{...entry,labor:20000}]).length,2);});
