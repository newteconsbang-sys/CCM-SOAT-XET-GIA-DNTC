"""Opt-in local service test. Approves then revokes a synthetic example."""
import json
import live_ai_smoke as service
api=service.api
if __name__=='__main__':
    service.token=api('session')['token']
    q={'id':'TEST-MEMORY','description':'DỮ LIỆU KIỂM THỬ Ống PVC 20 phụ kiện','unit':'m','errors':[]}
    rate={'mode':'single','value':100,'source':'Nguồn minh họa kiểm thử'}
    m=api('ai/memories',{'business':'NCC','approved':True,'lesson':'Khi so sánh ống PVC phải kiểm tra phụ kiện đã gồm trong giá hay chưa.',
        'review':{'quote':q,'confirmed':True,'status':'yellow','material':rate,'labor':rate,'reason':'Đã kiểm tra mẫu minh họa, cần đối chiếu phạm vi phụ kiện.'}})
    try:
        assert any(x['memoryId']==m['memoryId'] for x in api('ai/memories?business=NCC'))
        p=api('ai/prepare',{'business':'NCC','quote':q,'references':[],'useMemories':True})
        examples=p['context']['approvedExamples']
        assert any(x['memoryId']==m['memoryId'] for x in examples)
        assert all('value' not in json.dumps(x) for x in examples)
        result=api('ai/analyze',{'planId':p['planId'],'digest':p['digest'],'confirmed':True})
        assert result['reference'] is None and result['materialPrice'] is None and result['laborPrice'] is None
        assert result['status']=='yellow'
        (service.ROOT/'docs/local-memory-live.json').write_text(json.dumps({'memoryId':m['memoryId'],'context':examples,'result':result},ensure_ascii=False,indent=2),encoding='utf-8')
        print('PASS: approved example retrieved by real service; local Ollama inference preserved missing current prices.')
    finally:
        api('ai/memories/revoke',{'business':'NCC','memoryId':m['memoryId']})
        assert not any(x['memoryId']==m['memoryId'] for x in api('ai/memories?business=NCC'))
