"""Preserve source rates as text, never evaluate them as monetary prices."""
import re

EXPRESSION = re.compile(r'^(?:=?\s*\d+(?:[.,]\d+)?\s*%|(?:VT|VTP|NC)\b.*[+*/-])', re.I)


def expression(value):
    if not isinstance(value, str) or not EXPRESSION.match(value.strip()):
        return ''
    return value.strip()


def expression_fields(entry):
    result = {}
    for role in ('material', 'labor'):
        key = role + 'Expression'
        value = entry.get(key, '')
        if not isinstance(value, str) or len(value) > 2000 or value and not expression(value):
            raise ValueError('Tỷ lệ / biểu thức giá cần là nội dung gốc, tối đa 2.000 ký tự.')
        if value and entry.get(role) is not None:
            raise ValueError('Một cột giá chỉ có giá số hoặc tỷ lệ / biểu thức.')
        if value:
            result[key] = value.strip()
    return result


def expression_conditions(entry):
    notes = [entry.get('conditions') or '']
    for role, label in (('material', 'Vật tư'), ('labor', 'Nhân công')):
        value = entry.get(role + 'Expression')
        if value:
            note = f'{label} theo tỷ lệ / biểu thức: {value}'
            if note not in notes[0]:
                notes.append(note)
    if not entry.get('unit', '').strip() and any(entry.get(r+'Expression') for r in ('material', 'labor')):
        note = 'Nguồn chưa có ĐVT; cần xác nhận.'
        if note not in notes[0]:
            notes.append(note)
    return '; '.join(filter(None, notes))
