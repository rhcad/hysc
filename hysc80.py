import re
import json
import fire
from docx import Document, oxml
from docx.oxml.ns import qn

re_punc = re.compile('[，。：；？“”‘’「」○《》〈〉．、▲]')
re_kuo_embed = re.compile(r'「[^「」]+[『』\u3000，。；][^「」]*」')
re_kuo_dup = re.compile('「[^「」『』]{0,5}『[^「」『』]+』[等者]?」')
re_kuo_long = re.compile('「[^「」]{16,100}」')
re_kuo_sign = re.compile('[「」『』]')
re_kuo_end = re.compile('[，：；、]”$')


def fix_kuo(text):
    """修正段落文本中的括号"""
    text = re_kuo_dup.sub(lambda _: '「' + re_kuo_sign.sub('', _.group()) + '」', text.strip())
    text = re_kuo_embed.sub(fix_kuo_embed, text)
    text = re_kuo_long.sub(lambda _: _.group().replace('「', '“').replace('」', '”'), text)
    text = re_kuo_end.sub('”', re_kuo_sign.sub('', text))
    return text


def fix_kuo_embed(m):
    """修正嵌套引号的文本"""
    text = re.sub('『[^『』]{0,4}』', lambda _: _.group()[1: -1], m.group())
    return text.replace('『', '‘').replace('』', '’').replace('「', '“').replace('」', '”')


def find_kp(kp_map, juan, text, kp_idx):
    """在科判数据中查找行文本对应的科判项"""
    nodes = kp_map['id2node'][str(juan)]
    for i, v in enumerate(nodes.values()):
        if i >= kp_idx and v['text'].endswith(' ' + text):
            kp_idx = i + 1
            text = v['text']
            break
    return text, kp_idx


def set_outline_level(p, level):
    """
    设置文档段落的大纲级别
    :param p: 段落
    :param level: 大纲级别，1~9，0为正文文本
    """
    p_element = p._p
    p_pr = p_element.find(qn('w:pPr'))
    if p_pr is None:
        p_pr = oxml.OxmlElement('w:pPr')
        p_element.insert(0, p_pr)
    outline_lvl = p_pr.find(qn('w:outlineLvl'))
    if outline_lvl is not None:
        p_pr.remove(outline_lvl)
    new_outline_lvl = oxml.OxmlElement('w:outlineLvl')
    new_outline_lvl.set(qn('w:val'), str(min(8, level - 1)) if level else '-1')
    p_pr.append(new_outline_lvl)


def main():
    doc = Document('hysc80.docx')
    with open('hysc.json', encoding='utf-8') as f:
        kp_map = json.load(f)

    with open('10-80.txt', encoding='utf-8') as f:
        rows = f.read().split('\n')

    ps, last_style, juan, kp_idx, kp_idx0 = [], '', 9, 0, 0
    for i, r in enumerate(rows):
        r = fix_kuo(r)
        if not r:
            continue
        style, item = last_style, {}
        if r.startswith('大方廣佛華嚴經疏鈔會本卷第'):
            style = '卷'
            juan += 1
            kp_idx = 0
        elif not rows[i].startswith('　'):
            style = '经'
        else:
            if r.startswith('疏：'):
                style = '疏'
            elif r.startswith('鈔：'):
                style = '钞'
            elif not re_punc.search(r) or re.match('^“.{4}”善根迴向$', r):
                style = '科判'
                r, kp_idx = find_kp(kp_map, juan, rows[i].strip('　'), kp_idx)
                if not re.match(r'^\d', r):
                    if kp_idx == 0:
                        r, kp_idx0 = find_kp(kp_map, juan - 1, rows[i].strip('　'), kp_idx0)
                kp_idx0 = kp_idx
                level = re.search(r'^\d+', r)
                item['level'] = int(level.group()) if level else 40
                r = re.sub(r'^\d+:', '', re.sub(r'^\d+:…', lambda _: _.group().replace(':', ' '), r))

        last_style = style if style in '疏钞经' else '钞'
        ps.append(dict(text=r, style=style, juan=juan, **item))

    last_style = ''
    for i, r in enumerate(ps):
        text, style, juan = r['text'], r['style'], r['juan']
        if style == '卷':
            kp = [p for p in ps if p['juan'] == juan and p['style'] == '科判']
            min_level = min([p['level'] for p in kp])
            if 3 < min_level < 40:
                for p in kp:
                    p['level'] -= min_level - 4
            jing = [p for p in ps if p['juan'] == juan and p['style'] == '经']
            j, last_r = -1, ''
            while j + 1 < len(jing):
                j += 1
                t = jing[j]['text'].strip('“”')
                if re_punc.sub('', t) == t:
                    song = [j]
                    for k in range(j + 1, len(jing)):
                        t1 = jing[k]['text'].strip('“”')
                        if re_punc.sub('', t1) == t1 and len(t1) == len(t):
                            song.append(k)
                        else:
                            break
                    jing[j]['style'] = '经颂'
                    if len(song) > 1 or '　' in t:
                        for si, k in enumerate(song):
                            jing[k]['style'] = '经颂'
                            if si % 2 == 1 and '　' not in jing[k]['text'] \
                                    and '　' not in jing[k-1]['text']:
                                if ps.index(jing[k-1]) + 1 == ps.index(jing[k]):
                                    jing[k - 1]['text'] += '　　' + jing[k]['text']
                                    jing[k]['text'] = ''
                        j = song[-1]

        if style in '疏钞':
            text = re.sub(f'^{style}：', '', text)
        if style in '经疏钞' and style != last_style:
            text = f'【{style}】' + text

        p = text and doc.add_paragraph(text, style)
        if p and style == '科判':
            set_outline_level(p, r['level'])
        last_style = style

    doc.save('华严经疏钞.docx')


if __name__ == '__main__':
    fire.Fire(main)
