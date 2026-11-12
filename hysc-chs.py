import re
import fire
from glob import glob
from docx import Document, oxml
from docx.oxml.ns import qn
import zhconv

data = {}
trans = {
    '轝': '舆', '廼': '乃', '徧': '遍', '氷': '冰', '𥧌': '寐', '㵎': '涧', '閇': '闭', '薗': '园',
    '羣': '群', '麤': '粗', '虗': '虛', '葢': '盖', '䟦': '跋', '翛': '倏', '諠': '喧',
    '鶴': '鹤', '壍': '堑', '鬭': '斗', '鬬': '斗', '濵': '滨', '叅': '参', '癎': '癫',
    '蘂': '蕊', '瓈': '璃', '荅': '答', '詺': '名', '镮': '环', '磒': '陨', '伱': '尔',
}
re_cjk_b = re.compile(r'[\U00020000-\U0002a6d6]')
re_cjk_c = re.compile(r'[\U0002a700-\U0002fa1d\U000e0100-\U000e0176]')


def convert_doc(fn, offen=None, fix_kp=False):
    doc = Document(fn)
    count = 0
    if fix_kp:
        for p in doc.paragraphs:
            if p.style.name[0] == '卷' or p.text.endswith('(终)'):
                fix_kp_level()
                data['kp_ps'] = []
            if p.style.name == '科判':
                count += 1
                data['kp_ps'].append(p)

        if count:
            print(fn)
            doc.save(fn)
        return

    dest_fn = fn.replace('繁体', '简体')
    if dest_fn == fn:
        return
    for p in doc.paragraphs:
        for r in p.runs:
            text = zhconv.convert(r.text, 'zh-hans')
            for k, v in trans.items():
                if k in text:
                    text = text.replace(k, v)

            cjk = re_cjk_b.findall(text) + re_cjk_c.findall(text)
            if cjk and 'cjk-' not in r.style.name:
                for c in set(cjk):
                    t = r.text[text.index(c)]
                    text = text.replace(c, t)
            if offen:
                text = offen.sub('', text)

            if text != r.text:
                r.text = text
                count += 1
        if not p.text:
            p.style = '空行'

    if count:
        print(dest_fn)
        doc.save(dest_fn)


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


def get_outline_level(p):
    """得到文档段落的大纲级别，0为正文文本"""
    p_element = p._p
    p_pr = p_element.find(qn('w:pPr'))
    if p_pr is None:
        return 0
    outline_lvl = p_pr.find(qn('w:outlineLvl'))
    if outline_lvl is None:
        return 0
    return int(outline_lvl.get(qn('w:val'))) + 1


def fix_kp_level():
    """修正一卷中的科判级别，让其有4级项"""
    if not data.get('kp_ps'):
        return
    off, min_level = 0, min([get_outline_level(p) for p in data['kp_ps']])
    while 4 < min_level < 40 and off < 3:
        lv4 = 0
        for p in data['kp_ps']:
            level = get_outline_level(p)
            level -= min_level + off - 4
            if level <= 4:
                lv4 += 1
                if lv4 > 4:
                    continue
            set_outline_level(p, level)
        if lv4 > 1:
            break
        off += 1
    if data.get('remove_kp5'):
        for p in data['kp_ps']:
            level = get_outline_level(p)
            if level > 6:
                set_outline_level(p, 0)


def main(doc_path='.', remove_kp5=0, fix_kp=0):
    data.update(dict(remove_kp5=remove_kp5))
    try:
        with open('offen.txt', encoding='utf-8') as f:
            offen = f.read()
        offen = ''.join(list(set(list(offen)))) + '“”·，。、：？！…»《》（）【】()○▲<>.0-9A-Za-z-'
        offen = re.compile(f'[{offen}]')
    except FileNotFoundError:
        offen = None

    for fn in sorted(glob(doc_path + '/*.docx')):
        convert_doc(fn, offen, fix_kp)


if __name__ == '__main__':
    fire.Fire(main)
