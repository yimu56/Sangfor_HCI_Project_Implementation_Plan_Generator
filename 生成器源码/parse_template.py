# -*- coding: utf-8 -*-
"""
模板解析器 —— 深信服云计算平台实施方案（HCI）模板
==================================================
功能：
  1. 解析 .docx 模板的章节结构（标题层级 / 编号 / 正文 / 表格 / 图片）
  2. 提取全部“占位字段”（正文段落文本 + 表格单元格文本）
  3. 生成可供前端填充程序使用的模板模型 template_model.json

核心思路（保证 100% 保留原模板的框架、顺序与格式）：
  不重建文档，而是以原 .docx 包为基底，仅把「文本节点」替换成唯一占位符 __TK_xxx__；
  生成时把占位符换回用户输入即可。段落属性(pPr)、字符属性(rPr)、表格合并、
  页眉页脚、样式、编号、图片、页面设置全部原样保留。

用法：
  python parse_template.py [模板路径] [-o 输出目录]
"""
import base64
import json
import os
import re
import sys
import time
import zipfile
import zlib
from xml.parsers import expat

# ---------------------------------------------------------------- 常量
W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
IMAGE_TAGS = ('w:drawing', 'w:pict', 'w:object', 'w:txbxContent',
              'w:fldSimple', 'w:instrText', 'w:fldChar', 'w:AlternateContent')
TOKEN_PREFIX = '__TK_'
TOKEN_SUFFIX = '__'

# 样式名 -> (层级, 语义)。模板使用 SANGFOR 私有样式名
STYLE_MAP = {
    'SANGFOR11': (1, 'h1'),
    'SANGFOR22': (2, 'h2'),
    'SANGFOR33': (3, 'h3'),
    'SANGFOR44': (4, 'h4'),
    'SANGFOR55': (5, 'h5'),
    'SANGFORa':  (0, 'body'),
    'SANGFORb':  (0, 'body'),
    'SANGFORf':  (0, 'label'),
    'SANGFOR20': (0, 'cover'),
    'Title':     (0, 'title'),
    'ItemListCharSANGFOR': (0, 'li'),
}


# ---------------------------------------------------------------- XML 定位
TAG_RE = re.compile(
    r'<(/?)([A-Za-z_][\w:.\-]*)'
    r'((?:"[^"]*"|\'[^\']*\'|[^>"\'])*?)'
    r'(/?)>', re.S)
CDATA_RE = re.compile(r'<!\[CDATA\[.*?\]\]>', re.S)
COMMENT_RE = re.compile(r'<!--.*?-->', re.S)


class XmlIndexer:
    """自建 XML 标签扫描器：在原始 XML 文本上精确建立元素区间索引。

    只关心标签边界，不做实体解析，因此可以 100% 原样保留原始 XML 文本。
    """

    def __init__(self, text):
        self.text = text
        # 把注释/CDATA 屏蔽掉，避免其内部的 '<' 干扰
        masked = list(text)
        for m in list(COMMENT_RE.finditer(text)) + list(CDATA_RE.finditer(text)):
            for i in range(m.start(), m.end()):
                masked[i] = ' '
        masked = ''.join(masked)
        self.nodes = []          # [tag, start, end, depth, parent, self_closing]
        stack = []
        for m in TAG_RE.finditer(masked):
            closing, name, _attrs, selfclose = m.group(1), m.group(2), m.group(3), m.group(4)
            if m.group(0).startswith('<?'):
                continue
            if closing:
                if stack:
                    idx = stack.pop()
                    self.nodes[idx][2] = m.end()
                continue
            self.nodes.append([name, m.start(), m.end(), len(stack),
                               stack[-1] if stack else -1, selfclose == '/'])
            if selfclose != '/':
                stack.append(len(self.nodes) - 1)

    def children(self, tag, parent=-1):
        return [i for i, n in enumerate(self.nodes)
                if n[0] == tag and n[4] == parent and not n[5]]

    def find_first(self, tag, parent=None):
        for i, n in enumerate(self.nodes):
            if n[0] == tag and (parent is None or n[4] == parent) and not n[5]:
                return i
        return -1

    def raw(self, i):
        n = self.nodes[i]
        return self.text[n[1]:n[2]]


T_TAG = re.compile(r'<w:t(?: [^>]*)?>(.*?)</w:t>', re.S)
RPR_RE = re.compile(r'<w:r>(?:\s*)(<w:rPr>.*?</w:rPr>)?', re.S)
BM_RE = re.compile(r'<w:bookmark(?:Start|End)\b[^>]*/>')
PPR_RE = re.compile(r'<w:pPr(?:\s[^>]*)?(?:/>|>.*?</w:pPr>)', re.S)
TC_PR_RE = re.compile(r'<w:tcPr(?:\s[^>]*)?(?:/>|>.*?</w:tcPr>)', re.S)
TRPR_RE = re.compile(r'<w:trPr(?:\s[^>]*)?(?:/>|>.*?</w:trPr>)', re.S)


def deflate_raw(b):
    """raw deflate（无 zlib 头），对应浏览器 DecompressionStream('deflate-raw')"""
    c = zlib.compressobj(9, zlib.DEFLATED, -15)
    return c.compress(b) + c.flush()


def unescape(s):
    return (s.replace('&lt;', '<').replace('&gt;', '>')
             .replace('&quot;', '"').replace('&apos;', "'").replace('&amp;', '&'))


def esc(s):
    return (s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))


def para_text(raw):
    return ''.join(unescape(t) for t in T_TAG.findall(raw))


# ---------------------------------------------------------------- 解析器
class TemplateParser:
    def __init__(self, path):
        self.path = path
        self.zip = zipfile.ZipFile(path)
        self.doc = self.zip.read('word/document.xml').decode('utf-8')
        self.ix = XmlIndexer(self.doc)
        self.blocks = []      # 章节结构（人类可读）
        self.tokens = {}      # token -> 默认文本
        self._seq = 0
        self._repl = []       # (start, end, new_text)

    # -------- 工具
    def new_token(self):
        self._seq += 1
        return f'{TOKEN_PREFIX}{self._seq:04d}{TOKEN_SUFFIX}'

    def style_of(self, raw):
        m = re.search(r'<w:pStyle w:val="([^"]+)"', raw)
        return m.group(1) if m else None

    def num_of(self, raw):
        m = re.search(r'<w:numPr>.*?<w:ilvl w:val="(\d+)".*?(?:<w:numId w:val="(\d+)")?', raw, re.S)
        n = re.search(r'<w:numId w:val="(\d+)"', raw)
        l = re.search(r'<w:ilvl w:val="(\d+)"', raw)
        if n or l:
            return (n.group(1) if n else None, l.group(1) if l else None)
        return None

    def outline_of(self, raw):
        m = re.search(r'<w:outlineLvl w:val="(\d+)"', raw)
        return m.group(1) if m else None

    # -------- 段落
    def tokenize_paragraph(self, raw, kind, is_cell=False):
        """把一个段落替换成只含占位符的段落，返回 (新XML, token)"""
        if any(t in raw for t in IMAGE_TAGS):
            return raw, None
        text = para_text(raw)
        if not text and not is_cell:
            return raw, None
        ppr = PPR_RE.search(raw)
        ppr = ppr.group(0) if ppr else ''
        rpr = RPR_RE.search(raw)
        rpr = rpr.group(1) if (rpr and rpr.group(1)) else ''
        bms = ''.join(BM_RE.findall(raw))
        tok = self.new_token()
        self.tokens[tok] = text
        new = (f'<w:p>{ppr}{bms}<w:r>{rpr}'
               f'<w:t xml:space="preserve">{tok}</w:t></w:r></w:p>')
        return new, tok

    # -------- 表格单元格
    def tokenize_cell(self, cell_raw):
        if any(t in cell_raw for t in IMAGE_TAGS) or '<w:tbl' in cell_raw:
            return cell_raw, None
        m = TC_PR_RE.search(cell_raw)
        tcpr = m.group(0) if m else ''
        paras = re.split(r'(?=<w:p[ >])', cell_raw)
        texts, first_ppr, first_rpr, bms = [], '', '', ''
        for pr in paras:
            if not pr.lstrip().startswith('<w:p'):
                continue
            if not first_ppr:
                pm = PPR_RE.search(pr)
                first_ppr = pm.group(0) if pm else ''
                rm = RPR_RE.search(pr)
                first_rpr = rm.group(1) if (rm and rm.group(1)) else ''
                bms = ''.join(BM_RE.findall(pr))
            texts.append(para_text(pr))
        text = '\n'.join(t for t in texts)
        tok = self.new_token()
        self.tokens[tok] = text
        return (f'<w:tc>{tcpr}<w:p>{first_ppr}{bms}<w:r>{first_rpr}'
                f'<w:t xml:space="preserve">{tok}</w:t></w:r></w:p></w:tc>'), tok

    def tokenize_table(self, tbl_idx):
        t_start, t_end = self.ix.nodes[tbl_idx][1], self.ix.nodes[tbl_idx][2]
        trs = self.ix.children('w:tr', tbl_idx)
        cells = []
        for tr in trs:
            row = []
            for tc in self.ix.children('w:tc', tr):
                s, e = self.ix.nodes[tc][1], self.ix.nodes[tc][2]
                new, tok = self.tokenize_cell(self.doc[s:e])
                if tok:
                    self._repl.append((s, e, new))
                row.append(tok)
            cells.append(row)
        return cells

    # -------- 主流程
    def run(self):
        body = self.ix.find_first('w:body')
        if body < 0:
            raise RuntimeError('未找到 w:body')
        chapter = []          # 当前章节路径
        self.structure = []
        blocks = [i for i, n in enumerate(self.ix.nodes)
                  if n[0] in ('w:p', 'w:tbl') and n[4] == body and not n[5]]
        for child in sorted(blocks, key=lambda i: self.ix.nodes[i][1]):
            tag = self.ix.nodes[child][0]
            s, e = self.ix.nodes[child][1], self.ix.nodes[child][2]
            raw = self.doc[s:e]
            if tag == 'w:tbl':
                cells = self.tokenize_table(child)
                self.structure.append({
                    'kind': 'table',
                    'chapter': list(chapter),
                    'rows': len(cells), 'cols': max(len(r) for r in cells) if cells else 0,
                    'cells': cells,
                })
                continue
            # 段落
            style = self.style_of(raw)
            level, kind = STYLE_MAP.get(style, (0, 'body'))
            text = para_text(raw)
            has_img = any(t in raw for t in IMAGE_TAGS)
            num = self.num_of(raw)
            outline = self.outline_of(raw)
            if outline is not None and level == 0:
                level = int(outline) + 1
                kind = f'h{level}'
            if level and text.strip():
                chapter = chapter[:level - 1] + [text.strip()]
            new, tok = self.tokenize_paragraph(raw, kind)
            if tok:
                self._repl.append((s, e, new))
            self.structure.append({
                'kind': kind if not has_img else 'media',
                'level': level, 'style': style, 'num': num,
                'chapter': list(chapter), 'text': text, 'token': tok,
            })

        # 从后往前替换，避免位移
        out = self.doc
        for s, e, new in sorted(self._repl, key=lambda x: -x[0]):
            out = out[:s] + new + out[e:]
        return out

    # -------- 输出
    def build_model(self, out_dir):
        doc_xml = self.run()
        # 目录域（TOC）所在的段落未做替换，目录仍由 Word 域自动生成
        parts = {}
        for n in self.zip.infolist():
            if n.filename in ('word/document.xml',):
                continue
            parts[n.filename] = base64.b64encode(
                deflate_raw(self.zip.read(n.filename))).decode()
        model = {
            'meta': {
                'source': os.path.basename(self.path),
                'generated_at': time.strftime('%Y-%m-%d %H:%M:%S'),
                'token_count': len(self.tokens),
                'paragraph_count': sum(1 for b in self.structure if b['kind'] != 'table'),
                'table_count': sum(1 for b in self.structure if b['kind'] == 'table'),
                'structure': self.structure,
            },
            'package_parts': parts,     # 原 .docx 其余部件（deflate+base64）
            'doc_xml': base64.b64encode(deflate_raw(doc_xml.encode())).decode(),
            'tokens': self.tokens,
        }
        os.makedirs(out_dir, exist_ok=True)
        p = os.path.join(out_dir, 'template_model.json')
        with open(p, 'w', encoding='utf-8') as f:
            json.dump(model, f, ensure_ascii=False, separators=(',', ':'))
        print(f'[OK] 模板模型已生成: {p}  ({os.path.getsize(p)/1024:.0f} KB)')
        print(f'     占位字段 {len(self.tokens)} 个 | 段落 {model["meta"]["paragraph_count"]} '
              f'| 表格 {model["meta"]["table_count"]}')
        # 人类可读的章节结构
        lines = []
        for b in self.structure:
            if b['kind'] == 'table':
                lines.append(f'{"  "*len(b["chapter"])}[表格 {b["rows"]}行×{b["cols"]}列]')
            elif b['kind'] == 'media':
                lines.append(f'{"  "*len(b["chapter"])}[第{b["level"] or 0}级/图形]')
            elif b['level']:
                lines.append(f'{b["level"]*"#"*4} {b["text"]}')
            elif b['text'].strip():
                lines.append(f'{"  "*len(b["chapter"])}{b["text"][:80]}')
        with open(os.path.join(out_dir, 'template_outline.txt'), 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines))
        with open(os.path.join(out_dir, 'template_fields.json'), 'w', encoding='utf-8') as f:
            json.dump(self.tokens, f, ensure_ascii=False, indent=1)
        return model


if __name__ == '__main__':
    argv = sys.argv[1:]
    # 目录可整体搬迁：源码在 <项目根>/生成器源码/，模板 docx 在 <项目根>/
    BUILD = os.path.dirname(os.path.abspath(__file__))
    ROOT = os.path.dirname(BUILD)
    tpl = argv[0] if argv and not argv[0].startswith('-') else \
        os.path.join(ROOT, '深信服云计算平台实施方案-XX集团.docx')
    out = BUILD
    if '-o' in argv:
        out = argv[argv.index('-o') + 1]
    TemplateParser(tpl).build_model(out)
