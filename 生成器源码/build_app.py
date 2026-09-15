# -*- coding: utf-8 -*-
"""
生成本地 HTML 应用：把模板模型 + 表单结构 + 原始文档包 注入 app_template.html
输出：深信服云计算平台实施方案生成器.html（单文件，离线可用）
"""
import base64, json, os, re, sys, zipfile, zlib


def deflate_raw(b):
    c = zlib.compressobj(9, zlib.DEFLATED, -15)
    return c.compress(b) + c.flush()

# 目录可整体搬迁：源码在 <项目根>/生成器源码/，产物与模板 docx 在 <项目根>/
BUILD = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BUILD)
SRC_DOCX = os.path.join(ROOT, '深信服云计算平台实施方案-XX集团.docx')
OUT = os.path.join(ROOT, '深信服云计算平台实施方案生成器.html')

# ------------------------------------------------------------------ 参数定义
# (key, 分组, 标签, 默认值, 类型, 提示)
PARAMS = [
    ('customer',      '基本信息', '客户名称',            'XX集团',                'text',   '会自动替换全文中的“XX集团”'),
    ('projectTitle',  '基本信息', '项目名称',            'XX集团超融合实施方案',   'text',   '用于封面标题与文件名'),
    ('docVersion',    '基本信息', '文档版本',            '01',                    'text',   ''),
    ('docDate',       '基本信息', '发布日期',            '2026-09-14',            'date',   ''),
    ('clusterName',   '平台规模', '集群名称',            'HCI-Cluster',           'text',   ''),
    ('nodes',         '平台规模', '节点数量（台）',       '6',                     'number', '影响项目背景/目标/范围与 IP 规划'),
    ('serverModel',   '平台规模', '服务器型号',          'aServer-X5-2305P',      'text',   ''),
    ('mgmtPrefix',    '网络规划', '管理网段前缀',        '192.168.131.',          'text',   '含结尾的点，如 192.168.131.'),
    ('mgmtStart',     '网络规划', '管理 IP 主机位起始',   '173',                   'number', '第 1 台服务器管理口末位'),
    ('bmcStart',      '网络规划', 'MGMT/BMC 主机位起始',  '181',                   'number', '第 1 台服务器带外管理口末位'),
    ('mgmtGateway',   '网络规划', '管理网关',            '192.168.131.254',       'text',   ''),
    ('scpIp',         '网络规划', 'SCP 管理 IP',         '192.168.131.180',       'text',   ''),
    ('aniIp',         '网络规划', 'aNI 管理 IP',         '192.168.131.187',       'text',   ''),
    ('brainIp',       '网络规划', '云脑管理 IP',         '192.168.131.188',       'text',   ''),
    ('storageNet',    '网络规划', '存储网段',            '10.251.251.X',          'text',   ''),
    ('startDate',     '项目组',   '项目启动日期',        '2026-09-15',            'date',   '用于推算实施周期完成时间'),
    ('custPM',        '项目组',   '甲方项目经理',        '吴XX',                  'text',   ''),
    ('vendPM',        '项目组',   '乙方项目经理',        '包XX',                  'text',   ''),
    ('engDeliver',    '项目组',   '交付工程师',          '鲍X',                   'text',   ''),
    ('engExpert',     '项目组',   '交付专家',            '王XX',                  'text',   ''),
    ('salesMgr',      '项目组',   '销售经理',            '刘X',                   'text',   ''),
    ('presalesMgr',   '项目组',   '售前经理',            '薛X',                   'text',   ''),
]

# ------------------------------------------------------------------ 文本联动规则
LITERAL_RULES = [
    ('XX集团超融合实施方案', '{projectTitle}'),
    ('文档版本  01', '文档版本  {docVersion}'),
    ('XX集团', '{customer}'),
    ('2026-09-14', '{docDate}'),
    ('HCI-Cluster', '{clusterName}'),
    ('aServer-X5-2305P', '{serverModel}'),
    ('192.168.131.254', '{mgmtGateway}'),
    ('192.168.131.180', '{scpIp}'),
    ('192.168.131.188', '{brainIp}'),
    ('192.168.131.187', '{aniIp}'),
    ('新建一套6节点的私有云', '新建一套{nodes}节点的私有云'),
    ('建设一套6节点的超融合私有云', '建设一套{nodes}节点的超融合私有云'),
    ('本次需管理网IP共计12个', '本次需管理网IP共计{nodes*2}个'),
    ('本次需存储网IP共计6个', '本次需存储网IP共计{nodes}个'),
    ('vxlan接口IP为6个', 'vxlan接口IP为{nodes}个'),
]

# ------------------------------------------------------------------ 表格元信息
TABLE_META = {
    1:  ('修订记录', 1, 1, ''),
    2:  ('符号说明 · 图形标志', 1, 0, ''),
    3:  ('符号说明 · 界面格式', 1, 0, ''),
    4:  ('物料清单 · 硬件设备', 1, 1, ''),
    5:  ('物料清单 · 线材', 1, 1, ''),
    6:  ('设备规格参数', 1, 1, ''),
    7:  ('设备命名规范', 1, 1, ''),
    8:  ('机架部署规划（U 位图）', 2, 0, ''),
    9:  ('网络平面说明', 1, 1, ''),
    10: ('服务器 IP 规划', 1, 1, 'ip'),
    11: ('设备接线表', 1, 1, ''),
    12: ('平台资源规划', 2, 0, ''),
    13: ('备份容量规划', 1, 1, ''),
    14: ('平台组件 IP 规划', 1, 1, ''),
    15: ('SCP 配置推荐', 1, 1, ''),
    16: ('云主机最佳实践 · 计算', 1, 1, ''),
    17: ('云主机最佳实践 · 内存/磁盘/网络', 1, 1, ''),
    18: ('业务迁移方式', 1, 1, ''),
    19: ('项目组成员', 2, 1, ''),
    20: ('项目实施周期', 1, 1, 'sch'),
    21: ('实施前准备 · 工具材料', 1, 1, ''),
    22: ('平台可靠性测试项', 1, 1, ''),
    23: ('平台功能及性能测试项', 1, 1, ''),
    24: ('产品培训计划', 1, 1, ''),
    25: ('项目验收材料', 1, 1, ''),
}

# 实施周期默认工期（天）
SCHEDULE_DAYS = ['3', '5', '5', '5', '10', '7', '2', '3']


def build():
    model = json.load(open(os.path.join(BUILD, 'template_model.json'), encoding='utf-8'))
    tokens = model['tokens']
    struct = model['meta']['structure']

    # ---------- 1. 表格（按文档顺序编号） ----------
    table_list = []      # table_list[表号][行号][列号] -> token
    for blk in struct:
        if blk['kind'] == 'table':
            table_list.append(blk['cells'])
    if len(table_list) != 25:
        print('[WARN] 模板共 %d 张表格，与 TABLE_META 预设的 25 张不一致；'
              '超出部分会用默认标题，请同步 TABLE_META 与下面的表号。' % len(table_list))

    def tc(ti, ri, ci):
        """按 表号/行号/列号 取单元格 token；换模板导致表号错位时返回 None 而不是崩溃"""
        try:
            return table_list[ti][ri][ci]
        except (IndexError, TypeError):
            return None

    # ---------- 2. 公式 ----------
    formulas = {}

    def add_formula(tok, tpl):
        if tok and tok in tokens:
            formulas[tok] = tpl

    for tok, txt in tokens.items():
        new = txt
        for a, b in LITERAL_RULES:
            if a in new:
                new = new.replace(a, b)
        if new != txt:
            formulas[tok] = new

    # 表 1 修订记录 · 版本号
    add_formula(tc(0, 1, 1), '{docVersion}')
    # 表 4 物料清单 · 数量
    add_formula(tc(3, 1, 1), '{nodes}')
    # 表 19 项目组成员 · 姓名（第 2~7 行第 2 列）
    member_map = {2: '{custPM}', 3: '{vendPM}', 4: '{engDeliver}',
                  5: '{engExpert}', 6: '{salesMgr}', 7: '{presalesMgr}'}
    for r, tpl in member_map.items():
        add_formula(tc(18, r, 1), tpl)
    # 表 10 服务器 IP 规划（每台服务器占 5 行：管理口/vxlan/存储口/业务口/MGMT口）
    for k in range(6):
        base = 1 + k * 5
        add_formula(tc(9, base, 0), '超融合服务器%d' % (k + 1))
        add_formula(tc(9, base, 3), '{mgmtPrefix}{mgmtStart+' + str(k) + '}')
        add_formula(tc(9, base, 5), '{mgmtGateway}')
        add_formula(tc(9, base + 4, 3), '{mgmtPrefix}{bmcStart+' + str(k) + '}')
    # 表 20 项目实施周期：预填各阶段工期
    for i, d in enumerate(SCHEDULE_DAYS):
        tok = tc(19, i + 1, 2)
        if tok:
            tokens[tok] = d

    # ---------- 3. 表单结构 ----------
    sections = []
    cur = None
    prev_num = None
    list_item = None
    n1 = n2 = 0
    cover_done = False

    def new_section(title, level, adv=0, note=''):
        s = {'title': title, 'level': level, 'adv': adv, 'note': note, 'items': []}
        sections.append(s)
        return s

    cur = new_section('封面与文档信息', 0, adv=1,
                      note='封面标题、文档版本、发布时间由左侧“项目关键参数”自动填充。')
    ti = 0

    for b in struct:
        if b['kind'] == 'table':
            ti += 1
            title, nhead, repeat, gen = TABLE_META.get(ti, ('表格 %d' % ti, 1, 1, ''))
            cur['items'].append({'kind': 'table', 'ti': ti, 'title': title,
                                 'rows': b['cells'], 'nhead': nhead,
                                 'repeat': repeat, 'gen': gen})
            prev_num = None
            list_item = None
            continue
        if b['kind'] == 'media':
            cur['items'].append({'kind': 'media',
                                 'note': '模板内置图形（网络拓扑图 / 部署流程图），生成后原样保留。'})
            prev_num = None
            list_item = None
            continue
        tok = b.get('token')
        if not tok:
            continue
        lvl = b.get('level') or 0
        if lvl:
            if lvl == 1:
                n1 += 1
                n2 = 0
                cur = new_section('%d. %s' % (n1, b['text']), 1)
            elif lvl == 2:
                n2 += 1
                cur = new_section('%d.%d %s' % (n1, n2, b['text']), 2)
            cur['items'].append({'kind': 'h', 'tok': tok, 'level': lvl})
            prev_num = None
            list_item = None
            # 封面三个基础字段由参数控制，不再在正文里重复出现
            continue
        # 封面参数字段跳过
        if tok in ('__TK_0001__', '__TK_0002__', '__TK_0003__'):
            continue
        num = b.get('num')
        nid = num[0] if num else None
        if nid:
            if list_item and prev_num == nid:
                list_item['toks'].append(tok)
            else:
                list_item = {'kind': 'list', 'toks': [tok],
                             'style': 'bullet' if nid in ('18', '22', '4') else 'number'}
                cur['items'].append(list_item)
            prev_num = nid
            continue
        cur['items'].append({'kind': 'p', 'tok': tok})
        prev_num = None
        list_item = None

    # ---------- 4. 覆盖 settings.xml / core.xml ----------
    z = zipfile.ZipFile(SRC_DOCX)
    settings = z.read('word/settings.xml').decode('utf-8')
    if 'updateFields' not in settings:
        settings = re.sub(r'(<w:settings[^>]*>)', r'\1<w:updateFields w:val="true"/>',
                          settings, count=1)
    core = z.read('docProps/core.xml').decode('utf-8')
    def set_tag(xml, tag, val):
        if re.search(r'<%s[^>]*>.*?</%s>' % (tag, tag), xml):
            return re.sub(r'<%s[^>]*>.*?</%s>' % (tag, tag), '<%s>%s</%s>' % (tag, val, tag), xml)
        return xml.replace('</cp:coreProperties>', '<%s>%s</%s></cp:coreProperties>' % (tag, val, tag))
    core = set_tag(core, 'dc:title', '@@TITLE@@')
    core = set_tag(core, 'dc:creator', '@@AUTHOR@@')
    core = set_tag(core, 'cp:lastModifiedBy', '@@AUTHOR@@')
    core = re.sub(r'<dcterms:modified[^>]*>.*?</dcterms:modified>',
                  '<dcterms:modified xsi:type="dcterms:W3CDTF">@@DATE@@</dcterms:modified>', core)

    overrides = {
        'word/settings.xml': base64.b64encode(deflate_raw(settings.encode())).decode(),
        'docProps/core.xml': base64.b64encode(deflate_raw(core.encode())).decode(),
    }

    # ---------- 5. 组装 appData ----------
    app = {
        'meta': {'source': model['meta']['source'],
                 'token_count': len(tokens),
                 'paragraph_count': model['meta']['paragraph_count'],
                 'table_count': model['meta']['table_count']},
        'params': [{'k': k, 'grp': g, 'label': l, 'val': v, 'type': t, 'hint': h}
                   for (k, g, l, v, t, h) in PARAMS],
        'formulas': formulas,
        'tokens': tokens,
        'sections': sections,
    }
    stats = {'sections': len(sections),
             'visible_tokens': len(tokens) - len(formulas),
             'formulas': len(formulas)}

    tpl = open(os.path.join(BUILD, 'app_template.html'), encoding='utf-8').read()
    pkg_b64 = base64.b64encode(open(SRC_DOCX, 'rb').read()).decode()
    html = (tpl
            .replace('__APP_DATA_JSON__', json.dumps(app, ensure_ascii=False, separators=(',', ':')))
            .replace('__PKG_B64__', pkg_b64)
            .replace('__DOCXML_B64__', model['doc_xml'])
            .replace('__OVERRIDES_JSON__', json.dumps(overrides, ensure_ascii=False)))
    with open(OUT, 'w', encoding='utf-8') as f:
        f.write(html)
    print('[OK] 生成 %s  (%.2f MB)' % (OUT, os.path.getsize(OUT) / 1048576))
    print('     %s' % stats)
    for s in sections:
        kinds = {}
        for i in s['items']:
            kinds[i['kind']] = kinds.get(i['kind'], 0) + 1
        print('   -', s['title'], kinds)


if __name__ == '__main__':
    build()
