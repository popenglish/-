#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""中1 英語テスト 改善版を .docx で生成する"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

# ページ設定 (A4)
sec = doc.sections[0]
sec.page_width = Cm(21.0)
sec.page_height = Cm(29.7)
sec.top_margin = Cm(1.8)
sec.bottom_margin = Cm(1.8)
sec.left_margin = Cm(2.0)
sec.right_margin = Cm(2.0)

# 既定フォント
style = doc.styles['Normal']
style.font.name = 'Arial'
style.font.size = Pt(11)
style.element.rPr.rFonts.set(qn('w:eastAsia'), 'MS Gothic')
style.paragraph_format.space_after = Pt(4)


def para(text='', bold=False, size=None, align=None, space_before=None, space_after=None):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = bold
    if size:
        r.font.size = Pt(size)
    r.font.name = 'Arial'
    r.element.rPr.rFonts.set(qn('w:eastAsia'), 'MS Gothic')
    if align is not None:
        p.alignment = align
    if space_before is not None:
        p.paragraph_format.space_before = Pt(space_before)
    if space_after is not None:
        p.paragraph_format.space_after = Pt(space_after)
    return p


def set_cell_borders(cell, sz=8):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), str(sz))
        el.set(qn('w:color'), '000000')
        borders.append(el)
    tcPr.append(borders)


def goku_box(text):
    """語群を枠で囲んで表示"""
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = t.cell(0, 0)
    set_cell_borders(cell)
    cell.width = Cm(16.5)
    p = cell.paragraphs[0]
    r = p.add_run(text)
    r.font.size = Pt(11)
    r.font.name = 'Arial'
    r.element.rPr.rFonts.set(qn('w:eastAsia'), 'MS Gothic')
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para('', space_after=2)


def daimon(num, title):
    p = para(f'大問{num}　{title}', bold=True, size=12, space_before=10, space_after=4)
    # 下線
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:color'), '000000')
    pBdr.append(bottom)
    pPr.append(pBdr)


# ============ ヘッダー ============
para('中1 ②　英語テスト　2026 前期（水曜 #2 / #3）', bold=True, size=14,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
para('Name（　　　　　　　　　　　　　　　　　）　　　　得点（　　　　　／ 50 ）',
     size=12, space_after=6)
para('※ 答えはすべて 記号（きごう）を（　）に書くか、〇で囲（かこ）む問題です。'
     'ゆっくり読んで答えましょう。', size=10, space_after=8)

# ============ 大問1 ============
daimon(1, '主語（しゅご）と be動詞')
para('A）次の英語の意味をア～キから選び、（　）に記号を書きましょう。', space_after=2)
goku_box('ア あなた　　イ 彼ら（彼女ら）　　ウ 私　　エ それ　　オ 彼女　　カ 私たち　　キ 彼')
para('(1) I（　　）　(2) you（　　）　(3) he（　　）　(4) she（　　）')
para('(5) it（　　）　(6) we（　　）　(7) they（　　）', space_after=6)
para('B）次の主語に合う be動詞を選び、〇で囲みましょう。', space_after=2)
para('(1) I（ am ・ are ・ is ）　　(2) You（ am ・ are ・ is ）')
para('(3) He（ am ・ are ・ is ）　　(4) She（ am ・ are ・ is ）')
para('(5) It（ am ・ are ・ is ）　　(6) We（ am ・ are ・ is ）')
para('(7) They（ am ・ are ・ is ）')

# ============ 大問2 ============
daimon(2, 'be動詞の疑問文（ぎもんぶん）・否定文（ひていぶん）')
para('それぞれ正しい文をア～ウから選び、（　）に記号を書きましょう。', space_after=4)

be_items = [
    ('You are tall.', 'あなたは背が高いです',
     ['Are you tall?', 'You are tall?', 'Do you tall?'],
     ['You not are tall.', 'You are not tall.', "You don't tall."]),
    ('It is small.', 'それは小さいです',
     ['It is small?', 'Is it small?', 'Does it small?'],
     ['It is not small.', 'It not is small.', "It don't small."]),
    ('She is kind.', '彼女は親切です',
     ['Are she kind?', 'Is she kind?', 'She is kind?'],
     ['She is not kind.', 'She not kind.', "She doesn't kind."]),
    ('They are happy.', '彼らは幸せです',
     ['Is they happy?', 'Are they happy?', 'They are happy?'],
     ['They are not happy.', 'They not happy.', "They don't happy."]),
]
for i, (en, ja, q, n) in enumerate(be_items, 1):
    para(f'({i}) {en}（{ja}）', bold=True, space_before=4, space_after=2)
    para(f'　疑問文（　　）　ア {q[0]}　　イ {q[1]}　　ウ {q[2]}', space_after=1)
    para(f'　否定文（　　）　ア {n[0]}　　イ {n[1]}　　ウ {n[2]}', space_after=3)

# ============ 大問3 ============
daimon(3, 'can（～できる）')
para('A）それぞれ正しい文をア～ウから選び、（　）に記号を書きましょう。', space_after=4)
can_items = [
    ('You can swim.', 'あなたは泳げます',
     ['You can swim?', 'Can you swim?', 'Do you can swim?'],
     ['You cannot swim.', 'You can not swimming.', "You don't can swim."]),
    ('He can run fast.', '彼は速く走れます',
     ['Can he run fast?', 'Can he runs fast?', 'Does he can run fast?'],
     ['He cannot run fast.', "He can't runs fast.", "He doesn't can run fast."]),
    ('They can play tennis.', '彼らはテニスができます',
     ['Are they play tennis?', 'Can they play tennis?', 'Can they plays tennis?'],
     ['They cannot play tennis.', "They can't playing tennis.", "They don't can play tennis."]),
]
for i, (en, ja, q, n) in enumerate(can_items, 1):
    para(f'({i}) {en}（{ja}）', bold=True, space_before=4, space_after=2)
    para(f'　疑問文（　　）　ア {q[0]}　　イ {q[1]}　　ウ {q[2]}', space_after=1)
    para(f'　否定文（　　）　ア {n[0]}　　イ {n[1]}　　ウ {n[2]}', space_after=3)
para('B）正しい語順（ごじゅん）の文をア～ウから選び、（　）に記号を書きましょう。',
     space_before=4, space_after=2)
para('（　　）(1) 私は上手に歌えます。')
para('　ア I can sing well.　　イ I sing can well.　　ウ I well can sing.', space_after=4)
para('（　　）(2) あなたはピアノをひけますか。')
para('　ア You can play the piano?　　イ Can you play the piano?　　ウ Can you the piano play?')

# ============ 大問4 ============
daimon(4, '疑問詞（ぎもんし）（5W1H）')
para('A）次の英語の意味をア～カから選び、（　）に記号を書きましょう。', space_after=2)
goku_box('ア いつ　　イ なぜ　　ウ 何　　エ どのように　　オ どこ　　カ だれ')
para('(1) what（　　）　(2) when（　　）　(3) where（　　）')
para('(4) who（　　）　(5) why（　　）　(6) how（　　）', space_after=6)
para('B）（　）に合う疑問詞をア～ウから選び、記号を書きましょう。（be動詞の文です）', space_after=4)
q4 = [
    ('これは何ですか。', '(　　) is this?', 'What', 'Who', 'Where'),
    ('彼女はだれですか。', '(　　) is she?', 'When', 'Who', 'Why'),
    ('あなたのぼうしはどこですか。', '(　　) is your cap?', 'How', 'What', 'Where'),
    ('なぜあなたはつかれているのですか。', '(　　) are you tired?', 'Why', 'Who', 'When'),
    ('あなたのお母さんはいつ家にいますか。', '(　　) is your mother at home?', 'Who', 'When', 'Why'),
    ('あなたの学校はどうですか。', '(　　) is your school?', 'When', 'How', 'Where'),
]
for i, (ja, en, a, b, c) in enumerate(q4, 1):
    para(f'（　　）({i}) {ja}　{en}')
    para(f'　ア {a}　　イ {b}　　ウ {c}', space_after=4)

# ============ 大問5 ============
daimon(5, '時（とき）を表す語')
para('A）次の英語の意味をア～キから選び、（　）に記号を書きましょう。', space_after=2)
goku_box('ア 今日　　イ 夜　　ウ 朝　　エ 明日　　オ 午後　　カ 昨日　　キ 夕方')
para('(1) morning（　　）　(2) afternoon（　　）　(3) evening（　　）')
para('(4) night（　　）　(5) today（　　）　(6) yesterday（　　）')
para('(7) tomorrow（　　）', space_after=6)
para('B）正しい語順の文をア～ウから選び、（　）に記号を書きましょう。', space_after=4)
q5 = [
    ('私は夜に散歩をします。',
     ['I take a walk at night.', 'I at night take a walk.', 'At night a walk I take.']),
    ('私たちは毎朝バスに乗ります。',
     ['We get a bus on every morning.', 'We get on a bus every morning.',
      'Every morning on a bus we get.']),
    ('彼らは毎年夏に旅行へ行きます。',
     ['They go every summer on a trip.', 'They on a trip go every summer.',
      'They go on a trip every summer.']),
]
for i, (ja, opts) in enumerate(q5, 1):
    para(f'（　　）({i}) {ja}')
    para(f'　ア {opts[0]}　　イ {opts[1]}　　ウ {opts[2]}', space_after=4)

# ============ 大問6 ============
daimon(6, '副詞（ふくし）')
para('A）次の英語の意味をア～カから選び、（　）に記号を書きましょう。', space_after=2)
goku_box('ア 遅く　　イ 上手に　　ウ ゆっくり　　エ 早く　　オ 一生けんめいに　　カ 注意深く')
para('(1) slowly（　　）　(2) hard（　　）　(3) late（　　）')
para('(4) carefully（　　）　(5) well（　　）　(6) early（　　）', space_after=6)
para('B）正しい語順の文をア～ウから選び、（　）に記号を書きましょう。', space_after=4)
q6 = [
    ('私は学校に遅れて来ます。',
     ['I come to school late.', 'I to school late come.', 'Late come I to school.']),
    ('彼らは皿を注意深く洗います。',
     ['They wash carefully the dishes.', 'They wash the dishes carefully.',
      'They the dishes wash carefully.']),
    ('あなたはとても上手に花を植えます。',
     ['You plant flowers very well.', 'You very well plant flowers.',
      'You plant very well flowers.']),
]
for i, (ja, opts) in enumerate(q6, 1):
    para(f'（　　）({i}) {ja}')
    para(f'　ア {opts[0]}　　イ {opts[1]}　　ウ {opts[2]}', space_after=4)

# ============ 大問7 ============
daimon(7, '前置詞（ぜんちし）')
para('A）次の英語の意味をア～カから選び、（　）に記号を書きましょう。', space_after=2)
goku_box('ア ～まで　　イ ～について　　ウ ～と一緒に　　エ ～から　　オ ～へ　　カ ～のために')
para('(1) with（　　）　(2) until（　　）　(3) to（　　）')
para('(4) from（　　）　(5) for（　　）　(6) about（　　）', space_after=6)
para('B）正しい文をア～ウから選び、（　）に記号を書きましょう。', space_after=4)
q7 = [
    ('私は朝から晩まで歌います。',
     ['I sing from morning until night.', 'I sing until morning from night.',
      'I from morning sing until night.']),
    ('私は父と一緒に駅で母を待ちます。',
     ['I wait for my mother with my father at the station.',
      'I wait my mother for with my father at the station.',
      'I wait for my father with my mother at the station.']),
    ('あなたは毎日宿題についてたずねます。',
     ['You ask the homework about every day.', 'You ask about the homework every day.',
      'Every day you about the homework ask.']),
]
for i, (ja, opts) in enumerate(q7, 1):
    para(f'（　　）({i}) {ja}')
    para(f'　ア {opts[0]}', space_after=1)
    para(f'　イ {opts[1]}', space_after=1)
    para(f'　ウ {opts[2]}', space_after=4)

# ============ 大問8 ============
daimon(8, '場所（ばしょ）を表す前置詞')
para('A）次の英語の意味をア～カから選び、（　）に記号を書きましょう。', space_after=2)
goku_box('ア ～の後ろに　　イ ～の上に　　ウ ～の前に　　エ ～の下に　　オ ～のそばに　　カ ～の間に')
para('(1) in front of（　　）　(2) behind（　　）')
para('(3) between（　　）　(4) by（　　）')
para('(5) under（　　）　(6) on（　　）', space_after=6)
para('B）次の文が表す絵を、わくの中にかきましょう。（かんたんな絵でOKです）', space_after=4)

draw_items = [
    '(1) The dog is in front of the cup.',
    '(2) The pen is behind the box.',
    '(3) The cat is between the desks.',
    '(4) The apple is by the tree.',
    '(5) The ball is under the chair.',
    '(6) The cat is on the table.',
]
table = doc.add_table(rows=6, cols=2)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
for r in range(6):
    table.rows[r].height = Cm(4.0) if r % 2 == 1 else Cm(0.8)
for idx, sentence in enumerate(draw_items):
    r = (idx // 2) * 2      # 文の行
    c = idx % 2
    cell = table.cell(r, c)
    set_cell_borders(cell)
    p = cell.paragraphs[0]
    run = p.add_run(sentence)
    run.bold = True
    run.font.size = Pt(10.5)
    run.font.name = 'Arial'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), 'MS Gothic')
    # 絵をかく空きスペース
    blank = table.cell(r + 1, c)
    set_cell_borders(blank)

doc.save('/home/user/-/中1_英語テスト_2026前期_改善版.docx')
print('saved')
