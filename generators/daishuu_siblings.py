"""兄弟姉妹の代襲相続が生じている場合（配偶者あり／なし共用）"""
from openpyxl import Workbook
from .common import (F14, al, set_col_widths, set_row_heights,
                     merge, fill_yellow, set_border, set_cell, draw_creator_box)


def generate(has_spouse: bool, has_half_sib: bool, siblings_data: list,
             appl_start: str = 'X24', appl_end: str = 'AA25') -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)
    _build_sheet(wb, has_spouse, has_half_sib, siblings_data, appl_start, appl_end)
    return wb


# ── ヘルパー ────────────────────────────────────────────────────

def _header(ws):
    merge(ws, 'I2', 'M2', '被相続人', F14, alignment=al(3))
    fill_yellow(ws, 2, 14, 2, 20); merge(ws, 'N2', 'T2')
    merge(ws, 'U2', 'AA2', '法定相続情報', F14, alignment=al(2))


def _creator(ws, row):
    set_cell(ws, row,   11, '作成日：', alignment=al(0))
    fill_yellow(ws, row, 14, row, 24); merge(ws, f'N{row}', f'X{row}')
    set_cell(ws, row+1, 11, '作成者：', alignment=al(0))
    set_cell(ws, row+1, 14, '住所', alignment=al(1))
    fill_yellow(ws, row+1, 16, row+1, 28); merge(ws, f'P{row+1}', f'AB{row+1}')
    set_cell(ws, row+2, 14, '氏名', alignment=al(2))
    fill_yellow(ws, row+2, 16, row+2, 22); merge(ws, f'P{row+2}', f'V{row+2}')
    draw_creator_box(ws, row - 1)


def _compute_blocks(siblings_data):
    """
    各兄弟姉妹・甥姪ブロックの開始行を計算する。

    生存兄弟姉妹:    住所/出生/続柄/氏名×2 = 5行 + gap1行 → 計6行
    死亡兄弟姉妹:    住所/出生/死亡/続柄/氏名×2 = 6行 + gap1行 → 計7行
    甥姪 (per child): 住所/出生/続柄/氏名×2 = 5行 + gap1行 → 計6行
    """
    cur = 21
    blocks = []
    for i, sib in enumerate(siblings_data):
        blk = {
            'num': i + 1,
            'start': cur,
            'alive': sib['alive'],
            'num_children': sib.get('num_children', 0),
            'nephew_starts': [],
        }
        if sib['alive']:
            cur += 6
        else:
            cur += 7
            for _ in range(sib.get('num_children', 0)):
                blk['nephew_starts'].append(cur)
                cur += 6
        blocks.append(blk)
    return blocks, cur   # cur = 以下余白 行


def _build_heights(blocks, wr, has_spouse):
    h = {2: 585, 3: 90, 9: 300, 10: 300, 11: 300,
         12: 300, 13: 300, 14: 300, 15: 300, 16: 300, 17: 300,
         18: 150, 19: 150, 20: 300}
    if has_spouse:
        h.update({5: 300, 6: 300, 7: 300, 8: 300})
    for blk in blocks:
        s = blk['start']
        if blk['alive']:
            h[s]=300; h[s+1]=300; h[s+2]=300; h[s+3]=150; h[s+4]=150; h[s+5]=300
        else:
            h[s]=300; h[s+1]=300; h[s+2]=300; h[s+3]=300
            h[s+4]=150; h[s+5]=150; h[s+6]=300
            for ns in blk['nephew_starts']:
                h[ns]=300; h[ns+1]=300; h[ns+2]=300
                h[ns+3]=150; h[ns+4]=150; h[ns+5]=300
    h[wr]=300; h[wr+2]=150; h[wr+3]=300; h[wr+4]=300; h[wr+5]=300
    return h


def _draw_sibling(ws, s, alive):
    """兄弟姉妹1ブロックを描画。name fill 開始行を返す。"""
    merge(ws, f'P{s}',   f'Q{s}',   '住所',   alignment=al(1))
    fill_yellow(ws, s,   18, s,   31); merge(ws, f'R{s}',   f'AE{s}')
    merge(ws, f'P{s+1}', f'Q{s+1}', '出生  ', alignment=al(1))
    fill_yellow(ws, s+1, 18, s+1, 27); merge(ws, f'R{s+1}', f'AA{s+1}')
    if alive:
        merge(ws, f'P{s+2}', f'Q{s+2}', '（　）', alignment=al(1))
        fill_yellow(ws, s+3, 16, s+4, 23); merge(ws, f'P{s+3}', f'W{s+4}')
        return s + 3   # name fill 開始行
    else:
        merge(ws, f'P{s+2}', f'Q{s+2}', '死亡  ', alignment=al(1))
        fill_yellow(ws, s+2, 18, s+2, 27); merge(ws, f'R{s+2}', f'AA{s+2}')
        merge(ws, f'P{s+3}', f'Q{s+3}', '（　）', alignment=al(1))
        fill_yellow(ws, s+4, 16, s+5, 23); merge(ws, f'P{s+4}', f'W{s+5}')
        return s + 4   # name fill 開始行


def _draw_nephew(ws, s):
    """甥姪1ブロックを描画。"""
    merge(ws, f'P{s}',   f'Q{s}',   '住所',   alignment=al(1))
    fill_yellow(ws, s,   18, s,   31); merge(ws, f'R{s}',   f'AE{s}')
    merge(ws, f'P{s+1}', f'Q{s+1}', '出生  ', alignment=al(1))
    fill_yellow(ws, s+1, 18, s+1, 27); merge(ws, f'R{s+1}', f'AA{s+1}')
    merge(ws, f'P{s+2}', f'Q{s+2}', '（　）', alignment=al(1))
    fill_yellow(ws, s+3, 16, s+4, 23); merge(ws, f'P{s+3}', f'W{s+4}')


def _draw_borders(ws, blocks, has_spouse):
    """関係線を描画する。

    文法（子の代襲相続と同じ考え方）:
      K列(11)幹線 : 被相続人・各兄弟姉妹を結ぶ縦線。
                    各兄弟姉妹の氏名欄中央の高さに横枝線(K-O列)。
      M列(13)小幹線: 死亡した兄弟姉妹の横枝線から下ろし、
                    各甥姪の氏名欄中央の高さに横枝線(M-O列)で接続。
                    ※甥姪はK列幹線には接続しない（兄弟姉妹の子のため）
    """
    # 配偶者の婚姻線（配偶者がいる場合のみ）
    if has_spouse:
        set_border(ws, 9,  17, right=6)
        set_border(ws, 10, 18, left=6)

    # 被相続人→兄弟姉妹幹線の接続部（行18/19境界）
    set_border(ws, 18, 11, bottom=1)
    for c in range(12, 16): set_border(ws, 18, c, bottom=1)
    set_border(ws, 19, 11, top=1)
    for c in range(12, 16): set_border(ws, 19, c, top=1)

    # 父母の婚姻線（E列二重線）
    set_border(ws, 22, 5, top=1, left=6)
    for c in range(6, 11): set_border(ws, 22, c, top=1)
    set_border(ws, 22, 10, right=1)
    set_border(ws, 20, 5, left=6); set_border(ws, 21, 5, left=6)
    set_border(ws, 23, 5, left=6)
    set_border(ws, 24, 10, right=1)

    trunk_end = 19
    for blk in blocks:
        s = blk['start']
        # この兄弟姉妹への横枝線（氏名欄中央の高さ、K-O列）
        tooth = (s + 3) if blk['alive'] else (s + 4)
        set_border(ws, tooth, 11, bottom=1)
        for c in range(12, 16): set_border(ws, tooth, c, bottom=1)
        trunk_end = max(trunk_end, tooth)

        # 甥姪：死亡兄弟姉妹の横枝線からM列の小幹線で接続
        if blk['nephew_starts']:
            sub_teeth = [ns + 3 for ns in blk['nephew_starts']]
            for r in range(tooth + 1, max(sub_teeth) + 1):
                set_border(ws, r, 13, left=1)
            for g in sub_teeth:
                set_border(ws, g, 13, bottom=1)
                set_border(ws, g, 14, bottom=1)
                set_border(ws, g, 15, bottom=1)

    # K列幹線（行19から最後の兄弟姉妹の横枝線まで連続）
    for r in range(19, trunk_end + 1):
        set_border(ws, r, 11, left=1)


# ── メイン ──────────────────────────────────────────────────────

def _build_sheet(wb, has_spouse, has_half_sib, siblings_data, appl_start='X24', appl_end='AA25'):
    sheet_name = '兄弟姉妹代襲（配偶者あり）' if has_spouse else '兄弟姉妹代襲（配偶者なし）'
    ws = wb.create_sheet(sheet_name)
    set_col_widths(ws)

    blocks, wr = _compute_blocks(siblings_data)
    set_row_heights(ws, _build_heights(blocks, wr, has_spouse))

    _header(ws)

    # 配偶者
    if has_spouse:
        set_cell(ws, 5, 16, '住所', alignment=al(1))
        fill_yellow(ws, 5, 18, 5, 31); merge(ws, 'R5', 'AE5')
        merge(ws, 'P6', 'Q6', '出生  ', alignment=al(1))
        fill_yellow(ws, 6, 18, 6, 27); merge(ws, 'R6', 'AA6')
        fill_yellow(ws, 7, 16, 7, 18); merge(ws, 'P7', 'R7', '（　　）', alignment=al(1))
        fill_yellow(ws, 8, 16, 8, 23); merge(ws, 'P8', 'W8')
        merge(ws, 'P7', 'S7', '（配偶者）', alignment=al(1))

    # 被相続人
    set_cell(ws, 11, 16, '最後の住所　', alignment=al(1))
    fill_yellow(ws, 12, 16, 12, 31); merge(ws, 'P12', 'AE12')
    set_cell(ws, 13, 16, '最後の本籍', alignment=al(1))
    merge(ws, 'P13', 'T13')
    fill_yellow(ws, 14, 16, 14, 31); merge(ws, 'P14', 'AE14')
    set_cell(ws, 15, 16, '出生  ', alignment=al(1))
    fill_yellow(ws, 15, 18, 15, 26); merge(ws, 'R15', 'Z15')
    set_cell(ws, 16, 16, '死亡  ', alignment=al(1))
    fill_yellow(ws, 16, 18, 16, 26); merge(ws, 'R16', 'Z16')
    set_cell(ws, 17, 16, '（被相続人）', alignment=al(1))

    # 父ボックス（左側 rows 18-19）
    fill_yellow(ws, 18, 16, 19, 23); merge(ws, 'P18', 'W19')

    # 母ボックス（左側 rows 24-25）
    fill_yellow(ws, 24, 16, 25, 23); merge(ws, 'P24', 'W25')
    # 申出人ラベルを選択した相続人の氏名欄右に配置
    merge(ws, appl_start, appl_end, '(申出人)', alignment=al(2))

    # 親ラベル（半血の場合は（　））
    flabel = '（　）' if has_half_sib else '（父）'
    mlabel = '（　）' if has_half_sib else '（母）'
    merge(ws, 'D22', 'E22', flabel, alignment=al(2))
    merge(ws, 'D27', 'E27', mlabel, alignment=al(2))

    # 兄弟姉妹・甥姪ブロック描画
    first_name_row = None
    for blk in blocks:
        name_row = _draw_sibling(ws, blk['start'], blk['alive'])
        if first_name_row is None:
            first_name_row = name_row
        for ns in blk['nephew_starts']:
            _draw_nephew(ws, ns)

    # 以下余白
    merge(ws, f'P{wr}', f'T{wr}', '以下余白', alignment=al(2))

    # 作成者
    _creator(ws, wr + 3)

    # 罫線
    _draw_borders(ws, blocks, has_spouse)
