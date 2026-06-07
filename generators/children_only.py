"""子（１人～４人まで対応）である場合（配偶者なし）"""
from openpyxl import Workbook
from .common import (F11, F14, YEL, al, set_col_widths, set_row_heights,
                     merge, fill_yellow, set_border, set_cell, draw_creator_box)


def generate(num_children: int,
             appl_start: str = 'Y13', appl_end: str = 'AB14') -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)
    _build_sheet(wb, num_children, appl_start, appl_end)
    return wb


def _build_sheet(wb, n, appl_start='Y13', appl_end='AB14'):
    ws = wb.create_sheet(f'子{n}人の場合')
    set_col_widths(ws)
    if   n == 1: _s1(ws)
    elif n == 2: _s2(ws)
    elif n == 3: _s3(ws)
    elif n == 4: _s4(ws)
    else:        _sN(ws, n)
    # 申出人ラベルを選択した相続人の氏名欄右に配置
    merge(ws, appl_start, appl_end, '(申出人)', alignment=al(2))


# ── 拡張ヘルパー（子5人以上） ──────────────────────────────────
def _co_child_start(n):
    """子n（n≥3）の開始行: 32 + (n-3)*6"""
    return 32 + (n - 3) * 6


def _co_add_child_block(ws, start):
    """子3以降の5行ブロック"""
    merge(ws, f'P{start}',   f'Q{start}',   '住所', alignment=al(1))
    fill_yellow(ws, start,   18, start,   31); merge(ws, f'R{start}',   f'AE{start}')
    merge(ws, f'P{start+1}', f'Q{start+1}', '出生', alignment=al(1))
    fill_yellow(ws, start+1, 18, start+1, 25); merge(ws, f'R{start+1}', f'Y{start+1}')
    fill_yellow(ws, start+2, 16, start+2, 18)
    merge(ws, f'P{start+2}', f'R{start+2}', '（　　）', alignment=al(1))
    fill_yellow(ws, start+3, 16, start+4, 24); merge(ws, f'P{start+3}', f'X{start+4}')


def _co_add_m_borders(ws, start):
    """M列区切り罫線（子3以降のセクション境界）"""
    set_border(ws, start-2, 13, top=1, left=1)
    set_border(ws, start-2, 14, top=1); set_border(ws, start-2, 15, top=1)
    for r in range(start-1, start+3):
        set_border(ws, r, 13, left=1)
    set_border(ws, start+3, 13, bottom=1, left=1)
    set_border(ws, start+3, 14, bottom=1); set_border(ws, start+3, 15, bottom=1)


# ── 共通ヘッダー ──────────────────────────────
def _header(ws):
    merge(ws,'I4','M4','被相続人',F14,alignment=al(3))
    fill_yellow(ws,4,14,4,20); merge(ws,'N4','T4')
    merge(ws,'U4','AA4','法定相続情報',F14,alignment=al(2))


def _creator(ws, row):
    """row = 作成日の行"""
    set_cell(ws, row, 11, '作成日：', alignment=al(0))
    fill_yellow(ws, row, 14, row, 24); merge(ws, f'N{row}', f'X{row}')
    set_cell(ws, row+1, 11, '作成者：', alignment=al(0))
    merge(ws, f'N{row+1}', f'O{row+1}', '住所', alignment=al(1))
    fill_yellow(ws, row+1, 16, row+1, 28); merge(ws, f'P{row+1}', f'AB{row+1}')
    set_cell(ws, row+2, 14, '氏名', alignment=al(2))
    fill_yellow(ws, row+2, 16, row+2, 22); merge(ws, f'P{row+2}', f'V{row+2}')
    draw_creator_box(ws, row - 1)


# ── Sheet 1 ───────────────────────────────────
def _s1(ws):
    set_row_heights(ws, {4:585,5:90,6:300,7:300,8:300,9:300,10:300,11:300,12:300,
        13:150,14:150,15:300,16:150,17:150,18:150,19:150,20:300,21:300,
        22:150,23:150,24:300,26:150,27:300,28:300,29:300,30:300,
        32:300,33:300,34:300,35:300,36:300,37:300,38:300,39:300})
    _header(ws)
    merge(ws,'A6','E6','最後の住所　',alignment=al(1))
    fill_yellow(ws,7,1,7,13); merge(ws,'A7','M7')
    merge(ws,'A8','E8','最後の本籍',alignment=al(1))
    fill_yellow(ws,9,1,9,13); merge(ws,'A9','M9')
    merge(ws,'A10','B10','出生  ',alignment=al(1))
    fill_yellow(ws,10,3,10,11); merge(ws,'C10','K10')
    set_cell(ws,10,16,'住所',alignment=al(1))
    fill_yellow(ws,10,18,10,31); merge(ws,'R10','AE10')
    set_cell(ws,11,1,'死亡  ',alignment=al(1))
    fill_yellow(ws,11,3,11,11); merge(ws,'C11','K11')
    merge(ws,'P11','Q11','出生  ',alignment=al(1))
    fill_yellow(ws,11,18,11,26); merge(ws,'R11','Z11')
    merge(ws,'A12','E12','（被相続人）',alignment=al(2))
    fill_yellow(ws,12,16,12,18); merge(ws,'P12','R12','（　　）',alignment=al(1))
    fill_yellow(ws,13,1,14,8); merge(ws,'A13','H14')
    fill_yellow(ws,13,16,14,24); merge(ws,'P13','X14')

    merge(ws,'P16','T17','以下余白',alignment=al(2))
    fill_yellow(ws,18,16,19,24); merge(ws,'P18','X19')
    merge(ws,'P20','Q20')
    fill_yellow(ws,20,18,20,27); merge(ws,'R20','AA20')
    fill_yellow(ws,22,16,23,24); merge(ws,'P22','X23')
    _creator(ws, 27)
    for c in range(9,15): set_border(ws,14,c,top=1)


# ── Sheet 2 ───────────────────────────────────
def _s2(ws):
    set_row_heights(ws, {4:585,5:90,6:240,7:300,8:300,9:300,10:300,
        11:150,12:150,13:150,14:150,15:300,16:150,17:150,18:150,19:150,
        20:150,21:150,22:300,23:300,24:150,25:150,26:300,27:300,28:300,
        29:150,30:300,31:300,32:300,33:300,35:300,36:300,37:300,38:300,
        39:300,40:300,41:300,42:300})
    _header(ws)
    merge(ws,'A7','K7','最後の住所',alignment=al(1))
    fill_yellow(ws,8,1,8,13); merge(ws,'A8','M8')
    merge(ws,'A9','K9','最後の本籍',alignment=al(1))
    merge(ws,'P9','Q9','住所',alignment=al(1))
    fill_yellow(ws,9,18,9,31); merge(ws,'R9','AE9')
    fill_yellow(ws,10,1,10,13); merge(ws,'A10','M10')
    merge(ws,'P10','Q10','出生  ',alignment=al(1))
    fill_yellow(ws,10,18,10,27); merge(ws,'R10','AA10')
    merge(ws,'A11','B12','出生  ',alignment=al(1))
    fill_yellow(ws,11,3,12,11); merge(ws,'C11','K12')
    fill_yellow(ws,11,16,12,18); merge(ws,'P11','R12','（　　）',alignment=al(1))
    fill_yellow(ws,11,25,12,28); merge(ws,'Y11','AB12')
    merge(ws,'A13','B14','死亡  ',alignment=al(1))
    fill_yellow(ws,13,3,14,11); merge(ws,'C13','K14')
    fill_yellow(ws,13,16,14,24); merge(ws,'P13','X14')

    merge(ws,'A15','E15','（被相続人）',alignment=al(2))
    fill_yellow(ws,16,1,17,8); merge(ws,'A16','H17')
    merge(ws,'P16','Q16'); merge(ws,'R16','AA16')
    # 子1
    merge(ws,'P20','Q21','住所',alignment=al(1))
    fill_yellow(ws,20,18,21,31); merge(ws,'R20','AE21')
    merge(ws,'P22','Q22','出生  ',alignment=al(1))
    fill_yellow(ws,22,18,22,27); merge(ws,'R22','AA22')
    fill_yellow(ws,23,16,23,18); merge(ws,'P23','R23','（　　）',alignment=al(1))
    fill_yellow(ws,24,16,25,24); merge(ws,'P24','X25')
    merge(ws,'P27','T27','以下余白',alignment=al(2))
    _creator(ws, 30)
    # 罫線
    set_border(ws,14,13,right=1); set_border(ws,14,14,top=1,left=1); set_border(ws,14,15,top=1)
    for r in range(15,25):
        set_border(ws,r,13,right=1); set_border(ws,r,14,left=1)
    for c in range(9,13): set_border(ws,17,c,top=1)
    set_border(ws,17,13,top=1,right=1)
    set_border(ws,24,14,bottom=1); set_border(ws,24,15,bottom=1)
    for c in range(16,21): set_border(ws,28,c,bottom=1)


# ── Sheet 3 ───────────────────────────────────
def _s3(ws):
    set_row_heights(ws, {4:585,5:90,6:300,7:300,8:300,9:300,10:300,
        11:150,12:150,13:150,14:150,15:150,16:150,17:300,18:150,19:150,
        20:150,21:150,22:150,23:150,24:300,25:150,26:150,27:300,28:300,
        29:150,30:150,31:300,32:300,34:150,35:300,36:300,37:300,38:300,
        40:300,41:300,42:300,43:300,44:300,45:300,46:300,47:300})
    _header(ws)
    merge(ws,'A6','E6','最後の住所　',alignment=al(1))
    fill_yellow(ws,7,1,7,12); merge(ws,'A7','L7')
    merge(ws,'A8','E8','最後の本籍',alignment=al(1))
    fill_yellow(ws,9,1,9,12); merge(ws,'A9','L9')
    merge(ws,'P9','Q9','住所',alignment=al(1))
    fill_yellow(ws,9,18,9,31); merge(ws,'R9','AE9')
    merge(ws,'A10','B10','出生  ',alignment=al(1))
    fill_yellow(ws,10,3,10,11); merge(ws,'C10','K10')
    merge(ws,'P10','Q10','出生  ',alignment=al(1))
    fill_yellow(ws,10,18,10,25); merge(ws,'R10','Y10')
    merge(ws,'A11','B12','死亡  ',alignment=al(1))
    fill_yellow(ws,11,3,12,11); merge(ws,'C11','K12')
    fill_yellow(ws,11,16,12,18); merge(ws,'P11','R12','（　　）',alignment=al(1))
    fill_yellow(ws,11,25,12,28); merge(ws,'Y11','AB12')
    fill_yellow(ws,13,1,14,5); merge(ws,'A13','E14','（被相続人）',alignment=al(1))
    fill_yellow(ws,13,16,14,24); merge(ws,'P13','X14')

    fill_yellow(ws,15,1,16,8); merge(ws,'A15','H16')
    merge(ws,'P15','Q15')
    fill_yellow(ws,15,18,15,27); merge(ws,'R15','AA15')
    # 子1
    merge(ws,'P17','Q17','住所',alignment=al(1))
    fill_yellow(ws,17,18,17,31); merge(ws,'R17','AE17')
    merge(ws,'P18','Q19','出生',alignment=al(1))
    fill_yellow(ws,18,18,19,25); merge(ws,'R18','Y19')
    fill_yellow(ws,20,16,21,18); merge(ws,'P20','R21','（　　）',alignment=al(1))
    fill_yellow(ws,22,16,23,24); merge(ws,'P22','X23')
    # 子2
    merge(ws,'P25','Q26','住所',alignment=al(1))
    fill_yellow(ws,25,18,26,31); merge(ws,'R25','AE26')
    merge(ws,'P27','Q27','出生',alignment=al(1))
    fill_yellow(ws,27,18,27,25); merge(ws,'R27','Y27')
    fill_yellow(ws,28,16,28,18); merge(ws,'P28','R28','（　　）',alignment=al(1))
    fill_yellow(ws,29,16,30,24); merge(ws,'P29','X30')
    merge(ws,'P32','T32','以下余白',alignment=al(2))
    _creator(ws, 35)
    # 罫線
    set_border(ws,13,13,bottom=1); set_border(ws,13,14,bottom=1); set_border(ws,13,15,bottom=1)
    set_border(ws,14,13,top=1,left=1); set_border(ws,14,14,top=1); set_border(ws,14,15,top=1)
    set_border(ws,15,13,left=1)
    for c in range(9,12): set_border(ws,16,c,top=1)
    set_border(ws,16,12,top=1,right=1); set_border(ws,16,13,left=1)
    set_border(ws,17,13,left=1); set_border(ws,18,13,left=1)
    set_border(ws,19,12,right=1); set_border(ws,19,13,left=1)
    for r in range(20,23): set_border(ws,r,13,left=1)
    set_border(ws,23,13,top=1,left=1); set_border(ws,23,14,top=1); set_border(ws,23,15,top=1)
    for r in range(24,29): set_border(ws,r,13,left=1)
    set_border(ws,29,13,bottom=1,left=1); set_border(ws,29,14,bottom=1); set_border(ws,29,15,bottom=1)


# ── Sheet 4 ───────────────────────────────────
def _s4(ws):
    set_row_heights(ws, {4:585,5:90,6:300,7:300,8:300,9:300,10:300,
        11:150,12:150,13:150,14:150,15:150,16:150,17:300,18:150,19:150,
        20:150,21:150,22:150,23:150,24:300,25:150,26:150,27:300,28:300,
        29:150,30:150,31:300,32:300,33:300,34:300,35:150,36:150,37:300,
        38:300,39:300,40:150,41:300,42:300,43:300,44:300})
    _header(ws)
    merge(ws,'A6','E6','最後の住所　',alignment=al(1))
    fill_yellow(ws,7,1,7,12); merge(ws,'A7','L7')
    merge(ws,'A8','E8','最後の本籍',alignment=al(1))
    fill_yellow(ws,9,1,9,12); merge(ws,'A9','L9')
    merge(ws,'P9','Q9','住所',alignment=al(1))
    fill_yellow(ws,9,18,9,31); merge(ws,'R9','AE9')
    merge(ws,'A10','B10','出生  ',alignment=al(1))
    fill_yellow(ws,10,3,10,11); merge(ws,'C10','K10')
    merge(ws,'P10','Q10','出生  ',alignment=al(1))
    fill_yellow(ws,10,18,10,25); merge(ws,'R10','Y10')
    merge(ws,'A11','B12','死亡  ',alignment=al(1))
    fill_yellow(ws,11,3,12,11); merge(ws,'C11','K12')
    fill_yellow(ws,11,16,12,18); merge(ws,'P11','R12','（　　）',alignment=al(1))
    fill_yellow(ws,11,25,12,28); merge(ws,'Y11','AB12')
    fill_yellow(ws,13,1,14,5); merge(ws,'A13','E14','（被相続人）',alignment=al(1))
    fill_yellow(ws,13,16,14,24); merge(ws,'P13','X14')

    fill_yellow(ws,15,1,16,8); merge(ws,'A15','H16')
    merge(ws,'P15','Q15')
    fill_yellow(ws,15,18,15,27); merge(ws,'R15','AA15')
    # 子1
    merge(ws,'P17','Q17','住所',alignment=al(1))
    fill_yellow(ws,17,18,17,31); merge(ws,'R17','AE17')
    merge(ws,'P18','Q19','出生',alignment=al(1))
    fill_yellow(ws,18,18,19,25); merge(ws,'R18','Y19')
    fill_yellow(ws,20,16,21,18); merge(ws,'P20','R21','（　　）',alignment=al(1))
    fill_yellow(ws,22,16,23,24); merge(ws,'P22','X23')
    # 子2
    merge(ws,'P25','Q26','住所',alignment=al(1))
    fill_yellow(ws,25,18,26,31); merge(ws,'R25','AE26')
    merge(ws,'P27','Q27','出生',alignment=al(1))
    fill_yellow(ws,27,18,27,25); merge(ws,'R27','Y27')
    fill_yellow(ws,28,16,28,18); merge(ws,'P28','R28','（　　）',alignment=al(1))
    fill_yellow(ws,29,16,30,24); merge(ws,'P29','X30')
    # 子3
    merge(ws,'P32','Q32','住所',alignment=al(1))
    fill_yellow(ws,32,18,32,31); merge(ws,'R32','AE32')
    merge(ws,'P33','Q33','出生',alignment=al(1))
    fill_yellow(ws,33,18,33,25); merge(ws,'R33','Y33')
    fill_yellow(ws,34,16,34,18); merge(ws,'P34','R34','（　　）',alignment=al(1))
    fill_yellow(ws,35,16,36,24); merge(ws,'P35','X36')
    merge(ws,'P38','T38','以下余白',alignment=al(2))
    _creator(ws, 41)
    # 罫線（s3と同じ基本構造＋追加）
    set_border(ws,13,13,bottom=1); set_border(ws,13,14,bottom=1); set_border(ws,13,15,bottom=1)
    set_border(ws,14,13,top=1,left=1); set_border(ws,14,14,top=1); set_border(ws,14,15,top=1)
    set_border(ws,15,13,left=1)
    for c in range(9,12): set_border(ws,16,c,top=1)
    set_border(ws,16,12,top=1,right=1); set_border(ws,16,13,left=1)
    set_border(ws,17,13,left=1); set_border(ws,18,13,left=1)
    set_border(ws,19,12,right=1); set_border(ws,19,13,left=1)
    for r in range(20,23): set_border(ws,r,13,left=1)
    set_border(ws,23,13,top=1,left=1); set_border(ws,23,14,top=1); set_border(ws,23,15,top=1)
    for r in range(24,29): set_border(ws,r,13,left=1)
    set_border(ws,29,13,bottom=1,left=1); set_border(ws,29,14,bottom=1); set_border(ws,29,15,bottom=1)
    set_border(ws,30,13,top=1,left=1); set_border(ws,30,14,top=1); set_border(ws,30,15,top=1)
    for r in range(31,35): set_border(ws,r,13,left=1)
    set_border(ws,35,13,bottom=1,left=1); set_border(ws,35,14,bottom=1); set_border(ws,35,15,bottom=1)
    for c in range(16,25): set_border(ws,39,c,bottom=1)


# ── Sheet N（子5人以上） ──────────────────────────────────────
def _sN(ws, n):
    """子n人（n≥5）の拡張シート: s4のベース + 子4以降の追加ブロック"""
    # 行高さ: s4の固定部分（子1-3 + gap） + 子4以降の各ブロック + 以下余白 + 作成者
    heights = {4:585,5:90,6:300,7:300,8:300,9:300,10:300,
        11:150,12:150,13:150,14:150,15:150,16:150,17:300,18:150,19:150,
        20:150,21:150,22:150,23:150,24:300,25:150,26:150,27:300,28:300,
        29:150,30:150,31:300,32:300,33:300,34:300,35:150,36:150,37:300}
    for i in range(4, n):          # 子4〜子(n-1) のブロック
        s = _co_child_start(i)     # i=4→38, i=5→44, ...
        heights[s]=300; heights[s+1]=300; heights[s+2]=300
        heights[s+3]=150; heights[s+4]=150; heights[s+5]=300  # s+5 = gap
    wr = 38 + (n - 4) * 6         # 以下余白 行
    heights[wr]=300; heights[wr+1]=300; heights[wr+2]=150
    heights[wr+3]=300; heights[wr+4]=300; heights[wr+5]=300
    set_row_heights(ws, heights)

    _header(ws)

    # 被相続人情報（s4と同一）
    merge(ws,'A6','E6','最後の住所　',alignment=al(1))
    fill_yellow(ws,7,1,7,12); merge(ws,'A7','L7')
    merge(ws,'A8','E8','最後の本籍',alignment=al(1))
    fill_yellow(ws,9,1,9,12); merge(ws,'A9','L9')
    merge(ws,'P9','Q9','住所',alignment=al(1))
    fill_yellow(ws,9,18,9,31); merge(ws,'R9','AE9')
    merge(ws,'A10','B10','出生  ',alignment=al(1))
    fill_yellow(ws,10,3,10,11); merge(ws,'C10','K10')
    merge(ws,'P10','Q10','出生  ',alignment=al(1))
    fill_yellow(ws,10,18,10,25); merge(ws,'R10','Y10')
    merge(ws,'A11','B12','死亡  ',alignment=al(1))
    fill_yellow(ws,11,3,12,11); merge(ws,'C11','K12')
    fill_yellow(ws,11,16,12,18); merge(ws,'P11','R12','（　　）',alignment=al(1))
    fill_yellow(ws,11,25,12,28); merge(ws,'Y11','AB12')
    fill_yellow(ws,13,1,14,5); merge(ws,'A13','E14','（被相続人）',alignment=al(1))
    fill_yellow(ws,13,16,14,24); merge(ws,'P13','X14')

    fill_yellow(ws,15,1,16,8); merge(ws,'A15','H16')
    merge(ws,'P15','Q15')
    fill_yellow(ws,15,18,15,27); merge(ws,'R15','AA15')

    # 子1（s3/s4と同一）
    merge(ws,'P17','Q17','住所',alignment=al(1))
    fill_yellow(ws,17,18,17,31); merge(ws,'R17','AE17')
    merge(ws,'P18','Q19','出生',alignment=al(1))
    fill_yellow(ws,18,18,19,25); merge(ws,'R18','Y19')
    fill_yellow(ws,20,16,21,18); merge(ws,'P20','R21','（　　）',alignment=al(1))
    fill_yellow(ws,22,16,23,24); merge(ws,'P22','X23')

    # 子2
    merge(ws,'P25','Q26','住所',alignment=al(1))
    fill_yellow(ws,25,18,26,31); merge(ws,'R25','AE26')
    merge(ws,'P27','Q27','出生',alignment=al(1))
    fill_yellow(ws,27,18,27,25); merge(ws,'R27','Y27')
    fill_yellow(ws,28,16,28,18); merge(ws,'P28','R28','（　　）',alignment=al(1))
    fill_yellow(ws,29,16,30,24); merge(ws,'P29','X30')

    # 子3（s4と同一）
    _co_add_child_block(ws, 32)

    # 子4以降
    for i in range(4, n):
        _co_add_child_block(ws, _co_child_start(i))

    # 以下余白
    merge(ws, f'P{wr}', f'T{wr}', '以下余白', alignment=al(2))

    # 作成者
    _creator(ws, wr + 3)

    # 罫線（s4と同じ基本構造）
    set_border(ws,13,13,bottom=1); set_border(ws,13,14,bottom=1); set_border(ws,13,15,bottom=1)
    set_border(ws,14,13,top=1,left=1); set_border(ws,14,14,top=1); set_border(ws,14,15,top=1)
    set_border(ws,15,13,left=1)
    for c in range(9,12): set_border(ws,16,c,top=1)
    set_border(ws,16,12,top=1,right=1); set_border(ws,16,13,left=1)
    set_border(ws,17,13,left=1); set_border(ws,18,13,left=1)
    set_border(ws,19,12,right=1); set_border(ws,19,13,left=1)
    for r in range(20,23): set_border(ws,r,13,left=1)
    set_border(ws,23,13,top=1,left=1); set_border(ws,23,14,top=1); set_border(ws,23,15,top=1)
    for r in range(24,29): set_border(ws,r,13,left=1)
    set_border(ws,29,13,bottom=1,left=1); set_border(ws,29,14,bottom=1); set_border(ws,29,15,bottom=1)
    set_border(ws,30,13,top=1,left=1); set_border(ws,30,14,top=1); set_border(ws,30,15,top=1)
    for r in range(31,35): set_border(ws,r,13,left=1)
    set_border(ws,35,13,bottom=1,left=1); set_border(ws,35,14,bottom=1); set_border(ws,35,15,bottom=1)
    # 子4以降のM列罫線
    for i in range(4, n):
        _co_add_m_borders(ws, _co_child_start(i))
    for c in range(16,25): set_border(ws, wr+1, c, bottom=1)
