"""父母の一方のみを同じくする兄弟姉妹がいる場合"""
from openpyxl import Workbook
from .common import (F11, F14, YEL, al, set_col_widths, set_row_heights,
                     merge, fill_yellow, set_border, set_cell, draw_creator_box)


def generate(num_siblings: int,
             appl_start: str = 'X24', appl_end: str = 'AA25') -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)
    _build_sheet(wb, num_siblings, appl_start, appl_end)
    return wb


def _header(ws):
    merge(ws,'I2','M2','被相続人',F14,alignment=al(3))
    fill_yellow(ws,2,14,2,20); merge(ws,'N2','T2')
    merge(ws,'U2','AA2','法定相続情報',F14,alignment=al(2))


def _base_content(ws):
    # 配偶者（右側上部）
    set_cell(ws,5,16,'住所',alignment=al(1))
    fill_yellow(ws,5,18,5,31); merge(ws,'R5','AE5')
    merge(ws,'P6','Q6','出生  ',alignment=al(1))
    fill_yellow(ws,6,18,6,27); merge(ws,'R6','AA6')
    fill_yellow(ws,7,16,7,18); merge(ws,'P7','R7','（　　）',alignment=al(1))

    # 被相続人情報（右側中部）
    set_cell(ws,11,16,'最後の住所　',alignment=al(1))
    fill_yellow(ws,12,16,12,31); merge(ws,'P12','AE12')
    set_cell(ws,13,16,'最後の本籍',alignment=al(1))
    merge(ws,'P13','T13')
    fill_yellow(ws,14,16,14,31); merge(ws,'P14','AE14')
    set_cell(ws,15,16,'出生  ',alignment=al(1))
    fill_yellow(ws,15,18,15,26); merge(ws,'R15','Z15')
    set_cell(ws,16,16,'死亡  ',alignment=al(1))
    fill_yellow(ws,16,18,16,26); merge(ws,'R16','Z16')
    set_cell(ws,17,16,'（被相続人）',alignment=al(1))

    # 親エリア（左側）
    fill_yellow(ws,18,16,19,23); merge(ws,'P18','W19')

    # 兄弟姉妹1（右側）
    merge(ws,'P21','Q21','住所',alignment=al(1))
    fill_yellow(ws,21,18,21,31); merge(ws,'R21','AE21')
    merge(ws,'P22','Q22','出生  ',alignment=al(1))
    fill_yellow(ws,22,18,22,27); merge(ws,'R22','AA22')
    merge(ws,'P23','Q23','（　）',alignment=al(1))

    # 第2親エリア（左側）
    fill_yellow(ws,24,16,25,23); merge(ws,'P24','W25')


def _creator_s(ws, row):
    set_cell(ws, row, 11, '作成日：', alignment=al(0))
    fill_yellow(ws, row, 14, row, 24); merge(ws, f'N{row}', f'X{row}')
    set_cell(ws, row+1, 11, '作成者：', alignment=al(0))
    set_cell(ws, row+1, 14, '住所', alignment=al(1))
    fill_yellow(ws, row+1, 16, row+1, 28); merge(ws, f'P{row+1}', f'AB{row+1}')
    set_cell(ws, row+2, 14, '氏名', alignment=al(2))
    fill_yellow(ws, row+2, 16, row+2, 22); merge(ws, f'P{row+2}', f'V{row+2}')
    draw_creator_box(ws, row - 1)


def _base_borders(ws):
    set_border(ws,9,17,right=6)
    set_border(ws,10,18,left=6)
    # K列縦線
    set_border(ws,18,11,bottom=1); set_border(ws,18,12,bottom=1)
    set_border(ws,18,13,bottom=1); set_border(ws,18,14,bottom=1); set_border(ws,18,15,bottom=1)
    set_border(ws,19,11,top=1,left=1); set_border(ws,19,12,top=1)
    set_border(ws,19,13,top=1); set_border(ws,19,14,top=1); set_border(ws,19,15,top=1)
    set_border(ws,20,11,left=1); set_border(ws,21,11,left=1)
    # E列二重線（基本分）
    set_border(ws,22,5,top=1,left=6)
    for c in range(6,11): set_border(ws,22,c,top=1)
    set_border(ws,22,10,right=1); set_border(ws,22,11,left=1)
    set_border(ws,20,5,left=6); set_border(ws,21,5,left=6)
    set_border(ws,23,5,left=6); set_border(ws,23,11,left=1)
    set_border(ws,24,10,right=1); set_border(ws,24,11,left=1)
    for c in range(11,16): set_border(ws,25,c,top=1)


def _sN(ws, n):
    """半血兄弟姉妹N人（N≥4）拡張シート"""
    heights = {2:585,3:90,5:300,6:300,7:300,8:300,9:300,10:300,11:300,
        12:300,13:300,14:300,15:300,16:300,17:300,18:150,19:150,20:300,21:300,
        22:300,23:300,24:150,25:150,26:300,27:300,28:300,29:300,30:150,31:150,
        32:300,33:300,34:300,35:300,36:150,37:150,38:300}
    for i in range(4, n+1):
        s = 21 + (i-1)*6
        heights[s]=300; heights[s+1]=300; heights[s+2]=300
        heights[s+3]=150; heights[s+4]=150; heights[s+5]=300
    wr = 21 + n*6  # 以下余白行
    heights[wr]=300; heights[wr+2]=150; heights[wr+3]=300; heights[wr+4]=300; heights[wr+5]=300
    set_row_heights(ws, heights)

    _header(ws); _base_content(ws)
    fill_yellow(ws,8,16,8,23); merge(ws,'P8','W8')
    merge(ws,'P7','S7','（配偶者）',alignment=al(1))

    # 兄弟姉妹2, 3（静的）
    for s in [27, 33]:
        merge(ws,f'P{s}',f'Q{s}','住所',alignment=al(1))
        fill_yellow(ws,s,18,s,31); merge(ws,f'R{s}',f'AE{s}')
        merge(ws,f'P{s+1}',f'Q{s+1}','出生  ',alignment=al(1))
        fill_yellow(ws,s+1,18,s+1,27); merge(ws,f'R{s+1}',f'AA{s+1}')
        merge(ws,f'P{s+2}',f'Q{s+2}','（　）',alignment=al(1))
        fill_yellow(ws,s+3,16,s+4,23); merge(ws,f'P{s+3}',f'W{s+4}')

    # 兄弟姉妹4以降（動的）
    for i in range(4, n+1):
        s = 21 + (i-1)*6
        merge(ws,f'P{s}',f'Q{s}','住所',alignment=al(1))
        fill_yellow(ws,s,18,s,31); merge(ws,f'R{s}',f'AE{s}')
        merge(ws,f'P{s+1}',f'Q{s+1}','出生  ',alignment=al(1))
        fill_yellow(ws,s+1,18,s+1,27); merge(ws,f'R{s+1}',f'AA{s+1}')
        merge(ws,f'P{s+2}',f'Q{s+2}','（　）',alignment=al(1))
        fill_yellow(ws,s+3,16,s+4,23); merge(ws,f'P{s+3}',f'W{s+4}')

    merge(ws,f'P{wr}',f'T{wr}','以下余白',alignment=al(2))
    _creator_s(ws, wr+3)

    _base_borders(ws)
    # 半血：（　）ラベル
    merge(ws,'D18','E19','（　）',alignment=al(2))
    merge(ws,'D24','E25','（　）',alignment=al(2))
    merge(ws,f'D{wr-1}',f'E{wr-1}','（　）',alignment=al(2))

    # E列延長二重線（半血特有）：row 26 〜 creator行まで
    for r in range(26, wr+4): set_border(ws,r,5,left=6)

    # 兄弟姉妹2のK列罫線（静的）
    set_border(ws,25,11,top=1,left=1)
    for c in range(12,16): set_border(ws,25,c,top=1)
    set_border(ws,26,11,left=1)
    for r in range(27,30): set_border(ws,r,11,left=1)
    set_border(ws,30,11,bottom=1,left=1)
    for c in range(12,16): set_border(ws,30,c,bottom=1)

    # 兄弟姉妹3のK列罫線（静的）
    set_border(ws,31,11,top=1,left=1)
    for c in range(12,16): set_border(ws,31,c,top=1)
    set_border(ws,32,11,left=1)
    for r in range(33,36): set_border(ws,r,11,left=1)
    set_border(ws,36,11,bottom=1,left=1)
    for c in range(12,16): set_border(ws,36,c,bottom=1)

    # 兄弟姉妹4以降のK列罫線（動的）
    for i in range(4, n+1):
        s = 21 + (i-1)*6
        set_border(ws,s-2,11,top=1,left=1)
        for c in range(12,16): set_border(ws,s-2,c,top=1)
        set_border(ws,s-1,11,left=1)
        for r in range(s,s+3): set_border(ws,r,11,left=1)
        set_border(ws,s+3,11,bottom=1,left=1)
        for c in range(12,16): set_border(ws,s+3,c,bottom=1)


def _build_sheet(wb, n, appl_start='X24', appl_end='AA25'):
    ws = wb.create_sheet(f'半血兄弟姉妹{n}人')
    set_col_widths(ws)

    if n == 1:
        set_row_heights(ws, {2:585,3:90,5:300,6:300,7:300,8:300,9:300,10:300,11:300,
            12:300,13:300,14:300,15:300,16:300,17:300,18:150,19:150,20:300,21:300,
            22:300,23:300,24:150,25:150,26:300,27:300,28:300,29:300,30:300,31:300,
            32:300,33:300,34:300,36:150,37:300,38:300,39:300,40:300,
            42:300,43:300,44:300,45:300,46:300,47:300,48:300,49:300})
        _header(ws); _base_content(ws)
        fill_yellow(ws,8,16,8,22); merge(ws,'P8','V8')
        merge(ws,'P27','T27','以下余白',alignment=al(2))
        _creator_s(ws, 37)
        _base_borders(ws)
        # 半血：親ラベルを（　）に
        merge(ws,'D18','E19','（　）',alignment=al(2))
        merge(ws,'D24','E25','（　）',alignment=al(2))
        # E列延長二重線（半血特有）
        for r in range(26,38): set_border(ws,r,5,left=6)

    elif n == 2:
        set_row_heights(ws, {2:585,3:90,5:300,6:300,7:300,8:300,9:300,10:300,11:300,
            12:300,13:300,14:300,15:300,16:300,17:300,18:150,19:150,20:300,21:300,
            22:300,23:300,24:150,25:150,26:300,27:300,28:300,29:300,30:150,31:150,
            32:300,33:300,34:300,35:300,36:150,37:150,38:300,40:150,41:300,42:300,
            43:300,44:300,46:300,47:300,48:300,49:300,50:300,51:300,52:300,53:300})
        _header(ws); _base_content(ws)
        fill_yellow(ws,8,16,8,23); merge(ws,'P8','W8')
        # 兄弟姉妹2
        merge(ws,'P27','Q27','住所',alignment=al(1))
        fill_yellow(ws,27,18,27,31); merge(ws,'R27','AE27')
        merge(ws,'P28','Q28','出生  ',alignment=al(1))
        fill_yellow(ws,28,18,28,27); merge(ws,'R28','AA28')
        merge(ws,'P29','Q29','（　）',alignment=al(1))
        fill_yellow(ws,30,16,31,23); merge(ws,'P30','W31')
        merge(ws,'P33','T33','以下余白',alignment=al(2))
        _creator_s(ws, 41)
        _base_borders(ws)
        merge(ws,'D18','E19','（　）',alignment=al(2))
        merge(ws,'D24','E25','（　）',alignment=al(2))
        for r in range(26,42): set_border(ws,r,5,left=6)
        # 兄弟姉妹2のK列罫線
        set_border(ws,25,11,top=1,left=1)
        for c in range(12,16): set_border(ws,25,c,top=1)
        set_border(ws,26,11,left=1)
        set_border(ws,27,11,left=1); set_border(ws,28,11,left=1); set_border(ws,29,11,left=1)
        set_border(ws,30,11,bottom=1,left=1)
        for c in range(12,16): set_border(ws,30,c,bottom=1)
        set_border(ws,31,11,top=1)
        for c in range(12,16): set_border(ws,31,c,top=1)

    elif n == 3:
        set_row_heights(ws, {2:585,3:90,5:300,6:300,7:300,8:300,9:300,10:300,11:300,
            12:300,13:300,14:300,15:300,16:300,17:300,18:150,19:150,20:300,21:300,
            22:300,23:300,24:150,25:150,26:300,27:300,28:300,29:300,30:150,31:150,
            32:300,33:300,34:300,35:300,36:150,37:150,38:300,39:300,41:150,42:300,
            43:300,44:300,45:300,47:300,48:300,49:300,50:300,51:300,52:300,53:300,54:300})
        _header(ws); _base_content(ws)
        fill_yellow(ws,8,16,8,23); merge(ws,'P8','W8')
        merge(ws,'P7','S7','（配偶者）',alignment=al(1))
        # 兄弟姉妹2
        merge(ws,'P27','Q27','住所',alignment=al(1))
        fill_yellow(ws,27,18,27,31); merge(ws,'R27','AE27')
        merge(ws,'P28','Q28','出生  ',alignment=al(1))
        fill_yellow(ws,28,18,28,27); merge(ws,'R28','AA28')
        merge(ws,'P29','Q29','（　）',alignment=al(1))
        fill_yellow(ws,30,16,31,23); merge(ws,'P30','W31')
        # 兄弟姉妹3
        merge(ws,'P33','Q33','住所',alignment=al(1))
        fill_yellow(ws,33,18,33,31); merge(ws,'R33','AE33')
        merge(ws,'P34','Q34','出生  ',alignment=al(1))
        fill_yellow(ws,34,18,34,27); merge(ws,'R34','AA34')
        merge(ws,'P35','Q35','（　）',alignment=al(1))
        fill_yellow(ws,36,16,37,23); merge(ws,'P36','W37')
        merge(ws,'P39','T39','以下余白',alignment=al(2))
        _creator_s(ws, 42)
        _base_borders(ws)
        merge(ws,'D18','E19','（　）',alignment=al(2))
        merge(ws,'D24','E25','（　）',alignment=al(2))
        merge(ws,'D38','E38','（　）',alignment=al(2))
        for r in range(26,43): set_border(ws,r,5,left=6)
        # 兄弟姉妹2のK列罫線
        set_border(ws,25,11,top=1,left=1)
        for c in range(12,16): set_border(ws,25,c,top=1)
        set_border(ws,26,11,left=1)
        for r in range(27,30): set_border(ws,r,11,left=1)
        set_border(ws,30,11,bottom=1,left=1)
        for c in range(12,16): set_border(ws,30,c,bottom=1)
        # 兄弟姉妹3のK列罫線
        set_border(ws,31,11,top=1,left=1)
        for c in range(12,16): set_border(ws,31,c,top=1)
        set_border(ws,32,11,left=1); set_border(ws,33,11,left=1)
        set_border(ws,34,11,left=1); set_border(ws,35,11,left=1)
        set_border(ws,36,11,bottom=1,left=1)
        for c in range(12,16): set_border(ws,36,c,bottom=1)

    else:
        _sN(ws, n)

    # 申出人ラベルを選択した相続人の氏名欄右に配置
    merge(ws, appl_start, appl_end, '(申出人)', alignment=al(2))

    return ws
