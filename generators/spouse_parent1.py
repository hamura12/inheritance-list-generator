"""配偶者・親１名（父又は母）である場合"""
from openpyxl import Workbook
from .common import (F11, F14, YEL, al, set_col_widths, set_row_heights,
                     merge, fill_yellow, set_border, set_cell, draw_creator_box)


def generate(which_parent: str = '父',
             appl_start: str = 'X21', appl_end: str = 'AA21') -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('配偶者・親１名の場合')
    set_col_widths(ws)
    set_row_heights(ws, {2:585,3:90,4:300,5:300,6:300,7:300,8:300,9:300,10:300,
        11:300,12:300,13:150,14:150,15:300,16:300,17:300,18:300,19:300,
        20:300,21:300,22:300,23:300,24:300,25:150,26:300,27:300,28:300,
        29:300,30:300,31:300,32:300,33:300,34:300,35:300,36:300,37:300,38:300})

    # ヘッダー（行2）
    merge(ws,'F2','J2','被相続人',F14,alignment=al(3))
    fill_yellow(ws,2,11,2,20); merge(ws,'K2','T2')
    merge(ws,'U2','AD2','法定相続情報',F14,alignment=al(2))

    merge(ws,'A4','E4')  # 空白結合

    # 被相続人情報（右側、P列）
    merge(ws,'P6','T6','最後の住所　',alignment=al(1))
    fill_yellow(ws,7,16,7,31); merge(ws,'P7','AE7')
    merge(ws,'P8','T8','最後の本籍',alignment=al(1))
    fill_yellow(ws,9,16,9,31); merge(ws,'P9','AE9')

    # 配偶者情報（左側、A列）
    merge(ws,'A10','B10','住所',alignment=al(1))
    fill_yellow(ws,10,3,10,15); merge(ws,'C10','O10')
    merge(ws,'P10','Q10','出生  ',alignment=al(1))
    fill_yellow(ws,10,18,10,26); merge(ws,'R10','Z10')

    merge(ws,'A11','B11','出生  ',alignment=al(1))
    fill_yellow(ws,11,3,11,11); merge(ws,'C11','K11')
    merge(ws,'P11','Q11','死亡  ',alignment=al(1))
    fill_yellow(ws,11,18,11,26); merge(ws,'R11','Z11')

    merge(ws,'A12','B12','（　）',alignment=al(1))
    merge(ws,'P12','T12','（被相続人）',alignment=al(2))

    fill_yellow(ws,13,1,14,8); merge(ws,'A13','H14')
    fill_yellow(ws,13,16,14,23); merge(ws,'P13','W14')

    # 親情報（右側下部）
    merge(ws,'P18','Q18','住所',alignment=al(1))
    fill_yellow(ws,18,18,18,31); merge(ws,'R18','AE18')
    merge(ws,'P19','Q19','出生',alignment=al(1))
    fill_yellow(ws,19,18,19,27); merge(ws,'R19','AA19')
    fill_yellow(ws,20,16,20,18); merge(ws,'P20','R20','（　　）',alignment=al(1))
    fill_yellow(ws,21,16,21,23); merge(ws,'P21','W21')
    merge(ws, appl_start, appl_end, '(申出人)', alignment=al(2))

    merge(ws,'P23','T23','以下余白',alignment=al(2))

    # 作成者（N=住所 単体セル）
    set_cell(ws,26,11,'作成日：',alignment=al(0))
    fill_yellow(ws,26,14,26,24); merge(ws,'N26','X26')
    set_cell(ws,27,11,'作成者：',alignment=al(0))
    set_cell(ws,27,14,'住所',alignment=al(1))
    fill_yellow(ws,27,16,27,28); merge(ws,'P27','AB27')
    set_cell(ws,28,14,'氏名',alignment=al(2))
    fill_yellow(ws,28,16,28,22); merge(ws,'P28','V28')
    draw_creator_box(ws, 25)

    # 罫線（被相続人→親の接続二重線）
    for c in range(9,15): set_border(ws,14,c,top=1)  # I14:O14 top
    set_border(ws,15,18,left=6)  # R15 double left（被相続人→親）
    set_border(ws,16,18,left=6)
    set_border(ws,17,18,left=6)
    # 申出人下ライン
    for c in range(16,28): set_border(ws,24,c,bottom=1)

    return wb
