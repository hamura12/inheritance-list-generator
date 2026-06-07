"""配偶者・親２名（父及び母）である場合"""
from openpyxl import Workbook
from .common import (F11, F14, YEL, al, set_col_widths, set_row_heights,
                     merge, fill_yellow, set_border, set_cell, draw_creator_box)


def generate(appl_start: str = 'X23', appl_end: str = 'AA23') -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('配偶者・親２名の場合')
    set_col_widths(ws)
    set_row_heights(ws, {2:585,3:90,4:300,5:300,6:300,7:300,8:300,9:300,10:300,
        11:300,12:300,13:150,14:150,15:300,16:300,17:300,18:300,19:300,
        20:300,21:300,22:300,23:300,24:300,25:300,26:300,27:150,28:300,
        29:300,30:300,31:300,32:300,33:300,34:300,35:300,36:300,37:300,
        38:300,39:300,40:300})

    # ヘッダー（行2）
    merge(ws,'F2','J2','被相続人',F14,alignment=al(3))
    fill_yellow(ws,2,11,2,20); merge(ws,'K2','T2')
    merge(ws,'U2','AD2','法定相続情報',F14,alignment=al(2))

    merge(ws,'A4','E4')  # 空白

    # 被相続人情報（右側）
    merge(ws,'P6','T6','最後の住所　',alignment=al(1))
    fill_yellow(ws,7,16,7,31); merge(ws,'P7','AE7')

    # 親1（父）情報（左側上）
    merge(ws,'A8','B8','住所',alignment=al(1))
    fill_yellow(ws,8,3,8,15); merge(ws,'C8','O8')
    merge(ws,'P8','T8','最後の本籍　',alignment=al(1))
    fill_yellow(ws,9,16,9,31); merge(ws,'P9','AE9')
    merge(ws,'A9','B9','出生  ',alignment=al(1))
    fill_yellow(ws,9,3,9,11); merge(ws,'C9','K9')

    merge(ws,'A10','B10','（　）',alignment=al(1))
    merge(ws,'P10','Q10','出生  ',alignment=al(1))
    fill_yellow(ws,10,18,10,26); merge(ws,'R10','Z10')

    merge(ws,'P11','Q11','死亡  ',alignment=al(1))
    fill_yellow(ws,11,1,11,7); merge(ws,'A11','G11')
    fill_yellow(ws,11,18,11,26); merge(ws,'R11','Z11')

    merge(ws,'P12','T12','（被相続人）',alignment=al(2))
    fill_yellow(ws,13,16,14,23); merge(ws,'P13','W14')

    # 親2（母）情報（左側下）
    merge(ws,'A16','B16','住所',alignment=al(1))
    fill_yellow(ws,16,3,16,15); merge(ws,'C16','O16')
    merge(ws,'A17','B17','出生',alignment=al(1))
    fill_yellow(ws,17,3,17,11); merge(ws,'C17','K17')
    merge(ws,'A18','B18','（　）',alignment=al(1))
    fill_yellow(ws,19,1,19,7); merge(ws,'A19','G19')

    # 配偶者情報（右側下）
    merge(ws,'P20','Q20','住所',alignment=al(1))
    fill_yellow(ws,20,18,20,31); merge(ws,'R20','AE20')
    merge(ws,'P21','Q21','出生',alignment=al(1))
    fill_yellow(ws,21,18,21,27); merge(ws,'R21','AA21')
    fill_yellow(ws,22,16,22,18); merge(ws,'P22','R22','（　　）',alignment=al(1))
    fill_yellow(ws,23,16,23,23); merge(ws,'P23','W23')
    merge(ws, appl_start, appl_end, '(申出人)', alignment=al(2))

    merge(ws,'P25','T25','以下余白',alignment=al(2))
    merge(ws,'P26','Q26')
    fill_yellow(ws,26,18,26,27); merge(ws,'R26','AA26')

    # 作成者
    set_cell(ws,28,11,'作成日：',alignment=al(0))
    fill_yellow(ws,28,14,28,24); merge(ws,'N28','X28')
    set_cell(ws,29,11,'作成者：',alignment=al(0))
    set_cell(ws,29,14,'住所',alignment=al(1))
    fill_yellow(ws,29,16,29,28); merge(ws,'P29','AB29')
    set_cell(ws,30,14,'氏名',alignment=al(2))
    fill_yellow(ws,30,16,30,22); merge(ws,'P30','V30')
    draw_creator_box(ws, 27)

    # 罫線
    # 親1（A列下→被相続人方向）
    set_border(ws,12,3,left=6); set_border(ws,13,3,left=6)
    set_border(ws,14,3,top=1,left=6)
    for c in range(4,15): set_border(ws,14,c,top=1)
    set_border(ws,15,3,left=6)
    # 被相続人→配偶者の二重線（R列）
    for r in range(15,20): set_border(ws,r,18,left=6)

    return wb
