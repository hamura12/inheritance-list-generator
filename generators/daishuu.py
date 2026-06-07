"""代襲相続が生じている場合（配偶者・子複数名・子について代襲相続）"""
from openpyxl import Workbook
from .common import (F11, F14, YEL, al, set_col_widths, set_row_heights,
                     merge, fill_yellow, set_border, set_cell, draw_creator_box)


def generate() -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('配偶者・子2人・孫1人の場合')
    set_col_widths(ws)
    set_row_heights(ws, {4:585,5:90,6:300,7:300,8:300,9:300,10:300,11:150,12:150,
        13:150,14:150,15:300,16:300,17:300,18:150,19:150,20:150,21:150,22:150,
        23:150,24:300,25:150,26:150,27:300,28:300,29:150,30:150,31:150,32:150,
        33:300,34:300,36:150,37:300,38:300,39:300,40:300,
        42:300,43:300,44:300,45:300,46:300,47:300,48:300,49:300})

    # ヘッダー（行4）
    merge(ws,'I4','M4','被相続人',F14,alignment=al(3))
    fill_yellow(ws,4,14,4,20); merge(ws,'N4','T4')
    merge(ws,'U4','AA4','法定相続情報',F14,alignment=al(2))

    # 被相続人情報（左側）
    merge(ws,'A7','E7','最後の住所',alignment=al(1))
    fill_yellow(ws,8,1,8,11); merge(ws,'A8','K8')
    fill_yellow(ws,8,15,9,23); merge(ws,'O8','W9')
    merge(ws,'A9','E9','最後の本籍',alignment=al(1))
    merge(ws,'M9','N9','住所',alignment=al(1))
    fill_yellow(ws,10,1,10,11); merge(ws,'A10','K10')
    fill_yellow(ws,10,15,10,22); merge(ws,'O10','V10')
    merge(ws,'M10','N10','出生  ',alignment=al(1))
    merge(ws,'A11','B12','出生  ',alignment=al(1))
    fill_yellow(ws,11,3,12,10); merge(ws,'C11','J12')
    merge(ws,'M11','O12','（　）',alignment=al(1))
    merge(ws,'A13','B14','死亡  ',alignment=al(1))
    fill_yellow(ws,13,3,14,10); merge(ws,'C13','J14')
    fill_yellow(ws,13,13,14,19); merge(ws,'M13','S14')
    merge(ws,'T13','W14','(申出人)',alignment=al(2))
    merge(ws,'A15','E15','（被相続人）',alignment=al(2))

    # 配偶者（左側名前エリア）
    fill_yellow(ws,16,1,16,9); merge(ws,'A16','I16')
    fill_yellow(ws,16,24,17,32); merge(ws,'X16','AF17')

    # 孫（代襲者1）右側
    merge(ws,'V17','W17','住所',alignment=al(1))
    fill_yellow(ws,18,24,19,31); merge(ws,'X18','AE19')
    merge(ws,'V18','W19','出生  ',alignment=al(1))
    merge(ws,'V20','AB21','（孫・代襲者）',alignment=al(2))
    fill_yellow(ws,20,3,23,11); merge(ws,'C20','K23')

    # 子1（左側下）
    merge(ws,'A22','B23','住所　',alignment=al(1))
    fill_yellow(ws,22,22,23,28); merge(ws,'V22','AB23')
    merge(ws,'A24','B24','出生  ',alignment=al(1))
    fill_yellow(ws,24,3,24,10); merge(ws,'C24','J24')

    # 被代襲者
    merge(ws,'A25','D26','（　　　）',alignment=al(1))
    merge(ws,'M25','R26','被代襲者',alignment=al(2))
    fill_yellow(ws,25,24,27,32); merge(ws,'X25','AF27')

    merge(ws,'L27','T27','（　　　　　　　　　　　　）',alignment=al(1))
    fill_yellow(ws,27,1,27,9); merge(ws,'A27','I27')

    # 孫（代襲者2）右側
    merge(ws,'V27','W27','住所',alignment=al(1))
    fill_yellow(ws,28,24,28,31); merge(ws,'X28','AE28')
    merge(ws,'V28','W28','出生',alignment=al(1))
    merge(ws,'V29','AB30','（孫・代襲者）',alignment=al(2))
    fill_yellow(ws,31,22,32,28); merge(ws,'V31','AB32')

    merge(ws,'V34','Z34','以下余白',alignment=al(2))

    # 作成者（N:O結合あり）
    set_cell(ws,37,11,'作成日：',alignment=al(0))
    fill_yellow(ws,37,14,37,24); merge(ws,'N37','X37')
    set_cell(ws,38,11,'作成者：',alignment=al(0))
    merge(ws,'N38','O38','住所',alignment=al(1))
    fill_yellow(ws,38,16,38,26); merge(ws,'P38','Z38')
    set_cell(ws,39,14,'氏名',alignment=al(2))
    fill_yellow(ws,39,16,39,22); merge(ws,'P39','V39')
    draw_creator_box(ws, 36)

    # 罫線
    set_border(ws,14,12,top=1,left=1)
    set_border(ws,15,12,left=1); set_border(ws,16,12,left=1)
    set_border(ws,17,3,left=6); set_border(ws,17,12,left=1)
    set_border(ws,18,3,left=6); set_border(ws,18,12,left=1)
    set_border(ws,19,3,top=1,left=6)
    for c in range(4,10): set_border(ws,19,c,top=1)
    set_border(ws,19,10,top=1); set_border(ws,19,12,left=1)
    set_border(ws,20,2,right=6); set_border(ws,20,11,right=1); set_border(ws,20,12,left=1)
    set_border(ws,21,2,right=6); set_border(ws,21,11,right=1); set_border(ws,21,12,left=1)
    set_border(ws,22,11,right=1); set_border(ws,22,12,left=1)
    set_border(ws,23,11,right=1); set_border(ws,23,12,left=1)
    set_border(ws,23,21,top=1,left=1)
    set_border(ws,24,12,left=1); set_border(ws,24,21,left=1)
    set_border(ws,25,12,bottom=1,left=1); set_border(ws,25,21,left=1)
    set_border(ws,26,19,top=1); set_border(ws,26,20,top=1); set_border(ws,26,21,left=1)
    set_border(ws,27,20,right=1); set_border(ws,27,21,left=1)
    for r in range(28,31): set_border(ws,r,21,left=1)
    set_border(ws,31,21,bottom=1,left=1)

    return wb
