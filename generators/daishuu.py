"""代襲相続が生じている場合（動的生成版）"""
from openpyxl import Workbook
from .common import (F11, F14, YEL, al, set_col_widths, set_row_heights,
                     merge, fill_yellow, set_border, set_cell, draw_creator_box)

# 各ブロックの行数定数
CHILD_ROWS = 6   # 子1人あたりの左ゾーン行数（住所2+出生2+氏名2）
GC_ROWS    = 6   # 孫1人あたりの右ゾーン行数（住所2+出生2+氏名2）


def generate(has_spouse=True, children_data=None,
             appl_start=None, appl_end=None) -> Workbook:
    """
    has_spouse   : 配偶者の有無
    children_data: [{'alive': bool, 'num_grandchildren': int}, ...]
    appl_start   : (申出人)ラベルの左上セルアドレス（省略時はデフォルト位置）
    appl_end     : (申出人)ラベルの右下セルアドレス
    """
    if not children_data:
        children_data = [{'alive': False, 'num_grandchildren': 1}]

    wb = Workbook()
    wb.remove(wb.active)

    # ── シート名 ──
    nc = len(children_data)
    total_gc = sum(c.get('num_grandchildren', 0)
                   for c in children_data if not c['alive'])
    parts = []
    if has_spouse:
        parts.append('配偶者')
    parts.append(f'子{nc}人')
    parts.append(f'孫{total_gc}人（代襲）')
    sname = '・'.join(parts) + 'の場合'
    if len(sname) > 31:
        sname = '代襲相続の場合'

    ws = wb.create_sheet(sname)
    set_col_widths(ws)

    # ── 行の高さ（固定ヘッダー部分）──
    rh = {4: 585, 5: 90}
    for r in range(6, 16):
        rh[r] = 300
    rh[11] = 150
    rh[12] = 150
    rh[13] = 150
    rh[14] = 150

    # ── 動的セクションのレイアウト計算 ──
    cur = 16

    sp_row = None
    if has_spouse:
        sp_row = cur
        rh[cur]     = 300
        rh[cur + 1] = 300
        cur += 2

    # 各子グループのレイアウト
    groups = []
    for idx, child in enumerate(children_data, start=1):
        g_start = cur
        if child['alive']:
            g_rows = CHILD_ROWS
        else:
            num_gc = child.get('num_grandchildren', 0)
            g_rows = max(CHILD_ROWS, num_gc * GC_ROWS)

        for r in range(cur, cur + g_rows):
            rh[r] = 300

        groups.append({
            'start': g_start,
            'rows':  g_rows,
            'idx':   idx,
            'child': child,
        })
        cur += g_rows

    # 以下余白
    yuuhaku_row = cur
    rh[yuuhaku_row] = 300
    cur += 2  # 以下余白行 + 空白1行

    # 作成者ボックス
    box_row = cur
    for r in range(box_row, box_row + 6):
        rh[r] = 300

    set_row_heights(ws, rh)

    # ══════════════════════════════════════════
    # ヘッダー（行4）
    # ══════════════════════════════════════════
    merge(ws, 'I4', 'M4', '被相続人', F14, alignment=al(3))
    fill_yellow(ws, 4, 14, 4, 20)
    merge(ws, 'N4', 'T4')
    merge(ws, 'U4', 'AA4', '法定相続情報', F14, alignment=al(2))

    # ══════════════════════════════════════════
    # 被相続人情報（行7-15）
    # ══════════════════════════════════════════
    merge(ws, 'A7', 'E7', '最後の住所', alignment=al(1))
    fill_yellow(ws, 8, 1, 8, 11)
    merge(ws, 'A8', 'K8')
    fill_yellow(ws, 8, 15, 9, 23)
    merge(ws, 'O8', 'W9')
    merge(ws, 'A9', 'E9', '最後の本籍', alignment=al(1))
    merge(ws, 'M9', 'N9', '住所', alignment=al(1))
    fill_yellow(ws, 10, 1, 10, 11)
    merge(ws, 'A10', 'K10')
    fill_yellow(ws, 10, 15, 10, 22)
    merge(ws, 'O10', 'V10')
    merge(ws, 'M10', 'N10', '出生  ', alignment=al(1))
    merge(ws, 'A11', 'B12', '出生  ', alignment=al(1))
    fill_yellow(ws, 11, 3, 12, 10)
    merge(ws, 'C11', 'J12')
    merge(ws, 'M11', 'O12', '（　）', alignment=al(1))
    merge(ws, 'A13', 'B14', '死亡  ', alignment=al(1))
    fill_yellow(ws, 13, 3, 14, 10)
    merge(ws, 'C13', 'J14')
    fill_yellow(ws, 13, 13, 14, 19)
    merge(ws, 'M13', 'S14')

    # 申出人ラベル
    if appl_start and appl_end:
        merge(ws, appl_start, appl_end, '(申出人)', alignment=al(2))
    else:
        merge(ws, 'T13', 'W14', '(申出人)', alignment=al(2))

    merge(ws, 'A15', 'E15', '（被相続人）', alignment=al(2))

    # ══════════════════════════════════════════
    # 配偶者（左ゾーン＋右ゾーン）
    # ══════════════════════════════════════════
    if sp_row is not None:
        r = sp_row
        fill_yellow(ws, r, 1, r, 9)
        merge(ws, f'A{r}', f'I{r}')
        fill_yellow(ws, r, 24, r + 1, 32)
        merge(ws, f'X{r}', f'AF{r + 1}')

    # ══════════════════════════════════════════
    # 子グループ
    # ══════════════════════════════════════════
    for g in groups:
        r     = g['start']
        idx   = g['idx']
        child = g['child']

        # ── 左ゾーン：子の住所・出生・氏名 ──
        merge(ws, f'A{r}',     f'B{r + 1}', '住所　', alignment=al(1))
        fill_yellow(ws, r, 3, r + 1, 11)
        merge(ws, f'C{r}',     f'K{r + 1}')

        merge(ws, f'A{r + 2}', f'B{r + 3}', '出生  ', alignment=al(1))
        fill_yellow(ws, r + 2, 3, r + 3, 10)
        merge(ws, f'C{r + 2}', f'J{r + 3}')

        fill_yellow(ws, r + 4, 1, r + 5, 9)
        merge(ws, f'A{r + 4}', f'I{r + 5}')

        if child['alive']:
            # ── 生存子のラベル（センター） ──
            merge(ws, f'M{r + 4}', f'T{r + 5}',
                  f'（子{idx}・相続人）', alignment=al(2))

        else:
            # ── 被代襲者ラベル（センター） ──
            merge(ws, f'M{r + 4}', f'R{r + 5}',
                  '被代襲者', alignment=al(2))

            # ── 右ゾーン：孫グループ ──
            num_gc = child.get('num_grandchildren', 0)
            for j in range(num_gc):
                gc = r + j * GC_ROWS   # 孫jの開始行

                # 住所
                merge(ws, f'V{gc}',     f'W{gc + 1}', '住所', alignment=al(1))
                fill_yellow(ws, gc, 24, gc + 1, 31)
                merge(ws, f'X{gc}',     f'AE{gc + 1}')

                # 出生
                merge(ws, f'V{gc + 2}', f'W{gc + 3}', '出生', alignment=al(1))
                fill_yellow(ws, gc + 2, 24, gc + 3, 31)
                merge(ws, f'X{gc + 2}', f'AE{gc + 3}')

                # 氏名
                fill_yellow(ws, gc + 4, 24, gc + 5, 32)
                merge(ws, f'X{gc + 4}', f'AF{gc + 5}')
                merge(ws, f'M{gc + 4}', f'T{gc + 5}',
                      f'（孫{j + 1}・代襲者）', alignment=al(2))

    # ══════════════════════════════════════════
    # 以下余白
    # ══════════════════════════════════════════
    merge(ws, f'V{yuuhaku_row}', f'Z{yuuhaku_row}', '以下余白', alignment=al(2))

    # ══════════════════════════════════════════
    # 作成者
    # ══════════════════════════════════════════
    draw_creator_box(ws, box_row)
    set_cell(ws, box_row + 1, 11, '作成日：', alignment=al(0))
    fill_yellow(ws, box_row + 1, 14, box_row + 1, 24)
    merge(ws, f'N{box_row + 1}', f'X{box_row + 1}')
    set_cell(ws, box_row + 2, 11, '作成者：', alignment=al(0))
    merge(ws, f'N{box_row + 2}', f'O{box_row + 2}', '住所', alignment=al(1))
    fill_yellow(ws, box_row + 2, 16, box_row + 2, 26)
    merge(ws, f'P{box_row + 2}', f'Z{box_row + 2}')
    set_cell(ws, box_row + 3, 14, '氏名', alignment=al(2))
    fill_yellow(ws, box_row + 3, 16, box_row + 3, 22)
    merge(ws, f'P{box_row + 3}', f'V{box_row + 3}')

    return wb
