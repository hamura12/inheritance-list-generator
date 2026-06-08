"""法定相続情報一覧図 自動生成アプリ"""
import io
from datetime import date
import streamlit as st
from decision import decide, TEMPLATES


# ── 申出人候補の計算 ────────────────────────────────────────────

def _compute_candidates(has_spouse, has_children, num_children,
                        parents_status, has_siblings, siblings_data):
    """
    現在の入力状態から申出人候補の (表示ラベル, appl_start, appl_end) リストを返す。
    各位置はジェネレータの merge() に渡すセルアドレス文字列。
    """
    cands = []
    n = int(num_children)

    if has_children:
        if has_spouse:
            # ── 配偶者・子 ──
            if n == 1:
                cands = [
                    ('配偶者',  'Y16', 'AB17'),
                    ('子1',     'Y18', 'AB19'),
                ]
            elif n == 2:
                cands = [
                    ('配偶者',  'Y13', 'AB14'),
                    ('子1',     'Y24', 'AB25'),
                    ('子2',     'AC11', 'AF12'),
                ]
            elif n == 3:
                cands = [
                    ('配偶者',  'Y13', 'AB14'),
                    ('子1',     'Y21', 'AB22'),
                    ('子2',     'Y28', 'AB29'),
                    ('子3',     'AC11', 'AF12'),
                ]
            elif n == 4:
                cands = [
                    ('配偶者',  'Y13', 'AB14'),
                    ('子1',     'Y21', 'AB22'),
                    ('子2',     'Y28', 'AB29'),
                    ('子3',     'Y34', 'AB35'),
                    ('子4',     'AC11', 'AF12'),
                ]
            else:
                # n >= 5: 拡張シート
                # 子3と子4は同行(rows 34-35)になるため子3位置を共用
                cands = [
                    ('配偶者',  'Y13', 'AB14'),
                    ('子1',     'Y21', 'AB22'),
                    ('子2',     'Y28', 'AB29'),
                    ('子3・子4', 'Y34', 'AB35'),
                ]
                # 子5以降: _child_start(i) = 31 + (i-4)*6, 氏名欄 = start+3:start+4
                for i in range(5, n + 1):
                    s = 31 + (i - 4) * 6
                    cands.append((f'子{i}', f'Y{s+3}', f'AB{s+4}'))
        else:
            # ── 子のみ ──
            if n == 1:
                cands = [('子1', 'Y13', 'AB14')]
            elif n == 2:
                cands = [
                    ('子1', 'Y13', 'AB14'),
                    ('子2', 'Y24', 'AB25'),
                ]
            elif n == 3:
                cands = [
                    ('子1', 'Y13', 'AB14'),
                    ('子2', 'Y22', 'AB23'),
                    ('子3', 'Y29', 'AB30'),
                ]
            elif n == 4:
                cands = [
                    ('子1', 'Y13', 'AB14'),
                    ('子2', 'Y22', 'AB23'),
                    ('子3', 'Y29', 'AB30'),
                    ('子4', 'Y35', 'AB36'),
                ]
            else:
                # n >= 5: _co_child_start(i) = 32 + (i-3)*6, 氏名欄 = start+3:start+4
                cands = [
                    ('子1', 'Y13', 'AB14'),
                    ('子2', 'Y22', 'AB23'),
                ]
                for i in range(3, n + 1):
                    s = 32 + (i - 3) * 6
                    cands.append((f'子{i}', f'Y{s+3}', f'AB{s+4}'))

    elif parents_status == '両親とも存命':
        # ── 配偶者・父母 ──
        cands = [
            ('配偶者', 'X23', 'AA23'),
            ('父',     'H11', 'L11'),
            ('母',     'H19', 'L19'),
        ]

    elif parents_status == '父または母のみ存命':
        # ── 配偶者・親１名 ──
        cands = [
            ('配偶者', 'I13', 'M14'),
            ('親',     'X21', 'AA21'),
        ]

    elif has_siblings:
        any_deceased = any(not s['alive'] for s in siblings_data)

        if has_spouse:
            cands.append(('配偶者', 'X8', 'AA9'))

        if any_deceased:
            # 代襲あり: 生存兄弟姉妹 + 甥姪
            cur = 21
            for i, sib in enumerate(siblings_data):
                if sib['alive']:
                    name_row = cur + 3
                    cands.append((f'兄弟姉妹{i+1}（生存）',
                                  f'X{name_row}', f'AA{name_row+1}'))
                    cur += 6
                else:
                    cur += 7
                    for j in range(sib.get('num_children', 0)):
                        name_row = cur + 3
                        cands.append((f'兄弟姉妹{i+1}の子{j+1}（甥・姪）',
                                      f'X{name_row}', f'AA{name_row+1}'))
                        cur += 6
        else:
            # 代襲なし: 全兄弟姉妹
            for i in range(len(siblings_data)):
                name_row = 24 + i * 6
                cands.append((f'兄弟姉妹{i+1}',
                               f'X{name_row}', f'AA{name_row+1}'))

    return cands


# ── 相続関係図プレビュー ──────────────────────────────────────────

def _esc(text: str) -> str:
    """SVG/HTMLに埋め込む文字列をエスケープしてXSSを防ぐ。"""
    return (str(text)
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;')
            .replace("'", '&#x27;'))


def _generate_preview_svg(has_spouse, has_children, num_children,
                           parents_status, has_siblings, siblings_data,
                           which_parent, has_daishuu, appl_label,
                           children_data=None):
    """相続関係のSVGプレビューを生成してHTML文字列を返す。"""
    BW, BH  = 100, 42
    NAVY    = '#1b2436'
    GOLD    = '#a9853f'
    WHITE   = '#FFFFFF'
    HEIR_BG = '#f5f0e6'
    DEAD_BG = '#cfc7b4'
    APPL_BG = '#fdf5e4'
    TDARK   = '#262420'
    TDEAD   = '#6c6555'
    LC      = '#6c7a8a'
    W       = 720

    def box(cx, cy, label, dec=False, appl=False, dead=False):
        x, y = cx - BW // 2, cy - BH // 2
        if dec:
            bg, sc, sw, tc = NAVY, NAVY, 2, WHITE
        elif appl:
            bg, sc, sw, tc = APPL_BG, GOLD, 2.5, TDARK
        elif dead:
            bg, sc, sw, tc = DEAD_BG, '#8A7A6A', 1.5, TDEAD
        else:
            bg, sc, sw, tc = HEIR_BG, '#8A9AB8', 1.5, TDARK
        s = [
            f'<rect x="{x}" y="{y}" width="{BW}" height="{BH}" rx="4" '
            f'fill="{bg}" stroke="{sc}" stroke-width="{sw}"/>',
            f'<text x="{cx}" y="{cy+1}" text-anchor="middle" dominant-baseline="middle" '
            f'font-size="12" fill="{tc}" font-weight="bold" '
            f'font-family="Noto Sans JP,sans-serif">{_esc(label)}</text>',
        ]
        if dec:
            s.append(
                f'<text x="{cx}" y="{y-10}" text-anchor="middle" '
                f'font-size="10" fill="#4A5A7A" font-family="sans-serif">（被相続人）</text>'
            )
        if appl:
            s.append(
                f'<rect x="{x-2}" y="{y-2}" width="{BW+4}" height="{BH+4}" '
                f'rx="5" fill="none" stroke="{GOLD}" stroke-width="2" stroke-dasharray="5,3"/>'
            )
            s.append(
                f'<text x="{x+BW}" y="{y-8}" text-anchor="end" font-size="9" '
                f'fill="{GOLD}" font-family="sans-serif" font-weight="bold">▶ 申出人</text>'
            )
        if dead:
            s.append(
                f'<text x="{cx}" y="{y-8}" text-anchor="middle" font-size="9" '
                f'fill="#8A6A5A" font-family="sans-serif">（故）</text>'
            )
        return ''.join(s)

    def ml(x1, x2, y):   # 婚姻（二重線）
        return (
            f'<line x1="{x1}" y1="{y-2}" x2="{x2}" y2="{y-2}" '
            f'stroke="{LC}" stroke-width="1.5"/>'
            f'<line x1="{x1}" y1="{y+2}" x2="{x2}" y2="{y+2}" '
            f'stroke="{LC}" stroke-width="1.5"/>'
        )

    def hl(x1, x2, y):
        return (f'<line x1="{x1}" y1="{y}" x2="{x2}" y2="{y}" '
                f'stroke="{LC}" stroke-width="1.5"/>')

    def vl(x, y1, y2):
        return (f'<line x1="{x}" y1="{y1}" x2="{x}" y2="{y2}" '
                f'stroke="{LC}" stroke-width="1.5"/>')

    def wrap(h, els):
        return (
            '<div style="background:#fffdf8;border-radius:2px;'
            'padding:16px 20px 10px;margin:16px 0;'
            'border:1px solid #cfc7b4;border-top:3px solid #1b2436;">'
            '<p style="text-align:center;font-size:0.78rem;color:#6c6555;'
            'margin:0 0 8px;font-weight:600;letter-spacing:2px;">相 続 関 係 プ レ ビ ュ ー</p>'
            f'<svg viewBox="0 0 {W} {h}" width="100%" xmlns="http://www.w3.org/2000/svg">'
            + ''.join(els) +
            '</svg></div>'
        )

    els = []

    # ─── 代襲相続 ───────────────────────────────────────────
    if has_daishuu and children_data:
        cd   = children_data or []
        n    = len(cd)
        SHOW = min(n, 4)

        if has_spouse:
            dec_cx, dec_cy = W // 2 - 80, 90
            sp_cx          = W // 2 + 80
            els.append(ml(dec_cx + BW // 2, sp_cx - BW // 2, dec_cy))
            els.append(box(dec_cx, dec_cy, '被相続人', dec=True))
            els.append(box(sp_cx,  dec_cy, '配偶者'))
        else:
            dec_cx, dec_cy = W // 2, 90
            els.append(box(dec_cx, dec_cy, '被相続人', dec=True))

        child_y = 230
        step    = min(140, (W - 60) // max(SHOW, 1))
        start_x = W // 2 - step * (SHOW - 1) // 2
        mid_y   = (dec_cy + BH // 2 + child_y - BH // 2) // 2

        els.append(vl(dec_cx, dec_cy + BH // 2, mid_y))
        if SHOW > 1:
            els.append(hl(start_x, start_x + step * (SHOW - 1), mid_y))

        max_y = child_y
        for j in range(SHOW):
            cx    = start_x + j * step
            child = cd[j] if j < len(cd) else {'alive': True, 'num_grandchildren': 0}
            alive = child.get('alive', True)
            num_gc = child.get('num_grandchildren', 0)
            lbl   = (f'子{j + 1}〜子{n}'
                     if (n > SHOW and j == SHOW - 1) else f'子{j + 1}')
            els.append(vl(cx, mid_y, child_y - BH // 2))
            els.append(box(cx, child_y, lbl, dead=not alive))

            if not alive and num_gc > 0:
                SHOW_GC = min(num_gc, 2)
                gc_y    = child_y + 120
                max_y   = max(max_y, gc_y)
                gc_step = 75
                gc_start_x = cx - (SHOW_GC - 1) * gc_step // 2
                gc_mid_y   = (child_y + BH // 2 + gc_y - BH // 2) // 2
                els.append(vl(cx, child_y + BH // 2, gc_mid_y))
                if SHOW_GC > 1:
                    els.append(hl(gc_start_x,
                                  gc_start_x + gc_step * (SHOW_GC - 1),
                                  gc_mid_y))
                for k in range(SHOW_GC):
                    gcx = gc_start_x + k * gc_step
                    els.append(vl(gcx, gc_mid_y, gc_y - BH // 2))
                    label = (f'孫{k + 1}' if num_gc <= SHOW_GC
                             else (f'孫{k + 1}' if k < SHOW_GC - 1
                                   else f'孫{k + 1}〜{num_gc}'))
                    els.append(box(gcx, gc_y, label))

        return wrap(max_y + BH // 2 + 50, els)

    # ─── 子あり ──────────────────────────────────────────────
    if has_children:
        n    = int(num_children)
        SHOW = min(n, 5)

        if has_spouse:
            dec_cx, dec_cy = W // 2 - 75, 100
            sp_cx          = W // 2 + 75
            is_a_sp        = (appl_label == '配偶者')
            els.append(ml(dec_cx + BW // 2, sp_cx - BW // 2, dec_cy))
            els.append(box(dec_cx, dec_cy, '被相続人', dec=True))
            els.append(box(sp_cx,  dec_cy, '配偶者',   appl=is_a_sp))
        else:
            dec_cx, dec_cy = W // 2, 90
            els.append(box(dec_cx, dec_cy, '被相続人', dec=True))

        child_y = 230
        step    = min(130, (W - 60) // max(SHOW, 1))
        start_x = W // 2 - step * (SHOW - 1) // 2
        mid_y   = (dec_cy + BH // 2 + child_y - BH // 2) // 2

        els.append(vl(dec_cx, dec_cy + BH // 2, mid_y))
        if SHOW > 1:
            els.append(hl(start_x, start_x + step * (SHOW - 1), mid_y))
        for j in range(SHOW):
            cx  = start_x + j * step
            lbl = (f'子{j+1}〜子{n}' if (n > SHOW and j == SHOW - 1) else f'子{j+1}')
            is_a_c = (appl_label == f'子{j+1}')
            els.append(vl(cx, mid_y, child_y - BH // 2))
            els.append(box(cx, child_y, lbl, appl=is_a_c))

        return wrap(300, els)

    # ─── 両親とも存命 ─────────────────────────────────────────
    if parents_status == '両親とも存命':
        f_cx, f_cy   = 195, 85
        m_cx, m_cy   = 385, 85
        dec_cx, dec_cy = 285, 215
        sp_cx          = dec_cx + BW + 50

        mid_x = (f_cx + BW // 2 + m_cx - BW // 2) // 2
        els += [
            ml(f_cx + BW // 2, m_cx - BW // 2, f_cy),
            vl(mid_x, f_cy + BH // 2, dec_cy - BH // 2),
            ml(dec_cx + BW // 2, sp_cx - BW // 2, dec_cy),
            box(f_cx,   f_cy,   '父',    appl=(appl_label == '父')),
            box(m_cx,   m_cy,   '母',    appl=(appl_label == '母')),
            box(dec_cx, dec_cy, '被相続人', dec=True),
            box(sp_cx,  dec_cy, '配偶者',  appl=(appl_label == '配偶者')),
        ]
        return wrap(290, els)

    # ─── 親１名 ───────────────────────────────────────────────
    if parents_status == '父または母のみ存命':
        pl             = which_parent or '親'
        p_cx,  p_cy   = 220, 85
        dec_cx, dec_cy = 380, 215
        sp_cx          = dec_cx + BW + 50
        is_a_p  = (appl_label in ('親', pl))
        els += [
            (f'<line x1="{p_cx}" y1="{p_cy + BH // 2}" '
             f'x2="{dec_cx}" y2="{dec_cy - BH // 2}" '
             f'stroke="{LC}" stroke-width="1.5"/>'),
            ml(dec_cx + BW // 2, sp_cx - BW // 2, dec_cy),
            box(p_cx,   p_cy,   pl,        appl=is_a_p),
            box(dec_cx, dec_cy, '被相続人', dec=True),
            box(sp_cx,  dec_cy, '配偶者',   appl=(appl_label == '配偶者')),
        ]
        return wrap(290, els)

    # ─── 兄弟姉妹 ─────────────────────────────────────────────
    if has_siblings:
        dec_cx, dec_cy = 100, 215
        sp_cx          = dec_cx + BW + 50  # 250

        SHOW_S    = min(len(siblings_data), 3)
        sib_start = 390
        sib_step  = 140
        sib_cxs   = [sib_start + j * sib_step for j in range(SHOW_S)]
        last_x    = sib_cxs[-1] if sib_cxs else dec_cx

        # 親（故人）
        mid_scene = (dec_cx + last_x) // 2
        f_cx = mid_scene - 85
        m_cx = mid_scene + 85
        f_cy = m_cy = 75
        bar_y    = (f_cy + BH // 2 + dec_cy - BH // 2) // 2
        mid_par  = (f_cx + BW // 2 + m_cx - BW // 2) // 2

        els.append(ml(f_cx + BW // 2, m_cx - BW // 2, f_cy))
        els.append(box(f_cx, f_cy, '父', dead=True))
        els.append(box(m_cx, m_cy, '母', dead=True))

        # 横棒（親→子供たち）
        all_cx = [dec_cx] + sib_cxs
        els.append(vl(mid_par, f_cy + BH // 2, bar_y))
        if len(all_cx) > 1:
            els.append(hl(all_cx[0], all_cx[-1], bar_y))
        for cx in all_cx:
            els.append(vl(cx, bar_y, dec_cy - BH // 2))

        # 被相続人
        els.append(box(dec_cx, dec_cy, '被相続人', dec=True))

        # 配偶者
        if has_spouse:
            is_a_sp = (appl_label == '配偶者')
            els.append(ml(dec_cx + BW // 2, sp_cx - BW // 2, dec_cy))
            els.append(box(sp_cx, dec_cy, '配偶者', appl=is_a_sp))

        # 兄弟姉妹 ＋ 甥姪
        max_y = dec_cy
        for j, scx in enumerate(sib_cxs):
            sib   = siblings_data[j]
            alive = sib.get('alive', True)
            nc    = sib.get('num_children', 0)
            is_a_s = (appl_label in (
                f'兄弟姉妹{j+1}', f'兄弟姉妹{j+1}（生存）'))
            els.append(box(scx, dec_cy, f'兄弟姉妹{j+1}', appl=is_a_s, dead=not alive))

            if not alive and nc > 0:
                SHOW_N  = min(nc, 2)
                niece_y = dec_cy + 115
                max_y   = max(max_y, niece_y)
                n_step  = 80
                n_start = scx - (SHOW_N - 1) * n_step // 2
                n_mid   = (dec_cy + BH // 2 + niece_y - BH // 2) // 2
                els.append(vl(scx, dec_cy + BH // 2, n_mid))
                if SHOW_N > 1:
                    els.append(hl(n_start, n_start + n_step * (SHOW_N - 1), n_mid))
                for k in range(SHOW_N):
                    ncx   = n_start + k * n_step
                    is_a_n = (appl_label == f'兄弟姉妹{j+1}の子{k+1}（甥・姪）')
                    els.append(vl(ncx, n_mid, niece_y - BH // 2))
                    els.append(box(ncx, niece_y, f'甥・姪{k+1}', appl=is_a_n))

        extra = len(siblings_data) - SHOW_S
        if extra > 0:
            els.append(
                f'<text x="{last_x + 70}" y="{dec_cy + 6}" font-size="12" '
                f'fill="#6B5E4A" font-family="sans-serif">…他{extra}名</text>'
            )

        return wrap(max_y + BH // 2 + 55, els)

    # Fallback
    els.append(box(W // 2, 100, '被相続人', dec=True))
    return wrap(200, els)


import generators.spouse_children as spouse_children
import generators.children_only as children_only
import generators.spouse_parent1 as spouse_parent1
import generators.spouse_parent2 as spouse_parent2
import generators.spouse_siblings as spouse_siblings
import generators.siblings_only as siblings_only
import generators.daishuu as daishuu
import generators.daishuu_siblings as daishuu_siblings
import generators.half_siblings as half_siblings

GENERATOR_MAP = {
    'spouse_children': spouse_children,
    'children_only':   children_only,
    'spouse_parent1':  spouse_parent1,
    'spouse_parent2':  spouse_parent2,
    'spouse_siblings': spouse_siblings,
    'siblings_only':   siblings_only,
    'daishuu':          daishuu,
    'daishuu_siblings': daishuu_siblings,
    'half_siblings':   half_siblings,
}

st.set_page_config(page_title='法定相続情報一覧図 作成ツール', layout='centered')

# ── リセット用カウンター（全ウィジェットのキーに付与して確実リセット）────
if 'run_id' not in st.session_state:
    st.session_state.run_id = 0
_k = st.session_state.run_id

# ── カスタムCSS ────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+JP:wght@400;600;700&family=Noto+Sans+JP:wght@400;500;700&display=swap');

/* ── 全体 ── */
.stApp {
    background-color: #e9e4d8;
    background-image: radial-gradient(circle at 50% -100px, #f4efe4, #e9e4d8 65%);
    font-family: 'Noto Sans JP', 'Hiragino Kaku Gothic ProN', sans-serif;
    overflow-x: hidden;
}
.block-container {
    padding: 0 1rem 3rem 1rem !important;
    max-width: 760px;
}

/* ── ヘッダー（画面幅いっぱいに拡張） ── */
.app-header {
    background: linear-gradient(180deg, #232f47 0%, #1b2436 100%);
    border-bottom: 3px solid #a9853f;
    box-shadow: 0 3px 16px rgba(0,0,0,.28);
    padding: 2.2rem 2.5rem 1.8rem;
    margin-bottom: 1.6rem;
    margin-left: calc(50% - 50vw);
    width: 100vw;
    box-sizing: border-box;
    text-align: center;
}
.app-header-label {
    color: #a9853f;
    font-size: 0.72rem;
    font-weight: 700;
    letter-spacing: 3px;
    text-transform: uppercase;
    margin-bottom: 0.5rem;
}
.app-header h1 {
    color: #f5f0e6 !important;
    font-family: 'Noto Serif JP', serif !important;
    font-size: 1.75rem !important;
    font-weight: 600 !important;
    margin: 0 0 0 0 !important;
    letter-spacing: 0.1em;
    line-height: 1.3 !important;
    text-shadow: 0 1px 2px rgba(0,0,0,.3);
}
.app-header p {
    color: #cabf9f !important;
    font-size: 0.83rem !important;
    margin: 0 !important;
    letter-spacing: 0.02em;
    line-height: 1.8 !important;
}
.app-header-rule {
    width: 60px;
    height: 2px;
    background: #a9853f;
    margin: 14px auto 14px;
    border: none;
}

/* ── セクション見出し ── */
h2, [data-testid="stSubheader"] > div {
    color: #1b2436 !important;
    font-family: 'Noto Serif JP', serif !important;
    font-size: 1.05rem !important;
    font-weight: 700 !important;
    letter-spacing: 0.05em !important;
    background: #fffdf8 !important;
    padding: 0.65rem 1rem 0.65rem 1.1rem !important;
    border: 1px solid #cfc7b4 !important;
    border-top: 3px solid #1b2436 !important;
    border-radius: 2px !important;
    margin-top: 1rem !important;
    margin-bottom: 0.8rem !important;
    box-shadow: 0 1px 6px rgba(40,36,32,.09) !important;
}

/* ── 入力エリア ── */
.stRadio, .stCheckbox, .stNumberInput, .stSelectbox {
    background-color: #fffdf8;
    padding: 0.3rem 0.8rem;
    border-radius: 2px;
    border: 1px solid #e8e2d6;
    margin-bottom: 2px;
}
.stRadio label, .stCheckbox label {
    color: #262420 !important;
    font-weight: 400 !important;
}

/* ── 生成ボタン ── */
.stButton > button[kind="primary"] {
    background: linear-gradient(180deg, #2c3a57 0%, #1b2436 100%) !important;
    color: #f5f0e6 !important;
    border: 1px solid #1b2436 !important;
    border-radius: 2px !important;
    padding: 1rem 2rem !important;
    font-size: 1.05rem !important;
    font-weight: 700 !important;
    font-family: 'Noto Serif JP', serif !important;
    letter-spacing: 0.08em !important;
    width: 100% !important;
    box-shadow: 0 3px 12px rgba(27,36,54,.32) !important;
    transition: border-color .2s ease, box-shadow .2s ease !important;
}
.stButton > button[kind="primary"]:hover,
.stButton > button[kind="primary"]:focus,
.stButton > button[kind="primary"]:active {
    background: linear-gradient(180deg, #2c3a57 0%, #1b2436 100%) !important;
    color: #f5f0e6 !important;
    border-color: #a9853f !important;
    box-shadow: 0 6px 18px rgba(27,36,54,.42) !important;
}

/* ── ダウンロードボタン ── */
.stDownloadButton > button {
    background: linear-gradient(180deg, #2c3a57, #1b2436) !important;
    color: #f5f0e6 !important;
    border: 1px solid #a9853f !important;
    border-radius: 2px !important;
    width: 100% !important;
    font-size: 1.05rem !important;
    font-weight: 700 !important;
    font-family: 'Noto Serif JP', serif !important;
    letter-spacing: 0.05em !important;
    padding: 0.75rem !important;
    box-shadow: 0 3px 12px rgba(27,36,54,.3) !important;
}

/* ── リセットボタン ── */
.stButton > button[kind="secondary"] {
    background: transparent !important;
    color: #6c6555 !important;
    border: 1px solid #cfc7b4 !important;
    border-radius: 2px !important;
    font-size: 0.8rem !important;
    font-weight: 500 !important;
    padding: 0.3rem 0.9rem !important;
    width: auto !important;
    box-shadow: none !important;
    letter-spacing: 0.3px !important;
}
.stButton > button[kind="secondary"]:hover {
    background: rgba(27,36,54,.06) !important;
    border-color: #1b2436 !important;
    color: #1b2436 !important;
    box-shadow: none !important;
}

/* ── 区切り線 ── */
hr {
    border: none !important;
    border-top: 1px solid #cfc7b4 !important;
    margin: 1.2rem 0 !important;
}

/* ── キャプション・リンク ── */
.stCaption { color: #6c6555 !important; font-size: 0.82rem !important; }
.stMarkdown a { color: #a9853f !important; }
.stMarkdown a:hover { color: #1b2436 !important; }

/* ── セレクトボックス ── */
.stSelectbox > div > div { border-radius: 2px !important; }

/* ── アニメーション定義 ── */
@keyframes fadeInDown {
    from { opacity: 0; transform: translateY(-14px); }
    to   { opacity: 1; transform: translateY(0); }
}
@keyframes fadeInUp {
    from { opacity: 0; transform: translateY(10px); }
    to   { opacity: 1; transform: translateY(0); }
}
@keyframes expandRule {
    from { width: 0; opacity: 0; }
    to   { width: 60px; opacity: 1; }
}
@keyframes subtleGlow {
    0%, 100% { box-shadow: 0 3px 12px rgba(27,36,54,.32); }
    50%       { box-shadow: 0 4px 22px rgba(169,133,63,.5),
                            0 0 0 3px rgba(169,133,63,.18); }
}

/* ── ヘッダーアニメーション ── */
.app-header-label {
    animation: fadeInDown 0.55s ease both;
}
.app-header h1 {
    animation: fadeInUp 0.65s ease 0.2s both;
}
.app-header-rule {
    animation: expandRule 0.6s ease 0.5s both;
}
.app-header p {
    animation: fadeInUp 0.6s ease 0.65s both;
}

/* ── 生成ボタン：控えめなゴールドグロー脈動 ── */
.stButton > button[kind="primary"] {
    animation: subtleGlow 3.5s ease-in-out infinite !important;
}
.stButton > button[kind="primary"]:hover,
.stButton > button[kind="primary"]:focus,
.stButton > button[kind="primary"]:active {
    animation: none !important;
}

/* ── ラジオボタン → トグルボタン風 ── */

/* radio-group を横並びボタン帯に */
[data-testid="stRadio"] [data-baseweb="radio-group"] {
    display: inline-flex !important;
    flex-wrap: wrap !important;
    border: 1px solid #1b2436 !important;
    border-radius: 2px !important;
    overflow: hidden !important;
    gap: 0 !important;
    background: #fffdf8 !important;
}

/* 各ラジオ項目 */
[data-testid="stRadio"] [data-baseweb="radio"] {
    display: flex !important;
    align-items: center !important;
    padding: 7px 18px !important;
    margin: 0 !important;
    cursor: pointer !important;
    background: #fffdf8 !important;
    border-right: 1px solid #1b2436 !important;
    transition: background 0.15s, color 0.15s !important;
    gap: 0 !important;
    min-height: unset !important;
}
[data-testid="stRadio"] [data-baseweb="radio"]:last-child {
    border-right: none !important;
}

/* 丸いインジケーターを非表示 */
[data-testid="stRadio"] [data-baseweb="radio"] > div:first-child {
    display: none !important;
}

/* ラジオのテキスト */
[data-testid="stRadio"] [data-baseweb="radio"] p {
    font-size: 0.88rem !important;
    color: #1b2436 !important;
    margin: 0 !important;
    line-height: 1 !important;
}

/* 選択中: ネイビー塗りつぶし */
[data-testid="stRadio"] [data-baseweb="radio"]:has(input:checked) {
    background: #1b2436 !important;
}
[data-testid="stRadio"] [data-baseweb="radio"]:has(input:checked) p {
    color: #f5f0e6 !important;
}

/* ラジオ全体の余白調整 */
[data-testid="stRadio"] {
    background-color: transparent !important;
    border: none !important;
    padding: 0.2rem 0.8rem !important;
}
[data-testid="stRadio"] > label {
    font-size: 0.9rem !important;
    color: #262420 !important;
    margin-bottom: 6px !important;
}
</style>
""", unsafe_allow_html=True)

# ── ヘッダー ──────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
    <div class="app-header-label">Legal Document Tool</div>
    <h1>法定相続情報一覧図　作成ツール</h1>
    <div class="app-header-rule"></div>
    <p>法務局の法定相続情報証明制度に対応したExcelファイルを自動生成します。<br>
    氏名・住所・日付等は出力後に手入力してください。</p>
</div>
""", unsafe_allow_html=True)

# ── 参考リンク ───────────────────────────────────────────────
st.markdown(
    '📎 参考：[法務局「法定相続情報証明制度について」]'
    '(https://houmukyoku.moj.go.jp/homu/page7_000014.html)',
)

# ── 個人情報の取り扱いに関するお願い ──────────────────────────
st.markdown(
    '<div style="background:#eef3f8;border:1px solid #b9c7d8;'
    'border-left:5px solid #1b2436;border-radius:4px;'
    'padding:0.85rem 1.1rem;margin:0.8rem 0;line-height:1.7;">'
    '<strong>🔒 個人情報の入力について（お願い）</strong><br>'
    'このツールでは、<strong>氏名・住所・生年月日などの個人情報は入力しないでください。</strong><br>'
    'これらの情報は、ファイルを<strong>ダウンロードした後</strong>、'
    'お手元のExcel（黄色のセル）にご自身で入力してください。<br>'
    '<span style="font-size:0.85rem;color:#6c6555;">'
    '※ 本ツールは空欄のひな形を作成するもので、入力内容を保存・送信することはありません。</span>'
    '</div>',
    unsafe_allow_html=True,
)

# ── リセットボタン ────────────────────────────────────────────
_sp, _rst = st.columns([5, 1])
with _rst:
    if st.button('↺ リセット', key=f'reset_btn_{_k}',
                 help='入力内容をすべてクリアして最初からやり直します'):
        _new_k = _k + 1
        st.session_state.clear()
        st.session_state.run_id = _new_k
        st.rerun()

# ── ① 配偶者 ────────────────────────────────────────────────
st.subheader('① 配偶者について')
has_spouse = st.radio(
    '被相続人に配偶者はいますか？',
    ['はい', 'いいえ'],
    horizontal=True,
    key=f'has_spouse_{_k}',
) == 'はい'

# ── ② 子 ────────────────────────────────────────────────────
st.subheader('② 子について')
has_children = st.radio(
    '被相続人に子（実子・養子を含む）はいますか？',
    ['はい', 'いいえ'],
    horizontal=True,
    key=f'has_children_{_k}',
) == 'はい'

# 子あり → 人数 ＋ 各子の状態
num_children  = 1
has_daishuu   = False
children_data = []

if has_children:
    num_children = st.number_input(
        '子は何人いますか？',
        min_value=1, max_value=99, value=1, step=1,
        key=f'num_children_{_k}',
    )
    st.caption('各お子さまの状態を入力してください。')
    for _i in range(1, int(num_children) + 1):
        _cols = st.columns([3, 3])
        with _cols[0]:
            _child_status = st.radio(
                f'子{_i}',
                ['生存', '死亡（代襲相続あり）'],
                horizontal=True,
                key=f'child_{_i}_status_{_k}',
            )
        _alive = (_child_status == '生存')
        _num_gc = 0
        if not _alive:
            with _cols[1]:
                _num_gc = st.number_input(
                    f'孫（子{_i}の子）の人数',
                    min_value=1, max_value=99, value=1, step=1,
                    key=f'child_{_i}_gc_{_k}',
                )
        children_data.append({'alive': _alive, 'num_grandchildren': int(_num_gc)})
    has_daishuu = any(not c['alive'] for c in children_data)

# ── ③ 親（子なしの場合のみ表示）────────────────────────────
parents_status = '両親とも死亡'
which_parent   = None

if not has_children:
    st.subheader('③ 親について')
    parents_status = st.radio(
        '被相続人の親（父・母）は存命ですか？',
        ['両親とも存命', '父または母のみ存命', '両親とも死亡'],
        horizontal=False,
        key=f'parents_status_{_k}',
    )
    if parents_status == '父または母のみ存命':
        which_parent = st.radio(
            '存命の親は？',
            ['父', '母'],
            horizontal=True,
            key=f'which_parent_{_k}',
        )

# ── ④ 兄弟姉妹（子なし・親なしの場合のみ表示）──────────────
has_siblings  = False
num_siblings  = 1
has_half_sib  = False
siblings_data = []

if not has_children and parents_status == '両親とも死亡':
    st.subheader('④ 兄弟姉妹について')
    has_siblings = st.radio(
        '被相続人に兄弟姉妹はいますか？',
        ['はい', 'いいえ'],
        horizontal=True,
        key=f'has_siblings_{_k}',
    ) == 'はい'
    if has_siblings:
        num_siblings = st.number_input(
            '兄弟姉妹は何人いますか？',
            min_value=1, max_value=99, value=1, step=1,
            key=f'num_siblings_{_k}',
        )
        has_half_sib = st.checkbox(
            '父母の一方のみを同じくする兄弟姉妹（半血兄弟姉妹）がいる',
            key=f'has_half_sib_{_k}',
        )
        st.caption('各兄弟姉妹の状態を入力してください。')
        siblings_data = []
        for i in range(1, int(num_siblings) + 1):
            cols = st.columns([3, 3])
            with cols[0]:
                sib_status = st.radio(
                    f'兄弟姉妹{i}',
                    ['生存', '死亡（代襲相続あり）', '死亡（代襲相続なし）'],
                    horizontal=True,
                    key=f'sib_{i}_alive_{_k}',
                )
            alive = (sib_status == '生存')
            num_nieces = 0
            if sib_status == '死亡（代襲相続あり）':
                with cols[1]:
                    num_nieces = st.number_input(
                        f'子（甥・姪）の人数',
                        min_value=1, max_value=99, value=1, step=1,
                        key=f'sib_{i}_nieces_{_k}',
                    )
            siblings_data.append({'alive': alive, 'num_children': int(num_nieces)})

# ── 申出人の選択（代襲相続以外の全ケースで表示）──────────────
appl_start = None
appl_end   = None
appl_label = None

if not has_daishuu:
    candidates = _compute_candidates(
        has_spouse, has_children, num_children,
        parents_status, has_siblings, siblings_data,
    )
    if len(candidates) >= 2:
        st.divider()
        st.subheader('申出人について')
        labels = [c[0] for c in candidates]
        selected = st.selectbox(
            '申出人（法定相続情報の申出を行う相続人）はどなたですか？',
            labels,
            key=f'applicant_sel_{_k}',
        )
        _, appl_start, appl_end = next(c for c in candidates if c[0] == selected)
        appl_label = selected
    elif len(candidates) == 1:
        # 候補が1人だけの場合は自動確定（表示のみ）
        appl_start = candidates[0][1]
        appl_end   = candidates[0][2]
        appl_label = candidates[0][0]

# ── 相続関係プレビュー ────────────────────────────────────────
st.markdown(
    _generate_preview_svg(
        has_spouse, has_children, num_children,
        parents_status, has_siblings, siblings_data,
        which_parent, has_daishuu, appl_label,
        children_data=children_data,
    ),
    unsafe_allow_html=True,
)

# ── 生成ボタン ───────────────────────────────────────────────
st.divider()

if has_daishuu:
    st.markdown(
        '<div style="background:#fff8e1;border:1px solid #f0a500;border-left:5px solid #f0a500;'
        'border-radius:4px;padding:0.8rem 1rem;margin-bottom:0.8rem;">'
        '<strong>⚠️ この書式は現在調整中です</strong><br>'
        '「②の子について 代襲相続が発生している場合」のフォーマットは現在整備中のため、'
        '出力されるファイルは完全ではない場合があります。'
        '</div>',
        unsafe_allow_html=True,
    )

if st.button('フォーマットを生成する', type='primary'):
    answers = {
        'has_spouse':         has_spouse,
        'has_children':       has_children,
        'num_children':       int(num_children),
        'children_data':      children_data,
        'both_parents_alive': parents_status == '両親とも存命',
        'one_parent_alive':   parents_status == '父または母のみ存命',
        'which_parent':       which_parent,
        'has_siblings':       has_siblings,
        'num_siblings':       int(num_siblings),
        'has_half_sibling':   has_half_sib,
        'siblings_data':      siblings_data if has_siblings else [],
        'appl_start':         appl_start,
        'appl_end':           appl_end,
        'has_daishuu':        has_daishuu,
    }

    template_key, params = decide(answers)

    if template_key is None:
        st.error(f'エラー: {params}')
        st.stop()

    try:
        gen = GENERATOR_MAP[template_key]
        wb  = gen.generate(**params)

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        today = date.today()
        file_name = f'{today.strftime("%y%m%d")} 法定相続情報一覧図.xlsx'

        st.download_button(
            label='📥 エクセルファイルをダウンロード',
            data=buf,
            file_name=file_name,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
    except NotImplementedError as e:
        st.warning(f'未実装: {e}')
    except Exception as e:
        st.error(f'生成中にエラーが発生しました: {e}')
