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

# ── カスタムCSS ────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+JP:wght@400;700&family=Noto+Sans+JP:wght@400;500;700&display=swap');

/* 全体背景・フォント */
.stApp {
    background-color: #C8B89A;
    font-family: 'Noto Sans JP', sans-serif;
}

/* メインコンテナ */
.block-container {
    padding-top: 0 !important;
    padding-bottom: 3rem;
    max-width: 760px;
}

/* ── ヘッダー ── */
.app-header {
    background: linear-gradient(160deg, #0D1B33 0%, #1A3056 60%, #0D1B33 100%);
    padding: 2.2rem 2.5rem 2rem 2.5rem;
    margin-bottom: 1.8rem;
    border-bottom: 3px solid #B8960C;
    position: relative;
    overflow: hidden;
}
.app-header::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0; bottom: 0;
    background: repeating-linear-gradient(
        45deg,
        transparent,
        transparent 40px,
        rgba(255,255,255,0.015) 40px,
        rgba(255,255,255,0.015) 80px
    );
}
.app-header-label {
    color: #B8960C;
    font-size: 0.72rem;
    font-weight: 700;
    letter-spacing: 3px;
    text-transform: uppercase;
    margin-bottom: 0.4rem;
}
.app-header h1 {
    color: #FFFFFF !important;
    font-family: 'Noto Serif JP', serif !important;
    font-size: 1.7rem !important;
    font-weight: 700 !important;
    margin: 0 0 0.5rem 0 !important;
    letter-spacing: 1px;
    line-height: 1.3 !important;
}
.app-header p {
    color: rgba(255,255,255,0.65) !important;
    font-size: 0.82rem !important;
    margin: 0 !important;
    letter-spacing: 0.3px;
}
.app-header-rule {
    border: none;
    border-top: 1px solid rgba(184,150,12,0.4);
    margin: 1rem 0 0.8rem 0;
}

/* ── セクション見出し ── */
h2, [data-testid="stSubheader"] > div {
    color: #FFFFFF !important;
    font-weight: 700 !important;
    font-size: 0.88rem !important;
    letter-spacing: 1.5px !important;
    background: linear-gradient(90deg, #0D1B33, #1A3056) !important;
    padding: 0.55rem 1rem 0.55rem 1.1rem !important;
    border-left: 4px solid #B8960C !important;
    border-radius: 2px !important;
    margin-top: 0.8rem !important;
    margin-bottom: 0.8rem !important;
    box-shadow: 0 2px 8px rgba(13,27,51,0.2) !important;
}

/* ── 入力エリアの背景 ── */
.stRadio, .stCheckbox, .stNumberInput, .stSelectbox {
    background-color: #B5A282;
    padding: 0.3rem 0.8rem;
    border-radius: 4px;
}

/* ── ラジオ・チェックボックス ── */
.stRadio label, .stCheckbox label {
    color: #1A2A42 !important;
    font-weight: 500 !important;
}

/* ── 生成ボタン ── */
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #0D1B33 0%, #1A3056 100%) !important;
    color: #FFFFFF !important;
    border: 1px solid #B8960C !important;
    border-radius: 3px !important;
    padding: 0.75rem 2rem !important;
    font-size: 0.95rem !important;
    font-weight: 700 !important;
    letter-spacing: 2px !important;
    width: 100% !important;
    box-shadow: 0 4px 16px rgba(13,27,51,0.35), inset 0 1px 0 rgba(255,255,255,0.07) !important;
    transition: all 0.2s ease !important;
}
.stButton > button[kind="primary"]:hover {
    background: linear-gradient(135deg, #162340 0%, #223D6A 100%) !important;
    box-shadow: 0 6px 20px rgba(13,27,51,0.5), 0 0 0 1px #B8960C !important;
}

/* ── ダウンロードボタン ── */
.stDownloadButton > button {
    background: linear-gradient(135deg, #0D1B33, #1A3056) !important;
    color: #FFD166 !important;
    border: 1px solid #B8960C !important;
    border-radius: 3px !important;
    width: 100% !important;
    font-size: 0.95rem !important;
    font-weight: 700 !important;
    letter-spacing: 1.5px !important;
    padding: 0.75rem !important;
    box-shadow: 0 4px 16px rgba(13,27,51,0.35) !important;
}

/* ── 区切り線 ── */
hr {
    border: none !important;
    border-top: 1px solid #C8CDD8 !important;
    margin: 1.2rem 0 !important;
}

/* ── キャプション ── */
.stCaption { color: #6B7A95 !important; font-size: 0.8rem !important; }

/* ── セレクトボックス ── */
.stSelectbox > div > div { border-radius: 3px !important; }
</style>
""", unsafe_allow_html=True)

# ── ヘッダー ──────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
    <div class="app-header-label">Legal Document Tool</div>
    <h1>法定相続情報一覧図　作成ツール</h1>
    <hr class="app-header-rule">
    <p>必要最低限の情報を入力するだけで、法務局提出用のエクセルファイルを自動生成します。</p>
</div>
""", unsafe_allow_html=True)

# ── ① 配偶者 ────────────────────────────────────────────────
st.subheader('① 配偶者について')
has_spouse = st.radio(
    '被相続人に配偶者はいますか？',
    ['はい', 'いいえ'],
    horizontal=True,
    key='has_spouse',
) == 'はい'

# ── ② 子 ────────────────────────────────────────────────────
st.subheader('② 子について')
has_children = st.radio(
    '被相続人に子（実子・養子を含む）はいますか？',
    ['はい', 'いいえ'],
    horizontal=True,
    key='has_children',
) == 'はい'

# 子あり → 人数 ＋ 代襲相続
num_children = 1
has_daishuu  = False

if has_children:
    num_children = st.number_input(
        '子は何人いますか？',
        min_value=1, max_value=99, value=1, step=1,
        key='num_children',
    )
    has_daishuu = st.checkbox(
        '代襲相続が生じている（子が被相続人より先に死亡し、その子＝孫が相続人となる場合）',
        key='has_daishuu',
    )

# ── ③ 親（子なしの場合のみ表示）────────────────────────────
parents_status = '両親とも死亡'
which_parent   = None

if not has_children:
    st.subheader('③ 親について')
    parents_status = st.radio(
        '被相続人の親（父・母）は存命ですか？',
        ['両親とも存命', '父または母のみ存命', '両親とも死亡'],
        horizontal=False,
        key='parents_status',
    )
    if parents_status == '父または母のみ存命':
        which_parent = st.radio(
            '存命の親は？',
            ['父', '母'],
            horizontal=True,
            key='which_parent',
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
        key='has_siblings',
    ) == 'はい'
    if has_siblings:
        num_siblings = st.number_input(
            '兄弟姉妹は何人いますか？',
            min_value=1, max_value=99, value=1, step=1,
            key='num_siblings',
        )
        has_half_sib = st.checkbox(
            '父母の一方のみを同じくする兄弟姉妹（半血兄弟姉妹）がいる',
            key='has_half_sib',
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
                    key=f'sib_{i}_alive',
                )
            alive = (sib_status == '生存')
            num_nieces = 0
            if sib_status == '死亡（代襲相続あり）':
                with cols[1]:
                    num_nieces = st.number_input(
                        f'子（甥・姪）の人数',
                        min_value=1, max_value=99, value=1, step=1,
                        key=f'sib_{i}_nieces',
                    )
            siblings_data.append({'alive': alive, 'num_children': int(num_nieces)})

# ── 申出人の選択（代襲相続以外の全ケースで表示）──────────────
appl_start = None
appl_end   = None

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
            key='applicant_sel',
        )
        _, appl_start, appl_end = next(c for c in candidates if c[0] == selected)
    elif len(candidates) == 1:
        # 候補が1人だけの場合は自動確定（表示のみ）
        appl_start = candidates[0][1]
        appl_end   = candidates[0][2]

# ── 生成ボタン ───────────────────────────────────────────────
st.divider()
if st.button('フォーマットを生成する', type='primary'):
    answers = {
        'has_spouse':         has_spouse,
        'has_children':       has_children,
        'num_children':       int(num_children),
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
