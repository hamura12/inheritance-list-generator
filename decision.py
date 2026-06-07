"""
相続関係の入力情報からどのテンプレートを使うか判定するロジック。
戻り値: (template_key, params) または (None, error_message)
"""

TEMPLATES = {
    'spouse_children':  '配偶者・子（１人～４人まで対応）である場合',
    'children_only':    '子（１人～４人まで対応）である場合',
    'spouse_parent2':   '配偶者・親２名（父及び母）である場合',
    'spouse_parent1':   '配偶者・親１名（父又は母）である場合',
    'spouse_siblings':  '配偶者・兄弟姉妹がいる場合',
    'siblings_only':    '配偶者なし・兄弟姉妹のみの場合',
    'daishuu_siblings': '兄弟姉妹の代襲相続がある場合',
    'daishuu':          '代襲相続が生じている場合',
    'half_siblings':    '父母の一方のみを同じくする兄弟姉妹がいる場合',
}

def decide(answers: dict):
    """
    answers キー:
      has_spouse        : bool
      has_children      : bool
      num_children      : int (1以上, has_children=True のとき)
      both_parents_alive: bool | None
      one_parent_alive  : bool | None  (both=False のとき)
      which_parent      : str | None   ('父' or '母')
      has_siblings      : bool | None
      num_siblings      : int (1-3, has_siblings=True のとき)
      has_daishuu       : bool  (代襲相続)
      has_half_sibling  : bool  (半血兄弟姉妹)
      appl_start        : str   (申出人欄の左上セル, 例: 'X24')
      appl_end          : str   (申出人欄の右下セル, 例: 'AA25')

    戻り値: (template_key, params_dict) または (None, error_str)
    """
    has_spouse    = answers.get('has_spouse', False)
    has_children  = answers.get('has_children', False)
    num_children  = answers.get('num_children', 0)
    has_daishuu   = answers.get('has_daishuu', False)
    appl_start    = answers.get('appl_start', None)
    appl_end      = answers.get('appl_end', None)

    # ── 代襲相続（最優先）──────────────────────
    if has_daishuu:
        return 'daishuu', {}

    # ── 子あり ──────────────────────────────────
    if has_children:
        if num_children < 1:
            return None, "子の人数を1人以上入力してください"
        params = {'num_children': num_children}
        if appl_start and appl_end:
            params['appl_start'] = appl_start
            params['appl_end']   = appl_end
        if has_spouse:
            return 'spouse_children', params
        else:
            return 'children_only', params

    # ── 子なし ──────────────────────────────────
    both_parents  = answers.get('both_parents_alive', False)
    one_parent    = answers.get('one_parent_alive', False)
    which_parent  = answers.get('which_parent', None)
    has_siblings  = answers.get('has_siblings', False)
    num_siblings  = answers.get('num_siblings', 0)
    has_half_sib  = answers.get('has_half_sibling', False)

    if both_parents:
        if has_spouse:
            params = {}
            if appl_start and appl_end:
                params['appl_start'] = appl_start
                params['appl_end']   = appl_end
            return 'spouse_parent2', params
        return None, "配偶者なし・親2名の書式は現在未対応です"

    if one_parent:
        if has_spouse:
            params = {'which_parent': which_parent}
            if appl_start and appl_end:
                params['appl_start'] = appl_start
                params['appl_end']   = appl_end
            return 'spouse_parent1', params
        return None, "配偶者なし・親1名の書式は現在未対応です"

    # 親も子もいない → 兄弟姉妹
    if has_siblings:
        if num_siblings < 1:
            return None, "兄弟姉妹の人数を1人以上入力してください"
        siblings_data  = answers.get('siblings_data', [])
        any_deceased   = any(not s['alive'] for s in siblings_data)
        if any_deceased:
            params = {
                'has_spouse':    has_spouse,
                'has_half_sib':  has_half_sib,
                'siblings_data': siblings_data,
            }
            if appl_start and appl_end:
                params['appl_start'] = appl_start
                params['appl_end']   = appl_end
            return 'daishuu_siblings', params
        params = {'num_siblings': num_siblings}
        if appl_start and appl_end:
            params['appl_start'] = appl_start
            params['appl_end']   = appl_end
        if has_half_sib:
            return 'half_siblings', params
        if has_spouse:
            return 'spouse_siblings', params
        return 'siblings_only', params

    return None, "条件に合致するテンプレートが見つかりませんでした"
