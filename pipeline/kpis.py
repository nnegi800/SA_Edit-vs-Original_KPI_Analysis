"""
Compute all KPI metrics from df_edited + df_full.
Returns a structured dict ready for the Streamlit UI and xlsx export.
"""
import difflib
import json
import re
import statistics
from collections import Counter

import pandas as pd

EMOJI_RE = re.compile(
    "[\U0001F600-\U0001F64F"
    "\U0001F300-\U0001F5FF"
    "\U0001F680-\U0001F6FF"
    "\U0001F1E0-\U0001F1FF"
    "\U00002600-\U000027BF"
    "\U0001F900-\U0001F9FF"
    "\U0001FA00-\U0001FA6F"
    "\U0001FA70-\U0001FAFF"
    "\U00002702-\U000027B0]+",
    flags=re.UNICODE,
)

USE_CASES = [
    'ai360_action_plan_emotional_bonding',
    'ai360_action_plan_task',
    'ai360_action_plan_product_storytelling',
]
UC_LABELS = {
    'ai360_action_plan_emotional_bonding':    'Emotional Bonding',
    'ai360_action_plan_task':                 'Task',
    'ai360_action_plan_product_storytelling': 'Product Storytelling',
}


def _added_segments(a, b):
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    return "".join(b[j1:j2] for tag, i1, i2, j1, j2 in sm.get_opcodes() if tag in ('insert', 'replace'))


def _removed_segments(a, b):
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    return "".join(a[i1:i2] for tag, i1, i2, j1, j2 in sm.get_opcodes() if tag in ('delete', 'replace'))


def _classify(rec):
    summary   = str(rec.get('change_summary', '') or '')
    ai_en     = str(rec.get('ai_message_translated',  '') or '')
    sent_en   = str(rec.get('sent_message_translated', '') or '')
    ai_orig   = str(rec.get('ai_message_original',    '') or '')
    sent_orig = str(rec.get('sent_message_original',  '') or '')

    added_en   = _added_segments(ai_en, sent_en)
    removed_en = _removed_segments(ai_en, sent_en)
    added_orig = _added_segments(ai_orig, sent_orig)
    removed_orig = _removed_segments(ai_orig, sent_orig)
    diff_orig_all = added_orig + removed_orig

    ai_wc   = len(ai_orig)
    sent_wc = len(sent_orig)
    pct = ((sent_wc - ai_wc) / ai_wc * 100) if ai_wc else 0
    if pct <= -25:
        length = 'Shortened'
    elif pct >= 25:
        length = 'Expanded'
    elif abs(pct) >= 10:
        length = 'Moderate'
    else:
        length = 'Minor'

    loc_line = ""
    for line in summary.splitlines():
        if line.startswith("Location:"):
            loc_line = line.lower()
    in_opening = 'opening' in loc_line
    in_middle  = 'middle'  in loc_line
    in_closing = 'closing' in loc_line
    if 'entire' in loc_line or 'whole' in loc_line:
        in_opening = in_middle = in_closing = True

    emoji_added   = bool(EMOJI_RE.search(added_en)) or bool(EMOJI_RE.search(added_orig))
    emoji_removed = bool(EMOJI_RE.search(removed_en))

    CN_GREET    = ['亲爱', '您好', '你好', '小姐', '姐姐', '女士', '好久不见',
                   '先生', '太太', '早上好', '下午好', '晚上好', '嗨', '哈喽', '久违']
    CN_CTA      = ['有空', '看看', '联系', '随时', '过来', '试试', '试戴', '带来', '方便',
                   '来店', '到店', '欢迎', '预约', '安排', '拜访', '有机会', '欢迎光临']
    CN_PRODUCT  = ['珠宝', '卡地亚', '系列', '腕表', '戒指', 'LOVE', 'Love',
                   '猎豹', '玫瑰', '钻石', 'Juste', 'Clou', '手镯', '项链',
                   '手表', '吊坠', '耳环', '胸针', '蓝气球', '三环', 'Trinity',
                   '钉子', 'Santos', 'Tank', 'Clash', '宝石', '项圈']
    CN_SCARCITY = ['入手', '购买', '现在', '库存', '现货', '稀缺', '有限', '紧张', '难得', '最后']
    CN_CARE     = ['清洗', '服务', '清洁', '保养', '免费', '护理', '每年',
                   '维修', '抛光', '售后', '养护', '定期']
    CN_SEASON   = ['下午', '下个月', '去年', '上次',
                   '春', '夏', '秋', '冬', '节日', '新年', '圣诞', '礼物', '礼品',
                   '情人节', '母亲节', '七夕', '节庆', '假期', '生日', '周年']

    return {
        'use_case':          rec.get('use_case'),
        'length':            length,
        'in_opening':        in_opening,
        'in_middle':         in_middle,
        'in_closing':        in_closing,
        'emoji_added':       emoji_added,
        'emoji_removed':     emoji_removed,
        'cta_modified':      any(w in diff_orig_all for w in CN_CTA),
        'greeting_mod':      any(w in diff_orig_all for w in CN_GREET),
        'product_modified':  any(w in diff_orig_all for w in CN_PRODUCT),
        'scarcity':          any(w in diff_orig_all for w in CN_SCARCITY),
        'aftersales':        any(w in diff_orig_all for w in CN_CARE),
        'seasonal':          any(w in diff_orig_all for w in CN_SEASON),
        'ai_chars':          ai_wc,
        'sent_chars':        sent_wc,
    }


def _agg(subset):
    n = len(subset)
    if n == 0:
        return {}

    def pct(flag):
        return round(sum(1 for m in subset if m[flag]) / n * 100, 1)

    def pct_count(*buckets):
        return round(sum(1 for m in subset if m['length'] in buckets) / n * 100, 1)

    def avg_chars(key, bucket=None):
        rows = [m for m in subset if bucket is None or m['length'] == bucket]
        vals = [m[key] for m in rows if m[key] is not None]
        return round(sum(vals) / len(vals)) if vals else '—'

    return {
        'n':                        n,
        'pct_shortened':            f"{pct_count('Shortened')}%",
        'pct_expanded':             f"{pct_count('Expanded')}%",
        'pct_no_change':            f"{pct_count('Minor')}%",
        'avg_ai_chars':             avg_chars('ai_chars'),
        'avg_sent_chars':           avg_chars('sent_chars'),
        'avg_ai_chars_shortened':   avg_chars('ai_chars',  'Shortened'),
        'avg_sent_chars_shortened': avg_chars('sent_chars','Shortened'),
        'avg_ai_chars_expanded':    avg_chars('ai_chars',  'Expanded'),
        'avg_sent_chars_expanded':  avg_chars('sent_chars','Expanded'),
        'pct_opening':              f"{pct('in_opening')}%",
        'pct_middle':               f"{pct('in_middle')}%",
        'pct_closing':              f"{pct('in_closing')}%",
        'pct_emoji_added':          f"{pct('emoji_added')}%",
        'pct_emoji_removed':        f"{pct('emoji_removed')}%",
        'pct_cta':                  f"{pct('cta_modified')}%",
        'pct_greeting':             f"{pct('greeting_mod')}%",
        'pct_product_modified':     f"{pct('product_modified')}%",
        'pct_scarcity':             f"{pct('scarcity')}%",
        'pct_aftersales':           f"{pct('aftersales')}%",
        'pct_seasonal':             f"{pct('seasonal')}%",
    }


def _sa_distribution_kpis(sa_total, sa_edited, sa_total_uc, sa_edited_uc, uc=None, min_msgs=0):
    rates = []
    for sid in sa_total:
        if uc is None:
            total  = sa_total[sid]
            edited = sa_edited[sid]
        else:
            total  = sa_total_uc.get((sid, uc), 0)
            edited = sa_edited_uc.get((sid, uc), 0)
        if total >= max(min_msgs, 1):
            rates.append(edited / total * 100)
    n = len(rates)
    if n == 0:
        return {}
    mean_rate   = round(statistics.mean(rates), 1)
    median_rate = round(statistics.median(rates), 1)
    b0   = round(sum(1 for r in rates if r == 0)        / n * 100, 1)
    b1   = round(sum(1 for r in rates if 0 < r <= 20)   / n * 100, 1)
    b2   = round(sum(1 for r in rates if 20 < r <= 75)  / n * 100, 1)
    b3   = round(sum(1 for r in rates if r > 75)        / n * 100, 1)
    return {
        'sa_count':     n,
        'mean_sa_rate': f"{mean_rate}%",
        'median_sa_rate': f"{median_rate}%",
        'bucket_0':     f"{b0}%  ({sum(1 for r in rates if r == 0)} SAs)",
        'bucket_1_20':  f"{b1}%  ({sum(1 for r in rates if 0 < r <= 20)} SAs)",
        'bucket_20_75': f"{b2}%  ({sum(1 for r in rates if 20 < r <= 75)} SAs)",
        'bucket_75plus':f"{b3}%  ({sum(1 for r in rates if r > 75)} SAs)",
    }


def run(df_edited: pd.DataFrame, df_full: pd.DataFrame) -> dict:
    """
    Returns
    -------
    {
      'overall':  { kpi_key: value, ... },
      'Emotional Bonding':  { ... },
      'Task':               { ... },
      'Product Storytelling': { ... },
      'sections': [ (section_label, [ (kpi_label, kpi_key), ... ]), ... ],
      'median_msg_threshold': int,
      'sa_table': DataFrame  (per-SA: seller_id, total_msgs, edited_msgs, edit_rate_pct)
    }
    """
    # ── per-SA counters from the full raw dataset ──
    sa_total     = Counter()
    sa_edited    = Counter()
    sa_total_uc  = {}
    sa_edited_uc = {}
    total_by_uc  = Counter()
    edited_by_uc = Counter()

    for _, row in df_full.iterrows():
        uc  = row.get('use_case')
        sid = row.get('seller_id')
        cth = row.get('copy_to_chat_history')
        total_by_uc[uc] += 1
        sa_total[sid]   += 1
        sa_total_uc[(sid, uc)] = sa_total_uc.get((sid, uc), 0) + 1
        if cth:
            try:
                d  = json.loads(cth)
                lv = d.get(d.get('latest_version', 'v1'), {})
                if lv.get('is_change'):
                    edited_by_uc[uc]          += 1
                    sa_edited[sid]            += 1
                    sa_edited_uc[(sid, uc)] = sa_edited_uc.get((sid, uc), 0) + 1
            except Exception:
                pass

    total_all  = sum(total_by_uc.values())
    edited_all = sum(edited_by_uc.values())

    # ── per-row classification ──
    records = df_edited.to_dict('records')
    metrics = [_classify(r) for r in records]

    overall = _agg(metrics)
    by_uc   = {uc: _agg([m for m in metrics if m['use_case'] == uc]) for uc in USE_CASES}

    # ── edit rate (uses full-dataset totals) ──
    def edit_rate_fields(uc=None):
        total  = total_by_uc.get(uc, 0) if uc else total_all
        edited = edited_by_uc.get(uc, 0) if uc else edited_all
        rate   = round(edited / total * 100, 1) if total else 0
        return {'total_generated': total, 'total_edited': edited,
                'edit_rate': f"{rate}%  ({edited} / {total})"}

    overall.update(edit_rate_fields())
    for uc in USE_CASES:
        by_uc[uc].update(edit_rate_fields(uc))

    # ── SA distribution ──
    all_msg_counts       = list(sa_total.values())
    median_msg_threshold = int(statistics.median(all_msg_counts)) if all_msg_counts else 1

    sa_dist_overall     = _sa_distribution_kpis(sa_total, sa_edited, sa_total_uc, sa_edited_uc)
    sa_dist_filtered    = _sa_distribution_kpis(sa_total, sa_edited, sa_total_uc, sa_edited_uc, min_msgs=median_msg_threshold)
    sa_dist_by_uc       = {uc: _sa_distribution_kpis(sa_total, sa_edited, sa_total_uc, sa_edited_uc, uc) for uc in USE_CASES}
    sa_dist_filtered_uc = {uc: _sa_distribution_kpis(sa_total, sa_edited, sa_total_uc, sa_edited_uc, uc, median_msg_threshold) for uc in USE_CASES}

    overall.update(sa_dist_overall)
    for uc in USE_CASES:
        by_uc[uc].update(sa_dist_by_uc[uc])

    for k, v in sa_dist_filtered.items():
        overall[f'f_{k}'] = v
    for uc in USE_CASES:
        for k, v in sa_dist_filtered_uc[uc].items():
            by_uc[uc][f'f_{k}'] = v

    # ── SA volume (overall only) ──
    counts_desc = sorted(all_msg_counts, reverse=True)
    n_sa        = len(counts_desc)
    total_msgs  = sum(counts_desc)
    low_vol_n   = sum(1 for x in counts_desc if 1 <= x <= 5)
    top5_msgs   = sum(counts_desc[:5]) if len(counts_desc) >= 5 else sum(counts_desc)
    top5_pct    = round(top5_msgs / total_msgs * 100, 1) if total_msgs else 0
    top5_min    = counts_desc[4] if len(counts_desc) >= 5 else counts_desc[-1]
    top5_max    = counts_desc[0] if counts_desc else 0
    mean_msgs   = round(statistics.mean(counts_desc), 1) if counts_desc else 0
    median_msgs = round(statistics.median(counts_desc), 1) if counts_desc else 0
    ratio       = round(mean_msgs / median_msgs, 1) if median_msgs else 0

    overall.update({
        'vol_low_volume': f"{round(low_vol_n / n_sa * 100, 1) if n_sa else 0}%  ({low_vol_n} of {n_sa} SAs)",
        'vol_top5':       f"Top 5 SAs → {top5_pct}% of all messages  ({top5_min}–{top5_max} msgs each)",
        'vol_mean':       str(mean_msgs),
        'vol_median':     str(median_msgs),
        'vol_ratio':      f"Mean is {ratio}x the median",
    })

    # ── per-SA evidence table ──
    sa_rows = []
    for sid in sa_total:
        sa_rows.append({
            'seller_id':     sid,
            'total_msgs':    sa_total[sid],
            'edited_msgs':   sa_edited[sid],
            'edit_rate_pct': round(sa_edited[sid] / sa_total[sid] * 100, 1) if sa_total[sid] else 0,
        })
    sa_table = pd.DataFrame(sa_rows).sort_values('total_msgs', ascending=False).reset_index(drop=True)

    # ── sections definition (label, key) ──
    sections = [
        ("Edit Rate", [
            ("Total AI-generated messages",           'total_generated'),
            ("Total SA-edited messages",              'total_edited'),
            ("Edit rate  (edited / total generated)", 'edit_rate'),
        ]),
        ("Volume of Edited Messages", [
            ("Total edited messages", 'n'),
        ]),
        ("Length Change  (Chinese character count)", [
            ("% Shortened  (> 25% fewer characters)",          'pct_shortened'),
            ("Avg characters — AI message (shortened rows)",   'avg_ai_chars_shortened'),
            ("Avg characters — Sent message (shortened rows)", 'avg_sent_chars_shortened'),
            ("% Expanded   (> 25% more characters)",           'pct_expanded'),
            ("Avg characters — AI message (expanded rows)",    'avg_ai_chars_expanded'),
            ("Avg characters — Sent message (expanded rows)",  'avg_sent_chars_expanded'),
            ("% No significant change  (< 10%)",               'pct_no_change'),
            ("Avg characters — AI message (all)",              'avg_ai_chars'),
            ("Avg characters — Sent message (all)",            'avg_sent_chars'),
        ]),
        ("Location of Change", [
            ("% with edits in opening section", 'pct_opening'),
            ("% with edits in middle section",  'pct_middle'),
            ("% with edits in closing section", 'pct_closing'),
        ]),
        ("Content Type", [
            ("% modified greeting / salutation",  'pct_greeting'),
            ("% modified call-to-action",         'pct_cta'),
            ("% modified product references",     'pct_product_modified'),
            ("% added scarcity/urgency language", 'pct_scarcity'),
            ("% added after-sales / care ref",    'pct_aftersales'),
            ("% added seasonal/occasion ref",     'pct_seasonal'),
            ("% added emoji",                     'pct_emoji_added'),
            ("% removed emoji from AI",           'pct_emoji_removed'),
        ]),
        ("SA Distribution", [
            ("Total SAs with AI-generated messages",        'sa_count'),
            ("% of SAs who never edited  (0%)",             'bucket_0'),
            ("% of SAs with low edit rate  (1–20%)",        'bucket_1_20'),
            ("% of SAs with moderate edit rate  (20–75%)",  'bucket_20_75'),
            ("% of SAs with high edit rate  (> 75%)",       'bucket_75plus'),
            ("Mean edit rate per SA",                       'mean_sa_rate'),
            ("Median edit rate per SA",                     'median_sa_rate'),
        ]),
        ("SA Message Volume  (overall only)", [
            ("% of SAs who sent only 1–5 messages",    'vol_low_volume'),
            ("Top 5 SAs — share of total messages",    'vol_top5'),
            ("Mean messages per SA",                   'vol_mean'),
            ("Median messages per SA",                 'vol_median'),
            ("Mean / Median ratio  (skew indicator)",  'vol_ratio'),
        ]),
        (f"SA Distribution  (filtered: ≥ {median_msg_threshold} msgs)", [
            ("SAs included in filtered analysis",                    'f_sa_count'),
            ("% of SAs who never edited  (0%)",                      'f_bucket_0'),
            ("% of SAs with low edit rate  (1–20%)",                 'f_bucket_1_20'),
            ("% of SAs with moderate edit rate  (20–75%)",          'f_bucket_20_75'),
            ("% of SAs with high edit rate  (> 75%)",               'f_bucket_75plus'),
            ("Mean edit rate per SA  (filtered)",                    'f_mean_sa_rate'),
            ("Median edit rate per SA  (filtered)",                  'f_median_sa_rate'),
        ]),
    ]

    return {
        'overall':              overall,
        UC_LABELS[USE_CASES[0]]: by_uc[USE_CASES[0]],
        UC_LABELS[USE_CASES[1]]: by_uc[USE_CASES[1]],
        UC_LABELS[USE_CASES[2]]: by_uc[USE_CASES[2]],
        'sections':             sections,
        'median_msg_threshold': median_msg_threshold,
        'sa_table':             sa_table,
    }
