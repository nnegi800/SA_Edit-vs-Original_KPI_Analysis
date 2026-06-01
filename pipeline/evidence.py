"""
Maps each KPI key to a filter function that returns the relevant rows
from df_edited. Used by the Streamlit evidence expander.
"""
import pandas as pd

# Columns shown in the evidence table (keep it focused)
EVIDENCE_COLS = [
    'seller_id',
    'use_case',
    'ai_message_translated',
    'sent_message_translated',
    'change_summary',
]

# SA distribution / volume KPIs — evidence is the SA-level table, not message rows
SA_DIST_KEYS = {
    'sa_count', 'mean_sa_rate', 'median_sa_rate',
    'bucket_0', 'bucket_1_20', 'bucket_20_75', 'bucket_75plus',
    'f_sa_count', 'f_mean_sa_rate', 'f_median_sa_rate',
    'f_bucket_0', 'f_bucket_1_20', 'f_bucket_20_75', 'f_bucket_75plus',
    'vol_low_volume', 'vol_top5', 'vol_mean', 'vol_median', 'vol_ratio',
}

# KPIs with no meaningful row-level evidence
NO_EVIDENCE_KEYS = {
    'avg_ai_chars', 'avg_sent_chars',
    'avg_ai_chars_shortened', 'avg_sent_chars_shortened',
    'avg_ai_chars_expanded', 'avg_sent_chars_expanded',
    'total_generated',  # comes from full dataset, not edited rows
}


def _cs(df, term):
    """Case-insensitive substring match on change_summary."""
    return df[df['change_summary'].str.contains(term, case=False, na=False)]


def _loc(df, zone):
    """Match the Location: line in change_summary for a specific zone."""
    def matches(s):
        for line in str(s).splitlines():
            if line.strip().lower().startswith('location:') and zone in line.lower():
                return True
        return False
    return df[df['change_summary'].apply(matches)]


# Maps kpi_key → callable(df_edited) → filtered DataFrame
FILTERS = {
    # Edit rate
    'total_edited':         lambda df: df,
    'edit_rate':            lambda df: df,
    'n':                    lambda df: df,

    # Length
    'pct_shortened':        lambda df: df[df['change_summary'].str.contains('Shortened', na=False)],
    'pct_expanded':         lambda df: df[df['change_summary'].str.contains('Expanded', na=False)],
    'pct_no_change':        lambda df: df[df['change_summary'].str.contains('Similar length', na=False)],

    # Location
    'pct_opening':          lambda df: _loc(df, 'opening'),
    'pct_middle':           lambda df: _loc(df, 'middle'),
    'pct_closing':          lambda df: _loc(df, 'closing'),

    # Content type
    'pct_greeting':         lambda df: _cs(df, 'greeting'),
    'pct_cta':              lambda df: _cs(df, 'call-to-action'),
    'pct_product_modified': lambda df: _cs(df, 'product'),
    'pct_scarcity':         lambda df: _cs(df, 'scarcity'),
    'pct_aftersales':       lambda df: _cs(df, 'after-sales'),
    'pct_seasonal':         lambda df: _cs(df, 'seasonal'),
    'pct_emoji_added':      lambda df: df[
        df['change_summary'].str.contains('emoji', case=False, na=False) &
        ~df['change_summary'].str.contains('removed emoji', case=False, na=False)
    ],
    'pct_emoji_removed':    lambda df: _cs(df, 'removed emoji'),
}


def get_rows(kpi_key: str, df_edited: pd.DataFrame, use_case: str | None = None) -> pd.DataFrame | None:
    """
    Returns a filtered DataFrame for the evidence expander, or None if the KPI
    has no message-level evidence (SA dist / volume / avg chars).

    Parameters
    ----------
    kpi_key   : the internal key used in kpis.py (e.g. 'pct_opening')
    df_edited : the enriched 181-row DataFrame
    use_case  : optional use_case label (e.g. 'Emotional Bonding') to pre-filter
    """
    if kpi_key in SA_DIST_KEYS or kpi_key in NO_EVIDENCE_KEYS:
        return None

    filter_fn = FILTERS.get(kpi_key)
    if filter_fn is None:
        return None

    df = df_edited.copy()

    # Map the display label back to raw use_case value if needed
    UC_MAP = {
        'Emotional Bonding':    'ai360_action_plan_emotional_bonding',
        'Task':                 'ai360_action_plan_task',
        'Product Storytelling': 'ai360_action_plan_product_storytelling',
    }
    if use_case and use_case in UC_MAP:
        df = df[df['use_case'] == UC_MAP[use_case]]

    result = filter_fn(df)

    # Return only the display columns that exist in the dataframe
    cols = [c for c in EVIDENCE_COLS if c in result.columns]
    return result[cols].reset_index(drop=True)
