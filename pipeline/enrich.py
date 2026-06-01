"""
Add character-level diff highlights and a change_summary column to df_edited.
Accepts and returns a DataFrame (no file I/O).
"""
import difflib
import pandas as pd


def _char_diff(a, b):
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    segs_a, segs_b = [], []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'equal':
            segs_a.append(('equal',  a[i1:i2]))
            segs_b.append(('equal',  b[j1:j2]))
        elif tag == 'delete':
            segs_a.append(('delete', a[i1:i2]))
        elif tag == 'insert':
            segs_b.append(('insert', b[j1:j2]))
        elif tag == 'replace':
            segs_a.append(('delete', a[i1:i2]))
            segs_b.append(('insert', b[j1:j2]))
    return segs_a, segs_b


def _summarise(ai_en, sent_en):
    ai_words   = ai_en.split()
    sent_words = sent_en.split()
    n_ai, n_sent = len(ai_words), len(sent_words)

    removed_chunks, added_chunks = [], []
    sm = difflib.SequenceMatcher(None, ai_words, sent_words, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag in ('delete', 'replace'):
            removed_chunks.append(' '.join(ai_words[i1:i2]))
        if tag in ('insert', 'replace'):
            added_chunks.append(' '.join(sent_words[j1:j2]))

    diff_words = n_sent - n_ai
    pct = round(diff_words / n_ai * 100) if n_ai else 0
    if pct <= -30:
        length_note = f"Significantly shortened ({abs(pct)}% fewer words)"
    elif pct < -10:
        length_note = f"Shortened ({abs(pct)}% fewer words)"
    elif pct >= 30:
        length_note = f"Significantly expanded ({pct}% more words)"
    elif pct > 10:
        length_note = f"Expanded ({pct}% more words)"
    else:
        length_note = f"Similar length ({'+' if diff_words >= 0 else ''}{diff_words} words)"

    n_chars = max(len(ai_en), len(sent_en))
    sm_char = difflib.SequenceMatcher(None, ai_en, sent_en, autojunk=False)
    changed_positions = []
    for tag, i1, i2, j1, j2 in sm_char.get_opcodes():
        if tag != 'equal':
            changed_positions.append((i1 + i2) / 2)

    zones = {'opening': 0, 'middle': 0, 'closing': 0}
    for pos in changed_positions:
        rel = pos / n_chars if n_chars else 0.5
        if rel < 0.30:
            zones['opening'] += 1
        elif rel > 0.70:
            zones['closing'] += 1
        else:
            zones['middle'] += 1

    if all(v == 0 for v in zones.values()):
        zone_note = "whole message rewritten"
    elif sum(1 for v in zones.values() if v > 0) == 3:
        zone_note = "changes throughout entire message"
    elif sum(1 for v in zones.values() if v > 0) == 2:
        parts = [k for k, v in zones.items() if v > 0]
        zone_note = f"changes in {' & '.join(parts)}"
    else:
        dominant = max(zones, key=zones.get)
        zone_note = f"main change in {dominant}"

    content_notes = []
    all_removed = ' '.join(removed_chunks).lower()
    all_added   = ' '.join(added_chunks).lower()

    if any(w in all_removed for w in ['hello', 'dear', 'hi ', 'miss', 'mr', 'ms', 'madam']):
        content_notes.append("modified greeting/salutation")
    if any(w in all_added for w in ['hello', 'dear', 'hi ', 'miss', 'mr', 'ms', 'madam']):
        content_notes.append("personalised salutation")
    if any(w in all_removed for w in ['please', 'come', 'visit', 'welcome', 'contact', 'reach', 'appointment', 'arrange']):
        content_notes.append("modified call-to-action")
    if any(w in all_added for w in ['please', 'come', 'visit', 'welcome', 'contact', 'reach', 'appointment', 'arrange']):
        content_notes.append("added/changed call-to-action")
    if any(w in all_added for w in ['love', 'juste un clou', 'trinity', 'panthère', 'santos', 'tank', 'ballon', 'clash', 'bracelet', 'ring', 'necklace', 'watch']):
        content_notes.append("added product name/series reference")
    if any(w in all_removed for w in ['love', 'juste un clou', 'trinity', 'panthère', 'santos', 'tank', 'ballon', 'clash', 'bracelet', 'ring', 'necklace', 'watch']):
        content_notes.append("removed product reference")
    if any(w in all_added for w in ['stock', 'inventory', 'limited', 'available', 'rare', 'scarce', 'last']):
        content_notes.append("added scarcity/urgency")
    if any(w in all_removed for w in ['stock', 'inventory', 'limited', 'available', 'rare', 'scarce']):
        content_notes.append("removed scarcity language")
    if any(w in all_added for w in ['clean', 'maintain', 'service', 'repair', 'polish', 'care']):
        content_notes.append("added after-sales/care mention")
    if any(w in all_added for w in ['spring', 'summer', 'winter', 'autumn', 'holiday', 'festival', 'new year', 'christmas', 'gift']):
        content_notes.append("added seasonal/occasion reference")
    if n_sent > 0 and n_ai > 0 and len([c for c in added_chunks if len(c) > 60]) > 0:
        content_notes.append("added substantial new content")

    removed_preview = "; ".join(f'"{c}"' for c in removed_chunks if c.strip())[:300]
    added_preview   = "; ".join(f'"{c}"' for c in added_chunks if c.strip())[:300]

    parts = [length_note, f"Location: {zone_note}"]
    if content_notes:
        parts.append("Type: " + ", ".join(dict.fromkeys(content_notes)))
    if removed_preview:
        parts.append(f"Removed: {removed_preview}")
    if added_preview:
        parts.append(f"Added: {added_preview}")

    return "\n".join(parts)


def _diff_html(segs, mode):
    """
    Build an HTML string with inline spans for the evidence table display.
    mode='ai'   → deleted text in red
    mode='sent' → inserted text in bold
    """
    parts = []
    for tag, text in segs:
        escaped = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        if mode == 'ai' and tag == 'delete':
            parts.append(f'<span style="color:#CC0000">{escaped}</span>')
        elif mode == 'sent' and tag == 'insert':
            parts.append(f'<strong>{escaped}</strong>')
        else:
            parts.append(escaped)
    return ''.join(parts)


def run(df_edited: 'pd.DataFrame', progress_callback=None) -> 'pd.DataFrame':
    """
    Adds columns:
      - change_summary        : plain-text description of what changed
      - ai_message_diff_html  : AI message with deleted text marked in red (HTML)
      - sent_message_diff_html: sent message with added text marked bold (HTML)

    These HTML columns are used by the Streamlit evidence expander.
    The plain-text columns remain unchanged for KPI computation.
    """
    df = df_edited.copy()
    summaries       = []
    ai_diff_html    = []
    sent_diff_html  = []

    total = len(df)
    for i, row in df.iterrows():
        if progress_callback:
            progress_callback(i, total)

        ai_orig   = str(row.get('ai_message_original',   '') or '')
        sent_orig = str(row.get('sent_message_original', '') or '')
        ai_en     = str(row.get('ai_message_translated', '') or '')
        sent_en   = str(row.get('sent_message_translated','') or '')

        segs_orig_ai, segs_orig_sent = _char_diff(ai_orig, sent_orig)
        segs_en_ai,   segs_en_sent   = _char_diff(ai_en,   sent_en)

        summaries.append(_summarise(ai_en, sent_en))
        ai_diff_html.append(_diff_html(segs_en_ai,   'ai'))
        sent_diff_html.append(_diff_html(segs_en_sent, 'sent'))

    df['change_summary']         = summaries
    df['ai_message_diff_html']   = ai_diff_html
    df['sent_message_diff_html'] = sent_diff_html

    if progress_callback:
        progress_callback(total, total)

    return df
