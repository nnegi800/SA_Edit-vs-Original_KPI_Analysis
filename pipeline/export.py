"""
Generate downloadable xlsx files as BytesIO objects.
to_coloured_xlsx : SA_edits_coloured.xlsx equivalent (with diff highlights)
to_kpi_xlsx      : KPI_Summary.xlsx equivalent
"""
import difflib
import io

import pandas as pd
import xlsxwriter


def _char_diff(a, b):
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    segs_a, segs_b = [], []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'equal':
            segs_a.append(('eq',  a[i1:i2]))
            segs_b.append(('eq',  b[j1:j2]))
        elif tag == 'delete':
            segs_a.append(('del', a[i1:i2]))
        elif tag == 'insert':
            segs_b.append(('ins', b[j1:j2]))
        elif tag == 'replace':
            segs_a.append(('del', a[i1:i2]))
            segs_b.append(('ins', b[j1:j2]))
    return segs_a, segs_b


def to_coloured_xlsx(df_edited: pd.DataFrame) -> io.BytesIO:
    """
    Returns a BytesIO containing SA_edits_coloured.xlsx:
    - red  = text removed from the AI message
    - bold = text added in the SA's sent message
    """
    buf = io.BytesIO()
    wb  = xlsxwriter.Workbook(buf, {'in_memory': True})
    ws  = wb.add_worksheet("SA Edits")

    fill_a = '#EEF2FF'
    fill_b = '#FFFFFF'

    hdr_fmt = wb.add_format({
        'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#1F3864',
        'text_wrap': True, 'valign': 'vcenter', 'font_size': 11,
    })

    fmt_cache = {}
    def get_fmt(even, style='normal'):
        key = (even, style)
        if key not in fmt_cache:
            bg    = fill_a if even else fill_b
            props = {'text_wrap': True, 'valign': 'top', 'font_size': 11, 'bg_color': bg}
            if style == 'red':
                props['font_color'] = '#CC0000'
            elif style == 'bold':
                props['bold'] = True
            fmt_cache[key] = wb.add_format(props)
        return fmt_cache[key]

    # Columns to write — drop the HTML diff columns (those are UI-only)
    write_cols = [c for c in df_edited.columns if not c.endswith('_diff_html')]
    widths = {
        'seller_id': 12, 'unify_id': 18, 'use_case': 22, 'created': 22, 'send_time': 20,
        'time_to_send (days)': 14,
        'ai_action_plan_id': 18, 'country': 12, 'num_versions': 10,
        'conversation_starter_subject_original': 30,
        'conversation_starter_subject_translated': 30,
        'description_original': 40, 'description_translated': 40,
        'ai_message_original': 50, 'ai_message_translated': 50,
        'sent_message_original': 50, 'sent_message_translated': 50,
        'change_summary': 60,
    }

    for ci, h in enumerate(write_cols):
        ws.write(0, ci, h, hdr_fmt)
        ws.set_column(ci, ci, widths.get(h, 18))
    ws.set_row(0, 30)

    DIFF_COLS = {
        'ai_message_original', 'ai_message_translated',
        'sent_message_original', 'sent_message_translated',
    }

    def build_rich(segs, mode, even):
        parts = []
        has_markup = False
        for tag, text in segs:
            if not text:
                continue
            if mode == 'ai' and tag == 'del':
                parts += [get_fmt(even, 'red'), text]
                has_markup = True
            elif mode == 'sent' and tag == 'ins':
                parts += [get_fmt(even, 'bold'), text]
                has_markup = True
            else:
                parts += [get_fmt(even, 'normal'), text]
        return parts if has_markup else None

    for ri, (_, row) in enumerate(df_edited.iterrows()):
        even     = ri % 2 == 0
        xlsx_row = ri + 1

        ai_orig   = str(row.get('ai_message_original',    '') or '')
        sent_orig = str(row.get('sent_message_original',  '') or '')
        ai_en     = str(row.get('ai_message_translated',  '') or '')
        sent_en   = str(row.get('sent_message_translated','') or '')

        segs_ai_orig,   segs_sent_orig = _char_diff(ai_orig, sent_orig)
        segs_ai_en,     segs_sent_en   = _char_diff(ai_en,   sent_en)

        rich_map = {
            'ai_message_original':    build_rich(segs_ai_orig,   'ai',   even),
            'ai_message_translated':  build_rich(segs_ai_en,      'ai',   even),
            'sent_message_original':  build_rich(segs_sent_orig,  'sent', even),
            'sent_message_translated':build_rich(segs_sent_en,    'sent', even),
        }

        base_fmt = get_fmt(even, 'normal')
        for ci, h in enumerate(write_cols):
            val = row.get(h, '')
            if val is None:
                val = ''
            if h in DIFF_COLS and rich_map.get(h):
                ws.write_rich_string(xlsx_row, ci, *rich_map[h], base_fmt)
            elif h == 'time_to_send (days)':
                try:
                    ws.write_number(xlsx_row, ci, float(val), base_fmt)
                except Exception:
                    ws.write(xlsx_row, ci, str(val), base_fmt)
            else:
                ws.write(xlsx_row, ci, str(val) if val != '' else '', base_fmt)

    wb.close()
    buf.seek(0)
    return buf


def to_kpi_xlsx(kpi_results: dict) -> io.BytesIO:
    """
    Returns a BytesIO containing KPI_Summary.xlsx.
    kpi_results is the dict returned by pipeline/kpis.py run().
    """
    buf = io.BytesIO()
    wb  = xlsxwriter.Workbook(buf, {'in_memory': True})
    ws  = wb.add_worksheet("KPI Summary")
    wb.add_worksheet("Notes")

    UC_LABELS_ORDER = ['Emotional Bonding', 'Task', 'Product Storytelling']
    cols     = ['Overall'] + UC_LABELS_ORDER
    agg_data = [kpi_results['overall']] + [kpi_results.get(uc, {}) for uc in UC_LABELS_ORDER]

    title_fmt   = wb.add_format({'bold': True, 'font_size': 14, 'font_color': '#1F3864', 'bottom': 2})
    section_fmt = wb.add_format({'bold': True, 'font_size': 10, 'bg_color': '#1F3864',
                                  'font_color': '#FFFFFF', 'text_wrap': True})
    hdr_col_fmt = wb.add_format({'bold': True, 'font_size': 11, 'bg_color': '#1F3864',
                                  'font_color': '#FFFFFF', 'align': 'center',
                                  'border': 1, 'border_color': '#C0C8E0', 'text_wrap': True})
    metric_fmt  = wb.add_format({'font_size': 10, 'text_wrap': True, 'valign': 'vcenter',
                                  'bg_color': '#F4F6FB'})
    metric_fmt2 = wb.add_format({'font_size': 10, 'text_wrap': True, 'valign': 'vcenter',
                                  'bg_color': '#FFFFFF'})
    num_fmt     = wb.add_format({'font_size': 11, 'bold': True, 'align': 'center',
                                  'bg_color': '#EEF2FF', 'border': 1, 'border_color': '#C0C8E0'})
    num_fmt2    = wb.add_format({'font_size': 11, 'bold': True, 'align': 'center',
                                  'bg_color': '#FFFFFF', 'border': 1, 'border_color': '#C0C8E0'})

    ws.set_column(0, 0, 44)
    ws.set_column(1, len(cols), 22)

    ws.merge_range(0, 0, 0, len(cols), "SA Edit Pattern Analysis — KPI Summary", title_fmt)
    ws.set_row(0, 28)
    ws.write(1, 0, "KPI Metric", hdr_col_fmt)
    for ci, c in enumerate(cols, 1):
        ws.write(1, ci, c, hdr_col_fmt)
    ws.set_row(1, 30)
    ws.freeze_panes(2, 1)

    kpi_row     = 2
    parity      = 0
    for section_label, kpi_items in kpi_results['sections']:
        ws.merge_range(kpi_row, 0, kpi_row, len(cols), section_label, section_fmt)
        ws.set_row(kpi_row, 18)
        kpi_row += 1
        parity   = 0
        for label, key in kpi_items:
            even  = parity % 2 == 0
            lf    = metric_fmt  if even else metric_fmt2
            vf    = num_fmt     if even else num_fmt2
            ws.write(kpi_row, 0, label, lf)
            for ci, ag in enumerate(agg_data, 1):
                val = ag.get(key, '—')
                ws.write(kpi_row, ci, val if val is not None else '—', vf)
            ws.set_row(kpi_row, 20)
            kpi_row += 1
            parity  += 1

    wb.close()
    buf.seek(0)
    return buf
