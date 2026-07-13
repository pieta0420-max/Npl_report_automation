"""Fills the 자산 소개 section (pages 1-9) of the company's Word investment
committee report template with data from the same aggregator.py tables used
for the Excel output.

Like the Excel VF.xlsx template, the source .docx is DocumentSAFER-DRM
protected -- but COM automation can open, edit, and *save* it because Word
handles the DRM decryption/permission-check using the current Windows
login, the same access a person has double-clicking the file. So instead of
building a fresh document (which would require reproducing Word's native
paragraph/table styling from scratch), we copy the original template and
edit values in place via COM, leaving pages 10+ (LSPA/평가결과/시뮬레이션/
시장동향/투자금액, none of which are derivable from a Data Disk) untouched.

Table row counts are rebuilt to match the actual data (delete old data rows,
add new ones) rather than assumed fixed, since a different bank's DD will
have a different number of distinct property types/regions/amount tiers
than this WRB template happened to have. New rows created via Rows.Add()
turned out to inherit the header row's direct formatting (shading, bold,
center alignment) rather than plain body formatting, so every data cell's
formatting is set explicitly below instead of relying on inheritance.
"""
from __future__ import annotations

import shutil
from dataclasses import dataclass, field
from typing import List, Optional

import pandas as pd

import narrative

# Word constants (win32com late-binding doesn't expose the wdXxx names)
WD_ALIGN_LEFT = 0
WD_ALIGN_CENTER = 1
WD_ALIGN_RIGHT = 2
WD_BORDER_TOP = -1
WD_BORDER_BOTTOM = -3
WD_LINE_STYLE_NONE = 0
WD_LINE_STYLE_SINGLE = 1
WD_COLOR_BLACK = 0
WD_COLOR_AUTOMATIC = -16777216
WD_TEXTURE_NONE = 0
WD_LIST_NO_NUMBERING = 0

# #DAEEF3 (a light blue), converted to Word's BGR long-color format.
_HL_R, _HL_G, _HL_B = 0xDA, 0xEE, 0xF3
HIGHLIGHT_COLOR = (_HL_B << 16) | (_HL_G << 8) | _HL_R


@dataclass
class Cell:
    text: str
    bold: bool = False
    align: int = WD_ALIGN_RIGHT


@dataclass
class RowSpec:
    cells: List[Cell]
    border_top: bool = False
    border_bottom: bool = False
    highlight: bool = False


def L(text, bold: bool = False) -> Cell:
    return Cell(text, bold=bold, align=WD_ALIGN_LEFT)


def R(text, bold: bool = False) -> Cell:
    return Cell(text, bold=bold, align=WD_ALIGN_RIGHT)


def C(text, bold: bool = False) -> Cell:
    return Cell(text, bold=bold, align=WD_ALIGN_CENTER)


CM_TO_PT = 28.3464567
GROUP_TABLE_ROW_HEIGHT_CM = 0.61


# ---------------------------------------------------------------------------
# Formatting helpers (Word tables hold static text, not live formulas)
# ---------------------------------------------------------------------------

def _n(x) -> str:
    return f"{int(round(x)):,}"


def _pct(x) -> str:
    return f"{x * 100:.2f}%"


# ---------------------------------------------------------------------------
# COM table manipulation primitives
# ---------------------------------------------------------------------------

def _clear_data_rows(table, header_rows: int = 1) -> None:
    while table.Rows.Count > header_rows:
        table.Rows(header_rows + 1).Delete()


def _style_cell(cell, cell_spec: Cell, highlight: bool = False) -> None:
    cell.Range.Text = str(cell_spec.text)
    cell.Range.Font.Bold = True if highlight else cell_spec.bold
    cell.Range.Font.Color = WD_COLOR_BLACK
    cell.Range.ParagraphFormat.Alignment = cell_spec.align
    try:
        # Rows.Add() rows turned out to also inherit the header row's list
        # numbering (a small bullet/marker glyph appeared before every data
        # row's text) -- explicitly stripping it here, same fix pattern as
        # the earlier shading/bold/alignment inheritance issue.
        cell.Range.ListFormat.RemoveNumbers()
    except Exception:
        pass
    try:
        cell.Shading.Texture = WD_TEXTURE_NONE
        cell.Shading.BackgroundPatternColor = HIGHLIGHT_COLOR if highlight else WD_COLOR_AUTOMATIC
    except Exception:
        pass


def _add_row(table, row_spec: RowSpec, skip_highlight_cols: frozenset = frozenset()):
    row = table.Rows.Add()
    for c, cell_spec in enumerate(row_spec.cells, start=1):
        if c > table.Columns.Count:
            break
        hl = row_spec.highlight and c not in skip_highlight_cols
        _style_cell(row.Cells(c), cell_spec, highlight=hl)
    return row


def _rebuild_table(table, rows: List[RowSpec], header_rows: int = 1,
                    skip_highlight_cols: frozenset = frozenset()) -> None:
    """Text/font/shading is set while building each row, but NOT borders:
    adjacent rows share a border line in Word's table model, so setting a
    "no border" on one row's top edge while building it can silently
    overwrite the "single line" a previous row had just set on its own
    bottom edge (same physical line). Borders are applied in one pass after
    every row exists, so whichever setting is written for a given shared
    edge is unambiguous and not raced by the next row's construction.
    skip_highlight_cols excludes specific (1-indexed) columns from a row's
    highlight even when row_spec.highlight is True -- used for 2.4/2.5's
    Pool column, which keeps its own bold-on-pool-label styling but not the
    #DAEEF3 rank highlight the rest of the row gets."""
    _clear_data_rows(table, header_rows)
    for row_spec in rows:
        _add_row(table, row_spec, skip_highlight_cols=skip_highlight_cols)

    n_cols = table.Columns.Count
    for r in range(header_rows + 1, table.Rows.Count + 1):
        for c in range(1, n_cols + 1):
            cell = table.Cell(r, c)
            cell.Borders(WD_BORDER_TOP).LineStyle = WD_LINE_STYLE_NONE
            cell.Borders(WD_BORDER_BOTTOM).LineStyle = WD_LINE_STYLE_NONE
    for i, row_spec in enumerate(rows):
        r = header_rows + 1 + i
        if not (row_spec.border_top or row_spec.border_bottom):
            continue
        for c in range(1, n_cols + 1):
            cell = table.Cell(r, c)
            if row_spec.border_top:
                cell.Borders(WD_BORDER_TOP).LineStyle = WD_LINE_STYLE_SINGLE
            if row_spec.border_bottom:
                cell.Borders(WD_BORDER_BOTTOM).LineStyle = WD_LINE_STYLE_SINGLE


def _set_row_height(table, cm: float, header_rows: int = 1) -> None:
    """Applies a fixed (not merely minimum) row height to every data row --
    Rows.Add() rows inherited an oversized default height, and 'unify' means
    every data row should measure the same in Word's Table Properties dialog
    (which shows cm for this Korean-locale doc). The header row is left
    alone since it may wrap onto two lines and a fixed height would clip it."""
    pts = cm * CM_TO_PT
    for r in range(header_rows + 1, table.Rows.Count + 1):
        row = table.Rows(r)
        row.HeightRule = 2  # wdRowHeightExactly
        row.Height = pts


def _keep_table_together(table) -> None:
    """Prevents a single row's text from splitting across a page boundary,
    and chains every row to the next via KeepWithNext so Word treats the
    whole table as one unbreakable block -- if it doesn't fit on the current
    page, the entire table moves to the next page instead of splitting."""
    table.Rows.AllowBreakAcrossPages = False
    n_rows = table.Rows.Count
    for r in range(1, n_rows + 1):
        for cell in table.Rows(r).Cells:
            for p in cell.Range.Paragraphs:
                p.Format.KeepTogether = True
                if r < n_rows:
                    p.Format.KeepWithNext = True


def _merge_pool_and_category(table, row: int, text: str) -> None:
    """Rule (총괄표 합계행): the Pool cell (col 1) and 구분 cell (col 2) of
    the grand-total row become one merged, centered cell reading '합계'.
    Merging can clear borders on the merged range, so top/bottom borders are
    re-applied to the whole row afterward."""
    left = table.Cell(row, 1)
    right = table.Cell(row, 2)
    left.Merge(right)  # Merge is a Sub in the Word object model, not a function
    merged = table.Cell(row, 1)
    merged.Range.Text = text
    merged.Range.Font.Bold = True
    merged.Range.Font.Color = WD_COLOR_BLACK
    merged.Range.ParagraphFormat.Alignment = WD_ALIGN_CENTER
    for cell in table.Rows(row).Cells:
        cell.Borders(WD_BORDER_TOP).LineStyle = WD_LINE_STYLE_SINGLE
        cell.Borders(WD_BORDER_BOTTOM).LineStyle = WD_LINE_STYLE_SINGLE


def _set_pool_column_dividers(table, pool_block_last_rows: List[int], header_rows: int = 1,
                               close_bottom: bool = False) -> None:
    """Rule: in the Pool column (col 1) only, draw a line ONLY between
    consecutive pool blocks -- not at every subtotal/total row like the
    other columns (columns 2+ already got their correct borders from the
    general subtotal/total rule during _rebuild_table; only column 1 needs
    overriding here). `pool_block_last_rows` are the (1-indexed, within the
    rebuilt table) row numbers of each pool's last row, except the final
    pool (so len == n_pools - 1). close_bottom=True additionally draws a
    line under the table's very last row, closing off the Pool column the
    same way the other columns already are."""
    for r in range(header_rows + 1, table.Rows.Count + 1):
        cell = table.Cell(r, 1)
        cell.Borders(WD_BORDER_TOP).LineStyle = WD_LINE_STYLE_NONE
        cell.Borders(WD_BORDER_BOTTOM).LineStyle = WD_LINE_STYLE_NONE
    for r in pool_block_last_rows:
        table.Cell(r, 1).Borders(WD_BORDER_BOTTOM).LineStyle = WD_LINE_STYLE_SINGLE
    if close_bottom and table.Rows.Count > header_rows:
        table.Cell(table.Rows.Count, 1).Borders(WD_BORDER_BOTTOM).LineStyle = WD_LINE_STYLE_SINGLE


def _cell_text(cell) -> str:
    """Cell.Range.Text always carries a trailing cell-end marker (\\x07,
    preceded by \\r if the cell holds a full paragraph) that str.strip()
    doesn't remove since it isn't whitespace -- left in, a blank cell reads
    back as '\\x07' instead of '', which silently breaks exact-text
    comparisons (e.g. every "blank" cell then looks like a distinct,
    non-empty value)."""
    return cell.Range.Text.replace("\r", "").replace("\x07", "").strip()


def _detect_template_pool_count(table1) -> int:
    """The template's per-pool table sections (금액별/담보물종류별/담보물
    지역별) are physical tables, one set per pool -- unlike 총괄표 itself,
    which embeds every pool as rows in a single table and so is already
    pool-count-agnostic. A template built for a prior deal with fewer (or
    more) pools than the current one won't have the right number of table
    slots. Read this BEFORE any table is rebuilt: the template's own
    still-intact 총괄표 (from whatever deal it was last used for) already
    lists each pool's letter in column 1, so counting the distinct
    non-empty entries there tells us how many pools the template's table
    sections were actually built for."""
    pools = set()
    for r in range(2, table1.Rows.Count + 1):
        text = _cell_text(table1.Cell(r, 1))
        if text and text not in ("Pool", "합계"):
            pools.add(text)
    return len(pools) if pools else 1


def _detect_header_rows(table, pools) -> int:
    """Tables with a Pool column per row (경매/회생/신용보증서) are normally
    assumed to have a single header row, but this template's 신용보증서
    table turned out to have 2 (a merged category row plus a sub-label
    row) -- rebuilding from row 2 then clobbered the real sub-label row and
    crashed trying to delete a row inside a vertical merge. Detecting the
    true header row count directly, by finding the first row whose column
    1 is one of the deal's actual pool letters, sidesteps assuming any
    fixed header shape."""
    pool_set = set(pools)
    for r in range(1, table.Rows.Count + 1):
        try:
            text = _cell_text(table.Cell(r, 1))
        except Exception:
            continue
        if text in pool_set:
            return r - 1
    return 1


def _unmerge_header_vertical_cells(table, header_rows: int) -> None:
    """A vertically-merged header cell (e.g. "Pool" spanning both header
    rows) makes Rows.Delete() fail on ANY row in the table -- confirmed by
    testing: deleting the first data row and deleting the last row both
    failed identically with the same Word COM error, even though neither
    row was itself part of the merge. Splitting the merge back into one
    cell per row sidesteps this; our rebuild never touches header rows
    anyway, so losing the merge there has no effect on the content we
    write, and the header's own text (already correct from the template)
    is preserved -- Split() keeps existing text in the first resulting
    cell and leaves the rest blank."""
    if header_rows < 2:
        return
    n_cols = table.Columns.Count
    for c in range(1, n_cols + 1):
        try:
            table.Cell(2, c)
        except Exception:
            table.Cell(1, c).Split(NumRows=header_rows, NumColumns=1)


def _clone_table(doc, source_index: int, insert_after_index: int) -> None:
    """Copies the table at source_index (1-based, re-read fresh so this is
    safe to call repeatedly as indices shift with each insertion) and
    inserts the duplicate immediately after the table at insert_after_index
    (may be the same table, or a different one -- used to grow a per-pool
    table block by cloning the last pool's table(s) onto the end of the
    section). Pasting a copied table's Range directly against another
    table's end merges the two into one (confirmed by testing -- row counts
    added together instead of a new Tables entry appearing); inserting an
    empty paragraph first, then pasting after THAT, keeps them separate."""
    import pythoncom

    src = doc.Tables(source_index)
    target = doc.Tables(insert_after_index)
    end = target.Range.End
    gap = doc.Range(end, end)
    gap.InsertParagraphAfter()
    new_pos = end + 1
    src.Range.Copy()
    pythoncom.PumpWaitingMessages()
    doc.Range(new_pos, new_pos).Paste()
    pythoncom.PumpWaitingMessages()


def _rewrite_matrix_header(table, prop_groups: List[str]) -> None:
    """Unlike every other table's header (fixed labels like Pool/구분 that
    never change deal to deal, so leaving them untouched is exactly what we
    want), the (*) 지역별/물건별 OPB 비중 matrix's header names the deal's
    actual property-type groups -- which come from classification_config
    .json and can differ from whatever group set the template's original
    deal happened to use. Left untouched, the template's old column labels
    (e.g. a stale group name from a prior deal) end up sitting over this
    deal's differently-ordered/differently-named data."""
    headers = ["구분"] + list(prop_groups) + ["합계"]
    for i in range(min(len(headers), table.Columns.Count)):
        table.Cell(1, i + 1).Range.Text = headers[i]


def _clone_paragraph_after(doc, paragraph):
    """Duplicates a narrative paragraph (with its formatting) and inserts
    the copy immediately after it -- used when the template has fewer
    instances of a given narrative sentence than the deal has pools (e.g.
    a template built for 3 pools has no 4th "Pool D" sentence slot at
    all). Returns the newly created Paragraph object."""
    import pythoncom

    end = paragraph.Range.End
    paragraph.Range.Copy()
    pythoncom.PumpWaitingMessages()
    doc.Range(end, end).Paste()
    pythoncom.PumpWaitingMessages()
    return doc.Range(end, end + 1).Paragraphs(1)


def _ensure_anchor_count(doc, anchors: list, target_count: int) -> None:
    """Extends an anchors list (from _collect_narrative_anchors) up to
    target_count by cloning its last entry, once per missing pool."""
    while anchors and len(anchors) < target_count:
        anchors.append(_clone_paragraph_after(doc, anchors[-1]))


def _extend_pool_table_block(doc, first_index: int, block_size: int,
                              template_pools: int, target_pools: int) -> None:
    """A "block" is `block_size` consecutive tables per pool (1 for a plain
    per-pool table, 2 for a data+matrix pair), repeated `template_pools`
    times starting at `first_index`. If the deal has more pools than the
    template was built for, clones the LAST pool's block onto the end of
    the section, once per extra pool, so it ends up `target_pools` blocks
    wide before any content is written into it."""
    extra = target_pools - template_pools
    if extra <= 0:
        return
    last_index = first_index + template_pools * block_size - 1
    for _ in range(extra):
        block_start = last_index - block_size + 1
        insert_after = last_index
        for offset in range(block_size):
            _clone_table(doc, block_start + offset, insert_after)
            insert_after += 1
        last_index = insert_after


def _collect_narrative_anchors(doc, stop_before_table) -> dict:
    """One pass over the paragraphs in pages 1-9 (up to stop_before_table's
    start), bucketing the 4 kinds of intro sentence by a distinctive keyword
    so callers can grab "the Nth occurrence" (1st = Pool A, 2nd = Pool B).

    Earlier this used "the last non-empty paragraph immediately before a
    given table" (via Range boundary math), but Table.Range.Start turned out
    to sit at a position that made that lookup grab the table's own header
    cell instead, merging the sentence text into it. Matching by the
    sentence's own distinctive wording sidesteps that boundary ambiguity
    entirely."""
    anchors = {"overall": [], "amount": [], "proptype": [], "region": []}
    limit = stop_before_table.Range.Start
    for p in doc.Paragraphs:
        if p.Range.Start >= limit:
            break
        text = p.Range.Text.strip()
        if not text:
            continue
        if "매각 Program은 총" in text:
            anchors["overall"].append(p)
        elif "차주가 전체 OPB의" in text and "이상" in text:
            anchors["amount"].append(p)
        elif "가장 많은 비중을 차지" in text:
            anchors["proptype"].append(p)
        elif "소재 물건이" in text and "차지하고 있습니다" in text:
            anchors["region"].append(p)
    return anchors


def _replace_paragraph_text(doc, paragraph, text: str, align: Optional[int] = None,
                             flush_left: bool = False) -> None:
    """Setting Range.Text on a Paragraph object leaves that same object
    pointing at the wrong (now-empty, differently-formatted) range
    afterward -- confirmed by testing: re-reading paragraph.Range.Text right
    after the assignment came back empty, and a subsequent alignment change
    landed on that stale range instead of the new text. Re-deriving a fresh
    Range from the paragraph's original start position sidesteps that.

    flush_left=True additionally zeroes the paragraph's left/first-line
    indent -- the template paragraph style carries an indent that made
    left-aligned text sit noticeably off the left margin -- and prepends
    exactly one space so there's still a small, controlled gap instead."""
    if flush_left:
        text = " " + text
    start = paragraph.Range.Start
    paragraph.Range.Text = text + "\r"
    rng = doc.Range(start, start + len(text))
    if align is not None:
        rng.ParagraphFormat.Alignment = align
    if flush_left:
        # This template's paragraphs use CJK "character unit" indentation
        # (CharacterUnitFirstLineIndent=6, i.e. a 6-character first-line
        # indent) rather than plain points -- confirmed by testing that
        # setting FirstLineIndent alone silently reverted on the next
        # read/save because Word recalculates the point value from the
        # (untouched) character-unit value. Both must be zeroed together.
        fresh_para = rng.Paragraphs(1)
        fresh_para.Format.CharacterUnitLeftIndent = 0
        fresh_para.Format.CharacterUnitRightIndent = 0
        fresh_para.Format.CharacterUnitFirstLineIndent = 0
        fresh_para.Format.LeftIndent = 0
        fresh_para.Format.RightIndent = 0
        fresh_para.Format.FirstLineIndent = 0


# ---------------------------------------------------------------------------
# Section 1: 총괄표 (Table 1) + (*) 차주 형태별 구분 (Table 2)
# ---------------------------------------------------------------------------
# Both are [Pool, 구분, 차주수, 구성비, OPB, 구성비, 채권총액, 구성비]. Bold is
# limited to the Pool-label cell and to 소계/합계 rows (whole row); borders
# are top+bottom on 소계/합계 rows only, except column 1 (Pool) which only
# gets a divider between pool blocks (see _set_pool_column_dividers).

def _data_row_8col(pool_label: str, category: str, r: pd.Series) -> RowSpec:
    bold = bool(pool_label)
    return RowSpec([
        C(pool_label, bold=bold), L(category), R(_n(r["count"])), R(_pct(r["count_pct"])),
        R(_n(r["opb_million"])), R(_pct(r["opb_pct"])), R(_n(r["claim_million"])), R(_pct(r["claim_pct"])),
    ])


def _subtotal_row_8col(label: str, count, count_pct, opb, opb_pct, claim, claim_pct) -> RowSpec:
    return RowSpec([
        C("", bold=True), L(label, bold=True), R(_n(count), True), R(_pct(count_pct), True),
        R(_n(opb), True), R(_pct(opb_pct), True), R(_n(claim), True), R(_pct(claim_pct), True),
    ], border_top=True, border_bottom=True)


def _rows_overall_summary(agg_table: pd.DataFrame, config: dict):
    rows: List[RowSpec] = []
    pool_block_last_rows: List[int] = []
    pools = sorted(p for p in agg_table["pool_type"].unique() if p)
    for pool in pools:
        pool_rows = agg_table[agg_table["pool_type"] == pool]
        data = pool_rows[pool_rows["row_kind"] == "data"].copy()
        data["_order"] = data["category"].apply(lambda c: 0 if "특별" in c else 1)
        data = data.sort_values("_order")
        for i, (_, r) in enumerate(data.iterrows()):
            rows.append(_data_row_8col(pool if i == 0 else "", r["category"], r))
        sub = pool_rows[pool_rows["row_kind"] == "subtotal"].iloc[0]
        rows.append(_subtotal_row_8col("소계", sub["count"], sub["count_pct"], sub["opb_million"],
                                        sub["opb_pct"], sub["claim_million"], sub["claim_pct"]))
        pool_block_last_rows.append(len(rows) + 1)  # +1 for header row offset
    pool_block_last_rows.pop()  # no divider after the LAST pool (before 합계)

    total = agg_table[agg_table["row_kind"] == "total"].iloc[0]
    rows.append(RowSpec([
        C("합계", bold=True), L("", bold=True), R(_n(total["count"]), True), R("100.00%", True),
        R(_n(total["opb_million"]), True), R("100.00%", True), R(_n(total["claim_million"]), True), R("100.00%", True),
    ], border_top=True, border_bottom=True))
    return rows, pool_block_last_rows


def _rows_borrower_type(agg_table: pd.DataFrame, overall_summary_table: pd.DataFrame):
    rows: List[RowSpec] = []
    pool_block_last_rows: List[int] = []
    order = {"개인": 0, "개인사업자": 1, "법인": 2}
    for pool in sorted(agg_table["pool_type"].unique()):
        pool_rows = agg_table[agg_table["pool_type"] == pool]
        data = pool_rows[pool_rows["row_kind"] == "data"].copy()
        data["_order"] = data["category"].map(order).fillna(9)
        data = data.sort_values("_order")
        for i, (_, r) in enumerate(data.iterrows()):
            rows.append(_data_row_8col(pool if i == 0 else "", r["category"], r))
        sub = pool_rows[pool_rows["row_kind"] == "subtotal"].iloc[0]
        rows.append(_subtotal_row_8col("소계", sub["count"], sub["count_pct"], sub["opb_million"],
                                        sub["opb_pct"], sub["claim_million"], sub["claim_pct"]))
        pool_block_last_rows.append(len(rows) + 1)
    pool_block_last_rows.pop()

    grand = overall_summary_table[overall_summary_table["row_kind"] == "total"].iloc[0]
    rows.append(RowSpec([
        C("합계", bold=True), L("", bold=True), R(_n(grand["count"]), True), R("100.00%", True),
        R(_n(grand["opb_million"]), True), R("100.00%", True), R(_n(grand["claim_million"]), True), R("100.00%", True),
    ], border_top=True, border_bottom=True))
    return rows, pool_block_last_rows


# ---------------------------------------------------------------------------
# Section 2: 금액별 분류 (one table per pool)
# ---------------------------------------------------------------------------

def _rows_amount_bucket(pool_table: pd.DataFrame) -> List[RowSpec]:
    rows: List[RowSpec] = []
    data = pool_table[pool_table["row_kind"] == "data"]
    # Highlight the top 2 tiers by amount threshold itself (i.e. simply the
    # first 2 rows in the table's existing descending-threshold order --
    # 100억 이상, 50억 이상 -- not the 2 tiers with the largest OPB share).
    for i, (_, r) in enumerate(data.iterrows()):
        rows.append(RowSpec([
            L(r["category"]), R(_n(r["count"])), R(_n(r["opb_million"])), R(_pct(r["opb_pct"])),
            R(_n(r["cum_count"])), R(_n(r["cum_opb_million"])), R(_pct(r["cum_opb_pct"])),
        ], highlight=(i < 2)))
    sub = pool_table[pool_table["row_kind"] == "subtotal"].iloc[0]
    rows.append(RowSpec([
        L("합계", True), R(_n(sub["count"]), True), R(_n(sub["opb_million"]), True), R("100.00%", True),
        R("", True), R("", True), R("", True),
    ], border_top=True, border_bottom=True))
    return rows


# ---------------------------------------------------------------------------
# Section 3/4: 담보물 종류별 / 지역별 분류 (one table + one matrix per pool)
# ---------------------------------------------------------------------------
# [구분, 물건수, OPB, 구성비, 유사유형물건수, 유사유형OPB합계, 유사유형구성비] --
# the last 3 (group rollup) columns are populated only on the first row of
# each group's block, left blank afterward (mirrors the merged-cell look of
# the original template without needing real cell merges).

def _top_groups(pool_table: pd.DataFrame, n: int = 2) -> set:
    """Ranks each group by its group_opb_pct (유사유형 구성비) and returns
    the names of the top n -- these get the #DAEEF3 highlight + bold, both
    in the breakdown table itself and in the region x property matrix that
    shares the same region groupings."""
    data = pool_table[pool_table["row_kind"] == "data"]
    seen = {}
    for _, r in data.iterrows():
        seen.setdefault(r["group"], r["group_opb_pct"])
    ranked = sorted(seen.items(), key=lambda kv: kv[1], reverse=True)
    return {name for name, _ in ranked[:n]}


def _rows_group_breakdown(pool_table: pd.DataFrame, label_key: str):
    """[구분, 물건수, OPB, 구성비, 유사유형물건수, 유사유형OPB, 유사유형구성비].
    The group-rollup columns (5-7) are populated only on each group's first
    row; the group's own name is written as "(그룹명)" into column 7 (유사
    유형구성비), right-aligned, on the row directly below that first row --
    or, for a single-member group, on an extra row inserted just for the
    label. A divider line separates each group's block from the next. Every
    row belonging to one of the top-2 groups by share is highlighted."""
    rows: List[RowSpec] = []
    data = pool_table[pool_table["row_kind"] == "data"]
    top_groups = _top_groups(pool_table, 2)

    blocks: List[tuple] = []
    for _, r in data.iterrows():
        if blocks and blocks[-1][0] == r["group"]:
            blocks[-1][1].append(r)
        else:
            blocks.append((r["group"], [r]))

    for block_idx, (group_name, items) in enumerate(blocks):
        hl = group_name in top_groups
        block_start = len(rows)
        for i, r in enumerate(items):
            if i == 0:
                group_cells = [R(_n(r["group_count"])), R(_n(r["group_opb_million"])), R(_pct(r["group_opb_pct"]))]
            else:
                group_cells = [R(""), R(""), R("")]
            rows.append(RowSpec([
                L(r[label_key]), R(_n(r["count"])), R(_n(r["opb_million"])), R(_pct(r["opb_pct"])),
            ] + group_cells, highlight=hl))

        if len(items) >= 2:
            rows[block_start + 1].cells[6] = R(f"({group_name})")
            block_last = rows[-1]  # the block's true last row, not the label row
        else:
            block_last = RowSpec([L(""), R(""), R(""), R(""), R(""), R(""), R(f"({group_name})")], highlight=hl)
            rows.append(block_last)

        if block_idx < len(blocks) - 1:
            block_last.border_bottom = True

    total = pool_table[pool_table["row_kind"] == "total"].iloc[0]
    rows.append(RowSpec([
        L("합계", True), R(_n(total["count"]), True), R(_n(total["opb_million"]), True), R("100.00%", True),
        R("", True), R("", True), R("", True),
    ], border_top=True, border_bottom=True))
    return rows, top_groups


def _pct_or_dash(x: float) -> str:
    return "-" if abs(x) < 1e-9 else _pct(x)


def _rows_matrix(pool_matrix: pd.DataFrame, prop_groups: List[str], highlight_groups: set = frozenset()) -> List[RowSpec]:
    rows: List[RowSpec] = []
    row_total = {pg: 0.0 for pg in prop_groups}
    for _, r in pool_matrix.iterrows():
        vals = [R(_pct_or_dash(r[pg])) for pg in prop_groups]
        row_sum = sum(r[pg] for pg in prop_groups)
        for pg in prop_groups:
            row_total[pg] += r[pg]
        hl = r["region_group"] in highlight_groups
        rows.append(RowSpec([L(r["region_group"])] + vals + [R(_pct_or_dash(row_sum))], highlight=hl))
    col_totals = [R(_pct_or_dash(row_total[pg]), True) for pg in prop_groups]
    grand_total = sum(row_total.values())
    rows.append(RowSpec([L("합계", True)] + col_totals + [R(_pct_or_dash(grand_total), True)],
                         border_top=True, border_bottom=True))
    return rows


# ---------------------------------------------------------------------------
# Section 5/6/7: 경매절차 / 회생절차 / 신용보증서 (single combined table)
# ---------------------------------------------------------------------------

# Word's table column for 2.4 (경매절차) is narrow enough that "경매 미개시"
# wraps onto a second line; removing the space keeps it on one line. This is
# a display-only fix scoped to the Word table, not the shared aggregator
# category text (Excel's wider column doesn't have this problem).
_WORD_LABEL_TIGHTEN = {"경매 미개시": "경매미개시"}


def _rows_foreclosure(agg_table: pd.DataFrame):
    rows: List[RowSpec] = []
    pool_block_last_rows: List[int] = []
    for pool in sorted(agg_table["pool_type"].unique()):
        pool_rows = agg_table[agg_table["pool_type"] == pool]
        # Rows within each pool are shown biggest-share-first (by 채권총액
        # 구성비), with the top row highlighted -- if there's only one
        # category (e.g. 회생절차 with just 개시결정), that single row is
        # both first and highlighted.
        data = pool_rows[pool_rows["row_kind"] == "data"].sort_values("claim_pct", ascending=False)
        for i, (_, r) in enumerate(data.iterrows()):
            pool_label = pool if i == 0 else ""
            category = _WORD_LABEL_TIGHTEN.get(r["category"], r["category"])
            rows.append(RowSpec([
                C(pool_label, bold=bool(pool_label)), L(category), R(_n(r["count"])),
                R(_n(r["opb_million"])), R(_pct(r["opb_pct"])), R(_n(r["claim_million"])), R(_pct(r["claim_pct"])),
            ], highlight=(i == 0)))
        tot = pool_rows[pool_rows["row_kind"] == "total"].iloc[0]
        rows.append(RowSpec([
            C("", True), L("소계", True), R(_n(tot["count"]), True), R(_n(tot["opb_million"]), True),
            R("100.00%", True), R(_n(tot["claim_million"]), True), R("100.00%", True),
        ], border_top=True, border_bottom=True))
        pool_block_last_rows.append(len(rows) + 1)  # +1 for header row offset
    pool_block_last_rows.pop()  # no divider after the LAST pool
    return rows, pool_block_last_rows


def _rows_rehab(agg_table: pd.DataFrame):
    if agg_table.empty:
        return [], []
    return _rows_foreclosure(agg_table)  # identical shape (Pool/구분/count/opb/pct/claim/pct)


def _rows_guarantee(agg_table: pd.DataFrame) -> List[RowSpec]:
    rows: List[RowSpec] = []
    n = len(agg_table)
    for i, (_, r) in enumerate(agg_table.iterrows()):
        rows.append(RowSpec([
            C(r["pool_type"], bold=True), R(_n(r["total_opb_million"])), R(_n(r["target_borrower_count"])),
            R(_n(r["target_opb_million"])), R(_n(r["guarantee_balance_million"])), R(_pct(r["ratio_to_total_opb"])),
        ], border_bottom=(i == n - 1)))
    return rows


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------

def build_word_report(template_path: str, output_path: str, agg_result, frames: dict, config: dict,
                       bank_name: str, program_name: str, report_date: str) -> None:
    import pythoncom
    import win32com.client as win32

    shutil.copy(template_path, output_path)

    pythoncom.CoInitialize()
    word = win32.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        doc = word.Documents.Open(output_path)
        try:
            t = agg_result.tables
            borrower_df = frames["borrower"]

            # --- title page (bounded to the first few paragraphs so we can't
            # accidentally touch unrelated text with a similar shape deeper
            # in the document). Deliberately NOT using Range.Information() to
            # get the page number here: on this DRM-protected document, that
            # call hung indefinitely (multiple minutes, never returned) the
            # moment the document was opened for editing rather than
            # read-only -- open-for-edit apparently runs a DRM/pagination
            # hook that a plain paragraph-index bound avoids entirely. ---
            title_para_count = min(25, doc.Paragraphs.Count)
            for i in range(1, title_para_count + 1):
                p = doc.Paragraphs(i)
                text = p.Range.Text.strip()
                if not text:
                    continue
                if "입찰참여를 위한" in text:
                    _replace_paragraph_text(doc, p, f"{program_name} 입찰참여를 위한")
                elif text.count(".") == 2 and text.replace(".", "").replace(" ", "").isdigit():
                    _replace_paragraph_text(doc, p, report_date, align=WD_ALIGN_CENTER)

            tables = doc.Tables
            amt = t["amount_bucket"]
            pools = sorted(p for p in amt["pool_type"].unique())
            n_pools = len(pools)

            # --- the template's 금액별/담보물종류별/담보물지역별 sections are
            # physical tables, one set per pool (unlike 총괄표 itself, which
            # embeds every pool as rows in one table and needs no extra
            # slots). If this deal has more pools than the template was
            # last used for, clone the last pool's table(s) onto the end of
            # each section BEFORE writing anything, so there's a slot for
            # every pool. Must run in this order -- each section's start
            # index depends on the (now-extended) width of the ones before it. ---
            template_pools = _detect_template_pool_count(tables(1))
            _extend_pool_table_block(doc, 3, 1, template_pools, n_pools)
            proptype_start = 3 + n_pools
            _extend_pool_table_block(doc, proptype_start, 1, template_pools, n_pools)
            region_start = proptype_start + n_pools
            _extend_pool_table_block(doc, region_start, 2, template_pools, n_pools)
            tables = doc.Tables  # re-fetch: table insertions invalidate any cached collection

            amount_indices = list(range(3, 3 + n_pools))
            proptype_indices = list(range(3 + n_pools, 3 + 2 * n_pools))
            region_data_indices = [3 + 2 * n_pools + 2 * i for i in range(n_pools)]
            region_matrix_indices = [i + 1 for i in region_data_indices]
            foreclosure_idx = 3 + 4 * n_pools
            rehab_idx = 4 + 4 * n_pools
            guarantee_idx = 5 + 4 * n_pools

            # --- collect all 4 kinds of intro sentence in one pass, THEN
            # rewrite them (left-aligned), before touching any table ---
            anchors = _collect_narrative_anchors(doc, tables(foreclosure_idx))
            # template built for fewer pools than this deal has -> no Nth
            # narrative sentence slot exists yet for the extra pool(s)
            _ensure_anchor_count(doc, anchors["amount"], n_pools)
            _ensure_anchor_count(doc, anchors["proptype"], n_pools)
            _ensure_anchor_count(doc, anchors["region"], n_pools)
            if anchors["overall"]:
                _replace_paragraph_text(doc, anchors["overall"][0],
                                         narrative.overall_summary_sentence(bank_name, borrower_df),
                                         align=WD_ALIGN_LEFT, flush_left=True)
            for para, pool in zip(anchors["amount"], pools):
                pool_t = amt[amt["pool_type"] == pool]
                _replace_paragraph_text(doc, para, narrative.amount_bucket_sentence(pool, pool_t),
                                         align=WD_ALIGN_LEFT, flush_left=True)
            pt = t["property_type"]
            for para, pool in zip(anchors["proptype"], pools):
                pool_t = pt[pt["pool_type"] == pool]
                _replace_paragraph_text(doc, para, narrative.property_type_sentence(pool, pool_t),
                                         align=WD_ALIGN_LEFT, flush_left=True)
            rg = t["region"]
            for para, pool in zip(anchors["region"], pools):
                pool_t = rg[rg["pool_type"] == pool]
                _replace_paragraph_text(doc, para, narrative.region_sentence(pool, pool_t),
                                         align=WD_ALIGN_LEFT, flush_left=True)

            # --- now rebuild every table's data rows ---
            rows1, dividers1 = _rows_overall_summary(t["overall_summary"], config)
            _rebuild_table(tables(1), rows1)
            _set_pool_column_dividers(tables(1), dividers1)
            _merge_pool_and_category(tables(1), 1 + len(rows1), "합계")

            rows2, dividers2 = _rows_borrower_type(t["borrower_type"], t["overall_summary"])
            _rebuild_table(tables(2), rows2)
            _set_pool_column_dividers(tables(2), dividers2)
            _merge_pool_and_category(tables(2), 1 + len(rows2), "합계")

            for pool, table_idx in zip(pools, amount_indices):
                pool_t = amt[amt["pool_type"] == pool]
                _rebuild_table(tables(table_idx), _rows_amount_bucket(pool_t))

            group_tables = list(proptype_indices) + list(region_data_indices)
            for pool, table_idx in zip(pools, proptype_indices):
                pool_t = pt[pt["pool_type"] == pool]
                rows, _ = _rows_group_breakdown(pool_t, "property_type")
                _rebuild_table(tables(table_idx), rows)

            mx = t["region_x_property"]
            prop_groups = [c for c in mx.columns if c not in ("pool_type", "region_group")]
            for pool, table_idx, matrix_idx in zip(pools, region_data_indices, region_matrix_indices):
                pool_t = rg[rg["pool_type"] == pool]
                rows, top_groups = _rows_group_breakdown(pool_t, "region")
                _rebuild_table(tables(table_idx), rows)
                pool_mx = mx[mx["pool_type"] == pool]
                _rewrite_matrix_header(tables(matrix_idx), prop_groups)
                _rebuild_table(tables(matrix_idx), _rows_matrix(pool_mx, prop_groups, top_groups))

            for table_idx in group_tables:
                _set_row_height(tables(table_idx), GROUP_TABLE_ROW_HEIGHT_CM)

            # --- 2.4/2.5/2.6: 경매/회생/신용보증서 (no narrative) ---
            rows11, dividers11 = _rows_foreclosure(t["foreclosure"])
            _rebuild_table(tables(foreclosure_idx), rows11, skip_highlight_cols={1})
            _set_pool_column_dividers(tables(foreclosure_idx), dividers11, close_bottom=True)

            rows12, dividers12 = _rows_rehab(t["rehab"])
            _rebuild_table(tables(rehab_idx), rows12, skip_highlight_cols={1})
            _set_pool_column_dividers(tables(rehab_idx), dividers12, close_bottom=True)

            guarantee_table = tables(guarantee_idx)
            guarantee_header_rows = _detect_header_rows(guarantee_table, pools)
            _unmerge_header_vertical_cells(guarantee_table, guarantee_header_rows)
            _rebuild_table(guarantee_table, _rows_guarantee(t["guarantee"]), header_rows=guarantee_header_rows)

            # --- keep every table intact across a page break ---
            for table_idx in range(1, guarantee_idx + 1):
                _keep_table_together(tables(table_idx))

            doc.Save()
        finally:
            doc.Close(False)
    finally:
        word.Quit()
        pythoncom.CoUninitialize()
