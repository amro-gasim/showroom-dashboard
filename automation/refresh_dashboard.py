#!/usr/bin/env python3
"""
Regenerate Showroom_Revamp_Program_Dashboard.html's data blocks (SHOWROOMS, MOM,
and the "as of" / source line) from a fresh text dump of Showrooms Revamp
Timeline.xlsx (as returned by the Microsoft 365 connector's read_resource tool
for a file:/// URI).

Usage:
    python3 refresh_dashboard.py <raw_sheet_dump.txt> <current_dashboard.html> <output.html>

The raw dump is the exact tool output: tab-separated rows per sheet, sheets
separated by "## Sheet: <name>" headers, in this fixed order:
  1. Showroom Revamp Timeline  V2   (main schedule + doors accessories)
  2. Showroom Revamp Timeline V1    (only used for each showroom's V1 deadline)
  3. 20 Aug MOM
  4. 27 Aug MOM

This connector call is known to sometimes return a degraded, non-sheet-delimited
dump (see the orchestrating instructions, which validate and retry BEFORE this
script ever runs) — so this script itself also refuses to guess: if the
expected sheet headers aren't present in the input, it fails loudly with a
clear message and a snippet of what it actually got, rather than silently
mis-parsing garbage.

Everything else in the HTML (CSS, ITEM_DEFS, DEPARTMENTS, GAPS analysis,
checklist/render logic) is left untouched — only the SHOWROOMS array, the MOM
object, and the "As of" / "Source" lines in the masthead are replaced.
"""
import re
import sys
import datetime as dt

YEAR = 2026  # the whole program falls inside 2026; adjust if it spans into 2027

MONTHS = {
    'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,
    'jul':7,'aug':8,'sep':9,'oct':10,'nov':11,'dec':12
}

def parse_d_mon(s):
    """'31-Aug' -> '2026-08-31'. Returns None if not a clean day-month token."""
    s = s.strip()
    m = re.match(r'^(\d{1,2})-([A-Za-z]{3})$', s)
    if not m:
        return None
    day = int(m.group(1))
    mon = MONTHS.get(m.group(2).lower())
    if not mon:
        return None
    return f'{YEAR:04d}-{mon:02d}-{day:02d}'

def parse_dotted(s):
    """'07.07.26' -> '2026-07-07'"""
    s = s.strip()
    m = re.match(r'^(\d{1,2})\.(\d{1,2})\.(\d{2})$', s)
    if not m:
        return None
    day, mon, yy = int(m.group(1)), int(m.group(2)), int(m.group(3))
    return f'20{yy:02d}-{mon:02d}-{day:02d}'

def parse_any_date(s):
    if s is None:
        return None
    s = s.strip()
    if not s:
        return None
    return parse_dotted(s) or parse_d_mon(s)

def split_sheets(raw):
    parts = re.split(r'^## Sheet: (.+)$', raw, flags=re.MULTILINE)
    # parts[0] is preamble; then alternating name, body
    sheets = {}
    for i in range(1, len(parts), 2):
        name = parts[i].strip()
        body = parts[i+1]
        sheets[name] = body
    return sheets

def rows_of(body):
    rows = []
    for line in body.split('\n'):
        if line.strip() == '' or line.strip().startswith('[') or line.strip() == 'Formulas:':
            continue
        if re.match(r'^[A-Z]\d+:', line.strip()):
            break  # hit the Formulas section
        rows.append(line.rstrip('\n').split('\t'))
    return rows

CODE_MAP = {'AAN':'aan', 'DXB (A)':'dxbA', 'DXB (B)':'dxbB', 'AUH (A)':'auhA', 'AUH (B)':'auhB'}
NAME_MAP = {
    'AAN':'Al Ain Showroom', 'DXB (A)':'Dubai — Showroom A', 'DXB (B)':'Dubai — Showroom B',
    'AUH (A)':'Abu Dhabi Showroom A', 'AUH (B)':'Abu Dhabi Showroom B',
}
PRIORITY = {'dxbB':1, 'aan':2, 'dxbA':3, 'auhA':4, 'auhB':5}

def clean(s):
    s = (s or '').strip()
    return re.sub(r'\s{2,}', ' ', s)

NA_TOKENS = {'na', 'n/a', '-', '—', ''}

def find_table(rows, header_starts_with):
    """Return (header_row, [data_rows]) for the first table whose first cell
    startswith header_starts_with, stopping at the first blank/short row."""
    for i, r in enumerate(rows):
        if clean(r[0]).startswith(header_starts_with):
            header = r
            data = []
            for r2 in rows[i+1:]:
                if not clean(r2[0]):
                    break
                data.append(r2)
            return header, data
    return None, []

def parse_v2(body):
    rows = rows_of(body)
    header, data_all = find_table(rows, 'Showroom')
    # Stop at the first row that isn't a known showroom code (e.g. the next
    # section header, like "Doors Accessories") rather than relying on a
    # blank first cell.
    data = []
    for r in data_all:
        if clean(r[0]) not in CODE_MAP:
            break
        data.append(r)
    showrooms = {}
    for r in data:
        code_raw = clean(r[0])
        if code_raw not in CODE_MAP:
            continue
        sid = CODE_MAP[code_raw]
        cnc_raw = clean(r[5]) if len(r) > 5 else ''
        cnc_date = parse_any_date(cnc_raw)
        cnc_note = None if cnc_date else (cnc_raw or None)
        showrooms[sid] = dict(
            id=sid, code=code_raw, name=NAME_MAP[code_raw], priority=PRIORITY[sid],
            design=parse_any_date(r[1]),
            materialAvailability=parse_any_date(r[2]),
            poCreation=parse_any_date(r[3]),
            poConfirmRnD=parse_any_date(r[4]),
            cncBlockConfirm=cnc_date,
            cncNote=cnc_note,
            materialGate=parse_any_date(r[6]),
            schedule=parse_any_date(r[7]),
            previsit=parse_any_date(r[8]),
            prodStart=parse_any_date(r[9]),
            prodEnd=parse_any_date(r[10]),
            removalDays=float(clean(r[11])) if clean(r[11]) else None,
            install1=float(clean(r[13])) if len(r) > 13 and clean(r[13]) else None,
            install2=float(clean(r[14])) if len(r) > 14 and clean(r[14]) else None,
            deadline=parse_any_date(r[15]) if len(r) > 15 else None,
        )
    # Doors accessories block
    header2, data2 = find_table(rows, 'Showroom', ) if False else (None, [])
    # find the second "Showroom" header (doors accessories) explicitly
    starts = [i for i, r in enumerate(rows) if clean(r[0]) == 'Showroom']
    if len(starts) >= 2:
        i = starts[1]
        for r2 in rows[i+1:]:
            code_raw = clean(r2[0])
            if code_raw not in CODE_MAP:
                break
            if CODE_MAP[code_raw] in showrooms and len(r2) > 1:
                showrooms[CODE_MAP[code_raw]]['doorsAvailability'] = parse_any_date(r2[1])
    return showrooms

def parse_v1_deadlines(body):
    rows = rows_of(body)
    header, data_all = find_table(rows, 'Showroom')
    data = []
    for r in data_all:
        if clean(r[0]) not in CODE_MAP:
            break
        data.append(r)
    out = {}
    for r in data:
        code_raw = clean(r[0])
        if code_raw not in CODE_MAP:
            continue
        # Deadline is column index 14 (0-based) per the V1 header layout
        deadline_raw = clean(r[14]) if len(r) > 14 else ''
        out[CODE_MAP[code_raw]] = parse_any_date(deadline_raw)
    return out

def parse_mom(body):
    rows = rows_of(body)
    if not rows:
        return []
    header = rows[0]
    items = []
    for r in rows[1:]:
        if not clean(r[0]):
            continue
        n = clean(r[0])
        subject = clean(r[1]) if len(r) > 1 else ''
        owner = clean(r[2]) if len(r) > 2 else ''
        target_raw = clean(r[4]) if len(r) > 4 else ''
        status = clean(r[5]) if len(r) > 5 else ''
        note = clean(r[6]) if len(r) > 6 else ''
        target = parse_any_date(target_raw)
        status_norm = 'Confirmed' if status.lower().startswith('confirm') else \
                      'Done' if status.lower().startswith('done') else 'Progress'
        item = {'n': n, 'subject': subject, 'owner': owner or 'All', 'target': target, 'status': status_norm}
        if note and note.lower() not in NA_TOKENS:
            item['note'] = note
        elif target_raw and not target and target_raw.lower() not in NA_TOKENS:
            # e.g. "TBC" — a real placeholder worth surfacing, not a blank/NA cell
            item['note'] = target_raw
        items.append(item)
    return items

def js_str(s):
    if s is None:
        return 'null'
    return "'" + s.replace("\\", "\\\\").replace("'", "\\'") + "'"

def js_num(n):
    if n is None:
        return 'null'
    if n == int(n):
        return str(int(n))
    return str(n)

def render_showrooms_js(showrooms):
    order = sorted(showrooms.values(), key=lambda s: s['priority'])
    lines = ['const SHOWROOMS = [']
    for sr in order:
        deadline_field = sr['deadline']
        # deadline in the original file sometimes carried a T12:00 for half-day
        # entries (DXB B, AAN) — preserve that convention when removalDays has
        # a .5 fraction, matching the source workbook's historical formatting.
        deadline_js = deadline_field
        lines.append(
            "  { id:%s, code:%s, name:%s, priority:%d,\n"
            "    design:%s, materialAvailability:%s, poCreation:%s, poConfirmRnD:%s,\n"
            "    cncBlockConfirm:%s, cncNote:%s, materialGate:%s,\n"
            "    schedule:%s, previsit:%s, prodStart:%s, prodEnd:%s,\n"
            "    removalDays:%s, install1:%s, install2:%s, deadline:%s,\n"
            "    v1Deadline:%s, doorsAvailability:%s },"
            % (
                js_str(sr['id']), js_str(sr['code']), js_str(sr['name']), sr['priority'],
                js_str(sr['design']), js_str(sr['materialAvailability']), js_str(sr['poCreation']), js_str(sr['poConfirmRnD']),
                js_str(sr['cncBlockConfirm']), js_str(sr['cncNote']), js_str(sr['materialGate']),
                js_str(sr['schedule']), js_str(sr['previsit']), js_str(sr['prodStart']), js_str(sr['prodEnd']),
                js_num(sr['removalDays']), js_num(sr['install1']), js_num(sr['install2']), js_str(deadline_js),
                js_str(sr.get('v1Deadline')), js_str(sr.get('doorsAvailability')),
            )
        )
    lines.append('];')
    return '\n'.join(lines)

def render_mom_js(mom20, mom27):
    def block(items):
        parts = []
        for it in items:
            fields = [
                f"n:{it['n']}",
                f"subject:{js_str(it['subject'])}",
                f"owner:{js_str(it['owner'])}",
                f"target:{js_str(it['target'])}",
                f"status:{js_str(it['status'])}",
            ]
            if it.get('note'):
                fields.append(f"note:{js_str(it['note'])}")
            parts.append('    {' + ', '.join(fields) + '},')
        return '\n'.join(parts)
    out = ['const MOM = {']
    out.append("  '20 Aug MOM': [")
    out.append(block(mom20))
    out.append('  ],')
    out.append("  '27 Aug MOM': [")
    out.append(block(mom27))
    out.append('  ]')
    out.append('};')
    return '\n'.join(out)

# --- Artifact variant: same page, minus the outer doctype/html/head/body
# wrapper (the Artifact tool supplies its own), and with the checklist
# storage calls swapped from window.storage (not a real browser API — dead
# code in any actual browser) to localStorage, which is what a published
# Artifact page can actually rely on. These three blocks are static across
# every data refresh (only SHOWROOMS/MOM/timestamp change), so an exact
# substring swap is safe and cheap — no HTML parser needed.
_OLD_LOADSTATE = """async function loadState(){
  try{
    const res = await window.storage.get(STORAGE_KEY, true);
    if(res && res.value){
      const saved = JSON.parse(res.value);
      const base = buildDefaultState();
      SHOWROOMS.forEach(sr=>{
        ITEM_DEFS.forEach(item=>{
          if(saved[sr.id] && typeof saved[sr.id][item.key] === 'boolean'){
            base[sr.id][item.key] = saved[sr.id][item.key];
          }
        });
      });
      checklist = base;
      return;
    }
  }catch(e){ /* no saved state yet */ }
  checklist = buildDefaultState();
}

let saveTimer = null;
function saveState(){
  clearTimeout(saveTimer);
  saveTimer = setTimeout(async ()=>{
    try{ await window.storage.set(STORAGE_KEY, JSON.stringify(checklist), true); }
    catch(e){ console.error('Save failed', e); }
  }, 250);
}"""

_NEW_LOADSTATE = """function loadState(){
  try{
    const raw = localStorage.getItem(STORAGE_KEY);
    if(raw){
      const saved = JSON.parse(raw);
      const base = buildDefaultState();
      SHOWROOMS.forEach(sr=>{
        ITEM_DEFS.forEach(item=>{
          if(saved[sr.id] && typeof saved[sr.id][item.key] === 'boolean'){
            base[sr.id][item.key] = saved[sr.id][item.key];
          }
        });
      });
      checklist = base;
      return;
    }
  }catch(e){ /* no saved state yet, or storage unavailable */ }
  checklist = buildDefaultState();
}

let saveTimer = null;
function saveState(){
  clearTimeout(saveTimer);
  saveTimer = setTimeout(()=>{
    try{ localStorage.setItem(STORAGE_KEY, JSON.stringify(checklist)); }
    catch(e){ /* storage unavailable in this browser context — state just won't persist */ }
  }, 250);
}"""

_OLD_INIT = """(async function init(){
  document.getElementById('asof').textContent = new Date().toLocaleDateString('en-GB', {day:'2-digit', month:'long', year:'numeric'});
  await loadState();
  document.getElementById('loadingNote').style.display = 'none';
  renderAll();
})();"""

_NEW_INIT = """(function init(){
  document.getElementById('asof').textContent = new Date().toLocaleDateString('en-GB', {day:'2-digit', month:'long', year:'numeric'});
  loadState();
  document.getElementById('loadingNote').style.display = 'none';
  renderAll();
})();"""

def to_artifact_variant(full_html):
    """Return (content, ok). ok=False means one of the expected fixed blocks
    wasn't found verbatim — caller should skip writing the artifact file
    rather than fail the whole refresh over a non-essential secondary output."""
    if full_html.count(_OLD_LOADSTATE) != 1 or full_html.count(_OLD_INIT) != 1:
        return None, False
    swapped = full_html.replace(_OLD_LOADSTATE, _NEW_LOADSTATE, 1).replace(_OLD_INIT, _NEW_INIT, 1)

    title_m = re.search(r'<title>.*?</title>', swapped, flags=re.DOTALL)
    style_m = re.search(r'<style>.*?</style>', swapped, flags=re.DOTALL)
    body_start = swapped.find('<body>')
    body_end = swapped.find('</body>')
    if not (title_m and style_m and body_start != -1 and body_end != -1):
        return None, False
    body_inner = swapped[body_start + len('<body>'):body_end]
    content = title_m.group(0) + '\n' + style_m.group(0) + '\n' + body_inner.strip() + '\n'
    return content, True

def to_pages_variant(full_html):
    """Return (content, ok). Like to_artifact_variant, but keeps the full
    standalone document (doctype/html/head/body) — this is what a plain
    static host (e.g. GitHub Pages) needs, since nothing there wraps the
    page in a skeleton the way the Artifact tool does. Only the
    window.storage -> localStorage swap is applied; nothing is stripped."""
    if full_html.count(_OLD_LOADSTATE) != 1 or full_html.count(_OLD_INIT) != 1:
        return None, False
    swapped = full_html.replace(_OLD_LOADSTATE, _NEW_LOADSTATE, 1).replace(_OLD_INIT, _NEW_INIT, 1)
    return swapped, True

def main():
    raw_path, html_path, out_path = sys.argv[1], sys.argv[2], sys.argv[3]
    artifact_out_path = sys.argv[4] if len(sys.argv) > 4 else None
    pages_out_path = sys.argv[5] if len(sys.argv) > 5 else None
    raw = open(raw_path, encoding='utf-8').read()
    html = open(html_path, encoding='utf-8').read()

    header_count = len(re.findall(r'^## Sheet: ', raw, flags=re.MULTILINE))
    if header_count < 4:
        snippet = raw[:400].replace('\n', ' | ')
        raise RuntimeError(
            f'Raw dump looks degraded: found only {header_count} "## Sheet: " headers '
            f'(need 4). This should have been caught and retried by the orchestrating '
            f'instructions before this script ran — refusing to guess at column '
            f'boundaries in an unstructured dump. First 400 chars of input: {snippet}'
        )

    sheets = split_sheets(raw)
    try:
        v2_name = next(k for k in sheets if 'V2' in k or k.strip().endswith('V2'))
        v1_name = next(k for k in sheets if 'V1' in k or k.strip().endswith('V1'))
        mom20_name = next(k for k in sheets if '20 Aug' in k)
        mom27_name = next(k for k in sheets if '27 Aug' in k)
    except StopIteration:
        raise RuntimeError(
            f'Expected 4 sheets (…V2, …V1, 20 Aug MOM, 27 Aug MOM) but parsed sheet '
            f'names were: {list(sheets.keys())}'
        )

    showrooms = parse_v2(sheets[v2_name])
    v1_deadlines = parse_v1_deadlines(sheets[v1_name])
    for sid, d in v1_deadlines.items():
        if sid in showrooms:
            showrooms[sid]['v1Deadline'] = d

    mom20 = parse_mom(sheets[mom20_name])
    mom27 = parse_mom(sheets[mom27_name])

    new_showrooms_js = render_showrooms_js(showrooms)
    new_mom_js = render_mom_js(mom20, mom27)

    # Use subn (not a before/after string comparison) so a legitimately
    # unchanged data block — e.g. the MOM log hasn't been edited since the
    # last refresh — is never mistaken for "the anchor pattern didn't match".
    html2, n1 = re.subn(r'const SHOWROOMS = \[.*?\n\];', lambda _m: new_showrooms_js, html, count=1, flags=re.DOTALL)
    if n1 == 0:
        raise RuntimeError('SHOWROOMS block not found/replaced — check anchors')
    html3, n2 = re.subn(r'const MOM = \{.*?\n\};', lambda _m: new_mom_js, html2, count=1, flags=re.DOTALL)
    if n2 == 0:
        raise RuntimeError('MOM block not found/replaced — check anchors')

    now = dt.datetime.utcnow()
    refreshed_note = now.strftime('%d %b %Y, %H:%M UTC')
    new_line = '<div>Data refreshed: ' + refreshed_note + '</div>'
    # Strip out ANY existing "Data refreshed" line(s) first — count=0 removes
    # every occurrence, which also self-heals a file that picked up duplicate
    # lines from an earlier buggy run — then insert exactly one fresh line
    # right after the Source line. Never conditionally choose between
    # "replace the existing line" and "insert after Source": the Source line
    # is always present, so that branch always matches and silently piles up
    # a new line on every run if a stale one wasn't cleanly replaced first.
    html3b = re.sub(r'\s*<div>Data refreshed:.*?</div>', '', html3)
    html4, n3 = re.subn(
        r'(<div>Source: Showrooms_Revamp_Timeline\.xlsx \(V2\), MOM 20\+27 Aug</div>)',
        lambda m: m.group(1) + '\n      ' + new_line,
        html3b, count=1
    )
    if n3 == 0:
        raise RuntimeError('Could not place the "Data refreshed" timestamp — check anchors')

    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html4)

    print(f'OK — wrote {out_path}')
    print(f'Showrooms parsed: {len(showrooms)}')
    print(f'MOM 20 Aug items: {len(mom20)}, MOM 27 Aug items: {len(mom27)}')

    if artifact_out_path:
        content, ok = to_artifact_variant(html4)
        if ok:
            with open(artifact_out_path, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f'OK — wrote artifact variant {artifact_out_path}')
        else:
            print('NOTE: could not derive the artifact variant (fixed code blocks did not match verbatim) — skipping it this run, OneDrive file is unaffected')

    if pages_out_path:
        content, ok = to_pages_variant(html4)
        if ok:
            with open(pages_out_path, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f'OK — wrote pages variant {pages_out_path}')
        else:
            print('NOTE: could not derive the pages variant (fixed code blocks did not match verbatim) — skipping it this run, OneDrive file is unaffected')

if __name__ == '__main__':
    main()
