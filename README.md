# Debt Dashboard (บริหารหนี้ฝ่าย ชบ.)

Single-file HTML dashboard (`index.html`) hosted on GitHub Pages at
`https://surathin3320500281957-cmd.github.io`. No backend — all data is
compressed and embedded directly in the HTML file. Built with vanilla
JS + Chart.js (via CDN).

**When starting a new chat about this project, upload `index.html` plus this
README so Claude has full context immediately.**

---

## 1. What this dashboard does

Two tabs:

1. **บริหารหนี้ตามสัญญา** — debt across 4 contract types (เช่า, เช่าซื้อ,
   เช่าจัดประโยชน์, เช่าซื้อคบส): totals, KPI ("หนี้ตัวชี้วัด"), ranking
   charts, filterable detail table, Excel/PDF export.
2. **DPD ธอส.** — days-past-due tracking sourced from a separate sheet.

Three view modes on the debt tab: **เดือนนี้** (current), **เดือนก่อน**
(previous), **เปรียบเทียบ** (compare both side by side, deltas).

---

## 2. File map

| File | Purpose |
|---|---|
| `index.html` | The entire app — HTML/CSS/JS + embedded gzip+base64 data blobs (`DB64`, `PB64`) |
| `.nojekyll` | Empty file at repo root. **Required** — without it GitHub Pages runs Jekyll on the file and deploys take 8–9 min instead of ~10–25s. Already in the repo; don't remove it. |

No build step. Editing `index.html` and pushing to `main` is the entire
deploy process.

---

## 3. Updating data from a new Excel export

The source file is always named like `บริหารหนี้....xlsx` with sheets:
`เช่า`, `เช่าซื้อ`, `เช่าจัดประโยชน์`, `เช่าซื้อคบส`, `DPDธอส` (a `รวมชี้แจง`
sheet sometimes appears too — it's a different, unrelated data shape and is
**not** currently ingested).

### 3.1 Contract sheets → `DR` array → `DB64`

Each contract sheet has **merged headers spanning rows 0–2, data starts at
row 7** (`skiprows=7` in pandas). Key column indices (0-based, same across
all 4 contract sheets):

| Excel col | Meaning |
|---|---|
| 1 | ชนิดสัญญา (tp) |
| 3 | กอง (dv) |
| 4 | สาขา (br) |
| 5 | โครงการ (pj) |
| 6 | เลขผัง (pl) |
| 7 | เลขบ้าน (hs) |
| 14 | (เดือนก่อน) งวดค้าง |
| 15 | (เดือนก่อน) กรณีฟ้อง |
| 16 | (เดือนก่อน) ช่วงงวดค้าง |
| 17 | (เดือนก่อน) งวดค้าง — ⚠ same label as col 14, this is the numeric one used |
| 18 | (เดือนก่อน) หนี้ค้าง |
| 19 | (เดือนก่อน) หนี้อายุเกิน 3 งวด |
| 20 | (เดือนก่อน) หนี้ตัวชี้วัด |
| 21 | (เดือนก่อน) การติดตาม |
| 22 | (เดือนก่อน) รายละเอียดการติดตาม |
| 23 | (เดือนนี้) สถานะบริหารหนี้ |
| 25 | (เดือนนี้) ช่วงงวดค้าง |
| 26 | (เดือนนี้) งวดค้าง |
| 27 | (เดือนนี้) หนี้ค้าง |
| 28 | (เดือนนี้) อายุหนี้เกิน 3 งวด |
| 29 | (เดือนนี้) หนี้ตัวชี้วัด |
| 30 | (เดือนนี้) การติดตาม |
| 31 | (เดือนนี้) รายละเอียดการติดตาม |

**⚠ Known trap:** columns 28 ("หนี้อายุเกิน 3 งวด") and 29 ("หนี้ตัวชี้วัด")
sit right next to each other and were confused once before (shipped wrong
totals). **Always validate `kc`/`kp` sums against the sheet's own header
row** before embedding — see §3.3.

Each output row of `DR` (order matters — matches the `D` index map in
`index.html`):

```python
row = [S(r[1]),S(r[3]),S(r[4]),S(r[5]),S(r[6]),S(r[7]),
       S(r[23]),S(r[14]),S(r[25]),S(r[16]),
       N(r[26]),N(r[17]),N(r[27]),N(r[18]),N(r[29]),N(r[20]),
       S(r[30]),S(r[21]),S(r[31]),S(r[22])]
```

`S()` = string-clean (strip, blank for NaN, drop trailing `.0` on numbers
stored as float). `N()` = numeric-clean, round to 2dp, 0 for NaN.

In `index.html`, `const D={tp:0,dv:1,br:2,pj:3,pl:4,hs:5,sc:6,sp:7,rc:8,
rp:9,ic:10,ip:11,pc:12,pp:13,kc:14,kp:15,tc:16,tp2:17,dc:18,dp2:19,pf:20}`
— `*c`/`*2`/no-suffix = เดือนนี้ (current), `*p` = เดือนก่อน (previous).

### 3.2 DPD sheet → `PR` array → `PB64`

Header spans rows 0–4, data starts at row 5 (`skiprows=5`). Columns:
1=กอง(dv), 2=สาขา(br), 3=เลขผัง(ln), 4=เดือน(mo), 5=โครงการ(pj),
6=เลขบ้าน(hs), 7=คำนำหน้า(la), 8=ชื่อ(nm), 13=(เดือนนี้)หนี้ค้าง(pr),
16=(เดือนก่อน)หนี้ค้าง(dp), 17=(เดือนก่อน)ช่วง DPD(rp),
18=(เดือนก่อน)การติดตาม(tp), 15=(เดือนนี้)อายุหนี้(dep),
20=(เดือนนี้)หนี้ค้าง(dc), 21=(เดือนนี้)ช่วง DPD(rc),
22=(เดือนนี้)การติดตาม(tc), 23=รายละเอียด(dec_).

`const P={dv:0,br:1,ln:2,mo:3,pj:4,hs:5,nm:6,la:7,pr:8,dp:9,rp:10,tp:11,
dep:12,dc:13,rc:14,tc:15,dec_:16}`.

DPD range strings are like `"DPD 1-30 วัน"`, `"DPD 61-90 วัน"`,
`"DPD เกิน 90 วัน"` — full string including the `"DPD "` prefix.

### 3.3 Build + validate + embed (the actual repeatable recipe)

```python
import pandas as pd, math, json, gzip, base64

xl = pd.ExcelFile('บริหารหนี้....xlsx')

def S(v):
    if v is None or (isinstance(v,float) and math.isnan(v)): return ''
    if isinstance(v,float) and v.is_integer(): return str(int(v))
    return str(v).strip()
def N(v):
    if v is None or (isinstance(v,float) and math.isnan(v)): return 0
    try: return round(float(v),2)
    except: return 0

# --- validate BEFORE trusting any total: compare column sums against
#     each sheet's own header row (row idx 0/1/2, col 15 = เดือนก่อน value,
#     col 30 = เดือนนี้ value) ---
metrics = [('หนี้ค้างรวม (ก่อน)',0,15,18), ('หนี้ค้างรวม (นี้)',0,30,27),
           ('เกิน 3 เดือน (ก่อน)',1,15,19), ('เกิน 3 เดือน (นี้)',1,30,28),
           ('หนี้ตัวชี้วัด (ก่อน)',2,15,20), ('หนี้ตัวชี้วัด (นี้)',2,30,29)]
for s in ['เช่า','เช่าซื้อ','เช่าจัดประโยชน์','เช่าซื้อคบส']:
    hdr = pd.read_excel(xl, sheet_name=s, header=None, nrows=3)
    df  = pd.read_excel(xl, sheet_name=s, header=None, skiprows=7)
    for name,hr,hc,dc in metrics:
        exp, got = N(hdr.iloc[hr,hc]), sum(N(v) for v in df[dc])
        assert abs(exp-got) < 1, f'{s} {name}: header={exp} got={got}'

# --- build DR ---
DR = []
for s in ['เช่า','เช่าซื้อ','เช่าจัดประโยชน์','เช่าซื้อคบส']:
    df = pd.read_excel(xl, sheet_name=s, header=None, skiprows=7)
    for r in df.itertuples(index=False):
        r = list(r)
        DR.append([S(r[1]),S(r[3]),S(r[4]),S(r[5]),S(r[6]),S(r[7]),
                   S(r[23]),S(r[14]),S(r[25]),S(r[16]),
                   N(r[26]),N(r[17]),N(r[27]),N(r[18]),N(r[29]),N(r[20]),
                   S(r[30]),S(r[21]),S(r[31]),S(r[22])])

# --- build PR ---
PR = []
df = pd.read_excel(xl, sheet_name='DPDธอส', header=None, skiprows=5)
for r in df.itertuples(index=False):
    r = list(r)
    PR.append([S(r[1]),S(r[2]),S(r[3]),S(r[4]),S(r[5]),S(r[6]),S(r[8]),S(r[7]),
               N(r[13]),N(r[16]),S(r[17]),S(r[18]),N(r[15]),N(r[20]),
               S(r[21]),S(r[22]),S(r[23])])

db64 = base64.b64encode(gzip.compress(json.dumps(DR,ensure_ascii=False,separators=(',',':')).encode())).decode()
pb64 = base64.b64encode(gzip.compress(json.dumps(PR,ensure_ascii=False,separators=(',',':')).encode())).decode()

# --- embed into index.html ---
import re
with open('index.html', encoding='utf-8') as f: c = f.read()
c = re.sub(r'DB64="[^"]*"', 'DB64="'+db64+'"', c, count=1)
c = re.sub(r'PB64="[^"]*"', 'PB64="'+pb64+'"', c, count=1)
with open('index.html', 'w', encoding='utf-8') as f: f.write(c)
```

Also update the static header labels if the record counts / month changed:
`<span class="hbg">18,589 รายการ</span>`, the DPD count badge, and the
month badge (`มิ.ย. 2569` etc — DPD sheet's title row says the month, e.g.
`หนี้ค้าง DPD ธอส. ส.ค.69`).

**Before deciding data "didn't change," diff it** — two uploads with the
same `Contract_Active_Dm` date stamp have genuinely differed before (one
had 362 rows with retroactively-updated follow-up notes; another had
"หนี้ตัวชี้วัด (เดือนก่อน)" flip from all-zero to real values). Compare
row-by-row against the previously-embedded blob, don't assume identical
dates mean identical content.

### 3.4 After embedding: always render-test before shipping

Headless-render `index.html` (Playwright; a chromium binary may already be
at `/opt/pw-browsers` — set `PLAYWRIGHT_BROWSERS_PATH=/opt/pw-browsers`
before falling back to `npx playwright install`, which is usually blocked
by the sandbox's network allowlist). CDN scripts (Chart.js, datalabels
plugin, Google Fonts) are blocked in the sandbox — swap the two
`<script src="https://cdnjs...">` tags for local copies
(`npm install chart.js@4.4.1 chartjs-plugin-datalabels@2.2.0`, then point
the `<script>` tags at the local files) for local testing only; ship the
original CDN-referencing file.

Check: `DR.length` / `PR.length` match expected row counts, the 4 KPI cards
match the validated totals, switch to DPD tab and check its KPI row,
switch เดือนนี้/เดือนก่อน/เปรียบเทียบ and re-check, no console errors besides
the expected sandbox CDN 403s.

---

## 4. Ranking chart label design (why it looks the way it does)

The three "Ranking …" bar charts and the two DPD ranking charts don't use
Chart.js's normal axes. Long Thai labels used to get clipped to 2–3
characters when the page's 3-column grid squeezed each chart's y-axis on
mobile. Fix, iterated over several rounds:

- Both axes hidden (`scales:{x:{display:false},y:{display:false}}`).
- `chartjs-plugin-datalabels` draws the **item name inside the bar at the
  start** (`anchor:'start',align:'end',offset:8`, dark text + white stroke
  for contrast against any bar color) and the **value as a solid dark
  rounded badge right at the bar's tip** (`anchor:'end',align:'end',
  offset:6`, white text — keep this a flat number; an earlier dynamic
  offset-to-avoid-overlap formula was tried and explicitly rejected by the
  user as "too clever" — the opaque badge already reads fine even when it
  slightly overlaps the name text).
- `layout.padding.right` (110px) reserves room so the largest bar's badge
  doesn't clip off the canvas edge.
- Font size 14 / bold, container heights 460px (main 3) / 360px (DPD 2) —
  bumped up from the original 8px/310px for an executive audience with
  poor eyesight.
- `Chart.defaults.set('plugins.datalabels',{display:false})` globally,
  then explicitly `display:true` per-chart — **don't forget the explicit
  `display:true`**, it's easy to leave it inheriting the global default
  and get silently blank labels.

Truncation (`cl(label, N)`) is still applied per chart (~20–30 chars) since
labels overlay a finite-width canvas even without a reserved axis column.

---

## 5. Filter system (บริหารหนี้ตามสัญญา tab)

Two parallel filter mechanisms feed the same `DF` (filtered dataset):

- **Dropdowns** (`f-type`, `f-div`, `f-br`, `f-pj`, `f-st`, `f-tr`,
  `f-rng`, `f-inst`, `f-trend`, `f-trendi`, `f-kpi`) — top filter bar,
  read in `applyF()`.
- **Click-to-drill** (`DS.st`, `DS.br`, `DS.pj`, `DS.rng`, `DS.plan`) —
  clicking a bar/pie segment in a chart sets one of these; also read in
  `applyF()`. Both mechanisms AND together.

`f-inst` ("งวดค้าง", exact installment count) and `f-trend`/`f-trendi`
("แนวโน้ม…") are worth knowing about if extending this further:

- `f-inst` options are capped at 24 explicit values + a `"25+"` bucket
  (`value="25+"`) because the real installment-count distribution has a
  long tail out to 254 with only ~200 records above 20 — don't populate
  one `<option>` per distinct value found in the data, it balloons to
  250+ entries.
- `f-trend` compares **debt amount** (`pc - pp`); `f-trendi` compares
  **installment count** (`ic - ip`). These are genuinely different signals
  (a partial payment can shrink the amount owed without changing the
  installment count) — both exist side-by-side on purpose, not
  redundant. Only shown in "เปรียบเทียบ" view mode.

**Active-filter badges**: `bdg()` renders one colored, individually-
dismissible chip per active filter (dropdown or click-drill) into
`#d-bdg`, styled via `.fchip` + a per-category color class (`.fc-type`,
`.fc-div`, `.fc-br`, `.fc-pj`, `.fc-st`, `.fc-tr`, `.fc-rng`, `.fc-inst`,
`.fc-trend`, `.fc-trendi`, `.fc-kpi`, `.fc-ds` for click-drill). Each
chip's `✕` calls `clrF(id)` (dropdown) or `clrF('ds:'+key)` (click-drill)
to clear just that one filter and re-run `applyF()`.

**DPD ธอส. tab currently has no filters at all** — nothing to badge there.
If asked to add filters there too, that's new work, not a copy of the
existing system.

---

## 6. GitHub Pages deployment notes

- Repo: `surathin3320500281957-cmd.github.io`, deploy source = GitHub
  Actions ("pages build and deployment" workflow), branch `main`.
- `.nojekyll` must exist at repo root — without it, Jekyll processes the
  large single-file HTML unnecessarily and deploys take 8–9 minutes
  (still succeeds, just slow).
- Occasional single-run failures with `Error: Deployment failed, try
  again later.` are a transient GitHub-side issue, not a file problem —
  re-push (even a no-op change) to trigger a fresh run.
- To upload a new `index.html`: repo → Code → open `index.html` → edit
  (pencil icon) or **Add file → Upload files** to replace it wholesale →
  Commit directly to `main`.

---

## 7. History / decisions worth knowing before changing things again

- Font/layout was bumped up once already for readability (§4) — don't
  shrink it back without asking.
- The dynamic-offset badge-positioning experiment (computing offset from
  measured text width to avoid name/value overlap) was built, verified
  working, then **explicitly reverted** in favor of a flat `offset:4`
  because the user found the simple version equally readable and less
  fragile. Don't reintroduce dynamic positioning unless asked.
- "หนี้ตัวชี้วัด" (indicator debt) is legally/definitionally **หนี้อายุ
  เกิน 3 งวด, minus cases sued before the current fiscal year** — the
  Excel formula already bakes this in (col 29/20), don't recompute it
  from other columns.
- Same-dated Excel re-uploads have contained real content changes twice
  now (§3.3) — always diff, never assume "same date = skip."
