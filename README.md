# Debt Dashboard (บริหารหนี้ฝ่าย ชบ.)

Single-file HTML dashboard (`index.html`) hosted on GitHub Pages
(`https://surathin3320500281957-cmd.github.io`, repo `Debt_Dashboard`).
No backend: the data is gzip+base64 embedded in the HTML. Vanilla JS +
Chart.js 4.4.1 + chartjs-plugin-datalabels 2.2.0 + xlsx (SheetJS) via CDN.

**Starting a new chat: upload `index.html`, this README and `build_data.py`.**

Latest state (พ.ย. 2569): 18,822 contracts, data built from the new
3-sheet export (§3). Language of UI and of the user: Thai.

---

## 1. What the dashboard does

One page only (บริหารหนี้ตามสัญญา) over 4 contract types: เช่า, เช่าซื้อ,
เช่าจัดประโยชน์, เช่าซื้อ-คบส. — KPIs, ranking charts, filters, filterable
detail table, Excel/PDF export.

View modes (`DV`): **เดือนนี้** (`curr`), **เดือนก่อน** (`prev`),
**เปรียบเทียบ** (`cmp`). Each month shows its own contracts (rows are the
*union* of both months, field 20 `ch` = `B` both / `N` new this month /
`X` closed = prev only; `applyF()` hides `X` in curr and `N` in prev).
Only in cmp: filter `f-ch` (การเปลี่ยนแปลงสัญญา), 3 clickable summary cards
(`#d-chg`, `setCh()`), and a table/export column.
Also only in cmp: filter `f-sc` (เปลี่ยนแปลงสถานะบริหารหนี้: เปลี่ยนสถานะ / ไม่เปลี่ยนสถานะ,
compares `sc` vs `sp`, applies to `B` rows only — new/closed contracts are excluded from both
options), 2 clickable cards (`#d-chs`, `setSc()`), and the column สถานะบริหารหนี้(ก่อน)
before สถานะบริหารหนี้ (table + export). `applyF()` uses `pass(r,skip)` so each card set ignores its own filter. New = ทำสัญญาใหม่,
X = ปิดสัญญา (confirmed by the user).

**Removed on request (do not re-add unless asked):**
- DPD ธอส. tab (no `PB64`, no `PR`, no nav bar, no DPD sheet).
- "การติดตาม" (follow-up) filter, badge, table columns, export columns.
  The values still exist in the data array (fields 16/17) but are unused.

---

## 2. File map

| File | Purpose |
|---|---|
| `index.html` | Entire app + embedded data blob `DB64` |
| `build_data.py` | Turns the new Excel into `DR_new.json` (see §3) |
| `.nojekyll` | Empty file at repo root. **Required**: without it Pages runs Jekyll (8–9 min deploys). |

No build step; push `index.html` to `main`.

---

## 3. Updating data from a new Excel export (CURRENT format)

Source file name looks like `...Dashboard.xlsx`, sheets:

| Sheet | Content |
|---|---|
| `เดือนนี้` | one row per contract, current month (59 columns, English headers) |
| `เดือนก่อน` | same layout, previous month |
| `ฟ้องปีก่อน` | `CUST_CODE`, `DM_STATUS_NAME` — customers sued in the previous year |
| `โครงการ` | `PROJ_CODE`, `PROJ_NAME`, `BP_ABBR_NAME`, `ฝ่าย`, `กอง`, `สำนักงาน`, `ประเภทโครงการ` |

(An old format with 4 contract sheets + header-row totals + `DPDธอส` existed
until Nov 2569; it is obsolete.)

### 3.1 Field mapping → one row of `DR` (22 fields)

Rows = **union** of `เดือนนี้` and `เดือนก่อน`, joined by `CUST_CODE`
(unique in every sheet). Missing side gets 0 / '' and `ch` flag (field 20:
`B`/`N`/`X`). Current view = B+N, prev view = B+X, so each month's totals
match its own sheet.

| idx | key (`D.`) | source |
|---|---|---|
| 0 | `tp` | `CONTRACT_TYPE_CODE`: 20 เช่าซื้อ, 50 เช่า, 70 เช่าจัดประโยชน์, 80 `เช่าซื้อ-คบส.` (only these 4 are accepted; script asserts) |
| 1 | `dv` กอง | sheet `โครงการ`.`กอง` via `PROJ` = `PROJ_CODE` |
| 2 | `br` สาขา | sheet `โครงการ`.`สำนักงาน` via `PROJ` |
| 3 | `pj` | `PROJ_NAME` |
| 4 | `pl` | `HOUSE_NO_STD` with leading zeros stripped |
| 5 | `hs` | `HOUSE_NO_ACT` |
| 6 / 7 | `sc` / `sp` สถานะบริหารหนี้ | `DM_STATUS_NAME` this month / previous month |
| 8 / 9 | `rc` / `rp` | range label from `PERIOD` (0 ไม่ค้าง, 1-2 ค้าง 1-2 งวด, 3 ค้าง 3 งวด, >3 เกิน 3 งวด); `rp` '' if no prev row |
| 10 / 11 | `ic` / `ip` งวดค้าง | `PERIOD` (blank = 0) this / previous |
| 12 / 13 | `pc` / `pp` หนี้ค้างรวม | `ALL_PRICE` (rounded to whole baht) |
| 14 / 15 | `kc` / `kp` หนี้ตัวชี้วัด | **rule below** |
| 16–19 | unused | 'ยังไม่ติดตาม','ยังไม่ติดตาม','','' |
| 20 | `ch` | `B` both months, `N` only this month, `X` only previous month |
| 21 | `cc` | `CUST_CODE` (รหัสลูกหนี้), shown after บ้านเลขที่ in the table and exports |

**หนี้ตัวชี้วัด rule (confirmed by the user):**
`CUST_CODE` found in sheet `ฟ้องปีก่อน` → 0 (still counted in หนี้ค้างรวม);
not found → `M4_M6_PRICE + M7_M12_PRICE + M13_M24_PRICE + MORE24_PRICE`.
The same `ฟ้องปีก่อน` list is applied to both months. Customers sued
*this* year are not in that sheet, so they are still counted.
`ALL_PRICE` equals M1_M2+M3+M4_M6+M7_M12+M13_M24+MORE24 on every row.

`rngMap()` in the page merges 'ค้าง 3 งวด' and 'เกิน 3 งวด' into
**'เกิน 2 งวด'** for display/filter/export — the data keeps the old labels.

### 3.2 Recipe

```bash
python3 -I build_data.py "/path/to/new Dashboard.xlsx"   # writes DR_new.json
```
then embed (gzip → base64, one line replace of `DB64="..."`):
```python
import re,json,gzip,base64
c=open('index.html',encoding='utf-8').read()
new=json.load(open('DR_new.json'))
b=base64.b64encode(gzip.compress(json.dumps(new,ensure_ascii=False,separators=(',',':')).encode())).decode()
c=re.sub(r'DB64="[^"]*"',lambda m:'DB64="'+b+'"',c,count=1)
open('index.html','w',encoding='utf-8').write(c)
```
Then update the static labels: header badge `N รายการ` (`class="hbg"`) and
the month badge (`พ.ย. 2569`) — the month comes from the date stamp in the
file/CREATE_DATE. Decoder in the page: `dcmp()` using `DecompressionStream('gzip')`.

### 3.3 Validation checklist (do before shipping, report exact numbers)

1. Row count = rows in `เดือนนี้`; types ⊂ {20,50,70,80}; all `PROJ` found in `โครงการ`.
2. Σ`pc` = Σ`ALL_PRICE` of `เดือนนี้` (±rounding); prev total vs `เดือนก่อน`.
3. Σ`kc` recomputed independently from the rule above.
4. Diff against the previously embedded blob (decode `DB64` of the old
   `index.html`, match on `(pj,pl,hs)`): กอง/สาขา/ประเภท/สถานะ/งวด must match
   for continuing contracts; money may differ ≤ 0.5 baht (rounding) or by real changes.
5. Never assume "same date = same content" — always diff.
6. Playwright render test (below). Known harmless console error: Google Fonts blocked offline.

Reference result (Nov 2569 file): 18,931 rows = 18,780 B + 42 N + 109 X;
หนี้ค้างรวม 62,058,482 / ก่อน 60,064,204;
หนี้ตัวชี้วัด 2,556,493 / ก่อน 2,084,703; 871 sued-last-year rows zeroed;
626 rows with indicator debt.

### 3.4 Render test

Playwright (use pre-installed `/opt/pw-browsers/chromium`, `--no-sandbox`),
serve Chart.js / datalabels / xlsx from local `node_modules` via
`page.route`; wait ~6 s; call `setView('curr'|'prev'|'cmp')` and check
`DF.length`, KPI text (`.kc`), ranking rows, `exportXL('detail')`
download, no `pageerror`.

---

## 4. Ranking charts

The 3 rankings (`ch-rk1` โครงการ หนี้รวม, `ch-rk2` ตัวชี้วัด, `ch-rkp`
แปลง/โครงการ) are **HTML/CSS, not canvas** (`mkRank(id, labels, values)`):
rank number | capsule bar (hue by rank, width ∝ value) with a purple name
badge | dark value badge on the right. Fonts 16–17 px (executives with poor
eyesight — don't shrink). Cards are `auto-fit minmax(520px)` so names don't
truncate; long names wrap. Trade-off: bars are not click-to-drill any more.
Other charts still use Chart.js (`mkC`). `index-baa.html` (separate
dashboard) was deliberately left unchanged.

---

## 5. Filter system

- **Dropdowns**: `f-type, f-div, f-br, f-pj, f-st, f-rng, f-inst, f-kpi`
  plus `f-trend` (debt amount `pc-pp`), `f-trendi` (installments
  `ic-ip`) and `f-ch` (contract change), the last three visible only in เปรียบเทียบ. Read in `applyF()`.
- **Click-drill** `DS` (`st, br, pj, rng, plan`) from charts; ANDed with dropdowns.
- `f-inst`: ไม่ค้าง(0), 1..24, `25+` (cap at 24 — long tail up to 254).
- **Badges**: `bdg()` renders a colored dismissible chip per active filter
  into `#d-bdg` (`.fchip` + `.fc-*` classes); `clrF(id)` / `clrF('ds:key')`
  clears one; `resetF()` clears all.

Detail table: เดือนนี้ 18 columns (incl. 5 "(ก่อน)" columns after สถานะ),
เดือนก่อน 13, เปรียบเทียบ 18 (incl. สถานะ(ก่อน); last = การเปลี่ยนแปลงสัญญา); the Excel export has the same columns.
"รายละเอียด" column exists but is empty (the new file has no such column).

---

## 6. GitHub Pages

- Source: GitHub Actions "pages build and deployment", branch `main`.
- `.nojekyll` at repo root is required.
- "Deployment failed, try again later" = transient GitHub issue; re-commit.
- Upload: repo → Add file → Upload files → replace `index.html` → commit to `main`.
- The user often works from a tablet: give browser-only steps.

---

## 7. Decisions / history worth knowing

- Larger fonts for executives (§4) — keep.
- Dynamic badge-offset experiment was reverted; user prefers simple.
- หนี้ตัวชี้วัด = หนี้อายุเกิน 3 เดือน excluding customers sued last year
  (§3.1). Earlier (pre-Nov 2569) the number came from a column in the old
  Excel; the new rule reproduces it to within ~3 baht.
- Always validate against the source before reporting totals; an earlier
  round showed a wrong หนี้ตัวชี้วัด because the wrong header cell was used.
