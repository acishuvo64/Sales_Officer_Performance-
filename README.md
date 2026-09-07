# AI Sales Force Tracking Dashboard

Sales Officer der daily Activity + Attendance Excel file theke ekta full
Streamlit dashboard — **kono upload chara**. Director/Boss shudhu link
open korlei shesh update kora data dekhte pabe.

## Design

Dark navy sidebar + gradient KPI cards + gradient header title — screenshot
onujayi style kora hoyeche. Company naam/logo/color shob customizable
(niche dekhun).

## Daily Workflow (No Upload!)

1. Notun Excel file-ta ei naam-e rakhun: **`data/SO_Activity_Attendance.xlsx`**
   (purono file-ta overwrite kore dite hobe — sheet names obossoi
   "SO Activity" o "Attendance" thakte hobe).
2. GitHub repo-te commit + push korun:
   ```bash
   git add data/SO_Activity_Attendance.xlsx
   git commit -m "Update data - <date>"
   git push
   ```
3. Streamlit Community Cloud automatically redeploy hoye notun data show
   korbe (usually 30-60 seconds er moddhe). Boss-ke kichu korte hobe na —
   shudhu link refresh korlei notun data dekhbe.

## Branding Customize Korte (logo/company naam)

`dashboard.py` file-er upore ei 3-ta line change korun:

```python
COMPANY_NAME = "Your Company Name"
DASHBOARD_TITLE = "AI Sales Force Tracking Dashboard"
DASHBOARD_TAGLINE = f"{COMPANY_NAME} — Sales Officer Activity, Attendance & Productivity Analytics"
```

Nijer **logo** dite chaile:
- `assets/logo.png` → sidebar-er upore-r white card-e dekhabe (rectangular logo)
- `assets/badge.png` → circular badge hishebe sidebar o header-er dan-pashe dekhabe

Logo file na dile, automatically company-naam diye ekta styled placeholder
dekhabe — kono error hobe na.

## Ki ki dekhabe

1. **Attendance & Late** — proti SO er koi din Late, koi din on-time Present,
   Total Late %, group-wise average Late %.
2. **First Visit & Market Stay** — average first check-in time, o average
   market/field stay time (checkout recorded thakle) — full note app-e deওয়া
   ache.
3. **Performance (Outlet Visit)** — Active Days × 20 = Target, Achievement %,
   60% er niche hole 🔴 red mark, Top 10 Good / Bad Performer chart.
4. **Group & RSM Charts** — Group-wise (Kings/Royals/Captain) pie + column
   chart, RSM-wise Top/Low performer bar chart.
5. **Route & Repeat-Market Analysis** — SO protidin kon route e sob theke
   beshi gyeche, ar kon market bar bar (assigned Route Day-r baire) visit
   hocche.
6. **Full Data Table** — shob metric ekshathe, CSV export kora jay.

Sidebar-e Start Date / End Date sorasori dekha jay; "More Filters" expander-e
Group/Region/Zone/RSM filter kora jay.

## Kivabe Run Korben (local test)

```bash
pip install -r requirements.txt
streamlit run dashboard.py
```

Browser-e `http://localhost:8501` open hobe.

## Deploy Korben Kivabe (GitHub + Streamlit Community Cloud)

1. Ei folder-er shob file (`dashboard.py`, `requirements.txt`,
   `.streamlit/config.toml`, `data/SO_Activity_Attendance.xlsx`, `assets/`)
   ekta GitHub repo-te push korun.
   > **Note:** `.streamlit` folder ta hidden thakte pare file explorer-e —
   > kintu eta obossoi push korte hobe, karon eikhane theme/dark-sidebar
   > color-er config ache.
2. https://share.streamlit.io e giye GitHub diye login korun.
3. "New app" → repo select → main file `dashboard.py` → Deploy.
4. Ekta permanent link paben (e.g. `https://your-app.streamlit.app`) —
   eta Boss/Director-ke diye din. Tar kono upload/login lagbe na, shudhu
   link open korলেই dekhte pabe.

> Free tier-e app kichu shomoy inactive thakle "sleep" hote pare — first
> open-e kichu shomoy (10-20 sec) lagte pare, eta normal.

## Assumption gulo (Please Review)

- SO Activity sheet-e kono visit-time na thakay, **"First Visit"** = din-er
  prothom attendance check-in time (proxy). Alada kono time-field thakle
  janaben, dashboard update kore deওয়া jabe.
- **"Market Stay / Idle Time"** = Attendance sheet-er Total Hours column
  (checkout recorded thakle) — onek din checkout na thakay shei din bad
  deওয়া hoyeche (Checkout Days column-e koto din er data ta dekhano ache).
- **Good Performer** = Avg daily outlet visit ≥ 20; **Bad Performer** =
  Avg daily outlet visit < 15 — apnar dea criteria onujayi.
- **Target Achievement %** = Total Outlet Visit ÷ (Active Days × 20) × 100;
  60% er niche hole red mark.

Ei threshold gulo (`DAILY_OUTLET_TARGET`, `BAD_PERFORMER_OUTLET_LIMIT`,
`ACHIEVEMENT_RED_THRESHOLD`) `dashboard.py` file-er upore easily change kora
jay.
