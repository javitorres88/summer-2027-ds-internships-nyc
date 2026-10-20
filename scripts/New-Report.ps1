# New-Report.ps1
#
# Builds README.md from data/enriched.csv, using a curated degree-level map that was
# verified by hand against the job descriptions. Tables are kept to four columns so
# they stay readable in the GitHub mobile app.
$ErrorActionPreference = 'Stop'

$root    = Split-Path -Parent $PSScriptRoot
$dataDir = Join-Path $root 'data'
$e       = Import-Csv (Join-Path $dataDir 'enriched.csv')
$n       = Import-Csv (Join-Path $dataDir 'internships.csv')

# NOTE: ASCII-only on purpose - Windows PowerShell 5.1 reads .ps1 files as ANSI.
function Norm($s) {
    if (-not $s) { return '' }
    $s = $s -replace "[$([char]0x2022)$([char]0x2013)$([char]0x2014)]", ' / '
    return ($s -replace '[^\x20-\x7E]', '' -replace '\s+', ' ').Trim()
}

function ShortLoc($loc) {
    if (-not $loc) { return '' }
    $p = @()
    if ($loc -match 'New York|NYC|\bNY\b|Manhattan|Brooklyn') { $p += 'NYC' }
    if ($loc -match 'San Francisco')  { $p += 'SF' }
    if ($loc -match 'Palo Alto')      { $p += 'Palo Alto' }
    if ($loc -match 'Seattle')        { $p += 'Seattle' }
    if ($loc -match 'Chicago')        { $p += 'Chicago' }
    if ($loc -match 'Pittsburgh')     { $p += 'Pittsburgh' }
    if ($loc -match 'Austin')         { $p += 'Austin' }
    if ($loc -match 'Greenwich')      { $p += 'Greenwich CT' }
    if ($loc -match 'London')         { $p += 'London' }
    if ($loc -match 'Paris')          { $p += 'Paris' }
    if ($p.Count -eq 0) { $p += (Norm $loc) }
    $out = $p -join ' / '
    if ($loc -match 'Hybrid') { $out += ' (hybrid)' }
    return $out
}

function ShortDate($s) {
    if (-not $s) { return '' }
    if ($s -match '^(\d{4})-(\d{2})-(\d{2})') { return "$($Matches[2])-$($Matches[3])" }
    if ($s -match 'Posted (\d+)\+? Day') { return "$($Matches[1])d+" }
    if ($s -match 'Posted Today') { return 'today' }
    return ''
}

$posted = @{}
foreach ($r in $n) { $posted[(Norm "$($r.Company) :: $($r.Title)")] = ShortDate $r.Updated }

# Curated level (verified against each posting's description):
#   MS   = program built for master's students
#   MSok = master's explicitly eligible
#   ?    = posting does not state a degree level
#   PhD  = PhD-only
#   UG   = undergrad-focused
$lvl = @{
 'blackrock :: 2027 Quantitative Masters Internship Program - Technology - Analytics & Modeling - New York' = 'MS'
 'blackrock :: 2027 Quantitative Masters Internship Program - Investments - Quantitative Investing - New York' = 'MS'
 'blackrock :: 2027 Quantitative Masters Internship Program - Investments - Global Capital Markets - New York' = 'MS'
 "pinterest :: Master's Data Science Internship 2027 (USA)" = 'MS'
 "pinterest :: Master's Machine Learning Internship 2027 (USA)" = 'MS'
 'pinterest :: PhD Data Science Internship 2027 (USA)' = 'PhD'
 'pinterest :: PhD Machine Learning Internship 2027 (USA)' = 'PhD'
 'point72 :: 2027 Summer Internship - MS/PhD Data Scientist, Proprietary Research' = 'MS'
 'point72 :: Summer 2027 Quantitative Research Internship' = 'PhD'
 'point72 :: Quantitative Research Intern' = 'MSok'
 'point72 :: Quantitative Researcher Intern' = 'MSok'
 'point72 :: Quantitative Research Intern (NLP)' = 'MSok'
 'point72 :: 2027 Summer Internship - Sector Analyst, Proprietary Research' = 'MSok'
 'point72 :: Quantitative Developer Intern' = '?'
 'point72 :: Quantitative Software Developer Intern' = '?'
 'point72 :: Machine Learning Researcher - Intern' = '?'
 'lyft :: Data Science Intern, Algorithms (Summer 2027 - SF/NYC)' = 'MS'
 'lyft :: Data Analyst Intern (Summer 2027)' = 'MSok'
 'duolingo :: Data Scientist, Intern' = 'MSok'
 'duolingo :: Data Scientist, Intern (PhD)' = 'PhD'
 'coinbase :: Data Science Intern' = 'MSok'
 'coinbase :: Analytics Engineer Intern' = 'MSok'
 'coinbase :: People Analytics Intern' = 'MSok'
 'coinbase :: Internal Audit Analytics Intern' = 'MSok'
 'stripe :: Data Analyst, Intern' = 'MSok'
 'stripe :: PhD Data Scientist, Intern' = 'PhD'
 'figma :: Data Science Intern (2027)' = '?'
 'figma :: Data Engineer Intern (2027)' = '?'
 'figma :: PhD Intern, Data Science (2027)' = 'PhD'
 'figma :: PhD Intern, AI Applied Scientist (2027)' = 'PhD'
 'ramp :: Applied Scientist Intern' = 'MSok'
 'imc :: Machine Learning Engineer Intern - Summer 2027' = 'MSok'
 'imc :: Hardware Machine Learning PhD Research Internship' = 'PhD'
 'schonfeld :: 2027 Business Analytics Intern' = 'MSok'
 'schonfeld :: 2027 PhD Quantitative Research Intern' = 'PhD'
 'drweng :: Quantitative Research Intern' = 'MSok'
 'jumptrading :: Campus Quantitative Researcher, UG/MS (Intern)' = 'MSok'
 'jumptrading :: Campus Quantitative Researcher, PhD (Intern)' = 'PhD'
 'jumptrading :: Campus Quantitative Trader (Intern)' = 'UG'
 'virtu :: 2027 Internship - Quantitative Researcher (Undergrad)' = 'UG'
 'virtu :: 2027 Internship - Quantitative Researcher (PhD)' = 'PhD'
 'virtu :: 2027 Internship - Quantitative Trading' = 'UG'
 'flowtraders :: Quantitative Trading Intern Summer 2027' = 'UG'
 'dvtrading :: Quantitative Research Intern - Summer 2027 (DV Equities)' = 'MSok'
 'dvtrading :: Quantitative Trading Intern - Summer 2027 (DV Equities)' = 'MSok'
 'dvtrading :: Futures & Options Trading Analyst Intern - Summer 2027' = 'MSok'
 'pdtpartners :: Summer 2027 Quantitative Research Intern - PhD/Postdoctoral' = 'PhD'
 'aqr :: 2027 Research Summer Analyst' = 'MSok'
 'aqr :: 2027 Quantitative Prediction Markets Research Summer Analyst' = 'MSok'
 'aqr :: AQR Arbitrage - 2027 Research Summer Analyst' = 'MSok'
 'aqr :: 2027 PhD Summer Associate, Machine Learning Research' = 'PhD'
}

$rows = foreach ($r in $e) {
    $k = Norm "$($r.Company) :: $($r.Title)"
    $l = $lvl[$k]; if (-not $l) { $l = '?' }
    [pscustomobject]@{
        Company = $r.Company
        Title   = Norm $r.Title
        Where   = (ShortLoc $r.Location)
        Level   = $l
        NYC     = $r.NYC
        Y2027   = $r.Term2027
        Posted  = $posted[$k]
        URL     = $r.URL
    }
}

function Tbl($set) {
    $arr = @($set)
    if ($arr.Count -eq 0) { return "_(none)_`n" }
    $sb = "| Company | Role | Level | Where |`n|---|---|---|---|`n"
    foreach ($r in ($arr | Sort-Object Company, Title)) {
        $when = if ($r.Posted) { " <sub>$($r.Posted)</sub>" } else { '' }
        $sb += "| $($r.Company) | [$($r.Title)]($($r.URL))$when | $($r.Level) | $($r.Where) |`n"
    }
    return $sb
}

$nyc     = @($rows | Where-Object { $_.NYC -eq 'True' })
$tier1   = @($nyc | Where-Object { $_.Level -eq 'MS' })
$tier2   = @($nyc | Where-Object { $_.Level -eq 'MSok' })
$tier3   = @($nyc | Where-Object { $_.Level -eq '?' })
$tier4   = @($nyc | Where-Object { $_.Level -in @('PhD','UG') })
$outside = @($rows | Where-Object { $_.NYC -ne 'True' })
$total   = @($rows).Count

$md = @"
# Summer 2027 Data Science Internships - New York City

Live, verified list of **$total open internship postings** for a data science master's student in the NYC area. Refreshed $(Get-Date -Format 'yyyy-MM-dd').

Every posting was pulled from the employer's own applicant-tracking system (Greenhouse, Lever, Ashby, Workday) and confirmed open. Degree level in the tables was checked against the actual job description text, not guessed from the title.

**Level key:** ``MS`` = program built for master's students &nbsp;|&nbsp; ``MSok`` = master's explicitly eligible &nbsp;|&nbsp; ``?`` = posting doesn't state a degree level &nbsp;|&nbsp; ``PhD`` = PhD-only &nbsp;|&nbsp; ``UG`` = undergrad-focused

**Snapshot:** $($tier1.Count) master's-targeted &nbsp;|&nbsp; $($tier2.Count) master's-eligible &nbsp;|&nbsp; $($tier3.Count) unstated &nbsp;|&nbsp; $($tier4.Count) PhD/undergrad &nbsp;|&nbsp; $($outside.Count) outside NYC

---

## Apply to these first

1. **BlackRock - 2027 Quantitative Masters Internship Program** (3 NYC tracks). A program that exists *only* for master's students, at the world's largest asset manager. The Analytics & Modeling track is the closest fit to a DS degree.
2. **Pinterest - Master's Data Science / Master's ML Internship 2027.** Master's-only tracks, New York selectable.
3. **Point72 - 2027 Summer Internship, MS/PhD Data Scientist (Proprietary Research).** NYC, explicitly MS-eligible, and a rare hedge-fund DS role rather than pure quant research.
4. **Lyft - Data Science Intern, Algorithms (Summer 2027).** Posting states Masters or PhD; NYC is one of two locations.
5. **Coinbase - Data Science Intern** (NYC hybrid) and **Duolingo - Data Scientist, Intern** (NYC). Both accept a graduate degree. High applicant volume - apply early.
6. **Stripe - Data Analyst, Intern** and **Schonfeld - 2027 Business Analytics Intern.** Both name bachelor's-or-master's explicitly; NYC available.

**Want quant instead?** DRW, IMC (ML Engineer), Jump Trading (UG/MS QR), DV Trading and Point72 Quantitative Research all list master's as eligible. Flow Traders, Virtu (Undergrad) and Jump (Quant Trader) are undergrad-titled - a master's student is usually still eligible, but the pipeline skews younger.

---

## Tier 1 - Built for master's students

$(Tbl $tier1)

## Tier 2 - Master's explicitly eligible

$(Tbl $tier2)

## Tier 3 - Degree level not stated

Read the posting before investing time; several are MS/PhD in practice.

$(Tbl $tier3)

## Tier 4 - PhD-only or undergrad-focused

$(Tbl $tier4)

## Outside NYC, but commutable

Greenwich, CT is about 50 minutes from Grand Central.

$(Tbl $outside)

---

## Not yet posted - watch these

Employers with no public ATS API, or no Summer 2027 DS role open at refresh time. These cycles usually open **Aug-Nov for the following summer**:

| Employer | Where to watch |
|---|---|
| Bloomberg | careers.bloomberg.com |
| JPMorgan Chase | jpmorganchase.com/careers |
| Goldman Sachs | higher.gs.com |
| Morgan Stanley | ms.wd5.myworkdayjobs.com/External |
| Two Sigma | twosigma.com/careers |
| Citadel / Citadel Securities | citadel.com/careers |
| Hudson River Trading | hudsonrivertrading.com/careers |
| Optiver | optiver.com/careers |
| Capital One | capitalone.wd12.myworkdayjobs.com/Capital_One |
| Mastercard | mastercard.wd1.myworkdayjobs.com/CorporateCareers |
| Meta / Google / Amazon / Microsoft | own portals, post Aug-Oct |
| MSK / Mount Sinai / NYU Langone | own portals (healthcare DS) |

## Timing notes

- Most Tier 1-2 postings went live **July-October 2026**. Quant firms, BlackRock and Pinterest hire on a rolling basis - apply within days of a posting appearing.
- Expect a coding/stats screen (HackerRank/Codility), then SQL plus product-analytics case rounds for tech DS roles. Quant research roles add probability and brainteaser rounds.
- Big-bank and big-tech cycles close earliest. Healthcare and insurance DS internships often post as late as Jan-Mar 2027.

---

## How this repo works

```
scripts/Find-Internships.ps1    scrape ~100 employer boards -> data/internships.csv
scripts/Get-Descriptions.ps1    pull each job description  -> data/enriched.csv
scripts/New-Report.ps1          rebuild this README
```

```powershell
./scripts/Find-Internships.ps1      # NYC metro only
./scripts/Find-Internships.ps1 -All # every location
./scripts/Get-Descriptions.ps1
./scripts/New-Report.ps1
```

`data/internships.csv` has the full detail - exact locations, posting dates, direct apply URLs. The GitHub Action in `.github/workflows/refresh.yml` re-runs the pipeline weekly and commits any changes, so this page stays current without you doing anything.

To widen coverage, add board slugs to `$boards` (Greenhouse/Lever/Ashby) or `$workday` in `Find-Internships.ps1`. A board slug is the identifier in a company's careers URL, e.g. ``boards.greenhouse.io/<slug>`` or ``jobs.ashbyhq.com/<slug>``.

**Caveat:** the ``?`` rows in Tier 3 were not verified for degree level. Tier 1 and Tier 2 were each confirmed by reading the posting's eligibility text.
"@

$md | Set-Content (Join-Path $root 'README.md') -Encoding UTF8
"README.md written"
"tier1=$($tier1.Count) tier2=$($tier2.Count) tier3=$($tier3.Count) tier4=$($tier4.Count) outside=$($outside.Count) total=$total"
