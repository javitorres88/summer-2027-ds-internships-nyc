# Find-Internships.ps1
#
# Queries public job-board APIs (Greenhouse / Lever / Ashby / Workday) for internship
# postings matching data-science / ML / analytics keywords in the NYC metro area.
#
# Usage:
#   ./scripts/Find-Internships.ps1            # NYC metro only
#   ./scripts/Find-Internships.ps1 -All       # every location
#   ./scripts/Find-Internships.ps1 -OutCsv path.csv
param(
    [switch]$All,
    [string]$OutCsv
)
$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

$root = Split-Path -Parent $PSScriptRoot
$dataDir = Join-Path $root 'data'
if (-not (Test-Path $dataDir)) { New-Item -ItemType Directory -Path $dataDir | Out-Null }
if (-not $OutCsv) { $OutCsv = Join-Path $dataDir 'internships.csv' }

# PowerShell 5.1 decodes JSON responses as ISO-8859-1 when the Content-Type omits a
# charset, which turns UTF-8 into mojibake ("&#8226;" -> "a-euro-tm"). Repair it.
function Get-Latin1 {
    try { if ([System.Text.Encoding]::Latin1) { return [System.Text.Encoding]::Latin1 } } catch { }
    return [System.Text.Encoding]::GetEncoding(28591)
}
function FixEnc($s) {
    if (-not $s) { return $s }
    if ($s -notmatch '[^\x00-\x7F]') { return $s }
    try {
        $b = (Get-Latin1).GetBytes($s)
        $u = [System.Text.Encoding]::UTF8.GetString($b)
        if ($u -notmatch [char]0xFFFD) { return $u }
    } catch { }
    return $s
}

# ---------------------------------------------------------------- company boards
# s = ATS board token (from the careers URL), p = gh | lever | ashby
$boards = @(
    # Big tech / AI labs
    @{s='openai';p='ashby'}, @{s='anthropic';p='gh'}, @{s='databricks';p='gh'},
    @{s='snowflake';p='ashby'}, @{s='datadog';p='gh'}, @{s='mongodb';p='gh'},
    @{s='elastic';p='gh'}, @{s='cohere';p='ashby'}, @{s='perplexity';p='ashby'},
    @{s='sierra';p='ashby'}, @{s='harvey';p='ashby'}, @{s='modal';p='ashby'},
    @{s='anyscale';p='ashby'}, @{s='writer';p='ashby'}, @{s='twilio';p='gh'},
    @{s='labelbox';p='gh'},
    # Fintech / payments
    @{s='stripe';p='gh'}, @{s='robinhood';p='gh'}, @{s='coinbase';p='gh'},
    @{s='affirm';p='gh'}, @{s='chime';p='gh'}, @{s='brex';p='gh'},
    @{s='plaid';p='ashby'}, @{s='ramp';p='ashby'}, @{s='mercury';p='ashby'},
    @{s='alloy';p='gh'}, @{s='betterment';p='gh'}, @{s='gusto';p='gh'},
    @{s='justworks';p='gh'},
    # Quant / trading / asset management (NYC-heavy)
    @{s='janestreet';p='gh'}, @{s='imc';p='gh'}, @{s='akunacapital';p='gh'},
    @{s='pdtpartners';p='gh'}, @{s='point72';p='gh'}, @{s='aqr';p='gh'},
    @{s='drweng';p='gh'}, @{s='squarepointcapital';p='gh'}, @{s='schonfeld';p='gh'},
    @{s='virtu';p='gh'}, @{s='flowtraders';p='gh'}, @{s='dvtrading';p='gh'},
    @{s='jumptrading';p='gh'},
    # Data / analytics / infrastructure
    @{s='fivetran';p='gh'}, @{s='hex';p='ashby'}, @{s='mixpanel';p='gh'},
    @{s='amplitude';p='ashby'}, @{s='fullstory';p='ashby'}, @{s='asana';p='gh'},
    @{s='notion';p='ashby'}, @{s='figma';p='gh'}, @{s='checkr';p='gh'},
    @{s='persona';p='ashby'}, @{s='unit';p='ashby'}, @{s='substack';p='ashby'},
    # Consumer / marketplace / health
    @{s='airbnb';p='gh'}, @{s='instacart';p='gh'}, @{s='reddit';p='gh'},
    @{s='pinterest';p='gh'}, @{s='lyft';p='gh'}, @{s='dropbox';p='gh'},
    @{s='duolingo';p='gh'}, @{s='oscar';p='gh'}, @{s='zocdoc';p='gh'},
    @{s='flatironhealth';p='gh'}, @{s='benchling';p='ashby'},
    @{s='lattice';p='gh'}, @{s='greenhouse';p='gh'}, @{s='ashby';p='ashby'},
    # Logistics / hardware / mobility
    @{s='flexport';p='gh'}, @{s='project44';p='gh'}, @{s='fourkites';p='gh'},
    @{s='samsara';p='gh'}, @{s='motional';p='gh'}, @{s='waymo';p='gh'},
    @{s='nuro';p='gh'},
    # Consumer / media / martech
    @{s='peloton';p='gh'}, @{s='glossier';p='gh'}, @{s='zoominfo';p='gh'},
    @{s='similarweb';p='gh'}, @{s='taboola';p='gh'}, @{s='pubmatic';p='gh'},
    @{s='klaviyo';p='gh'}, @{s='braze';p='gh'}, @{s='iterable';p='ashby'},
    @{s='airship';p='gh'}, @{s='yotpo';p='gh'}, @{s='gorgias';p='ashby'},
    @{s='hellofresh';p='gh'}, @{s='sweetgreen';p='gh'}, @{s='away';p='ashby'},
    @{s='recharge';p='ashby'}, @{s='flexe';p='gh'}
)

# ---------------------------------------------------------------- Workday boards
# Large NYC employers that don't expose GH/Lever/Ashby boards.
$workday = @(
    @{t='ms';         w='wd5';  s='External'},
    @{t='blackrock';  w='wd1';  s='BlackRock_Professional'},
    @{t='mastercard'; w='wd1';  s='CorporateCareers'},
    @{t='capitalone'; w='wd12'; s='Capital_One'},
    @{t='pfizer';     w='wd1';  s='PfizerCareers'}
)

# ---------------------------------------------------------------- filters
$roleRx  = '\bintern(ship)?s?\b|co-?op\b|summer analyst|summer associate|placement student'
$fieldRx = 'data scien|machine learning|\bml\b|\bai\b|artificial intelligence|analytics|analyst|statistic|quantitative|research scientist|data engineer|applied scientist|decision scien|econometric|data analytics'
$locRx   = 'new york|nyc|\bny\b|manhattan|brooklyn|jersey city|hoboken|stamford|greenwich|paramus|newark'

$rows = @()
foreach ($b in $boards) {
    $jobs = @()
    try {
        switch ($b.p) {
            'gh'    { $jobs = (Invoke-RestMethod "https://boards-api.greenhouse.io/v1/boards/$($b.s)/jobs" -TimeoutSec 20).jobs |
                        ForEach-Object { [pscustomobject]@{ title=(FixEnc $_.title); loc=(FixEnc $_.location.name); url=$_.absolute_url; updated=$_.updated_at } } }
            'lever' { $jobs = Invoke-RestMethod "https://api.lever.co/v0/postings/$($b.s)?mode=json" -TimeoutSec 20 |
                        ForEach-Object { [pscustomobject]@{ title=(FixEnc $_.text); loc=(FixEnc $_.categories.location); url=$_.hostedUrl; updated=$_.createdAt } } }
            'ashby' { $jobs = (Invoke-RestMethod "https://api.ashbyhq.com/posting-api/job-board/$($b.s)" -TimeoutSec 20).jobs |
                        ForEach-Object { [pscustomobject]@{ title=(FixEnc $_.title); loc=(FixEnc ($_.location + ' ' + (($_.secondaryLocations.location) -join ' '))); url=$_.jobUrl; updated=$_.publishedAt } } }
        }
    } catch { continue }

    foreach ($j in $jobs) {
        if (-not $j.title) { continue }
        if ($j.title -notmatch $roleRx)  { continue }
        if ($j.title -notmatch $fieldRx) { continue }
        $isNYC = ($j.loc -match $locRx)
        if (-not $All -and -not $isNYC) { continue }
        $rows += [pscustomobject]@{
            Company  = $b.s
            Title    = $j.title
            Location = $j.loc
            NYC      = $isNYC
            Updated  = $j.updated
            URL      = $j.url
        }
    }
}

# ---- Workday sweep
foreach ($w in $workday) {
    $base = "https://$($w.t).$($w.w).myworkdayjobs.com"
    $api  = "$base/wday/cxs/$($w.t)/$($w.s)/jobs"
    for ($off = 0; $off -lt 120; $off += 20) {
        $body = @{ appliedFacets = @{}; limit = 20; offset = $off; searchText = 'intern' } | ConvertTo-Json -Compress
        try {
            $res = Invoke-RestMethod -Method Post -Uri $api -Body $body -ContentType 'application/json' `
                     -Headers @{ Accept = 'application/json' } -TimeoutSec 25
        } catch { break }
        if (-not $res.jobPostings) { break }
        foreach ($j in $res.jobPostings) {
            if ($j.title -notmatch $roleRx)  { continue }
            if ($j.title -notmatch $fieldRx) { continue }
            $isNYC = ($j.locationsText -match $locRx)
            if (-not $All -and -not $isNYC) { continue }
            $rows += [pscustomobject]@{
                Company  = $w.t
                Title    = FixEnc $j.title
                Location = FixEnc $j.locationsText
                NYC      = $isNYC
                Updated  = $j.postedOn
                URL      = "$base/en-US/$($w.s)$($j.externalPath)"
            }
        }
        if (($off + 20) -ge $res.total) { break }
    }
}

$rows = $rows | Sort-Object Company, Title -Unique
$rows | Export-Csv -Path $OutCsv -NoTypeInformation -Encoding UTF8
"MATCHES: $($rows.Count) -> $OutCsv"
$rows | Sort-Object @{Expression='NYC';Descending=$true}, Company | Format-Table Company, Title, Location -AutoSize -Wrap
