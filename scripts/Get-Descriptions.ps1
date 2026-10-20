# Get-Descriptions.ps1
#
# Pulls the full job description for every posting in data/internships.csv, then
# classifies each one (Summer 2027? master's eligible? PhD-only?) and writes
# data/enriched.csv plus data/descriptions.json.
$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

$root    = Split-Path -Parent $PSScriptRoot
$dataDir = Join-Path $root 'data'
$inCsv   = Join-Path $dataDir 'internships.csv'
$outCsv  = Join-Path $dataDir 'enriched.csv'
$outJson = Join-Path $dataDir 'descriptions.json'

if (-not (Test-Path $inCsv)) { throw "Missing $inCsv - run Find-Internships.ps1 first." }

function Get-Latin1 {
    try { if ([System.Text.Encoding]::Latin1) { return [System.Text.Encoding]::Latin1 } } catch { }
    return [System.Text.Encoding]::GetEncoding(28591)
}

function Strip($h) {
    if (-not $h) { return '' }
    $t = $h -replace '<[^>]+>', ' '
    $t = [System.Net.WebUtility]::HtmlDecode($t)
    # undo PS 5.1's Latin-1 JSON decoding
    if ($t -match '[^\x00-\x7F]') {
        try {
            $b = (Get-Latin1).GetBytes($t)
            $u = [System.Text.Encoding]::UTF8.GetString($b)
            if ($u -notmatch [char]0xFFFD) { $t = $u }
        } catch { }
    }
    ($t -replace '\s+', ' ').Trim()
}

# platform per company slug
$plat = @{}
foreach ($l in @(
  'openai:ashby','anthropic:gh','databricks:gh','snowflake:ashby','datadog:gh','mongodb:gh',
  'elastic:gh','cohere:ashby','perplexity:ashby','sierra:ashby','harvey:ashby','modal:ashby',
  'anyscale:ashby','writer:ashby','twilio:gh','labelbox:gh','stripe:gh','robinhood:gh',
  'coinbase:gh','affirm:gh','chime:gh','brex:gh','plaid:ashby','ramp:ashby','mercury:ashby',
  'alloy:gh','betterment:gh','gusto:gh','justworks:gh','janestreet:gh','imc:gh','akunacapital:gh',
  'pdtpartners:gh','point72:gh','aqr:gh','fivetran:gh','hex:ashby','mixpanel:gh','amplitude:ashby',
  'fullstory:ashby','asana:gh','notion:ashby','figma:gh','checkr:gh','persona:ashby','unit:ashby',
  'substack:ashby','airbnb:gh','instacart:gh','reddit:gh','pinterest:gh','lyft:gh','dropbox:gh',
  'duolingo:gh','oscar:gh','zocdoc:gh','flatironhealth:gh','benchling:ashby','lattice:gh',
  'greenhouse:gh','ashby:ashby','flexport:gh','project44:gh','fourkites:gh','samsara:gh',
  'motional:gh','waymo:gh','nuro:gh','drweng:gh','squarepointcapital:gh','schonfeld:gh',
  'virtu:gh','flowtraders:gh','dvtrading:gh','jumptrading:gh','peloton:gh','glossier:gh',
  'zoominfo:gh','similarweb:gh','taboola:gh','pubmatic:gh','klaviyo:gh','braze:gh',
  'iterable:ashby','airship:gh','yotpo:gh','gorgias:ashby','hellofresh:gh','sweetgreen:gh',
  'away:ashby','recharge:ashby','flexe:gh')) {
    $p = $l -split ':'; $plat[$p[0]] = $p[1]
}
$wd = @{ ms=@{w='wd5';s='External'}; blackrock=@{w='wd1';s='BlackRock_Professional'};
         mastercard=@{w='wd1';s='CorporateCareers'}; capitalone=@{w='wd12';s='Capital_One'};
         pfizer=@{w='wd1';s='PfizerCareers'} }

$in   = Import-Csv $inCsv
$out  = @()
$full = @{}

foreach ($row in $in) {
    $slug = $row.Company
    $desc = ''
    try {
        if ($wd.ContainsKey($slug)) {
            $w    = $wd[$slug]
            $path = ($row.URL -replace '^.*?/en-US/[^/]+', '')
            $api  = "https://$slug.$($w.w).myworkdayjobs.com/wday/cxs/$slug/$($w.s)$path"
            $d    = Invoke-RestMethod -Uri $api -Headers @{ Accept = 'application/json' } -TimeoutSec 25
            $desc = Strip $d.jobPostingInfo.jobDescription
        }
        elseif ($plat[$slug] -eq 'gh') {
            if ($row.URL -match 'gh_jid=(\d+)') { $id = $Matches[1] }
            else { $id = ($row.URL -split '/')[-1] -replace '\?.*$', '' }
            $d    = Invoke-RestMethod "https://boards-api.greenhouse.io/v1/boards/$slug/jobs/$id" -TimeoutSec 25
            $desc = Strip $d.content
        }
        elseif ($plat[$slug] -eq 'lever') {
            $all = Invoke-RestMethod "https://api.lever.co/v0/postings/$slug?mode=json" -TimeoutSec 25
            $m   = $all | Where-Object { $_.hostedUrl -eq $row.URL } | Select-Object -First 1
            if ($m) { $desc = Strip $m.descriptionPlain }
        }
        elseif ($plat[$slug] -eq 'ashby') {
            $all = (Invoke-RestMethod "https://api.ashbyhq.com/posting-api/job-board/$slug" -TimeoutSec 25).jobs
            $m   = $all | Where-Object { $_.jobUrl -eq $row.URL } | Select-Object -First 1
            if ($m) { $desc = Strip $m.descriptionPlain }
        }
    } catch { $desc = '' }

    $t  = $row.Title
    $d2 = $desc
    $is2027  = ($t + ' ' + $d2) -match '2027'
    $phdOnly = ($t -match 'PhD|Doctoral') -or (($d2 -match 'PhD') -and ($d2 -notmatch 'Master'))
    $masters = ($t + ' ' + $d2) -match "Master|MS\b|M\.S\.|graduate student|Master's"

    $snip = ''
    $mm = [regex]::Match($d2, '(eligib|qualification|what you.ll need|what we.re looking for|requirements|who you are|basic qualifications|minimum qualifications).{0,420}', 'IgnoreCase')
    if ($mm.Success) { $snip = $mm.Value.Trim() }
    elseif ($d2.Length -gt 0) { $snip = $d2.Substring(0, [Math]::Min(300, $d2.Length)) }

    # Strict NYC test. The scraper's $locRx deliberately includes commuter towns
    # (Greenwich, Stamford) so those rows get collected; for the report we only want
    # postings that are actually in the city.
    $nycStrict = ($row.Location -match 'New York|NYC|\bNY\b|Manhattan|Brooklyn')

    $out += [pscustomobject]@{
        Company   = $slug
        Title     = $t
        Location  = $row.Location
        Term2027  = $is2027
        MastersOK = $masters
        PhDOnly   = $phdOnly
        NYC       = $nycStrict
        DescLen   = $d2.Length
        Snippet   = $snip
        URL       = $row.URL
    }
    $full["$slug :: $t"] = $d2
}

$out  | Export-Csv $outCsv -NoTypeInformation -Encoding UTF8
$full | ConvertTo-Json -Depth 3 | Set-Content $outJson -Encoding UTF8

"enriched rows: $($out.Count)  (descriptions fetched: $(($out | Where-Object DescLen -gt 0).Count))"
