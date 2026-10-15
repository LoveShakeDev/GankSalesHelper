<#
.SYNOPSIS
    Gets your Gank sales and writes them to a CSV file.

.DESCRIPTION
    Downloads your Gank listings and earning transactions, then writes per-listing
    sales to GankSales-YYYY-MM-DD-HH-MM-SS.csv next to this script. Income totals
    are printed to the console.

    Downloaded data is cached in *.cache files next to this script, so later runs
    only need to fetch what's new. Delete them to force a full refresh.

    Works with Windows PowerShell 5.1 and PowerShell 7+.

.NOTES
    Getting a bearer token: log in at ganknow.com, open developer tools (F12), go to
    the Network tab, click around the site, pick a request to api.ganknow.com, and
    copy its "Authorization" request header.

    If Windows refuses to run the script, either right-click it > Properties > Unblock,
    or run: powershell -ExecutionPolicy Bypass -File .\GankSales.ps1
#>

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue' # Progress bars make web requests very slow in 5.1

# 5.1 bug: ConvertTo-Json can serialize arrays as {"value":[...],"Count":n} without this.
if ($PSVersionTable.PSEdition -ne 'Core') { Remove-TypeData System.Array -ErrorAction SilentlyContinue }

#region Constants

$MeUrl = 'https://api.ganknow.com/v1/users/me'
# {0} = page, {1} = page size. The listings URL also has a user ID inserted at runtime.
$ListingsUrlFormat = 'https://api.ganknow.com/v1/catalogs/services?userId=USER_ID&page={0}&per_page={1}&order_by=sort%20asc&order_by=createdAt%20desc'
$TransactionsUrlFormat = 'https://api-v2.ganknow.com/payment/transactions?page={0}&per_page={1}&usecases_client=WALLET_GANK_EARNING_TX'

# Transaction types
$TxProfileTip = 'DONATION'
$TxPostTip = 'POST_TIPPING'
$TxWithdrawal = 'WITHDRAWAL'
$TxDigitalGoods = 'DIGITAL_GOODS'

# Script directory ($PSScriptRoot is empty if the code was pasted into a console)
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).ProviderPath }
$SelfCachePath = Join-Path $ScriptDir 'Self.cache'
$ListingsCachePath = Join-Path $ScriptDir 'Listings.cache'
$TransactionsCachePath = Join-Path $ScriptDir 'Transactions.cache'

#endregion

#region Helpers

function Read-BearerToken {
    $raw = Read-Host 'Paste your Gank bearer token'
    # Accept "eyJ...", "Bearer eyJ...", or "Authorization: Bearer eyJ...", optionally quoted.
    $token = $raw.Trim().Trim('"', "'").Trim() -replace '^(Authorization:\s*)?Bearer\s+', ''
    if (-not $token) { throw 'No bearer token was entered.' }
    return $token
}

# GETs a Gank API URL and returns the parsed JSON. Decodes as UTF-8 explicitly so
# 5.1 doesn't mangle non-ASCII listing names.
function Invoke-GankApi([string]$Url) {
    $response = Invoke-WebRequest -Uri $Url -Headers $script:Headers -UseBasicParsing
    $json = [Text.Encoding]::UTF8.GetString($response.RawContentStream.ToArray())
    $parsed = ConvertFrom-Json -InputObject $json
    if ($null -eq $parsed -or $null -eq $parsed.data) { throw "Unexpected server response from $Url" }
    return $parsed
}

# Yields items from a paged endpoint, newest first, until it hits a known ID or runs out.
function Get-NewItems([string]$UrlFormat, [int]$PageSize, $KnownIds, [string]$IdProperty) {
    $page = 0
    while ($true) {
        $page++
        $items = @((Invoke-GankApi ($UrlFormat -f $page, $PageSize)).data)
        foreach ($item in $items) {
            if ($KnownIds.Contains([string]$item.$IdProperty)) { return }
            $item
        }
        if ($items.Count -lt $PageSize) { return }
    }
}

# Yields the elements of a cached JSON array, or nothing if the file doesn't exist.
# (5.1 and 7 differ in how ConvertFrom-Json returns arrays; foreach smooths that over.)
function Read-JsonArrayFile([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return }
    $parsed = ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($Path))
    foreach ($element in $parsed) { $element }
}

function Write-JsonFile([string]$Path, $Data) {
    $json = ConvertTo-Json -InputObject $Data -Depth 10 -Compress
    [IO.File]::WriteAllText($Path, $json, (New-Object Text.UTF8Encoding($false)))
}

# PowerShell 7 turns ISO date strings into DateTimes when parsing JSON; 5.1 doesn't.
function ConvertTo-DateTimeOffset($Value) {
    if ($Value -is [DateTimeOffset]) { return $Value }
    if ($Value -is [DateTime]) { return [DateTimeOffset]$Value }
    return [DateTimeOffset]::Parse([string]$Value, [Globalization.CultureInfo]::InvariantCulture)
}

#endregion

#region Data retrieval
# Cache files use the same shape as the GankHelper .NET app, so they're interchangeable.

function Get-UserId {
    if (Test-Path -LiteralPath $SelfCachePath) {
        $self = ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($SelfCachePath))
        if ($self.id) { return [string]$self.id }
    }

    $userId = [string](Invoke-GankApi $MeUrl).data.id
    Write-JsonFile $SelfCachePath ([pscustomobject]@{ id = $userId })
    return $userId
}

# Trims an API/cached listing down to the fields we cache.
function ConvertTo-CachedListing($Listing) {
    $prices = foreach ($price in $Listing.servicePrices) { [pscustomobject]@{ rate_usd = [decimal]$price.rate_usd } }
    [pscustomobject]@{
        id            = [string]$Listing.id
        name          = [string]$Listing.name
        createdAt     = (ConvertTo-DateTimeOffset $Listing.createdAt).ToUniversalTime().ToString('o')
        servicePrices = @($prices)
    }
}

# Trims an API/cached transaction down to the fields we cache.
function ConvertTo-CachedTransaction($Transaction) {
    $data = $Transaction.usecase_data.data
    [pscustomobject]@{
        human_id     = [string]$Transaction.human_id
        created_at   = (ConvertTo-DateTimeOffset $Transaction.created_at).ToUniversalTime().ToString('o')
        usecase      = [string]$Transaction.usecase
        price        = [decimal]$Transaction.price
        usecase_data = [pscustomobject]@{
            data = [pscustomobject]@{
                catalog_id            = $data.catalog_id
                catalog_title         = $data.catalog_title
                buyer_user_id         = $data.buyer_user_id
                buyer_user_nickname   = $data.buyer_user_nickname
                buyer_user_email      = $data.buyer_user_email
                donator_user_id       = $data.donator_user_id
                donator_user_nickname = $data.donator_user_nickname
                donator_user_email    = $data.donator_user_email
            }
        }
    }
}

# Loads cached items, appends anything new from the API, and saves the cache.
function Update-Cache([string]$CachePath, [string]$UrlFormat, [int]$FullPageSize, [int]$CatchUpPageSize, [string]$IdProperty, [scriptblock]$Convert) {
    $items = New-Object 'System.Collections.Generic.List[object]'
    $knownIds = New-Object 'System.Collections.Generic.HashSet[string]'

    foreach ($cached in Read-JsonArrayFile $CachePath) {
        $item = & $Convert $cached
        $items.Add($item)
        [void]$knownIds.Add($item.$IdProperty)
    }

    # A small page size is enough to catch up when there's already a cache.
    $pageSize = if ($items.Count -gt 0) { $CatchUpPageSize } else { $FullPageSize }
    foreach ($new in Get-NewItems $UrlFormat $pageSize $knownIds $IdProperty) {
        $items.Add((& $Convert $new))
    }

    Write-JsonFile $CachePath $items.ToArray()
    return , $items.ToArray()
}

#endregion

#region Report

function Write-SalesReport($Listings, $Transactions) {
    $tipEarnings = [decimal]0
    $itemEarnings = [decimal]0
    $withdrawn = [decimal]0
    $salesByListing = @{} # listing ID -> @{ SaleCount; Total }
    $skippedTypes = @{}   # unhandled transaction type -> count
    $seenIds = New-Object 'System.Collections.Generic.HashSet[string]'

    foreach ($tx in $Transactions) {
        if (-not $seenIds.Add($tx.human_id)) { continue } # Duplicate

        $amount = [decimal]$tx.price
        switch ($tx.usecase) {
            { $_ -eq $TxProfileTip -or $_ -eq $TxPostTip } { $tipEarnings += $amount; break }
            $TxWithdrawal { $withdrawn += $amount; break }
            $TxDigitalGoods {
                $itemEarnings += $amount
                $listingId = [string]$tx.usecase_data.data.catalog_id
                if (-not $salesByListing.ContainsKey($listingId)) { $salesByListing[$listingId] = @{ SaleCount = 0; Total = [decimal]0 } }
                $salesByListing[$listingId].SaleCount++
                $salesByListing[$listingId].Total += $amount
                break
            }
            default { $skippedTypes[$_] = 1 + [int]$skippedTypes[$_] }
        }
    }

    foreach ($type in $skippedTypes.Keys) {
        Write-Warning "Skipped $($skippedTypes[$type]) transaction(s) of unhandled type '$type'."
    }

    # One row per current listing, newest first.
    $rows = foreach ($listing in $Listings) {
        $sales = $salesByListing[$listing.id]
        $createdAt = ConvertTo-DateTimeOffset $listing.createdAt
        [pscustomobject]@{
            Name          = $listing.name
            SaleCount     = if ($sales) { $sales.SaleCount } else { 0 }
            TotalEarnings = if ($sales) { $sales.Total } else { [decimal]0 }
            CurrentPrice  = if (@($listing.servicePrices).Count -gt 0) { [decimal]@($listing.servicePrices)[0].rate_usd } else { [decimal]0 }
            CreatedDate   = $createdAt.LocalDateTime.ToString('yyyy-MM-dd HH:mm:ss')
            SortKey       = $createdAt
        }
    }
    $rows = @($rows | Sort-Object @{ Expression = 'SortKey'; Descending = $true }, Name | Select-Object -Property * -ExcludeProperty SortKey)

    # Write with a UTF-8 BOM so Excel shows non-ASCII names correctly (same result in 5.1 and 7).
    $csvPath = Join-Path $ScriptDir ('GankSales-{0}.csv' -f (Get-Date -Format 'yyyy-MM-dd-HH-mm-ss'))
    $csvLines = @($rows | ConvertTo-Csv -NoTypeInformation)
    [IO.File]::WriteAllLines($csvPath, [string[]]$csvLines, (New-Object Text.UTF8Encoding($true)))

    Write-Host ''
    Write-Host ('Earnings from items: {0:N2}' -f $itemEarnings)
    Write-Host ('Earnings from tips:  {0:N2}' -f $tipEarnings)
    Write-Host ('Amount withdrawn:    {0:N2}' -f $withdrawn)
    Write-Host ''
    Write-Host "Sales written to $csvPath"
}

#endregion

#region Main

try {
    # Windows PowerShell 5.1 may not enable TLS 1.2 by default.
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

    $script:Headers = @{ Authorization = 'Bearer ' + (Read-BearerToken) }

    Write-Host 'Getting listings...'
    $listingsUrlFormat = $ListingsUrlFormat.Replace('USER_ID', [Uri]::EscapeDataString((Get-UserId)))
    $listings = Update-Cache $ListingsCachePath $listingsUrlFormat 50 5 'id' ${function:ConvertTo-CachedListing}

    Write-Host 'Getting transactions...'
    $transactions = Update-Cache $TransactionsCachePath $TransactionsUrlFormat 500 20 'human_id' ${function:ConvertTo-CachedTransaction}

    Write-SalesReport $listings $transactions
}
catch {
    $status = $null
    if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
    if ($status -eq 401 -or $status -eq 403) {
        Write-Host 'Gank rejected the bearer token. It may be expired or mistyped; grab a fresh one and try again.' -ForegroundColor Red
    }
    else {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}
finally {
    # Keep the window open when launched via "Run with PowerShell".
    Read-Host 'Press Enter to exit' | Out-Null
}

#endregion
