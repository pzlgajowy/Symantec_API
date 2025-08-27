# =================================================================================
# Author:      Pawel Lesniewski
# Created:     2025.08.27
# Description: PowerShell script for removing duplicate SEPM clients via REST API
# Version:     1.0
# 
# WARNING:     Running this script may cause irreversible changes to the SEPM database.
#              Always back up and test in a development environment.
# 
#                   USE AT YOUR OWN RISK
# 
# https://{{SEPM_server_address}}:8446/sepm/restapidocs.html
# =================================================================================

# ============== VARIABLES ==============
# Adres IP lub nazwa FQDN serwera SEPM
$sepmServer = "{{SEPM_server_address}}"
# Port API (domyślnie 8446)
$sepmPort = 8446
# SEPM administrators login
$sepmUser = "$($env:USERNAME)"
# Users password (leave blank for security reasons, script will ask for pass)
$sepmPassword = ""

# Dry Run - Set to $true to show what will be deleted
# Set to $false to PERMANENTLY REMOVE clients
[bool]$dryRun = $true
#     $dryRun = $false

# =====================================================

function ConvertFrom-UnixTime ([UInt64]$EpochTime){
    if ($EpochTime -gt [uint32]::MaxValue) {
        $result = ([System.DateTimeOffset]::FromUnixTimeMilliseconds($EpochTime))
    } else {
        $result = ([System.DateTimeOffset]::FromUnixTimeSeconds($EpochTime))
    }
    return $result 
}

# Ignore SSL certificate errors (usefull for self-signed certs)
add-type @"
    using System.Net;
    using System.Security.Cryptography.X509Certificates;
    public class TrustAllCertsPolicy : ICertificatePolicy {
        public bool CheckValidationResult(ServicePoint srvPoint, X509Certificate certificate, WebRequest request, int certificateProblem) {
            return true;
        }
    }
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

# --- STEP 1: AUTHENTICATION AND TOKEN DOWNLOAD ---

if (-not $sepmPassword) {
    $sepmPassword = Read-Host -AsSecureString -Prompt "Logon as '$sepmUser'"
}

$baseApiUrl = "https://{0}:{1}/sepm/api/v1" -f $sepmServer, $sepmPort
$authUrl = "$baseApiUrl/identity/authenticate"
$headers = @{ "Content-Type" = "application/json" }

$authBody = @{
    "username" = $sepmUser
    "password" = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($sepmPassword))
    "domain"   = "" # Leave blank if you are using a local SEPM account
} | ConvertTo-Json

Write-Host "Step 1: Authentication on the $sepmServer server..."
try {
    $authResponse = Invoke-RestMethod -Uri $authUrl -Method Post -Headers $headers -Body $authBody
    $accessToken = $authResponse.token
    Write-Host "Authentication successful. Token obtained." -ForegroundColor Green
}
catch {
    Write-Error "Error while authenticating: $($_.Exception.Message)"
    exit
}

# --- STEP 2: DOWNLOAD THE LIST OF ALL CUSTOMERS ---

$computersUrl = "$baseApiUrl/computers"
$apiHeaders = @{
    "Content-Type"  = "application/json"
    "Authorization" = "Bearer $accessToken"
}

Write-Host "Step 2: Downloading a list of all clients... (this may take a while)"
try {
    # The API returns data in pages, the 'do-while' loop will fetch all pages
    $allClients = @()
    $pageIndex = 1
    do {
        $pagedUrl = "$($computersUrl)?pageIndex=$($pageIndex)&pageSize=1000"
        $pageResponse = Invoke-RestMethod -Uri $pagedUrl -Method Get -Headers $apiHeaders
        if ($pageResponse.content) {
            $allClients += $pageResponse.content
        }
        $pageIndex++
        Write-Host -NoNewline "."
    } while ($pageResponse.totalElements -gt $allClients.Count)
    
    Write-Host "`nA total of $($allClients.Count) clients were downloaded.." -ForegroundColor Green
}
catch {
    Write-Error "Error while retrieving customer list: $($_.Exception.Message)"
    exit
}


# --- STEP 3: IDENTIFY DUPLICATES ---

Write-Host "Step 3: Analyze and identify duplicates based on the Client Name (Computer Name)..."

# We group clients by 'hardwareKey' and filter those groups that have more than one member
# $duplicates = $allClients | Group-Object -Property hardwareKey | Where-Object { $_.Count -gt 2 }

# We group clients by 'computerName' and filter those groups that have more than one member
$duplicates = $allClients | Group-Object -Property computerName | Where-Object { $_.Count -gt 2 }

if ($duplicates.Count -eq 0) {
    Write-Host "No duplicates found. Finishing work." -ForegroundColor Green
    exit
}

Write-Host "A $($duplicates.Count) groups of duplicate clients found." -ForegroundColor Yellow

# --- STEP 4 and 5: SELECT THE LATEST AND REMOVE THE REST ---

if ($dryRun) {
    Write-Host "[TEST MODE (Dry Run)] The script will only display which clients would be deleted." -ForegroundColor Cyan
} else {
    Write-Host "[REAL MODE] The script will remove duplicate clients!" -ForegroundColor Red -BackgroundColor White
}

$totalDeletedCount = 0

foreach ($group in $duplicates) {
    Write-Host "--------------------------------------------------------"
    Write-Host "Analyzing duplicates by computer name: $($group.Name)"
    
    # We sort clients in a group by their last check-in date, from newest to oldest
    $sortedClients = $group.Group | Sort-Object -Property agentTimeStamp -Descending

    # The first one on the list is the newest one - we'll leave that one alone
    $clientToKeep = $sortedClients | Select-Object -First 1
    
    # The rest are duplicates to be removed
    $clientsToDelete = $sortedClients | Select-Object -Skip 1
    
    # Write-Host "  > client to leave: $($clientToKeep.computerName) (lastConn: $($clientToKeep.lastCheckinTime))" -ForegroundColor Green
    Write-Host "  > client to leave: $($clientToKeep.computerName), CompID: $($clientToKeep.uniqueId), HWkey: $($clientToKeep.hardwareKey), (lastConn: $((ConvertFrom-UnixTime $clientToKeep.agentTimeStamp).datetime))" -ForegroundColor Green

    foreach ($client in $clientsToDelete) {
        Write-Host "  > client to REMOVE: $($client.computerName), CompID: $($client.uniqueId), HWkey: $($client.hardwareKey), lastConn: $((ConvertFrom-UnixTime $client.agentTimeStamp).datetime))" -ForegroundColor Yellow
        if (-not $dryRun) {
            $deleteUrl = "$computersUrl/$($client.uniqueId)"
            try {
                Invoke-RestMethod -Uri $deleteUrl -Method Delete -Headers $apiHeaders
                Write-Host "    - SUCCESS: Client removed $($client.computerName)." -ForegroundColor DarkGreen -BackgroundColor Green
                $totalDeletedCount++
                # The API will execute 50 requests and then reject the next ones with error 429. To prevent this, the script waits 1,3 seconds after each request. 
                sleep -Milliseconds 1300
            }
            catch {
                Write-Host "    - ERROR: Failed to delete client $($client.computerName). Message: $($_.Exception.Message)" -ForegroundColor Red
                exit
            }
        }
    }
}

Write-Host "--------------------------------------------------------"
if ($dryRun) {
    Write-Host "Test mode completed. No changes made." -ForegroundColor Cyan
} else {
    Write-Host "Deletion process completed. A total of $totalDeletedCount clients were deleted." -ForegroundColor Green
}

# --- STEP 6: LOGOUT (INVALIDATE TOKEN) ---
$logoutUrl = "$baseApiUrl/identity/logout"
try {
    Invoke-RestMethod -Uri $logoutUrl -Method Post -Headers $apiHeaders
    Write-Host "The access token has been invalidated."
}
catch {
    Write-Warning "Failed to revoke access token: $($_.Exception.Message)"
}
