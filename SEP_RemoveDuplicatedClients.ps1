# =================================================================================
# Skrypt PowerShell do usuwania zduplikowanych klientów SEPM przez REST API
# Wersja: 1.0
#
# UWAGA: Uruchomienie tego skryptu może spowodować nieodwracalne zmiany w bazie
# danych SEPM. Zawsze wykonuj kopię zapasową i testuj w środowisku deweloperskim.
# USE FOR OWN RISK
# https://{{SEPM_server_address}}:8446/sepm/restapidocs.html
# =================================================================================

# ============== VARIABLES ==============
# Adres IP lub nazwa FQDN serwera SEPM
$sepmServer = "{{SEPM_server_address}}"
# Port API (domyślnie 8446)
$sepmPort = 8446
# Nazwa użytkownika administratora SEPM
$sepmUser = "$($env:USERNAME)"
# Users password (leave blank for security reasons, script will ask for pass)
$sepmPassword = ""

# Dry Run - Set to $true to show what will be deleted
# set to $false to REMOVE clients 
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

# --- KROK 1: UWIERZYTELNIENIE I POBRANIE TOKENA ---

if (-not $sepmPassword) {
    $sepmPassword = Read-Host -AsSecureString -Prompt "Logon as '$sepmUser'"
}

$baseApiUrl = "https://{0}:{1}/sepm/api/v1" -f $sepmServer, $sepmPort
$authUrl = "$baseApiUrl/identity/authenticate"
$headers = @{ "Content-Type" = "application/json" }

$authBody = @{
    "username" = $sepmUser
    "password" = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($sepmPassword))
    "domain"   = "" # Pozostaw puste, jeśli używasz konta lokalnego SEPM
} | ConvertTo-Json

Write-Host "Krok 1: Uwierzytelnianie na serwerze $sepmServer..."
try {
    $authResponse = Invoke-RestMethod -Uri $authUrl -Method Post -Headers $headers -Body $authBody
    $accessToken = $authResponse.token
    Write-Host "Uwierzytelnianie pomyślne. Uzyskano token." -ForegroundColor Green
}
catch {
    Write-Error "Błąd podczas uwierzytelniania: $($_.Exception.Message)"
    exit
}

# --- KROK 2: POBRANIE LISTY WSZYSTKICH KLIENTÓW ---

$computersUrl = "$baseApiUrl/computers"
$apiHeaders = @{
    "Content-Type"  = "application/json"
    "Authorization" = "Bearer $accessToken"
}

Write-Host "Krok 2: Pobieranie listy wszystkich klientów... (może to potrwać dłuższą chwilę)"
try {
    # API zwraca dane stronnicowo, pętla 'do-while' pobierze wszystkie strony
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
    
    Write-Host "`nPobrano łącznie $($allClients.Count) klientów." -ForegroundColor Green
}
catch {
    Write-Error "Błąd podczas pobierania listy klientów: $($_.Exception.Message)"
    exit
}


# --- KROK 3: IDENTYFIKACJA DUPLIKATÓW ---

Write-Host "Krok 3: Analiza i identyfikacja duplikatów na podstawie nazwy klienta (Computer Name)..."

# Grupujemy klientów po 'hardwareKey' i filtrujemy te grupy, które mają więcej niż jednego członka
# $duplicates = $allClients | Group-Object -Property hardwareKey | Where-Object { $_.Count -gt 1 }

# Grupujemy klientów po 'computerName' i filtrujemy te grupy, które mają więcej niż jednego członka
$duplicates = $allClients | Group-Object -Property computerName | Where-Object { $_.Count -gt 2 }

if ($duplicates.Count -eq 0) {
    Write-Host "Nie znaleziono żadnych duplikatów. Kończenie pracy." -ForegroundColor Green
    exit
}

Write-Host "Znaleziono $($duplicates.Count) grup zduplikowanych klientów." -ForegroundColor Yellow

# --- KROK 4 i 5: WYBÓR NAJNOWSZEGO I USUWANIE POZOSTAŁYCH ---

if ($dryRun) {
    Write-Host "[TRYB TESTOWY (Dry Run)] Skrypt tylko wyświetli, którzy klienci zostaliby usunięci." -ForegroundColor Cyan
} else {
    Write-Host "[TRYB RZECZYWISTY] Skrypt będzie usuwał zduplikowanych klientów!" -ForegroundColor Red -BackgroundColor White
}

$totalDeletedCount = 0

foreach ($group in $duplicates) {
    Write-Host "--------------------------------------------------------"
    Write-Host "Analizowanie duplikatów dla klucza sprzętowego: $($group.Name)"
    
    # Sortujemy klientów w grupie po dacie ostatniego zameldowania, od najnowszego do najstarszego
    $sortedClients = $group.Group | Sort-Object -Property agentTimeStamp -Descending

    # Pierwszy na liście jest najnowszy - tego zostawiamy
    $clientToKeep = $sortedClients | Select-Object -First 1
    
    # Pozostali to duplikaty do usunięcia
    $clientsToDelete = $sortedClients | Select-Object -Skip 1
    
    # Write-Host "  > Klient do pozostawienia: $($clientToKeep.computerName) (Ostatnie połączenie: $($clientToKeep.lastCheckinTime))" -ForegroundColor Green
    Write-Host "  > Klient do pozostawienia: $($clientToKeep.computerName), CompID: $($clientToKeep.uniqueId), HWkey: $($clientToKeep.hardwareKey), (Ostatnie połączenie: $((ConvertFrom-UnixTime $clientToKeep.agentTimeStamp).datetime))" -ForegroundColor Green

    foreach ($client in $clientsToDelete) {
        Write-Host "  > Klient do USUNIĘCIA: $($client.computerName), CompID: $($client.uniqueId), HWkey: $($client.hardwareKey), Ostatnie połączenie: $((ConvertFrom-UnixTime $client.agentTimeStamp).datetime))" -ForegroundColor Yellow
        if (-not $dryRun) {
            $deleteUrl = "$computersUrl/$($client.uniqueId)"
            try {
                Invoke-RestMethod -Uri $deleteUrl -Method Delete -Headers $apiHeaders
                Write-Host "    - SUKCES: Usunięto klienta $($client.computerName)." -ForegroundColor DarkGreen -BackgroundColor Green
                $totalDeletedCount++
                sleep -Seconds 2
            }
            catch {
                Write-Host "    - BŁĄD: Nie udało się usunąć klienta $($client.computerName). Komunikat: $($_.Exception.Message)" -ForegroundColor Red
                exit
            }
        }
    }
}

Write-Host "--------------------------------------------------------"
if ($dryRun) {
    Write-Host "Tryb testowy zakończony. Nie dokonano żadnych zmian." -ForegroundColor Cyan
} else {
    Write-Host "Proces usuwania zakończony. Usunięto łącznie $totalDeletedCount klientów." -ForegroundColor Green
}

# --- KROK 6: WYLOGOWANIE (UNIEWAŻNIENIE TOKENA) ---
$logoutUrl = "$baseApiUrl/identity/logout"
try {
    Invoke-RestMethod -Uri $logoutUrl -Method Post -Headers $apiHeaders
    Write-Host "Token dostępu został unieważniony."
}
catch {
    Write-Warning "Nie udało się unieważnić tokena dostępu: $($_.Exception.Message)"
}
