# ============================================================================
# Check-Cameras.ps1
# ============================================================================
# Beskrivelse:
#   Læser kamera-liste fra Excel, tester hver IP-adresse (ping + HTTP),
#   og skriver resultater til CSV-fil som Excel kan importere.
#
# Constraints:
#   - Ingen direkte Excel COM-objekter (for at undgå afhængigheder)
#   - Outputter kun CSV som Excel manuelt kan refreshe
#   - Ingen automatisk opdatering
#
# ============================================================================

param(
    [string]$ExcelFile = "$PSScriptRoot\..\excel\CameraDashboard.xlsx",
    [string]$OutputFile = "$PSScriptRoot\..\output\CameraStatus.csv",
    [int]$TimeoutSeconds = 2,
    [int]$HttpPort = 80
)

# ============================================================================
# FUNKTIONER
# ============================================================================

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        "ERROR" { "Red" }
        "WARNING" { "Yellow" }
        "SUCCESS" { "Green" }
        default { "White" }
    }
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

function Test-CameraConnection {
    param(
        [string]$IpAddress,
        [int]$TimeoutMs = 2000,
        [int]$HttpPort = 80
    )
    
    $result = [PSCustomObject]@{
        IpAddress = $IpAddress
        PingStatus = "FAIL"
        HttpStatus = "FAIL"
        OverallStatus = "OFFLINE"
        ResponseTime = 0
        LastChecked = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
    
    # Test 1: ICMP Ping
    try {
        # PowerShell 5.1 kompatibel syntaks (bruger millisekunder)
        $ping = Test-Connection -ComputerName $IpAddress -Count 1 -Quiet -ErrorAction SilentlyContinue
        if ($ping) {
            $result.PingStatus = "OK"
            # Få response time med fuld ping
            $pingFull = Test-Connection -ComputerName $IpAddress -Count 1 -ErrorAction SilentlyContinue
            if ($pingFull) {
                $result.ResponseTime = $pingFull.ResponseTime
            }
        }
    } catch {
        Write-Log "Ping fejlede for $IpAddress : $_" -Level "WARNING"
    }
    
    # Test 2: HTTP port check (kun hvis ping OK)
    if ($result.PingStatus -eq "OK") {
        try {
            $tcpClient = New-Object System.Net.Sockets.TcpClient
            $connect = $tcpClient.BeginConnect($IpAddress, $HttpPort, $null, $null)
            $wait = $connect.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
            
            if ($wait) {
                try {
                    $tcpClient.EndConnect($connect)
                    $result.HttpStatus = "OK"
                    $result.OverallStatus = "ONLINE"
                } catch {
                    $result.HttpStatus = "FAIL"
                }
            }
            $tcpClient.Close()
        } catch {
            Write-Log "HTTP check fejlede for $IpAddress : $_" -Level "WARNING"
        }
    }
    
    return $result
}

function Read-CameraListFromExcel {
    param([string]$ExcelPath)
    
    Write-Log "Læser kamera-liste fra Excel..."
    
    # Forsøg at læse Excel via COM (hvis tilgængelig)
    $cameras = @()
    
    try {
        # Metode 1: Brug Excel COM object (kræver Excel installeret)
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        $workbook = $excel.Workbooks.Open($ExcelPath)
        $worksheet = $workbook.Sheets.Item("CameraList")
        
        # Find tabel-område (antag tabel starter ved A1)
        $lastRow = $worksheet.Cells.Find("*", [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value, 1, 2).Row
        
        # Læs data (spring header over)
        for ($row = 2; $row -le $lastRow; $row++) {
            $id = $worksheet.Cells.Item($row, 1).Text
            $name = $worksheet.Cells.Item($row, 2).Text
            $ip = $worksheet.Cells.Item($row, 3).Text
            $location = $worksheet.Cells.Item($row, 4).Text
            
            if ($ip -and $ip -ne "") {
                $cameras += [PSCustomObject]@{
                    ID = $id
                    CameraName = $name
                    IpAddress = $ip
                    Location = $location
                }
            }
        }
        
        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        
        Write-Log "Læste $($cameras.Count) kameraer fra Excel (COM)" -Level "SUCCESS"
        
    } catch {
        Write-Log "Kunne ikke læse Excel via COM: $_" -Level "WARNING"
        
        # Metode 2: Fallback - Læs fra example CSV hvis Excel ikke kan åbnes
        $exampleFile = "$PSScriptRoot\..\examples\CameraList_Example.csv"
        if (Test-Path $exampleFile) {
            Write-Log "Bruger example CSV som fallback..." -Level "WARNING"
            $cameras = Import-Csv -Path $exampleFile -Delimiter ";"
        } else {
            Write-Log "FEJL: Kunne ikke læse Excel og ingen example CSV fundet!" -Level "ERROR"
            throw "Kan ikke læse kamera-liste. Sørg for Excel er installeret eller brug example CSV."
        }
    }
    
    return $cameras
}

# ============================================================================
# HOVEDPROGRAM
# ============================================================================

Write-Log "========================================" -Level "INFO"
Write-Log "  KAMERA DASHBOARD - TEST STARTER" -Level "INFO"
Write-Log "========================================" -Level "INFO"
Write-Log "Excel fil: $ExcelFile"
Write-Log "Output fil: $OutputFile"
Write-Log "Timeout: $TimeoutSeconds sekunder"
Write-Log ""

# Valider input
if (-not (Test-Path $ExcelFile)) {
    Write-Log "FEJL: Excel-filen findes ikke: $ExcelFile" -Level "ERROR"
    Write-Log "Kør først BUILD_STEPS.md for at oprette Excel-filen!" -Level "ERROR"
    Read-Host "Tryk Enter for at afslutte"
    exit 1
}

# Opret output mappe hvis ikke eksisterer
$outputDir = Split-Path -Parent $OutputFile
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    Write-Log "Oprettede output mappe: $outputDir"
}

# Læs kamera-liste
try {
    $cameras = Read-CameraListFromExcel -ExcelPath $ExcelFile
    
    if ($cameras.Count -eq 0) {
        Write-Log "ADVARSEL: Ingen kameraer fundet i Excel!" -Level "WARNING"
        Write-Log "Tilføj kameraer til 'CameraList' sheet i Excel." -Level "WARNING"
        Read-Host "Tryk Enter for at afslutte"
        exit 0
    }
    
} catch {
    Write-Log "FEJL ved læsning af Excel: $_" -Level "ERROR"
    Read-Host "Tryk Enter for at afslutte"
    exit 1
}

# Test alle kameraer
Write-Log ""
Write-Log "Tester $($cameras.Count) kameraer..." -Level "INFO"
Write-Log "----------------------------------------"

$results = @()
$onlineCount = 0

foreach ($camera in $cameras) {
    Write-Host "Testing $($camera.CameraName) ($($camera.IpAddress))..." -NoNewline
    
    $testResult = Test-CameraConnection -IpAddress $camera.IpAddress -TimeoutMs ($TimeoutSeconds * 1000) -HttpPort $HttpPort
    
    # Kombiner kamera info med test resultat
    $combinedResult = [PSCustomObject]@{
        ID = $camera.ID
        CameraName = $camera.CameraName
        IpAddress = $camera.IpAddress
        Location = $camera.Location
        Status = $testResult.OverallStatus
        PingStatus = $testResult.PingStatus
        HttpStatus = $testResult.HttpStatus
        ResponseTime = $testResult.ResponseTime
        LastChecked = $testResult.LastChecked
    }
    
    $results += $combinedResult
    
    if ($testResult.OverallStatus -eq "ONLINE") {
        $onlineCount++
        Write-Host " [OK] ONLINE ($($testResult.ResponseTime)ms)" -ForegroundColor Green
    } else {
        Write-Host " [FAIL] OFFLINE" -ForegroundColor Red
    }
}

# Gem resultater til CSV
Write-Log ""
Write-Log "Gemmer resultater til: $OutputFile"

try {
    $results | Export-Csv -Path $OutputFile -Delimiter ";" -Encoding UTF8 -NoTypeInformation
    Write-Log "Resultater gemt succesfuldt!" -Level "SUCCESS"
} catch {
    Write-Log "FEJL ved gemning af resultater: $_" -Level "ERROR"
    Read-Host "Tryk Enter for at afslutte"
    exit 1
}

# Sammenfatning
Write-Log ""
Write-Log "========================================" -Level "SUCCESS"
Write-Log "  TEST AFSLUTTET" -Level "SUCCESS"
Write-Log "========================================" -Level "SUCCESS"
Write-Log "Total kameraer: $($cameras.Count)"
Write-Log "Online: $onlineCount" -Level "SUCCESS"
Write-Log "Offline: $($cameras.Count - $onlineCount)" -Level $(if ($cameras.Count -eq $onlineCount) { "SUCCESS" } else { "WARNING" })
Write-Log ""
Write-Log "Næste skridt:" -Level "INFO"
Write-Log "1. Åbn CameraDashboard.xlsx" -Level "INFO"
Write-Log "2. Gå til 'Status' sheet" -Level "INFO"
Write-Log "3. Højreklik på tabellen > Refresh" -Level "INFO"
Write-Log ""

# Hold vinduet åbent
Read-Host "Tryk Enter for at afslutte"
