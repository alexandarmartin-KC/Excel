# 📊 Kamera Dashboard - Excel-baseret Netværksovervågning

> **Enterprise-ready Excel dashboard** til overvågning af IP-kameraer uden makroer eller automatisk kørsel.

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://github.com/PowerShell/PowerShell)
[![Excel](https://img.shields.io/badge/Excel-2013%2B-green.svg)](https://www.microsoft.com/excel)

---

## 📋 Indholdsfortegnelse

- [Oversigt](#-oversigt)
- [Features](#-features)
- [Sikkerhed & Compliance](#-sikkerhed--compliance)
- [Forudsætninger](#-forudsætninger)
- [Installation](#-installation)
- [Brug](#-brug)
- [Projekt-struktur](#-projekt-struktur)
- [Konfiguration](#-konfiguration)
- [Fejlfinding](#-fejlfinding)
- [Bidrag](#-bidrag)
- [Licens](#-licens)

---

## 🎯 Oversigt

Dette projekt leverer et **Excel-baseret dashboard** til overvågning af IP-kameraer i virksomhedsnetværk med **strenge IT-sikkerhedskrav**:

- ✅ **Ingen Excel-makroer** (kun .xlsx format)
- ✅ **Ingen automatisk kørsel** (kræver manuel brugerhandling)
- ✅ **Ingen VBA-kode** (helt ren datafil)
- ✅ **Manuel refresh** (ingen real-time opdateringer)
- ✅ **PowerShell-baseret testing** (standard Windows-værktøj)

### Hvordan virker det?

```
┌─────────────────┐       ┌──────────────────┐       ┌─────────────────┐
│   CameraList    │       │  Run-Test.cmd    │       │   Status        │
│   (Excel Input) │ ───>  │  (Manuel start)  │ ───>  │   (Excel Output)│
│                 │       │                  │       │                 │
│ • ID            │       │ 1. Læs liste     │       │ • ONLINE/OFFLINE│
│ • Name          │       │ 2. Test ping     │       │ • Response time │
│ • IP Address    │       │ 3. Test HTTP     │       │ • Timestamps    │
│ • Location      │       │ 4. Gem CSV       │       │ • Refresh data  │
└─────────────────┘       └──────────────────┘       └─────────────────┘
```

---

## ✨ Features

### Core Funktionalitet
- 📝 **Vedligehold kamera-liste** i Excel (tabel-format)
- 🔍 **Test kamera-status** via PowerShell (ping + HTTP port check)
- 📊 **Visuel status-oversigt** med betinget formatering
- 📈 **Dashboard** med nøgletal (total, online, offline)
- 🔄 **Manuel data-refresh** (ingen automation)

### Test-metoder
- **ICMP Ping**: Verificer netværksforbindelse
- **HTTP Port Check**: Tjek at web-interface er tilgængeligt (port 80)
- **Response Time**: Måling af ping-responstid i millisekunder

### Output
- CSV-fil med test-resultater (`/output/CameraStatus.csv`)
- Excel-import via Data Connection (manuel refresh)
- Farve-kodede status-indikatorer (grøn/rød)

---

## 🔒 Sikkerhed & Compliance

Dette projekt er designet til virksomheder med **strenge IT-sikkerhedspolicies**:

| Krav | Status | Beskrivelse |
|------|--------|-------------|
| **Ingen makroer** | ✅ | Excel-fil er .xlsx (ikke .xlsm) |
| **Ingen VBA** | ✅ | Ingen Visual Basic kode |
| **Ingen auto-run** | ✅ | Kræver manuel dobbeltklik på .cmd |
| **Ingen scheduled tasks** | ✅ | Ingen baggrundsprocesser |
| **Ingen real-time polling** | ✅ | Kun test når bruger starter det |
| **Ingen Excel COM automation** | ⚠️ | Kan deaktiveres (fallback til CSV) |
| **Ingen external dependencies** | ✅ | Kun standard Windows-værktøjer |

### IT-godkendelse

**Velegnet til**:
- Virksomheder med deaktiverede Excel-makroer
- Miljøer hvor VBA er blokeret
- Netværk med streng application control
- Compliance-krævende brancher (finans, sundhed, offentlig)

**Ikke-velegnet til**:
- Real-time overvågning (ingen live-opdateringer)
- Automatisk alerting (ingen notifikationer)
- Store netværk (>100 kameraer kan være langsomt)

---

## 🔧 Forudsætninger

### Software
- **Microsoft Excel** 2013 eller nyere
- **Windows PowerShell** 5.1 eller nyere (inkluderet i Windows)
- **Windows** 10 eller nyere (eller Windows Server 2016+)

### Netværksadgang
- Adgang til IP-kameraernes netværk
- Tilladelse til at sende ICMP ping
- Tilladelse til TCP port 80 (HTTP)

### Brugerrettigheder
- Læse/skrive adgang til projekt-mappen
- Tilladelse til at køre PowerShell scripts (ExecutionPolicy)

---

## 📦 Installation

### 1. Clone Repository

```bash
git clone https://github.com/alexandarmartin-KC/Excel.git
cd Excel
```

### 2. Verificer Fil-struktur

```
Excel/
├── README.md
├── LICENSE
├── .gitignore
├── excel/
│   ├── BUILD_STEPS.md          ⬅️ LÆST DETTE FØRST!
│   └── CameraDashboard.xlsx    (oprettes manuelt)
├── scripts/
│   ├── Check-Cameras.ps1       ⬅️ Test-script
│   └── Run-Test.cmd            ⬅️ Double-click runner
├── examples/
│   └── CameraList_Example.csv  (fallback data)
└── output/
    └── CameraStatus.csv        (genereres automatisk)
```

### 3. Opret Excel-fil

**VIGTIGT**: Excel-filen inkluderes IKKE i repo (binær fil).

Følg den detaljerede guide:
```
📖 Se: /excel/BUILD_STEPS.md
```

Det tager ca. **10 minutter** at oprette filen korrekt.

### 4. Konfigurer PowerShell (hvis nødvendigt)

Hvis du får fejl om "ExecutionPolicy", kør som Administrator:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

---

## 🚀 Brug

### Daglig Arbejdsgang

#### **Trin 1: Tilføj/Opdater Kameraer**

1. Åbn `CameraDashboard.xlsx`
2. Gå til **"CameraList"** sheet
3. Tilføj nye rækker i tabellen:
   
   | ID | CameraName | IpAddress | Location | Notes |
   |----|-----------|-----------|----------|-------|
   | 1 | Hovedindgang | 192.168.1.101 | Indgang 1. sal | Dome |
   | 2 | Parkering | 192.168.1.102 | P-kælder | PTZ |

4. Gem filen (`Ctrl+S`)

#### **Trin 2: Kør Test**

1. Dobbeltklik på: **`/scripts/Run-Test.cmd`**
2. Bekræft at du vil starte testen (tryk Enter)
3. Vent på testen afslutter (kan tage 10-30 sekunder)
4. Se output i konsol-vinduet
5. Tryk Enter for at lukke vinduet

**Eksempel output**:
```
========================================
  KAMERA DASHBOARD - TEST STARTER
========================================
Excel fil: C:\...\excel\CameraDashboard.xlsx
Output fil: C:\...\output\CameraStatus.csv

Tester 4 kameraer...
----------------------------------------
Testing Hovedindgang (192.168.1.101)... ✓ ONLINE (12ms)
Testing Parkering (192.168.1.102)... ✓ ONLINE (8ms)
Testing Reception (192.168.1.103)... ✗ OFFLINE
Testing Lager (192.168.1.104)... ✓ ONLINE (15ms)

========================================
  TEST AFSLUTTET
========================================
Total kameraer: 4
Online: 3
Offline: 1
```

#### **Trin 3: Se Resultater i Excel**

1. Åbn `CameraDashboard.xlsx` (hvis lukket)
2. Gå til **"Status"** sheet
3. Højreklik i tabellen
4. Vælg: **Refresh** (eller tryk `Ctrl+Alt+F5`)
5. Se opdaterede status:
   - 🟢 **ONLINE** = Grøn
   - 🔴 **OFFLINE** = Rød

#### **Trin 4: Se Oversigt** (valgfrit)

1. Gå til **"Dashboard"** sheet
2. Se sammenfatning:
   - Total Kameraer: 4
   - Online: 3
   - Offline: 1
   - Sidste Test: 2026-01-06 14:32:15

---

## 📁 Projekt-struktur

```
Excel/
│
├── 📄 README.md                    # Denne fil
├── 📄 LICENSE                      # MIT licens
├── 📄 .gitignore                   # Git ignore regler
│
├── 📂 excel/                       # Excel-filer
│   ├── 📖 BUILD_STEPS.md           # Guide til at oprette Excel-fil
│   └── 📊 CameraDashboard.xlsx     # Main dashboard (oprettes manuelt)
│
├── 📂 scripts/                     # PowerShell scripts
│   ├── 💻 Check-Cameras.ps1        # Test-logik (ping + HTTP)
│   └── ▶️  Run-Test.cmd            # Double-click runner
│
├── 📂 examples/                    # Eksempel-data
│   └── 📄 CameraList_Example.csv   # Fallback kamera-liste
│
└── 📂 output/                      # Genererede filer
    └── 📄 CameraStatus.csv         # Test-resultater (auto-genereret)
```

### Fil-beskrivelser

| Fil | Formål | Type |
|-----|--------|------|
| `CameraDashboard.xlsx` | Excel dashboard med data-input og visualisering | Excel Workbook |
| `Check-Cameras.ps1` | PowerShell script til at teste kameraer | PowerShell |
| `Run-Test.cmd` | CMD wrapper til at starte PowerShell script | Batch Script |
| `CameraStatus.csv` | Output fra test (importeres til Excel) | CSV Data |
| `BUILD_STEPS.md` | Detaljeret guide til Excel-opsætning | Markdown Doc |

---

## ⚙️ Konfiguration

### PowerShell Script Parametre

Åbn `Check-Cameras.ps1` og tilpas disse parametre:

```powershell
param(
    [string]$ExcelFile = "$PSScriptRoot\..\excel\CameraDashboard.xlsx",
    [string]$OutputFile = "$PSScriptRoot\..\output\CameraStatus.csv",
    [int]$TimeoutSeconds = 2,      # Ping timeout
    [int]$HttpPort = 80             # HTTP port (ændr til 443 for HTTPS)
)
```

### Test HTTPS i stedet for HTTP

Hvis dine kameraer bruger HTTPS (port 443), ændr:

```powershell
[int]$HttpPort = 443
```

### Tilpas Timeout

For langsomme netværk, øg timeout:

```powershell
[int]$TimeoutSeconds = 5
```

---

## 🐛 Fejlfinding

### Problem: "PowerShell script kan ikke køres"

**Fejl**: `execution of scripts is disabled on this system`

**Løsning**:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Problem: "Excel-fil ikke fundet"

**Fejl**: `FEJL: Excel-filen findes ikke: ...`

**Løsning**:
1. Tjek at Excel-filen er oprettet: `/excel/CameraDashboard.xlsx`
2. Følg guiden i: `/excel/BUILD_STEPS.md`

### Problem: "Alle kameraer vises som OFFLINE"

**Mulige årsager**:
- Firewall blokerer ICMP ping
- Forkert IP-adresser i Excel
- Kameraer er faktisk offline
- Netværks-segmentering (VLAN isolation)

**Debug**:
```powershell
# Test ping manuelt
ping 192.168.1.101

# Test HTTP port manuelt
Test-NetConnection -ComputerName 192.168.1.101 -Port 80
```

### Problem: "Kan ikke refreshe data i Excel"

**Løsning**:
1. Sørg for CSV-filen eksisterer: `/output/CameraStatus.csv`
2. Genopret data connection:
   - Gå til: **Data > Queries & Connections**
   - Slet eksisterende connection
   - Importer CSV igen via: **Data > From Text/CSV**

### Problem: "Excel COM error"

**Fejl**: `New-Object : Cannot create object "Excel.Application"`

**Løsning**:
- Excel er ikke installeret korrekt
- Brug fallback CSV-metode:
  ```powershell
  # Tilføj kamera-liste til /examples/CameraList_Example.csv
  # Scriptet vil automatisk bruge denne
  ```

---

## 🤝 Bidrag

Bidrag er velkomne! Følg disse guidelines:

### Rapporter Bugs
1. Tjek om issue allerede eksisterer
2. Opret ny issue med:
   - Klar beskrivelse af problemet
   - Trin til at reproducere
   - Forventet vs. faktisk adfærd
   - Screenshots (hvis relevant)

### Foreslå Features
1. Opret issue med label `enhancement`
2. Beskriv feature og use case
3. Diskuter approach med maintainers

### Pull Requests
1. Fork repository
2. Opret feature branch: `git checkout -b feature/AmazingFeature`
3. Commit ændringer: `git commit -m 'Add AmazingFeature'`
4. Push til branch: `git push origin feature/AmazingFeature`
5. Åbn Pull Request

### Code Standards
- PowerShell: Følg [PowerShell Best Practices](https://poshcode.gitbook.io/powershell-practice-and-style/)
- Dokumentation: Opdater README.md ved feature-ændringer
- Testing: Test på ren Windows-installation

---

## 📜 Licens

Dette projekt er licenseret under **MIT License** - se [LICENSE](LICENSE) filen for detaljer.

```
MIT License

Copyright (c) 2026 Alexandar Martin

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction...
```

---

## 📞 Support

### Dokumentation
- **Excel Setup**: [excel/BUILD_STEPS.md](excel/BUILD_STEPS.md)
- **Script Details**: Se kommentarer i PowerShell-filer

### Kontakt
- **GitHub Issues**: [https://github.com/alexandarmartin-KC/Excel/issues](https://github.com/alexandarmartin-KC/Excel/issues)
- **Email**: (tilføj hvis relevant)

### FAQ

**Q: Kan jeg bruge dette til andre enheder end kameraer?**  
A: Ja! Systemet kan teste alle IP-enheder (printere, servere, IoT-devices, etc.)

**Q: Virker det på Mac/Linux?**  
A: Nej, dette projekt kræver Windows (Excel + PowerShell).

**Q: Kan jeg automatisere testen?**  
A: Teknisk ja, men det bryder design-principperne for IT-compliance. Brug scheduled task på eget ansvar.

**Q: Understøtter det HTTPS/TLS?**  
A: Ja, ændr `$HttpPort` til 443 i scriptet.

**Q: Hvor mange kameraer kan det håndtere?**  
A: Testet med op til 50 kameraer (~60 sekunder test-tid). For >100 kameraer, overvej parallel processing.

---

## 🎓 Credits

Udviklet med fokus på:
- **Enterprise compliance** (ingen makroer/VBA)
- **IT-sikkerhed** (manuel kørsel)
- **Brugervenlighed** (Excel-interface)
- **Simplicitet** (standard Windows-værktøjer)

---

## 🔄 Versionshistorik

### v1.0.0 (2026-01-06)
- ✨ Initial release
- ✅ Excel-baseret dashboard (ingen makroer)
- ✅ PowerShell test-script (ping + HTTP)
- ✅ CMD double-click runner
- ✅ CSV output med manuel refresh
- ✅ Betinget formatering (ONLINE/OFFLINE)
- ✅ Dashboard med nøgletal
- ✅ Komplet dokumentation

---

<div align="center">

**⭐ Star dette repo hvis det er nyttigt! ⭐**

Made with ❤️ for IT-afdelinger der værdsætter sikkerhed

[Report Bug](https://github.com/alexandarmartin-KC/Excel/issues) · [Request Feature](https://github.com/alexandarmartin-KC/Excel/issues)

</div>