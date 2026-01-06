# BUILD_STEPS.md - Opret CameraDashboard.xlsx

**VIGTIGT**: Excel-filen skal oprettes manuelt, da dette repo ikke indeholder binære filer (.xlsx). Følg disse præcise trin.

---

## 📋 FORUDSÆTNINGER
- Microsoft Excel installeret (2013 eller nyere)
- 10 minutter til opsætning

---

## 🔨 TRIN-FOR-TRIN GUIDE

### TRIN 1: Opret ny Excel-fil

1. Åbn Microsoft Excel
2. Opret en ny tom projektmappe (Blank Workbook)
3. Gem filen som: `CameraDashboard.xlsx`
   - Placer den i mappen: `/excel/`
   - **VIGTIGT**: Gem som `.xlsx` (IKKE `.xlsm` - ingen makroer!)

---

### TRIN 2: Opret "CameraList" sheet

1. Omdøb `Sheet1` til: **`CameraList`**
2. I celle A1, indtast følgende kolonner:

   | A1 | B1 | C1 | D1 | E1 |
   |----|----|----|----|-----|
   | **ID** | **CameraName** | **IpAddress** | **Location** | **Notes** |

3. Formater header-række:
   - Markér A1:E1
   - **Fed skrift** (Ctrl+B)
   - **Baggrund**: Blå (eller din foretrukne farve)
   - **Tekstfarve**: Hvid

4. **VIGTIGT**: Konverter til Excel Table:
   - Markér A1:E1 (header række)
   - **Dansk Excel**: Gå til **Indsæt > Tabel** (eller tryk `Ctrl + L`)
   - **Engelsk Excel**: Tryk `Ctrl + T` eller gå til **Insert > Table**
   - Sørg for "Tabellen har overskrifter" / "My table has headers" er markeret
   - Klik OK
   - Omdøb tabellen til: **`tblCameras`**
     - Højreklik på tabellen > Tabel > Tabelnavn
     - Eller via: **Tabeldesign > Tabelnavn** feltet (øverst til venstre når tabel er valgt)

5. Tilføj eksempel-data (valgfrit):
   
   | ID | CameraName | IpAddress | Location | Notes |
   |----|-----------|-----------|----------|-------|
   | 1 | Hovedindgang | 192.168.1.101 | Hovedindgang 1. sal | Dome kamera |
   | 2 | Parkering A | 192.168.1.102 | P-kælder Nord | |
   | 3 | Reception | 192.168.1.103 | Reception | PTZ kamera |
   | 4 | Lager | 192.168.1.104 | Baglager | Infrarød |

6. Juster kolonnebredder (valgfrit):
   - Dobbeltklik på kolonne-dividers for auto-fit

---

### TRIN 3: Opret "Status" sheet

1. Tilføj nyt sheet:
   - Klik på `+` ved bundmenuen
   - Omdøb til: **`Status`**

2. I celle A1, indtast følgende kolonner:

   | A1 | B1 | C1 | D1 | E1 | F1 | G1 | H1 | I1 |
   |----|----|----|----|----|----|----|----|-----|
   | **ID** | **CameraName** | **IpAddress** | **Location** | **Status** | **PingStatus** | **HttpStatus** | **ResponseTime** | **LastChecked** |

3. Formater header-række:
   - Markér A1:I1
   - **Fed skrift** (Ctrl+B)
   - **Baggrund**: Grøn
   - **Tekstfarve**: Hvid

4. **VIGTIGT**: Importer data fra CSV:
   - Klik i celle A2 (første data-række)
   - **Dansk Excel**: Gå til **Data > Hent data > Fra fil > Fra tekst/CSV**
     - (I ældre Excel: **Data > Fra tekst/CSV**)
   - **Engelsk Excel**: Gå til **Data > Get Data > From File > From Text/CSV**
   - Vælg fil: `/output/CameraStatus.csv`
     - **BEMÆRK**: Denne fil eksisterer ikke endnu - den oprettes når du kører første test
     - Hvis filen ikke findes endnu, spring dette trin over og kom tilbage efter første test
   - Indstillinger:
     - **Delimiter**: Semicolon (`;`)
     - **File Origin**: UTF-8
   - Klik: **Load**

5. **VIGTIGT**: Opret "Data Connection" som kan refreshes:
   - Efter import, højreklik i tabellen
   - Vælg: **Data > Opdater alle** / **Refresh All**
   - Nu kan du altid opdatere ved at trykke "Opdater" / "Refresh"

6. **ALTERNATIV METODE** (hvis Hent Data ikke virker):
   - **Dansk Excel**: Brug **Data > Fra tekst (ældre)** / **From Text (Legacy)**
   - **Engelsk Excel**: Brug **Data > From Text (Legacy)**
   - Eller manuelt copy-paste fra CSV efter første test

---

### TRIN 4: Tilføj betinget formatering til Status sheet

1. Markér kolonne E (Status) - alle data-celler (E2:E1000)
2. **Dansk Excel**: Gå til **Hjem > Betinget formatering > Fremhæv celleregler > Tekst, der indeholder**
3. **Engelsk Excel**: Gå til **Home > Conditional Formatting > Highlight Cell Rules > Text that Contains**
4. Opret regel for ONLINE:
   - Tekst: `ONLINE`
   - Format: **Grøn fyld med mørkegrøn tekst**
5. Opret regel for OFFLINE:
   - Tekst: `OFFLINE`
   - Format: **Rød fyld med mørkerød tekst**

---

### TRIN 5: Opret "Dashboard" sheet (valgfrit, men anbefalet)

1. Tilføj nyt sheet: **`Dashboard`**

2. Opret sammenfatning (i celle A1):

   ```
   ┌─────────────────────────────────────┐
   │  KAMERA DASHBOARD - OVERSIGT         │
   └─────────────────────────────────────┘
   
   Total Kameraer:     [FORMEL]
   Online:             [FORMEL]
   Offline:            [FORMEL]
   Sidste Test:        [FORMEL]
   ```

3. Indsæt formler:
   - **B4** (Total): `=COUNTA(Status!A:A)-1`
   - **B5** (Online): `=COUNTIF(Status!E:E,"ONLINE")`
   - **B6** (Offline): `=COUNTIF(Status!E:E,"OFFLINE")`
   - **B7** (Sidste test): `=MAX(Status!I:I)`

4. Formater:
   - Gør celle-områder tydelige med borders
   - Brug stor skrift til tal (18pt)
   - Tilføj betinget formatering for farver

---

### TRIN 6: Tilføj "Instructions" sheet (valgfrit)

1. Tilføj nyt sheet: **`Instructions`**
2. Tilføj bruger-instruktioner:

   ```
   ═══════════════════════════════════════════════════
   KAMERA DASHBOARD - BRUGERVEJLEDNING
   ═══════════════════════════════════════════════════

   1. TILFØJ KAMERAER
      - Gå til "CameraList" sheet
      - Tilføj nye rækker i tabellen
      - Udfyld: ID, CameraName, IpAddress, Location

   2. KØR TEST
      - Dobbeltklik på: /scripts/Run-Test.cmd
      - Vent på testen afslutter
      - Luk PowerShell-vinduet

   3. SE RESULTATER
      - Åbn denne Excel-fil (hvis lukket)
      - Gå til "Status" sheet
      - Højreklik på tabellen > Refresh
      - Se opdaterede status for alle kameraer

   4. SE OVERSIGT
      - Gå til "Dashboard" sheet
      - Se sammenfatning af online/offline status

   ═══════════════════════════════════════════════════
   BEMÆRK:
   - Ingen makroer - helt sikkert!
   - Test kører KUN når du starter den manuelt
   - Ingen automatisk opdatering
   - Data gemmes i /output/CameraStatus.csv
   ═══════════════════════════════════════════════════
   ```

---

### TRIN 7: Færdiggør og gem

1. **Arranger sheets i rækkefølge**:
   - Dashboard (først)
   - CameraList
   - Status
   - Instructions

2. **Beskyt strukturen** (valgfrit):
   - **Dansk Excel**: Gå til **Gennemse > Beskyt projektmappe** / **Review > Protect Workbook**
   - **Engelsk Excel**: Gå til **Review > Protect Workbook**
   - Sæt password (kun hvis påkrævet)
   - Dette forhindrer utilsigtet sletning af sheets

3. **Gem filen**:
   - `Ctrl + S`
   - Bekræft at den er gemt som: `/excel/CameraDashboard.xlsx`

4. **Verificer**:
   - Luk Excel
   - Åbn filen igen
   - Tjek at alle sheets findes
   - Tjek at tabellen hedder `tblCameras`

---

## ✅ VERIFICERING

Før du går videre, tjek:

- [ ] Filen hedder `CameraDashboard.xlsx` (ikke .xlsm!)
- [ ] Filen ligger i `/excel/` mappen
- [ ] "CameraList" sheet eksisterer
- [ ] Excel Table med navn `tblCameras` eksisterer
- [ ] "Status" sheet eksisterer
- [ ] Kolonner matcher nøjagtigt som beskrevet
- [ ] Data connection til CSV er oprettet (eller klar til første import)

---

## 🚀 NÆSTE SKRIDT

1. Kør første test: Dobbeltklik på `/scripts/Run-Test.cmd`
2. Åbn Excel og refresh Status sheet
3. Se resultaterne!

---

## 🔧 FEJLFINDING
Hent data" / "Get Data" menu  
**Løsning**: Du har ældre Excel. Brug "Data > Fra tekst/CSV" eller "From Text (Legacy)"

**Problem**: Ctrl+T åbner typografi (dansk Excel)  
**Løsning**: I dansk Excel skal du bruge **Indsæt > Tabel** fra menuen, eller prøv `Ctrl+L`
**Løsning**: Du har ældre Excel. Brug "Data > From Text/CSV" eller "From Text (Legacy)"

**Problem**: CSV importeres ikke korrekt  
**Løsning**: Tjek at delimiter er `;` (semicolon), ikke `,` (comma)

**Problem**: Tabellen hedder ikke tblCameras  
**Løsning**: Højreklik på tabellen > Table > Table Name > Ændr til tblCameras

**Problem**: PowerShell scriptet kan ikke læse Excel  
**Løsning**: Sørg for Excel er installeret og fil-sti er korrekt

---

## 📝 NOTER TIL IT-AFDELINGEN

- **Ingen makroer**: Filen er .xlsx format
- **Ingen VBA kode**: Helt ren datafil
- **Ingen auto-refresh**: Kræver manuel brugerhandling
- **Ingen COM automation**: Scripts læser kun via COM (kan deaktiveres hvis nødvendigt)
- **CSV fallback**: Scripts kan læse fra CSV hvis COM blokeres
- **Ingen netværks-adgang fra Excel**: Kun fra PowerShell script
- **Ingen scheduled tasks**: Kun manuel kørsel

---

**Spørgsmål?** Se README.md eller kontakt systemadministrator.
