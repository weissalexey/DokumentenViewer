# DokumentenViewer (Chr. Carstensen Logistik)

Interne Desktop-App (Windows) zum schnellen Sichten, Zusammenführen und Exportieren von gescannten Dokumenten (PDF/JPG/PNG) inkl. TXT/LIS für DMS/GetMyInvoices.

> Fokus: schneller Operator-Workflow + minimale Klicks + stabile Standardpfade je Filiale.

---

## ✨ Highlights

- **Vorschau** für PDF und Bilder (JPG/PNG)
- **PDF-Merge**: mehrere Eingänge zu einer PDF pro Auftrag + Dokumenttyp
- **TXT/LIS-Generierung** passend zum PDF-Namen (z. B. `12345678_Eingangsbelege.txt`)
- **OCR per Maus**:
  - Bereich im Vorschaufenster markieren → Nummer wird erkannt
  - bei mehreren Treffern → Auswahl per Taste **1..9**
  - erkannte Bereiche werden **grün/rot** markiert
- **Filial-Logik** (10/40/50):
  - Filiale 10: AUFNR **genau 8-stellig**
  - Filiale 40/50: AUFNR **nur Ziffern**, Länge variabel (z. B. 10-stellig)
- **Alt+N** blendet nur „Ziel“ ein/aus (Standard: verborgen)
- **config.ini** wird beim Beenden gespeichert und beim Start geladen (AppData)

---

## 🖼️ Screenshots

> Lege Screenshots unter `docs/screenshots/` ab und committe sie, dann werden sie hier angezeigt.

### Main Window
![Main Window](docs/screenshots/main.png)

### OCR-Auswahl (1..9)
![OCR Selection](docs/screenshots/ocr_selection.png)

### Ziel ein-/ausblenden (Alt+N)
![Toggle Ziel](docs/screenshots/toggle_ziel.png)

---

## 🚀 Quickstart (Anwender)

1. Programm starten: `DokumentenViewer.exe`
2. **Load** → Dateien laden
3. AUFNR:
   - manuell eingeben **oder**
   - per OCR: Bereich mit Maus markieren
4. Dokumenttyp auswählen
5. **Save** → PDF+TXT erstellen, Quelldatei wird gelöscht, nächste Datei wird geladen

---

## ⌨️ Hotkeys

| Taste | Aktion |
|------:|--------|
| **F1** | Hilfe öffnen |
| **Ctrl+S** | Save |
| **Alt+N** | Ziel ein-/ausblenden |
| **← / →** | Vorherige / nächste Datei |
| **↑ / ↓** | PDF-Seite hoch / runter |
| **Delete** | Datei löschen |

---

## 🧠 OCR (Tesseract)

Die App nutzt `pytesseract` und benötigt eine installierte Tesseract-Version.

**Standardpfad:**
- `C:\Program Files\Tesseract-OCR\tesseract.exe`

Falls `tesseract --version` im CMD nicht funktioniert, ist das ok – die App kann trotzdem laufen,
wenn der Pfad im Code gesetzt ist (`TESSERACT_EXE`).

---

## 🏗️ Build (EXE)

### Voraussetzungen (Build-PC)
- Python 3.x
- `pip install -r requirements.txt`

### Build (PyInstaller)
Empfohlen (onefile + windowed + assets):

```bat
pyinstaller --onefile --windowed --name DokumentenViewer --icon assets\carstensen.ico ^
  --add-data "assets\carstensen.ico;." ^
  --add-data "assets\logo.png;." ^
  src\NEW.py
