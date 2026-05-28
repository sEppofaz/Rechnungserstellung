# Rechnungserstellung – Kargl

## Kerninfos

- **Service:** `/opt/kargl-invoice/`, Port 5002, systemd `kargl-invoice.service`
- **Lokale Arbeitskopie:** `~/Dropbox/Apps/Claude/Rechnungserstellung/app.py`
- **GitHub:** `https://github.com/sEppofaz/Rechnungserstellung`
- **Log:** `journalctl -u kargl-invoice -f`
- **nginx-Route:** `location /webhook-invoice → 127.0.0.1:5002`
- **Dropbox-Webhook-URL:** `https://umbenennen.duckdns.org/webhook-invoice`

### Deployment-Flow

```bash
# Code ändern → push → Server pull
git -C ~/Library/CloudStorage/Dropbox/Apps/Claude/Rechnungserstellung add app.py
git -C ~/Library/CloudStorage/Dropbox/Apps/Claude/Rechnungserstellung commit -m "..."
git -C ~/Library/CloudStorage/Dropbox/Apps/Claude/Rechnungserstellung push
ssh root@89.167.104.145 "git -C /opt/kargl-invoice/src pull && systemctl restart kargl-invoice"
```

### template.docx deployen

```bash
# Quelldatei: ~/Dropbox/Apps/Claude/Rechnung Kargl/VORL_Rechnungsformular 2026.dotx
# Schritt 1: .dotx → .docx konvertieren (Python, lokal)
cd ~/Library/CloudStorage/Dropbox/Apps/Claude/Rechnung\ Kargl
python3 - <<'EOF'
import zipfile
src = "VORL_Rechnungsformular 2026.dotx"
dst = "template.docx"
with zipfile.ZipFile(src, 'r') as zin, zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == '[Content_Types].xml':
            data = data.replace(b'wordprocessingml.template.main+xml',
                                b'wordprocessingml.document.main+xml')
        zout.writestr(item, data)
print("Fertig:", dst)
EOF

# Schritt 2: auf Server hochladen (kein Neustart nötig)
scp ~/Library/CloudStorage/Dropbox/Apps/Claude/Rechnung\ Kargl/template.docx \
  root@89.167.104.145:/opt/kargl-invoice/template.docx
```

---

## Architektur

**Primärer Flow (App):**
```
App /kargl/ → POST /kargl/api/ocr → Claude OCR → Felder anzeigen
→ POST /kargl/api/confirm → ODT → Dropbox Entwurf + Register
```

**Fallback (Dropbox-Webhook):**
```
Rechnungen_Input/ → POST /webhook-invoice → Claude OCR → ODT automatisch
```

### Server-Struktur

```
/opt/kargl-invoice/
├── src/
│   ├── app.py              ← gesamte Logik (standalone, kein shared-Import)
│   ├── kargl_app.html      ← Review-App PWA (git)
│   └── requirements.txt
├── template.docx           ← Word-Vorlage (außerhalb git, per scp deployen)
├── invoice_cursor.txt      ← Dropbox-Cursor (außerhalb git)
├── sessions/               ← temporäre OCR-Sessions (max 24h, auto-cleanup)
├── icons/                  ← generierte PWA-Icons (auto-generiert beim Start)
└── bin/                    ← Python venv
```

### Dropbox-Struktur

```
/_Austauschordner-Sandra-sEpp/Kargl-Rechnung/
    Rechnungen_Input/       ← Webhook-Fallback: Scan hier ablegen
    Rechnungen_Entwurf/     ← fertige .odt erscheint hier
    Rechnungen_Erledigt/    ← verarbeitete Originale
    Rechnungen_Fehler/      ← nicht verarbeitbare Dateien
    _Adressen.xlsx          ← Kundenadressen (automatisch gepflegt)
    _Rechnungsregister.xlsx ← alle Rechnungen mit Nr., Name, Beträgen
```

---

## Rechnungsnummer

Format `NNN00JJ`: z.B. `0170026` = laufende Nr. 017, Jahr 2026.
Wird automatisch aus `_Rechnungsregister.xlsx` hochgezählt. Jahreswechsel → Reset auf `001`.

---

## Template-Platzhalter

| Platzhalter | Inhalt |
|-------------|--------|
| `{{rechnungsnummer}}` | z.B. `0170026` |
| `{{anrede}}` | Firma / Herr / Frau |
| `{{name}}` | Kundenname |
| `{{strasse_nr}}` | Straße + Nr. (ggf. `[BITTE PRÜFEN]`) |
| `{{plz}}` / `{{ort}}` | Adresse |
| `{{datum}}` | Verarbeitungstag (TT.MM.JJJJ) |
| `{{beschreibungstext}}` | Leistungstext vom Zettel |
| `{{position1..6}}` | Menge (z.B. `1,66 cbm`) |
| `{{einzelpreis1..6}}` | Einzelpreis (z.B. `110,00 €`) |
| `{{gesamtpreis1..6}}` | Zeilenbetrag |
| `{{netto}}` / `{{mwst}}` / `{{brutto}}` | Beträge |
| `{{hinweis}}` | Zahlungshinweis (optional) |

---

## Zwei Verarbeitungs-Modi

**Normalfall:** cbm × Einzelpreis → Python rechnet nach, vergleicht mit Zettelwerten (±0,02 € Toleranz).

**Pauschalbetrag-Modus:** Kein Einzelpreis auf Zettel → Brutto direkt übernehmen, Netto/MwSt rückrechnen.

---

## Pitfalls

- **`INVOICE_CURSOR_FILE` und `INVOICE_TEMPLATE`** zeigen auf `/opt/kargl-invoice/`, nicht `/opt/rename-webhook/` – Verwechslung historisch möglich
- **`CLAUDE_API_KEY`** (nicht `ANTHROPIC_API_KEY`) – so benannt in `secrets.env`
- **Kein shared-Import:** `app.py` ist vollständig standalone – `MEDIA_TYPES`, `log` etc. sind direkt definiert, kein Import aus Vereinskalender
- **template.docx liegt außerhalb git** – bei Änderungen an `.dotx` konvertieren und per scp deployen (kein Service-Restart nötig)
- **Cursor liegt außerhalb git** – `invoice_cursor.txt` unter `/opt/kargl-invoice/`, nicht unter `/src/`
- **Claude-Antwort manchmal in Markdown-Backticks** – Code strippt ` ```json ``` ` vor `json.loads()`
- **Adresse bekannter Kunde** → wird aus `_Adressen.xlsx` übernommen, kein Nominatim-Aufruf
- **`[BITTE PRÜFEN]`** erscheint im Dokument wenn: Adresse korrigiert, Adresse nicht verifiziert, Rechenabweichung >0,02 €
- **nginx `sites-enabled` ist eine Kopie, kein Symlink** – Änderungen an `sites-available/rename-webhook` müssen immer mit `cp sites-available/rename-webhook sites-enabled/rename-webhook` übernommen werden, sonst bleibt nginx auf dem alten Stand
- **LibreOffice OCR-Timeout** – Konvertierung kann bis zu 30s dauern; bei sehr großen Dateien ggf. Timeout anpassen
- **`KARGL_APP_TOKEN` fehlt → App zeigt Login, aber alle API-Calls liefern 401** – Token in `/etc/pka/secrets.env` eintragen + `systemctl restart kargl-invoice`

---

## Häufige Änderungen

### Neues Feld hinzufügen
1. Platzhalter `{{neues_feld}}` ins `.dotx`-Template
2. Template konvertieren + per scp deployen
3. In `extract_invoice_data()` → `user_prompt` ins JSON-Schema aufnehmen
4. In `build_docx()` → `context`-Dict ergänzen
5. `app.py` committen + deployen

### MwSt-Satz ändern (derzeit 19%)
In `app.py`, Funktion `calculate_and_validate()`:
```python
mwst = round(netto * 0.19, 2)
```

### Modell wechseln (ohne Code-Änderung)
In `/etc/pka/secrets.env` (Josef fragen):
```
CLAUDE_INVOICE_MODEL=claude-opus-4-7
```

---

## Kargl Review-App (`/kargl/`)

**URL:** `https://umbenennen.duckdns.org/kargl/`
**Auth:** Token-Login (Bearer-Token, `KARGL_APP_TOKEN` in secrets.env)

### Neuer Flow (primär)

```
App: Foto aufnehmen / Datei wählen
    ↓ POST /kargl/api/ocr (multipart, Bearer-Token)
Server: Claude OCR → JSON-Felder zurück + session_id
    ↓ (~5–15 Sek)
App: Editierbare Felder + Bildvorschau
    ↓ POST /kargl/api/confirm (editierte Felder + session_id)
Server: Rechnungsnummer vergeben, docx → ODT, Dropbox Entwurf,
        Adressen.xlsx (neue Kunden), Rechnungsregister.xlsx
    ↓
App: "✅ Rechnung erstellt"
```

### Session-Files

- `/opt/kargl-invoice/sessions/{uuid}.jpg` – temporäre Scan-Kopie (max 24h)
- `/opt/kargl-invoice/sessions/{uuid}.json` – Session-Meta
- Automatische Bereinigung (>24h) beim nächsten OCR-Aufruf

### Icons

- Generiert beim Service-Start via PIL in `/opt/kargl-invoice/icons/`
- Falls Generierung fehlschlägt: 404 bei Icon-Request (unkritisch)

---

## Ausgabeformat

- **App-Flow:** ODT via LibreOffice headless (`libreoffice --headless --convert-to odt`)
- **Webhook-Fallback:** ebenfalls ODT; fällt auf `.docx` zurück wenn LibreOffice fehlt
- LibreOffice ist installiert unter `/usr/bin/libreoffice`

---

## Secrets (`/etc/pka/secrets.env`)

Niemals direkt lesen – Josef fragen. Benötigte Keys:
- `DROPBOX_INVOICE_REFRESH_TOKEN`, `DROPBOX_INVOICE_APP_KEY`, `DROPBOX_INVOICE_APP_SECRET`
- `CLAUDE_API_KEY`
- `KARGL_APP_TOKEN` – Zugangscode für die Review-App (Josef + Sandra)
- Optional: `CLAUDE_INVOICE_MODEL` (Default: `claude-sonnet-4-6`)
