# TextCraft Word-Add-in – Anleitung für Microsoft Word 2013

> **Sprache:** Deutsch | **Zielversion:** Microsoft Word 2013 (Windows Desktop)

---

## Inhaltsverzeichnis

1. [Voraussetzungen](#1-voraussetzungen)
2. [Office-Add-ins in Word 2013 installieren](#2-office-add-ins-in-word-2013-installieren)
3. [Was tun, wenn der Store / „Office-Add-ins" fehlt?](#3-was-tun-wenn-der-store--office-add-ins-fehlt)
4. [TextCraft öffnen und erste Schritte](#4-textcraft-öffnen-und-erste-schritte)
5. [Grundworkflow: Text bearbeiten mit TextCraft](#5-grundworkflow-text-bearbeiten-mit-textcraft)
6. [Häufige Probleme & Lösungen](#6-häufige-probleme--lösungen)
7. [Sicherheits- und Datenschutzhinweis](#7-sicherheits--und-datenschutzhinweis)
8. [FAQ](#8-faq)

---

## 1. Voraussetzungen

Bevor Sie TextCraft installieren, stellen Sie sicher, dass folgende Bedingungen erfüllt sind:

| Voraussetzung | Details |
|---|---|
| **Word-Version** | Microsoft Word 2013 (32-Bit oder 64-Bit), möglichst mit den neuesten Updates (SP1 oder höher) |
| **Betriebssystem** | Windows 7 oder neuer |
| **Microsoft-Konto** | Ein aktives Microsoft- oder Microsoft 365-Konto ist erforderlich, um den Office Add-in-Store zu nutzen |
| **Internetverbindung** | Für die Installation aus dem Store und für die KI-Funktionen von TextCraft |
| **Berechtigungen** | Sie müssen Add-ins installieren dürfen – bei Unternehmensgeräten ggf. IT-Abteilung kontaktieren |

---

## 2. Office-Add-ins in Word 2013 installieren

### Schritt 1 – Menüband öffnen

1. Starten Sie **Microsoft Word 2013**.
2. Klicken Sie im Menüband auf den Reiter **„Einfügen"**.

### Schritt 2 – Add-in-Bereich aufrufen

- Suchen Sie in der Gruppe **„Add-Ins"** nach einem der folgenden Einträge:
  - **„Office-Add-Ins"** (ältere Bezeichnung: **„Apps für Office"**)
  - **„Meine Add-Ins"**
- Klicken Sie auf das entsprechende Symbol.

> **Hinweis:** In Word 2013 heißt der Eintrag häufig noch **„Apps für Office"**. In neueren Service-Pack-Versionen kann er bereits als **„Office-Add-Ins"** erscheinen.

### Schritt 3 – Store öffnen

1. Im sich öffnenden Dialogfenster wählen Sie den Reiter **„Store"** (oder **„Office Store"**).
2. Geben Sie im Suchfeld **„TextCraft"** ein und drücken Sie **Enter**.

### Schritt 4 – TextCraft installieren

1. Klicken Sie in den Suchergebnissen auf **„TextCraft"**.
2. Klicken Sie auf **„Hinzufügen"** bzw. **„Abrufen"**.
3. Melden Sie sich mit Ihrem **Microsoft-Konto** an, falls Sie dazu aufgefordert werden.
4. Bestätigen Sie die Zustimmung zu den Nutzungsbedingungen.
5. Word fügt das Add-in automatisch hinzu und zeigt es in der Seitenleiste an.

---

## 3. Was tun, wenn der Store / „Office-Add-ins" fehlt?

Sollte der Menüpunkt **„Apps für Office"** oder **„Office-Add-Ins"** im Reiter „Einfügen" **nicht angezeigt** werden, kommen folgende Ursachen infrage:

### Mögliche Ursachen und Lösungen

| Ursache | Lösung |
|---|---|
| **Word 2013 nicht aktualisiert** | Installieren Sie **Service Pack 1 (SP1)** und alle aktuellen Windows-Updates. Office-Add-ins erfordern mindestens SP1. |
| **Keine Anmeldung mit Microsoft-Konto** | Klicken Sie oben rechts in Word auf **„Anmelden"** und melden Sie sich mit einem Microsoft- oder Geschäftskonto an. |
| **Unternehmens-/Gruppenrichtlinie blockiert Add-ins** | Wenden Sie sich an Ihre **IT-Abteilung**. Administratoren können Add-ins per Gruppenrichtlinie deaktivieren. |
| **Vertrauensstellungen für Add-ins deaktiviert** | Gehen Sie zu **Datei → Optionen → Trust Center → Einstellungen für das Trust Center → Vertrauenswürdige Add-In-Kataloge** und prüfen Sie die Einstellungen. |
| **Reiter „Einfügen" zeigt keine Add-in-Gruppe** | Klicken Sie mit der rechten Maustaste auf das Menüband → **„Menüband anpassen"** und prüfen Sie, ob die Gruppe „Add-Ins" in „Einfügen" aktiviert ist. |

### Alternative: Manuelle Sideload-Installation

Falls der Store nicht verfügbar ist, kann TextCraft über einen **freigegebenen Netzwerkordner (Katalog)** bereitgestellt werden:

1. Bitten Sie Ihren Administrator, einen freigegebenen Ordner als **vertrauenswürdigen Katalog** einzurichten.
2. Gehen Sie zu **Datei → Optionen → Trust Center → Einstellungen für das Trust Center → Vertrauenswürdige Add-In-Kataloge**.
3. Tragen Sie die **Netzwerkadresse** des Katalogs ein und klicken Sie auf **„Katalog hinzufügen"**.
4. Aktivieren Sie das Kontrollkästchen **„Im Menü anzeigen"** und klicken Sie auf **OK**.
5. Starten Sie Word neu und navigieren Sie zu **Einfügen → Meine Add-Ins → Freigegebener Ordner**.

---

## 4. TextCraft öffnen und erste Schritte

### Add-in nach der Installation öffnen

1. Klicken Sie auf **Einfügen → Meine Add-Ins**.
2. Wählen Sie im Dialogfenster unter **„Zuletzt verwendet"** oder **„Alle Add-Ins"** den Eintrag **„TextCraft"** aus.
3. Klicken Sie auf **„Hinzufügen"** (oder **„Einfügen"**).
4. Die **TextCraft-Seitenleiste** öffnet sich auf der rechten Seite Ihres Word-Dokuments.

### Übersicht der Seitenleiste

Die TextCraft-Seitenleiste enthält typischerweise:

- **Aktionsauswahl** – z. B. Text verbessern, zusammenfassen, übersetzen, umschreiben
- **Eingabefeld** – für zusätzliche Anweisungen oder Kontextinformationen
- **Ergebnisbereich** – zeigt den generierten Text an
- **Schaltflächen** – zum Übernehmen, Kopieren oder Verwerfen des Ergebnisses

---

## 5. Grundworkflow: Text bearbeiten mit TextCraft

Folgen Sie diesen Schritten, um Text in Ihrem Dokument mit TextCraft zu bearbeiten:

### Schritt 1 – Text markieren

Markieren Sie den Text in Ihrem Word-Dokument, den Sie bearbeiten möchten.

### Schritt 2 – Aktion auswählen

Wählen Sie in der TextCraft-Seitenleiste die gewünschte Aktion aus, z. B.:
- **„Text verbessern"** – stilistische und grammatikalische Verbesserungen
- **„Zusammenfassen"** – erstellt eine Kurzfassung des markierten Abschnitts
- **„Übersetzen"** – übersetzt den Text in eine andere Sprache
- **„Umschreiben"** – formuliert den Text anders (z. B. formeller / informeller)

### Schritt 3 – Anweisung verfeinern (optional)

Geben Sie im Eingabefeld der Seitenleiste zusätzliche Hinweise ein, z. B.:
> *„Bitte formal und auf Deutsch"*

### Schritt 4 – Ergebnis generieren

Klicken Sie auf **„Generieren"** (oder den entsprechenden Aktions-Button). TextCraft verarbeitet den Text und zeigt das Ergebnis im Ergebnisbereich an.

### Schritt 5 – Ergebnis übernehmen oder verwerfen

- **Übernehmen:** Klicken Sie auf **„Einfügen"** oder **„Ersetzen"**, um den markierten Text durch das Ergebnis zu ersetzen.
- **Kopieren:** Klicken Sie auf **„Kopieren"**, um das Ergebnis in die Zwischenablage zu kopieren.
- **Verwerfen:** Schließen Sie das Ergebnis, wenn Sie es nicht verwenden möchten. Ihr Originaldokument bleibt unverändert.

---

## 6. Häufige Probleme & Lösungen

### Problem: Add-in erscheint nicht unter „Meine Add-Ins"

**Mögliche Ursachen und Lösungen:**

1. **Deaktivierte Add-ins prüfen:**
   - Gehen Sie zu **Datei → Optionen → Add-Ins**.
   - Wählen Sie unten bei **„Verwalten"** den Eintrag **„Deaktivierte Elemente"** aus und klicken Sie auf **„Gehe zu"**.
   - Falls TextCraft in der Liste erscheint, wählen Sie es aus und klicken Sie auf **„Aktivieren"**.
   - Starten Sie Word anschließend neu.

2. **COM-Add-ins prüfen:**
   - Gehen Sie zu **Datei → Optionen → Add-Ins**.
   - Wählen Sie bei **„Verwalten"** den Eintrag **„COM-Add-Ins"** und klicken Sie auf **„Gehe zu"**.
   - Prüfen Sie, ob TextCraft vorhanden und aktiviert ist (Häkchen gesetzt).

### Problem: Seitenleiste öffnet sich nicht / bleibt leer

| Symptom | Lösung |
|---|---|
| Seitenleiste bleibt weiß/leer | Internetverbindung prüfen; Proxy/Firewall kann Add-in blockieren |
| Ladeanimation ohne Ergebnis | Word neu starten; Office-Cache leeren (siehe unten) |
| Fehlermeldung „Verbindung fehlgeschlagen" | Firewall-/Proxyeinstellungen überprüfen; ggf. IT-Abteilung informieren |

### Office-Cache leeren

1. Schließen Sie **alle Office-Programme**.
2. Löschen Sie den Inhalt des Ordners:
   ```
   %LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
   ```
   *(Diesen Pfad in den Windows-Explorer eingeben)*
3. Starten Sie Word neu und öffnen Sie TextCraft erneut.

### Problem: Add-in nach Neustart weg

- Stellen Sie sicher, dass Sie bei Ihrem **Microsoft-Konto** angemeldet sind.
- Installieren Sie TextCraft erneut über **Einfügen → Office-Add-Ins → Store**.

### Problem: Fehlermeldungen zu Proxy oder Firewall

TextCraft kommuniziert über HTTPS (Port 443) mit externen Servern. Bitten Sie Ihre IT-Abteilung, folgende Domains freizuschalten (falls bekannt):
- Die Domain des TextCraft-Dienstanbieters
- `*.office.com`, `*.microsoft.com` (für die Add-in-Infrastruktur)

---

## 7. Sicherheits- und Datenschutzhinweis

> ⚠️ **Wichtiger Hinweis:**

- Wenn Sie TextCraft verwenden, wird der **markierte Text aus Ihrem Dokument an einen externen KI-Dienst übertragen** und dort verarbeitet.
- Dies gilt für alle Aktionen (Verbessern, Übersetzen, Zusammenfassen usw.).
- **Übermitteln Sie keine vertraulichen, personenbezogenen oder sensiblen Daten** (z. B. Patientendaten, Geschäftsgeheimnisse, Passwörter) über TextCraft.
- Prüfen Sie die **Datenschutzrichtlinie** und die **Nutzungsbedingungen** von TextCraft, bevor Sie das Add-in in einem Unternehmensumfeld einsetzen.
- Bei Fragen zu Datenschutz und Compliance wenden Sie sich an Ihre **IT- oder Datenschutzabteilung**.

---

## 8. FAQ

**F: Kann ich TextCraft in Word 2013 ohne Microsoft-Konto nutzen?**

A: Für die Installation aus dem Office Store ist ein Microsoft-Konto erforderlich. Falls Ihr Administrator eine Sideload-Installation eingerichtet hat, kann TextCraft möglicherweise ohne Konto verwendet werden. Die KI-Funktionen selbst erfordern jedoch eine Internetverbindung.

---

**F: Kostet TextCraft etwas?**

A: TextCraft kann als kostenlose oder kostenpflichtige Version (mit Abonnement) verfügbar sein. Prüfen Sie die aktuellen Preisdetails im Office Store oder auf der Website des Anbieters.

---

**F: Funktioniert TextCraft auch im Word Online (Browser-Version)?**

A: TextCraft ist als Office-Add-in konzipiert und kann, sofern kompatibel, auch in Word Online verwendet werden. Die Anleitung hier bezieht sich auf die Desktop-Version Word 2013. Für Word Online gelten teilweise andere Installationsschritte.

---

**F: Was passiert mit meinen Daten, wenn ich TextCraft lösche?**

A: Nach der Deinstallation werden lokal keine weiteren Daten übertragen. Für Informationen über bereits übermittelte Daten und deren Löschung wenden Sie sich an den Anbieter von TextCraft gemäß seiner Datenschutzrichtlinie.

---

*Zuletzt aktualisiert: März 2026*
