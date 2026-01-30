MsgMetaWDX (WDX-Plugin, x64)
=============================

Funktion: Liefert Spalten für MSG-Dateien (From, To, Subject, Sent, Received, Message-ID, HasAttachments).
Voraussetzung: Outlook/Extended MAPI installiert.

**Schnellstart (empfohlen):**
1) Visual Studio 2019/2022 (Desktop C++) installieren.
2) In diesem Verzeichnis `vs\MsgMetaWDX.sln` öffnen.
3) Konfiguration `Release | x64` bauen – Ausgabe: `vs\x64\Release\MsgMetaWDX.wdx`.
4) `build_package.ps1` mit PowerShell ausführen. Das Script erstellt `MsgMetaWDX_v1.0_x64.zip` mit Auto-Install.
5) ZIP in Total Commander öffnen → "Installieren" bestätigen.

**Manuelle Installation:**
- Alternativ `MsgMetaWDX.wdx` direkt in TC unter *Einstellungen → Plugins → Inhaltsplugins (WDX) → Hinzufügen…* eintragen.

**Felder:**
- From (Name <E-Mail>)
- To
- Subject
- Sent (PR_CLIENT_SUBMIT_TIME)
- Received (PR_MESSAGE_DELIVERY_TIME)
- Message-ID (PR_INTERNET_MESSAGE_ID)
- HasAttachments (PR_HASATTACH)

**Hinweise:**
- Unicode, x64. Für x86 bitte Projekt kopieren und auf Win32 anpassen.
- Signierte/verschlüsselte Nachrichten können Anhänge kapseln; das Feld zeigt nur, ob *Anhänge vorhanden* sind.
