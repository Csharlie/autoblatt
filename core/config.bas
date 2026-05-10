Attribute VB_Name = "config"
Option Explicit

' ============================================================
' AUTOBLATT - KOZPONTI KONFIGURACIO
' ============================================================
' Minden hardkodolt cella, lap, shape es alapertelmezett ut
' itt van. A modulok ezeken a konstansokon keresztul hivatkoznak,
' igy egy helyen lehet modositani.
' ============================================================

' --- Alkalmazas metaadatok ---
Public Const APP_NAME           As String = "Autoblatt"
Public Const APP_VERSION        As String = "1.0.0"
Public Const APP_PREFIX         As String = "AB_"

' --- Cellanevek (Fotoblatt munkalapon) ---
Public Const CELL_USERNAME      As String = "A2"
Public Const CELL_FULLNAME      As String = "B3"
Public Const CELL_SHIFT         As String = "F3"
Public Const CELL_DATE          As String = "I3"
Public Const CELL_SUBJECT_1     As String = "A6"
Public Const CELL_SUBJECT_2     As String = "C6"
Public Const CELL_SUBJECT_3     As String = "E6"
Public Const CELL_SUBJECT_4     As String = "G8"
Public Const CELL_HIGHLIGHT     As String = "H6"
Public Const CELL_SHEETTYPE     As String = "B1"
Public Const CELL_NOTE          As String = "C15"
Public Const CELL_FOOTER        As String = "A9"

' --- Lapnevek ---
Public Const SHEET_FOTOBLATT    As String = "Fotoblatt"
Public Const SHEET_PERSONAL     As String = "Personal"
Public Const SHEET_DATA         As String = "Data"
Public Const SHEET_FILLPANEL    As String = "Fill Panel"
Public Const SHEET_HELPRE       As String = "Helpre Panel"

' --- ListObject (Excel Tablazat) nevek ---
Public Const TABLE_PERSONAL     As String = "tablePersonal"
Public Const COL_PERSONAL_ID    As String = "ID"
Public Const COL_PERSONAL_NAME  As String = "Name"
Public Const COL_PERSONAL_SHIFT As String = "Shift"

' --- Shape-nevek (kep import) ---
Public Const SHAPE_KEEP         As String = "Grafik 5"
Public Const IMAGE_START_CELL   As String = "A17"

' --- Mentesi mappastruktura ---
' A mentett fajlok a OneDrive\Sheets\<ev>\<mappa> alatt landolnak.
' (Megegyezo a regi autosheets viselkedessel.)
Public Const SAVE_ROOT_FOLDER   As String = "Sheets"
' Az evszamot futasidoben a Year(Date) adja, nem hardkodoljuk.

' --- Csere-szabalyok normalizalashoz ---
' Format: "kereses=csere;kereses=csere;..."
' A SaveToDrive utan minden lap minden cellajan lefut.
Public Const REPLACE_RULES      As String = "bpl=BPL;sw=SW;li=Li"

' --- Email tematika ---
' "outlook" = Outlook alapertelmezett alairas hasznalata
' "custom"  = Sajat alairasszoveg (lasd email-sender modul)
' "none"    = Nincs alairas
Public Const EMAIL_SIGNATURE_MODE As String = "outlook"

' --- Fallback utvonalak (system-config.json felulir) ---
' Ezek csak akkor szamitanak, ha a JSON config nem letezik vagy hianyos.
Public Const FALLBACK_DATA_FILENAME As String = "data.xlsm"

' --- Logolas ---
Public Const LOG_MAX_BYTES      As Long = 5242880 ' 5 MB

' --- Modulok kulcsai (a loader ezekkel azonositja oket) ---
Public Const MODULE_KEY_DRIVE   As String = "drive"
Public Const MODULE_KEY_IMAGE   As String = "image"
Public Const MODULE_KEY_EMAIL   As String = "email"
Public Const MODULE_KEY_FILL    As String = "fill"
Public Const MODULE_KEY_HELPRE  As String = "helpre"
Public Const MODULE_KEY_SETTINGS As String = "settings"

' --- VBA Modul nevek (importalt komponensek azonositoi) ---
Public Const VBMODULE_DRIVE     As String = "driveManager"
Public Const VBMODULE_IMAGE     As String = "imageImporter"
Public Const VBMODULE_EMAIL     As String = "emailSender"
Public Const VBMODULE_FILL      As String = "fillPanel"
Public Const VBMODULE_HELPRE    As String = "helprePanel"
