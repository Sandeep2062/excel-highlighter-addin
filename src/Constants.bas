Attribute VB_Name = "Constants"
'===============================================================================
' Module    : Constants
' Purpose   : Central location for all magic numbers, keys and enums used
'             across the add-in. Nothing outside this module should hard-code
'             a registry path, a name string or a colour literal.
' Author    : Excel Crosshair Highlighter contributors
'===============================================================================
Option Explicit

' --- Registry / SaveSetting section identifiers -------------------------------
' SaveSetting/GetSetting write under
' HKCU\Software\VB and VBA Program Settings\<APP_NAME>\<SECTION>
Public Const APP_NAME               As String = "ExcelCrosshairHighlighter"
Public Const SECTION_GENERAL        As String = "General"

Public Const KEY_ENABLED            As String = "Enabled"
Public Const KEY_MODE               As String = "Mode"
Public Const KEY_COLOUR_NAME        As String = "ColourName"
Public Const KEY_CUSTOM_RGB         As String = "CustomRGB"

' --- Defined name prefixes (per-workbook, hidden) ------------------------------
' Each monitored workbook gets its own pair of names so that multiple open
' workbooks can be highlighted independently and simultaneously.
Public Const NAME_ROW_PREFIX        As String = "_XLCH_Row"
Public Const NAME_COL_PREFIX        As String = "_XLCH_Col"

' A marker written into a custom document property so we can recognise
' worksheets/workbooks that already have our conditional formatting applied,
' and so cleanup can find every rule that belongs to us.
Public Const CF_TAG                 As String = "XLCH"

' --- Highlight modes ------------------------------------------------------------
Public Enum HighlightMode
    hmNone = 0
    hmRow = 1
    hmColumn = 2
    hmCrosshair = 3
End Enum

' --- Named colour palette -------------------------------------------------------
Public Enum HighlightColour
    hcYellow = 0
    hcGreen = 1
    hcOrange = 2
    hcCyan = 3
    hcBlue = 4
    hcPink = 5
    hcGrey = 6
    hcCustom = 7
End Enum

' RGB literals for the fixed palette. Kept as constants rather than inline
' magic numbers so the gallery, the CF builder and the About box all agree.
Public Const RGB_YELLOW              As Long = 16776960   ' RGB(255,255,0)
Public Const RGB_GREEN               As Long = 5296274    ' RGB(146, 208, 80)
Public Const RGB_ORANGE              As Long = 39423       ' RGB(255, 153, 0)
Public Const RGB_CYAN                As Long = 16777164    ' RGB(204,255,255)
Public Const RGB_BLUE                As Long = 16751001    ' RGB(153,204,255)
Public Const RGB_PINK                As Long = 13353215    ' RGB(255,204,204)
Public Const RGB_GREY                As Long = 13882323    ' RGB(211,211,211)

' --- Ribbon control ids (must match customUI14.xml) -----------------------------
Public Const CTRL_TOGGLE             As String = "btnToggleHighlight"
Public Const CTRL_MODE_ROW           As String = "btnModeRow"
Public Const CTRL_MODE_COLUMN        As String = "btnModeColumn"
Public Const CTRL_MODE_CROSSHAIR     As String = "btnModeCrosshair"
Public Const CTRL_GALLERY_COLOUR     As String = "galColour"
Public Const CTRL_CUSTOM_COLOUR      As String = "btnCustomColour"
Public Const CTRL_RESET              As String = "btnResetSettings"
Public Const CTRL_ABOUT              As String = "btnAbout"

' --- Misc ------------------------------------------------------------------------
Public Const APP_VERSION             As String = "1.0.0"
