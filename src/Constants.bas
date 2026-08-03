Attribute VB_Name = "Constants"
'===============================================================================
' Module    : Constants
' Purpose   : Central location for all magic numbers, keys and enums used
'             across the add-in. Nothing outside this module should hard-code
'             a registry path, a name string or a colour literal.
' Author    : Excel-Highlighter contributors
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
Public Const KEY_RECENT_COLOURS     As String = "RecentColours"
Public Const KEY_ALLOW_PROTECTED    As String = "AllowProtected"
Public Const KEY_HIGHLIGHT_STYLE    As String = "HighlightStyle"
Public Const KEY_INTERSECTION       As String = "IntersectionEnabled"
Public Const KEY_INTERSECTION_RGB   As String = "IntersectionRGB"
Public Const KEY_ANIMATED           As String = "AnimatedEnabled"
Public Const KEY_SELECTION_HISTORY  As String = "SelectionHistory"
Public Const KEY_PROFILES           As String = "Profiles"
Public Const KEY_ACTIVE_PROFILE     As String = "ActiveProfile"
Public Const KEY_DARK_MODE          As String = "DarkMode"
Public Const KEY_HOTKEY_TOGGLE      As String = "HotkeyToggle"
Public Const KEY_HOTKEY_HISTORY_BACK As String = "HotkeyHistoryBack"
Public Const KEY_HOTKEY_HISTORY_FWD As String = "HotkeyHistoryFwd"
Public Const KEY_PER_MODE_COLOURS   As String = "PerModeColours"
Public Const KEY_ROW_COLOUR         As String = "RowColour"
Public Const KEY_COL_COLOUR         As String = "ColColour"
Public Const KEY_SCOPE_ALL          As String = "ScopeAll"

' --- Defined name prefixes (per-workbook, hidden) ------------------------------
' Each monitored workbook gets its own set of names so that multiple open
' workbooks can be highlighted independently and simultaneously.
'
' IMPORTANT (Excel 2024 / recent 365): the names must NOT start with an
' underscore. Empirically verified on Excel 2024 build 16.0.20228: creating a
' name like "_XLCH_Row" throws "The syntax of this name isn't correct"
' (error 1004) - the modern defined-name parser rejects leading underscores
' combined with these tokens - while the clean "XLCH_Row" form succeeds. The
' whole add-in silently stopped highlighting on such builds because every
' EnsureNamesExist call died on the first name. The clean prefix is correct
' on every Excel version (there is no requirement for a leading underscore).
Public Const NAME_ROW_PREFIX        As String = "XLCH_Row"
Public Const NAME_ROW_END_PREFIX    As String = "XLCH_RowEnd"
Public Const NAME_COL_PREFIX        As String = "XLCH_Col"
Public Const NAME_COL_END_PREFIX    As String = "XLCH_ColEnd"
Public Const NAME_EXCLUDED          As String = "XLCH_Excluded"
' NOTE: NAME_EXCLUDED is a workbook-scoped defined name. When present and
' set to 1, the workbook is excluded from highlighting entirely. This lets
' the exclusion travel with the file itself rather than living in registry.
Public Const NAME_SHEET_EXCLUDED    As String = "XLCH_SheetExcluded"
' NOTE: NAME_SHEET_EXCLUDED is a worksheet-scoped defined name. When present
' and set to 1 on a specific worksheet, that sheet is excluded from
' highlighting while the rest of the workbook continues to work normally.

' --- Legacy name prefixes (pre-2.1.4, kept for migration) ----------------------
' Prior versions used a leading-underscore prefix (_XLCH_*) that modern Excel
' rejects with error 1004. Workbooks touched by an older version may still
' carry these names. We keep them read as a fallback (so old exclusions keep
' working) and sweep them during cleanup (so upgrading never orphans them).
Public Const NAME_LEGACY_ROW_PREFIX      As String = "_XLCH_Row"
Public Const NAME_LEGACY_ROW_END_PREFIX  As String = "_XLCH_RowEnd"
Public Const NAME_LEGACY_COL_PREFIX      As String = "_XLCH_Col"
Public Const NAME_LEGACY_COL_END_PREFIX  As String = "_XLCH_ColEnd"
Public Const NAME_LEGACY_EXCLUDED        As String = "_XLCH_Excluded"
Public Const NAME_LEGACY_SHEET_EXCLUDED  As String = "_XLCH_SheetExcluded"

' --- Highlight modes ------------------------------------------------------------
Public Enum HighlightMode
    hmNone = 0
    hmRow = 1
    hmColumn = 2
    hmCrosshair = 3
    hmCell = 4
End Enum

' --- Highlight style (fill vs border) -------------------------------------------
Public Enum HighlightStyle
    hsFill = 0
    hsBorder = 1
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
' NOTE: Excel's Long colour format is 0x00BBGGRR (blue in the high byte,
' red in the low byte), so the value for a colour is NOT the "RRGGBB" hex
' you might expect. RGB(255,255,0) is &H00FFFF = 65535, NOT &HFFFF00 (which
' would render as cyan). These constants were once written as if they were
' RRGGBB, which is why "Yellow" rendered cyan and "Blue" rendered violet.
Public Const RGB_YELLOW              As Long = 65535       ' RGB(255,255,0)
Public Const RGB_GREEN               As Long = 5296274    ' RGB(146, 208, 80)
Public Const RGB_ORANGE              As Long = 39423       ' RGB(255, 153, 0)
Public Const RGB_CYAN                As Long = 16777164    ' RGB(204,255,255)
Public Const RGB_BLUE                As Long = 16764057    ' RGB(153,204,255)
Public Const RGB_PINK                As Long = 13421823    ' RGB(255,204,204)
Public Const RGB_GREY                As Long = 13882323    ' RGB(211,211,211)

' --- Ribbon control ids (must match customUI14.xml) -----------------------------
Public Const CTRL_TOGGLE             As String = "btnToggleHighlight"
Public Const CTRL_MODE_ROW           As String = "btnModeRow"
Public Const CTRL_MODE_COLUMN        As String = "btnModeColumn"
Public Const CTRL_MODE_CROSSHAIR     As String = "btnModeCrosshair"
Public Const CTRL_MODE_CELL          As String = "btnModeCell"
Public Const CTRL_GALLERY_COLOUR     As String = "galColour"
Public Const CTRL_CUSTOM_COLOUR      As String = "btnCustomColour"
Public Const CTRL_EXCLUDE_WB         As String = "btnExcludeWorkbook"
Public Const CTRL_EXCLUDE_SHEET      As String = "btnExcludeSheet"
Public Const CTRL_ALLOW_PROTECTED    As String = "btnAllowProtected"
Public Const CTRL_STYLE_FILL         As String = "btnStyleFill"
Public Const CTRL_STYLE_BORDER       As String = "btnStyleBorder"
Public Const CTRL_INTERSECTION       As String = "btnIntersection"
Public Const CTRL_ANIMATED           As String = "btnAnimated"
Public Const CTRL_DARK_MODE          As String = "btnDarkMode"
Public Const CTRL_HISTORY_BACK       As String = "btnHistoryBack"
Public Const CTRL_HISTORY_FWD        As String = "btnHistoryForward"
Public Const CTRL_PROFILES           As String = "ddlProfiles"
Public Const CTRL_SAVE_PROFILE       As String = "btnSaveProfile"
Public Const CTRL_PER_MODE_COLOURS  As String = "btnPerModeColours"
Public Const CTRL_ROW_COLOUR_GAL    As String = "galRowColour"
Public Const CTRL_COL_COLOUR_GAL    As String = "galColColour"
Public Const CTRL_RESET              As String = "btnResetSettings"
Public Const CTRL_SCOPE_ALL          As String = "btnScopeAll"
Public Const CTRL_ABOUT              As String = "btnAbout"

' --- Default hotkey strings (can be overridden via registry) -------------------
Public Const DEFAULT_HOTKEY_TOGGLE       As String = "^+H"   ' Ctrl+Shift+H
Public Const DEFAULT_HOTKEY_HISTORY_BACK As String = "^+Z"   ' Ctrl+Shift+Z
Public Const DEFAULT_HOTKEY_HISTORY_FWD  As String = "^+X"   ' Ctrl+Shift+X

' --- Misc ------------------------------------------------------------------------
Public Const APP_VERSION             As String = "2.3.1"
Public Const RECENT_COLOURS_COUNT    As Long = 4
Public Const SELECTION_HISTORY_SIZE  As Long = 20
Public Const STATUS_BAR_PREFIX       As String = "XLCH: "
' NOTE: the constants above must stay in sync with any workbook that already
' has the old "_XLCH_*" names from a previous build - SafeDeleteName uses
' these constants, so old names on existing workbooks are simply left alone
' (harmless) and fresh names get created on the next selection change.
