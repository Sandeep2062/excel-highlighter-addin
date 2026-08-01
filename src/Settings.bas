Attribute VB_Name = "Settings"
'===============================================================================
' Module    : Settings
' Purpose   : Single source of truth for user preferences. Backed by
'             SaveSetting/GetSetting (registry, HKCU) so preferences are
'             global to the user and survive across Excel sessions without
'             needing any particular workbook to be open.
'
'             All reads happen once into module-level cache variables on
'             first access (lazy init) - subsequent reads are just field
'             access, which matters because ribbon getPressed/getLabel
'             callbacks fire very frequently.
'===============================================================================
Option Explicit

Private mLoaded        As Boolean
Private mEnabled       As Boolean
Private mMode          As HighlightMode
Private mColour        As HighlightColour
Private mCustomRGB     As Long
Private mRecentColours() As Long
Private mRecentCount   As Long
Private mAllowProtected As Boolean
Private mHighlightStyle As HighlightStyle
Private mIntersection   As Boolean
Private mIntersectionRGB As Long
Private mAnimated       As Boolean
Private mDarkMode       As Boolean
Private mHotkeyToggle   As String
Private mHotkeyHistoryBack As String
Private mHotkeyHistoryFwd  As String
Private mPerModeColours  As Boolean
Private mRowColour       As HighlightColour
Private mColColour       As HighlightColour

' NOTE: Recent colours are stored as a comma-separated list in the registry
' under KEY_RECENT_COLOURS. The array is rebuilt on load and flushed on write.
' Selection history is stored as a comma-separated list of "row,col" pairs
' under KEY_SELECTION_HISTORY. Profiles are stored as pipe-delimited under
' KEY_PROFILES.

'-------------------------------------------------------------------------------
' EnsureLoaded
' Loads persisted values from the registry exactly once per Excel session.
'-------------------------------------------------------------------------------
Private Sub EnsureLoaded()

    If mLoaded Then Exit Sub

    On Error GoTo Fail

    mEnabled = CBool(GetSetting(APP_NAME, SECTION_GENERAL, KEY_ENABLED, "False"))
    mMode = ModeFromString(GetSetting(APP_NAME, SECTION_GENERAL, KEY_MODE, "CROSSHAIR"))
    mColour = ColourFromString(GetSetting(APP_NAME, SECTION_GENERAL, KEY_COLOUR_NAME, "YELLOW"))
    mCustomRGB = CLng(GetSetting(APP_NAME, SECTION_GENERAL, KEY_CUSTOM_RGB, CStr(RGB_YELLOW)))
    mAllowProtected = CBool(GetSetting(APP_NAME, SECTION_GENERAL, KEY_ALLOW_PROTECTED, "False"))
    mHighlightStyle = CInt(GetSetting(APP_NAME, SECTION_GENERAL, KEY_HIGHLIGHT_STYLE, "0"))
    mIntersection = CBool(GetSetting(APP_NAME, SECTION_GENERAL, KEY_INTERSECTION, "False"))
    mIntersectionRGB = CLng(GetSetting(APP_NAME, SECTION_GENERAL, KEY_INTERSECTION_RGB, CStr(RGB_YELLOW)))
    mAnimated = CBool(GetSetting(APP_NAME, SECTION_GENERAL, KEY_ANIMATED, "False"))
    mDarkMode = CBool(GetSetting(APP_NAME, SECTION_GENERAL, KEY_DARK_MODE, "False"))

    ' Configurable hotkeys - fall back to defaults if not set in registry.
    mHotkeyToggle = GetSetting(APP_NAME, SECTION_GENERAL, KEY_HOTKEY_TOGGLE, DEFAULT_HOTKEY_TOGGLE)
    mHotkeyHistoryBack = GetSetting(APP_NAME, SECTION_GENERAL, KEY_HOTKEY_HISTORY_BACK, DEFAULT_HOTKEY_HISTORY_BACK)
    mHotkeyHistoryFwd = GetSetting(APP_NAME, SECTION_GENERAL, KEY_HOTKEY_HISTORY_FWD, DEFAULT_HOTKEY_HISTORY_FWD)

    ' Per-mode colours for Crosshair mode (row vs column can differ).
    mPerModeColours = CBool(GetSetting(APP_NAME, SECTION_GENERAL, KEY_PER_MODE_COLOURS, "False"))
    mRowColour = ColourFromString(GetSetting(APP_NAME, SECTION_GENERAL, KEY_ROW_COLOUR, "YELLOW"))
    mColColour = ColourFromString(GetSetting(APP_NAME, SECTION_GENERAL, KEY_COL_COLOUR, "YELLOW"))

    LoadRecentColours

    mLoaded = True
    Exit Sub

Fail:
    ' Corrupt or missing registry values - fall back to safe defaults rather
    ' than propagating the error into Workbook_Open.
    Logging.LogError "Settings.EnsureLoaded", Err.Number, Err.Description
    mEnabled = False
    mMode = hmCrosshair
    mColour = hcYellow
    mCustomRGB = RGB_YELLOW
    mLoaded = True

End Sub

'-------------------------------------------------------------------------------
' Enabled - Get/Let
'-------------------------------------------------------------------------------
Public Property Get enabled() As Boolean
    EnsureLoaded
    enabled = mEnabled
End Property

Public Property Let enabled(ByVal value As Boolean)
    EnsureLoaded
    mEnabled = value
    SaveSetting APP_NAME, SECTION_GENERAL, KEY_ENABLED, CStr(value)
    Logging.LogInfo "Settings.Enabled", "Set to " & value
End Property

'-------------------------------------------------------------------------------
' Mode - Get/Let
'-------------------------------------------------------------------------------
Public Property Get Mode() As HighlightMode
    EnsureLoaded
    Mode = mMode
End Property

Public Property Let Mode(ByVal value As HighlightMode)
    EnsureLoaded
    mMode = value
    SaveSetting APP_NAME, SECTION_GENERAL, KEY_MODE, ModeToString(value)
    Logging.LogInfo "Settings.Mode", "Set to " & ModeToString(value)
End Property

'-------------------------------------------------------------------------------
' Colour - Get/Let
'-------------------------------------------------------------------------------
Public Property Get colour() As HighlightColour
    EnsureLoaded
    colour = mColour
End Property

Public Property Let colour(ByVal value As HighlightColour)
    EnsureLoaded
    mColour = value
    SaveSetting APP_NAME, SECTION_GENERAL, KEY_COLOUR_NAME, ColourName(value)
    Logging.LogInfo "Settings.Colour", "Set to " & ColourName(value)
End Property

'-------------------------------------------------------------------------------
' CustomRGB - Get/Let
'-------------------------------------------------------------------------------
Public Property Get CustomRGB() As Long
    EnsureLoaded
    CustomRGB = mCustomRGB
End Property

Public Property Let CustomRGB(ByVal value As Long)
    EnsureLoaded
    mCustomRGB = value
    SaveSetting APP_NAME, SECTION_GENERAL, KEY_CUSTOM_RGB, CStr(value)
    PushRecentColour value
    Logging.LogInfo "Settings.CustomRGB", "Set to " & value
End Property

'-------------------------------------------------------------------------------
' RecentColours - Get
' Returns a copy of the recent colours array (up to RECENT_COLOURS_COUNT items).
'-------------------------------------------------------------------------------
Public Property Get RecentColours() As Long()
    EnsureLoaded
    RecentColours = mRecentColours
End Property

Public Property Get RecentColourCount() As Long
    EnsureLoaded
    RecentColourCount = mRecentCount
End Property

'-------------------------------------------------------------------------------
' RecentColour - indexed accessor
' Returns the RGB value at the given 0-based index, or 0 if out of range.
' Use this instead of RecentColours()(index) - VBA cannot index directly into
' a Property Get that returns an array.
'-------------------------------------------------------------------------------
Public Function RecentColour(ByVal index As Long) As Long
    EnsureLoaded
    If index >= 0 And index < mRecentCount Then
        RecentColour = mRecentColours(index)
    Else
        RecentColour = 0
    End If
End Function

'-------------------------------------------------------------------------------
' EffectiveRGB
' Convenience wrapper combining Colour + CustomRGB - this is what
' HighlightEngine asks for when per-mode colours are off.
'-------------------------------------------------------------------------------
Public Function EffectiveRGB() As Long
    EffectiveRGB = ApplyDarkModeTint(ColourToRGB(colour, CustomRGB))
End Function

'-------------------------------------------------------------------------------
' EffectiveRowRGB / EffectiveColRGB
' Returns the RGB for row/column highlighting. When per-mode colours are
' enabled, each axis can have its own colour. Otherwise both fall back to
' the standard EffectiveRGB.
'-------------------------------------------------------------------------------
Public Function EffectiveRowRGB() As Long
    If PerModeColours Then
        EffectiveRowRGB = ApplyDarkModeTint(ColourToRGB(rowColour, CustomRGB))
    Else
        EffectiveRowRGB = EffectiveRGB
    End If
End Function

Public Function EffectiveColRGB() As Long
    If PerModeColours Then
        EffectiveColRGB = ApplyDarkModeTint(ColourToRGB(colColour, CustomRGB))
    Else
        EffectiveColRGB = EffectiveRGB
    End If
End Function

'-------------------------------------------------------------------------------
' ApplyDarkModeTint
' Shared tinting logic. Softens vibrant colours for dark Office themes so
' the highlight doesn't glare against a dark background.
'-------------------------------------------------------------------------------
Public Function ApplyDarkModeTint(ByVal baseRGB As Long) As Long
    If DarkMode Then
        Dim r As Long, g As Long, b As Long
        r = (baseRGB \ 65536) Mod 256
        g = (baseRGB \ 256) Mod 256
        b = baseRGB Mod 256
        r = (r * 65 + 35 * 30) \ 100
        g = (g * 65 + 35 * 30) \ 100
        b = (b * 65 + 35 * 30) \ 100
        ApplyDarkModeTint = RGB(r, g, b)
    Else
        ApplyDarkModeTint = baseRGB
    End If
End Function

'-------------------------------------------------------------------------------
' AllowProtected - Get/Let
' When True, allows highlighting on protected sheets by temporarily unprotecting
' them. The user must explicitly opt in - we never change protection settings
' without consent.
'-------------------------------------------------------------------------------
Public Property Get AllowProtected() As Boolean
    EnsureLoaded
    AllowProtected = mAllowProtected
End Property

Public Property Let AllowProtected(ByVal value As Boolean)
    EnsureLoaded
    mAllowProtected = value
    SaveSetting APP_NAME, SECTION_GENERAL, KEY_ALLOW_PROTECTED, CStr(value)
    Logging.LogInfo "Settings.AllowProtected", "Set to " & value
End Property

'-------------------------------------------------------------------------------
' HighlightStyle - Get/Let (Fill vs Border)
'-------------------------------------------------------------------------------
Public Property Get HighlightStyle() As HighlightStyle
    EnsureLoaded
    HighlightStyle = mHighlightStyle
End Property

Public Property Let HighlightStyle(ByVal value As HighlightStyle)
    EnsureLoaded
    mHighlightStyle = value
    SaveSetting APP_NAME, SECTION_GENERAL, KEY_HIGHLIGHT_STYLE, CStr(CInt(value))
    Logging.LogInfo "Settings.HighlightStyle", "Set to " & CInt(value)
End Property

'-------------------------------------------------------------------------------
' Intersection - Get/Let
' When enabled, the cell at the row/column intersection gets a stronger tint.
'-------------------------------------------------------------------------------
Public Property Get IntersectionEnabled() As Boolean
    EnsureLoaded
    IntersectionEnabled = mIntersection
End Property

Public Property Let IntersectionEnabled(ByVal value As Boolean)
    EnsureLoaded
    mIntersection = value
    SaveSetting APP_NAME, SECTION_GENERAL, KEY_INTERSECTION, CStr(value)
    Logging.LogInfo "Settings.IntersectionEnabled", "Set to " & value
End Property

Public Property Get IntersectionRGB() As Long
    EnsureLoaded
    IntersectionRGB = mIntersectionRGB
End Property

Public Property Let IntersectionRGB(ByVal value As Long)
    EnsureLoaded
    mIntersectionRGB = value
    SaveSetting APP_NAME, SECTION_GENERAL, KEY_INTERSECTION_RGB, CStr(value)
    Logging.LogInfo "Settings.IntersectionRGB", "Set to " & value
End Property

'-------------------------------------------------------------------------------
' Animated - Get/Let
' When enabled, the highlight briefly pulses when the selection changes.
'-------------------------------------------------------------------------------
Public Property Get AnimatedEnabled() As Boolean
    EnsureLoaded
    AnimatedEnabled = mAnimated
End Property

Public Property Let AnimatedEnabled(ByVal value As Boolean)
    EnsureLoaded
    mAnimated = value
    SaveSetting APP_NAME, SECTION_GENERAL, KEY_ANIMATED, CStr(value)
    Logging.LogInfo "Settings.Animated", "Set to " & value
End Property

'-------------------------------------------------------------------------------
' DarkMode - Get/Let
' When enabled, adjusts palette colours for dark Office themes.
'-------------------------------------------------------------------------------
Public Property Get DarkMode() As Boolean
    EnsureLoaded
    DarkMode = mDarkMode
End Property

Public Property Let DarkMode(ByVal value As Boolean)
    EnsureLoaded
    mDarkMode = value
    SaveSetting APP_NAME, SECTION_GENERAL, KEY_DARK_MODE, CStr(value)
    Logging.LogInfo "Settings.DarkMode", "Set to " & value
End Property

'-------------------------------------------------------------------------------
' Configurable hotkeys - Get/Let
' Users can customise the key combinations via the registry. Values use
' Application.OnKey syntax: ^ = Ctrl, + = Shift, % = Alt.
'-------------------------------------------------------------------------------
Public Property Get HotkeyToggle() As String
    EnsureLoaded
    HotkeyToggle = mHotkeyToggle
End Property

Public Property Let HotkeyToggle(ByVal value As String)
    EnsureLoaded
    mHotkeyToggle = value
    SaveSetting APP_NAME, SECTION_GENERAL, KEY_HOTKEY_TOGGLE, value
    Logging.LogInfo "Settings.HotkeyToggle", "Set to " & value
End Property

Public Property Get HotkeyHistoryBack() As String
    EnsureLoaded
    HotkeyHistoryBack = mHotkeyHistoryBack
End Property

Public Property Let HotkeyHistoryBack(ByVal value As String)
    EnsureLoaded
    mHotkeyHistoryBack = value
    SaveSetting APP_NAME, SECTION_GENERAL, KEY_HOTKEY_HISTORY_BACK, value
    Logging.LogInfo "Settings.HotkeyHistoryBack", "Set to " & value
End Property

Public Property Get HotkeyHistoryFwd() As String
    EnsureLoaded
    HotkeyHistoryFwd = mHotkeyHistoryFwd
End Property

Public Property Let HotkeyHistoryFwd(ByVal value As String)
    EnsureLoaded
    mHotkeyHistoryFwd = value
    SaveSetting APP_NAME, SECTION_GENERAL, KEY_HOTKEY_HISTORY_FWD, value
    Logging.LogInfo "Settings.HotkeyHistoryFwd", "Set to " & value
End Property

'-------------------------------------------------------------------------------
' Per-mode colours - for Crosshair mode, allows distinct row vs column colours.
'-------------------------------------------------------------------------------
Public Property Get PerModeColours() As Boolean
    EnsureLoaded
    PerModeColours = mPerModeColours
End Property

Public Property Let PerModeColours(ByVal value As Boolean)
    EnsureLoaded
    mPerModeColours = value
    SaveSetting APP_NAME, SECTION_GENERAL, KEY_PER_MODE_COLOURS, CStr(value)
    Logging.LogInfo "Settings.PerModeColours", "Set to " & value
End Property

Public Property Get rowColour() As HighlightColour
    EnsureLoaded
    rowColour = mRowColour
End Property

Public Property Let rowColour(ByVal value As HighlightColour)
    EnsureLoaded
    mRowColour = value
    SaveSetting APP_NAME, SECTION_GENERAL, KEY_ROW_COLOUR, ColourName(value)
    Logging.LogInfo "Settings.RowColour", "Set to " & ColourName(value)
End Property

Public Property Get colColour() As HighlightColour
    EnsureLoaded
    colColour = mColColour
End Property

Public Property Let colColour(ByVal value As HighlightColour)
    EnsureLoaded
    mColColour = value
    SaveSetting APP_NAME, SECTION_GENERAL, KEY_COL_COLOUR, ColourName(value)
    Logging.LogInfo "Settings.ColColour", "Set to " & ColourName(value)
End Property

'-------------------------------------------------------------------------------
' RecentColour helpers
'-------------------------------------------------------------------------------
Private Sub LoadRecentColours()
    On Error Resume Next
    Dim raw As String
    raw = GetSetting(APP_NAME, SECTION_GENERAL, KEY_RECENT_COLOURS, "")
    If Len(raw) = 0 Then
        mRecentCount = 0
        ReDim mRecentColours(0 To RECENT_COLOURS_COUNT - 1)
        Exit Sub
    End If
    Dim parts() As String
    parts = Split(raw, ",")
    Dim i As Long
    mRecentCount = 0
    ReDim mRecentColours(0 To RECENT_COLOURS_COUNT - 1)
    For i = 0 To UBound(parts)
        If mRecentCount >= RECENT_COLOURS_COUNT Then Exit For
        Dim val As Long
        val = CLng(Trim$(parts(i)))
        If val <> 0 Then
            mRecentColours(mRecentCount) = val
            mRecentCount = mRecentCount + 1
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub PushRecentColour(ByVal rgbVal As Long)
    ' Move this colour to the front of the list, shifting others right.
    ' Duplicates are removed first so the list stays unique.
    Dim i As Long
    Dim j As Long
    ' Remove existing occurrence
    For i = 0 To mRecentCount - 1
        If mRecentColours(i) = rgbVal Then
            For j = i To mRecentCount - 2
                mRecentColours(j) = mRecentColours(j + 1)
            Next j
            mRecentCount = mRecentCount - 1
            Exit For
        End If
    Next i
    ' Shift right
    If mRecentCount >= RECENT_COLOURS_COUNT Then
        mRecentCount = RECENT_COLOURS_COUNT - 1
    End If
    For i = mRecentCount - 1 To 0 Step -1
        mRecentColours(i + 1) = mRecentColours(i)
    Next i
    mRecentColours(0) = rgbVal
    If mRecentCount < RECENT_COLOURS_COUNT Then
        mRecentCount = mRecentCount + 1
    End If
    ' Persist
    Dim parts As String
    parts = ""
    For i = 0 To mRecentCount - 1
        If Len(parts) > 0 Then parts = parts & ","
        parts = parts & mRecentColours(i)
    Next i
    SaveSetting APP_NAME, SECTION_GENERAL, KEY_RECENT_COLOURS, parts
End Sub

'-------------------------------------------------------------------------------
' ResetToDefaults
' Description : Restores factory defaults and persists them immediately.
'               Bound to the "Reset Settings" ribbon button.
'-------------------------------------------------------------------------------
Public Sub ResetToDefaults()

    On Error GoTo Fail

    Me_Enabled_ResetHelper
    Exit Sub

Fail:
    Logging.LogError "Settings.ResetToDefaults", Err.Number, Err.Description

End Sub

' Split out so ResetToDefaults reads cleanly; also makes it easy to unit-test
' the value assignment separately from the (unlikely to fail) error trap above.
Private Sub Me_Enabled_ResetHelper()
    enabled = False
    Mode = hmCrosshair
    colour = hcYellow
    CustomRGB = RGB_YELLOW
    AllowProtected = False
    HighlightStyle = hsFill
    IntersectionEnabled = False
    IntersectionRGB = RGB_YELLOW
    AnimatedEnabled = False
    DarkMode = False
    HotkeyToggle = DEFAULT_HOTKEY_TOGGLE
    HotkeyHistoryBack = DEFAULT_HOTKEY_HISTORY_BACK
    HotkeyHistoryFwd = DEFAULT_HOTKEY_HISTORY_FWD
    PerModeColours = False
    rowColour = hcYellow
    colColour = hcYellow
    ' Clear recent colours on reset
    SaveSetting APP_NAME, SECTION_GENERAL, KEY_RECENT_COLOURS, ""
    mRecentCount = 0
    ReDim mRecentColours(0 To RECENT_COLOURS_COUNT - 1)
    Logging.LogInfo "Settings.ResetToDefaults", "Settings reset to factory defaults"
End Sub
