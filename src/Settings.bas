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

'-------------------------------------------------------------------------------
' EnsureLoaded
' Loads persisted values from the registry exactly once per Excel session.
'-------------------------------------------------------------------------------
Private Sub EnsureLoaded()

    If mLoaded Then Exit Sub

    On Error GoTo Fail

    mEnabled   = CBool(GetSetting(APP_NAME, SECTION_GENERAL, KEY_ENABLED, "False"))
    mMode      = ModeFromString(GetSetting(APP_NAME, SECTION_GENERAL, KEY_MODE, "CROSSHAIR"))
    mColour    = ColourFromString(GetSetting(APP_NAME, SECTION_GENERAL, KEY_COLOUR_NAME, "YELLOW"))
    mCustomRGB = CLng(GetSetting(APP_NAME, SECTION_GENERAL, KEY_CUSTOM_RGB, CStr(RGB_YELLOW)))

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
Public Property Get Enabled() As Boolean
    EnsureLoaded
    Enabled = mEnabled
End Property

Public Property Let Enabled(ByVal value As Boolean)
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
Public Property Get Colour() As HighlightColour
    EnsureLoaded
    Colour = mColour
End Property

Public Property Let Colour(ByVal value As HighlightColour)
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
    Logging.LogInfo "Settings.CustomRGB", "Set to " & value
End Property

'-------------------------------------------------------------------------------
' EffectiveRGB
' Convenience wrapper combining Colour + CustomRGB - this is what
' HighlightEngine actually asks for.
'-------------------------------------------------------------------------------
Public Function EffectiveRGB() As Long
    EffectiveRGB = ColourToRGB(Colour, CustomRGB)
End Function

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
    Enabled = False
    Mode = hmCrosshair
    Colour = hcYellow
    CustomRGB = RGB_YELLOW
    Logging.LogInfo "Settings.ResetToDefaults", "Settings reset to factory defaults"
End Sub
