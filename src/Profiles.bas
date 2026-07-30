Attribute VB_Name = "Profiles"
'===============================================================================
' Module    : Profiles
' Purpose   : Named bundles of mode + colour + style + intersection settings
'             that can be saved and restored quickly via a dropdown.
'
'             Profiles are persisted in the registry under KEY_PROFILES as a
'             pipe-delimited string: "name|mode|colour|customRGB|style|intersect"
'             The active profile name is stored under KEY_ACTIVE_PROFILE.
'===============================================================================
Option Explicit

Private Type ProfileEntry
    Name        As String
    ModeIndex   As Integer
    ColourIndex As Integer
    CustomRGB   As Long
    StyleIndex  As Integer
    Intersect   As Boolean
End Type

Private mProfiles()  As ProfileEntry
Private mCount       As Long
Private mActiveName  As String

'-------------------------------------------------------------------------------
' Load / Save (registry)
'-------------------------------------------------------------------------------
Public Sub LoadProfiles()
    On Error Resume Next
    Dim raw As String
    raw = GetSetting(APP_NAME, SECTION_GENERAL, KEY_PROFILES, "")
    If Len(raw) = 0 Then
        ' Create a default profile.
        mCount = 1
        ReDim mProfiles(0 To 0)
        mProfiles(0).Name = "Default"
        mProfiles(0).ModeIndex = hmCrosshair
        mProfiles(0).ColourIndex = hcYellow
        mProfiles(0).CustomRGB = RGB_YELLOW
        mProfiles(0).StyleIndex = hsFill
        mProfiles(0).Intersect = False
        SaveToRegistry
    Else
        Dim lines() As String
        lines = Split(raw, "|")
        mCount = (UBound(lines) + 1) \ 6
        ReDim mProfiles(0 To mCount - 1)
        Dim i As Long, j As Long
        For i = 0 To mCount - 1
            j = i * 6
            mProfiles(i).Name = lines(j)
            mProfiles(i).ModeIndex = CInt(lines(j + 1))
            mProfiles(i).ColourIndex = CInt(lines(j + 2))
            mProfiles(i).CustomRGB = CLng(lines(j + 3))
            mProfiles(i).StyleIndex = CInt(lines(j + 4))
            mProfiles(i).Intersect = CBool(lines(j + 5))
        Next i
    End If
    mActiveName = GetSetting(APP_NAME, SECTION_GENERAL, KEY_ACTIVE_PROFILE, "Default")
End Sub

Private Sub SaveToRegistry()
    Dim parts As String
    Dim i As Long
    For i = 0 To mCount - 1
        If Len(parts) > 0 Then parts = parts & "|"
        parts = parts & mProfiles(i).Name & "|" & _
                mProfiles(i).ModeIndex & "|" & _
                mProfiles(i).ColourIndex & "|" & _
                mProfiles(i).CustomRGB & "|" & _
                mProfiles(i).StyleIndex & "|" & _
                mProfiles(i).Intersect
    Next i
    SaveSetting APP_NAME, SECTION_GENERAL, KEY_PROFILES, parts
End Sub

'-------------------------------------------------------------------------------
' Profile operations
'-------------------------------------------------------------------------------
Public Sub ApplyProfile(ByVal index As Integer)
    If index < 0 Or index >= mCount Then Exit Sub
    Dim p As ProfileEntry
    p = mProfiles(index)
    Settings.Mode = CInt(p.ModeIndex)
    Settings.Colour = CInt(p.ColourIndex)
    Settings.CustomRGB = p.CustomRGB
    Settings.HighlightStyle = CInt(p.StyleIndex)
    Settings.IntersectionEnabled = p.Intersect
    mActiveName = p.Name
    SaveSetting APP_NAME, SECTION_GENERAL, KEY_ACTIVE_PROFILE, mActiveName
    Logging.LogInfo "Profiles.ApplyProfile", "Applied profile: " & p.Name
End Sub

Public Sub SaveCurrentAsProfile(ByVal name As String)
    ' Check for duplicates.
    Dim i As Long
    For i = 0 To mCount - 1
        If LCase$(mProfiles(i).Name) = LCase$(name) Then
            mProfiles(i).ModeIndex = Settings.Mode
            mProfiles(i).ColourIndex = Settings.Colour
            mProfiles(i).CustomRGB = Settings.CustomRGB
            mProfiles(i).StyleIndex = Settings.HighlightStyle
            mProfiles(i).Intersect = Settings.IntersectionEnabled
            mActiveName = name
            SaveToRegistry
            SaveSetting APP_NAME, SECTION_GENERAL, KEY_ACTIVE_PROFILE, mActiveName
            Logging.LogInfo "Profiles.SaveCurrentAsProfile", "Updated profile: " & name
            Exit Sub
        End If
    Next i
    ' Add new profile.
    ReDim Preserve mProfiles(0 To mCount)
    mProfiles(mCount).Name = name
    mProfiles(mCount).ModeIndex = Settings.Mode
    mProfiles(mCount).ColourIndex = Settings.Colour
    mProfiles(mCount).CustomRGB = Settings.CustomRGB
    mProfiles(mCount).StyleIndex = Settings.HighlightStyle
    mProfiles(mCount).Intersect = Settings.IntersectionEnabled
    mCount = mCount + 1
    mActiveName = name
    SaveToRegistry
    SaveSetting APP_NAME, SECTION_GENERAL, KEY_ACTIVE_PROFILE, mActiveName
    Logging.LogInfo "Profiles.SaveCurrentAsProfile", "Created profile: " & name
End Sub

'-------------------------------------------------------------------------------
' Read-only access for the ribbon dropdown
'-------------------------------------------------------------------------------
Public Function ProfileCount() As Long
    ProfileCount = mCount
End Function

Public Function ProfileName(ByVal index As Long) As String
    If index >= 0 And index < mCount Then
        ProfileName = mProfiles(index).Name
    End If
End Function

Public Function ActiveProfileIndex() As Long
    Dim i As Long
    For i = 0 To mCount - 1
        If mProfiles(i).Name = mActiveName Then
            ActiveProfileIndex = i
            Exit Function
        End If
    Next i
    ActiveProfileIndex = 0
End Function

Public Sub Init()
    LoadProfiles
End Sub