VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================================
' Module    : ThisWorkbook
' Purpose   : The add-in's own workbook object. Stands up the Application-level
'             event sink (EventApp) when Excel loads the add-in, tears it down
'             on unload, and - critically - supplies the ribbon XML to Excel by
'             implementing IRibbonExtensibility.
'
'             RIBBON DELIVERY (why Implements lives here, not on a class):
'             This module is the add-in's root IDispatch object - the same
'             root-object pattern Word uses (ThisDocument) and, empirically,
'             the pattern commercial Excel add-ins such as ASAP Utilities rely
'             on when they deliver their ribbon from VBA. GetCustomUI here is
'             a FALLBACK: the primary delivery is the declarative
'             customUI14.xml package part injected by scripts/build-xlam.ps1
'             (the default build). When the part is present Excel uses it and
'             never calls GetCustomUI, so this interface is only exercised on
'             builds that do not honour the part.
'
'             WHY THE RIBBON MAY STILL NOT APPEAR (session findings, Excel
'             2024 build 16.0.20228, see CHANGELOG 2.1.3): earlier rounds
'             blamed the signature (WRONG - an unsigned copy of ASAP Utilities
'             renders from the same folder; the signature is not the gate) and
'             the customUI part (also not the gate - even a minimal static
'             part is ignored on this build). The confirmed difference between
'             this add-in and the one that renders (ASAP Utilities) is that
'             ASAP's VBA project is PASSWORD-LOCKED (precompiled). Excel's
'             ribbon engine discovers IRibbonExtensibility from the compiled
'             type info; an unlocked source-only project is queried before
'             lazy compilation exposes the interface, so Excel silently treats
'             the add-in as having no ribbon - no tab, no onLoad, no
'             GetCustomUI, no error. Fix if needed: run
'             scripts/lock-vba-project.ps1 once (Tools > VBAProject Properties
'             > Protection > Lock project for viewing + password), which
'             re-injects the ribbon wiring afterwards.
'
'             SIGNATURE TRAP: GetCustomUI must be a Function returning String -
'             the interface declares [out, retval] BSTR*, which VBA projects as
'             the return value. A `Sub ... ByRef RibbonXml` spelling does not
'             match the interface, fails to compile, and produces an add-in
'             that loads its VBA but never its ribbon (and, at startup, makes
'             Excel auto-open the VBE on the broken project).
'===============================================================================
Option Explicit

Implements IRibbonExtensibility

'-------------------------------------------------------------------------------
' Workbook_Open
' Fires when the xlam is loaded as an add-in (Excel loads add-ins by opening
' them, same as a normal workbook, just without showing a window for it).
'-------------------------------------------------------------------------------
Private Sub Workbook_Open()
    On Error GoTo ErrHandler
    AddinHost.StartUp
    Exit Sub
ErrHandler:
    Logging.LogError "ThisWorkbook.Workbook_Open", Err.Number, Err.Description
End Sub

'-------------------------------------------------------------------------------
' Workbook_AddinInstall
' Fires once, the first time the user ticks the add-in on in
' Excel Options > Add-ins. Workbook_Open also fires immediately after this,
' so we don't duplicate startup logic here - just log the milestone.
'-------------------------------------------------------------------------------
Private Sub Workbook_AddinInstall()
    On Error Resume Next
    Logging.LogInfo "ThisWorkbook.Workbook_AddinInstall", "Add-in installed, version " & APP_VERSION
End Sub

'-------------------------------------------------------------------------------
' Workbook_AddinUninstall
' Fires when the user unticks the add-in. We proactively strip our
' conditional formatting and names from every currently open workbook so
' nothing is left behind after the add-in disappears.
'-------------------------------------------------------------------------------
Private Sub Workbook_AddinUninstall()
    On Error GoTo ErrHandler
    AddinHost.ShutDown
    Exit Sub
ErrHandler:
    Logging.LogError "ThisWorkbook.Workbook_AddinUninstall", Err.Number, Err.Description
End Sub

'-------------------------------------------------------------------------------
' Workbook_BeforeClose
' Belt-and-braces cleanup for the (unusual) case where the add-in workbook
' itself is closed directly, e.g. via VBA automation, rather than through the
' normal add-ins UI.
'-------------------------------------------------------------------------------
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    AddinHost.ShutDown
End Sub

'-------------------------------------------------------------------------------
' IRibbonExtensibility_GetCustomUI
' Excel calls this when it needs the ribbon XML for the add-in. MUST be a
' Function returning String (see header - [out, retval] maps to the return
' value; the Sub+ByRef spelling is a compile error that silently kills the
' ribbon). Reads customUI14.xml from the add-in's own folder so the XML stays
' the single editable source of truth; falls back to a minimal one-button
' ribbon on any read failure so the tab is never silently missing.
'-------------------------------------------------------------------------------
Private Function IRibbonExtensibility_GetCustomUI(ByVal RibbonID As String) As String
    On Error GoTo ErrHandler
    Logging.LogInfo "ThisWorkbook.IRibbonExtensibility_GetCustomUI", "GetCustomUI called, RibbonID=" & RibbonID

    Dim xmlPath As String
    xmlPath = AddinHost.AddinFolder & "\customUI14.xml"

    ' Read as UTF-8 via ADODB.Stream (correctly decodes non-ASCII and strips a
    ' UTF-8 BOM; a raw binary read + StrConv would be an ANSI decode and can
    ' never strip the BOM).
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2          ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile xmlPath
    IRibbonExtensibility_GetCustomUI = stream.ReadText
    stream.Close
    Set stream = Nothing

    ' Prove the real file was read: the fallback ribbon also produces a tab,
    ' so a "tab appeared" test alone cannot distinguish the full ribbon from
    ' the one-button fallback. Logging the resolved path + read length does.
    Logging.LogInfo "ThisWorkbook.IRibbonExtensibility_GetCustomUI", _
        "Read " & Len(IRibbonExtensibility_GetCustomUI) & " chars from " & xmlPath
    Exit Function

ErrHandler:
    If Not stream Is Nothing Then
        On Error Resume Next
        stream.Close
        On Error GoTo ErrHandler
    End If
    Logging.LogError "ThisWorkbook.IRibbonExtensibility_GetCustomUI", Err.Number, Err.Description
    IRibbonExtensibility_GetCustomUI = MinimalRibbonXml()
End Function

'-------------------------------------------------------------------------------
' MinimalRibbonXml
' Safety-net ribbon when the real XML cannot be read. Keeps the tab present
' with a single About button so the user always sees the add-in has UI.
'-------------------------------------------------------------------------------
Private Function MinimalRibbonXml() As String
    MinimalRibbonXml = _
        "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & _
        "<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" " & _
        "onLoad=""RibbonCallbacks.onLoad"">" & _
        "<ribbon><tabs><tab id=""tabHighlighter"" label=""Highlighter"">" & _
        "<group id=""grpFallback"" label=""Highlighter"">" & _
        "<button id=""btnAboutFallback"" label=""About"" imageMso=""Info"" " & _
        "screentip=""About Excel Highlighter"" onAction=""RibbonCallbacks.OnAbout_Action""/>" & _
        "</group></tab></tabs></ribbon></customUI>"
End Function
