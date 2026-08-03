Attribute VB_Name = "ColourPicker"
'===============================================================================
' Module    : ColourPicker
' Purpose   : Replacement for the xlDialogEditColor custom colour flow.
'             Uses the Windows Common Dialog API (ChooseColor from comdlg32)
'             instead of borrowing palette slot 56 from the active workbook.
'
'             No palette slots are touched, no workbook colours are modified.
'             The dialog shows the standard Windows colour picker with a
'             full palette, custom colour area, and "Define Custom Colors"
'             section - exactly what users expect from a native app.
'===============================================================================
Option Explicit

' --- API structures and declarations ------------------------------------------

#If VBA7 Then
    Private Type CHOOSECOLOR
        lStructSize    As Long
        hwndOwner      As LongPtr
        hInstance      As LongPtr
        rgbResult      As Long
        lpCustColors   As LongPtr
        Flags          As Long
        lCustData      As LongPtr
        lpfnHook       As LongPtr
        lpTemplateName As LongPtr
    End Type

    Private Type PICTDESC
        cbSizeOfStruct As Long
        picType        As Long
        hBitmap        As LongPtr
        hPal           As LongPtr
    End Type

    Private Declare PtrSafe Function ChooseColorAPI Lib "comdlg32.dll" Alias "ChooseColorW" (pChoosecolor As CHOOSECOLOR) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
    Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
    Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
    Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr
    Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
    Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
    Private Declare PtrSafe Function SetRect Lib "user32" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Private Declare PtrSafe Function FillRect Lib "user32" (ByVal hDC As LongPtr, ByRef lpRect As RECT, ByVal hBrush As LongPtr) As Long
    Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (ByRef pPictDesc As PICTDESC, ByRef riid As GUID, ByVal fOwn As Long, ByRef ppvObj As Object) As Long
#Else
    Private Type CHOOSECOLOR
        lStructSize    As Long
        hwndOwner      As Long
        hInstance      As Long
        rgbResult      As Long
        lpCustColors   As Long
        Flags          As Long
        lCustData      As Long
        lpfnHook       As Long
        lpTemplateName As String
    End Type

    Private Type PICTDESC
        cbSizeOfStruct As Long
        picType        As Long
        hBitmap        As Long
        hPal           As Long
    End Type

    Private Declare Function ChooseColorAPI Lib "comdlg32.dll" Alias "ChooseColorW" (pChoosecolor As CHOOSECOLOR) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
    Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
    Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
    Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
    Private Declare Function SetRect Lib "user32" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
    Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (ByRef pPictDesc As PICTDESC, ByRef riid As GUID, ByVal fOwn As Long, ByRef ppvObj As Object) As Long
#End If

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Type GUID
    Data1    As Long
    Data2    As Integer
    Data3    As Integer
    Data4(7) As Byte
End Type

' Flags for the CHOOSECOLOR API.
Private Const CC_RGBINIT       As Long = &H1
Private Const CC_FULLOPEN      As Long = &H2
Private Const CC_ANYCOLOR      As Long = &H100

' 16 custom colour slots that the API remembers between calls.
Private mCustColors(0 To 15) As Long

' Swatch cache, keyed by RGB value as text. The ribbon calls getItemImage far
' more often than the underlying colour actually changes (every repaint,
' every hover, every Invalidate) - regenerating a GDI bitmap + IPicture
' wrapper from scratch on each of those calls was needlessly expensive and,
' over a long Excel session, the most likely source of GDI resource
' exhaustion (which shows up as broken/exclamation-mark icons and eventual
' instability). Once a colour's swatch exists we just hand back the same
' IPicture reference.
Private mSwatchCache As Object

'-------------------------------------------------------------------------------
' PromptForCustomRGB
' Description : Opens the native Windows colour picker dialog. This is the
'               replacement for the old xlDialogEditColor approach which
'               borrowed palette slot 56. No workbook palettes are touched.
' Returns     : Boolean - True if the user picked a colour (OK, not Cancel)
' Parameters  : result     - ByRef Long, the chosen RGB value when True
'               initialRGB - the colour to show as initially selected
'-------------------------------------------------------------------------------
Public Function PromptForCustomRGB(ByRef result As Long, _
                                   Optional ByVal initialRGB As Long = -1) As Boolean

    On Error GoTo ErrHandler

    ' Initialise custom colours on first call.
    Static custInited As Boolean
    If Not custInited Then
        InitCustomColours
        custInited = True
    End If

    Dim cc As CHOOSECOLOR
    cc.lStructSize = LenB(cc)
    cc.hwndOwner = GetExcelHWND()
    
    If initialRGB >= 0 Then
        cc.rgbResult = initialRGB
        cc.Flags = CC_RGBINIT Or CC_FULLOPEN Or CC_ANYCOLOR
    Else
        cc.Flags = CC_FULLOPEN Or CC_ANYCOLOR
    End If
    
    cc.lpCustColors = VarPtr(mCustColors(0))

    If ChooseColorAPI(cc) <> 0 Then
        result = cc.rgbResult
        PromptForCustomRGB = True
        Logging.LogInfo "ColourPicker.PromptForCustomRGB", "Colour chosen: " & result
    Else
        PromptForCustomRGB = False
    End If

    Exit Function

ErrHandler:
    Logging.LogError "ColourPicker.PromptForCustomRGB", Err.Number, Err.Description
    PromptForCustomRGB = False

End Function

'-------------------------------------------------------------------------------
' CreateDynamicColourSwatch
' Description : Creates a 32x32 solid-colour bitmap at runtime for dynamic
'               Ribbon gallery items (e.g. recent custom colours).
' Returns     : IPictureDisp object suitable for Ribbon getImage / getItemImage.
'-------------------------------------------------------------------------------
Public Function CreateDynamicColourSwatch(ByVal rgbColour As Long) As Object

    On Error GoTo ErrHandler

    If mSwatchCache Is Nothing Then Set mSwatchCache = CreateObject("Scripting.Dictionary")

    Dim cacheKey As String
    cacheKey = CStr(rgbColour)

    If mSwatchCache.Exists(cacheKey) Then
        Set CreateDynamicColourSwatch = mSwatchCache(cacheKey)
        Exit Function
    End If

    #If VBA7 Then
        Dim hScreenDC As LongPtr, hMemDC As LongPtr, hBitmap As LongPtr, hOldBmp As LongPtr, hBrush As LongPtr
    #Else
        Dim hScreenDC As Long, hMemDC As Long, hBitmap As Long, hOldBmp As Long, hBrush As Long
    #End If

    hScreenDC = GetDC(0&)
    hMemDC = CreateCompatibleDC(hScreenDC)
    hBitmap = CreateCompatibleBitmap(hScreenDC, 32, 32)
    hOldBmp = SelectObject(hMemDC, hBitmap)

    Dim rc As RECT
    SetRect rc, 0, 0, 32, 32
    hBrush = CreateSolidBrush(rgbColour)
    FillRect hMemDC, rc, hBrush

    SelectObject hMemDC, hOldBmp
    DeleteObject hBrush
    DeleteDC hMemDC
    ReleaseDC 0&, hScreenDC

    Dim pic As PICTDESC
    pic.cbSizeOfStruct = LenB(pic)
    pic.picType = 1   ' vbPicTypeBitmap
    pic.hBitmap = hBitmap
    pic.hPal = 0&

    Dim iPictureGUID As GUID
    With iPictureGUID
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With

    Dim objPic As Object
    If OleCreatePictureIndirect(pic, iPictureGUID, 1, objPic) = 0 Then
        Set CreateDynamicColourSwatch = objPic
        Set mSwatchCache(cacheKey) = objPic
    End If

    Exit Function

ErrHandler:
    Logging.LogError "ColourPicker.CreateDynamicColourSwatch", Err.Number, Err.Description
    Set CreateDynamicColourSwatch = Nothing

End Function

'-------------------------------------------------------------------------------
' ClearSwatchCache
' Description : Drops all cached swatch pictures. Call after Reset Settings
'               or when recent colours change substantially, so the cache
'               can't grow without bound across a very long Excel session.
'-------------------------------------------------------------------------------
Public Sub ClearSwatchCache()
    On Error Resume Next
    If Not mSwatchCache Is Nothing Then mSwatchCache.RemoveAll
End Sub

'-------------------------------------------------------------------------------
' InitCustomColours
' Seeds the 16 custom colour slots with sensible defaults so the dialog
' doesn't open to all-black custom swatches.
'-------------------------------------------------------------------------------
Private Sub InitCustomColours()

    Dim defaults As Variant
    ' Include the add-in's own palette colours plus a few extras.
    defaults = Array(RGB_YELLOW, RGB_GREEN, RGB_ORANGE, RGB_CYAN, _
                     RGB_BLUE, RGB_PINK, RGB_GREY, RGB(255, 255, 255), _
                     RGB(0, 0, 0), RGB(255, 0, 0), RGB(0, 255, 0), RGB(0, 0, 255), _
                     RGB(255, 255, 204), RGB(204, 255, 255), RGB(255, 204, 204), RGB(204, 204, 255))

    Dim i As Long
    For i = 0 To 15
        mCustColors(i) = CLng(defaults(i))
    Next i

End Sub

'-------------------------------------------------------------------------------
' GetExcelHWND
' Retrieves the main Excel window handle for the ChooseColor dialog to
' centre itself on. Uses FindWindow on the "XLMAIN" class.
'-------------------------------------------------------------------------------
#If VBA7 Then
Private Function GetExcelHWND() As LongPtr
#Else
Private Function GetExcelHWND() As Long
#End If
    On Error Resume Next
    GetExcelHWND = FindWindow("XLMAIN", vbNullString)
End Function
