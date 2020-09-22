Attribute VB_Name = "modDrawing"
Option Explicit
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" ( _
    ByVal dwFlags As Long, lpSource As Any, dwMessageID As Long, _
    ByVal dwLanguageID As Long, lpBuffer As String, _
    ByVal nSize As Long, Arguments As Long) As Long
' =====================================================================
' APIs used primarily for drawing/graphics
' =====================================================================
Private Declare Function StretchBlt Lib "gdi32" ( _
        ByVal hDC As Long, ByVal x As Long, ByVal y As Long, _
        ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
        ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, _
        ByVal nSrcHeight As Long, ByVal dwRop As Long) _
    As Long
Public Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBR As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CopyImage Lib "user32" (ByVal Handle As Long, ByVal ImageType As Long, ByVal newWidth As Long, _
    ByVal NewHeight As Long, ByVal lFlags As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" _
     (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, _
     ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Const NEWTRANSPARENT = 3 'use with SetBkMode()
Private Declare Function CreatePen Lib "gdi32" _
     (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" _
     (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" _
     (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function Rectangle Lib "gdi32" _
     (ByVal hDC As Long, ByVal x1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
' following public functions/types may not be used in these modules but are
' used in my CodeSafe program & are here for organizational reasons
Public Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Integer
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As NONCLIENTMETRICS, ByVal fuWinIni As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" _
     (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, _
     ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
     ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
     ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
     ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Const DT_CALCRECT = &H400
Public Const DT_LEFT = &H0
Public Const DT_SINGLELINE = &H20
Public Const DT_NOCLIP = &H100
Private Const DT_CENTER = &H1

' Types used for fonts & images
Public Type POINTAPI
    x As Long
    y As Long
End Type
Public Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte '0=false; 255=true
  lfUnderline As Byte '0=f; 255=t
  lfStrikeOut As Byte '0=f; 255=t
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName As String * 32
End Type
Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Private Type TEXTMETRIC
  tmHeight As Long
  tmAscent As Long
  tmDescent As Long
  tmInternalLeading As Long
  tmExternalLeading As Long
  tmAveCharWidth As Long
  tmMaxCharWidth As Long
  tmWeight As Long
  tmOverhang As Long
  tmDigitizedAspectX As Long
  tmDigitizedAspectY As Long
  tmFirstChar As Byte
  tmLastChar As Byte
  tmDefaultChar As Byte
  tmBreakChar As Byte
  tmItalic As Byte
  tmUnderlined As Byte
  tmStruckOut As Byte
  tmPitchAndFamily As Byte
  tmCharSet As Byte
End Type
Private Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As LOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONT
    lfStatusFont As LOGFONT
    lfMessageFont As LOGFONT
End Type
' Other constants used for graphics
Private Const WHITENESS = &HFF0062
Private Const MAGICROP = &HB8074A
Private Const DSna = &H220326 '0x00220326
'Color constants for GetSysColor
Public Enum ColConst
    COLOR_ACTIVEBORDER = 10
    COLOR_ACTIVECAPTION = 2
    COLOR_ADJ_MAX = 100
    COLOR_ADJ_MIN = -100
    COLOR_APPWORKSPACE = 12
    COLOR_BACKGROUND = 1
    COLOR_BTNFACE = 15
    COLOR_BTNHIGHLIGHT = 20
    COLOR_BTNLIGHT = 22
    COLOR_BTNSHADOW = 16
    COLOR_BTNTEXT = 18
    COLOR_CAPTIONTEXT = 9
    COLOR_GRAYTEXT = 17
    COLOR_HIGHLIGHT = 13
    COLOR_HIGHLIGHTTEXT = 14
    COLOR_INACTIVEBORDER = 11
    COLOR_INACTIVECAPTION = 3
    COLOR_INACTIVECAPTIONTEXT = 19
    COLOR_MENU = 4
    COLOR_MENUTEXT = 7
    COLOR_SCROLLBAR = 0
    COLOR_WINDOW = 5
    COLOR_WINDOWFRAME = 6
    COLOR_WINDOWTEXT = 8
End Enum
' local variables
Private m_hDC As Long               ' reference to DC being drawn in
Private m_Font(0 To 1) As Long      ' local copy of menu font
Private m_FontOld As Long           ' font of DC prior to replacing with menu font

Public Sub DrawRect(ByVal x1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
' =====================================================================
'   Simply draws a rectangle using the current DC's background color
' =====================================================================
     If m_hDC = 0 Then Exit Sub
     Call Rectangle(m_hDC, x1, Y1, X2, Y2)
End Sub

Public Function GetPen(ByVal nWidth As Long, ByVal Clr As Long) As Long
' =====================================================================
' Creates a colored pen for drawing
' =====================================================================
    GetPen = CreatePen(0, nWidth, Clr)
End Function

Public Sub DrawCaption(ByVal x As Long, ByVal y As Long, tRect As RECT, _
    ByVal hStr As String, hAccel As String, iTab As Integer, _
    ByVal Clr As Long, Optional bCenter As Boolean = False, _
    Optional iOffset As Integer = 0)
    
' =====================================================================
' Prints text to current DC in the coodinates & colors provided
' =====================================================================
    If m_hDC = 0 Then Exit Sub
    'Equivalent to setting a form's property FontTransparent = True
    SetBkMode m_hDC, NEWTRANSPARENT
    Dim OT As Long, x1 As Long
    ' set text color and set x,y coordinates for printing
    OT = GetTextColor(m_hDC)
    SetTextColor m_hDC, Clr
    x1 = tRect.Right
    ' print the caption/text
    If bCenter Then
        DrawText m_hDC, hStr, Len(hStr), tRect, DT_NOCLIP Or DT_CALCRECT Or DT_CALCRECT Or DT_SINGLELINE
        tRect.Left = (x1 - iOffset - tRect.Right) \ 2 + iOffset
        tRect.Right = tRect.Left + tRect.Right
    Else
        tRect.Left = x + iOffset
    End If
    tRect.Top = y
    tRect.Bottom = tRect.Bottom + y
    DrawText m_hDC, hStr, Len(hStr), tRect, DT_SINGLELINE Or DT_NOCLIP Or DT_LEFT
    If Len(hAccel) Then
        ' here we will print an acceleraor key if needed
        tRect.Left = tRect.Left + iTab - iOffset
        tRect.Top = y
        DrawText m_hDC, hAccel, Len(hAccel), tRect, DT_LEFT Or DT_NOCLIP Or DT_SINGLELINE
    End If
    'Restore old text color
    SetTextColor m_hDC, OT
End Sub

Public Property Let TargethDC(ByVal vNewValue As Long)
' =====================================================================
' Maintain a local reference to the DC being drawn in
' simply prevents having to pass it to each call to a drawing routine
' =====================================================================
     m_hDC = vNewValue
End Property

Public Sub ThreeDbox(ByVal x1 As Long, ByVal Y1 As Long, _
ByVal X2 As Long, ByVal Y2 As Long, bSelected As Boolean, _
Optional Sunken As Boolean = False)
' =====================================================================
'   Draw/erase a raised/sunken box around the specified coordinates.
' =====================================================================

     If m_hDC = 0 Then Exit Sub

     Dim CurPen As Long, OldPen As Long
     Dim dm As POINTAPI, iOffset As Integer
     
     ' select colors, offset when set indicates erasing
     iOffset = Abs(CInt(bSelected)) + 1
     If Sunken = False Then
         CurPen = GetPen(1, GetSysColor(Choose(iOffset, COLOR_MENU, COLOR_BTNHIGHLIGHT)))
     Else
          CurPen = GetPen(1, GetSysColor(Choose(iOffset, COLOR_MENU, COLOR_BTNSHADOW)))
     End If
     OldPen = SelectObject(m_hDC, CurPen)
     
     'First - Light Line
     MoveToEx m_hDC, x1 + 2, Y2, dm
     LineTo m_hDC, x1 + 2, Y1
     LineTo m_hDC, X2 - 2, Y1

     SelectObject m_hDC, OldPen
     DeleteObject CurPen
     ' Next - Dark line
     If Sunken = False Then
          CurPen = GetPen(1, GetSysColor(Choose(iOffset, COLOR_MENU, COLOR_BTNSHADOW)))
     Else
          CurPen = GetPen(1, GetSysColor(Choose(iOffset, COLOR_MENU, COLOR_BTNHIGHLIGHT)))
     End If
     OldPen = SelectObject(m_hDC, CurPen)
     
     MoveToEx m_hDC, X2 - 2, Y1, dm
     LineTo m_hDC, X2 - 2, Y2
     LineTo m_hDC, x1 + 2, Y2

     ' Replace pen & delete temp pen
     SelectObject m_hDC, OldPen
     DeleteObject CurPen
End Sub

Public Sub DrawMenuIcon(lImageHdl As Long, ImageType As Long, _
    rt As RECT, bdisabled As Boolean, Optional bInColor As Boolean = True, _
    Optional bForceTransparency As Long = 0, Optional iOffset As Integer = 0, _
    Optional yOffset As Integer, Optional IMGwidth As Integer = 16, _
    Optional IMGheight As Integer = 16, Optional lMask As Long = -1)
' =====================================================================
'   Draws imagelist image on destined DC
' =====================================================================

' ensure the requested image exists
If lImageHdl = 0 Then Exit Sub

Dim lImageSmall As Long, lDrawType As Long, lImageType As Long
Const DSS_DISABLED = &H20
Const DSS_NORMAL = &H0
Const DSS_BITMAP = &H4
Const DSS_ICON = &H3
Const CI_BITMAP = &H0
Const CI_ICON = &H1
Dim rcImage As RECT

If ImageType < 2 Then
    lDrawType = DSS_BITMAP
    lImageType = CI_BITMAP
Else
    lDrawType = DSS_ICON
    lImageType = CI_ICON
End If
lImageSmall = CopyImage(lImageHdl, lImageType, IMGwidth, IMGheight, DSS_NORMAL)
If lImageSmall = 0 Then ' failed to make a copy from the imagetype passed, try the other settings
    If lDrawType = DSS_BITMAP Then
        lDrawType = DSS_ICON
        lImageType = CI_ICON
    Else
        lDrawType = DSS_BITMAP
        lImageType = CI_BITMAP
    End If
    lImageSmall = CopyImage(lImageHdl, lImageType, IMGwidth, IMGheight, DSS_NORMAL)
End If
If lImageSmall = 0 Then Exit Sub

If bdisabled = False Then
    ' if not disabled, then straightforward extraction/drawing on coords
    If ((lImageType = CI_ICON And bForceTransparency < 2) Or bForceTransparency = 2) Then
        DrawState m_hDC, 0, 0, lImageSmall, 0, rt.Left + iOffset, rt.Top + yOffset, 0, 0, lDrawType
    Else
        MakeTransparentBitmap lImageSmall, rt.Left + iOffset, rt.Top + yOffset, IMGwidth, IMGheight, , , lMask
    End If
    DeleteObject lImageSmall
Else
' =====================================================================
'   This function is from Paul DiLascia's DrawEmbossed function
'   which draws colored disabled pictures.
'   To be fair to him, I modified several lines of code so
'   it is customized for CodeSafe & should it fail -- not his fault
' =====================================================================

'  // create mono or color bitmap
    Dim hBitmap As Long
    If bInColor Then
        hBitmap& = CreateCompatibleBitmap(m_hDC&, IMGwidth, IMGheight)
    Else
        hBitmap& = CreateBitmap(IMGwidth, IMGheight, 1, 1, vbNull)
    End If
'  // draw image into memory DC--fill BG white first
'  // create memory dc
    Dim hOldBitmap As Long
    Dim hmemDC As Long
    hmemDC& = CreateCompatibleDC(m_hDC&)
    hOldBitmap = SelectObject(hmemDC&, hBitmap&)
    Call PatBlt(hmemDC&, 0, 0, IMGwidth, IMGheight, WHITENESS)
    If (lImageType = CI_ICON And bForceTransparency < 2) Or bForceTransparency = 2 Then
        DrawState hmemDC&, 0, 0, lImageSmall, 0, 0, 0, 0, 0, lDrawType
    Else
        MakeTransparentBitmap lImageSmall, 0, 0, IMGwidth, IMGheight, , , lMask, hmemDC
    End If
    DeleteObject lImageSmall
'  // This seems to be required.
  Dim hOldBackColor As Long
  hOldBackColor& = SetBkColor(m_hDC&, RGB(255, 255, 255))
'  // Draw using hilite offset by (1,1), then shadow
  Dim hbrShadow As Long, hbrHilite As Long
  hbrShadow& = CreateSolidBrush(GetSysColor(COLOR_BTNSHADOW))
  hbrHilite& = CreateSolidBrush(GetSysColor(COLOR_BTNHIGHLIGHT))
  
  Dim hOldBrush As Long
  hOldBrush& = SelectObject(m_hDC&, hbrHilite&)

  Call BitBlt(m_hDC&, rt.Left + 1 + iOffset, rt.Top + 1 + yOffset, IMGwidth, IMGheight, hmemDC&, 0, 0, MAGICROP)
  Call SelectObject(m_hDC&, hbrShadow&)
  Call BitBlt(m_hDC&, rt.Left + iOffset, rt.Top + yOffset, IMGwidth, IMGheight, hmemDC&, 0, 0, MAGICROP)
  
  Call SelectObject(m_hDC&, hOldBrush&)
  Call SetBkColor(m_hDC&, hOldBackColor&)
  Call SelectObject(hmemDC&, hOldBitmap&)
  Call DeleteObject(hOldBrush&)
  Call DeleteObject(hbrHilite&)
  Call DeleteObject(hbrShadow&)
  Call DeleteObject(hOldBackColor&)
  Call DeleteObject(hOldBitmap&)
  Call DeleteObject(hBitmap&)
  Call DeleteDC(hmemDC&)
End If
End Sub

Private Sub MakeTransparentBitmap(imgHdl As Long, _
               ByVal xDest As Long, _
               ByVal yDest As Long, _
               ByVal Width As Long, _
               ByVal Height As Long, _
               Optional xSrc As Long = 0, _
               Optional ByVal ySrc As Long = 0, _
               Optional clrMask As OLE_COLOR = -1, _
               Optional destDC As Long = 0)
' =====================================================================
' Borrowed and modified - creates a transparent bitmap
' =====================================================================

    Dim hdcSrc As Long
    Dim hbmMemSrcOld As Long
    Dim hdcMask As Long     'HDC of the created mask image
    Dim hdcColor As Long    'HDC of the created color image
    Dim hbmMask As Long     'Bitmap handle to the mask image
    Dim hbmColor As Long    'Bitmap handle to the color image
    Dim hbmColorOld As Long
    Dim hbmMaskOld As Long
    Dim hPalOld As Long
    Dim hdcScreen As Long
    Dim hdcScnBuffer As Long 'Buffer to do all work on
    Dim hbmScnBuffer As Long
    Dim hbmScnBufferOld As Long
    Dim hPalBufferOld As Long
    Dim lMaskColor As Long
    Const hPal As Long = 0
    
    On Error Resume Next
    hdcScreen = GetDC(0&)
    
    If destDC = 0 Then destDC = m_hDC
    hdcSrc = CreateCompatibleDC(hdcScreen)
    hbmMemSrcOld = SelectObject(hdcSrc, imgHdl)
    RealizePalette hdcSrc
    
    If clrMask < 0 Then clrMask = GetPixel(hdcSrc, 0, 0)
    OleTranslateColor clrMask, hPal, lMaskColor

    'Create a color bitmap to server as a copy of the destination
    'Do all work on this bitmap and then copy it back over the
    'destination when it's done.
    hbmScnBuffer = CreateCompatibleBitmap(hdcScreen, Width, Height)
    'Create DC for screen buffer
    hdcScnBuffer = CreateCompatibleDC(hdcScreen)
    hbmScnBufferOld = SelectObject(hdcScnBuffer, hbmScnBuffer)
    hPalBufferOld = SelectPalette(hdcScnBuffer, hPal, True)
    RealizePalette hdcScnBuffer
    'Copy the destination to the screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, destDC, xDest, yDest, vbSrcCopy
    'Create a (color) bitmap for the cover (can't use
    'CompatibleBitmap with hdcSrc, because this will create a
    'DIB section if the original bitmap is a DIB section)
    hbmColor = CreateCompatibleBitmap(hdcScreen, Width, Height)
    'Now create a monochrome bitmap for the mask
    hbmMask = CreateBitmap(Width, Height, 1, 1, ByVal 0&)
    'First, blt the source bitmap onto the cover.  We do this
    'first and then use it instead of the source bitmap
    'because the source bitmap may be
    'a DIB section, which behaves differently than a bitmap.
    '(Specifically, copying from a DIB section to a monochrome
    'bitmap does a nearest-color selection rather than painting
    'based on the backcolor and forecolor.
    hdcColor = CreateCompatibleDC(hdcScreen)
    hbmColorOld = SelectObject(hdcColor, hbmColor)
    hPalOld = SelectPalette(hdcColor, hPal, True)
    RealizePalette hdcColor
    'In case hdcSrc contains a monochrome bitmap, we must set
    'the destination foreground/background colors according to
    'those currently set in hdcSrc (because Windows will
    'associate these colors with the two monochrome colors)
    SetBkColor hdcColor, GetBkColor(hdcSrc)
    SetTextColor hdcColor, GetTextColor(hdcSrc)
    BitBlt hdcColor, 0, 0, Width, Height, hdcSrc, xSrc, ySrc, vbSrcCopy
    'Paint the mask.  What we want is white at the transparent
    'color from the source, and black everywhere else.
    hdcMask = CreateCompatibleDC(hdcScreen)
    hbmMaskOld = SelectObject(hdcMask, hbmMask)

    'When bitblt'ing from color to monochrome, Windows sets to 1
    'all pixels that match the background color of the source DC.
    'All other bits are set to 0.
    SetBkColor hdcColor, lMaskColor
    SetTextColor hdcColor, vbWhite
    BitBlt hdcMask, 0, 0, Width, Height, hdcColor, 0, 0, _
        vbSrcCopy
    'Paint the rest of the cover bitmap.
    '
    'What we want here is black at the transparent color,
    'and the original colors everywhere else.  To do this,
    'we first paint the original onto the cover (which we
    'already did), then we AND the inverse of the mask onto
    'that using the DSna ternary raster operation
    '(0x00220326 - see Win32 SDK reference, Appendix,
    '"Raster Operation Codes", "Ternary
    'Raster Operations", or search in MSDN for 00220326).
    'DSna [reverse polish] means "(not SRC) and DEST".
    '
    'When bitblt'ing from monochrome to color, Windows
    'transforms all white bits (1) to the background color
    'of the destination hdc.  All black (0)
    'bits are transformed to the foreground color.
    SetTextColor hdcColor, vbBlack
    SetBkColor hdcColor, vbWhite
    BitBlt hdcColor, 0, 0, Width, Height, hdcMask, 0, 0, DSna
    'Paint the Mask to the Screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hdcMask, 0, 0, vbSrcAnd
    'Paint the Color to the Screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hdcColor, 0, 0, vbSrcPaint
    'Copy the screen buffer to the screen
    BitBlt destDC, xDest, yDest, Width, Height, hdcScnBuffer, 0, 0, vbSrcCopy
    'All done!
    DeleteObject SelectObject(hdcColor, hbmColorOld)
    SelectPalette hdcColor, hPalOld, True
    RealizePalette hdcColor
    DeleteDC hdcColor
    DeleteObject SelectObject(hdcScnBuffer, hbmScnBufferOld)
    SelectPalette hdcScnBuffer, hPalBufferOld, True
    RealizePalette hdcScnBuffer
    DeleteDC hdcScnBuffer

    DeleteObject SelectObject(hdcMask, hbmMaskOld)
    DeleteDC hdcMask
    ReleaseDC 0&, hdcScreen
    
    SelectObject hdcSrc, hbmMemSrcOld
    RealizePalette hdcSrc
    DeleteDC hdcSrc
    ReleaseDC 0&, hdcScreen
    
End Sub

Public Sub SetMenuFont(bSet As Boolean, Optional hDC As Long, _
    Optional bReduced As Boolean = False, Optional otherFont As Long = 0)
' =====================================================================
'   This creates system menu fonts for the destination DC if needed
'   and either sets it or removes it from the DC
' =====================================================================

'   reference the current DC for all drawing
    If hDC Then m_hDC = hDC
    If bSet Then
        ' in order to set the font, we must first determine what it is
        If m_Font(0) = 0 And otherFont = 0 Then
            Dim ncm As NONCLIENTMETRICS, newFont As LOGFONT
            ncm.cbSize = Len(ncm)
            ' this will return the system font info along with other stuff
            SystemParametersInfo 41, 0, ncm, 0
            ' here we create a memory font based off of system menu font
            newFont = ncm.lfMenuFont
            m_Font(0) = CreateFontIndirect(newFont)
            ' now we are going to try to create a scalable font for
            ' separator bar text just in case the computer's menu font
            ' is not scalable. The following is a shortcut way of creating
            ' the font & I hope it works on all systems!
            newFont.lfFaceName = "Tahoma" & Chr$(0)
            newFont.lfCharSet = 1
            newFont.lfHeight = (7.5 * -20) / Screen.TwipsPerPixelY
            'newFont.lfHeight = newFont.lfHeight + 1
            m_Font(1) = CreateFontIndirect(newFont)
        End If
        ' add the font to the DC & keep reference to old font
        ' Calling routines responsbile for restoring original font
        ' with call back to this routine & a FALSE parameter
        If m_FontOld = 0 And otherFont = 0 Then
            m_FontOld = SelectObject(m_hDC, m_Font(Abs(CInt(bReduced))))
        Else
            If otherFont Then
                SelectObject m_hDC, otherFont
            Else
                SelectObject m_hDC, m_Font(Abs(CInt(bReduced)))
            End If
        End If
    Else
        ' Restoring old font
        If m_hDC = 0 Then Exit Sub
        SelectObject m_hDC, m_FontOld
    End If
End Sub

Public Sub DestroyMenuFont()
' =====================================================================
' Simply destroy the memory font to free up resources
' =====================================================================
    On Error Resume Next
    SelectObject m_hDC, m_FontOld
    DeleteObject m_Font(0)
    DeleteObject m_Font(1)
    m_Font(0) = 0
    m_Font(1) = 0
End Sub

Public Sub DoGradientBkg(lColor As Long, tRect As RECT, hwnd As Long)
'=======================================================================

'=======================================================================

Dim sColor As String, i As Integer, tmpSB As PictureBox, formID As Long
Dim R As Integer, B As Integer, G As Integer
Dim lColorStep As Long, lNewColor As Long

On Error GoTo GradientErrors
' we are going to create a picturebox to draw the gradient in
' tried drawing directly to hdc via MoveTo & LineTo APIs, but
' everytime, it failed -- maybe Win98, maybe my graphics card?
' This works though a little slower
formID = GetFormHandle(hwnd)
' create the picture box in memory
Set tmpSB = Forms(formID).Controls.Add("VB.PictureBox", "pic___tmp_s_b", Forms(formID))
With tmpSB
    ' set picturebox attributes
    .Visible = False
    .BorderStyle = 0
    .AutoRedraw = True
    .DrawMode = 13
    .DrawWidth = 1
    .Height = .ScaleY(tRect.Bottom, vbPixels, .ScaleMode)
    .Width = .ScaleX(tRect.Right, vbPixels, .ScaleMode)
    .ScaleMode = vbPixels
    lNewColor = lColor
    ' loop thru each line & color it
    For i = 1 To tRect.Bottom - 1
        ' this line is used to subtract/add colors
        ' for a more dramatic fade, increment the #2 below
        lColorStep = (2 / tRect.Bottom) * i
        ' modify the current color
        B = ((lNewColor \ &H10000) Mod &H100) - lColorStep
        G = ((lNewColor \ &H100) Mod &H100) - lColorStep
        R = (lNewColor And &HFF) - lColorStep
        ' ensure the Red, Green, Blue values are in acceptable ranges
        If R < 0 Then
            R = 0
        ElseIf R > 255 Then
            R = 255
        End If
        If G < 0 Then
            G = 0
        ElseIf G > 255 Then
            G = 255
        End If
        If B < 0 Then
            B = 0
        ElseIf B > 255 Then
            B = 255
        End If
        lNewColor = RGB(R, G, B)    ' cache the color & draw the line
        tmpSB.Line (0, i - 1)-(tRect.Right, i - 1), lNewColor, BF
    Next
    ' now that the gradient has been drawn, copy it to the menu panel
    BitBlt m_hDC, 0, 0, tRect.Right, tRect.Bottom, .hDC, 0, 0, vbSrcCopy
End With
GradientErrors:
On Error Resume Next
' clean up
Forms(formID).Controls.Remove "pic___tmp_s_b"
Set tmpSB = Nothing
End Sub

Public Sub DrawCheckMark(pRect As RECT, lColor As Long, _
    bdisabled As Boolean, Optional xtraOffset As Long = 0)
' =====================================================================
' Simple little check mark drawing, looks good 'nuf I think
' =====================================================================

Dim CurPen As Long, OldPen As Long
Dim dm As POINTAPI
Dim yOffset As Integer, xOffset As Integer
Dim x1 As Integer, X2 As Integer
Dim Y1 As Integer, Y2 As Integer

CurPen = GetPen(1, lColor)
OldPen = SelectObject(m_hDC, CurPen)

xOffset = 6 + xtraOffset
yOffset = pRect.Top + 6

' Here we are simply tracing the outline of a check box
' Created by opening a 8x8 bitmap editor and drawing a
' simple checkmark from left to right, bottom to top
MoveToEx m_hDC, 1 + xOffset, 4 + yOffset, dm
LineTo m_hDC, 2 + xOffset, 4 + yOffset
LineTo m_hDC, 2 + xOffset, 5 + yOffset
LineTo m_hDC, 3 + xOffset, 5 + yOffset
LineTo m_hDC, 3 + xOffset, 6 + yOffset
LineTo m_hDC, 4 + xOffset, 6 + yOffset
LineTo m_hDC, 4 + xOffset, 4 + yOffset
LineTo m_hDC, 5 + xOffset, 4 + yOffset
LineTo m_hDC, 5 + xOffset, 2 + yOffset
LineTo m_hDC, 6 + xOffset, 2 + yOffset
LineTo m_hDC, 6 + xOffset, 1 + yOffset
LineTo m_hDC, 7 + xOffset, 1 + yOffset
LineTo m_hDC, 7 + xOffset, 0 + yOffset

' replace original pen
SelectObject m_hDC, OldPen
DeleteObject CurPen
End Sub
