VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIBSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Type SAFEARRAYBOUND
  cElements As Long
  lLbound As Long
End Type

Private Type SAFEARRAY2D
  cDims As Integer
  fFeatures As Integer
  cbElements As Long
  cLocks As Long
  pvData As Long
  Bounds(0 To 1) As SAFEARRAYBOUND
End Type

Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias _
        "VarPtr" (Ptr() As Any) As Long

Private Type RGBQUAD
  rgbBlue As Byte
  rgbGreen As Byte
  rgbRed As Byte
  rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type

Private Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmiColors As RGBQUAD
End Type

Private Declare Function CreateCompatibleDC Lib "GDI32" _
        (ByVal hDC As Long) As Long
        
Private Declare Function GetDC Lib "USER32" (ByVal hWnd As _
        Long) As Long
        
Private Declare Function GetDesktopWindow Lib "USER32" () As Long

Private Declare Function CreateDIBSection Lib "GDI32" (ByVal hDC _
        As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, _
        lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) _
        As Long
        
Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As _
        Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As _
        Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
        ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As _
        Long) As Long
        
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC _
        As Long, ByVal hObject As Long) As Long
        
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject _
        As Long) As Long
        
Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As _
        Long) As Long
        
Private Declare Function LoadImage Lib "USER32" Alias "LoadImageA" _
        (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As _
        Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As _
        Long) As Long

Private Const BI_RGB = 0&
Private Const BI_RLE4 = 2&
Private Const BI_RLE8 = 1&
Private Const DIB_RGB_COLORS = 0

Private Type BITMAP
  bmType As Long
  BmWidth As Long
  BmHeight As Long
  bmWidthBytes As Long
  BmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type

Private Declare Function GetObjectAPI Lib "GDI32" Alias "GetObjectA" _
        (ByVal hObject As Long, ByVal nCount As Long, lpObject As _
        Any) As Long
        
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal _
        hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) _
        As Long

Private Declare Function OpenClipboard Lib "USER32" (ByVal hWnd _
        As Long) As Long
        
Private Declare Function CloseClipboard Lib "USER32" () As Long

Private Declare Function SetClipboardData Lib "USER32" (ByVal _
        wFormat As Long, ByVal hMem As Long) As Long
        
Private Declare Function EmptyClipboard Lib "USER32" () As Long

Private Const CF_BITMAP = 2
Private Const CF_DIB = 8

Private m_hDIb As Long
Private m_hBmpOld As Long
Private m_hDC As Long
Private m_lPtr As Long
Private m_tBI As BITMAPINFO

Public Function CopyToClipboard(Optional ByVal bAsDIB As _
                                Boolean = True) As Boolean
  Dim lhDCDesktop As Long
  Dim lhDC As Long
  Dim lhBmpOld As Long
  Dim hObj As Long
  Dim lFmt As Long
  Dim b() As Byte
  Dim tBI As BITMAPINFO
  Dim lPtr As Long
  Dim hDibCopy As Long

    lhDCDesktop = GetDC(GetDesktopWindow())
    If (lhDCDesktop <> 0) Then
      lhDC = CreateCompatibleDC(lhDCDesktop)
      If (lhDC <> 0) Then
        If (bAsDIB) Then
          MsgBox "I don't know how to put a DIB on the clipboard! Copy as bitmap instead!!!"
        Else
          hObj = CreateCompatibleBitmap(lhDCDesktop, Width, Height)
          If (hObj <> 0) Then
            lhBmpOld = SelectObject(lhDC, hObj)
            PaintPicture lhDC
            SelectObject lhDC, lhBmpOld
            lFmt = CF_BITMAP
            If (OpenClipboard(0) <> 0) Then
              EmptyClipboard
              If (SetClipboardData(lFmt, hObj) <> 0) Then
                CopyToClipboard = True
              End If
              CloseClipboard
            End If
          End If
        End If
        DeleteDC lhDC
      End If
      DeleteDC lhDCDesktop
    End If
End Function

Public Function CreateDIB(ByVal lhDC As Long, ByVal lWidth As Long, _
                          ByVal lHeight As Long, ByRef hDib As Long _
                          ) As Boolean
  With m_tBI.bmiHeader
    .biSize = Len(m_tBI.bmiHeader)
    .biWidth = lWidth
    .biHeight = lHeight
    .biPlanes = 1
    .biBitCount = 24
    .biCompression = BI_RGB
    .biSizeImage = BytesPerScanLine * .biHeight
  End With
  
  hDib = CreateDIBSection(lhDC, m_tBI, DIB_RGB_COLORS, m_lPtr, 0, 0)
  CreateDIB = (hDib <> 0)
End Function

Public Function CreateFromPicture(ByRef picThis As StdPicture)
  Dim lhDC As Long
  Dim lhDCDesktop As Long
  Dim lhBmpOld As Long
  Dim tBMP As BITMAP
  On Error GoTo fehler
    GetObjectAPI picThis.handle, Len(tBMP), tBMP
    If (Create(tBMP.BmWidth, tBMP.BmHeight)) Then
      lhDCDesktop = GetDC(GetDesktopWindow())
      If (lhDCDesktop <> 0) Then
        lhDC = CreateCompatibleDC(lhDCDesktop)
        DeleteDC lhDCDesktop
        If (lhDC <> 0) Then
          lhBmpOld = SelectObject(lhDC, picThis.handle)
          LoadPictureBlt lhDC
          SelectObject lhDC, lhBmpOld
          DeleteObject lhDC
        End If
      End If
    End If
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in CreateFromPicture/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function

Public Function Create(ByVal lWidth&, ByVal lHeight&) As Boolean
  ClearUp
  m_hDC = CreateCompatibleDC(0)
  If (m_hDC <> 0) Then
    If (CreateDIB(m_hDC, lWidth, lHeight, m_hDIb)) Then
      m_hBmpOld = SelectObject(m_hDC, m_hDIb)
      Create = True
    Else
      DeleteObject m_hDC
      m_hDC = 0
    End If
  End If
End Function

Public Property Get BytesPerScanLine() As Long
  BytesPerScanLine = (m_tBI.bmiHeader.biWidth * 3 + 3) And &HFFFFFFFC
End Property

Public Property Get Width() As Long
  Width = m_tBI.bmiHeader.biWidth
End Property

Public Property Get Height() As Long
  Height = m_tBI.bmiHeader.biHeight
End Property

Public Sub LoadPictureBlt(ByVal lhDC&, Optional ByVal lSrcLeft& = 0, _
                          Optional ByVal lSrcTop& = 0, Optional _
                          ByVal lSrcWidth& = -1, Optional ByVal _
                          lSrcHeight& = -1, Optional ByVal eRop As _
                          RasterOpConstants = vbSrcCopy)
                          
  If lSrcWidth < 0 Then lSrcWidth = m_tBI.bmiHeader.biWidth
  If lSrcHeight < 0 Then lSrcHeight = m_tBI.bmiHeader.biHeight
  BitBlt m_hDC, 0, 0, lSrcWidth, lSrcHeight, lhDC, lSrcLeft, lSrcTop, eRop
End Sub

Public Sub PaintPicture(ByVal lhDC As Long, Optional ByVal _
                        lDestLeft& = 0, Optional ByVal lDestTop& _
                        = 0, Optional ByVal lDestWidth& = -1, _
                        Optional ByVal lDestHeight& = -1, Optional _
                        ByVal lSrcLeft& = 0, Optional ByVal _
                        lSrcTop& = 0, Optional ByVal eRop As _
                        RasterOpConstants = vbSrcCopy)
                        
  If (lDestWidth < 0) Then lDestWidth = m_tBI.bmiHeader.biWidth
  If (lDestHeight < 0) Then lDestHeight = m_tBI.bmiHeader.biHeight
  BitBlt lhDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, m_hDC, lSrcLeft, lSrcTop, eRop
End Sub

Public Property Get hDC() As Long
    hDC = m_hDC
End Property
Public Property Get hDib() As Long
  hDib = m_hDIb
End Property

Public Property Get DIBSectionBitsPtr() As Long
  DIBSectionBitsPtr = m_lPtr
End Property

Public Sub RandomiseBits(Optional ByVal bGray As Boolean = False)
  Dim bDib() As Byte, x&, Y&, lC&, tSA As SAFEARRAY2D, xEnd&

    With tSA
      .cbElements = 1
      .cDims = 2
      .Bounds(0).lLbound = 0
      .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
      .Bounds(1).lLbound = 0
      .Bounds(1).cElements = BytesPerScanLine()
      .pvData = m_lPtr
    End With
    
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4
    Randomize Timer
    
    xEnd = (Width - 1) * 3
    If (bGray) Then
      For Y = 0 To m_tBI.bmiHeader.biHeight - 1
        For x = 0 To xEnd Step 3
          lC = Rnd * 255
          bDib(x, Y) = lC
          bDib(x + 1, Y) = lC
          bDib(x + 2, Y) = lC
        Next x
      Next Y
    Else
      For x = 0 To xEnd Step 3
        For Y = 0 To m_tBI.bmiHeader.biHeight - 1
          bDib(x, Y) = 0
          bDib(x + 1, Y) = Rnd * 255
          bDib(x + 2, Y) = Rnd * 255
        Next Y
      Next x
    End If
    
    CopyMemory ByVal VarPtrArray(bDib), 0&, 4
End Sub

Public Sub ClearUp()
  If (m_hDC <> 0) Then
    If (m_hDIb <> 0) Then
      SelectObject m_hDC, m_hBmpOld
      DeleteObject m_hDIb
    End If
    DeleteObject m_hDC
  End If
  m_hDC = 0: m_hDIb = 0: m_hBmpOld = 0: m_lPtr = 0
End Sub

Public Function Resample(ByVal lNewHeight&, ByVal lNewWidth&) _
                         As cDIBSection
  Dim cDib As cDIBSection
    Set cDib = New cDIBSection
    If cDib.Create(lNewWidth, lNewHeight) Then
      If (lNewWidth <> m_tBI.bmiHeader.biWidth) Or (lNewHeight <> m_tBI.bmiHeader.biHeight) Then
        ResampleDib cDib
      Else
        cDib.LoadPictureBlt m_hDC
      End If
        Set Resample = cDib
    End If
End Function

Private Function ResampleDib(ByRef cDibTo As cDIBSection) As Boolean
  Dim bDibFrom() As Byte, bDibTo() As Byte
  Dim tSAFrom As SAFEARRAY2D, tSATo As SAFEARRAY2D

    With tSAFrom
      .cbElements = 1
      .cDims = 2
      .Bounds(0).lLbound = 0
      .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
      .Bounds(1).lLbound = 0
      .Bounds(1).cElements = BytesPerScanLine()
      .pvData = m_lPtr
    End With
    
    CopyMemory ByVal VarPtrArray(bDibFrom()), VarPtr(tSAFrom), 4

    With tSATo
      .cbElements = 1
      .cDims = 2
      .Bounds(0).lLbound = 0
      .Bounds(0).cElements = cDibTo.Height
      .Bounds(1).lLbound = 0
      .Bounds(1).cElements = cDibTo.BytesPerScanLine()
      .pvData = cDibTo.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDibTo()), VarPtr(tSATo), 4

    Dim xScale As Single, yScale As Single, x&, Y&, xEnd&, xOut&
    Dim fX As Single, fY As Single, ifY&, ifX&
    Dim dX As Single, dy As Single, r&
    Dim r1 As Single, r2 As Single, r3 As Single, r4 As Single
    Dim g&, g1 As Single, g2 As Single, g3 As Single, g4 As Single
    Dim b&, b1 As Single, b2 As Single, b3 As Single, b4 As Single
    Dim ir1&, ig1&, ib1&, ir2&, ig2&, ib2&

    xScale = (Width - 1) / cDibTo.Width
    yScale = (Height - 1) / cDibTo.Height
    xEnd = cDibTo.Width - 1
        
    For Y = 0 To cDibTo.Height - 1
      fY = Y * yScale
      ifY = Int(fY)
      dy = fY - ifY
      
      For x = 0 To xEnd
        fX = x * xScale
        ifX = Int(fX)
        dX = fX - ifX
        
        ifX = ifX * 3
        b1 = bDibFrom(ifX, ifY): g1 = bDibFrom(ifX + 1, ifY): r1 = bDibFrom(ifX + 2, ifY)
        b2 = bDibFrom(ifX + 3, ifY): g2 = bDibFrom(ifX + 4, ifY): r2 = bDibFrom(ifX + 5, ifY)
        b3 = bDibFrom(ifX, ifY + 1): g3 = bDibFrom(ifX + 1, ifY + 1): r3 = bDibFrom(ifX + 2, ifY + 1)
        b4 = bDibFrom(ifX + 3, ifY + 1): g4 = bDibFrom(ifX + 4, ifY + 1): r4 = bDibFrom(ifX + 5, ifY + 1)
        ir1 = r1 * (1 - dy) + r3 * dy: ig1 = g1 * (1 - dy) + g3 * dy: ib1 = b1 * (1 - dy) + b3 * dy
        ir2 = r2 * (1 - dy) + r4 * dy: ig2 = g2 * (1 - dy) + g4 * dy: ib2 = b2 * (1 - dy) + b4 * dy
        r = ir1 * (1 - dX) + ir2 * dX: g = ig1 * (1 - dX) + ig2 * dX: b = ib1 * (1 - dX) + ib2 * dX

        If (r < 0) Then r = 0
        If (r > 255) Then r = 255
        If (g < 0) Then g = 0
        If (g > 255) Then g = 255
        If (b < 0) Then b = 0
        If (b > 255) Then
            b = 255
        End If
        xOut = x * 3
        bDibTo(xOut, Y) = b
        bDibTo(xOut + 1, Y) = g
        bDibTo(xOut + 2, Y) = r
      Next x
    Next Y

    CopyMemory ByVal VarPtrArray(bDibFrom), 0&, 4
    CopyMemory ByVal VarPtrArray(bDibTo), 0&, 4
End Function

Private Sub Class_Terminate()
  ClearUp
End Sub


