Attribute VB_Name = "ModuleMain"
'//////////////////////////////////////////////////////////////////////////////////////
'//
'// ModuleMain : global definitions
'//
'//////////////////////////////////////////////////////////////////////////////////////
'// This is a part of the Innoveda-PowerPCB OLE Automation server PPCBAutoGeom sample.
'// Copyright (C) 2003 Mentor Graphics Corp.
'// All rights reserved.
'//
'// This source code is only intended as a supplement to the Innoveda-PowerPCB OLE
'// Automation Server API Help file.
'//////////////////////////////////////////////////////////////////////////////////////

' The one and only Picture object (to be accessed through GetPicture function (below)).
Private pictureObject As New Picture

' These are resource IDs used.

Public Const AFX_IDS_IDLEMESSAGE = 57345
Public Const IDS_OLEINITFAILED = 129
Public Const IDS_PICTUREBUILDFAILED = 130
Public Const IDS_BADVARIANT = 131
Public Const IDS_SERVERSTARTCONNECTION = 132
Public Const IDS_BADROUTESEGMENTARC = 133
Public Const IDS_DRAWINGMESSAGE = 134
Public Const IDS_LOADDATAMESSAGE = 135

' The following structures and procedures declaration reflect those of Windows API.

' If you wish to use Unicode versions of Windows API functions, please, set this constant to True.
#Const Unicode = False

Public Type Rect
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Public Type Point
    X As Long
    Y As Long
End Type

Public Type XFORM
    eM11 As Single
    eM12 As Single
    eM21 As Single
    eM22 As Single
    eDx As Single
    eDy As Single
End Type

Public Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
End Type

' This structure is used by GetVersionEx function.
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
#If Unicode Then
    szCSDVersion(128 * 2) As Byte
#Else
    szCSDVersion(128) As Byte
#End If
End Type

Declare Function GetVersionExW Lib "KERNEL32" _
 (lpVersionInformation As OSVERSIONINFO) As Long
Declare Function GetVersionExA Lib "KERNEL32" _
 (lpVersionInformation As OSVERSIONINFO) As Long
Declare Function GetSystemMetrics Lib "USER32" _
 (ByVal nIndex As Long) As Long
Declare Function DeleteObject Lib "GDI32" _
 (ByVal hObject As Long) As Long
Declare Function GetDC Lib "USER32" _
 (ByVal hWnd As Long) As Long
Declare Function ReleaseDC Lib "GDI32" _
 (ByVal hWnd As Long, ByVal hdc As Long) As Long
Declare Function SelectObject Lib "GDI32" _
 (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function CreatePen Lib "GDI32" _
 (ByVal fnPenStyle As Long, ByVal nWidth As Long, _
 ByVal crColor As Long) As Long
Declare Function CreateHatchBrush Lib "GDI32" _
 (ByVal fnStyle As Long, ByVal clrref As Long) As Long
Declare Function CreateCompatibleDC Lib "GDI32" _
 (ByVal hdc As Long) As Long
Declare Function WindowFromDC Lib "USER32" _
 (ByVal hdc As Long) As Long
Declare Function BitBlt Lib "GDI32" _
 (ByVal hdcDest As Long, ByVal nXDest As Long, _
 ByVal nYDest As Long, ByVal nWidth As Long, _
 ByVal nHeight As Long, ByVal hdcSrc As Long, _
 ByVal nXSrc As Long, ByVal nYSrc As Long, _
 ByVal dwRop As Long) As Long
Declare Function SaveDC Lib "GDI32" _
 (ByVal hdc As Long) As Long
Declare Function RestoreDC Lib "GDI32" _
 (ByVal hdc As Long, ByVal nSavedDC As Long) As Long
Declare Function GetStockObject Lib "GDI32" _
 (ByVal fnObject As Long) As Long
Declare Function SetBkMode Lib "GDI32" _
 (ByVal hdc As Long, ByVal iBkMode As Long) As Long
Declare Function ExtTextOutW Lib "GDI32" _
 (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, _
 ByVal fuOptions As Long, lprc As Rect, lpString As String, _
 ByVal cbCount As Long, ByVal lpDx As Long) As Long
Declare Function ExtTextOutA Lib "GDI32" _
 (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, _
 ByVal fuOptions As Long, lprc As Rect, lpString As String, _
 ByVal cbCount As Long, ByVal lpDx As Long) As Long
Declare Function CreateCompatibleBitmap Lib "GDI32" _
 (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function SetMapMode Lib "GDI32" _
 (ByVal hdc As Long, ByVal fnMapMode As Long) As Long
Declare Function SetWindowExtEx Lib "GDI32" _
 (ByVal hdc As Long, ByVal nXExtent As Long, ByVal nYExtent As Long, _
 ByVal lpSize As Long) As Long
Declare Function SetWindowOrgEx Lib "GDI32" _
 (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, _
 ByVal lpPoint As Long) As Long
Declare Function GetClientRect Lib "USER32" _
 (ByVal hdc As Long, lpRect As Rect) As Long
Declare Function SetViewportExtEx Lib "GDI32" _
 (ByVal hdc As Long, ByVal nXExtent As Long, ByVal nYExtent As Long, _
 ByVal lpSize As Long) As Long
Declare Function Ellipse Lib "GDI32" _
 (ByVal hdc As Long, ByVal nLeftRect As Long, ByVal nTopRect As Long, _
 ByVal nRightRect As Long, ByVal nBottomRect As Long) As Long
Declare Function CreateFontW Lib "GDI32" _
 (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, _
 ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, _
 ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, _
 ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, _
 ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Declare Function CreateFontA Lib "GDI32" _
 (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, _
 ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, _
 ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, _
 ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, _
 ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Declare Function GetTextMetricsW Lib "GDI32" _
 (ByVal hdc As Long, lptm As TEXTMETRIC) As Long
Declare Function GetTextMetricsA Lib "GDI32" _
 (ByVal hdc As Long, lptm As TEXTMETRIC) As Long
Declare Function SetGraphicsMode Lib "GDI32" _
 (ByVal hdc As Long, ByVal iMode As Long) As Long
Declare Function SetWorldTransform Lib "GDI32" _
 (ByVal hdc As Long, lpXform As XFORM) As Long
Declare Function ModifyWorldTransform Lib "GDI32" _
 (ByVal hdc As Long, pXform As XFORM, ByVal iMode As Long) As Long
Declare Function TextOutW Lib "GDI32" _
 (ByVal hdc As Long, ByVal nXStart As Long, ByVal nYStart As Long, _
 ByVal lpString As String, ByVal cbString As Long) As Long
Declare Function TextOutA Lib "GDI32" _
 (ByVal hdc As Long, ByVal nXStart As Long, ByVal nYStart As Long, _
 ByVal lpString As String, ByVal cbString As Long) As Long
Declare Function BeginPath Lib "GDI32" _
 (ByVal hdc As Long) As Long
Declare Function EndPath Lib "GDI32" _
 (ByVal hdc As Long) As Long
Declare Function StrokeAndFillPath Lib "GDI32" _
 (ByVal hdc As Long) As Long
Declare Function MoveToEx Lib "GDI32" _
 (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, _
 ByVal lpPoint As Long) As Long
Declare Function LineTo Lib "GDI32" _
 (ByVal hdc As Long, ByVal nXEnd As Long, ByVal nYEnd As Long) As Long
Declare Function SetArcDirection Lib "GDI32" _
 (ByVal hdc As Long, ByVal ArcDirection As Long) As Long
Declare Function ArcTo Lib "GDI32" _
 (ByVal hdc As Long, ByVal nLeftRect As Long, _
 ByVal nTopRect As Long, ByVal nRightRect As Long, _
 ByVal nBottomRect As Long, ByVal nXRadial1 As Long, _
 ByVal nYRadial1 As Long, ByVal nXRadial2 As Long, _
 ByVal nYRadial2 As Long) As Long
Declare Function Arc Lib "GDI32" _
 (ByVal hdc As Long, ByVal nLeftRect As Long, _
 ByVal nTopRect As Long, ByVal nRightRect As Long, _
 ByVal nBottomRect As Long, ByVal nXRadial1 As Long, _
 ByVal nYRadial1 As Long, ByVal nXRadial2 As Long, _
 ByVal nYRadial2 As Long) As Long
Declare Function GetPath Lib "GDI32" _
 (ByVal hdc As Long, lpPoints As Point, lpTypes As Byte, _
 ByVal nSize As Long) As Long
Declare Function FillPath Lib "GDI32" _
 (ByVal hdc As Long) As Long
Declare Function PolyDraw Lib "GDI32" _
 (ByVal hdc As Long, lppt As Point, _
 ByVal lpbTypes As Byte, ByVal cCount As Long) As Long
Declare Function CloseFigure Lib "GDI32" _
 (ByVal hdc As Long) As Long
Declare Function PolyBezierTo Lib "GDI32" _
 (ByVal hdc As Long, lppt As Point, _
 ByVal cCount As Long) As Long
Declare Function GetTextColor Lib "GDI32" _
 (ByVal hdc As Long) As Long
Declare Function SetTextAlign Lib "GDI32" _
 (ByVal hdc As Long, ByVal fMode As Long) As Long
Declare Function LoadCursorW Lib "USER32" _
 (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Declare Function LoadCursorA Lib "USER32" _
 (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Declare Function SetCursor Lib "USER32" _
 (ByVal hCursor As Long) As Long

Function GetVersionEx(lpVersionInformation As OSVERSIONINFO) As Long
#If Unicode Then
    GetVersionEx = GetVersionExW(lpVersionInformation)
#Else
    GetVersionEx = GetVersionExA(lpVersionInformation)
#End If
End Function

Function ExtTextOut(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, _
                    ByVal fuOptions As Long, lprc As Rect, lpString As String, _
                    ByVal cbCount As Long, ByVal lpDx As Long) As Long
#If Unicode Then
    ExtTextOut = ExtTextOutW(hdc, X, Y, fuOptions, lprc, lpString, cbCount, lpDx)
#Else
    ExtTextOut = ExtTextOutA(hdc, X, Y, fuOptions, lprc, lpString, cbCount, lpDx)
#End If
End Function

Function CreateFont(ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, _
                    ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, _
                    ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, _
                    ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, _
                    ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
#If Unicode Then
    CreateFont = CreateFontW(nHeight, nWidth, nEscapement, nOrientation, fnWeight, fdwItalic, _
                            fdwUnderline, fdwStrikeOut, fdwCharSet, fdwOutputPrecision, _
                            fdwClipPrecision, fdwQuality, fdwPitchAndFamily, lpszFace)
#Else
    CreateFont = CreateFontA(nHeight, nWidth, nEscapement, nOrientation, fnWeight, fdwItalic, _
                            fdwUnderline, fdwStrikeOut, fdwCharSet, fdwOutputPrecision, _
                            fdwClipPrecision, fdwQuality, fdwPitchAndFamily, lpszFace)
#End If
End Function

Function GetTextMetrics(ByVal hdc As Long, lptm As TEXTMETRIC) As Long
#If Unicode Then
    GetTextMetrics = GetTextMetricsW(hdc, lptm)
#Else
    GetTextMetrics = GetTextMetricsA(hdc, lptm)
#End If
End Function

Function TextOut(ByVal hdc As Long, ByVal nXStart As Long, ByVal nYStart As Long, _
                ByVal lpString As String, ByVal cbString As Long) As Long
#If Unicode Then
    TextOut = TextOutW(hdc, nXStart, nYStart, lpString, cbString)
#Else
    TextOut = TextOutA(hdc, nXStart, nYStart, lpString, cbString)
#End If
End Function
                    
Function LoadCursor(ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
#If Unicode Then
    LoadCursor = LoadCursorW(hInstance, lpCursorName)
#Else
    LoadCursor = LoadCursorA(hInstance, lpCursorName)
#End If
End Function

' This function detects the Windows platform the program is running on.
Function IsWindowsNT() As Boolean
    Dim verInfo As OSVERSIONINFO
    verInfo.dwOSVersionInfoSize = Len(verInfo) - 1
    Const VER_PLATFORM_WIN32_NT = 2
    If GetVersionEx(verInfo) <> 0 Then
        IsWindowsNT = (verInfo.dwPlatformId = VER_PLATFORM_WIN32_NT)
    Else
        IsWindowsNT = False
    End If
End Function

' This function returns the Picture object for the program to work with.
Function GetPicture() As Picture
    Set GetPicture = pictureObject
End Function
