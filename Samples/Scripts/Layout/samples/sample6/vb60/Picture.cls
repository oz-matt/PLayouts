VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Picture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//////////////////////////////////////////////////////////////////////////////////////
'//
'// Picture : picture storage class
'//
'//////////////////////////////////////////////////////////////////////////////////////
'// This is a part of the Innoveda-PowerPCB OLE Automation server PPCBAutoGeom sample.
'// Copyright (C) 2003 Mentor Graphics Corp.
'// All rights reserved.
'//
'// This source code is only intended as a supplement to the Innoveda-PowerPCB OLE
'// Automation Server API Help file.
'//////////////////////////////////////////////////////////////////////////////////////

' Picture: this class stores screen image as a bitmap.

Private bmp As Long ' picture bitmap handle
Private winNT As Boolean ' current Windows is NT/2000

Private Sub Class_Initialize()
    CleanUp
    winNT = IsWindowsNT()
End Sub

Private Sub Class_Terminate()
    CleanUp
End Sub

Private Sub CleanUp()
    If bmp <> 0 Then
        DeleteObject bmp
        bmp = 0
    End If
End Sub

' Build a picture bitmap based on Automation-retrieved data.
Sub Refresh()
    CleanUp ' discard what we may have had.

    Dim dc As Long
    dc = GetDC(FormMain.hWnd)   ' to be "Release"d!

    ' Let's create a full-screen bitmap to allow to increase window's size without picture rebuild, if desired.
    Const SM_CXFULLSCREEN = 16
    Const SM_CYFULLSCREEN = 17
    Dim width, height As Long
    width = GetSystemMetrics(SM_CXFULLSCREEN)
    height = GetSystemMetrics(SM_CYFULLSCREEN)

    ' One may wonder, why we re-create a full-screen bitmap instead of reusing it.
    ' Well, screen resolution might have been increased...
    bmp = CreateCompatibleBitmap(dc, width, height)

    If bmp Then
        Dim autoDrawObj As New AutoDraw
        autoDrawObj.Init winNT
        autoDrawObj.Refresh dc, bmp, width, height
    Else  ' if there's a trouble, report it
        MsgBox LoadResString(IDS_PICTUREBUILDFAILED), vbExclamation
    End If

    If dc <> Null Then
        ReleaseDC FormMain.hWnd, dc ' prevent a GDI resource leak
    End If
End Sub

' draw the picture in a window
Sub Draw(hdc As Long)
    ' create a DC for stored bitmap
    Dim memDc As Long
    memDc = CreateCompatibleDC(hdc)
    If bmp <> 0 And memDc <> 0 Then
        Dim oldBmp As Long
        oldBmp = SelectObject(memDc, bmp)

        Dim clientRect As Rect
        GetClientRect WindowFromDC(hdc), clientRect ' if there's no window, there's nothing to trust in

        ' let's (proudly?) show our bitmap
        Const SRCCOPY = &HCC0020
        BitBlt hdc, 0, 0, clientRect.right, clientRect.bottom, memDc, 0, 0, SRCCOPY

        SelectObject memDc, oldBmp
    End If
End Sub

