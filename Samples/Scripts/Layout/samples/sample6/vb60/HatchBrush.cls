VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HatchBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//////////////////////////////////////////////////////////////////////////////////////
'//
'// HatchBrush : Windows GDI brush representation class
'//
'//////////////////////////////////////////////////////////////////////////////////////
'// This is a part of the Innoveda-PowerPCB OLE Automation server PPCBAutoGeom sample.
'// Copyright (C) 2003 Mentor Graphics Corp.
'// All rights reserved.
'//
'// This source code is only intended as a supplement to the Innoveda-PowerPCB OLE
'// Automation Server API Help file.
'//////////////////////////////////////////////////////////////////////////////////////

' HatchBrush: this class encapsulates a Windows graphics device interface (GDI) brush
' and provides automatic device context cleanup after GDI object selection.

Private hdc As Long ' Windows device context handle
Private hOldObject As Long ' Windows GDI object handle

Private Sub Class_Initialize()
    hdc = 0
    hOldObject = 0
End Sub

Private Sub Class_Terminate()
    If hdc <> 0 And hOldObject <> 0 Then
        SelectObject hdc, hOldObject
    End If
End Sub

Public Property Let dc(ByVal aHDC As Long)
    hdc = aHDC
End Property

Public Property Let oldObject(ByVal aOldObject As Long)
    hOldObject = aOldObject
End Property

