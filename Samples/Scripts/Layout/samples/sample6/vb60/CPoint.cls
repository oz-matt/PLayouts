VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//////////////////////////////////////////////////////////////////////////////////////
'//
'// CPoint : Windows POINT representation class
'//
'//////////////////////////////////////////////////////////////////////////////////////
'// This is a part of the Innoveda-PowerPCB OLE Automation server PPCBAutoGeom sample.
'// Copyright (C) 2003 Mentor Graphics Corp.
'// All rights reserved.
'//
'// This source code is only intended as a supplement to the Innoveda-PowerPCB OLE
'// Automation Server API Help file.
'//////////////////////////////////////////////////////////////////////////////////////

Private data As Point

Public Sub Init(ByVal ax As Long, ByVal ay As Long)
    X = ax
    Y = ay
End Sub

Public Property Let X(ByVal ax As Long)
    data.X = ax
End Property

Public Property Get X() As Long
    X = data.X
End Property

Public Property Let Y(ByVal ay As Long)
    data.Y = ay
End Property

Public Property Get Y() As Long
    Y = data.Y
End Property

Public Function Subtract(ByVal value As CPoint) As CPoint
    Set Subtract = New CPoint
    Subtract.X = X - value.X
    Subtract.Y = Y - value.Y
End Function

Public Function Add(ByVal value As CPoint) As CPoint
    Set Add = New CPoint
    Add.X = X + value.X
    Add.Y = Y + value.Y
End Function

