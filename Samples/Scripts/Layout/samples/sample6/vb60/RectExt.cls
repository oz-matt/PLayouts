VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RectExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//////////////////////////////////////////////////////////////////////////////////////
'//
'// RectExt : custom Windows RECT representation class
'//
'//////////////////////////////////////////////////////////////////////////////////////
'// This is a part of the Innoveda-PowerPCB OLE Automation server PPCBAutoGeom sample.
'// Copyright (C) 2003 Mentor Graphics Corp.
'// All rights reserved.
'//
'// This source code is only intended as a supplement to the Innoveda-PowerPCB OLE
'// Automation Server API Help file.
'//////////////////////////////////////////////////////////////////////////////////////

' RectExt: this class facilitates rectangle (square) definition
' based on the center point and radius of the circle inscribed.

Private rct As Rect

' Initialize the rectangle.
Public Sub Init(ByVal X As Double, ByVal Y As Double, ByVal radius As Double)
    left = X - radius
    top = Y - radius
    right = X + radius
    bottom = Y + radius
End Sub

Property Let left(ByVal val As Long)
    rct.left = val
End Property

Property Get left() As Long
    left = rct.left
End Property

Property Let top(ByVal val As Long)
    rct.top = val
End Property

Property Get top() As Long
    top = rct.top
End Property

Property Let right(ByVal val As Long)
    rct.right = val
End Property

Property Get right() As Long
    right = rct.right
End Property

Property Let bottom(ByVal val As Long)
    rct.bottom = val
End Property

Property Get bottom() As Long
    bottom = rct.bottom
End Property

