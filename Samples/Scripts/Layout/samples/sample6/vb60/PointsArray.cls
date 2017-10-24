VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PointsArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//////////////////////////////////////////////////////////////////////////////////////
'//
'// PointsArray : PowerPCB Automation points array representation class
'//
'//////////////////////////////////////////////////////////////////////////////////////
'// This is a part of the Innoveda-PowerPCB OLE Automation server PPCBAutoGeom sample.
'// Copyright (C) 2003 Mentor Graphics Corp.
'// All rights reserved.
'//
'// This source code is only intended as a supplement to the Innoveda-PowerPCB OLE
'// Automation Server API Help file.
'//////////////////////////////////////////////////////////////////////////////////////

' PointsArray: this class facilitates data access for Automation arrays of points.

Option Explicit ' make sure enum values are respected

Private points As Variant ' Automation array

' column structure indexes
Public Enum Indexes
    pointX = 1
    pointY = 2
    routeDim2 = 2
    angle = 3
    polylineDim2 = 3
    numDims = 2
End Enum

Public Property Let value(arrayVar As Variant)
    points = arrayVar
End Property

Public Function Check(ByVal dim2 As Long) As Boolean
    ' array dimensions bounds (first (lower), second (lower,upper))
    Dim lBound1, lBound2, uBound2 As Long
    
    lBound1 = LBound(points, 1)
    lBound2 = LBound(points, 2)
    uBound2 = UBound(points, 2)
    
    If VarType(points) <> (vbArray + vbDouble) Or lBound1 <> 1 Or lBound2 <> 1 Or uBound2 <> dim2 Then
        MsgBox LoadResString(IDS_BADVARIANT), vbCritical
        Check = False
    End If
    Check = True
End Function

Public Function GetSize() As Long
    Dim uBound1 As Long
    uBound1 = UBound(points, 1)
    GetSize = uBound1
End Function

Public Function GetPoint(ByVal corner As Long) As CPoint
    Set GetPoint = New CPoint
    GetPoint.X = GetElement(corner, pointX)
    GetPoint.Y = GetElement(corner, pointY)
End Function

Public Function GetAngle(ByVal corner As Long) As Long
    GetAngle = GetElement(corner, angle)
End Function

Private Function GetElement(ByVal i As Long, ByVal j As Long) As Double
    GetElement = points(i, j)
End Function
