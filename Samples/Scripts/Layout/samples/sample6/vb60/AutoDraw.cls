VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AutoDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//////////////////////////////////////////////////////////////////////////////////////
'//
'// AutoDraw : picture building class
'//
'//////////////////////////////////////////////////////////////////////////////////////
'// This is a part of the Innoveda-PowerPCB OLE Automation server PPCBAutoGeom sample.
'// Copyright (C) 2003 Mentor Graphics Corp.
'// All rights reserved.
'//
'// This source code is only intended as a supplement to the Innoveda-PowerPCB OLE
'// Automation Server API Help file.
'//////////////////////////////////////////////////////////////////////////////////////

' PowerPCB Automation enums used

Private Enum PPcbObjectType
    ppcbObjectTypeUnknown = 0
    ppcbObjectTypeComponent = 1
    ppcbObjectTypeNet = 2
    ppcbObjectTypePin = 3
    ppcbObjectTypeVia = 4
    ppcbObjectTypeConnection = 5
    ppcbObjectTypeRouteSegment = 6
    ppcbObjectTypeJumper = 7
    ppcbObjectTypePartType = 8
    ppcbObjectTypeCBP = 9
    ppcbObjectTypeSBP = 10
    ppcbObjectTypeWirebond = 11
    ppcbObjectTypeNetClass = 12
    ppcbObjectTypeDrawing = 13
    ppcbObjectTypeText = 14
    ppcbObjectTypeLabel = 15
    ppcbObjectTypePolyline = 16
    ppcbObjectTypeCircle = 17
    ppcbObjectTypeLibrary = 18
    ppcbObjectTypeLibraryItem = 19
    ppcbObjectTypeAll = 9999
    ppcbObjectTypeApplication = 20
    ppcbObjectTypeAttribute = 21
    ppcbObjectTypeAttributeType = 22
    ppcbObjectTypeDocument = 23
    ppcbObjectTypeMeasure = 24
    ppcbObjectTypeView = 25
    ppcbObjectTypeAssemblyOptions = 26
    ppcbObjectTypeAttributes = 27
    ppcbObjectTypeAttributeTypes = 28
    ppcbObjectTypeObjects = 29
End Enum

Private Enum PPcbHorizontalJustification
    ppcbJustifyLeft = 0
    ppcbJustifyHCenter = 1
    ppcbJustifyRight = 2
End Enum

Private Enum PPcbVerticalJustification
    ppcbJustifyTop = 0
    ppcbJustifyVCenter = 1
    ppcbJustifyBottom = 2
End Enum

Private Enum PPcbRightReadingStatus
    ppcbRightReadingNone = 0
    ppcbRightReadingOrthogonal = 1
    ppcbRightReadingAngled = 2
End Enum

Private Enum PPcbShapeType
    ppcbShapeTypeOpen = 0
    ppcbShapeTypeHollow = 1
    ppcbShapeTypeFilled = 2
    ppcbShapeTypeVoid = 3
End Enum

Private Enum PPcbUnit
    ppcbUnitCurrent = 0
    ppcbUnitDatabase = 1
    ppcbUnitMils = 2
    ppcbUnitInch = 3
    ppcbUnitMetric = 4
End Enum

Private Enum PPcbOriginType
    ppcbOriginTypeDesign = 0
    ppcbOriginTypeParent = 1
End Enum

Private Enum DrawRouteSegment_Constants
    startPnt = 1
    endPnt = 2
    midPnt = 3
    segmentSize = 2
    arcSize = 3
End Enum

Option Explicit ' make sure enum values are respected

' class properties

Private hdc As Long ' Windows device context handle to use
Private winNT As Boolean ' current Windows is NT/2000
Private doc As Object ' we need a PowerPCB Automation document to query it's contents
Private unit, origin As Long  ' PPCB Automation parameters values to use
Private no_powerpcb As Boolean
Private height As Long
' The following data members are needed only to support operations under Win95-98.
' They are not used under NT.
Private graphicsMode, ArcDirection As Long
Private xfrm As XFORM

Private Sub Class_Initialize()
    unit = ppcbUnitMils
    origin = ppcbOriginTypeDesign
    no_powerpcb = False
End Sub

Public Sub Init(ByVal aWinNT As Boolean)
    Const GM_COMPATIBLE = 1
    Const AD_COUNTERCLOCKWISE = 1
    winNT = aWinNT
    graphicsMode = GM_COMPATIBLE
    ArcDirection = AD_COUNTERCLOCKWISE
    SetIdentityTrasnform
End Sub

' refresh the drawing using Automation data
Public Sub Refresh(ByVal aHDC As Long, ByVal bmp As Long, ByVal width As Long, ByVal height As Long)
    no_powerpcb = False
    hdc = CreateCompatibleDC(aHDC)
    If hdc Then

        Dim index As Long
        index = SaveDC(hdc)

        Dim hBrush As Long
        Const HOLLOW_BRUSH = 5
        hBrush = GetStockObject(HOLLOW_BRUSH)

        If index Then
            SelectObject hdc, bmp
            SelectObject hdc, hBrush
            Const TRANSPARENT = 1
            SetBkMode hdc, TRANSPARENT

            ' paint the background
            Dim rct As Rect
            rct.left = 0
            rct.top = 0
            rct.right = width
            rct.bottom = height
            Dim str As String
            Const ETO_OPAQUE = &H2
            ExtTextOut hdc, 0, 0, ETO_OPAQUE, rct, str, 0, 0
            
            On Error Resume Next
            
            DoDraw
            
            If Err.Number <> 0 Then
                If no_powerpcb Then
                    MsgBox "Open a design in PowerPCB first."
                Else
                    MsgBox Err.Description
                End If
            End If

            RestoreDC hdc, index
        End If
    End If
End Sub

Private Sub SelectPen(pen As SolidPen, ByVal width As Double)
    Dim hPen As Long
    Const PS_SOLID = 0
    hPen = CreatePen(PS_SOLID, CLng(width), GetTextColor(hdc))
    pen.dc = hdc
    pen.oldObject = SelectObject(hdc, hPen)
End Sub

Private Sub SelectBrush(brush As HatchBrush)
    Dim hBrush As Long
    Const HS_DIAGCROSS = 5
    hBrush = CreateHatchBrush(HS_DIAGCROSS, GetTextColor(hdc))
    brush.dc = hdc
    brush.oldObject = SelectObject(hdc, hBrush)
End Sub

' refresh the drawing using Automation data
Private Sub DoDraw()

    ConnectToPowerPCB
    
    Dim view As Object
    Set view = doc.ActiveView() ' the picture we want is inside a view
    
    Const MM_ISOTROPIC = 7
    SetMapMode hdc, MM_ISOTROPIC

    ' we'll need it to refine picture's placement
    Dim topLeft As New CPoint
    Dim bottomRight As New CPoint
    Dim addValue As New CPoint
    topLeft.Init view.TopLeftX(unit), view.TopLeftY(unit)
    bottomRight.Init view.BottomRightX(unit), view.BottomRightY(unit)
    addValue.Init 1, 0

    ' set picture display parameters (units conversion, axes orientation, position)
    Dim winExt As CPoint
    Set winExt = bottomRight.Subtract(topLeft).Add(addValue)
    SetWindowExtEx hdc, winExt.X, winExt.Y, 0
    SetWindowOrgEx hdc, topLeft.X, topLeft.Y, 0  ' let's respect view's (PowerPCB's window scrolling) current position

    Dim clientRect As Rect
    GetClientRect FormMain.hWnd, clientRect
    SetViewportExtEx hdc, clientRect.right, clientRect.bottom, 0

    ' Drawings
    DrawObjects ppcbObjectTypeDrawing
    ' RouteSegments
    DrawObjects ppcbObjectTypeRouteSegment
    ' Pins
    DrawObjects ppcbObjectTypePin
    ' Vias
    DrawObjects ppcbObjectTypeVia
    ' Texts
    DrawObjects ppcbObjectTypeText
    ' Labels
    DrawObjects ppcbObjectTypeLabel
    
End Sub

' connect to PowerPCB and initialize the document object
Private Sub ConnectToPowerPCB()
    Dim app As Object
    On Error Resume Next
    Set app = GetObject(Class:="powerpcb.application")
    If Err.Number <> 0 Then
        no_powerpcb = True
        Exit Sub
    End If
    Set doc = app.ActiveDocument()
End Sub

' extract 'pcbObjectType' items from "objects" and draw them
Private Sub DrawObjects(ByVal pcbObjectType As Long)
    DoDrawObjects pcbObjectType, doc.GetObjects(pcbObjectType, "", False)
End Sub

' select a method to draw a typed object
Private Sub DoDrawObjects(ByVal pcbObjectType As Long, objects As Object)
    Dim var As Variant
    
    Dim i As Long
    Dim count As Long
    count = objects.count()
    
    ' for all objects in the collection
    For i = 1 To count
        var = i
        If objects.ItemType(var) = pcbObjectType Then
            Select Case pcbObjectType
                Case ppcbObjectTypeCircle
                    DrawCircle objects.Item(var)
                Case ppcbObjectTypeDrawing
                    DrawDrawing objects.Item(var)
                Case ppcbObjectTypeLabel
                    DrawLabel objects.Item(var)
                Case ppcbObjectTypePolyline
                    DrawPolyline objects.Item(var)
                Case ppcbObjectTypeRouteSegment
                    DrawRouteSegment objects.Item(var)
                Case ppcbObjectTypeText
                    DrawText objects.Item(var)
                Case ppcbObjectTypePin
                    DrawPin objects.Item(var)
                Case ppcbObjectTypeVia
                    DrawVia objects.Item(var)
            End Select
        End If
    Next i
End Sub

' extract data from "Circle" PowerPCB Automation object and draw
Private Sub DrawCircle(aCircle As Object)
    Dim brush As New HatchBrush
    If IsShapeFilled(aCircle.ShapeType) Then
        SelectBrush brush
    End If

    Dim pen As New SolidPen
    SelectPen pen, aCircle.LineWidth(unit)

    Dim rct As New RectExt
    rct.Init aCircle.CenterX(unit, origin), aCircle.CenterY(unit, origin), aCircle.radius(unit)
    Ellipse hdc, rct.left, rct.top, rct.right, rct.bottom
End Sub

' extract data from "Drawing" PowerPCB Automation object and draw
Private Sub DrawDrawing(aDrawing As Object)
    Dim geometry As Object
    Set geometry = aDrawing.geometry()
    DoDrawObjects ppcbObjectTypeCircle, geometry
    DoDrawObjects ppcbObjectTypePolyline, geometry
End Sub

' extract data from "Label" PowerPCB Automation object and draw
Private Sub DrawLabel(aLabel As Object)
    DoDrawText aLabel.text(), CLng(aLabel.PositionX(unit, origin)), CLng(aLabel.PositionY(unit, origin)), _
                CLng(aLabel.height(unit)), aLabel.HorzJustification(), aLabel.VertJustification(), _
                CLng(aLabel.orientation(origin)), aLabel.Mirror(origin), "Arial", aLabel.RightReading()
End Sub

' Polyline segments drawing helper: draws all segments
Private Sub DrawSegments(polyline As Object)
    Dim var As Variant
    var = polyline.points(unit, origin)
    Dim points As New PointsArray
    points.value = var

    MoveToEx hdc, points.GetPoint(1).X, points.GetPoint(1).Y, 0
    Dim corner As Long
    For corner = 2 To points.GetSize()
        Dim dest As CPoint
        Set dest = points.GetPoint(corner)
        Dim angle As Double
        angle = points.GetAngle(corner - 1)
        If angle <> 0 Then ' it's not over yet
            Dim start As CPoint
            Set start = points.GetPoint(corner - 1)
            Dim rct As New RectExt
            rct.Init polyline.CenterX(corner - 1, unit, origin), _
                     polyline.CenterY(corner - 1, unit, origin), _
                     polyline.radius(corner - 1, unit)
            Const AD_COUNTERCLOCKWISE = 1
            Const AD_CLOCKWISE = 2
            If angle > 0# Then
                SetArcDirectionCustom AD_COUNTERCLOCKWISE
            Else
                SetArcDirectionCustom AD_CLOCKWISE
            End If
            ArcToCustom rct, start, dest
        Else
            LineTo hdc, dest.X, dest.Y
        End If
    Next corner
End Sub

' extract data from "Polyline" PowerPCB Automation object and draw
Private Sub DrawPolyline(aPolyline As Object)
    Dim pen As New SolidPen
    SelectPen pen, aPolyline.LineWidth(unit)

    If IsShapeFilled(aPolyline.ShapeType()) Then
        ' fill the shape
        BeginPath hdc
        DrawSegments aPolyline
        EndPath hdc
        Dim brush As New HatchBrush
        SelectBrush brush
        StrokeAndFillPath hdc
    Else
        DrawSegments aPolyline
    End If
End Sub

' check if both r =- r and d + r fit into "long" data type
Private Function IsConvertibleToLong(ByVal d As Double, ByVal r As Double) As Boolean
    Const LONG_MAX = 2147483647
    Const LONG_MIN = -2147483647 - 1
    IsConvertibleToLong = (CDbl(LONG_MAX) >= d + r) And (d - r >= CDbl(LONG_MIN))
End Function

' Calculate arc rectangle (actually, square) by three points
' returns:
'      true  if rectangle can be calculated
'      false if rectangle cannot be calculated
'          (all points on a line, or the result does not fit into RECT)
'      rct - calculated arc rectangle
'      ccw - calcucated arc direction
'      start - start point on the circle
'      interm - intermediate point on the circle
'      end - last point on the circle

Private Function arcRect(rct As RectExt, ccw As Boolean, startP As CPoint, _
                        intermP As CPoint, endP As CPoint) As Boolean
    ' a = {ax, ay} is vector from middle point to start point
    Dim ax As Double
    Dim ay As Double
    ax = startP.X - intermP.X
    ay = startP.Y - intermP.Y
    ' c = {cx, cy} is vector from middle point to end point
    Dim cx As Double
    Dim cy As Double
    cx = endP.X - intermP.X
    cy = endP.Y - intermP.Y
    Dim aa As Double
    Dim cc As Double
    aa = ax * ax + ay * ay ' length of vector "a" squared
    cc = cx * cx + cy * cy ' length of vector "c" squared
    ' ac is doubled cross product "a x c" or four times triangle area (signed)
    Dim ac As Double
    ac = 2# * (ax * cy - ay * cx)

    ' if ac is zero, the points are on a line (not on arc)
    If ac <> 0# Then
        ' {rx, ry} is vertor from middle point to arc center
        Dim rx As Double
        Dim ry As Double
        rx = (cy * aa - ay * cc) / ac
        ry = (ax * cc - cx * aa) / ac
        ' length of vertor "r" is the arc radius
        Dim rad As Double
        rad = Sqr(rx * rx + ry * ry)

        rx = rx + intermP.X
        ry = ry + intermP.Y
        ' now r is the arc center position

        ' check if arc rectangle fit into "long" RECT fields
        If Not IsConvertibleToLong(rx, rad) Then
            arcRect = False
            Exit Function
        End If
        If Not IsConvertibleToLong(ry, rad) Then
            arcRect = False
            Exit Function
        End If

        ' construct the arc rectangle
        rct.Init rx, ry, rad
        ' sign of ac defines arc direction
        ccw = (ac < 0#)
        arcRect = True
    Else
        arcRect = False
    End If
End Function

' extract data from "RouteSegment" PowerPCB Automation object and draw
Private Sub DrawRouteSegment(aRouteSegment As Object)
    Dim var As Variant
    var = aRouteSegment.points(unit)
    Dim points As New PointsArray
    points.value = var

    Dim pen As New SolidPen
    SelectPen pen, aRouteSegment.width(unit)
    
    Dim size As Long
    size = points.GetSize()
    If size >= segmentSize Then
        Dim startP As CPoint
        Dim endP As CPoint
        Set startP = points.GetPoint(startPnt)
        Set endP = points.GetPoint(endPnt)
        MoveToEx hdc, startP.X, startP.Y, 0 ' set starting point

        If size = segmentSize Then
            ' 2 points - draw a line
            LineTo hdc, endP.X, endP.Y
        Else
            If size = arcSize Then
                ' 3 points - draw an arc
                Dim intermP As CPoint
                Set intermP = points.GetPoint(midPnt)
                Dim arcRct As New RectExt
                Dim ccw As Boolean
                ' calculate arc rectangle and direction by three points
                If arcRect(arcRct, ccw, startP, intermP, endP) Then
                    If ccw Then
                        Const AD_COUNTERCLOCKWISE = 1
                        SetArcDirectionCustom AD_COUNTERCLOCKWISE
                    Else
                        Const AD_CLOCKWISE = 2
                        SetArcDirectionCustom AD_CLOCKWISE
                    End If
                    ArcToCustom arcRct, startP, endP
                Else
                    MsgBox LoadResString(IDS_BADROUTESEGMENTARC), vbExclamation
                End If
            Else
                MsgBox LoadResString(IDS_BADVARIANT), vbCritical
            End If
        End If
    End If
End Sub

' extract data from "Text" PowerPCB Automation object and draw
Private Sub DrawText(aText As Object)
    DoDrawText aText.text(), CLng(aText.PositionX(unit, origin)), CLng(aText.PositionY(unit, origin)), _
                CLng(aText.height(unit)), aText.HorzJustification(), aText.VertJustification(), _
                CLng(aText.orientation(origin)), aText.Mirror(origin), "Times New Roman", ppcbRightReadingNone
End Sub

' extract data from "Pin" PowerPCB Automation object and draw
Private Sub DrawPin(aPin As Object)
    Dim rct As New RectExt
    rct.Init aPin.PositionX(unit), aPin.PositionY(unit), aPin.DrillSize(unit) / 2#
    Ellipse hdc, rct.left, rct.top, rct.right, rct.bottom
End Sub

' extract data from "Via" PowerPCB Automation object and draw
Private Sub DrawVia(aVia As Object)
    Dim rct As New RectExt
    rct.Init aVia.PositionX(unit), aVia.PositionY(unit), aVia.DrillSize(unit) / 2#
    Ellipse hdc, rct.left, rct.top, rct.right, rct.bottom
End Sub

' Draw text (for "Test" or "Label") object in bitmap
' using parameters common for both objects.
Private Sub DoDrawText(text As String, ByVal X As Long, ByVal Y As Long, ByVal aHeight As Long, _
                        ByVal horzJust As PPcbHorizontalJustification, ByVal vertJust As PPcbVerticalJustification, _
                        ByVal aOrientation As Long, ByVal mirrored As Boolean, fontName As String, _
                        ByVal rrs As PPcbRightReadingStatus)
    ' skip empty texts
    If text = "" Then Exit Sub
    
    ' ensure the orientation is in 0 - 359 range
    Dim orientation As Long
    orientation = aOrientation - 360 * (aOrientation \ 360)
    If orientation < 0 Then orientation = orientation + 360

    ' orientation ajdusted to take "Right Reading" into account
    Dim adjustedOrientation As Long
    adjustedOrientation = orientation

    If rrs = ppcbRightReadingAngled Then
        If orientation > 90 And orientation < 270 Then
            orientation = orientation + 180
        End If
    Else
        If rrs = ppcbRightReadingOrthogonal Then
            Dim slack As Long
            slack = orientation - 180 * (orientation \ 180)
            If slack > 45 And slack < 135 Then
                orientation = 90
            Else
                orientation = 0
            End If
            adjustedOrientation = orientation
        End If
    End If

    ' create font in left-to-right (zero) orientation
    Dim hFont As Long
    Const FW_NORMAL = 400
    Const ANSI_CHARSET = 0
    Const OUT_DEFAULT_PRECIS = 0
    Const CLIP_LH_ANGLES = &H10
    Const CLIP_DEFAULT_PRECIS = 0
    Const DEFAULT_QUALITY = 0
    Const DEFAULT_PITCH = 0
    Const FF_DONTCARE = &H0
    hFont = CreateFont(-aHeight, 0, 0, 0, _
                        FW_NORMAL, False, False, False, ANSI_CHARSET, _
                        OUT_DEFAULT_PRECIS, CLIP_LH_ANGLES Or CLIP_DEFAULT_PRECIS, _
                        DEFAULT_QUALITY, DEFAULT_PITCH Or FF_DONTCARE, fontName)

    Dim hOldFont As Long
    hOldFont = SelectObject(hdc, hFont)

    ' PowerPCB vertical justification does not match Win32 GDI types.
    ' y (vertical) position is adjusted to place desired text point at y = 0 line.
    Dim offset As Long
    offset = 0
    Dim tm As TEXTMETRIC
    GetTextMetrics hdc, tm
    Select Case vertJust
        Case ppcbJustifyTop
            offset = -tm.tmInternalLeading
        Case ppcbJustifyVCenter
            offset = -(tm.tmAscent + tm.tmInternalLeading) / 2
        Case ppcbJustifyBottom
            offset = -tm.tmAscent
    End Select

    ' horizontal justification is the same as Win32 setting
    Const TA_TOP = 0
    Const TA_NOUPDATECP = 0
    Const TA_LEFT = 0
    Const TA_CENTER = 6
    Const TA_RIGHT = 2
    Select Case horzJust
        Case ppcbJustifyLeft
            SetTextAlign hdc, TA_TOP Or TA_LEFT Or TA_NOUPDATECP
        Case ppcbJustifyHCenter
            SetTextAlign hdc, TA_TOP Or TA_CENTER Or TA_NOUPDATECP
        Case ppcbJustifyRight
            SetTextAlign hdc, TA_TOP Or TA_RIGHT Or TA_NOUPDATECP
    End Select
    ' convert angle from degrees to radians
    Dim angle As Double
    angle = (CDbl(adjustedOrientation) / 180#) * 3.14159265358979
    ' if test in mirrored, we should negate x coordinate
    Dim m As Double
    If mirrored Then m = -1# Else m = 1
    Dim xf As XFORM
    ' for rotation transformation, see SetWorldTransform Help
    xf.eM11 = CSng(m * Cos(angle))
    xf.eM21 = CSng(Sin(angle))
    xf.eM12 = CSng(m * Sin(angle))
    xf.eM22 = CSng(-Cos(angle))
    ' move text to proper position
    xf.eDx = CSng(X)
    xf.eDy = CSng(Y)

    height = aHeight

    ' "advanced" GDI mode required for mirrored texts
    Const GM_ADVANCED = 2
    SetGraphicsModeCustom GM_ADVANCED
    SetWorldTransformCustom xf
    TextOutCustom 0, offset, text
    ' cleanup "advanced" mode and transformation
    Const MWT_IDENTITY = 1
    ModifyWorldTransformCustom xf, MWT_IDENTITY ' "xf" should be ignored
    Const GM_COMPATIBLE = 1
    SetGraphicsModeCustom GM_COMPATIBLE
    ' restore old font
    SelectObject hdc, hOldFont
End Sub

Private Function IsShapeFilled(ByVal sType As PPcbShapeType) As Boolean
    IsShapeFilled = (winNT And sType = ppcbShapeTypeFilled)
End Function

' The following methods are closely related to Win32 GDI functions.
' You could use Win32 Help for descriptions.

' The following methods are overloaded to support operation under Win95-98.

Private Function SetArcDirectionCustom(ByVal nArcDirection As Long) As Long
    If winNT Then
        SetArcDirectionCustom = SetArcDirection(hdc, nArcDirection)
    Else
        SetArcDirectionCustom = ArcDirection
        ArcDirection = nArcDirection
    End If
End Function

Private Function ArcToCustom(rct As RectExt, start As CPoint, dest As CPoint) As Boolean
    If winNT Then
        ArcToCustom = ArcTo(hdc, rct.left, rct.top, rct.right, rct.bottom, _
                            start.X, start.Y, dest.X, dest.Y)
    Else
        Const AD_COUNTERCLOCKWISE = 1
        If ArcDirection = AD_COUNTERCLOCKWISE Then
            ArcToCustom = Arc(hdc, rct.left, rct.top, rct.right, rct.bottom, _
                              start.X, start.Y, dest.X, dest.Y)
        Else
            Const AD_CLOCKWISE = 2
            If ArcDirection = AD_CLOCKWISE Then
                ArcToCustom = Arc(hdc, rct.left, rct.top, rct.right, rct.bottom, _
                                  dest.X, dest.Y, start.X, start.Y)
            End If
        End If
        MoveToEx hdc, dest.X, dest.Y, 0
    End If
End Function

Private Function TextOutCustom(ByVal X As Long, ByVal Y As Long, str As String) As Boolean
    Const GM_COMPATIBLE = 1
    If winNT Or graphicsMode = GM_COMPATIBLE Then
        TextOutCustom = TextOut(hdc, X, Y, str, Len(str))
    Else
        
        ' record text in path
        BeginPath hdc
        TextOut hdc, X, 0, str, Len(str)
        EndPath hdc

        ' extract path points
        ReDim points(1) As Point
        ReDim flags(1) As Byte
        Dim n As Long
        n = GetPath(hdc, points(1), flags(1), 0)
        ReDim points(n)
        ReDim flags(n)
        Dim i As Long
        For i = 1 To n
            points(i).X = 0
            points(i).Y = 0
            flags(i) = 0
        Next i
        GetPath hdc, points(1), flags(1), n

        ' calculate Y coordinates range
        Dim y_min As Long, y_max As Long
        Const INT_MAX = 2147483647
        Const INT_MIN = -INT_MAX - 1 '–2147483648
        y_min = INT_MAX
        y_max = INT_MIN
        For i = 1 To n
            Dim yTmp As Long
            yTmp = points(i).Y
            If yTmp < y_min Then y_min = yTmp
            If yTmp > y_max Then y_max = yTmp
        Next i
        ' calcucate additional Y coordinate transformation parameters
        Dim ky As Double
        ky = 0.65 * CDbl(height) / CDbl(y_max - y_min)

        ' trasform path points
        For i = 1 To n
            Dim t As Point
            Dim py As Long
            py = Y + CLng(ky * (points(i).Y - y_min) + 0.25 * height)
            t.X = CLng(points(i).X * xfrm.eM11 + py * xfrm.eM21 + xfrm.eDx)
            t.Y = CLng(points(i).X * xfrm.eM12 + py * xfrm.eM22 + xfrm.eDy)
            points(i) = t
        Next i

        ' record transformed poly in path
        BeginPath hdc
        PolyDrawCustom points, flags, n
        EndPath hdc

        ' draw filled patt
        Dim hBrush As Long
        Const BLACK_BRUSH = 4
        hBrush = GetStockObject(BLACK_BRUSH)
        Dim hOldBrush As Long
        hOldBrush = SelectObject(hdc, hBrush)
        TextOutCustom = FillPath(hdc)
        SelectObject hdc, hOldBrush
    End If
End Function

' The following methods are trivial wrappers of Win32 GDI functions under NT-Win2000.
' They are re-implemented under Win95-98.

Private Function PolyDrawCustom(points() As Point, types() As Byte, ByVal count As Long) As Boolean
    If winNT Then
        PolyDrawCustom = PolyDraw(hdc, points(1), types(1), count)
    Else
        Dim i As Long
        For i = 1 To count
            Dim t As Byte
            Const PT_CLOSEFIGURE = &H1
            t = types(i) And Not PT_CLOSEFIGURE
            Const PT_MOVETO = &H6
            If t = PT_MOVETO Then
                MoveToEx hdc, points(i).X, points(i).Y, 0
            Else
                Const PT_LINETO = &H2
                If t = PT_LINETO Then
                    LineTo hdc, points(i).X, points(i).Y
                    If types(i) And PT_CLOSEFIGURE Then
                        CloseFigure hdc
                    End If
                Else
                    Const PT_BEZIERTO = &H4
                    If t = PT_BEZIERTO Then
                        PolyBezierTo hdc, points(i), 3
                        i = i + 2
                        If types(i) And PT_CLOSEFIGURE Then
                            CloseFigure hdc
                        End If
                    End If
                End If
            End If
        Next i
        PolyDrawCustom = True
    End If
End Function

Private Sub SetGraphicsModeCustom(ByVal mode As Long)
    If winNT Then
        SetGraphicsMode hdc, mode
    Else
        graphicsMode = mode
    End If
End Sub

Private Sub SetWorldTransformCustom(xf As XFORM)
    If winNT Then
        SetWorldTransform hdc, xf
    Else
        Const GM_ADVANCED = 2
        If graphicsMode = GM_ADVANCED Then xfrm = xf
    End If
End Sub

Private Sub ModifyWorldTransformCustom(xf As XFORM, ByVal operation As Long)
    If winNT Then
        ModifyWorldTransform hdc, xf, operation
    Else
        Const GM_ADVANCED = 2
        Const MWT_IDENTITY = 1
        If graphicsMode = GM_ADVANCED And operation = MWT_IDENTITY Then
            SetIdentityTrasnform
        End If
    End If
End Sub

Private Sub SetIdentityTrasnform()
    xfrm.eM11 = 1#
    xfrm.eM22 = 1#
    xfrm.eM12 = 0#
    xfrm.eM21 = 0#
    xfrm.eDx = 0#
    xfrm.eDy = 0#
End Sub

