//////////////////////////////////////////////////////////////////////////////////////
//
// CPicture class implementation
//
//////////////////////////////////////////////////////////////////////////////////////
// This is a part of the Innoveda-PowerPCB OLE Automation server PPCBAutoGeom sample.
// Copyright (C) 2003 Mentor Graphics Corp.
// All rights reserved.
//
// This source code is only intended as a supplement to the Innoveda-PowerPCB OLE
// Automation Server API Help file.
//////////////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "PPCBAutoGeom.h"

#include "Picture.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


BOOL CPicture::winNT = FALSE; // to be set appropriately by CPicture constructor

CBitmap CPicture::bmp; // the bitmap we'll store the picture in (one and only instance)

CPicture CPicture::object; // picture instance

// Build a picture bitmap based on Automation-retrieved data.
void CPicture::Refresh()
{
	CleanUp(); // discard what we may have had.

	// Let's assume there is a window and there's no need to be careful accessing it through a pointer.
	CDC *dc = ((CMainFrame*)AfxGetMainWnd())->GetViewWnd().GetDC(); // to be "Release"d!

	// Let's create a full-screen bitmap to allow to increase window's size without picture rebuild, if desired.
	const width = GetSystemMetrics(SM_CXFULLSCREEN);
	const height = GetSystemMetrics(SM_CYFULLSCREEN);

	// One may wonder, why we re-create a full-screen bitmap instead of reusing it.
	// Well, screen resolution might have been increased...

	if ( bmp.CreateCompatibleBitmap( dc, width, height ) ) {
		CAutoDraw autoDraw(winNT);
		autoDraw.Refresh(dc, bmp, width, height);
	}
	else { // if there's a trouble, report it
		AfxMessageBox( IDS_PICTUREBUILDFAILED, MB_ICONEXCLAMATION );
	}

	if ( dc ) {
		((CMainFrame*)AfxGetMainWnd())->GetViewWnd().ReleaseDC(dc); // prevent a GDI resource leak
	}
}

// draw the picture in a window
void CPicture::Draw( CDC &dc )
{
	// create a DC for stored bitmap
	CDC memDc;
	if ( (HBITMAP)bmp && memDc.CreateCompatibleDC( &dc ) ) {
		CBitmap *oldBmp = memDc.SelectObject( &bmp );

		CRect clientRect;
		dc.GetWindow()->GetClientRect(&clientRect); // if there's no window, there's nothing to trust in

		// let's (proudly?) show our bitmap
		dc.BitBlt(0,0,clientRect.right,clientRect.bottom,&memDc,0,0,SRCCOPY);

		memDc.SelectObject( oldBmp );
	}
}

const long CAutoDraw::unit = ppcbUnitMils;			
const long CAutoDraw::origin = ppcbOriginTypeDesign;
bool no_powerpcb = false;

// re-read all data from PowerPCB and redraw them in the bitmap
void CAutoDraw::Refresh(CDC * dc, CBitmap& bmp, int width, int height)
{
	no_powerpcb = false;
	if ( CreateCompatibleDC( dc ) ) {

		const index = SaveDC();

		CBrush brush;
		brush.CreateStockObject( HOLLOW_BRUSH );

		if (index) {
			SelectObject( &bmp );
			SelectObject( &brush );
			SetBkMode(TRANSPARENT);

			// paint the background
			FillSolidRect( 0, 0, width, height, GetBkColor() ); 

			TRY {
				Do();
			}
			CATCH ( CException, error ) {
				if (no_powerpcb) {
					AfxMessageBox("Open a design in PowerPCB first.");
				} else {
					error->ReportError();
				}
			}
			END_CATCH

			RestoreDC(index);
		}
	}
}

// connnect to PowerPCb and initiate document object
void CAutoDraw::ConnectToPowerPCB()
{
	COleException error;
	IPowerPCBApp app;

	// connect to PowerPCB Application (main) object
	{
		// map prog id to CLSID
		CLSID clsid;
		SCODE sc = AfxGetClassIDFromString(_T("powerpcb.application"), &clsid); // CLSIDFromProgID
		if (FAILED(sc)) {
			AfxThrowOleException(sc);
		}
		// get object
		LPUNKNOWN lpUnknown = NULL;
		sc = GetActiveObject( clsid, NULL, &lpUnknown );
		if ( FAILED(sc) ) {
			no_powerpcb = true;
			AfxThrowOleException(sc);
		}
		// query for IDispatch interface
		LPDISPATCH lpDispatch = NULL;
		sc = lpUnknown->QueryInterface(IID_IDispatch, (void **)&lpDispatch);
		lpUnknown->Release();
		if (FAILED(sc)) {
			AfxThrowOleException(sc);
		}
		app.AttachDispatch(lpDispatch);
	}

	doc.AttachDispatch(app.GetActiveDocument());
}


// Having separated routine preparation actions into wrappers, let's finally play the Automation game.
void CAutoDraw::Do()
{
	ConnectToPowerPCB();

	IPowerPCBView view(doc.GetActiveView()); // the picture we want is inside a view
	{
		SetMapMode(MM_ISOTROPIC);

		// we'll need it to refine picture's placement
		CPoint topLeft((int)view.GetTopLeftX(unit), (int)view.GetTopLeftY(unit)); 
		// set picture display parameters (units conversion, axes orientation, position)
		CPoint bottomRight( (int)view.GetBottomRightX(unit), (int)view.GetBottomRightY(unit) );

		SetWindowExt( (bottomRight - topLeft) + CSize(1) );
		SetWindowOrg( topLeft ); // let's respect view's (PowerPCB's window scrolling) current position

		CRect clientRect;
		AfxGetMainWnd()->GetClientRect( &clientRect );
		SetViewportExt( clientRect.Size() );
	}

	// Drawings 
	DrawObjects( ppcbObjectTypeDrawing );
	// RouteSegments
	DrawObjects( ppcbObjectTypeRouteSegment );
	// Pins
	DrawObjects( ppcbObjectTypePin );
	// Vias
	DrawObjects( ppcbObjectTypeVia );
	// Texts
	DrawObjects( ppcbObjectTypeText );
	// Labels
	DrawObjects( ppcbObjectTypeLabel );
}

// select a method to draw a typed object
void CAutoDraw::DrawObjects( long pcbObjectType, IPowerPCBObjs &objects )
{
	VARIANT var;
	var.vt = VT_I4;
	
	// for all objects in the collection
	for ( long i = 1, count = objects.GetCount(); i <= count; ++i ) {
		var.lVal = i;
		if ( objects.GetItemType(var) == pcbObjectType ) {
			switch ( pcbObjectType ) {
				case ppcbObjectTypeCircle:
					DrawCircle( objects.GetItem(var) );
					break;
				case ppcbObjectTypeDrawing:
					DrawDrawing( objects.GetItem(var) );
					break;
				case ppcbObjectTypeLabel:
					DrawLabel( objects.GetItem(var) );
					break;
				case ppcbObjectTypePolyline:
					DrawPolyline( objects.GetItem(var) );
					break;
				case ppcbObjectTypeRouteSegment:
					DrawRouteSegment( objects.GetItem(var) );
					break;
				case ppcbObjectTypeText:
					DrawText( objects.GetItem(var) );
					break;
				case ppcbObjectTypePin:
					DrawPin( objects.GetItem(var) );
					break;
				case ppcbObjectTypeVia:
					DrawVia( objects.GetItem(var) );
					break;
				default:
					ASSERT(0);
					break;
			} // switch
		} // if
	} // for
}

// extract data from "Circle" PowerPCB Automation object and draw
void CAutoDraw::DrawCircle(LPDISPATCH aCircle )
{
	IPowerPCBCircle circle( aCircle );

	CHatchBrush brush;
	if (IsShapeFilled((PPcbShapeType)circle.GetShapeType())) {
		SelectBrush( brush );
	}

	CSolidPen pen;
	SelectPen( pen, circle.GetLineWidth(unit) );

	Ellipse( CRectExt(circle.GetCenterX(unit,origin), circle.GetCenterY(unit,origin), circle.GetRadius(unit)) );
}

// extract data from "Drawing" PowerPCB Automation object and draw
void CAutoDraw::DrawDrawing(LPDISPATCH aDrawing )
{
	IPowerPCBDrawing drw(aDrawing);

	IPowerPCBObjs geometry(drw.GetGeometry());
	DrawObjects( ppcbObjectTypeCircle, geometry );
	DrawObjects( ppcbObjectTypePolyline, geometry );
}

// extract data from "Pin" PowerPCB Automation object and draw
void CAutoDraw::DrawPin(LPDISPATCH aPin )
{
	IPowerPCBPin pin( aPin );

	Ellipse( CRectExt(pin.GetPositionX(unit), pin.GetPositionY(unit), pin.GetDrillSize(unit) / 2.) );
}

// extract data from "Via" PowerPCB Automation object and draw
void CAutoDraw::DrawVia(LPDISPATCH aVia )
{
	IPowerPCBVia via( aVia );

	Ellipse( CRectExt(via.GetPositionX(unit), via.GetPositionY(unit), via.GetDrillSize(unit) / 2.) );
}

// Draw text (for "Test" or "Label") object in bitmap 
// using parameters common for both objects
void CAutoDraw::DoDrawText(const CString &text, int x, int y, int height,
					PPcbHorizontalJustification horzJust, PPcbVerticalJustification vertJust,
					int aOrientation, BOOL mirrored, LPCTSTR fontName, PPcbRightReadingStatus rrs )
{
	// skip empty tests
	if ( text.IsEmpty() ) {
		return;
	}
	// else

	// ensure the orientation is in 0 - 359 range
	int orientation = aOrientation % 360;
	if ( orientation < 0 ) {
		orientation += 360;
	}

	// orientation ajdusted to take "Right Reading" into account
	int adjustedOrientation = orientation; 

	if ( rrs == ppcbRightReadingAngled ) {
		if ( orientation > 90 && orientation < 270 ) {
			orientation += 180;
		}
	}
	else if ( rrs == ppcbRightReadingOrthogonal ) {
		const slack = orientation % 180;
		if ( slack > 45 && slack < 135 ) {
			orientation = 90;
		}
		else {
			orientation = 0;
		}
		adjustedOrientation = orientation;
	}

	// create font in laft-to-right (zero) orientation
	CFont font;
	VERIFY( font.CreateFont( -height, 0, 0, 0,
								FW_NORMAL, FALSE, FALSE, FALSE, ANSI_CHARSET,
								OUT_DEFAULT_PRECIS, CLIP_LH_ANGLES | CLIP_DEFAULT_PRECIS,
								DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, fontName ));

	CFont *oldFont = SelectObject( &font );

	// PowerPCB vertical justification does not match Win32 GDI types
	// y (vertical) position is adjusted to place desired text point at y = 0 line
	int offset = 0;
	{
		TEXTMETRIC tm;
		GetTextMetrics(&tm);
		switch ( vertJust ) {
			case ppcbJustifyTop:
				offset = - tm.tmInternalLeading;
				break;
			case ppcbJustifyVCenter:
				offset =  - (tm.tmAscent + tm.tmInternalLeading) / 2;
				break;
			case ppcbJustifyBottom:
				offset =  - tm.tmAscent;
				break;
		}
	}
	// horizontal justification is the same, so Win32 setting 
	switch ( horzJust ) {
		case ppcbJustifyLeft:
			SetTextAlign(TA_TOP | TA_LEFT | TA_NOUPDATECP);
			break;
		case ppcbJustifyHCenter:
			SetTextAlign(TA_TOP | TA_CENTER | TA_NOUPDATECP);
			break;
		case ppcbJustifyRight:
			SetTextAlign(TA_TOP | TA_RIGHT | TA_NOUPDATECP);
			break;
	}
	// convert angle from degrees to radian
	const double angle = (double(adjustedOrientation) / 180.) * 3.1415926535897932384626433832795;
	// if test in mirrored, we should negate x coordinate
	const double m = (mirrored) ? -1. : 1; 
	XFORM xf;
	// for rotation transformation, see SetWorldTransform Help
	xf.eM11 = float(m * cos(angle));
	xf.eM21 = float(sin(angle));
	xf.eM12 = float(m * sin(angle));
	xf.eM22 = float(- cos(angle));
	// move text in proper position
	xf.eDx = float(x);
	xf.eDy = float(y);

	m_height = height;

	// "advanced" GDI mode required for mirrored texts
	SetGraphicsMode(GM_ADVANCED);
	SetWorldTransform(&xf);
	VERIFY(TextOut( 0, offset, text ));
	// cleanup "advanced" mode and transformation
	ModifyWorldTransform(NULL, MWT_IDENTITY);
	SetGraphicsMode(GM_COMPATIBLE);
	// restore old font
	SelectObject( oldFont );
}

// extract data from "Label" PowerPCB Automation object and draw
void CAutoDraw::DrawLabel(LPDISPATCH aLabel )
{
	IPowerPCBLabel label( aLabel );

	DoDrawText(label.GetText(), (int)label.GetPositionX(unit,origin), (int)label.GetPositionY(unit,origin),
				(int)label.GetHeight(unit), (PPcbHorizontalJustification)label.GetHorzJustification(),
				(PPcbVerticalJustification)label.GetVertJustification(), (int)label.GetOrientation(origin),
				label.GetMirror(origin), _T("Arial"), (PPcbRightReadingStatus)label.GetRightReading() );
}

// extract data from "Text" PowerPCB Automation object and draw
void CAutoDraw::DrawText(LPDISPATCH aText )
{
	IPowerPCBText text( aText );

	DoDrawText(text.GetText(), (int)text.GetPositionX(unit,origin), (int)text.GetPositionY(unit,origin),
				(int)text.GetHeight(unit), (PPcbHorizontalJustification)text.GetHorzJustification(),
				(PPcbVerticalJustification)text.GetVertJustification(), (int)text.GetOrientation(origin),
				text.GetMirror(origin), _T("Times New Roman") );
}

// extract data from "Polyline" PowerPCB Automation object and draw
void CAutoDraw::DrawPolyline(LPDISPATCH aPolyline )
{
	IPowerPCBPolyline polyline( aPolyline );

	CSolidPen pen;
	SelectPen(pen, polyline.GetLineWidth(unit));

	if ( IsShapeFilled((PPcbShapeType)polyline.GetShapeType()) ) {
		// fill the shape
		VERIFY(BeginPath());
		DrawSegments(polyline);
		VERIFY(EndPath());
		CHatchBrush brush;
		SelectBrush(brush);
		VERIFY(StrokeAndFillPath());
	} else {
		DrawSegments(polyline);
	}
}

// Polyline segments drawing helper: draws all segments
void CAutoDraw::DrawSegments(IPowerPCBPolyline& polyline)
{
	CPointsArray array = polyline.GetPoints(unit,origin); // ~COleSafeArray will clean it up

	MoveTo(array.GetPoint(1));
	for ( long corner = 2; corner <= array.GetSize(); corner++ ) {
		CPoint dest = array.GetPoint(corner);
		const double angle = array.GetAngle(corner - 1);
		if ( angle != 0 ) { // it's not over yet
			CPoint start = array.GetPoint(corner - 1);
			CRectExt rect(polyline.GetCenterX(corner - 1, unit, origin), 
											polyline.GetCenterY(corner - 1, unit, origin),
											polyline.GetRadius(corner - 1, unit) );
			SetArcDirection((angle > 0.) ? AD_COUNTERCLOCKWISE : AD_CLOCKWISE);
			ArcTo(rect, start, dest);
		} else {
			LineTo(dest);
		}
	}
}

// check if both r =- r and d + r fit into "long" data type
static inline bool IsConvertibleToLong(double d, double r)
{
	return (double(LONG_MAX) >= d + r) && (d - r >= double(LONG_MIN));
}

// Calculate arc rectangle (actually, square) by three points
// returns:
//		true  if rectangle can be calculated
//		false if rectangle cannot be calculated 
//          (all points on a line, or the result does not fit into RECT)
static bool ArcRect(
	RECT& rect,		// calculated arc rectangle
	bool& ccw,		// calcucated arc direction
	POINT start,	// start point on the circle
	POINT interm, 	// intermediate point on the circle
	POINT end 		// last point on the circle
) {
	// a = {ax, ay} is vector from middle point to start point
	double ax = start.x - interm.x;
	double ay = start.y - interm.y;
	// c = {cx, cy} is vector from middle point to end point
	double cx = end.x - interm.x;
	double cy = end.y - interm.y;
	double aa = ax * ax + ay * ay; // length of vector "a" squared
	double cc = cx * cx + cy * cy; // length of vector "c" squared
	// ac is doubled cross product "a x c" or four times triangle area (signed)
	double ac = 2. * (ax * cy - ay * cx);

	// if ac is zero, the points are on a line (not on arc)
	if (ac != 0.) {
		// {rx, ry} is vertor from middle point to arc center
		double rx = (cy * aa - ay * cc) / ac;
		double ry = (ax * cc - cx * aa) / ac;
		// length of vertor "r" is the arc radius
		double rad = sqrt(rx * rx + ry * ry); 

		rx += interm.x;
		ry += interm.y;
		// now r is the arc center position

		// check if arc rectangle fit into "long" RECT fields
		if (! IsConvertibleToLong(rx, rad)) return false;
		if (! IsConvertibleToLong(ry, rad)) return false;

		// construct the arc rectangle
		rect = CRectExt(rx, ry, rad);
		// sign of ac defines arc direction
		ccw = ac < 0.;
		return true;
	} else {
		return false;
	}
}

// extract data from "RouteSegment" PowerPCB Automation object and draw
void CAutoDraw::DrawRouteSegment(LPDISPATCH aRouteSegment )
{
	IPowerPCBRouteSegment routeSegment( aRouteSegment );

	CPointsArray array = routeSegment.GetPoints(unit); // ~COleSafeArray will clean it up

	enum {startPnt = 1, endPnt = 2, midPnt = 3, segmentSize = 2, arcSize = 3};

	CSolidPen pen;
	SelectPen( pen, routeSegment.GetWidth(unit) );
	
	const int size = array.GetSize();
	if (size >= segmentSize) {
		CPoint start( array.GetPoint(startPnt) );
		CPoint end( array.GetPoint(endPnt) );
		MoveTo( start ); // set starting point

		if (size == segmentSize) {
			// 2 points - draw a line
			LineTo( end );
		} else if (size == arcSize) {
			// 3 points - draw an arc
			CPoint interm( array.GetPoint(midPnt) );
			CRect arcRect;
			bool ccw;
			// calculate arc rectangle and direction by three points
			if ( ArcRect(arcRect, ccw, start, interm, end) ) {
				SetArcDirection(ccw ? AD_COUNTERCLOCKWISE : AD_CLOCKWISE);
				ArcTo(arcRect, start, end);
			}
			else {
				AfxMessageBox(IDS_BADROUTESEGMENTARC, MB_ICONEXCLAMATION);
			}
		} else {
			AfxMessageBox(IDS_BADVARIANT, MB_ICONSTOP);
		}
	}
}
