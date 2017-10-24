//////////////////////////////////////////////////////////////////////////////////////
//
// CPicture : drawing representation class
//
//////////////////////////////////////////////////////////////////////////////////////
// This is a part of the Innoveda-PowerPCB OLE Automation server PPCBAutoGeom sample.
// Copyright (C) 2003 Mentor Graphics Corp.
// All rights reserved.
//
// This source code is only intended as a supplement to the Innoveda-PowerPCB OLE
// Automation Server API Help file.
//////////////////////////////////////////////////////////////////////////////////////

#ifndef __PICTURE_H__
#define __PICTURE_H__

#include "PPCBAutoGeom.h" // application class

#include "import.h" // Automation interfaces

// the class stores screen image as bitmap
class CPicture
{
public:

	static CPicture object; // the one and only instance

	CPicture()
	{
		ASSERT(this == &object);
		CleanUp();
		OSVERSIONINFO verInfo;
		verInfo.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);
		winNT = (GetVersionEx(&verInfo))? (verInfo.dwPlatformId == VER_PLATFORM_WIN32_NT) : FALSE;
	};

	~CPicture() { CleanUp(); }

	void CleanUp()
	{
		if ( (HBITMAP)bmp ) {
			bmp.DeleteObject();
		}
	}

	void Refresh(); // build the bitmap to show

	void Draw( CDC &dc ); // show bitmap on screen

protected:

	static CBitmap bmp; // picture bitmap
	static BOOL winNT; // current Windows is NT/2000
};

// the class facilitate rectangle (square) definition
// based in center point and radius of inscribed circle
class CRectExt : public CRect {
public:
	CRectExt(double x, double y, double radius) :
		CRect( int(x - radius), int(y - radius), int(x + radius), int(y + radius))
	{}
};

// automate DC cleanup after GDI object selection
template < class GDI_OBJ >
class CAutoObject : public GDI_OBJ
{
	friend class CAutoDraw;
	CAutoObject() { 
		pDC = NULL;
		pOldObject = NULL;
	}
	~CAutoObject() {
		if (pDC != NULL && pOldObject != NULL) {
			pDC->SelectObject(pOldObject);
		}
	}
	GDI_OBJ * pOldObject;
	CDC * pDC;
};

typedef CAutoObject<CPen> CSolidPen;
typedef CAutoObject<CBrush> CHatchBrush;

// facilitate data access for Automation arrays of points
class CPointsArray : public COleSafeArray
{
public:
	CPointsArray(const VARIANT& var) : COleSafeArray(var) {}

	// column structure indexes
	enum {
		pointX = 1,
		pointY = 2,
		routeDim2 = 2,
		angle = 3,
		polylineDim2 = 3,
		numDims = 2
	};

	bool Check(long dim2) {
		long lBound1, lBound2, uBound2; // array dimensions bounds (first (lower,upper), second (lower,upper))
		GetLBound( 1, &lBound1 );
		GetLBound( 2, &lBound2 );
		GetUBound( 2, &uBound2 );

		if ( !(vt & (VT_ARRAY | VT_R8)) || GetDim() != numDims ||
					lBound1 != 1 || lBound2 != 1 || uBound2 != dim2) {
			AfxMessageBox(IDS_BADVARIANT, MB_ICONSTOP);
			return false;
		}
		return true;
	}
	
	long GetSize() {
		long uBound1;
		GetUBound( 1, &uBound1 );
		return uBound1;
	}

	CPoint GetPoint (long corner) {
		return CPoint((int)GetElement(corner, pointX), (int)GetElement(corner, pointY));
	}

	int GetAngle(long corner) {
		return (int)GetElement(corner, angle);
	}

protected:
	double GetElement(long i, long j) {
		double ret;
		long indexes[numDims] = {i,j};
		COleSafeArray::GetElement( indexes, &ret );
		return ret;
	}

};


class CAutoDraw : public CDC
{
public:

	CAutoDraw(BOOL aWinNT)	{
		winNT = aWinNT;
		m_iGraphicsMode = GM_COMPATIBLE;
		m_nArcDirection = AD_COUNTERCLOCKWISE;
		SetIdentityTransform();
	}

	// refresh the drawing using Automation data (throws COleException and COleDispatchException)
	void Refresh(CDC * dc, CBitmap& bmp, int width, int height);

protected:

	BOOL winNT;	

	void SelectPen(CSolidPen& pen, double width) {
		pen.CreatePen( PS_SOLID, int(width), GetTextColor() );
		pen.pDC = this;
		pen.pOldObject = SelectObject(&pen);
	}

	void SelectBrush(CHatchBrush& brush) {
		brush.CreateHatchBrush(HS_DIAGCROSS, GetTextColor());
		brush.pDC = this;
		brush.pOldObject = SelectObject(&brush);
	}


	// refresh the drawing using Automation data (throws COleException and COleDispatchException)
	void Do();

	void ConnectToPowerPCB();

	// extract 'pcbObjectType' items from "objects" and draw them
	void DrawObjects( long pcbObjectType )	{
		DrawObjects( pcbObjectType, IPowerPCBObjs( doc.GetObjects( pcbObjectType, _T(""), FALSE )) );
	}

	void DrawObjects( long pcbObjectType, IPowerPCBObjs &objects );

	// PowerPCB Automation objects drawing methods
	void DrawCircle( LPDISPATCH aCircle );
	void DrawDrawing(LPDISPATCH aDrawing );
	void DrawLabel( LPDISPATCH aLabel );
	void DrawSegments(IPowerPCBPolyline& polyline);
	void DrawPolyline( LPDISPATCH aPolyline );
	void DrawRouteSegment( LPDISPATCH aRouteSegment );
	void DrawText( LPDISPATCH aText );
	void DrawPin( LPDISPATCH aPin );
	void DrawVia( LPDISPATCH aVia );

	void DoDrawText(const CString &text, int x, int y, int height,
					PPcbHorizontalJustification horzJust, PPcbVerticalJustification vertJust,
					int aOrientation, BOOL mirrored, LPCTSTR fontName, PPcbRightReadingStatus rrs = ppcbRightReadingNone );

	bool IsShapeFilled(PPcbShapeType type) {return winNT && type == ppcbShapeTypeFilled; }

	IPowerPCBDoc doc; // we need a document to query it's contents
	static const long unit, origin; // PPCB Automation parameters values to use

// following methods are closely related to Win32 GDI functions
// you could use Win32 and MFC Help for descriptions

	// following methods are overloaded to support operation under Win95-98
	int SetArcDirection(int nArcDirection );
	BOOL ArcTo(const CRect& rect, const CPoint &start, const CPoint &dest);
	BOOL TextOut( int x, int y, const CString& str );
	// following methods are trivial wrappers of Win32 GDI functions under NT-Win2000
	// they are re-implemented under Win95-98
	BOOL PolyDraw( const POINT* lpPoints, const BYTE* lpTypes, int nCount );
	void SetGraphicsMode(int mode);
	void SetWorldTransform(XFORM * xf);
	void ModifyWorldTransform(XFORM * xf, int operation); 


	int m_height;
private:
// following data members are needed only to support operations under Win95-98
// they are not used under NT
	void SetIdentityTransform();

	int m_iGraphicsMode;
	int m_nArcDirection;
	XFORM m_xForm;
};

inline CPicture *AfxGetPicture() { return &CPicture::object; }

#endif // __PICTURE_H__
