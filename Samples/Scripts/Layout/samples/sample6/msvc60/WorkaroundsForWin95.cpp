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

//
// This file includes only workarounds for Win 95 and Win 98 shortcoming
//
// You probably may skip it
//
// See Win32 and MFC Help for methods descriptions

int CAutoDraw::SetArcDirection(
  int nArcDirection   // new arc direction
) {
	if (winNT) {
		return CDC::SetArcDirection(nArcDirection);
	} else {
		int ret = m_nArcDirection;
		m_nArcDirection = nArcDirection;
		return ret;
	}
}


BOOL CAutoDraw::ArcTo(const CRect& rect, const CPoint &start, const CPoint &dest)
{
	if (winNT) {
		return CDC::ArcTo( rect, start, dest );
	}
	else {
		BOOL rc = FALSE;
		if (m_nArcDirection == AD_COUNTERCLOCKWISE) {
			rc = Arc( rect, start, dest );
		} else if (m_nArcDirection == AD_CLOCKWISE) {
			rc = Arc( rect, dest, start );
		}
		MoveTo(dest);
		return rc;
	}
}

BOOL CAutoDraw::TextOut( int x, int y, const CString& str )
{
	if (winNT || m_iGraphicsMode == GM_COMPATIBLE) {
		return CDC::TextOut(x, y, str);
	} 
	else {
		
		// record text in path
		VERIFY(BeginPath());
		CDC::TextOut(x, 0, str);
		VERIFY(EndPath());

		// extract path points
		int n = GetPath(NULL, NULL, 0);
		POINT * points = new POINT[n];
		BYTE * flags = new BYTE[n];
		for (int i = 0; i < n; i++) {
			points[i].x = 0;
			points[i].y = 0;
			flags[i] = 0;
		}
		VERIFY(GetPath(points, flags, n) == n);

		// calculate Y coordinates range
		int y_min = INT_MAX, y_max = INT_MIN;
		for (i = 0; i < n; i++) {
			int y = points[i].y;
			if (y < y_min) y_min = y;
			if (y > y_max) y_max = y;
		}
		// calcucate additional Y coordinate transformation parameters
		double ky = 0.65 * double(m_height) / double(y_max - y_min);

		// trasform path points
		for (i = 0; i < n; i++) {
			POINT t;
			POINT &p = points[i]; 
			int py = y + int(ky * (p.y - y_min) + 0.25 * m_height);
			t.x = long(p.x * m_xForm.eM11 + py * m_xForm.eM21 + m_xForm.eDx);
			t.y = long(p.x * m_xForm.eM12 + py * m_xForm.eM22 + m_xForm.eDy);
			p = t;
		}

		// record transformed poly in path
		VERIFY(BeginPath());
		PolyDraw(points, flags, n);
		VERIFY(EndPath());
		delete [] points;
		delete [] flags;

		// draw filled patt
		CBrush brush;
		brush.CreateStockObject(BLACK_BRUSH);
		CBrush * oldBrush = SelectObject(&brush);
		BOOL rc = FillPath();
		SelectObject(oldBrush);
		return rc;
	}
}


BOOL CAutoDraw::PolyDraw( const POINT* lpPoints, const BYTE* lpTypes, int nCount )
{
	if (winNT) {
		return PolyDraw(lpPoints, lpTypes, nCount);
	} else {
		for (int i = 0; i < nCount; i++) {
			BYTE t = lpTypes[i] & ~ PT_CLOSEFIGURE;
			if (t == PT_MOVETO) {
				MoveTo(lpPoints[i]);
			} 
			else if (t == PT_LINETO) {
				LineTo(lpPoints[i]);
				if (lpTypes[i] & PT_CLOSEFIGURE) {
					CloseFigure();
				}
			} 
			else if (t == PT_BEZIERTO) {
				ASSERT(i + 2 < nCount);
				PolyBezierTo(lpPoints + i, 3);
				i += 2;
				if (lpTypes[i] & PT_CLOSEFIGURE) {
					CloseFigure();
				}
			}
			else {
				ASSERT(FALSE);
			}

		}
		return TRUE;
	}
}


void CAutoDraw::SetGraphicsMode(int mode)
{
	if (winNT) {
		VERIFY(::SetGraphicsMode(m_hDC, mode));
	} else {
		m_iGraphicsMode = mode;
	}
}


void CAutoDraw::SetWorldTransform(XFORM * xf) 
{
	if (winNT) {
		VERIFY(::SetWorldTransform(m_hDC, xf)); 
	} else {
		ASSERT(m_iGraphicsMode == GM_ADVANCED);
		m_xForm = *xf;
	}
}

void CAutoDraw::ModifyWorldTransform(XFORM * xf, int operation)
{
	if (winNT) {
		VERIFY(::ModifyWorldTransform(m_hDC, xf, operation)); 
	} else {
		ASSERT(m_iGraphicsMode == GM_ADVANCED);
		ASSERT(operation == MWT_IDENTITY);
		SetIdentityTransform();
	}
}

void CAutoDraw::SetIdentityTransform()
{
	m_xForm.eM11 = 1.f;
	m_xForm.eM22 = 1.f;
	m_xForm.eM12 = 0.f;
	m_xForm.eM21 = 0.f;
	m_xForm.eDx = 0.f;
	m_xForm.eDy = 0.f;
}
