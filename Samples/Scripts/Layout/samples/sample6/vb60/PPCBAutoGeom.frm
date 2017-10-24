VERSION 5.00
Begin VB.Form FormMain 
   BackColor       =   &H80000005&
   Caption         =   "PPCBAutoGeom"
   ClientHeight    =   3870
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4725
   Icon            =   "PPCBAutoGeom.frx":0000
   ScaleHeight     =   3870
   ScaleMode       =   0  'User
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.Label StatusBar 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ready"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3600
      Width           =   4695
   End
   Begin VB.Menu Menu_File 
      Caption         =   "&File"
      Begin VB.Menu Menu_File_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu Menu_View 
      Caption         =   "&View"
      Begin VB.Menu Menu_View_Refresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu Menu_Separator 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_View_StatusBar 
         Caption         =   "&Status Bar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu Menu_Help 
      Caption         =   "&Help"
      Begin VB.Menu Menu_Help_About 
         Caption         =   "&About PPCBAutoGeom..."
      End
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//////////////////////////////////////////////////////////////////////////////////////
'//
'// FormMain : user interface definition
'//
'//////////////////////////////////////////////////////////////////////////////////////
'// This is a part of the Innoveda-PowerPCB OLE Automation server PPCBAutoGeom sample.
'// Copyright (C) 2003 Mentor Graphics Corp.
'// All rights reserved.
'//
'// This source code is only intended as a supplement to the Innoveda-PowerPCB OLE
'// Automation Server API Help file.
'//////////////////////////////////////////////////////////////////////////////////////

Private Sub Form_DblClick()
    Menu_View_Refresh_Click
End Sub

Private Sub Form_Initialize()
    StatusBar.Caption = LoadResString(AFX_IDS_IDLEMESSAGE)
    StatusBar.Visible = Menu_View_StatusBar.Checked ' Ensure correspondance.
    currentCursor = 0
End Sub

Private Sub Form_Paint()
    GetPicture().Draw hdc
End Sub

Private Sub Form_Resize()
    ' Adjust status bar's size+position
    StatusBar.top = ScaleHeight - StatusBar.height
    StatusBar.width = width
End Sub

Private Sub Menu_File_Exit_Click()
    ' Say "Good bye!".
    Unload FormMain
End Sub

Private Sub Menu_Help_About_Click()
    ' Display the "About..." dialog as modal.
    FormAbout.Show vbModal
End Sub

Private Sub Menu_View_Refresh_Click()
    StatusBar.Caption = LoadResString(IDS_LOADDATAMESSAGE)
    Const IDC_WAIT = 32514
    Dim prevCursor As Long
    prevCursor = SetCursor(LoadCursor(0, IDC_WAIT))
    
    GetPicture.Refresh
    
    SetCursor prevCursor
    StatusBar.Caption = LoadResString(AFX_IDS_IDLEMESSAGE)
    
    Refresh
End Sub

Private Sub Menu_View_StatusBar_Click()
    ' Invert the menu item's checked + status bar's visible status
    Menu_View_StatusBar.Checked = Not Menu_View_StatusBar.Checked
    StatusBar.Visible = Not StatusBar.Visible
End Sub
