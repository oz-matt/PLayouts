VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Sample3 (Visual Basic 6.0 Version)"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox LstDocObjects 
      Height          =   2400
      Left            =   5040
      TabIndex        =   12
      Top             =   360
      Width           =   2172
   End
   Begin VB.TextBox EdtDocSelectedComps 
      Enabled         =   0   'False
      Height          =   288
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   4692
   End
   Begin VB.TextBox EdtDocComps 
      Enabled         =   0   'False
      Height          =   288
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   4692
   End
   Begin VB.CommandButton BtnOk 
      Caption         =   "Close"
      Height          =   372
      Left            =   3720
      TabIndex        =   5
      Top             =   2280
      Width           =   1092
   End
   Begin VB.CommandButton BtnRefresh 
      Caption         =   "Refresh"
      Height          =   372
      Left            =   2520
      TabIndex        =   4
      Top             =   2280
      Width           =   1092
   End
   Begin VB.CommandButton BtnDisconnect 
      Caption         =   "Disconnect"
      Height          =   372
      Left            =   1320
      TabIndex        =   3
      Top             =   2280
      Width           =   1092
   End
   Begin VB.CommandButton BtnConnect 
      Caption         =   "Connect"
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1092
   End
   Begin VB.TextBox EdtDocName 
      Enabled         =   0   'False
      Height          =   288
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4692
   End
   Begin VB.Label Label4 
      Caption         =   "PowerPCB Doc Components"
      Height          =   252
      Left            =   5040
      TabIndex        =   11
      Top             =   120
      Width           =   2172
   End
   Begin VB.Label Label3 
      Caption         =   "PowerPCB Selected Component Count"
      Height          =   252
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   4692
   End
   Begin VB.Label Label2 
      Caption         =   "PowerPCB Component Count"
      Height          =   252
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   4692
   End
   Begin VB.Label Status 
      Caption         =   "Status Line"
      Height          =   252
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   4692
   End
   Begin VB.Label Label1 
      Caption         =   "PowerPCB Document Name"
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4692
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SAMPLE3.BAS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This is a part of the PADS-PowerPCB OLE Automation server SAMPLE3 sample.
' Copyright (C) 2003 Mentor Graphics Corp.
' All rights reserved.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This source code is only intended as a supplement to the PADS-PowerPCB
' Automation Server API Help file.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim WithEvents powerPCBApp As PowerPCB.Application
Attribute powerPCBApp.VB_VarHelpID = -1
Dim WithEvents powerPCBDoc As PowerPCB.Document
Attribute powerPCBDoc.VB_VarHelpID = -1

Private Sub RefreshGeneralInfos()
    ' If we are not connected, exit procedure
    If powerPCBDoc Is Nothing Then
        EdtDocName.Text = "???????"
        EdtDocComps.Text = "???????"
        EdtDocSelectedComps.Text = "???????"
        Exit Sub
    End If

    On Error GoTo OnErrorServerAccess
    
    ' Output the PowerPCB server Document name to our editbox
    EdtDocName.Text = powerPCBApp.ActiveDocument.Name
    
    ' Output the PowerPCB component objects
    mcount = powerPCBDoc.GetObjects(1).Count
    EdtDocComps.Text = mcount & " component object(s)"
    
    ' Output the PowerPCB selected component objects
    mcount = powerPCBDoc.GetObjects(1, "", True).Count
    EdtDocSelectedComps.Text = mcount & " selected component object(s)"

Exit Sub        ' Exit to avoid handler.

OnErrorServerAccess:
    Msg = "Are you sure you are connected??"
    MsgBox Msg, , "Sample2"

End Sub

Public Sub RefreshObjectsInfos()
    ' Clear the listbox
    LstDocObjects.Clear
    
    ' If we are not connected, exit procedure
    If powerPCBApp Is Nothing Then Exit Sub

    On Error GoTo OnErrorServerAccess
    
    ' Get all PowerPCB objects of type COMPONENT from PowerPCB database
    ' and output each component in the listbox
    countComp = 0
    For Each nextComp In powerPCBDoc.GetObjects(1)
        LstDocObjects.AddItem nextComp.Name
        countComp = countComp + 1
        Status.Caption = "Loading " & countComp & " component(s) from PowerPCB..."
        Status.Refresh
    Next nextComp
    Status.Caption = ""
Exit Sub        ' Exit to avoid handler.

OnErrorServerAccess:
    Msg = "Are you sure you are connected??"
    MsgBox Msg, , "Sample3"
    
End Sub

Private Sub BtnConnect_Click()
    ' Enable error-handling routine.
    On Error GoTo OnErrorGetObject
    
    ' Connect to a running instance of PowerPCB server
    Set powerPCBApp = GetObject(, "PowerPCB.Application")
    Set powerPCBDoc = powerPCBApp.ActiveDocument
    
    ' Enable and disable correct buttons
    BtnConnect.Enabled = False
    BtnDisconnect.Enabled = True
    
    ' Set status line text
    Status.Caption = "Connected to PowerPCB."
    
Exit Sub        ' Exit to avoid handler.

OnErrorGetObject:
    Msg = "Cannot connect to a running PowerPCB Server!"
    MsgBox Msg, , "Sample2"
    
End Sub

Private Sub BtnDisconnect_Click()
    ' Disconnect from PowerPCB server
    Set powerPCBDoc = Nothing
    Set powerPCBApp = Nothing
    ' Enable and disable correct buttons
    BtnConnect.Enabled = True
    BtnDisconnect.Enabled = False
         ' Set status line text
    Status.Caption = "Not connected to PowerPCB."
End Sub
Private Sub BtnOk_Click()
    End
End Sub

Private Sub BtnRefresh_Click()
    RefreshGeneralInfos
    RefreshObjectsInfos
End Sub

Private Sub Form_Load()
    BtnConnect.Enabled = True
    BtnDisconnect.Enabled = False
    Status.Caption = "Not connected to PowerPCB."
End Sub


Private Sub powerPCBApp_OpenDocument(ByVal Doc As PowerPCB.Document)
    RefreshGeneralInfos
    RefreshObjectsInfos
End Sub

Private Sub powerPCBApp_Quit()
    BtnDisconnect_Click
    BtnRefresh_Click
End Sub

Private Sub powerPCBDoc_SelectionChange()
    RefreshGeneralInfos
End Sub
