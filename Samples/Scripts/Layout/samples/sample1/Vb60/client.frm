VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Sample1 (Visual Basic 6.0 Version)"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BtnOk 
      Caption         =   "Close"
      Height          =   372
      Left            =   3720
      TabIndex        =   5
      Top             =   840
      Width           =   1092
   End
   Begin VB.CommandButton BtnRefresh 
      Caption         =   "Refresh"
      Height          =   372
      Left            =   2520
      TabIndex        =   4
      Top             =   840
      Width           =   1092
   End
   Begin VB.CommandButton BtnDisconnect 
      Caption         =   "Disconnect"
      Height          =   372
      Left            =   1320
      TabIndex        =   3
      Top             =   840
      Width           =   1092
   End
   Begin VB.CommandButton BtnConnect 
      Caption         =   "Connect"
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   840
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
   Begin VB.Label Status 
      Caption         =   "Status Line"
      Height          =   252
      Left            =   120
      TabIndex        =   6
      Top             =   1320
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
' SAMPLE1.BAS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This is a part of the PADS-PowerPCB OLE Automation server SAMPLE1 sample.
' Copyright (C) 2003 Mentor Graphics Corp.
' All rights reserved.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This source code is only intended as a supplement to the PADS-PowerPCB
' Automation Server API Help file.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim powerPCBApp As Object
Attribute powerPCBApp.VB_VarHelpID = -1

Private Sub RefreshDialog()
    ' If we are not connected, exit procedure
    If powerPCBApp Is Nothing Then
        EdtDocName.Text = "???????"
        Exit Sub
    End If

    On Error GoTo OnErrorServerAccess
    
    ' Output the PowerPCB server Document name to our editbox
    EdtDocName.Text = powerPCBApp.ActiveDocument.Name
    
Exit Sub        ' Exit to avoid handler.

OnErrorServerAccess:
    Msg = "Are you sure you are connected??"
    MsgBox Msg, , "Sample1"

End Sub

Private Sub BtnConnect_Click()
    ' Enable error-handling routine.
    On Error GoTo OnErrorGetObject
    
    ' Connect to a running instance of PowerPCB server
    Set powerPCBApp = GetObject(, "PowerPCB.Application")
    
    ' Enable and disable correct buttons
    BtnConnect.Enabled = False
    BtnDisconnect.Enabled = True
    
    ' Set status line text
    Status.Caption = "Connected to PowerPCB."
    
Exit Sub        ' Exit to avoid handler.

OnErrorGetObject:
    Msg = "Cannot connect to a running PowerPCB Server!"
    MsgBox Msg, , "Sample1"
    
End Sub

Private Sub BtnDisconnect_Click()
    ' Disconnect from PowerPCB server
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
    RefreshDialog
End Sub

Private Sub Form_Load()
    BtnConnect.Enabled = True
    BtnDisconnect.Enabled = False
    Status.Caption = "Not connected to PowerPCB."
End Sub



