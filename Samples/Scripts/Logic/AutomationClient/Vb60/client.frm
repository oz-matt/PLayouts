VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Client (Visual Basic 5.0 Version)"
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
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
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
      Caption         =   "PADS Logic Doc Components"
      Height          =   252
      Left            =   5040
      TabIndex        =   11
      Top             =   120
      Width           =   2172
   End
   Begin VB.Label Label3 
      Caption         =   "PADS Logic Selected Component Count"
      Height          =   252
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   4692
   End
   Begin VB.Label Label2 
      Caption         =   "PADS Logic Component Count"
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
      Caption         =   "PADS Logic Document Name"
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
' Client.BAS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This is a part of the PADS Logic OLE Automation server Client sample.
' Copyright Mentor Graphics Corporation 2007
' All rights reserved.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This source code is only intended as a supplement to the PADS Logic
' Automation Server API Help file.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim WithEvents powerLogicApp As PowerLogic.Application
Attribute powerLogicApp.VB_VarHelpID = -1
Dim WithEvents powerLogicDoc As PowerLogic.Document
Attribute powerLogicDoc.VB_VarHelpID = -1

' This global flag is used to work around a VB problem:
' When the listbox Selected property is used, VB automatically
' will call the Click event handler. That is wrong.
Dim bDoNotListboxClickCallbackUponSelectionChange As Boolean

Private Sub RefreshGeneralInfos()
    ' If we are not connected, exit procedure
    If powerLogicDoc Is Nothing Then
        EdtDocName.Text = "???????"
        EdtDocComps.Text = "???????"
        EdtDocSelectedComps.Text = "???????"
        Exit Sub
    End If

    On Error GoTo OnErrorServerAccess
    
    ' Output the PADS Logic server Document name to our editbox
    EdtDocName.Text = powerLogicApp.ActiveDocument.Name
    
    ' Output the PADS Logic component objects
    mcount = powerLogicDoc.GetObjects(1).Count
    EdtDocComps.Text = mcount & " component object(s)"
    
    ' Output the PADS Logic selected component objects
    mcount = powerLogicDoc.GetObjects(1, "", True).Count
    EdtDocSelectedComps.Text = mcount & " selected component object(s)"

Exit Sub        ' Exit to avoid handler.

OnErrorServerAccess:
    Msg = "Are you sure you are connected??"
    MsgBox Msg, , "Client"

End Sub

Public Sub RefreshObjectsInfos()
    ' Clear the listbox
    LstDocObjects.Clear
    
    ' If we are not connected, exit procedure
    If powerLogicApp Is Nothing Then Exit Sub

    On Error GoTo OnErrorServerAccess
    
    ' Get all PADS Logic objects of type COMPONENT from PADS Logic database
    ' and output each component in the listbox
    countComp = 0
    For Each nextComp In powerLogicDoc.GetObjects(1)
        LstDocObjects.AddItem nextComp.Name
        countComp = countComp + 1
        Status.Caption = "Loading " & countComp & " component(s) from PADS Logic..."
        Status.Refresh
    Next nextComp
    Status.Caption = ""
Exit Sub        ' Exit to avoid handler.

OnErrorServerAccess:
    Msg = "Are you sure you are connected??"
    MsgBox Msg, , "Client"
    
End Sub

Public Sub RefreshSelectionInfos()
    ' If we are not connected, exit procedure
    If powerLogicDoc Is Nothing Then Exit Sub

    On Error GoTo OnErrorServerAccess
    
    bDoNotListboxClickCallbackUponSelectionChange = True
    
    ' Clear the selection in the listbox
    ' LstDocObjects.ListIndex = -1 ' This doesn't work on multi-select listboxes!!!
    For I = 0 To LstDocObjects.ListCount - 1
        LstDocObjects.Selected(I) = False
    Next I
    
    ' For each component selected in PADS Logic...
    For Each nextComp In powerLogicDoc.GetObjects(1, "", True)
        ' ... search this component in our listbox in the listbox...
        For J = 0 To LstDocObjects.ListCount - 1
            ' ... if this is the one select it, otherwise, unselect it.
            If LstDocObjects.List(J) = nextComp.Name Then
                LstDocObjects.Selected(J) = True
            End If
        Next J
    Next nextComp
    
    bDoNotListboxClickCallbackUponSelectionChange = False
    
Exit Sub        ' Exit to avoid handler.

OnErrorServerAccess:
    Msg = "Are you sure you are connected??"
    MsgBox Msg, , "Client"
    
End Sub

Public Sub LocateSelection()
    'look for selection in current sheet: if it has selection then do nothing
    If powerLogicDoc.ActiveSheet.GetObjects(plogObjectTypeComponent, , True).Count > 0 Then Exit Sub
    For Each Sheet In powerLogicDoc.Sheets
        If Sheet.GetObjects(plogObjectTypeComponent, , True).Count > 0 Then
            Sheet.Activate
            Exit Sub
        End If
    Next
End Sub

Private Sub BtnConnect_Click()
    ' Enable error-handling routine.
    On Error GoTo OnErrorGetObject
    
    ' Connect to a running instance of PADS Logic server
    Set powerLogicApp = GetObject(, "PowerLogic.Application")
    Set powerLogicDoc = powerLogicApp.ActiveDocument
    
    ' Enable and disable correct buttons
    BtnConnect.Enabled = False
    BtnDisconnect.Enabled = True
    
    ' Set status line text
    Status.Caption = "Connected to PADS Logic."
    
Exit Sub        ' Exit to avoid handler.

OnErrorGetObject:
    Msg = "Cannot connect to a running PADS Logic Server!"
    MsgBox Msg, , "Client"
    
End Sub

Private Sub BtnDisconnect_Click()
    ' Disconnect from PADS Logic server
    Set powerLogicDoc = Nothing
    Set powerLogicApp = Nothing
    ' Enable and disable correct buttons
    BtnConnect.Enabled = True
    BtnDisconnect.Enabled = False
         ' Set status line text
    Status.Caption = "Not connected to PADS Logic."
End Sub
Private Sub BtnOk_Click()
    End
End Sub

Private Sub BtnRefresh_Click()
    RefreshGeneralInfos
    RefreshObjectsInfos
    RefreshSelectionInfos
End Sub

Private Sub Form_Load()
    BtnConnect.Enabled = True
    BtnDisconnect.Enabled = False
    Status.Caption = "Not connected to PADS Logic."
End Sub


Private Sub powerLogicApp_OpenDocument(ByVal Doc As PowerLogic.Document)
    RefreshGeneralInfos
    RefreshObjectsInfos
    RefreshSelectionInfos
End Sub

Private Sub powerLogicApp_Quit()
    BtnDisconnect_Click
    BtnRefresh_Click
End Sub

Private Sub powerLogicDoc_SelectionChange()
    If bDoNotListboxClickCallbackUponSelectionChange = True Then Exit Sub
    RefreshGeneralInfos
    RefreshSelectionInfos
End Sub

Private Sub LstDocObjects_Click()
    ' If we are not connected, exit procedure
    If powerLogicDoc Is Nothing Then Exit Sub
    
    If bDoNotListboxClickCallbackUponSelectionChange = True Then Exit Sub
    
    bDoNotListboxClickCallbackUponSelectionChange = True
    
    MousePointer = 11
    
    ' If there is at least 1 item selected in the listbox...
    If LstDocObjects.SelCount > 0 Then
        ' .. reset the server selection list.
        powerLogicDoc.SelectObjects 9999, "", False
        ' For each item in the listbox...
        For I = 0 To LstDocObjects.ListCount - 1
            ' ... if this item is selected...
            If LstDocObjects.Selected(I) = True Then
                ' ... select its equivalent in PADS Logic.
                objectName = LstDocObjects.List(I)
                powerLogicDoc.SelectObjects 1, objectName, True
            End If
        Next I
    End If

    LocateSelection

    bDoNotListboxClickCallbackUponSelectionChange = False

    MousePointer = 0
End Sub


Private Sub LstDocObjects_DblClick()
    ' If we are not connected, exit procedure
    If powerLogicDoc Is Nothing Then Exit Sub

    MousePointer = 11
    
    ' If there is at least 1 item selected in the listbox...
    If LstDocObjects.SelCount > 0 Then
        ' For each item in the listbox...
        For I = 0 To LstDocObjects.ListCount - 1
            ' ... if this item is selected...
            If LstDocObjects.Selected(I) = True Then
                objectName = LstDocObjects.List(I)
                ' ... retrieve it.
                Dim powerLogicComp As Object
                Set powerLogicComp = powerLogicDoc.Components(objectName)
                    
                ' Get the object name
                Dim string1 As String
                string1 = "Object Name: " & powerLogicComp.Name
          
                ' Get the object part type
                Dim string2 As String
                string2 = "Object Type: " & powerLogicComp.PartType
            
                ' Get the object decal
                Dim string3 As String
                string3 = "Object PCB Decal: " & powerLogicComp.PCBDecal
            
                ' Get the object gates
                Dim string4 As String
                string4 = "Object Gates: " & powerLogicComp.Gates.Count + powerLogicComp.UnusedGates.Count & " gates."
            
                ' Get the object unused gates
                Dim string5 As String
                string5 = "Object Unused Gates: " & powerLogicComp.UnusedGates.Count & " gates."
                
                ' Get the object pin count
                Dim string6 As String
                string6 = "Object Pins: " & powerLogicComp.Pins.Count & " pins."
            
                ' Display all these infos in a dialog box
                msgString = string1 & vbCrLf & string2 & vbCrLf & string3 & vbCrLf & string4 & vbCrLf & string5 & vbCrLf & string6
                MsgBox msgString, , "Client"
                Set powerLogicComp = Nothing
            End If
        Next
    End If
    MousePointer = 0
End Sub



