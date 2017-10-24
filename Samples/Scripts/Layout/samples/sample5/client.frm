VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Sample5: PowerPCB Intelligent Report Plug-In"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6240
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.CommandButton BTNShow 
      Caption         =   "Show"
      Height          =   320
      Left            =   4920
      TabIndex        =   28
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton BTNHelp 
      Caption         =   "Help"
      Height          =   372
      Left            =   6120
      TabIndex        =   12
      Top             =   1680
      Width           =   1332
   End
   Begin VB.CommandButton BTNReset 
      Caption         =   "Reset Settings"
      Height          =   372
      Left            =   6120
      TabIndex        =   8
      ToolTipText     =   "Delete saved settings and reset to default ones"
      Top             =   720
      Width           =   1332
   End
   Begin VB.Frame Frame3 
      Caption         =   "Report configuration"
      Height          =   3132
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   5895
      Begin VB.Frame Frame5 
         Caption         =   "Others"
         Height          =   1812
         Left            =   3240
         TabIndex        =   24
         Top             =   1200
         Width           =   2535
         Begin VB.CheckBox CHKNoTitles 
            Caption         =   "No field titles"
            Height          =   252
            Left            =   120
            TabIndex        =   27
            ToolTipText     =   "Do not report field names"
            Top             =   240
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox CHKGroup 
            Caption         =   "Group by the first selected field in the listbox"
            Height          =   675
            Left            =   120
            TabIndex        =   26
            ToolTipText     =   "Group report by the first checked field"
            Top             =   600
            Width           =   2295
         End
         Begin VB.CheckBox CHKDataLink 
            Caption         =   "Data Link (Excel only)"
            Height          =   375
            Left            =   120
            TabIndex        =   25
            ToolTipText     =   "Enable data link between the report and PowerPCB"
            Top             =   1320
            Width           =   2295
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Fields"
         Height          =   2772
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   3012
         Begin VB.CommandButton BTNFieldsSelectNone 
            Caption         =   "Uncheck all"
            Height          =   360
            Left            =   1920
            TabIndex        =   21
            ToolTipText     =   "Uncheck all fields"
            Top             =   2280
            Width           =   972
         End
         Begin VB.CommandButton BTNFieldsSelectAll 
            Caption         =   "Check all"
            Height          =   360
            Left            =   1920
            TabIndex        =   20
            ToolTipText     =   "Check all fields"
            Top             =   1920
            Width           =   972
         End
         Begin VB.CommandButton BTNMoveDown 
            Caption         =   "Move down"
            Height          =   360
            Left            =   1920
            TabIndex        =   19
            ToolTipText     =   "Move selected field down"
            Top             =   600
            Width           =   972
         End
         Begin VB.CommandButton BTNMoveUp 
            Caption         =   "Move up"
            Height          =   360
            Left            =   1920
            TabIndex        =   18
            ToolTipText     =   "Move selected field up"
            Top             =   240
            Width           =   972
         End
         Begin VB.ListBox LSTFields 
            Height          =   2310
            ItemData        =   "client.frx":0000
            Left            =   120
            List            =   "client.frx":0002
            Style           =   1  'Checkbox
            TabIndex        =   17
            ToolTipText     =   "Report fields to output"
            Top             =   240
            Width           =   1692
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Parts"
         Height          =   852
         Left            =   3240
         TabIndex        =   13
         Top             =   240
         Width           =   2535
         Begin VB.OptionButton OPTSelected 
            Caption         =   "Selected (9999 parts)"
            Height          =   252
            Left            =   120
            TabIndex        =   15
            ToolTipText     =   "Report on selected parts of the design only"
            Top             =   480
            Width           =   1812
         End
         Begin VB.OptionButton OPTAll 
            Caption         =   "All (9999 parts)"
            Height          =   252
            Left            =   120
            TabIndex        =   14
            ToolTipText     =   "Report on all parts of the design"
            Top             =   240
            Width           =   1452
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Report file"
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   5895
      Begin VB.ComboBox CBXReportFormat 
         Height          =   315
         ItemData        =   "client.frx":0004
         Left            =   720
         List            =   "client.frx":001A
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Format in which the report is output"
         Top             =   600
         Width           =   5055
      End
      Begin VB.TextBox EDTReportFile 
         Height          =   288
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "File name for the report output"
         Top             =   240
         Width           =   5655
      End
      Begin VB.Label Label1 
         Caption         =   "Format:"
         Height          =   252
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   612
      End
   End
   Begin VB.CommandButton BTNClose 
      Caption         =   "Close"
      Height          =   372
      Left            =   6120
      TabIndex        =   2
      ToolTipText     =   "Close the dialog"
      Top             =   1200
      Width           =   1332
   End
   Begin VB.CommandButton BTNOk 
      Caption         =   "Run Report"
      Height          =   372
      Left            =   6120
      TabIndex        =   1
      ToolTipText     =   "Generate the report"
      Top             =   240
      Width           =   1332
   End
   Begin VB.Frame Frame1 
      Caption         =   "PowerPCB design file"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.ComboBox CBXAssOpt 
         Height          =   288
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   22
         ToolTipText     =   "Assembly option to report"
         Top             =   600
         Width           =   3252
      End
      Begin VB.ComboBox CBXPowerPCBFile 
         Height          =   288
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "PCB design currently open in PowerPCB"
         Top             =   240
         Width           =   4572
      End
      Begin VB.CommandButton BTNBrowse 
         Caption         =   "Browse..."
         Height          =   320
         Left            =   4800
         TabIndex        =   7
         ToolTipText     =   "Browse and open a new PowerPCB design"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Assembly option:"
         Height          =   252
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   1332
      End
   End
   Begin VB.Label Status 
      Caption         =   "Status"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5760
      Width           =   7335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' Allow highest degree of type checking

' Declare Win32 functions that we use here. These 2 are used to get the defaukt
' application registered to handle a given file type (TXT and HTML in our case).
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' Declare and instanciate connection (with events) with PowerPCB
Dim WithEvents ppcbApp As PowerPCB.Application
Attribute ppcbApp.VB_VarHelpID = -1
Dim WithEvents ppcbDoc As PowerPCB.Document
Attribute ppcbDoc.VB_VarHelpID = -1

' Declare a global index since we use temp indexes for "for/next" loops very often.
Dim index As Long


Private Sub SaveRegistrySettings()
    ' CBXPowerPCBFile
    For index = 0 To CBXPowerPCBFile.ListCount - 1
        If index < 10 Then SaveSetting "PowerPCB Sample5", "Startup", "PCBFile" & index + 1, CBXPowerPCBFile.List(index)
    Next index
        
    ' EDTReportFile.Text
    SaveSetting "PowerPCB Sample5", "Startup", "Report File", EDTReportFile.Text
    
    ' CBXReportFormat
    SaveSetting "PowerPCB Sample5", "Startup", "Report Format", CBXReportFormat.List(CBXReportFormat.ListIndex)
    
    ' OPTSelected
    SaveSetting "PowerPCB Sample5", "Startup", "Selected", OPTSelected.Value

    ' CHKNoTitles
    SaveSetting "PowerPCB Sample5", "Startup", "Titles", CHKNoTitles.Value

    ' CHKDataLink
    SaveSetting "PowerPCB Sample5", "Startup", "DataLink", CHKDataLink.Value

    ' CHKGroup
    SaveSetting "PowerPCB Sample5", "Startup", "Group", CHKGroup.Value

    ' LSTFields
    For index = 0 To LSTFields.ListCount - 1
        SaveSetting "PowerPCB Sample5", "Startup", "Field" & index + 1 & "Name", LSTFields.List(index)
        SaveSetting "PowerPCB Sample5", "Startup", "Field" & index + 1 & "Value", LSTFields.Selected(index)
    Next index

End Sub

Private Sub RestoreRegistrySettings()
    Dim strSetting As String
    
    ' CBXPowerPCBFile
    CBXPowerPCBFile.Clear
    For index = 0 To 10
        strSetting = GetSetting("PowerPCB Sample5", "Startup", "PCBFile" & index + 1, "")
        If strSetting <> "" Then CBXPowerPCBFile.AddItem (strSetting)
    Next index
    
    ' EDTReportFile.Text
    EDTReportFile.Text = GetSetting("PowerPCB Sample5", "Startup", "Report File", "")
    
    ' CBXReportFormat
    CBXReportFormat.ListIndex = 0
    For index = 0 To CBXReportFormat.ListCount - 1
        If CBXReportFormat.List(index) = GetSetting("PowerPCB Sample5", "Startup", "Report Format", "") Then
            CBXReportFormat.ListIndex = index
        End If
    Next index

    ' OPTSelected
    OPTAll.Value = False
    OPTSelected.Value = False
    strSetting = GetSetting("PowerPCB Sample5", "Startup", "Selected", "")
    If strSetting = "True" Then OPTSelected.Value = True Else OPTAll.Value = True
    
    ' CHKNoTitles
    CHKNoTitles.Value = 0
    strSetting = GetSetting("PowerPCB Sample5", "Startup", "Titles", "")
    If strSetting = "1" Then CHKNoTitles.Value = 1

    ' CHKDataLink
    CHKDataLink.Value = 0
    strSetting = GetSetting("PowerPCB Sample5", "Startup", "DataLink", "")
    If strSetting = "1" Then CHKDataLink.Value = 1

    ' CHKGroup
    CHKGroup.Value = 0
    strSetting = GetSetting("PowerPCB Sample5", "Startup", "Group", "")
    If strSetting = "1" Then CHKGroup.Value = 1

    ' LSTFields
    LSTFields.Clear
    For index = 0 To 15
        strSetting = GetSetting("PowerPCB Sample5", "Startup", "Field" & index + 1 & "Name", "")
        If strSetting <> "" Then
            LSTFields.AddItem (strSetting)
        Else
            Select Case index
                Case 0:  LSTFields.AddItem ("RefDes")
                Case 1:  LSTFields.AddItem ("PartType")
                Case 2:  LSTFields.AddItem ("Decal")
                Case 3:  LSTFields.AddItem ("Logic")
                Case 4:  LSTFields.AddItem ("Description")
                Case 5:  LSTFields.AddItem ("Attributes")
                Case 6:  LSTFields.AddItem ("IsSMD")
                Case 7:  LSTFields.AddItem ("Pins")
                Case 8:  LSTFields.AddItem ("X")
                Case 9:  LSTFields.AddItem ("Y")
                Case 10: LSTFields.AddItem ("Placed")
                Case 11: LSTFields.AddItem ("Orientation")
                Case 12: LSTFields.AddItem ("Layer")
                Case 13: LSTFields.AddItem ("Glued")
                Case 14: LSTFields.AddItem ("Installed")
                Case 15: LSTFields.AddItem ("Substituted")
            End Select
        End If
        strSetting = GetSetting("PowerPCB Sample5", "Startup", "Field" & index + 1 & "Value", "")
        If strSetting = "False" Then LSTFields.Selected(index) = False Else LSTFields.Selected(index) = True
    Next index
End Sub

Private Sub Form_Load()
    On Error GoTo PPCBConnectError
    Set ppcbApp = GetObject(, "PowerPCB.Application")
    Set ppcbDoc = ppcbApp.ActiveDocument
    On Error GoTo 0
    RestoreRegistrySettings
    RefreshDialog
    Exit Sub
    
PPCBConnectError:
    MsgBox "Cannot connect to a running PowerPCB server. Aborting...", , "Sample5"
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveRegistrySettings
    On Error GoTo errorHandler:
    Set ppcbDoc = Nothing
    Set ppcbApp = Nothing
errorHandler:
End Sub

Private Sub BTNReset_Click()
    On Error GoTo errorNoSetting
    DeleteSetting "PowerPCB Sample5", "Startup"
    On Error GoTo 0
errorNoSetting:
    RestoreRegistrySettings
    RefreshDialog
End Sub

Private Sub CBXPowerPCBFile_Click()
    Dim newFile As String: newFile = CBXPowerPCBFile.List(CBXPowerPCBFile.ListIndex)
    If newFile <> "" Then
        If newFile <> ppcbDoc.FullName Then
            Status.Caption = "Opening " & newFile & ". Please wait...": Status.Refresh
            Form1.MousePointer = vbHourglass: Form1.Refresh
            ppcbApp.OpenDocument (newFile)
            Form1.MousePointer = vbDefault: Form1.Refresh
        End If
    End If
End Sub

Private Sub BTNBrowse_Click()
    CommonDialog1.CancelError = False
    CommonDialog1.DialogTitle = "Open PowerPCB file"
    CommonDialog1.Filter = "PowerPCB Files (*.pcb)|*.pcb"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist
    CommonDialog1.InitDir = ppcbApp.DefaultFilePath
    CommonDialog1.ShowOpen
    If CommonDialog1.filename <> "" Then
        Status.Caption = "Opening " & CommonDialog1.filename & ". Please wait...": Status.Refresh
        Form1.MousePointer = vbHourglass: Form1.Refresh
        ppcbApp.OpenDocument (CommonDialog1.filename)
        Form1.MousePointer = vbDefault: Form1.Refresh
    End If
End Sub
 
Private Sub CBXReportFormat_Click()
    RefreshDialog
End Sub

Private Sub BTNMoveUp_Click()
    Dim moveIndex As String: moveIndex = LSTFields.ListIndex
    If moveIndex >= 0 Then
        If moveIndex > 0 Then
            Dim savTitle As String: savTitle = LSTFields.List(moveIndex)
            Dim savChecked As Boolean: savChecked = LSTFields.Selected(moveIndex)
            LSTFields.RemoveItem (moveIndex)
            moveIndex = moveIndex - 1
            LSTFields.AddItem savTitle, moveIndex
            LSTFields.Selected(moveIndex) = savChecked
            LSTFields.ListIndex = moveIndex
        End If
    End If
End Sub
 
Private Sub BTNMoveDown_Click()
    Dim moveIndex As String: moveIndex = LSTFields.ListIndex
    If moveIndex >= 0 Then
        If moveIndex < LSTFields.ListCount - 1 Then
            Dim savTitle As String: savTitle = LSTFields.List(moveIndex)
            Dim savChecked As Boolean: savChecked = LSTFields.Selected(moveIndex)
            LSTFields.RemoveItem (moveIndex)
            moveIndex = moveIndex + 1
            LSTFields.AddItem savTitle, moveIndex
            LSTFields.Selected(moveIndex) = savChecked
            LSTFields.ListIndex = moveIndex
        End If
    End If
End Sub

Private Sub BTNFieldsSelectAll_Click()
    Dim savTopIndex As Long: savTopIndex = LSTFields.TopIndex
    Dim savListIndex As Long: savListIndex = LSTFields.ListIndex
    For index = 0 To LSTFields.ListCount - 1
        LSTFields.Selected(index) = True
    Next index
    LSTFields.TopIndex = savTopIndex
    LSTFields.ListIndex = savListIndex
End Sub

Private Sub BTNFieldsSelectNone_Click()
    Dim savTopIndex As Long: savTopIndex = LSTFields.TopIndex
    Dim savListIndex As Long: savListIndex = LSTFields.ListIndex
    For index = 0 To LSTFields.ListCount - 1
        LSTFields.Selected(index) = False
    Next index
    LSTFields.TopIndex = savTopIndex
    LSTFields.ListIndex = savListIndex
End Sub

Private Sub CHKDataLink_Click()
    RefreshDialog
End Sub

Private Sub CHKGroup_Click()
    RefreshDialog
End Sub

Private Sub OPTAll_Click()
    RefreshDialog
End Sub

Private Sub OPTSelected_Click()
    RefreshDialog
End Sub

Private Sub ppcbApp_OpenDocument(ByVal Doc As PowerPCB.Document)
    RefreshDialog
End Sub

Private Sub ppcbDoc_Save()
    RefreshDialog
End Sub

Private Sub ppcbDoc_SelectionChange()
    RefreshDialog
End Sub

Private Sub ppcbApp_Quit()
    BTNClose_Click
End Sub


Private Sub BTNShow_Click()
    ppcbApp.Visible = False
    ppcbApp.Visible = True
End Sub

Private Sub BTNClose_Click()
    Form_Unload (0) ' Somehow, End doesn't call it, so we call it manually
    End
End Sub

Private Sub BTNHelp_Click()
    CommonDialog1.HelpFile = "PPCBOLE.HLP"
    CommonDialog1.HelpCommand = cdlHelpContext
    CommonDialog1.HelpContext = 999999 '
    CommonDialog1.ShowHelp
End Sub

Private Sub RefreshDialog()
    Dim bNeedsUpdate As Boolean
    
    ' Update the PowerPCB file name if needed
    bNeedsUpdate = CBXPowerPCBFile.List(CBXPowerPCBFile.ListIndex) <> ppcbDoc.FullName
    If bNeedsUpdate = True Then
        CBXPowerPCBFile.ListIndex = -1
        For index = 0 To CBXPowerPCBFile.ListCount - 1
            If CBXPowerPCBFile.List(index) = ppcbDoc.FullName Then
                CBXPowerPCBFile.ListIndex = index
            End If
        Next index
        If CBXPowerPCBFile.ListIndex = -1 Then
            CBXPowerPCBFile.AddItem ppcbDoc.FullName, 0
            CBXPowerPCBFile.ListIndex = 0
        End If
    End If
    
    ' Update the PowerPCB assembly option combobox
    bNeedsUpdate = False
    Dim assOpts As Object: Set assOpts = ppcbDoc.AssemblyOptions
    If assOpts.Count = 0 Then bNeedsUpdate = True
    For index = 1 To assOpts.Count
        If Left(assOpts.Item(index).Name, InStr(1, assOpts.Item(index).Name, ":") - 1) <> CBXAssOpt.List(index) Then
            bNeedsUpdate = True
        End If
    Next index
    If bNeedsUpdate = True Then
        CBXAssOpt.Clear
        If assOpts.Count = 0 Then
            CBXAssOpt.Enabled = False
            CBXAssOpt.AddItem ("<<none>>")
            CBXAssOpt.ListIndex = 0
        Else
            CBXAssOpt.Enabled = True
            CBXAssOpt.AddItem ("<<Default>>")
            Dim nextDoc As Object
            For Each nextDoc In assOpts
                CBXAssOpt.AddItem (Left(nextDoc.Name, InStr(1, nextDoc.Name, ":") - 1))
            Next nextDoc
            If CBXAssOpt.ListIndex = -1 Then CBXAssOpt.ListIndex = 0
        End If
    End If
    
    ' Configuration panel settings
    LSTFields.ListIndex = -1
    EDTReportFile.Text = ""
    CHKGroup.Enabled = False
    CHKNoTitles.Enabled = False
    CHKDataLink.Enabled = False
    If ppcbDoc.FullName <> "" Then
        ' Report name, based on report type
        Dim reportExt As String: reportExt = ""
        Select Case CBXReportFormat.ListIndex
            Case 0: reportExt = "txt"
            Case 1: reportExt = "csv"
            Case 2: reportExt = "xls"
            Case 3: reportExt = "doc"
            Case 4: reportExt = "htm"
            Case 5: reportExt = "htm"
        End Select
        
        ' COnfiguration panel
        CHKDataLink.Enabled = reportExt = "xls"
        CHKGroup.Enabled = CHKDataLink.Enabled = False Or CHKDataLink.Value = False
        CHKNoTitles.Enabled = CHKDataLink.Enabled = False Or CHKDataLink.Value = False
        
        
        Dim theDotPos As Long: theDotPos = InStr(UCase(ppcbDoc.FullName), ".PCB")
        If theDotPos = 0 Then theDotPos = InStr(UCase(ppcbDoc.FullName), ".JOB")
        If theDotPos = 0 Then theDotPos = InStr(UCase(ppcbDoc.FullName), ".ASC")
        If theDotPos = 0 Then
            EDTReportFile.Text = ppcbDoc.FullName
        Else
            EDTReportFile.Text = Left(ppcbDoc.FullName, theDotPos - 1) & "_report"
        End If
        
        For index = 1 To 1000
            If Dir(EDTReportFile.Text & index & "." & "*") = "" Then
                EDTReportFile.Text = EDTReportFile.Text & index & "." & reportExt
                Exit For
            End If
        Next index
    End If
    
    Dim nbParts As Long: nbParts = ppcbDoc.Components.Count
    Dim nbSelParts As Long: nbSelParts = ppcbDoc.GetObjects(1, "", True).Count
    
    OPTAll.Caption = "All (" & nbParts & " parts)"
    OPTSelected.Caption = "Selected (" & nbSelParts & " parts)"
    
    BTNOk.Enabled = (ppcbDoc.FullName <> "" And nbParts > 0 And (OPTSelected.Value = False Or nbSelParts > 0))
    If ppcbDoc.FullName <> "" Then
        Status.Caption = "Connected to " & ppcbDoc.Name
    Else
        Status.Caption = "Connected to PowerPCB (no document loaded)"
    End If

End Sub

Private Sub BTNOk_Click()
    
    ' Check that the output report file (in the output report file editbox)
    ' doesn't exist already. This will avoid us to later deal with Excel, Word...
    ' exceptions when trying to save to a file that exist and might be already
    ' opened. If the file already exist, just report it and return.
    If Dir(EDTReportFile.Text) <> "" Then
        MsgBox "The output report file " & EDTReportFile.Text & " already exist. Please choose another file name.", , "Sample5"
        EDTReportFile.SetFocus
        Exit Sub
    End If
    
    ' Set proper assembly option
    Dim currentDocument As Object
    If CBXAssOpt.ListIndex > 0 Then
        Set currentDocument = ppcbDoc.AssemblyOptions(CBXAssOpt.List(CBXAssOpt.ListIndex))
    Else
        Set currentDocument = ppcbDoc
    End If
    
    ' Get all components in the current design and calculate the component count
    Dim objComps As Object
    If OPTSelected.Value = True Then
        Set objComps = currentDocument.GetObjects(1, "", True)
    Else
        Set objComps = currentDocument.Components
    End If
    Dim nbComps As Long: nbComps = objComps.Count
    
    ' Prepare the main data array. This array is a table listing all components
    ' values for the fields that has been selected to output. We hardcode a max
    ' number of columns to 20 just to make sure we have space.
    Dim reportArrayRaw() As String
    ReDim reportArrayRaw(20, nbComps + 1) ' the +1 is because the first row is the title row
    
    ' Output the field (column) names in the array
    Dim nbFields As Long: nbFields = 0
    For index = 0 To LSTFields.ListCount - 1
        If LSTFields.Selected(index) = True Then
            reportArrayRaw(nbFields, 0) = LSTFields.List(index)
            nbFields = nbFields + 1
        End If
    Next index
    
    ' If we need to group by something, add a Count field at the end
    If (CHKGroup.Enabled = True And CHKGroup.Value = 1) Then
        reportArrayRaw(nbFields, 0) = "Count"
    End If

    ' Exit if there is no field selected (that would be an empty report)
    If nbFields = 0 Then MsgBox "You need at least one field to output", , "Sample5": Exit Sub
    
    ' Check that if Excel DataLink is checked, the RefDes field in checked
    Dim bRefDesFound As Boolean: bRefDesFound = False
    For index = 0 To nbFields - 1
        If reportArrayRaw(index, 0) = "RefDes" Then bRefDesFound = True
    Next index
    If CHKDataLink.Enabled = True And CHKDataLink.Value = 1 And bRefDesFound = False Then
        MsgBox "You need to select the RefDes field for dynamically linked Excel reports", , "Sample5"
        Exit Sub
    End If
    
    ' Check that the application chosen to output the report is installed
    ' on that system. Here, I can't find how to check for the presence of
    ' a server without actually starting. So we start it...
    Dim hostApp As Object
    Dim hostAppName As String: hostAppName = ""
    Select Case CBXReportFormat.ListIndex
        Case 0
            Set hostApp = Nothing
            hostAppName = "NotePad"
        Case 1
            Set hostApp = Nothing
            hostAppName = "NotePad"
        Case 2
            On Error GoTo noExcel
            Status.Caption = "Starting Microsoft Excel...": Status.Refresh
            Set hostApp = CreateObject("Excel.Application")
            hostAppName = hostApp.Name
        Case 3
            On Error GoTo noWord
            Status.Caption = "Starting Microsoft Word...": Status.Refresh
            Set hostApp = CreateObject("Word.Application")
            hostAppName = hostApp.Name
        Case 4
            ' Create a dummy HTML file (required when using FindExecutable)
            Dim dummyFile As String: dummyFile = ppcbApp.DefaultFilePath & "\dummy html file ah ah ah.htm"
            Dim fileHandle As Long: fileHandle = FreeFile
            Open dummyFile For Output As #fileHandle
                Write #fileHandle, "<HTML></HTML>"
            Close #fileHandle
            Dim browserEXE As String * 255: browserEXE = Space(255)
            Dim retval As Long
            Dim directory As String
            retval = FindExecutable(dummyFile, directory, browserEXE)
            Kill dummyFile
            If retval <= 32 Or IsEmpty(browserEXE) Then GoTo noInternetBrowser
            Set hostApp = Nothing
            hostAppName = Trim(browserEXE)
        Case 5
            On Error GoTo noInternetExplorer
            Status.Caption = "Starting Microsoft Internet Explorer...": Status.Refresh
            Set hostApp = CreateObject("InternetExplorer.Application")
            hostAppName = hostApp.Name
    End Select
    On Error GoTo 0
    
    ' Output the component data (only the fields selected)
    Dim arrayIndex As Long: arrayIndex = 1
    Dim statusCounter As Long: statusCounter = 1
    Dim nextComp As Object
    ' For each component in the current PowerPCB design...
    For Each nextComp In objComps
        Dim tmpArrayRaw() As String
        ReDim tmpArrayRaw(20)
        ' ... get all the properties we want and output in correct field (column)
        On Error Resume Next
        For index = 0 To nbFields - 1
            tmpArrayRaw(index) = "?????"
            Select Case reportArrayRaw(index, 0)
                Case "RefDes":      tmpArrayRaw(index) = nextComp.Name: If tmpArrayRaw(index) = "" Then tmpArrayRaw(index) = "None"
                Case "PartType":    tmpArrayRaw(index) = nextComp.PartType: If tmpArrayRaw(index) = "" Then tmpArrayRaw(index) = "None"
                Case "Decal":       tmpArrayRaw(index) = nextComp.Decal: If tmpArrayRaw(index) = "" Then tmpArrayRaw(index) = "None"
                Case "Logic":       tmpArrayRaw(index) = nextComp.PartTypeLogic: If tmpArrayRaw(index) = "" Then tmpArrayRaw(index) = "None"
                Case "Description": tmpArrayRaw(index) = nextComp.Attributes("DESCRIPTION"): If tmpArrayRaw(index) = "" Then tmpArrayRaw(index) = "None"
                Case "Attributes":  tmpArrayRaw(index) = nextComp.Attributes.Count
                Case "IsSMD":       If nextComp.IsSMD = False Then tmpArrayRaw(index) = "False" Else tmpArrayRaw(index) = "True"
                Case "Pins":        tmpArrayRaw(index) = nextComp.Pins.Count
                Case "X":           tmpArrayRaw(index) = nextComp.PositionX
                Case "Y":           tmpArrayRaw(index) = nextComp.PositionY
                Case "Placed":      If nextComp.Placed = False Then tmpArrayRaw(index) = "False" Else tmpArrayRaw(index) = "True"
                Case "Orientation": tmpArrayRaw(index) = nextComp.Orientation
                Case "Layer":       tmpArrayRaw(index) = nextComp.LayerName: If tmpArrayRaw(index) = "" Then tmpArrayRaw(index) = "None"
                Case "Glued":       If nextComp.Glued = False Then tmpArrayRaw(index) = "False" Else tmpArrayRaw(index) = "True"
                Case "Installed":   If nextComp.Installed = False Then tmpArrayRaw(index) = "False" Else tmpArrayRaw(index) = "True"
                Case "Substituted": If nextComp.Substituted = False Then tmpArrayRaw(index) = "False" Else tmpArrayRaw(index) = "True"
            End Select
        Next index
        On Error GoTo 0
        
        ' If user selected to group report, handle it
        If (CHKGroup.Enabled = True And CHKGroup.Value = 1) Then
            ' Lookup that group key (the group key is always the first column)
            Dim keyIndex As Long: keyIndex = -1
            For index = 1 To arrayIndex - 1
                If reportArrayRaw(0, index) = tmpArrayRaw(0) Then keyIndex = index
            Next index
            ' If that group key has not yet been entered, add it
            If keyIndex = -1 Then
                For index = 0 To nbFields - 1
                    reportArrayRaw(index, arrayIndex) = tmpArrayRaw(index)
                Next index
                reportArrayRaw(nbFields, arrayIndex) = 1
                arrayIndex = arrayIndex + 1
            ' Otherwise add a new record
            Else
                For index = 1 To nbFields - 1
                    reportArrayRaw(index, keyIndex) = reportArrayRaw(index, keyIndex) & ", " & tmpArrayRaw(index)
                Next index
                reportArrayRaw(nbFields, keyIndex) = 1 + reportArrayRaw(nbFields, keyIndex)
            End If
        Else
            For index = 0 To nbFields - 1
                reportArrayRaw(index, arrayIndex) = tmpArrayRaw(index)
            Next index
            arrayIndex = arrayIndex + 1
        End If
        
        Status.Caption = "Generating data (" & CInt(100 * statusCounter / nbComps) & "%)...": Status.Refresh
        statusCounter = statusCounter + 1
    Next nextComp
    
    If (CHKGroup.Enabled = True And CHKGroup.Value = 1) Then
        nbFields = nbFields + 1 ' Include the count field
    End If

    Dim reportName As String: reportName = EDTReportFile.Text
    Select Case CBXReportFormat.ListIndex
        Case 0: Call OutputText(reportName, reportArrayRaw, nbFields, CHKNoTitles.Value, arrayIndex)
        Case 1: Call OutputCSV(reportName, reportArrayRaw, nbFields, CHKNoTitles.Value, arrayIndex)
        Case 2
            If CHKDataLink.Value = False Then
                Call OutputExcel(hostApp, reportName, reportArrayRaw, nbFields, CHKNoTitles.Value, arrayIndex)
            Else
                Call OutputExcelDynamic(hostApp, reportName, reportArrayRaw, nbFields, 0, arrayIndex)
            End If
        Case 3: Call OutputWord(hostApp, reportName, reportArrayRaw, nbFields, CHKNoTitles.Value, arrayIndex)
        Case 4: Call OutputInternetBrowser(hostAppName, reportName, reportArrayRaw, nbFields, CHKNoTitles.Value, arrayIndex)
        Case 5: Call OutputIE4(hostApp, reportName, reportArrayRaw, nbFields, CHKNoTitles.Value, arrayIndex)
    End Select
    Set hostApp = Nothing
    
    RefreshDialog
     
    Status.Caption = "Report " & reportName & " done.": Status.Refresh
    
    Exit Sub
    
noExcel:
    MsgBox "Microsoft Excel is not installed on your system.", , "Sample5": Exit Sub
noWord:
    MsgBox "Microsoft Word is not installed on your system.", , "Sample5": Exit Sub
noInternetBrowser:
    MsgBox "There is no default Internet browser installed on your system.", , "Sample5": Exit Sub
noInternetExplorer:
    MsgBox "Microsoft Internet Explorer is not installed on your system.", , "Sample5": Exit Sub
End Sub

Private Sub OutputText(reportName As String, data() As String, nbFields As Long, startData As Long, nbData As Long)
    Dim columnSize() As Long
    ReDim columnSize(nbFields)
    
    For index = startData To nbData - 1
        Dim index2 As Long
        For index2 = 0 To nbFields - 1
            If columnSize(index2) <= Len(data(index2, index)) Then
                columnSize(index2) = Len(data(index2, index)) + 1
            End If
        Next index2
    Next index
        
    Dim fileHandle As Long: fileHandle = FreeFile
    Open reportName For Output As #fileHandle
    For index = startData To nbData - 1
        For index2 = 0 To nbFields - 1
            Print #fileHandle, data(index2, index); Space(columnSize(index2) - Len(data(index2, index)));
            If index2 = nbFields - 1 Then Print #fileHandle, ""
        Next index2
    Next index
    Close #fileHandle
    Dim retval As Long
    retval = ShellExecute(Me.hwnd, "open", "Notepad", reportName, "", 1)
End Sub

Private Sub OutputCSV(reportName As String, data() As String, nbFields As Long, startData As Long, nbData As Long)
    Dim fileHandle As Long: fileHandle = FreeFile
    Open reportName For Output As #fileHandle
    For index = startData To nbData - 1
        Dim index2 As Long
        For index2 = 0 To nbFields - 1
            Print #fileHandle, Chr$(34); data(index2, index); Chr$(34);
            If index2 <> nbFields - 1 Then Print #fileHandle, ","; Else Print #fileHandle, ""
        Next index2
    Next index
    Close #fileHandle
    Dim retval As Long
    retval = ShellExecute(Me.hwnd, "open", "Notepad", reportName, "", 1)
End Sub

Private Sub OutputExcel(excelApp As Object, reportName As String, data() As String, nbFields As Long, startData As Long, nbData As Long)
    ' Open temporarly text file
    Randomize
    Dim filename As String
    filename = ppcbApp.DefaultFilePath & "\tmp" & CInt(Rnd() * 10000) & ".csv"
    
    Dim fileHandle As Long: fileHandle = FreeFile
    Open filename For Output As #fileHandle

    ' Output data
    For index = startData To nbData - 1
        Dim index2 As Long
        For index2 = 0 To nbFields - 1
            Print #fileHandle, Chr$(34); data(index2, index); Chr$(34);
            If index2 <> nbFields - 1 Then Print #fileHandle, ","; Else Print #fileHandle, ""
        Next index2
    Next index

    ' Close the text file
    Close #fileHandle
    
    excelApp.Visible = True
    excelApp.Workbooks.OpenText filename:=filename
    excelApp.Rows("1:1").Select
    With excelApp.Selection
        .Font.Bold = True: .Font.Italic = True
    End With
    excelApp.Range("A1").Select
    excelApp.ActiveSheet.Name = ppcbDoc.Name & " Report"
    excelApp.Range("A2").Select
    excelApp.ActiveWindow.FreezePanes = True
    excelApp.ActiveWorkbook.SaveAs filename:=reportName, FileFormat:=xlNormal
    Kill filename
End Sub

Private Sub OutputExcelDynamic(excelApp As Object, reportName As String, data() As String, nbFields As Long, startData As Long, nbData As Long)
    Dim assoptName As String: assoptName = ""
    If CBXAssOpt.ListIndex > 0 Then assoptName = CBXAssOpt.List(CBXAssOpt.ListIndex)
    
    excelApp.Visible = True
    Dim thePath As String: thePath = App.Path & "\client.xls"
    excelApp.Workbooks.Open (thePath)
    excelApp.Run "Sheet1.Sample5", assoptName, data, nbFields, startData, nbData
    
    excelApp.ActiveWorkbook.SaveAs filename:=reportName, FileFormat:=xlNormal
End Sub

Private Sub OutputWord(wordApp As Object, reportName As String, data() As String, nbFields As Long, startData As Long, nbData As Long)
    ' Open temporarly text file
    Randomize
    Dim filename As String
    filename = ppcbApp.DefaultFilePath & "\tmp" & CInt(Rnd() * 10000) & ".csv"
    
    Dim fileHandle As Long: fileHandle = FreeFile
    Open filename For Output As #fileHandle

    ' Output data
    For index = startData To nbData - 1
        Dim index2 As Long
        For index2 = 0 To nbFields - 1
            Print #fileHandle, data(index2, index);
            If index2 <> nbFields - 1 Then Print #fileHandle, Chr$(9); Else Print #fileHandle, ""
        Next index2
    Next index

    ' Close the text file
    Close #fileHandle
    
    wordApp.Visible = True
    wordApp.Documents.Open filename:=filename
    wordApp.Selection.WholeStory
    wordApp.Selection.ConvertToTable Separator:=wdSeparateByTabs, NumColumns:=nbFields, NumRows:=nbData
    wordApp.ActiveDocument.SaveAs filename:=reportName, FileFormat:=wdFormatDocument
    ' Little trick here: Word still keeps our temp file opened, so we can't remove it.
    ' So, we close the Word doc and reopen it to force Word to release all file handles
    wordApp.ActiveWindow.Close
    wordApp.Documents.Open filename:=reportName
    Kill filename
End Sub


Private Sub OutputInternetBrowser(browserEXE As String, reportName As String, data() As String, nbFields As Long, startData As Long, nbData As Long)
    Dim fileHandle As Long: fileHandle = FreeFile
    Open reportName For Output As #fileHandle

    Print #fileHandle, "<HTML>"
    Print #fileHandle, "<HEAD><TITLE>" & reportName & "</TITLE></HEAD>"
    Print #fileHandle, "<BODY>"
    Print #fileHandle, "<TABLE BORDER CELLSPACING=1 CELLPADDING=9>"

    ' Output data
    For index = startData To nbData - 1
        Print #fileHandle, "<TR>"
        Dim index2 As Long
        For index2 = 0 To nbFields - 1
            If index = 0 Then
                Print #fileHandle, "<TD><B><P ALIGN='CENTER'>" & data(index2, index) & "</P></B></TD>"
            Else
                Print #fileHandle, "<TD><P>" & data(index2, index) & "</P></TD>"
            End If
        Next index2
        Print #fileHandle, "</TR>"
    Next index

    Print #fileHandle, "</TABLE>"
    Print #fileHandle, "</BODY>"
    Print #fileHandle, "</HTML>"

    ' Close the text file
    Close #fileHandle
    
    Dim retval As Long: Dim directory As String
    retval = ShellExecute(Me.hwnd, "open", browserEXE, reportName, directory, 1)
End Sub

Private Sub OutputIE4(IE4App As Object, reportName As String, data() As String, nbFields As Long, startData As Long, nbData As Long)
    Dim fileHandle As Long: fileHandle = FreeFile
    Open reportName For Output As #fileHandle

    Print #fileHandle, "<html>"
    Print #fileHandle, "<head><title>" & reportName & "</title></head>"
    Print #fileHandle, "<body bgColor='#C0C0C0'>"
    Print #fileHandle, ""
    Print #fileHandle, "<script language='JavaScript'>"
    Print #fileHandle, "function MakeBright() {"
    Print #fileHandle, "   window.event.srcElement.bgColor = '#DDDDDD';"
    Print #fileHandle, "}"
    Print #fileHandle, "function MakeNormal() {"
    Print #fileHandle, "   window.event.srcElement.bgColor = '#C0C0C0';"
    Print #fileHandle, "}"
    Print #fileHandle, "function MakeDown() {"
    Print #fileHandle, "   window.event.srcElement.bgColor = '#A2A2A2';"
    Print #fileHandle, "}"
    Print #fileHandle, "</script>"
    Print #fileHandle, ""
    
    Print #fileHandle, "<table border='1' bgColor='#C0C0C0' bordercolorlight='#000000' bordercolordark='#FFFFFF'>"

    ' Output data
    For index = startData To nbData - 1
        Print #fileHandle, "<tr>"
        Dim index2 As Long
        For index2 = 0 To nbFields - 1
            If index = 0 Then
                Print #fileHandle, "   <td><b><P ALIGN='CENTER'>" & data(index2, index) & "</P></b></td>"
            Else
                Print #fileHandle, "   <td onmouseover='MakeBright()' onmouseout='MakeNormal()' onmousedown='MakeDown()' onmouseup='MakeBright()'>" & data(index2, index) & "</td>"
            End If
        Next index2
        Print #fileHandle, "</tr>"
    Next index

    Print #fileHandle, ""
    Print #fileHandle, "</table>"
    Print #fileHandle, "</body>"
    Print #fileHandle, "</html>"

    ' Close the text file
    Close #fileHandle
        
    IE4App.Visible = True
    IE4App.Navigate (reportName)
End Sub
