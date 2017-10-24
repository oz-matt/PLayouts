Attribute VB_Name = "Module1"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CLIENT.BAS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This is a part of the PADS Logic OLE Automation server Client sample.
' Copyright (C) 2007 Mentor Graphics Corp.
' All rights reserved.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This source code is only intended as a supplement to the PADS Logic
' Automation Server API Help file.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim ListComps$(10000)

Sub Main()
        Index = 0
        For Each nextComp In ActiveDocument.Components
                ListComps$(Index) = nextComp.Name
                Index = Index + 1
        Next nextComp

        Begin Dialog UserDialog 640,182,"Client (PADS Logic Basic Version)",.CallbackFunc ' %GRID:10,7,1,1
                Text 10, 7, 420, 14, "PADS Logic Document Name", .Text1
                TextBox 10, 21, 420, 21, .EdtDocname
                TextBox 10, 63, 420, 21, .EdtCompCount
                TextBox 10, 105, 420, 21, .EdtSelCompCount
                PushButton 220, 133, 100, 21, "Refresh", .BtnRefresh
                PushButton 330, 133, 100, 21, "Close", .BtnClose
                Text 10, 161, 410, 14, "Connected to PADS Logic.", .Status
                Text 10, 49, 420, 14, "PADS Logic Component Count", .Text2
                Text 10, 91, 420, 14, "PADS Logic Selected Component Count", .Text3
                ListBox 440, 21, 190, 140, ListComps(), .ListBox1
                Text 440, 7, 190, 14, "PADS Logic Doc Components", .Text4
        End Dialog
        Dim dlg As UserDialog
        Dialog dlg
End Sub

Rem See DialogFunc help topic for more information.
Private Function CallbackFunc(DlgItem$, Action%, SuppValue%) As Boolean
        Select Case Action%
        Case 1 ' Dialog box initialization
        Case 2 ' Value changing or button pressed
                CallbackFunc = True ' Prevent button press from closing the dialog box
                If DlgItem$ = "BtnClose" Then CallbackFunc = False
                If DlgItem$ = "BtnRefresh" Then
                        DlgText "EdtDocName", ActiveDocument.Name
                        DlgText "EdtCompCount", ActiveDocument.Components.Count & " components object(s)"
                        DlgText "EdtSelCompCount", ActiveDocument.GetObjects(plogObjectTypeComponent, "", True).Count & " selected component object(s)"
                End If
                If DlgItem$ = "ListBox1" Then
                                ActiveDocument.SelectObjects(9999, "", False)
                                ActiveDocument.SelectObjects(1, ListComps(SuppValue%), True)
                                LocateSelection
                End If
        Case 3 ' TextBox or ComboBox text changed
        Case 4 ' Focus changed
        Case 5 ' Idle
                Rem CallbackFunc = True ' Continue getting idle actions
        Case 6 ' Function key
        End Select
End Function

Public Sub Application_OpenDocument(Doc As PowerLogic.Document)
        CallbackFunc("BtnRefresh", 2, 0)
End Sub

Public Sub Document_SelectionChange()
        CallbackFunc("BtnRefresh", 2, 0)
End Sub

Sub LocateSelection()
        'look for selection in current sheet: if it has selection then do nothing
        If ActiveDocument.ActiveSheet.GetObjects(plogObjectTypeComponent, , True).Count > 0 Then Exit Sub
        For Each Sheet In ActiveDocument.Sheets
                If Sheet.GetObjects(plogObjectTypeComponent, , True).Count > 0 Then
                        Sheet.Activate
                        Exit Sub
                End If
        Next
End Sub
