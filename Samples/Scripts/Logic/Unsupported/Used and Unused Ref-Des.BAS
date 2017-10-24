'Unused and Last Used Reference Designators Report"
'Compatible with PowerPCB v3.0 and PowerLogic v3.0
'
'
Sub Main()

        Dim RefDesCount As Integer

        Dim res As Boolean
        Dim s, smain As String
        Dim smain_prev As String
        Dim snum As String
        Dim snum_prev As String
        
        'Get component collection of design
        RefDesCount = 0
        Dim flag_exist As Boolean
        Dim Comps As objects
        Dim RefDes As String
        
        'Report about last used and unused Ref Des to Wordpad
        filename = DefaultFilePath & "\Last Used and Unused RefDes.rtf"
        Open filename For Output As #1
        
        Dim tab As Integer
        tab=4
        smain_prev = ""
        snum_prev = ""
        Set Comps = Document.Components
        For Each comp In Comps
                RefDes = comp.Name
                SplitRefDes(RefDes, smain, snum)
                If Not smain = smain_prev Then
                        If Not smain_prev = "" Then Print #1, Space(tab); "Last used Ref Des is: " & snum_prev
                        snum_prev = snum
                    Print #1, "For component " & smain & ":"
                    Print #1, Space(tab); "First used Ref Des is: " & snum
                Else
                        If Not IncreaseString(snum_prev) = snum Then
                        Print #1, Space(tab); "There are unused Ref Des between " & snum_prev & " and " & snum
                        End If
                    snum_prev = snum
                End If
                smain_prev = smain
    
        Next comp
        
        Print #1, Space(tab); "Last used Ref Des is: " & snum_prev
        
        Close #1
        Shell "Write " & """" & filename & """", 1
        
        Exit Sub
p0:
        MsgBox "Unknown error"

End Sub

'Following function get string (Ref Des) and split it at main part and number part
'function returns False if rd is empty string
Function SplitRefDes(rd As String, ByRef smain As String, ByRef snum As String) As Boolean
        Dim length, num, n As Integer
        Dim sn, tmp As String
        
        If Len(rd) = 0 Then
                SplitRefDes = False
                Exit Function
        End If
        
        On Error GoTo SplitString
        n = CInt(rd)
        snum = rd
        smain = """Empty"""
        SplitRefDes = True
        Exit Function

SplitString:
        snum = ""
        length = Len(rd)
        For k = length To k = 1 Step -1
                sn = Mid(rd, k, 1)
                On Error GoTo NotInt
                n = CInt(sn)
                On Error GoTo f0
                snum = sn + snum
        Next k
NotInt:
        If snum = "" Then snum = """Empty"""
        smain = Mid(rd, 1, length - (length - k))
        If smain = "" Then smain = """Empty"""
        SplitRefDes = True
        Exit Function
        
f0:
        MsgBox "Unknown error"
End Function

Function IncreaseString(sn As String) As String
        Dim l As Integer
        
        If sn = """Empty""" Then
                IncreaseString = "1"
                Exit Function
        End If
        
        IncreaseString = CStr(CInt(sn) + 1)
        
End Function

Sub Clear(ByRef temp_objs As objects)
        For k = 1 To temp_objs.Count
                temp_objs.Remove (1)
        Next k
End Sub

