' Sample 13: Route All.BAS
' 
' This sample demonstrates a real life usage of PADS Router Basic.
'
' This little script displays a dialog to select a folder.
' In the folder, all files with PCB extension will be routed and saved
' under a different name.
'
' For more details, please refer to the PADS Router Visual Basic Editor Help File.
'

Sub Main
	Dim fl$()  ' list of file names
	D$ = BrowseDirectory ' activate dialog to get folder
	D$ = D$ + "\"	' append backslash to separate folder name for the file name
	F$ = Dir$ (D$ + "*.pcb")	' set search for *.pcb files and get the first file
	' count files
	n% = 0
	While F$ <> ""	' if the *pcb file exists
		If Right(F$, 11) <> "_routed.pcb" Then ' file is not routed yet
			n% = n% + 1		
		End If
		F$ = Dir$	' get next *.pcb file in the folder
	Wend
	' create file list
	ReDim fl(n%)
	i% = 1
	F$ = Dir$ (D$ + "*.pcb")	' set search for *.pcb files and get the first file
	While F$ <> ""	' if the *pcb file exists
		If Right(F$, 11) <> "_routed.pcb" Then ' file is not routed yet
			If i% <= n% Then
				fl$(i%) = F$ ' store the file name in the list
			End If
			i% = i% + 1
		End If
		F$ = Dir$	' get next *.pcb file in the folder
	Wend
	' route files
	For i% = 1 To n%
		F$ = fl$(i%)
		OpenDocument D$ + F$	' open the file
		ActiveDocument.Router.Run = True		' start routing
		While ActiveDocument.Router.Run	' wait for the end of routing
			Wait 5
		Wend
		idx = InStrRev(F$, ".")		' find the end of file name
		F$ = Left(F$, idx - 1) + "_routed" + Mid(F$, idx)	' add _routed to the end of file
		ActiveDocument.SaveAs D$ + F$		' save routed file under the new name
	Next
End Sub
' following code uses system functions from Windows dll to inpot folder name
Type BROWSEINFO	' declare data structure for system function SHBrowseForFolder
	hwndOwner As Long 
	pidlRoot As Long
	pszDisplayName As Long 
	lpszTitle As Long 
	ulFlags As Long
	lpfn As Long
	lParam As Long
	iImage As Long
End Type
' declare used system functions
Private Declare Function SHBrowseForFolder Lib "shell32.dll" (ByRef bi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal dirName As String) As Long

' call system functions to get folder name
Function BrowseDirectory As String
	Dim bi As BROWSEINFO, s As String*256, pos
	bi.ulFlags = 9	' BIF_RETURNONLYFSDIRS Or BIF_RETURNFSANCESTORS
	bi.hwndOwner = 0
	Dim pidl As Long
	pidl = SHBrowseForFolder(bi)
	If pidl <> 0 Then
		If SHGetPathFromIDList(pidl, s) <> 0 Then
			pos = InStr(s, Chr(0))
			BrowseDirectory = Left(s, pos-1)
		End If
	End If
End Function
