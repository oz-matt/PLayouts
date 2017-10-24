' Sample: RouteAll.MCR
' 
' This sample demonstrates a real life usage of PADS Router Macro.
'
' Set the folder name in the first line of this script
' In the folder, all files with PCB extension will be routed and saved
' under a different name.
'
' For more details, please refer to the Help File.
'

D = "c:\padspwr\files\temp"  ' set he folder name here as needed

If D <> "" Then	' if folder name is not empty
	D = D + "\"		' append backslash separator to the folder name
	ext  = ".pcb"	' set file extention to be "pcb"
	F = Dir (D + "*" + ext)		' set the filter and get the first file name
	Dim FileList(1 to 1000)		' declare file names list
	idx = 1									' set stating index
	routed = "_routed"					' set string to applend to file name
	While F <> ""
		if InStr(F, routed) = 0 then		' if the file is not routed
			FileList(idx) = D & F			' store path and file name
			idx = idx + 1						' increment the index
		end if
		F = Dir									' get next file
	Wend
	count = idx - 1						' the number of files to route
	For idx = 1 to count					' for each file
		Application.OpenDocument(FileList(idx))				' open the file
		Application.ExecuteCommand("CLI_BATCH_ROUTE")		' start router
		L = Len( FileList(idx) )									' find lenght of the file name
		name = Left( FileList(idx), L - Len(ext) )			' define new name
		Document.SaveAs(name & routed & ext)					' save routed design under a new name
	Next idx									' continue with next file in the list
End If
