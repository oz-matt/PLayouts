Sub Main
	' This script wil generate the ASCII file for ROUTES of the currently loaded PowerPCB design
	' The name of the ASCII file is the following:
	outputFileName = DefaultFilePath & "\ROUTES.ASC"

	' Open temporarly text file
	Randomize
	filename = DefaultFilePath & "\tmp"  & CInt(Rnd()*10000) & ".mcr"
	Open filename For Output As #1

	' Output PowerPCB MACRO file
	Print #1,  "MACRO ROUTES"
	Print #1, "VERSION 2.0"
	Print #1,  "KEY"
	Print #1,  "DATE 05/29/98, 10:09"
	Print #1,  "DESCRIPTION "
	Print #1,  "BEGIN"
	Print #1,  "MODE MOUSEMOVE COMPRESS"
	Print #1,  "MODE PROMPT RECORD"
	Print #1,  "MODE COORDINATE ABSOLUTE"
	Print #1,  "SETTING DRC TRUEOFF"
	Print #1,  "SETTING UNIT MILL"
	Print #1,  "SETTING PICKRADIUS 5"
	Print #1,  "SETTING RESOLUTION 640x480 632x376"
	Print #1,  "COMMAND 34"
	Print #1,  "PARAMETER STRING 1 " & Chr$(34) & outputFileName & Chr$(34)
	Print #1,  "COMMAND 339"
	Print #1,  "PARAMETER INTEGER 1200 1024"
	Print #1,  "PARAMETER VOID 1201"
	Print #1,  "PARAMETER INTEGER 1206 6"
	Print #1,  "PARAMETER INTEGER 1203 0"
	Print #1,  "PARAMETER STRING 1207 " & Chr$(34) & Chr$(34)
	Print #1,  "PARAMETER STRING 1204 " & Chr$(34) & outputFileName & Chr$(34)
	Print #1,  "COMMAND 8"
	Print #1,  "END"
	Close #1
	RunMacro(filename, "ROUTES")
End Sub
