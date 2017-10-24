 'Count smt pads per side.BAS
' 
' This will check each pin on the pcb to see if it is a smd pin.  Then it will check which layer it is on.
'It outputs: total smd on top, total smd on bottom, total smd, total thru-hole and total pins
'mb 12-23-98

Dim Listpins$(10000)
Sub Main
	counttop=0
	countbot=0
	countsmd=0
	countthr=0
	index = 0
	Dim nextpin As Object
	For Each nextpin In ActiveDocument.GetObjects(3, "", False)
		If nextpin.IsSMD Then countsmd=countsmd+1 Else countthr=countthr+1
		If nextpin.Component.layer =1 And nextpin.IsSMD Then counttop=counttop+1
		If nextpin.Component.layer >1 And nextpin.IsSMD Then countbot=countbot+1
		Listpins$(index) = nextpin.Name +" " +Str(nextpin.IsSMD) + Str(nextpin.Component.layer)+" " +Str(countsmd) + " " +Str(counttop)
		index = index + 1
	Next nextpin
	
	Begin Dialog UserDialog 310,189
		OKButton 110,161,90,21
		Text 40,14,270,14,"Total SMD pins on top:  " +counttop
		Text 40,42,220,14,"Total SMD pins on bottom:  " +countbot
		Text 40,70,220,14,"Total SMD pins:  " + countsmd
		Text 40,98,220,14,"Total thru-hole pins:  " +countthr
		Text 40,126,190,14,"Total pins:  " + index
		
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg
	
End Sub
