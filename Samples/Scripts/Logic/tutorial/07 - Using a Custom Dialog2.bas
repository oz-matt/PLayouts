' Sample 07: Using a Custom Dialog2.BAS
' 
' This sample demonstrates using the PADS Logic Basic Dialog Editor.
'
' The PADS Logic Basic Editor contains a dialog editor, accessed through
' the second button (starting from the end) of the toolbar of the Basic
' Editor dialog. The PADS Logic Basic editor is a graphic editor which
' automatically generates the code for you.
' The code below was not written by hand but automatically generated by the
' PADS Logic Basic Dialog Editor.
'
' For more details, please refer to the PADS Logic Basic Editor Help File.
'
Sub Main
	Begin Dialog UserDialog 220,182
		CheckBox 10,119,90,14,"Check Box",.CheckBox1
		GroupBox 10,49,90,63,"Options",.GroupBox1
		OptionGroup .Group1
			OptionButton 20,63,70,14,"Option 1",.OptionButton1
			OptionButton 20,77,70,14,"Option 2",.OptionButton2
			OptionButton 20,91,70,14,"Option 3",.OptionButton3
		TextBox 10,21,200,21,.TextBox1
		OKButton 60,147,90,21
		Text 10,7,200,14,"This is a text field",.Text1
		PushButton 120,56,90,21,"Button 1",.PushButton1
		PushButton 120,84,90,21,"Button 2",.PushButton2
		PushButton 120,112,90,21,"Button 3",.PushButton3
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg
	
End Sub
