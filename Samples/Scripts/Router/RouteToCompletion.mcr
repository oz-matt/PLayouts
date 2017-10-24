' Sample: RouteToCompletion.MCR
' 
' This sample demonstrates usage of PADS Router Macro language.
'
' Route design multiple times increasing routing and optimizing intensity until 
' number of unroutes and via count cannot be improved
'
' For more details on macro language, please refer to the Macro Language Reference.
'

RouteToCompletion


Function RouteToCompletion

' Start with lowest intensity
intensity = 0

' While following variable is true - run router pass, otherwise run optimization only
route_pass = True

' Cleanup strategy parameters
Application.OpenOptionsDialog()
DlgOptions.ActiveTab = 5
DlgStrategyOptions.Pass.Cell("Fanout", "Pass")		= "0"
DlgStrategyOptions.Pass.Cell("Patterns", "Pass")		= "0"
DlgStrategyOptions.Pass.Cell("Route", "Pass")		= "0"
DlgStrategyOptions.Pass.Cell("Optimize", "Pass")		= "0"
DlgStrategyOptions.Pass.Cell("Miters", "Pass")		= "0"
DlgStrategyOptions.Pass.Cell("Test Point", "Pass")	= "0"
DlgStrategyOptions.Pass.Cell("Tune", "Pass")			= "0"
DlgOptions.Ok.Click()

' For intensity Low (0) to High (2)
For intensity = 0 To 2 Step 1

	'Get number of unroutes before routing pass
	old_unroute_count = Document.UnrouteCount

	'Route
	If route_pass = True Then
		Route(intensity)
	End If

	'Get number of unroutes after routing pass
	unroute_count = Document.UnrouteCount

	If unroute_count = 0 Then 
		'Design is routed 100 %
		route_pass = False
	End If

	'Get number of vias before optimization pass
	old_via_count = Document.Vias.Count

	Optimize(intensity)

	'Get number of vias after optimization pass
	via_count = Document.Vias.Count

	If (intensity = 2) And (via_count < old_via_count) Then 
		' Number of vias decreased during optimization pass with highest intensity
		' Try to route and optimize again. Decrease intensity, to stay in For loop.
		intensity = 1
	ElseIf (intensity = 2) And (unroute_count < old_unroute_count) Then 
		' Number of unroutes decreased during routing pass with highest intensity
		' Try to route and optimize again. Decrease intensity, to stay in For loop.
		intensity = 1
	End If

Next intensity

' Report results
via_count		= Document.Vias.Count
unroute_count	= Document.UnrouteCount
msg_str			= "Routing complete. Unroutes - " + Str(unroute_count) + ". Vias - " + Str(via_count)
MsgBox(msg_str)

End Function


Function Route(route_intensity)
	'Setup strategy parameters for routing
	Application.OpenOptionsDialog()
	DlgOptions.ActiveTab = 5
	DlgStrategyOptions.Pass.Cell("Optimize", "Pass") 	= "0"
	DlgStrategyOptions.Pass.Cell("Route", "Pass") 	= "1"
	DlgStrategyOptions.Pass.Cell("Route", "Intensity") = route_intensity
	DlgOptions.Ok.Click()
	'Route
	Document.Router.Run = True
	Delay
End Function

Function Optimize(optimize_intensity)
	'Setup strategy parameters for optimization
	Application.OpenOptionsDialog()
	DlgOptions.ActiveTab = 5
	DlgStrategyOptions.Pass.Cell("Route", "Pass") 		= "0"
	DlgStrategyOptions.Pass.Cell("Optimize", "Pass") 		= "1"
	DlgStrategyOptions.Pass.Cell("Optimize", "Intensity") 	= optimize_intensity
	DlgOptions.Ok.Click()
	'Optimize
	Document.Router.Run = True
	Delay
End Function



' Allow UI to process routing command
Function Delay
	for i = 0 to 100 step 1
		DoEvents
	next i
End Function

