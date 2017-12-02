Type=Activity
Version=6
ModulesStructureVersion=1
B4A=true
@EndOfDesignText@
#Region  Activity Attributes 
	#FullScreen: False
	#IncludeTitle: True
#End Region

Sub Process_Globals
	'These global variables will be declared once when the application starts.
	'These variables can be accessed from all modules.

End Sub

Sub Globals
	'These global variables will be redeclared each time the activity is created.
	'These variables can only be accessed from this module.
	Dim mapWebView As WebView
	Dim t1 As Timer
	Dim counter As Int
End Sub

Sub Activity_Create(FirstTime As Boolean)
	'Do not forget to load the layout file created with the visual designer. For example:
	'Activity.LoadLayout("Layout1")
	mapWebView.Initialize ("mapWebView")
	Activity.AddView (mapWebView,0%x,0%y,100%x,100%y)
	t1.Initialize ("t1",1000)
	t1.Enabled = True 
End Sub

Sub mapWebView_PageFinished (Url As String)
	ProgressDialogHide
End Sub

Sub t1_Tick
	counter = counter + 1
	If counter = 8 Then
		counter = 0
		
	Else If counter =1 Then
		Main.sqlResult = Main.sql.Query ("Select * from EON_PUVLocation where plateno = '"& frmFindPUV.selectedPUV  &"'")
		If Main.sqlResult.RowCount > 0 Then
'		ProgressDialogShow("Updating the bus' position")
'		'mapWebView.LoadUrl("file:///android_asset/location_map.htm?lat="& Main.sqlResult.GetString(3) & "&lng=" & Main.sqlResult.GetString(4) & "&zoom=10")		
'		mapWebView.LoadHtml("file:///android_asset/location_map.htm?lat=13.1213&lng=121.11312&zoom=10")
		ShowMap(Main.sqlResult.GetString(3), Main.sqlResult.GetString(4),10)
		End If
		Main.sqlResult.Close
	End If	
End Sub


Sub ShowMap(Latitude As Float, Longitude As Float, Zoom As Int)
 Dim HtmlCode As String
 HtmlCode = "<!DOCTYPE html><html><head><meta name='viewport' content='initial-scale=1.0, user-scalable=no' /><style type='text/css'> html { height: 100% } body { height: 100%; margin: 0px; padding: 0px }#map_canvas { height: 100% }</style><script type='text/javascript' src='http://maps.google.com/maps/api/js?sensor=true'></script><script type='text/javascript'> function initialize() { var latlng = new google.maps.LatLng(" & Latitude & "," & Longitude & "); var myOptions = { zoom: "&Zoom&", center: latlng, mapTypeId: google.maps.MapTypeId.ROADMAP }; var map = new google.maps.Map(document.getElementById('map_canvas'), myOptions); var marker0 = new google.maps.Marker({ position: new google.maps.LatLng(" & Latitude & "," & Longitude & "),map: map,title: '',clickable: false,icon: '' }); }</script></head><body onload='initialize()'> <div id='map_canvas' style='width:100%; height:100%'></div></body></html>"
 mapWebView.LoadHtml(HtmlCode)
End Sub

Sub Activity_Resume

End Sub

Sub Activity_Pause (UserClosed As Boolean)

End Sub
