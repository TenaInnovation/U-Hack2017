Type=Activity
Version=6
ModulesStructureVersion=1
B4A=true
@EndOfDesignText@
#Region  Activity Attributes 
	#FullScreen: False
	#IncludeTitle: False
#End Region

Sub Process_Globals
	'These global variables will be declared once when the application starts.
	'These variables can be accessed from all modules.
	Dim selectedPUV As String 
End Sub

Sub Globals
	'These global variables will be redeclared each time the activity is created.
	'These variables can only be accessed from this module.

	Private btnSearch As Button
	Private lstLocations As ListView
	Private spnLocations As Spinner
	Dim driverName, driverLocation As String
	Dim cd As ColorDrawable
End Sub

Sub Activity_Create(FirstTime As Boolean)
	'Do not forget to load the layout file created with the visual designer. For example:
	InitializeView
	ShowAvailableLocation
End Sub


Sub ShowAvailableLocation
	
	lstLocations.Clear()
	
	If spnLocations.SelectedIndex = -1 Then
		Main.sqlResult = Main.sql.Query("Select * from EON_PUVLocation")
		
		If Main.sqlResult.RowCount > 0 Then
		For i = 0 To Main.sqlResult.RowCount - 1
			Main.sqlResult.Position = i
			driverName = Main.sqlResult.GetString(2)
			driverLocation = Main.sqlResult.GetString(1)
			lstLocations.AddTwoLines(driverName,driverLocation)
		Next
		Else
			lstLocations.AddSingleLine("No available PUV for this route")
		End If
	Else		
		Main.sqlResult = Main.sql.Query("Select * from EON_PUVLocation where location = '"& spnLocations.SelectedItem &"'")
		If Main.sqlResult.RowCount > 0 Then
		For i = 0 To Main.sqlResult.RowCount - 1
			Main.sqlResult.Position = i
			driverName = Main.sqlResult.GetString(2)
			driverLocation = Main.sqlResult.GetString(1)
			lstLocations.AddTwoLines(driverName,driverLocation)
		Next
		Else
			lstLocations.AddSingleLine("No available PUV for this route")
		End If
	End If
	
End Sub

Sub InitializeView
	'Activity.LoadLayout("frmFindPUV")
	
	
	spnLocations.Initialize("spnLocations")
	spnLocations.Add("San Juan")
	spnLocations.Add("Rosario")
	spnLocations.Add("Garcia")
	spnLocations.Add("Lipa")
	spnLocations.Add("Malvar")
	spnLocations.Add("Tanauan")
	spnLocations.Add("Sto. Tomas")
	Activity.AddView(spnLocations, 10%x, 5%y, 80%x, 8%y)
	
	lstLocations.Initialize("lstLocations")
	Activity.AddView(lstLocations,0%x,13%y,100%x,87%y)
End Sub

Sub Activity_Resume

End Sub

Sub Activity_Pause (UserClosed As Boolean)

End Sub


Sub lstLocations_ItemClick (Position As Int, Value As Object)
	selectedPUV = Value
	StartActivity(frmPUVLocation)
End Sub

Sub btnSearch_Click
	ShowAvailableLocation
End Sub

Sub spnLocations_ItemClick (Position As Int, Value As Object)
	ShowAvailableLocation
End Sub