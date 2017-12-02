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
	Dim selectedPUV As String 
End Sub

Sub Globals
	'These global variables will be redeclared each time the activity is created.
	'These variables can only be accessed from this module.

	Private btnSearch As Button
	Private lstLocations As ListView
	Private spnLocations As Spinner
	Dim driverName, driverLocation As String
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
		For i = 0 To Main.sqlResult.RowCount - 1
			Main.sqlResult.Position = i
			driverName = Main.sqlResult.GetString(2)
			driverLocation = Main.sqlResult.GetString(1)
			lstLocations.AddTwoLines(driverName,driverLocation)
		Next
	Else		
		Main.sqlResult = Main.sql.Query("Select * from EON_PUVLocation where location = '"& spnLocations.SelectedItem &"'")
		For i = 0 To Main.sqlResult.RowCount - 1
			Main.sqlResult.Position = i
			driverName = Main.sqlResult.GetString(2)
			driverLocation = Main.sqlResult.GetString(1)
			lstLocations.AddTwoLines(driverName,driverLocation)
		Next
	End If
	
End Sub

Sub InitializeView
	Activity.LoadLayout("frmFindPUV")
	
	spnLocations.Add("San Juan")
	spnLocations.Add("Rosario")
	spnLocations.Add("Garcia")
	spnLocations.Add("Lipa")
	spnLocations.Add("Malvar")
	spnLocations.Add("Tanauan")
	spnLocations.Add("Sto. Tomas")
	
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