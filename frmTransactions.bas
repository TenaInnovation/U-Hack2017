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
	
	Private lstTransactions As ListView
	Dim transaction As String
	Dim transactionDate As String
End Sub

Sub Activity_Create(FirstTime As Boolean)
	'Do not forget to load the layout file created with the visual designer. For example:
	Activity.LoadLayout("frmTransactions")
	ShowTransaction
End Sub

Sub ShowTransaction
	Main.sqlResult = Main.sql.Query("Select * from EON_Account_Logs where username = '"& Global_declaration.username &"'")
	For i = 0 To Main.sqlResult.RowCount - 1
		Main.sqlResult.Position = 0
		transaction = Main.sqlResult.GetString(3)
		transactionDate = Main.sqlResult.GetString(4)
		lstTransactions.AddTwoLines(transaction,transactionDate)
	Next
End Sub

Sub Activity_Resume

End Sub

Sub Activity_Pause (UserClosed As Boolean)

End Sub


Sub lstTransactions_ItemClick (Position As Int, Value As Object)
	
End Sub