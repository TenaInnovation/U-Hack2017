Type=Activity
Version=6
ModulesStructureVersion=1
B4A=true
@EndOfDesignText@
#Region  Activity Attributes 
	#FullScreen: True	
	#IncludeTitle: False
#End Region

Sub Process_Globals
	'These global variables will be declared once when the application starts.
	'These variables can be accessed from all modules.

End Sub

Sub Globals
	'These global variables will be redeclared each time the activity is created.
	'These variables can only be accessed from this module.
	
	' --------- Views --------------------------------------
	Dim lvMenu As ListView
	Dim FakeActionBar, UnderActionBar As Panel
	Dim PanelWithSidebar As ClsSlidingSidebar
	Dim btnMenu As Button
	Dim lblmenu As Label
	Dim BarSize As Int: BarSize = 60dip
	Dim LightCyan As Int: LightCyan = Colors.Transparent
	' -------------------------------------------------------
	
	' -------- QR -------------------
	Dim QRResult As String
	Dim scanner As ABZxing
	'--------------------------------
	
	'----- Others ------------------
	Dim currentBalance As Double
	'-------------------------------
End Sub

Sub Activity_Create(FirstTime As Boolean)
	'Do not forget to load the layout file created with the visual designer. For example:
	'Activity.LoadLayout("Layout1")
	Activity.SetBackgroundImage(LoadBitmap(File.DirAssets,"Generic_Background.jpg"))
	InitializeView
End Sub

Sub InitializeView
	
	lblmenu.Initialize("")
	lblmenu.Text = "MENU"
	lblmenu.Gravity = Gravity.CENTER
	
	
	FakeActionBar.Initialize("")
	FakeActionBar.Color = Colors.Transparent
	Activity.AddView(FakeActionBar, 0, 0, 100%x, BarSize)

	
	UnderActionBar.Initialize("")
	UnderActionBar.Color = LightCyan
	Activity.AddView(UnderActionBar, 0, BarSize, 100%x, 100%y - BarSize)

	PanelWithSidebar.Initialize(UnderActionBar, 65%y, 2, 1, 500, 500)
	PanelWithSidebar.ContentPanel.Color = LightCyan
	PanelWithSidebar.Sidebar.Background = PanelWithSidebar.LoadDrawable("popup_inline_error")
	PanelWithSidebar.SetOnChangeListeners(Me, "Menu_onFullyOpen", "Menu_onFullyClosed", "Menu_onMove")
	
	
	lvMenu.Initialize("lvMenu")
	lvMenu.AddTwoLinesAndBitmap("Find PUV","",LoadBitmap(File.DirAssets,"find_puv_icon.png"))
	lvMenu.AddTwoLinesAndBitmap("Pay thru QR","",LoadBitmap(File.DirAssets,"qr_icon.png"))
	lvMenu.AddTwoLinesAndBitmap("Remaining Balance","",LoadBitmap(File.DirAssets,"balance_icon.png"))
	lvMenu.AddTwoLinesAndBitmap("Transaction","",LoadBitmap(File.DirAssets,"trans_log_icon.png"))
	'lvMenu.SingleLineLayout.Label.TextColor = Colors.Black
	lvMenu.Color = Colors.Transparent
	lvMenu.ScrollingBackgroundColor = Colors.Transparent
	lvMenu.TwoLinesAndBitmap.Label.TextColor = Colors.Black
	PanelWithSidebar.Sidebar.AddView(lvMenu, 20dip, 25dip, -1, -1)
	
	btnMenu.Initialize("")
	btnMenu.SetBackgroundImage(LoadBitmap(File.DirAssets, "menu.png"))
	FakeActionBar.AddView(btnMenu, 100%x - BarSize, 0, BarSize, BarSize)
	FakeActionBar.AddView(lblmenu,73%x,0%y, 20%x,BarSize)
	PanelWithSidebar.SetOpenCloseButton(btnMenu)	
End Sub

Sub lvMenu_ItemClick (Position As Int, Value As Object)
	If Position = 0 Then
		' check PUV location
		StartActivity(frmFindPUV)
	else if Position = 1 Then
		' initialize QR Scanner
		scanner.ABGetBarcode("scanner","QR_CODES_TYPES")
	Else If Position = 2 Then
		' check remaining balance
		Main.sqlResult = Main.sql.Query("Select * from EON_Account where username = '"& Global_declaration.username &"' ")
		If Main.sqlResult.RowCount > 0 Then
			Main.sqlResult.Position = 0
			currentBalance = Main.sqlResult.GetInt(2)
			Msgbox("Your current load balance is " & currentBalance, "")
		End If
	else if Position = 3 Then
		' check transactions
		StartActivity(frmTransactions)
	End If
End Sub

Sub scanner_BarcodeFound (barCode As String, formatName As String)
	
	QRResult = barCode
	
	
	Main.sql.Exec("Update eon_account set balance = "& QRResult &"")
	Msgbox("Your have successfully paid your fare","")
	
End Sub

Sub Activity_Resume

End Sub

Sub Activity_Pause (UserClosed As Boolean)

End Sub
