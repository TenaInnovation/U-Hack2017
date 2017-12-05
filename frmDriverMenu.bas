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
	
	
	
	' --------- Views --------------------------------------
	Dim lvMenu As ListView
	Dim FakeActionBar, UnderActionBar As Panel
	Dim PanelWithSidebar As ClsSlidingSidebar
	Dim btnMenu As Button
	Dim lblmenu As Label
	Dim BarSize As Int: BarSize = 60dip
	Dim LightCyan As Int: LightCyan = Colors.Transparent
	Dim txtRFID As EditText
	' -------------------------------------------------------
	
End Sub

Sub Activity_Create(FirstTime As Boolean)
	'Do not forget to load the layout file created with the visual designer. For example:
	'Activity.LoadLayout("Layout1")
	InitializeViews
End Sub

Sub InitializeViews
	txtRFID.Initialize("txtRFID")
	Activity.AddView(txtRFID,10%x,BarSize,80%x,8%y)
	txtRFID.RequestFocus
	
	
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
	lvMenu.AddTwoLinesAndBitmap("QR","",LoadBitmap(File.DirAssets,"transaction.png"))
	lvMenu.SingleLineLayout.Label.TextColor = Colors.Black
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
	StartActivity(frmPayByQR)
End Sub
Sub txtRFID_TextChanged (Old As String, New As String)
	
End Sub

Sub txtRFID_EnterPressed
	
	Dim rfid As String = txtRFID.Text
		
	ToastMessageShow(rfid,True)
	
	txtRFID.RequestFocus
End Sub

Sub Activity_Resume

End Sub

Sub Activity_Pause (UserClosed As Boolean)

End Sub
