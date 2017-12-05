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
	Dim cmbCurrentLocation , cmbDestination As Spinner
	
	Dim messageQR As String
	Dim err As String
    Dim mask_pattern As Int
	Dim rsie As RSImageEffects
    Dim rsip As RSImageProcessing
	Dim b As Bitmap
	Dim shape As String
	Dim imgqr As ImageView
	
	
	Dim startDestination,endDestination As String 
End Sub

Sub Activity_Create(FirstTime As Boolean)
	'Do not forget to load the layout file created with the visual designer. For example:
	'Activity.LoadLayout("Layout1")
	
	initializeViews
	
End Sub

Sub initializeViews
	
	cmbCurrentLocation.Add("San Juan")
	cmbCurrentLocation.Add("Rosario")
	cmbCurrentLocation.Add("Garcia")
	cmbCurrentLocation.Add("Lipa")
	cmbCurrentLocation.Add("Malvar")
	cmbCurrentLocation.Add("Tanauan")
	cmbCurrentLocation.Add("Sto. Tomas")
	
	cmbDestination.Add("San Juan")
	cmbDestination.Add("Rosario")
	cmbDestination.Add("Garcia")
	cmbDestination.Add("Lipa")
	cmbDestination.Add("Malvar")
	cmbDestination.Add("Tanauan")
	cmbDestination.Add("Sto. Tomas")
	
	imgqr.Initialize ("")
	imgqr.Gravity = Gravity.FILL
	
	
	Activity.AddView(cmbCurrentLocation, 5%x, 5%y, 42%x, 8%y)
	Activity.AddView(cmbDestination, 52%x, 5%y, 40%x, 8%y)
	
	
	Activity.AddView (imgqr,5%x,15%y,90%x,70%y)
End Sub

Sub cmbCurrentLocation_ItemClick (Position As Int, Value As Object)
	startDestination = ((Position * cmbDestination.SelectedIndex) * -1) * 8
	generateQR(startDestination)
End Sub

Sub cmbDestination_ItemClick (Position As Int, Value As Object)
	endDestination = ((Position * cmbCurrentLocation.SelectedIndex) * -1) * 8
	generateQR(startDestination)
End Sub




Sub generateQR(code As String)
	
	messageQR = code
	shape = "s"
	mask_pattern = 3              
	err = "H"
	
	QRcode.Draw_QR_Code(messageQR,err,mask_pattern,Colors.White,Colors.Black,shape)    'Pass the message, the error level, the masking pattern, the background color, and the foreground/pixel color 

	b = LoadBitmap(File.DirRootExternal, "/pictures/QRcode.png")                       'Load the created png file into ImageView1
    b = rsip.createScaledBitmap(b, 445, 445, False)                                    'createScaledBitmap seems to produce a better compressed image
    imgqr.Bitmap = rsie.RoundCorner(b, 15)   
End Sub



Sub Activity_Resume

End Sub

Sub Activity_Pause (UserClosed As Boolean)

End Sub
