Type=StaticCode
Version=6
ModulesStructureVersion=1
B4A=true
@EndOfDesignText@
'Code module
'Subs in this code module will be accessible from all modules.
 Private Sub Process_Globals
	'These global variables will be declared once when the application starts.
	'These variables can be accessed from all modules.
	Private Rect1 As Rect    

	Private sf As StringFunctions
	Private gf As Int = 256  'define the Size & Prime Polynomial of this Galois field...
	Private PP As Int = 285  '...for QR codes
	Private logg(256) As Int
    Private alogg(512) As Int
	Private gf_mul_result As Int
	Private s4c4(37) As Int
	Private s4c24(35)
	Private s4c1(301)
	Private z(4000)
	Private message_coeff(2956) As Int   'might have to change to accommodate versions >15 with max data
	Private s5cij(180,180) As Int        'dimensioned for up to version 40 QR codes (177 x 177)
	Private version_no, max_mod, qr_size As Int
    Private error_level As String
	Private a As String
	Private b As String
	Private q As String
	Private tel2 As Int
	Private anew As String
	Private ECC_woord As String
	Private no_of_error_words,mess_coef_teller As Int
	Private block1,block2,dw1,dw2,ew1,ew2 As Int
	Private max_data_bits As Int
	Private maxlen As Int
	Private tipe As String
	Private maskp As Int
    Private dat As String
	Private ver1(), ver2(), ver3(), ver4(), ver5(), ver6(), ver7(), ver8(), ver9(), ver10() As Int	
	Private ver11(), ver12(), ver13(), ver14(), ver15(), ver16(), ver17(), ver18(), ver19(), ver20()As Int
	Private ver21(), ver22(), ver23(), ver24(), ver25(), ver26(), ver27(), ver28(), ver29(), ver30() As Int	
	Private ver31(), ver32(), ver33(), ver34(), ver35(), ver36(), ver37(), ver38(), ver39(), ver40()As Int

	 ver1 = Array As Int (1,21,441,0,30,1,0,0,0,10,192,233,208,208,0,26,7,10,13,17,152,128,104,72,1,26,19,7,0,0,0,0,1,26,16,10,0,0,0,0,1,26,13,13,0,0,0,0,1,26,9,17,0,0,0,0,0,0,0,0,0,0,0,0,17,14,11,7,40,33,26,16,24,19,15,9,19,16,13,9)
	 ver2 = Array As Int (2,25,625,0,30,1,1,0,25,18,192,266,359,352,7,44,10,16,22,28,272,224,176,128,1,44,34,10,0,0,0,0,1,44,28,16,0,0,0,0,1,44,22,22,0,0,0,0,1,44,16,28,0,0,0,0,1,6,18,0,0,0,0,0,32,26,20,14,76,62,47,33,46,37,28,19,34,28,22,16)
	 ver3 = Array As Int (3,29,841,0,30,1,1,0,25,26,192,274,567,560,7,70,15,26,36,44,440,352,272,208,1,70,55,15,0,0,0,0,1,70,44,26,0,0,0,0,2,35,17,18,0,0,0,0,2,35,13,22,0,0,0,0,1,6,22,0,0,0,0,0,53,42,32,24,126,100,76,57,76,60,46,34,55,44,17,13)
	 ver4 = Array As Int (4,33,1089,0,30,1,1,0,25,34,192,282,807,800,7,100,20,36,52,64,640,512,384,288,1,100,80,20,0,0,0,0,2,50,32,18,0,0,0,0,2,50,24,26,0,0,0,0,4,25,9,16,0,0,0,0,1,6,26,0,0,0,0,0,78,62,46,34,113,89,66,49,113,89,66,49,80,32,24,9)
	 ver5 = Array As Int (5,37,1369,0,30,1,1,0,25,42,192,290,1079,1072,7,134,26,48,72,88,864,688,496,368,1,134,108,26,0,0,0,0,2,67,43,24,0,0,0,0,2,33,15,18,2,34,16,18,2,33,11,22,2,34,12,22,1,6,30,0,0,0,0,0,106,84,60,44,153,121,86,63,153,121,86,63,108,43,16,12)
	 ver6 = Array As Int (6,41,1681,0,30,1,1,0,25,50,192,298,1383,1376,7,172,36,64,96,112,1088,864,608,480,2,86,68,18,0,0,0,0,4,43,27,16,0,0,0,0,4,43,19,24,0,0,0,0,4,43,15,28,0,0,0,0,1,6,34,0,0,0,0,0,134,106,74,58,194,153,107,83,194,153,107,83,68,27,19,15)
	 ver7 = Array As Int (7,45,2025,36,30,1,6,1,150,48,192,457,1568,1568,0,196,40,72,108,130,1248,992,704,528,2,98,78,20,0,0,0,0,4,49,31,18,0,0,0,0,2,32,14,18,4,33,15,18,4,39,13,26,1,40,14,26,6,6,22,38,0,0,0,0,154,122,86,64,223,177,124,92,223,177,124,92,78,31,15,14)
	 ver8 = Array As Int (8,49,2401,36,30,1,6,1,150,56,192,465,1936,1936,0,242,48,88,132,156,1552,1232,880,688,2,121,97,24,0,0,0,0,2,60,38,22,2,61,39,22,4,40,18,22,2,41,19,22,4,40,14,26,2,41,15,26,6,6,24,42,0,0,0,0,192,152,108,84,278,220,156,121,278,220,156,121,97,39,19,15)
	 ver9 = Array As Int (9,53,2809,36,30,1,6,1,150,64,192,473,2336,2336,0,292,60,110,160,192,1856,1456,1056,800,2,146,116,30,0,0,0,0,3,58,36,22,2,59,37,22,4,36,16,20,4,37,17,20,4,36,12,24,4,37,13,24,6,6,26,46,0,0,0,0,230,180,130,98,551,431,311,234,334,261,188,142,116,37,17,13)
	 ver10 = Array As Int (10,57,3249,36,30,1,6,1,150,72,192,481,2768,2768,0,346,72,130,192,224,2192,1728,1232,976,2,86,68,18,2,87,69,18,4,69,43,26,1,70,44,26,6,43,19,24,2,44,20,24,6,43,15,28,2,44,16,28,6,6,28,50,0,0,0,0,271,213,151,119,651,512,363,287,394,310,220,173,69,44,20,16)
	 ver11 = Array As Int (11,61,3721,36,30,1,6,1,150,80,192,489,3232,3232,0,404,80,150,224,264,2592,2032,1440,1120,4,101,81,20,0,0,0,0,1,80,50,30,4,81,51,30,4,50,22,28,4,51,23,28,3,36,12,24,8,37,13,24,6,6,30,54,0,0,0,0,321,251,177,137,771,603,426,330,467,365,258,199,81,51,23,13)
	 ver12 = Array As Int (12,65,4225,36,30,1,6,1,150,88,192,497,3728,3728,0,466,96,176,260,308,2960,2320,1648,1264,2,116,92,24,2,117,93,24,6,58,36,22,2,59,37,22,4,46,20,26,6,47,21,26,7,42,14,28,4,43,15,28,6,6,32,58,0,0,0,0,367,287,203,155,882,690,488,373,534,418,295,226,93,37,21,15)
	 ver13 = Array As Int (13,69,4761,36,30,1,6,1,150,96,192,505,4256,4256,0,532,104,198,288,352,3424,2672,1952,1440,4,133,107,26,0,0,0,0,8,59,37,22,1,60,38,22,8,44,20,24,4,45,21,24,12,33,11,22,4,34,12,22,6,6,34,62,0,0,0,0,425,331,241,177,1021,795,579,426,618,482,351,258,107,38,21,12)
	 ver14 = Array As Int (14,73,5329,36,30,1,13,2,325,94,192,678,4651,4648,3,581,120,216,320,384,3688,2920,2088,1576,3,145,115,30,1,146,116,30,4,64,40,24,5,65,41,24,11,36,16,20,5,37,17,20,11,36,12,24,5,37,13,24,13,6,26,46,66,0,0,0,458,362,258,194,1100,870,620,467,666,527,375,282,116,41,17,13)
	 ver15 = Array As Int (15,77,5929,36,30,1,13,2,325,102,192,686,5243,5240,3,655,132,240,36,432,4184,3320,4952,1784,5,109,87,22,1,110,88,22,5,65,41,24,5,66,42,24,5,54,24,30,7,55,25,30,11,36,12,24,7,37,13,24,13,6,26,48,70,0,0,0,520,412,292,220,1249,990,702,529,757,599,425,320,88,42,25,13)
	 ver16 = Array As Int (16,81,6561,36,30,1,13,2,325,110,192,694,5867,5864,3,733,144,280,408,480,4712,3624,2600,2024,5,122,98,24,1,123,99,24,7,73,45,28,3,74,46,28,15,43,19,24,2,44,20,24,3,45,15,30,13,46,16,30,13,6,26,50,74,0,0,0,586,450,322,250,1407,1081,774,601,853,655,469,364,99,46,20,16)
	 ver17 = Array As Int (17,85,7225,36,30,1,13,2,325,118,192,702,6523,6520,3,815,168,308,448,532,5176,4056,2936,2264,1,135,107,28,5,136,108,28,10,74,46,28,1,75,47,14,1,50,22,28,15,51,23,28,2,42,14,28,17,43,15,28,13,6,30,54,78,0,0,0,644,504,364,280,1547,1211,875,673,937,733,530,407,108,47,23,15)
	 ver18 = Array As Int (18,89,7921,36,30,1,13,2,325,126,192,710,7211,7208,3,901,180,338,504,588,5768,4504,3176,2504,5,150,120,30,1,151,121,30,9,69,43,26,4,70,44,26,17,50,22,28,1,51,23,28,2,42,14,28,19,43,15,28,13,6,30,56,82,0,0,0,718,560,394,310,1724,1345,947,745,1045,815,573,451,121,44,23,15)
	 ver19 = Array As Int (19,93,8649,36,30,1,13,2,325,134,192,718,7931,7928,3,991,196,364,546,650,6360,5016,3560,2728,3,141,113,28,4,142,114,28,3,70,44,26,11,71,45,26,17,47,21,26,4,48,22,26,9,39,13,26,16,40,14,26,13,6,30,58,86,0,0,0,792,624,442,338,192,1499,1062,812,1152,908,643,492,114,45,22,14)
	 ver20 = Array As Int (20,97,9409,36,30,1,13,2,325,142,192,726,8683,8680,3,1085,224,416,600,700,6888,5352,3880,3080,3,135,107,28,5,136,108,28,3,67,41,26,13,68,42,26,15,54,24,30,5,55,25,30,15,43,15,28,10,44,16,28,13,6,34,62,90,0,0,0,858,666,482,382,2060,1599,1158,918,1248,969,701,556,108,42,25,16)
	 ver21 = Array As Int (21,101,10201,36,30,1,22,3,550,140,192,949,9252,9248,4,1156,224,442,644,750,7456,5712,4096,3248,4,144,116,28,4,145,117,28,17,68,42,26,0,0,0,0,17,50,22,28,6,51,23,28,19,46,16,30,6,47,17,30,22,6,28,50,72,94,0,0,929,711,509,403,2231,1707,1223,968,1351,1034,741,586,117,42,23,17)
	 ver22 = Array As Int (22,105,11025,36,30,1,22,3,550,148,192,957,10068,10064,4,1258,252,476,690,816,8048,6256,4544,3536,2,139,111,28,7,140,112,28,17,74,46,28,0,0,0,0,7,54,24,30,16,55,25,30,34,37,13,24,0,0,0,0,22,6,26,50,74,98,0,0,1003,779,565,439,2408,1871,1357,1055,1459,1133,822,639,112,46,25,13)
	 ver23 = Array As Int (23,109,11881,36,30,1,22,3,550,156,192,965,10916,10912,4,1364,270,504,750,900,8752,6880,4912,3712,4,151,121,30,5,152,122,30,4,75,47,28,14,76,48,28,11,54,24,30,14,55,25,30,16,45,15,30,14,46,16,30,22,6,30,54,78,102,0,0,1091,857,611,461,2619,2058,1467,1107,1587,1247,88,671,122,48,25,16)
	 ver24 = Array As Int (24,113,12769,36,30,1,22,3,550,164,192,973,11796,11792,4,1474,300,560,810,960,9392,7312,5312,4112,6,147,117,30,4,148,118,30,6,73,45,28,14,74,46,28,11,54,24,30,16,55,25,30,30,46,16,30,2,47,17,30,22,6,28,54,80,106,0,0,1171,911,661,511,2811,2187,1587,1227,1703,1325,962,743,118,46,25,17)
	 ver25 = Array As Int (25,117,13689,36,30,1,22,3,550,172,192,981,12708,12704,4,1588,312,588,870,1050,10208,8000,5744,4304,8,132,106,26,4,133,107,26,8,75,47,28,13,76,48,28,7,54,24,30,22,55,25,30,22,45,15,30,13,46,16,30,22,6,32,58,84,110,0,0,1273,997,715,535,3056,2394,1717,1285,1852,1450,1040,778,107,48,25,16)
	 ver26 = Array As Int (26,121,14641,36,30,1,22,3,550,180,192,989,13652,13648,4,1706,336,644,952,1110,10960,8496,6032,4768,10,142,114,28,2,143,115,28,19,74,46,28,4,75,47,28,28,50,22,28,6,51,23,28,33,46,16,30,4,47,17,30,22,6,30,58,86,114,0,0,1367,1059,751,593,3282,2543,1803,1424,1989,1541,1093,863,115,47,23,17)
	 ver27 = Array As Int (27,125,15625,36,30,1,22,3,550,188,192,997,14628,14624,4,1828,360,700,1020,1200,11744,9024,6464,5024,8,152,122,30,4,153,123,30,22,73,45,28,3,74,46,28,8,53,23,30,26,54,24,30,12,45,15,30,28,26,16,30,22,6,34,62,90,118,0,0,1465,1125,805,625,3516,2700,1932,1500,2131,1636,1171,909,123,46,24,16)
	 ver28 = Array As Int (28,129,16641,36,30,1,33,4,825,186,192,1270,15371,15368,3,1921,390,728,1050,1260,12248,9544,6968,5288,3,147,117,30,10,148,118,30,3,73,45,28,23,74,46,28,4,54,24,30,31,55,25,30,11,45,15,30,31,46,16,30,33,6,26,50,74,98,122,0,1528,1190,868,658,3668,2856,2084,1580,2222,1731,1262,957,118,46,25,16)
	 ver29 = Array As Int (29,133,17689,36,30,1,33,4,825,194,192,1278,16411,16408,3,2051,420,784,1140,1350,13048,10136,7288,5608,7,146,116,30,7,147,117,30,21,73,45,28,7,74,46,28,1,53,23,30,37,54,24,30,19,45,15,30,26,46,16,30,33,6,30,54,78,102,126,0,1628,1264,908,698,398,3034,2180,1676,2368,1838,1321,1015,117,46,24,16)
	 ver30 = Array As Int (30,137,18769,36,30,1,33,4,825,202,192,1286,17483,17480,3,2185,450,812,1200,1440,13880,10984,7880,5960,5,145,115,30,10,146,116,30,19,75,47,28,10,76,48,28,15,54,24,30,25,55,25,30,23,45,15,30,25,46,16,30,33,6,26,52,78,104,130,0,1732,1370,982,742,4157,3288,2357,1781,2519,1993,1428,1079,116,48,25,16)
	 ver31 = Array As Int (31,141,19881,36,30,1,33,4,825,210,192,1294,18587,18584,3,2323,480,868,1290,1530,14744,11640,8264,6344,13,145,115,30,3,146,116,30,2,74,46,28,29,75,47,28,42,54,24,30,1,55,25,30,23,45,15,30,28,46,16,30,33,6,30,56,82,108,134,0,1840,1452,1030,790,4416,3485,2472,1896,2676,2112,1498,1149,116,47,25,16)
	 ver32 = Array As Int (32,145,21025,36,30,1,33,4,825,218,192,1302,19723,19720,3,2465,510,924,1350,1620,15640,12328,8920,6760,17,145,115,30,0,0,0,0,10,74,46,28,23,75,47,28,10,54,24,30,35,55,25,30,19,45,15,30,35,46,16,30,33,6,34,60,86,112,138,0,1952,1538,1112,842,4685,3692,2669,2021,2839,2237,1617,1225,115,47,25,16)
	 ver33 = Array As Int (33,149,22201,36,30,1,33,4,825,226,192,1310,20891,20888,3,2611,540,980,1440,1710,16568,13048,9368,7208,17,145,115,30,1,146,116,30,14,74,46,28,21,75,47,28,29,54,24,30,19,55,25,30,11,45,15,30,46,46,16,30,33,6,30,58,86,114,142,0,2068,1628,1168,898,4964,3908,2804,2156,3008,2368,1699,1306,116,47,25,16)
	 ver34 = Array As Int (34,153,23409,36,30,1,33,4,825,234,192,1318,22091,22088,3,2761,570,1036,1530,1800,17528,13800,9848,7688,13,145,115,30,6,146,116,30,14,74,46,28,23,75,47,28,44,54,24,30,7,55,25,30,59,46,16,30,1,47,17,30,33,6,34,62,90,118,146,0,2188,1722,1228,958,5252,4133,2948,2300,3182,2505,1786,1393,116,47,25,17)
	 ver35 = Array As Int (35,157,24649,36,30,1,46,5,1150,232,192,1641,23008,23008,0,2876,570,1064,1590,1890,18448,14496,10288,7888,12,151,121,30,7,152,122,30,12,75,47,28,26,76,48,28,39,54,24,30,14,55,25,30,22,45,15,30,41,46,16,30,46,6,30,54,78,102,126,150,2303,1809,1283,983,5528,4342,3080,2360,3350,2631,1866,1430,122,48,25,16)
	 ver36 = Array As Int (36,161,25921,36,30,1,46,5,1150,240,192,1649,24272,24272,0,3034,600,1120,1680,1980,19472,15312,10832,8432,6,151,121,30,14,152,122,30,6,75,47,28,34,76,48,28,46,54,24,30,10,55,25,30,2,45,15,30,64,46,16,30,46,6,24,50,76,102,128,154,2431,1911,1351,1051,5835,4587,3243,2523,3536,2779,1965,1529,122,48,25,16)
	 ver37 = Array As Int (37,165,27225,36,30,1,46,5,1150,248,192,1657,25568,25568,0,3196,630,1204,1770,2100,20528,15936,11408,8768,17,152,122,30,4,153,123,30,29,74,46,28,14,75,47,28,49,54,24,30,10,55,25,30,24,45,15,30,46,46,16,30,46,6,28,54,80,106,132,158,2563,1989,1423,1093,6152,4774,3416,2624,3728,2893,2070,1590,123,47,25,16)
	 ver38 = Array As Int (38,169,28561,36,30,1,46,5,1150,256,192,1665,26896,26896,0,3362,660,1260,1860,2220,21616,16816,12016,9136,4,152,122,30,18,153,123,30,13,74,46,28,32,75,47,28,48,54,24,30,14,55,25,30,42,45,15,30,32,46,16,30,46,6,32,58,84,110,136,162,2699,2099,1499,1139,6478,5038,3598,2734,3926,3053,2180,1657,123,47,25,16)
	 ver39 = Array As Int (39,173,29929,36,30,1,46,5,1150,264,192,1673,28256,28256,0,3532,720,1316,1950,2310,22496,17728,12656,9776,20,147,117,30,4,148,118,30,40,75,47,28,7,76,48,28,43,54,24,30,22,55,25,30,10,45,15,30,67,46,16,30,46,6,26,54,82,110,138,166,2809,2213,1579,1219,6742,5312,3790,2926,4086,3219,2297,1773,118,48,25,16)
	 ver40 = Array As Int (40,177,31329,36,30,1,46,5,1150,272,192,1681,29648,29648,0,3706,750,1372,2040,2430,23648,18672,13328,10208,19,148,118,30,6,149,119,30,18,75,47,28,31,76,48,28,34,54,24,30,34,55,25,30,20,45,15,30,61,46,16,30,46,6,30,58,86,114,142,170,2953,2331,1663,1273,7088,5595,3992,3056,4295,3390,2419,1851,119,48,25,16)

	 Private bcolor As Int
	 Private fcolor As Int
	 Private shp As String

End Sub

'Description: Generates a QR code and saves it in the /pictures folder as "QRcode.png" from where it can be loaded into for eg an ImageView
'
'Method of calling:
'QRcode.Draw_QR_Code(message,error_level,mask_pattern,back_color,fore_color,shape)
'
'where 'message' = message to be encoded - type String
'where 'error_level' = the required error level ("L", "M", "Q", "H") - type String
'where 'mask_pattern' = the masking pattern to apply (0 to 7) - type Int
'where 'back_color' = background color of the QR Code - type Int
'where 'fore_color' = foreground color of the QR Code (pixel color) - type Int
'where 'shape' = the shape of the pixels (c or C = circle, s or S = square) - type string
'example of parameters to pass:
'message = "Hallo, how are you today"
'error_level = "H"
'mask_pattern = 7
'back_color = Colors.Yellow
'fore_color = Colors.Black
'shape = "c"
Public Sub Draw_QR_Code(message As String, err_lvl As String, mask As Int,bc As Int,fc As Int, shape As String) As Int                      '(need to add VERSION, ERROR CORRECTION,ENCODING MODE,MASKING PATTERN


Dim QRSTRING As String
Dim len_dat As Int
Dim rev_len_dat_bin As String
Dim len_dat_bin As String
bcolor = bc
fcolor = fc
sf.Initialize

shp = shape
dat = message
'version_no = ver_numb
error_level = err_lvl
tipe = "8-bit-Byte"
maskp = mask

'set_initial_value
ECC_woord = ""

determine_max_len

select_block

QRSTRING = ""

Select Case tipe              'get the correct mode indicator
  Case "Numeric"
    QRSTRING = "0001"
  Case "Alphanumeric"
    QRSTRING = "0010"
  Case "8-bit-Byte"
    QRSTRING = "0100"
End Select

Generate_log_antilog
'determine_max_len

If sf.Len(dat) > maxlen Then   'only encode the maximum alowable number of chars based on selection of Version, ECC level, encoding mode
  dat = sf.Left(dat, maxlen)
End If

len_dat = sf.Len(dat)
rev_len_dat_bin = ""       'this will hold a binary string but it will be in reverse order i.e LSB to MSB

Dim a1 As Int
Dim b1 As Int
Dim c1 As Int

a1 = len_dat    'get the binary string for the length of data. It is initially in reverse order

Do While a1 <> 0          
  b1 = Floor(a1 / 2)
  c1 = a1 - (b1 * 2)
  If c1 = 1 Then rev_len_dat_bin = rev_len_dat_bin & "1"
  If c1 = 0 Then rev_len_dat_bin = rev_len_dat_bin & "0"
  a1 = b1
Loop

len_dat_bin = ""     'now reverse the sting that is in reversed order
For i = 1 To sf.Len(rev_len_dat_bin)
  len_dat_bin = sf.Mid(rev_len_dat_bin, i, 1) & len_dat_bin    'reverse the binary string to be MSB to LSB
Next 

rev_len_dat_bin = sf.Trim(rev_len_dat_bin)
len_dat_bin = sf.Trim(len_dat_bin)

If tipe = "8-bit-Byte" Then
    If version_no >= 1 And version_no <= 9 Then  'version 1 to 9 for 8-bit byte mode
        If sf.Len(len_dat_bin) < 8 Then
          Do While sf.Len(len_dat_bin) < 8
            len_dat_bin = "0" & len_dat_bin    'add a leading zero to get the length indicator to be 8 bits long
          Loop       
        End If
    Else If version_no > 9 Then
        If sf.Len(len_dat_bin) < 16 Then
          Do While sf.Len(len_dat_bin) < 16 
            len_dat_bin = "0" & len_dat_bin   'add a leading zero to get the length indicator to be 16 bits long
          Loop        
        End If
    End If
Else If tipe = "Numeric" Then
    If version_no >= 1 And version_no <= 9 Then  'version 1 to 9 for 8-bit byte mode
        If sf.Len(len_dat_bin) < 10 Then
          Do While sf.Len(len_dat_bin) < 10
            len_dat_bin = "0" & len_dat_bin  'add a leading zero to get the length indicator to be 10 bits long
          Loop        
        End If
    Else If version_no >= 10 And version_no <= 26 Then
        If sf.Len(len_dat_bin) < 12 Then
          Do While sf.Len(len_dat_bin) < 12 
            len_dat_bin = "0" & len_dat_bin   'add a leading zero to get the length indicator to be 12 bits long
          Loop       
        End If
    Else If version_no > 26 Then
        If sf.Len(len_dat_bin) < 14 Then
          Do While sf.Len(len_dat_bin) < 14  
            len_dat_bin = "0" & len_dat_bin 'add a leading zero to get the length indicator to be 14 bits long
          Loop      
        End If
    End If
Else If tipe = "Alphanumeric" Then
    If version_no >= 1 And version_no <= 9 Then  'version 1 to 9 for 8-bit byte mode
        If sf.Len(len_dat_bin) < 9 Then
          Do While sf.Len(len_dat_bin) < 9 
            len_dat_bin = "0" & len_dat_bin 'add a leading zero to get the length indicator to be 9 bits long
          Loop       
        End If
    Else If version_no > 9 And version_no < 27 Then
        If sf.Len(len_dat_bin) < 11 Then
          Do While sf.Len(len_dat_bin) < 11 
            len_dat_bin = "0" & len_dat_bin 'add a leading zero to get the length indicator to be 11 bits long
          Loop       
        End If
    Else If version_no > 26 Then
        If sf.Len(len_dat_bin) < 13 Then
          Do While sf.Len(len_dat_bin) < 13  
            len_dat_bin = "0" & len_dat_bin 'add a leading zero to get the length indicator to be 13 bits long
          Loop     
        End If
    End If
End If

QRSTRING = QRSTRING & len_dat_bin 'data type number (4 bits long - type) + input string length (padded to correct length)

Dim char_to_convert As String
Dim waarde As Int
Dim getal As String

If tipe = "8-bit-Byte" Then
  For i = 1 To len_dat
    char_to_convert = sf.Mid(dat, i, 1)   'get each individual char in the string
    waarde = conv(char_to_convert)       'get the ASCII value for the chars
    z(i) = waarde                         'store the ASCII values in z()
  Next
  For i = 1 To len_dat
    getal = convert_to_binary(z(i))     'convert each ASCII value to binary equivalent
    If sf.Len(getal) < 8 Then            'and then pad it with leading zeros to be 8 bits long
      Do While sf.Len(getal) < 8
        getal = "0" & getal
      Loop 
    End If

    QRSTRING = QRSTRING & sf.Trim(getal)
  Next
End If

If tipe = "Alphanumeric" Then
    a1 = Floor(len_dat / 2)
    b1 = a1 * 2
    
    For i = 1 To len_dat
      char_to_convert = sf.Mid(dat, i, 1)
      waarde = anconv(char_to_convert)
      z(i) = waarde
    Next
    
    If len_dat - b1 = 1 Then                  'uneven number of chars in the input string
      For i = 1 To len_dat - 1 Step 2
        waarde = (z(i) * 45) + z(i + 1)
        getal = convert_to_binary(waarde)
        If sf.Len(getal) < 11 Then
          Do While sf.Len(getal) < 11 
            getal = "0" & getal
          Loop        'alphanumeric "words" must be 11 bits long
        End If
        QRSTRING = QRSTRING & sf.Trim(getal)
       Next 
       waarde = z(len_dat)
       getal = convert_to_binary(waarde)
       If sf.Len(getal) < 6 Then               'do the uneven numbered char
         Do While sf.Len(getal) < 6  
           getal = "0" & getal
         Loop         'pad it with zeros - it needs to be 6 bits long
       End If
       QRSTRING = QRSTRING & sf.Trim(getal)
    End If
    
    If len_dat - b1 = 0 Then                  'even number of chars in the input sting
      For i = 1 To len_dat - 1 Step 2
        waarde = (z(i) * 45) + z(i + 1)
        getal = convert_to_binary(waarde)
        If sf.Len(getal) < 11 Then
          Do While sf.Len(getal) < 11
            getal = "0" & getal
          Loop
        End If
        QRSTRING = QRSTRING & sf.Trim(getal)
      Next 
    End If
End If

If tipe = "Numeric" Then

    a1 = Floor(len_dat / 3 )                       'going to pack 3 chars per 10 bits
    b1 = a1 * 3                               'how many multiples of 3 chars in the input string
    
    If len_dat - b1 = 2 Then                  'remainder of two chars
      For i = 1 To len_dat - 2 Step 3
        waarde = sf.Val(sf.Mid(dat, i, 3))
        getal = convert_to_binary(waarde)
        If sf.Len(getal) < 10 Then
          Do While sf.Len(getal) < 10 
            getal = "0" & getal
          Loop        'numeric "words" must be 10 bits long
        End If
        QRSTRING = QRSTRING & sf.Trim(getal)
       Next 
       waarde = sf.Val(sf.Mid(dat, i, 2))
       getal = convert_to_binary(waarde)
       If sf.Len(getal) < 7 Then               '2 remaining chars represented by 7 bits
         Do While sf.Len(getal) < 7 
           getal = "0" & getal
         Loop          'pad it with zeros - it needs to be 7 bits long
       End If
       QRSTRING = QRSTRING & sf.Trim(getal)
    End If

  
    If len_dat - b1 = 1 Then                  'remainder of 1 char
      For i = 1 To len_dat - 1 Step 3
        waarde = sf.Val(sf.Mid(dat, i, 3))
        getal = convert_to_binary(waarde)
        If sf.Len(getal) < 10 Then
          Do While sf.Len(getal) < 10  
            getal = "0" & getal
          Loop       'numeric "words" must be 10 bits long
        End If
        QRSTRING = QRSTRING & sf.Trim(getal)
       Next 
       waarde = sf.Val(sf.Mid(dat, i, 1))
       getal = convert_to_binary(waarde)
       If sf.Len(getal) < 4 Then               'one remaining char represented by 4 bits
         Do While sf.Len(getal) < 4 
           getal = "0" & getal
         Loop          'pad it with zeros - it needs to be 4 bits long
       End If
       QRSTRING = QRSTRING & sf.Trim(getal)
    End If
      
    If len_dat - b1 = 0 Then                  'multiple of 3 chars in the input string
      For i = 1 To len_dat Step 3
        waarde = sf.Val(sf.Mid(dat, i, 3))
        getal = convert_to_binary(waarde)
        If sf.Len(getal) < 10 Then
          Do While sf.Len(getal) < 10  
            getal = "0" & getal
          Loop       'numeric "words" must be 10 bits long
        End If
        QRSTRING = QRSTRING & sf.Trim(getal)
       Next 
    End If

End If

If sf.Len(QRSTRING) <= max_data_bits - 4 Then      'add 4 zeros if length is not at least max_data_bits minus 4 long
   QRSTRING = QRSTRING & "0000"
 Else
   Do While sf.Len(QRSTRING) < max_data_bits
    QRSTRING = QRSTRING & "0"
   Loop   'else if longer then add just enough zeros until it is max_data_bits long
 End If



a1 = sf.Len(QRSTRING)
b1 = a1 Mod 8         'check if the sting is multiples of 8 bits

If b1 <> 0 Then
  Do While sf.Len(QRSTRING) Mod 8 <> 0
   QRSTRING = QRSTRING & "0"         'if not, then pad with zeros until it is multiples of 8 bits
  Loop 
End If

Dim last_flag As Int
last_flag = 0
If sf.Len(QRSTRING) < max_data_bits Then           'if still less than max_data_bits then add stings below alternatively
  Do While sf.Len(QRSTRING) <> max_data_bits  
    If last_flag = 0 Then
      QRSTRING = QRSTRING & "11101100"    'pad the sting alternatively with.....
      last_flag = 1
    Else
     QRSTRING = QRSTRING & "00010001"     '....and with....
     last_flag = 0
    End If
  Loop         'sting must be max_data_bits long
End If

Dim teller As Int

If dw2 <> 0 Then
  Dim data_block(block1 + block2 + 1, dw2 + 1) As Int
  For j = 1 To dw2
    For i = 1 To block1 + block2
      data_block(i, j) = 0              'load the array with zeros
    Next 
  Next 
  teller = 1
  For i = 1 To block1                     'load the order in which the data block will assembled to for the interleaved data string
    For j = 1 To dw1
      data_block(i, j) = teller
      teller = teller + 1
    Next 
  Next 
  For i = block1 + 1 To block1 + block2                      'load the order in which the data block will assembled to for the interleaved data string
    For j = 1 To dw2
      data_block(i, j) = teller
      teller = teller + 1
    Next 
  Next 
Else
  Dim data_block(block1 + block2 + 1 , dw1 + 1) As Int
  For j = 1 To dw1
    For i = 1 To block1
      data_block(i, j) = 0              'load the array with zeros
    Next 
  Next 
  teller = 1
  For i = 1 To block1                      'load the order in which the data block will assembled to for the interleaved data string
    For j = 1 To dw1
      data_block(i, j) = teller
      teller = teller + 1
    Next 
  Next 
End If

Dim ecc_block(block1 + block2 + 1, ew1 + 1) As Int  'could also be ew2 -> they are the same size
For j = 1 To ew1
  For i = 1 To block1 + block2
    ecc_block(i, j) = 0              'load the array with zeros
  Next 
Next 

teller = 1
For i = 1 To block1 + block2                      'load the order in which the data block will assembled to for the interleaved data string
  For j = 1 To ew1
    ecc_block(i, j) = teller
    teller = teller + 1
  Next 
Next 

Dim begin,einde,vlag,dec_answer As Int
Dim string_to_convert As String

begin = 1
einde = dw1 * 8
vlag = 0
Do While vlag <> block1 + block2  
  no_of_error_words = ew1
  mess_coef_teller = 1
  If vlag < block1 Then
    For i = begin To einde Step 8
      string_to_convert = sf.Mid(QRSTRING, i, 8)
      dec_answer = convert_to_decimal(string_to_convert)   'get the decimal value
      message_coeff(mess_coef_teller) = dec_answer          'this is the message polynomial coefficients
      mess_coef_teller = mess_coef_teller + 1
   Next 
    begin = einde + 1
    If vlag = block1 - 1 Then
      einde = einde + dw2 * 8
    Else
      einde = einde + dw1 * 8
    End If
  End If

  If vlag >= block1 Then
    For i = begin To einde Step 8    'take portion of the data sting to calculate ECC words for
      string_to_convert = sf.Mid(QRSTRING, i, 8)
      dec_answer = convert_to_decimal(string_to_convert)   'get the decimal value
      message_coeff(mess_coef_teller) = dec_answer          'this is the message polynomial coeefficients
      mess_coef_teller = mess_coef_teller + 1
    Next 
    begin = einde + 1
    einde = einde + dw2 * 8
  End If
  
  mess_coef_teller = mess_coef_teller - 1

  For i = 1 To 196
    s4c1(i) = ""
  Next 

  For i = 1 To mess_coef_teller
    s4c1(i) = message_coeff(i)
  Next 
  
  generator_polynomial(no_of_error_words)    'create the generator polynomial with 30 ECC words for each half of the complete data string i.e half of 1856 bits

  q = ECC_words          'calculate the ECC words for each half of the message polynomial
  vlag = vlag + 1
Loop    'do it for as many times as the number of data blocks that are required

a = QRSTRING
b = q   'first block of error correction words

anew = ""                       'now build a new interleaved string

If dw2 <> 0 Then
  For j = 1 To dw2
    For i = 1 To block1 + block2
      If data_block(i, j) <> 0 Then
        anew = anew & sf.Mid(a, data_block(i, j) * 8 - 7, 8) 'take 8 bits from the first half of the data sting...
      End If
    Next      'repeat this process until the whole data string has been interleaved
  Next 
End If

If dw2 = 0 Then
  For j = 1 To dw1
    For i = 1 To block1 + block2
      If data_block(i, j) <> 0 Then
        anew = anew & sf.Mid(a, data_block(i, j) * 8 - 7, 8) 'take 8 bits from the first half of the data sting...
      End If
    Next     'repeat this process until the whole data string has been interleaved
  Next  
End If

' anew now holds the interleaved string

Dim d, dnew As String
d = ""                           'now do the same interleaving with the two ECC word blocks
d = b
dnew = ""                        'this will be the interleaved ECC word string
  For j = 1 To ew1
    For i = 1 To block1 + block2
      If ecc_block(i, j) <> 0 Then
        dnew = dnew & sf.Mid(d, ecc_block(i, j) * 8 - 7, 8) 'take 8 bits from the first half of the data sting...
      End If
    Next      'repeat this process until the whole data string has been interleaved
  Next 

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

anew = anew & dnew       'now build one long string of interleaved data followed by interleaved ECC words

place_matrix_data

Return version_no

End Sub

Private Sub conv (karak As String) As Int

Dim char1 As Int

'Log (Chr(199))
Select Case karak     'decimal values for ASCII chars
  Case Chr(13)        'Carriage Return    
    char1 = 13
  Case Chr(10)        'Line Feed
    char1 = 10
  Case " "            'Space
   char1 = 32
  Case "!"
   char1 = 33
  Case """"
   char1 = 34
  Case "#"
   char1 = 35
  Case "$"
   char1 = 36
  Case "%"
   char1 = 37
  Case "&"
   char1 = 38
  Case "'"
   char1 = 39
  Case "("
   char1 = 40
  Case ")"
   char1 = 41
  Case "*"
   char1 = 42
  Case "+"
   char1 = 43
  Case ","
   char1 = 44
  Case "-"
   char1 = 45
  Case "."
   char1 = 46
  Case "/"
   char1 = 47
  Case "0"
   char1 = 48
  Case "1"
   char1 = 49
  Case "2"
   char1 = 50
  Case "3"
   char1 = 51
  Case "4"
   char1 = 52
  Case "5"
   char1 = 53
  Case "6"
   char1 = 54
  Case "7"
   char1 = 55
  Case "8"
   char1 = 56
  Case "9"
   char1 = 57
  Case ":"
   char1 = 58
  Case ";"
   char1 = 59
  Case "<"
   char1 = 60
  Case "="
   char1 = 61
  Case ">"
   char1 = 62
  Case "?"
   char1 = 63
  Case "@"
   char1 = 64
  Case "A"
   char1 = 65
  Case "B"
   char1 = 66
  Case "C"
   char1 = 67
  Case "D"
   char1 = 68
  Case "E"
   char1 = 69
  Case "F"
   char1 = 70
  Case "G"
   char1 = 71
  Case "H"
   char1 = 72
  Case "I"
   char1 = 73
  Case "J"
   char1 = 74
  Case "K"
   char1 = 75
  Case "L"
   char1 = 76
  Case "M"
   char1 = 77
  Case "N"
   char1 = 78
  Case "O"
   char1 = 79
  Case "P"
   char1 = 80
  Case "Q"
   char1 = 81
  Case "R"
   char1 = 82
  Case "S"
   char1 = 83
  Case "T"
   char1 = 84
  Case "U"
   char1 = 85
  Case "V"
   char1 = 86
  Case "W"
   char1 = 87
  Case "X"
   char1 = 88
  Case "Y"
   char1 = 89
  Case "Z"
   char1 = 90
  Case "["
   char1 = 91     
  Case "]"
   char1 = 93     
  Case "^"
   char1 = 94
  Case "_"
   char1 = 95
  Case "`"
   char1 = 96
  Case "a"
   char1 = 97
  Case "b"
   char1 = 98
  Case "c"
   char1 = 99
  Case "d"
   char1 = 100
  Case "e"
   char1 = 101
  Case "f"
   char1 = 102
  Case "g"
   char1 = 103
  Case "h"
   char1 = 104
  Case "i"
   char1 = 105
  Case "j"
   char1 = 106
  Case "k"
   char1 = 107
  Case "l"
   char1 = 108
  Case "m"
   char1 = 109
  Case "n"
   char1 = 110
  Case "o"
   char1 = 111
  Case "p"
   char1 = 112
  Case "q"
   char1 = 113
  Case "r"
   char1 = 114
  Case "s"
   char1 = 115
  Case "t"
   char1 = 116
  Case "u"
   char1 = 117
  Case "v"
   char1 = 118
  Case "w"
   char1 = 119
  Case "x"
   char1 = 120
  Case "y"
   char1 = 121
  Case "z"
   char1 = 122
  Case "{"
   char1 = 123
  Case "|"
   char1 = 124
  Case "}"
   char1 = 125
  Case Chr(126)
   char1 = 126
  Case Chr(160)
   char1 = 160   
  Case Chr(161)
   char1 = 161      
  Case Chr(162)
   char1 = 162      
  Case Chr(163)
   char1 = 163     
  Case Chr(164)
   char1 = 164      
  Case Chr(165)
   char1 = 165      
  Case Chr(166)
   char1 = 166      
  Case Chr(167)
   char1 = 167         
  Case Chr(168)
   char1 = 168         
  Case Chr(169)
   char1 = 169         
  Case Chr(170)
   char1 = 170
  Case Chr(171)
   char1 = 171   
  Case Chr(172)
   char1 = 172  
  Case Chr(173)
   char1 = 173   
  Case Chr(174)
   char1 = 174  
  Case Chr(175)
   char1 = 175   
  Case Chr(176)
   char1 = 176   
  Case Chr(177)
   char1 = 177   
  Case Chr(178)
   char1 = 178   
  Case Chr(179)
   char1 = 179   
  Case Chr(180)
   char1 = 180   
  Case Chr(181)
   char1 = 181     
  Case Chr(182)
   char1 = 182     
  Case Chr(183)
   char1 = 183     
  Case Chr(184)
   char1 = 184     
  Case Chr(185)
   char1 = 185     
  Case Chr(186)
   char1 = 186     
  Case Chr(187)
   char1 = 187     
  Case Chr(188)
   char1 = 188     
  Case Chr(189)
   char1 = 189        
  Case Chr(190)
   char1 = 190        
  Case Chr(191)
   char1 = 191      
  Case Chr(192)
   char1 = 192         
  Case Chr(193) 
   char1 = 193
  Case Chr(194) 
   char1 = 194
  Case Chr(195) 
   char1 = 195
  Case Chr(196) 
   char1 = 196
  Case Chr(197) 
   char1 = 197
  Case Chr(198) 
   char1 = 198
  Case Chr(199) 
   char1 = 199
  Case Chr(200)
   char1 = 200
  Case Chr(201)
   char1 = 201
  Case Chr(202) 
   char1 = 202
  Case Chr(203)
   char1 = 203
  Case Chr(204)
   char1 = 204
  Case Chr(205)
   char1 = 205
  Case Chr(206) 
   char1 = 206
  Case Chr(207)
   char1 = 207
  Case Chr(208)
   char1 = 208
  Case Chr(209)
   char1 = 209
  Case Chr(210)
   char1 = 210
  Case Chr(211)
   char1 = 211
  Case Chr(212)
   char1 = 212
  Case Chr(213)
   char1 = 213
  Case Chr(214)
   char1 = 214
  Case Chr(215)
   char1 = 215
  Case Chr(216) 
   char1 = 216
  Case Chr(217)
   char1 = 217
  Case Chr(218)
   char1 = 218
  Case Chr(219)
   char1 = 219
  Case Chr(220)
   char1 = 220
  Case Chr(221)
   char1 = 221
  Case Chr(222) 
   char1 = 222
  Case Chr(223)
   char1 = 223
  Case Chr(224) 
   char1 = 224
  Case Chr(225)
   char1 = 225
  Case Chr(226)
   char1 = 226
  Case Chr(227)
   char1 = 227
  Case Chr(228)
   char1 = 228
  Case Chr(229)
   char1 = 229
  Case Chr(230)
   char1 = 230
  Case Chr(231) 
   char1 = 231
  Case Chr(232)
   char1 = 232
  Case Chr(233)
   char1 = 233
  Case Chr(234)
   char1 = 234
  Case Chr(235)
   char1 = 235
  Case Chr(236)
   char1 = 236
  Case Chr(237) 
   char1 = 237
  Case Chr(238)
   char1 = 238
  Case Chr(239) 
   char1 = 239
  Case Chr(240) 
   char1 = 240
  Case Chr(241)
   char1 = 241
  Case Chr(242) 
   char1 = 242
  Case Chr(243)
   char1 = 243
  Case Chr(244)
   char1 = 244
  Case Chr(245)
   char1 = 245
  Case Chr(246)
   char1 = 246
  Case Chr(247)
   char1 = 247
  Case Chr(248)
   char1 = 248
  Case Chr(249) 
   char1 = 249
  Case Chr(250)
   char1 = 250
  Case Chr(251)
   char1 = 251
  Case Chr(252) 
   char1 = 252
  Case Chr(253) 
   char1 = 253
  Case Chr(254)
   char1 = 254
  Case Chr(255)
   char1 = 255
 End Select

Return char1

End Sub

Private Sub convert_to_decimal(bb As String) As Int 

	Dim waarde As Int
	waarde = 0    'convert binary to decimal
	For i = 1 To 8
	 If sf.Mid(bb, i, 1) = 1 Then
	   waarde = waarde + Power(2,(8 - i))
	 End If
	Next

Return waarde

End Sub

Private Sub Generate_log_antilog

logg(0) = 1
alogg(0) = 1

For i = 1 To 255       'set up log / anti-log tables
  alogg(i) = alogg(i - 1) * 2
  If alogg(i) >= gf Then 
    alogg(i) = Bit.Xor(alogg(i), PP)
  End If	
  logg(alogg(i)) = i
Next

For i = 256 To 511
   alogg(i) = alogg(i - 255)
Next

logg(alogg(255)) = 255 '# Set last missing elements in gf_log[]

'Log("done")

End Sub

Private Sub convert_to_binary(getal1 As String) As String

Dim rev_len_dat_bin, len_dat_bin As String
Dim e,f,g As Int

rev_len_dat_bin = ""
e = sf.Val(getal1)
If e <> 0 Then
	Do While e <> 0
	  f = Floor(e / 2)
	  g = e - (f * 2)
	  If g = 1 Then rev_len_dat_bin = rev_len_dat_bin & "1"
	  If g = 0 Then rev_len_dat_bin = rev_len_dat_bin & "0"
	  e = f
	Loop
	'reverse the binary string to get the order right
	len_dat_bin = ""
	For x = 1 To sf.len(rev_len_dat_bin)
	  len_dat_bin = sf.Mid(rev_len_dat_bin, x, 1) & len_dat_bin
	Next

	rev_len_dat_bin = sf.Trim(rev_len_dat_bin)
	len_dat_bin = sf.Trim(len_dat_bin)

	getal1 = len_dat_bin
	
	Return getal1
Else
	getal1 = "00000000"
	Return getal1
End If

End Sub

Private Sub anconv(char1 As String) As Int

Dim agetal As Int

Select Case char1
  Case "0"
   agetal = 0
  Case "1"
   agetal = 1
  Case "2"
   agetal = 2
  Case "3"
   agetal = 3
  Case "4"
   agetal = 4
  Case "5"
   agetal = 5
  Case "6"
   agetal = 6
  Case "7"
   agetal = 7
  Case "8"
   agetal = 8
  Case "9"
   agetal = 9
  Case "A"
   agetal = 10
  Case "a"
   agetal = 10
  Case "B"
   agetal = 11
  Case "b"
   agetal = 11
  Case "C"
   agetal = 12
  Case "c"
   agetal = 12
  Case "D"
   agetal = 13
  Case "d"
   agetal = 13
  Case "E"
   agetal = 14
  Case "e"
   agetal = 14
  Case "F"
   agetal = 15
  Case "f"
   agetal = 15
  Case "G"
   agetal = 16
  Case "g"
   agetal = 16
  Case "H"
   agetal = 17
  Case "h"
   agetal = 17
  Case "I"
   agetal = 18
  Case "i"
   agetal = 18
  Case "J"
   agetal = 19
  Case "j"
   agetal = 19
  Case "K"
   agetal = 20
  Case "k"
   agetal = 20
  Case "L"
   agetal = 21
  Case "l"
   agetal = 21
  Case "M"
   agetal = 22
  Case "m"
   agetal = 22
  Case "N"
   agetal = 23
  Case "n"
   agetal = 23
  Case "O"
   agetal = 24
  Case "o"
   agetal = 24
  Case "P"
   agetal = 25
  Case "p"
   agetal = 25
  Case "Q"
   agetal = 26
  Case "q"
   agetal = 26
  Case "R"
   agetal = 27
  Case "r"
   agetal = 27
  Case "S"
   agetal = 28
  Case "s"
   agetal = 28
  Case "T"
   agetal = 29
  Case "t"
   agetal = 29
  Case "U"
   agetal = 30
  Case "u"
   agetal = 30
  Case "V"
   agetal = 31
  Case "v"
   agetal = 31
  Case "W"
   agetal = 32
  Case "w"
   agetal = 32
  Case "X"
   agetal = 33
  Case "x"
   agetal = 33
  Case "Y"
   agetal = 34
  Case "y"
   agetal = 34
  Case "Z"
   agetal = 35
  Case "z"
   agetal = 35
  Case " "
   agetal = 36 '(space)
  Case "$"
   agetal = 37
  Case "%"
   agetal = 38
  Case "*"
   agetal = 39
  Case "+"
   agetal = 40
  Case "-"
   agetal = 41
  Case "."
   agetal = 42
  Case "/"
   agetal = 43
  Case ":"
   agetal = 44
End Select

Return agetal

End Sub

Private Sub generator_polynomial(noew As Int) As Int

Dim a3() As Int

'the below is the x-coefficients of the generator polynomials and not the alpha coefficients (i.e alpha converted to x with logg / alogg table)
Select Case noew
  Case  7
    a3 = Array As Int(1, 127, 122, 154, 164, 11, 68, 117) 
  Case  10
    a3 = Array As Int(1, 216, 194, 159, 111, 199, 94, 95, 113, 157, 193) 
  Case  13
    a3 = Array As Int(1, 137, 73, 227, 17, 177, 17, 52, 13, 46, 43, 83, 132, 120)
  Case  15
    a3 = Array As Int(1, 29, 196, 111, 163, 112, 74, 10, 105, 105, 139, 132, 151, 32, 134, 26) 
  Case  16
    a3 = Array As Int(1, 59, 13, 104, 189, 68, 209, 30, 8, 163, 65, 41, 229, 98, 50, 36, 59) 
  Case  17
    a3 = Array As Int(1, 119, 66, 83, 120, 119, 22, 197, 83, 249, 41, 143, 134, 85, 53, 125, 99, 79) 
  Case  18
    a3 = Array As Int(1, 239, 251, 183, 113, 149, 175, 199, 215, 240, 220, 73, 82, 173, 75, 32, 67, 217, 146) 
  Case  20
    a3 = Array As Int(1, 152, 185, 240, 5, 111, 99, 6, 220, 112, 150, 69, 36, 187, 22, 228, 198, 121, 121, 165, 174) 
  Case  22
    a3 = Array As Int(1, 89, 179, 131, 176, 182, 244, 19, 189, 69, 40, 28, 137, 29, 123, 67, 253, 86, 218, 230, 26, 145, 245) 
  Case  24
    a3 = Array As Int(1, 122, 118, 169, 70, 178, 237, 216, 102, 115, 150, 229, 73, 130, 72, 61, 43, 206, 1, 237, 247, 127, 217, 144, 117) 
  Case  26
    a3 = Array As Int(1, 246, 51, 183, 4, 136, 98, 199, 152, 77, 56, 206, 24, 145, 40, 209, 117, 233, 42, 135, 68, 70, 144, 146, 77, 43, 94) 
  Case  28
    a3 = Array As Int(1, 252, 9, 28, 13, 18, 251, 208, 150, 103, 174, 100, 41, 167, 12, 247, 56, 117, 119, 233, 127, 181, 100, 121, 147, 176, 74, 58, 197) 
  Case  30
    a3 = Array As Int(1, 212, 246, 77, 73, 195, 192, 75, 98, 5, 70, 103, 177, 22, 217, 138, 51, 181, 246, 72, 25, 18, 46, 228, 74, 216, 195, 11, 106, 130, 150) 

End Select

For i = 1 To 35
  s4c4(i) = 0      
Next

For i = 1 To noew + 1
  s4c4(i) = a3(i - 1)
Next

End Sub

Private Sub ECC_words                     'sub that calculates the error correction words

   Dim msg_out(2956) As Int
   Dim coef,bollie As Int

   For i =0 To 2955
	  msg_out(i) = 0
   Next  
   For i = 0 To mess_coef_teller - 1
      msg_out(i) = message_coeff(i + 1)
   Next
   For i = 0 To mess_coef_teller - 1
      coef = msg_out(i)
      If coef <> 0 Then
         For j = 0 To no_of_error_words
		    gf_mul_result = gf_mul(s4c4(j + 1), coef)
            bollie = Bit.Xor(msg_out(i + j), gf_mul_result)
            msg_out(i + j) = bollie
         Next
      End If
   Next
   
   For i = 0 To mess_coef_teller - 1
      msg_out(i) = message_coeff(i + 1)
   Next


For i = 1 To no_of_error_words
     s4c24(i) = msg_out(mess_coef_teller - 1 + i)
Next

Dim getal As String
Dim getalle As Int
For i = 1 To no_of_error_words
    getalle = s4c24(i)
    getal = convert_to_binary(getalle)
    Do While sf.Len(getal) < 8
      If sf.Len(getal) < 8 Then
       getal = "0" & getal
      End If
    Loop 
      ECC_woord = ECC_woord & getal
      q = ECC_woord
Next
  
Return q

End Sub

Private Sub gf_mul(x As Int, y As Int) As Int
   
   If x = 0 OR y = 0 Then
      gf_mul_result = 0
   Else 
     gf_mul_result = alogg(logg(x) + logg(y))
   End If

   Return gf_mul_result   
End Sub

Private Sub place_matrix_data

Dim x1(31335) As Int   'stores the x coordinate of data modules
Dim y1(31335) As Int    'stores the y coordinate of data modules


'The below code determines the x,y coordinates for each of the bits in the complete
'interleaved string and store them in x(i) and y(i) -> rows and columns

draw_qr      'call procedure to draw the QR matrix - not populated but selected mask pattern will be added

Dim ry,col1,col2,rgt,teller,up,a3,i As Int

ry = qr_size                     'we start on excel row qr size +1
col1 = qr_size                   'we start first column on excel column qr size +1
col2 = col1 - 1              'we do two columns at a time when packing bits - thus column y is used along with y+1, etc
rgt = 1                      'marker for the righthand/lefthand column of the two columns that we are currently working with
teller = 1                   'keeps count of the number of valid cells that can take a bit (the count will exclude format info, version info and all patterns)
up = 1                       'keep track of the direction that we are going ---> up or down


Do While teller <> max_mod + 1 
    If rgt = 1 AND up = 1 Then        'in the right column of the two columns and going up
      a3 = s5cij(ry, col1)
      If a3 <> 9 AND a3 <> 10 AND ry > 0 Then     'it is a cell that can accommodate a bit
        x1(teller) = ry
        y1(teller) = col1
        rgt = 0                       'next bit to go in the left column
        teller = teller + 1
      Else
        rgt = 0
      End If
	
    Else If rgt = 0 AND up = 1 Then    'in the left column of the two columns and going up
       a3 = s5cij(ry, col2)
       If a3 <> 9 AND a3 <> 10 AND ry > 0 Then
         x1(teller) = ry
         y1(teller) = col2
         rgt = 1                      'next bit to go in the right column
         teller = teller + 1
         ry = ry - 1                  'go up 1 row
       Else
         rgt = 1
         ry = ry - 1
        End If
    
    Else If rgt = 1 AND up = 0 Then    'in the right column of the two columns and going down
       a3 = s5cij(ry, col1)
       If a3 <> 9 AND a3 <> 10 AND ry < qr_size + 1 Then    '***** qr_size + 2
         x1(teller) = ry
         y1(teller) = col1
         rgt = 0
         teller = teller + 1
       Else
         rgt = 0
       End If
    
    Else If rgt = 0 AND up = 0 Then   'in the left column of the two columns and going down
       a3 = s5cij(ry, col2)
       If a3 <> 9 AND a3 <> 10 AND ry < qr_size + 1 Then                '***** qr_size + 2
         x1(teller) = ry
         y1(teller) = col2
         rgt = 1
         teller = teller + 1
         ry = ry + 1             'go down 1 row
       Else
         rgt = 1
         ry = ry + 1
       End If
    End If
    
    If ry = 0 AND col1 > 1 AND up = 1 AND rgt = 1 Then
      ry = 1
      rgt = 1
      up = 0
      col1 = col1 - 2
      If col1 = 7 Then col1 = 6    'the vertical timing pattern does not take bits - so skip this column
      col2 = col1 - 1
    Else If ry = qr_size + 1 AND col1 > 1 AND up = 0 AND rgt = 1 Then  '******* qr_size + 2
      ry = qr_size
      rgt = 1
      up = 1
      col1 = col1 - 2
      If col1 = 7 Then col1 = 6
      col2 = col1 - 1
    End If
Loop    'we now have the x,y coordinates of all the data modules and it is stored in arrays x(i) and y(i)

'The code that follows fill the matix from the string of data bits

If sf.Len(anew) < max_mod Then    'if the string length is less than the required length then pad it with trailing zeros
  Do While sf.Len(anew) <> max_mod
    anew = anew & "1"   'Changed from 0 to 1 on 28 March 2014
  Loop 
End If

teller = 1                    'keep track of the bits to place according to x(teller) and y(teller)

Dim bitt As Int

Do While teller <> max_mod + 1                             'place the black/white modules in the matrix
 bitt = sf.Val(sf.Mid(anew, teller, 1))  'get a bit at a time

 If s5cij(x1(teller), y1(teller)) = 0 AND bitt = 0 Then
   s5cij(x1(teller), y1(teller)) = 10   'decide what color to make the module baased on the bit value and the mask bit in that specific module -> 1 = black and 2 = white
 Else If s5cij(x1(teller), y1(teller)) = 0 AND bitt = 1 Then
     s5cij(x1(teller), y1(teller)) = 9

 Else If s5cij(x1(teller), y1(teller)) = 1 AND bitt = 0 Then
    s5cij(x1(teller), y1(teller)) = 9
 Else If s5cij(x1(teller), y1(teller)) = 1 AND bitt = 1 Then
     s5cij(x1(teller), y1(teller)) = 10
 End If
 teller = teller + 1
Loop    'if teller reaches max_mod + 1 then we have placed all the bits in the string
                            'inside the matrix
teken1 
							
End Sub

Private Sub draw_qr

'the pattern of a finding pattern
Dim fp() As Int
fp = Array As Int(1, 1, 1, 1, 1, 1, 1, _
                  1, 0, 0, 0, 0, 0, 1, _
                  1, 0, 1, 1, 1, 0, 1, _
                  1, 0, 1, 1, 1, 0, 1, _
                  1, 0, 1, 1, 1, 0, 1, _
                  1, 0, 0, 0, 0, 0, 1, _
                  1, 1, 1, 1, 1, 1, 1)
'the pattern of an alignment pattern
Dim ap() As Int
ap = Array As Int(1, 1, 1, 1, 1, _
                  1, 0, 0, 0, 1, _
                  1, 0, 1, 0, 1, _
                  1, 0, 0, 0, 1, _
                  1, 1, 1, 1, 1)

Dim apx(50)   'store x coordinates of alignment pattern central modules
Dim apy(50)   'store y coordinates of alignment pattern central modules
Dim co_x(8)
Dim co_y(8)

'Log(maskp)

If maskp = 0 Then
  For i = 1 To qr_size                  'load mask pattern 0 in the whole matrix
    For j = 1 To qr_size                'we will null the mask in all format/version modules as we place these modules
      s5cij(i, j) = ((i - 1) + (j - 1)) Mod 2
    Next 
  Next 
End If

If maskp = 1 Then
  For i = 1 To qr_size                  'load mask pattern 1 in the whole matrix
    For j = 1 To qr_size                'we will null the mask in all format/version modules as we place these modules
      s5cij(i, j) = (i - 1) Mod 2
    Next 
  Next 
End If

If maskp = 2 Then
  For i = 1 To qr_size                  'load mask pattern 2 in the whole matrix
    For j = 1 To qr_size                'we will null the mask in all format/version modules as we place these modules
      s5cij(i, j) = (j - 1) Mod 3
      If s5cij(i, j) > 1 Then s5cij(i, j) = 1
    Next 
  Next 
End If

If maskp = 3 Then
  For i = 1 To qr_size                  'load mask pattern 3 in the whole matrix
    For j = 1 To qr_size                'we will null the mask in all format/version modules as we place these modules
      s5cij(i, j) = ((i - 1) + (j - 1)) Mod 3
      If s5cij(i, j) > 1 Then s5cij(i, j) = 1
    Next 
  Next 
End If

If maskp = 4 Then
  For i = 1 To qr_size                  'load mask pattern 4 in the whole matrix
    For j = 1 To qr_size                'we will null the mask in all format/version modules as we place these modules
      s5cij(i, j) = ((Floor((i - 1) / 2  )) + (Floor((j - 1) / 3))) Mod 2
    Next 
  Next 
End If

If maskp = 5 Then
  For i = 1 To qr_size                  'load mask pattern 5 in the whole matrix
    For j = 1 To qr_size                'we will null the mask in all format/version modules as we place these modules
      s5cij(i, j) = (((i - 1) * (j - 1)) Mod 2) + (((i - 1) * (j - 1)) Mod 3)
      If s5cij(i, j) > 1 Then s5cij(i, j) = 1
    Next 
  Next 
End If

If maskp = 6 Then
  For i = 1 To qr_size                  'load mask pattern 6 in the whole matrix
    For j = 1 To qr_size                'we will null the mask in all format/version modules as we place these modules
      s5cij(i, j) = ((((i - 1) * (j - 1)) Mod 2) + (((i - 1) * (j - 1)) Mod 3)) Mod 2
    Next 
  Next 
End If

If maskp = 7 Then
  For i = 1 To qr_size                  'load mask pattern 7 in the whole matrix
    For j = 1 To qr_size                'we will null the mask in all format/version modules as we place these modules
      s5cij(i, j) = ((((i - 1) + (j - 1)) Mod 2) + (((i - 1) * (j - 1)) Mod 3)) Mod 2
    Next 
  Next 
End If
Dim teller As Int
teller = 1              'put down the 3 finding patterns
For i = 1 To 7
  For j = 1 To 7
    If fp(teller - 1) = 1 Then     'fp() stores the patterns of the finding pattern
      s5cij(i, j) = 10         'it will become a white square
      s5cij(i, qr_size - 7 + j) = 10
      s5cij(qr_size - 7 + i, j) = 10
      teller = teller + 1
    Else
      s5cij(i, j) = 9  'it will become a black square
      s5cij(i, qr_size - 7 + j) = 9
      s5cij(qr_size - 7 + i, j) = 9
      teller = teller + 1
    End If
  Next 
Next 
For j = 9 To qr_size - 7    'place horizontal timing pattern
  If j Mod 2 = 1 Then
    s5cij(7, j) = 10
  Else
    s5cij(7, j) = 9
  End If
Next 
For i = 9 To qr_size - 7         'place vertical timing pattern
  If i Mod 2 = 1 Then
    s5cij(i, 7) = 10
  Else
    s5cij(i, 7) = 9
  End If
Next
s5cij(qr_size - 7, 9) = 10  'place reserved bit to the right and top of bottom left finding pattern

Dim vi As String

Select Case version_no
  Case 7
     vi = "001010010011111000"
  Case 8
     vi = "001111011010000100"
   Case 9
     vi = "100110010101100100"
   Case 10
     vi = "110010110010010100"
   Case 11
     vi = "011011111101110100"
   Case 12
     vi = "010001101110001100"
   Case 13
     vi = "111000100001101100"
   Case 14
     vi = "101100000110011100"
   Case 15
     vi = "000101001001111100"
   Case 16
     vi = "000111101101000010"
   Case 17
     vi = "101110100010100010"
   Case 18
     vi = "111010000101010010"
   Case 19
     vi = "010011001010110010"
   Case 20
     vi = "011001011001001010"
   Case 21
     vi = "110000010110101010"
   Case 22
     vi = "100100110001011010"
   Case 23
     vi = "001101111110111010"
   Case 24
     vi = "001000110111000110"
   Case 25
     vi = "100001111000100110"
   Case 26
     vi = "110101011111010110"
   Case 27
     vi = "011100010000110110"
   Case 28
     vi = "010110000011001110"
   Case 29
     vi = "111111001100101110"
   Case 30
     vi = "101011101011011110"
   Case 31
     vi = "000010100100111110"
   Case 32
     vi = "101010111001000001"
   Case 33
     vi = "000011110110100001"
   Case 34
     vi = "010111010001010001"
   Case 35
     vi = "111110011110110001"
   Case 36
     vi = "110100001101001001"
   Case 37
     vi = "011101000010101001"
   Case 38
     vi = "001001100101011001"
   Case 39
     vi = "100000101010111001"
   Case 40
     vi = "100101100011000101"
 End Select

If version_no >= 7 Then
    teller = 1         'put version information to left of top right finding pattern
    For i = 1 To 6
      For j = 1 To 3
        If sf.Mid(vi, teller, 1) = "1" Then
          s5cij(i, qr_size - 11 + j) = 10
          teller = teller + 1
        Else
          s5cij(i, qr_size - 11 + j) = 9
          teller = teller + 1
        End If
      Next 
    Next 
    teller = 1        'put version information above bottom left finding pattern
    For j = 1 To 6
      For i = 1 To 3
        If sf.Mid(vi, teller, 1) = "1" Then
          s5cij(qr_size - 11 + i, j) = 10
           teller = teller + 1
        Else
          s5cij(qr_size - 11 + i, j) = 9
           teller = teller + 1
        End If
      Next 
    Next 
End If

'nullify horizontal mask patterns around alignment patterns
For j = 1 To 8
  s5cij(8, j) = 9
  s5cij(qr_size - 7, j) = 9
  s5cij(8, qr_size - 7 + j) = 9
Next 

'nullify vertical mask patterns around alignment patterns
For i = 1 To 8
  s5cij(i, 8) = 9
  s5cij(i, qr_size - 7) = 9
  s5cij(qr_size - 7 + i, 8) = 9
Next

'nullify horizontal mask patterns around where format info goes
For j = 1 To 9
  If s5cij(9, j) <> 10 Then s5cij(9, j) = 9 
  If s5cij(9, qr_size - 8 + j) <> 10 Then s5cij(9, qr_size - 8 + j) = 9
Next

'nullify vertical mask patterns around alignment patterns
For i = 1 To 8
  If s5cij(i, 9) <> 10 Then s5cij(i, 9) = 9
  If s5cij(qr_size - 8 + i, 9) <> 10 Then s5cij(qr_size - 8 + i, 9) = 9
Next

'now place the format information
Dim tipe As String
If error_level = "L" AND maskp = 0 Then tipe = "111011111000100"
If error_level = "L" AND maskp = 1 Then tipe = "111001011110011"
If error_level = "L" AND maskp = 2 Then tipe = "111110110101010"
If error_level = "L" AND maskp = 3 Then tipe = "111100010011101"
If error_level = "L" AND maskp = 4 Then tipe = "110011000101111"
If error_level = "L" AND maskp = 5 Then tipe = "110001100011000"
If error_level = "L" AND maskp = 6 Then tipe = "110110001000001"
If error_level = "L" AND maskp = 7 Then tipe = "110100101110110"
If error_level = "M" AND maskp = 0 Then tipe = "101010000010010"
If error_level = "M" AND maskp = 1 Then tipe = "101000100100101"
If error_level = "M" AND maskp = 2 Then tipe = "101111001111100"
If error_level = "M" AND maskp = 3 Then tipe = "101101101001011"
If error_level = "M" AND maskp = 4 Then tipe = "100010111111001"
If error_level = "M" AND maskp = 5 Then tipe = "100000011001110"
If error_level = "M" AND maskp = 6 Then tipe = "100111110010111"
If error_level = "M" AND maskp = 7 Then tipe = "100101010100000"
If error_level = "Q" AND maskp = 0 Then tipe = "011010101011111"
If error_level = "Q" AND maskp = 1 Then tipe = "011000001101000"
If error_level = "Q" AND maskp = 2 Then tipe = "011111100110001"
If error_level = "Q" AND maskp = 3 Then tipe = "011101000000110"
If error_level = "Q" AND maskp = 4 Then tipe = "010010010110100"
If error_level = "Q" AND maskp = 5 Then tipe = "010000110000011"
If error_level = "Q" AND maskp = 6 Then tipe = "010111011011010"
If error_level = "Q" AND maskp = 7 Then tipe = "010101111101101"
If error_level = "H" AND maskp = 0 Then tipe = "001011010001001"
If error_level = "H" AND maskp = 1 Then tipe = "001001110111110"
If error_level = "H" AND maskp = 2 Then tipe = "001110011100111"
If error_level = "H" AND maskp = 3 Then tipe = "001100111010000"
If error_level = "H" AND maskp = 4 Then tipe = "000011101100010"
If error_level = "H" AND maskp = 5 Then tipe = "000001001010101"
If error_level = "H" AND maskp = 6 Then tipe = "000110100001100"
If error_level = "H" AND maskp = 7 Then tipe = "000100000111011"

teller = 1    'place horizontal format info
For j = 1 To 15
  If sf.Mid(tipe, j, 1) = "1" Then
    s5cij(9, teller) = 10
  Else
    s5cij(9, teller) = 9
  End If
  teller = teller + 1
  If teller = 7 Then teller = teller + 1
  If teller = 9 Then teller = qr_size - 7   'was 7
Next 
teller = qr_size  'place vertical format info
For i = 1 To 15
  If sf.Mid(tipe, i, 1) = "1" Then
    s5cij(teller, 9) = 10
  Else
    s5cij(teller, 9) = 9
  End If
  teller = teller - 1
  If teller = qr_size - 7 Then teller = 9
  If teller = 7 Then teller = 6
Next

Dim k As Int
For k = 1 To 7
  co_x(k) = 0
  co_y(k) = 0
Next 

Select Case version_no
Case 2
  co_x(1) = ver2(57)
  co_x(2) = ver2(58)
  co_x(3) = ver2(59)
  co_x(4) = ver2(60)
  co_x(5) = ver2(61)
  co_x(6) = ver2(62)
  co_x(7) = ver2(63)
  co_y(1) = ver2(57)
  co_y(2) = ver2(58)
  co_y(3) = ver2(59)
  co_y(4) = ver2(60)
  co_y(5) = ver2(61)
  co_y(6) = ver2(62)
  co_y(7) = ver2(63)  
Case 3
  co_x(1) = ver3(57)
  co_x(2) = ver3(58)
  co_x(3) = ver3(59)
  co_x(4) = ver3(60)
  co_x(5) = ver3(61)
  co_x(6) = ver3(62)
  co_x(7) = ver3(63)
  co_y(1) = ver3(57)
  co_y(2) = ver3(58)
  co_y(3) = ver3(59)
  co_y(4) = ver3(60)
  co_y(5) = ver3(61)
  co_y(6) = ver3(62)
  co_y(7) = ver3(63)  
Case 4
  co_x(1) = ver4(57)
  co_x(2) = ver4(58)
  co_x(3) = ver4(59)
  co_x(4) = ver4(60)
  co_x(5) = ver4(61)
  co_x(6) = ver4(62)
  co_x(7) = ver4(63)
  co_y(1) = ver4(57)
  co_y(2) = ver4(58)
  co_y(3) = ver4(59)
  co_y(4) = ver4(60)
  co_y(5) = ver4(61)
  co_y(6) = ver4(62)
  co_y(7) = ver4(63)  
Case 5
  co_x(1) = ver5(57)
  co_x(2) = ver5(58)
  co_x(3) = ver5(59)
  co_x(4) = ver5(60)
  co_x(5) = ver5(61)
  co_x(6) = ver5(62)
  co_x(7) = ver5(63)
  co_y(1) = ver5(57)
  co_y(2) = ver5(58)
  co_y(3) = ver5(59)
  co_y(4) = ver5(60)
  co_y(5) = ver5(61)
  co_y(6) = ver5(62)
  co_y(7) = ver5(63)  
Case 6
  co_x(1) = ver6(57)
  co_x(2) = ver6(58)
  co_x(3) = ver6(59)
  co_x(4) = ver6(60)
  co_x(5) = ver6(61)
  co_x(6) = ver6(62)
  co_x(7) = ver6(63)
  co_y(1) = ver6(57)
  co_y(2) = ver6(58)
  co_y(3) = ver6(59)
  co_y(4) = ver6(60)
  co_y(5) = ver6(61)
  co_y(6) = ver6(62)
  co_y(7) = ver6(63)  
Case 7
  co_x(1) = ver7(57)
  co_x(2) = ver7(58)
  co_x(3) = ver7(59)
  co_x(4) = ver7(60)
  co_x(5) = ver7(61)
  co_x(6) = ver7(62)
  co_x(7) = ver7(63)
  co_y(1) = ver7(57)
  co_y(2) = ver7(58)
  co_y(3) = ver7(59)
  co_y(4) = ver7(60)
  co_y(5) = ver7(61)
  co_y(6) = ver7(62)
  co_y(7) = ver7(63)  
Case 8
  co_x(1) = ver8(57)
  co_x(2) = ver8(58)
  co_x(3) = ver8(59)
  co_x(4) = ver8(60)
  co_x(5) = ver8(61)
  co_x(6) = ver8(62)
  co_x(7) = ver8(63)
  co_y(1) = ver8(57)
  co_y(2) = ver8(58)
  co_y(3) = ver8(59)
  co_y(4) = ver8(60)
  co_y(5) = ver8(61)
  co_y(6) = ver8(62)
  co_y(7) = ver8(63)  
Case 9
  co_x(1) = ver9(57)
  co_x(2) = ver9(58)
  co_x(3) = ver9(59)
  co_x(4) = ver9(60)
  co_x(5) = ver9(61)
  co_x(6) = ver9(62)
  co_x(7) = ver9(63)
  co_y(1) = ver9(57)
  co_y(2) = ver9(58)
  co_y(3) = ver9(59)
  co_y(4) = ver9(60)
  co_y(5) = ver9(61)
  co_y(6) = ver9(62)
  co_y(7) = ver9(63)  
Case 10
  co_x(1) = ver10(57)
  co_x(2) = ver10(58)
  co_x(3) = ver10(59)
  co_x(4) = ver10(60)
  co_x(5) = ver10(61)
  co_x(6) = ver10(62)
  co_x(7) = ver10(63)
  co_y(1) = ver10(57)
  co_y(2) = ver10(58)
  co_y(3) = ver10(59)
  co_y(4) = ver10(60)
  co_y(5) = ver10(61)
  co_y(6) = ver10(62)
  co_y(7) = ver10(63)  
Case 11
  co_x(1) = ver11(57)
  co_x(2) = ver11(58)
  co_x(3) = ver11(59)
  co_x(4) = ver11(60)
  co_x(5) = ver11(61)
  co_x(6) = ver11(62)
  co_x(7) = ver11(63)
  co_y(1) = ver11(57)
  co_y(2) = ver11(58)
  co_y(3) = ver11(59)
  co_y(4) = ver11(60)
  co_y(5) = ver11(61)
  co_y(6) = ver11(62)
  co_y(7) = ver11(63)  
Case 12
  co_x(1) = ver12(57)
  co_x(2) = ver12(58)
  co_x(3) = ver12(59)
  co_x(4) = ver12(60)
  co_x(5) = ver12(61)
  co_x(6) = ver12(62)
  co_x(7) = ver12(63)
  co_y(1) = ver12(57)
  co_y(2) = ver12(58)
  co_y(3) = ver12(59)
  co_y(4) = ver12(60)
  co_y(5) = ver12(61)
  co_y(6) = ver12(62)
  co_y(7) = ver12(63)  
Case 13
  co_x(1) = ver13(57)
  co_x(2) = ver13(58)
  co_x(3) = ver13(59)
  co_x(4) = ver13(60)
  co_x(5) = ver13(61)
  co_x(6) = ver13(62)
  co_x(7) = ver13(63)
  co_y(1) = ver13(57)
  co_y(2) = ver13(58)
  co_y(3) = ver13(59)
  co_y(4) = ver13(60)
  co_y(5) = ver13(61)
  co_y(6) = ver13(62)
  co_y(7) = ver13(63)  
Case 14
  co_x(1) = ver14(57)
  co_x(2) = ver14(58)
  co_x(3) = ver14(59)
  co_x(4) = ver14(60)
  co_x(5) = ver14(61)
  co_x(6) = ver14(62)
  co_x(7) = ver14(63)
  co_y(1) = ver14(57)
  co_y(2) = ver14(58)
  co_y(3) = ver14(59)
  co_y(4) = ver14(60)
  co_y(5) = ver14(61)
  co_y(6) = ver14(62)
  co_y(7) = ver14(63)  
Case 15
  co_x(1) = ver15(57)
  co_x(2) = ver15(58)
  co_x(3) = ver15(59)
  co_x(4) = ver15(60)
  co_x(5) = ver15(61)
  co_x(6) = ver15(62)
  co_x(7) = ver15(63)
  co_y(1) = ver15(57)
  co_y(2) = ver15(58)
  co_y(3) = ver15(59)
  co_y(4) = ver15(60)
  co_y(5) = ver15(61)
  co_y(6) = ver15(62)
  co_y(7) = ver15(63)  
Case 16
  co_x(1) = ver16(57)
  co_x(2) = ver16(58)
  co_x(3) = ver16(59)
  co_x(4) = ver16(60)
  co_x(5) = ver16(61)
  co_x(6) = ver16(62)
  co_x(7) = ver16(63)
  co_y(1) = ver16(57)
  co_y(2) = ver16(58)
  co_y(3) = ver16(59)
  co_y(4) = ver16(60)
  co_y(5) = ver16(61)
  co_y(6) = ver16(62)
  co_y(7) = ver15(63)    
 Case 17
  co_x(1) = ver17(57)
  co_x(2) = ver17(58)
  co_x(3) = ver17(59)
  co_x(4) = ver17(60)
  co_x(5) = ver17(61)
  co_x(6) = ver17(62)
  co_x(7) = ver17(63)
  co_y(1) = ver17(57)
  co_y(2) = ver17(58)
  co_y(3) = ver17(59)
  co_y(4) = ver17(60)
  co_y(5) = ver17(61)
  co_y(6) = ver17(62)
  co_y(7) = ver17(63)   
Case 18
  co_x(1) = ver18(57)
  co_x(2) = ver18(58)
  co_x(3) = ver18(59)
  co_x(4) = ver18(60)
  co_x(5) = ver18(61)
  co_x(6) = ver18(62)
  co_x(7) = ver18(63)
  co_y(1) = ver18(57)
  co_y(2) = ver18(58)
  co_y(3) = ver18(59)
  co_y(4) = ver18(60)
  co_y(5) = ver18(61)
  co_y(6) = ver18(62)
  co_y(7) = ver18(63)    
Case 19
  co_x(1) = ver19(57)
  co_x(2) = ver19(58)
  co_x(3) = ver19(59)
  co_x(4) = ver19(60)
  co_x(5) = ver19(61)
  co_x(6) = ver19(62)
  co_x(7) = ver19(63)
  co_y(1) = ver19(57)
  co_y(2) = ver19(58)
  co_y(3) = ver19(59)
  co_y(4) = ver19(60)
  co_y(5) = ver19(61)
  co_y(6) = ver19(62)
  co_y(7) = ver19(63)    
 Case 20
  co_x(1) = ver20(57)
  co_x(2) = ver20(58)
  co_x(3) = ver20(59)
  co_x(4) = ver20(60)
  co_x(5) = ver20(61)
  co_x(6) = ver20(62)
  co_x(7) = ver20(63)
  co_y(1) = ver20(57)
  co_y(2) = ver20(58)
  co_y(3) = ver20(59)
  co_y(4) = ver20(60)
  co_y(5) = ver20(61)
  co_y(6) = ver20(62)
  co_y(7) = ver20(63)   
 Case 21
  co_x(1) = ver21(57)
  co_x(2) = ver21(58)
  co_x(3) = ver21(59)
  co_x(4) = ver21(60)
  co_x(5) = ver21(61)
  co_x(6) = ver21(62)
  co_x(7) = ver21(63)
  co_y(1) = ver21(57)
  co_y(2) = ver21(58)
  co_y(3) = ver21(59)
  co_y(4) = ver21(60)
  co_y(5) = ver21(61)
  co_y(6) = ver21(62)
  co_y(7) = ver21(63)   
 Case 22
  co_x(1) = ver22(57)
  co_x(2) = ver22(58)
  co_x(3) = ver22(59)
  co_x(4) = ver22(60)
  co_x(5) = ver22(61)
  co_x(6) = ver22(62)
  co_x(7) = ver22(63)
  co_y(1) = ver22(57)
  co_y(2) = ver22(58)
  co_y(3) = ver22(59)
  co_y(4) = ver22(60)
  co_y(5) = ver22(61)
  co_y(6) = ver22(62)
  co_y(7) = ver22(63)   
 Case 23
  co_x(1) = ver23(57)
  co_x(2) = ver23(58)
  co_x(3) = ver23(59)
  co_x(4) = ver23(60)
  co_x(5) = ver23(61)
  co_x(6) = ver23(62)
  co_x(7) = ver23(63)
  co_y(1) = ver23(57)
  co_y(2) = ver23(58)
  co_y(3) = ver23(59)
  co_y(4) = ver23(60)
  co_y(5) = ver23(61)
  co_y(6) = ver23(62)
  co_y(7) = ver23(63)   
 Case 24
  co_x(1) = ver24(57)
  co_x(2) = ver24(58)
  co_x(3) = ver24(59)
  co_x(4) = ver24(60)
  co_x(5) = ver24(61)
  co_x(6) = ver24(62)
  co_x(7) = ver24(63)
  co_y(1) = ver24(57)
  co_y(2) = ver24(58)
  co_y(3) = ver24(59)
  co_y(4) = ver24(60)
  co_y(5) = ver24(61)
  co_y(6) = ver24(62)
  co_y(7) = ver24(63)   
 Case 25
  co_x(1) = ver25(57)
  co_x(2) = ver25(58)
  co_x(3) = ver25(59)
  co_x(4) = ver25(60)
  co_x(5) = ver25(61)
  co_x(6) = ver25(62)
  co_x(7) = ver25(63)
  co_y(1) = ver25(57)
  co_y(2) = ver25(58)
  co_y(3) = ver25(59)
  co_y(4) = ver25(60)
  co_y(5) = ver25(61)
  co_y(6) = ver25(62)
  co_y(7) = ver25(63)   
 Case 26
  co_x(1) = ver26(57)
  co_x(2) = ver26(58)
  co_x(3) = ver26(59)
  co_x(4) = ver26(60)
  co_x(5) = ver26(61)
  co_x(6) = ver26(62)
  co_x(7) = ver26(63)
  co_y(1) = ver26(57)
  co_y(2) = ver26(58)
  co_y(3) = ver26(59)
  co_y(4) = ver26(60)
  co_y(5) = ver26(61)
  co_y(6) = ver26(62)
  co_y(7) = ver26(63)    
 Case 27
  co_x(1) = ver27(57)
  co_x(2) = ver27(58)
  co_x(3) = ver27(59)
  co_x(4) = ver27(60)
  co_x(5) = ver27(61)
  co_x(6) = ver27(62)
  co_x(7) = ver27(63)
  co_y(1) = ver27(57)
  co_y(2) = ver27(58)
  co_y(3) = ver27(59)
  co_y(4) = ver27(60)
  co_y(5) = ver27(61)
  co_y(6) = ver27(62)
  co_y(7) = ver27(63)    
 Case 28
  co_x(1) = ver28(57)
  co_x(2) = ver28(58)
  co_x(3) = ver28(59)
  co_x(4) = ver28(60)
  co_x(5) = ver28(61)
  co_x(6) = ver28(62)
  co_x(7) = ver28(63)
  co_y(1) = ver28(57)
  co_y(2) = ver28(58)
  co_y(3) = ver28(59)
  co_y(4) = ver28(60)
  co_y(5) = ver28(61)
  co_y(6) = ver28(62)
  co_y(7) = ver28(63)  
 Case 29
  co_x(1) = ver29(57)
  co_x(2) = ver29(58)
  co_x(3) = ver29(59)
  co_x(4) = ver29(60)
  co_x(5) = ver29(61)
  co_x(6) = ver29(62)
  co_x(7) = ver29(63)
  co_y(1) = ver29(57)
  co_y(2) = ver29(58)
  co_y(3) = ver29(59)
  co_y(4) = ver29(60)
  co_y(5) = ver29(61)
  co_y(6) = ver29(62)
  co_y(7) = ver29(63)    
 Case 30
  co_x(1) = ver30(57)
  co_x(2) = ver30(58)
  co_x(3) = ver30(59)
  co_x(4) = ver30(60)
  co_x(5) = ver30(61)
  co_x(6) = ver30(62)
  co_x(7) = ver30(63)
  co_y(1) = ver30(57)
  co_y(2) = ver30(58)
  co_y(3) = ver30(59)
  co_y(4) = ver30(60)
  co_y(5) = ver30(61)
  co_y(6) = ver30(62)
  co_y(7) = ver30(63)    
 Case 31
  co_x(1) = ver31(57)
  co_x(2) = ver31(58)
  co_x(3) = ver31(59)
  co_x(4) = ver31(60)
  co_x(5) = ver31(61)
  co_x(6) = ver31(62)
  co_x(7) = ver31(63)
  co_y(1) = ver31(57)
  co_y(2) = ver31(58)
  co_y(3) = ver31(59)
  co_y(4) = ver31(60)
  co_y(5) = ver31(61)
  co_y(6) = ver31(62)
  co_y(7) = ver31(63)    
 Case 32
  co_x(1) = ver32(57)
  co_x(2) = ver32(58)
  co_x(3) = ver32(59)
  co_x(4) = ver32(60)
  co_x(5) = ver32(61)
  co_x(6) = ver32(62)
  co_x(7) = ver32(63)
  co_y(1) = ver32(57)
  co_y(2) = ver32(58)
  co_y(3) = ver32(59)
  co_y(4) = ver32(60)
  co_y(5) = ver32(61)
  co_y(6) = ver32(62)
  co_y(7) = ver32(63)    
 Case 33
  co_x(1) = ver33(57)
  co_x(2) = ver33(58)
  co_x(3) = ver33(59)
  co_x(4) = ver33(60)
  co_x(5) = ver33(61)
  co_x(6) = ver33(62)
  co_x(7) = ver33(63)
  co_y(1) = ver33(57)
  co_y(2) = ver33(58)
  co_y(3) = ver33(59)
  co_y(4) = ver33(60)
  co_y(5) = ver33(61)
  co_y(6) = ver33(62)
  co_y(7) = ver33(63)    
 Case 34
  co_x(1) = ver34(57)
  co_x(2) = ver34(58)
  co_x(3) = ver34(59)
  co_x(4) = ver34(60)
  co_x(5) = ver34(61)
  co_x(6) = ver34(62)
  co_x(7) = ver34(63)
  co_y(1) = ver34(57)
  co_y(2) = ver34(58)
  co_y(3) = ver34(59)
  co_y(4) = ver34(60)
  co_y(5) = ver34(61)
  co_y(6) = ver34(62)
  co_y(7) = ver34(63)    
 Case 35
  co_x(1) = ver35(57)
  co_x(2) = ver35(58)
  co_x(3) = ver35(59)
  co_x(4) = ver35(60)
  co_x(5) = ver35(61)
  co_x(6) = ver35(62)
  co_x(7) = ver35(63)
  co_y(1) = ver35(57)
  co_y(2) = ver35(58)
  co_y(3) = ver35(59)
  co_y(4) = ver35(60)
  co_y(5) = ver35(61)
  co_y(6) = ver35(62)
  co_y(7) = ver35(63)    
 Case 36
  co_x(1) = ver36(57)
  co_x(2) = ver36(58)
  co_x(3) = ver36(59)
  co_x(4) = ver36(60)
  co_x(5) = ver36(61)
  co_x(6) = ver36(62)
  co_x(7) = ver36(63)
  co_y(1) = ver36(57)
  co_y(2) = ver36(58)
  co_y(3) = ver36(59)
  co_y(4) = ver36(60)
  co_y(5) = ver36(61)
  co_y(6) = ver36(62)
  co_y(7) = ver36(63)    
 Case 37
  co_x(1) = ver37(57)
  co_x(2) = ver37(58)
  co_x(3) = ver37(59)
  co_x(4) = ver37(60)
  co_x(5) = ver37(61)
  co_x(6) = ver37(62)
  co_x(7) = ver37(63)
  co_y(1) = ver37(57)
  co_y(2) = ver37(58)
  co_y(3) = ver37(59)
  co_y(4) = ver37(60)
  co_y(5) = ver37(61)
  co_y(6) = ver37(62)
  co_y(7) = ver37(63)    
 Case 38
  co_x(1) = ver38(57)
  co_x(2) = ver38(58)
  co_x(3) = ver38(59)
  co_x(4) = ver38(60)
  co_x(5) = ver38(61)
  co_x(6) = ver38(62)
  co_x(7) = ver38(63)
  co_y(1) = ver38(57)
  co_y(2) = ver38(58)
  co_y(3) = ver38(59)
  co_y(4) = ver38(60)
  co_y(5) = ver38(61)
  co_y(6) = ver38(62)
  co_y(7) = ver38(63)    
 Case 39
  co_x(1) = ver39(57)
  co_x(2) = ver39(58)
  co_x(3) = ver39(59)
  co_x(4) = ver39(60)
  co_x(5) = ver39(61)
  co_x(6) = ver39(62)
  co_x(7) = ver39(63)
  co_y(1) = ver39(57)
  co_y(2) = ver39(58)
  co_y(3) = ver39(59)
  co_y(4) = ver39(60)
  co_y(5) = ver39(61)
  co_y(6) = ver39(62)
  co_y(7) = ver39(63)    
 Case 40
  co_x(1) = ver40(57)
  co_x(2) = ver40(58)
  co_x(3) = ver40(59)
  co_x(4) = ver40(60)
  co_x(5) = ver40(61)
  co_x(6) = ver40(62)
  co_x(7) = ver40(63)
  co_y(1) = ver40(57)
  co_y(2) = ver40(58)
  co_y(3) = ver40(59)
  co_y(4) = ver40(60)
  co_y(5) = ver40(61)
  co_y(6) = ver40(62)
  co_y(7) = ver40(63)    
End Select

Dim flag,last As Int
'find coordinates of alignment patterns
If version_no > 1 Then  'only required for version 2 upwards - no alignment patterns in version 1
  flag = 0
  For k = 1 To 7
    If co_x(k) = 0 Then   'And flag = 0
      last = co_x(k - 1)
      Exit
    Else
      last = co_x(k)
    End If
  Next 
End If

teller = 1
For i = 1 To 7
  For j = 1 To 7
    apx(teller) = co_x(i)
    apy(teller) = co_y(j)
    teller = teller + 1
  Next 
Next 

'******************************************************************
Dim first As Int
first = 6    'this is always the first x or y cooridinate of one of the alignment patterns

'got last number other than zero stored in "last"

'put down the alignment patterns
flag = 0
For i = 1 To 49
  teller = 1
  If apx(i) = 0 OR apy(i) = 0 Then flag = 1
  If apx(i) = first AND apy(i) = first Then flag = 1   'exclude this combination - occupied by finding pattern i.e coordinate 6,6
  If apx(i) = first AND apy(i) = last Then flag = 1    'exclude this combination - occupied by finding pattern i.e coordinate last,6
  If apx(i) = last AND apy(i) = first Then flag = 1    'exclude this combination - occupied by finding pattern i.e coordinate 6,last
        If flag = 0 Then                               'a valid x,y position for an alignment pattern
          For k = apx(i) To apx(i) + 4
            For l = apy(i) To apy(i) + 4
              If ap(teller - 1) = 1 Then
                s5cij(k-1, l-1) = 10
                teller = teller + 1
              Else
                s5cij(k-1, l-1) = 9
                teller = teller + 1
              End If
            Next 
          Next
       Else
          flag = 0
       End If
Next

teken1

End Sub

Private Sub select_block

Select Case version_no

Case 1
    max_mod = ver1(12)
	qr_size = ver1(1)

    If error_level = "L" Then max_data_bits = ver1(20)
    If error_level = "M" Then max_data_bits = ver1(21)
    If error_level = "Q" Then max_data_bits = ver1(22)	
    If error_level = "H" Then max_data_bits = ver1(23)
    
	If error_level = "L" Then block1 = ver1(24)
    If error_level = "L" Then block2 = ver1(28)
    If error_level = "L" Then dw1 = ver1(26)
    If error_level = "L" Then dw2 = ver1(30)
	If error_level = "L" Then ew1 = ver1(27)
	If error_level = "L" Then ew2 = ver1(31)
	
	If error_level = "M" Then block1 = ver1(32)
    If error_level = "M" Then block2 = ver1(36)
    If error_level = "M" Then dw1 = ver1(34)
    If error_level = "M" Then dw2 = ver1(38)
	If error_level = "M" Then ew1 = ver1(35)
	If error_level = "M" Then ew2 = ver1(39)	
	
	If error_level = "Q" Then block1 = ver1(40)
    If error_level = "Q" Then block2 = ver1(44)
    If error_level = "Q" Then dw1 = ver1(42)
    If error_level = "Q" Then dw2 = ver1(46)
	If error_level = "Q" Then ew1 = ver1(43)
	If error_level = "Q" Then ew2 = ver1(47)		
	
	If error_level = "H" Then block1 = ver1(48)
    If error_level = "H" Then block2 = ver1(52)
    If error_level = "H" Then dw1 = ver1(50)
    If error_level = "H" Then dw2 = ver1(54)
	If error_level = "H" Then ew1 = ver1(51)
	If error_level = "H" Then ew2 = ver1(55)			
	
	If error_level = "L" Then tel2 = ver1(76)
	If error_level = "M" Then tel2 = ver1(77)
	If error_level = "Q" Then tel2 = ver1(78)
	If error_level = "H" Then tel2 = ver1(79)
Case 2
    max_mod = ver2(12)
	qr_size = ver2(1)
 	
    If error_level = "L" Then max_data_bits = ver2(20)
    If error_level = "M" Then max_data_bits = ver2(21)
    If error_level = "Q" Then max_data_bits = ver2(22)	
    If error_level = "H" Then max_data_bits = ver2(23)
	
	If error_level = "L" Then block1 = ver2(24)
    If error_level = "L" Then block2 = ver2(28)
    If error_level = "L" Then dw1 = ver2(26)
    If error_level = "L" Then dw2 = ver2(30)
	If error_level = "L" Then ew1 = ver2(27)
	If error_level = "L" Then ew2 = ver2(31)
	
	If error_level = "M" Then block1 = ver2(32)
    If error_level = "M" Then block2 = ver2(36)
    If error_level = "M" Then dw1 = ver2(34)
    If error_level = "M" Then dw2 = ver2(38)
	If error_level = "M" Then ew1 = ver2(35)
	If error_level = "M" Then ew2 = ver2(39)	
	
	If error_level = "Q" Then block1 = ver2(40)
    If error_level = "Q" Then block2 = ver2(44)
    If error_level = "Q" Then dw1 = ver2(42)
    If error_level = "Q" Then dw2 = ver2(46)
	If error_level = "Q" Then ew1 = ver2(43)
	If error_level = "Q" Then ew2 = ver2(47)		
	
	If error_level = "H" Then block1 = ver2(48)
    If error_level = "H" Then block2 = ver2(52)
    If error_level = "H" Then dw1 = ver2(50)
    If error_level = "H" Then dw2 = ver2(54)
	If error_level = "H" Then ew1 = ver2(51)
	If error_level = "H" Then ew2 = ver2(55)		
	
	If error_level = "L" Then tel2 = ver2(76)
	If error_level = "M" Then tel2 = ver2(77)
	If error_level = "Q" Then tel2 = ver2(78)
	If error_level = "H" Then tel2 = ver2(79)	
Case 3
    max_mod = ver3(12)
	qr_size = ver3(1)
	
    If error_level = "L" Then max_data_bits = ver3(20)
    If error_level = "M" Then max_data_bits = ver3(21)
    If error_level = "Q" Then max_data_bits = ver3(22)	
    If error_level = "H" Then max_data_bits = ver3(23)

	If error_level = "L" Then block1 = ver3(24)
    If error_level = "L" Then block2 = ver3(28)
    If error_level = "L" Then dw1 = ver3(26)
    If error_level = "L" Then dw2 = ver3(30)
	If error_level = "L" Then ew1 = ver3(27)
	If error_level = "L" Then ew2 = ver3(31)
	
	If error_level = "M" Then block1 = ver3(32)
    If error_level = "M" Then block2 = ver3(36)
    If error_level = "M" Then dw1 = ver3(34)
    If error_level = "M" Then dw2 = ver3(38)
	If error_level = "M" Then ew1 = ver3(35)
	If error_level = "M" Then ew2 = ver3(39)	
	
	If error_level = "Q" Then block1 = ver3(40)
    If error_level = "Q" Then block2 = ver3(44)
    If error_level = "Q" Then dw1 = ver3(42)
    If error_level = "Q" Then dw2 = ver3(46)
	If error_level = "Q" Then ew1 = ver3(43)
	If error_level = "Q" Then ew2 = ver3(47)		
	
	If error_level = "H" Then block1 = ver3(48)
    If error_level = "H" Then block2 = ver3(52)
    If error_level = "H" Then dw1 = ver3(50)
    If error_level = "H" Then dw2 = ver3(54)
	If error_level = "H" Then ew1 = ver3(51)
	If error_level = "H" Then ew2 = ver3(55)	

	If error_level = "L" Then tel2 = ver3(76)
	If error_level = "M" Then tel2 = ver3(77)
	If error_level = "Q" Then tel2 = ver3(78)
	If error_level = "H" Then tel2 = ver3(79)	
Case 4
    max_mod = ver4(12)
	qr_size = ver4(1)
 	
    If error_level = "L" Then max_data_bits = ver4(20)
    If error_level = "M" Then max_data_bits = ver4(21)
    If error_level = "Q" Then max_data_bits = ver4(22)	
    If error_level = "H" Then max_data_bits = ver4(23)
	
	If error_level = "L" Then block1 = ver4(24)
    If error_level = "L" Then block2 = ver4(28)
    If error_level = "L" Then dw1 = ver4(26)
    If error_level = "L" Then dw2 = ver4(30)
	If error_level = "L" Then ew1 = ver4(27)
	If error_level = "L" Then ew2 = ver4(31)
	
	If error_level = "M" Then block1 = ver4(32)
    If error_level = "M" Then block2 = ver4(36)
    If error_level = "M" Then dw1 = ver4(34)
    If error_level = "M" Then dw2 = ver4(38)
	If error_level = "M" Then ew1 = ver4(35)
	If error_level = "M" Then ew2 = ver4(39)	
	
	If error_level = "Q" Then block1 = ver4(40)
    If error_level = "Q" Then block2 = ver4(44)
    If error_level = "Q" Then dw1 = ver4(42)
    If error_level = "Q" Then dw2 = ver4(46)
	If error_level = "Q" Then ew1 = ver4(43)
	If error_level = "Q" Then ew2 = ver4(47)		
	
	If error_level = "H" Then block1 = ver4(48)
    If error_level = "H" Then block2 = ver4(52)
    If error_level = "H" Then dw1 = ver4(50)
    If error_level = "H" Then dw2 = ver4(54)
	If error_level = "H" Then ew1 = ver4(51)
	If error_level = "H" Then ew2 = ver4(55)
	
	If error_level = "L" Then tel2 = ver4(76)
	If error_level = "M" Then tel2 = ver4(77)
	If error_level = "Q" Then tel2 = ver4(78)
	If error_level = "H" Then tel2 = ver4(79)	
Case 5
    max_mod = ver5(12)
	qr_size = ver5(1)

    If error_level = "L" Then max_data_bits = ver5(20)
    If error_level = "M" Then max_data_bits = ver5(21)
    If error_level = "Q" Then max_data_bits = ver5(22)	
    If error_level = "H" Then max_data_bits = ver5(23)
	
	If error_level = "L" Then block1 = ver5(24)
    If error_level = "L" Then block2 = ver5(28)
    If error_level = "L" Then dw1 = ver5(26)
    If error_level = "L" Then dw2 = ver5(30)
	If error_level = "L" Then ew1 = ver5(27)
	If error_level = "L" Then ew2 = ver5(31)
	
	If error_level = "M" Then block1 = ver5(32)
    If error_level = "M" Then block2 = ver5(36)
    If error_level = "M" Then dw1 = ver5(34)
    If error_level = "M" Then dw2 = ver5(38)
	If error_level = "M" Then ew1 = ver5(35)
	If error_level = "M" Then ew2 = ver5(39)	
	
	If error_level = "Q" Then block1 = ver5(40)
    If error_level = "Q" Then block2 = ver5(44)
    If error_level = "Q" Then dw1 = ver5(42)
    If error_level = "Q" Then dw2 = ver5(46)
	If error_level = "Q" Then ew1 = ver5(43)
	If error_level = "Q" Then ew2 = ver5(47)		
	
	If error_level = "H" Then block1 = ver5(48)
    If error_level = "H" Then block2 = ver5(52)
    If error_level = "H" Then dw1 = ver5(50)
    If error_level = "H" Then dw2 = ver5(54)
	If error_level = "H" Then ew1 = ver5(51)
	If error_level = "H" Then ew2 = ver5(55)	
	
	If error_level = "L" Then tel2 = ver5(76)
	If error_level = "M" Then tel2 = ver5(77)
	If error_level = "Q" Then tel2 = ver5(78)
	If error_level = "H" Then tel2 = ver5(79)	
Case 6
    max_mod = ver6(12)
	qr_size = ver6(1)
	
    If error_level = "L" Then max_data_bits = ver6(20)
    If error_level = "M" Then max_data_bits = ver6(21)
    If error_level = "Q" Then max_data_bits = ver6(22)	
    If error_level = "H" Then max_data_bits = ver6(23)
	
	If error_level = "L" Then block1 = ver6(24)
    If error_level = "L" Then block2 = ver6(28)
    If error_level = "L" Then dw1 = ver6(26)
    If error_level = "L" Then dw2 = ver6(30)
	If error_level = "L" Then ew1 = ver6(27)
	If error_level = "L" Then ew2 = ver6(31)
	
	If error_level = "M" Then block1 = ver6(32)
    If error_level = "M" Then block2 = ver6(36)
    If error_level = "M" Then dw1 = ver6(34)
    If error_level = "M" Then dw2 = ver6(38)
	If error_level = "M" Then ew1 = ver6(35)
	If error_level = "M" Then ew2 = ver6(39)	
	
	If error_level = "Q" Then block1 = ver6(40)
    If error_level = "Q" Then block2 = ver6(44)
    If error_level = "Q" Then dw1 = ver6(42)
    If error_level = "Q" Then dw2 = ver6(46)
	If error_level = "Q" Then ew1 = ver6(43)
	If error_level = "Q" Then ew2 = ver6(47)		
	
	If error_level = "H" Then block1 = ver6(48)
    If error_level = "H" Then block2 = ver6(52)
    If error_level = "H" Then dw1 = ver6(50)
    If error_level = "H" Then dw2 = ver6(54)
	If error_level = "H" Then ew1 = ver6(51)
	If error_level = "H" Then ew2 = ver6(55)	
	
	If error_level = "L" Then tel2 = ver6(76)
	If error_level = "M" Then tel2 = ver6(77)
	If error_level = "Q" Then tel2 = ver6(78)
	If error_level = "H" Then tel2 = ver6(79)	
Case 7
    max_mod = ver7(12)
	qr_size = ver7(1)
	
    If error_level = "L" Then max_data_bits = ver7(20)
    If error_level = "M" Then max_data_bits = ver7(21)
    If error_level = "Q" Then max_data_bits = ver7(22)	
    If error_level = "H" Then max_data_bits = ver7(23)
	
	If error_level = "L" Then block1 = ver7(24)
    If error_level = "L" Then block2 = ver7(28)
    If error_level = "L" Then dw1 = ver7(26)
    If error_level = "L" Then dw2 = ver7(30)
	If error_level = "L" Then ew1 = ver7(27)
	If error_level = "L" Then ew2 = ver7(31)
	
	If error_level = "M" Then block1 = ver7(32)
    If error_level = "M" Then block2 = ver7(36)
    If error_level = "M" Then dw1 = ver7(34)
    If error_level = "M" Then dw2 = ver7(38)
	If error_level = "M" Then ew1 = ver7(35)
	If error_level = "M" Then ew2 = ver7(39)	
	
	If error_level = "Q" Then block1 = ver7(40)
    If error_level = "Q" Then block2 = ver7(44)
    If error_level = "Q" Then dw1 = ver7(42)
    If error_level = "Q" Then dw2 = ver7(46)
	If error_level = "Q" Then ew1 = ver7(43)
	If error_level = "Q" Then ew2 = ver7(47)		
	
	If error_level = "H" Then block1 = ver7(48)
    If error_level = "H" Then block2 = ver7(52)
    If error_level = "H" Then dw1 = ver7(50)
    If error_level = "H" Then dw2 = ver7(54)
	If error_level = "H" Then ew1 = ver7(51)
	If error_level = "H" Then ew2 = ver7(55)	
	
	If error_level = "L" Then tel2 = ver7(76)
	If error_level = "M" Then tel2 = ver7(77)
	If error_level = "Q" Then tel2 = ver7(78)
	If error_level = "H" Then tel2 = ver7(79)	
Case 8
    max_mod = ver8(12)
	qr_size = ver8(1)
	
    If error_level = "L" Then max_data_bits = ver8(20)
    If error_level = "M" Then max_data_bits = ver8(21)
    If error_level = "Q" Then max_data_bits = ver8(22)	
    If error_level = "H" Then max_data_bits = ver8(23)

	If error_level = "L" Then block1 = ver8(24)
    If error_level = "L" Then block2 = ver8(28)
    If error_level = "L" Then dw1 = ver8(26)
    If error_level = "L" Then dw2 = ver8(30)
	If error_level = "L" Then ew1 = ver8(27)
	If error_level = "L" Then ew2 = ver8(31)
	
	If error_level = "M" Then block1 = ver8(32)
    If error_level = "M" Then block2 = ver8(36)
    If error_level = "M" Then dw1 = ver8(34)
    If error_level = "M" Then dw2 = ver8(38)
	If error_level = "M" Then ew1 = ver8(35)
	If error_level = "M" Then ew2 = ver8(39)	
	
	If error_level = "Q" Then block1 = ver8(40)
    If error_level = "Q" Then block2 = ver8(44)
    If error_level = "Q" Then dw1 = ver8(42)
    If error_level = "Q" Then dw2 = ver8(46)
	If error_level = "Q" Then ew1 = ver8(43)
	If error_level = "Q" Then ew2 = ver8(47)		
	
	If error_level = "H" Then block1 = ver8(48)
    If error_level = "H" Then block2 = ver8(52)
    If error_level = "H" Then dw1 = ver8(50)
    If error_level = "H" Then dw2 = ver8(54)
	If error_level = "H" Then ew1 = ver8(51)
	If error_level = "H" Then ew2 = ver8(55)	
	
	If error_level = "L" Then tel2 = ver8(76)
	If error_level = "M" Then tel2 = ver8(77)
	If error_level = "Q" Then tel2 = ver8(78)
	If error_level = "H" Then tel2 = ver8(79)	
Case 9
    max_mod = ver9(12)
	qr_size = ver9(1)
	
    If error_level = "L" Then max_data_bits = ver9(20)
    If error_level = "M" Then max_data_bits = ver9(21)
    If error_level = "Q" Then max_data_bits = ver9(22)	
    If error_level = "H" Then max_data_bits = ver9(23)
	
	If error_level = "L" Then block1 = ver9(24)
    If error_level = "L" Then block2 = ver9(28)
    If error_level = "L" Then dw1 = ver9(26)
    If error_level = "L" Then dw2 = ver9(30)
	If error_level = "L" Then ew1 = ver9(27)
	If error_level = "L" Then ew2 = ver9(31)
	
	If error_level = "M" Then block1 = ver9(32)
    If error_level = "M" Then block2 = ver9(36)
    If error_level = "M" Then dw1 = ver9(34)
    If error_level = "M" Then dw2 = ver9(38)
	If error_level = "M" Then ew1 = ver9(35)
	If error_level = "M" Then ew2 = ver9(39)	
	
	If error_level = "Q" Then block1 = ver9(40)
    If error_level = "Q" Then block2 = ver9(44)
    If error_level = "Q" Then dw1 = ver9(42)
    If error_level = "Q" Then dw2 = ver9(46)
	If error_level = "Q" Then ew1 = ver9(43)
	If error_level = "Q" Then ew2 = ver9(47)		
	
	If error_level = "H" Then block1 = ver9(48)
    If error_level = "H" Then block2 = ver9(52)
    If error_level = "H" Then dw1 = ver9(50)
    If error_level = "H" Then dw2 = ver9(54)
	If error_level = "H" Then ew1 = ver9(51)
	If error_level = "H" Then ew2 = ver9(55)	
	
	If error_level = "L" Then tel2 = ver9(76)
	If error_level = "M" Then tel2 = ver9(77)
	If error_level = "Q" Then tel2 = ver9(78)
	If error_level = "H" Then tel2 = ver9(79)	
Case 10
    max_mod = ver10(12)
	qr_size = ver10(1)
	
    If error_level = "L" Then max_data_bits = ver10(20)
    If error_level = "M" Then max_data_bits = ver10(21)
    If error_level = "Q" Then max_data_bits = ver10(22)	
    If error_level = "H" Then max_data_bits = ver10(23)
	
	If error_level = "L" Then block1 = ver10(24)
    If error_level = "L" Then block2 = ver10(28)
    If error_level = "L" Then dw1 = ver10(26)
    If error_level = "L" Then dw2 = ver10(30)
	If error_level = "L" Then ew1 = ver10(27)
	If error_level = "L" Then ew2 = ver10(31)
	
	If error_level = "M" Then block1 = ver10(32)
    If error_level = "M" Then block2 = ver10(36)
    If error_level = "M" Then dw1 = ver10(34)
    If error_level = "M" Then dw2 = ver10(38)
	If error_level = "M" Then ew1 = ver10(35)
	If error_level = "M" Then ew2 = ver10(39)	
	
	If error_level = "Q" Then block1 = ver10(40)
    If error_level = "Q" Then block2 = ver10(44)
    If error_level = "Q" Then dw1 = ver10(42)
    If error_level = "Q" Then dw2 = ver10(46)
	If error_level = "Q" Then ew1 = ver10(43)
	If error_level = "Q" Then ew2 = ver10(47)		
	
	If error_level = "H" Then block1 = ver10(48)
    If error_level = "H" Then block2 = ver10(52)
    If error_level = "H" Then dw1 = ver10(50)
    If error_level = "H" Then dw2 = ver10(54)
	If error_level = "H" Then ew1 = ver10(51)
	If error_level = "H" Then ew2 = ver10(55)	
	
	If error_level = "L" Then tel2 = ver10(76)
	If error_level = "M" Then tel2 = ver10(77)
	If error_level = "Q" Then tel2 = ver10(78)
	If error_level = "H" Then tel2 = ver10(79)	
Case 11
    max_mod = ver11(12)
	qr_size = ver11(1)
	
    If error_level = "L" Then max_data_bits = ver11(20)
    If error_level = "M" Then max_data_bits = ver11(21)
    If error_level = "Q" Then max_data_bits = ver11(22)	
    If error_level = "H" Then max_data_bits = ver11(23)
	
	If error_level = "L" Then block1 = ver11(24)
    If error_level = "L" Then block2 = ver11(28)
    If error_level = "L" Then dw1 = ver11(26)
    If error_level = "L" Then dw2 = ver11(30)
	If error_level = "L" Then ew1 = ver11(27)
	If error_level = "L" Then ew2 = ver11(31)
	
	If error_level = "M" Then block1 = ver11(32)
    If error_level = "M" Then block2 = ver11(36)
    If error_level = "M" Then dw1 = ver11(34)
    If error_level = "M" Then dw2 = ver11(38)
	If error_level = "M" Then ew1 = ver11(35)
	If error_level = "M" Then ew2 = ver11(39)	
	
	If error_level = "Q" Then block1 = ver11(40)
    If error_level = "Q" Then block2 = ver11(44)
    If error_level = "Q" Then dw1 = ver11(42)
    If error_level = "Q" Then dw2 = ver11(46)
	If error_level = "Q" Then ew1 = ver11(43)
	If error_level = "Q" Then ew2 = ver11(47)		
	
	If error_level = "H" Then block1 = ver11(48)
    If error_level = "H" Then block2 = ver11(52)
    If error_level = "H" Then dw1 = ver11(50)
    If error_level = "H" Then dw2 = ver11(54)
	If error_level = "H" Then ew1 = ver11(51)
	If error_level = "H" Then ew2 = ver11(55)	
	
	If error_level = "L" Then tel2 = ver11(76)
	If error_level = "M" Then tel2 = ver11(77)
	If error_level = "Q" Then tel2 = ver11(78)
	If error_level = "H" Then tel2 = ver11(79)	
Case 12
    max_mod = ver12(12)
	qr_size = ver12(1)
	
    If error_level = "L" Then max_data_bits = ver12(20)
    If error_level = "M" Then max_data_bits = ver12(21)
    If error_level = "Q" Then max_data_bits = ver12(22)	
    If error_level = "H" Then max_data_bits = ver12(23)
	
	If error_level = "L" Then block1 = ver12(24)
    If error_level = "L" Then block2 = ver12(28)
    If error_level = "L" Then dw1 = ver12(26)
    If error_level = "L" Then dw2 = ver12(30)
	If error_level = "L" Then ew1 = ver12(27)
	If error_level = "L" Then ew2 = ver12(31)
	
	If error_level = "M" Then block1 = ver12(32)
    If error_level = "M" Then block2 = ver12(36)
    If error_level = "M" Then dw1 = ver12(34)
    If error_level = "M" Then dw2 = ver12(38)
	If error_level = "M" Then ew1 = ver12(35)
	If error_level = "M" Then ew2 = ver12(39)	
	
	If error_level = "Q" Then block1 = ver12(40)
    If error_level = "Q" Then block2 = ver12(44)
    If error_level = "Q" Then dw1 = ver12(42)
    If error_level = "Q" Then dw2 = ver12(46)
	If error_level = "Q" Then ew1 = ver12(43)
	If error_level = "Q" Then ew2 = ver12(47)		
	
	If error_level = "H" Then block1 = ver12(48)
    If error_level = "H" Then block2 = ver12(52)
    If error_level = "H" Then dw1 = ver12(50)
    If error_level = "H" Then dw2 = ver12(54)
	If error_level = "H" Then ew1 = ver12(51)
	If error_level = "H" Then ew2 = ver12(55)	
	
	If error_level = "L" Then tel2 = ver12(76)
	If error_level = "M" Then tel2 = ver12(77)
	If error_level = "Q" Then tel2 = ver12(78)
	If error_level = "H" Then tel2 = ver12(79)	
Case 13
    max_mod = ver13(12)
	qr_size = ver13(1)
	
    If error_level = "L" Then max_data_bits = ver13(20)
    If error_level = "M" Then max_data_bits = ver13(21)
    If error_level = "Q" Then max_data_bits = ver13(22)	
    If error_level = "H" Then max_data_bits = ver13(23)
	
	If error_level = "L" Then block1 = ver13(24)
    If error_level = "L" Then block2 = ver13(28)
    If error_level = "L" Then dw1 = ver13(26)
    If error_level = "L" Then dw2 = ver13(30)
	If error_level = "L" Then ew1 = ver13(27)
	If error_level = "L" Then ew2 = ver13(31)
	
	If error_level = "M" Then block1 = ver13(32)
    If error_level = "M" Then block2 = ver13(36)
    If error_level = "M" Then dw1 = ver13(34)
    If error_level = "M" Then dw2 = ver13(38)
	If error_level = "M" Then ew1 = ver13(35)
	If error_level = "M" Then ew2 = ver13(39)	
	
	If error_level = "Q" Then block1 = ver13(40)
    If error_level = "Q" Then block2 = ver13(44)
    If error_level = "Q" Then dw1 = ver13(42)
    If error_level = "Q" Then dw2 = ver13(46)
	If error_level = "Q" Then ew1 = ver13(43)
	If error_level = "Q" Then ew2 = ver13(47)		
	
	If error_level = "H" Then block1 = ver13(48)
    If error_level = "H" Then block2 = ver13(52)
    If error_level = "H" Then dw1 = ver13(50)
    If error_level = "H" Then dw2 = ver13(54)
	If error_level = "H" Then ew1 = ver13(51)
	If error_level = "H" Then ew2 = ver13(55)	
	
	If error_level = "L" Then tel2 = ver13(76)
	If error_level = "M" Then tel2 = ver13(77)
	If error_level = "Q" Then tel2 = ver13(78)
	If error_level = "H" Then tel2 = ver13(79)	
Case 14
    max_mod = ver14(12)
	qr_size = ver14(1)
	
    If error_level = "L" Then max_data_bits = ver14(20)
    If error_level = "M" Then max_data_bits = ver14(21)
    If error_level = "Q" Then max_data_bits = ver14(22)	
    If error_level = "H" Then max_data_bits = ver14(23)
	
	If error_level = "L" Then block1 = ver14(24)
    If error_level = "L" Then block2 = ver14(28)
    If error_level = "L" Then dw1 = ver14(26)
    If error_level = "L" Then dw2 = ver14(30)
	If error_level = "L" Then ew1 = ver14(27)
	If error_level = "L" Then ew2 = ver14(31)
	
	If error_level = "M" Then block1 = ver14(32)
    If error_level = "M" Then block2 = ver14(36)
    If error_level = "M" Then dw1 = ver14(34)
    If error_level = "M" Then dw2 = ver14(38)
	If error_level = "M" Then ew1 = ver14(35)
	If error_level = "M" Then ew2 = ver14(39)	
	
	If error_level = "Q" Then block1 = ver14(40)
    If error_level = "Q" Then block2 = ver14(44)
    If error_level = "Q" Then dw1 = ver14(42)
    If error_level = "Q" Then dw2 = ver14(46)
	If error_level = "Q" Then ew1 = ver14(43)
	If error_level = "Q" Then ew2 = ver14(47)		
	
	If error_level = "H" Then block1 = ver14(48)
    If error_level = "H" Then block2 = ver14(52)
    If error_level = "H" Then dw1 = ver14(50)
    If error_level = "H" Then dw2 = ver14(54)
	If error_level = "H" Then ew1 = ver14(51)
	If error_level = "H" Then ew2 = ver14(55)	
	
	If error_level = "L" Then tel2 = ver14(76)
	If error_level = "M" Then tel2 = ver14(77)
	If error_level = "Q" Then tel2 = ver14(78)
	If error_level = "H" Then tel2 = ver14(79)	
Case 15
    max_mod = ver15(12)
	qr_size = ver15(1)
	
    If error_level = "L" Then max_data_bits = ver15(20)
    If error_level = "M" Then max_data_bits = ver15(21)
    If error_level = "Q" Then max_data_bits = ver15(22)	
    If error_level = "H" Then max_data_bits = ver15(23)
	
	If error_level = "L" Then block1 = ver15(24)
    If error_level = "L" Then block2 = ver15(28)
    If error_level = "L" Then dw1 = ver15(26)
    If error_level = "L" Then dw2 = ver15(30)
	If error_level = "L" Then ew1 = ver15(27)
	If error_level = "L" Then ew2 = ver15(31)
	
	If error_level = "M" Then block1 = ver15(32)
    If error_level = "M" Then block2 = ver15(36)
    If error_level = "M" Then dw1 = ver15(34)
    If error_level = "M" Then dw2 = ver15(38)
	If error_level = "M" Then ew1 = ver15(35)
	If error_level = "M" Then ew2 = ver15(39)	
	
	If error_level = "Q" Then block1 = ver15(40)
    If error_level = "Q" Then block2 = ver15(44)
    If error_level = "Q" Then dw1 = ver15(42)
    If error_level = "Q" Then dw2 = ver15(46)
	If error_level = "Q" Then ew1 = ver15(43)
	If error_level = "Q" Then ew2 = ver15(47)		
	
	If error_level = "H" Then block1 = ver15(48)
    If error_level = "H" Then block2 = ver15(52)
    If error_level = "H" Then dw1 = ver15(50)
    If error_level = "H" Then dw2 = ver15(54)
	If error_level = "H" Then ew1 = ver15(51)
	If error_level = "H" Then ew2 = ver15(55)	
	
	If error_level = "L" Then tel2 = ver15(76)
	If error_level = "M" Then tel2 = ver15(77)
	If error_level = "Q" Then tel2 = ver15(78)
	If error_level = "H" Then tel2 = ver15(79)
Case 16
    max_mod = ver16(12)
	qr_size = ver16(1)
	
    If error_level = "L" Then max_data_bits = ver16(20)
    If error_level = "M" Then max_data_bits = ver16(21)
    If error_level = "Q" Then max_data_bits = ver16(22)	
    If error_level = "H" Then max_data_bits = ver16(23)
	
	If error_level = "L" Then block1 = ver16(24)
    If error_level = "L" Then block2 = ver16(28)
    If error_level = "L" Then dw1 = ver16(26)
    If error_level = "L" Then dw2 = ver16(30)
	If error_level = "L" Then ew1 = ver16(27)
	If error_level = "L" Then ew2 = ver16(31)
	
	If error_level = "M" Then block1 = ver16(32)
    If error_level = "M" Then block2 = ver16(36)
    If error_level = "M" Then dw1 = ver16(34)
    If error_level = "M" Then dw2 = ver16(38)
	If error_level = "M" Then ew1 = ver16(35)
	If error_level = "M" Then ew2 = ver16(39)	
	
	If error_level = "Q" Then block1 = ver16(40)
    If error_level = "Q" Then block2 = ver16(44)
    If error_level = "Q" Then dw1 = ver16(42)
    If error_level = "Q" Then dw2 = ver16(46)
	If error_level = "Q" Then ew1 = ver16(43)
	If error_level = "Q" Then ew2 = ver16(47)		
	
	If error_level = "H" Then block1 = ver16(48)
    If error_level = "H" Then block2 = ver16(52)
    If error_level = "H" Then dw1 = ver16(50)
    If error_level = "H" Then dw2 = ver16(54)
	If error_level = "H" Then ew1 = ver16(51)
	If error_level = "H" Then ew2 = ver16(55)	
	
	If error_level = "L" Then tel2 = ver16(76)
	If error_level = "M" Then tel2 = ver16(77)
	If error_level = "Q" Then tel2 = ver16(78)
	If error_level = "H" Then tel2 = ver16(79)	
Case 17
    max_mod = ver17(12)
	qr_size = ver17(1)
	
    If error_level = "L" Then max_data_bits = ver17(20)
    If error_level = "M" Then max_data_bits = ver17(21)
    If error_level = "Q" Then max_data_bits = ver17(22)	
    If error_level = "H" Then max_data_bits = ver17(23)
	
	If error_level = "L" Then block1 = ver17(24)
    If error_level = "L" Then block2 = ver17(28)
    If error_level = "L" Then dw1 = ver17(26)
    If error_level = "L" Then dw2 = ver17(30)
	If error_level = "L" Then ew1 = ver17(27)
	If error_level = "L" Then ew2 = ver17(31)
	
	If error_level = "M" Then block1 = ver17(32)
    If error_level = "M" Then block2 = ver17(36)
    If error_level = "M" Then dw1 = ver17(34)
    If error_level = "M" Then dw2 = ver17(38)
	If error_level = "M" Then ew1 = ver17(35)
	If error_level = "M" Then ew2 = ver17(39)	
	
	If error_level = "Q" Then block1 = ver17(40)
    If error_level = "Q" Then block2 = ver17(44)
    If error_level = "Q" Then dw1 = ver17(42)
    If error_level = "Q" Then dw2 = ver17(46)
	If error_level = "Q" Then ew1 = ver17(43)
	If error_level = "Q" Then ew2 = ver17(47)		
	
	If error_level = "H" Then block1 = ver17(48)
    If error_level = "H" Then block2 = ver17(52)
    If error_level = "H" Then dw1 = ver17(50)
    If error_level = "H" Then dw2 = ver17(54)
	If error_level = "H" Then ew1 = ver17(51)
	If error_level = "H" Then ew2 = ver17(55)	
	
	If error_level = "L" Then tel2 = ver17(76)
	If error_level = "M" Then tel2 = ver17(77)
	If error_level = "Q" Then tel2 = ver17(78)
	If error_level = "H" Then tel2 = ver17(79)		
Case 18
    max_mod = ver18(12)
	qr_size = ver18(1)
	
    If error_level = "L" Then max_data_bits = ver18(20)
    If error_level = "M" Then max_data_bits = ver18(21)
    If error_level = "Q" Then max_data_bits = ver18(22)	
    If error_level = "H" Then max_data_bits = ver18(23)
	
	If error_level = "L" Then block1 = ver18(24)
    If error_level = "L" Then block2 = ver18(28)
    If error_level = "L" Then dw1 = ver18(26)
    If error_level = "L" Then dw2 = ver18(30)
	If error_level = "L" Then ew1 = ver18(27)
	If error_level = "L" Then ew2 = ver18(31)
	
	If error_level = "M" Then block1 = ver18(32)
    If error_level = "M" Then block2 = ver18(36)
    If error_level = "M" Then dw1 = ver18(34)
    If error_level = "M" Then dw2 = ver18(38)
	If error_level = "M" Then ew1 = ver18(35)
	If error_level = "M" Then ew2 = ver18(39)	
	
	If error_level = "Q" Then block1 = ver18(40)
    If error_level = "Q" Then block2 = ver18(44)
    If error_level = "Q" Then dw1 = ver18(42)
    If error_level = "Q" Then dw2 = ver18(46)
	If error_level = "Q" Then ew1 = ver18(43)
	If error_level = "Q" Then ew2 = ver18(47)		
	
	If error_level = "H" Then block1 = ver18(48)
    If error_level = "H" Then block2 = ver18(52)
    If error_level = "H" Then dw1 = ver18(50)
    If error_level = "H" Then dw2 = ver18(54)
	If error_level = "H" Then ew1 = ver18(51)
	If error_level = "H" Then ew2 = ver18(55)	
	
	If error_level = "L" Then tel2 = ver18(76)
	If error_level = "M" Then tel2 = ver18(77)
	If error_level = "Q" Then tel2 = ver18(78)
	If error_level = "H" Then tel2 = ver18(79)		
Case 19
    max_mod = ver19(12)
	qr_size = ver19(1)
	
    If error_level = "L" Then max_data_bits = ver19(20)
    If error_level = "M" Then max_data_bits = ver19(21)
    If error_level = "Q" Then max_data_bits = ver19(22)	
    If error_level = "H" Then max_data_bits = ver19(23)
	
	If error_level = "L" Then block1 = ver19(24)
    If error_level = "L" Then block2 = ver19(28)
    If error_level = "L" Then dw1 = ver19(26)
    If error_level = "L" Then dw2 = ver19(30)
	If error_level = "L" Then ew1 = ver19(27)
	If error_level = "L" Then ew2 = ver19(31)
	
	If error_level = "M" Then block1 = ver19(32)
    If error_level = "M" Then block2 = ver19(36)
    If error_level = "M" Then dw1 = ver19(34)
    If error_level = "M" Then dw2 = ver19(38)
	If error_level = "M" Then ew1 = ver19(35)
	If error_level = "M" Then ew2 = ver19(39)	
	
	If error_level = "Q" Then block1 = ver19(40)
    If error_level = "Q" Then block2 = ver19(44)
    If error_level = "Q" Then dw1 = ver19(42)
    If error_level = "Q" Then dw2 = ver19(46)
	If error_level = "Q" Then ew1 = ver19(43)
	If error_level = "Q" Then ew2 = ver19(47)		
	
	If error_level = "H" Then block1 = ver19(48)
    If error_level = "H" Then block2 = ver19(52)
    If error_level = "H" Then dw1 = ver19(50)
    If error_level = "H" Then dw2 = ver19(54)
	If error_level = "H" Then ew1 = ver19(51)
	If error_level = "H" Then ew2 = ver19(55)	
	
	If error_level = "L" Then tel2 = ver19(76)
	If error_level = "M" Then tel2 = ver19(77)
	If error_level = "Q" Then tel2 = ver19(78)
	If error_level = "H" Then tel2 = ver19(79)		
Case 20
    max_mod = ver20(12)
	qr_size = ver20(1)
	
    If error_level = "L" Then max_data_bits = ver20(20)
    If error_level = "M" Then max_data_bits = ver20(21)
    If error_level = "Q" Then max_data_bits = ver20(22)	
    If error_level = "H" Then max_data_bits = ver20(23)
	
	If error_level = "L" Then block1 = ver20(24)
    If error_level = "L" Then block2 = ver20(28)
    If error_level = "L" Then dw1 = ver20(26)
    If error_level = "L" Then dw2 = ver20(30)
	If error_level = "L" Then ew1 = ver20(27)
	If error_level = "L" Then ew2 = ver20(31)
	
	If error_level = "M" Then block1 = ver20(32)
    If error_level = "M" Then block2 = ver20(36)
    If error_level = "M" Then dw1 = ver20(34)
    If error_level = "M" Then dw2 = ver20(38)
	If error_level = "M" Then ew1 = ver20(35)
	If error_level = "M" Then ew2 = ver20(39)	
	
	If error_level = "Q" Then block1 = ver20(40)
    If error_level = "Q" Then block2 = ver20(44)
    If error_level = "Q" Then dw1 = ver20(42)
    If error_level = "Q" Then dw2 = ver20(46)
	If error_level = "Q" Then ew1 = ver20(43)
	If error_level = "Q" Then ew2 = ver20(47)		
	
	If error_level = "H" Then block1 = ver20(48)
    If error_level = "H" Then block2 = ver20(52)
    If error_level = "H" Then dw1 = ver20(50)
    If error_level = "H" Then dw2 = ver20(54)
	If error_level = "H" Then ew1 = ver20(51)
	If error_level = "H" Then ew2 = ver20(55)	
	
	If error_level = "L" Then tel2 = ver20(76)
	If error_level = "M" Then tel2 = ver20(77)
	If error_level = "Q" Then tel2 = ver20(78)
	If error_level = "H" Then tel2 = ver20(79)	
Case 21
    max_mod = ver21(12)
	qr_size = ver21(1)
	
    If error_level = "L" Then max_data_bits = ver21(20)
    If error_level = "M" Then max_data_bits = ver21(21)
    If error_level = "Q" Then max_data_bits = ver21(22)	
    If error_level = "H" Then max_data_bits = ver21(23)
	
	If error_level = "L" Then block1 = ver21(24)
    If error_level = "L" Then block2 = ver21(28)
    If error_level = "L" Then dw1 = ver21(26)
    If error_level = "L" Then dw2 = ver21(30)
	If error_level = "L" Then ew1 = ver21(27)
	If error_level = "L" Then ew2 = ver21(31)
	
	If error_level = "M" Then block1 = ver21(32)
    If error_level = "M" Then block2 = ver21(36)
    If error_level = "M" Then dw1 = ver21(34)
    If error_level = "M" Then dw2 = ver21(38)
	If error_level = "M" Then ew1 = ver21(35)
	If error_level = "M" Then ew2 = ver21(39)	
	
	If error_level = "Q" Then block1 = ver21(40)
    If error_level = "Q" Then block2 = ver21(44)
    If error_level = "Q" Then dw1 = ver21(42)
    If error_level = "Q" Then dw2 = ver21(46)
	If error_level = "Q" Then ew1 = ver21(43)
	If error_level = "Q" Then ew2 = ver21(47)		
	
	If error_level = "H" Then block1 = ver21(48)
    If error_level = "H" Then block2 = ver21(52)
    If error_level = "H" Then dw1 = ver21(50)
    If error_level = "H" Then dw2 = ver21(54)
	If error_level = "H" Then ew1 = ver21(51)
	If error_level = "H" Then ew2 = ver21(55)	
	
	If error_level = "L" Then tel2 = ver21(76)
	If error_level = "M" Then tel2 = ver21(77)
	If error_level = "Q" Then tel2 = ver21(78)
	If error_level = "H" Then tel2 = ver21(79)	
Case 22
    max_mod = ver22(12)
	qr_size = ver22(1)
	
    If error_level = "L" Then max_data_bits = ver22(20)
    If error_level = "M" Then max_data_bits = ver22(21)
    If error_level = "Q" Then max_data_bits = ver22(22)	
    If error_level = "H" Then max_data_bits = ver22(23)
	
	If error_level = "L" Then block1 = ver22(24)
    If error_level = "L" Then block2 = ver22(28)
    If error_level = "L" Then dw1 = ver22(26)
    If error_level = "L" Then dw2 = ver22(30)
	If error_level = "L" Then ew1 = ver22(27)
	If error_level = "L" Then ew2 = ver22(31)
	
	If error_level = "M" Then block1 = ver22(32)
    If error_level = "M" Then block2 = ver22(36)
    If error_level = "M" Then dw1 = ver22(34)
    If error_level = "M" Then dw2 = ver22(38)
	If error_level = "M" Then ew1 = ver22(35)
	If error_level = "M" Then ew2 = ver22(39)	
	
	If error_level = "Q" Then block1 = ver22(40)
    If error_level = "Q" Then block2 = ver22(44)
    If error_level = "Q" Then dw1 = ver22(42)
    If error_level = "Q" Then dw2 = ver22(46)
	If error_level = "Q" Then ew1 = ver22(43)
	If error_level = "Q" Then ew2 = ver22(47)		
	
	If error_level = "H" Then block1 = ver22(48)
    If error_level = "H" Then block2 = ver22(52)
    If error_level = "H" Then dw1 = ver22(50)
    If error_level = "H" Then dw2 = ver22(54)
	If error_level = "H" Then ew1 = ver22(51)
	If error_level = "H" Then ew2 = ver22(55)	
	
	If error_level = "L" Then tel2 = ver22(76)
	If error_level = "M" Then tel2 = ver22(77)
	If error_level = "Q" Then tel2 = ver22(78)
	If error_level = "H" Then tel2 = ver22(79)	
Case 23
    max_mod = ver23(12)
	qr_size = ver23(1)
	
    If error_level = "L" Then max_data_bits = ver23(20)
    If error_level = "M" Then max_data_bits = ver23(21)
    If error_level = "Q" Then max_data_bits = ver23(22)	
    If error_level = "H" Then max_data_bits = ver23(23)
	
	If error_level = "L" Then block1 = ver23(24)
    If error_level = "L" Then block2 = ver23(28)
    If error_level = "L" Then dw1 = ver23(26)
    If error_level = "L" Then dw2 = ver23(30)
	If error_level = "L" Then ew1 = ver23(27)
	If error_level = "L" Then ew2 = ver23(31)
	
	If error_level = "M" Then block1 = ver23(32)
    If error_level = "M" Then block2 = ver23(36)
    If error_level = "M" Then dw1 = ver23(34)
    If error_level = "M" Then dw2 = ver23(38)
	If error_level = "M" Then ew1 = ver23(35)
	If error_level = "M" Then ew2 = ver23(39)	
	
	If error_level = "Q" Then block1 = ver23(40)
    If error_level = "Q" Then block2 = ver23(44)
    If error_level = "Q" Then dw1 = ver23(42)
    If error_level = "Q" Then dw2 = ver23(46)
	If error_level = "Q" Then ew1 = ver23(43)
	If error_level = "Q" Then ew2 = ver23(47)		
	
	If error_level = "H" Then block1 = ver23(48)
    If error_level = "H" Then block2 = ver23(52)
    If error_level = "H" Then dw1 = ver23(50)
    If error_level = "H" Then dw2 = ver23(54)
	If error_level = "H" Then ew1 = ver23(51)
	If error_level = "H" Then ew2 = ver23(55)	
	
	If error_level = "L" Then tel2 = ver23(76)
	If error_level = "M" Then tel2 = ver23(77)
	If error_level = "Q" Then tel2 = ver23(78)
	If error_level = "H" Then tel2 = ver23(79)	
Case 24
    max_mod = ver24(12)
	qr_size = ver24(1)
	
    If error_level = "L" Then max_data_bits = ver24(20)
    If error_level = "M" Then max_data_bits = ver24(21)
    If error_level = "Q" Then max_data_bits = ver24(22)	
    If error_level = "H" Then max_data_bits = ver24(23)
	
	If error_level = "L" Then block1 = ver24(24)
    If error_level = "L" Then block2 = ver24(28)
    If error_level = "L" Then dw1 = ver24(26)
    If error_level = "L" Then dw2 = ver24(30)
	If error_level = "L" Then ew1 = ver24(27)
	If error_level = "L" Then ew2 = ver24(31)
	
	If error_level = "M" Then block1 = ver24(32)
    If error_level = "M" Then block2 = ver24(36)
    If error_level = "M" Then dw1 = ver24(34)
    If error_level = "M" Then dw2 = ver24(38)
	If error_level = "M" Then ew1 = ver24(35)
	If error_level = "M" Then ew2 = ver24(39)	
	
	If error_level = "Q" Then block1 = ver24(40)
    If error_level = "Q" Then block2 = ver24(44)
    If error_level = "Q" Then dw1 = ver24(42)
    If error_level = "Q" Then dw2 = ver24(46)
	If error_level = "Q" Then ew1 = ver24(43)
	If error_level = "Q" Then ew2 = ver24(47)		
	
	If error_level = "H" Then block1 = ver24(48)
    If error_level = "H" Then block2 = ver24(52)
    If error_level = "H" Then dw1 = ver24(50)
    If error_level = "H" Then dw2 = ver24(54)
	If error_level = "H" Then ew1 = ver24(51)
	If error_level = "H" Then ew2 = ver24(55)	
	
	If error_level = "L" Then tel2 = ver24(76)
	If error_level = "M" Then tel2 = ver24(77)
	If error_level = "Q" Then tel2 = ver24(78)
	If error_level = "H" Then tel2 = ver24(79)	
Case 25
    max_mod = ver25(12)
	qr_size = ver25(1)
	
    If error_level = "L" Then max_data_bits = ver25(20)
    If error_level = "M" Then max_data_bits = ver25(21)
    If error_level = "Q" Then max_data_bits = ver25(22)	
    If error_level = "H" Then max_data_bits = ver25(23)
	
	If error_level = "L" Then block1 = ver25(24)
    If error_level = "L" Then block2 = ver25(28)
    If error_level = "L" Then dw1 = ver25(26)
    If error_level = "L" Then dw2 = ver25(30)
	If error_level = "L" Then ew1 = ver25(27)
	If error_level = "L" Then ew2 = ver25(31)
	
	If error_level = "M" Then block1 = ver25(32)
    If error_level = "M" Then block2 = ver25(36)
    If error_level = "M" Then dw1 = ver25(34)
    If error_level = "M" Then dw2 = ver25(38)
	If error_level = "M" Then ew1 = ver25(35)
	If error_level = "M" Then ew2 = ver25(39)	
	
	If error_level = "Q" Then block1 = ver25(40)
    If error_level = "Q" Then block2 = ver25(44)
    If error_level = "Q" Then dw1 = ver25(42)
    If error_level = "Q" Then dw2 = ver25(46)
	If error_level = "Q" Then ew1 = ver25(43)
	If error_level = "Q" Then ew2 = ver25(47)		
	
	If error_level = "H" Then block1 = ver25(48)
    If error_level = "H" Then block2 = ver25(52)
    If error_level = "H" Then dw1 = ver25(50)
    If error_level = "H" Then dw2 = ver25(54)
	If error_level = "H" Then ew1 = ver25(51)
	If error_level = "H" Then ew2 = ver25(55)	
	
	If error_level = "L" Then tel2 = ver25(76)
	If error_level = "M" Then tel2 = ver25(77)
	If error_level = "Q" Then tel2 = ver25(78)
	If error_level = "H" Then tel2 = ver25(79)	
Case 26
    max_mod = ver26(12)
	qr_size = ver26(1)
	
    If error_level = "L" Then max_data_bits = ver26(20)
    If error_level = "M" Then max_data_bits = ver26(21)
    If error_level = "Q" Then max_data_bits = ver26(22)	
    If error_level = "H" Then max_data_bits = ver26(23)
	
	If error_level = "L" Then block1 = ver26(24)
    If error_level = "L" Then block2 = ver26(28)
    If error_level = "L" Then dw1 = ver26(26)
    If error_level = "L" Then dw2 = ver26(30)
	If error_level = "L" Then ew1 = ver26(27)
	If error_level = "L" Then ew2 = ver26(31)
	
	If error_level = "M" Then block1 = ver26(32)
    If error_level = "M" Then block2 = ver26(36)
    If error_level = "M" Then dw1 = ver26(34)
    If error_level = "M" Then dw2 = ver26(38)
	If error_level = "M" Then ew1 = ver26(35)
	If error_level = "M" Then ew2 = ver26(39)	
	
	If error_level = "Q" Then block1 = ver26(40)
    If error_level = "Q" Then block2 = ver26(44)
    If error_level = "Q" Then dw1 = ver26(42)
    If error_level = "Q" Then dw2 = ver26(46)
	If error_level = "Q" Then ew1 = ver26(43)
	If error_level = "Q" Then ew2 = ver26(47)		
	
	If error_level = "H" Then block1 = ver26(48)
    If error_level = "H" Then block2 = ver26(52)
    If error_level = "H" Then dw1 = ver26(50)
    If error_level = "H" Then dw2 = ver26(54)
	If error_level = "H" Then ew1 = ver26(51)
	If error_level = "H" Then ew2 = ver26(55)	
	
	If error_level = "L" Then tel2 = ver26(76)
	If error_level = "M" Then tel2 = ver26(77)
	If error_level = "Q" Then tel2 = ver26(78)
	If error_level = "H" Then tel2 = ver26(79)	
Case 27
    max_mod = ver27(12)
	qr_size = ver27(1)
 	
    If error_level = "L" Then max_data_bits = ver27(20)
    If error_level = "M" Then max_data_bits = ver27(21)
    If error_level = "Q" Then max_data_bits = ver27(22)	
    If error_level = "H" Then max_data_bits = ver27(23)
	
	If error_level = "L" Then block1 = ver27(24)
    If error_level = "L" Then block2 = ver27(28)
    If error_level = "L" Then dw1 = ver27(26)
    If error_level = "L" Then dw2 = ver27(30)
	If error_level = "L" Then ew1 = ver27(27)
	If error_level = "L" Then ew2 = ver27(31)
	
	If error_level = "M" Then block1 = ver27(32)
    If error_level = "M" Then block2 = ver27(36)
    If error_level = "M" Then dw1 = ver27(34)
    If error_level = "M" Then dw2 = ver27(38)
	If error_level = "M" Then ew1 = ver27(35)
	If error_level = "M" Then ew2 = ver27(39)	
	
	If error_level = "Q" Then block1 = ver27(40)
    If error_level = "Q" Then block2 = ver27(44)
    If error_level = "Q" Then dw1 = ver27(42)
    If error_level = "Q" Then dw2 = ver27(46)
	If error_level = "Q" Then ew1 = ver27(43)
	If error_level = "Q" Then ew2 = ver27(47)		
	
	If error_level = "H" Then block1 = ver27(48)
    If error_level = "H" Then block2 = ver27(52)
    If error_level = "H" Then dw1 = ver27(50)
    If error_level = "H" Then dw2 = ver27(54)
	If error_level = "H" Then ew1 = ver27(51)
	If error_level = "H" Then ew2 = ver27(55)	
	
	If error_level = "L" Then tel2 = ver27(76)
	If error_level = "M" Then tel2 = ver27(77)
	If error_level = "Q" Then tel2 = ver27(78)
	If error_level = "H" Then tel2 = ver27(79)	
Case 28
    max_mod = ver28(12)
	qr_size = ver28(1)
 	
    If error_level = "L" Then max_data_bits = ver28(20)
    If error_level = "M" Then max_data_bits = ver28(21)
    If error_level = "Q" Then max_data_bits = ver28(22)	
    If error_level = "H" Then max_data_bits = ver28(23)
	
	If error_level = "L" Then block1 = ver28(24)
    If error_level = "L" Then block2 = ver28(28)
    If error_level = "L" Then dw1 = ver28(26)
    If error_level = "L" Then dw2 = ver28(30)
	If error_level = "L" Then ew1 = ver28(27)
	If error_level = "L" Then ew2 = ver28(31)
	
	If error_level = "M" Then block1 = ver28(32)
    If error_level = "M" Then block2 = ver28(36)
    If error_level = "M" Then dw1 = ver28(34)
    If error_level = "M" Then dw2 = ver28(38)
	If error_level = "M" Then ew1 = ver28(35)
	If error_level = "M" Then ew2 = ver28(39)	
	
	If error_level = "Q" Then block1 = ver28(40)
    If error_level = "Q" Then block2 = ver28(44)
    If error_level = "Q" Then dw1 = ver28(42)
    If error_level = "Q" Then dw2 = ver28(46)
	If error_level = "Q" Then ew1 = ver28(43)
	If error_level = "Q" Then ew2 = ver28(47)		
	
	If error_level = "H" Then block1 = ver28(48)
    If error_level = "H" Then block2 = ver28(52)
    If error_level = "H" Then dw1 = ver28(50)
    If error_level = "H" Then dw2 = ver28(54)
	If error_level = "H" Then ew1 = ver28(51)
	If error_level = "H" Then ew2 = ver28(55)	
	
	If error_level = "L" Then tel2 = ver28(76)
	If error_level = "M" Then tel2 = ver28(77)
	If error_level = "Q" Then tel2 = ver28(78)
	If error_level = "H" Then tel2 = ver28(79)	
Case 29
    max_mod = ver29(12)
	qr_size = ver29(1)
	
    If error_level = "L" Then max_data_bits = ver29(20)
    If error_level = "M" Then max_data_bits = ver29(21)
    If error_level = "Q" Then max_data_bits = ver29(22)	
    If error_level = "H" Then max_data_bits = ver29(23)
	
	If error_level = "L" Then block1 = ver29(24)
    If error_level = "L" Then block2 = ver29(28)
    If error_level = "L" Then dw1 = ver29(26)
    If error_level = "L" Then dw2 = ver29(30)
	If error_level = "L" Then ew1 = ver29(27)
	If error_level = "L" Then ew2 = ver29(31)
	
	If error_level = "M" Then block1 = ver29(32)
    If error_level = "M" Then block2 = ver29(36)
    If error_level = "M" Then dw1 = ver29(34)
    If error_level = "M" Then dw2 = ver29(38)
	If error_level = "M" Then ew1 = ver29(35)
	If error_level = "M" Then ew2 = ver29(39)	
	
	If error_level = "Q" Then block1 = ver29(40)
    If error_level = "Q" Then block2 = ver29(44)
    If error_level = "Q" Then dw1 = ver29(42)
    If error_level = "Q" Then dw2 = ver29(46)
	If error_level = "Q" Then ew1 = ver29(43)
	If error_level = "Q" Then ew2 = ver29(47)		
	
	If error_level = "H" Then block1 = ver29(48)
    If error_level = "H" Then block2 = ver29(52)
    If error_level = "H" Then dw1 = ver29(50)
    If error_level = "H" Then dw2 = ver29(54)
	If error_level = "H" Then ew1 = ver29(51)
	If error_level = "H" Then ew2 = ver29(55)	
	
	If error_level = "L" Then tel2 = ver29(76)
	If error_level = "M" Then tel2 = ver29(77)
	If error_level = "Q" Then tel2 = ver29(78)
	If error_level = "H" Then tel2 = ver29(79)	
Case 30
    max_mod = ver30(12)
	qr_size = ver30(1)
 	
    If error_level = "L" Then max_data_bits = ver30(20)
    If error_level = "M" Then max_data_bits = ver30(21)
    If error_level = "Q" Then max_data_bits = ver30(22)	
    If error_level = "H" Then max_data_bits = ver30(23)
	
	If error_level = "L" Then block1 = ver30(24)
    If error_level = "L" Then block2 = ver30(28)
    If error_level = "L" Then dw1 = ver30(26)
    If error_level = "L" Then dw2 = ver30(30)
	If error_level = "L" Then ew1 = ver30(27)
	If error_level = "L" Then ew2 = ver30(31)
	
	If error_level = "M" Then block1 = ver30(32)
    If error_level = "M" Then block2 = ver30(36)
    If error_level = "M" Then dw1 = ver30(34)
    If error_level = "M" Then dw2 = ver30(38)
	If error_level = "M" Then ew1 = ver30(35)
	If error_level = "M" Then ew2 = ver30(39)	
	
	If error_level = "Q" Then block1 = ver30(40)
    If error_level = "Q" Then block2 = ver30(44)
    If error_level = "Q" Then dw1 = ver30(42)
    If error_level = "Q" Then dw2 = ver30(46)
	If error_level = "Q" Then ew1 = ver30(43)
	If error_level = "Q" Then ew2 = ver30(47)		
	
	If error_level = "H" Then block1 = ver30(48)
    If error_level = "H" Then block2 = ver30(52)
    If error_level = "H" Then dw1 = ver30(50)
    If error_level = "H" Then dw2 = ver30(54)
	If error_level = "H" Then ew1 = ver30(51)
	If error_level = "H" Then ew2 = ver30(55)	
	
	If error_level = "L" Then tel2 = ver30(76)
	If error_level = "M" Then tel2 = ver30(77)
	If error_level = "Q" Then tel2 = ver30(78)
	If error_level = "H" Then tel2 = ver30(79)	
Case 31
    max_mod = ver31(12)
	qr_size = ver31(1)
	
    If error_level = "L" Then max_data_bits = ver31(20)
    If error_level = "M" Then max_data_bits = ver31(21)
    If error_level = "Q" Then max_data_bits = ver31(22)	
    If error_level = "H" Then max_data_bits = ver31(23)
	
	If error_level = "L" Then block1 = ver31(24)
    If error_level = "L" Then block2 = ver31(28)
    If error_level = "L" Then dw1 = ver31(26)
    If error_level = "L" Then dw2 = ver31(30)
	If error_level = "L" Then ew1 = ver31(27)
	If error_level = "L" Then ew2 = ver31(31)
	
	If error_level = "M" Then block1 = ver31(32)
    If error_level = "M" Then block2 = ver31(36)
    If error_level = "M" Then dw1 = ver31(34)
    If error_level = "M" Then dw2 = ver31(38)
	If error_level = "M" Then ew1 = ver31(35)
	If error_level = "M" Then ew2 = ver31(39)	
	
	If error_level = "Q" Then block1 = ver31(40)
    If error_level = "Q" Then block2 = ver31(44)
    If error_level = "Q" Then dw1 = ver31(42)
    If error_level = "Q" Then dw2 = ver31(46)
	If error_level = "Q" Then ew1 = ver31(43)
	If error_level = "Q" Then ew2 = ver31(47)		
	
	If error_level = "H" Then block1 = ver31(48)
    If error_level = "H" Then block2 = ver31(52)
    If error_level = "H" Then dw1 = ver31(50)
    If error_level = "H" Then dw2 = ver31(54)
	If error_level = "H" Then ew1 = ver31(51)
	If error_level = "H" Then ew2 = ver31(55)	
	
	If error_level = "L" Then tel2 = ver31(76)
	If error_level = "M" Then tel2 = ver31(77)
	If error_level = "Q" Then tel2 = ver31(78)
	If error_level = "H" Then tel2 = ver31(79)		
Case 32
    max_mod = ver32(12)
	qr_size = ver32(1)
	
    If error_level = "L" Then max_data_bits = ver32(20)
    If error_level = "M" Then max_data_bits = ver32(21)
    If error_level = "Q" Then max_data_bits = ver32(22)	
    If error_level = "H" Then max_data_bits = ver32(23)
	
	If error_level = "L" Then block1 = ver32(24)
    If error_level = "L" Then block2 = ver32(28)
    If error_level = "L" Then dw1 = ver32(26)
    If error_level = "L" Then dw2 = ver32(30)
	If error_level = "L" Then ew1 = ver32(27)
	If error_level = "L" Then ew2 = ver32(31)
	
	If error_level = "M" Then block1 = ver32(32)
    If error_level = "M" Then block2 = ver32(36)
    If error_level = "M" Then dw1 = ver32(34)
    If error_level = "M" Then dw2 = ver32(38)
	If error_level = "M" Then ew1 = ver32(35)
	If error_level = "M" Then ew2 = ver32(39)	
	
	If error_level = "Q" Then block1 = ver32(40)
    If error_level = "Q" Then block2 = ver32(44)
    If error_level = "Q" Then dw1 = ver32(42)
    If error_level = "Q" Then dw2 = ver32(46)
	If error_level = "Q" Then ew1 = ver32(43)
	If error_level = "Q" Then ew2 = ver32(47)		
	
	If error_level = "H" Then block1 = ver32(48)
    If error_level = "H" Then block2 = ver32(52)
    If error_level = "H" Then dw1 = ver32(50)
    If error_level = "H" Then dw2 = ver32(54)
	If error_level = "H" Then ew1 = ver32(51)
	If error_level = "H" Then ew2 = ver32(55)	
	
	If error_level = "L" Then tel2 = ver32(76)
	If error_level = "M" Then tel2 = ver32(77)
	If error_level = "Q" Then tel2 = ver32(78)
	If error_level = "H" Then tel2 = ver32(79)		
Case 33
    max_mod = ver33(12)
	qr_size = ver33(1)
	
    If error_level = "L" Then max_data_bits = ver33(20)
    If error_level = "M" Then max_data_bits = ver33(21)
    If error_level = "Q" Then max_data_bits = ver33(22)	
    If error_level = "H" Then max_data_bits = ver33(23)
	
	If error_level = "L" Then block1 = ver33(24)
    If error_level = "L" Then block2 = ver33(28)
    If error_level = "L" Then dw1 = ver33(26)
    If error_level = "L" Then dw2 = ver33(30)
	If error_level = "L" Then ew1 = ver33(27)
	If error_level = "L" Then ew2 = ver33(31)
	
	If error_level = "M" Then block1 = ver33(32)
    If error_level = "M" Then block2 = ver33(36)
    If error_level = "M" Then dw1 = ver33(34)
    If error_level = "M" Then dw2 = ver33(38)
	If error_level = "M" Then ew1 = ver33(35)
	If error_level = "M" Then ew2 = ver33(39)	
	
	If error_level = "Q" Then block1 = ver33(40)
    If error_level = "Q" Then block2 = ver33(44)
    If error_level = "Q" Then dw1 = ver33(42)
    If error_level = "Q" Then dw2 = ver33(46)
	If error_level = "Q" Then ew1 = ver33(43)
	If error_level = "Q" Then ew2 = ver33(47)		
	
	If error_level = "H" Then block1 = ver33(48)
    If error_level = "H" Then block2 = ver33(52)
    If error_level = "H" Then dw1 = ver33(50)
    If error_level = "H" Then dw2 = ver33(54)
	If error_level = "H" Then ew1 = ver33(51)
	If error_level = "H" Then ew2 = ver33(55)	
	
	If error_level = "L" Then tel2 = ver33(76)
	If error_level = "M" Then tel2 = ver33(77)
	If error_level = "Q" Then tel2 = ver33(78)
	If error_level = "H" Then tel2 = ver33(79)		
Case 34
    max_mod = ver34(12)
	qr_size = ver34(1)
	
    If error_level = "L" Then max_data_bits = ver34(20)
    If error_level = "M" Then max_data_bits = ver34(21)
    If error_level = "Q" Then max_data_bits = ver34(22)	
    If error_level = "H" Then max_data_bits = ver34(23)
	
	If error_level = "L" Then block1 = ver34(24)
    If error_level = "L" Then block2 = ver34(28)
    If error_level = "L" Then dw1 = ver34(26)
    If error_level = "L" Then dw2 = ver34(30)
	If error_level = "L" Then ew1 = ver34(27)
	If error_level = "L" Then ew2 = ver34(31)
	
	If error_level = "M" Then block1 = ver34(32)
    If error_level = "M" Then block2 = ver34(36)
    If error_level = "M" Then dw1 = ver34(34)
    If error_level = "M" Then dw2 = ver34(38)
	If error_level = "M" Then ew1 = ver34(35)
	If error_level = "M" Then ew2 = ver34(39)	
	
	If error_level = "Q" Then block1 = ver34(40)
    If error_level = "Q" Then block2 = ver34(44)
    If error_level = "Q" Then dw1 = ver34(42)
    If error_level = "Q" Then dw2 = ver34(46)
	If error_level = "Q" Then ew1 = ver34(43)
	If error_level = "Q" Then ew2 = ver34(47)		
	
	If error_level = "H" Then block1 = ver34(48)
    If error_level = "H" Then block2 = ver34(52)
    If error_level = "H" Then dw1 = ver34(50)
    If error_level = "H" Then dw2 = ver34(54)
	If error_level = "H" Then ew1 = ver34(51)
	If error_level = "H" Then ew2 = ver34(55)	
	
	If error_level = "L" Then tel2 = ver34(76)
	If error_level = "M" Then tel2 = ver34(77)
	If error_level = "Q" Then tel2 = ver34(78)
	If error_level = "H" Then tel2 = ver34(79)		
Case 35
    max_mod = ver35(12)
	qr_size = ver35(1)
	
    If error_level = "L" Then max_data_bits = ver35(20)
    If error_level = "M" Then max_data_bits = ver35(21)
    If error_level = "Q" Then max_data_bits = ver35(22)	
    If error_level = "H" Then max_data_bits = ver35(23)
	
	If error_level = "L" Then block1 = ver35(24)
    If error_level = "L" Then block2 = ver35(28)
    If error_level = "L" Then dw1 = ver35(26)
    If error_level = "L" Then dw2 = ver35(30)
	If error_level = "L" Then ew1 = ver35(27)
	If error_level = "L" Then ew2 = ver35(31)
	
	If error_level = "M" Then block1 = ver35(32)
    If error_level = "M" Then block2 = ver35(36)
    If error_level = "M" Then dw1 = ver35(34)
    If error_level = "M" Then dw2 = ver35(38)
	If error_level = "M" Then ew1 = ver35(35)
	If error_level = "M" Then ew2 = ver35(39)	
	
	If error_level = "Q" Then block1 = ver35(40)
    If error_level = "Q" Then block2 = ver35(44)
    If error_level = "Q" Then dw1 = ver35(42)
    If error_level = "Q" Then dw2 = ver35(46)
	If error_level = "Q" Then ew1 = ver35(43)
	If error_level = "Q" Then ew2 = ver35(47)		
	
	If error_level = "H" Then block1 = ver35(48)
    If error_level = "H" Then block2 = ver35(52)
    If error_level = "H" Then dw1 = ver35(50)
    If error_level = "H" Then dw2 = ver35(54)
	If error_level = "H" Then ew1 = ver35(51)
	If error_level = "H" Then ew2 = ver35(55)	
	
	If error_level = "L" Then tel2 = ver35(76)
	If error_level = "M" Then tel2 = ver35(77)
	If error_level = "Q" Then tel2 = ver35(78)
	If error_level = "H" Then tel2 = ver35(79)	
Case 36
    max_mod = ver36(12)
	qr_size = ver36(1)
	
    If error_level = "L" Then max_data_bits = ver36(20)
    If error_level = "M" Then max_data_bits = ver36(21)
    If error_level = "Q" Then max_data_bits = ver36(22)	
    If error_level = "H" Then max_data_bits = ver36(23)
	
	If error_level = "L" Then block1 = ver36(24)
    If error_level = "L" Then block2 = ver36(28)
    If error_level = "L" Then dw1 = ver36(26)
    If error_level = "L" Then dw2 = ver36(30)
	If error_level = "L" Then ew1 = ver36(27)
	If error_level = "L" Then ew2 = ver36(31)
	
	If error_level = "M" Then block1 = ver36(32)
    If error_level = "M" Then block2 = ver36(36)
    If error_level = "M" Then dw1 = ver36(34)
    If error_level = "M" Then dw2 = ver36(38)
	If error_level = "M" Then ew1 = ver36(35)
	If error_level = "M" Then ew2 = ver36(39)	
	
	If error_level = "Q" Then block1 = ver36(40)
    If error_level = "Q" Then block2 = ver36(44)
    If error_level = "Q" Then dw1 = ver36(42)
    If error_level = "Q" Then dw2 = ver36(46)
	If error_level = "Q" Then ew1 = ver36(43)
	If error_level = "Q" Then ew2 = ver36(47)		
	
	If error_level = "H" Then block1 = ver36(48)
    If error_level = "H" Then block2 = ver36(52)
    If error_level = "H" Then dw1 = ver36(50)
    If error_level = "H" Then dw2 = ver36(54)
	If error_level = "H" Then ew1 = ver36(51)
	If error_level = "H" Then ew2 = ver36(55)	
	
	If error_level = "L" Then tel2 = ver36(76)
	If error_level = "M" Then tel2 = ver36(77)
	If error_level = "Q" Then tel2 = ver36(78)
	If error_level = "H" Then tel2 = ver36(79)	
Case 37
    max_mod = ver37(12)
	qr_size = ver37(1)
	
    If error_level = "L" Then max_data_bits = ver37(20)
    If error_level = "M" Then max_data_bits = ver37(21)
    If error_level = "Q" Then max_data_bits = ver37(22)	
    If error_level = "H" Then max_data_bits = ver37(23)
	
	If error_level = "L" Then block1 = ver37(24)
    If error_level = "L" Then block2 = ver37(28)
    If error_level = "L" Then dw1 = ver37(26)
    If error_level = "L" Then dw2 = ver37(30)
	If error_level = "L" Then ew1 = ver37(27)
	If error_level = "L" Then ew2 = ver37(31)
	
	If error_level = "M" Then block1 = ver37(32)
    If error_level = "M" Then block2 = ver37(36)
    If error_level = "M" Then dw1 = ver37(34)
    If error_level = "M" Then dw2 = ver37(38)
	If error_level = "M" Then ew1 = ver37(35)
	If error_level = "M" Then ew2 = ver37(39)	
	
	If error_level = "Q" Then block1 = ver37(40)
    If error_level = "Q" Then block2 = ver37(44)
    If error_level = "Q" Then dw1 = ver37(42)
    If error_level = "Q" Then dw2 = ver37(46)
	If error_level = "Q" Then ew1 = ver37(43)
	If error_level = "Q" Then ew2 = ver37(47)		
	
	If error_level = "H" Then block1 = ver37(48)
    If error_level = "H" Then block2 = ver37(52)
    If error_level = "H" Then dw1 = ver37(50)
    If error_level = "H" Then dw2 = ver37(54)
	If error_level = "H" Then ew1 = ver37(51)
	If error_level = "H" Then ew2 = ver37(55)	
	
	If error_level = "L" Then tel2 = ver37(76)
	If error_level = "M" Then tel2 = ver37(77)
	If error_level = "Q" Then tel2 = ver37(78)
	If error_level = "H" Then tel2 = ver37(79)	
Case 38
    max_mod = ver38(12)
	qr_size = ver38(1)
 	
    If error_level = "L" Then max_data_bits = ver38(20)
    If error_level = "M" Then max_data_bits = ver38(21)
    If error_level = "Q" Then max_data_bits = ver38(22)	
    If error_level = "H" Then max_data_bits = ver38(23)
	
	If error_level = "L" Then block1 = ver38(24)
    If error_level = "L" Then block2 = ver38(28)
    If error_level = "L" Then dw1 = ver38(26)
    If error_level = "L" Then dw2 = ver38(30)
	If error_level = "L" Then ew1 = ver38(27)
	If error_level = "L" Then ew2 = ver38(31)
	
	If error_level = "M" Then block1 = ver38(32)
    If error_level = "M" Then block2 = ver38(36)
    If error_level = "M" Then dw1 = ver38(34)
    If error_level = "M" Then dw2 = ver38(38)
	If error_level = "M" Then ew1 = ver38(35)
	If error_level = "M" Then ew2 = ver38(39)	
	
	If error_level = "Q" Then block1 = ver38(40)
    If error_level = "Q" Then block2 = ver38(44)
    If error_level = "Q" Then dw1 = ver38(42)
    If error_level = "Q" Then dw2 = ver38(46)
	If error_level = "Q" Then ew1 = ver38(43)
	If error_level = "Q" Then ew2 = ver38(47)		
	
	If error_level = "H" Then block1 = ver38(48)
    If error_level = "H" Then block2 = ver38(52)
    If error_level = "H" Then dw1 = ver38(50)
    If error_level = "H" Then dw2 = ver38(54)
	If error_level = "H" Then ew1 = ver38(51)
	If error_level = "H" Then ew2 = ver38(55)	
	
	If error_level = "L" Then tel2 = ver38(76)
	If error_level = "M" Then tel2 = ver38(77)
	If error_level = "Q" Then tel2 = ver38(78)
	If error_level = "H" Then tel2 = ver38(79)	
Case 39
    max_mod = ver39(12)
	qr_size = ver39(1)
	
    If error_level = "L" Then max_data_bits = ver39(20)
    If error_level = "M" Then max_data_bits = ver39(21)
    If error_level = "Q" Then max_data_bits = ver39(22)	
    If error_level = "H" Then max_data_bits = ver39(23)
	
	If error_level = "L" Then block1 = ver39(24)
    If error_level = "L" Then block2 = ver39(28)
    If error_level = "L" Then dw1 = ver39(26)
    If error_level = "L" Then dw2 = ver39(30)
	If error_level = "L" Then ew1 = ver39(27)
	If error_level = "L" Then ew2 = ver39(31)
	
	If error_level = "M" Then block1 = ver39(32)
    If error_level = "M" Then block2 = ver39(36)
    If error_level = "M" Then dw1 = ver39(34)
    If error_level = "M" Then dw2 = ver39(38)
	If error_level = "M" Then ew1 = ver39(35)
	If error_level = "M" Then ew2 = ver39(39)	
	
	If error_level = "Q" Then block1 = ver39(40)
    If error_level = "Q" Then block2 = ver39(44)
    If error_level = "Q" Then dw1 = ver39(42)
    If error_level = "Q" Then dw2 = ver39(46)
	If error_level = "Q" Then ew1 = ver39(43)
	If error_level = "Q" Then ew2 = ver39(47)		
	
	If error_level = "H" Then block1 = ver39(48)
    If error_level = "H" Then block2 = ver39(52)
    If error_level = "H" Then dw1 = ver39(50)
    If error_level = "H" Then dw2 = ver39(54)
	If error_level = "H" Then ew1 = ver39(51)
	If error_level = "H" Then ew2 = ver39(55)	
	
	If error_level = "L" Then tel2 = ver39(76)
	If error_level = "M" Then tel2 = ver39(77)
	If error_level = "Q" Then tel2 = ver39(78)
	If error_level = "H" Then tel2 = ver39(79)	
Case 40
    max_mod = ver40(12)
	qr_size = ver40(1)
	
    If error_level = "L" Then max_data_bits = ver40(20)
    If error_level = "M" Then max_data_bits = ver40(21)
    If error_level = "Q" Then max_data_bits = ver40(22)	
    If error_level = "H" Then max_data_bits = ver40(23)
	
	If error_level = "L" Then block1 = ver40(24)
    If error_level = "L" Then block2 = ver40(28)
    If error_level = "L" Then dw1 = ver40(26)
    If error_level = "L" Then dw2 = ver40(30)
	If error_level = "L" Then ew1 = ver40(27)
	If error_level = "L" Then ew2 = ver40(31)
	
	If error_level = "M" Then block1 = ver40(32)
    If error_level = "M" Then block2 = ver40(36)
    If error_level = "M" Then dw1 = ver40(34)
    If error_level = "M" Then dw2 = ver40(38)
	If error_level = "M" Then ew1 = ver40(35)
	If error_level = "M" Then ew2 = ver40(39)	
	
	If error_level = "Q" Then block1 = ver40(40)
    If error_level = "Q" Then block2 = ver40(44)
    If error_level = "Q" Then dw1 = ver40(42)
    If error_level = "Q" Then dw2 = ver40(46)
	If error_level = "Q" Then ew1 = ver40(43)
	If error_level = "Q" Then ew2 = ver40(47)		
	
	If error_level = "H" Then block1 = ver40(48)
    If error_level = "H" Then block2 = ver40(52)
    If error_level = "H" Then dw1 = ver40(50)
    If error_level = "H" Then dw2 = ver40(54)
	If error_level = "H" Then ew1 = ver40(51)
	If error_level = "H" Then ew2 = ver40(55)	
	
	If error_level = "L" Then tel2 = ver40(76)
	If error_level = "M" Then tel2 = ver40(77)
	If error_level = "Q" Then tel2 = ver40(78)
	If error_level = "H" Then tel2 = ver40(79)	
End Select

End Sub

Private Sub determine_max_len

If error_level = "L" AND sf.len(dat) >= 0 Then version_no = 1
If error_level = "L" AND sf.len(dat) > ver1(64) Then version_no = 2
If error_level = "L" AND sf.len(dat) > ver2(64) Then version_no = 3
If error_level = "L" AND sf.len(dat) > ver3(64) Then version_no = 4
If error_level = "L" AND sf.len(dat) > ver4(64) Then version_no = 5
If error_level = "L" AND sf.len(dat) > ver5(64) Then version_no = 6
If error_level = "L" AND sf.len(dat) > ver6(64) Then version_no = 7
If error_level = "L" AND sf.len(dat) > ver7(64) Then version_no = 8
If error_level = "L" AND sf.len(dat) > ver8(64) Then version_no = 9
If error_level = "L" AND sf.len(dat) > ver9(64) Then version_no = 10
If error_level = "L" AND sf.len(dat) > ver10(64) Then version_no = 11
If error_level = "L" AND sf.len(dat) > ver11(64) Then version_no = 12
If error_level = "L" AND sf.len(dat) > ver12(64) Then version_no = 13
If error_level = "L" AND sf.len(dat) > ver13(64) Then version_no = 14
If error_level = "L" AND sf.len(dat) > ver14(64) Then version_no = 15
If error_level = "L" AND sf.len(dat) > ver15(64) Then version_no = 16
If error_level = "L" AND sf.len(dat) > ver16(64) Then version_no = 17
If error_level = "L" AND sf.len(dat) > ver17(64) Then version_no = 18
If error_level = "L" AND sf.len(dat) > ver18(64) Then version_no = 19
If error_level = "L" AND sf.len(dat) > ver19(64) Then version_no = 20
If error_level = "L" AND sf.len(dat) > ver20(64) Then version_no = 21
If error_level = "L" AND sf.len(dat) > ver21(64) Then version_no = 22
If error_level = "L" AND sf.len(dat) > ver22(64) Then version_no = 23
If error_level = "L" AND sf.len(dat) > ver23(64) Then version_no = 24
If error_level = "L" AND sf.len(dat) > ver24(64) Then version_no = 25
If error_level = "L" AND sf.len(dat) > ver25(64) Then version_no = 26
If error_level = "L" AND sf.len(dat) > ver26(64) Then version_no = 27
If error_level = "L" AND sf.len(dat) > ver27(64) Then version_no = 28
If error_level = "L" AND sf.len(dat) > ver28(64) Then version_no = 29
If error_level = "L" AND sf.len(dat) > ver29(64) Then version_no = 30
If error_level = "L" AND sf.len(dat) > ver30(64) Then version_no = 31
If error_level = "L" AND sf.len(dat) > ver31(64) Then version_no = 32
If error_level = "L" AND sf.len(dat) > ver32(64) Then version_no = 33
If error_level = "L" AND sf.len(dat) > ver33(64) Then version_no = 34
If error_level = "L" AND sf.len(dat) > ver34(64) Then version_no = 35
If error_level = "L" AND sf.len(dat) > ver35(64) Then version_no = 36
If error_level = "L" AND sf.len(dat) > ver36(64) Then version_no = 37
If error_level = "L" AND sf.len(dat) > ver37(64) Then version_no = 38
If error_level = "L" AND sf.len(dat) > ver38(64) Then version_no = 39
If error_level = "L" AND sf.len(dat) > ver39(64) Then version_no = 40

If error_level = "M" AND sf.len(dat) >= 0 Then version_no = 1
If error_level = "M" AND sf.len(dat) > ver1(65) Then version_no = 2
If error_level = "M" AND sf.len(dat) > ver2(65) Then version_no = 3
If error_level = "M" AND sf.len(dat) > ver3(65) Then version_no = 4
If error_level = "M" AND sf.len(dat) > ver4(65) Then version_no = 5
If error_level = "M" AND sf.len(dat) > ver5(65) Then version_no = 6
If error_level = "M" AND sf.len(dat) > ver6(65) Then version_no = 7
If error_level = "M" AND sf.len(dat) > ver7(65) Then version_no = 8
If error_level = "M" AND sf.len(dat) > ver8(65) Then version_no = 9
If error_level = "M" AND sf.len(dat) > ver9(65) Then version_no = 10
If error_level = "M" AND sf.len(dat) > ver10(65) Then version_no = 11
If error_level = "M" AND sf.len(dat) > ver11(65) Then version_no = 12
If error_level = "M" AND sf.len(dat) > ver12(65) Then version_no = 13
If error_level = "M" AND sf.len(dat) > ver13(65) Then version_no = 14
If error_level = "M" AND sf.len(dat) > ver14(65) Then version_no = 15
If error_level = "M" AND sf.len(dat) > ver15(65) Then version_no = 16
If error_level = "M" AND sf.len(dat) > ver16(65) Then version_no = 17
If error_level = "M" AND sf.len(dat) > ver17(65) Then version_no = 18
If error_level = "M" AND sf.len(dat) > ver18(65) Then version_no = 19
If error_level = "M" AND sf.len(dat) > ver19(65) Then version_no = 20
If error_level = "M" AND sf.len(dat) > ver20(65) Then version_no = 21
If error_level = "M" AND sf.len(dat) > ver21(65) Then version_no = 22
If error_level = "M" AND sf.len(dat) > ver22(65) Then version_no = 23
If error_level = "M" AND sf.len(dat) > ver23(65) Then version_no = 24
If error_level = "M" AND sf.len(dat) > ver24(65) Then version_no = 25
If error_level = "M" AND sf.len(dat) > ver25(65) Then version_no = 26
If error_level = "M" AND sf.len(dat) > ver26(65) Then version_no = 27
If error_level = "M" AND sf.len(dat) > ver27(65) Then version_no = 28
If error_level = "M" AND sf.len(dat) > ver28(65) Then version_no = 29
If error_level = "M" AND sf.len(dat) > ver29(65) Then version_no = 30
If error_level = "M" AND sf.len(dat) > ver30(65) Then version_no = 31
If error_level = "M" AND sf.len(dat) > ver31(65) Then version_no = 32
If error_level = "M" AND sf.len(dat) > ver32(65) Then version_no = 33
If error_level = "M" AND sf.len(dat) > ver33(65) Then version_no = 34
If error_level = "M" AND sf.len(dat) > ver34(65) Then version_no = 35
If error_level = "M" AND sf.len(dat) > ver35(65) Then version_no = 36
If error_level = "M" AND sf.len(dat) > ver36(65) Then version_no = 37
If error_level = "M" AND sf.len(dat) > ver37(65) Then version_no = 38
If error_level = "M" AND sf.len(dat) > ver38(65) Then version_no = 39
If error_level = "M" AND sf.len(dat) > ver39(65) Then version_no = 40

If error_level = "Q" AND sf.len(dat) >= 0 Then version_no = 1
If error_level = "Q" AND sf.len(dat) > ver1(66) Then version_no = 2
If error_level = "Q" AND sf.len(dat) > ver2(66) Then version_no = 3
If error_level = "Q" AND sf.len(dat) > ver3(66) Then version_no = 4
If error_level = "Q" AND sf.len(dat) > ver4(66) Then version_no = 5
If error_level = "Q" AND sf.len(dat) > ver5(66) Then version_no = 6
If error_level = "Q" AND sf.len(dat) > ver6(66) Then version_no = 7
If error_level = "Q" AND sf.len(dat) > ver7(66) Then version_no = 8
If error_level = "Q" AND sf.len(dat) > ver8(66) Then version_no = 9
If error_level = "Q" AND sf.len(dat) > ver9(66) Then version_no = 10
If error_level = "Q" AND sf.len(dat) > ver10(66) Then version_no = 11
If error_level = "Q" AND sf.len(dat) > ver11(66) Then version_no = 12
If error_level = "Q" AND sf.len(dat) > ver12(66) Then version_no = 13
If error_level = "Q" AND sf.len(dat) > ver13(66) Then version_no = 14
If error_level = "Q" AND sf.len(dat) > ver14(66) Then version_no = 15
If error_level = "Q" AND sf.len(dat) > ver15(66) Then version_no = 16
If error_level = "Q" AND sf.len(dat) > ver16(66) Then version_no = 17
If error_level = "Q" AND sf.len(dat) > ver17(66) Then version_no = 18
If error_level = "Q" AND sf.len(dat) > ver18(66) Then version_no = 19
If error_level = "Q" AND sf.len(dat) > ver19(66) Then version_no = 20
If error_level = "Q" AND sf.len(dat) > ver20(66) Then version_no = 21
If error_level = "Q" AND sf.len(dat) > ver21(66) Then version_no = 22
If error_level = "Q" AND sf.len(dat) > ver22(66) Then version_no = 23
If error_level = "Q" AND sf.len(dat) > ver23(66) Then version_no = 24
If error_level = "Q" AND sf.len(dat) > ver24(66) Then version_no = 25
If error_level = "Q" AND sf.len(dat) > ver25(66) Then version_no = 26
If error_level = "Q" AND sf.len(dat) > ver26(66) Then version_no = 27
If error_level = "Q" AND sf.len(dat) > ver27(66) Then version_no = 28
If error_level = "Q" AND sf.len(dat) > ver28(66) Then version_no = 29
If error_level = "Q" AND sf.len(dat) > ver29(66) Then version_no = 30
If error_level = "Q" AND sf.len(dat) > ver30(66) Then version_no = 31
If error_level = "Q" AND sf.len(dat) > ver31(66) Then version_no = 32
If error_level = "Q" AND sf.len(dat) > ver32(66) Then version_no = 33
If error_level = "Q" AND sf.len(dat) > ver33(66) Then version_no = 34
If error_level = "Q" AND sf.len(dat) > ver34(66) Then version_no = 35
If error_level = "Q" AND sf.len(dat) > ver35(66) Then version_no = 36
If error_level = "Q" AND sf.len(dat) > ver36(66) Then version_no = 37
If error_level = "Q" AND sf.len(dat) > ver37(66) Then version_no = 38
If error_level = "Q" AND sf.len(dat) > ver38(66) Then version_no = 39
If error_level = "Q" AND sf.len(dat) > ver39(66) Then version_no = 40

If error_level = "H" AND sf.len(dat) >= 0 Then version_no = 1
If error_level = "H" AND sf.len(dat) > ver1(67) Then version_no = 2
If error_level = "H" AND sf.len(dat) > ver2(67) Then version_no = 3
If error_level = "H" AND sf.len(dat) > ver3(67) Then version_no = 4
If error_level = "H" AND sf.len(dat) > ver4(67) Then version_no = 5
If error_level = "H" AND sf.len(dat) > ver5(67) Then version_no = 6
If error_level = "H" AND sf.len(dat) > ver6(67) Then version_no = 7
If error_level = "H" AND sf.len(dat) > ver7(67) Then version_no = 8
If error_level = "H" AND sf.len(dat) > ver8(67) Then version_no = 9
If error_level = "H" AND sf.len(dat) > ver9(67) Then version_no = 10
If error_level = "H" AND sf.len(dat) > ver10(67) Then version_no = 11
If error_level = "H" AND sf.len(dat) > ver11(67) Then version_no = 12
If error_level = "H" AND sf.len(dat) > ver12(67) Then version_no = 13
If error_level = "H" AND sf.len(dat) > ver13(67) Then version_no = 14
If error_level = "H" AND sf.len(dat) > ver14(67) Then version_no = 15
If error_level = "H" AND sf.len(dat) > ver15(67) Then version_no = 16
If error_level = "H" AND sf.len(dat) > ver16(67) Then version_no = 17
If error_level = "H" AND sf.len(dat) > ver17(67) Then version_no = 18
If error_level = "H" AND sf.len(dat) > ver18(67) Then version_no = 19
If error_level = "H" AND sf.len(dat) > ver19(67) Then version_no = 20
If error_level = "H" AND sf.len(dat) > ver20(67) Then version_no = 21
If error_level = "H" AND sf.len(dat) > ver21(67) Then version_no = 22
If error_level = "H" AND sf.len(dat) > ver22(67) Then version_no = 23
If error_level = "H" AND sf.len(dat) > ver23(67) Then version_no = 24
If error_level = "H" AND sf.len(dat) > ver24(67) Then version_no = 25
If error_level = "H" AND sf.len(dat) > ver25(67) Then version_no = 26
If error_level = "H" AND sf.len(dat) > ver26(67) Then version_no = 27
If error_level = "H" AND sf.len(dat) > ver27(67) Then version_no = 28
If error_level = "H" AND sf.len(dat) > ver28(67) Then version_no = 29
If error_level = "H" AND sf.len(dat) > ver29(67) Then version_no = 30
If error_level = "H" AND sf.len(dat) > ver30(67) Then version_no = 31
If error_level = "H" AND sf.len(dat) > ver31(67) Then version_no = 32
If error_level = "H" AND sf.len(dat) > ver32(67) Then version_no = 33
If error_level = "H" AND sf.len(dat) > ver33(67) Then version_no = 34
If error_level = "H" AND sf.len(dat) > ver34(67) Then version_no = 35
If error_level = "H" AND sf.len(dat) > ver35(67) Then version_no = 36
If error_level = "H" AND sf.len(dat) > ver36(67) Then version_no = 37
If error_level = "H" AND sf.len(dat) > ver37(67) Then version_no = 38
If error_level = "H" AND sf.len(dat) > ver38(67) Then version_no = 39
If error_level = "H" AND sf.len(dat) > ver39(67) Then version_no = 40

If version_no = 1 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver1(64)	
If version_no = 1 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver1(65)	
If version_no = 1 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver1(66)	
If version_no = 1 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver1(67)	
If version_no = 1 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver1(68)	
If version_no = 1 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver1(69)	
If version_no = 1 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver1(70)	
If version_no = 1 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver1(71)	
If version_no = 1 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver1(72)	
If version_no = 1 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver1(73)	
If version_no = 1 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver1(74)	
If version_no = 1 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver1(75)	

If version_no = 2 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver2(64)	
If version_no = 2 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver2(65)	
If version_no = 2 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver2(66)	
If version_no = 2 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver2(67)	
If version_no = 2 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver2(68)	
If version_no = 2 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver2(69)	
If version_no = 2 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver2(70)	
If version_no = 2 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver2(71)	
If version_no = 2 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver2(72)	
If version_no = 2 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver2(73)	
If version_no = 2 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver2(74)	
If version_no = 2 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver2(75)	

If version_no = 3 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver3(64)	
If version_no = 3 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver3(65)	
If version_no = 3 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver3(66)	
If version_no = 3 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver3(67)	
If version_no = 3 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver3(68)	
If version_no = 3 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver3(69)	
If version_no = 3 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver3(70)	
If version_no = 3 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver3(71)	
If version_no = 3 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver3(72)	
If version_no = 3 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver3(73)	
If version_no = 3 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver3(74)	
If version_no = 3 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver3(75)	

If version_no = 4 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver4(64)	
If version_no = 4 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver4(65)	
If version_no = 4 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver4(66)	
If version_no = 4 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver4(67)	
If version_no = 4 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver4(68)	
If version_no = 4 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver4(69)	
If version_no = 4 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver4(70)	
If version_no = 4 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver4(71)	
If version_no = 4 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver4(72)	
If version_no = 4 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver4(73)	
If version_no = 4 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver4(74)	
If version_no = 4 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver4(75)

If version_no = 5 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver5(64)	
If version_no = 5 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver5(65)	
If version_no = 5 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver5(66)	
If version_no = 5 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver5(67)	
If version_no = 5 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver5(68)	
If version_no = 5 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver5(69)	
If version_no = 5 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver5(70)	
If version_no = 5 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver5(71)	
If version_no = 5 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver5(72)	
If version_no = 5 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver5(73)	
If version_no = 5 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver5(74)	
If version_no = 5 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver5(75)

If version_no = 6 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver6(64)	
If version_no = 6 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver6(65)	
If version_no = 6 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver6(66)	
If version_no = 6 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver6(67)	
If version_no = 6 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver6(68)	
If version_no = 6 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver6(69)	
If version_no = 6 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver6(70)	
If version_no = 6 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver6(71)	
If version_no = 6 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver6(72)	
If version_no = 6 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver6(73)	
If version_no = 6 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver6(74)	
If version_no = 6 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver6(75)

If version_no = 7 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver7(64)	
If version_no = 7 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver7(65)	
If version_no = 7 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver7(66)	
If version_no = 7 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver7(67)	
If version_no = 7 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver7(68)	
If version_no = 7 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver7(69)	
If version_no = 7 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver7(70)	
If version_no = 7 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver7(71)	
If version_no = 7 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver7(72)	
If version_no = 7 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver7(73)	
If version_no = 7 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver7(74)	
If version_no = 7 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver7(75)

If version_no = 8 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver8(64)	
If version_no = 8 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver8(65)	
If version_no = 8 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver8(66)	
If version_no = 8 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver8(67)	
If version_no = 8 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver8(68)	
If version_no = 8 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver8(69)	
If version_no = 8 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver8(70)	
If version_no = 8 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver8(71)	
If version_no = 8 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver8(72)	
If version_no = 8 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver8(73)	
If version_no = 8 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver8(74)	
If version_no = 8 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver8(75)

If version_no = 9 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver9(64)	
If version_no = 9 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver9(65)	
If version_no = 9 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver9(66)	
If version_no = 9 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver9(67)	
If version_no = 9 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver9(68)	
If version_no = 9 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver9(69)	
If version_no = 9 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver9(70)	
If version_no = 9 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver9(71)	
If version_no = 9 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver9(72)	
If version_no = 9 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver9(73)	
If version_no = 9 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver9(74)	
If version_no = 9 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver9(75)

If version_no = 10 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver10(64)	
If version_no = 10 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver10(65)	
If version_no = 10 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver10(66)	
If version_no = 10 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver10(67)	
If version_no = 10 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver10(68)	
If version_no = 10 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver10(69)	
If version_no = 10 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver10(70)	
If version_no = 10 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver10(71)	
If version_no = 10 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver10(72)	
If version_no = 10 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver10(73)	
If version_no = 10 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver10(74)	
If version_no = 10 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver10(75)

If version_no = 11 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver11(64)	
If version_no = 11 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver11(65)	
If version_no = 11 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver11(66)	
If version_no = 11 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver11(67)	
If version_no = 11 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver11(68)	
If version_no = 11 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver11(69)	
If version_no = 11 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver11(70)	
If version_no = 11 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver11(71)	
If version_no = 11 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver11(72)	
If version_no = 11 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver11(73)	
If version_no = 11 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver11(74)	
If version_no = 11 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver11(75)

If version_no = 12 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver12(64)	
If version_no = 12 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver12(65)	
If version_no = 12 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver12(66)	
If version_no = 12 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver12(67)	
If version_no = 12 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver12(68)	
If version_no = 12 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver12(69)	
If version_no = 12 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver12(70)	
If version_no = 12 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver12(71)	
If version_no = 12 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver12(72)	
If version_no = 12 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver12(73)	
If version_no = 12 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver12(74)	
If version_no = 12 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver12(75)

If version_no = 13 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver13(64)	
If version_no = 13 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver13(65)	
If version_no = 13 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver13(66)	
If version_no = 13 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver13(67)	
If version_no = 13 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver13(68)	
If version_no = 13 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver13(69)	
If version_no = 13 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver13(70)	
If version_no = 13 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver13(71)	
If version_no = 13 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver13(72)	
If version_no = 13 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver13(73)	
If version_no = 13 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver13(74)	
If version_no = 13 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver13(75)

If version_no = 14 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver14(64)	
If version_no = 14 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver14(65)	
If version_no = 14 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver14(66)	
If version_no = 14 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver14(67)	
If version_no = 14 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver14(68)	
If version_no = 14 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver14(69)	
If version_no = 14 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver14(70)	
If version_no = 14 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver14(71)	
If version_no = 14 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver14(72)	
If version_no = 14 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver14(73)	
If version_no = 14 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver14(74)	
If version_no = 14 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver14(75)

If version_no = 15 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver15(64)	
If version_no = 15 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver15(65)	
If version_no = 15 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver15(66)	
If version_no = 15 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver15(67)	
If version_no = 15 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver15(68)	
If version_no = 15 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver15(69)	
If version_no = 15 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver15(70)	
If version_no = 15 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver15(71)	
If version_no = 15 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver15(72)	
If version_no = 15 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver15(73)	
If version_no = 15 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver15(74)	
If version_no = 15 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver15(75)

If version_no = 16 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver16(64)	
If version_no = 16 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver16(65)	
If version_no = 16 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver16(66)	
If version_no = 16 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver16(67)	
If version_no = 16 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver16(68)	
If version_no = 16 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver16(69)	
If version_no = 16 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver16(70)	
If version_no = 16 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver16(71)	
If version_no = 16 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver16(72)	
If version_no = 16 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver16(73)	
If version_no = 16 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver16(74)	
If version_no = 16 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver16(75)

If version_no = 17 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver17(64)	
If version_no = 17 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver17(65)	
If version_no = 17 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver17(66)	
If version_no = 17 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver17(67)	
If version_no = 17 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver17(68)	
If version_no = 17 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver17(69)	
If version_no = 17 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver17(70)	
If version_no = 17 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver17(71)	
If version_no = 17 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver17(72)	
If version_no = 17 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver17(73)	
If version_no = 17 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver17(74)	
If version_no = 17 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver17(75)

If version_no = 18 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver18(64)	
If version_no = 18 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver18(65)	
If version_no = 18 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver18(66)	
If version_no = 18 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver18(67)	
If version_no = 18 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver18(68)	
If version_no = 18 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver18(69)	
If version_no = 18 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver18(70)	
If version_no = 18 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver18(71)	
If version_no = 18 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver18(72)	
If version_no = 18 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver18(73)	
If version_no = 18 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver18(74)	
If version_no = 18 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver18(75)

If version_no = 19 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver19(64)	
If version_no = 19 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver19(65)	
If version_no = 19 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver19(66)	
If version_no = 19 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver19(67)	
If version_no = 19 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver19(68)	
If version_no = 19 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver19(69)	
If version_no = 19 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver19(70)	
If version_no = 19 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver19(71)	
If version_no = 19 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver19(72)	
If version_no = 19 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver19(73)	
If version_no = 19 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver19(74)	
If version_no = 19 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver19(75)

If version_no = 20 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver20(64)	
If version_no = 20 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver20(65)	
If version_no = 20 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver20(66)	
If version_no = 20 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver20(67)	
If version_no = 20 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver20(68)	
If version_no = 20 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver20(69)	
If version_no = 20 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver20(70)	
If version_no = 20 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver20(71)	
If version_no = 20 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver20(72)	
If version_no = 20 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver20(73)	
If version_no = 20 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver20(74)	
If version_no = 20 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver20(75)

If version_no = 21 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver21(64)	
If version_no = 21 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver21(65)	
If version_no = 21 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver21(66)	
If version_no = 21 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver21(67)	
If version_no = 21 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver21(68)	
If version_no = 21 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver21(69)	
If version_no = 21 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver21(70)	
If version_no = 21 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver21(71)	
If version_no = 21 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver21(72)	
If version_no = 21 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver21(73)	
If version_no = 21 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver21(74)	
If version_no = 21 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver21(75)

If version_no = 22 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver22(64)	
If version_no = 22 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver22(65)	
If version_no = 22 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver22(66)	
If version_no = 22 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver22(67)	
If version_no = 22 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver22(68)	
If version_no = 22 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver22(69)	
If version_no = 22 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver22(70)	
If version_no = 22 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver22(71)	
If version_no = 22 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver22(72)	
If version_no = 22 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver22(73)	
If version_no = 22 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver22(74)	
If version_no = 22 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver22(75)

If version_no = 23 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver23(64)	
If version_no = 23 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver23(65)	
If version_no = 23 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver23(66)	
If version_no = 23 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver23(67)	
If version_no = 23 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver23(68)	
If version_no = 23 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver23(69)	
If version_no = 23 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver23(70)	
If version_no = 23 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver23(71)	
If version_no = 23 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver23(72)	
If version_no = 23 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver23(73)	
If version_no = 23 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver23(74)	
If version_no = 23 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver23(75)

If version_no = 24 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver24(64)	
If version_no = 24 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver24(65)	
If version_no = 24 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver24(66)	
If version_no = 24 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver24(67)	
If version_no = 24 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver24(68)	
If version_no = 24 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver24(69)	
If version_no = 24 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver24(70)	
If version_no = 24 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver24(71)	
If version_no = 24 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver24(72)	
If version_no = 24 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver24(73)	
If version_no = 24 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver24(74)	
If version_no = 24 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver24(75)

If version_no = 25 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver25(64)	
If version_no = 25 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver25(65)	
If version_no = 25 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver25(66)	
If version_no = 25 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver25(67)	
If version_no = 25 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver25(68)	
If version_no = 25 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver25(69)	
If version_no = 25 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver25(70)	
If version_no = 25 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver25(71)	
If version_no = 25 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver25(72)	
If version_no = 25 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver25(73)	
If version_no = 25 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver25(74)	
If version_no = 25 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver25(75)

If version_no = 26 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver26(64)	
If version_no = 26 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver26(65)	
If version_no = 26 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver26(66)	
If version_no = 26 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver26(67)	
If version_no = 26 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver26(68)	
If version_no = 26 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver26(69)	
If version_no = 26 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver26(70)	
If version_no = 26 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver26(71)	
If version_no = 26 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver26(72)	
If version_no = 26 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver26(73)	
If version_no = 26 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver26(74)	
If version_no = 26 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver26(75)

If version_no = 27 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver27(64)	
If version_no = 27 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver27(65)	
If version_no = 27 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver27(66)	
If version_no = 27 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver27(67)	
If version_no = 27 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver27(68)	
If version_no = 27 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver27(69)	
If version_no = 27 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver27(70)	
If version_no = 27 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver27(71)	
If version_no = 27 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver27(72)	
If version_no = 27 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver27(73)	
If version_no = 27 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver27(74)	
If version_no = 27 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver27(75)

If version_no = 28 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver28(64)	
If version_no = 28 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver28(65)	
If version_no = 28 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver28(66)	
If version_no = 28 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver28(67)	
If version_no = 28 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver28(68)	
If version_no = 28 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver28(69)	
If version_no = 28 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver28(70)	
If version_no = 28 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver28(71)	
If version_no = 28 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver28(72)	
If version_no = 28 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver28(73)	
If version_no = 28 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver28(74)	
If version_no = 28 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver28(75)

If version_no = 29 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver29(64)	
If version_no = 29 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver29(65)	
If version_no = 29 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver29(66)	
If version_no = 29 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver29(67)	
If version_no = 29 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver29(68)	
If version_no = 29 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver29(69)	
If version_no = 29 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver29(70)	
If version_no = 29 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver29(71)	
If version_no = 29 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver29(72)	
If version_no = 29 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver29(73)	
If version_no = 29 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver29(74)	
If version_no = 29 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver29(75)

If version_no = 30 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver30(64)	
If version_no = 30 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver30(65)	
If version_no = 30 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver30(66)	
If version_no = 30 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver30(67)	
If version_no = 30 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver30(68)	
If version_no = 30 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver30(69)	
If version_no = 30 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver30(70)	
If version_no = 30 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver30(71)	
If version_no = 30 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver30(72)	
If version_no = 30 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver30(73)	
If version_no = 30 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver30(74)	
If version_no = 30 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver30(75)

If version_no = 31 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver31(64)	
If version_no = 31 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver31(65)	
If version_no = 31 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver31(66)	
If version_no = 31 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver31(67)	
If version_no = 31 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver31(68)	
If version_no = 31 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver31(69)	
If version_no = 31 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver31(70)	
If version_no = 31 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver31(71)	
If version_no = 31 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver31(72)	
If version_no = 31 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver31(73)	
If version_no = 31 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver31(74)	
If version_no = 31 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver31(75)

If version_no = 32 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver32(64)	
If version_no = 32 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver32(65)	
If version_no = 32 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver32(66)	
If version_no = 32 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver32(67)	
If version_no = 32 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver32(68)	
If version_no = 32 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver32(69)	
If version_no = 32 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver32(70)	
If version_no = 32 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver32(71)	
If version_no = 32 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver32(72)	
If version_no = 32 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver32(73)	
If version_no = 32 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver32(74)	
If version_no = 32 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver32(75)

If version_no = 33 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver33(64)	
If version_no = 33 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver33(65)	
If version_no = 33 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver33(66)	
If version_no = 33 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver33(67)	
If version_no = 33 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver33(68)	
If version_no = 33 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver33(69)	
If version_no = 33 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver33(70)	
If version_no = 33 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver33(71)	
If version_no = 33 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver33(72)	
If version_no = 33 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver33(73)	
If version_no = 33 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver33(74)	
If version_no = 33 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver33(75)

If version_no = 34 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver34(64)	
If version_no = 34 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver34(65)	
If version_no = 34 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver34(66)	
If version_no = 34 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver34(67)	
If version_no = 34 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver34(68)	
If version_no = 34 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver34(69)	
If version_no = 34 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver34(70)	
If version_no = 34 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver34(71)	
If version_no = 34 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver34(72)	
If version_no = 34 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver34(73)	
If version_no = 34 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver34(74)	
If version_no = 34 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver34(75)

If version_no = 35 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver35(64)	
If version_no = 35 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver35(65)	
If version_no = 35 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver35(66)	
If version_no = 35 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver35(67)	
If version_no = 35 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver35(68)	
If version_no = 35 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver35(69)	
If version_no = 35 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver35(70)	
If version_no = 35 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver35(71)	
If version_no = 35 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver35(72)	
If version_no = 35 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver35(73)	
If version_no = 35 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver35(74)	
If version_no = 35 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver35(75)

If version_no = 36 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver36(64)	
If version_no = 36 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver36(65)	
If version_no = 36 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver36(66)	
If version_no = 36 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver36(67)	
If version_no = 36 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver36(68)	
If version_no = 36 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver36(69)	
If version_no = 36 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver36(70)	
If version_no = 36 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver36(71)	
If version_no = 36 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver36(72)	
If version_no = 36 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver36(73)	
If version_no = 36 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver36(74)	
If version_no = 36 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver36(75)

If version_no = 37 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver37(64)	
If version_no = 37 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver37(65)	
If version_no = 37 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver37(66)	
If version_no = 37 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver37(67)	
If version_no = 37 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver37(68)	
If version_no = 37 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver37(69)	
If version_no = 37 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver37(70)	
If version_no = 37 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver37(71)	
If version_no = 37 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver37(72)	
If version_no = 37 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver37(73)	
If version_no = 37 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver37(74)	
If version_no = 37 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver37(75)

If version_no = 38 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver38(64)	
If version_no = 38 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver38(65)	
If version_no = 38 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver38(66)	
If version_no = 38 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver38(67)	
If version_no = 38 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver38(68)	
If version_no = 38 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver38(69)	
If version_no = 38 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver38(70)	
If version_no = 38 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver38(71)	
If version_no = 38 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver38(72)	
If version_no = 38 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver38(73)	
If version_no = 38 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver38(74)	
If version_no = 38 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver38(75)

If version_no = 39 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver39(64)	
If version_no = 39 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver39(65)	
If version_no = 39 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver39(66)	
If version_no = 39 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver39(67)	
If version_no = 39 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver39(68)	
If version_no = 39 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver39(69)	
If version_no = 39 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver39(70)	
If version_no = 39 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver39(71)	
If version_no = 39 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver39(72)	
If version_no = 39 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver39(73)	
If version_no = 39 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver39(74)	
If version_no = 39 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver39(75)

If version_no = 40 AND error_level = "L" AND tipe  = "8-bit-Byte" Then maxlen = ver40(64)	
If version_no = 40 AND error_level = "M" AND tipe  = "8-bit-Byte" Then maxlen = ver40(65)	
If version_no = 40 AND error_level = "Q" AND tipe  = "8-bit-Byte" Then maxlen = ver40(66)	
If version_no = 40 AND error_level = "H" AND tipe  = "8-bit-Byte" Then maxlen = ver40(67)	
If version_no = 40 AND error_level = "L" AND tipe  = "Numeric" Then maxlen = ver40(68)	
If version_no = 40 AND error_level = "M" AND tipe  = "Numeric" Then maxlen = ver40(69)	
If version_no = 40 AND error_level = "Q" AND tipe  = "Numeric" Then maxlen = ver40(70)	
If version_no = 40 AND error_level = "H" AND tipe  = "Numeric" Then maxlen = ver40(71)	
If version_no = 40 AND error_level = "L" AND tipe  = "Alphanumeric" Then maxlen = ver40(72)	
If version_no = 40 AND error_level = "M" AND tipe  = "Alphanumeric" Then maxlen = ver40(73)	
If version_no = 40 AND error_level = "Q" AND tipe  = "Alphanumeric" Then maxlen = ver40(74)	
If version_no = 40 AND error_level = "H" AND tipe  = "Alphanumeric" Then maxlen = ver40(75)

End Sub

'Store the QR Code code as a PNG file in the pictures folder (File.DirRootExternal, "pictures/Aztec.png")
'and load it into ImageView1 
Private Sub teken1

Dim p1 As Path
Dim b1 As Bitmap
Dim siz1 As Int
siz1 = 8   'was 9 27 May
b1.InitializeMutable(qr_size*siz1+50dip,qr_size*siz1+50dip)      'This will set a border of 50dip around the image
                                                        'siz is the size of the matrix/QR Code code eg 37 x 37 modules  
Dim c As Canvas
Dim x As Int
Dim y As Int

c.initialize2(b1)
Rect1.Initialize(0, 0, 0, 0)
Rect1.Left = 0                                          'Set the size of the drawing area and fill it with WHITE
Rect1.Right = qr_size * siz1 + 50dip        
Rect1.Top = 0            
Rect1.Bottom = qr_size * siz1 + 50dip         

'c.DrawRect(Rect1,bcolor,True,0)

'***********************************************************************
			Dim gd As GradientDrawable
			Dim colours(4) As Int
			colours(0) = Colors.ARGB(255,255,255,255)
			colours(1) = Colors.ARGB(255,255,255,255)
			colours(2) = Colors.ARGB(255,255,255,255)
			colours(3) = Colors.ARGB(255,255,255,255)			
			gd.Initialize("TR_BL",colours)

			Dim r As Reflector
		    Dim style As Int
			style = Rnd(0,3)	
			
			r.Target = gd
'			r.runmethod4("setColors",Array As Object(Array As Int(0xFFFF00FF,0xFFFF4500,0xFF9932CC,0xFF6495ED)), Array As String("[I"))
			r.RunMethod2("setGradientType",style,"java.lang.int")                          '0 = linear, 1 = radial, 2 = sweep/angled/diamond
			r.RunMethod2("setGradientRadius",(qr_size*siz1+50dip)/1.5,"java.lang.float")
            r.RunMethod2("setCornerRadius",20,"java.lang.float")
'            r.RunMethod4("setCornerRadii", Array As Object(Array As Float(20,20,20,20,20,20,20,20)), Array As String("[F"))

			c.DrawDrawable(gd,Rect1)
'***********************************************************************

x = Floor (((qr_size * siz1 + 50dip) - (qr_size * siz1))/2)     'Find the centre points and step back to make it appear centre in the PNG file 
y = Floor (((qr_size * siz1 + 50dip) - (qr_size * siz1))/2)
         

For i =  1 To qr_size
  For j = 1 To qr_size
    If s5cij(i,j) = 10  Then 'OR s5cij(i,j) = 1 Then
		Rect1.Left = x       
		Rect1.Right = x + siz1            
		Rect1.Top = y            
		Rect1.Bottom = y + siz1
'		Rect1.Left = x + 1      
'		Rect1.Right = x + siz1 - 2           
'		Rect1.Top = y + 1            
'		Rect1.Bottom = y + siz1 - 2	   	      	 

		If shp = "s" OR shp = "S" Then
          c.DrawRect(Rect1, fcolor, True, siz1)                  'draw squares
		End If  

	Else If s5cij(i,j) = 9  Then 'OR s5cij(i,j) = 0 Then
		Rect1.Left = x       
		Rect1.Right = x + siz1            
		Rect1.Top = y            
		Rect1.Bottom = y + siz1

	Else                              'only if there is an error in the code that populates the QR code matrix i.e in array "s5cij"
		Rect1.Left = x                'this "Else" statement is pupely for debugging purposes - making sure "s5cij" had been packed correctly  
		Rect1.Right = x + siz1            
		Rect1.Top = y            
		Rect1.Bottom = y + siz1            	 
        c.DrawRect(Rect1, Colors.Red, True, siz1)			
	End If	
	x = x + siz1
  Next
    x = Floor (((qr_size * siz1 + 50dip) - (qr_size * siz1))/2)      
	y = y + siz1                                             
Next 

'See Erel's posting at http://www.basic4ppc.com/android/forum/threads/saving-a-picture.10200/#post-56686 for detail of the code below
'Marginally changed to suite my requirements

Dim out As OutputStream
out = File.OpenOutput(File.DirRootExternal, "/pictures/QRcode.png", False)
b1.WriteToStream(out, 100, "PNG")                                           'Export the generated bitmap to a PNG and save in in the pictures folder
out.flush                                                                   'Make sure nothing is stuck that should have been written
out.Close

'send the intent that asks the media scanner to scan the file
Dim a123 As Intent
a123.Initialize("android.intent.action.MEDIA_SCANNER_SCAN_FILE", _
    "file://" & File.Combine(File.DirRootExternal, "/pictures/QRcode.png"))   '...so that the file is visible....    
Dim p As Phone
p.SendBroadcastIntent(a123)


End Sub


