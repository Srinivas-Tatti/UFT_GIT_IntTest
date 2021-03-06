'-------------------------------------------------------------------------------
'
'	Script:			HTMLReporter
'	Author:			Multiple
'	Date Created:	Thursday 23 April, 2015
'
'	Description:
'		Classes to create an HTML Report after test execution
'
'-------------------------------------------------------------------------------

Option Explicit

Public IMSAppscript_HTML_MasterReport
Public strSummaryReport
Set IMSAppscript_HTML_MasterReport = New CHTMLMasterReport

Class CHTMLMasterReport
	
	''' <value type="Scripting.FileSystemObject"/>
	Private vMasterReportFSO
	''' <value type="String"/>
	Private vSystemName
	''' <value type="DateTime"/>
	Private vStartTime
	''' <value type="DateTime"/>
	Private vEndTime
	''' <value type="String"/>
	Private vFilePath
	''' <value type="Integer"/>
	Private vScenarioCount
	''' <value type="Integer"/>
	Private vScenarioPassedCount
	''' <value type="TextStream"/>
	Private vTextStream	
	''' <value type="String"/>
	Private vBaseFolder
	

	''' <summary>
	''' Gets/Sets the FileSystemObject Associated with the Master Report
	''' </summary>
	''' <value type="Scripting.FileSystemObject"/>
	Public Property Get MasterReportFSO
		Set MasterReportFSO = vMasterReportFSO
	End Property
	Public Property Set MasterReportFSO(ByRef vData)
		Set vMasterReportFSO = vData
	End Property
		
	''' <summary>
	''' Gets/Sets the System Name for the System Under Test
	''' </summary>
	''' <value type="String"/>
	Public Property Get SystemName
		SystemName = vSystemName
	End Property
	Public Property Let SystemName(ByVal vData)
		vSystemName = vData
	End Property
		
	''' <summary>
	''' Gets/Sets the Start Time of the Test
	''' </summary>
	''' <value type="DateTime"/>
	Public Property Get StartTime
		StartTime = vStartTime
	End Property
	Public Property Let StartTime(ByVal vData)
		vStartTime = vData
	End Property
	
	''' <summary>
	''' Gets/Sets the End Time for the Test
	''' </summary>
	''' <value type="DateTime"/>
	Public Property Get EndTime
		EndTime = vEndTime
	End Property
	Public Property Let EndTime(ByVal vData)
		vEndTime = vData
	End Property
	
	''' <summary>
	''' Gets/Sets the FilePath to the HTML Report File
	''' </summary>
	''' <value type="String"/>
	Public Property Get FilePath
		FilePath = vFilePath
	End Property
	Public Property Let FilePath(ByVal vData)
		vFilePath = vData
	End Property
	
	''' <summary>
	''' summaryInfo
	''' </summary>
	''' <value type="String"/>
	Public Property Get BaseFolder
		BaseFolder = vBaseFolder
	End Property
	Public Property Let BaseFolder(ByVal vData)
		vBaseFolder = vData
	End Property
	
	''' <summary>
	''' Returns a string representing the duration of the test event which is recorded in the StartTime and EndTime properties
	''' </summary>
	''' <returns type="String">String in HH:MM:SS Format</returns>
	Public Function TestDuration()
		TestDuration = FormatDuration(StartTime, EndTime)		
	End Function
	
	''' <summary>
	''' Gets the TextStream Object in order to write to the HTML file.
	''' </summary>
	''' <value type="TextStream"/>
	Private Property Get TextStreamObject()
		Set TextStreamObject = vTextStream
	End Property	
	
	''' <summary>
	''' Adds a scenario to the report row of the Master Report
	''' </summary>
	''' <param name="objScenario" type="CTestScenario"></param>

	Public Sub AddScenario(objScenario)
		''' <value type="String"/>
		Dim strHTML, strLinkToScenarioReport
		''' <value type="CTestCase"/>
		Dim objTestCase
		vScenarioCount = vScenarioCount + 1
		
		Dim strStatus
		
		If objScenario.TestCases.Count > 0 Then
			strLinkToScenarioReport = CreateScenarioReport(objScenario)	
		
		End If
		
		Select Case objScenario.Status
			Case StatusTypes.Fail:
				strStatus = "Fail"
			Case StatusTypes.Warning:
				strStatus = "Warning"
				vScenarioPassedCount = vScenarioPassedCount + 1
			Case StatusTypes.Pass:
				strStatus = "Pass"
				vScenarioPassedCount = vScenarioPassedCount + 1
			Case Else:
				strStatus = "Information"
		End Select
		
		TextStreamObject.WriteLine(vbTab & vbTab & vbTab & "<tr Class=Rows>")
		TextStreamObject.WriteLine(vbTab & vbTab & vbTab & vbTab & "<td Class=Cols>" & vScenarioCount & "</td>")
		
		If strLinkToScenarioReport = "" Then
			TextStreamObject.WriteLine(vbTab & vbTab & vbTab & vbTab & "<td Class=Cols>" & objScenario.TestScenarioID & "</td>")
		Else
			TextStreamObject.WriteLine(vbTab & vbTab & vbTab & vbTab & "<td Class=Cols><a href=" & Chr(34) & strLinkToScenarioReport & Chr(34) & ">" & objScenario.TestScenarioID & " </a></td>")
		End If

		TextStreamObject.WriteLine(vbTab & vbTab & vbTab & vbTab & "<td Class=ColsNA>"& objScenario.TestScenarioDescription &"</td>")
		TextStreamObject.WriteLine(vbTab & vbTab & vbTab & vbTab & "<td Class="& LCase(strStatus) &" Class=Cols>"& strStatus &"</td>")
		'TextStreamObject.WriteLine(vbTab & vbTab & vbTab & vbTab & "<td Class=ColsNA>"& FormatDuration(objScenario.StartTime, objScenario.EndTime) &"</td>")
		TextStreamObject.WriteLine(vbTab & vbTab & vbTab & "</tr>")
		
	End Sub
		
	Private Sub AddHeader()
		TextStreamObject.WriteLine(vbTab & "<head>")
		
		TextStreamObject.WriteLine(vbTab & vbTab & "<style>")
		TextStreamObject.WriteLine(vbTab & vbTab & vbTab & ".Header {color: DarkRed;font-size: 13px;font-family: arial;border-collapse: collapse;}")
		TextStreamObject.WriteLine(vbTab & vbTab & vbTab & ".Header.Col2 {background-color: #C0C0C0;font-weight: bold;border-color:#6C6C6C;border-width: 1px 1px 1px 1px;border-style:solid; padding: 2px;}")
		TextStreamObject.WriteLine(vbTab & vbTab & vbTab & ".Header2Margin{margin-top:-85px;margin-bottom:0px;margin-right:0px;margin-left:950px;}")
		TextStreamObject.WriteLine(vbTab & vbTab & vbTab & ".HeaderPadding{padding: 5}")
		TextStreamObject.WriteLine()
		TextStreamObject.WriteLine(vbTab & vbTab & vbTab & ".MainTbl {background-color:#ffffff; margin-top:70px;margin-bottom:0px;margin-right:0px;margin-left:0px;}")
		TextStreamObject.WriteLine(vbTab & vbTab & vbTab & ".MainTbl.RowHeader {background-color:#CC99CC; color: White ; font-size: 12px; font-family: verdana;}")
		TextStreamObject.WriteLine(vbTab & vbTab & vbTab & ".MainTbl.Footer {background-color:#CC99CC; color: White ; font-size: 12px; font-family: verdana;text-align:Left }")
		TextStreamObject.WriteLine(vbTab & vbTab & vbTab & ".MainTbl.subSection {background-color:#CCCCFF}")
		TextStreamObject.WriteLine(vbTab & vbTab & vbTab & ".MainTbl.subSection.Col {padding: 3; color: DarkBlue; font-size:12px; font-family: verdana;}")
		TextStreamObject.WriteLine(vbTab & vbTab & vbTab & ".MainTbl.Rows {background-color:#F0F0F8}")
		TextStreamObject.WriteLine(vbTab & vbTab & vbTab & ".MainTbl.Rows.Cols {padding: 2; color: DarkBlue; font-size: 11px; font-family: verdana;text-align:center}")
		TextStreamObject.WriteLine(vbTab & vbTab & vbTab & ".MainTbl.Rows.ColsNA {padding: 2; color: DarkBlue; font-size: 11px; font-family: verdana;text-align:Left}")		
		TextStreamObject.WriteLine()
		TextStreamObject.WriteLine(vbTab & vbTab & vbTab & ".pass{padding: 2; color: green; font-size: 13px; font-family: verdana;text-align:center}")
		TextStreamObject.WriteLine(vbTab & vbTab & vbTab & ".fail{padding: 2; color: red; font-size: 13px; font-family: verdana;text-align:center}")
		TextStreamObject.WriteLine(vbTab & vbTab & vbTab & ".warning{padding: 2; color:#FF8C00; font-size: 13px; font-family: verdana;text-align:center}")
		TextStreamObject.WriteLine(vbTab & vbTab & "</style>")
		
		
		TextStreamObject.WriteLine(vbTab & "</head>")
	End Sub
	
	Private Sub AddBody()
			TextStreamObject.WriteLine(vbTab & "<body bgcolor='#dddddd'  link='#0000ff' vlink='#FF9900' alink='#993366'>")
			TextStreamObject.WriteLine(vbTab & vbTab & "<table width='99%'>")			
			TextStreamObject.WriteLine(vbTab & vbTab & vbTab & "<tr><td width='100%' align='left'><img src='"&Environment.Value("LogoPath")&"' height=124 width=272></td></td></tr>")
			TextStreamObject.WriteLine(vbTab & vbTab & "</table>")
			TextStreamObject.WriteLine(vbTab & vbTab & "<br>")
			TextStreamObject.WriteLine(vbTab & vbTab & "<table Class=Header width='30%' bgcolor='#81F7F3' cellspacing='1' colspan='2' cellpadding='2' border='1'>")
			TextStreamObject.WriteLine(vbTab & vbTab & vbTab & "<tr>")
			TextStreamObject.WriteLine(vbTab & vbTab & vbTab & vbTab & "<td width='30%'>Application</td>")
			TextStreamObject.WriteLine(vbTab & vbTab & vbTab & vbTab & "<td Class=Col2 width='70%'>" & Environment.Value("Application") & "</td>")
			TextStreamObject.WriteLine(vbTab & vbTab & vbTab & "</tr>")
			TextStreamObject.WriteLine(vbTab & vbTab & vbTab & "<tr>")
			TextStreamObject.WriteLine(vbTab & vbTab & vbTab & vbTab & "<td width='30%'>Environment</td>")
			TextStreamObject.WriteLine(vbTab & vbTab & vbTab & vbTab & "<td Class=Col2 width='70%'>" & Environment.Value("TestEnv") & "</td>")
			TextStreamObject.WriteLine(vbTab & vbTab & vbTab & "</tr>")
			TextStreamObject.WriteLine(vbTab & vbTab & vbTab & "<tr>")
			TextStreamObject.WriteLine(vbTab & vbTab & vbTab & vbTab & "<td width='30%'>Execution Start Time</td>")
			TextStreamObject.WriteLine(vbTab & vbTab & vbTab & vbTab & "<td Class=Col2 width='70%'>" & FormatDateTime(Environment.Value("StartTime"),vbGeneralDate) & "</td>")
			TextStreamObject.WriteLine(vbTab & vbTab & vbTab & "</tr>")
			TextStreamObject.WriteLine(vbTab & vbTab & vbTab & "<tr>")
			TextStreamObject.WriteLine(vbTab & vbTab & vbTab & vbTab & "<td width='30%'>Execution End Time</td>")			
			TextStreamObject.WriteLine(vbTab & vbTab & vbTab & vbTab & "<td Class=Col2 width='70%'>" & FormatDateTime(Environment.Value("EndTime"),vbGeneralDate) & "</td>")
			TextStreamObject.WriteLine(vbTab & vbTab & vbTab & "</tr>")			
			'<Shweta - 15/11/2016 > Add Total Execution Element to Results - Start
			TextStreamObject.WriteLine(vbTab & vbTab & vbTab & "<tr>")
			TextStreamObject.WriteLine(vbTab & vbTab & vbTab & vbTab & "<td width='30%'>Total Execution Time (Hours:Minutes:Seconds)</td>")			
			TextStreamObject.WriteLine(vbTab & vbTab & vbTab & vbTab & "<td Class=Col2 width='70%'>" &Environment.Value("TotalExecutionTime")& "</td>")
			TextStreamObject.WriteLine(vbTab & vbTab & vbTab & "</tr>")			
			'<Shweta - 15/11/2016 > Add Total Execution Element to Results - End
			TextStreamObject.WriteLine(vbTab & vbTab & "</table>")
			TextStreamObject.WriteLine(vbTab & vbTab & "<br>")			
			TextStreamObject.WriteLine(vbTab & vbTab & "<table Class=MainTbl cellspacing='1' colspan='5' cellpadding='5' border='1' Width='100%'>")
			TextStreamObject.WriteLine(vbTab & vbTab &  vbTab & "<tr Class=RowHeader>")
			TextStreamObject.WriteLine(vbTab & vbTab &  vbTab & vbTab & "<th Class=HeaderPadding Width='5%'>S No.</th>")
			TextStreamObject.WriteLine(vbTab & vbTab &  vbTab & vbTab & "<th Class=HeaderPadding Width='15%'>Test Scenario#</th>")
			TextStreamObject.WriteLine(vbTab & vbTab &  vbTab & vbTab & "<th Class=HeaderPadding Width='60%'>Test Scenario Description</th>")
			TextStreamObject.WriteLine(vbTab & vbTab &  vbTab & vbTab & "<th Class=HeaderPadding Width='10%'>Status</th>")
			'TextStreamObject.WriteLine(vbTab & vbTab &  vbTab & vbTab & "<th Class=HeaderPadding Width='10%'>Duration<br/>in<br/>(HH:MM:SS)</th>")
			TextStreamObject.WriteLine(vbTab & vbTab &  vbTab & "</tr>")
			
			''' <value type="CTestScenario"/>
			Dim tsenarios()
			i=0
			For Each objScenario In IMS_Framework_TestCaseHandler.TestScenarios.Items
			
			
			ReDim Preserve tsenarios(i)
			tsenarios(i)=objScenario.TestScenarioID
			i=i+1
			Next
			
			
			For Each objScenario In IMS_Framework_TestCaseHandler.TestScenarios.Items
				If tsenarios(Environment("TC"))=objScenario.TestScenarioID Then
					Environment("TC")=Environment("TC")+1
					AddScenario(objScenario)
					Exit For
				End If
			Next
			
			'TextStreamObject.WriteLine(vbTab & vbTab & "</table>")			
			
			'TextStreamObject.WriteLine(vbTab & "</body>")
			
	End Sub
	
	Private Sub uppendBody()
			'TextStreamObject.WriteLine(vbTab & vbTab & "<br>")
			'TextStreamObject.WriteLine("<table Class=MainTbl cellspacing='1' colspan='1' cellpadding='1' border='1' Width='100%'>")
			'TextStreamObject.WriteLine(vbTab & vbTab &  vbTab & "<tr Class=RowHeader>")
	
	'TextStreamObject.WriteLine(vbTab & vbTab &  vbTab & "</tr>")
	Dim tsenarios()
			i=0
			For Each objScenario In IMS_Framework_TestCaseHandler.TestScenarios.Items
			
			
			ReDim Preserve tsenarios(i)
			tsenarios(i)=objScenario.TestScenarioID
			i=i+1
			Next
			
			
			For Each objScenario In IMS_Framework_TestCaseHandler.TestScenarios.Items
				If tsenarios(Environment("TC"))=objScenario.TestScenarioID Then
					Environment("TC")=Environment("TC")+1
					AddScenario(objScenario)
					Exit For
				End If
			Next
			
			'TextStreamObject.WriteLine(vbTab & vbTab & "</table>")			
			
			'TextStreamObject.WriteLine(vbTab & "</body>")
	End Sub
	''' <summary>
	''' Creates a Scenario Sub-Report and returns the link to the SubReport
	''' </summary>
	''' <param name="objScenario" type="CTestScenario">Scenario to create the Sub Report</param>
	''' <returns type="String">Path to the Sub Report</returns>
	Private Function CreateScenarioReport(objScenario)
		''' <value type="CScenarioReport"/>
		Dim objScenarioReport
		Set objScenarioReport = New CScenarioReport
		CreateScenarioReport = objScenarioReport.CreateScenarioReport(BaseFolder, objScenario)
		Set objScenarioReport = Nothing
	End Function
	
	Public Function CreateMasterReport()
						
		If (Not objUtils.fnFolderExist( Environment.Value("TestResultsFolderPath") )) Then
			Call objUtils.fnCreateFolder(Environment.Value("ResultsPath"), Environment.Value("TestResultsFolderName"))
		End If
		
		BaseFolder = Environment.Value("TestResultsFolderPath")
		
		' FilePath = BaseFolder & Environment.Value("ApplicationShortName") & "_Run_" & FormatFileTime(StartTime, "_") & ".html"
		FilePath = BaseFolder & Environment.Value("TestName") & "_Run_" & FormatFileTime(StartTime, "_") & ".html" 
		 

		strSummaryReport = FilePath
		
		If MasterReportFSO.FileExists(FilePath)<> True Then
		Set vTextStream = MasterReportFSO.CreateTextFile(FilePath, True, True)
		TextStreamObject.WriteLine("<html>")
		AddHeader()
		AddBody()
		TextStreamObject.WriteLine("</html>")
		else
		Set vTextStream = MasterReportFSO.OpenTextFile(FilePath, 8, True, True)
		AddHeader()
		uppendBody()
		TextStreamObject.WriteLine("</html>")
		End If
		

		
		'TODO: Copy Images to Correct Location
		'CopyImagesToImageFolder(BaseFolder)
		
		


		
		'Copies the New report to the report folder
		'Call objUtils.fnCopyFolder(Environment.Value("QAStockReport"), BaseFolder, False)
		'TODO: Remove the HTML Reporter and just include the file copy script once report v2 is satisfactory
		'Dim oFSO
		'Set oFSO = CreateObject("Scripting.FileSystemObject")
		'Call oFSO.CopyFile(Environment.Value("QAStockReport") & "\*.*", Environment.Value("TestResultsFolderPath"), False)
		
		
		CreateMasterReport = FilePath
		TextStreamObject.Close
	End Function
	
	''' <summary>
	''' Launches the HTML Report
	''' </summary>
	Public Sub ShowHTMLReport( )
		
		systemutil.Run  FilePath
		'systemUtil.Run Environment.Value("TestResultsFolderPath") & "Results.html"
	End Sub
	
	
	Sub CopyImagesToImageFolder(Path_Name)
		''' <value type="Scripting.FileSystemObject"/>
		Dim fso
		''' <value type="String"/>
		Dim strPath
		
		Set fso = CreateObject("Scripting.FileSystemObject")

		If (objUtils.fnFolderExist(Path_Name) And objUtils.fnFileExist(Environment.Value("LogoPath"))) Then

			If Not objUtils.fnFolderExist(Path_Name & "Images\") Then
				strPath = objUtils.fnCreateFolder(Path_Name, "Images\")
			End If
			'TODO: Create objUtils.fnCopyFile
			Call fso.CopyFile(Environment.Value("LogoPath"), strPath & "IMS.jpg")
	   End If
	End Sub
	
	Private Sub Class_Initialize()
		Set MasterReportFSO = CreateObject("Scripting.FileSystemObject")
		StartTime = Now
		EndTime = Now
		SystemName = ""
		vScenarioCount = 0
		vScenarioPassedCount = 0
	End Sub
	
End Class

''' <summary>
''' Takes the difference between a Start DateTime and an End DateTime and parses it into a string format
''' </summary>
''' <param name="startDateTime" type="DateTime">Start DateTime</param>
''' <param name="endDateTime" type="DateTime">End DateTime</param>
''' <returns type="String">The difference between the Start and End DateTimes in string format: HH:MM:SS</returns>
Public Function FormatDuration(startDateTime, endDateTime)
	
	''' <value type="Integer"/>
	Dim iSeconds
	
	iSeconds = DateDiff("d",startDateTime, endDateTime)
	
	If iSeconds < 0 Then
	     FormatDuration = "Error"
		 Exit Function
	End If
	
	''' <value type="Integer"/>
	Dim vHours, vMinutes, vSeconds
	''' <value type="String"/>
	Dim strDuration
	
	vHours = Int(iSeconds / 3600)
	vMinutes = Int((iSeconds Mod 3600 ) / 60 )
	vSeconds = Int((iSeconds Mod 3600 ) Mod 60 )
			
	strDuration = vHours
		
	If vMinutes < 10 Then
	    strDuration = strDuration & ":" & "0" & vMinutes
	Else
	    strDuration = strDuration & ":" & vMinutes
	End If
	
	If vSeconds < 10 Then
	    strDuration = strDuration & ":" & "0" & vSeconds
	Else
		strDuration = strDuration & ":" & vSeconds
	End If
	
	FormatDuration = strDuration
End Function

''' <summary>
''' Takes a DateTime and convertes it into a string format: HH_MM_SS
''' </summary>
''' <param name="dttm" type="DateTime">DateTime that needs to be formatted</param>
''' <param name="chrDelimeter" type="String"></param>
''' <returns type="String">String In Format: HH_MM_SS where "_" is the </returns>
Public Function FormatFileTime(dttm, chrDelimeter)

	''' <value type="String"/>
	Dim strTime, strAMPM 
	strTime = FormatDateTime(dttm,vbLongTime)
	strAMPM = Right(strTime,2)
	
	''' <value type="Array"/>
	Dim a
	
	a = Split(Left(strTime,Len(strTime)-3),":")
	
	''' <value type="Integer"/>
	Dim vHours, vMinutes, vSeconds
	''' <value type="String"/>
	Dim strDuration
	
	
	vHours = 	CInt(a(0))
	vMinutes = 	CInt(a(1))
	vSeconds = 	CInt(a(2))
	
	If strAMPM = "PM"  And vHours < 12 Then
		vHours = vHours + 12
	End If
	
	If strAMPM = "AM" And vHours = 12 Then
		vHours = 0
	End If
	
	If vHours < 10 Then
		strDuration = "0" & vHours
	Else
		strDuration = vHours
	End If
		
	If vMinutes < 10 Then
	    strDuration = strDuration & chrDelimeter & "0" & vMinutes
	Else
	    strDuration = strDuration & chrDelimeter & vMinutes
	End If
	
	If vSeconds < 10 Then
	    strDuration = strDuration & chrDelimeter & "0" & vSeconds
	Else
		strDuration = strDuration & chrDelimeter & vSeconds
	End If
	
	FormatFileTime = strDuration
End Function

Class CScenarioReport
	''' <value type="CTestScenario"/>
	Private vScenario
	''' <value type="Scripting.FileSystemObject"/>
	Private vReportFSO	
	''' <value type="TextStream"/>
	Private vTextStreamObject	
	''' <value type="String"/>
	Private vFilePath
	''' <value type="String"/>
	Private vBaseFolder	
	
	''' <summary>
	''' Get/Set the Test Scenario Object
	''' </summary>
	''' <value type="CTestScenario"/>
	Public Property Get Scenario
		Set Scenario = vScenario
	End Property
	Public Property Set Scenario(ByRef vData)
		Set vScenario = vData
	End Property
	
	''' <summary>
	''' Gets/Sets the FilePath String for the Scenario Report
	''' </summary>
	''' <value type="String"/>
	Public Property Get FilePath
		FilePath = vFilePath
	End Property
	Public Property Let FilePath(ByVal vData)
		vFilePath = vData
	End Property
	
	''' <summary>
	''' Gets the Report FileSystemObject
	''' </summary>
	''' <value type="Scripting.FileSystemObject"/>
	Private Property Get ReportFSO
		Set ReportFSO = vReportFSO
	End Property
	''' <summary>
	''' Gets the Text Stream Object of the Report
	''' </summary>
	''' <value type="TextStream"/>
	Private Property Get TextStreamObject
		Set TextStreamObject = vTextStreamObject
	End Property

	''' <summary>
	''' Gets/Sets the BaseFolder of the Report
	''' </summary>
	''' <value type="String"/>
	Public Property Get BaseFolder
		BaseFolder = vBaseFolder
	End Property
	Public Property Let BaseFolder(ByVal vData)
		vBaseFolder = vData
	End Property
	

	''' <summary>
	''' Adds HTML Information to the report
	''' </summary>
	''' <param name="strHTML" type="String">HTML to add to the Report</param>
	Private Sub InjectHTML(strHTML)
		TextStreamObject.WriteLine(strHTML)
	End Sub
		
	Private Sub AddHeader()		
		TextStreamObject.WriteLine(vbTab & "<head>")
		TextStreamObject.WriteLine(vbTab & vbTab & "<title>QTP Result File " & Scenario.TestScenarioID & "</title>")
		TextStreamObject.WriteLine(vbTab & vbTab & "<style>")
		TextStreamObject.WriteLine(vbTab & vbTab & vbTab & ".pass {font-family:verdana;color:green;font-size:11px}")
		TextStreamObject.WriteLine(vbTab & vbTab & vbTab & ".fail {font-family:verdana;color:red;font-size:12px}")
		TextStreamObject.WriteLine(vbTab & vbTab & vbTab & ".warning {font-family:verdana;color:#FF8C00;font-size:12px}")
		TextStreamObject.WriteLine(vbTab & vbTab & vbTab & ".done {font-family:verdana;color:#000000 ;font-size:12px}")
		TextStreamObject.WriteLine(vbTab & vbTab & vbTab & ".info {font-family:verdana;color:Navy  ;font-size:12px}")
		TextStreamObject.WriteLine(vbTab & vbTab & "</style>")
		TextStreamObject.WriteLine(vbTab & "</head>")
	End Sub
	
	Private Sub AddBody()

		TextStreamObject.WriteLine(	vbTab & "<body bgcolor='#dddddd'>")
		
		If Scenario.TestCases.Count = 0 Then
			TextStreamObject.WriteLine(	vbTab & "</body>")
			Exit Sub
		End If	
		
		TextStreamObject.WriteLine(	vbTab & vbTab & vbTab & "<tr bgcolor='#C0C0C0'><td><font size = 4 face = 'arial' color='DarkRed'> <B>Test Scenario:&nbsp" & Scenario.TestScenarioID & "</B></font></td></tr>")
		TextStreamObject.WriteLine(	vbTab & vbTab & "</br>")
		TextStreamObject.WriteLine(	vbTab & vbTab & "<table>")
		TextStreamObject.WriteLine( vbTab & vbTab & vbTab & "<style type=text/css><!--#layer1{position:relative;left:5px;top:13px;width:40px;height:10px;background-color:green;z-index:0;}body{background-color: #dddddd;}--></style><div id=layer1></div>")
		TextStreamObject.WriteLine(	vbTab & vbTab & "</br>")
		TextStreamObject.WriteLine( vbTab & vbTab & vbTab & "<style type=text/css><!--#layer2{position:relative;left:5px;top:13px;width:40px;height:10px;background-color:red;z-index:0;}body{background-color: #dddddd;}--></style><div id=layer2></div>")
		TextStreamObject.WriteLine(	vbTab & vbTab & "</br>")
		TextStreamObject.WriteLine(	vbTab & vbTab & vbTab & "<style type=text/css><!--#layer3{position:relative;left:5px;top:13px;width:40px;height:10px;background-color:Navy;z-index:0;}body{background-color: #dddddd;}--></style><div id=layer3></div>")
		TextStreamObject.WriteLine(	vbTab & vbTab & "</br>")
		TextStreamObject.WriteLine(	vbTab & vbTab & vbTab & "<span style=position:absolute;top:38px;left:60px;color:green;><b>Pass</b></span>")
		TextStreamObject.WriteLine(	vbTab & vbTab & vbTab & "<span style=position:absolute;top:67px;left:60px;color:red><b>Fail</b></span>")
		TextStreamObject.WriteLine(	vbTab & vbTab & vbTab & "<span style=position:absolute;top:96px;left:60px;color:Navy;><b>Information</b></span>")
		TextStreamObject.WriteLine(	vbTab & vbTab & "<table>")
		TextStreamObject.WriteLine(	vbTab & vbTab & "</br>")
		TextStreamObject.WriteLine(	vbTab & vbTab & "</table>")

		TextStreamObject.WriteLine(vbTab & vbTab & vbTab & vbTab & "<td Class=Cols><a href=" &strSummaryReport& "> Click here to go summary page </a></td>")
		TextStreamObject.WriteLine(	vbTab & vbTab & "</br>")
		TextStreamObject.WriteLine(	vbTab & vbTab & "</br>")
		
		TextStreamObject.WriteLine(	vbTab & vbTab & vbTab & "<table bgcolor='#F2F5A9' cellspacing='0' colspan='5' cellpadding='5' border='1' Width='100%'>")
		TextStreamObject.WriteLine(	vbTab & vbTab & vbTab & vbTab & "<tr Class=RowHeader>")
		TextStreamObject.WriteLine(	vbTab & vbTab & vbTab & vbTab & vbTab & "<th Class=HeaderPadding Width='8%'>Step No.</th>")
		TextStreamObject.WriteLine(	vbTab & vbTab & vbTab & vbTab & vbTab & "<th Class=HeaderPadding Width='50%'>Description</th>")
		TextStreamObject.WriteLine(	vbTab & vbTab & vbTab & vbTab & vbTab & "<th Class=HeaderPadding Width='25%'>Location</th>")
		TextStreamObject.WriteLine(	vbTab & vbTab & vbTab & vbTab & vbTab & "<th Class=HeaderPadding Width='5%'>Status</th>")
		TextStreamObject.WriteLine(	vbTab & vbTab & vbTab & vbTab & vbTab & "<th Class=HeaderPadding Width='15%'>Time</th>")
		TextStreamObject.WriteLine(	vbTab & vbTab & vbTab & vbTab & "</tr>")
		
		''' <value type="CTestCase"/>
		Dim objTestCase
		''' <value type="CTestCaseStep"/>
		Dim objTestStep
		
		For Each objTestCase In Scenario.TestCases.Items
			''' <value type="Integer"/>
			Dim vStepCounter

			vStepCounter = 0
			
			TextStreamObject.WriteLine(	vbTab & vbTab & vbTab & vbTab & "<tr bgcolor='#F0F0F8'>")
			TextStreamObject.WriteLine(	vbTab & vbTab & vbTab & vbTab & vbTab & "<td class='info'><Br></td>")
			TextStreamObject.WriteLine(	vbTab & vbTab & vbTab & vbTab & vbTab & "<td class='info'>Test Case : " & objTestCase.TestCaseID & "</td>")
			TextStreamObject.WriteLine(	vbTab & vbTab & vbTab & vbTab & vbTab & "<td class='info'>Test Case</td>")
			TextStreamObject.WriteLine(	vbTab & vbTab & vbTab & vbTab & vbTab & "<td class='info'><B>INFORMATION </B>  </td>")
			TextStreamObject.WriteLine(	vbTab & vbTab & vbTab & vbTab & vbTab & "<td class='info'>" & FormatDateTime(objTestCase.StartTime, vbGeneralDate) & "</td>")
			TextStreamObject.WriteLine(	vbTab & vbTab & vbTab & vbTab & "</tr>")
			
		 
			For Each objTestStep In objTestCase.TestCaseSteps.Items
				''' <value type="String"/>
				Dim strClassValue

				vStepCounter = vStepCounter + 1				
				
				Select Case objTestStep.StepStatus
					Case StatusTypes.Pass:
						strClassValue = "pass"
					Case StatusTypes.Warning:
						strClassValue = "warning"
					Case StatusTypes.Fail:
						strClassValue = "fail"
					Case Else:
						strClassValue = "info"
				End Select
				
				TextStreamObject.WriteLine(	vbTab & vbTab & vbTab & vbTab & "<tr bgcolor='#F0F0F8'>")
				TextStreamObject.WriteLine(	vbTab & vbTab & vbTab & vbTab & vbTab & "<td class='"& strClassValue &"'>" & vStepCounter & "</td>")
				TextStreamObject.WriteLine(	vbTab & vbTab & vbTab & vbTab & vbTab & "<td class='"& strClassValue &"'>" & objTestStep.ActualResult & "</td>")
				TextStreamObject.WriteLine(	vbTab & vbTab & vbTab & vbTab & vbTab & "<td class='"& strClassValue &"'>" & objTestStep.Location & "</td>")
				TextStreamObject.WriteLine(	vbTab & vbTab & vbTab & vbTab & vbTab & "<td class='"& strClassValue &"'>" & UCase(strClassValue) & "</td>")
				TextStreamObject.WriteLine(	vbTab & vbTab & vbTab & vbTab & vbTab & "<td class='"& strClassValue &"'>" & FormatDateTime(objTestStep.TimeStamp, vbGeneralDate) & "</td>")
				TextStreamObject.WriteLine(	vbTab & vbTab & vbTab & vbTab & "</tr>")
			Next
		Next
		
		TextStreamObject.WriteLine(	vbTab & "</body>")
		
	End Sub
	
	''' <summary>
	''' Creates a Scenario Report HTML File based on the supplied Scenario Object
	''' </summary>
	''' <param name="strBaseFolderPath" type="String">Path to the Base Folder</param>
	''' <param name="objScenario" type="CTestScenario">The scenario to base the report on</param>
	''' <returns type="String">Path to the newly created HTML Report</returns>
	Public Function CreateScenarioReport(strBaseFolderPath, objScenario)
		
		Set Scenario = objScenario

		
		If (Not objUtils.fnFolderExist(strBaseFolderPath & "Scenarios\")) Then
			BaseFolder = objUtils.fnCreateFolder(strBaseFolderPath, "Scenarios\")
		Else
			BaseFolder = strBaseFolderPath & "Scenarios\"
		End If
		
'		FilePath = "./" & "Scenarios\Scenario_"& Scenario.TestScenarioID & "_" & FormatFileTime(Scenario.StartTime, "_") & ".html"
		
		FilePath = "\\?\" & BaseFolder & "Scenario_"& Scenario.TestScenarioID & "_" & FormatFileTime(Scenario.StartTime, "_") & ".html"
		
			
		Set vTextStreamObject = ReportFSO.CreateTextFile(FilePath,True,True)
		
		FilePath = "./" & "Scenarios\Scenario_"& Scenario.TestScenarioID & "_" & FormatFileTime(Scenario.StartTime, "_") & ".html"
		
		TextStreamObject.WriteLine("<html>")
		
		AddHeader()	
		AddBody()
		
		TextStreamObject.WriteLine("</html>")
		
		TextStreamObject.Close()
		
		CreateScenarioReport = FilePath 
	End Function
	
	Private Sub Class_Initialize()
		Set Scenario = Nothing
		Set vReportFSO = CreateObject("Scripting.FileSystemObject")
	End Sub
End Class



'=========================================================================================================================================================
'FunctionName	 :		Func_Updatesuitexcelresults
'Purpose		 :      	To update the excel with the results file
'ReturnValue	 :		Null
'Parameters		 : 		Null
'Author          :		-
'Created Date	 :		24th Aug 2011
'Modified By	 :		Madhu
'Modified Time   :		-
'Reason for Modification:		Added the Funtion Header
'=========================================================================================================================================================
'TODO: Find a Home for This
'Update to Frameworkhandler
Public Sub Func_Updatesuitexcelresults() 
	
	Dim intSheetRow, xlApp, TDSheet,i
	
    '@ No of rows in the given Datatable source sheet
    intSheetRow= Datatable.GetSheet("TestScenarios").GetRowCount
    Set xlApp = CreateObject("Excel.Application")
    Set TDSheet=xlAPP.workbooks.open(Environment.Value("ControlSheetPath"))
    
    For i=1 To intSheetRow+1
		  Datatable.GetSheet("TestScenarios").SetCurrentRow(i)	
		  TDSheet.Sheets("TestScenarios").Cells(i+1,8).value=Datatable("Status","TestScenarios")	

		  If Datatable("Status","TestScenarios")="fail" Then
				TDSheet.Sheets("TestScenarios").Cells(i+1,7).value="YES"
			Elseif Datatable("Status","TestScenarios")="Pass" Then
				TDSheet.Sheets("TestScenarios").Cells(i+1,7).value="NO"
		  End If

		  If Datatable("Status","TestScenarios")="pass" Then
				TDSheet.Sheets("TestScenarios").Cells(i+1,8).Interior.Color=RGB(193, 255,193)
			ElseIf Datatable("Status","TestScenarios")="No Run" Then
				TDSheet.Sheets("TestScenarios").Cells(i+1,8).Interior.Color=RGB(185,211,238)
			ElseIf Datatable("Status","TestScenarios")="fail" Then
				TDSheet.Sheets("TestScenarios").Cells(i+1,8).Interior.Color=RGB(255, 106,106)
		  End If
    Next

	TDSheet.Application.ActiveWorkbook.save
    TDSheet.Application.ActiveWorkbook.Close

    xlApp.Quit
    Set TDSheet = Nothing 
    Set xlApp=Nothing
End Sub
