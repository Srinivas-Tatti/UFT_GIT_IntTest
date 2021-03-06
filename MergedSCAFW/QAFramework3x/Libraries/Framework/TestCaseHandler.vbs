'-------------------------------------------------------------------------------
'
'	Script:			TestCaseHandler
'	Author:			Multiple
'	Date Created:	Thursday 23 April, 2015
'
'	Description:
'		Contains Classes that perform handling and execution of Test Cases
'
'-------------------------------------------------------------------------------

Option Explicit

''' <summary>
''' Global Instance of the CTestCaseHandler
''' </summary>
''' <value type="CTestCaseHandler"/>
Public IMS_Framework_TestCaseHandler
''' <summary>
''' Global Instance of the CStepEnum class
''' </summary>
''' <value type="CStepEnum"/>
Public StatusTypes


Set IMS_Framework_TestCaseHandler = New CTestCaseHandler
Set StatusTypes = New CStepEnum


''' <summary>
''' Adds a Step Event to the HTML Report for the current running Test Case
''' </summary>
''' <param name="vStepStatus" type="Integer">An Integer from the StatusTypes Enum {CStepEnum}</param>
''' <param name="strActualResult" type="String">Actual Result of the Step</param>
''' <param name="strLocation" type="String">Location the step occurs. E.G. Function | Screen Field | etc.</param>
Public Sub ReportStep(vStepStatus, strActualResult, strLocation)
	'Shweta <24/8/2016> - Start
	On Error Resume Next
	Dim Err
	'Shweta <24/8/2016> - End
	
	Call IMS_Framework_TestCaseHandler.CurrentTestCase.AddStep(vStepStatus,strActualResult,strLocation,Now)
	
	if vStepStatus=4 Then 
	  Reporter.ReportEvent micFail,strLocation,strActualResult
	  
	  'Shweta <24/8/2016> - 1) Added Err Number and Err description 2) Exiting from Action when any failure step is executed- Start
	  '1) Added Err Number and Err description 
	  If Err.Number <> 0 Then
		  Reporter.ReportEvent micFail,"Error found in location: "& Err.Number&" And Error Description is: " &Err.Description
		  'ReportStep StatusTypes.Fail, "Error number is: "&Err.number, "Error Description is: "&Err.Description
	  	  Err.Clear 	
	      On Error GoTo 0
	  'Else
	  '	  Reporter.ReportEvent micInfo,"Error found in location: "& Err.Number," And Error Description is: " &Err.Description
	  End If
	
'	  '2) Exiting from Action when any failure step is executed
'	  '15/11/2016 - Start
'	  Reporter.ReportEvent micInfo,"Exit from current Action "&Environment.Value("TestName")&" execution", "Exited from Action "&Environment.Value("TestName")&" Execution when validating "&strActualResult&" in "&strLocation
'	  Call IMS_Framework_TestCaseHandler.SerializeResultsToXML()
'	  Call IMSAppscript_HTML_MasterReport.CreateMasterReport()
'	  ExitAction
'	  '15/11/2016 - End
	  'Shweta <24/8/2016> - Added Err Number and Err description 2) Exiting from Action when any failure step is executed - End
	else
	   Reporter.ReportEvent micpass,strLocation,strActualResult
	End if
		
End Sub




Class CTestCaseHandler
	
	''' <value type="CTestScenario"/>
	Private vCurrScenario
	''' <value type="CTestCase"/>
	Private vCurrTestCase
	''' <value type="Scripting.Dictionary"/>
	Private vTestScenarios
	
	''' <summary>
	''' Gets/Sets the Current Scenario string. E.G. "AD_S001"
	''' </summary>
	''' <value type="CTestScenario"/>
	Public Property Get CurrentScenario
		Set CurrentScenario = vCurrScenario
	End Property
	Public Property Set CurrentScenario(ByVal vData)
		Set vCurrScenario = vData
	End Property
	
	''' <summary>
	''' Gets/Sets the CurrentTestCase string. E.G. "AD_S001_T001"
	''' </summary>
	''' <value type="CTestCase"/>
	Public Property Get CurrentTestCase
		Set CurrentTestCase = vCurrTestCase
	End Property
	Public Property Set CurrentTestCase(ByVal vData)
		Set vCurrTestCase = vData
	End Property
	
	''' <summary>
	''' Read Only property of the TestScenarios Data Sheet name
	''' </summary>
	''' <value type="String"/>
	Public Property Get TestScenariosSheetName
		TestScenariosSheetName = "TestScenarios"
	End Property
	
	''' <summary>
	''' Read Only property of TestCases sheet name
	''' </summary>
	''' <value type="String"/>
	Public Property Get TestCasesSheetName
		TestCasesSheetName = "TestCases"
	End Property
	

	''' <summary>
	''' Gets the Test Scenarios Dictionary Object
	''' </summary>
	''' <value type="Scripting.Dictionary"/>
	Public Property Get TestScenarios
		Set TestScenarios = vTestScenarios
	End Property
	
	''' <summary>
	''' Gets the Status of the Test Execution {CStatusEnum}
	''' </summary>
	''' <value type="Integer"/>
	Private Property Get Status
		''' <value type="Integer"/>
		Dim vFailCount, vWarnCount
		vFailCount = 0
		vWarnCount = 0
		
		''' <value type="CTestScenario"/>
		Dim objScenarioStatus
	
		For Each objScenarioStatus In TestScenarios.Items()
		     Select Case objScenarioStatus.Status
				Case StatusTypes.Fail:
					vFailCount = vFailCount + 1
				Case StatusTypes.Warning:
					vWarnCount = vWarnCount + 1
				Case Else:
				
			End Select
		Next
		
		If vFailCount <> 0 Then
			Status = StatusTypes.Fail
		Else
			If vWarnCount <> 0 Then
				Status = StatusTypes.Warning
			Else
				Status = StatusTypes.Pass
			End If
		End If		
	End Property	
	
	Private Sub Class_Initialize()
		Set vTestScenarios = CreateObject("Scripting.Dictionary")
		vTestScenarios.RemoveAll()
		Set vCurrScenario = Nothing
		Set vCurrTestCase = Nothing
	End Sub
	
	''' <summary>
	''' Loads Scenarios from Suite Data
	''' </summary>
	Private Sub LoadScenariosFromTestScenariosData()
		''' <value type="Integer"/>
		Dim i, SCENARIO_ROW_START, SCENARIO_COLUMN_ID, SCENARIO_COLUMN_DESCRIPTION, SCENARIO_COLUMN_EXECFLAG
		
		SCENARIO_ROW_START			= 1
		SCENARIO_COLUMN_ID			= 4
		SCENARIO_COLUMN_DESCRIPTION	= 5
		SCENARIO_COLUMN_EXECFLAG	= 7
		
		i = SCENARIO_ROW_START
   Datatable.GetSheet(TestScenariosSheetName).SetCurrentRow(i)
		
		

		While (Trim(DataTable.Value(SCENARIO_COLUMN_ID,TestScenariosSheetName))<> "")			
			If (UCase(Trim(DataTable.Value(SCENARIO_COLUMN_ID,TestScenariosSheetName))) = UCASE(Environment("TestName"))) Then
    			''' <value type="CTestScenario"/>
    			Dim scenarioObj
    			Set scenarioObj = New CTestScenario
				
    			scenarioObj.TestScenarioID = CStr(Trim(DataTable.Value(SCENARIO_COLUMN_ID,TestScenariosSheetName)))
    			scenarioObj.TestScenarioDescription = CStr(Trim(DataTable.Value(SCENARIO_COLUMN_DESCRIPTION,TestScenariosSheetName)))
    			Call vTestScenarios.Add(scenarioObj.TestScenarioID,scenarioObj)
				
    			Set scenarioObj = Nothing
    			Exit sub
				
			End If

			i = i + 1
    		DataTable.GetSheet(TestScenariosSheetName).SetCurrentRow(i) 
    		
		Wend
		
	End Sub
	
	''' <summary>
	''' Loads Test Case Data from DataTable
	''' </summary>
	Private Sub LoadCasesFromTestCaseData()
		If TestScenarios.Count = 0 Then
		     Exit Sub
		End If

		'Region Constants
			'''<value type="Integer"/>
			Const SCENARIO_COLUMN_ID = 1
			'''<value type="Integer"/>
			Const CASE_COLUMN_ID = 2
			'''<value type="Integer"/>
			Const CASE_COLUMN_EXECFLAG = 3
			'''<value type="Integer"/>
			Const DATA_COLUMN_START = 4
			'''<value type="Integer"/>
			Const TEST_CASE_ROW_START = 2
		'End Region
		
		'Region Variables
			''' <value type="Integer"/>
			''' <summary>Row designation for loop</summary>
			Dim i
			''' <value type="Integer"/>
			''' <summary>Column designation for loop</summary>
			Dim j
			''' <value type="Integer"/>
			''' <summary>Designates the scenario's header row in the TestCases datatable</summary>
			Dim SCENARIO_ROW_HEADER
			''' <value type="Integer"/>
			''' <summary>Designates the Test Case's row in the TestCases DataTable</summary>
			Dim CASE_ROW_ID
			''' <value type="CTestScenario"/>
			Dim objScenario	
		'End Region
			
		'For Each Scenario to be executed
		'	Loop through the DataTable and find the Header row for the Scenario
		'	Loop through each test case for the scenario
		'		If the test case is set to be Executed
		'			Create new TestCase object and add it to the TestCases for the Scenario
		For Each objScenario In TestScenarios.Items()
			i = TEST_CASE_ROW_START
			DataTable.GetSheet(TestCasesSheetName).SetCurrentRow(i)
			SCENARIO_ROW_HEADER = -1
			
			While SCENARIO_ROW_HEADER < 0 And i < DataTable.GetSheet(TestCasesSheetName).GetRowCount()
				If (DataTable.Value(SCENARIO_COLUMN_ID,TestCasesSheetName) = objScenario.TestScenarioID) Then
				    SCENARIO_ROW_HEADER = i
					objScenario.TestCaseDataHeaderRow = i	
				
				Else
				    i = i + 1
					DataTable.GetSheet(TestCasesSheetName).SetCurrentRow(i)
				End If
				
			Wend
			
			CASE_ROW_ID = objScenario.TestCaseDataHeaderRow + 1
			DataTable.GetSheet(TestCasesSheetName).SetCurrentRow(CASE_ROW_ID)
			
			
			For envnt  = 1 To 6 Step 1
				If Ucase(DataTable.Value(SCENARIO_COLUMN_ID,TestCasesSheetName)) = Ucase(Environment("Testing_Environment")) Then
		     		While (Trim(DataTable.Value(CASE_COLUMN_ID, TestCasesSheetName)) <> "")
						If  (UCase(Trim(DataTable.Value(CASE_COLUMN_EXECFLAG,TestCasesSheetName))) = "YES") Then
							''' <value type="CTestCase"/>
							Dim objTestCase
							Set objTestCase = New CTestCase
					
							objTestCase.ParentScenarioID = objScenario.TestScenarioID
							objTestCase.ParentScenarioHeaderRow = objScenario.TestCaseDataHeaderRow
							objTestCase.TestCaseDescription = objScenario.TestScenarioDescription
							objTestCase.TestCaseDataRow = CASE_ROW_ID
							objTestCase.TestCaseID = Trim(DataTable.Value(CASE_COLUMN_ID,TestCasesSheetName))
					
							Call objScenario.TestCases.Add(objTestCase.TestCaseID,objTestCase)
						End If

						Exit For 
						'CASE_ROW_ID = CASE_ROW_ID + 1
						'DataTable.GetSheet(TestCasesSheetName).SetCurrentRow(CASE_ROW_ID)								
					Wend
				Else
				    CASE_ROW_ID = CASE_ROW_ID + 1
				    DataTable.GetSheet(TestCasesSheetName).SetCurrentRow(CASE_ROW_ID)
					If Ucase(DataTable.Value(SCENARIO_COLUMN_ID,TestCasesSheetName)) = "" Then
						Exit For
					End If				   
				End If
			Next
						
		Next		
	End Sub
	
	''' <summary>
	''' Loads the Data for a Test Case from the DataTable
	''' </summary>
	''' <param name="objTestCase" type="CTestCase"></param>
	Private Sub LoadTestCaseData(ByRef objTestCase)
	
		'Region Constants
			Const COLUMN_START = 4
		'End Region
		
		'Region Variables
			Dim j
		'End Region
		
		j = 0
		objTestCase.TestCaseData.Data.RemoveAll()
		
		DataTable.GetSheet( TestCasesSheetName ).SetCurrentRow( objTestCase.ParentScenarioHeaderRow )
		
		
		While Trim(DataTable.Value(COLUMN_START + j, TestCasesSheetName)) <> ""
			If Trim(DataTable.Value(COLUMN_START + j, TestCasesSheetName)) <> "-" Then
				''' <value type="String"/>
				Dim strKey, strValue
				
				strKey = Trim(DataTable.Value(COLUMN_START + j, TestCasesSheetName))
				DataTable.GetSheet(TestCasesSheetName).SetCurrentRow(objTestCase.TestCaseDataRow)
				strValue = Trim(DataTable.Value(COLUMN_START + j, TestCasesSheetName))
				Call objTestCase.TestCaseData.Data.Add(strKey, strValue)
			     
			End If

			j = j + 1
			DataTable.GetSheet(TestCasesSheetName).SetCurrentRow(objTestCase.ParentScenarioHeaderRow)			

		Wend
		
		'Find Scenario header in datasheet
			'Find TestCase from Scenario Header
			'Add data to Data Dictionary object
		
		
		'Return data dictionary object
		
	End Sub
	
	''' <summary>
	''' Executes Test Cases.
	''' 1 - Loads DataTable
	''' 2 - Loads Scenarios and Test Cases
	''' 3 - Calls Test Handler
	''' </summary>
	Public Sub ExecuteTestCases()
		Environment("LogFileLocation")=""
		'TODO: Find a way lessen the load for SerializeResultsToXML()
		Call LoadDataTable()
		Call LoadScenariosFromTestScenariosData()
		Call LoadCasesFromTestCaseData()
		
		''' <value type="CTestScenario"/>
		Dim objScenario
		''' <value type="CTestCase"/>
		Dim objCase
		
		'shweta - 23/2/2016 - Start
		Dim TestCaseCounter
		TestCaseCounter = 0
		'shweta - 23/2/2016 - End
		
		For Each objScenario In TestScenarios.Items()
			'shweta - 23/2/2016 - Start
			TestCaseCounter = TestCaseCounter+1
			If (TestCaseCounter mod 5) = 0  Then
				wait 300
				Systemutil.CloseProcessByName "iexplore.exe"
				objUtils.ClearBrowserHistory()
			End If
			'shweta - 23/2/2016 - End
			
		    Set CurrentScenario = objScenario
			CurrentScenario.StartTime = Now
			
			For Each objCase In objScenario.TestCases.Items()
				LoadTestCaseData(objCase)
				Set CurrentTestCase = objCase
				CurrentTestCase.StartTime = Now
				Call ExecuteTestHandler()
				
				CurrentTestCase.EndTime = Now
				Call SerializeResultsToXML()
			Next
			
			CurrentScenario.EndTime = Now
			Call SerializeResultsToXML()
			
			
		Call IMSAppscript_HTML_MasterReport.CreateMasterReport()
		
		
		Next
		
		Environment.Value("EndTime") = Now
		Call SerializeResultsToXML()
	'Rajesh - 04/10/2016 - Start
		'###########	TODO - Disable ShowHTMLReport based on bulk execution status  ########
	'IMSAppscript_HTML_MasterReport.ShowHTMLReport
'Rajesh - 04/10/2016 - End
  Set Fso=CreateObject("Scripting.FileSystemobject")
 'Fso.CopyFile Environment("LogFileLocation"),Environment.Value("TestDocumentsFolderPath")&"\"&Environment("TestName")&".xls",True
  Set Fso=Nothing






End Sub

	''' <summary>
	''' Calls determines the functions to call for the currently running test case
	''' </summary>
	Private Sub ExecuteTestHandler()
		'TODO: Add New Application Calls Here:
		
		''' <value type="Array"/>
		Dim arrScenario : arrScenario = Split(CurrentScenario.TestScenarioID,"_")
		Dim strTestScenarioApp
		if arrScenario(0) = "OCRF" Then 
			strTestScenarioApp = arrScenario(0)
		ElseIf arrScenario(0) = "SCA" Then 
			strTestScenarioApp = arrScenario(3)
		End If 
		
		Select Case strTestScenarioApp	
			Case "Ph1":
			    Environment.Value("ReportCreationFile")="InputFiles\IMSSCAWeb\Phase1\SCAReportSheet2356.xls"
				Call  IMSSCA.Scenarios.Handler()
			Case "Ph2":
				Environment.Value("ReportCreationFile")="InputFiles\IMSSCAWeb\Phase2\SCAReportSheet2358.xls"
				Call  IMSSCA.Scenarios.Handler()
				'<Date 11/24/2016><Phase3 scripts rename><By poornima>
			Case "Ph3":
				Environment.Value("ReportCreationFile")="InputFiles\IMSSCAWeb\Phase3\SCAReportSheetOpsInt.xls"	
				Environment.Value("SCAConfigurationOutputFile")="InputFiles\IMSSCAWeb\Phase3\SCAConfigurationOutput.xls"				
				Call  IMSSCA.Scenarios.Handler()
			Case "Ph4":
				Environment.Value("DCURL") = "http://sit.imsbi.rxcorp.com/clients/scascriptphase04"
				Environment.Value("ReportCreationFile")="InputFiles\IMSSCAWeb\Phase4\SCAReportSheet2359.xls"
				Call  IMSSCA.Scenarios.Handler()
			Case "IAMSubs":
				Environment.Value("ReportCreationFile")="InputFiles\IMSSCAWeb\SCAReportCreation.xls"
				Environment.Value("DCURL") = Environment("URLDCSite") & Environment("IAMClientName") 
				Call  IMSSCA.Scenarios.Handler()
			Case "OCRF":
				'Environment.Value("ReportCreationFile")="InputFiles\IMSSCAWeb\SCAReportCreation.xls"
				'Environment.Value("DCURL") = Environment("URLDCSite") & Environment("IAMClientName") 
				Dim thisScenario: thisScenario=IMS_Framework_TestCaseHandler.CurrentScenario.TestScenarioID
				Dim thisDictionaryObj: Set thisDictionaryObj=IMS_Framework_TestCaseHandler.CurrentTestCase.TestCaseData.Data

				Dim run : run = "Call " & thisScenario & "(thisDictionaryObj)"
				'print  thisScenario
				'Exitrun 
				Execute run
				'Call  IMSSCA.Scenarios.Handler()				
			Case Else:
				Environment.Value("DCURL") = Environment("URLDCSite") & Environment("ClientName") 
				'"http://sit.imsbi.rxcorp.com/clients/scascriptphase04"   ' For SIT
'				Environment.Value("DCURL") = "https://uat.imsbi.com/clients/uatauto" ' For UAT			    
			    
			    Environment.Value("ReportCreationFile")="InputFiles\IMSSCAWeb\SCAReportCreation.xls"	
				Environment.Value("SCAConfigurationOutputFile")="InputFiles\IMSSCAWeb\ReportValidation\SCAConfigurationOutput.xls"				
				Call  IMSSCA.Scenarios.Handler()
				'ReportStep StatusTypes.Warning, "Test Scenario Not Found: " & CurrentScenario.TestScenarioID, "Test Case Handler"
		End Select
		
	End Sub

	''' <summary>
	''' Loads the DataTable from the TestCases spreadsheet
	''' </summary>
	Private Sub LoadDataTable()
		
		Datatable.AddSheet(TestScenariosSheetName) 																
		Datatable.AddSheet(TestCasesSheetName)
		
		Datatable.ImportSheet Environment.Value("TestScenariosFile"),TestScenariosSheetName,TestScenariosSheetName
		Datatable.ImportSheet Environment.Value("TestCasesFile"),TestCasesSheetName,TestCasesSheetName

	End Sub
	
	Public Sub SerializeResultsToXML()
		Dim oXML
		Dim oXMLElementResults
		Dim oXMLElementApplication
		Dim oXMLElementEnvironment
		Dim oXMLElementExecutionStart
		Dim oXMLElementExecutionEnd
		Dim oXMLElementTotalExecutionTime
		Dim oXMLElementReleaseNo
		Dim oXMLElementStatus
		Dim oXMLElementTestScenarios

		'Create XML Object
		Set oXML = CreateObject("Msxml2.DOMDocument.6.0")

		'Create Results Element
		Set oXMLElementResults = oXML.createElement( "Results" )

		'Create and add Application Element to Results Element
		Set oXMLElementApplication = oXML.createElement( "Application" )
		oXMLElementApplication.Text = Environment.Value("Application")
		oXMLElementResults.appendChild oXMLElementApplication

		'Create and add Environment Element to Results Element
		Set oXMLElementEnvironment = oXML.createElement( "Environment" )
		oXMLElementEnvironment.Text = Environment.Value("TestEnv")
		oXMLElementResults.appendChild oXMLElementEnvironment

		'Create and add Execution Start Element to Results Element
		Set oXMLElementExecutionStart = oXML.createElement( "ExecutionStart" )
		oXMLElementExecutionStart.Text = Environment.Value("StartTime")
		oXMLElementResults.appendChild oXMLElementExecutionStart
		Environment.Value("EndTime") = Now
		'Create and add Execution End Element to Results Element
		Set oXMLElementExecutionEnd = oXML.createElement( "ExecutionEnd" )
		oXMLElementExecutionEnd.Text = Environment.Value("EndTime")
		oXMLElementResults.appendChild oXMLElementExecutionEnd		
		
		'<Shweta - 15/11/2016 > Create and add Total Execution Element to Results Element - Start
		Set oXMLElementTotalExecutionTime = oXML.createElement( "TotalExecutionTime" )
		iNumSec = DateDiff("s", Environment.Value("StartTime"), Environment.Value("EndTime"))
	        ' calculates whole hours 
	        NumHour = iNumSec\3600 
	     	' calculate remaining number of seconds 
	        NumSec = iNumSec Mod 3600 
	     	' calculate whole number of minutes 
		    NumMin = NumSec\60 
	        ' calculate remaining number of seconds 
	      	NumSec = NumSec Mod 60 
		
		Environment.Value("TotalExecutionTime") = NumHour&":"&NumMin&":"&NumSec
		oXMLElementTotalExecutionTime.Text =  Environment.Value("TotalExecutionTime")
		oXMLElementResults.appendChild oXMLElementTotalExecutionTime		
		
		'<Shweta - 15/11/2016 > Create and add Execution Element to Results Element - End
		
		'Create and add ReleaseNo Element to Results Element
		Set oXMLElementReleaseNo = oXML.createElement( "ReleaseNo" )
		oXMLElementReleaseNo.Text = Environment.Value("ReleaseNo")
		oXMLElementResults.appendChild oXMLElementReleaseNo
		
		'Create and add Status Element to Results Element
		Set oXMLElementStatus = oXML.createElement( "Status" )
		oXMLElementStatus.Text = StatusTypes.GetStringFromStatusType( Status )
		oXMLElementResults.appendChild oXMLElementStatus		
		
		
		'Create TestScenarios Element
		Set oXMLElementTestScenarios = oXML.createElement( "TestScenarios" )
		
		'For each Test Scenario - Create and add TestScenario Elements to TestScenarios Element - Calls TestScenarioXML()
		''' <value type="CTestScenario"/>
		Dim objScenario
		For Each objScenario In TestScenarios.Items()
			oXMLElementTestScenarios.appendChild( objScenario.SerializeToXML() )
		Next

		'Add Scenarios Element to TestCase Element
		oXMLElementResults.appendChild( oXMLElementTestScenarios )
		oXML.appendChild( oXMLElementResults )
		
		'Save XML File to Results Directory
		If ( Not objUtils.fnFolderExist( Environment.Value("TestResultsFolderPath") ) ) Then
			Call objUtils.fnCreateFolder(Environment.Value("ResultsPath"), Environment.Value("TestResultsFolderName"))
		End If	
		
		oXML.save(Environment.Value("TestResultsFolderPath") & "Data.xml")
		
	End Sub
	
End Class	
	
	
Class CTestScenario

	''' <value type="String"/>
	Private vTestScenarioID
	''' <value type="String"/>
	Private vTestScenarioDescription	
	''' <value type="Integer"/>
	Private vTestCaseDataRow
	''' <value type="Scripting.Dictionary"/>
	Private vTestCasesDictionary
	''' <value type="DateTime"/>
	Private vEndTime
	''' <value type="DateTime"/>
	Private vStartTime	
	
	''' <summary>
	''' Get/Set the TestScenarioID
	''' </summary>
	''' <value type="String"/>
	Public Property Get TestScenarioID
		TestScenarioID = vTestScenarioID
	End Property
	Public Property Let TestScenarioID(ByVal vData)
		vTestScenarioID = vData
	End Property
	
	''' <summary>
	''' Get/Set the TestScenarioDescription
	''' </summary>
	''' <value type="String"/>
	Public Property Get TestScenarioDescription
		TestScenarioDescription = vTestScenarioDescription
	End Property
	Public Property Let TestScenarioDescription(ByVal vData)
		vTestScenarioDescription = vData
	End Property
	
	''' <summary>
	''' Get/Set TestCaseDataHeaderRow
	''' </summary>
	''' <value type="Integer"/>
	Public Property Get TestCaseDataHeaderRow
		TestCaseDataHeaderRow = vTestCaseDataRow
	End Property
	Public Property Let TestCaseDataHeaderRow(ByVal vData)
		vTestCaseDataRow = vData
	End Property
	
	''' <summary>
	''' Get the Dictionary containing the test cases
	''' </summary>
	''' <value type="Scripting.Dictionary"/>
	Public Property Get TestCases
		Set TestCases = vTestCasesDictionary
	End Property
		
	''' <summary>
	''' Get/Set the Time the Test Ends
	''' </summary>
	''' <value type="DateTime"/>
	Public Property Get EndTime
		EndTime = vEndTime
	End Property
	Public Property Let EndTime(ByVal vData)
		vEndTime = vData
	End Property
		
	''' <summary>
	''' Get/Set the Time the Test Starts
	''' </summary>
	''' <value type="DateTime"/>
	Public Property Get StartTime
		StartTime = vStartTime
	End Property
	Public Property Let StartTime(ByVal vData)
		vStartTime = vData
	End Property
	
	''' <summary>
	''' Gets the Status of the Test Scenario {CStatusEnum}
	''' </summary>
	''' <value type="Integer"/>
	Public Property Get Status
		''' <value type="Integer"/>
		Dim vFailCount, vWarnCount
		vFailCount = 0
		vWarnCount = 0
		
		''' <value type="CTestCase"/>
		Dim objCaseStatus
	
		For Each objCaseStatus In TestCases.Items()
		     Select Case objCaseStatus.Status
				Case StatusTypes.Fail:
					vFailCount = vFailCount + 1
				Case StatusTypes.Warning:
					vWarnCount = vWarnCount + 1
				Case Else:
				
			End Select
		Next
		
		If vFailCount <> 0 Then
			Status = StatusTypes.Fail
		Else
			If vWarnCount <> 0 Then
				Status = StatusTypes.Warning
			Else
				Status = StatusTypes.Pass
			End If
		End If		
	End Property
	
	Private Sub Class_Initialize()
		TestScenarioID = "Unknown"
		TestScenarioDescription = "Unknown"
		TestCaseDataHeaderRow = -1
		Set vTestCasesDictionary = CreateObject("Scripting.Dictionary")
		vTestCasesDictionary.RemoveAll()
	End Sub
	
	Public Function SerializeToXML()
		Dim oXML
		Dim oXMLElementTestScenario
		Dim oXMLAttributeTestScenarioID
		Dim oXMLElementDescription
		Dim oXMLElementStatus
		Dim oXMLElementDuration
		Dim oXMLElementTestCases

		Set oXML = CreateObject("Msxml2.DOMDocument.6.0")

		'Create TestScenario Element
		Set oXMLElementTestScenario = oXML.createElement( "TestScenario" )

		'Create and add ID Attribute to TestScenario Element
		Set oXMLAttributeTestScenarioID = oXML.createAttribute( "ID" )
		oXMLAttributeTestScenarioID.nodeValue = TestScenarioID
		oXMLElementTestScenario.setAttributeNode( oXMLAttributeTestScenarioID )

		'Create and add Description Element to TestScenario Element
		Set oXMLElementDescription = oXML.createElement( "Description" )
		oXMLElementDescription.Text = TestScenarioDescription
		oXMLElementTestScenario.appendChild oXMLElementDescription

		'Create and add Status Element to TestScenario Element
		Set oXMLElementStatus = oXML.createElement( "Status" )
		oXMLElementStatus.Text = StatusTypes.GetStringFromStatusType(Status)
		oXMLElementTestScenario.appendChild oXMLElementStatus

		'Create and add Duration Element to TestScenario Element
		Set oXMLElementDuration = oXML.createElement( "Duration" )
		oXMLElementDuration.Text = StartTime
		oXMLElementTestScenario.appendChild oXMLElementDuration
		
		'Create TestCases Element
		Set oXMLElementTestCases = oXML.createElement( "TestCases" )
		
		'For each TestCase - Create and add TestCase Element to Steps Element - Calls SerializeToXML() for each Test Case
		''' <value type="CTestCase"/>
		Dim objTestCase
		For Each objTestCase In TestCases.Items()
			oXMLElementTestCases.appendChild( objTestCase.SerializeToXML() )     
		Next
		
		'Add TestCases Element to TestScenario Element
		oXMLElementTestScenario.appendChild( oXMLElementTestCases )
	
		Set SerializeToXML = oXMLElementTestScenario
	End Function
	
End Class

	
Class CTestCase

	''' <value type="String"/>
	Private vParentScenarioID
	''' <value type="Integer"/>
	Private vParentScenarioHeaderRow	
	''' <value type="String"/>
	Private vTestCaseID	
	''' <value type="Integer"/>
	Private vTestCaseDataRow	
	''' <value type="CTestCaseData"/>
	Private vTestCaseData	
	''' <value type="Scripting.Dictionary"/>
	Private vTestCaseSteps
	''' <value type="Integer"/>
	Private vTestCaseStepCounter	
	''' <value type="DateTime"/>
	Private vStartTime	
	''' <value type="DateTime"/>
	Private vEndTime
	''' <value type="String"/>
	Private vTestCaseDescription
	
	''' <summary>
	''' Gets/Sets the Test Case's Description
	''' </summary>
	''' <value type="String"/>
	Public Property Get TestCaseDescription
		TestCaseDescription = vTestCaseDescription
	End Property
	Public Property Let TestCaseDescription(ByVal vData)
		vTestCaseDescription = vData
	End Property
	
	
	''' <summary>
	''' Get/Set the Parent Scenario. E.G "AD_S001"
	''' </summary>
	''' <value type="String"/>
	Public Property Get ParentScenarioID
		ParentScenarioID= vParentScenarioID
	End Property
	Public Property Let ParentScenarioID(ByVal vData)
		vParentScenarioID = vData
	End Property
	
	''' <summary>
	''' The parent scenario's header row.
	''' </summary>
	''' <value type="Integer"/>
	Public Property Get ParentScenarioHeaderRow
		ParentScenarioHeaderRow = vParentScenarioHeaderRow
	End Property
	Public Property Let ParentScenarioHeaderRow(ByVal vData)
		vParentScenarioHeaderRow = vData
	End Property
	
	''' <summary>
	''' Gets/Sets the TestCaseName
	''' </summary>
	''' <value type="String"/>
	Public Property Get TestCaseID
		TestCaseID = vTestCaseID
	End Property
	Public Property Let TestCaseID(ByVal vData)
		vTestCaseID = vData
	End Property
	
	''' <summary>
	''' Gets/Sets the TestCaseDataRow
	''' </summary>
	''' <value type="Integer"/>
	Public Property Get TestCaseDataRow
		TestCaseDataRow = vTestCaseDataRow
	End Property
	Public Property Let TestCaseDataRow(ByVal vData)
		vTestCaseDataRow = vData
	End Property
	
	''' <summary>
	''' Gets the Dictionary object containing the Test Case's Data
	''' </summary>
	''' <value type="CTestCaseData"/>
	Public Property Get TestCaseData
		Set TestCaseData = vTestCaseData
	End Property
	
	''' <summary>
	''' Gets the Dictionary Object containing all the Test Case's Steps
	''' </summary>
	''' <value type="Scripting.Dictionary"/>
	Public Property Get TestCaseSteps
		Set TestCaseSteps = vTestCaseSteps
	End Property
	
	''' <summary>
	''' Gets the Status of a Test Case. Returns a {CStepEnum} value 
	''' </summary>
	''' <value type="Integer"/>
	Public Property Get Status
		Dim vFailCount, vWarnCount
		vFailCount = 0
		vWarnCount = 0
		
		''' <value type="CTestCaseStep"/>
		Dim objStep
		
		For Each objStep In TestCaseSteps.Items()
		    Select Case objStep.StepStatus
				Case StatusTypes.Fail:
					vFailCount = vFailCount + 1
				Case StatusTypes.Warning:
					vWarnCount = vWarnCount + 1
				Case Else:
				
			End Select
		Next
		
		
		If vFailCount <> 0 Then
			Status = StatusTypes.Fail
		Else
			If vWarnCount <> 0 Then
				Status = StatusTypes.Warning
			Else
				Status = StatusTypes.Pass
			End If
		End If
	End Property
	
	''' <summary>
	''' Get/Set the Start Time
	''' </summary>
	''' <value type="DateTime"/>
	Public Property Get StartTime
		StartTime = vStartTime
	End Property
	Public Property Let StartTime(ByVal vData)
		vStartTime = vData
	End Property
	
	''' <summary>
	''' Get/Set the End Time
	''' </summary>
	''' <value type="DateTime"/>
	Public Property Get EndTime
		EndTime = vEndTime
	End Property
	Public Property Let EndTime(ByVal vData)
		vEndTime = vData
	End Property	
	
	''' <summary>
	''' Adds a Test Case Step to the Test Case
	''' </summary>
	''' <param name="vStepStatus" type="Integer">StatusTypes Object {CStepEnum}</param>
	''' <param name="vActualResult" type="String">Actual Result of the Step</param>
	''' <param name="vLocation" type="String">Location the step occurs. E.G. Function | Screen Field | etc.</param>
	''' <param name="vTimeStamp" type="DateTime">The Time that the step occurred</param>
	Public Sub AddStep(vStepStatus, vActualResult, vLocation, vTimeStamp)

		'<22/9/2016> Added log information - Start
		' create a Text file for every test script. 
		Set Fso=CreateObject("Scripting.FileSystemObject")
		strLogFilePath=Environment.Value("TestDocumentsFolderPath")&"\"&Environment("TestName")&".txt"
		bcreationfile="Notcreated"
		If Fso.FileExists(strLogFilePath) = False Then 
			Fso.CreateTextFile(strLogFilePath)
			detailLoginfo="Sr No" & Chr(9) & "STEP STATUS" & Chr(9) & "Location" & Chr(9)  & "Description" & Chr(9) & "DateTime"
		        Set columnfile=Fso.OpenTextFile(strLogFilePath, 8, True) 
				columnfile.Writeline detailLoginfo
				columnfile.Close
				Set columnfile=Nothing 
				bcreationfile="created"
		End if 
		 
		' log the information and save the file
		'Set fsowrite = fso.OpenTextFile(strLogFilePath,2, True)
		Set fsoAppend = fso.OpenTextFile(strLogFilePath, 8, True)
		
		'<15/11/2016 HTMLFile> - Start
		' create a LogHTML file for every test script. 
		Set FsoHTML=CreateObject("Scripting.FileSystemObject")
		strHTMLLogFilePath=Environment.Value("TestDocumentsFolderPath")&"\"&Environment("TestName")&".html"
		
		If FsoHTML.FileExists(strHTMLLogFilePath) = False Then 
		       FsoHTML.CreateTextFile(strHTMLLogFilePath)
				'test 21/11/2016 - start
						       Set fsoHTMLBody = FsoHTML.OpenTextFile(strHTMLLogFilePath, 8, True)
						       fsoHTMLBody.WriteLine("<html>")
						       fsoHTMLBody.WriteLine("<body>")
						       fsoHTMLBody.WriteLine("<pre>")
						       'fsoHTMLBody.WriteLine("<tr>test name</tr>")
						       'fsoHTMLBody.WriteLine("<tr>test action name</tr>")
						       'fsoHTMLBody.WriteLine("<tr>test iteration name</tr>")
						       'fsoHTMLBody.WriteLine("</pre>")
						       'fsoHTMLBody.WriteLine("</body>")
						       'fsoHTMLBody.WriteLine("</html>")
							   fsoHTMLBody.Close
							   set fsoHTMLBody = nothing					
				'test 21/11/2016 - end
		End if 
		 
		' log the information and save the file
		Set fsoHTMLAppend = FsoHTML.OpenTextFile(strHTMLLogFilePath, 8, True)
		'<15/11/2016 HTMLFile> - End
		'<22/9/2016> Added log information - End

		If vTestCaseStepCounter = 0 And TestCaseSteps.Count <> 0 Then
		     TestCaseSteps.RemoveAll()
		End If
		
		vTestCaseStepCounter = vTestCaseStepCounter + 1

		Dim objStep
		Set objStep = New CTestCaseStep
		objStep.ID = vTestCaseStepCounter
		objStep.StepStatus = vStepStatus
		objStep.ActualResult = vActualResult
		objStep.Location = vLocation
		objStep.TimeStamp = vTimeStamp
		
		'<22/9/2016> Added log information - Start
		'vDefault = -1;	'vPass = 1;	'vWarn = 2;	'vInfo = 3;	'vFail = 4
		If vStepStatus=1 Then
		   stepStatus="PASS"
		ElseIf vStepStatus=2 Then
		   stepStatus="WARNING"
		ElseIf vStepStatus=3 Then
		   stepStatus="INFORMATION"
		ElseIf vStepStatus=4 Then 
		   stepStatus="FAIL"
	    End if
	    detailLoginfo=vTestCaseStepCounter & ":" & stepStatus &space(3)& vLocation & ":" &  ":"  & vActualResult & ":" &space(3)& vTimeStamp
		fsoAppend.Writeline detailLoginfo
	   '<22/9/2016> Added log information - End
	   
		If vStepStatus = 4 Then
			Dim strFileName : strFileName = IMS_Framework_TestCaseHandler.CurrentTestCase.TestCaseID &"_Failure_"&objStep.ID
			Call SCA.Screenshot(strFileName)
			
		
		'<23/11/2016> changed to give relative path  to see failure screenshots <By Poornima Appoji> 
			'objStep.ActualResult=objStep.ActualResult & " <a href=' "&Environment.Value("TestDocumentsFolderPath")& IMS_Framework_TestCaseHandler.CurrentTestCase.TestCaseID&"\"&strFileName&".png'> Click here to see the failure screenshot </a>"		 
		 objStep.ActualResult=objStep.ActualResult & "<a href='../Test_Documents\"& IMS_Framework_TestCaseHandler.CurrentTestCase.TestCaseID&"\"&strFileName&".png'>Click here to see the failure screenshot </a>"
		'<22/9/2016> Added Reporter message to capture pass/fail events - Start
			Reporter.ReportEvent micFail,vLocation,objStep.ActualResult
		End If

		If vStepStatus =1 Then
		      fsoAppend.Writeline detailLoginfo
		     Reporter.ReportEvent micpass,vLocation,vActualResult
		    
		End if 
		'<22/9/2016> Added Reporter message to capture pass/fail events - Start
		
		'shweta 30/8/2016 - Added Screenshot capture for important specific functions - Start
		If vStepStatus = -1 Then
			Dim strDefFileName : strDefFileName = IMS_Framework_TestCaseHandler.CurrentTestCase.TestCaseID &"_FunctionStep_"&objStep.ID
			Call SCA.Screenshot(strDefFileName)
			'<23/11/2016> changed to give relative path  to see failure screenshots <By Poornima Appoji>
			objStep.ActualResult=objStep.ActualResult & "<a href='../Test_Documents\"& IMS_Framework_TestCaseHandler.CurrentTestCase.TestCaseID&"\"&strFileName&".png'>Click here to see the failure screenshot </a>"
		End If
		'shweta 30/8/2016 - Added Screenshot capture for important specific functions - End	
				
		'<15/11/2016 HTMLFile> - Start
		If vStepStatus = 4 Then
				htmlDetailLogInfo=vTestCaseStepCounter & ":" & stepStatus &space(3)& vLocation & ":" &  ":"  & vActualResult & ":" &space(3)& vTimeStamp
				'<23/11/2016> changed to give relative path  to see failure screenshots <By Poornima Appoji>
				'fsoHTMLAppend.WriteLine htmlDetailLogInfo&space(3)& " <a href=' "&Environment.Value("TestDocumentsFolderPath")& IMS_Framework_TestCaseHandler.CurrentTestCase.TestCaseID&"\"&strFileName&".png'> Click here to see the failure screenshot </a>"
				fsoHTMLAppend.WriteLine htmlDetailLogInfo&space(3)& "<a href='../Test_Documents\"& IMS_Framework_TestCaseHandler.CurrentTestCase.TestCaseID&"\"&strFileName&".png'> Click here to see the failure screenshot </a>"
		ElseIf  vStepStatus = -1 Then
				htmlDetailLogInfo=vTestCaseStepCounter & ":" & stepStatus &space(3)& vLocation & ":" &  ":"  & vActualResult & ":" &space(3)& vTimeStamp
				
				'<23/11/2016> changed to give relative path  to see failure screenshots <By Poornima Appoji>
				'fsoHTMLAppend.WriteLine htmlDetailLogInfo&Space(3)&" <a href=' "&Environment.Value("TestDocumentsFolderPath")& IMS_Framework_TestCaseHandler.CurrentTestCase.TestCaseID&"\"&strDefFileName&".png'> Click here to see the Function Step screenshot </a>"			
                fsoHTMLAppend.WriteLine htmlDetailLogInfo&Space(3)&"<a href='../Test_Documents\"& IMS_Framework_TestCaseHandler.CurrentTestCase.TestCaseID&"\"&strFileName&".png'> Click here to see the Function Step screenshot </a>"			
        Else
				htmlDetailLogInfo=vTestCaseStepCounter & ":" & stepStatus &space(3)& vLocation & ":" &  ":"  & vActualResult & ":" &space(3)& vTimeStamp
				fsoHTMLAppend.WriteLine htmlDetailLogInfo
		End If
		'<15/11/2016 HTMLFile> - End
		
		
		Call TestCaseSteps.Add(vTestCaseStepCounter, objStep)
						
		fsoAppend.Close
		
	    Set fsoAppend=Nothing 
	    Set Fso=Nothing 
	   
	   Environment("LogFileLocation")=strLogFilePath  ' declared environment log file location path.
	   'Msgbox Environment("LogFileLocation")
	End Sub 
	
	Private Sub Class_Initialize()
		ParentScenarioID = "Unknown"
		TestCaseID = "Unknown"
		TestCaseDataRow = -1
		Set vTestCaseData = New CTestCaseData
		vTestCaseData.Data.RemoveAll()
		Set vTestCaseSteps = CreateObject("Scripting.Dictionary")
		vTestCaseSteps.RemoveAll()
		vTestCaseStepCounter = 0
	End Sub
	
	Public Function SerializeToXML()
		Dim oXML
		Dim oXMLElementTestCase
		Dim oXMLAttributeTestCaseID
		Dim oXMLElementStatus
		Dim oXMLElementSteps
		Dim oXMLElementDescription

		'Create XML Dom Object
		Set oXML = CreateObject("Msxml2.DOMDocument.6.0")

		'Create TestCase Element
		Set oXMLElementTestCase = oXML.createElement( "TestCase" )

		'Create and add ID Attribute to TestCase Element
		Set oXMLAttributeTestCaseID = oXML.createAttribute( "ID" )
		oXMLAttributeTestCaseID.nodeValue = TestCaseID
		oXMLElementTestCase.setAttributeNode( oXMLAttributeTestCaseID )	

		'Create and add Status Element to TestCase Element
		Set oXMLElementStatus = oXML.createElement( "Status" )
		oXMLElementStatus.Text = StatusTypes.GetStringFromStatusType( Status )
		oXMLElementTestCase.appendChild oXMLElementStatus
		
	
		'Create and add Description Element to TestCase Element
		Set oXMLElementDescription = oXML.createElement( "Description" )
		oXMLElementDescription.Text = TestCaseDescription
		oXMLElementTestCase.appendChild oXMLElementDescription
		
		'Create Steps Element
		Set oXMLElementSteps = oXML.createElement( "Steps" )

		'For each Step - Create and add Step Element to Steps Element - Calls SerializeToXML() for that Step
		''' <value type="CTestCaseStep"/>
		Dim objStep
		For Each objStep In TestCaseSteps.Items()
			oXMLElementSteps.appendChild( objStep.SerializeToXML() )     
		Next
		
		'Add Steps Element to TestCase Element
		oXMLElementTestCase.appendChild( oXMLElementSteps )

		Set SerializeToXML = oXMLElementTestCase
	End Function
End Class


Class CTestCaseStep

	''' <value type="String"/>
	Private vActualResult
	''' <value type="String"/>
	Private vLocation
	''' <value type="Integer"/>
	Private vStepStatus
	''' <value type="DateTime"/>
	Private vTimeStamp
	''' <value type="Integer"/>
	Private vID

	''' <summary>
	''' Get/Set the Actual Result of the step
	''' </summary>
	''' <value type="String"/>
	Public Property Get ActualResult
		ActualResult = vActualResult
	End Property
	Public Property Let ActualResult(ByVal vData)
		vActualResult = vData
	End Property

	''' <summary>
	''' Get/Set the Location of the step. E.G The function name or application field, etc.
	''' </summary>
	''' <value type="String"/>
	Public Property Get Location
		Location = vLocation
	End Property
	Public Property Let Location(ByVal vData)
		vLocation = vData
	End Property

	''' <summary>
	''' Get/Set the TimeStamp that the Step occurred
	''' </summary>
	''' <value type="DateTime"/>
	Public Property Get TimeStamp
		TimeStamp = vTimeStamp
	End Property
	Public Property Let TimeStamp(ByVal vData)
		vTimeStamp = vData
	End Property

	''' <summary>
	''' Get/Set the StatusType for the step {CStepEnum}
	''' </summary>
	''' <value type="Integer"/>
	Public Property Get StepStatus
		StepStatus = vStepStatus
	End Property
	Public Property Let StepStatus(ByVal vData)
		vStepStatus = vData
	End Property
	
	''' <summary>
	''' Gets/Sets the ID for this Step.
	''' </summary>
	''' <value type=""/>
	Public Property Get ID
		ID = vID
	End Property
	Public Property Let ID(ByVal vData)
		vID = vData
	End Property
	

	Public Function SerializeToXML()
		Dim oXML
		Dim oXMLElementStep
		Dim oXMLAttributeStepID
		Dim oXMLElementStatus
		Dim oXMLElementLocation
		Dim oXMLElementDescription
		Dim oXMLElementExecutionStart

		'Create XML Dom Object
		Set oXML = CreateObject("Msxml2.DOMDocument.6.0")

		'Create Step Element
		Set oXMLElementStep = oXML.createElement( "Step" )

		'Create and add ID Attribute to Step Element
		Set oXMLAttributeStepID = oXML.createAttribute( "ID" )
		oXMLAttributeStepID.nodeValue = ID
		oXMLElementStep.setAttributeNode( oXMLAttributeStepID )

		'Create and add Status element to Step Element
		Set oXMLElementStatus = oXML.createElement( "Status" )
		oXMLElementStatus.Text = StatusTypes.GetStringFromStatusType(StepStatus)
		oXMLElementStep.appendChild oXMLElementStatus

		'Create and add Location element to Step Element
		Set oXMLElementLocation = oXML.createElement( "Location" )
		oXMLElementLocation.Text = Location
		oXMLElementStep.appendChild oXMLElementLocation

		'Create and add Description element to Step Element
		Set oXMLElementDescription = oXML.createElement( "Description" )
		oXMLElementDescription.Text = ActualResult
		oXMLElementStep.appendChild oXMLElementDescription

		'Create and add ExecutionStart element to Step Element
		Set oXMLElementExecutionStart = oXML.createElement( "Time" )
		oXMLElementExecutionStart.Text = TimeStamp
		oXMLElementStep.appendChild oXMLElementExecutionStart
	
		Set SerializeToXML = oXMLElementStep
	End Function

	Private Sub Class_Initialize
		ActualResult = ""
		Location = ""
		TimeStamp = Now
		StepStatus = StatusTypes.StepDefault
	End Sub

End Class

''' <summary>
''' Contains Status Types for a TestCase Step. I.E. Pass|Fail|etc.
''' </summary>
Class CStepEnum

	''' <value type="Integer"/>
	Private vPass
	''' <value type="Integer"/>
	Private vWarn
	''' <value type="Integer"/>
	Private vInfo
	''' <value type="Integer"/>
	Private vFail
	''' <value type="Integer"/>
	Private vDefault


	
	''' <summary>
	''' Pass Type
	''' </summary>
	''' <value type="Integer"/>
	Public Property Get Pass
		Pass = vPass
	End Property
	
	''' <summary>
	''' Warning Type
	''' </summary>
	''' <value type="Integer"/>
	Public Property Get Warning
		Warning = vWarn
	End Property
	
	''' <summary>
	''' Information Type
	''' </summary>
	''' <value type="Integer"/>	
	Public Property Get Information
		Information = vInfo
	End Property
	
	''' <summary>
	''' Fail Type
	''' </summary>
	''' <value type="Integer"/>	
	Public Property Get Fail
		Fail = vFail
	End Property

	''' <summary>
	''' Default Type
	''' </summary>
	''' <value type="Integer"/>
	Public Property Get StepDefault
		StepDefault = vDefault
	End Property
	
	
	Public Function GetStringFromStatusType(vStepStatus)
		Select Case vStepStatus
			Case StatusTypes.Pass:
				GetStringFromStatusType = "Pass"
			Case StatusTypes.Warning:
				GetStringFromStatusType = "Warning"
			Case StatusTypes.Fail:
				GetStringFromStatusType = "Fail"
			Case StatusTypes.Information:
				GetStringFromStatusType = "Information"
			Case StatusTypes.StepDefault:
				GetStringFromStatusType = "Default"
			Case Else:
				GetStringFromStatusType = "Default"
		End Select
	End Function
	
	
	Private Sub Class_Initialize
		vDefault = -1
		vPass = 1
		vWarn = 2
		vInfo = 3
		vFail = 4
	End Sub
	
	
End Class


Class CTestCaseData
	
	''' <value type="Scripting.Dictionary"/>
	Private vDictionary
	
	''' <summary>
	''' summaryInfo
	''' </summary>
	''' <value type="Scripting.Dictionary"/>
	Public Property Get Data
		Set Data = vDictionary
	End Property
	Public Property Set Data(ByRef vData)
		Set vDictionary = vData
	End Property
	
	Private Sub Class_Initialize()
		Set vDictionary = CreateObject("Scripting.Dictionary")
		vDictionary.RemoveAll()
	End Sub
	
End Class
