'-------------------------------------------------------------------------------
'
'	Script:			OSUtilities.vbs
'	Author:			Multiple
'	Date Created:	June 25, 2012
'
'	Description:
'	Utilities used with QTP
'
'-------------------------------------------------------------------------------
Option Explicit

''' <summary>
''' Public Global Instance of CUtils used for Utility Functions. E.G. creating a folder.
''' </summary>
''' <value type="CUtils"/>
Public objUtils : Set objUtils = New CUtils

Public objINIFile : Set objINIFile = New ReadINIFile

''' <summary>
''' The CUtils class deals with the general utilities used with QTP and the Framework.
''' </summary>
''' <author>Kishore Sura; Indium</author> 
''' <datecreated>6/29/2011</datecreated>
Class CUtils
	'TODO: Create a Function to Copy Files

	''' <summary>
	''' Function to Launch Internet Explorer and navigate to the specified URL.
	''' </summary>
	''' <param name="Url" type="String">The URL to navigate to after launching IE.</param>
	Public Sub fnLaunchingIE(Url)
		'<shweta 16/6/2015 Closing All opened Excel Files- start
		'objUtils.KillProcess("excel.exe")
		'<shweta 16/6/2015 Closing All opened Excel Files- end
		'objUtils.KillProcess("iexplore.exe")
		
'		Systemutil.CloseProcessByName "excel.exe"
'		
'		wait 3
'		Systemutil.CloseProcessByName "iexplore.exe"
		
	  	On Error Resume Next
	   	Dim IE,objHwnd
		Set IE = CreateObject("InternetExplorer.Application")
		IE.Visible = True	
        'IE.Navigate Url	
		IE.Navigate2 Url
		
		'Systemutil.Run Url
		
		
		Window("hwnd:="&IE.HWND).Maximize
		'objUtils.fnDeleteIECookies()
		Set IE=Nothing
	End Sub
	
	Public Sub fnDeleteIECookies()
		'Object
		Dim oWebUtil
	    Set oWebUtil = CreateObject("Mercury.GUI_WebUtil")
	    oWebUtil.DeleteCookies()
        Set oWebUtil =Nothing 	    
	End Sub
	
	Public Function fnBrowserCreationTime()
		Dim brwCnt, blnBrowser
		brwCnt=0
		Do 
			blnBrowser=Browser("micclass:=Browser","Creationtime:="&brwCnt).Exist(2)
			If blnBrowser Then brwCnt=brwCnt+1
		Loop Until blnBrowser=False
		fnBrowserCreationTime=brwCnt-1
	End Function

	Public Function fnUniqueDateTime()
		Dim strNow, strSecond, strMinute, strHour, strDay, strMonth, strYear
		
		strNow		= Now
		strSecond	= Second(strNow)
		strMinute	= Minute(strNow)
		strHour		= Hour(strNow)
		strDay		= Day(strNow)
		strMonth 	= Month(strNow)
		strYear 	= Year(strNow)

		If CInt(strSecond) < 10 Then
			strSecond = "0" & strSecond
		End If

		If CInt(strMinute) < 10 Then
			strMinute = "0" & strMinute
		End If

		If CInt(strHour) < 10 Then
			strHour = "0" & strHour
		End If

		If CInt(strDay) < 10 Then
			strDay = "0" & strDay
		End If

		If CInt(strMonth) < 10 Then
			strMonth = "0" & strMonth
		End If

		fnUniqueDateTime="_" & strYear & strMonth & strDay & "_" & strHour & strMinute & strSecond
	End Function

	Public Function fnUniqueDateTimeByVal(ByVal dte)
		Dim strSecond, strMinute, strHour, strDay, strMonth, strYear
		
		strSecond	= Second(dte)
		strMinute	= Minute(dte)
		strHour		= Hour(dte)
		strDay		= Day(dte)
		strMonth 	= Month(dte)
		strYear 	= Year(dte)

		If CInt(strSecond) < 10 Then
			strSecond = "0" & strSecond
		End If

		If CInt(strMinute) < 10 Then
			strMinute = "0" & strMinute
		End If

		If CInt(strHour) < 10 Then
			strHour = "0" & strHour
		End If

		If CInt(strDay) < 10 Then
			strDay = "0" & strDay
		End If

		If CInt(strMonth) < 10 Then
			strMonth = "0" & strMonth
		End If

		fnUniqueDateTimeByVal="_" & strYear & strMonth & strDay & "_" & strHour & strMinute & strSecond
	End Function


	Public Function fnUniqueTime()
		Dim strNow, strSecond, strMinute, strHour
		strNow		= Now
		strSecond	= Second(strNow)
		strMinute	= Minute(strNow)
		strHour		= Hour(strNow)

		If CInt(strSecond) < 10 Then
			strSecond = "0" & strSecond
		End If

		If CInt(strMinute) < 10 Then
			strMinute = "0" & strMinute
		End If

		If CInt(strHour) < 10 Then
			strHour = "0" & strHour
		End If

		fnUniqueTime="_" & strHour & strMinute & strSecond

	End Function
		
	Public Function fnFileExist(strFileName)
		Dim FileSysObj
		
		fnFileExist=True
		
		Set FileSysObj = CreateObject("Scripting.FileSystemObject")
		
		If Not FileSysObj.FileExists(strFileName) Then 
			fnFileExist=False
		End If
	End Function
	
	Public Function fnFolderExist(strFolderName)
		Dim FolderObj
		fnFolderExist=True
		Set FolderObj = CreateObject("Scripting.FileSystemObject")
		If Not FolderObj.FolderExists(strFolderName) Then fnFolderExist=False
	End Function
	
	Public Function fnScreenShot(HWND)
		'TODO: Update for AddStep Function / New Reporter. Tip: Use Images\ Folder
		Dim strBrowserName, intcounter, ResultFolder, ImageFolder, GPS_ErrorImg
		
		Dim strFilename,LocWinID
		LocWinID = Trim(HWND)
		ImageFolder=Environment.Value("ImageFolders")
		strBrowserName = Window("HWND:="&LocWinID).GetROProperty("name")
		If (fnFolderExist(ImageFolder)) Then
			Randomize 
			
			strFilename = Environment.Value("ImageFolders")&"\Images"&"_"&Round(Rnd(100)*1000,0)&"_"&intcounter
			intcounter=intcounter+1
			If (LocWinID <> "") Then
				GPS_ErrorImg = strFilename &".png"
				'Window("HWND:="&LocWinID).CaptureBitmap GPS_ErrorImg
				Desktop.CaptureBitmap GPS_ErrorImg
				'fnReportEvent "info","","","Screenshot taken for browser title:- ["&strBrowserName&"]",UCase("Screenshot:")&"-"&strFilename &".png","SCREENSHOT"
			Else
				'fnReportEvent "warning","fnscreenShot()","","Screenshot not taken as requested WinID not found - "&ResultFolder,"Screenshot not taken for browser","<B> Picture cannot saved!!! - WinID not found </B>"
			End If
		Else
			'fnReportEvent "warning","fnscreenShot()","","Screenshot not taken as requested Result folder not found - "&ResultFolder,"Screenshot not taken for browser","<B> Picture cannot saved!!! - Result folder not found </B>"
		End If
		fnscreenShot = LocWinID
	End Function
	
	Public Function fnCreateFolder(Path_Name,Folder_Name)
		Dim fso, f
		Set fso = CreateObject("Scripting.FileSystemObject")
		If (fso.FolderExists(Path_Name))Then
			'Reporter.ReportEvent 2, "Result folder found", "Result Folder found"
		Else
			Set f = fso.CreateFolder(Path_Name)
		End If
	
		'ResultFolder1 = strProjFolder &"Results\"& strFunctionName
		If (fso.FolderExists(Path_Name & Folder_Name))Then
			'Reporter.ReportEvent 2, ModuleName&"folder found", ModuleName&"Result Folder found"
		Else
			Set f = fso.CreateFolder(Path_Name & Folder_Name)
		End If
		Set fso=Nothing
		Set f=Nothing
		fnCreateFolder = Path_Name & Folder_Name
	End Function
	
	Public Function fnCopyFile(ByVal strSrcDir, ByVal strFileName, ByVal strDestDir)
				
		If Not fnFileExist(strSrcDir & "\" & strFileName) Then
		     fnCopyFile = False
			 Exit Function
		End If
		
		If Not fnFolderExist(strDestDir) Then
			fnCopyFile = False
			Exit Function
		End If
		
		''' <value type="Scripting.FileSystemObject"/>
		Dim oFSO
		
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		
		Call oFSO.CopyFile(strSrcDir & "\" & strFileName, strDestDir)
		
		If Not fnFileExist(strDestDir & "\" & strFileName) Then
			fnCopyFile = False
			Exit Function
		End If
		
		fnCopyFile = True
		
	End Function
	
	Public Function fnCopyFolder(ByVal strSrcDir, ByVal strDestDir, ByVal blnRecursive)
		
		If Not fnFolderExist(strSrcDir) Then
			fnCopyFolder = False
			Exit Function
		End If
		
		If Not fnFolderExist(strDestDir) Then
			fnCopyFolder = False
			Exit Function
		End If
		
		Dim oFSO
		''' <value type="IFolder"/>
		Dim f
		
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		
		Call oFSO.CopyFolder(strSrcDir, strDestDir, True)
		
		If blnRecursive	Then
			Call oFSO.CopyFolder(strSrcDir & "\*", strDestDir, True)
		End If
		
		fnCopyFolder = True
	End Function

	''' <summary>
	''' Kills a process in task manager. Need to send in the process name in single quotes. I.E. 'excel.exe'
	''' </summary>
	''' <author>Kishore</author>
	''' <datecreated>6/29/2012</datecreated>
	''' <param name="strProcess" type="String">Process name to kill</param>
	Public Sub KillProcess(strProcess)
		On Error Resume Next
		Dim strComputer, objWMIService, colProcesses, objProcess
		strComputer = "." 
		Set objWMIService = GetObject("winmgmts:\\"&strComputer&"\root\cimv2") 
		Set colProcesses = objWMIService.ExecQuery ("Select * From Win32_Process Where Name = '"&strProcess&"'") 
		For Each objProcess In colProcesses
			objProcess.Terminate() 
		Next		
	End Sub	
	
	'Shweta - 23/2/2016 - Start
	''' <summary>
	''' To delete IE browser history
	''' </summary>
	Public Sub ClearBrowserHistory()
		Exit sub
		Dim WshShell
		Set WshShell = CreateObject("WScript.Shell")
		'Clear cookies
		WshShell.run "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2"
		'wait 2
		
		'Temporary Internet Files
		WshShell.run "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8"
		'wait 2
		
		'Cookies
		WshShell.run "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2"
		'wait 2
		
'		'History
'		WshShell.run "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 1"
'		wait 2
		
'		'Form Data
'		WshShell.run "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 16"
'		wait 2
'		
'		'Passwords
'		WshShell.run "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 32"
'		wait 2
'		
'		'Delete All
'		WshShell.run "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255"
		
		Set WshShell=Nothing
	End Sub
	'Shweta - 23/2/2016 - End
	
	
	''' <summary>
	''' Removes a key and item from the Dictionary Object.
	''' </summary>
	''' <param name="objData" type="Scripting.Dictionary">The Dictionary object that holds the key</param>
	''' <param name="strFilterKey" type="String">The key to remove</param>
	Public Sub RemoveKey(ByRef objData, ByVal strFilterKey)
		If objData.Exists(strFilterKey) Then
		     objData.Remove(strFilterKey)
		End If	
	End Sub
	
	Private Sub class_Initialize()


	End Sub  
	
End Class


''' <summary>
''' Class used to retrieve values from INI files
''' </summary>
Class ReadINIFile

	Private m_strFileName
	
	''' <summary>
	''' Gets/Sets the filename for the INI file.
	''' </summary>
	''' <value type="String"/>
	Public Property Get FileName()
		FileName = m_strFileName
	End Property
	Public Property Let FileName(strData)
		m_strFileName = strData
	End Property
	
	''' <summary>
	''' Looks up the value of the Section:Key combination provided in the parameters.
	''' </summary>
	''' <param name="strSection" type="String">Section to find the key in</param>
	''' <param name="strKeyName" type="String">Key to look up under the section</param>
	''' <returns type="String">Returns value of the Section:Key combination</returns>
	Public Function GetKeyValue(ByVal strSection, ByVal strKeyName)
		'On Error Resume Next
		Dim strData
		
		extern.GetPrivateProfileString strSection, strKeyName,"", strData, 3072, m_strFileName
		If strData = "" Then
			'GetKeyValue = strData
			GetKeyValue = strKeyName
		Else
			GetKeyValue = strData
		End If
	End Function 
		
	Private Sub Class_Initialize()
		Dim m_FileName
		m_FileName = ""
		Extern.Declare micInteger, "GetPrivateProfileString", "kernel32", "GetPrivateProfileStringA", _
		micString, micString, micString, micString + micByRef, micDWord, micString
	End Sub
	
	
	''' <summary>
	''' Initialized Environment Variables specific to the Community Investment (Web) System Test Cases
	''' </summary>
	Public Sub InitializeVariables_SCAWeb()
		
		'Region Variable Declarations
			''' <value type="String"/>
			Dim strConfigFileFolder
			''' <value type="String"/>
			Dim strConfigFileName	
		'End Region

		'Grab the base directory of the framework
		'**Environment.Value("CurrDir") - Value is passed through QTP.
		'Environment.Value("CurrDir") = "C:\Hexaware\Sakthi\QAFramework3x\"
		strConfigFileFolder = "EnvironmentFiles\"
		
		'Sumit Modification for IAM Subs - 10/10/2017
		
		'strConfigFileName   = "IMSSCAWeb_Config.ini"
		
		  'ConfigurationDirectory=Mid(Environment("TestDir"),1,Instr(1,Environment("TestDir"),"\QAFramework3x\"))
          'Environment.Value("CurrDir")=ConfigurationDirectory&"QAFramework3x\"
          
          
          'Environment.Value("CurrDir")="C:\BIMergedFramework\4AprilQAFramework3x\QAFramework3x\"
          
          
		If UCase(Environment.Value("Testing_Environment")) = "SITLIVE" Then 
			strConfigFileName   = "IMSSCAWeb_Config_SITLive.ini"
		ElseIf UCase(Environment.Value("Testing_Environment")) = "SITNONLIVE" Then
			strConfigFileName   = "IMSSCAWeb_Config_SITNonLive.ini"
		ElseIf UCase(Environment.Value("Testing_Environment")) = "UATLIVE" Then
			strConfigFileName   = "IMSSCAWeb_Config_UATLive.ini"
		ElseIf UCase(Environment.Value("Testing_Environment")) = "UATNONLIVE" Then
			strConfigFileName   = "IMSSCAWeb_Config_UATNonLive.ini"
		ElseIf UCase(Environment.Value("Testing_Environment")) = "DSTLIVE" Then
			strConfigFileName   = "IMSSCAWeb_Config_DSTLive.ini"
		ElseIf UCase(Environment.Value("Testing_Environment")) = "DSTNONLIVE" Then
			strConfigFileName   = "IMSSCAWeb_Config_DSTNonLive.ini"
		ElseIf Environment.Value("Testing_Environment") = "" Then
			strConfigFileName   = "IMSSCAWeb_Config.ini"
		End If 	
		
		'Sumit - End - 10/10/2017

		
		'shweta - start
		FileName= Environment.Value("CurrDir") & strConfigFileFolder & strConfigFileName
		
		'shweta - end
		Environment.Value("InputFiles") = Environment.Value("CurrDir") & "InputFiles\"
		Environment.Value("ResultsPath") = Environment.Value("CurrDir") & "ResultFiles\"
		Environment.Value("StartTime") = Now
		
		Environment.Value("LogoPath") = Environment.Value("CurrDir") & "Logos\IMS.jpg"
		
		
		
		
		
'		'--Test Data Input Files
'	Environment.Value("TestScenariosFile")		= Environment.Value("InputFiles") & GetKeyValue ("FilePath:TestInput","TestScenarios") 
'   Environment.Value("TestCasesFile")			= Environment.Value("InputFiles") & GetKeyValue ("FilePath:TestInput","TestCases") 

  If Instr(UCASE(Environment("TestName")),"_PH1")<>0  Then
     Environment.Value("TestScenariosFile")= Environment.Value("InputFiles") & "IMSSCAWeb\Phase1\TestScenarios.xls"
     Environment.Value("TestCasesFile")= Environment.Value("InputFiles") & "IMSSCAWeb\Phase1\TestCases.xls"
  ElseIf Instr(UCASE(Environment("TestName")),"_PH3")<>0  Then
      Environment.Value("TestScenariosFile")= Environment.Value("InputFiles") & "IMSSCAWeb\Phase3\TestScenarios.xls"
       Environment.Value("TestCasesFile")= Environment.Value("InputFiles") & "IMSSCAWeb\Phase3\TestCases.xls"
  ElseIf Instr(UCASE(Environment("TestName")),"PH2")<>0  Then
        Environment.Value("TestScenariosFile")= Environment.Value("InputFiles") & "IMSSCAWeb\Phase2\TestScenarios.xls"
       Environment.Value("TestCasesFile")= Environment.Value("InputFiles") & "IMSSCAWeb\Phase2\TestCases.xls"
  ElseIf Instr(UCASE(Environment("TestName")),"SCA2359")<>0 or InStrRev(UCASE(Environment("TestName")),"_PH4")<>0 Then
       Environment.Value("TestScenariosFile")= Environment.Value("InputFiles") & "IMSSCAWeb\Phase4\TestScenarios.xls"
       Environment.Value("TestCasesFile")= Environment.Value("InputFiles") & "IMSSCAWeb\Phase4\TestCases.xls"
  ElseIf InStrRev(UCASE(Environment("TestName")),"_PH5")<>0 Then
       Environment.Value("TestScenariosFile")= Environment.Value("InputFiles") & "IMSSCAWeb\Phase5\TestScenarios.xls"
       Environment.Value("TestCasesFile")= Environment.Value("InputFiles") & "IMSSCAWeb\Phase5\TestCases.xls"
  ElseIf InStrRev(UCASE(Environment("TestName")),"_PH6")<>0 or InStrRev(UCASE(Environment("TestName")),"_PH7")<>0 or InStrRev(UCASE(Environment("TestName")),"_PH8")<>0 or InStrRev(UCASE(Environment("TestName")),"_2361")<>0 Then
       Environment.Value("TestScenariosFile")= Environment.Value("InputFiles") & "IMSSCAWeb\Phase6\TestScenarios.xls"
       Environment.Value("TestCasesFile")= Environment.Value("InputFiles") & "IMSSCAWeb\Phase6\TestCases.xls"
  ElseIf InStr(UCASE(Environment("TestName")),"_IAMSUBS")<>0 Then
       Environment.Value("TestScenariosFile")= Environment.Value("InputFiles") & "IMSSCAWeb\IAMSubscription\TestScenarios.xls"
       Environment.Value("TestCasesFile")= Environment.Value("InputFiles") & "IMSSCAWeb\IAMSubscription\TestCases.xls"
  ElseIf InStr(UCASE(Environment("TestName")),"OCRF_")<>0 Then
  	Environment.Value("TestScenariosFile")= Environment.Value("InputFiles") & "IMSSCAWeb\OCRF\TestScenarios.xls"
     	Environment.Value("TestCasesFile")= Environment.Value("InputFiles") & "IMSSCAWeb\OCRF\TestCases.xls"
  End If

'		
		'--System Specific Paths/Objects
		Environment.Value("Application")			= GetKeyValue ("Application:Properties","Application")
		Environment.Value("ApplicationType")		= GetKeyValue ("Application:Properties","ApplicationType")
		Environment.Value("ApplicationShortName")	= GetKeyValue ("Application:Properties","ApplicationShortName")
		Environment.Value("TestEnv")				= GetKeyValue ("Application:Properties","TestEnv")
		Environment.Value("TestDB")					= GetKeyValue ("Application:Properties","TestDB")
		Environment.Value("ReleaseNo")				= GetKeyValue ("Application:Properties","ReleaseNo")
		
		
		
		Environment.Value("SCAReportCreationFile") =  GetKeyValue ("SCA:ReportCreationFile","ReportCreationFile")
		
		'--Application Specific Values
		'Appscript-URL
		Environment.Value("ScaURL") =  GetKeyValue ("SCA:URL","URL")
		'Chnaged by Poornima 2/13/2018
        Environment.Value("OCRFURL") =  GetKeyValue ("OCRF:URL","URL")
        Environment.Value("MSTRURL") =  GetKeyValue ("MSTR:URL","URL")
    	Environment.Value("AUTHURL") =  GetKeyValue ("AUTH:URL","URL")
		Environment.Value("DCURL") =  GetKeyValue ("DC:URL","URL")
		Environment.Value("HostingPath") =  GetKeyValue ("Hosting:Path","Path") 
		Environment.Value("SourceFile") =  GetKeyValue ("Source:File","File")
		Environment.Value("OPSURL") =  GetKeyValue ("OPS:URL","URL")
		Environment.Value("PBIServerURL") =  GetKeyValue ("PBIServer:URL","URL")
		Environment.Value("InternalOPSURL") =  GetKeyValue ("OPSInternal:URLInternal","URLInternal")
		Environment.Value("OwnerId") = GetKeyValue ("SCA:aReportTagKeyval2","aReportTagKeyval2")
		Environment.Value("ReportName") = GetKeyValue ("SCA:aReportTagKeyval1","aReportTagKeyval1")
		Environment.Value("ComponentDataSrcKeyVal1") =  GetKeyValue ("SCA:aCompDataSrcKeyVal1","aCompDataSrcKeyVal1") 
		Environment.Value("ComponentDataSrcKeyVal2") =  GetKeyValue ("SCA:aCompDataSrcKeyVal2","aCompDataSrcKeyVal2")
		Environment.Value("ComponentDataSrcKeyVal3") =  GetKeyValue ("SCA:aCompDataSrcKeyVal3","aCompDataSrcKeyVal3")
		Environment.Value("ComponentDataSrcKeyVal4") =  GetKeyValue ("SCA:aCompDataSrcKeyVal4","aCompDataSrcKeyVal4")
		Environment.Value("DataSrcKeyVal") =  GetKeyValue ("SCA:aDataSrcKeyVal","aDataSrcKeyVal")
		'shweta- Added intCounterStart and intCounterMaxLimit to control while/For loop max values <21/3/2016>- Start
		Environment.Value("intCounterStart") =  GetKeyValue ("SCA:intCounterStart","intCounterStart")
		Environment.Value("intCounterMaxLimit") =  GetKeyValue ("SCA:intCounterMaxLimit","intCounterMaxLimit")
		Environment.Value("URLIAMSubs") = GetKeyValue ("Project:IAMSubscription","URLIAMSubs")
		Environment.Value("URLDownloadCenter") = GetKeyValue ("Project:IAMSubscription","URLDownloadCenter")
		Environment.Value("URLDCSite") = GetKeyValue ("Project:IAMSubscription","URLDCSite")
		Environment.Value("IAMClientName") = GetKeyValue ("SCA:ClientName","IAMClientName")
		Environment.Value("ClientName") = GetKeyValue ("SCA:ClientName","ClientName")
		Environment.Value("cubeName")		= GetKeyValue ("OCRF:cubeName","cubeName")
		Environment.Value("cubeLocation")		= GetKeyValue ("OCRF:cubeLocation","cubeLocation")
		Environment.Value("flaPath")		= GetKeyValue ("OCRF:flaPath","flaPath") 
		
		Environment.Value("SITNONLIVEURL")		= GetKeyValue ("OCRF:SITNONLIVEURL","SITNONLIVEURL")
		Environment.Value("NONLiveURL")		= GetKeyValue ("OCRF:NONLiveURL","NONLiveURL")
		Environment.Value("flaServerPath")		= GetKeyValue ("OCRF:flaServerPath","flaServerPath")
		Environment.Value("HostServerName") = GetKeyValue ("OCRF:HostServerName","HostServerName")
		'shweta- Added intCounterStart and intCounterMaxLimit to control while/For loop max values <21/3/2016>- End
		
		'--Final Save Directory
		Environment.Value("TestResultsFolderName")	= Environment.Value("TestName") & objUtils.fnUniqueDateTimeByVal(Environment.Value("StartTime")) & "\"
		Environment.Value("TestResultsFolderPath")	= Environment.Value("ResultsPath") & Environment.Value("TestResultsFolderName")
			
		If ( Not objUtils.fnFolderExist( Environment.Value("TestResultsFolderPath") ) ) Then
			Call objUtils.fnCreateFolder(Environment.Value("ResultsPath"), Environment.Value("TestResultsFolderName"))
		End If
		
		'--Test Documents Save Directory		
		Environment.Value("TestDocumentsFolderPath") = Environment.Value("TestResultsFolderPath") & "Test_Documents\"
		
		If ( Not objUtils.fnFolderExist( Environment.Value("TestDocumentsFolderPath") ) ) Then
			Call objUtils.fnCreateFolder(Environment.Value("TestResultsFolderPath"), "Test_Documents")
		End If
	
		'System Prep - E.G. Launch PB App or I.E.
		'shweta - 4/1/2016 - Start
		'objUtils.KillProcess("iexplore.exe")
		Systemutil.CloseProcessByName "excel.exe"
		wait 3
		'Systemutil.CloseProcessByName "iexplore.exe"
		'shweta - 4/1/2016 - End
		
		'objUtils.fnLaunchingIE( Environment.Value("ScaURL") )	
		
		'shweta - 23/2/2016 - Start
		objUtils.ClearBrowserHistory()
		'shweta - 23/2/2016 - End
		 
		

	End Sub 		
		
End Class

