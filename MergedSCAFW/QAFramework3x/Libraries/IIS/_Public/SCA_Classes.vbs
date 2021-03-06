Option Explicit


Public SCA : Set SCA = New CIMSSCALibrary



Class CIMSSCALibrary
	
	
	''' <summary>
	''' Retrieves the URL From the browser
	''' </summary>
	''' <returns type="String">The Browser's URL as a String</returns>
	Public Function GetCurrentBrowserURL()
		On Error Resume Next
	
		GetCurrentBrowserURL = Browser("micclass:=Browser").Page("micclass:=Page").GetROProperty("URL")
		
		If Err.Number <> 0 Then
			ReportStep StatusTypes.Fail, "GetCurrentBrowserURL", "Error Found in GetCurrentBrowserURL: " & Err.Description
			Err.Clear
		End If
		
		On Error GoTo 0
	End Function
	
	
	
	
	
		
	''' <summary>
	''' Navigate to IMSAppscriptDashboard home page
	''' </summary>
	''' <param name="strHomePageUrl" type="String (Url)"></param>
	Public Sub NavigateHome( ByVal strHomePageUrl )
		Browser("Browser").Navigate(strHomePageUrl)
		'ValidateIMSAppscriptDashboardLoaded()
	End Sub
	
	''' <summary>
	''' Takes snapshot of the available page
	''' </summary>
	''' <param name="strImageName" type="String">File name of the image to save in the results folder</param>
	Public Sub Screenshot(ByVal strImageName)
		
		''' <value type="String"/>	
		Dim strFileName
		
		If Len(Trim(strImageName)) > 0 Then
			strFileName = strImageName
		Else
			Call ReportStep(StatusTypes.Warning, "Image name cannot be empty." , "Function Screenshot")
			Exit Sub
		End If
		
		If (Not objUtils.fnFolderExist( Environment.Value("TestDocumentsFolderPath") & IMS_Framework_TestCaseHandler.CurrentTestCase.TestCaseID)) Then
			Call objUtils.fnCreateFolder(Environment.Value("TestDocumentsFolderPath") , IMS_Framework_TestCaseHandler.CurrentTestCase.TestCaseID)
		End If
		
'		If Browser("Analyzer").Page("title:=.*").Exist(2) Then
'			Browser("Analyzer").Page("title:=.*").CaptureBitmap (Environment.Value("TestDocumentsFolderPath") & IMS_Framework_TestCaseHandler.CurrentTestCase.TestCaseID &"\"&strFileName&".png")
'		End If

      Desktop.CaptureBitmap (Environment.Value("TestDocumentsFolderPath") & IMS_Framework_TestCaseHandler.CurrentTestCase.TestCaseID &"\"&strFileName&".png")
		
	End Sub
	
	''' <summary>Verifies the IMSAppscript home browser is present</summary>
	Public Sub PerformBrowserCheck()
	
		'Validating the presence of Browser
		If Not Browser("Browser").Exist(1) Then
			Call ReportStep(StatusTypes.Fail, "Error: Browser isn't loaded.",  "Appscript")
			'Relaunching Browser and Loading IMSAppscriptURL for Smooth Execution of Succeeding Scripts
			objUtils.fnLaunchingIE( Environment.Value("AppscriptURL") )			
		End If
		
	End Sub
	
	''' <summary>Click on radio option</summary>
	''' <param name="objRdbGroup" type="Object">Radio group Object</param>
	''' <param name="strRdbOption" type="Scripting.Dictionary">Option to select</param>
	''' <param name="strPageName" type="String">Screen location</param>
	Public Sub ClickOnRadio(ByVal objRdbGroup, ByVal strRdbOption ,ByVal strPageName)

		''' <value type="String"/>
		Dim strRdbVal
		
		strRdbVal=""
		
		If objRdbGroup.Exist(2) Then
			If Trim(UCase(strRdbOption))="ANY" Or Trim(UCase(strRdbOption))="OR" Then
				strRdbVal="or"
				objRdbGroup.Select strRdbVal
			ElseIf Trim(UCase(strRdbOption))="ALL" Or Trim(UCase(strRdbOption))="AND" Then
				strRdbVal="and"
				objRdbGroup.Select strRdbVal			
			ElseIf strRdbVal="" Then
				objRdbGroup.Select strRdbOption
			End If
			Call ReportStep (StatusTypes.Pass, "Selected '"&strRdbOption&"' as radio option", strPageName)
		Else
			Call ReportStep (StatusTypes.Fail, "Radio group does not exist", strPageName)
		End If
			
		
	End Sub	
	
	''' <summary>Set the value in the field</summary>
	''' <param name="objWebEdit" type="Object">Field object</param>
	''' <param name="strFieldName" type="String">Field name of the value to enter</param>
	''' <param name="strInputValue" type="Scripting.Dictionary">Data value to enter</param>
	''' <param name="strPageName" type="String">Screen location</param>
	Public Sub SetTextOldFunction(ByVal objWebEdit, ByVal strInputValue, ByVal strFieldName, ByVal strPageName)
	
	if  objWebEdit.Exist(15) Then 
		objWebEdit.RefreshObject
		If Not(objWebEdit.WaitProperty("html tag", "INPUT")) Then
			Call ReportStep (StatusTypes.Fail, "'"&strFieldName&"' field is not displayed.", strPageName)
			Exit Sub
		End If
	
		If CInt(objWebEdit.GetROProperty("disabled")) = 1 Then
			Call ReportStep (StatusTypes.Fail, "Field '"&strFieldName&"' is disabled. You cannot enter the value.", strPageName)
			Exit Sub
		End If
		
		objWebEdit.Set Trim(strInputValue)
		Call ReportStep (StatusTypes.Pass, "'"&strInputValue&"' is entered in "&strFieldName&" field.", strPageName)
		Wait(2)
		
		Else
		Call ReportStep (StatusTypes.Fail, "'"&strFieldName&"' is not exist in" , strPageName) 
	End if 
	
	End Sub
	
	
	''' <summary>Set the value in the field For Mutiple Values</summary>
	''' <param name="objWebEdit" type="Object">Field object</param>
	''' <param name="strFieldName" type="String">Field name of the value to enter</param>
	''' <param name="strInputValue" type="Scripting.Dictionary">Data value to enter</param>
	''' <param name="strPageName" type="String">Screen location</param>
	Public Sub SetText_MultipleLineArea(ByVal objWebEdit, ByVal strInputValue, ByVal strFieldName, ByVal strPageName)
	
		objWebEdit.RefreshObject
		If Not(objWebEdit.WaitProperty("html tag", "TEXTAREA")) Then
			Call ReportStep (StatusTypes.Fail, "'"&strFieldName&"' field is not displayed.", strPageName)
			Exit Sub
		End If
	
		If CInt(objWebEdit.GetROProperty("disabled")) = 1 Then
			Call ReportStep (StatusTypes.Fail, "Field '"&strFieldName&"' is disabled. You cannot enter the value.", strPageName)
			Exit Sub
		End If
		
		objWebEdit.Set Trim(strInputValue)
		Call ReportStep (StatusTypes.Pass, "'"&strInputValue&"' is entered in "&strFieldName&" field.", strPageName)
		Wait(2)
	
	End Sub
	
	
	
	
	
	''' <summary>Click the link or button</summary>
	''' <param name="objLnkOrButton" type="Object">Link|Button object</param>
	''' <param name="strLnkOrButtonName" type="String">Link|Button name to click</param>
	''' <param name="strPageName" type="String">Screen location</param>
	Public Sub ClickOn(ByVal objLnkOrButton, ByVal strLnkOrButtonName, ByVal strPageName)
	
		
		If Not(objLnkOrButton.Exist) Then
			objLnkOrButton.RefreshObject
			Call ReportStep (StatusTypes.Fail," "&strLnkOrButtonName&" is not displayed.", strPageName)
	     	Exit Sub
		End If
		
		objLnkOrButton.RefreshObject
		objLnkOrButton.Click
		wait 5
		'Browser("Analyzer").Sync
		wait(1.5)
		Call ReportStep (StatusTypes.Pass,"Clicked on "&strLnkOrButtonName&"", strPageName)
			
	End Sub
	
	''' <summary>Returns the value in the field</summary>
	''' <param name="objField" type="Object">Field object</param>
	''' <param name="strFieldName" type="String">Field name of the value to enter</param>
	''' <param name="strPageName" type="String">Screen location</param>
	Public Function GetText(ByVal objField, ByVal strFieldName, ByVal strPageName)
		
		If Not(objField.Exist(2)) Then
			Call ReportStep (StatusTypes.Fail, "Field '"&strFieldName&"' does not exist", strPageName)
			Exit Function
		End If
		
		If Trim(objField.GetROProperty("micclass"))="WebEdit" Then
			GetText = Trim(objField.GetROProperty("value"))
		ElseIf Trim(objField.GetROProperty("micclass"))="WebList" Then
			GetText = Trim(objField.GetROProperty("default value"))
		Else
			Call ReportStep (StatusTypes.Warning, "Object passed to the function is not a WebEdit or WebList. Please check the object for the field : '"&strFieldName&"'", strPageName)	
			GetText="Invalid_Object"
		End If		
					
	End Function
	
	''' <summary>To set a Checkbox Object as checked or unchecked</summary>
	''' <param name="webCheckBoxObject" type="QTPWebCheckBoxObject">WebCheckBox object to set</param>
	''' <param name="strValue" type="String">Value to set. "ON" to check the checkbox and "OFF" to uncheck it.</param>
	''' <param name="strPageName" type="String">Screen location</param>
	Public Sub SetCheckBox(ByVal webCheckBoxObject,ByVal strCheckboxName, ByVal strValue, ByVal strPageName)

		If Not (webCheckBoxObject.Exist(2)) Then
			Call ReportStep (StatusTypes.Fail, "Field '"&strCheckboxName&"' does not exist", strPageName)		
			Exit Sub
		End If
		If UCase(strValue)<>"ON" And UCase(strValue)<>"OFF" Then
			Call ReportStep (StatusTypes.Warning, "Invalid value: '"&strValue&"' to set checkbox. Please use 'ON' or 'OFF'.", strPageName)
			Exit Sub
		End If
		webCheckBoxObject.Set strValue
		If UCase(strValue) = "ON" And webCheckBoxObject.GetROProperty("checked") = 1 Then
			Call ReportStep (StatusTypes.Pass, "Checked: '"&strCheckboxName&"'", strPageName)
		ElseIf UCase(strValue) = "OFF" And webCheckBoxObject.GetROProperty("checked") = 0 Then
			Call ReportStep (StatusTypes.Pass, "Un-Checked: '"&strCheckboxName&"'", strPageName)
		Else
			Call ReportStep (StatusTypes.Fail, "Set CheckBox action failed.", strPageName)
		End If

	End Sub
	
	''' <summary>Selects the value from dropdown</summary>
	''' <param name="strSelectValue" type="String">Value to select from the dropdown</param>
	Public Function SelectFromDropdown(ByVal objWebList, ByVal strSelectValue)
			
		''' <value type="Integer"/>
		Dim listCount, i,k
		''' <value type="Boolean"/>
		Dim found	
		For k=1 to 1000
			'<Shweta 3/6/2016> -Added wait stmt- Start
			wait 2
			'<Shweta 3/6/2016> -Added wait stmt- End
          If objWebList.GetROProperty("disable")=0  Then
			Exit For
		End If
		Next
     
		
		objWebList.RefreshObject		
		listCount = objWebList.GetROProperty("items count")
		i = 1
		found = False
		If UCase(objWebList.GetROProperty("select type")) <> "COMBOBOX SELECT" Then
			Reporter.ReportEvent micFail,"Non dropdown field object sent to SelectFromDropdown","Reusable Function - SelectFromDropdown"
			SelectFromDropdown = False
		End If
		
		While i <= listCount And Not found
			objWebList.RefreshObject
            wait 1			
			If Trim(objWebList.GetItem(i)) = Trim(strSelectValue) Then
				objWebList.Select i-1
				wait(2)
				found = True
			End If
			i = i + 1						
		Wend
		
		objWebList.RefreshObject
		If Trim(objWebList.GetROProperty("selection")) = Trim(strSelectValue)Then
			SelectFromDropdown = True
		Else
			SelectFromDropdown = False
		End If
	
	End Function
	
	''' <summary>Gets the random number</summary>
	''' <param name="LengthOfRandomNumber" type="Integer">Specifying the length of the random number to generate</param>
	Public Function GetRandomNumber(ByVal LengthOfRandomNumber)

		Dim sMaxVal : sMaxVal = ""
		Dim iL,iTmp,iLen,iLength : iLength = LengthOfRandomNumber
		
		'Find the maximum value for the given number of digits
		For iL = 1 to iLength
			sMaxVal = sMaxVal & "9"
		Next
		sMaxVal = Int(sMaxVal)
		'Find Random Value
		Randomize
		iTmp = Int((sMaxVal * Rnd) + 1)
		'Add Trailing Zeros if required
		iLen = Len(iTmp)
		GetRandomNumber = iTmp * (10 ^(iLength - iLen))
	
	End Function
	
	
	
	Private Sub wScript()
         
     
	End Sub
	''' <summary>To Verify the Existance of the Webtable and  Working with the Webtable</summary>
	''' <param name="webCheckBoxObject" type="QTPWebCheckBoxObject">WebCheckBox object to set</param>
	''' <param name="strValue" type="String">Value to set. "ON" to check the checkbox and "OFF" to uncheck it.</param>
	''' <param name="strPageName" type="String">Screen location</param>
	Function  Webtable(ByVal webtableObject,ByVal strWebtableName,ByVal strOperationSelection,ByVal strPageName,ByVal intCompColumnName,ByVal intClickColumnName,ByVAl strEXpcellValue,ByVAl strElementType)
	
	   Dim rc,cc,i,j,strcellValue,strActvaluecell, checkBoxValue
	   Dim resArr()
	   Dim objClick
	   checkBoxValue = 0
	   
	   Select Case strOperationSelection
	   	
	   	Case "Retriving_DataTableValue"
	   	
		   	If Not (webtableObject.Exist(2)) Then
			 Call ReportStep (StatusTypes.Fail, "Field '"&strWebtableName&"' does not exist", strPageName)		
			 Exit Function
			End If
			
			rc=webtableObject.rowcount
	        cc=webtableObject.Columncount(rc)
			ReDim resArr(rc,cc)
			For i=1 to rc
				For j=1 to cc
				 resArr(i-1,j-1)=webtableObject.getcelldata(i,j)
				Next
			Next
		    Webtable=resArr
	   	
	   	Case "ClickOn_ROW_Condition"
	   	
		   	If Not (webtableObject.Exist(2)) Then
			 Call ReportStep (StatusTypes.Fail, "Field '"&strWebtableName&"' does not exist", strPageName)		
			 Exit Function
			End If
		   	
'		   	'Shweta- Start
'		   	Do
'			   	rc=webtableObject.rowcount
'		        cc=webtableObject.Columncount(rc)
'		        
'		        For i= 1 To rc  
'		        
'		           strActvaluecell=webtableObject.getcelldata(i,intCompColumnName)           
'		           If Strcomp(TRIM(strActvaluecell),TRIM(strEXpcellValue))=0 Then
'		           
'		           	set objClick=webtableObject.ChildItem(i,intClickColumnName,strElementType,0)
'		           	objClick.Click
'		           	checkBoxValue = 1
'		           	Exit For
'		           	
'		           End If 
'		        	
'		        Next
'		        
'				If checkBoxValue = 0 Then
'					set objNextPageClick= Browser("Analyzer").Page("Home").Frame("Frame").WebElement("welNextPage").Click
'					Browser("Analyzer").Page("Home").Sync	
'				End If
'
'			Loop While checkBoxValue = 0
'			
'			If checkBoxValue = 0 Then
'					Call ReportStep (StatusTypes.Fail, "Could Not select member "&strEXpcellValue& " successfully", "")		
'					Exit Sub
'			End If
'		   	'Shweta- End
		   	
		   	rc=webtableObject.rowcount
	        cc=webtableObject.Columncount(rc)
	        
	        For i= 1 To rc  
	        
	           strActvaluecell=webtableObject.getcelldata(i,intCompColumnName)           
	           If Strcomp(TRIM(strActvaluecell),TRIM(strEXpcellValue))=0 Then
	           
	           	set objClick=webtableObject.ChildItem(i,intClickColumnName,strElementType,0)
	           	objClick.Click
	           	Exit For
	           	
	           End If 
	        	
	        Next  	   	
	   	
	   End Select	

	End Function
    
	Private Sub Class_Initialize()
		
	End Sub



''' <summary>Set the value in the field</summary>
	''' <param name="objWebEdit" type="Object">Field object</param>
	''' <param name="strFieldName" type="String">Field name of the value to enter</param>
	''' <param name="strInputValue" type="Scripting.Dictionary">Data value to enter</param>
	''' <param name="strPageName" type="String">Screen location</param>


Public Sub SetText(ByRef objWebEdit, ByVal strInputValue, ByVal strFieldName, ByVal strPageName)
	
	if  objWebEdit.Exist(30) Then 
		objWebEdit.RefreshObject
'		If Not(objWebEdit.WaitProperty("html tag", "INPUT")) Then
'			Call ReportStep (StatusTypes.Fail, "'"&strFieldName&"' field is not displayed.", strPageName)
'			Exit Sub
'		End If
	
'		If CInt(objWebEdit.GetROProperty("disabled")) = 1 Then
'			Call ReportStep (StatusTypes.Fail, "Field '"&strFieldName&"' is disabled. You cannot enter the value.", strPageName)
'			Exit Sub
'		End If
		
		objWebEdit.Set Trim(strInputValue)
		Call ReportStep (StatusTypes.Pass, "'"&strInputValue&"' is entered in "&strFieldName&" field.", strPageName)
		Wait(2)
		
		Else
		Call ReportStep (StatusTypes.Fail, "'"&strFieldName&"' is not exist in" , strPageName) 
	End if 
	
	End Sub
End Class
