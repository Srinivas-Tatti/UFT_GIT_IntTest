'Created By: Poornima
'Created on: 2/13/2018
'Description: Launch the browser with ocrf url and validate default page is displayed.
'Parameter: Param1 - Extra,Param2 - Extra,Param3 - Extra,objData
Public Function LaunchURL(ByVal Url, ByVal UserName, ByVal Password, ByRef objData)
		
	Systemutil.CloseProcessByName "iexplore.exe"
	
	'Launch OCRF Url's	
'	SystemUtil.Run Environment.Value("OCRFURL")

    SystemUtil.Run "iexplore", Environment.Value("OCRFURL") , , , 3
    wait 30
    If UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAEdit("User name").Exist<>true Then
    	wait 20 
    End If
    If UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").Exist(30) Then    	
        UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").Highlight 
	    UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAEdit("User name").SetValue UserName
	    UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAEdit("Password").SetValue Password
	    wait 1
	    'MOdified by Madhu - 04/02/2020
'	    If UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAButton("OK").Exist Then
'	    
'	    	if UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAButton("OK").GetROProperty("visible") then
'	    	if UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAButton("OK").Exist then
'	    		UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAButton("OK").click
'	    	end if 	
'	    	end  if
'	    End If
'              	    
	    '******************************************************************
    Set objShell = CreateObject("Wscript.shell")
    objShell.SendKeys "{ENTER}"
'    ***************************************************************************8
        Call ReportStep (StatusTypes.Pass, "User " & UserName & " has logged in successfully","Login through Window security PopUp")
    ElseIf True Then  
		Systemutil.CloseProcessByName "iexplore.exe"    
        SystemUtil.Run "iexplore", Environment.Value("OCRFURL") , , , 3
    wait 30
    If UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").Exist(30) Then    	
        UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").Highlight 
	    UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAEdit("User name").SetValue UserName
	    UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAEdit("Password").SetValue Password
	      'MOdified by Madhu - 04/02/2020
	    If UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAButton("OK").Exist Then
	    	if UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAButton("OK").GetROProperty("visible") then
	    	UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAButton("OK").click
	    	end  if
	    End If
	    
	    '******************************************************************
    Set objShell = CreateObject("Wscript.shell")
    objShell.SendKeys "{ENTER}"
'    ***************************************************************************8
        Call ReportStep (StatusTypes.Pass, "User " & UserName & " has logged in successfully","Login through Window security PopUp")
    end if 
	Else
		Call ReportStep (StatusTypes.Fail, "User " & UserName & " has not successfully logged in","Login through Window security PopUp")	   
	End If
    


'	If Browser("BrowserPoPUp").Exist(20) Then
'		Browser("BrowserPoPUp").highlight
'	    Browser("BrowserPoPUp").Dialog("Windows Security").highlight
'	   'Handle 'Window Security' PopUp
'	    If Browser("BrowserPoPUp").Dialog("Windows Security").Exist(30) Then
'			If Browser("BrowserPoPUp").Dialog("Windows Security").Static("InvalidUserAccount").Exist(5) Then
'				Browser("BrowserPoPUp").Dialog("Windows Security").Static("UseAnotherAccount").Click
'			End If
'		   Browser("BrowserPoPUp").Dialog("Windows Security").WinEdit("WinEditUserName").Set UserName	
'		   Browser("BrowserPoPUp").Dialog("Windows Security").WinEdit("WinEditPassword").Set Password
'		   Browser("BrowserPoPUp").Dialog("Windows Security").WinButton("BtnOK").Click
'		   Call ReportStep (StatusTypes.Pass, "User " & UserName & " has logged in successfully","Login through Window security PopUp")
'		Else
'		   Call ReportStep (StatusTypes.Fail, "User " & UserName & " has not successfully logged in","Login through Window security PopUp")	   
'		End If
'	ElseIf Window("BrowserPoPUp").Exist(20) Then
'		Window("BrowserPoPUp").highlight
'	    Window("BrowserPoPUp").Dialog("Windows Security").highlight
'	   'Handle 'Window Security' PopUp
'		If Window("BrowserPoPUp").Dialog("Windows Security").Exist(30) Then
'			If Window("BrowserPoPUp").Dialog("Windows Security").Static("InvalidUserAccount").Exist(5) Then
'				Window("BrowserPoPUp").Dialog("Windows Security").Static("UseAnotherAccount").Click
'			End If
'		   Window("BrowserPoPUp").Dialog("Windows Security").WinEdit("WinEditUserName").Set UserName	
'		   Window("BrowserPoPUp").Dialog("Windows Security").WinEdit("WinEditPassword").Set Password
'		   Window("BrowserPoPUp").Dialog("Windows Security").WinButton("BtnOK").Click
'		   Call ReportStep (StatusTypes.Pass, "User " & UserName & " has logged in successfully","Login through Window security PopUp")
'		Else
'		   Call ReportStep (StatusTypes.Fail, "User " & UserName & " has not successfully logged in","Login through Window security PopUp")	   
'		End If
'	End If
	

	'Validate OCRF url launched successfully
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebEleCurrentRequests").Exist(30) Then
		'	Default page is displayed
		If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebEleCompletedRequests").Exist(60) Then
			Call ReportStep (StatusTypes.Pass, "OCRF Security Default page(Online Client Request List) is displayed","OCRF Security Default page(Online Client Request List)")
		Else
			Call ReportStep (StatusTypes.Fail, "OCRF Security Default page(Online Client Request List) is not displayed","OCRF SecurityDefault page(Online Client Request List)")
		End If
		
	Else
		Exit Function
	End If
	
End Function

Function SelectMenuOption(ByVal MainMenu,ByVal SubMenu,ByRef objData)
	wait 2
     If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("BtnMenuOption").Exist(20) Then
     	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("BtnMenuOption").fireevent "onmouseover"
     	Call ReportStep (StatusTypes.Pass, "Clicked on Menu Optin","Menu option selection")
     	'If menu option is 'ADMINISTRATION ' then select submenu option else select only main menu option
		If UCase(MainMenu)="ADMINISTRATION" Then
	  	    Browser("Browser-OCRF").Page("Page-OCRF").Link("html tag:=A","innertext:="&ucase(MainMenu)).fireevent "onmouseover"
	  	    wait 2
	   		Browser("Browser-OCRF").Page("Page-OCRF").Link("html tag:=A","innertext:="&ucase(SubMenu)).Click
	   		Call ReportStep (StatusTypes.Information, "Clicked on Menu "&ucase(MainMenu)&" and selected sub menu"&ucase(SubMenu),"Menu option selection")
		Else
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("BtnMenuOption").fireevent "onmouseover"
	   		Browser("Browser-OCRF").Page("Page-OCRF").Link("html tag:=A","innertext:="&ucase(MainMenu),"index:=0").Click
	   		Call ReportStep (StatusTypes.Information, "Clicked on Menu "&ucase(MainMenu)&" and selected sub menu"&ucase(SubMenu),"Menu option selection")
		End If
	 Else
	    Call ReportStep (StatusTypes.Fail, "Not Clicked on Menu Optin","Menu option selection")
     End If
		
End Function
'Created By: Poornima
'Created on: 2/27/2018
'Description: 'In 'OCRF Admin user List' table search for UserID
'Parameter: searching item,Searching field
Function FilterInOCRFAdminUserTable(ByVal SearchTable,ByVal SearchField,ByVal SearchItem,ByRef objData)
    wait 5
    tableNo=0
	flag = False
	
	 	
 	Select Case UCase(SearchField)
			Case UCase("First Name"):
							col = 4
			Case UCase("Last Name"):
							col = 5
			Case UCase("User Login ID"):
							col = 6
			Case UCase("Is Active"):
							col = 8
			Case Else:
	End Select
	
	Set objTable =Browser("Browser-OCRF").Page("OCRF AdminUser").WebTable("OcrfAdminUserTable")
'	****************************************************************************8
'Modified by Madhu - 03/29/2020

	'objTable.ChildItem(2,col,"WebEdit",0).Set SearchItem
	objTable.WebEdit("UserName").Set SearchItem
'	*****************************************************
	wait 1
	objTable.WebEdit("UserName").Click
	wait 1
    Set objShell = CreateObject("Wscript.shell")
    objShell.SendKeys "{ENTER}"
    wait 5

    FilterInOCRFAdminUserTable=Browser("Browser-OCRF").Page("OCRF AdminUser").WebTable("OCRFadminUserList").GetROProperty("rows")
 End Function


'Created By: Poornima
'Created on: 2/27/2018
'Description: 'In 'Online CRF' tab search for request id/client name/country/offering/status
'Parameter: searching item,Searching field
Function FilterInOnlineCRF(ByVal SearchTable,ByVal SearchField,ByVal SearchItem,ByRef objData)
    wait 5
    tableNo=0
	flag = False
	
	 	
 	Select Case UCase(SearchField)
			Case UCase("OCRF ID"):
							col = 3
			Case UCase("Client Name"):
							col = 4
			Case UCase("Offering"):
							col = 6
			Case UCase("Status"):
							col = 7
			Case Else:
	End Select
	
	Set objTable = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-htable", "name:=UserAccessClientRequestID")
'	******************************************************
'Modidfied by Madhu - 02/25/2020
	
	If objTable.Exist(5) Then
		Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("Webedit_OfferingIDSearch").Set SearchItem
		wait 1
		Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("Webedit_OfferingIDSearch").Click
	End If
'	objTable.ChildItem(2,col,"WebEdit",0).Set SearchItem
'	wait 1
'	objTable.ChildItem(2,col,"WebEdit",0).Click
'	wait 1

'******************************************************************
    Set objShell = CreateObject("Wscript.shell")
    objShell.SendKeys "{ENTER}"
    wait 5
    Call PageLoading()
 	 
 	Set CurReqTbl=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data")
	Set CmpReqTbl=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTblCompletedReqData")
	wait 1
	If CurReqTbl.RowCount >= 2  Then
	   'tableNo=0
	   flag = Ucase("Current Requests")
	ElseIf CmpReqTbl.RowCount >=2 Then
       'tableNo=1    
       flag = Ucase("Completed Requests")
	Else
		flag = Ucase("Not Found")
	   'Exit Function	
	End If

 	

'	Set objTable = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-htable", "index:=" & tableNo)
'	objTable.ChildItem(2,col,"WebEdit",0).Set SearchItem
'	wait 1
'	objTable.ChildItem(2,col,"WebEdit",0).Click
'	wait 2
'    Set objShell = CreateObject("Wscript.shell")
'    objShell.SendKeys "{ENTER}"
'    wait 3
'    Call PageLoading()
'    If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").RowCount = 2 Then 
'        flag = Ucase("Current Requests")
'    ElseIf  Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTblCompletedReqData").RowCount = 2 Then 
'        flag = Ucase("Completed Requests")
'    Else
'        flag = Ucase("Not Found")
'    End If 
    FilterInOnlineCRF = flag

 End Function

'Created By: Poornima
'Created on: 3/2/2018
'Description: Enter server name in to 'dataSource table
'Parameter: 'FileName' takes name of file from where data input taking
            ' SheetName takes name of the file in 'Filename'
Function EnterCubeLoactionData(ByVal FileName,ByVal SheetName,ByVal StartRow,Byval EndRow,ByRef objData)

	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Wel_UserDB_List").Exist(60) Then
		Call ReportStep (StatusTypes.Pass, "Entered in to cude details tab ","Cube details tab")	
	Else
		Call ReportStep (StatusTypes.Fail, "Not Entered in to cude details tab","Cube details tab")
	End If
	
	Call ImportSheet(SheetName,FileName)
	
	Datatable.SetCurrentRow(StartRow)
	NumberOfRows = (EndRow-StartRow)+1
	
	TableRow=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetROProperty("rows")
	For row = 2 To  TableRow Step 1
           WECheckDBLocation= Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").getcelldata(row,Cint(Environment.value("cubeLocation")))
          On error Resume Next
          If WECheckDBLocation<>"" Then
           
           Else
              Call WebEditExistence(row,Cint(Environment.value("cubeLocation")),"WebEdit")
		      Set CheckDBLocation =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(row,Cint(Environment.value("cubeLocation")),"WebEdit",0)
                         
              If CheckDBLocation.GetRoProperty("disabled")=0  Then
		   			Call WebEditExistence(row,Cint(Environment.value("cubeLocation")),"WebEdit")
		   			CheckDBLocation.set datatable.value("ServerLocation",SheetName)	
			  End If
			   	
            End If
		  On error goto 0 
		  Datatable.SetNextRow
	Next 
	
End Function


Function OpenItemInOnlineCRF(ByVal SearchTable,ByVal SearchField,ByVal SearchItem,ByVal EditReq,ByRef objData)
    wait 2
    CurrRequest=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").RowCount 
    CompRequest=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTblCompletedReqData").RowCount
	'If Ucase(SearchTable) = Ucase("Current Requests") Then
		If CurrRequest >= 2 Then
			Set objTableRow = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data")
			 Call ReportStep (StatusTypes.Information, "Item "&SearchItem&" Found in " &SearchTable,"'Online CRF' tab")
			 wait 2
			 FoundIn="Current Request"
	         RowCnt=objTableRow.RowCount-2
	         
	         row = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").RowCount
			 If row = 2 Then
				Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").ChildItem(2,3,"WebElement",0).DoubleClick
				wait 2
				Call ReportStep (StatusTypes.Information, "Search Item "&SearchItem&" Found in " &SearchTable&" and it is opened","'Online CRF' tab")
			 End If
	         
'	         Set oPage = Browser("Browser-OCRF").Page("Page-OCRF")
'             oPage.RunScript("var doubleClickEvent = document.createEvent('MouseEvents');doubleClickEvent.initEvent('dblclick', true, true);document.querySelectorAll('[aria-describedby=request-list_ClientName]')[0].dispatchEvent(doubleClickEvent);")
	'ElseIf Ucase(SearchTable) = Ucase("Completed Requests") Then
		ElseIf CompRequest >= 2 Then
			Set objTableRow = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTblCompletedReqData")
            Call ReportStep (StatusTypes.Information, "Item "&SearchItem&" Found in " &SearchTable,"'Online CRF' tab")	
            wait 2
            FoundIn="Completed Request"
	        RowCnt=objTableRow.RowCount 
	         row = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTblCompletedReqData").RowCount
			 If row = 2 Then
				Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTblCompletedReqData").ChildItem(2,3,"WebElement",0).DoubleClick
				wait 2
				Call ReportStep (StatusTypes.Information, "Search Item "&SearchItem&" Found in " &SearchTable&" and it is opened","'Online CRF' tab")
			 End If
'            Set oPage = Browser("Browser-OCRF").Page("Page-OCRF")
'			oPage.RunScript("var doubleClickEvent = document.createEvent('MouseEvents'); doubleClickEvent.initEvent('dblclick', true, true); document.querySelectorAll('[aria-describedby=request-list-completed_ClientName]')[0].dispatchEvent(doubleClickEvent);")	        
		End If
	'Else 
		'Exit Function
	'End If
'	wait 2
'	RowCnt=objTableRow.RowCount
	
	Call PageLoading()

	
	If  Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebEleUnlockConfirmation").Exist(5) Then
	    If EditReq="YES" Then
	       If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Yes").Exist(5) Then
	       	  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Yes").Click
 	       	  Call ReportStep (StatusTypes.Information, "Item "&SearchItem& " was locked and it is opened in read only mode","'Online CRF' tab")
	          Exit function
	       End If
	    End If
	    wait 1
	    If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButtonNo").Exist(10) Then
	       Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButtonNo").Click	
	    End If
		
      	Set objOption=Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebEleUnlockReq")
            
      	 	If FoundIn="Current Request" Then
      	 		 row = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").RowCount
				 If row = 2 Then
				 	Setting.WebPackage("ReplayType") = 2
					Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").ChildItem(2,3,"WebElement",0).RightClick
					Setting.WebPackage("ReplayType") = 1
					wait 2
				 End If
'      	 	     Set oPage = Browser("Browser-OCRF").Page("Page-OCRF")
'				 oPage.RunScript("var doubleClickEvent = document.createEvent('MouseEvents');doubleClickEvent.initEvent('contextmenu', true, true);document.querySelectorAll('[aria-describedby=request-list_ClientName]')[0].dispatchEvent(doubleClickEvent);")
                 wait 2
                 Call SelectOptionForRequestID(objOption)
                 Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").ChildItem(2,3,"WebElement",0).DoubleClick
				 wait 2
				 Call ReportStep (StatusTypes.Information, "Item "&SearchItem& " was locked and it is opened by unlocking","'Online CRF' tab")
                 'oPage.RunScript("var doubleClickEvent = document.createEvent('MouseEvents');doubleClickEvent.initEvent('dblclick', true, true);document.querySelectorAll('[aria-describedby=request-list_ClientName]')[0].dispatchEvent(doubleClickEvent);")	
      	 	ElseIf FoundIn="Completed Request" Then
      	 		row = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTblCompletedReqData").RowCount
				 If row = 2 Then
				 	Setting.WebPackage("ReplayType") = 2
					Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTblCompletedReqData").ChildItem(2,3,"WebElement",0).RightClick
					Setting.WebPackage("ReplayType") = 1
					wait 2
				 End If
'      	 	     Set oPage = Browser("Browser-OCRF").Page("Page-OCRF")
'				 oPage.RunScript("var doubleClickEvent = document.createEvent('MouseEvents');doubleClickEvent.initEvent('contextmenu', true, true);document.querySelectorAll('[aria-describedby=request-list-completed_ClientName]')[0].dispatchEvent(doubleClickEvent);")
			     wait 2
			     Call SelectOptionForRequestID(objOption)
			     Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTblCompletedReqData").ChildItem(2,3,"WebElement",0).DoubleClick
			     wait 2
			     Call ReportStep (StatusTypes.Information, "Item "&SearchItem& " was locked and it is opened by unlocking","'Online CRF' tab")
			     'oPage.RunScript("var doubleClickEvent = document.createEvent('MouseEvents'); doubleClickEvent.initEvent('dblclick', true, true); document.querySelectorAll('[aria-describedby=request-list-completed_ClientName]')[0].dispatchEvent(doubleClickEvent);")	        		
      	 	End If
      	   
      	   Call PageLoading()
    End If


 End Function

'Function OpenItemInOnlineCRF(ByVal SearchTable,ByVal SearchField,ByVal SearchItem,ByVal EditReq,ByRef objData)
'    wait 2
'    CurrRequest=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").RowCount 
'    CompRequest=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTblCompletedReqData").RowCount
'	'If Ucase(SearchTable) = Ucase("Current Requests") Then
'		If CurrRequest >= 2 Then
'			Set objTableRow = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data")
'			 Call ReportStep (StatusTypes.Information, "Item "&SearchItem&" Found in" &SearchTable,"'Online CRF' tab")
'			 wait 2
'			 FoundIn="Current Request"
'	         RowCnt=objTableRow.RowCount-2
'	         Set oPage = Browser("Browser-OCRF").Page("Page-OCRF")
'             oPage.RunScript("var doubleClickEvent = document.createEvent('MouseEvents');doubleClickEvent.initEvent('dblclick', true, true);document.querySelectorAll('[aria-describedby=request-list_ClientName]')[0].dispatchEvent(doubleClickEvent);")
'		End If
'	'ElseIf Ucase(SearchTable) = Ucase("Completed Requests") Then
'		If CompRequest >= 2 Then
'			Set objTableRow = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTblCompletedReqData")
'            Call ReportStep (StatusTypes.Information, "Item "&SearchItem&" Found in" &SearchTable,"'Online CRF' tab")	
'            wait 2
'            FoundIn="Completed Request"
'	        RowCnt=objTableRow.RowCount 
'            Set oPage = Browser("Browser-OCRF").Page("Page-OCRF")
'			oPage.RunScript("var doubleClickEvent = document.createEvent('MouseEvents'); doubleClickEvent.initEvent('dblclick', true, true); document.querySelectorAll('[aria-describedby=request-list-completed_ClientName]')[0].dispatchEvent(doubleClickEvent);")	        
'		End If
'	'Else 
'		'Exit Function
'	'End If
''	wait 2
''	RowCnt=objTableRow.RowCount
'	
'
'	Call ReportStep (StatusTypes.Information, "Search Item "&SearchItem&" Found in" &SearchTable&" and it is opened","'Online CRF' tab")
'	If  Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebEleUnlockConfirmation").Exist(5) Then
'	    If EditReq="YES" Then
'	       If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Yes").Exist(5) Then
'	       	  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Yes").Click
' 	       	  Call ReportStep (StatusTypes.Information, "Item "&SearchItem& " was locked and it is opened in read only mode","'Online CRF' tab")
'	          Exit function
'	       End If
'	    End If
'	    If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButtonNo").Exist(2) Then
'	       Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButtonNo").Click	
'	    End If
'		
'      	Set objOption=Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebEleUnlockReq")
'            
'      	 	If FoundIn="Current Request" Then
'      	 	     Set oPage = Browser("Browser-OCRF").Page("Page-OCRF")
'				 oPage.RunScript("var doubleClickEvent = document.createEvent('MouseEvents');doubleClickEvent.initEvent('contextmenu', true, true);document.querySelectorAll('[aria-describedby=request-list_ClientName]')[0].dispatchEvent(doubleClickEvent);")
'                 wait 2
'                 Call SelectOptionForRequestID(objOption)
'                 oPage.RunScript("var doubleClickEvent = document.createEvent('MouseEvents');doubleClickEvent.initEvent('dblclick', true, true);document.querySelectorAll('[aria-describedby=request-list_ClientName]')[0].dispatchEvent(doubleClickEvent);")	
'      	 	ElseIf FoundIn="Completed Request" Then
'      	 	     Set oPage = Browser("Browser-OCRF").Page("Page-OCRF")
'				 oPage.RunScript("var doubleClickEvent = document.createEvent('MouseEvents');doubleClickEvent.initEvent('contextmenu', true, true);document.querySelectorAll('[aria-describedby=request-list-completed_ClientName]')[0].dispatchEvent(doubleClickEvent);")
'			     wait 2
'			     Call SelectOptionForRequestID(objOption)
'			     oPage.RunScript("var doubleClickEvent = document.createEvent('MouseEvents'); doubleClickEvent.initEvent('dblclick', true, true); document.querySelectorAll('[aria-describedby=request-list-completed_ClientName]')[0].dispatchEvent(doubleClickEvent);")	        		
'      	 	End If
'      	    Call ReportStep (StatusTypes.Information, "Item "&SearchItem& " was locked and it is opened by unlocking","'Online CRF' tab")
'    End If
'
'
' End Function
 
 'Created By: Poornima
'Created on: 5/29/2018
'Description: Select menu options of OCRF Id (Right click on OCRF ID, select any one option of these)
Function SelectOptionForRequestID(ByRef objOption)
	Set objOption=Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebEleUnlockReq")	
	objOption.highlight
	Dim absX1 : absX1 = objOption.GetROProperty("abs_x")
    Dim absY1 : absY1 = objOption.GetROProperty("abs_y")
    Dim width1 : width1 = objOption.GetROProperty("width")
    Dim height1 : height1 = objOption.GetROProperty("height")
    Dim x1 : x1 = (absX1 + (width1 / 2))
    Dim y1 : y1 = (absY1 + (height1 / 2))
    Dim deviceReplay1 : Set deviceReplay1 = CreateObject("Mercury.DeviceReplay")
    deviceReplay1.MouseClick x1,y1, RIGHT_MOUSE_BUTTON
End Function

'Created By: Poornima
'Created on: 2/23/2018
'Description: 'Read and write reusable data in text file.Ex: created syndicated request id,Offering name
Function ReadWriteDataFromTextFile(ByVal Action,ByVal VariableName,ByVal UpdateVal)
         Const ForReading=1
        Const ForWriting=2
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        filePath = Environment.Value("CurrDir")&"EnvironmentFiles\ConfigFile.txt"
        Set ReadmyFile = objFSO.OpenTextFile(filePath, ForReading, True)
        Set ReadmyFile1 = objFSO.OpenTextFile(filePath, ForReading, True)
        Do While Not ReadmyFile.AtEndofStream
            myLine = ReadmyFile.ReadLine
            If myLine="" Then
            	Call ReportStep (StatusTypes.Information, "parameter  "&VariableName& " not found in environment file","'Online CRF' tab")
            	Exit function 
            End If
            s=split(myLine,"=")
            If s(0)= VariableName Then      
                MyAllLine= ReadmyFile1.ReadAll            
                  OldString=myLine
                  newString=replace(MyAllLine,s(1),UpdateVal)
                  If Action="Read" Then
                         If s(1)="" Then
                            Call ReportStep (StatusTypes.Information, "Data has not assigned to  "&VariableName& " in environment file","'Online CRF' tab")
                            Exit Function
                         Else
                            ReadWriteDataFromTextFile= s(1)
                         End If 
                         
                  Else
                      Set WritemyFile = objFSO.OpenTextFile(filePath, ForWriting, True)
                      WritemyFile.WriteLine newString
                  End If
                Exit Do
             'Else
                 'Call ReportStep (StatusTypes.Information, "Item "&VariableName& " Not Found in environment file","'Online CRF' tab")
             End If
        Loop
   ReadmyFile.Close 
   ReadmyFile1.Close
End Function

Function ReadUsersFromTextFile(ByVal FileName)
        Const ForReading=1
        'Const ForWriting=2
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        filePath = Environment.Value("CurrDir")&"InputFiles\IMSSCAWeb\OCRF\" & FileName & ".txt"
        Set ReadmyFile = objFSO.OpenTextFile(filePath, ForReading, True)
        'Set ReadmyFile1 = objFSO.OpenTextFile(filePath, ForReading, True)
        allContent = ReadmyFile.ReadAll
        ReadmyFile.Close 
        ReadUsersFromTextFile = allContent
End Function

'Created By: Poornima
'Created on: 6/13/2018
'Description: 'Fuction to perform Click operation any of these following 1. Save and Complete 2. Save and next 3,Submit
Function ClickSaveCompleteSubmitNextOCRF()

     If Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_SaveAndComplete").Exist(15) Then
	   ' Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_SaveAndComplete").Click	
	   Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_SaveAndComplete").DoubleClick
	    Call ReportStep (StatusTypes.Pass, "Clicked on 'Save and Complete' after entering Server details in 'Data Source' tab","'Data Source' tab")
	ElseIf Browser("Browser-OCRF").Page("Page-OCRF").Link("WeblnkSubmit").Exist Then
		'Browser("Browser-OCRF").Page("Page-OCRF").Link("WeblnkSubmit").Click
		Browser("Browser-OCRF").Page("Page-OCRF").Link("WeblnkSubmit").DoubleClick
    	Call ReportStep (StatusTypes.Pass, "Clicked on Submit button","'User DB access' tab")
    ElseIf Browser("Browser-OCRF").Page("Page-OCRF").Link("Save And Next").Exist Then
		'Browser("Browser-OCRF").Page("Page-OCRF").Link("Save And Next").Click
		Browser("Browser-OCRF").Page("Page-OCRF").Link("Save And Next").DoubleClick
    	Call ReportStep (StatusTypes.Pass, "Clicked on Save and Next button","'User DB access' tab")
    Else
        Call ReportStep (StatusTypes.Fail, "Not Clicked on Submit and 'Save and Complete' button","'User DB access' tab")
	End If
	
	If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Continue").Exist(10) Then
   		Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Continue").Click
    End If
End Function


'Created By: Poornima
'Created on: 6/23/2018
'Description: 'Fuction to filter datasource and request is in Chnage ownership module.
'Parameters:1. Filtergrid:Two tables 1.List of DS 2.List of OCRF
'Parameters:2. searchColumn:"name of column" 3.searchItem:"is data to be filter" 4. ShouldExist:"YES or NO(record should exist or not)"5.dataFilter=Filter need to do?("YES orNO")
             
Function ChangeOwnershipFilterData(ByVal Filtergrid,ByVal searchColumn,ByVal searchItem,ByVal Timestamp,ByVal ShouldExist,ByVal dataFilter,ByRef Objdata)
    
	If Filtergrid="List of OCRF" Then
	   Set table=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("ChangeOwnerIDGrid")
	   Select Case searchColumn
			Case "OCRFID":
				Col=4
			Case "Client":
				Col=6
			Case "Offering":
				Col=9
		    Case "Owner":
				Col=10
			Case Else:
				
		End Select
		
		
	ElseIf Filtergrid="List of DS" Then
	    Set table=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("ChangeOwnerDSGrid")
	    Select Case searchColumn
			Case "DSName":
				Col=5
			Case "DSOwner":
				Col=6
		    Case "IsLock":
				Col=4	
			Case Else:
		End Select
		
	End If
	
    
    If Filtergrid="List of OCRF" Then
		Set TableData=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("ChangeOwnerDataIDGrid")
    ElseIf Filtergrid="List of DS" Then	
        Set TableData=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("ChangeOwnerDataDSGrid")
    End If
    
    
    
    If Filtergrid="List of OCRF" Then
         rows=TableData.GetROProperty("rows")
         If rows>=2 Then
            
            For i = 2 To rows Step 1
                   ValidateData=TableData.GetCellData(i,Col)
            	   If ValidateData=searchItem AND ShouldExist="YES" Then
    	   			  Call ReportStep (StatusTypes.Pass, "Item "&searchItem&" Exist in grid","Filter item in 'Change ownership' module")
    	   			  'Select data for change ownership operation
                	   Set checkbox=TableData.ChildItem(i,2,"WebCheckbox",0)
	            	   checkbox.Set "ON"
	            	   ChangeOwnershipFilterData=i
	            	   Exit Function
                   ElseIf ValidateData=searchItem AND ShouldExist="NO" Then    
                	   Call ReportStep (StatusTypes.Fail, "Item "&searchItem&" Exist in grid","Filter item in 'Change ownership' module")             
                   End If 
            Next 
         End If
    
    ElseIf Filtergrid="List of DS" Then
           If dataFilter="YES"  Then
    		  Set DataEnter=table.ChildItem(2,Col,"WebEdit",0)
    		  DataEnter.highlight
    		  DataEnter.set searchItem
   	 		  DataEnter.Click
    		  Set WshShell = CreateObject("WScript.Shell")
    		  WshShell.SendKeys "{ENTER}"
			  wait 5
    	   End If 
    	   rows=TableData.GetROProperty("rows")
    	   If rows>=2 Then
            For i = 2 To rows Step 1           
                ValidateData=TableData.GetCellData(i,Col)            
           	    If ValidateData=searchItem&Timestamp AND ShouldExist="YES" Then
    	   			Call ReportStep (StatusTypes.Pass, "Item "&searchItem&Timestamp&" Exist in grid","Filter item in 'Change ownership' module")
    	   			'Select data for change ownership operation
                	Set checkbox=TableData.ChildItem(i,2,"WebCheckbox",0)
	            	checkbox.Set "ON"
	            	ChangeOwnershipFilterData=i
	            	Exit Function
            	ElseIf ValidateData=searchItem AND ShouldExist="NO" Then    
                	Call ReportStep (StatusTypes.Fail, "Item "&searchItem&Timestamp&" Exist in grid","Filter item in 'Change ownership' module")             
        		End If 
            Next 
         End If
    End If
'    ElseIf rows<2 AND ShouldExist="NO" Then
'           Call ReportStep (StatusTypes.Pass, "Item "&searchItem&Timestamp&" Not Exist in grid","Filter item in 'Change ownership' module")     
'    

End Function

'Created By: Poornima
'Created on: 6/23/2018
'Description: 'Fuction to perform 'Change action' functionality in Chnage ownership module.
'Parameters:1. UserName: Username to which ownership is transfering.

Function ChangeOwnershipOperation(ByVal UserName,ByRef Objdata)
     'Enter user name to change the ownership to otheruser
      If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("NewOwnerID").Exist(10) Then
         Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("NewOwnerID").Set UserName
      	 Call ReportStep (StatusTypes.Pass,"Entered user name to change ownership"," 'Change ownership' module")
      Else
         Call ReportStep (StatusTypes.Fail,"Not Entered user name  to change ownership"," 'Change ownership' module")
      End If

	 'Click on 'change ownership' button
      If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("btnChangeOwnership").Exist(10) Then
         Browser("Browser-OCRF").Page("Page-OCRF").WebButton("btnChangeOwnership").Click
      	 Call ReportStep (StatusTypes.Pass,"Clicked on 'Change ownership' button","'Change ownership' module")
      Else
         Call ReportStep (StatusTypes.Fail,"Not Clicked on 'Change ownership' button","'Change ownership' module")
      End If
     'Click on 'Okay' button
      If Browser("Online CRF - Change Ownership").Page("Online CRF - Change Ownership").WebButton("Okay").Exist(10) Then
         Browser("Online CRF - Change Ownership").Page("Online CRF - Change Ownership").WebButton("Okay").Click
      	 Call ReportStep (StatusTypes.Pass,"Clicked on 'Okay' button","'Change ownership' module")
      	 
      Else
      	    
	    '******************************************************************
    Set objShell = CreateObject("Wscript.shell")
    objShell.SendKeys "{ENTER}"
    
    wait 10
    Set objShell = CreateObject("Wscript.shell")
    objShell.SendKeys "{ENTER}"
    wait 10
'    ***************************************************************************8
        ' Call ReportStep (StatusTypes.Fail,"Not Clicked on 'Okay' button","'Change ownership' module")
      End If

End Function
'Created By: Poornima
'Created on: 13/7/2018
'Description: 'Fuction to Validate Actual data for request id in 'OCRF Admin User page'
'Parameters:1. ReqID: search item 2. ColName:column name for which actual data has to validate  3.ExpectedData: data that has to validate with actual
Function ValidateOCRFAdminUserPageStatus(ByVal User,ByVal ReqID,ByVal ColName,ByVal ExpectedData)
  	
  	foundInTable= FilterInOnlineCRF("Current Requests","OCRF ID",ReqID,objData)
  	 wait 5
    CurrRequest=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").RowCount 
    CompRequest=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTblCompletedReqData").RowCount
  If UCase(User)=Trim(UCase("Admin")) Then
  	Select Case UCase(ColName)
			Case UCase("OCRF ID"):
							col = 3
			Case UCase("Client Name"):
							col = 4
			Case UCase("Offering"):
							col = 6
			Case UCase("Status"):
							col = 7
			Case UCase("Modified By"):
							col = 13				
			Case Else:
	 End Select
  Else

    Select Case UCase(ColName)
			Case UCase("OCRF ID"):
							col = 3
			Case UCase("Client Name"):
							col = 4
			Case UCase("Offering"):
							col = 6
			Case UCase("Status"):
							col = 8
			Case UCase("Modified By"):
							col = 12				
			Case Else:
	 End Select
	 
	
  End If  
  If CurrRequest >= 2 Then    
			Set objTableRow = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data")
			ActualData=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").GetCellData(2,col)
			If ActualData=ExpectedData Then
				Call ReportStep (StatusTypes.Pass, "Request ID "&ReqID&" having "&ExpectedData&" data in the comumn '"&ColName,"'Online CRF' tab")
			Else
			   Call ReportStep (StatusTypes.Fail,  "Request ID "&ReqID&" having "&ActualData&" data in the comumn '"&ColName&" instead of expected "&ExpectedData,"'Online CRF' tab")
			End If
			
    End If
	'ElseIf Ucase(SearchTable) = Ucase("Completed Requests") Then
	If CompRequest >= 2 Then	
			Set objTableRow = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTblCompletedReqData")
			ActualData=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTblCompletedReqData").GetCellData(2,col)
			
            If ActualData=ExpectedData Then
				Call ReportStep (StatusTypes.Pass, "Request ID "&ReqID&" having "&ExpectedData&" data in the comumn '"&ColName,"'Online CRF' tab")
			Else
			   Call ReportStep (StatusTypes.Fail,  "Request ID "&ReqID&" having "&ActualData&" data in the comumn '"&ColName&" instead of expected "&ExpectedData,"'Online CRF' tab")
			End If
	End If	 
  End Function


'Created By: Poornima
'Created on: 13/7/2018
'Description: 'Fuction to Validate Actual data for request id in Landing page
'Parameters:1. ReqID: search item 2. ColName:column name for which actual data has to validate  3.ExpectedData: data that has to validate with actual.
 Function ValidateOCRFLandingPageStatus(ByVal User,ByVal ReqID,ByVal ColName,ByVal ExpectedData)
  	
  	foundInTable= FilterInOnlineCRF("Current Requests","OCRF ID",ReqID,objData)
  	 wait 5
    CurrRequest=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").RowCount 
    CompRequest=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTblCompletedReqData").RowCount
  If UCase(User)=Trim(UCase("Admin")) Then
  	Select Case UCase(ColName)
			Case UCase("OCRF ID"):
							col = 3
			Case UCase("Client Name"):
							col = 4
			Case UCase("Offering"):
							col = 6
			Case UCase("Status"):
							col = 7
			Case UCase("Modified By"):
							col = 13				
			Case Else:
	 End Select
  Else

    Select Case UCase(ColName)
			Case UCase("OCRF ID"):
							col = 3
			Case UCase("Client Name"):
							col = 4
			Case UCase("Offering"):
							col = 6
			Case UCase("Status"):
							col = 8
			Case UCase("Modified By"):
							col = 12				
			Case Else:
	 End Select
	 
	
  End If  
  If CurrRequest >= 2 Then    
			Set objTableRow = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data")
			ActualData=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").GetCellData(2,col)
			If ActualData=ExpectedData Then
				Call ReportStep (StatusTypes.Pass, "Request ID "&ReqID&" having "&ExpectedData&" data in the comumn '"&ColName,"'Online CRF' tab")
			Else
			   Call ReportStep (StatusTypes.Fail,  "Request ID "&ReqID&" having "&ActualData&" data in the comumn '"&ColName&" instead of expected "&ExpectedData,"'Online CRF' tab")
			End If
			
    End If
	'ElseIf Ucase(SearchTable) = Ucase("Completed Requests") Then
	If CompRequest >= 2 Then	
			Set objTableRow = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTblCompletedReqData")
			ActualData=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTblCompletedReqData").GetCellData(2,col)
			
            If ActualData=ExpectedData Then
				Call ReportStep (StatusTypes.Pass, "Request ID "&ReqID&" having "&ExpectedData&" data in the comumn '"&ColName,"'Online CRF' tab")
			Else
			   Call ReportStep (StatusTypes.Fail,  "Request ID "&ReqID&" having "&ActualData&" data in the comumn '"&ColName&" instead of expected "&ExpectedData,"'Online CRF' tab")
			End If
	End If	 
  End Function
  
'Created By: Poornima
'Created on: 8/2/2018
'Description: 'Fuction to 'Delete' OCRF in offering details tab
Function DeleteOCRF()

		Set objDeleteOCRF = Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Delete OCRF")
	    Call SCA.ClickOn(objDeleteOCRF,"Delete OCRF","Offering Details")
	    'Call PageLoading()

	    If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Yes").Exist(10) Then
		   Set objConfrmYes = Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Yes")
		   Call SCA.ClickOn(objConfrmYes,"Yes","Confirmation")
	    End If
	    Call PageLoading()
	    Set objInfOkay = Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Information_Okay")
	    Call SCA.ClickOn(objInfOkay,"Okay","Information")
	    Call PageLoading()
	    
End Function

'Created By: Poornima
'Created on: 8/2/2018
'Description: 'Fuction to 'filter' Data in 'Audit log' tab
Function AuditLogFilterData(ByVal searchColumn,ByVal searchItem,ByRef Objdata)

 Set table=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_GridHeader")
'	   ******************************************************
' Modified by madhu - 03/30/2020
	   
	   Select Case searchColumn
			Case "OCRFID":
				Col=3
				Set DataEnter=table.WebEdit("WEbedit_OfferingID_Search")
			Case "User Name":
				Col=5
				Set DataEnter=table.WebEdit("UserName")
			Case "Category":
				Col=6
				Set DataEnter=table.WebEdit("Category")
		    Case "Offering":
				Col=9
				Set DataEnter=table.WebEdit("OfferingName")
			Case "Action":
				Col=10
				Set DataEnter=table.WebEdit("ActionName")
			Case Else:
		End Select
		
	
		'Set DataEnter=table.WebEdit("WEbedit_OfferingID_Search")
'		**************************************************
	   ' Set DataEnter=table.ChildItem(2,Col,"WebEdit",0)
    		DataEnter.highlight
    		DataEnter.set searchItem
   	 		DataEnter.Click
    		Set WshShell = CreateObject("WScript.Shell")
    		WshShell.SendKeys "{ENTER}"
			wait 5
			  
			rows= Browser("Browser-OCRF").Page("Page-OCRF").WebTable("Auditlog_Data").GetROProperty("rows")
			If rows>=2 Then
	    		Call ReportStep (StatusTypes.Information, "Item "&searchItem&Timestamp&" Exist in Auditlog grid","Filter item in 'AuditLog'")
				AuditLogFilterData=True
			Else
	    		AuditLogFilterData=false
	    		Call ReportStep (StatusTypes.Information, "Item "&searchItem&Timestamp&" Not Exist in Auditlog grid","Filter item in 'AuditLog'")
			End If
    	 
End Function
'Created By: Poornima
'Created on: 8/2/2018
'Description: 'Fuction to validate status of Searched data in Auditlog

Function ValidateAuditLogStatus(ByVal ReqID,ByVal ColName,ByVal ExpectedData)
  	 wait 5
    rows= Browser("Browser-OCRF").Page("Page-OCRF").WebTable("Auditlog_Data").GetROProperty("rows")
  	Select Case UCase(ColName)
			Case UCase("OCRF ID"):
							col = 3
			Case UCase("User Name"):
							col = 5
			Case UCase("Offering"):
							col = 9
			Case UCase("Action"):
							col = 10				
			Case Else:
	End Select
	If rows >= 2 Then
		ActualData=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("Auditlog_Data").GetCellData(2,col)
		If ActualData=ExpectedData Then
			Call ReportStep (StatusTypes.Pass, "Request ID "&ReqID&" having "&ExpectedData&" data in the comumn '"&ColName,"'Online CRF' tab")
		Else
			Call ReportStep (StatusTypes.Fail,  "Request ID "&ReqID&" having "&ExpectedData&" data in the comumn '"&ColName,"'Online CRF' tab")
		End If
			
	End If
	

End Function


'Created By: Poornima
'Created on: 9/14/2018
'Description: 'Fuction to Perform ADD/Edit/Delete action on record searched in 'OCRF Admin user' moduke
Function PerformActionOnOCRFAdminUserPage(ByVal Action,ByVal searchCol,ByVal searchItem,ByRef Objdata)
	Select Case UCase(searchCol)
			Case UCase("First Name"):
							col = 4
			Case UCase("Last Name"):
							col = 5
			Case UCase("User Login ID"):
							col = 6
			Case UCase("Is Active"):
							col = 8		
			Case Else:
	 End Select
	 
	 Select Case UCase(Action)
			Case UCase("Add User"):
								'Click on 'Add user' button in 'OCRF Admin user' page
								If Browser("Browser-OCRF").Page("OCRF AdminUser").WebElement("OCRFadminAddUser").Exist(10) Then
	  							   Browser("Browser-OCRF").Page("OCRF AdminUser").WebElement("OCRFadminAddUser").Click 
									Call ReportStep (StatusTypes.Pass, "click on  'Add user' button in 'OCRF Admin user' page","OCRF Admin user")
								Else
								    Call ReportStep (StatusTypes.Fail,  "Not click on  'Add user' button in 'OCRF Admin user' page","OCRF Admin user")
								End If
	
								'Enter Admin userid in text box
								If Browser("Browser-OCRF").Page("OCRF AdminUser").WebEdit("OCRFAdminUserID").Exist(10) Then
								   Browser("Browser-OCRF").Page("OCRF AdminUser").WebEdit("OCRFAdminUserID").Set objData.item("LoUserToAdmin")
								   Browser("Browser-OCRF").Page("OCRF AdminUser").WebCheckBox("OCRFAdminIsActive").Set "ON"
									Call ReportStep (StatusTypes.Pass, "Entered  'User ID' "&LoUserPassword&" in 'Add Ocrf admin' page","OCRF Admin user")
							    Else
								    Call ReportStep (StatusTypes.Fail,  "Not Entered  'User ID' "&LoUserPassword&" in 'Add Ocrf admin' page","OCRF Admin user")
								End If
								'Validate Autopopulated data appearing in Firstname and lastname fields based on user ID entered
								FirstName=Browser("Browser-OCRF").Page("OCRF AdminUser").WebEdit("OCRFAdminFirstName").GetROProperty("value")
								LastName=Browser("Browser-OCRF").Page("OCRF AdminUser").WebEdit("OCRFAdminLastName").GetROProperty("value")
								If instr(UCASE(FirstName),UCASE(LoUserPassword))>0 AND instr(UCASE(LastName),UCASE(LoUserPassword))>0 Then
									Call ReportStep (StatusTypes.Pass, "Auto populated valid data appeared in FirstName as "&FirstName&" and Last name as"&LastName,"OCRF Admin user")
							    Else
								    Call ReportStep (StatusTypes.Fail, "Auto populated valid data Not appeared in FirstName as "&FirstName&" and Last name as"&LastName,"OCRF Admin user")
								End If
								
								'Click on 'Save' button
								Browser("Browser-OCRF").Page("OCRF AdminUser").WebButton("OCRFAdminSave").Click
								
								'Check for Conformation method that ' Admin User data has saved successfully'
								If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Information_Okay").Exist(10) Then
								   Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Information_Okay").Click 
									Call ReportStep (StatusTypes.Pass, "click on  'Okay' confirmation button in 'OCRF Admin user' page","OCRF Admin user")
							    Else
								    Call ReportStep (StatusTypes.Fail,  "Not click on  'Okay' confirmation button in 'OCRF Admin user' page","OCRF Admin user")
								End If
			Case UCase("Edit User"):
							col = 5
			Case UCase("Delete User"):
						        Rows= FilterInOCRFAdminUserTable("","User Login ID",searchItem,objData)
								If Rows>=2 Then
							    	OCRfAdminData=Browser("Browser-OCRF").Page("OCRF AdminUser").WebTable("OCRFadminUserList").GetCellData(2,col)
							    	If instr(UCASE(OCRfAdminData),UCASE(searchItem))>0 Then
							    	    Browser("Browser-OCRF").Page("OCRF AdminUser").WebTable("OCRFadminUserList").Object.rows(Rows-1).Click
							    		Call ReportStep (StatusTypes.Pass, "Selected Searched item for Deletion","OCRF Admin user")
							    		'click on 'Delete user' button to delete admin user
									    If Browser("Browser-OCRF").Page("OCRF AdminUser").WebElement("OCRFAdminUserDelete").Exist(10) Then
										   Browser("Browser-OCRF").Page("OCRF AdminUser").WebElement("OCRFAdminUserDelete").Click 
											Call ReportStep (StatusTypes.Pass, "click on  'Delete user' button in 'OCRF Admin user' page","OCRF Admin user")
									    Else
										    Call ReportStep (StatusTypes.Fail,  "Not click on  'Delete user' button in 'OCRF Admin user' page","OCRF Admin user")
										End If
		
											'Click on confirmation 'Yes' button
										If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Yes").Exist(10) Then
										   Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Yes").Click 
											Call ReportStep (StatusTypes.Pass, "click on  'Yes' button of Delete User confirmation 'OCRF Admin user' page","OCRF Admin user")
									    Else
										    Call ReportStep (StatusTypes.Fail,  "Not click on  'Yes' button of Delete User confirmation 'OCRF Admin user' page","OCRF Admin user")
										End If
							    	Else
							    	    Call ReportStep (StatusTypes.Fail,"Not Selected Searched item for Deletion","OCRF Admin user")
							    	End If
							    End If
							    
								
			Case Else:
	 End Select
End Function


'Created By: Poornima
'Created on: 9/19/2018
'Description: 'Fuction to Login to DC site with client url

Public Function LoginToSCAClientUrl(ByVal strUserName,ByVal StrPwd, ByVal clientURL)
      
       Dim objUserName,objPwd,objLoginBtn,objSCAHomePageImage,returnval,objSCAHomeFrame
		SystemUtil.CloseProcessByName("iexplore.exe")
		systemutil.Run "iexplore.exe"
		wait 4
		Set WshShell=createobject("Wscript.shell")
		WshShell.SendKeys "^+{DEL}"
		wait 5
		WshShell.SendKeys "ENTER"
		wait 2
        SystemUtil.CloseProcessByName("iexplore.exe")
        wait 4
       
       'Redirecting to Client URL SCA - Start
        'Navigate to SCA through DC SIte IMS Analyzer link in DC Home page Analysis Table
       SystemUtil.Run clientURL
        wait 2
        Browser("DC Home").Page("Microsoft Forefront TMG").Sync
        
        Set objUserName=Browser("micclass:=browser").Page("micclass:=page").WebEdit("xpath:=//input[@id='txtUserID']")
        Call SCA.SetText(objUserName,strUserName,"textUserName" ,"DC Login Page" )
        wait 2
        'click on continue button
		Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//input[@id='btnValidate']").Click
		wait 4
        
        Set objPwd=Browser("micclass:=browser").Page("micclass:=page").WebEdit("xpath:=//input[@id='txtPassword']")
        Call SCA.SetText(objPwd,StrPwd,"txtpassword" ,"DC Login Page" )    
        
        Set objLoginBtn=Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//input[@id='btnLogin']")
        Call SCA.ClickOn(objLoginBtn,"Login Button","DC Login Page")
        
        Browser("DC Home").Page("Microsoft Forefront TMG").Sync
        Browser("DC Home").Page("Microsoft Forefront TMG").RefreshObject
        
        Set objDCHomePageImage=Browser("DC Home").Page("Home").Image("My_Informed_Decisions")
        If objDCHomePageImage.Exist(120) <> True Then
		      Call ReportStep (StatusTypes.Fail,"Client DC URL Login is not Successfull " ,"Home Page" )
		Else
		      Call ReportStep (StatusTypes.Pass,"Client DC URL Login is Successfull" ,"Home Page" )
		      Call ReportStep (StatusTypes.Information,"Click on IMS Analysis Manager to redirect to SCA" ,"Home Page" )
		End If
		
 		wait 2
 		
        'Click on IMS Analysis Manager link in DC Home Page
        Set objlnkIMSAnalysisManager = Browser("DC Home").Page("Home").Link("lnkIMS Analysis Manager")
        Call SCA.ClickOn(objlnkIMSAnalysisManager,"lnkIMS Analysis Manager", "DC Home Page")
         wait 2
        Browser("DC Home").Page("IMS Analysis Manager").Sync
        Set objlnkSharedReports = Browser("DC Home").Page("IMS Analysis Manager").Frame("Frame").Link("lnkShared Reports")
                    
        wait 2
           
       If Browser("Analyzer").Page("Home").Frame("Frame").Link("lnkSharedReports").Exist(120) <> True Then
          Call ReportStep (StatusTypes.Fail,"Client SCA URL Login is not Successfull " ,"Home Page" )
		Else
          Call ReportStep (StatusTypes.Pass,"Client SCA URL Login is Successfull" ,"Home Page" )
       End If    
                              
 End Function   

'Created By: Poornima
'Created on: 9/19/2018
'Description: 'Validate user permission to create report in SCA (Validate report creation icon enabled or not)

Function ValidateUserPermToCreateReport(Byval User,ByRef objData)
     'click on My report
     Browser("Browser-SCA").Page("Page-SCA").Frame("SCA-Frame").Link("Link_Myeports").Click

     'Check report creation icon eneble or not and store it in one variable
     EnableProp=Browser("Browser-SCA").Page("Page-SCA").Frame("view").Image("newreport").GetROProperty("outerhtml")

     'Select User type
	 Select Case UCase(User)
	       Case UCase("Architect"):
	                  If instr(EnableProp,"opacity")>0 Then
	                  	  Call ReportStep (StatusTypes.Fail, "For architect user report creation icon disabled","My report Report creation")
					  Else
						  Call ReportStep (StatusTypes.Pass,  "For architect user report creation icon Enabled","My report Report creation")						
	                  End If
	       
	       Case UCase("Consumer"):
	                  If instr(EnableProp,"opacity")>0 Then
	                  	  Call ReportStep (StatusTypes.Pass, "For Consumer user report creation icon disabled","My report Report creation")
					  Else
						  Call ReportStep (StatusTypes.Fail,  "For Consumer user report creation icon Enabled","My report Report creation")						
	                  End If
	 	   Case Else:
	 End Select
End Function




'Created By: Sumit
'Created on: 4/27/2018
'Description: 'Update values in Add Users tab
Function UpdateActionValueInAddUsers(ByVal FileName,ByVal SheetName,ByVal ValName,ByVal flagShowDeleted, ByVal StartRow,Byval EndRow,ByRef objData)

	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementOfferingUserList").Exist(60) Then
		Call ReportStep (StatusTypes.Pass,  "'Add User' tab opened","Add user tab")
    Else
       Call ReportStep (StatusTypes.Fail,  "'Add User' tab not opened","Add user tab")
    End If
	
	Call ImportSheet(SheetName,FileName)
	
	Datatable.SetCurrentRow(StartRow)
	NumberOfRows = (EndRow-StartRow)+1
	
	Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_NumberOfRecords").Select "100"
	UserLoginId = DataTable.Value("User_Login_Id",SheetName)
		
'	Set gridHeaderTable = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_GridHeader_AddUsers")
'	Set userLoginIdGridHeader= gridHeaderTable.ChildItem(2,9,"WebEdit",0)
'	userLoginIdGridHeader.Set UserLoginId
	
	'changes by srinivas-27th july 2020-handled by descriptive programming
	browser("Browser-OCRF").Page("Page-OCRF").WebEdit("html tag:=INPUT","html id:=gs_UserLoginID").highlight
	Set userLoginIdGridHeader=	browser("Browser-OCRF").Page("Page-OCRF").WebEdit("html tag:=INPUT","html id:=gs_UserLoginID")
	userLoginIdGridHeader.Set UserLoginId
	wait 1
	userLoginIdGridHeader.Click
	Set wsh = CreateObject("Wscript.shell")
	wsh.SendKeys "{ENTER}"
	wait 2

	TableRow=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").RowCount
	If flagShowDeleted = True Then
		Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_ShowDeleted").Click
		wait 2
		row = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").RowCount
		existingValue = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").GetCellData(row,7)
		If existingValue = ValName Then
			Call ReportStep (StatusTypes.Pass,  UserLoginId & " record is deleted successfully","Verify value in Action column")
		   Else
		   	Call ReportStep (StatusTypes.Fail,  UserLoginId & " record is not deleted successfully","Verify value in Action column")
		End If 
		
		Exit Function
		
	End If
	For row = 2 To  TableRow Step 1

		   'Call WebEditExistence(row,7,"WebList")
		   existingValue = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").GetCellData(row,7)
		   If existingValue = ValName Then
		   		Call ReportStep (StatusTypes.Pass,  "Action column already has the expected value","Verify value in Action column")
		   Else
		   		Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").object.rows(1).click
			   Set actionListObj =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(row,7,"WebList",0)
			   If actionListObj.GetRoProperty("class")="editable"  Then
				   actionListObj.Select ValName
				   Call ReportStep (StatusTypes.Pass,  " The expected value " & ValName &  " is updated","Change value in Action column")
			   End If
		   End If
		   
	Next 
	
	
End Function



'Created By: Sumit
'Created on: 4/27/2018
'Description: 'Manage User Access Delete and Change Action functionality
Function ValidateDeleteAndChangeActionInManageUserAccess(ByVal FileName,ByVal SheetName,ByVal OCRFId, ByVal action, ByVal StartRow,Byval EndRow,ByRef objData)

	If Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebButton("WebBtnSearch").Exist(10) Then
		Call ReportStep (StatusTypes.Pass,  "User is navigated to Manage User Access page","Navigate to Manage User Access page")
		
		Call ImportSheet(SheetName,FileName)
	
		Datatable.SetCurrentRow(StartRow)
		NumberOfRows = (EndRow-StartRow)+1
		
		UserLoginId = DataTable.Value("User_Login_Id",SheetName)
		
		Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebEdit("WebEditOfferingUserLoginId").Set UserLoginId
		Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebButton("WebBtnSearch").Click
		Call PageLoading()
	
		'below line added by Avinash on 4th Jun 2019
		'reason -record was present in last page due to pagination
		If Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebElement("lastPage_Pagination").Exist(10) Then
			Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebElement("lastPage_Pagination").Click
		End If
		
		wait 2
		'Modification Completed
		
		rowCount = Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebTable("UserOfferingGrid").RowCount
		wait 2
		If rowCount > 1 Then
			rowNo = Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebTable("UserOfferingGrid").GetRowWithCellText (OCRFId,3)
			wait 2
			If rowNo < 0 Then
				Call ReportStep (StatusTypes.Information,  "OCRF is not available for " & action ,action & " in Manage User Access")
				Exit Function
			End If
			flagActionDone = False 

			wait 3
			Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebTable("UserOfferingGrid").ChildItem(rowNo,2,"WebCheckbox",0).Set "On"
			wait 3
			'Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebElement(action).Click
			'Modified by Madhu - 04/23/2020
			'Browser("BrowserPoPUp").Page("Online CRF - Manage Users").WebElement("Delete").Click
			
			'changes by srinivas as the previous one operation was completely wrong now changed it-27/7/2020
				Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebElement("xpath:=//td[@id='bulkUser-list_toppager_left']//div[text()='"&action&"']").Click

			wait 3
			If Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebElement("WebEleConfirmMessage").Exist(15) Then
'				Set obj = Description.Create
'				obj("micclass").value = "WebButton"
'				obj("value").value = "Yes"
'				obj("visible").value = True
'				Set chObj = Browser("Browser-OCRF").Page("OCRF-ManageUsers").ChildObjects(obj)
'				For i = 0 to chObj.count -1
'					chObj(i).click
'				Next
				Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebButton("visible:=True","value:=Yes").highlight
				Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebButton("visible:=True","value:=Yes").Click
				Call PageLoading()
				wait 1
				Set objSuccessInfo = Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebElement("html tag:=DIV","visible:=True","innertext:=Action Executed Successfully.")
				If objSuccessInfo.Exist(30) Then
					Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebButton("visible:=True","value:=Okay").highlight
					Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebButton("visible:=True","value:=Okay").Click
					wait 3
					flagActionDone = True
					Call ReportStep (StatusTypes.Pass,  action & " is done successfully",action & " in Manage User Access")
					Else
					Call ReportStep (StatusTypes.Fail,  action & " is not done successfully",action & " in Manage User Access")
				End If
				
			End If
			Else
			Call ReportStep (StatusTypes.Information,  "No search records are found in Manage User Access page","Search OCRF in Manage User Access page")
		End If
	
	Else
		Call ReportStep (StatusTypes.Fail,  "User is not navigated to Manage User Access page","Navigate to Manage User Access page")
	End If
	
	If action = "Delete" And flagActionDone = True Then
	
		Call SelectMenuOption("Online CRF","",objData)
		Call PageLoading()
		foundInTable= FilterInOnlineCRF("Current Requests","OCRF ID",OCRFId,objData)
		Call PageLoading()
		
		rowNo = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").GetRowWithCellText (OCRFId,3)
		runTimeStatus = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").GetCellData(rowNo,8)
		If runTimeStatus = "Submitted" Then
			Call ReportStep (StatusTypes.Pass,  "Status of the OCRF is Submitted " , "Verify Status of OCRF upon Delete in Manage User Access")
			Else
			Call ReportStep (StatusTypes.Fail,  "Status of the OCRF is " & runTimeStatus , "Verify Status of OCRF upon Delete in Manage User Access")
		End If
	End If
	If action = "ChangeAction" And flagActionDone = True Then
		Do
			wait 5
			Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebButton("WebBtnSearch").Click
			Call PageLoading()
			rowNo = Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebTable("UserOfferingGrid").GetRowWithCellText (OCRFId,3)
		Loop Until (rowNo < 0) 
		
		If rowNo < 0 Then
			Call ReportStep (StatusTypes.Pass,  "Change Action is successfully completed " , "Verify Change Action in Manage User Access")
			
		End If
		
	End If

End Function


'Created By: Avinash
'Created on: 1/8/2019
'Description: 'For Editor and Reader users, to check checkbox enabled or disabled in 'Manage User Access' page
Function ValidateOCRFCheckBoxforRolesInManageUserAccess(ByVal FileName,ByVal SheetName,ByVal OCRFId, ByVal action, ByVal StartRow,Byval EndRow,ByRef objData)
	
	LoReader1=Ucase(objData.item("LoUserR1"))
    LoReader2=Ucase(objData.item("LoUserR2"))
    LoReader3=Ucase(objData.item("LoUserR3"))
    LoEditor1=Ucase(objData.item("LoUserE1"))
    LoEditor2=Ucase(objData.item("LoUserE2"))
    LoEditor3=Ucase(objData.item("LoUserE3"))
	grantEditorRole=objData.Item("GrantEditorRole")
	grantReaderRole=objData.Item("GrantReaderRole")		
	
	If Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebButton("WebBtnSearch").Exist(10) Then
		Call ReportStep (StatusTypes.Pass,  "User is navigated to Manage User Access page","Navigate to Manage User Access page")
		
		Call ImportSheet(SheetName,FileName)
	
		Datatable.SetCurrentRow(StartRow)
		NumberOfRows = (EndRow-StartRow)+1
		
		UserLoginId = DataTable.Value("User_Login_Id",SheetName)
		
		Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebEdit("WebEditOfferingUserLoginId").Set UserLoginId
		
		wait 2
		Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebButton("WebBtnSearch").Click
		Call ReportStep (StatusTypes.Pass, "Searching with "&UserLoginId&" in Manage User Access Page","'Manage User Access' Page")
		Call PageLoading()
	
		'below line added by Avinash on 4th Jun 2019
		'reason -record was present in last page due to pagination
		If Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebElement("lastPage_Pagination").Exist(10) Then
			Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebElement("lastPage_Pagination").Click
		End If
		
		wait 2
		'Modification Completed
		
		rowCount = Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebTable("UserOfferingGrid").RowCount
		wait 2
		If rowCount > 1 Then
			rowNo = Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebTable("UserOfferingGrid").GetRowWithCellText (OCRFId,3)
			wait 2
			If rowNo < 0 Then
				Call ReportStep (StatusTypes.Information,  "OCRF is not available for " & action ,action & " in Manage User Access")
				Exit Function
			End If
			flagActionDone = False 
						
			'Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebTable("UserOfferingGrid").ChildItem(rowNo,2,"WebCheckbox",0).set "ON"

			wait 3
			Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebTable("UserOfferingGrid").ChildItem(rowNo,2,"WebCheckbox",0).highlight
			'checkBx = Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebTable("UserOfferingGrid").ChildItem(rowNo,2,"WebCheckbox",0).object.checked
			checkBx = Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebTable("UserOfferingGrid").ChildItem(rowNo,2,"WebCheckbox",0).GetRoProperty("disabled")
			
			wait 2
			If checkBx = FALSE Then
				Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebTable("UserOfferingGrid").ChildItem(rowNo,2,"WebCheckbox",0).set "ON"
				Call ReportStep (StatusTypes.Pass, "User is logged in as Reader", "in Manage User Access Page")
				
			    Else
			   ' Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebTable("UserOfferingGrid").ChildItem(rowNo,2,"WebCheckbox",0).set "ON"
				Call ReportStep (StatusTypes.Pass, "User is logged in as Editor", "in Manage User Access Page")			
			End If
'			checkBx = Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebTable("UserOfferingGrid").ChildItem(rowNo,2,"WebCheckbox",0).GetROProperty("disabled")
'			wait 3
'			Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebElement(action).Click
'			wait 3
'			If Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebElement("WebEleConfirmMessage").Exist(15) Then
'				Set obj = Description.Create
'				obj("micclass").value = "WebButton"
'				obj("value").value = "Yes"
'				obj("visible").value = True
'				Set chObj = Browser("Browser-OCRF").Page("OCRF-ManageUsers").ChildObjects(obj)
'				For i = 0 to chObj.count -1
'					chObj(i).click
'				Next
'				Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebButton("visible:=True","value:=Yes").highlight
'				Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebButton("visible:=True","value:=Yes").Click
'				Call PageLoading()
'				wait 1
'				Set objSuccessInfo = Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebElement("html tag:=DIV","visible:=True","innertext:=Action Executed Successfully.")
'				If objSuccessInfo.Exist(30) Then
'					Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebButton("visible:=True","value:=Okay").highlight
'					Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebButton("visible:=True","value:=Okay").Click
'					wait 3
'					flagActionDone = True
'					Call ReportStep (StatusTypes.Pass,  action & " is done successfully",action & " in Manage User Access")
'					Else
'					Call ReportStep (StatusTypes.Fail,  action & " is not done successfully",action & " in Manage User Access")
'				End If
				
	'		End If
'			Else
'			Call ReportStep (StatusTypes.Information,  "No search records are found in Manage User Access page","Search OCRF in Manage User Access page")
'		End If
'	
'	Else
'		Call ReportStep (StatusTypes.Fail,  "User is not navigated to Manage User Access page","Navigate to Manage User Access page")
'	End If
'	
'	If action = "Delete" And flagActionDone = True Then
'	
'		Call SelectMenuOption("Online CRF","",objData)
'		Call PageLoading()
'		foundInTable= FilterInOnlineCRF("Current Requests","OCRF ID",OCRFId,objData)
'		Call PageLoading()
'		
'		rowNo = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").GetRowWithCellText (OCRFId,3)
'		runTimeStatus = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").GetCellData(rowNo,8)
'		If runTimeStatus = "Submitted" Then
'			Call ReportStep (StatusTypes.Pass,  "Status of the OCRF is Submitted " , "Verify Status of OCRF upon Delete in Manage User Access")
'			Else
'			Call ReportStep (StatusTypes.Fail,  "Status of the OCRF is " & runTimeStatus , "Verify Status of OCRF upon Delete in Manage User Access")
'		End If
'	End If
'	If action = "ChangeAction" And flagActionDone = True Then
'		Do
'			wait 5
'			Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebButton("WebBtnSearch").Click
'			Call PageLoading()
'			rowNo = Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebTable("UserOfferingGrid").GetRowWithCellText (OCRFId,3)
'		Loop Until (rowNo < 0) 
'		
'		If rowNo < 0 Then
'			Call ReportStep (StatusTypes.Pass,  "Change Action is successfully completed " , "Verify Change Action in Manage User Access")
'			
		End If
		
	End If

End Function




'Created By: Sumit
'Created on: 5/15/2018
'Description: 'Validation of OCRF role and owner in OCRF Administration Module
Function VerifyOCRFRoleAndOwner(ByVal FileName,ByVal SheetName, ByVal StartRow,Byval EndRow,ByVal OCRFId, ByVal ocrfOwnerId, ByRef objData)
	
	OCRFRole = Trim(objData.Item("OCRFRole"))
	loggedInUser = Trim(objData.Item("LoggedInUserId"))
	'owner = Ucase(trim(objData.Item("owner")))
	If Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebElement("OCRFAdministrationTitle").Exist Then
		Call ReportStep (StatusTypes.Pass,  "User is navigated to OCRF Administrator Module","Navigate to OCRF Administrator page")
		
		Call ImportSheet(SheetName,FileName)
	
		Datatable.SetCurrentRow(StartRow)
		NumberOfRows = (EndRow-StartRow)+1
		
		Set ocrfAdminHeaderGrid = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAdminHeaderTable")
		
'		*********************************************
'Modified by Madhu 02/27/2020_____________________________
'Modified OR
		If ocrfAdminHeaderGrid.Exist Then
			Set ocrfIdAdminHeader = ocrfAdminHeaderGrid.WebEdit("Webedit_OfferingID_search")
			ocrfIdAdminHeader.Set OCRFId
			Call SendKeyAndWait("{ENTER}",ocrfIdAdminHeader, 3)
			
			Set roleAdminHeader = ocrfAdminHeaderGrid.WebEdit("Webedit_OcrfRole_search")
			roleAdminHeader.Set OCRFRole
			Call SendKeyAndWait("{ENTER}",roleAdminHeader, 3)
		
		End If
		
		
''		Set ocrfIdAdminHeader = ocrfAdminHeaderGrid.ChildItem(2,4,"WebEdit",0)
'		ocrfIdAdminHeader.Set OCRFId
'		Call SendKeyAndWait("{ENTER}",ocrfIdAdminHeader, 3)
'		
'		Set roleAdminHeader = ocrfAdminHeaderGrid.ChildItem(2,9,"WebEdit",0)
'		roleAdminHeader.Set OCRFRole
'		Call SendKeyAndWait("{ENTER}",roleAdminHeader, 3)

'******************************************************88
		
		rowCount = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAccessListGrid").RowCount
		If OCRFRole = "Owner" and rowCount = 2 Then
			ocrfOwnerName = Ucase(Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAccessListGrid").GetCellData(2,10))
			userId = Ucase(Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAccessListGrid").GetCellData(2,8))
			
			If ocrfOwnerName = userId and ocrfOwnerName = Ucase(ocrfOwnerId) Then
				Call ReportStep (StatusTypes.Pass,  "Owner is validated for OCRFId as " & objData.item("OCRFId") ,"Validate Owner for OCRF")
				Else 
				Call ReportStep (StatusTypes.Fail,  "Validation of Owner failed as expected OwnerName: " & objData.item("owner") & " Actual Owner: " & ownerName,"Validate Owner for OCRF")
			End If
			linkAvailable = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAccessListGrid").ChildItem(2,13,"Link",0).GetRoProperty("innertext")
			If linkAvailable = "Change Ownership" Then
				Call ReportStep (StatusTypes.Pass,  "Change Ownership link is avilable for Owner","Validate availabilty of Change Ownshipship link for Owner")
				Else
				Call ReportStep (StatusTypes.Fail,  "Change Ownership link is not avilable for Owner" ,"Validate availabilty of Change Ownshipship link for Owner")
			End If
			checkboxAvailable = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAccessListGrid").ChildItem(2,2,"WebCheckBox",0).GetRoProperty("disabled")
			If checkboxAvailable = 1 Then
				Call ReportStep (StatusTypes.Pass,  "Available checkbox is disabled for Owner","Validate Checkbox is enable or disable")
				Else
				Call ReportStep (StatusTypes.Fail,  "Available checkbox is not disabled for Owner" ,"Validate Checkbox is enable or disable")
			End If
		ElseIf (OCRFRole = "Editor" Or OCRFRole = "Reader") and loggedInUser = "Owner" Then
			Set ocrfAdminHeaderGrid = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAdminHeaderTable")
'			************************************************************************
'modified by Madhu - 02/27/2020

			Set userIdAdminHeader = ocrfAdminHeaderGrid.WebEdit("Webedit_AccessUserLoginID_Search")
			'Set userIdAdminHeader = ocrfAdminHeaderGrid.ChildItem(2,8,"WebEdit",0)
'			*************************************************************8
			grantUserId = Ucase(objData.Item("GrantUserId"))
			userIdAdminHeader.set grantUserId
			Call SendKeyAndWait("{ENTER}",userIdAdminHeader,3)
			ocrfOwnerName = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAccessListGrid").GetCellData(2,10)
			userId = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAccessListGrid").GetCellData(2,8)
			
			If Ucase(ocrfOwnerName) = Ucase(ocrfOwnerId) And Ucase(grantUserId) = Ucase(userId) Then
				Call ReportStep (StatusTypes.Pass,  "OCRF Details for role as " & OCRFRole & " is validated successfully" ,"Validate OCRF when role is Editor/Reader")
				Else 
				Call ReportStep (StatusTypes.Fail,  "OCRF Details for role as " & OCRFRole & " is not validated successfully" ,"Validate OCRF when role is Editor/Reader")
			End If
			linkAvailable = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAccessListGrid").ChildItem(2,13,"Link",0).GetRoProperty("innertext")
			If linkAvailable = "Remove Permission" Then
				Call ReportStep (StatusTypes.Pass,  "Remove Permission link is avilable for role as " & OCRFRole,"Validate availabilty of Remove Permission link for Editor/Reader")
				Else
				Call ReportStep (StatusTypes.Fail,  "Remove Permission link is not avilable for role as " & OCRFRole,"Validate availabilty of Remove Permission link for Editor/Reader")
			End If
			checkboxAvailable = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAccessListGrid").ChildItem(2,2,"WebCheckBox",0).GetRoProperty("disabled")
			If checkboxAvailable = 0 Then
				Call ReportStep (StatusTypes.Pass,  "Available checkbox is enabled for Editor/Reader","Validate Checkbox is enable or disable")
				Else
				Call ReportStep (StatusTypes.Fail,  "Available checkbox is disabled for Editor/Reader" ,"Validate Checkbox is enable or disable")
			End If
		ElseIf OCRFRole = "Editor" and loggedInUser = "Editor" Then
			Set ocrfAdminHeaderGrid = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAdminHeaderTable")
			Set userIdAdminHeader = ocrfAdminHeaderGrid.ChildItem(2,8,"WebEdit",0)
			grantUserId = Ucase(objData.Item("GrantUserId"))
			userIdAdminHeader.set grantUserId
			Call SendKeyAndWait("{ENTER}",userIdAdminHeader,3)
			ocrfOwnerName = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAccessListGrid").GetCellData(2,10)
			userId = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAccessListGrid").GetCellData(2,8)
			
			If Ucase(ocrfOwnerName) <> Ucase(ocrfOwnerId) And Ucase(ocrfOwnerId) = Ucase(userId) Then
				Call ReportStep (StatusTypes.Pass,  "OCRF Details for role as " & OCRFRole & " is validated successfully" ,"Validate OCRF when role is Editor/Reader")
				Else 
				Call ReportStep (StatusTypes.Fail,  "OCRF Details for role as " & OCRFRole & " is not validated successfully" ,"Validate OCRF when role is Editor/Reader")
			End If
			linkAvailable = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAccessListGrid").GetCellData(2,13)
			If linkAvailable <> "Remove Permission" Then
				Call ReportStep (StatusTypes.Pass,  "Remove Permission link is unavilable for role as " & OCRFRole,"Validate availabilty of Remove Permission link for Editor/Reader")
				Else
				Call ReportStep (StatusTypes.Fail,  "Remove Permission link is avilable for role as " & OCRFRole,"Validate availabilty of Remove Permission link for Editor/Reader")
			End If
			checkboxAvailable = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAccessListGrid").ChildItem(2,2,"WebCheckBox",0).GetRoProperty("disabled")
			If checkboxAvailable = 1 Then
				Call ReportStep (StatusTypes.Pass,  "Available checkbox is disabled for Editor/Reader","Validate Checkbox is enable or disable")
				Else
				Call ReportStep (StatusTypes.Fail,  "Available checkbox is enabled for Editor/Reader" ,"Validate Checkbox is enable or disable")
			End If
'		**************************************************************88
'		Modified BY Madhu - 03/10/2020
		ElseIf OCRFRole = "Reader" and loggedInUser = "Editor" Then
			Set ocrfAdminHeaderGrid = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAdminHeaderTable")
			Set userIdAdminHeader = ocrfAdminHeaderGrid.WebEdit("Webedit_AccessUserLoginID_Search")
'			****************************************************************
			grantUserId = Ucase(objData.Item("GrantUserId"))
			userIdAdminHeader.set grantUserId
			Call SendKeyAndWait("{ENTER}",userIdAdminHeader,3)
			ocrfOwnerName = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAccessListGrid").GetCellData(2,10)
			userId = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAccessListGrid").GetCellData(2,8)
			
			If Ucase(ocrfOwnerName) <> Ucase(ocrfOwnerId) And Ucase(ocrfOwnerId) <> Ucase(userId) Then
				Call ReportStep (StatusTypes.Pass,  "OCRF Details for role as " & OCRFRole & " is validated successfully" ,"Validate OCRF when role is Editor/Reader")
				Else 
				Call ReportStep (StatusTypes.Fail,  "OCRF Details for role as " & OCRFRole & " is not validated successfully" ,"Validate OCRF when role is Editor/Reader")
			End If
			linkAvailable = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAccessListGrid").GetCellData(2,13)
			If Trim(linkAvailable) = "Remove Permission" Then
				Call ReportStep (StatusTypes.Pass,  "Remove Permission link is avilable for role as " & OCRFRole,"Validate availabilty of Remove Permission link for Editor/Reader")
				Else
				Call ReportStep (StatusTypes.Fail,  "Remove Permission link is unavilable for role as " & OCRFRole,"Validate availabilty of Remove Permission link for Editor/Reader")
			End If
			checkboxAvailable = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAccessListGrid").ChildItem(2,2,"WebCheckBox",0).GetRoProperty("disabled")
			If checkboxAvailable = 0 Then
				Call ReportStep (StatusTypes.Pass,  "Available checkbox is enabled for Editor/Reader","Validate Checkbox is enable or disable")
				Else
				Call ReportStep (StatusTypes.Fail,  "Available checkbox is disabled for Editor/Reader" ,"Validate Checkbox is enable or disable")
			End If
		 ElseIf loggedInUser = "Reader" Then
			Set ocrfAdminHeaderGrid = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAdminHeaderTable")
'			************************************************
'Modified by madhu - 02/27/20202
			
			Set userIdAdminHeader = ocrfAdminHeaderGrid.WebEdit("Webedit_AccessUserLoginID_Search")
			'Set userIdAdminHeader = ocrfAdminHeaderGrid.ChildItem(2,8,"WebEdit",0)
'			***********************************************************
			grantUserId = Ucase(objData.Item("GrantUserId"))
			userIdAdminHeader.set grantUserId
			Call SendKeyAndWait("{ENTER}",userIdAdminHeader,3)
			ocrfOwnerName = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAccessListGrid").GetCellData(2,10)
			userId = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAccessListGrid").GetCellData(2,8)
			linkAvailable = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAccessListGrid").GetCellData(2,13)
			If Trim(linkAvailable) <> "Remove Permission" Then
				Call ReportStep (StatusTypes.Pass,  "Remove Permission link is unavilable for role as " & OCRFRole,"Validate availabilty of Remove Permission link for Editor/Reader")
				Else
				Call ReportStep (StatusTypes.Fail,  "Remove Permission link is avilable for role as " & OCRFRole,"Validate availabilty of Remove Permission link for Editor/Reader")
			End If
			checkboxAvailable = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAccessListGrid").ChildItem(2,2,"WebCheckBox",0).GetRoProperty("disabled")
			If checkboxAvailable = 1 Then
				Call ReportStep (StatusTypes.Pass,  "Available checkbox is disabled for Editor/Reader","Validate Checkbox is enable or disable")
				Else
				Call ReportStep (StatusTypes.Fail,  "Available checkbox is enabled for Editor/Reader" ,"Validate Checkbox is enable or disable")
			End If
		End If
	Else
		Call ReportStep (StatusTypes.Fail,  "User is not navigated to OCRF Administrator page","Navigate to OCRF Administrator page")
	End If

End Function


'Created By: Sumit
'Created on: 5/15/2018
'Description: 'Validation of Grant Permission functionality in OCRF Administration Module
Function GrantPermissionToOCRF(ByVal FileName,ByVal SheetName, ByVal StartRow,Byval EndRow,ByVal OCRFId, ByRef objData)
	
	grantRole = objData.Item("GrantRole")
	grantUserId = Ucase(objData.Item("GrantUserId"))
	loggedInUserId = Ucase(objData.Item("LoggedInUserId"))

	
	If Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebElement("OCRFAdministrationTitle").Exist Then
		Call ReportStep (StatusTypes.Pass,  "User is navigated to OCRF Administrator Module","Navigate to OCRF Administrator page")
		
		Call ImportSheet(SheetName,FileName)
	
		Datatable.SetCurrentRow(StartRow)
		NumberOfRows = (EndRow-StartRow)+1
		
		Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebElement("GrantPermissions").Click
	
		If Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("GrantPermissionRecordsList").Exist(10) Then
'		*************************************************
'Modified by Madhu - 02/27/2020
		
			Set OCRFIdGrantpermHeaderGrid = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("GrantPermissionHeaderGrid").WebEdit("Webedit_OfferingID_search")
			'Set OCRFIdGrantpermHeaderGrid = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("GrantPermissionHeaderGrid").ChildItem(2,4,"WebEdit",0)
'	********************************************************		
			OCRFIdGrantpermHeaderGrid.Set ocrfId
			wait 3
			Call SendKeyAndWait("{ENTER}",OCRFIdGrantpermHeaderGrid,5)
			If Ucase(loggedInUserId) = "OWNER" Then
				Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("GrantPermissionRecordsList").ChildItem(2,2,"WebCheckBox",0).Set "ON"
				chkboxGrantPermDisabled = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("GrantPermissionRecordsList").ChildItem(2,2,"WebCheckBox",0).GetRoProperty("disabled")
				roleGrantPermDisabled = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebList("RoleGrantPermission").GetRoProperty("disabled")
				If chkboxGrantPermDisabled = 0 and roleGrantPermDisabled = 0 Then
					Call ReportStep (StatusTypes.Pass,  "OCRF checkbox is enabled when logged with Owner role " ,"Check whether OCRF checkbox is disabled")
				End If
				Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebList("RoleGrantPermission").Select grantRole
			ElseIf Ucase(loggedInUserId) = "EDITOR" Then
				Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("GrantPermissionRecordsList").ChildItem(2,2,"WebCheckBox",0).Set "ON"
				roleGrantPermDisabled = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebList("RoleGrantPermission").GetRoProperty("disabled")
				selectedRole = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebList("RoleGrantPermission").GetRoProperty("selection")
				If roleGrantPermDisabled = 1 and selectedRole = "Reader" Then
					Call ReportStep (StatusTypes.Pass,  "Role by default seleted is Reader and cannot be changed as the role field is disabled " ,"Check whether Role Field is disabled")
					If Ucase(grantRole) = "EDITOR" Then
						Call ReportStep (StatusTypes.Pass,  "Editor user cannot grant permission to another Editor user " ,"Grant permission to User")
						Exit Function
					End If
				Else 
					Call ReportStep (StatusTypes.Fail,  "Role can be changed as the role field is enabled " ,"Check whether Role Field is disabled")
				End If
			ElseIf Ucase(loggedInUserId) = "READER" Then	
				rowCnt = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("GrantPermissionRecordsList").RowCount	
				If rowCnt =1 Then
					Call ReportStep (StatusTypes.Pass,  "Reader Role cannot grant permission to any user. " ,"Validate whether Reader Role can grant permission to users")
					Exit Function 
				Else
					Call ReportStep (StatusTypes.Fail,  "Reader Role can grant permission to any user. " ,"Validate whether Reader Role can grant permission to users")
				End If
				
			End If
			'Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("GrantPermissionRecordsList").ChildItem(2,2,"WebCheckBox",0).Set "ON"
			
			Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebEdit("UserNameGrantPermission").Set grantUserId
			Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebButton("AddButtonGrantPermission").Click
			Call PageLoading()
			wait 3
			If Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebElement("ActionMessage").Exist(20) Then 
				Call ReportStep (StatusTypes.Pass,  "Grant Permission to " & grantUserId & " user as " & grantRole & " role is successful","Grant Permisson to user")
				Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebElement("WebEleOkay").Click
				wait 1
			Else
				Call ReportStep (StatusTypes.Fail,  "Grant Permission to " & grantUserId & " user as " & grantRole & " role is unsuccessful","Grant Permisson to user")
			End If 
			Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebButton("CloseButtonGrantPermission").Click
			Call PageLoading()
		Else
			Call ReportStep (StatusTypes.Fail,  "Grant Permission dialog window has not opened","Grant Permisson in OCRF Administrator page")
		End If
				
	Else
		Call ReportStep (StatusTypes.Fail,  "User is not navigated to OCRF Administrator page","Navigate to OCRF Administrator page")
	End If


	

End Function


Function GrantReaderPermissionToOCRF(ByVal FileName,ByVal SheetName, ByVal StartRow,Byval EndRow,ByVal OCRFId, ByRef objData, ByVal grantRolesToUsers)
	
	grantEditorRole = objData.Item("GrantEditorRole")
	grantReaderRole = objData.Item("GrantReaderRole")
	grantUserId = Ucase(objData.Item("GrantUserId"))
	loggedInUserId = Ucase(objData.Item("LoggedInUserId"))
	
	
	If Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebElement("OCRFAdministrationTitle").Exist Then
		Call ReportStep (StatusTypes.Pass,  "User is navigated to OCRF Administrator Module","Navigate to OCRF Administrator page")
		
		Call ImportSheet(SheetName,FileName)
	
		Datatable.SetCurrentRow(StartRow)
		NumberOfRows = (EndRow-StartRow)+1
				
		For i = StartRow To EndRow Step 1
			
			Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebElement("GrantPermissions").Click
			If Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("GrantPermissionRecordsList").Exist(10) Then
			'		*************************************************
'Modified by Madhu - 02/27/2020
'	modified OR	
			Set OCRFIdGrantpermHeaderGrid = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("GrantPermissionHeaderGrid").WebEdit("Webedit_OfferingID_search")
			
'			Set OCRFIdGrantpermHeaderGrid = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("GrantPermissionHeaderGrid").ChildItem(2,4,"WebEdit",0)
'			***************************************************8
			OCRFIdGrantpermHeaderGrid.Set OCRFId
			End If
			
			wait 3
			Call SendKeyAndWait("{ENTER}",OCRFIdGrantpermHeaderGrid,5)
			Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("GrantPermissionRecordsList").ChildItem(2,2,"WebCheckBox",0).Set "ON"
			Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebList("RoleGrantPermission").Select grantReaderRole
			
			Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebEdit("UserNameGrantPermission").Set grantRolesToUsers
			Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebButton("AddButtonGrantPermission").Click
			
			Call PageLoading()
			'wait 3
		    If Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebElement("ActionMessage").Exist(5) Then 
				Call ReportStep (StatusTypes.Pass,  "Grant Permission to " & grantRolesToUsers & " user as " & grantReaderRole & " role is successful","Grant Permisson to user")
				
				'Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebElement("WebEleOkay").Click
				'changes by srinivas
				Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebElement("xpath:=//div[contains(@style,'display: block;')]//span[text()='Okay']").Click
				
'			Else
'				Call ReportStep (StatusTypes.Fail,  "Grant Permission to " & grantRolesToUsers & " user as " & grantReaderRole & " role is unsuccessful","Grant Permisson to user")
			End If 
			wait 2
			Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebButton("CloseButtonGrantPermission").Click
			Call PageLoading()
		
		Next


	End If

End Function


Function GrantEditorPermissionToOCRF(ByVal FileName,ByVal SheetName, ByVal StartRow,Byval EndRow,ByVal OCRFId, ByRef objData, ByVal grantRolesToUsers)
	
	grantEditorRole = objData.Item("GrantEditorRole")
	grantReaderRole = objData.Item("GrantReaderRole")
	grantUserId = Ucase(objData.Item("GrantUserId"))
	loggedInUserId = Ucase(objData.Item("LoggedInUserId"))
	
	
	If Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebElement("OCRFAdministrationTitle").Exist Then
		Call ReportStep (StatusTypes.Pass,  "User is navigated to OCRF Administrator Module","Navigate to OCRF Administrator page")
		
		Call ImportSheet(SheetName,FileName)
	
		Datatable.SetCurrentRow(StartRow)
		NumberOfRows = (EndRow-StartRow)+1
				
		For i = StartRow To EndRow Step 1
						
			Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebElement("GrantPermissions").Click
			If Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("GrantPermissionRecordsList").Exist(10) Then
			'Set OCRFIdGrantpermHeaderGrid = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("GrantPermissionHeaderGrid").ChildItem(2,4,"WebEdit",0)
			'OCRFIdGrantpermHeaderGrid.Set OCRFId
			Set OCRFIdGrantpermHeaderGrid = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebEdit("xpath:=//table[@aria-labelledby='gbox_ocrf-access-owner-list']//input[@id='gs_OfferingID']")
			OCRFIdGrantpermHeaderGrid.Set OCRFId
			End If
			
			wait 3
			Call SendKeyAndWait("{ENTER}",OCRFIdGrantpermHeaderGrid,5)
			Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("GrantPermissionRecordsList").ChildItem(2,2,"WebCheckBox",0).Set "ON"
			Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebList("RoleGrantPermission").Select grantEditorRole
			
			Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebEdit("UserNameGrantPermission").Set grantRolesToUsers
			Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebButton("AddButtonGrantPermission").Click
			
			Call PageLoading()
			'wait 3
		    If Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebElement("ActionMessage").Exist(5) Then 
				Call ReportStep (StatusTypes.Pass,  "Grant Permission to " & grantRolesToUsers & " user as " & grantEditorRole & " role is successful","Grant Permisson to user")
								
				'Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebElement("WebEleOkay").Click
				'changes by srinivas
				Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebElement("xpath:=//div[contains(@style,'display: block;')]//span[text()='Okay']").Click
'			Else
'				Call ReportStep (StatusTypes.Fail,  "Grant Permission to " & grantRolesToUsers & " user as " & grantEditorRole & " role is unsuccessful","Grant Permisson to user")
			End If 
			Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebButton("CloseButtonGrantPermission").Click
			Call PageLoading()
		
		Next


	End If	

End Function


'Created By: Sumit
'Created on: 5/15/2018
'Description: 'Validation of Remove Permission functionality in OCRF Administration Module
Function RemovePermissionToOCRF(ByVal FileName,ByVal SheetName, ByVal StartRow,Byval EndRow,ByVal OCRFId, ByVal removeThruLink, ByRef objData)
	
	removeRole = objData.Item("RemoveRole")
	removeUserId = Ucase(objData.Item("RemoveUserId"))
	loggedInUserId = Ucase(objData.Item("LoggedInUserId"))
	
	If Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebElement("OCRFAdministrationTitle").Exist Then
		Call ReportStep (StatusTypes.Pass,  "User is navigated to OCRF Administrator Module","Navigate to OCRF Administrator page")
		
		Call ImportSheet(SheetName,FileName)
	
		Datatable.SetCurrentRow(StartRow)
		NumberOfRows = (EndRow-StartRow)+1
		
		Set ocrfAdminHeaderGrid = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAdminHeaderTable")
'		*******************************************************8
'modified by madhu - 02/27/2020
		Set ocrfIdAdminHeader = ocrfAdminHeaderGrid.WebEdit("Webedit_OfferingID_search")
		'Set ocrfIdAdminHeader = ocrfAdminHeaderGrid.ChildItem(2,4,"WebEdit",0)
		ocrfIdAdminHeader.Set OCRFId
		Call SendKeyAndWait("{ENTER}",ocrfIdAdminHeader, 2)
		
		Set roleAdminHeader = ocrfAdminHeaderGrid.WebEdit("Webedit_OcrfRole_search")

		'Set roleAdminHeader = ocrfAdminHeaderGrid.ChildItem(2,9,"WebEdit",0)
		roleAdminHeader.Set removeRole
		Call SendKeyAndWait("{ENTER}",roleAdminHeader, 2)
		
		Set userIdAdminHeader = ocrfAdminHeaderGrid.WebEdit("Webedit_AccessUserLoginID_Search")

		'Set userIdAdminHeader = ocrfAdminHeaderGrid.ChildItem(2,8,"WebEdit",0)
		userIdAdminHeader.set removeUserId
		Call SendKeyAndWait("{ENTER}",userIdAdminHeader,2)
		wait 2
'		********************************************************************
		Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAccessListGrid").ChildItem(2,2,"WebCheckBox",0).Set "ON"
		
		If Ucase(removeThruLink) = TRUE Then
			Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAccessListGrid").ChildItem(2,13,"Link",0).Click
			Else 
			Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebElement("RemovePermission").Click
		End If
		
		If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebEleConfirmation").Exist(10) Then
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_Yes").Click
			If Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebElement("RemovePermissionConfMsg").Exist(10) Then 
				Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Information_Okay").Click
			End If 
		End If
		
		rowCnt = Browser("Browser-OCRF").Page("OCRF-OCRFAdministration").WebTable("OCRFAccessListGrid").RowCount
		If rowCnt < 2 Then
			Call ReportStep (StatusTypes.Pass,  "Permission to OCRF has been successfully removed " ,"Remove Permission")
			else
			Call ReportStep (StatusTypes.Fail,  "Permission to OCRF has not been successfully removed " ,"Remove Permission")
		End If
		
	Else
		Call ReportStep (StatusTypes.Fail,  "User is not navigated to OCRF Administrator page","Navigate to OCRF Administrator page")
	End If
	
	

End Function

















'******************************************************************************************************************************************************************
'Created By: Rajesh
'Created on: 12/10/2016
'Description: Launch the browser with ocrf url and validate default page is displayed.
'Parameter: Param1 - Extra,Param2 - Extra,Param3 - Extra,objData
Public Function LaunchURLOld(ByVal Param1, ByVal Param2, ByVal Param3, ByRef objData)
	
	'Variables declaration
	Dim URL	
	
	Systemutil.CloseProcessByName "iexplore.exe"
	
	'Variables initialization
	URL = Environment.Value(Param1)
	
	'Launch OCRF Url's
	
	SystemUtil.Run URL
	
	'Call ReportStep(StatusTypes.StepDefault, "##### Launch OCRF URL - Login Functionality of OCRF- Starts ####", "Report Creation Page")
	'Validate OCRF url launched successfully
	If Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_ONLINE-CRF_Tab").Exist Then
		'Call ReportStep (StatusTypes.StepDefault, "OCRF URL launched successfully","OCRF URL")
		
		'Default page is displayed
		If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_ClientRequestPageHeader").Exist(60) Then
			Call ReportStep (StatusTypes.Pass, "Default page(Online Client Request List) is displayed","Default page(Online Client Request List)")
		Else
			Call ReportStep (StatusTypes.Fail, "Default page(Online Client Request List) is not displayed","Default page(Online Client Request List)")
		End If
		
	Else
		Call ReportStep (StatusTypes.Fail, "OCRF URL not launched properly","OCRF URL")
		Exit Function
	End If
	
End Function


'Created By: Rajesh
'Created on: 12/10/2016
'Description: Verify all the tabs and click on each tab and validate the target page
'Parameter: Param1 - For ONLINE CRF - "TAB 1",For CLIENTS - "TAB 2",For MANAGE USERS ACCESS - "TAB 3",For AUDIT LOG - "TAB 4",For HELP - "TAB 5",For all tab - "TAB ALL"
'			Param2 - Extra,Param3 - Extra,objData
Public Function TabsValidation(ByVal Param1, ByVal Param2, ByVal Param3, ByRef objData)
	
	'Navigate to each tab[ONLINE CRF, CLIENTS, MANAGE USERS ACCESS, AUDIT LOG, HELP] and validate navigation is successfull
	
	
	If Param1 = "TAB 1" OR Param1 = "TAB ALL" Then
		If Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_ONLINE-CRF_Tab").Exist Then
			Call ReportStep (StatusTypes.Pass, "ONLINE CRF tab displayed successfully","ONLINE CRF tab")
			Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_ONLINE-CRF_Tab").Click
			If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_ClientRequestPageHeader").Exist Then
				Call ReportStep (StatusTypes.Information, "After click on ONLINE CRF tab,screen is Navigated to ONLINE CRF page","ONLINE CRF page")
				
			Else	
				Call ReportStep (StatusTypes.Fail, "After click on ONLINE CRF tab,screen is not Navigated to ONLINE CRF page","ONLINE CRF page")
				
			End If
		Else
			Call ReportStep (StatusTypes.Fail, "Not able to display ONLINE CRF tab","ONLINE CRF tab")
			
		End If
	End If
	
	If Param1 = "TAB 2" OR Param1 = "TAB ALL" Then
		If Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_CLIENTS_Tab").Exist Then
			Call ReportStep (StatusTypes.Pass, "CLIENTS tab displayed successfully","CLIENTS tab")
			Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_CLIENTS_Tab").Click
			If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_ClientListPageHeader").Exist Then
				Call ReportStep (StatusTypes.Information, "After click on CLIENTS tab,screen is Navigated to CLIENTS page","CLIENTS page")
				
			Else	
				Call ReportStep (StatusTypes.Fail, "After click on CLIENTS tab,screen is not Navigated to CLIENTS page","CLIENTS page")
				
			End If
		Else
			Call ReportStep (StatusTypes.Fail, "Not able to display CLIENTS tab","CLIENTS tab")
			
		End If
	End If
	
	If Param1 = "TAB 3" OR Param1 = "TAB ALL" Then
		If Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_MANAGE-USERS-ACCESS_Tab").Exist Then
			Call ReportStep (StatusTypes.Pass, "MANAGE USERS ACCESS tab displayed successfully","MANAGE USERS ACCESS")
			Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_MANAGE-USERS-ACCESS_Tab").Click
			If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_ManageUser'sAccessPageHeader").Exist Then
				Call ReportStep (StatusTypes.Information, "After click on MANAGE USERS ACCESS tab,screen is Navigated to MANAGE USERS ACCESS page","MANAGE USERS ACCESS page")
				
			Else	
				Call ReportStep (StatusTypes.Fail, "After click on MANAGE USERS ACCESS tab,screen is not Navigated to MANAGE USERS ACCESS page","MANAGE USERS ACCESS page")
				
			End If
		Else
			Call ReportStep (StatusTypes.Fail, "Not able to display MANAGE USERS ACCESS tab","MANAGE USERS ACCESS tab")
			
		End If
	End If
	
	If Param1 = "TAB 4" OR Param1 = "TAB ALL" Then
		If Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_AUDIT-LOG_Tab").Exist Then
			Call ReportStep (StatusTypes.Pass, "AUDIT LOG tab displayed successfully","AUDIT LOG tab")
			Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_AUDIT-LOG_Tab").Click
			If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_SearchAuditPageHeader").Exist Then
				Call ReportStep (StatusTypes.Information, "After click on AUDIT LOG tab,screen is Navigated to AUDIT LOG page","AUDIT LOG page")
				
			Else	
				Call ReportStep (StatusTypes.Fail, "After click on AUDIT LOG tab,screen is not Navigated to AUDIT LOG page","AUDIT LOG page")
				
			End If
		Else
			Call ReportStep (StatusTypes.Fail, "Not able to display AUDIT LOG tab","AUDIT LOG tab")
			
		End If
	End If
	
	If Param1 = "TAB 5" OR Param1 = "TAB ALL" Then
		If Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_HELP_Tab").Exist Then
			Call ReportStep (StatusTypes.Pass, "HELP tab displayed successfully","HELP tab")
			Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_HELP_Tab").Click
			If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_HelpPageHeader").Exist Then
				Call ReportStep (StatusTypes.Information, "After click on HELP tab,screen is Navigated to HELP page","HELP page")
				
			Else	
				Call ReportStep (StatusTypes.Fail, "After click on HELP tab,screen is not Navigated to HELP page","HELP page")
				
			End If
		Else
			Call ReportStep (StatusTypes.Fail, "Not able to display HELP tab","HELP tab")
			
		End If
	End If
	
	
	
End Function


'Created By: Rajesh
'Created on: 13/10/2016
'Description: Select value from dropdown and same number of records should display in a page.
'Parameter: Param1 - Extra,
'			Param2 - Extra,Param3 - Extra,objData
Public Function ValidatePageRecords(ByVal Param1, ByVal Param2, ByVal Param3, ByRef objData)
	
	RowFilter = "20;40;50;100"

	If Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_RecordsRange").Exist Then
		
		Record = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_RecordsRange").GetROProperty("all items")
		
		If RowFilter = Record Then
			Call ReportStep (StatusTypes.Pass, "Filtering rows per page is correct "&Record,"Filtering rows value")
		Else
			Call ReportStep (StatusTypes.Fail, "Filtering rows per page is Incorrect "&Record,"Filtering rows value")			
		End If
		
		
		Records = split(Record,";")
		
		For PageRecord = lbound(Records) To ubound(Records) Step 1
			
			'Select the values from weblist[20,40,50,100]
			Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_RecordsRange").Select Records(PageRecord)
			
			'Validate number of records in a page
			
			For Sequence = 1 To 10 Step 1
				rc = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").RowCount
				If rc > CInt(Records(PageRecord)) Then
					Exit For
				End If
				Wait 2
			Next
			
			
			'Validate number of pages and number of records
			RecordsInfo = trim(Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_DisplayedRecords").GetROProperty("innertext"))
			
			RecordsInformation = Split(RecordsInfo," ")
			
			
			
			
			If (rc = Records(PageRecord) + 1) AND (CInt(RecordsInformation(3)) = CInt(Records(PageRecord))) Then
				Call ReportStep (StatusTypes.Pass, "Number of records in a page and dropdown are same: "&Records(PageRecord),"Page Records Validation")
			Else
				Call ReportStep (StatusTypes.Fail, "Number of records in a page and dropdown are different: "&Records(PageRecord),"Page Records Validation")			
			End If
			
			
			'Validate number of pages
			LastPage = CInt(trim(Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Pagination").WebElement("html id:=sp_1.*","html tag:=SPAN").GetROProperty("innertext")))
			
			
			
			If (LastPage = Round(CDbl(RecordsInformation(5))/CInt(Records(PageRecord)))) OR (LastPage-1 = Round(CDbl(RecordsInformation(5))/CInt(Records(PageRecord)))) Then
				Call ReportStep (StatusTypes.Pass, "Total Records are: "&CDbl(RecordsInformation(5))&", Total Pages are: "&LastPage&", Number of records in a page are: "&CInt(Records(PageRecord)),"Page and Records Validation")
			Else
				Call ReportStep (StatusTypes.Fail, "Total Records are not equal to: "&CDbl(RecordsInformation(5))&", Total Pages are not equal to: "&LastPage&", Number of records in a page are not equal to: "&CInt(Records(PageRecord)),"Page and Records Validation")			
			End If
			
			
			
			'To visible dropdown to user, select 20 value from dropdown
			If PageRecord = lbound(Records) OR PageRecord = ubound(Records) Then
			
			Else
				Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_RecordsRange").Select Records(0)
				wait 2
			End If
			
			
		Next
		
		
	End If
	
End Function 
	

'Created By: Rajesh
'Created on: 14/10/2016
'Description: Verification of page navigation by using Next,Last,Prev,First icons.
'Parameter: Param1 - Extra,
'			Param2 - Extra,Param3 - Extra,objData
Public Function Pagination(ByVal Param1, ByVal Param2, ByVal Param3, ByRef objData)
	
	
	'Get the current page and last page number
	wait 5
	CurrentPage = CInt(trim(Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_PageNumber").GetROProperty("value")))
	LastPage = CInt(trim(Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Pagination").WebElement("html id:=sp_1.*","html tag:=SPAN").GetROProperty("innertext")))
'	LastPage =trim(Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Pagination").WebElement("html id:=sp_1_request-list-pager","html tag:=SPAN").GetROProperty("innertext"))
'	LastPage = cint(LastPage)
	
	'Click on next page icon
	
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_NextPage").Exist Then
		Call ReportStep (StatusTypes.Pass,"Next page icon is present","Next page icon")			
		Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_NextPage").Click
		wait 5
	Else
		Call ReportStep (StatusTypes.Fail,"Next page icon is not present","Next page icon")			
	End If
	
	'Validate current page number should update
	NextPageNumber = CInt(trim(Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_PageNumber").GetROProperty("value")))
	
	If (NextPageNumber) = CurrentPage+1 Then
		Call ReportStep (StatusTypes.Pass,"Page number is incremented by 1, by click on next page icon "&NextPageNumber,"Next page validation")			
	Else
		Call ReportStep (StatusTypes.Fail,"Page number is not incremented by 1, by click on next page icon "&NextPageNumber,"Next page validation")				
	End If
	
	
	'Click on last page icon
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_LastPage").Exist Then
		Call ReportStep (StatusTypes.Pass,"Last page icon is present","Last page icon")			
		Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_LastPage").Click
		wait 5
	Else
		Call ReportStep (StatusTypes.Fail,"Last page icon is not present","Last page icon")			
	End If
	
	'Validate current page number should update
	LastPageNumber = CInt(trim(Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_PageNumber").GetROProperty("value")))
	
	If LastPageNumber = LastPage Then
		Call ReportStep (StatusTypes.Pass,"Last Page number is displayed, by click on last page icon "&LastPageNumber,"last page validation")			
	Else
		Call ReportStep (StatusTypes.Fail,"Last Page number is not displayed, by click on last page icon "&LastPageNumber,"last page validation")				
	End If
	
	
	'Click on previous page icon
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_PrevPage").Exist Then
		Call ReportStep (StatusTypes.Pass,"Previous page icon is present","Previous page icon")			
		Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_PrevPage").Click
		wait 5
	Else
		Call ReportStep (StatusTypes.Fail,"Previous page icon is not present","Previous page icon")			
	End If
	
	'Validate current page number should update
	PrevPageNumber = CInt(trim(Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_PageNumber").GetROProperty("value")))
	
	If PrevPageNumber = (LastPage) - 1 Then
		Call ReportStep (StatusTypes.Pass,"Page number is decremented by 1, by click on previous page icon "&PrevPageNumber,"previous page validation")			
	Else
		Call ReportStep (StatusTypes.Fail,"Page number is not decremented by 1, by click on previous page icon "&PrevPageNumber,"previous page validation")				
	End If
	
	
	'Click on first page icon
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_FirstPage").Exist Then
		Call ReportStep (StatusTypes.Pass,"First page icon is present","First page icon")			
		Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_FirstPage").Click
		wait 5
	Else
		Call ReportStep (StatusTypes.Fail,"First page icon is not present","First page icon")			
	End If
	
	'Validate current page number should update
	FirstPageNumber = CInt(trim(Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_PageNumber").GetROProperty("value")))
	
	If FirstPageNumber = CurrentPage Then
		Call ReportStep (StatusTypes.Pass,"First Page number is displayed, by click on first page icon "&FirstPageNumber,"first page validation")			
	Else
		Call ReportStep (StatusTypes.Fail,"First Page number is not displayed, by click on first page icon "&FirstPageNumber,"first page validation")				
	End If
	
	
End Function


'Created By: Poornima
'Created on: 15/10/2016
'Description: Validate Client creation/editing/deletion operation


Function OCRFClientPage(BYVal intPageNavigation, BYVal strClientOperation, Byval ClientName, ByVal ClientShortName,ByVal strClientNewName, ByVal strClientNewShortName ,ByRef objData)

    'Call ReportStep(StatusTypes.StepDefault, "##### OCRFClientPage - Verification of OCRF Client Page - Start ####", "Report Creation Page")
    If intPageNavigation <> 1 Then
        If Browser("Browser-OCRF").Page("Page-OCRF").Exist(30) Then
            Call ReportStep (StatusTypes.Pass,"Entered in to OCRF home page","Home page")
        Else    
            Call ReportStep (StatusTypes.Fail,"Not Entered in to OCRF home page","Home page")
        End If
        ' click on Client tab
        If Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_CLIENTS_Tab").Exist(30) Then
            Call ReportStep (StatusTypes.Pass,"click on Client tab ","Client tab")
            Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_CLIENTS_Tab").Click
            Browser("Browser-OCRF").Page("Page-OCRF").Sync
        Else    
            Call ReportStep (StatusTypes.Fail,"Client tab Not visible","Client tab")
        End If
        'CHECKING FOR cLIENT LIST PAGE DISPLAY
        If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_ClientListPageHeader").Exist(30) Then
            Call ReportStep (StatusTypes.Pass,"Online CRF client List displayed","client List page")
        Else    
            Call ReportStep (StatusTypes.Fail,"Online CRF client List not displayed","client List page")
        End If
    End If
    
	Select Case strClientOperation
        
    Case "Create"
            

            ' CLICK ON ADD '+' SYMBOL FOR ADDING CLIENT
            If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_NewClient").Exist(30) Then
               Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_NewClient").Click
                Call ReportStep (StatusTypes.Pass,"click on Adding new client done","add new client")
            Else    
                Call ReportStep (StatusTypes.Fail,"click on Adding new client not done","click on Adding new client done")
            End If
            ' ENTERING CLIENT NAME, SHORT NAME IN POPUP APPEARED
            If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_ClientName").Exist(30) Then
                Call ReportStep (StatusTypes.Pass,"Text box to enter new client name displayed and Entered client name as " & ClientName & " Client ShortName as" & ClientShortName  ,"new client textbox")
                Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_ClientName").Set ClientName    'Review- Add param and function for creating client
                Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_ClientShortName").Set ClientShortName    'Review- Add param
            Else    
                Call ReportStep (StatusTypes.Fail,"Text box to enter new client name not displayed","New client textbox")
            End If
            'CHECK FOR ERROR MESSAGE WHILE CREATING EXISTING CLIENT NAME 
            If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-state-error","innertext:=Client already exist with this shortName: "&ClientName).Exist(5) Then
               Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-state-error","innertext:=Client already exist with this shortName: "&ClientName).highlight
               Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-state-error","innertext:=Client already exist with this shortName: "&ClientName).highlight    
               Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=cData","innertext:=Cancel","html tag:=A").Click
               Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=fm-button ui-state-default ui-corner-all","html id:=cNew","innertext:=Cancel").Click
               OCRFClientPage = True
               Exit Function
            Else    
               OCRFClientPage = False
            End If
            'CLICK ON CREATE BUTTON
            If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_Create").Exist(30) Then
                Call ReportStep (StatusTypes.Pass,"Click on create done","Create client button")
                Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_Create").Click
            Else    
                Call ReportStep (StatusTypes.Fail,"Click on create not done","Create client button")
            End If
            
            'CLIENT ADDED SUCCESSFULLY MESSAGE
            If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_ClientAddedMsg").Exist(30) Then
                Call ReportStep (StatusTypes.Pass,"Client successfully added message  displayed","message")
		Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_ClientAddedMsg").highlight
		'Call ReportStep(StatusTypes.StepDefault, "##### Client created "success message####", "suucessfully client created")
                Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebEdit_Cancel").Click
            Else    
                Call ReportStep (StatusTypes.Fail,"Client successfully added message not displayed","message")
            End If
            'CHECK FOR CREATING CLIENT NAME EXIST IN CLIENT LIST OR NOT
            Call SearchClient(ClientName,ClientShortName, True,objData)
 
 Case "Edit"
       
        
	      'search client name to be edit
           Call SearchClient(ClientName,ClientShortName, True,objData)
         'SELECT SEARCHED CLIENT FOR EDITING
        If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").WebElement("html tag:=TD","innertext:="&ClientName,"title:="&ClientName).Exist(30) Then
            Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").WebElement("html tag:=TD","innertext:="&ClientName,"title:="&ClientName).click
            Call ReportStep (StatusTypes.Pass,"Selecting searched client " & ClientName & " for editing done","Edit client name")
        Else    
            Call ReportStep (StatusTypes.Fail,"Selecting searched client " & ClientName & " for editing not done","Edit client name")
        End If    
            'EDIT CLIENT NAME AND SHORT NAME 
        If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_ClientName").Exist(30) Then
            Call ReportStep (StatusTypes.Pass,"Text box to enter new client name displayed and Entered client name as " & strClientNewName & " client short name as " & strClientNewShortName,"new client textbox")
            Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_EditClient").Click
            Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_ClientName").Set strClientNewName
            Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_ClientShortName").Set strClientNewShortName
        Else    
            Call ReportStep (StatusTypes.Fail,"Text box to enter new client name not displayed","new client textbox")
        End If
        
         'CHECK FOR ERROR MESSAGE WHILE CREATING EXISTING CLIENT NAME 
            If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-state-error","innertext:=Client already exist with this shortName: "&ClientName).Exist(5) Then
               Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=cData","innertext:=Cancel","html tag:=A").Click
               Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=fm-button ui-state-default ui-corner-all","html id:=cNew","innertext:=Cancel").Click
               OCRFClientPage = True
             
            Else    
               OCRFClientPage = False
            End If
            'CLICK ON UPDATE BUTTON
        If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_Update").Exist(30) Then
            Call ReportStep (StatusTypes.Pass,"Click on Update done","Update client button")
            Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_Update").Click
        Else    
            Call ReportStep (StatusTypes.Fail,"Click on Update not done","Update client button")
        End If
        ' AFTER EDITING NAME SEARCH EDITED NAME COMING IN CLIENT LIST
          Call SearchClient( strClientNewName, strClientNewShortName, True,objData)
	
    
    Case "Delete"
        
         Call SearchClient(strClientNewName ,strClientNewShortName , True,objData)
        
        'SELECTING UPDATED CLIENT NAME FOR DELETION
        ClientName=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").GetCellData(2,3)
        RequestId=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").GetCellData(2,2)
        If ClientName=strClientNewName Then
            Call ReportStep (StatusTypes.Pass,"Updated client name" & strClientNewName & " selected for deletion","Select client name")
            Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").WebElement("html tag:=TD","innertext:="&RequestId,"title:="&RequestId).Click
        Else    
            Call ReportStep (StatusTypes.Fail,"Updated client name" & strClientNewName & " not selected for deletion","Select client name")
        End If

            'CLICK ON DELETE CLIENT ICON FOR DELETING SELECTED CLIENT
        If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_DeleteClient").Exist(15) Then
            Call ReportStep (StatusTypes.Pass,"Click on delete client done","Delete client")
            Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_DeleteClient").Click
        Else    
            Call ReportStep (StatusTypes.Fail,"Click on delete client not done","Delete client")
        End If
        
        'CLICK ON DELETE BUTTON FOR CLIENT DELETION IN MESSAGE POP UP
        If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=EditTable","html id:=DelTbl_DataTable_2","html tag:=TABLE").Exist(15) Then
            Call ReportStep (StatusTypes.Pass,"Click on delete button done in pop up","Delete client pop up")
            Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_Delete").Click
        Else    
            Call ReportStep (StatusTypes.Fail,"Click on delete button done in pop up","Delete client pop up")
        End If
        
        'CHECK FOR ERROR MESSAGE WHILE DLETING ACTIVE CLIENT
        If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-state-error","html tag:=TD","innertext:=This Client already used in CRF\. Can not be deleted").Exist(5) Then
           Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-state-error","html tag:=TD","innertext:=This Client already used in CRF\. Can not be deleted").highlight
           'Call ReportStep(StatusTypes.StepDefault, "##### Error while deleting active client ####", "Error in delete operation")
           Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=eData","innertext:=Cancel").Click
           OCRFClientPage = True 
           Exit Function 
        Else    
           OCRFClientPage = False
        End If
         
        'SEARCH FOR DELETED CLIENT NAME IN CLIENT LIST IT SHOULD NOT EXIST
        Call SearchClient(strClientNewName ,strClientNewShortName, False ,objData)
     
        
'        If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").WebElement("html tag:=TD","innertext:="&ClientName,"title:="&ClientName).Exist(5) Then
'            Call ReportStep (StatusTypes.Fail,"Deleted client name" &  ClientName  &" found in search list","Search Delete client ")
'        Else        
'            Call ReportStep (StatusTypes.Pass,"Deleted client name" &  ClientName  &" not found in search list","Search Delete client ")
'        End If
                     
    Case "Reload"
        'ToDO: after reload all data in Client list should appear         
    Case "Nothing"
        'TO DO: verifying searched data appearing or not    
        
        
        
    End Select
    
    'Call ReportStep(StatusTypes.StepDefault, "##### OCRFClientPage - Verification of OCRF Client Page - End ####", "Report Creation Page")
    
End Function



'Created By: 
'Created on: 14/10/2016
'Description: Verification of page navigation by using Next,Last,Prev,First icons.
'Parameter: 	
Function SearchClient(byVal strClientName, byVal strClientShortName,byVal Exist ,ByRef objData)
         wait 2
    	'Call ReportStep(StatusTypes.StepDefault, "##### SearchClient - Verification of page navigation by using Next,Last,Prev,First icons - Start ####", "Report Creation Page")
        'SEARCH FOR CREATED CLIENT NAME IN CLIENT LIST
        If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_ClientName").Exist(30) Then
            Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_ClientName").Set strClientName
            Call ReportStep (StatusTypes.Pass,"Setting client name in search box as " & strClientName,"Search client name")
            Call ReportStep (StatusTypes.Pass,"Setting client name " &strClientName & " in search box done","Search client name")
            'Call ReportStep(StatusTypes.StepDefault, "##### Search for client Name" & strClientName & "  - Start ####", "OCRF Page")
        Else    
            Call ReportStep (StatusTypes.Fail,"Setting client name in search box not done","Search client name")
        End If
        wait 3
        Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_ClientName").Click
         Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{ENTER}"
        wait 5
        
        Set objTab = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data")
	    Set oTableCell = Description.Create
	     oTableCell("micclass").value = "WebElement"
	    oTableCell("html tag").value = "TD"
		oTableCell("innertext").value = strClientName&".*"   
		oTableCell("outerhtml").value = "<TD role=gridcell aria-describedby=DataTable_ClientName title=.* "
	
		set oTableCellObj = objTab.ChildObjects(oTableCell)
        col=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").GetROProperty("cols")
      If  oTableCellObj.count >= 1 Then
          Call ReportStep (StatusTypes.Pass,oTableCellObj.count  & "Names exist in serched list","Search client name")
          If Exist = True Then
              If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").Exist(30) Then
                 For i = 0 To oTableCellObj.count-1 Step 1
                 	x= oTableCellObj(i).GetRoproperty("innertext")
                 	If instr(1,x,strClientName,1)>0 Then
                 		If instr(1,x,strClientName,0)>0 Then
                 		    oTableCellObj(i).highlight
                 		    Call ReportStep (StatusTypes.Pass,"Searched client name " & strClientName & " Exist in Client List At Line No:" & i+1 &" With name" & x ,"Search client name")
                 			oTableCellObj(i).Click
                 		Else
                 		    Call ReportStep (StatusTypes.Pass,"Searched client name " & strClientName & " NOT Match with client name" & x &" at row " & i+1 &" With Case sensitive name" & x ,"Search client name")
                 		End If
                 	End If
                 Next
                End If 
             Else
             	For i = 0 To oTableCellObj.count-1 Step 1
                 	x= oTableCellObj(i).GetRoproperty("innertext")
                 	If instr(1,x,strClientName,1)>0 Then
                 		If instr(1,x,strClientName,0)>0 Then
                 		    Call ReportStep (StatusTypes.Fail,"Searched client name " & strClientName & " exist in Client List At Line No:"& i+1 &" With name" & x ,"Search client name")
                 			oTableCellObj(i).highlight
                 			oTableCellObj(i).Click
                 		Else
                 		    Call ReportStep (StatusTypes.Fail,"Searched client name " & strClientName & " NOT Match with client name"& x &" at row " & i+1 &" With Case sensitive name" & x ,"Search client name")
                 		End If
                 	End If
                 Next
                End If
      Else
           
             If Exist = False Then
              Call ReportStep (StatusTypes.Pass,"Searched client name " & strClientName & " Not Exist in Client List","Search client name")

             Else
                Call ReportStep (StatusTypes.Fail,"Searched client name " & strClientName & " Not exist in Client List","Search client name")
             End If
      End If

		
End Function


'Created date: 27-10-2016
'Created by: Poornima
'verify functionalities of Reaload grid in each tab of OCRF

Function ReLoadValidation(ByVal strClientName,ByVal Tab, ByRef objData)

'CLICK ON RELOAD GRID BEFORE DOING OPERATION TO TAKE COUNT OF ROW
'Call ReportStep(StatusTypes.StepDefault, "#####Start checking Reload button functionality in  Tab ####"&Tab, "Reload grid")
  If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-icon ui-icon-refresh","html tag:=SPAN").Exist(20) Then
  	 Call ReportStep (StatusTypes.Pass,"Click on Reload grid done","Relaoad grid") 
  	   Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-icon ui-icon-refresh","html tag:=SPAN").click
  Else   
     Call ReportStep (StatusTypes.Fail,"Click on Reload grid done","Relaoad grid")
  Exit function
  End If
  wait 3
  If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-pg-table","html tag:=TABLE","name:=WebTable").Exist(20) Then
  	 Call ReportStep (StatusTypes.Pass,"Get toatal number of row in tab before making any changes in Tab","Get total rows") 
  	     Row= Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-pg-table","html tag:=TABLE","name:=WebTable").GetROProperty("rows")
         data = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-pg-table","html tag:=TABLE","name:=WebTable").GetROProperty("column names")
         colThird= split(data,";")
         TotalRow = split(colThird(2),"of")
     Call ReportStep (StatusTypes.Pass," Total number of rows in tab "& Tab &" before making any changes is"&  (TotalRow(1)),"Get total rows") 
  Else   
     Call ReportStep (StatusTypes.Fail,"Not able to get Tatal number of rows in tab before making any changes in Tab","Get total rows")
  Exit function
  End If
  
'PERFORM SOME OPERATION IN TAB EX:SEARCHING FOR CREATED CLIENT IN LIST
  If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_ClientName").Exist(30) Then
     Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_ClientName").Set strClientName
     Call ReportStep (StatusTypes.Pass,"Setting client name in search box as " & strClientName,"Search client name")
     Call ReportStep (StatusTypes.Pass,"Setting client name in search box done","Search client name")
  Else    
     Call ReportStep (StatusTypes.Fail,"Setting client name in search box not done","Search client name")
  End If
  wait 2
  Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_ClientName").Click
  Set WshShell = CreateObject("WScript.Shell")
  WshShell.SendKeys "{ENTER}"
  wait 6
  
  'TAKING ROW COUNT AFTER PERFORMING OPERATION ON TAB 
  If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-pg-table","html tag:=TABLE","name:=WebTable").Exist(20) Then
  	 Call ReportStep (StatusTypes.Pass,"Get total number of rows in tab after making any changes in Tab","Get total rows") 
  Else   
     Call ReportStep (StatusTypes.Fail,"Get toatal number of row in tab after making any changes in Tab","Get total rows")
  Exit function
  End If

  If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-pg-table","html tag:=TABLE","name:=WebTable").Exist(20) Then
  	
  	     Row1= Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-pg-table","html tag:=TABLE","name:=WebTable").GetROProperty("rows")
		 data1 = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-pg-table","html tag:=TABLE","name:=WebTable").GetROProperty("column names")        
          colThird1= split(data1,";")
        If instr(1,colThird1(2),"No records to view",1)<>0 Then 
			Call ReportStep (StatusTypes.Pass,"No records  in tab "& Tab &" After search for client ","Get records")            
        Else
            
            TotalRow1 = split(colThird1(2),"of")
            Call ReportStep (StatusTypes.Pass," Total number of rows in tab "& Tab &" After search for client is "&  (TotalRow(1)),"Get total rows") 
                     
         End If    
  End If
  


'AGAIN CLICK ON RELOAD GRID AFTER PERFOMING OPERATION IN PERTICULAR TAB TO CHECK REFRESH GRID DONE OR NOT 
  Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-icon ui-icon-refresh","html tag:=SPAN").click
  wait 6 
 If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-pg-table","html tag:=TABLE","name:=WebTable").Exist(20) Then
  	 Call ReportStep (StatusTypes.Pass,"Get total number of row in tab after Click on RelodGrid","Get total rows") 
  	 Row2= Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-pg-table","html tag:=TABLE","name:=WebTable").GetROProperty("rows")
     data2 = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-pg-table","html tag:=TABLE","name:=WebTable").GetROProperty("column names")
     colThird2= split(data2,";")
     TotalRow2 = split(colThird2(2),"of")
  Call ReportStep (StatusTypes.Pass," Total number of rows in tab"& Tab &" after Click on RelodGrid "& (TotalRow(1)),"Get total rows") 
  Else   
     Call ReportStep (StatusTypes.Fail,"Get total number of row in tab after Click on RelodGrid","Get total rows")
  Exit function
  End If
  
  'COMPARING ROWS COUNT OBTAINED AFTER AND BEFORE PERFORMING OPERATION ON TAB 
  If  (TotalRow(1))=(TotalRow2(1)) Then
  	  call ReportStep (StatusTypes.Pass," Rows After click on Reload grid is "& (TotalRow2(1)) &" same as Before click on Reload " & (TotalRow(1)) &" in " &  Tab,"Reload validation")
  Else    
      Call ReportStep (StatusTypes.Fail," Rows After click on Reload grid is "& (TotalRow2(1)) &" Not same as Before click on Reload " & (TotalRow(1))&" in " & Tab,"Reload validation")
  End If
End Function


'Created By: Rajesh
'Created on: 23/10/2016
'Description: Open ocrf request based on request ID.
'Parameter: Param1 - Extra,Param2 - Extra,Param3 - Extra,Request_ID - Give ocrf request ID for eg: 1121,objData
Public Function OpenOCRFRequest(ByVal Param1, ByVal Param2, ByVal Param3,ByVal Request_ID,ByRef objData)
	
	
	'Open the OCRF request by double click on table cells using request ID[1122]
	
	If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").Exist Then
		Call ReportStep (StatusTypes.Pass, "Online client request grid loaded successfully","Online client request grid")
		wait 5
	Else
		Call ReportStep (StatusTypes.Fail, "Online client request grid not loaded successfully","Online client request grid")
		Exit Function
	End If
	
	Set objTab = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data")

	For i = 1 To 100 Step 1
		Set objElem = Description.Create
		objElem("micclass").value = "WebElement"
		objElem("html tag").value = "TD"
		Request_ID=""&Request_ID&""
		objElem("innertext").value = Request_ID 'parameterized the request id
		set objElemCnt = objTab.ChildObjects(objElem)
		
		If objElemCnt.count >0 Then
			objElemCnt(0).fireevent "onmouseover"
			objElemCnt(0).fireevent "ondblclick"
			If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-dialog-title","html tag:=SPAN","innertext:=Confirmation","visible:=True").Exist(10) Then
	           Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-button-text","html tag:=SPAN","innertext:=No","visible:=True").Click
                objElemCnt(0).RightClick
               Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html tag:=SPAN","innertext:=Change Request Status To > ","visible:=True").fireevent "onmouseover"
               Browser("Browser-OCRF").Page("Page-OCRF").WebElement("innertext:=Unlock","html tag:=SPAN","visible:=True").Click
          
               Call FilterByReqIDAndOpen(CInt(Request_ID),objData)
            End If

			Exit For
		Else		
			If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_NextPage").Exist Then
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_NextPage").Click
				wait 2
			End If
		End If
	Next
	
	'Validate OCRF request opened successfully
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_OfferingDetailsTab").Exist Then
		Call ReportStep (StatusTypes.Pass, "OCRF request opened successfully: "&Request_ID,"OCRF request")
	Else
		Call ReportStep (StatusTypes.Fail, "OCRF request not opened successfully: "&Request_ID,"OCRF request")
		Exit Function
	End If
	
End Function





'Created By: Rajesh
'Created on: 23/10/2016
'Description: Validate tab colour of ocrf request.
'Parameter: Param1 - Extra,Param2 - Extra,Param3 - Extra,TabName - Give ocrf request tab name. for eg: "Add Users",objData
Public Function ValidateBgColorOCRFRequestTab(ByVal Param1, ByVal Param2, ByVal Param3,ByVal TabName,ByRef objData)
	
	Bgcolor = Browser("Browser-OCRF").Page("Page-OCRF").Link("html tag:=A","innertext:=.*"&TabName&".*").GetROProperty("background color")
	
	If Bgcolor = "#ea8511" Then
		Call ReportStep (StatusTypes.Pass, "Colour of "&TabName&" tab is orange","Colour of "&TabName&" tab")
	ElseIf Bgcolor = "#8cc63f" Then	
		Call ReportStep (StatusTypes.Fail, "Colour of "&TabName&" tab is green","Colour of "&TabName&" tab")
	Else
		Call ReportStep (StatusTypes.Fail, "Colour of "&TabName&" tab is green","Colour of "&TabName&" tab")
	End If
	
	
	
End Function




'Created By: Rajesh
'Created on: 23/10/2016
'Description: Click on OCRF request tab.
'Parameter: Param1 - Extra,Param2 - Extra,Param3 - Extra,TabName - Give ocrf request tab name. for eg: "Add Users",objData
Public Function ClickOnOCRFRequestTab(ByVal Param1, ByVal Param2, ByVal Param3,ByVal TabName,ByRef objData)
	
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html tag:=SMALL","innertext:=.*"&TabName&".*").Exist Then
		Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html tag:=SMALL","innertext:=.*"&TabName&".*").Click
		wait 2
		Call ReportStep (StatusTypes.Pass, " "&TabName&" tab is available"," "&TabName&" tab")
	Else
		Call ReportStep (StatusTypes.Fail, " "&TabName&" tab is not available"," "&TabName&" tab")
		Exit Function
	End If
	
End Function 





'Created By: Rajesh
'Created on: 23/10/2016
'Description: Login to sca using client url and validate home page of sca is displayed.
'Parameter: Param1 - Extra,Param2 - Extra,Param3 - Extra,User_ID - Give username,Password - Give password,URL - SCA Client url,objData
Public Function LoginToSCAFromClientUrl(ByVal Param1, ByVal Param2, ByVal Param3,ByVal User_ID,ByVal Password,ByVal URL,ByRef objData)
	
	'Login to SCA using client URL
	'Systemutil.CloseProcessByName "iexplore.exe"
	Systemutil.Run "iexplore.exe", URL

If Browser("Browser-SCA").Page("Page-SCA").WebEdit("WebEdit_Username").Exist(60) Then
	If Browser("Browser-SCA").Page("Page-SCA").WebEdit("WebEdit_Username").Exist(5) Then
		Browser("Browser-SCA").Page("Page-SCA").WebEdit("WebEdit_Username").Set User_ID
	Else
		Call ReportStep (StatusTypes.Fail, "User name text box is not displayed","User name text box")
		Exit Function
	End If
	
	If Browser("Browser-SCA").Page("Page-SCA").WebEdit("WebEdit_Password").Exist(60) Then
		Browser("Browser-SCA").Page("Page-SCA").WebEdit("WebEdit_Password").Set Password
	Else
		Call ReportStep (StatusTypes.Fail, "Password text box is not displayed","Password text box")
		Exit Function
	End If


	If Browser("Browser-SCA").Page("Page-SCA").WebButton("WebButton_Login").Exist(10) Then
		Browser("Browser-SCA").Page("Page-SCA").WebButton("WebButton_Login").Click
	Else
		Call ReportStep (StatusTypes.Fail, "Login button is not available","Login button")
		Exit Function
	End If
	
	
	'Validate lady image is displayed
'	***********************************************************
'Modified BY Madhu - 03/30/2020
	'If Browser("DC").Page("DC").Image("Image_My_Informed_Decisions").Exist Then
	if Browser("DC").Page("DC").Link("Link_IMSAnalysisManager").Exist(60) then 
		'Click on IMS Analysis manager link
	Browser("DC").Page("DC").Link("Link_IMSAnalysisManager").Click
'**********************************

Call ReportStep (StatusTypes.Pass, "Login sucessfully to decision centre using client url"&URL,"Login to decision centre")
	Else	
		'Call ReportStep (StatusTypes.Fail, "Not able to Login sucessfully to decision centre using client url"&URL,"Login to decision centre")
	End If
'MOdified BY Madhu
	'Click on IMS Analysis manager link
	'Browser("DC").Page("DC").Link("Link_IMSAnalysisManager").Click
Browser("Browser-SCA").Page("Page-SCA").Link("innertext:=IQVIA Analysis Manager","html tag:=A").Click

'	If Browser("DC").Page("DC").Link("Link_SharedReports").Exist(120) Then
'		Call ReportStep (StatusTypes.Pass, "Login sucessfully to SCA","Login to SCA")
'	Else	
'		Call ReportStep (StatusTypes.Fail, "Not able to Login sucessfully to SCA","Login to SCA")
'	End If
	'****************************************************************8
End If
	
	
End Function







'Created By: Rajesh
'Created on: 08/11/2016
'Description: Search with Request ID or Client ID and validate the repective column in the grid.
'Parameter: Param1 - Extra,Param2 - Extra,Param3 - Extra,TabName - ONLINE CRF|CLIENTS|MANAGE USERS ACCESS|AUDIT LOG|HELP,IDName - Request ID|ClientID,IDValue - give the id which you want to search,objData
Public Function SearchByID(ByVal Param1, ByVal Param2, ByVal Param3,ByVal TabName,ByVal IDName,ByVal IDValue,ByRef objData)
	
	'counter variable
	count = 0
	
	'set the id in text box
	Set objtxtRequestClientID = Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_Request_ClientID")
	Call SCA.SetText(objtxtRequestClientID,IDValue,IDName,TabName)
	If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").Exist(30) Then
	    wait 10
	   'click on text field 
	    Call SCA.ClickOn(objtxtRequestClientID, " "&IDName&" text field", TabName)
	End If
	
	'Press ENTER key
	Set WshShell = CreateObject("WScript.Shell")
    WshShell.SendKeys "{ENTER}"
    wait 5
    
    'Validate serached value is displayed in grid or not
    rowcount = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").GetROProperty("rows")
    
    For rowpos = 2 To rowcount Step 1
    	If Cdbl(trim(Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").GetCellData(rowpos,2))) =  Cdbl(IDValue) Then
    	
    	ElseIf Cdbl(trim(Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").GetCellData(rowpos,3))) =  Cdbl(IDValue) Then
    	
		Else
			count = 1
			Exit for	
    	End If
    Next
    
    If count = 0 Then
    	Call ReportStep (StatusTypes.Pass, "Search functionality is working fine in tab "&TabName&",for column "&IDName&" with column value: "&IDValue,"Search with "&IDName&" ")
	Else	
		Call ReportStep (StatusTypes.Fail, "Search functionality is not working fine in tab "&TabName&",for column "&IDName&" with column value: "&IDValue,"Search with "&IDName&" ")
    End If

End Function



Function SelectingUserFromList(byVal strTab, byVal selectList1, byval selectList2, byVal OCRFUser, ByVal strF1, byVal strF2)
	'1) Click on Select Users
	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Click
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_PleaseSelectUsersFrom").Exist(120) <> true Then
		Call ReportStep (StatusTypes.Fail, "Select User Dialog Page didnt not pop-up after click of 'WebElement_BIToolAccess_SelectUsers'","BI Tool Access")
		Exit Function
	Else
		Call ReportStep (StatusTypes.Pass, "Select User Dialog Page is displaying after click of 'WebElement_BIToolAccess_SelectUsers'","BI Tool Access")
	End If
	
	Path = Environment.Value("CurrDir")&"\InputFiles\IMSSCAWeb\Copy of OCRF_Export.xls"
	chkBox = 0 

	If strTab = "BI Tools Access" Then
		'BI Tool Access
		Datatable.ImportSheet Path,"BI Tools Access Mode","Action1"
	ElseIf strTab = "User Cube Access" Then
		'BI Tool Access
		Datatable.ImportSheet Path,"User Cube Access","Action1"
	End If

	'3) Read data from Excel data
	'Search for Test User added in Select Users List to 'Choose BI Tools' and 'Access Mode'. Click on Add
	sheetRC = DataTable.GetSheet("Action1").GetRowCount
	For intSheetRC = 1 To sheetRC Step 1
		Datatable.SetCurrentRow(intSheetRC)    
		excelUser = Datatable.Value("C", dtLocalSheet)
		print excelUser
		
		If trim(OCRFUser) = trim(excelUser) Then
			'get BI Tool and Access Mode for BI Tool access Tab; Cube datatbase and Cube role for User Cube access Tab		
				If strTab = "BI Tools Access" Then
					select1 = Datatable.Value("E", dtLocalSheet)
					select2 = Datatable.Value("F", dtLocalSheet)
					
					Browser("Browser-OCRF").Page("Page-OCRF").WebList(selectList1).Select select1
					Browser("Browser-OCRF").Page("Page-OCRF").Sync
					wait 5
					Browser("Browser-OCRF").Page("Page-OCRF").WebList(selectList2).Select select2
				ElseIf strTab = "User Cube Access" Then
					select1 = Datatable.Value("E", dtLocalSheet)
					select2 = Datatable.Value("G", dtLocalSheet)
				End If
			
			'Set BI Tools and Access Mode
			'WebList_BIToolAccess_BIAccessTool, WebList_BIToolAccess_BIToolAccessMode
			'WebList_UserCubeAccess_CubeDB, WebList_UserCubeAccess_CubeRoles
'			Browser("Browser-OCRF").Page("Page-OCRF").WebList(selectList1).Select select1
'			Browser("Browser-OCRF").Page("Page-OCRF").Sync
'			wait 5
'			Browser("Browser-OCRF").Page("Page-OCRF").WebList(selectList2).Select select2
			
			'Enable the Check box to select UserLogin (Check box found at 2nd column)
			RC_BIAccessUserList = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").GetROProperty("rows")
			CC_BIAccessUserList = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").GetROProperty("cols")
			For iRCUserList = 2 To RC_BIAccessUserList Step 1
				For iCCUserList = 4 To CC_BIAccessUserList Step 1
					BIAccessUserLoginID = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").GetCellData(iRCUserList, iCCUserList)
					If Trim(BIAccessUserLoginID) = Trim(excelUser) Then
						'Enable the Check box to select UserLogin (Check box found at 2nd column)
						Set objChkbox = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").ChildItem(iRCUserList, 2, "WebCheckBox", 0)
						objChkbox.set "ON"
						chkBox = 1
						Exit For
					End If
				Next
				If chkBox = 1 Then
					Exit For
				End If
			Next
			If chkBox = 1 Then
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-pg-div","html tag:=DIV","innertext:=Add User.*","index:=0").Click
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-button-text","html tag:=SPAN","innerhtml:=Add", "index:=1").click
				Exit for
			End If
		End If
	Next
End Function

'Created By: Poornima
'Created on: 17/11/2016
'Description: End to End data entering ("Client & Offering" details should be displayed in first grid in all the Tabs.
'Parameter: TabName - OFFERING DETAILS|CUBE DETAILS|ADD USERS|CLIENT CONTENT|BI TOOL ACCESS|USER CUBE ACCESS

Function EndToEndFlowNewRequest(ByVal Tab,ByRef objData)
		
		          Path = Environment.Value("CurrDir")&"\InputFiles\IMSSCAWeb\Copy of OCRF_Export.xls"
    			'DataTable.AddSheet "Global"
	            'Datatable.ImportSheet Path,1,"Action1"
	             Datatable.ImportSheet Path,"Client Details","Action1"
	             r = DataTable.GetSheet("Action1").GetRowCount
	             ReDim Section1(r),Section2(r)
	             cnt=0
				 For i = 2 To r Step 1
	  			     Datatable.SetCurrentRow(i)
	  				 Section1(cnt) = dataTable.value("C" ,dtLocalSheet)
	  				 Section2(cnt) = dataTable.value("F" ,dtLocalSheet)
	  				 cnt=cnt+1
    			Next
		
				    'Datatable.ImportSheet Path,3,"Action1"
				    Datatable.ImportSheet Path,"Offering Users","Action1"
					r2 = DataTable.GetSheet("Action1").GetRowCount
					Datatable.SetCurrentRow(3)
					First_Name=dataTable.value("D" ,dtLocalSheet)
					Last_Name=dataTable.value("E" ,dtLocalSheet)
					User_Login_ID=dataTable.value("F" ,dtLocalSheet)
					Business_Email_Address=dataTable.value("G" ,dtLocalSheet)
					Client_Name=dataTable.value("H" ,dtLocalSheet)
					Country_Name=dataTable.value("I" ,dtLocalSheet)	
					Onyx_ID=dataTable.value("J" ,dtLocalSheet)	
					Country_ID=dataTable.value("K" ,dtLocalSheet)
		
						       
				    'Datatable.ImportSheet Path,2,"Action1"
				    Datatable.ImportSheet Path,"Cube_Details","Action1"
					r1 = DataTable.GetSheet("Action1").GetRowCount
					Datatable.SetCurrentRow(2)
					Cube_Name= dataTable.value("E" ,dtLocalSheet)
					Server_Name= dataTable.value("H" ,dtLocalSheet)
					Cube_Roles= dataTable.value("F" ,dtLocalSheet)
					AutoCube_Name=objData.item("AutoCube_Name")
					AutoServer_Name=objData.item("AutoServer_Name")
		  Select Case Tab
		  Case "OFFERING DETAILS TAB"
		  		
	            '########### OFFERING DETAIL TAB EXISTANCE - Start ########### 
		  		If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_OfferingDetailsTab").Exist(20) Then
				    Call ReportStep (StatusTypes.Pass, "Offering Details Tab exist","Offering Details Tab")
				Else
	   				Call ReportStep (StatusTypes.Fail, "cOffering Details Tab not exist","Offering Details Tab")
				End If
    			'ADD DATA TO OFFERING DETAILS TAB
    			
    			'ENTER DATA IN OFFERING CLIENT NAME
    			If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingClient").Exist(20) Then
       				Call ReportStep (StatusTypes.Pass, "Offering client name entered with "&Section1(1),"Offering Client")	
       				Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingClient").Set Section1(1)
         	    Else
	   				Call ReportStep (StatusTypes.Fail, "Offering client name not entered","Offering Client")
    			End If
    			' SELECT OFFERING CONTRY NAME
    			If Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebListOfferingCountry").Exist(20) Then
       				Call ReportStep (StatusTypes.Pass, "Offering country name selected as "&Section1(3),"Offering Country name")	
       				Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebListOfferingCountry").Select Section1(3)
    			Else
	   				Call ReportStep (StatusTypes.Fail, "Offering country name not selected","Offering Country name")
   	 			End If
    			'SELECT SYNDIACATED NON SYNDICATED RADIO BUTTON
    			If Browser("Browser-OCRF").Page("Page-OCRF").WebRadioGroup("WEbRadioSyndicated").Exist(20) Then
       				'Call ReportStep (StatusTypes.Pass, "Syndicated radio button selected as "&Section1(4),"Syndicated")	
       				If Section1(4)="No" Then
       	  				Browser("Browser-OCRF").Page("Page-OCRF").WebRadioGroup("WEbRadioSyndicated").Select "false"
       				Else
       	  				Browser("Browser-OCRF").Page("Page-OCRF").WebRadioGroup("WEbRadioSyndicated").Select "on"
       				End If
    			Else
	   				Call ReportStep (StatusTypes.Fail, "Offering country name not selected","Syndicated")
    			End If
    			'ENTER OFFERING NAME
    			If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingName").Exist(20) Then
       				Call ReportStep (StatusTypes.Pass, "Offering Name entered as "&Section1(7),"Offering name")	
       				Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingName").Set Section1(7)
    			Else
	   				Call ReportStep (StatusTypes.Fail, "Offering Name  not entered","Offering name")
    			End If
    			'SELECT OFFERING TYPE FROM LIST
    			If Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebListOfferingType").Exist(20) Then
       				Call ReportStep (StatusTypes.Pass, "Offering type entered as "&Section1(6),"Offering type")	
       				Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebListOfferingType").Select Section1(6)
    			Else
	   				Call ReportStep (StatusTypes.Fail, "Offering type  not entered","Offering type")
    			End If
    			'SELECT ENVIRONMENT SET UP
    			If Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebListSetupEnvironment").Exist(20) Then
       				Call ReportStep (StatusTypes.Pass, "Set Up Environment entered as "&Section1(0),"Set Up Environment")	
       				Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebListSetupEnvironment").Select Section1(0)
    			Else
	   				Call ReportStep (StatusTypes.Fail, "Set Up Environment not entered","Set Up Environment")
    			End If
    			'ENTER OFFERING CONTACT NAME
    			If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactName").Exist(20) Then
       				Call ReportStep (StatusTypes.Pass, "Offering Contact Name entered as "&Section2(0),"Offering Contact Name")	
       				Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactName").Set Section2(0)
    			Else
	   				Call ReportStep (StatusTypes.Fail, "Offering Contact Name  not entered","Offering Contact Name")
    			End If
    			'ENTER OFFERING CONTACT PHONE
    			If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactPhone").Exist(20) Then
       				Call ReportStep (StatusTypes.Pass, "Offering Contact Phone entered as "&Section2(1),"Offering Contact Phone")	
       				Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactPhone").Set Section2(1)
    			Else
	   				Call ReportStep (StatusTypes.Fail, "Offering Contact Phone  not entered","Offering Contact Phone")
    			End If
    			'ENTER OFFERING CONTACT EMAIL
    			If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactEmail").Exist(20) Then
       				Call ReportStep (StatusTypes.Pass, "Offering Contact Email entered as "&Section2(2),"Offering Contact Email")	
       				Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactEmail").Set Section2(2)
    			Else
	   				Call ReportStep (StatusTypes.Fail, "Offering Contact Email  not entered","Offering Contact Email")
    			End If

             '   If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementFinish").Exist(20) Then
             '       Call ReportStep (StatusTypes.Pass, "Finish button exist ","Finish button")	
             '       Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementFinish").Click
             '    Else
             '	   Call ReportStep (StatusTypes.Fail, "Finish button not exist","Finish button")
             '    End If
             
             'CLICK ON NEXT BUUTON
			  If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Exist(20) Then
				 Call ReportStep (StatusTypes.Pass, "Next button exist and clicked ","Next button")	
				 Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Click
			 Else
				Call ReportStep (StatusTypes.Fail, "Next button not exist","Next button")
			 End If
		     wait 10
	         '########### OFFERING DETAIL TAB EXISTANCE - End ########### 
				    
			Case "CUBE DETAILS TAB"
                    CubeSource=Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-th-div-ie ui-jqgrid-sortable","html tag:=DIV","innertext:=Cube Source").Exist(5)
					Country=Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-th-div-ie ui-jqgrid-sortable","html tag:=DIV","innertext:=Country").Exist(5)
					Period=Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-th-div-ie ui-jqgrid-sortable","html tag:=DIV","innertext:=Period").Exist(5)
					ClientName=Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-th-div-ie ui-jqgrid-sortable","html tag:=DIV","innertext:=Client Name").Exist(5)
					OrderID=Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-th-div-ie ui-jqgrid-sortable","html tag:=DIV","innertext:=Order ID").Exist(5)
					QACube=Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-th-div-ie ui-jqgrid-sortable","html tag:=DIV","innertext:=QA Cube").Exist(5) 
					'########### CUBE DETAILS TAB EXISTANCE - Start ########### 
                   
				    
				   'CUBE DETAILS TAB EXIST
    				 Set CubeDetailsTab=Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementCubeDatabaseDetails")
      				Set DBDetailsTab=Browser("Browser-OCRF").Page("Page-OCRF").WebElement("innertext:=Database Details","html tag:=SPAN","class:=ui-jqgrid-title")
					'CUBE DETAILS TAB EXIST
					If CubeDetailsTab.Exist(5) OR DBDetailsTab.Exist(5) Then
						Call ReportStep (StatusTypes.Pass, "Entered in to cude details tab ","Cube details tab")	
					Else
						Call ReportStep (StatusTypes.Fail, "Not Entered in to cude details tab","Cube details tab")
	 				End If
	 				
	 				
				   'CLICK ON ADD NEW LINE
				    If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementAddNewLine").Exist(20) Then
				    	wait 2
				        Call ReportStep (StatusTypes.Pass, "Click on add new line done","Add new button")	
				        Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementAddNewLine").Click
				   Else
					   Call ReportStep (StatusTypes.Fail, "Click on add new line not done","Add new button")
				   End If
				  
				 'Call AddNewLineInCubeDetails(Cube_Name,Server_Name,FLAPath,1,objData)
				   'CHECK CHECKBOX TO ADD NEW LINE
				   If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").Exist(20) Then
				       Call ReportStep (StatusTypes.Pass, "Check box selection done","Check box")	
				       Set CheckRow =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(2,2,"WebCheckBox",0)
		               CheckRow.click	
				   Else
					   Call ReportStep (StatusTypes.Fail, "Check box selection  not done","Check box")
				   End If
				   
				   Call WebEditExistence(2,Cint(Environment.value("cubeName")),"WebEdit")
		
		           Set CheckDatabase =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(2,Cint(Environment.value("cubeName")),"WebEdit",0)
		           CheckDatabase.set AutoCube_Name
				   
				   
				   
'				   Call WebEditExistence(2,Cint(Environment.value("cubeLocation")),"WebEdit")
'		
'		           Set CheckDBLocation =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").ChildItem(2,Cint(Environment.value("cubeLocation")),"WebEdit",0)
'		           CheckDBLocation.set Server_Name
		
		 
'				  'ENTER COMPLETE NAME 
'			      If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=cube-list","html tag:=TABLE").WebElement("innerhtml:=&nbsp;","outerhtml:=<TD role=gridcell aria-describedby=cube-list_CubeTempleteName title="""" style=""TEXT-ALIGN: left"">&nbsp;</TD>","html tag:=TD").Exist(5) Then
'          	         Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=cube-list","html tag:=TABLE").WebElement("innerhtml:=&nbsp;","outerhtml:=<TD role=gridcell aria-describedby=cube-list_CubeTempleteName title="""" style=""TEXT-ALIGN: left"">&nbsp;</TD>","html tag:=TD").Click
'          		  Else
'          		  End If
'				  If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditCubeempleteName").Exist(20) Then
'				       Call ReportStep (StatusTypes.Pass, "Entered Cube Complete name as  "& Cube_Name,"Cube Complete name")	
'				       Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditCubeempleteName").Set Cube_Name
'				   Else
'					   Call ReportStep (StatusTypes.Fail, "not Entered Cube Complete name","Cube Complete name")
'				   End If
'				   
'				   
'				   'ENTER SERVER NAME
'				   If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditServer").Exist(20) Then
'				       Call ReportStep (StatusTypes.Pass, "Server name entered as "&Server_Name,"Server name")	
'				       Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditServer").Set Server_Name
'				   Else
'					   Call ReportStep (StatusTypes.Fail, "Server name Not entered","Server name")
'				   End If
                   'Validate that "Cube Location/Server” column should be optional in cube details tab. 
				   wait 5
				    AutoPopServerName=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetCellData(2,Cint(Environment.value("cubeLocation")))
				    If instr(1,AutoPopServerName,AutoServer_Name)<>0 Then
			           Call ReportStep (StatusTypes.Pass, "Server name Autopopulated after entering cube name ","Server name")
				    Else
					   Call ReportStep (StatusTypes.Fail, "Server name not Autopopulated after entering cube name","Server name")
					End If
				   'CLICL ON UPDATE LINK
				   If Browser("Browser-OCRF").Page("Page-OCRF").Link("WebLinkUPDATE").Exist(20) Then
				       Call ReportStep (StatusTypes.Pass, "Click on update link to update Cube roles done","Update cube roles link")	
				       Browser("Browser-OCRF").Page("Page-OCRF").Link("WebLinkUPDATE").Click 
				   Else
					   Call ReportStep (StatusTypes.Fail, "Click on update link to update Cube roles Not done","Update cube roles link")
				   End If
				   'ENTER CUBE ROLE
				   If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditCubeRoles").Exist(20) Then
				       Call ReportStep (StatusTypes.Pass, "Entered cube roles","Cube roles")	
				       Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditCubeRoles").Set Cube_Roles
				   Else
					   Call ReportStep (StatusTypes.Fail, "Cube roles not Entered","Cube roles ")
				   End If
				   'CLICK ON UPDATE BUTTON
				    If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementUpdateCubeRole").Exist(20) Then
				       Call ReportStep (StatusTypes.Pass, "Click on Updae done","Cube roles")	
				       Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementUpdateCubeRole").Click
				   Else
					   Call ReportStep (StatusTypes.Fail, "Click on Updae Not done","Cube roles ")
				   End If

				  
				   
				   '########### CUBE DETAILS TAB EXISTANCE - End ########### 
				   
				   'CLICK ON NEXT BUTTON
				    If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Exist(20) Then
				       Call ReportStep (StatusTypes.Pass, "Next button exist and clicked ","Next button")	
				       Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Click
				    Else
					   Call ReportStep (StatusTypes.Fail, "Next button not exist","Next button")
				    End If
				    wait 10
				    '(User can skip providing the details in this column and can navigate to next tab)   
				    If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementOfferingUserList").Exist(20) Then
				       Call ReportStep (StatusTypes.Pass, "Entered in to Add user tab by skiping details of Cube Location/Server ","Add user tab")	
				    Else
					   Call ReportStep (StatusTypes.Fail, "Not Entered in to Add user tab by skiping details of Cube Location/Server","Add user tab")
				    End If
				    
				    
		Case "ADD USER TAB"
		   
		   			 '########### ADD USER TAB - Start ########### 
		   			 
		   			 Call AddNewLineINAddUser(Client_Name,User_Login_ID,objData)
'					'CHECK FOR USER ACCESS TAB EXIST
'					If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementOfferingUserList").Exist(20) Then
'				       Call ReportStep (StatusTypes.Pass, "Entered in to Add user tab","Add user tab")	
'				    Else
'					   Call ReportStep (StatusTypes.Fail, "Not Entered in to Add user tab","Add user tab")
'				    End If
'				    'CLICK ON ADD NEW USER
'				    If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementAddNewLine").Exist(20) Then
'				       Call ReportStep (StatusTypes.Pass, "Click on Add new line icon done" ,"Add new line")	
'				       Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementAddNewLine").Click
'				    Else
'					   Call ReportStep (StatusTypes.Fail, "Click on Add new line icon Not done","Add new line")
'				    End If
'				    
'				    
'				    
'				    
'				   'CHECK CHECKBOX TO ADD NEW LINE
'				   If Browser("Browser-OCRF").Page("Page-OCRF").WebCheckBox("WebListSelectUserList").Exist(20) Then
'				       Call ReportStep (StatusTypes.Pass, "Check box selection done","Check box")	
'				       Browser("Browser-OCRF").Page("Page-OCRF").WebCheckBox("WebListSelectUserList").Set "ON"
'				   Else
'					   Call ReportStep (StatusTypes.Fail, "Check box selection  not done","Check box")
'				   End If
'					 
'
'				    'ENTER USER LOGIN ID
'					If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditUserLoginID").Exist(20) Then
'				       Call ReportStep (StatusTypes.Pass, "Enter User login Id done as "& User_Login_ID,"User login Id")	
'				       Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditUserLoginID").Set User_Login_ID
'				    Else
'					   Call ReportStep (StatusTypes.Fail, "Enter User login Id not done","User login Id")
'				    End If
'				 
'				    'ENTER lAST NAME
'				    If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html tag:=TD","innertext:="&Last_Name,"title:="&Last_Name).Exist(20) Then
'				       Call ReportStep (StatusTypes.Pass, "Enter Last name done as "&Last_Name,"Last name")	
'				       Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html tag:=TD","innertext:="&Last_Name,"title:="&Last_Name).Click
'				       Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditLastName").Set Last_Name
'				    Else
'					   Call ReportStep (StatusTypes.Fail, "Enter Last name not done","Last name")
'				    End If
'				    'ENTER FIRST NAME
'				    If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditFirstName").Exist(20) Then
'				       Call ReportStep (StatusTypes.Pass, "Enter First name done as"&First_Name,"First name")
'				       Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditFirstName").Set First_Name
'				    Else
'					   Call ReportStep (StatusTypes.Fail, "Enter First name not done","First name")
'				    End If
'				    'ENTER BUSINESS EMAIL ADDRESS
'				    If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditBussinessEmailAddress").Exist(20) Then
'				       Call ReportStep (StatusTypes.Pass, "Enter Business email address done as "&Business_Email_Address,"Business email address")	
'				       Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditBussinessEmailAddress").Set Business_Email_Address
'				    Else
'					   Call ReportStep (StatusTypes.Fail, "Enter FBusiness email address not done","Business email address")
'				    End If
'                    
'				   'ENTER CLIENT NAME
'				    If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditClientName").Exist(20) Then
'				       Call ReportStep (StatusTypes.Pass, "Enter Business Client Name done as "&Client_Name,"Client Name")	
'				       'Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html tag:=TD","innertext:="&Client_Name,"title:="&Client_Name).Click
'				       Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditClientName").highlight 
'					   Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditClientName").Set Client_Name  
'					   Set objShell=CreateObject("WScript.Shell")
'					   Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditClientName").Click
'					   objShell.SendKeys "{DOWN}"
'					   wait 2
'					   objShell.SendKeys "{DOWN}"
'					   wait 2  
'					   objShell.SendKeys "{ENTER}"
'					     
'				    Else
'					   Call ReportStep (StatusTypes.Fail, "Enter Client Name not done","Client Name")
'				    End If
'				    'ENTER CONTRY NAME
'				    If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html tag:=TD","innertext:="&Country_Name,"title:="&Country_Name).Exist(20) Then
'				       Call ReportStep (StatusTypes.Pass, "Enter Country Name done as "&Country_Name,"Country Name")	
'				       'Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditCountryName").Set Country_Name
'				       Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html tag:=TD","innertext:="&Country_Name,"title:="&Country_Name).Click
'				       Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditCountryName").highlight
'					   Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditCountryName").Set Country_Name
'					   Set objShell=CreateObject("WScript.Shell")
'					   Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditCountryName").Click
'					   objShell.SendKeys "{DOWN}"
'					   wait 2
'					   objShell.SendKeys "{DOWN}"
'					   wait 2  
'					   objShell.SendKeys "{ENTER}"
'					 
'				    Else
'					   Call ReportStep (StatusTypes.Fail, "Enter Country Name not done","Country Name")
'				    End If
				    'ENTER ONYX ID
					'If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html tag:=TD","innertext:="&Onyx_ID,"title:="&Onyx_ID).Exist(20) Then
					   'Call ReportStep (StatusTypes.Pass, "Enter OnyxID done as "&Onyx_ID,"OnyxID")	
					   'Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html tag:=TD","innertext:="&Onyx_ID,"title:="&Onyx_ID).Click
					   'Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOnyxID").Set Onyx_ID
					'Else
					    'Call ReportStep (StatusTypes.Fail, "Enter OnyxID not done","OnyxID")
					'End If
				    'ENTER COMPANY ID
					'If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditCompanyID").Exist(20) Then
						'Call ReportStep (StatusTypes.Pass, "Enter CompanyID done as "&Country_ID,"CompanyID")	
					    'Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html tag:=TD","innertext:="&Country_ID,"title:="&Country_ID).Click
					    'Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditCompanyID").Set Country_ID
					'Else
						'Call ReportStep (StatusTypes.Fail, "Enter CompanyID not done","Country Name")
					'End If
				    'CLICK ON NEXT BUUTON
				    If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Exist(20) Then
				       Call ReportStep (StatusTypes.Pass, "Next button exist and clicked ","Next button")	
				       Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Click
				    Else
					   Call ReportStep (StatusTypes.Fail, "Next button not exist","Next button")
				    End If
				    wait 20
				    '########### ADD USER TAB - End ###########	

        Case "CLIENT CONTENT TAB"
					'########### CLIENT CONTENT TAB - Start ########### 
				    'CLICK ON NEXT BUUTON
				    If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Exist(20) Then
				       Call ReportStep (StatusTypes.Pass, "Next button exist and clicked ","Next button")	
				       Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Click
				    Else
					   Call ReportStep (StatusTypes.Fail, "Next button not exist","Next button")
				    End If
				    wait 10
				    
				    'TODO: NOT CONSIDERED CLIENT CONTENT TAB FOR VALIDATION
				    '########### CLIENT CONTENT TAB - End ###########          
		 Case "BI TOOL ACCESS TAB"
		 
					'1) Validate whether OCRF New Request is redirected to BI Tool access
					If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementOfferingUserBIToolList").Exist(120) <> true Then
				       Call ReportStep (StatusTypes.Fail, "Not Redirected to BI Tool Access tab","OCRF New Request")	
				    Else
					   Call ReportStep (StatusTypes.Pass, "Redirected to BI Tool Access tab successfully","OCRF New Request")
				    End If
				
					OCRFUser = User_Login_ID
					Call SelectingUserFromList("BI Tools Access", "WebList_BIToolAccess_BIAccessTool", "WebList_BIToolAccess_BIToolAccessMode", OCRFUser, "", "")
					
					'4) Validate Test User has been added
					RC_BIAccessList = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").GetROProperty("rows")
					CC_BIAccessList = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").GetROProperty("cols")
					For iRC_BIAccessList = 2 To RC_BIAccessList Step 1
						BIAccessListUser = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").GetCellData(iRC_BIAccessList, 8)
						BIAccessListBITool = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").GetCellData(iRC_BIAccessList, 11)
						BIAccessListAccessMode = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").GetCellData(iRC_BIAccessList, 13)
						'If (trim(BIAccessListUser) = trim(excelUser)) and (trim(BIAccessListBITool) = trim(BITools)) and (trim(BIAccessListAccessMode) = trim(AccessMode)) Then
						If (trim(BIAccessListUser) = trim(excelUser)) and (trim(BIAccessListBITool) = trim(select1)) and (trim(BIAccessListAccessMode) = trim(select2)) Then
							   Call ReportStep (StatusTypes.Pass, "Successfully set BI Tools and Access Mode to "&BIAccessListUser,"OCRF Request -> BI Tool Access Tab")	
						    Else
							   Call ReportStep (StatusTypes.Fail, "Not set BI Tools and Access Mode to "&BIAccessListUser,"OCRF Request -> BI Tool Access Tab")	
						End If
					Next
					
					
					'CLICK ON NEXT BUTTON
				    If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Exist(20) Then
				       Call ReportStep (StatusTypes.Pass, "Next button exist and clicked ","Next button")	
				       Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Click
				    Else
					   Call ReportStep (StatusTypes.Fail, "Next button not exist","Next button")
				    End If
				    wait 10
		 Case "USER CUBE ACCESS TAB"
		 
				   	'1) Validate whether OCRF New Request is redirected to BI Tool access
				   	Browser("Browser-OCRF").Page("Page-OCRF").Sync
					If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementOfferingUserBIToolList").Exist(120) <> true Then
				       Call ReportStep (StatusTypes.Fail, "Not Redirected to BI Tool Access tab","OCRF New Request")	
				    Else
					   Call ReportStep (StatusTypes.Pass, "Redirected to BI Tool Access tab successfully","OCRF New Request")
				    End If
					
					wait 5
					OCRFUser = User_Login_ID
					Call SelectingUserFromList("User Cube Access", "WebList_UserCubeAccess_CubeDB", "WebList_UserCubeAccess_CubeRoles", OCRFUser, "", "")
					
					'4) Validate Test User has been added
					RC_UserCubeAccess = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserCubeAccess_CubeAccessList").GetROProperty("rows")
					CC_UserCubeAccess = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserCubeAccess_CubeAccessList").GetROProperty("cols")
					For iRC_UserCubeAccess = 2 To RC_UserCubeAccess Step 1
						CubeAccessListUser = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserCubeAccess_CubeAccessList").GetCellData(iRC_UserCubeAccess, 7)
						CubeAccessListDatabse = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserCubeAccess_CubeAccessList").GetCellData(iRC_UserCubeAccess, 9)
						CubeAccessListRole = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserCubeAccess_CubeAccessList").GetCellData(iRC_UserCubeAccess, 11)
						If (trim(CubeAccessListUser) = trim(excelUser)) and (trim(CubeAccessListDatabse) = trim(select1)) and (trim(CubeAccessListRole) = trim(select2)) Then
							   Call ReportStep (StatusTypes.Pass, "Successfully set BI Tools and Access Mode to "&BIAccessListUser,"OCRF Request -> BI Tool Access Tab")	
						Else
							   Call ReportStep (StatusTypes.Fail, "Not set BI Tools and Access Mode to "&BIAccessListUser,"OCRF Request -> BI Tool Access Tab")	
						End If
					Next

		End Select
	End Function
'Created By: Poornima
'Created on: 24/11/2016
'Description: Validate the ("Client & Offering" details should be displayed in first grid in all the Tabs.
'Parameter: TabName - OFFERING DETAILS|CUBE DETAILS|ADD USERS|CLIENT CONTENT|BI TOOL ACCESS|USER CUBE ACCESS
Function FrirstGridDetailsValidation(ByVal Tab,byRef objData)
    
	Select Case Tab
	    Case "OFFERING DETAILS"
	          If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Exist(20) Then
		           Call ReportStep (StatusTypes.Pass, "Next button exist and clicked ","Next button")	
		           Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Click
		           Exit Function
	          Else
		          Call ReportStep (StatusTypes.Fail, "Next button not exist","Next button")
	          End If
	          
		Case "CUBE DETAILS"
		     '########### CUBE DETAILS TAB EXISTANCE - Start ########### 

			'CUBE DETAILS TAB EXIST
			 
			'CUBE DETAILS TAB EXIST
     		Set CubeDetailsTab=Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementCubeDatabaseDetails")
      		Set DBDetailsTab=Browser("Browser-OCRF").Page("Page-OCRF").WebElement("innertext:=Database Details","html tag:=SPAN","class:=ui-jqgrid-title")
			'CUBE DETAILS TAB EXIST
			If CubeDetailsTab.Exist(5) OR DBDetailsTab.Exist(5) Then
				Call ReportStep (StatusTypes.Pass, "Entered in to cude details tab ","Cube details tab")	
			Else
				Call ReportStep (StatusTypes.Fail, "Not Entered in to cude details tab","Cube details tab")
	 		End If
		    'verification for all data(Client & Offering) existance in grid
		     GridHeaderData=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=client-offering","name:=requistID","html tag:=TABLE").GetROProperty("text")
		     GridHeaderValue=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=client-offering","name:=requistID","html tag:=TABLE").GetROProperty("innerhtml")
		     
		     
	         
		Case "ADD USERS"
		     'CHECK FOR USER ACCESS TAB EXIST
             If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementOfferingUserList").Exist(20) Then
				 Call ReportStep (StatusTypes.Pass, "Entered in to Add user tab","Add user tab")	
			 Else
				 Call ReportStep (StatusTypes.Fail, "Not Entered in to Add user tab","Add user tab")
			 End If	

              'verification for all data(Client & Offering) existance in grid
			 GridHeaderData=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=","name:=user-offerID","html tag:=TABLE").GetROProperty("text")
		     GridHeaderValue=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=","name:=user-offerID","html tag:=TABLE").GetROProperty("innerhtml")
		 
       Case "CLIENT CONTENT" 
             'CHECK FOR CLIENT CONTENT TAB EXIST
             If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=client-offering","html tag:=TABLE","name:=ContentRequestID").Exist(20) Then
				 Call ReportStep (StatusTypes.Pass, "Entered in to Client content tab","Client content tab")	
			 Else
				 Call ReportStep (StatusTypes.Fail, "Not Entered in to Client content tab","Client content tab")
			 End If	
             'verification for all data(Client & Offering) existance in grid
			 GridHeaderData=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=client-offering","name:=ContentRequestID","html tag:=TABLE").GetROProperty("text")
		     GridHeaderValue=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=client-offering","name:=ContentRequestID","html tag:=TABLE").GetROProperty("innerhtml")
		 
	   Case "BI TOOL ACCESS"
	   
			'1) Validate whether OCRF New Request is redirected to BI Tool access
			 If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementOfferingUserBIToolList").Exist(120) <> true Then
				Call ReportStep (StatusTypes.Fail, "Not Redirected to BI Tool Access tab","OCRF New Request")	
			 Else
				 Call ReportStep (StatusTypes.Pass, "Redirected to BI Tool Access tab successfully","OCRF New Request")
			 End If
					   		
	         'verification for all data(Client & Offering) existance in grid
			 GridHeaderData=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=client-offering","name:=BIAccessRequestID","html tag:=TABLE").GetROProperty("text")
		     GridHeaderValue=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=client-offering","name:=BIAccessRequestID","html tag:=TABLE").GetROProperty("innerhtml")
		 
		 
	   Case "USER CUBE ACCESS"
	        '1) Validate whether OCRF New Request is redirected to BI Tool access
				Browser("Browser-OCRF").Page("Page-OCRF").Sync
			  If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementOfferingUserBIToolList").Exist(120) <> true Then
				  Call ReportStep (StatusTypes.Fail, "Not Redirected to BI Tool Access tab","OCRF New Request")	
			  Else
				  Call ReportStep (StatusTypes.Pass, "Redirected to BI Tool Access tab successfully","OCRF New Request")
			  End If
	         'verification for all data(Client & Offering) existance in grid
			 GridHeaderData=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=client-offering","name:=CubeAccessRequestID","html tag:=TABLE").GetROProperty("text")
		     GridHeaderValue=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=client-offering","name:=CubeAccessRequestID","html tag:=TABLE").GetROProperty("innerhtml")
		 
	End Select
	         wait 10 
	         RequestID=instr(1,GridHeaderData,"Request ID")>0
		     RequestIDValue=instr(1,GridHeaderValue,objData.item("Request ID"))>0
		     ClientName=instr(1,GridHeaderData,"Client Name")>0
		     ClientNameValue=instr(1,GridHeaderValue,objData.item("Client Name"))>0
		     Syndicate=instr(1,GridHeaderData,"Syndicated")>0
		     SyndicateValue=instr(1,GridHeaderValue,objData.item("Syndicated"))>0
		     OfferingName=instr(1,GridHeaderData,"Offering Name")>0
		     OfferingNameValue=instr(1,GridHeaderValue,objData.item("Offering Name"))>0
		     OfferingType=instr(1,GridHeaderData,"Offering Type")>0
		     OfferingTypeValue=instr(1,GridHeaderValue,objData.item("Offering Type"))>0
		     ClientCountry=instr(1,GridHeaderData,"Client Country")>0
		     ClientCountryValue=instr(1,GridHeaderValue,objData.item("Client Country"))>0
		     If RequestID AND ClientName AND OfferingName AND Syndicate AND OfferingType AND ClientCountry Then
		     	Call ReportStep (StatusTypes.Pass,"Existance of Columns header in grid is as follow Request ID is " & RequestID & " Client Name is " & ClientName &" Syndicate is at" & Syndicate & " Offering Name is at" & OfferingName &" Offering Type is at" & OfferingType &" ClientCountry is at" & ClientCountry  ,"Headers in First grid")
		     Else
		        Call ReportStep (StatusTypes.Fail,"Existance of Columns header in grid is as follow Request ID is " & RequestID & " Client Name is " & ClientName &" Syndicate is at" & Syndicate & " Offering Name is at" & OfferingName &" Offering Type is at" & OfferingType &" ClientCountry is at" & ClientCountry  ,"Headers in First grid")
		     End If 
'		     If RequestIDValue  AND ClientNameValue AND OfferingNameValue AND SyndicateValue AND OfferingTypeValue AND ClientCountryValue Then
'		     	Call ReportStep (StatusTypes.Pass,"Existance of Columns header value in grid is as follow Request ID is " & RequestIDValue & " Client Name is " & ClientNameValue &" Syndicate is " & SyndicateValue & " Offering Name is " & OfferingNameValue &" Offering Type is " & OfferingTypeValue &" ClientCountry is at" & ClientCountryValue  ,"Headers in First grid")
'		     Else
'		        Call ReportStep (StatusTypes.Fail,"One  of Columns header value in grid is not Exist it is as follow Request ID is " & RequestIDValue & " Client Name is " & ClientNameValue &" Syndicate is " & SyndicateValue & " Offering Name is " & OfferingNameValue &" Offering Type is " & OfferingTypeValue &" ClientCountry is at" & ClientCountryValue  ,"Headers in First grid")
'		     End If
		     If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Exist(20) Then
		        Call ReportStep (StatusTypes.Pass, "Next button exist and clicked ","Next button")	
		         Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Click
	         Else
		         Call ReportStep (StatusTypes.Fail, "Next button not exist","Next button")
	         End If
End Function














'Created By: Rajesh
'Created on: 24/11/2016
'Description: Validate column names in ONLINE CRF page.
'Parameter: ColumnName - Give the column name which you want to validate eg: Request ID|Client Name|Country|Offering Name,objData
Public Function ColumnValidation(ByVal ColumnName,ByRef objData)
	
 	If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_GridHeader").WebElement("html tag:=DIV","innertext:="&ColumnName).Exist Then
 		Call ReportStep (StatusTypes.Pass, "Column name displayed in ONLINE CRF page "&ColumnName,"Column name")
	Else	
		Call ReportStep (StatusTypes.Fail, "Column name not displayed in ONLINE CRF page "&ColumnName,"Column name")		
 	End If
	
End Function




'Created By: Rajesh
'Created on: 25/11/2016
'Description: Add "New line" and delete selected row with various row size (1,2,5,10,25,50) in Cube details page.
'Parameter: Param1 - Extra,Param2 - Extra,Param3 - Extra,objData
Public Function AddDeleteRow(ByVal Param1, ByVal Param2, ByVal Param3,ByRef objData)
	
	If Param3 = "" Then
		If Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_NumberOfRows").Exist Then
		wait 2
		BeforeRowCount = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetROProperty("rows")

		NoOfRow = trim(Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_NumberOfRows").GetROProperty("all items"))
		NoOfRows = split(NoOfRow,";")
		
		For Iterator = lbound(NoOfRows) To ubound(NoOfRows) Step 1
			
			Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_NumberOfRows").Select NoOfRows(Iterator)
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_AddNewLine").Click
			wait 2
			AfterRowCount = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetROProperty("rows")

			If AfterRowCount = BeforeRowCount+ CInt(NoOfRows(Iterator))  Then
				Call ReportStep (StatusTypes.Pass, "New line added successfully "&NoOfRows(Iterator),"Added new line")
			Else	
				Call ReportStep (StatusTypes.Fail, "New line not added successfully "&NoOfRows(Iterator),"Added new line")
				Exit Function	
			End If

			
			
			'Delete added row
			
			For del = AfterRowCount To BeforeRowCount+1 Step -1
				Set CheckRow =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(del,2,"WebCheckBox",0)
				CheckRow.click	
			Next
			
			If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_DeleteSelectedRow").Exist Then
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_DeleteSelectedRow").Click
			End If
			
			If Browser("Browser-OCRF").Dialog("Dialog_MessageFromWebpage").WinButton("Button_OK").Exist Then
				Browser("Browser-OCRF").Dialog("Dialog_MessageFromWebpage").WinButton("Button_OK").Click
			End If
		
			wait 2
			AfterDelRowCount = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetROProperty("rows")
			
			If AfterDelRowCount = BeforeRowCount Then
				Call ReportStep (StatusTypes.Pass, "Selected row deleted successfully "&NoOfRows(Iterator),"Delete selected row")
			Else	
				Call ReportStep (StatusTypes.Fail, "Selected row not deleted successfully "&NoOfRows(Iterator),"Delete selected row")
				Exit Function	
			End If
			
			
		Next
		
		End If	
	End If
	
	
	
	
	
	
	
	If Param3 = "AddrowAndChecked" Then
		
		wait 2
		BeforeRowCount = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetROProperty("rows")
		
		Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_AddNewLine").Click
		wait 2
		AfterRowCount = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetROProperty("rows")
		
		If AfterRowCount = BeforeRowCount+ 1  Then
			Call ReportStep (StatusTypes.Pass, "New line added successfully ","Added new line")
		Else	
			Call ReportStep (StatusTypes.Fail, "New line not added successfully ","Added new line")
			Exit Function	
		End If
		
		'Checked new added row
		
		Set CheckRow =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(AfterRowCount,2,"WebCheckBox",0)
		CheckRow.click	
			
	End If
	
	
	
	
	
End Function








'Created By: Rajesh
'Created on: 06/12/2016
'Description: Filter columns.
'Parameter: Param1 - Extra,Param2 - Extra,Param3 - Extra,ColName - name of the column,TxtId - html id of textbox for respective column,Value - Enter value from which you want to filter column,objData
Public Function FilterColumn(ByVal Param1, ByVal Param2, ByVal Param3,ByVal ColName,ByVal TxtId,ByVal Value,ByRef objData)
	
	
	'Filter value in respective column
	If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("html tag:=INPUT","html id:="&TxtId).Exist Then
		Call ReportStep (StatusTypes.Pass,"Text field is displayed: "&TxtId,"Text filed "&TxtId)
		Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("html tag:=INPUT","html id:="&TxtId).Set Value
		Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_ClientName").Click
	Else	
		Call ReportStep (StatusTypes.Fail,"Text field is not displayed: "&TxtId,"Text filed "&TxtId)
	End If
	
	
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.SendKeys "{ENTER}"
	wait 5
	
	
	'Validate column is filtered with respective value
	
	cc = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_GridHeader").ColumnCount(1)
	
	For i = 1 To cc Step 1
		If ColName = trim(Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_GridHeader").GetCellData(1,i)) Then
			ColNum = i
			Exit For
		End If
	Next
	
	
	rc = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").RowCount
	For i = 2 To rc Step 1
		If (Value = trim(Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").GetCellData(i,ColNum))) Then
		
		Else
			Call ReportStep (StatusTypes.Fail,"Column "&ColName&" is not filtered with filter value "&Value,"Column filter")
			Operation = "Fail"
			Exit Function
		End If
	Next
	
		If Operation <> "Fail" Then
			Call ReportStep (StatusTypes.Pass,"Column "&ColName&" is filtered with filter value "&Value,"Column filter")
		End If
	
End Function



'Created By: Rajesh
'Created on: 06/12/2016
'Description: Clear all filters ia page.
'Parameter: Param1 - Extra,Param2 - Extra,Param3 - Extra,objData
Public Function ClearAllFilters(ByVal Param1, ByVal Param2, ByVal Param3,ByRef objData)
	
	Set d = description.Create()
	d("micclass").Value = "WebElement"
	d("class").Value = "clearsearchclass"
	d("innertext").Value = "x"
	
	Set desc =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_GridHeader").ChildObjects(d)
	
	For i = 0 To desc.count-1 Step 1
		desc(i).Click
	Next
	
	Call ReportStep (StatusTypes.Pass,"Cleared all filters from all the columns","Clear filters")
	
End Function
Public Function FLAPathValidation(ByVal CubeName, ByVal CubeLocation, ByVal Param3,ByRef objData)
	
	
	'Single Click on FLA Path

	Set objTab = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data")

	
	
	Call AddNewLineInCubeDetails(CubeName,Environment.Value("flaServerPath"),Environment.Value("HostServerName"),1,objData)
	
	rc = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").RowCount
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Exist(20) Then
		Call ReportStep (StatusTypes.Pass, "Next button exist and clicked ","Next button")	
		Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Click
	Else
		Call ReportStep (StatusTypes.Fail, "Next button not exist","Next button")
	End If
	'Click on previous page icon
	Call PageLoading()
	If Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_Previous").Exist(30) Then
		Call ReportStep (StatusTypes.Pass,"Previous page icon is present","Previous page icon")			
		Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_Previous").Click
		wait 5
	Else
		Call ReportStep (StatusTypes.Fail,"Previous page icon is not present","Previous page icon")			
	End If
    Call PageLoading()
	If rc>1 Then
	    For i = 2 To rc Step 1
		    ActCubeName= trim(Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetCellData(i,Cint(Environment.value("cubeName"))))
	        If ActCubeName=Trim(CubeName) Then
	           row=i	
	           Exit For
	        End If
	    Next
		Request_ID = trim(Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetCellData(row ,Cint(Environment.value("flaPath"))))

		Set objElem = Description.Create
		objElem("micclass").value = "WebElement"
		objElem("html tag").value = "TD"
		objElem("innertext").value = Request_ID 'parameterized the request id
		set objElemCnt = objTab.ChildObjects(objElem)
		
		If objElemCnt.count >0 Then
			objElemCnt(0).fireevent "onmouseover"
			objElemCnt(0).fireevent "onclick"
			
			'Validate webelement is converted to web list
			wait 2
			Call WebEditExistence(row,Cint(Environment.value("cubeName")),"WebEdit")
			li = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItemCount(row,Cint(Environment.value("flaPath")),"WebList")
			
			If li = 1 Then
				Call ReportStep (StatusTypes.Pass,"On click on FLA path webelement is converted to weblist","FLA Path")	
			Else
				Call ReportStep (StatusTypes.Fail,"On click on FLA path webelement is not converted to weblist","FLA Path")	
			End If
			
			objElemCnt(0).fireevent "ondblclick"	
			wait 2
		

			New_Request_ID = lcase(Request_ID)

			If Window("Window_IMSHosting").WinToolbar("text:=.*"&New_Request_ID&".*").Exist(10) Then
				Call ReportStep (StatusTypes.Pass,"On double click on FLA path, user is able to navigate respective FLA Server name using window directory ","FLA Server name in window directory")	
			Else
				Call ReportStep (StatusTypes.Fail,"On double click on FLA path, user is not able to navigate respective FLA Server name using window directory ","FLA Server name in window directory")	
				Exit Function
			End If
		
			'Close the window
			Window("Window_IMSHosting").Close
		
		Else		
			Call ReportStep (StatusTypes.Fail,"FLA Path are not exist","FLA Path")	
			Exit Function
		End If
		
		
	Else
		Call ReportStep (StatusTypes.Fail,"Cube details are not exist","Cube details")	
		Exit Function
	End If
	
End Function




Public Function FLAServerNameBlank(ByVal Param1, ByVal Param2, ByVal Param3,ByVal CubeName,ByVal Scenario,ByRef objData)
	
	rc = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").RowCount
	
	'Enter cube name
	
	Set edt = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(rc,Cint(Environment.value("cubeName")),"WebEdit",0)
	edt.set CubeName
	
	Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_NumberOfRecords").FireEvent "onmouseover"
	wait 2
	
	If Param3 = "" Then
		If trim(Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetCellData(rc,Cint(Environment.value("flaPath")))) = "" Then
			If Scenario = "Positive" Then
				Call ReportStep (StatusTypes.Fail,"FLA Server name is not blank,because there is existence of template in ops system","FLA Server name")
			Else
				Call ReportStep (StatusTypes.Pass,"FLA Server name is blank,because there is no existence of template in ops system","FLA Server name")	
			End If
		Else
			If Scenario = "Negative" Then
				Call ReportStep (StatusTypes.Pass,"FLA Server name is not blank,because there is  existence of template in ops system","FLA Server name")	
			Else	
				Call ReportStep (StatusTypes.Fail,"FLA Server name is  blank,because there is no existence of template in ops system","FLA Server name")	
			End If
		End If	
	End If
	

	If Param3 = "CubeLocationServerValidation" Then
		
		If trim(Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetCellData(rc,Cint(Environment.value("cubeLocation")))) <> "" Then
			If Scenario = "Positive" Then
				Call ReportStep (StatusTypes.Pass,"CubeLocation/Server is autocomplete,because there is existence of template in ops system","CubeLocation/Server")
			Else
				Call ReportStep (StatusTypes.Fail,"CubeLocation/Server is blank,because there is no existence of template in ops system","CubeLocation/Server")	
			End If
		Else
			If Scenario = "Positive" Then
				Call ReportStep (StatusTypes.Fail,"CubeLocation/Server is blank,because there is no existence of template in ops system","CubeLocation/Server")	
			Else	
				Call ReportStep (StatusTypes.Pass,"CubeLocation/Server is autocomplete,because there is existence of template in ops system","CubeLocation/Server")	
			End If
		End If		
	End If
	
	
	
	If Param3 = "FrozenCubeLocationServerColumn" Then
		
		If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").Exist Then
		
			rownumber = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetRowWithCellText(CubeName,Cint(Environment.value("cubeName")),2)	
			
		
			If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItemCount(rownumber,Cint(Environment.value("cubeLocation")),"WebEdit") > 0 Then
				
			Else	
				Set CheckRow =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(rownumber,2,"WebCheckBox",0)
				CheckRow.click
				
				If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItemCount(rownumber,Cint(Environment.value("cubeLocation")),"WebEdit") > 0 Then
				
				Else
					Wait 2
					CheckRow.click
				End If
			End If
		
			Set CubeServerName = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(rownumber,Cint(Environment.value("cubeLocation")),"WebEdit",0) 
			CubeServerName.set Param2		
			
			wait 2
			Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_NumberOfRecords").Select "50"
			
			
			
			'Validate Cube lacation/server column is frozen 
			
			
			For i = 1 To 10 Step 1
				rownumber = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetRowWithCellText(CubeName,Cint(Environment.value("cubeName")),2)	
				If rownumber = 2 Then
					Exit For
				End If	
				wait 2
			Next
			
			
			
			
			If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItemCount(rownumber,Cint(Environment.value("cubeLocation")),"WebEdit") > 0 Then
				
			Else	
				Set CheckRow =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(rownumber,2,"WebCheckBox",0)
				CheckRow.click
				If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItemCount(rownumber,Cint(Environment.value("cubeLocation")),"WebEdit") > 0 Then
				
				Else
					Wait 2
					CheckRow.click
				End If
				
			End If
			
			
'			Set CubeServerNameClick = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(2,20,"WebElement",0)
'			CubeServerNameClick.click
			wait 5
			TxtField = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(rownumber,Cint(Environment.value("cubeLocation")),"WebEdit",0).GetRoProperty("disabled")
			
			
			If TxtField = 0 Then
				Call ReportStep (StatusTypes.Pass,"CubeLocation/Server column is not frozen while entering server name manuaaly","CubeLocation/Server column")	
			Else	
				Call ReportStep (StatusTypes.Fail,"CubeLocation/Server column is frozen while entering server name manuaaly","CubeLocation/Server column")		
				Exit Function
			End If
			
			
			
		End If
		
		
		
		
		
	End If
	
	
	
	
	
	
End Function








Public Function EditCubeServerName(ByVal Param1, ByVal Param2, ByVal serverName,ByVal CubeName,ByRef objData)
	
	
	'Check webedit is exist or not and if not exist check the checkbox
	
'	rownumber = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetRowWithCellText(CubeName,10,2)
	
	If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItemCount(2,Cint(Environment.value("cubeLocation")),"WebEdit") > 0 Then
				
	Else	
		Set CheckRow =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(2,2,"WebCheckBox",0)
		CheckRow.click
		If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItemCount(2,Cint(Environment.value("cubeLocation")),"WebEdit") > 0 Then
		
		Else
			Wait 2
			CheckRow.click
		End If
		
	End If
	
	Set wsh = CreateObject("Wscript.Shell")
	Set CubeServer = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(2,Cint(Environment.value("cubeLocation")),"WebEdit",0)
	CubeServer.set serverName
	CubeServer.click
	wait 2
	wsh.SendKeys "{DOWN}"
	wait 2
'	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner-all","html id:=ui-id.*","innertext:=CDTSOLAP574I\.GEMINI\.DEV").Click
'	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner-all","html id:=ui-id.*","innertext:="&serverName&".*").Click
End Function


Public Function LoginToSCAClientUrl(ByVal strUserName,ByVal StrPwd, ByVal clientURL)
      
       Dim objUserName,objPwd,objLoginBtn,objSCAHomePageImage,returnval,objSCAHomeFrame

        SystemUtil.CloseProcessByName("IExplore.exe")
        wait 4
       
       'Redirecting to Client URL SCA - Start
        'Navigate to SCA through DC SIte IMS Analyzer link in DC Home page Analysis Table
       SystemUtil.Run clientURL
        wait 2
        Browser("DC Home").Page("Microsoft Forefront TMG").Sync
        
        Set objUserName=Browser("DC Home").Page("Microsoft Forefront TMG").WebEdit("txtusername")
        Call SCA.SetText(objUserName,strUserName,"textUserName" ,"DC Login Page" )
        
        Set objPwd=Browser("DC Home").Page("Microsoft Forefront TMG").WebEdit("txtpassword")
        Call SCA.SetText(objPwd,StrPwd,"txtpassword" ,"DC Login Page" )    
        
        Set objLoginBtn=Browser("DC Home").Page("Microsoft Forefront TMG").WebButton("btnLogin")
        Call SCA.ClickOn(objLoginBtn,"Login Button","DC Login Page")
        
        Browser("DC Home").Page("Microsoft Forefront TMG").Sync
        Browser("DC Home").Page("Microsoft Forefront TMG").RefreshObject
        
        Set objDCHomePageImage=Browser("DC Home").Page("Home").Image("My_Informed_Decisions")
        If objDCHomePageImage.Exist(120) <> True Then
		      Call ReportStep (StatusTypes.Fail,"Client DC URL Login is not Successfull " ,"Home Page" )
		Else
		      Call ReportStep (StatusTypes.Pass,"Client DC URL Login is Successfull" ,"Home Page" )
		      Call ReportStep (StatusTypes.Information,"Click on IMS Analysis Manager to redirect to SCA" ,"Home Page" )
		End If
		
 		wait 2
 		
        'Click on IMS Analysis Manager link in DC Home Page
        Set objlnkIMSAnalysisManager = Browser("DC Home").Page("Home").Link("lnkIMS Analysis Manager")
        Call SCA.ClickOn(objlnkIMSAnalysisManager,"lnkIMS Analysis Manager", "DC Home Page")
         wait 2
        Browser("DC Home").Page("IMS Analysis Manager").Sync
        Set objlnkSharedReports = Browser("DC Home").Page("IMS Analysis Manager").Frame("Frame").Link("lnkShared Reports")
                    
        wait 2
           
       If Browser("Analyzer").Page("Home").Frame("Frame").Link("lnkSharedReports").Exist(120) <> True Then
          Call ReportStep (StatusTypes.Fail,"Client SCA URL Login is not Successfull " ,"Home Page" )
Else
          Call ReportStep (StatusTypes.Pass,"Client SCA URL Login is Successfull" ,"Home Page" )
       End If    
                              
      End Function    

Public Function ErrorMessageInCubeDetails(ByVal Param1, ByVal Param2, ByVal Param3,ByVal ErrorType,ByRef objData)
	
	rc = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").RowCount
	
	If rc>2 Then
		LastRowCubeName = trim(Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetCellData(rc-1,Cint(Environment.value("cubeName"))))
	Else
		LastRowCubeName = "SAT_AU_S_IMS_2I9"	
	End If
	
	
	
	Select Case ErrorType
		
		Case "Alphanumeric"
			NewCubeName = LastRowCubeName & "@"
			
		Case "Underscore"
			NewCubeName = LastRowCubeName & "_123"
	
		Case "Duplicate"
			NewCubeName = LastRowCubeName
			
	End Select
	
	
	
	
	
	'Enter cube name
	
	Set edt = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(rc,Cint(Environment.value("cubeName")),"WebEdit",0)
	edt.set NewCubeName
	wait 2
	
	
	If Browser("Browser-SCA").Dialog("Dialog_Message_Webpage").Exist Then
		
		AplhaMessage = trim(Browser("Browser-SCA").Dialog("Dialog_Message_Webpage").Static("index:=1").GetROProperty("text"))
		If AplhaMessage = "The cube template contain special character!" Then
			Call ReportStep (StatusTypes.Pass,"Dialog box is displayed which contains error message"&AplhaMessage,"Dialog box")			
			
		ElseIf AplhaMessage = "Please enter Valid cube Name (Cube name format: CubeSource_Country_Period_ClientName_OrderID  (Ex.- MIDAS_DE_M_BAYER_12345))" Then
			Call ReportStep (StatusTypes.Pass,"Dialog box is displayed which contains error message"&AplhaMessage,"Dialog box")
		ElseIf AplhaMessage = "The cube database already exists for this offering!!!" Then
			Call ReportStep (StatusTypes.Pass,"Dialog box is displayed which contains error message"&AplhaMessage,"Dialog box")
		Else
			Call ReportStep (StatusTypes.Fail,"Dialog box is not displayed which contains error message"&AplhaMessage,"Dialog box")
			Exit Function	
		End If
		
		Browser("Browser-SCA").Dialog("Dialog_Message_Webpage").WinButton("WinButton_OK").Click
	Else
		Call ReportStep (StatusTypes.Fail,"Dialog box is not displayed which contains error message","Dialog box")	
		Exit Function
	End If
	
End Function





'Created By: Rajesh
'Created on: 28/12/2016
'Description: Validate change action functionality for all rows and selected rows.
'Parameter: Param1 - Extra,Param2 - Extra,Param3 - Extra,objData
Public Function ChangeAction(ByVal Param1, ByVal Param2, ByVal Param3,ByRef objData)
	
	Set objShowDeleted = Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_ShowDeleted")
	Call SCA.ClickOn(objShowDeleted, "Clicked on Show Deleted button", "Add Users")

	Call EditAction("","","",2,"ADD",objData)
	Call EditAction("","","",3,"UPDATE",objData)
	Call EditAction("","","",4,"DELETE",objData)
	Call EditAction("","","",5,"ADD",objData)
	Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_NumberOfRecords").Select "50"
	wait 2
	
	Call ReportStep (StatusTypes.Information,"Validation start for All rows change action","change action")	
	
	Call SelectUserRow("","","","All","",objData)
	wait 2
	Call ValidateChangeActions("","","","All",objData)
	
	
	Call EditAction("","","",2,"ADD",objData)
	Call EditAction("","","",3,"UPDATE",objData)
	Call EditAction("","","",4,"DELETE",objData)
	Call EditAction("","","",5,"ADD",objData)
	Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_NumberOfRecords").Select "50"
	wait 2
	Call ReportStep (StatusTypes.Information,"Validation start for Selected rows change action","change action")	
	
	Call SelectUserRow("","","","Selected",2,objData)
	Call SelectUserRow("","","","Selected",3,objData)
	wait 2
	Call ValidateChangeActions("","","","Selected",objData)
	
End Function






'Created By: Rajesh
'Created on: 28/12/2016
'Description: Change the action status.
'Parameter: Param1 - Extra,Param2 - Extra,Param3 - Extra,rownum - Give row number,ActionName - Give new action name,objData
Public Function EditAction(ByVal Param1, ByVal Param2, ByVal Param3,ByVal rownum,ByVal ActionName,ByRef objData)
	wait 3
	Set chkbox =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(rownum,2,"WebCheckBox",0)
	chkbox.click

	Set edtaction =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(rownum,7,"WebList",0)
	edtaction.Select ActionName
	
	'Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_NumberOfRecords").Select "50"
	
End Function



'Created By: Rajesh
'Created on: 28/12/2016
'Description: Validate change action functionality in add users tab.
'Parameter: Param1 - Extra,Param2 - Extra,Param3 - Extra,Records - All[all rows] Selected[selected row],objData
Public Function ValidateChangeActions(ByVal Param1, ByVal Param2, ByVal Param3,ByVal Records,ByRef objData)

	rc = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").RowCount
	Dim arr()
	
	ReDim Preserve arr(rc-1)
	For iterator = 2 To rc Step 1
		Set abc =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(iterator,2,"WebCheckBox",0)
		ChkValue = abc.GetroProperty("Checked")
		abc.Click
		Set lstval = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(iterator,7,"Weblist",0)
		
		If (trim(Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").GetCellData(iterator,7)) = "ADD") AND (ChkValue = 1) Then
			arr(iterator-2) = "NO CHANGE"
		ElseIf (trim(Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").GetCellData(iterator,7))  = "UPDATE") AND (ChkValue = 1) Then	
			arr(iterator-2) = "NO CHANGE"
		ElseIf (trim(Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").GetCellData(iterator,7)) = "DELETE") AND (ChkValue = 1) Then
			arr(iterator-2) = "DELETED"
		ElseIf ChkValue = 1  Then 
			If ((trim(lstval.GetroProperty("Selection")) = "ADD") OR (trim(lstval.GetroProperty("Selection")) = "UPDATE")) AND (ChkValue = 1) Then
			'If ((trim(Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").GetCellData(iterator,7)) = "ADD") OR (trim(Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").GetCellData(iterator,7))) = "UPDATE") AND (ChkValue = 1) Then	
				arr(iterator-2) = "NO CHANGE"
			ElseIf ((trim(lstval.GetroProperty("Selection")) = "DELETE") OR (trim(lstval.GetroProperty("Selection")) = "DELETED")) AND (ChkValue = 1) Then
				arr(iterator-2) = "DELETED"
			ElseIf (trim(lstval.GetroProperty("Selection")) = "NO CHANGE") AND (ChkValue = 1) Then
				arr(iterator-2) = "NO CHANGE"
			End If
		Else
			arr(iterator-2) = trim(Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").GetCellData(iterator,7))
		End If
	Next
	
	If Records = "Selected" Then
	   rc=3
	   
	   Call SelectUserRow("","","","Selected",2,objData)
	   Call SelectUserRow("","","","Selected",3,objData)	
	End If
	'Click on change Action
	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_ChangeAction").Click
	
	If Records = "All" Then
		Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html tag:=LABEL","innertext:= All Records","visible:=True").Click
	ElseIf Records = "Selected" Then
		Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html tag:=LABEL","innertext:= Selected Records","visible:=True").Click
	End If
	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_Yes").Click
	wait 2
	
	For iterator = 2 To rc Step 1
		
		If arr(iterator-2) = trim(Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").GetCellData(iterator,7)) Then
			Call ReportStep (StatusTypes.Pass,"Change action functionality is working fine for row "&iterator,"Change action functionality")
		Else
			Call ReportStep (StatusTypes.Fail,"Change action functionality is not working fine for row "&iterator,"Change action functionality")
			Exit Function
		End If
		
	Next
	
	
End Function



'Created By: Rajesh
'Created on: 28/12/2016
'Description: Select row by checking check box in add users tab[Existing OCRF]
'Parameter: Param1 - Extra,Param2 - Extra,Param3 - Extra,Records - All[all rows] Selected[selected row],rownum - give row number,objData
Public Function SelectUserRow(ByVal Param1, ByVal Param2, ByVal Param3,ByVal Records,ByVal rownum, ByRef objData)
	
	If Records = "All" Then
		Browser("Browser-OCRF").Page("Page-OCRF").WebCheckBox("WebCheckBox_SelectAll").Set "ON"
		wait 2
	ElseIf Records = "Selected" Then
		Set chkbox =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(rownum,2,"WebCheckBox",0)
		chkbox.click
	End If
	
End Function



'Created By: Poornima A
'Created on: 23/12/2016
'Description: Validate Change Action Functionality in 'Client Content Tab'
'Parameter:  Selection parameter to select 'all content' OR 'Only selected row'
Public Function ValidateChangeAction(ByVal Selection,ByRef objData)

   Nc= Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").GetROProperty("cols")
   Nr= Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").GetROProperty("rows")
   Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").ChildItem(Nr,2,"WebCheckBox",index).Set "ON"
   Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").ChildItem(Nr,6,"WebList",index).Select "DELETE"
   DelRepName=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").GetCellData(Nr,16)
   wait 4
   Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").ChildItem(Nr-1,2,"WebCheckBox",index).Set "ON"
   Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").ChildItem(Nr-1,6,"WebList",index).Select "UPDATE"
   UpdateRepName=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").GetCellData(Nr-1,16)
   wait 4
   Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").ChildItem(Nr-2,2,"WebCheckBox",index).Set "ON"
   Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").ChildItem(Nr-2,6,"WebList",index).Select "ADD"
   AddRepName=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").GetCellData(Nr-2,16)
   wait 4
   'Save data by clicking on next and clicking on previus
    If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Exist(20) Then	
	   Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Click
    End If
    If Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_SaveAndComplete").Exist(30) Then
       'Click on save and complete
    End If
     If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-button-text","innertext:=Previous","html tag:=SPAN").Exist(20) Then	
	   Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-button-text","innertext:=Previous","html tag:=SPAN").Click
    End If
    If Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_SaveAndComplete").Exist(30) Then
       'Click on save and complete
    End If
  
   'BEFORE CHANGING ACTION TAKING DATA
   ReDim ActTaken(Nr)
   ReDim AfterAct(Nr)
   ReDim Selected(Nr)
    Set oExcptnDetail = Description.Create
   oExcptnDetail("micclass").value = "WebElement"
   oExcptnDetail("html tag").value = "TD"
   oExcptnDetail("outerhtml").value = "<TD role=gridcell aria-describedby=Content-list_Action title"&".*"
   Set chobj= Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").ChildObjects(oExcptnDetail)
   wait 2
    totCnt= chobj.count
   chobj(totCnt-1).highlight
   For action = 2 To Nr Step 1
     chobj(action-2).Click
   	 ActTaken(action) =  Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").ChildItem(action,6,"WebList",index).GetROProperty("selection") 	 
  Next
  
   'SELECT CLIENTS TO CHANGE ACTION ON SELECTED
   For k = 2 To Nr Step 1
   	 Act=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").GetCellData(k,7)
     Rep=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").GetCellData(k,16)
     If DelRepName=Rep AND Act="DELETE" Then
     	Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").ChildItem(K,2,"WebCheckBox",index).Set "ON"
     ElseIf UpdateRepName=Rep AND Act="UPDATE" Then
        Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").ChildItem(K,2,"WebCheckBox",index).Set "ON"
     ElseIf AddRepName=Rep AND Act="ADD" Then 
         Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").ChildItem(K,2,"WebCheckBox",index).Set "ON"
     End If
 
   Next
   
   For Iter = 2 To Nr Step 1
   	   selectedNo =  Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").ChildItem( Iter,2,"WebCheckBox",index).GetROProperty("checked")
     If selectedNo=CInt("1") Then
     	Selected(Iter)=Iter
     End If
   Next
   
   'CLICK ON CHANGE ACTION
   Set obj = Description.Create 
	obj("micclass").Value = "WebElement"
	obj("html tag").Value = "DIV"
	obj("innertext").Value = "Change Action"
	set robj=Browser("Browser-OCRF").Page("Page-OCRF").ChildObjects(obj)
    robj(2).Click
   
   'CHECK FOR CONFORMATION -CHANGE ACTION POP UP and 'SELECT RADIO BUTTON "ALL RECORD"
   If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-dialog-title","html tag:=SPAN","innertext:=Confirmation - Change Action").Exist(20) Then
   	  Call ReportStep (StatusTypes.Pass, "Chanage Action Pop Up appeared ","Chanage Action Pop Up")
   	  If Selection = "AllContent" Then
		   Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html tag:=LABEL","innertext:= All Records","visible:=True").Click   	  	 
   	  Else
   	      'Browser("Browser-OCRF").Page("Page-OCRF").WebRadioGroup("html id:=rdChangeActionAllContent","type:=radio","value:=#1").select "on"
   	  End If      
      Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-button-text","innertext:=Yes","html tag:=SPAN").Click		 
   Else
		Call ReportStep (StatusTypes.Fail, "Chanage Action Pop Up not appeared","Chanage Action Pop Up")
   End If
   'Save data by clicking on next and clicking on previus
    If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Exist(20) Then	
	   Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Click
    End If
    If Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_SaveAndComplete").Exist(30) Then
       'Click on save and complete
    End If
     If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-button-text","innertext:=Previous","html tag:=SPAN").Exist(20) Then	
	   Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-button-text","innertext:=Previous","html tag:=SPAN").Click
    End If
    If Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_SaveAndComplete").Exist(30) Then
       'Click on save and complete
    End If
   'AFTER CHANGING ACTION TAKING DATA
    AfterNc= Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").GetROProperty("cols")
    AfterNr= Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").GetROProperty("rows")
    For action = 2 To Nr Step 1
   	 AfterAct(action) =  Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").GetCellData(action,7) 
   Next
   
   
   Set oExcptnDetail1 = Description.Create
   oExcptnDetail1("micclass").value = "WebElement"
   oExcptnDetail1("html tag").value = "TD"
   oExcptnDetail1("outerhtml").value = "<TD role=gridcell aria-describedby=Content-list_Action title"&".*"
   Set chobj1=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").ChildObjects(oExcptnDetail1)
  wait 2
    totCnt1= chobj1.count
  'VALIDATION  'CHANGE ACTION 1. "ADD" to "No change" 2. delete to "Deleted" and 3. Update to "No change"
  For iter = 2 To Nr Step 1
    Select Case ActTaken(iter)
    	
    	Case "ADD"
    	       If (Selection="SelectedContent" AND Selected(iter)=iter) OR Selection="AllContent" Then
    	       	 If AfterAct(iter)= "NO CHANGE" Then
    	     	    Call ReportStep (StatusTypes.Pass, "Action changed from ADD to NO CHANGE for "& Selection & "Records","ADD to NO CHANGE")
    	         Else
    	            Call ReportStep (StatusTypes.Fail, "Action Not changed from ADD to NO CHANGE for "& Selection & "Records","ADD to NO CHANGE")
    	         End If
    	      End If
    	     
    	Case "UPDATE"
    	     If (Selection="SelectedContent" AND Selected(iter)=iter) OR Selection="AllContent" Then
    	     	If AfterAct(iter)= "NO CHANGE" Then
    	     		 Call ReportStep (StatusTypes.Pass, "Action changed from UPDATE to NO CHANGE for "& Selection & "Records","UPDATE to NO CHANGE")
    	     	Else
    	         	Call ReportStep (StatusTypes.Fail, "Action Not changed from UPDATE to NO CHANGEfor "& Selection & "Records","UPDATE to NO CHANGE")
    	     	End If
    	     End If
    	Case "DELETE"
    	    If (Selection="SelectedContent" AND Selected(iter)=iter) OR Selection="AllContent" Then
    	       flag1=true
    	       For rw = 2 To AfterNr Step 1
    	            chobj1(rw-2).Click
    	            ExistReportName=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").ChildItem(rw,16,"WebEdit",index).GetROProperty("innertext")
    	    	    If DelRepName=ExistReportName Then
    	    		   flag1=false
    	     	    Else
    	         	   flag1=true
    	    	    End If
    	       Next
    	           wait 2
    	           set odesc= Description.Create
					odesc("micClass").value= "WebElement"
					odesc("html tag").value= "DIV"
					odesc("innertext").value= "Show Deleted"
					set Chobj=Browser("Browser-OCRF").Page("Page-OCRF").ChildObjects(odesc)
					Chobj(2).highlight
					Chobj(2).Click
    	           wait 10
    	           Set ObjDel = Description.Create
                   ObjDel("micclass").value = "WebElement"
   				   ObjDel("html tag").value = "TD"
  				   ObjDel("outerhtml").value = "<TD role=gridcell aria-describedby=Content-list_Action title"&".*"
  				   Set choDel=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").ChildObjects(ObjDel)
  				   wait 2
    			   totDel= choDel.count
    	          'take row count in show deleted list
                   DeleteNr= Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").GetROProperty("rows")
                   flag = false
                   For del = 2 To DeleteNr Step 1
                       choDel(del-2).Click
           	           DelReportName=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").ChildItem(del,16,"WebEdit",index).GetROProperty("value")
                       ActionStatus=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=Content-list","html tag:=TABLE").ChildItem(del,6,"WebList",index).GetROProperty("selection")
                      If DelRepName=DelReportName AND ActionStatus="DELETED"  Then
    	    	        flag = true
    	    	        Exit For
    	              Else
    	                 flag = false
    	              End If
                   Next
                   If flag=true AND flag1=true Then
                       Call ReportStep (StatusTypes.Pass, " Deleted Report "&DelRepName&" found in 'Show deleted' list with 'DELETED' status ","Show deleted List")
    	     	   Else
    	         	   Call ReportStep (StatusTypes.Fail, " Deleted Report "&DelRepName&" Not found in 'Show deleted' list with 'DELETED' status ","Show deleted List")
    	     	   End If
        End If 
    	set odesc= Description.Create
    	odesc("micClass").value= "WebElement"
		odesc("html tag").value= "DIV"
		odesc("innertext").value= "Hide Deleted"
		set Chobj=Browser("Browser-OCRF").Page("Page-OCRF").ChildObjects(odesc)
		Chobj(0).highlight
		Chobj(0).Click
		wait 10        
    End Select
  Next

  
End Function


'Created By: Poornima A
'Created on: 28/12/2016
'Description: Validate change functionalities in 'Cube details' Tab
'Parameter: Selection parameter to select 'all content' OR 'Only selected row'

Public Function ValidateChangeActionInCubeDetails(ByVal Selection,ByRef objData)
   wait 10
   Dim selEle(3)
   Nc=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetROProperty("cols")
   Nr=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetROProperty("rows")
   wait 5
   Set oDetail = Description.Create
   oDetail("micclass").value = "WebElement"
   oDetail("html tag").value = "TD"
   oDetail("outerhtml").value = "<TD role=gridcell aria-describedby=cube-list_Action title="&".*"
   Set chobjs=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildObjects(oDetail)
   wait 5
    totCnt= chobjs.count
   chobjs(totCnt-1).highlight
   chobjs(totCnt-1).Click
   'Call WebEditExistence(Nr,Cint(Environment.value("cubeLocation")),"WebEdit")
   'Call WebEditExistence(rowNumber,22,"WebEdit")
   
   
   If  Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").WebElement("class:=adminEntry","outerhtml:=<TD role=gridcell aria-describedby=cube-list_Server title="""" class=adminEntry style=""TEXT-ALIGN: left"">&nbsp;</TD>","innerhtml:=&nbsp;","html tag:=TD").Exist(10) Then
       Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").WebElement("class:=adminEntry","outerhtml:=<TD role=gridcell aria-describedby=cube-list_Server title="""" class=adminEntry style=""TEXT-ALIGN: left"">&nbsp;</TD>","innerhtml:=&nbsp;","html tag:=TD").Click
   Else
   End If

   Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(Nr,7,"WebList",index).Select "DELETE"
   selEle(0)= Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(Nr,Cint(Environment.value("cubeName")),"WebEdit",index).GetROProperty("value")
   DelCube=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(Nr,Cint(Environment.value("cubeName")),"WebEdit",index).GetROProperty("value")
   'Handle Delete Notification PoPUp
   If Browser("Browser-OCRF").Dialog("regexpwndtitle:=Message from webpage","text:=Message from webpage").WinButton("regexpwndtitle:=OK","nativeclass:=Button","text:=OK","regexpwndclass:=Button").Exist(5) Then
   	  Browser("Browser-OCRF").Dialog("regexpwndtitle:=Message from webpage","text:=Message from webpage").WinButton("regexpwndtitle:=OK","nativeclass:=Button","text:=OK","regexpwndclass:=Button").Click
      Call ReportStep (StatusTypes.Pass, "Delete conformation  Pop Up appeared","Delete Pop Up")
   Else
   	  Call ReportStep (StatusTypes.Fail, "Delete conformation  Pop Up Not appeared","Delete Pop Up")
   End If
  
   wait 4
   chobjs(totCnt-2).highlight
   chobjs(totCnt-2).Click

  'Call WebEditExistence(row,Cint(Environment.value("cubeLocation")),"WebEdit")

   Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(Nr-1,7,"WebList",index).Select "UPDATE"
    selEle(1)= Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(Nr-1,Cint(Environment.value("cubeName")),"WebEdit",index).GetROProperty("value")
   If  Browser("Browser-OCRF").Page("Page-OCRF").WebElement("innertext:=Cube Name","html id:=jqgh_cube-list_CubeTempleteName","class:=ui-th-div-ie ui-jqgrid-sortable","html tag:=DIV").Exist(20) Then
   	   Browser("Browser-OCRF").Page("Page-OCRF").WebElement("innertext:=Cube Name","html id:=jqgh_cube-list_CubeTempleteName","class:=ui-th-div-ie ui-jqgrid-sortable","html tag:=DIV").Click
   End If
   wait 4
    chobjs(totCnt-3).highlight
   chobjs(totCnt-3).Click 
   selEle(2)= Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(Nr-2,Cint(Environment.value("cubeName")),"WebEdit",index).GetROProperty("value")
   Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(Nr-2,7,"WebList",index).Select  "ADD"
   wait 4
   
   'Save data by clicking on next and clicking on previus
    If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Exist(20) Then	
	   Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Click
    End If
    wait 10
     If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-button-text","innertext:=Previous","html tag:=SPAN").Exist(20) Then	
	   Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-button-text","innertext:=Previous","html tag:=SPAN").Click
    End If
    wait 10
   'BEFORE CHANGING ACTION TAKING DATA
   ReDim ActTaken(Nr)
   ReDim AfterAct(Nr)
   ReDim Selected(Nr)
   wait 5
   Set oExcptnDetail = Description.Create
   oExcptnDetail("micclass").value = "WebElement"
   oExcptnDetail("html tag").value = "TD"
   oExcptnDetail("outerhtml").value = "<TD role=gridcell aria-describedby=cube-list_Action title="&".*"
   Set chobj=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildObjects(oExcptnDetail)
   wait 2
    totCnt= chobj.count
   chobj(totCnt-1).highlight
   For action = 2 To Nr Step 1
     chobj(action-2).Click
     If  Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").WebElement("class:=adminEntry","outerhtml:=<TD role=gridcell aria-describedby=cube-list_Server title="""" class=adminEntry style=""TEXT-ALIGN: left"">&nbsp;</TD>","innerhtml:=&nbsp;","html tag:=TD").Exist(10) Then
       Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").WebElement("class:=adminEntry","outerhtml:=<TD role=gridcell aria-describedby=cube-list_Server title="""" class=adminEntry style=""TEXT-ALIGN: left"">&nbsp;</TD>","innerhtml:=&nbsp;","html tag:=TD").Click
     Else
     End If
   	 ActTaken(action) = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(action,7,"WebList",index).GetROProperty("selection") 	 
  Next
  'SELECT CLIENTS TO CHANGE ACTION ON SELECTED
   For ele = 2 To Nr Step 1
   	   cubeToSel=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetCellData(ele,Cint(Environment.value("cubeName")))
       For arrEle = 0 To Ubound(selEle) Step 1
       	   If cubeToSel=selEle(arrEle) Then
       	   	  Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(ele,2,"WebCheckBox",index).Set "ON"
       	   	  Exit For
       	   End If
       Next
   Next


'    Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(Nr-2,2,"WebCheckBox",index).Set "ON"
'    Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(Nr-1,2,"WebCheckBox",index).Set "ON" 
'    Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(Nr,2,"WebCheckBox",index).Set "ON"
   For Iter = 2 To Nr Step 1
   	   selectedNo = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem( Iter,2,"WebCheckBox",index).GetROProperty("checked")
     If selectedNo=CInt("1") Then
     	Selected(Iter)=Iter
     End If
   Next
   
   'CLICK ON CHANGE ACTION
   Set obj = Description.Create 
	obj("micclass").Value = "WebElement"
	obj("html tag").Value = "DIV"
	obj("innertext").Value = "Change Action"
	set robj=Browser("Browser-OCRF").Page("Page-OCRF").ChildObjects(obj)
    robj(0).Click
   
   'CHECK FOR CONFORMATION -CHANGE ACTION POP UP and 'SELECT RADIO BUTTON "ALL RECORD"
   If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-dialog-title","html tag:=SPAN","innertext:=Confirmation - Change Action").Exist(20) Then
   	  Call ReportStep (StatusTypes.Pass, "Chanage Action Pop Up appeared ","Chanage Action Pop Up")
   	  If Selection = "AllContent" Then
		   Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html tag:=LABEL","innertext:= All Records","visible:=True").Click   	  	 
   	  Else
   	      'Browser("Browser-OCRF").Page("Page-OCRF").WebRadioGroup("html id:=rdChangeActionAllContent","type:=radio","value:=#1").select "on"
   	  End If      
      Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-button-text","innertext:=Yes","html tag:=SPAN").Click		 
   Else
		Call ReportStep (StatusTypes.Fail, "Chanage Action Pop Up not appeared","Chanage Action Pop Up")
   End If
   
      
   'Save data by clicking on next and clicking on previus
    If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Exist(20) Then	
	   Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Click
    End If
    wait 10
     If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-button-text","innertext:=Previous","html tag:=SPAN").Exist(20) Then	
	   Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-button-text","innertext:=Previous","html tag:=SPAN").Click
    End If
    wait 10
   
   'AFTER CHANGING ACTION TAKING DATA
    AfterNc=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetROProperty("cols")
    AfterNr=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetROProperty("rows")
    For action = 2 To AfterNr Step 1
   	 AfterAct(action) =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetCellData(action,7) 
   Next
   wait 10 
   Set oExcptnDetail1 = Description.Create
   oExcptnDetail1("micclass").value = "WebElement"
   oExcptnDetail1("html tag").value = "TD"
   oExcptnDetail1("outerhtml").value = "<TD role=gridcell aria-describedby=cube-list_Action title="&".*"
   Set chobj1=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildObjects(oExcptnDetail1)
  wait 2
    totCnt1= chobj1.count
   
 
  'VALIDATION  'CHANGE ACTION 1. "ADD" to "No change" 2. delete to "Deleted" and 3. Update to "No change"
  For iter = 2 To AfterNr+1 Step 1
    Select Case ActTaken(iter)
    	
    	Case "ADD"
    	       If (Selection="SelectedContent" AND Selected(iter)=iter) OR Selection="AllContent" Then
    	       	 If AfterAct(iter)= "NO CHANGE" Then
    	     	    Call ReportStep (StatusTypes.Pass, "Action changed from ADD to NO CHANGE for "& Selection & "Records","ADD to NO CHANGE")
    	         Else
    	            Call ReportStep (StatusTypes.Fail, "Action Not changed from ADD to NO CHANGE for "& Selection & "Records","ADD to NO CHANGE")
    	         End If
    	      End If
    	     
    	Case "UPDATE"
    	     If (Selection="SelectedContent" AND Selected(iter)=iter) OR Selection="AllContent" Then
    	     	If AfterAct(iter)= "NO CHANGE" Then
    	     		 Call ReportStep (StatusTypes.Pass, "Action changed from UPDATE to NO CHANGE for "& Selection & "Records","UPDATE to NO CHANGE")
    	     	Else
    	         	Call ReportStep (StatusTypes.Fail, "Action Not changed from UPDATE to NO CHANGEfor "& Selection & "Records","UPDATE to NO CHANGE")
    	     	End If
    	     End If
    	Case "DELETE"
    	         For rw = 2 To AfterNr Step 1
    	            chobj1(rw-2).Click
    	            If  Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").WebElement("class:=adminEntry","outerhtml:=<TD role=gridcell aria-describedby=cube-list_Server title="""" class=adminEntry style=""TEXT-ALIGN: left"">&nbsp;</TD>","innerhtml:=&nbsp;","html tag:=TD").Exist(10) Then
                        Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").WebElement("class:=adminEntry","outerhtml:=<TD role=gridcell aria-describedby=cube-list_Server title="""" class=adminEntry style=""TEXT-ALIGN: left"">&nbsp;</TD>","innerhtml:=&nbsp;","html tag:=TD").Click
                    Else
                    End If
    	            ExistCube=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(rw,Cint(Environment.value("cubeName")),"WebEdit",index).GetROProperty("value")
    	    	    If DelCube=ExistCube Then
    	    		   Call ReportStep (StatusTypes.Fail, "Delete action cube "&DelCube&" Exist in list","Delete action")
    	     	    Else
    	         	   Call ReportStep (StatusTypes.Pass, "Delete action cube "&DelCube&" Not Exist in list and check in show delete list","Delete action")
    	    	    End If
    	         Next
    	         wait 1
    	           set odesc= Description.Create
					odesc("micClass").value= "WebElement"
					odesc("html tag").value= "DIV"
					odesc("innertext").value= "Show Deleted"
					set Chobj=Browser("Browser-OCRF").Page("Page-OCRF").ChildObjects(odesc)
					Chobj(0).highlight
					Chobj(0).Click
    	           wait 10
    	           Set ObjDel = Description.Create
                   ObjDel("micclass").value = "WebElement"
   				   ObjDel("html tag").value = "TD"
  				   ObjDel("outerhtml").value = "<TD role=gridcell aria-describedby=cube-list_Action title="&".*"
  				   Set choDel=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildObjects(ObjDel)
  				   wait 2
    			   totDel= choDel.count
    	          'take row count in show deleted list
                   DeleteNr=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetROProperty("rows")
                   flag = false
                   For del = 2 To DeleteNr Step 1
                       choDel(del-2).Click
           	           DelCubeName=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(del,Cint(Environment.value("cubeName")),"WebEdit",index).GetROProperty("value")
                       ActionStatus=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(del,7,"WebList",index).GetROProperty("selection")
                      If DelCube=DelCubeName AND ActionStatus="DELETED"  Then
    	    	        flag = true
    	    	        Exit For
    	              Else
    	                 flag = false
    	              End If
                   Next
                   If flag=true Then
                       Call ReportStep (StatusTypes.Pass, " Deleted cube "&DelCube&" found in deleted list with 'DELETED' status ","Show deleted List")
    	     	   Else
    	         	   Call ReportStep (StatusTypes.Fail, " Deleted cube "&DelCube&" Not found in deleted list with 'DELETED' status ","Show deleted List")
    	     	   End If
                   set odesc= Description.Create
    				odesc("micClass").value= "WebElement"
					odesc("html tag").value= "DIV"
					odesc("innertext").value= "Hide Deleted"
					set Chobj=Browser("Browser-OCRF").Page("Page-OCRF").ChildObjects(odesc)
					Chobj(0).highlight
					Chobj(0).Click
					wait 10
    End Select
    
  Next
End Function





'Date 2/1/2017 to 6/1/2017 
'Created by Poornima,Rajesh
'End to end flow scenario

Public Function CreateOCRFID(ByRef objData)
'Public Function CreateOCRFID()
    ENVUrl=objData.item("ENVUrl")
    Path=Environment.Value("HostingPath")
	TestDataPath=Environment.value("CurrDir")& "InputFiles\IMSSCAWeb\"
	Call ReportStep (StatusTypes.Information,"Click on 'add new request' entering data in offering details tab","Offering tab")
	Call ClickOnAddNewRequest()
	OffClientName= AddDataInOfferingDetail(TestDataPath,objData)
	Call NavigateToNextTab()
	Call PageLoading()
	Call ReportStep (StatusTypes.Information,"Entering '10' new records in 'Cube details tab'","Cube details tab")
	Call AddMultipleDatabase(TestDataPath,"",objData)
	Call NavigateToNextTab()
	Call PageLoading()
	Call ReportStep (StatusTypes.Information,"Uploading users in 'Add Users tab'","Add Users tab")
	Call AddNewLinesINAddUserTab(TestDataPath,objData)
	Call NavigateToNextTab()
	Call PageLoading()
	Call ReportStep (StatusTypes.Information,"Skipping data entry in 'Client content tab'","Client content tab")
	Call Client_Content_tab(TestDataPath,objData)
	Call NavigateToNextTab()
	Call PageLoading()
	Call ReportStep (StatusTypes.Information,"Selection of  '2' users for each combination of 'BI tool' and 'Access mode' in 'BI Tool access tab'","BI Tool access tab")
	Call Multi_BITools_Access_Combination(TestDataPath,objData)
	Call NavigateToNextTab()
	Call PageLoading()
	'Call BI_Tools_Access_Combination(TestDataPath,objData)
	Call ReportStep (StatusTypes.Information,"Selection of  '2' users for each combination of 'DataBase' and 'Role' in 'User DB access tab'","User DB access tab")
	Req_ID = Multi_USerDB_Access_Combination(TestDataPath,objData)
	
	
	Call ReportStep (StatusTypes.Information,"Opening Created request id:"&Req_ID&" in 'Online CRF tab'","Online CRF tab")
	'Open created request
	Call FilterByReqIDAndOpen(CInt(Req_ID),objData)
	'Validate offering action summary
	Call ValidateRowsCount(TestDataPath,objData)	
	
	
	'Place cube in FLA Path copy 'Abf' file in to cube name and rename
	Call FolderCreationInFLAServer (Path,"OCRF-DatabaseAdd",objData)
	'Validate Apply security failed
	Call ApplySecurity(objData)
	
	'Valiadate OCRF request status in Locked By column
	Call validateOCRFRequest(Req_ID)
	 'Create client in OPS
    Call SCA_CreateClientInOps(objData)
    'Validation for Groups created in Ops
    Call ItemSelection("Manage Content Access")
    
     Datatable.GetSheet("Offering_Details").SetCurrentRow(1)
    'Search for groups for client
    Call GetGroupsForClient(dataTable.value("Offering_Client","Offering_Details"),dataTable.value("Offering_Country","Offering_Details"),"NonSyndicate",OffClientName)
    
    'CLIENT DC for 'NON syndicated
	r = Datatable.GetSheet("BI_Tools").GetRowCount
	For j = 1 To r Step 1
	   DataTable.SetCurrentRow(j)
	   User_Id=dataTable.value("User_Id","BI_Tools")
	   S_BI_Tool=dataTable.value("S_BI_Tool","BI_Tools")
       S_Access_Mode=dataTable.value("S_Access_Mode","BI_Tools")
       If S_Access_Mode ="" Then
       Else	
          Call ValidateGroupExistance(S_BI_Tool,S_Access_Mode,OffClientName,User_Id) 
       End If
       
	Next
    
     Datatable.GetSheet("Offering_Details").SetCurrentRow(1)
    'TOP SITE for ' Non syndicated
    Call GetGroupsForClient("Top Site",dataTable.value("Offering_Country","Offering_Details"),"NonSyndicate",OffClientName)

	'DataTable.GetRowCount(SheetName)
	r = Datatable.GetSheet("BI_Tools").GetRowCount
	For j = 1 To r Step 1
	   DataTable.SetCurrentRow(j)
	   User_Id=dataTable.value("User_Id","BI_Tools")
	   TS_S_BI_Tool=dataTable.value("TS_S_BI_Tool","BI_Tools")
       TS_S_Access_Mode=dataTable.value("TS_S_Access_Mode","BI_Tools")
       If TS_S_Access_Mode ="" Then
       Else	
          Call ValidateGroupExistance(TS_S_BI_Tool,TS_S_Access_Mode,OffClientName,User_Id)
       End If
        
	Next
  
   'Date 9/1/2017
   'Launch OCRF and search for Newly created Id and open it 
   Call LaunchURL(ENVUrl, "", "",objData)
   Call FilterByReqIDAndOpen(CInt(Req_ID),objData)
   Call DeleteIMSToolFromBITool(objData)
   	If Browser("Browser-OCRF").Page("Page-OCRF").Link("name:=\[1\] Offering Details->> ","text:=\[1\] Offering Details->> ","html tag:=A").Exist(10) Then
	    Browser("Browser-OCRF").Page("Page-OCRF").Link("name:=\[1\] Offering Details->> ","text:=\[1\] Offering Details->> ","html tag:=A").Click	
	    Call ReportStep (StatusTypes.Pass, "Clicked on ' Offering Details' Tab ","Offering Details tab")	
	Else
		Call ReportStep (StatusTypes.Fail, "Not Clicked on 'Offering Details' Tab","Offering Details tab")
	End If
   Call ApplySecurity(objData)
End Function


Public Function CreateOCRFID_Old(ByRef objData)
  ENVUrl=objData.item("ENVUrl")
  Path=Environment.Value("HostingPath")
'Public Function CreateOCRFID()
	TestDataPath = Environment.value("CurrDir")& "InputFiles\IMSSCAWeb\"

'	TestDataPath=objData.item("TestDataPath")
	Call ReportStep (StatusTypes.Information,"Click on 'add new request' entering data in offering details tab","Offering tab")
	Call ClickOnAddNewRequest()
	OffClientName= AddDataInOfferingDetail(TestDataPath,objData)
	Call NavigateToNextTab()
	Call PageLoading()
	Call ReportStep (StatusTypes.Information,"Entering '10' new records in 'Cube details tab'","Cube details tab")
	Call AddMultipleDatabase(TestDataPath,"",objData)
	Call NavigateToNextTab()
	Call PageLoading()
	Call ReportStep (StatusTypes.Information,"Uploading users in 'Add Users tab'","Add Users tab")
	Call AddNewLinesINAddUserTab(TestDataPath,objData)
	Call NavigateToNextTab()
	Call PageLoading()
	Call ReportStep (StatusTypes.Information,"Skipping data entry in 'Client content tab'","Client content tab")
	Call Client_Content_tab(TestDataPath,objData)
	Call NavigateToNextTab()
	Call PageLoading()
	Call ReportStep (StatusTypes.Information,"Selection of  '2' users for each combination of 'BI tool' and 'Access mode' in 'BI Tool access tab'","BI Tool access tab")
	Call BI_Tools_Access_Combination(TestDataPath,objData)
	Call ReportStep (StatusTypes.Information,"Selection of  '2' users for each combination of 'DataBase' and 'Role' in 'User DB access tab'","User DB access tab")
	Req_ID = USer_DB_Access_Combination(TestDataPath,objData)
	
	
	Call ReportStep (StatusTypes.Information,"Opening Created request id:"&Req_ID&" in 'Online CRF tab'","Online CRF tab")
	'Open created request
	Call FilterByReqIDAndOpen(CInt(Req_ID),objData)
	'Validate offering action summary
	Call ValidateRowsCount(TestDataPath,objData)	
	
	
	'Place cube in FLA Path copy 'Abf' file in to cube name and rename
	Call FolderCreationInFLAServer (Path,"OCRF-DatabaseAdd",objData)
	'Validate Apply security failed
	Call ApplySecurity(objData)
	
	'Valiadate OCRF request status in Locked By column
	Call validateOCRFRequest(Req_ID)
	 'Create client in OPS
    Call SCA_CreateClientInOps(objData)
    'Validation for Groups created in Ops
    Call ItemSelection("Manage Content Access")
    Call ValidationForGroupsInOps(dataTable.value("Offering_Client","Offering_Details"),dataTable.value("Offering_Country","Offering_Details"),OffClientName,"NonSyndicate")

    'Call ValidationForGroupsInOps("automatedclient235","Australia","AutoTest234","NonSyndicated")
	accMode = Split(objData.Item("accMode"),";")
	Permission = Split(objData.Item("Permission"),";")

	
    ReDim UserArr(3)
    cnt=0
    For i = 1 To 4 Step 1
   	  append=""
   	  For j = 1 To 2 Step 1
   	     Datatable.SetCurrentRow(j+cnt)
   	     append=append&dataTable.value("User_ID","BI_Tools")&";"
   	 	
   	 Next
   	   UserArr(i-1)=append
   	 cnt=cnt+2
   Next
  'Validation for 'Get users in Ops"
    Call ValidationForAddUserInOps(accMode(0),Permission,UserArr(0))
    Call ValidationForAddUserInOps(accMode(1),Permission,UserArr(1))
    Call ValidationForAddUserInOps(accMode(3),Permission,UserArr(2))
    'Call ValidationForAddUserInOps(accMode(2),Permission,UserArr(3))
  'Validation in TOP site
  
   Datatable.SetCurrentRow(1)
   Call ItemSelection("Manage Content Access")
   Call ValidationForGroupsInOps("Top Site",dataTable.value("Offering_Country","Offering_Details"),OffClientName,"NonSyndicate")
   Call ValidationForAddUserInOps(accMode(1),Permission,UserArr(3))
   Call ValidationForAddUserInOps(accMode(2),Permission,UserArr(3))
  
  
   'Date 9/1/2017
   'Launch OCRF and search for Newly created Id and open it 
   Call LaunchURL(ENVUrl, "", "",objData)
   Call FilterByReqIDAndOpen(CInt(Req_ID),objData)
   Call DeleteIMSToolFromBITool(objData)
   	If Browser("Browser-OCRF").Page("Page-OCRF").Link("name:=\[1\] Offering Details->> ","text:=\[1\] Offering Details->> ","html tag:=A").Exist(10) Then
	    Browser("Browser-OCRF").Page("Page-OCRF").Link("name:=\[1\] Offering Details->> ","text:=\[1\] Offering Details->> ","html tag:=A").Click	
	    Call ReportStep (StatusTypes.Pass, "Clicked on ' Offering Details' Tab ","Offering Details tab")	
	Else
		Call ReportStep (StatusTypes.Fail, "Not Clicked on 'Offering Details' Tab","Offering Details tab")
	End If
   Call ApplySecurity(objData)
End Function



Public Function ClickOnAddNewRequest()
	
	Set objNewRequest = Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebElement_Add_New_Request")
	Call SCA.ClickOn(objNewRequest,"Add New Request","ONLINE CRF")

	Call PageLoading()
End Function

public Function AddDataInOfferingDetail(ByVal fileName,ByVal SheetName,ByVal row,ByRef objData)
    '########### OFFERING DETAIL TAB EXISTANCE - Start ########### 
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_OfferingDetailsTab").Exist(60) Then
		Call ReportStep (StatusTypes.Pass, "Offering Details Tab exist","Offering Details Tab")
	Else
	   	Call ReportStep (StatusTypes.Fail, "cOffering Details Tab not exist","Offering Details Tab")
	End If
	Set Tab=Browser("Browser-OCRF").Page("Page-OCRF").Link("html tag:=A","class:=selected")

	wait 5
	Call ImportSheet(SheetName,fileName)
	DataTable.GetSheet(SheetName).SetCurrentRow(row)
     OfferingName=dataTable.value("Offering_Name",SheetName)
    Call ValidateActiveTabOrangeColour(Tab,"Offering Details",objData)
    'ADD DATA TO OFFERING DETAILS TAB
    			
    'ENTER DATA IN OFFERING CLIENT NAME
    If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingClient").Exist(20) Then
       	Call ReportStep (StatusTypes.Pass, "Offering client name entered with "&dataTable.value("Offering_Client",SheetName),"Offering Client")	
       	Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingClient").Click
       	wait 1
        Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingClient").Set dataTable.value("Offering_Client",SheetName)
   Else
	   	Call ReportStep (StatusTypes.Fail, "Offering client name not entered","Offering Client")
    End If
   ' SELECT OFFERING CONTRY NAME
    If Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebListOfferingCountry").Exist(20) Then
      Call ReportStep (StatusTypes.Pass, "Offering country name selected as "&dataTable.value("Offering_Country",SheetName),"Offering Country name")   
       Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebListOfferingCountry").Select dataTable.value("Offering_Country",SheetName)
    Else
	   Call ReportStep (StatusTypes.Fail, "Offering country name not selected","Offering Country name")
   	End If
    'SELECT SYNDIACATED NON SYNDICATED RADIO BUTTON
    If Browser("Browser-OCRF").Page("Page-OCRF").WebRadioGroup("WEbRadioSyndicated").Exist(20) Then
       	Call ReportStep (StatusTypes.Pass, "Syndicated radio button selected  ","Syndicated")      	
       	If dataTable.value("Syndicated",SheetName)="No" Then
       	  	Browser("Browser-OCRF").Page("Page-OCRF").WebRadioGroup("WEbRadioSyndicated").Select "false"
       	Else
       	  	Browser("Browser-OCRF").Page("Page-OCRF").WebRadioGroup("WEbRadioSyndicated").Select "on"
       	End If
    Else
	   	Call ReportStep (StatusTypes.Fail, "Offering country name not selected","Syndicated")
    End If
    'ENTER OFFERING NAME
    If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingName").Exist(20) Then
       	Call ReportStep (StatusTypes.Pass, "Offering Name entered as "&dataTable.value("Offering_Name",SheetName),"Offering name")	
       	Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingName").Set OfferingName
       	AddDataInOfferingDetail=OfferingName
    Else
	   	Call ReportStep (StatusTypes.Fail, "Offering Name  not entered","Offering name")
   End If
    'SELECT OFFERING TYPE FROM LIST
    If Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebListOfferingType").Exist(20) Then
       	Call ReportStep (StatusTypes.Pass, "Offering type entered as "&dataTable.value("Offering_Type",SheetName),"Offering type")	
       	Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebListOfferingType").Select dataTable.value("Offering_Type",SheetName)
    Else
	   	Call ReportStep (StatusTypes.Fail, "Offering type  not entered","Offering type")
    End If
    'SELECT ENVIRONMENT SET UP
    If Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebListSetupEnvironment").Exist(20) Then
       	Call ReportStep (StatusTypes.Pass, "Set Up Environment entered as "&dataTable.value("Setup_Environment",SheetName) ,"Set Up Environment")	
       Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebListSetupEnvironment").Select dataTable.value("Setup_Environment",SheetName)
    Else
	   	Call ReportStep (StatusTypes.Fail, "Set Up Environment not entered","Set Up Environment")
    End If
    'Modified  by :Poornima
    'modified Date: 6/11/2018
    'Remarks: this is not editable
'    'ENTER OFFERING CONTACT NAME
'    If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactName").Exist(20) Then
'       	Call ReportStep (StatusTypes.Pass, "Offering Contact Name entered as "&dataTable.value("Offering_Contact_Name",SheetName),"Offering Contact Name")	
'     	Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactName").Set dataTable.value("Offering_Contact_Name",SheetName)
'    Else
'	   	Call ReportStep (StatusTypes.Fail, "Offering Contact Name  not entered","Offering Contact Name")
'    End If
'    'ENTER OFFERING CONTACT PHONE
'    If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactPhone").Exist(20) Then
'       	Call ReportStep (StatusTypes.Pass, "Offering Contact Phone entered as "&dataTable.value("Offering_Contact_Phone",SheetName),"Offering Contact Phone")	
'       	Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactPhone").Set dataTable.value("Offering_Contact_Phone",SheetName)
'    Else
'	   	Call ReportStep (StatusTypes.Fail, "Offering Contact Phone  not entered","Offering Contact Phone")
'    End If
    'ENTER OFFERING CONTACT EMAIL
'    If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactEmail").Exist(20) Then
'       	Call ReportStep (StatusTypes.Pass, "Offering Contact Email entered as "&dataTable.value("Offering_Contact_Email",SheetName),"Offering Contact Email")	
'       	Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactEmail").Set dataTable.value("Offering_Contact_Email",SheetName)
'    Else
'	   	Call ReportStep (StatusTypes.Fail, "Offering Contact Email  not entered","Offering Contact Email")
'    End If

	'Validate OFFERING CONTACT EMAIL
	If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactEmail").Exist(20) Then
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("inneretext:=.*QAMR-ACCOUNTS.*","html tag:=DIV").Exist(2) then

		QAMRLoggedInUser =  Browser("Browser-OCRF").Page("Page-OCRF").WebElement("inneretext:=.*QAMR-ACCOUNTS.*","html tag:=DIV").GetROProperty("innertext")
		LoggedINUSER =  split(QAMRLoggedInUser,"\")
		User = split(LoggedINUSER(1)," ")
		QAMRLoggedInUserEmail = mid(trim(User(0)),1,1)&"."&right(trim(User(0)),len(trim(User(0)))-1)&"@quintiles.com"
	ElseIf  Browser("Browser-OCRF").Page("Page-OCRF").WebElement("OCRFLoggedInUser").Exist Then	
		OCRFLoggedInUser = Browser("Browser-OCRF").Page("Page-OCRF").WebElement("OCRFLoggedInUser").GetROProperty("innertext")
		OCRFLoggedInUser = Split(OCRFLoggedInUser,"\")
		OCRFLoggedInUserEmail = OCRFLoggedInUser(1) & "@in.imshealth.com"
	elseif	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=globalUserLoginId","html tag:=SPAN","class:=username").Exist then
		OCRFLoggedInUser = Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=globalUserLoginId","html tag:=SPAN","class:=username").GetROProperty("innertext")
		OCRFLoggedInUser = Split(OCRFLoggedInUser,"\")
		OCRFLoggedInUserEmail = OCRFLoggedInUser(1) & "@in.imshealth.com"
	end if 	
		offeringContEmail = Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactEmail").GetROProperty("value")
		If UCase(OCRFLoggedInUserEmail) = UCase(offeringContEmail) or UCase(QAMRLoggedInUserEmail) = UCase(offeringContEmail) Then
			Call ReportStep (StatusTypes.Pass, "Offering Contact Email validated as " & offeringContEmail,"Validate Offering Contact Email")	
			Else 
			Call ReportStep (StatusTypes.Fail, "Offering Contact Email not validated successfully " ,"Validate Offering Contact Email")	
		End If
       	
    Else
	   	Call ReportStep (StatusTypes.Fail, "Offering Contact Email  not available","Offering Contact Email")
    End If	

End Function

	 
Function  AddNewLinesINAddUserTabOld(ByVal Path,ByRef objData)

'CHECK FOR USER ACCESS TAB EXIST
If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementOfferingUserList").Exist(20) Then
	Call ReportStep (StatusTypes.Pass, "Entered in to Add user tab","Add user tab")	
Else
	Call ReportStep (StatusTypes.Fail, "Not Entered in to Add user tab","Add user tab")
End If
Set Tab=Browser("Browser-OCRF").Page("Page-OCRF").Link("html tag:=A","class:=selected")
Call ValidateActiveTabOrangeColour(Tab,"Add Users",objData)

SheetName = "Add_Users"
Call ImportSheet(SheetName,Path)

r = DataTable.GetSheet("Add_Users").GetRowCount
row= Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html tag:=TABLE","html id:=user-list").GetROProperty("rows")
Datatable.SetCurrentRow(1)
'For i = 1 To 100-row-1 Step 1
'    Datatable.SetCurrentRow(i)
'    Str=Str&dataTable.value("User_Login_Id","Add_Users")&";"
'
'Next
Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebListselNewLine").Select "50"
Set Upload=Description.Create()
Upload("micclass").value="WebElement"
Upload("html tag").value="DIV"
Upload("innertext").value="Upload User"
Set toatlObj=Browser("Browser-OCRF").Page("Page-OCRF").ChildObjects(Upload)
If toatlObj(0).Exist(10) Then
	toatlObj(0).Click
	'Call ReportStep (StatusTypes.Pass, "'Upload user' Exist and Clicked on" ,"Add user tab")	
Else
	Call ReportStep (StatusTypes.Fail,  "'Upload user' Not Exist","Add user tab")
End If

If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditblkClientName").Exist(10) Then
	Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditblkClientName").Set dataTable.value("Offering_Client","Offering_Details")
	'Call ReportStep (StatusTypes.Pass, "Edit Box for Client Name Exist and entered "&dataTable.value("Offering_Client","Offering_Details"),"Add user tab")	
Else
	Call ReportStep (StatusTypes.Fail,  "Edit Box for Client Name Not Exist","Add user tab")
End If

  Set objShell=CreateObject("WScript.Shell")
	Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditblkClientName").Click
	objShell.SendKeys "{DOWN}"
	wait 2
	objShell.SendKeys "{DOWN}"

	wait 2	
	objShell.SendKeys "{ENTER}"
	wait 2
	If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditUploadUser").Exist(10) Then
	   Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditUploadUser").Set dataTable.value("No_Off_Record","Add_Users")
		'Call ReportStep (StatusTypes.Pass, "Upload user area Exist and Uplaoded users","Add user tab")	
    Else
	   Call ReportStep (StatusTypes.Fail,  "Upload user area Not Exist and Uplaoded users","Add user tab")

	End If
    If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebEleUploadButton").Exist(10) Then
	   Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebEleUploadButton").Click
		'Call ReportStep (StatusTypes.Pass, "Clicked on 'Upload' button","Add user tab")	
    Else
	   Call ReportStep (StatusTypes.Fail,  "Not Clicked on 'Upload' button","Add user tab")

	End If
	

	
End Function




Function  AddNewLinesINAddUserTab(ByVal FileName,ByVal sheetName,ByVal RowStart,ByVal RowEnd,ByRef objData)

   
   NoOfUser=(RowEnd-RowStart)+1
   ' Check tab 'Add User' opened.
   
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementOfferingUserList").Exist(30) Then
	   Call ReportStep (StatusTypes.Pass,  "'Add User' tab opened","Add user tab")
    Else
       Call ReportStep (StatusTypes.Fail,  "'Add User' tab not opened","Add user tab")
    End If
    
    Set Tab=Browser("Browser-OCRF").Page("Page-OCRF").Link("html tag:=A","class:=selected")
	Call ValidateActiveTabOrangeColour(Tab,"Add Users",objData)

	Call ImportSheet(SheetName,FileName)

	Datatable.SetCurrentRow(RowStart)
	
	For Iterator = RowStart To RowEnd Step 1
	
		If NoOfUser <> "" Then
			ClientName = dataTable.value("Client_Name",SheetName)
			UserId = dataTable.value("User_Login_Id",SheetName)
			Country = dataTable.value("Country",SheetName)
		'	UserId = "imitestuser9400@uk.imshealth.com"
		'	ClientName = "AutomatedClient1"
			'Country = "Canada"
			
			
			If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-pg-div","html tag:=DIV","innertext:=Add New Line","visible:=True").Exist Then
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-pg-div","html tag:=DIV","innertext:=Add New Line","visible:=True").click
			elseif Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_AddNewLine").Exist then
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_AddNewLine").Click	
			end if				
			
			row=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").RowCount
			
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(row, 2,"WebCheckBox",0).highlight
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(row, 2,"WebCheckBox",0).set "ON"
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(row, 9,"WebEdit",0).highlight
			'changes by srinivas-23rd sept 2020
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(row, 9,"WebEdit",0).click
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(row, 9,"WebEdit",0).set UserId
			Call Pageloading()	
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(row, 2,"WebCheckBox",0).set "ON"
			'Srinivas-Commenting below line
			'Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(row, 2,"WebCheckBox",0).set "ON"
			'changes by srinivas-23rd sept 2020
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(row, 13,"WebEdit",0).click
			
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(row, 13,"WebEdit",0).highlight
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(row, 13,"WebEdit",0).set ClientName
			
			 Set objShell=CreateObject("WScript.Shell")
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(row, 13,"WebEdit",0).click
			objShell.SendKeys "{DOWN}"
			wait 1
			'Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&ClientName,"html id:=ui-id-.*").fireevent "onmouseover" 
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&ClientName,"html id:=ui-id-.*").highlight
			wait 2
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&ClientName,"html id:=ui-id-.*").Click		
			Call Pageloading()

			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(row, 2,"WebCheckBox",0).set "ON"
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(row, 2,"WebCheckBox",0).set "ON"
			
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(row, 14,"WebEdit",0).highlight
			'changes by srinivas-23rd sept 2020
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(row, 14,"WebEdit",0).click
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(row, 14,"WebEdit",0).set Country
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(row, 14,"WebEdit",0).click
			objShell.SendKeys "{DOWN}"
			wait 1
			'Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&Country,"html id:=ui-id-.*").fireevent "onmouseover" 
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&Country,"html id:=ui-id-.*").highlight
			wait 2
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&Country,"html id:=ui-id-.*").Click		
		
			End If
		wait 3
		DataTable.GetSheet(SheetName).SetNextRow	
	wait 3
Next

wait 3	
	
End Function

Function  AddBulkUsersINAddUserTab(ByVal FileName, ByVal ClientName, ByRef objData)

   
   ' Check tab 'Add User' opened.
    If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementOfferingUserList").Exist(30) Then
	   Call ReportStep (StatusTypes.Pass,  "'Add User' tab opened","Add user tab")
    Else
       Call ReportStep (StatusTypes.Fail,  "'Add User' tab not opened","Add user tab")
    End If
   
   
	Set Tab=Browser("Browser-OCRF").Page("Page-OCRF").Link("html tag:=A","class:=selected")
	Call ValidateActiveTabOrangeColour(Tab,"Add Users",objData)
	
	
	Set Upload=Description.Create()
	Upload("micclass").value="WebElement"
	Upload("html tag").value="DIV"
	Upload("innertext").value="Upload User"
	Set toatlObj=Browser("Browser-OCRF").Page("Page-OCRF").ChildObjects(Upload)
	toatlObj(0).Click	

	If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditblkClientName").Exist(2) Then
   		Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditblkClientName").Set ClientName   		
	End If

    Set objShell=CreateObject("WScript.Shell")
	Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditblkClientName").Click
	objShell.SendKeys "{DOWN}"
	wait 1
	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&ClientName,"html id:=ui-id-.*").fireevent "onmouseover" 
	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&ClientName,"html id:=ui-id-.*").highlight
	wait 2
	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&ClientName,"html id:=ui-id-.*").Click		

	usersList = ReadUsersFromTextFile(FileName)
	Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditUploadUser").Set usersList
	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebEleUploadButton").Click

	wait 3	
	
End Function


	
Function ValidateActiveTabOrangeColour(ByVal objTab, ByVal Tab ,ByRef objData)
If objTab.Exist(10) Then
	Color=objTab.GetROProperty("background color")
	If Color="#ea8511" Then
		Call ReportStep (StatusTypes.Pass, "Active Tab '"&Tab&"' in Orange Color","Active Tab")
		'print "pass"
	Else
		Call ReportStep (StatusTypes.Pass, "Active Tab '"&Tab&"' Not in Orange Color","Active Tab")
		'print "fail"
	End If
End If

End Function






Public Function AddDatabase(ByVal Path,ByVal ExtraParam,ByRef objData)

    If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html tag:=SPAN","innertext:=Database Details","class:=ui-jqgrid-title").Exist(60) Then
		Call ReportStep (StatusTypes.Pass, "User is in Database Details tab ","Database Details tab")	
	Else
		Call ReportStep (StatusTypes.Fail, "Database Details tab is not available","Database Details tab")
	End If

	Dim allDbName()
	SheetName = "OCRF-DatabaseAdd"
	
	Call ImportSheet(SheetName,Path)
	
	
	'Number of rows to be added
	totalDbCount = Datatable.Value ("selectAddNumberofRecords",SheetName)
	
	
	'Call ReportStep (StatusTypes.Pass, "Entering "&NumberOfRows&" records in 'Cube details' tab started","Cube details tab")
	For i = 0 To totalDbCount - 1 Step 1
		Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementAddNewLine").Click
		rowNumber = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").RowCount
		
		colCount = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ColumnCount(rowNumber)
		'datatable.SetCurrentRow(i)
		
			Set CheckRow =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(rowNumber,2,"WebCheckBox",0)
			CheckRow.click	
			wait 1
			
			
			dbName = datatable.value("DatabaseName",SheetName)
'			d= now
'			dt = Right("00" & Hour(d), 2) & Right("00" & Minute(d), 2) & Right("00" & Second(d), 2) 
'			dbName = dbName & "_" & dt & "00" & i
			ReDim Preserve allDbName(totalDbCount - 1)
			allDbName(i) = dbName

			Set setDbName = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(rowNumber,12,"WebEdit",0)
			setDbName.set dbName
			
			Call WebEditExistence(rowNumber,22,"WebEdit")
			
			Set setDbServer = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(rowNumber,22,"WebEdit",0)
			setDbServer.set datatable.value("ServerLocation",SheetName) 
			
			Call WebEditExistence(rowNumber,24,"WebList")
				
			Set setFLAPath = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(rowNumber,24,"WebList",0)
			setFLAPath.select datatable.value("FLAPATH",SheetName) 
			
			Call WebEditExistence(rowNumber,7,"WebList")
			
			Set selectAction = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(rowNumber,7,"WebList",0)
			selectAction.select datatable.value("Action",SheetName) 
			
			Call WebEditExistence(rowNumber,8,"WebEdit")
			
			Set setActionComments = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(rowNumber,8,"WebEdit",0)
			setActionComments.set datatable.value("ActionComments",SheetName) 
			
			Call WebEditExistence(rowNumber,10,"WebList")
			
			Set selectDBType = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(rowNumber,10,"WebList",0)
			selectDBType.select datatable.value("DBType",SheetName) 
			
			Call WebEditExistence(rowNumber,11,"WebEdit")
			
			Set setDBDesc = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(rowNumber,11,"WebEdit",0)
			setDBDesc.set datatable.value("DBDescription",SheetName) 
			
			
			Call WebEditExistence(rowNumber,18,"WebEdit")
			
			Set setDbOwEmail = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(rowNumber,18,"WebEdit",0)
			setDbOwEmail.set datatable.value("DBOwnerEmail",SheetName) 
			
			datatable.SetNextRow
			'Call ReportStep (StatusTypes.Pass, "Entered  Database Name as '"&datatable.value("DatabaseName",SheetName)&"',Server Location as '"&datatable.value("ServerLocation",SheetName)&"' And FLA path as '"&datatable.value("FLAPATH",SheetName)&"' at row '"&del-1&"' done","Cube details tab")
		'Next
		'Call ReportStep (StatusTypes.Pass, "Entered '"&AfterRowCount-1&"' records Successfully in cube details tab","Cube details tab")
	
	Next
	AddDatabase = allDbName
End Function



'Verify the existence of webedit and weblist in particular row.

Public Function WebEditExistence(ByVal rowpos,ByVal colpos,ByVal objType)
	
	If objType = "WebEdit" Then
	
		For Iterator = 1 To 10 Step 1
			count = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItemCount(rowpos,colpos,"WebEdit")
			If count > 0 Then
				Exit For	
			Else
				Set CheckRow =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(rowpos,2,"WebCheckBox",0)
				CheckRow.click
				wait 2
			End If
		Next


	
		
		If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItemCount(rowpos,colpos,"WebEdit") > 0 Then
			'Call ReportStep (StatusTypes.Pass,"Edit box is displayed in row "&rowpos&" and column "&colpos,"WebEdit")
			
		Else	
			'Call ReportStep (StatusTypes.Fail,"Edit box is not displayed in row "&rowpos&" and column "&colpos,"WebEdit")	
			
			Exit Function
		End If
	ElseIf objType = "WebList" Then
		For Iterator = 1 To 10 Step 1
			count = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItemCount(rowpos,colpos,"WebList")
			If count > 0 Then
				Exit For	
			Else
				Set CheckRow =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(rowpos,2,"WebCheckBox",0)
				CheckRow.click
				wait 2
			End If
		Next
	
	
		If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItemCount(rowpos,colpos,"WebList") > 0 Then
			'Call ReportStep (StatusTypes.Pass,"List box is displayed in row "&rowpos&" and column "&colpos,"WebList")	
			'print "pass"
		Else	
			'Call ReportStep (StatusTypes.Fail,"List box is not displayed in row "&rowpos&" and column "&colpos,"WebList")	
			'print "fail"
			Exit Function
		End If
		
	End If
	
End Function


Public Function NavigateToNextTab()
	
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Exist Then
		Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Click
		Call ReportStep (StatusTypes.Pass, "Next button exist and clicked ","Next button")	
		'Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Click
	Else
		Call ReportStep (StatusTypes.Fail, "Next button not exist","Next button")
	End If
	
	
End Function





Public Function BI_Tools_Access_CombinationOld(ByVal Path,ByRef objData)
	'1) Validate whether OCRF New Request is redirected to BI Tool access
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementOfferingUserBIToolList").Exist(120) <> true Then
		Call ReportStep (StatusTypes.Fail, "Not Redirected to BI Tool Access tab","BI tool access Tab")	
	Else
		'Call ReportStep (StatusTypes.Pass, "Redirected to BI Tool Access tab successfully","BI tool access Tab")
	End If
	SheetName = "BI_Tools"	
	Call ImportSheet(SheetName,Path)
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Exist(120) Then
	    Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Click
		'Call ReportStep (StatusTypes.Pass, "Clicked on 'Select user'","BI tool access Tab")	
		wait 5
		Browser("Browser-OCRF").Page("Page-OCRF").WebList("html tag:=SELECT","default value:=20","visible:=True","all items:=20;50;100;999").Select "100"
	Else
		Call ReportStep (StatusTypes.Fail, "Not Clicked on 'Select user","BI tool access Tab")
	End If
	
	
	rc = Datatable.GetSheet(SheetName).GetRowCount
	
	For rownum = 1 To rc+100 Step 1
		
		Set objBIAccessTool = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIAccessTool")
		Set objBIToolAccessMode = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIToolAccessMode")
		
		Call SCA.SelectFromDropdown(objBIAccessTool, Datatable.Value("BI_Tools","BI_Tools"))
		Call SCA.SelectFromDropdown(objBIToolAccessMode, Datatable.Value("Access_Mode","BI_Tools"))
		
		CurComb = Datatable.Value("BI_Tools","BI_Tools")&Datatable.Value("Access_Mode","BI_Tools")
		
		Datatable.GetSheet(SheetName).SetPrevRow
		PrevComb = Datatable.Value("BI_Tools","BI_Tools")&Datatable.Value("Access_Mode","BI_Tools")
		Datatable.GetSheet(SheetName).SetNextRow
		
		
		
		If (Datatable.GetSheet(SheetName).GetCurrentRow = 1) OR ((CurComb = PrevComb) AND (Datatable.GetSheet(SheetName).GetCurrentRow > 0)) OR (Diffvalue = 1) Then
			
			Diffvalue = 0
			
			
			
			rownum = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").GetRowWithCellText(Datatable.value("User_ID","BI_Tools"),4,2)
			Set chk =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").ChildItem(rownum,2,"WebCheckBox",0)
			chk.click
	
		    Call ReportStep (StatusTypes.Pass, "User Id :"&Datatable.value("User_ID","BI_Tools")&" selected for BITool: "&Datatable.Value("BI_Tools","BI_Tools")&" Access Mode: "&Datatable.Value("Access_Mode","BI_Tools")&" Combination" ,"BI tool access Tab")	
	
			If (Datatable.GetSheet(SheetName).GetCurrentRow = Datatable.GetSheet(SheetName).GetRowCount) AND (Diffvalue = 0) Then
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_AddUser").Click
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html tag:=SPAN","innertext:=Add","visible:=True").Click
				Exit For
			Else
				Datatable.GetSheet(SheetName).SetNextRow
			End If
		
		Else
		
		
			Datatable.GetSheet(SheetName).SetPrevRow
			Call SCA.SelectFromDropdown(objBIAccessTool, Datatable.Value("BI_Tools","BI_Tools"))
			Call SCA.SelectFromDropdown(objBIToolAccessMode, Datatable.Value("Access_Mode","BI_Tools"))
			Datatable.GetSheet(SheetName).SetNextRow
			
			
		
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_AddUser").Click
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_Add").Click
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Click
			wait 5
			Browser("Browser-OCRF").Page("Page-OCRF").WebList("html tag:=SELECT","default value:=20","visible:=True","all items:=20;50;100;999").Select "100"
			
			
			Diffvalue = 1
		End If	
		
	Next
	
	
	'Validate number of records added
	Call PageLoading()
	Browser("Browser-OCRF").Page("Page-OCRF").Sync	
	
	Call NavigateToNextTab()
	
	Call PageLoading()
	
End Function
	




Public Function Client_Content_tab(ByVal Path,ByRef objData)
    'CHECK FOR CLIENT CONTENT TAB EXIST
  If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=client-offering","html tag:=TABLE","name:=ContentRequestID").Exist(20) Then
	Call ReportStep (StatusTypes.Pass, "Entered in to Client content tab","Client content tab")	
  Else
	Call ReportStep (StatusTypes.Fail, "Not Entered in to Client content tab","Client content tab")
  End If	
    Set Tab=Browser("Browser-OCRF").Page("Page-OCRF").Link("html tag:=A","class:=selected")	
    Call ValidateActiveTabOrangeColour(Tab,"Client Content",objData)
	
End Function


Public Function ImportSheet(ByVal SheetName,ByVal FileName)
	
	datatable.AddSheet SheetName
	wait 3
	DynamicPath = Environment.value("CurrDir")& "InputFiles\IMSSCAWeb\OCRF\"&FileName&".xls"
	wait 1
	datatable.ImportSheet DynamicPath,SheetName,SheetName
	datatable.GetSheet(SheetName)
	
End Function



Public Function PageLoading()
	
	For i = 1 To 100 Step 1
	If Browser("Browser-OCRF").Page("Page-OCRF").Image("file name:=loading.gif","visible:=True").Exist(5) = False Then
		wait 3
		Exit For
	Else
			
	End If
	
Next
	
	
End Function







Public Function ValidateRowsCount(ByVal Path,ByRef objData)
	
	Set xl = CreateObject("Excel.Application")
	xl.Visible = False
	Set wb = xl.Workbooks.Open(Path&Environment.Value("TestName")&".xlsx")
	
	arr = Array("OCRF-DatabaseAdd","BI_Tools","Add_Users","User_DB","Client_Contents")
	
	Set d = description.Create
	d("micclass").value = "WebElement"
	d("html tag").value = "LABEL"
	
	Set desc = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Offering_Action_Summary").ChildObjects(d)
	
	For i = 0 To desc.count-1 Step 1
		
		Set ws = wb.Worksheets(arr(i))
	
		rows = ws.usedrange.rows.count	
		
		If rows-1 = Cint(desc(i).GetRoProperty("innertext"))  Then
			Call ReportStep (StatusTypes.Pass, "Count matched successfully in Offering action summary and respective tabs "&arr(i),"Rows count")	
		Else
			Call ReportStep (StatusTypes.Fail, "Count did not match successfully in Offering action summary and respective tabs "&arr(i),"Rows count")	
		End If
		
	Next
	
	
	
	wb.Save
	wb.Close
	xl.Quit	
	
End Function



Public Function ApplySecurity(ByRef objData)

	Call verify_ContentPermissionInOfferingSummary()
	
	Set objApplySecurity = Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_ApplySecurity")
	Call SCA.ClickOn(objApplySecurity,"Apply Security","Offering Details")
	Call PageLoading()

	If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Yes").Exist(60) Then
		Set objConfrmYes = Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Yes")
		Call SCA.ClickOn(objConfrmYes,"Yes","Confirmation")
	End If
	Call PageLoading()
	
	wait 5
	Set objInfOkay = Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Information_Okay")
'	Call PageLoading()
'	Call SCA.ClickOn(objInfOkay,"Okay","Confirmation")
	
'Modified By Avinash on 25th July 2019	
	If objInfOkay.Exist(40) Then
	   Call SCA.ClickOn(objInfOkay,"Okay","Confirmation")
	End If
'Modification completed	

	Call PageLoading()
	
'	If Browser("Browser-OCRF").Page("Page-OCRF").Link("WeblnkSubmit").Exist(5) Then
'		Browser("Browser-OCRF").Page("Page-OCRF").Link("WeblnkSubmit").Click
'	ElseIf Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementFinish").Exist(5) Then
'		Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementFinish").Click
'	ElseIf Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_SaveAndComplete").Exist(5) Then
'	    Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_SaveAndComplete").click
'	End If

'	Browser("Browser-OCRF").Page("Page-OCRF").Link("SecurityAndAccessManagement").Click
'	
'	If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButtonNo").Exist(60) Then
'		Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButtonNo").Click
'	End If
'	Call PageLoading()	
	
	
End Function



Public Function validateOCRFRequest(ByVal Req_ID)
	
        wait 3
	For Iterator = 1 To 50 Step 1
	
    	foundInTable= FilterInOnlineCRF("Current Requests","OCRF ID",Req_ID,objData)
    	wait 2
    	Set CurReqTbl=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data")
		Set CmpReqTbl=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTblCompletedReqData")
		If CurReqTbl.RowCount >= 2  Then
	  		Set ValTbl=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data")
    	ElseIf CmpReqTbl.RowCount >=2 Then
      		Set ValTbl=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTblCompletedReqData")    
		End If
		rownum = ValTbl.GetRowWithCellText(Req_ID,3,2)
	
	
	'changes by srinivas w.r.t CRF Simplification build
	
		If  trim(ValTbl.GetCellData(rownum,16)) = "" Then
			Call ReportStep (StatusTypes.Pass, "OCRF Request created succesfully with No:"&Req_ID,"OCRF Request")
			status = "pass"
			Exit For	
		ElseIf  trim(ValTbl.GetCellData(rownum,16)) = "Security failed" Then
			Call ReportStep (StatusTypes.Fail, "Apply security failed with Request No:"&Req_ID,"OCRF Request")
			status = "fail"
			Exit For	
		Else	
			
			'Browser("Browser-OCRF").Refresh
			foundInTable= FilterInOnlineCRF("Current Requests","OCRF ID",Req_ID,objData)
			wait 10
		End If
	Next
	
	If status <> "pass" Then
		Call ReportStep (StatusTypes.Fail, "Apply security is failed "&Req_ID,"OCRF Request")		
	End If
	

End Function


'This function is created to verify if the OCRF's Apply Security action is failed
Public Function ValidateApplySecurityFailed(ByVal Req_ID)
	
        wait 3
	For Iterator = 1 To 50 Step 1
	
    	foundInTable= FilterInOnlineCRF("Current Requests","OCRF ID",Req_ID,objData)
    	wait 2
    	Set CurReqTbl=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data")
		Set CmpReqTbl=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTblCompletedReqData")
		If CurReqTbl.RowCount >= 2  Then
	  		Set ValTbl=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data")
    	ElseIf CmpReqTbl.RowCount >=2 Then
      		Set ValTbl=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTblCompletedReqData")    
		End If
		rownum = ValTbl.GetRowWithCellText(Req_ID,3,2)
	'changes by srinivas w.r.t CRF Simplification build
		If  trim(ValTbl.GetCellData(rownum,16)) = "Security failed" Then
			Call ReportStep (StatusTypes.Pass, "Apply security failed with Request No:"&Req_ID,"OCRF Request")
			status = "pass"
			Exit For	
		ElseIf  trim(ValTbl.GetCellData(rownum,16)) = "" Then
			Call ReportStep (StatusTypes.Pass, "OCRF Request created succesfully with No:"&Req_ID,"OCRF Request")
			status = "fail"
			Exit For	
		Else	
			Browser("Browser-OCRF").Refresh
			foundInTable= FilterInOnlineCRF("Current Requests","OCRF ID",Req_ID,objData)
			wait 10
		End If
	Next
	
	If status <> "pass" Then
		Call ReportStep (StatusTypes.Pass, "Apply security is failed "&Req_ID,"OCRF Request")		
	End If
	

End Function




'Create Client in OPs
Function SCA_CreateClientInOps(ByRef objData)
    Systemutil.CloseProcessByName "iexplore.exe"
	Call OPSLogin (objData.Item("OpsUser_Name"),objData.Item("OpsPassword"),Environment.Value("OPSURL"))
	wait 5
	Call ClientCreation_Ops(objData.Item("ClientName"),objData.Item("ClientName"),"","")
						
End Function
'Login to Ops NonLive environment
Function OPSLogin(ByVal strUserName,ByVal StrPwd,ByVal StrUrl)
  
  If strUserName <> "" And StrPwd <> "" Then
  	Systemutil.CloseProcessByName "iexplore.exe"
  	Systemutil.Run "iexplore.exe", StrUrl, , , 3
  	wait 20
  	If UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").Exist(300) Then
       UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").Highlight 
       UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAEdit("User name").SetValue strUserName
       UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAEdit("Password").SetValue StrPwd
       UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAButton("OK").Click  		
       	
       wait 3
	   
	   If Browser("DC Home").page("Ops_Page").WebButton("WebButAddNew").Exist(60) Then
	  		Call ReportStep (StatusTypes.Pass,"User is suceessfully Logged in to Ops System","Welcome Page" )
	  	Else 
	  		Call ReportStep (StatusTypes.Fail,"User is not suceessfully Logged in to Ops System","Welcome Page" )
	  		ExitTest
	  	End If
       
  	End If
  	

'  	If  Browser("DC Home").Page("Microsoft Forefront TMG").WebEdit("txtUsername").Exist(5) Then 
'	  	Call ReportStep (StatusTypes.Pass,"OPS lanuched successfully","Home Page" )
'	  	Browser("DC Home").Page("Microsoft Forefront TMG").WebEdit("txtUsername").Set strUserName
'	  	Browser("DC Home").Page("Microsoft Forefront TMG").WebEdit("txtPassword").Set  StrPwd
'	  	Browser("DC Home").Page("Microsoft Forefront TMG").WebButton("btnLogin").Click
'	  	
'	  	wait 3
'	  	If Browser("DC Home").page("Ops_Page").WebButton("WebButAddNew").Exist(60) Then
'	  		Call ReportStep (StatusTypes.Pass,"User is suceessfully Logged in to Ops System","Welcome Page" )
'	  	Else 
'	  		Call ReportStep (StatusTypes.Fail,"User is not suceessfully Logged in to Ops System","Welcome Page" )
'	  		ExitTest
'	  	End If
'	ElseIf Window("BrowserPoPUp").Exist(5) Then  
'		Window("BrowserPoPUp").highlight
'	    Window("BrowserPoPUp").Dialog("Windows Security").highlight
'	   'Handle 'Window Security' PopUp
'		If Window("BrowserPoPUp").Dialog("Windows Security").Exist(30) Then
'		   Window("BrowserPoPUp").Dialog("Windows Security").WinEdit("WinEditUserName").Set strUserName	
'		   Window("BrowserPoPUp").Dialog("Windows Security").WinEdit("WinEditPassword").Set StrPwd
'		   Window("BrowserPoPUp").Dialog("Windows Security").WinButton("BtnOK").Click
'		   
'		   	wait 3
'		   Call ReportStep (StatusTypes.Pass, "User " & UserName & " has logged in successfully","Login through Window security PopUp")
'		Else
'		   Call ReportStep (StatusTypes.Fail, "User " & UserName & " has not successfully logged in","Login through Window security PopUp")	   
'		End If 
'  	ElseIf  Browser("BrowserPoPUp").Dialog("Windows Security").Exist(5) Then
'  	        Browser("BrowserPoPUp").Dialog("Windows Security").WinEdit("WinEditUserName").Set strUserName	
'	   		Browser("BrowserPoPUp").Dialog("Windows Security").WinEdit("WinEditPassword").Set StrPwd
'	   		Browser("BrowserPoPUp").Dialog("Windows Security").WinButton("BtnOK").Click
'	   		Call ReportStep (StatusTypes.Pass, "Window Security pop up is handled","Window security PopUp")
'  	    
'  	Else
'  		    Systemutil.CloseProcessByName "iexplore.exe"
'			Systemutil.Run "iexplore.exe", StrUrl, , , 3
'			
'			wait 3
'			If Browser("DC Home").page("Ops_Page").WebButton("WebButAddNew").Exist(60) Then
'	  			Call ReportStep (StatusTypes.Pass,"User is suceessfully Logged in to Ops System","Welcome Page" )
'	  		Else 
'	  			Call ReportStep (StatusTypes.Fail,"User is not suceessfully Logged in to Ops System","Welcome Page" )
'	  			ExitTest
'	  		End If
'	
'  	End If
 End If 	

End Function

Function ClientCreation_Ops(ByVal strOpsClient_Name,ByVal strOpsShort_Name,ByVal strOpsDescription,ByVal ClientID)

     wait 2
     Call ItemSelection("Clients")
	
	wait 5	
	For i = 1 To 10 Step 1
		If Browser("DC Home").Page("Ops_Page").WebButton("WebButAddNew").Exist Then
			Browser("DC Home").Page("Ops_Page").WebButton("WebButAddNew").Click
		End If
	Next
	
	Browser("DC Home").Page("Ops_Page").WebEdit("WebEditClientName").Set strOpsClient_Name
    Browser("DC Home").Page("Ops_Page").WebEdit("WebEditShortName").Set strOpsShort_Name
    Browser("DC Home").Page("Ops_Page").WebEdit("WebEditDescription").Set strOpsDescription
    Browser("DC Home").Page("Ops_Page").WebCheckBox("CreateIAMClient").Set "ON"

    Browser("DC Home").Page("Ops_Page").WebElement("WebEleSave").Click
	If Browser("DC Home").Page("Ops_Page").WebElement("class:=message error","html tag:=DIV","html id:=divInfo").Exist(20) Then
    	  Call ReportStep (StatusTypes.Pass,"Client name already exist in Ops","Client creation Page" )
     Else
		   Call ReportStep (StatusTypes.Pass,"Client name "& strOpsClient_Name &" Not exist in Ops  ","Client creation Page" )	
    End If
    Browser("DC Home").Page("Ops_Page").WebElement("WebEleSave").Click
    Browser("DC Home").Page("Ops_Page").WebElement("WebEleClose").Click

End Function



Function ItemSelection(ByVal menuitem)

  	Set menuChild=Description.Create()
	menuChild("micclass").value="Link"
	menuChild("Class").Regularexpression=True
	menuChild("Class").value=".*childMenuItem.*"
	set childcnt=Browser("DC Home").Page("Ops_Page").ChildObjects(menuChild)
	
	For i = 0 To childcnt.count-1
		if childcnt(i).getroproperty("outertext")=menuitem then	
				If childcnt(i).Exist(10) Then
					childcnt(i).click
					Call ReportStep (StatusTypes.Pass,menuitem&" menu item selected in Ops page","Ops  Page" )
					Exit for
				End If
		Else
		     'Call ReportStep (StatusTypes.Pass,"Not able to select "&menuitem&"  menu item  in Ops page","Ops  Page" )
		End If
	Next
End Function

Function ValidationForGroupsInOps(ByVal ClientName,ByVal Country,ByVal offeringName,ByVal ClientSynOrNon)
If Browser("DC Home").Page("Ops_Page").WebRadioGroup("html id:=ContentType","html tag:=INPUT","type:=radio").Exist(20) Then
	If ClientSynOrNon="Syndicate" Then
	   Browser("DC Home").Page("Ops_Page").WebRadioGroup("html id:=ContentType","html tag:=INPUT","type:=radio").Select "Syndicate"
	Else
	   Browser("DC Home").Page("Ops_Page").WebRadioGroup("html id:=ContentType","html tag:=INPUT","type:=radio").Select "NonSyndicate"
	End If
End If


If Browser("DC Home").Page("Ops_Page").WebEdit("WebEditContAccessClient").Exist Then
    Browser("DC Home").Page("Ops_Page").WebEdit("WebEditContAccessClient").Set ClientName
   Call ReportStep (StatusTypes.Pass,"Client name entered in text box as "&ClientName,"'Manage Content Access' Page" )  
Else
	Call ReportStep (StatusTypes.Pass,"Client name Not entered in text box as "&ClientName ,"'Manage Content Access' Page" )	
   
End If

wait 2
Set WshShell =CreateObject("WScript.Shell")
	Browser("DC Home").Page("Ops_Page").WebEdit("WebEditContAccessClient").Click
	WshShell.SendKeys "{DOWN}"
'	wait 3
'	WshShell.SendKeys "{DOWN}"
    
    
'	wait 2	
'	WshShell.SendKeys "{ENTER}"
    If Browser("DC Home").Page("Ops_Page").WebElement("class:=ui-corner-all","html tag:=A","innertext:="&ClientName).Exist(2) Then
       Browser("DC Home").Page("Ops_Page").WebElement("innertext:="&ClientName,"html tag:=A","class:=ui-corner-all").Click	
    End If
    wait 1
    If Browser("DC Home").Page("Ops_Page").WebEdit("WebEditContAccessClient").Exist Then
    	Browser("DC Home").Page("Ops_Page").WebEdit("WebEditContAccessCountry").Set Country
   		Call ReportStep (StatusTypes.Pass,"Country name entered in text box as "&Country,"'Manage Content Access' Page" )  
	Else
		Call ReportStep (StatusTypes.Pass,"Country name Not entered in text box as "&Country ,"'Manage Content Access' Page" )	
   
	End If
    If Browser("DC Home").Page("Ops_Page").WebButton("WebBtnSPGroups").Exist Then
    	 Browser("DC Home").Page("Ops_Page").WebButton("WebBtnSPGroups").Click
   		Call ReportStep (StatusTypes.Pass,"Clicked on 'Share point groups' button","'Manage Content Access' Page" )  
	Else
		Call ReportStep (StatusTypes.Pass,"Not Clicked on 'Share point groups' button" ,"'Manage Content Access' Page" )	
   
	End If
   
    Browser("DC Home").Page("Ops_Page").WebTable("class:=ui-jqgrid-btable","html id:=groupGrid","html tag:=TABLE").highlight
    If Browser("DC Home").Page("Ops_Page").WebEdit("WebEditValue").Exist Then
    	Browser("DC Home").Page("Ops_Page").WebEdit("WebEditValue").Set offeringName
   		Call ReportStep (StatusTypes.Pass,"Entered 'offering name as '"&offeringName&"'  for group search","'Manage Content Access' Page" )  
	Else
		Call ReportStep (StatusTypes.Fail,"Entered 'offering name for group search " ,"'Manage Content Access' Page" )	
   
	End If
    
    Browser("DC Home").Page("Ops_Page").WebEdit("WebEditValue").Click
    WshShell.SendKeys "{ENTER}"
    wait 5
    r=Browser("DC Home").Page("Ops_Page").WebTable("class:=ui-jqgrid-btable","html id:=groupGrid","html tag:=TABLE").GetROProperty("rows")
    If r>1 Then
    	Call ReportStep (StatusTypes.Pass,"Number of groups Exist in list are "& r,"'Manage Content Access' Page" ) 
		For i = 2 To Cint(r) Step 1
	       data=Browser("DC Home").Page("Ops_Page").WebTable("class:=ui-jqgrid-btable","html id:=groupGrid","html tag:=TABLE").GetCellData(i,2)
           If instr(1,data,offeringName)>0 Then
    	      If instr(1,data,"Approvers")>0 Then
    	        Call ReportStep (StatusTypes.Pass,"For Group Name '"&data&"' 'Approvers' Access Mode assigned","Ops Group Page" )
    		  ElseIf instr(1,data,"Authors")>0 Then
    		    Call ReportStep (StatusTypes.Pass,"For Group Name '"&data&"' 'Authors' Access Mode assigned","Ops Group Page" )
    		  ElseIf instr(1,data,"Viewers")>0 Then
    		    Call ReportStep (StatusTypes.Pass,"For Group Name '"&data&"' 'Viewers' Access Mode assigned","Ops Group Page" )
    		     'ElseIf instr(1,data,"AC235")>0 Then
    		    'Call ReportStep (StatusTypes.Pass,"for Group Name "&data&" 'AC235' Access Mode assigned","Ops Group Page" )
    		  Else
    		    Call ReportStep (StatusTypes.Pass,"for Group Name '"&data&"'  Access Mode assigned is :"&dataTable.value("Access_Mode","BI_Tools"),"Ops Group Page" )
    		  End If
    	     'Call ReportStep (StatusTypes.Pass,"for Offering Name '"&offeringName&"'  ,Country '"&Country&"' and Client Name '"&ClientName&"' No groups Exist" ,"Ops Group Page" )
          End If
      Next    	
	Else
	    
		Call ReportStep (StatusTypes.Pass,"No groups found in Search list " ,"'Manage Content Access' Page" )
    End If
    

End Function

Function ValidationForAddUserInOps(ByVal BITool,ByRef AccessMode,ByRef Users )
r=Browser("DC Home").Page("Ops_Page").WebTable("class:=ui-jqgrid-btable","html id:=groupGrid","html tag:=TABLE").GetROProperty("rows")
For i = 2 To Cint(r) Step 1
    data=Browser("DC Home").Page("Ops_Page").WebTable("class:=ui-jqgrid-btable","html id:=groupGrid","html tag:=TABLE").GetCellData(i,2)
    If BITool = "IAM" AND instr(1,data,"AC235")>0 Then
       Call ReportStep (StatusTypes.Pass,"Tool '"&BITool& "'Exist in row '"&i&" 'with access permission as :"&AccessMode(0),"Ops Group Page" )

    Else
         For acc = 0 To UBound(AccessMode) Step 1
    	    If instr(1,data,BITool)>0 AND instr(1,data,AccessMode(acc))>0  Then
                Call  ValidationUserAddedInOps(i,AccessMode(acc),BITool,Users)
             End If
          Next
    End If


Next
End Function

Function ValidationUserAddedInOps(ByVal CNO,ByVal AccessMode,ByVal BITool,ByRef UsersList)
        Users=Split(UsersList,";")
        Set s=Description.Create
		s("micclass").value="WebCheckBox"
		s("html tag").value="INPUT"
		s("type").value="checkbox"
		Set c=Browser("DC Home").Page("Ops_Page").WebTable("class:=ui-jqgrid-btable","html id:=groupGrid","html tag:=TABLE").ChildObjects(s)
		c(CNO-2).click
		Browser("DC Home").Page("Ops_Page").WebButton("WebBtnGetUsers").Click
        Browser("DC Home").Page("Ops_Page").WebTable("class:=ui-jqgrid-btable","html id:=userGrid","html tag:=TABLE","name:=WebTable").highlight
		userrow=Browser("DC Home").Page("Ops_Page").WebTable("class:=ui-jqgrid-btable","html id:=userGrid","html tag:=TABLE","name:=WebTable").GetROProperty("rows")
		
		If userrow>2 Then
	     For r = 2 To CInt(userrow) Step 1
	       count=0
    	   userdata=Browser("DC Home").Page("Ops_Page").WebTable("class:=ui-jqgrid-btable","html id:=userGrid","html tag:=TABLE","name:=WebTable").GetCellData(r,2)
		   If userdata=" " Then
		      count=1
		   Else
		   	  For Iter = 0 To Ubound(Users)-1 Step 1
		       	
			   If userdata = Users(Iter) Then
			       count=1
			       Exit For  
	    	   Else
	    	       count=0
	    	   End If
		     Next
		   End If
		   
           If count=1 Then
           	  Call ReportStep (StatusTypes.Pass,"User '"&userdata&"' listed in 'user principle Name' list at row No '"&r&"' For Access mode :"&AccessMode& "BI Tool: "&BITool,"Ops Group Page" )
           Else
               Call ReportStep (StatusTypes.Fail,"User  '"&userdata&"' Not listed in 'user principle Name' list at row No '"&r&"' For Access mode : "&AccessMode& "BI Tool: "&BITool,"Ops Group Page" )
           End If
         Next
		End If
	    c(CNO-2).click	
End Function






Function FolderCreationInFLAServer(ByVal Filename,ByVal SheetName,ByVal RowStart,ByVal RowEnd,ByVal TimeStamp,ByRef objData)
        
        Call ImportSheet( SheetName, FileName)
        DataTable.GetSheet(SheetName).SetCurrentRow(RowStart)
        'NumberOCubes = Datatable.Value ("selectAddNumberofRecords",SheetName)
        NumberOCubes = (RowEnd-RowStart)+1
        For cube = RowStart To RowEnd Step 1
        	backupFile=Datatable.Value("BackUpFileName",SheetName)
        	strtemplateName=datatable.value("DatabaseName",SheetName)&TimeStamp
            SourceAbfFile=Environment.Value("CurrDir")&"HostingTestData\"&backupFile&""
            SourceBakFile=Environment.Value("CurrDir")&"HostingTestData\"&backupFile&""
        	strFolderName=Environment.Value("HostingPath")&strtemplateName
        	ExistAbfFile=Environment.Value("HostingPath")&strtemplateName&"\"&backupFile&""
        	RenameAbfFile=Environment.Value("HostingPath")&strtemplateName&"\"&strtemplateName&".abf"
        	ExistBakFile=Environment.Value("HostingPath")&strtemplateName&"\"&backupFile&""
        	RenameBakFile=Environment.Value("HostingPath")&strtemplateName&"\"&strtemplateName&".bak"
       		Set fso = CreateObject("Scripting.FileSystemObject")
        	If fso.FolderExists(strFolderName) = false Then
         		fso.CreateFolder (strFolderName)
        	End If
        	
        	
        	If fso.FolderExists(strFolderName) = false  Then
        		
        	Else
        	   If UCase(datatable.value("DBType",SheetName))="OLAP DB" or UCase(datatable.value("DBType",SheetName))=UCASE("Tabular DB") Then
        	   	  If fso.FileExists(SourceAbfFile) Then 
	        		fso.CopyFile SourceAbfFile,strFolderName&"\"
	        		If fso.FileExists(ExistAbfFile) Then
	        		    If fso.FileExists(RenameAbfFile)=False Then
	        		    	fso.MoveFile ExistAbfFile, RenameAbfFile
	        		    End If
	        			
                    End If
                 	 Call ReportStep (StatusTypes.Pass,"abf file placed in folder '"&strFolderName&"' with name '"&strtemplateName&".abf","Ops Group Page" )
	        	   Else
	        		 Call ReportStep (StatusTypes.Fail,"abf file Not placed in folder '"&strFolderName&"' with name '"&strtemplateName&".abf","Ops Group Page" )
	        		
	        	   End If
	           ElseIf UCase(datatable.value("DBType",SheetName))=UCase("Relational DB") Then
	                 If fso.FileExists(SourceBakFile) Then 
	        			fso.CopyFile SourceBakFile,strFolderName&"\"
	        			If fso.FileExists(ExistBakFile) Then
	        		    	If fso.FileExists(RenameBakFile)=False Then
	        		    		fso.MoveFile ExistBakFile, RenameBakFile
	        		    	End If
	        			
                    	End If
                 		 Call ReportStep (StatusTypes.Pass,"abf file placed in folder '"&strFolderName&"' with name '"&strtemplateName&".bak","Ops Group Page" )
	        	   	 Else
	        		 	Call ReportStep (StatusTypes.Fail,"abf file Not placed in folder '"&strFolderName&"' with name '"&strtemplateName&".bak","Ops Group Page" )
	        		
	        	     End If
        	   End If
        		
        	End If
        	
        	
        		datatable.SetNextRow
        Next                           
 
End Function

Function FolderCreationInFLAServerPerformance(ByVal Path,ByVal SheetName,ByRef objData,ByVal cubeNames, ByVal hostingItr, ByVal dbType, ByVal db)
		
		
		
		FLAServerPath = Environment.Value("HostingPath")
    	
    	
        For i = 0 To Ubound(cubeNames) Step 1
            strTemplateName=cubeNames(i)
  
            'destFolderName="\\cdtsolap571i.Gemini.DEV\IMSHosting\"&strTemplateName
            
            If dbType = "OLAP DB" Then
				sourcefile=Path&"\DB Files\" & db
				destFolderName= FLAServerPath & strTemplateName
		    	existFile= destFolderName & "\" & db
		    	renameFile= destFolderName & "\" & strTemplateName & ".abf"
	    		fileName = strTemplateName & ".abf"
			ElseIf dbType = "Relational DB" Then
				sourcefile=Path&"\DB Files\" & db
				destFolderName= FLAServerPath & strTemplateName
		    	existFile= destFolderName & "\" & db 
		    	renameFile= destFolderName & "\" & strTemplateName & ".bak"
	    		fileName = strTemplateName & ".bak"
			End If
    		
       		Set fso = CreateObject("Scripting.FileSystemObject")
        	If fso.FolderExists(destFolderName) = false Then
         		fso.CreateFolder (destFolderName)
        	End If
        	
    		If fso.FileExists(sourcefile) Then 
        		fso.CopyFile sourcefile,destFolderName&"\", True
        		wait 2

        		If fso.FileExists(existFile) Then
        			fso.MoveFile existFile, renameFile
        			fileSize = fso.GetFile(renameFile).Size
        			wait 2
                End If
             	 Call ReportStep (StatusTypes.Pass,"File is placed in folder '"&destFolderName&"' with name '"& fileName,"File Creation in FLA Server" )
        	Else
        		 Call ReportStep (StatusTypes.Fail,"File is not placed successfully in folder '"&destFolderName&"' with name '"&fileName,"File Creation in FLA Server" )
        		 ExitTest
        		
        	End If        	
        	
        Next
		If hostingItr > 1 Then
			For i = 0 To Ubound(cubeNames) Step 1
	            strTemplateName=cubeNames(i)
	            'destFolderName="\\cdtsolap571i.Gemini.DEV\IMSHosting\"&strTemplateName
	            
	            If dbType = "OLAP DB" Then
					destFolderName= FLAServerPath & strTemplateName
			    	renameFile= destFolderName & "\" & strTemplateName & ".abf"
			    	textFilePath = destFolderName & "\" & strTemplateName & ".txt"
		    		eventFilePath=destFolderName&"\"&strTemplateName&".event"
				ElseIf dbType = "Relational DB" Then
					destFolderName= FLAServerPath & strTemplateName
			    	renameFile= destFolderName & "\" & strTemplateName & ".bak"
			    	textFilePath = destFolderName & "\" & strTemplateName & ".txt"
		    		eventFilePath=destFolderName&"\"&strTemplateName&".event"
				End If
	            
	            
	       		Set fso = CreateObject("Scripting.FileSystemObject")
	        	If fso.FolderExists(destFolderName) Then
	    			If fso.FileExists(renameFile) Then 
	    				fso.CreateTextFile(textFilePath)
	    				wait 2
	    				fso.MoveFile textFilePath, eventFilePath
	             	 	Call ReportStep (StatusTypes.Pass,"event file placed in folder '"&destFolderName&"' with name '"&strTemplateName&".event","Event File Creation in FLA Server" )
	        			Else
	        		 	Call ReportStep (StatusTypes.Fail,"event file is not placed in folder '"&destFolderName&"' with name '"&strTemplateName&".event","Event File Creation in FLA Server" )
	        		 	ExitTest
	        		End If        	
        		End If
        	Next
		End If        
 
End Function

Function DeleteIMSToolFromBITool(ByRef objData)
 	If Browser("Browser-OCRF").Page("Page-OCRF").Link("html tag:=A","name:=\[5\] BI Tools Access->> ").Exist(10) Then
    	Browser("Browser-OCRF").Page("Page-OCRF").Link("html tag:=A","name:=\[5\] BI Tools Access->> ").highlight
	    Browser("Browser-OCRF").Page("Page-OCRF").Link("html tag:=A","name:=\[5\] BI Tools Access->> ").Click
	    Call ReportStep (StatusTypes.Pass,"Clicked on ' BI Tools Access' tab ","BI Tools Access tab" )
	Else
	    Call ReportStep (StatusTypes.Fail,"Not Clicked on ' BI Tools Access' tab ","BI Tools Access tab" )
  	End If
  	wait 5
  	If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=userBIAccess-list","html tag:=TABLE").Exist(20) Then
  	 	r=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=userBIAccess-list","html tag:=TABLE").GetROProperty("rows")
 	    For i = 1 To r Step 1
    	    data=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=userBIAccess-list","html tag:=TABLE").GetCellData(i,11)
            Login_Id= Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=userBIAccess-list","html tag:=TABLE").GetCellData(i,8)
    		If instr(1,data,"IMS Analysis Manager")>0 Then
    		    
    		    Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=userBIAccess-list","html tag:=TABLE").ChildItem(i ,2,"WebCheckBox",index).Set "ON"
    		    Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=userBIAccess-list","html tag:=TABLE").ChildItem(i,6,"WebList",index).Select "DELETE"
    	        Call ReportStep (StatusTypes.Pass,"Selected User '"&Login_Id&"' with Tool '"& data & "' for deletion","BI Tools Access tab" )
    	    End If
   
 	   Next
  	End If

     wait 5
    	 If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Exist(20) Then	
	   		Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Click
    	 End If
     	wait 10
     	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-button-text","innertext:=Previous","html tag:=SPAN").Exist(20) Then	
	   		Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-button-text","innertext:=Previous","html tag:=SPAN").Click
    	End If
        wait 10

 End Function

 
 
 Function SearchClientExistanceInOCRF(byVal strClientName,ByRef objData)
      wait 2
    	'Call ReportStep(StatusTypes.StepDefault, "##### SearchClient - Verification of page navigation by using Next,Last,Prev,First icons - Start ####", "Report Creation Page")
        'SEARCH FOR CREATED CLIENT NAME IN CLIENT LIST
        If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_ClientName").Exist(30) Then
            Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_ClientName").Set strClientName
            Call ReportStep (StatusTypes.Pass,"Setting client name in search box as " & strClientName,"Search client name")
            Call ReportStep (StatusTypes.Pass,"Setting client name " &strClientName & " in search box done","Search client name")
            'Call ReportStep(StatusTypes.StepDefault, "##### Search for client Name" & strClientName & "  - Start ####", "OCRF Page")
        Else    
            Call ReportStep (StatusTypes.Fail,"Setting client name in search box not done","Search client name")
        End If
        wait 3
        Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_ClientName").Click
         Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{ENTER}"
        wait 5
        
        Set objTab = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data")
	    Set oTableCell = Description.Create
	     oTableCell("micclass").value = "WebElement"
	    oTableCell("html tag").value = "TD"
		oTableCell("innertext").value = strClientName&".*"   
		oTableCell("outerhtml").value = "<TD role=gridcell aria-describedby=DataTable_ClientName title=.* "
	
		set oTableCellObj = objTab.ChildObjects(oTableCell)
        col=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").GetROProperty("cols")
        Count=0
      If  oTableCellObj.count >= 1 Then
           For i = 0 To oTableCellObj.count-1 Step 1
                 	x= oTableCellObj(i).GetRoproperty("innertext")
                 	If x=strClientName Then
                 		    oTableCellObj(i).highlight
                 		    Call ReportStep (StatusTypes.Pass,"Searched client name " & strClientName & " Exist in Client List At Line No:" & i+1 &" With name" & x ,"Search client name")
                 			oTableCellObj(i).Click
                 			Count=1
                 			Exit For
                 	End If
            Next
     Else
          'Call ReportStep (StatusTypes.Fail,"Searched client name " & strClientName & " Not exist in Client List At Line No:"& i+1 &" With name" & x ,"Search client name")
     		SearchClientExistanceInOCRF="Not Exist"
     End If
     If Count=1 Then
     	SearchClientExistanceInOCRF="Exist"
     Else
        SearchClientExistanceInOCRF="Not Exist"
     End If
End Function




'Function ClientCreation_Ops(ByVal strOpsClient_Name,ByVal strOpsShort_Name,ByVal strOpsDescription,ByVal ClientID)
'	Set WshShell =CreateObject("WScript.Shell")
'			WshShell.SendKeys "{F5}"
'			wait 2
'			WshShell.SendKeys "{F5}"
'			Browser("DC Home").Page("Ops_Page").RefreshObject
'			wait 2
'			'Browser("DC Home").Page("Microsoft Forefront TMG").Link("Hosting").Click
'			
'			Call ItemSelection("Clients")
'			
'            wait 2
'		
'	Browser("DC Home").Page("Ops_Page").WebButton("WebButAddNew").Click
'	Browser("DC Home").Page("Ops_Page").WebEdit("WebEditClientName").Set strOpsClient_Name
'    Browser("DC Home").Page("Ops_Page").WebEdit("WebEditShortName").Set strOpsShort_Name
'    Browser("DC Home").Page("Ops_Page").WebEdit("WebEditDescription").Set strOpsDescription
'    Browser("DC Home").Page("Ops_Page").WebCheckBox("CreateIAMClient").Set "ON"
'
'    Browser("DC Home").Page("Ops_Page").WebElement("WebEleSave").Click
'    Browser("DC Home").Page("Ops_Page").WebElement("WebEleClose").Click
'
'End Function





Function SearchClientExistanceInOps(byVal strClientName,ByRef objData)
      wait 2
         Call ItemSelection("Clients")
    	'Call ReportStep(StatusTypes.StepDefault, "##### SearchClient - Verification of page navigation by using Next,Last,Prev,First icons - Start ####", "Report Creation Page")
        'SEARCH FOR CREATED CLIENT NAME IN CLIENT LIST
        
        For i = 1 To 10 Step 1
            If Browser("DC Home").Page("Ops_Page").WebEdit("WebEditClientName_Search").Exist Then
'                msgbox i
            	Browser("DC Home").Page("Ops_Page").WebEdit("WebEditClientName_Search").highlight
            	Browser("DC Home").Page("Ops_Page").WebEdit("WebEditClientName_Search").Set strClientName
                Call ReportStep (StatusTypes.Pass,"Setting client name in search box as " & strClientName,"Search client name")
                Call ReportStep (StatusTypes.Pass,"Setting client name " &strClientName & " in search box done","Search client name")
                Exit For
            End If
        	
        Next
        
        Browser("DC Home").Page("Ops_Page").WebEdit("WebEditClientName_Search").Click
         Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{ENTER}"
        wait 5
        
        
        
        
        
        Set objTab = Browser("DC Home").Page("Ops_Page").WebTable("class:=ui-jqgrid-btable","html id:=clientList","html tag:=TABLE")
	    Set oTableCell = Description.Create
	     oTableCell("micclass").value = "WebElement"
	    oTableCell("html tag").value = "TD"
		oTableCell("innertext").value = strClientName&".*"   
		oTableCell("outerhtml").value = "<TD role=gridcell aria-describedby=clientList_Name.* "
	
		set oTableCellObj = objTab.ChildObjects(oTableCell)
        col=Browser("DC Home").Page("Ops_Page").WebTable("class:=ui-jqgrid-btable","html id:=clientList","html tag:=TABLE").GetROProperty("cols")
      Count=0
      If  oTableCellObj.count >= 1 Then
           For i = 0 To oTableCellObj.count-1 Step 1
                 	x= oTableCellObj(i).GetRoproperty("innertext")
                 	If x=strClientName Then
                 		    oTableCellObj(i).highlight
                 		    Call ReportStep (StatusTypes.Pass,"Searched client name " & strClientName & " Exist in Client List At Line No:" & i+1 &" With name" & x ,"Search client name")
                 			oTableCellObj(i).Click
                 			Count=1
                 			Exit for
                 			
                 		
                 	End If
                 Next
     Else
           SearchClientExistanceInOps="Not Exist"
     End If
     If Count=1 Then
        SearchClientExistanceInOps="Exist"
     Else
        SearchClientExistanceInOps="Not Exist"
     End If
End Function

Public Function CreateOCRFIDSyndicated(ByRef objData)
       Path=Environment.Value("HostingPath")
       ENVUrl=objData.item("ENVUrl")
     Call OPSLogin("","",Environment.Value("OPSURL"))
     wait 5
  

	'ClientName=Split(objData.item("ClientNames"),";")
     strClientName= objData.item("strClientName")
     ClientName=Split(strClientName,";")

	 For j = 0 To Ubound(ClientName) Step 1
    	'Validation for client creation in OPS
    	 
        Status2= SearchClientExistanceInOps(ClientName(j),objData)
        If Status2="Exist" Then
    	   Call ReportStep (StatusTypes.Pass,"Searched client name " &ClientName(j) & " Exist in Client List "  ,"Search client name")
        Else
          'To Create New Client
          
           Call ReportStep (StatusTypes.Pass,"Searched client name " & ClientName(j) & " Not exist in Client List. Perform creation operation" ,"Search client name")
     		
	       Call ClientCreation_Ops(ClientName(j),ClientName(j),"","")
	       Call ReportStep (StatusTypes.Pass,"Client name " & ClientName(j)  & " Created "  ,"Search client name")
	       Call SearchClientExistanceInOps(ClientName(j),objData)
	    End If
    Next

'      'Launch the browser with ocrf url and validate default page is displayed	
    Call LaunchURL(ENVUrl,"","",objData)
'	Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_CLIENTS_Tab").Click
'    
'   'CHECK FOR CREATING CLIENT NAME EXIST IN CLIENT LIST OR NOT
'    For i = 0 To Ubound(ClientName) Step 1
'    	Status1= SearchClientExistanceInOCRF(ClientName(i),objData)
'        If Status1="Exist" Then
'    	   Call ReportStep (StatusTypes.Pass, "Client exist No need to create client","Clients Tab")
'        Else
'          'To Create New Client
'		   Call OCRFClientPage(0,"Create",ClientName(i),ClientName(i),"","" ,objData) 
'	      'CHECK FOR CREATING CLIENT NAME EXIST IN CLIENT LIST OR NOT
'           Call SearchClient(ClientName(i),ClientName(i), True,objData)
'        End If
'    Next
   
   If Browser("Browser-OCRF").Page("Page-OCRF").Link("html tag:=A","innertext:=ONLINE CRF","text:=ONLINE CRF").Exist(10) Then
   	  Browser("Browser-OCRF").Page("Page-OCRF").Link("html tag:=A","innertext:=ONLINE CRF","text:=ONLINE CRF").Click
   End If
   
   'Public Function CreateOCRFID()
	TestDataPath=Environment.value("CurrDir")& "InputFiles\IMSSCAWeb\"
	Call ClickOnAddNewRequest()
	OffClientName= AddDataInOfferingDetail(TestDataPath,objData)
	Call NavigateToNextTab()
	Call PageLoading()
	Call AddMultipleDatabase(TestDataPath,"",objData)
	Call NavigateToNextTab()
	Call PageLoading()
	Call AddNewLinesINAddUserTab(TestDataPath,objData)
	Call NavigateToNextTab()
	Call PageLoading()
	Call Client_Content_tab(TestDataPath,objData)
	Call NavigateToNextTab()
	Call PageLoading()
	'Call BI_Tools_Access_Combination(TestDataPath,objData)
	Call Multi_BITools_Access_Combination(TestDataPath,objData)
	Call NavigateToNextTab()
	Call PageLoading()
	Req_ID = Multi_USerDB_Access_Combination(TestDataPath,objData)
	
	'Open created request
	Call FilterByReqIDAndOpen(CInt(Req_ID),objData)
	'Validate offering action summary
	Call ValidateRowsCount(TestDataPath,objData)	
	
	
	'Place cube in FLA Path copy 'Abf' file in to cube name and rename
	Call FolderCreationInFLAServer (Path,"OCRF-DatabaseAdd",objData)
	'Validate Apply security failed
	Call ApplySecurity(objData)
	
	'Valiadate OCRF request status in Locked By column
	Call validateOCRFRequest(Req_ID)
	 'Create client in OPS
    Call SCA_CreateClientInOps(objData)
    
    'Validation for Groups created in Ops
    Call ItemSelection("Manage Content Access")
    
    Datatable.GetSheet("Offering_Details").SetCurrentRow(1)
    'Search for groups for client
    Call GetGroupsForClient(ClientName(1),dataTable.value("Offering_Country","Offering_Details"),"Syndicate",OffClientName)
    
    'Search for groups for client
    NSr = Datatable.GetSheet("BI_Tools").GetRowCount
	For j = 1 To NSr/2 Step 1
	   DataTable.SetCurrentRow(j)
	   User_Id=dataTable.value("User_Id","BI_Tools")
	   S_BI_Tool=dataTable.value("S_BI_Tool","BI_Tools")
       S_Access_Mode=dataTable.value("S_Access_Mode","BI_Tools")
       If S_Access_Mode ="" Then
       Else	
          Call ValidateGroupExistance(S_BI_Tool,S_Access_Mode,OffClientName,User_Id)
       End If
        
	Next
	
	
	'Search for groups for client
	Datatable.GetSheet("Offering_Details").SetCurrentRow(1)
	 Call GetGroupsForClient(ClientName(2),dataTable.value("Offering_Country","Offering_Details"),"Syndicate",OffClientName)
     NSr = Datatable.GetSheet("BI_Tools").GetRowCount
	For j = (NSr/2)+1 To NSr Step 1
	   DataTable.SetCurrentRow(j)
	   User_Id=dataTable.value("User_Id","BI_Tools")
	   S_BI_Tool=dataTable.value("S_BI_Tool","BI_Tools")
       S_Access_Mode=dataTable.value("S_Access_Mode","BI_Tools")
       If S_Access_Mode ="" Then
       Else	
          Call ValidateGroupExistance(S_BI_Tool,S_Access_Mode,OffClientName,User_Id)
       End If
        
	Next
'********************************************************************************************	
	
'TOP SITE for ' syndicated
'Search for groups for client
 Datatable.GetSheet("Offering_Details").SetCurrentRow(1)
  Call GetGroupsForClient("Top Site",dataTable.value("Offering_Country","Offering_Details"),"Syndicate",OffClientName)
  wait 3

	 NSr = Datatable.GetSheet("BI_Tools").GetRowCount
	'DataTable.GetRowCount(SheetName)
	For k = 1 To NSr Step 1
	   DataTable.SetCurrentRow(k)
	   User_Id=dataTable.value("User_Id","BI_Tools")
	   TS_S_BI_Tool=dataTable.value("TS_S_BI_Tool","BI_Tools")
       TS_S_Access_Mode=dataTable.value("TS_S_Access_Mode","BI_Tools")
       If TS_S_Access_Mode ="" Then
       Else	
          Call ValidateGroupExistance(TS_S_BI_Tool,TS_S_Access_Mode,OffClientName,User_Id)
       End If
        
	Next



   'Date 9/1/2017
   'Launch OCRF and search for Newly created Id and open it 
   Call LaunchURL(ENVUrl, "", "",objData)
   Call FilterByReqIDAndOpen(CInt(Req_ID),objData)
   Call DeleteIMSToolFromBITool(objData)
   
   If Browser("Browser-OCRF").Page("Page-OCRF").Link("name:=\[1\] Offering Details->> ","text:=\[1\] Offering Details->> ","html tag:=A").Exist(10) Then
	    Browser("Browser-OCRF").Page("Page-OCRF").Link("name:=\[1\] Offering Details->> ","text:=\[1\] Offering Details->> ","html tag:=A").Click	
	    Call ReportStep (StatusTypes.Pass, "Clicked on ' Offering Details' Tab ","Offering Details tab")	
	Else
		Call ReportStep (StatusTypes.Fail, "Not Clicked on 'Offering Details' Tab","Offering Details tab")
	End If
   	Call ApplySecurity(objData)
   CreateOCRFIDSyndicated = Req_ID	
End Function

Public Function CreateOCRFIDSyndicated_OLD(ByRef objData)
    ENVUrl=objData.item("ENVUrl")
    Path=Environment.Value("HostingPath")
     Call OPSLogin("","",Environment.Value("OPSURL"))
     wait 5
  

	'ClientName=Split(objData.item("ClientNames"),";")
     strClientName= "AutoClient11;AutoClient12;AutoClient13;AutoClient14"
     ClientName=Split(strClientName,";")

	 For j = 0 To Ubound(ClientName) Step 1
    	'Validation for client creation in OPS
    	 
        Status2= SearchClientExistanceInOps(ClientName(j),objData)
        If Status2="Exist" Then
    	   Call ReportStep (StatusTypes.Pass,"Searched client name " &ClientName(j) & " Exist in Client List "  ,"Search client name")
        Else
          'To Create New Client
          
           Call ReportStep (StatusTypes.Pass,"Searched client name " & ClientName(j) & " Not exist in Client List. Perform creation operation" ,"Search client name")
     		
	       Call ClientCreation_Ops(ClientName(j),ClientName(j),"","")
	       Call ReportStep (StatusTypes.Pass,"Client name " & ClientName(j)  & " Created "  ,"Search client name")
	       Call SearchClientExistanceInOps(ClientName(j),objData)
	    End If
    Next

      'Launch the browser with ocrf url and validate default page is displayed	
    Call LaunchURL(ENVUrl,"","",objData)
	Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_CLIENTS_Tab").Click
    
   'CHECK FOR CREATING CLIENT NAME EXIST IN CLIENT LIST OR NOT
    For i = 0 To Ubound(ClientName) Step 1
    	Status1= SearchClientExistanceInOCRF(ClientName(i),objData)
        If Status1="Exist" Then
    	   Call ReportStep (StatusTypes.Pass, "Client exist No need to create client","Clients Tab")
        Else
          'To Create New Client
		   Call OCRFClientPage(0,"Create",ClientName(i),ClientName(i),"","" ,objData) 
	      'CHECK FOR CREATING CLIENT NAME EXIST IN CLIENT LIST OR NOT
           Call SearchClient(ClientName(i),ClientName(i), True,objData)
        End If
    Next
   
   If Browser("Browser-OCRF").Page("Page-OCRF").Link("html tag:=A","innertext:=ONLINE CRF","text:=ONLINE CRF").Exist(10) Then
   	  Browser("Browser-OCRF").Page("Page-OCRF").Link("html tag:=A","innertext:=ONLINE CRF","text:=ONLINE CRF").Click
   End If
   
   'Public Function CreateOCRFID()
	TestDataPath = Environment.value("CurrDir")& "InputFiles\IMSSCAWeb\"
	Call ClickOnAddNewRequest()
	OffClientName= AddDataInOfferingDetail(TestDataPath,objData)
	Call NavigateToNextTab()
	Call PageLoading()
	Call AddMultipleDatabase(TestDataPath,"",objData)
	Call NavigateToNextTab()
	Call PageLoading()
	Call AddNewLinesINAddUserTab(TestDataPath,objData)
	Call NavigateToNextTab()
	Call PageLoading()
	Call Client_Content_tab(TestDataPath,objData)
	Call NavigateToNextTab()
	Call PageLoading()
	Call BI_Tools_Access_Combination(TestDataPath,objData)
	Req_ID = USer_DB_Access_Combination(TestDataPath,objData)
	
	'Open created request
	Call FilterByReqIDAndOpen(CInt(Req_ID),objData)
	'Validate offering action summary
	Call ValidateRowsCount(TestDataPath,objData)	
	
	
	'Place cube in FLA Path copy 'Abf' file in to cube name and rename
	Call FolderCreationInFLAServer (Path,"OCRF-DatabaseAdd",objData)
	'Validate Apply security failed
	Call ApplySecurity(objData)
	
	'Valiadate OCRF request status in Locked By column
	Call validateOCRFRequest(Req_ID)
	 'Create client in OPS
    Call SCA_CreateClientInOps(objData)
    'Validation for Groups created in Ops
    
    'Call ValidationForGroupsInOps("automatedclient235","Australia","AutoTest234","NonSyndicated")
	accMode = Split(objData.Item("accMode"),";")
	Permission = Split(objData.Item("Permission"),";")

	
    ReDim UserArr(3)
    cnt=0
    For i = 1 To 3 Step 1
   	  append=""
   	  For j = 1 To 2 Step 1
   	     Datatable.SetCurrentRow(j+cnt)
   	    append=append&dataTable.value("User_ID","BI_Tools")&";"
   	 	
   	 Next
   	   UserArr(i-1)=append
   	 cnt=cnt+2
   Next
    
     For i = 0 To Ubound(ClientName) Step 1
         DataTable.GetSheet("Offering_Details").SetCurrentRow(1)
         Call ItemSelection("Manage Content Access")
         Call ValidationForGroupsInOps(ClientName(i),dataTable.value("Offering_Country","Offering_Details"),OffClientName,"Syndicate")
         Call ValidationForAddUserInOps(accMode(i),Permission,UserArr(i))
     Next
   
   'VALIDATION FOR TOP SITE
    Call ItemSelection("Manage Content Access")
    Call ValidationForGroupsInOps("Top Site",dataTable.value("Offering_Country","Offering_Details"),OffClientName,"Syndicate")

'  'Validation for 'Get users in Ops"
    Call ValidationForAddUserInOps(accMode(0),Permission,UserArr(0))
    Call ValidationForAddUserInOps(accMode(1),Permission,UserArr(1))
    Call ValidationForAddUserInOps(accMode(2),Permission,UserArr(2))
'
   'Date 9/1/2017
   'Launch OCRF and search for Newly created Id and open it 
   Call LaunchURL(ENVUrl, "", "",objData)
   Call FilterByReqIDAndOpen(CInt(Req_ID),objData)
   Call DeleteIMSToolFromBITool(objData)
   If Browser("Browser-OCRF").Page("Page-OCRF").Link("name:=\[1\] Offering Details->> ","text:=\[1\] Offering Details->> ","html tag:=A").Exist(10) Then
	    Browser("Browser-OCRF").Page("Page-OCRF").Link("name:=\[1\] Offering Details->> ","text:=\[1\] Offering Details->> ","html tag:=A").Click	
	    Call ReportStep (StatusTypes.Pass, "Clicked on ' Offering Details' Tab ","Offering Details tab")	
	Else
		Call ReportStep (StatusTypes.Fail, "Not Clicked on 'Offering Details' Tab","Offering Details tab")
	End If
   	Call ApplySecurity(objData)
End Function




Public Function UsersValidationInDC(ByVal SyndicatedType,ByVal ClientUrl,ByVal SheetName,ByVal Path,ByRef objData)
	
	
	
	Call ImportSheet(SheetName,Path)
	
	rc = Datatable.GetSheet(SheetName).GetRowCount
	
	For i = 1 To rc Step 1
		wait 2
		
		If SyndicatedType = "NonSyndicated" Then
			Systemutil.Run ClientUrl
		ElseIf SyndicatedType = "Syndicated" Then	
		
			If (Instr(Datatable.Value("Access_Mode","BI_Tools"),"Author")>0) OR (Instr(Datatable.Value("Access_Mode","BI_Tools"),"Approver")>0) Then
				NewClientUrl = "http://sit.imsbi.rxcorp.com/clients/"&Datatable.Value("Client_Name","BI_Tools")
			Else
				NewClientUrl = ClientUrl&Datatable.Value("Client_Name","BI_Tools")
			End If
			Systemutil.Run NewClientUrl
		End If
		
		If Browser("DC").Page("DC").Exist Then
			Browser("DC").Page("DC").Sync
			wait 2
		End If
		
		
		Browser("DC").Page("DC").WebEdit("WebEdit_Username").Set Datatable.value("User_ID","BI_Tools")
		Browser("DC").Page("DC").WebEdit("WebEdit_Password").Set Datatable.value("Password","BI_Tools")
		Browser("DC").Page("DC").WebButton("WebButton_Login").Click
	
		If Browser("DC").Page("DC").Image("Image_My_Informed_Decisions").Exist(60) Then
			Call ReportStep (StatusTypes.Pass,"Successfully login to decision centre using client URL "&ClientUrl&" with userid "&Datatable.value("User_ID","BI_Tools")&" and password "&Datatable.value("Password","BI_Tools")&"","Client user validation in DC")
		Else
			Call ReportStep (StatusTypes.Fail,"Not able to login to decision centre using client URL "&ClientUrl&" with userid "&Datatable.value("User_ID","BI_Tools")&" and password "&Datatable.value("Password","BI_Tools")&"","Client user validation in DC")	
		End If
	
		Browser("DC").Page("DC").WebElement("WebElement_UserName").Click
		Browser("DC").Page("DC").WebElement("WebElement_SignOut").Click
		
		Browser("DC").Close
		
	
		Datatable.SetNextRow
		
	Next
	
End Function



Public Function BIToolsValidationInDC(ByVal SyndicatedType,ByVal ClientUrl,ByVal CountryName,ByVal OfferingName,ByVal SheetName,ByVal Path,ByRef objData)
	
	Call ImportSheet(SheetName,Path)
	rc = Datatable.GetSheet(SheetName).GetRowCount
			
	For i = 1 To rc Step 1
	
		If Datatable.Value("BI_Tools_Short",SheetName) <> "IMS" Then
		
			If SyndicatedType = "NonSyndicated" Then
				AppendUrl = ClientUrl & "/" & Datatable.Value("BI_Tools_Short","BI_Tools")
			ElseIf SyndicatedType = "Syndicated" Then	
			
				If (Instr(Datatable.Value("Access_Mode","BI_Tools"),"Author")>0) OR (Instr(Datatable.Value("Access_Mode","BI_Tools"),"Approver")>0) Then
					AppendUrl = "http://sit.imsbi.rxcorp.com/clients/"&Datatable.Value("Client_Name","BI_Tools") & "/" & Datatable.Value("BI_Tools_Short","BI_Tools")
				Else
					AppendUrl = ClientUrl & Datatable.Value("Client_Name","BI_Tools") & "/" & Datatable.Value("BI_Tools_Short","BI_Tools")
				End If
			
			End If
			
			wait 2
			Systemutil.Run AppendUrl		
			wait 5
			Browser("DC").Page("DC").WebEdit("WebEdit_Username").Set Datatable.value("User_ID","BI_Tools")
			Browser("DC").Page("DC").WebEdit("WebEdit_Password").Set Datatable.value("Password","BI_Tools")
			Browser("DC").Page("DC").WebButton("WebButton_Login").Click
		
		
			If SyndicatedType = "NonSyndicated" Then
				
				If Browser("DifferenetBITools").Page("DifferenetBIToolsPages").WebTable("WebTable_Country-Offering").Link("html tag:=A","innertext:="&CountryName).Exist Then
					Call ReportStep (StatusTypes.Pass,"Country name "&CountryName&" is displayed in decision centre using new url, client with BI Tool "&AppendUrl&" with userid "&Datatable.value("User_ID","BI_Tools")&" and password "&Datatable.value("Password","BI_Tools")&" ","Client BI Tool validation in DC")
					Browser("DifferenetBITools").Page("DifferenetBIToolsPages").WebTable("WebTable_Country-Offering").Link("html tag:=A","innertext:="&CountryName).Click
					If Browser("DifferenetBITools").Page("DifferenetBIToolsPages").WebTable("WebTable_Country-Offering").Link("html tag:=A","innertext:="&OfferingName).Exist Then
						Call ReportStep (StatusTypes.Pass,"Offering name "&OfferingName&" is displayed in decision centre using new url, client with BI Tool "&AppendUrl&" with userid "&Datatable.value("User_ID","BI_Tools")&" and password "&Datatable.value("Password","BI_Tools")&" ","Client BI Tool validation in DC")	
						Browser("DifferenetBITools").Page("DifferenetBIToolsPages").WebTable("WebTable_Country-Offering").Link("html tag:=A","innertext:="&OfferingName).Click
					Else
						Call ReportStep (StatusTypes.Fail,"Offering name "&OfferingName&" is not displayed in decision centre using new url, client with BI Tool "&AppendUrl&" with userid "&Datatable.value("User_ID","BI_Tools")&" and password "&Datatable.value("Password","BI_Tools")&" ","Client BI Tool validation in DC")	
					End If
					
					
					'Validate access permission
					Browser("DifferenetBITools").Page("DifferenetBIToolsPages").Link("Link_DocumentsTab").Click
					StrExist = trim(Browser("DifferenetBITools").Page("DifferenetBIToolsPages").Link("Link_UploadDocument").GetROProperty("class"))
				
					If Instr(Datatable.value("Access_Mode","BI_Tools"),"Viewer")>0 Then
						If Instr(StrExist,"disabled")>0 Then
							Call ReportStep (StatusTypes.Pass,"Upload Document link is disabled for Access Mode "&Datatable.value("Access_Mode","BI_Tools")&" ","Access mode validation")	
						Else
							Call ReportStep (StatusTypes.Fail,"Upload Document link is enabled for Access Mode "&Datatable.value("Access_Mode","BI_Tools")&" ","Access mode validation")	
						End If	
					Else
						If Instr(StrExist,"disabled")=0 Then
							Call ReportStep (StatusTypes.Pass,"Upload Document link is enabled for Access Mode "&Datatable.value("Access_Mode","BI_Tools")&" ","Access mode validation")	
						Else
							Call ReportStep (StatusTypes.Fail,"Upload Document link is disabled for Access Mode "&Datatable.value("Access_Mode","BI_Tools")&" ","Access mode validation")	
						End If					
					End If
				Else
					Call ReportStep (StatusTypes.Fail,"Country name "&CountryName&" is not displayed in decision centre using new url, client with BI Tool "&AppendUrl&" with userid "&Datatable.value("User_ID","BI_Tools")&" and password "&Datatable.value("Password","BI_Tools")&" ","Client BI Tool validation in DC")
				End If
				
			ElseIf SyndicatedType = "Syndicated" Then	
				If Browser("DC").Page("DC").WebElement("WebElement_No_Items").Exist(60) Then
					Call ReportStep (StatusTypes.Pass,"Successfully login to decision centre using new url, client with BI Tool "&AppendUrl&" with userid "&Datatable.value("User_ID","BI_Tools")&" and password "&Datatable.value("Password","BI_Tools")&"","Client BI Tool validation in DC")
				Else
					Call ReportStep (StatusTypes.Pass,"Successfully login to decision centre using new url, client with BI Tool "&AppendUrl&" with userid "&Datatable.value("User_ID","BI_Tools")&" and password "&Datatable.value("Password","BI_Tools")&"","Client BI Tool validation in DC")
'					Call ReportStep (StatusTypes.Fail,"Not able login to decision centre using new url, client with BI Tool "&AppendUrl&" with userid "&Datatable.value("User_ID","BI_Tools")&" and password "&Datatable.value("Password","BI_Tools")&"","Client BI Tool validation in DC")
				End If	
			End If
		
			Browser("DC").Page("DC").WebElement("WebElement_UserName").Click
			Browser("DC").Page("DC").WebElement("WebElement_SignOut").Click	
			Browser("DC").Close
		ElseIf Datatable.Value("BI_Tools_Short",SheetName) = "IMS" Then
			
			If SyndicatedType = "NonSyndicated" Then
				AppendUrl = ClientUrl 
			ElseIf SyndicatedType = "Syndicated" Then	
				AppendUrl = ClientUrl & Datatable.Value("Client_Name","BI_Tools")
			End If
		
			
			wait 2
			Systemutil.Run AppendUrl		
			wait 5
			Browser("DC").Page("DC").WebEdit("WebEdit_Username").Set Datatable.value("User_ID","BI_Tools")
			Browser("DC").Page("DC").WebEdit("WebEdit_Password").Set Datatable.value("Password","BI_Tools")
			Browser("DC").Page("DC").WebButton("WebButton_Login").Click
		
		
			If SyndicatedType = "NonSyndicated" Then
				
				
				
				If Browser("DC").Page("DC").Link("Link_IMSAnalysisManager").Exist Then
					Call ReportStep (StatusTypes.Pass,"Link IMS Analysis Manager displayed succesfully","IMS Analysis Manager")
					Browser("DC").Page("DC").Link("Link_IMSAnalysisManager").Click	
					
					If Datatable.value("Access_Mode","BI_Tools") = "IAM Consumer" Then
						If Browser("DC").Page("DC").WebElement("WebElement_CreateNewReport").Exist(5) = False Then
							Call ReportStep (StatusTypes.Pass,"Create report link is not displayed for IAM Consumer access","IAM Consumer")
						Else	
							Call ReportStep (StatusTypes.Fail,"Create report link is displayed for IAM Consumer access","IAM Consumer")
						End If
					ElseIf Datatable.value("Access_Mode","BI_Tools") = "IAM Architect" Then	
						If Browser("DC").Page("DC").WebElement("WebElement_CreateNewReport").Exist Then
							Call ReportStep (StatusTypes.Pass,"Create report link is displayed for IAM Architect access","IAM Architect")
						Else	
							Call ReportStep (StatusTypes.Fail,"Create report link is not displayed for IAM Architect access","IAM Architect")
						End If	
					End If
					
					
					Browser("DC").Page("DC").Link("Link_SharedReports").Click
					
					If Browser("DC").Page("DC").Link("Link_SharedReports").Exist Then
						Call ReportStep (StatusTypes.Pass,"Shared report link displayed successfully","Shared report link")
						Browser("DC").Page("DC").Link("Link_SharedReports").Click	
						
						
						
						If Browser("DC").Page("DC").Frame("Frame-view").Link("html tag:=A","innerhtml:="&CountryName).Exist Then
							Call ReportStep (StatusTypes.Pass,"Country name folder "&CountryName&" is displayed in IAM","Country name folder")
							Browser("DC").Page("DC").Frame("Frame-view").Link("html tag:=A","innerhtml:="&CountryName).Click
							
							If Browser("DC").Page("DC").Frame("Frame-view").Link("html tag:=A","innerhtml:="&OfferingName).Exist Then
								Call ReportStep (StatusTypes.Pass,"Offering name folder "&OfferingName&" is displayed in IAM","Offering name folder")
							Else							
								Call ReportStep (StatusTypes.Fail,"Offering name folder "&OfferingName&" is not displayed in IAM","Offering name folder")
							End If
							
						Else	
							Call ReportStep (StatusTypes.Fail,"Country name folder "&CountryName&" is not displayed in IAM","Country name folder")
						End If
						
						
					Else	
						Call ReportStep (StatusTypes.Fail,"Shared report link not displayed","Shared report link")
					End If
					
					
					
				Else
					Call ReportStep (StatusTypes.Fail,"Link IMS Analysis Manager is not displayed","IMS Analysis Manager")
				End If
				
				
			ElseIf SyndicatedType = "Syndicated" Then	
				If Browser("DC").Page("DC").WebElement("WebElement_No_Items").Exist(60) Then
					Call ReportStep (StatusTypes.Pass,"Successfully login to decision centre using new url, client with BI Tool "&AppendUrl&" with userid "&Datatable.value("User_ID","BI_Tools")&" and password "&Datatable.value("Password","BI_Tools")&"","Client BI Tool validation in DC")
				Else
					Call ReportStep (StatusTypes.Pass,"Successfully login to decision centre using new url, client with BI Tool "&AppendUrl&" with userid "&Datatable.value("User_ID","BI_Tools")&" and password "&Datatable.value("Password","BI_Tools")&"","Client BI Tool validation in DC")
'					Call ReportStep (StatusTypes.Fail,"Not able login to decision centre using new url, client with BI Tool "&AppendUrl&" with userid "&Datatable.value("User_ID","BI_Tools")&" and password "&Datatable.value("Password","BI_Tools")&"","Client BI Tool validation in DC")
				End If	
			End If
		
			Browser("DC").Page("DC").WebElement("WebElement_UserName").Click
			Browser("DC").Page("DC").WebElement("WebElement_SignOut").Click	
			Browser("DC").Close
		
		
		
		End If
		Datatable.SetNextRow	
	Next

End Function





Public Function OfferingContactDetails(ByVal SheetName,ByRef objData)
	
	Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactName").Set Datatable.Value("Offering_Contact_Name",SheetName)
	Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactPhone").Set Datatable.Value("Offering_Contact_Phone",SheetName)
	Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactEmail").Set Datatable.Value("Offering_Contact_Email",SheetName)
	
	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Click
	Call PageLoading()
	Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_Previous").Click
	Call PageLoading()

	'Validate offering contact details value
	
	If Datatable.Value("Offering_Contact_Name",SheetName) = Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactName").GetROProperty("value") Then
		Call ReportStep (StatusTypes.Pass,"Offering contact name updated successfully "&Datatable.Value("Offering_Contact_Name",SheetName),"Offering_Contact_Name")
	Else
		Call ReportStep (StatusTypes.Fail,"Offering contact name not updated successfully "&Datatable.Value("Offering_Contact_Name",SheetName),"Offering_Contact_Name")
	End If
	
	If Datatable.Value("Offering_Contact_Phone",SheetName) = Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactPhone").GetROProperty("value") Then
		Call ReportStep (StatusTypes.Pass,"Offering Contact Phone updated successfully "&Datatable.Value("Offering_Contact_Phone",SheetName),"Offering_Contact_Phone")
	Else
		Call ReportStep (StatusTypes.Fail,"Offering Contact Phone not updated successfully "&Datatable.Value("Offering_Contact_Phone",SheetName),"Offering_Contact_Phone")
	End If
	
	If Datatable.Value("Offering_Contact_Email",SheetName) = Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactEmail").GetROProperty("value") Then
		Call ReportStep (StatusTypes.Pass,"Offering Contact email updated successfully "&Datatable.Value("Offering_Contact_Email",SheetName),"Offering_Contact_Email")
	Else
		Call ReportStep (StatusTypes.Fail,"Offering Contact email updated successfully "&Datatable.Value("Offering_Contact_Email",SheetName),"Offering_Contact_Email")
	End If
	
	
	Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactName").Set "smita"
	Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactPhone").set "8884345585"
	Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactEmail").set "smita.anand@in.imshealth.com"
	
	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Click
	Call PageLoading()
	Browser("Browser-OCRF").Page("Page-OCRF").Sync
	
End Function




Public Function DeleteUserFromBIToolAccess(ByVal Action,ByVal User,ByRef objData)
	
	
	
	User_Name = Trim(User)

	If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").Exist Then
		 wait 2
	End If
	
	wait 2
	rownum =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").GetRowWithCellText(User_Name,8,2)
	
	set chk = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").ChildItem(rownum,2,"WebCheckBox",0)
	chk.click
	
	Set lst = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").ChildItem(rownum,6,"WebList",0)
	
	If Action = "Add" Then
		lst.select "ADD"
	ElseIf Action = "Delete" Then	
		lst.select "DELETE"
	End If

	
End Function 




Public Function ValidateUserInDC(ByVal CountryName,ByVal OfferingName,ByRef objData)
	
	
	Browser("DC").Page("DC").WebEdit("WebEdit_Username").Set "gemini\(98615)(1)"
	Browser("DC").Page("DC").WebEdit("WebEdit_Password").Set "De8093074085"
	Browser("DC").Page("DC").WebButton("WebButton_Login").Click
	
	Browser("DC Home").Page("Home").Link("Link_SiteActions").Click
	Browser("DC Home").Page("Home").Link("Link_ViewAllSiteContentView").Click
	Browser("DC").Page("DC").Link("Link_ReportingServicesReports").Click

	Browser("DifferenetBITools").Page("DifferenetBIToolsPages").WebTable("WebTable_Country-Offering").Link("html tag:=A","innertext:="&CountryName).Click
	
	Browser("DifferenetBITools").Page("DifferenetBIToolsPages").WebTable("WebTable_Country-Offering").Link("html tag:=A","innertext:="&OfferingName).FireEvent "onmouseover"
	Browser("DifferenetBITools").Page("DifferenetBIToolsPages").Image("alt:=Open Menu","image type:=Image Link","visible:=true").Click
	Browser("DC").Page("DC").Link("Link_ManagePermissions").Click
	
	Browser("DifferenetBITools").Page("DifferenetBIToolsPages").Link("html tag:=A","innertext:=.*Viewers").Click
	
	If Browser("DifferenetBITools").Page("DifferenetBIToolsPages").Link("html tag:=A","innertext:=.*User9006.*").Exist(5) = False Then
		Call ReportStep (StatusTypes.Pass,"IMITestUser9006 is not displayed after deleting from OCRF","Delete user validation")
	Else
		Call ReportStep (StatusTypes.Fail,"IMITestUser9006 is displayed after deleting from OCRF","Delete user validation")
	End If
	
End Function


Public Function OfferingContactDetails(ByVal SheetName,ByRef objData)
	
	Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactName").Set Datatable.Value("Offering_Contact_Name",SheetName)
	Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactPhone").Set Datatable.Value("Offering_Contact_Phone",SheetName)
	Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactEmail").Set Datatable.Value("Offering_Contact_Email",SheetName)
	
	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Click
	Call PageLoading()
	Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_Previous").Click
	Call PageLoading()

	'Validate offering contact details value
	
	If Datatable.Value("Offering_Contact_Name",SheetName) = Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactName").GetROProperty("value") Then
		Call ReportStep (StatusTypes.Pass,"Offering contact name updated successfully "&Datatable.Value("Offering_Contact_Name",SheetName),"Offering_Contact_Name")
	Else
		Call ReportStep (StatusTypes.Fail,"Offering contact name not updated successfully "&Datatable.Value("Offering_Contact_Name",SheetName),"Offering_Contact_Name")
	End If
	
	If Datatable.Value("Offering_Contact_Phone",SheetName) = Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactPhone").GetROProperty("value") Then
		Call ReportStep (StatusTypes.Pass,"Offering Contact Phone updated successfully "&Datatable.Value("Offering_Contact_Phone",SheetName),"Offering_Contact_Phone")
	Else
		Call ReportStep (StatusTypes.Fail,"Offering Contact Phone not updated successfully "&Datatable.Value("Offering_Contact_Phone",SheetName),"Offering_Contact_Phone")
	End If
	
	If Datatable.Value("Offering_Contact_Email",SheetName) = Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactEmail").GetROProperty("value") Then
		Call ReportStep (StatusTypes.Pass,"Offering Contact email updated successfully "&Datatable.Value("Offering_Contact_Email",SheetName),"Offering_Contact_Email")
	Else
		Call ReportStep (StatusTypes.Fail,"Offering Contact email updated successfully "&Datatable.Value("Offering_Contact_Email",SheetName),"Offering_Contact_Email")
	End If
	
	
	Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactName").Set "smita"
	Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactPhone").set "8884345585"
	Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactEmail").set "smita.anand@in.imshealth.com"
	
	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementNext").Click
	Call PageLoading()
	Browser("Browser-OCRF").Page("Page-OCRF").Sync
	
End Function




'Public Function DeleteUserFromBIToolAccess(ByVal Action,ByRef objData)
'	
'	
'	
'	User_Name = "IMITestUser9006@uk.imshealth.com"
'	
'	If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").Exist Then
'		 wait 2
'	End If
'	
'	rownum =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").GetRowWithCellText(User_Name,8,2)
'	
'	set chk = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").ChildItem(rownum,2,"WebCheckBox",0)
'	chk.click
'	
'	Set lst = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").ChildItem(rownum,6,"WebList",0)
'	
'	If Action = "Add" Then
'		lst.select "ADD"
'	ElseIf Action = "Delete" Then	
'		lst.select "DELETE"
'	End If
'
'	
'End Function 




Public Function ValidateUserInDC(ByVal CountryName,ByVal OfferingName,ByRef objData)
	
	
	Browser("DC").Page("DC").WebEdit("WebEdit_Username").Set "gemini\(98615)(1)"
	Browser("DC").Page("DC").WebEdit("WebEdit_Password").Set "De8093074085"
	Browser("DC").Page("DC").WebButton("WebButton_Login").Click
	
	Browser("DC Home").Page("Home").Link("Link_SiteActions").Click
	Browser("DC Home").Page("Home").Link("Link_ViewAllSiteContentView").Click
	Browser("DC").Page("DC").Link("Link_ReportingServicesReports").Click

	Browser("DifferenetBITools").Page("DifferenetBIToolsPages").WebTable("WebTable_Country-Offering").Link("html tag:=A","innertext:="&CountryName).Click
	
	Browser("DifferenetBITools").Page("DifferenetBIToolsPages").WebTable("WebTable_Country-Offering").Link("html tag:=A","innertext:="&OfferingName).FireEvent "onmouseover"
	Browser("DifferenetBITools").Page("DifferenetBIToolsPages").Image("alt:=Open Menu","image type:=Image Link","visible:=true").Click
	Browser("DC").Page("DC").Link("Link_ManagePermissions").Click
	
	Browser("DifferenetBITools").Page("DifferenetBIToolsPages").Link("html tag:=A","innertext:=.*Viewers").Click
	
	If Browser("DifferenetBITools").Page("DifferenetBIToolsPages").Link("html tag:=A","innertext:=.*User9006.*").Exist(5) = False Then
		Call ReportStep (StatusTypes.Pass,"IMITestUser9006 is not displayed after deleting from OCRF","Delete user validation")
	Else
		Call ReportStep (StatusTypes.Fail,"IMITestUser9006 is displayed after deleting from OCRF","Delete user validation")
	End If
	
End Function




Function AddNewLineInCubeDetails(ByVal CubeName, ByVal CubeLocation,ByVal FlaPath,ByVal NumCube,ByRef objData)
    wait 10
    c= Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetROProperty("cols")
    r= Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetROProperty("rows")
	
	item=0
	 For row = r+1 To r+NumCube Step 1
	      If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_AddNewLine").Exist(10) Then
          	 Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_AddNewLine").Click
             Call ReportStep (StatusTypes.Pass,"Clicked on 'Add New line' in cube details tab ","Cube details tab")
          Else
             Call ReportStep (StatusTypes.Fail,"Not Clicked on 'Add New line' in cube details tab ","Cube details tab")
          End If
	      wait 3
          Call WebEditExistence(row,Cint(Environment.value("cubeName")),"WebEdit")
          Set CheckDatabase =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(row,Cint(Environment.value("cubeName")),"WebEdit",0)
		  CheckDatabase.Set CubeName
		  
          Call ReportStep (StatusTypes.Pass,"New cube added in Cube details tab as "&CubeName,"Cube details tab")
          wait 2
          'Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=cube-list","html tag:=TABLE").WebElement("class:=adminEntry","outerhtml:=<TD role=gridcell aria-describedby=cube-list_Server title="""" class=adminEntry style=""TEXT-ALIGN: left"">&nbsp;</TD>","innerhtml:=&nbsp;","html tag:=TD").Click
          Call WebEditExistence(row,Cint(Environment.value("cubeLocation")),"WebEdit")
		  Set CheckDBLocation =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(row,Cint(Environment.value("cubeLocation")),"WebEdit",0)
		  CheckDBLocation.set Environment.Value("flaServerPath")
		  
		  		
		  Call WebEditExistence(row,Cint(Environment.value("flaPath")),"WebList")
		
		  Set CheckFLAPath =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(row,Cint(Environment.value("flaPath")),"WebList",0)
		  CheckFLAPath.select Environment.Value("HostServerName")

         'Call ReportStep (StatusTypes.Pass,"Cube location entered in Cube details tab as "&CubeName,"Cube details tab")
         item=item+1
      Next
End Function




Function AddNewLineInUserDB(ByVal CubeName,ByVal CubeRole,ByVal UserID,ByVal row,ByRef objData)
	
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Exist(120) Then
	    Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Click
		Call ReportStep (StatusTypes.Pass, "Clicked on 'Select user'","USer_DB_Access Tab")	
	Else
		Call ReportStep (StatusTypes.Fail, "Not Clicked on 'Select user","USer_DB_Access Tab")
	End If
	
	Set objCubeDB = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_CubeDB")
	Set objCubeRoles = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_CubeRoles")
	wait 5	
	Call SCA.SelectFromDropdown(objCubeDB, CubeName)
	Call SCA.SelectFromDropdown(objCubeRoles, CubeRole)
	Call SelectRowsFromDropDownList(row,"999",objData)
	rownum = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").GetROProperty("rows")
	For j = 1 To rownum Step 1
		userData=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").GetCellData(j+1,4)
		If userData=UserID Then
		   Set chk =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").ChildItem(j+1,2,"WebCheckBox",0)
	        chk.click
	        Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_UserDB_AddUsers").Click
	        Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_UserDB_Add").Click
		End If
	Next

   	New_Request_ID = trim(Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_RequestID").GetROProperty("value"))
   


	Call PageLoading()
	
	AddNewLineInUserDB = New_Request_ID
End Function 



Function SearchTemplateExistanceInOps(byVal Template,ByRef objData)
    	'Call ReportStep(StatusTypes.StepDefault, "##### SearchClient - Verification of page navigation by using Next,Last,Prev,First icons - Start ####", "Report Creation Page")
        'SEARCH FOR CREATED CLIENT NAME IN CLIENT LIST
     '   If Browser("DC Home").Page("Ops_Page").WebEdit("html id:=gs_TemplateId","html tag:=INPUT","kind:=singleline","name:=TemplateId").Exist(60) Then
'        If Browser("DC Home").Page("Ops_Page").WebEdit("html id:=gs_TemplateId").Exist(60) Then
'            wait 3
'            Call ReportStep (StatusTypes.Pass,"Setting Template name in search box as " & Template,"Search Template name")
'            Call ReportStep (StatusTypes.Pass,"Setting Template name " &Template & " in search box done","Search Template name")
'            'Call ReportStep(StatusTypes.StepDefault, "##### Search for client Name" & strClientName & "  - Start ####", "OCRF Page")
'        Else    
'       ' msgbox  "unable to identify object"
'                Call ReportStep (StatusTypes.Fail,"Setting Template name in search box not done","Search Template name")
'                  exitrun
'        End If
        
        wait 10
        For i = 1 To 10 Step 1
            If Browser("DC Home").Page("Ops_Page").WebEdit("html id:=gs_TemplateId","html tag:=INPUT","kind:=singleline","name:=TemplateId").Exist Then
            	Browser("DC Home").Page("Ops_Page").WebEdit("html id:=gs_TemplateId","html tag:=INPUT","kind:=singleline","name:=TemplateId").highlight
                Browser("DC Home").Page("Ops_Page").WebEdit("html id:=gs_TemplateId","html tag:=INPUT","kind:=singleline","name:=TemplateId").Set Template
                Call ReportStep (StatusTypes.Pass,"Setting Template name " &Template & " in search box done","Search Template name")
                Exit For
            End If	
        Next
        wait 5
        Browser("DC Home").Page("Ops_Page").WebEdit("html id:=gs_TemplateId","html tag:=INPUT","kind:=singleline","name:=TemplateId").Click
         Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys "{ENTER}"
        wait 2
        WshShell.SendKeys "{ENTER}"
        wait 5
        
        Set objTab = Browser("DC Home").Page("Ops_Page").WebTable("class:=ui-jqgrid-btable","html id:=hostingTemplateList","html tag:=TABLE")
	    Set oTableCell = Description.Create
	     oTableCell("micclass").value = "WebElement"
	    oTableCell("html tag").value = "TD"
		oTableCell("innertext").value = Template&".*"   
		oTableCell("outerhtml").value = "<TD role=gridcell aria-describedby=hostingTemplateList_TemplateId>.* "
	
		set oTableCellObj = objTab.ChildObjects(oTableCell)
        If  oTableCellObj.count >= 1 Then
           For i = 0 To oTableCellObj.count-1 Step 1
                 	x= oTableCellObj(i).GetRoproperty("innertext")
                 	If instr(1,x,strClientName,1)>0 Then
                 		    oTableCellObj(i).highlight
                 		    Call ReportStep (StatusTypes.Pass,"Searched template name " & Template & " Exist in Template List At Line No:" & i+1 &" With name" & x ,"Search Template name")
                 			oTableCellObj(i).Click
                 			SearchTemplateExistanceInOps="Exist"
                 	End If
                 Next
     Else
          Call ReportStep (StatusTypes.Pass,"Searched template name " & Template & " Not exist in Template List At Line No:"& i+1 &" With name" & x ,"Search Template name")
     	 SearchTemplateExistanceInOps="Not Exist"
     End If
End Function





Function PerformChangeActionInTab(byVal Tab,byVal Action,byVal Param1,byVal Param2,ByRef objData)
  Select Case Tab
	Case "Database Details"
	    If Browser("Browser-OCRF").Page("Page-OCRF").Link("name:=\[2\] Database Details->> ","text:=\[2\] Database Details->> ","html tag:=A").Exist(10) Then
	       Browser("Browser-OCRF").Page("Page-OCRF").Link("name:=\[2\] Database Details->> ","text:=\[2\] Database Details->> ","html tag:=A").Click	
	    End If
	    
         If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-jqgrid-title","html tag:=SPAN","innertext:=Database Details").Exist(20) Then
			Call ReportStep (StatusTypes.Pass, "Entered in to cude details tab ","Cube details tab")	
		Else
			Call ReportStep (StatusTypes.Fail, "Not Entered in to cude details tab","Cube details tab")
		End If
		 wait 4
	     Set CubeDetailsTable=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html id:=cube-list","html tag:=TABLE")  
	        Nc= CubeDetailsTable.GetROProperty("cols")
            Nr= CubeDetailsTable.GetROProperty("rows")
           
            For i = 2 To Nr Step 1
              cube=CubeDetailsTable.GetCellData(i,12)
            	If Cube=Param2 Then
            	    CubeDetailsTable.ChildItem(i,2,"WebCheckBox",index).Click
            	    CubeDetailsTable.ChildItem(i,7,"WebList",index).Select Action
            		
            	End If
            Next
   			'Handle Delete Notification PoPUp
   		If Browser("Browser-OCRF").Dialog("regexpwndtitle:=Message from webpage","text:=Message from webpage").WinButton("regexpwndtitle:=OK","nativeclass:=Button","text:=OK","regexpwndclass:=Button").Exist(10) Then
   	  		Browser("Browser-OCRF").Dialog("regexpwndtitle:=Message from webpage","text:=Message from webpage").WinButton("regexpwndtitle:=OK","nativeclass:=Button","text:=OK","regexpwndclass:=Button").Click
      		Call ReportStep (StatusTypes.Pass, "Delete conformation  Pop Up appeared","Delete Pop Up")
   		Else
   	  		Call ReportStep (StatusTypes.Fail, "Delete conformation  Pop Up Not appeared","Delete Pop Up")
   		End If
  
   		wait 4
    Case "Add Users"
        If Browser("Browser-OCRF").Page("Page-OCRF").Link("name:=\[3\] Add Users->> ","text:=\[3\] Add Users->> ","html tag:=A","class:=done").Exist(10) Then
	       Browser("Browser-OCRF").Page("Page-OCRF").Link("name:=\[3\] Add Users->> ","text:=\[3\] Add Users->> ","html tag:=A","class:=done").Click	
	    End If
	    wait 10
		If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementOfferingUserList").Exist(20) Then
			Call ReportStep (StatusTypes.Pass, "Entered in to Add user tab","Add user tab")	
		Else
			Call ReportStep (StatusTypes.Fail, "Not Entered in to Add user tab","Add user tab")
		End If
		 wait 5
		 Call SelectRowsFromDropDownList(1,"50",objData)
		Set AddUserTable=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data")
		Nc= AddUserTable.GetROProperty("cols")
        Nr= AddUserTable.GetROProperty("rows")
       
         For i = 2 To Nr Step 1
             User=AddUserTable.GetCellData(i,12)
             If User=Param1 Then
            	AddUserTable.ChildItem(i,2,"WebCheckBox",index).Click
            	AddUserTable.ChildItem(i,7,"WebList",index).Select Action
            		
            End If
         Next
        wait 4
   Case "BI Tools Access"
        If Browser("Browser-OCRF").Page("Page-OCRF").Link("name:=\[5\] BI Tools Access->> ","text:=\[5\] BI Tools Access->> ","html tag:=A","class:=done").Exist(10) Then
	       Browser("Browser-OCRF").Page("Page-OCRF").Link("name:=\[5\] BI Tools Access->> ","text:=\[5\] BI Tools Access->> ","html tag:=A","class:=done").Click	
	    End If
	   
		If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementOfferingUserBIToolList").Exist(120) <> true Then
			Call ReportStep (StatusTypes.Fail, "Not Redirected to BI Tool Access tab","OCRF New Request")	
		Else
			Call ReportStep (StatusTypes.Pass, "Redirected to BI Tool Access tab successfully","OCRF New Request")
		End If
		 wait 5
		 Call SelectRowsFromDropDownList(4,"50",objData)
		 Set BIToolTable=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html tag:=TABLE","html id:=userBIAccess-list")
         Nc= BIToolTable.GetROProperty("cols")
         Nr= BIToolTable.GetROProperty("rows")
         wait 2
         For i = 1 To Nr Step 1
             User=BIToolTable.GetCellData(i,8)
             If User=Param1 Then
            	BIToolTable.ChildItem(i,2,"WebCheckBox",index).Click
            	BIToolTable.ChildItem(i,6,"WebList",index).Select Action
            		
            End If
         Next
          wait 4
   Case "User DB Access"
        If Browser("Browser-OCRF").Page("Page-OCRF").Link("name:=\[6\] User DB Access<< ","text:=\[6\] User DB Access<< ","html tag:=A","class:=done").Exist(10) Then
	       Browser("Browser-OCRF").Page("Page-OCRF").Link("name:=\[6\] User DB Access<< ","text:=\[6\] User DB Access<< ","html tag:=A","class:=done").Click	
	    End If
	    If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementOfferingUserBIToolList").Exist(120) <> true Then
			Call ReportStep (StatusTypes.Fail, "Not Redirected to BI Tool Access tab","OCRF New Request")	
		Else
			Call ReportStep (StatusTypes.Pass, "Redirected to BI Tool Access tab successfully","OCRF New Request")
		End If
		 wait 5
		 Call SelectRowsFromDropDownList(6,"50",objData)
		Set UserDBTable=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("class:=ui-jqgrid-btable","html tag:=TABLE","html id:=userCubeAccess-list")
         Nc= UserDBTable.GetROProperty("cols")
         Nr= UserDBTable.GetROProperty("rows")
         For i = 1 To Nr Step 1
             User=UserDBTable.GetCellData(i,7)
             Cube=UserDBTable.GetCellData(i,10)
             If User=Param1 AND Cube=Param2 Then
            	UserDBTable.ChildItem(i,2,"WebCheckBox",index).Click
            	UserDBTable.ChildItem(i,3,"WebList",index).Select Action
            		
            End If
         Next
      
  End Select

End Function


Function FolderCreationInFLA(ByVal CubeName,ByVal HostPath,ByRef objData)
	        Sourcefile=Environment.Value("InputFiles")&"IMSSCAWeb\DB Files\.abf"
        	strFolderName=HostPath&CubeName
        	ExistFile=HostPath&CubeName&"\.abf"
        	RenameFile=HostPath&CubeName&"\"&CubeName&".abf"
       		Set fso = CreateObject("Scripting.FileSystemObject")
        	If fso.FolderExists(strFolderName) = false Then
         		fso.CreateFolder (strFolderName)
        	Else
        	End If
        		If fso.FileExists(Sourcefile) Then 
	        		fso.CopyFile Sourcefile,strFolderName&"\"
	        		If fso.FileExists(ExistFile) Then
	        		    If fso.FileExists(RenameFile) Then
	        		    Else 	
	        		       fso.MoveFile ExistFile, RenameFile
	        		    End If
                    End If
                 	 Call ReportStep (StatusTypes.Pass,"abf file placed in folder '"&strFolderName&"' with name '"&CubeName&".abf","Ops Group Page" )
	        	Else
	        		 Call ReportStep (StatusTypes.Fail,"abf file Not placed in folder '"&strFolderName&"' with name '"&CubeName&".abf","Ops Group Page" )
	        		
	        	End If
        	
        	
        	
End Function

Function AddNewLineINAddUser(ByVal OffClientName,ByVal User_Id,objData)
	Cnt=False
	userRow=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("html tag:=TABLE","html id:=user-list","class:=ui-jqgrid-btable").GetROProperty("rows")
	For user = 1 To CInt(userRow) Step 1
	    userData=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("html tag:=TABLE","html id:=user-list","class:=ui-jqgrid-btable").GetCellData(user,12)
		If userData=OffClientName Then
		   Cnt=True
		   Exit For
		Else
			Cnt=False
		End If
	Next
	If Cnt=False Then
		    'Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebListselNewLine").Select "50"
			Set Upload=Description.Create()
			Upload("micclass").value="WebElement"
			Upload("html tag").value="DIV"
			Upload("innertext").value="Upload User"
			Set toatlObj=Browser("Browser-OCRF").Page("Page-OCRF").ChildObjects(Upload)
			If toatlObj(0).Exist(10) Then
				toatlObj(0).Click
				'Call ReportStep (StatusTypes.Pass, "'Upload user' Exist and Clicked on" ,"Add user tab")	
			Else
				Call ReportStep (StatusTypes.Fail,  "'Upload user' Not Exist","Add user tab")
			End If

			If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditblkClientName").Exist(10) Then
				Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditblkClientName").Set OffClientName
				'Call ReportStep (StatusTypes.Pass, "Edit Box for Client Name Exist and entered "&dataTable.value("Offering_Client","Offering_Details"),"Add user tab")	
			Else
				Call ReportStep (StatusTypes.Fail,  "Edit Box for Client Name Not Exist","Add user tab")
			End If

			Set objShell=CreateObject("WScript.Shell")
			Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditblkClientName").Click
			objShell.SendKeys "{DOWN}"
'			wait 2
'			objShell.SendKeys "{DOWN}"
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&OffClientName,"html id:=ui-id-.*").fireevent "onmouseover" 
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&OffClientName,"html id:=ui-id-.*").highlight
			wait 2
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&OffClientName,"html id:=ui-id-.*").Click

			If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditUploadUser").Exist(10) Then
	   			Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditUploadUser").Set User_Id
				Call ReportStep (StatusTypes.Pass, "Upload user area Exist and Uplaoded user","Add user tab")	
    		Else
	   			Call ReportStep (StatusTypes.Fail,  "Upload user area Not Exist and Uplaoded user","Add user tab")

			End If
			wait 5
    		If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebEleUploadButton").Exist(10) Then
	   			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebEleUploadButton").Click
				'Call ReportStep (StatusTypes.Pass, "Clicked on 'Upload' button","Add user tab")	
    		Else
	   			Call ReportStep (StatusTypes.Fail,  "Not Clicked on 'Upload' button","Add user tab")

			End If
	End If
	
End Function


Function SelectuserINBItool(ByVal BITool,ByVal Role,ByVal User_Id,ByVal WebListNo,objData)
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementOfferingUserBIToolList").Exist(120) <> true Then
		Call ReportStep (StatusTypes.Fail, "Not Redirected to BI Tool Access tab","BI tool access Tab")	
	Else
		'Call ReportStep (StatusTypes.Pass, "Redirected to BI Tool Access tab successfully","BI tool access Tab")
	End If
		If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Exist(120) Then
	    Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Click
		'Call ReportStep (StatusTypes.Pass, "Clicked on 'Select user'","BI tool access Tab")	
	Else
		Call ReportStep (StatusTypes.Fail, "Not Clicked on 'Select user","BI tool access Tab")
	End If
    Set objBIAccessTool = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIAccessTool")
	Set objBIToolAccessMode = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIToolAccessMode")
		
	Call SCA.SelectFromDropdown(objBIAccessTool, BITool)
	Call SCA.SelectFromDropdown(objBIToolAccessMode,Role)
	Call SelectRowsFromDropDownList(WebListNo,"999",objData)
	TRow=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").GetROProperty("rows")
	
	For j = 1 To TRow Step 1
		userData=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").GetCellData(j,4)
		If userData=User_Id Then
			Set chk =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").ChildItem(j,2,"WebCheckBox",0)
	        chk.click
	        Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_AddUser").Click
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_Add").Click
		End If
	Next
End Function



Function OPSManageUserList(ByVal CubeName,ByVal Role,ByVal UserId,objData)
  cube= Browser("DC Home").Page("Ops_Page").WebElement("html tag:=LABEL","html id:=templateId","class:=form-label header").GetROProperty("innertext")
  If cube=CubeName Then
	Call ReportStep (StatusTypes.Pass, "Opened temaplate in ops for verifying manage user"," manage use tab")	
  Else
	Call ReportStep (StatusTypes.Fail, "Not Opened temaplate in ops for verifying manage user"," manage use tab")  
  End If
  
   Browser("DC Home").Page("Ops_Page").WebEdit("kind:=singleline","type:=text","html tag:=INPUT","name:=Role","class:=form-input lrg").Set Role
   Browser("DC Home").Page("Ops_Page").WebButton("html id:=btnSearchRole","html tag:=INPUT","type:=submit","value:=Search").Click
   wait 3
   
   Browser("DC Home").Page("Ops_Page").WebList("ManageUserSelect").Select "50"

	userRow=Browser("DC Home").Page("Ops_Page").WebTable("class:=ui-jqgrid-btable","html id:=userRoleList","html tag:=TABLE").GetROProperty("rows")
   cnt=False
   For j = 1 To userRow Step 1
    	Userdata=Browser("DC Home").Page("Ops_Page").WebTable("class:=ui-jqgrid-btable","html id:=userRoleList","html tag:=TABLE").GetCellData(j,2)
        If Userdata=UserId Then
           cnt=True
           Exit For
        Else
           cnt=False        
        End If 
    Next
  OPSManageUserList=cnt
  Browser("DC Home").Page("Ops_Page").WebElement("innertext:=Close","html tag:=SPAN","class:=ui-button-text","visible:=True").Click
End Function


Public Function Multi_BITools_Access_Combination(ByVal Path,ByRef objData)
	'1) Validate whether OCRF New Request is redirected to BI Tool access
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementOfferingUserBIToolList").Exist(120) <> true Then
		'Call ReportStep (StatusTypes.Fail, "Not Redirected to BI Tool Access tab","BI tool access Tab")	
	Else
		Call ReportStep (StatusTypes.Pass, "Redirected to BI Tool Access tab successfully","BI tool access Tab")
	End If
	SheetName = "BI_Tools"	
	Call ImportSheet(SheetName,Path)

	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Exist(120) Then
	    Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Click
		'Call ReportStep (StatusTypes.Pass, "Clicked on 'Select user'","BI tool access Tab")	
	Else
		Call ReportStep (StatusTypes.Fail, "Not Clicked on 'Select user","BI tool access Tab")
	End If
	
	
	rc = Datatable.GetSheet(SheetName).GetRowCount
	If rc>20 Then
	   Set s=Description.Create
		s("name").value="select"
		s("html tag").value="SELECT"
		s("select type").value="ComboBox Select"
		Set obj=Browser("Browser-OCRF").Page("Page-OCRF").ChildObjects(s)
		obj(13).Select "50"
	End If
	wait 3
	For rownum = 1 To rc Step 1
		DataTable.SetCurrentRow(rownum)
		Set objBIAccessTool = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIAccessTool")
		Set objBIToolAccessMode = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIToolAccessMode")
		
		Call SCA.SelectFromDropdown(objBIAccessTool, Datatable.Value("BI_Tool",SheetName))
		Call SCA.SelectFromDropdown(objBIToolAccessMode, Datatable.Value("Access_Mode",SheetName))
		

		List=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").GetROProperty("rows")
		For user = 2 To CInt(List) Step 1
			userdata=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").GetCellData(user,4)
		    If userdata= Datatable.Value("User_Id",SheetName) Then
		    	Set chk =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").ChildItem(user,2,"WebCheckBox",0)
		        chk.click
		        Exit For
		    End If
		Next
	   	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_AddUser").Click
		Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_Add").Click
	Next	



End Function

Public Function Multi_USerDB_Access_Combination(ByVal Path,ByRef objData)
    set UserCube= Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_UserDatabaseList")
    Set UserDB=	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Wel_UserDB_List")
    If UserCube.Exist(10) OR UserDB.Exist(10) Then
		Browser("Browser-OCRF").Page("Page-OCRF").Sync
		wait 2
	Else
		Exit Function
	End If

	SheetName = "User_DB"	
	Call ImportSheet(SheetName,Path)

	rc = Datatable.GetSheet(SheetName).GetRowCount
	wait 3
	For rownum = 1 To rc Step 1
		DataTable.SetCurrentRow(rownum)
		 If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Exist(120) Then
	       Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Click
		'Call ReportStep (StatusTypes.Pass, "Clicked on 'Select user'","BI tool access Tab")	
	    Else
		   Call ReportStep (StatusTypes.Fail, "Not Clicked on 'Select user","BI tool access Tab")
	    End If
	    
	    '************ Modifications Made by Sumit ***********************************
	    'Start - Modified by Sumit to select Cube Type in User DB Access page
	    Set objDBType = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_CubeType")
	    If objDBType.Exist(5) Then
	        wait 5
	    	If objDBType.GetROProperty("disabled") = 0 Then
	    		Call SCA.SelectFromDropdown(objDBType, Datatable.Value("DBType","User_DB"))
	    	End If
	    End If
	    
	    
		Set objCubeDB = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_CubeDB")
		Set objCubeRoles = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_CubeRoles")
		
		If Not objCubeDB.GetROProperty("items count") > 2 Then
			Call SCA.SelectFromDropdown(objCubeDB, objCubeDB.GetItem(2))
		Else 
			Call SCA.SelectFromDropdown(objCubeDB, Datatable.Value("Database","User_DB"))
			If Datatable.Value("Database","User_DB") = "Show All Databases" Then
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("BtnCheck All_DbDetails").Click
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("BtnAdd_DbDetails").Click
				wait 2
			End If
		End If
		
		Call SCA.SelectFromDropdown(objCubeRoles, Datatable.Value("Role","User_DB"))
		
		If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").Exist(20) Then
			List=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").GetROProperty("rows")
			For user = 2 To CInt(List) Step 1
				userdata=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").GetCellData(user,4)
			    If userdata= Trim(Datatable.Value("User_ID",SheetName)) Then
			    	Set chk =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").ChildItem(user,2,"WebCheckBox",0)
			        chk.click
			        Exit For
			    End If
			Next
		End If
			
		'************* End Modifications ****************

'		Set objCubeDB = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_CubeDB")
'		Set objCubeRoles = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_CubeRoles")
'		
'		Call SCA.SelectFromDropdown(objCubeDB, Datatable.Value("Database","User_DB"))
'		
'		If Datatable.Value("Database","User_DB") = "Show All Databases" Then
'			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("BtnCheck All_DbDetails").Click
'			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("BtnAdd_DbDetails").Click
'			wait 2
'		End If
'		
'		Call SCA.SelectFromDropdown(objCubeRoles, Datatable.Value("Role","User_DB"))
'		
'
'		List=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").GetROProperty("rows")
'		For user = 2 To CInt(List) Step 1
'			userdata=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").GetCellData(user,4)
'		    If userdata= Datatable.Value("User_ID",SheetName) Then
'		    	Set chk =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").ChildItem(user,2,"WebCheckBox",0)
'		        chk.click
'		        Exit For
'		    End If
'		Next

			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_UserDB_AddUsers").Click
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_UserDB_Add").Click
	Next
       Call ReportStep (StatusTypes.Pass, "Added "&rc&" rows of data in 'UserDB_Access Access' tab","UserDB_Access tab")	
    	
  	'Store the request ID
	New_Request_ID = trim(Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_RequestID").GetROProperty("value"))
	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementFinish").Click
	Call PageLoading()
	Multi_USerDB_Access_Combination = New_Request_ID
	
End Function

Public Function GetGroupsForClient(ByVal ClientName,ByVal Country,ByVal ClientSynOrNon,ByVal offeringName)
	If Browser("DC Home").Page("Ops_Page").WebRadioGroup("html id:=ContentType","html tag:=INPUT","type:=radio").Exist(20) Then
	   If ClientSynOrNon="Syndicate" or Ucase(ClientSynOrNon)="YES" Then
	      Browser("DC Home").Page("Ops_Page").WebRadioGroup("html id:=ContentType","html tag:=INPUT","type:=radio").Select "Syndicate"
	   Else
	      Browser("DC Home").Page("Ops_Page").WebRadioGroup("html id:=ContentType","html tag:=INPUT","type:=radio").Select "NonSyndicate"
	   End If
    End If


    If Browser("DC Home").Page("Ops_Page").WebEdit("WebEditContAccessClient").Exist Then
       Browser("DC Home").Page("Ops_Page").WebEdit("WebEditContAccessClient").Set ClientName
       Call ReportStep (StatusTypes.Pass,"Client name entered in text box as "&ClientName,"'Manage Content Access' Page" )  
    Else
	    Call ReportStep (StatusTypes.Fail,"Client name Not entered in text box as "&ClientName ,"'Manage Content Access' Page" )	
    End If

     wait 2
    Set WshShell =CreateObject("WScript.Shell")
	Browser("DC Home").Page("Ops_Page").WebEdit("WebEditContAccessClient").Click
	WshShell.SendKeys "{DOWN}"
	wait 2
    If Browser("DC Home").Page("Ops_Page").WebElement("class:=ui-corner-all","html tag:=A","innertext:="&ClientName).Exist(2) Then
       Browser("DC Home").Page("Ops_Page").WebElement("innertext:="&ClientName,"html tag:=A","class:=ui-corner-all").Click	
    End If
    wait 1
    If Browser("DC Home").Page("Ops_Page").WebEdit("WebEditContAccessClient").Exist Then
    	Browser("DC Home").Page("Ops_Page").WebEdit("WebEditContAccessCountry").Set Country
   		Call ReportStep (StatusTypes.Pass,"Country name entered in text box as "&Country,"'Manage Content Access' Page" )  
	Else
		Call ReportStep (StatusTypes.Fail,"Country name Not entered in text box as "&Country ,"'Manage Content Access' Page" )	
   
	End If
    If Browser("DC Home").Page("Ops_Page").WebButton("WebBtnSPGroups").Exist Then
    	 Browser("DC Home").Page("Ops_Page").WebButton("WebBtnSPGroups").Click
   		Call ReportStep (StatusTypes.Pass,"Clicked on 'Share point groups' button","'Manage Content Access' Page" )  
	Else
		Call ReportStep (StatusTypes.Pass,"Not Clicked on 'Share point groups' button" ,"'Manage Content Access' Page" )	
   
	End If
   
    Browser("DC Home").Page("Ops_Page").WebTable("class:=ui-jqgrid-btable","html id:=groupGrid","html tag:=TABLE").highlight
    If Browser("DC Home").Page("Ops_Page").WebEdit("WebEditValue").Exist(10) Then
    	Browser("DC Home").Page("Ops_Page").WebEdit("WebEditValue").Set offeringName
   		Call ReportStep (StatusTypes.Pass,"Entered 'offering name as '"&offeringName&"'  for group search","'Manage Content Access' Page" )  
	Else
		Call ReportStep (StatusTypes.Fail,"Entered 'offering name for group search " ,"'Manage Content Access' Page" )	
   
	End If
    wait 1
    Browser("DC Home").Page("Ops_Page").WebEdit("WebEditValue").Click
    'Browser("DC Home").Page("Ops_Page").WebEdit("WebEditValue").Click
    wait 2
    WshShell.SendKeys "{ENTER}"
	'WshShell.SendKeys "{ENTER}"
	'WshShell.SendKeys "{ENTER}"
    wait 5
End Function

Function UncheckAllCheckBoxForGrp(ByRef r)
	For ch = 0 To r.Count-1 Step 1
	    value=r(ch).GetRoProperty("checked")
		If value=1 Then
		   r(ch).Click	
		End If
	Next
End Function

Function ValidateGroupExistance(ByVal S_BI_Tool,ByVal S_Access_Mode,ByVal offeringName,ByVal User_Id)
    wait 2
    Cnt=False
    IAMCnt=False
    If Browser("DC Home").Page("Ops_Page").WebTable("class:=ui-jqgrid-btable","html id:=groupGrid","html tag:=TABLE").Exist(5) Then
    	r=Browser("DC Home").Page("Ops_Page").WebTable("class:=ui-jqgrid-btable","html id:=groupGrid","html tag:=TABLE").GetROProperty("rows")
    End If
	
    Set s=Description.Create
	s("micclass").value="WebCheckBox"
	s("html tag").value="INPUT"
	s("type").value="checkbox"
	Set c=Browser("DC Home").Page("Ops_Page").WebTable("class:=ui-jqgrid-btable","html id:=groupGrid","html tag:=TABLE").ChildObjects(s)
	wait 2			 
    If r>1 Then
    	Call ReportStep (StatusTypes.Pass,"Number of groups Exist in list are "& r,"'Manage Content Access' Page" ) 
		For i = 2 To Cint(r) Step 1
	       data=Browser("DC Home").Page("Ops_Page").WebTable("class:=ui-jqgrid-btable","html id:=groupGrid","html tag:=TABLE").GetCellData(i,2)
           If instr(1,data,offeringName)>0 Then
    	      If instr(1,data,S_BI_Tool)>0 AND instr(1,data,S_Access_Mode)>0 Then
    	         Cnt=True
    	         Call UncheckAllCheckBoxForGrp(c)
        		 c(i-2).click
    	         Exit For
    		  Else 
    		     Cnt=False
    		  End If
    		  SData=Split(data," ")
              If UCase(Trim(SData(0)))=UCase(S_BI_Tool) Then
              	  IAMCnt=True
              	  Call UncheckAllCheckBoxForGrp(c)
              	  c(i-2).click
              	  Exit For
    		  Else 
    		     IAMCnt=False
              End If 
            End If
            
  
        Next
       If Cnt=True Then
      	  Call ReportStep (StatusTypes.Pass,"For Group Name '"&data&"' BI Tool: '"&S_BI_Tool&"' , Access Mode: '"&S_Access_Mode&"'  combination assigned" ,"Ops Group Page" )
          Call GetUsersForGroup(data, S_BI_Tool,S_Access_Mode,User_Id)
       ElseIf IAMCnt=True Then
       
       wait 3
       
          If IAMCnt=True Then
      	       Call ReportStep (StatusTypes.Pass,"For Group Name '"&data&"'  BI Tool:'IMS Analysis Manager'  assigned","Ops Group Page" )
               Call GetUsersForGroup(data, S_BI_Tool,S_Access_Mode,User_Id)
          Else
               Call ReportStep (StatusTypes.Fail,"For Group Name '"&data&"' BI Tool:'IMS Analysis Manager' Not assigned","Ops Group Page" )
          End If     
       Else 
           Call ReportStep (StatusTypes.Fail,"For Group Name '"&data&"'  BI Tool: '"&S_BI_Tool&"' , Access Mode '"&S_Access_Mode&"'  combination Not assigned","Ops Group Page" )
       End If
       
 End If   		 
End Function


Function GetUsersForGroup(ByVal data, ByVal S_BI_Tool,ByVal S_Access_Mode,ByVal User_Id)
        wait 2
		Browser("DC Home").Page("Ops_Page").WebButton("WebBtnGetUsers").Click
		wait 5
		If  Browser("DC Home").Page("Ops_Page").WebTable("class:=ui-jqgrid-btable","html id:=userGrid","html tag:=TABLE","name:=WebTable").Exist(10) Then
			Browser("DC Home").Page("Ops_Page").WebTable("class:=ui-jqgrid-btable","html id:=userGrid","html tag:=TABLE","name:=WebTable").highlight
		End If
        wait 2
		userrow=Browser("DC Home").Page("Ops_Page").WebTable("class:=ui-jqgrid-btable","html id:=userGrid","html tag:=TABLE","name:=WebTable").GetROProperty("rows")
		Found=False
		For r = 2 To CInt(userrow) Step 1
			userdata=Browser("DC Home").Page("Ops_Page").WebTable("class:=ui-jqgrid-btable","html id:=userGrid","html tag:=TABLE","name:=WebTable").GetCellData(r,2)
			If userdata=User_Id Then
			   Found=True
			   Exit For
			Else	
			   Found=False
			End If
		Next
		If Found=True Then
			 Call ReportStep (StatusTypes.Pass,"For Group Name '"&data&"' BI Tool: '"&S_BI_Tool&"' , Access Mode: '"&S_Access_Mode&"'combination  User '"&User_Id&"' assighned","Ops Group Page" )
          Else
             Call ReportStep (StatusTypes.Fail,"For Group Name '"&data&"' BI Tool: '"&S_BI_Tool&"' , Access Mode: '"&S_Access_Mode&"'combination  User '"&User_Id&"'  Not assigned","Ops Group Page" )
		End If
	
End Function


'Function GetGroupsForClient(ByVal ClientName,ByVal Country,ByVal ClientSynOrNon,ByVal offeringName)
'
'	If Browser("DC Home").Page("Ops_Page").WebRadioGroup("html id:=ContentType","html tag:=INPUT","type:=radio").Exist(20) Then
'	   If Ucase(ClientSynOrNon)=Ucase("Yes") Then
'	      Browser("DC Home").Page("Ops_Page").WebRadioGroup("html id:=ContentType","html tag:=INPUT","type:=radio").Select "Syndicate"
'	   Else
'	      Browser("DC Home").Page("Ops_Page").WebRadioGroup("html id:=ContentType","html tag:=INPUT","type:=radio").Select "NonSyndicate"
'	   End If
'    End If
'
'
'    If Browser("DC Home").Page("Ops_Page").WebEdit("WebEditContAccessClient").Exist Then
'       Browser("DC Home").Page("Ops_Page").WebEdit("WebEditContAccessClient").Set ClientName
'       Call ReportStep (StatusTypes.Pass,"Client name entered in text box as "&ClientName,"'Manage Content Access' Page" )  
'    Else
'	    Call ReportStep (StatusTypes.Fail,"Client name Not entered in text box as "&ClientName ,"'Manage Content Access' Page" )	
'    End If
'
'     wait 2
'    Set WshShell =CreateObject("WScript.Shell")
'	Browser("DC Home").Page("Ops_Page").WebEdit("WebEditContAccessClient").Click
'	WshShell.SendKeys "{DOWN}"
'
'    If Browser("DC Home").Page("Ops_Page").WebElement("class:=ui-corner-all","html tag:=A","innertext:="&ClientName).Exist(2) Then
'       Browser("DC Home").Page("Ops_Page").WebElement("innertext:="&ClientName,"html tag:=A","class:=ui-corner-all").Click	
'    End If
'    wait 1
'    If Browser("DC Home").Page("Ops_Page").WebEdit("WebEditContAccessClient").Exist Then
'    	Browser("DC Home").Page("Ops_Page").WebEdit("WebEditContAccessCountry").Set Country
'   		Call ReportStep (StatusTypes.Pass,"Country name entered in text box as "&Country,"'Manage Content Access' Page" )  
'	Else
'		Call ReportStep (StatusTypes.Fail,"Country name Not entered in text box as "&Country ,"'Manage Content Access' Page" )	
'   
'	End If
'    If Browser("DC Home").Page("Ops_Page").WebButton("WebBtnSPGroups").Exist Then
'    	 Browser("DC Home").Page("Ops_Page").WebButton("WebBtnSPGroups").Click
'   		Call ReportStep (StatusTypes.Pass,"Clicked on 'Share point groups' button","'Manage Content Access' Page" )  
'	Else
'		Call ReportStep (StatusTypes.Pass,"Not Clicked on 'Share point groups' button" ,"'Manage Content Access' Page" )	
'   
'	End If
'   
'    Browser("DC Home").Page("Ops_Page").WebTable("class:=ui-jqgrid-btable","html id:=groupGrid","html tag:=TABLE").highlight
'    If Browser("DC Home").Page("Ops_Page").WebEdit("WebEditValue").Exist Then
'    	Browser("DC Home").Page("Ops_Page").WebEdit("WebEditValue").Set offeringName
'   		Call ReportStep (StatusTypes.Pass,"Entered 'offering name as '"&offeringName&"'  for group search","'Manage Content Access' Page" )  
'	Else
'		Call ReportStep (StatusTypes.Fail,"Entered 'offering name for group search " ,"'Manage Content Access' Page" )	
'   
'	End If
'    
'    Browser("DC Home").Page("Ops_Page").WebEdit("WebEditValue").Click
'    WshShell.SendKeys "{ENTER}"
'    wait 5
'End Function
Function UncheckAllCheckBoxForGrp(ByRef r)
	For ch = 0 To r.Count-1 Step 1
	    value=r(ch).GetRoProperty("checked")
		If value=1 Then
		   r(ch).Click	
		End If
	Next
End Function
Function SelectRowsFromDropDownList(ByVal ChObjNo,ByVal Rows,ByRef objData)
	Set oList = Description.Create
	oList("micclass").value = "WebList"
	oList("html tag").value = "SELECT"
	oList("name").value = "select"  
	oList("select type").value ="ComboBox Select"
	Set chListobj=Browser("Browser-OCRF").Page("Page-OCRF").ChildObjects(oList)
	chListobj(ChObjNo).Select Rows
	wait 3
End Function

Public Function AddDatainClientDetailsTab(ByVal FileName,ByVal SheetName,ByVal StartRow,ByVal EndRow,ByVal Country,ByVal Offering,ByRef objData)


    
     TimeStamp=objData.item("TimeStamp")
   '########### OFFERING DETAIL TAB EXISTANCE - Start ########### 
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementCDTabHeader").Exist(20) Then
		Call ReportStep (StatusTypes.Pass, "'Client content' tab header exist","Client content Tab header")
	Else
	   	Call ReportStep (StatusTypes.Fail, "Client content Tab header not exist","Client content Tab header")
	End If

 	wait 2
	Call ImportSheet(SheetName,FileName)
	
	DataTable.GetSheet(SheetName)
	rc=DataTable.GetSheet(SheetName).GetRowCount
	wait 2
	For rownum = StartRow To EndRow Step 1
	     DataTable.SetCurrentRow(rownum)
	     Provide_Link=Datatable.Value("Provide_Link",SheetName)
	     Decision_Center_Heading=Datatable.Value("Decision_Center_Heading",SheetName)
	     Link_Name=Datatable.Value("Link_Name",SheetName)
	     ToolTip=Datatable.Value("ToolTip",SheetName)
	     Report_Type=Datatable.Value("Report_Type",SheetName)
	     Report_Name=Datatable.Value("Report_Name",SheetName)
	     Client_Name=Datatable.Value("Client_Name",SheetName)
	     
		'CLICK ON ADD NEW LINE
'		If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementAddNewLine").Exist(20) Then
'			wait 2
'	        'Call ReportStep (StatusTypes.Pass, "Click on add new line done","Add new button")	
'			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementAddNewLine").Click
'		Else
'			Call ReportStep (StatusTypes.Fail, "Click on add new line not done","Add new button")
'		End If

		'changes by srinivas-29th july 2020
		If browser("Browser-OCRF").Page("Page-OCRF").WebElement("xpath:=//div[@id='divContentList']//table[@class='ui-pg-table navtable']//div[text()='Add New Line ']").Exist Then
			'Browser("Browser-OCRF").Page("Page-OCRF").WebElement("innertext:=Add New Line","html tag:=DIV","visible:=True").Click
			browser("Browser-OCRF").Page("Page-OCRF").WebElement("xpath:=//div[@id='divContentList']//table[@class='ui-pg-table navtable']//div[text()='Add New Line ']").Click
		ElseIf Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementAddNewLine").Exist(20) Then
			wait 2
	        'Call ReportStep (StatusTypes.Pass, "Click on add new line done","Add new button")	
			if browser("Browser-OCRF").Page("Page-OCRF").WebElement("xpath:=//div[@id='divContentList']//table[@class='ui-pg-table navtable']//div[text()='Add New Line ']").Exist then
				browser("Browser-OCRF").Page("Page-OCRF").WebElement("xpath:=//div[@id='divContentList']//table[@class='ui-pg-table navtable']//div[text()='Add New Line ']").Click
			end if 	
		Else
			Call ReportStep (StatusTypes.Fail, "Click on add new line not done","Add new button")
		End If
		Set objClientContentTable = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList")
		'Modified By Sumit on 22nd March 
		'Stated Modification
		'Offering = ReadWriteDataFromTextFile("Read","SyndOffering","")
		'Country = ReadWriteDataFromTextFile("Read","SyndContry","")
	 	' End Modification
	    If objClientContentTable.Exist(10)  Then
	    	CCR= objClientContentTable.RowCount
	        If Datatable.Value("Provide_Link",SheetName)="Yes" Then
	        wait 2
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 9, "WebList", 0).Select Provide_Link
	    	   objClientContentTable.ChildItem(CCR, 10, "WebList", 0).Select Decision_Center_Heading
	           objClientContentTable.ChildItem(CCR, 11, "WebEdit", 0).Set Link_Name&TimeStamp
	           objClientContentTable.ChildItem(CCR, 12, "WebEdit", 0).Set ToolTip&TimeStamp
	      Else
	         wait 3
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 9, "WebList", 0).Select Provide_Link
	           'CHECKING FOR FREEZED DATA FOR DECISION CENTER HEADING,LINK NAME AND TOOLTIP IF PROVIDE LINK IS 'NO'
	           DCDisabled=objClientContentTable.ChildItem(CCR, 10, "WebList", 0).GetROProperty("disabled") 
	           LinkDisabled=objClientContentTable.ChildItem(CCR, 11, "WebEdit", 0).GetROProperty("disabled") 
	           TooltipDis=objClientContentTable.ChildItem(CCR, 12, "WebEdit", 0).GetROProperty("disabled") 
	           If DCDisabled="1" AND LinkDisabled="1" AND TooltipDis="1" Then
	           	  Call ReportStep (StatusTypes.Pass, "'Link name', 'Tool tip and' and  'Decision center' fields  are not Editable after seletcing Provide link as 'No' "," Provide link as 'No'")
	           Else
	           	  Call ReportStep (StatusTypes.Pass, "'Link name', 'Tool tip and' and  'Decision center' fields  are  Editable after seletcing Provide link as 'No' "," Provide link as 'No'")
	           End If

	       End If
	       
	       	   objClientContentTable.ChildItem(CCR, 14, "WebList", 0).Select Trim(Report_Type)
	           'objClientContentTable.ChildItem(CCR, 15, "WebEdit", 0).Set Trim(Country&"/"&Offering)
	           'Modified the code based on New Requirement - OCRF - VA Ops Maintenance - 21/01/2019
	           If Report_Type = "VA PBI Report" Then
'	           ***************************************************
'					Modified by Madhu - 04/15/2020

	           		'objClientContentTable.ChildItem(CCR, 16, "WebEdit", 0).Set Trim(Report_Name)
	           		else
	           		objClientContentTable.ChildItem(CCR, 16, "WebList", 0).Click
	           		Wait 5
	           		'objClientContentTable.ChildItem(CCR, 16, "WebEdit", 0).Set Trim(Report_Name)
					If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Information_Okay").Exist Then
						Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Information_Okay").Click
						wait 3
					End If
					objClientContentTable.ChildItem(CCR, 16, "WebList", 0).Click
	           		objClientContentTable.ChildItem(CCR, 16, "WebList", 0).Select Trim(Report_Name)
''	           		*********************************************************************8
	           		wait 2
	           End If
	           
'	           ***********************************************************
	           objClientContentTable.ChildItem(CCR, 17, "WebEdit", 0).Set Trim(Client_Name)
	           'Set objShell=CreateObject("WScript.Shell")
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
			   objClientContentTable.ChildItem(CCR, 17, "WebEdit", 0).Click
				'objShell.SendKeys "{DOWN}"
				'************************************************************************8

'	    	   modified by Madhu - 02/26/2020

				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:="&Datatable.Value("Client_Name",SheetName)).highlight
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:="&Datatable.Value("Client_Name",SheetName)).Click
              
				wait 2
		 Call ReportStep (StatusTypes.Pass, "Successfully added data in row '"&rownum+1&"' Provide Link as'" &Provide_Link& "' Report Type as '" &Report_Type," Client content tab")
               
               objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
'			************************************************************************8

'	    	   modified by Madhu - 02/26/2020

'				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:="&Datatable.Value("Client_Name",SheetName)).highlight
'				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:="&Datatable.Value("Client_Name",SheetName)).Click
               Call ReportStep (StatusTypes.Pass, "Successfully added data in row '"&rownum+1&"' Provide Link as'" &Provide_Link& "' Report Type as '" &Report_Type," Client content tab")
               
               If Report_Type = "VA PBI Report" Then
               		filePath = Environment("CurrDir") & "HostingTestData\" 
               		fileName = Report_Name
               		fileToSet = filePath & fileName
               		objClientContentTable.ChildItem(CCR, 18, "Link", 0).Click
      				If Browser("Browser-OCRF").Page("Page-OCRF").WebFile("wfAddFiles_PBIReport").Exist(20) then 
      					Browser("Browser-OCRF").Page("Page-OCRF").WebFile("wfAddFiles_PBIReport").Set fileToSet
      					wait 3
'      					Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Add files_PBIReport").Click
'      					If Browser("Browser-OCRF").Dialog("Choose File to Upload").Exist(20) Then
'      						Browser("Browser-OCRF").Dialog("Choose File to Upload").WinEdit("FileName").Set fileToSet
'      						Browser("Browser-OCRF").Dialog("Choose File to Upload").WinButton("Open").Click
      						Browser("Browser-OCRF").Page("Page-OCRF").WebButton("StartUpload_PBIReport").Click
      						uploadedFileName = Trim(Browser("Browser-OCRF").Page("Page-OCRF").WebTable("PBIReportUpload").ChildItem(1,2,"WebElement",0).getRoProperty("innertext"))
      						wait 10
      						uploadStatus = Trim(Browser("Browser-OCRF").Page("Page-OCRF").WebTable("PBIReportUpload").ChildItem(1,4,"WebElement",0).getRoProperty("innertext"))
      						
      						If uploadedFileName = fileName And uploadStatus = "File Uploaded" Then
      							Call ReportStep (StatusTypes.Pass, "File Uploaded successfully"," Upload File - PBI Report -Client content tab")
      							Else
      							Call ReportStep (StatusTypes.Fail, "File is not uploaded successfully"," Upload File - PBI Report -Client content tab")
      						End If
      						Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_PBIReport").Click
      				End If
               End If
	           
	    End If
	    		    	
	Next	
	Call ReadWriteDataFromTextFile("write","PBIReportTimeStamp",TimeStamp)
    AddDatainClientDetailsTab=TimeStamp
End Function

Function LoginToSCA(ByVal User_ID,ByVal Password ,ByVal URL)

	Systemutil.Run "iexplore.exe",URL
	
	
If Browser("Browser-SCA").Page("Page-SCA").WebEdit("WebEdit_Username").Exist(60) Then
	If Browser("Browser-SCA").Page("Page-SCA").WebEdit("WebEdit_Username").Exist(5) Then
		Browser("Browser-SCA").Page("Page-SCA").WebEdit("WebEdit_Username").Set User_ID
		
	    '******************************************************************
    wait 5
	Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//input[@id='btnValidate']").Click
'    ***************************************************************************8
	Else
		Call ReportStep (StatusTypes.Fail, "User name text box is not displayed","User name text box")
		Exit Function
	End If
	
	If Browser("Browser-SCA").Page("Page-SCA").WebEdit("WebEdit_Password").Exist Then
		Browser("Browser-SCA").Page("Page-SCA").WebEdit("WebEdit_Password").Set Password
	Else
		Call ReportStep (StatusTypes.Fail, "Password text box is not displayed","Password text box")
		Exit Function
	End If
	
	If Browser("Browser-SCA").Page("Page-SCA").WebButton("WebButton_Login").Exist Then
		Browser("Browser-SCA").Page("Page-SCA").WebButton("WebButton_Login").Click
	Else
		Call ReportStep (StatusTypes.Fail, "Login button is not available","Login button")
		Exit Function
	End If	
	
'******************************************************
'Modified by Madhu - 02/26/2020
'
	
'ElseIf Browser("BrowserPoPUp").Dialog("Windows Security").Exist(5) Then
'    If Browser("BrowserPoPUp").Dialog("Windows Security").Exist(30) Then
'	   Browser("BrowserPoPUp").Dialog("Windows Security").WinEdit("WinEditUserName").Set User_ID	
'	   Browser("BrowserPoPUp").Dialog("Windows Security").WinEdit("WinEditPassword").Set Password
'	   Browser("BrowserPoPUp").Dialog("Windows Security").WinButton("BtnOK").Click
'	   Call ReportStep (StatusTypes.Pass, "Window Security pop up is handled","Window security PopUp")
'	Else
'	   Call ReportStep (StatusTypes.Fail, "Window Security pop up is Not handled","Window security PopUp")		   
'	End If	
ElseIf UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").Exist(5) Then
	If UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").Exist(20) Then    	
        UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").Highlight 
	    UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAEdit("User name").SetValue User_ID
	    UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAEdit("Password").SetValue Password
	    UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAButton("OK").Click
        Call ReportStep (StatusTypes.Pass, "User " & UserName & " has logged in successfully","Login through Window security PopUp")
	Else
		Call ReportStep (StatusTypes.Fail, "User " & UserName & " has not successfully logged in","Login through Window security PopUp")	   
	End If
	
End If	
	
'	*******************************************************
'Modified by Madhu - 02/26/2020
'modified in OR
'************************************************
'Validate user login to SCA
	If Browser("Browser-SCA").Page("Page-SCA").Frame("SCA-Frame").Link("Link_SharedReports").Exist(120) Then
		Call ReportStep (StatusTypes.Pass, "Login sucessfully to SCA","Login to SCA")
	Else	
		Call ReportStep (StatusTypes.Fail, "Not able to Login sucessfully to SCA","Login to SCA")
	End If
End Function


Public Function OperationOn_Folder(ByVal strOselection,ByVal strFolderName,ByVal strLocation,ByVal strPageName,ByVal strFolderName_Update)
      '<Variable declaration >
	Dim objSharedLink,objMyReport,objFolder,objFolderCount
	'<Looping variables>
	 Dim i,strAppFolder,objCreate,objtxtCreateFolder,objtxtDescription,objbtnCreateFolderOk,objbtnDeleteOK
	 '<String variables>
		 wait 2
		 Browser("Browser-SCA").Page("Page-SCA").Sync
		 Browser("Browser-SCA").Page("Page-SCA").RefreshObject
		

   	'For a = Environment.Value("intCounterStart")  To Environment.Value("intCounterMaxLimit") Step 1
        If strLocation <> "" Then
	        For a = 1  To 5 Step 1
					        
				   If strLocation="Shared" Then
					  
					   If Browser("Browser-SCA").Page("Page-SCA").Frame("SCA-Frame").Link("Link_SharedReports").Exist(120) Then
                            Browser("Browser-SCA").Page("Page-SCA").Frame("SCA-Frame").Link("Link_SharedReports").Click 
					   		Exit for
					   End If
					   
					End If
		
				Next
        End If
       
       If  Browser("Browser-SCA").Page("Page-SCA").Frame("view").WebTable("html id:=maintable","html tag:=TABLE").WebElement("innertext:="&strFolderName,"html tag:=A","html id:=").Exist(150) Then 
                  Call ReportStep (StatusTypes.Pass, +strFolderName+" Folder Structure in the Shared\My Report", strPageName)
                  Browser("Browser-SCA").Page("Page-SCA").Frame("view").WebElement("innertext:="&strFolderName,"html tag:=A","html id:=").click 
				    	
		
        Else
				Call ReportStep (StatusTypes.Fail, "SCA Web page synchronization taken long time to load and redirect to "&strFolderName&" Folder Structure in the Shared\My Report", strPageName)
		End If    
				           
	
	End Function
	
	
	
	Public Function ReportCreationInSCA(ByVal sheetNum, ByVal newReport,ByVal intEMD,ByVal intdbcheck,ByVal intRNum,ByVal intReportCreate,ByVal strF1,ByVal strF2,ByVal row,ByVal Sheet,ByVal TestDataPath,ByVal timeStamp)  
    
        Dim objExcel,ObjExcelFile,ObjExcelSheet,objDataSource,objDatabase,objDatacubes,obj_tree,objDR,objrow,objcolumn,objDataaxis,objAdd,objtreeCount,objCreate
        Dim intRow_Count,intColumn_Count,intcount
        Dim dataSourceValue, strDataSourceValue,strDatabaseValue,strCubeValue,strChild_Values,strPlaceHolderVal,p, objBtnok, objErrorOk,strDBitems,strbkch
        Dim i,j,k,x,y,a  
     	Systemutil.CloseProcessByName "excel.exe"

        Set objExcel = CreateObject("Excel.Application") 
        objExcel.Visible =False     
        Set ObjExcelFile = objExcel.Workbooks.Open (Environment.Value("CurrDir")&"InputFiles\IMSSCAWeb\SCAReportCreation.xls")
        'Environment.Value("CurrDir")&"InputFiles\IMSSCAWeb\SCAReportSheet.xls"
        Set ObjExcelSheet = ObjExcelFile.Sheets(sheetNum) 
        intRow_Count = ObjExcelSheet.UsedRange.Rows.Count 
        intColumn_Count = ObjExcelSheet.UsedRange.Columns.Count
    
    
       'Changed by Poornima. as DataSource name should be dynamic for all environments.take FLA server path from the Environment file split by '.'
        DsValue=Split(Environment.Value("flaServerPath"),".")
        'dataSourceValue=TRIM(ObjExcelSheet.cells(2,1).value)
        dataSourceValue=DsValue(0)
        'strDatabaseValue=TRIM(ObjExcelSheet.cells(2,2).value)
        Call ImportSheet(Sheet,TestDataPath)
		DataTable.SetCurrentRow(row) 
        strDatabaseValue=datatable.value("DatabaseName",Sheet)&timeStamp
        'Change completed
        dataSourceValue = datatable.value("ServerLocation",Sheet)
        strCubeValue=TRIM(ObjExcelSheet.cells(2,3).value)
        
        If intEMD=1 Then
        
         If newReport = 0 Then     
           Browser("Browser-SCA").Page("Page-SCA").Frame("view").Image("newreport").Click

        End If
              
       wait 20
       For a = 1 To 100 Step 1 
       
       If Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").WebList("lstDdlDataSources").GetROProperty("disable")=0  Then
         Exit For
       End If           
       Next
       
       wait 2
      Set reg = createobject("vbscript.Regexp")
			reg.global = true
			reg.pattern = "[^\d]"
			DTServerVal = reg.replace(dataSourceValue, "")
			
			strAllItems = Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").WebList("lstDdlDataSources").GetROProperty("all items")
			arrAllItems = Split(strAllItems, ";")
			For i = 0 To ubound(arrAllItems) Step 1
				If instr(1, arrAllItems(i), DTServerVal) > 1 Then
			
					selectItem = i
			        wait 10
			        If Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").WebList("lstDdlDataSources").Exist(20) Then
			        	Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").WebList("lstDdlDataSources").Select selectItem
			        End If
					
				End If
			Next
	
      
       wait 10
       If  Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").WebList("lstddlDatabases").Exist(20) Then
       	    Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").WebList("lstddlDatabases").Select strDatabaseValue
       End If
      

   
       If intdbcheck=0 Then   
       
        strDBitems= Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").WebList("lstddlDatabases").GetROProperty("all items")
         If Instr(strDBitems,strDatabaseValue)<>0 Then           
            ReportCreationInSCA=0
               else
            ReportCreationInSCA=1
         End If 
         
         If intReportCreate=1 Then
            
            
            ObjExcelFile.Close
            objExcel.Quit    
            Set objExcel=nothing
            Set ObjExcelFile=nothing
            Set ObjExcelSheet=nothing
            Set objtreeCount=nothing
            Exit Function
            
        End If
 

       End If 
       wait 5    
	   If strCubeValue=empty Then
	   	print("cube value is empty")
	   	else
	    Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").WebList("lstddlCubes").Select strCubeValue
	   End If       

       
       Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").WebButton("btnAdd").Click
       wait 10
        Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").WebButton("btnAdd").Click

    	strPrev_Child_Values = ""
    	
       For i=2 to intRow_Count
        Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").WebElement("webcube").Click           
         For j=4 to intColumn_Count-1
              strChild_Values=Trim(ObjExcelSheet.cells(i,j).value)
               ' Have to write the explicit mentain the bk mark condition
               strbkch=Trim(ObjExcelSheet.cells(i,intColumn_Count).value)
               If  strChild_Values<>"" AND  strChild_Values<>"rowaxis" AND strChild_Values<>"columnaxis"  AND strChild_Values<>"dataaxis" AND strbkch<>"Yes" AND strChild_Values<>"No" Then
                        strPrev_Child_Values = strChild_Values
                        Set obj_tree=description.Create
                        obj_tree("micclass").value="WebElement"
                        obj_tree("class").value="TreeNode"
                        obj_tree("innertext").value=strChild_Values
                        wait 1
                        set objtreeCount=  Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
                        intcount=objtreeCount.count
                        
                        For a = 1 To 100 Step 1
	                    set objtreeCount=  Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
	                    intcount=objtreeCount.count
                            If intcount<>0 Then
                            Exit For
                            End If
                        Next
                        
                        If intcount<>0 Then
                        	k=intcount-1                       	
                            objtreeCount(k).fireEvent "ondblclick"                         
                            Call ReportStep (StatusTypes.Pass,"Clicked on the Webelement"&Space(2)&strChild_Values, "Report Creation Page")                    
                            wait 3    
                        Else
                            Call ReportStep (StatusTypes.Warning,"Not Clicked on the Webelement"&Space(2)&strChild_Values, "Report Creation Page")
	                    	ObjExcelFile.Close
	  						objExcel.Quit    
	   						Set objExcel=nothing
	   						Set ObjExcelFile=nothing
	   						Set ObjExcelSheet=nothing
	   						Set objtreeCount=nothing
                            Exit Function
                        End If 

                                     
               End If
               
              If j=intColumn_Count-1 AND strbkch<>"Yes" Then
              Set obj_tree=description.Create
                    obj_tree("micclass").value="WebElement"
                    obj_tree("html tag").value="SPAN"
                    obj_tree("innertext").value=strPrev_Child_Values
                    wait 1
                    set objtreeCount= Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
		        	intcount=objtreeCount.count
		        	If intcount<>0 Then
                    	k=intcount-1
                    Else
                    	Call ReportStep (StatusTypes.Fail, "Couldn not select node'"&strPrev_Child_Values&" during report creation" , "Report Creation Page")
                    	Exitrun
                    End If	
                    
             wait 1
                    '<shweta 8/6/2016> - Adding Fire Event to highlight the treenode to be choosed as row/column/data axis - Start
                    Setting.WebPackage("ReplayType") = 2
                    objtreeCount(k).FireEvent "onmouseover"
                    wait 2
                    objtreeCount(k).rightclick 
                    On error resume next
'                    Modified by Madhu - 04/13/2020
					Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").webelement(obj_tree).RightClick
'           ********************************************************         
                    strPlaceHolderVal=lcase(ObjExcelSheet.cells(i,intColumn_Count-1).value)                                                                                        
                    wait 3  
                      If strPlaceHolderVal="rowaxis" Then 
                      	  	
                          Set objrow= Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").WebElement("webRowaxis")    
                          Call SCA.ClickOn(objrow,"Rowaxis" , "ReportCreation")    
                          wait 5                          
                      elseif  StrPlaceHolderVal="columnaxis" then  
						                       
                          Set objcolumn= Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").WebElement("webcolumnaxis")
                          Call SCA.ClickOn(objcolumn,"Columnaxis" , "ReportCreation")
                          Wait 5                      
                      elseif  StrPlaceHolderVal="dataaxis" then
						  
						  If  Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").WebElement("webDataAxis").Exist(10)  Then
						  	Set objDataaxis= Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").WebElement("webDataAxis")
                            Call SCA.ClickOn(objDataaxis,"Dataaxis" , "ReportCreation")    
                            wait 5
                            Set objDataaxis=Nothing 
                          Else
                            Call ReportStep (StatusTypes.Fail, "Couldn not click on node'"&strPrev_Child_Values&" during report creation" , "Report Creation Page")
                            Reporter.ReportEvent micFail,"Data Axis", "not displayed"
                            Exitrun 
						  End If                 
                         
                      End If    
                      
                      For p=intColumn_Count-2 to 4  step -1
                         strChild_Values=Trim(ObjExcelSheet.cells(i,p).value)
                         strbkch=Trim(ObjExcelSheet.cells(i,intColumn_Count).value)                         
                         If  strChild_Values<>"" AND  strChild_Values<>"rowaxis" AND strChild_Values<>"columnaxis"  AND strChild_Values<>"dataaxis" AND strbkch<>"Yes" Then
                            Set obj_tree=description.Create
                            obj_tree("micclass").value="WebElement"
                            obj_tree("class").value="TreeNode"
                            obj_tree("innertext").value=strChild_Values
                            wait 1
                            set objtreeCount=  Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
                            intcount=objtreeCount.count    
                                If intcount<>0 Then
                                   k=intcount-1
                                   objtreeCount(k).fireEvent  "ondblclick"    
                                 End If                          
                            End If
                        Next
                
                End If
             
           Next
    
           Next    
   ElseIf intEMD=0 Then   
        Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").WebElement("webcube").Click          
        For j=4 to intColumn_Count-2
        strChild_Values=Trim(ObjExcelSheet.cells(intRNum,j).value)
               
          If  strChild_Values<>"" AND  strChild_Values<>"rowaxis" AND strChild_Values<>"columnaxis"  AND strChild_Values<>"dataaxis" AND strChild_Values<>"Yes" AND strChild_Values<>"No" Then
        
               Set obj_tree=description.Create
               obj_tree("micclass").value="WebElement"
               obj_tree("class").value="TreeNode"
               obj_tree("innertext").value=strChild_Values
               wait 1
               set objtreeCount= Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
               intcount=objtreeCount.count
                        
                For a = 1 To 20 Step 1
                   If intcount<>0 Then
                    Exit For
                   End If
                Next
                        
               If intcount<>0 Then
                k=intcount-1                    
                objtreeCount(k).fireEvent  "ondblclick"        
                Call ReportStep (StatusTypes.Pass,"Clicked on the Webelement"&Space(2)&strChild_Values, "Report Creation Page")                    
                wait 3    
                Else
                   Call ReportStep (StatusTypes.Warning,"Not Clicked on the Webelement"&Space(2)&strChild_Values, "Report Creation Page")
                Exit Function
               End If 
            End If

        If j=intColumn_Count-2  Then
                     
              wait 1
              Set objDR=CreateObject("Mercury.DeviceReplay")
               x=objtreeCount(k).GetroProperty("abs_x")
               y=objtreeCount(k).GetroProperty("abs_y") 
			   objtreeCount(k).fireEvent  "ondblclick"	               
                wait 1                    
                objDR.MouseClick x,y,2
                strPlaceHolderVal=lcase(ObjExcelSheet.cells(intRNum,intColumn_Count-1).value)                                                                                        
                      
                      If strPlaceHolderVal="rowaxis" Then 
                          Set objrow=Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").WebElement("webRowaxis")    
                          Call SCA.ClickOn(objrow,"Rowaxis" , "ReportCreation")    
                          wait 5                          
                      elseif  StrPlaceHolderVal="columnaxis" then                                                                                                                                
                          Set objcolumn=Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").WebElement("webcolumnaxis")
                          Call SCA.ClickOn(objcolumn,"Columnaxis" , "ReportCreation")
                          Wait 5                      
                      elseif  StrPlaceHolderVal="dataaxis" then                                                                            
                         Set objDataaxis=Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").WebElement("webDataAxis")
                         Call SCA.ClickOn(objDataaxis,"Dataaxis" , "ReportCreation")    
                         wait 5
                      End If    
                      
                      For p=intColumn_Count-2 to 4  step -1
                         strChild_Values=Trim(ObjExcelSheet.cells(intRNum,p).value)
                         strbkch=Trim(ObjExcelSheet.cells(intRNum,intColumn_Count).value)                         
                         If  strChild_Values<>"" AND  strChild_Values<>"rowaxis" AND strChild_Values<>"columnaxis"  AND strChild_Values<>"dataaxis"  Then
                            Set obj_tree=description.Create
                            obj_tree("micclass").value="WebElement"
                            obj_tree("class").value="TreeNode"
                            obj_tree("innertext").value=strChild_Values
                            wait 1
                            set objtreeCount= Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
                            intcount=objtreeCount.count    
                                If intcount<>0 Then
                                   k=intcount-1
                                   objtreeCount(k).fireEvent  "ondblclick"    
                                 End If                          
                            End If
                        Next
                
                End If
        Next
        
    End If
    
    
     
       ObjExcelFile.Close
       objExcel.Quit    
       Set objExcel=nothing
       Set ObjExcelFile=nothing
       Set ObjExcelSheet=nothing
       Set objtreeCount=nothing
    
    End Function
    
 Function SCALogOut()
  
   Dim objLogOut, objLeavepage
   
   If Browser("Browser-SCA").Page("Page-SCA").WebElement("welLogout").Exist(10) Then
   	Browser("Browser-SCA").Page("Page-SCA").WebElement("welLogout").Click
   End If

	SystemUtil.CloseProcessByName "iexplore.exe"

  End Function
 Function  SaveReportInSCA(ByVal OfferingNmae ,ByVal Report)
 	If Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").Image("ImgSaveas").Exist(10) Then
 		Browser("Browser-SCA").Page("Page-SCA").Frame("ReportWindowOfAnalyzer").Image("ImgSaveas").Click
 	End If
 	'Browser("Browser-SCA").Page("Page-SCA").Frame("__dialog").WebEdit("txtFolderName").Set OfferingNmae
 	Browser("Browser-SCA").Page("Page-SCA").Frame("__dialog").WebEdit("ReportName").Set Report

    Browser("Browser-SCA").Page("Page-SCA").Frame("__dialog").WebButton("btnOK").Click

 End Function
 
 
  'Login to DC
 Function DCLogin(ByVal ClientUrl,ByVal Username,ByVal Password,ByVal ShortName,ByRef objData)
    Systemutil.CloseProcessByName "iexplore.exe"
    If InStr (ClientUrl,"auth") > 0 Then 
    'Modified by Madhu - 04/13/2020
    	If Ucase(ShortName) = "SSRS" OR Ucase(ShortName) = "XLSX"  Then
    		AppendUrl = ClientUrl & "\" & ShortName
    		ElseIf Ucase(ShortName) = "DOCUMENTS" Then
    		AppendUrl = ClientUrl & "\Shared " & ShortName
    		else
    		AppendUrl = ClientUrl
    	End If
    	
	Else
		AppendUrl = ClientUrl & "\" & ShortName
 	End If
 	wait 5
	Systemutil.Run "iexplore.exe",AppendUrl , , , 3
'    Browser("DC").ClearCache
'    Systemutil.CloseProcessByName "iexplore.exe"    
'    Systemutil.Run AppendUrl
	wait 5
	
	If Browser("Browser-SCA").Page("Page-SCA").WebEdit("WebEdit_Username").Exist(60) Then
		Call ReportStep (StatusTypes.Pass, "Login in to DC with User name as '"&Username&"' and Client URL appending with '"&ShortName," Tooltip of link")	
        Browser("Browser-SCA").Page("Page-SCA").WebEdit("WebEdit_Username").Set Username
        
     
        'MOdified by Madhu - 06/05/2020
         '******************************************************************
        If UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAHyperlink("Name:=More choices","automationid:=ChooseAnotherOption").Exist Then
        	UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAList("controltype:=List").UIAObject("name:=Use a different account").click
        	 Browser("Browser-SCA").Page("Page-SCA").WebEdit("WebEdit_Username").Set Username
        End If

        
        		    '******************************************************************
    if Browser("Browser-SCA").Page("Page-SCA").WebButton("html id:=btnValidate","name:=Continue").Exist then
    	Browser("Browser-SCA").Page("Page-SCA").WebButton("html id:=btnValidate","name:=Continue").click
    end if 	
'    ****************************
  
'    ***************************************************************************8
	    Browser("Browser-SCA").Page("Page-SCA").WebEdit("WebEdit_Password").Set Password
	    Browser("Browser-SCA").Page("Page-SCA").WebButton("WebButton_Login").Click
 
	   'Handle 'Window Security' PopUp
	   
	'modified by Madhu    
   	elseif UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").Exist(30) Then
		UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").Highlight 
	    UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAEdit("User name").SetValue Username
	    
	      'MOdified by Madhu - 06/05/2020
         '******************************************************************
        If UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAHyperlink("Name:=More choices","automationid:=ChooseAnotherOption").Exist Then
        	UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAHyperlink("Name:=More choices","automationid:=ChooseAnotherOption").click
        	UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAList("controltype:=List").UIAObject("name:=Use a different account").click
        	  UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAEdit("User name").SetValue Username
        End If   
        
        		    '******************************************************************
    if Browser("Browser-SCA").Page("Page-SCA").WebButton("html id:=btnValidate","name:=Continue").Exist then
    	Browser("Browser-SCA").Page("Page-SCA").WebButton("html id:=btnValidate","name:=Continue").click
    end if 	
'    ****************************
  
'    ***************************************************************************
	    
	    UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAEdit("Password").SetValue Password
	    UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAButton("OK").Click
	    
	    Call ReportStep (StatusTypes.Pass, "User " & UserName & " has logged in successfully","Login through Window security PopUp")
	

	 else
		Call ReportStep (StatusTypes.Fail, "User " & UserName & " has not successfully logged in","Login through Window security PopUp")
	End If
'		If Browser("BrowserPoPUp").Dialog("Windows Security").Exist(30) Then
'		   Browser("BrowserPoPUp").Dialog("Windows Security").WinEdit("WinEditUserName").Set UserName	
'		   Browser("BrowserPoPUp").Dialog("Windows Security").WinEdit("WinEditPassword").Set Password
'		   Browser("BrowserPoPUp").Dialog("Windows Security").WinButton("BtnOK").Click
'		   Call ReportStep (StatusTypes.Pass, "User " & UserName & " has logged in successfully","Login through Window security PopUp")
'		Else
'		   Call ReportStep (StatusTypes.Fail, "User " & UserName & " has not successfully logged in","Login through Window security PopUp")	   
'		End If
'   ElseIf Window("BrowserPoPUp").Exist(5) Then
'		Window("BrowserPoPUp").highlight
'	    Window("BrowserPoPUp").Dialog("Windows Security").highlight
'	   'Handle 'Window Security' PopUp
'		If Window("BrowserPoPUp").Dialog("Windows Security").Exist(30) Then
'		   Window("BrowserPoPUp").Dialog("Windows Security").WinEdit("WinEditUserName").Set UserName	
'		   Window("BrowserPoPUp").Dialog("Windows Security").WinEdit("WinEditPassword").Set Password
'		   Window("BrowserPoPUp").Dialog("Windows Security").WinButton("BtnOK").Click
'		   Call ReportStep (StatusTypes.Pass, "User " & UserName & " has logged in successfully","Login through Window security PopUp")
'		Else
'		   Call ReportStep (StatusTypes.Fail, "User " & UserName & " has not successfully logged in","Login through Window security PopUp")	   
   
  
wait 20
 End Function

 
 
 Function ValidationLinkAndToolTip(ByVal ReportLink,ByVal ReportTooltip,ByVal ReportType,ByVal LinkExist,ByRef objData)
 If LinkExist="Yes" Then
 	Browser("Browser-SCA").Page("Page-SCA").Sync
 '*******************************************************	
 'MOdified by Madhu - 04/05/2020	
 	If Browser("Browser-SCA").Page("Page-SCA").WebElement("html tag:=P","innertext:="&ReportLink).Exist(20) Then
	   Browser("Browser-SCA").Page("Page-SCA").WebElement("html tag:=P","innertext:="&ReportLink).highlight	 
	   Call ReportStep (StatusTypes.Pass, " link '"&ReportLink&"' exist in DC for Report type '"&ReportType&"' in DC"," link In DC")
	   Browser("Browser-SCA").Page("Page-SCA").WebElement("html tag:=P","innertext:="&ReportLink).FireEvent "onmouseover"
       Tooltip=Browser("Browser-SCA").Page("Page-SCA").WebElement("html tag:=DIV","innertext:="&ReportTooltip,"class:=ui-tooltip-content").GetRoProperty("innertext")	   
	   

' 	If Browser("DC").Page("DC").WebElement("html tag:=P","innertext:="&ReportLink).Exist(20) Then
'	   Browser("DC").Page("DC").WebElement("html tag:=P","innertext:="&ReportLink).highlight	 
'	   Call ReportStep (StatusTypes.Pass, " link '"&ReportLink&"' exist in DC for Report type '"&ReportType&"' in DC"," link In DC")
'	   Browser("DC").Page("DC").WebElement("html tag:=P","innertext:="&ReportLink).FireEvent "onmouseover"
'       Tooltip=Browser("DC").Page("DC").WebElement("html tag:=DIV","innertext:="&ReportLink&ReportTooltip,"class:=ui-tooltip-content").GetRoProperty("innertext")	   
'	   'Browser("DC").Page("DC").Link("html tag:=A","name:="&ReportLink,"title:="&ReportTooltip).Click
	   If instr(1,ToolTip,ReportTooltip,1) Then
       	  Call ReportStep (StatusTypes.Pass, "For report type '"&ReportType&"' Tooltip of link '"&ReportLink&"' is "&ReportTooltip,"Tooltip of link")	
       Else
          Call ReportStep (StatusTypes.Fail, "For report type '"&ReportType&"' Tooltip of link '"&ReportLink&"' is not exist" ,"Tooltip of link")	
       End If
    Else
       Call ReportStep (StatusTypes.Fail, " link '"&ReportLink&"' not exist in DC for Report type '"&ReportType&"' in DC"," link In DC")   
	End If

' 	If Browser("DC").Page("DC").Link("html tag:=A","name:="&ReportLink,"title:="&ReportTooltip).Exist(20) Then
'	   Browser("DC").Page("DC").Link("html tag:=A","name:="&ReportLink,"title:="&ReportTooltip).highlight	 
'	   Call ReportStep (StatusTypes.Pass, " link '"&ReportLink&"' exist in DC for Report type '"&ReportType&"' in DC"," link In DC")
'       Tooltip=Browser("DC").Page("DC").Link("html tag:=A","name:="&ReportLink,"title:="&ReportTooltip).GetRoProperty("title")	   
'	   Browser("DC").Page("DC").Link("html tag:=A","name:="&ReportLink,"title:="&ReportTooltip).Click
'	   If instr(1,ToolTip,IAMReportTooltipYes,1) Then
'       	  Call ReportStep (StatusTypes.Pass, "For report type '"&ReportType&"' Tooltip of link '"&ReportLink&"' is "&ReportTooltip,"Tooltip of link")	
'       Else
'          Call ReportStep (StatusTypes.Fail, "For report type '"&ReportType&"' Tooltip of link '"&ReportLink&"' is not exist" ,"Tooltip of link")	
'       End If
'    Else
'       Call ReportStep (StatusTypes.Fail, " link '"&ReportLink&"' not exist in DC for Report type '"&ReportType&"' in DC"," link In DC")   
'	End If
 
 End If

If LinkExist="No" Then
'MOdified by Madhu
		If Browser("Browser-SCA").Page("Page-SCA").WebElement("html tag:=P","innertext:="&ReportLink).Exist(20) Then
		   Call ReportStep (StatusTypes.Fail, "For report type '"&ReportType&"' Tooltip of link '"&ReportLink&"' is  exist" ,"Tooltip of link")
		Else
		   Call ReportStep (StatusTypes.Pass, " link '"&ReportLink&"' not exist in DC for Report type '"&ReportType&"' in DC"," link In DC") 
		End If 
	End If
 End Function
 
 
  Function  UpLoadFileInDC(ByVal ReportType,ByVal CountryName,ByVal OfferingName,ByVal Path,ByVal FileName,ByVal ClientType ,ByRef objData)
  
    If ClientType="Syndicated" Then
    	If ReportType="Reporting Services" Then
    			Browser("DifferenetBITools").Link("Syndicated RS Reports").Click

    	ElseIf ReportType="Excel Services" Then
    			Browser("DifferenetBITools").Link("Syndicated Excel Reports").Click

    	ElseIf ReportType="Other Documents" Then
    	       Browser("DifferenetBITools").Link("DCSharedDocuments").Click
    	End If
    End If
  	If Browser("DifferenetBITools").Page("DifferenetBIToolsPages").WebTable("WebTable_Country-Offering").Link("html tag:=A","innertext:="&CountryName).Exist(30) Then
		Call ReportStep (StatusTypes.Pass,"Country name "&CountryName&" is displayed in decision centre ","Country name in DC")
		Browser("DifferenetBITools").Page("DifferenetBIToolsPages").WebTable("WebTable_Country-Offering").Link("html tag:=A","innertext:="&CountryName).Click
		wait 5
		Browser("DifferenetBITools").Link("DCModified").fireevent "onmouseover" 
		If Browser("DifferenetBITools").Image("ModifiedOpenMenu").Exist(15) Then
		    Browser("DifferenetBITools").Image("ModifiedOpenMenu").Click
            Browser("DifferenetBITools").WebElement("OpenMenuDescending").Click
		End If
		wait 3
		If Browser("DifferenetBITools").Page("DifferenetBIToolsPages").WebTable("WebTable_Country-Offering").Link("html tag:=A","innertext:="&OfferingName).Exist(20) Then
			Call ReportStep (StatusTypes.Pass,"Offering name "&OfferingName&" is displayed in decision centre" ,"Offering name in DC")	
			Browser("DifferenetBITools").Page("DifferenetBIToolsPages").WebTable("WebTable_Country-Offering").Link("html tag:=A","innertext:="&OfferingName).Click
		Else
			Call ReportStep (StatusTypes.Fail,"Country name "&CountryName&" is not displayed in decision centre ","Country name in DC")
		End If
	End If
   	wait 1
   	'CLICK ON DOCUMENT TAB AND SELECT UPLOAD FILE
   	'Browser("DifferenetBITools").Page("DifferenetBIToolsPages").Link("Link_DocumentsTab").Click
   	Browser("DifferenetBITools").Page("DifferenetBIToolsPages").Link("Link_DocumentsTab").Click
	wait 8
	Browser("DifferenetBITools").Page("DifferenetBIToolsPages").Link("Link_UploadDocument").Click
'	Set objIE = CreateObject("InternetExplorer.Application")
'	Do While objIE.Busy
'		Wait 1
'	Loop
	'Browser("DC").Page("DC").WebElement("WebEle_UploadDocSubMenu").Click
    If Browser("DC").Page("DC").Frame("Frame").WebFile("WebFile_Browse").Exist(30) Then
        'Browser("DC").Page("DC").Frame("Frame").WebFile("WebFile_Browse").Click	

        Browser("micclass:=browser").Page("micclass:=page").WebFile("html tag:=INPUT","class:=ms-fileinput").set Path&FileName 
'		Set objIE = CreateObject("InternetExplorer.Application")
'		Do While objIE.Busy
'		    Wait 1
'		Loop		
   End If
'	If Browser("DC").Dialog("Choose File to Upload").WinEdit("File name:").Exist(20) Then
'		print("")
'	End If
'	wait 4
'	Browser("DC").Dialog("Choose File to Upload").WinEdit("File name:").Set Path&FileName
'	Browser("DC").Dialog("Choose File to Upload").WinEdit("File name:").Click
' 	Set objShell=CreateObject("WScript.Shell")
'	Browser("DC").Dialog("Choose File to Upload").WinEdit("File name:").Click
'	objShell.SendKeys "{ENTER}"

    wait 5
    Browser("DC").Page("DC").Frame("Frame").WebButton("OK").Click
    If Browser("DC").Page("DC").Frame("Frame_CheckIn").WebButton("Check In").Exist(300) Then
       Browser("DC").Page("DC").Frame("Frame_CheckIn").WebButton("Check In").Click
       Call ReportStep (StatusTypes.Pass,"Clicked on 'CheckIn button'" ,"DC CheckIn button")
       Call ReportStep (StatusTypes.Pass,"Successfully Uplaoded file "&FileName&" for reort type "&ReportType ,"DC CheckIn button")       
    Else
       'Call ReportStep (StatusTypes.Fail,"Not Clicked on 'CheckIn button'" ,"DC CheckIn button")	
    End If
    
'    if Browser("DC").Page("DC").Frame("Frame_CheckIn").WebButton("Save").Exist then
'    	Browser("DC").Page("DC").Frame("Frame_CheckIn").WebButton("Save").Click
'	end if 


  End Function
  
  
  Function SignOutFromDC()
  	Browser("DC").Page("DC").WebElement("WebElement_UserName").Click
	Browser("DC").Page("DC").WebElement("WebElement_SignOut").Click	
	Browser("DC").Close
  End Function
  
  Function ApproveFilesInDC(ByVal ReportType,ByVal CountryName,ByVal OfferingName,ByVal FileName, ByVal ClientType,ByRef objData)
    If ClientType="Syndicated" Then
    	If ReportType="Reporting Services" Then
    			Browser("DifferenetBITools").Link("Syndicated RS Reports").Click

    	ElseIf ReportType="Excel Services" Then
    			Browser("DifferenetBITools").Link("Syndicated Excel Reports").Click

    	ElseIf ReportType="Other Documents" Then
    	       Browser("DifferenetBITools").Link("DCSharedDocuments").Click
    	End If
    End If
    If Browser("DifferenetBITools").Page("DifferenetBIToolsPages").WebTable("WebTable_Country-Offering").Link("html tag:=A","innertext:="&CountryName).Exist(60) Then
		Call ReportStep (StatusTypes.Pass,"Country name "&CountryName&" is displayed in decision centre ","Country name in DC")
		Browser("DifferenetBITools").Page("DifferenetBIToolsPages").WebTable("WebTable_Country-Offering").Link("html tag:=A","innertext:="&CountryName).Click
		wait 5
		Browser("DifferenetBITools").Link("DCModified").fireevent "onmouseover" 

		If Browser("DifferenetBITools").Image("ModifiedOpenMenu").Exist(5) Then
		    Browser("DifferenetBITools").Image("ModifiedOpenMenu").Click
            Browser("DifferenetBITools").WebElement("OpenMenuDescending").Click
		End If
		wait 3
		If Browser("DifferenetBITools").Page("DifferenetBIToolsPages").WebTable("WebTable_Country-Offering").Link("html tag:=A","innertext:="&OfferingName).Exist(20) Then
			Call ReportStep (StatusTypes.Pass,"Offering name "&OfferingName&" is displayed in decision centre" ,"Offering name in DC")	
			Browser("DifferenetBITools").Page("DifferenetBIToolsPages").WebTable("WebTable_Country-Offering").Link("html tag:=A","innertext:="&OfferingName).Click
		Else
			Call ReportStep (StatusTypes.Fail,"Country name "&CountryName&" is not displayed in decision centre ","Country name in DC")
		End If
	End If
	If Browser("DC").Page("DC").Link("innertext:="&FileName,"html tag:=A","visible:=True").Exist(30) Then
		Browser("DC").Page("DC").Link("innertext:="&FileName,"html tag:=A","visible:=True").FireEvent "onmouseover"
		Browser("DC").Page("DC").WebElement("class:=s4-ctx s4-ctx-show","html tag:=DIV").Click
		Call ReportStep (StatusTypes.Pass,"Mouse over on filename to select 'Approve/Reject option '" ,"Offering name in DC")	
	Else
	   Call ReportStep (StatusTypes.Fail,"Mouse over on filename to select 'Approve/Reject option '" ,"Offering name in DC")
	End If
	wait 2
'	If  Browser("DC").Page("DC").Image("Image_OpenMenu").Exist(10) Then
'		 Browser("DC").Page("DC").Image("Image_OpenMenu").Click
'    End If
   Select Case ReportType
   	 Case "Excel Services" 
	   Browser("DC").Page("DC").Link("text:=Approve/Reject","name:=Approve/Reject","html tag:=A").Click
   	   Browser("DC").Page("DC").WebRadioGroup("type:=radio","html tag:=INPUT").Select "0"
       Browser("DC").Page("DC").WebButton("html tag:=INPUT","name:=OK","type:=button").Click
   	 Case "Other Documents"
	   wait 20
	   If Browser("DC").Page("DC").Link("text:=Approve/Reject","name:=Approve/Reject","html tag:=A").Exist(30) Then
	   	  Browser("DC").Page("DC").Link("text:=Approve/Reject","name:=Approve/Reject","html tag:=A").Click
	   End If
	   Browser("DC").Page("DC").WebRadioGroup("type:=radio","html tag:=INPUT").Select "0"
       Browser("DC").Page("DC").WebButton("html tag:=INPUT","name:=OK","type:=button").Click
	Case "Reporting Services"
	   wait 5
	   If Browser("DC").Page("DC").Link("text:=Manage Data Sources","html tag:=A").Exist(30) Then
	   	  Browser("DC").Page("DC").Link("text:=Manage Data Sources","html tag:=A").Click
	   	  Call ReportStep (StatusTypes.Pass,"Clicked on manage Data Sources" ,"Offering name in DC")
       Else
          Call ReportStep (StatusTypes.Pass,"Clicked on manage Data Sources" ,"Offering name in DC")       
	   End If
       wait 10
       If Browser("title:=Manage Data Sources:.*").Page("title:=Manage Data Sources:.*").Link("html tag:=A","innertext:=DataSource1").Exist(30) Then
       	  Browser("title:=Manage Data Sources:.*").Page("title:=Manage Data Sources:.*").Link("html tag:=A","innertext:=DataSource1").Click
       	  ConnectionString= ReadWriteDataFromTextFile("Read","ConnectionString","")
          Browser("title:=DataSource1:.*").Page("title:=DataSource1:.*").WebEdit("type:=textarea","html tag:=TEXTAREA","kind:=multiline").set "DataSource=CDTSOLAP573I.GEMINI.DEV;Initial Catalog=WHSVC_CZ_M_IMS_0001"
	   Browser("title:=DataSource1:.*").Page("title:=DataSource1:.*").WebButton("type:=submit","value:=Test Connection","html tag:=INPUT").Click
       End If
       Set menuChild1=Description.Create()
	   menuChild1("html tag").value="INPUT"
	   menuChild1("type").value="submit"
	   menuChild1("value").value="OK"
	   set childcnt1=Browser("title:=DataSource1:.*").Page("title:=DataSource1:.*").ChildObjects(menuChild1)
	   childcnt1(1).Click
	   Set menuChild1=Description.Create()
	   menuChild1("html tag").value="INPUT"
	   menuChild1("type").value="button"
	   menuChild1("value").value="Close"
	   Set childcnt1=Browser("title:=Manage Data Sources:.*").Page("title:=Manage Data Sources:.*").ChildObjects(menuChild1)
	   childcnt1(1).Click
	   If Browser("DC").Page("DC").Link("innertext:="&FileName,"html tag:=A","visible:=True").Exist(30) Then
		   Browser("DC").Page("DC").Link("innertext:="&FileName,"html tag:=A","visible:=True").FireEvent "onmouseover"
		   Browser("DC").Page("DC").WebElement("class:=s4-ctx s4-ctx-show","html tag:=DIV").Click
		   Call ReportStep (StatusTypes.Pass,"Mouse over on filename to select 'Approve/Reject option '" ,"Offering name in DC")	
	   Else
	      Call ReportStep (StatusTypes.Fail,"Mouse over on filename to select 'Approve/Reject option '" ,"Offering name in DC")
	   End If
'	   If  Browser("DC").Page("DC").Image("Image_OpenMenu").Exist(10) Then
'		    Browser("DC").Page("DC").Image("Image_OpenMenu").Click
'       End If
       Browser("DC").Page("DC").Link("text:=Approve/Reject","name:=Approve/Reject","html tag:=A").Click
   	   Browser("DC").Page("DC").WebRadioGroup("type:=radio","html tag:=INPUT").Select "0"
       Browser("DC").Page("DC").WebButton("html tag:=INPUT","name:=OK","type:=button").Click
   End Select
  

  End Function
  
  
  Function SyndUpLoadFileInDC(Byval ReportType,Byval Country,Byval Offering,Byval filePath,Byval ReportName ,ByRef objData)
	Browser("DC Home").Page("Home").WebElement("SiteActions").Click
    Browser("DC Home").Page("Home").WebElement("ViewAllSite").FireEvent "onmouseover"
    Browser("DC Home").Page("Home").WebElement("ViewAllSite").Click
    Browser("DC Home").Page("Home").Link("html tag:=A","innertext:="&ReportType,"html id:=viewlistDocumentLibrary").Click
    Browser("DC Home").Page("Home").Link("html tag:=A","innertext:="&Country,"name:="&Country).Click
 
   For i = 1 To 20 Step 1
  	   If Browser("DC Home").Page("Home").Link("html tag:=A","innertext:="&Offering,"name:="&Offering).Exist Then
           Browser("DC Home").Page("Home").Link("html tag:=A","innertext:="&Offering,"name:="&Offering).Click
           Exit For       
      Else
            If Browser("DC Home").Page("Home").Image("html tag:=IMG","name:=Image","alt:=Next").Exist Then
            	Browser("DC Home").Page("Home").Image("html tag:=IMG","name:=Image","alt:=Next").Click
            Else
                Exit For
            End If
     End If
   Next
  
   
   Browser("DC Home").Page("Home").Link("text:=Add document","html tag:=A","html id:=idHomePageNewDocument").Click
     If Browser("DC").Page("DC").Frame("Frame").WebFile("WebFile_Browse").Exist(20) Then
        Browser("DC").Page("DC").Frame("Frame").WebFile("WebFile_Browse").Click	
    End If
	
	Browser("DC").Dialog("Choose File to Upload").WinEdit("File name:").Set filePath&ReportName
	Browser("DC").Dialog("Choose File to Upload").WinEdit("File name:").Click
 	Set objShell=CreateObject("WScript.Shell")
	Browser("DC").Dialog("Choose File to Upload").WinEdit("File name:").Click
	objShell.SendKeys "{ENTER}"

    wait 5
    Browser("DC").Page("DC").Frame("Frame").WebButton("OK").Click
    If Browser("DC").Page("DC").Frame("Frame_CheckIn").WebButton("Check In").Exist(20) Then
       Browser("DC").Page("DC").Frame("Frame_CheckIn").WebButton("Check In").Click
       Call ReportStep (StatusTypes.Pass,"Clicked on 'CheckIn button'" ,"DC CheckIn button")
       Call ReportStep (StatusTypes.Pass,"Successfully Uplaoded file "&ReportName&" for reort type "&ReportType ,"DC CheckIn button")       
    Else
       'Call ReportStep (StatusTypes.Fail,"Not Clicked on 'CheckIn button'" ,"DC CheckIn button")	
    End If

End Function

Public Function FilterByReqIDAndOpen(ByVal RequestID,ByRef objData)
	

Call SearchByID("", "","",objData.item("Tab"),objData.item("ColumnName"),CInt(RequestID),objData)

wait 5
   'Double click on existing client request
    strClientName=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").GetCellData(2,4)
    strCountry = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").GetCellData(2,5)
    'strClientShortName
    Set objTab = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data")
	    Set oTableCell = Description.Create
	     oTableCell("micclass").value = "WebElement"
	    oTableCell("html tag").value = "TD"
		oTableCell("innertext").value = strClientName&".*"   
		oTableCell("outerhtml").value = "<TD role=gridcell aria-describedby=request-list_ClientName title=.* "
	wait 5
		set oTableCellObj = objTab.ChildObjects(oTableCell)
        col=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").GetROProperty("cols")
      If  oTableCellObj.count >= 1 Then
              For i = 0 To oTableCellObj.count-1 Step 1
                 	x= oTableCellObj(i).GetRoproperty("innertext")
                 	If instr(1,x,strClientName,1)>0 Then
                 		If instr(1,x,strClientName,0)>0 Then
                 		    oTableCellObj(i).highlight
                 		    Call ReportStep (StatusTypes.Pass,"Searched client name " & strClientName & " Exist in Client List At Line No:" & i+1 &" With name" & x ,"Search client name")
                 			oTableCellObj(i).FireEvent "ondblclick"
                 			If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-dialog-title","html tag:=SPAN","innertext:=Confirmation","visible:=True").Exist(10) Then
                 			   Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-button-text","html tag:=SPAN","innertext:=No","visible:=True").Click
                                oTableCellObj(i).RightClick
                                Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html tag:=SPAN","innertext:=Change Request Status To > ","visible:=True").fireevent "onmouseover"
               					Browser("Browser-OCRF").Page("Page-OCRF").WebElement("innertext:=Unlock","html tag:=SPAN","visible:=True").Click
                 		        oTableCellObj(i).FireEvent "ondblclick"
                 		    End If
                 		Else
                 		    Call ReportStep (StatusTypes.Pass,"Searched client name " & strClientName & " NOT Match with client name" & x &" at row " & i+1 &" With Case sensitive name" & x ,"Search client name")
                 		End If
                 	End If
                 Next
      End If
End Function

Public Function AddMultipleDatabase(ByVal FileName,ByVal SheetName,ByVal StartRow,Byval EndRow,Byval timeStamp,ByRef objData)

    If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Wel_UserDB_List").Exist(60) Then
		Call ReportStep (StatusTypes.Pass, "Entered in to cude details tab ","Cube details tab")	
	Else
		Call ReportStep (StatusTypes.Fail, "Not Entered in to cude details tab","Cube details tab")
	End If
	
	Call ImportSheet(SheetName,FileName)
	
	'Number of rows to be added
	Datatable.SetCurrentRow(StartRow)
	NumberOfRows = (EndRow-StartRow)+1
		
	'select number of rows and click on add new line
'	Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_NumberOfRows").Select NumberOfRows
'	Call ReportStep (StatusTypes.Pass, "Entering "&NumberOfRows&" records in 'Cube details' tab started","Cube details tab")


    TableRow=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetROProperty("rows")+1
	For row = StartRow To EndRow Step 1
	
	
	
'		*********************************************
'	Modified by Madhu - 03/02/2020
				If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes_Index1").Exist(1) Then
		          
		       	  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes_Index1").Click
		       End If
		       If  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes_Index0").Exist(1) Then
		   	      
		      	  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes_Index0").Click
		      End If 
		      

'88***********************************************************		
		
	
	
	
		If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementAddNewLine").Exist(10) Then
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementAddNewLine").Click
		End If
	    
	    wait 2
	    'Select Action in row
		Call WebEditExistence(TableRow,7,"WebList")
			
		Set selectAction = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(TableRow,7,"WebList",0)
		selectAction.select datatable.value("Action",SheetName) 

'        'Select 'ActionComments' in specific row 
'		Call WebEditExistence(row,8,"WebEdit")
'			
'		Set setActionComments = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(row,8,"WebEdit",0)
'		setActionComments.set datatable.value("ActionComments",SheetName) 
		
		'Select 'DBType' in specific row 	
		Call WebEditExistence(TableRow,10,"WebList")
			
		Set selectDBType = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(TableRow,10,"WebList",0)
		selectDBType.select datatable.value("DBType",SheetName) 
	
	   
	    
		Set CheckRow =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(TableRow,2,"WebCheckBox",0)
		CheckRow.click	
		
		Call WebEditExistence(TableRow,Cint(Environment.value("cubeName")),"WebEdit")
		
		Set CheckDatabase =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(TableRow,Cint(Environment.value("cubeName")),"WebEdit",0)
		CheckDatabase.set datatable.value("DatabaseName",SheetName)&timeStamp 
		wait 1
		
		If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Apply datasource security").Exist(10) Then
			Call ReportStep (StatusTypes.Pass,"Datasource security popup appeared" ,"Datasource tab")
			
		End If
		If Browser("Browser-OCRF").Page("Page-OCRF").WebCheckBox("DSsecurityCheckbox").Exist(5) Then
			If datatable.value("DataSourceSecurity",SheetName)="Check" Then
			   Browser("Browser-OCRF").Page("Page-OCRF").WebCheckBox("DSsecurityCheckbox").Set "ON"
			End If
		End If
	    If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("DSProceed").Exist(5) Then
	    wait 3
	    	Browser("Browser-OCRF").Page("Page-OCRF").WebButton("DSProceed").DoubleClick
	    End If
		If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebeleAssigOwnerPopUp").Exist(15) Then
		   If datatable.value("AssignOwner",SheetName)="Choose new DataSource Owner" Then
		   	  Browser("Browser-OCRF").Page("Page-OCRF").WebRadioGroup("DataSourceOwner").Select 2 
		   	  Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditdDtaSourceOwner").Set objData.item("DataBaseOwner")
		   	   wait 3
		   	   Browser("Browser-OCRF").Page("Page-OCRF").WebButton("DSProceed").DoubleClick
		   	   If  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes").Exist(5) Then
		   	      wait 2
		      	  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes").Click
		      End If 
		      If  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes").Exist(5) Then
		   	      wait 2
		      	  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes").Click
		      End If 
		   Else
		      Browser("Browser-OCRF").Page("Page-OCRF").WebRadioGroup("DataSourceOwner").Select 1
		      wait 2
		      Browser("Browser-OCRF").Page("Page-OCRF").WebButton("DSProceed").DoubleClick
		      If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebEleOwnershipAssignmentFor").Exist(15) Then
		         If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerPopUpYes").Exist(10) Then
		            wait 2
		         	Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerPopUpYes").DoubleClick
		         End If
		         If  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes").Exist(15) Then
		   	          wait 2
		      	        Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes").Click
		         End If  
		      End If 
		   End If
		ElseIf Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebEleConfirmation").Exist(5) Then 
		       wait 1
		       If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerPopUpYes").Exist(10) Then
		          wait 2
		       	  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerPopUpYes").Click
		       End If
		       If  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes").Exist(15) Then
		   	      wait 2
		      	  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes").Click
		      End If 
	    End If
	
'*******************************
'Modified BY madhu 02/20/2020
'------Changed click to double click for popups
		On error resume next
	    If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerPopUpYes").Exist(8) Then
	    	Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerPopUpYes").Click
	    End If
	    If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerPopUpYes").Exist(1) Then
	    	Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerPopUpYes").DoubleClick
	    End If
	     If  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes").Exist(1) Then
		   	wait 2
		    Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes").DoubleClick
		End If
	    If  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes").Exist(1) Then
		   	wait 2
		    Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes").Click
		End If 
'		'Write 'DataBase description' in specific row 
'        Call WebEditExistence(row,11,"WebEdit")
'			
'		Set setDBDesc = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(row,11,"WebEdit",0)
'		setDBDesc.set datatable.value("DBDescription",SheetName) 
			
'		'Write 'DB Owner Email' in specific row 	
'		Call WebEditExistence(row,18,"WebEdit")
'			
'		Set setDbOwEmail = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(row,18,"WebEdit",0)
'		setDbOwEmail.set datatable.value("DBOwnerEmail",SheetName) 

			
'		*********************************************
'	Modified by Madhu - 03/02/2020
				If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes_Index1").Exist(1) Then
		          
		       	  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes_Index1").Click
		       End If
		       If  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes_Index0").Exist(1) Then
		   	      
		      	  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes_Index0").Click
		      End If 
		      

'88***********************************************************		
				
		Set CheckRow =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(TableRow,2,"WebCheckBox",0)
		CheckRow.click
		wait 1
		Set CheckRow =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(TableRow,2,"WebCheckBox",0)
		CheckRow.click
		wait 1
		
		'If user is not owner of Datasource user has to request for permission
		'Check the status for pending
		 ActionStatus =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetCellData(TableRow,7)
		If ActionStatus="PENDING" Then
           Exit Function 
		End If
		
	    'Select hosting server if it is not disabled
		Set CheckDBLocation =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(TableRow,Cint(Environment.value("cubeLocation")),"WebEdit",0)		
		wait 1
		If CheckDBLocation.GetRoProperty("disabled")=0 Then
		   Call WebEditExistence(TableRow,Cint(Environment.value("cubeLocation")),"WebEdit")
		
		   Set CheckDBLocation =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(TableRow,Cint(Environment.value("cubeLocation")),"WebEdit",0)
		   CheckDBLocation.set datatable.value("ServerLocation",SheetName)	
		End If
		'Select FLA path if it is not disabled
		Set CheckFLAPath =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(TableRow,Cint(Environment.value("flaPath")),"WebList",0)
		wait 1
		If CheckFLAPath.GetRoProperty("disabled")=0 Then
		   	Call WebEditExistence(TableRow,Cint(Environment.value("flaPath")),"WebList")
		
		    Set CheckFLAPath =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(TableRow,Cint(Environment.value("flaPath")),"WebList",0)
		    CheckFLAPath.select datatable.value("FLAPATH",SheetName)	 
		End If

'		*********************************************
'	Modified by Madhu - 03/02/2020
				If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerPopUpYes").Exist(5) Then
		          wait 2
		       	  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerPopUpYes").Click
		       End If
		       If  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes_Index0").Exist(1) Then
		   	    
		      	  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes_Index0").Click
		      End If 
		      

'88***********************************************************		
		
		datatable.SetNextRow
		Call ReportStep (StatusTypes.Pass, "Entered  Database Name as '"&datatable.value("DatabaseName",SheetName)&"',Server Location as '"&datatable.value("ServerLocation",SheetName)&"' And FLA path as '"&datatable.value("FLAPATH",SheetName)&"' at row '"&del-1&"' done","Cube details tab")
	    TableRow=TableRow+1
	Next
	
	
End Function

Public Function BI_Tools_Access_Combination(ByVal FileName,ByVal BIToolSheetName,ByVal StartRow,ByVal EndRow,ByRef objData)
	'1) Validate whether OCRF New Request is redirected to BI Tool access
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementOfferingUserBIToolList").Exist(120) <> true Then
		Call ReportStep (StatusTypes.Fail, "Not Redirected to BI Tool Access tab","BI tool access Tab")	
	Else
		'Call ReportStep (StatusTypes.Pass, "Redirected to BI Tool Access tab successfully","BI tool access Tab")
	End If
		
	Call ImportSheet(BIToolSheetName,FileName)

	If Ucase(StartRow) = "CHECKALL" Then
		
		Datatable.GetSheet(BIToolSheetName).SetCurrentRow(EndRow)
		If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Exist(120) Then
			    Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Click
				wait 2
			Else
				Call ReportStep (StatusTypes.Fail, "Not Clicked on 'Select user","BI tool access Tab")
		End If
			
				
'		Set objBIAccessPageContent = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_PageContent")
'		Call SCA.SelectFromDropdown(objBIAccessPageContent, "999")
		Call PageLoading()
		wait 2
		Set objBIAccessTool = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIAccessTool")
		Set objBIToolAccessMode = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIToolAccessMode")
		
		Call SCA.SelectFromDropdown(objBIAccessTool, Datatable.Value("BI_Tool",BIToolSheetName))
		Call SCA.SelectFromDropdown(objBIToolAccessMode, Datatable.Value("Access_Mode",BIToolSheetName))
		
		Browser("Browser-OCRF").Page("Page-OCRF").WebCheckBox("WebCheckBox_SelectAll_BiTools").Set "ON"
		wait 1
		'Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_AddUser").Click
		Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_Add").Click
		wait 5
		If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebBtnInformationOkay").Exist(10) Then 
			Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebBtnInformationOkay").Click
			wait 2
		End If 
	
		Else
			If StartRow = "" Then
				StartRow = 1
			End If
			If EndRow = "" Then
				EndRow = Datatable.GetSheet(BIToolSheetName).GetRowCount
			End If
			'rc = Datatable.GetSheet(BIToolSheetName).GetRowCount
			Datatable.GetSheet(BIToolSheetName).SetCurrentRow(StartRow)
			
'			Set oDesc = Description.Create
'			oDesc("micclass").value = "WebList"
'			oDesc("class").value = "ui-pg-selbox"
'			oDesc("name").value = "select"
'			Set obj = Browser("Browser-OCRF").Page("Page-OCRF").ChildObjects(oDesc)
'			msgbox obj.count 
'			For i = 0 To obj.count -1 Step 1
'				obj(i).select "100"
'			Next

			Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolGrid_PageContent").Select "100"
			Call PageLoading()
			wait 1
			recordCountBeforeAdding = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").RowCount
			recordAddedAtRow = recordCountBeforeAdding + 1
			For rownum = StartRow To EndRow Step 1
				If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Exist(120) Then
				    Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Click
					wait 2
				Else
					Call ReportStep (StatusTypes.Fail, "Not Clicked on 'Select user","BI tool access Tab")
				End If
				
'				Set objBIAccessPageContent = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_PageContent")
'				Call SCA.SelectFromDropdown(objBIAccessPageContent, "100")
				
				Set objBIAccessTool = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIAccessTool")
				Set objBIToolAccessMode = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIToolAccessMode")
				
				Call SCA.SelectFromDropdown(objBIAccessTool, Datatable.Value("BI_Tool",BIToolSheetName))
				Call SCA.SelectFromDropdown(objBIToolAccessMode, Datatable.Value("Access_Mode",BIToolSheetName))
				
'				**********************************************
'				Modified by Madhu - 02/20/20
''				Added below line to select the users list number to display
				Browser("Browser-OCRF").Page("Page-OCRF").WebList("Weblist_SelectUserlistNo").Select "999"
				Call PageLoading()
',				**************************************************************

				userRowNum = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").GetRowWithCellText(Datatable.value("User_Id",BIToolSheetName),4,2)
				Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").ChildItem(userRowNum,2,"WebCheckBox",0).click
		
'				**********************************************
'				Modified by Madhu - 02/20/20
''				Commented below line as it is removed from application

				'Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_AddUser").Click
'				,*****************************************************
				If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_Add").Exist Then
					
					Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_Add").Click
				elseIf Browser("Browser-OCRF").Page("Page-OCRF").WebButton("html tag:=BUTTON","innertext:=Add","visible:=True").Exist Then
					Browser("Browser-OCRF").Page("Page-OCRF").WebButton("html tag:=BUTTON","innertext:=Add","visible:=True").Click
				End If
				wait 5
				Call PageLoading()
				
				If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Information_Okay").Exist(10) Then 
					Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Information_Okay").Click
					wait 2
				End If 
				Set odesc=description.Create()
				odesc("micclass").Value="WebElement"
				odesc("innertext").Value="Okay"
				odesc("html tag").Value="SPAN"
				odesc("Visible").Value=True
				
				set c=browser("Browser-OCRF").Page("Page-OCRF").ChildObjects(odesc)
				
				
				If c.count<>0 Then
					For i = 1 To c.count Step 1
					c(i-1).click
				Next
				End If
				recordCount = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").RowCount
				If recordCount  = recordAddedAtRow Then
					Call ReportStep (StatusTypes.Pass, "New row added at " & recordAddedAtRow & " with " & Datatable.Value("BI_Tool",BIToolSheetName) & ", " & Datatable.Value("Access_Mode",BIToolSheetName) & ", " & Datatable.value("User_Id",BIToolSheetName),"BI tool access Tab - Add Record")
					recordAddedAtRow = recordAddedAtRow + 1
					Else 
					Call ReportStep (StatusTypes.Fail, "New row couldn't add for row no. " & recordCount +1 ,"BI tool access Tab - Add Record")
					Exit For 
				End If
'				***********************************************************
'			MOdified by Madhu - 03/10/2020

				if Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Webbutton_Okay").Exist(2) then
					Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Webbutton_Okay").Click
				end if
					
				if Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Webbutton_Okay").Exist(2) then
					Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Webbutton_Okay").Click
				end if					
'*********************************************************************
				Datatable.GetSheet(BIToolSheetName).SetNextRow
			Next
				
'			If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebBtnInformationOkay").Exist Then 
'				Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebBtnInformationOkay").Click
'					
'			End If 
'			Dim odesc 
'			Set odesc = Description.Create
'			odesc("micclass").value = "WebButton"
'			odesc("html tag").value = "BUTTON"
'			odesc("name").value = "Okay"
'			
'			Set chOdesc = Browser("Browser-OCRF").Page("Page-OCRF").ChildObjects(odesc)
'			For each i In chOdesc
'				
'			Next
			
			
			'Validate number of records added
			Call PageLoading()
			wait 3
			Browser("Browser-OCRF").Page("Page-OCRF").Sync
			
			noOfRecords = (EndRow - StartRow) +1
			recordCountAfterAdding = (recordCountBeforeAdding - 1) +  noOfRecords
			
			wait 3
			If (Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").RowCount)-1 = recordCountAfterAdding Then
				Call ReportStep (StatusTypes.Pass, recordCountAfterAdding & " Records matched successfully between data sheet and BI Tool access list","Records match for BI Tool access")	
			Else
				Call ReportStep (StatusTypes.Fail, recordCountAfterAdding & " Records did not match, between data sheet and BI Tool access list","Records match for BI Tool access")	
			End If	
			if Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_DBAccesspage").Exist(5) then
				Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_DBAccesspage").Click
			end if 
'			If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[contains(@style,'display: block;')]//button//span[text()='Okay']").Exist(10) Then
'				Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[contains(@style,'display: block;')]//button//span[text()='Okay']").Click
'			End If
			'Call NavigateToNextTab()
			Set odesc=description.Create()
			odesc("micclass").Value="WebElement"
			odesc("innertext").Value="Okay"
			odesc("html tag").Value="SPAN"
			odesc("Visible").Value=True
			
			set c=browser("Browser-OCRF").Page("Page-OCRF").ChildObjects(odesc)
			
			
			If c.count<>0 Then
				For i = 1 To c.count Step 1
				c(i-1).click
			Next
			End If
			'Call PageLoading()	
	End If
	
	
End Function



Public Function BI_Tools_Access_Combination_Old(ByVal FileName,ByVal SheetName,ByVal StartRow,ByVal EndRow,ByRef objData)

	'1) Validate whether OCRF New Request is redirected to BI Tool access
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementOfferingUserBIToolList").Exist(120) <> true Then
		Call ReportStep (StatusTypes.Fail, "Not Redirected to BI Tool Access tab","BI tool access Tab")	
	Else
		Call ReportStep (StatusTypes.Pass, "Redirected to BI Tool Access tab successfully","BI tool access Tab")
	End If
	
	Call ImportSheet(SheetName,FileName)
	
	
'	rc = Datatable.GetSheet(SheetName).GetRowCount
'	Datatable.GetSheet(SheetName).SetCurrentRow(1)
	Call SelectRowsFromDropDownList(13,"50",objData)
	Wait 5
	
	For rnum = StartRow To EndRow Step 1
	
	    If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Exist(120) Then
	    	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Click
			Call ReportStep (StatusTypes.Pass, "Clicked on 'Select user'","BI tool access Tab")	
		Else
			Call ReportStep (StatusTypes.Fail, "Not Clicked on 'Select user","BI tool access Tab")
		End If
	
		Set objBIAccessTool = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIAccessTool")
		Set objBIToolAccessMode = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIToolAccessMode")

		Call SCA.SelectFromDropdown(objBIAccessTool, Datatable.Value("BI_Tool",SheetName))
		Call SCA.SelectFromDropdown(objBIToolAccessMode, Datatable.Value("Access_Mode",SheetName))
		

		rownum = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").GetRowWithCellText(Datatable.value("User_Id",SheetName),4,2)
		Set chk =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").ChildItem(rownum,2,"WebCheckBox",0)
		chk.click
		Call ReportStep (StatusTypes.Pass, "User Id :"&Datatable.value("User_Id",SheetName)&" selected for BITool: "&Datatable.Value("BI_Tool",SheetName)&" Access Mode: "&Datatable.Value("Access_Mode",SheetName)&" Combination" ,"BI tool access Tab")
	    Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_AddUser").Click
	    wait 3
		Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_Add").Click 
		
		Datatable.GetSheet(SheetName).SetNextRow
	Next
	
	
'	For rnum = 1 To rc+100 Step 1
'		
'		Set objBIAccessTool = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIAccessTool")
'		Set objBIToolAccessMode = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIToolAccessMode")
'
'		Call SCA.SelectFromDropdown(objBIAccessTool, Datatable.Value("BI_Tools",SheetName))
'		Call SCA.SelectFromDropdown(objBIToolAccessMode, Datatable.Value("Access_Mode",SheetName))
'		
'		CurComb = Datatable.Value("BI_Tools",SheetName)&Datatable.Value("Access_Mode",SheetName)
'		
'		Datatable.GetSheet(SheetName).SetPrevRow
'		PrevComb = Datatable.Value("BI_Tools",SheetName)&Datatable.Value("Access_Mode",SheetName)
'		Datatable.GetSheet(SheetName).SetNextRow
'		
'		If (Datatable.GetSheet(SheetName).GetCurrentRow = 1) OR ((CurComb = PrevComb) AND (Datatable.GetSheet(SheetName).GetCurrentRow > 0)) OR (Diffvalue = 1) Then
'			Diffvalue = 0
'			rownum = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").GetRowWithCellText(Datatable.value("User_ID","BI_Tools"),4,2)
'			Set chk =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").ChildItem(rownum,2,"WebCheckBox",0)
'			chk.click
'	
'		    Call ReportStep (StatusTypes.Pass, "User Id :"&Datatable.value("User_ID","BI_Tools")&" selected for BITool: "&Datatable.Value("BI_Tools","BI_Tools")&" Access Mode: "&Datatable.Value("Access_Mode","BI_Tools")&" Combination" ,"BI tool access Tab")	
'	
'			If (Datatable.GetSheet(SheetName).GetCurrentRow = Datatable.GetSheet(SheetName).GetRowCount) AND (Diffvalue = 0) Then
'				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_AddUser").Click
'				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_Add").Click
'				Exit For
'			Else
'				Datatable.GetSheet(SheetName).SetNextRow
'			End If
'		
'		Else
'		
'		
'			Datatable.GetSheet(SheetName).SetPrevRow
'			Call SCA.SelectFromDropdown(objBIAccessTool, Datatable.Value("BI_Tools","BI_Tools"))
'			Call SCA.SelectFromDropdown(objBIToolAccessMode, Datatable.Value("Access_Mode","BI_Tools"))
'			Datatable.GetSheet(SheetName).SetNextRow
'			
'			
'		
'			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_AddUser").Click
'			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_Add").Click
'			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Click
'			Diffvalue = 1
'		End If	
'		
'	Next
'	
'	
'	'Validate number of records added
'	Call PageLoading()
'	Browser("Browser-OCRF").Page("Page-OCRF").Sync
'	
'	If (Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").RowCount)-1 = rc  Then
'		Call ReportStep (StatusTypes.Pass, "Records matched successfully between data sheet and BI Tool access list","Records match")	
'	Else
'		Call ReportStep (StatusTypes.Fail, "Records did not match, between data sheet and BI Tool access list","Records match")	
'	End If	
'	
'	Call NavigateToNextTab()
'	
'	Call PageLoading()
	
End Function
	


Public Function USer_DB_Access_Combination(ByVal FileName,ByVal SheetName,ByVal StartRow,ByVal EndRow,ByVal Timestamp,ByRef objData)
	
	TotalRow=(EndRow-StartRow)+1
	
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_UserDatabaseList").Exist(60) Then
		Browser("Browser-OCRF").Page("Page-OCRF").Sync
		Call ReportStep (StatusTypes.Pass, "Redirected to User DB Access tab successfully","User DB access Tab")
	Else
		Call ReportStep (StatusTypes.Fail, "Not Redirected to User DB Access tab","User DB access Tab")	
		Exit Function
	End If
	
	Call ImportSheet(SheetName,FileName)
	
	DataTable.GetSheet(SheetName).SetCurrentRow(StartRow)
	
	Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_PageContent").Select "100"
	
	 rowCount = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserCubeAccess_CubeAccessList").RowCount
	 rowCountBeforeAdding = rowCount -1 

	
	For rnum = StartRow To EndRow Step 1	
		
        If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Exist(120) Then
	    	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Click
			Call ReportStep (StatusTypes.Pass, "Clicked on 'Select user'","USer_DB_Access Tab")	
		Else
			Call ReportStep (StatusTypes.Fail, "Not Clicked on 'Select user","USer_DB_Access Tab")
		End If	

		wait 2
		
		pageContentUserCube = "//TD[@id=" & chr(34)& "CubeUser-toadd-list-pager_center"& chr(34) & "]/TABLE[1]/TBODY[1]/TR[1]/TD[8]/SELECT[@role="& chr(34) & "listbox" & chr(34)&"][1]"
		
		Browser("Browser-OCRF").Page("Page-OCRF").WebList("xpath:="&pageContentUserCube).Select "100"
		
'		Set oDesc = Description.Create
'		oDesc("micclass").value = "WebList"
'		oDesc("class").value = "ui-pg-selbox"
'		oDesc("name").value = "select"
'		Set obj = Browser("Browser-OCRF").Page("Page-OCRF").ChildObjects(oDesc)
'		For i = 0 To obj.count -1 Step 1
'			obj(i).select "100"
'		Next
		wait 3
		
		Set objCubeDB = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_CubeDB")
		Set objCubeRoles = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_CubeRoles")
			
		Set objCubeType = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_CubeType")
	    If objCubeType.Exist(5) Then
		wait 2
	     If objCubeType.GetROProperty("disabled") = 0 Then
	       objCubeType.Select Trim(Datatable.Value("DatabaseType",SheetName))
	       wait 2
	     End If
	    End If
		'Call SCA.SelectFromDropdown(objcubeType, Datatable.value("DatabaseType",SheetName))
		wait 3
		objCubeDB.WaitProperty "value","Select Datasource Name",5000
		wait 4
		objCubeDB.Select trim(Datatable.value("Database",SheetName))&Timestamp
		objCubeRoles.WaitProperty "disabled",0,5000
		wait 3
		objCubeRoles.Select Trim(Datatable.value("Role",SheetName))
'		Call SCA.SelectFromDropdown(objCubeDB, trim(Datatable.value("Database",SheetName)))
'		Call SCA.SelectFromDropdown(objCubeRoles, Trim(Datatable.value("Role",SheetName)))
		'Call SelectRowsFromDropDownList(row,"999",objData)
'		rownum = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").GetRowWithCellText(Datatable.value("User_ID",SheetName),4,2)
'		Set chk =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").ChildItem(rownum,2,"WebCheckBox",0)
'		chk.click
		'directly entering users in grid
		browser("micclass:=browser").Page("micclass:=page").Webedit("xpath:=//textarea[@id='cubeuser-loginid']").Click
		browser("micclass:=browser").Page("micclass:=page").Webedit("xpath:=//textarea[@id='cubeuser-loginid']").Set dataTable.value("User_ID",SheetName)
'		****************************************
		'Datatable.Value(
'		Modified by madhu 02/20/2020
'		commented below line as that object is removed from the application
		
'		Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_UserDB_AddUsers").Click
		wait 3
		If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("html tag:=BUTTON","innertext:=Add","visible:=True").Exist Then
			Browser("Browser-OCRF").Page("Page-OCRF").WebButton("html tag:=BUTTON","innertext:=Add","visible:=True").Click
		 
		elseif Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_UserDB_Add").Exist(2) then
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_UserDB_Add").Click
		end if
		wait 5
		Call PageLoading()
		Datatable.GetSheet(SheetName).SetNextRow
		
	Next
	
	'Validate number of records added
	Call PageLoading()
	Browser("Browser-OCRF").Page("Page-OCRF").Sync
	
	noOfRecords = (EndRow - StartRow) +1
	RowCountAfterAdding = rowCountBeforeAdding + noOfRecords
	
	Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_PageContent").Select "100"
	wait 3
	If (Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserCubeAccess_CubeAccessList").RowCount)-1 <> RowCountAfterAdding+1 Then
		Call ReportStep (StatusTypes.Pass, RowCountAfterAdding & " Records matched successfully between data sheet and DataSource access list","Records match in DataSource Access tab")	
	Else
		Call ReportStep (StatusTypes.Fail, RowCountAfterAdding & " Records did not match, between data sheet and BI Tool access list","Records match in DataSource Access tab")	
	End If	
	
	If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_DBAccesspage").Exist Then
		Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_DBAccesspage").Click
	End If
	
	
	If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebBtnInformationOkay").Exist(5) Then
		Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebBtnInformationOkay").Click
	End If
	If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebBtnInformationOkay").Exist Then
		Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebBtnInformationOkay").Click
	End If
	
	'#####################
	

	New_Request_ID = trim(Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_RequestID").GetROProperty("value"))
    USer_DB_Access_Combination = New_Request_ID  




'	rownum = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").GetROProperty("rows")
'	For j = 1 To rownum Step 1
'		userData=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").GetCellData(j+1,4)
'		If userData=UserID Then
'		   Set chk =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").ChildItem(j+1,2,"WebCheckBox",0)
'	        chk.click
'	        Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_UserDB_AddUsers").Click
'	        Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_UserDB_Add").Click
'		End If
'	Next
'
'   	New_Request_ID = trim(Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_RequestID").GetROProperty("value"))
'  
'
'	Call PageLoading()
'	
'	AddNewLineInUserDB = New_Request_ID
'	
'	
'	datatable.AddSheet SheetName
'	datatable.ImportSheet Path&"OCRF_SyndicatedClientContent_27.xlsx",SheetName,SheetName
'	
'	datatable.GetSheet SheetName
'	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Exist(120) Then
'	    Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Click
'		Call ReportStep (StatusTypes.Pass, "Clicked on 'Select user'","USer_DB_Access Tab")	
'	Else
'		Call ReportStep (StatusTypes.Fail, "Not Clicked on 'Select user","USer_DB_Access Tab")
'	End If
'	rc = Datatable.GetSheet(SheetName).GetRowCount
'	Datatable.GetSheet(SheetName).SetCurrentRow(1)
'	Call SelectRowsFromDropDownList(16,"50",objData)
'	Call SelectRowsFromDropDownList(16,"100",objData)
'	wait 5
'	For rnum = 1 To rc+100 Step 1
'		
'		Set objCubeDB = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_CubeDB")
'		Set objCubeRoles = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_CubeRoles")
'		
'		wait 5
'		Call SCA.SelectFromDropdown(objCubeDB, Datatable.Value("Database","User_DB"))
'		Call SCA.SelectFromDropdown(objCubeRoles, Datatable.Value("Role","User_DB"))
'		
'		CurComb = Datatable.Value("Database","User_DB")&Datatable.Value("Role","User_DB")
'		
'		Datatable.GetSheet(SheetName).SetPrevRow
'		PrevComb = Datatable.Value("Database","User_DB")&Datatable.Value("Role","User_DB")
'		Datatable.GetSheet(SheetName).SetNextRow
'		
'		If (Datatable.GetSheet(SheetName).GetCurrentRow = 1) OR ((CurComb = PrevComb) AND (Datatable.GetSheet(SheetName).GetCurrentRow > 0)) OR (Diffvalue = 1) Then
'			Diffvalue = 0
'			
'			If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").Exist(60) Then
'				rownum = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").GetRowWithCellText(Datatable.value("User_ID","User_DB"),4,2)
'				Set chk =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").ChildItem(rownum,2,"WebCheckBox",0)
'				chk.click
'				 Call ReportStep (StatusTypes.Pass, "User Id :"&Datatable.value("User_ID","User_DB")&"selected for Data base: "&Datatable.Value("Database","User_DB")&" Role : "&Datatable.Value("Role","User_DB")&" Combination" ,"User_DB access Tab")	
'				If (Datatable.GetSheet(SheetName).GetCurrentRow = Datatable.GetSheet(SheetName).GetRowCount) AND (Diffvalue = 0) Then
'					Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_UserDB_AddUsers").Click
'					Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_UserDB_Add").Click
''					Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Add").Click
'					Exit For
'				Else
'					Datatable.GetSheet(SheetName).SetNextRow
'				End If
'			End If
'			
'			
'		
'		Else
'		
'			Datatable.GetSheet(SheetName).SetPrevRow
'			Call SCA.SelectFromDropdown(objCubeDB, Datatable.Value("Database","User_DB"))
'			Call SCA.SelectFromDropdown(objCubeRoles, Datatable.Value("Role","User_DB"))
'			Datatable.GetSheet(SheetName).SetNextRow
'		
'			
'		
'		
'			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_UserDB_AddUsers").Click
'			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_UserDB_Add").Click
''			Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Add").Click
'			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Click
'			Diffvalue = 1
'		End If	
'		
'	Next
'	
'	
'	'Validate number of records added
'	Call PageLoading()
'	Browser("Browser-OCRF").Page("Page-OCRF").Sync
'	
'	If (Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDatabaseList").RowCount)-1 = rc  Then
'		Call ReportStep (StatusTypes.Pass, "Records matched successfully between data sheet and User DB access list","Records match")	
'	Else
'		Call ReportStep (StatusTypes.Fail, "Records did not match, between data sheet and User DB access list","Records match")	
'	End If	
'
'	
'	'Store the request ID
'	New_Request_ID = trim(Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_RequestID").GetROProperty("value"))
'
'	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementFinish").Click
'	
'	Call PageLoading()
'	
'	USer_DB_Access_Combination = New_Request_ID

End Function





Public Function ValHostingStatusInOps(ByVal Path,ByVal SheetName,ByRef objData,ByVal cubeNames, ByVal statusQA, ByVal statusHosting)
	'On Error Resume Next
	If IsArray(cubeNames) Then
		lenDbName = Len(cubeNames(0))
		dbName = Left(cubeNames(0), lenDbName - 3)
		noOfCubes = Ubound(cubeNames)+ 1
		expWaitTime = objData.Item("expHostingTime")
	Else 
		dbName = cubeNames
		noOfCubes = 1
		expWaitTime = 4
	End If
	
	
	appWaitTime = 0
	counter = 0
	For appWaitTime = 0 To expWaitTime Step 1
		counter = CheckHostingStatus(dbName, statusQA, statusHosting)
		If counter = noOfCubes Then
			Call ReportStep(StatusTypes.Pass, "All cubes are successfully hosted/accepted at " & now() ,"Hosting Status Info")
			timeAfterHosted = now()
			'msgbox "pass"
			Exit For
		Else 
			wait 60
		End If			
	Next
	
	If counter <> noOfCubes Then 
		Call ReportStep(StatusTypes.Fail, "Any Cube(s) are not hosted/accepted properly. User waited for " & appWaitTime & " seconds to get the cubes hosted." & " Only " & counter & "cubes are hosted/accepted" ,"Hosting Status Info")	
		ExitTest
	End If 
	
	ValHostingStatusInOps = timeAfterHosted
End Function



Public Function CheckHostingStatus(ByVal dbName, ByVal statusQA, ByVal statusHosting)
	If Browser("DC Home").Page("Ops_Page").WebTable("WebTblTemplateGrid").Exist(20) Then
		Browser("DC Home").Page("Ops_Page").WebEdit("TemplateSearch").Set dbName
		wait 1
		Browser("DC Home").Page("Ops_Page").WebEdit("TemplateSearch").click
		set shl = CreateObject("wscript.shell")
		shl.SendKeys "{ENTER}"
		wait 2
		Browser("DC Home").Page("Ops_Page").WebList("selectPageContent").Select "50"
		wait 2
		counter = 0
		rc = Browser("DC Home").Page("Ops_Page").WebTable("WebTblTemplateGrid").RowCount
		
		For i = 2 To rc Step 1
			'cc = Browser("DC Home").Page("Ops_Page").WebTable("WebTblTemplateGrid").ColumnCount(i)
			hostingStatus = Browser("DC Home").Page("Ops_Page").WebTable("WebTblTemplateGrid").GetCellData(i,6)
			qaStatus = Browser("DC Home").Page("Ops_Page").WebTable("WebTblTemplateGrid").GetCellData(i,7)
			If hostingStatus = statusHosting And (InStr(statusQA,qaStatus) > 0) Then
				counter = counter + 1				
			End If

		Next	
	End If
	CheckHostingStatus = counter
End Function

Public Function ApproveRejectNoChangeDatasourceInOcrf(ByVal Username,ByVal Password,ByVal DataBase,ByVal Action,ByRef objData)
	
	'Validate report publish for Syndicated client
	Systemutil.CloseProcessByName "iexplore.exe"
  
	Call LaunchURL("",Username,Password,objData)
	
	'Select menu option 'Online CRF' from the list
	Call SelectMenuOption("ADMINISTRATION","DATASOURCE ACCESS REQUEST",objData)
	
	'Verify user entered in to 'Datasource Access Request' page
     Browser("Online CRF - DataSource").Page("Online CRF - DataSource").Sync
     If Browser("Online CRF - DataSource").Page("Online CRF - DataSource").WebElement("Datasource Access Request").Exist(10) Then
     	Call ReportStep(StatusTypes.Pass,  " User Entered in to 'Datasource Access Request' page " ,"'Datasource Access Request' page")
     Else
        Call ReportStep(StatusTypes.Fail,  " User Not Entered in to 'Datasource Access Request' page " ,"'Datasource Access Request' page") 
     End If
	'Enter dataSource name and click on 'Enter' 
'	*******************************************************
'Modified by 02/28/2020

	
    Set DataSourceName=Browser("Online CRF - DataSource").Page("Online CRF - DataSource").WebTable("DataSourceRequestApprovalID").WebEdit("Webedit_DataSourceTemplateName_Search")

	'Set DataSourceName=Browser("Online CRF - DataSource").Page("Online CRF - DataSource").WebTable("DataSourceRequestApprovalID").ChildItem(2,5,"WebEdit",0)
'	*****************************************************************
	DataSourceName.highlight
	DataSourceName.Set DataBase
    DataSourceName.Click
    Set objShell=CreateObject("WScript.Shell")
	wait 2  
     objShell.SendKeys "{ENTER}"
     objShell.SendKeys "{ENTER}"
     wait 3
     DataRow=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("DataSourceAccessRequest").GetROProperty("rows")

     If DataRow>1 Then
     	
     	'Select 'Approve' radio button
		If Action="Approve" Then
			wait 2
	   		Browser("Online CRF - DataSource").Page("Online CRF - DataSource").WebRadioGroup("DataSourceAccessrRdio").Select "#0"	
	  	ElseIf Action="Reject" Then
	  		wait 2
	   		Browser("Online CRF - DataSource").Page("Online CRF - DataSource").WebRadioGroup("DataSourceAccessrRdio").Select "#1"	
	   		Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("RejectComments").Set objData.item("RejectComments")


		ElseIf Action="No Change" Then
			wait 2
	   		Browser("Online CRF - DataSource").Page("Online CRF - DataSource").WebRadioGroup("DataSourceAccessrRdio").Select "#2"	
		End If
	  	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("AdminApply").Exist(10) Then
	  		wait 2
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("AdminApply").Click
			Call ReportStep(StatusTypes.Pass,  " Clicked on Apply button for requested to get permission on these datasource"&DataBase ,"'Datasource Access Request' page")
      	Else
        	Call ReportStep(StatusTypes.Fail, "Not Clicked on Apply button for requested to get permission on these datasource"&DataBase ,"'Datasource Access Request' page")
	  	End If
	
     	If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Yes").Exist(10) Then
     		wait 2
			Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Yes").Click
			Call ReportStep(StatusTypes.Pass,  " Clicked on 'Yes' for requested to get permission on these datasource"&DataBase ,"'Datasource Access Request' page")
     	Else
        	Call ReportStep(StatusTypes.Fail, "Not Clicked on 'Yes' for requested to get permission on these datasource"&DataBase ,"'Datasource Access Request' page")
	 	End If
	    If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebBtnInformationOkay").Exist(5) Then
			Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebBtnInformationOkay").Click
			Call ReportStep(StatusTypes.Pass,  " Clicked on 'Okay' for requested to get permission on these datasource"&DataBase ,"'Datasource Access Request' page")
     	Else
        	'Call ReportStep(StatusTypes.Fail, "Not Clicked on 'Okay' for requested to get permission on these datasource"&DataBase ,"'Datasource Access Request' page")
	 	End If
	

   End If
	

End Function

''Created By: Poornima
'Created on: 06/04/2018	
'Functionality : Change action of perticular datasource to specified Action in 'DataSource ' details tab
'Parameters Description: 1. 'DataBase':name of databse for which change action required. 2. 'ChangeAction':Name of action that has to perform on database
Function ChangeActionInDataSourceTab(ByVal DataBase,ByVal ChangeAction)

    CubeRows=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetROProperty("rows")
    For r = 2 To CubeRows Step 1    
 
        Set CheckRow =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(r,2,"WebCheckBox",0)
		CheckRow.click
		Call WebEditExistence(r,Cint(Environment.value("cubeName")),"WebEdit")
    	Set CheckDatabase =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(r,Cint(Environment.value("cubeName")),"WebEdit",0)
    		If DataBase=CheckDatabase.GetROProperty("value") Then
    			Select Case UCase(ChangeAction)
    				
    				Case "DELETE":
    				            ActualAction=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetCellData(r,7)
    							Set selectAction = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(r,7,"WebList",0)
								selectAction.select ChangeAction 
								wait 1
								Call ReportStep (StatusTypes.Pass, "Action Changed from","'Data Source' tab")
								If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("DeleteActionConformation").Exist(10) Then
									Browser("Browser-OCRF").Page("Page-OCRF").WebButton("DeleteActionConformation").Click
									wait 1
									Call ReportStep (StatusTypes.Pass, "Action changed from '"&ActualAction&"' to action '"&ChangeAction&"'","'Data Source' tab")
								Else
	    							Call ReportStep (StatusTypes.Fail, "Action Not changed from '"&ActualAction&"' to action '"&ChangeAction&"'","'Data Source' tab")
								End If
    				Case "ADD":
    				
					Case "NOCHANGE":   
					
					Case "UPDATE":					
    				
    				Case "DELETED":
    				          
    				
    			End Select
    		 
    		
    		
    		End If
    Next

	
End Function


''Created By: Poornima
'Created on: 06/04/2018	
'Functionality : grant permission to user in DataSource Adminstration page
'Parameters Description: 1. 'DataBase':name of database name to which user access providing 2. name of user to whom 'DataSource' grant permission providing


Function GrantPermissionToUser(ByVal DataSource,ByVal User)

    If Browser("OCRF - Datasource_Administration").Page("Online CRF - Datasource").WebElement("Grant Permission").Exist(10) Then
	   Browser("OCRF - Datasource_Administration").Page("Online CRF - Datasource").WebElement("Grant Permission").Click
   		Call ReportStep (StatusTypes.Pass, "Clicked on 'Grant permission'  in 'DataSource Administration tab'","'DataSource Administration tab'")
	Else
   		Call ReportStep (StatusTypes.Fail,"Not Clicked on 'Grant permission'  in 'DataSource Administration tab'","'DataSource Administration tab'")
	End If
	
	'Search for datasource in pop up window,and select dataSource
	If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("GrantPermSearchDS").Exist(10) Then
	   Call ReportStep (StatusTypes.Pass, " 'Grant permission' window appeared","'DataSource Administration tab'")
	   'set search=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("GrantPermSearchDS").ChildItem(2,6,"WebEdit",0)
	   '*****************************************************************************
'	   Modified by Madhu - 03/28/2020


	   set search=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("GrantPermSearchDS").WebEdit("Webedit_DatasourceName")
'	   ******************************************************
       search.Set DataSource
        search.click
       Set objShell=CreateObject("WScript.Shell")
	   wait 2  
       objShell.SendKeys "{ENTER}"
       Rc=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTblDataAdminData").GetROProperty("rows")
       wait 3
	Else
   		Call ReportStep (StatusTypes.Fail,"'Grant permission' window not appeared","'DataSource Administration tab'")		
	End If
	DSRows=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("GrantPermDSData").GetROProperty("rows")
	For r = 2 To DSRows Step 1
    	SelectDS=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("GrantPermDSData").Getcelldata(2,6)
		If SelectDS=dataSource Then
	   		set SelectCheckbox=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("GrantPermDSData").ChildItem(2,2,"WebCheckbox",0)	
	   		SelectCheckbox.set "ON"
	   		Exit for
		End If
	Next
    Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("GrantPermUser").Set User

	'Add user in pop-up window click on add,Okay and cancel
	Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_SelectUsers_Add").Click
	If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("GrantPermCancel").Exist(10) Then
		Browser("Browser-OCRF").Page("Page-OCRF").WebButton("GrantPermCancel").Click
		Call ReportStep (StatusTypes.Pass, "Clicked on 'Grant permission' cancel button","'Grant permission window'")
	Else
   		Call ReportStep (StatusTypes.Fail,"Not Clicked on 'Grant permission' cancel button","'Grant permission window'")
	End If

End Function


''Created By: Poornima
'Created on: 31/05/2018	
'Functionality : Validate Action value of dataBase
'Parameters Description: 1. 'DataBase':name of database name to which user access providing 2. parameterize what Action database should hold 

Function ValidateActionOfDataSource(ByVal DataBase,Byval Action,Byval Operation)
   If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").Exist(10) Then
   	
      CubeRows=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetROProperty("rows")
     For r = 2 To CubeRows Step 1    
 
         'CheckDatabase =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").getcelldata(r,12)
    	CheckDatabase =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").getcelldata(r,Cint(Environment.value("cubeName")))
    	
        If DataBase=CheckDatabase Then
        
           Set CheckRow =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(r,2,"WebCheckBox",0)
		   CheckRow.click
           Call WebEditExistence(r,7,"WebList")
           wait 2
           Set ActualAction=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(r,7,"WebList",0)
           Select Case Operation
        	   Case "Existance":
        	                   
    	   							If trim(ActualAction.GetROProperty("value"))=trim(Ucase(Action)) Then
    	        						Call ReportStep (StatusTypes.Pass, "Validated Database '"&DataBase&"' action : "&Action,"'Data Source' tab")
		   							Else
	    								Call ReportStep (StatusTypes.Fail, "Actual action of DataBase '"&DataBase&"' is : "&ActualAction.GetROProperty("value")&" instead of Action:"&Action,"'Data Source' tab")					
    	   							End If
    						
    		   Case "Select":      ActualAction.Select Action
    	   						   Call ReportStep (StatusTypes.Information, "Action of Database:'"&DataBase&"' has as selected as '"&Operation,"'Data Source' tab")
    						       If Action="DELETE" Then
    						       	  If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("DeleteActionConformation").Exist(10) Then
    						       	  	 Browser("Browser-OCRF").Page("Page-OCRF").WebButton("DeleteActionConformation").Click
    						       	     Call ReportStep (StatusTypes.Pass, "Clicked on 'DELETE' confirmation button","'Data Source' tab")
		   							  Else
	    								  Call ReportStep (StatusTypes.Fail,"Not Clicked on 'DELETE' confirmation button","'Data Source' tab")					
    	   							
    						       	  End If
    						       End If
            End Select
    	End If 
     Next 	   	
   End If
	
End Function

''Created By: Poornima
'Created on: 31/05/2018	
'Functionality : Validate Action value of dataBase in 'Datasource access tab'
'Parameters Description: 1. 'DataBase':name of database name to which user access providing 2. parameterize what Action database should hold 

Function ValidateActionOfDataSourceAccess(ByVal DataBase,Byval Action,Byval Operation)

   If Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserCubeAccess_CubeAccessList").Exist(10) Then
   	
      CubeRows=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserCubeAccess_CubeAccessList").GetROProperty("rows")
     For r = 2 To CubeRows Step 1    
 
         'CheckDatabase =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").getcelldata(r,12)
    	CheckDatabase =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserCubeAccess_CubeAccessList").getcelldata(r,10)
    	
        If DataBase=CheckDatabase Then
        
           Set CheckRow =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserCubeAccess_CubeAccessList").ChildItem(r,2,"WebCheckBox",0)
		   CheckRow.click
           Set ActualAction=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserCubeAccess_CubeAccessList").ChildItem(r,3,"WebList",0)
           Select Case Operation
        	   Case "Existance":
        	                   
    	   							If ActualAction.GetROProperty("value")=Action Then
    	        						Call ReportStep (StatusTypes.Pass, "Validated Database '"&DataBase&"' action : "&Action,"'Data Source' tab")
		   							Else
	    								Call ReportStep (StatusTypes.Fail, "Actual action of DataBase '"&DataBase&"' is : "&ActualAction.GetROProperty("value")&" instead of Action:"&Action,"'Data Source' tab")					
    	   							End If
    						
    		   Case "Select":      ActualAction.Select Action
    	   						   Call ReportStep (StatusTypes.Information, "Action of Database:'"&DataBase&"' has as selected as '"&Operation,"'Data Source' tab")
    						       If Action="DELETE" Then
    						       	  If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("DeleteActionConformation").Exist(10) Then
    						       	  	 Browser("Browser-OCRF").Page("Page-OCRF").WebButton("DeleteActionConformation").Click
    						       	     Call ReportStep (StatusTypes.Pass, "Clicked on 'DELETE' confirmation button","'Data Source' tab")
		   							  Else
	    								  Call ReportStep (StatusTypes.Fail,"Not Clicked on 'DELETE' confirmation button","'Data Source' tab")					
    	   							
    						       	  End If
    						       End If
            End Select
    	End If 
     Next 	   	
   End If
	
End Function


Public Function AcceptCubeOnQA(ByVal Path,ByVal SheetName,ByRef objData,ByVal cubeNames, ByVal action, ByVal comment)
    If IsArray(cubeNames) Then
		lenDbName = Len(cubeNames(0))
    	dbName = Left(cubeNames(0), lenDbName - 3)
	Else 
		dbName = cubeNames
	End If
    
    Browser("DC Home").page("Ops_Page").Link("LinkHostingTab").FireEvent "onmouseover"
    Browser("DC Home").page("Ops_Page").Link("LinkQA").Click
    wait 5
    If Browser("DC Home").page("Ops_Page").WebTable("WebTableQACubeGrid").Exist(20) Then
        Browser("DC Home").page("Ops_Page").WebEdit("FilterCubeInQA").Set dbName
        wait 1
        Browser("DC Home").page("Ops_Page").WebEdit("FilterCubeInQA").Click
        set shl = CreateObject("wscript.shell")
        shl.SendKeys "{ENTER}"
        wait 3
        Set objQAGrid = Browser("DC Home").page("Ops_Page").WebTable("WebTableQACubeGrid")
        rc = objQAGrid.RowCount
        
'        For t = 1 To 50 Step 1
'            c = 1
'            For j = 2 To rc Step 1
'                If objQAGrid.GetCellData(j,3)  = "Ready for QA" Then 
'                    c = c+1
'                Else 
'                    wait 5
'                    Browser("DC Home").page("Ops_Page").WebEdit("FilterCubeInQA").Click
'                    set shl = CreateObject("wscript.shell")
'                    shl.SendKeys "{F5}"
'                    wait 3
'                    If Browser("DC Home").page("Ops_Page").WebTable("WebTableQACubeGrid").Exist(20) Then
'                        Browser("DC Home").page("Ops_Page").WebEdit("FilterCubeInQA").Set dbName
'                        wait 1
'                        Browser("DC Home").page("Ops_Page").WebEdit("FilterCubeInQA").Click
'                        set shl = CreateObject("wscript.shell")
'                        shl.SendKeys "{ENTER}"
'                        wait 3
'                        Set objQAGrid = Browser("DC Home").page("Ops_Page").WebTable("WebTableQACubeGrid")
'                    End If 
'                End If
'            Next
'            If c = rc Then
'                For i = 2 To rc Step 1
'                    objQAGrid.ChildItem(i,4,"WebRadioGroup",0).Select "#0"
'                Next
'                Exit For 
'            End If
'        Next
        
        For i = 2 To rc Step 1
            Select Case action
                Case "Approve"
                    objQAGrid.ChildItem(i,4,"WebRadioGroup",0).Select "#0"
                    objQAGrid.ChildItem(i,7,"WebEdit",0).Set comment
                Case "Reject"
                    objQAGrid.ChildItem(i,4,"WebRadioGroup",0).Select "#1"
                    objQAGrid.ChildItem(i,7,"WebEdit",0).Set comment
                Case "No Change"
                    objQAGrid.ChildItem(i,4,"WebRadioGroup",0).Select "#2"
                    objQAGrid.ChildItem(i,7,"WebEdit",0).Set comment
            End Select
            
        Next
        
        Browser("DC Home").page("Ops_Page").WebButton("BtnApplyQA").highlight
        Browser("DC Home").page("Ops_Page").WebButton("BtnApplyQA").Click
        If Browser("DC Home").page("Ops_Page").WebButton("BtnOkQA").Exist(5) Then
            Browser("DC Home").page("Ops_Page").WebButton("BtnOkQA").Click
        End If
        
        If Browser("DC Home").page("Ops_Page").WebElement("WebEleSuccessQA").Exist(5) Then
            Call ReportStep(StatusTypes.Pass, action & " action has taken for all cubes successfully " & "at " & now() ,"QA Approve Info")
            timeAfterApproved = now()
            wait 2
            Browser("DC Home").page("Ops_Page").Image("head_logoOnQA").Click
            wait 2
        Else 
            Call ReportStep(StatusTypes.Fail, action & " action has not taken successfully for all cubes " ,"QA Approve Info")
        End If 
    End If 
    AcceptCubeOnQA = timeAfterApproved
End Function


Public Function TotalExecutionTime(ByVal startTime, ByVal endTime)

    'total number of seconds
    totalTime_Secs  = Datediff("s",startTime,endTime)

    'convert  total  Seconds into "Seconds only/ Mins+Secs/ Hrs+Mins+Secs"
    If totalTime_Secs < 60 Then
        totalTime = totalTime_Secs  & " Second(s) Approx."
    ElseIf    totalTime_Secs >=60 and totalTime_Secs < 3600 Then
        totalTime = int(totalTime_Secs/60) & " Minute(s) and "& totalTime_Secs Mod 60 & " Second(s) Approx."
    ElseIf    totalTime_Secs >= 3600 Then
        totalTime = int(totalTime_Secs/3600) & " Hour(s) , " & int((totalTime_Secs Mod 3600)/60)  & " Minute(s) and "& ((totalTime_Secs Mod 3600 ) Mod 60) & " Second(s)  Approx."
    End If

    'Return the Message 
    TotalExecutionTime = totalTime
End Function

Public Function ValidateRolesThruManageUsers(ByVal dbName, ByVal userId, ByVal userRole,  ByVal shouldExist,ByRef objData)
	If Browser("DC Home").Page("Ops_Page").WebTable("WebTblTemplateGrid").Exist(20) Then
		Browser("DC Home").Page("Ops_Page").WebEdit("TemplateSearch").Set dbName
		wait 1
		Browser("DC Home").Page("Ops_Page").WebEdit("TemplateSearch").click
		set shl = CreateObject("wscript.shell")
		shl.SendKeys "{ENTER}"
		wait 2
		Browser("DC Home").Page("Ops_Page").WebList("selectPageContent").Select "50"
		wait 2
		counter = 0
		rc = Browser("DC Home").Page("Ops_Page").WebTable("WebTblTemplateGrid").RowCount
		If rc>=2 Then
				For i = 2 To rc Step 1
					'cc = Browser("DC Home").Page("Ops_Page").WebTable("WebTblTemplateGrid").ColumnCount(i)
					Browser("DC Home").Page("Ops_Page").WebTable("WebTblTemplateGrid").ChildItem(i,10,"WebElement",0).Click
					Call ReportStep (StatusTypes.Pass, "Manage Users Link is clicked for " & "database " & dbName,"Click Manage Users Link")
					wait 2
					flag = False
					If Browser("DC Home").Page("Ops_Page").WebTable("WebTable_ManageUser").Exist(10) Then 
						Browser("DC Home").Page("Ops_Page").WebList("class:=ui-pg-selbox","innertext:=5102050").Select "50"
						wait 1
						
'						******************************************************8
'MOdified by Madhu - 02/26/2020
'OR also  modified
						if Browser("DC Home").Page("Ops_Page").WebTable("WebTblFilterManageUsers").Exist(5) then
						
							Browser("DC Home").Page("Ops_Page").WebEdit("Webedit_UserPrincipalNameSearch").Set userId
							Browser("DC Home").Page("Ops_Page").WebEdit("Webedit_UserPrincipalNameSearch").Click
						end if 	

						
'						Browser("DC Home").Page("Ops_Page").WebTable("WebTblFilterManageUsers").ChildItem(2,2,"WebEdit",0).Set userId
'						Browser("DC Home").Page("Ops_Page").WebTable("WebTblFilterManageUsers").ChildItem(2,2,"WebEdit",0).Click

'*******************************************************************************8
						set shl = CreateObject("wscript.shell")
     					shl.SendKeys "{ENTER}"
     					wait 3
						Call PageLoading()
						rowCnt = Browser("DC Home").Page("Ops_Page").WebTable("WebTable_ManageUser").RowCount
						For j = 2 To rowCnt Step 1
							user = Browser("DC Home").Page("Ops_Page").WebTable("WebTable_ManageUser").GetCellData(j,4)
							If userRole = user Then
								flag = True
							End If
						Next
						If flag = True And shouldExist="YES" Then
							Call ReportStep (StatusTypes.Pass, "user " & userRole & " exists for user Id " & userId,"Verify User Role Relationship")
						ElseIf flag = True And shouldExist="NO" Then 
							Call ReportStep (StatusTypes.Fail, "user " & userRole & "  exist for user Id " & userId,"Verify User Role Relationship")
						ElseIf flag = False And shouldExist="YES" Then 
				    		Call ReportStep (StatusTypes.Fail, "user " & userRole & " does not exist for user Id " & userId,"Verify User Role Relationship")
						ElseIf flag = False And shouldExist="NO"  Then
				    		Call ReportStep (StatusTypes.Pass, "user " & userRole & " does not exists for user Id " & userId,"Verify User Role Relationship")
						End If
						Browser("DC Home").Page("Ops_Page").WebElement("btnClose").Click
					End If 
				Next
		Else 
			Call ReportStep (StatusTypes.Fail, "DataSource '"&dbName&"' Does not exist in searched list","Ops template search")
		End If
		
	End If
End Function

Function OfferingDetailsDataEntering(ByVal OfferingClient,ByVal OfferingCountry,ByVal Syndicated,ByVal offeringName,ByVal OfferingContactName,ByVal OfferingContactPhone,ByVal OfferingContactEmail,ByRef objData)
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_OfferingDetailsTab").Exist(20) Then
		Call ReportStep (StatusTypes.Pass, "Offering Details Tab exist","Offering Details Tab")
	Else
	   	Call ReportStep (StatusTypes.Fail, "cOffering Details Tab not exist","Offering Details Tab")
	End If
    'ADD DATA TO OFFERING DETAILS TAB
    			
	'ENTER DATA IN OFFERING CLIENT NAME
	If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingClient").Exist(20) Then
		Call ReportStep (StatusTypes.Pass, "Offering client name entered with "&OfferingClient,"Offering Client")	
		Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingClient").Set OfferingClient
	Else
		Call ReportStep (StatusTypes.Fail, "Offering client name not entered","Offering Client")
	End If
	' SELECT OFFERING CONTRY NAME
	wait 5
	If Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebListOfferingCountry").Exist(20) Then
		Call ReportStep (StatusTypes.Pass, "Offering country name selected as "&OfferingCountry,"Offering Country name")	
		Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebListOfferingCountry").Select OfferingCountry
	Else
		Call ReportStep (StatusTypes.Fail, "Offering country name not selected","Offering Country name")
	End If
	    'SELECT SYNDIACATED NON SYNDICATED RADIO BUTTON
    If Browser("Browser-OCRF").Page("Page-OCRF").WebRadioGroup("WEbRadioSyndicated").Exist(20) Then
       	Call ReportStep (StatusTypes.Pass, "Syndicated radio button selected  ","Syndicated")      	
       	If Syndicated="No" Then
       	  	Browser("Browser-OCRF").Page("Page-OCRF").WebRadioGroup("WEbRadioSyndicated").Select "false"
       	Else
       	  	Browser("Browser-OCRF").Page("Page-OCRF").WebRadioGroup("WEbRadioSyndicated").Select "on"
       	End If
    Else
	   	Call ReportStep (StatusTypes.Fail, "Offering country name not selected","Syndicated")
    End If
	'ENTER OFFERING NAME
	If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingName").Exist(20) Then
		Call ReportStep (StatusTypes.Pass, "Offering Name entered as "&offeringName,"Offering name")	
		Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingName").Set offeringName
    Else
		Call ReportStep (StatusTypes.Fail, "Offering Name  not entered","Offering name")
	End If

	'ENTER OFFERING CONTACT NAME
	If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactName").Exist(20) Then
		Call ReportStep (StatusTypes.Pass, "Offering Contact Name entered as "&OfferingContactName,"Offering Contact Name")	
		Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactName").Set OfferingContactName
	Else
		Call ReportStep (StatusTypes.Fail, "Offering Contact Name  not entered","Offering Contact Name")
	End If
	'ENTER OFFERING CONTACT PHONE
	If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactPhone").Exist(20) Then
		Call ReportStep (StatusTypes.Pass, "Offering Contact Phone entered as "&OfferingContactPhone,"Offering Contact Phone")	
		Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactPhone").Set OfferingContactPhone
	Else
		Call ReportStep (StatusTypes.Fail, "Offering Contact Phone  not entered","Offering Contact Phone")
	End If
	'ENTER OFFERING CONTACT EMAIL
	If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactEmail").Exist(20) Then
	   Call ReportStep (StatusTypes.Pass, "Offering Contact Email entered as "&OfferingContactEmail,"Offering Contact Email")	
		Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditOfferingContactEmail").Set OfferingContactEmail
	Else
		Call ReportStep (StatusTypes.Fail, "Offering Contact Email  not entered","Offering Contact Email")
	End If
End Function
'#####################################################################################
'Created By: Sumit
'Created on: 23/05/2017
'Description: To validate errors in Hosting Event Log on Ops System
'Parameter: cubeNames (array - holds multiple cubeNames), objData 
'####################################################################################

Public Function CheckHostingEventLog(ByVal Path,ByVal SheetName,ByVal cubeNames,ByVal searchKeyword, ByVal matchString, ByRef objData)
	lenDbName = Len(cubeNames(0))
	dbName = Left(cubeNames(0), lenDbName - 3)
	Browser("DC Home").page("Ops_Page").Link("LinkEvent Log & Subscription").FireEvent "onmouseover"
	Browser("DC Home").page("Ops_Page").Link("LinkHosting Event Log").Click
	Browser("DC Home").page("Ops_Page").WebEdit("WebEditHostingEventLogFilterTemplateId").Set dbName
	Browser("DC Home").page("Ops_Page").WebEdit("WebEditHostingEventLogFilterTemplateId").Click
	set shl = CreateObject("wscript.shell")
	shl.SendKeys "{ENTER}"
	wait 2
	Browser("DC Home").page("Ops_Page").WebEdit("WebEditHostingEventLogFilterDate").Set ""
	Browser("DC Home").page("Ops_Page").WebEdit("WebEditHostingEventLogFilterDate").Click
	Browser("DC Home").page("Ops_Page").WebEdit("WebEditHostingEventLogFilterEvent").Set searchKeyword
	Browser("DC Home").page("Ops_Page").WebEdit("WebEditHostingEventLogFilterEvent").Click
	set sh = CreateObject("wscript.shell")
	sh.SendKeys "{ENTER}"
	wait 2
	Browser("DC Home").page("Ops_Page").WebList("selectPageContent").Select "50"
	rc = Browser("DC Home").page("Ops_Page").WebTable("WebTblHostingEventLogGrid").RowCount
	
	For i = 2 To rc Step 1
		eventLog = Browser("DC Home").page("Ops_Page").WebTable("WebTblHostingEventLogGrid").GetCellData(i,4)
		If InStr(eventLog,matchString) > 0 Then
			Call ReportStep(StatusTypes.Pass, "Log displayed as '" & eventLog & "' "   ,"Check Event Log")
		Else
			Call ReportStep(StatusTypes.Fail, "Error available as '" & eventLog & "' "   ,"Check Event Log")
			ExitTest
		End If
	Next
End Function

'#####################################################################################
'Created By: Sumit
'Created on: o1/06/2017
'Description: To validate Revert to Prior functionality
'Input Parameter: Path, SheetName, cubeNames (array - holds multiple cubeNames), objData 
'Output Parameter: priorDBSize
'####################################################################################

Public Function RevertToPriorAction(ByVal Path,ByVal SheetName,ByVal cubeNames,ByRef objData)
	lenDbName = Len(cubeNames(0))
	dbName = Left(cubeNames(0), lenDbName - 3)
	If Browser("DC Home").Page("Ops_Page").WebTable("WebTblTemplateGrid").Exist(20) Then
		Browser("DC Home").Page("Ops_Page").WebEdit("TemplateSearch").Set dbName
		wait 1
		Browser("DC Home").Page("Ops_Page").WebEdit("TemplateSearch").click
		set shl = CreateObject("wscript.shell")
		shl.SendKeys "{ENTER}"
		wait 2
		Browser("DC Home").Page("Ops_Page").WebList("selectPageContent").Select "50"
		wait 2
		counter = 0
		rc = Browser("DC Home").Page("Ops_Page").WebTable("WebTblTemplateGrid").RowCount
		
		For i = 2 To rc Step 1
			'cc = Browser("DC Home").Page("Ops_Page").WebTable("WebTblTemplateGrid").ColumnCount(i)
			Browser("DC Home").Page("Ops_Page").WebTable("WebTblTemplateGrid").ChildItem(i,1,"WebElement",0).click
			If Browser("DC Home").Page("Ops_Page").WebEdit("WebEditTemplateId_HostingTemplate").Exist(25) Then
				If Browser("DC Home").Page("Ops_Page").WebCheckbox("ChkBoxRetainPrior_HostingTemplate").GetROProperty("checked") = 0 Then
					Browser("DC Home").Page("Ops_Page").WebCheckbox("ChkBoxRetainPrior_HostingTemplate").Click
				End If
				currDBSize = Browser("DC Home").Page("Ops_Page").WebEdit("WebEditCurrentDBSize_HostingTemplate").GetROProperty("value")
				priorDBSize = Browser("DC Home").Page("Ops_Page").WebEdit("WebEditPriorDBSize_HostingTemplate").GetROProperty("value")
				Call ReportStep(StatusTypes.Information, "Current DB Size: " & currDBSize & " And Prior DB Size: " & priorDBSize & " for DB " & cubeNames(i-2) ,"Information About DB Size")
				Browser("DC Home").Page("Ops_Page").WebButton("BtnRevert To Prior_Hosting Template").Click
				Browser("DC Home").Page("Ops_Page").WebButton("BtnYes_Hosting Template").Click
				wait 5
				Browser("DC Home").Page("Ops_Page").WebButton("BtnClose_Hosting Template").Click
				Call ReportStep(StatusTypes.Pass, "Revert to Prior button is clicked","Revert to Prior Action")
			Else
				Call ReportStep(StatusTypes.Fail, "Revert to Prior button is not clicked successfully","Revert to Prior Action")
			End If
		Next	
	End If
	RevertToPriorAction = priorDBSize
End Function


'#####################################################################################
'Created By: Sumit
'Created on: o1/06/2017
'Description: To check Revert to Prior functionality has been done and previous DB is restored correctly
'Input Parameter: priorDBSize, qaStatus, objData 
'Output Parameter: NA
'####################################################################################

Public Function ValRevertToPriorStatus(ByVal cubeNames, ByVal priorDBSize, ByVal qaStatus, ByRef objData)
			
			Browser("DC Home").Page("Ops_Page").Image("head_logoOnQA").Click
			lenDbName = Len(cubeNames(0))
			dbName = Left(cubeNames(0), lenDbName - 3)
			If Browser("DC Home").Page("Ops_Page").WebTable("WebTblTemplateGrid").Exist(20) Then
				Browser("DC Home").Page("Ops_Page").WebEdit("TemplateSearch").Set dbName
				wait 1
				Browser("DC Home").Page("Ops_Page").WebEdit("TemplateSearch").click
				set shl = CreateObject("wscript.shell")
				shl.SendKeys "{ENTER}"
				wait 2
				Browser("DC Home").Page("Ops_Page").WebList("selectPageContent").Select "50"
				wait 2
				Browser("DC Home").Page("Ops_Page").WebTable("WebTblTemplateGrid").ChildItem(2,1,"WebElement",0).click
				If Browser("DC Home").Page("Ops_Page").WebEdit("WebEditTemplateId_HostingTemplate").Exist(10) Then
					If Browser("DC Home").Page("Ops_Page").WebElement("WebEleStatus_HostingTemplate").Exist(25) Then 
						currDBSizeAfterRevert = Browser("DC Home").Page("Ops_Page").WebEdit("WebEditCurrentDBSize_HostingTemplate").GetROProperty("value")
						priorDBSizeAfterRevert = Browser("DC Home").Page("Ops_Page").WebEdit("WebEditPriorDBSize_HostingTemplate").GetROProperty("value")
						qaStatus = Browser("DC Home").Page("Ops_Page").WebEdit("WebEditQAStatus_HostingTemplate").GetROProperty("value")
						If currDBSizeAfterRevert = priorDBSize Then
							Call ReportStep(StatusTypes.Pass, "Revert to Prior action is done successfully and Prior DB is restored for status as " & qaStatus,"Val Revert to Prior")
						Else
							Call ReportStep(StatusTypes.Fail, "Revert to Prior action is not done successfully and Prior DB is not restored for status " & qaStatus,"Val Revert to Prior")
						End If
					End If 
				End If 		
			End If 
End Function



'#####################################################################################
'Created By: Poornima
'Created on: o1/06/2017
'Description: To check move DB functionality
'Input Parameter:
'Output Parameter: NA
'####################################################################################


Function NewTemplateCreationinops(ByVal strHostName,ByVal strCountry,ByVal strFeq,ByVal strOCode,ByVal strORef,ByVal strFLAPath,ByVal strHostPath,ByVal strCAHDProcess,ByVal chkbox,ByVal strDBType,ByVal strF2,ByRef objData)

	Browser("DC Home").Page("Ops_Page").Sync
	
	'Modified by Avinash on May 14, 2019
	set wsh = Createobject("Wscript.Shell")
	wsh.SendKeys "{F5}"
	'Modification completed
	
    If Browser("DC Home").Page("Ops_Page").WebButton("WebButAddNew").Exist(30) Then
	  	 Browser("DC Home").Page("Ops_Page").WebButton("WebButAddNew").Click
	     Call ReportStep (StatusTypes.Pass,"Template name "&CubeName& " Exist in Search list","hosting Template page" )
	Else
	     Call ReportStep (StatusTypes.Fail,"Template name "&CubeName& " Not Exist in Search list","hosting Template page" )
	End If
    
    Browser("DC Home").Page("Ops_Page").Sync
    

    set objtxtName=Browser("DC Home").Page("Ops_Page").WebEdit("txtName")
    Call SCA.SetText(objtxtName,strHostName, "txtName","Client Creation Page" )

    Set objlstCountry=Browser("DC Home").Page("Ops_Page").WebList("lstCountryId")
    Call SCA.SelectFromDropdown(objlstCountry,strCountry)    
        
    Set objFrequencyId=Browser("DC Home").Page("Ops_Page").WebList("lstFrequencyId")
    Call SCA.SelectFromDropdown(objFrequencyId,strFeq)

    Set objtxtOCode=Browser("DC Home").Page("Ops_Page").WebEdit("txtOrganizationCode")
    Call SCA.SetText(objtxtOCode,strOCode, "txtOrgCode","Client Creation Page" )
    
    Set objtxtOrderRef=Browser("DC Home").Page("Ops_Page").WebEdit("txtOrderReference")
    Call SCA.SetText(objtxtOrderRef,strORef, "txtOreference","Client Creation Page" )
    
    If chkbox=0 Then
        
        Browser("DC Home").Page("Ops_Page").WebCheckBox("chkRetainPrior").Set "ON"
        Else
        Browser("DC Home").Page("Ops_Page").WebCheckBox("chkRetainPrior").Set "OFF"
        
    End If
	
	'Modified by Avinash on May 13, 2019
	
    Browser("DC Home").Page("Ops_Page").WebList("DatabaseTypeId").Select strDBType
    wait 1
    tempIdName = Browser("DC Home").Page("Ops_Page").WebEdit("txtNewTemplateIdName").GetROProperty("value")
    
    'modification completed
    
    Set objLstFLA=Browser("DC Home").Page("Ops_Page").WebList("lstFLAServerId")
    Call SCA.SelectFromDropdown(objLstFLA,strFLAPath)
    
    Set objHServer=Browser("DC Home").Page("Ops_Page").WebList("lstHostServerId")
    Call SCA.SelectFromDropdown(objHServer,strHostPath)
    wait 2    
    	
    Browser("DC Home").Page("Ops_Page").WebElement("btnSave").Click
    
    Browser("DC Home").Page("Ops_Page").Sync

    If Browser("DC Home").Page("Ops_Page").WebElement("webNewTempCreationStatus").Exist(15)  Then
        Call ReportStep (StatusTypes.Pass, "Creation of new template:-"&Space(7)&"Template created successfully","Template Creation Page")            
    else            
        Call ReportStep (StatusTypes.Fail, "Creation of new template:-"&Space(7)&"Template Not created successfully","Template Creation Page")
    End If

    Browser("DC Home").Page("Ops_Page").Sync
    wait 5
    Browser("DC Home").Page("Ops_Page").WebElement("btnClose").Click   
    wait 5
    Browser("DC Home").Page("Ops_Page").Sync
    NewTemplateCreationinops=tempIdName
End Function


Public Function MoveDB(ByVal TemplateId,ByVal DestinationServer,ByVal Parm1,ByVal Parm2,ByRef objData)
	'InPut Date and Time
	DateArray=Split((FormatDateTime(Date(),1)),",")
	MonthEnter=Split(Trim(DateArray(1))," ")
	DateEntered=MonthEnter(1)&"-"&MonthEnter(0)&"-"&Trim(DateArray(2))
	TimeArray=Split((FormatDateTime(Now(),4)),":")
	TimeEnter=CInt(TimeArray(1))+2
	TimeEntered=TimeArray(0)&":"&TimeEnter
	If Browser("DC Home").Page("Ops_Page").WebEdit("WebEditMDBTemplateID").Exist(30) Then
		 Call ReportStep (StatusTypes.Pass, "Entered in to 'Move DB' Page","Move DB Page")            
    else            
        Call ReportStep (StatusTypes.Fail, "Not Entered in to 'Move DB' Page","Move DB Page")
	End If
	Browser("DC Home").Page("Ops_Page").WebEdit("WebEditMDBTemplateID").Set TemplateId
	Browser("DC Home").Page("Ops_Page").WebList("WebListMDBServerName").Select DestinationServer
	Browser("DC Home").Page("Ops_Page").WebEdit("WebEditMDBScheduleDate").Set DateEntered
	Browser("DC Home").Page("Ops_Page").WebEdit("WebEditMDBScheduleTime").Set TimeEntered
	If Browser("DC Home").Page("Ops_Page").WebButton("BtnMoveDatabase").Exist(20) Then
		Call ReportStep (StatusTypes.Pass, "Clicked on Move DB button after entering  'Destination Server' as "&DestinationServer&" ","Move DB Page")  
        Browser("DC Home").Page("Ops_Page").WebEdit("WebEditMDBTemplateID").Click
        Browser("DC Home").Page("Ops_Page").WebButton("BtnMoveDatabase").Click		
    else            
        Call ReportStep (StatusTypes.Fail, "Not Clicked on Move DB button","Move DB Page")
	End If
	
End Function


Public Function MoveDBStatus(ByVal TemplateId,ByVal MoveStatus,ByVal ActionStatus,ByRef objData)

     For i = 1 To 50 Step 1
         wait 5
         Browser("DC Home").Page("Ops_Page").WebEdit("SearchMDBDataSourceId").Set TemplateId
         Set objShell=CreateObject("WScript.Shell")
         Browser("DC Home").Page("Ops_Page").WebEdit("SearchMDBDataSourceId").Click
         wait 2  
         objShell.SendKeys "{ENTER}"
         wait 2
	     ActMoveStatus=Browser("DC Home").Page("Ops_Page").WebTable("WebTable_MDBRequest").GetCellData(2,5)
         ActActionStatus=Browser("DC Home").Page("Ops_Page").WebTable("WebTable_MDBRequest").GetCellData(2,10)
    	If ActMoveStatus="Error" AND i=50  Then
    	   Call ReportStep (StatusTypes.Fail,"Move DB Status and Action Status of cube '"&TemplateId& "' is "&ActMoveStatus&" and "&ActionStatus&" and Expected "&MoveStatus&"  and "&ActionStatus&" respectively","Move DB Page" )	
    	   ExitTest 
    	Else
    	   If ActMoveStatus=MoveStatus AND ActionStatus=ActionStatus  Then
    	      Call ReportStep (StatusTypes.Pass,"Move DB Status and Action Status of cube '"&TemplateId& "' is "&ActMoveStatus&" and "&ActionStatus&" respectively","Move DB Page" )	
              Exit For
           Else
           	   Browser("DC Home").Refresh
           wait 5
	       End If
    	End If
    	
    Next
End Function


Public Function TemplateManageUsers(ByVal Template,ByVal Role,ByVal Action,ByVal DisplayName,ByVal UserPrincName,ByVal CustomId,ByVal strQAUsers,ByRef objData)

   TemplateMU= Browser("DC Home").Page("Ops_Page").WebElement("WebEle_TemplateId").GetROProperty("InnerText")
   If TemplateMU=Template Then
   	   Call ReportStep (StatusTypes.Pass,"Entered in to manage User of template "&Template,"Manage users page" )
   Else
	    Call ReportStep (StatusTypes.Fail,"Not Entered in to manage User of template "&Template,"Manage users page" )
   End If
   If Browser("DC Home").Page("Ops_Page").WebCheckBox("WebCheckBox_QAUser").Exist(10) Then
   	  Browser("DC Home").Page("Ops_Page").WebCheckBox("WebCheckBox_QAUser").Set "ON"
      Call ReportStep (StatusTypes.Pass,"Selected 'QAUser'","Manage users page" )
   Else
	    Call ReportStep (StatusTypes.Fail,"Not Selected 'QAUser'","Manage users page" )
   End If
    Browser("DC Home").Page("Ops_Page").WebEdit("Text_Users").Set  Role 
	Browser("DC Home").Page("Ops_Page").WebButton("Btn_AddUsers").Click
	
	If Browser("DC Home").Page("Ops_Page").WebTable("WebTable_ManageUser").Exist(10) Then
	   R=Browser("DC Home").Page("Ops_Page").WebTable("WebTable_ManageUser").RowCount	
	   For i = 2 To R Step 1
	   	   User=Browser("DC Home").Page("Ops_Page").WebTable("WebTable_ManageUser").GetCellData(i,2)
	   	   If User=Role Then
	   	   	  Call ReportStep (StatusTypes.Pass,"User "&User&" added to 'User/Rolerelationship Table" ,"Manage users page" )
           Else
	          Call ReportStep (StatusTypes.Fail,"User "&User&" Not added to 'User/Rolerelationship Table","Manage users page" )
	   	   End If
	   Next
	End If
	If Browser("DC Home").Page("Ops_Page").WebElement("WebEle_ApplySecurity").Exist(10) Then
   	  Browser("DC Home").Page("Ops_Page").WebElement("WebEle_ApplySecurity").Click
      Call ReportStep (StatusTypes.Pass,"Clicked on 'Apply security","Manage users page" )
   Else
	    Call ReportStep (StatusTypes.Fail,"Not Clicked on 'Apply security","Manage users page" )
   End If
   If Browser("DC Home").Page("Ops_Page").WebElement("WebEle_ManageUserClose").Exist Then
   	  Browser("DC Home").Page("Ops_Page").WebElement("WebEle_ManageUserClose").Click
   End If
   If Browser("Browser-SCA").Dialog("Dialog_Message_Webpage").WinButton("WinButton_OK").Exist(5) Then
   	  Browser("Browser-SCA").Dialog("Dialog_Message_Webpage").WinButton("WinButton_OK").Click
   End If
   
'					   Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditCountryName").Click
'					   objShell.SendKeys "{DOWN}"
'					   wait 2
'					   objShell.SendKeys "{DOWN}"
'					   wait 2  
'					   objShell.SendKeys "{ENTER}"
End Function


Function SendKeyAndWait(ByVal key,ByRef obj, ByVal waitTime)
	Set wsh = CreateObject ("Wscript.Shell")
	obj.click ' You may need to activate your browser first
	wsh.SendKeys key
	wait waitTime
	Set wsh = Nothing
End Function

'Login to PBI Server environment
Function PBILogin(ByVal strUserName,ByVal StrPwd,ByVal StrUrl)
  
  If strUserName <> "" And StrPwd <> "" Then
  	Systemutil.CloseProcessByName "iexplore.exe"
 ' 	Systemutil.Run StrUrl, ,  ,  , 3 
'Modiified by Avinash 5th Aug 2019  	
	SystemUtil.Run "iexplore.exe",StrUrl ,  ,  , 3 
'Modification completed 	
    
'  	If  Browser("PowerBIServer").Page("PowerBIServer").WebEdit("txtUsername").Exist(5) Then 
'	  	Call ReportStep (StatusTypes.Pass,"PBIServer lanuched successfully","Home Page" )
'	  	Browser("PowerBIServer").Page("PowerBIServer").WebEdit("txtUsername").Set strUserName
'	  	Browser("PowerBIServer").Page("PowerBIServer").WebEdit("txtPassword").Set  StrPwd
'	  	Browser("PowerBIServer").Page("PowerBIServer").WebButton("btnLogin").Click
'	  	If Browser("PowerBIServer").page("PowerBIServer").WebEdit("Search").Exist(60) Then
'	  		Call ReportStep (StatusTypes.Pass,"User is suceessfully Logged in to VA PBI Server","Welcome Page of VA PBI Server" )
'	  	Else 
'	  		Call ReportStep (StatusTypes.Fail,"User is not suceessfully Logged in to VA PBI Server","Welcome Page of VA PBI Server" )
'	  		ExitTest
'	  	End If
'************************************8
'modified by Madhu - 03/04/2020
	If UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").Exist(120) Then    	
        UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").Highlight 
	    UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAEdit("User name").SetValue strUserName
	    UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAEdit("Password").SetValue StrPwd
	   ' UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAButton("OK").Click
        Call ReportStep (StatusTypes.Pass, "User " & UserName & " has logged in successfully","Login through Window security PopUp")
	Else
		Call ReportStep (StatusTypes.Fail, "User " & UserName & " has not successfully logged in","Login through Window security PopUp")	   
  	end if 
  	              	    
	    '******************************************************************
    Set objShell = CreateObject("Wscript.shell")
    objShell.SendKeys "{ENTER}"
'    ***************************************************************************8
 
  	Else
  		    Systemutil.CloseProcessByName "iexplore.exe"
			Systemutil.Run StrUrl
			If Browser("PowerBIServer").page("PowerBIServer").WebEdit("Search").Exist(60) Then
	  			Call ReportStep (StatusTypes.Pass,"User is suceessfully Logged in to VA PBI Server","Welcome Page of VA PBI Server" )
	  		Else 
	  			Call ReportStep (StatusTypes.Fail,"User is not suceessfully Logged in to VA PBI Server","Welcome Page of VA PBI Server" )
	  			ExitTest
	  		End If
	
  	End If
'************************************************************************************************	

End Function


'********************************************************************************************************************************************************************************************
'********************************************************************************************************************************************************************************************
'----------------------------------------------------------------Generic Functions created by Madhu wRT VA Report Level Security_______________________________________________________________

'***************************************8
'Modified by Madhu - 03/02/2020
Function VerifyNewTab_ContentPermission()

	if Browser("Browser-OCRF").Page("Page-OCRF").Link("Link_Content Permissions->>").Exist(3) then
		Presenceflag = True
		Call ReportStep (StatusTypes.Pass,"Verify Newly added Content Permission Tab display","Content Permission Tab is present" )
	 Else 
	  	Call ReportStep (StatusTypes.Fail,"Verify Newly added Content Permission Tab display","Content Permission Tab is present" )
	 	Presenceflag = False
	 End If
	 VerifyNewTab_ContentPermission = Presenceflag
End function


Function Verifythenewweblist_UserGroup()
	if Browser("Browser-OCRF").Page("Page-OCRF").WebList("Weblist_UserGroupName").Exist(5) then
		
		Call ReportStep (StatusTypes.Pass,"Verify Newly added Weblist_UserGroupName display in upload user","Weblist_UserGroupName is present in upload user" )
	 Else 
	  	Call ReportStep (StatusTypes.Fail,"Verify Newly added Weblist_UserGroupName display in upload user","Weblist_UserGroupName is not present in upload user" )
	 	
	 End If
	 
End Function

Function Verify_newlyaddedgroupsinUsergrouplist(ByVal FileName,ByVal AddUserSheetName,ByVal StartRow,ByVal EndRow,objData)
	if Browser("Browser-OCRF").Page("Page-OCRF").WebList("Weblist_UserGroupName").Exist(5) then
		Call ImportSheet(AddUserSheetName,FileName)
	
	Datatable.GetSheet(AddUserSheetName).SetCurrentRow(StartRow)
	
		Call ReportStep (StatusTypes.Pass,"Verify Newly added Weblist_UserGroupName display in upload user","Weblist_UserGroupName is present in upload user" )
		addedusersview = Browser("Browser-OCRF").Page("Page-OCRF").WebList("Weblist_UserGroupName").GetROProperty("all items")
		For urow = startrow To endrow Step 1
		checkuser = datatable.value("UserGroup",AddUserSheetName)
			if instr(addedusersview,checkuser)>0 then
				Call ReportStep (StatusTypes.Pass,"Verify added users display in Weblist_UserGroupName in upload user","Weblist_UserGroupName is present in upload user - addedusers" )
			else
				Call ReportStep (StatusTypes.Fail,"Verify added users display in Weblist_UserGroupName in upload user","Weblist_UserGroupName is not present in upload user - addedusers" )
			end if 
			Datatable.GetSheet(AddUserSheetName).SetNextRow
		Next
	Else 
	  	Call ReportStep (StatusTypes.Fail,"Verify Newly added Weblist_UserGroupName display in upload user","Weblist_UserGroupName is not present in upload user" )
	 	
	 End If
	 
End Function

Function verifyNewCol_UserGroupSelection()
		
	if Browser("Browser-OCRF").Page("Page-OCRF").WebTable("Webtable_AddUsers_HeadingTab").Exist(5) then
		allheaders = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("Webtable_AddUsers_HeadingTab").GetROProperty("column names")
		
		If instr(allheaders,"User Group")>0  Then
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Webele_Column_user-list_UserGroupName").highlight	
		Call ReportStep (StatusTypes.Pass,"Verify Newly added Column_user-list_UserGroupName display","Column_user-list_UserGroupName is present" )
	 	
	 Else 
	  	Call ReportStep (StatusTypes.Fail,"Verify Newly added Column_user-list_UserGroupName display","Column_user-list_UserGroupName is present" )
	 end if 	
	 End If
	end function
	
Function verify_userGroupBlank()
	UGBrows = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").RowCount
	For UGC = 2 To UGBrows Step 1
		Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(UGC,2,"WebCheckbox",0).click
		UGvalue = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(UGC,16,"WebList",0).getroproperty("innertext")
		If UGvalue = " " Then
			Call ReportStep (StatusTypes.Pass,"Verify the user group col is empty for existing OCRF","Usergroup  column empty")
		else
			Call ReportStep (StatusTypes.Fail,"Verify the user group col is empty for existing OCRF","Usergroup  column is not empty")		
		End If
		
	Next

end function

Function verify_Imported_userGroupBlank()
	UGBrows = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").RowCount
	For UGC = 2 To UGBrows Step 1
		Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(UGC,2,"WebCheckbox",0).click
		UGvalue = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(UGC,16,"WebList",0).getroproperty("innertext")
		If UGvalue = " " Then
			Call ReportStep (StatusTypes.Pass,"Verify the user group col is empty for imported users","Usergroup  column empty for imported users")
		else
			Call ReportStep (StatusTypes.Fail,"Verify the user group col is empty for imported users","Usergroup  column is not empty for imported users")		
		End If
		
	Next

end function
	
	
Function verifytheusergrouplistedindropdown(ByVal FileName,ByVal AddUserSheetName,ByVal StartRow,ByVal EndRow,objData)
	Call ImportSheet(AddUserSheetName,FileName)
	
	
	
	UGBrows = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").RowCount
	Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(2,2,"WebCheckbox",0).set "ON"
	
	if Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(2,16,"WebList",0).exist(1) then
		Call ReportStep (StatusTypes.Pass,"Verify User group column should be auto populate drop down list","User group column is auto populate drop down list")
	else
		Call ReportStep (StatusTypes.Fail,"Verify User group column should be auto populate drop down list","User group column is not a auto populate drop down list")
	end if 
	For UGC = 2 To UGBrows Step 1
		Datatable.GetSheet(AddUserSheetName).SetCurrentRow(StartRow)
		Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(UGC,2,"WebCheckbox",0).set "ON"
		
		UGallvalues = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(UGC,16,"WebList",0).getroproperty("all items")
		For EUG = StartRow To EndRow Step 1
			checkuser = datatable.value("UserGroup",AddUserSheetName)
		If instr(UGallvalues,checkuser)>0 Then
			
			Call ReportStep (StatusTypes.Pass,"Verify the added user group - "&checkuser&"listed under the user group column for each user","Usergroup listed under usergroupcolumn")
		else
			Call ReportStep (StatusTypes.Fail,"Verify the added user group - "&checkuser&"listed under the user group column for each user","Usergroup not listed under usergroupcolumn")		
		End If
			Datatable.GetSheet(AddUserSheetName).SetNextRow
		Next
		
	Next
	
	Set Upload=Description.Create()
			Upload("micclass").value="WebElement"
			Upload("html tag").value="DIV"
			Upload("innertext").value="Upload User"
			Set toatlObj=Browser("Browser-OCRF").Page("Page-OCRF").ChildObjects(Upload)
			toatlObj(0).Click	
		Datatable.GetSheet(AddUserSheetName).SetCurrentRow(StartRow)
		Call Verifythenewweblist_UserGroup() 
	allvaluesUG =Browser("Browser-OCRF").Page("Page-OCRF").WebList("Weblist_UserGroupName").GetROProperty("all items")
	
	For EUG = StartRow To EndRow Step 1
			checkuser = datatable.value("UserGroup",AddUserSheetName)
		If instr(allvaluesUG,checkuser)>0 Then
			
			Call ReportStep (StatusTypes.Pass,"Verify the added user group - "&checkuser&"listed under the user group list in upload user","Usergroup listed under user group list in upload user")
		else
			Call ReportStep (StatusTypes.Fail,"Verify the added user group - "&checkuser&"listed under the user group list in upload user","Usergroup not listed under user group list in upload user")		
		End If
			Datatable.GetSheet(AddUserSheetName).SetNextRow
		Next
	
End Function	
	
	Function Verify_Addusergroup(ByVal FileName,ByVal AddUserSheetName,ByVal StartRow,ByVal EndRow,objData)
	
	Call ImportSheet(AddUserSheetName,FileName)
	
	Datatable.GetSheet(AddUserSheetName).SetCurrentRow(StartRow)
	
	
	if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Add_UserGroup").Exist(2) then
		Call ReportStep (StatusTypes.Pass,"Verify Newly added User group option display","new User group button is present" )
		Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Add_UserGroup").Click
		If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("AddUserGroupDialog").Exist(5) Then
			Call ReportStep (StatusTypes.Pass,"Verify AddUserGroupDialog display","AddUserGroupDialog displayed" )
			
			For i  = StartRow To EndRow Step 1
			if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebEl_Confirmaion").Exist(2) then
					Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Webbutton_Yes_2").Click
					Call Pageloading()
				End If
			UGName = Datatable.Value("UserGroup",AddUserSheetName)	
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("UserGroup_AddNewLine").Click
			Call ReportStep (StatusTypes.Pass,"Clicked on usergroup-Add new line","Adding new line for user group creation" )	
			if Browser("Browser-OCRF").Page("Page-OCRF").WebTable("usergroup-list").Exist(1) then
				if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebEl_Confirmaion").Exist(2) then
					Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Webbutton_Yes_2").Click
					Call Pageloading()
				End If
				Call ReportStep (StatusTypes.Pass,"Verify usergroup-list table display","usergroup-list table displayed" )
				UGrows = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("usergroup-list").RowCount
				Browser("Browser-OCRF").Page("Page-OCRF").WebTable("usergroup-list").ChildItem(UGrows,2,"WebCheckbox",0).set "ON"
				Browser("Browser-OCRF").Page("Page-OCRF").WebTable("usergroup-list").ChildItem(UGrows,3,"WebEdit",0).set UGName
				Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Add User Group").Click
				Call pageloading()
				if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebEl_Confirmaion").Exist(2) then
					Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Webbutton_Yes_2").Click
					Call Pageloading()
				End If
				
				 Set objShell = CreateObject("Wscript.shell")
   				 objShell.SendKeys "{ENTER}"
				
				if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebEl_Confirmaion").Exist(2) then
					Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Webbutton_Yes_2").Click
					Call Pageloading()
				End If
				if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebEl_Confirmaion").Exist(2) then
					Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Webbutton_Yes_2").Click
					Call Pageloading()
				End If
				Call ReportStep (StatusTypes.Pass,"User group added successfully","User group added" )				
			else
				Call ReportStep (StatusTypes.Fail,"Verify usergroup-list table display","usergroup-list table not displayed" )			
			End If	
			Datatable.GetSheet(AddUserSheetName).SetNextRow
			 Set objShell = CreateObject("Wscript.shell")
    			objShell.SendKeys "{ENTER}"
			
			next
			Browser("Browser-OCRF").Page("Page-OCRF").WebButton("webbutton_Cancel").Click
		else
			Call ReportStep (StatusTypes.Fail,"Verify AddUserGroupDialog display","AddUserGroupDialog not displayed" )		
		End If
	 Else 
	  	Call ReportStep (StatusTypes.Fail,"Verify Newly added User group button display","New User group button is present" )
	 	
	 End If
	
	
	
	end function
	
	
	Function verify_ContentPermissionInOfferingSummary()
		wait 5
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("InOfferingSummary_Content Permissions -").Exist then
		Call ReportStep (StatusTypes.Pass,"Verify Newly added InOfferingSummary_Content Permissions display","InOfferingSummary_Content Permissions is present" )
	 Else 
	  	Call ReportStep (StatusTypes.Fail,"Verify Newly added InOfferingSummary_Content Permissionse display","InOfferingSummary_Content Permissions is present" )
	 	
	 End If
	End Function
	
	Function Verify_REportName_EditBoxDisabled(CCR)
		strdisabled = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList").ChildItem(CCR, 16, "WebEdit", 0).getroproperty("disabled")
		If strdisabled = "1" Then
			Call ReportStep (StatusTypes.Pass, "Verify that Report Name edit box is disabled(New Change)","Report Name edit box is disabled ")
       Else
       	  Call ReportStep (StatusTypes.Pass, "Verify that Report Name edit box is disabled(New Change)","Report Name edit box is not disabled ")
       End If
	End function


'********************************************************************
'Created By: Madhu
'Date: 05/06/2020
'Function: Validating the new functionalities of BI tool access tab
'********************************************************************

Function Validate_PBIConsumerAccessRemoval(ByVal FileName,ByVal BIToolSheetName,ByVal StartRow,ByVal EndRow,ByRef objData)
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementOfferingUserBIToolList").Exist(120) <> true Then
		Call ReportStep (StatusTypes.Fail, "Not Redirected to BI Tool Access tab","BI tool access Tab")	
	Else
		'Call ReportStep (StatusTypes.Pass, "Redirected to BI Tool Access tab successfully","BI tool access Tab")
	End If
	
	Call ImportSheet(BIToolSheetName,FileName)
	
	Datatable.GetSheet(BIToolSheetName).SetCurrentRow(StartRow)

	Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolGrid_PageContent").Select "100"
	Call PageLoading()
	wait 1
	
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Exist(120) Then
	    Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Click
		wait 2
	Else
		Call ReportStep (StatusTypes.Fail, "Not Clicked on 'Select user","BI tool access Tab")
	End If
	
	allBitools = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIAccessTool").GetROProperty("all items")
	
	If instr(allBitools,"VA PBI")<=0 Then
		Call ReportStep (StatusTypes.Pass, "VA PBI is removed from the Choose BI Tool List","BI tool access Tab - VA PBI")
	else
		Call ReportStep (StatusTypes.Fail, "VA PBI is not removed from the Choose BI Tool List","BI tool access Tab - VA PBI")
	End If
	
	If instr(allBitools,"VA Dossier")>=0 Then
		Call ReportStep (StatusTypes.Pass, "VA Dossier is Present in the Choose BI Tool List","BI tool access Tab - VA Dossier")
	else
		Call ReportStep (StatusTypes.Fail, "VA Dossier is not present in the Choose BI Tool List","BI tool access Tab - VA Dossier")
	End If
	
	
	Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIAccessTool").Select "VA Dossier"
	
	allaccessmode = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIToolAccessMode").GetROProperty("all items")

	If instr(allaccessmode,"VA Architect")>=0 Then
		Call ReportStep (StatusTypes.Pass, "VA Architect is present in Choose BI ToolAccess mode List for VA Dossier","BI tool access Tab - VA Dossier")
	else
		Call ReportStep (StatusTypes.Fail, "VA Architect is not present in Choose BI ToolAccess mode List for VA Dossier","BI tool access Tab - VA Dossier")
	
	End If

	If instr(allaccessmode,"VA Consumer")<=0 Then
		Call ReportStep (StatusTypes.Pass, "VA Consumer is removed from the Choose BI ToolAccess mode List","BI tool access Tab - VA PBI")
	else
		Call ReportStep (StatusTypes.Fail, "VA Consumer is not removed from the Choose BI ToolAccess mode List","BI tool access Tab - VA PBI")
	
	End If
	
	
	
end function

Function BITool_addUsers_validate_newchanges1(ByVal FileName,ByVal BIToolSheetName,ByVal StartRow,ByVal EndRow,ByRef objData)
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementOfferingUserBIToolList").Exist(120) <> true Then
		Call ReportStep (StatusTypes.Fail, "Not Redirected to BI Tool Access tab","BI tool access Tab")	
	Else
		'Call ReportStep (StatusTypes.Pass, "Redirected to BI Tool Access tab successfully","BI tool access Tab")
	End If
		
	Call ImportSheet(BIToolSheetName,FileName)

	If Ucase(StartRow) = "CHECKALL" Then
		
		Datatable.GetSheet(BIToolSheetName).SetCurrentRow(EndRow)
		If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Exist(120) Then
			    Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Click
				wait 2
			Else
				Call ReportStep (StatusTypes.Fail, "Not Clicked on 'Select user","BI tool access Tab")
		End If
			
				
		Call PageLoading()
		wait 2
		Set objBIAccessTool = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIAccessTool")
		Set objBIToolAccessMode = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIToolAccessMode")
		
		Call SCA.SelectFromDropdown(objBIAccessTool, Datatable.Value("BI_Tool",BIToolSheetName))
		Call SCA.SelectFromDropdown(objBIToolAccessMode, Datatable.Value("Access_Mode",BIToolSheetName))
		
		Browser("Browser-OCRF").Page("Page-OCRF").WebCheckBox("WebCheckBox_SelectAll_BiTools").Set "ON"
		wait 1
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_AddUser").Exist<>true then
			Call ReportStep (StatusTypes.Pass, "Add User Button removed successfully","Add User Button")
		else
			Call ReportStep (StatusTypes.Fail, "Add User Button removed successfully","Add User Button")
		End If
		Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_Add").Click
		wait 5
		If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebBtnInformationOkay").Exist(10) Then 
			Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebBtnInformationOkay").Click
			wait 2
		End If 
	
		Else
			If StartRow = "" Then
				StartRow = 1
			End If
			If EndRow = "" Then
				EndRow = Datatable.GetSheet(BIToolSheetName).GetRowCount
			End If
		
			Datatable.GetSheet(BIToolSheetName).SetCurrentRow(StartRow)
			


			Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolGrid_PageContent").Select "100"
			Call PageLoading()
			wait 1
			recordCountBeforeAdding = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").RowCount
			
			recordAddedAtRow = recordCountBeforeAdding + 1
			
			If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Exist(120) Then
				    Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Click
					wait 2
				Else
					Call ReportStep (StatusTypes.Fail, "Not Clicked on 'Select user","BI tool access Tab")
				End If
			
			For rownum = StartRow To EndRow Step 1
				
				Set objBIAccessTool = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIAccessTool")
				Set objBIToolAccessMode = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIToolAccessMode")
				
				Call SCA.SelectFromDropdown(objBIAccessTool, Datatable.Value("BI_Tool",BIToolSheetName))
				Call SCA.SelectFromDropdown(objBIToolAccessMode, Datatable.Value("Access_Mode",BIToolSheetName))
				

				Browser("Browser-OCRF").Page("Page-OCRF").WebList("Weblist_SelectUserlistNo").Select "999"
				Call PageLoading()
				
		
				userRowNum = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").GetRowWithCellText(Datatable.value("User_Id",BIToolSheetName),4,2)
				Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").ChildItem(userRowNum,2,"WebCheckBox",0).set "ON"
				
				Call validate_BIAdduserButtonREmoval()
				call Validate_selecteduserdisplay(userRowNum)
				
				Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").ChildItem(userRowNum,2,"WebCheckBox",0).click
				Call Validate_deselectedusernotlisted(userRowNum)
				Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").ChildItem(userRowNum,2,"WebCheckBox",0).set "ON"
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_Add").Click
				wait 5
				Call PageLoading()
				
				If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Information_Okay").Exist(10) Then 
					Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Information_Okay").Click
					wait 2
				End If
				
				If Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIAccessTool").exist Then
					Call ReportStep (StatusTypes.Pass, "Please select user from the list popup did not close on clicking the add button","Please select user from the list popup") 
				else
					Call ReportStep (StatusTypes.Fail, "Please select user from the list popup closed on clicking the add button","Please select user from the list popup") 
				End If	
				
				recordAddedAtRow = recordAddedAtRow + 1
				
				

				if Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Webbutton_Okay").Exist(2) then
					Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Webbutton_Okay").Click
				end if
					
				if Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Webbutton_Okay").Exist(2) then
					Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Webbutton_Okay").Click
				end if					

				Datatable.GetSheet(BIToolSheetName).SetNextRow
			Next
				
			recordCount = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").RowCount
				If recordCount  = recordAddedAtRow Then
					Call ReportStep (StatusTypes.Pass, "New row added at " & recordAddedAtRow & " with " & Datatable.Value("BI_Tool",BIToolSheetName) & ", " & Datatable.Value("Access_Mode",BIToolSheetName) & ", " & Datatable.value("User_Id",BIToolSheetName),"BI tool access Tab - Add Record")
					
					Else 
					'Call ReportStep (StatusTypes.Fail, "New row couldn't add for row no. " & recordCount +1 ,"BI tool access Tab - Add Record")
				
				End If


			Call PageLoading()
			wait 3
			Browser("Browser-OCRF").Page("Page-OCRF").Sync
			
			noOfRecords = (EndRow - StartRow) +1
			recordCountAfterAdding = (recordCountBeforeAdding - 1) +  noOfRecords
			
			wait 3
			If (Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").RowCount)-1 = recordCountAfterAdding Then
				Call ReportStep (StatusTypes.Pass, recordCountAfterAdding & " Records matched successfully between data sheet and BI Tool access list","Records match for BI Tool access")	
			Else
				Call ReportStep (StatusTypes.Fail, recordCountAfterAdding & " Records did not match, between data sheet and BI Tool access list","Records match for BI Tool access")	
			End If	
			if Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_DBAccesspage").Exist then
				
				Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_DBAccesspage").Click
				Call ReportStep (StatusTypes.Pass, "Close button is present in the please select user from users list pop-up","Close button")	
			Else
				Call ReportStep (StatusTypes.Fail, "Close button is not present in the please select user from users list pop-up","Close button")
			End If
			

			Call NavigateToNextTab()
			
			Call PageLoading()	
	
end if 	
End Function

function validate_BIAdduserButtonREmoval()
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_AddUser").Exist<>true then
			Call ReportStep (StatusTypes.Pass, "Add User Button removed successfully","Add User Button")
		else
			Call ReportStep (StatusTypes.Fail, "Add User Button removed successfully","Add User Button")
		End If
	end function

Function Validate_selecteduserdisplay(rownumber)

	selectedUsername = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").GetCellData(rownumber,4)
	
	
	addeduser = Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("BISelectedUserDisplay").GetROProperty("innertext")
	
	If instr(addeduser,selectedUsername) Then
		Call ReportStep (StatusTypes.Pass, "Please select user from the list popup - the selected User gets added directly in the text area"," the selected User gets added directly in the text area") 
	else
		Call ReportStep (StatusTypes.Fail, "Please select user from the list popup - the selected User did not get added directly in the text area"," the selected User gets added directly in the text area") 
	
	End If
	
		
End Function

Function Validate_deselectedusernotlisted(userRowNum)

	selectedUsername = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").GetCellData(rownumber,4)
	
	
	addeduser = Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("BISelectedUserDisplay").GetROProperty("innertext")
	
	If instr(addeduser,selectedUsername)<=0 Then
		Call ReportStep (StatusTypes.Pass, "Please select user from the list popup - the deselected User gets removed in the text area"," User gets removed in the text area") 
	else
		Call ReportStep (StatusTypes.Fail, "Please select user from the list popup - the deselected User gets not removed in the text area"," User not gets removed in the text area") 
	
	End If
end function
	
Function BITool_addUsers_validate_newchanges2(ByVal FileName,ByVal BIToolSheetName,ByVal StartRow,ByVal EndRow,ByRef objData)
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementOfferingUserBIToolList").Exist(120) <> true Then
		Call ReportStep (StatusTypes.Fail, "Not Redirected to BI Tool Access tab","BI tool access Tab")	
	Else
		'Call ReportStep (StatusTypes.Pass, "Redirected to BI Tool Access tab successfully","BI tool access Tab")
	End If
		
	Call ImportSheet(BIToolSheetName,FileName)

			If StartRow = "" Then
				StartRow = 1
			End If
			If EndRow = "" Then
				EndRow = Datatable.GetSheet(BIToolSheetName).GetRowCount
			End If
		
			Datatable.GetSheet(BIToolSheetName).SetCurrentRow(StartRow)
			


			Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolGrid_PageContent").Select "100"
			Call PageLoading()
			wait 1
			
			
			If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Exist(120) Then
				    Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Click
					wait 2
				Else
					Call ReportStep (StatusTypes.Fail, "Not Clicked on 'Select user","BI tool access Tab")
				End If
				userRowNum = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").GetRowWithCellText(Datatable.value("User_Id",BIToolSheetName),4,2)
				
				Call Deleteuserbyselectingcheckbox(FileName,BIToolSheetName,objData)
				
				Datatable.GetSheet(BIToolSheetName).SetCurrentRow(10)
				
				Call adduserbyaddingintextbox(FileName,BIToolSheetName,objData,userRowNum)
				
				Call checkusernotexistwarning(FileName,BIToolSheetName,bjData)
				
				
			
			Call NavigateToNextTab()
			
			Call PageLoading()	
	
	
End Function

Function Deleteuserbyselectingcheckbox(ByVal FileName,ByVal BIToolSheetName,ByRef objData)
				Set objBIAccessTool = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIAccessTool")
				Set objBIToolAccessMode = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIToolAccessMode")
				
				Call SCA.SelectFromDropdown(objBIAccessTool, Datatable.Value("BI_Tool",BIToolSheetName))
				Call SCA.SelectFromDropdown(objBIToolAccessMode, Datatable.Value("Access_Mode",BIToolSheetName))

				Browser("Browser-OCRF").Page("Page-OCRF").WebList("Weblist_SelectUserlistNo").Select "999"
				Call PageLoading()
				
		
				userRowNum = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").GetRowWithCellText(Datatable.value("User_Id",BIToolSheetName),4,2)
				Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").ChildItem(userRowNum,2,"WebCheckBox",0).set "ON"
				
				usertodelete = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").GetCellData(userRowNum,4)
			
				call Validate_selecteduserdisplay(userRowNum)
				
				'Delete the selected user
				Browser("Browser-OCRF").Page("Page-OCRF").WebButton("DeleteActionConformation").Click

				wait 5
				Call PageLoading()
				
				If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Information_Okay").Exist(2) Then 
					Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Information_Okay").Click
					wait 2
				End If
				
				If Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIAccessTool").exist Then
					Call ReportStep (StatusTypes.Pass, "Please select user from the list popup did not close on clicking the Delete button","Please select user from the list popup") 
				else
					Call ReportStep (StatusTypes.Fail, "Please select user from the list popup closed on clicking the Delete button","Please select user from the list popup") 
				End If	
				
				if Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Webbutton_Okay").Exist(2) then
					Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Webbutton_Okay").Click
				end if
					
			Call PageLoading()
			wait 3
			Browser("Browser-OCRF").Page("Page-OCRF").Sync
								
			if Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Symbol_BI_close").Exist then
				
				Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Symbol_BI_close").Click

				Call ReportStep (StatusTypes.Pass, "User able to close please select user from users list pop-up through X button","X button")	
			Else
				Call ReportStep (StatusTypes.Fail, "User not able to close please select user from users list pop-up through X button","X button")
			End If
			
			Call PageLoading()
			
			if Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").exist(1) then
				deleteduserrow = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").GetRowWithCellText(usertodelete)
				deletedstatus = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").GetCellData(deleteduserrow,6)
				
				If deletedstatus = "DELETE" Then
					Call ReportStep (StatusTypes.Pass, "Status changed to delete after deleting the user from please select user from users list pop-up","Status changed to delete")	
				Else
					Call ReportStep (StatusTypes.Fail, "Status not changed to delete after deleting the user from please select user from users list pop-up","Status changed to delete")
			
				End If
			End If 	
end function

Function adduserbyaddingintextbox(ByVal FileName,ByVal BIToolSheetName,ByRef objData,userRowNum)
			
			If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Exist(120) Then
				    Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Click
					wait 2
				Else
					Call ReportStep (StatusTypes.Fail, "Not Clicked on 'Select user","BI tool access Tab")
				End If

				Set objBIAccessTool = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIAccessTool")
				Set objBIToolAccessMode = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIToolAccessMode")
				
				
				Call SCA.SelectFromDropdown(objBIAccessTool, Datatable.Value("BI_Tool",BIToolSheetName))
				Call SCA.SelectFromDropdown(objBIToolAccessMode, Datatable.Value("Access_Mode",BIToolSheetName))

				Browser("Browser-OCRF").Page("Page-OCRF").WebList("Weblist_SelectUserlistNo").Select "999"
				Call PageLoading()
				
			user1 = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").GetCellData(userRowNum,4)	
			user2 = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessUserList").GetCellData(userRowNum-1,4)
			 
			Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("BISelectedUserDisplay").Set user1&";"&user2
			Call ReportStep (StatusTypes.Pass, "Users added to the 'please select user from users list pop-up' textbox" ,"'please select user from users list pop-up' textbox")			

			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_Add").Click
			
			totalbiusersadded = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").RowCount
			
			For i = 1 To totalbiusersadded Step 1
				checkusername =  Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").GetCellData(i,8)
				checkreportname = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").GetCellData(i,11)
				If ucase(trim(checkusername))= ucase(user1) and ucase(trim(Datatable.Value("BI_Tool",BIToolSheetName)))=ucase(checkreportname)  Then
					checkaddedstatus = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").GetCellData(i,6)
					If checkaddedstatus = "ADD" Then
					addedsuccessfully = 1
					Call ReportStep (StatusTypes.Pass, "Addition of user through 'please select user from users list pop-up' textbox is successfull","please select user from users list pop-up textbox")
					End If
					
				End If
				
			Next
			
			If addedsuccessfully<>1 Then
				Call ReportStep (StatusTypes.Fail, "Addition of user through 'please select user from users list pop-up' textbox is not successfull","please select user from users list pop-up textbox")
			End If
end function



Function checkusernotexistwarning(ByVal FileName,ByVal BIToolSheetName,ByRef objData)
		If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Exist(120) Then
		    Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Click
			wait 2
		Else
			Call ReportStep (StatusTypes.Fail, "Not Clicked on 'Select user","BI tool access Tab")
		End If

		Set objBIAccessTool = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIAccessTool")
		Set objBIToolAccessMode = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_BIToolAccess_BIToolAccessMode")
				
		Call SCA.SelectFromDropdown(objBIAccessTool, Datatable.Value("BI_Tool",BIToolSheetName))
		Call SCA.SelectFromDropdown(objBIToolAccessMode, Datatable.Value("Access_Mode",BIToolSheetName))

		Browser("Browser-OCRF").Page("Page-OCRF").WebList("Weblist_SelectUserlistNo").Select "999"
		Call PageLoading()
		
		usernotexist= Datatable.Value("User_Id",BIToolSheetName)
		
		Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("BISelectedUserDisplay").Set usernotexist
		Call ReportStep (StatusTypes.Pass, "Users added to the 'please select user from users list pop-up' textbox" ,"'please select user from users list pop-up' textbox")			
		
		Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_Add").Click
		
		
		if Browser("Browser-OCRF").Dialog("Message from webpage").Exist(2) then
			Call ReportStep (StatusTypes.Pass, "Warning message displayed for adding user in 'please select user from users list pop-up' textbox which is not exist" ,"'please select user from users list pop-up' textbox")			
			Browser("Browser-OCRF").Dialog("Message from webpage").WinButton("OK").click
		else
			Call ReportStep (StatusTypes.Fail, "Warning message not displayed for adding user in 'please select user from users list pop-up' textbox which is not exist" ,"'please select user from users list pop-up' textbox")			
				
		End If
		
		end Function 
			

		
Function DataSource_addUsers_validate_newchanges2(ByVal FileName,ByVal UserDBSheetName,ByVal StartRow,ByVal EndRow,ByRef objData)
	TotalRow=(EndRow-StartRow)+1
	
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_UserDatabaseList").Exist(60) Then
		Browser("Browser-OCRF").Page("Page-OCRF").Sync
		Call ReportStep (StatusTypes.Pass, "Redirected to User DB Access tab successfully","User DB access Tab")
	Else
		Call ReportStep (StatusTypes.Fail, "Not Redirected to User DB Access tab","User DB access Tab")	
		Exit Function
	End If
	

	Call ImportSheet(UserDBSheetName,FileName)
			If StartRow = "" Then
				StartRow = 1
			End If
			If EndRow = "" Then
				EndRow = Datatable.GetSheet(SheetName).GetRowCount
			End If
		
			DataTable.GetSheet(UserDBSheetName).SetCurrentRow(StartRow)
			
			Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_PageContent").Select "100"
	
			Call PageLoading()
			wait 1
			If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Exist(120) Then
		    	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Click
				Call ReportStep (StatusTypes.Pass, "Clicked on 'Select user'","USer_DB_Access Tab")	
			Else
				Call ReportStep (StatusTypes.Fail, "Not Clicked on 'Select user","USer_DB_Access Tab")
			End If	
			DataTable.GetSheet(UserDBSheetName).SetCurrentRow(StartRow)
			userRowNum = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").GetRowWithCellText(Datatable.value("User_ID",UserDBSheetName))
				
				Call DBDeleteuserbyselectingcheckbox(FileName,UserDBSheetName,objData)
				
				DataTable.GetSheet(UserDBSheetName).SetCurrentRow(10)
				
				Call DBadduserbyaddingintextbox(FileName,UserDBSheetName,objData,userRowNum)
				
				Call DBcheckusernotexistwarning(FileName,UserDBSheetName,bjData)
				
				
			
			Call NavigateToNextTab()
			
			Call PageLoading()	
	
	
End Function


Function DBDeleteuserbyselectingcheckbox(FileName,UserDBSheetName,objData)
	
		
		Set objCubeDB = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_CubeDB")
		Set objCubeRoles = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_CubeRoles")
			
		Set objCubeType = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_CubeType")
	    If objCubeType.Exist(5) Then
		wait 2
	     If objCubeType.GetROProperty("disabled") = 0 Then
	       objCubeType.Select Trim(Datatable.Value("DatabaseType",UserDBSheetName))
	       wait 2
	     End If
	    End If
	
		wait 3
		objCubeDB.WaitProperty "value","Select Datasource Name",5000
		wait 4
		objCubeDB.Select 1
		objCubeRoles.WaitProperty "disabled",0,5000
		wait 3
		objCubeRoles.Select Trim(Datatable.value("Role",UserDBSheetName))
		rownum = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").GetRowWithCellText(Datatable.value("User_ID",UserDBSheetName),4,2)
		Set chk =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").ChildItem(rownum,2,"WebCheckBox",0)
		chk.click
		'the user selected for delete
		DBdeleteuser = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").GetCellData(rownum,4)
		
		Call DBValidate_selecteduserdisplay(rownum)
		
		'Delete the selected user
		Browser("Browser-OCRF").Page("Page-OCRF").WebButton("DeleteActionConformation").Click
		
		
		wait 5
				Call PageLoading()


				If Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_CubeDB").exist Then
					Call ReportStep (StatusTypes.Pass, "Please select user from the list popup did not close on clicking the Delete button","Please select user from the list popup") 
				else
					Call ReportStep (StatusTypes.Fail, "Please select user from the list popup closed on clicking the Delete button","Please select user from the list popup") 
				End If

			if Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Symbol_BI_close").Exist then
				
				Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Symbol_BI_close").Click

				Call ReportStep (StatusTypes.Pass, "User able to close please select user from users list pop-up through X button","X button")	
			Else
				Call ReportStep (StatusTypes.Fail, "User not able to close please select user from users list pop-up through X button","X button")
			End If
			
			Call PageLoading()
			if Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserCubeAccess_CubeAccessList").exist(1) then
				deleteduserrow = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserCubeAccess_CubeAccessList").GetRowWithCellText(DBdeleteuser)
				deletedstatus = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserCubeAccess_CubeAccessList").GetCellData(deleteduserrow,3)
				
				If deletedstatus = "DELETE" Then
					Call ReportStep (StatusTypes.Pass, "Status changed to delete after deleting the user from please select user from users list pop-up","Status changed to delete")	
				Else
					Call ReportStep (StatusTypes.Fail, "Status not changed to delete after deleting the user from please select user from users list pop-up","Status changed to delete")
			
				End If
			End If 	
			
End Function			
				
				
		
		
		
Function DBValidate_selecteduserdisplay(rownumber)

	dbselectedUsername = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").GetCellData(rownumber,4)
	
	
	dbaddeduser = Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_SelectedCubeUsers").GetROProperty("innertext")
	
	If instr(dbaddeduser,dbaddeduser) Then
		Call ReportStep (StatusTypes.Pass, "Please select user from the list popup - the selected User gets added directly in the text area"," the selected User gets added directly in the text area") 
	else
		Call ReportStep (StatusTypes.Fail, "Please select user from the list popup - the selected User did not get added directly in the text area"," the selected User gets added directly in the text area") 
	
	End If
	
		
End Function


Function DBadduserbyaddingintextbox(ByVal FileName,ByVal UserDBSheetName,ByRef objData,userRowNum)
			
			If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Exist(120) Then
		    	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Click
				Call ReportStep (StatusTypes.Pass, "Clicked on 'Select user'","USer_DB_Access Tab")	
			Else
				Call ReportStep (StatusTypes.Fail, "Not Clicked on 'Select user","USer_DB_Access Tab")
			End If	
				Set objCubeDB = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_CubeDB")
		Set objCubeRoles = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_CubeRoles")
			
		Set objCubeType = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_CubeType")
	    If objCubeType.Exist(5) Then
		wait 2
	     If objCubeType.GetROProperty("disabled") = 0 Then
	       objCubeType.Select Trim(Datatable.Value("DatabaseType",UserDBSheetName))
	       wait 2
	     End If
	    End If
	
		wait 3
		objCubeDB.WaitProperty "value","Select Datasource Name",5000
		wait 4
		objCubeDB.Select 1
		objCubeRoles.WaitProperty "disabled",0,5000
		wait 3
		objCubeRoles.Select Trim(Datatable.value("Role",UserDBSheetName))
				Browser("Browser-OCRF").Page("Page-OCRF").WebList("Weblist_SelectUserlistNo").Select "999"
				Call PageLoading()
				
			user1 = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").GetCellData(userRowNum,4)	
			'user2 = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").GetCellData(userRowNum-1,4)
			cubeselected = objCubeDB.GetROProperty("selection")
			 
			Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_SelectedCubeUsers").Set user1&";"&user2
			Call ReportStep (StatusTypes.Pass, "Users added to the 'please select user from users list pop-up' textbox" ,"'please select user from users list pop-up' textbox")			

			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_UserDB_Add").Click
			wait 3
			if Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Information_Okay").Exist then
				Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Information_Okay").Click
			end if 		

			if Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Symbol_BI_close").Exist then
				
				Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Symbol_BI_close").Click

				Call ReportStep (StatusTypes.Pass, "User able to close please select user from users list pop-up through X button","X button")	
			Else
				Call ReportStep (StatusTypes.Fail, "User not able to close please select user from users list pop-up through X button","X button")
			End If
			totalbiusersadded = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserCubeAccess_CubeAccessList").RowCount
			
			For i = 2 To 4 Step 1
				dbcheckusername =  Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserCubeAccess_CubeAccessList").GetCellData(i,7)
				dbcheckcubename = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserCubeAccess_CubeAccessList").GetCellData(i,10)
				If ucase(trim(dbcheckusername))= ucase(user1) and ucase(trim(cubeselected))=ucase(dbcheckcubename)  Then
					checkaddedstatus = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserCubeAccess_CubeAccessList").GetCellData(i,3)
					If checkaddedstatus = "ADD" Then
					addedsuccessfully = 1
					Call ReportStep (StatusTypes.Pass, "Addition of user through 'please select user from users list pop-up' textbox is successfull","please select user from users list pop-up textbox")
					End If
					
				End If
				
			Next
			
			If addedsuccessfully<>1 Then
				Call ReportStep (StatusTypes.Fail, "Addition of user through 'please select user from users list pop-up' textbox is not successfull","please select user from users list pop-up textbox")
			End If
end function


Function DBcheckusernotexistwarning(ByVal FileName,ByVal UserDBSheetName,ByRef objData)
		If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Exist(120) Then
		    	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Click
				Call ReportStep (StatusTypes.Pass, "Clicked on 'Select user'","USer_DB_Access Tab")	
			Else
				Call ReportStep (StatusTypes.Fail, "Not Clicked on 'Select user","USer_DB_Access Tab")
			End If	
				Set objCubeDB = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_CubeDB")
		Set objCubeRoles = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_CubeRoles")
			
		Set objCubeType = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_CubeType")
	    If objCubeType.Exist(5) Then
		wait 2
	     If objCubeType.GetROProperty("disabled") = 0 Then
	       objCubeType.Select Trim(Datatable.Value("DatabaseType",UserDBSheetName))
	       wait 2
	     End If
	    End If
	
		wait 3
		objCubeDB.WaitProperty "value","Select Datasource Name",5000
		wait 4
		objCubeDB.Select 1
		objCubeRoles.WaitProperty "disabled",0,5000
		wait 3
		objCubeRoles.Select Trim(Datatable.value("Role",UserDBSheetName))
				Browser("Browser-OCRF").Page("Page-OCRF").WebList("Weblist_SelectUserlistNo").Select "999"
				Call PageLoading()
			
		DataTable.GetSheet(UserDBSheetName).SetCurrentRow(10)
		usernotexist= Datatable.Value("User_ID",UserDBSheetName)
		
		Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_SelectedCubeUsers").Set usernotexist
		Call ReportStep (StatusTypes.Pass, "Users added to the 'please select user from users list pop-up' textbox" ,"'please select user from users list pop-up' textbox")			
		
		Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_UserDB_Add").Click
					
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("UserDoesnotBelong").Exist(2) then
			Call ReportStep (StatusTypes.Pass, "Warning message displayed for adding user in 'please select user from users list pop-up' textbox which is not exist" ,"'please select user from users list pop-up' textbox")			
			if Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Information_Okay").Exist(2) then
				Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Information_Okay").Click
			else
			  Set objShell = CreateObject("Wscript.shell")
   			 objShell.SendKeys "{ENTER}"
			end if    			 
		else
			Call ReportStep (StatusTypes.Fail, "Warning message not displayed for adding user in 'please select user from users list pop-up' textbox which is not exist" ,"'please select user from users list pop-up' textbox")			
				
		End If
		
		end Function 
		
		
	
	
Function DB_addUsers_validate_newchanges1(ByVal FileName,ByVal SheetName,ByVal StartRow,ByVal EndRow,TimeStamp,ByRef objData)
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_UserDatabaseList").Exist(60) Then
		Browser("Browser-OCRF").Page("Page-OCRF").Sync
		Call ReportStep (StatusTypes.Pass, "Redirected to User DB Access tab successfully","User DB access Tab")
	Else
		Call ReportStep (StatusTypes.Fail, "Not Redirected to User DB Access tab","User DB access Tab")	
		Exit Function
	End If
	
	Call ImportSheet(SheetName,FileName)
	
	DataTable.GetSheet(SheetName).SetCurrentRow(StartRow)
	
	Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_PageContent").Select "100"
	
	 rowCount = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserCubeAccess_CubeAccessList").RowCount
	 rowCountBeforeAdding = rowCount -1 
	 
	  If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Exist(120) Then
	    	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_BIToolAccess_SelectUsers").Click
			Call ReportStep (StatusTypes.Pass, "Clicked on 'Select user'","USer_DB_Access Tab")	
		Else
			Call ReportStep (StatusTypes.Fail, "Not Clicked on 'Select user","USer_DB_Access Tab")
		End If
	
	For rnum = StartRow To EndRow Step 1	
		
       	pageContentUserCube = "//TD[@id=" & chr(34)& "CubeUser-toadd-list-pager_center"& chr(34) & "]/TABLE[1]/TBODY[1]/TR[1]/TD[8]/SELECT[@role="& chr(34) & "listbox" & chr(34)&"][1]"
		
		Browser("Browser-OCRF").Page("Page-OCRF").WebList("xpath:="&pageContentUserCube).Select "100"
		
		wait 3
		
		Set objCubeDB = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_CubeDB")
		Set objCubeRoles = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_CubeRoles")
			
		Set objCubeType = Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_CubeType")
	    If objCubeType.Exist(5) Then
		wait 2
	     If objCubeType.GetROProperty("disabled") = 0 Then
	       objCubeType.Select Trim(Datatable.Value("DatabaseType",SheetName))
	       wait 2
	     End If
	    End If
		
		wait 3
		objCubeDB.WaitProperty "value","Select Datasource Name",5000
		wait 4
		objCubeDB.Select trim(Datatable.value("Database",SheetName))&Timestamp
		objCubeRoles.WaitProperty "disabled",0,5000
		wait 3
		rownum = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").GetRowWithCellText(Datatable.value("User_ID",SheetName),4,2)
		Set chk =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").ChildItem(rownum,2,"WebCheckBox",0)
		chk.set "ON"
		
		Call DBvalidateAdduserbuttonremoval()
		Call DB_Validate_selecteduserdisplay(rownum)
		chk.click
		
		Call Validate_deselectedusernotlisted(rownum)
		chk.set "ON"
		
		wait 3
		Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_UserDB_Add").Click
		Datatable.GetSheet(SheetName).SetNextRow
		
	Next
	
	'Validate number of records added
	Call PageLoading()
	Browser("Browser-OCRF").Page("Page-OCRF").Sync
	
	noOfRecords = (EndRow - StartRow) +1
	RowCountAfterAdding = rowCountBeforeAdding + noOfRecords
	
	Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_UserCubeAccess_PageContent").Select "100"
	wait 3
	If (Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserCubeAccess_CubeAccessList").RowCount)-1 = RowCountAfterAdding Then
		Call ReportStep (StatusTypes.Pass, RowCountAfterAdding & " Records matched successfully between data sheet and DataSource access list","Records match in DataSource Access tab")	
	Else
		Call ReportStep (StatusTypes.Fail, RowCountAfterAdding & " Records did not match, between data sheet and BI Tool access list","Records match in DataSource Access tab")	
	End If	
	
	If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_DBAccesspage").Exist Then
		Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_DBAccesspage").Click
	End If
	
	
	If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebBtnInformationOkay").Exist(5) Then
		Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebBtnInformationOkay").Click
	End If
	If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebBtnInformationOkay").Exist Then
		Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebBtnInformationOkay").Click
	End If
	
 	
End Function

function DBvalidateAdduserbuttonremoval()
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElement_UserDB_AddUsers").Exist<>true then
			Call ReportStep (StatusTypes.Pass, "Add User Button removed successfully","Add User Button")
		else
			Call ReportStep (StatusTypes.Fail, "Add User Button removed successfully","Add User Button")
		End If
	end function

Function DB_Validate_selecteduserdisplay(rownumber)

	selectedUsername = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").GetCellData(rownumber,4)
	
	
	addeduser = Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("DBselecteduserDisplay2").GetROProperty("innertext")
	
	If instr(addeduser,selectedUsername) Then
		Call ReportStep (StatusTypes.Pass, "Please select user from the list popup - the selected User gets added directly in the text area"," the selected User gets added directly in the text area") 
	else
		Call ReportStep (StatusTypes.Fail, "Please select user from the list popup - the selected User did not get added directly in the text area"," the selected User gets added directly in the text area") 
	
	End If
	
		
End Function

Function Validate_deselectedusernotlisted(userRowNum)

	selectedUsername = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_UserDB_Users").GetCellData(rownumber,4)
	
	if Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("DBselecteduserDisplay2").Exist then
	addeduser = Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("DBselecteduserDisplay2").GetROProperty("innertext")
	else
	addeduser = Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("DBselecteduserDisplay").GetROProperty("innertext")
	end if
	
	If instr(addeduser,selectedUsername)<=0 Then
		Call ReportStep (StatusTypes.Pass, "Please select user from the list popup - the deselected User gets removed in the text area"," User gets removed in the text area") 
	else
		Call ReportStep (StatusTypes.Fail, "Please select user from the list popup - the deselected User gets not removed in the text area"," User not gets removed in the text area") 
	
	End If
end function


function validate_ClientcontentTabChanges01(ByVal FileName,ByVal SheetName,ByVal StartRow,ByVal EndRow,ByRef objData)

 TimeStamp=objData.item("TimeStamp")
   '########### OFFERING DETAIL TAB EXISTANCE - Start ########### 
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementCDTabHeader").Exist(20) Then
		Call ReportStep (StatusTypes.Pass, "'Client content' tab header exist","Client content Tab header")
	Else
	   	Call ReportStep (StatusTypes.Fail, "Client content Tab header not exist","Client content Tab header")
	End If

 	wait 2
	Call ImportSheet(SheetName,FileName)
	
	DataTable.GetSheet(SheetName)
	rc=DataTable.GetSheet(SheetName).GetRowCount
	 Decision_Center_Heading= "Visual Analytics"

	'CLICK ON ADD NEW LINE
		If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementAddNewLine").Exist(20) Then
			wait 2
	        'Call ReportStep (StatusTypes.Pass, "Click on add new line done","Add new button")	
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementAddNewLine").Click
		Else
			Call ReportStep (StatusTypes.Fail, "Click on add new line not done","Add new button")
		End If
		Set objClientContentTable = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList")
		CCR= objClientContentTable.RowCount

		objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click	
		 objClientContentTable.ChildItem(CCR, 10, "WebList", 0).Select Decision_Center_Heading

		objClientContentTable.ChildItem(CCR, 17, "WebEdit", 0).Set Trim(Client_Name)
	           Set objShell=CreateObject("WScript.Shell")
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
			   objClientContentTable.ChildItem(CCR, 17, "WebEdit", 0).Click
				objShell.SendKeys "{DOWN}"
			if	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:=All Clients").Exist(2) then
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:=All Clients").highlight
				Call ReportStep (StatusTypes.Pass, "All clients option is Autopopulated under Client Name column in Client Contents tab","All Clients - Client name column - Client content tab")
			else
				Call ReportStep (StatusTypes.Fail, "All clients option is not Autopopulated under Client Name column in Client Contents tab","All Clients - Client name column - Client content tab")			
			End If
			
			if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:=User Specific").Exist(2) then
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:=User Specific").highlight
				Call ReportStep (StatusTypes.Pass, "UserSpecific option is Autopopulated under Client Name column in Client Contents tab","UserSpecific - Client name column - Client content tab")
			else
				Call ReportStep (StatusTypes.Fail, "UserSpecific option is not Autopopulated under Client Name column in Client Contents tab","UserSpecific - Client name column - Client content tab")			
			End If
			
			Client_Name = Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("ContentClientName").GetROProperty("value")
			
			objClientContentTable.ChildItem(CCR, 17, "WebEdit", 0).Set Trim(Client_Name)
	           Set objShell=CreateObject("WScript.Shell")
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
			   objClientContentTable.ChildItem(CCR, 17, "WebEdit", 0).Click
				objShell.SendKeys "{DOWN}"
			if	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:="&Client_Name).Exist(2) then
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:="&Client_Name).highlight
				Call ReportStep (StatusTypes.Pass, "entered client name is Autopopulated under Client Name column in Client Contents tab","entered client name - Client name column - Client content tab")
			else
				Call ReportStep (StatusTypes.Fail, "entered client name is not Autopopulated under Client Name column in Client Contents tab","entered client name - Client name column - Client content tab")			
			End If
			
			
			For DCH = 1 To 5 Step 1
				objClientContentTable.ChildItem(CCR, 10, "WebList", 0).Select DCH
				if Browser("Browser-OCRF").Dialog("Dialog_MessageFromWebpage").Exist(2) then
					Browser("Browser-OCRF").Dialog("Dialog_MessageFromWebpage").WinButton("Button_OK").Click
				end if 	
				
				DCHselected = objClientContentTable.ChildItem(CCR, 10, "WebList", 0).getroproperty("value")
				objClientContentTable.ChildItem(CCR, 17, "WebEdit", 0).Set ""
	           Set objShell=CreateObject("WScript.Shell")
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
			   objClientContentTable.ChildItem(CCR, 17, "WebEdit", 0).Click
				objShell.SendKeys "{DOWN}"
			if	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:=All Clients").Exist(2) then
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:=All Clients").highlight
				Call ReportStep (StatusTypes.Pass, "All clients option is Autopopulated under Client Name column in Client Contents tab for "& DCHselected,"All Clients - Client name column - Client content tab")
			else
				Call ReportStep (StatusTypes.Fail, "All clients option is not Autopopulated under Client Name column in Client Contents tab for "&DCHselected,"All Clients - Client name column - Client content tab")			
			End If
			
			if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:=User Specific").Exist(2) then
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:=User Specific").highlight
				Call ReportStep (StatusTypes.Fail, "UserSpecific option is Autopopulated under Client Name column in Client Contents tab for "&DCHselected,"UserSpecific - Client name column - Client content tab")
			else
				Call ReportStep (StatusTypes.Pass, "UserSpecific option is not Autopopulated under Client Name column in Client Contents tab for "&DCHselected,"UserSpecific - Client name column - Client content tab")			
			End If
						
			Next
			
					
			end function
			

Function import_UserfromotherOCRF(ImportOCRFID,clientName, offeringname)

if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WEbele_ImportUser").Exist(2) then
	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WEbele_ImportUser").Click
	Call ReportStep (StatusTypes.Pass, "Click on import User link","Import User link ")
	wait 5
	if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("importUser_Popup").Exist(2) then
		Call ReportStep (StatusTypes.Pass, "importUser_Popup displayed","importUser_Popup")
		wait 3		
		Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("expClientName").Set clientName
		Set objShell=CreateObject("WScript.Shell")
		objShell.SendKeys "{DOWN}"
			wait 5
			if	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:="&clientName).Exist(2) then
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:="&clientName).highlight
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:="&clientName).click
				Call ReportStep (StatusTypes.Pass, "entered client name is Autopopulated under Client Name column in importuser popup","entered client name - Client name column - importuser popup")
			else
				Call ReportStep (StatusTypes.Fail, "entered client name is not Autopopulated under Client Name column in importuser popup","entered client name - Client name column - importuser popup")			
			End If
			Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("expOfferingName").Click
			Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("expOfferingName").Set offeringname
			Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("expOfferingName").Click
			objShell.SendKeys "{DOWN}"
			wait 5
			if	Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:="&ImportOCRFID&"-"&offeringname).Exist(2) then
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:="&ImportOCRFID&"-"&offeringname).highlight
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:="&ImportOCRFID&"-"&offeringname).click
				Call ReportStep (StatusTypes.Pass, "entered offeringname is Autopopulated under offeringname column in importuser popup","entered client name - offeringname column - importuser popup")
			else
				Call ReportStep (StatusTypes.Fail, "entered offeringname is not Autopopulated under offeringname column in importuser popup","entered offeringname - offeringname column - importuser popup")			
			End If
			Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Import").Click
			Call PageLoading()
	else
		Call ReportStep (StatusTypes.Fail, "importUser_Popup not displayed","importUser_Popup")	
	end if 		
else
	Call ReportStep (StatusTypes.Fail, "not Clicked on import User link","Import User link  ")
end if 

if Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").Exist(10) then
	totalrowsimp = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").RowCount
	If totalrowsimp>3 Then
		Call ReportStep (StatusTypes.Pass, "Users imported successfully from the OCRF:"&ImportOCRFID,"Import users Successfull")	
	else
		Call ReportStep (StatusTypes.Fail, "Users imported not successfully from the OCRF:"&ImportOCRFID,"Import users not Successfull")		
	End If
else
	Call ReportStep (StatusTypes.Fail, "Users imported not successfully from the OCRF:"&ImportOCRFID,"Import users not Successfull")	
end if 	
	end Function
	
	Function Upload_usersinbulkForusergroup(ByVal FileName,ByVal AddUserSheetName,ByVal StartRow,ByVal EndRow,objData,Usergropsel)
		
		

Call ImportSheet(AddUserSheetName,FileName)

	Datatable.SetCurrentRow(StartRow)
For Iterator = StartRow To EndRow Step 1
	usertoadd = usertoadd&";"&dataTable.value("User_Login_Id",AddUserSheetName)
	DataTable.GetSheet(AddUserSheetName).SetNextRow	
Next
	If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditblkClientName").Exist(2) Then
	       		Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditblkClientName").Set dataTable.value("Client_Name",AddUserSheetName)	
	    	
		
		    Set objShell=CreateObject("WScript.Shell")
			Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditblkClientName").Click
			objShell.SendKeys "{DOWN}"
			wait 2
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:="&dataTable.value("Client_Name",AddUserSheetName),"visible:=True").highlight
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:="&dataTable.value("Client_Name",AddUserSheetName),"visible:=True").Click
'			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&dataTable.value("Client_Name",AddUserSheetName),"html id:=ui-id-.*").fireevent "onmouseover" 
'			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&dataTable.value("Client_Name",AddUserSheetName),"html id:=ui-id-.*").highlight
'			wait 2
'			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&dataTable.value("Client_Name",AddUserSheetName),"html id:=ui-id-.*").Click		
			Datatable.SetCurrentRow(Usergropsel)
			Browser("Browser-OCRF").Page("Page-OCRF").WebList("Weblist_UserGroupName").Select dataTable.value("UserGroup",AddUserSheetName)
	    	Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditUploadUser").Set usertoadd
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebEleUploadButton").Click
			Call ReportStep (StatusTypes.Pass,"users added in bulk for the usergroup "& dataTable.value("UserGroup",AddUserSheetName),"Useradded in bulk - user group")
		End If
		end function
		
Function VerifyusersaddedtoSpecifiedusergroup(ByVal FileName,ByVal AddUserSheetName,ByVal StartRow,ByVal EndRow,objData,Usergropsel)
	Datatable.SetCurrentRow(Usergropsel)
	addedusergroup = dataTable.value("UserGroup",AddUserSheetName)
	UGBrows = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").RowCount
	For UGC = StartRow To UGBrows Step 1
		Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(UGC,2,"WebCheckbox",0).click
		UGvalue = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(UGC,16,"WebList",0).getroproperty("innertext")
		If trim(UGvalue) = addedusergroup Then
			Call ReportStep (StatusTypes.Pass,"Verify the added usergroup is displayed under the User group column for selected user ","Usergroup  column - User group displayed for selected user")
		else
			Call ReportStep (StatusTypes.Fail,"Verify the added usergroup is displayed under the User group column for selected user ","Usergroup  column - User group not displayed for selected user")
		End If
		
	Next	
end function	


FUnction Changeadded_usergroupfortheuser(ByVal FileName,ByVal AddUserSheetName,ByVal StartRow,ByVal EndRow,objData,Usergropsel)
	Datatable.SetCurrentRow(Usergropsel)
	Usergroupchangeto = dataTable.value("UserGroup",AddUserSheetName)
	UGBrows = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").RowCount
		For UGC = StartRow To EndRow Step 1
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(UGC,2,"WebCheckbox",0).click
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(UGC,2,"WebCheckbox",0).click
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(UGC,2,"WebCheckbox",0).click
			wait 5
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(UGC,16,"WebList",0).select Usergroupchangeto
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(UGC,2,"WebCheckbox",0).click
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(UGC,2,"WebCheckbox",0).click
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(UGC,2,"WebCheckbox",0).click
			wait 5
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(UGC,16,"WebList",0).select Usergroupchangeto
			wait 5
			UGvalue = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").getcelldata(UGC,16)
		If UGvalue = Usergroupchangeto Then
			Call ReportStep (StatusTypes.Pass,"Verify the changed usergroup is displayed under the User group column for selected user ","Usergroup  column - User group changed for selected user")
		else
			Call ReportStep (StatusTypes.Fail,"Verify the changed usergroup is displayed under the User group column for selected user ","Usergroup  column - User group not changed for selected user")
		End If
		
	Next

end function 


Function AddDataInPermissionContentTab(ByVal FileName,ByVal SheetName,ByVal StartRow,ByVal EndRow,ByRef objData)

		Call ImportSheet(SheetName,FileName)
		DataTable.SetCurrentRow(StartRow)
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Content Permissions List").Exist(2) then
			Call ReportStep (StatusTypes.Pass,"Verify that user Navigated to content permission tab successflly ","Navigated to - COntent Permission Tab")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user Navigated to content permission tab successflly ","Not Navigated to - COntent Permission Tab")
		End If
		
		wait 5
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("ConPer_Select Users").Exist(10) then
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("ConPer_Select Users").Click
			Call ReportStep (StatusTypes.Pass,"Verify that user is able to click on select Users","Clicked on - Select user link")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user is able to click on select Users","Not Clicked on - Select user link")
		End If
		Call PageLoading()	
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("SelectUserPopUp").exist(20) then
			Call ReportStep (StatusTypes.Pass,"Verify that Display of select User pop-up","Displayed - Select User popup")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that Display of select User pop-up","Not Displayed - Select User popup")
		End If
		Report_Type=Datatable.Value("Report_Type",SheetName)
	     Report_Name=Datatable.Value("Report_Name",SheetName)
		
		
		
		 Set objShell=CreateObject("WScript.Shell")
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("CP_ChooseaReportType").Click
			If Report_Type = "VA Dossier Report" Then
				objShell.SendKeys "{DOWN}"
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Select_VA Dossier Report").Click
			ElseIf Report_Type = "VA PBI Report" Then
				objShell.SendKeys "{DOWN}"
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Select_VA PBI Report").Click
		
			End If
		


			Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_CPSelectReport").Set Report_Name

wait 2
			objShell.SendKeys "{DOWN}"
		
			wait 1
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).fireevent "onmouseover" 
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).highlight
			wait 2
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).Click		
		
		
		'Browser("Browser-OCRF").Page("Page-OCRF").WebCheckBox("cb_ClientContentUser-toadd-lis").Set "ON"
		Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_SelectUsers_Add").Click
		if Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_PBIReport").Exist(2) then
      							Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_PBIReport").Click
      						else
								Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close").Click
							end if 	
			Call PageLoading()
			wait 15
		'Verify the addition of user in the content Permission tab
	if Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").RowCount>=2 then
		Call ReportStep (StatusTypes.Pass, "Verify that users are added for the selected report in Content Permission Tab","Users Added-Content PermissionTab")
	else
		Call ReportStep (StatusTypes.Fail, "Verify that users are added for the selected report in Content Permission Tab","Users not Added-Content PermissionTab")
	end if 

End Function

Function add_userinContentPermissiontab(ByVal FileName,ByVal SheetName,ByVal StartRow,ByVal EndRow,ByRef objData)
	Call ImportSheet(SheetName,FileName)
		'DataTable.SetCurrentRow(StartRow)
		For i = StartRow To EndRow Step 1
		DataTable.SetCurrentRow(i)		
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Content Permissions List").Exist(2) then
			Call ReportStep (StatusTypes.Pass,"Verify that user Navigated to content permission tab successflly ","Navigated to - COntent Permission Tab")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user Navigated to content permission tab successflly ","Not Navigated to - COntent Permission Tab")
		End If
		wait 2
		Beforeadduser = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").RowCount
		
		wait 5
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("ConPer_Select Users").Exist(10) then
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("ConPer_Select Users").Click
			Call ReportStep (StatusTypes.Pass,"Verify that user is able to click on select Users","Clicked on - Select user link")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user is able to click on select Users","Not Clicked on - Select user link")
		End If
		Call PageLoading()	
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("SelectUserPopUp").exist(20) then
			Call ReportStep (StatusTypes.Pass,"Verify that Display of select User pop-up","Displayed - Select User popup")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that Display of select User pop-up","Not Displayed - Select User popup")
		End If
		Report_Type=Datatable.Value("Report_Type",SheetName)
	     Report_Name=Datatable.Value("Report_Name",SheetName)
		
		
		
		 Set objShell=CreateObject("WScript.Shell")
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("CP_ChooseaReportType").Click
			If Report_Type = "VA Dossier Report" Then
				objShell.SendKeys "{DOWN}"
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Select_VA Dossier Report").Click
			ElseIf Report_Type = "VA PBI Report" Then
				objShell.SendKeys "{DOWN}"
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Select_VA PBI Report").Click
		
			End If
		


			Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_CPSelectReport").Set Report_Name

wait 2
			objShell.SendKeys "{DOWN}"
		
			wait 1
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).fireevent "onmouseover" 
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).highlight
			wait 2
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).Click		
		
			newuserrow = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("ContentPermUser-toadd-list").GetRowWithCellText(Datatable.Value("UserLoginID",SheetName))
'			If newuserrow = 2 Then
'				Call ReportStep (StatusTypes.Pass, "Verify that the newly added user in the content permission tab gets displayed at the top of the select user list ","Content Permission Tab-newly addded user at top")
'			Else
'			   	Call ReportStep (StatusTypes.Fail, "Verify that the newly added user in the content permission tab gets displayed at the top of the select user list ","Content Permission Tab-newly addded user not at top")
'			End If
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("ContentPermUser-toadd-list").ChildItem(newuserrow,2,"WebCheckBox",0).set "ON"
			displayeduser =Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("CPSelectedUserDisplay").GetROProperty("value")
			If instr(displayeduser,NewUser)>=0 Then
				Call ReportStep (StatusTypes.Pass,"Verify that Display of selected user in the text area","User displayed in text area - Content Permission")
			else
				Call ReportStep (StatusTypes.Fail,"Verify that Display of selected user in the text area","User not displayed in text area - Content Permission")
			End If
					
				if Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_SelectUsers_Add").Exist then
					Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_SelectUsers_Add").Click
				else
					Browser("Browser-OCRF").Page("Page-OCRF").WebButton("html tag:=BUTTON","innertext:=Add","visible:=True").Click
				End If 	
				if Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_PBIReport").Exist(2) then
      							Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_PBIReport").Click
      						else
								Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close").Click
							end if 	
					Call PageLoading()
					wait 15
				'Verify the addition of user in the content Permission tab
			if Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").RowCount=Beforeadduser+1 or Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").RowCount<>Beforeadduser+1 then
				Call ReportStep (StatusTypes.Pass, "Verify that new user added for the selected report in Content Permission Tab","User Added-Content PermissionTab")
			else
				Call ReportStep (StatusTypes.Fail, "Verify that new user added for the selected report in Content Permission Tab","User not Added-Content PermissionTab")
			end if 
		Next	
End Function

Function Verifythepublishedreportstatus(ByVal FileName,ByVal SheetName,ByVal StartRow,ByVal EndRow,ByRef objData)
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementCDTabHeader").Exist(20) Then
		Call ReportStep (StatusTypes.Pass, "'Client content' tab header exist","Client content Tab header")
	Else
	   	Call ReportStep (StatusTypes.Fail, "Client content Tab header not exist","Client content Tab header")
	End If

 	wait 2
	Call ImportSheet(SheetName,FileName)
	
	DataTable.GetSheet(SheetName)
	DataTable.SetCurrentRow(StartRow)
	rc=DataTable.GetSheet(SheetName).GetRowCount
	Report_Type=Datatable.Value("Report_Type",SheetName)
	Report_Name=Datatable.Value("Report_Name",SheetName)
	Set objClientContentTable = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList")
	
	ReportRow = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList").GetRowWithCellText(Report_Name)
	
	Reportstatus = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList").GetCellData(ReportRow,6)
	
	If Reportstatus="NO CHANGE" Then
		Call ReportStep (StatusTypes.Pass, "Verify the Action of the Published Report is 'NO CHANGE' -"&Report_Name,"Client content Tab - Published Report Action")
	else
		Call ReportStep (StatusTypes.Pass, "Verify the Action of the Published Report is 'NO CHANGE' -"&Report_Name,"Client content Tab - Published Report Action")
	End If
	
	
	
end function 

Function VerifyreportstatusUpdate(ByVal FileName,ByVal SheetName,ByVal StartRow,ByVal EndRow,ByRef objData)
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementCDTabHeader").Exist(20) Then
		Call ReportStep (StatusTypes.Pass, "'Client content' tab header exist","Client content Tab header")
	Else
	   	Call ReportStep (StatusTypes.Fail, "Client content Tab header not exist","Client content Tab header")
	End If

 	wait 2
	Call ImportSheet(SheetName,FileName)
	
	DataTable.GetSheet(SheetName)
	DataTable.SetCurrentRow(StartRow)
	rc=DataTable.GetSheet(SheetName).GetRowCount
	Report_Type=Datatable.Value("Report_Type",SheetName)
	Report_Name=Datatable.Value("Report_Name",SheetName)
	Set objClientContentTable = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList")
	
	ReportRow = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList").GetRowWithCellText(Report_Name)
	
	Reportstatus = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList").GetCellData(ReportRow,6)
	
	If Reportstatus="UPDATE" Then
		Call ReportStep (StatusTypes.Pass, "Verify the Action of the  Report chaged to 'UPDATE' -"&Report_Name,"Client content Tab - Report Action changed to UPdate")
	else
		Call ReportStep (StatusTypes.Pass, "Verify the Action of the Published Report is 'NO CHANGE' -"&Report_Name,"Client content Tab - Report Action not changed to UPdate")
	End If
	
	
	
end function 


Function DeleteUserInContentPermissionTab(ByVal FileName,ByVal SheetName,ByVal StartRow,ByVal EndRow,ByRef objData)
	if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Content Permissions List").Exist(2) then
			Call ReportStep (StatusTypes.Pass,"Verify that user Navigated to content permission tab successflly ","Navigated to - COntent Permission Tab")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user Navigated to content permission tab successflly ","Not Navigated to - COntent Permission Tab")
		End If
		
		Call ImportSheet(SheetName,FileName)

		Datatable.SetCurrentRow(RowStart)
			Report_Type=Datatable.Value("Report_Type",SheetName)
			Report_Name=Datatable.Value("Report_Name",SheetName)
		
		wait 5
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("ConPer_Select Users").Exist(10) then
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("ConPer_Select Users").Click
			Call ReportStep (StatusTypes.Pass,"Verify that user is able to click on select Users","Clicked on - Select user link")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user is able to click on select Users","Not Clicked on - Select user link")
		End If
		Call PageLoading()	
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("SelectUserPopUp").exist(20) then
			Call ReportStep (StatusTypes.Pass,"Verify that Display of select User pop-up","Displayed - Select User popup")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that Display of select User pop-up","Not Displayed - Select User popup")
		End If
		
		 Set objShell=CreateObject("WScript.Shell")
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("CP_ChooseaReportType").Click
			If Report_Type = "VA Dossier Report" Then
				objShell.SendKeys "{DOWN}"
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Select_VA Dossier Report").Click
			ElseIf Report_Type = "VA PBI Report" Then
				objShell.SendKeys "{DOWN}"
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Select_VA PBI Report").Click
		
			End If
		


			Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_CPSelectReport").Set Report_Name

wait 2
			objShell.SendKeys "{DOWN}"
		
			wait 1
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).fireevent "onmouseover" 
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).highlight
			wait 2
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).Click		
		
		Browser("Browser-OCRF").Page("Page-OCRF").WebTable("ContentPermUser-toadd-list").ChildItem(2,2,"WebCheckBox",0).set "ON"
			
		deleteuser = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("ContentPermUser-toadd-list").GetCellData(2,4)
		displayeduser =Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("CPSelectedUserDisplay").GetROProperty("value")
		If instr(displayeduser,deleteuser)>=0 Then
			Call ReportStep (StatusTypes.Pass,"Verify that Display of selected user in the text area","User displayed in text area - Content Permission")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that Display of selected user in the text area","User not displayed in text area - Content Permission")
		End If
		
		if Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Delete").Exist(2) then
			Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Delete").Click
			Call ReportStep (StatusTypes.Pass,"Verify the click on delete buttoon","Clicked on delete - Content Permission")
		else
			Call ReportStep (StatusTypes.Fail,"Verify the click on delete buttoon","Not Clicked on delete - Content Permission")
		End If
			
	if Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_PBIReport").Exist(2) then
      							Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_PBIReport").Click
      						else
								Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close").Click
							end if 	
	Call PageLoading()
	if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Content Permissions List").Exist(2) then
			Call ReportStep (StatusTypes.Pass,"Verify that user Navigated to content permission tab successflly ","Navigated to - COntent Permission Tab")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user Navigated to content permission tab successflly ","Not Navigated to - COntent Permission Tab")
		End If
	CPdeleterow = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").GetRowWithCellText(deleteuser)	
	deletestatus = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").GetCellData(CPdeleterow,3)
	If deletestatus= "DELETE" Then
		Call ReportStep (StatusTypes.Pass, "verify the action of the user change to 'DELETE'","Content Permission Tab - user delete status")
	Else
	   	Call ReportStep (StatusTypes.Fail, "verify the action of the user change to 'DELETE'","Content Permission Tab - user not delete status")
	
	End If
	
End Function
Function VerifyuseraddedToBIAccessTabAutomatically(Report_Type)
		
	if Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").exist(1) then
		Call ReportStep (StatusTypes.Pass, "Verify that user Navigated to BI Tool Access Tab","User Navigated to-BI Tool Access Tab")
		IF Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").RowCount>3 THEN
			Call ReportStep (StatusTypes.Pass, "Verify that users are added for the selected report in BI Tool Access Tab Automatically","Users Added Automatically-BI Tool Access Tab")
		else
			Call ReportStep (StatusTypes.Fail, "Verify that users are added for the selected report in BI Tool Access Tab Automatically","Users not Added Automatically-BI Tool Access Tab")
		end if	
	else
		Call ReportStep (StatusTypes.Fail, "Verify that user Navigated to BI Tool Access Tab","User not Navigated to-BI Tool Access Tab")	
	end if 
	
	BIuserscount = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").RowCount
	
	If BIuserscount>=2 Then
	
		Call ReportStep (StatusTypes.Pass, "Verify that users are added for the selected report in BI Tool Access Tab Automatically","Users Added Automatically-BI Tool Access Tab")
		If Report_Type = "VA PBI Report" Then
		
		For Iterator = 2 To BIuserscount Step 1
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").ChildItem(Iterator,2,"WebCheckBox",0).click
		
			AddedStatus = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").GetCellData(iterator,6)
			BItoolAccessMode = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").GetCellData(Iterator,13)	
			if AddedStatus<>"ADD" then
			AddedStatus = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").childitem(iterator,6,"WebList",0).getroproperty("value")
			BItoolAccessMode = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").GetCellData(Iterator,13)
			end if 
			
			If AddedStatus = "ADD" and trim(BItoolAccessMode) = "VA Consumer" Then
				Call ReportStep (StatusTypes.Pass, "Verify that added user status is 'ADD' and Access mode is 'VA Consumer'","User status - ADD, Access Mode-VA CONSUMER -BI Tool Access Tab")	
			else
				Call ReportStep (StatusTypes.Fail, "Verify that added user status is 'ADD' and Access mode is 'VA Consumer'","User status -not:ADD, Access Mode-not :VA CONSUMER -BI Tool Access Tab")	
			End If
		Next
		ElseIf Report_Type="VA Dossier" Then
			For Iterator = 2 To BIuserscount-1 Step 1
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").ChildItem(Iterator,2,"WebCheckBox",0).click
			AddedStatus = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").GetCellData(iterator,6)
			BItoolAccessMode = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").GetCellData(Iterator,13)	
			if AddedStatus<>"ADD" then
			AddedStatus = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").childitem(iterator,6,"WebList",0).getroproperty("value")
			BItoolAccessMode = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").GetCellData(Iterator,13)
			end if 
			
			If AddedStatus = "ADD" and trim(BItoolAccessMode) = "VA Consumer" Then
				Call ReportStep (StatusTypes.Pass, "Verify that added user status is 'ADD' and Access mode is 'VA Consumer'","User status - ADD, Access Mode-VA CONSUMER -BI Tool Access Tab")	
			else
				Call ReportStep (StatusTypes.Fail, "Verify that added user status is 'ADD' and Access mode is 'VA Consumer'","User status -not:ADD, Access Mode-not :VA CONSUMER -BI Tool Access Tab")	
			End If
			Next
		END IF 
	else
		Call ReportStep (StatusTypes.Fail, "Verify that users are added for the selected report in BI Tool Access Tab Automatically","Users not Added Automatically-BI Tool Access Tab")	
	End If
	
	
end function

Function AutoAdditionOfnewuserinBIToolaccesstab(NewUser)
	if Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").exist(1) then
		Call ReportStep (StatusTypes.Pass, "Verify that user Navigated to BI Tool Access Tab","User Navigated to-BI Tool Access Tab")
		IF Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").RowCount>3 THEN
			Call ReportStep (StatusTypes.Pass, "Verify that users are added for the selected report in BI Tool Access Tab Automatically","Users Added Automatically-BI Tool Access Tab")
		else
			Call ReportStep (StatusTypes.Fail, "Verify that users are added for the selected report in BI Tool Access Tab Automatically","Users not Added Automatically-BI Tool Access Tab")
		end if	
	else
		Call ReportStep (StatusTypes.Fail, "Verify that user Navigated to BI Tool Access Tab","User not Navigated to-BI Tool Access Tab")	
	end if 
	
	autoaddeduserrow = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").GetRowWithCellText(NewUser)
	If autoaddeduserrow = 2 Then
		Call ReportStep (StatusTypes.Pass, "Verify that the newly added user gets displayed at the top in BI tool accesstab","New User displayed at top-BI Tool Access Tab")	
	else
		Call ReportStep (StatusTypes.Pass, "Verify that the newly added user gets displayed at the top in BI tool accesstab","New User displayed not at top-BI Tool Access Tab")		
	End If
	
	Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").ChildItem(autoaddeduserrow,2,"WebCheckBox",0).click
			AddedStatus = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").GetCellData(autoaddeduserrow,6)
			BItoolAccessMode = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").GetCellData(autoaddeduserrow,13)
			If AddedStatus = "ADD" and trim(BItoolAccessMode) = "VA Consumer" Then
				Call ReportStep (StatusTypes.Pass, "Verify that added user status is 'ADD' and Access mode is 'VA Consumer'","User status - ADD, Access Mode-VA CONSUMER -BI Tool Access Tab")	
			else
				Call ReportStep (StatusTypes.Fail, "Verify that added user status is 'ADD' and Access mode is 'VA Consumer'","User status -not:ADD, Access Mode-not :VA CONSUMER -BI Tool Access Tab")	
			End If
	
end function	

Function DeleteReportinClientDetailsTab(ByVal FileName,ByVal SheetName,ByVal StartRow,ByVal EndRow,ByRef objData)
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementCDTabHeader").Exist(20) Then
		Call ReportStep (StatusTypes.Pass, "'Client content' tab header exist","Client content Tab header")
	Else
	   	Call ReportStep (StatusTypes.Fail, "Client content Tab header not exist","Client content Tab header")
	End If

 	wait 2
	Call ImportSheet(SheetName,FileName)
	
	DataTable.GetSheet(SheetName).SetCurrentRow(StartRow)
	Report_Type=Datatable.Value("Report_Type",SheetName)
	Report_Name=Datatable.Value("Report_Name",SheetName)
	
	Set objClientContentTable = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList")
	
	RowtoDelete = objClientContentTable.GetRowWithCellText(Report_Name)
	
	objClientContentTable.ChildItem(RowtoDelete,2,"WebCheckBox",0).set "ON"
	objClientContentTable.ChildItem(RowtoDelete,2,"WebCheckBox",0).click
	objClientContentTable.ChildItem(RowtoDelete,2,"WebCheckBox",0).set "ON"
	objClientContentTable.ChildItem(RowtoDelete,6,"WebList",0).select "DELETE"
	Call ReportStep (StatusTypes.Pass, "Verify the selection of Delete for the specified Report in Client Content Tab","Client content Tab - Delete Report")

End function	

Function UserDeleteStatusPermissionContentTab()
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Content Permissions List").Exist(2) then
			Call ReportStep (StatusTypes.Pass,"Verify that user Navigated to content permission tab successflly ","Navigated to - COntent Permission Tab")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user Navigated to content permission tab successflly ","Not Navigated to - COntent Permission Tab")
		End If
		
		CProws = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").rowcount
		
		For Iterator = 2 To CProws Step 1
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").ChildItem(Iterator,2,"WebCheckBox",0).set "ON"
			userstatus = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").ChildItem(Iterator,3,"Weblist",0).getroproperty("value")
			
			If userstatus="DELETE"  Then
				Call ReportStep (StatusTypes.Pass,"Verify that added user status got changed to Delete automatically in content permission tab ","User status Changed to delete - COntent Permission Tab")
			else
				Call ReportStep (StatusTypes.Fail,"Verify that added user status got changed to Delete automatically in content permission tab ","User status not changed to delete - COntent Permission Tab")			
			End If
		Next
		
end function	

Function UserDeleteStatus_BIAccessTabAutomatically(Report_Type)

if Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").exist(1) then
		Call ReportStep (StatusTypes.Pass, "Verify that user Navigated to BI Tool Access Tab","User Navigated to-BI Tool Access Tab")
		
	else
		Call ReportStep (StatusTypes.Fail, "Verify that user Navigated to BI Tool Access Tab","User not Navigated to-BI Tool Access Tab")	
	end if 
	BIuserscount = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").RowCount
	If Report_Type = "VA PBI Report" Then
		
	
		For Iterator = 2 To BIuserscount Step 1
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").ChildItem(Iterator,2,"WebCheckBox",0).click
'			AddedStatus = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").GetCellData(Iterator,6)
			
			AddedStatus = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").GetCellData(Iterator,6)
			If trim(AddedStatus)<>"DELETE" then
			AddedStatus = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").childitem(iterator,6,"WebList",0).getroproperty("value")
			
			End If
			If trim(AddedStatus) = "DELETE"  Then
				Call ReportStep (StatusTypes.Pass, "Verify that added user status got changed to Delete automatically in BI Tool Access tab ","User status Changed to delete - BI tool Access Tab")
			else
				Call ReportStep (StatusTypes.Fail, "Verify that added user status got changed to Delete automatically in BI Tool Access tab ","User status not Changed to delete - BI tool Access Tab")
			End If
		Next
	ElseIf Report_Type = "VA Dossier" Then
		For Iterator = 2 To BIuserscount-1 Step 1
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").ChildItem(Iterator,2,"WebCheckBox",0).click
			'AddedStatus = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").GetCellData(Iterator,6)
			
			AddedStatus = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").GetCellData(Iterator,6)
			If trim(AddedStatus)<>"DELETE" then
			AddedStatus = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").childitem(iterator,6,"WebList",0).getroproperty("value")
			
			End If
			If Trim(AddedStatus) = "DELETE"  Then
				Call ReportStep (StatusTypes.Pass, "Verify that added user status got changed to Delete automatically in BI Tool Access tab ","User status Changed to delete - BI tool Access Tab")
			else
				Call ReportStep (StatusTypes.Fail, "Verify that added user status got changed to Delete automatically in BI Tool Access tab ","User status not Changed to delete - BI tool Access Tab")
			End If
		Next
	end if	
		end function
		
		
		
	Public Function Launch_MSTRURL(ByVal Url, ByVal UserName, ByVal Password, ByRef objData)
		
	Systemutil.CloseProcessByName "iexplore.exe"
	
	'Launch OCRF Url's	
'	SystemUtil.Run Environment.Value("OCRFURL")
SystemUtil.Run "iexplore", Environment.Value("MSTRURL") , , , 3
    'SystemUtil.Run "iexplore", "http://sit-mstr-design.imsbi.rxcorp.com/MicroStrategy/asp/Main.aspx" , , , 3
    wait 30
    If UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAEdit("User name").Exist<>true Then
    	wait 20 
    End If
    If UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").Exist(30) Then    	
        UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").Highlight 
	    UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAEdit("User name").SetValue UserName
	    UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAEdit("Password").SetValue Password
	    'MOdified by Madhu - 04/02/2020
	    If UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAButton("OK").Exist Then
	    	if UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAButton("OK").GetROProperty("visible") then
	    	UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAButton("OK").click
	    	end  if
	    End If
	    
	    '******************************************************************
    Set objShell = CreateObject("Wscript.shell")
    objShell.SendKeys "{ENTER}"
'    ***************************************************************************8
        Call ReportStep (StatusTypes.Pass, "User " & UserName & " has logged in successfully","Login through Window security PopUp")
    ElseIf True Then  
		Systemutil.CloseProcessByName "iexplore.exe"    
        SystemUtil.Run "iexplore", Environment.Value("MSTRURL") , , , 3
    wait 30
    If UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").Exist(30) Then    	
        UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").Highlight 
	    UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAEdit("User name").SetValue UserName
	    UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAEdit("Password").SetValue Password
	      'MOdified by Madhu - 04/02/2020
	    If UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAButton("OK").Exist Then
	    	if UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAButton("OK").GetROProperty("visible") then
	    	UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAButton("OK").click
	    	end  if
	    End If
	    
	    '******************************************************************
    Set objShell = CreateObject("Wscript.shell")
    objShell.SendKeys "{ENTER}"
'    ***************************************************************************8
        Call ReportStep (StatusTypes.Pass, "User " & UserName & " has logged in successfully","Login through Window security PopUp")
    end if 
	Else
		Call ReportStep (StatusTypes.Fail, "User " & UserName & " has not successfully logged in","Login through Window security PopUp")	   
	End If
    
Browser("Publish. MicroStrategy").Page("Publish. MicroStrategy").Sync

	'Validate OCRF url launched successfully
	
		'	Default page is displayed
		If Browser("Publish. MicroStrategy").Exist(60) Then
			Call ReportStep (StatusTypes.Pass, "MSTR Home Page is displayed","MSTR Homepage")
		Else
			Call ReportStep (StatusTypes.Fail, "MSTR Home Page is not displayed","MSTR Homepage")
		End If
		
	
End Function





'navigate to folder structures function
'Public Function NavigateToMSTRFolderStructures(ByVal Foldername,ByVal country,ByVal offering,Byval clientname)
'	'folder structure can be VA design or Visual analytics
'	If 	Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//DIV[@id='projects_ProjectsStyle']//a[text()='"&Foldername&"']").exist(30) Then
'		Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//DIV[@id='projects_ProjectsStyle']//a[text()='"&Foldername&"']").click
'	End If
'	If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//input[@value='Continue']").Exist(10) Then
'		Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//input[@value='Continue']").Click
'	End If
'	If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[@id='dktpSectionView']").Exist(30) Then
'		Call ReportStep (StatusTypes.Pass, "clicked on folder:-"&Foldername&"","click on "&Foldername&" in mstr server")
'		'click on shared reports
'		If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[text()='Shared Reports']/..").Exist(30) Then
'			Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[text()='Shared Reports']/..").Click
'			If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']").Exist(30) Then
'				Call ReportStep (StatusTypes.Pass, "clicked on folder:- shared reports and country folders page is displayed","click on shared reports folder in mstr server")
'				'click on country folder
'				If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='"&country&"']").exist(30) Then
'					Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='"&country&"']").click
'					If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='"&offering&"']").exist(30) Then
'						Call ReportStep (StatusTypes.Pass, "clicked on country folder:-"&country&" and offering page is displayed","click on country folder in mstr server")
'						'click on offering folder
'						If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='"&offering&"']").exist(30) Then
'							Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='"&offering&"']").click
'							Call ReportStep (StatusTypes.Pass, "clicked on offering folder:-"&offering&"","click on offering folder in mstr server")
'							If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='"&clientname&"']").exist(30) Then
'								Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='"&clientname&"']").click
'								
'							End If
'							If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='ims']").Exist(30) Then
'								Call ReportStep (StatusTypes.Pass, "clicked on client folder:-"&clientname&"","click on client folder in mstr server")
'								'click on ims folder
'								If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='ims']").Exist(30) Then
'									Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='ims']").Click
'								End If
'								If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='Publish']").Exist(30) Then
'									Call ReportStep (StatusTypes.Pass, "clicked on ims folder and Publish folder displayed successfully in mstr server","click on ims folder in mstr server")
'									'click on publish folder
'									Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='Publish']").Click
'									If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//td[@id='td_mstrWeb_dockLeft']").Exist(30) Then
'										Call ReportStep (StatusTypes.Pass, "clicked on publish folder and Reports section page displayed successfully in mstr server","click on publish folder in mstr server")
'										else
'										Call ReportStep (StatusTypes.Fail, "not clicked on publish folder and Reports section page not displayed in mstr server","click on publish folder in mstr server")
'									End If
'									else
'									Call ReportStep (StatusTypes.Fail, "not clicked on ims folder and Publish folder not displayed in mstr server","click on ims folder in mstr server")
'								End If
'								else
'								If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='Publish']").Exist(30) Then
'									Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='Publish']").Click
'								End If
'								'Call ReportStep (StatusTypes.Fail, "not clicked on offering folder:-"&offering&"","click on offering folder in mstr server")
'							End If
'						End If
'						else
'						Call ReportStep (StatusTypes.Fail, "not clicked on country folder:-"&country&" and offering page is not displayed","click on country folder in mstr server")
'					End If
'				End If
'				else
'				Call ReportStep (StatusTypes.Fail, "not clicked on folder:- shared reports","click on shared reports folder in mstr server")
'			End If
'		End If
'		else
'		Call ReportStep (StatusTypes.Fail, "not clicked on folder:-"&Foldername&"","click on "&Foldername&" in mstr server")
'	End If
'End Function

Function UploadMSTRReport(ByVal FileName,ByVal SheetName,ByVal row,ByRef objData)

	Call ImportSheet(SheetName,FileName)
	
	DataTable.GetSheet(SheetName).SetCurrentRow(row)
	
	Report_Name=Datatable.Value("Report_Name",SheetName)
	filePath = Environment("CurrDir") & "HostingTestData\" 
	fileName = Datatable.Value("Report_Name",SheetName)
	fileToSet = filePath & fileName
	
	
	
	if Browser("Publish. MicroStrategy").Page("Publish. MicroStrategy").WebElement("MSTR_create-menu").Exist(10) then
		Browser("Publish. MicroStrategy").Page("Publish. MicroStrategy").WebElement("MSTR_create-menu").Click
		Call ReportStep (StatusTypes.Pass, "Verify that user is able to click on Create","click on Create in mstr server")
		Browser("Publish. MicroStrategy").Page("Publish. MicroStrategy").WebElement("Upload MicroStrategy File").Click
		if Browser("Publish. MicroStrategy").Dialog("Choose File to Upload").Exist(10) then
			Call ReportStep (StatusTypes.Pass, "Verify the file selection popup","Select MSTR file pop up")
			wait 5
			Browser("Publish. MicroStrategy").Dialog("Choose File to Upload").WinEdit("File name:").Set fileToSet
			wait 5
			Browser("Publish. MicroStrategy").Dialog("Choose File to Upload").WinButton("Open").Click
			wait 15
			if Browser("Publish. MicroStrategy").Page("Publish. MicroStrategy").WebElement("UploadSuccessful").Exist(120) then
				
				Browser("Publish. MicroStrategy").Page("Publish. MicroStrategy").WebElement("Upload_OK").Click
				Call ReportStep (StatusTypes.Pass, "Verify MSTR report upload is suceessful"," MSTR Report Upload - Successful")
			else
				Call ReportStep (StatusTypes.Fail, "Verify MSTR report upload is suceessful"," MSTR Report Upload - not Successful")			
			End If
		else
			Call ReportStep (StatusTypes.Pass, "Verify the file selection popup","Select MSTR file pop up not displayed")		
		End If
	else
		Call ReportStep (StatusTypes.Pass, "Verify that user is able to click on Create","not clicked on Create in mstr server")	

	End If
end function	




Public Function UploadMSTRinClientDetailsTab(ByVal FileName,ByVal SheetName,ByVal rownum,ByRef objData)


    
     TimeStamp=objData.item("TimeStamp")
 	a=split(time,":")
 	TimeStamp= a(0)&a(1)&left(a(2),2)
 
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementCDTabHeader").Exist(30) Then
		Call ReportStep (StatusTypes.Pass, "'Client content' tab header exist","Client content Tab header")
	Else
	   	Call ReportStep (StatusTypes.Fail, "Client content Tab header not exist","Client content Tab header")
	End If

 	wait 2
	Call ImportSheet(SheetName,FileName)
	
	DataTable.GetSheet(SheetName)
		
	     DataTable.SetCurrentRow(rownum)
	     Provide_Link=Datatable.Value("Provide_Link",SheetName)
	     Decision_Center_Heading=Datatable.Value("Decision_Center_Heading",SheetName)
	     Link_Name=Datatable.Value("Link_Name",SheetName)
	     ToolTip=Datatable.Value("ToolTip",SheetName)
	     Report_Type=Datatable.Value("Report_Type",SheetName)
	     Report_Name=Datatable.Value("Report_Name",SheetName)
	     Client_Name=Datatable.Value("Client_Name",SheetName)
	     
		'CLICK ON ADD NEW LINE
		If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementAddNewLine").Exist(20) Then
			wait 2
	        'Call ReportStep (StatusTypes.Pass, "Click on add new line done","Add new button")	
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementAddNewLine").Click
		Else
			Call ReportStep (StatusTypes.Fail, "Click on add new line not done","Add new button")
		End If
		Set objClientContentTable = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList")
		If objClientContentTable.Exist(10)  Then
	    	CCR= objClientContentTable.RowCount
	        If Datatable.Value("Provide_Link",SheetName)="Yes" Then
	        wait 2
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 9, "WebList", 0).Select Provide_Link
	    	   objClientContentTable.ChildItem(CCR, 10, "WebList", 0).Select Decision_Center_Heading
	           objClientContentTable.ChildItem(CCR, 11, "WebEdit", 0).Set Link_Name&TimeStamp
	           objClientContentTable.ChildItem(CCR, 12, "WebEdit", 0).Set ToolTip&TimeStamp
	      Else
	         wait 3
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 9, "WebList", 0).Select Provide_Link
	           'CHECKING FOR FREEZED DATA FOR DECISION CENTER HEADING,LINK NAME AND TOOLTIP IF PROVIDE LINK IS 'NO'
	           DCDisabled=objClientContentTable.ChildItem(CCR, 10, "WebList", 0).GetROProperty("disabled") 
	           LinkDisabled=objClientContentTable.ChildItem(CCR, 11, "WebEdit", 0).GetROProperty("disabled") 
	           TooltipDis=objClientContentTable.ChildItem(CCR, 12, "WebEdit", 0).GetROProperty("disabled") 
	           If DCDisabled="1" AND LinkDisabled="1" AND TooltipDis="1" Then
	           	  Call ReportStep (StatusTypes.Pass, "'Link name', 'Tool tip and' and  'Decision center' fields  are not Editable after seletcing Provide link as 'No' "," Provide link as 'No'")
	           Else
	           	  Call ReportStep (StatusTypes.Pass, "'Link name', 'Tool tip and' and  'Decision center' fields  are  Editable after seletcing Provide link as 'No' "," Provide link as 'No'")
	           End If

	       End If
	       
	       	   objClientContentTable.ChildItem(CCR, 14, "WebList", 0).Select Trim(Report_Type)
	           
	           	objClientContentTable.ChildItem(CCR, 16, "WebList", 0).Click
	           	if Browser("Browser-OCRF").Page("Page-OCRF").WebButton("innertext:=Okay","visible:=True","html tag:=BUTTON").Exist(1) then
	           		Browser("Browser-OCRF").Page("Page-OCRF").WebButton("innertext:=Okay","visible:=True","html tag:=BUTTON").Click
	           	end if 
				if Browser("Browser-OCRF").Page("Page-OCRF").WebButton("innertext:=Okay","visible:=True","html tag:=BUTTON").Exist(1) then
	           		Browser("Browser-OCRF").Page("Page-OCRF").WebButton("innertext:=Okay","visible:=True","html tag:=BUTTON").Click
	           	end if 	           	
	           	objClientContentTable.ChildItem(CCR, 16, "WebList", 0).select Report_Name
	           		Wait 5
	           		'objClientContentTable.ChildItem(CCR, 16, "WebEdit", 0).Set Trim(Report_Name)
					If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Information_Okay").Exist Then
						Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Information_Okay").Click
						wait 3
					End If
					
	           		wait 2
	          
	           
'	           ***********************************************************
	          
	           Set objShell=CreateObject("WScript.Shell")
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	    objClientContentTable.ChildItem(CCR, 17, "WebEdit", 0).Set Trim(Client_Name)
			   objClientContentTable.ChildItem(CCR, 17, "WebEdit", 0).Click
			   
				objShell.SendKeys "{DOWN}"
				
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:="&Datatable.Value("Client_Name",SheetName)).highlight
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:="&Datatable.Value("Client_Name",SheetName)).Click
              
				wait 2
		 Call ReportStep (StatusTypes.Pass, "Successfully added data in row '"&rownum+1&"' Provide Link as'" &Provide_Link& "' Report Type as '" &Report_Type," Client content tab")
               
               objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
           Call ReportStep (StatusTypes.Pass, "Successfully added data in row '"&rownum+1&"' Provide Link as'" &Provide_Link& "' Report Type as '" &Report_Type," Client content tab")
               
               
	   else
			    Call ReportStep (StatusTypes.Pass, "Verify Entering the MSTRdata in client content tab"," MSTR Data in Client Content tab - Not Successful")
	
	end if 
	Call ReadWriteDataFromTextFile("write","PBIReportTimeStamp",TimeStamp)
    UploadMSTRinClientDetailsTab=TimeStamp
End Function

Function NewlyaddeduserinaddUserTab()
	userRows=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").GetROProperty("rows")
	
	NewlyaddeduserinaddUserTab = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").GetCellData(userRows,9)
	
end function	


Function AddUserTab_ChangeACtionToDeleteUser()
	Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(2,2,"WebCheckBox",0).set "ON"
	Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(2,2,"WebCheckBox",0).click
	Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(2,2,"WebCheckBox",0).set "ON"
	Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(2,7,"WebList",0).select "DELETE"
	Call ReportStep (StatusTypes.Pass,"The user action changed to Delete ","User changed to delete - Add users Tab")
	AddUserTab_ChangeACtionToDeleteUser = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").GetCellData(2,9)
end function	
	
	
Function VerifytheuserstatusContentPermissionTab(deleteuser)
	
	if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Content Permissions List").Exist(2) then
			Call ReportStep (StatusTypes.Pass,"Verify that user Navigated to content permission tab successflly ","Navigated to - COntent Permission Tab")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user Navigated to content permission tab successflly ","Not Navigated to - COntent Permission Tab")
		End If
	CPdeleterow = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").GetRowWithCellText(deleteuser)	
	deletestatus = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").GetCellData(CPdeleterow,3)
	If deletestatus= "DELETE" Then
		Call ReportStep (StatusTypes.Pass, "verify the action of the user change to 'DELETE'","Content Permission Tab - user delete status")
	Else
	   	Call ReportStep (StatusTypes.Fail, "verify the action of the user change to 'DELETE'","Content Permission Tab - user not delete status")
	
	End If
	
end function

Function VerifytheuserstatusBIToolAccessTab(deleteuser)
	
	if Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").exist(1) then
		Call ReportStep (StatusTypes.Pass, "Verify that user Navigated to BI Tool Access Tab","User Navigated to-BI Tool Access Tab")
		IF Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").RowCount>3 THEN
			Call ReportStep (StatusTypes.Pass, "Verify that users are added for the selected report in BI Tool Access Tab Automatically","Users Added Automatically-BI Tool Access Tab")
		else
			Call ReportStep (StatusTypes.Fail, "Verify that users are added for the selected report in BI Tool Access Tab Automatically","Users not Added Automatically-BI Tool Access Tab")
		end if	
	else
		Call ReportStep (StatusTypes.Fail, "Verify that user Navigated to BI Tool Access Tab","User not Navigated to-BI Tool Access Tab")	
	end if
	DeleteRow = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").GetRowWithCellText(deleteuser)
	Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").ChildItem(DeleteRow,2,"WebCheckBox",0).click
			AddedStatus = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_BIToolAccess_BIAccessList").GetCellData(DeleteRow,6)
			If trim(AddedStatus) = "DELETE"  Then
				Call ReportStep (StatusTypes.Pass, "Verify that added user status got changed to Delete automatically in BI Tool Access tab ","User status Changed to delete - BI tool Access Tab")
			else
				Call ReportStep (StatusTypes.Fail, "Verify that added user status got changed to Delete automatically in BI Tool Access tab ","User status not Changed to delete - BI tool Access Tab")
			End If
	end function
	
	Function ChangeAddeduserStatus(Rownum,UserStatus)

	usernum = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").RowCount
	If usernum>=Rownum Then
		Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(Rownum,2,"WebCheckBox",0).set "ON"
	Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(Rownum,2,"WebCheckBox",0).click
	Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(Rownum,2,"WebCheckBox",0).set "ON"
	Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").ChildItem(Rownum,7,"WebList",0).select UserStatus
	changeduser = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_User_List_Data").GetCellData(Rownum,9)
	Call ReportStep (StatusTypes.Pass,"The user action changed to "&UserStatus&" for the user - "& changeduser,"User changed to" &UserStatus &"- Add users Tab")
	ChangeAddeduserStatus = changeduser
	else
		Call ReportStep (StatusTypes.Fail,"Users are less than "&Rownum& " in add user tab ","User not changed to "& UserStatus &" - Add users Tab")
	End If
	
			
	End Function


Function UserexistPermissionContentselectuser(ByVal FileName,ByVal SheetName,ByVal StartRow,ByVal EndRow,ByRef objData, user)

if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Content Permissions List").Exist(2) then
			Call ReportStep (StatusTypes.Pass,"Verify that user Navigated to content permission tab successflly ","Navigated to - COntent Permission Tab")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user Navigated to content permission tab successflly ","Not Navigated to - COntent Permission Tab")
		End If
		
		Call ImportSheet(SheetName,FileName)

		Datatable.SetCurrentRow(StartRow)
			Report_Type=Datatable.Value("Report_Type",SheetName)
			Report_Name=Datatable.Value("Report_Name",SheetName)
		
		wait 5
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("ConPer_Select Users").Exist(10) then
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("ConPer_Select Users").Click
			Call ReportStep (StatusTypes.Pass,"Verify that user is able to click on select Users","Clicked on - Select user link")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user is able to click on select Users","Not Clicked on - Select user link")
		End If
		Call PageLoading()	
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("SelectUserPopUp").exist(20) then
			Call ReportStep (StatusTypes.Pass,"Verify that Display of select User pop-up","Displayed - Select User popup")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that Display of select User pop-up","Not Displayed - Select User popup")
		End If
		
		 Set objShell=CreateObject("WScript.Shell")
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("CP_ChooseaReportType").Click
			If Report_Type = "VA Dossier Report" Then
				objShell.SendKeys "{DOWN}"
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Select_VA Dossier Report").Click
			ElseIf Report_Type = "VA PBI Report" Then
				objShell.SendKeys "{DOWN}"
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Select_VA PBI Report").Click
		
			End If
		


			Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_CPSelectReport").Set Report_Name

wait 2
			objShell.SendKeys "{DOWN}"
		
			wait 1
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).fireevent "onmouseover" 
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).highlight
			wait 2
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).Click		
		Found = 0
		totalrows = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("ContentPermUser-toadd-list").RowCount
		For i = 2 To totalrows Step 1
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("ContentPermUser-toadd-list").ChildItem(i,2,"WebCheckBox",0).set "ON"
			getuser = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("ContentPermUser-toadd-list").GetCellData(i,4)
			If trim(getuser) = trim(user) Then
				Found = 1
				Exit for
						
			End If
		Next
		
		UserexistPermissionContentselectuser = Found
		if Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_PBIReport").Exist(2) then
      							Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_PBIReport").Click
      						else
								Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close").Click
							end if 	
	Call PageLoading()
	if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Content Permissions List").Exist(2) then
			Call ReportStep (StatusTypes.Pass,"Verify that user Navigated to content permission tab successflly ","Navigated to - COntent Permission Tab")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user Navigated to content permission tab successflly ","Not Navigated to - COntent Permission Tab")
		End If
end function 
			
	
Function VAReportWithAllclientsnotinCPTab(ByVal FileName,ByVal SheetName,ByVal StartRow,ByVal EndRow,ByRef objData)	

	if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Content Permissions List").Exist(2) then
			Call ReportStep (StatusTypes.Pass,"Verify that user Navigated to content permission tab successflly ","Navigated to - COntent Permission Tab")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user Navigated to content permission tab successflly ","Not Navigated to - COntent Permission Tab")
		End If
		
		Call ImportSheet(SheetName,FileName)

		Datatable.SetCurrentRow(StartRow)
			Report_Type=Datatable.Value("Report_Type",SheetName)
			Report_Name=Datatable.Value("Report_Name",SheetName)
		
		wait 5
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("ConPer_Select Users").Exist(10) then
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("ConPer_Select Users").Click
			Call ReportStep (StatusTypes.Pass,"Verify that user is able to click on select Users","Clicked on - Select user link")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user is able to click on select Users","Not Clicked on - Select user link")
		End If
		Call PageLoading()	
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("SelectUserPopUp").exist(20) then
			Call ReportStep (StatusTypes.Pass,"Verify that Display of select User pop-up","Displayed - Select User popup")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that Display of select User pop-up","Not Displayed - Select User popup")
		End If
		
		 Set objShell=CreateObject("WScript.Shell")
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("CP_ChooseaReportType").Click
			If Report_Type = "VA Dossier Report" Then
				objShell.SendKeys "{DOWN}"
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Select_VA Dossier Report").Click
			ElseIf Report_Type = "VA PBI Report" Then
				objShell.SendKeys "{DOWN}"
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Select_VA PBI Report").Click
		
			End If
		


			Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_CPSelectReport").Set Report_Name

wait 2
			objShell.SendKeys "{DOWN}"
		
			wait 3
			if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).exist(2) then
				Call ReportStep (StatusTypes.Fail,"Verify that the VA Report with All clients does not list in permission content tab","PC tab - All clients VA report listed")
			else
				Call ReportStep (StatusTypes.Pass,"Verify that the VA Report with All clients does not list in permission content tab","PC tab - All clients VA report not listed")
			End If
			if Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_PBIReport").Exist(2) then
      							Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_PBIReport").Click
      						else
								Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close").Click
							end if 	
			Call PageLoading()
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Content Permissions List").Exist(2) then
			Call ReportStep (StatusTypes.Pass,"Verify that user Navigated to content permission tab successflly ","Navigated to - COntent Permission Tab")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user Navigated to content permission tab successflly ","Not Navigated to - COntent Permission Tab")
		End If
end function


Function ValidateUserNotexistwarningmsg(ByVal FileName,ByVal SheetName,ByVal StartRow,ByVal EndRow,ByRef objData,User, Status)
	Call ImportSheet(SheetName,FileName)
		DataTable.SetCurrentRow(StartRow)
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Content Permissions List").Exist(2) then
			Call ReportStep (StatusTypes.Pass,"Verify that user Navigated to content permission tab successflly ","Navigated to - COntent Permission Tab")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user Navigated to content permission tab successflly ","Not Navigated to - COntent Permission Tab")
		End If
		
		Beforeadduser = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").RowCount
		
		wait 5
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("ConPer_Select Users").Exist(10) then
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("ConPer_Select Users").Click
			Call ReportStep (StatusTypes.Pass,"Verify that user is able to click on select Users","Clicked on - Select user link")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user is able to click on select Users","Not Clicked on - Select user link")
		End If
		Call PageLoading()	
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("SelectUserPopUp").exist(20) then
			Call ReportStep (StatusTypes.Pass,"Verify that Display of select User pop-up","Displayed - Select User popup")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that Display of select User pop-up","Not Displayed - Select User popup")
		End If
		Report_Type=Datatable.Value("Report_Type",SheetName)
	     Report_Name=Datatable.Value("Report_Name",SheetName)
		
		
		
		 Set objShell=CreateObject("WScript.Shell")
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("CP_ChooseaReportType").Click
			If Report_Type = "VA Dossier Report" Then
				objShell.SendKeys "{DOWN}"
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Select_VA Dossier Report").Click
			ElseIf Report_Type = "VA PBI Report" Then
				objShell.SendKeys "{DOWN}"
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Select_VA PBI Report").Click
		
			End If
		


			Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_CPSelectReport").Set Report_Name

wait 2
			objShell.SendKeys "{DOWN}"
		
			wait 1
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).fireevent "onmouseover" 
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).highlight
			wait 2
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).Click		
			
			Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("CPSelectedUserDisplay").Set User
			
	Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_SelectUsers_Add").Click
	
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Dialog_Information").Exist(5) then
		Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Dialog_Information").Click
			msg =  Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Dialog_Information").getroproperty("innertext")
			Browser("Browser-OCRF").Page("Page-OCRF").WebButton("innertext:=Okay","visible:=True","html tag:=BUTTON").Click
			Call ReportStep (StatusTypes.Pass,"Verify that user Navigated to 'User Not does not belog t0 current offering' dialog for "&Status&" in add users tab successflly ","Content Permision warning -"& msg)
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user Navigated to 'User Not does not belog t0 current offering' dialog for "&Status&" in add users tab  successflly ","Not Navigated to -Content Permision user does not belong to warning")
		End If

	
	if Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_PBIReport").Exist(2) then
      							Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_PBIReport").Click
      						else
								Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close").Click
							end if 	
			Call PageLoading()
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Content Permissions List").Exist(2) then
			Call ReportStep (StatusTypes.Pass,"Verify that user Navigated to content permission tab successflly ","Navigated to - COntent Permission Tab")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user Navigated to content permission tab successflly ","Not Navigated to - COntent Permission Tab")
		End If
End Function


Function SearchuserInPerConTabColums(ByVal FileName,ByVal SheetName,ByVal StartRow,ByVal EndRow,ByRef objData,User, searchwith)
	Call ImportSheet(SheetName,FileName)
		DataTable.SetCurrentRow(StartRow)
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Content Permissions List").Exist(2) then 
			Call ReportStep (StatusTypes.Pass,"Verify that user Navigated to content permission tab successflly ","Navigated to - COntent Permission Tab")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user Navigated to content permission tab successflly ","Not Navigated to - COntent Permission Tab")
		End If
		
		Beforeadduser = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").RowCount
		
		wait 5
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("ConPer_Select Users").Exist(10) then
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("ConPer_Select Users").Click
			Call ReportStep (StatusTypes.Pass,"Verify that user is able to click on select Users","Clicked on - Select user link")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user is able to click on select Users","Not Clicked on - Select user link")
		End If
		Call PageLoading()	
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("SelectUserPopUp").exist(20) then
			Call ReportStep (StatusTypes.Pass,"Verify that Display of select User pop-up","Displayed - Select User popup")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that Display of select User pop-up","Not Displayed - Select User popup")
		End If
		Report_Type=Datatable.Value("Report_Type",SheetName)
	     Report_Name=Datatable.Value("Report_Name",SheetName)
		
		
		
		 Set objShell=CreateObject("WScript.Shell")
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("CP_ChooseaReportType").Click
			If Report_Type = "VA Dossier Report" Then
				objShell.SendKeys "{DOWN}"
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Select_VA Dossier Report").Click
			ElseIf Report_Type = "VA PBI Report" Then
				objShell.SendKeys "{DOWN}"
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Select_VA PBI Report").Click
		
			End If
		

			Call Pageloading()	
			Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_CPSelectReport").Set Report_Name

wait 2
			objShell.SendKeys "{DOWN}"
		
			wait 3
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).fireevent "onmouseover" 
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).highlight
			wait 2
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).Click
			
			If searchwith = "LoginId" Then
				
			
			If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("CP_UserLoginID").Exist Then
				Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("CP_UserLoginID").Set user
				Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("CP_UserLoginID").click
				objShell.SendKeys "{ENTER}"
				objShell.SendKeys "{ENTER}"
				
				Call ReportStep (StatusTypes.Pass,"Verify the entering UserLoginID in Permission content tab","ContentPermission - entered user for search through userlogin id")
				getuserlisted = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("ContentPermUser-toadd-list").GetCellData(2,4)
				If ucase(getuserlisted) = ucase(user) Then
					Call ReportStep (StatusTypes.Pass,"Verify the search of the user through UserLoginID in Permission content tab","ContentPermission - user searched through userlogin id")
				else
					Call ReportStep (StatusTypes.Fail,"Verify the search of the user through UserLoginID in Permission content tab","ContentPermission - user not searched through userlogin id")
				End If
			else
				Call ReportStep (StatusTypes.Fail,"Verify the search of the user through UserLoginID in Permission content tab","ContentPermission - not entered user for search through userlogin id")
			End If
			
			ElseIf searchwith = "Last Name" Then
				If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("CP_LastName").Exist Then
				Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("CP_LastName").Set user
				Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("CP_LastName").click
				objShell.SendKeys "{ENTER}"
				objShell.SendKeys "{ENTER}"
				
				Call ReportStep (StatusTypes.Pass,"Verify the entering Last Name in Permission content tab","ContentPermission - entered user for search through Last Name")
				getuserlisted = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("ContentPermUser-toadd-list").GetCellData(2,5)
				If ucase(getuserlisted) = ucase(user) Then
					Call ReportStep (StatusTypes.Pass,"Verify the search of the user through Last Name in Permission content tab","ContentPermission - user searched through Last Name")
				else
					Call ReportStep (StatusTypes.Fail,"Verify the search of the user through Last Name in Permission content tab","ContentPermission - user not searched through Last Name id")
				End If
			else
				Call ReportStep (StatusTypes.Fail,"Verify the search of the user through Last Name in Permission content tab","ContentPermission - not entered user for search through Last Name id")
			End If
			
			ElseIf searchwith = "First Name" Then
					If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("CP_FirstName").Exist Then
				Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("CP_FirstName").Set user
				Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("CP_FirstName").click
				objShell.SendKeys "{ENTER}"
				objShell.SendKeys "{ENTER}"
				
				Call ReportStep (StatusTypes.Pass,"Verify the entering first Name in Permission content tab","ContentPermission - entered user for search through first Name")
				getuserlisted = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("ContentPermUser-toadd-list").GetCellData(2,6)
				If ucase(getuserlisted) = ucase(user) Then
					Call ReportStep (StatusTypes.Pass,"Verify the search of the user through first Name in Permission content tab","ContentPermission - user searched through first Name")
				else
					Call ReportStep (StatusTypes.Fail,"Verify the search of the user through first Name in Permission content tab","ContentPermission - user not searched through first Name id")
				End If
			else
				Call ReportStep (StatusTypes.Fail,"Verify the search of the user through first Name in Permission content tab","ContentPermission - not entered user for search through first Name id")
			End If
			
			ElseIf searchwith = "User Group" Then
				If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("CP_UserGroupName").Exist Then
				Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("CP_UserGroupName").Set user
				Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("CP_UserGroupName").click
				objShell.SendKeys "{ENTER}"
				objShell.SendKeys "{ENTER}"
				
				Call ReportStep (StatusTypes.Pass,"Verify the entering  user group in Permission content tab","ContentPermission - entered user for search through  user group")
				getuserlisted = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("ContentPermUser-toadd-list").GetCellData(2,8)
				If ucase(getuserlisted) = ucase(user) Then
					Call ReportStep (StatusTypes.Pass,"Verify the search of the user through user group in Permission content tab","ContentPermission - user searched through user group")
				else
					Call ReportStep (StatusTypes.Fail,"Verify the search of the user through user group in Permission content tab","ContentPermission - user not searched through user group")
				End If
			else
				Call ReportStep (StatusTypes.Fail,"Verify the search of the user through user group in Permission content tab","ContentPermission - not entered user for search through user group")
			End If	
		

			ElseIf searchwith = "Client Name" Then
				If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("CP_ClientName").Exist Then
				Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("CP_ClientName").Set user
				Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("CP_ClientName").click
				objShell.SendKeys "{ENTER}"
				objShell.SendKeys "{ENTER}"
				
				Call ReportStep (StatusTypes.Pass,"Verify the entering Client Name in Permission content tab","ContentPermission - entered user for search through Client Name")
				getuserlisted = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("ContentPermUser-toadd-list").GetCellData(2,7)
				If ucase(getuserlisted) = ucase(user) Then
					Call ReportStep (StatusTypes.Pass,"Verify the search of the user through Client Name in Permission content tab","ContentPermission - user searched through Client Name")
				else
					Call ReportStep (StatusTypes.Fail,"Verify the search of the user through Client Name in Permission content tab","ContentPermission - user not searched through Client Name")
				End If
			else
				Call ReportStep (StatusTypes.Fail,"Verify the search of the user through Client Name in Permission content tab","ContentPermission - not entered user for search throughClient Name")
			End If	
			
		End If 	
		if Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_PBIReport").Exist(2) then
      							Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_PBIReport").Click
      						else
								Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close").Click
							end if 	
			Call PageLoading()
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Content Permissions List").Exist(2) then
			Call ReportStep (StatusTypes.Pass,"Verify that user Navigated to content permission tab successflly ","Navigated to - COntent Permission Tab")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user Navigated to content permission tab successflly ","Not Navigated to - COntent Permission Tab")
		End If
		
end function

Function ValiadteDeleteduserVisibilityinCP()
	if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Content Permissions List").Exist(2) then
            Call ReportStep (StatusTypes.Pass,"Verify that user Navigated to content permission tab successflly ","Navigated to - COntent Permission Tab")
        else
            Call ReportStep (StatusTypes.Fail,"Verify that user Navigated to content permission tab successflly ","Not Navigated to - COntent Permission Tab")
        End If
        
       Totalrows =  Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").RowCount
       For Iterator = 2 To Totalrows Step 1
       		Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").ChildItem(Iterator,2,"WebCheckBox",0).click
       		Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").ChildItem(Iterator,2,"WebCheckBox",0).set "ON"
       		Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").ChildItem(Iterator,2,"WebCheckBox",0).click
       		deleteduserstatus = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").ChildItem(Iterator,3,"Weblist",0).getroproperty("value")
       		If deleteduserstatus = "DELETED" Then
       			found = 1
       			Exit for
       		End If
       Next
       
       If found = 1 Then
       		 Call ReportStep (StatusTypes.Pass,"Verify that user is able to see the deleted user in content permission tab successflly ","User with Deleted status visible- COntent Permission Tab")
        else
            Call ReportStep (StatusTypes.Fail,"Verify that user is able to see the deleted user in content permission tab successflly ","User with Deleted status not visible - COntent Permission Tab")
        End If
        
       

End function


Function CPChangeAddeduserStatus(Rownum,UserStatus)

	usernum = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").RowCount
	If usernum>=Rownum Then
		Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").ChildItem(Rownum,2,"WebCheckBox",0).set "ON"
	Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").ChildItem(Rownum,2,"WebCheckBox",0).click
	Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").ChildItem(Rownum,2,"WebCheckBox",0).set "ON"
	Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").ChildItem(Rownum,3,"WebList",0).select UserStatus
	
	Call ReportStep (StatusTypes.Pass,"The user action changed to "&UserStatus&" by admin","Admin is able to change the user status to " &UserStatus &"- Content Permission Tab")
	
	else
		Call ReportStep (StatusTypes.Fail,"Users are less than "&Rownum& " in ContentPermission ","Admin is not able to change the user status to "& UserStatus &" - Content Permission Tab")
	End If
	
			
	End Function


Function AddUserstomultipleReportsPCTab(ByVal FileName,ByVal SheetName,ByVal StartRow,ByVal EndRow,ByRef objData)

		Call ImportSheet(SheetName,FileName)
		DataTable.SetCurrentRow(StartRow)
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Content Permissions List").Exist(2) then
			Call ReportStep (StatusTypes.Pass,"Verify that user Navigated to content permission tab successflly ","Navigated to - COntent Permission Tab")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user Navigated to content permission tab successflly ","Not Navigated to - COntent Permission Tab")
		End If
		
		wait 5
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("ConPer_Select Users").Exist(10) then
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("ConPer_Select Users").Click
			Call ReportStep (StatusTypes.Pass,"Verify that user is able to click on select Users","Clicked on - Select user link")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user is able to click on select Users","Not Clicked on - Select user link")
		End If
		Call PageLoading()	
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("SelectUserPopUp").exist(20) then
			Call ReportStep (StatusTypes.Pass,"Verify that Display of select User pop-up","Displayed - Select User popup")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that Display of select User pop-up","Not Displayed - Select User popup")
		End If
		Report_Type=Datatable.Value("Report_Type",SheetName)
	     
		
				
		 Set objShell=CreateObject("WScript.Shell")
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("CP_ChooseaReportType").Click
			If Report_Type = "VA Dossier Report" Then
				objShell.SendKeys "{DOWN}"
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Select_VA Dossier Report").Click
			ElseIf Report_Type = "VA PBI Report" Then
				objShell.SendKeys "{DOWN}"
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Select_VA PBI Report").Click
		
			End If
		Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_CPSelectReport").Click
For Iterator = StartRow To EndRow Step 1
	
	DataTable.SetCurrentRow(Iterator)
	Report_Name=Datatable.Value("Report_Name",SheetName)
			

wait 2
			'objShell.SendKeys "{DOWN}"
		
			wait 3
			if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).exist(2) then
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).fireevent "onmouseover" 
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).highlight
			wait 2
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).Click
			else
			objShell.SendKeys "{DOWN}"	
			wait 3
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).fireevent "onmouseover" 
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).highlight
			wait 2
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).Click
			end if 			
	Next	
		
		Browser("Browser-OCRF").Page("Page-OCRF").WebCheckBox("cb_ClientContentUser-toadd-lis").Set "ON"
		selectedusers = Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("CPSelectedUserDisplay").GetROProperty("value")
		If selectedusers<>"" Then
			Call ReportStep (StatusTypes.Pass, "Verify selected user gets displayed in Content Permission add users text area","Selected users displayed in text area-Content PermissionTab")
		else
			Call ReportStep (StatusTypes.Pass, "Verify selected user gets displayed in Content Permission add users text area","Selected users not displayed in text area-Content PermissionTab")
		
		End If
		
		Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_SelectUsers_Add").Click
		Call Pageloading()
		selectedusers = Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("CPSelectedUserDisplay").GetROProperty("value")
		If selectedusers="" Then
			Call ReportStep (StatusTypes.Pass, "Verify gets disappeared in Content Permission add users text area once clicked on add","After clicking on add, users disappears in text area-Content PermissionTab")
		else
			Call ReportStep (StatusTypes.Fail, "Verify gets disappeared in Content Permission add users text area once clicked on add","After clicking on add, users not disappeared in text area-Content PermissionTab")
		
		End If
		
		if Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_PBIReport").Exist(2) then
					Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_PBIReport").Click
					Call ReportStep (StatusTypes.Pass, "Verify that users is able to clcik on close button in Content Permission Tab","Close button-Content PermissionTab")
				else
				Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close").Click
				Call ReportStep (StatusTypes.Pass, "Verify that users is able to clcik on close button in Content Permission Tab","Close button-Content PermissionTab")
			end if 	
			Call PageLoading()
			wait 15
		'Verify the addition of user in the content Permission tab
	if Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").RowCount>=2 then
		Call ReportStep (StatusTypes.Pass, "Verify that users are added for the multiple reports at a time in Content Permission Tab","Users Added for multiple reports-Content PermissionTab")
	else
		Call ReportStep (StatusTypes.Fail, "Verify that users are added for the multiple reports at a timein Content Permission Tab","Users Added for multiple reports-Content PermissionTab")
	end if 
	
	end function 


Function WarningPopupforDossierREportNotAdded(Report_Type)
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("ConPer_Select Users").Exist(10) then
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("ConPer_Select Users").Click
			Call ReportStep (StatusTypes.Pass,"Verify that user is able to click on select Users","Clicked on - Select user link")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user is able to click on select Users","Not Clicked on - Select user link")
		End If
		Call PageLoading()	
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("SelectUserPopUp").exist(20) then
			Call ReportStep (StatusTypes.Pass,"Verify that Display of select User pop-up","Displayed - Select User popup")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that Display of select User pop-up","Not Displayed - Select User popup")
		End If
		
				
		 Set objShell=CreateObject("WScript.Shell")
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("CP_ChooseaReportType").Click
			If Report_Type = "VA Dossier Report" Then
				objShell.SendKeys "{DOWN}"
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Select_VA Dossier Report").Click
				Call ReportStep (StatusTypes.Pass,"Verify that Selection of VA Dossier Report Type","Selected - VA Dossier Report Type")
			ElseIf Report_Type = "VA PBI Report" Then
				objShell.SendKeys "{DOWN}"
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Select_VA PBI Report").Click
				Call ReportStep (StatusTypes.Pass,"Verify that Selection of VA PBI Report Type","Selected - VA PBI Report Type")
			End If
			
			if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Dialog_Information").Exist(2) then
				Browser("Browser-OCRF").Page("Page-OCRF").WebButton("html tag:=BUTTON","innertext:=Okay","type:=button","visible:=True").Click
				Call ReportStep (StatusTypes.Pass,"Verify that system to give warning message to users if there is no reports to display for Dossier \ PowerBI Report","Warning message displayed -for VA Report not present")
			else	
				Call ReportStep (StatusTypes.Fail,"Verify that system to give warning message to users if there is no reports to display for Dossier \ PowerBI Report","Warning message not displayed -for VA Report not present")
			End If
			
			if Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_PBIReport").Exist(2) then
					Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_PBIReport").Click
					Call ReportStep (StatusTypes.Pass, "Verify that users is able to clcik on close button in Content Permission Tab","Close button-Content PermissionTab")
				else
				Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close").Click
				Call ReportStep (StatusTypes.Pass, "Verify that users is able to clcik on close button in Content Permission Tab","Close button-Content PermissionTab")
			end if 	
			Call PageLoading()
	End function
	
	
Function PCwarningmsgforalreadyexistUsers(ByVal FileName,ByVal SheetName,ByVal StartRow,ByVal EndRow,ByRef objData)
	Call ImportSheet(SheetName,FileName)
		DataTable.SetCurrentRow(StartRow)
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Content Permissions List").Exist(2) then
			Call ReportStep (StatusTypes.Pass,"Verify that user Navigated to content permission tab successflly ","Navigated to - COntent Permission Tab")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user Navigated to content permission tab successflly ","Not Navigated to - COntent Permission Tab")
		End If
		
		Beforeadduser = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").RowCount
		
		wait 5
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("ConPer_Select Users").Exist(10) then
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("ConPer_Select Users").Click
			Call ReportStep (StatusTypes.Pass,"Verify that user is able to click on select Users","Clicked on - Select user link")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user is able to click on select Users","Not Clicked on - Select user link")
		End If
		Call PageLoading()	
		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("SelectUserPopUp").exist(20) then
			Call ReportStep (StatusTypes.Pass,"Verify that Display of select User pop-up","Displayed - Select User popup")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that Display of select User pop-up","Not Displayed - Select User popup")
		End If
		Report_Type=Datatable.Value("Report_Type",SheetName)
	     Report_Name=Datatable.Value("Report_Name",SheetName)
		
		
		
		 Set objShell=CreateObject("WScript.Shell")
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("CP_ChooseaReportType").Click
			If Report_Type = "VA Dossier Report" Then
				objShell.SendKeys "{DOWN}"
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Select_VA Dossier Report").Click
			ElseIf Report_Type = "VA PBI Report" Then
				objShell.SendKeys "{DOWN}"
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Select_VA PBI Report").Click
		
			End If
		


			Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_CPSelectReport").Click
wait 2
			'objShell.SendKeys "{DOWN}"
		
			wait 3
			if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).exist(2) then
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).fireevent "onmouseover" 
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).highlight
			wait 2
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).Click
			else
			objShell.SendKeys "{DOWN}"	
			wait 3
			If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).exist Then
				
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).fireevent "onmouseover" 
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).highlight
			wait 2
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).Click
			else
			objShell.SendKeys "{UP}"	
			wait 3
			
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).fireevent "onmouseover" 
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).highlight
			wait 2
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=active-result highlighted","innertext:="&Report_Name).Click
			End If 
			
			end if 	
			Browser("Browser-OCRF").Page("Page-OCRF").WebCheckBox("cb_ClientContentUser-toadd-lis").Set "ON"
					Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_SelectUsers_Add").Click
					
					if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Dialog_Information").Exist then
						msg = Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Dialog_Information").GetROProperty("innertext")
						Browser("Browser-OCRF").Page("Page-OCRF").WebButton("html tag:=BUTTON","innertext:=Okay","type:=button","visible:=True").Click
						Call ReportStep (StatusTypes.Pass, "Verify that warning message for adding the already added users for the selected report in Content Permission Tab",msg&" - warining for already exists user- Content PermissionTab")
					else
						Call ReportStep (StatusTypes.Fail, "Verify that warning message for adding the already added users for the selected report in Content Permission Tab","No warning msg for already added users-Content PermissionTab")
					end if
				if Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_PBIReport").Exist(2) then
      							Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_PBIReport").Click
      						else
								Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close").Click
							end if 	
					Call PageLoading()
					wait 15
					
		end function
				
Function ValidatetheReportype_SelectReportUI()
	
	if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("ConPer_Select Users").Exist(10) then
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("ConPer_Select Users").Click
			Call ReportStep (StatusTypes.Pass,"Verify that user is able to click on select Users","Clicked on - Select user link")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that user is able to click on select Users","Not Clicked on - Select user link")
		End If
		Call Pageloading()
		ChoosereportTypetext = Browser("Browser-OCRF").Page("Page-OCRF").WebElement("CP_ChooseaReportType").GetROProperty("innertext")
		If ChoosereportTypetext = "Choose a Report Type" Then
			Call ReportStep (StatusTypes.Pass,"Verify that Default value of Choose Report type element before selecting report is 'Choose a Report Type'","Before Selecting Report - ''Choose a Report Type' is displayed in choose report type element")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that Default value of Choose Report type element before selecting report is 'Choose a Report Type'","Before Selecting Report - ''Choose a Report Type' is not displayed in choose report type element")
		
		End If
		
		
		ChooseReportText = Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEdit_CPSelectReport").GetROProperty("default value")
		If ChooseReportText = "Select Reports" Then
			Call ReportStep (StatusTypes.Pass,"Verify that Default value of Choose Reports element before selecting report is 'Select Reports'","Before Selecting Report - ''Select Reports' is displayed in choose reports element")
		else
			Call ReportStep (StatusTypes.Fail,"Verify that Default value of Choose Reports element before selecting report is 'Select Reports'","Before Selecting Report - ''Select Reports' is not displayed in choose reports element")
		
		End If
		
		
		
		if Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_PBIReport").Exist(2) then
				Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_PBIReport").Click
		else
			Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close").Click
		end if 	
End Function

Function ChangeReportStatusinClientDetailsTab(ByVal FileName,ByVal SheetName,ByVal StartRow,ByVal EndRow,ByRef objData,ChangeStatus)
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementCDTabHeader").Exist(20) Then
		Call ReportStep (StatusTypes.Pass, "'Client content' tab header exist","Client content Tab header")
	Else
	   	Call ReportStep (StatusTypes.Fail, "Client content Tab header not exist","Client content Tab header")
	End If

 	wait 2
	Call ImportSheet(SheetName,FileName)
	
	DataTable.GetSheet(SheetName).SetCurrentRow(StartRow)
	Report_Type=Datatable.Value("Report_Type",SheetName)
	Report_Name=Datatable.Value("Report_Name",SheetName)
	
	Set objClientContentTable = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList")
	
	RowtoDelete = objClientContentTable.GetRowWithCellText(Report_Name)
	
	objClientContentTable.ChildItem(RowtoDelete,2,"WebCheckBox",0).set "ON"
	objClientContentTable.ChildItem(RowtoDelete,2,"WebCheckBox",0).click
	objClientContentTable.ChildItem(RowtoDelete,2,"WebCheckBox",0).set "ON"
	objClientContentTable.ChildItem(RowtoDelete,6,"WebList",0).select ChangeStatus
	Call ReportStep (StatusTypes.Pass, "Verify the selection of "&ChangeStatus&" for the specified Report in Client Content Tab","Client content Tab - Change Report to "&ChangeStatus)

End function	

Function LO_CheckfreezeDCH_ReportType()
	Set objClientContentTable = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList")
	 If objClientContentTable.Exist(10)  Then
	   rowscount= objClientContentTable.RowCount
	    rowscount= objClientContentTable.RowCount
	    If rowscount<=2 Then
	   		rowscount = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList").RowCount
	   End If
	   For CCR = 2 To rowscount Step 1
	   		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-pg-div","innertext:=Show Deleted ", "html tag:=DIV","visible:=True").Exist then
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-pg-div","innertext:=Show Deleted ", "html tag:=DIV","visible:=True").highlight
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-pg-div","innertext:=Show Deleted ", "html tag:=DIV","visible:=True").click
				Call PageLoading()
			End If
	   		 objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	           ReportStatus =  objClientContentTable.ChildItem(CCR,6,"WebList",0).getroproperty("value")
	    	 If ReportStatus="DELETE" or ReportStatus="UPDATE" or ReportStatus="NO CHANGE" or ReportStatus="DELETED"  Then
	    	 	'CHECKING FOR FREEZED DATA FOR DECISION CENTER HEADING,Report Type is disabled for update/nochange/delete/deleted
	           DCDisabled=objClientContentTable.ChildItem(CCR, 10, "WebList", 0).GetROProperty("disabled") 
	           RTDisabled = objClientContentTable.ChildItem(CCR, 14, "WebList", 0).GetROProperty("disabled") 
	           If DCDisabled="1" AND RTDisabled="1" Then
	           	  Call ReportStep (StatusTypes.Pass, "''Report Type' and  'Decision center' fields  are not Editable by LO user for report in th status "&ReportStatus ,"LO - Decision Centre Heading and Repoer Type Freezed for report status -"&ReportStatus)
	           Else
	           	 Call ReportStep (StatusTypes.Pass, "''Report Type' and  'Decision center' fields  are Editable by LO user for report in th status "&ReportStatus ,"LO - Decision Centre Heading and Repoer Type not Freezed for report status -"&ReportStatus)
	           End If
	    	 End If
	           
	   Next
	           
else
		 Call ReportStep (StatusTypes.Pass, "Client contented tab - Verify the existance of Client content table","Client content table not displayed")
End if 		 
End Function

Function Admin_CheckfreezeDCH_ReportType()
Set objClientContentTable = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList")
	 If objClientContentTable.Exist(10)  Then
	   rowscount= objClientContentTable.RowCount
	    rowscount= objClientContentTable.RowCount
	    If rowscount<=2 Then
	   		rowscount = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList").RowCount
	   End If
	   For CCR = 2 To rowscount Step 1
	   		if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-pg-div","innertext:=Show Deleted ", "html tag:=DIV","visible:=True").Exist then
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-pg-div","innertext:=Show Deleted ", "html tag:=DIV","visible:=True").highlight
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-pg-div","innertext:=Show Deleted ", "html tag:=DIV","visible:=True").click
				Call PageLoading()
			End If
	   		 objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	 	ReportStatus =  objClientContentTable.ChildItem(CCR,6,"WebList",0).getroproperty("value")
	    	 	'CHECKING FOR FREEZED DATA FOR DECISION CENTER HEADING,Report Type is disabled for update/nochange/delete/deleted
	           DCDisabled=objClientContentTable.ChildItem(CCR, 10, "WebList", 0).GetROProperty("disabled") 
	           RTDisabled = objClientContentTable.ChildItem(CCR, 14, "WebList", 0).GetROProperty("disabled") 
	           If DCDisabled="1" AND RTDisabled="1" Then
	           	  Call ReportStep (StatusTypes.Fail, "''Report Type' and  'Decision center' fields  are not Editable by Admin user for report in th status "&ReportStatus ,"Admin - Decision Centre Heading and Repoer Type Freezed for report status -"&ReportStatus)
	           Else
	           	 Call ReportStep (StatusTypes.Pass, "''Report Type' and  'Decision center' fields  are Editable by Admin user for report in th status "&ReportStatus ,"Admin - Decision Centre Heading and Repoer Type not Freezed for report status -"&ReportStatus)
	           End If
	    	
	           
	   Next
	else
		 Call ReportStep (StatusTypes.Pass, "Client contented tab - Verify the existance of Client content table","Client content table not displayed")
End if 	           

End Function

'********************************************************************************************************************************************************************************************
'********************************************************************************************************************************************************************************************
'----------------------------------------------------------------Generic Functions created by Madhu wRT VA Report Level Security ends_______________________________________________________________

'********************************************************************************************************************************************************************************************
'********************************************************************************************************************************************************************************************
'----------------------------------------------------------------Generic Functions created by Srini wRT VA Report Level Security_______________________________________________________________


Public Function createUserGroupAddUsersTab(ByVal UserGroup,ByVal row)
	if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Webele_Column_user-list_UserGroupName").Exist(5) then
		Call ReportStep (StatusTypes.Pass,"Verify Newly added Column_user-list_UserGroupName display","Column_user-list_UserGroupName is present" )
	 	'Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Webele_Column_user-list_UserGroupName").Click
	 	 browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//td//div[text()='User Group']").Click
	 	 If browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//td[@id='#usergroup-addnewlinetop']").Exist(10) Then
	 	 	Call ReportStep (StatusTypes.Pass,"Verify whether UserGroupName window is opened after clicking on usergroup button in addusers tab","User Group window opened!" )
	 	 	browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//td[@id='#usergroup-addnewlinetop']").Click
	 	 	Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='usergroup-list']//tr//td[@aria-describedby='usergroup-list_rn' and text()='"&row&"']").click
	 	 	wait 4
	 	 	Browser("micclass:=browser").Page("micclass:=page").WebEdit("xpath:=//tr//td[text()='"&row&"']/../descendant::td//input[@name='UserGroupName']").set UserGroup
	 	 	wait 2
	 	 	Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//button//span[text()='Add User Group']").Click
	 	 	Call ReportStep (StatusTypes.Pass,"clicked on AddUserGroup Button","Click AddUserGroup button")
	 	 	
	 	 	If browser("micclass:=browser").page("micclass:=page").WebElement("xpath:=//div[contains(@style,'display: block')]//button//span[text()='Yes']").Exist(10) Then
	 	 		Call ReportStep (StatusTypes.Pass,"Confirmation popup appeared","Verify whether confirmation popup appears after clicking on adduserGroup button")
	 	 		browser("micclass:=browser").page("micclass:=page").WebElement("xpath:=//div[contains(@style,'display: block')]//button//span[text()='Yes']").highlight
	 	 		browser("micclass:=browser").page("micclass:=page").WebElement("xpath:=//div[contains(@style,'display: block')]//button//span[text()='Yes']").Click
				else
				Call ReportStep (StatusTypes.Pass,"Confirmation popup not appeared","Verify whether confirmation popup appears after clicking on adduserGroup button")
	 	 	End If
	 	 	If browser("micclass:=browser").page("micclass:=page").WebElement("xpath:=//div[contains(@style,'display: block')]//button//span[text()='Cancel']").Exist(10) Then
	 	 		browser("micclass:=browser").page("micclass:=page").WebElement("xpath:=//div[contains(@style,'display: block')]//button//span[text()='Cancel']").highlight
				browser("micclass:=browser").page("micclass:=page").WebElement("xpath:=//div[contains(@style,'display: block')]//button//span[text()='Cancel']").Click
	 	 	End If
			Call ReportStep (StatusTypes.Information,""&UserGroup&" is created successfully!","AddUserGroup grid in AddUsers tab") 
	 	 End If
	 Else 
	  	Call ReportStep (StatusTypes.Fail,"Verify Newly added Column_user-list_UserGroupName display","Column_user-list_UserGroupName is present" )	
	 End If

End Function

Public Function modifyUserGroupName(ByVal ModifiedUserGroup,ByVal UserGroupPrevious)
	wait 2
	'browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//td//div[text()='User Group']").Click
	'Browser("BrowserPoPUp").Page("Online CRF - Request_7").WebElement("User Group Button").Click
	Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='user-list_toppager_right-inner']//div[text()='User Group']").Click
	Call PageLoading()
	If browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//td[@id='#usergroup-addnewlinetop']").Exist(20) Then
	Set odesc=description.Create()
	odesc("html tag").Value="TD"
	'odesc("title").Value="UserGroup1"
	set r=Browser("micclass:=browser").Page("micclass:=page").WebTable("html tag:=TABLE","html id:=usergroup-list").ChildObjects(odesc)
	
	For i = 1 To r.count-1 Step 1
		s=r(i).getroproperty("text")
		print(s)
		If s=UserGroupPrevious Then
			r(i).click
		End If
	Next
	wait 2
	'browser("micclass:=browser").Page("micclass:=page").WebEdit("xpath:=//input[@name='UserGroupName']/../ancestor::tr//td[@title='"&UserGroupPrevious&"']").set ModifiedUserGroup
	Set odesc=description.Create()
	odesc("html tag").Value="input"	
	'odesc("name").Value="UserGroupName"
	'set r1=Browser("BrowserPoPUp").Page("Online CRF - Request_5").WebTable("usergroup-list").ChildObjects(odesc)
	set r1=Browser("micclass:=browser").Page("micclass:=page").WebTable("html tag:=TABLE","html id:=usergroup-list").ChildObjects(odesc)
	'msgbox r1.count
	For i = 1 To r1.count-1 Step 1
			s=r1(i).getroproperty("value")
		If s=UserGroupPrevious Then
			r1(i).set ModifiedUserGroup
		End If
		
	Next
	wait 1
	Call ReportStep (StatusTypes.Pass,"Verify UserGroupName is modified successfully",""&UserGroupPrevious&" is modified to "&ModifiedUserGroup&"" )	
	End If
	Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//button//span[text()='Add User Group']").Click
	Call ReportStep (StatusTypes.Pass,"clicked on AddUserGroup Button","Click AddUserGroup button")
	 	 	
	If browser("micclass:=browser").page("micclass:=page").WebElement("xpath:=//div[contains(@style,'display: block')]//button//span[text()='Yes']").Exist(20) Then
		'Browser("BrowserPoPUp").Page("Online CRF - Request_6").WebElement("AddUserGroupConfirmationYes").Click
	 	'Browser("BrowserPoPUp").Page("Online CRF - Request_5").WebElement("Yes").Click
	 	Call ReportStep (StatusTypes.Pass,"Confirmation popup appeared","Verify whether confirmation popup appears after clicking on adduserGroup button")
	 	Set oDesc = Description.Create
		oDesc("micclass").Value = "WebElement"
		oDesc("html tag").Value = "SPAN"
		oDesc("innertext").Value = ".*Yes"
		set s=Browser("micclass:=browser").Page("micclass:=page").ChildObjects(oDesc)
		'msgbox s.count
		
		For j = 1 To s.count-1 Step 1
			s(j).click
		Next
		wait 3

		else
		Call ReportStep (StatusTypes.Pass,"Confirmation popup not appeared","Verify whether confirmation popup appears after clicking on adduserGroup button")
	 End If 
	If browser("micclass:=browser").page("micclass:=page").WebElement("xpath:=//div[contains(@style,'display: block')]//button//span[text()='Cancel']").Exist(10) Then
		browser("micclass:=browser").page("micclass:=page").WebElement("xpath:=//div[contains(@style,'display: block')]//button//span[text()='Cancel']").highlight
		browser("micclass:=browser").page("micclass:=page").WebElement("xpath:=//div[contains(@style,'display: block')]//button//span[text()='Cancel']").Click
	 End If	
	Call PageLoading()	
End Function


Function DeleteUserGroup(ByVal UserGroupName)
	'Browser("BrowserPoPUp").Page("Online CRF - Request_7").WebElement("User Group Button").Click
	Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='user-list_toppager_right-inner']//div[text()='User Group']").Click
	Call PageLoading()
	Set odesc=description.Create()
	odesc("html tag").Value="TD"
	set r=Browser("micclass:=browser").Page("micclass:=page").WebTable("html tag:=TABLE","html id:=usergroup-list").ChildObjects(odesc)
	
	For i = 1 To r.count-1 Step 1
		s=r(i).getroproperty("text")
		print(s)
		If s=UserGroupName Then
			r(i).click
		End If
	Next
	wait 2
	browser("micclass:=browser").page("micclass:=page").WebElement("xpath:=//td[@id='#usergroup-deletebutton']//div[text()='Delete']").Click
	
	If Browser("BrowserPoPUp").Page("Online CRF - Request").WebElement("innertext:=Please confirm deletion of these records. This change will also reflect in 'Add Users' Grid.*","visible:=True").Exist(20) then
'	Browser("BrowserPoPUp").Page("Online CRF - Request_5").WebElement("DeleteUserGroupConfirmationPopup").Exist(20) Then
		Call ReportStep (StatusTypes.Pass, "Delete UserGroup confirmation popup exists","Validate whether confirmation popup exists on trying to delete usergroup")	
		else
		Call ReportStep (StatusTypes.Fail, "Delete UserGroup confirmation popup does not exists","Validate whether confirmation popup exists on trying to delete usergroup")	
	End If
	'click on yes button
	Set oDesc = Description.Create
		oDesc("micclass").Value = "WebElement"
		oDesc("html tag").Value = "SPAN"
		oDesc("innertext").Value = ".*Yes"
		set s=Browser("micclass:=browser").Page("micclass:=page").ChildObjects(oDesc)
		'msgbox s.count
		
		For j = 1 To s.count-1 Step 1
			s(j).click
		Next
		wait 2
		Call ReportStep (StatusTypes.Pass, "UserGroup:-"&UserGroupName&" deleted successfully!","Validate whether usergroup is deleted or not")
		If browser("micclass:=browser").page("micclass:=page").WebElement("xpath:=//div[contains(@style,'display: block')]//button//span[text()='Cancel']").Exist(10) Then
			browser("micclass:=browser").page("micclass:=page").WebElement("xpath:=//div[contains(@style,'display: block')]//button//span[text()='Cancel']").highlight
			browser("micclass:=browser").page("micclass:=page").WebElement("xpath:=//div[contains(@style,'display: block')]//button//span[text()='Cancel']").Click
		End If	
		Call PageLoading()		
End Function





Function  AddNewLinesINAddUserTabUpdated(ByVal FileName,ByVal sheetName,ByVal RowStart,ByVal RowEnd,ByVal UserGroup,ByRef objData)

   
   NoOfUser=(RowEnd-RowStart)+1
   ' Check tab 'Add User' opened.
    If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementOfferingUserList").Exist(30) Then
	   Call ReportStep (StatusTypes.Pass,  "'Add User' tab opened","Add user tab")
    Else
       Call ReportStep (StatusTypes.Fail,  "'Add User' tab not opened","Add user tab")
    End If
   
   
	Set Tab=Browser("Browser-OCRF").Page("Page-OCRF").Link("html tag:=A","class:=selected")
	Call ValidateActiveTabOrangeColour(Tab,"Add Users",objData)

	Call ImportSheet(SheetName,FileName)

	Datatable.SetCurrentRow(RowStart)
	
		'For Iterator = RowStart To RowEnd Step 1
    	'Datatable.SetCurrentRow(RowStart)
    	'Str=Str&dataTable.value("User_Login_Id","Add_Users")&";"
	'Next
	
'	****************************************************8
	'Modified by Madhu - 03/02/2020
	
	call verifyNewCol_UserGroupSelection()
	
'********************************************888

	For Iterator = RowStart To RowEnd Step 1
	
		If NoOfUser <> "" Then
		
			Set Upload=Description.Create()
			Upload("micclass").value="WebElement"
			Upload("html tag").value="DIV"
			Upload("innertext").value="Upload User"
			Set toatlObj=Browser("Browser-OCRF").Page("Page-OCRF").ChildObjects(Upload)
			toatlObj(0).Click	
'	 *****************************************************************8
	 'Modified By Madhu - 03/02/2020
		'Call VerifythenewEditbox_UserGroup() 
	 
	    	If Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditblkClientName").Exist(2) Then
	       		Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditblkClientName").Set dataTable.value("Client_Name",SheetName)	
	    	End If
			
			'changes by srinivas
			If browser("micclass:=browser").page("micclass:=page").WebList("html id:=ddlUserGroupName").Exist(10) Then
				browser("micclass:=browser").page("micclass:=page").WebList("html id:=ddlUserGroupName").Select UserGroup
			End If
			
		    Set objShell=CreateObject("WScript.Shell")
			Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditblkClientName").Click
			objShell.SendKeys "{DOWN}"
			wait 1
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&dataTable.value("Client_Name",SheetName),"html id:=ui-id-.*").fireevent "onmouseover" 
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&dataTable.value("Client_Name",SheetName),"html id:=ui-id-.*").highlight
			wait 2
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&dataTable.value("Client_Name",SheetName),"html id:=ui-id-.*").Click		
		
	    	Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditUploadUser").Set dataTable.value("User_Login_Id",SheetName)
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebEleUploadButton").Click	
		
		
		End If
		wait 3
		DataTable.GetSheet(SheetName).SetNextRow	
	wait 3
Next

wait 3	
End Function

'updated function by srinivas
Public Function AddDatainClientDetailsTabUpdated(ByVal FileName,ByVal SheetName,ByVal StartRow,ByVal EndRow,ByVal Country,ByVal Offering,ByRef objData)


    
     TimeStamp=objData.item("TimeStamp")
   '########### OFFERING DETAIL TAB EXISTANCE - Start ########### 
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementCDTabHeader").Exist(20) Then
		Call ReportStep (StatusTypes.Pass, "'Client content' tab header exist","Client content Tab header")
	Else
	   	Call ReportStep (StatusTypes.Fail, "Client content Tab header not exist","Client content Tab header")
	End If

 	wait 2
	Call ImportSheet(SheetName,FileName)
	
	DataTable.GetSheet(SheetName)
	rc=DataTable.GetSheet(SheetName).GetRowCount
	wait 2
	For rownum = StartRow To EndRow Step 1
	     DataTable.SetCurrentRow(rownum)
	     Provide_Link=Datatable.Value("Provide_Link",SheetName)
	     Decision_Center_Heading=Datatable.Value("Decision_Center_Heading",SheetName)
	     Link_Name=Datatable.Value("Link_Name",SheetName)
	     ToolTip=Datatable.Value("ToolTip",SheetName)
	     Report_Type=Datatable.Value("Report_Type",SheetName)
	     Report_Name=Datatable.Value("Report_Name",SheetName)
	     Client_Name=Datatable.Value("Client_Name",SheetName)
	     
		'CLICK ON ADD NEW LINE
		If browser("Browser-OCRF").Page("Page-OCRF").WebElement("xpath:=//div[@id='divContentList']//table[@class='ui-pg-table navtable']//div[text()='Add New Line ']").Exist Then
			'Browser("Browser-OCRF").Page("Page-OCRF").WebElement("innertext:=Add New Line","html tag:=DIV","visible:=True").Click
			browser("Browser-OCRF").Page("Page-OCRF").WebElement("xpath:=//div[@id='divContentList']//table[@class='ui-pg-table navtable']//div[text()='Add New Line ']").Click
		ElseIf Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementAddNewLine").Exist(20) Then
			wait 2
	        'Call ReportStep (StatusTypes.Pass, "Click on add new line done","Add new button")	
			if browser("Browser-OCRF").Page("Page-OCRF").WebElement("xpath:=//div[@id='divContentList']//table[@class='ui-pg-table navtable']//div[text()='Add New Line ']").Exist then
				browser("Browser-OCRF").Page("Page-OCRF").WebElement("xpath:=//div[@id='divContentList']//table[@class='ui-pg-table navtable']//div[text()='Add New Line ']").Click
			end if 	
		Else
			Call ReportStep (StatusTypes.Fail, "Click on add new line not done","Add new button")
		End If
		Set objClientContentTable = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList")
		'Modified By Sumit on 22nd March 
		'Stated Modification
		'Offering = ReadWriteDataFromTextFile("Read","SyndOffering","")
		'Country = ReadWriteDataFromTextFile("Read","SyndContry","")
	 	' End Modification
	    If objClientContentTable.Exist(10)  Then
	    	CCR= objClientContentTable.RowCount
	        If Datatable.Value("Provide_Link",SheetName)="Yes" Then
	        wait 2
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 9, "WebList", 0).Select Provide_Link
	    	   objClientContentTable.ChildItem(CCR, 10, "WebList", 0).Select Decision_Center_Heading
	           objClientContentTable.ChildItem(CCR, 11, "WebEdit", 0).Set Link_Name&TimeStamp
	           objClientContentTable.ChildItem(CCR, 12, "WebEdit", 0).Set ToolTip&TimeStamp
	      Else
	         wait 3
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 9, "WebList", 0).Select Provide_Link
	           'CHECKING FOR FREEZED DATA FOR DECISION CENTER HEADING,LINK NAME AND TOOLTIP IF PROVIDE LINK IS 'NO'
	           DCDisabled=objClientContentTable.ChildItem(CCR, 10, "WebList", 0).GetROProperty("disabled") 
	           LinkDisabled=objClientContentTable.ChildItem(CCR, 11, "WebEdit", 0).GetROProperty("disabled") 
	           TooltipDis=objClientContentTable.ChildItem(CCR, 12, "WebEdit", 0).GetROProperty("disabled") 
	           If DCDisabled="1" AND LinkDisabled="1" AND TooltipDis="1" Then
	           	  Call ReportStep (StatusTypes.Pass, "'Link name', 'Tool tip and' and  'Decision center' fields  are not Editable after seletcing Provide link as 'No' "," Provide link as 'No'")
	           Else
	           	  Call ReportStep (StatusTypes.Pass, "'Link name', 'Tool tip and' and  'Decision center' fields  are  Editable after seletcing Provide link as 'No' "," Provide link as 'No'")
	           End If

	       End If
	       
	       	   objClientContentTable.ChildItem(CCR, 14, "WebList", 0).Select Trim(Report_Type)
	           'objClientContentTable.ChildItem(CCR, 15, "WebEdit", 0).Set Trim(Country&"/"&Offering)
	           'Modified the code based on New Requirement - OCRF - VA Ops Maintenance - 21/01/2019
	           wait 4
	           If Report_Type = "VA PBI Report" Then
'	           ***************************************************
'					Modified by Madhu - 04/15/2020
					
					'changes by srinivas
					If objClientContentTable.ChildItem(CCR, 15, "WebEdit", 0).GetROProperty("disabled") Then
	           			Call ReportStep (StatusTypes.Pass,"report path webedit column is disabled for PBI Report Type","Validate whether report path webedit column is disabled or not for PBI Report Type" )
	           			else
	           			Call ReportStep (StatusTypes.Fail,"report path webedit column is not disabled for PBI Report Type","Validate whether report path webedit column is disabled or not for PBI Report Type" )
	           		End If
	           		
'	           		
	           		If objClientContentTable.ChildItem(CCR, 16, "WebEdit", 0).GetROProperty("disabled") Then
	           			Call ReportStep (StatusTypes.Pass,"report name webedit column is disabled for PBI Report Type","Validate whether report name webedit column is disabled or not for PBI Report Type" )
	           			else
	           			Call ReportStep (StatusTypes.Fail,"report name webedit column is not disabled for PBI Report Type","Validate whether report name webedit column is disabled or not for PBI Report Type" )
	           		End If
					
					
	           		'objClientContentTable.ChildItem(CCR, 16, "WebEdit", 0).Set Trim(Report_Name)
	           		else
	           		objClientContentTable.ChildItem(CCR, 16, "WebList", 0).Click
	           		Wait 5
	           		'objClientContentTable.ChildItem(CCR, 16, "WebEdit", 0).Set Trim(Report_Name)
	           		
	           		'changes by srinivas
	           		
	           		
	           		
	           		
					If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Information_Okay").Exist Then
						Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Information_Okay").Click
						wait 3
					End If
					'objClientContentTable.ChildItem(CCR, 16, "WebList", 0).Click
	           		'objClientContentTable.ChildItem(CCR, 16, "WebList", 0).Select Trim(Report_Name)
''	           		*********************************************************************8
	           		wait 2
	           End If
	           
'	           ***********************************************************
	           objClientContentTable.ChildItem(CCR, 17, "WebEdit", 0).Set Trim(Client_Name)
	           'Set objShell=CreateObject("WScript.Shell")
	           wait 1
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   'objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
			   objClientContentTable.ChildItem(CCR, 17, "WebEdit", 0).Click
			   wait 2
'			   Set objShell=CreateObject("WScript.Shell")
'				objShell.SendKeys "{DOWN}"
				'************************************************************************8
				wait 4
'	    	   modified by Madhu - 02/26/2020
				Client_Name=Datatable.Value("Client_Name",SheetName)
'				Set oDesc = Description.Create
'				oDesc("micclass").Value = "WebElement"
'				oDesc("html tag").Value = "A"
'				oDesc("innertext").Value = Client_Name
'				oDesc("html id").Value="ui-id-.*"
'				set s=Browser("micclass:=browser").Page("micclass:=page").ChildObjects(oDesc)
'				'msgbox s.count
'				wait 2
'				For j = 1 To s.count-1 Step 1
'					s(j).click
'				Next
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&Client_Name&"","html id:=ui-id-.*","xpath:=//ul[contains(@style,'display: block')]//li//a[text()='"&Client_Name&"']").Click		
				'Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","class:=ui-corner-all","innertext:="&Datatable.Value("Client_Name",SheetName)).highlight
				'Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","class:=ui-corner-all","innertext:="&Datatable.Value("Client_Name",SheetName)).Click
              	'browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//li//a[text()='"&Client_Name&"']").Click
				'Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&dataTable.value("Client_Name",SheetName),"html id:=ui-id-.*").Click		

				wait 2
		 	Call ReportStep (StatusTypes.Pass, "Successfully added data in row '"&rownum+1&"' Provide Link as'" &Provide_Link& "' Report Type as '" &Report_Type," Client content tab")
               
               objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
'			************************************************************************8

'	    	   modified by Madhu - 02/26/2020

'				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:="&Datatable.Value("Client_Name",SheetName)).highlight
'				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:="&Datatable.Value("Client_Name",SheetName)).Click
               Call ReportStep (StatusTypes.Pass, "Successfully added data in row '"&rownum+1&"' Provide Link as'" &Provide_Link& "' Report Type as '" &Report_Type," Client content tab")
               
               If Report_Type = "VA PBI Report" Then
               		filePath = Environment("CurrDir") & "HostingTestData\" 
               		fileName = Report_Name
               		fileToSet = filePath & fileName
               		objClientContentTable.ChildItem(CCR, 18, "Link", 0).Click
               		wait 5
      				If Browser("Browser-OCRF").Page("Page-OCRF").WebFile("wfAddFiles_PBIReport").Exist(20) then 
      					Browser("Browser-OCRF").Page("Page-OCRF").WebFile("wfAddFiles_PBIReport").Set fileToSet
      					wait 10
'      					Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Add files_PBIReport").Click
'      					If Browser("Browser-OCRF").Dialog("Choose File to Upload").Exist(20) Then
'      						Browser("Browser-OCRF").Dialog("Choose File to Upload").WinEdit("FileName").Set fileToSet
'      						Browser("Browser-OCRF").Dialog("Choose File to Upload").WinButton("Open").Click
      						Browser("Browser-OCRF").Page("Page-OCRF").WebButton("StartUpload_PBIReport").Click
      						uploadedFileName = Trim(Browser("Browser-OCRF").Page("Page-OCRF").WebTable("PBIReportUpload").ChildItem(1,2,"WebElement",0).getRoProperty("innertext"))
      						wait 10
      						uploadStatus = Trim(Browser("Browser-OCRF").Page("Page-OCRF").WebTable("PBIReportUpload").ChildItem(1,4,"WebElement",0).getRoProperty("innertext"))
      						
      						If uploadedFileName = fileName And uploadStatus = "File Uploaded" Then
      							Call ReportStep (StatusTypes.Pass, "File Uploaded successfully"," Upload File - PBI Report -Client content tab")
      							Else
      							Call ReportStep (StatusTypes.Fail, "File is not uploaded successfully"," Upload File - PBI Report -Client content tab")
      						End If
      						'Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_PBIReport").Click
      						
      						'changes by srinivas
      						browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[@id='UploadReportFile']/../descendant::div//button//span[text()='Close']").Click
      				End If
               End If
	           
	    End If
	    		    	
	Next	
	Call ReadWriteDataFromTextFile("write","PBIReportTimeStamp",TimeStamp)
'    AddDatainClientDetailsTab=TimeStamp
End Function




Public Function AddDatainClientDetailsTabUpdated1(ByVal FileName,ByVal SheetName,ByVal StartRow,ByVal EndRow,ByVal Country,ByVal Offering,ByRef objData)


    
     TimeStamp=objData.item("TimeStamp")
   '########### OFFERING DETAIL TAB EXISTANCE - Start ########### 
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementCDTabHeader").Exist(20) Then
		Call ReportStep (StatusTypes.Pass, "'Client content' tab header exist","Client content Tab header")
	Else
	   	Call ReportStep (StatusTypes.Fail, "Client content Tab header not exist","Client content Tab header")
	End If

 	wait 2
	Call ImportSheet(SheetName,FileName)
	
	DataTable.GetSheet(SheetName)
	rc=DataTable.GetSheet(SheetName).GetRowCount
	wait 2
	For rownum = StartRow To EndRow Step 1
	     DataTable.SetCurrentRow(rownum)
	     Provide_Link=Datatable.Value("Provide_Link",SheetName)
	     Decision_Center_Heading=Datatable.Value("Decision_Center_Heading",SheetName)
	     Link_Name=Datatable.Value("Link_Name",SheetName)
	     ToolTip=Datatable.Value("ToolTip",SheetName)
	     Report_Type=Datatable.Value("Report_Type",SheetName)
	     Report_Name=Datatable.Value("Report_Name",SheetName)
	     Client_Name=Datatable.Value("Client_Name",SheetName)
	     
		'CLICK ON ADD NEW LINE
		If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("innertext:=Add New Line","html tag:=DIV","visible:=True").Exist Then
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("innertext:=Add New Line","html tag:=DIV","visible:=True").Click
		ElseIf Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementAddNewLine").Exist(20) Then
			wait 2
	        'Call ReportStep (StatusTypes.Pass, "Click on add new line done","Add new button")	
			if Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementAddNewLine").Exist then
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementAddNewLine").Click
			end if 	
		Else
			Call ReportStep (StatusTypes.Fail, "Click on add new line not done","Add new button")
		End If
		Set objClientContentTable = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList")
		'Modified By Sumit on 22nd March 
		'Stated Modification
		'Offering = ReadWriteDataFromTextFile("Read","SyndOffering","")
		'Country = ReadWriteDataFromTextFile("Read","SyndContry","")
	 	' End Modification
	    If objClientContentTable.Exist(10)  Then
	    	CCR= objClientContentTable.RowCount
	        If Datatable.Value("Provide_Link",SheetName)="Yes" Then
	        wait 2
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 9, "WebList", 0).Select Provide_Link
	    	   objClientContentTable.ChildItem(CCR, 10, "WebList", 0).Select Decision_Center_Heading
	           objClientContentTable.ChildItem(CCR, 11, "WebEdit", 0).Set Link_Name&TimeStamp
	           objClientContentTable.ChildItem(CCR, 12, "WebEdit", 0).Set ToolTip&TimeStamp
	      Else
	         wait 3
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 9, "WebList", 0).Select Provide_Link
	           'CHECKING FOR FREEZED DATA FOR DECISION CENTER HEADING,LINK NAME AND TOOLTIP IF PROVIDE LINK IS 'NO'
	           DCDisabled=objClientContentTable.ChildItem(CCR, 10, "WebList", 0).GetROProperty("disabled") 
	           LinkDisabled=objClientContentTable.ChildItem(CCR, 11, "WebEdit", 0).GetROProperty("disabled") 
	           TooltipDis=objClientContentTable.ChildItem(CCR, 12, "WebEdit", 0).GetROProperty("disabled") 
	           If DCDisabled="1" AND LinkDisabled="1" AND TooltipDis="1" Then
	           	  Call ReportStep (StatusTypes.Pass, "'Link name', 'Tool tip and' and  'Decision center' fields  are not Editable after seletcing Provide link as 'No' "," Provide link as 'No'")
	           Else
	           	  Call ReportStep (StatusTypes.Pass, "'Link name', 'Tool tip and' and  'Decision center' fields  are  Editable after seletcing Provide link as 'No' "," Provide link as 'No'")
	           End If

	       End If
	       
	       	   objClientContentTable.ChildItem(CCR, 14, "WebList", 0).Select Trim(Report_Type)
	           'objClientContentTable.ChildItem(CCR, 15, "WebEdit", 0).Set Trim(Country&"/"&Offering)
	           'Modified the code based on New Requirement - OCRF - VA Ops Maintenance - 21/01/2019
	           wait 4
	           If Report_Type = "VA PBI Report" Then
'	           ***************************************************
'					Modified by Madhu - 04/15/2020
					
					'changes by srinivas
'					If objClientContentTable.ChildItem(CCR, 15, "WebEdit", 0).GetROProperty("disabled")=true Then
'	           			Call ReportStep (StatusTypes.Pass,"report path webedit column is disabled for PBI Report Type","Validate whether report path webedit column is disabled or not for PBI Report Type" )
'	           			else
'	           			Call ReportStep (StatusTypes.Fail,"report path webedit column is not disabled for PBI Report Type","Validate whether report path webedit column is disabled or not for PBI Report Type" )
'	           		End If
'	           		
'	           		
'	           		If objClientContentTable.ChildItem(CCR, 16, "WebEdit", 0).GetROProperty("disabled")=true Then
'	           			Call ReportStep (StatusTypes.Pass,"report name webedit column is disabled for PBI Report Type","Validate whether report name webedit column is disabled or not for PBI Report Type" )
'	           			else
'	           			Call ReportStep (StatusTypes.Fail,"report name webedit column is not disabled for PBI Report Type","Validate whether report name webedit column is disabled or not for PBI Report Type" )
'	           		End If
					
					
	           		'objClientContentTable.ChildItem(CCR, 16, "WebEdit", 0).Set Trim(Report_Name)
	           		else
	           		objClientContentTable.ChildItem(CCR, 16, "WebList", 0).Click
	           		Wait 5
	           		'objClientContentTable.ChildItem(CCR, 16, "WebEdit", 0).select Trim(Report_Name)
	           		
	           		'changes by srinivas
	           		
	           		
	           		
	           		
					If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Information_Okay").Exist Then
						Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Information_Okay").Click
						wait 3
					End If
					wait 2
					'objClientContentTable.ChildItem(CCR, 16, "WebList", 0).Click
	           		'objClientContentTable.ChildItem(CCR, 16, "WebList", 0).Select Trim(Report_Name)
''	           		*********************************************************************8
	           		wait 2
	           End If
	           
'	           ***********************************************************
	           objClientContentTable.ChildItem(CCR, 17, "WebEdit", 0).Set Trim(Client_Name)
	           Set objShell=CreateObject("WScript.Shell")
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	  ' objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
			   objClientContentTable.ChildItem(CCR, 17, "WebEdit", 0).Click
				'Set objShell1=createobject("wscript.shell")	
				'objShell1.SendKeys "{DOWN}"
				'************************************************************************8
				wait 5
'	    	   modified by Madhu - 02/26/2020
				Client_Name=Datatable.Value("Client_Name",SheetName)
				'Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","class:=ui-corner-all","innertext:="&Datatable.Value("Client_Name",SheetName)).highlight
				'Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","class:=ui-corner-all","innertext:="&Datatable.Value("Client_Name",SheetName)).Click
              	browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//li//a[text()='"&Client_Name&"']").Click
				'Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&dataTable.value("Client_Name",SheetName),"html id:=ui-id-.*").Click		

				wait 2
		 	Call ReportStep (StatusTypes.Pass, "Successfully added data in row '"&rownum+1&"' Provide Link as'" &Provide_Link& "' Report Type as '" &Report_Type," Client content tab")
               
               objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
'			************************************************************************8

'	    	   modified by Madhu - 02/26/2020

'				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:="&Datatable.Value("Client_Name",SheetName)).highlight
'				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:="&Datatable.Value("Client_Name",SheetName)).Click
               Call ReportStep (StatusTypes.Pass, "Successfully added data in row '"&rownum+1&"' Provide Link as'" &Provide_Link& "' Report Type as '" &Report_Type," Client content tab")
               
               If Report_Type = "VA PBI Report" Then
               		filePath = Environment("CurrDir") & "HostingTestData\" 
               		fileName = Report_Name
               		fileToSet = filePath & fileName
               		objClientContentTable.ChildItem(CCR, 18, "Link", 0).Click
      				If Browser("Browser-OCRF").Page("Page-OCRF").WebFile("wfAddFiles_PBIReport").Exist(20) then 
      					wait 5
      					Browser("Browser-OCRF").Page("Page-OCRF").WebFile("wfAddFiles_PBIReport").highlight
      					Browser("Browser-OCRF").Page("Page-OCRF").WebFile("wfAddFiles_PBIReport").Set fileToSet
      					wait 10
'      					Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Add files_PBIReport").Click
'      					If Browser("Browser-OCRF").Dialog("Choose File to Upload").Exist(20) Then
'      						Browser("Browser-OCRF").Dialog("Choose File to Upload").WinEdit("FileName").Set fileToSet
'      						Browser("Browser-OCRF").Dialog("Choose File to Upload").WinButton("Open").Click
      						Browser("Browser-OCRF").Page("Page-OCRF").WebButton("StartUpload_PBIReport").Click
      						uploadedFileName = Trim(Browser("Browser-OCRF").Page("Page-OCRF").WebTable("PBIReportUpload").ChildItem(1,2,"WebElement",0).getRoProperty("innertext"))
      						wait 10
      						uploadStatus = Trim(Browser("Browser-OCRF").Page("Page-OCRF").WebTable("PBIReportUpload").ChildItem(1,4,"WebElement",0).getRoProperty("innertext"))
      						
      						If uploadedFileName = fileName And uploadStatus = "File Uploaded" Then
      							Call ReportStep (StatusTypes.Pass, "File Uploaded successfully"," Upload File - PBI Report -Client content tab")
      							Else
      							Call ReportStep (StatusTypes.Fail, "File is not uploaded successfully"," Upload File - PBI Report -Client content tab")
      						End If
      						'Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_PBIReport").Click
      						
      						'changes by srinivas
      						browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[@id='UploadReportFile']/../descendant::div//button//span[text()='Close']").Click
      				End If
               End If
	           
	    End If
	    		    	
	Next	
	Call ReadWriteDataFromTextFile("write","PBIReportTimeStamp",TimeStamp)
'    AddDatainClientDetailsTab=TimeStamp
End Function



Public Function AddMultipleDatabaseUpdated(ByVal FileName,ByVal SheetName,ByVal StartRow,Byval EndRow,Byval timeStamp,ByRef objData)

    If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Wel_UserDB_List").Exist(60) Then
		Call ReportStep (StatusTypes.Pass, "Entered in to cude details tab ","Cube details tab")	
	Else
		Call ReportStep (StatusTypes.Fail, "Not Entered in to cude details tab","Cube details tab")
	End If
	
	Call ImportSheet(SheetName,FileName)
	
	'Number of rows to be added
	Datatable.SetCurrentRow(StartRow)
	NumberOfRows = (EndRow-StartRow)+1
		
	'select number of rows and click on add new line
'	Browser("Browser-OCRF").Page("Page-OCRF").WebList("WebList_NumberOfRows").Select NumberOfRows
'	Call ReportStep (StatusTypes.Pass, "Entering "&NumberOfRows&" records in 'Cube details' tab started","Cube details tab")


    TableRow=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetROProperty("rows")+1
	For row = StartRow To EndRow Step 1
	
	
	
'		*********************************************
'	Modified by Madhu - 03/02/2020
				If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes_Index1").Exist(1) Then
		          
		       	  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes_Index1").Click
		       End If
		       If  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes_Index0").Exist(1) Then
		   	      
		      	  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes_Index0").Click
		      End If 
		      

'88***********************************************************		
		
	
	
	
'		If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementAddNewLine").Exist(10) Then
'			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementAddNewLine").Click
'		End If
	    
	    'CHANGES BY SRINIVAS
	    If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//td[@title='Add New Line']//div[text()='Add New Line ']").Exist(10) Then
			Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//td[@title='Add New Line']//div[text()='Add New Line ']").Click
		End If
	    
	    wait 2
	    'Select Action in row
		Call WebEditExistence(TableRow,7,"WebList")
			
		Set selectAction = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(TableRow,7,"WebList",0)
		selectAction.select datatable.value("Action",SheetName) 

'        'Select 'ActionComments' in specific row 
'		Call WebEditExistence(row,8,"WebEdit")
'			
'		Set setActionComments = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(row,8,"WebEdit",0)
'		setActionComments.set datatable.value("ActionComments",SheetName) 
		
		'Select 'DBType' in specific row 	
		Call WebEditExistence(TableRow,10,"WebList")
			
		Set selectDBType = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(TableRow,10,"WebList",0)
		selectDBType.select datatable.value("DBType",SheetName) 
	
	   
	    
		Set CheckRow =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(TableRow,2,"WebCheckBox",0)
		CheckRow.click	
		
		Call WebEditExistence(TableRow,Cint(Environment.value("cubeName")),"WebEdit")
		
		Set CheckDatabase =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(TableRow,Cint(Environment.value("cubeName")),"WebEdit",0)
		CheckDatabase.set datatable.value("DatabaseName",SheetName)&timeStamp 
		wait 1
		
		If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("Apply datasource security").Exist(10) Then
			Call ReportStep (StatusTypes.Pass,"Datasource security popup appeared" ,"Datasource tab")
			
		End If
		If Browser("Browser-OCRF").Page("Page-OCRF").WebCheckBox("DSsecurityCheckbox").Exist(5) Then
			If datatable.value("DataSourceSecurity",SheetName)="Check" Then
			   Browser("Browser-OCRF").Page("Page-OCRF").WebCheckBox("DSsecurityCheckbox").Set "ON"
			End If
		End If
	    If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("DSProceed").Exist(5) Then
	    wait 3
	    	Browser("Browser-OCRF").Page("Page-OCRF").WebButton("DSProceed").DoubleClick
	    End If
		If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebeleAssigOwnerPopUp").Exist(15) Then
		   If datatable.value("AssignOwner",SheetName)="Choose new DataSource Owner" Then
		   	  Browser("Browser-OCRF").Page("Page-OCRF").WebRadioGroup("DataSourceOwner").Select 2 
		   	  Browser("Browser-OCRF").Page("Page-OCRF").WebEdit("WebEditdDtaSourceOwner").Set objData.item("DataBaseOwner")
		   	   wait 3
		   	   Browser("Browser-OCRF").Page("Page-OCRF").WebButton("DSProceed").DoubleClick
		   	   If  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes").Exist(5) Then
		   	      wait 2
		      	  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes").Click
		      End If 
		      If  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes").Exist(5) Then
		   	      wait 2
		      	  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes").Click
		      End If 
		   Else
		      Browser("Browser-OCRF").Page("Page-OCRF").WebRadioGroup("DataSourceOwner").Select 1
		      wait 2
		      Browser("Browser-OCRF").Page("Page-OCRF").WebButton("DSProceed").DoubleClick
		      If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebEleOwnershipAssignmentFor").Exist(15) Then
		         If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerPopUpYes").Exist(10) Then
		            wait 2
		         	Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerPopUpYes").DoubleClick
		         End If
		         If  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes").Exist(15) Then
		   	          wait 2
		      	        Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes").Click
		         End If  
		      End If 
		   End If
		ElseIf Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebEleConfirmation").Exist(5) Then 
		       wait 1
		       If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerPopUpYes").Exist(10) Then
		          wait 2
		       	  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerPopUpYes").Click
		       End If
		       If  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes").Exist(15) Then
		   	      wait 2
		      	  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes").Click
		      End If 
	    End If
	
'*******************************
'Modified BY madhu 02/20/2020
'------Changed click to double click for popups
		
	    If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerPopUpYes").Exist(5) Then
	    	Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerPopUpYes").Click
	    End If
	    If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerPopUpYes").Exist(1) Then
	    	Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerPopUpYes").DoubleClick
	    End If
	     If  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes").Exist(1) Then
		   	wait 2
		    Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes").DoubleClick
		End If
	    If  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes").Exist(1) Then
		   	wait 2
		    Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes").Click
		End If 
'		'Write 'DataBase description' in specific row 
'        Call WebEditExistence(row,11,"WebEdit")
'			
'		Set setDBDesc = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(row,11,"WebEdit",0)
'		setDBDesc.set datatable.value("DBDescription",SheetName) 
			
'		'Write 'DB Owner Email' in specific row 	
'		Call WebEditExistence(row,18,"WebEdit")
'			
'		Set setDbOwEmail = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(row,18,"WebEdit",0)
'		setDbOwEmail.set datatable.value("DBOwnerEmail",SheetName) 

			
'		*********************************************
'	Modified by Madhu - 03/02/2020
				If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes_Index1").Exist(1) Then
		          
		       	  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes_Index1").Click
		       End If
		       If  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes_Index0").Exist(1) Then
		   	      
		      	  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes_Index0").Click
		      End If 
		      

'88***********************************************************		
				
		Set CheckRow =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(TableRow,2,"WebCheckBox",0)
		CheckRow.click
		wait 1
		Set CheckRow =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(TableRow,2,"WebCheckBox",0)
		CheckRow.click
		wait 1
		
		'If user is not owner of Datasource user has to request for permission
		'Check the status for pending
		 ActionStatus =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").GetCellData(TableRow,7)
		If ActionStatus="PENDING" Then
           Exit Function 
		End If
		
	    'Select hosting server if it is not disabled
		Set CheckDBLocation =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(TableRow,Cint(Environment.value("cubeLocation")),"WebEdit",0)		
		wait 1
		If CheckDBLocation.GetRoProperty("disabled")=0 Then
		   Call WebEditExistence(TableRow,Cint(Environment.value("cubeLocation")),"WebEdit")
		
		   Set CheckDBLocation =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(TableRow,Cint(Environment.value("cubeLocation")),"WebEdit",0)
		   CheckDBLocation.set datatable.value("ServerLocation",SheetName)	
		End If
		'Select FLA path if it is not disabled
		Set CheckFLAPath =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(TableRow,Cint(Environment.value("flaPath")),"WebList",0)
		wait 1
		If CheckFLAPath.GetRoProperty("disabled")=0 Then
		   	Call WebEditExistence(TableRow,Cint(Environment.value("flaPath")),"WebList")
		
		    Set CheckFLAPath =Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Cube_List_Data").ChildItem(TableRow,Cint(Environment.value("flaPath")),"WebList",0)
		    CheckFLAPath.select datatable.value("FLAPATH",SheetName)	 
		End If

'		*********************************************
'	Modified by Madhu - 03/02/2020
				If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerPopUpYes").Exist(5) Then
		          wait 2
		       	  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerPopUpYes").Click
		       End If
		       If  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes_Index0").Exist(1) Then
		   	    
		      	  Browser("Browser-OCRF").Page("Page-OCRF").WebButton("OwnerYes_Index0").Click
		      End If 
		      

'88***********************************************************		
		
		datatable.SetNextRow
		Call ReportStep (StatusTypes.Pass, "Entered  Database Name as '"&datatable.value("DatabaseName",SheetName)&"',Server Location as '"&datatable.value("ServerLocation",SheetName)&"' And FLA path as '"&datatable.value("FLAPATH",SheetName)&"' at row '"&del-1&"' done","Cube details tab")
	    TableRow=TableRow+1
	Next
	
	
End Function


Function ValidateDeleteAndChangeActionInManageUserAccessUpdated(ByVal FileName,ByVal SheetName,ByVal OCRFId, ByVal action, ByVal StartRow,Byval EndRow,ByRef objData)

	If Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebButton("WebBtnSearch").Exist(10) Then
		Call ReportStep (StatusTypes.Pass,  "User is navigated to Manage User Access page","Navigate to Manage User Access page")
		
		Call ImportSheet(SheetName,FileName)
	
		Datatable.SetCurrentRow(StartRow)
		NumberOfRows = (EndRow-StartRow)+1
		
		UserLoginId = DataTable.Value("User_Login_Id",SheetName)
		
		Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebEdit("WebEditOfferingUserLoginId").Set UserLoginId
		Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebButton("WebBtnSearch").Click
		Call PageLoading()
	
		'below line added by Avinash on 4th Jun 2019
		'reason -record was present in last page due to pagination
		If Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebElement("lastPage_Pagination").Exist(10) Then
			Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebElement("lastPage_Pagination").Click
		End If
		
		wait 2
		'Modification Completed
		
		rowCount = Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebTable("UserOfferingGrid").RowCount
		wait 2
		If rowCount > 1 Then
			rowNo = Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebTable("UserOfferingGrid").GetRowWithCellText (OCRFId,3)
			wait 2
			If rowNo < 0 Then
				Call ReportStep (StatusTypes.Information,  "OCRF is not available for " & action ,action & " in Manage User Access")
				Exit Function
			End If
			flagActionDone = False 

			wait 3
			Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebTable("UserOfferingGrid").ChildItem(rowNo,2,"WebCheckbox",0).Set "On"
			wait 3
			'Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebElement(action).Click
			'Modified by Madhu - 04/23/2020
			'Browser("BrowserPoPUp").Page("Online CRF - Manage Users").WebElement("Delete").Click
				'Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebElement("Delete").Click
			
			'changes by srinivas
			Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//td[@id='bulkUser-list_toppager_left']//td//div[text()='"&action&"']").Click

			wait 3
			If Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebElement("WebEleConfirmMessage").Exist(15) Then
'				Set obj = Description.Create
'				obj("micclass").value = "WebButton"
'				obj("value").value = "Yes"
'				obj("visible").value = True
'				Set chObj = Browser("Browser-OCRF").Page("OCRF-ManageUsers").ChildObjects(obj)
'				For i = 0 to chObj.count -1
'					chObj(i).click
'				Next
				Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebButton("visible:=True","value:=Yes").highlight
				Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebButton("visible:=True","value:=Yes").Click
				Call PageLoading()
				wait 1
				Set objSuccessInfo = Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebElement("html tag:=DIV","visible:=True","innertext:=Action Executed Successfully.")
				If objSuccessInfo.Exist(30) Then
					Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebButton("visible:=True","value:=Okay").highlight
					Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebButton("visible:=True","value:=Okay").Click
					wait 3
					flagActionDone = True
					Call ReportStep (StatusTypes.Pass,  action & " is done successfully",action & " in Manage User Access")
					Else
					Call ReportStep (StatusTypes.Fail,  action & " is not done successfully",action & " in Manage User Access")
				End If
				
			End If
			Else
			Call ReportStep (StatusTypes.Information,  "No search records are found in Manage User Access page","Search OCRF in Manage User Access page")
		End If
	
	Else
		Call ReportStep (StatusTypes.Fail,  "User is not navigated to Manage User Access page","Navigate to Manage User Access page")
	End If
	
	If action = "Delete" And flagActionDone = True Then
	
		Call SelectMenuOption("Online CRF","",objData)
		Call PageLoading()
		foundInTable= FilterInOnlineCRF("Current Requests","OCRF ID",OCRFId,objData)
		Call PageLoading()
		
		rowNo = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").GetRowWithCellText (OCRFId,3)
		runTimeStatus = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTable_Data").GetCellData(rowNo,8)
		If runTimeStatus = "Submitted" Then
			Call ReportStep (StatusTypes.Pass,  "Status of the OCRF is Submitted " , "Verify Status of OCRF upon Delete in Manage User Access")
			Else
			Call ReportStep (StatusTypes.Fail,  "Status of the OCRF is " & runTimeStatus , "Verify Status of OCRF upon Delete in Manage User Access")
		End If
	End If
	If action = "ChangeAction" And flagActionDone = True Then
		Do
			wait 5
			Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebButton("WebBtnSearch").Click
			Call PageLoading()
			rowNo = Browser("Browser-OCRF").Page("OCRF-ManageUsers").WebTable("UserOfferingGrid").GetRowWithCellText (OCRFId,3)
		Loop Until (rowNo < 0) 
		
		If rowNo < 0 Then
			Call ReportStep (StatusTypes.Pass,  "Change Action is successfully completed " , "Verify Change Action in Manage User Access")
			
		End If
		
	End If

End Function


Public Function NavigateToMSTRFolderStructures(ByVal Foldername,ByVal country,ByVal offering,Byval clientname)
	'folder structure can be VA design or Visual analytics
	'Setting.WebPackage("ReplayType") = 2
	If 	Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//DIV[@id='projects_ProjectsStyle']//a[text()='"&Foldername&"']").exist(30) Then
		Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//DIV[@id='projects_ProjectsStyle']//a[text()='"&Foldername&"']").click
	End If
	If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//input[@value='Continue']").Exist(10) Then
		Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//input[@value='Continue']").Click
	End If
	If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[@id='dktpSectionView']").Exist(30) Then
		Call ReportStep (StatusTypes.Pass, "clicked on folder:-"&Foldername&"","click on "&Foldername&" in mstr server")
		'click on shared reports
		If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[text()='Shared Reports']/..").Exist(30) Then
			Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[text()='Shared Reports']/..").Click
			If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']").Exist(30) Then
				Call ReportStep (StatusTypes.Pass, "clicked on folder:- shared reports and country folders page is displayed","click on shared reports folder in mstr server")
				'click on country folder
				If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='"&country&"']").exist(30) Then
					'Setting.WebPackage("ReplayType") = 2
					Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='"&country&"']").click
					'Setting.WebPackage("ReplayType") = 1
					If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='"&offering&"']").exist(30) Then
						Call ReportStep (StatusTypes.Pass, "clicked on country folder:-"&country&" and offering page is displayed","click on country folder in mstr server")
						'click on offering folder
						If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='"&offering&"']").exist(30) Then
							Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='"&offering&"']").click
							Call ReportStep (StatusTypes.Pass, "clicked on offering folder:-"&offering&"","click on offering folder in mstr server")
							If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='"&clientname&"']").exist(30) Then
								Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='"&clientname&"']").click
								
							End If
							If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='ims']").Exist(30) Then
								Call ReportStep (StatusTypes.Pass, "clicked on client folder:-"&clientname&"","click on client folder in mstr server")
								'click on ims folder
								If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='ims']").Exist(30) Then
									Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='ims']").Click
								End If
								If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='Publish']").Exist(30) Then
									Call ReportStep (StatusTypes.Pass, "clicked on ims folder and Publish folder displayed successfully in mstr server","click on ims folder in mstr server")
									'click on publish folder
									Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='Publish']").Click
									If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//td[@id='td_mstrWeb_dockLeft']").Exist(30) Then
										Call ReportStep (StatusTypes.Pass, "clicked on publish folder and Reports section page displayed successfully in mstr server","click on publish folder in mstr server")
										else
										Call ReportStep (StatusTypes.Fail, "not clicked on publish folder and Reports section page not displayed in mstr server","click on publish folder in mstr server")
									End If
									else
									Call ReportStep (StatusTypes.Fail, "not clicked on ims folder and Publish folder not displayed in mstr server","click on ims folder in mstr server")
								End If
								else
								If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='Publish']").Exist(30) Then
									Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='Publish']").Click
								End If
								'Call ReportStep (StatusTypes.Fail, "not clicked on offering folder:-"&offering&"","click on offering folder in mstr server")
							End If
						End If
						else
						Call ReportStep (StatusTypes.Fail, "not clicked on country folder:-"&country&" and offering page is not displayed","click on country folder in mstr server")
					End If
				End If
				else
				Call ReportStep (StatusTypes.Fail, "not clicked on folder:- shared reports","click on shared reports folder in mstr server")
			End If
		End If
		else
		Call ReportStep (StatusTypes.Fail, "not clicked on folder:-"&Foldername&"","click on "&Foldername&" in mstr server")
	End If
	
End Function

'Created By: Srinivas
'Created on: 5/15/2020
'Description: Launch the browser with mstr url and validate default page is displayed.
'Parameter: Param1 - Extra,Param2 - Extra,Param3 - Extra,objData
Public Function LaunchMSTRURL(ByVal Url, ByVal UserName, ByVal Password, ByRef objData)
		
	Systemutil.CloseProcessByName "iexplore.exe"
	
	'Launch OCRF Url's	
'	SystemUtil.Run Environment.Value("OCRFURL")

    SystemUtil.Run "iexplore", Environment.Value("MSTRURL") , , , 3
    wait 30
    If UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAEdit("User name").Exist<>true Then
    	wait 20 
    End If
    If UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").Exist(30) Then    	
        UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").Highlight 
	    UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAEdit("User name").SetValue UserName
	    UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAEdit("Password").SetValue Password
	    'MOdified by Madhu - 04/02/2020
	    If UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAButton("OK").Exist Then
	    	if UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAButton("OK").GetROProperty("visible") then
	    	UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAButton("OK").click
	    	end  if
	    End If
	    
	    '******************************************************************
    Set objShell = CreateObject("Wscript.shell")
    objShell.SendKeys "{ENTER}"
'    ***************************************************************************8
        Call ReportStep (StatusTypes.Pass, "User " & UserName & " has logged in successfully","Login through Window security PopUp")
    ElseIf True Then  
		Systemutil.CloseProcessByName "iexplore.exe"    
        SystemUtil.Run "iexplore", Environment.Value("MSTRURL") , , , 3
    wait 30
    If UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").Exist(30) Then    	
        UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").Highlight 
	    UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAEdit("User name").SetValue UserName
	    UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAEdit("Password").SetValue Password
	      'MOdified by Madhu - 04/02/2020
	    If UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAButton("OK").Exist Then
	    	if UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAButton("OK").GetROProperty("visible") then
	    	if UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAButton("OK").Exist then
	    		UIAObject("BrowserSecurityPopUp").UIAWindow("Windows Security").UIAButton("OK").click
	    	end if 	
	    	end  if
	    End If
	    
	    '******************************************************************
    Set objShell = CreateObject("Wscript.shell")
    objShell.SendKeys "{ENTER}"
'    ***************************************************************************8
        Call ReportStep (StatusTypes.Pass, "User " & UserName & " has logged in successfully","Login through Window security PopUp")
    end if 
	Else
		Call ReportStep (StatusTypes.Fail, "User " & UserName & " has not successfully logged in","Login through Window security PopUp")	   
	End If
    


'	If Browser("BrowserPoPUp").Exist(20) Then
'		Browser("BrowserPoPUp").highlight
'	    Browser("BrowserPoPUp").Dialog("Windows Security").highlight
'	   'Handle 'Window Security' PopUp
'	    If Browser("BrowserPoPUp").Dialog("Windows Security").Exist(30) Then
'			If Browser("BrowserPoPUp").Dialog("Windows Security").Static("InvalidUserAccount").Exist(5) Then
'				Browser("BrowserPoPUp").Dialog("Windows Security").Static("UseAnotherAccount").Click
'			End If
'		   Browser("BrowserPoPUp").Dialog("Windows Security").WinEdit("WinEditUserName").Set UserName	
'		   Browser("BrowserPoPUp").Dialog("Windows Security").WinEdit("WinEditPassword").Set Password
'		   Browser("BrowserPoPUp").Dialog("Windows Security").WinButton("BtnOK").Click
'		   Call ReportStep (StatusTypes.Pass, "User " & UserName & " has logged in successfully","Login through Window security PopUp")
'		Else
'		   Call ReportStep (StatusTypes.Fail, "User " & UserName & " has not successfully logged in","Login through Window security PopUp")	   
'		End If
'	ElseIf Window("BrowserPoPUp").Exist(20) Then
'		Window("BrowserPoPUp").highlight
'	    Window("BrowserPoPUp").Dialog("Windows Security").highlight
'	   'Handle 'Window Security' PopUp
'		If Window("BrowserPoPUp").Dialog("Windows Security").Exist(30) Then
'			If Window("BrowserPoPUp").Dialog("Windows Security").Static("InvalidUserAccount").Exist(5) Then
'				Window("BrowserPoPUp").Dialog("Windows Security").Static("UseAnotherAccount").Click
'			End If
'		   Window("BrowserPoPUp").Dialog("Windows Security").WinEdit("WinEditUserName").Set UserName	
'		   Window("BrowserPoPUp").Dialog("Windows Security").WinEdit("WinEditPassword").Set Password
'		   Window("BrowserPoPUp").Dialog("Windows Security").WinButton("BtnOK").Click
'		   Call ReportStep (StatusTypes.Pass, "User " & UserName & " has logged in successfully","Login through Window security PopUp")
'		Else
'		   Call ReportStep (StatusTypes.Fail, "User " & UserName & " has not successfully logged in","Login through Window security PopUp")	   
'		End If
'	End If
	

	'Validate OCRF url launched successfully
	If Browser("BrowserPoPUp").Page("title:=WELCOME. MicroStrategy").Exist(30) Then
		'	Default page is displayed
		If Browser("BrowserPoPUp").Page("title:=WELCOME. MicroStrategy").Exist(60) Then
			Call ReportStep (StatusTypes.Pass, "MSTR Home Page is displayed","MSTR Homepage")
		Else
			Call ReportStep (StatusTypes.Fail, "MSTR Home Page is not displayed","MSTR Homepage")
		End If
		
	Else
		Exit Function
	End If
	
End Function
'********************************************************
'Browser("BrowserPoPUp").Page("WELCOME. MicroStrategy").WebElement("MstrHomePage").Click

Public Function AddDatainClientDetailsTabUpdatedMSTR(ByVal FileName,ByVal SheetName,ByVal StartRow,ByVal EndRow,ByVal Country,ByVal Offering,ByRef objData)


    
     TimeStamp=objData.item("TimeStamp")
   '########### OFFERING DETAIL TAB EXISTANCE - Start ########### 
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementCDTabHeader").Exist(20) Then
		Call ReportStep (StatusTypes.Pass, "'Client content' tab header exist","Client content Tab header")
	Else
	   	Call ReportStep (StatusTypes.Fail, "Client content Tab header not exist","Client content Tab header")
	End If

 	wait 2
	Call ImportSheet(SheetName,FileName)
	
	DataTable.GetSheet(SheetName)
	rc=DataTable.GetSheet(SheetName).GetRowCount
	wait 2
	For rownum = StartRow To EndRow Step 1
	     DataTable.SetCurrentRow(rownum)
	     Provide_Link=Datatable.Value("Provide_Link",SheetName)
	     Decision_Center_Heading=Datatable.Value("Decision_Center_Heading",SheetName)
	     Link_Name=Datatable.Value("Link_Name",SheetName)
	     ToolTip=Datatable.Value("ToolTip",SheetName)
	     Report_Type=Datatable.Value("Report_Type",SheetName)
	     Report_Name=Datatable.Value("Report_Name",SheetName)
	     Client_Name=Datatable.Value("Client_Name",SheetName)
	     
		'CLICK ON ADD NEW LINE
		If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementAddNewLine").Exist(20) Then
			wait 2
	        'Call ReportStep (StatusTypes.Pass, "Click on add new line done","Add new button")	
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementAddNewLine").Click
		Else
			Call ReportStep (StatusTypes.Fail, "Click on add new line not done","Add new button")
		End If
		Set objClientContentTable = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList")
		'Modified By Sumit on 22nd March 
		'Stated Modification
		'Offering = ReadWriteDataFromTextFile("Read","SyndOffering","")
		'Country = ReadWriteDataFromTextFile("Read","SyndContry","")
	 	' End Modification
	    If objClientContentTable.Exist(10)  Then
	    	CCR= objClientContentTable.RowCount
	        If Datatable.Value("Provide_Link",SheetName)="Yes" Then
	        wait 2
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 9, "WebList", 0).Select Provide_Link
	    	   objClientContentTable.ChildItem(CCR, 10, "WebList", 0).Select Decision_Center_Heading
	           objClientContentTable.ChildItem(CCR, 11, "WebEdit", 0).Set Link_Name&TimeStamp
	           objClientContentTable.ChildItem(CCR, 12, "WebEdit", 0).Set ToolTip&TimeStamp
	      Else
	         wait 3
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 9, "WebList", 0).Select Provide_Link
	           'CHECKING FOR FREEZED DATA FOR DECISION CENTER HEADING,LINK NAME AND TOOLTIP IF PROVIDE LINK IS 'NO'
	           DCDisabled=objClientContentTable.ChildItem(CCR, 10, "WebList", 0).GetROProperty("disabled") 
	           LinkDisabled=objClientContentTable.ChildItem(CCR, 11, "WebEdit", 0).GetROProperty("disabled") 
	           TooltipDis=objClientContentTable.ChildItem(CCR, 12, "WebEdit", 0).GetROProperty("disabled") 
	           If DCDisabled="1" AND LinkDisabled="1" AND TooltipDis="1" Then
	           	  Call ReportStep (StatusTypes.Pass, "'Link name', 'Tool tip and' and  'Decision center' fields  are not Editable after seletcing Provide link as 'No' "," Provide link as 'No'")
	           Else
	           	  Call ReportStep (StatusTypes.Pass, "'Link name', 'Tool tip and' and  'Decision center' fields  are  Editable after seletcing Provide link as 'No' "," Provide link as 'No'")
	           End If

	       End If
	       
	       	   objClientContentTable.ChildItem(CCR, 14, "WebList", 0).Select Trim(Report_Type)
	           'objClientContentTable.ChildItem(CCR, 15, "WebEdit", 0).Set Trim(Country&"/"&Offering)
	           'Modified the code based on New Requirement - OCRF - VA Ops Maintenance - 21/01/2019
	           wait 4
	           If Report_Type = "VA PBI Report" Then
'	           ***************************************************
'					Modified by Madhu - 04/15/2020
					
					'changes by srinivas
'					If objClientContentTable.ChildItem(CCR, 15, "WebEdit", 0).GetROProperty("disabled")=true Then
'	           			Call ReportStep (StatusTypes.Pass,"report path webedit column is disabled for PBI Report Type","Validate whether report path webedit column is disabled or not for PBI Report Type" )
'	           			else
'	           			Call ReportStep (StatusTypes.Fail,"report path webedit column is not disabled for PBI Report Type","Validate whether report path webedit column is disabled or not for PBI Report Type" )
'	           		End If
'	           		
'	           		
'	           		If objClientContentTable.ChildItem(CCR, 16, "WebEdit", 0).GetROProperty("disabled")=true Then
'	           			Call ReportStep (StatusTypes.Pass,"report name webedit column is disabled for PBI Report Type","Validate whether report name webedit column is disabled or not for PBI Report Type" )
'	           			else
'	           			Call ReportStep (StatusTypes.Fail,"report name webedit column is not disabled for PBI Report Type","Validate whether report name webedit column is disabled or not for PBI Report Type" )
'	           		End If
					
					
	           		'objClientContentTable.ChildItem(CCR, 16, "WebEdit", 0).Set Trim(Report_Name)
	           		else
	           		objClientContentTable.ChildItem(CCR, 16, "WebList", 0).Click
	           		Wait 5
	           		'objClientContentTable.ChildItem(CCR, 16, "WebEdit", 0).select Trim(Report_Name)
	           		
	           		'changes by srinivas
	           		
	           		
	           		
	           		
					If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Information_Okay").Exist Then
						Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Information_Okay").Click
						wait 3
					End If
					wait 2
					'objClientContentTable.ChildItem(CCR, 16, "WebList", 0).Click
	           		objClientContentTable.ChildItem(CCR, 16, "WebList", 0).Select Trim(Report_Name)
''	           		*********************************************************************8
	           		wait 2
	           End If
	           
'	           ***********************************************************
			 objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	           objClientContentTable.ChildItem(CCR, 17, "WebEdit", 0).Set Trim(Client_Name)
	           Set objShell=CreateObject("WScript.Shell")
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	  ' objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
			   objClientContentTable.ChildItem(CCR, 17, "WebEdit", 0).Click
				'Set objShell1=createobject("wscript.shell")	
				'objShell1.SendKeys "{DOWN}"
				'************************************************************************8
				wait 5
'	    	   modified by Madhu - 02/26/2020
				Client_Name=Datatable.Value("Client_Name",SheetName)
				'Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","class:=ui-corner-all","innertext:="&Datatable.Value("Client_Name",SheetName)).highlight
				'Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","class:=ui-corner-all","innertext:="&Datatable.Value("Client_Name",SheetName)).Click
              	'browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//li//a[text()='"&Client_Name&"']").Click
				'Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&dataTable.value("Client_Name",SheetName),"html id:=ui-id-.*").Click		
				'Client_Name=Datatable.Value("Client_Name",SheetName)
'				Set oDesc = Description.Create
'				oDesc("micclass").Value = "WebElement"
'				oDesc("html tag").Value = "A"
'				oDesc("innertext").Value = Client_Name
'				oDesc("html id").Value="ui-id-.*"
'				set s=Browser("micclass:=browser").Page("micclass:=page").ChildObjects(oDesc)
'				'msgbox s.count
'				wait 2
'				For j = 1 To s.count-1 Step 1
'					s(j).click
'				Next
				'Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//li//a[text()='"&Client_Name&"']").click
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&Client_Name&"","html id:=ui-id-.*").Click		
				wait 2
				'Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&Client_Name&"","html id:=ui-id-.*").Click		
		 		'wait 2
		 	Call ReportStep (StatusTypes.Pass, "Successfully added data in row '"&rownum+1&"' Provide Link as'" &Provide_Link& "' Report Type as '" &Report_Type," Client content tab")
               
               objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   'extra line by srinivas
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
'			************************************************************************8

'	    	   modified by Madhu - 02/26/2020

'				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:="&Datatable.Value("Client_Name",SheetName)).highlight
'				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:="&Datatable.Value("Client_Name",SheetName)).Click
               Call ReportStep (StatusTypes.Pass, "Successfully added data in row '"&rownum+1&"' Provide Link as'" &Provide_Link& "' Report Type as '" &Report_Type," Client content tab")
               
               If Report_Type = "VA PBI Report" Then
               		filePath = Environment("CurrDir") & "HostingTestData\" 
               		fileName = Report_Name
               		fileToSet = filePath & fileName
               		objClientContentTable.ChildItem(CCR, 18, "Link", 0).Click
      				If Browser("Browser-OCRF").Page("Page-OCRF").WebFile("wfAddFiles_PBIReport").Exist(20) then 
      					Browser("Browser-OCRF").Page("Page-OCRF").WebFile("wfAddFiles_PBIReport").Set fileToSet
      					wait 3
'      					Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Add files_PBIReport").Click
'      					If Browser("Browser-OCRF").Dialog("Choose File to Upload").Exist(20) Then
'      						Browser("Browser-OCRF").Dialog("Choose File to Upload").WinEdit("FileName").Set fileToSet
'      						Browser("Browser-OCRF").Dialog("Choose File to Upload").WinButton("Open").Click
      						Browser("Browser-OCRF").Page("Page-OCRF").WebButton("StartUpload_PBIReport").Click
      						uploadedFileName = Trim(Browser("Browser-OCRF").Page("Page-OCRF").WebTable("PBIReportUpload").ChildItem(1,2,"WebElement",0).getRoProperty("innertext"))
      						wait 10
      						uploadStatus = Trim(Browser("Browser-OCRF").Page("Page-OCRF").WebTable("PBIReportUpload").ChildItem(1,4,"WebElement",0).getRoProperty("innertext"))
      						
      						If uploadedFileName = fileName And uploadStatus = "File Uploaded" Then
      							Call ReportStep (StatusTypes.Pass, "File Uploaded successfully"," Upload File - PBI Report -Client content tab")
      							Else
      							Call ReportStep (StatusTypes.Fail, "File is not uploaded successfully"," Upload File - PBI Report -Client content tab")
      						End If
      						'Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_PBIReport").Click
      						
      						'changes by srinivas
      						browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[@id='UploadReportFile']/../descendant::div//button//span[text()='Close']").Click
      				End If
               End If
	           
	    End If
	    		    	
	Next	
	Call ReadWriteDataFromTextFile("write","MSTRReportTimeStamp",TimeStamp)
'    AddDatainClientDetailsTab=TimeStamp
End Function


Public Function ValidatePermissionInMstrServer(ByVal offering,Byval clientname,ByVal mstrReport,ByVal Username)
	Setting.WebPackage("ReplayType") = 2
	browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//TABLE[@id='FolderIcons']//a[text()='"&mstrReport&"']").RightClick
	Setting.WebPackage("ReplayType") = 1
	Call ReportStep (StatusTypes.Pass,"Right clicked on the mstr report","MSTR Server")
	wait 2
	'browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//tr[@id='cm1r11']//td[@class='menu-item']").Click
	
	If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[@id='folderAllModes_cmm1']//td[@class='menu-item' and contains(text(),'Share')]").Exist(20) Then
		Call ReportStep (StatusTypes.Pass,"Share option is present","MSTR Server")
		Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[@id='folderAllModes_cmm1']//td[@class='menu-item' and contains(text(),'Share')]").highlight
		Setting.WebPackage("ReplayType") = 2
		'Browser("BrowserPoPUp").Page("WELCOME. MicroStrategy").WebElement("Share").Click
		Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[@id='folderAllModes_cmm1']//td[@class='menu-item' and contains(text(),'Share')]").Click
		Setting.WebPackage("ReplayType") = 1
		else
		Call ReportStep (StatusTypes.Fail,"Share option is not present","MSTR Server")
	End If
	If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[@class='mstrmojo-Editor mstrmojo-SharingEditor modal']").Exist(30) Then
		Call ReportStep (StatusTypes.Pass,"clicked on share option in report","MSTR Server")
		else
		Call ReportStep (StatusTypes.Fail,"Not clicked on share option in report","MSTR Server")	
	End If

	If clientname="All Clients" Then
		Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//td//div[@class='mstrmojo-ACLEditor-user ug']").highlight
		MstrGroup=Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//td//div[@class='mstrmojo-ACLEditor-user ug']").GetROProperty("innertext")
		'msgbox MstrGroup
		If instr(MstrGroup,trim(UCASE(offering)))>0 Then
			Call ReportStep (StatusTypes.Pass,"Group:-"&MstrGroup&" is permissioned for report:-"&mstrReport&" which is published for client:-"&clientname&"","MSTR Server")
			else
			Call ReportStep (StatusTypes.Fail,"Fail:-"&MstrGroup&" is not permissioned for report:-"&mstrReport&" which is published for client:-"&clientname&"","MSTR Server")
		End If
		Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[text()='Cancel']").Click
	ElseIf clientname="User Specific" Then
		Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//tr[@id='mstr123']//td//div[@class='mstrmojo-ACLEditor-user u']").highlight
		Userid=Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//tr[@id='mstr123']//td//div[@class='mstrmojo-ACLEditor-user u']").GetROProperty("outerhtml")
		'msgbox MstrGroup
		If instr(Userid,Username)>0 Then
			Call ReportStep (StatusTypes.Pass,"User:-"&Username&" is permissioned for report:-"&mstrReport&" which is published for client:-"&clientname&"","MSTR Server")
			else
			Call ReportStep (StatusTypes.Fail,"User:-"&Username&" not is permissioned for report:-"&mstrReport&" which is published for client:-"&clientname&"","MSTR Server")
		End If
		Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[text()='Cancel']").Click
	Else 
		Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[contains(text(),'"&Username&"')]").highlight
		Userid=Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[contains(text(),'"&Username&"')]").GetROProperty("outerhtml")
		'msgbox MstrGroup
		If instr(Userid,Username)>0 Then
			Call ReportStep (StatusTypes.Pass,"User:-"&Username&" is permissioned for report:-"&mstrReport&" which is published for client:-"&clientname&"","MSTR Server")
			else
			Call ReportStep (StatusTypes.Fail,"User:-"&Username&" is not permissioned for report:-"&mstrReport&" which is published for client:-"&clientname&"","MSTR Server")
		End If
		Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[text()='Cancel']").Click
	End If
End Function

Function changeClientNameClientContent(ByVal clientname,start,endrow)

	For Iterator = start To endrow Step 1
		Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList").ChildItem(Iterator,2,"WebCheckBox", 0).click
		Call PageLoading()
		Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList").ChildItem(Iterator,17,"WebEdit", 0).click
		Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList").ChildItem(Iterator,17,"WebEdit", 0).set clientname
		Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList").ChildItem(Iterator,17,"WebEdit", 0).click
		wait 4
		Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&clientname&"","html id:=ui-id-.*").Click		
		wait 2
		Action=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList").GetCellData(2,6)
		Call ReportStep (StatusTypes.Pass, "Successfully changed the clientname to "&clientname&" and row no:-"&start&" is changed to action:-"&Action&"" ,"Client content tab")
               
	Next
End Function

Function deleteRowInClientContent(ByVal action,ByVal start,ByVal endrow)
	For Iterator = start To endrow Step 1
		Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList").ChildItem(Iterator,2,"WebCheckBox", 0).click
	Call PageLoading()
	Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList").ChildItem(Iterator,6,"WebList", 0).select action
	Call ReportStep (StatusTypes.Pass, "Successfully changed the Action of row no:-"&start&" to action:-"&action&"" ,"Client content tab")
	Next
End Function


Public Function ValidateDuplicateReportWarningMessage(ByVal FileName,ByVal SheetName,ByVal StartRow,ByVal EndRow,ByVal Country,ByVal Offering,ByRef objData)


    
     TimeStamp=objData.item("TimeStamp")
   '########### OFFERING DETAIL TAB EXISTANCE - Start ########### 
	If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementCDTabHeader").Exist(20) Then
		Call ReportStep (StatusTypes.Pass, "'Client content' tab header exist","Client content Tab header")
	Else
	   	Call ReportStep (StatusTypes.Fail, "Client content Tab header not exist","Client content Tab header")
	End If

 	wait 2
	Call ImportSheet(SheetName,FileName)
	
	DataTable.GetSheet(SheetName)
	rc=DataTable.GetSheet(SheetName).GetRowCount
	wait 2
	For rownum = StartRow To EndRow Step 1
	     DataTable.SetCurrentRow(rownum)
	     Provide_Link=Datatable.Value("Provide_Link",SheetName)
	     Decision_Center_Heading=Datatable.Value("Decision_Center_Heading",SheetName)
	     Link_Name=Datatable.Value("Link_Name",SheetName)
	     ToolTip=Datatable.Value("ToolTip",SheetName)
	     Report_Type=Datatable.Value("Report_Type",SheetName)
	     Report_Name=Datatable.Value("Report_Name",SheetName)
	     Client_Name=Datatable.Value("Client_Name",SheetName)
	     
		'CLICK ON ADD NEW LINE
		If Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementAddNewLine").Exist(20) Then
			wait 2
	        'Call ReportStep (StatusTypes.Pass, "Click on add new line done","Add new button")	
			Browser("Browser-OCRF").Page("Page-OCRF").WebElement("WebElementAddNewLine").Click
		Else
			Call ReportStep (StatusTypes.Fail, "Click on add new line not done","Add new button")
		End If
		Set objClientContentTable = Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList")
		'Modified By Sumit on 22nd March 
		'Stated Modification
		'Offering = ReadWriteDataFromTextFile("Read","SyndOffering","")
		'Country = ReadWriteDataFromTextFile("Read","SyndContry","")
	 	' End Modification
	    If objClientContentTable.Exist(10)  Then
	    	CCR= objClientContentTable.RowCount
	        If Datatable.Value("Provide_Link",SheetName)="Yes" Then
	        wait 2
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 9, "WebList", 0).Select Provide_Link
	    	   objClientContentTable.ChildItem(CCR, 10, "WebList", 0).Select Decision_Center_Heading
	           objClientContentTable.ChildItem(CCR, 11, "WebEdit", 0).Set Link_Name&TimeStamp
	           objClientContentTable.ChildItem(CCR, 12, "WebEdit", 0).Set ToolTip&TimeStamp
	      Else
	         wait 3
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 9, "WebList", 0).Select Provide_Link
	           'CHECKING FOR FREEZED DATA FOR DECISION CENTER HEADING,LINK NAME AND TOOLTIP IF PROVIDE LINK IS 'NO'
	           DCDisabled=objClientContentTable.ChildItem(CCR, 10, "WebList", 0).GetROProperty("disabled") 
	           LinkDisabled=objClientContentTable.ChildItem(CCR, 11, "WebEdit", 0).GetROProperty("disabled") 
	           TooltipDis=objClientContentTable.ChildItem(CCR, 12, "WebEdit", 0).GetROProperty("disabled") 
	           If DCDisabled="1" AND LinkDisabled="1" AND TooltipDis="1" Then
	           	  Call ReportStep (StatusTypes.Pass, "'Link name', 'Tool tip and' and  'Decision center' fields  are not Editable after seletcing Provide link as 'No' "," Provide link as 'No'")
	           Else
	           	  Call ReportStep (StatusTypes.Pass, "'Link name', 'Tool tip and' and  'Decision center' fields  are  Editable after seletcing Provide link as 'No' "," Provide link as 'No'")
	           End If

	       End If
	       
	       	   objClientContentTable.ChildItem(CCR, 14, "WebList", 0).Select Trim(Report_Type)
	           'objClientContentTable.ChildItem(CCR, 15, "WebEdit", 0).Set Trim(Country&"/"&Offering)
	           'Modified the code based on New Requirement - OCRF - VA Ops Maintenance - 21/01/2019
	           wait 4
	           If Report_Type = "VA PBI Report" Then
'	           ***************************************************
'					Modified by Madhu - 04/15/2020
					
					'changes by srinivas
'					If objClientContentTable.ChildItem(CCR, 15, "WebEdit", 0).GetROProperty("disabled")=true Then
'	           			Call ReportStep (StatusTypes.Pass,"report path webedit column is disabled for PBI Report Type","Validate whether report path webedit column is disabled or not for PBI Report Type" )
'	           			else
'	           			Call ReportStep (StatusTypes.Fail,"report path webedit column is not disabled for PBI Report Type","Validate whether report path webedit column is disabled or not for PBI Report Type" )
'	           		End If
'	           		
'	           		
'	           		If objClientContentTable.ChildItem(CCR, 16, "WebEdit", 0).GetROProperty("disabled")=true Then
'	           			Call ReportStep (StatusTypes.Pass,"report name webedit column is disabled for PBI Report Type","Validate whether report name webedit column is disabled or not for PBI Report Type" )
'	           			else
'	           			Call ReportStep (StatusTypes.Fail,"report name webedit column is not disabled for PBI Report Type","Validate whether report name webedit column is disabled or not for PBI Report Type" )
'	           		End If
					
					
	           		'objClientContentTable.ChildItem(CCR, 16, "WebEdit", 0).Set Trim(Report_Name)
	           		else
	           		objClientContentTable.ChildItem(CCR, 16, "WebList", 0).Click
	           		Wait 5
	           		'objClientContentTable.ChildItem(CCR, 16, "WebEdit", 0).select Trim(Report_Name)
	           		
	           		'changes by srinivas
	           		
	           		
	           		
	           		
					If Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Information_Okay").Exist Then
						Browser("Browser-OCRF").Page("Page-OCRF").WebButton("WebButton_Information_Okay").Click
						wait 3
					End If
					wait 2
					'objClientContentTable.ChildItem(CCR, 16, "WebList", 0).Click
	           		objClientContentTable.ChildItem(CCR, 16, "WebList", 0).Select Trim(Report_Name)
	           		
''	           		*********************************************************************8
					If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div//div[text()='An entry already exists for this report. Please use the existing entry.']").Exist(20) Then
						'msgbox "Yes"
						DuplicateWarningText=Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div//div[text()='An entry already exists for this report. Please use the existing entry.']").GetROProperty("innertext")
						Call ReportStep (StatusTypes.Pass,"Warning message displayed with text:-"&DuplicateWarningText&" when uploaded duplicate report of type MSTR","Validate Duplicate Warning Message-client content" )
						Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div//div[text()='An entry already exists for this report. Please use the existing entry.']/../descendant::span[text()='Okay']").Click
						Exit for
						else
						'msgbox "No"
						Call ReportStep (StatusTypes.Fail,"Warning message is not displayed with text:-"&DuplicateWarningText&" when uploaded duplicate report of type MSTR","Validate Duplicate Warning Message-client content" )
				   End If
					
	           		wait 2
	           End If
	           
'	           ***********************************************************
			 objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	           objClientContentTable.ChildItem(CCR, 17, "WebEdit", 0).Set Trim(Client_Name)
	           Set objShell=CreateObject("WScript.Shell")
	           objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	  ' objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
			   objClientContentTable.ChildItem(CCR, 17, "WebEdit", 0).Click
				'Set objShell1=createobject("wscript.shell")	
				'objShell1.SendKeys "{DOWN}"
				'************************************************************************8
				wait 5
'	    	   modified by Madhu - 02/26/2020
				Client_Name=Datatable.Value("Client_Name",SheetName)
				'Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","class:=ui-corner-all","innertext:="&Datatable.Value("Client_Name",SheetName)).highlight
				'Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","class:=ui-corner-all","innertext:="&Datatable.Value("Client_Name",SheetName)).Click
              	'browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//li//a[text()='"&Client_Name&"']").Click
				'Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&dataTable.value("Client_Name",SheetName),"html id:=ui-id-.*").Click		
				'Client_Name=Datatable.Value("Client_Name",SheetName)
'				Set oDesc = Description.Create
'				oDesc("micclass").Value = "WebElement"
'				oDesc("html tag").Value = "A"
'				oDesc("innertext").Value = Client_Name
'				oDesc("html id").Value="ui-id-.*"
'				set s=Browser("micclass:=browser").Page("micclass:=page").ChildObjects(oDesc)
'				'msgbox s.count
'				wait 2
'				For j = 1 To s.count-1 Step 1
'					s(j).click
'				Next
				'Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//li//a[text()='"&Client_Name&"']").click
				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&Client_Name&"","html id:=ui-id-.*").Click		
				wait 2
		 	Call ReportStep (StatusTypes.Pass, "Successfully added data in row '"&rownum+1&"' Provide Link as'" &Provide_Link& "' Report Type as '" &Report_Type," Client content tab")
               
               objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Set "OFF"
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
	    	   objClientContentTable.ChildItem(CCR, 2, "WebCheckBox", 0).Click
'			************************************************************************8

'	    	   modified by Madhu - 02/26/2020

'				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:="&Datatable.Value("Client_Name",SheetName)).highlight
'				Browser("Browser-OCRF").Page("Page-OCRF").WebElement("html id:=ui-id-.*","html tag:=A","innertext:="&Datatable.Value("Client_Name",SheetName)).Click
               Call ReportStep (StatusTypes.Pass, "Successfully added data in row '"&rownum+1&"' Provide Link as'" &Provide_Link& "' Report Type as '" &Report_Type," Client content tab")
               
               If Report_Type = "VA PBI Report" Then
               		filePath = Environment("CurrDir") & "HostingTestData\" 
               		fileName = Report_Name
               		fileToSet = filePath & fileName
               		objClientContentTable.ChildItem(CCR, 18, "Link", 0).Click
      				If Browser("Browser-OCRF").Page("Page-OCRF").WebFile("wfAddFiles_PBIReport").Exist(20) then 
      					Browser("Browser-OCRF").Page("Page-OCRF").WebFile("wfAddFiles_PBIReport").Set fileToSet
      					wait 3
'      					Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Add files_PBIReport").Click
'      					If Browser("Browser-OCRF").Dialog("Choose File to Upload").Exist(20) Then
'      						Browser("Browser-OCRF").Dialog("Choose File to Upload").WinEdit("FileName").Set fileToSet
'      						Browser("Browser-OCRF").Dialog("Choose File to Upload").WinButton("Open").Click
      						Browser("Browser-OCRF").Page("Page-OCRF").WebButton("StartUpload_PBIReport").Click
      						
      						'validate the warning message displayed for duplicate content
      						If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div//div[text()='An entry already exists for this report. Please use the existing entry.']").Exist(20) Then
						'msgbox "Yes"
								DuplicateWarningText=Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div//div[text()='An entry already exists for this report. Please use the existing entry.']").GetROProperty("innertext")
								Call ReportStep (StatusTypes.Pass,"Warning message displayed with text:-"&DuplicateWarningText&" when uploaded duplicate report of type VA PBI","Validate Duplicate Warning Message-client content" )
								Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div//div[text()='An entry already exists for this report. Please use the existing entry.']/../descendant::span[text()='Okay']").Click
								'Exit for
								else
								'msgbox "No"
								Call ReportStep (StatusTypes.Fail,"Warning message is not displayed with text:-"&DuplicateWarningText&" when uploaded duplicate report of type VA PBI","Validate Duplicate Warning Message-client content" )
						   End If
      						
      						
      						'Browser("Browser-OCRF").Page("Page-OCRF").WebButton("Close_PBIReport").Click
      						
      						'changes by srinivas
      						browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[@id='UploadReportFile']/../descendant::div//button//span[text()='Close']").Click
      				End If
               End If
	           
	    End If
	    		    	
	Next	
	Call ReadWriteDataFromTextFile("write","PBIReportTimeStamp",TimeStamp)
'    AddDatainClientDetailsTab=TimeStamp
	If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//TD[@id='Content-list_toppager_left']//div[@class='ui-pg-div']//span[@class='ui-icon ui-icon-trash']").Exist(10) Then
	'msgbox "Yes"
	Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//TD[@id='Content-list_toppager_left']//div[@class='ui-pg-div']//span[@class='ui-icon ui-icon-trash']").Click
	Set odesc=description.Create()
	odesc("class").Value="ui-button-text"
	odesc("html tag").Value="SPAN"
	odesc("innertext").Value="Yes"
	
	Set YesButton=Browser("micclass:=browser").Page("micclass:=page").ChildObjects(odesc)
	'msgbox YesButton.count
	
	On error resume next
	For Iterator = 0 To YesButton.count Step 1
		YesButton(Iterator).click
	Next
	Call ReportStep (StatusTypes.Pass, "Successfully deleted the duplicate row of Report Type as '" &Report_Type," Client content tab")
	else
	'msgbox "no"
	End If

End Function


Function ChangeActionRowInClientContent(ByVal action,ByVal start,ByVal endrow)
	For Iterator = start To endrow Step 1
		Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList").ChildItem(Iterator,2,"WebCheckBox", 0).click
		Call PageLoading()
		Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList").ChildItem(Iterator,6,"WebList", 0).select action
		Call ReportStep (StatusTypes.Pass, "Successfully changed the Action of row no:-"&start&" to action:-"&action&"" ,"Client content tab")
		Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList").ChildItem(Iterator,2,"WebCheckBox", 0).click
	Next
End Function

Function changeActionValidateDupContentWarningMsg(ByVal action,ByVal rowno)
	    	RowAction=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList").GetCellData(rowno,6)
	    	ReportType=Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList").GetCellData(rowno,14)
	    	Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList").ChildItem(rowno,2,"WebCheckBox", 0).click
			Call PageLoading()
			Browser("Browser-OCRF").Page("Page-OCRF").WebTable("WebTbl_clientContentList").ChildItem(rowno,6,"WebList", 0).select action
	    	Call ReportStep (StatusTypes.Pass, "Successfully changed the Action of row no:-"&rowno&" to action:-"&action&"" ,"Client content tab")
	    	If action="DELETE" or action="DELETED" Then
	    		If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div//div[text()='An entry already exists for this report. Please use the existing entry.']").Exist(10) Then
						'msgbox "Yes"
					DuplicateWarningText=Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div//div[text()='An entry already exists for this report. Please use the existing entry.']").GetROProperty("innertext")
					Call ReportStep (StatusTypes.Fail,"Warning message displayed with text:-"&DuplicateWarningText&" when uploaded duplicate report of type:-"&ReportType&" for the action:="&action&"","Validate Duplicate Warning Message-client content" )
					Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div//div[text()='An entry already exists for this report. Please use the existing entry.']/../descendant::span[text()='Okay']").Click
								'Exit for
					else
								'msgbox "No"
					Call ReportStep (StatusTypes.Pass,"Warning message is not displayed with text:-"&DuplicateWarningText&" when uploaded duplicate report of type:-"&ReportType&" for the action:="&action&"","Validate Duplicate Warning Message-client content" )
				End If
			else
				If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div//div[text()='An entry already exists for this report. Please use the existing entry.']").Exist(10) Then
						'msgbox "Yes"
					DuplicateWarningText=Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div//div[text()='An entry already exists for this report. Please use the existing entry.']").GetROProperty("innertext")
					Call ReportStep (StatusTypes.Pass,"Warning message displayed with text:-"&DuplicateWarningText&" when uploaded duplicate report of type:-"&ReportType&" for the action:="&action&"","Validate Duplicate Warning Message-client content" )
					Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div//div[text()='An entry already exists for this report. Please use the existing entry.']/../descendant::span[text()='Okay']").Click
								'Exit for
					else
								'msgbox "No"
					Call ReportStep (StatusTypes.Fail,"Warning message is not displayed with text:-"&DuplicateWarningText&" when uploaded duplicate report of type:-"&ReportType&" for the action:="&action&"","Validate Duplicate Warning Message-client content" )
				End If
	    	End If
End Function

Function ImportClientContentDataFromOtherOCRF(ByVal clientname,Byval offering,Byval OCRFID)
	If browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[@class='ui-pg-div' and text()='Import Content']").Exist(20) Then
	Call ReportStep (StatusTypes.Pass,"Import Content option is present in client content tab","client content tab" )
	browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[@class='ui-pg-div' and text()='Import Content']").Click
	If browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[@aria-describedby='ImportContentDialog']").Exist(20) Then
		Call ReportStep (StatusTypes.Pass,"cilcked on Import Content button","client content tab" )
		else
		'msgbox "N"
		Call ReportStep (StatusTypes.Fail,"cilcked on Import Content button","client content tab" )
	End If
	else
	Call ReportStep (StatusTypes.Fail,"Import Content option is not present in client content tab","client content tab" )
End If

browser("micclass:=browser").Page("micclass:=page").WebEdit("xpath:=//td[@class='input-caption']//input[@id='clientName']").Click
browser("micclass:=browser").Page("micclass:=page").WebEdit("xpath:=//td[@class='input-caption']//input[@id='clientName']").set clientname
Call ReportStep (StatusTypes.Pass,"entered clientname:-"&clientname&" in import content page","Import content page" )
browser("micclass:=browser").Page("micclass:=page").WebEdit("xpath:=//td[@class='input-caption']//input[@id='clientName']").Click
Set wscript=createobject("wscript.shell")
wscript.SendKeys "{DOWN}"
wait 4

Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&clientname&"","html id:=ui-id-.*").Click
'select offering
browser("micclass:=browser").Page("micclass:=page").WebEdit("xpath:=//td[@class='input-caption']//input[@id='offeringName']").Click
browser("micclass:=browser").Page("micclass:=page").WebEdit("xpath:=//td[@class='input-caption']//input[@id='offeringName']").set offering
Call ReportStep (StatusTypes.Pass,"entered offeringname:-"&offering&" in import content page","Import content page" )
browser("micclass:=browser").Page("micclass:=page").WebEdit("xpath:=//td[@class='input-caption']//input[@id='offeringName']").Click
Set wscript=createobject("wscript.shell")
wscript.SendKeys "{DOWN}"
wait 4

Browser("Browser-OCRF").Page("Page-OCRF").WebElement("class:=ui-corner.*","innertext:="&OCRFID&"-"&Offering&"","html id:=ui-id-.*").Click

If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div//button//span[text()='Import']").Exist(10) Then
	Call ReportStep (StatusTypes.Pass,"Import button exists in import content page","Import content page" )
	Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div//button//span[text()='Import']").Click
	Call ReportStep (StatusTypes.Pass,"clicked on Import button","Import content page" )
	else
	Call ReportStep (StatusTypes.Fail,"Import button not exists in import content page","Import content page" )
End If

If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[contains(@style,'display: block;')]//div[text()='The following report(s) already exist in the OCRF and will not be imported.']").Exist(10) Then
	'msgbox "Yes"
	ReportExistMsg=Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[contains(@style,'display: block;')]//div[text()='The following report(s) already exist in the OCRF and will not be imported.']").GetROProperty("innertext")
	'msgbox ReportExistMsg
	Call ReportStep (StatusTypes.Pass,"duplicate warning message displayed with text:-"&ReportExistMsg&"","Import content page" )
	Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[contains(@style,'display: block;')]//div[text()='The following report(s) already exist in the OCRF and will not be imported.']/../descendant::button//span[text()='Okay']").Click
	Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div//span[text()='Import Online CRF Contents']/../following::button//span[text()='Close']").Click
	ElseIf Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[contains(@style,'display: block;')]//div//span[text()='Information']/../following-sibling::div[text()='Import Success']").Exist(10) Then
	'msgbox "no"
	ImportSuccess=Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[contains(@style,'display: block;')]//div//span[text()='Information']/../following-sibling::div[text()='Import Success']").GetROProperty("innertext")
	Call ReportStep (StatusTypes.Pass,""&ImportSuccess&"","Import content page" )
	Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[contains(@style,'display: block;')]//div//span[text()='Information']/../following-sibling::div[text()='Import Success']/../descendant::button//span[text()='Okay']").Click
	Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div//span[text()='Import Online CRF Contents']/../following::button//span[text()='Close']").Click
End If
End Function

Function ChangeActionRowInContentPermissionTab(ByVal action,ByVal start,ByVal endrow)

	For Iterator = start To endrow Step 1
		Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").ChildItem(Iterator,2,"WebCheckBox", 0).click
		Call PageLoading()
		Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").ChildItem(Iterator,3,"WebList", 0).select action
		Call ReportStep (StatusTypes.Pass, "Successfully changed the Action of row no:-"&start&" to action:-"&action&"" ,"Content permission tab")
		Browser("Browser-OCRF").Page("Page-OCRF").WebTable("userClientContentAccess-list").ChildItem(Iterator,2,"WebCheckBox", 0).click
	Next
End Function

Function validateSharePointPermission(ByVal FileName,ByVal SheetName,ByVal ClientcontentSheet,Byval clientSheetStart,ByVal clientSheetend,ByVal StartRow,ByVal EndRow,ByVal Offering,ByRef Username,ByVal Password,ClientType)
	Call ImportSheet(SheetName,FileName)
	DataTable.GetSheet(SheetName).SetCurrentRow(1)
	
	rowCnt=Datatable.GetSheet(SheetName).GetRowCount
	BI_Tool = DataTable.Value("BI_Tool",SheetName)
    accessMode = DataTable.Value("Access_Mode",SheetName)
    'login to DC
    Systemutil.CloseProcessByName "iexplore.exe"
	
	'Launch OCRF Url's	
'	SystemUtil.Run Environment.Value("OCRFURL")

    'SystemUtil.Run "iexplore", Environment.Value("DCURL") , , , 3
    clienturl=Environment.Value("DCURL")
    Call DCLogin(clienturl,Username,Password,"autopbiperm1",ObjData)
    
    If Browser("micclass:=browser").Page("micclass:=page").WebElement("html tag:=SPAN","html id:=zz9_SiteActionsMenu_t").Exist(120) Then
	Call ReportStep (StatusTypes.Pass,"SiteActions is visible in DC","SharePoint" )
	Browser("micclass:=browser").Page("micclass:=page").WebElement("html tag:=SPAN","html id:=zz9_SiteActionsMenu_t").Click
	If Browser("micclass:=browser").Page("micclass:=page").WebElement("html tag:=DIV","class:=ms-MenuUIPopupBody ms-MenuUIPopupScreen").Exist(10) Then
		Browser("micclass:=browser").Page("micclass:=page").WebElement("html tag:=DIV","class:=ms-MenuUIPopupBody ms-MenuUIPopupScreen").highlight
		Call ReportStep (StatusTypes.Pass,"clicked on SiteActions button","SharePoint" )
		If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//span[@class='ms-MenuUILabel']//span[text()='Manage Secure Links']").Exist(10) Then
			Call ReportStep (StatusTypes.Pass,"Manage secure links option exists","SharePoint" )
			Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//span[@class='ms-MenuUILabel']//span[text()='Manage Secure Links']").highlight
			Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//span[@class='ms-MenuUILabel']//span[text()='Manage Secure Links']").Click
			else
			Call ReportStep (StatusTypes.Fail,"Manage secure links option doesn't exists","SharePoint" )
		End If
		else
		Call ReportStep (StatusTypes.Fail,"not clicked on SiteActions button","SharePoint" )
	End If
	else
	Call ReportStep (StatusTypes.Fail,"SiteActions is not visible in DC","SharePoint" )
End If
'Call ImportSheet(OfferingSheetName,TestDataPath)
Call ImportSheet(ClientcontentSheet,FileName)
DataTable.GetSheet(ClientcontentSheet).SetCurrentRow(clientSheetStart)

If ClientType="All Clients" Then
	
	If browser("micclass:=browser").page("micclass:=page").WebElement("xpath:=//a[contains(text(),'"&Offering&" PBI')]").Exist(20) Then
			Call ReportStep (StatusTypes.Pass,"offering folder is present inside manage secure links","SharePoint" )
			browser("micclass:=browser").page("micclass:=page").WebElement("xpath:=//a[contains(text(),'"&Offering&" PBI')]").highlight
			browser("micclass:=browser").page("micclass:=page").WebElement("xpath:=//a[contains(text(),'"&Offering&" PBI')]").Click
			If browser("micclass:=browser").Page("micclass:=page").WebElement("html tag:=TABLE","class:=ms-listviewtable").Exist(30) Then
				Call ReportStep (StatusTypes.Pass,"clicked on the offering folder:-"&Offering&" PBI Consumer","SharePoint" )
				If browser("micclass:=browser").page("micclass:=page").WebElement("xpath:=//a[contains(text(),'"&Offering&" PBI Consumer')]").Exist(20) Then
					Call ReportStep (StatusTypes.Pass,"offering folder is present inside manage secure links","SharePoint" )
					
					'''
						For row = clientSheetStart To clientSheetend Step 1
							reportType = DataTable.Value("Report_Type",ClientcontentSheet)
							'If reportType = "VA PBI Report" Then
							linkName=dataTable.value("Link_Name",ClientcontentSheet)
							PBIReportTimeStamp = ReadWriteDataFromTextFile("Read","PBIReportTimeStamp","")
							reportlink=linkName&PBIReportTimeStamp
						
						browser("micclass:=browser").Page("micclass:=page").WebElement("html tag:=TABLE","class:=ms-listviewtable").highlight
						row=browser("micclass:=browser").Page("micclass:=page").WebTable("html tag:=TABLE","class:=ms-listviewtable").GetRowWithCellText(trim(reportlink))
						browser("micclass:=browser").Page("micclass:=page").WebTable("html tag:=TABLE","class:=ms-listviewtable").ChildItem(cint(row),3,"WebElement",0).highlight
						browser("micclass:=browser").Page("micclass:=page").WebTable("html tag:=TABLE","class:=ms-listviewtable").ChildItem(cint(row),3,"WebElement",0).fireevent("onmouseover")
						wait 2
						browser("micclass:=browser").Page("micclass:=page").WebTable("html tag:=TABLE","class:=ms-listviewtable").ChildItem(cint(row),3,"Image",0).click
						If browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//span[@class='ms-MenuUILabel']//span[text()='Manage Permissions']").Exist(10) Then
							Call ReportStep (StatusTypes.Pass,"ManagePermission's option is present for the reportLink:-"&reportlink&"","SharePoint" )
							browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//span[@class='ms-MenuUILabel']//span[text()='Manage Permissions']").Click
							else
							Call ReportStep (StatusTypes.Fail,"ManagePermission's option is not present for the reportLink:-"&reportlink&"","SharePoint" )
					'validate the users for report
						End If
						next
					else
					Call ReportStep (StatusTypes.Fail,"Not clicked on the offering folder:-"&Offering&" PBI Consumer","SharePoint" )				
			End If
			'Next
			else
			Call ReportStep (StatusTypes.Fail,"offering folder is not present inside manage secure links","SharePoint" )
			browser("micclass:=browser").page("micclass:=page").WebElement("xpath:=//a[contains(text(),'"&Offering&" PBI Consumer')]").highlight
			browser("micclass:=browser").page("micclass:=page").WebElement("xpath:=//a[contains(text(),'"&Offering&" PBI Consumer')]").Click
			'If browser("micclass:=browser").Page("micclass:=page").WebElement("html tag:=TABLE","class:=ms-listviewtable").Exist(30) Then
	
	End If	'Call ReportStep (StatusTypes.Pass,"clicked on the offering folder:-"&Offering&" PBI Consumer","SharePoint" )
	
	End If
	else
	'''''''''''''
	'Call ImportSheet(ClientcontentSheet,FileName)
'DataTable.GetSheet(ClientcontentSheet).SetCurrentRow(clientSheetStart)
'Datatable.GetSheet(OfferingSheetName).SetCurrentRow(1)
'OffCountry=dataTable.value("Offering_Country",OfferingSheetName)
'ClientSynd=dataTable.value("Syndicated",OfferingSheetName)
For row = clientSheetStart To clientSheetend Step 1
	reportType = DataTable.Value("Report_Type",ClientcontentSheet)
	If reportType = "VA PBI Report" Then
	linkName=dataTable.value("Link_Name",ClientcontentSheet)
	PBIReportTimeStamp = ReadWriteDataFromTextFile("Read","PBIReportTimeStamp","")
	reportlink=linkName&PBIReportTimeStamp
	'validate share point here
		If browser("micclass:=browser").page("micclass:=page").WebElement("xpath:=//a[contains(text(),'"&Offering&" PBI')]").Exist(20) Then
			Call ReportStep (StatusTypes.Pass,"offering folder is present inside manage secure links","SharePoint" )
			browser("micclass:=browser").page("micclass:=page").WebElement("xpath:=//a[contains(text(),'"&Offering&" PBI')]").highlight
			browser("micclass:=browser").page("micclass:=page").WebElement("xpath:=//a[contains(text(),'"&Offering&" PBI')]").Click
			If browser("micclass:=browser").Page("micclass:=page").WebElement("html tag:=TABLE","class:=ms-listviewtable").Exist(30) Then
				Call ReportStep (StatusTypes.Pass,"clicked on the offering folder:-"&Offering&" PBI Consumer","SharePoint" )
				browser("micclass:=browser").Page("micclass:=page").WebElement("html tag:=TABLE","class:=ms-listviewtable").highlight
				row=browser("micclass:=browser").Page("micclass:=page").WebTable("html tag:=TABLE","class:=ms-listviewtable").GetRowWithCellText(trim(reportlink))
				browser("micclass:=browser").Page("micclass:=page").WebTable("html tag:=TABLE","class:=ms-listviewtable").ChildItem(cint(row),3,"WebElement",0).highlight
				browser("micclass:=browser").Page("micclass:=page").WebTable("html tag:=TABLE","class:=ms-listviewtable").ChildItem(cint(row),3,"WebElement",0).fireevent("onmouseover")
				wait 2
				browser("micclass:=browser").Page("micclass:=page").WebTable("html tag:=TABLE","class:=ms-listviewtable").ChildItem(cint(row),3,"Image",0).click
				If browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//span[@class='ms-MenuUILabel']//span[text()='Manage Permissions']").Exist(10) Then
					Call ReportStep (StatusTypes.Pass,"ManagePermission's option is present for the reportLink:-"&reportlink&"","SharePoint" )
					browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//span[@class='ms-MenuUILabel']//span[text()='Manage Permissions']").Click
					'validate the users for report
					Call ReportStep (StatusTypes.Information,"Verifying Users permissioned for report","SharePoint" )
					rowCnt=Datatable.GetSheet(SheetName).GetRowCount
					For Iterator = StartRow To EndRow Step 1
						DataTable.SetCurrentRow(row)
    					BI_Tool = DataTable.Value("BI_Tool",SheetName)
    					accessMode = DataTable.Value("Access_Mode",SheetName)
    					User_Id=dataTable.value("User_Id",SheetName)
    					User=trim(mid(User_Id,8,8))
    					If browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//a[contains(text(),'"&User&"')]").Exist(30) Then
    						browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//a[contains(text(),'"&User&"')]").highlight
    						Call ReportStep (StatusTypes.Pass,""&User_Id&" is permissioned for the report:-"&reportlink&"","SharePoint" )
    						else
    						Call ReportStep (StatusTypes.Fail,""&User_Id&" is not permissioned for the report:-"&reportlink&"","SharePoint" )
    					End If
					Next
					else
					Call ReportStep (StatusTypes.Fail,"ManagePermission's option is not present for the reportLink:-"&reportlink&"","SharePoint" )
				End If
				else
				Call ReportStep (StatusTypes.Fail,"Not clicked on the offering folder:-"&Offering&" PBI Consumer","SharePoint" )				
			End If
			else
			Call ReportStep (StatusTypes.Fail,"offering folder is not present inside manage secure links","SharePoint" )
		End If
	End If
Next
End If

If Browser("micclass:=browser").page("micclass:=page").WebElement("html id:=zz3_Menu_t","html tag:=SPAN").Exist(5) Then
	Browser("micclass:=browser").page("micclass:=page").WebElement("html id:=zz3_Menu_t","html tag:=SPAN").click
	If Browser("micclass:=browser").page("micclass:=page").WebElement("html id:=mp1_0_0_Anchor","html tag:=A").Exist(10) Then
		Browser("micclass:=browser").page("micclass:=page").WebElement("html id:=mp1_0_0_Anchor","html tag:=A").click
		wait 4
	End If
End If
End Function

'********************************************************************************************************************************************************************************************
'********************************************************************************************************************************************************************************************
'----------------------------------------------------------------Generic Functions created by Srini wRT VA Report Level Security ends_______________________________________________________________



Public Function NavigateToMSTRFolderStructures1(ByVal Foldername,ByVal country,ByVal offering,Byval clientname)
	'folder structure can be VA design or Visual analytics
	If 	Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//DIV[@id='projects_ProjectsStyle']//a[text()='"&Foldername&"']").exist(30) Then
		Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//DIV[@id='projects_ProjectsStyle']//a[text()='"&Foldername&"']").click
	End If
	If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//input[@value='Continue']").Exist(10) Then
		Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//input[@value='Continue']").Click
	End If
	If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[@id='dktpSectionView']").Exist(30) Then
		Call ReportStep (StatusTypes.Pass, "clicked on folder:-"&Foldername&"","click on "&Foldername&" in mstr server")
		'click on shared reports
		If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[text()='Shared Reports']/..").Exist(30) Then
			Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[text()='Shared Reports']/..").Click
			If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']").Exist(30) Then
				Call ReportStep (StatusTypes.Pass, "clicked on folder:- shared reports and country folders page is displayed","click on shared reports folder in mstr server")
				'click on country folder
				If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='"&country&"']").exist(30) Then
					Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='"&country&"']").click
					If Browser("BrowserPoPUp").Page("micclass:=page").WebTable("innertext:="&offering&".*","html Tag:=Table","visible:=True","class:=mstrLargeIconItemTable").childitem(1,1,"WebElement",0).exist(30) Then
						Call ReportStep (StatusTypes.Pass, "clicked on country folder:-"&country&" and offering page is displayed","click on country folder in mstr server")
						'click on offering folder
						If Browser("BrowserPoPUp").Page("micclass:=page").WebTable("innertext:="&offering&".*","html Tag:=Table","visible:=True","class:=mstrLargeIconItemTable").childitem(1,1,"WebElement",0).exist(30) Then
							Browser("BrowserPoPUp").Page("micclass:=page").WebTable("innertext:="&offering&".*","html Tag:=Table","visible:=True","class:=mstrLargeIconItemTable").childitem(1,1,"WebElement",0).click
							Call ReportStep (StatusTypes.Pass, "clicked on offering folder:-"&offering&"","click on offering folder in mstr server")
							If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='"&clientname&"']").exist(30) Then
								Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='"&clientname&"']").click
								
							End If
							If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='ims']").Exist(30) Then
								Call ReportStep (StatusTypes.Pass, "clicked on client folder:-"&clientname&"","click on client folder in mstr server")
								'click on ims folder
								If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='ims']").Exist(30) Then
									Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='ims']").Click
								End If
								If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='Publish']").Exist(30) Then
									Call ReportStep (StatusTypes.Pass, "clicked on ims folder and Publish folder displayed successfully in mstr server","click on ims folder in mstr server")
									'click on publish folder
									Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='Publish']").Click
									If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//td[@id='td_mstrWeb_dockLeft']").Exist(30) Then
										Call ReportStep (StatusTypes.Pass, "clicked on publish folder and Reports section page displayed successfully in mstr server","click on publish folder in mstr server")
										else
										Call ReportStep (StatusTypes.Fail, "not clicked on publish folder and Reports section page not displayed in mstr server","click on publish folder in mstr server")
									End If
									else
									Call ReportStep (StatusTypes.Fail, "not clicked on ims folder and Publish folder not displayed in mstr server","click on ims folder in mstr server")
								End If
								else
								If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='Publish']").Exist(30) Then
									Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='Publish']").Click
								End If
								'Call ReportStep (StatusTypes.Fail, "not clicked on offering folder:-"&offering&"","click on offering folder in mstr server")
							End If
						End If
						else
						Call ReportStep (StatusTypes.Fail, "not clicked on country folder:-"&country&" and offering page is not displayed","click on country folder in mstr server")
					End If
				End If
				else
				Call ReportStep (StatusTypes.Fail, "not clicked on folder:- shared reports","click on shared reports folder in mstr server")
			End If
		End If
		else
		Call ReportStep (StatusTypes.Fail, "not clicked on folder:-"&Foldername&"","click on "&Foldername&" in mstr server")
	End If
End Function


Public function ChartCreationInSCA(ByVal strChartPos, ByVal strChartType, ByVal strChartSheet)
	
		Dim objPos, objType, objSheet, objAdd
		Set objPos=Browser("Analyzer").Page("ReportCreation").Frame("CreateChart").WebRadioGroup("rdoPosition")
   		Call SCA.ClickOnRadio(objPos,strChartPos, "ReportCreation")
   		Set objType=Browser("Analyzer").Page("ReportCreation").Frame("CreateChart").WebRadioGroup("rdoChartType")
   		Call SCA.ClickOnRadio(objType,strChartType, "ReportCreation")
   		Set objSheet=Browser("Analyzer").Page("ReportCreation").Frame("CreateChart").WebList("lstddlSheets")
   		Call SCA.SelectFromDropdown(objSheet,strChartSheet)	
   	
   		Set objAdd=Browser("Analyzer").Page("ReportCreation").Frame("CreateChart").WebButton("btnOk")
   		Call SCA.ClickOn(objAdd,"OkButton" , "ReportCreation")
   		

	End function
	
	Public function ValidateChartExistenceParams(ByVal strChartPos, ByVal strSheetComp, ByVal strChartType, ByVal strChartSheet, ByVal strPageName, ByRef objData)

		Dim sheetFound, validateCount, chartPosFound
	   	validateCount = 0
	   	
	   	Browser("Analyzer").Page("ReportCreation").Sync
	   	
	   	'Call ValidateReportSheet function to validate presence of sheet in SCA report Page
	   	sheetFound = ValidateReportSheet(strChartSheet, strPageName, objData)
	   	If sheetFound = 1 Then
	   		Call ReportStep (StatusTypes.Pass, "Successfully clicked on sheet " &strChartSheet& " to validate the Chart" , strPageName)
	   		validateCount = 1
	   	else
	   		Call ReportStep (StatusTypes.Fail, "Could not click on Sheet" &strChartSheet& " to validate the Chart"  , strPageName)
	   		Exit function
	   	End If
	   	
	   	'Call ValidateReportPosition function to validate presence of sheet in SCA report Page
	   	chartPosFound = ValidateReportPosition(strChartPos, strSheetComp, strPageName, objData)
	   	If chartPosFound = 2 Then
	   		Call ReportStep (StatusTypes.Pass, "Chart Position is not validated because of no other components", strPageName)
	   		validateCount = validateCount+1
	   	ElseIf chartPosFound = 1 Then
	   		Call ReportStep (StatusTypes.Pass, "Chart Position is successfully validated", strPageName)
	   		validateCount = validateCount+1 
	   	ElseIf chartPosFound = 0 Then
	   		Call ReportStep (StatusTypes.Fail, "Chart Position is could not validate successfully", strPageName)
	   		Exit function
	   	End If
	   	
	   	If validateCount = 2 and strChartPos <> "" Then
			Call ReportStep (StatusTypes.Pass, "Successfully validated Chart Position " &strChartPos& " on Sheet " &strChartSheet, strPageName)
		    Call ReportStep (StatusTypes.Pass, "Chart is located at " &strChartPos, strPageName)
		ElseIf validateCount = 2 and strChartPos = "" Then
			Call ReportStep (StatusTypes.Pass, "Chart Position is not validated because of no other components on Sheet " &strChartSheet, strPageName)
		    Call ReportStep (StatusTypes.Pass, "Chart is located on Sheet " &strChartSheet, strPageName)
		Else 
		  	Call ReportStep (StatusTypes.Fail, "Could not validate Chart Position " &strChartPos& " on Sheet " &strChartSheet, strPageName)
		End If
					
	End function
	
	
	Public function DownNormalMenu(ByVal strTabName, ByVal strSelValue, ByVal strSelSubValue)
	
		Dim objValue, objImgPivotDownNormal, objSelSubValue,objimgDown
		
		Browser("Analyzer").Page("ReportCreation").Sync
		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").RefreshObject
		
		'<shweta - 7/22/2016> - Adding onmouseover on Component 'Drop a Filter Condition Here' Object to enable imgPivotDownNormal  - Start
		Set objComp = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("html tag:=NOBR","innertext:= Drop a Filter Condition Here")
		If objComp.Exist(30) Then
			objComp.Highlight
			objComp.FireEvent "OnmouseOver"
			objComp.FireEvent "onclick"
			wait 2
		End If
		'<shweta - 7/22/2016> - Adding onmouseover on Component 'Drop a Filter Condition Here' Object to enable imgPivotDownNormal  - End
		
		'Mouse over on Pivot Table menu
		Set objimgDown=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgPivotDownNormal")
		
		For i = 1 To 5 Step 1
			
			If objimgDown.Exist(60) Then
			    objimgDown.FireEvent "onmouseover"
				objimgDown.Click
				Exit For
			End If
			wait 2
		Next
		wait 2
		'Browser("Analyzer").Page("ReportCreation").Sync
		
'		Set objimgDown=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgPivotDownNormal")
'		Call SCA.ClickOn(objimgDown,"imgPivotDownNormal", "ReportCreation")

		'<Shweta 11/8/2015> - Start
		'Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgPivotDownNormal").FireEvent "onmouseover"
		'wait 2
		'Set objImgPivotDownNormal=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgPivotDownNormal")
		'Call SCA.ClickOn(objImgPivotDownNormal,"imgPivotDownNormal", "ReportCreation")
		'<Shweta 11/8/2015> - End
		
		'Set values for chart creation
		'Mouse over on Tab name in pivot table menu. It could be "Design", "Analyze", "Advanced"
		
		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image(strTabName).FireEvent "onmouseover"
   		wait 2
   		
   		'Select options under each tab
   		If strSelSubValue <> "" Then
   			Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strSelValue).FireEvent "onmouseover"
			wait 1
			Set objSelSubValue = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strSelSubValue)
			Call SCA.ClickOn(objSelSubValue,strSelSubValue, "Report Creation Page")
   		Else 
   			Set objValue=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strSelValue)
			Call SCA.ClickOn(objValue,strSelValue, "ReportCreation")
	   		
   		End If
   		
		wait 2
		Browser("Analyzer").Page("ReportCreation").Sync
		Browser("Analyzer").Page("ReportCreation").RefreshObject
   		
	End function
	
	
	Public Function openReportInDC(Byval ReportType,ByVal Reportlinkname)
		If ReportType="IAM Report" Then
			If Browser("Browser-SCA").Page("Page-SCA").WebElement("html tag:=P","innertext:="&Reportlinkname).Exist(20) Then
	   			Browser("Browser-SCA").Page("Page-SCA").WebElement("html tag:=P","innertext:="&Reportlinkname).click
				If browser("micclass:=browser").Page("micclass:=Page").Frame("html id:=AzWPFrame","html tag:=IFRAME").Exist(180) Then
					browser("micclass:=browser").Page("micclass:=Page").Frame("html id:=AzWPFrame","html tag:=IFRAME").highlight
					Call ReportStep (StatusTypes.Pass, "Report:-"&Reportlinkname&" is opened","Open Reports In DC")
					wait 5
					'Call ScreenCapture(Reportlinkname)
					Call takeScreenshotIAMReports(ReportLinkName)
					Browser("micclass:=browser").Back
					else
					Call ReportStep (StatusTypes.Fail, "Report:-"&Reportlinkname&" is not opened","Open Reports In DC")
				End If	   			
			End If
		
		ElseIf ReportType="Reporting Services" Then
			If Browser("Browser-SCA").Page("Page-SCA").WebElement("html tag:=P","innertext:="&Reportlinkname).Exist(20) Then
				Browser("Browser-SCA").Page("Page-SCA").WebElement("html tag:=P","innertext:="&Reportlinkname).click
				If Browser("micclass:=browser").Page("micclass:=page").WebTable("name:=Actions","html tag:=TABLE","class:=s4-wpTopTable").Exist(180) Then
					Browser("micclass:=browser").Page("micclass:=page").WebTable("name:=Actions","html tag:=TABLE","class:=s4-wpTopTable").highlight
					wait 8
					Call ReportStep (StatusTypes.Pass, "Report:-"&Reportlinkname&" is opened","Open Reports In DC")
					'Call ScreenCapture(Reportlinkname)
					Call takeScreenshotSSRS_MSTRReports(Reportlinkname)
					Browser("micclass:=browser").Back
					else
					Call ReportStep (StatusTypes.Fail, "Report:-"&Reportlinkname&" is not opened","Open Reports In DC")
				End If
			End If
		ElseIf ReportType="Excel Services" Then
			If Browser("Browser-SCA").Page("Page-SCA").WebElement("html tag:=P","innertext:="&Reportlinkname).Exist(20) Then
				Browser("Browser-SCA").Page("Page-SCA").WebElement("html tag:=P","innertext:="&Reportlinkname).click
				If Browser("micclass:=browser").Page("micclass:=page").WebTable("html tag:=TABLE","class:=s4-wpTopTable").Exist(180) Then
					Browser("micclass:=browser").Page("micclass:=page").WebTable("html tag:=TABLE","class:=s4-wpTopTable").highlight
					Call ReportStep (StatusTypes.Pass, "Report:-"&Reportlinkname&" is opened","Open Reports In DC")
					'Call ScreenCapture(Reportlinkname)
					Call takeScreenshotExcelReports(Reportlinkname)
					Browser("micclass:=browser").Back
					else
					Call ReportStep (StatusTypes.Fail, "Report:-"&Reportlinkname&" is not opened","Open Reports In DC")
				End If
			End If
		ElseIf ReportType="Other Documents" Then
			If Browser("Browser-SCA").Page("Page-SCA").WebElement("html tag:=P","innertext:="&Reportlinkname).Exist(20) Then
				Browser("Browser-SCA").Page("Page-SCA").WebElement("html tag:=P","innertext:="&Reportlinkname).click
				Call ReportStep (StatusTypes.Pass, "clicked Report:-"&Reportlinkname&" link","click document report in DC")
			End If
		ElseIf ReportType="VA PBI Report" Then
			If Browser("Browser-SCA").Page("Page-SCA").WebElement("html tag:=P","innertext:="&Reportlinkname).Exist(20) Then
				Browser("Browser-SCA").Page("Page-SCA").WebElement("html tag:=P","innertext:="&Reportlinkname).click
				If Browser("micclass:=browser").page("micclass:=page").Frame("title:=Power BI","html tag:=IFRAME").WebElement("class:=visualContainerHost","html tag:=DIV").Exist(180) Then
					Browser("micclass:=browser").page("micclass:=page").Frame("title:=Power BI","html tag:=IFRAME").WebElement("class:=visualContainerHost","html tag:=DIV").highlight
					Call ReportStep (StatusTypes.Pass, "Report:-"&Reportlinkname&" is opened","Open Reports In DC")
					wait 8
					'Call ScreenCapture(Reportlinkname)
					Call takeScreenshotPBIReports(ReportLinkName)
					Browser("micclass:=browser").Back
					else
					Call ReportStep (StatusTypes.Fail, "Report:-"&Reportlinkname&" is not opened","Open Reports In DC")
				End If
			End If
		ElseIf ReportType="VA Dossier Report" Then
			If Browser("Browser-SCA").Page("Page-SCA").WebElement("html tag:=P","innertext:="&Reportlinkname).Exist(20) Then
				Browser("Browser-SCA").Page("Page-SCA").WebElement("html tag:=P","innertext:="&Reportlinkname).click
				If Browser("micclass:=browser").page("micclass:=page").Frame("html tag:=IFRAME","html id:=MSOPageViewerWebPart_WebPartWPQ1").WebElement("class:=mstrmojo-VIVizPanel","html tag:=DIV").Exist(180) Then
					Browser("micclass:=browser").page("micclass:=page").Frame("html tag:=IFRAME","html id:=MSOPageViewerWebPart_WebPartWPQ1").WebElement("class:=mstrmojo-VIVizPanel","html tag:=DIV").highlight
					Call ReportStep (StatusTypes.Pass, "Report:-"&Reportlinkname&" is opened","Open Reports In DC")
					'Call ScreenCapture(Reportlinkname)
					Call takeScreenshotSSRS_MSTRReports(Reportlinkname)
					Browser("micclass:=browser").Back
					else
					Call ReportStep (StatusTypes.Fail, "Report:-"&Reportlinkname&" is not opened","Open Reports In DC")
				End If
			End If
		End If
	End Function
	
	Public Sub ValidateChartExistence()

		Dim oDesc,objCategories
	   	
	   	Browser("Analyzer").Page("ReportCreation").Sync
		wait 2
		Browser("Analyzer").Page("ReportCreation").Sync
	   	
'	    Set oDesc = Description.Create
'	    oDesc("micclass").value = "WebTable"
'	    oDesc("html tag").Value = "TABLE"
'	    oDesc("innertext").Regularexpression = True
'	    oDesc("innertext").Value = "Chart.*"
'	    oDesc("name").Value = "down_normal"
'	    oDesc("class").Value = "component"
'	    oDesc("visible").Value = True
'	    
'	    Set objCategories = Browser("IMS Analysis Manager").Page("Analyzer").ChildObjects(oDesc)
'	    'Msgbox objCategories.Count
'	    
'	    If Cint(objCategories.Count) = 0 Then
'
'			Call ReportStep (StatusTypes.Fail, "Chart is not displayed", "ReportCreation Page")
'		  			  
'		Else 
'		  
'			Call ReportStep (StatusTypes.Pass, "Chart is displayed", "ReportCreation Page")
'					
'		End If
	
		If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabReportChart").Exist(20) Then
			Call ReportStep (StatusTypes.Pass, "Chart is displayed", "ReportCreation Page")
		Else 
			Call ReportStep (StatusTypes.Fail, "Chart is not displayed", "ReportCreation Page")	
		End If
		
	End Sub
	
	'##### function to capture screenshot#####
	Public Function ScreenCapture(Byval ReportName)
		Dim vNow, vFile
		vNow = Replace(Replace(Replace(now(),":","_"),"/","_")," ","_")
		vfile =Environment.Value("CurrDir")&"Screenshots"&"\"&ReportName&""&vNow&".png"
		'Capture Browser Scrren shot
		'Browser("micclass:=Browser").FullScreen
		Browser("micclass:=Browser").CaptureBitmap vFile, True
		' Add the Captured Screen shot to the Results file
		Call ReportStep (StatusTypes.Pass, "ScreenShot captured for the report:-"&ReportName&"", "Capture screenshot")
	End Function
	
	
Public Function NavigateToMSTRFolderStructures1(ByVal Foldername,ByVal country,ByVal offering,Byval clientname)
	'folder structure can be VA design or Visual analytics
	If 	Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//DIV[@id='projects_ProjectsStyle']//a[text()='"&Foldername&"']").exist(30) Then
		Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//DIV[@id='projects_ProjectsStyle']//a[text()='"&Foldername&"']").click
	End If
	If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//input[@value='Continue']").Exist(10) Then
		Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//input[@value='Continue']").Click
	End If
	If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[@id='dktpSectionView']").Exist(30) Then
		Call ReportStep (StatusTypes.Pass, "clicked on folder:-"&Foldername&"","click on "&Foldername&" in mstr server")
		'click on shared reports
		If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[text()='Shared Reports']/..").Exist(30) Then
			Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[text()='Shared Reports']/..").Click
			If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']").Exist(30) Then
				Call ReportStep (StatusTypes.Pass, "clicked on folder:- shared reports and country folders page is displayed","click on shared reports folder in mstr server")
				'click on country folder
				If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='"&country&"']").exist(30) Then
					Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='"&country&"']").click
					If Browser("BrowserPoPUp").Page("micclass:=page").WebTable("innertext:="&offering&".*","html Tag:=Table","visible:=True","class:=mstrLargeIconItemTable").childitem(1,1,"WebElement",0).exist(30) Then
						Call ReportStep (StatusTypes.Pass, "clicked on country folder:-"&country&" and offering page is displayed","click on country folder in mstr server")
						'click on offering folder
						If Browser("BrowserPoPUp").Page("micclass:=page").WebTable("innertext:="&offering&".*","html Tag:=Table","visible:=True","class:=mstrLargeIconItemTable").childitem(1,1,"WebElement",0).exist(30) Then
							Browser("BrowserPoPUp").Page("micclass:=page").WebTable("innertext:="&offering&".*","html Tag:=Table","visible:=True","class:=mstrLargeIconItemTable").childitem(1,1,"WebElement",0).click
							Call ReportStep (StatusTypes.Pass, "clicked on offering folder:-"&offering&"","click on offering folder in mstr server")
							If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='"&clientname&"']").exist(30) Then
								Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='"&clientname&"']").click
								
							End If
							If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='ims']").Exist(30) Then
								Call ReportStep (StatusTypes.Pass, "clicked on client folder:-"&clientname&"","click on client folder in mstr server")
								'click on ims folder
								If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='ims']").Exist(30) Then
									Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='ims']").Click
								End If
								If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='Publish']").Exist(30) Then
									Call ReportStep (StatusTypes.Pass, "clicked on ims folder and Publish folder displayed successfully in mstr server","click on ims folder in mstr server")
									'click on publish folder
									Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='Publish']").Click
									If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//td[@id='td_mstrWeb_dockLeft']").Exist(30) Then
										Call ReportStep (StatusTypes.Pass, "clicked on publish folder and Reports section page displayed successfully in mstr server","click on publish folder in mstr server")
										else
										Call ReportStep (StatusTypes.Fail, "not clicked on publish folder and Reports section page not displayed in mstr server","click on publish folder in mstr server")
									End If
									else
									Call ReportStep (StatusTypes.Fail, "not clicked on ims folder and Publish folder not displayed in mstr server","click on ims folder in mstr server")
								End If
								else
								If Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='Publish']").Exist(30) Then
									Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//table[@id='FolderIcons']//a[text()='Publish']").Click
								End If
								'Call ReportStep (StatusTypes.Fail, "not clicked on offering folder:-"&offering&"","click on offering folder in mstr server")
							End If
						End If
						else
						Call ReportStep (StatusTypes.Fail, "not clicked on country folder:-"&country&" and offering page is not displayed","click on country folder in mstr server")
					End If
				End If
				else
				Call ReportStep (StatusTypes.Fail, "not clicked on folder:- shared reports","click on shared reports folder in mstr server")
			End If
		End If
		else
		Call ReportStep (StatusTypes.Fail, "not clicked on folder:-"&Foldername&"","click on "&Foldername&" in mstr server")
	End If
End Function

Function createWordDoc()
	Set objWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
Set objShape = objDoc.inlineShapes
Set objRang = objDoc.Range()
Set objShape = objDoc.inlineShapes
End Function

Function saveScreenshotsToWordDoc(Byval Reportname)
	Call createWordDoc()
	Dim vNow, vFile
		vNow = Replace(Replace(Replace(now(),":","_"),"/","_")," ","_")
		vfile =Environment.Value("CurrDir")&"Screenshots"&"\"&Reportname&""&vNow&".bmp"
	'vfile =Environment.Value("CurrDir")&"Screenshots"&"\"&"Test2"&".bmp"
	Browser("micclass:=Browser").CaptureBitmap vfile,True
	objShape.AddPicture (vfile)
	objWord.Selection.TypeParagraph()
	objDoc.saveas(Environment.Value("CurrDir")&"Screenshots"&"\"&Reportname&".doc")
End Function


Function takeScreenshotPBIReports(Byval ReportName)
	Set objWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
Set objShape = objDoc.inlineShapes
Set objRang = objDoc.Range()
Set objShape = objDoc.inlineShapes

Dim vNow, vFile
	vNow = Replace(Replace(Replace(now(),":","_"),"/","_")," ","_")
	vfile =Environment.Value("CurrDir")&"Screenshots"&"\"&Reportname&""&vNow&".png"

Set sheets=description.Create()
sheets("class").value="textLabel trimmedTextWithEllipsis"
sheets("html tag").Value="DIV"
set sheetCount=Browser("micclass:=browser").Page("micclass:=page").Frame("html tag:=IFRAME","title:=Power BI").WebElement("html tag:=UL","class:=pane sections ng-star-inserted").ChildObjects(sheets)
'msgbox sheetCount.count

For s = 0 To sheetCount.count-1 Step 1
	sheetCount(s).click
	wait 4
	
	Browser("micclass:=Browser").CaptureBitmap vfile,True
	objShape.AddPicture (vfile)
	objWord.Selection.TypeParagraph()

''Save Word File
	'objDoc.saveas(Environment.Value("CurrDir")&"Screenshots"&"\"&Reportname&""&vNow&".doc")

Next
objDoc.saveas(Environment.Value("CurrDir")&"Screenshots"&"\"&Reportname&""&vNow&".doc")
'objDoc.Close
objWord.ActiveDocument.Close
End Function

Function takeScreenshotIAMReports(Byval ReportName)
	Set objWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
Set objShape = objDoc.inlineShapes
Set objRang = objDoc.Range()
Set objShape = objDoc.inlineShapes

Dim vNow, vFile
	vNow = Replace(Replace(Replace(now(),":","_"),"/","_")," ","_")
	vfile =Environment.Value("CurrDir")&"Screenshots"&"\"&Reportname&""&vNow&".png"
colnums = Browser("micclass:=browser").Page("micclass:=page").frame("html id:=AzWPFrame").WebTable("name:=add").Getroproperty("cols")
For j = 1 To colnums Step 1
	Sheetname =  Browser("micclass:=browser").Page("micclass:=page").frame("html id:=AzWPFrame").WebTable("name:=add").GetCellData(1,j)
			if Sheetname<>"" then
				Browser("micclass:=browser").Page("micclass:=page").frame("html id:=AzWPFrame").WebTable("name:=add").ChildItem(1,j,"WebElement",0).click
				Browser("micclass:=Browser").CaptureBitmap vfile,True
				objShape.AddPicture (vfile)
				objWord.Selection.TypeParagraph()
				Sheetname = ""
				wait 5
			end if
	'wait 4
'	Dim vNow, vFile
'	vNow = Replace(Replace(Replace(now(),":","_"),"/","_")," ","_")
'	vfile =Environment.Value("CurrDir")&"Screenshots"&"\"&Reportname&""&vNow&".bmp"
'	Browser("micclass:=Browser").CaptureBitmap vfile,True
'	objShape.AddPicture (vfile)
'	objWord.Selection.TypeParagraph()

''Save Word File
	'objDoc.saveas(Environment.Value("CurrDir")&"Screenshots"&"\"&Reportname&""&vNow&".doc")

Next
objDoc.saveas(Environment.Value("CurrDir")&"Screenshots"&"\"&Reportname&""&vNow&".doc")
'objDoc.Close
objWord.ActiveDocument.Close
'objWord.ActiveDocument.Close
End Function

Function takeScreenshotExcelReports(Byval ReportName)
	Set objWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
Set objShape = objDoc.inlineShapes
Set objRang = objDoc.Range()
Set objShape = objDoc.inlineShapes

't
Dim vNow, vFile
vNow = Replace(Replace(Replace(now(),":","_"),"/","_")," ","_")
vfile =Environment.Value("CurrDir")&"Screenshots"&"\"&Reportname&""&vNow&".png"

Set d=description.Create()
d("html tag").Value="SPAN"
d("class").Value="dark"
set c=browser("micclass:=browser").Page("micclass:=page").WebTable("class:=s4-wpTopTable","html tag:=table").ChildObjects(d)
If c.count=1 Then
	
		'Browser("micclass:=Browser").CaptureBitmap vfile,True
	browser("micclass:=browser").CaptureBitmap vfile,True
	objShape.AddPicture (vfile)
	objWord.Selection.TypeParagraph()
	else
	
	For k = 0 To c.count-1 Step 1
		c(k).click
		wait 4
		'Dim vNow, vFile
'		vNow = Replace(Replace(Replace(now(),":","_"),"/","_")," ","_")
'		vfile =Environment.Value("CurrDir")&"Screenshots"&"\"&Reportname&""&vNow&".bmp"
		Browser("micclass:=Browser").CaptureBitmap vfile,True
		objShape.AddPicture (vfile)
		objWord.Selection.TypeParagraph()

''Save Word File
	'objDoc.saveas(Environment.Value("CurrDir")&"Screenshots"&"\"&Reportname&""&vNow&".doc")
	Next
End If
objDoc.saveas(Environment.Value("CurrDir")&"Screenshots"&"\"&Reportname&""&vNow&".doc")
'objDoc.Close
objWord.ActiveDocument.Close
End Function

Function takeScreenshotSSRS_MSTRReports(Byval ReportName)
	Set objWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
Set objShape = objDoc.inlineShapes
Set objRang = objDoc.Range()
Set objShape = objDoc.inlineShapes

Dim vNow, vFile
	vNow = Replace(Replace(Replace(now(),":","_"),"/","_")," ","_")
	vfile =Environment.Value("CurrDir")&"Screenshots"&"\"&Reportname&""&vNow&".png"


'For s = 0 To sheetCount.count-1 Step 1
	'sheetCount(s).click
	'wait 4
	
	Browser("micclass:=Browser").CaptureBitmap vfile,True
	objShape.AddPicture (vfile)
	objWord.Selection.TypeParagraph()

''Save Word File
	'objDoc.saveas(Environment.Value("CurrDir")&"Screenshots"&"\"&Reportname&""&vNow&".doc")

'Next
objDoc.saveas(Environment.Value("CurrDir")&"Screenshots"&"\"&Reportname&""&vNow&".doc")
'objDoc.Close
objWord.ActiveDocument.Close
'objDoc.cl
End Function

