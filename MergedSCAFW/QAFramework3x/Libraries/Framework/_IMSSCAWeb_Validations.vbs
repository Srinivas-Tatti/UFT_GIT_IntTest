'
'	Script:			_IMSSCA_Validations
'	Author:			Shilpi, Shobha, Shweta, Mani
'	Solution:		IMSSCA
'	Project:		IMSSCA
'	Date Created:	Thursday 23 April, 2015
'
'	Description:
'		<Enter Description>
'
'-------------------------------------------------------------------------------

Option Explicit

''' <summary>
''' 
''' </summary>
''' <returns type="IMSSCAValidationsLibrary"></returns>
''' <author>Shobha</author>
Public Function [New CIMSSCAValidationsLibrary]()
	
	Dim objValidations : Set objValidations = New CIMSSCAValidationsLibrary
	Set [New CIMSSCAValidationsLibrary]= objValidations

End Function


Class CIMSSCAValidationsLibrary




''' <summary>Validation of the ADDMAX Page and Captions of the each attribute</summary>
''''<Author>Shobha<Author>

Public Sub ADDMAXValidation_Page()
	
	If Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebElement("webDimension:").Exist(2) then
		   Call ReportStep (StatusTypes.Pass, "Dimension Element is present" , "Add\MDX Page")
        else
           Call ReportStep (StatusTypes.Fail, "Dimension Element is not present" , "Add\MDX Page")    

		End If
		
		
		If Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebElement("webAttribute:").Exist(2) then
		 Call ReportStep (StatusTypes.Pass, "Attribute Caption is present" , "Add\MDX Page")
        else
           Call ReportStep (StatusTypes.Fail, "Attribute Caption is not present" , "Add\MDX Page") 
		End If
		
		
		If Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebElement("webName").Exist(2) then
		 Call ReportStep (StatusTypes.Pass, "Name Caption is present" , "Add\MDX Page")
        else
         Call ReportStep (StatusTypes.Fail, "Name Caption is not present" , "Add\MDX Page") 
		End If
		
		
		If Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebElement("webCaption").Exist(2) then
		 Call ReportStep (StatusTypes.Pass, "Web caption is present" , "Add\MDX Page")
        else
         Call ReportStep (StatusTypes.Fail, "Web Caption is not present" , "Add\MDX Page") 
		End If
		
		
		If Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebElement("webFormat String").Exist(2) then
		 Call ReportStep (StatusTypes.Pass, "Format String  is present" , "Add\MDX Page")
        else
         Call ReportStep (StatusTypes.Fail, "Format String is not present" , "Add\MDX Page") 
		End If
		
		
		If Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebElement("webSolve Order").Exist(2) then
		 Call ReportStep (StatusTypes.Pass, "Solve Order is present" , "Add\MDX Page")
        else
           Call ReportStep (StatusTypes.Fail, "Solve Order is not present" , "Add\MDX Page") 
		End If
		
		
		If Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebElement("webVisible").Exist(2) then
		 Call ReportStep (StatusTypes.Pass, "webVisible is present" , "Add\MDX Page")
        else
           Call ReportStep (StatusTypes.Fail, "webVisible is not present" , "Add\MDX Page") 
		End If
		
		
		If Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebElement("webExpandable").Exist(3) then
		 Call ReportStep (StatusTypes.Pass, "Expandable Element is present" , "Add\MDX Page")
        else
           Call ReportStep (StatusTypes.Fail, "Expandable Element is not present" , "Add\MDX Page") 
		End If
		
		
		If Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebElement("webCalculable").Exist(2) then
		 Call ReportStep (StatusTypes.Pass, "Calculable Element is present" , "Add\MDX Page")
        else
         Call ReportStep (StatusTypes.Fail, "CalculableElement is not present" , "Add\MDX Page") 
		End If
		
		
		If Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebElement("webDisplay As").Exist(3) then
		 Call ReportStep (StatusTypes.Pass, "Display As Element is present" , "Add\MDX Page")
        else
         Call ReportStep (StatusTypes.Fail, "Display As Element is not present" , "Add\MDX Page") 
		End If
		
		If Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebElement("webGroup").Exist(2) then
		  Call ReportStep (StatusTypes.Pass, "Group Element is present" , "Add\MDX Page")
        else
          Call ReportStep (StatusTypes.Fail, "Group Element is not present" , "Add\MDX Page")
		End If	
		
		If Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebElement("webIndividual Members").Exist(3) then
		  Call ReportStep (StatusTypes.Pass, "Individual Members is present" , "Add\MDX Page")
        else
          Call ReportStep (StatusTypes.Fail, "Individual Members is not present" , "Add\MDX Page")
		End If
		
		If Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebElement("webBehavior").Exist(2) then
		  Call ReportStep (StatusTypes.Pass, "Behavior Element is present" , "Add\MDX Page")
        else
          Call ReportStep (StatusTypes.Fail, "Behavior Element is not present" , "Add\MDX Page")
		End If
		
		If Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebElement("webStatic").Exist(2) then
		   Call ReportStep (StatusTypes.Pass, "Static Element is present" , "Add\MDX Page")
        else
          Call ReportStep (StatusTypes.Fail, "Static Element is not present" , "Add\MDX Page")
		End If
		
		If Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebElement("webDynamic").Exist(2) then
			Call ReportStep (StatusTypes.Pass, "Dynamic Element is present" , "Add\MDX Page")
        else
          Call ReportStep (StatusTypes.Fail, "Dynamic Element is not present" , "Add\MDX Page")
		End If
		
		
		If Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebElement("webContent Type").Exist(3) then
			Call ReportStep (StatusTypes.Pass, "Content Type Element is present" , "Add\MDX Page")
        else
          Call ReportStep (StatusTypes.Fail, "Content Type Element is not present" , "Add\MDX Page")
		End If
		
		If Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebElement("weblistMember").Exist(2) then
			Call ReportStep (StatusTypes.Pass, "listMember Element is present" , "Add\MDX Page")
        else
           Call ReportStep (StatusTypes.Fail, "listMember Element is not present" , "Add\MDX Page")
		End If
		
		If Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebElement("webMDX").Exist(2) then
			Call ReportStep (StatusTypes.Pass, "MDX Element is present" , "Add\MDX Page")
        else
           Call ReportStep (StatusTypes.Fail, "MDX Element is not present" , "Add\MDX Page")
		End If
		
		
		If Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebElement("webDelimiter").Exist(2) then
		  Call ReportStep (StatusTypes.Pass, "Delimiter Element is present" , "Add\MDX Page")
        else
          Call ReportStep (StatusTypes.Fail, "Delimiter Element is not present" , "Add\MDX Page")
		End If
		
		If Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebElement("webomma").Exist(3) then
			Call ReportStep (StatusTypes.Pass, "Comma Element is present" , "Add\MDX Page")
        else
          Call ReportStep (StatusTypes.Fail, "Comma Element is not present" , "Add\MDX Page")
		End If

		If Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebElement("webLine").Exist(4) then
		  Call ReportStep (StatusTypes.Pass, "Line Element is present" , "Add\MDX Page")
        else
          Call ReportStep (StatusTypes.Fail, "Line Element is not present" , "Add\MDX Page")
		End If
		If Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebElement("webSpace").Exist(3) then
		  Call ReportStep (StatusTypes.Pass, "Space Element is present" , "Add\MDX Page")
        else
          Call ReportStep (StatusTypes.Fail, "Space Element is not present" , "Add\MDX Page")
		End If

		If Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebElement("webSemicolon").Exist(2) Then
		
		Call ReportStep (StatusTypes.Pass, "Semicolon Element is present" , "Add\MDX Page")
        else
         Call ReportStep (StatusTypes.Fail, "Semicolon Element is not present" , "Add\MDX Page")
		End If
	
End Sub




''' <summary>Validate of the Edit Group page</summary>
''''<Author>Shobha<Author>
Public Sub ValidateEditGroupPage( )

        'Strings
        
        'Objects
        
        If  Browser("Analyzer").Page("GroupCreationPage").Frame("EditGroupDialog").Image("imgAddFilter").exist(10) then
                Call ReportStep (StatusTypes.Pass, "Add Filter Icon Found" , "Edit Filter Page")
        else
                Call ReportStep (StatusTypes.Fail, "Add Filter Icon not Found" , "Edit Filter Page")    
        End If
        
        If  Browser("Analyzer").Page("GroupCreationPage").Frame("EditGroupDialog").Image("imgAddRange").exist(10) then
                Call ReportStep (StatusTypes.Pass, "Add Range Icon Found" , "Edit Filter Page")
        else
                Call ReportStep (StatusTypes.Fail, "Add Range Icon not Found" , "Edit Filter Page")    
        End If
        
        If  Browser("Analyzer").Page("GroupCreationPage").Frame("EditGroupDialog").Image("imgAddCalculation").exist(10) then
                Call ReportStep (StatusTypes.Pass, "Add Calculation Icon Found" , "Edit Filter Page")
        else
                Call ReportStep (StatusTypes.Fail, "Add Calculation Icon not Found" , "Edit Filter Page")    
        End If
        
        If  Browser("Analyzer").Page("GroupCreationPage").Frame("EditGroupDialog").Image("imgAddAllMembers").exist(10) then
                Call ReportStep (StatusTypes.Pass, "Add All Members Icon Image Found" , "Edit Filter Page")
        else
                Call ReportStep (StatusTypes.Fail, "Add All Members Icon not Found" , "Edit Filter Page")    
        End If
        
        If  Browser("Analyzer").Page("GroupCreationPage").Frame("EditGroupDialog").Image("imgAddAlOtherMembers").exist(10) then
                Call ReportStep (StatusTypes.Pass, "Add All Other Members Icon Found" , "Edit Filter Page")
        else
                Call ReportStep (StatusTypes.Fail, "Add All Other Members Icon not Found" , "Edit Filter Page")    
        End If        
        
        If  Browser("Analyzer").Page("GroupCreationPage").Frame("EditGroupDialog").Image("imgDuplicateGroup").exist(10) then
                Call ReportStep (StatusTypes.Pass, "Duplicate Icon  Found" , "Edit Filter Page")
        else
                Call ReportStep (StatusTypes.Fail, "Duplicate Icon not Found" , "Edit Filter Page")    
        End If    
        
        If  Browser("Analyzer").Page("GroupCreationPage").Frame("EditGroupDialog").Image("imgDeleteGroup").exist(10) then
                Call ReportStep (StatusTypes.Pass, "Delete Icon Found" , "Edit Filter Page")
        else
                Call ReportStep (StatusTypes.Fail, "Delete Icon not Found" , "Edit Filter Page")    
        End If    
    End Sub    


''' <summary>Validation of the Group display in the Report create Page</summary>
''' <param strGroupCap_Name>GroupCaption Name</param>	
''' <param strGroupMembers>Members of the Group</param>
''''<param strF>Value to handle in next release </param>
''''<Author>Shobha<Author>
Public Function Groupdisplay_V(ByVal strGroupCap_Name,ByVal strGroupMembers,ByVal strF)
	
	Dim objArrow,objGrouptab,b,i,j,strtabvalue,strGrouptabvalues,objArrow_P,strCapreturn,returnGroupMembers,retuenVal
	Dim m,n,strtabvalue_Expanding,strGrouptabvalues_E
	
	strCapreturn=1
	objArrow_P=1
	returnGroupMembers=1
	
	
	Set objArrow=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgExpandMembers")
				
	Set objGrouptab=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGroupArea")
	b=SCA.Webtable(objGrouptab,"Grouptab Report Creation","Retriving_DataTableValue","Report Creation Page","","","","")
	
	For i = 0 To ubound(b,1)-1 Step 1
		For j = 0 To ubound(b,2)-1 Step 1			
			strtabvalue=b(i,j)
			strGrouptabvalues=strGrouptabvalues&strtabvalue			
		Next		
	Next
	
	If instr(strGrouptabvalues,strGroupCap_Name)<>0 Then
		
		strCapreturn=0
		else
		strCapreturn=1
		
	End If
	
	If objArrow.Exist(3) Then
		Set objArrow=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgExpandMembers")
		Call SCA.ClickOn(objArrow,"GroupArrow", "GroupCreation Report Area")
		objArrow_P=0
		else
		objArrow_P=1		
	End If
	
	wait 2
	
	Set objGrouptab=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGroupArea")
	b=SCA.Webtable(objGrouptab,"Grouptab Report Creation","Retriving_DataTableValue","Report Creation Page","","","","")
	
	For m = 0 To ubound(b,1)-1 Step 1
		For n = 0 To ubound(b,2)-1 Step 1			
			strtabvalue_Expanding=b(m,n)
			strGrouptabvalues_E=strGrouptabvalues_E&strtabvalue_Expanding			
		Next		
	Next
		
	If instr(TRIM(strGrouptabvalues_E),TRIM(strGroupMembers))<>0 Then		
		returnGroupMembers=0
		else
		returnGroupMembers=1		
	End If
	
   If returnGroupMembers=0 AND objArrow_P=0 AND strCapreturn=0 Then
   
	retuenVal=0
	
   ElseIf returnGroupMembers=0 AND objArrow_P=1 AND strCapreturn=1 Then

	retuenVal=1
	ElseIf returnGroupMembers=1 AND objArrow_P=1 AND strCapreturn=0 Then
	' To validate the expandable members
	retuenVal=2
	
End If

Groupdisplay_V=retuenVal
	
End Function


''' <summary>Delimiter Validation in MDX Group creation</summary>
''' <param strSelection_O>Delimiter selection</param>	
''' <param StrMemberValues>Members of the delimiter</param>
''''<Author>Shobha<Author>
Public Function delimiterValidation_MDXGroup(ByVal strSelection_O,ByVal StrMemberValues,ByVal Delimiter_Lselection)

 Dim arr,i,strdelimit,strMembertext,strMemberCount,MemberVal,arrline,strlineMembers
 Dim objrdoDelimiter,objtxtMember,objevalBtn
 
 Set objrdoDelimiter=Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebRadioGroup("rdoDelimiter")
 Call SCA.ClickOnRadio(objrdoDelimiter,strSelection_O,"GroupCreation Dialouge")

 If Delimiter_Lselection=0 Then
 		
    arrline=split(StrMemberValues," ")
 	For i = 0 To ubound(arrline) Step 1
 		strlineMembers=strlineMembers&arrline(i)&vbnewline 		
 	Next
	'StrMemberValues=strlineMembers 	
	Set objtxtMember=Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebEdit("txtMembers")
   Call SCA.SetText_MultipleLineArea(objtxtMember, strlineMembers, "txtMember", "GroupCreation Dialouge")	

else
	
   Set objtxtMember=Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebEdit("txtMembers")
   Call SCA.SetText_MultipleLineArea(objtxtMember, StrMemberValues, "txtMember", "GroupCreation Dialouge")	

 End If
 
 
 Set objevalBtn=Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebButton("btnEvaluate")
 Call SCA.ClickOn(objevalBtn,"btnEvaluate", "GroupCreation Dialouge")
		 

 
 Select Case strSelection_O
 	
 	Case "rdoComma"
 	
 	arr=split(StrMemberValues,",")
 	For i = 0 To ubound(arr)
 		strdelimit=strdelimit&" "&arr(i) 		
 	Next
 	
 	
 	Case "rdoSpace"
 	
 	arr=split(StrMemberValues," ")
 	For i = 0 To ubound(arr)
 		strdelimit=strdelimit&" "&arr(i) 		
 	Next
 	
 	Case "rdoLine"
 	
 	arr=split(StrMemberValues," ")
 	For i = 0 To ubound(arr)
 		strdelimit=strdelimit&" "&arr(i) 		
 	Next
 	
 	Case "rdoSemicolon"
 	
 	arr=split(StrMemberValues,";")
 	
 	For i = 0 To ubound(arr)
 
 	strdelimit=strdelimit&" "&arr(i) 
 	
   Next
 	
 End Select
 
 
 
 strMemberCount=trim(Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebElement("webMembercount").GetROProperty("innertext"))
 strMembertext=trim(Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebList("lstMembers").GetROProperty("innertext"))
 
 
 
 If instr(strMemberCount,ubound(arr)+1)<>0 AND Strcomp(trim(strMembertext),trim(strdelimit))=0 Then
 	
 	MemberVal=0
 	else
 	MemberVal=1   
		
 End If
 
delimiterValidation_MDXGroup=MemberVal	
	
End Function
	

''' <summary> Validation of the Group Creation in the Report creation page</summary>
''' <param strGroupname>GroupName</param>	
''' <param strGroupCaption>GroupCaption Name for the validation</param>
''''<Author>Shobha<Author>	
Public Function Groupcreation_Verification(strGroupname,strGroupCaption,strF)

Dim objgroup,objg,Gpresence,objbtnapply,rc,cc,i,j,strcellval,Gcap_presence,G_creation
Gpresence=1
Gcap_presence=1
G_creation=1

Set objgroup=description.Create
objgroup("micclass").value="WebElement"
objgroup("innertext").value=strGroupname
objgroup("class").value="trvItemsNode TreeNode"

Set objg=Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebTable("tabGroupName").ChildObjects(objgroup)
if objg.Count=1 then
Gpresence=0
End if

set objbtnapply=Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebButton("btnApply")
Call SCA.ClickOn(objbtnapply,"objbtnapply", "GroupCreation Dialouge")

rc=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGroupArea").RowCount
cc=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGroupArea").ColumnCount(rc)

For i= 1 To rc Step 1
For j = 1 To cc Step 1
	
	strcellval=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGroupArea").GetCellData(i,j) 
	If instr(TRIM(strcellval),TRIM(strGroupCaption))<>0 Then
		Gcap_presence=0
		i=rc
		
		Exit For
		
	End If
Next

Next

If Gpresence=0 AND Gcap_presence=0 Then
	
	G_creation=0
	ElseIf Gpresence=0 AND Gcap_presence=1 Then
	G_creation=1
	
End If 

Groupcreation_Verification=G_creation


End Function


''' <summary>Verification of the Bookmark in the BookList</summary>
''' <param name="strBookName" type="By Value">Field  Value of BookMark</param>	
''' <param name="strPageName" type="By Value">Screen location</param>
Public Sub VerificationReportValues(ByVal rowcount,ByVal ColumnCount,ByVal strPageName)

	If rowcount>2 Then
	 Call ReportStep (StatusTypes.Pass,"Creating A Report using Pivot Table:-"&Space(5)&"Report is created using Pivot Table", strPageName)
	 Else
	 Call ReportStep (StatusTypes.Fail,"Creating A Report using Pivot Table:-"&Space(5)&"Report is not created using Pivot Table", strPageName)	
	End If
	
End Sub


''' <summary>Verification of the Bookmark in the BookList</summary>
''' <param name="strBookName" type="By Value">Field  Value of BookMark</param>	
''' <param name="strPageName" type="By Value">Screen location</param>
Public Sub ValidateCreationof_Weblink(ByVal strLink,ByVal strPageName)

	Dim strActLinkname
	
	strActLinkname=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webAutomation").GetROProperty("innertext")
	If strcomp(TRIM(strLink),TRIM(strActLinkname))=0 Then
	
		Call ReportStep (StatusTypes.Pass, "Adding a Web Link:-"&"The Web Link"&Space(2)&strLink&"is displayed on the same page as the report.", strPageName)
			Else
		Call ReportStep (StatusTypes.Fail, "Adding a Web Link:-"&"The Web Link"&Space(2)&strLink&"isnot displayed on the same page as the report.", strPageName)
	
	End If



End Sub


''' <summary>Verification of the Bookmark in the BookList</summary>
''' <param name="strBookName" type="By Value">Field  Value of BookMark</param>	
''' <param name="strPageName" type="By Value">Screen location</param>
Public Sub Bookmark_Validation(ByVal strBookName,ByVal strPageName)

	Dim strActBookName	  
	strActBookName=Browser("Analyzer").Page("ReportCreation").Frame("FrameBookmark").WebTable("tblbookmark").GetROProperty("innertext")
		
	If instr(strActBookName,strBookName)<>0 Then	
	Call ReportStep (StatusTypes.Pass, +strBookName+"Bookmark is created for the Report ", strPageName)
	Else
	Call ReportStep (StatusTypes.Fail, +strBookName+"Bookmark is not created for the Report ", strPageName)		
	End If

End Sub

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'''' <summary>Validation of Report Creation in SCA</summary>
'''' <param name="strReportName" type="By Value">Field  Value of ReportName</param>	
'''' <param name="strFolderName" type="By Value">Folder Location of the Report</param>
'''' <param name="strPageName" type="By Value">Screen Location</param>
Public Function ReportCreation_Validation(ByVal strReportName,ByVal strFolderName,ByVal strPageName)

	Dim strFElements,objlink,objlinks,strlinkName,i,strreturnval
	
	Set objlink=description.Create
	objlink("micclass").Value="Link"
	objlink("name").value=strReportName
	objlink("html tag").value="A"
	Set objlinks=Browser("Analyzer").Page("Shared_Folder").Frame("view").ChildObjects(objlink)
	
	For i=0 to objlinks.count-1
	
		strlinkName=objlinks(i).GetROProperty("innertext")
		If strcomp(strlinkName,strReportName)=0 Then
		strreturnval=0	
		Else
		strreturnval=1		
		 Exit For 
		 
		End If
	
	Next

ReportCreation_Validation=strreturnval

End Function
''&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&


''' <summary>Verify thereport properties is set correctly</summary>
''' <param name="strReportName" type="By Value">Report Name</param>
''' <param name="strReportDescription" type="By Value">Report Name</param>
''' <param name="strReportVideoOrDoc" type="By Value">Report Video or Document</param>
''' <param name="strChkAutoRefreshReport" type="By Value">Report AutoRefresh Parameter</param>
''' <param name="strAutoRefreshInterval" type="By Value">Report AutoRefresh Interval</param>
''' <param name="strPageName" type="By Value">Report Page Name</param>
Public Sub ReportPropertiesValidation(ByVal strReportName,ByVal strReportDescription,ByVal strReportVideoOrDoc, ByVal strChkAutoRefreshReport, Byval strAutoRefreshInterval, ByVal strPageName)

	Dim objFrame, FrameName, objReportNameTxt, objReportDescription, objReportVideoOrDoc, objchkAutoRefreshReport, objAutoRefreshInterval
	
	Set objFrame = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer")
	FrameName = objFrame.GetROProperty("name")
	Call IMSSCA.General.ReportToolBar("edit.svg", FrameName)
	
	wait 2
	
	Set objReportNameTxt=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebEdit("txtReportSettingsName")
	Set objReportDescription = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebEdit("txtReportPropertyDescription")
	Set objReportVideoOrDoc = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebEdit("txtVideoOrDoc")
'		Set objReportAccessUrl = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebEdit("txtReportAccessUrl")
	Set objchkAutoRefreshReport = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebCheckBox("chkAutoRefreshReport")
	Set objAutoRefreshInterval = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebEdit("txtAutoRefreshInterval")
					
	'Validate Report Properties field values
	'Validate strReportName field
	Call ValidateFieldText(objReportNameTxt, "txtReportSettingsName", strReportName, strPageName)
	
	'Validate strReportDescription field
	Call ValidateFieldText(objReportDescription, "txtReportPropertyDescription", strReportDescription, strPageName)
	
	'Validate strReportDescription field
	Call ValidateFieldText(objReportVideoOrDoc, "txtVideoOrDoc", strReportVideoOrDoc, strPageName)
			
	'Validate strReportDescription field
	Call VerifyCheckBox(objchkAutoRefreshReport, "chkAutoRefreshReport", strChkAutoRefreshReport, strPageName)
	
	'Validate strReportDescription field
	Call ValidateFieldText(objAutoRefreshInterval, "txtAutoRefreshInterval", strAutoRefreshInterval, strPageName)


End Sub


''' <summary>Verify the page is loaded</summary>
''' <param name="strPageName" type="By Value">Screen location</param>
Public Sub ValidatePageLoaded(ByVal strPageName)

	Dim oDesc,objVisible
	
	Set oDesc = Description.Create()
	oDesc("micclass").Value = "WebElement"
	oDesc("class").Value = "c"
	oDesc("html tag").Value = "TD"
	oDesc("innertext").Value = strPageName
	
	If Browser("micclass:=Browser").Page("micclass:=Page").WebElement(oDesc).Exist(6) Then
		Call ReportStep (StatusTypes.Pass, "Validated "+strPageName+" Page", strPageName)
	Else
		Call ReportStep (StatusTypes.Fail, "Could not validated "+strPageName+" Page", strPageName)
	
	End If

End Sub	
	
	
''' <summary>Verify the value in the fields (i.e WebEdit and WebList)</summary>
''' <param name="objField" type="Object">Field object</param>
''' <param name="strFieldName" type="By Value">Field name of the value to get</param>
''' <param name="strCompareValue" type="Scripting.Dictionary">Data value to compare</param>
''' <param name="strPageName" type="By Value">Screen location</param>
Public Sub ValidateFieldText(ByRef objField , ByVal strFieldName , ByVal strCompareValue , ByVal strPageName)

	If Not (objField.Exist(2)) Then
		Call ReportStep (StatusTypes.Fail, "Could not validate the '"&strFieldName&"' :field as it does not exist", strPageName)
		Exit Sub
	End If
	
	If (Not(objField.GetROProperty("micclass")="WebEdit")) And (Not(objField.GetROProperty("micclass")="WebList")) Then
		Call ReportStep (StatusTypes.Warning, "Object passed to the function is not a WebEdit or WebList. Please check the object for the WebEdit or WebList : '"&strFieldName&"'", strPageName)
		Exit Sub
	End If
	
	If Trim(objField.GetROProperty("default value"))=Trim(strCompareValue) Or Trim(objField.GetROProperty("innertext"))=Trim(strCompareValue) Or Trim(objField.GetROProperty("value"))=Trim(strCompareValue) Then
		Call ReportStep (StatusTypes.Pass, "Supplied text '"&strCompareValue&"' is displayed in "&strFieldName, strPageName)
	Else
		Call ReportStep (StatusTypes.Fail, "Supplied text '"&strCompareValue&"' is not displayed in "&strFieldName, strPageName)
	End If

End Sub


''' <summary>Verify the date in the field</summary>
''' <param name="objField" type="Object">Field object</param>
''' <param name="strFieldName" type="By Value">Field name of the value to get</param>
''' <param name="strCompareValue" type="Scripting.Dictionary">Data value to compare</param>
''' <param name="strPageName" type="By Value">Screen location</param>
Public Sub ValidateFieldDate(ByRef objField , ByVal strFieldName , ByVal strCompareValue , ByVal strPageName)
	
	If Not (objField.Exist(2)) Then
		Call ReportStep (StatusTypes.Fail, "Could not validate the date field as it does not exist: '"&strFieldName&"'", strPageName)
		Exit Sub
	End If
	
	If (Not(objField.GetROProperty("micclass")="WebEdit")) Then
		Call ReportStep (StatusTypes.Warning, "Object passed to the function is not a WebEdit (Date field). Please check the object for the WebEdit (Date field) : '"&strFieldName&"'", strPageName)
		Exit Sub
	End If
	
	If Instr(CDate(Trim(objField.GetROProperty("value"))),CDate(Trim(strCompareValue)))<>0 Then
		Call ReportStep (StatusTypes.Pass, "Given date: '"&objField.GetROProperty("value")&"' is in "&strFieldName, strPageName)
	Else
		Call ReportStep (StatusTypes.Fail, "Given date: '"&objField.GetROProperty("value")&"' is not in "&strFieldName, strPageName)
	End If
	
End Sub
	
	
	
''' <summary>Verify the selected radio option from the radio group</summary>
''' <param name="objRadioGroup" type="Object">Radio group object</param>
''' <param name="strCompareValue" type="Scripting.Dictionary">Option value To compare</param>
''' <param name="strPageName" type="By Value">Screen location</param>
Public Sub ValidateRadioSelection(ByRef objRadioGroup , ByVal strCompareValue , ByVal strPageName)
	
	If (Not(objRadioGroup.GetROProperty("micclass")="WebRadioGroup")) Then
		Call ReportStep (StatusTypes.Warning, "Object passed to the function is not a WebRadioGroup. Please check the object.", strPageName)
		Exit Sub
	End If
	
	If Trim(objRadioGroup.GetRoProperty("value"))=Trim(strCompareValue) Then
		Call ReportStep(StatusTypes.Pass, "Selected radio value is - "&strCompareValue, strPageName)
	Else
		Call ReportStep(StatusTypes.Pass, "Compared value does match", strPageName)
	End If	
	
End Sub
	
	
	
''' <summary>Verify the value in the field contains the given value</summary>
''' <param name="objField" type="Object">Field object</param>
''' <param name="strFieldName" type="By Value">Field name of the value to get</param>
''' <param name="strCompareValue" type="Scripting.Dictionary">Data value to compare</param>
''' <param name="strPageName" type="By Value">Screen location</param>
Public Sub ValidateFieldContainsText(ByRef objField , ByVal strFieldName , ByVal strCompareValue , ByVal strPageName)

	If objField.Exist(2) Then
		If Instr(Trim(objField.GetROProperty("default value")),Trim(strCompareValue))<>0 Or Instr(Trim(objField.GetROProperty("innertext")),Trim(strCompareValue))<>0 Then
			Call ReportStep (StatusTypes.Pass, strFieldName&" field contains '"&strCompareValue&"'.", strPageName)
		Else
			Call ReportStep (StatusTypes.Fail, strFieldName&" field does not contain '"&strCompareValue&"'.", strPageName)
		End If
	Else
		Call ReportStep (StatusTypes.Fail, "Could not validate the '"&strFieldName&"' :field as it does not exist", strPageName)
	End If
	
End Sub
	
	
''' <summary>Verifying the categories displayed on the left side of OCRF Web page</summary>
Public Sub VerifyCategoryList()

    Dim oDesc,objCategories
    Dim strCategory
    Dim i
    
    Set oDesc = Description.Create
    oDesc("micclass").value = "WebElement"
    oDesc( "html tag" ).Value = "LI"
    oDesc( "innerHtml" ).Value = ".*indicator fa-li fa fa.*"
    oDesc( "visible" ).Value = True
    
    Set objCategories = Browser("micclass:=Browser").Page("micclass:=Page").ChildObjects(oDesc)
    'Msgbox objCategories.Count
	If Cint(objCategories.Count) = 0 Then
	  Call ReportStep (StatusTypes.Fail, "Categories are not displayed", "Categories Page")
	  Exit sub
	End If
       
    For i = 0 to Cint(objCategories.Count) - 1        
        strCategory = Trim(objCategories(i).GetROProperty("innertext"))
        Call ReportStep (StatusTypes.Pass, "Category: '"&strCategory&"' is displayed on left side of the all apps table.", "Categories Page")
        'Msgbox strCategory
    Next
    
End Sub


	
	
''' <summary>To verify Checkbox Object is checked or unchecked</summary>
''' <param name="webCheckBoxObject" type="QTPWebCheckBoxObject">WebCheckBox object to check</param>
''' <param name="strCompareValue" type="String">Value to compare with the value of  WebCheckBox as "ON" to check the checkbox and "OFF" to uncheck it.</param>	
Public Sub VerifyCheckBox(ByVal webCheckBoxObject,ByVal strCheckboxName, ByVal strCompareValue, ByVal strPageName)

	If Not webCheckBoxObject.Exist(2) Then
		Call ReportStep (StatusTypes.Fail, "Field '"&strCheckboxName&"' does not exist", strPageName)		
		Exit Sub
	End If
	If UCase(strCompareValue) = "ON" Then
		If webCheckBoxObject.GetROProperty("checked") = 1 Then
			Call ReportStep (StatusTypes.Pass,"Check box '"&strCheckboxName&"'is checked", strPageName)
		Else
			Call ReportStep (StatusTypes.Fail,"Check box '"&strCheckboxName&"'is not checked", strPageName)
		End If
	ElseIf UCase(strCompareValue) = "OFF" Then
		If webCheckBoxObject.GetROProperty("checked") = 0 Then
			Call ReportStep (StatusTypes.Pass,"Check box "&strCheckboxName&"'is unchecked", strPageName)
		Else
			Call ReportStep (StatusTypes.Fail,"Check box '"&strCheckboxName&"'is not unchecked", strPageName)
		End If
	End If
	
End Sub


	''' <summary>Verify the existence of Chart with params/summary>
	''' <param name="strChartPos" type="String">Chart Pos with respect to Another component(Say Pivot Table). It could be TOP, BOTTOM, LEFT, RIGHT</param>
	''' <param name="strSheetComp" type="String">to validate graph’s position with resoect to Component selected in the sheet</param>
	''' <param name="strChartType" type="String">Chart Type. It could be different types of basic/advanced graphs</param>
	''' <param name="strChartSheet" type="String">Type of sheet where Graph has created</param>
	''' <param name="strPageName" type="String">Report Page Name</param>
	''' <param name="objData" type="String">Reference to objData</param>
	Public Sub ValidateChartExistenceParams(ByVal strChartPos, ByVal strSheetComp, ByVal strChartType, ByVal strChartSheet, ByVal strPageName, ByRef objData)

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
	   		Exit Sub
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
	   		Exit Sub
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
					
	End Sub
	
	
	''' <summary>Verifies the Position of Chart with params/summary>
	''' <param name="strChartPos" type="String">Chart Pos with respect to Another component(Say Pivot Table). It could be TOP, BOTTOM, LEFT, RIGHT</param>
	''' <param name="strSheetComp" type="String">to validate graph’s position with resoect to Component selected in the sheet</param>
	''' <param name="strPageName" type="String">Report Page Name</param>
	''' <param name="objData" type="String">Reference to objData</param>
	Public Function ValidateReportPosition(ByVal strChartPos, ByVal strSheetComp, ByVal strPageName, ByRef objData)
		
		Dim objDesignerSplitter, xobjDesignerSplitter, yobjDesignerSplitter, objVerticalSplitter, xobjVerticalSplitter, yobjVerticalSplitter, objHorizontalSplitter, xobjHorizontalSplitter, yobjHorizontalSplitter, xChart, yChart, oDescPivot, oDescPivotObj, xPivotTab, yPivotTab, strChartPosFound, diff
		ValidateReportPosition = 0
		
		'Splitter C and Y co ordinate values - Start
				set objDesignerSplitter = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welDesignerSplitter")
			  	If objDesignerSplitter.Exist(5) Then
			  		xobjDesignerSplitter= objDesignerSplitter.GetROProperty("x")
			  		yobjDesignerSplitter= objDesignerSplitter.GetROProperty("y")
			  	else
			  		Call ReportStep (StatusTypes.Information, "Desginer Splitter welDesignerSplitter doesnt exist" , strPageName)
			  	End If
			    
			    set objVerticalSplitter = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welVerticalSplitter")
			  	If objVerticalSplitter.Exist(5) Then
			  		xobjVerticalSplitter= objVerticalSplitter.GetROProperty("x")
			  		yobjVerticalSplitter= objVerticalSplitter.GetROProperty("y")
			  	else
			  		Call ReportStep (StatusTypes.Information, "Vertical Splitter welVerticalSplitter doesnt exist" , strPageName)
			  	End If
			  	
			  	set objHorizontalSplitter = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welHorizontalSplitter")
			  	If objHorizontalSplitter.Exist(5) Then
			  		xobjHorizontalSplitter= objHorizontalSplitter.GetROProperty("x")
			  		yobjHorizontalSplitter= objHorizontalSplitter.GetROProperty("y")
			  	else
			  		Call ReportStep (StatusTypes.Information, "Horizontal Splitter welHorizontalSplitter doesnt exist" , strPageName)
			  	End If
		'Splitter C and Y co ordinate values - End
		
		'Get x and Y co-ordinates of Chart
				xChart = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabReportChart").GetROProperty("x")
				yChart = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabReportChart").GetROProperty("y")
		
		'Chart Position is not required to validate, if sheet created doesnt contain other components
				If strChartPos = "" Then
					Call ReportStep (StatusTypes.Information, "Chart is created in new sheet. This Sheet doesnt contain other components", strPageName)
					ValidateReportPosition = 2
					Exit Function
				End if

		'Chart Position is required to validate, if sheet created contains other components
		Select Case strSheetComp
			Case "Pivot Table"
						'Descriptive Programming for Pivot table
						Set oDescPivot = Description.Create()
						oDescPivot("micclass").value = "WebTable"
						oDescPivot("class").value = "component"
						oDescPivot("column names").regularexpression=True
						oDescPivot("column names").value = "PivotTable.*"
						oDescPivot("html tag").value = "TABLE"
						oDescPivot("name").value = "down_normal"
						'oDescPivot("outertext").regularexpression=True
						'oDescPivot("outertext").value = ".*PivotTable.*Drop a Filter Condition Here.*"
									    		
						Set oDescPivotObj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(oDescPivot)
						
						'Get x and Y co-ordinates of Pivot Table
						xPivotTab = oDescPivotObj(0).GetROProperty("x")
						yPivotTab = oDescPivotObj(0).GetROProperty("y")
										
						Select Case strChartPos
					    	Case "rdoBottom"
					    		diff = yChart - yobjHorizontalSplitter
					    		If xChart= xPivotTab and diff <=20 Then
					    			strChartPosFound= 1
					    		Else 
					    			strChartPosFound= 0
					    		End If
					    		
					    	Case "rdoTop"
					    		diff = yPivotTab - yobjHorizontalSplitter
					    		If xChart= xPivotTab and diff <=20 Then
					    			strChartPosFound= 1
					    		Else 
					    			strChartPosFound= 0
					    		End If
					    		
					    	Case "rdoLeft"
					    		diff = xPivotTab - xobjVerticalSplitter
					    		If yChart= yPivotTab and diff <=20 Then
					    			strChartPosFound= 1
					    		Else 
					    			strChartPosFound= 0
					    		End If
					    		
					    	Case "rdoRight"
					    		diff = xChart - xobjVerticalSplitter
					    		If yChart= yPivotTab and diff <=20 Then
					    			strChartPosFound= 1
					    		Else 
					    			strChartPosFound= 0
					    		End If
					    End Select
					    
		End Select
				
		If strChartPosFound= 1 Then
			Call ReportStep (StatusTypes.Pass, "Chart is located at " &strChartPos& " of " &strSheetComp, strPageName)
			ValidateReportPosition = 1
		Else
			Call ReportStep (StatusTypes.Fail, "Chart is not located at " &strChartPos& " of " &strSheetComp, strPageName)
			ValidateReportPosition = 0
		End If
			
	End Function
	
	
	''' <summary>Validates the Report sheet /summary>
	''' <param name="strChartSheet" type="String">Sheet where Chart is created</param>
	''' <param name="strPageName" type="By Value">Screen location</param>
	''' <param name="objData" type="String">Reference to objData</param>
	Public function ValidateReportSheet(ByVal strChartSheet, ByVal strPageName, ByRef objData)

		Dim sheetFound, oDescSheet, oDescSheetObj, i, strSheetText
	   	sheetFound = 0
	   	
	   	'Descriptive Programming for Sheet
 				Set oDescSheet = Description.Create()
				oDescSheet("micclass").value = "WebElement"
				oDescSheet("html tag").value = "NOBR"
				oDescSheet("outertext").regularexpression=True
				oDescSheet("outertext").value = ".*Sheet.*"
				Set oDescSheetObj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(oDescSheet)
				wait 2
				
				For i = 0 To oDescSheetObj.count-1
					strSheetText = oDescSheetObj(i).GetROProperty("innertext")					
					If Trim(UCase(strSheetText)) = Trim(UCase(strChartSheet)) Then
						sheetFound = 1
						ValidateReportSheet = sheetFound
						Call ReportStep (StatusTypes.Pass, "Successfully found Sheet" &strChartSheet , strPageName)
						oDescSheetObj(i).click
						wait 2
						Exit For
					End If
				Next
				
				If sheetFound = 0 Then
					ValidateReportSheet = sheetFound
					Call ReportStep (StatusTypes.Fail, "Could not find Sheet" &strChartSheet , strPageName)
				End If
	
	End Function
	
	
	
	''' <summary>Validate the Report theme color for different components /summary>
	''' <param name="tabComponent" type="String">SCA Components</param>
	''' <param name="strColorTheme" type="String">Value to compare color theme of compenent</param>
	''' <param name="strPageName" type="By Value">Screen location</param>
	Public Sub ValidateReportColorTheme(ByVal tabComponent,ByVal strColorTheme,ByVal strPageName)

		Dim compOuterHtml,compTheme
	   	
	   	compOuterHtml = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable(tabComponent).GetROProperty("outerhtml")
		compTheme = InStr(1, Ucase(compOuterHtml),ucase(strColorTheme))
		
		If compTheme > 1 Then
			Call ReportStep (StatusTypes.Pass, "Choosed Component color theme " &strColorTheme , strPageName)
			Call ReportStep (StatusTypes.Pass, "Successfully Changed Component " &tabComponent& " with the color theme " &strColorTheme , strPageName)
		else
			Call ReportStep (StatusTypes.Fail, "Choosed Component color theme " &strColorTheme , strPageName)
			Call ReportStep (StatusTypes.Fail, "Could not changed Component " &tabComponent& " with the color theme successfully" &strColorTheme , strPageName)
		End If
	   	
	
	End Sub
	
	
	
	''' <summary>Validate the Pivot Table cell color /summary>
    ''' <param name="strColor" type="String">Color Values in RGB</param>
    ''' <param name="strConstVal" type="String">Constant Value of Cell data</param>
    ''' <param name="strFormulaOperator" type="By Value">Operator value to compare cell data</param>
    ''' <param name="strPagename" type="By Value">Web Page name</param>
    ''' <param name="objData" type="By Ref"></param>
    Public Function ValidatePivotTabCellColor(ByVal strColor, ByVal strConstVal, ByVal strFormulaOperator, ByVal strPagename, ByRef objData)
        Dim oTableCell, i, oTableCellObj, cellColor, cellData, FormulaOperator
        
        ValidatePivotTabCellColor = 0
        
        Browser("Analyzer").Page("ReportCreation").Sync
        Browser("Analyzer").Page("ReportCreation").RefreshObject
        
        'Block Keyboard & Mouse Inputs
        SystemUtil.BlockInput
        
        'Descriptive programme for Deny CheckBox
        Set oTableCell = Description.Create()
        oTableCell("micclass").value = "WebElement"
        oTableCell("class").value = "col000ms0"
        oTableCell("html id").value = ""
        oTableCell("html tag").value = "TD"
                                                                                                                        
        Set oTableCellObj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabTestPivotTable").ChildObjects(oTableCell)
        'msgbox oTableCellObj.count
        
        For i = 0 To oTableCellObj.count-1
            Browser("Analyzer").Page("ReportCreation").Sync
            Browser("Analyzer").Page("ReportCreation").RefreshObject
            cellColor = oTableCellObj(i).GetRoproperty("outerhtml")
            'msgbox cellColor
            wait 1
            cellData = oTableCellObj(i).GetRoproperty("innertext")
            'msgbox cellData
            If InStr(1, cellColor, "rgb(174, 255, 128)") <> 0 Then
            Select Case strFormulaOperator
                Case "Less Than"
                                If Trim(cellData) < CINT(strConstVal) Then
                                                ValidatePivotTabCellColor = ValidatePivotTabCellColor + 1
                                End If
                                
                Case "Less Than or Equal"
                                If Trim(cellData) <= CINT(strConstVal) Then
                                                ValidatePivotTabCellColor = ValidatePivotTabCellColor + 1
                                End If
                                
                Case "Equal"
                                If Trim(cellData) = CINT(strConstVal) Then
                                                ValidatePivotTabCellColor = ValidatePivotTabCellColor + 1
                                End If
                                
                Case "Not Equal"
                                If Trim(cellData) <> CINT(strConstVal)Then
                                                ValidatePivotTabCellColor = ValidatePivotTabCellColor + 1
                                End If
                                
                Case "Greater Than or Equal"
                                If Trim(cellData) >= CINT(strConstVal) Then
                                                ValidatePivotTabCellColor = ValidatePivotTabCellColor + 1
                                End If
                                
                Case "Greater Than"
                                If Trim(cellData) > CINT(strConstVal) Then
                                                ValidatePivotTabCellColor = ValidatePivotTabCellColor + 1
                                End If
                                
                Case "Between"
                                'TODO
            End Select
                            
            ElseIf i = oTableCellObj.count-1 and ValidatePivotTabCellColor = 0 Then
            	Call ReportStep (StatusTypes.Fail, "Found "&ValidatePivotTabCellColor&" cells with background color "&strColor& " set for cell data "&strFormulaOperator&" Constant Value " &strConstVal, strPagename)
            End If
		Next
	
		If ValidatePivotTabCellColor > 0 Then
		   Call ReportStep (StatusTypes.Pass, "Found "&ValidatePivotTabCellColor&" cells with background color "&strColor& " set for cell data "&strFormulaOperator&" Constant Value " &strConstVal, strPagename)
		End If
		
		Set oTableCell=Nothing
		
		'Unblock Keyboard & Mouse Inputs
		SystemUtil.UnBlockInput
	                    
	End Function
	
	
	''' <summary>Validate Sorting order of the Pivot Table by selecting column number of the table/summary>
	''' <param name="ColVal" type="By Value">Column to be sorted</param>
	''' <param name="strSortType" type="By Value">Type of sorting operation to be performed</param>
	''' <param name="rowTotal" type="By Value">Total no of rows in table</param>
	''' <param name="objData" type="By Ref"></param>
	Public Function ValidationSorting(ByVal ColVal, ByVal strSortType, Byval rowTotal, ByRef objTable)
		
		Dim rowStart, rowSecondary, sDataA, sDataB
				
		Select Case strSortType
			Case "Ascending"
							For rowStart = 1 to rowTotal-2
								For rowSecondary = rowStart to rowTotal-1
									If objTable.GetCellData(rowStart,ColVal) <> "-" and objTable.GetCellData(rowSecondary,ColVal) <> "-" Then
							        	sDataA = cdbl(objTable.GetCellData(rowStart,ColVal))
							            sDataB = cdbl(objTable.GetCellData(rowSecondary,ColVal))
										If (sDataB > sDataA) Then
							            	ValidationSorting = "True" 
										ElseIf (sDataB= sDataA) Then
							                ValidationSorting = "True" 
										ElseIf (sDataB<sDataA) < 0 Then
							                ValidationSorting = "False" 
							            	Exit For
							            End If
							
							            If rowSecondary = rowTotal Then
							            	Exit For
							            End If
							       	End if
								Next
							    
							    If ValidationSorting = "False" Then
									Exit For
							    End If
							Next
							
			Case "Descending"
							'TODO
		End Select
		
	End Function
	
	
	Public Function VerifyEmail(ByVal strSubject)
		
		'Object
		Dim objOutlookApp,objNameSpace,inboxFolders,junkFolder,inboxCount,junkCount
		'Integer
		Dim mailWaitTime,i
		'String
		Dim strBody,strMailSubject,strinboxMail,strjunkMail
		'Boolean
		Dim found
		
		mailWaitTime = 1800
		strBody = "false"
		
		Set objOutlookApp = CreateObject("Outlook.Application")
		Set objNameSpace = objOutlookApp.GetNamespace("MAPI")
		Set inboxFolders = objNameSpace.GetDefaultFolder(6)
		'Set junkFolder = objNameSpace.GetDefaultFolder(23)
		
		'Msgbox junkFolder.Items.Count
		'Msgbox inboxFolders.Items.Count
	
		Set inboxCount = inboxFolders.Items
		'Set junkCount = junkFolder.Items
		
		i = 1
		found = False
		While i <= mailWaitTime And Not found
			For each strinboxMail in inboxCount
				If strinboxMail.unread Then
					strMailSubject = strinboxMail.Subject
					If Instr(LCase(strMailSubject),LCase(strSubject)) <> 0 then 
						strBody = strinboxMail.body
						found = True
						'Msgbox "Email matched - " &strMailSubject
						'Msgbox strBody
						'Msgbox strinboxMail.sendername
					End If
				End if
			Next
			i = i + 1						
		Wend
			
		REM i = 1
		REM found = False
		REM While i <= mailWaitTime And Not found
			REM For each strjunkMail in junkCount
				REM If strjunkMail.unread Then
					REM strMailSubject = strjunkMail.Subject
					REM If Instr(LCase(strMailSubject),LCase(strSubject)) <> 0 then 
						REM strBody = strjunkMail.body
						REM found = True
						REM 'Msgbox strBody
						REM 'Msgbox "Email matched - " &strMailSubject
						REM 'Msgbox strjunkMail.sendername
					REM End If
				REM End if 
			REM Next
		REM Wend
		
		VerifyEmail = Trim(strBody)
	
	End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	
	


	

	'<Composed By : Shweta B Nagaral>
	'<Clicks on Down Normal menu of any component to validate options under each menu>
    '<strPageName: name of report creation page :  String>
    '<strTabName: Tab name in component down normal menu. It could be "Design", "Analyze", "Advanced"  :  String>
    '<strCompVal1, strCompVal2, strCompVal3, strCompVal4, strCompVal5, strCompVal6, strCompVal7, strCompVal8, strCompVal9, strCompVal10: Menu validation options under each tab :  String>
	Public Sub ValidateDownNormalMenu(ByVal strPageName, ByVal strTabName, ByVal strCompVal1, ByVal strCompVal2, ByVal strCompVal3, ByVal strCompVal4, ByVal strCompVal5, ByVal strCompVal6, ByVal strCompVal7, ByVal strCompVal8, ByVal strCompVal9, ByVal strCompVal10)
	
		Dim objImgPivotDownNormal, objMenuOptionVal1, objMenuOptionVal2, objMenuOptionVal3, objMenuOptionVal4, objMenuOptionVal5, objMenuOptionVal6, objMenuOptionVal7, objMenuOptionVal8
		
		Browser("Analyzer").Page("ReportCreation").Sync
		wait 2
		'Mouse over on Pivot Table menu
		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgPivotDownNormal").FireEvent "onmouseover"
		wait 2
		Set objImgPivotDownNormal=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgPivotDownNormal")
		Call SCA.ClickOn(objImgPivotDownNormal,"imgPivotDownNormal", "ReportCreation")
		
		Browser("Analyzer").Page("ReportCreation").Sync
		
		'Set values for chart creation
		'Mouse over on Tab name in pivot table menu. It could be "Design", "Analyze", "Advanced"
		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image(strTabName).FireEvent "onmouseover"
   		wait 2
   		   		  		
   		'Descriptive program to find total count of Options under each tab
   		
   		Select Case strTabName
   			Case "imgDesign"
   					
   					Set objMenuOptionVal1 = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strCompVal1)
					Call ExistanceWebelements_Verification(objMenuOptionVal1, strCompVal1, strPageName, 1)
   					Set objMenuOptionVal2 = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strCompVal2)
					Call ExistanceWebelements_Verification(objMenuOptionVal2, strCompVal2, strPageName, 1)
					Set objMenuOptionVal3 = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strCompVal3)
					Call ExistanceWebelements_Verification(objMenuOptionVal3, strCompVal3, strPageName, 1)
					Set objMenuOptionVal4 = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strCompVal4)
					Call ExistanceWebelements_Verification(objMenuOptionVal4, strCompVal4, strPageName, 1)
					Set objMenuOptionVal5 = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strCompVal5)
					Call ExistanceWebelements_Verification(objMenuOptionVal5, strCompVal5, strPageName, 1)
					Set objMenuOptionVal6 = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strCompVal6)
					Call ExistanceWebelements_Verification(objMenuOptionVal6, strCompVal6, strPageName, 1)
					Set objMenuOptionVal7 = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strCompVal7)
					Call ExistanceWebelements_Verification(objMenuOptionVal7, strCompVal7, strPageName, 1)
					Set objMenuOptionVal8 = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strCompVal8)
					Call ExistanceWebelements_Verification(objMenuOptionVal8, strCompVal8, strPageName, 1)
   					
   			Case "imgAnalyticChart"
   				'TODO if required or In Scope 
   			Case "Advanced"
   				'TODO if required or In Scope
   		End Select
   		
	End Sub
	
	
	
	'''' <summary>Verification of the Webelements\button\Radio or any objects of Application</summary>
	'	''' <param name="objName" type="By Reference">object Name</param>	
	'	''' <param name="objValue" type="By Reference">object Value</param>	
	'	''' <param name="strPageName" type="By Value">Screen location</param>
	'	''' <param name="retStatus" type="By Reference"> retStatus i.e. 0 or 1. Mention retStatus = 0 if existence status of object is required in calling function </param>
	Public Function ExistanceWebelements_Verification(ByRef objName, ByVal objValue, ByVal strPageName, ByVal retStatus)
	 
		If objName.exist(4) Then	
			Call ReportStep (StatusTypes.Pass, objValue& " Object Exists", strPageName)
			Else
			Call ReportStep (StatusTypes.Fail, objValue& " Object deos not Exist" , strPageName)
		End If
		
		If retStatus = 0 Then
			ExistanceWebelements_Verification=objName.exist(4)
		End If
		
	End Function


	'<Composed By : Shweta B Nagaral>
	'<Validates values/sub-values of Down Normal menu of any component>
    '<strTabName: Tab name in pivot table menu. It could be "Design", "Analyze", "Advanced"  :  String>
    '<strTabOption: Select options under each tab :  String>
    '<strTabSubOption: Select sub-options inside options under each tab :  String>
    '<strCompareTabVal: Actual Compare value to be validated against run time values/sub-values of Down Normal menu of any component :  String>
	Public Function ValidateTabInnerText(ByVal strCompTab, ByVal strTabName, ByVal strTabOption, ByVal strTabSubOption, ByVal strCompareTabVal)
			Dim strActualTabVal, objValue, objImgPivotDownNormal
			
			Browser("Analyzer").Page("ReportCreation").Sync
			
			'<shweta - 20/8/2015 Added "strCompTab" parameter> - Starts
'			'Mouse over on Component Down normal Menu
'			Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgPivotDownNormal").FireEvent "onmouseover"
'			wait 2
'	        'Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgPivotDownNormal").Click
'			Set objImgPivotDownNormal=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgPivotDownNormal")
'			Call SCA.ClickOn(objImgPivotDownNormal,"imgPivotDownNormal", "ReportCreation")
	
			'Mouse over on Component Down normal Menu
			Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image(strCompTab).FireEvent "onmouseover"
			wait 2
	        'Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgPivotDownNormal").Click
			Set objImgPivotDownNormal=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image(strCompTab)
			Call SCA.ClickOn(objImgPivotDownNormal, strCompTab, "ReportCreation")
			'<shweta - 20/8/2015 Added "strCompTab" parameter> - End
			
			Browser("Analyzer").Page("ReportCreation").Sync
			wait 5
			
			'Set values for chart creation
			'Mouse over on Tab name in pivot table menu. It could be "Design", "Analyze", "Advanced"
			'Browser("Hosting Templates - Ops").Page("IMS Analysis Manager_2").Frame("ReportWindowOfAnalyzer_3").Image("strTabName").FireEvent "onmouseover"
			Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image(strTabName).FireEvent "onmouseover"
	   		wait 2


	   		'Select options under each tab
	   		If strTabSubOption <> "" Then
	   			'Shweta - TODO: Need to add last param object details - Start
		   		'			Select sub-options inside options under each tab
				'   		Set objSheet=Browser("Analyzer").Page("ReportCreation").Frame("CreateChart").WebList("ddlSheets")
				'   		Call SCA.SelectFromDropdown(objSheet,strChartSheet)
				'Shweta - TODO: Need to add last param object details - End
	   		Else 
	   			Set objValue=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strTabOption)
	   			wait 1
				strActualTabVal = objValue.GetROProperty("InnerText")
				If Trim(Ucase(strActualTabVal)) = Trim(Ucase(strCompareTabVal)) Then
					Call ReportStep (StatusTypes.Pass, strCompareTabVal& " is displaying under Design Tab", "Report Creation Page")
					Browser("Analyzer").Page("ReportCreation").Sync
					Browser("Analyzer").Page("ReportCreation").RefreshObject
					ValidateTabInnerText = 0
					'objValue.click
					Call SCA.ClickOn(objValue,strTabOption, "ReportCreation")
					Browser("Analyzer").Page("ReportCreation").Sync
					Browser("Analyzer").Page("ReportCreation").RefreshObject
													
				else
					Call ReportStep (StatusTypes.Fail, strCompareTabVal& " is not displaying under Design Tab", "Report Creation Page")
					ValidateTabInnerText = 1
				End If
				
	   		End If
	   		
	   		Browser("Analyzer").Page("ReportCreation").Sync
			Browser("Analyzer").Page("ReportCreation").RefreshObject
			wait 2
	End Function


	'<Composed By : Shweta B Nagaral>
	'<Checks no of rows/cols set when valiating rows/cols per page functionality of Down Normal menu of Group Table component>
    '<pageType: welRowsPerPage/welColumnsPerPage option to validate :  String>
    '<selectionVal: row/col count values :  String>
    '<rowCnt: Row count to be expected :  String>
    '<colCnt: Col count to be expected :  String>
    '<tabAtrributeName: tabGrpTableMeasureData/tabGrpTableColAttributeData to validate :  String>
    '<strCellValues_Before: Array Value of Group Table when complete table is displayed:  String>
	'<objData: Reference to objData>
	Public Function ValidationRowColPage(ByVal pageType, ByVal selectionVal, ByVal rowCnt, ByVal colCnt, ByVal tabAtrributeName, ByVal strCellValues_Before, ByVal objData)
		
		Dim oComp, oCompObj, objNewwebtable, ArrValue_After, strCellValue_After, strCellValues_After, reVal
		Dim m, n, lenAfter
		
		'1) strArrayFullTable
		Browser("Analyzer").Page("ReportCreation").Sync
		Browser("Analyzer").Page("ReportCreation").RefreshObject
		
'		'Shweta <28/12/2016> - added Higlight and firevents on 'tabDesignStepGrpTable' - Start
'		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable(tabAtrributeName).Highlight
'		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable(tabAtrributeName).FireEvent "onmouseover"
'		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable(tabAtrributeName).FireEvent "onclick"
'		'Shweta <28/12/2016> - added Higlight and firevents on 'tabDesignStepGrpTable' - End
		
'		'Shweta <28/12/2016> - added Higlight and firevents on 'tabDesignStepGrpTable' - Start
'		Set objFilterCnd = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("html tag:=NOBR","innertext:= Drop a Filter Condition Here")
'		objFilterCnd.FireEvent "onclick"
'		wait 2
'		objFilterCnd.FireEvent "onmouseover"
'		'Shweta <28/12/2016> - added Higlight and firevents on 'tabDesignStepGrpTable' - End
		
		'2) Click on Row/Col Per Page
		'Call IMSSCA.General.DownNormalMenu("imgDesign", pageType, selectionVal)
		Call IMSSCA.General.ComponentDownNormalMenu(tabAtrributeName, "imgDesign", pageType, selectionVal)
		Browser("Analyzer").Page("ReportCreation").Sync
		Browser("Analyzer").Page("ReportCreation").RefreshObject
		
		'3) descriptive program to find total no of table for "tabAtrributeName"
		If tabAtrributeName = "tabGrpTableColAttributeData" Then
			Set oComp  = Description.Create()
			oComp("micclass").value = "WebTable"
			oComp("cols").value = 1
			oComp("html tag").value = "TABLE"
			oComp("name").value = "gcollapse"
			oComp("rows").value = rowCnt+1
			
			Set oCompObj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabPivotTable").ChildObjects(oComp)
			'msgbox oCompObj.count
			
		ElseIf tabAtrributeName = "tabGrpTableMeasureData" Then
		
			Set oComp  = Description.Create()
			oComp("micclass").value = "WebTable"
			oComp("cols").value = colCnt+1
			oComp("html tag").value = "TABLE"
			oComp("name").value = "WebTable"
			oComp("rows").value = rowCnt
			
			Set oCompObj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabPivotTable").ChildObjects(oComp)
			'msgbox oCompObj.count
			
		End If
		
		wait 2
		Browser("Analyzer").Page("ReportCreation").Sync
		Browser("Analyzer").Page("ReportCreation").RefreshObject
			
		'4) table Comparision
		If oCompObj.count > 0 Then
			ValidationRowColPage = 1
			Set objNewwebtable=oCompObj(0)
			ArrValue_After=SCA.Webtable(objNewwebtable,"Dimension Value table",TRIM("Retriving_DataTableValue"),"Report Creation","","","","")
			For m=0 to ubound(ArrValue_After,1)
			    For n=0 to ubound(ArrValue_After,2)
			        strCellValue_After= ArrValue_After(m,n)
			        strCellValues_After=strCellValues_After&";"&strCellValue_After
			    Next
			Next
			
			If tabAtrributeName = "tabGrpTableColAttributeData" Then
					reVal=IMSSCA.Validations.FilterVerification(strCellValues_Before,strCellValues_After)
					If reVal >= 1 Then
						Call ReportStep (StatusTypes.Pass, "########################Complete Group Table values are greater than table values when "&rowCnt&" Row Per Page  is set", "ReportCreation Page")
					ElseIf reVal = 0 Then
						Call ReportStep (StatusTypes.Pass, "########################Complete Group Table values are same as table values when "&rowCnt&" Row Per Page  is set", "ReportCreation Page")
					ElseIf reVal < 1 Then
						Call ReportStep (StatusTypes.Fail, "########################Complete Group Table values are lesser than table values when "&rowCnt&" Row Per Page  is set", "ReportCreation Page")
					End If
			
			ElseIf tabAtrributeName = "tabGrpTableMeasureData" Then
					lenAfter = Len(strCellValues_After)
					If strCellValues_Before > lenAfter Then
						Call ReportStep (StatusTypes.Pass, "########################Complete Group Table values are greater than table values when "&colCnt&" Columns Per Page  is set", "ReportCreation Page")
					ElseIf strCellValues_Before = lenAfter Then
						Call ReportStep (StatusTypes.Pass, "########################Complete Group Table values are same as table values when "&colCnt&" Columns Per Page  is set", "ReportCreation Page")
					ElseIf strCellValues_Before < lenAfter  Then
						Call ReportStep (StatusTypes.Fail, "########################Complete Group Table values are lesser than table values when "&colCnt&" Columns Per Page  is set", "ReportCreation Page")
					End If
				
			End If
			
		Else
		
			ValidationRowColPage = 0
		End If
	
	End Function
	
	
	''' <summary>Filter condition verification on the Pivotal table</summary>
	''' <param name="strBFilterValues" type="By Value">Field  Value of Report Before Filteration</param>	
	''' <param name="strAFilterValues" type="By Value">Field  Value of Report After Filteration</param>
	''' <param name="strReportName" type="By Value">Field  Value of Report Name</param>
	''' <param name="strPageName" type="By Value">Screen Location</param>
	
	Public Function FilterVerification(ByVal strBOValues,ByVal strAOValues)
	
	Dim returnVAl
	  
	     returnVAl=Strcomp(strBOValues,strAOValues)
	     FilterVerification=returnVAl
		    
	End Function
	
	
	''' <summary>Verify the renamed value in the fields (i.e WebEdit, WebElement and WebList)</summary>
	''' <param name="objField" type="Object">Field object</param>
	''' <param name="strFieldName" type="By Value">Field name of the value to get</param>
	''' <param name="strActualFieldValue" type="By Value"> Actual Field value before renaming Field object</param>
	''' <param name="strCompareValue" type="Scripting.Dictionary">Data value to compare</param>
	''' <param name="strPageName" type="By Value">Screen location</param>

	Public Sub ValidateReplaceField(ByRef objField , ByVal strFieldName , ByVal strActualFieldValue, ByVal strCompareValue , ByVal strPageName)
		wait 5
'		If Not (objField.Exist(2)) Then
'			Call ReportStep (StatusTypes.Fail, "Could not validate the '"&strFieldName&"' :field as it does not exist", strPageName)
'			Exit Sub
'		End If
		
		If (Not(objField.GetROProperty("micclass")="WebEdit")) And (Not(objField.GetROProperty("micclass")="WebList"))  And (Not(objField.GetROProperty("micclass")="WebElement")) Then
			Call ReportStep (StatusTypes.Warning, "Object passed to the function is not a WebEdit or WebList. Please check the object for the WebEdit or WebList : '"&strFieldName&"'", strPageName)
			Exit Sub
		End If
		
		wait 2
		If strActualFieldValue <> strCompareValue Then
					Call ReportStep (StatusTypes.Pass, "Successfully Renamed '"&strActualFieldValue&"' with the Supplied text '"&strCompareValue&"' in "&strFieldName, strPageName)
		Else		
					Call ReportStep (StatusTypes.Fail, "Could not successfully rename '"&strActualFieldValue&"' with the Supplied text '"&strCompareValue&"' in "&strFieldName, strPageName)
		End If
		
		
		If Trim(objField.GetROProperty("default value"))=Trim(strCompareValue) Or Trim(objField.GetROProperty("innertext"))=Trim(strCompareValue) Or Trim(objField.GetROProperty("value"))=Trim(strCompareValue) Then
					Call ReportStep (StatusTypes.Pass, "Supplied text '"&strCompareValue&"' is displayed in "&strFieldName, strPageName)
		Else
					Call ReportStep (StatusTypes.Fail, "Supplied text '"&strCompareValue&"' is not displayed in "&strFieldName, strPageName)
		End If
		
	End Sub
	
	
'################################################################################################################

	
	
	
	

	
	




'################################################################################################################

	
	''' <summary>Validate Group Table Filter Dialogue page existence, </summary>
	''' <param name="strOptions" type="By Value"> Group Table Name </param>	
	''' <Author> Shilpi Jha</Author>
	''' <Date Created> 5-5-2015</Date>
		Public Sub ValidateGroupFilterPageExistence()
		
'	If Browser("Analyzer").Frame("GroupFilter_dialogBox") Then
'			Call ReportStep (StatusTypes.Pass,"Filter Group Table Dialog exists", "Group Table Filter page")
'		Else
'			Call ReportStep (StatusTypes.Fail,"Group Table Filter Dialog does not exist", "Group Table Filter page")
'		End If
		
		If Browser("Analyzer").Page("GroupCreationPage").Frame("GroupFilterCreation").WebEdit("txtName").Exist(10) And Browser("Analyzer").Page("GroupCreationPage").Frame("GroupFilterCreation").WebEdit("txtCaption").Exist(10) Then
			Call ReportStep (StatusTypes.Pass,"Filter Group Table Dialog components", "Group Table Filter page")
		ElseIf Browser("Analyzer").Page("GroupCreationPage").Frame("GroupFilterCreation").WebEdit("txtName").Exist(10) And Not Browser("Analyzer").Page("GroupCreationPage").Frame("GroupFilterCreation").WebEdit("txtCaption").Exist(5) Then
			Call ReportStep (StatusTypes.Fail,"Caption Edit Box", "Group Table Filter page")	
        ElseIf NOT Browser("Analyzer").Page("GroupCreationPage").Frame("GroupFilterCreation").WebEdit("txtName").Exist(10) And Not Browser("Analyzer").Page("GroupCreationPage").Frame("GroupFilterCreation").WebEdit("txtCaption").Exist(5) Then	
		    Call ReportStep (StatusTypes.Fail,"Name Edit Box", "Group Table Filter page")	
		End If
	End Sub
		
	
	''' <summary>Verify the existence of Chart /summary>
	''' <param name="" type=""> </param>
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
	
	
	
	'===========================================================================================================================================================   
	    
	    '    '<Validates XmL tags based on input array contents>
	    '    '<aReportTagKeyval :- array values to Report Tag key values>
	    '    '<aReportTagCompval :- array values to Report Tag Comparing values>
	    '	 '<aCompDataSrcKeyVal:- array values to ComponentDataSource Tag key values>
	    '	 '<aCompDataSrcCompVal:- array values to ComponentDataSource Tag Comparing values>
	    '	 '<aDataSrcKeyVal:- array values to DataSources Tag key values>
	    '	 '<aDataSrcCompVal:- array values to DataSources Tag Comparing values>
	    '    '<author :-Shweta Nagaral>'
	'===========================================================================================================================================================
	
    Public Function ValidateSCAXMLReport(ByVal intReportCnt, ByVal intDataVal, ByVal intCloseXML, ByVal aReportTagKeyval, ByVal aReportTagCompval, ByVal aCompDataSrcKeyVal, ByVal aCompDataSrcCompVal, ByVal aDataSrcKeyVal, ByVal aDataSrcCompVal, ByVal ReportOwner,ByVal strFileName,ByVal strTagName,ByVal strAttribute)
	
		Dim i, j, k, m
		Dim xmlVer, strReportOwner, reportCompSearch, strCompDataSrc, valCompSearch, strDataSrc, valDataSearch, reportOpened
		Dim AttName()
		reportOpened = 0
		
'		'shweta 21/4/2016 Commented Not required to oprn the xml file when using XML dom object to read/parse XML file - Start
'		Dim WshShell
'		Set WshShell = CreateObject("WScript.Shell")
'		WshShell.run strFileName
'		Set WshShell = Nothing
'		wait 2
'		'shweta 21/4/2016 Commented Not required to oprn the xml file when using XML dom object to read/parse XML file - End
		
''		For k = 1 To 10 Step 1
'			'Shweta 5/1/2016- Added waitproperty as sync stmt - Start
'		 	Browser("OtherReports").Page("XMLExportedReport").WebElement("welXmlVersion").WaitProperty "visible",True,30000
'		 	'Shweta 5/1/2016- Added waitproperty as sync stmt - End
'		 	
'			If Browser("OtherReports").Page("XMLExportedReport").WebElement("welXmlVersion").Exist(1) Then
'				xmlVer = Browser("OtherReports").Page("XMLExportedReport").WebElement("welXmlVersion").GetROProperty("innertext")
'				Call ReportStep (StatusTypes.Pass, "Validation of XML file opened in browser. XML file is opened in browser", "SCA Exported XML File")
'				reportOpened = 1
'				Exit for
'			End If
'			
'			If k=10 and reportOpened = 0 Then
'				Call ReportStep (StatusTypes.Warning, "Validation of XML file opening in browser", "SCA Exported XML File")
'				Call ReportStep (StatusTypes.Warning, "XML file is not opened in browser", "SCA Exported XML File")
'			End If
'		Next
		
    	If ReportOwner=0 Then
    		ReDim AttName(ubound(strAttribute))
    		set  xmldom = Createobject("MSXML.DOMDocument")
			xmldom.async = "False"
			xmldom.Load(strFileName)
			Set nodelist = xmldom.getElementsByTagName(strTagName)
			
				If nodelist.length > 0 then
    				For each x in nodelist
    					For i = 0 to ubound(strAttribute) Step 1
    						AttName(i)=x.getattribute(strAttribute(i))  	
    					Next
        			
    				Next
				Else
  				Call ReportStep (StatusTypes.Fail,"Field Not Found", "Xml Validation Page") 
				End If
    		
    		ValidateSCAXMLReport=AttName
    		Exit Function
    		
    	End If 
		If UBound(aReportTagKeyval) <> -1 or UBound(aReportTagCompval) <> -1 Then
			
			Set oReportTag  = Description.Create()
			oReportTag("micclass").value = "Link"
			oReportTag("html tag").value = "a"
			oReportTag("class").value = "collapse"
			oReportTag("innertext").regularexpression=True
			oReportTag("innertext").value="<Report.*ReportName.*"
			Set oReportTagObj = Browser("OtherReports").Page("XMLExportedReport").ChildObjects(oReportTag)
			'msgbox oReportTagObj.count
			wait 1
			
			If intReportCnt <= oReportTagObj.count Then
				
				'search for ReportName inside <Report> Tag 
				strReportOwner = oReportTagObj(intReportCnt-1).GetROProperty("innertext")
				If UBound(aReportTagKeyval) = UBound(aReportTagCompval) Then
					For i = 0 To UBound(aReportTagCompval) Step 1
						reportCompSearch = instr(1, Trim(Ucase(strReportOwner)), Trim(Ucase(aReportTagKeyval(i)&"="&chr(34)&aReportTagCompval(i))))
						If reportCompSearch > 0 Then
							Call ReportStep (StatusTypes.Pass, aReportTagKeyval(i)&"="&chr(34)&aReportTagCompval(i)&" found successfully inside Report Tag", "SCA Exported XML File")
						Else 
							Call ReportStep (StatusTypes.Fail, aReportTagKeyval(i)&"="&chr(34)&aReportTagCompval(i)&" is not found inside Report Tag", "SCA Exported XML File")
						End If
					Next
				Else
						Call ReportStep (StatusTypes.Fail, "Could not test values in Report tag  successfully. Input params mismatch", "SCA Exported XML File")
				End If
				
			End If
			
		Else
		
			Call ReportStep (StatusTypes.Information, "Skipping Validation of test values in Report tag", "SCA Exported XML File")
		End If
		
		If UBound(aCompDataSrcKeyVal) <> -1 or UBound(aCompDataSrcCompVal) <> -1 Then
			
			Set oCompTag  = Description.Create()
			oCompTag("micclass").value = "WebElement"
			oCompTag("html tag").value = "span"
			oCompTag("class").value = ""
			oCompTag("innertext").regularexpression=True
			oCompTag("innertext").value="<ComponentDataSources>.*OldDataSrcId.*OldDatabase.*OldCube.*NewDataSrcId.*NewDatabase.*NewCube.*"
			Set oCompTagObj = Browser("OtherReports").Page("XMLExportedReport").ChildObjects(oCompTag)
			'msgbox oCompTagObj.count
			wait 1
			
			If intReportCnt <= oCompTagObj.count Then
				
				'search for ComponentDataSources values inside <ComponentDataSources> Tag 	
				strCompDataSrc = oCompTagObj(intReportCnt-1).GetROProperty("innertext")
				If UBound(aCompDataSrcKeyVal) = UBound(aCompDataSrcCompVal) Then
					For j = 0 To UBound(aCompDataSrcCompVal) Step 1
						valCompSearch = instr(1, Trim(Ucase(strCompDataSrc)), Trim(Ucase(aCompDataSrcKeyVal(j)&"="&chr(34)&aCompDataSrcCompVal(j))))			
						If valCompSearch > 0 Then
							Call ReportStep (StatusTypes.Pass, aCompDataSrcKeyVal(j)&"="&chr(34)&aCompDataSrcCompVal(j)&" found successfully inside Component Tag", "SCA Exported XML File")
						Else 
							Call ReportStep (StatusTypes.Fail, aCompDataSrcKeyVal(j)&"="&chr(34)&aCompDataSrcCompVal(j)&" is not found inside Component Tag", "SCA Exported XML File")
						End If
					Next
				Else
						Call ReportStep (StatusTypes.Fail, "Could not test values in Components tag successfully. Input params mismatch", "SCA Exported XML File")
				End If
			
			End If
			
		Else
			Call ReportStep (StatusTypes.Information, "Skipping Validation of test values in Components tag", "SCA Exported XML File")
		End If
		
		If intDataVal = 1 Then
			
			If UBound(aDataSrcKeyVal) <> -1 or UBound(aDataSrcCompVal) <> -1 Then
				
				Set oDataSourceTag  = Description.Create()
				oDataSourceTag("micclass").value = "WebElement"
				'oDataSourceTag("html tag").value = "DataSource"
				oDataSourceTag("html tag").value = "a"
				oDataSourceTag("innertext").regularexpression=True
				oDataSourceTag("innertext").value="<DataSource.*"
				oDataSourceTag("outertext").regularexpression=True
				oDataSourceTag("outertext").value="<DataSource.*"
				Set oDataSourceTagObj = Browser("OtherReports").Page("XMLExportedReport").ChildObjects(oDataSourceTag)
				'msgbox oDataSourceTagObj.count
				wait 1
				
				'search for DataSources values inside <DataSources> Tag
				strDataSrc = oDataSourceTagObj(0).GetROProperty("innertext")
				If UBound(aDataSrcKeyVal) = UBound(aDataSrcCompVal) Then
					For m = 0 To UBound(aDataSrcCompVal) Step 1
						valDataSearch = instr(1, Trim(Ucase(strDataSrc)), Trim(Ucase(aDataSrcKeyVal(m)&"="&chr(34)&aDataSrcCompVal(m))))			
						If valDataSearch > 0 Then
							Call ReportStep (StatusTypes.Pass, aDataSrcKeyVal(m)&"="&chr(34)&aDataSrcCompVal(m)&" found successfully inside DataSource Tag", "SCA Exported XML File")
						Else 
							Call ReportStep (StatusTypes.Fail, aDataSrcKeyVal(m)&"="&chr(34)&aDataSrcCompVal(m)&" is not found inside DataSource Tag", "SCA Exported XML File")
						End If
					Next
				Else
						Call ReportStep (StatusTypes.Fail, "Could not test values in DataSource tag successfully. Input params mismatch", "SCA Exported XML File")
				End If
				
			Else
				Call ReportStep (StatusTypes.Information, "Skipping Validation of test values in DataSource tag", "SCA Exported XML File")
			End If
			
		End If
		
		'Close the xml opened file
		If intCloseXML = 1 Then
		'TODO : Get Runtime Browser("CreationTime") and close it
			For r = 1 To 10 Step 1
				If Browser("CreationTime:=1").Exist(2) Then
					 Browser("CreationTime:=1").Close
					 Call ReportStep (StatusTypes.Pass, "Successfully closed XML opened in IE browser", "XML Exported Page")
					 Exit For
				End If 
				
				If r = 10 Then
					Call ReportStep (StatusTypes.Fail, "Could not close XML opened in IE browser successfully", "XML Exported Page")
				End If
			Next
		End If
		
	End Function
	
	
'		
'    Public Function ValidateSCAXMLReportBackup(ByVal intReportCnt, ByVal intDataVal, ByVal intCloseXML, ByVal aReportTagKeyval, ByVal aReportTagCompval, ByVal aCompDataSrcKeyVal, ByVal aCompDataSrcCompVal, ByVal aDataSrcKeyVal, ByVal aDataSrcCompVal, ByVal ReportOwner,ByVal strFileName,ByVal strTagName,ByVal strAttribute)
'	
'		Dim i, j, k, m
'		Dim xmlVer, strReportOwner, reportCompSearch, strCompDataSrc, valCompSearch, strDataSrc, valDataSearch, reportOpened
'		reportOpened = 0
'		
'		'shweta 21/4/2016 - Start
'		Dim WshShell
'		Set WshShell = CreateObject("WScript.Shell")
'		WshShell.run strFileName
'		Set WshShell = Nothing
'		wait 2
'		'shweta 21/4/2016 - End
'		
'''		For k = 1 To 10 Step 1
''			'Shweta 5/1/2016- Added waitproperty as sync stmt - Start
''		 	Browser("OtherReports").Page("XMLExportedReport").WebElement("welXmlVersion").WaitProperty "visible",True,30000
''		 	'Shweta 5/1/2016- Added waitproperty as sync stmt - End
''		 	
''			If Browser("OtherReports").Page("XMLExportedReport").WebElement("welXmlVersion").Exist(1) Then
''				xmlVer = Browser("OtherReports").Page("XMLExportedReport").WebElement("welXmlVersion").GetROProperty("innertext")
''				Call ReportStep (StatusTypes.Pass, "Validation of XML file opened in browser. XML file is opened in browser", "SCA Exported XML File")
''				reportOpened = 1
''				Exit for
''			End If
''			
''			If k=10 and reportOpened = 0 Then
''				Call ReportStep (StatusTypes.Warning, "Validation of XML file opening in browser", "SCA Exported XML File")
''				Call ReportStep (StatusTypes.Warning, "XML file is not opened in browser", "SCA Exported XML File")
''			End If
''		Next
'		
'    	If ReportOwner=0 Then
'    		
'    		set  xmldom = Createobject("MSXML.DOMDocument")
'			xmldom.async = "False"
'			xmldom.Load(strFileName)
'			Set nodelist = xmldom.getElementsByTagName(strTagName)
'			
'				If nodelist.length > 0 then
'    				For each x in nodelist
'        				AttName=x.getattribute(strAttribute)  
'        			
'    				Next
'				Else
'  				Call ReportStep (StatusTypes.Pass,"Field Not Found", "Xml Validation Page") 
'				End If
'    		
'    		ValidateSCAXMLReport=AttName
'    		Exit Function
'    		
'    	End If 
'		If UBound(aReportTagKeyval) <> -1 or UBound(aReportTagCompval) <> -1 Then
'			
'			Set oReportTag  = Description.Create()
'			oReportTag("micclass").value = "Link"
'			oReportTag("html tag").value = "a"
'			oReportTag("class").value = "collapse"
'			oReportTag("innertext").regularexpression=True
'			oReportTag("innertext").value="<Report.*ReportName.*"
'			Set oReportTagObj = Browser("OtherReports").Page("XMLExportedReport").ChildObjects(oReportTag)
'			'msgbox oReportTagObj.count
'			wait 1
'			
'			If intReportCnt <= oReportTagObj.count Then
'				
'				'search for ReportName inside <Report> Tag 
'				strReportOwner = oReportTagObj(intReportCnt-1).GetROProperty("innertext")
'				If UBound(aReportTagKeyval) = UBound(aReportTagCompval) Then
'					For i = 0 To UBound(aReportTagCompval) Step 1
'						reportCompSearch = instr(1, Trim(Ucase(strReportOwner)), Trim(Ucase(aReportTagKeyval(i)&"="&chr(34)&aReportTagCompval(i))))
'						If reportCompSearch > 0 Then
'							Call ReportStep (StatusTypes.Pass, aReportTagKeyval(i)&"="&chr(34)&aReportTagCompval(i)&" found successfully inside Report Tag", "SCA Exported XML File")
'						Else 
'							Call ReportStep (StatusTypes.Fail, aReportTagKeyval(i)&"="&chr(34)&aReportTagCompval(i)&" is not found inside Report Tag", "SCA Exported XML File")
'						End If
'					Next
'				Else
'						Call ReportStep (StatusTypes.Fail, "Could not test values in Report tag  successfully. Input params mismatch", "SCA Exported XML File")
'				End If
'				
'			End If
'			
'		Else
'		
'			Call ReportStep (StatusTypes.Information, "Skipping Validation of test values in Report tag", "SCA Exported XML File")
'		End If
'		
'		If UBound(aCompDataSrcKeyVal) <> -1 or UBound(aCompDataSrcCompVal) <> -1 Then
'			
'			Set oCompTag  = Description.Create()
'			oCompTag("micclass").value = "WebElement"
'			oCompTag("html tag").value = "span"
'			oCompTag("class").value = ""
'			oCompTag("innertext").regularexpression=True
'			oCompTag("innertext").value="<ComponentDataSources>.*OldDataSrcId.*OldDatabase.*OldCube.*NewDataSrcId.*NewDatabase.*NewCube.*"
'			Set oCompTagObj = Browser("OtherReports").Page("XMLExportedReport").ChildObjects(oCompTag)
'			'msgbox oCompTagObj.count
'			wait 1
'			
'			If intReportCnt <= oCompTagObj.count Then
'				
'				'search for ComponentDataSources values inside <ComponentDataSources> Tag 	
'				strCompDataSrc = oCompTagObj(intReportCnt-1).GetROProperty("innertext")
'				If UBound(aCompDataSrcKeyVal) = UBound(aCompDataSrcCompVal) Then
'					For j = 0 To UBound(aCompDataSrcCompVal) Step 1
'						valCompSearch = instr(1, Trim(Ucase(strCompDataSrc)), Trim(Ucase(aCompDataSrcKeyVal(j)&"="&chr(34)&aCompDataSrcCompVal(j))))			
'						If valCompSearch > 0 Then
'							Call ReportStep (StatusTypes.Pass, aCompDataSrcKeyVal(j)&"="&chr(34)&aCompDataSrcCompVal(j)&" found successfully inside Component Tag", "SCA Exported XML File")
'						Else 
'							Call ReportStep (StatusTypes.Fail, aCompDataSrcKeyVal(j)&"="&chr(34)&aCompDataSrcCompVal(j)&" is not found inside Component Tag", "SCA Exported XML File")
'						End If
'					Next
'				Else
'						Call ReportStep (StatusTypes.Fail, "Could not test values in Components tag successfully. Input params mismatch", "SCA Exported XML File")
'				End If
'			
'			End If
'			
'		Else
'			Call ReportStep (StatusTypes.Information, "Skipping Validation of test values in Components tag", "SCA Exported XML File")
'		End If
'		
'		If intDataVal = 1 Then
'			
'			If UBound(aDataSrcKeyVal) <> -1 or UBound(aDataSrcCompVal) <> -1 Then
'				
'				Set oDataSourceTag  = Description.Create()
'				oDataSourceTag("micclass").value = "WebElement"
'				'oDataSourceTag("html tag").value = "DataSource"
'				oDataSourceTag("html tag").value = "a"
'				oDataSourceTag("innertext").regularexpression=True
'				oDataSourceTag("innertext").value="<DataSource.*"
'				oDataSourceTag("outertext").regularexpression=True
'				oDataSourceTag("outertext").value="<DataSource.*"
'				Set oDataSourceTagObj = Browser("OtherReports").Page("XMLExportedReport").ChildObjects(oDataSourceTag)
'				'msgbox oDataSourceTagObj.count
'				wait 1
'				
'				'search for DataSources values inside <DataSources> Tag
'				strDataSrc = oDataSourceTagObj(0).GetROProperty("innertext")
'				If UBound(aDataSrcKeyVal) = UBound(aDataSrcCompVal) Then
'					For m = 0 To UBound(aDataSrcCompVal) Step 1
'						valDataSearch = instr(1, Trim(Ucase(strDataSrc)), Trim(Ucase(aDataSrcKeyVal(m)&"="&chr(34)&aDataSrcCompVal(m))))			
'						If valDataSearch > 0 Then
'							Call ReportStep (StatusTypes.Pass, aDataSrcKeyVal(m)&"="&chr(34)&aDataSrcCompVal(m)&" found successfully inside DataSource Tag", "SCA Exported XML File")
'						Else 
'							Call ReportStep (StatusTypes.Fail, aDataSrcKeyVal(m)&"="&chr(34)&aDataSrcCompVal(m)&" is not found inside DataSource Tag", "SCA Exported XML File")
'						End If
'					Next
'				Else
'						Call ReportStep (StatusTypes.Fail, "Could not test values in DataSource tag successfully. Input params mismatch", "SCA Exported XML File")
'				End If
'				
'			Else
'				Call ReportStep (StatusTypes.Information, "Skipping Validation of test values in DataSource tag", "SCA Exported XML File")
'			End If
'			
'		End If
'		
'		'Close the xml opened file
'		If intCloseXML = 1 Then
'		'TODO : Get Runtime Browser("CreationTime") and close it
'			For r = 1 To 10 Step 1
'				If Browser("CreationTime:=1").Exist(2) Then
'					 Browser("CreationTime:=1").Close
'					 Call ReportStep (StatusTypes.Pass, "Successfully closed XML opened in IE browser", "XML Exported Page")
'					 Exit For
'				End If 
'				
'				If r = 10 Then
'					Call ReportStep (StatusTypes.Fail, "Could not close XML opened in IE browser successfully", "XML Exported Page")
'				End If
'			Next
'		End If
'		
'	End Function
'	
'	
	
    Private Sub Class_Initialize()
		
	End Sub
	
	
End Class
	
	
