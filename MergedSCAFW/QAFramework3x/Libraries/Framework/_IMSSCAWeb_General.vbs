'	Script:			_IMSSCA_General
'	Author:			Shilpi, Shobha, Shweta
'	Solution:		IMSSCA
'	Project:		IMSSCA
'	Date Created:	Thursday 23 April, 2015
'
'	Description:
'		It contains user defined functions/sub-procedures created to validate 
'		application specific functionalities that can be re-used across 
'		different application components.
'''
'
'
'
'-------------------------------------------------------------------------------

Option Explicit

''' <summary>
''' Creates and Returns an object to CGeneralLibrary
''' </summary>
''' <returns type="IMSSCAGeneralLibrary">The newly created IMSSCAGeneralLibrary object</returns>
''' <browsable>never</browsable>
Public Function [New CIMSSCAGeneralLibrary]()
	
	Dim objGeneral : Set objGeneral = New CIMSSCAGeneralLibrary	
	Set  [New CIMSSCAGeneralLibrary] = objGeneral
	
End Function

''' <summary>
''' Contains Return values of Projects Table return status
''' </summary>
Public ReturnStatus

Set ReturnStatus = New CReturnStatus

''' <summary>
''' Contains Return values after CReturnTypes I.E. Pass|Fail.
''' </summary>
Class CReturnStatus
	
	''' <summary>
	''' Return Type Pass
	''' </summary>
	''' <value type="String"/>
	Public Property Get Pass
		Pass = "Pass"
	End Property
	
	''' <summary>
	''' Return Type Fail
	''' </summary>
	''' <value type="String"/>
	Public Property Get Fail
		Fail = "Fail"
	End Property
	
	''' <summary>
	''' Return Type Fail
	''' </summary>
	''' <value type="String"/>
	Public Property Get Warn
		Warn = "Warn"
	End Property
	
	
End Class


Class CIMSSCAGeneralLibrary
	

	'<Operation on the different Folder Structure>
	'<strOselection: to select the operation>
    '<strFolderName : Foldername>
	'<	strLocation: location to perform the operations>
	'< strFolderName_Update:Folder name to update the Folder name>	
	'<Author>Shobha<Author>
	Public Function OperationOn_Folder(ByVal strOselection,ByVal strFolderName,ByVal strLocation,ByVal strPageName,ByVal strFolderName_Update)
      '<Variable declaration >
	Dim objSharedLink,objMyReport,objFolder,objFolderCount
	'<Looping variables>
	 Dim i,strAppFolder,objCreate,objtxtCreateFolder,objtxtDescription,objbtnCreateFolderOk,objbtnDeleteOK
	 '<String variables>
		 wait 2
		 Browser("Analyzer").Page("Shared_Folder").Sync
		 Browser("Analyzer").Page("Shared_Folder").RefreshObject
		
		'<Shweta - 2/2/2015> Added Sync and waitproprtty stmts - start		
'		   If strLocation="Shared" Then
'					
'		   Set objSharedLink=Browser("Analyzer").Page("Home").Frame("Frame").Link("lnkSharedReports")
'		   Call SCA.ClickOn(objSharedLink,"SharedReport", "Home")	 
'            ElseIf strLocation="MyFolder" Then
'		   Set objMyReportlink=Browser("Analyzer").Page("Home").Frame("Frame").Link("lnkMy Reports") 
'		   Call SCA.ClickOn(objMyReportlink,"MyReport", "Home")	 
'					  
'		  End If

		Set objSharedLink=Browser("Analyzer").Page("Home").Frame("Frame").Link("lnkSharedReports")
		Set objMyReportlink=Browser("Analyzer").Page("Home").Frame("Frame").Link("lnkMy Reports") 
		'180000
	    
        	'shweta 17/3/2016 > Testing While Loop- Start 
        	'For a = Environment.Value("intCounterStart")  To Environment.Value("intCounterMaxLimit") Step 1
        If strLocation <> "" Then
	        For a = 1  To 5 Step 1
					        
				   If strLocation="Shared" Then
					  
					   If objSharedLink.Exist(30) Then
					    objSharedLink.WaitProperty "visible", true, 10000
					   		Call ReportStep(StatusTypes.Pass, "Successfully click on 'Shared Reports' Page in SCA", "SCA Home Page")
					   		Call SCA.ClickOn(objSharedLink,"SharedReport", "Home")
					   		Exit for
					   End If
					   	 
		           ElseIf strLocation="MyFolder" Then
					   
					   If objMyReportlink.Exist(30) Then
					   objMyReportlink.WaitProperty "visible", true, 10000
					   		Call ReportStep(StatusTypes.Pass, "Successfully click on 'MyReport Reports' Page in SCA", "SCA Home Page")
					   		Call SCA.ClickOn(objMyReportlink,"MyReport", "Home")	 
					   		Exit for
					   End If
					   
					End If
					
					'If a=Environment.Value("intCounterMaxLimit") Then
					If a=5 Then
		                Call ReportStep(StatusTypes.Fail, "Could not successfully click on "&strLocation&" Page in SCA", "SCA Home Page")
		                'it should be warning down here
		                Call ReportStep(StatusTypes.Warning, "Could not successfully click on "&strLocation&" Page in SCA. Further functionality Validation steps may fail", "SCA Home Page")
		            	'msgbox a
		            End If
		
				Next
        End If

'			'Commented on 22/3/2016 after testing while loop - start
'			If strLocation="Shared" Then
'					Browser("Analyzer").Page("Shared_Folder").Sync
'					Browser("Analyzer").Page("Shared_Folder").RefreshObject
'				    Do
'						objSharedLink.WaitProperty  "visible", true
'						'print "waiting"
'					Loop While objSharedLink.Exist(2) = False
'					
'					Call ReportStep(StatusTypes.Pass, "Successfully click on 'Shared Reports' Page in SCA", "SCA Home Page")
'					Call SCA.ClickOn(objSharedLink,"SharedReport", "Home")
'					Browser("Analyzer").Page("Shared_Folder").Sync
'					Browser("Analyzer").Page("Shared_Folder").RefreshObject
'					
'	        ElseIf strLocation="MyFolder" Then
'				   
'				   objMyReportlink.WaitProperty "visible", true, 10000
'				   If objMyReportlink.Exist(1) Then
'				   		Call ReportStep(StatusTypes.Pass, "Successfully click on 'MyReport Reports' Page in SCA", "SCA Home Page")
'				   		Call SCA.ClickOn(objMyReportlink,"MyReport", "Home")	 
'				   End If
'				   
'			End If
'			'Commented on 22/3/2016 after testing while loop - End

        	'shweta 17/3/2016 > Testing While Loop- End

          
		  '<Shweta - 2/2/2015> Added Sync and waitproprtty stmts - End


			  Select Case strOselection

					  Case "fSearch"
						  wait  2	
						  
						  '<shweta 12/4/2015> Verify the existence of create New report icon to handle web page synchronization - start
				          'Browser("Analyzer").Page("Shared_Folder").Frame("view").Image("imgnewreport")
				          If Browser("Analyzer").Page("Shared_Folder").Frame("view").Image("imgnewreport").Exist(180) Then
'				           	  Set objFolder=Description.Create
'							  objFolder("Micclass").Value="Link"
'							   objFolder("innertext").Value=strFolderName
'							  
'							  Set objFolderCount=Browser("Analyzer").Page("Shared_Folder").Frame("view").ChildObjects(objFolder)
'							  msgbox  objFolderCount.count
							  
							  if  Browser("Analyzer").Page("Shared_Folder").Frame("view").WebElement("innertext:="&strFolderName,"html tag:=A").Exist(120) Then 
                                 Call ReportStep (StatusTypes.Pass, +strFolderName+"Folder Structure in the Shared\My Report", strPageName)
                                  Browser("Analyzer").Page("Shared_Folder").Frame("view").WebElement("innertext:="&strFolderName,"html tag:=A").click 
				    	
						     End if 
				
'									   For i=0 to objFolderCount.count-1
'												 strAppFolder=objFolderCount(i).getroproperty("innertext")
'												 If strcomp(TRIM(strAppFolder),TRIM(strFolderName))=0 Then												 
'													 objFolderCount(i).Click
'													 Exit For							 
'													 
'												 End If						
'										 Next		  
									'wait 2
'								    If strcomp(TRIM(strAppFolder),TRIM(strFolderName))=0 Then
'										 Call ReportStep (StatusTypes.Pass, +strFolderName+"Folder Structure in the Shared\My Report", strPageName)	
'										 Else
'										 Call ReportStep (StatusTypes.Fail, +strFolderName+"Folder Structure in the Shared\My Report is not present", strPageName)
'									End If 
						   Else
						  		Call ReportStep (StatusTypes.Fail, "SCA Web page synchronization taken long time to load and redirect to "&strFolderName&" Folder Structure in the Shared\My Report", strPageName)
				           End If    
				           
						  '<shweta 12/4/2015>Verify the existence of create New report icon to handle web page synchronization - end...
						  
						   
					   Case "fCreate"

							 Set objCreate=Browser("Analyzer").Page("Shared_Folder").Frame("view").Image("imgnewfolder")
                             Call SCA.ClickOn(objCreate,"newFoldercreation","Report Creation")
                             

                             Set objtxtCreateFolder=Browser("Analyzer").Page("Shared_Folder").Frame("FolderOperation").WebEdit("txtCreateFolder")
                             Call SCA.SetText(objtxtCreateFolder,strFolderName, "Folder Name", "Create Folder in Shared Report")
                              

                             Set objtxtDescription=Browser("Analyzer").Page("Shared_Folder").Frame("FolderOperation").WebEdit("txtDescriptionFolder")
                             Call SCA.SetText(objtxtDescription,"Automation Report Create" , "Folder Description", "Create Folder in Shared Report")
                             
                             Set objbtnCreateFolderOk=Browser("Analyzer").Page("Shared_Folder").Frame("FolderOperation").WebButton("btnCreateFolderOK")
                             Call SCA.ClickOn(objbtnCreateFolderOk,"Folder Create Button","Create Folder in Shared Report")
                             
                             Set objFolder=Description.Create
						     objFolder("Micclass").Value="Link"
						     Set objFolderCount=Browser("Analyzer").Page("Shared_Folder").Frame("view").ChildObjects(objFolder)
			
								   For i=0 to objFolderCount.count-1
								   
											 strAppFolder=objFolderCount(i).getroproperty("innertext")											 
											 If strcomp(TRIM(strAppFolder),TRIM(strFolderName))=0 Then	
											 
												 Exit For							 
												 
											 End If						
								  Next		  
                            
							    If strcomp(TRIM(strAppFolder),TRIM(strFolderName))=0 Then
							
									 Call ReportStep (StatusTypes.Pass, +strFolderName+"Folder Structure is Created Successfully in the Shared\My Report", strPageName)	
									 Else
									 Call ReportStep (StatusTypes.Fail, +strFolderName+"Folder Structure is not Created Successfully in the Shared\My Report", strPageName)
									 
							   End If 


					  Case "fDelete"
					  		
					  		Set objFolder=Description.Create
						    objFolder("Micclass").Value="WebElement"
						    objFolder("innertext").Value=strFolderName
						  
						     Set objFolderCount=Browser("Analyzer").Page("Shared_Folder").Frame("view").ChildObjects(objFolder)
			
								   For i=0 to objFolderCount.count-1
											 strAppFolder=objFolderCount(i).getroproperty("innertext")
											 If strcomp(TRIM(strAppFolder),TRIM(strFolderName))=0 Then												 
												 'objFolderCount(i).Click
												 'Setting.WebPackage("ReplayType") = 2
												 objFolderCount(i).RightCLick
												 'Setting.WebPackage("ReplayType") = 1
												 Exit For							 
												 
											 End If						
									 Next		  
					  			
					  			CAll ReportFunctions("Delete",strFolderName,"Folder CReation in Shared Reports","")
					  			
'					  			Set objbtnDeleteOK=Browser("Analyzer").Dialog("Message from webpage").WinButton("btnOK")
'                                Call SCA.ClickOn(objbtnDeleteOK,"weblink button", "Dialogue")
					  			
					  			Set objFolderCount=Browser("Analyzer").Page("Shared_Folder").Frame("view").ChildObjects(objFolder)								  	  
					  			
					  			If objFolderCount.count=0 Then
							
									 Call ReportStep (StatusTypes.Pass, +strFolderName+"Folder Structure in the Shared\My Report", strPageName)	
									 Else
									 Call ReportStep (StatusTypes.Fail, +strFolderName+"Folder Structure in the Shared\My Report is not present", strPageName)
									 
							   End If 
					     
					  Case "fUpdate"
					  
					        Set objFolder=Description.Create
						    objFolder("Micclass").Value="WebElement"
						    objFolder("innertext").Value=strFolderName
						  
						     Set objFolderCount=Browser("Analyzer").Page("Shared_Folder").Frame("view").ChildObjects(objFolder)
			
								   For i=0 to objFolderCount.count-1
											 strAppFolder=objFolderCount(i).getroproperty("innertext")
											 If strcomp(TRIM(strAppFolder),TRIM(strFolderName))=0 Then												 
												 ' Setting.WebPackage("ReplayType") = 2
												  objFolderCount(i).RightCLick
												  'Setting.WebPackage("ReplayType") = 1
												 Exit For							 
												 
											 End If						
									 Next		  
					  			
					  			CAll ReportFunctions("Config",strFolderName,"Folder CReation in Shared Reports","")
					  			
					  		  Set objtxtCreateFolder=Browser("Analyzer").Page("Shared_Folder").Frame("FolderOperation").WebEdit("txtCreateFolder")
                              Call SCA.SetText(objtxtCreateFolder,strFolderName_Update, "Folder Name", "Create Folder in Shared Report")
                               
                             Set objbtnCreateFolderOk=Browser("Analyzer").Page("Shared_Folder").Frame("FolderOperation").WebButton("btnCreateFolderOK")
                             Call SCA.ClickOn(objbtnCreateFolderOk,"Folder Create Button","Create Folder in Shared Report")
                             
                             Wait 3
                             
                             Set objFolder=Description.Create
						    objFolder("Micclass").Value="WebElement"
						    objFolder("innertext").Value=strFolderName_Update
						  
						     Set objFolderCount=Browser("Analyzer").Page("Shared_Folder").Frame("view").ChildObjects(objFolder)
			
								   For i=0 to objFolderCount.count-1
											 strAppFolder=objFolderCount(i).getroproperty("innertext")
											 If strcomp(TRIM(strAppFolder),TRIM(strFolderName_Update))=0 Then												 
												  Exit For										 
											 End If						
								   Next	
                             If strcomp(TRIM(strAppFolder),TRIM(strFolderName_Update))=0 Then
                             
		                         Call ReportStep (StatusTypes.Pass, +strFolderName+Space(2)+"Folder is updated to"+Space(2)&strFolderName_Update, strPageName)	
									 Else
								 Call ReportStep (StatusTypes.Fail, +strFolderName+Space(2)+"Folder is not updated to"+Space(2)&strFolderName_Update, strPageName)										 										 
							 End If	                                
					  					
					  
					  Case "fVerification"
                      
                      Set objFolder=Description.Create
                          objFolder("Micclass").Value="Link"
                          Set objFolderCount=Browser("Analyzer").Page("Shared_Folder").Frame("view").ChildObjects(objFolder)
            
                                   For i=0 to objFolderCount.count-1
                                             strAppFolder=objFolderCount(i).getroproperty("innertext")
                                             If strcomp(TRIM(strAppFolder),TRIM(strFolderName))=0 Then                             
                                                Exit For    
                                             End If                        
                                     Next 

						 If strcomp(TRIM(strAppFolder),TRIM(strFolderName))=0 Then                                                 
                              BfolderExist=0    
                              else
                              BfolderExist=1                                                  
                         End If	
                         
                        OperationOn_Folder=BfolderExist
                        
                	Case "foldercount"
	                      fcount=0
	                      Set objFolder=Description.Create
	                          objFolder("Micclass").Value="Link"
	                          Set objFolderCount=Browser("Analyzer").Page("Shared_Folder").Frame("view").ChildObjects(objFolder)
	            
	                                   For i=0 to objFolderCount.count-1
	                                             strAppFolder=objFolderCount(i).getroproperty("innertext")
	                                             If strcomp(TRIM(strAppFolder),TRIM(strFolderName))=0 Then                             
	                                                 fcount=fcount+1   
	                                             End If                        
	                                     Next 
	                         
	                        OperationOn_Folder=fcount
                        
				  Case "fRImage_Verification"
					  wait 2
					  Set objFolder=Description.Create
						  objFolder("Micclass").Value="WebTable"
						  objFolder("outertext").RegularExpression=True
						  objFolder("outertext").Value=strFolderName&".*"
					
						  
						  Set objFolderCount=Browser("Analyzer").Page("Shared_Folder").Frame("view").ChildObjects(objFolder)
							For i=0 to objFolderCount.count-1
							  strRName=objFolderCount(i).getroproperty("name")
							  strRType=strRType&Space(1)&strRName
											 					
						   Next	  
						OperationOn_Folder=strRType			
			  End Select


	End Function	
	
	
	'<Selects SCA components from ReportAreaPermierSettings>
    '<strSettingsLoc:  ReportAreaPermierSettings can be selected from "NewSheet", "VerticalDownNormal", "HorizontalDownNormal" :  String>
    '<strOlapMenuValue: Select component options under each tab :  String>
    Public Sub ReportAreaPermierSettings(ByVal strSettingsLoc, ByVal strOlapMenuValue)
	
		Dim strOlapMenuCreated, RowCnt, ColCnt, i, j, PivotTableSheetVal, oDescOlap, OlapObj, objNewSheet, objVerticalCntxtMenu, objHorizontalCntxtMenu
		
		strOlapMenuCreated = 0
		
		Browser("Analyzer").Page("ReportCreation").Sync
		Browser("Analyzer").Page("ReportCreation").RefreshObject
		
		wait 5
		Select Case strSettingsLoc
			Case "NewSheet"
					Set objNewSheet=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgNewSheet")
	        		Call SCA.ClickOn(objNewSheet,"imgNewSheet","Report Creation")
	        		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgNewSheet").FireEvent "oncontextmenu"
					
        	Case "welVerticalDownNormal"
        			Set objVerticalCntxtMenu=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welVerticalDownNormal")
	        		Call SCA.ClickOn(objVerticalCntxtMenu,"welVerticalDownNormal","Report Creation")
	        		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welVerticalDownNormal").FireEvent "oncontextmenu"
	        		
	        Case "welHorizontalDownNormal"
        			Set objHorizontalCntxtMenu=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welHorizontalDownNormal")
	        		Call SCA.ClickOn(objHorizontalCntxtMenu,"welHorizontalDownNormal","Report Creation")
	        		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welHorizontalDownNormal").FireEvent "oncontextmenu"		
	        
        	
		End Select
		
		wait 10
		RowCnt= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabReportAreaPivotTable").GetROProperty("rows")
		ColCnt= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabReportAreaPivotTable").GetROProperty("cols")
		
		For i = 1 To RowCnt Step 1
		
			For j = 1 To ColCnt
			
				PivotTableSheetVal = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabReportAreaPivotTable").GetCellData(i, j)
				If UCase(Trim(PivotTableSheetVal)) = UCase(Trim(strOlapMenuValue)) Then
						' Shobha changes		
						Set oDescOlap  = Description.Create()
						oDescOlap("micclass").value = "WebElement"
						oDescOlap("innertext").value=strOlapMenuValue
'						oDescOlap("innerhtml").regularexpression=True
'						oDescOlap("innerhtml").value=".*nbsp"
						oDescOlap("html tag").value="TD"
					        wait 1						
						Set OlapObj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabReportAreaPivotTable").ChildObjects(oDescOlap)
						wait 1
						OlapObj(0).Click
						strOlapMenuCreated = 1
						Exit For
				End If
				
			Next
			
			If strOlapMenuCreated = 1 Then
				'Reporter Statement
				Call ReportStep (StatusTypes.Pass, "Successfully created " +strOlapMenuValue+ " from ReportAreaPermierSettings", "ReportCreation Page")
				Exit For
			End If
			
		Next
		
		If strOlapMenuCreated = 0 Then
			Call ReportStep (StatusTypes.Fail, "Could not create " +strOlapMenuValue+ " from ReportAreaPermierSettings", "ReportCreation Page")
			Exit Sub
		End If
   		
		Browser("Analyzer").Page("Shared_Folder").Sync
		Browser("Analyzer").Page("Shared_Folder").RefreshObject
		 
	End Sub

	
	'<To create the Report in SCA using the Pivotal table>
    '<sheetNum :- Name of the sheet>
    '<newReport :- to create the new Report or Not>
    '<author :-Shobha>    
    Public Function ReportCreationInSCA(ByVal sheetNum, ByVal newReport,ByVal intEMD,ByVal intdbcheck,ByVal intRNum,ByVal intReportCreate,ByVal strF1,ByVal strF2)  
    
        Dim objExcel,ObjExcelFile,ObjExcelSheet,objDataSource,objDatabase,objDatacubes,obj_tree,objDR,objrow,objcolumn,objDataaxis,objAdd,objtreeCount,objCreate
        Dim intRow_Count,intColumn_Count,intcount
        Dim dataSourceValue, strDataSourceValue,strDatabaseValue,strCubeValue,strChild_Values,strPlaceHolderVal,p, objBtnok, objErrorOk,strDBitems,strbkch
        '<Looping Variables for the row and column>
        Dim i,j,k,x,y,a  
        
        '<shweta 11/9/2015 Closing All opened Excel Files- start
		'objUtils.KillProcess("excel.exe")
		Systemutil.CloseProcessByName "excel.exe"
		'<shweta 11/9/2015 Closing All opened Excel Files- end
        
        Set objExcel = CreateObject("Excel.Application") 
        objExcel.Visible =False     
    
       ' Set ObjExcelFile = objExcel.Workbooks.Open(Environment.Value("CurrDir")&"InputFiles\IMSSCAWeb\SCAReportSheet.xls")
        Set ObjExcelFile = objExcel.Workbooks.Open(Environment.Value("CurrDir")&Environment.Value("ReportCreationFile"))
        Set ObjExcelSheet = ObjExcelFile.Sheets(sheetNum) 
        intRow_Count = ObjExcelSheet.UsedRange.Rows.Count 
        intColumn_Count = ObjExcelSheet.UsedRange.Columns.Count
    
        dataSourceValue=TRIM(ObjExcelSheet.cells(2,1).value)
        
        '<Shweta - 16/12/2016> handled change in DataSource Name for "+ CDTSOLAP573I*", "CDTSOLAP573I*", "CDTSOLAP573I.GEMINI.DEV*" format - Start 
        'strDataSourceValue = "+ "&dataSourceValue       
        '<Shweta - 16/12/2016> handled change in DataSource Name for "+ CDTSOLAP573I*", "CDTSOLAP573I*", "CDTSOLAP573I.GEMINI.DEV*" format - End
        
        strDatabaseValue=TRIM(ObjExcelSheet.cells(2,2).value)
        strCubeValue=TRIM(ObjExcelSheet.cells(2,3).value)
        
        If intEMD=1 Then
        
         If newReport = 0 Then
            Set objCreate=Browser("Analyzer").Page("Shared_Folder").Frame("view").Image("imgnewreport")        
           'Set objCreate=Browser("Analyzer").Page("Shared_Folder").Image("imgnewreport")-------Ask Shobha
           Call SCA.ClickOn(objCreate,"newreport","Report Creation")    
        End If
              
       'Capture Database Access Error - Start
       If Browser("Analyzer").Frame("ErrorFrame").WebElement("Database access error.").Exist(1) Then
               Set objErrorOk=Browser("Analyzer").Frame("ErrorFrame").WebButton("btnOK")
               Call SCA.ClickOn(objErrorOk,"btnOK" , "Database Access Error during Report Creation")
               Call ReportStep (StatusTypes.Fail,"Not Able to create a report. Database Access Error during Report Creation", "Report Creation Page")    
                    
            'Closing Excel
            ObjExcelFile.Close
            objExcel.Quit    
            Set objExcel=nothing
            Set  ObjExcelFile=nothing
            Set  ObjExcelSheet=nothing
            Set  objtreeCount=nothing
            ReportCreationInSCA = 1
            Exit Function
       End If
       'Capture Database Access Error - End
       
       'Handled Application Time Out in Report Creation Page - Start
       Set objBtnok = Browser("Analyzer").Dialog("Message from webpage").WinButton("btnOK")
       If objBtnok.Exist(5) Then
            Call SCA.ClickOn(objBtnok,"btnOK" , "Application Time Out in Report Creation Page")
       End If    
       'Handled Application Time Out in Report Creation Page - End
       wait 10
       For a = 1 To 100 Step 1           
       If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDataSources").GetROProperty("disable")=0  Then
         Exit For
       End If           
       Next
       
       wait 2
       
       '<Shweta - 16/12/2016> handled change in DataSource Name for "+ CDTSOLAP573I*", "CDTSOLAP573I*", "CDTSOLAP573I.GEMINI.DEV*" format - Start 
	       'Set objDataSource=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDataSources")   
	       'Call SCA.SelectFromDropdown(objDataSource,strDataSourceValue)
	       
	        Set reg = createobject("vbscript.Regexp")
			reg.global = true
			reg.pattern = "[^\d]"
			DTServerVal = reg.replace(dataSourceValue, "")
			
			strAllItems = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstDdlDataSources").GetROProperty("all items")
			arrAllItems = Split(strAllItems, ";")
			For i = 0 To ubound(arrAllItems) Step 1
				If instr(1, arrAllItems(i), DTServerVal) > 1 Then
					'print arrAllItems(i)
					selectItem = i
					'print selectItem
					Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstDdlDataSources").Select selectItem
				End If
			Next
	
       '<Shweta - 16/12/2016> handled change in DataSource Name for "+ CDTSOLAP573I*", "CDTSOLAP573I*", "CDTSOLAP573I.GEMINI.DEV*" format - Start 
       
       wait 2
       
       Set objDatabase=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDatabases")
       Call SCA.SelectFromDropdown(objDatabase,strDatabaseValue)       
       
       If intdbcheck=0 Then   
       
        strDBitems=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDatabases").GetROProperty("all items")
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
       
       
       Set objDatacubes=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlCubes")
       Call SCA.SelectFromDropdown(objDatacubes,strCubeValue)
       
       Set objAdd=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebButton("btnAdd")
       Call SCA.ClickOn(objAdd,"AddButton" , "ReportCreation")
    
    	'<Shweta 22/6/2016 > - Start
    	strPrev_Child_Values = ""
    	'<Shweta 22/6/2016 > - End
    	
       For i=2 to intRow_Count
        Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcube").Click           
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
                        set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
                        intcount=objtreeCount.count
                        
                        For a = 1 To 100 Step 1
	                    set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
	                    intcount=objtreeCount.count
                            If intcount<>0 Then
                            Exit For
                            End If
                        Next
                        
                        If intcount<>0 Then
                        	'shweta - 31/3/2016 - start
                        	k=intcount-1
							'print "K Value: "&k&" for strChild_Values: "&strChild_Values                        	
                            objtreeCount(k).fireEvent "ondblclick" 
                            'objtreeCount(0).fireEvent  "ondblclick"  
							'shweta - 31/3/2016 - End                            
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
'                  '<shweta 22/6/2016>-checking trvSchemaSel SelectedNode count before clicking on Row/Column/Data Axis - Start  
                	Set obj_tree=description.Create
                    obj_tree("micclass").value="WebElement"
                    obj_tree("html tag").value="SPAN"
                    obj_tree("innertext").value=strPrev_Child_Values
                    wait 1
                    set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
		        	intcount=objtreeCount.count
		        	If intcount<>0 Then
                    	k=intcount-1
                    Else
                    	Call ReportStep (StatusTypes.Fail, "Couldn not select node'"&strPrev_Child_Values&" during report creation" , "Report Creation Page")
                    	Exitrun
                    End If	
                    
'                    '<shweta 8/6/2016>-checking tree node count before clicking on Row/Column/Data Axis - Start
'            		set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
'		        	intcount=objtreeCount.count
'		        	newK=intcount-1
'		        	If newK <> -1 Then
'		        		k = newK
'		        	End If
''		        	'<shweta 8/6/2016>-checking tree node count before clicking on Row/Column/Data Axis - End   

'                  '<shweta 22/6/2016>-checking trvSchemaSel SelectedNode count before clicking on Row/Column/Data Axis - End

                    wait 1
                    '<shweta 8/6/2016> - Adding Fire Event to highlight the treenode to be choosed as row/column/data axis - Start
                    objtreeCount(k).FireEvent "onmouseover"
                    wait 2
                    objtreeCount(k).highlight
                    wait 1
                   ' objtreeCount(k).rightclick 
                   'changes by srinivas
                    Setting.WebPackage("ReplayType") = 2
					objtreeCount(k).RightClick
					Setting.WebPackage("ReplayType") = 1
                    '<shweta 8/6/2016> - Adding Fire Event to highlight the treenode to be choosed as row/column/data axis - End
                    
                    '<shweta 8/6/2016> - Commented beacuse Devicereplay is inconsistent. So added "objtreeCount(k).rightclick"  in above step - Start
					'   Set objDR=CreateObject("Mercury.DeviceReplay")
					'   x=objtreeCount(k).GetroProperty("abs_x")
					'   y=objtreeCount(k).GetroProperty("abs_y")  
					'   wait 1                    
					'   objDR.MouseClick x,y,2
                    '<shweta 8/6/2016> - Commented beacuse Devicereplay is inconsistent. So added "objtreeCount(k).rightclick"  in above step - End
                    strPlaceHolderVal=lcase(ObjExcelSheet.cells(i,intColumn_Count-1).value)                                                                                        
                    wait 3  
                      If strPlaceHolderVal="rowaxis" Then 
                      	  	
                          Set objrow=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webRowaxis")    
                          Call SCA.ClickOn(objrow,"Rowaxis" , "ReportCreation")    
                          wait 5                          
                      elseif  StrPlaceHolderVal="columnaxis" then  
						                       
                          Set objcolumn=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcolumnaxis")
                          Call SCA.ClickOn(objcolumn,"Columnaxis" , "ReportCreation")
                          Wait 5                      
                      elseif  StrPlaceHolderVal="dataaxis" then
						  
						  If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webDataAxis").Exist(10)  Then
						  	Set objDataaxis=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webDataAxis")
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
                            set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
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
           
'           Set objExcel = CreateObject("Excel.Application") 
'        objExcel.Visible =False     
'        
'        Set ObjExcelFile = objExcel.Workbooks.Open(Environment.Value("CurrDir")&"InputFiles\IMSSCAWeb\SCAReportSheet.xls")
'        Set ObjExcelSheet = ObjExcelFile.Sheets(intsheetName)
'        
'        intRow_Count = ObjExcelSheet.UsedRange.Rows.Count 
'        intColumn_Count = ObjExcelSheet.UsedRange.Columns.Count   
        Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcube").Click          
        For j=4 to intColumn_Count-2
        strChild_Values=Trim(ObjExcelSheet.cells(intRNum,j).value)
               
          If  strChild_Values<>"" AND  strChild_Values<>"rowaxis" AND strChild_Values<>"columnaxis"  AND strChild_Values<>"dataaxis" AND strChild_Values<>"Yes" AND strChild_Values<>"No" Then
        
               Set obj_tree=description.Create
               obj_tree("micclass").value="WebElement"
               obj_tree("class").value="TreeNode"
               obj_tree("innertext").value=strChild_Values
               wait 1
               set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
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
                'objDR.MouseClick x,y,2
                Setting.WebPackage("ReplayType") = 2
                objtreeCount(k).rightclick
                Setting.WebPackage("ReplayType") = 1
                strPlaceHolderVal=lcase(ObjExcelSheet.cells(intRNum,intColumn_Count-1).value)                                                                                        
                      
                      If strPlaceHolderVal="rowaxis" Then 
                          Set objrow=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webRowaxis")    
                          Call SCA.ClickOn(objrow,"Rowaxis" , "ReportCreation")    
                          wait 5                          
                      elseif  StrPlaceHolderVal="columnaxis" then                                                                                                                                
                          Set objcolumn=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcolumnaxis")
                          Call SCA.ClickOn(objcolumn,"Columnaxis" , "ReportCreation")
                          Wait 5                      
                      elseif  StrPlaceHolderVal="dataaxis" then                                                                            
                         Set objDataaxis=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webDataAxis")
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
                            set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
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

	
'	
'	<To perform the DRag and drop operation while group creation>
'	<intgroupcreation: to select either the group is created or not>
'	<sheetNum: Selection of the test data from the excel sheet>
'	<newReport: selection for the new Report or for the existance Report>
'	<intRow_Count: set the row number to expand the tree>
'	<strF and strF1: to do the modification in future>
'   <author:shobha>
	Public Sub GroupTableCreationInSCAOld(ByVal intgroupcreation,ByVal sheetNum, ByVal newReport, ByVal intRow_Count, ByVal strF,Byval strF1)  
    
        Dim objExcel,ObjExcelFile,ObjExcelSheet,objDataSource,objDatabase,objDatacubes,obj_tree,objDR,objrow,objcolumn,objDataaxis,objAdd,objtreeCount,objCreate
        Dim intColumn_Count,intcount
        Dim dataSourceValue, strDataSourceValue,strDatabaseValue,strCubeValue,strChild_Values,strPlaceHolderVal,p, objBtnok, objErrorOk
        '<Looping Variables for the row and column>
        Dim i,j,k,x,y,a
        
        '<shweta 16/6/2015 Closing All opened Excel Files- start
		'objUtils.KillProcess("excel.exe")
		Systemutil.CloseProcessByName "excel.exe"
		'<shweta 16/6/2015 Closing All opened Excel Files- end
		
        Set objExcel = CreateObject("Excel.Application") 
        objExcel.Visible =False     
    
        Set ObjExcelFile = objExcel.Workbooks.Open(Environment.Value("CurrDir")&Environment.Value("ReportCreationFile"))
        Set ObjExcelSheet = ObjExcelFile.Sheets(sheetNum) 
        intColumn_Count = ObjExcelSheet.UsedRange.Columns.Count
    
        dataSourceValue=TRIM(ObjExcelSheet.cells(2,1).value)
        strDataSourceValue = "+ "&dataSourceValue
        strDatabaseValue=TRIM(ObjExcelSheet.cells(2,2).value)
        strCubeValue=TRIM(ObjExcelSheet.cells(2,3).value)
        
		If intgroupcreation=0 Then		
    
    	
          
		
		
		If newReport = 0 Then
    		Set objCreate=Browser("Analyzer").Page("Shared_Folder").Frame("view").Image("imgnewreport")      
	        Call SCA.ClickOn(objCreate,"newreport","Report Creation")
	        
	        
	        'Capture Database Access Error - Start
       If Browser("Analyzer").Frame("ErrorFrame").WebElement("Database access error.").Exist(0) Then
       		Set objErrorOk=Browser("Analyzer").Frame("ErrorFrame").WebButton("btnOK")
       		Call SCA.ClickOn(objErrorOk,"btnOK" , "Database Access Error during Report Creation")
       		Call ReportStep (StatusTypes.Fail,"Not Able to create a report. Database Access Error during Report Creation", "Report Creation Page")	
					
			'Closing Excel
			ObjExcelFile.Close
			objExcel.Quit	
			Set objExcel=nothing
			Set  ObjExcelFile=nothing
			Set  ObjExcelSheet=nothing
			Set  objtreeCount=nothing
			Exit Sub
	   End If
       'Capture Database Access Error - End
       
       'Handled Application Time Out in Report Creation Page - Start
	   Set objBtnok = Browser("Analyzer").Dialog("Message from webpage").WinButton("btnOK")
	   If objBtnok.Exist(10) Then
			Call SCA.ClickOn(objBtnok,"btnOK" , "Application Time Out in Report Creation Page")
	   End If	
	   'Handled Application Time Out in Report Creation Page - End  
	        
	        
			Call IMSSCA.General.ReportAreaPermierSettings(TRIM("welVerticalDownNormal"),TRIM("  GroupTable ") )	        
    	End If
		
		
		
             
       	For a = 1 To 100 Step 1
        	If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDataSources").GetROProperty("disable")=0  Then
            Exit For
        	End If
		Next
       
       wait 2
       
       Set objDataSource=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDataSources")   
       Call SCA.SelectFromDropdown(objDataSource,strDataSourceValue)
       
       	For a = 1 To 50 Step 1
        	If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDatabases").GetROProperty("disable")=0  Then
            Exit For
        	End If
		Next
       
       Set objDatabase=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDatabases")
       Call SCA.SelectFromDropdown(objDatabase,strDatabaseValue)
       
       	For a = 1 To 50 Step 1
        	If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlCubes").GetROProperty("disable")=0  Then
            Exit For
        	End If
		Next
       
       Set objDatacubes=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlCubes")
       Call SCA.SelectFromDropdown(objDatacubes,strCubeValue)
       
       Set objAdd=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebButton("btnAdd")
       Call SCA.ClickOn(objAdd,"AddButton" , "ReportCreation")
       
       
       
       
       ' dynamically we have to fetch the row value. So "i" should be dynamic
    
       For i=intRow_Count to intRow_Count
    	  ' Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcube").Click
           Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webCubes").Click

           For j=4 to intColumn_Count-1
    		  
	           strChild_Values=Trim(ObjExcelSheet.cells(i,j).value)
	           If  strChild_Values<>"" AND  strChild_Values<>"rowaxis" AND strChild_Values<>"columnaxis"  AND strChild_Values<>"dataaxis" AND strChild_Values<>"addtofav" AND  strChild_Values<>"deletepattern" Then
	                    Set obj_tree=description.Create
	                    obj_tree("micclass").value="WebElement"
	                    obj_tree("class").value="TreeNode"
	                    obj_tree("innertext").value=strChild_Values
	                    wait 1
	                    set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
	                    intcount=objtreeCount.count
	                    
	                    For a = 1 To 20 Step 1
		                    If intcount<>0 Then
		                    Exit For
		                    End If
	                    Next
	                    
	                    If intcount<>0 Then
	                    If intcount=3 Then
	                    	k=intcount-2                    
		                    objtreeCount(k).fireEvent  "ondblclick"  
							                    
		                    Call ReportStep (StatusTypes.Pass,"Clicked on the Webelement"&Space(2)&strChild_Values, "Report Creation Page")                    
		                    wait 3  
	                    	ElseIf intcount=2 Then
	                    	k=intcount-2                    
		                    objtreeCount(k).fireEvent  "ondblclick"        
		                    Call ReportStep (StatusTypes.Pass,"Clicked on the Webelement"&Space(2)&strChild_Values, "Report Creation Page")                    
		                    wait 3  
	                    	Else
	                    	k=intcount-1                    
		                    objtreeCount(k).fireEvent  "ondblclick"        
		                    Call ReportStep (StatusTypes.Pass,"Clicked on the Webelement"&Space(2)&strChild_Values, "Report Creation Page")                    
		                    wait 3  
	                    End If
		                      
	                    Else
	                    	Call ReportStep (StatusTypes.Warning,"Not Clicked on the Webelement"&Space(2)&strChild_Values, "Report Creation Page")
		                    Exit Sub
	                    End If    
	                    
	           End If
	           
	           If j=intColumn_Count-1 Then
		            wait 1
		            Set objDR=CreateObject("Mercury.DeviceReplay")
		            x=objtreeCount(k).GetroProperty("abs_x")
		            y=objtreeCount(k).GetroProperty("abs_y") 
					'changes by srinivas		            
		            Setting.WebPackage("ReplayType") = 2
		            'objDR.MouseClick x,y,2
		            objtreeCount(k).rightclick
		            Setting.WebPackage("ReplayType") = 1
		            strPlaceHolderVal=lcase(ObjExcelSheet.cells(i,j+1).value)                                                                                        
		              
		              If strPlaceHolderVal="rowaxis" Then 
		                  Set objrow=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webRowaxis")    
		                  Call SCA.ClickOn(objrow,"Rowaxis" , "ReportCreation")    
		                  wait 5                          
		              elseif  StrPlaceHolderVal="columnaxis" then                                                                                                                                
		                  Set objcolumn=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcolumnaxis")
		                  Call SCA.ClickOn(objcolumn,"Columnaxis" , "ReportCreation")
		                  Wait 5                      
		              elseif  StrPlaceHolderVal="dataaxis" then                                                                            
		                 Set objDataaxis=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webDataAxis")
		                 Call SCA.ClickOn(objDataaxis,"Dataaxis" , "ReportCreation")    
		                 wait 5
		              'Shweta - start 
                      elseif  StrPlaceHolderVal="addtofav" then                                                                            
                         Set objAddToFavorite=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welAddToFavorite")
                         Call SCA.ClickOn(objAddToFavorite,"AddToFavorite" , "ReportCreation")    
                         wait 5
                       elseif  StrPlaceHolderVal="deletepattern" then                                                                            
                       	 wait 2
                         Set objDelPattern=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welDeletePattern")
                         Call SCA.ClickOn(objDelPattern,"DeletePattern" , "ReportCreation")    
                         wait 5
                         
                      'Shweta - End
		              End If    
		              
		              
	            
	           End If
             
           Next
    
       Next
       
		
	 
	
    ElseIf intgroupcreation=1 Then         
       
       For p=intColumn_Count-1 to 4  step -1
		  strChild_Values=Trim(ObjExcelSheet.cells(intRow_Count,p).value)
	      If  strChild_Values<>"" AND  strChild_Values<>"rowaxis" AND strChild_Values<>"columnaxis"  AND strChild_Values<>"dataaxis" AND strChild_Values<>"addtofav" AND  strChild_Values<>"deletepattern" Then
		                            
		    Set obj_tree=description.Create
		    obj_tree("micclass").value="WebElement"
		    obj_tree("class").value="TreeNode"
		    obj_tree("innertext").value=strChild_Values
		    wait 1
		    set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
		    intcount=objtreeCount.count    
		       If intcount<>0 Then
		       If intcount=3 Then
		       	k=intcount-3
		       	objtreeCount(k).fireEvent  "ondblclick"  
		       	Else
		       	k=intcount-1
		        objtreeCount(k).fireEvent  "ondblclick"   
		       	
		       End If
		          
		       End If                            
		                            
		End If
	 Next
	 
	 ElseIf intgroupcreation=2 then
	 
	 For p=intColumn_Count-1 to 4  step -1
		  strChild_Values=Trim(ObjExcelSheet.cells(intRow_Count,p).value)
	      If  strChild_Values<>"" AND  strChild_Values<>"rowaxis" AND strChild_Values<>"columnaxis"  AND strChild_Values<>"dataaxis" Then
		                            
		    Set obj_tree=description.Create
		    obj_tree("micclass").value="WebElement"
		    obj_tree("class").value="TreeNode"
		    obj_tree("innertext").value=strChild_Values
		    wait 1		    
		    
		    set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
		    intcount=objtreeCount.count    
		       If intcount<>0 Then
		       If intcount=2 Then
		       	 k=intcount-1
		       	 objtreeCount(k).fireEvent  "ondblclick" 
		       	 else
		       	 k=intcount-1
		       	  objtreeCount(k).fireEvent  "ondblclick" 
		       End If
		        		            
		       End If 


			 Set objDR=CreateObject("Mercury.DeviceReplay")
		            x=objtreeCount(k).GetroProperty("abs_x")
		            y=objtreeCount(k).GetroProperty("abs_y")        
		            'objDR.MouseClick x,y,2
		            'changes by srinivas
		            Setting.WebPackage("ReplayType") = 2
		            objtreeCount(k).rightclick
		            Setting.WebPackage("ReplayType") = 1
		            strPlaceHolderVal=lcase(ObjExcelSheet.cells(intRow_Count,intColumn_Count).value)                                                                                        
		              
		              If strPlaceHolderVal="rowaxis" Then 
		                  Set objrow=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webRowaxis")    
		                  Call SCA.ClickOn(objrow,"Rowaxis" , "ReportCreation")    
		                  wait 5                          
		              elseif  StrPlaceHolderVal="columnaxis" then                                                                                                                                
		                  Set objcolumn=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcolumnaxis")
		                  Call SCA.ClickOn(objcolumn,"Columnaxis" , "ReportCreation")
		                  Wait 5                      
		              elseif  StrPlaceHolderVal="dataaxis" then                                                                            
		                 Set objDataaxis=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webDataAxis")
		                 Call SCA.ClickOn(objDataaxis,"Dataaxis" , "ReportCreation")    
		                 wait 5
		             End If    
		    
		   Exit For                         
		End If
		
	 Next
	 
    
  End If 
  
    ObjExcelFile.Close
	objExcel.Quit    
	Set objExcel=nothing
	Set ObjExcelFile=nothing
	Set ObjExcelSheet=nothing
	Set objtreeCount=nothing	
  
 End Sub


'	<To perform the DRag and drop operation while group creation>
'	<intgroupcreation: to select either the group is created or not>
'	<sheetNum: Selection of the test data from the excel sheet>
'	<newReport: selection for the new Report or for the existance Report>
'	<intRow_Count: set the row number to expand the tree>
'	<strF and strF1: to do the modification in future>
'   <author:shobha>
	Public Sub GroupTableCreationInSCA(ByVal intgroupcreation,ByVal sheetNum, ByVal newReport, ByVal intRow_Count, ByVal strF,Byval strF1)  
        wait 5
        Dim objExcel,ObjExcelFile,ObjExcelSheet,objDataSource,objDatabase,objDatacubes,obj_tree,objDR,objrow,objcolumn,objDataaxis,objAdd,objtreeCount,objCreate
        Dim intColumn_Count,intcount
        Dim dataSourceValue, strDataSourceValue,strDatabaseValue,strCubeValue,strChild_Values,strPlaceHolderVal,p, objBtnok, objErrorOk
        '<Looping Variables for the row and column>
        Dim i,j,k,x,y,a
        
        '<shweta 16/6/2015 Closing All opened Excel Files- start
		'objUtils.KillProcess("excel.exe")
		Systemutil.CloseProcessByName "excel.exe"
		'<shweta 16/6/2015 Closing All opened Excel Files- end
		
        Set objExcel = CreateObject("Excel.Application") 
        objExcel.Visible =False     
    
        Set ObjExcelFile = objExcel.Workbooks.Open(Environment.Value("CurrDir")&Environment.Value("ReportCreationFile"))
        Set ObjExcelSheet = ObjExcelFile.Sheets(sheetNum) 
        intColumn_Count = ObjExcelSheet.UsedRange.Columns.Count
    
        dataSourceValue=TRIM(ObjExcelSheet.cells(2,1).value)
        '<Shweta - 16/12/2016> handled change in DataSource Name for "+ CDTSOLAP573I*", "CDTSOLAP573I*", "CDTSOLAP573I.GEMINI.DEV*" format - Start 
        'strDataSourceValue = "+ "&dataSourceValue
        '<Shweta - 16/12/2016> handled change in DataSource Name for "+ CDTSOLAP573I*", "CDTSOLAP573I*", "CDTSOLAP573I.GEMINI.DEV*" format - End
        strDatabaseValue=TRIM(ObjExcelSheet.cells(2,2).value)
        strCubeValue=TRIM(ObjExcelSheet.cells(2,3).value)
        
		If intgroupcreation=0 Then		
    
    	
          
		
		
		If newReport = 0 Then
    		Set objCreate=Browser("Analyzer").Page("Shared_Folder").Frame("view").Image("imgnewreport")      
	        Call SCA.ClickOn(objCreate,"newreport","Report Creation")
	        
	        
	        'Capture Database Access Error - Start
       If Browser("Analyzer").Frame("ErrorFrame").WebElement("Database access error.").Exist(5) Then
       		Set objErrorOk=Browser("Analyzer").Frame("ErrorFrame").WebButton("btnOK")
       		Call SCA.ClickOn(objErrorOk,"btnOK" , "Database Access Error during Report Creation")
       		Call ReportStep (StatusTypes.Fail,"Not Able to create a report. Database Access Error during Report Creation", "Report Creation Page")	
					
			'Closing Excel
			ObjExcelFile.Close
			objExcel.Quit	
			Set objExcel=nothing
			Set  ObjExcelFile=nothing
			Set  ObjExcelSheet=nothing
			Set  objtreeCount=nothing
			Exit Sub
	   End If
       'Capture Database Access Error - End
       
       'Handled Application Time Out in Report Creation Page - Start
	   Set objBtnok = Browser("Analyzer").Dialog("Message from webpage").WinButton("btnOK")
	   If objBtnok.Exist(20) Then
			Call SCA.ClickOn(objBtnok,"btnOK" , "Application Time Out in Report Creation Page")
	   End If	
	   'Handled Application Time Out in Report Creation Page - End  
	        
	        
			Call IMSSCA.General.ReportAreaPermierSettings(TRIM("welVerticalDownNormal"),TRIM("  GroupTable ") )	        
    	End If
		
		
		
             
       	For a = 1 To 100 Step 1
        	If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDataSources").GetROProperty("disable")=0  Then
            Exit For
        	End If
		Next
       
       wait 2
       
        '<Shweta - 16/12/2016> handled change in DataSource Name for "+ CDTSOLAP573I*", "CDTSOLAP573I*", "CDTSOLAP573I.GEMINI.DEV*" format - Start 
	       'Set objDataSource=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDataSources")   
	       'Call SCA.SelectFromDropdown(objDataSource,strDataSourceValue)
	       
	        Set reg = createobject("vbscript.Regexp")
			reg.global = true
			reg.pattern = "[^\d]"
			DTServerVal = reg.replace(dataSourceValue, "")
			
			strAllItems = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstDdlDataSources").GetROProperty("all items")
			arrAllItems = Split(strAllItems, ";")
			For i = 0 To ubound(arrAllItems) Step 1
				If instr(1, arrAllItems(i), DTServerVal) > 1 Then
					wait 2
					'print arrAllItems(i)
					selectItem = i
					'print selectItem
					Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstDdlDataSources").Select selectItem
				End If
			Next
	
       '<Shweta - 16/12/2016> handled change in DataSource Name for "+ CDTSOLAP573I*", "CDTSOLAP573I*", "CDTSOLAP573I.GEMINI.DEV*" format - Start 
       
       
       	For a = 1 To 50 Step 1
        	If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDatabases").GetROProperty("disable")=0  Then
            Exit For
        	End If
		Next
       
       Set objDatabase=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDatabases")
       Call SCA.SelectFromDropdown(objDatabase,strDatabaseValue)
       
       	For a = 1 To 50 Step 1
        	If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlCubes").GetROProperty("disable")=0  Then
            Exit For
        	End If
		Next
       
       Set objDatacubes=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlCubes")
       Call SCA.SelectFromDropdown(objDatacubes,strCubeValue)
       
       Set objAdd=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebButton("btnAdd")
       Call SCA.ClickOn(objAdd,"AddButton" , "ReportCreation")
       
       
       
       
       ' dynamically we have to fetch the row value. So "i" should be dynamic
    
       For i=intRow_Count to intRow_Count
    	  ' Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcube").Click
           Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webCubes").Click

           For j=4 to intColumn_Count-1
    		  
	           strChild_Values=Trim(ObjExcelSheet.cells(i,j).value)
	           If  strChild_Values<>"" AND  strChild_Values<>"rowaxis" AND strChild_Values<>"columnaxis"  AND strChild_Values<>"dataaxis" AND strChild_Values<>"addtofav" AND  strChild_Values<>"deletepattern" Then
	                    Set obj_tree=description.Create
	                    obj_tree("micclass").value="WebElement"
	                    obj_tree("class").value="TreeNode"
	                    obj_tree("html tag").value="SPAN"
	                    obj_tree("innertext").value=strChild_Values
	                    wait 1
	                    set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
	                    intcount=objtreeCount.count
	                    
	                    For a = 1 To 20 Step 1
		                    If intcount<>0 Then
		                    Exit For
		                    End If
	                    Next
	                    
	                    If intcount<>0 Then
	                    If intcount=3 Then
	                    	k=intcount-2                    
		                    objtreeCount(k).fireEvent  "ondblclick"  
							                    
		                    Call ReportStep (StatusTypes.Pass,"Clicked on the Webelement"&Space(2)&strChild_Values, "Report Creation Page")           

         
		                    wait 3  
	                    	ElseIf intcount=2 Then
	                    	k=intcount-2                    
		                    objtreeCount(k).fireEvent  "ondblclick"        
		                    Call ReportStep (StatusTypes.Pass,"Clicked on the Webelement"&Space(2)&strChild_Values, "Report Creation Page")           

         
		                    wait 3  
	                    	Else
	                    	k=intcount-1                    
		                    objtreeCount(k).fireEvent  "ondblclick"        
		                    Call ReportStep (StatusTypes.Pass,"Clicked on the Webelement"&Space(2)&strChild_Values, "Report Creation Page")           

         
		                    wait 3  
	                    End If
		                      
	                    Else
	                    	Call ReportStep (StatusTypes.Warning,"Not Clicked on the Webelement"&Space(2)&strChild_Values, "Report Creation Page")
		                    Exit Sub
	                    End If    
	                    
	           End If
	           
	           If j=intColumn_Count-1 Then
		            wait 1
		            Set objDR=CreateObject("Mercury.DeviceReplay")
		            x=objtreeCount(k).GetroProperty("abs_x")
		            y=objtreeCount(k).GetroProperty("abs_y")        
		            Setting.WebPackage("ReplayType") = 2
		           ' objDR.MouseClick x,y,2
		            objtreeCount(k).rightclick
		            Setting.WebPackage("ReplayType") = 1
		            strPlaceHolderVal=lcase(ObjExcelSheet.cells(i,j+1).value)                                                                         

               
		              
		              If strPlaceHolderVal="rowaxis" Then 
		                  Set objrow=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webRowaxis")    
		                  Call SCA.ClickOn(objrow,"Rowaxis" , "ReportCreation")    
		                  wait 5                          
		              elseif  StrPlaceHolderVal="columnaxis" then                                                                                     

                                           
		                  Set objcolumn=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcolumnaxis")
		                  Call SCA.ClickOn(objcolumn,"Columnaxis" , "ReportCreation")
		                  Wait 5                      
		              elseif  StrPlaceHolderVal="dataaxis" then                                                                            
		                 Set objDataaxis=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webDataAxis")
		                 Call SCA.ClickOn(objDataaxis,"Dataaxis" , "ReportCreation")    
		                 wait 5
		              'Shweta - start 
                      elseif  StrPlaceHolderVal="addtofav" then                                                                            
                         Set objAddToFavorite=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welAddToFavorite")
                         Call SCA.ClickOn(objAddToFavorite,"AddToFavorite" , "ReportCreation")    
                         wait 5
                       elseif  StrPlaceHolderVal="deletepattern" then                                                                            
                       	 wait 2
                         Set objDelPattern=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welDeletePattern")
                         Call SCA.ClickOn(objDelPattern,"DeletePattern" , "ReportCreation")    
                         wait 5
                      'Shweta - End
		              End If    
		              
		              
	            
	           End If
             
           Next
    
       Next
       
		
	 
	
    ElseIf intgroupcreation=1 Then         
       
       For p=intColumn_Count-1 to 4  step -1
		  strChild_Values=Trim(ObjExcelSheet.cells(intRow_Count,p).value)
	      If  strChild_Values<>"" AND  strChild_Values<>"rowaxis" AND strChild_Values<>"columnaxis"  AND strChild_Values<>"dataaxis" AND strChild_Values<>"addtofav" AND  strChild_Values<>"deletepattern" Then
		                            
		    Set obj_tree=description.Create
		    obj_tree("micclass").value="WebElement"
		    obj_tree("class").value="TreeNode"
		    obj_tree("innertext").value=strChild_Values
		    wait 1
		    set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
		    intcount=objtreeCount.count    
		       If intcount<>0 Then
		       If intcount=3 Then
		       	k=intcount-3
		       	objtreeCount(k).fireEvent  "ondblclick"  
		       	Else
		       	k=intcount-1
		        objtreeCount(k).fireEvent  "ondblclick"   
		       	
		       End If
		          
		       End If                            
		                            
		End If
	 Next
	 
	 ElseIf intgroupcreation=2 then
	 
	 For p=intColumn_Count-1 to 4  step -1
		  strChild_Values=Trim(ObjExcelSheet.cells(intRow_Count,p).value)
	      If  strChild_Values<>"" AND  strChild_Values<>"rowaxis" AND strChild_Values<>"columnaxis"  AND strChild_Values<>"dataaxis" Then
		                            
		    Set obj_tree=description.Create
		    obj_tree("micclass").value="WebElement"
		    obj_tree("class").value="TreeNode"
		    obj_tree("innertext").value=strChild_Values
		    wait 1		    
		    
		    set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
		    intcount=objtreeCount.count    
		       If intcount<>0 Then
		       If intcount=2 Then
		       	 k=intcount-1
		       	 objtreeCount(k).fireEvent  "ondblclick" 
		       	 else
		       	 k=intcount-1
		       	  objtreeCount(k).fireEvent  "ondblclick" 
		       End If
		        		            
		       End If 


			 Set objDR=CreateObject("Mercury.DeviceReplay")
		            x=objtreeCount(k).GetroProperty("abs_x")
		            y=objtreeCount(k).GetroProperty("abs_y") 
					Setting.WebPackage("ReplayType") = 2		            
		            'objDR.MouseClick x,y,2
					objtreeCount(k).rightclick
		            Setting.WebPackage("ReplayType") = 1
		            strPlaceHolderVal=lcase(ObjExcelSheet.cells(intRow_Count,intColumn_Count).value)                                                  

                                      
		              
		              If strPlaceHolderVal="rowaxis" Then 
		                  Set objrow=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webRowaxis")    
		                  Call SCA.ClickOn(objrow,"Rowaxis" , "ReportCreation")    
		                  wait 5                          
		              elseif  StrPlaceHolderVal="columnaxis" then                                                                                     

                                           
		                  Set objcolumn=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcolumnaxis")
		                  Call SCA.ClickOn(objcolumn,"Columnaxis" , "ReportCreation")
		                  Wait 5                      
		              elseif  StrPlaceHolderVal="dataaxis" then                                                                            
		                 Set objDataaxis=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webDataAxis")
		                 Call SCA.ClickOn(objDataaxis,"Dataaxis" , "ReportCreation")    
		                 wait 5
		             End If    
		    
		   Exit For                         
		End If
		
	 Next
	 
    
  End If 
  
    ObjExcelFile.Close
	objExcel.Quit    
	Set objExcel=nothing
	Set ObjExcelFile=nothing
	Set ObjExcelSheet=nothing
	Set objtreeCount=nothing	
  
 End Sub	
	
	Public Sub GroupTableCreationInSCARajeshCode(ByVal intgroupcreation,ByVal sheetNum, ByVal newReport, ByVal intRow_Count, ByVal strF,Byval strF1)  
    
        Dim objExcel,ObjExcelFile,ObjExcelSheet,objDataSource,objDatabase,objDatacubes,obj_tree,objDR,objrow,objcolumn,objDataaxis,objAdd,objtreeCount,objCreate
        Dim intColumn_Count,intcount
        Dim dataSourceValue, strDataSourceValue,strDatabaseValue,strCubeValue,strChild_Values,strPlaceHolderVal,p, objBtnok, objErrorOk
        '<Looping Variables for the row and column>
        Dim i,j,k,x,y,a
        
        '<shweta 16/6/2015 Closing All opened Excel Files- start
		'objUtils.KillProcess("excel.exe")
		Systemutil.CloseProcessByName "excel.exe"
		'<shweta 16/6/2015 Closing All opened Excel Files- end
		
        Set objExcel = CreateObject("Excel.Application") 
        'objExcel.Visible =True     
    
        Set ObjExcelFile = objExcel.Workbooks.Open(Environment.Value("CurrDir")&Environment.Value("ReportCreationFile"))
        Set ObjExcelSheet = ObjExcelFile.Sheets(sheetNum) 
        intColumn_Count = ObjExcelSheet.UsedRange.Columns.Count
    
        dataSourceValue=TRIM(ObjExcelSheet.cells(2,1).value)
        strDataSourceValue = "+ "&dataSourceValue
        strDatabaseValue=TRIM(ObjExcelSheet.cells(2,2).value)
        strCubeValue=TRIM(ObjExcelSheet.cells(2,3).value)
        
		If intgroupcreation=0 Then		
    
    	
          
		
		
		If newReport = 0 Then
    		Set objCreate=Browser("Analyzer").Page("Shared_Folder").Frame("view").Image("imgnewreport")      
	        Call SCA.ClickOn(objCreate,"newreport","Report Creation")
	        
	        
	        'Capture Database Access Error - Start
       If Browser("Analyzer").Frame("ErrorFrame").WebElement("Database access error.").Exist(0) Then
       		Set objErrorOk=Browser("Analyzer").Frame("ErrorFrame").WebButton("btnOK")
       		Call SCA.ClickOn(objErrorOk,"btnOK" , "Database Access Error during Report Creation")
       		Call ReportStep (StatusTypes.Fail,"Not Able to create a report. Database Access Error during Report Creation", "Report Creation Page")	
					
			'Closing Excel
			ObjExcelFile.Close
			objExcel.Quit	
			Set objExcel=nothing
			Set  ObjExcelFile=nothing
			Set  ObjExcelSheet=nothing
			Set  objtreeCount=nothing
			Exit Sub
	   End If
       'Capture Database Access Error - End
       
       'Handled Application Time Out in Report Creation Page - Start
	   Set objBtnok = Browser("Analyzer").Dialog("Message from webpage").WinButton("btnOK")
	   If objBtnok.Exist(5) Then
			Call SCA.ClickOn(objBtnok,"btnOK" , "Application Time Out in Report Creation Page")
	   End If	
	   'Handled Application Time Out in Report Creation Page - End  
	        
	        
			Call IMSSCA.General.ReportAreaPermierSettings(TRIM("welVerticalDownNormal"),TRIM("  GroupTable ") )	        
    	End If
		
		
		
             
       	For a = 1 To 100 Step 1
        	If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDataSources").GetROProperty("disable")=0  Then
            Exit For
        	End If
		Next
       
       wait 2
       
       Set objDataSource=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDataSources")   
       Call SCA.SelectFromDropdown(objDataSource,strDataSourceValue)
       
       	For a = 1 To 50 Step 1
        	If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDatabases").GetROProperty("disable")=0  Then
            Exit For
        	End If
		Next
       
       Set objDatabase=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDatabases")
       Call SCA.SelectFromDropdown(objDatabase,strDatabaseValue)
       
       	For a = 1 To 50 Step 1
        	If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlCubes").GetROProperty("disable")=0  Then
            Exit For
        	End If
		Next
       
       Set objDatacubes=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlCubes")
       Call SCA.SelectFromDropdown(objDatacubes,strCubeValue)
       
       Set objAdd=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebButton("btnAdd")
       Call SCA.ClickOn(objAdd,"AddButton" , "ReportCreation")
       
       
       
       
       ' dynamically we have to fetch the row value. So "i" should be dynamic
    
       For i=intRow_Count to intRow_Count
    	  ' Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcube").Click
           Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webCubes").Click

           For j=4 to intColumn_Count-1
    		  
	           strChild_Values=Trim(ObjExcelSheet.cells(i,j).value)
	           If  strChild_Values<>"" AND  strChild_Values<>"rowaxis" AND strChild_Values<>"columnaxis"  AND strChild_Values<>"dataaxis" AND strChild_Values<>"addtofav" AND  strChild_Values<>"deletepattern" Then
	                    Set obj_tree=description.Create
	                    obj_tree("micclass").value="WebElement"
	                    obj_tree("class").value="TreeNode"
	                    obj_tree("innertext").value=strChild_Values
	                    wait 1
	                    set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
	                    intcount=objtreeCount.count
	                    
	                    For a = 1 To 20 Step 1
		                    If intcount<>0 Then
		                    Exit For
		                    End If
	                    Next
	                    
	                    If intcount<>0 Then
	                    If intcount=3 Then
	                    	k=intcount-2                    
		                    objtreeCount(k).fireEvent  "ondblclick"  
							                    
		                    Call ReportStep (StatusTypes.Pass,"Clicked on the Webelement"&Space(2)&strChild_Values, "Report Creation Page")                    
		                    wait 3  
	                    	ElseIf intcount=2 Then
	                    	    If strF = 1 AND strChild_Values = "Additional attributes" Then
	                    	    	k=intcount-1                    
				                    objtreeCount(k).fireEvent  "ondblclick"        
				                    Call ReportStep (StatusTypes.Pass,"Clicked on the Webelement"&Space(2)&strChild_Values, "Report Creation Page")                    
				                    wait 3 
	                    	    
	                    	    Else
	                    	    
	                    	    	k=intcount-2                   
				                    objtreeCount(k).fireEvent  "ondblclick"        
				                    Call ReportStep (StatusTypes.Pass,"Clicked on the Webelement"&Space(2)&strChild_Values, "Report Creation Page")                    
				                    wait 3 
	                    	    
	                    	    End If
'		                    	k=intcount-2                    
'			                    objtreeCount(k).fireEvent  "ondblclick"        
'			                    Call ReportStep (StatusTypes.Pass,"Clicked on the Webelement"&Space(2)&strChild_Values, "Report Creation Page")                    
'			                    wait 3 
	                    	Else
	                    	k=intcount-1                    
		                    objtreeCount(k).fireEvent  "ondblclick"        
		                    Call ReportStep (StatusTypes.Pass,"Clicked on the Webelement"&Space(2)&strChild_Values, "Report Creation Page")                    
		                    wait 3  
	                    End If
		                      
	                    Else
	                    	Call ReportStep (StatusTypes.Warning,"Not Clicked on the Webelement"&Space(2)&strChild_Values, "Report Creation Page")
		                    Exit Sub
	                    End If    
	                    
	           End If
	           
	           If j=intColumn_Count-1 Then
		            wait 1
		            
'		            Set objDR=CreateObject("Mercury.DeviceReplay")
'		            x=objtreeCount(k).GetroProperty("abs_x")
'		            y=objtreeCount(k).GetroProperty("abs_y")        
'		            objDR.MouseClick x,y,2
		            
		            objtreeCount(k).RightClick
		            
		            strPlaceHolderVal=lcase(ObjExcelSheet.cells(i,j).value)                                                                                        
		              
		              If strPlaceHolderVal="rowaxis" Then 
		                  Set objrow=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webRowaxis")    
		                  Call SCA.ClickOn(objrow,"Rowaxis" , "ReportCreation")    
		                  wait 5                          
		              elseif  StrPlaceHolderVal="columnaxis" then                                                                                                                                
		                  Set objcolumn=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcolumnaxis")
		                  Call SCA.ClickOn(objcolumn,"Columnaxis" , "ReportCreation")
		                  Wait 5                      
		              elseif  StrPlaceHolderVal="dataaxis" then                                                                            
		                 Set objDataaxis=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webDataAxis")
		                 Call SCA.ClickOn(objDataaxis,"Dataaxis" , "ReportCreation")    
		                 wait 5
		              'Shweta - start 
                      elseif  StrPlaceHolderVal="addtofav" then                                                                            
                         Set objAddToFavorite=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welAddToFavorite")
                         Call SCA.ClickOn(objAddToFavorite,"AddToFavorite" , "ReportCreation")    
                         wait 5
                       elseif  StrPlaceHolderVal="deletepattern" then                                                                            
                       	 wait 2
                         Set objDelPattern=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welDeletePattern")
                         Call SCA.ClickOn(objDelPattern,"DeletePattern" , "ReportCreation")    
                         wait 5
                      'Shweta - End
		              End If    
		              
		              
	            
	           End If
             
           Next
    
       Next
       
		
	 
	
    ElseIf intgroupcreation=1 Then         
       
       For p=intColumn_Count-1 to 4  step -1
		  strChild_Values=Trim(ObjExcelSheet.cells(intRow_Count,p).value)
	      If  strChild_Values<>"" AND  strChild_Values<>"rowaxis" AND strChild_Values<>"columnaxis"  AND strChild_Values<>"dataaxis" AND strChild_Values<>"addtofav" AND  strChild_Values<>"deletepattern" Then
		                            
		    Set obj_tree=description.Create
		    obj_tree("micclass").value="WebElement"
		    obj_tree("class").value="TreeNode"
		    obj_tree("innertext").value=strChild_Values
		    wait 1
		    set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
		    intcount=objtreeCount.count    
		       If intcount<>0 Then
		       If intcount=3 Then
		       	k=intcount-3
		       	objtreeCount(k).fireEvent  "ondblclick"  
		       	ElseIf intcount=2 Then
	                    	    If strF = 1 AND strChild_Values = "Additional attributes" Then
	                    	    	k=intcount-1                    
				                    objtreeCount(k).fireEvent  "ondblclick"        
				                    Call ReportStep (StatusTypes.Pass,"Clicked on the Webelement"&Space(2)&strChild_Values, "Report Creation Page")                    
				                    wait 3 
	                    	    
	                    	    Else
	                    	    
	                    	    	k=intcount-2                    
				                    objtreeCount(k).fireEvent  "ondblclick"        
				                    Call ReportStep (StatusTypes.Pass,"Clicked on the Webelement"&Space(2)&strChild_Values, "Report Creation Page")                    
				                    wait 3 
	                    	    
	                    	    End If
		       	Else
		       	k=intcount-1
		        objtreeCount(k).fireEvent  "ondblclick"   
		       	
		       End If
		          
		       End If                            
		                            
		End If
	 Next
	 
	 ElseIf intgroupcreation=2 then
	 
	 For p=intColumn_Count-1 to 4  step -1
		  strChild_Values=Trim(ObjExcelSheet.cells(intRow_Count,p).value)
	      If  strChild_Values<>"" AND  strChild_Values<>"rowaxis" AND strChild_Values<>"columnaxis"  AND strChild_Values<>"dataaxis" Then
		                            
		    Set obj_tree=description.Create
		    obj_tree("micclass").value="WebElement"
		    obj_tree("class").value="TreeNode"
		    obj_tree("innertext").value=strChild_Values
		    wait 1		    
		    
		    set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
		    intcount=objtreeCount.count    
		       If intcount<>0 Then
		       If intcount = 2 AND strChild_Values <> "Additional attributes" Then
		       	 k=intcount-2
		       	 objtreeCount(k).fireEvent  "ondblclick" 
		       	 else
		       	 k=intcount-1
		       	  objtreeCount(k).fireEvent  "ondblclick" 
		       End If
		        		            
		       End If 


'			 Set objDR=CreateObject("Mercury.DeviceReplay")
'             x=objtreeCount(k).GetroProperty("abs_x")
'             y=objtreeCount(k).GetroProperty("abs_y")        
'             objDR.MouseClick x,y,2
		            
		     objtreeCount(k).RightClick       
		            
		            strPlaceHolderVal=lcase(ObjExcelSheet.cells(intRow_Count,intColumn_Count).value)                                                                                        
		              
		              If strPlaceHolderVal="rowaxis" Then 
		                  Set objrow=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webRowaxis")    
		                  Call SCA.ClickOn(objrow,"Rowaxis" , "ReportCreation")    
		                  wait 5                          
		              elseif  StrPlaceHolderVal="columnaxis" then                                                                                                                                
		                  Set objcolumn=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcolumnaxis")
		                  Call SCA.ClickOn(objcolumn,"Columnaxis" , "ReportCreation")
		                  Wait 5                      
		              elseif  StrPlaceHolderVal="dataaxis" then                                                                            
		                 Set objDataaxis=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webDataAxis")
		                 Call SCA.ClickOn(objDataaxis,"Dataaxis" , "ReportCreation")    
		                 wait 5
		             End If    
		    
		   Exit For                         
		End If
		
	 Next
	 
    
  End If 
  
    ObjExcelFile.Close
	objExcel.Quit    
	Set objExcel=nothing
	Set ObjExcelFile=nothing
	Set ObjExcelSheet=nothing
	Set objtreeCount=nothing	
  
 End Sub	
	
	Public Function ReportFilterExtractionCompression(ByVal intsheetName, ByVal intrownum)
        Dim objExcel,ObjExcelFile,ObjExcelSheet,intRow_Count,intColumn_Count,j,strChild_Values,obj_tree,objtreeCount,intcount,k,obj_tree1,objtreeCount1,intcount1,m,x1,y1,StrvalueData,i	    
	    Dim objDR,x,y,strPlaceHolderVal,objrow,objcolumn,objDataaxis    
	    
	    Set objExcel = CreateObject("Excel.Application") 
	    objExcel.Visible =False     
	    
	    Set ObjExcelFile = objExcel.Workbooks.Open(Environment.Value("CurrDir")&Environment.Value("ReportCreationFile"))
	    Set ObjExcelSheet = ObjExcelFile.Sheets(intsheetName)
	    
	    intRow_Count = ObjExcelSheet.UsedRange.Rows.Count 
	    intColumn_Count = ObjExcelSheet.UsedRange.Columns.Count   
         
         For p=intColumn_Count-2 to 4  step -1
	         strChild_Values=Trim(ObjExcelSheet.cells(intrownum,p).value)
	         strbkch=Trim(ObjExcelSheet.cells(intrownum,intColumn_Count).value)                         
	         If  strChild_Values<>"" AND  strChild_Values<>"rowaxis" AND strChild_Values<>"columnaxis"  AND strChild_Values<>"dataaxis" AND strbkch<>"Yes" Then
	            Set obj_tree=description.Create
	            obj_tree("micclass").value="WebElement"
	            obj_tree("class").value="TreeNode"
	            obj_tree("innertext").value=strChild_Values
	            wait 1
	            set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
	            intcount=objtreeCount.count    
	            If intcount<>0 Then
	            	If intcount = 2 Then
	            		 k=intcount-2
	            	Else
	            		 k=intcount-1
	            	End If
	               objtreeCount(k).fireEvent  "ondblclick"    
	             End If                          
	         End If
	      Next
	End Function
	
	'	<Enters Datasource/databse/cub details. Captures x/y co-ordinates of the tree hierarchy elements after expanding and returns x/y co-ordinates of the tree hierarchy elements>
	Public Function ReportFilterExtraction(ByVal intsheetName,ByVal intaddpivotal,ByVal intrownum,ByVal treeDetails,ByVal strF2)
    
	    Dim objExcel,ObjExcelFile,ObjExcelSheet,intRow_Count,intColumn_Count,j,strChild_Values,obj_tree,objtreeCount,intcount,k,obj_tree1,objtreeCount1,intcount1,m,x1,y1,StrvalueData,i	    
	    Dim objDR,x,y,strPlaceHolderVal,objrow,objcolumn,objDataaxis    
	    
	    Set objExcel = CreateObject("Excel.Application") 
	    objExcel.Visible =False     
	    
	    Set ObjExcelFile = objExcel.Workbooks.Open(Environment.Value("CurrDir")&Environment.Value("ReportCreationFile"))
	    Set ObjExcelSheet = ObjExcelFile.Sheets(intsheetName)
	    
	    intRow_Count = ObjExcelSheet.UsedRange.Rows.Count 
	    intColumn_Count = ObjExcelSheet.UsedRange.Columns.Count   

		'######## Shweta 20/9/2016 - Added Datasource/Datacube/Cube/Database entering section - Start
		If treeDetails <> "" Then
		
			Set objCreate=Browser("Analyzer").Page("Shared_Folder").Frame("view").Image("imgnewreport")        
            Call SCA.ClickOn(objCreate,"newreport","Report Creation")    
		
			'Handled Application Time Out in Report Creation Page - Start
	       Set objBtnok = Browser("Analyzer").Dialog("Message from webpage").WinButton("btnOK")
	       If objBtnok.Exist(5) Then
	            Call SCA.ClickOn(objBtnok,"btnOK" , "Application Time Out in Report Creation Page")
	       End If    
	       'Handled Application Time Out in Report Creation Page - End
	       
			dataSourceValue=TRIM(ObjExcelSheet.cells(2,1).value)
	        strDataSourceValue = "+ "&dataSourceValue       
	        strDatabaseValue=TRIM(ObjExcelSheet.cells(2,2).value)
	        strCubeValue=TRIM(ObjExcelSheet.cells(2,3).value)
	
			For a = 1 To 100 Step 1           
		       If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDataSources").GetROProperty("disable")=0  Then
		         Exit For
		       End If           
	        Next
	       
	        wait 2
	       
	       Set objDataSource=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDataSources")   
	       Call SCA.SelectFromDropdown(objDataSource,strDataSourceValue)
	       
	       wait 2
	       
	       Set objDatabase=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDatabases")
	       Call SCA.SelectFromDropdown(objDatabase,strDatabaseValue)       
		
			wait 2
		   Set objDatacubes=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlCubes")
	       Call SCA.SelectFromDropdown(objDatacubes,strCubeValue)
	       
	       wait 2
	       Set objAdd=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebButton("btnAdd")
	       Call SCA.ClickOn(objAdd,"AddButton" , "ReportCreation")
	       
	       wait 2
           Browser("Analyzer").Page("ReportCreation").Sync
           Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").RefreshObject
                      
		End If
		
		'######## Shweta 20/9/2016 - Added Datasource/Datacube/Cube/Database entering section - End
		
	    For j=4 to intColumn_Count-2
	    
	         strChild_Values=Trim(ObjExcelSheet.cells(intrownum,j).value)
	         If  strChild_Values<>"" AND  strChild_Values<>"rowaxis" AND strChild_Values<>"columnaxis"  AND strChild_Values<>"dataaxis" Then
                 Set obj_tree=description.Create
	             obj_tree("micclass").value="WebElement"
	             obj_tree("class").value="TreeNode"
	             obj_tree("innertext").value=strChild_Values
	             wait 3
	             set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
	             intcount=objtreeCount.count                                                                                
	             If intcount<>0   Then
	                 k=intcount-1
	                objtreeCount(k).fireEvent  "ondblclick"                 
	             End If 
                    wait 10	             
	                StrvalueData=Trim(ObjExcelSheet.cells(intrownum,j+1).value)
	                If StrvalueData="" OR StrvalueData="rowaxis" Then
	                   Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcube").Click
	                   Set obj_tree1=description.Create
	                   obj_tree1("micclass").value="WebElement"
	                   obj_tree1("class").value="TreeNode"
	                   obj_tree1("innertext").value=strChild_Values  
					   'objtreeCount(k).fireEvent  "ondblclick"
	                    wait 3
	                    set objtreeCount1= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree1)
	                    intcount1=objtreeCount1.count 
	                    
	                    Set obj_trees=description.Create
	                   obj_trees("micclass").value="WebElement"
	                   obj_trees("class").value="trvSchemaHover HoverNode HoverNode"
	                   obj_trees("innertext").value=strChild_Values  
					   'objtreeCount(k).fireEvent  "ondblclick"
	                    wait 3
	                    set objtreeCounts= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_trees)
	                    intcounts=objtreeCounts.count
	                    If intcounts<>0 Then
	                    	If intcounts=3 Then
	                          m=intcounts-2
	                       ElseIf intcounts=2 Then
	                           m=intcounts-2
	                       Else
	                           m=objtreeCounts.count-1
	                       End If 
	                          objtreeCounts(m).highlight
	                       x1=objtreeCounts(m).GetroProperty("x")
	                       y1=objtreeCounts(m).GetroProperty("y")
	                       print ":trvSchemaHover HoverNode HoverNode"&y1
	                       ReportFilterExtraction=x1&"-"&y1
	                       Exit For
	                    End If
	                    
	                    
	                    
	                    If intcount1<>0 Then
	                    	If intcount1=3 Then
	                          m=intcount1-2
	                       ElseIf intcount1=2 Then
	                           m=intcount1-2
	                       Else
	                           m=objtreeCount1.count-1
	                       End If 
	                       objtreeCount1(m).highlight
	                       x1=objtreeCount1(m).GetroProperty("x")
	                       y1=objtreeCount1(m).GetroProperty("y")
	                       print ":::TreeNode"&y1
	                       ReportFilterExtraction=x1&"-"&y1
	                       Exit For
	                    End If
	                       
	                        
	               End If
                     If j=intColumn_Count-2 AND intaddpivotal=0 Then
	            		wait 1
	            		Set objDR=CreateObject("Mercury.DeviceReplay")
	            		x=objtreeCount(k).GetroProperty("abs_x")
	            		y=objtreeCount(k).GetroProperty("abs_y") 
						wait 1	    
						Setting.WebPackage("ReplayType") = 2						
	            		objDR.MouseClick x,y,2
	            		Setting.WebPackage("ReplayType") = 1
	            		strPlaceHolderVal=lcase(ObjExcelSheet.cells(intrownum,j+1).value)                                                                                        
	              		If strPlaceHolderVal="rowaxis" Then 
	                  		Set objrow=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webRowaxis")    
	                  		Call SCA.ClickOn(objrow,"Rowaxis" , "ReportCreation")    
	                  		wait 5                          
	              		elseif  StrPlaceHolderVal="columnaxis" then                                                                                                                                
	                  		Set objcolumn=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcolumnaxis")
	                  		Call SCA.ClickOn(objcolumn,"Columnaxis" , "ReportCreation")
	                  		Wait 5                      
	              		elseif  StrPlaceHolderVal="dataaxis" then                                                                            
	                 		Set objDataaxis=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webDataAxis")
	                 		Call SCA.ClickOn(objDataaxis,"Dataaxis" , "ReportCreation")    
	                 		wait 5
	            		End If    
	
	          		End If
	           
  
	         End If
	            
	
	           
	               
	       Next  	
				  
		ObjExcelFile.Close
		objExcel.Quit    
		Set objExcel=nothing
		Set  ObjExcelFile=nothing
		Set  ObjExcelSheet=nothing
		Set  objtreeCount=nothing
	
	End Function
	
    '..........................................................................................................................................	

	'<	To Perform the click operation on the Group Dialouge box to create the type of group>
	'<  objclick:- Name of the object >
	'<Author>shobha<author>
	
    '..........................................................................................................................................	
    
	Public Sub GrouptabToolbar(Byval objclick)
	Dim objtoobar,objtab	
	
		Set objtoobar=description.Create
		objtoobar("micclass").value="Image"
		'objtoobar("micclass").value="WebElement"
		objtoobar("title").value=objclick		
		set objtab=Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebTable("tabbar_toolbar").ChildObjects(objtoobar)
		objtab(0).Click		

	End Sub
	
	
	
   '''<To Create the ADDMDXGroup depending on the selection>
   '''<strO_selection:- To select the operation on Group>
   '''<strName :- Name of the Group>
   '''<strCaption:- Name of the Group Caption>
   '''<strFormat:-selection of the Forrmat>
   '''<strDisplayAs:-Display check box selection>
   '''<strBehaviour:-Behaviour selection>
   '''<strContenttype
   '''<strDelimiter:
   '''<strValues:
   '''<strMemberValues: Memberrs of the group>
   '''< strupdatename
   '''< strupdatecap
   '''<	strTC:
   '''< intlineDelimiterselection: line delimiter selection> 
   '''< Author>shobha<Author>    
	
	Public Sub ADDMDXGroup(ByVal strO_selection,Byval strName,ByVal strCaption,ByVal strFormat,ByVal strDisplayAs, ByVal strBehaviour,ByVal strContenttype, ByVal strDelimiter,ByVal strValues,ByVal strupdatename,ByVal strupdatecap,ByVal strTC,ByVal DelimiterLineSelection)
		
	 'Dim strO_selection
	 Dim objtxtname,objtxtCaption,objchkformat,objrdodisplay,objrdoBehaviour,objrdocontenttype,objrdoDelimiter,objtxtMember,objevalBtn,objweblist,objokbtn,intR_Val 
	 Dim objedit,objgroupname,objgroup,objeditdrop,d_returnVal
	

	 If strO_selection="CreateGroup" Or strO_selection="CancelGroup" Then
	 	
	 	 Set objtxtname=Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebEdit("txtName")
	 	 Call SCA.SetText(objtxtname,strName, "txtGroupName","GroupCreation Dialouge" )
	 	 
	 	 
         Set objtxtCaption=Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebEdit("txtCaption")
         Call SCA.SetText(objtxtCaption,strCaption, "txtGroupCaption","GroupCreation Dialouge" )
         
         Set objchkformat=Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebCheckBox(strFormat)
         Call SCA.SetCheckBox(objchkformat,"chkFormatString","On","GroupCreation Dialouge")
         
         
         Set objrdodisplay=Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebRadioGroup("rdoDisplayas")
         Call SCA.ClickOnRadio(objrdodisplay,strDisplayAs ,"GroupCreation Dialouge")
         
         
		 Set objrdoBehaviour=Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebRadioGroup("rdoBehaviour")
		 Call SCA.ClickOnRadio(objrdoBehaviour,strBehaviour ,"GroupCreation Dialouge")
		 
		 
		 Set objrdocontenttype= Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebRadioGroup("rdocontenttype")
		 Call SCA.ClickOnRadio(objrdocontenttype,strContenttype ,"GroupCreation Dialouge")		  

		 
		 d_returnVal=IMSSCA.Validations.delimiterValidation_MDXGroup(strDelimiter,strValues,DelimiterLineSelection)
		 
		 
		 If strTC="positiveTC" Then
		 	
		 If d_returnVal=0 Then
		 	
		 Call ReportStep (StatusTypes.Pass,strDelimiter&Space(3)&"Members are Evaluated successfully","MDX Group table creation" )
	     Else
	     Call ReportStep (StatusTypes.Fail,strDelimiter&Space(3)&"Members not are Evaluated successfully","MDX Group table creation" )
		 End If	
		 
		 ElseIf strTC="negetiveTC" Then
		 
		 
		 If d_returnVal=1 Then
		 	
		 Call ReportStep (StatusTypes.Pass,strDelimiter&Space(3)&"Members are not Evaluated successfully","MDX Group table creation" )
	     Else
	     Call ReportStep (StatusTypes.Fail,strDelimiter&Space(3)&"Members are Evaluated successfully","MDX Group table creation" )
		 End If	
		 	
		 End If
		 	 
		 
	 	
	 End If
	 	 
	   
	 
	 Select Case strO_selection
	 
		Case "CreateGroup"
		Set objokbtn=Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebButton("btnOK")
		Call SCA.ClickOn(objokbtn,"btnoK", "GroupCreation Dialouge")
		intR_Val=IMSSCA.Validations.Groupcreation_Verification(strName,strCaption,"")	
		If intR_Val=0 Then
			Call ReportStep (StatusTypes.Pass, "Create Group Validation:-"&Space(7)&strName&Space(5)&"Group is created and Reflected in the Group dialouge box"&Space(5)&strCaption&Space(5)&"created in the Report creation area","GroupCreation Dialouge box")
			ElseIf intR_Val=1 Then
			Call ReportStep (StatusTypes.Pass, "Create Group Validation:-"&Space(7)&strName&Space(5)&"Group is created and Reflected in the Group dialouge box"&Space(5)&strCaption&Space(5)&"not created in the Report creation area","GroupCreation Dialouge box")
			else
			Call ReportStep (StatusTypes.Fail, "Create Group Validation:-"&Space(7)&strName&Space(5)&"Group is not created and not Reflected in the Group dialouge box"&Space(5)&strCaption&Space(5)&"not created in the Report creation area","GroupCreation Dialouge box")
		End If

		
	    Case  "CancelGroup"
	    
	    Set objokbtn=Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebButton("btnCancel")
		Call SCA.ClickOn(objokbtn,"btnoK", "GroupCreation Dialouge")
		intR_Val=IMSSCA.Validations.Groupcreation_Verification(strName,strCaption,"")	
		If intR_Val<>0 Then
			Call ReportStep (StatusTypes.Pass, "Cancel Group Validation:-"&Space(7)&strName&Space(5)&"Group is not created and Reflected in the Group dialouge box"&Space(5)&strCaption&Space(5)&"not created in the Report creation area","GroupCreation Dialouge box")
			else
			Call ReportStep (StatusTypes.Fail, "Cancel Group Validation:-"&Space(7)&strName&Space(5)&"Group is created and Reflected in the Group dialouge box"&Space(5)&strCaption&Space(5)&"created in the Report creation area","GroupCreation Dialouge box")
		End If
		
		Case  "EditGroup"
				
            set objeditdrop=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("Editdown_normal")
            Call SCA.ClickOn(objeditdrop,"GroupEditdropdown", "GroupCreation Dialouge")
            
            Set objedit=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webeditgroup")
            Call SCA.ClickOn(objedit,"GroupEdit", "GroupCreation Dialouge")

            
            Set objgroupname=description.Create
            objgroupname("micclass").value="WebElement"
            objgroupname("class").value="trvItemsNode TreeNode"
            objgroupname("innerhtml").value=strName
            
            Set objgroup=Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebTable("tabGroupName").ChildObjects(objgroupname)
            objgroup(0).Click   
			
			wait 2
			
		  Set objtxtname=Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebEdit("txtName")
	 	  Call SCA.SetText(objtxtname,strupdatename, "txtGroupName","GroupCreation Dialouge" )
	 	 
	 	 
          Set objtxtCaption=Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebEdit("txtCaption")
          Call SCA.SetText(objtxtCaption,strupdatecap, "txtGroupCaption","GroupCreation Dialouge" )
          
          
          Set objchkformat=Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebCheckBox(strFormat)
         Call SCA.SetCheckBox(objchkformat,"chkFormatString","On","GroupCreation Dialouge")
         
         
         Set objrdodisplay=Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebRadioGroup("rdoDisplayas")
         Call SCA.ClickOnRadio(objrdodisplay,strDisplayAs ,"GroupCreation Dialouge")
         
         
		 Set objrdoBehaviour=Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebRadioGroup("rdoBehaviour")
		 Call SCA.ClickOnRadio(objrdoBehaviour,strBehaviour ,"GroupCreation Dialouge")
		 
		 
		 Set objrdocontenttype= Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebRadioGroup("rdocontenttype")
		 Call SCA.ClickOnRadio(objrdocontenttype,strContenttype ,"GroupCreation Dialouge")		  

		       
          Set objokbtn=Browser("Analyzer").Page("GroupCreationPage").Frame("Grouptab_dialog").WebButton("btnOK")
		  Call SCA.ClickOn(objokbtn,"btnoK", "GroupCreation Dialouge")
		  wait 2
		 
		 intR_Val=IMSSCA.Validations.Groupcreation_Verification(strupdatename,strupdatecap,"")	
		 If intR_Val=0 Then
			Call ReportStep (StatusTypes.Pass,"Edit Group Validatio:-"&Space(7)&strName&Space(4)&"Updated to "&Space(2)&strupdatename ,"GroupCreation Dialouge box")
			ElseIf intR_Val=1 Then
			Call ReportStep (StatusTypes.Pass, "Edit Group Validatio:-"&Space(7)&strName&Space(4)&"Updated to "&Space(2)&strupdatename,"GroupCreation Dialouge box")
			else
			Call ReportStep (StatusTypes.Fail, "Edit Group Validatio:-"&Space(7)&strName&Space(4)&"Updated to "&Space(2)&strupdatename,"GroupCreation Dialouge box")
		 End If						
			
		End Select	
		
		
	End Sub
	
'	
'   '''<To Create the GroupFilter group depending on the selection>
'   '''<strO_selection:- To select the operation on Group>
'   '''<strName :- Name of the Group>
'   '''<strCaption:- Name of the Group Caption>
'   '''<strDisplayVal:-Display check box selection>
'   '''<strBehaviour:-Behaviour selection>  
'   '''<strmemberselection: Memberrs of the group>
'   '''<	strlistSelection:Selecting the Filter Values either equal or starts with
'   '''< intlineDelimiterselection: line delimiter selection> 
'   '''< Author>shobha<Author> 
'
'




  '''<To Create the AddCalculation Group depending on the selection>
   '''<strO_selection:- To select the operation on Group>
   '''<strName :- Name of the Group>
   '''<strCaption:- Name of the Group Caption>
   '''<strMDX:- MDX Query to pass> 
   '''< Author>shobha<Author> 

Public Sub AddCalculation(ByVal strName,ByVal strCaption,ByVal strMDX)


	Dim objtxtname,objtxtcaption,objMDX,objchkformat,objchkformat1,objokbtn,intR_Val


	Set objtxtname=Browser("Analyzer").Page("GroupCreationPage").Frame("Add Calculation").WebEdit("txtName")
    Call SCA.SetText(objtxtname,strName, "txtGroupName","Add Calculation GroupCreation Dialouge" )
    
    Set objtxtcaption=Browser("Analyzer").Page("GroupCreationPage").Frame("Add Calculation").WebEdit("txtCaption")
    Call SCA.SetText(objtxtcaption,strCaption, "txtGroupCaption","Add Calculation GroupCreation Dialouge" )
    
	Set objMDX=Browser("Analyzer").Page("GroupCreationPage").Frame("Add Calculation").WebEdit("txtExpression")
	Call SCA.SetText_MultipleLineArea(objMDX,strMDX, "txtGroupCaption","Add Calculation GroupCreation Dialouge" )
	
	Set objchkformat=Browser("Analyzer").Page("GroupCreationPage").Frame("Add Calculation").WebCheckBox("WebCheckBox")
    Call SCA.SetCheckBox(objchkformat,"Filter Member","On","Add Calculation GroupCreation Dialouge")
    
    
    Set objchkformat1=Browser("Analyzer").Page("GroupCreationPage").Frame("Add Calculation").WebCheckBox("WebCheckBox_2")
    Call SCA.SetCheckBox(objchkformat1,"Filter Member","On","Add Calculation GroupCreation Dialouge")
	
	Set objokbtn=Browser("Analyzer").Page("GroupCreationPage").Frame("Add Calculation").WebButton("btnOK")
	Call SCA.ClickOn(objokbtn,"btnoK", "GroupCreation Dialouge")	
	
	 intR_Val=IMSSCA.Validations.Groupcreation_Verification(strName,strCaption,"")    
        If intR_Val=0 Then
            Call ReportStep (StatusTypes.Pass, "Create Group Validation:-"&Space(7)&strName&Space(5)&"Group is created and Reflected in the Group dialouge box"&Space(5)&strCaption&Space(5)&"created in the Report creation area","GroupCreation Dialouge box")
        ElseIf intR_Val=1 Then
            Call ReportStep (StatusTypes.Pass, "Create Group Validation:-"&Space(7)&strName&Space(5)&"Group is created and Reflected in the Group dialouge box"&Space(5)&strCaption&Space(5)&"not created in the Report creation area","GroupCreation Dialouge box")
        else
            Call ReportStep (StatusTypes.Fail, "Create Group Validation:-"&Space(7)&strName&Space(5)&"Group is not created and not Reflected in the Group dialouge box"&Space(5)&strCaption&Space(5)&"not created in the Report creation area","GroupCreation Dialouge box")
        End If

	
	
End Sub

'‘’’’’’’’’’’’’’’’’’’’’’’’’’’’’’’’’’’’’’’
	'<Composed By : Shweta B Nagaral>
	'<Exports EXCEL, HTML, CSV from Design tab or Menu Bar>
	'<export: Type of Export to perform: String>
	'<exportSheet: SheetName to be exported if any: String>
	'<exportFormat: Export Format. It could be Dashboard or Component: String>
	'<dataRange: Specify Datarange based on type of exports : String>
	'<cellStyle: Specify cellStyle based on type of exports: String>
	'<chkLeadingSpace: Specify chkLeadingSpace based on type of exports: String>
	'<excelOption: Applicable only for excel Exports. Specify Excel options, It could be "Excel 97"/"Excel 2007" : String>
	'<objData: Reference to objData>
	Public Sub ExportOptions(ByVal export, ByVal exportSheet, ByVal exportFormat, ByVal dataRange, ByVal cellStyle, ByVal chkLeadingSpace, ByVal excelOption, ByVal objData)
    
	    Dim objExport, objExcelFormat, objExcelSelection, objChkLeadingSpace, objExportCStyle, objExportDataRange, objDelimiter, objChkColumnHeader, objTextQualifier
	    Dim objFormatComp, objDdlExcelSheets
    
    	If Browser("Analyzer").Page("ReportCreation").Frame("__dialog").Exist(120) = False Then
    		Call ReportStep (StatusTypes.Fail, "Export Options Dialog didnt appear","GroupCreation Dialouge box")
    		Exit Sub
    	End If
    	
        Select Case export
            
            Case "Excel"
                    'For Excel Export Option from Drop Down Design Tab
                    Set objExcelFormat=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebList("lstddlExcelFormat")
                    Call SCA.SelectFromDropdown(objExcelFormat,export)
                    
                    Set objExportDataRange=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebRadioGroup("rdoExportDataRange")
                    Call SCA.ClickOnRadio(objExportDataRange,dataRange, "ReportCreation")
                    
                    Set objExportCStyle=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebRadioGroup("rdoExportCStyle")
                    Call SCA.ClickOnRadio(objExportCStyle,cellStyle, "ReportCreation")
                    
                    Set objChkLeadingSpace=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebCheckBox("chkLeadingSpace")
                    Call SCA.SetCheckBox(objChkLeadingSpace, "chkLeadingSpace", chkLeadingSpace, "Excel Export Options Page")
    
                    Set objExcelSelection=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebRadioGroup("rdoExcelSelection")
                    Call SCA.ClickOnRadio(objExcelSelection,excelOption, "ReportCreation")
                    
            Case "MenuExcel"
                    'For Excel Export Option on Menu bar    
                    Set objDdlExcelSheets=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebList("lstDdlExcelSheets")
                    Call SCA.SelectFromDropdown(objDdlExcelSheets,exportSheet)
                     
                    Set objFormatComp = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebRadioGroup("rdoExcelFormat")
                    Call SCA.ClickOnRadio(objFormatComp,exportFormat, "ReportCreation")
                    
                    Set objExportDataRange=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebRadioGroup("rdoExportDataRange")
                    Call SCA.ClickOnRadio(objExportDataRange,dataRange, "ReportCreation")
                    
                    Set objExportCStyle=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebRadioGroup("rdoExportCStyle")
                    Call SCA.ClickOnRadio(objExportCStyle,cellStyle, "ReportCreation")
                    
                    Set objChkLeadingSpace=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebCheckBox("chkLeadingSpace")
                    Call SCA.SetCheckBox(objChkLeadingSpace, "chkLeadingSpace", chkLeadingSpace, "Export Options Page")
    
                    Set objExcelSelection=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebRadioGroup("rdoExcelSelection")
                    Call SCA.ClickOnRadio(objExcelSelection,excelOption, "ReportCreation")
                    
            Case "HTML"
                    'For HTML Export Option from Drop Down Design Tab
                    Set objExcelFormat=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebList("lstddlExcelFormat")
                    Call SCA.SelectFromDropdown(objExcelFormat,export)
                    
                    Set objExportDataRange=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebRadioGroup("rdoExportDataRange")
                    Call SCA.ClickOnRadio(objExportDataRange,dataRange, "ReportCreation")
                    
                    Set objExportCStyle=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebRadioGroup("rdoExportCStyle")
                    Call SCA.ClickOnRadio(objExportCStyle,cellStyle, "ReportCreation")
            
            Case "CSV"
                    'For CSV Export Option from Drop Down Design Tab
                    Set objExcelFormat=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebList("lstddlExcelFormat")
                    Call SCA.SelectFromDropdown(objExcelFormat,export)
                    
                    Set objDelimiter=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebRadioGroup("rdoDelimiter")
                    Call SCA.ClickOnRadio(objDelimiter,dataRange, "ReportCreation")
                
                    Set objTextQualifier=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebRadioGroup("rdoTextQualifier")
                    Call SCA.ClickOnRadio(objTextQualifier,cellStyle, "ReportCreation")
                
                    Set objChkColumnHeader=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebCheckBox("chkColumnHeader")
                    Call SCA.SetCheckBox(objChkColumnHeader, "objChkColumnHeader", chkLeadingSpace, "CSV Export Options Page")
            
        End Select
            
        'If User wants to exports, clock on export button
        Set objExport=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebButton("btnExport")
        Call SCA.ClickOn(objExport,"ExportButton" , "ReportCreation")
            
    End Sub


	'<Composed By : Shweta B Nagaral>
	'<Adds all members for a selected attribute when creating its group in Group Table>
	'<strName: Name of Group: String>
	'<strCaption: Caption for group: String>
	'<chkExpandable: Expandable option used to expand group: String>
	'<rdoDisplayAs: Provides options to either display as Group/Individual Members: String>
	Public Sub AddAllMembers(ByVal strName, ByVal strCaption, ByVal chkExpandable, ByVal rdoDisplayAs)
    If Browser("Analyzer").Page("GroupCreationPage").Frame("__dialogAddAllMembers").Image("imgAddAllMembers").Exist(30) Then
    	 Browser("Analyzer").Page("GroupCreationPage").Frame("__dialogAddAllMembers").Image("imgAddAllMembers").Click   
    End If
   
    If strName <> "" Then
        Browser("Analyzer").Page("GroupCreationPage").Frame("__dialogAddAllMembers").WebEdit("txtName").Set strName
    End If
    Browser("Analyzer").Page("GroupCreationPage").Frame("__dialogAddAllMembers").WebEdit("txtCaption").Set strCaption
    Browser("Analyzer").Page("GroupCreationPage").Frame("__dialogAddAllMembers").WebCheckBox("chkExpandable").Set chkExpandable   
    Browser("Analyzer").Page("GroupCreationPage").Frame("__dialogAddAllMembers").WebRadioGroup("rdoDisplayAs").Select rdoDisplayAs
    Browser("Analyzer").Page("GroupCreationPage").Frame("__dialogAddAllMembers").WebButton("btnOK").Click
    Browser("Analyzer").Page("GroupCreationPage").Sync
    Browser("Analyzer").Page("GroupCreationPage").RefreshObject
    Browser("Analyzer").Page("GroupCreationPage").Frame("__dialogAddAllMembers").WebButton("btnApply").Click
    Browser("Analyzer").Page("GroupCreationPage").Sync
    Browser("Analyzer").Page("GroupCreationPage").RefreshObject
    
End Sub

	
	Public Sub OperationOn_Report(ByVal strOselection,ByVal strFolderName,ByVal strReportName,ByVal strPagename)

		Dim objFrame,FrameName,objFolderNameTxt,objReportnameTxt,objbutton,objwebHome
		
		wait 1
		Browser("Analyzer").Page("ReportCreation").Sync
		Browser("Analyzer").Page("ReportCreation").RefreshObject
		
	   Select Case strOselection

         Case "SaveReport"

                Set objFrame = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer")
				'FrameName = objFrame.GetROProperty("name")
				Call ReportToolBar("save.svg",objFrame)
				If Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebEdit("txtFolderName").Exist(5) Then
				Set objFolderNameTxt=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebEdit("txtFolderName")
				Call SCA.SetText(objFolderNameTxt," " , "Report Saving Folder Name", "Report Save Dialogue")
				Call SCA.SetText(objFolderNameTxt,strFolderName , "Report Saving Folder Name", "Report Save Dialogue")

                Set objReportnameTxt=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebEdit("txtReportName")
				Call SCA.SetText(objReportnameTxt," " , "Report Name", "Report Save Dialogue")
				Call SCA.SetText(objReportnameTxt,strReportName , "Report Name", "Report Save Dialogue") 				

                Set objbutton=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebButton("btnOK")
				Call SCA.ClickOn(objbutton,"Reportsave Button", "Report Save Dialogue")

				Set objwebHome=Browser("Analyzer").Page("Shared_Folder").WebElement("webHome")
				Call SCA.ClickOn(objwebHome,"HomeWebElement", "Report Creation Page ")

				Call IMSSCA.Validations.ReportCreation_Validation(strReportName,strFolderName,"ReportCreated Folder Page")	
					
				End If
				
                 
		 Case "ExportToexcel"

		 Case "MoveReport"

		 Case ""

	      
	   End Select
	   
	   	Browser("Analyzer").Page("ReportCreation").Sync
		Browser("Analyzer").Page("ReportCreation").RefreshObject

	End Sub


	'<Composed By : Shweta B Nagaral>
	'<Saves SCA Report in Excel format to local machine and returns local path where report has been saved
	'It validates Report Name, Sheet Name and Excel format in while saving SCA report. 
	'SCA Report name should be same as Excel file name>
	'<ExcelName: Name of the excel while saving: String>
	Public Function SaveExcelReport(ByVal ExcelName)
	
	Dim i,toolBarNameFormat,strReport,a,b,j,strReportName,strReportFormat,path,formatPos,myVar,objLstWinButton ,objbtnSave,objYes,k,objClose,objbtnCloseExport, midVal
	midVal = len(ExcelName)+5
	   ' Browser("Hosting Templates - Ops").WinObject("Notification").WinButton("SaveAsDropDownWinButton").Click

	    'Report saving
	      For i = 1 To 50 Step 1
	       If Browser("Hosting Templates - Ops").WinObject("Notification").WinButton("SaveAsDropDownWinButton").Exist(2) then
	       Exit For
	       End If
	      Next
	      wait 2                
	      toolBarNameFormat = Browser("Analyzer").WinObject("Notification bar").GetROProperty("text")
	      'Retreiving name and format of report in Notification tool bar
	      'strReport Splitting
	      strReport = mid(toolBarNameFormat, InStr(1, Trim(Ucase(toolBarNameFormat)), Trim(Ucase(ExcelName))), midVal)
	           a=Split(Trim(strReport), ".")
	           b= UBound(a)
	            For j = 0 To b Step 2
	              strReportName = a(j)
	              strReportFormat = "."&a(j+1)
	            Next
	                
	       path = Environment.Value("CurrDir")&"Exportresults\"&Trim(strReport)
	    
	       formatPos = InStr(Trim(Ucase(toolBarNameFormat)), Trim(Ucase(strReportFormat)))
	           If formatPos > 0  Then
	             myVar = Mid(toolBarNameFormat, formatPos, 5 )
	                     
	              If strReportFormat = ".xlsx" Then
	                If StrComp(Trim(UCase(strReportFormat)), Trim(Ucase(myVar))) = 0 Then
	                   Call ReportStep (StatusTypes.Pass, "Saving Excel Report " &strReportName& " with Report Format " &strReportFormat, "Export Options Page")
	                   ElseIf StrComp(Trim(UCase(strReportFormat)), Trim(Ucase(myVar))) = 1 Then  
	                       
	                   Call ReportStep (StatusTypes.Fail, "Saving Excel Report " &strReportName& " with Report Format " &strReportFormat, "Export Options Page")
	                 End If
	                                                
	                Else
	                    If StrComp(Trim(UCase(strReportFormat)), Trim(Ucase(myVar))) = 0 Then                   
	                    Call ReportStep (StatusTypes.Pass, "Saving Excel Report " &strReportName& " with Report Format " &strReportFormat, "Export Options Page")
	                    ElseIf StrComp(Trim(UCase(strReportFormat)), Trim(Ucase(myVar))) = -1 Then                   
	                    Call ReportStep (StatusTypes.Fail, "Saving Excel Report " &strReportName& " with Report Format " &strReportFormat, "Export Options Page")
	                    End If
	                    End If
	                                
	                ElseIf formatPos = 0 Then
	                   'Format Mismtach
	                   Call ReportStep (StatusTypes.Fail, "Saving Excel Report " &strReportName& " with Report Format " &strReportFormat, "Export Options Page")
	                   Call ReportStep (StatusTypes.Fail, "Format Mismtach. Expected Excel Report format is " &strReportFormat& " But Notification tool bar is showing --> "&toolBarNameFormat& " Report format", "Export Options Page")
	             End If
	               
	                'Shweta - <11/1/2016> - Click on SaveAsContextMenu Menu - Start 
					'	                Set objLstWinButton=Browser("Analyzer").WinObject("Notification bar").WinButton("lstWinButton")
					'	                Call SCA.ClickOn(objLstWinButton,"lstWinButton", "Excel Report Export Dialouge")
					'	                
					'	                'Browser("Analyzer").WinMenu("ContextMenu").Select "Save as"
					'					Browser("Analyzer").WinMenu("SaveAsContextMenu").WaitProperty "visible",True,10000
					'					Browser("Analyzer").WinMenu("SaveAsContextMenu").Select "Save as"

	                Set objLstWinButton=Browser("Hosting Templates - Ops").WinObject("Notification").WinButton("SaveAsDropDownWinButton")
	                intLoopCounter = 0
	                Do
	                	If intLoopCounter=15 Then
							Exit Do 
					    End If
					 	intLoopCounter=intLoopCounter+1
	                	Call SCA.ClickOn(objLstWinButton,"lstWinButton", "Excel Report Export Dialouge")
	                Loop Until Browser("Analyzer").WinMenu("SaveAsContextMenu").Exist(2) = true
	                
	                If intLoopCounter=15 Then
	                	Call ReportStep (StatusTypes.Fail, "Context menu of Windows Notification bar was not clicked", "Export Options Page")
	                End If
	                'Shweta - <11/1/2016> - Click on SaveAsContextMenu Menu - Start 
	                wait 2
	                'Browser("Analyzer").WinMenu("ContextMenu").Select "Save as"
					Browser("Analyzer").WinMenu("SaveAsContextMenu").Select "Save as"
					'Shweta - <11/1/2016> - Click on SaveAsContextMenu Menu - End
	                wait 2
	               
	               	'Shweta - <11/1/2016> - Click on Click on Windows Save Location Weblist - Start
	               	'Browser("Analyzer").Dialog("Save As").WinEdit("lnkFile name:").WaitProperty "visible",True,10000
	               	if Browser("Analyzer").Dialog("Save As").WinEdit("lnkFile name:").Exist(120) = False Then
	               		Call ReportStep (StatusTypes.Fail, "Didnt enter "&path&" location to save the excel file", "Export Options Page")
	               	End if
	               	
	               'Browser("Analyzer").Dialog("Save As").WinObject("Items View").WinList("Items View").Highlight
	               	'Browser("Analyzer").Dialog("Save As").WinObject("Items View").WinList("Items View").Click
	               	wait 2
	               	'Shweta - <11/1/2016> - Click on Click on Windows Save Location Weblist - End
	               	
	                Browser("Analyzer").Dialog("Save As").WinEdit("lnkFile name:").Set path
	              
	                wait 2
	                Set objbtnSave=Browser("Analyzer").Dialog("Save As").WinButton("btnSave")
	                objbtnSave.WaitProperty "visible",True,10000
	                Call SCA.ClickOn(objbtnSave,"btnSave", "Excel Report Export Dialouge")
	                wait 2
	                
	                IF Browser("Analyzer").Dialog("Save As").Dialog("Confirm Save As").WinButton("Yes").Exist(5) Then
	                    Set objYes=Browser("Analyzer").Dialog("Save As").Dialog("Confirm Save As").WinButton("Yes")
	                    Call SCA.ClickOn(objYes,"Yes", "Excel Report Export Dialouge")
	                    wait 2
	                End If
	                
	                For k = 1 To 3 Step 1
	                   If Browser("Analyzer").WinObject("Notification bar").WinButton("Close").Exist(1) Then
	                      Set objClose=Browser("Analyzer").WinObject("Notification bar").WinButton("Close")
	                       Call SCA.ClickOn(objClose,"Close", "Excel Report Export Dialouge, Notification Bar")
	                       Exit For
	                    End If
	                Next
	                
	                wait 2
	                Set objbtnCloseExport=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebButton("btnClose")
	                objbtnCloseExport.WaitProperty "Visible",True,10000
	                Call SCA.ClickOn(objbtnCloseExport,"CloseExportbtn", "Excel Report Export Dialouge")
	               
	              
	    SaveExcelReport=path
	End Function
	
	
	'<Composed By : Shweta B Nagaral/Shobha Shivarudraiah>
	'1) Validates Sheet Name and Excel format after saving SCA report to local machine. Sheet Name should be same as SCA Application Report Sheet Name.
	'2) Validates SCA Application Report table with respect to Excel Sheet Values 
	'<path: local machine Path where Excel Report is saved: String>
	'<sheetName: Sheet Name after saving SCA report to local machine. Sheet Name should be same as SCA Application Report Sheet Name.: String>
	Public Sub ReportValidation_Excel(ByVal path, ByVal sheetName)
	Dim fso,objExcel,ObjExcelFile,ObjExcelSheet,ExcelSheetName
	Dim  intRow_Count,intColumn_Count,objGrpTabColData,rcAttribute,ccAttribute,p,q,x,y,strexcelcellvalue,strcellvalue
	Dim  objGrpTabMeasure,m,n,s,t,ccMeasure,rcMeasure ,toolBarNameFormat
	                
	
	               
	               'Vaildating name and format of report in Local Machine Path
	                Set fso = CreateObject("Scripting.FileSystemObject")
	                If (fso.FileExists(path)) Then
	                   Call ReportStep (StatusTypes.Pass, path & " exists.", "Local Machine Path")
	                   Call ReportStep (StatusTypes.Pass, "Successfully found sved SCA application Excel Report " &toolBarNameFormat, "Local Machine Path")
	                
	                Else
	                   Call ReportStep (StatusTypes.Fail, path & " doesn't exist.", "Local Machine Path")
	                   Call ReportStep (StatusTypes.Fail, "Could not find SCA application Excel Report " &toolBarNameFormat& " successfully", "Local Machine Path")
	                End If
	                                     
	                'Excel content validation wrt application group table data - Start
	                Set objExcel = CreateObject("Excel.Application") 
	                objExcel.Visible =False  
	                Set ObjExcelFile = objExcel.Workbooks.Open(path)
	                Set ObjExcelSheet = ObjExcelFile.Sheets(sheetName)
	                                
	                'Vaildating Excel Sheet Name wrt Application Report Creatio Page Sheet Name. 
	                'Get the Sheet Name of the Excel workbook
	                ExcelSheetName = ObjExcelSheet.Name
	                If StrComp(sheetName,ExcelSheetName) = 0 Then         
	                  'msgbox "SheetName Same"
	                  Call ReportStep (StatusTypes.Pass, "Vaildating Excel Sheet Name wrt Application Report Creation Page Sheet Name", "")
	                  Call ReportStep (StatusTypes.Pass, "Excel Sheet Name is same as Application Report Creation Page Sheet Name " &sheetName, "")
	                Else
	                  Call ReportStep (StatusTypes.Fail, "Vaildating Excel Sheet Name wrt Application Report Creation Page Sheet Name", "")
	                  Call ReportStep (StatusTypes.Fail, "Excel Sheet Name is diffrent wrt Application Report Creation Page Sheet Name " &sheetName, "")            
	                End If
	               
	                intRow_Count = ObjExcelSheet.UsedRange.Rows.Count 
	                intColumn_Count = ObjExcelSheet.UsedRange.Columns.Count
	                
	                'Validating Excel Sheet data and Group Table Dimensions Values
	                 Set objGrpTabColData=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrpTableColAttributeData")
	                 rcAttribute = objGrpTabColData.GetROProperty("rows")
	                 ccAttribute = objGrpTabColData.GetROProperty("cols")
	                   'p=5
	                   p=6
	                   q=1
	                   For x = 1 To rcAttribute-2
	                   For y = 1 To ccAttribute 
	                              
	                    If x=1 Then
	                       strexcelcellvalue=ObjExcelSheet.cells(p,q).value
	                    strcellvalue=objGrpTabColData.GetCellData(x,y)     
	                      
	                    ElseIf x<>1 Then
	    
	                    If  y<2 AND x<>1 Then
	                   ' q=ccAttribute
	                    strexcelcellvalue=ObjExcelSheet.cells(p,q).value
	                    
	                    If strexcelcellvalue="" Then
	                    q=ccAttribute
	                    strexcelcellvalue=ObjExcelSheet.cells(p,q).value
	                    	
	                    End If
	                    
	                    strcellvalue=objGrpTabColData.GetCellData(x,y) 
	                    End If
	                    End If
	                   
	                   
	                   
	                    If Strcomp(TRIM(strcellvalue),TRIM(strexcelcellvalue))=0 Then
	                    Call ReportStep (StatusTypes.Pass,"Validation of Report Exporting to Excel:-"&Space(5)&"Report Exported Successfully and Validated for the Dimension"&Space(3)&strexcelcellvalue,"Report Creation Page")
	                    Else
	                    Call ReportStep (StatusTypes.Fail,"Validation of Report Exporting to Excel:-"&Space(5)&"Report not Exported Successfully and Validated for the Dimension"&Space(3)&strexcelcellvalue,"Report Creation Page")                
	                    End If
	                    q=q+1
	                    Next
	                    q=1
	                    p=p+1   
	                    Next
	                 
	                 'Validating Excel Sheet data and Group Table measure values
	                 Set objGrpTabMeasure = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("class:=pivot_table","name:=cube").ChildItem(2,2,"WebTable",0)
	                 'Set objGrpTabMeasure=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrpTableMeasureData")
	                 rcMeasure = objGrpTabMeasure.GetROProperty("rows")
	                 ccMeasure = objGrpTabMeasure.GetROProperty("cols")
	                    
	                 m = 5
	                 n = ccAttribute+1
	                  For s = 1 To rcMeasure
	                  For t = 1 To ccMeasure 
	                     Browser("Analyzer").Page("ReportCreation").Sync
	                     Browser("Analyzer").Page("ReportCreation").RefreshObject
	                     strcellvalue=objGrpTabMeasure.GetCellData(s,t)
	                     strexcelcellvalue=ObjExcelSheet.cells(m,n).value
	                                                               
	                       If strcellvalue<>"-" AND strexcelcellvalue<>"-" Then
	                         strcellvalue=Round(strcellvalue,2)
	                         strexcelcellvalue=Round(strexcelcellvalue,2)
	                                                                    
	                         If Strcomp(TRIM(strcellvalue),TRIM(strexcelcellvalue))=0 Then
	                            Call ReportStep (StatusTypes.Pass,"Validation of Report Exporting to Excel:-"&Space(5)&"Report Exported Successfully and Validated for the Measure"&Space(3)&strexcelcellvalue,"Report Creation Page")
	                            'Reporter.ReportEvent micPass,"",""
	                          Else
	                            Call ReportStep (StatusTypes.Fail,"Validation of Report Exporting to Excel:-"&Space(5)&"Report not Exported Successfully and Validated for the Measure"&Space(3)&strexcelcellvalue,"Report Creation Page")                
	                            ' Reporter.ReportEvent  micFail,"",""
	                         End If
	                                
	                         End If
	                         n= n + 1                                
	                         
	                         Next
	'                         If n = ccMeasure Then
	'                          n = ccAttribute+1
	'                         End If
	                          n = ccAttribute+1    
	                         m = m + 1
	                         Next
	                
	                                
	                Browser("Analyzer").Page("ReportCreation").Sync
	                Browser("Analyzer").Page("ReportCreation").RefreshObject
	                wait 2
	
	                ObjExcelFile.Saved = True
	                ObjExcelFile.Close
	                
	                objExcel.Quit
	                
	                Set ObjExcelSheet = Nothing
	                Set ObjExcelFile = Nothing
	                Set objExcel = Nothing
	End Sub



	'< To Select the Menu Items of the Tool Bar>
    '<strFileName: File NAme of the Object to be Clicked:  String>
    '<objFrame: Frame Name where Object is located:  String>
    Public Sub ReportToolBar(ByVal strFileName, ByVal objFrame)
        wait 3
        Dim obj_Image, objchildImage 
        Set obj_Image=Description.Create
        obj_Image("micclass").Value="Image"
        obj_Image("file name").value=strFileName        
        
        'Set objchildImage=objFrame.ChildObjects(obj_Image)
        Set objchildImage=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_Image)
        objchildImage(0).click

    End Sub
    
'==================================================================================================================================================    
    
	  '<Login to SCA>
	  '<strUserName: UserName  String>
	  '<StrPwd: Password  String>
'==================================================================================================================================================
	  Public Sub SCALogin(ByVal strUserName,ByVal StrPwd, ByVal DCSiteLogin)
	  
	   Dim objUserName,objPwd,objLoginBtn,objSCAHomePageImage,returnval,objSCAHomeFrame

		SystemUtil.CloseProcessByName("iexplore.exe")
		systemutil.Run "iexplore.exe"
		wait 4
		Set WshShell=createobject("Wscript.shell")
		WshShell.SendKeys "^+{DEL}"
		wait 5
		WshShell.SendKeys "ENTER"
		wait 2
		Systemutil.CloseProcessByName "iexplore.exe"
		wait 4
		'<Shweta - 26/11/2015> Redirecting to SCA through DC site - Start
		If DCSiteLogin = 1 Then
			'Navigate to SCA through DC SIte IMS Analyzer link in DC Home page Analysis Table
			'msgbox StrPwd
			Call IMSSCA.General.DCLogin(strUserName, StrPwd)
			wait 2
			'Click on IMS Analysis Manager link in DC Home Page
			Set objlnkIMSAnalysisManager = Browser("DC Home").Page("Home").Link("lnkIMS Analysis Manager")
			Call SCA.ClickOn(objlnkIMSAnalysisManager,"lnkIAM Analysis Manager", "DC Home Page")
		
			wait 2
			Browser("DC Home").Page("IMS Analysis Manager").Sync
			Browser("DC Home").Page("IMS Analysis Manager").RefreshObject
			
			Set objlnkSharedReports = Browser("DC Home").Page("IMS Analysis Manager").Frame("Frame").Link("lnkShared Reports")
			
		else
'			SystemUtil.Run Environment.Value("ScaURL"),,,,3
			SystemUtil.Run "iexplore.exe",Environment.Value("ScaURL")
			
			wait 2
			Browser("Analyzer").Page("Home").Sync
			
			Set objUserName=Browser("micclass:=browser").Page("micclass:=page").WebEdit("xpath:=//input[@id='txtUserID']")
		    Call SCA.SetText(objUserName,strUserName,"textUserName" ,"Home Page" )
		   	'click on continue button
		   	wait 5
			Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//input[@id='btnValidate']").Click
			wait 4
		   	Set objPwd=Browser("micclass:=browser").Page("micclass:=page").WebEdit("xpath:=//input[@id='txtPassword']")
		   	Call SCA.SetText(objPwd,StrPwd,"txtpassword" ,"Home Page" )	
		   
		   	wait 2
		   	Set objLoginBtn=Browser("Analyzer").Page("Home").WebButton("btnLogin")
		   	Call SCA.ClickOn(objLoginBtn,"Login Button","Home Page")
		End If
		'<Shweta - 26/11/2015> Redirecting to SCA through DC site - End
		
		wait 2
		Browser("Analyzer").Page("Home").Sync
		Browser("Analyzer").Page("Home").RefreshObject
	   
	   For i=1 = 1 To 100 Step 1
	   	
	   If Browser("Analyzer").Page("Home").Frame("Frame").Link("lnkSharedReports").Exist(0) Then
	   	
	   	  
	   Set objSCAHomePageImage=Browser("Analyzer").Page("Home").Frame("Frame").Link("lnkSharedReports")
	   	Exit for
	   	
	   End If	
	   	
	   	
	   Next
	   
	   
	   
	   Set objSCAHomePageImage=Browser("Analyzer").Page("Home").Frame("Frame").Link("lnkSharedReports")
	   Set objSCAHomeFrame=Browser("Analyzer").Page("Home").Frame("Frame")
	   wait 3.5
	   
	   returnval=IMSSCA.Validations.ExistanceWebelements_Verification(objSCAHomePageImage, "lnkSharedReports", "Home Page", 0)
	   
	   If returnval=True Then
	   	Call ReportStep (StatusTypes.Pass,"Login is Successfull for SCA ","Home Page" )
		Else
		Call ReportStep (StatusTypes.Fail,"Login is notSuccessfull for SCA" ,"Home Page" )
	   End If
	    	
	  End Sub	
  

  '< Logs Out from SCA>
  '<strPageName: Page name of SCA application>
  Public Sub SCALogOut(ByVal strPageName)
  
   Dim objLogOut, objLeavepage
   
   Set objLogOut=Browser("Analyzer").Page("Home").WebElement("welLogout")
   objLogOut.Click
   'Call SCA.ClickOn(objLogOut,"welLogout" , strPageName)
   
   If Browser("Analyzer").Dialog("Windows Internet Explorer").Exist(5) Then
			Set objLeavepage=Browser("Analyzer").Dialog("Windows Internet Explorer").WinButton("btnLeave this page")
   			'Call SCA.ClickOn(objLeavepage,"Windows Internet Explorer Leave this page button" , strPageName)
			Browser("Analyzer").Dialog("Windows Internet Explorer").WinButton("btnLeave this page").Click
			wait 1
   End If

	SystemUtil.CloseProcessByName "iexplore.exe"

  End Sub




	'<Clicks on Pivot Table menu to create Report components>
    '<strTabName: Tab name in pivot table menu. It could be "Design", "Analyze", "Advanced"  :  String>
    '<strSelValue: Select options under each tab :  String>
    '<strSelSubValue: Select sub-options inside options under each tab :  String>
	
	
	
	'<Clicks on Pivot Table menu to create Report components>
    '<strTabName: Tab name in pivot table menu. It could be "Design", "Analyze", "Advanced"  :  String>
    '<strSelValue: Select options under each tab :  String>
    '<strSelSubValue: Select sub-options inside options under each tab :  String>
	Public Sub ComponentDownNormalMenu(ByVal compTableName, ByVal strTabName, ByVal strSelValue, ByVal strSelSubValue)
	
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
		wait 2
		'Mouse over on Pivot Table menu
		Set objimgDown=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgGroupTab_down_normal")
		'		Set objimgDown=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgPivotDownNormal")
		For i = 1 To 10 Step 1
		
			'Shweta <28/12/2016> - added Higlight and firevents on 'tabDesignStepGrpTable' - Start
			Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable(compTableName).Highlight
			Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable(compTableName).FireEvent "onmouseover"
			Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable(compTableName).FireEvent "onclick"
			'Shweta <28/12/2016> - added Higlight and firevents on 'tabDesignStepGrpTable' - End

			objimgDown.FireEvent "onmouseover"
			If objimgDown.Exist(5) Then
				objimgDown.Click
				Exit For
			End If
			wait 2
		Next
		wait 2
		'Browser("Analyzer").Page("ReportCreation").Sync
		
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
   		
	End Sub
	
	
	'< Searchs for Report created. In the Report Working Area, open the folder where the report is saved.Selects the options from menu.>
	'<strOselection: Selection of report property: String>
	'<strReportName: Report Name whose properties has to be accessed: String>
	'<strPageName: Page name where Report Properties should be accessed>
	'<objData: Reference to objData>
	Public Sub ReportFunctions(ByVal strOselection, ByVal strReportName, ByVal strPagename, ByRef objData)

		Dim objlink, objlinks, i, strlinkName, objReportTab, objReportTabCnt, objSetDefaultReport, objExecute, objConfig, objDelete, objDeleteOk
		
		Browser("Analyzer").Page("Shared_Folder").Sync
		Browser("Analyzer").Page("Shared_Folder").Frame("view").RefreshObject
		wait 2
		
		'Descriptive program to search for report created.
		Set objlink=description.Create
		objlink("micclass").Value="Link"
		objlink("name").value=strReportName
		
		Set objlinks=Browser("Analyzer").Page("Shared_Folder").Frame("view").ChildObjects(objlink)
		
		For i = 0 To 100 Step 1
			If objlinks.count = 0 Then
				Wait 2
			Else
				Exit for	
			End If
			
		Next
		
		For i=0 to objlinks.count-1
			wait 1
			Browser("Analyzer").Page("Shared_Folder").Frame("view").RefreshObject
			strlinkName=objlinks(i).GetROProperty("innertext")
			wait 1
			Browser("Analyzer").Page("Shared_Folder").Frame("view").RefreshObject
			If strcomp(strlinkName,strReportName)=0 Then		
				'If report link is found, Right-click on the required report to select options from menu
				Setting.WebPackage("ReplayType") = 2
				objlinks(i).RightClick
				Setting.WebPackage("ReplayType") = 1
				Exit For 
			End If
		Next
		
		If strcomp(strlinkName,strReportName)=0 Then		
			Call ReportStep (StatusTypes.Pass, +strReportName+Space(4)+"Report found Successfully", strPageName)
			Else
			Call ReportStep (StatusTypes.Fail, +strReportName+Space(4)+"Report is Not found Successfully", strPageName)   		
		End If
			
		'Select the report options from menu
	   	Select Case strOselection

         Case "Config"
           		'Select the "Config" option from the menu.
				Set objConfig = Browser("Analyzer").Page("Shared_Folder").Frame("view").WebElement("welConfig")
				Call SCA.ClickOn(objConfig,"welConfig", strPagename)
				
				wait 2
				Browser("Analyzer").Page("Shared_Folder").Sync
				
				'Call ConfigGeneralTab to set fields in General tab
				'Call ConfigSecurityTab to set fields in Security tab
				'TODO BookMarks 
				
		 Case "SetDefaultReport"
		 		'Select the "Set as default Report" option from the menu.
		 		Set objSetDefaultReport=Browser("Analyzer").Page("Shared_Folder").Frame("view").WebElement("welSetDefaultReport")
		 		Call SCA.ClickOn(objSetDefaultReport,"welSetDefaultReport", strPagename)
				wait 5
				Browser("Analyzer").Page("Shared_Folder").Sync
				wait 2
				
		 Case "Execute"
		 		'Select the "Execute" option from the menu.
		 		wait 2
		 		Set objExecute = Browser("Analyzer").Page("Shared_Folder").Frame("view").WebElement("welExecute")
		 		'Msgbox "here"
		 		'Set objExecute=Browser("micclass:=browser").page("micclass:=page").Frame("html id:=view").WebElement("xpath:=//span[text()='Execute']")
		 		Call SCA.ClickOn(objExecute,"welExecute", strPagename)
				wait 2
				Browser("Analyzer").Page("ReportCreation").Sync
				Browser("Analyzer").Page("Shared_Folder").RefreshObject
		
		'<13/5/2016>-Added create shortcut case - start
		 Case "CreateShortcut"
		 		wait 2
		 		'Select the "Create Shortcute" option from the menu.
		 		'Set objExecute = Browser("Analyzer").Page("Shared_Folder").Frame("view").WebElement("welExecute")
		 		'Set objCreateShortcut = Browser("Analyzer").Page("Shared_Folder").Frame("view").WebElement("welCreateShortcut")
		 		
		 		'Solving by descriptive programming
		 		Set objCreateShortcut=Browser("micclass:=browser").Page("micclass:=page").Frame("html id:=view").WebElement("html id:=CreateShortcut")
		 		
		 		Call SCA.ClickOn(objCreateShortcut,"welExecute", strPagename)
				wait 2
				Browser("Analyzer").Page("ReportCreation").Sync
				Browser("Analyzer").Page("Shared_Folder").RefreshObject
		'<13/5/2016>-Added create shortcut case - end
				
		 Case "Delete"
		 		'Select the "Execute" option from the menu.
		 		Set objDelete = Browser("Analyzer").Page("Shared_Folder").Frame("view").WebElement("webDelete")
		 		Call SCA.ClickOn(objDelete,"webDelete", strPagename)
		 		
				Browser("Analyzer").Page("ReportCreation").Sync
				If Browser("Analyzer").Dialog("Message from webpage").Static("Move this object to Recycle").Exist(10) Then
					Set objDeleteOk=Browser("Analyzer").Dialog("Message from webpage").WinButton("btnOK")
					Call SCA.ClickOn(objDeleteOk,"Move this object to Recycle" , strPageName)
					Call ReportStep (StatusTypes.Pass, "Successfully Deleted Report "&strReportName , strPagename)
				Else
					Call ReportStep (StatusTypes.Fail, "Report "&strReportName&" is not deleted" , strPagename) 
				End If
				
				Browser("Analyzer").Page("Shared_Folder").Sync
				Browser("Analyzer").Page("Shared_Folder").RefreshObject
		
	   End Select
		wait 2
		Browser("Analyzer").Page("ReportCreation").Sync
		Browser("Analyzer").Page("Shared_Folder").RefreshObject
	End Sub


	'<Configures general tab for report created>
	'<strConfigReportName : Configuration Report Name: String>
	'<strConfigReportDescription: Configuration Report Description: String>
	'<strReportName : ReportName for which general tab has to be configured>
	'<strPagename : Page name where Report Properties should be accessed>
	'<objData : Reference to objData>
	Public Sub ConfigGeneralTab(ByVal strConfigReportName, ByVal strConfigReportDescription,  ByVal strReportName, ByVal strPagename, ByRef objData)
				
				Dim objConfigGeneralTab, objConfigReportName, objConfigReportDescription, objbutton
				
				'Set Parameters in General Tab
				Set objConfigGeneralTab = Browser("Analyzer").Page("Home").Frame("Frame").WebElement("welConfigGeneralTab")
				Call SCA.ClickOn(objConfigGeneralTab,"welConfigGeneralTab", strPagename)
				
				set objConfigReportName=Browser("Analyzer").Page("Home").Frame("Frame").WebEdit("txtConfigReportName")
				Call SCA.SetText(objConfigReportName, strConfigReportName, "txtConfigReportName", strPagename)
				
				wait 2
				
				set objConfigReportDescription =Browser("Analyzer").Page("Home").Frame("Frame").WebEdit("txtConfigReportDescription")
				Call SCA.SetText_MultipleLineArea(objConfigReportDescription, strConfigReportDescription, "txtConfigReportDescription", strPagename)
				
				'Click on OK button in Config Report Page
				Set objbutton = Browser("Analyzer").Page("Home").Frame("Frame").WebButton("btnFolderOk")
				Call SCA.ClickOn(objbutton,"btnFolderOk", strPagename)
				
				Browser("Analyzer").Page("Home").Sync
		
	End Sub
	
	
   	'<configures Security tab for report created>
    '<strMemberOper : Configure report created by adding or removing members/users: String>
    '<strMember : Member to be configured for report created: String>
    '<strReportName : ReportName for which general tab has to be configured>
    '<strPagename : Page name where Report Properties should be accessed>
    '<objData : Reference to objData>
    Public Sub ConfigSecurityTab(ByVal strMemberOper, ByVal strMember, ByVal strReportName, ByVal strPagename, ByRef objData)
                Dim i, j, k , objAddMember, objRemoveMember, objConfigSecurityTab, showUserRow, strShowMemberName, strShowMemberType, selectedMembersRow, strUsername, strUserType, memberAdded, objShowUser,objUserSearch,objImgAddMember, objAddbutton, objConfigUser, objCheckUsersAll, objNextPage
                memberAdded = 0
                'Check User1 is able to perform below operations on selected report
                'i.e: List, Execute, Export, Config, Write, Delete Permission.
                
                'Click on Security tab and Set Parameters in Security Tab
                Set objConfigSecurityTab = Browser("Analyzer").Page("Home").Frame("Frame").WebElement("welConfigSecurityTab")
                Call SCA.ClickOn(objConfigSecurityTab,"welConfigSecurityTab", strPagename)
                
                Browser("Analyzer").Page("Home").Sync
                Browser("Analyzer").Page("Home").RefreshObject
                
                'Click on Add/Remove memeber User1 and check whether User1 has selected permission for report.
                Select Case strMemberOper
                    Case "Add"
                        Set objAddMember = Browser("Analyzer").Page("Home").Frame("Frame").Image("imgAdd")
                        Call SCA.ClickOn(objAddMember,"imgAdd",strPagename)
                                                
                        'Search for User to grant/deny persmissions for report
                        For i = 1 To 10
                            set objShowUser = Browser("Analyzer").Page("Home").Frame("Frame").Image("imgShowUser")
                            Call SCA.ClickOn(objShowUser,"imgShowUser",strPagename)
                            Browser("Analyzer").Page("Home").Sync
                            Browser("Analyzer").Page("Home").Sync
                            wait 2
                            Browser("Analyzer").Page("Home").Frame("Frame").WebEdit("txtConfigUser").FireEvent "onmouseover"
                            Browser("Analyzer").Page("Home").Frame("Frame").WebEdit("txtConfigUser").FireEvent "ondblclick"
                            set objConfigUser=Browser("Analyzer").Page("Home").Frame("Frame").WebEdit("txtConfigUser")
                            Call SCA.SetText(objConfigUser, strMember, "txtConfigUser", strPagename)
                            wait 2
                            Browser("Analyzer").Page("Home").Sync
                            set objUserSearch = Browser("Analyzer").Page("Home").Frame("Frame").Image("imgUserSearch")
                            Call SCA.ClickOn(objUserSearch,"imgUserSearch",strPagename)
                            wait 2
                            Browser("Analyzer").Page("Home").Sync
                        
                            showUserRow=Browser("Analyzer").Page("Home").Frame("Frame").WebTable("tabShowUser").GetROProperty("rows")
                                If showUserRow >=2 Then
                                    wait 2
                                    strShowMemberName = Browser("Analyzer").Page("Home").Frame("Frame").WebTable("tabShowUser").GetCellData(2, 2)
                                    strShowMemberType = Browser("Analyzer").Page("Home").Frame("Frame").WebTable("tabShowUser").GetCellData(2, 3)
                                    Set objCheckUsersAll=Browser("Analyzer").Page("Home").Frame("Frame").WebCheckBox("chkCheckUsersAll")
                                    Call SCA.SetCheckBox(objCheckUsersAll, "chkCheckUsersAll", "ON", strPagename)
                                    set objImgAddMember = Browser("Analyzer").Page("Home").Frame("Frame").Image("imgAddMember")
                                    Call SCA.ClickOn(objImgAddMember,"imgAddMember",strPagename)
                                    Browser("Analyzer").Page("Home").Sync
                                    'Click on "Add" button to add the selected User
                                    Set objAddbutton = Browser("Analyzer").Page("Home").Frame("Frame").WebButton("btnAdd")
                                    Call SCA.ClickOn(objAddbutton,"btnAdd", strPagename)                
                                    Exit For
                                End If
                        Next
                        
                        Browser("Analyzer").Page("Home").Sync
                        wait 2
 


'                'Removal of Do While Loop - Start
					j=1
                    Do
	                  	If Browser("Hosting Templates - Ops").Page("Analyzer").Frame("Frame").Image("next_enabled").Exist(10) AND j<>1 Then
		  	 				Browser("Hosting Templates - Ops").Page("Analyzer").Frame("Frame").Image("next_enabled").Click
		  				End If
                            Browser("Analyzer").Page("Home").Sync
                            Browser("Analyzer").Page("Home").RefreshObject
                            selectedMembersRow=Browser("Analyzer").Page("Home").Frame("Frame").WebTable("tabSelectedMembers").GetROProperty("rows")
                            For j = 1 To selectedMembersRow
                                strUsername = Browser("Analyzer").Page("Home").Frame("Frame").WebTable("tabSelectedMembers").GetCellData(j, 2)
                                strUserType = Browser("Analyzer").Page("Home").Frame("Frame").WebTable("tabSelectedMembers").GetCellData(j, 3)
                                'shweta <7/8/2015> - Start
                                'If Trim(UCase(strUsername)) = Trim(UCase(strShowMemberName)) and Trim(UCase(strUserType)) = Trim(UCase(strShowMemberType)) Then
                                '    Call ReportStep (StatusTypes.Pass, "Selected Memeber "&strMember& " successfully", strPageName)
                                '    memberAdded = 1
                                '    Exit For
                                'End If
                                
                                If InStr(1, strShowMemberName, strUsername) > 0 and InStr(1, strShowMemberType, strUserType) > 0 Then
                                    Call ReportStep (StatusTypes.Pass, "Selected Memeber "&strMember& " successfully", strPageName)
                                    memberAdded = 1
                                    Exit For
                                End If
                                'shweta <7/8/2015> - End
                            Next
                            
                            If memberAdded = 1 Then
                                    Exit Do
                            Elseif memberAdded = 0 Then
                                Set objNextPage = Browser("Analyzer").Page("Home").Frame("Frame").WebElement("welNextPage")
                                Call SCA.ClickOn(objNextPage,"welNextPage", strPagename)
                                Browser("Analyzer").Page("Home").Sync    
                            End If
                

                    	j=j+1
				Loop While Browser("Hosting Templates - Ops").Page("Analyzer").Frame("Frame").Image("next_enabled").Exist(10)
                    

'                        Do
'                            selectedMembersRow=Browser("Analyzer").Page("Home").Frame("Frame").WebTable("tabSelectedMembers").GetROProperty("rows")
'                            For j = 1 To selectedMembersRow
'                                strUsername = Browser("Analyzer").Page("Home").Frame("Frame").WebTable("tabSelectedMembers").GetCellData(j, 2)
'                                strUserType = Browser("Analyzer").Page("Home").Frame("Frame").WebTable("tabSelectedMembers").GetCellData(j, 3)
'                                If Trim(UCase(strUsername)) = Trim(UCase(strShowMemberName)) and Trim(UCase(strUserType)) = Trim(UCase(strShowMemberType)) Then
'                                    Call ReportStep (StatusTypes.Pass, "Selected Memeber "&strMember& " successfully", strPageName)
'                                    memberAdded = 1
'                                    Exit For
'                                End If
'                            Next
'                            
'                            If memberAdded = 0 Then
'                                Set objNextPage = Browser("Analyzer").Page("Home").Frame("Frame").WebElement("welNextPage")
'                                Call SCA.ClickOn(objNextPage,"welNextPage", strPagename)
'                                Browser("Analyzer").Page("Home").Sync    
'                            End If
'                        Loop While memberAdded = 0
'                'Removal of Do While Loop - End
                        If memberAdded = 0 Then
                                Call ReportStep (StatusTypes.Fail, "Could Not select member "&strMember& " successfully", strPageName)        
                        End If
                        'copy end
                                                
                    Case "Remove"
                        Set objRemoveMember = Browser("Analyzer").Page("Home").Frame("Frame").Image("imgRemove")
                        Call SCA.ClickOn(objRemoveMember,"imgRemove","Report Creation")
                End Select
                
'                Set objchkPermissions=Browser("Analyzer").Page("Home").Frame("Frame").WebCheckBox("chkPermissions")
'                Call SCA.SetCheckBox(objchkPermissions, "chkPermissions", objData.Item(""), strPagename)
        
    End Sub
    
    
	'<Composed By : Shweta B Nagaral>
	'<Closes component as required by User .>
	'<strComponent: Mention component name. Example: "PivotTable", "GroupTable": String>
	'<strPageName: Web Page name>
	'<objData: Reference to objData>
	Public Sub CloseComponent(ByVal strComponent, ByVal strPagename, ByRef objData)
		
		Dim oComp, oCompObj, oDelete, oDeleteObj, i, objButtonOK, strWarnMsg, compCounter
		compCounter = 0
		
		For i = 0 To 10 Step 1
			wait 1
			Browser("Analyzer").Page("ReportCreation").Sync
			Browser("Analyzer").Page("ReportCreation").RefreshObject
			wait 4
			'Descriptive programe to find Components
			Set oComp  = Description.Create()
			oComp("micclass").value = "WebTable"
			oComp("cols").value = 4
			oComp("column names").regularexpression=True
			oComp("column names").value = ";"&strComponent&".*;;"
			oComp("html tag").value = "TABLE"
			oComp("name").value = "down_normal"
			oComp("innertext").regularexpression=True
			oComp("innertext").value = strComponent&".*"
			
			Set oCompObj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabPivotTable").ChildObjects(oComp)
			'msgbox oCompObj.count
			
			If oCompObj.count = 0 Then
					Call ReportStep (StatusTypes.Pass, strComponent& " Component not available in the Report Page" , strPagename)
					Exit For 
			ElseIf oCompObj.count > 0 Then
					wait 1
					Browser("Analyzer").Page("ReportCreation").Sync
					Browser("Analyzer").Page("ReportCreation").RefreshObject
					
					'Descriptive programe to find delete button to close components
					Set oDelete  = Description.Create()
					oDelete("micclass").value = "Image"
					oDelete("file name").regularexpression=True
					oDelete("file name").value = "delete_normal\.gif"
					oDelete("html tag").value = "IMG"
					oDelete("image type").value = "Plain Image"
					oDelete("name").value = "Image"
					oDelete("title").value = "Delete"
				
					Set oDeleteObj = oCompObj(0).ChildObjects(oDelete)
					'msgbox oDeleteObj.count
					oDeleteObj(0).FireEvent "onmouseover"
					wait 1
					oDeleteObj(0).Click
					wait 2
					If Browser("Analyzer").Dialog("Message from webpage").Exist(3) then
						strWarnMsg = Browser("Analyzer").Dialog("Message from webpage").GetROProperty("innerText")
						Call ReportStep (StatusTypes.Pass, "WARNING MESSAGE: "&strWarnMsg , strPagename)
						Call ReportStep (StatusTypes.Pass, "Please add any other component before deleting only "&strComponent& " component available in the Report Page" , strPagename)
						Set objButtonOK = Browser("Analyzer").Dialog("Message from webpage").WinButton("btnOK")
						Call SCA.ClickOn(objButtonOK,"Dialog Ok Button", "ReportCreation Page")
						wait 1
						Exit For
					End If
					
					compCounter = compCounter+1
					Browser("Analyzer").Page("ReportCreation").Sync
					Browser("Analyzer").Page("ReportCreation").RefreshObject
			End If
			Browser("Analyzer").Page("ReportCreation").Sync
			Browser("Analyzer").Page("ReportCreation").RefreshObject
		Next
		
		If compCounter > 0 Then
			Call ReportStep (StatusTypes.Pass, "Successfully deleted "&compCounter&Space(2) &strComponent& " components available in the Report Page" , strPagename)
		End If
		
		Browser("Analyzer").Page("ReportCreation").Sync
		Browser("Analyzer").Page("ReportCreation").RefreshObject
	End Sub
	

	'<Composed By : Shweta B Nagaral>	
	'<Performs Design tab Operations like First Step, Next Step, Prior Step>
	'<stepType: Mention Design tab Operations to perform like First Step, Next Step, Prior Step: String>
	'<strPageName: Web Page name>
	'<objData: Reference to objData>
	Public Function GroupTableDesignSteps(ByVal stepType, ByVal strPagename, ByRef objData)
	
	Dim validate, objwebtable, objNewwebtable, ArrValue_Before, ArrValue_After, strCellValue_Before, strCellValues_Before, strWarnMsg, strCellValue_After, strCellValues_After, strPriorInsertMeasure, objButtonOK
	Dim retPriorInsertMeasure, retPriorExpandGroup, retPriorEditGroupItems, retDisabledNextStep, retNextInsertMeasure, retRowPriorExpandGroup, retRowPriorEditGroupItems, StepListOper, oDescStepList, objDescStepList, EditedGrp, grpTabVal, StepListArrValue_After, strStepListCellValue_After, strStepListCellValues_After, reVal
	Dim oGrpTabDownNormal, oGrpTabDownNormalObj, expandDisabled, expandMemberText, retBackToToplevel
	Dim i, j, k, l, m, n, p, q, o, r
	validate = 0
	
	Browser("Analyzer").Page("ReportCreation").Sync
	Browser("Analyzer").Page("ReportCreation").RefreshObject
	
	Call ReportStep (StatusTypes.Information, "Performing "&stepType&" Design Tab operation on Group Table", "Report Creation Page")
	Set objwebtable=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDesignStepGrpTable")
	ArrValue_Before=SCA.Webtable(objwebtable,"Dimension Value table",TRIM("Retriving_DataTableValue"),"Report Creation","","","","")
	For i=0 to ubound(ArrValue_Before,1)
	    For j=0 to ubound(ArrValue_Before,2)
	        strCellValue_Before= ArrValue_Before(i,j)
	        strCellValues_Before=strCellValues_Before&";"&strCellValue_Before
	    Next
	Next
	
	Browser("Analyzer").Page("ReportCreation").Sync
	Browser("Analyzer").Page("ReportCreation").RefreshObject
	wait 2
	
	Select Case stepType
		Case "First Step"

				'Click on First Step 
				'Call IMSSCA.General.DownNormalMenu(objData.item("strSelValue"),objData.item("MenuOption1"), "")
				Call IMSSCA.General.ComponentDownNormalMenu("tabDesignStepGrpTable", "imgDesign", "welFirstStep", "")
		
				For r = 1 To 10 Step 1
					'Click on Dialog to perfor m action on Group Table
					If Browser("Analyzer").Dialog("Message from webpage").Exist(2) then
						strWarnMsg = Browser("Analyzer").Dialog("Message from webpage").Static("DialogMessage").GetROProperty("text")
						'msgbox strWarnMsg
						Call ReportStep (StatusTypes.Pass, "Warning Message before performing First Step Function", strPagename)
						Call ReportStep (StatusTypes.Pass, "WARNING MESSAGE: "&strWarnMsg , strPagename)
						Set objButtonOK = Browser("Analyzer").Dialog("Message from webpage").WinButton("btnOK")
						Call SCA.ClickOn(objButtonOK,"Dialog Ok Button", "ReportCreation Page")
						Exit For
					End If
					
					If r = 10 Then
						Call ReportStep (StatusTypes.Fail, "Warning Message before performing First Step Function didnt appear", strPagename)
					End If
				Next
							
				wait 2	
				
				Browser("Analyzer").Page("ReportCreation").Sync
			    Browser("Analyzer").Page("ReportCreation").RefreshObject
	
				'Validate Existence of Empty Group Table 
				For k = 0 To 10 Step 1
					If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabEmptyGroupTable").Exist(2) Then
						validate = 1
						Exit For
					End If
				Next
	
				If validate = 1 Then
					GroupTableDesignSteps = 1
					'Msgbox "validated First Step Functionality of Group Table Design Tab. Working successfully"
				ElseIf validate = 0 Then	
					GroupTableDesignSteps = 0
				End If
	
		Case "Prior Step"
				
				'Step 1
				'Validate 'welPriorStep' Innertext value should be -> "Prior Step (Insert Measure)"
					retPriorInsertMeasure = IMSSCA.Validations.ValidateTabInnerText("imgGroupDownNormal", "imgDesign", "welPriorStep", "", "Prior Step (Insert Measure)")
					If retPriorInsertMeasure = 1 Then
						Call ReportStep (StatusTypes.Fail, "'Prior Step (Insert Measure)' not found", strPagename)
						Call ReportStep (StatusTypes.Fail, "Could not successfully validate Prior Step Functionality of Group Table Design Tab", strPagename)
						Exit Function
					End If
			   		
				'Step 2
				'Validate 'welPriorStep' Innertext value should be -> "Prior Step (Expand Group) for Column Attribute added"
					retPriorExpandGroup = IMSSCA.Validations.ValidateTabInnerText("imgGroupDownNormal", "imgDesign", "welPriorStep", "", "Prior Step (Expand Group)")
					If retPriorExpandGroup = 1 Then
						Call ReportStep (StatusTypes.Fail, "'Prior Step (Expand Group)' not found", strPagename)
						Call ReportStep (StatusTypes.Fail, "Could not successfully validate Prior Step Functionality of Group Table Design Tab", strPagename)
						Exit Function
					End If
									
				'Step 3
				'Validate 'welPriorStep' Innertext value should be -> "Prior Step (Edit Group Items) for Column Attribute added"
					retPriorEditGroupItems = IMSSCA.Validations.ValidateTabInnerText("imgGroupDownNormal", "imgDesign", "welPriorStep", "", "Prior Step (Edit Group Items)")
					If retPriorEditGroupItems = 1 Then
						Call ReportStep (StatusTypes.Fail, "'Prior Step (Edit Group Items)' not found", strPagename)
						Call ReportStep (StatusTypes.Fail, "Could not successfully validate Prior Step Functionality of Group Table Design Tab", strPagename)
						Exit Function
					End If
					
				'Step 4
				'Validate 'welPriorStep' Innertext value should be -> "Prior Step (Expand Group) for Row Attribute added"
					retRowPriorExpandGroup = IMSSCA.Validations.ValidateTabInnerText("imgGroupDownNormal", "imgDesign", "welPriorStep", "", "Prior Step (Expand Group)")
					If retRowPriorExpandGroup = 1 Then
						Call ReportStep (StatusTypes.Fail, "'Prior Step (Expand Group)' not found", strPagename)
						Call ReportStep (StatusTypes.Fail, "Could not successfully validate Prior Step Functionality of Group Table Design Tab", strPagename)
						Exit Function
					End If
									
				'Step 5
				'Validate 'welPriorStep' Innertext value should be -> "Prior Step (Edit Group Items) for Row Attribute added"
					retRowPriorEditGroupItems = IMSSCA.Validations.ValidateTabInnerText("imgGroupDownNormal", "imgDesign", "welPriorStep", "", "Prior Step (Edit Group Items)")
					If retRowPriorEditGroupItems = 1 Then
						Call ReportStep (StatusTypes.Fail, "'Prior Step (Edit Group Items)' not found", strPagename)
						Call ReportStep (StatusTypes.Fail, "Could not successfully validate Prior Step Functionality of Group Table Design Tab", strPagename)
						Exit Function
					End If
					
				'Validate Existence of Empty Group Table to confirm "Prior Step" Functionality
				For l = 0 To 10 Step 1
					Browser("Analyzer").Page("ReportCreation").Sync
					Browser("Analyzer").Page("ReportCreation").RefreshObject
					If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabEmptyGroupTable").Exist(2) Then
						validate = 1
						Exit For
					End If
				Next
	
				If validate = 1 Then
					'Validated Prior Step Functionality of Group Table Design Tab. Working successfully
					GroupTableDesignSteps = 1
				ElseIf validate = 0 Then	
					'Validated Prior Step Functionality of Group Table Design Tab. Not Working successfully
					GroupTableDesignSteps = 0
				End If
		
		Case "Next Step"
		
				'Step 1
				'Validate 'welNextStep' Innertext value should be -> "Next Step". "Next Step" option is disabled.
					'Mouse over on Pivot Table menu
					retDisabledNextStep = IMSSCA.Validations.ValidateTabInnerText("imgGroupDownNormal", "imgDesign", "welNextStep", "", "Next Step")
					If retDisabledNextStep = 1 Then
						Call ReportStep (StatusTypes.Fail, "'Next Step' should be disabled", strPagename)
						Call ReportStep (StatusTypes.Fail, "Validating Next Step Functionality of Group Table Design Tab", strPagename)
					ElseIf retDisabledNextStep = 0 Then
						Call ReportStep (StatusTypes.Pass, "'Next Step' option should be disabled", strPagename)
						Call ReportStep (StatusTypes.Pass, "'Next Step' is disabled for first time", strPagename)
					End If
	
				'Step 2
				'Validate 'welPriorStep' Innertext value should be -> "Prior Step (Insert Measure)"
					retPriorInsertMeasure = IMSSCA.Validations.ValidateTabInnerText("imgGroupDownNormal", "imgDesign", "welPriorStep", "", "Prior Step (Insert Measure)")
					If retPriorInsertMeasure = 1 Then
						Call ReportStep (StatusTypes.Fail, "'Prior Step (Insert Measure)' not found", strPagename)
						Call ReportStep (StatusTypes.Fail, "Could not successfully validate Next Step Functionality of Group Table Design Tab", strPagename)
						Exit Function
					End If
					
				'Step 3
				'Validate 'welNextStep Innertext value should be -> "Next Step (Insert Measure)"
					retNextInsertMeasure = IMSSCA.Validations.ValidateTabInnerText("imgGroupDownNormal", "imgDesign", "welNextStep", "", "Next Step (Insert Measure)")
					If retNextInsertMeasure = 1 Then
						Call ReportStep (StatusTypes.Fail, "'Next Step (Insert Measure)' not found", strPagename)
						Call ReportStep (StatusTypes.Fail, "Could not successfully validate Next Step Functionality of Group Table Design Tab", strPagename)
						Exit Function
					End If
					
				'Validate Existence of Complete Group Table Created to confirm "Next Step" Functionality
				For m = 0 To 10 Step 1
					Browser("Analyzer").Page("ReportCreation").Sync
					Browser("Analyzer").Page("ReportCreation").RefreshObject
					If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDesignStepGrpTable").Exist(2) Then
						validate = 1
						Exit For
					End If
				Next
				
				Browser("Analyzer").Page("ReportCreation").Sync
				Browser("Analyzer").Page("ReportCreation").RefreshObject
				
				If validate = 1 Then
					Browser("Analyzer").Page("ReportCreation").Sync
					Browser("Analyzer").Page("ReportCreation").RefreshObject
					Set objNewwebtable=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDesignStepGrpTable")
					ArrValue_After=SCA.Webtable(objNewwebtable,"Dimension Value table",TRIM("Retriving_DataTableValue"),"Report Creation","","","","")
					For m=0 to ubound(ArrValue_After,1)
					    For n=0 to ubound(ArrValue_After,2)
					        strCellValue_After= ArrValue_After(m,n)
					        strCellValues_After=strCellValues_After&";"&strCellValue_After
					    Next
					Next
					
					wait 2
					reVal=IMSSCA.Validations.FilterVerification(strCellValues_Before,strCellValues_After)
			    	If reVal = 0 Then
			    		Call ReportStep (StatusTypes.Pass, "Successfully Validated Group Table after performing Next Step Functionality. Same Original Created Group Table Contents are displayed", strPagename)
				      	Call ReportStep (StatusTypes.Pass, "Validated Next Step Functionality of Group Table Design Tab. Working successfully", strPagename)
				      	GroupTableDesignSteps = 1
						Else		
						Call ReportStep (StatusTypes.Fail, "Successfully Validated Group Table after performing Next Step Functionality. Group Table Contents are different", strPagename)
						GroupTableDesignSteps = 0
					End If    
				ElseIf validate = 0 Then	
					'Validated Next Step Functionality of Group Table Design Tab. Not Working successfully
					GroupTableDesignSteps = 0
				End If
				
				'"Next Step", "Next Step (Expand Group)"
				'Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welNextStep").Click
				'Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welNextStepInsertMeasure").Click
				'Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welPriorStep").Click
				'Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welPriorStepExpandGrp").Click
	
		Case "Back to top level"
				expandDisabled = 0
				wait 2
				Browser("Analyzer").Page("ReportCreation").Sync
				Browser("Analyzer").Page("ReportCreation").RefreshObject
				
				'By using "Expand" or "Drill down" option expand all the member's in the Group table - Start
				For o = 1 To 10 Step 1
					'Descriptive programe to click on Down Normal image of row attribute
					Set oGrpTabDownNormal  = Description.Create()
					oGrpTabDownNormal("micclass").value = "Image"
					oGrpTabDownNormal("file name").regularexpression=True
					oGrpTabDownNormal("file name").value = "down_normal.gif"
					oGrpTabDownNormal("html tag").value = "IMG"
					oGrpTabDownNormal("image type").value = "Plain Image"
					oGrpTabDownNormal("name").value = "Image"
					oGrpTabDownNormal("outerhtml").regularexpression=True
					oGrpTabDownNormal("outerhtml").value = ".*CURSOR: hand.*"
									
					Set oGrpTabDownNormalObj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(oGrpTabDownNormal)
					'msgbox oGrpTabDownNormalObj.count
					wait 2
					oGrpTabDownNormalObj(oGrpTabDownNormalObj.count-2).FireEvent "onmouseover"
					wait 2
					oGrpTabDownNormalObj(oGrpTabDownNormalObj.count-2).Click
					wait 2
					Browser("Analyzer").Page("ReportCreation").Sync
					Browser("Analyzer").Page("ReportCreation").RefreshObject
					expandMemberText=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welExpandMembers").GetROProperty("outerhtml")
					If InStr(1, expandMemberText, "lightgrey")>0 Then
						expandDisabled = 1
						Call ReportStep (StatusTypes.Pass, "Expand Members Option Disbled. Successfully expanded members to each level completely", strPagename)
						Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welExpandMembers").Click
						Browser("Analyzer").Page("ReportCreation").Sync
						Browser("Analyzer").Page("ReportCreation").RefreshObject
						Exit For
					Else
						Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welExpandMembers").Click
					End If
					
					wait 8
					If oGrpTabDownNormalObj.count >= 3 Then
						wait 3
						Browser("Analyzer").Page("ReportCreation").Sync
						Browser("Analyzer").Page("ReportCreation").RefreshObject
					End If
					Browser("Analyzer").Page("ReportCreation").Sync
					Browser("Analyzer").Page("ReportCreation").RefreshObject
					
					If expandDisabled = 0 and o = 10 Then
						Call ReportStep (StatusTypes.Fail, "Expand Members Option is still not Disbled. Could not expand members to each level completely", strPagename)
						Exit Function
					End If
				Next
				'By using "Expand" or "Drill down" option expand all the member's in the Group table - End
					
				wait 2
				Browser("Analyzer").Page("ReportCreation").Sync
				Browser("Analyzer").Page("ReportCreation").RefreshObject
				
				'Save table contents in strCellValues_After and validate "Back To Top Level"
				'Click on the Group table drop down menu , Click on option "Back to Top level". Verify that all the member's are collapsed and displayed.
				retBackToToplevel = IMSSCA.Validations.ValidateTabInnerText("imgGroupDownNormal", "imgDesign", "welBackToTopLevel", "", "Back to Top Level")
				If retBackToToplevel = 1 Then
					Call ReportStep (StatusTypes.Fail, "'Back to top level' not found", strPagename)
					Call ReportStep (StatusTypes.Fail, "Could not successfully validate 'Back to top level' Functionality of Group Table Design Tab", strPagename)
					Exit Function
				End If
				
				wait 2
				Browser("Analyzer").Page("ReportCreation").Sync
				Browser("Analyzer").Page("ReportCreation").RefreshObject
						
				Set objNewwebtable=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDesignStepGrpTable")
				ArrValue_After=SCA.Webtable(objNewwebtable,"Dimension Value table",TRIM("Retriving_DataTableValue"),"Report Creation","","","","")
				For m=0 to ubound(ArrValue_After,1)
				    For n=0 to ubound(ArrValue_After,2)
				        strCellValue_After= ArrValue_After(m,n)
				        strCellValues_After=strCellValues_After&";"&strCellValue_After
				    Next
				Next
				
				wait 2
				reVal=IMSSCA.Validations.FilterVerification(strCellValues_Before,strCellValues_After)
		    	If reVal = 0 Then
			      	Call ReportStep (StatusTypes.Pass, "Completed Validation of Group Table after performing 'Back to top level' Functionality. Group Table Contents are same with respect to first table", strPageName)
			      	Call ReportStep (StatusTypes.Pass, "Successfully Validated Group Table after performing 'Back to top level' Functionality. After expanding values within the table and clicking on 'Back to Top Level', Group tabe contents are same with respect to first table", strPageName)
			      	GroupTableDesignSteps = 1
				Else
					Call ReportStep (StatusTypes.Fail, "Completed Validation of Group Table after performing 'Back to top level' Functionality. Group Table Contents are different", strPageName)
			      	Call ReportStep (StatusTypes.Fail, "'Back to top level' Functionality of Group Table Design Tab not working correctly", strPageName)
			      	GroupTableDesignSteps = 0
				End If  
				
				wait 1
				Browser("Analyzer").Page("ReportCreation").Sync
				Browser("Analyzer").Page("ReportCreation").RefreshObject			
		
		Case "Step List"
				'Shweta 29/12/2016 - Start
'				'Verify the steps performed on Group Table while creating groups
'				'Identify weather performed steps are there on Step List 
'				'Click on all steps staring from first to check the fuctionality
'				
'				'Click on Step List From Design tab of Group Table Drop Down menu
'				'Call IMSSCA.General.DownNormalMenu(objData.item("strSelValue"),objData.item("MenuOption1"), "")
'				Call IMSSCA.General.DownNormalMenu("imgDesign", "welStepList", "")

				tabAtrributeName = "tabGrpTableColAttributeData"
				
				Call IMSSCA.General.ComponentDownNormalMenu(tabAtrributeName, "imgDesign", "welStepList", "")
				Browser("Analyzer").Page("ReportCreation").Sync
				Browser("Analyzer").Page("ReportCreation").RefreshObject
				'Shweta 29/12/2016 - End
				
				StepListOper = Cint(objData.item("NoOfStepListOperations"))
				For i = 1 To 20 Step 1
					If Browser("Analyzer").Frame("FrameStepList").WebTable("tabBack to First Step").Exist(1) Then
						'Descriptive program for Step List operations
						Set oDescStepList  = Description.Create()
						oDescStepList("micclass").value = "WebElement"
						oDescStepList("class").regularexpression=True
						oDescStepList("class").value="UltraWebListbar.*"
						oDescStepList("html tag").value="SPAN"
						'oDescOlap("innertext").value=stepListOperation
						Set objDescStepList =Browser("Analyzer").Frame("FrameStepList").WebTable("tabBack to First Step").ChildObjects(oDescStepList)
						'msgbox objDescStepList.count
						
						If objDescStepList.count = CINT(StepListOper) Then
							Call ReportStep (StatusTypes.Pass, "Step List dialogue box is displayed on the right side of the screen. Listing all the operations user performed on the group table", strPagename)
						Else 
							Call ReportStep (StatusTypes.Fail,  "Step List dialogue box is displayed on the right side of the screen. But could not find all list of the operations user performed on the group table", strPagename)
						End If
						
						Exit For
					End If
					
					If i = 20 Then
						Call ReportStep (StatusTypes.Fail,  "Step List dialogue box is not displayed on the right side of the screen. Could not find list of the operations user performed on the group table", strPagename)
						Exit Function
					End If
				Next
				
				For j = 0 To objDescStepList.count - 1 Step 1
					objDescStepList(j).Click
					StepListOper = objDescStepList(j).GetRoProperty("innertext")
					wait 2
					Browser("Analyzer").Page("ReportCreation").Sync
					Browser("Analyzer").Page("ReportCreation").RefreshObject
					
					'EditedGrp = 0
					
					Select Case TRim((StepListOper))
						Case "Edit Group Items"
							For l = 1 To 20 Step 1
								wait 3
								Browser("Analyzer").Page("ReportCreation").Sync
								Browser("Analyzer").Page("ReportCreation").RefreshObject
								
								If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDesignStepGrpTable").Exist(1) Then
									grpTabVal = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDesignStepGrpTable").GetROProperty("innertext")
									If InStr(1, Trim(Ucase(grpTabVal)), Trim(Ucase("Drop a Column Dimension Here")))> 0 and InStr(1, Trim(Ucase(grpTabVal)), Trim(Ucase("Drop a Measure Here")))> 0 Then
											Call ReportStep (StatusTypes.Pass, "Successfully validated Edit Group Items Table for First ----> Edited Group Items step. Edited Group Items successfully from Step List", strPagename)
											Exit For
									ElseIf InStr(1, Trim(Ucase(grpTabVal)), Trim(Ucase("Drop a Measure Here")))> 0 Then
											'TODO: Table validation wrt "First ----> Expand Group step"
											Call ReportStep (StatusTypes.Pass, "Successfully validated Edit Group Items Table for Second ----> Edited Group Items step. Edited Group Items successfully from Step List", strPagename)
											Exit For 
									End If
								End If	
								
								If l = 20 Then
									Call ReportStep (StatusTypes.Fail, "Edit Group Items Table not found. Could not test Edit Group Items Step Successfully", strPagename)
								End If
							Next
				
					
						Case "Expand Group"
							For m = 1 To 20 Step 1
								wait 3
								Browser("Analyzer").Page("ReportCreation").Sync
								Browser("Analyzer").Page("ReportCreation").RefreshObject
								
								If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDesignStepGrpTable").Exist(1) Then
									grpTabVal = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDesignStepGrpTable").GetROProperty("innertext")
									If InStr(1, Trim(Ucase(grpTabVal)), Trim(Ucase("Drop a Column Dimension Here")))> 0 and InStr(1, Trim(Ucase(grpTabVal)), Trim(Ucase("Drop a Measure Here")))> 0 Then
											'TODO: validate for table contents wrt First ----> Edited Group Items Group step
											Call ReportStep (StatusTypes.Pass, "Successfully validated Expand Group Table for First ----> Expand Group step. Expanded Group successfully from Step List", strPagename)
											Exit For
									ElseIf InStr(1, Trim(Ucase(grpTabVal)), Trim(Ucase("Drop a Measure Here")))> 0 Then
											'TODO: validate for table contents wrt Second ----> Edited Group Items step
											Call ReportStep (StatusTypes.Pass, "Successfully validated Edit Group Items Table for Second ----> Edited Groups step. Edited Groups successfully from Step List", strPagename)
											Exit For 
									End If
								End If	
								
								If m = 20 Then
									Call ReportStep (StatusTypes.Fail, "Expand Group Table not found. Could not test Expand Group Step Successfully", strPagename)
								End If
							Next
						
						Case "Insert Measure"
							For n = 1 To 20 Step 1
								wait 3			
								Browser("Analyzer").Page("ReportCreation").Sync
								Browser("Analyzer").Page("ReportCreation").RefreshObject
								
								If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDesignStepGrpTable").Exist(1) Then
									'TODO table Validation
										Set objNewwebtable=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDesignStepGrpTable")
										StepListArrValue_After=SCA.Webtable(objNewwebtable,"Insert Measure table",TRIM("Retriving_DataTableValue"),"Report Creation","","","","")
										For p=0 to ubound(StepListArrValue_After,1)
										    For q=0 to ubound(StepListArrValue_After,2)
										        strStepListCellValue_After= StepListArrValue_After(p,q)
										        strStepListCellValues_After=strStepListCellValues_After&";"&strStepListCellValue_After
										    Next
										Next
										
										wait 2
										reVal=IMSSCA.Validations.FilterVerification(strCellValues_Before,strStepListCellValues_After)
								    	If reVal = 0 Then
									      	Call ReportStep (StatusTypes.Pass, "Successfully Validated Group Table after performing Step List ----------> Insert Measure Step. Same Original Created Group Table Contents are displayed", strPagename)
									      	Call ReportStep (StatusTypes.Pass, "Validated Step List Functionality of Group Table Design Tab. Step List Working successfully", strPagename)
									      	GroupTableDesignSteps = 1
									      	Exit For
										Else		
											Call ReportStep (StatusTypes.Fail, "Successfully Validated Group Table after performing Step List ----------> Insert Measure Step. Group Table Contents are different", strPagename)
											Call ReportStep (StatusTypes.Fail, "Validated Step List Functionality of Group Table Design Tab. Step List Not Working Correctly", strPagename)
											GroupTableDesignSteps = 0
											Exit For
										End If
								End If	
								
								If n = 20 Then
									Call ReportStep (StatusTypes.Fail, "Insert Measure Table not found. Could not test Expand Group Step Successfully", strPagename)
								End If
							Next
							
							
					End Select
					
				Next
				'End of case "Step List"
			
		End Select
		
	End Function

	
	'<Composed By : Shweta B Nagaral>
	'<It performs creation of GroupTable by adding row attribute, column attribute and measure attribute. 
	'It Clicks on AddAllMembers option to creates groups for various row and column attributes>
	Public Sub TestAddAllMembersGroupTable() 
		Dim k, l
		
		'Call Grouptable creation method to create grp table with 2 dimension (Row and Column Attribute) and Measure - Start
				'1) Add Row attribute
				Browser("Analyzer").Page("ReportCreation").Sync
				Browser("Analyzer").Page("ReportCreation").RefreshObject
				Call IMSSCA.General.GroupTableCreationInSCA(0,"SCAData",1,2,"","" )
				
				'Call AddAllMembers
				Call IMSSCA.General.AddAllMembers("", "", "ON", "rdoGroup")
				
				'Expand the grp created
				For k = 1 To 10 Step 1
					If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgGexpand").Exist(1) Then
						wait 1
						Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgGexpand").Click
						wait 2
						Browser("Analyzer").Page("ReportCreation").Sync
						Browser("Analyzer").Page("ReportCreation").RefreshObject
						Exit for
					End If
					If  k = 10 Then
						Call ReportStep (StatusTypes.Fail,"Could not expand Row attribute group created" ,"Report Creation Page")	
					End If
				Next
				
				Browser("Analyzer").Page("ReportCreation").Sync
				Browser("Analyzer").Page("ReportCreation").RefreshObject
				Call IMSSCA.General.GroupTableCreationInSCA(1,"SCAData",1,2,"","" )
			  
				'2) Add Col attribute
				Browser("Analyzer").Page("ReportCreation").Sync
				Browser("Analyzer").Page("ReportCreation").RefreshObject
				Call IMSSCA.General.GroupTableCreationInSCA(0,"SCAData",1,4,"","")
				'Call AddAllMembers
				Call IMSSCA.General.AddAllMembers("", "", "ON", "rdoGroup")
				
				'Expand the grp created
				For l = 1 To 10 Step 1
					If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgGexpand").Exist(1) Then
						wait 1
						Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgGexpand").Click
						wait 2
						Browser("Analyzer").Page("ReportCreation").Sync
						Browser("Analyzer").Page("ReportCreation").RefreshObject
						Exit for
					End If
					If  l = 10 Then
						Call ReportStep (StatusTypes.Fail,"Could not expand Row attribute group created", "Report Creation Page")	
					End If
				Next
				
				Browser("Analyzer").Page("ReportCreation").Sync
				Browser("Analyzer").Page("ReportCreation").RefreshObject
				Call IMSSCA.General.GroupTableCreationInSCA(1,"SCAData",1,4,"","" )
				 
				'3) Add Measure attribute
				Browser("Analyzer").Page("ReportCreation").Sync
				Browser("Analyzer").Page("ReportCreation").RefreshObject
				Call IMSSCA.General.GroupTableCreationInSCA(0,"SCAData",1,5,"","")
		
				Call IMSSCA.General.GroupTableCreationInSCA(1,"SCAData",1,5,"","" )
				wait 1
				Browser("Analyzer").Page("ReportCreation").Sync
				Browser("Analyzer").Page("ReportCreation").RefreshObject
		'Call Grouptable creation method to create grp table with 2 dimension (Row and Column Attribute) and Measure - End
			
	End Sub
	
	
	'<Composed By : Shweta B Nagaral>
	'Saves Analytical Pattern by clicking on Down Normal image of row attribute. Opens 'SCAReportSheet.xls' Excel and writes attribute value in "AnalyticalPattern" sheet
	'<strPatName: name of Analytical Pattern to be saved: String>
	'<strDispFolder: Display Folder selection: String>
	'<PatternSave: Shared/Personal Pattern folder option: String>
	'<StrPage: Page name: String>
	Public Function SaveAnalyticalPatterns(ByVal strPatName, ByVal strDispFolder, ByVal PatternSave, ByVal StrPage)
		Dim o,i, SaveAnalyticalPatternAttributeName, expandDisabled, oGrpTabDownNormal, oGrpTabDownNormalObj
		Dim objExcel, objWorkbook
		expandDisabled = 0
				
		'Descriptive programe to click on Down Normal image of row attribute
		Set oGrpTabDownNormal  = Description.Create()
		oGrpTabDownNormal("micclass").value = "Image"
		oGrpTabDownNormal("file name").regularexpression=True
		oGrpTabDownNormal("file name").value = "down_normal.gif"
		oGrpTabDownNormal("html tag").value = "IMG"
		oGrpTabDownNormal("image type").value = "Plain Image"
		oGrpTabDownNormal("name").value = "Image"
		oGrpTabDownNormal("outerhtml").regularexpression=True
		oGrpTabDownNormal("outerhtml").value = ".*CURSOR: hand.*"
						
		Set oGrpTabDownNormalObj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(oGrpTabDownNormal)
		'msgbox oGrpTabDownNormalObj.count
		wait 2
		oGrpTabDownNormalObj(0).FireEvent "onmouseover"
		wait 1
		oGrpTabDownNormalObj(0).Click
		wait 2
		Browser("Analyzer").Page("ReportCreation").Sync
		Browser("Analyzer").Page("ReportCreation").RefreshObject
		
		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welSaveAnalyticalPattern").Click
					
		For i = 1 To 10 Step 1
			If Browser("Analyzer").Page("ReportCreation").Frame("__dialog").Exist(1) Then
				expandDisabled = 1
				
				Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebEdit("txtName").Set strPatName
				SaveAnalyticalPatterns = strPatName
				Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebList("ddlDisplayFolder").Select strDispFolder
				Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebRadioGroup("rdoFolderPattern").Select PatternSave
				SaveAnalyticalPatternAttributeName = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebElement("welSaveAnalyticalPatternAttributeName").GetROProperty("InnerText")
				
				'Open 'SCAReportSheet.xls' Excel and write attribute value in 3rd sheet "AnalyticalPattern" - Start
				Set objExcel = CreateObject("Excel.Application")
				Set objWorkbook = objExcel.Workbooks.Open(Environment.Value("CurrDir")&Environment.Value("ReportCreationFile"))
				
				objExcel.Application.Visible = False
				objWorkbook.sheets("AnalyticalPattern").cells(3,7)=strPatName
				
				'Saving Analytical Pattern name at 5th row in Excel to validate 'Adding Analytical Pattern From Personal to Favorites'
				objWorkbook.sheets("AnalyticalPattern").cells(5,7)=strPatName
				
				'Saving Analytical Pattern name at 6th row in Excel to validate 'Adding Analytical Pattern From Personal to Favorites'
				objWorkbook.sheets("AnalyticalPattern").cells(6,6)=strPatName
				
				'Saving Analytical Pattern name at 7th row in Excel to validate 'Deleting Analytical Pattern'
				objWorkbook.sheets("AnalyticalPattern").cells(7,7)=strPatName
				
'				'Saving Analytical Pattern Attribute name at 3rd and 6th row in Excel to validate 'Adding Analytical Pattern From Personal to Favorites'
'				objWorkbook.sheets("AnalyticalPattern").cells(3,10)=SaveAnalyticalPatternAttributeName
'				objWorkbook.sheets("AnalyticalPattern").cells(6,7)=SaveAnalyticalPatternAttributeName
				
				objWorkbook.Save
				objExcel.ActiveWorkbook.Close
				objExcel.Application.Quit
				Set objExcel=Nothing
				'Open 'SCAReportSheet.xls' Excel and write attribute value in 3rd sheet "AnalyticalPattern" - End
				
				Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebButton("btnOK").Click
				Exit For
			End If
		Next
		
		If expandDisabled = 0 and o = 10 Then
			Call ReportStep (StatusTypes.Fail, "Could not set valus while saving analtical patterns", "StrPage")
		End If
		
		wait 2
		Browser("Analyzer").Page("ReportCreation").Sync
		Browser("Analyzer").Page("ReportCreation").RefreshObject
		
	End Function


	'<Composed By : Shweta B Nagaral>	
	'<Performs Following validations:
	'	1) Saves of HTMl Export to local machine
	'	2) Validaes HTML exported FileName, Sheet if any and HTML exported Extension if any
	'	3) Compares Report Creation Page Group Table component wrt locally Saved HTML report >
	'<strReportSavedName: Group Table Name of the Report created in SCA application>
	'<objData: Reference to objData>
	Public Sub SaveHTMLOption(ByVal strReportSavedName, ByVal objData)
	
		Dim a, b, i, j, p, q, x, y, m, n, s, t, k, r
		Dim path, toolBarNameFormat, formatPos, myVar, strReport, strReportName, strReportFormat, intRow_Count, intColumn_Count, rcAttribute, ccAttribute, strcellvalue, ccMeasure, rcMeasure, strRowGrandTotalValue, strColGrandTotalValue
		Dim objbtnExport, objbtnCloseExport, fso, objGrpTabColData, objGrpTabMeasure, objLstWinButton, objbtnSave, objYes, objClose
		Dim val1, val2, rc, cc, trimCompareVal, TrimVal, compareVal, strHTMLReportSaved, val2Len, compareValLen
			
		'HTML Report saving
		For i = 1 To 50 Step 1
			If Browser("Hosting Templates - Ops").WinObject("Notification").WinButton("SaveAsDropDownWinButton").Exist(2) then
			 Exit For
			End If
		Next
		wait 2
		
		toolBarNameFormat = Browser("Analyzer").WinObject("Notification bar").GetROProperty("text")
		
		'Retreiving name and format of HTML report in Notification tool bar
		'strReport Splitting
		strReport = mid(toolBarNameFormat, InStr(1, Trim(Ucase(toolBarNameFormat)), Trim(Ucase("GroupTable"))), 15)
			a=Split(strReport, ".")
			b= UBound(a)
			For j = 0 To b Step 2
				strReportName = a(j)
				strReportFormat = "."&a(j+1)
			Next
		
		path = Environment.Value("CurrDir")&"Exportresults\"&TRim(strReport)
		
		'Validate Report extension, name of the report. HTML report name should be same as group table name in SCA.
		Call ReportStep (StatusTypes.Pass, "##########################Validating HTML Report Extension. Report extension should be '.htm'", "HTML Export Options Page")
		If UCASE(strReportFormat) = UCASE(".htm") Then
			Call ReportStep (StatusTypes.Pass, "Found '.htm' as file extension for HTML Export saving Report. Saving Report with '.htm' extension", "HTML Export Options Page")
		Else	
			Call ReportStep (StatusTypes.Fail, "Could not find '.htm' as file texnstion for HTML Export saving Report. Saving Report with different extension", "HTML Export Options Page")
		End If
		
		'HTML report name should be same as group table name in SCA.
		Call ReportStep (StatusTypes.Pass, "##########################Validation of HTML Report Name. HTML report name should be same as Group Table name in SCA", "HTML Export Options Page")
		If strReportName = strReportSavedName Then
			Call ReportStep (StatusTypes.Pass, "HTML Report Name is as Group Table name in SCA . Saving "&strReportName&" Report with '.htm' extension", "HTML Export Options Page")
		Else
			Call ReportStep (StatusTypes.Fail, "HTML Report Name is not same as Group Table name in SCA . Saving "&strReportName&" Report with '.htm' extension", "HTML Export Options Page")
		End If
		wait 4
		'changes by srinivas
		'Set objLstWinButton=Browser("Analyzer").WinObject("Notification bar").WinButton("lstWinButton")
		Set objLstWinButton=Browser("Hosting Templates - Ops").WinObject("Notification").WinButton("SaveAsDropDownWinButton")
		Call SCA.ClickOn(objLstWinButton,"lstWinButton", "HTML Report Export Dialouge")
		wait 2
		Browser("Analyzer").WinMenu("SaveAsContextMenu").Select "Save as"
		wait 2
		Browser("Analyzer").Dialog("Save As").WinEdit("lnkFile name:").Set path
		wait 2
		Set objbtnSave=Browser("Analyzer").Dialog("Save As").WinButton("btnSave")
		Call SCA.ClickOn(objbtnSave,"btnSave", "HTML Report Export Dialouge")
		wait 2
		
		IF Browser("Analyzer").Dialog("Save As").Dialog("Confirm Save As").WinButton("Yes").Exist(5) Then
			Set objYes=Browser("Analyzer").Dialog("Save As").Dialog("Confirm Save As").WinButton("Yes")
			Call SCA.ClickOn(objYes,"Yes", "HTML Report Export Dialouge")
			wait 2
		End If
		
		For k = 1 To 3 Step 1
			If Browser("Analyzer").WinObject("Notification bar").WinButton("Close").Exist(1) Then
				Set objClose=Browser("Analyzer").WinObject("Notification bar").WinButton("Close")
				Call SCA.ClickOn(objClose,"Close", "HTML Report Export Dialouge, Notification Bar")
				Exit For
			End If
		Next
		
		wait 2
		Set objbtnCloseExport=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebButton("btnClose")
		Call SCA.ClickOn(objbtnCloseExport,"CloseExportbtn", "HTML Report Export Dialouge")
		
		'Vaildating name and format of report in Local Machine Path
		Set fso = CreateObject("Scripting.FileSystemObject")
		If (fso.FileExists(path)) Then
		   Call ReportStep (StatusTypes.Pass, path & " exists.", "Local Machine Path")
		   Call ReportStep (StatusTypes.Pass, "Successfully found locally saved SCA application HTML Report " &strReportName, "Local Machine Path")
		   'Code to Open Saved Html file in explorer - Start
		   	Dim objIE
			'Create an IE object
			Set objIE = CreateObject("InternetExplorer.Application")
			'Open file
			objIE.Navigate path
		
		Else
		   Call ReportStep (StatusTypes.Fail, path & " doesn't exist.", "Local Machine Path")
		   Call ReportStep (StatusTypes.Fail, "Could not find locally saved SCA application HTML Report " &strReportName& " successfully", "Local Machine Path")
		End If
		
'		'Code to Open Saved Html file in explorer - Start
'		Set fso = CreateObject("Scripting.FileSystemObject")
'		If (fso.FileExists(path)) Then
'		  	'msgbox "Pass"
'		  	Dim objIE
'			'Create an IE object
'			Set objIE = CreateObject("InternetExplorer.Application")
'			'Open file
'			objIE.Navigate path
'		Else
'		 	'msgbox "Fail"
'		End If
'		'Code to Open Saved Html file in explorer - End
	
		'Capture Report Name from Open HTML File
		strHTMLReportSaved = Browser("HTMLreport").Page("Page").WebElement("welGroupTable").GetROProperty("innertext")

		'HTML report name should be same as group table name in SCA.
		Call ReportStep (StatusTypes.Pass, "##########################Validation of Saved HTML Report Name. HTML report name should be same as Group Table name in SCA", "HTML Export Options Page")
		If strHTMLReportSaved = strReportSavedName Then
			Call ReportStep (StatusTypes.Pass, "HTML Report Name is as same as Group Table name in SCA . Saved "&strReportName&" Report with '.htm' extension", "HTML Export Options Page")
		Else
			Call ReportStep (StatusTypes.Fail, "HTML Report Name is not same as Group Table name in SCA . Saved "&strReportName&" Report with '.htm' extension", "HTML Export Options Page")
		End If
		
		'Comparision of SAC application table and HTML File saved - Start
		Call ReportStep (StatusTypes.Pass, "###################Comparision of SCA application table and HTML File saved", "")
		val2 = ""
		rc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDesignStepGrpTable").GetRoProperty("rows")
		cc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDesignStepGrpTable").GetRoProperty("cols")
		For m = 1 To rc Step 1
			For n = 1 To cc Step 1
				TrimVal = Trim(Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDesignStepGrpTable").GetCellData(m, n))
				val1 = Replace(TrimVal, " ", "")
				val2 = val2&val1
				val2Len = len(val2)
			Next
		Next
		
		rc = Browser("HTMLreport").Page("Page").WebTable("tabHTMLData").GetRoProperty("rows")
		cc = Browser("HTMLreport").Page("Page").WebTable("tabHTMLData").GetRoProperty("cols")
		For p = 1 To rc Step 1
			For q = 1 To cc Step 1
				trimCompareVal = Trim(Browser("HTMLreport").Page("Page").WebTable("tabHTMLData").GetCellData(p, q))
				compareVal= Replace(trimCompareVal, " ", "")
				compareValLen = Len(compareVal)		
			Next
		Next
		
		'Commented For alternate validation wrt len(val2) and Len(compareVal) - Start		
		'		If StrComp(val2, compareVal) = 0 Then
		'			msgbox "Same"
		'			Call ReportStep (StatusTypes.Pass, "########################Report Creation Page Group Table values are same as HTML File values", "ReportCreation Page")
		'		ElseIf StrComp(val2, compareVal) = -1 Then
		'			msgbox "string1 < string2"
		'			Call ReportStep (StatusTypes.Fail, "########################Report Creation Page Group Table values are not same as HTML File values", "ReportCreation Page")
		'			Call ReportStep (StatusTypes.Fail, "########################Report Creation Page Group Table values are lesser than as HTML File values", "ReportCreation Page")
		'		ElseIf StrComp(val2, compareVal) = 1 AND len(val2) = len(compareVal) Then
		'			msgbox "string1 > string2"
		'			msgbox "same contents"
		'		End If
		'Commented For alternate validation wrt len(val2) and Len(compareVal) - End

		If val2Len > compareValLen Then
			Call ReportStep (StatusTypes.Fail, "########################Report Creation Page Group Table values are not same as HTML File values", "ReportCreation Page")
			Call ReportStep (StatusTypes.Fail, "########################Report Creation Page Group Table values are lesser than as HTML File values", "ReportCreation Page")
		ElseIf val2Len = compareValLen Then
			Call ReportStep (StatusTypes.Pass, "########################Report Creation Page Group Table values are same as HTML File values", "ReportCreation Page")
		ElseIf val2Len < compareValLen  Then
			Call ReportStep (StatusTypes.Fail, "########################Report Creation Page Group Table values are not same as HTML File values", "ReportCreation Page")
			Call ReportStep (StatusTypes.Fail, "########################Report Creation Page Group Table values are greater than as HTML File values", "ReportCreation Page")
		End If
		
		'Comparision of SAC application table and HTML File saved - End
		
		'Close HTML opened file
		For r = 1 To 10 Step 1
			If Browser("CreationTime:=1").Exist(2) Then
				 Browser("CreationTime:=1").Close
				 Call ReportStep (StatusTypes.Pass, "Successfully closed HTML opened in IE browser", "HTML Saved Page")
				 Exit For
			End If 
			
			If r = 10 Then
				Call ReportStep (StatusTypes.Fail, "Could not close HTML opened in IE browser successfully", "HTML Saved Page")
			End If
		Next
		
	End Sub
	
	
	'<Creates Chart in SCA with respect to Pivot table>
    '<strChartPos: Position where chart has to be created, It could be Top, Bottom, Left, Right :  String>
    '<strChartType: Chart Type to be selected. It could be Basic Charts, Advanced Charts :  String>
    '<strChartSheet: Sheet where chart has to be created. It could be New Sheet, Sheet1 :  String>
	Public Sub ChartCreationInSCA(ByVal strChartPos, ByVal strChartType, ByVal strChartSheet)
	
		Dim objPos, objType, objSheet, objAdd
		Set objPos=Browser("Analyzer").Page("ReportCreation").Frame("CreateChart").WebRadioGroup("rdoPosition")
   		Call SCA.ClickOnRadio(objPos,strChartPos, "ReportCreation")
   		Set objType=Browser("Analyzer").Page("ReportCreation").Frame("CreateChart").WebRadioGroup("rdoChartType")
   		Call SCA.ClickOnRadio(objType,strChartType, "ReportCreation")
   		Set objSheet=Browser("Analyzer").Page("ReportCreation").Frame("CreateChart").WebList("lstddlSheets")
   		Call SCA.SelectFromDropdown(objSheet,strChartSheet)	
   	
   		Set objAdd=Browser("Analyzer").Page("ReportCreation").Frame("CreateChart").WebButton("btnOk")
   		Call SCA.ClickOn(objAdd,"OkButton" , "ReportCreation")
   		

	End Sub

	
	'<Settings on Sheet created in Report Creation Page>
	'<SheetName: Name of sheet to be renamed :  String>
    '<strSheetOperation: Operation on Sheet to be performed :  String>
    '<strOperVal: For Sheet rename operation, rename value should be included :  String>
	Public Sub SheetSettings(BYVal SheetName, ByVal strSheetOperation, ByVal strOperVal )
	
		Dim webElem, webElemInnertText, oDescEditTxt, oDescEditTxtObj, WshShell, objSheetOper, objDialogOk
		
		Browser("Analyzer").Page("ReportCreation").Sync
		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").RefreshObject
		
		If strSheetOperation <> "ClikOperation" Then
		    wait 5
			set webElem = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(SheetName)
			webElemInnertText = webElem.GetROProperty("innertext")
			Setting.WebPackage("ReplayType") = 2
			webElem.RightClick
			Setting.WebPackage("ReplayType") = 1
			wait 2
			
			Set objSheetOper = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strSheetOperation)
			wait 2
		 	Call SCA.ClickOn(objSheetOper,strSheetOperation, "ReportCreation Page")
	   		wait 2
		End If
		
		Select Case strSheetOperation
		
		'<shweta - 28/8/2015> - Start
		Case "ClikOperation"
				'Perform click on different sheets in SCA Report Creation Page
						Set oSheet  = Description.Create()
						oSheet("micclass").value = "WebElement"
						oSheet("html tag").value = "NOBR"
						oSheet("innerhtml").regularexpression=True
						oSheet("innerhtml").value = "&nbsp.*"&strOperVal&"&nbsp.*"
						oSheet("innertext").value = strOperVal
										
						Set oSheetObj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(oSheet)
						'msgbox oSheetObj.count
						wait 1
						If oSheetObj.count = 1 Then
							oSheetObj(0).click
						Else
							Call ReportStep (StatusTypes.Fail, "Cannot perform Click operation on Sheet. More than one sheet objects", "ReportCreation Page")
						End If
						
						wait 2
						Browser("Analyzer").Page("ReportCreation").Sync
						Browser("Analyzer").Page("ReportCreation").RefreshObject
				
		'<shweta - 28/8/2015> - End	
		
		Case "welDeleteSheet"
						If Browser("Analyzer").Dialog("Message from webpage").Exist(30) = True Then
								
								'Click on Ok
								Set objDialogOk = Browser("Analyzer").Dialog("Message from webpage").WinButton("btnOK")
								Call SCA.ClickOn(objDialogOk,"Dialog Ok Button", "ReportCreation Page")
						   		wait 6
						   		
		    					If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welSheetForOperation").Exist(0) Then
					   				Call ReportStep (StatusTypes.Fail, "Could not delete '"&webElemInnertText&"' Sheet successfully" , "ReportCreation Page")
								Else 
									Call ReportStep (StatusTypes.Pass, "Successfully Deleted '"&webElemInnertText&"' Sheet" , "ReportCreation Page")
								End If
								
						else
		
								Call ReportStep (StatusTypes.Fail, "Delete Dialog box doesnt exist", "ReportCreation Page")
						
						End If
			
		Case "welRenameSheet"
			
						'Descriptive programe to edit the text for renaming
						Set oDescEditTxt  = Description.Create()
						oDescEditTxt("micclass").value = "WebEdit"
					    oDescEditTxt("html tag").value = "INPUT"
					    oDescEditTxt("type").value = "text"
					    oDescEditTxt("kind").value = "singleline"
					    oDescEditTxt("value").regularexpression=True
					    oDescEditTxt("value").value = ".*Sheet*."
					    
					    'oDescEditTxt("xpath").value="//TD[@id='SheetActivated']/NOBR[1]/INPUT[1]"				
						wait 2
						Set oDescEditTxtObj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(oDescEditTxt)
						'oDescEditTxtObj(i).Set "Shweta"
						Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebEdit(oDescEditTxt).Set strOperVal
						
						'Using Shell object to enter
						Set WshShell = CreateObject("WScript.Shell")
						WshShell.SendKeys "{ENTER}"
						Set WshShell = Nothing
						wait 8
						Call IMSSCA.Validations.ValidateReplaceField(webElem, "Rename", webElemInnertText, strOperVal, "ReportCreation Page")
						
		End Select

		Browser("Analyzer").Page("ReportCreation").Sync
		Browser("Analyzer").Page("ReportCreation").RefreshObject
	End Sub
	
	
'	'<Settings on Sheet created in Report Creation Page>
'	'<SheetName: Name of sheet to be renamed :  String>
'    '<strSheetOperation: Operation on Sheet to be performed :  String>
'    '<strOperVal: For Sheet rename operation, rename value should be included :  String>
'	Public Sub SheetSettings(ByVal strSheetOperation, BYVal SheetName, ByVal strOperVal)
'	
'		Dim webElem, webElemInnertText, oDescEditTxt, oDescEditTxtObj, WshShell, objSheetOper, objDialogOk, oSheet, oSheetObj
'		
'		set webElem = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(SheetName)
'		webElemInnertText = webElem.GetROProperty("innertext")
'		Setting.WebPackage("ReplayType") = 2
'		webElem.RightClick
'		wait 2
'		
'		Set objSheetOper = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strSheetOperation)
'		wait 2
'	 	Call SCA.ClickOn(objSheetOper,strSheetOperation, "ReportCreation Page")
'   		wait 2
'		
'		Select Case strSheetOperation
'		'<shweta - 28/8/2015> - Start
'		Case "ClikOperation"
'				'Perform click on different sheets in SCA Report Creation Page
'						Set oSheet  = Description.Create()
'						oSheet("micclass").value = "WebElement"
'						oSheet("html tag").value = "NOBR"
'						oSheet("innerhtml").regularexpression=True
'						oSheet("innerhtml").value = "&nbsp.*"&strOperVal&"&nbsp.*"
'						oSheet("innertext").value = strOperVal
'										
'						Set oSheetObj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(oSheet)
'						msgbox oSheetObj.count
'						wait 1
'						If oSheetObj.count = 1 Then
'							oSheetObj(0).click
'						Else
'							Mgsbox "More than one sheet objects"
'						End If
'						
'						wait 2
'						Browser("Analyzer").Page("ReportCreation").Sync
'						Browser("Analyzer").Page("ReportCreation").RefreshObject
'				
'		'<shweta - 28/8/2015> - End				
'		Case "welDeleteSheet"
'						If Browser("Analyzer").Dialog("Message from webpage").Exist = True Then
'								
'								'Click on Ok
'								Set objDialogOk = Browser("Analyzer").Dialog("Message from webpage").WinButton("btnOK")
'								Call SCA.ClickOn(objDialogOk,"Dialog Ok Button", "ReportCreation Page")
'						   		wait 2
'						   		
'		    					If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welSheetForOperation").Exist(0) Then
'					   				Call ReportStep (StatusTypes.Fail, "Could not delete '"&webElemInnertText&"' Sheet successfully" , "ReportCreation Page")
'								Else 
'									Call ReportStep (StatusTypes.Pass, "Successfully Deleted '"&webElemInnertText&"' Sheet" , "ReportCreation Page")
'								End If
'								
'						else
'		
'								Call ReportStep (StatusTypes.Fail, "Delete Dialog box doesnt exist", "ReportCreation Page")
'						
'						End If
'			
'		Case "welRenameSheet"
'			
'						'Descriptive programe to edit the text for renaming
'						Set oDescEditTxt  = Description.Create()
'						oDescEditTxt("micclass").value = "WebEdit"
'					    oDescEditTxt("html tag").value = "INPUT"
'					    oDescEditTxt("type").value = "text"
'					    oDescEditTxt("kind").value = "singleline"
'					    oDescEditTxt("value").regularexpression=True
'					    oDescEditTxt("value").value = ".*Sheet*."
'					    				
'						Set oDescEditTxtObj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(oDescEditTxt)
'						'oDescEditTxtObj(i).Set "Shweta"
'						Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebEdit(oDescEditTxt).Set strOperVal
'						
'						'Using Shell object to enter
'						Set WshShell = CreateObject("WScript.Shell")
'						WshShell.SendKeys "{ENTER}"
'						Set WshShell = Nothing
'						
'						Call IMSSCA.Validations.ValidateReplaceField(webElem, "Rename", webElemInnertText, strOperVal, "ReportCreation Page")
'						
'		End Select
'   		   		
'	End Sub

	
	'<Composed By : Shweta B Nagaral>	
	'<Performs Following validations:
	'	1) Saves of CSV Export to local machine
	'	2) Validaes CSV exported FileName, SheetName and CSV exported Extension
	'	3) Compares Report Creation Page Group Table component wrt locally Saved CSV report >
	'<strReportSavedName: Saved Report Name:  String>
	'<sheetVal: Excel sheet value:  Number>
	'<sheetName: Excel sheet Name:  Number>
	'<objData: Reference to objData>
	Public Sub SaveCSVOption(ByVal strReportSavedName, ByVal sheetVal, ByVal sheetName)
		
		Dim a, b, i, j, p, q, x, y, m, n, s, t, k
		Dim path, toolBarNameFormat, formatPos, myVar, strReportName, strReportFormat, ExcelSheetName, intRow_Count, intColumn_Count, rcAttribute, ccAttribute, strcellvalue, strexcelcellvalue, ccMeasure, rcMeasure, strRowGrandTotalValue, strColGrandTotalValue
		Dim objbtnExport, objbtnCloseExport, fso, objExcel, objGrpTabColData, objGrpTabMeasure, objLstWinButton, objbtnSave, objYes, objClose, ObjExcelFile, ObjExcelSheet, strReport
			
		'CSV Report saving
		For i = 1 To 50 Step 1
			If Browser("Analyzer").WinObject("Notification bar").WinButton("lstWinButton").Exist(0) then
			 Exit For
			End If
		Next
		wait 2
		
		toolBarNameFormat = Browser("Analyzer").WinObject("Notification bar").GetROProperty("text")
	
		'Retreiving name and format of CSV report in Notification tool bar
		'strReport Splitting
		strReport = mid(toolBarNameFormat, InStr(1, Trim(Ucase(toolBarNameFormat)), Trim(Ucase("GroupTable"))), 15)
				a=Split(strReport, ".")
				b= UBound(a)
				For j = 0 To b Step 2
					strReportName = a(j)
					strReportFormat = "."&a(j+1)
				Next
					
		path = Environment.Value("CurrDir")&"Exportresults\"&Trim(strReport)
		
		'Validate the extension of CSV file
		Call ReportStep (StatusTypes.Pass, "Validating extension of CSV file. Report extension should be '.CSV'", "CSV Export Options Page")
		If UCASE(strReportFormat) = UCASE(".csv") Then
			Call ReportStep (StatusTypes.Pass, "Found '.csv' as file extension for CSV Export saving Report. Saving Report with '.CSV' extension", "CSV Export Options Page")
		Else	
			Call ReportStep (StatusTypes.Fail, "Could not find '.csv' as file texnstion for CSV Export saving Report. Saving Report with different extension", "CSV Export Options Page")
		End If
		
		'TODO: CSV report name should be same as group table name in SCA and sheet name should be same as group table name in SCA.
		Call ReportStep (StatusTypes.Pass, "Validation of CSV Report Name. CSV report name should be same as Group Table name in SCA", "CSV Export Options Page")
		If strReportName = strReportSavedName Then
			Call ReportStep (StatusTypes.Pass, "CSV Report Name is as Group Table name in SCA . Saving "&strReportName&" Report with '.CSV' extension", "CSV Export Options Page")
		Else
			Call ReportStep (StatusTypes.Fail, "CSV Report Name is not same as Group Table name in SCA . Saving "&strReportName&" Report with '.CSV' extension", "CSV Export Options Page")
		End If
			
		'Shweta - <11/1/2016> - Click on SaveAsContextMenu Menu - Start 			
		'		Set objLstWinButton=Browser("Analyzer").WinObject("Notification bar").WinButton("lstWinButton")
		'		Call SCA.ClickOn(objLstWinButton,"lstWinButton", "Excel Report Export Dialouge")
		'		
		'		'Browser("Analyzer").WinMenu("ContextMenu").Select "Save as"
		'		Browser("Analyzer").WinMenu("SaveAsContextMenu").Select "Save as"
		Set objLstWinButton=Browser("Analyzer").WinObject("Notification bar").WinButton("lstWinButton")
        intLoopCounter = 0
        Do
            If intLoopCounter=15 Then
                Exit Do 
            End If
             intLoopCounter=intLoopCounter+1
            Call SCA.ClickOn(objLstWinButton,"lstWinButton", "Excel Report Export Dialouge")
        Loop Until Browser("Analyzer").WinMenu("SaveAsContextMenu").Exist(2) = true
        
        If intLoopCounter=15 Then
            Call ReportStep (StatusTypes.Fail, "Context menu of Windows Notification bar for CSV file downloading was not clicked", "Export Options Page")
        End If

		Browser("Analyzer").WinMenu("SaveAsContextMenu").Select "Save as"
		wait 2
		 'Shweta - <11/1/2016> - Click on SaveAsContextMenu Menu - End 
		 
        'Shweta - <11/1/2016> - Click on Click on Windows Save Location Weblist - Start
       'Browser("Analyzer").Dialog("Save As").WinEdit("lnkFile name:").WaitProperty "visible",True,10000
       if Browser("Analyzer").Dialog("Save As").WinEdit("lnkFile name:").Exist(120) = False Then
           Call ReportStep (StatusTypes.Fail, "Didnt enter "&path&" location to save the excel file", "Export Options Page")
       End if
       
       Browser("Analyzer").Dialog("Save As").WinObject("Items View").WinList("Items View").Highlight
       Browser("Analyzer").Dialog("Save As").WinObject("Items View").WinList("Items View").Click
       wait 2
       'Shweta - <11/1/2016> - Click on Click on Windows Save Location Weblist - End
		 
		Browser("Analyzer").Dialog("Save As").WinEdit("lnkFile name:").Set path
		wait 2
		Set objbtnSave=Browser("Analyzer").Dialog("Save As").WinButton("btnSave")
		Call SCA.ClickOn(objbtnSave,"btnSave", "CSV Report Export Dialouge")
		wait 2
		
		IF Browser("Analyzer").Dialog("Save As").Dialog("Confirm Save As").WinButton("Yes").Exist(5) Then
			Set objYes=Browser("Analyzer").Dialog("Save As").Dialog("Confirm Save As").WinButton("Yes")
			Call SCA.ClickOn(objYes,"Yes", "CSV Report Export Dialouge")
			wait 2
		End If
		
		For k = 1 To 3 Step 1
			If Browser("Analyzer").WinObject("Notification bar").WinButton("Close").Exist(1) Then
				Set objClose=Browser("Analyzer").WinObject("Notification bar").WinButton("Close")
				Call SCA.ClickOn(objClose,"Close", "CSV Report Export Dialouge, Notification Bar")
				Exit For
			End If
		Next
		
		Set objbtnCloseExport=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebButton("btnClose")
		Call SCA.ClickOn(objbtnCloseExport,"CloseExportbtn", "CSV Report Export Dialouge")
		
		'Vaildating name and format of report in Local Machine Path
		Set fso = CreateObject("Scripting.FileSystemObject")
		If (fso.FileExists(path)) Then
		   Call ReportStep (StatusTypes.Pass, path & " exists.", "Local Machine Path")
		   Call ReportStep (StatusTypes.Pass, "Successfully found locally saved SCA application CSV Report " &strReportName, "Local Machine Path")
		
		Else
		   Call ReportStep (StatusTypes.Fail, path & " doesn't exist.", "Local Machine Path")
		   Call ReportStep (StatusTypes.Fail, "Could not find locally saved SCA application CSV Report " &strReportName, "Local Machine Path")
		   Exit Sub
		End If
			
	
		'Excel content validation wrt application group table data - Start
		Set objExcel = CreateObject("Excel.Application") 
		objExcel.Visible =False  
		Set ObjExcelFile = objExcel.Workbooks.Open(path)
		Set ObjExcelSheet = ObjExcelFile.Sheets(sheetVal)
			
		'Vaildating Excel Sheet Name wrt Application Report Creatio Page Sheet Name. 
		'Get the Sheet Name of the Excel workbook
		Call ReportStep (StatusTypes.Pass, "Vaildating CSV Sheet Name wrt Application Report Creation Page Sheet Name", "")
		ExcelSheetName = ObjExcelSheet.Name
		If StrComp(sheetName,ExcelSheetName) = 0 Then	
			Call ReportStep (StatusTypes.Pass, "CSV Excel Sheet Name is same as Application Report Creation Page Sheet Name " &sheetName, "")
		Else
			Call ReportStep (StatusTypes.Fail, "CSV Excel Sheet Name is diffrent wrt Application Report Creation Page Sheet Name " &sheetName, "")	
		End If
		
		intRow_Count = ObjExcelSheet.UsedRange.Rows.Count 
		intColumn_Count = ObjExcelSheet.UsedRange.Columns.Count
		
		'Validating CSV Excel Sheet data and Group Table Dimensions Values
		Call ReportStep (StatusTypes.Pass,"#######################Validating CSV Excel Sheet data and Group Table Dimensions Values", "")
			Set objGrpTabColData=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrpTableColAttributeData")
			rcAttribute = objGrpTabColData.GetROProperty("rows")
			ccAttribute = objGrpTabColData.GetROProperty("cols")
		
			p=2
			q=1
			For x = 1 To rcAttribute
				For y = 1 To ccAttribute 
					
					strcellvalue=objGrpTabColData.GetCellData(x,y)
					strexcelcellvalue=ObjExcelSheet.cells(p,q).value
					
					If Strcomp(TRIM(strcellvalue),TRIM(strexcelcellvalue))=0 Then
						Call ReportStep (StatusTypes.Pass,"Validation of Report Exported to CSV:-"&Space(5)&"Report Exported Successfully and Validated for the Dimension"&Space(3)&strexcelcellvalue,"Report Creation Page")
					Else
						Call ReportStep (StatusTypes.Fail,"Validation of Report Exported to CSV:-"&Space(5)&"Report not Exported Successfully and Validated for the Dimension"&Space(3)&strexcelcellvalue,"Report Creation Page")	
					End If
					
				Next
			p=p+1	
			Next
		
		'Validating Excel Sheet data and Group Table measure values
		Call ReportStep (StatusTypes.Pass,"#######################Validating Excel Sheet data and Group Table measure values", "")
			Set objGrpTabMeasure=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrpTableMeasureData")
			rcMeasure = objGrpTabMeasure.GetROProperty("rows")
			ccMeasure = objGrpTabMeasure.GetROProperty("cols")
			
			m = 2
			n = 2
			For s = 1 To rcMeasure
				For t = 1 To ccMeasure 
					Browser("Analyzer").Page("ReportCreation").Sync
					Browser("Analyzer").Page("ReportCreation").RefreshObject
					strcellvalue=objGrpTabMeasure.GetCellData(s,t)
					strexcelcellvalue=ObjExcelSheet.cells(m,n).value
					
					If strcellvalue<>"-" AND strexcelcellvalue<>"-" Then
						'strcellvalue=Cint(strcellvalue)
						strcellvalue=Round(strcellvalue,2)
						
						'strexcelcellvalue=Cint(strexcelcellvalue)
						strexcelcellvalue=Round(strexcelcellvalue,2)
						
						If Strcomp(TRIM(strcellvalue),TRIM(strexcelcellvalue))=0 Then
							Call ReportStep (StatusTypes.Pass,"Validation of Report Exported to CSV:-"&Space(5)&"Report Exported Successfully and Validated for the Dimension"&Space(3)&strexcelcellvalue,"Report Creation Page")
						Else
							Call ReportStep (StatusTypes.Fail,"Validation of Report Exported to CSV:-"&Space(5)&"Report  not Exported Successfully and Validated for the Dimension"&Space(3)&strexcelcellvalue,"Report Creation Page")	
						End If
						
					End If
					
					n= n + 1		
					If n = ccMeasure+2 Then
						n = 2
					End If
				Next
				m = m + 1
			Next
		
		'Validating End of comparison of SCA Grp Table and exported Excel sheet data
			If m = 2 + rcMeasure Then
				strRowGrandTotalValue=ObjExcelSheet.cells(m-1,1).value
				strColGrandTotalValue=ObjExcelSheet.cells(1,ccMeasure+1).value
				If Trim(Ucase(strRowGrandTotalValue)) = Trim(Ucase("Grand Total")) and Trim(Ucase(strColGrandTotalValue)) = Trim(Ucase("Grand Total")) Then
					Call ReportStep (StatusTypes.Pass, "Validating end of comparison of SCA Grp Table and exported CSV sheet data","Report Creation Page")
					Call ReportStep (StatusTypes.Pass, "SCA Grp Table and exported CSV sheet data are same","Report Creation Page")
				Else 
					Call ReportStep (StatusTypes.Fail, "Validating end of comparison of SCA Grp Table and exported CSV sheet data","Report Creation Page")
					Call ReportStep (StatusTypes.Fail, "Could not compare CSV data completely","Report Creation Page")
				End If
			
			else
				Call ReportStep (StatusTypes.Fail, "Validating end of comparison of SCA Grp Table and exported CSV sheet data","Report Creation Page")
				Call ReportStep (StatusTypes.Fail, "Could not compare CSV data completely. Data Mismatch","Report Creation Page")
				Call ReportStep (StatusTypes.Fail, "SCA Grp Table and exported CSV sheet data are different","Report Creation Page")
				
			End If
			
		'Excel content validation wrt application group table data - End
		Browser("Analyzer").Page("ReportCreation").Sync
		Browser("Analyzer").Page("ReportCreation").RefreshObject
		wait 2
		
		ObjExcelFile.Saved = True
		ObjExcelFile.Close
		
		objExcel.Quit
		
		Set ObjExcelSheet = Nothing
		Set ObjExcelFile = Nothing
		Set objExcel = Nothing
		
	End Sub
	
	
	'<Composed By : Shweta B Nagaral>	
	'<Performs Group Table Row Page Validation
	'To check "Rows per page "option , ie To check whether report is displaying the rows as per selected in rows per page option
	'<objData: Reference to objData>
	Public Sub GrpTabRowPageValidation(ByRef objData)
		Dim tabAtrributeName, objwebtable, ArrValue_Before, strCellValue_Before, strCellValues_Before, compAttributeRowCnt, compAttributeColCnt
		Dim completeValidation, validate10, validate15, validate20, validate25, validate30, validate50, validate80, validate100, validate200, validateAll
		Dim i, j
		
		tabAtrributeName = "tabGrpTableColAttributeData"
		
		Set objwebtable=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable(tabAtrributeName)
		ArrValue_Before=SCA.Webtable(objwebtable,"Dimension Value table",TRIM("Retriving_DataTableValue"),"Report Creation","","","","")
		For i=0 to ubound(ArrValue_Before,1)
		    For j=0 to ubound(ArrValue_Before,2)
		        strCellValue_Before= ArrValue_Before(i,j)
		        strCellValues_Before=strCellValues_Before&";"&strCellValue_Before
		    Next
		Next
		
		compAttributeRowCnt = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable(tabAtrributeName).GetROProperty("rows")
		compAttributeColCnt = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable(tabAtrributeName).GetROProperty("cols")
			
		Browser("Analyzer").Page("ReportCreation").Sync
		Browser("Analyzer").Page("ReportCreation").RefreshObject

		validate10 =  IMSSCA.Validations.ValidationRowColPage("welRowsPerPage", "wel10", 10, 1, tabAtrributeName, strCellValues_Before, objData)
		If validate10 = 1 Then
			completeValidation= 1
			Call ReportStep (StatusTypes.Pass, "########################validated RowPage with 10 rows", "ReportCreation Page")
			validate15 = IMSSCA.Validations.ValidationRowColPage("welRowsPerPage", "wel15", 15, 1, tabAtrributeName, strCellValues_Before, objData)
			If validate15 = 1 Then
				completeValidation= 2
				Call ReportStep (StatusTypes.Pass, "########################validated RowPage with 15 rows", "ReportCreation Page")
				validate20 = IMSSCA.Validations.ValidationRowColPage("welRowsPerPage", "wel20", 20, 1, tabAtrributeName, strCellValues_Before, objData)
				If validate20 = 1 Then
					completeValidation= 3
					Call ReportStep (StatusTypes.Pass, "########################validated RowPage with 20 rows", "ReportCreation Page")
					validate25 = IMSSCA.Validations.ValidationRowColPage("welRowsPerPage", "wel25", 25, 1, tabAtrributeName, strCellValues_Before, objData)
					If validate25 = 1 Then
						completeValidation= 4
						Call ReportStep (StatusTypes.Pass, "########################validated RowPage with 25 rows", "ReportCreation Page")
						validate30 = IMSSCA.Validations.ValidationRowColPage("welRowsPerPage", "wel30", 30, 1, tabAtrributeName, strCellValues_Before, objData)
						If validate30 = 1 Then
							completeValidation= 5
							Call ReportStep (StatusTypes.Pass, "########################validated RowPage with 30 rows", "ReportCreation Page")
							validate50 = IMSSCA.Validations.ValidationRowColPage("welRowsPerPage", "wel50", 30, 1, tabAtrributeName, strCellValues_Before, objData)
							If validate50 = 1 Then
								completeValidation= 6
								Call ReportStep (StatusTypes.Pass, "########################validated RowPage with 50 rows", "ReportCreation Page")
								validate80 = IMSSCA.Validations.ValidationRowColPage("welRowsPerPage", "wel80", 80, 1, tabAtrributeName, strCellValues_Before, objData)
								If validate80 = 1 Then
									completeValidation= 7
									Call ReportStep (StatusTypes.Pass, "########################validated RowPage with 80 rows", "ReportCreation Page")
									validate100 = IMSSCA.Validations.ValidationRowColPage("welRowsPerPage", "wel100", 100, 1, tabAtrributeName, strCellValues_Before, objData)
									If validate100 = 1 Then
										completeValidation= 8
										Call ReportStep (StatusTypes.Pass, "########################validated RowPage with 100 rows", "ReportCreation Page")
										validate200 = IMSSCA.Validations.ValidationRowColPage("welRowsPerPage", "wel200", 200, 1, tabAtrributeName, strCellValues_Before, objData)
										If validate200 = 1 Then
											completeValidation= 9
											Call ReportStep (StatusTypes.Pass, "########################validated RowPage with 200 rows. Validating last pagination option 'ALL'", "ReportCreation Page")
											validateAll = IMSSCA.Validations.ValidationRowColPage("welRowsPerPage", "welAll", compAttributeRowCnt-1, 1, tabAtrributeName, strCellValues_Before, objData)
											If validateAll = 1 Then
												completeValidation= 10
												Call ReportStep (StatusTypes.Pass, "########################validated RowPage for all rows options successfully", "ReportCreation Page")
											End If										
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
			
		If completeValidation <> 10 Then
			Call ReportStep (StatusTypes.Pass, "########################validated Rows Per Page for first "&completeValidation&" options successfully", "ReportCreation Page")
			validateAll = IMSSCA.Validations.ValidationRowColPage("welRowsPerPage", "welAll", compAttributeRowCnt-1, 1, tabAtrributeName, strCellValues_Before, objData)
			If validateAll = 1 Then
				Call ReportStep (StatusTypes.Pass, "########################validated Rows Per Page for 'ALL' rows options successfully", "ReportCreation Page")
			End If
		End If
		
	End Sub
	
	
	'<Composed By : Shweta B Nagaral>	
	'<Performs Group Table Column Page Validation
	'To check "Column per page "option , ie To check whether report is displaying the columns as per selected in cols per page option
	'<objData: Reference to objData>
	Public Sub GrpTabColPageValidation(ByRef objData)
		Dim tabAtrributeName, objwebtable, ArrValue_Before, strCellValue_Before, strCellValues_Before, compMeasureRowCnt, compMeasureColCnt
		Dim completeValidation, validate10, validate15, validate20, validate25, validate30, validate50, validate80, validate100, validate200, validateAll
		Dim i, j, lenBefore
		
		tabAtrributeName = "tabGrpTableMeasureData"
		
		Set objwebtable=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable(tabAtrributeName)
		ArrValue_Before=SCA.Webtable(objwebtable,"Dimension Value table",TRIM("Retriving_DataTableValue"),"Report Creation","","","","")
		For i=0 to ubound(ArrValue_Before,1)
		    For j=0 to ubound(ArrValue_Before,2)
		        strCellValue_Before= ArrValue_Before(i,j)
		        strCellValues_Before=strCellValues_Before&";"&strCellValue_Before
		        lenBefore = Len(strCellValues_Before)
		    Next
		Next
		
		compMeasureRowCnt = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable(tabAtrributeName).GetROProperty("rows")
		compMeasureColCnt = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable(tabAtrributeName).GetROProperty("cols")
			
		Browser("Analyzer").Page("ReportCreation").Sync
		Browser("Analyzer").Page("ReportCreation").RefreshObject
		
		validate10 =  IMSSCA.Validations.ValidationRowColPage("welColumnsPerPage", "wel10", compMeasureRowCnt, 10, tabAtrributeName, lenBefore, objData)
		If validate10 = 1 Then
			completeValidation= 1
			Call ReportStep (StatusTypes.Pass, "########################validated Columns Per Page with 10 Cols", "ReportCreation Page")
			validate15 = IMSSCA.Validations.ValidationRowColPage("welColumnsPerPage", "wel15", compMeasureRowCnt, 15, tabAtrributeName, lenBefore, objData)
			If validate15 = 1 Then
				completeValidation= 2
				Call ReportStep (StatusTypes.Pass, "########################validated Columns Per Page with 15 Cols", "ReportCreation Page")
				validate20 = IMSSCA.Validations.ValidationRowColPage("welColumnsPerPage", "wel20", compMeasureRowCnt, 20, tabAtrributeName, lenBefore, objData)
				If validate20 = 1 Then
					completeValidation= 3
					Call ReportStep (StatusTypes.Pass, "########################validated Columns Per Page with 20 Cols", "ReportCreation Page")
					validate25 = IMSSCA.Validations.ValidationRowColPage("welColumnsPerPage", "wel25", compMeasureRowCnt, 25, tabAtrributeName, lenBefore, objData)
					If validate25 = 1 Then
						completeValidation= 4
						Call ReportStep (StatusTypes.Pass, "########################validated Columns Per Page with 25 Cols", "ReportCreation Page")
						validate30 = IMSSCA.Validations.ValidationRowColPage("welColumnsPerPage", "wel30", compMeasureRowCnt, 30, tabAtrributeName, lenBefore, objData)
						If validate30 = 1 Then
							completeValidation= 5
							Call ReportStep (StatusTypes.Pass, "########################validated Columns Per Page with 30 Cols", "ReportCreation Page")
							validate50 = IMSSCA.Validations.ValidationRowColPage("welColumnsPerPage", "wel50", compMeasureRowCnt, 50, tabAtrributeName, lenBefore, objData)
							If validate50 = 1 Then
								completeValidation= 6
								Call ReportStep (StatusTypes.Pass, "########################validated Columns Per Page with 50 Cols", "ReportCreation Page")
								validate80 = IMSSCA.Validations.ValidationRowColPage("welColumnsPerPage", "wel80", compMeasureRowCnt, 80, tabAtrributeName, lenBefore, objData)
								If validate80 = 1 Then
									completeValidation= 7
									Call ReportStep (StatusTypes.Pass, "########################validated Columns Per Page with 80 Cols", "ReportCreation Page")
									validate100 = IMSSCA.Validations.ValidationRowColPage("welColumnsPerPage", "wel100", compMeasureRowCnt, 100, tabAtrributeName, lenBefore, objData)
									If validate100 = 1 Then
										completeValidation= 8
										Call ReportStep (StatusTypes.Pass, "########################validated Columns Per Page with 100 Cols", "ReportCreation Page")
										validate200 = IMSSCA.Validations.ValidationRowColPage("welColumnsPerPage", "wel200", compMeasureRowCnt, 200, tabAtrributeName, lenBefore, objData)
										If validate200 = 1 Then
											completeValidation= 9
											Call ReportStep (StatusTypes.Pass, "########################validated Columns Per Page with cols. Validating last pagination option 'ALL'",  "ReportCreation Page")
											validateAll = IMSSCA.Validations.ValidationRowColPage("welColumnsPerPage", "welAll", compMeasureRowCnt, compMeasureColCnt, tabAtrributeName, lenBefore, objData)
											If validateAll = 1 Then
												completeValidation= 10
												Call ReportStep (StatusTypes.Pass, "########################validated Columns Per Page for 'ALL' Cols options successfully", "ReportCreation Page")
											End If										
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
			
		If completeValidation <> 10 Then
			Call ReportStep (StatusTypes.Pass, "########################validated Columns Per Page for first "&completeValidation&" options successfully", "ReportCreation Page")
			validateAll = IMSSCA.Validations.ValidationRowColPage("welColumnsPerPage", "welAll", compMeasureRowCnt, compMeasureColCnt-1, tabAtrributeName, lenBefore, objData)
			If validateAll = 1 Then
				Call ReportStep (StatusTypes.Pass, "########################validated Columns Per Page for 'ALL' cols options successfully", "ReportCreation Page")
			End If
		End If
		
	End Sub
	
	
 	'''<To Create the GroupFilter group depending on the selection>
    '''<strO_selection:- To select the operation on Group>
    '''<strName :- Name of the Group>
   	'''<strCaption:- Name of the Group Caption>
   	'''<strDisplayVal:-Display check box selection>
   	'''<strBehaviour:-Behaviour selection>  
   	'''<strmemberselection: Memberrs of the group>
   	'''<	strlistSelection:Selecting the Filter Values either equal or starts with
   	'''< intlineDelimiterselection: line delimiter selection> 
   	'''< Author>shobha<Author> 
	Public Sub FilterGroupCreation(ByVal strName,ByVal strCaption,byval strDisplayVal,ByVal strFilterVal,ByVal strlistSelection,ByVal strmemberselection, ByVal chkVisible)
	    
	    Dim objtxtname,objtxtcaption,objrdoDisplay,objrdotype,imgobjEditmem,objframechk,objchkcount,objbtnok,objFieldok,intR_Val
	    Dim i,j, objexpand, objVisible,objtxtSelection,btnAddcriteria,objddlselection, objEditMembers
	    
	    wait 1
	    Browser("Analyzer").Page("GroupCreationPage").Sync
	    Browser("Analyzer").Page("GroupCreationPage").RefreshObject
	    
	    Set objtxtname=Browser("Analyzer").Page("GroupCreationPage").Frame("GroupFilterCreation").WebEdit("txtName")
	    Call SCA.SetText(objtxtname,strName, "txtGroupName","Filter GroupCreation Dialouge" )
	    
	    Set objtxtcaption=Browser("Analyzer").Page("GroupCreationPage").Frame("GroupFilterCreation").WebEdit("txtCaption")
	    Call SCA.SetText(objtxtcaption,strCaption, "txtGroupCaption","Filter GroupCreation Dialouge" )
	    
	    '<shweta-2/6/2015- Start>
	    Set objVisible=Browser("Analyzer").Page("GroupCreationPage").Frame("GroupFilterCreation").WebCheckBox("chkVisible")
	    Call SCA.SetCheckBox(objVisible,"Visible Check Box", chkVisible, "Group Filter Page")
	    '<shweta-2/6/2015- End>
	     
	    Set objrdoDisplay=Browser("Analyzer").Page("GroupCreationPage").Frame("GroupFilterCreation").WebRadioGroup("rdoDisplayas")
	    Call SCA.ClickOnRadio(objrdoDisplay,strDisplayVal ,"Filter GroupCreation Dialouge")
	    
	    Set objrdotype=Browser("Analyzer").Page("GroupCreationPage").Frame("GroupFilterCreation").WebRadioGroup("rdoFiltertypes")
	    Call SCA.ClickOnRadio(objrdotype,strFilterVal ,"Filter GroupCreation Dialouge")  
	    
	    
	    If strlistSelection<>"" Then
	    	
		    Set objddlselection=Browser("Analyzer").Page("GroupCreationPage").Frame("GroupFilterCreation").WebList("ddlLabel")
			Call SCA.SelectFromDropdown(objddlselection,strlistSelection)
			
			Set objtxtSelection=Browser("Analyzer").Page("GroupCreationPage").Frame("GroupFilterCreation").WebEdit("txtFirstLabel")
			Call SCA.SetText(objtxtSelection,strmemberselection, "txtMemberSelection","Filter GroupCreation Dialouge" )
				
			Set btnAddcriteria=Browser("Analyzer").Page("GroupCreationPage").Frame("GroupFilterCreation").WebButton("btnAddCriteria")
			Call SCA.ClickOn(btnAddcriteria,"Add Criteria", "Filter GroupCreation Dialouge")
		
		Else
			wait 4
			'Set objEditMembers=Browser("Analyzer").Page("GroupCreationPage").Frame("GroupFilterCreation").Image("EditMembers")
			Set objEditMembers=Browser("micclass:=browser").Page("micclass:=page").Frame("html id:=__dialog").WebElement("xpath:=//img[@id='btnMember']")
		    Call SCA.ClickOn(objEditMembers,"EditMembers", "Filter GroupCreation Dialouge")
		    'Set imgobjEditmem=
		    'Call SCA.ClickOn(imgobjEditmem,"btnoK", "Filter GroupCreation Dialouge")
		    wait 4
			Set objframechk=Description.Create
			objframechk("micclass").value="WebCheckBox"
			objframechk("type").value="checkbox"
			'objframechk("xpath").value="//input[contains(@id,'lstMembers_')]"
			
			'Set objchkcount=Browser("Analyzer").Page("GroupCreationPage").Frame("EditGroupDialog").ChildObjects(objframechk)
			'Set objchkcount=Browser("Hosting Templates - Ops").Page("Analyzer_2").Frame("Frame").WebTable("Angiologue").ChildObjects(objframechk)
			Set objchkcount=Browser("Hosting Templates - Ops").Page("Analyzer_2").Frame("Frame").WebTable("Equals Does not equal").ChildObjects(objframechk)
'			msgbox objchkcount.count
			'Browser("micclass:=browser").Page("micclass:=page").Frame("html id:=__dialog").WebElement("xpath:=//tr//div//table[@id='lstMembers']").highlight
'			Set objchkcount=Browser("micclass:=browser").Page("micclass:=page").Frame("xpath:=//iframe[contains(@src,'popup/PropertySelector.aspx')]").ChildObjects(objframechk)
			'msgbox objchkcount.count
			'Browser("Analyzer").Page("GroupCreationPage").Frame("FrameFilterMembers").WebTable("Angiologue").GetCellData
			For i = 1 To strmemberselection Step 1



				objchkcount(i).click
		
			Next
			 
			wait 8	 
			'changes by srinivas
			Set objbtnok=Browser("Hosting Templates - Ops").Page("Analyzer_2").Frame("Frame").WebButton("OkButtonInSelectchkboxes")

			'Set objbtnok=Browser("Analyzer").Page("GroupCreationPage").Frame("FrameFilterMembers").WebButton("btnOK")
			Call SCA.ClickOn(objbtnok,"objbtnok", "Filter GroupCreation Dialouge") 
			
	    End If
	    
	    wait 8
	   ' msgbox "here"
	   ' Set objFieldok=Browser("Analyzer").Page("GroupCreationPage").Frame("GroupFilterCreation").WebButton("FieldOk")
	   	'Browser("micclass:=browser").Page("micclass:=page").Frame("html id:=__dialog").WebElement("xpath:=//input[@id='btnOk' and @type='submit']").Click
	    'Call SCA.ClickOn(objFieldok,"objFieldok", "Filter GroupCreation Dialouge")    
	    Set objFieldok=Browser("Analyzer").Page("GroupCreationPage").Frame("__dialogAddAllMembers").WebButton("btnOK")
		Call SCA.ClickOn(objFieldok,"objFieldok", "Filter GroupCreation Dialouge")
	    intR_Val=IMSSCA.Validations.Groupcreation_Verification(strName,strCaption,"")    
	    If intR_Val=0 Then
	        Call ReportStep (StatusTypes.Pass, "Create Group Validation:-"&Space(7)&strName&Space(5)&"Group is created and Reflected in the Group dialouge box"&Space(5)&strCaption&Space(5)&"created in the Report creation area","GroupCreation Dialouge box")
	    ElseIf intR_Val=1 Then
	        Call ReportStep (StatusTypes.Pass, "Create Group Validation:-"&Space(7)&strName&Space(5)&"Group is created and Reflected in the Group dialouge box"&Space(5)&strCaption&Space(5)&"not created in the Report creation area","GroupCreation Dialouge box")
	    else
	        Call ReportStep (StatusTypes.Fail, "Create Group Validation:-"&Space(7)&strName&Space(5)&"Group is not created and not Reflected in the Group dialouge box"&Space(5)&strCaption&Space(5)&"not created in the Report creation area","GroupCreation Dialouge box")
	    End If
	
		'<shweta-2/6/2015- Start>
	    If chkVisible = "ON" Then
			For j = 1 To 10 Step 1
				If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgGexpand").Exist(2) Then
		    	 	Call ReportStep (StatusTypes.Pass, "New Group "&strName&" Created is visibe in Homepage after enabling 'Visible' checkbox in Add Filter Page", "ReportCreation Page")
		    	 	Set objexpand=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgGexpand")
	   		 		Call SCA.ClickOn(objexpand,"expand", "Report Creation Dialouge")
	   		 		Exit for
		    	ElseIf intR_Val=1 Then
	            	Exit For
		    	End If
		    	
		    	If j = 10 Then
		    		Call ReportStep (StatusTypes.Fail, "New Group "&strName&" Created is not visibe in Homepage after enabling 'Visible' checkbox in Add Filter Page", , "ReportCreation Page")
		    	End If
	    	Next
	   	 
	   	ElseIf chkVisible = "OFF" Then
				If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgGexpand").Exist(2) <> True  Then
	    	 	Call ReportStep (StatusTypes.Pass, "New Group "&strName&" Created is not visibe in Homepage after disabling 'Visible' checkbox in Add Filter Page", "ReportCreation Page")
			Else
				Call ReportStep (StatusTypes.Fail, "New Group "&strName&" Created is visibe in Homepage after disabling 'Visible' checkbox in Add Filter Page", "ReportCreation Page")
	    	End If
		    	
	    End If
	   '<shweta-2/6/2015- End>
	    
	End Sub

	
	'<Changes Component Theme of report created>
	'<strColorTheme : Report Theme Color: String>
	'<strCompColorTheme : Component Specific Color theme: String>
	'<strReportName : ReportName for which general tab has to be configured>
	'<strPagename : Page name where Report Properties should be accessed>
	'<objData : Reference to objData>
	Public Sub ChangeTheme(ByVal strColorTheme, ByVal strCompColorTheme, ByVal strReportName, ByVal strPagename, ByRef objData)

		Dim objFrame, FrameName, objbutton, objRdoConst
	
		If strCompColorTheme = "" Then
			Set objFrame = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer")
			FrameName = objFrame.GetROProperty("name")
			Call ReportToolBar("colortheme.svg", FrameName)
		else
			Call DownNormalMenu(objData.item("TabName"),objData.item("CompSelectValue"), objData.item("SelectSubValue"))
		End If
		
		wait 2
		Browser("Analyzer").Page("ReportCreation").Sync
		Set objRdoConst=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebRadioGroup("rdoTheme")
		
		'Select color theme
		Select Case strColorTheme
			
			Case "Blue"
				'Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebRadioGroup("rdoTheme").Select 0
				Call SCA.ClickOnRadio(objRdoConst,"0", strPagename)
			Case "Cyan"
				'Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebRadioGroup("rdoTheme").Select 1
				Call SCA.ClickOnRadio(objRdoConst,"1", strPagename)
			Case "Green"
				'Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebRadioGroup("rdoTheme").Select 2
				Call SCA.ClickOnRadio(objRdoConst,"2", strPagename)
			Case "Orange"
				'Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebRadioGroup("rdoTheme").Select 3
				Call SCA.ClickOnRadio(objRdoConst,"3", strPagename)
			Case "Purple"
				'Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebRadioGroup("rdoTheme").Select 4
				Call SCA.ClickOnRadio(objRdoConst,"4", strPagename)
			Case "Red"
				'Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebRadioGroup("rdoTheme").Select 5
				Call SCA.ClickOnRadio(objRdoConst,"5", strPagename)
			Case "Silver"
				'Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebRadioGroup("rdoTheme").Select 6
				Call SCA.ClickOnRadio(objRdoConst,"6", strPagename)
		End Select
		
		Set objbutton=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebButton("btnOK")
		Call SCA.ClickOn(objbutton,"Color Theme change Ok Button", "Setting Color Theme for Report")
		
		Browser("Analyzer").Page("ReportCreation").Sync

	End Sub
	
	
	'<Sets report properties>
	'<strReportName : ReportName for which general tab has to be configured>
	'<strPagename : Page name where Report Properties should be accessed>
	'<objData : Reference to objData>
	Public Sub SetReportProperties(ByVal strReportName, ByVal strPagename, ByRef objData)

		Dim objFrame, FrameName, objReportNameTxt, objReportDescription, objReportVideoOrDoc, objchkAutoRefreshReport, objAutoRefreshInterval, objbutton
	
		Set objFrame = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer")
		FrameName = objFrame.GetROProperty("name")
		Call ReportToolBar("edit.svg", FrameName)
		 		
		Set objReportNameTxt=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebEdit("txtReportName")
		Call SCA.SetText(objReportNameTxt, objData.Item("ReportName"), "txtReportName", strPagename)
			
		Set objReportDescription = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebEdit("txtReportPropertyDescription")
		Call SCA.SetText_MultipleLineArea(objReportDescription, objData.Item("ReportDescription"), "txtReportPropertyDescription", strPagename)
				
		Set objReportVideoOrDoc = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebEdit("txtVideoOrDoc")
		Call SCA.SetText(objReportVideoOrDoc, objData.Item("RelatedVideoDocPath"), "txtVideoOrDoc", strPagename)
				
'		Set objReportAccessUrl = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebEdit("txtReportAccessUrl")
'		Call SCA.SetText(objReportAccessUrl, objData.Item("RelatedVideoDocPath"), "txtReportAccessUrl", "Report Properties")
				
		Set objchkAutoRefreshReport = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebCheckBox("chkAutoRefreshReport")
		Call SCA.SetCheckBox(objchkAutoRefreshReport, "chkAutoRefreshReport", objData.Item("AutoRefreshReportCheckBox"), strPagename)
							
		Set objAutoRefreshInterval = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebEdit("txtAutoRefreshInterval")
		Call SCA.SetText(objAutoRefreshInterval, objData.Item("AutoRefreshInterval"), "txtAutoRefreshInterval", strPagename)

		Set objbutton = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebButton("btnOK")
		Call SCA.ClickOn(objbutton,"Report Properties Ok Button", strPagename)
				
	End Sub
	
	
	'< To select the Menu Items from measure field>
    '<ColVal: Measure Column Count from where menu table has to be clicked:  Number>
    '<strTabName: Tab name of measure menu table. It could be "Sort", "Calculations", "Measures", "Visualization Rules":  String>
    '<strSelValue: Select options under each tab :  String>
    '<strSelSubValue: Select sub-options inside options under each tab :  String>
	Public Sub MeasureDownNormal(ByVal ColVal, ByVal strTabName, ByVal strSelValue, ByVal strSelSubValue)
		
		Dim oMeasureDownNormal, oMeasureDownNormalObj, objSelSubValue, objSelValue
	    wait 10
		Set oMeasureDownNormal = Description.Create()
		oMeasureDownNormal("micclass").value = "Image"
		oMeasureDownNormal("html id").value = ""
		oMeasureDownNormal("html tag").value = "IMG"
		oMeasureDownNormal("image type").value = "Plain Image"
		oMeasureDownNormal("outerhtml").regularexpression=True
		oMeasureDownNormal("outerhtml").value = ".*onMeasureContextClick.*"
		
		Set oMeasureDownNormalObj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDropColDimension").ChildObjects(oMeasureDownNormal)
				
		Browser("Analyzer").Page("ReportCreation").Sync
		'Mouse over on Pivot Table menu
		wait 4
		oMeasureDownNormalObj(ColVal).FireEvent "onmouseover"
		
		oMeasureDownNormalObj(ColVal).Click
		wait 6
		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image(strTabName).FireEvent "onmouseover"
		wait 3
		
		'First part of if condition will be executed, if sub-options has to be selected
		If strSelSubValue <> "" Then
			Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strSelValue).FireEvent "onmouseover"
			wait 2
			Set objSelSubValue = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strSelSubValue)
			Call SCA.ClickOn(objSelSubValue,strSelSubValue, "Report Creation Page")
		Else
			'Clicks on options options under each tab
			Set objSelValue = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strSelValue)
			Call SCA.ClickOn(objSelValue,strSelValue, "Report Creation Page")
		End If
		wait 2
		
	End Sub
	
	
	
	'< Define alerts for Pivot table created>
	'<strFormulaName: Alert Formula Name:  String>
	'<strFormula: Select Measure value to apply Alert Formula:  String>
	'<strFormulaOperator: Select Formula operator for Measure value:  String>
	'<rdoConstant: Select Constant Value radio button:  String>
	'<rdoConstantValue: Constant Value:  String>
    '<strColor: Select background color to highlight the alert created:  String>
    '<strReportName : ReportName for which general tab has to be configured>
    '<strPagename: Page name where Alet has to be created :  String>
    '<objData: Refrence to objData :  >
	Public Sub DefineAlerts(ByVal strFormulaName, ByVal strFormula, ByVal strFormulaOperator, ByVal rdoConstant, ByVal rdoConstantValue, ByVal strColor, ByVal strReportName, ByVal strPagename,  ByVal rdoConstantValue2,ByRef objData)
	
		Dim opalette, opaletteObj, strOuterHtml, objOk, objtxtBMarkName, objlstDdlMeasureA, objlstDdlOperator, objrdoConstant, objtxtConstB, objImgPalette, objtxtDefineAlertsBckGrndColor
		
		Set objtxtBMarkName=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebEdit("txtBMarkName")
		Call SCA.SetText(objtxtBMarkName, strFormulaName, "txtBMarkName", strPagename)
		
		Set objlstDdlMeasureA=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebList("lstDdlMeasureA")
		Call SCA.SelectFromDropdown(objlstDdlMeasureA,strFormula)
		
		Set objlstDdlOperator=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebList("lstDdlOperator")
		Call SCA.SelectFromDropdown(objlstDdlOperator,strFormulaOperator)
		
		Set objrdoConstant=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebRadioGroup("rdoConstant")
   		Call SCA.ClickOnRadio(objrdoConstant,rdoConstant, strPagename)
   		
   		Set objtxtConstB=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebEdit("txtConstB")
   		Call SCA.SetText(objtxtConstB, rdoConstantValue, "txtConstB", strPagename)
   		
   		If strFormulaOperator= "Between" Then
   		   Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebRadioGroup("type:=radio","html tag:=INPUT","name:=c").Select "rdoConstC"
   			Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebEdit("html id:=txtConstC","html tag:=INPUT","type:=text").Set rdoConstantValue2
   		End If
   		
   		
		set objImgPalette=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").Image("imgPalette")
		Call SCA.ClickOn(objImgPalette,"imgPalette",strPagename)
		wait 2
		If Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebTable("tabDefault").Exist(10) Then
			'<Shweta - 11/8/2015> - Start
'			Set opalette = Description.Create()
'			opalette("micclass").value = "WebElement"
'			opalette("html id").value = ""
'			opalette("html tag").value = "TD"
'			opalette("outerhtml").regularexpression=True
'			opalette("outerhtml").value = ".*bgColor.*"&strColor&".*"
'										
'			Set opaletteObj = 	Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebTable("tabDefault").ChildObjects(opalette)
'			opaletteObj(0).FireEvent "onmouseover"
'			wait 2
'			opaletteObj(0).Click
			'<Shweta - 11/8/2015> - End
			Set objtxtDefineAlertsBckGrndColor = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebEdit("txtDefineAlertsBckGrndColor")
			Call SCA.SetText(objtxtDefineAlertsBckGrndColor, strColor, "txtDefineAlertsBckGrndColor", strPagename)
			
			Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebButton("btnSelect").Click
			wait 1
		Else
			Call ReportStep (StatusTypes.Fail, "Could not display color table to select color "&strColor& " for report " &strReportName, strPagename)
		End If
		wait 2
		' = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebElement("welBackgroundColor").GetROProperty("outerhtml")

		strOuterHtml = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebElement("welBackgroundColor").Object.style.backgroundColor
		'strOuterHtml1=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebElement("welBackgroundColor").object.getAttribute("backgroundcolor")
		'msgbox strOuterHtml1
'		s=""&col&""
'		a=split(s,")")
'		n=replace(a(0),"RGB","")
'		t=replace(n,"(","")
'		res=split(t,",")
'		'rgb values
'		r=res(0)
'		g=res(1)
'		b=res(2)
'		strOuterHtml=RGB(r,g,b)
		'msgbox strOuterHtml

		If InStr(1, trim(strOuterHtml), "rgb(174, 255, 128)") <> 0  Then
			Call ReportStep (StatusTypes.Pass, "Successfully selected background color as "&strColor, strPagename)
		Else
			Call ReportStep (StatusTypes.Fail, "Could not successfully select background color as "&strColor, strPagename)
		End If
	
		Set objOk=Browser("Analyzer").Page("ReportCreation").Frame("CreateChart").WebButton("btnOk")
   		Call SCA.ClickOn(objOk,"OkButton" , "ReportCreation")
   		
   		Browser("Analyzer").Page("ReportCreation").Sync
   		
	End Sub
	
	
	
	'<DataSourceFunctions redirects Home Page to Manage data source>
	'<strSrcName: Mention Data Source name :  String>
    '<objData: Refrence to objData :  >
	Public Sub DataSourceFunctions(ByVal strSrcName, ByRef objData)
		
		Dim rc, cc, i, strDataSourceName		
		
		rc = Browser("Analyzer").Page("Home").Frame("Frame").WebTable("tabDataSource").GetROProperty("rows")
		cc = Browser("Analyzer").Page("Home").Frame("Frame").WebTable("tabDataSource").GetROProperty("cols")
		j=1
		Do
		  If Browser("Hosting Templates - Ops").Page("Analyzer").Frame("Frame").Image("next_enabled").Exist(10) AND j<>1 Then
		  	 Browser("Hosting Templates - Ops").Page("Analyzer").Frame("Frame").Image("next_enabled").Click
		  End If
		  For i = 1 To rc
			strDataSourceName = Browser("Analyzer").Page("Home").Frame("Frame").WebTable("tabDataSource").GetCellData(i, 2)
			If InStr(1, strDataSourceName, strSrcName) <> 0 Then
				Browser("Analyzer").Page("Home").Frame("Frame").WebTable("tabDataSource").ChildItem(i, 1, "WebCheckBox",0).click
				Browser("Analyzer").Page("Home").Frame("Frame").WebTable("tabDataSource").ChildItem(i, 2, "WebElement",0).click
				Exit Do
		
			End If
		  Next
		   j=j+1
		Loop While Browser("Hosting Templates - Ops").Page("Analyzer").Frame("Frame").Image("next_enabled").Exist(10)
		
		
		Browser("Analyzer").Page("Home").Sync

	End Sub
	
	Public Sub GroupPivotTable(ByVal strGrpName, ByVal strPagename, ByRef objData)
		Dim oTable, oTableObj, oSetGroup, oSetGroupObj, i, j, GroupPivotTable, objTabGroup, rc, cc, tabGroupVal
		GroupPivotTable = 0
		'Descriptive programme to Count No of Pivot Tables
		Set oTable = Description.Create()
		oTable("micclass").value = "WebTable"
		oTable("cols").value=4
		oTable("column names").regularexpression=True
		oTable("column names").value = ".*PivotTable.*"
		oTable("html tag").value = "TABLE"
		oTable("name").value = "down_normal"
		oTable("innertext").regularexpression=True
		oTable("innertext").value = "PivotTable.*"
		
		Set oTableObj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabPivotTable").ChildObjects(oTable)
		'msgbox oTableObj.count
		
		'Descriptive programme to Count No of Grouping Images in Tables
		Set oSetGroup = Description.Create()
		oSetGroup("micclass").value = "Image"
		oSetGroup("file name").regularexpression=True
		oSetGroup("file name").value = "g0\.gif"
		oSetGroup("html tag").value = "IMG"
		oSetGroup("image type").value = "Plain Image"
		oSetGroup("name").value = "Image"
		oSetGroup("title").value = "Set Group"
		
		Set oSetGroupObj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabPivotTable").ChildObjects(oSetGroup)
		'msgbox oSetGroupObj.count
		
		For i = 0 To oSetGroupObj.count-1
			oSetGroupObj(i).Click
			wait 2
			set objTabGroup = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGroup")
			rc = objTabGroup.GetROProperty("rows")
			cc = objTabGroup.GetROProperty("cols")
			For j = 1 To rc
				tabGroupVal = objTabGroup.GetCellData(j, 2)
				If InStr(1, Trim(Ucase(tabGroupVal)), Trim(Ucase(strGrpName)) ) > 0 Then
					'msgbox "Found"
					Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGroup").ChildItem(j, 0, "WebElement", 2).Click
					wait 2
					Browser("Analyzer").Page("ReportCreation").Sync
					GroupPivotTable = GroupPivotTable + 1
					Exit For
				End If 
			Next
		Next
		
		If GroupPivotTable = oSetGroupObj.count Then
			Call ReportStep (StatusTypes.Pass, "Successfully selected "&strGrpName&" for "&oTableObj.count& " Pivot Tables", strPagename)
		Else
			Call ReportStep (StatusTypes.Fail, "Could not select "&strGrpName&" for "&oTableObj.count& " Pivot Tables", strPagename)
			Call ReportStep (StatusTypes.Fail, "Could able to select " &oSetGroupObj.count&space(5)&strGrpName&" out of "&oTableObj.count& " Pivot Tables", strPagename)
		End If
		
	End Sub
	
'	'< To select the Menu Items from measure field>
'    '<ColVal: Measure Column Count from where menu table has to be clicked:  Number>
'    '<strTabName: Tab name of measure menu table. It could be "Sort", "Calculations", "Measures", "Visualization Rules":  String>
'    '<strSelValue: Select options under each tab :  String>
'    '<strSelSubValue: Select sub-options inside options under each tab :  String>
'	Public Sub FilterDownNormal(ByVal strfilterValue, Byval strSelValue, ByVal strSelSubValue)
'		
'		Dim oFilterCond, oFilterCondObj, oFilter, oFilterObj, objSelSubValue, objSelValue, oFilterList, oFilterListObj, i
'	
'		'Descriptive programme to Count No of Filter Table
'		Set oFilterCond = Description.Create()
'		oFilterCond("micclass").value = "WebTable"
'		oFilterCond("cols").value=1
'		oFilterCond("rows").value=1
'		oFilterCond("column names").regularexpression=True
'		oFilterCond("column names").value = strfilterValue
'		oFilterCond("html tag").value = "TABLE"
'		oFilterCond("name").value = "down_normal"
'		oFilterCond("innertext").regularexpression=True
'		oFilterCond("innertext").value = strfilterValue&".*"
'		
'		Set oFilterCondObj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabPivotTable").ChildObjects(oFilterCond)
'		'msgbox oFilterCondObj.count
'		
'		For i = 0 To oFilterCondObj.count-1
'			'Descriptive programme to Count No of Filter Table
'			Set oFilter = Description.Create()
'			oFilter("micclass").value = "Image"
'			oFilter("file name").regularexpression=True
'			oFilter("file name").value="down_normal\.gif"
'			oFilter("html tag").value = "IMG"
'			oFilter("image type").value="Plain Image"
'			oFilter("name").value = "Image"
'										
'			Set oFilterObj = oFilterCondObj(i).ChildObjects(oFilter)
'			'msgbox oFilterObj.count
'			
'			oFilterObj(0).FireEvent "onmouseover"
'			oFilterObj(0).Click
'			
'			'First part of if condition will be executed, if sub-options has to be selected
'			If strSelSubValue <> "" Then
'				Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strSelValue).FireEvent "onmouseover"
'				wait 1
'				Set objSelSubValue = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strSelSubValue)
'				Call SCA.ClickOn(objSelSubValue,strSelSubValue, "Report Creation Page")
'			Else
'				'Clicks on options options under each tab
'				Set objSelValue = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strSelValue)
'				Call SCA.ClickOn(objSelValue,strSelValue, "Report Creation Page")
'			End If
'			
'			wait 2
'			Browser("Analyzer").Page("ReportCreation").Sync
'		Next
'		
'		
''		'Descriptive programme to Select Filter Value from List - Start
''		Set oFilterList= Description.Create()
''		oFilterList("micclass").value = "WebList"
''		oFilterList("html tag").value = "SELECT"
''		oFilterList("name").value = "select"
''		oFilterList("select type").value= "ComboBox Select"
''									
''		Set oFilterListObj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabPivotTable").ChildObjects(oFilterList)
''		msgbox oFilterListObj.count
''		Call SCA.SelectFromDropdown(oFilterListObj(0),strFilterVal)
''		
''		wait 2
''		Browser("Analyzer").Page("ReportCreation").Sync
''		'Descriptive programme to Select Filter Value from List - End
'		
'	End Sub
	
	
	'< To select the Menu Items from measure field>
    '<ColVal: Measure Column Count from where menu table has to be clicked:  Number>
    '<strTabName: Tab name of measure menu table. It could be "Sort", "Calculations", "Measures", "Visualization Rules":  String>
    '<strSelValue: Select options under each tab :  String>
    '<strSelSubValue: Select sub-options inside options under each tab :  String>
	Public Sub FilterDownNormal(ByVal strFilterAll, ByVal strFilterName, Byval strFilterMethod, ByVal strfilterSelectValue, ByRef objData)
		
		Dim oFilterCond, oFilterCondObj, oFilter, oFilterObj, objSelSubValue, objSelValue, oFilterList, oFilterListObj, i
	
		'Descriptive programme to Count No of Filter Table
		Set oFilterCond = Description.Create()
		oFilterCond("micclass").value = "WebTable"
		oFilterCond("cols").value=1
		oFilterCond("rows").value=1
		oFilterCond("column names").regularexpression=True
		oFilterCond("column names").value = strFilterName&".*"
		oFilterCond("html tag").value = "TABLE"
		oFilterCond("name").value = "down_normal"
		oFilterCond("innertext").regularexpression=True
		oFilterCond("innertext").value = strFilterName&".*"
		
		Set oFilterCondObj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabPivotTable").ChildObjects(oFilterCond)
		'msgbox oFilterCondObj.count
		
		'Descriptive programme to Count No of "None" to choose filter condition to select from "strFilterName" - Start
				If strFilterAll = "" Then
						Set oFilter = Description.Create()
						oFilter("micclass").value = "WebElement"
						oFilter("html tag").value = "SPAN"
						oFilter("class").value = ""
						oFilter("html id").value = ""
						oFilter("innertext").regularexpression=True
						oFilter("innertext").value = ".*None.*"
						oFilter("outertext").regularexpression=True
						oFilter("outertext").value = ".*None.*"
						oFilter("outerhtml").regularexpression=True
						oFilter("outerhtml").value = ".*COLOR.*666666.*"	
						Set oFilterObj = oFilterCondObj(0).ChildObjects(oFilter)
						'msgbox oFilterObj.count
						oFilterObj(0).FireEvent "onmouseover"
						oFilterObj(0).Click
						wait 1
						Call FilterConditionSettings("Set Filter Member", strFilterMethod, strfilterSelectValue, objData)
				else
				
						For i = 0 To oFilterCondObj.count-1
							Set oFilter = Description.Create()
							oFilter("micclass").value = "WebElement"
							oFilter("html tag").value = "SPAN"
							oFilter("class").value = ""
							oFilter("html id").value = ""
							oFilter("innertext").regularexpression=True
							oFilter("innertext").value = ".*None.*"
							oFilter("outertext").regularexpression=True
							oFilter("outertext").value = ".*None.*"
							oFilter("outerhtml").regularexpression=True
							oFilter("outerhtml").value = ".*COLOR.*666666.*"	
							Set oFilterObj = oFilterCondObj(i).ChildObjects(oFilter)
							'msgbox oFilterObj.count
							oFilterObj(0).FireEvent "onmouseover"
							oFilterObj(0).Click
							wait 1
							'TODO
							call FilterConditionSettings()
						Next
				End If
		'Descriptive programme to Count No of "None" to choose filter condition to select from "strFilterName" - End
		
	End Sub

	Public Sub FilterConditionSettings(Byval strFilterCase, Byval strFilterMethod, ByVal strfilterSelectValue, ByRef objData)
		Dim oFilterVal, oFilterValObj
		
		Select Case strFilterCase
			Case "Set Filter Member"
					Browser("Analyzer").Page("ReportCreation").Frame("FilterConditionSettings").WebList("lstDdlFilterType").Select strFilterMethod		
					Browser("Analyzer").Page("ReportCreation").Frame("FilterConditionSettings").WebElement("welSetFilterMem").Click
					'Set off 
					Browser("Analyzer").Page("ReportCreation").Frame("FilterConditionSettings").Image("imgSelectAll").Click
					wait 2
					
					Set oFilterVal = Description.Create()
					oFilterVal("micclass").value = "WebElement"
					oFilterVal("html tag").value = "SPAN"
					oFilterVal("innertext").regularexpression=True
					oFilterVal("innertext").value = strfilterSelectValue
					
					Set oFilterValObj = Browser("Analyzer").Page("ReportCreation").Frame("FilterConditionSettings").WebTable("Set Filter Member").ChildObjects(oFilterVal)
					'msgbox oFilterValObj.count
					oFilterValObj(0).FireEvent "onmouseover"
					oFilterValObj(0).Click
					'Browser("Analyzer").Page("ReportCreation").Frame("FilterConditionSettings").WebButton("btnOK").Click
										
			Case "Default Value"
					'TODO
		End Select
		
		wait 2
		Browser("Analyzer").Page("ReportCreation").Sync
		Browser("Analyzer").Page("ReportCreation").RefreshObject
		
	End Sub


  	Function Func_GetQueryResult(sQuery,selection)
		Dim con,rs ,loopCounter,var,counter
		counter=0
		Dim resArr()
		Set con=CreateObject("ADODB.Connection")
		Set rs=CreateObject("ADODB.recordset")    
		
	    'con.open"Driver={SQL Server};server=CDTCSIT06I;database=StrategyCompanionAnalyzer" 
	    con.open"Driver={SQL Server};server=10.121.20.11;database=StrategyCompanionAnalyzer"
	    rs.open sQuery,con
	
	    If selection=1 Then
			While not rs.EOF 'Count no of rows in your result
	                	loopCounter=loopCounter+1
				rs.MoveNext
			Wend
	       If loopCounter = 0 Then
		   Reporter.ReportEvent micWarning,"Empty Result Set","Query is: " & sQuery
		   Func_GetQueryResult=False
		   Exit Function
		   End If
	       rs.MoveFirst 'move the recordset object to first row again
		   ReDim resArr(loopCounter,rs.fields.count-1)
		   Do until rs.EOF 'Rows
		   For var=0 to rs.fields.count-1 'cols       ‘                                           ‘ 
		'		If counter=0 Then
		'			resArr(counter,var) =rs.fields(var).name
																													'		else
					resArr(counter,var) =rs.fields(var).value
		'		end if
			Next
			rs.MoveNext
			counter=counter+1
			Loop
			rs.Close
			con.Close
		
			Func_GetQueryResult=resArr

		End If	
	
	
	End Function


	Function toDeletetheReport(ReportName)
		Dim strquery1,strquery2,b,c,Report_Name
		
		' Query to Execute 
		strquery1="Delete  from dbo.Reports  where ReportName='"&Report_Name&"'"
		strquery2="Delete from dbo.Shortcuts  where ShortcutName='"&Report_Name&"'"
		
		' Calling All the Functions
		b=Func_GetQueryResult(strquery1,0)
		
		c=Func_GetQueryResult(strquery2,0)
	
	End Function

    
    '<Assigning Report Permissions to Added User>
    '<strUser : UserName to whom permissions has to be assigned: String>
    '<strPermName : Name of  Permission: String>
    '<strPertype : Type Of Permission>
    '<strPagename : Page name where Report Properties should be accessed>
    '<objData : Reference to objData>
    Public Sub Permission(ByVal strUser, ByVal strPermName, ByVal strPertype, ByVal strPagename, ByRef objData)
                Dim selectedMembersRow, objwelPreviousPage, k, m, j, strUsername, rowCount, i, xPathPermission, xPathPerName, oPermName, xPathPerTypeGrant, xPathPerTypeDeny, objGrant, objDeny, objApply, objbutton, checkBoxValue, objNextPage, strOuterHtml
                checkBoxValue = 0
                
'                'Search for User to add permission to it
'                Browser("Analyzer").Page("Home").Sync
'                Browser("Analyzer").Page("Home").Sync
'                wait 2

'                'Removal of Do While Loop - Start

'Commented by Poornima
'               j=1
'               Do
'               	 If Browser("Hosting Templates - Ops").Page("Analyzer").Frame("Frame").Image("previous_enabled").Exist(10) AND j<>1 Then
'		  	        Browser("Hosting Templates - Ops").Page("Analyzer").Frame("Frame").Image("previous_enabled").Click
'		         End If
'                    strOuterHtml = Browser("Analyzer").Page("Home").Frame("Frame").WebElement("welPreviousPage").GetROProperty("outerhtml")
'                    
'                    If InStr(1, strOuterHtml, "gray") > 0 Then
'                        Exit Do
'                    ElseIf InStr(1, strOuterHtml, "gray") = 0 Then
'                        Set objwelPreviousPage = Browser("Analyzer").Page("Home").Frame("Frame").WebElement("welPreviousPage")
'                        Call SCA.ClickOn(objwelPreviousPage,"welPreviousPage", strPagename)
'                        wait 1
'                        Browser("Analyzer").Page("Home").Sync
'                    End If  
'		         
'                 j=j+1
'			   Loop While Browser("Hosting Templates - Ops").Page("Analyzer").Frame("Frame").Image("previous_enabled").Exist(10)
			   If Browser("Hosting Templates - Ops").Page("Analyzer").Frame("Frame").Image("first_enabled").Exist(10) Then
			   	  Browser("Hosting Templates - Ops").Page("Analyzer").Frame("Frame").Image("first_enabled").Click
			   End If
               

'                Do
'                    Browser("Analyzer").Page("Home").Frame("Frame").WebElement("welPreviousPage").Click
'                    wait 1
'                    Browser("Analyzer").Page("Home").Sync
'                    strOuterHtml = Browser("Analyzer").Page("Home").Frame("Frame").WebElement("welPreviousPage").GetROProperty("outerhtml")
'                Loop While InStr(1, strOuterHtml, "gray") = 0
'                'Removal of Do While Loop - End
                
                Browser("Analyzer").Page("Home").Sync
                
'                'Removal of Do While Loop - Start
	j=1
	Do
	 	If Browser("Hosting Templates - Ops").Page("Analyzer").Frame("Frame").Image("next_enabled").Exist(10) AND j<>1 Then
			Browser("Hosting Templates - Ops").Page("Analyzer").Frame("Frame").Image("next_enabled").Click
		End If
		Browser("Analyzer").Page("Home").Sync
        Browser("Analyzer").Page("Home").RefreshObject
        selectedMembersRow=Browser("Analyzer").Page("Home").Frame("Frame").WebTable("tabSelectedMembers").GetROProperty("rows")
        For j = 1 To selectedMembersRow
            strUsername = Browser("Analyzer").Page("Home").Frame("Frame").WebTable("tabSelectedMembers").GetCellData(j, 2)
            If InStr(1, Trim(UCase(strUsername)),Trim(UCase(strUser))) <> 0 Then
               Browser("Analyzer").Page("Home").Frame("Frame").WebCheckBox("chkCheckUsersAll").Set "ON"
               Browser("Analyzer").Page("Home").Frame("Frame").WebCheckBox("chkCheckUsersAll").Set "OFF"
               Browser("Analyzer").Page("Home").Frame("Frame").WebTable("tabSelectedMembers").ChildItem(j,1,"WebCheckBox",0).click
               checkBoxValue = 1
               Exit For
            End If
       Next
                    
            If checkBoxValue = 1 Then
               Exit Do
            Elseif checkBoxValue = 0 Then
               Set objNextPage = Browser("Analyzer").Page("Home").Frame("Frame").WebElement("welNextPage")
               Call SCA.ClickOn(objNextPage,"welNextPage", strPagename)
               Browser("Analyzer").Page("Home").Sync    
            End If
		j=j+1
     Loop While Browser("Hosting Templates - Ops").Page("Analyzer").Frame("Frame").Image("next_enabled").Exist(10)


'                Do
'                    selectedMembersRow=Browser("Analyzer").Page("Home").Frame("Frame").WebTable("tabSelectedMembers").GetROProperty("rows")
'                    For j = 1 To selectedMembersRow
'                        strUsername = Browser("Analyzer").Page("Home").Frame("Frame").WebTable("tabSelectedMembers").GetCellData(j, 2)
'                        If InStr(1, Trim(UCase(strUsername)),Trim(UCase(strUser))) <> 0 Then
'                            Browser("Analyzer").Page("Home").Frame("Frame").WebCheckBox("chkCheckUsersAll").Set "ON"
'                            Browser("Analyzer").Page("Home").Frame("Frame").WebCheckBox("chkCheckUsersAll").Set "OFF"
'                            Browser("Analyzer").Page("Home").Frame("Frame").WebTable("tabSelectedMembers").ChildItem(j,1,"WebCheckBox",0).click
'                            checkBoxValue = 1
'                            Exit For
'                        End If
'                    Next
'                    
'                    If checkBoxValue = 0 Then
'                        Set objNextPage = Browser("Analyzer").Page("Home").Frame("Frame").WebElement("welNextPage")
'                        Call SCA.ClickOn(objNextPage,"welNextPage", strPagename)
'                        Browser("Analyzer").Page("Home").Sync    
'                    End If
'                    
'                Loop While checkBoxValue = 0
'                'Removal of Do While Loop - End            
        
                If checkBoxValue = 0 Then
                        Call ReportStep (StatusTypes.Fail, "Could Not select member "&strMember& " successfully", strPageName)        
                        Exit Sub
                End If
                
                rowCount = Browser("Analyzer").Page("Home").Frame("Frame").WebTable("tabPermission").GetROProperty("rows")
                For i = 3 To rowCount Step 1
                    
                    xPathPermission = "/html/body/form/table/tbody/tr[1]/td/table/tbody/tr[2]/td/div[2]/span/table/tbody/tr[1]/td/table/tbody/tr[3]/td[2]/table/tbody/tr["&i&"]/"
                    
                    'Descriptive programme for strPermName using Xpath of PermissionName
                    xPathPerName = xPathPermission&"td[2]"
                    oPermName = Browser("Analyzer").Page("Home").Frame("Frame").WebElement("xpath:="&xPathPerName).GetROProperty("innerhtml")
                    
                    If Trim(UCase(oPermName)) = Trim(UCase(strPermName)) Then
                
                        Select Case strPertype
                                Case "Grant"
                                    'Descriptive programme for strPermType:= Grant using Xpath
                                    xPathPerTypeGrant = xPathPermission&"td[3]/span/img"
                                    set objGrant = Browser("Analyzer").Page("Home").Frame("Frame").Image("xpath:="&xPathPerTypeGrant)
                                    Call SCA.ClickOn(objGrant,"Grant Image",strPagename)
                                    
                                Case "Deny"
                                    'Descriptive programme for strPermType:= Grant using Xpath
                                    xPathPerTypeDeny = xPathPermission&"td[4]/span/img"
                                    set objDeny = Browser("Analyzer").Page("Home").Frame("Frame").Image("xpath:="&xPathPerTypeDeny)
                                    Call SCA.ClickOn(objDeny,"Deny Image",strPagename)
                        End Select
                        Exit For
                    End If
                Next
                
                set objApply = Browser("Analyzer").Page("Home").Frame("Frame").WebButton("btnApply")
                Call SCA.ClickOn(objApply,"btnApply",strPagename)
                
                'Click on OK button in Config Report Page
                Set objbutton = Browser("Analyzer").Page("Home").Frame("Frame").WebButton("btnFolderOk")
                Call SCA.ClickOn(objbutton,"btnFolderOk", strPagename)
                
                Browser("Analyzer").Page("Home").Sync
                
    End Sub
    
    Public Sub popUp_Measure_Dimension(ByVal strSelection,ByVal strOperationSelection)
	
		Dim objMeasure,objDimension,objClickWebElements
		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").RefreshObject
		
		   Select Case strSelection
		    Case "Row"
		    
				Set objMeasure=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("AdvanceFilter_down_normal")
				Call SCA.ClickOn(objMeasure,"Measure Popup", "Report Creation Page ")			
				wait 2
			Case "Column"
			
				Set objDimension=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("xpath:=.//*[@id='columnsDrop']/span/img")
				Call SCA.ClickOn(objDimension,"Dimension Popup", "Report Creation Page ")            
				
		   End Select

		   
		Set objClickWebElements=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strOperationSelection)
		Call SCA.ClickOn(objClickWebElements,"Selection","Report Creation Page")
    End Sub
    
    '< To Select the Menu Items of the Tool Bar>
    '<strPage: Page name where QuickChart is created:  String>
    '<objData: Reference to objData>
    'Commented Params: ByVal sheetName, ByVal strQuickChartOper, 
    Public Sub QuickChart(ByVal strPage, ByRef objData)

        Dim quickChartText, dragx,  dragy, quickChartNewText, desWebArea
        Dim  strreturnval,obj_tree1,StrXY,objidentity,objdropdown,obj,k,x2,y2,objdropdownloop, objchartcnt, outertext,validate, i, j, objImgQuickChartMeasure
        validate = 0
        
        Browser("Analyzer").Page("ReportCreation").Sync
        Browser("Analyzer").Page("ReportCreation").RefreshObject
        
        'Validate whether measure has been droppped into QuickChart
        wait 10
        quickChartNewText = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabQuickChart").GetROProperty("innertext")
        'msgbox  quickChartNewText 
		If Browser("micclass:=browser").Page("micclass:=page").Frame("html id:=tabPages_frame2").WebElement("html id:=qc_QuickChart1_container").Exist(20) and InStr(1, quickChartNewText, "Drop a Filter Condition Here")> 1 Then
			Call ReportStep (StatusTypes.Pass, "Successfully added measure to quick chart created", strPage)
			'msgbox "pass"
		Else
			'msgbox "fail"
			Call ReportStep (StatusTypes.Fail, "Could not add measure to quick chart successfully", strPage)
		End If

'        If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgRenderBinary").Exist(15) and InStr(1, quickChartNewText, "Drop a Filter Condition Here")> 1 then
'            Call ReportStep (StatusTypes.Pass, "Successfully added measure to quick chart created", strPage)
'        else
'            Call ReportStep (StatusTypes.Fail, "Could not add measure to quick chart successfully", strPage)
'        End if
        
        For i = 0 To 10 Step 1
            Browser("Analyzer").Page("ReportCreation").Sync
            Browser("Analyzer").Page("ReportCreation").RefreshObject
            If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebArea("imgQuickChartMeasure").Exist(2) Then
                Set objImgQuickChartMeasure = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebArea("imgQuickChartMeasure")
                Setting.WebPackage("ReplayType") = 2
                objImgQuickChartMeasure.RightClick
        
                Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebArea("imgQuickChartMeasure").FireEvent "onmouseover"
                Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebArea("imgQuickChartMeasure").FireEvent "ondblclick"
                wait 2
                If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabQuickChartToolTip").Exist(2) Then
                    outertext=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabQuickChartToolTip").GetROProperty("outertext")
                    validate = 1
                    Exit For
                End If
            End If
        Next
        
        If validate = 1 Then
            Call ReportStep (StatusTypes.Pass,"Tool tip value of measure added to quick chart is " &outertext, strPage)
        End If
'        'Shweta - Start        
'        'Drag and drop the first Measure for the chart from the schema tree
'        dragx=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webAvg Daily Posology").GetROProperty("x")
'        dragy=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webAvg Daily Posology").GetROProperty("y")
'        Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webAvg Daily Posology").Drag dragx,dragy
'        Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabQuickChart").Drop
'        wait 2
'        Browser("Analyzer").Page("ReportCreation").Sync
'        Browser("Analyzer").Page("ReportCreation").Sync
'        'Shweta - End
        
'        'Shweta - <6/4/2015 - Code modified because of change in application objects. Object Identification problem> - Start
'        set desWebArea=Description.Create()
'        desWebArea("micclass").value="WebArea"
'        desWebArea("class").value=""
'        desWebArea("html tag").value="AREA"
'        desWebArea("image type").value="Client Side ImageMap"
'        desWebArea("map name").regularexpression=True
'        desWebArea("map name").value="QuickChart.*ImageMap"
'        
'        set objchartcnt=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabQuickChart").ChildObjects(desWebArea)
'        'msgbox objchartcnt.count
''        
'        For i = 0 To objchartcnt.count-1 Step 1
'            objchartcnt(0).FireEvent "onfocus"
'            objchartcnt(0).FireEvent "onmouseover"
'            wait 1
'            For j = 0 To 10 Step 1
'                If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabQuickChartToolTip").Exist(2) Then
'                    outertext=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabQuickChartToolTip").GetROProperty("outertext")
'                    validate = 1
'                    Exit For
'                End If
'            Next
'            
'            If validate = 1 Then
'                If InStr(1, Trim(Ucase(outertext)), Trim(Ucase("value"))) > 0 and InStr(1, Trim(Ucase(outertext)), Trim(Ucase("legend"))) > 0 and InStr(1, Trim(Ucase(outertext)), Trim(Ucase("category"))) > 0 Then
'                    Call ReportStep (StatusTypes.Pass,"Tool tip value of measure added to quick chart is " &outertext, strPage)
'                    Exit For    
'                End If
'            End If
'            
'        Next
'        'Shweta - <6/4/2015 - Code modified because of change in application objects. Object Identification problem> - End
            
    End Sub
'=========================================================================================================================================================================
'Phase 3 functions
'=========================================================================================================================================================
'Function Name: OPSLogin
'Description : To login to OPS application
'Parameters: strUserName-> valid username to login to application, StrPwd-> valid password to log into application
'Creation Date: 08th July 2015
'Author : IMS Health

'========================================================================================================================================================= 
	
Public Sub OPSLogin(ByVal strUserName,ByVal StrPwd)
	Systemutil.CloseProcessByName "iexplore.exe"
  
   Dim objUserName,objPwd,objLoginBtn,objSCAHomePageImage,returnval,objSCAHomeFrame
	'objUtils.KillProcess("iexplore.exe")

' # CH -001 - Shwetha - Start --
	Systemutil.Run Environment.Value("OPSURL")

	If Browser("Hosting Templates - Ops").Exist(200)  Then
	   If  Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebEdit("txtusername").Exist(60) Then 
	   		Call ReportStep (StatusTypes.Pass,"OPS lanuched successfully","Home Page" )
       Else
			Call ReportStep (StatusTypes.Fail,"OPS not lanuched successfully","Home Page" )
            Exitrun 
	   End if 
	End If
					
	  Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
	  If Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebEdit("txtusername").Exist(4) Then
	  	
	    Set objUserName=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebEdit("txtusername")
		Call SCA.SetText(objUserName,strUserName,"textUserName" ,"Home Page" )
		
		Set objPwd=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebEdit("txtpassword")
		Call SCA.SetText(objPwd,StrPwd,"txtpassword" ,"Home Page" )	
		
		Set objLoginBtn=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebButton("btnLogin")
		Call SCA.ClickOn(objLoginBtn,"Login Button","Home Page")   
	  
	  End If 
	' # CH -001 - Shwetha - End  --	

	Set objSCAHomePageImage=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebElement("welHostingTemplates")
   wait 3.5
   returnval=IMSSCA.Validations.ExistanceWebelements_Verification(objSCAHomePageImage, "", "Home Page", 0)
   
   If returnval=True Then
   	Call ReportStep (StatusTypes.Pass,"Login is Successfull for OPS ","Home Page" )
	Else
	Call ReportStep (StatusTypes.Fail,"Login is notSuccessfull for OPS" ,"Home Page" )
   End If
    	
  End Sub
  
  '=========================================================================================================================================================================
'Phase 3 functions
'=========================================================================================================================================================
'Function Name: OPSLogin
'Description : To login to OPS application
'Parameters: strUserName-> valid username to login to application, StrPwd-> valid password to log into application
'Creation Date: 08th July 2015
'Author : shobha

'========================================================================================================================================================= 
    
Public Sub OPSLogin_Internal()
  
   Dim objUserName,objPwd,objLoginBtn,objSCAHomePageImage,returnval,objSCAHomeFrame
    'objUtils.KillProcess("iexplore.exe")
    Systemutil.CloseProcessByName "iexplore.exe"
    
    '<shweta - 1/6/2016> - Used SystemUtil.run to launch InternalOPSURL - Start
    'objUtils.fnLaunchingIE( Environment.Value("InternalOPSURL") )
    Systemutil.Run Environment.Value("InternalOPSURL")
    '<shweta - 1/6/2016> - Used SystemUtil.run to launch InternalOPSURL - End
    wait 2
    Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
   
   Set objSCAHomePageImage=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebElement("welHostingTemplates")
   wait 3.5
   returnval=IMSSCA.Validations.ExistanceWebelements_Verification(objSCAHomePageImage, "", "Home Page", 0)
   
   If returnval=True Then
       Call ReportStep (StatusTypes.Pass,"Login is Successfull for OPS ","Home Page" )
    Else
    Call ReportStep (StatusTypes.Fail,"Login is notSuccessfull for OPS" ,"Home Page" )
   End If
        
  End Sub


  
  
  
  
  
  
  
  
  
  
  
  
'=========================================================================================================================================================
'Function Name: OPSLogout
'Description : To logout from OPS application
'Parameters: NIL
'Creation Date: 22th July 2015
'Author : Shweta Nagaral

'========================================================================================================================================================= 
	
Public Sub OPSLogout()
  
   	Dim objLogOut, i
	
	Set objLogOut=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Link("lnkLog out")
	'Call SCA.ClickOn(objLogOut,"lnkLog out" , "Ops Page")
	
	'shweta 27/4/2016 - Start
	If Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Link("lnkLog out").Exist(60) Then
		Call SCA.ClickOn(objLogOut,"lnkLog out" , "Ops Page")
	End if
	'shweta 27/4/2016 - End
	
	
	For i = 1 To 20 Step 1
		If Browser("Hosting Templates - Ops").Page("SignOut - Ops System").WebElement("welSignedOut").Exist(1) Then
			Call ReportStep (StatusTypes.Pass,"Logged out from OPS Successfully","OPS Page")
			Exit For
		End If
	
		If i = 20 Then
			Call ReportStep (StatusTypes.Warning,"Could not log out from OPS Successfully","OPS Page")
		End If
	Next
		
	'objUtils.KillProcess("iexplore.exe")
	Systemutil.CloseProcessByName "iexplore.exe"
	
End Sub

'=========================================================================================================================================================
'Function Name: menuitemselectOldFunction
'Description : To select child menu item from the menu
'Parameters: appobject-> req menu object where it has been displayed, menuitem-> required menuitem
'Creation Date: 08th July 2015
'Author : IMS Health

'========================================================================================================================================================= 

  Sub menuitemselectOldFunction(ByVal appobject,ByVal menuitem)
  	itemClick = 0
  	Set menuChild=Description.Create()
	menuChild("micclass").value="Link"
	menuChild("Class").Regularexpression=True
	menuChild("Class").value=".*childMenuItem.*"
	set childcnt=appobject.ChildObjects(menuChild)
	
	For i = 0 To childcnt.count-1
		if childcnt(i).getroproperty("outertext")=menuitem then
			'shweta - start
			
			'For m = 1 To 5 Step 1
			'	'Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Link("Hosting").Click
			'	childcnt(i).WaitProperty "visible",True,10000
			'	If childcnt(i).Exist(2) Then
			'		childcnt(i).click
			'		itemClick = 1
			'		Exit for
			'	End If
			'Next
			'
			'If itemClick = 1 Then
			'	Exit For
			'End If

			'shweta 22/3/2016 Added for intCounterStart/intCounterMaxLimit counter - Start
			For m = Environment.Value("intCounterStart") To Environment.Value("intCounterMaxLimit") Step 1
				'Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Link("Hosting").Click
				childcnt(i).WaitProperty "Visible",True,10000
				If childcnt(i).Exist(2) Then
					childcnt(i).click
					itemClick = 1
					Call ReportStep (StatusTypes.Pass,"Sucessfully re-directed to requested menuitem "&menuitem&" Page from OPS Home Page","OPS Home Page")
					Exit for
				End If
			Next
			
			If itemClick = 1 Then
				Exit For
			End If
			'shweta 22/3/2016 Added for intCounterStart/intCounterMaxLimit counter - End
	
			'shweta - End
		End If
	Next
	
	'shweta 22/3/2016 - Added Report fail stmt if could not re-direct to menuitem Page - start
	If itemClick = 0 Then
		Call ReportStep (StatusTypes.Fail,"Could not re-direct to requested menuitem "&menuitem&" Page from OPS Home Page","OPS Home Page")
	End If
	'shweta 22/3/2016 - Added Report fail stmt if could not re-direct to menuitem Page - End
	
	'shweta 22/3/2016 - Added ReadyState method to check if browser loading is completed  - start
	intLoopCounter = Environment.Value("intCounterStart")
	Do
		wait (1)
		If intLoopCounter=Environment.Value("intCounterMaxLimit") Then
	             Call ReportStep (StatusTypes.Fail,"Browser loading is not completed menuitem "&menuitem&" Page from OPS Home Page","OPS Home Page")
	             Exit Do 
	    End If
		appobject.sync
		intLoopCounter=intLoopCounter+1
		'Call ReportStep (StatusTypes.Information, appobject&" Page from OPS loaded successfully","OPS Home Page")
	Loop While (( appobject.Object.ReadyState) <>"complete")
	'shweta 22/3/2016 - Added ReadyState method to check if browser loading is completed  - End
	
	appobject.sync
	appobject.RefreshObject
  End Sub
  
  '=====================================================================================================================
'Function Name: Creation of client in ops
'Description : To create a client in Ops system
'Parameters: strOpsClient_Name-> Client name , strOpsShort_Name-> Shart name for the client, strOpsDescription-> Brief description ofr the client
'Client ID-> Unique Client Id ( Use formulas in test data to pass Unique Client id ), Checkbox-> to check crate iam role check box needd to pass check if want to check
'( you can pass it data from data table or else you can hardcode)
'Creation Date: 08th July 2015
'Author : IMS Health

'=====================================================================================================================
	Public Sub ClientCreation_Ops(ByVal strOpsClient_Name,ByVal strOpsShort_Name,ByVal strOpsDescription,ByVal ClientID,ByVal Checkbox)

	strscriptstatus="Fail"
	Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
	Set appobject=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops")	'
	'Call IMSSCA.General.menuitemselect(appobject,"Clients")
	
	'<shweta 14/3/2016> - start
	'wait  6
	''Browser("Clients - Ops System").Page("Clients - Ops System").Link("Clients").Click
	'Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Link("lnkHosting Templates").Click
	'Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Link("Clients").Click

'	clientClick = 0
'	Set WshShell =CreateObject("WScript.Shell")
'	set appClient = Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Link("Clients")
'
'	For i = 1 To 5 Step 1
'		WshShell.SendKeys "{F5}"
'		wait 2
'		WshShell.SendKeys "{F5}"
'		Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
'		Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").RefreshObject
'		
'		Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Link("lnkHosting Templates").Click
'		appClient.WaitProperty "visible", true, 10000
'		If appClient.Exist(2) = true Then
'			appClient.Click
'			Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
'			Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").RefreshObject
'			clientClick = 1
'			Exit for
'		End If
'	Next
'	
'	If clientClick = 0 Then
'		Call ReportStep (StatusTypes.Fail,"Could not click on Client in the Ops System. Further steps in Test Case may fail", "OPS Hosting Page")
'	End If
	
		Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").RefreshObject
		Set appobject=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops")
		Set objClientName=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtClientName")
		Set WshShell =CreateObject("WScript.Shell")
		
		For a = 1 To 5 Step 1
			WshShell.SendKeys "{F5}"
			wait 2
			WshShell.SendKeys "{F5}"
			Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
			Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").RefreshObject
			
			Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Link("Hosting").Click
			
			Call IMSSCA.General.menuitemselect(appobject,"Clients")
			
			Browser("Hosting Templates - Ops").Page("Clients - Ops System").Sync
			Browser("Hosting Templates - Ops").Page("Clients - Ops System").RefreshObject
			
			objClientName.WaitProperty "Visible",True,10000
			
			If objClientName.Exist(5) Then
				objClientName.WaitProperty "Visible",True,10000
				Call ReportStep(StatusTypes.Pass, "Successfully redirected to 'Clients' Page in OPS", "OPS Client Page")
				Exit For
			End If
			
			If a=5 Then
				Call ReportStep(StatusTypes.Fail, "Could not successfully redirected to 'Clients' Page in OPS", "OPS Client Page")
				Call ReportStep(StatusTypes.Warning, "Could not successfully redirected to 'Clients' Page in OPS. Further functionality Validation steps may fail", "OPS Client Page")
			End If
		Next
	
	'<shweta 14/3/2016> - End

	Browser("Hosting Templates - Ops").Page("Clients - Ops System").Sync
	Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("BtnAdd New").Click ' 
    ' Entering All the Values to create the Clients 
	wait 5    
	 If Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("webNew Client").Exist(20)  then
	 
	 	Set objtxtName=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtName")
	 	Call SCA.SetText(objtxtName,strOpsClient_Name, "txtName","NewClient")	 
	 
	 	Set objshname=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtShortName")
	 	Call SCA.SetText(objshname,strOpsShort_Name, "txtShortName","NewClient")
	 	
	    Set objdes=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtDescription")
	    Call SCA.SetText(objdes,strOpsDescription, "txtDescription","NewClient")
	    
	    Set objcomid=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtCompanyIds")
	    Call SCA.SetText_MultipleLineArea(objcomid,ClientID,"txtCompanyIds","NewClient")
	    If Checkbox="Check" Then
		    Set objchkiamclient=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebCheckBox("chkCreateIAMClient")
		    Call SCA.SetCheckBox(objchkiamclient,"chkCreateIAMClient","ON","NewClient")
		    Else
		    Set objchkiamclient=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebCheckBox("chkCreateIAMClient")
		    Call SCA.SetCheckBox(objchkiamclient,"chkCreateIAMClient","OFF","NewClient")		    
	    End If  
	    Set objsave=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnSave")
	    Call SCA.ClickOn(objsave,"btnSave","NewClient")
	 End if

	'   Checking the Existance of the Successful message,  if exist  passing the script    
	If Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("webOPS Client created. DC").exist(20)  then
        strscriptstatus="pass" 
        Set objClose=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnClose")
        Call SCA.ClickOn(objClose,"btnClose","NewClient")
         If strscriptstatus="pass"  Then
       	 	Call ReportStep (StatusTypes.Pass,"Add Client in the Ops System", strOpsClient_Name&" Added successfully")
  		 	else
    	 	Call ReportStep (StatusTypes.Fail,"Add Client in the Ops System", strOpsClient_Name&" not Added successfully")
 		End If
     End if 

		
	End Sub


'=========================================================================================================================================================
'Function Name: selectuserprofile
'Description : To select required user profile and to do required operation based on the scenario
'Parameters: UserName-> required username, Operation-> required opearation to perform on user profile, toSelect->which one to select
'Creation Date: 16th July 2015
'Author : IMS Health

'========================================================================================================================================================= 

  Sub selectuserprofile(UserName,Operation,toSelect)
  
    Set objAdminLink=Browser("Analyzer").Page("Home").Frame("Frame").Link("lnkSystem Administration")
	Call SCA.ClickOn(objAdminLink,"lnkSystem Administration","Home Page")
	Browser("Analyzer").Page("Home").Frame("Analyzer AdminCenter").Link("lnkManage User Profiles").Click
	
	Browser("Analyzer").Page("Home").Frame("UserProfile").WebEdit("txtFilter").Set UserName
	Browser("Analyzer").Page("Home").Frame("UserProfile").WebCheckBox("chkFiltered").Set "ON"
	wait 8
	introwcount=Browser("Analyzer").Page("Home").Frame("UserProfile").WebTable("grdUserProfilesdata").GetROProperty("rows")
	Environment("strUserProfile")="NotFound"
	If introwcount>1 Then

		Select Case Operation 
			Case "Delete"
				For i=1 to introwcount
					If Ucase(Browser("Analyzer").Page("Home").Frame("UserProfile").WebTable("grdUserProfilesdata").GetCellData(i,2))=Ucase(UserName) Then
					Environment("strUserProfile")="Found"
		   			Browser("Analyzer").Page("Home").Frame("UserProfile").WebTable("grdUserProfilesdata").ChildItem(i,1,"WebCheckBox",0).click
		   			Exit For
					End If
				Next
				wait 2				
				Browser("Analyzer").Page("Home").Frame("UserProfile").Image("delete").Click
				wait 2
				Browser("Analyzer").Page("Home").Frame("Report").WebButton("Delete User Profile").Click
				wait 4
			Case "Search"
				For i=1 to introwcount
					If Ucase(Browser("Analyzer").Page("Home").Frame("UserProfile").WebTable("grdUserProfilesdata").GetCellData(i,2))=Ucase(UserName) Then
					Environment("strUserProfile")="Found"
					Environment("stractCompany")=Browser("Analyzer").Page("Home").Frame("UserProfile").WebTable("grdUserProfilesdata").GetCellData(i,4)
					Select Case toSelect
						Case "User"
							Browser("Analyzer").Page("Home").Frame("UserProfile").WebTable("grdUserProfilesdata").ChildItem(i,2,"WebElement",0).Click
							wait 2
					End Select
		   			Exit For
					End If
				Next
			
			End Select
	End If
	End Sub
'============================================================================================================================'=============================
'Function Name: Addordeluserstoaopsclientoldfunction1
'Description : To add or delete users to a specific client
'Parameters: 
'clientname-> Name of the client,reqAction-> Action to perform on the clint,OPSoperation->Operation to perform on the clint, objData-> This is constant as need to use test data inside the function
'CountryName-> Name of the country,Offering->Offering name ,User->User name,Role->Role name
'Creation Date: 16th July 2015
'Author : IMS Health
'Last Modified Date:NA
'Last Modifyed by :NA

'============================================================================================================================'============================= 
	Sub Addordeluserstoaopsclientoldfunction1(clientname,reqAction,OPSoperation,ByRef objData,CountryName,Offering,User,Role)
		
		'shweta <30/3/2016> Added validation point to checking Server Response Time / Error captured while creating a new client - start
	    'wait 300
	    'shweta <30/3/2016> Added validation point to checking Server Response Time / Error captured while creating a new client - End
	   
		Set WshShell =CreateObject("WScript.Shell")
		WshShell.SendKeys "{F5}"
		wait 2
		WshShell.SendKeys "{F5}"
		Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
		Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").RefreshObject
		Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
		
		'shweta05/1/2016. Added waitproperty and sync stmts - Start
		Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").RefreshObject
		Set appobject=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops")
		Set objClientName=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtClientName")
		
		For a = 1 To 5 Step 1
			WshShell.SendKeys "{F5}"
			wait 2
			WshShell.SendKeys "{F5}"
			Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
			Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").RefreshObject
			
			Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Link("Hosting").Click
			
			Call IMSSCA.General.menuitemselect(appobject,"Clients")
			Browser("Hosting Templates - Ops").Page("Clients - Ops System").Sync
			Browser("Hosting Templates - Ops").Page("Clients - Ops System").RefreshObject
			
			objClientName.WaitProperty "visible",True,10000
			If objClientName.Exist(1) Then
				Call ReportStep(StatusTypes.Pass, "Successfully redirected to 'Clients' Page in OPS", "OPS Client Page")
				Exit For
			End If
			
			If a=5 Then
				Call ReportStep(StatusTypes.Fail, "Could not successfully redirected to 'Clients' Page in OPS", "OPS Client Page")
				Call ReportStep(StatusTypes.Warning, "Could not successfully redirected to 'Clients' Page in OPS. Further functionality Validation steps may fail", "OPS Client Page")
			End If
		Next
		
 		'Set objClientName=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtClientName")
		'shweta05/1/2016. Added waitproperty and sync stmts - End
		
		Call SCA.SetText(objclientname,clientname,"txtClientName","Clients - Ops System")
	    set WshShell = CreateObject("WScript.Shell")
		Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtClientName").MiddleClick
		WshShell.SendKeys "{TAB}"
		WshShell.SendKeys "{ENTER}"
		wait 5
		introwcount=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("tablClientList").GetROProperty("rows")
		If introwcount>1 Then
			For i=1 to introwcount
				If ucase(Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("tablClientList").GetCellData(i,1))=ucase(clientname) Then
'				   'shweta <30/3/2016> Added validation point to checking Server Response Time / Error captured while creating a new client - start
'				   wait 600
'				   'shweta <30/3/2016> Added validation point to checking Server Response Time / Error captured while creating a new client - End
				   intwebelmcnt=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("tablClientList").ChildItemCount(i,4,"WebElement")
				   For j=0 to intwebelmcnt-1
					   If Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("tablClientList").ChildItem(i,4,"WebElement",j).getroproperty("outertext")=reqAction Then
					   	  wait 2
					   	  Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("tablClientList").ChildItem(i,4,"WebElement",j).Click
					   	  wait 5
					   	  Exit For
					   End If
				   Next
	   		      Exit For 
				End If
			Next
		
			Select Case reqAction
				Case "Manage IAM Users"
					Select Case OPSoperation
						Case "Add"
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebList("Client").Select clientname
							wait 2
							Set objgetroles=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnGet Roles")
							Call SCA.ClickOn(objgetroles,"Get Roles","Add User")
							wait 2
							arrroles=Split(Role,";")
							
							For i = 0 To ubound(arrroles)
								Set rolecheck=Description.Create()
								rolecheck("MicClass").value="WebCheckBox"
								rolecheck("name").value=arrroles(i)
								Set chrolecheck=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(rolecheck)
								chrolecheck(0).Set "ON"
							Next
							
							wait 2
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtUsers").Set User
		'					Set objuser=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtUsers")
		'					Call SCA.SetText(objuser,objData.Item("User"),"Users","Add User")
							'wait 2
							Set objadd=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Add")
							Call SCA.ClickOn(objadd,"Add","Add User")
							Wait 10
						Case "Addusertoroal"
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebList("Client").Select clientname
							wait 2
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtRoles").Set Role
	
							Set objgetroles=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnGet Roles")
							Call SCA.ClickOn(objgetroles,"Get Roles","Add User")
							wait 2
							arrroles=Split(Role,";")
							
							For i = 0 To ubound(arrroles)
								Set rolecheck=Description.Create()
								rolecheck("MicClass").value="WebCheckBox"
								rolecheck("name").value=arrroles(i)
								Set chrolecheck=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(rolecheck)
								chrolecheck(0).Set "ON"
							Next
							
							wait 2
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtUsers").Set User
		'					Set objuser=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtUsers")
		'					Call SCA.SetText(objuser,objData.Item("User"),"Users","Add User")
							'wait 2
							Set objadd=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Add")
							Call SCA.ClickOn(objadd,"Add","Add User")
							Wait 10
							
						Case "Delusertoroal"
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebList("Client").Select clientname
							wait 2
							'<shweta - 7/12/2015> - Start
							offerpermssions=Split(Role,";")
							For i = 0 To UBound(offerpermssions) Step 1
								Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtRoles").Set offerpermssions(i)
								If i = 0 Then
									Set objgetroles=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnGet Roles")
									Call SCA.ClickOn(objgetroles,"Get Roles","Add User")
									wait 2
								End If
								
								Set rolecheck=Description.Create()
								rolecheck("MicClass").value="WebCheckBox"
								rolecheck("name").value=offerpermssions(i)
								Set chrolecheck=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(rolecheck)
								chrolecheck(0).Set "ON"
								
							Next
							'<shweta - 7/12/2015> - End
							
'							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtRoles").Set Role
'	
'							Set objgetroles=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnGet Roles")
'							Call SCA.ClickOn(objgetroles,"Get Roles","Add User")
'							wait 2
'							Set rolecheck=Description.Create()
'							rolecheck("MicClass").value="WebCheckBox"
'							rolecheck("name").value=Role
'							Set chrolecheck=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(rolecheck)
'							chrolecheck(0).Set "ON"
							'Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Select All").Click
							wait 2
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtUsers").Set User
		'					Set objuser=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtUsers")
		'					Call SCA.SetText(objuser,objData.Item("User"),"Users","Add User")
							'wait 2
							Set objDelete=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Delete")
							Call SCA.ClickOn(objDelete,"Delete","Delete User")
							Wait 10
							'Set objDelete=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Delete IAM User")
							'Call SCA.ClickOn(objDelete,"Delete IAM User","Delete User")
							Wait 5
					End Select
				
				Case "Manage IAM Roles"
					Select Case OPSoperation
						Case "CreateOfferingRole"




'==============================================================================================================================

						If clientname="IMS Health" Then
							
						
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("Country").Click
							set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys CountryName
							wait 5
							Set Desc=Description.Create()
							Desc("MicClass").value="WebElement"
							Desc("class").Value="ui-corner-all"
							Set chobj=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(Desc)
							For i=0 to chobj.count-1
								If chobj(i).getroproperty("innertext")=CountryName Then
									chobj(i).Click
								End If 
							Next
											
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("Offering").Click
							WshShell.SendKeys Offering
							wait 5
							Set Desc=Description.Create()
							Desc("MicClass").value="WebElement"
							Desc("class").Value="ui-corner-all"
							Set chobj=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(Desc)
							For i=0 to chobj.count-1
								If chobj(i).getroproperty("innertext")=Offering Then
									chobj(i).Click
								End If 
							Next
						
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnGet Roles").Click
						wait 2
	
						stractDisplayinrolegrid=objData.Item("ShortName")&" "&objData.Item("CountryShortName")&" "&Replace(Offering," ","")
						
						stroffering="notfound"
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RoleId").ChildItem(2,3,"WebEdit",0).Set stractDisplayinrolegrid
						'Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("RoleName").Set stractDisplayinrolegrid
						'set WshShell = CreateObject("WScript.Shell")
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("RoleName").MiddleClick
						WshShell.SendKeys "{TAB}"
						WshShell.SendKeys "{ENTER}"
						wait 4
						intofferscnt=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").RowCount
						For i=1 to intofferscnt
						 If Ucase(Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").GetCellData(i,3))=Ucase(stractDisplayinrolegrid) Then
						 	stroffering="found"
						 	Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").ChildItem(i,1,"Webcheckbox",0).Set "ON"
						 	Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnDelete Selected Roles").Click
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Ok").Click
							wait 10
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").WaitProperty "Exist",True,30
							strroledeletemsg=" "
							If Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").Exist Then
								strroledeletemsg=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").GetROProperty("innertext")
								If instr(1,Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").GetROProperty("innertext"),"deleted successfully")<>0 Then
									Call ReportStep(StatusTypes.Pass,"After deleting offering Folder "&Offering&" success message should display ","After deleting offering Folder "&Offering&" success message has been displayed ")
								End If
							End If														
							
						 	Exit For
						 End If
						Next
					Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnClose").Click	
					Browser("Hosting Templates - Ops").Page("Clients - Ops System").RefreshObject
						
					intwebelmcnt=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("tablClientList").ChildItemCount(2,4,"WebElement")
				   For j=0 to intwebelmcnt-1
					   If Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("tablClientList").ChildItem(2,4,"WebElement",j).getroproperty("outertext")=reqAction Then
					   	  wait 2
					   	  Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("tablClientList").ChildItem(2,4,"WebElement",j).Click
					   	  wait 5
					   	  Exit For
					   End If
				   Next
               End If
'==============================================================================================================================
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("Country").Click
							set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys CountryName
							wait 2
							Set Desc=Description.Create()
							Desc("MicClass").value="WebElement"
							Desc("class").Value="ui-corner-all"
							Set chobj=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(Desc)
							For i=0 to chobj.count-1
								If chobj(i).getroproperty("innertext")=CountryName Then
									chobj(i).Click
								End If 
							Next
											
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("Offering").Click
							WshShell.SendKeys Offering
							wait 2
							Set Desc=Description.Create()
							Desc("MicClass").value="WebElement"
							Desc("class").Value="ui-corner-all"
							Set chobj=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(Desc)
							For i=0 to chobj.count-1
								If chobj(i).getroproperty("innertext")=Offering Then
									chobj(i).Click
								End If 
							Next
			
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Create Role").Click
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Ok").Click
							wait 10
							Environment("strresmessage")=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").GetROProperty("innertext")
	
						Case "DeleteOfferingRole"
						
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("Country").Click
							set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys CountryName
							wait 5
							Set Desc=Description.Create()
							Desc("MicClass").value="WebElement"
							Desc("class").Value="ui-corner-all"
							Set chobj=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(Desc)
							For i=0 to chobj.count-1
								If chobj(i).getroproperty("innertext")=CountryName Then
									chobj(i).Click
								End If 
							Next
											
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("Offering").Click
							WshShell.SendKeys Offering
							wait 5
							Set Desc=Description.Create()
							Desc("MicClass").value="WebElement"
							Desc("class").Value="ui-corner-all"
							Set chobj=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(Desc)
							For i=0 to chobj.count-1
								If chobj(i).getroproperty("innertext")=Offering Then
									chobj(i).Click
								End If 
							Next
						
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnGet Roles").Click
						wait 2
	
						stractDisplayinrolegrid=objData.Item("ShortName")&" "&objData.Item("CountryShortName")&" "&Replace(Offering," ","")
						
						stroffering="notfound"
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RoleId").ChildItem(2,3,"WebEdit",0).Set stractDisplayinrolegrid
						'Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("RoleName").Set stractDisplayinrolegrid
						'set WshShell = CreateObject("WScript.Shell")
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("RoleName").MiddleClick
						WshShell.SendKeys "{TAB}"
						WshShell.SendKeys "{ENTER}"
						wait 4
						intofferscnt=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").RowCount
						For i=1 to intofferscnt
						 If Ucase(Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").GetCellData(i,3))=Ucase(stractDisplayinrolegrid) Then
						 	stroffering="found"
						 	Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").ChildItem(i,1,"Webcheckbox",0).Set "ON"
						 	Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnDelete Selected Roles").Click
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Ok").Click
							wait 10
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").WaitProperty "Exist",True,30
							strroledeletemsg=" "
							If Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").Exist Then
								strroledeletemsg=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").GetROProperty("innertext")
								If instr(1,Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").GetROProperty("innertext"),"deleted successfully")<>0 Then
									Call ReportStep(StatusTypes.Pass,"After deleting offering Folder "&Offering&" success message should display ","After deleting offering Folder "&Offering&" success message has been displayed ")
								End If
							End If														
							
						 	Exit For
						 End If
						Next
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnGet Roles").Click
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RoleId").ChildItem(2,3,"WebEdit",0).Set stractDisplayinrolegrid
						'Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("RoleName").Set stractDisplayinrolegrid
						'set WshShell = CreateObject("WScript.Shell")
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("RoleName").MiddleClick
						WshShell.SendKeys "{TAB}"
						WshShell.SendKeys "{ENTER}"
						wait 4
						intofferscnt=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").RowCount
						strdelrole="deleted"
						For l=2 to  intofferscnt
						If Ucase(Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").GetCellData(l,3))=Ucase(stractDisplayinrolegrid)	Then
							strdelrole="notdeleted"
							Exit For
						End If
						Next
						If strdelrole="deleted" Then
							Call ReportStep (StatusTypes.Pass,"Deleted role offer: "&stractDisplayinrolegrid&" should not display in the result grid", "Deleted role offer: "&stractDisplayinrolegrid&" is not displayed in the result grid")
	  		 				Else
	    	 				Call ReportStep (StatusTypes.Fail,"Deleted role offer: "&stractDisplayinrolegrid&" should not be display in the result grid", "Deleted role offer: "&stractDisplayinrolegrid&" is displayed in the result grid")
						End If
						If instr(1,strroledeletemsg,"deleted successfully")=0 Then
						
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnClose").Click
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").Link("Event Log & Subscription").Click
							Set appobject=Browser("Hosting Templates - Ops").Page("Clients - Ops System")
							Call menuitemselect(appobject,"Hosting Event Log")
							Set objclientname=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtClientName")
							Call SCA.SetText(objclientname,clientname,"txtClientName","Clients - Ops System")
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtClientName").MiddleClick
							WshShell.SendKeys "{TAB}"
							WshShell.SendKeys "{ENTER}"
							wait 5
							introwcount=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("Template/Client Id").GetROProperty("rows")
							If introwcount>1 Then
								For i=1 to introwcount
									If ucase(Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("Template/Client Id").GetCellData(i,1))=ucase(clientname) Then
									   strEventlogmsg=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("Template/Client Id").GetCellData(i,4)
									   Exit For
									End If
								Next
								If instr(1,strEventlogmsg,"deleted by")<>0 Then
									Call ReportStep(StatusTypes.Pass,"After deleting offering Folder "&Offering&" success message should display ","After deleting offering Folder "&Offering&" success message has been displayed in event log")
									Else
									Call ReportStep(StatusTypes.Fail,"After deleting offering Folder "&Offering&" success message should display ","After deleting offering Folder "&Offering&" success message has not been displayed in event log as well")					
								End If
							End If
						
						End If
						
						Case "DeleteOfferRoleholdmsg"
						
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("Country").Click
							set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys CountryName
							wait 5
							Set Desc=Description.Create()
							Desc("MicClass").value="WebElement"
							Desc("class").Value="ui-corner-all"
							Set chobj=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(Desc)
							For i=0 to chobj.count-1
								If chobj(i).getroproperty("innertext")=CountryName Then
									chobj(i).Click
								End If 
							Next
											
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("Offering").Click
							WshShell.SendKeys Offering
							wait 5
							Set Desc=Description.Create()
							Desc("MicClass").value="WebElement"
							Desc("class").Value="ui-corner-all"
							Set chobj=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(Desc)
							For i=0 to chobj.count-1
								If chobj(i).getroproperty("innertext")=Offering Then
									chobj(i).Click
								End If 
							Next
						
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnGet Roles").Click
						wait 2
	
						stractDisplayinrolegrid=objData.Item("ShortName")&" "&objData.Item("CountryShortName")&" "&Replace(Offering," ","")
						
						stroffering="notfound"
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RoleId").ChildItem(2,3,"WebEdit",0).Set stractDisplayinrolegrid
						'Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("RoleName").Set stractDisplayinrolegrid
						'set WshShell = CreateObject("WScript.Shell")
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("RoleName").MiddleClick
						WshShell.SendKeys "{TAB}"
						WshShell.SendKeys "{ENTER}"
						wait 10
						intofferscnt=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").RowCount
						For i=1 to intofferscnt
						 If Ucase(Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").GetCellData(i,3))=Ucase(stractDisplayinrolegrid) Then
						 	stroffering="found"
						 	Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").ChildItem(i,1,"Webcheckbox",0).Set "ON"
						 	Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnDelete Selected Roles").Click
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Ok").Click
							wait 10
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").WaitProperty "Exist",True,30
							Environment("strdeletemsg")=" "
							msgexist=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").Exist
							
							If msgexist=True Then
								if instr(1,Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").GetROProperty("innertext"),"error")<>0 then
									Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("PartialMessage").Click
									wait 2
									Environment("strdeletemsg")=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Messageadescription").GetROProperty("innertext")
									Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Ok").Click
								Else
								Environment("strdeletemsg")=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").GetROProperty("innertext")	
								End If
																	
							End If						
							
						 	Exit For
						 End If
						Next
	
						If msgexist=False Then
						
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnClose").Click
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").Link("Event Log & Subscription").Click
							Set appobject=Browser("Hosting Templates - Ops").Page("Clients - Ops System")
							Call menuitemselect(appobject,"Hosting Event Log")
							Set objclientname=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtClientName")
							Call SCA.SetText(objclientname,clientname,"txtClientName","Clients - Ops System")
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtClientName").MiddleClick
							WshShell.SendKeys "{TAB}"
							WshShell.SendKeys "{ENTER}"
							wait 5
							introwcount=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("Template/Client Id").GetROProperty("rows")
							If introwcount>1 Then
								For i=1 to introwcount
									If ucase(Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("Template/Client Id").GetCellData(i,1))=ucase(clientname) Then
									   Environment("strdeletemsg")=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("Template/Client Id").GetCellData(i,4)
									   Exit For
									End If
								Next
	
							End If
						
						End If
						
					End Select
					
				Case "IAM DataSource Permission"
					strDataSource = User
					strRole=Role
							
					Select Case OPSoperation
						Case "Add"
							
							'TODO: Ask Mani to click on "IAM DataSource Permission"
							Set WshShell = CreateObject("WScript.Shell")
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtAddDatasource").Click
					        WshShell.SendKeys strDataSource
					        wait 2 ' To wait for List to display as per keyword entered above
					        WshShell.SendKeys "{DOWN}" 
					        wait 2 'To select item from list and wait
					        WshShell.SendKeys "{ENTER}"
					        Set WshShell = Nothing

							Set WshShell = CreateObject("WScript.Shell")
							Set objtxtAddDataSourceRole=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtAddDataSourceRole")
							Call SCA.SetText(objtxtAddDataSourceRole,strRole,"txtAddDataSourceRole","IAM DataSource Permission")
							wait 2
				            WshShell.SendKeys "{ENTER}"
				            wait 1
				            Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtAddDataSourceRole").Click
							WshShell.SendKeys "{ENTER}"
				            Set WshShell = Nothing
							wait 2
							
							set objchkAddDatasourceWithRoles= Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebCheckBox("chkAddDatasourceWithRoles")
							Call SCA.SetCheckBox(objchkAddDatasourceWithRoles,"chkAddDatasourceWithRoles","ON","IAM DataSource Permission")		    
							
							Set objbtnAddDatasourceWithRoles = Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnAddDatasourceWithRoles")
							Call SCA.ClickOn(objbtnAddDatasourceWithRoles,"btnAddDatasourceWithRoles","IAM DataSource Permission")
							
							'Validating "Are you sure you want to add selected Role(s)?"
							If Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Are you sure you want to add selected Roles").Exist(2) then
								Set objbtnOk = Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Ok")
								Call SCA.ClickOn(objbtnOk,"Ok","IAM DataSource Permission")
								
								Call ReportStep (StatusTypes.Pass,"Successfully Added added Datasoure "&strDataSource&" with Role(s) "&strRole&" mapping","IAM DataSource Permission")
							End If
							
						Case "Remove"
							
							Set WshShell = CreateObject("WScript.Shell")
							Set objtxtRemoveDatasourceRoleName= Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtRemoveDatasourceRoleName")
							Call SCA.SetText(objtxtRemoveDatasourceRoleName,strRole, "txtRemoveDatasourceRoleName","IAM DataSource Permission")
							wait 2
				            WshShell.SendKeys "{ENTER}"
				            wait 1
				            Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtRemoveDatasourceRoleName").Click
							WshShell.SendKeys "{ENTER}"
				            wait 2
							
							Set objtxtRemoveDataSourceName = Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtRemoveDataSourceName")
							Call SCA.SetText(objtxtRemoveDataSourceName,strDataSource, "txtRemoveDataSourceName","IAM DataSource Permission")
							wait 2
				            WshShell.SendKeys "{ENTER}"
				            wait 1
				            Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtRemoveDataSourceName").Click
							WshShell.SendKeys "{ENTER}"
				            Set WshShell = Nothing
							wait 2
							
							rc = Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("tabRemoveDatasourceWithRoles").GetROProperty("rows")
							If rc=2 Then
								set objchkRemoveDatasourceWithRoles= Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebCheckBox("chkRemoveDatasourceWithRoles")
								Call SCA.SetCheckBox(objchkRemoveDatasourceWithRoles,"chkRemoveDatasourceWithRoles","ON","IAM DataSource Permission")		    
							
								set objbtnRemoveDatasourceWithRoles = Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnRemoveDatasourceWithRoles")
								Call SCA.ClickOn(objbtnRemoveDatasourceWithRoles,"btnRemoveDatasourceWithRoles","IAM DataSource Permission")
								
								'Validating "Are you sure you want to delete selected Datasoure with Role(s) mapping?"
								If Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Are you sure you want to delete selected Datasoure").Exist(2) then
									Set objbtnOk = Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Ok")
									Call SCA.ClickOn(objbtnOk,"Ok","IAM DataSource Permission")
									
									Call ReportStep (StatusTypes.Pass,"Successfully deleted selected Datasoure "&strDataSource&" with Role(s) "&strRole&" mapping","IAM DataSource Permission")
								End If
							else
								Call ReportStep (StatusTypes.Fail,"Could not delete selected Datasoure "&strDataSource&" with Role(s) "&strRole&" mapping","IAM DataSource Permission")
							End If
							
					End Select
					
					Set objClose=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnClose")
        			'shweta 27/4/2016 - Start
        			'objClose.Exist(1000)
        			'shweta 27/4/2016 - End
        			Call SCA.ClickOn(objClose,"btnClose","NewClient")
					
					Browser("Hosting Templates - Ops").Page("Clients - Ops System").Sync
					Browser("Hosting Templates - Ops").Page("Clients - Ops System").RefreshObject
			End Select
		End If
	End Sub
	
'=========================================================================================================================================================
'Function Name: selectcompanyprofile
'Description : To select required company profile and to do required operation based on the scenario
'Parameters: CompName-> required CompName, Operation-> required opearation to perform on company profile, toSearch-> To link to select for the comapny
'Creation Date: 16th July 2015
'Author : IMS Health	

'========================================================================================================================================================= 

  Sub selectcompanyprofile(CompName,Operation,toSelect)
  
    Set objAdminLink=Browser("Analyzer").Page("Home").Frame("Frame").Link("lnkSystem Administration")
	Call SCA.ClickOn(objAdminLink,"lnkSystem Administration","Home Page")
	Browser("Analyzer").Page("Home").Frame("Analyzer AdminCenter").Link("Manage Companies").Click
	
	Browser("Analyzer").Page("Home").Frame("ManageCompanies").WebEdit("txtFilter").Set CompName
	Browser("Analyzer").Page("Home").Frame("ManageCompanies").WebCheckBox("chkFiltered").Set "ON"
	wait 4
	introwcount=Browser("Analyzer").Page("Home").Frame("ManageCompanies").WebTable("grdCompanydata").GetROProperty("rows")
	Environment("strCompProfile")="NotFound"
	If introwcount>1 Then

		Select Case Operation 
'			Case "Delete" 'Manikumar: edit and use below code if u want to delete 
'				For i=1 to introwcount
'					If Browser("Analyzer").Page("Home").Frame("ManageCompanies").WbfGrid("grdCompanydata").GetCellData(i,2)=CompName Then
'					Environment("strCompProfile")="Found"
'		   			Browser("Analyzer").Page("Home").Frame("ManageCompanies").WbfGrid("grdCompanydata").ChildItem(i,1,"WebCheckBox",0).click
'		   			Exit For
'					End If
'				Next		
'				Browser("Analyzer").Page("Home").Frame("ManageCompanies").Image("delete").Click
'				Browser("Analyzer").Page("Home").Frame("Report").WebButton("Delete User Profile").Click
			Case "Search"
				For i=1 to introwcount
					If Browser("Analyzer").Page("Home").Frame("ManageCompanies").WebTable("grdCompanydata").GetCellData(i,2)=CompName Then
					Environment("strCompProfile")="Found"
					
					Select Case toSelect
						Case "Manage Roles"
							intactionscnt=Browser("Analyzer").Page("Home").Frame("ManageCompanies").WebTable("grdCompanydata").ChildItemCount(i,5,"WebElement")
							For j=0 to intactionscnt-1
							If Trim(Browser("Analyzer").Page("Home").Frame("ManageCompanies").WebTable("grdCompanydata").ChildItem(i,5,"WebElement",j).getroproperty("innertext"))="Manage Roles" Then
								Browser("Analyzer").Page("Home").Frame("ManageCompanies").WebTable("grdCompanydata").ChildItem(i,5,"WebElement",j).click
								Exit For
							End If
								
							Next
					End Select
		   			Exit For
					End If
				Next
			
			End Select
	End If
	End Sub

'===========================================================================================================================================================
'Function Name: selectofferingrole
'Description : To select required offeringrole and it’s details
'Parameters: Rolename-> Role Name, Selectrole-> Need to click on role or not if need to click on role "Yes", RoleDetails-> Which role detail need to select
'Creation Date: 16th July 2015
'Author : IMS Health


'===========================================================================================================================================================
	Sub selectofferingrole(Rolename,Selectrole,RoleDetails)
	 
		If Browser("Analyzer").Page("Home").Frame("Frame").Link("lnkSystem Administration").Exist Then
			Set objAdminLink=Browser("Analyzer").Page("Home").Frame("Frame").Link("lnkSystem Administration")
			Call SCA.ClickOn(objAdminLink,"lnkSystem Administration","Home Page")	
		End If
	
			
		Set objManagerole=Browser("Analyzer").Page("Home").Frame("Analyzer AdminCenter").Link("Manage Roles")
		Call SCA.ClickOn(objManagerole,"Manage Roles","Admin Page")
		Browser("Analyzer").Page("Home").Frame("UserProfile").WebEdit("txtFilter").Set ""
		Browser("Analyzer").Page("Home").Frame("UserProfile").WebCheckBox("chkFiltered").Set "OFF"
		wait 5
		Browser("Analyzer").Page("Home").Frame("UserProfile").WebEdit("txtFilter").Set Rolename
		Browser("Analyzer").Page("Home").Frame("UserProfile").WebCheckBox("chkFiltered").Set "ON"
		wait 4
		introwcount=Browser("Analyzer").Page("Home").Frame("UserProfile").WebTable("grdRolesdata").GetROProperty("rows")
		Environment("strRoledata")="NotFound"
		If introwcount>1 Then
			For i=1 to introwcount
				If Browser("Analyzer").Page("Home").Frame("UserProfile").WebTable("grdRolesdata").GetCellData(i,2)=Rolename Then
					Environment("strRoledata")="Found"
					If Selectrole="Yes" Then
						intactionscnt=Browser("Analyzer").Page("Home").Frame("UserProfile").WebTable("grdRolesdata").ChildItemCount(i,2,"WebElement")
						For j=0 to intactionscnt-1
							If Trim(Browser("Analyzer").Page("Home").Frame("UserProfile").WebTable("grdRolesdata").ChildItem(i,2,"WebElement",j).getroproperty("innertext"))=Rolename Then
								Browser("Analyzer").Page("Home").Frame("UserProfile").WebTable("grdRolesdata").ChildItem(i,2,"WebElement",j).click
								Exit For
							End If
						Next	
					End If
					
					Select Case RoleDetails
						Case "Permission"
						Set objpermision=Browser("Analyzer").Page("Home").Frame("RolesEditor").WebElement("Permission")
						'Call SCA.ClickOn(objpermision,"Permission","Roles")
						Browser("Analyzer").Page("Home").Frame("RolesEditor").WebElement("Permission").Click
						Wait 2
					End Select
					Exit For
				End If
			Next
				
		
		End If
	End Sub

'===========================================================================================================================================================
  '<To create New Template in the ops System>
    '<strOselection: to select the operation>
    '<strFolderName : Foldername>
    '<    strOName: Name of the Template>
    '< strCountry:Name of the country>    
    '<strF: Frequency>
    '<strOcode: Organization Code>
    '<StrORef: Reference code>
    '<strFLAServer: FLA Server Name>
    '<strHServer:- Host Server Name>
    '<Author>Shobha<Author>
 '===========================================================================================================================================================
    Public Function NewTemplateCreationinops(ByVal strOName,ByVal strCountry,ByVal strF,ByVal strOcode,ByVal strORef,ByVal strFLAServer,ByVal strHServer,ByVal strCAHDProcess,ByVal chkbox,ByVal strF1,ByVal strF2)
    
    Dim objbtnAdd,objtxtName,objlstCountry,objFrequencyId,objtxtOCode,objtxtOrderRef
    Dim objLstFLA,objHServer,tempIdName,objbtnSave,errMsg,errorcnt
    Dim StrdupCubeTemp
    '<looping variables>
    Dim i
    
    Set objbtnAdd=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebButton("btnAdd New")
    Call SCA.ClickOn(objbtnAdd,"AddButton" , "Client Creation Page")
    
    Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync

    set objtxtName=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebEdit("txtName")
    Call SCA.SetText(objtxtName,strOName, "txtName","Client Creation Page" )

    Set objlstCountry=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebList("lstCountryId")
    Call SCA.SelectFromDropdown(objlstCountry,strCountry)    
        
    Set objFrequencyId=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebList("lstFrequencyId")
    Call SCA.SelectFromDropdown(objFrequencyId,strF)

    Set objtxtOCode=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebEdit("txtOrganizationCode")
    Call SCA.SetText(objtxtOCode,strOcode, "txtOrgCode","Client Creation Page" )
    
    Set objtxtOrderRef=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebEdit("txtOrderReference")
    Call SCA.SetText(objtxtOrderRef,strORef, "txtOreference","Client Creation Page" )
    
    If chkbox=0 Then
        
        Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebCheckBox("chkRetainPrior").Set "ON"
        Else
        Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebCheckBox("chkRetainPrior").Set "OFF"
        
    End If

    
    Set objLstFLA=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebList("lstFLAServerId")
    Call SCA.SelectFromDropdown(objLstFLA,strFLAServer)
    
    Set objHServer=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebList("lstHostServerId")
    Call SCA.SelectFromDropdown(objHServer,strHServer)
    wait 2    
    tempIdName = Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebEdit("txtNewTemplateIdName").GetROProperty("value")
    

    Set objbtnSave=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebButton("btnSave")
    Call SCA.ClickOn(objbtnSave,"TemplatebtnSave" , "Template Creation Page")
    
    Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
    

    'Descriptive programming to find the Error message while creating duplicate cube templates
    Set errMsg = Description.Create()
    errMsg("micclass").value = "WebElement"
    errMsg("class").value = "message error"
    errMsg("innertext").Regularexpression=True
    errMsg("innertext").value = "Hosting Template\:.*.already exists"
    set errorcnt=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").ChildObjects(errMsg)


    'Reporter statement for Duplicate Hosting Templates
    If errorcnt.count >= 1 Then
        StrdupCubeTemp = Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebElement("webHosting Template Status").GetROProperty("innertext")        
        Call ReportStep (StatusTypes.Fail, "Creating new cube template:-"&Space(3)&"New cube template cannot be created Because Hosting Template already exists with the same details"&Space(3)&StrdupCubeTemp,"Templete Creation Page")        
        Exit Function
    End If    
    
    wait 4

    For i=1 to 180
    If Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebElement("webNewTempCreationStatus").Exist(15)  Then            
    Exit For            
    End If
    Next
    

    If Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebElement("webNewTempCreationStatus").Exist(15)  Then
            Call ReportStep (StatusTypes.Pass, "Creation of new template:-"&Space(7)&"Template created successfully","Template Creation Page")            
            else            
            Call ReportStep (StatusTypes.Fail, "Creation of new template:-"&Space(7)&"Template Not created successfully","Template Creation Page")
    End If


    Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
    Set objbtnSave=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebButton("btnClose")
    Call SCA.ClickOn(objbtnSave,"btnSave" , "Template Creation Page")    
    Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
    NewTemplateCreationinops=tempIdName
    
End Function
    
'===========================================================================================================================================================        
    '<Creation of Folder Structure in FLA>
    '<strtemplateName :-Template Name to create the Folder Structure>    
    '<author :-Shobha> 
'===========================================================================================================================================================    
    Public Sub CreationofFolderStructure_FLAServer(ByVal strtemplateName,ByVal strabffile)
        
        
        'Source path , from where we have to copy the abf and event file
        SourcePath=Environment.Value("SourceFile")
        'msgbox SourcePath
        
        'Folder path, which is to be created in the FLA
        strFolderName=Environment.Value("HostingPath")&strtemplateName
        'creating of  file System object 
        Set fso = CreateObject("Scripting.FileSystemObject")

        If fso.FolderExists(strFolderName) = false Then
        ' creating the Folder Structure
         fso.CreateFolder (strFolderName)
        End If

        Set DataFolder = fso.GetFolder(SourcePath) 
        ' To  get the Files in the Folder Structure
        Set DataFiles = DataFolder.Files 

            '  Retriving the File name one by one 
            For Each oFiles in DataFiles 
            FileName=oFiles.name
            Fileextension=fso.GetExtensionName(oFiles.name)            
               fso.MoveFile SourcePath&"\"&FileName, SourcePath&"\"&strtemplateName&"."&Fileextension
            fso.CopyFile SourcePath&"\"&strtemplateName&"."&Fileextension,strFolderName&"\"
                                
            Next 

        For i=1 to 400

            wait 2

            If fso.FolderExists(strFolderName) = false Then
            Call ReportStep (StatusTypes.Pass, "Hosting Template:-"&Space(7)&strtemplateName&Space(5)&"Template Hosted Successfully","Hosting Server Window")
            
            Exit for 

           End If
        Next


    If  fso.FolderExists(strFolderName) =True Then
    Call ReportStep (StatusTypes.Fail, "Hosting Template:-"&Space(7)&strtemplateName&Space(5)&"Template Not Hosted Successfully","Hosting Server Window")
    End If


    'releasing the Object created in the previous steps
    Set DataFolder=nothing
    set DataFiles=nothing
    Set fso=nothing    
        
    End Sub

'===========================================================================================================================================================
    '<Adding the User to the Template either for the QA User OR Cube User and for Cube Admin>
    '<strTemplateID :- Template Name>
    '<strManage_Users: Param to hold list of users to be added for Manage Users Section>
	'<strRoleSel: role selection>
	'<strF1: temprory param>
	'<strF2: temprory param>
    '<Applysecuritysel :-Apply the Security to the Added Users>
    '<operation: Operation of the Manager Users either to delete the User or Add the Users>
    '<Author : Shweta Nagaral>    
'===========================================================================================================================================================    '    
    Public Sub TemplateManageUsers(ByVal strTemplateID,ByVal strManage_Users,ByVal Applysecuritysel,ByVal operation,ByVal strRoleSel,ByVal strF1,ByVal strF2)
    
    Dim objlnkTemplate,objtxtTemptab,WshShell,objtemplatetab1,strproperty,objwebManager,objtxttemTab2,objtemplatetab2,objbtnApplySecurity,btncloseok
    '<looping Variables>
    Dim j,h
    Dim objtxtUser,Objbtnclose,objbtnAddUser,objbtnCancel,objbtnMaDeleteUser
    
    
    Set objlnkTemplate=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Link("lnkTemplates")
    Call SCA.ClickOn(objlnkTemplate,"LnkTemplate" , "Hosting Templates")

    Set objtxtTemptab=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebEdit("txtTemplateIdTab")
    Call SCA.SetText(objtxtTemptab,strTemplateID, "txtTemplateTab","Hosting Templates Page" )
    
    set WshShell = CreateObject("WScript.Shell")
    
    Set objtemplatetab1=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebEdit("txtTemplateIdTab")
    Call SCA.ClickOn(objtemplatetab1,"LnkTemplate" , "Hosting Templates Page" )
    
    wait 1
    WshShell.SendKeys "{ENTER}"
    wait 4


    For j=1 to 60
        strproperty=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebTable("tabTemplatename").GetCellData(2,7)
        
        If strproperty="Manage Users" Then
            Set objwebManager=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebElement("webManageUsers")
            Call SCA.ClickOn(objwebManager,"WebManager Users" , "Hosting Templates Page")
        Exit For
        else
        wait 3
        
        Set objtxttemTab2=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebEdit("txtTemplateIdTab")
        Call SCA.SetText(objtxttemTab2,strTemplateID, "txtTemplateTab","Hosting Templates Page" )
        
        Set objtemplatetab2=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebEdit("txtTemplateIdTab")
        Call SCA.ClickOn(objtemplatetab2,"LnkTemplate" , "Hosting Templates Page" )
        WshShell.SendKeys "{ENTER}"
        
        End If

    Next
    wait 2
    
    'Shweta<10/9/2015 - Start
'    Select Case strRoleSel
'        
'        Case "QAUsers"
'        
'         Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebCheckBox("chkQAUser").Set "ON"
'         
'        Case "CubeAdmin"
'        
'        Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebCheckBox("CubeAdmin").Set "ON"
'        
'        Case "CubeUsers"
'        
'        Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebCheckBox("CubeUser").Set "ON"
'        
'        
'    End Select

	arrroles=Split(strRoleSel,";")
	
	For i = 0 To ubound(arrroles)
	    Set rolecheck=Description.Create()
	    rolecheck("MicClass").value="WebCheckBox"
	    rolecheck("name").value=arrroles(i)
	    Set chrolecheck=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").ChildObjects(rolecheck)
	    chrolecheck(0).Set "ON"
	Next
	wait 2

	'Shweta<10/9/2015 - End
    
    'Enter Users 
     Set objtxtUser=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebEdit("txtEnterUsers")
     Call SCA.SetText_MultipleLineArea(objtxtUser,strManage_Users, "CubeUserName","Hosting Templates Manager User" )
    'Shobha changes Ends here


    'Managing Add/Delete Users roles
    If  operation = "AddUsers" Then

            Set objbtnAddUser=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebButton("btnManage Add Users")
            Call SCA.ClickOn(objbtnAddUser,"Manage Add Users" , "Hosting Templates Manager User" )
            Wait 4
            Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
            
            If Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebElement("webUsers added to role").Exist then
              Call ReportStep (StatusTypes.Pass, "Managing Add Users roles:-"&Space(7)&"Successfully added user" &strManage_Users& " to cube " &strTemplateID,"Hosting Templates Manager User")        
              
            elseif Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebElement("webOne or more users Warn Msg").Exist Then             
              Call ReportStep (StatusTypes.Warning, "Applying security status:-"&Space(7)&"Selected User " &strManage_Users& " already exists","Hosting Templates Manager User")    

            elseif     Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebElement("webError already applying security").Exist Then
               Call ReportStep (StatusTypes.Fail, "Managing Add Users roles:-"&Space(7)&"Error Template " &strTemplateID& " is already applying security. Please try again later.Error occured during add operation","Hosting Templates Manager User")            
                                        
                Set Objbtnclose=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebButton("btnClose")
                Call SCA.ClickOn(Objbtnclose,"Manage close button" , "Hosting Templates Manager User" )
                
                If Browser("Hosting Templates - Ops").Dialog("Message from webpage").Exist Then
                    Set objbtnCancel=Browser("Hosting Templates - Ops").Dialog("Message from webpage").WinButton("btnCancel")
                    Call SCA.ClickOn(objbtnCancel,"Manage Cancel button" , "Hosting Templates Manager User" )
                End If

                                
                Exit sub    
                else
                    Call ReportStep (StatusTypes.Fail, "Managing Add Users roles:-"&Space(7)&"Could not add user" &strManage_Users& " to cube " &strTemplateID& " successfully","Hosting Templates Manager User")            
                    
                End if


    elseif operation = "DelUsers" Then

             Set objbtnMaDeleteUser=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebButton("btnManage Delete Users")
             Call SCA.ClickOn(objbtnMaDeleteUser,"Manage Delete User button" , "Hosting Templates Manager User" )
             Wait 4
              Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
              
               If Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebElement("webUsers deleted from role").Exist then
                  Call ReportStep (StatusTypes.Pass, "Managing Delete Users roles:-"&Space(7)&"Successfully deleted user " &strManage_Users& " to cube " &strTemplateID,"Hosting Templates Manager User")            
                  
             elseif Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebElement("webOne or more users Warn Msg").Exist Then
                 Call ReportStep (StatusTypes.Warning, "Applying security status:-"&Space(7)&"Selected user " &strManage_Users& " already deleted","Hosting Templates Manager User")    
                  
                    
             elseif Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebElement("webError already applying security").Exist Then
                 Call ReportStep (StatusTypes.Fail, "Managing Delete Users roles:-"&Space(7)&"Error Template " &strTemplateID& " is already applying security. Please try again later.Error occured during delete operation","Hosting Templates Manager User")    
                 
                  Set objbtnclose=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebButton("btnClose")
                  Call SCA.ClickOn(Objbtnclose,"Manage close button" , "Hosting Templates Manager User" )
                  
            If Browser("Hosting Templates - Ops").Dialog("Message from webpage").Exist Then
                Set objbtnCancel=Browser("Hosting Templates - Ops").Dialog("Message from webpage").WinButton("btnCancel")
                Call SCA.ClickOn(objbtnCancel,"Manage Cancel button" , "Hosting Templates Manager User" )
            End If
            
            Exit Sub
                            
    else
                
            Call ReportStep (StatusTypes.Fail, "Managing Add Users roles:-"&Space(7)&"Could not delete user" &strManage_Users& " to cube " &strTemplateID& " successfully","Hosting Templates Manager User")            
                    
            End if

    End If    
    
    If Applysecuritysel=0 Then
        
        wait 2
        Set objbtnApplySecurity=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebButton("btnApply Security")
        Call SCA.ClickOn(objbtnApplySecurity,"Apply Security btn" , "Hosting Templates Manager User")
        wait 2
            For h=1 to 50
                If Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebElement("webApplying security in progress").Exist Then
                Call ReportStep (StatusTypes.Pass, "Applying security status:-"&Space(7)&"Successfully applied security to cube " &cube& " for user " &strManage_Users,"Hosting Templates Manager User")            
                     Exit For    
                End If
                wait 3
                Set objbtnApplySecurity=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebButton("btnApply Security")
                Call SCA.ClickOn(objbtnApplySecurity,"Apply Security btn" , "Hosting Templates Manager User")
            Next
        If Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebElement("webApplying security in progress").Exist(0)=False Then
            Call ReportStep (StatusTypes.Warning, "Applying security status:-"&Space(7)&"Cannot apply security to cube " &strTemplateID& " to check the existance of DSE File in Dynamic Security File","Hosting Templates Manager User")    
            
        End If
    End If
    'Close the Manage Users dialog
    Set btncloseok=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebButton("btnClose")
    Call SCA.ClickOn(btncloseok,"btnclose" , "Hosting Templates Manager User")
    set WshShell=nothing
    
    End Sub

'===========================================================================================================================================================
    '<Template Delection from Ops System>
    '<strtemp_id :- Template Name>
    '<Author : Shweta Nagaral>    
'===========================================================================================================================================================
    Public Sub TemplateDelection_OpsSystem(ByVal strtemp_id,ByVal strF1,ByVal strF2)
    
    Dim objimghead_logo,oShell,objtxtTemID,objtxtTemplateTab,objClick,objDelDataTem,objbtnOK
    Dim i
        
    Set objimghead_logo=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Image("imghead_logo")
    If objimghead_logo.Exist(2) Then
		Call SCA.ClickOn(objimghead_logo,"imghead_logo" , "Ops Hosting Templates")
    End If
    
	Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync 
	Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").RefreshObject
	
	Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebEdit("txtTemplateIdTab").Set strtemp_id
	wait 1
	
	set WshShell = CreateObject("WScript.Shell")
	wait 1
	Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebEdit("txtTemplateIdTab").Click
	wait 1
	WshShell.sendkeys "{ENTER}"
	wait 1
	WshShell.sendkeys "{ENTER}"
	wait 1
	Set objClick=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebTable("tabTemplatename").ChildItem(2,1,"WebElement",0)
	objClick.click
	wait 2
    
    Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync 
    Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").RefreshObject
    Call ReportStep (StatusTypes.Information,"Template Delection in Ops System", "Ops Page")
    
    Set objDelDataTem=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebButton("btnDelete Template and Data")
    Call SCA.ClickOn(objDelDataTem,"Delete Template and Data" , "Ops Hosting Templates")
    
    wait 1
    
    Set objbtnOK=Browser("Hosting Templates - Ops").Dialog("Message from webpage").WinButton("btnOK")
    If objbtnOK.Exist(2) Then
    	Call SCA.ClickOn(objbtnOK,"Ok Button" , "Ops Hosting Templates")
    End If
    
    For i = 1 To 10 Step 1
    	If  Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebElement("webHosting Template Deleted.").Exist(2) Then
	        Call ReportStep (StatusTypes.Pass, strtemp_id&space(5)&"deleted from the Ops Successfully", "Ops Page")
	        Exit For
	    End If
	    
	    If i = 10 Then
	    	Else
	        Call ReportStep (StatusTypes.Fail,"Template Deletion in Ops System "&strtemp_id&" Not Deleted from the Ops Successfully", "Ops Page")
	    End If
    Next
    
    Set objbtnClose=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebButton("btnClose")
    Call SCA.ClickOn(objbtnClose,"btnClose" , "Ops Hosting Templates")
    
    set oShell=Nothing
    Set objClick=Nothing

   End Sub
    
'===========================================================================================================================================================   
    
    '    '<Cube Acceptance and rejection in Ops System>
    '    '<strtemp_id :- Strtemplate ID>
    '    '<num_cubesel :-Selection for the acceptance and Rejection , 0 for acceptance and 1 for rejection>
    '    '<author :-Shobha>'
'===========================================================================================================================================================
    Public Sub AcceptRejectCube_Inops(ByVal strtemp_id,ByVal num_cubesel,ByVal strF1,ByVal strF2)
        
    '<Shweta.Nagaral <30/6/2015> -Start>
    Set appobject=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops")
    
    Set objlnkHosting = Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Link("lnkHosting")
    Call SCA.ClickOn(objlnkHosting,"lnkReport" , "Hosting Templates")
    
    Call IMSSCA.General.menuitemselect(appobject,"QA")
    '<Shweta.Nagaral <30/6/2015> -End>
    
    Browser("Hosting Templates - Ops").Page("QA Page - Ops System").Sync
    Browser("Hosting Templates - Ops").Page("QA Page - Ops System").Sync
    Browser("Hosting Templates - Ops").Page("QA Page - Ops System").WebEdit("txtCube").Set strtemp_id
    Browser("Hosting Templates - Ops").Page("QA Page - Ops System").WebEdit("txtCube").Click

    Set wobj=createobject("WScript.Shell")
    wobj.SendKeys "{ENTER}"
    wait 3
    For i=1 to 60
        If Browser("Hosting Templates - Ops").Page("QA Page - Ops System").WebElement("webready for QA").Exist(2) Then
            Exit For
        End If
    ' Imp This Should be added as a part of the code Ask shweta'

''            wait 3
'            Browser("Hosting Templates - Ops").Close
'            call OpenApp(Environment("OPSUrl"), "i")
'            call OpenApp(url, "i")
'            RunAction "Action1 [Re-Usable Actions\Login_OpsSystem]", oneIteration,Parameter("UserName"),Parameter("Pwd")
'            wait 2    
'            'Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Link("QA").Click
'            Browser("Hosting Templates - Ops").Page("QA Page - Ops System").WebElement("webLoading ...").WaitProperty "visible","False",10000
'            Browser("Hosting Templates - Ops").Page("QA Page - Ops System").WebEdit("txtCube").Set strtemp_id
'            Browser("Hosting Templates - Ops").Page("QA Page - Ops System").WebEdit("txtCube").Click
'            wobj.SendKeys "{ENTER}"
''
    Next

    If Browser("Hosting Templates - Ops").Page("QA Page - Ops System").WebElement("webready for QA").Exist=False Then
    Call ReportStep (StatusTypes.Warning,"Not able to accept or Reject the cube","Ops QA Page")
    Exit Sub
    End If

    If num_cubesel=0 Then
        Browser("Hosting Templates - Ops").Page("QA Page - Ops System").WebRadioGroup("Class Name:=WebRadioGroup;selected item index:=1").Select "#0"
     Call ReportStep (StatusTypes.Pass,"Cube Acceptance in Ops System. "&strtemp_id&"is accepted","Ops QA Page")
     Browser("Hosting Templates - Ops").Page("QA Page - Ops System").WebButton("btnApply").Click
     wait 1
     Browser("Hosting Templates - Ops").Page("QA Page - Ops System").WebButton("btnOk").Click
        If Browser("Hosting Templates - Ops").Page("QA Page - Ops System").WebElement("webSuccess").Exist(3)  Then
            Call ReportStep (StatusTypes.Pass,"Cube "&strtemp_id&"is accepted successfully","Ops QA Page")
        Else
            Call ReportStep (StatusTypes.Pass,"Cube "&trtemp_id&"is Not accepted successfully","Ops QA Page")
        End If
    Else    
    Browser("Hosting Templates - Ops").Page("QA Page - Ops System").WebRadioGroup("Class Name:=WebRadioGroup;selected item index:=1").Select "#1"
    Call ReportStep (StatusTypes.Pass,"Cube Rejection in Ops System "&strtemp_id&"is Rejected","Ops QA Page")
    Browser("Hosting Templates - Ops").Page("QA Page - Ops System").WebEdit("txtRejecteditBox").Set "Automation Testing"
    Browser("Hosting Templates - Ops").Page("QA Page - Ops System").WebButton("btnApply").Click
    wait 1
    Browser("Hosting Templates - Ops").Page("QA Page - Ops System").WebButton("btnOk").Click
          If Browser("Hosting Templates - Ops").Page("QA Page - Ops System").WebElement("webSuccess").Exist(3)  Then
            Call ReportStep (StatusTypes.Pass,"Cube "&strtemp_id&"is rejected successfully","Ops QA Page")
        Else
               Call ReportStep (StatusTypes.Pass,"Cube "&strtemp_id&"is Not reject successfully","Ops QA Page")
        End If
    End If

    Set wobj=nothing
   
        
    End Sub
'===========================================================================================================================================================
    
    '    '<Cube Acceptance and rejection in Ops System>
    '    '<strtemp_id :- Strtemplate ID>
    '    '<num_cubesel :-Selection for the acceptance and Rejection , 0 for acceptance and 1 for rejection>
    '    '<author :-Shobha>'
'===========================================================================================================================================================
	
  Public Function MoveDBInOpsSystem(ByVal strtemp_id, ByVal strServerName, ByVal intFMove, ByVal intFinalize, ByVal intInitialSetUp, ByVal strF2)	
		
   Dim strNow, strSecond, strMinute, strHour, strDay, strMonth, strYear, strmonthName, strRDate, strScheduleTime, objpage ,strstatusCap
   '<Objects Declaration >
   Dim objtxtTempID,lstServerName, objlstMonth, objlstYear, objtxtScdeduleDate,objlnkDate , objScheduleTime , objbtnMove, objDataSourceIdtxt, DataSourectxt, objlnkHosting
    
    
	strNow		= Now
	strSecond	= Second(strNow)
	strHour		= Hour(strNow)
	strDay		= Cstr(Day(strNow))
	strMonth 	= Cstr(Month(strNow))
	strYear 	= Cstr(Year(strNow))
	strmonthName=MonthName(strMonth,True)
	wait 5
	Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Link("lnkHosting").Click
	Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Link("lnkMove Database").Click

	'Shweta - 20/4/2016 - Initial Setup to validate whether CUbe template(DataSource) is repointing to expected DataSource or not- Start
	If intInitialSetUp <> "" Then
			Set WshShell = CreateObject("WScript.Shell")	
			Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebEdit("txtDataSourceId").Set strtemp_id
			Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebEdit("txtDataSourceId").Click
			
			wait 2
			WshShell.SendKeys "{ENTER}"
			
			If Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebTable("tabMoveDatabase").GetROProperty("rows") >=2 Then
				actualServer=Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebTable("tabMoveDatabase").GetCellData(2, 3)
				If actualServer = strServerName Then
					Exit Function
				End If
			Else
				Exit Function
			
			End If
	End If
	'Shweta - 20/4/2016 - Initial Setup to validate whether CUbe template(DataSource) is repointing to expected DataSource or not- End
	
	
	'<shweta  27/10/2015> - Start
'	Set objpage=Browser("Hosting Templates - Ops").Page("Move Database - Ops System")
'    Call IMSSCA.General.menuitemselect(objpage,"Move Database")
	
    '<shweta  27/10/2015> - End
	If strHour<9 AND strHour>0 Then
		strHour=0&strHour
		
	End If
	
	 
	
	'False
	If intFinalize=1 Then
	
'<shweta  27/10/2015> - Start
'    Set objpage=Browser("Hosting Templates - Ops").Page("Move Database - Ops System")
'    Call IMSSCA.General.menuitemselect(objpage,"Move Database")
'<shweta  27/10/2015> - End
	
	set objtxtTempID=Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebEdit("txtTemplateDatasourceId")
	Call SCA.SetText(objtxtTempID,strtemp_id , "TemplateDatasourceId","Move Database Page" )
	
	Set lstServerName=Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebList("lstServerName")
	Call SCA.SelectFromDropdown(lstServerName,strServerName)
	
	Set objtxtScdeduleDate=Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebEdit("txtScheduleDate")
	Call SCA.ClickOn(objtxtScdeduleDate,"ScheduleDate", "Move Database Page")
	
	Set objlstMonth=Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebList("lstcalMonth")
	Call SCA.SelectFromDropdown(objlstMonth,strmonthName)
	
	Set objlstYear=Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebList("lstcalYear")
	Call SCA.SelectFromDropdown(objlstYear,strYear)
	
	Browser("Hosting Templates - Ops").Page("Move Database - Ops System").Sync
	
	 
	Set objlnkDate=Description.Create
	objlnkDate("micclass").Value="Link"	
	objlnkDate("name").Value=strDay
	Set objlnkDate=Browser("Hosting Templates - Ops").Page("Move Database - Ops System").Link(objlnkDate)
	Call SCA.ClickOn(objlnkDate,"DateLink", "Move Database Page")



    strMinute = Minute(strNow)
    strScheduleTime=strHour&":"&strMinute+2
    '<Shweta - 1/6/2016> Changed time by 1 min - Start
    'strScheduleTime=strHour&":"&strMinute+1
    '<Shweta - 1/6/2016> Changed time by 1 min - End
    
    If strMinute<9 AND strMinute>0 Then
        strMinute=0&strMinute
        
    End If
    If strMinute=59  Then        
        strHour=strHour+1
        strMinute=00        
    End If    
    
	
	Set objScheduleTime=Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebEdit("txtScheduleTime")
	Call SCA.SetText(objScheduleTime,strScheduleTime , "ScheduleTime","Move Database Page" )
	
	Set objbtnMove=Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebButton("btnMove Database")
	Call SCA.ClickOn(objbtnMove,"Move Database Btn", "Move Database Page")
	
	 
	'<shweta 22/4/2016> - start
	For i = 1 To Environment.Value("intCounterMaxLimit") Step 1
		If Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebElement("txtInformation saved successfully").Exist(5) Then
			Call ReportStep (StatusTypes.Pass,"Clicked on Btn Move Database Successfully", "Move Database Page") 	  
			Exit for
		End If	
		
		If i = Environment.Value("intCounterMaxLimit") Then
			Call ReportStep (StatusTypes.Fail,"Not Clicked on Btn Move Database Successfully", "Move Database Page") 
		End If
	Next
	
'	If Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebElement("txtInformation saved successfully").Exist(0) Then
'	   
'	 Call ReportStep (StatusTypes.Pass,"Clicked on Btn Move Database Successfully", "Move Database Page") 	   
'	 else
'	 Call ReportStep (StatusTypes.Fail,"Not Clicked on Btn Move Database Successfully", "Move Database Page") 
'		
'	End If
	'<shweta 22/4/2016> - end
	
    Set WshShell = CreateObject("WScript.Shell")	
    Set objDataSourceIdtxt=Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebEdit("txtDataSourceId")
    Call SCA.SetText(objDataSourceIdtxt,strtemp_id , "DataSource Id Txt","Move Database Page" )
    
    Set DataSourectxt=Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebEdit("txtDataSourceId")
    Call SCA.ClickOn(DataSourectxt,"txt DataSource Id", "Move Database Page")
    wait 2
    WshShell.SendKeys "{ENTER}"
    strstatus=Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebTable("tabMoveDatabase").GetCellData(2,4)
    
    If strstatus="Scheduled" Then
     
     For k = 1 To 70  Step 1
     Set objlnkHosting=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Link("lnkHosting")
	 Call SCA.ClickOn(objlnkHosting,"Link Hosting", "Move Database Page")
	 
	 Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Link("lnkMove Database").Click
   '  Call IMSSCA.General.menuitemselect(objpage,"Move Database")
     Set objDataSourceIdtxt=Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebEdit("txtDataSourceId")
     Call SCA.SetText(objDataSourceIdtxt,strtemp_id , "DataSource Id Txt","Move Database Page" )
    
     Set DataSourectxt=Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebEdit("txtDataSourceId")
     Call SCA.ClickOn(DataSourectxt,"txt DataSource Id", "Move Database Page")
     WshShell.SendKeys "{ENTER}"
     wait 4
     strstatus=Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebTable("tabMoveDatabase").GetCellData(2,4)     
     If strstatus="Completed" OR strstatus="Error" OR strstatus="Finalized" Then
     	
     	MoveDBInOpsSystem=strstatus
     	Exit For
     End If     
    	
   	 Next
        	
    End If
    
    If strstatus="Completed" AND intFMove=0 Then
    	
    	Set ObjFC=Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebTable("tabMoveDatabase").ChildItem(2,9,"WebElement",0)
		ObjFC.click	
		wait 3		
		strstatus=Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebTable("tabMoveDatabase").GetCellData(2,4) 
		For i = 1 To 150 Step 1
		wait  2
		strstatusCap=Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebTable("tabMoveDatabase").GetCellData(2,4)
		If strstatusCap="Finalized" Then
			Exit For
		End If	
			
		Next
    	MoveDBInOpsSystem=strstatusCap
    	
    End If
	 
	ElseIf intFinalize=0  then
	
	 Set objpage=Browser("Hosting Templates - Ops").Page("Move Database - Ops System")	
	 Call IMSSCA.General.menuitemselect(objpage,"Move Database")
     Set WshShell = CreateObject("WScript.Shell")

	 Set objDataSourceIdtxt=Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebEdit("txtDataSourceId")
     Call SCA.SetText(objDataSourceIdtxt,strtemp_id , "DataSource Id Txt","Move Database Page" )
    
     Set DataSourectxt=Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebEdit("txtDataSourceId")
     Call SCA.ClickOn(DataSourectxt,"txt DataSource Id", "Move Database Page")
     
     wait 2
     WshShell.SendKeys "{ENTER}"
     strstatus=Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebTable("tabMoveDatabase").GetCellData(2,4)
     wait 3
     If strstatus="Completed" Then
     	
		'Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebElement("webFinalize Cube Move").Click
		Set ObjFC=Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebTable("tabMoveDatabase").ChildItem(2,9,"WebElement",0)
		ObjFC.click
		strstatus=Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebTable("tabMoveDatabase").GetCellData(2,4) 
		wait 3
		For i = 1 To 150 Step 1
		wait 3
		strstatusCap=Browser("Hosting Templates - Ops").Page("Move Database - Ops System").WebTable("tabMoveDatabase").GetCellData(2,4)
		If strstatusCap="Finalized" Then
			Exit For
		End If	
			
		Next
		
    	MoveDBInOpsSystem=strstatusCap
     	
     End If
	
	
	End If
	
	Set WshShell = Nothing
     
    
		
		
	End Function
	
	
'===========================================================================================================================================================   
    
    '    '<BooKMarkCreation in SCA for the Created Report>
    '    '<strBName :-BookMark Name>
    '    '<strOSelection :-Selection for Required operations , ToSave/To Open BookMark>
    '    '<author :-Shobha>'
'===========================================================================================================================================================
	
	Public Function BooKMarkCreation(ByVal strBName,ByVal strOSelection, ByVal strF1, ByVal strF2)
	
	Dim rc,cc,objBM,objBMChild,strBMName,ArrValue_BFilter,i,j,strCellValue_BFilter,strCellValues_BFilter,m,n,objwebtable
	'<Object Declaration >
	Dim ObjImdaddBM,objtxtBName, objtxtDescr, objokbtn,objimgDBM
	
    Browser("Analyzer").Page("ReportCreation").Sync
    Browser("Analyzer").Page("ReportCreation").RefreshObject
	Select Case strOSelection
		
		Case "SaveBM"
		
		 Set ObjImdaddBM=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgAddbookmark")
		 Call SCA.ClickOn(ObjImdaddBM,"Add BookMark btn", "Report Creation Page")
		 
	   	 Set objtxtBName=Browser("Analyzer").Page("ReportCreation").Frame("BookMarkFrame").WebEdit("txtName")
	   	 Call SCA.SetText(objtxtBName,strBName, "BooK Mark txt","Report Creation Page" )
	   	 
	   	 Set objtxtDescr=Browser("Analyzer").Page("ReportCreation").Frame("BookMarkFrame").WebEdit("txtDescription")
	   	 Call SCA.SetText_MultipleLineArea(objtxtDescr,"Automation Testing", "BooK Mark Description txt","Report Creation Page" )
	   	 
	   	 Set objokbtn=Browser("Analyzer").Page("ReportCreation").Frame("BookMarkFrame").WebButton("OK")
		 Call SCA.ClickOn(objokbtn,"btn Ok", "Report Creation Page")	   	 
		
		Case "Display_BM"
		
		 If Browser("Analyzer").Page("ReportCreation").Frame("BMDisplayFrame").Exist(0) Then
		 	Browser("Analyzer").Page("ReportCreation").Sync
		 	else
		 	 Set objimgDBM=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgDisplaybookmarks")
		 	 Call SCA.ClickOn(objimgDBM,"Display Book Mark", "Report Creation Page")
		 	
		 End If
		
		 wait 1
		 Set objBM=Description.Create
		 objBM("micclass").value="WebElement"
		 objBM("class").value="UltraWebListbar1ctl0BookmarkList1trvBookmarksNode TreeNode"
		 objBM("html tag").value="SPAN"
		 
				 
		 Set objBMChild=Browser("Analyzer").Page("ReportCreation").Frame("BMDisplayFrame").ChildObjects(objBM)
		 
		 For i = 0 To objBMChild.count-1 Step 1
		 	strBMName=objBMChild(i).GetRoproperty("innertext")
		 	If Strcomp(strBMName,strBName)=0 Then
		 		objBMChild(i).Click
		 		Call ReportStep (StatusTypes.Pass,strBName&Space(3)&"Displayed in the BookMark Pannel of SCA Report Creation","ReportCReation Page in SCA" )
		 		wait(2.5)
		 		Set objwebtable=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabReportC")
	            ArrValue_BFilter=SCA.Webtable(objwebtable,"Dimension Value table","Retriving_DataTableValue","Report Creation","","","","")
	
	             For m=0 to ubound(ArrValue_BFilter,1)
		          For n=0 to ubound(ArrValue_BFilter,2)
			         strCellValue_BFilter= ArrValue_BFilter(i,j)
			         strCellValues_BFilter=strCellValues_BFilter&";"&strCellValue_BFilter
		         Next
	             Next	
			BooKMarkCreation=strCellValues_BFilter	
			Exit For	
		 	End If
		 	
		 	
		 Next	
		
		
	End Select
	 
    Browser("Analyzer").Page("ReportCreation").Sync
    Browser("Analyzer").Page("ReportCreation").RefreshObject
 End Function


'===========================================================================================================================================================

    '    '<Report Re-Point page fields validation in Ops System>
    '    '<strtemp_id :- Strtemplate ID>
    '     '<'radioDBtype:= Live;QA;Prior>
    '    '<num_cubesel :-Selection for the acceptance and Rejection , 0 for acceptance and 1 for rejection>
    '    '<Author : Shweta Nagaral>
'===========================================================================================================================================================
Public Function ReportRepoint(ByVal strReportName, ByVal strActualDatabase, ByVal strActualCube, ByVal strOldDataSrcName, ByVal repoint, ByVal strNewDataSrcName, ByVal bookmarkValidation, ByVal aBookMark, ByVal radioDBtype, ByVal reportExists)
    
    Dim objReportName, WshShell, strDataBase, strCube, strDataSrcName, objDatabase, objNewDataSrcName, objbtnOk, strBookMarks, aBookMarkEdit, objReportLink, objCloseOk, compareVal
    Dim i, j, k, l
    aBookMarkEdit = ""
    
    
    '<Shweta 2/2/2015> - Commented to handled sync issues - Start 
'    Set WshShell =CreateObject("WScript.Shell")
'	WshShell.SendKeys "{F5}"
'	wait 2
'	WshShell.SendKeys "{F5}"
'
'	Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
'	Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").RefreshObject
'	
'    set objHostingPage = Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops")
'    
'    Set objReportLink = Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Link("lnkReport")
'    Call SCA.ClickOn(objReportLink,"lnkReport" , "Hosting Templates")
'    wait 2
'    
'    Call IMSSCA.General.menuitemselect(objHostingPage,"Repoint")

	'commented
	Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
	Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").RefreshObject
    set objHostingPage = Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops")
    Set objReportLink = Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Link("lnkReport")
    Set objReportName = Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtReportName")
    Set WshShell =CreateObject("WScript.Shell")

	For a = 1 To 5 Step 1
        WshShell.SendKeys "{F5}"
        wait 2
        WshShell.SendKeys "{F5}"
        Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
        Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").RefreshObject

		Call SCA.ClickOn(objReportLink,"lnkReport" , "Hosting Templates")
		wait 2
		
		'Call IMSSCA.General.menuitemselect(objHostingPage,"Repoint")
		Call menuitemselect(objHostingPage,"Repoint")
        Browser("Hosting Templates - Ops").Page("Clients - Ops System").Sync
        Browser("Hosting Templates - Ops").Page("Clients - Ops System").RefreshObject
        
        objReportName.WaitProperty "visible",True,10000
        If objReportName.Exist(1) Then
            Call ReportStep(StatusTypes.Pass, "Successfully redirected to 'Report Repiont' Page in OPS", "OPS Report Repiont Page")
            Exit For
        End If
        
        If a=5 Then
            Call ReportStep(StatusTypes.Fail, "Could not successfully redirected to 'Report Repiont' Page in OPS", "OPS Report Repiont Page")
            Call ReportStep(StatusTypes.Warning, "Could not successfully redirected to 'Report Repiont' Page in OPS. Further functionality Validation steps may fail", "OPS Report Repiont Page")
        End If
    Next
	
	'Enter ReportName in grid tab and Validate DataSorce/Cube/DataBase details
    '<Shweta 2/2/2015> - Commented to handle sync issues- End
    
    
    '<Shweta - 25/11/2015 > Entering report name if its not null - Start
    If strReportName <> "" Then
    		'Using Shell object to enter
		    Set WshShell = CreateObject("WScript.Shell")
		    'Set objReportName = Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtReportName")
		    Call SCA.SetText(objReportName,strReportName, "txtReportName","Report Repoint Page" )
		    wait 2
		    WshShell.SendKeys "{ENTER}"
		    wait 1
		    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtReportName").Click
		    WshShell.SendKeys "{ENTER}"
		    Set WshShell = Nothing
		     	
		 	'Shweta - 26/7/2016> - Start
		 	Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtReportName").RefreshObject
		 	'Shweta - 26/7/2016> - End
    End If
    '<Shweta - 25/11/2015 > Entering report name if its not null - End

    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").Sync
    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").RefreshObject
    wait 2
    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").Sync
    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").RefreshObject
            
    For l = 1 To 10 Step 1
    	Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").RefreshObject
        If Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebTable("tabReport Repoint").Exist(1) Then
            rc = Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebTable("tabReport Repoint").GetROProperty("rows")
        else
            Call ReportStep(StatusTypes.Fail, "Report Repoint grid table is not displaying", "Report Re-point page")            

        End if
        
     	If rc > 2 Then
	    	Set WshShell = CreateObject("WScript.Shell")
	    	Set objtxtGridValueDatabase = Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtGridValueDatabase")
		    Call SCA.SetText(objtxtGridValueDatabase,strActualDatabase, "txtGridValueDatabase","Report Repoint Page" )
		    wait 2
		    WshShell.SendKeys "{ENTER}"
		    wait 1
		    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtReportName").Click
		    WshShell.SendKeys "{ENTER}"
		    Set WshShell = Nothing
		    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").Sync
		    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").RefreshObject
		    wait 1
		    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").Sync
    		Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").RefreshObject
    		 	
		 	'Shweta - 26/7/2016> - Start
		 	Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtReportName").RefreshObject
		 	'Shweta - 26/7/2016> - End
    	End If
    	
    	rc = Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebTable("tabReport Repoint").GetROProperty("rows")
    	
    	If rc > 2 Then
			Set WshShell = CreateObject("WScript.Shell")
			Set objtxtGridValueCube = Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtCube")
		    Call SCA.SetText(objtxtGridValueCube,strActualCube, "txtCube","Report Repoint Page" )
		    wait 2
		    WshShell.SendKeys "{ENTER}"
		    wait 1
		    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtReportName").Click
		    WshShell.SendKeys "{ENTER}"
		    Set WshShell = Nothing
		    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").Sync
		    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").RefreshObject
		    wait 1
		    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").Sync
    		Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").RefreshObject
    		
    		'Shweta - 26/7/2016> - Start
		 	Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtReportName").RefreshObject
		 	'Shweta - 26/7/2016> - End
   		End If
    
    	'Shweta - 26/7/2016> - Start
	 	Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtReportName").RefreshObject
	 	'Shweta - 26/7/2016> - End
	 	
        If reportExists = 0 Then
            If rc >= 2 Then
                Reporter.ReportEvent micPass, strReportName&" Report saved in SCA found and displaying in Report Repoint grid at "&rc&" row", "Report Re-point page"
                Call ReportStep(StatusTypes.Pass, strReportName&" Report saved in SCA found and displaying in Report Repoint grid at "&rc&" row", "Report Re-point page")
                ReportRepoint = "exists"
                Exit For
            End if
            
            If l =10 Then
                'Reporter.ReportEvent micFail, strReportName&" Report saved in SCA is not displaying in Report Repoint grid. But Report is supposed to display in repoint grid", "Report Re-point page"
                'Reporter.ReportEvent micFail, strReportName&" Report saved in SCA could not find in Report Repoint grid", "Report Re-point page"
                ReportRepoint = "Doesnt exists"
            End If
        else
            If rc = 1 Then
                'Reporter.ReportEvent micPass, strReportName&" Report saved in SCA is not displaying in Report Repoint grid", "Report Re-point page"
                ReportRepoint = "Doesnt exists"
                Exit function
            else
                'Reporter.ReportEvent micFail, strReportName&" Report saved in SCA is displaying in Report Repoint grid. Report is not supposed to display in repoint grid but displaying at "&rc&" row", "Report Re-point page"
                ReportRepoint = "exists"
            End if
            Exit For        
        End If
    Next
    
    'Shweta - 26/7/2016> - Start
 	Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtReportName").RefreshObject
 	'Shweta - 26/7/2016> - End
 	
	Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").Sync
	Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").RefreshObject
	Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebTable("tabTemplatename").RefreshObject
	wait 5
	'shweta - 3/6/2016 - Start	
	For i = 1 To 5 Step 1
		Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").RefreshObject
		Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebTable("tabTemplatename").RefreshObject
		wait 2
		rowCount = Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebTable("tabTemplatename").RowCount
		If rowCount = 2 Then
			Call ReportStep(StatusTypes.INformation, "No rows found in Report Repoint grid for Report "&strReportName, "Report Re-point page")
			Exit For
		End If
	Next
	'shweta - 3/6/2016 - End
	
	'Validate DataSorce/Cube/DataBase details
    strDataBase=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebTable("tabTemplatename").GetCellData(2,8)
    strCube=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebTable("tabTemplatename").GetCellData(2,9)
    strDataSrcName=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebTable("tabTemplatename").GetCellData(2,11)
    
    wait 2
    If Trim(Ucase(strDataBase)) = Trim(Ucase(strActualDatabase)) Then
        Reporter.ReportEvent micPass, strActualDatabase&" database found in grid for Report "&strReportName, "Report Re-point page"
        Call ReportStep(StatusTypes.Pass, strActualDatabase&" database found in grid for Report "&strReportName, "Report Re-point page")
    Else
    	Call ReportStep(StatusTypes.Information, strActualDatabase&" database with specific to its Hosting status found in grid for Report "&strReportName, "Report Re-point page")
'        Reporter.ReportEvent Infromation, strActualDatabase&" database could not find in grid for Report "&strReportName, "Report Re-point page"
'        Call ReportStep(StatusTypes.Fail, strActualDatabase&" database could not find in grid for Report "&strReportName, "Report Re-point page")
'        Call ReportStep(StatusTypes.Fail, "Mismatch between expected Database:= "&strActualDatabase&" and Application displaying database:= "&strDataBase, "Report Re-point page")
    End If
    
    wait 2
    If Trim(Ucase(strCube)) = Trim(Ucase(strActualCube)) Then
        Reporter.ReportEvent micPass, strActualCube&" cube found in grid for Report "&strReportName, "Report Re-point page"
        Call ReportStep(StatusTypes.Pass, strActualCube&" cube found in grid for Report "&strReportName, "Report Re-point page")
    Else
        Reporter.ReportEvent micFail, strActualCube&" cube could not find in grid for Report "&strReportName, "Report Re-point page"
        Call ReportStep(StatusTypes.Fail, strActualCube&" cube could not find in grid for Report "&strReportName, "Report Re-point page")
        Call ReportStep(StatusTypes.Fail, "Mismatch between expected Cube:= "&strActualCube&" and Application displaying Cube:= "&strCube, "Report Re-point page")
    End If
        
    wait 2
   
'Shweta 3/5/2016 Due to change in application Grid display of DataSource appended "*" in strOldDataSrcName based on application demand - Start 
   If instr(1, strDataSrcName, "*")>0 Then
		If instr(1, strOldDataSrcName, "*")<=0 Then
			strOldDataSrcName = strOldDataSrcName&"*"
		End If
	End If
'Shweta 3/5/2016 Due to change in application Grid display of DataSource appended "*" in strOldDataSrcName based on application demand - End	

    If Trim(Ucase(strDataSrcName)) = Trim(Ucase(strOldDataSrcName)) Then
        Reporter.ReportEvent micPass, strOldDataSrcName&" datasource found in grid for Report "&strReportName, "Report Re-point page"
        Call ReportStep(StatusTypes.Pass, strOldDataSrcName&" datasource found in grid for Report "&strReportName, "Report Re-point page")
    Else
        Reporter.ReportEvent micFail, strOldDataSrcName&" datasource could not find in grid for Report "&strReportName, "Report Re-point page"
        Call ReportStep(StatusTypes.Fail, strOldDataSrcName&" datasource could not find in grid for Report "&strReportName, "Report Re-point page")
        Call ReportStep(StatusTypes.Fail, "Mismatch between expected Datasource:= "&strOldDataSrcName&" and Application displaying Datasource:= "&strDataSrcName, "Report Re-point page")
    End If

    
    If bookmarkValidation = 0 Then
        strBookMarks=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebTable("tabTemplatename").GetCellData(2,6)
        For j = 0 To UBound(aBookMark) Step 1
            If j <> UBound(aBookMark) Then
                aBookMarkEdit = aBookMarkEdit&aBookMark(j)&", "
            Else
                aBookMarkEdit = aBookMarkEdit&aBookMark(j)
            End If
        Next
        
        '<Shweta 15/9/2015> - Start
'        If Trim(Ucase(strBookMarks)) = Trim(Ucase(aBookMarkEdit)) Then
'            Reporter.ReportEvent micPass, strBookMarks&" BookMarks found in grid for Report "&strReportName, "Report Re-point page"
'        Else
'            Reporter.ReportEvent micFail, strBookMarks&" BookMarks could not find in grid for Report "&strReportName, "Report Re-point page"
'        End If

		compareVal = InStr(1, Trim(Ucase(strBookMarks)), Trim(Ucase(aBookMarkEdit)))
	    If compareVal >= 1 Then
	        Reporter.ReportEvent micPass, strBookMarks&" BookMark/BookMarks found in grid for Report "&strReportName, "Report Re-point page"
	         Call ReportStep(StatusTypes.Pass, strBookMarks&" BookMark/BookMarks found in grid for Report "&strReportName, "Report Re-point page")
	    Else
	    	Reporter.ReportEvent micFail, strBookMarks&" BookMark/BookMarks could not find in grid for Report "&strReportName, "Report Re-point page"
	    	Call ReportStep(StatusTypes.Fail, strBookMarks&" BookMark/BookMarks could not find in grid for Report "&strReportName, "Report Re-point page")
	    End If
	    '<Shweta 15/9/2015> - End
	    
    End If
    
    
    'Enter details in Destination move and provide Environment, Confirm the report repoint
    If repoint = 0 Then
        rc = Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebTable("tabReport Repoint").GetROProperty("rows")
        If rc = 2 Then
            Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebTable("tabTemplatename").ChildItem(2, 1, "WebCheckBox", 0).click    
        '<Shweta - 25/11/2015 > If report name value is null - Start
        ElseIf rc >= 2 Then
        	Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebCheckBox("chkReportRepointSelection").Set "ON"
        '<Shweta - 25/11/2015 > If report name value is null - End
        End If

        'Confirm the report repoint
        '<Shweta - 27/7/2016> - Start Handled Report Repoint status message through Do until loop - Start
        intLoopCounter = 0
		Do
			If intLoopCounter=5 Then
				Exit Do 
		    End If
		 	intLoopCounter=intLoopCounter+1
      		
      		If Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebButton("btnConfirmationClose").Exist(2) = True Then
      			Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebButton("btnConfirmationClose").Click
      			Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").Sync
      		End If 
      		
      		Set WshShell = CreateObject("WScript.Shell")
      		Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtDestinationDataSource").Click
			Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtDestinationDataSource").Set ""
			Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtDestinationDataSource").Click
	        WshShell.SendKeys strNewDataSrcName
	        wait 2 ' To wait for List to display as per keyword entered above
	        WshShell.SendKeys "{DOWN}" 
	        wait 2 'To select item from list and wait
	        WshShell.SendKeys "{ENTER}"
	        'WshShell.SendKeys "{TAB}"
	        WshShell.SendKeys "{DOWN}" 
	        wait 2 'To select item from list and wait
	        WshShell.SendKeys "{ENTER}"
	        
	        wait 2
	        
	        Set WshShell1 = CreateObject("WScript.Shell")
	        Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtDestinationDatabase").Click
	        Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtDestinationDatabase").Set ""
	        Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtDestinationDatabase").Click
	        WshShell1.SendKeys strActualDatabase
	        wait 2 ' To wait for List to display as per keyword entered above
	        WshShell1.SendKeys "{DOWN}" 
	        wait 2 'To select item from list and wait
	        WshShell1.SendKeys "{ENTER}"
	        Set  WshShell1 = Nothing
	        wait 2
	        '<shweta 24/8/2015> - End
	        
	        Set objradioDBtype = Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebRadioGroup("rdoDBType")
	        Call SCA.ClickOnRadio(objradioDBtype,radioDBtype, "Report Repoint Page")
	        
	        Set objbtnRepoint=Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebButton("btnRepoint Selected Reports")
	        Call SCA.ClickOn(objbtnRepoint,"btnRepoint Selected Reports" , "Report Repoint Page")
        
        Loop Until Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebElement("welReportRepointMsg").Exist(2) = True
        '<Shweta - 27/7/2016> - Start Handled Report Repoint status message through Do until loop - End

		If Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebElement("welReportRepointMsg").Exist(2) Then
                repointMsg = Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebTable("tabConfirmation").GetCellData(1, 2)
                
                Set objbtnOk = Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebButton("btnConfirmationOk")
                Call SCA.ClickOn(objbtnOk,"btnConfirmationOk" , "Report Repoint Page")
                
                wait 2
                Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").Sync
        ElseIf intLoopCounter=5 Then 
	         '<shweta - 31/5/2016> - Changed report statements - start
	'                Reporter.ReportEvent micFail, "Report "&strReportName&" could not repoint successfully", "Report Re-point page"
	'                Call ReportStep(StatusTypes.Fail, "Report "&strReportName&" could not repoint successfully", "Report Re-point page")
	            Reporter.ReportEvent micInfo, "No success message after repointing for Report "&strReportName, "Report Re-point page"
	            Call ReportStep(StatusTypes.Information, "No success message after repointing for Report "&strReportName, "Report Re-point page")
	        '<shweta - 31/5/2016> - Changed report statements - end
        End If

        Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").Sync
        Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").RefreshObject
        
        For k = 1 To 20 Step 1
        	wait 2
        	Browser("Hosting Templates - Ops").Sync
            If Trim(UCase(Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebTable("tabConfirmation").GetCellData(1, 3))) = Trim(UCase("Success")) Then
                Reporter.ReportEvent micPass, "Report "&strReportName&" saved at "&repointMsg& " and repointed successfully", "Re-port Repoint Page"
                Call ReportStep(StatusTypes.Pass, "Report "&strReportName&" saved at "&repointMsg& " and repointed successfully", "Re-port Repoint Page")
                Exit for
            End If 
            
            If k = 20 Then
                Reporter.ReportEvent micInfo, "Report "&strReportName&" could not repoint successfully", "Report Re-point Page"
                Call ReportStep(StatusTypes.Information, "Report "&strReportName&" could not repoint successfully", "Report Re-point Page")
            End If
        Next
        
        Set objCloseOk = Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebButton("btnConfirmationClose")
        If objCloseOk.Exist(5) Then
            Call SCA.ClickOn(objCloseOk,"btnConfirmationOk" , "Report Repoint Page")
            Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").Sync
            Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").RefreshObject
        End If
        
        'TODO 
        'Based on the requirement:    Refresh the grid and capture the DataSorce/DataBase details for validation wrt new Exported file
        'rc = Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebTable("tabReport Repoint").GetROProperty(rows)    
    End If
    
    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").Sync
    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").RefreshObject
        
End Function


'Public Function ReportRepoint(ByVal strReportName, ByVal strActualDatabase, ByVal strActualCube, ByVal strOldDataSrcName, ByVal repoint, ByVal strNewDataSrcName, ByVal bookmarkValidation, ByVal aBookMark, ByVal radioDBtype, ByVal reportExists)
'    
'    Dim objReportName, WshShell, strDataBase, strCube, strDataSrcName, objDatabase, objNewDataSrcName, objbtnOk, strBookMarks, aBookMarkEdit, objReportLink, objCloseOk, compareVal
'    Dim i, j, k, l
'    aBookMarkEdit = ""
'    
'    
'    '<Shweta 2/2/2015> - Commented to handled sync issues - Start 
''    Set WshShell =CreateObject("WScript.Shell")
''	WshShell.SendKeys "{F5}"
''	wait 2
''	WshShell.SendKeys "{F5}"
''
''	Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
''	Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").RefreshObject
''	
''    set objHostingPage = Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops")
''    
''    Set objReportLink = Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Link("lnkReport")
''    Call SCA.ClickOn(objReportLink,"lnkReport" , "Hosting Templates")
''    wait 2
''    
''    Call IMSSCA.General.menuitemselect(objHostingPage,"Repoint")
'
'	'commented
'	Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
'	Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").RefreshObject
'    set objHostingPage = Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops")
'    Set objReportLink = Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Link("lnkReport")
'    Set objReportName = Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtReportName")
'    Set WshShell =CreateObject("WScript.Shell")
'
'	For a = 1 To 5 Step 1
'        WshShell.SendKeys "{F5}"
'        wait 2
'        WshShell.SendKeys "{F5}"
'        Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
'        Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").RefreshObject
'
'		Call SCA.ClickOn(objReportLink,"lnkReport" , "Hosting Templates")
'		wait 2
'		
'		'Call IMSSCA.General.menuitemselect(objHostingPage,"Repoint")
'		Call menuitemselect(objHostingPage,"Repoint")
'        Browser("Hosting Templates - Ops").Page("Clients - Ops System").Sync
'        Browser("Hosting Templates - Ops").Page("Clients - Ops System").RefreshObject
'        
'        objReportName.WaitProperty "visible",True,10000
'        If objReportName.Exist(1) Then
'            Call ReportStep(StatusTypes.Pass, "Successfully redirected to 'Report Repiont' Page in OPS", "OPS Report Repiont Page")
'            Exit For
'        End If
'        
'        If a=5 Then
'            Call ReportStep(StatusTypes.Fail, "Could not successfully redirected to 'Report Repiont' Page in OPS", "OPS Report Repiont Page")
'            Call ReportStep(StatusTypes.Warning, "Could not successfully redirected to 'Report Repiont' Page in OPS. Further functionality Validation steps may fail", "OPS Report Repiont Page")
'        End If
'    Next
'	
'	'Enter ReportName in grid tab and Validate DataSorce/Cube/DataBase details
'    '<Shweta 2/2/2015> - Commented to handle sync issues- End
'    
'    
'    '<Shweta - 25/11/2015 > Entering report name if its not null - Start
'    If strReportName <> "" Then
'    		'Using Shell object to enter
'		    Set WshShell = CreateObject("WScript.Shell")
'		    'Set objReportName = Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtReportName")
'		    Call SCA.SetText(objReportName,strReportName, "txtReportName","Report Repoint Page" )
'		    wait 2
'		    WshShell.SendKeys "{ENTER}"
'		    wait 1
'		    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtReportName").Click
'		    WshShell.SendKeys "{ENTER}"
'		    Set WshShell = Nothing
'		     	
'		 	'Shweta - 26/7/2016> - Start
'		 	Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtReportName").RefreshObject
'		 	'Shweta - 26/7/2016> - End
'    End If
'    '<Shweta - 25/11/2015 > Entering report name if its not null - End
'
'    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").Sync
'    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").RefreshObject
'    wait 2
'    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").Sync
'    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").RefreshObject
'            
'    For l = 1 To 10 Step 1
'    	Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").RefreshObject
'        If Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebTable("tabReport Repoint").Exist(1) Then
'            rc = Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebTable("tabReport Repoint").GetROProperty("rows")
'        else
'            Call ReportStep(StatusTypes.Fail, "Report Repoint grid table is not displaying", "Report Re-point page")            
'
'        End if
'        
'     	If rc > 2 Then
'	    	Set WshShell = CreateObject("WScript.Shell")
'	    	Set objtxtGridValueDatabase = Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtGridValueDatabase")
'		    Call SCA.SetText(objtxtGridValueDatabase,strActualDatabase, "txtGridValueDatabase","Report Repoint Page" )
'		    wait 2
'		    WshShell.SendKeys "{ENTER}"
'		    wait 1
'		    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtReportName").Click
'		    WshShell.SendKeys "{ENTER}"
'		    Set WshShell = Nothing
'		    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").Sync
'		    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").RefreshObject
'		    wait 1
'		    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").Sync
'    		Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").RefreshObject
'    		 	
'		 	'Shweta - 26/7/2016> - Start
'		 	Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtReportName").RefreshObject
'		 	'Shweta - 26/7/2016> - End
'    	End If
'    	
'    	rc = Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebTable("tabReport Repoint").GetROProperty("rows")
'    	
'    	If rc > 2 Then
'			Set WshShell = CreateObject("WScript.Shell")
'			Set objtxtGridValueCube = Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtCube")
'		    Call SCA.SetText(objtxtGridValueCube,strActualCube, "txtCube","Report Repoint Page" )
'		    wait 2
'		    WshShell.SendKeys "{ENTER}"
'		    wait 1
'		    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtReportName").Click
'		    WshShell.SendKeys "{ENTER}"
'		    Set WshShell = Nothing
'		    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").Sync
'		    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").RefreshObject
'		    wait 1
'		    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").Sync
'    		Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").RefreshObject
'    		
'    		'Shweta - 26/7/2016> - Start
'		 	Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtReportName").RefreshObject
'		 	'Shweta - 26/7/2016> - End
'   		End If
'    
'    	'Shweta - 26/7/2016> - Start
'	 	Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtReportName").RefreshObject
'	 	'Shweta - 26/7/2016> - End
'	 	
'        If reportExists = 0 Then
'            If rc >= 2 Then
'                Reporter.ReportEvent micPass, strReportName&" Report saved in SCA found and displaying in Report Repoint grid at "&rc&" row", "Report Re-point page"
'                Call ReportStep(StatusTypes.Pass, strReportName&" Report saved in SCA found and displaying in Report Repoint grid at "&rc&" row", "Report Re-point page")
'                ReportRepoint = "exists"
'                Exit For
'            End if
'            
'            If l =10 Then
'                'Reporter.ReportEvent micFail, strReportName&" Report saved in SCA is not displaying in Report Repoint grid. But Report is supposed to display in repoint grid", "Report Re-point page"
'                'Reporter.ReportEvent micFail, strReportName&" Report saved in SCA could not find in Report Repoint grid", "Report Re-point page"
'                ReportRepoint = "Doesnt exists"
'            End If
'        else
'            If rc = 1 Then
'                'Reporter.ReportEvent micPass, strReportName&" Report saved in SCA is not displaying in Report Repoint grid", "Report Re-point page"
'                ReportRepoint = "Doesnt exists"
'                Exit function
'            else
'                'Reporter.ReportEvent micFail, strReportName&" Report saved in SCA is displaying in Report Repoint grid. Report is not supposed to display in repoint grid but displaying at "&rc&" row", "Report Re-point page"
'                ReportRepoint = "exists"
'            End if
'            Exit For        
'        End If
'    Next
'    
'    'Shweta - 26/7/2016> - Start
' 	Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtReportName").RefreshObject
' 	'Shweta - 26/7/2016> - End
' 	
'	Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").Sync
'	Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").RefreshObject
'	
'	'shweta - 3/6/2016 - Start	
'	For i = 1 To 5 Step 1
'		Browser("Hosting Templates - Ops").Sync
'		wait 1
'		rowCount = Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebTable("tabTemplatename").RowCount
'		If rowCount = 2 Then
'			Exit For
'		End If
'	Next
'	'shweta - 3/6/2016 - End
'	
'	'Validate DataSorce/Cube/DataBase details
'    strDataBase=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebTable("tabTemplatename").GetCellData(2,8)
'    strCube=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebTable("tabTemplatename").GetCellData(2,9)
'    strDataSrcName=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebTable("tabTemplatename").GetCellData(2,11)
'    
'    wait 2
'    If Trim(Ucase(strDataBase)) = Trim(Ucase(strActualDatabase)) Then
'        Reporter.ReportEvent micPass, strActualDatabase&" database found in grid for Report "&strReportName, "Report Re-point page"
'        Call ReportStep(StatusTypes.Pass, strActualDatabase&" database found in grid for Report "&strReportName, "Report Re-point page")
'    Else
'    	Call ReportStep(StatusTypes.Information, strActualDatabase&" database with specific to its Hosting status found in grid for Report "&strReportName, "Report Re-point page")
''        Reporter.ReportEvent Infromation, strActualDatabase&" database could not find in grid for Report "&strReportName, "Report Re-point page"
''        Call ReportStep(StatusTypes.Fail, strActualDatabase&" database could not find in grid for Report "&strReportName, "Report Re-point page")
''        Call ReportStep(StatusTypes.Fail, "Mismatch between expected Database:= "&strActualDatabase&" and Application displaying database:= "&strDataBase, "Report Re-point page")
'    End If
'    
'    wait 2
'    If Trim(Ucase(strCube)) = Trim(Ucase(strActualCube)) Then
'        Reporter.ReportEvent micPass, strActualCube&" cube found in grid for Report "&strReportName, "Report Re-point page"
'        Call ReportStep(StatusTypes.Pass, strActualCube&" cube found in grid for Report "&strReportName, "Report Re-point page")
'    Else
'        Reporter.ReportEvent micFail, strActualCube&" cube could not find in grid for Report "&strReportName, "Report Re-point page"
'        Call ReportStep(StatusTypes.Fail, strActualCube&" cube could not find in grid for Report "&strReportName, "Report Re-point page")
'        Call ReportStep(StatusTypes.Fail, "Mismatch between expected Cube:= "&strActualCube&" and Application displaying Cube:= "&strCube, "Report Re-point page")
'    End If
'        
'    wait 2
'   
''Shweta 3/5/2016 Due to change in application Grid display of DataSource appended "*" in strOldDataSrcName based on application demand - Start 
'   If instr(1, strDataSrcName, "*")>0 Then
'		If instr(1, strOldDataSrcName, "*")<=0 Then
'			strOldDataSrcName = strOldDataSrcName&"*"
'		End If
'	End If
''Shweta 3/5/2016 Due to change in application Grid display of DataSource appended "*" in strOldDataSrcName based on application demand - End	
'
'    If Trim(Ucase(strDataSrcName)) = Trim(Ucase(strOldDataSrcName)) Then
'        Reporter.ReportEvent micPass, strOldDataSrcName&" datasource found in grid for Report "&strReportName, "Report Re-point page"
'        Call ReportStep(StatusTypes.Pass, strOldDataSrcName&" datasource found in grid for Report "&strReportName, "Report Re-point page")
'    Else
'        Reporter.ReportEvent micFail, strOldDataSrcName&" datasource could not find in grid for Report "&strReportName, "Report Re-point page"
'        Call ReportStep(StatusTypes.Fail, strOldDataSrcName&" datasource could not find in grid for Report "&strReportName, "Report Re-point page")
'        Call ReportStep(StatusTypes.Fail, "Mismatch between expected Datasource:= "&strOldDataSrcName&" and Application displaying Datasource:= "&strDataSrcName, "Report Re-point page")
'    End If
'
'    
'    If bookmarkValidation = 0 Then
'        strBookMarks=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebTable("tabTemplatename").GetCellData(2,6)
'        For j = 0 To UBound(aBookMark) Step 1
'            If j <> UBound(aBookMark) Then
'                aBookMarkEdit = aBookMarkEdit&aBookMark(j)&", "
'            Else
'                aBookMarkEdit = aBookMarkEdit&aBookMark(j)
'            End If
'        Next
'        
'        '<Shweta 15/9/2015> - Start
''        If Trim(Ucase(strBookMarks)) = Trim(Ucase(aBookMarkEdit)) Then
''            Reporter.ReportEvent micPass, strBookMarks&" BookMarks found in grid for Report "&strReportName, "Report Re-point page"
''        Else
''            Reporter.ReportEvent micFail, strBookMarks&" BookMarks could not find in grid for Report "&strReportName, "Report Re-point page"
''        End If
'
'		compareVal = InStr(1, Trim(Ucase(strBookMarks)), Trim(Ucase(aBookMarkEdit)))
'	    If compareVal >= 1 Then
'	        Reporter.ReportEvent micPass, strBookMarks&" BookMark/BookMarks found in grid for Report "&strReportName, "Report Re-point page"
'	         Call ReportStep(StatusTypes.Pass, strBookMarks&" BookMark/BookMarks found in grid for Report "&strReportName, "Report Re-point page")
'	    Else
'	    	Reporter.ReportEvent micFail, strBookMarks&" BookMark/BookMarks could not find in grid for Report "&strReportName, "Report Re-point page"
'	    	Call ReportStep(StatusTypes.Fail, strBookMarks&" BookMark/BookMarks could not find in grid for Report "&strReportName, "Report Re-point page")
'	    End If
'	    '<Shweta 15/9/2015> - End
'	    
'    End If
'    
'    
'    'Enter details in Destination move and provide Environment, Confirm the report repoint
'    If repoint = 0 Then
'        rc = Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebTable("tabReport Repoint").GetROProperty("rows")
'        If rc = 2 Then
'            Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebTable("tabTemplatename").ChildItem(2, 1, "WebCheckBox", 0).click    
'        '<Shweta - 25/11/2015 > If report name value is null - Start
'        ElseIf rc >= 2 Then
'        	Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebCheckBox("chkReportRepointSelection").Set "ON"
'        '<Shweta - 25/11/2015 > If report name value is null - End
'        End If
'        
'        '<shweta 24/8/2015> - Start
'        '        Set objNewDataSrcName = Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtDestinationDataSource")
'        '        Call SCA.SetText(objNewDataSrcName,strNewDataSrcName, "txtDestinationDataSource","Report Repoint Page")
'        '        wait 2
'        '        
'        '        Set objDatabase = Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtDestinationDatabase")
'        '        Call SCA.SetText(objDatabase,strDatabase, "txtDestinationDatabase","Report Repoint Page")
'        
'        Set WshShell = CreateObject("WScript.Shell")
'        Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtDestinationDataSource").Click
'        WshShell.SendKeys strNewDataSrcName
'        wait 2 ' To wait for List to display as per keyword entered above
'        WshShell.SendKeys "{DOWN}" 
'        wait 2 'To select item from list and wait
'        WshShell.SendKeys "{ENTER}"
'        'WshShell.SendKeys "{TAB}"
'        WshShell.SendKeys "{DOWN}" 
'        wait 2 'To select item from list and wait
'        WshShell.SendKeys "{ENTER}"
'        
'        wait 2
'        
''        'Shweta - 26/7/2016> - Start
''        'Handled str splitting to seperate "_QA" and  "Prior" from strActualDatabase 
''        'Eg: strActualDatabase = "OPSINTSCA_IN_W_1_35162922_QA" | "OPSINTSCA_IN_W_1_35162922_Prior" | "OPSINTSCA_IN_W_1_35162922"
''        Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtDestinationDatabase").RefreshObject
''         If instr(1, Trim(UCase(strActualDatabase)), Trim(UCase("_QA")))> 0 then
''         	wait 2
''		 	arrTempId = split(Ucase(strActualDatabase), Ucase("_QA"))
''		 	strActualDatabase = arrTempId(0)
''		 ElseIf instr(1, Trim(UCase(strActualDatabase)), Trim(UCase("_Prior")))> 0 then
''		 	arrTempId = split(Ucase(strActualDatabase), UCase("_Prior"))
''		 	strActualDatabase = arrTempId(0)
''		 Else
''		 	'Dim arrTempId(0)
''		 	'arrTempId(0) = strActualDatabase
''		 End if
''		 wait 2
''		Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").Sync 
''		 Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtDestinationDatabase").RefreshObject
''        'Shweta - 26/7/2016> - End
'        
'        Set WshShell1 = CreateObject("WScript.Shell")
'        Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebEdit("txtDestinationDatabase").Click
'        WshShell1.SendKeys strActualDatabase
'        wait 2 ' To wait for List to display as per keyword entered above
'        WshShell1.SendKeys "{DOWN}" 
'        wait 2 'To select item from list and wait
'        WshShell1.SendKeys "{ENTER}"
'        Set  WshShell1 = Nothing
'        wait 2
'        '<shweta 24/8/2015> - End
'        
'        Set objradioDBtype = Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebRadioGroup("rdoDBType")
'        Call SCA.ClickOnRadio(objradioDBtype,radioDBtype, "Report Repoint Page")
'        
'        Set objbtnRepoint=Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebButton("btnRepoint Selected Reports")
'        Call SCA.ClickOn(objbtnRepoint,"btnRepoint Selected Reports" , "Report Repoint Page")
'        
'        'Confirm the report repoint
'        For j = 1 To 20 Step 1
'            If Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebElement("welReportRepointMsg").Exist(2) Then
'                repointMsg = Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebTable("tabConfirmation").GetCellData(1, 2)
'                
'                Set objbtnOk = Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebButton("btnConfirmationOk")
'                Call SCA.ClickOn(objbtnOk,"btnConfirmationOk" , "Report Repoint Page")
'                
'                wait 2
'                Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").Sync
'                Exit for
'            End If
'            
'            If j=20 Then
'            '<shweta - 31/5/2016> - Changed report statements - start
''                Reporter.ReportEvent micFail, "Report "&strReportName&" could not repoint successfully", "Report Re-point page"
''                Call ReportStep(StatusTypes.Fail, "Report "&strReportName&" could not repoint successfully", "Report Re-point page")
'                Reporter.ReportEvent micInfo, "No success message after repointing for Report "&strReportName, "Report Re-point page"
'                Call ReportStep(StatusTypes.Information, "No success message after repointing for Report "&strReportName, "Report Re-point page")
'            '<shweta - 31/5/2016> - Changed report statements - end
'            End If
'        Next
'        
'        Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").Sync
'        Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").RefreshObject
'        
'        For k = 1 To 20 Step 1
'        	wait 2
'        	Browser("Hosting Templates - Ops").Sync
'            If Trim(UCase(Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebTable("tabConfirmation").GetCellData(1, 3))) = Trim(UCase("Success")) Then
'                Reporter.ReportEvent micPass, "Report "&strReportName&" saved at "&repointMsg& " and repointed successfully", "Re-port Repoint Page"
'                Call ReportStep(StatusTypes.Pass, "Report "&strReportName&" saved at "&repointMsg& " and repointed successfully", "Re-port Repoint Page")
'                Exit for
'            End If 
'            
'            If k = 20 Then
'                Reporter.ReportEvent micInfo, "Report "&strReportName&" could not repoint successfully", "Report Re-point Page"
'                Call ReportStep(StatusTypes.Information, "Report "&strReportName&" could not repoint successfully", "Report Re-point Page")
'            End If
'        Next
'        
'        Set objCloseOk = Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebButton("btnConfirmationClose")
'        If objCloseOk.Exist(5) Then
'            Call SCA.ClickOn(objCloseOk,"btnConfirmationOk" , "Report Repoint Page")
'            Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").Sync
'            Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").RefreshObject
'        End If
'        
'        'TODO 
'        'Based on the requirement:    Refresh the grid and capture the DataSorce/DataBase details for validation wrt new Exported file
'        'rc = Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").WebTable("tabReport Repoint").GetROProperty(rows)    
'    End If
'    
'    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").Sync
'    Browser("Hosting Templates - Ops").Page("Report Repoint - Ops System").RefreshObject
'        
'End Function
'

'===========================================================================================================================================================   
    
    '    '<Opens Excel, reads and edits selected fields in Excel sheet>
    '    '<StrOper :- Excel(read/write) operation to be perfomed>
    '    '<excelName :- Mention Excel File Name>
    '    '<excelSheetName :- Mention Excel sheetname within file>
    '	 '<aRow:- Excel Row Val>
    '	 '<aCol:- Excel Col Val>
    '	 '<aVal:- mention Value to set at Excel Row Val(aRow), Excel Col Val(aCol)>
    '    '<author :-Shweta Nagaral>'
'===========================================================================================================================================================
Public Function ExcelOperation(BYVal StrOper, ByVal excelName, ByVal excelSheetName, ByVal aRow, ByVal aCol, ByVal aVal)
	
	Dim objReportName, WshShell, strDataBase, strCube, strDataSrcName, objDatabase, objNewDataSrcName, objbtnOk
	
	Dim aCellVal() 'Declaring empty array
	
	SystemUtil.CloseProcessByName "excel.exe"
	
	'Open 'SCAReportSheet.xls' Excel and write attribute value in SheetName mentioned in "" - Start
    Set objExcel = CreateObject("Excel.Application")
    Set objWorkbook = objExcel.Workbooks.Open(Environment.Value("CurrDir")&excelName)
   
    objExcel.Application.Visible = False
	
	Select Case StrOper
		Case "Read"
					If UBound(aRow) = UBound(aCol) Then
							ReDim aCellVal(UBound(aRow))
							
							For i = 0 To UBound(aRow) Step 1
								aValue=objWorkbook.Worksheets(excelSheetName).cells(aRow(i),aCol(i))
								aCellVal(i) = aValue
							Next
					Else
							Call ReportStep (StatusTypes.Fail, "Row and Col count are different. Check 'ExcelOperation' Funtion input args","Excel File Reading Opertaion")
							Call ReportStep (StatusTypes.Fail, "Could not fetch values from "&excelName&" inside "&excelSheetName&" Sheet","Excel File Reading Opertaion")
					End If
					
					If UBound(aCol) = UBound(aCellVal) Then
						ExcelOperation = aCellVal
					Else
						Call ReportStep (StatusTypes.Fail, "Could not fetch all values from "&excelName&" inside "&excelSheetName&" Sheet","Excel File Reading Opertaion")
					End If
					
		
		Case "Write"
					If UBound(aRow) = UBound(aCol) and UBound(aCol) = UBound(aVal) Then
							For i = 0 To UBound(aRow) Step 1
							if aVal(i)<>""  Then 
								objWorkbook.sheets(excelSheetName).cells(aRow(i),aCol(i))=aVal(i)
							End if 
						Next
					Else
							Call ReportStep (StatusTypes.Fail, "Row and Col count are different. Check 'ExcelOperation' Funtion input args","Excel File writing Opertaion")
							Call ReportStep (StatusTypes.Fail, "Could not save values in "&excelName&" inside "&excelSheetName&" Sheet","Excel File writing Opertaion")
					End If
		
	End Select	
	
	'Saves workbook and closes it
    objWorkbook.Save
    objExcel.ActiveWorkbook.Close
    objExcel.Application.Quit
    Set objExcel=Nothing
    'Open 'SCAReportS
	
End Function


'
'===========================================================================================================================================================   
    
    '    '<Checks hosting status of the cube templates>
    '    '<cubeTempName :- Cube Template Name>
    '    '<author :-Shweta Nagaral>'
'===========================================================================================================================================================
Public Function CheckHostingTemplateStatus(ByVal cubeTempName)
	Dim WshShell, rowCount, i, cubeName, hostStatus, qaStatus, statusFound, aReturnVal, appobject
	
	'To Redirect from QA Acceptance page - Start
	If Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Image("imghead_logo").Exist(2) Then
		Set imghead_logo = Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Image("imghead_logo")
		Call SCA.ClickOn(imghead_logo,"imghead_logo","Report Creation")
		Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
		Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").RefreshObject
	End If
	'To Redirect from QA Acceptance page - End
	
	Set appobject=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops")	
	Call IMSSCA.General.menuitemselect(appobject,"Templates")
	
	Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
	wait 2
	Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
	
	'Search for Cube Template ID in Hosting Template section
	Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebEdit("txtTemplateIdTab").Set cubeTempName
	
	'Using Shell object to filter template id and  to list the same template id in Hosting Templates List
	Set WshShell = CreateObject("WScript.Shell")
	wait 2
	WshShell.SendKeys "{ENTER}"
	Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebEdit("txtTemplateIdTab").Click
	WshShell.SendKeys "{ENTER}"
	wait 2
	Set WshShell = Nothing
	
	
	rowCount = Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebTable("tabTemplatename").GetROProperty("rows")
	'msgbox rowCount
	
	If rowCount  > 1  Then
	
	        For i = 1 to rowCount
	                cubeName = Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebTable("tabTemplatename").GetCellData(i, 1)
	                If  cubeName = cubeTempName  Then
	                    hostStatus = Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebTable("tabTemplatename").GetCellData(i, 4)
	                    qaStatus = Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").WebTable("tabTemplatename").GetCellData(i, 5)
	                    Call ReportStep (StatusTypes.Pass, "Validation of  hosting template status in Hosting Section. Cube template " &cubeTempName& " is listing in Hosting Template section", "OPS Page")
	                    Call ReportStep (StatusTypes.Pass, cubeTempName& " Hosting status is " &hostStatus& " and QAStatus is " &qaStatus, "OPS Page")
	                    statusFound = 1
	                End If
	        Next
	
	End If
	
	If  statusFound = 0 Then
	        Call ReportStep (StatusTypes.Fail,  "Validation of  hosting template status in Hosting Section. Cube template " &cubeTempName& " is not listing, Please create cube template again", "OPS Page")
	        hostStatus = "Not Defined"
	        qaStatus = "Not Defined"
	End If
	
	aReturnVal = Array(hostStatus, qaStatus)
	CheckHostingTemplateStatus = aReturnVal
	
End Function






'""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
	'< Operation on the different Folder Structure>
	'< strOselection: to select the operation>
    '< strFolderName : Foldername>
	'< strLocation: location to perform the operations>
	'< strFolderName_Update:Folder name to update the Folder name>	
	'<Author>Shobha<Author>	
'""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

 Public Function ReportPublish_Ops(ByVal strOSelection,ByVal strClientName,ByVal strReportName,ByVal strCountry,ByVal stroffering)
 	
   Dim arr,arrele,i,objclick,WshShell,strSerach
   
   '<Object Declaration>
   Dim objpage,objlnkReport, SynReportPublish , objlstCountry , objlstOffering ,objbtnBrowse , objSCA , objbtnOKS
   Dim objNonSyn,onjCountry,objBrowserbtn,objwebSCAAutomation,objbtnOKk, lstC,lstCountry , objtxtReportName, objbtnPublish, objPublish 
   Dim objtxtReportNamecl
 
   wait 3
   Set WshShell = CreateObject("WScript.Shell")
   
   Set objlnkReport=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").Link("lnkReport")
   Call SCA.ClickOn(objlnkReport,"lnkReport", "Report Publish")
	 wait 1
  Set objReportPublish=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").Link("lnkPublish")
  Call SCA.ClickOn(objReportPublish,"lnk Publish", "Report Publish")
 
	 
 
'   Set objpage=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops")  
'   Call IMSSCA.General.menuitemselect(objpage,"Publish")	
	
	
	
   Select Case strOSelection
 
   	 Case "SyndicateReport"
   	 
   	 	Set objReporttype=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebRadioGroup("ReportType") 	
		Call SCA.ClickOnRadio(objReporttype,"SynReportPublish" ,"SynReportPublish")
		
   	 	wait 1
   	 	Set objlstCountry=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebList("lstCountry")
   	 	Call SCA.SelectFromDropdown(objlstCountry,strCountry)
   	 	wait 2
		Set objlstOffering=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebList("lstSrc_Offerings")
		Call SCA.SelectFromDropdown(objlstOffering,stroffering)
		wait 1
		Set objbtnBrowse=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebButton("btnBrowse")	
		Call SCA.ClickOn(objbtnBrowse,"btn Browse", "Report Publish")
		
		Set objSCA=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebElement("webSCA_AutomationReport")
		Call SCA.ClickOn(objSCA,"Automation Report Folder Selection", "Report Publish")
		
		Set objbtnOKS=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebElement("btnOK")
		Call SCA.ClickOn(objbtnOKS,"Btn OK", "Report Publish")
		
			
		wait 1		
		arr=Split(strClientName,";")		
		arrele=ubound(arr)
		strSerach=arr(0)
		Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebEdit("txtSearchWebEdit").Set Left(strSerach,10)
		 For i = 0 To arrele Step 1
		  Set objchk=Description.Create
		  objchk("micclass").Value="WebCheckBox"
		  objchk("type").Value="checkbox"
		  objchk("title").Value=arr(i)
		
		  wait 1
		  Set objchkC= Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebElement("webFilter:Check allUncheck").ChildObjects(objchk)
		  wait 1
		  objchkC(0).click  		 	
		 Next
		 
		wait 1
		
		Set objtxtReportName=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebEdit("txtReportName")		
		Call SCA.SetText(objtxtReportName,strReportName, "ReportName","Report Publish" )
		
		Set objtxtReportNamecl=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebEdit("txtReportName")
		Call SCA.ClickOn(objtxtReportNamecl,"txtReportName", "Report Publish")
		
		Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").Sync
		Set WshShell = CreateObject("WScript.Shell")
		WshShell.SendKeys "{ENTER}"
		wait 3
		Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").Sync		
        Set objclick=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebTable("tabPublish").ChildItem(2,1,"WebCheckBox",0)
		objclick.click
		
		Set objbtnPublish=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebButton("btnPublish")
		Call SCA.ClickOn(objbtnPublish,"btn Report Publish", "Report Publish")
		wait 1
		Set objPublish=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebElement("btnReportPublishOk")
		Call SCA.ClickOn(objPublish,"Report Publish Ok btn", "Report Publish")
		
		
        'Capture the Msg here
   	 Case "NonsyndicateReport"
   	 
   	 	Set objReporttype=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebRadioGroup("ReportType")   	 	
		Call SCA.ClickOnRadio(objReporttype,"NonSynReportPublish" ,"SynReportPublish")	
   	 	
   	 	Set onjCountry=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebList("lstCountry")
   	 	Call SCA.SelectFromDropdown(onjCountry,strCountry)
   	 	
		Set objBrowserbtn=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebButton("btnBrowse")
		Call SCA.ClickOn(objBrowserbtn,"btn Browse", "Report Publish")
		
		Set objwebSCAAutomation=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebElement("webSCA_AutomationReport")
		Call SCA.ClickOn(objwebSCAAutomation,"SCAAutomation", "Report Publish")
		
		Set objbtnOKk=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebElement("btnOK")
		Call SCA.ClickOn(objbtnOKk,"btn OK", "Report Publish")
	
		Set lstC=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebList("lstClient")
		Call SCA.SelectFromDropdown(lstC,strClientName)
		wait 1
		
		Set lstCountry=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebList("lstCountry")
		Call SCA.SelectFromDropdown(lstCountry,strCountry)
 		wait 1
 		
		Set objlstOffering=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebList("lstOfferings")
		Call SCA.SelectFromDropdown(objlstOffering,stroffering)
		wait 1
		
		rc=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebTable("tabPublish").RowCount
		If rc=0 Then
			
		Set objBrowserbtn=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebButton("btnBrowse")
		Call SCA.ClickOn(objBrowserbtn,"btn Browse", "Report Publish")
		
		Set objwebSCAAutomation=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebElement("webSCA_AutomationReport")
		Call SCA.ClickOn(objwebSCAAutomation,"SCAAutomation", "Report Publish")
		
		Set objbtnOKk=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebElement("btnOK")
		Call SCA.ClickOn(objbtnOKk,"btn OK", "Report Publish")
	

		End If
		
		
		Set objtxtReportName=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebEdit("txtReportName")		
		Call SCA.SetText(objtxtReportName,strReportName, "ReportName","Report Publish" )
		
		Set objtxtReportNamecl=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebEdit("txtReportName")
		Call SCA.ClickOn(objtxtReportNamecl,"txtReportName", "Report Publish")	
		
		
		wait 1		
		Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").Sync
		Set WshShell = CreateObject("WScript.Shell")
		WshShell.SendKeys "{ENTER}"
		wait 1
		Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").Sync
		Set objclick=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebTable("tabPublish").ChildItem(2,1,"WebCheckBox",0)
		objclick.click		
		
		Set objbtnPublish=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebButton("btnPublish")
		Call SCA.ClickOn(objbtnPublish,"btn Report Publish", "Report Publish")
		wait 1
		Set objPublish=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebElement("btnReportPublishOk")
		Call SCA.ClickOn(objPublish,"Report Publish Ok btn", "Report Publish")		

		
   	 Case "Republish"
   	 	
   	 	Set objNonSyn=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebRadioGroup("ReportType")   	 	
   	 	Call SCA.ClickOnRadio(objNonSyn,"ReportRepublish" ,"SynReportPublish")	
   	 	
   	 	Set onjCountry=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebList("lstCountry")
   	 	Call SCA.SelectFromDropdown(onjCountry,strCountry)
   	 	
		Set objBrowserbtn=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebButton("btnBrowse")
		Call SCA.ClickOn(objBrowserbtn,"btn Browse", "Report Publish")
		
		Set objwebSCAAutomation=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebElement("webSCA_AutomationReport")
		Call SCA.ClickOn(objwebSCAAutomation,"SCAAutomation", "Report Publish")
		
		Set objbtnOKk=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebElement("btnOK")
		Call SCA.ClickOn(objbtnOKk,"btn OK", "Report Publish")
	
		Set lstC=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebList("lstClient")
		Call SCA.SelectFromDropdown(lstC,strClientName)
		wait 1
		
		Set lstCountry=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebList("lstCountry")
		Call SCA.SelectFromDropdown(lstCountry,strCountry)
 		wait 1
 		
		Set objlstOffering=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebList("lstOfferings")
		Call SCA.SelectFromDropdown(objlstOffering,stroffering)
		wait 1
		
		rc=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebTable("tabPublish").RowCount
		If rc=0 Then
			
		Set objBrowserbtn=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebButton("btnBrowse")
		Call SCA.ClickOn(objBrowserbtn,"btn Browse", "Report Publish")
		
		Set objwebSCAAutomation=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebElement("webSCA_AutomationReport")
		Call SCA.ClickOn(objwebSCAAutomation,"SCAAutomation", "Report Publish")
		
		Set objbtnOKk=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebElement("btnOK")
		Call SCA.ClickOn(objbtnOKk,"btn OK", "Report Publish")
	

		End If
		
		
		Set objtxtReportName=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebEdit("txtrepublishReportName")		
		Call SCA.SetText(objtxtReportName,strReportName, "ReportName","Report Publish" )
		
		Set objtxtReportNamecl=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebEdit("txtrepublishReportName")
		Call SCA.ClickOn(objtxtReportNamecl,"txtReportName", "Report Publish")	
		
		
		wait 1		
		Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").Sync
	    Set WshShell = CreateObject("WScript.Shell")
		WshShell.SendKeys "{ENTER}"
		wait 1
		Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").Sync
		Set objclick=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebTable("tabPublish").ChildItem(2,1,"WebCheckBox",0)
		objclick.click		
		
		Set objbtnPublish=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebButton("btnPublish")
		Call SCA.ClickOn(objbtnPublish,"btn Report Publish", "Report Publish")
		wait 1
		Set objPublish=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebElement("btnReportPublishOk")
		Call SCA.ClickOn(objPublish,"Report Publish Ok btn", "Report Publish")		

   	
   End Select
 	wait 4
 	If Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebElement("webReportStatus").Exist(5) Then
 	
 		strReportstatus=Browser("Hosting Templates - Ops").Page("Publish Reports - Ops").WebElement("webReportStatus").GetROProperty("innertext")
		ReportPublish_Ops=strReportstatus
 		
 	End If
 	 	

 Set WshShell = Nothing	
 	
 End Function
 
 
 
 
'""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
	'<Report OwnerShip>
	'<strOselection: to select the operation>
    '<strFolderName : Foldername>
	'<	strLocation: location to perform the operations>
	'< strFolderName_Update:Folder name to update the Folder name>	
	'<Author>Shobha<Author>	
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""" 
 
 

 Function ReportOwnerShip_InOps(ByVal strOSelection,ByVal strReportName,ByVal strEReportOwner,ByVal strNReportOwner,ByVal strClientName,ByVal strBookmarkN)
 
 Dim Objwshell,objpage,objReDescriptionB,objReDescription,objCountR,objCountB,intCountR,intCountB
	
	wait 2
	Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Link("lnkReport").Click
	
	
	wait 2
	Browser("Hosting Templates - Ops").Page("Report Ownership Change").Link("lnkOwnership").Click

	

	
'	Call IMSSCA.General.menuitemselect(objpage,"Ownership")
'	Browser("Hosting Templates - Ops").Page("Report Ownership Change").Sync

	Select Case strOSelection
		
		Case "Report"
		  Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebRadioGroup("chkROType").Select "Report"
		Case "Bookmark"
		 Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebRadioGroup("chkROType").Select "Bookmark"
		Case "Folder"
		 Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebRadioGroup("chkROType").Select "Folder"
	End Select
	Browser("Hosting Templates - Ops").Page("Report Ownership Change").Sync
	Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebList("lstClientselect").Select strClientName
	
	Browser("Hosting Templates - Ops").Page("Report Ownership Change").Sync
	
	If strOSelection="Report" OR strOSelection="Bookmark" Then
		 Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebEdit("txtExistingOwner").Set strEReportOwner
		 Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebButton("btnSearchE").Click
		 Browser("Hosting Templates - Ops").Page("Report Ownership Change").Sync 
	else
		 Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebButton("btnFolderBrowse").Click
		 wait 2
		 If Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebElement("webA2Folder").Exist(3) then
			Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebElement("webA2Folder").Click
		
		 End If
		 Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebButton("btnBrowseOK").Click	
		 wait 1
		'		For k = 0 To 20 Step 1	
		'		Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebElement("webMove").Click
		'		Set WshShell =CreateObject("WScript.Shell")
		'		WshShell.SendKeys "{DOWN}"
		'		wait 1
		'		If Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebElement("webA2Folder").Exist(3) then
		'		Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebElement("webA2Folder").Click
		'		Exit For
		'		End If
		'		Next
		'		
		'		For k = 0 To 50 Step 1	
		'		Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebElement("webMove").Click
		'		Set WshShell =CreateObject("WScript.Shell")
		'		WshShell.SendKeys "{DOWN}"
		'		wait 1
		'		If Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebElement("WebIndia").Exist(3) then
		'		Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebElement("WebIndia").Click
		'		Exit For
		'		End If
		'		Next
		'		
		'		For k = 0 To 30 Step 1	
		'		Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebElement("webMove").Click
		'		Set WshShell =CreateObject("WScript.Shell")
		'		WshShell.SendKeys "{DOWN}"
		'		wait 1
		'		If Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebElement("WebSCAAutomation").Exist(3) then
		'		Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebElement("WebSCAAutomation").Click
		'		Exit For
		'		End If
		'		Next
		'	 Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebButton("btnBrowseOK").Click	
	End If

'shweta - 3/5/2016 - Start
'	Browser("Hosting Templates - Ops").Page("Report Ownership Change").Sync
'	Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebEdit("txtPath").Click
'	Browser("Hosting Templates - Ops").Page("Report Ownership Change").Sync
'	
'	Set Objwshell = CreateObject("WScript.Shell")
'	Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebEdit("txtPath").Set strReportName
'	Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebEdit("txtPath").Click
'	wait 1
'	Objwshell.SendKeys "{ENTER}"
'	wait 2
	Set Objwshell = CreateObject("WScript.Shell")
	intLoopCounter = Environment.Value("intCounterStart")
	
	Do
		If intLoopCounter=Environment.Value("intCounterMaxLimit") Then
			Exit Do 
	    End If
	
	 	intLoopCounter=intLoopCounter+1
	 	
		Browser("Hosting Templates - Ops").Page("Report Ownership Change").Sync
		Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebEdit("txtPath").Click
		Browser("Hosting Templates - Ops").Page("Report Ownership Change").Sync
		
		Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebEdit("txtPath").Set strReportName
		Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebEdit("txtPath").Click
		wait 1
		Objwshell.SendKeys "{ENTER}"
		wait 2
		
	Loop Until Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebTable("tabReportOwnership").Exist(5) = True
'shweta - 3/5/2016 - End

'	If Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebEdit("txtBookmarkName").Exist(0)  Then
'		Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebEdit("txtBookmarkName").Click
'		Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebEdit("txtBookmarkName").Set strBookmarkN
'		Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebEdit("txtBookmarkName").Click
'		wait 1
'		Objwshell.SendKeys "{ENTER}"
'		wait 2
'		
'	End If
'	
	set objReportName= Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebTable("tabReportOwnership").ChildItem(2,1,"WebCheckBox",0)
	objReportName.click
	
	Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebEdit("txtNewOwner").Set strNReportOwner
	Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebButton("btnSerachN").Click
	Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebButton("btnUpdate").Click
	
	
	
	Set objReDescription=Description.Create
	objReDescription("micclass").value="WebElement"
	objReDescription("html tag").value="LI"
	objReDescription("innertext").Regularexpression=True
	objReDescription("innertext").Value=".*"&strReportName
	
	
	Set objReDescriptionB=Description.Create
	objReDescriptionB("micclass").value="WebElement"
	objReDescriptionB("html tag").value="LI"
	objReDescriptionB("innertext").Regularexpression=True
	objReDescriptionB("innertext").Value=".*"&strBookmarkN
	
	
	Set objCountR=Browser("Hosting Templates - Ops").Page("Report Ownership Change").ChildObjects(objReDescription)
	intCountR=objCountR.count
	Set objCountB=Browser("Hosting Templates - Ops").Page("Report Ownership Change").ChildObjects(objReDescriptionB)
	intCountB= objCountB.count	
	
	
	If intCountR=1 AND intCountB<>1 Then	
	 Call ReportStep (StatusTypes.Pass,strReportName&space(2)&"Report Selection in Report Ownesrhip:-"&space(2)&"Selected Successfully", "Report Ownership Page")
	ElseIf intCountB=1 Then
     Call ReportStep (StatusTypes.Pass,strBookmarkN&space(2)&"BookMark Selection in Report Ownesrhip:-"&space(2)&"Selected Successfully", "Report Ownership Page")	
   	Else	
	 Call ReportStep (StatusTypes.Fail,strReportName&space(2)&"Report Selection in Report Ownesrhip:-"&space(2)&"Not Selected Successfully", "Report Ownership Page")	
	End If				 
	
  	
		Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebButton("btnBrowseOK").Click
		Browser("Hosting Templates - Ops").Page("Report Ownership Change").Sync
		
		If Browser("Hosting Templates - Ops").Page("Report Ownership Change").WebElement("webOwnership change successfully").Exist(100) Then
 		Call ReportStep (StatusTypes.Pass,"Report Owner for the"&strReportName&"is:-"&strEReportOwner, "Report Ownership Page")
		else		
		Call ReportStep (StatusTypes.Fail,"Report Owner for the"&strReportName&":-"&strEReportOwner, "Report Ownership Page")

		End If
	
 End Function


'===========================================================================================================================================================   
    
    '    '<ReportImportExport perfomring Exporting/Importing of SCA reports saved inside Shared Reports/ My Reports/ Shared Data Sources >
    '    '<reportOperation :- Mention type of operation i.e. exporting/importin>
    '    '<sheetNum :- Mention Excel sheetname to open>
    '	 '<strXmlOper:- operation on XML such as Open Xml/Save Xml/openLocalSavedXml>
    '	 '<reportOwner:- Select Report owner value to list all reports listing under specific user>
    '    '<author :-Shweta Nagaral>'
'===========================================================================================================================================================

Public Function ReportImportExport(byVal reportOperation, byVal sheetNum, byVal strXmlOper, byVal reportOwner)
    
    Dim objExcel, ObjExcelFile, ObjExcelSheet, intRow_Count, intColumn_Count, strChild_Values, obj_tree, intcount, objtreeCount, objwebele, objchwebelm, desc, chcnt, objlnkReportImportExport
    Dim j, a, k
    Dim strChild_ValRow2,strPreviousRow
    
    Browser("Analyzer").Page("Shared_Folder").Sync
    Browser("Analyzer").Page("Shared_Folder").RefreshObject
    
    'Click on ReportImportExport                
    Set objlnkReportImportExport = Browser("Analyzer").Page("Home").Frame("Frame").Link("lnkReportImportExport")
    Call SCA.ClickOn(objlnkReportImportExport,"lnkReportImportExport", "SCA Home Page")
    
    Browser("Analyzer").Page("Shared_Folder").Sync
    Browser("Analyzer").Page("Shared_Folder").RefreshObject
    
    Select Case ReportOperation
        
        Case "ReportImport"
        
        Case "ReportExport"
                            'Set Report Owner
                            Browser("Analyzer").Page("Home").Frame("Report Export Import Frame").WebList("lstReportOwner").Select reportOwner                            
                            '<shweta 16/6/2015 Closing All opened Excel Files- start
'                            objUtils.KillProcess("excel.exe")
                            '<shweta 16/6/2015 Closing All opened Excel Files- end
                            
                            Set objExcel = CreateObject("Excel.Application") 
                            objExcel.Visible =False     
                            
                            'Shweta <30/11/2015> - Start
                            'Set ObjExcelFile = objExcel.Workbooks.Open(Environment.Value("CurrDir")&"InputFiles\IMSSCAWeb\SCAReportSheet.xls")
                            'Shweta <30/11/2015> - End
                            Set ObjExcelFile = objExcel.Workbooks.Open(Environment.Value("CurrDir")&Environment.Value("ReportCreationFile"))
                            Set ObjExcelSheet = ObjExcelFile.Sheets(sheetNum) 
                            intRow_Count = ObjExcelSheet.UsedRange.Rows.Count 
                            intColumn_Count = ObjExcelSheet.UsedRange.Columns.Count
                            
                            'For i=2 to intRow_Count
                            'Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcube").Click
                            For i=2 to intRow_Count                            
                                                           
                               For j=1 to intColumn_Count-1         
                                                        
                                   strChild_Values=Trim(ObjExcelSheet.cells(i,j).value)
                                   strPreviousRow=Trim(ObjExcelSheet.cells(i-1,j).value)
                                   If strComp(strChild_Values,strPreviousRow)<>0 Then
                                   
                                       If  strChild_Values<>"" Then
                                            For a = 1 To 20 Step 1
                                                Set obj_tree=description.Create
                                                obj_tree("micclass").value="WebElement"
                                                obj_tree("class").value="mainTabctl1RptExport1trvReportsNode TreeNode"
                                                obj_tree("innertext").value=strChild_Values
                                                wait 1
                                                'set objtreeCount= Browser("Analyzer").Page("Home").Frame("Report Export Import Frame").ChildObjects(obj_tree)
                                                set objtreeCount= Browser("Analyzer").Page("Home").Frame("Report Export Import Frame").WebTable("tabReportExportImport").ChildObjects(obj_tree)
                                                wait 2
                                                intcount=objtreeCount.count

                                                If intcount<>0 Then
                                                    Exit For
                                                End If
                                            Next
                                            
                                            If intcount<>0 Then
                                                k=intcount-1                    
                                                objtreeCount(k).fireEvent  "ondblclick"
                                                wait 1.5
                                                If j>=3 Then
                                                    wait 2
                                                End If                                                
                                                Browser("Analyzer").Page("Home").Sync
                                                Browser("Analyzer").Page("Home").RefreshObject
                                               ' Call ReportStep (StatusTypes.Pass,"Clicked on the Webelement"&Space(2)&strChild_Values, "Report Import/Export Page")                    
                                                wait 2   
                                            Else
                                               ' Call ReportStep (StatusTypes.Fail,"Not Clicked on the Webelement"&Space(2)&strChild_Values, "Report Import/Export Page")
                                                Exit function
                                            End If    
                                            
                                   End If
                                   
                                   
                                       
                                   End If
                                   
                                Next
                                                                
                                Browser("Analyzer").Page("Home").Sync
                                Browser("Analyzer").Page("Home").RefreshObject
                                aReportName =ObjExcelFile.Worksheets(sheetNum).cells(i,intColumn_Count)
                                
                                Set objwebele=Description.Create()
                                objwebele("micclass").value="WebElement"
                                objwebele("html tag").value="DIV"
                                objwebele("innertext").value=aReportName
                                Set objchwebelm= Browser("Analyzer").Page("Home").Frame("Report Export Import Frame").ChildObjects(objwebele)
                                
                                Set desc=Description.Create()
                                desc("micclass").value="WebCheckBox"
                                Set chcnt=objchwebelm(0).ChildObjects(desc)                                
                                wait 2
                                chcnt(0).set "ON"
                                wait 1
                                
                         For p=intColumn_Count-1 to 1  step -1
                         strChild_Values=Trim(ObjExcelSheet.cells(i,p).value)
                         strChild_ValRow2=Trim(ObjExcelSheet.cells(i+1,p).value)
                         If StrComp(strChild_Values,strChild_ValRow2)<>0 Then                             
                             If  strChild_Values<>"" AND i>intRow_Count+1 Then
                            Set obj_tree=description.Create
                            obj_tree("micclass").value="WebElement"
                            obj_tree("class").value="TreeNode"
                            obj_tree("innertext").value=strChild_Values
                            wait 1
                            set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
                            intcount=objtreeCount.count    
                                If intcount<>0 Then
                                   k=intcount-1
                                   objtreeCount(k).fireEvent  "ondblclick"    
                                 End If                          
                            End If
                             
                             
                         End If
                                                 
                         
                        Next
                                
                            
                            Next
                                
                            ObjExcelFile.Close
                            objExcel.Quit    
                            Set objExcel=nothing
                            Set ObjExcelFile=nothing
                            Set ObjExcelSheet=nothing
                            Set objtreeCount=nothing
                            
                            wait 1
                            Browser("Analyzer").Page("Home").Sync
                            Browser("Analyzer").Page("Home").RefreshObject                                                                
                            
                            Browser("Analyzer").Page("Home").Frame("Report Export Import Frame").WebButton("btnExport").Click
                            wait 2
                            'Call Notification Bar to download the selected report
                            'Call function to export XML file
                            ReportImportExport = IMSSCA.General.XMLReportOperation("", strXmlOper)
                            
                            Browser("Analyzer").Page("Home").Sync
                            Browser("Analyzer").Page("Home").RefreshObject 
    End Select
    
End Function


'===========================================================================================================================================================   
    
    '    '<Saves Report in XML format to local machine and returns local path where report has been saved>
	'    '<strXMlName :- XML name>
    '    '<operation :- operation on XML report file>    
	'  	 '<author :-Shweta Nagaral>'
'===========================================================================================================================================================

Public Function XMLReportOperation(ByVal strXMlName, ByVal operation)
	
	Dim i,toolBarNameFormat,strReport,a,b,j,strReportName,strReportFormat,path,formatPos,myVar,objLstWinButton ,objbtnSave,objYes,k,objClose,objbtnCloseExport, midVal
	
	path = Environment.Value("CurrDir")&"Exportresults\"&Trim(strXMlName)&".xml"

	Select Case operation
	
		Case "openXml"
					'Report saving
					For i = 1 To 50 Step 1
						If Browser("Analyzer").WinObject("Notification bar").WinButton("lstWinButton").Exist(2) then
							Exit For
						End If
						If i = 50 Then
							Call ReportStep (StatusTypes.Fail, "Notification bar didnt not popped up", "Report Import/Export Page")
						End If	
					Next
					
					Set objbtnOpen=Browser("Analyzer").WinObject("Notification bar").WinButton("btnOpen")
					Call SCA.ClickOn(objbtnOpen,"btnOpen", "XML Report Export Dialouge")
					wait 2
					
					If Browser("Analyzer").WinObject("Notification bar").WinButton("Close").Exist(2) Then
				      Set objClose=Browser("Analyzer").WinObject("Notification bar").WinButton("Close")
				      Call SCA.ClickOn(objClose,"Close", "XML Report Export Dialouge, Notification Bar")
				    End If
					XMLReportOperation = ""
		Case "saveXml"
					'Report saving
					If Browser("Analyzer").WinObject("Notification bar").Exist(50) then
					    Call ReportStep (StatusTypes.Pass, "Notification bar Exist", "Report Import/Export Page")'Exit For
					Else
					    Call ReportStep (StatusTypes.Fail, "Notification bar didnt not popped up", "Report Import/Export Page")
					    Exitrun 
					End If
					wait 2
					Notificationmessage=Browser("Analyzer").WinObject("Notification bar").GetROProperty("text")
					
					'Msgbox  Notificationmessage
					
					SubNotifcationmessage=Mid(Notificationmessage,Instr(1,Notificationmessage,"export_"),Instr(1,Notificationmessage,".xml")-Instr(1,Notificationmessage,"export_"))
					
					Browser("Analyzer").WinObject("Notification bar").WinButton("Save").Click
					
					Do
						wait 1
						If intLoopCounter=Environment.Value("intCounterMaxLimit") Then
							Exit Do 
					    End If
					
					 	intLoopCounter=intLoopCounter+1
					 	saveMessage=Browser("Analyzer").WinObject("Notification bar").GetROProperty("text")
					Loop Until InStr(1, TRIM(UCASE(saveMessage)), TRIM(UCASE("download has completed"))) > 0
					
					FilePath=Trim(SubNotifcationmessage)
					
					SubNotifcationmessage=SubNotifcationmessage &".xml" 
					
					Set FSo=CreateObject("Scripting.FileSystemObject")
					For i = 1 To 10 Step 1
						If Fso.FileExists("C:\Users\" & Environment("UserName")&"\Downloads\"&SubNotifcationmessage) Then
							Call ReportStep (StatusTypes.Pass, SubNotifcationmessage& "has been saved in Downloads folder successfully", "Report Import/Export Page")
							XMLReportOperation="C:\Users\" & Environment("UserName")&"\Downloads\"&SubNotifcationmessage
							Exit for
					    End if
					    
					    If i = 10 Then
					    	Call ReportStep (StatusTypes.Fail, SubNotifcationmessage& "has not been saved in Downloads folder successfully", "Report Import/Export Page")
						End if
					Next
					
		Case "saveAsXml"
					'Report saving
					For i = 1 To 50 Step 1
						If Browser("Analyzer").WinObject("Notification bar").WinButton("lstWinButton").Exist(2) then
							Exit For
						End If
						If i = 50 Then
							Call ReportStep (StatusTypes.Fail, "Notification bar didnt not popped up", "Report Import/Export Page")
						End If	
					Next
					               
					Set objLstWinButton=Browser("Analyzer").WinObject("Notification bar").WinButton("lstWinButton")
					Call SCA.ClickOn(objLstWinButton,"lstWinButton", "XML Report Export Dialouge")
					
					'Browser("Analyzer").WinMenu("ContextMenu").Select "Save as"
					Browser("Analyzer").WinMenu("SaveAsContextMenu").Select "Save as"
					wait 2
					
					Browser("Analyzer").Dialog("Save As").WinEdit("lnkFile name:").Set path
					
					wait 2
					Set objbtnSave=Browser("Analyzer").Dialog("Save As").WinButton("btnSave")
					Call SCA.ClickOn(objbtnSave,"btnSave", "XML Report Export Dialouge")
					wait 2
				
					IF Browser("Analyzer").Dialog("Save As").Dialog("Confirm Save As").WinButton("Yes").Exist(5) Then
					    Set objYes=Browser("Analyzer").Dialog("Save As").Dialog("Confirm Save As").WinButton("Yes")
					    Call SCA.ClickOn(objYes,"Yes", "XML Report Export Dialouge")
					    wait 2
					End If
					
					For k = 1 To 3 Step 1
					   If Browser("Analyzer").WinObject("Notification bar").WinButton("Close").Exist(1) Then
					      Set objClose=Browser("Analyzer").WinObject("Notification bar").WinButton("Close")
					       Call SCA.ClickOn(objClose,"Close", "XML Report Export Dialouge, Notification Bar")
					       Exit For
					    End If
					Next
					
					wait 2
					Set objbtnCloseExport=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebButton("btnClose")
					Call SCA.ClickOn(objbtnCloseExport,"CloseExportbtn", "XML Report Export Dialouge")
					
					XMLReportOperation=path
					
		Case "openLocalSavedXml"
					
					'Vaildating name and format of report in Local Machine Path
					Set fso = CreateObject("Scripting.FileSystemObject")
					If (fso.FileExists(path)) Then
					   Call ReportStep (StatusTypes.Pass, path & " exists.", "Local Machine Path")
					   Call ReportStep (StatusTypes.Pass, "Successfully found locally saved SCA application XML Report " &strXMlName, "Local Machine Path")
					   'Code to Open Saved Html file in explorer - Start
					   	Dim objIE
						'Create an IE object
						Set objIE = CreateObject("InternetExplorer.Application")
						'Open file
						objIE.Navigate path
					
					Else
					   Call ReportStep (StatusTypes.Fail, path & " doesn't exist.", "Local Machine Path")
					   Call ReportStep (StatusTypes.Fail, "Could not find locally saved SCA application HTML Report " &strXMlName& " successfully", "Local Machine Path")
					End If
					wait 2
		
	End Select
	
End Function


'===========================================================================================================================================================   
    
    '    '<Browses through SCA folders. Search for IMS Health Folder inside Shared Reports>
	'    '<strRefreshPage :- Boolean Vlaue to refresh the page>
    '  	 '<author :-Shweta Nagaral>'
'===========================================================================================================================================================

Public Function OpenReport(ByVal strRefreshPage)
 	If strRefreshPage = 1 Then
 		Set objwebHome= Browser("Analyzer").Page("Shared_Folder").WebElement("webHome")
 		Call SCA.ClickOn(objwebHome,"webHome","Report Creation")
 		
 		Set WshShell =CreateObject("WScript.Shell")
    	WshShell.SendKeys "{F5}"
		wait 2    	
 	End If
 	
	'Search for IMS Health Folder inside Shared Reports
	Browser("Analyzer").Page("Shared_Folder").Sync
	Browser("Analyzer").Page("Shared_Folder").RefreshObject
	Call IMSSCA.General.OperationOn_Folder("fSearch","IMS Health ","Shared", "FolderCreation","")
	Browser("Analyzer").Page("Shared_Folder").Sync
	Call IMSSCA.General.OperationOn_Folder("fSearch","A2. IMS Report Creation","", "FolderCreation", "")
	Browser("Analyzer").Page("Shared_Folder").Sync
	Call IMSSCA.General.OperationOn_Folder("fSearch","India","", "FolderCreation", "")
	Browser("Analyzer").Page("Shared_Folder").Sync
	Call IMSSCA.General.OperationOn_Folder("fSearch","SCA_AutomationReport","", "FolderCreation", "")
	Browser("Analyzer").Page("Shared_Folder").Sync
	Browser("Analyzer").Page("Shared_Folder").RefreshObject
End Function

'==================================================================================================================================================    
    
	  '<Login to Decision Center>
	  '<strUserName: UserName  String>
	  '<StrPwd: Password  String>
'==================================================================================================================================================
	  Public Sub DCLogin(ByVal strUserName,ByVal StrPwd)
	  
		Dim objUserName,objPwd,objLoginBtn,objDCHomePageImage,returnval
		
		SystemUtil.CloseProcessByName("iexplore.exe")
'		SystemUtil.Run Environment.Value("DCURL"),,,,3
		'SystemUtil.Run Environment.Value("DCURL")
		wait 2
		SystemUtil.Run "iexplore.exe", Environment.Value("DCURL"), , ,3     ' 3 is for opening IE in maximized mode
		wait 2
		
		Browser("DC Home").Page("Microsoft Forefront TMG").Sync
		
'		Set objUserName=Browser("DC Home").Page("Microsoft Forefront TMG").WebEdit("txtUsername")
'		wait 2
'		Call SCA.SetText(objUserName,strUserName,"txtUserName" ,"DC Login Page" )
		'new code
		Set objUserName=Browser("micclass:=browser").Page("micclass:=page").WebEdit("xpath:=//input[@id='txtUserID']")
		wait 2
		Call SCA.SetText(objUserName,strUserName,"txtUserName" ,"DC Login Page" )
		'click on continue button
		Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//input[@id='btnValidate']").Click
		wait 4
		'msgbox StrPwd
		Set objPwd=Browser("micclass:=browser").Page("micclass:=page").WebEdit("xpath:=//input[@id='txtPassword']")
		
		Call SCA.SetText(objPwd,trim(StrPwd),"txtpassword" ,"DC Login Page" )
		Set objLoginBtn=Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//input[@id='btnLogin']")
		Call SCA.ClickOn(objLoginBtn,"Login Button","DC Login Page")
		wait 8
		Set wscript=createobject("wscript.shell")
		wscript.SendKeys "{ENTER}"
'		Set objPwd=Browser("DC Home").Page("Microsoft Forefront TMG").WebEdit("txtPassword")
'		Call SCA.SetText(objPwd,StrPwd,"txtpassword" ,"DC Login Page" )	
'		
'		Set objLoginBtn=Browser("DC Home").Page("Microsoft Forefront TMG").WebButton("btnLogin")
'		Call SCA.ClickOn(objLoginBtn,"Login Button","DC Login Page")
		
		Browser("DC Home").Page("Microsoft Forefront TMG").Sync
		Browser("DC Home").Page("Microsoft Forefront TMG").RefreshObject
		
		'Changed by :Poornima  Date:4/19/2018 this object not exist in home page due to branding
		'Set objDCHomePageImage=Browser("DC Home").Page("Home").Image("My_Informed_Decisions")
		wait 2
		'returnval=IMSSCA.Validations.ExistanceWebelements_Verification(objDCHomePageImage, "My_Informed_Decisions", "DC Home Page", 0)
		
'		If returnval=True Then
'			Call ReportStep (StatusTypes.Pass,"Logged into DC Successfully","DC Login Page" )
'		Else
'			Call ReportStep (StatusTypes.Fail,"Coudl not log into DC Successfully","DC Login Page" )
'		End If

		'making to http
'		currurl=Browser("micclass:=browser").Page("micclass:=page").GetROProperty("URL")
'
'		print(currurl)
'		n=replace(currurl,"res://ieframe.dll/invalidcert.htm?SSLError=50331648#","")
'		newurl=replace(n,"https","http")
'		'msgbox newurl
'		Browser("micclass:=browser").Navigate newurl
	  End Sub	
  
'==================================================================================================================================================    
    
	'< Logs Out from DC Application>
'==================================================================================================================================================
	Public Sub DCLogOut()
  
		Dim objlnk_OpenMenu, objlnk_SignOUT
		
		Browser("DC Home").Page("Home").Sync
		'clicking on the Open Menu
		
		Set objlnk_OpenMenu = Browser("DC Home").Page("Home").Link("lnk_Open Menu")
		Call SCA.ClickOn(objlnk_OpenMenu,"lnk_Open Menu", "DC Page")
		
		Browser("DC Home").Page("Home").Sync
		'  Clicking on the Sign out  link
		
		Set objlnk_SignOUT = Browser("DC Home").Page("Home").Link("lnk_Sign OUT")
		Call SCA.ClickOn(objlnk_SignOUT,"lnk_Sign OUT", "DC Page")
		
		Call ReportStep (StatusTypes.Information,"Logged out from DC Successfully","DC Page")
		
	End Sub

'===========================================================================================================================================================   
    
    '    '<Function to remove cubes listing under Cube Node>
	'    '<strCubeName :- Cube Name to be removed>
    '  	 '<author :-Shweta Nagaral>'
'===========================================================================================================================================================

	Public Function RemoveCube(ByVal strCubeName)
		
		Dim cubeName, cubeNameObj
		
		Browser("Analyzer").Page("ReportCreation").Sync
		Browser("Analyzer").Page("ReportCreation").RefreshObject
		
		'Handled Application Time Out in Report Creation Page - Start
       	Set objBtnok = Browser("Analyzer").Dialog("Message from webpage").WinButton("btnOK")
       	If objBtnok.Exist(4) Then
        	Call SCA.ClickOn(objBtnok,"btnOK" , "Application Time Out in Report Creation Page")
       	End If    
        'Handled Application Time Out in Report Creation Page – End
				
		'Descriptive programe to click on Down Normal image of row attribute
		Set cubeName  = Description.Create()
		cubeName("micclass").value = "WebElement"
		cubeName("class").value = "TreeNode"
		cubeName("html tag").value = "SPAN"
		cubeName("innerhtml").value = strCubeName
						
		Set cubeNameObj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(cubeName)
		'msgbox cubeNameObj.count
		
		If cubeNameObj.count = 1 Then
			cubeNameObj(0).RightClick
		ElseIf  cubeNameObj.count = 0 Then
			'TODO based on further enhancements
			Call ReportStep (StatusTypes.Fail,"Could not delete cube "&strCubeName,"Report Creation Page")
			Exit Function
		ElseIf  cubeNameObj.count > 1 Then
			'TODO based on further enhancements
			Call ReportStep (StatusTypes.Fail,"Could not delete cube "&strCubeName&". Multiple cubes with same cube name","Report Creation Page")
			Exit Function
		End If
		
		Set objwelRemoveCube = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welRemoveCube")
 		Call SCA.ClickOn(objwelRemoveCube,"welRemoveCube","Report Creation Page")
 		Call ReportStep (StatusTypes.Pass,"Deleted cube "&strCubeName&" successfully","Report Creation Page")
        
 		'Refresh the cube input details
 		Set objimgRefresh = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgRefresh")
		Call SCA.ClickOn(objimgRefresh,"imgRefresh","Report Creation")
 		
 		'Handled Application Time Out in Report Creation Page - Start
       	Set objBtnok = Browser("Analyzer").Dialog("Message from webpage").WinButton("btnOK")
       	If objBtnok.Exist(4) Then
        	Call SCA.ClickOn(objBtnok,"btnOK" , "Application Time Out in Report Creation Page")
       	End If    
        'Handled Application Time Out in Report Creation Page – End
        
 		wait 1
		Browser("Analyzer").Page("ReportCreation").Sync
		Browser("Analyzer").Page("ReportCreation").RefreshObject
		
		
	End Function
	
	
	 '===========================================================================================================================================================   
        
        '    '<Validates Analytical Validation in the SCA Home Page After Moving Db>
        '	 '<strCAp:- Name of the Analytical pattern created>
		'	  '< strpage:- page where we are creating the Ap>               
        '    '<author :-shobha>'
  '===========================================================================================================================================================
    
    Function AnalyticalPattern(ByVal strCAp,ByVal strPageName)
    
    Dim rc,cc,i,strAPname,strServerName
    	
'    	Browser("Analyzer").Page("Home").Frame("Frame").Link("lnkMy Analytical Patterns").Click
'    	Browser("Analyzer").Page("Home").Sync
'    	Browser("Analyzer").Page("Home").RefreshObject
'		rc=Browser("Analyzer").Page("Home").Frame("AnalyticalPatternFrame").WebTable("tabAnalyticalPattern").RowCount
'		cc=Browser("Analyzer").Page("Home").Frame("AnalyticalPatternFrame").WebTable("tabAnalyticalPattern").ColumnCount(rc)
'
'
'		For i = 1 To rc Step 1
'	
'		strAPname=Browser("Analyzer").Page("Home").Frame("AnalyticalPatternFrame").WebTable("tabAnalyticalPattern").GetCellData(i,2)
'			If StrComp(strAPname,strCAp)=0 Then
'			strServerName=Browser("Analyzer").Page("Home").Frame("AnalyticalPatternFrame").WebTable("tabAnalyticalPattern").GetCellData(i,11)		
'			Exit For
'			End If
'	
'		Next
'
'    	AnalyticalPattern=strServerName

	Browser("Analyzer").Page("Home").Frame("Frame").Link("lnkMy Analytical Patterns").Click
	Browser("Analyzer").Page("Home").Sync
	Browser("Analyzer").Page("Home").RefreshObject
	
	Do
		rc=Browser("Analyzer").Page("Home").Frame("AnalyticalPatternFrame").WebTable("tabAnalyticalPattern").RowCount
		cc=Browser("Analyzer").Page("Home").Frame("AnalyticalPatternFrame").WebTable("tabAnalyticalPattern").ColumnCount(rc)
		
		For i = 1 To rc Step 1
			strAPname=Browser("Analyzer").Page("Home").Frame("AnalyticalPatternFrame").WebTable("tabAnalyticalPattern").GetCellData(i,2)
			If StrComp(strAPname,strCAp)=0 Then
				strServerName=Browser("Analyzer").Page("Home").Frame("AnalyticalPatternFrame").WebTable("tabAnalyticalPattern").GetCellData(i,11)		
				patternFound = 1
				Exit For
			End If
		Next
	    
	    If patternFound = 1 Then
	    	Exit do
	   	else
	        If instr (1, Browser("Analyzer").Page("Home").Frame("Frame").WebElement("welNextPage").GetROProperty("outerhtml"), "COLOR: blue")>0 Then
	        	Set objNextPage = Browser("Analyzer").Page("Home").Frame("Frame").WebElement("welNextPage")
		        Call SCA.ClickOn(objNextPage,"welNextPage", strPagename)
		        Browser("Analyzer").Page("Home").Sync
			Else
				Exit Do	
	        End If
	        
	    End If
	    
	Loop until Browser("Analyzer").Page("Home").Frame("AnalyticalPatternFrame").WebTable("tabAnalyticalPattern").RowCount = 0
	
	AnalyticalPattern=strServerName
    	
    	
    End Function
	
	
Public Function Windows_Login()
	

Browser("Hosting Templates - Ops").Dialog("Windows Security").WinEdit("user_name").Set "internal\dpatro"
Browser("Hosting Templates - Ops").Dialog("Windows Security").WinEdit("Password").Set "deba0223@@@@"
Browser("Hosting Templates - Ops").Dialog("Windows Security").WinButton("OK").Click
End Function	
	
	
	
	
	
	
	
	
	
	
	
	
	





'============================================================================================================================'=============================
'Function Name: AddordeluserstoaopsclientOldFunction
'Description : To add or delete users to a specific client
'Parameters: 
'clientname-> Name of the client,reqAction-> Action to perform on the clint,OPSoperation->Operation to perform on the clint, objData-> This is constant as need to use test data inside the function
'CountryName-> Name of the country,Offering->Offering name ,User->User name,Role->Role name
'Creation Date: 18th April 2015
'Author : IMS Health
'Last Modified Date:NA
'Last Modifyed by :NA

'============================================================================================================================'============================= 



Sub AddordeluserstoaopsclientOldFunction(clientname,reqAction,OPSoperation,ByRef objData,CountryName,Offering,User,Role)
		
		'shweta <30/3/2016> Added validation point to checking Server Response Time / Error captured while creating a new client - start
	    'wait 300
	    'shweta <30/3/2016> Added validation point to checking Server Response Time / Error captured while creating a new client - End
	   
		Set WshShell =CreateObject("WScript.Shell")
		WshShell.SendKeys "{F5}"
		wait 2
		WshShell.SendKeys "{F5}"
		Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
		Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").RefreshObject
		Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
		
		'shweta05/1/2016. Added waitproperty and sync stmts - Start
		Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").RefreshObject
		Set appobject=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops")
		Set objClientName=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtClientName")
		
		For a = 1 To 5 Step 1
			WshShell.SendKeys "{F5}"
			wait 2
			WshShell.SendKeys "{F5}"
			Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
			Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").RefreshObject
			
			Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Link("Hosting").Click
			
			Call IMSSCA.General.menuitemselect(appobject,"Clients")
			Browser("Hosting Templates - Ops").Page("Clients - Ops System").Sync
			Browser("Hosting Templates - Ops").Page("Clients - Ops System").RefreshObject
			
			objClientName.WaitProperty "visible",True,10000
			If objClientName.Exist(1) Then
				Call ReportStep(StatusTypes.Pass, "Successfully redirected to 'Clients' Page in OPS", "OPS Client Page")
				Exit For
			End If
			
			If a=5 Then
				Call ReportStep(StatusTypes.Fail, "Could not successfully redirected to 'Clients' Page in OPS", "OPS Client Page")
				Call ReportStep(StatusTypes.Warning, "Could not successfully redirected to 'Clients' Page in OPS. Further functionality Validation steps may fail", "OPS Client Page")
			End If
		Next
		
 		'Set objClientName=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtClientName")
		'shweta05/1/2016. Added waitproperty and sync stmts - End
		
		Call SCA.SetText(objclientname,clientname,"txtClientName","Clients - Ops System")
	    set WshShell = CreateObject("WScript.Shell")
		Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtClientName").MiddleClick
		WshShell.SendKeys "{TAB}"
		WshShell.SendKeys "{ENTER}"
		wait 5
		introwcount=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("tablClientList").GetROProperty("rows")
		If introwcount>1 Then
			For i=1 to introwcount
				If ucase(Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("tablClientList").GetCellData(i,1))=ucase(clientname) Then
'				   'shweta <30/3/2016> Added validation point to checking Server Response Time / Error captured while creating a new client - start
'				   wait 600
'				   'shweta <30/3/2016> Added validation point to checking Server Response Time / Error captured while creating a new client - End
				   intwebelmcnt=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("tablClientList").ChildItemCount(i,4,"WebElement")
				   For j=0 to intwebelmcnt-1
					   If Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("tablClientList").ChildItem(i,4,"WebElement",j).getroproperty("outertext")=reqAction Then
					   	  wait 2
					   	  Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("tablClientList").ChildItem(i,4,"WebElement",j).Click
					   	  wait 5
					   	  Exit For
					   End If
				   Next
	   		      Exit For 
				End If
			Next
		
			Select Case reqAction
				Case "Manage IAM Users"
					Select Case OPSoperation
						Case "Add"
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebList("Client").Select clientname
							wait 2
							Set objgetroles=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnGet Roles")
							Call SCA.ClickOn(objgetroles,"Get Roles","Add User")
							wait 2
							arrroles=Split(Role,";")
							
							For i = 0 To ubound(arrroles)
								Set rolecheck=Description.Create()
								rolecheck("MicClass").value="WebCheckBox"
								rolecheck("name").value=arrroles(i)
								Set chrolecheck=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(rolecheck)
								chrolecheck(0).Set "ON"
							Next
							
							wait 2
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtUsers").Set User
		'					Set objuser=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtUsers")
		'					Call SCA.SetText(objuser,objData.Item("User"),"Users","Add User")
							'wait 2
							Set objadd=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Add")
							Call SCA.ClickOn(objadd,"Add","Add User")
							Wait 10
						Case "Addusertoroal"
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebList("Client").Select clientname
							wait 2
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtRoles").Set Role
	
							Set objgetroles=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnGet Roles")
							Call SCA.ClickOn(objgetroles,"Get Roles","Add User")
							wait 2
							arrroles=Split(Role,";")
							
							For i = 0 To ubound(arrroles)
								Set rolecheck=Description.Create()
								rolecheck("MicClass").value="WebCheckBox"
								rolecheck("name").value=arrroles(i)
								Set chrolecheck=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(rolecheck)
								chrolecheck(0).Set "ON"
							Next
							
							wait 2
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtUsers").Set User
		'					Set objuser=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtUsers")
		'					Call SCA.SetText(objuser,objData.Item("User"),"Users","Add User")
							'wait 2
							Set objadd=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Add")
							Call SCA.ClickOn(objadd,"Add","Add User")
							Wait 10
							
						Case "Delusertoroal"
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebList("Client").Select clientname
							wait 2
							'<shweta - 7/12/2015> - Start
							offerpermssions=Split(Role,";")
							For i = 0 To UBound(offerpermssions) Step 1
								Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtRoles").Set offerpermssions(i)
								If i = 0 Then
									Set objgetroles=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnGet Roles")
									Call SCA.ClickOn(objgetroles,"Get Roles","Add User")
									wait 2
								End If
								
								Set rolecheck=Description.Create()
								rolecheck("MicClass").value="WebCheckBox"
								rolecheck("name").value=offerpermssions(i)
								Set chrolecheck=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(rolecheck)
								chrolecheck(0).Set "ON"
								
							Next
							'<shweta - 7/12/2015> - End
							
'							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtRoles").Set Role
'	
'							Set objgetroles=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnGet Roles")
'							Call SCA.ClickOn(objgetroles,"Get Roles","Add User")
'							wait 2
'							Set rolecheck=Description.Create()
'							rolecheck("MicClass").value="WebCheckBox"
'							rolecheck("name").value=Role
'							Set chrolecheck=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(rolecheck)
'							chrolecheck(0).Set "ON"
							'Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Select All").Click
							wait 2
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtUsers").Set User
		'					Set objuser=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtUsers")
		'					Call SCA.SetText(objuser,objData.Item("User"),"Users","Add User")
							'wait 2
							Set objDelete=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Delete")
							Call SCA.ClickOn(objDelete,"Delete","Delete User")
							Wait 10
							'Set objDelete=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Delete IAM User")
							'Call SCA.ClickOn(objDelete,"Delete IAM User","Delete User")
							Wait 5
					End Select
				
				Case "Manage IAM Roles"
					Select Case OPSoperation
						Case "CreateOfferingRole"




'============================================================================================================================'==

						If clientname="IMS Health" Then
							
						
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("Country").Click
							set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys CountryName
							wait 5
							Set Desc=Description.Create()
							Desc("MicClass").value="WebElement"
							Desc("class").Value="ui-corner-all"
							Set chobj=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(Desc)
							For i=0 to chobj.count-1
								If chobj(i).getroproperty("innertext")=CountryName Then
									chobj(i).Click
								End If 
							Next
											
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("Offering").Click
							WshShell.SendKeys Offering
							wait 5
							Set Desc=Description.Create()
							Desc("MicClass").value="WebElement"
							Desc("class").Value="ui-corner-all"
							Set chobj=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(Desc)
							For i=0 to chobj.count-1
								If chobj(i).getroproperty("innertext")=Offering Then
									chobj(i).Click
								End If 
							Next
						
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnGet Roles").Click
						wait 2
	
						stractDisplayinrolegrid=objData.Item("ShortName")&" "&objData.Item("CountryShortName")&" "&Replace(Offering," ","")
						
						stroffering="notfound"
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RoleId").ChildItem(2,3,"WebEdit",0).Set stractDisplayinrolegrid
						'Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("RoleName").Set stractDisplayinrolegrid
						'set WshShell = CreateObject("WScript.Shell")
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("RoleName").MiddleClick
						WshShell.SendKeys "{TAB}"
						WshShell.SendKeys "{ENTER}"
						wait 4
						
						
						
						
						
						
						intofferscnt=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").RowCount
						For i=1 to intofferscnt
						 If Ucase(Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").GetCellData(i,3))=Ucase(stractDisplayinrolegrid) Then
						 	stroffering="found"
						 	Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").ChildItem(i,1,"Webcheckbox",0).Set "ON"
						 	Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnDelete Selected Roles").Click
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Ok").Click
							wait 10
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").WaitProperty "Exist",True,30
							strroledeletemsg=" "
							If Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").Exist Then
								strroledeletemsg=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").GetROProperty("innertext")
								If instr(1,Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").GetROProperty("innertext"),"deleted successfully")<>0 Then
									Call ReportStep(StatusTypes.Pass,"After deleting offering Folder "&Offering&" success message should display ","After deleting offering Folder "&Offering&" success message has been displayed ")
								End If
							End If														
							
						 	Exit For
						 End If
						Next
					Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnClose").Click	
					Browser("Hosting Templates - Ops").Page("Clients - Ops System").RefreshObject
						
					intwebelmcnt=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("tablClientList").ChildItemCount(2,4,"WebElement")
				   For j=0 to intwebelmcnt-1
					   If Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("tablClientList").ChildItem(2,4,"WebElement",j).getroproperty("outertext")=reqAction Then
					   	  wait 2
					   	  Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("tablClientList").ChildItem(2,4,"WebElement",j).Click
					   	  wait 5
					   	  Exit For
					   End If
				   Next
               End If
'============================================================================================================================'==
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("Country").Click
							set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys CountryName
							wait 2
							Set Desc=Description.Create()
							Desc("MicClass").value="WebElement"
							Desc("class").Value="ui-corner-all"
							Set chobj=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(Desc)
							For i=0 to chobj.count-1
								If chobj(i).getroproperty("innertext")=CountryName Then
									chobj(i).Click
								End If 
							Next
											
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("Offering").Click
							WshShell.SendKeys Offering
							wait 2
							Set Desc=Description.Create()
							Desc("MicClass").value="WebElement"
							Desc("class").Value="ui-corner-all"
							Set chobj=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(Desc)
							For i=0 to chobj.count-1
								If chobj(i).getroproperty("innertext")=Offering Then
									chobj(i).Click
								End If 
							Next
			
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Create Role").Click
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Ok").Click
							wait 10
							Environment("strresmessage")=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").GetROProperty("innertext")
	
						Case "DeleteOfferingRole"
						
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("Country").Click
							set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys CountryName
							wait 5
							Set Desc=Description.Create()
							Desc("MicClass").value="WebElement"
							Desc("class").Value="ui-corner-all"
							Set chobj=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(Desc)
							For i=0 to chobj.count-1
								If chobj(i).getroproperty("innertext")=CountryName Then
									chobj(i).Click
								End If 
							Next
											
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("Offering").Click
							WshShell.SendKeys Offering
							wait 5
							Set Desc=Description.Create()
							Desc("MicClass").value="WebElement"
							Desc("class").Value="ui-corner-all"
							Set chobj=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(Desc)
							For i=0 to chobj.count-1
								If chobj(i).getroproperty("innertext")=Offering Then
									chobj(i).Click
								End If 
							Next
						
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnGet Roles").Click
						wait 2
	
						stractDisplayinrolegrid=objData.Item("ShortName")&" "&objData.Item("CountryShortName")&" "&Replace(Offering," ","")
						
						stroffering="notfound"
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RoleId").ChildItem(2,3,"WebEdit",0).Set stractDisplayinrolegrid
						'Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("RoleName").Set stractDisplayinrolegrid
						'set WshShell = CreateObject("WScript.Shell")
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("RoleName").MiddleClick
						WshShell.SendKeys "{TAB}"
						WshShell.SendKeys "{ENTER}"
						wait 4
						intofferscnt=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").RowCount
						For i=1 to intofferscnt
						 If Ucase(Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").GetCellData(i,3))=Ucase(stractDisplayinrolegrid) Then
						 	stroffering="found"
						 	Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").ChildItem(i,1,"Webcheckbox",0).Set "ON"
						 	Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnDelete Selected Roles").Click
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Ok").Click
							wait 10
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").WaitProperty "Exist",True,30
							strroledeletemsg=" "
							If Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").Exist Then
								strroledeletemsg=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").GetROProperty("innertext")
								If instr(1,Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").GetROProperty("innertext"),"deleted successfully")<>0 Then
									Call ReportStep(StatusTypes.Pass,"After deleting offering Folder "&Offering&" success message should display ","After deleting offering Folder "&Offering&" success message has been displayed ")
								End If
							End If														
							
						 	Exit For
						 End If
						Next
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnGet Roles").Click
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RoleId").ChildItem(2,3,"WebEdit",0).Set stractDisplayinrolegrid
						'Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("RoleName").Set stractDisplayinrolegrid
						'set WshShell = CreateObject("WScript.Shell")
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("RoleName").MiddleClick
						WshShell.SendKeys "{TAB}"
						WshShell.SendKeys "{ENTER}"
						wait 4
						
						'Rajesh- 18/04/2016
						
						If (Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("welNoRecords").Exist(10)) Then
							
'							intofferscnt=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").RowCount
							strdelrole="deleted"
						Else
							strdelrole="notdeleted"
							
						End If
						
						
'						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("welNoRecords").Exist(10)
'						intofferscnt=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").RowCount
'						strdelrole="deleted"
						'Rajesh- 18/04/2016
'						For l=2 to  intofferscnt
'						If Ucase(Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").GetCellData(l,3))=Ucase(stractDisplayinrolegrid)	Then
'							strdelrole="notdeleted"
'							Exit For
'						End If
'						Next
						'Rajesh- 18/04/2016
						If strdelrole="deleted" Then
							Call ReportStep (StatusTypes.Pass,"Deleted role offer: "&stractDisplayinrolegrid&" should not display in the result grid", "Deleted role offer: "&stractDisplayinrolegrid&" is not displayed in the result grid")
	  		 				Else
	    	 				Call ReportStep (StatusTypes.Fail,"Deleted role offer: "&stractDisplayinrolegrid&" should not be display in the result grid", "Deleted role offer: "&stractDisplayinrolegrid&" is displayed in the result grid")
						End If
						If instr(1,strroledeletemsg,"deleted successfully")=0 Then
						
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnClose").Click
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").Link("Event Log & Subscription").Click
							Set appobject=Browser("Hosting Templates - Ops").Page("Clients - Ops System")
							Call menuitemselect(appobject,"Hosting Event Log")
							Set objclientname=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtClientName")
							Call SCA.SetText(objclientname,clientname,"txtClientName","Clients - Ops System")
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtClientName").MiddleClick
							WshShell.SendKeys "{TAB}"
							WshShell.SendKeys "{ENTER}"
							wait 5
							introwcount=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("Template/Client Id").GetROProperty("rows")
							If introwcount>1 Then
								For i=1 to introwcount
									If ucase(Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("Template/Client Id").GetCellData(i,1))=ucase(clientname) Then
									   strEventlogmsg=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("Template/Client Id").GetCellData(i,4)
									   Exit For
									End If
								Next
								If instr(1,strEventlogmsg,"deleted by")<>0 Then
									Call ReportStep(StatusTypes.Pass,"After deleting offering Folder "&Offering&" success message should display ","After deleting offering Folder "&Offering&" success message has been displayed in event log")
									Else
									Call ReportStep(StatusTypes.Fail,"After deleting offering Folder "&Offering&" success message should display ","After deleting offering Folder "&Offering&" success message has not been displayed in event log as well")					
								End If
							End If
						
						End If
						
						Case "DeleteOfferRoleholdmsg"
						
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("Country").Click
							set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys CountryName
							wait 5
							Set Desc=Description.Create()
							Desc("MicClass").value="WebElement"
							Desc("class").Value="ui-corner-all"
							Set chobj=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(Desc)
							For i=0 to chobj.count-1
								If chobj(i).getroproperty("innertext")=CountryName Then
									chobj(i).Click
								End If 
							Next
											
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("Offering").Click
							WshShell.SendKeys Offering
							wait 5
							Set Desc=Description.Create()
							Desc("MicClass").value="WebElement"
							Desc("class").Value="ui-corner-all"
							Set chobj=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(Desc)
							For i=0 to chobj.count-1
								If chobj(i).getroproperty("innertext")=Offering Then
									chobj(i).Click
								End If 
							Next
						
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnGet Roles").Click
						wait 2
	
						stractDisplayinrolegrid=objData.Item("ShortName")&" "&objData.Item("CountryShortName")&" "&Replace(Offering," ","")
						
						stroffering="notfound"
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RoleId").ChildItem(2,3,"WebEdit",0).Set stractDisplayinrolegrid
						'Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("RoleName").Set stractDisplayinrolegrid
						'set WshShell = CreateObject("WScript.Shell")
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("RoleName").MiddleClick
						WshShell.SendKeys "{TAB}"
						WshShell.SendKeys "{ENTER}"
						wait 10
						intofferscnt=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").RowCount
						For i=1 to intofferscnt
						 If Ucase(Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").GetCellData(i,3))=Ucase(stractDisplayinrolegrid) Then
						 	stroffering="found"
						 	Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").ChildItem(i,1,"Webcheckbox",0).Set "ON"
						 	Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnDelete Selected Roles").Click
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Ok").Click
							wait 10
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").WaitProperty "Exist",True,30
							Environment("strdeletemsg")=" "
							msgexist=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").Exist
							
							If msgexist=True Then
								if instr(1,Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").GetROProperty("innertext"),"error")<>0 then
									Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("PartialMessage").Click
									wait 2
									Environment("strdeletemsg")=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Messageadescription").GetROProperty("innertext")
									Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Ok").Click
								Else
								Environment("strdeletemsg")=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").GetROProperty("innertext")	
								End If
																	
							End If						
							
						 	Exit For
						 End If
						Next
	
						If msgexist=False Then
						
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnClose").Click
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").Link("Event Log & Subscription").Click
							Set appobject=Browser("Hosting Templates - Ops").Page("Clients - Ops System")
							Call menuitemselect(appobject,"Hosting Event Log")
							Set objclientname=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtClientName")
							Call SCA.SetText(objclientname,clientname,"txtClientName","Clients - Ops System")
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtClientName").MiddleClick
							WshShell.SendKeys "{TAB}"
							WshShell.SendKeys "{ENTER}"
							wait 5
							introwcount=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("Template/Client Id").GetROProperty("rows")
							If introwcount>1 Then
								For i=1 to introwcount
									If ucase(Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("Template/Client Id").GetCellData(i,1))=ucase(clientname) Then
									   Environment("strdeletemsg")=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("Template/Client Id").GetCellData(i,4)
									   Exit For
									End If
								Next
	
							End If
						
						End If
						
					End Select
					
				Case "IAM DataSource Permission"
					strDataSource = User
					strRole=Role
							
					Select Case OPSoperation
						Case "Add"
							
							'TODO: Ask Mani to click on "IAM DataSource Permission"
							Set WshShell = CreateObject("WScript.Shell")
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtAddDatasource").Click
					        WshShell.SendKeys strDataSource
					        wait 2 ' To wait for List to display as per keyword entered above
					        WshShell.SendKeys "{DOWN}" 
					        wait 2 'To select item from list and wait
					        WshShell.SendKeys "{ENTER}"
					        Set WshShell = Nothing

							Set WshShell = CreateObject("WScript.Shell")
							Set objtxtAddDataSourceRole=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtAddDataSourceRole")
							Call SCA.SetText(objtxtAddDataSourceRole,strRole,"txtAddDataSourceRole","IAM DataSource Permission")
							wait 2
				            WshShell.SendKeys "{ENTER}"
				            wait 1
				            Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtAddDataSourceRole").Click
							WshShell.SendKeys "{ENTER}"
				            Set WshShell = Nothing
							wait 2
							
							set objchkAddDatasourceWithRoles= Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebCheckBox("chkAddDatasourceWithRoles")
							Call SCA.SetCheckBox(objchkAddDatasourceWithRoles,"chkAddDatasourceWithRoles","ON","IAM DataSource Permission")		    
							
							Set objbtnAddDatasourceWithRoles = Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnAddDatasourceWithRoles")
							Call SCA.ClickOn(objbtnAddDatasourceWithRoles,"btnAddDatasourceWithRoles","IAM DataSource Permission")
							
							'Validating "Are you sure you want to add selected Role(s)?"
							If Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Are you sure you want to add selected Roles").Exist(2) then
								Set objbtnOk = Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Ok")
								Call SCA.ClickOn(objbtnOk,"Ok","IAM DataSource Permission")
								
								Call ReportStep (StatusTypes.Pass,"Successfully Added added Datasoure "&strDataSource&" with Role(s) "&strRole&" mapping","IAM DataSource Permission")
							End If
							
						Case "Remove"
							
							Set WshShell = CreateObject("WScript.Shell")
							Set objtxtRemoveDatasourceRoleName= Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtRemoveDatasourceRoleName")
							Call SCA.SetText(objtxtRemoveDatasourceRoleName,strRole, "txtRemoveDatasourceRoleName","IAM DataSource Permission")
							wait 2
				            WshShell.SendKeys "{ENTER}"
				            wait 1
				            Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtRemoveDatasourceRoleName").Click
							WshShell.SendKeys "{ENTER}"
				            wait 2
							
							Set objtxtRemoveDataSourceName = Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtRemoveDataSourceName")
							Call SCA.SetText(objtxtRemoveDataSourceName,strDataSource, "txtRemoveDataSourceName","IAM DataSource Permission")
							wait 2
				            WshShell.SendKeys "{ENTER}"
				            wait 1
				            Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtRemoveDataSourceName").Click
							WshShell.SendKeys "{ENTER}"
				            Set WshShell = Nothing
							wait 2
							
							rc = Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("tabRemoveDatasourceWithRoles").GetROProperty("rows")
							If rc=2 Then
								set objchkRemoveDatasourceWithRoles= Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebCheckBox("chkRemoveDatasourceWithRoles")
								Call SCA.SetCheckBox(objchkRemoveDatasourceWithRoles,"chkRemoveDatasourceWithRoles","ON","IAM DataSource Permission")		    
							
								set objbtnRemoveDatasourceWithRoles = Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnRemoveDatasourceWithRoles")
								Call SCA.ClickOn(objbtnRemoveDatasourceWithRoles,"btnRemoveDatasourceWithRoles","IAM DataSource Permission")
								
								'Validating "Are you sure you want to delete selected Datasoure with Role(s) mapping?"
								If Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Are you sure you want to delete selected Datasoure").Exist(2) then
									Set objbtnOk = Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Ok")
									Call SCA.ClickOn(objbtnOk,"Ok","IAM DataSource Permission")
									
									Call ReportStep (StatusTypes.Pass,"Successfully deleted selected Datasoure "&strDataSource&" with Role(s) "&strRole&" mapping","IAM DataSource Permission")
								End If
							else
								Call ReportStep (StatusTypes.Fail,"Could not delete selected Datasoure "&strDataSource&" with Role(s) "&strRole&" mapping","IAM DataSource Permission")
							End If
							
					End Select
					
					Set objClose=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnClose")
        			Call SCA.ClickOn(objClose,"btnClose","NewClient")
					
					Browser("Hosting Templates - Ops").Page("Clients - Ops System").Sync
					Browser("Hosting Templates - Ops").Page("Clients - Ops System").RefreshObject
			End Select
		End If
	End Sub


'============================================================================================================================'=============================
'Function Name: menuitemselect
'Description : To select child menu item from the menu
'Parameters: appobject-> req menu object where it has been displayed, menuitem-> required menuitem
'Creation Date: 08th July 2015
'Author : IMS Health

'============================================================================================================================'============================= 





Sub menuitemselect(ByVal appobject,ByVal menuitem)
  	itemClick = 0
  	Set menuChild=Description.Create()
	menuChild("micclass").value="Link"
	menuChild("Class").Regularexpression=True
	menuChild("Class").value=".*childMenuItem.*"
	set childcnt=appobject.ChildObjects(menuChild)
	
	For i = 0 To childcnt.count-1
		if childcnt(i).getroproperty("outertext")=menuitem then
			'shweta - start
			
			'For m = 1 To 5 Step 1
			'	'Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Link("Hosting").Click
			'	childcnt(i).WaitProperty "visible",True,10000
			'	If childcnt(i).Exist(2) Then
			'		childcnt(i).click
			'		itemClick = 1
			'		Exit for
			'	End If
			'Next
			'
			'If itemClick = 1 Then
			'	Exit For
			'End If

			'shweta 22/3/2016 Added for intCounterStart/intCounterMaxLimit counter - Start
			For m = Environment.Value("intCounterStart") To Environment.Value("intCounterMaxLimit") Step 1
				'Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Link("Hosting").Click
'				childcnt(i).WaitProperty "Visible",True,10000
				If childcnt(i).Exist(10) Then
					childcnt(i).click
					itemClick = 1
					Call ReportStep (StatusTypes.Pass,"Sucessfully re-directed to requested menuitem "&menuitem&" Page from OPS Home Page","OPS Home Page")
					Exit for
				End If
			Next
			
			If itemClick = 1 Then
				Exit For
			End If
			'shweta 22/3/2016 Added for intCounterStart/intCounterMaxLimit counter - End
	
			'shweta - End
		End If
	Next
	
	'shweta 22/3/2016 - Added Report fail stmt if could not re-direct to menuitem Page - start
	If itemClick = 0 Then
		Call ReportStep (StatusTypes.Fail,"Could not re-direct to requested menuitem "&menuitem&" Page from OPS Home Page","OPS Home Page")
	End If
	'shweta 22/3/2016 - Added Report fail stmt if could not re-direct to menuitem Page - End
	
	'shweta 22/3/2016 - Added ReadyState method to check if browser loading is completed  - start
	intLoopCounter = Environment.Value("intCounterStart")
	Do
		wait (1)
		If intLoopCounter=Environment.Value("intCounterMaxLimit") Then
	             Call ReportStep (StatusTypes.Fail,"Browser loading is not completed menuitem "&menuitem&" Page from OPS Home Page","OPS Home Page")
	             Exit Do 
	    End If
		appobject.sync
		intLoopCounter=intLoopCounter+1
		'Call ReportStep (StatusTypes.Information, appobject&" Page from OPS loaded successfully","OPS Home Page")
	Loop While (( appobject.Object.ReadyState) <>"complete")
	'shweta 22/3/2016 - Added ReadyState method to check if browser loading is completed  - End
	
	appobject.sync
	appobject.RefreshObject
  End Sub



'================================================================================================='=========================='==============================
'Function Name: Addordeluserstoaopsclient
'Description : To add or delete users to a specific client
'Parameters: 
'clientname-> Name of the client,reqAction-> Action to perform on the clint,OPSoperation->Operation to perform on the clint, objData-> This is constant as need to use test data inside the function
'CountryName-> Name of the country,Offering->Offering name ,User->User name,Role->Role name
'Creation Date: 18th April 2015
'Author : IMS Health
'Last Modified Date:NA
'Last Modifyed by :NA

'================================================================================================='=========================='============================= 





Sub Addordeluserstoaopsclient(clientname,reqAction,OPSoperation,ByRef objData,CountryName,Offering,User,Role)
		
		'shweta <30/3/2016> Added validation point to checking Server Response Time / Error captured while creating a new client - start
	    'wait 300
	    'shweta <30/3/2016> Added validation point to checking Server Response Time / Error captured while creating a new client - End
	   
		Set WshShell =CreateObject("WScript.Shell")
		WshShell.SendKeys "{F5}"
		wait 2
		WshShell.SendKeys "{F5}"
		Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
		Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").RefreshObject
		Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
		
		'shweta05/1/2016. Added waitproperty and sync stmts - Start
		Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").RefreshObject
		Set appobject=Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops")
		Set objClientName=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtClientName")
		
		'<Shweta - 7/2016> - created btnClose object - Start 
		Set objClose=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnClose")
		'<Shweta - 7/2016> - created btnClose object - End
		
		For a = 1 To 5 Step 1
			WshShell.SendKeys "{F5}"
			wait 2
			WshShell.SendKeys "{F5}"
			Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Sync
			Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").RefreshObject
			
			Browser("Hosting Templates - Ops").Page("Hosting Templates - Ops").Link("Hosting").Click
			
			Call IMSSCA.General.menuitemselect(appobject,"Clients")
			Browser("Hosting Templates - Ops").Page("Clients - Ops System").Sync
			Browser("Hosting Templates - Ops").Page("Clients - Ops System").RefreshObject
			
			'objClientName.WaitProperty "visible",True,10000
			
			If objClientName.Exist(30) Then
				Call ReportStep(StatusTypes.Pass, "Successfully redirected to 'Clients' Page in OPS", "OPS Client Page")
				Exit For
			End If
			
			If a=5 Then
				Call ReportStep(StatusTypes.Fail, "Could not successfully redirected to 'Clients' Page in OPS", "OPS Client Page")
				Call ReportStep(StatusTypes.Warning, "Could not successfully redirected to 'Clients' Page in OPS. Further functionality Validation steps may fail", "OPS Client Page")
			End If
		Next
		
 		'Set objClientName=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtClientName")
		'shweta05/1/2016. Added waitproperty and sync stmts - End
		
		Call SCA.SetText(objclientname,clientname,"txtClientName","Clients - Ops System")
	    set WshShell = CreateObject("WScript.Shell")
		Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtClientName").MiddleClick
		WshShell.SendKeys "{TAB}"
		WshShell.SendKeys "{ENTER}"
		wait 5
		introwcount=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("tablClientList").GetROProperty("rows")
		If introwcount>1 Then
			For i=1 to introwcount
				If ucase(Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("tablClientList").GetCellData(i,1))=ucase(clientname) Then
'				   'shweta <30/3/2016> Added validation point to checking Server Response Time / Error captured while creating a new client - start
'				   wait 600
'				   'shweta <30/3/2016> Added validation point to checking Server Response Time / Error captured while creating a new client - End
				   intwebelmcnt=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("tablClientList").ChildItemCount(i,4,"WebElement")
				   For j=0 to intwebelmcnt-1
					   If Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("tablClientList").ChildItem(i,4,"WebElement",j).getroproperty("outertext")=reqAction Then
					   	  wait 2
					   	  Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("tablClientList").ChildItem(i,4,"WebElement",j).Click
					   	  wait 5
					   	  Exit For
					   End If
				   Next
	   		      Exit For 
				End If
			Next
		
			Select Case reqAction
				Case "Manage IAM Users"
					Select Case OPSoperation
						Case "Add"
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebList("Client").Select clientname
							wait 2
							Set objgetroles=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnGet Roles")
							Call SCA.ClickOn(objgetroles,"Get Roles","Add User")
							wait 2
							arrroles=Split(Role,";")
							
							For i = 0 To ubound(arrroles)
								Set rolecheck=Description.Create()
								rolecheck("MicClass").value="WebCheckBox"
								rolecheck("name").value=arrroles(i)
								Set chrolecheck=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(rolecheck)
								chrolecheck(0).Set "ON"
							Next
							
							wait 2
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtUsers").Set User
		'					Set objuser=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtUsers")
		'					Call SCA.SetText(objuser,objData.Item("User"),"Users","Add User")
							'wait 2
							Set objadd=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Add")
							Call SCA.ClickOn(objadd,"Add","Add User")
							Wait 10
						Case "Addusertoroal"
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebList("Client").Select clientname
							wait 2
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtRoles").Set Role
	
							Set objgetroles=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnGet Roles")
							Call SCA.ClickOn(objgetroles,"Get Roles","Add User")
							wait 2
							arrroles=Split(Role,";")
							
							For i = 0 To ubound(arrroles)
								Set rolecheck=Description.Create()
								rolecheck("MicClass").value="WebCheckBox"
								rolecheck("name").value=arrroles(i)
								Set chrolecheck=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(rolecheck)
								chrolecheck(0).Set "ON"
							Next
							
							wait 2
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtUsers").Set User
		'					Set objuser=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtUsers")
		'					Call SCA.SetText(objuser,objData.Item("User"),"Users","Add User")
							'wait 2
							Set objadd=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Add")
							Call SCA.ClickOn(objadd,"Add","Add User")
							Wait 10
							
						Case "Delusertoroal"
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebList("Client").Select clientname
							wait 2
							'<shweta - 7/12/2015> - Start
							offerpermssions=Split(Role,";")
							For i = 0 To UBound(offerpermssions) Step 1
								Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtRoles").Set offerpermssions(i)
								If i = 0 Then
									Set objgetroles=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnGet Roles")
									Call SCA.ClickOn(objgetroles,"Get Roles","Add User")
									wait 2
								End If
								
								Set rolecheck=Description.Create()
								rolecheck("MicClass").value="WebCheckBox"
								rolecheck("name").value=offerpermssions(i)
								Set chrolecheck=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(rolecheck)
								chrolecheck(0).Set "ON"
								
							Next
							'<shweta - 7/12/2015> - End
							
'							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtRoles").Set Role
'	
'							Set objgetroles=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnGet Roles")
'							Call SCA.ClickOn(objgetroles,"Get Roles","Add User")
'							wait 2
'							Set rolecheck=Description.Create()
'							rolecheck("MicClass").value="WebCheckBox"
'							rolecheck("name").value=Role
'							Set chrolecheck=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(rolecheck)
'							chrolecheck(0).Set "ON"
							'Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Select All").Click
							wait 2
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtUsers").Set User
		'					Set objuser=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtUsers")
		'					Call SCA.SetText(objuser,objData.Item("User"),"Users","Add User")
							'wait 2
							Set objDelete=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Delete")
							Call SCA.ClickOn(objDelete,"Delete","Delete User")
							Wait 10
							'Set objDelete=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Delete IAM User")
							'Call SCA.ClickOn(objDelete,"Delete IAM User","Delete User")
							Wait 5
					End Select
				
				Case "Manage IAM Roles"
					Select Case OPSoperation
						Case "CreateOfferingRole"




'============================================================================================================================'==

						If clientname="IMS Health" Then
							
						
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("Country").Click
							set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys CountryName
							wait 5
							Set Desc=Description.Create()
							Desc("MicClass").value="WebElement"
							Desc("class").Value="ui-corner-all"
							Set chobj=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(Desc)
							For i=0 to chobj.count-1
								If chobj(i).getroproperty("innertext")=CountryName Then
									chobj(i).Click
								End If 
							Next
											
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("Offering").Click
							WshShell.SendKeys Offering
							wait 5
							Set Desc=Description.Create()
							Desc("MicClass").value="WebElement"
							Desc("class").Value="ui-corner-all"
							Set chobj=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(Desc)
							For i=0 to chobj.count-1
								If chobj(i).getroproperty("innertext")=Offering Then
									chobj(i).Click
								End If 
							Next
						
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnGet Roles").Click
						wait 2
	
						stractDisplayinrolegrid=objData.Item("ShortName")&" "&objData.Item("CountryShortName")&" "&Replace(Offering," ","")
						
						stroffering="notfound"
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RoleId").ChildItem(2,3,"WebEdit",0).Set stractDisplayinrolegrid
						'Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("RoleName").Set stractDisplayinrolegrid
						'set WshShell = CreateObject("WScript.Shell")
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("RoleName").MiddleClick
						WshShell.SendKeys "{TAB}"
						WshShell.SendKeys "{ENTER}"
						wait 4
						
						'Rajesh- 19/04/2016
						
						If (Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("welNoRecords").Exist(10)) Then
							
						Else
							intofferscnt=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").RowCount
							For i=1 to intofferscnt
							 If Ucase(Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").GetCellData(i,3))=Ucase(stractDisplayinrolegrid) Then
							 	stroffering="found"
							 	Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").ChildItem(i,1,"Webcheckbox",0).Set "ON"
							 	Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnDelete Selected Roles").Click
								Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Ok").Click
								wait 10
'								Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").WaitProperty "Exist",True,30
								strroledeletemsg=" "
								If Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").Exist(30) Then
									strroledeletemsg=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").GetROProperty("innertext")
									If instr(1,Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").GetROProperty("innertext"),"deleted successfully")<>0 Then
										Call ReportStep(StatusTypes.Pass,"After deleting offering Folder "&Offering&" success message should display ","After deleting offering Folder "&Offering&" success message has been displayed ")
									End If
								End If														
								
							 	Exit For
							 End If
							Next							
							
					   End If
						
						
						
						'Rajesh- 19/04/2016
						
						
						
					Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnClose").Click	
					Browser("Hosting Templates - Ops").Page("Clients - Ops System").RefreshObject
						
					intwebelmcnt=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("tablClientList").ChildItemCount(2,4,"WebElement")
				   For j=0 to intwebelmcnt-1
					   If Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("tablClientList").ChildItem(2,4,"WebElement",j).getroproperty("outertext")=reqAction Then
					   	  wait 2
					   	  Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("tablClientList").ChildItem(2,4,"WebElement",j).Click
					   	  wait 5
					   	  Exit For
					   End If
				   Next
               End If
'============================================================================================================================'==
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("Country").Click
							set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys CountryName
							wait 2
							Set Desc=Description.Create()
							Desc("MicClass").value="WebElement"
							Desc("class").Value="ui-corner-all"
							Set chobj=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(Desc)
							For i=0 to chobj.count-1
								If chobj(i).getroproperty("innertext")=CountryName Then
									chobj(i).Click
								End If 
							Next
											
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("Offering").Click
							WshShell.SendKeys Offering
							wait 2
							Set Desc=Description.Create()
							Desc("MicClass").value="WebElement"
							Desc("class").Value="ui-corner-all"
							Set chobj=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(Desc)
							For i=0 to chobj.count-1
								If chobj(i).getroproperty("innertext")=Offering Then
									chobj(i).Click
								End If 
							Next
			
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Create Role").Click
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Ok").Click
							wait 10
							Environment("strresmessage")=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").GetROProperty("innertext")
	
						Case "DeleteOfferingRole"
						
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("Country").Click
							set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys CountryName
							wait 5
							Set Desc=Description.Create()
							Desc("MicClass").value="WebElement"
							Desc("class").Value="ui-corner-all"
							Set chobj=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(Desc)
							For i=0 to chobj.count-1
								If chobj(i).getroproperty("innertext")=CountryName Then
									chobj(i).Click
								End If 
							Next
											
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("Offering").Click
							WshShell.SendKeys Offering
							wait 5
							Set Desc=Description.Create()
							Desc("MicClass").value="WebElement"
							Desc("class").Value="ui-corner-all"
							Set chobj=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(Desc)
							For i=0 to chobj.count-1
								If chobj(i).getroproperty("innertext")=Offering Then
									chobj(i).Click
								End If 
							Next
						
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnGet Roles").Click
						wait 2
	
						stractDisplayinrolegrid=objData.Item("ShortName")&" "&objData.Item("CountryShortName")&" "&Replace(Offering," ","")
						
						stroffering="notfound"
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RoleId").ChildItem(2,3,"WebEdit",0).Set stractDisplayinrolegrid
						'Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("RoleName").Set stractDisplayinrolegrid
						'set WshShell = CreateObject("WScript.Shell")
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("RoleName").MiddleClick
						WshShell.SendKeys "{TAB}"
						WshShell.SendKeys "{ENTER}"
						wait 4
						intofferscnt=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").RowCount
						For i=1 to intofferscnt
						 If Ucase(Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").GetCellData(i,3))=Ucase(stractDisplayinrolegrid) Then
						 	stroffering="found"
						 	Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").ChildItem(i,1,"Webcheckbox",0).Set "ON"
						 	Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnDelete Selected Roles").Click
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Ok").Click
							wait 10
'							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").WaitProperty "Exist",True,30
							strroledeletemsg=" "
							If Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").Exist(30) Then
								strroledeletemsg=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").GetROProperty("innertext")
								If instr(1,Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").GetROProperty("innertext"),"deleted successfully")<>0 Then
									Call ReportStep(StatusTypes.Pass,"After deleting offering Folder "&Offering&" success message should display ","After deleting offering Folder "&Offering&" success message has been displayed ")
								End If
							End If														
							
						 	Exit For
						 End If
						Next
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnGet Roles").Click
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RoleId").ChildItem(2,3,"WebEdit",0).Set stractDisplayinrolegrid
						'Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("RoleName").Set stractDisplayinrolegrid
						'set WshShell = CreateObject("WScript.Shell")
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("RoleName").MiddleClick
						WshShell.SendKeys "{TAB}"
						WshShell.SendKeys "{ENTER}"
						wait 4
						
						'Rajesh- 18/04/2016
						
						If (Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("welNoRecords").Exist(10)) Then
							
'							intofferscnt=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").RowCount
							strdelrole="deleted"
						Else
							strdelrole="notdeleted"
							
						End If
						
						
'						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("welNoRecords").Exist(10)
'						intofferscnt=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").RowCount
'						strdelrole="deleted"
						'Rajesh- 18/04/2016
'						For l=2 to  intofferscnt
'						If Ucase(Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").GetCellData(l,3))=Ucase(stractDisplayinrolegrid)	Then
'							strdelrole="notdeleted"
'							Exit For
'						End If
'						Next
						'Rajesh- 18/04/2016
						If strdelrole="deleted" Then
							Call ReportStep (StatusTypes.Pass,"Deleted role offer: "&stractDisplayinrolegrid&" should not display in the result grid", "Deleted role offer: "&stractDisplayinrolegrid&" is not displayed in the result grid")
	  		 				Else
	    	 				Call ReportStep (StatusTypes.Fail,"Deleted role offer: "&stractDisplayinrolegrid&" should not be display in the result grid", "Deleted role offer: "&stractDisplayinrolegrid&" is displayed in the result grid")
						End If
						If instr(1,strroledeletemsg,"deleted successfully")=0 Then
						
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnClose").Click
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").Link("Event Log & Subscription").Click
							Set appobject=Browser("Hosting Templates - Ops").Page("Clients - Ops System")
							Call menuitemselect(appobject,"Hosting Event Log")
							Set objclientname=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtClientName")
							Call SCA.SetText(objclientname,clientname,"txtClientName","Clients - Ops System")
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtClientName").MiddleClick
							WshShell.SendKeys "{TAB}"
							WshShell.SendKeys "{ENTER}"
							wait 5
							introwcount=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("Template/Client Id").GetROProperty("rows")
							If introwcount>1 Then
								For i=1 to introwcount
									If ucase(Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("Template/Client Id").GetCellData(i,1))=ucase(clientname) Then
									   strEventlogmsg=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("Template/Client Id").GetCellData(i,4)
									   Exit For
									End If
								Next
								If instr(1,strEventlogmsg,"deleted by")<>0 Then
									Call ReportStep(StatusTypes.Pass,"After deleting offering Folder "&Offering&" success message should display ","After deleting offering Folder "&Offering&" success message has been displayed in event log")
									Else
									Call ReportStep(StatusTypes.Fail,"After deleting offering Folder "&Offering&" success message should display ","After deleting offering Folder "&Offering&" success message has not been displayed in event log as well")					
								End If
							End If
						
						End If
						
						Case "DeleteOfferRoleholdmsg"
						
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("Country").Click
							set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys CountryName
							wait 5
							Set Desc=Description.Create()
							Desc("MicClass").value="WebElement"
							Desc("class").Value="ui-corner-all"
							Set chobj=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(Desc)
							For i=0 to chobj.count-1
								If chobj(i).getroproperty("innertext")=CountryName Then
									chobj(i).Click
								End If 
							Next
											
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("Offering").Click
							WshShell.SendKeys Offering
							wait 5
							Set Desc=Description.Create()
							Desc("MicClass").value="WebElement"
							Desc("class").Value="ui-corner-all"
							Set chobj=Browser("Hosting Templates - Ops").Page("Clients - Ops System").ChildObjects(Desc)
							For i=0 to chobj.count-1
								If chobj(i).getroproperty("innertext")=Offering Then
									chobj(i).Click
								End If 
							Next
						
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnGet Roles").Click
						wait 2
	
						stractDisplayinrolegrid=objData.Item("ShortName")&" "&objData.Item("CountryShortName")&" "&Replace(Offering," ","")
						
						stroffering="notfound"
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RoleId").ChildItem(2,3,"WebEdit",0).Set stractDisplayinrolegrid
						'Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("RoleName").Set stractDisplayinrolegrid
						'set WshShell = CreateObject("WScript.Shell")
						Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("RoleName").MiddleClick
						WshShell.SendKeys "{TAB}"
						WshShell.SendKeys "{ENTER}"
						wait 10
						intofferscnt=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").RowCount
						For i=1 to intofferscnt
						 If Ucase(Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").GetCellData(i,3))=Ucase(stractDisplayinrolegrid) Then
						 	stroffering="found"
						 	Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("RolesGrid").ChildItem(i,1,"Webcheckbox",0).Set "ON"
						 	Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnDelete Selected Roles").Click
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Ok").Click
							wait 10
'							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").WaitProperty "Exist",True,30
							Environment("strdeletemsg")=" "
							msgexist=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").Exist(30)
							
							If msgexist=True Then
								if instr(1,Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").GetROProperty("innertext"),"error")<>0 then
									Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("PartialMessage").Click
									wait 2
									Environment("strdeletemsg")=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Messageadescription").GetROProperty("innertext")
									Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Ok").Click
								Else
								Environment("strdeletemsg")=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Message").GetROProperty("innertext")	
								End If
																	
							End If						
							
						 	Exit For
						 End If
						Next
	
						If msgexist=False Then
						
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnClose").Click
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").Link("Event Log & Subscription").Click
							Set appobject=Browser("Hosting Templates - Ops").Page("Clients - Ops System")
							Call menuitemselect(appobject,"Hosting Event Log")
							Set objclientname=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtClientName")
							Call SCA.SetText(objclientname,clientname,"txtClientName","Clients - Ops System")
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtClientName").MiddleClick
							WshShell.SendKeys "{TAB}"
							WshShell.SendKeys "{ENTER}"
							wait 5
							introwcount=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("Template/Client Id").GetROProperty("rows")
							If introwcount>1 Then
								For i=1 to introwcount
									If ucase(Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("Template/Client Id").GetCellData(i,1))=ucase(clientname) Then
									   Environment("strdeletemsg")=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("Template/Client Id").GetCellData(i,4)
									   Exit For
									End If
								Next
	
							End If
						
						End If
						
					End Select
					
				Case "IAM DataSource Permission"
					strDataSource = User
					strRole=Role
							
					Select Case OPSoperation
						Case "Add"
							
							'TODO: Ask Mani to click on "IAM DataSource Permission"
							Set WshShell = CreateObject("WScript.Shell")
							Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtAddDatasource").Click
					        WshShell.SendKeys strDataSource
					        wait 2 ' To wait for List to display as per keyword entered above
					        WshShell.SendKeys "{DOWN}" 
					        wait 2 'To select item from list and wait
					        WshShell.SendKeys "{ENTER}"
					        Set WshShell = Nothing

							Set WshShell = CreateObject("WScript.Shell")
							Set objtxtAddDataSourceRole=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtAddDataSourceRole")
							Call SCA.SetText(objtxtAddDataSourceRole,strRole,"txtAddDataSourceRole","IAM DataSource Permission")
							wait 2
				            WshShell.SendKeys "{ENTER}"
				            wait 1
				            Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtAddDataSourceRole").Click
							WshShell.SendKeys "{ENTER}"
				            Set WshShell = Nothing
							wait 2
							
							set objchkAddDatasourceWithRoles= Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebCheckBox("chkAddDatasourceWithRoles")
							Call SCA.SetCheckBox(objchkAddDatasourceWithRoles,"chkAddDatasourceWithRoles","ON","IAM DataSource Permission")		    
							
							Set objbtnAddDatasourceWithRoles = Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnAddDatasourceWithRoles")
							Call SCA.ClickOn(objbtnAddDatasourceWithRoles,"btnAddDatasourceWithRoles","IAM DataSource Permission")
							
							'Validating "Are you sure you want to add selected Role(s)?"
							If Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Are you sure you want to add selected Roles").Exist(2) then
								Set objbtnOk = Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Ok")
								Call SCA.ClickOn(objbtnOk,"Ok","IAM DataSource Permission")
								
								Call ReportStep (StatusTypes.Pass,"Successfully Added added Datasoure "&strDataSource&" with Role(s) "&strRole&" mapping","IAM DataSource Permission")
							End If
							
						Case "Remove"
							
							Set WshShell = CreateObject("WScript.Shell")
							Set objtxtRemoveDatasourceRoleName= Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtRemoveDatasourceRoleName")
							Call SCA.SetText(objtxtRemoveDatasourceRoleName,strRole, "txtRemoveDatasourceRoleName","IAM DataSource Permission")
							wait 2
				            WshShell.SendKeys "{ENTER}"
				            wait 1
				            Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtRemoveDatasourceRoleName").Click
							WshShell.SendKeys "{ENTER}"
				            wait 2
							
							Set objtxtRemoveDataSourceName = Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtRemoveDataSourceName")
							Call SCA.SetText(objtxtRemoveDataSourceName,strDataSource, "txtRemoveDataSourceName","IAM DataSource Permission")
							wait 2
				            WshShell.SendKeys "{ENTER}"
				            wait 1
				            Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebEdit("txtRemoveDataSourceName").Click
							WshShell.SendKeys "{ENTER}"
				            Set WshShell = Nothing
							wait 2
							
							rc = Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebTable("tabRemoveDatasourceWithRoles").GetROProperty("rows")
							If rc=2 Then
								set objchkRemoveDatasourceWithRoles= Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebCheckBox("chkRemoveDatasourceWithRoles")
								Call SCA.SetCheckBox(objchkRemoveDatasourceWithRoles,"chkRemoveDatasourceWithRoles","ON","IAM DataSource Permission")		    
							
								set objbtnRemoveDatasourceWithRoles = Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnRemoveDatasourceWithRoles")
								Call SCA.ClickOn(objbtnRemoveDatasourceWithRoles,"btnRemoveDatasourceWithRoles","IAM DataSource Permission")
								
								'Validating "Are you sure you want to delete selected Datasoure with Role(s) mapping?"
								If Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebElement("Are you sure you want to delete selected Datasoure").Exist(2) then
									Set objbtnOk = Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("Ok")
									Call SCA.ClickOn(objbtnOk,"Ok","IAM DataSource Permission")
									
									Call ReportStep (StatusTypes.Pass,"Successfully deleted selected Datasoure "&strDataSource&" with Role(s) "&strRole&" mapping","IAM DataSource Permission")
								End If
							else
								Call ReportStep (StatusTypes.Fail,"Could not delete selected Datasoure "&strDataSource&" with Role(s) "&strRole&" mapping","IAM DataSource Permission")
							End If
							
					End Select

'<Shweta - 7/2016> - created btnClose object - Start 
'					Set objClose=Browser("Hosting Templates - Ops").Page("Clients - Ops System").WebButton("btnClose")
'        			Call SCA.ClickOn(objClose,"btnClose","NewClient")
					If objClose.Exist(5) Then
						Call SCA.ClickOn(objClose,"btnClose","NewClient")
					End If
'<Shweta - 7/2016> - created btnClose object - End
					
					Browser("Hosting Templates - Ops").Page("Clients - Ops System").Sync
					Browser("Hosting Templates - Ops").Page("Clients - Ops System").RefreshObject
			End Select
		End If
	End Sub




	'<To create the Report in SCA using the Pivotal table>
    '<sheetNum :- Name of the sheet>
    '<newReport :- to create the new Report or Not>
    '<author :-Shobha>    
    Public Function ReportCreationInSCARajeshcode2(ByVal sheetNum, ByVal newReport,ByVal intEMD,ByVal intdbcheck,ByVal intRNum,ByVal intReportCreate,ByVal strF1,ByVal strF2)  
    
        Dim objExcel,ObjExcelFile,ObjExcelSheet,objDataSource,objDatabase,objDatacubes,obj_tree,objDR,objrow,objcolumn,objDataaxis,objAdd,objtreeCount,objCreate
        Dim intRow_Count,intColumn_Count,intcount
        Dim dataSourceValue, strDataSourceValue,strDatabaseValue,strCubeValue,strChild_Values,strPlaceHolderVal,p, objBtnok, objErrorOk,strDBitems,strbkch
        '<Looping Variables for the row and column>
        Dim i,j,k,x,y,a  
        
        '<shweta 11/9/2015 Closing All opened Excel Files- start
		'objUtils.KillProcess("excel.exe")
		Systemutil.CloseProcessByName "excel.exe"
		'<shweta 11/9/2015 Closing All opened Excel Files- end
        
        Set objExcel = CreateObject("Excel.Application") 
        'objExcel.Visible =True     
    
       ' Set ObjExcelFile = objExcel.Workbooks.Open(Environment.Value("CurrDir")&"InputFiles\IMSSCAWeb\SCAReportSheet.xls")
        Set ObjExcelFile = objExcel.Workbooks.Open(Environment.Value("CurrDir")&Environment.Value("ReportCreationFile"))
        Set ObjExcelSheet = ObjExcelFile.Sheets(sheetNum) 
        intRow_Count = ObjExcelSheet.UsedRange.Rows.Count 
        intColumn_Count = ObjExcelSheet.UsedRange.Columns.Count
    
        dataSourceValue=TRIM(ObjExcelSheet.cells(2,1).value)
        strDataSourceValue = "+ "&dataSourceValue       
        strDatabaseValue=TRIM(ObjExcelSheet.cells(2,2).value)
        strCubeValue=TRIM(ObjExcelSheet.cells(2,3).value)
        
        If intEMD=1 Then
        
         If newReport = 0 Then
            Set objCreate=Browser("Analyzer").Page("Shared_Folder").Frame("view").Image("imgnewreport")        
           'Set objCreate=Browser("Analyzer").Page("Shared_Folder").Image("imgnewreport")-------Ask Shobha
           Call SCA.ClickOn(objCreate,"newreport","Report Creation")    
        End If
              
       'Capture Database Access Error - Start
       If Browser("Analyzer").Frame("ErrorFrame").WebElement("Database access error.").Exist(1) Then
               Set objErrorOk=Browser("Analyzer").Frame("ErrorFrame").WebButton("btnOK")
               Call SCA.ClickOn(objErrorOk,"btnOK" , "Database Access Error during Report Creation")
               Call ReportStep (StatusTypes.Fail,"Not Able to create a report. Database Access Error during Report Creation", "Report Creation Page")    
                    
            'Closing Excel
            ObjExcelFile.Close
            objExcel.Quit    
            Set objExcel=nothing
            Set  ObjExcelFile=nothing
            Set  ObjExcelSheet=nothing
            Set  objtreeCount=nothing
            ReportCreationInSCARajeshcode2 = 1
            Exit Function
       End If
       'Capture Database Access Error - End
       
       'Handled Application Time Out in Report Creation Page - Start
       Set objBtnok = Browser("Analyzer").Dialog("Message from webpage").WinButton("btnOK")
       If objBtnok.Exist(5) Then
            Call SCA.ClickOn(objBtnok,"btnOK" , "Application Time Out in Report Creation Page")
       End If    
       'Handled Application Time Out in Report Creation Page - End
       wait 4
       For a = 1 To 100 Step 1           
       If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDataSources").GetROProperty("disable")=0  Then
         Exit For
       End If           
       Next
       
       wait 2
       
       Set objDataSource=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDataSources")   
       Call SCA.SelectFromDropdown(objDataSource,strDataSourceValue)
       
       wait 2
       
       Set objDatabase=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDatabases")
       Call SCA.SelectFromDropdown(objDatabase,strDatabaseValue)       
       
       If intdbcheck=0 Then   
       
        strDBitems=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDatabases").GetROProperty("all items")
         If Instr(strDBitems,strDatabaseValue)<>0 Then           
            ReportCreationInSCARajeshcode2=0
               else
            ReportCreationInSCARajeshcode2=1
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
       
       
       Set objDatacubes=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlCubes")
       Call SCA.SelectFromDropdown(objDatacubes,strCubeValue)
       
       Wait 2
   	   Set objAdd=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebButton("btnAdd")
   	   Call SCA.ClickOn(objAdd,"AddButton" , "ReportCreation")
       
       
       
    
       For i=2 to intRow_Count
        Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcube").Click           
         For j=4 to intColumn_Count-1
              strChild_Values=Trim(ObjExcelSheet.cells(i,j).value)
               ' Have to write the explicit mentain the bk mark condition
               strbkch=Trim(ObjExcelSheet.cells(i,intColumn_Count).value)
               If  strChild_Values<>"" AND  strChild_Values<>"rowaxis" AND strChild_Values<>"columnaxis"  AND strChild_Values<>"dataaxis" AND strbkch<>"Yes" AND strChild_Values<>"No" Then
                        Set obj_tree=description.Create
                        obj_tree("micclass").value="WebElement"
                        obj_tree("class").value="TreeNode"
                        obj_tree("innertext").value=strChild_Values
                        wait 1
                        set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
                        intcount=objtreeCount.count
                        
                        For a = 1 To 100 Step 1
	                    set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
	                    intcount=objtreeCount.count
                            If intcount<>0 Then
                            Exit For
                            End If
                        Next
                        
                        If intcount<>0 Then
                        '' ch - 25 - 001 - May - Rajesh - Start 
'                        objtreeCount(0).fireEvent  "ondblclick" 
                        	'shweta - 31/3/2016 - start
                        	If intcount = 2 AND strF1 <> "" Then
                        		k=intcount-2                    
                                objtreeCount(k).fireEvent  "ondblclick"  	
                        	
                        	ElseIf intcount = 2 AND strChild_Values <> "Additional attributes" Then
                        		k=intcount-2                    
                                objtreeCount(k).fireEvent  "ondblclick"  	
                        	Else
								k=intcount-1                    
                                objtreeCount(k).fireEvent  "ondblclick"  
                        	End If
                        	'' ch - 25 - 001 - May - Rajesh - End..
 
							'shweta - 31/3/2016 - End                            
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
                       
                    wait 5
                    
'                    Set objDR=CreateObject("Mercury.DeviceReplay")
'                    x=objtreeCount(k).GetroProperty("abs_x")
'                    y=objtreeCount(k).GetroProperty("abs_y")  
'                    wait 1                    
'                    objDR.MouseClick x,y,2

'obj_tree("micclass").value="WebElement"
'                            obj_tree("class").value="TreeNode"
'                            obj_tree("innertext").value=strChild_Values
'
'    Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("class:=TreeNode","index:=1","innertext:="&strChild_Values).RightClick
                    Setting.WebPackage("ReplayType") = 2
                    objtreeCount(k).RightClick
                    Setting.WebPackage("ReplayType") = 1
                    
                    strPlaceHolderVal=lcase(ObjExcelSheet.cells(i,intColumn_Count-1).value)                                                                                        
                    wait 1  
                      If strPlaceHolderVal="rowaxis" Then 
                      	  	
                          Set objrow=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webRowaxis")    
                          Call SCA.ClickOn(objrow,"Rowaxis" , "ReportCreation")    
                          wait 5                          
                      elseif  StrPlaceHolderVal="columnaxis" then  
						                       
                          Set objcolumn=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcolumnaxis")
                          Call SCA.ClickOn(objcolumn,"Columnaxis" , "ReportCreation")
                          Wait 5                      
                      elseif  StrPlaceHolderVal="dataaxis" then
						                    
                         Set objDataaxis=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webDataAxis")
                         Call SCA.ClickOn(objDataaxis,"Dataaxis" , "ReportCreation")    
                         wait 5
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
                            set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
                            intcount=objtreeCount.count    
                                If intcount<>0 Then
                                
                                   If intcount = 2 AND strF1 <> "" Then
		                        		k=intcount-2                    
		                                objtreeCount(k).fireEvent  "ondblclick"  	
	
                                   ElseIf intcount = 2 AND strChild_Values <> "Additional attributes" Then
                                   	  k=intcount-2
                                      objtreeCount(k).fireEvent  "ondblclick" 
								   Else
									   k=intcount-1
                                       objtreeCount(k).fireEvent  "ondblclick"    
                                   End If
                                   
'                                   k=intcount-1
'                                   objtreeCount(k).fireEvent  "ondblclick"    
                                 End If                          
                            End If
                        Next
                
                End If
             
           Next
    
           Next
            
   
       
   ElseIf intEMD=0 Then
           
'           Set objExcel = CreateObject("Excel.Application") 
'        objExcel.Visible =True     
'        
'        Set ObjExcelFile = objExcel.Workbooks.Open(Environment.Value("CurrDir")&"InputFiles\IMSSCAWeb\SCAReportSheet.xls")
'        Set ObjExcelSheet = ObjExcelFile.Sheets(intsheetName)
'        
'        intRow_Count = ObjExcelSheet.UsedRange.Rows.Count 
'        intColumn_Count = ObjExcelSheet.UsedRange.Columns.Count   
        Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcube").Click          
        For j=4 to intColumn_Count-2
        strChild_Values=Trim(ObjExcelSheet.cells(intRNum,j).value)
               
          If  strChild_Values<>"" AND  strChild_Values<>"rowaxis" AND strChild_Values<>"columnaxis"  AND strChild_Values<>"dataaxis" AND strChild_Values<>"Yes" AND strChild_Values<>"No" Then
        
               Set obj_tree=description.Create
               obj_tree("micclass").value="WebElement"
               obj_tree("class").value="TreeNode"
               obj_tree("innertext").value=strChild_Values
               wait 1
               set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
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
              
'              Set objDR=CreateObject("Mercury.DeviceReplay")
'               x=objtreeCount(k).GetroProperty("abs_x")
'               y=objtreeCount(k).GetroProperty("abs_y") 
'			   objtreeCount(k).fireEvent  "ondblclick"	               
'                wait 1                    
'                objDR.MouseClick x,y,2
                Setting.WebPackage("ReplayType") = 2
              	objtreeCount(k).RightClick
              	Setting.WebPackage("ReplayType") = 1
               strPlaceHolderVal=lcase(ObjExcelSheet.cells(intRNum,intColumn_Count-1).value)                                                                                        
                      
                      If strPlaceHolderVal="rowaxis" Then 
                          Set objrow=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webRowaxis")    
                          Call SCA.ClickOn(objrow,"Rowaxis" , "ReportCreation")    
                          wait 5                          
                      elseif  StrPlaceHolderVal="columnaxis" then                                                                                                                                
                          Set objcolumn=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcolumnaxis")
                          Call SCA.ClickOn(objcolumn,"Columnaxis" , "ReportCreation")
                          Wait 5                      
                      elseif  StrPlaceHolderVal="dataaxis" then                                                                            
                         Set objDataaxis=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webDataAxis")
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
                            set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
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


	
	'<Composed By : Shweta B Nagaral>
	'<Saves SCA Report in Excel format to local machine and returns local path where report has been saved
	'It validates Report Name, Sheet Name and Excel format in while saving SCA report. 
	'SCA Report name should be same as Excel file name>
	'<ExcelName: Name of the excel while saving: String>
	Public Function SaveExcelReportRajesh(ByVal ExcelName)
	
	Dim i,toolBarNameFormat,strReport,a,b,j,strReportName,strReportFormat,path,formatPos,myVar,objLstWinButton ,objbtnSave,objYes,k,objClose,objbtnCloseExport, midVal
	midVal = len(ExcelName)+5
	    
	    'Report saving
	      For i = 1 To 50 Step 1
	       If Browser("Analyzer").WinObject("Notification bar").WinButton("lstWinButton").Exist(2) then
	       Exit For
	       End If
	      Next
	      wait 2                
	      toolBarNameFormat = Browser("Analyzer").WinObject("Notification bar").GetROProperty("text")
	      'Retreiving name and format of report in Notification tool bar
	      'strReport Splitting
	      strReport = mid(toolBarNameFormat, InStr(1, Trim(Ucase(toolBarNameFormat)), Trim(Ucase(ExcelName))), midVal)
	           a=Split(Trim(strReport), ".")
	           b= UBound(a)
	            For j = 0 To b Step 2
	              strReportName = a(j)
	              strReportFormat = "."&a(j+1)
	            Next
	                
	       path = Environment.Value("CurrDir")&"Exportresults\"&Trim(strReport)
	    
	       formatPos = InStr(Trim(Ucase(toolBarNameFormat)), Trim(Ucase(strReportFormat)))
	           If formatPos > 0  Then
	             myVar = Mid(toolBarNameFormat, formatPos, 5 )
	                     
	              If strReportFormat = ".xlsx" Then
	                If StrComp(Trim(UCase(strReportFormat)), Trim(Ucase(myVar))) = 0 Then
	                   Call ReportStep (StatusTypes.Pass, "Saving Excel Report " &strReportName& " with Report Format " &strReportFormat, "Export Options Page")
	                   ElseIf StrComp(Trim(UCase(strReportFormat)), Trim(Ucase(myVar))) = 1 Then  
	                       
	                   Call ReportStep (StatusTypes.Fail, "Saving Excel Report " &strReportName& " with Report Format " &strReportFormat, "Export Options Page")
	                 End If
	                                                
	                Else
	                    If StrComp(Trim(UCase(strReportFormat)), Trim(Ucase(myVar))) = 0 Then                   
	                    Call ReportStep (StatusTypes.Pass, "Saving Excel Report " &strReportName& " with Report Format " &strReportFormat, "Export Options Page")
	                    ElseIf StrComp(Trim(UCase(strReportFormat)), Trim(Ucase(myVar))) = -1 Then                   
	                    Call ReportStep (StatusTypes.Fail, "Saving Excel Report " &strReportName& " with Report Format " &strReportFormat, "Export Options Page")
	                    End If
	                    End If
	                                
	                ElseIf formatPos = 0 Then
	                   'Format Mismtach
	                   Call ReportStep (StatusTypes.Fail, "Saving Excel Report " &strReportName& " with Report Format " &strReportFormat, "Export Options Page")
	                   Call ReportStep (StatusTypes.Fail, "Format Mismtach. Expected Excel Report format is " &strReportFormat& " But Notification tool bar is showing --> "&toolBarNameFormat& " Report format", "Export Options Page")
	             End If
	               
	                Set objLstWinButton=Browser("Analyzer").WinObject("Notification bar").WinButton("lstWinButton")
	                Call SCA.ClickOn(objLstWinButton,"lstWinButton", "Excel Report Export Dialouge")
	                
	                'Browser("Analyzer").WinMenu("ContextMenu").Select "Save as"
					Browser("Analyzer").WinMenu("SaveAsContextMenu").WaitProperty "visible",True,10000
					Browser("Analyzer").WinMenu("SaveAsContextMenu").Select "Save as"
	                wait 2
	                
	                abc =  Browser("Analyzer").Dialog("Save As").WinToolbar("WinPath").GetROProperty("regexpwndtitle")
					abd = Mid(abc,10)
	                Paths = abd&"\"&Trim(strReport)

	                
	              
'	               	Browser("Analyzer").Dialog("Save As").WinToolbar("WinPath").WaitProperty "visible",True,10000
'	                Browser("Analyzer").Dialog("Save As").WinToolbar("WinPath").Set path
	              
	                wait 2
	                Set objbtnSave=Browser("Analyzer").Dialog("Save As").WinButton("btnSave")
	                objbtnSave.WaitProperty "visible",True,10000
	                Call SCA.ClickOn(objbtnSave,"btnSave", "Excel Report Export Dialouge")
	                wait 2
	                
	                IF Browser("Analyzer").Dialog("Save As").Dialog("Confirm Save As").WinButton("Yes").Exist(5) Then
	                    Set objYes=Browser("Analyzer").Dialog("Save As").Dialog("Confirm Save As").WinButton("Yes")
	                    Call SCA.ClickOn(objYes,"Yes", "Excel Report Export Dialouge")
	                    wait 2
	                End If
	                
	                For k = 1 To 3 Step 1
	                   If Browser("Analyzer").WinObject("Notification bar").WinButton("Close").Exist(1) Then
	                      Set objClose=Browser("Analyzer").WinObject("Notification bar").WinButton("Close")
	                       Call SCA.ClickOn(objClose,"Close", "Excel Report Export Dialouge, Notification Bar")
	                       Exit For
	                    End If
	                Next
	                
	                wait 2
	                Set objbtnCloseExport=Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebButton("btnClose")
	                objbtnCloseExport.WaitProperty "Visible",True,10000
	                Call SCA.ClickOn(objbtnCloseExport,"CloseExportbtn", "Excel Report Export Dialouge")
	               
	              
	    SaveExcelReportRajesh=Paths
	End Function


	Public Function RollingSumYTD()
	
	Dim rc,cc,j,i,intMeasure,intMeasYTDSUM,intMeasPrevYTDSum,intMeasureSum,intMeasCompYTDSum
	    Set obj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("class:=pivot_table","name:=cube").ChildItem(2,2,"WebTable",0)
		rc =obj.GetROProperty("rows")
		cc=obj.GetROProperty("cols")
		'rc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabRollingSum").GetROProperty("rows")
		'cc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabRollingSum").GetROProperty("cols")
		
		cc=4
		
		If rc<>0 and cc<>0 Then
		    For j = 1 To cc Step 2
		        For i = 1 To rc Step 1
		        	
		        	Call ReportStep (StatusTypes.Information,"Row Number: "&i&"Col Number: "&j, "Pivot table")
		        	
		           
		            intMeasure = obj.GetCellData(i, j)
		            intMeasYTDSUM= obj.GetCellData(i,j+1)    
		            
		            If intMeasYTDSUM = "-" and i = rc Then
		            
		            	Call ReportStep(StatusTypes.Pass, "No YTD Sum Value calulation required. Reached End of the table. Total sum values of Col "&j&" is "&intMeasure, "Pivot table")
		                Exit for
		            ElseIf intMeasYTDSUM = "-" Then
		            	
		            	Call ReportStep(StatusTypes.Fail, "No YTD Sum Value found", "SCA Home Page")
		            
		            End If
		            
		            If cDbl(intMeasure) = cDbl(intMeasYTDSUM) Then
		                
		               	Call ReportStep(StatusTypes.Pass,"Measure and YTD Sum values are same"&intMeasure, "Pivot table") 
		                
		               


		            else
		                    If i<>1 Then
		                        intMeasPrevYTDSum = obj.GetCellData(i-1, j+1)
		                        intMeasureSum = cDbl(intMeasure)+cDbl(intMeasPrevYTDSum)
		                        intMeasCompYTDSum=obj.GetCellData(i, j+1)
		                        
		                        If Round(cDbl(intMeasureSum), 4) = Round(cDbl(intMeasCompYTDSum),4) Then
		                        
		                        	Call ReportStep(StatusTypes.Pass,"cell data value at row i, j and i, j+1 are same", "Pivot table")
		                            
		          
		                        Else
		                          
		                            Call ReportStep(StatusTypes.Fail, "Fail mismatch in Measure sum value at row i, j", "Pivot table")
		                            
		                        End If
		                        
		                    ElseIf i=1 Then
		                        Call ReportStep(StatusTypes.Fail, "cell data value at row i, j and i, j+1 are not same", "Pivot table")
		                        
		                    End If
		            End If        
		        Next
		    Next
		    
		End If
		
		
		wait 2	
			
		
	End Function 
		
	
	Public Sub RowDimensionDownNormal(ByVal ColVal,ByVal strTabName, ByVal strSelValue, ByVal strSelSubValue)
		
		Dim oDimensionDownNormal, oDimensionDownNormalObj, objSelSubValue, objSelValue
	
		Set oDimensionDownNormal = Description.Create()
		oDimensionDownNormal("micclass").value = "Image"
		oDimensionDownNormal("html id").value = ""
		oDimensionDownNormal("html tag").value = "IMG"
		oDimensionDownNormal("image type").value = "Plain Image"
		oDimensionDownNormal("outerhtml").regularexpression=True
		oDimensionDownNormal("outerhtml").value = ".*onLevelContextClick.*"
		Set obj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("class:=pivot_table","name:=cube").ChildItem(1,1,"WebTable",0)
		'Set oDimensionDownNormalObj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDropRowDimension").ChildObjects(oDimensionDownNormal)
		Set oDimensionDownNormalObj = obj.ChildObjects(oDimensionDownNormal)		
		Browser("Analyzer").Page("ReportCreation").Sync
		'Mouse over on Pivot Table menu
		oDimensionDownNormalObj(ColVal).highlight
		oDimensionDownNormalObj(ColVal).FireEvent "onmouseover"
		wait 2
		oDimensionDownNormalObj(ColVal).Click
		wait 2
		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image(strTabName).FireEvent "onmouseover"
		wait 1
		
		'First part of if condition will be executed, if sub-options has to be selected
		If strSelSubValue <> "" Then
			Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strSelValue).FireEvent "onmouseover"
			wait 1
			Set objSelSubValue = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strSelSubValue)
			Call SCA.ClickOn(objSelSubValue,strSelSubValue, "Report Creation Page")
		Else
			'Clicks on options options under each tab
			Set objSelValue = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strSelValue)
			Call SCA.ClickOn(objSelValue,strSelValue, "Report Creation Page")
		End If
		wait 2
		
	End Sub

	
	
	
	
	Public Function RollingSumQTDOld()
		
	Dim rc,cc,i,j,intMeasure,intMeasQTDSUM,intMeasPrevQTDSum,intMeasureSum,intMeasCompQTDSum
		
		
		
	rc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabRollingSum").GetROProperty("rows")
	cc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabRollingSum").GetROProperty("cols")

	cc = 4
If rc<>0 and cc<>0 Then
    For j = 1 To cc Step 2
        For i = 1 To rc Step 1
        
        	Call ReportStep (StatusTypes.Information,"Row Number: "&i&"Col Number: "&j, "Pivot table")
            intMeasure = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabRollingSum").GetCellData(i, j)
            intMeasQTDSUM= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabRollingSum").GetCellData(i,j+1)    
            
            If intMeasQTDSUM = "-" and i = rc Then
            
            	Call ReportStep(StatusTypes.Pass, "No QTD Sum Value calulation required. Reached End of the table. Total sum values of Col "&j&" is "&intMeasure, "Pivot table")
                
                Exit for
            ElseIf intMeasQTDSUM = "-" Then
            
            	Call ReportStep(StatusTypes.Fail, "No QTD Sum Value found", "SCA Home Page")
                
            End If
            
            If cDbl(intMeasure) = cDbl(intMeasQTDSUM) Then
            
            		Call ReportStep(StatusTypes.Pass,"Measure and QTD Sum values are same"&intMeasure, "Pivot table") 
                   
            else
                    If i<>1 Then
                        intMeasPrevQTDSum = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabRollingSum").GetCellData(i-1, j+1)
                        intMeasureSum = cDbl(intMeasure)+cDbl(intMeasPrevQTDSum)
                        intMeasCompQTDSum=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabRollingSum").GetCellData(i, j+1)
                        
                        If Round(cDbl(intMeasureSum), 4) = Round(cDbl(intMeasCompQTDSum),4) Then
                        
                        	Call ReportStep(StatusTypes.Pass,"cell data value at row i, j and i, j+1 are same", "Pivot table")
                            
                        Else
                        
                        	Call ReportStep(StatusTypes.Fail, "Fail mismatch in Measure sum value at row i, j", "Pivot table")
                        
                            
                        End If
                        
                    ElseIf i=1 Then
                    
                    	Call ReportStep(StatusTypes.Fail, "cell data value at row i, j and i, j+1 are not same", "Pivot table")
                        
                    End If
            End If        
        Next
    Next
    
End If


wait 2

		
		
		
	End Function







Public Function RollingSumQTD()
		
	Dim rc,cc,i,j,intMeasure,intMeasQTDSUM,intMeasPrevQTDSum,intMeasureSum,intMeasCompQTDSum
		
		
		
	rc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrpTableMeasureData1").GetROProperty("rows")
	cc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrpTableMeasureData1").GetROProperty("cols")

	cc = 4
If rc<>0 and cc<>0 Then
    For j = 1 To cc Step 2
        For i = 1 To rc Step 1
        
        	Call ReportStep (StatusTypes.Information,"Row Number: "&i&"Col Number: "&j, "Pivot table")
            intMeasure = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrpTableMeasureData1").GetCellData(i, j)
            intMeasQTDSUM= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrpTableMeasureData1").GetCellData(i,j+1)    
            
            If intMeasQTDSUM = "-" and i = rc Then
            
            	Call ReportStep(StatusTypes.Pass, "No QTD Sum Value calulation required. Reached End of the table. Total sum values of Col "&j&" is "&intMeasure, "Pivot table")
                
                Exit for
            ElseIf intMeasQTDSUM = "-" Then
            
            	Call ReportStep(StatusTypes.Fail, "No QTD Sum Value found", "SCA Home Page")
                
            End If
            
            If cDbl(intMeasure) = cDbl(intMeasQTDSUM) Then
            
            		Call ReportStep(StatusTypes.Pass,"Measure and QTD Sum values are same"&intMeasure, "Pivot table") 
                   
            else
                    If i<>1 Then
                        intMeasPrevQTDSum = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrpTableMeasureData1").GetCellData(i-1, j+1)
                        intMeasureSum = cDbl(intMeasure)+cDbl(intMeasPrevQTDSum)
                        intMeasCompQTDSum=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrpTableMeasureData1").GetCellData(i, j+1)
                        
                        If Round(cDbl(intMeasureSum), 4) = Round(cDbl(intMeasCompQTDSum),4) Then
                        
                        	Call ReportStep(StatusTypes.Pass,"cell data value at row i, j and i, j+1 are same", "Pivot table")
                            
                        Else
                        
                        	Call ReportStep(StatusTypes.Fail, "Fail mismatch in Measure sum value at row i, j", "Pivot table")
                        
                            
                        End If
                        
                    ElseIf i=1 Then
                    
                    	Call ReportStep(StatusTypes.Fail, "cell data value at row i, j and i, j+1 are not same", "Pivot table")
                        
                    End If
            End If        
        Next
    Next
    
End If


wait 2

		
		
		
	End Function














	
	
	Public Function RollingAVGYTD()
		
	Dim rc,cc,divide,i,j,intMeasure,intMeasYTDAVG,intMeasPrevYTDAVG,avj,intMeasureAVG,intMeasCompYTDAVG	
	Set obj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("class:=pivot_table","name:=cube").ChildItem(2,2,"WebTable",0)
	cc =obj.GetROProperty("cols")	
	rc	=obj.GetROProperty("rows")
	'	rc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabRollingSum").GetROProperty("rows")
'cc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabRollingSum").GetROProperty("cols")

divide = 0
cc = 4
If rc<>0 and cc<>0 Then
    For j = 1 To cc Step 2
        For i = 1 To rc Step 1
        	
        	Call ReportStep (StatusTypes.Information,"Row Number: "&i&"Col Number: "&j, "Pivot table")           
            intMeasure = obj.GetCellData(i, j)
            intMeasYTDAVG= obj.GetCellData(i,j+1)    
            
            If intMeasYTDAVG = "-" and i = rc Then
            
            	Call ReportStep(StatusTypes.Pass, "No YTD AVG Value calulation required. Reached End of the table. Total sum values of Col "&j&" is "&intMeasure, "Pivot table")
                
                Exit for
            ElseIf intMeasYTDAVG = "-" Then
            
            	Call ReportStep(StatusTypes.Fail, "No YTD AVG Value found", "SCA Home Page")
               
            End If
            
            If cDbl(intMeasure) = cDbl(intMeasYTDAVG) Then
            
            		Call ReportStep(StatusTypes.Pass,"Measure and YTD AVG values are same"&intMeasure, "Pivot table")                
                    divide =1
                    
            else
                    If i<>1 Then
                    
                    	divide = divide+1
                    	
                        intMeasPrevYTDAVG = obj.GetCellData(i-1, j+1)
                        
                        avj = (intMeasPrevYTDAVG*(divide-1))
                        
                        intMeasureAVG = (cDbl(intMeasure)+cDbl(avj))/divide
                        intMeasCompYTDAVG=obj.GetCellData(i, j+1)
                        
                        If Round(cDbl(intMeasureAVG),4) = Round(cDbl(intMeasCompYTDAVG),4) Then
                        
                        	Call ReportStep(StatusTypes.Pass,"cell data value at row i, j and i, j+1 are same", "Pivot table")
                           
                        Else
                        
                        	Call ReportStep(StatusTypes.Fail, "Fail mismatch in Measure AVG value at row i, j", "Pivot table")
                            
                        End If
                        
                    ElseIf i=1 Then
                    
                    	Call ReportStep(StatusTypes.Fail, "cell data value at row i, j and i, j+1 are not same", "Pivot table")
                        
                    End If
            End If        
        Next
    Next
    
End If


wait 2
			
		
	End Function
	
	
	
	
	
	
	

Public Function RollingAVGQTD()
		
	Dim rc,cc,divide,i,j,intMeasure,intMeasYTDAVG,intMeasPrevYTDAVG,avj,intMeasureAVG,intMeasCompYTDAVG	
		
		
		rc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrpTableMeasureData1").GetROProperty("rows")
cc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrpTableMeasureData1").GetROProperty("cols")

divide = 0
cc = 4
If rc<>0 and cc<>0 Then
    For j = 1 To cc Step 2
        For i = 1 To rc Step 1
        	
        	Call ReportStep (StatusTypes.Information,"Row Number: "&i&"Col Number: "&j, "Pivot table")           
            intMeasure = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrpTableMeasureData1").GetCellData(i, j)
            intMeasYTDAVG= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrpTableMeasureData1").GetCellData(i,j+1)    
            
            If intMeasYTDAVG = "-" and i = rc Then
            
            	Call ReportStep(StatusTypes.Pass, "No QTD AVG Value calulation required. Reached End of the table. Total sum values of Col "&j&" is "&intMeasure, "Pivot table")
                
                Exit for
            ElseIf intMeasYTDAVG = "-" Then
            
            	Call ReportStep(StatusTypes.Fail, "No QTD AVG Value found", "SCA Home Page")
               
            End If
            
            If cDbl(intMeasure) = cDbl(intMeasYTDAVG) Then
            
            		Call ReportStep(StatusTypes.Pass,"Measure and QTD AVG values are same"&intMeasure, "Pivot table")                
                    divide =1
                    
            else
                    If i<>1 Then
                    
                    	divide = divide+1
                    	
                        intMeasPrevYTDAVG = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrpTableMeasureData1").GetCellData(i-1, j+1)
                        
                        avj = (intMeasPrevYTDAVG*(divide-1))
                        
                        intMeasureAVG = (cDbl(intMeasure)+cDbl(avj))/divide
                        intMeasCompYTDAVG=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrpTableMeasureData1").GetCellData(i, j+1)
                        
                        If Round(cDbl(intMeasureAVG),4) = Round(cDbl(intMeasCompYTDAVG),4) Then
                        
                        	Call ReportStep(StatusTypes.Pass,"cell data value at row i, j and i, j+1 are same", "Pivot table")
                           
                        Else
                        
                        	Call ReportStep(StatusTypes.Fail, "Fail mismatch in Measure AVG value at row i, j", "Pivot table")
                            
                        End If
                        
                    ElseIf i=1 Then
                    
                    	Call ReportStep(StatusTypes.Fail, "cell data value at row i, j and i, j+1 are not same", "Pivot table")
                        
                    End If
            End If        
        Next
    Next
    
End If


wait 2
			
		
	End Function


	Public Function RowTotalRank()
		
		Dim arr(),rc,cc,count,ColSum,ColRank,SalesValue,SalesValue1,k,m,tmp,z,p,abcd,abcde,abcdef,j,i,Cnt
		
		'find the row and column count from pivot table
		rc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabRowTotalRankValidation").GetROProperty("rows")
		cc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabRowTotalRankValidation").GetROProperty("cols")
		
		cc = 4
		
		'redefine the array size
		ReDim arr(rc-2)
		
		
		If rc<>0 AND cc<>0 Then	
		
			
			For j = 1 To cc-1 Step 2
				count = 1    
					'Stores the column value in an array
		    		For i = 1 To rc Step 1
		        	
		        		If (i = rc) Then
		        		
			        		ColSum = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabRowTotalRankValidation").GetCellData(i,j)
			        		ColRank = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabRowTotalRankValidation").GetCellData(i,j+1)
			        		
			        		Call ReportStep (StatusTypes.Information,"No rank verification required. Reached End of the table. Total sum values of Col "&j&" is "&ColSum&" and rank is"&ColRank,"Pivot table")     		
							Exit For
		        		End If
		        	
		        	
				        SalesValue = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabRowTotalRankValidation").GetCellData(i,j)
				        SalesValue1 = Replace(SalesValue,",","")
				        arr(i-1) = SalesValue1
			        
		    		Next
		    
		    		'Make array in ascending order
			        Cnt = UBound(arr) 
			        For k = 0 To Cnt-1 Step 1
			        For m = 0 To Cnt-1-k Step 1
			        If arr(m+1) < arr(m) Then 
			        tmp = arr(m) 
			        arr(m) = arr(m+1) 
			        arr(m+1) = tmp 
			        End If 
			        Next 
			        Next
		    
		            
		                For z = ubound(arr) To 0 Step -1
		                
		
		                
		                    For p = 1 To rc Step 1
		                        
		                        abcd = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabRowTotalRankValidation").GetCellData(p,j)
		                    
		                        abcde = replace(abcd,",","")    
		                        
		                        If arr(z) = abcde Then
		                            
		                            abcdef = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabRowTotalRankValidation").GetCellData(p,j+1)
		                            
		                            If Cint(abcdef) = count Then
		                            	
		                            	Call ReportStep(StatusTypes.Pass, "rank validated for row "&p&" and column "&j&", sales value "&arr(z)&" and sales rank is" &count , "Pivot table")
		                            		                           
		                                Count = count+1
		                                Exit For
		                             Else

										Call ReportStep(StatusTypes.Fail, "Could not successfully validate the rank for row "&p&" and column "&j&", sales value "&arr(z)&" and sales rank is" &count, "Pivot table")
		                                
		                            End If    
		                            
		'                            Count = count+1
		'                            Exit For
		                            
		                        End If
		                        
		                    Next
		                
		                Next
		    
		    
		    
			Next
		
		End If



		
		
	End Function
	
	
	
	
	
	
	
	
	
	
	
	
	Public Function YearlyGrowthRate()
		
		Dim rc,cc,y,z,i,x,abc,abd,abcd,abe,varr,abede,abedef,Count,count1,Firstcolumn,m,n,o,p
		
		'Find the row and column count
		rc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").RowCount
		cc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").ColumnCount (1)
		
		cc = 4
		
		'First column yearly growth rate validation
		
		For m = 1 To 1 Step 1
			
			For n = 1 To rc Step 1
			
			o = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").GetCellData(n,m)
				
			p = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").GetCellData(n,m+1)
			
				If (p = "-") Then
					
					Call ReportStep(StatusTypes.Pass, "Yearly growth rate is null for first column, for row "&n&" and column "&m&" is"&p, "Pivot table")
				Else

					Call ReportStep(StatusTypes.Fail,"Yearly growth rate is not null for first column, for row "&n&" and column "&m&" is"&p, "Pivot table")
					
					
				End If	
			
			Next	
			
			
		Next
		
		
		'yearly growth rate validation
		
		y = 3 
		z = 4
		
		For i = 1 To cc-3 Step 2
			
			For x = 1 To rc Step 1
				
				abc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").GetCellData(x,i)
				
		'		y= 3
				
				abd = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").GetCellData(x,y)
				
				
				If abc = "0.0000" OR abc = "-" OR abd = "-" Then
					
				abcd = "-"	
					
'				ElseIf (z = cc  ) Then
				
'				abcd = "-"	
		
				Else
				
				abcd = Round((abd/abc*100-100),4)&" %"
					
				End If
				
		'		abcd = Round((abd/abc*100-100),4)&" %"
				
		'		abcde = Round(abcd,4)&" %"
			
		'		z=4
				
				abe = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").GetCellData(x,z)
				
				If abe <> "-" Then
					
					varr = split(abe," ")
				
					abede = Round(varr(0),4)
				
					abedef = replace(abede,",","")&" %"
				
				Else
				
					abedef = abe
					
				End If
				
				
				
				
				
				If (abedef = abcd) Then
				
				Call ReportStep(StatusTypes.Pass, "Yearly growth rate is display correctly for row "&x&" and column "&i&" is"&abedef, "Pivot table")
				
		
		
				Else
				
				Call ReportStep(StatusTypes.Fail, "Yearly growth rate is  not display correctly for row "&x&" and column "&i&" is"&abedef, "Pivot table")
				
				

				End If
				
				wait 2
			Next
			
			y = y+2
			
			z = z+2
			
			
		Next
		
		

	
		
	End Function
	
	
	
	
	
	Public Sub RMeasureDownNormal(ByVal ColVal, ByVal strTabName, ByVal strSelValue, ByVal strSelSubValue)
		
		Dim oMeasureDownNormal, oMeasureDownNormalObj, objSelSubValue, objSelValue
	
		Set oMeasureDownNormal = Description.Create()
		oMeasureDownNormal("micclass").value = "Image"
		oMeasureDownNormal("html id").value = ""
		oMeasureDownNormal("html tag").value = "IMG"
		oMeasureDownNormal("image type").value = "Plain Image"
		oMeasureDownNormal("outerhtml").regularexpression=True
		oMeasureDownNormal("outerhtml").value = ".*onMeasureContextClick.*"
		
		Set oMeasureDownNormalObj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrowthRateDropdowns").ChildObjects(oMeasureDownNormal)
				
		Browser("Analyzer").Page("ReportCreation").Sync
		'Mouse over on Pivot Table menu
		oMeasureDownNormalObj(ColVal).FireEvent "onmouseover"
		wait 2
		oMeasureDownNormalObj(ColVal).Click
		wait 2
		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image(strTabName).FireEvent "onmouseover"
		wait 1
		
		'First part of if condition will be executed, if sub-options has to be selected
		If strSelSubValue <> "" Then
			Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strSelValue).FireEvent "onmouseover"
			wait 1
			Set objSelSubValue = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strSelSubValue)
			Call SCA.ClickOn(objSelSubValue,strSelSubValue, "Report Creation Page")
		Else
			'Clicks on options options under each tab
			Set objSelValue = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strSelValue)
			Call SCA.ClickOn(objSelValue,strSelValue, "Report Creation Page")
		End If
		wait 2
		
	End Sub
	
	
	
	Public Function YearlyGrowth()
		
		Dim rc,cc,y,z,i,x,abc,abd,abcd,pro,abe,abede,abedef,m,n,o,p
		
	
		
		'To find the row and column count
		rc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").RowCount
		cc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").ColumnCount (1)
	
		cc = 4



		'First column yearly growth validation
		
		For m = 1 To 1 Step 1
			
			For n = 1 To rc Step 1
			
			o = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").GetCellData(n,m)
				
			p = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").GetCellData(n,m+1)
			
				If (p = o) Then
					
					Call ReportStep(StatusTypes.Pass, "Sales value and Yearly growth is equal for first column, for row "&n&" and column "&m&" is"&p, "Pivot table")
				Else

					Call ReportStep(StatusTypes.Fail,"Sales value and Yearly growth is not equal for first column, for row "&n&" and column "&m&" is"&p, "Pivot table")
					
					
				End If	
			
			Next	
			
			
		Next
		





		y = 3 
		z = 4
		
		For i = 1 To cc-3 Step 2
			
			For x = 1 To rc Step 1
				
				abc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").GetCellData(x,i)
				
		'		y= 3
				
				abd = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").GetCellData(x,y)
				
				
				If ((abc = "-") AND (abd = "0.0000")) Then
					
				abcd = Round(abd,4)	
					
				ElseIf abc = "-" OR abd = "-" Then
					
				abcd = "-"	
					
'				ElseIf (z = cc ) Then
				
'				abcd = Round(abd,4)
				
		'		abcd = "-"	
		
				Else
				
				abcd = Round((abd-abc),4)
					
				End If
				
'				If abc = "-" OR abd = "-" Then
'					
'				abcd = "-"	
'					
'				ElseIf (z = cc ) Then
'				
'				abcd = Round(abd,4)
'				
'		'		abcd = "-"	
'		
'				Else
'				
'				abcd = Round((abd-abc),4)
'					
'				End If
'				
				pro = Cstr(abcd)
				
		'		abcd = Round((abd/abc*100-100),4)&" %"
				
		'		abcde = Round(abcd,4)&" %"
			
		'		z=4
				
				abe = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").GetCellData(x,z)
				
				If abe <> "-" Then
					
		'			varr = split(abe," ")
				
					abede = Round(abe,4)
				
					abedef = replace(abede,",","")
				
				Else
				
					abedef = abe
					
				End If
				
				
				
				
				
				If (abedef = pro) Then
				
				Call ReportStep(StatusTypes.Pass, "Yearly growth  is display correctly for row "&x&" and column "&i&" is"&abedef, "Pivot table")
				
				Else
				
				Call ReportStep(StatusTypes.Fail, "Yearly growth is  not display correctly for row "&x&" and column "&i&" is"&abedef, "Pivot table")
				
				
				End If
				
				wait 2
			Next
			
			y = y+2
			
			z = z+2
			
			
		Next
		
	
		
		
	End Function
	
	
	
	
	
	
	
	
	
	
	Public Sub RGMeasureDownNormal(ByVal ColVal, ByVal strTabName, ByVal strSelValue, ByVal strSelSubValue)
		
		Dim oMeasureDownNormal, oMeasureDownNormalObj, objSelSubValue, objSelValue
	
		Set oMeasureDownNormal = Description.Create()
		oMeasureDownNormal("micclass").value = "Image"
		oMeasureDownNormal("html id").value = ""
		oMeasureDownNormal("html tag").value = "IMG"
		oMeasureDownNormal("image type").value = "Plain Image"
		oMeasureDownNormal("outerhtml").regularexpression=True
		oMeasureDownNormal("outerhtml").value = ".*onMeasureContextClick.*"
		
		Set oMeasureDownNormalObj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrowthRateDropdowns").ChildObjects(oMeasureDownNormal)
				
		Browser("Analyzer").Page("ReportCreation").Sync
		'Mouse over on Pivot Table menu
		oMeasureDownNormalObj(ColVal).FireEvent "onmouseover"
		wait 2
		oMeasureDownNormalObj(ColVal).Click
		wait 2
		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image(strTabName).FireEvent "onmouseover"
		wait 1
		
		'First part of if condition will be executed, if sub-options has to be selected
		If strSelSubValue <> "" Then
			Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strSelValue).FireEvent "onmouseover"
			wait 1
			Set objSelSubValue = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strSelSubValue)
			Call SCA.ClickOn(objSelSubValue,strSelSubValue, "Report Creation Page")
		Else
			'Mouse Over on options options under each tab
			Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strSelValue).FireEvent "onmouseover"
		
		End If
		wait 2
		
	End Sub
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	Public Sub RGRowTotalRankMeasureDownNormal(ByVal ColVal, ByVal strTabName, ByVal strSelValue, ByVal strSelSubValue)
		
		Dim oMeasureDownNormal, oMeasureDownNormalObj, objSelSubValue, objSelValue
	
		Set oMeasureDownNormal = Description.Create()
		oMeasureDownNormal("micclass").value = "Image"
		oMeasureDownNormal("html id").value = ""
		oMeasureDownNormal("html tag").value = "IMG"
		oMeasureDownNormal("image type").value = "Plain Image"
		oMeasureDownNormal("outerhtml").regularexpression=True
		oMeasureDownNormal("outerhtml").value = ".*onMeasureContextClick.*"
		
		Set oMeasureDownNormalObj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabRowTotalRank").ChildObjects(oMeasureDownNormal)
				
		Browser("Analyzer").Page("ReportCreation").Sync
		'Mouse over on Pivot Table menu
		oMeasureDownNormalObj(ColVal).FireEvent "onmouseover"
		wait 2
		oMeasureDownNormalObj(ColVal).Click
		wait 2
		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image(strTabName).FireEvent "onmouseover"
		wait 1
		
		'First part of if condition will be executed, if sub-options has to be selected
		If strSelSubValue <> "" Then
			Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strSelValue).FireEvent "onmouseover"
			wait 1
			Set objSelSubValue = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strSelSubValue)
			Call SCA.ClickOn(objSelSubValue,strSelSubValue, "Report Creation Page")
		Else
			'Mouse Over on options options under each tab
			Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strSelValue).FireEvent "onmouseover"
		
		End If
		wait 2
		
	End Sub
	
	
	
	
	
	
	
	
	
	
	
	Public Sub RJRowTotalRankMeasureDownNormal(ByVal ColVal, ByVal strTabName, ByVal strSelValue, ByVal strSelSubValue)
		
		Dim oMeasureDownNormal, oMeasureDownNormalObj, objSelSubValue, objSelValue
	
		Set oMeasureDownNormal = Description.Create()
		oMeasureDownNormal("micclass").value = "Image"
		oMeasureDownNormal("html id").value = ""
		oMeasureDownNormal("html tag").value = "IMG"
		oMeasureDownNormal("image type").value = "Plain Image"
		oMeasureDownNormal("outerhtml").regularexpression=True
		oMeasureDownNormal("outerhtml").value = ".*onMeasureContextClick.*"
		
		Set oMeasureDownNormalObj = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrowthRateDropdowns").ChildObjects(oMeasureDownNormal)
				
		Browser("Analyzer").Page("ReportCreation").Sync
		'Mouse over on Pivot Table menu
		oMeasureDownNormalObj(ColVal).FireEvent "onmouseover"
		wait 2
		oMeasureDownNormalObj(ColVal).Click
		wait 2
		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image(strTabName).FireEvent "onmouseover"
		wait 1
		
		'First part of if condition will be executed, if sub-options has to be selected
		If strSelSubValue <> "" Then
			Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strSelValue).FireEvent "onmouseover"
			wait 1
			Set objSelSubValue = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strSelSubValue)
			Call SCA.ClickOn(objSelSubValue,strSelSubValue, "Report Creation Page")
		Else
			'Mouse Over on options options under each tab
			Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement(strSelValue).Click
		
		End If
		wait 2
		
	End Sub
	
	
	
	
	
	
	
	
	
	
	
	
	Public Function PeriodGrowthRate()
		
		Dim rc,cc,y,z,i,x,abc,abd,abcd,abe,varr,abede,abedef,m,n,o,p
		
		'To find the row and column count
		rc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").RowCount
		cc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").ColumnCount (1)
		
		
		cc = 4
		
		'First column period growth rate validation
		
		For m = 1 To 1 Step 1
			
			For n = 1 To rc Step 1
			
			o = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").GetCellData(n,m)
				
			p = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").GetCellData(n,m+1)
			
				If (p = "-") Then
					
					Call ReportStep(StatusTypes.Pass, "Period growth rate is null for first column, for row "&n&" and column "&m&" is"&p, "Pivot table")
				Else

					Call ReportStep(StatusTypes.Fail,"Period growth rate is not null for first column, for row "&n&" and column "&m&" is"&p, "Pivot table")
					
					
				End If	
			
			Next	
			
			
		Next
		
		
		
		
		
		
		
		
		y = 3 
		z = 4
		
		For i = 1 To cc-3 Step 2
			
			For x = 1 To rc Step 1
				
				abc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").GetCellData(x,i)
				
		'		y= 3
				
				abd = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").GetCellData(x,y)
				
				
				If abc = "0.0000" OR abc = "-" OR abd = "-" Then
					
				abcd = "-"	
					
'				ElseIf (z = cc  ) Then
'				Print "Period growth rate is null for grand total column"
'				abcd = "-"	
		
				Else
				
				abcd = Round((abd/abc*100-100),4)&" %"
					
				End If
				
		'		abcd = Round((abd/abc*100-100),4)&" %"
				
		'		abcde = Round(abcd,4)&" %"
			
		'		z=4
				
				abe = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").GetCellData(x,z)
				
				If abe <> "-" Then
					
					varr = split(abe," ")
				
					abede = Round(varr(0),4)
				
					abedef = replace(abede,",","")&" %"
				
				Else
				
					abedef = abe
					
				End If
				
				
				
				
				
				If (abedef = abcd) Then
		        
				Call ReportStep(StatusTypes.Pass, "Period growth rate  is display correctly for row "&x&" and column "&i&" is"&abedef, "Pivot table")
				
				Else
				
				Call ReportStep(StatusTypes.Fail, "Period growth rate  is not correct for row "&x&" and column "&i&" is"&abcd, "Pivot table")
				
				End If
				
				wait 2
			Next
			
			y = y+2
			
			z = z+2
			
			
		Next
		
		
	End Function
	
	
	
	
	
	Public Function PeriodGrowth()
		
		Dim rc,cc,y,z,i,x,abc,abd,abcd,pro,abe,abede,abedef,m,n,o,p
		
		
		'To find the row and column
		rc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").RowCount
		cc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").ColumnCount (1)
		
		cc = 4
		
		
		'First column period growth validation
		
		For m = 1 To 1 Step 1
			
			For n = 1 To rc Step 1
			
			o = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").GetCellData(n,m)
				
			p = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").GetCellData(n,m+1)
			
				If (p = o) Then
					
					Call ReportStep(StatusTypes.Pass, "Sales value and period growth is equal for first column, for row "&n&" and column "&m&" is"&p, "Pivot table")
				Else

					Call ReportStep(StatusTypes.Fail,"Sales value and period growth is not equal for first column, for row "&n&" and column "&m&" is"&p, "Pivot table")
					
					
				End If	
			
			Next	
			
			
		Next
		
		
		
		
		
		
		y = 3 
		z = 4
		
		For i = 1 To cc-3 Step 2
			
			For x = 1 To rc Step 1
				
				abc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").GetCellData(x,i)
				
		'		y= 3
				
				abd = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").GetCellData(x,y)
				
				
				If abc = "-" OR abd = "-" Then
					
				abcd = "-"	
					
'				ElseIf (z = cc ) Then
				
'				Call ReportStep (StatusTypes.Information, "For grand total sales value and period growth are same"&abd,"Pivot table")
				
'				abcd = Round(abd,4)	
		
				Else
				
				abcd = Round((abd-abc),4)
					
				End If
				
				pro = Cstr(abcd)
				
		'		abcd = Round((abd/abc*100-100),4)&" %"
				
		'		abcde = Round(abcd,4)&" %"
			
		'		z=4
				
				abe = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabSalesValueGrowth").GetCellData(x,z)
				
				If abe <> "-" Then
					
		'			varr = split(abe," ")
				
					abede = Round(abe,4)
				
					abedef = replace(abede,",","")
				
				Else
				
					abedef = abe
					
				End If
				
				
				
				
				
				If (abedef = pro) Then
				
				Call ReportStep(StatusTypes.Pass, "Period growth is display correctly for row "&x&" and column "&i&" is"&abedef, "Pivot table")
				
		
				Else
				
				Call ReportStep(StatusTypes.Fail, "Period growth is not correct for row "&x&" and column "&i&" is"&abedef, "Pivot table")
			
				
				End If
				
				wait 2
			Next
			
			y = y+2
			
			z = z+2
			
			
		Next
		
		
	End Function

	Public Function Hierarchy(ByVal strFilterCondition, ByRef objData)
		
		
	
'	Select Case strFilterCondition
		
'		Case "BeginWith"
		
		
		
		Dim a()
		rc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrpTableColAttributeData").GetROProperty("rows")
		ReDim preserve a(rc-1,0) 
			 For i = 1 To rc Step 1
			 
				 abc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrpTableColAttributeData").GetCellData(i,1)
		'	 	 print abc
				 print(abc)
			 	 
			 	 If (Instr(abc,"A") = 2) Then
			 	 	
		'	 	 	print "rajesh"
			 	 	
			 	 	a(i-1,0) = abc
'			 	 print a(i-1,0)
			 	 	print(a(i-1,0))
			 	 End If
			 	 
			 Next
		
		
		
		'Set the filter
		
		'Click on set filter functionality
		Call IMSSCA.General.RowDimensionDownNormal(0, objData.item("MeasureDownNormalTabName"),objData.item("SelectValue"), objData.item("SelectSubValue"))
		
		
		Set objAllMembers = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").Image("imgAllMembers")
		Call SCA.ClickOn(objAllMembers,"All Members", "Filter Condition Settings")
		
		
		Set objFilterMembers = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").Image("imgFilterMembers")
		Call SCA.ClickOn(objFilterMembers,"Filter Members", "Filter Condition Settings")
		
		Set objOperationOnMembers = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebList("wliOperationOnMembers")
		Call SCA.SelectFromDropdown(objOperationOnMembers,objData.item("strOperationOnMembers"))
		
		
		Set objTxt = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebEdit("txtfilterMembers")
		Call SCA.SetText(objTxt,objData.Item("FilterString"),"Filter String","Filter Members")
		
		Set objFilter = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebButton("btnFilter")
		Call SCA.ClickOn(objFilter,"Filter Members", "Filter Condition Settings")
		
		
		
			'validate the member against webtable after giving filter condition

		Set d = description.Create
		d("micclass").value = "WebElement"
		d("html tag").value = "SPAN"
		d("innertext").RegularExpression = TRUE
		d("innertext").value = "A.*-.*"
		
		
		Set obj = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebElement("tabAllMembers").ChildObjects(d)
		
		obj.count
		Dim b()
		If (obj.count>0) Then
		ReDim Preserve b(obj.count-1)	
			For j = 1 To obj.count-1 Step 1
				b(j-1) = obj(j).getroproperty("innertext")
				msgbox trim(a(j,0))
'				msgbox obj(j).getroproperty("innertext")
				
				If (trim(a(j-1,0)) = obj(j).getroproperty("innertext")) Then
				
					Call ReportStep (StatusTypes.Pass, "Values in pivot table(Begins with 'A') and after giving filter condition(Begins with 'A') are same"&trim(a(j-1,0)),"Filter condition settings")
'					print "members are filtered with given criteria"
'					Print a(j-1,0)
				


				Else 
				
					Call ReportStep (StatusTypes.Fail,"Values in pivot table(Begins with 'A') and after giving filter condition(Begins with 'A') are not same"&trim(a(j-1,0)),"Filter condition settings")
				



				End If
				
				
				
			Next
				
			
		End If
		
	
	
	
	
	'Click on all members check box
		wait 2
		Call SCA.ClickOn(objAllMembers,"All Members", "Filter Condition Settings")


		'clicking on random check box


'		Set desc = description.Create
'		desc("micclass").value = "Image"
'		desc("image type").value = "Plain Image"
'		desc("html tag").value = "IMG"
'		desc("name").value = "Image"
'		desc("outerhtml").RegularExpression = TRUE
'		desc("outerhtml").value = "<IMG style=.*"
		
		Set desc = description.Create
		desc("micclass").value = "Image"
		desc("image type").value = "Plain Image"
		desc("html tag").value = "IMG"
		desc("file name").value = "u\.gif"
		
		Set obj1 = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebElement("tabAllMembers").ChildObjects(desc)
		
		obj1.count
		
		For K = 1 To obj1.count-1 Step 1
			
			If k <> 6 Then
				
				obj1(K).click
				
			Else
				Exit For
			End If
			
			
			
		Next
		
		
		
			Set objOK =Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebButton("btnOK")
		Call SCA.ClickOn(objOK,"OK", "Filter Condition Settings")





		'validate the values after filter
		wait 2
		rc1 = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrpTableColAttributeData").GetROProperty("rows")
		'ReDim preserve a(rc1-1,0) 
			 For m = 1 To rc1-1 Step 1
			 
				 abcd = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrpTableColAttributeData").GetCellData(m,1)
				 
				 If (trim(abcd) = b(m-1)) Then
				 	
'				 	print "value is verified"&abcd
				 	
				 	
				 	Call ReportStep (StatusTypes.Pass, "Values selected in filter condition settings and pivot table are same"&trim(abcd),"Pivot table")
				


				Else 
				
					Call ReportStep (StatusTypes.Fail,"Values selected in filter condition settings and pivot table are not same"&trim(abcd),"Pivot table")
				
				 		
				 	
				 	
				 End If
				
			 
			 Next
		
			wait 2

		
		
		Set objFilteredApplied = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgFilteredApplied")
		Call SCA.ClickOn(objFilteredApplied,"Filter", "Pivot table")
		
		Set objClearFilter =Browser("Analyzer").Page("ReportCreation").Frame("__dialog").Image("imgClearFilter")
		Call SCA.ClickOn(objClearFilter,"Filter", "Filter condition settings")

'		Set objOK =Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebButton("btnOK")
		Call SCA.ClickOn(objOK,"OK", "Filter Condition Settings")
		
		
	
		
	End Function
	
	
	
	Public Function Hierarchy1(ByVal strFilterCondition, ByRef objData)
		
		
	
	'2. Does not begins with a

		Dim a()
		
		n =0
		rc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrpTableColAttributeData").GetROProperty("rows")
		ReDim preserve a(rc-1,0) 
		
			For i = 1 To rc-1 Step 1
			 
				 abc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrpTableColAttributeData").GetCellData(i,1)
		'	 	 print abc
			 	 
			 	 If (Instr(abc,"A") <> 2) Then
			 	 	
		'	 	 	print "rajesh"
			 	 	
			 	 	a(n,0) = abc
'			 	 	print a(n,0)
			 	 	n = n+1
			 	 End If
			 	 
			 Next
			 
			 wait 2




		'Set the filter	 
		

		'Click on set filter functionality
		Call IMSSCA.General.RowDimensionDownNormal(0, objData.item("MeasureDownNormalTabName"),objData.item("SelectValue"), objData.item("SelectSubValue"))
		
		
		Set objAllMembers = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").Image("imgAllMembers")
		Call SCA.ClickOn(objAllMembers,"All Members", "Filter Condition Settings")
		
		
		Set objFilterMembers = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").Image("imgFilterMembers")
		Call SCA.ClickOn(objFilterMembers,"Filter Members", "Filter Condition Settings")
		
		Set objOperationOnMembers = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebList("wliOperationOnMembers")
		Call SCA.SelectFromDropdown(objOperationOnMembers,objData.item("strOperationOnMembers1"))
		
		
		Set objTxt = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebEdit("txtfilterMembers")
		Call SCA.SetText(objTxt,objData.Item("FilterString"),"Filter String","Filter Members")
		
		Set objFilter = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebButton("btnFilter")
		Call SCA.ClickOn(objFilter,"Filter Members", "Filter Condition Settings")
		


		'validate the member against webtable after giving filter condition
		
		Set d = description.Create
		d("micclass").value = "WebElement"
		d("html tag").value = "SPAN"
		d("innertext").RegularExpression = TRUE
		d("innertext").value = ".*-.*"
		
		
		Set obj = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebTable("html id:=igtabtabMain","html tag:=TABLE","name:=WebTable").ChildObjects(d)
		

		Dim b()
		If (obj.count>0) Then
		ReDim Preserve b(obj.count-1)	
			For j = 1 To obj.count-3 Step 1
				b(j-1) = obj(j).getroproperty("innertext")
				If (trim(a(j-1,0)) = obj(j).getroproperty("innertext")) Then
					
					Call ReportStep (StatusTypes.Pass, "Values in pivot table(Does not Begin with 'A') and after giving filter condition(Does not begin with 'A') are same"&trim(a(j-1,0)),"Filter condition settings")

				Else 
				
					Call ReportStep (StatusTypes.Fail,"Values in pivot table(Does not Begin with 'A') and after giving filter condition(Does not begin with 'A') are not same"&trim(a(j-1,0)),"Filter condition settings")
					
				End If
				
			Next
				
		End If
		
		
		'Click on all members check box
		Call SCA.ClickOn(objAllMembers,"All Members", "Filter Condition Settings")




		'clicking on random check box
		
'		Set desc = description.Create
'		desc("micclass").value = "Image"
'		desc("image type").value = "Plain Image"
'		desc("html tag").value = "IMG"
'		desc("name").value = "Image"
'		desc("outerhtml").RegularExpression = TRUE
'		desc("outerhtml").value = "<IMG style=.*"

		Set desc = description.Create
		desc("micclass").value = "Image"
		desc("image type").value = "Plain Image"
		desc("html tag").value = "IMG"
		desc("file name").value = "u\.gif"
		
		'Set obj1 = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebElement("tabAllMembers").ChildObjects(desc)
		Set obj1 =  Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebTable("html id:=igtabtabMain","html tag:=TABLE","name:=WebTable").ChildObjects(desc)
		obj1.count
		
		For K = 1 To obj1.count-1 Step 1
			
			If k <> 6 Then
				
				obj1(K).click
				
			Else
				Exit For
			End If
			
			
		Next
		
		Set objOK =Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebButton("btnOK")
		Call SCA.ClickOn(objOK,"OK", "Filter Condition Settings")
		

		'validate the values after filter
		
		wait 5
		
		Set obj1=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("class:=pivot_table","name:=scube").ChildItem(2,1,"WebTable",0)
		rc1 = obj1.GetROProperty("rows")
		'ReDim preserve a(rc1-1,0) 
			 For m = 1 To rc1-1 Step 1
			 
				 abcd = obj1.GetCellData(m,1)
				 
				 If (trim(abcd) = b(m-1)) Then
				 	
'				 	print "value is verified"&abcd
				 	
				 	Call ReportStep (StatusTypes.Pass, "Values selected in filter condition settings and pivot table are same"&trim(abcd),"Pivot table")
				
				Else 
				
					Call ReportStep (StatusTypes.Fail,"Values selected in filter condition settings and pivot table are not same"&trim(abcd),"Pivot table")
				
				 	
				 End If
				
			 
			 Next
		
		wait 5
		
		
		Set objFilteredApplied = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgFilteredApplied")
		Call SCA.ClickOn(objFilteredApplied,"Filter", "Pivot table")
		
		Set objClearFilter =Browser("Analyzer").Page("ReportCreation").Frame("__dialog").Image("imgClearFilter")
		Call SCA.ClickOn(objClearFilter,"Filter", "Filter condition settings")

'		Set objOK =Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebButton("btnOK")
		Call SCA.ClickOn(objOK,"OK", "Filter Condition Settings")
		Wait 5
		
		
		
		
		
		
	End Function


	Public Function Hierarchy2(ByVal strFilterCondition, ByRef objData)
		
		
		
		
		
		'3. Contains the word "ANTI"
		
		Dim a()
		n =0
		rc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrpTableColAttributeData").GetROProperty("rows")
		ReDim preserve a(rc-1,0) 
			 For i = 1 To rc-1 Step 1
			 
				 abc = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrpTableColAttributeData").GetCellData(i,1)
		'	 	 print abc
			 	 
			 	 If (Instr(abc,"ANTI") > 0) Then
			 	 	
		'	 	 	print "rajesh"
			 	 	
			 	 	a(n,0) = abc
'			 	 	print a(n,0)
			 	 	n = n+1
			 	 End If
			 	 
			 Next
			 
			 wait 2



		'Set the filter	 
		
		'Click on set filter functionality
		Call IMSSCA.General.RowDimensionDownNormal(0, objData.item("MeasureDownNormalTabName"),objData.item("SelectValue"), objData.item("SelectSubValue"))
		
		
		Set objAllMembers = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").Image("imgAllMembers")
		Call SCA.ClickOn(objAllMembers,"All Members", "Filter Condition Settings")
		
		
		Set objFilterMembers = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").Image("imgFilterMembers")
		Call SCA.ClickOn(objFilterMembers,"Filter Members", "Filter Condition Settings")
		
		Set objOperationOnMembers = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebList("wliOperationOnMembers")
		Call SCA.SelectFromDropdown(objOperationOnMembers,objData.item("strOperationOnMembers2"))
		
		
		Set objTxt = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebEdit("txtfilterMembers")
		Call SCA.SetText(objTxt,objData.Item("FilterString1"),"Filter String","Filter Members")
		
		Set objFilter = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebButton("btnFilter")
		Call SCA.ClickOn(objFilter,"Filter Members", "Filter Condition Settings")
		



		



		'validate the member against webtable after giving filter condition
		
		Set d = description.Create
		d("micclass").value = "WebElement"
		d("html tag").value = "SPAN"
		d("innertext").RegularExpression = TRUE
		d("innertext").value = ".*-.*"
		
		
		Set obj = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebElement("tabAllMembers").ChildObjects(d)
		
		obj.count
		Dim b()
		If (obj.count>0) Then
		ReDim Preserve b(obj.count-1)	
			For j = 1 To obj.count-2 Step 1
				b(j-1) = obj(j).getroproperty("innertext")
				If (trim(a(j-1,0)) = obj(j).getroproperty("innertext")) Then
				
					
					Call ReportStep (StatusTypes.Pass, "Values in pivot table(Contains 'ANTI') and after giving filter condition(Contains 'ANTI') are same"&trim(a(j-1,0)),"Filter condition settings")

				Else 
				
					Call ReportStep (StatusTypes.Fail,"Values in pivot table(Contains 'ANTI') and after giving filter condition(Contains 'ANTI') are not same"&trim(a(j-1,0)),"Filter condition settings")
					
					
					
				
				
'					print "members are filtered with given criteria"
'					Print a(j-1,0)
					
				End If
				
				
				
			Next
				
			
		End If
		
		'Click on all members check box
		Call SCA.ClickOn(objAllMembers,"All Members", "Filter Condition Settings")

		
		
'		Browser("Analyzer").Page("ReportCreation").Frame("__dialog").Image("imgAllMembers").Click



		'clicking on random check box
		
		
'		Set desc = description.Create
'		desc("micclass").value = "Image"
'		desc("image type").value = "Plain Image"
'		desc("html tag").value = "IMG"
'		desc("name").value = "Image"
'		desc("outerhtml").RegularExpression = TRUE
'		desc("outerhtml").value = "<IMG style=.*"

		Set desc = description.Create
		desc("micclass").value = "Image"
		desc("image type").value = "Plain Image"
		desc("html tag").value = "IMG"
		desc("file name").value = "u\.gif"
		
		Set obj1 = Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebElement("tabAllMembers").ChildObjects(desc)
		
		obj1.count
		
		For K = 1 To obj1.count-1 Step 1
			
			If k <> 6 Then
				
				obj1(K).click
				
			Else
				Exit For
			End If
		
		Next
		
		Set objOK =Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebButton("btnOK")
		Call SCA.ClickOn(objOK,"OK", "Filter Condition Settings")
		
		
		
		
'		Browser("Analyzer").Page("GroupCreationPage").Frame("__dialogAddAllMembers").WebButton("btnOK").Click




		'validate the values after filter
		wait 2
		rc1 = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrpTableColAttributeData").GetROProperty("rows")
		'ReDim preserve a(rc1-1,0) 
			 For m = 1 To rc1-1 Step 1
			 
				 abcd = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabGrpTableColAttributeData").GetCellData(m,1)
				 
				 If (trim(abcd) = b(m-1)) Then
				 	
				 	
				 	Call ReportStep (StatusTypes.Pass, "Values selected in filter condition settings and pivot table are same"&trim(abcd),"Pivot table")
				
				Else 
				
					Call ReportStep (StatusTypes.Fail,"Values selected in filter condition settings and pivot table are not same"&trim(abcd),"Pivot table")
	
'				 	print "value is verified"&abcd
				 	
				 End If
				
			 
			 Next
		
		wait 2

		
		Set objFilteredApplied = Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").Image("imgFilteredApplied")
		Call SCA.ClickOn(objFilteredApplied,"Filter", "Pivot table")
		
		Set objClearFilter =Browser("Analyzer").Page("ReportCreation").Frame("__dialog").Image("imgClearFilter")
		Call SCA.ClickOn(objClearFilter,"Filter", "Filter condition settings")

'		Set objOK =Browser("Analyzer").Page("ReportCreation").Frame("__dialog").WebButton("btnOK")
		Call SCA.ClickOn(objOK,"OK", "Filter Condition Settings")
		

		
		
	End Function
	
	






'' changed code,.....
''<author :-Shobha>  
'    Public Function ReportCreationInSCAShobhaCode(ByVal sheetNum, ByVal newReport,ByVal intEMD,ByVal intdbcheck,ByVal intRNum,ByVal intReportCreate,ByVal strF1,ByVal strF2)  
'    
'        Dim objExcel,ObjExcelFile,ObjExcelSheet,objDataSource,objDatabase,objDatacubes,obj_tree,objDR,objrow,objcolumn,objDataaxis,objAdd,objtreeCount,objCreate
'        Dim intRow_Count,intColumn_Count,intcount
'        Dim dataSourceValue, strDataSourceValue,strDatabaseValue,strCubeValue,strChild_Values,strPlaceHolderVal,p, objBtnok, objErrorOk,strDBitems,strbkch
'        '<Looping Variables for the row and column>
'        Dim i,j,k,x,y,a  
'        
'        '<shweta 11/9/2015 Closing All opened Excel Files- start
'		'objUtils.KillProcess("excel.exe")
'		Systemutil.CloseProcessByName "excel.exe"
'		'<shweta 11/9/2015 Closing All opened Excel Files- end
'        
'        Set objExcel = CreateObject("Excel.Application") 
'        objExcel.Visible =True     
'    
'       ' Set ObjExcelFile = objExcel.Workbooks.Open(Environment.Value("CurrDir")&"InputFiles\IMSSCAWeb\SCAReportSheet.xls")
'        Set ObjExcelFile = objExcel.Workbooks.Open(Environment.Value("CurrDir")&Environment.Value("ReportCreationFile"))
'        Set ObjExcelSheet = ObjExcelFile.Sheets(sheetNum) 
'        intRow_Count = ObjExcelSheet.UsedRange.Rows.Count 
'        intColumn_Count = ObjExcelSheet.UsedRange.Columns.Count
'    
'        dataSourceValue=TRIM(ObjExcelSheet.cells(2,1).value)
'        strDataSourceValue = "+ "&dataSourceValue       
'        strDatabaseValue=TRIM(ObjExcelSheet.cells(2,2).value)
'        strCubeValue=TRIM(ObjExcelSheet.cells(2,3).value)
'        
'        If intEMD=1 Then
'        
'         If newReport = 0 Then
'            Set objCreate=Browser("Analyzer").Page("Shared_Folder").Frame("view").Image("imgnewreport")        
'           'Set objCreate=Browser("Analyzer").Page("Shared_Folder").Image("imgnewreport")-------Ask Shobha
'           Call SCA.ClickOn(objCreate,"newreport","Report Creation")    
'        End If
'              
'       'Capture Database Access Error - Start
'       If Browser("Analyzer").Frame("ErrorFrame").WebElement("Database access error.").Exist(1) Then
'               Set objErrorOk=Browser("Analyzer").Frame("ErrorFrame").WebButton("btnOK")
'               Call SCA.ClickOn(objErrorOk,"btnOK" , "Database Access Error during Report Creation")
'               Call ReportStep (StatusTypes.Fail,"Not Able to create a report. Database Access Error during Report Creation", "Report Creation Page")    
'                    
'            'Closing Excel
'            ObjExcelFile.Close
'            objExcel.Quit    
'            Set objExcel=nothing
'            Set  ObjExcelFile=nothing
'            Set  ObjExcelSheet=nothing
'            Set  objtreeCount=nothing
'            ReportCreationInSCA = 1
'            Exit Function
'       End If
'       'Capture Database Access Error - End
'       
'       'Handled Application Time Out in Report Creation Page - Start
'       Set objBtnok = Browser("Analyzer").Dialog("Message from webpage").WinButton("btnOK")
'       If objBtnok.Exist(5) Then
'            Call SCA.ClickOn(objBtnok,"btnOK" , "Application Time Out in Report Creation Page")
'       End If    
'       'Handled Application Time Out in Report Creation Page - End
'       
'       For a = 1 To 100 Step 1           
'       If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDataSources").GetROProperty("disable")=0  Then
'         Exit For
'       End If           
'       Next
'       
'       wait 2
'       
'       Set objDataSource=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDataSources")   
'       Call SCA.SelectFromDropdown(objDataSource,strDataSourceValue)
'       
'       wait 2
'       
'       Set objDatabase=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDatabases")
'       Call SCA.SelectFromDropdown(objDatabase,strDatabaseValue)       
'       
'       If intdbcheck=0 Then   
'       
'        strDBitems=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDatabases").GetROProperty("all items")
'         If Instr(strDBitems,strDatabaseValue)<>0 Then           
'            ReportCreationInSCA=0
'               else
'            ReportCreationInSCA=1
'         End If 
'         
'         If intReportCreate=1 Then
'            
'            
'            ObjExcelFile.Close
'            objExcel.Quit    
'            Set objExcel=nothing
'            Set ObjExcelFile=nothing
'            Set ObjExcelSheet=nothing
'            Set objtreeCount=nothing
'            Exit Function
'            
'        End If
'     
'       End If    
'       
'       
'       Set objDatacubes=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlCubes")
'       Call SCA.SelectFromDropdown(objDatacubes,strCubeValue)
'       
'       Wait 2
'   	   Set objAdd=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebButton("btnAdd")
'   	   Call SCA.ClickOn(objAdd,"AddButton" , "ReportCreation")
'       
'       
'       
'    
'       For i=2 to intRow_Count
'        Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcube").Click           
'         For j=4 to intColumn_Count-1
'              strChild_Values=Trim(ObjExcelSheet.cells(i,j).value)
'               ' Have to write the explicit mentain the bk mark condition
'               strbkch=Trim(ObjExcelSheet.cells(i,intColumn_Count).value)
'               If  strChild_Values<>"" AND  strChild_Values<>"rowaxis" AND strChild_Values<>"columnaxis"  AND strChild_Values<>"dataaxis" AND strbkch<>"Yes" AND strChild_Values<>"No" Then
'                        Set obj_tree=description.Create
'                        obj_tree("micclass").value="WebElement"
'                        obj_tree("class").value="TreeNode"
'                        obj_tree("innertext").value=strChild_Values
'                        wait 1
'                        set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
'                        intcount=objtreeCount.count
'                        
'                        For a = 1 To 100 Step 1
'	                    set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
'	                    intcount=objtreeCount.count
'                            If intcount<>0 Then
'                            Exit For
'                            End If
'                        Next
'                        
'                        If intcount<>0 Then
'                        '' ch - 25 - 001 - May - Rajesh - Start 
''                        objtreeCount(0).fireEvent  "ondblclick" 
'                        	'shweta - 31/3/2016 - start
'                        	If intcount = 2 AND strF1 <> "" Then
'                        		k=intcount-2                    
'                                objtreeCount(k).fireEvent  "ondblclick"  	
'                        	
'                        	ElseIf intcount = 2 AND strChild_Values <> "Additional attributes" Then
'                        		k=intcount-2                    
'                                objtreeCount(k).fireEvent  "ondblclick"  	
'                        	Else
'								k=intcount-1                    
'                                objtreeCount(k).fireEvent  "ondblclick"  
'                        	End If
'                        	'' ch - 25 - 001 - May - Rajesh - End..
' 
'							'shweta - 31/3/2016 - End                            
'                            Call ReportStep (StatusTypes.Pass,"Clicked on the Webelement"&Space(2)&strChild_Values, "Report Creation Page")                    
'                            wait 3    
'                        Else
'                            Call ReportStep (StatusTypes.Warning,"Not Clicked on the Webelement"&Space(2)&strChild_Values, "Report Creation Page")
'	                    	ObjExcelFile.Close
'	  						objExcel.Quit    
'	   						Set objExcel=nothing
'	   						Set ObjExcelFile=nothing
'	   						Set ObjExcelSheet=nothing
'	   						Set objtreeCount=nothing
'                            Exit Function
'                        End If 
'
'                                     
'               End If
'               
'              If j=intColumn_Count-1 AND strbkch<>"Yes" Then
'                       
'                    wait 5
'                    
''                    Set objDR=CreateObject("Mercury.DeviceReplay")
''                    x=objtreeCount(k).GetroProperty("abs_x")
''                    y=objtreeCount(k).GetroProperty("abs_y")  
''                    wait 1                    
''                    objDR.MouseClick x,y,2
'
''obj_tree("micclass").value="WebElement"
''                            obj_tree("class").value="TreeNode"
''                            obj_tree("innertext").value=strChild_Values
''
''    Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("class:=TreeNode","index:=1","innertext:="&strChild_Values).RightClick
'                    objtreeCount(k).RightClick
'                    
'                    
'                    strPlaceHolderVal=lcase(ObjExcelSheet.cells(i,intColumn_Count-1).value)                                                                                        
'                    wait 1  
'                      If strPlaceHolderVal="rowaxis" Then 
'                      	  	
'                          Set objrow=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webRowaxis")    
'                          Call SCA.ClickOn(objrow,"Rowaxis" , "ReportCreation")    
'                          wait 5                          
'                      elseif  StrPlaceHolderVal="columnaxis" then  
'						                       
'                          Set objcolumn=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcolumnaxis")
'                          Call SCA.ClickOn(objcolumn,"Columnaxis" , "ReportCreation")
'                          Wait 5                      
'                      elseif  StrPlaceHolderVal="dataaxis" then
'						                    
'                         Set objDataaxis=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webDataAxis")
'                         Call SCA.ClickOn(objDataaxis,"Dataaxis" , "ReportCreation")    
'                         wait 5
'                      End If    
'                      
'                      For p=intColumn_Count-2 to 4  step -1
'                         strChild_Values=Trim(ObjExcelSheet.cells(i,p).value)
'                         strbkch=Trim(ObjExcelSheet.cells(i,intColumn_Count).value)                         
'                         If  strChild_Values<>"" AND  strChild_Values<>"rowaxis" AND strChild_Values<>"columnaxis"  AND strChild_Values<>"dataaxis" AND strbkch<>"Yes" Then
'                            Set obj_tree=description.Create
'                            obj_tree("micclass").value="WebElement"
'                            obj_tree("class").value="TreeNode"
'                            obj_tree("innertext").value=strChild_Values
'                            wait 1
'                            set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
'                            intcount=objtreeCount.count    
'                                If intcount<>0 Then
'                                
'                                   If intcount = 2 AND strF1 <> "" Then
'		                        		k=intcount-2                    
'		                                objtreeCount(k).fireEvent  "ondblclick"  	
'	
'                                   ElseIf intcount = 2 AND strChild_Values <> "Additional attributes" Then
'                                   	  k=intcount-2
'                                      objtreeCount(k).fireEvent  "ondblclick" 
'								   Else
'									   k=intcount-1
'                                       objtreeCount(k).fireEvent  "ondblclick"    
'                                   End If
'                                   
''                                   k=intcount-1
''                                   objtreeCount(k).fireEvent  "ondblclick"    
'                                 End If                          
'                            End If
'                        Next
'                
'                End If
'             
'           Next
'    
'           Next
'            
'   
'       
'   ElseIf intEMD=0 Then
'           
''           Set objExcel = CreateObject("Excel.Application") 
''        objExcel.Visible =True     
''        
''        Set ObjExcelFile = objExcel.Workbooks.Open(Environment.Value("CurrDir")&"InputFiles\IMSSCAWeb\SCAReportSheet.xls")
''        Set ObjExcelSheet = ObjExcelFile.Sheets(intsheetName)
''        
''        intRow_Count = ObjExcelSheet.UsedRange.Rows.Count 
''        intColumn_Count = ObjExcelSheet.UsedRange.Columns.Count   
'        Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcube").Click          
'        For j=4 to intColumn_Count-2
'        strChild_Values=Trim(ObjExcelSheet.cells(intRNum,j).value)
'               
'          If  strChild_Values<>"" AND  strChild_Values<>"rowaxis" AND strChild_Values<>"columnaxis"  AND strChild_Values<>"dataaxis" AND strChild_Values<>"Yes" AND strChild_Values<>"No" Then
'        
'               Set obj_tree=description.Create
'               obj_tree("micclass").value="WebElement"
'               obj_tree("class").value="TreeNode"
'               obj_tree("innertext").value=strChild_Values
'               wait 1
'               set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
'               intcount=objtreeCount.count
'                        
'                For a = 1 To 20 Step 1
'                   If intcount<>0 Then
'                    Exit For
'                   End If
'                Next
'                        
'               If intcount<>0 Then
'                k=intcount-1                    
'                objtreeCount(k).fireEvent  "ondblclick"        
'                Call ReportStep (StatusTypes.Pass,"Clicked on the Webelement"&Space(2)&strChild_Values, "Report Creation Page")                    
'                wait 3    
'                Else
'                   Call ReportStep (StatusTypes.Warning,"Not Clicked on the Webelement"&Space(2)&strChild_Values, "Report Creation Page")
'                Exit Function
'               End If 
'            End If
'
'        If j=intColumn_Count-2  Then
'                     
'              wait 1
'              
''              Set objDR=CreateObject("Mercury.DeviceReplay")
''               x=objtreeCount(k).GetroProperty("abs_x")
''               y=objtreeCount(k).GetroProperty("abs_y") 
''			   objtreeCount(k).fireEvent  "ondblclick"	               
''                wait 1                    
''                objDR.MouseClick x,y,2
'              
'              	objtreeCount(k).RightClick
'              
'               strPlaceHolderVal=lcase(ObjExcelSheet.cells(intRNum,intColumn_Count-1).value)                                                                                        
'                      
'                      If strPlaceHolderVal="rowaxis" Then 
'                          Set objrow=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webRowaxis")    
'                          Call SCA.ClickOn(objrow,"Rowaxis" , "ReportCreation")    
'                          wait 5                          
'                      elseif  StrPlaceHolderVal="columnaxis" then                                                                                                                                
'                          Set objcolumn=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcolumnaxis")
'                          Call SCA.ClickOn(objcolumn,"Columnaxis" , "ReportCreation")
'                          Wait 5                      
'                      elseif  StrPlaceHolderVal="dataaxis" then                                                                            
'                         Set objDataaxis=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webDataAxis")
'                         Call SCA.ClickOn(objDataaxis,"Dataaxis" , "ReportCreation")    
'                         wait 5
'                      End If    
'                      
'                      For p=intColumn_Count-2 to 4  step -1
'                         strChild_Values=Trim(ObjExcelSheet.cells(intRNum,p).value)
'                         strbkch=Trim(ObjExcelSheet.cells(intRNum,intColumn_Count).value)                         
'                         If  strChild_Values<>"" AND  strChild_Values<>"rowaxis" AND strChild_Values<>"columnaxis"  AND strChild_Values<>"dataaxis"  Then
'                            Set obj_tree=description.Create
'                            obj_tree("micclass").value="WebElement"
'                            obj_tree("class").value="TreeNode"
'                            obj_tree("innertext").value=strChild_Values
'                            wait 1
'                            set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
'                            intcount=objtreeCount.count    
'                                If intcount<>0 Then
'                                   k=intcount-1
'                                   objtreeCount(k).fireEvent  "ondblclick"    
'                                 End If                          
'                            End If
'                        Next
'                
'                End If
'        Next
'        
'    End If
'    
'    
'     
'       ObjExcelFile.Close
'       objExcel.Quit    
'       Set objExcel=nothing
'       Set ObjExcelFile=nothing
'       Set ObjExcelSheet=nothing
'       Set objtreeCount=nothing
'    
'    End Function
'


Public Function ReportOLD (ByVal axis,Byref objData)
	
	
	If axis = "row" Then
	
		dataSourceValue=objData.item("datasource")
	    strDataSourceValue = "+ "&dataSourceValue       
	    strDatabaseValue=objData.item("database")
	    strCubeValue=objData.item("Cube")
	     
	    Set objCreate=Browser("Analyzer").Page("Shared_Folder").Frame("view").Image("imgnewreport")        
	    Call SCA.ClickOn(objCreate,"newreport","Report Creation")    
	         
	   'Capture Database Access Error - Start
	   If Browser("Analyzer").Frame("ErrorFrame").WebElement("Database access error.").Exist(1) Then
	       Set objErrorOk=Browser("Analyzer").Frame("ErrorFrame").WebButton("btnOK")
	       Call SCA.ClickOn(objErrorOk,"btnOK" , "Database Access Error during Report Creation")
	       Call ReportStep (StatusTypes.Fail,"Not Able to create a report. Database Access Error during Report Creation", "Report Creation Page")    
	       Report = 1
	       Exit Function
	   End If
	   'Capture Database Access Error - End
	   
	   'Handled Application Time Out in Report Creation Page - Start
	   Set objBtnok = Browser("Analyzer").Dialog("Message from webpage").WinButton("btnOK")
	   If objBtnok.Exist(5) Then
	        Call SCA.ClickOn(objBtnok,"btnOK" , "Application Time Out in Report Creation Page")
	   End If    
	   'Handled Application Time Out in Report Creation Page - End
	   wait 2
	   
	   Set objDataSource=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDataSources")   
	   Call SCA.SelectFromDropdown(objDataSource,strDataSourceValue)
	   
	   wait 2
	   
	   Set objDatabase=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDatabases")
	   Call SCA.SelectFromDropdown(objDatabase,strDatabaseValue)       
	   
	   Wait 2
	   Set objDatacubes=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlCubes")
	   Call SCA.SelectFromDropdown(objDatacubes,strCubeValue)
	   
	   Wait 5
	   
	   Set objAdd=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebButton("btnAdd")
	   Call SCA.ClickOn(objAdd,"AddButton" , "ReportCreation")	
		
	End If
	
	
	If axis <> "row" Then
		
		strCubeValue=objData.item("Cube")
		Set objAdd=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebButton("btnAdd")
	End If
	
	

	
	
    
	If axis = "row" Then
		Hierarchys = objData.item("RowHierarchy")
	ElseIf axis = "column" Then	
		Hierarchys = objData.item("ColumnHierarchy")
	ElseIf axis = "data" Then
		Hierarchys = objData.item("DataHierarchy")
	End If
'	Hierarchy = objData.item("Hierarchy")
'	Cube = "SMP_IMS_M_TRN_0001"
	HierarcyValues = split(Hierarchys,";")
	
	
	
	
	
	For i = 0 To ubound(HierarcyValues) Step 1
		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDataSource").WebElement("innertext:="&HierarcyValues(i),"class:=TreeNode").FireEvent "ondblclick"	
	Next
	
	
	
	If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDataSource").WebElement("innertext:="&HierarcyValues(i-1),"class:=TreeNode","index:=1").exist(2) Then
		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDataSource").WebElement("innertext:="&HierarcyValues(i-1),"class:=TreeNode","index:=1").RightClick	
	ElseIf Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDataSource").WebElement("innertext:="&HierarcyValues(i-1),"class:=trvSchemaSel SelectedNode","index:=1").Exist(2) Then
		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDataSource").WebElement("innertext:="&HierarcyValues(i-1),"class:=trvSchemaSel SelectedNode","index:=1").RightClick
	End If
	
	
	'select row, column and data axis
	If axis = "row" Then
		Set objrow=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webRowaxis")    
        Call SCA.ClickOn(objrow,"Rowaxis" , "ReportCreation")   
	ElseIf axis = "column" Then
		Set objcolumn=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcolumnaxis")
        Call SCA.ClickOn(objcolumn,"Columnaxis" , "ReportCreation")
	ElseIf axis = "data" Then
		Set objDataaxis=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webDataAxis")
        Call SCA.ClickOn(objDataaxis,"Dataaxis" , "ReportCreation")    
	End If
	
'	Browser("Analyzer").Page("Analyzer").Frame("ReportWindowOfAnalyzer").WebElement("apc").Click
	wait 2
	
	If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("class:=null","html tag:=SPAN").Exist(2) Then
		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("class:=null","html tag:=SPAN").RightClick
	Else	
		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("class:=TreeNode","html tag:=SPAN","innertext:="&strCubeValue).RightClick
	End If
	
'	Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("removecube").Click
	Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welRemoveCube").Click

	
	
'	Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebButton("Add").Click
	Call SCA.ClickOn(objAdd,"AddButton" , "ReportCreation")	
	
	
End Function





Public Function Report (ByVal axis,ByVal HierarchyValue,ByVal NewReport,Byref objData)
	
	
	If axis = "row" Then
	
		dataSourceValue=objData.item("datasource")
	    strDataSourceValue = "+ "&dataSourceValue       
	    strDatabaseValue=objData.item("database")
	    strCubeValue=objData.item("Cube")
	     
	    If NewReport = 1 Then
	    	Set objCreate=Browser("Analyzer").Page("Shared_Folder").Frame("view").Image("imgnewreport")        
	    	Call SCA.ClickOn(objCreate,"newreport","Report Creation")    
	    End If 
	     
	    
	         
	   'Capture Database Access Error - Start
	   If Browser("Analyzer").Frame("ErrorFrame").WebElement("Database access error.").Exist(1) Then
	       Set objErrorOk=Browser("Analyzer").Frame("ErrorFrame").WebButton("btnOK")
	       Call SCA.ClickOn(objErrorOk,"btnOK" , "Database Access Error during Report Creation")
	       Call ReportStep (StatusTypes.Fail,"Not Able to create a report. Database Access Error during Report Creation", "Report Creation Page")    
	       Report = 1
	       Exit Function
	   End If
	   'Capture Database Access Error - End
	   
	   'Handled Application Time Out in Report Creation Page - Start
	   Set objBtnok = Browser("Analyzer").Dialog("Message from webpage").WinButton("btnOK")
	   If objBtnok.Exist(5) Then
	        Call SCA.ClickOn(objBtnok,"btnOK" , "Application Time Out in Report Creation Page")
	   End If    
	   'Handled Application Time Out in Report Creation Page - End
	   wait 2
	   
	   Set objDataSource=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDataSources")   
	   Call SCA.SelectFromDropdown(objDataSource,strDataSourceValue)
	   
	   wait 2
	   
	   Set objDatabase=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDatabases")
	   Call SCA.SelectFromDropdown(objDatabase,strDatabaseValue)       
	   
	   Wait 2
	   Set objDatacubes=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlCubes")
	   Call SCA.SelectFromDropdown(objDatacubes,strCubeValue)
	   
	   Wait 5
	   
	   Set objAdd=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebButton("btnAdd")
	   Call SCA.ClickOn(objAdd,"AddButton" , "ReportCreation")	
		
	End If
	
	
	If axis <> "row" Then
		
		strCubeValue=objData.item("Cube")
		Set objAdd=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebButton("btnAdd")
	End If
	
	

	
	
    
	If axis = "row" Then
'		Hierarchys = objData.item("RowHierarchy")
		Hierarchys = HierarchyValue
	ElseIf axis = "column" Then	
'		Hierarchys = objData.item("ColumnHierarchy")
		Hierarchys = HierarchyValue
	ElseIf axis = "data" Then
'		Hierarchys = objData.item("DataHierarchy")
		Hierarchys = HierarchyValue
	End If
'	Hierarchy = objData.item("Hierarchy")
'	Cube = "SMP_IMS_M_TRN_0001"
	HierarcyValues = split(Hierarchys,";")
	
	
	
	
	HierarcyValues = split(Hierarchys,";")
	
	For i = 0 To ubound(HierarcyValues) Step 1
	    Set obj_tree=description.Create
         obj_tree("micclass").value="WebElement"
         obj_tree("class").value="TreeNode"
         obj_tree("innertext").value=HierarcyValues(i)
           wait 1
         set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
         intcount=objtreeCount.count
         If intcount<>0 Then
             k=intcount-1                       	
             objtreeCount(k).fireEvent "ondblclick"                           
             Call ReportStep (StatusTypes.Pass,"Clicked on the Webelement"&Space(2)&strChild_Values, "Report Creation Page")                    
             wait 3    
         Else
		'Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDataSource").WebElement("innertext:="&HierarcyValues(i),"class:=TreeNode").FireEvent "ondblclick"	
	     End If 
	Next
	wait 5
	Set obj_tree=description.Create
    obj_tree("micclass").value="WebElement"
    obj_tree("class").value="TreeNode"
    obj_tree("innertext").value=HierarcyValues(i-1)
    
      wait 3
    set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
    objtreeCount((objtreeCount.Count)-1).FireEvent "onmouseover"
    wait 4
    Setting.WebPackage("ReplayType") = 2
    objtreeCount((objtreeCount.Count)-1).rightclick
    Setting.WebPackage("ReplayType") = 1
    wait 4
    
   
    
'	
'	For i = 0 To ubound(HierarcyValues) Step 1
'		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDataSource").WebElement("innertext:="&HierarcyValues(i),"class:=TreeNode").FireEvent "ondblclick"	
'	Next
'	wait 5
'	
'	
'	If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDataSource").WebElement("innertext:="&HierarcyValues(i-1),"class:=TreeNode","index:=1").exist(2) Then
'		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDataSource").WebElement("innertext:="&HierarcyValues(i-1),"class:=TreeNode","index:=1").RightClick	
'	ElseIf Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDataSource").WebElement("innertext:="&HierarcyValues(i-1),"class:=trvSchemaSel SelectedNode","index:=1").Exist(2) Then
'		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDataSource").WebElement("innertext:="&HierarcyValues(i-1),"class:=trvSchemaSel SelectedNode","index:=1").RightClick
'	End If
	
	
	'select row, column and data axis
	If axis = "row" Then
		Set objrow=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webRowaxis")    
        Call SCA.ClickOn(objrow,"Rowaxis" , "ReportCreation")   
	ElseIf axis = "column" Then
		Set objcolumn=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcolumnaxis")
        Call SCA.ClickOn(objcolumn,"Columnaxis" , "ReportCreation")
	ElseIf axis = "data" Then
		Set objDataaxis=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webDataAxis")
        Call SCA.ClickOn(objDataaxis,"Dataaxis" , "ReportCreation")    
	End If
	
'	Browser("Analyzer").Page("Analyzer").Frame("ReportWindowOfAnalyzer").WebElement("apc").Click
	wait 2
	
	If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("class:=null","html tag:=SPAN").Exist(2) Then
		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("class:=null","html tag:=SPAN").RightClick
	Else	
		Setting.WebPackage("ReplayType") = 2
		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("class:=TreeNode","html tag:=SPAN","innertext:="&strCubeValue).RightClick
		Setting.WebPackage("ReplayType") = 1
	End If
	
'	Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("removecube").Click
	Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welRemoveCube").Click

	
	
'	Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebButton("Add").Click
	Call SCA.ClickOn(objAdd,"AddButton" , "ReportCreation")	
	
	
End Function


	



Public Function GroupTable (ByVal newreport,ByVal axis,Byref objData)
	
	
	If axis = "row" Then
	
	
	
	  dataSourceValue=objData.item("datasource")
	    strDataSourceValue = "+ "&dataSourceValue       
	    strDatabaseValue=objData.item("database")
	    strCubeValue=objData.item("Cube")
	
	
	  If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("class:=null","html tag:=SPAN").Exist(2) Then
	  		Setting.WebPackage("ReplayType") = 2
			Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("class:=null","html tag:=SPAN").RightClick
			Setting.WebPackage("ReplayType") = 1
		Else	
			Setting.WebPackage("ReplayType") = 2
			Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("class:=TreeNode","html tag:=SPAN","innertext:="&strCubeValue).RightClick
			Setting.WebPackage("ReplayType") = 1
		End If
		
	'	Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("removecube").Click
		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welRemoveCube").Click
	
	'	Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebButton("Add").Click
'		Call SCA.ClickOn(objAdd,"AddButton" , "ReportCreation")
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
		
	     
	     If newreport = 1 Then
	     	Set objCreate=Browser("Analyzer").Page("Shared_Folder").Frame("view").Image("imgnewreport")        
	    	Call SCA.ClickOn(objCreate,"newreport","Report Creation")    
	     
	     
		   'Capture Database Access Error - Start
		   If Browser("Analyzer").Frame("ErrorFrame").WebElement("Database access error.").Exist(1) Then
		       Set objErrorOk=Browser("Analyzer").Frame("ErrorFrame").WebButton("btnOK")
		       Call SCA.ClickOn(objErrorOk,"btnOK" , "Database Access Error during Report Creation")
		       Call ReportStep (StatusTypes.Fail,"Not Able to create a report. Database Access Error during Report Creation", "Report Creation Page")    
		       Report = 1
		       Exit Function
		   End If
		   'Capture Database Access Error - End
		   
		   'Handled Application Time Out in Report Creation Page - Start
		   Set objBtnok = Browser("Analyzer").Dialog("Message from webpage").WinButton("btnOK")
		   If objBtnok.Exist(5) Then
		        Call SCA.ClickOn(objBtnok,"btnOK" , "Application Time Out in Report Creation Page")
		   End If    
		   'Handled Application Time Out in Report Creation Page - End
		   wait 2
	   
	    End If
	   
	   Set objDataSource=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDataSources")   
	   Call SCA.SelectFromDropdown(objDataSource,strDataSourceValue)
	   
	   wait 2
	   
	   Set objDatabase=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDatabases")
	   Call SCA.SelectFromDropdown(objDatabase,strDatabaseValue)       
	   
	   Wait 2
	   Set objDatacubes=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlCubes")
	   Call SCA.SelectFromDropdown(objDatacubes,strCubeValue)
	   
	   Wait 5
	   
	   Set objAdd=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebButton("btnAdd")
	   
'	   If newreport = 1 Then
	   	   Call SCA.ClickOn(objAdd,"AddButton" , "ReportCreation")
'	   End If
	   
	   
'	   Call SCA.ClickOn(objAdd,"AddButton" , "ReportCreation")	
		
	End If
	
	
	If axis <> "row" Then
		
		strCubeValue=objData.item("Cube")
		Set objAdd=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebButton("btnAdd")
		
		If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("class:=null","html tag:=SPAN").Exist(2) Then
			Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("class:=null","html tag:=SPAN").RightClick
		Else	
			Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("class:=TreeNode","html tag:=SPAN","innertext:="&strCubeValue).RightClick
		End If
		
	'	Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("removecube").Click
		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welRemoveCube").Click
	
	'	Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebButton("Add").Click
		Call SCA.ClickOn(objAdd,"AddButton" , "ReportCreation")	
	End If
	
	

	
	
    
	If axis = "row" Then
		Hierarchys = objData.item("RowHierarchy")
	ElseIf axis = "column" Then	
		Hierarchys = objData.item("ColumnHierarchy")
	ElseIf axis = "data" Then
		Hierarchys = objData.item("DataHierarchy")
	End If
'	Hierarchy = objData.item("Hierarchy")
'	Cube = "SMP_IMS_M_TRN_0001"
	HierarcyValues = split(Hierarchys,";")
	
	For i = 0 To ubound(HierarcyValues) Step 1
		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDataSource").WebElement("innertext:="&HierarcyValues(i),"class:=TreeNode").FireEvent "ondblclick"	
	Next
	
	
	
	If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDataSource").WebElement("innertext:="&HierarcyValues(i-1),"class:=TreeNode","index:=1").exist(2) Then
		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDataSource").WebElement("innertext:="&HierarcyValues(i-1),"class:=TreeNode","index:=1").RightClick	
	ElseIf Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDataSource").WebElement("innertext:="&HierarcyValues(i-1),"class:=trvSchemaSel SelectedNode","index:=1").Exist(2) Then
		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDataSource").WebElement("innertext:="&HierarcyValues(i-1),"class:=trvSchemaSel SelectedNode","index:=1").RightClick
	End If
	
	
	'select row, column and data axis
	If axis = "row" Then
		Set objrow=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webRowaxis")    
        Call SCA.ClickOn(objrow,"Rowaxis" , "ReportCreation")   
	ElseIf axis = "column" Then
		Set objcolumn=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcolumnaxis")
        Call SCA.ClickOn(objcolumn,"Columnaxis" , "ReportCreation")
	ElseIf axis = "data" Then
		Set objDataaxis=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webDataAxis")
        Call SCA.ClickOn(objDataaxis,"Dataaxis" , "ReportCreation")    
	End If
	
'	Browser("Analyzer").Page("Analyzer").Frame("ReportWindowOfAnalyzer").WebElement("apc").Click
	wait 2
	
'	If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("class:=null","html tag:=SPAN").Exist(2) Then
'		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("class:=null","html tag:=SPAN").RightClick
'	Else	
'		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("class:=TreeNode","html tag:=SPAN","innertext:="&strCubeValue).RightClick
'	End If
'	
''	Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("removecube").Click
'	Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welRemoveCube").Click
'
'	
'	
''	Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebButton("Add").Click
'	Call SCA.ClickOn(objAdd,"AddButton" , "ReportCreation")	
'	
	
End Function

'#######
'Created by:srinivas
'Description:-This function is only for single script in phase 8
Public Function ReportNewFunction(ByVal axis,ByVal HierarchyValue,ByVal NewReport,Byref objData)
	
	
	If axis = "row" Then
	
		dataSourceValue=objData.item("datasource")
	    strDataSourceValue = "+ "&dataSourceValue       
	    strDatabaseValue=objData.item("database")
	    strCubeValue=objData.item("Cube")
	     
	    If NewReport = 1 Then
	    	Set objCreate=Browser("Analyzer").Page("Shared_Folder").Frame("view").Image("imgnewreport")        
	    	Call SCA.ClickOn(objCreate,"newreport","Report Creation")    
	    End If 
	     
	    
	         
	   'Capture Database Access Error - Start
	   If Browser("Analyzer").Frame("ErrorFrame").WebElement("Database access error.").Exist(1) Then
	       Set objErrorOk=Browser("Analyzer").Frame("ErrorFrame").WebButton("btnOK")
	       Call SCA.ClickOn(objErrorOk,"btnOK" , "Database Access Error during Report Creation")
	       Call ReportStep (StatusTypes.Fail,"Not Able to create a report. Database Access Error during Report Creation", "Report Creation Page")    
	       Report = 1
	       Exit Function
	   End If
	   'Capture Database Access Error - End
	   
	   'Handled Application Time Out in Report Creation Page - Start
	   Set objBtnok = Browser("Analyzer").Dialog("Message from webpage").WinButton("btnOK")
	   If objBtnok.Exist(5) Then
	        Call SCA.ClickOn(objBtnok,"btnOK" , "Application Time Out in Report Creation Page")
	   End If    
	   'Handled Application Time Out in Report Creation Page - End
	   wait 2
	   
	   Set objDataSource=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDataSources")   
	   Call SCA.SelectFromDropdown(objDataSource,strDataSourceValue)
	   
	   wait 2
	   
	   Set objDatabase=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlDatabases")
	   Call SCA.SelectFromDropdown(objDatabase,strDatabaseValue)       
	   
	   Wait 2
	   Set objDatacubes=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebList("lstddlCubes")
	   Call SCA.SelectFromDropdown(objDatacubes,strCubeValue)
	   
	   Wait 5
	   
	   Set objAdd=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebButton("btnAdd")
	   Call SCA.ClickOn(objAdd,"AddButton" , "ReportCreation")	
		
	End If
	
	
	If axis <> "row" Then
		
		strCubeValue=objData.item("Cube")
		Set objAdd=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebButton("btnAdd")
	End If
	
	

	
	
    
	If axis = "row" Then
'		Hierarchys = objData.item("RowHierarchy")
		Hierarchys = HierarchyValue
	ElseIf axis = "column" Then	
'		Hierarchys = objData.item("ColumnHierarchy")
		Hierarchys = HierarchyValue
	ElseIf axis = "data" Then
'		Hierarchys = objData.item("DataHierarchy")
		Hierarchys = HierarchyValue
	End If
'	Hierarchy = objData.item("Hierarchy")
'	Cube = "SMP_IMS_M_TRN_0001"
	HierarcyValues = split(Hierarchys,";")
	
	
	
	
	HierarcyValues = split(Hierarchys,";")
	
	For i = 0 To ubound(HierarcyValues) Step 1
	    Set obj_tree=description.Create
         obj_tree("micclass").value="WebElement"
         obj_tree("class").value="TreeNode"
         obj_tree("innertext").value=HierarcyValues(i)
           wait 1
         set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
         intcount=objtreeCount.count
         If intcount<>0 Then
             k=intcount-1                       	
             objtreeCount(k).fireEvent "ondblclick"                           
             Call ReportStep (StatusTypes.Pass,"Clicked on the Webelement"&Space(2)&strChild_Values, "Report Creation Page")                    
             wait 3    
         Else
		'Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDataSource").WebElement("innertext:="&HierarcyValues(i),"class:=TreeNode").FireEvent "ondblclick"	
	     End If 
	Next
	wait 5
	Set obj_tree=description.Create
    obj_tree("micclass").value="WebElement"
    obj_tree("class").value="TreeNode"
    obj_tree("innertext").value=HierarcyValues(i-1)
    
      wait 3
    set objtreeCount= Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").ChildObjects(obj_tree)
    objtreeCount((objtreeCount.Count)-1).FireEvent "onmouseover"
    wait 4
    objtreeCount((objtreeCount.Count)-1).rightclick
    wait 4
    
    'little changes by srinivas it wont impact the function
    If Browser("micclass:=browser").Page("micclass:=page").Frame("html id:=tabPages_frame2").WebElement("xpath:=//div[@id='trvSchema_1_1_2_23_1']//span[text()='SALES UNITS']").Exist(5) Then
    	Browser("micclass:=browser").Page("micclass:=page").Frame("html id:=tabPages_frame2").WebElement("xpath:=//div[@id='trvSchema_1_1_2_23_1']//span[text()='SALES UNITS']").FireEvent "onmouseover"
    	Setting.WebPackage("ReplayType") = 2
    	Browser("micclass:=browser").Page("micclass:=page").Frame("html id:=tabPages_frame2").WebElement("xpath:=//div[@id='trvSchema_1_1_2_23_1']//span[text()='SALES UNITS']").RightClick
    	Setting.WebPackage("ReplayType") = 1
    End If
    
'	
'	For i = 0 To ubound(HierarcyValues) Step 1
'		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDataSource").WebElement("innertext:="&HierarcyValues(i),"class:=TreeNode").FireEvent "ondblclick"	
'	Next
'	wait 5
'	
'	
'	If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDataSource").WebElement("innertext:="&HierarcyValues(i-1),"class:=TreeNode","index:=1").exist(2) Then
'		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDataSource").WebElement("innertext:="&HierarcyValues(i-1),"class:=TreeNode","index:=1").RightClick	
'	ElseIf Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDataSource").WebElement("innertext:="&HierarcyValues(i-1),"class:=trvSchemaSel SelectedNode","index:=1").Exist(2) Then
'		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebTable("tabDataSource").WebElement("innertext:="&HierarcyValues(i-1),"class:=trvSchemaSel SelectedNode","index:=1").RightClick
'	End If
	
	
	'select row, column and data axis
	If axis = "row" Then
		Set objrow=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webRowaxis")    
        Call SCA.ClickOn(objrow,"Rowaxis" , "ReportCreation")   
	ElseIf axis = "column" Then
		Set objcolumn=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webcolumnaxis")
        Call SCA.ClickOn(objcolumn,"Columnaxis" , "ReportCreation")
	ElseIf axis = "data" Then
		Set objDataaxis=Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("webDataAxis")
        Call SCA.ClickOn(objDataaxis,"Dataaxis" , "ReportCreation")    
	End If
	
'	Browser("Analyzer").Page("Analyzer").Frame("ReportWindowOfAnalyzer").WebElement("apc").Click
	wait 2
	
	If Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("class:=null","html tag:=SPAN").Exist(2) Then
		Setting.WebPackage("ReplayType") = 2
		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("class:=null","html tag:=SPAN").RightClick
		Setting.WebPackage("ReplayType") = 1
	Else
		Setting.WebPackage("ReplayType") = 2	
		Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("class:=TreeNode","html tag:=SPAN","innertext:="&strCubeValue).RightClick
		Setting.WebPackage("ReplayType") = 1
	End If
	
'	Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("removecube").Click
	Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebElement("welRemoveCube").Click

	
	
'	Browser("Analyzer").Page("ReportCreation").Frame("ReportWindowOfAnalyzer").WebButton("Add").Click
	Call SCA.ClickOn(objAdd,"AddButton" , "ReportCreation")	
	
	
End Function


	






















End Class


'function by srinivas for converting color code component
Sub CorrectRGBComponent(ByRef component)
  component = aqConvert.VarToInt(component)
  If (component < 0) Then
    component = 0
  Else
    If (component > 255) Then
      component = 255
    End If
  End If
End Sub

'function by srinivas just pass RGB Values
Function RGB(r, g, b)
  Call CorrectRGBComponent(r)
  Call CorrectRGBComponent(g)
  Call CorrectRGBComponent(b)
  RGB = r + (g * 256) + (b * 65536)
End Function





