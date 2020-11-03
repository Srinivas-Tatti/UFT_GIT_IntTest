'.########################################################################################################################################
'# Project Name: OCRF
'# Purpose/Description: Please refer this test script specific function in OCRFScenarios library file
'# Pre Conditions: Please refer this test script specific function in OCRFScenarios library file if any
'#########################################################################################################################################

Environment("TC")=0

Call objINIFile.InitializeVariables_SCAWeb()

Call IMS_Framework_TestCaseHandler.ExecuteTestCases()

				
'Call ImportSheet("NonLive_NonSynd_DataSource","SIT_LIVE_NONLIVE_DATA")

'Browser("micclass:=browser").Page("micclass:=page").WebElement("xpath:=//div[contains(@style,'display: block;')]/../descendant::span[text()='Okay']").click

'Browser("micclass:=browser").Page("micclass:=page").WebList("xpath:=//div[@id='pg_user-list_toppager']//select[@class='ui-pg-selbox']").Select "100"




'Browser("micclass:=browser").Page("micclass:=page").WebFile("html tag:=INPUT","class:=ms-fileinput").set "C:\Data Backup\C Drive Backup\Automation_BaseLineDataSetup\MergedSCAFW\QAFramework3x\files\OCRF_SyndicatedDC_23.xlsx"



