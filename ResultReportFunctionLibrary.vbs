Public ResultsFolder,reportFile,TestStepStartTime,reportFilePath
Public  totalConsolidatedPassed , totalConsolidatedFailed , totalConsolidatedSkipped, totalConsolidatedTestCases, consolidatedReportStartTime, consolidatedReportEndTime
Public DBConnection_Results,DBConnection_Repository
Public DBConnection_TestData,IterationCount
Public ExecutionStartTime,ExecutionEndTime
'Public ExecutionStartTime,ExecutionEndTime
Public RepositoryObject,HTMLResultsStreamWriter	
Public testparameters, componentparameters,compIteration
Public API_Tests_Executed
Set testparameters  = CreateObject("Scripting.Dictionary")
totalConsolidatedPassed =0
totalConsolidatedFailed = 0
totalConsolidatedSkipped = 0
totalConsolidatedTestCases = 0
Public Function InitializeConsolidatedReport()
Set obFSO = CreateObject("Scripting.FileSystemObject")
		' Create a new colsolidated results folder
	ConsolidatedResultsFolder = Environment("RESULTS_FOLDER") & "\ConsolidatedReport_" & Replace(FormatDateTime(Date),"/","-") & " " & Replace(FormatDateTime(Time),":","")
	obFSO.CreateFolder ConsolidatedResultsFolder
	If Not obFSO.FolderExists(ConsolidatedResultsFolder) Then
		strError = "Consolidated Results folder is not created in the path : " & ConsolidatedResultsFolder & "     :-" & Err.Description
		Set obFSO = Nothing
		Exit Function
	End If
	Environment("CONSOLIDATED_RESULTS_FOLDER") = ConsolidatedResultsFolder
	obFSO.CreateFolder ConsolidatedResultsFolder&"\img"
	consolidatedResultsHTMLPath = ConsolidatedResultsFolder & "\ConsolidatedResults.html"
	Environment("CONSOLIDATED_RESULTS_HTML") = consolidatedResultsHTMLPath
	baxterLogoPath = Environment.Value("RESOURCES_IMG") & "\" &  "transparentLogo.png"
		tArray = split(baxterLogoPath,"\")
	baxterLogoFile = Environment.Value("CONSOLIDATED_RESULTS_FOLDER")  & "\img\" & tArray(UBOUND(tArray))

	TemplateBaxterLogoFile = Environment("RESOURCES_IMG") & "\" & tArray(UBOUND(tArray))
	If Not obFSO.FileExists(TemplateBaxterLogoFile) Then
		strError = "TemplateBaxterLogoFile is not available in the path : " & TemplateBaxterLogoFile 
		Exit Function
	End If

	obFSO.CopyFile TemplateBaxterLogoFile, Environment.Value("CONSOLIDATED_RESULTS_FOLDER") & "\img\"
	If Not obFSO.FileExists(baxterLogoFile) Then
		Reporter.ReportEvent micFail , "Copy template file to result folder ", "Failed to copy baxterLogoFile to "& ConsolidatedResultsFolder
		Exit Function
	End If
	
	baxterBackGroundPath = Environment.Value("RESOURCES_IMG") & "\" &  "BaxterBackGround.png"
tArray = split(baxterBackGroundPath,"\")
	baxterBackGroundFile = Environment.Value("CONSOLIDATED_RESULTS_FOLDER")  & "\img\" & tArray(UBOUND(tArray))

	TemplateBaxterBackGroundFile = Environment("RESOURCES_IMG") & "\" & tArray(UBOUND(tArray))
	If Not obFSO.FileExists(TemplateBaxterBackGroundFile) Then
		strError = "TemplatebaxterBackGroundPath file is not available in the path : " & TemplateBaxterBackGroundFile 
		Exit Function
	End If

	obFSO.CopyFile TemplateBaxterBackGroundFile, Environment.Value("CONSOLIDATED_RESULTS_FOLDER") & "\img\"
	If Not obFSO.FileExists(baxterBackGroundFile) Then
		Reporter.ReportEvent micFail , "Copy template file to result folder ", "Failed to copy baxterBackGroundPath to "& ConsolidatedResultsFolder
		Exit Function
	End If
	
	Call InitializeConsolidatedReportHTML(consolidatedResultsHTMLPath)
	If Err.Number <> 0 Then   
		 strError = "Unable to create " & consolidatedResultsHTMLPath & VbCrLf & Err.Description
		 InitializeConsolidatedReport = False
		 Set obFSO = Nothing
		Exit Function
	End If 
	InitializeConsolidatedReport = True
	Set obFSO = Nothing
End  Function
Public Function InitializeReport()
        
	On Error Resume Next
	Dim TemplateFile,ResultsDBFile,objFSO, addtimestamp, ResultsHtmFile
	strError = ""
   	InitializeReport = False
								
	Set obFSO = CreateObject("Scripting.FileSystemObject")

	ResultsFolder = Environment("RESULTS_FOLDER")
	timestamp = Replace(FormatDateTime(Date),"/","-") & " " & Replace(FormatDateTime(Time),":","")
	addtimestamp = "ExecRes_" & Environment.Value("CURRENT_TEST_SLNO") & "_" & Environment.Value("CURRENT_TESTCASE_NAME") & "_" & Environment.Value("CURRENT_TEST_ITERATION")  & "_"& timestamp
	
	' Create a new folder
	ResultsFolder = ResultsFolder & "\" & addtimestamp
	obFSO.CreateFolder ResultsFolder
	If Not obFSO.FolderExists(ResultsFolder) Then
		strError = "Results folder is not created in the path : " & ResultsFolder & "     :-" & Err.Description
		Exit Function
	End If
	
	Environment("CURRENT_RESULTS_FOLDER") = ResultsFolder
	
	obFSO.CreateFolder ResultsFolder&"\img"
	
	statusACCDB_TEMPLATE = DownloadResourceFromQC (Environment("ACCDB_TEMPLATE"), "RESOURCES")
	If UCASE(statusACCDB_TEMPLATE)="ERROR" Then
		Reporter.ReportEvent micFail , "ACCDB_TEMPLATE Download" , "Failed to download ACCDB_TEMPLATE : -"
		Exit Function
	End If
	
	If UCASE(TRIM(Environment("TYPE_OF_TEST"))) <> "API" Then
		statusHTMLRESULTS_TEMPLATE = DownloadResourceFromQC (Environment("HTMLRESULTS_TEMPLATE"), "RESOURCES")
		If UCASE(statusHTMLRESULTS_TEMPLATE)="ERROR" Then
			Reporter.ReportEvent micWarning , "HTMLRESULTS_TEMPLATE Download" , "Failed to download HTMLRESULTS_TEMPLATE : -"
			If Not (Environment("GenerateHTMLResult") = "YES" or Environment("GenerateHTMLResult") = "Y" )Then
				Exit Function
			End If
		End If
	End  If
	

	
	tArray = split(statusACCDB_TEMPLATE,"\")
	ResultsDBFile = Environment.Value("CURRENT_RESULTS_FOLDER")  & "\" & tArray(UBOUND(tArray))
	TemplateFile = Environment("RESOURCES_FOLDER") & "\" & tArray(UBOUND(tArray))
	obFSO.CopyFile TemplateFile, Environment.Value("CURRENT_RESULTS_FOLDER") & "\"
	If Not obFSO.FileExists(TemplateFile) Then
		strError = "Results_DBaseFile.accdb file is not available in the path : " & TemplateFile 
		Exit Function
	End If

	obFSO.CopyFile TemplateFile, Environment.Value("CURRENT_RESULTS_FOLDER") & "\"
	If Not obFSO.FileExists(ResultsDBFile) Then
		Reporter.ReportEvent micFail , "Copy template file to result folder ", "Failed to copy DB template to "& ResultsFolder
		Exit Function
	End If
If UCASE(TRIM(Environment("TYPE_OF_TEST"))) <> "API" Then
	tArray = split(statusHTMLRESULTS_TEMPLATE,"\")
	TemplateFile = Environment("RESOURCES_FOLDER") & "\" & tArray(UBOUND(tArray))
	If Not obFSO.FileExists(TemplateFile) Then
		strError = "ResultsHtmFile file is not available in the path : " & TemplateFile 
		Exit Function
	End If

	ResultsHtmFile = Environment.Value("CURRENT_RESULTS_FOLDER") & "\" & tArray(UBOUND(tArray))
	obFSO.CopyFile TemplateFile, Environment.Value("CURRENT_RESULTS_FOLDER") & "\"
	If Not obFSO.FileExists(ResultsHtmFile) Then
		strError = "ResultsHtmFile file is not available in the path : " & ResultsHtmFile 
		Exit Function
	End If
End  If
					strTempFolder = addtimestamp
					strTempfileOutput = "OutputValues.xls" 
					If obFSO.FileExists(strTempfileOutput) Then
							obFSO.DeleteFile (strTempfileOutput)
							If obFSO.FileExists(strTempfileOutput) Then
									strError = "Unable to delete existing Output value file in the path : " & strTempfileOutput
									Exit Function
							End If
					End if 
            
					TestDataOutputFile = Environment("OUTPUTVALUE_FILE") 
					tArray = split(TestDataOutputFile,"\")
					TestDataOutputFile = Environment.Value("CURRENT_RESULTS_FOLDER")  & "\" & tArray(UBOUND(tArray))					
					statusOUTPUTVALUE_FILE = DownloadResourceFromQC (Environment("OUTPUTVALUE_FILE"), "RESOURCES")
					If UCASE(statusOUTPUTVALUE_FILE)="ERROR" Then
						Reporter.ReportEvent micFail , "OUTPUTVALUE_FILE Download" , "Failed to download OUTPUTVALUE_FILE"
					Else
						obFSO.CopyFile statusOUTPUTVALUE_FILE, Environment.Value("CURRENT_RESULTS_FOLDER") & "\"
					End If 
					If Not obFSO.FileExists(TestDataOutputFile) Then						
							strError="Output value file not present in the path - " & TestDataOutputFile
							Exit Function
					End If	
	'Connect to Results_DBaseFile.accdb file
	Dim dbProvider,dbSource,strSQL,rs
	dbProvider = "PROVIDER=Microsoft.ACE.OLEDB.12.0;"
	dbSource = "Data Source = " + ResultsDBFile
	Set DBConnection_Results = CreateObject("ADODB.Connection")
	' Connect to the database
	DBConnection_Results.Open dbProvider & dbSource
	If Err.Number <> 0 Then   
		strError = "Unable to open the connection, dbSource : " & dbSource & VbCrLf & Err.Description
		Exit Function
	End If 


	strSQL = "Delete * from TestResults"
	Set rs = DBConnection_Results.Execute(strSQL)
	baxterLogoPath = Environment.Value("RESOURCES_IMG") & "\" &  "transparentLogo.png"
		tArray = split(baxterLogoPath,"\")
	baxterLogoFile = Environment.Value("CURRENT_RESULTS_FOLDER")  & "\img\" & tArray(UBOUND(tArray))

	TemplateBaxterLogoFile = Environment("RESOURCES_IMG") & "\" & tArray(UBOUND(tArray))
	If Not obFSO.FileExists(TemplateBaxterLogoFile) Then
		strError = "TemplateBaxterLogoFile is not available in the path : " & TemplateBaxterLogoFile 
		Exit Function
	End If

	obFSO.CopyFile TemplateBaxterLogoFile, Environment.Value("CURRENT_RESULTS_FOLDER") & "\img\"
	If Not obFSO.FileExists(baxterLogoFile) Then
		Reporter.ReportEvent micFail , "Copy template file to result folder ", "Failed to copy baxterLogoFile to "& ResultsFolder
		Exit Function
	End If
	
	baxterBackGroundPath = Environment.Value("RESOURCES_IMG") & "\" &  "BaxterBackGround.png"
tArray = split(baxterBackGroundPath,"\")
	baxterBackGroundFile = Environment.Value("CURRENT_RESULTS_FOLDER")  & "\img\" & tArray(UBOUND(tArray))

	TemplateBaxterBackGroundFile = Environment("RESOURCES_IMG") & "\" & tArray(UBOUND(tArray))
	If Not obFSO.FileExists(TemplateBaxterBackGroundFile) Then
		strError = "TemplatebaxterBackGroundPath file is not available in the path : " & TemplateBaxterBackGroundFile 
		Exit Function
	End If

	obFSO.CopyFile TemplateBaxterBackGroundFile, Environment.Value("CURRENT_RESULTS_FOLDER") & "\img\"
	If Not obFSO.FileExists(baxterBackGroundFile) Then
		Reporter.ReportEvent micFail , "Copy template file to result folder ", "Failed to copy baxterBackGroundPath to "& ResultsFolder
		Exit Function
	End If
	
	If Err.Number <> 0 Then   
'			Msgbox "Error while executing the SQL" & Err.Description
		    strError = "Unable to execute the SQL : " & strSQL & VbCrLf & Err.Description
			Exit Function
	End If 

	Set rs = Nothing    
	Set obFSO = Nothing
	InitializeReport = True
	ExecutionStartTime = Now

End Function

Public Sub ReportResult()

	On Error Resume Next

	Dim strSQL,rs,FieldList,ValuesList

	Const adVarWChar = 202
	Const adSingle = 4
	Const adLockOptimistic = 3
	Const adOpenDynamic = 2


	FieldList = Array("SNo","TestCaseName", "Iteration", "StepName","StepDescription", "ExpectedResult","AutomationStepDescription","StepResult","Duration","FailureDescription", "Screenshot", "Param1", "Param2")
	ValuesList = Array(valslNo,valTestCaseName,strIteration,valStepName, mid(valStepDescription,1,254), mid(valExpectedResult,1,254),mid(valAutomationStepDescription,1,254),StepResult,stepDuration,mid(actualResult,1,254),stepScreenshot,isNullisEmptyCheck(valParam1),isNullisEmptyCheck(valParam2))

	Set rs = CreateObject("ADODB.Recordset")
	rs.Open "TestResults", DBConnection_Results, adOpenDynamic, adLockOptimistic
	rs.AddNew FieldList,ValuesList
	rs.Update
	rs.Save
	If err.Number <> 0 Then   
'			 MsgBox err.description
			 rs.Close    
			 Set rs = Nothing    
			 Exit Sub
	End If

	rs.Close
	Set rs = Nothing

End Sub

Public Sub ReportTestResult()
	
	if (UCASE(Environment.Value("TRIGGER_FROMQC")) = "YES" or UCASE(Environment.Value("TRIGGER_FROMQC"))="Y") Then
	
		Set objQCCurrentTestSet = QCUtil.CurrentTestSet
		strCurrentTestSetid = objQCCurrentTestSet.id
		Set var_CurrentTest = QCUtil.CurrentTest
		setcurrentTestid=var_CurrentTest.Field("TS_TEST_ID")
	
	Else 
	
		strCurrentTestSetid = Environment.Value("CURRENT_TEST_SLNO")
		setcurrentTestid = Environment.Value("CURRENT_TEST_ITERATION")
		TestcaseName = Environment.Value("CURRENT_TESTCASE_NAME") 

	End if
		TestDescription = Environment.Value("CURRENT_TEST_DESCRIPTION")
	On Error Resume Next

	Dim strSQL,rs,FieldList,ValuesList

	Const adVarWChar = 202
	Const adSingle = 4
	Const adLockOptimistic = 3
	Const adOpenDynamic = 2

	FieldList = Array("TestSetid","TestCaseid","TestCase", "TestDescription","TestResult", "TestDuration")
	ValuesList = Array(strCurrentTestSetid,setcurrentTestid,TestcaseName,TestDescription,TestResult,GetExecutionTime(TestStartTime,TestEndTime))

	Set rs = CreateObject("ADODB.Recordset")
	rs.Open "CompleteResults", DBConnection_Results, adOpenDynamic, adLockOptimistic
	rs.AddNew FieldList,ValuesList
	rs.Update
	rs.Save

	If err.Number <> 0 Then   
		 rs.Close    
		 Set rs = Nothing    
		 Exit Sub
	End If

	rs.Close
	Set rs = Nothing
	Set var_CurrentTest = Nothing
	Set objQCCurrentTestSet = Nothing

End Sub

Function ReportToQC()

			On Error Resume Next
			Err.Clear
			If Trim(Ucase( Environment.Value("TRIGGER_FROMQC"))) = "NO" Then
				If mstepResult = "Passed" Then
					Reporter.ReportEvent micPass,stepDescription,stepExpected
				Else
					Reporter.ReportEvent micFail,stepDescription,actualResult1
				End If
				Exit Function
			End If 

			Dim myCurentRun,myStepFactory,myStepList,nStepKey,ObjCurrentTest,AttachmentFactory,ObjAttch

			Set myCurentRun = QCUtil.CurrentRun
			Set myStepFactory = myCurentRun.StepFactory
			myStepFactory.AddItem(valStepName)
			Set myStepList = myStepFactory.NewList("")
			nStepKey = myStepList.Count 'This sets the step count
			myStepList.Item(nStepKey).Field("ST_STATUS") = mstepResult
			myStepList.Item(nStepKey).Field("ST_DESCRIPTION") = stepDescription
			myStepList.Item(nStepKey).Field("ST_EXPECTED") = stepExpected
			myStepList.Item(nStepKey).Field("ST_ACTUAL") = actualResult1
			myStepList.Post

			If  Not stepScreenshot = ""  Then
				Set ObjCurrentTest = QCUtil.CurrentRun.StepFactory.NewList("").Item(nStepKey) 
				Set AttachmentFactory = ObjCurrentTest.attachments
				Set ObjAttch = AttachmentFactory.AddItem(null)
				ObjAttch.FileName = stepScreenshot
				ObjAttch.Type = 1
				ObjAttch.Post
				ObjAttch.Refresh 
			End If

            If  Not stepScreenshot1 = ""  Then

						zipFolder = Environment.Value("CURRENT_TESTCASE_NAME") & "_" & Environment.Value("CURRENT_TEST_ITERATION") & valStepName
						Set objFol = CreateObject("Scripting.FileSystemObject")
						Set objFolder = objFol.GetFolder(ResultsFolder)
						Set objFiles = objFolder.Files
						
						Set newReg = new RegExp
						newReg.Global = True
						newReg.IgnoreCase = True
						newReg.Pattern = Environment.Value("CURRENT_TESTCASE_NAME") & "_" & Environment.Value("CURRENT_TEST_ITERATION") & valStepName & ".*"
						newFol = ResultsFolder & "\" & zipFolder 
						If not objFol.FolderExists(newFol) Then
						
							Set newFol = objFol.CreateFolder(newFol)
						End If
                        						
						For each filename in objFiles
						
							If (newReg.Test(filename.name)=True) Then
						
									reqFile =  filename.name			
									reqFilePath = filename.path			
									objFol.CopyFile reqFilePath, newFol & "\", true
						
							End If
						
						Next
				
						sFile1 = newFol & ".zip"
						mvar1 = """"&sFile1&""""
							
						sFile3 =  newFol & "\*.png"
						mvar3 = """"&sFile3&""""
				
						While  now <= lastModifiedTime
						Wend
						
						w7zipfldr="C:\Program Files\7-Zip"
							Set oShell =CreateObject("WScript.Shell")
						oShell.Run  chr(34) &  w7zipfldr & "\7z.exe" & chr(34) & " a "& mvar1 & " "& mvar3 
						fileToAttach = sFile1
				
						Set ObjCurrentTest = QCUtil.CurrentRun.StepFactory.NewList("").Item(nStepKey) 
						Set AttachmentFactory = ObjCurrentTest.attachments
						Set ObjAttch = AttachmentFactory.AddItem(null)
						ObjAttch.FileName = fileToAttach
						ObjAttch.Type = 1
						ObjAttch.Post
						ObjAttch.Refresh 

			End if

			stepScreenshot = ""
			stepScreenshot1 = ""

			Set myStepList = Nothing
			Set myStepFactory = Nothing
			Set myCurentRun = Nothing
			Set ObjCurrentTest = Nothing
			Set AttachmentFactory = Nothing
			Set ObjAttch = Nothing

End Function

Public Function addAPIRequestResponseInZip(zipExe,ResultsFolder, strFolder ,fullZip)

strZip = zipFileName

'Create the command line string
strCommand = "cmd /c cd /d " & chr(34) & ResultsFolder & chr(34) &" "& chr(38) & " " & chr(34) &  zipExe & chr(34) & " a "  & chr(34) & fullZip & chr(34) & " " & chr(34) & strFolder & chr(92) & chr(42) & chr(34) & " -r"

'Execute the command
Set objShell_folder = CreateObject("WScript.Shell")
Set objExec = objShell_folder.Run (strCommand, 1, True)
Reporter.ReportEvent micDone, "cmd Execution command of " & strCommand , objExec.StdErr.ReadAll
Set objShell_folder = Nothing
Set objExec = nothing
	If err.number<>0 Then
		err.clear
		Exit Function
	End If
End Function

Public sub AttachResultsToQC()   

	Dim strArguments
	Dim strZipApp,sFile,obFSO,zipFileName

	On Error Resume Next
	Err.Clear

	Set obFSO = CreateObject("Scripting.FileSystemObject") 
	zipFileName = obFSO.GetBaseName(ResultsFolder)

	folderspec=ResultsFolder
	filenames=ShowFolderList(folderspec)
	If instr(1,filenames,"pdf") Then
		pdfFileExist="True"
	End If

	sFile1 = ResultsFolder & "\" & zipFileName & ".zip"
	mvar1 = """"&sFile1&""""
	sFile2 = ResultsFolder & "\*.htm"
	mvar2 = """"&sFile2&""""

	sFile3 = ResultsFolder & "\*.png"
	mvar3 = """"&sFile3&""""

	sFile4 = ResultsFolder & "\*.pdf"
	mvar4 = """"&sFile4&""""

	sFile5 = ResultsFolder & "\*.xls"
	mvar5 = """"&sFile5&""""


	Programfilespath = Environment("ZIPSW_PATH")
	winzipfldr=Programfilespath &"\WinZip"
	w7zipfldr=Programfilespath &"\7-Zip"
	Set oShell =CreateObject("WScript.Shell")

	If (obFSO.FolderExists(w7zipfldr)) Then
		If pdfFileExist="True" Then
			
			oShell.Run  chr(34) &  w7zipfldr & "\7z.exe" & chr(34) & " a "& mvar1 &" "& mvar2 & " "& mvar3 &" "& mvar4&" "& mvar5
		else
		
			oShell.Run  chr(34) &  w7zipfldr & "\7z.exe" & chr(34) & " a "& mvar1 &" "& mvar2 & " "& mvar3 &" "& mvar5
		End If
	ElseIf (obFSO.FolderExists(winzipfldr)) Then
'		Set aShell = CreateObject("Wscript.Shell")
		If pdfFileExist="True" Then
			'aShell.Exec("" & "C:\Program Files\WinZip\WINZIP32.exe" & "" & " " & " -a " & mvar1 & "  " & mvar2)' & "  " & mvar3)
			oShell.Exec("" & "C:\Program Files\WinZip\WINZIP32.exe" & "" & " " & " -a " & mvar1 & "  " & mvar2 & "  " & mvar3 & "  " & mvar4&" "& mvar5)
		Else
			oShell.Exec("" & "C:\Program Files\WinZip\WINZIP32.exe" & "" & " " & " -a " & mvar1 & "  " & mvar2 & "  " & mvar3&" "& mvar4)
		End if
	End If
	wait(conTen)
	
	If API_Tests_Executed Then
		zipExe =  w7zipfldr & "\7z.exe" 
		requestFolder =  "Requests"
		responsesFolder = "\Responses"
		fullZip =   zipFileName & ".zip"
		Call 	addAPIRequestResponseInZip(zipExe, ResultsFolder, requestFolder , fullZip)
		Call 	addAPIRequestResponseInZip(zipExe, ResultsFolder, responsesFolder ,fullZip)
	End If
	
	
'	If Err.Number <> 0 Then
'		Reporter.ReportEvent micFail, "Create Zip ", "Unable to create Zip:- " & err.description
'		Exit Sub
'	End If
	

	
	If QCUtil.IsConnected = true Then

					Dim objnowRun,attachmentPath,nowAttachment
					Set objnowRun = QCUtil.CurrentRun
					Set attachmentPath = objnowRun.Attachments
					Set nowAttachment = attachmentPath.AddItem(Null)
					nowAttachment.FileName = sFile1
					nowAttachment.Type = 1
					nowAttachment.Post()

					objnowRun.Refresh

					Set objnowRun = nothing
					Set attachmentPath = nothing
					Set nowAttachment = nothing
	End If
	If Err.Number <> 0 Then   
		 Err.Clear
		 Exit Sub
	End If


	Set obFSO = Nothing
	Set oShell = Nothing

End Sub

Public Sub CompleteReport()
   	ExecutionEndTime = Now
   	If UCASE(TRIM(Environment("TYPE_OF_TEST"))) <> "API" Then
   		AddTestResultToHTML
   	End If
	If API_Tests_Executed Then
		AddAPITestResultToHTML
	End If
	'DBConnection_Results.Close
	'Set DBConnection_Results = Nothing
	If Not(Trim(Ucase( Environment("TRIGGER_FROMQC"))) = "NO") Then
	AttachResultsToQC
	End If 
	call AddTestsToConsolidatedHTMLReport()
End Sub


Public Sub AddTestResultToHTML()

	On Error Resume Next
	Dim strSQL,rs,strFile,objFSO
	Dim Count_TC,Count_Passed,Count_Failed,Count_NotRun
	Count_TC = 0
	Count_Passed = 0
	Count_Failed = 0
	Count_NotRun = 0

	Const ForAppending = 8

	'Opens the Execution Results.htm file
	strFile = ResultsFolder & "\ExecutionResults.htm"
					
	set objFSO = CreateObject("Scripting.FileSystemObject")					
	set HTMLResultsStreamWriter = objFSO.OpenTextFile(strFile, ForAppending, True)

	'Gets the test execution results
	strSQL = "SELECT * FROM CompleteResults"
	Set rs = DBConnection_Results.Execute(strSQL)

	Do While not rs.eof

			Dim ttestsetid, ttestcaseid, tName,tResult,tDuration
			Count_TC = Count_TC + 1
			ttestsetid	=	rs.Fields.Item("TestSetid").Value
			ttestcaseid	=	rs.Fields.Item("TestCaseid").Value
			tName		=	rs.Fields.Item("TestCase").Value
			tDescription =  rs.Fields.Item("TestDescription").Value
			tResult		=	rs.Fields.Item("TestResult").Value
			tDuration 	=	rs.Fields.Item("TestDuration").Value

			If tResult = "Passed" Then
				Count_Passed = Count_Passed + 1
			ElseIf tResult = "Failed" Then
				Count_Failed = Count_Failed + 1
			ElseIf tResult = "Not Run" Then
				Count_NotRun = Count_NotRun + 1
			End If

			Call AddStepResultsToHTML(ttestsetid,ttestcaseid, tName,tDescription,tResult,tDuration,Count_TC)
			rs.MoveNext 
	Loop

	'Finishes the HTML Report
	HTMLResultsStreamWriter.WriteLine("</table>")
	
	'HTMLResultsStreamWriter.WriteLine("<footer style=""margin:5% 0%;""><span style=""float: right;"">Baxter Confidential - Do not distribute with out prior approval</span>"&vblf&"</footer>" )
	' <img src = ""./img/transparentLogo.png"" alt=""BaxterFooterLogo"" width=""90"" height=""20"" style=""float: center; vertical-align:bottom; "">"&vblf& &vblf&"</img>")
	HTMLResultsStreamWriter.Close()
	'HTMLResultsStreamWriter = Nothing

	Set rs = Nothing    
	Set HTMLResultsStreamWriter = Nothing
	ReplaceSummary Count_Passed,Count_Failed,Count_NotRun,Count_TC
End Sub

Public Sub AddStepResultsToHTML(tTestSetid,tTestCaseid,TestName,tDescription,TResult,TDuration,Count_TC)

	Dim strSQL,rs
	'Gets the test Iteration results
	strSQL = "SELECT Iteration FROM TestResults WHERE TestCaseName='" & TestName & "' GROUP BY Iteration"
	Set rs = DBConnection_Results.Execute(strSQL)

	Dim bgColor

	If Count_TC Mod 2 = 1 Then
		bgColor = "#FFFFFF"
	Else
		bgColor = "#D6EBFC"
	End If

	Dim failColor

	If TResult = "Passed" Then
		failColor = "#00D100"
	ElseIf TResult = "Failed" Then
		failColor = "#FF0000"
	ElseIf TResult = "Not Run" Then
		failColor = "#F6C6AD"
	End If

	'Writes out the full test execution result
			' "<td style=""vertical-align: middle; text-align: center; font-family: Calibri; padding: 0px; margin: 0px; background-color: " & bgColor & ";""> " &_
		'	  Count_TC & "</td>" & vbLf &_
	HTMLResultsStreamWriter.Write(_
		"<tr>" & vbLf &_
		"<td id=""main" & Count_TC & """ style=""vertical-align: middle; text-align: center; font-family: Calibri; padding: 0px; margin: 0px; background-color: " & bgColor & ";""> " &_
			"&nbsp;<a href=""javascript:void(0)"" onclick=""toggle(" & Count_TC & ", 'open')"" class=""style7"">Expand Table</a>&nbsp;</td>" & vbLf &_
		 "<td style=""vertical-align: middle; text-align: center; font-family: Calibri; padding: 0px; margin: 0px; background-color: " & bgColor & ";""> " &_
			tTestSetid & "</td>" & vbLf &_
		 "<td style=""vertical-align: middle; text-align: center; font-family: Calibri; padding: 0px; margin: 0px; background-color: " & bgColor & ";""> " &_
			TestName & "</td>" & vbLf &_
		 "<td style=""vertical-align: middle; text-align: center; font-family: Calibri; padding: 0px; margin: 0px; background-color: " & bgColor & ";""> " &_
			 tTestCaseid & "</td>" & vbLf &_
		"<td style=""vertical-align: middle; text-align: center; font-family: Calibri; padding: 0px; margin: 0px; background-color: " & bgColor & ";""> " &_
			tDescription & "</td>" & vbLf &_
		 "<td style=""vertical-align: middle; text-align: center; font-family: Calibri; padding: 0px; margin: 0px; background-color: " & bgColor & "; color: " & failColor & ";""> " &_
			TResult & "</td>" & vbLf &_
		 "<td style=""vertical-align: middle; text-align: center; font-family: Calibri; padding: 0px; margin: 0px; background-color: " & bgColor & ";""> " &_
		  TDuration & "</td>" & vbLf &_
		"</tr>" & vbLf)

	HTMLResultsStreamWriter.WriteLine(_
		 "<tr id=""subItem" & Count_TC & """ style=""display :none ; width :100%; position : relative;"" > " & vbLf &_
					 "<td colspan=""6"" style=""vertical-align: middle; text-align: center; font-family: Calibri;  margin: 0px; background-color: #FFFFFF;""  > " & vbLf)

	Do While not rs.eof

			'Creates a header for the sub table in the report
			HTMLResultsStreamWriter.WriteLine("<p></p>")
			Dim citeration
			citeration = rs.Fields.Item("Iteration").Value

			'HTMLResultsStreamWriter.WriteLine(_
			'	"<p style=""font-family: 'Trebuchet MS'; font-size: small; font-weight: bold; text-decoration: underline; text-align: left"">Iteration " & citeration & "</p>" & vbLf)


			HTMLResultsStreamWriter.WriteLine(_
				"<table align=""center"" style=""border-style: ridge; border-width: 2px; width:800px; position: absolute;"" cellpadding=""0"" cellspacing=""0"" > " & vbLf &_
				"<tr>" & vbLf &_
				   "<th width=""5%"" style=""border-style: solid; border-width: 1px; background-color: #00B5F0"">Steps</th> " & vbLf &_
				   "<th width=""20%"" style=""border-style: solid; border-width: 1px; background-color: #00B5F0"">Step Description</th> " & vbLf &_
				   "<th width=""35%"" style=""border-style: solid; border-width: 1px; background-color: #00B5F0"">Expected Result</th> " & vbLf &_
				   "<th width=""35%"" style=""border-style: solid; border-width: 1px; background-color: #00B5F0"">Actual Result</th> " & vbLf &_
				   "<th width=""10%"" style=""border-style: solid; border-width: 1px; background-color: #00B5F0"">Status </th> " & vbLf &_
				   "<th style=""border-style: solid; border-width: 1px; background-color: #00B5F0"">Screenshot</th> " & vbLf &_
			   "</tr>" & vbLf)
					'"<th style=""border-style: solid; border-width: 1px; background-color: #00B5F0"">Failure Description</th> " & vbLf &_
					
			Dim rs_new
			'Adds the test results step-by-step 
			strSQL = "SELECT * FROM TestResults WHERE TestCaseName='" & TestName & "' And Iteration='" & citeration & "'"
			Set rs_new = DBConnection_Results.Execute(strSQL)

		'	Msgbox rs_new.RecordCount

			Dim bManualStepNo
             
			bManualStepNo = 1
			Dim res_prev_stepName
			Do While not rs_new.eof		
			
			res_stepName = rs_new.Fields.Item("StepName").Value
			res_stepDescription = rs_new.Fields.Item("StepDescription").Value
			res_stepExpectedResult = rs_new.Fields.Item("ExpectedResult").Value
			res_stepActualResult = rs_new.Fields.Item("FailureDescription").Value
			'res_stepStatus = rs_new.Fields.Item("StepResult").Value
			res_stepScrenshot = isNullisEmptyCheck(rs_new.Fields.Item("Screenshot").Value)	
			If isNullisEmptyCheck( res_stepScrenshot)="" Then
				res_stepScreenshotFlag = "No"
			else
				res_stepScreenshotFlag = "Yes"
			End If
			
			If res_prev_stepName = res_stepName Then
				res_stepName=""
				res_stepDescription = ""
				res_stepExpectedResult = ""
			End If
			
			
						HTMLResultsStreamWriter.WriteLine(_
							"<tr>" & vbLf &_
									   "<td  style=""border-style: solid; border-width: 1px;""> " &_
										   res_stepName & "</td>" & vbLf &_
										"<td style=""border-style: solid; text-align: left; border-width: 1px;""> " &_
										  res_stepDescription & "</td>" & vbLf &_
									   "<td style=""border-style: solid; text-align: left; border-width: 1px;""> " &_
										  res_stepExpectedResult & "</td>" & vbLf   &_
									  "<td style=""border-style: solid; text-align: left; border-width: 1px;""> " &_
									  res_stepActualResult & "</td>" & vbLf)
'										   res_stepStatus & "</td>" & vbLf )
'										  &_
'										"<td style=""border-style: solid; text-align: left; border-width: 1px;""> " &_
'										   res_stepStatus & "</td>" & vbLf)

						'Else
						'HTMLResultsStreamWriter.WriteLine(_
						'	"<tr>" & vbLf &_
						'			   "<td  style=""border-style: solid; border-width: 1px;""> " &_
						'					 "" & "</td>" & vbLf &_
						'				"<td style=""border-style: solid; text-align: left; border-width: 1px;""> " &_
						'				  "" & "</td>" & vbLf &_
						'			   "<td style=""border-style: solid; text-align: left; border-width: 1px;""> " &_
						'				  "" & "</td>" & vbLf)
		
	
					Dim sResult 
					screenshotValue = rs_new.Fields.Item("Screenshot").Value
					sResult = rs_new.Fields.Item("StepResult").Value
					If UCase(sResult) = "PASSED" Then
						If screenshotValue="" or isnull(screenshotValue) or isempty(screenshotValue) Then
							HTMLResultsStreamWriter.WriteLine("<td style=""color:#00D100;border-style: solid; border-width: 1px;"" > " &_
									  sResult & "</td>" & vbLf)
						Else
							HTMLResultsStreamWriter.WriteLine("<td style=""border-style: solid; border-width: 1px;"" > " &_
									   "<a style=""color:#00D100;"" href= """ & rs_new.Fields.Item("Screenshot").Value & """ target=""_blank"">" & sResult & "</a> </td>" & vbLf)
						End If
						
					ElseIf UCase(sResult) = "FAILED" Then
						HTMLResultsStreamWriter.WriteLine("<td style=""border-style: solid; border-width: 1px;"" > " &_
									   "<a style=""color:#FF0000;"" href= """ & rs_new.Fields.Item("Screenshot").Value & """target=""_blank"">" & sResult & "</a> </td>" & vbLf)
						Reporter.ReportEvent micFail, "Overall Status", "Run failed. Please check html file attached to Run"
					ElseIf UCase(sResult) = "NOT RUN" Then
						HTMLResultsStreamWriter.WriteLine("<td style=""border-style: solid;"" > " &_
									   "<a style=""color:#F6C6AD;"" href= """ & rs_new.Fields.Item("Screenshot").Value & """ target=""_blank"">" & sResult & "</a> </td>" & vbLf)
					End If
	
'					Dim failDesc
'					failDesc = rs_new.Fields.Item("FailureDescription").Value
'					If not isnull(faildesc) Then
'						If Not Cstr(failDesc) = ""  Then
'						'If Not Cstr(failDesc) = "" Then
'							failDesc = Replace(Cstr(failDesc),vbLf,"<br>")
'						else
'							failDesc = "N/A"
'						End If
'					else
'						failDesc = "N/A"
'					End If
	
				HTMLResultsStreamWriter.WriteLine("<td style=""border-style: solid; text-align: left; border-width: 1px;""> " &_
										res_stepScreenshotFlag & "</td>" & vbLf &_
						"</tr>" & vbLf)
						If  not (res_stepName="") Then
							res_prev_stepName = res_stepName
						End If
					
					rs_new.MoveNext 
			Loop
			Set rs_new = Nothing    
			HTMLResultsStreamWriter.WriteLine("</table>" & vbLf)
			rs.MoveNext 
			
	Loop
	Set rs = Nothing    

	'Finishes writing out the sub table
	HTMLResultsStreamWriter.WriteLine("<p></p>" & vbLf & "</td> </tr>")

'                If IsDBNull(dtrow("FailureDescription")) Then
'                    dtrow("FailureDescription") = "   "
'                End If
'                If dtrow("FailureDescription") = "" Then
'                    dtrow("FailureDescription") = "   "
'                End If
'Adds step results to the sub table
End Sub

Public Sub ReplaceSummary(Count_Passed,Count_Failed,Count_NotRun,Total_TC)
					On Error Resume Next

					Dim objFSO,objTextFile,strFile,strText
					Const ForReading = 1
					Const ForWriting = 2

					'Opens the execution Results.htm file
					strFile = ResultsFolder & "\ExecutionResults.htm"
   					set objFSO = CreateObject("Scripting.FileSystemObject")					
					Set objTextFile = objFSO.OpenTextFile(strFile, ForReading)
					strText = objTextFile.ReadAll

					strText = Replace(strText,"&amp;ProjectName&amp;", isNullisEmptyCheck( Environment.Value("ProjectName")))
					strText = Replace(strText,"&amp;Host&amp;",  isNullisEmptyCheck(Environment.Value("LocalHostName")))
					strText = Replace(strText,"&amp;Executed By&amp;",  isNullisEmptyCheck(Environment.Value("UserName")))
          			strText = Replace(strText,"&amp;OS&amp;", isNullisEmptyCheck(Environment.Value("OS")))
					strText = Replace(strText,"&amp;Start Time&amp;",  isNullisEmptyCheck(ExecutionStartTime))
					strText = Replace(strText,"&amp;End Time&amp;", isNullisEmptyCheck(ExecutionEndTime))
					strText = Replace(strText,"&amp;Total TC&amp;",  isNullisEmptyCheck(Total_TC))
					strText = Replace(strText,"&amp;Passed&amp;",  isNullisEmptyCheck(Count_Passed))
					strText = Replace(strText,"&amp;Failed&amp;",  isNullisEmptyCheck(Count_Failed))
					strText = Replace(strText,"&amp;Not Run&amp;",  isNullisEmptyCheck(Count_NotRun))
					If  isNullisEmptyCheck(BrowserVersion)<>"" Then
						strText = Replace(strText,"&amp;Browser Name&amp;",  isNullisEmptyCheck(BrowserVersion))
					Else
						strText = Replace(strText,"&amp;Browser Name&amp;",  isNullisEmptyCheck(Environment("ApplicationName")))
					End If
					
					objTextFile.Close
					Set objTextFile = Nothing

					Set objTextFile = objFSO.OpenTextFile(strFile, ForWriting)
					objTextFile.Write strText
					objTextFile.Close
					Set objTextFile = Nothing

					If Err.Number <> 0 Then   
	                         Reporter.ReportEvent micFail,"ReplaceSummary - Update the pass/fail summary in the html report","Failed to update the pass/fail summary - " & Err.Description
							 Err.Clear
							 Exit Sub
					End If
End Sub

Function ShowFolderList(folderspec)
   Dim fso, f, f1, fc, s
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.GetFolder(folderspec)
   Set fc = f.Files
   For Each f1 in fc
      s = s & f1.name 
      s = s & "<BR>"
   Next
   ShowFolderList = s
   Set fc = Nothing
   Set f = Nothing
   Set fso = Nothing
End Function

'Consolidated Report Generation

Sub InitializeConsolidatedReportHTML(reportPath)
On error resume next
err.clear
    reportFilePath = reportPath
    Set reportFile = CreateObject("Scripting.FileSystemObject").CreateTextFile(reportFilePath, True)

    reportFile.WriteLine "<!DOCTYPE html>"
    reportFile.WriteLine "<html lang='en'>"
    reportFile.WriteLine "<head>"
    reportFile.WriteLine "    <meta charset='UTF-8'>"
    reportFile.WriteLine "    <meta name='viewport' content='width=device-width, initial-scale=1.0'>"
    reportFile.WriteLine "    <title>Enhanced Report</title>"
    reportFile.WriteLine "    <link rel='stylesheet' href='https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css'>"
    reportFile.WriteLine "    <script src='https://cdn.jsdelivr.net/npm/chart.js'></script>"
    reportFile.WriteLine "    <script src='https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels'></script>"
    reportFile.WriteLine "    <link rel=""stylesheet"" href=""https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css"">"
    reportFile.WriteLine "    <script src=""https://ajax.googleapis.com/ajax/libs/jquery/3.7.1/jquery.min.js""></script>"
    reportFile.WriteLine "    <script src=""https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/js/bootstrap.min.js""></script>"
    
    reportFile.WriteLine "    <style>"
    reportFile.WriteLine "        body { font-family: Arial, sans-serif; margin: 0; padding: 0; display: flex; flex-direction:column; min-height:100vh; background-image : url('C:/TCOE/JDE_Auto/Results/ExecRes_3_RDCTC_1_11-20-2024%2041515%20AM/img/BaxterBackGround.png'); background-size: contain; overflow:auto;height:100vh; }"
    reportFile.WriteLine "        .sidebar { z-index: 999; background-color: #2c3e50; color: white; padding: 20px; position: fixed; height: 100%; width : 10%; min-width: 120px; top:0;bottom:0;}"
    reportFile.WriteLine "        .sidebar a { color: white; text-decoration: none; display: block; margin: 10px 0; padding: 10px; border-radius: 5px; font-size: 12px; }"
    reportFile.WriteLine "        .sidebar a:hover, .sidebar a.active { background-color: #1abc9c; }"
    reportFile.WriteLine "        .content { display: flex; flex-direction:column; min-height:100vh; margin-left: 10%; padding: 20px; flex-grow: 1; margin-top: 5% ; min-height: 100%; position: relative;}"
    reportFile.WriteLine "        .hidden { display: none; }"
    'reportFile.WriteLine "        .navbar { background-color: #2c3e50; color: white; padding: 15px; text-align: center; }"
    reportFile.WriteLine "        .test-case { padding: 0; border-bottom: 1px solid #ddd; cursor: pointer; }"
    reportFile.WriteLine "        .details { font-size: 14px; margin-top: 5px; margin-left: 1%}"
    reportFile.WriteLine "        .steps { margin-left: 20px; background: transparent;}"  ' Modify to remove hidden by default
    'reportFile.WriteLine "        .step { margin-bottom: 10px; display: flex; justify-content: space-between; }"
    reportFile.WriteLine "        h3 {font-size: 24px;}"
    reportFile.WriteLine "        html {height : 100%}"
    reportFile.WriteLine "        h4.pass {color: green; }"
    reportFile.WriteLine "        h4.fail { color: red; }"
    reportFile.WriteLine "        h4.skip { color: orange; }"
    reportFile.WriteLine "         div.col-sm-2.pass { color: green; }"
    reportFile.WriteLine "         div.col-sm-2.fail { color: red; }"
    reportFile.WriteLine "         div.col-sm-2.skip { color: orange; }"
    reportFile.WriteLine "        #tempimg { max-width: 80vw; max-height : 80vh; margin: 10px; }"
    reportFile.WriteLine "        .chart-container { position: relative; margin: 20px auto; height: 400px; width: 400px; }"
    reportFile.WriteLine "        .icon { font-size: 24px; margin-right: 15px; }"
    reportFile.WriteLine "        .sidebar a i { margin-right: 8px; }"
    reportFile.WriteLine "        td { text-align: left; vertical-align: middle;padding-left: 2px; }"
    reportFile.WriteLine "        th { text-align: left; vertical-align: middle; }"
    reportFile.WriteLine "        .test-case { z-index: 2; }"
    reportFile.WriteLine "        img.screenshot { z-index: 1; overflow : clip }"
    'reportFile.WriteLine "        button { display: flex; width : 150px; height : 150px; }"
    reportFile.WriteLine "        button.buttonClose {width : 20px; height : 10px; float : right;}"
    'reportFile.WriteLine "        button img { width: 100%; height: 100%; object-fit: cover;  }"
    reportFile.WriteLine "       table {width:100%; border-collapse:collapse; margin-top : 10px;}"
    reportFile.WriteLine "       .row-header {font-weight:bold}"
    reportFile.WriteLine "       .table {width:80vw; border-collapse:collapse; margin-top : 10px; background: transparent !important;}"
    
    reportFile.WriteLine "    /* Top Navbar Styles */"
reportFile.WriteLine "    .navbar {"
reportFile.WriteLine "        position: fixed; display:flex; align-items: center; justify-content: space-between;" ' Fix the navbar at the top of the page
reportFile.WriteLine "        top: 0;"
reportFile.WriteLine "        left: 10%;"
reportFile.WriteLine "        width: 100%; height: 12%;"
'reportFile.WriteLine "        background-color: #333;"
reportFile.WriteLine "        padding: 10px 20px;"
reportFile.WriteLine "        color: #030733; background-color :#c5cfe0; "
reportFile.WriteLine "        z-index: 999;" ' Ensure it stays on top of other content"
reportFile.WriteLine "    }"
reportFile.WriteLine "    .navbar img {"
reportFile.WriteLine "        height: 40px;" ' Adjust logo size
reportFile.WriteLine "        width: auto;"
reportFile.WriteLine "        float: Right; margin : 0% 11%; vertial-align: middle;"
reportFile.WriteLine "    }"
reportFile.WriteLine "    .navbar .test-time {"
reportFile.WriteLine "        font-size: 16px;"
reportFile.WriteLine "        font-weight: bold;"
'reportFile.WriteLine "        float: Right; margin : 0% 11%; vertial-align: middle;"
reportFile.WriteLine "    }"
reportFile.WriteLine ""
    

reportFile.WriteLine "    .dashboardDetails {"
reportFile.WriteLine "        width: 100%;"
reportFile.WriteLine "        height: 50%;"
reportFile.WriteLine "        border-collapse: collapse;"
reportFile.WriteLine "        table-layout: fixed;" ' Ensures cells have equal width
reportFile.WriteLine "    }"
reportFile.WriteLine ""
reportFile.WriteLine "    .dashboardDetails td {"
reportFile.WriteLine "        padding: 8px;"
reportFile.WriteLine "        text-align: left;"
reportFile.WriteLine "        vertical-align: middle;"
reportFile.WriteLine "        border: 1px solid #ddd;"
reportFile.WriteLine "    }"
reportFile.WriteLine ""
reportFile.WriteLine "    #chart {"
reportFile.WriteLine "        width: 100% ;"
reportFile.WriteLine "        height: 100% ;"
reportFile.WriteLine "        border: 1px solid #000;"
reportFile.WriteLine "    }"
reportFile.WriteLine ""
reportFile.WriteLine "    .dashboardDetails td:not(:last-child) {"
reportFile.WriteLine "        width: 50%;"
reportFile.WriteLine "    }"
reportFile.WriteLine "    div.logo {"
reportFile.WriteLine "      display: inline-block; float : right; margin-right : 11%;vertical-align : middle;flex:auto;"
reportFile.WriteLine "    }"
    
    
reportFile.WriteLine "  .popup {"
reportFile.WriteLine "    position: fixed;"
reportFile.WriteLine "    top: 50%;"
reportFile.WriteLine "    left: 50%;"
reportFile.WriteLine "    transform: translate(-50%, -50%);"
reportFile.WriteLine "    background-color: white;"
reportFile.WriteLine "    padding: 20px;"
reportFile.WriteLine "    box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);"
reportFile.WriteLine "    border-radius: 8px;"
reportFile.WriteLine "    z-index: 1000;"
reportFile.WriteLine "  }"
reportFile.WriteLine "  .popup-overlay {"
reportFile.WriteLine "    position: fixed;"
reportFile.WriteLine "    top: 10%;"
reportFile.WriteLine "    left: 10%;"
reportFile.WriteLine "    width: 80vw;"
reportFile.WriteLine "    height: 80vh;"
reportFile.WriteLine "    background-color: rgba(0, 0, 0, 0.5);"
reportFile.WriteLine "    z-index: 999;"
reportFile.WriteLine "    display: none;"
reportFile.WriteLine "  }"
    
    reportFile.WriteLine "    </style>"
    reportFile.WriteLine "<script>"
    
    
    reportFile.WriteLine "    window.onload = function() {"
    

    
    reportFile.WriteLine "        // Get all elements with the class containing 'test-case'"

    reportFile.WriteLine "        var footerElement = document.getElementById('footer');"
    reportFile.WriteLine "        var bodyElement = document.getElementById('body');"
    reportFile.WriteLine "        if(footerElement!=null){footerElement.remove();};"
    'reportFile.WriteLine "        bodyElement.appendChild(footerElement);"
    reportFile.WriteLine "        var testCases = document.querySelectorAll('.test-case');"
    reportFile.WriteLine "        "
    reportFile.WriteLine "        // Get the dashboard view container"
    reportFile.WriteLine "        var testView = document.getElementById('testView');"
    reportFile.WriteLine "        const elements = document.querySelectorAll('[id^=popupOverlay]');"
    reportFile.WriteLine "        elements.forEach(function(element) { element.remove(); });"
    
    reportFile.WriteLine "        const popupOverlay = document.createElement('div');"
    reportFile.WriteLine "        popupOverlay.id='popupOverlay' "
    reportFile.WriteLine "       popupOverlay.class='popupOverlay' "

    reportFile.WriteLine "        const popupDiv = document.createElement('div');popupDiv.id = 'screenshotPopup'; popupDiv.className = 'popup';popupDiv.style.display = 'none';const closeButton = document.createElement('button');closeButton.class = 'buttonClose';closeButton.textContent = 'Close';closeButton.onclick = closePopup;popupDiv.appendChild(closeButton);"
    
    reportFile.WriteLine "        popupOverlay.appendChild(popupDiv);"
    reportFile.WriteLine "        testView.appendChild(popupOverlay);"
    reportFile.WriteLine "        // Loop through all test cases and move them to the dashboard view"
    reportFile.WriteLine "        testCases.forEach(function(testCase) {"
    reportFile.WriteLine "            // Remove the test case from its current position in the DOM"
    reportFile.WriteLine "            testCase.remove();"
    reportFile.WriteLine "            "
    reportFile.WriteLine "            // Append it to the dashboard view"
    reportFile.WriteLine "            testView.appendChild(testCase);"
    reportFile.WriteLine "        });    "
    reportFile.WriteLine "    };"
    
    reportFile.WriteLine "function toggleTestSteps(testCaseDivID) {"
    reportFile.WriteLine "    var stepsDivList = document.querySelectorAll('#'+testCaseDivID);"
    reportFile.WriteLine "        stepsDivList.forEach(function(stepsDiv) {"
    reportFile.WriteLine "           stepsDiv.style.border = 'collapse'; "
    reportFile.WriteLine "    console.log('Toggling test steps for: ' + testCaseDivID);"
    reportFile.WriteLine "    if (stepsDiv.style.display === 'none' || stepsDiv.style.display === '') {"
    reportFile.WriteLine "        stepsDiv.style.display = 'block';"
    reportFile.WriteLine "    } else {"
    reportFile.WriteLine "        stepsDiv.style.display = 'none';"
    reportFile.WriteLine "    }"
    reportFile.WriteLine "        });"

    reportFile.WriteLine "}"

    reportFile.WriteLine "function changeActiveLink(LinkID) {"
    reportFile.WriteLine "    document.getElementById('dashboardLink').classList.remove('active');"
    reportFile.WriteLine "    document.getElementById('testLink').classList.remove('active');"
    reportFile.WriteLine "    document.getElementById('exceptionLink').classList.remove('active');"
    reportFile.WriteLine "    document.getElementById(LinkID).classList.add('active');"
    reportFile.WriteLine "};"
    reportFile.WriteLine "function showView(viewID) {"
    reportFile.WriteLine "    document.getElementById('dashboardView').classList.add('hidden');"
    reportFile.WriteLine "    document.getElementById('testView').classList.add('hidden');"
    reportFile.WriteLine "    document.getElementById('exceptionView').classList.add('hidden');"
    reportFile.WriteLine "    document.getElementById(viewID).classList.remove('hidden');"
    
    reportFile.WriteLine "    if(viewID==='exceptionView'){"
    reportFile.WriteLine "        // Get all elements with the class containing 'test-case failed'"

    reportFile.WriteLine "        "
    reportFile.WriteLine "        // Get the dashboard view container"
    reportFile.WriteLine "        var exceptionView = document.getElementById('exceptionView');"
    reportFile.WriteLine "        const elements = document.querySelectorAll('[id^=popupOverlay]');"
    reportFile.WriteLine "        elements.forEach(function(element) { element.remove(); });"
    
    reportFile.WriteLine "        const popupOverlay = document.createElement('div');"
    reportFile.WriteLine "        popupOverlay.id='popupOverlay' "
    reportFile.WriteLine "       popupOverlay.class='popupOverlay' "

    reportFile.WriteLine "        const popupDiv = document.createElement('div');popupDiv.id = 'screenshotPopup'; popupDiv.className = 'popup';popupDiv.style.display = 'none';const closeButton = document.createElement('button');closeButton.textContent = 'Close';closeButton.class = 'buttonClose';closeButton.onclick = closePopup;popupDiv.appendChild(closeButton);"
    
    reportFile.WriteLine "        popupOverlay.appendChild(popupDiv);"
    reportFile.WriteLine "        exceptionView.appendChild(popupOverlay);"
    

     reportFile.WriteLine "        const classlistExceptionView = exceptionView.classList.contains('alreadyNavigated')"
     
    reportFile.WriteLine "        if(!classlistExceptionView){"
    reportFile.WriteLine "        var testCasesFailed = document.querySelectorAll('.test-case.fail');"
    reportFile.WriteLine "        var testCasesSkipped = document.querySelectorAll('.test-case.skip');"
    reportFile.WriteLine "        // Loop through all test cases and move them to the dashboard view"
    reportFile.WriteLine "        testCasesFailed.forEach(function(testCaseFailed) {"
    reportFile.WriteLine "           const testCaseFailedTemp = testCaseFailed.cloneNode(true); "
    reportFile.WriteLine "            // Append it to the exception view"
    reportFile.WriteLine "            exceptionView.appendChild(testCaseFailedTemp);"
    reportFile.WriteLine "        });"
    reportFile.WriteLine "        testCasesSkipped.forEach(function(testCaseSkipped) {"
    reportFile.WriteLine "            const testCaseSkippedTemp = testCaseSkipped.cloneNode(true); "
    reportFile.WriteLine "            // Append it to the exception view"
    reportFile.WriteLine "            exceptionView.appendChild(testCaseSkippedTemp);"
    reportFile.WriteLine "        });"
    reportFile.WriteLine "       exceptionView.classList.add('alreadyNavigated');"
    reportFile.WriteLine "       }"
    reportFile.WriteLine "      }"
    
    reportFile.WriteLine "    if(viewID==='testView'){"
    reportFile.WriteLine "        // Get all elements with the class containing 'test-case'"
    reportFile.WriteLine "        var testCases = document.querySelectorAll('.test-case');"
    reportFile.WriteLine "        "
    reportFile.WriteLine "        // Get the dashboard view container"
    reportFile.WriteLine "        var testView = document.getElementById('testView');"
    
    reportFile.WriteLine "        const elements = document.querySelectorAll('[id^=popupOverlay]');"
    reportFile.WriteLine "        elements.forEach(function(element) { element.remove(); });"
    
    reportFile.WriteLine "        const popupOverlay = document.createElement('div');"
    reportFile.WriteLine "        popupOverlay.id='popupOverlay' "
    reportFile.WriteLine "       popupOverlay.class='popupOverlay' "

    reportFile.WriteLine "        const popupDiv = document.createElement('div');popupDiv.id = 'screenshotPopup'; popupDiv.className = 'popup';popupDiv.style.display = 'none';const closeButton = document.createElement('button');closeButton.textContent = 'Close';closeButton.onclick = closePopup;popupDiv.appendChild(closeButton);closeButton.class = 'buttonClose';"
    
    reportFile.WriteLine "        popupOverlay.appendChild(popupDiv);"
    reportFile.WriteLine "        testView.appendChild(popupOverlay);"
    
    
    
    reportFile.WriteLine "       }"
    
    
    reportFile.WriteLine "}"
    
    reportFile.WriteLine "function showPopup(imageSrc) {"
    reportFile.WriteLine "     try{closePopup() ;}catch(err){}; var imgElement = document.createElement('img'); "
    reportFile.WriteLine "    imgElement.id = 'tempimg'; "
    reportFile.WriteLine "    imgElement.src = imageSrc; "
    reportFile.WriteLine "    document.getElementById('screenshotPopup').appendChild(imgElement)"
    reportFile.WriteLine "    document.getElementById('popupOverlay').style.display = 'block';"
    reportFile.WriteLine "    document.getElementById('screenshotPopup').style.display = 'block'; checkAndDisableBackGround();"
    reportFile.WriteLine "		var buttons = document.querySelectorAll('button');"
    reportFile.WriteLine " 		buttons.forEach(function(element) {"
    reportFile.WriteLine "		element.addEventListener('click', function(event){ event.stopPropagation(); }); });"
    reportFile.WriteLine "}"
    reportFile.WriteLine "function closePopup() {"
    reportFile.WriteLine "    document.getElementById('popupOverlay').style.display = 'none';"
    reportFile.WriteLine "    document.getElementById('screenshotPopup').style.display = 'none';"
    reportFile.WriteLine "    document.getElementById('tempimg').remove(); checkAndDisableBackGround();"
    reportFile.WriteLine "}"
    
' 	reportFile.WriteLine "document.addEventListener('click',  function handleClickOutsideBox(event) {"
'     reportFile.WriteLine "const box = document.getElementById('screenshotPopup');"
'    reportFile.WriteLine "if (!box.contains(event.target)) {"
'      reportFile.WriteLine " closePopup();  } }, );"

 reportFile.WriteLine " function checkAndDisableBackGround(){"
 reportFile.WriteLine " 	var screenshotElement = document.getElementById('screenshotPopup');"
 reportFile.WriteLine " 	if(screenshotElement && screenshotElement.style.display!='none'){"
 reportFile.WriteLine " 		document.body.style.pointerEvents=""none"";"
 reportFile.WriteLine " 		screenshotElement.style.pointerEvents=""auto"";"
 reportFile.WriteLine " 	}"
 reportFile.WriteLine " 	else{"
 reportFile.WriteLine " 		document.body.style.pointerEvents=""auto"";"
 reportFile.WriteLine " 		screenshotElement.style.pointerEvents=""auto"";"
 reportFile.WriteLine " 	}"
 reportFile.WriteLine " }"



    reportFile.WriteLine "</script>"

    reportFile.WriteLine "</head>"
    reportFile.WriteLine "<body id='body'>"
    reportFile.WriteLine "<div class='navbar'>"
    reportFile.WriteLine "      <table class=""reportDetails"" style="" width:30% ; font-size:12px; line-height : normal; display: inline-block; flex:none"">"
    reportFile.WriteLine "      <tr>"
    reportFile.WriteLine "      <td>Project</td>"
    reportFile.WriteLine "      <td>:&nbsp</td>"
    reportFile.WriteLine "      <td>"& isNullisEmptyCheck( Environment("ProjectName") )&"</td>"
    reportFile.WriteLine "      </tr>"
    reportFile.WriteLine "      <tr>"
    reportFile.WriteLine "      <td>Application Name</td>"
    reportFile.WriteLine "      <td>:</td>"
    reportFile.WriteLine "      <td>" & isNullisEmptyCheck( Environment("ApplicationName") )& "</td>"
    reportFile.WriteLine "      </tr>"
    reportFile.WriteLine "      <tr>"
    reportFile.WriteLine "      <td>Version</td>"
    reportFile.WriteLine "      <td>:</td>"
    reportFile.WriteLine "      <td>" & isNullisEmptyCheck( Environment("Version")) & "</td>"
    reportFile.WriteLine "      </tr>"
    reportFile.WriteLine "      <tr>"
    reportFile.WriteLine "      <td>Release#</td>"
    reportFile.WriteLine "      <td>:</td>"
    reportFile.WriteLine "      <td>" & isNullisEmptyCheck( Environment("Release")) & "</td>"
    reportFile.WriteLine "      </tr>"
    reportFile.WriteLine "      </table>"
    reportFile.WriteLine "    <div class='logo'><img src="".\img\transparentLogo.png"" alt='Logo'>"
    reportFile.WriteLine "    </div>"
    reportFile.WriteLine "</div>"
    reportFile.WriteLine "    <div class='sidebar'>"
    reportFile.WriteLine "        <a href=""#"" class=""active"" id=""dashboardLink"" onclick=""showView('dashboardView');changeActiveLink('dashboardLink');""><i class='fas fa-tachometer-alt icon'></i>Dashboard</a>"
    reportFile.WriteLine "        <a href=""#"" class="""" id=""testLink"" onclick=""showView('testView');changeActiveLink('testLink');""><i class='fas fa-flask icon'></i>TestResult</a>"
    reportFile.WriteLine "        <a href=""#"" class="""" id=""exceptionLink"" onclick=""showView('exceptionView');changeActiveLink('exceptionLink');""><i class ='fas fa-exclamation-circle icon'></i>Failures</a>"
    reportFile.WriteLine "    </div>"
    reportFile.WriteLine "    <div class='content'>"
    reportFile.WriteLine "        <div id='dashboardView'>"
    reportFile.WriteLine "            <h3>Test Summary</br></br></h3>"
    reportFile.WriteLine "      <div style='width: 100%; height: 50%; display: flex; justify-content: center;'>"
    reportFile.WriteLine "    <table  class='dashboardDetails'>"
    reportFile.WriteLine "    <tr>  <td>Host</td>                        <td>"& Environment.Value("LocalHostName")  &"</td>        <td style=""border:None""></td> <td style=""border:None""></td>  <td rowspan='20' style='width:100%; height: 100%;'> <canvas id='chart'></canvas>    </td>   </tr>"
    reportFile.WriteLine "    <tr>  <td>Executed By</td>                 <td>"& Environment.Value("UserName") &"</td> <td style=""border:None""></td> <td style=""border:None""></td>  <td></td></tr>"
    reportFile.WriteLine "    <tr>  <td>OS</td>                          <td>" & Environment.Value("OS")  & "</td>          <td style=""border:None""></td> <td style=""border:None""></td>  <td></td></tr>"
    consolidatedReportStartTime = FormatDateTime(Date) & " " & FormatDateTime(Time)    
   reportFile.WriteLine "    <tr>  <td>Start Time</td>                  <td>"& consolidatedReportStartTime &"</td>  <td style=""border:None""></td> <td style=""border:None""></td>  <td></td></tr>"
    reportFile.WriteLine "    <tr>  <td>End Time</td>                    <td>&amp;End Time&amp;</td>    <td style=""border:None""></td> <td style=""border:None""></td>  <td></td></tr>"
    'reportFile.WriteLine "    <tr><td>Application Name :</td><td>&Application Name&</td><td></td></tr>"
    reportFile.WriteLine "    <tr>  <td>Total Test Cases Executed</td>   <td>&amp;Executed&amp;</td>    <td style=""border:None""></td> <td style=""border:None""></td>  <td></td></tr>"
    reportFile.WriteLine "    <tr>  <td>Total Test Cases Passed</td>     <td>&amp;Passed&amp;</td>      <td style=""border:None""></td> <td style=""border:None""></td>  <td></td></tr>"
    reportFile.WriteLine "    <tr>  <td>Total Test Cases Failed</td>     <td>&amp;Failed&amp;</td>      <td style=""border:None""></td> <td style=""border:None""></td>  <td></td></tr>"
    reportFile.WriteLine "    <tr>  <td>Total Test Cases Skipped</td>    <td>&amp;Skipped&amp;</td>     <td style=""border:None""></td> <td style=""border:None""></td>  <td></td></tr>"
    
    reportFile.WriteLine "    "
    reportFile.WriteLine "    </table>"
    reportFile.WriteLine "</div>"
    reportFile.WriteLine "        </div>"
    reportFile.WriteLine "        <div id='testView' class='hidden'>"
    reportFile.WriteLine "          "
    reportFile.WriteLine "            <h3>Test Cases</br></br></h3>"
    reportFile.WriteLine "        </div>"
    reportFile.WriteLine "        <div id='exceptionView' class='hidden'>"
    reportFile.WriteLine "            <h3>Failures/Skipped</br></br></h3>"
    reportFile.WriteLine "        </div>"
    reportFile.WriteLine "    </div>"
    
    If err.number<>0 Then
    	Reporter.ReportEvent micFail, "initialize consolidated report", "failed to initialize consolidated report " & err.description
    	Exit Sub
    End If
    
    End Sub
    
    Sub consolidatedReportFinilizer(totalConsolidatedCases, totalConsolidatedPassed, totalConsolidatedFailed, totalConsolidatedSkipped)
    err.clear
    On error resume next
    ' Append closing tags for the HTML structure
    reportFile.WriteLine "    </div>" ' Close content div
        ' Add chart.js initialization (if you want to visualize test data)
reportFile.WriteLine "<script>"
reportFile.WriteLine "var ctx = document.getElementById('chart').getContext('2d');"
reportFile.WriteLine "var myChart = new Chart(ctx, {"
reportFile.WriteLine "    type: 'pie',"
reportFile.WriteLine "    data: {"
reportFile.WriteLine "        labels: ['Passed', 'Failed', 'Skipped'],"
reportFile.WriteLine "        datasets: [{"
reportFile.WriteLine "            label: 'Test Results',"
reportFile.WriteLine "            data: [" & totalConsolidatedPassed & ", " & totalConsolidatedFailed & ", " & totalConsolidatedSkipped & "],"
reportFile.WriteLine "            backgroundColor: ['#4caf50', '#f44336', '#ffeb3b'],"
reportFile.WriteLine "        }]"
reportFile.WriteLine "    },"
reportFile.WriteLine "    options: {"
reportFile.WriteLine "        responsive: true,"
reportFile.WriteLine "        plugins: {"
reportFile.WriteLine "            legend: {"
reportFile.WriteLine "                position: 'top'"
reportFile.WriteLine "            },"
reportFile.WriteLine "            datalabels: {"
reportFile.WriteLine "                formatter: function(value, context) {"
reportFile.WriteLine "                    var total = context.dataset.data.reduce((a, b) => a + b, 0);"
reportFile.WriteLine "                    var percentage = ((value / total) * 100).toFixed(1) + '%';"
reportFile.WriteLine "                    return percentage;"
reportFile.WriteLine "                },"
reportFile.WriteLine "                color: '#fff',"
reportFile.WriteLine "                font: {"
reportFile.WriteLine "                    weight: 'bold'"
reportFile.WriteLine "                }"
reportFile.WriteLine "            }"
reportFile.WriteLine "        }"
reportFile.WriteLine "    },"
reportFile.WriteLine "    plugins: [ChartDataLabels]"
reportFile.WriteLine "});"
reportFile.WriteLine "</script>"

''reportFile.WriteLine "<footer id='footer' style='display: flex;justify-content: flex-end; align-items: right; bottom:0;top:95%;padding: 1% 5%; border-top: 1px solid #ccc; position: relative; width: 100vw; margin-top: auto;'>"
'reportFile.WriteLine "  <img src='C:\\output\\logo.jpeg' style='object-fit: contain; width: 180px; height: 50px; margin-left: 15px; vertical-align: middle;' />"
'reportFile.WriteLine "  <span style='float: right;font-size: 14px; color: #555;filter: opacity(0.5);'>Baxter Confidential - Do not distribute without prior approval</span>"
'reportFile.WriteLine "</footer>"
    reportFile.WriteLine "</body>"
    reportFile.WriteLine "</html>"

    ' Close the file
    reportFile.Close
    
    If err.number<>0 Then
    	Reporter.ReportEvent micFail, "close reporter","unable to write the close statements for consolidated report" & err.description
    	Exit Sub
    End If
    
End Sub

Public Function ReplaceConsolidatedData()
	err.clear
	On error resume next
	consolidatedReportEndTime = FormatDateTime(Date) & " " & FormatDateTime(Time) 
	set objFSO = CreateObject("Scripting.FileSystemObject")	
	strFile = Environment("CONSOLIDATED_RESULTS_HTML")	
		Set objTextFile = objFSO.OpenTextFile(strFile, 1)
		strText = objTextFile.ReadAll
		strText = Replace(strText,"&amp;End Time&amp;", consolidatedReportEndTime)
		strText = Replace(strText,"&amp;Executed&amp;", totalConsolidatedTestCases)
		strText = Replace(strText,"&amp;Passed&amp;", totalConsolidatedPassed)
		strText = Replace(strText,"&amp;Failed&amp;", totalConsolidatedFailed)
		strText = Replace(strText,"&amp;Skipped&amp;", totalConsolidatedSkipped)
		objTextFile.Close
		Set objTextFile = Nothing
		Set objTextFile = objFSO.OpenTextFile(strFile, 2)
		objTextFile.Write strText
		objTextFile.Close
		Set objTextFile = Nothing

		If Err.Number <> 0 Then   
	             Reporter.ReportEvent micFail,"ReplaceSummary - Update the pass/fail summary in the consolidated html report","Failed to update the pass/fail summary  in consolidated report- " & Err.Description
		Err.Clear
		Exit Function
		End If
	Set objFSO = Nothing
	If err.number<>0 Then
		Reporter.ReportEvent micFail,"Replace value in consolidated report", "failed to replace value in consolidated report" & err.description
	End If
End Function

Public Function AddTestsToConsolidatedHTMLReport()
    
    Err.Clear
    On Error Resume Next
    
    Dim strSQL, APIResultSet, strFile, objFSO
    TCResultSet = "SELECT * FROM CompleteResults"
    Set TCResultSet = DBConnection_Results.Execute(TCResultSet)
    Do While Not TCResultSet.EOF
    		totalConsolidatedTestCases = totalConsolidatedTestCases+1
		TestCaseName = isNullisEmptyCheck(TCResultSet.Fields.Item("TestCase").Value)
		TestSetID = isNullisEmptyCheck(TCResultSet.Fields.Item("TestSetid").Value)
		TestDescription = isNullisEmptyCheck(TCResultSet.Fields.Item("TestDescription").Value)
		TestResult = isNullisEmptyCheck(TCResultSet.Fields.Item("TestResult").Value)
		TestDuration = isNullisEmptyCheck(TCResultSet.Fields.Item("TestDuration").Value)
		TestStart = ExecutionStartTime
		TestEnd = ExecutionEndTime
		
		'Add test steps to test case details
		Set TestStepResultListDict = CreateObject("Scripting.Dictionary")
		TCStepNo = 0
		
		'fetch unique test steps and details and add them to TestStepResultListDict
		TSStepUniqueResultSetQuery = "SELECT Val(SNo) as SerialNo, TestResults.[StepName] AS StepNames FROM TestResults;"
		Set TSStepUniqueResultSet = DBConnection_Results.Execute(TSStepUniqueResultSetQuery)
		Set StepNosList = CreateObject("System.Collections.ArrayList")
		Do While Not TSStepUniqueResultSet.EOF
			stepNamesVal = isNullisEmptyCheck(TSStepUniqueResultSet.Fields.Item("StepNames").Value)
			If Not StepNosList.Contains(stepNamesVal) Then
				StepNosList.Add(stepNamesVal)
			End If
		TSStepUniqueResultSet.MoveNext
		Loop
		For Iterator = 0 To StepNosList.Count-1 Step 1
			Set TestStepDict = CreateObject("Scripting.Dictionary")
			strStepName = StepNosList(Iterator)
			'"Step " & isNullisEmptyCheck(TSStepUniqueResultSet.Fields.Item("StepNames").Value)
			Print strStepName
			TSResultSet = "SELECT * FROM TestResults where StepName='" & strStepName & "'"
			Set TestStepResultSet = DBConnection_Results.Execute(TSResultSet)
			TestStepTimeStamp = ""
			TestStepStatus = ""
			TestStepError = ""
			TestStepDescription = ""
			Set TestStepScreenShot = CreateObject("System.Collections.ArrayList")
			ActualTestStepStartTime =""' isNullisEmptyCheck(TestStepResultSet.Fields.Item("TestStepStartTime").Value)
			TestStepStartTime = ActualTestStepStartTime
			ActualTestStepDescription = isNullisEmptyCheck(TestStepResultSet.Fields.Item("StepDescription").Value)
			TestStepDescription = ActualTestStepDescription
			Do While Not TestStepResultSet.EOF
				ActualTestStepStatus = isNullisEmptyCheck(TestStepResultSet.Fields.Item("StepResult").Value)
				If Trim(LCase(ActualTestStepStatus)) = "pass" Or Trim(LCase(ActualTestStepStatus)) = "passed" Then
					ActualTestStepStatus = "pass"
				ElseIf Trim(LCase(ActualTestStepStatus)) = "fail" Or Trim(LCase(ActualTestStepStatus)) = "failed" Then
					ActualTestStepStatus = "fail"
				ElseIf Trim(LCase(ActualTestStepStatus)) = "skip" Or Trim(LCase(ActualTestStepStatus)) = "skipped" Then
					ActualTestStepStatus = "skip"
				End If
				TestStepStatus = GetTestStepStatus(TestStepStatus, ActualTestStepStatus)
				
				If ActualTestStepStatus = "fail" Or ActualTestStepStatus = "skip" Then
					ActualTestStepError = Trim(isNullisEmptyCheck(TestStepResultSet.Fields.Item("FailureDescription").Value))
				Else
					ActualTestStepError = ""
				End If
				If not ActualTestStepError = "" Then
					If not TestStepError = "" Then
						TestStepError = TestStepError & vbCrLf & ActualTestStepError
					Else
						TestStepError = ActualTestStepError
					End If
				End If
				ActualTestStepScreenShot = isNullisEmptyCheck(TestStepResultSet.Fields.Item("Screenshot").Value)
				If Not ActualTestStepScreenShot = "" Then
                    			TestStepScreenShot.Add (ActualTestStepScreenShot)
                		End If
				TestStepResultSet.MoveNext
			Loop
			TestStepDict.Add "stepStatus" , TestStepStatus
			TestStepDict.Add "stepTimeStamp" , TestStepStartTime
			TestStepDict.Add "stepDescription" , TestStepDescription
			TestStepDict.Add "stepError" , TestStepError
			TestStepDict.Add "stepScreenShot" , TestStepScreenShot
			TCStepNo = TCStepNo+1
			TestStepResultListDict.Add TCStepNo, TestStepDict
			Set TestStepScreenShot = Nothing
			'TSStepUniqueResultSet.MoveNext
		Next
		TCResultSet.MoveNext
	Loop
	Call AddTestCaseToConsolidatedHTMLReport(TestResult,TestCaseName,TestSetID,TestDescription,TestDuration,TestStart,TestEnd, TestStepResultListDict )
	 If err.number<>0 Then
    	Reporter.ReportEvent micFail, "Add tests to consolidated report","unable to add tests to consolidated report - " & err.description
    End If
End Function

Public Function AddTestCaseToConsolidatedHTMLReport(TestResult,TestCaseName,TestSetID,TestDescription,TestDuration,TestStartTime,TestEndTime, TestStepResultListDict)
	err.clear
	On error resume next
	' Assign appropriate color classes for status
    If UCase(TestResult) = "PASS" or UCase(TestResult) = "PASSED" Then
        statusClass = "pass"
        totalConsolidatedPassed= totalConsolidatedPassed+1
    ElseIf UCase(TestResult) = "FAIL" or UCase(TestResult) = "FAILED" Then
        statusClass = "fail"
        totalConsolidatedFailed = totalConsolidatedFailed+1
    Else
        statusClass = "skip"
        totalConsolidatedSkipped = totalConsolidatedSkipped+1
    End If
	' Generate a unique div ID for each test case for easy referencing in JS
	testCaseNo = TestCaseName & TestSetID
    testCaseDivID = "testCase" & testCaseNo
    If CInt(TestDuration)>60 Then
    	TestDuration = FormatNumber((CInt(TestDuration)/60),0) & " mins " & CInt(TestDuration) Mod 60 & "secs"
    End If
    reportFile.WriteLine "                <div class=""test-case " & statusClass & """ onclick=""toggleTestSteps('" & testCaseDivID & "')"">"
    reportFile.WriteLine "                 <h4 class='" & statusClass & "'>" & testCaseNo & " - " & TestDescription & " ( " & statusClass & " )</h3>"
    reportFile.WriteLine "                    <div class='details'> <div class='table'><div class='row'><div class='col-sm-3'>Test ID:  " & testCaseNo & "</div><div class='col-sm-3'> Start:  " & TestStartTime & "</div><div class='col-sm-3'> End:  " & TestEndTime & "</div><div class='col-sm-3'> Duration:  " & TestDuration & "</div>"
    If Not isNullisEmptyCheck(BrowserVersion) = "" Then
    	reportFile.WriteLine "     <div class='col-sm-4'>Tested on Broswer: " & BrowserVersion & "</div>"
    Else
    	reportFile.WriteLine "     <div class='col-sm-4'>Tested on Application:" & isNullisEmptyCheck(Environment("ApplicationName") ) & " " & isNullisEmptyCheck(Environment("Version") ) & "</div>"
    End If
    reportFile.WriteLine "</div><div>"

	'Add Steps
	reportFile.WriteLine "                    <div class=""steps"" id=""" & testCaseDivID & """ style=""display: None; border: 1px solid blue;"">"
    reportFile.WriteLine "                        <div class='table' >"
    reportFile.WriteLine "                            <div class='row row-header'>"
    reportFile.WriteLine "                                <div class='col-sm-2'>Status</div><div class='col-sm-2'>Timestamp</div><div class='col-sm-3'>Test Step Expectation</div><div class='col-sm-2'>Error</div><div class='col-sm-2'>Screenshot</div>"
   
    
    reportFile.WriteLine "                            </div>"
    
    keys = TestStepResultListDict.Keys
    For i = 0 To TestStepResultListDict.Count - 1
    	 Set TestStepResultDict = TestStepResultListDict(keys(i))
    	' Color code steps based on their status
    	stepStatus=""
    	stepTimeStamp=""
    	stepDescription=""
    	stepError=""
    	Set stepScreenShot= Nothing
    	stepStatus = TestStepResultDict("stepStatus")
    	stepTimeStamp = TestStepResultDict("stepTimeStamp")
    	stepDescription = TestStepResultDict("stepDescription")
    	stepError = TestStepResultDict("stepError")
    	Set stepScreenShot = TestStepResultDict("stepScreenShot")
    	
    	
    	
        If UCase(stepStatus) = "PASS"  or UCase(stepStatus) = "PASSED" Then
            stepStatusClass = "pass"
        ElseIf UCase(stepStatus) = "FAIL" or UCase(stepStatus) = "FAILED" Then
            stepStatusClass = "fail"
        Else
            stepStatusClass = "skip"
        End If
    	
    	
    	' Add screenshot if available
		screenshotTag=""
        If (not IsNull(stepScreenShot)) and stepScreenShot.Count <> 0 Then
			for nCnt = 0 To stepScreenShot.Count-1
				strStepScreenShot = stepScreenShot.Item(CInt(nCnt) )
				strScreenshotTag = "<button onclick=""showPopup('" & Replace(Replace(strStepScreenShot, "\", "\\"), "'", "\'") & "')"" ><img class=""screenshot"" id=""popupImage"" src='" & strStepScreenShot & "' alt='Screenshot' style = ""height : 200px; width : 200px; object-fit: contain; display: block;""></button> "
				screenshotTag=screenshotTag & vbCrLf & strScreenshotTag
			Next
        Else
            screenshotTag = ""
        End If
        
            ' Write step details
        reportFile.WriteLine "                            <div class='row'>"
        reportFile.WriteLine "                                <div class='col-sm-2 " & stepStatusClass & "'>" & stepStatus & "</div>"
        reportFile.WriteLine "                                <div class='col-sm-2'>" & isNullisEmptyCheck(stepTimeStamp) & "</div>"
        reportFile.WriteLine "                                <div class='col-sm-3'>" & isNullisEmptyCheck(stepDescription) & "</div>"
        reportFile.WriteLine "                                <div class='col-sm-2'>" & isNullisEmptyCheck(stepError) & "</div>"
        reportFile.WriteLine "                                <div class='col-sm-2'>" & isNullisEmptyCheck(screenshotTag) & "</div>"
        reportFile.WriteLine "                            </div>"
    	
    Next
    	reportFile.WriteLine "                       </div>"
    	reportFile.WriteLine "                     </div>"
    	reportFile.WriteLine "                   </div>"
    	reportFile.WriteLine "                  </div>"
    	
    	If err.number<>0 Then
    	Reporter.ReportEvent micFail, "Add test case and steps to consolidated report","unable to add testcase and steps to consolidated report - " & err.description
    End If
    	
End Function

Public Function GetTestStepStatus(testStepStatus,ActualTestStepStatus)
	'-----------------------------------------------------------------------------------------
				If (TestStepStatus = "") and ActualTestStepStatus = "pass" Then
					TestStepStatus = "pass"
				ElseIf TestStepStatus="" and ActualTestStepStatus="skip" Then
					TestStepStatus = "skip"
				
				ElseIf TestStepStatus="" and ActualTestStepStatus="fail" Then
					TestStepStatus = "fail"
					'-------------------------------------------------------------------------------------
				ElseIf TestStepStatus="pass" and ActualTestStepStatus="fail" Then
					TestStepStatus = "fail"
				ElseIf TestStepStatus="pass" and ActualTestStepStatus="skip" Then
					TestStepStatus = "skip"
				
				ElseIf TestStepStatus="pass" and ActualTestStepStatus="pass" Then
					TestStepStatus = "pass"
				
				'-----------------------------------------------------------------------------------------
				ElseIf TestStepStatus="fail" and ActualTestStepStatus="fail" Then
					TestStepStatus = "fail"
				
				ElseIf TestStepStatus="fail" and ActualTestStepStatus="skip" Then
					TestStepStatus = "fail"
				
				ElseIf TestStepStatus="fail" and ActualTestStepStatus="pass" Then
					TestStepStatus = "fail"
				
				'-----------------------------------------------------------------------------------------
				ElseIf TestStepStatus="skip" and ActualTestStepStatus="pass" Then
					TestStepStatus = "skip"
				
				ElseIf TestStepStatus="skip" and ActualTestStepStatus="skip" Then
					TestStepStatus = "skip"
				
				ElseIf TestStepStatus="skip" and ActualTestStepStatus="fail" Then
					TestStepStatus = "fail"
				End If
			GetTestStepStatus = TestStepStatus
End Function
