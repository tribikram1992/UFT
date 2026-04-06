Public pdfFileExist,tsrfileDownload,txtfileDownload,InitializeAPIReportCount
Public sRow,tRow
Public resExcel,resWorkbook,resStepsWorksheet,resTestWorksheet
Dim strIteration
Public adOpenStatic,adLockOptimistic ,adCmdText, adOpenDynamic
Public TestEndTime, TestStartTime
adOpenStatic = 3
adLockOptimistic = 3
adCmdText = &H0001
adOpenDynamic = 2
Public TestResult
Public DBConnection_DriverFile
Dim arrIteration
Public Const HEADER_ROW = 10
Public strParameters,strObjectName,strScreenName,stepExpected,stepActual,StepResult,stepDuration,stepErrDescription,stepScreenshot,stepScreenshot1,stepDescription,mstepResult, actualResult,actualResult1,defaultstepExpected,lastModifiedTime,StepStartTime
Public blnExitIteration, blnExitTestCase,blnSetProperty,strExitTest,blnQCPrintVal
Public CONST componentWiseData = False
Public Const conTwo = 2, conThree = 3, conFive = 5, conTen = 10, conTwenty = 20, conFifteen = 15, conThirty = 30, conSixty = 60
Public Const conExist = 20
Public strBrowserTitle
strBrowserTitle = ".*"
Public inputparameters
Public outputparameters
Public valTestCaseName,valslNo, valStepName, valStepDescription, valExpectedResult, valActualResult,valProvingorNonProving, valAutomationStepDescription, valScreenName
Public valObject, valAction, valParam1, valParam2, valParam3, valParam4, valParam5, valErrorHandler, valScreenShot
Public  valPreviousStepName
Public valPreviousStepDescription
Public valPreviousStepExpectedResult
Public valPreviousStepProvingorNonProving
Public strAction,strErrorHandler
Public qtTest


Public Function ExecuteTestCase(ExcelFile, iterationCnt)
	On Error Resume Next
	Set inputparameters = CreateObject("Scripting.Dictionary")
	Set outputparameters = CreateObject("Scripting.Dictionary")
	API_Tests_Executed = False
	Set qtApp = CreateObject("QuickTest.Application")
	Set qtTest = qtApp.Test
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	downloadDriverFile()
	strDriverFile = Environment("DRIVER_FILE")
	If Not objFSO.FileExists(strDriverFile) Then
		Reporter.ReportEvent micFail, "Check driver File existance", "Driver File does not exist"
		Set objFSO = nothing
		Exit Function
	End If
	GetConfigurationDetails strDriverFile
	
	If Not FolderStructureCheck Then
		Reporter.ReportEvent micFail, "Check for Folder Structure", "Folder structure is not correct or could not be maintained "
	End If
	strProjectFolder = Environment.Value("FOLDERSTRUCTURE")
	newDriverLocation = "C:/TCOE" & "/" & strProjectFolder  & "/DriverFile/" & objFSO.getFileName(strDriverFile)
	objFSO.CopyFile strDriverFile, objFSO.GetParentFolderName(newDriverLocation) & "/", TRUE
	strDriverFile = newDriverLocation
	Environment("DRIVER_FILE").Value = strDriverFile
	Set objFSO = nothing
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	
	checkTestCaseSheet = checkSheetPresent(strDriverFile, "TestCase")
	if Not checkTestCaseSheet Then
		Reporter.ReportEvent micFail, "Check driver File for TestCase", "TestCase Sheet DoesNot exist in driverFile"
		Exit Function
	Else 			
		Set DBConnection_DriverFile = InitializeDriverFile()
		Dim  strSQL,objRS
		Set objRS= CreateObject("ADODB.Recordset")
		strSQL = "Select SlNo, TestCaseName, Iteration, TestDescription, Execution From [TestCase$]" 
		If Not ExeSQL(DBConnection_DriverFile,strSQL, objRS, True, IterationCount) then
			strError = "Get test case details: " & strError
			Reporter.ReportEvent micFail, "Get test case details", strError
					Exit function
				End if
				IterationCount = objRS.RecordCount
				If Not InitializeTestData() Then
			Reporter.ReportEvent micFail,"InitializeTestData - Connect to test data file", strError
			Exit Function
		End If
		
		If Not downloadReportImage() Then
			Reporter.ReportEvent micFail,"downloadReportImage - ", strError
			Exit Function
		End If
		
		If Environment.Value("TYPE_OF_TEST")="API" Then
			InitializeAPIReportCount = 0
			End  IF
			If Environment.Value("TYPE_OF_TEST")<>"API" Then
				If Not InitializeRepository() Then
					Reporter.ReportEvent micFail,"InitializeRepository - Connect to Repository file",strError
					Exit Function
				End If
			Else
				Reporter.ReportEvent micDone, "Initialize Repository", "Initialize Repository not required for API tests"
			End If
			If Not InitializeConsolidatedReport() Then
				Reporter.ReportEvent micFail,"Failed to Initialize the Consolidated HTML Report",strError
				Exit Function
			End If
			totalConsolidatedTestCases = 0
			totalConsolidatedPassed = 0
			totalConsolidatedFailed = 0
			totalConsolidatedSkipped = 0
					For iCount = 1 to IterationCount
					 valPreviousStepName=""
					 valPreviousStepDescription=""
					 valPreviousStepExpectedResult=""
					 valPreviousStepProvingorNonProving=""
					 valScreenShot = ""
				strSlNoTCS = Trim(objRS.Fields.Item(0).Value)
				strTestCaseName = Trim(objRS.Fields.Item(1).Value)
				strIteration = Trim(objRS.Fields.Item(2).Value)
				strTestDescription = Trim(objRS.Fields.Item(3).Value)
				strExecution = Trim(objRS.Fields.Item(4).Value)
				If UCASE(strExecution)= "Y" or UCASE(strExecution)="YES" Then
					ExecuteScript ="Y"
					Environment.Value("CURRENT_TESTCASE_NAME") = strTestCaseName
					Environment.Value("CURRENT_TEST_ITERATION") = strIteration
					Environment.Value("CURRENT_TEST_DESCRIPTION") = strTestDescription
					Environment.Value("CURRENT_TEST_SLNO") = strSlNoTCS
					TCStep = 0 'Added to keep track of API tests
					If Not InitializeReport() Then
						Reporter.ReportEvent micFail,"InitializeReport - Initialize the HTML Report",strError
						Exit Function
					End If
					
					
					
					
					Dim ContinueExecution
					ContinueExecution = True
					
					While ContinueExecution
						Select Case Trim(Ucase( Environment.Value("TRIGGER_FROMQC")))
						Case "YES"
														Reporter.ReportEvent micPass,"ExecuteTestCase","Triggering the execution from Quality Center."
							IterationToRun = Cint(strIteration)
							ContinueExecution = False
							ExecuteScript = "Y"
							TestcaseName = strTestCaseName
							End  Select
							TestResult = "Passed"
							InitializeAPIReportCount = 0
							Do
								If Not Trim(UCase(ExecuteScript)) = "Y" Then
									TestResult = "Not Run"
									Exit Do
								End If
								
								If Not GetTestDataIterations() Then
									TestResult = "Not Run"
									Reporter.ReportEvent micFail,"GetTestDataIterations", strError
									Exit Do
								End If 
								TestStartTime = Now
								Dim iNo
								For Each iNo in arrIteration	
									StartNewIteration
									CurrentIteration = iNo
									Reporter.ReportEvent micPass,"Start  iteration - " & CurrentIteration,"Starting the Iteration no [" & CurrentIteration & "] of the test case [" & Environment.Value("CURRENT_TESTCASE_NAME") & "]" 
									Set Testparameters  = CreateObject("Scripting.Dictionary")
									If Not RetrieveTestData(CurrentIteration) Then
										Reporter.ReportEvent micFail, "RetrieveTestData: iteration - " & CurrentIteration, strError
										TestResult = "Not Run"
										Exit Do
									End If		
									Call QTP_Driver(DBConnection_DriverFile)
									Call TestTermination
									
									If  blnExitTestCase Then
										Exit For
									End If
									If strExitTest="Yes" then
										ExitTest
									End if	
									
								Next
								TestEndTime = Now
								Exit Do
							Loop While False
							If Trim(UCase(ExecuteScript)) = "Y" Then
								ReportTestResult()
								ContinueExecution= false
							End If							
						WEnd
'====================Dispose all connections to MS Excel====================
						CompleteReport()
						
						
						If Err.Number <> 0 Then   
								Err.Clear
						End If
					End If
					objRS.MoveNext
					If Err.Number <> 0 Then
						strError =  "Error while performing MoveNext method on recordset" & VbCrLf & Err.Description
						objRS.Close: Set objRS = nothing
						Exit Function
					End If
				Next
				Call consolidatedReportFinilizer(totalConsolidatedCases, totalConsolidatedPassed, totalConsolidatedFailed, totalConsolidatedSkipped)
				Call ReplaceConsolidatedData()
				objRS.Close: Set objRS = nothing
				End  If
				DBConnection_Results.Close
				Set DBConnection_Results = Nothing
				DisposeDBConnection()
				KillProcess()		
				Set objFSO = nothing
				
			End Function
			
			
			Public Sub DisposeDBConnection
				DBConnection_TestData.Close
				Set DBConnection_TestData = nothing
			End Sub
			
'==================================================
'* Function Name -	GetConfigurationDetails								
'* Function Description -	Retrieves all config details from driverfile and stores in Environment						
'* Created By -	Tribikram Tripathy									
'* Created Date -	  04/15/2024								
'* Input Parameter -	DRIVER_FILE sheet path							
'* Output Parameter -   NA								
'* Pre-Conditions -	DRIVER_FILE needs to be present								
'* Post Conditions -	All config details are retrieved  from driverfile and stored in Environment							
'==================================================
			
			
			Public Sub GetConfigurationDetails(ExecutionExcel)
				
				On Error Resume Next
				
				If Not checkSheetPresent(ExecutionExcel, "Config") Then
					Reporter.ReportEvent micWarning, "check config File present in driver file", "config File not present in driver file"
					Exit Sub
				End If
				
				Set objExcel = CreateObject("Excel.Application")
					
				 	objExcel.Visible = False
   				 	objExcel.DisplayAlerts = False
					Set objWorkbook = objExcel.Workbooks.Open(ExecutionExcel, False, True)
					Set objSheet = objWorkbook.Worksheets("Config")
					columncount = objSheet.UsedRange.Columns.Count
					RowCount = objSheet.UsedRange.Rows.Count
				For i = 1 To columncount
						Environment.Value(Trim(objSheet.Cells(1,i).value))= Trim(objSheet.Cells(2,i).value)
				Next
					Set objSheet  = Nothing
					Set objWorkbook = Nothing
					Set objExcel = Nothing
				
				If Err.Number <> 0 Then 
						Msgbox err.Description
						Set objSheet  = Nothing
							Set objWorkbook = Nothing
							Set objExcel = Nothing
				End If
				
			End Sub
			
			Public Function FolderStructureCheck()
				On Error Resume Next
				Dim fso
					Set fso = CreateObject("Scripting.FileSystemObject")
					strProjectFolder = Environment.Value("FOLDERSTRUCTURE")
					strTempFolder = "C:\TCOE"
					If Not (fso.FolderExists(strTempFolder)) Then
							fso.CreateFolder ( strTempFolder )	
					Reporter.ReportEvent micDone, "Checking Temp folder", "Temp folder created successfully"    		
					End If
					If Not (fso.FolderExists(strTempFolder & "\" & strProjectFolder)) Then
							fso.CreateFolder ( strTempFolder & "\" & strProjectFolder )	
					Reporter.ReportEvent micDone, "Checking Project folder", "Project folder created successfully"    		
					End If
					strDriverFolder = strTempFolder & "\" & strProjectFolder & "\" & "DriverFile"
					strFunctionLibraryFolder =  strTempFolder & "\" & strProjectFolder & "\" & "FunctionLibrary"
					strResourcesFolder = strTempFolder & "\" & strProjectFolder & "\" & "Resources"
					strResultsFolder = strTempFolder & "\" & strProjectFolder & "\" & "Results"
					strRepositoryFolder = strTempFolder & "\" & strProjectFolder & "\" & "Repository"
					If Not (fso.FolderExists(strDriverFolder)) Then
							fso.CreateFolder ( strDriverFolder )	
					Reporter.ReportEvent micDone, "Checking Driver folder", "Driver folder created successfully"
					End If
					Environment.Value("DRIVER_FOLDER") = 	strDriverFolder
					If Not (fso.FolderExists(strFunctionLibraryFolder)) Then
							fso.CreateFolder ( strFunctionLibraryFolder )	
					Reporter.ReportEvent micDone, "Checking FunctionLibrary folder", "FunctionLibrary folder created successfully" 
					End If
					Environment.Value("FUNCTION_LIBRARY_FOLDER") = 	strFunctionLibraryFolder
					If Not (fso.FolderExists(strResourcesFolder)) Then
							fso.CreateFolder ( strResourcesFolder )	
					Reporter.ReportEvent micDone, "Checking Resources folder", "Resources folder created successfully" 
					End If
					Environment.Value("RESOURCES_FOLDER") = 	strResourcesFolder
					If Not (fso.FolderExists(strResultsFolder)) Then
							fso.CreateFolder ( strResultsFolder )	
					Reporter.ReportEvent micDone, "Checking Results folder", "Results folder created successfully" 
					End If
					
					strResourcesImgFolder = Environment("RESOURCES_FOLDER")& "\img"
					If Not (fso.FolderExists(strResourcesImgFolder)) Then
							fso.CreateFolder ( strResourcesImgFolder )
							Reporter.ReportEvent micDone, "Checking Image folder", "Image folder created successfully" 
				End if 
					Environment.Value("RESOURCES_IMG") = 	strResourcesImgFolder
					
					Environment.Value("RESULTS_FOLDER") = 	strResultsFolder
					If Not (fso.FolderExists(strRepositoryFolder)) Then
							fso.CreateFolder ( strRepositoryFolder )	
					Reporter.ReportEvent micDone, "Checking Repository folder", "Repository folder created successfully" 
					End If
					Environment.Value("REPOSITORY_FOLDER") = 	strRepositoryFolder
					Set fso = nothing
					If Err.Number <> 0 Then
							Reporter.ReportEvent micFail, "Checking Folder structure", "Folder structure is incorrect " & Err.Description
							FolderStructureCheck = false
					End If
					
					Set fso = nothing
					FolderStructureCheck = True
					
					
			End Function
			
			
			
			Public Function InitializeDriverFile()
				On Error Resume Next
					Err.Clear
					strError = ""
					strDriverFile = Environment.Value("DRIVER_FILE")
				Set DBConnection_DriverFile = ConnectToExcel(Environment.Value("DRIVER_FILE"),"YES")
				If Not DBConnection_DriverFile.State = 1 Then
					strError = "Failed to open the test data file as database present in the parth - " & strDriverFile
					Reporter.ReportEvent micFail,"InitializeDriverFile - Connect to Driver File","Failed to open the Driver file as a database - " & strDriverFile & "."
					Exit Function
				End If
				InitializeDriverFile = DBConnection_DriverFile
			End Function
			
			Public Function downloadDriverFile()
				On Error Resume Next
					Err.Clear
					strError = ""
					strDriverFile = Environment.Value("DRIVER_FILE")
					tarray =Split (strDriverFile,"\")
					fileName = tarray(Ubound(tarray))
					If Instr(strDriverFile,":\") = 0 Then
							On Error Resume Next
					Err.Clear
					Dim objFSO1,newFile
					Set objFSO1 = CreateObject("Scripting.FileSystemObject")
					Err.Clear    
					downloadPath = "C:\TCOE"
					statusDriverFile = DownloadResourceFromQC(strDriverFile,downloadPath)	
					If not objFSO1.FileExists(statusDriverFile) Then
						Reporter.ReportEvent micFail, "driver file download", "Could not download driver file "
						Exit Function
					End If
					Set objFSO1 = nothing
					Environment.Value("DRIVER_FILE") = 	downloadPath	& "\" & fileName
					End If
			End Function
			
			Function DownloadResourceFromQC(FilePath,fileType)
					On Error Resume Next
				Err.Clear
				
				Select Case UCASE(fileType)
				Case "DRIVER"
					folderPath = Environment.Value("DRIVER_FOLDER")
				Case "FUNCTIONLIBRARY"
					folderPath = Environment.Value("FUNCTION_LIBRARY_FOLDER")
				Case "RESOURCES"
					folderPath = Environment.Value("RESOURCES_FOLDER")
				Case "REPOSITORY"
					folderPath = Environment.Value("REPOSITORY_FOLDER")
				Case "REQUEST"
					folderPath = Environment.Value("API_REQUEST_FOLDER")
				Case "RESPONSE"
					folderPath = Environment.Value("API_RESPONSE_FOLDER")
				Case Else
					folderPath = 	 fileType					
				End Select
				
				If Instr(FilePath,":\") = 0 Then
					
					Dim tempArray,tempString,oResourceFolder
					Set qcConn = QCUtil.QCConnection
					Set oResourceFolder = qcConn.QCResourceFolderFactory
					tempString1 = Replace(FilePath,"[QualityCenter\Resources] ","")
					tempString = Replace(tempString1,"Resources\","")
					tempArray = Split(tempString,"\")
					For i= 0 To UBound(tempArray) - 1
						Set oFilter = oResourceFolder.Filter
						oFilter.Filter("RFO_NAME") = "'" & tempArray(i) & "'"
						Set oResourceFolderList = oFilter.NewList
						If oResourceFolderList.Count = 1 Then
							Set oResourceFolder = oResourceFolderList.Item(1)
						End If
						If Not UBound(tempArray) - 1 = i Then
							Set oResourceFolder = oResourceFolder.QCResourceFolderFactory
						End If
					Next
					
					Set oResource = oResourceFolder.QCResourceFactory
					Set oFilter = oResource.Filter
					oFilter.Filter("RSC_FILE_NAME") = tempArray(UBound(tempArray))
					Set oResourceList = oFilter.NewList
					If oResourceList.Count = 1 Then
						Set oFile = oResourceList.Item(1)
						oFile.FileName = tempArray(UBound(tempArray))
						oFile.DownloadResource folderPath, True
					End If
					Set oFile = Nothing
					Set oResourceList = Nothing
									Set oFilter = Nothing
					Set oResource = Nothing
					Set oResourceFolderList = Nothing
					Set oResourceFolder = Nothing
					Set qcConn = Nothing					
'								Set oFlieList = Nothing
					
					
					DownloadResourceFromQC = folderPath & "\" &  tempArray(UBound(tempArray))
				Else
					Set oFSO = CreateObject("Scripting.FileSystemObject")
					tempArray = Split(FilePath,"\")
					strDestinationFile = folderPath & "\" &  tempArray(UBound(tempArray))
					oFSO.CopyFile FilePath, strDestinationFile
					DownloadResourceFromQC = strDestinationFile
					Set oFSO = Nothing
				End If
				
				If Err.Number <> 0 Then
					Reporter.ReportEvent micFail,"error downloading " & FilePath , err.Description
					DownloadResourceFromQC = "ERROR"
					If not ( isNull(oFSO)  or isEmpty(oFSO)) Then
						Set oFSO = Nothing
					End If
				End If                   				
			End Function
			
			Public Function InitializeRepository()
				RepositoriesCollection.RemoveAll()
				tsrfile = Environment.Value("TSR_FILE_LIST")
				txtfile = Environment.Value("TXT_FILE_LIST")
				On Error Resume Next
				Dim RepositoryFile,Obfso
				Dim arrTempTsrFiles, arrTempTxtFiles, iCount
				
				InitializeRepository = False
				strError = ""
				saveToPath = "repository"
				
				arrTempTsrFiles =  Split(tsrfile, ",")
					arrTempTxtFiles = Split(txtfile, ",")
				RepositoriesCollection.RemoveAll() 
				
				Set Obfso = CreateObject("Scripting.FileSystemObject")
				For iCount = 0 to UBound(arrTempTsrFiles) 			
					strTempfile =  Environment.Value("REPOSITORY_FOLDER") & "\" & Trim(arrTempTsrFiles(iCount))
					If Obfso.FileExists(strTempfile) Then
						Obfso.DeleteFile (strTempfile)                     	
						If Obfso.FileExists(strTempfile) Then
							strError = "Unable to delete existing TSR file in the path : " & strTempfile & VbCrLf & Err.Description 
							Exit Function
						End if 
					End If
					
					RepositoryFile = Environment("REPOSITORY_PATH") & "\TSR\" & Trim(arrTempTsrFiles(iCount))
					RepositoryFile =  DownloadResourceFromQC(RepositoryFile, "REPOSITORY") 
					If NOT Obfso.FileExists(RepositoryFile) Then
						strError = ".TSR file is not found in the path : " & RepositoryFile
						Exit Function
					End If
					RepositoriesCollection.Add RepositoryFile,1
					qtTest.LoadRepository RepositoryFile
'ObjectRepositoryUtil.Load(RepositoryFile)
					
				Next
				
				
				
				For iCount = 0 to UBound(arrTempTxtFiles) 				
					strTempfile =  Environment.Value("REPOSITORY_FOLDER") & "\" & Trim(arrTempTxtFiles(iCount))
					If Obfso.FileExists(strTempfile) Then
						Obfso.DeleteFile (strTempfile) 
						If Obfso.FileExists(strTempfile) Then
								strError = "Unable to delete existing  TXT file in the path : " & strTempfile & VbCrLf & Err.Description 
							Exit Function
						End if 
					End If
					
					RepositoryFile = Environment("REPOSITORY_PATH") & "\TXT\" & Trim(arrTempTxtFiles(iCount))
					RepositoryFile = DownloadResourceFromQC(RepositoryFile, "REPOSITORY")
					If NOT Obfso.FileExists(RepositoryFile) Then
						strError = ".TXT  file is not found in the path : " & RepositoryFile
						Exit Function
					End If
						Call LoadObjectsFromRepository(RepositoryFile)
					
				Next 
				
				Set arrTempTsrFiles = Nothing: Set arrTempTxtFiles = Nothing
				
				Set Obfso = Nothing
				isDescriptive = False
				InitializeRepository = True
				
			End Function
			
			Public Function downloadReportImage
				
				On Error Resume Next
				
				Dim BaxterBackGroundFile,objFSO
				downloadReportImage = False
				strError= ""
				
				Set objFSO = CreateObject("Scripting.FileSystemObject") 
				strTempfile =  Environment.Value("RESOURCES_IMG") & "\" &  "BaxterBackGround.png"
				If objFSO.FileExists(strTempfile) Then
					objFSO.DeleteFile (strTempfile)
					If objFSO.FileExists(strTempfile) Then
						strError = "Unable to delete existing Test Data file in the path : " & strTempfile
						Exit Function
					End If
				End if 
				
				BaxterBackGroundFile = Environment("IMG_LOGO") &"\BaxterBackGround.png"
				BaxterBackGroundFile = DownloadResourceFromQC(BaxterBackGroundFile, Environment.Value("RESOURCES_IMG"))
				If Not objFSO.FileExists(BaxterBackGroundFile) Then
					strError="Test data file not present in the path - " & strTempfile
					Exit Function
				End If
				
				Set objFSO = CreateObject("Scripting.FileSystemObject") 
				strTempfile =  Environment.Value("RESOURCES_IMG") & "\" &  "transparentLogo.png"
				If objFSO.FileExists(strTempfile) Then
					objFSO.DeleteFile (strTempfile)
					If objFSO.FileExists(strTempfile) Then
						strError = "Unable to delete existing Test Data file in the path : " & strTempfile
						Exit Function
					End If
				End if 
				
				TransparentLogoFile = Environment("IMG_LOGO") &"\transparentLogo.png"
				TransparentLogoFile = DownloadResourceFromQC(TransparentLogoFile, Environment.Value("RESOURCES_IMG"))
				If Not objFSO.FileExists(TransparentLogoFile) Then
					strError="Test data file not present in the path - " & strTempfile
					Exit Function
				End If
				
				downloadReportImage = True
				
				Set objFSO = Nothing
				If err.number<>0 Then
					strError = err.description
				End If
				
			End Function
			
			Public Function InitializeTestData
				
				On Error Resume Next
				
				Dim TestDataFile,objFSO,strTemp,strTemp1,strTempfile1
				InitializeTestData = False
				strError= ""
				
				Set objFSO = CreateObject("Scripting.FileSystemObject") 
					
				tArray = Split(Environment("TESTDATA_FILE"),"\")
				testDataFileName = tArray(UBOUND(tArray))
				strTempfile =  Environment.Value("RESOURCES_FOLDER") & "\" &  testDataFileName
				
				If objFSO.FileExists(strTempfile) Then
					objFSO.DeleteFile (strTempfile)
					If objFSO.FileExists(strTempfile) Then
						strError = "Unable to delete existing Test Data file in the path : " & strTempfile
						Exit Function
					End If
				End if 
				
				TestDataFile = Environment("TESTDATA_FILE")
				TestDataFile = DownloadResourceFromQC(TestDataFile,"RESOURCES")
				If Not objFSO.FileExists(TestDataFile) Then
					strError="Test data file not present in the path - " & strTempfile
					Exit Function
				End If
				
'Connect to the MS Excel testdata file
				Set DBConnection_TestData = ConnectToExcel(TestDataFile,"YES")
				
				If Not DBConnection_TestData.State = 1 Then
					strError = "Failed to open the test data file as database present in the parth - " & TestDataFile
					Exit Function
				End If
				
				tArray = Split(Environment("APPURLSXLS_PATH"),"\")
				AppURLsXLSFileName = tArray(UBOUND(tArray))
				strTempfile1 =  Environment.Value("RESOURCES_FOLDER") & "\" & AppURLsXLSFileName
				
				If objFSO.FileExists(strTempfile1) Then
					objFSO.DeleteFile (strTempfile1)
					If objFSO.FileExists(strTempfile1) Then
						strError = "Unable to delete existing Env App file in the path : " & strTempfile1
						Exit Function
					End If
				End if 
				
				AppURLsXLSFile = Environment("APPURLSXLS_PATH")
				AppURLsXLSFile = DownloadResourceFromQC(AppURLsXLSFile,"RESOURCES")
'					Set objFSO = CreateObject("Scripting.FileSystemObject") 
				
				If Not objFSO.FileExists(AppURLsXLSFile) Then
					strError = "Env app urls  file not present  in the path - " & strTempfile1
					Exit Function
				End If
				
				InitializeTestData = True
				
				Set objFSO = Nothing
				
			End Function
			
			Public Function GetTestDataIterations()
				
				GetTestDataIterations = False
				On Error Resume Next
				
				Err.Clear
				strError = ""
				
				Dim TestModule, strSQL,objRS
				Dim striteration, striterations
				TestModule = "["&  Environment("MODULE") & "$]"
				Set objRS= CreateObject("ADODB.Recordset")
				
				strSQL =  "SELECT DISTINCT(Iteration) FROM " & TestModule & " where TestName = """ & Environment.Value("CURRENT_TESTCASE_NAME") &  """ and Execute = 'Y' and  Header = ""Header Data"""
				If Not ExeSQL(DBConnection_TestData,strSQL, objRS, True, Environment.Value("CURRENT_TEST_ITERATION")) then
					strError = "GetTestDataIterations: " & strError
					Exit function
				End if 
				IterationCountTD = objRS.RecordCount
				For iCount = 1 to IterationCountTD
					strIteration = Trim(objRS.Fields.Item(0).Value)
					strIterations = strIterations & "," & strIteration
					objRS.MoveNext
					If Err.Number <> 0 Then
						strError =  "Error while performing MoveNext method on recordset" & VbCrLf & Err.Description
						objRS.Close: Set objRS = nothing
						Exit Function
					End If
				Next
				
				strIterations = Right(strIterations, Len(strIterations)-1)
				arrIteration = Split(strIterations, ",")
				
				objRS.Close: Set objRS = nothing
				GetTestDataIterations = True
				
			End Function
			
			Public Function RetrieveTestData(iterationNo)
				RetrieveTestData = False
				On Error Resume Next
				Err.Clear
				strError = ""
				Dim strSQL1,objRS1, rowCount1, colCount1, col
				Dim strSQL2,objRS2, rowCount2
				Dim strSQL3, objRS3, rowCount3
				Dim strLineItem, strLineItems, rowLineCount
				Dim blnExeSQL3, TestModule
				
				TestModule = "[" &  Environment("MODULE") & "$]"
				Set objRS1= CreateObject("ADODB.Recordset")
				Set objRS2= CreateObject("ADODB.Recordset")
				Set objRS3= CreateObject("ADODB.Recordset")
				
				strSQL1 = "SELECT * FROM " & TestModule & " where TestName = '" & Environment.Value("CURRENT_TESTCASE_NAME") &  "' and Execute = 'Execute' and  Header = 'Field Name'"
				If Not ExeSQL(DBConnection_TestData,strSQL1, objRS1,  True, rowCount1) Then 
					Exit function
				End if 
				
				If rowCount1 > 1 Then
					strError =  "More than one Header record found:: SQL - " & strSQL1 & VbCrLf & Err.Description
					objRS1.Close: Set objRS1 = nothing
					Exit Function
				End If
				strSQL2 = "SELECT * FROM " & TestModule & " where TestName = '" & Environment.Value("CURRENT_TESTCASE_NAME") &  "' and Execute = 'Y' and Header = 'Header Data' and Iteration = " & "'" & iterationNo & "'" 
				If Not ExeSQL(DBConnection_TestData,strSQL2, objRS2, True, rowCount2) Then 
					Exit function
				End if
				If rowCount2 > 1 Then
					strError =  "More than one Data record found for iteration " & iterationNo &" : SQL - " & strSQL2 & VbCrLf & Err.Description
					objRS2.Close: Set objRS2 = nothing
					Exit Function
				End If
				strSQL3 = "SELECT * FROM " & TestModule & " where TestName = '" & Environment.Value("CURRENT_TESTCASE_NAME") &  "' and Execute = 'Y' and Header = 'Grid Data' and Iteration = " & "'" & iterationNo & "'" 
				
				colCount1=objRS1.Fields.Count
				blnExeSQL3 = True
				For col =2 To colCount1
					Dim pName, pValue
					pName = Trim(objRS1.Fields.Item(col).Value)
					If IsNull(pName) Or IsEmpty(pName) Then
						Exit For
					End If
					pValue = Trim(objRS2.Fields.Item(col).Value)
					If Left(pName,1) = "#" Then
						If  blnExeSQL3 Then
							If Not ExeSQL(DBConnection_TestData,strSQL3, objRS3, False, rowCount3) Then 
								Exit function
							End if
							blnExeSQL3 = False
						End If
						If rowCount3 > 0 Then
							objRS3.MoveFirst
							If Err.Number <> 0 Then
								strError =  "Error while performing MoveFirst method on recordset 3" & VbCrLf & Err.Description
								objRS1.Close: Set objRS1 = nothing: objRS2.Close: Set objRS2 = nothing: objRS3.Close: Set objRS3 = nothing
								Exit Function
							End If
							strLineItem = "": strLineItems = ""
							
							For rowLineCount = 1 to rowCount3
								strLineItem = Trim(objRS3.Fields.Item(col).Value)
								If  strLineItems <> "" Then
									strLineItems = strLineItems & "||" & strLineItem
								Else
									strLineItems = strLineItem
								End If
								
								objRS3.MoveNext
								If Err.Number <> 0 Then
									strError =  "Error while performing MoveNext method on recordset 3" & VbCrLf & Err.Description
									objRS1.Close: Set objRS1 = nothing: objRS2.Close: Set objRS2 = nothing: objRS3.Close: Set objRS3 = nothing
									Exit Function
								End If
							Next
							pValue = strLineItems
						End If
					End If
					TestParameters.Add pName, pValue
					If Err.Number <> 0  then   
						If Err.number <>94  Then
							strError =  "Error while performing add method on TestParameters object. " & pname  & "-" & pvalue & VbCrLf & Err.Description
							objRS1.Close: Set objRS1 = nothing: objRS2.Close: Set objRS2 = nothing: objRS3.Close: Set objRS3 = nothing
							Exit Function
						End If
					End If 
				Next
				
				
				objRS1.Close: Set objRS1 = nothing
				objRS2.Close: Set objRS2 = nothing
				objRS3.Close: Set objRS3 = nothing
				RetrieveTestData = True
				
			End Function
			
			Function ReplaceFirstOccurrence(str, find, replaceWith)
' Find position of substring
				pos = InStr(str, find)
				
' If substring is found
				If pos > 0 Then
' Replace first occurrence
					str = Left(str, pos - 1) & replaceWith & Mid(str, pos + Len(find))
				End If
				
				ReplaceFirstOccurrence = str
			End Function
			
			Public Function GetParameterValue(pName)
				
				On Error Resume Next
				If Left(TRIM(Ucase(pName)),10) = "USERSCRIPT" Then
					pName = Mid(Trim(Mid( Trim( pName), 11 ) ), 2)
					pName = Left(pName, Len(pName) - 1)
					pName = Eval(pName)
					If err.number<>0 Then
						pName = ""
						Reporter.ReportEvent micFail , "Fetch parameter value ", "Error while fetching parameter value for USERSCRIPT " & err.description
						Exit Function
					End If
				ElseIf  RIGHT(TRIM(pName), 1) = chr(34) and RIGHT(TRIM(pName), 1) = chr(34) Then
					pName =  MID(TRIM(pName), 2, LEN(TRIM(pName))-2)
				Else
					If inputparameters.Exists(pName) Then
						GetParameterValue = inputparameters.Item(pName)
					ElseIf outputparameters.Exists(pName) Then
						GetParameterValue = outputparameters.Item(pName)
					ElseIf testparameters.Exists(pName) Then
						GetParameterValue = testparameters.Item(pName)
					End If
					End  If
					If  GetParameterValue = "" Then
						If pName = "Application_EnvironmentName" Then
							GetParameterValue = "" 
						ElseIf pName = "Environment_Type" Then
							GetParameterValue = "" 
						Else
							GetParameterValue = pName
						End If
					End If
					On error goto 0
				End Function
				
				Public Function QTP_Driver(DBConnection_DriverFile)
					
					On Error Resume Next 
					Err.Clear
'Dim excel_header_row 
'If IsEmpty(Environment.Value("HEADER_ROW")) or Environment.Value("HEADER_ROW")="" then
'	excel_header_row = HEADER_ROW
'Else
'	excel_header_row = Environment.Value("HEADER_ROW")
'End if
					
'Dim SheetNumber
'SheetNumber = -1
'Set oRS = DBConnection_DriverFile.OpenSchema(20)
'Do Until oRS.EOF 
'	Dim tName 
'	sSheetName = oRS.Fields("table_name").Value
'	if Left(sSheetName,Len("TestScript")+1) = "TestScript$" then
'		SheetNumber = Mid(sSheetName,Len("TestScript")+2, Len(sSheetName) - Len("TestScript") -2)
'		Exit Do
'	End if
'oRS.MoveNext
'Loop
					data_range = "[TestScript$]"
'data_range = "["&"TestScript"&"$"& Cells(excel_header_row+1,1).Address & ":" & Cells(Rows.count,1).End(xlUp).Address & "]"
					
					sqlQuery = "Select * from " & data_range & "WHERE TestCaseName = '"& Environment.Value("CURRENT_TESTCASE_NAME") & "'"
					blnOneRec = false
					Set RS= CreateObject("ADODB.Recordset")
					ExeSQL DBConnection_DriverFile, sqlQuery, RS, blnOneRec, numRecCnt 
					
					If strExitTest="Yes" then
						Exit function
					End if
					DIM fTestCaseName,fslNo, fStepName, fStepDescription, fExpectedResult, fProvingorNonProving, fAutomationStepDescription, fScreenName
					DIM fObject, fAction, fParam1, fParam2, fParam3, fParam4, fParam5, fErrorHandler, fScreenShot
					
					Set fTestCaseName = RS.Fields.Item("TestCaseName")
					Set fslNo = RS.Fields.Item("SNo")
					Set fStepName = RS.Fields.Item("Step Name")
					Set fStepDescription = RS.Fields.Item("Step Description")
					Set fExpectedResult = RS.Fields.Item("Expected Result")
					Set fProvingorNonProving = RS.Fields.Item("Proving/NonProving")
					Set fAutomationStepDescription = RS.Fields.Item("Automation Step Description")
					Set fScreenName = RS.Fields.Item("ScreenName")
					Set fObject = RS.Fields.Item("Object")
					Set fAction = RS.Fields.Item("Action")
					Set fParam1 = RS.Fields.Item("Param1")
					Set fParam2 = RS.Fields.Item("Param2")
					Set fParam3 = RS.Fields.Item("Param3")
					Set fParam4 = RS.Fields.Item("Param4")
					Set fParam5 = RS.Fields.Item("Param5")
					Set fErrorHandler = RS.Fields.Item("Error Handler")
					Set fScreenShot = RS.Fields.Item("ScreenShot")
					
					
					
					Do While not RS.EOF
						
						blnQCPrintVal = True
						blnSkipRslt = True
						If strExitTest="Yes" then
							Exit function
						End if
						
						valTestCaseName = fTestCaseName.Value
						If isNull(valTestCaseName) or isEmpty(valTestCaseName) Then
							valTestCaseName = ""
						End If
						valslNo = fslNo.Value
						If isNull(valslNo) or isEmpty(valslNo) Then
							valslNo = ""
						End If
						valStepName = fStepName.Value
						If isNull(valStepName) or isEmpty(valStepName) Then
							valStepName = ""
						End If
						valStepDescription = fStepDescription.Value
						If isNull(valStepDescription) or isEmpty(valStepDescription) Then
							valStepDescription = ""
						End If
						valExpectedResult = fExpectedResult.Value
						If isNull(valExpectedResult) or isEmpty(valExpectedResult) Then
							valExpectedResult = ""
						End If
						valProvingorNonProving = fProvingorNonProving.Value
						If isNull(valProvingorNonProving) or isEmpty(valProvingorNonProving) Then
							valProvingorNonProving = ""
						End If
						valAutomationStepDescription = fAutomationStepDescription.Value
						If isNull(valAutomationStepDescription) or isEmpty(valAutomationStepDescription) Then
							valAutomationStepDescription = ""
						End If
						valScreenName = fScreenName.Value
						If isNull(valScreenName) or isEmpty(valScreenName) Then
							valScreenName = ""
						End If
						valObject = fObject.Value
						If isNull(valObject) or isEmpty(valObject) Then
							valObject = ""
						End If
						valAction = fAction.Value
						If isNull(valAction) or isEmpty(valAction) Then
							valAction = ""
						End If
						valParam1 = fParam1.Value
						If isNull(valParam1) or isEmpty(valParam1) Then
							valParam1 = ""
						End If
						valParam2 = fParam2.Value
						If isNull(valParam2) or isEmpty(valParam2) Then
							valParam2 = ""
						End If
						valParam3 = fParam3.Value
						If isNull(valParam3) or isEmpty(valParam3) Then
							valParam3 = ""
						End If
						valParam4 = fParam4.Value
						If isNull(valParam4) or isEmpty(valParam4) Then
							valParam4 = ""
						End If
						valParam5 = fParam5.Value
						If isNull(valParam5) or isEmpty(valParam5) Then
							valParam5 = ""
						End If
						valErrorHandler = fErrorHandler.Value
						If isNull(valErrorHandler) or isEmpty(valErrorHandler) Then
							valErrorHandler = ""
						End If
						valScreenShot = fScreenShot.Value
						If isNull(valScreenShot) or isEmpty(valScreenShot) Then
							valScreenShot = ""
						End If
						
						if not (valStepName = "" or isnull(valStepName) or isempty(valStepName)) then
							valPreviousStepName = valStepName
							Environment("valPreviousStepName").Value = valPreviousStepName
						end if
						
						if not (valStepDescription = "" or isnull(valStepDescription) or isempty(valStepDescription)) then
							valPreviousStepDescription = valStepDescription
							Environment("valPreviousStepDescription").Value = valPreviousStepDescription
						end if
						
						if not (valExpectedResult = "" or isnull(valExpectedResult) or isempty(valExpectedResult)) then
							valPreviousStepExpectedResult = valExpectedResult
							Environment("valPreviousStepExpectedResult").Value = valPreviousStepExpectedResult
						end if
						
						if not (valProvingorNonProving = "" or isnull(valProvingorNonProving) or isempty(valProvingorNonProving)) then
							valPreviousStepProvingorNonProving = valProvingorNonProving
							Environment("valPreviousStepProvingorNonProving").Value = valPreviousStepProvingorNonProving
						end if
						
						Call InitializeStepExecution()
						
						If Trim(Environment.Value("OVERWRITE_ERRORHANDLER")) = "" or isEmpty(Environment.Value("OVERWRITE_ERRORHANDLER")) or UCASE(Environment.Value("OVERWRITE_ERRORHANDLER"))="NO" or UCASE(Environment.Value("OVERWRITE_ERRORHANDLER"))="N"Then
							valErrorHandler = valErrorHandler
						Else
							valErrorHandler = Trim(Environment.Value("OVERWRITE_ERRORHANDLER"))
						End if 	
						stepResult = "Passed"
						
						
						strScreenName = valScreenName
						strObjectName = valObject
						strAction = valAction
						strErrorHandler = valErrorHandler
						strParameters = valParam1
						
						Call performAction(strAction)
						
						FinishStepExecution()
'Call ReportToQC()
						stepErrDescription = ""
						If  blnExitIteration Or blnExitTestCase Then
							Exit Do
						End If
						If strExitTest="Yes" then
							Exit function
						End if
						RS.MoveNext
					Loop
					
					Set RS = Nothing
					Set fTestCaseName = Nothing
					Set fslNo = Nothing
					Set fStepName= Nothing
					Set fStepDescription = Nothing
					Set fExpectedResult = Nothing
					Set fProvingorNonProving = Nothing
					Set fAutomationStepDescription = Nothing
					Set fScreenName = Nothing
					Set fObject = Nothing
					Set fAction = Nothing
					Set fParam1 = Nothing
					Set fParam2 = Nothing
					Set fParam3 = Nothing
					Set fParam4 = Nothing
					Set fParam5 = Nothing
					Set fErrorHandler = Nothing
					Set fScreenShot = Nothing
					
				End Function
				
				
				Sub InitializeStepExecution()
					
					On Error Resume Next
					Err.Clear
					
					Call ClearErrors()
					If valStepName = "" OR ISEMPTY(valStepName) OR ISNULL(valStepName)Then
						valStepName = valPreviousStepName
					End If
					If (valStepDescription = "" OR ISEMPTY(valStepDescription) OR ISNULL(valStepDescription)) Then
						valStepDescription = valPreviousStepDescription
					End If 
					If (valExpectedResult = "" OR ISEMPTY(valExpectedResult) OR ISNULL(valExpectedResult) ) Then
						valExpectedResult = valPreviousStepExpectedResult
					End If
					If (valProvingorNonProving = ""  OR ISEMPTY(valProvingorNonProving) OR ISNULL(valProvingorNonProving) ) Then
						valProvingorNonProving = valPreviousStepProvingorNonProving
					End If
					StepResult = "Passed"
					mstepResult = "Passed"
					actualResult  = ""
					actualResult1  = ""
					stepActual=""
					StepStartTime = Now
					stepScreenshot = ""
					defaultstepExpected = ""
					
					
				End Sub
				
				
				Sub FinishStepExecution()
					
					On Error Resume Next
					If actualResult = "" Then
						If Not stepErrDescription = "" Then
							actualResult = stepErrDescription
							Reporter.ReportEvent micFail, valAutomationStepDescription & ":-" &UCASE(strAction), UCASE(strAction) & " - " & stepErrDescription
						Else
							actualResult = UCASE(strAction) & " - " & stepActual
							If blnQCPrintVal = False Then
								actualResult = ""
							End if
						End If
					End If
					
					If actualResult1 = "" Then
						If Not stepErrDescription = "" Then
							actualResult1 = stepErrDescription
						Else
								If blnQCPrintVal = True Then
									actualResult1 = valAutomationStepDescription & ":-" & UCASE(strAction) & "-" &stepActual
								End If
						End If
					Else
						If Not stepErrDescription = "" Then
							actualResult1 = actualResult1& vbLf & stepErrDescription
						Else
								If blnQCPrintVal = True Then
									actualResult1 =  actualResult1& vbLf & UCASE(strAction) & "-" &stepActual
							End If
						End If
					End If
					If  strAction <> "CALL_API_TEST" Then
						
						
						If UCase(Environment.Value("SCREENSHOT_ALLSTEPS")) = "YES" and UCASE(TRIM(Environment("TYPE_OF_TEST"))) <> "API" Then
'Capture the screenshot for all steps executed
							stepScreenshot = ResultsFolder & "\" & Environment.Value("CURRENT_TEST_SLNO") & Environment.Value("CURRENT_TESTCASE_NAME") & "_" & Environment.Value("CURRENT_TEST_ITERATION") & "_" & compIteration & "_" & valStepName & "_" & valslNo & ".png"
'Desktop.CaptureBitmap stepScreenshot
							call CaptureHilightedScreenShot(strObjectName , stepScreenshot )
						Else
'Capture the screenshot for all failed steps
							If (StepResult <> "Passed" or UCASE(valScreenShot) = "Y" or UCASE(valScreenShot) = "YES") and UCASE(TRIM(Environment("TYPE_OF_TEST"))) <> "API"  Then
								If stepScreenshot = "" Then
									stepScreenshot = ResultsFolder & "\" & Environment.Value("CURRENT_TEST_SLNO") &Environment.Value("CURRENT_TESTCASE_NAME") & "_" & Environment.Value("CURRENT_TEST_ITERATION") & "_"  & compIteration & "_" & valStepName & "_" & valslNo & ".png"
'Desktop.CaptureBitmap stepScreenshot
									call CaptureHilightedScreenShot(strObjectName , stepScreenshot )
								End If
							End If
							
						End If
					End If
					If StepResult <> "Passed" Then
						mstepResult = StepResult
						TestResult = StepResult
					End If
					
					Call ErrorHandler()
					If  blnExitIteration Or blnExitTestCase Then
						If strAction <> "CALLTEST" Then
							StepEndTime = Now
							stepDuration = GetExecutionTime(StepStartTime,StepEndTime)
							Call ReportResult()
							Call ReportToQC()
							actualResult  = ""
							actualResult1  = ""
							Exit Sub
						End If
					End If
					
					If strAction <> "CALLTEST"  Then
						StepEndTime = Now
						stepDuration = GetExecutionTime(StepStartTime,StepEndTime)
						Call ReportResult( )
						actualResult  = ""
						Call ReportToQC()
						actualResult1  = ""
						Exit Sub
					End if		
				End Sub
				
				Function ErrorHandler 
					If Len ( Trim ( valErrorHandler )) = 0 Then 
						Exit Function
					End If
					If  stepErrDescription = "" Then
						Exit Function
					End If
					
					If UCase(strAction) = "PRINT" Then
						Exit Function
					End If
					
					Select Case Ucase ( Trim ( strErrorHandler ) )
					Case "NEXT_TESTCASE"
						blnExitTestCase = True
					Case "NEXT_ITERATION" 
						blnExitIteration= True
					End Select
				End Function
				
				Public Function LoadObjectsFromRepository(RepositoryFile)
					Set ObjFSO = CreateObject("Scripting.FileSystemObject")
					Set objTextFile = ObjFSO.OpenTextFile(RepositoryFile, 1, false)
					Do Until objTextFile.AtEndOfStream
						strLine = objTextFile.ReadLine
						If  LEFT(TRIM(strLine),3)="Set" Then
							ExecuteGlobal strLine
						End If
					Loop
					objTextFile.Close
					Set objTextFile =nothing
					Set ObjFSO = nothing
				End Function
				
				Function ClearErrors
					
					stepErrDescription = ""
'strErrorHandler = ""
					Err.Clear
					
				End Function
				Function CaptureHilightedScreenShot(strObjectName , screenshotName )
					Dim oldBorder
					If  (isNullisEmptyCheck( strObjectName)<>"" ) Then
						If IsObject(eval(strObjectName)) Then
							If   Eval( strObjectName & ".Exist") Then
								oldBorder = Eval(strObjectName).Object.style.border
								Reporter.ReportEvent micDone, "Highlight Object", "Highlighted the object"
													Eval(strObjectName).Object.style.border = "4px solid orange"
							End If	
						End If
					End If
					
					Desktop.CaptureBitmap stepScreenshot
					
					If  ( isNullisEmptyCheck( strObjectName)<>"" ) Then
						If IsObject(eval(strObjectName)) Then
							If   Eval("strObjectName.Exists") Then
												 	Eval(strObjectName).Object.style.border = oldBorder
												 	Reporter.ReportEvent micDone, "De - Highlight Object", "De - Highlighted the object"
							End If
						End If	
					End If
				End Function
