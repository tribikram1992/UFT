Public strError,strTempfile
Dim batchNumber
Dim orderNumber,BrowserVersion

Public Function performAction(strAction)

Select  case strAction
	Case "INPUTPARAM"
		InputParam_Function()
	Case "OUTPARAM"
		oName = inpWorksheet.Cells(row, OBJECT_COLUMN).Value
		oValue =  inpWorksheet.Cells(row, PARAMETER_COLUMN).Value
		outputparameters.Add oName,GetParameterValue(oValue)
	Case "EXTRACT_RO_PROPERTY_AND_SAVE"
		extractROPropertyAndSave()
	'''Case "CALLTEST"
		'CalledExcelFile = inpWorksheet.Cells(row, PARAMETER_COLUMN).Value
				'Set tempws = inpWorksheet
				'QTP_Driver(CalledExcelFile)
				'Set inpWorksheet = tempws
				'''strAction = "CALLTEST"
	Case "INPUT"
		Call Input_Function()
	Case "INPUTGRIDWEBEDIT"
		Call  InputGridWebEdit()
	Case "GRIDPARAMVALUE"
		Call GRIDPARAMVALUEFN
	Case "CLICK"
		Call Click_Function()	
	Case "PRINT"
			'Usage:1
			'OBJECT		| ACTION	| 	| PARAMETERS
			'objEdit		| PRINT		|	| text msg to be printed (without comma), objectpropertyname_whose_value to be printed
			
			'Usage:2
			'OBJECT		| ACTION	| 	| PARAMETERS
			'				| PRINT		|	| text msg to be printed (without comma), %%variable or executable string
			Call PrintToReport(strObjectName, strParameters)
	Case "ENTERPASSWORD"
			objName = valObject
			strValue = GetParameterValue(valParam1)
			Call UdfEditSetSecure (objName, strValue)
	Case "SENDKEYS"
			'OBJECT		| ACTION	| 	| PARAMETERS
			'			| SENDKEYS	|	| ENTER
			'			| SENDKEYS	|	| VK_ENTER
			'TESCREEN	| SENDKEYS	|	| TE_ENTER
			'stepExpected = " Send the keys " & strParameters & " to the application/web page"
			Call SendKeys_Function()								
	Case "CLOSEAPP"
			Call CloseAllBrowsersExceptQc 
	Case "EXIST"
			Call Exist_Function()
	Case "VALIDATE"
			Call Validate_Function()
	Case "VALREGEXP"
			Call ValidateByRegExpObjectsProperty(strObjectName, strParameters)
	Case "WAIT"
			Call Wait_Function
			blnQCPrintVal = False
	Case "WAITUNTIL"
			Call WaitUntil_Function()		
	Case "CALL"
			Call CallFunction()		
	Case "USERSCRIPT"
			Call UserScriptFunction()					
	Case "OPERATION"
			Call Operation_Function()		
	Case "VARIABLE"
			Call RetrieveProperty_Function()
	Case "STOREVALUE"
			Call RetrieveValue_Function()
	Case "CLICKANYLINK"
			Call Clickanylink_Function()
	Case "ENTERGRIDLINES"
			Call EnterCustGridLines_Function()
	Case "DOWNLOADTESTSCRIPT"
			Call DownloadTestScript_Function()
	Case "GRIDRECORDCOUNT"
			Call GridRecordCount_Function()
	Case "DATEDAYSELECT"
			Call fnDateDaySelect() 	
	Case "VERIFY"
			Call Verify_Function()
	Case "ITEMCHECKLIST"
			Call ItemCheckList_Function()
	Case "ITEMSORTLIST"
			Call ItemSortList_Function()
	Case "PDFCLOSING"
			Call OpenPDFclosing_Function()
	Case "MULTISELECT"
			Call fnDropDownMultiSelection()		
	Case  "WEBTABLEWEBEDIT"
			call fnUpdateWebeditInWebtable()
	Case "CHECKBOXINWEBTABLE"
			Call fnCheckboxInWebtable()   	
	Case "SCROLLBAR"
			Call ScrollBar_Function()
	Case "OPENWEBAPPLICATION"
			Call UDFOpenURL()	
	Case "INPUT_CURRENT_DATE_AND_TIME"
			Call UDFInputDateAndCurrentTime()	
	Case "PAGESYNC"
			Call fnSynchronisePage()
	Case "EXTRACTBATCHNUMBER"
			Call fnBatchNumber()
	Case "DROPDOWNSELECTBYVALUE"
			Call dropDownSelectByValue()
	Case "GETDATA"
			Call extractDataFromElement()
	Case "CALL_API_TEST"
			Call APITest()
	Case else
		stepErrDescription =  "Keyword " & chr(34) & strAction & chr(34) & " is invalid"
		stepResult = "Failed"
End Select
End  Function

Function KillProcess()

   On Error Resume Next 
   Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}" & "!\\.\root\cimv2")
 
   Set colProcess = objWMIService.ExecQuery ("Select * From Win32_Process")
   For Each objProcess in colProcess
      If LCase(objProcess.Name) = LCase("excel.exe") Then
         objProcess.Terminate()
      End If
   Next
   Set colProcess = Nothing
   Set objWMIService = Nothing

End Function

Public Function GetExecutionTime(StartTime,EndTime)
					GetExecutionTime = datediff("s",StartTime,EndTime)
End Function

Public Function checkSheetPresent(ExcelFile, strSheetName)
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = False
   	objExcel.DisplayAlerts = False
	check = false
     	Set objWorkbook = objExcel.Workbooks.Open(ExcelFile, false, true)
     	For each Sheet In  objWorkbook.Sheets
     		If  Sheet.Name = strSheetName Then
            		check = True
            		Exit For
        	End If
     	Next
     	
    	objWorkbook.Close
    	Set objExcel = nothing 
	checkSheetPresent = check    	
End Function

Public Function getColumnValueByName(objSheet,intRowNum,strColName)
	Dim returnVal
	columncount = objSheet.UsedRange.Columns.Count
	For i = 1 To columncount
		If strColName = objSheet.Cells(1, i).value Then
			returnVal = objSheet.Cells(intRowNum, i).value
			Exit For
		End If
	Next
	getColumnValueByName = returnVal
End Function

Public Function DateString(dDate)
    DateString = Year(dDate)& right("0" & Month(dDate),2) & right("0" & Day(dDate),2) & right("0" & Hour(dDate),2) & right("0" & Minute(dDate),2) & right("0" & second(dDate),2)
End Function

Public Function ConnectToExcel(ExcelPath,HDRValue)
     On Error Resume Next
     Err.Clear

     Dim objConnection,dbProvider,dbSource,dbExtended
     Set objConnection = CreateObject("ADODB.Connection")
     'dbProvider =  "Provider=Microsoft.Jet.OLEDB.4.0;"
     dbProvider =  "Provider=Microsoft.ACE.OLEDB.16.0;"
     dbSource =  "Data Source=" & ExcelPath & ";"
     'dbExtended =    "Extended Properties=""Excel 8.0;HDR=" & HDRValue & ";IMEX=1" & ";"""
     dbExtended =    "Extended Properties=""Excel 12.0 XML;HDR=" & HDRValue & ";IMEX=1" & ";"""

     objConnection.Open dbProvider & dbSource & dbExtended
     If Err.Number <> 0 Then   
        strError = Err.description
		Reporter.ReportEvent micFail,"Connect to Excel File","Failed in connecting to the Excel - " & ExcelPath & VbCrLf & strError
        ConnectToExcel = ""
        Set objConnection = Nothing    
        Exit Function
     End If

     Set ConnectToExcel = objConnection

End Function

Public Function GetExecutionTimeInMillis(StartTime,EndTime)

	GetExecutionTime = (EndTime - StartTime)

End Function

Function ExeSQL(DBConnection, strSQL, objRS, blnOneRec, numRecCnt)
	On Error Resume Next
    	ExeSQL = False
    	objRS.Open strSQL, DBConnection, adOpenStatic, adLockOptimistic, adCmdText
    	If Err.Number <> 0  Then   
    		strError =  "Error while executing the SQL - " & strSQL & VbCrLf & Err.Description
		objRS.Close: Set objRS = nothing
		Exit Function
     	End if
   	numRecCnt = objRS.RecordCount
   	If Err.Number <> 0  Then   
	   	numRecCnt = 0
   	End if
	If numRecCnt < 1 Then
			If blnOneRec = True Then
				  strError =  "No records found: SQL - " & strSQL & VbCrLf & Err.Description
				  objRS.Close: Set objRS = nothing
				  Exit Function
			End If 
	Else
		objRS.MoveFirst
		If Err.Number <> 0  Then   
				strError =  "Error while performing move method on the recordset. SQL - " & strSQL1 & VbCrLf & Err.Description
				objRS.Close: Set objRS = nothing
				Exit Function
		   End If 
	End If

   ExeSQL = True

End Function



'==================================================
'* Function Name -					fnSynchroniseObjects
'* Function Description -			This function synchronizes objects
'* Created By -									
'* Created Date -										
'* Input Parameter -				objPage [Broswer.Page Object]
'									objObjectToSynchronise [Object]
'									intWaitValut [Any positive wait time]
'* Output Parameter -								
'* Pre-Conditions -										
'* Post Conditions -									
'==================================================

Public Function fnSynchroniseObjects(objPage,objObjectToSynchronise,intWaitValue)
	Dim blnResult
	blnWaitProprtyValue=CBool(blnWaitProprtyValue)
	intWaitValue=(Int(intWaitValue)*1000)
	objPage.Sync
	If err.number>0 Then
		err.clear
		Exit Function
	End If
	objObjectToSynchronise.WaitProperty "visible", True , intWaitValue
	If err.number>0 Then
		err.clear
		Exit Function
	End If
	objObjectToSynchronise.WaitProperty "attribute/readyState", "complete", intWaitValue
	If err.number>0 Then
		err.clear
		Exit Function
	End If
End Function

'==================================================

Public Function UdfLinkClick(objName)

	'Check if the objName exists
	If objName.Exist Then
		If objName.GetROProperty("disabled") = False Then
			objName.Click
			UdfLinkClick = UDTRUE
		Else
'			UdsLogMsg "UdfLinkClick() WebLink Object ["& objName.GetROProperty("name") &"] is DISABLED "
			UdfLinkClick = OBJ_DISABLED
		End If
	Else
'		UdsLogMsg "UdfLinkClick() WebLink Object DOES NOT EXIST "
		UdfLinkClick = OBJ_NOT_FOUND
	End If
End Function

'==================================================

Public Function UdfButtonClick(objName)

	'Check if the objName exists
	If objName.Exist Then
		If objName.GetROProperty("disabled") = False Then
			objName.Click
			UdfButtonClick = UDTRUE
		Else
			'UdsLogMsg "UdfButtonClick() WebButton Object ["& objName.GetROProperty("name") &"] is DISABLED "
			UdfButtonClick = OBJ_DISABLED
		End If
	Else
		'UdsLogMsg "UdfButtonClick() WebButton Object DOES NOT EXIST "
		UdfButtonClick = OBJ_NOT_FOUND
	End If
End Function

'==================================================

Public Function UdfEditSetSecure(objName, strValue)

	'Check if the objName exists
	strobjName = objName
	If Eval(objName & ".Exist") Then
		Set objName = Eval(objName)
		If objName.GetROProperty("disabled") = False Then
			objName.SetSecure strValue
			UdfEditSetSecure = UDTRUE
			stepErrDescription = ""
			stepActual = "Password entered successfully in " & strobjName
			stepResult = "Passed"
			'Eval(objName &".Highlight")
		Else
			'UdsLogMsg "UdfEditSetSecure() WebEdit Object ["& objName.GetROProperty("name") &"] is DISABLED "
			UdfEditSetSecure = OBJ_DISABLED
			stepErrDescription = strobjName & " is Disabled"
			stepResult = "Failed"
		End If
	Else
		'UdsLogMsg "UdfEditSetSecure() WebEdit Object DOES NOT EXIST "
		UdfEditSetSecure = OBJ_NOT_FOUND
		stepErrDescription = strobjName & " Not Found"
		stepResult = "Failed"
	End If
End Function

'==================================================

Public Function UdfEditSet(objName, strValue)
	'Check if the objName exists
	strobjName = objName
	objName = eval(objName)
	If objName.Exist Then
		If objName.GetROProperty("disabled") = False Then
			objName.Set strValue
			UdfEditSet = UDTRUE
			stepErrDescription = ""
			stepActual = "Value entered successfully in " & strobjName
			stepResult = "Passed"
		Else
			UdfEditSet = OBJ_DISABLED
			stepErrDescription = strobjName & " - Object is disabled "
			stepResult = "Failed"
		End If
	Else
		UdfEditSet = OBJ_NOT_FOUND
		stepErrDescription = strobjName & " - Object not found "
		stepResult = "Failed"
	End If
End Function

'==================================================

Function RetrieveProperty_Function()
	'Usage 1:
	'OBJECT                                | ACTION             | PARAMETERS
	'tblObject            | VALIDATE         | strRowWithCellText, intColumnToValidate, strExpectedTxt
	'Usage 2:
	'OBJECT                                | ACTION             | PARAMETERS
	'tblObject            | VALIDATE         | intRow, intColumnToValidate, strExpectedTxt
	'Usage 3:
	'OBJECT                                | ACTION             | PARAMETERS
	'tblObject            | VALIDATE         | intRow, intColumnToValidate, %%Variable

	'NON Web Table object
	'Usage 4:
	'OBJECT                                | ACTION             | PARAMETERS
	'tblObject            | VALIDATE         | propertyName=propertyExpectedValue
	On Error Resume Next
	Err.Clear
	
	If Not IsObject(eval(strObjectName)) Then
		stepErrDescription = UCASE(strAction) & " - " & "Object " & strObjectName & " not found"
		StepResult = "Failed"
		Exit Function
	End If
	If  Not VerifyObjExist(strObjectName)  Then
		stepErrDescription = UCASE(strAction) & " - " & "Object " & strObjectName & " not present "
		StepResult = "Failed"
		Exit Function
	End If    

	Dim strobjClass,SplitParameters,strRowWithCellText,intColumnToValidate,strPropertyValue_Actual
	Execute (strObjectName & ".Init")
	strobjClass = eval(strObjectName & ".GetROProperty(" & Chr(34) & "micClass" & Chr(34) & ")")

	Select Case strobjClass
	
	Case "WebTable"
	
		SplitParameters = Split(strParameters, ",")
		strRowWithCellText = SplitParameters(0) 
		intColumnToValidate = SplitParameters(1)
		strRowWithCellText = GetParameterValue(strRowWithCellText)
		Select Case isNumeric ( Trim(strRowWithCellText))
			Case True
				intTblRow = strRowWithCellText
			Case False
				intTblRow = eval(strObjectName & ".GetRowWithCellText(" & Chr(34) & strRowWithCellText & Chr(34) & ")")
		End Select
		strPropertyValue_Actual = eval(strObjectName & ".GetCellData(" & intTblRow & ", " & intColumnToValidate & ")")
	
	Case "SiebList"
		SplitParameters = Split(strParameters, ",")
		strRowWithCellText = SplitParameters(0) 
		intColumnToValidate = SplitParameters(1)
		Select Case isNumeric ( Trim(strRowWithCellText))
			Case True
				intTblRow = strRowWithCellText
			'Usage is strRowWithCellText, column, strText
			Case False
				intTblRow = eval(strObjectName & ".GetRowWithCellText(" & Chr(34) & strRowWithCellText & Chr(34) & ")")
		End Select
	
		Execute("strPropertyValue_Actual = " & strObjectName & ".GetCellText(" & intColumnToValidate & ", " & intTblRow & ")")
	
	Case Else
	
		strPropertyName = strParameters
		strPropertyValue_Actual = eval(strObjectName & ".GetROProperty(" & Chr(34) & strPropertyName & Chr(34) & ")")
	End Select

	If Err.Number <> 0 Then
				stepErrDescription = UCASE(strAction) & " - " & "RetrieveProperty_Function - Failed to retrieve the property [" & strObjectName & "] from the object [" & testObjectName & "]." & Err.Description
				StepResult = "Failed"
				Exit Function
	Else
				stepActual= "Actual value retrieved from the field: [" & strObjectName & "] is " & strPropertyValue_Actual & "."
				StepResult = "Passed"
				stepErrDescription = ""
	End If
									
	testparameters.Add strObjectName,strPropertyValue_Actual
	'Adds variable name and value to data sheet
	DataTable.GetSheet("Action1").AddParameter strObjectName,strPropertyValue_Actual
	'Call StoreFile_Function(strObjectName,"xml")
	'stepErrDescription = ""
	'StepResult = "Passed"

End Function

'==================================================

Function Clickanylink_Function()

	On Error Resume Next
	Err.Clear
	
	If Not IsObject(eval(strObjectName)) Then
		'"Object " & strTestObject & "   Not Declared"
		stepErrDescription = UCASE(strAction) & " - " & "Object " & strObjectName & " not found"
		StepResult = "Failed"
		Exit Function
	End If

	'If  Not Eval(strObjectName & ".Exist(5)") Then
    If  Not VerifyObjExist(strObjectName)   Then
		stepErrDescription = UCASE(strAction) & " - " & "Object " & strObjectName & " not present "
		StepResult = "Failed"
		Exit Function
	End If    
	
	Execute (strObjectName & ".Init")
	
	Dim astrSingleParam,LinkObj
	
	If instr(strParameters,",") Then
		astrSingleParam = Split(strParameters, ",")
																																		
		Set LinkObj=Description.Create()
		LinkObj("html tag").Value="A"
		LinkObj("text").Value=astrSingleParam(0)
		LinkObj("Index").Value=astrSingleParam(1)
	
	Else
	
		Set LinkObj=Description.Create()
		LinkObj("html tag").Value="A"
		LinkObj("text").Value=strParameters
		LinkObj("Index").Value=0
	
	End If


	ActualLink = strObjectName & ".Link(LinkObj)"
	Execute ( ActualLink & ".Click")
	Set LinkObj=Nothing      

	If Err.Number <> 0 Then
		stepErrDescription = UCASE(strAction) & " - " & Err.Description
		StepResult = "Failed"
		Exit Function
	End If
									
	'stepErrDescription = ""
	'StepResult = "Passed"
End Function

'==================================================

Public Function UserScriptFunction()
	
	On Error Resume Next
	Err.Clear
	
'	If  InStr(1, Trim(strParameters), "=") > 0 Then
'			SplitParameters = Split(strParameters, "=")
'			Execute Trim("ReturnValue = " & SplitParameters(1))
'			inputparameters.Add SplitParameters(0), ReturnValue
	'Else
			Execute Trim(strParameters)
	'End If
	
	'stepExpected = "Execute userscript:" 	& Trim(strParameters)
	If Err.Number <> 0 Then
				stepErrDescription = UCASE(strAction) & " - [" & strParameters & "] is failed " & Err.Description
				Err.Clear
				stepResult ="Failed"
	Else
		    stepResult ="Passed"
			stepActual = " The  Steps [" & strParameters & "] are performed ."
	End If

End Function

'==================================================

Public Function Input_Function()

   'Usage: 1
	'OBJECT		| ACTION	| PARAMETERS
	'tblObject	| Input	| strExpectedRowText, intTblTxtCol, intTblObjCol, objClass, intObjIndex, strValue
	'the below selects the row containing text KUMAR FROM table at col 3 and checks(ON) check box 
	'in the same row at col 1,(with index 0) from tblToList 
	'Screen		Object		| Action	| Parameters
	'			tblToList	| Input		| KUMAR ,3,1,WebCheckBox,0,ON
	'in the place of 'KUMAR ' actual row number can also be given where in, the next param
	'col should be skipped with a blank comma
	'Usage: 2
	'OBJECT		| ACTION	| PARAMETERS
	'NONtblObject	| Input	| strValue

	'Usage:3 
	'Web List box, selection by Index
	'OBJECT			| ACTION	| PARAMETERS
	'lstTestbyindex | INPUT 	| #3

	On Error Resume Next
	Err.Clear

	Dim SplitParameters

	Execute (strObjectName & ".Init")

	strObjectClass = Eval(strObjectName & ".GetROProperty(" & Chr(34) & "micClass" & Chr(34) & ")")
	strObjectClass = UCase(Trim(strObjectClass))

	Select Case strObjectClass
		
		Case "WEBTABLE"
		
			SplitParameters = Split(strParameters, ",")
		
			strExpectedRowText 	= SplitParameters(0)
			intTblTxtCol 		= SplitParameters(1)
			intTblObjCol 		= SplitParameters(2)
			objClass 			= SplitParameters(3)
			intObjIndex 		= SplitParameters(4)
			strValue			= SplitParameters(5)
		
			strExpectedRowText = GetParameterValue(Trim(strExpectedRowText))
			strValue = GetParameterValue(Trim(strValue))
			'strValue = Chr(34) & strValue & Chr(34)
			'The below call will set the reference in object variable "object"
		
			Call objFromTable ( Eval(strObjectName), strExpectedRowText, intTblTxtCol, intTblObjCol, objClass, intObjIndex ) 
			Call SetObjectValue ("object",  strValue )
		
		Case Else
			strValue = GetParameterValue(strParameters)
			'Below is commented out to stop using literal mouse position
			'Setting.WebPackage("ReplayType") = 2
			Call SetObjectValue (strObjectName, strValue)
			'Setting.WebPackage("ReplayType") = 1

	End Select


End Function

'==================================================

Public Function Click_Function
   
			'Usage: 1
			'OBJECT		| ACTION	| PARAMETERS
			'tblObject	| CLICK	| strExpectedRowText, intTblTxtCol, intTblObjCol, objClass, intObjIndex, strValue

			'The below lines clicks the radio button (indexed 0) at column 1 in webtable tblObjectXmlFrame 
			'at row containing text 'Personal Email' at col 4
			'And similarly for the radio button at row containing text 'Work Email'

			'OBJECT				ACTION			PARAMETERS
			'tblObjectXmlFrame	Operation		Personal Email ,4,1,WebRadioGroup,0,Click
			'tblObjectXmlFrame	Operation		Work Email ,4,1,WebRadioGroup,0,Click

			'Usage: 2
			'OBJECT			| ACTION	| PARAMETERS
			'NONtblObject	| CLICK	

	On Error Resume Next
	
	Execute (strObjectName & ".Init")
	Dim SplitParameters
	
	strObjectClass = Eval(strObjectName & ".GetROProperty(" & Chr(34) & "micClass" & Chr(34) & ")")
	strObjectClass = Ucase(Trim(strObjectClass))
	
	Select Case strObjectClass
		
		Case "WEBTABLE"
		
			SplitParameters = Split(strParameters, ",")
			strExpectedRowText 	= SplitParameters(0)
			intTblTxtCol 		= SplitParameters(1)
			intTblObjCol 		= SplitParameters(2)
			objClass 			= SplitParameters(3)
			intObjIndex 		= SplitParameters(4)
			strValue			= SplitParameters(5)
		
			strExpectedRowText = GetParameterValue(Trim(strExpectedRowText))
			strValue = GetParameterValue(Trim(strValue))
		
			Call objFromTable ( Eval(strObjectName), strExpectedRowText, intTblTxtCol, intTblObjCol, objClass, intObjIndex )
			Call PerformObjectOperation (intLoop, "object",  "Click" )
		
		Case Else
		
			Call PerformObjectOperation (strObjectName, "Click")
	
	End Select

End Function

'==================================================

Public Function Operation_Function

	'Usage: 1
	'OBJECT		| ACTION	| PARAMETERS
	'tblObject	| OPERATION	| strExpectedRowText, intTblTxtCol, intTblObjCol, objClass, intObjIndex, strValue

	'Usage: 2
	'OBJECT			| ACTION	| PARAMETERS
	'NONtblObject	| OPERATION	| strValue
	Dim SplitParameters

	If Not IsObject(eval(strObjectName)) Then
		'"Object " & strTestObject & "   Not Declared"
		stepErrDescription =  UCASE(strAction) & " - " & "Object " & strObjectName & " not found"
		StepResult = "Failed"
		Exit Function
	End If

	'If  Not Eval(strObjectName & ".Exist(5)") Then
    If  Not VerifyObjExist(strObjectName)    Then
		stepErrDescription =  UCASE(strAction) & " - " & "Object " & strObjectName & " not present "
		StepResult = "Failed"
		Exit Function
	End If

	Execute (strObjectName & ".Init")
	strObjectClass = Eval(strObjectName & ".GetROProperty(" & Chr(34) & "micClass" & Chr(34) & ")")
	strObjectClass = Ucase(Trim (strObjectClass))

	Select Case strObjectClass
		
		Case "WEBTABLE"
			SplitParameters = Split(strParameters, ",")
			astrSingleParam = Split(strParameters, ",")
			strExpectedRowText 	= SplitParameters(0)
			intTblTxtCol 		= SplitParameters(1)
			intTblObjCol 		= SplitParameters(2)
			objClass 			= SplitParameters(3)
			intObjIndex 		= SplitParameters(4)
			strSingleOperation	= SplitParameters(5)
		
			strExpectedRowText = GetParameterValue(Trim(strExpectedRowText))
			strSingleOperation = GetParameterValue(Trim(strSingleOperation))
		
			Call objFromTable ( Eval(strObjectName), strExpectedRowText, intTblTxtCol, intTblObjCol, objClass, intObjIndex )
			Call PerformObjectOperation ("object",  strSingleOperation )
		
		Case Else
		
			Call PerformObjectOperation (strObjectName, strParameters)

	End Select
		
End Function

'==================================================

Public Function Exist_Function

	On Error Resume Next
	
	If Not IsObject(eval(strObjectName)) Then
		'"Object " & strTestObject & "   Not Declared"
		stepErrDescription =  UCASE(strAction) & " - " & "Object " & strObjectName & " not found"
		StepResult = "Failed"
		Exit Function
	End If
	
	If Len ( trim(strParameters) )  < 1 Then 
	strParameters = "0"		
	End If
	
	If  Eval(strObjectName & ".Exist(" & strParameters & ")") Then
		stepResult ="Passed"
		stepActual = " The Object [" & strObjectName & "] Exists."
	Else
		stepErrDescription = UCASE(strAction) & " - " & "Object " & strObjectName & " not present "
		StepResult = "Failed"
	End If	
	'Eval(strObjectName & ".Highlight")
	If Err.Number <> 0 Then
		strErrDescription = Err.Description
		StepResult = "Failed"
	End If
		
End Function

'==================================================

Public Function Wait_Function

	On Error Resume Next
	Err.Clear

	If  strParameters = "" Then
		strParameters = 0
	End If

	Wait CInt(strParameters)
	If Err.Number <> 0 Then
		strErrDescription = Err.Description
		StepResult = "Failed"
	Else
		stepActual = "Waited for " & strParameters & " seconds"
		StepResult = "Passed"
	End If
	
End Function

'==================================================

Public Function WaitUntil_Function()
	
	'Object		Action		OnError				Parameters
	'fromPort	WaitUntil	NextTestCase		value=London,2	'time in seconds, default is 1 sec
	
	Dim SpitParameters
	SpitParameters = Split(strParameters, "=")
	strPropertyName = SpitParameters(0)
	astrSingleParamExpValAndTimeOut = Split ( SpitParameters(1) , "," )
	strPropertyvalue_expected = astrSingleParamExpValAndTimeOut(0)
	
	'Setting optional time outs 
	If UBound(astrSingleParamExpValAndTimeOut) > 0 Then
		strTimeOut = astrSingleParamExpValAndTimeOut(1)
	Else 
		strTimeOut = ""
	End If

	On Error Resume Next
	Err.Clear
	
	If Not IsObject(eval(strObjectName)) Then
		'"Object " & strTestObject & "   Not Declared"
		stepErrDescription = UCASE(strAction) & " - " & "  " & strObjectName & " not found"
		StepResult = "Failed"
		Exit Function
	End If
	
	'If  Not Eval(strObjectName & ".Exist(50)") Then
	If  Not VerifyObjExist(strObjectName) Then
		stepErrDescription = UCASE(strAction) & " - " & "Object " & strObjectName & " not present "
		StepResult = "Failed"
'		Exit Function
	End If
	
	Execute (strObjectName & ".Init")   ' equiv to .RefreshObject of QTP 10.0
	
	If Len (strTimeOut) = 0  Then
		'Call without explicit timeout value
		strResult = Eval ( strObjectName & ".WaitProperty (" & Chr(34) & strPropertyName & Chr(34) & ", " & Chr(34) & strPropertyvalue_expected & Chr(34) & ")" )
	Else
		'Call with explicit timeout value
		strResult = Eval ( strObjectName & ".WaitProperty (" & Chr(34) & strPropertyName & Chr(34) & ", " & Chr(34) & strPropertyvalue_expected & Chr(34) & ", " & strTimeOut * 1000 & ")" )
	End If
	
	'stepExpected = "" & strObjectName & "'s property " & strPropertyName & "=" & strPropertyvalue_expected
	
	
	If Not strResult Then	'Pass 
		stepErrDescription = UCASE(strAction) & " - " & strObjectName & "'s property " & strPropertyName & " is NOT Equal to " & strPropertyvalue_expected
		stepResult = "Failed"
		Else
		stepResult = "Passed"
		stepActual =  strObjectName & "'s property " & strPropertyName & " is Equal to " & strPropertyvalue_expected
		stepErrDescription = ""
	End If
	
	'Eval(strObjectName & ".Highlight")
	
	If Err.Number <> 0 Then
		strErrDescription = Err.Description
		StepResult = "Failed"
	End If	
End Function

'==================================================



Sub BuildPath (strFilePath) 

   Dim objFS1
	Set objFS1 = CreateObject("Scripting.FileSystemObject")
	If Not (objFS1.FolderExists(strFilePath)) Then
		BuildPath 	objFS1.GetParentFolderName(strFilePath)
		objFS1.CreateFolder strFilePath
	End If
	Set objFS1 = Nothing

End Sub

'==================================================

Public Function ClearCache_CmdLaunch(ByVal strString)
                REM SystemUtil.Run "cmd.exe", "", "C:\Users", "open"
                REM Do until Window( "object class:=ConsoleWindowClass" ).Exist(0)
                                REM wait(1)
                REM Loop

                REM Window( "object class:=ConsoleWindowClass" ).Type strString
                REM Window( "object class:=ConsoleWindowClass" ).Type micReturn
				REM Do Until (Right(Trim(Window( "object class:=ConsoleWindowClass" ).GetVisibleText), 6) = "Users>" )
								REM wait (1)
				REM Loop
				REM Window( "object class:=ConsoleWindowClass" ).Close                   
				SystemUtil.Run "cmd.exe", "", "C:\Users", "open"
                Do until Window( "object class:=ConsoleWindowClass" ).Exist(0)
                wait(1)
                Loop
                Window( "object class:=ConsoleWindowClass" ).Type strString
                Window("object class:=ConsoleWindowClass" ).Type micReturn
                Str =  Window( "object class:=ConsoleWindowClass" ).GetVisibleText
                    If  (Instr(Str, "recognized")=0) Then
                        Reporter.ReportEvent micpass, "Command executed successfully",""
                    else 
                        Reporter.ReportEvent micfail, "Command execution failed",""
               End If
             Window( "object class:=ConsoleWindowClass" ).Close  


End Function

'==================================================

Public Function CloseAllBrowsersExceptQc ( )

	Dim oDesc, colBrowser,intBrowserIndex

	On Error Resume Next
	'Creates a description object
	Set oDesc = Description.Create
	oDesc( "micclass" ).Value = "Browser"

	Set colBrowser = Desktop.ChildObjects(oDesc)

	If colBrowser.Count > 0 Then 
		For intBrowserIndex = 0 to colBrowser.Count - 1
            strCurrBrowTitle = ""
			strCurrBrowTitle = colBrowser(intBrowserIndex).GetRoProperty("title")
			'If Instr(strCurrBrowTitle, "HP Quality Center 10.00") = 0 Then
			
			If (Instr(strCurrBrowTitle, "HP Quality Center 10.00") = 0) And (Instr(strCurrBrowTitle, "Office Preload ") = 0) And (Trim(strCurrBrowTitle) <> "") Then
				colBrowser(intBrowserIndex).Close
			End If
		Next
	End If


	If Err.Number <> 0 Then   
		stepErrDescription = Err.Description
		Err.Clear 
		stepResult = "Failed"
	End If 

	Set colBrowser = Nothing
	Set  oDesc = Nothing

'	Call ClearCache_CmdLaunch ("rundll32 dfshim CleanOnlineAppCache")

Call fnClearTempFilesCacheCookies()

End Function
'==================================================

Function SendKeys (strParameters)		

	Const VK_ESCAPE 	= 1
	Const VK_CONTROL 	= 29
	Const VK_MULTIPLY 	= 55
	Const VK_O		 	= 24
	Const VK_SHIFT	 	= 42
	Const VK_RETURN	 	= 28
	Const VK_TAB        = 15
	

	strParameters = Trim (strParameters)	
	Set WshShell = CreateObject("WScript.Shell")

	Select Case UCase(strParameters)

		Case "ENTER"
					WshShell.SendKeys "{ENTER}"
			
		Case "F1"
					WshShell.SendKeys "{F1}"
			
		Case "F2"
					WshShell.SendKeys "{F2}"
			
		Case "F3"
					WshShell.SendKeys "{F3}"
			
		Case "F4"
					WshShell.SendKeys "{F4}"
			
		Case "F5"
					WshShell.SendKeys "{F5}"
			
		Case "F6"
					WshShell.SendKeys "{F6}"
			
		Case "F7"
					WshShell.SendKeys "{F7}"
			
		Case "F8"
					WshShell.SendKeys "{F8}"
			
		Case "F9"
					WshShell.SendKeys "{F9}"
			
		Case "F10"
					WshShell.SendKeys "{F10}"
			
		Case "F11"
					WshShell.SendKeys "{F11}"
			
		Case "F12"
					WshShell.SendKeys "{F12}"
			
		Case "ESCAPE" 
					WshShell.SendKeys "{ESCAPE}"
			
		Case "TAB" 
					WshShell.SendKeys "{TAB}"
			
		Case "DELETE"
					WshShell.SendKeys "{DEL}"
			
		Case "END"
					WshShell.SendKeys "{END}"

		Case "Screen_Maximize"
		
                    WshShell.SendKeys "%"
					WshShell.SendKeys " "
					wait 2
					WshShell.SendKeys "x"
            
		Case "VK_MULTIPLY" 	'NUMPAD * 
			Set oDeviceReplay = CreateObject("Mercury.DeviceReplay")
			oDeviceReplay.PressKey (VK_MULTIPLY )
			Set oDeviceReplay = Nothing
            
		Case "VK_ESCAPE"
			Set oDeviceReplay = CreateObject("Mercury.DeviceReplay")
			oDeviceReplay.PressKey (VK_ESCAPE )
			Set oDeviceReplay = Nothing
            
		Case "VK_RETURN"	
			Set oDeviceReplay = CreateObject("Mercury.DeviceReplay")
			oDeviceReplay.PressKey (VK_RETURN )
			Set oDeviceReplay = Nothing
			
		Case "VK_TAB"	
			Set oDeviceReplay = CreateObject("Mercury.DeviceReplay")
			oDeviceReplay.PressKey (VK_TAB )
			Set oDeviceReplay = Nothing
			
		Case "BATCH"
			WshShell.SendKeys batchNumber
			
		Case "ORDER"
			WshShell.SendKeys orderNumber
            			
		Case Else
			WshShell.SendKeys strParameters	
    End Select

		If Err.Number <> 0 Then
				stepErrDescription = UCASE(strAction) & " - " & strParameters& "-"&" Sendkey  action is not perfomed . "
				StepResult = "Failed"
				Exit Function
		Else
				stepActual= strParameters& "-"&" Sendkey  action is perfomed successfully . "
				StepResult = "Passed"
				stepErrDescription = ""
		End If
					

	Set WshShell=Nothing
'	wait 2
	wait conTwo
End Function

'==================================================


Public Function SendKeys_Function
	If isNull(strObjectName) or isEmpty(strObjectName) Then
		strObjectName =""
	End If
	If Len(Trim(strObjectName)) > 0 Then

		strObjectClass = Eval(strObjectName & ".GetROProperty(" & Chr(34) & "micClass" & Chr(34) & ")")							

		Select Case Trim(strObjectClass)

			Case "TeScreen"	'PCOM / Terminal Emulator Screen
				'Execute (strObjectName & ".Sendkey " & Chr(34) & "micClass" & Chr(34) & ") strParameters)
				Execute (strObjectName & ".Sendkey " & "TE_" & strParameters )
				Execute (strObjectName & ".Sync " )
				stepActual = strObjectClass
				'Eval(strObjectName & ".Highlight")
        Case Else
					'
		End Select
	Else
				Call SendKeys (strParameters)	

	End If

End Function
'==================================================

'* Function Name -										intGetRowContainingTextAtGivenColumn
'* Function Description -							This function returns the row number of the web table that has the required text at a given column. It is used by the "INPUT", "CLICLK", "VALIDATE", "VALREGEXP" action keywords.
'* Created By -												
'* Created Date -										
'* Input Parameter -									
'* Output Parameter -								Row number of the web table object.
'* Pre-Conditions -										
'* Post Conditions -									
'==================================================

Function intGetRowContainingTextAtGivenColumn ( objTbl, strExpectedRowText, intTblTxtCol )

	intGetRowContainingTextAtGivenColumn = -1
	For intIndexRowCount = 1 to objTbl.RowCount
		strRcActualText = objTbl.GetCellData( intIndexRowCount, intTblTxtCol )
		If Ucase (Trim ( strRcActualText) ) = Ucase (Trim ( strExpectedRowText )) Then
			intGetRowContainingTextAtGivenColumn = intIndexRowCount
			Exit Function
		End If
	Next

End Function

'==================================================

'==================================================
'* Function Name -										objFromTable
'* Function Description -							This function returns the child object of the web table in the OBJECT column used by the keywords "INPUT", "CLICK", "VALIDATE", "VALREGEXP"
'* Created By -												
'* Created Date -										
'* Input Parameter -									
'* Output Parameter -								Child object of web table for the class type spcified in alias string"object"
'* Pre-Conditions -										
'* Post Conditions -									
'==================================================

Function objFromTable ( objTbl, strExpectedRowText, intTblTxtCol, intTblObjCol, objClass, intObjIndex ) 

	If isNumeric( Trim (strExpectedRowText) ) Then
		Set object = objTbl.ChildItem (strExpectedRowText, intTblObjCol, objClass, intObjIndex)	'.Click
	Else
		intTblRow = intGetRowContainingTextAtGivenColumn (objTbl, strExpectedRowText, intTblTxtCol )
		Set object = objTbl.ChildItem (intTblRow , intTblObjCol, objClass, intObjIndex)
	End If

End Function

'==================================================
'* Function Name -			PrintToReport
'* Function Description -	This function is used by the "PRINT" keyword. It is used to output any messages in both QTP and HTMLresult.
'* Created By -									
'* Created Date -								
'* Input Parameter -									
'* Output Parameter -								
'* Pre-Conditions -										
'* Post Conditions -									
'==================================================

Function PrintToReport(strObjAlias, strParameters)

	Dim SplitParameters,strMsgToBePrinted,strPropertyToBePrinted, strValueToBePrinted
	Err.Clear

	SplitParameters = Split(strParameters, ",", 2)
	strMsgToBePrinted = SplitParameters(0)
	strPropertyToBePrinted	= Trim(SplitParameters(1))

	strProperty = GetParameterValue(strPropertyToBePrinted)
	strPropertyToBePrinted = Chr(34) & GetParameterValue(strPropertyToBePrinted) & Chr(34)
	'strPropertyToBePrinted = GetParameterValue(strPropertyToBePrinted)

	If Len(Trim(strObjAlias)) > 0  Then
		strValueToBePrinted = eval(strObjAlias & ".GetROProperty(" & strPropertyToBePrinted & ")" )	'09-Aug-10
		If Err.Number <> 0 Then
				stepErrDescription =  strParameters  & " is failed " & Err.Description
				stepResult = "Failed"
				Exit Function
		End If
		'stepErrDescription = " Object : " & strObjAlias & "'s "  &  " value is  " & strValueToBePrinted
	'	 stepErrDescription = UCASE(strAction) & " - " & "Actual Result : "& strMsgToBePrinted & " :- " & strValueToBePrinted
		'stepErrDescription = "Print : Object " & strObjAlias & "'s "  & strPropertyToBePrinted & " value is  " & strValueToBePrinted
		'stepErrDescription = strPropertyToBePrinted & " value is : " & strValueToBePrinted

		stepActual = "Actual Result : "& strMsgToBePrinted & " :- " & strValueToBePrinted
		stepResult = "Passed"	
	
	Else
		'stepErrDescription = "Print : " & strMsgToBePrinted & " " & GetParameterValue(strPropertyToBePrinted)
		stepActual = "Actual Result : " & strMsgToBePrinted &" :- " & GetParameterValue(strPropertyToBePrinted)
'		stepActual = Replace(stepErrDescription,Chr(34),"")
		stepResult = "Passed"

	End If
			
End Function

'==================================================
'* Function Name -										ValidateByRegExpObjectsProperty
'* Function Description -							This function validates the test object's property name & value pairs. It is similar to "ValidateObjectsProperty" except that here the expected text can have regular expression patterns.
'* Created By -	
'* Created Date -	
'* Input Parameter -									intStepNo [Contains the step from the input sheet]
'																		strItemName [Contains the screen name followed after  Validate_screen]
'																		strParameters [Contains the property name &value pairs]
'																		NOTE - Currently, this function does not validate an object in a table but validates the text in a web table cell.
'* Output Parameter -								
'* Pre-Conditions -										Application should be opened and the object is found in the current page.
'* Post Conditions -									Outputs a pass/fail status in both the QTPReporter.ReportEvent and the HTMLresult file.
'==================================================

Function ValidateByRegExpObjectsProperty(strObjAlias, strParameters)

	'Usage 1:
	'OBJECT		| ACTION 	| PARAMETERS
	'tblObject	| VALIDATE	| strRowWithCellText, intColumnToValidate, strExpectedTxt
	'Usage 2:
	'OBJECT		| ACTION 	| PARAMETERS
	'tblObject	| VALIDATE	| intRow, intColumnToValidate, strExpectedTxt
	'Usage 3:
	'OBJECT		| ACTION 	| PARAMETERS
	'tblObject	| VALIDATE	| intRow, intColumnToValidate, %%Variable

	Dim regEx, Match, Matches   ' Create variable.

	If Not IsObject(eval(strObjAlias)) Then
		'"Object " & strTestObject & "   Not Declared"
		stepErrDescription = UCASE(strAction) & " - " & "Object " & strObjAlias & " not found"
		StepResult = "Failed"
		Exit Function
	End If

	'If  Not Eval(strObjAlias & ".Exist(5)") Then
 	If  Not  VerifyObjExist(strObjAlias)  Then
		stepErrDescription = UCASE(strAction) & " - " & "Object " & strObjAlias & " not present "
		StepResult = "Failed"
		Exit Function
	End If	

	Execute (strObjAlias & ".Init")

	Dim strobjClass,SplitParameters,strRowWithCellText,intColumnToValidate,strPropertyvalue_expected,strPropertyValue_Actual
	strobjClass = eval(strObjAlias & ".GetROProperty(" & Chr(34) & "micClass" & Chr(34) & ")")

	Select Case strobjClass

	Case "WebTable"

		SplitParameters = Split(strParameters, ",")
		strRowWithCellText = SplitParameters(0) 'dt_ param have to be taken from the parameters sheet
		intColumnToValidate = SplitParameters(1)
		strPropertyvalue_expected =  SplitParameters(2) 	'dt_ param have to be taken from the parameters sheet

		strRowWithCellText = GetParameterValue(strRowWithCellText)
		strPropertyvalue_expected = GetParameterValue(strPropertyvalue_expected)


		Select Case isNumeric ( Trim(strRowWithCellText))
			'Usage is intRow, intColumnToValidate, strExpectedTxt
			Case True
				intTblRow = strRowWithCellText
			'Usage is strRowWithCellText, column, strText
			Case False
				intTblRow = eval(strObjAlias & ".GetRowWithCellText(" & Chr(34) & strRowWithCellText & Chr(34) & ")")
		End Select

		strPropertyValue_Actual = eval(strObjAlias & ".GetCellData(" & intTblRow & ", " & intColumnToValidate & ")")

		Set regEx = New RegExp   'Creates a regular expression.
		regEx.Pattern = strPropertyvalue_expected   'Sets a pattern.
		regEx.IgnoreCase = True   'Sets the case insensitivity.
		regEx.Global = True   'Sets the global applicability.

		Set Matches = regEx.Execute(strPropertyValue_Actual)   'Executes a search.

		'stepExpected = "Table Object " & strObjAlias & "'s  row,col (" & intTblRow & ", " & intColumnToValidate &")= [RegExpMatch] " & 	strPropertyvalue_expected
		If Matches.Count <= 0 Then
			stepErrDescription = UCASE(strAction) & " - " & "Validation_RegEx Failed: Object " & strObjAlias & "'s row,col (" & intTblRow  & ", " & intColumnToValidate &") " & " is not equal to  " & strPropertyvalue_expected
			stepResult = "Failed"
		End If	
		Set regEx = Nothing
		Set Matches = Nothing

	Case Else

		SplitParameters = Split(strParameters, "=")
		strPropertyName = SplitParameters(0)
		strPropertyvalue_expected = SplitParameters(1)
	
		Set regEx = New RegExp   'Creates a regular expression.
		regEx.Pattern = strPropertyvalue_expected   'Sets a pattern.
		regEx.IgnoreCase = True   'Sets the case insensitivity.
		regEx.Global = True   'Sets the global applicability.

		strPropertyValue_Actual = eval(strObjAlias & ".GetROProperty(" & Chr(34) & strPropertyName & Chr(34) & ")")
		Set Matches = regEx.Execute(strPropertyValue_Actual)   ' Execute search.

		'stepExpected = " Object " & strObjAlias & "'s  property " & strPropertyName & "= [RegExpMatch] " & strPropertyvalue_expected
		If Matches.Count <= 0 Then
			stepErrDescription = UCASE(strAction) & " - " & "Validation_RegEx Failed: Object " & strObjAlias & "'s property " & strPropertyName & " is not equal to " & strPropertyvalue_expected
			stepResult = "Failed"
		End If	
		Set regEx = Nothing
		Set Matches = Nothing

	End Select


   'RegExpTest = RetStr
End Function

'==================================================

'* Function Name -										ValidateObjectsProperty
'* Function Description -							This function validates the test object's property name & value pairs.
'* Created By -										
'* Created Date -									
'* Input Parameter -									intStepNo [Contains the step from the input sheet]
'																		strItemName [Contains the screen name followed after  Validate_screen]
'																		strParameters [Contains the property name &value pairs]
'																		NOTE - Currently, this function does not validate an object in a table but validates the text in a web table cell.
'* Output Parameter -								
'* Pre-Conditions -										Application should be opened and the object is found in the current page.
'* Post Conditions -									Outputs a pass/fail status in both the QTPReporter.ReportEvent and the HTMLresult file.
'==================================================

Function ValidateObjectsProperty(strObjAlias, strParameters)

	'Usage 1:
	'OBJECT		| ACTION 	| PARAMETERS
	'tblObject	| VALIDATE	| strRowWithCellText, intColumnToValidate, strExpectedTxt
	'Usage 2:
	'OBJECT		| ACTION 	| PARAMETERS
	'tblObject	| VALIDATE	| intRow, intColumnToValidate, strExpectedTxt
	'Usage 3:
	'OBJECT		| ACTION 	| PARAMETERS
	'tblObject	| VALIDATE	| intRow, intColumnToValidate, %%Variable

	'NON Web Table object
	'Usage 4:
	'OBJECT		| ACTION 	| PARAMETERS
	'tblObject	| VALIDATE	| propertyName=propertyExpectedValue

	If Not IsObject(eval(strObjAlias)) Then
		'"Object " & strTestObject & "   Not Declared"
		stepErrDescription =  UCASE(strAction) & " - " & "Object " & strObjAlias & " not found"
		StepResult = "Failed"
		Exit Function
	End If

	'If  Not Eval(strObjAlias & ".Exist(5)") Then
	If  Not VerifyObjExist(strObjAlias) Then
		stepErrDescription =  UCASE(strAction) & " - " & "Object " & strObjAlias & " not Present "
		StepResult = "Failed"
		Exit Function
	End If	

	Dim strobjClass,SplitParameters,strRowWithCellText,intColumnToValidate,strPropertyvalue_expected,strPropertyValue_Actual
	strobjClass = eval(strObjAlias & ".GetROProperty(" & Chr(34) & "micClass" & Chr(34) & ")")

	Select Case strobjClass
   
		Case "WebTable"
		
			SplitParameters = Split(strParameters, ",")
			strRowWithCellText = SplitParameters(0) 'It can have dt_ param in the parameters sheet
			intColumnToValidate = SplitParameters(1)
			strPropertyvalue_expected =  SplitParameters(2) 'It can have dt_ param in the parameters sheet
		
			strRowWithCellText = GetParameterValue(strRowWithCellText)
			strPropertyvalue_expected = GetParameterValue(strPropertyvalue_expected)
		
			Select Case isNumeric ( Trim(strRowWithCellText))
				'Usage is intRow, intColumnToValidate, strExpectedTxt
				Case True
					intTblRow = strRowWithCellText
				'Usage is strRowWithCellText, column, strText
				Case False
					intTblRow = eval(strObjAlias & ".GetRowWithCellText(" & Chr(34) & strRowWithCellText & Chr(34) & ")")
			End Select
		
			strPropertyValue_Actual = eval(strObjAlias & ".GetCellData(" & intTblRow & ", " & intColumnToValidate & ")")
		
			'stepExpected = "Table Object " & strObjAlias & "'s  row,col (" & intTblRow & ", " & intColumnToValidate &  ") = " & strPropertyvalue_expected
		
			If Not UCase(Trim(strPropertyValue_Actual)) = UCase(Trim(strPropertyvalue_expected)) Then
				strErrDescription = "Validation Failed:  Object " & strObjAlias & "'s row,col (" & intTblRow & ", " & intColumnToValidate &  ") is not equal to " & strPropertyvalue_expected
				stepResult = "Failed"
			End If
		
		Case "SiebList"
		
			SplitParameters = Split(strParameters, ",")
			strRowWithCellText = SplitParameters(0) 'It can have dt_ param in the parameters sheet
			intColumnToValidate = SplitParameters(1)
			strPropertyvalue_expected =  SplitParameters(2) 'It can have dt_ param in the parameters sheet
		
			'strRowWithCellText = GetParameterValue(strRowWithCellText)
			strPropertyvalue_expected = GetParameterValue(strPropertyvalue_expected)
		
			Select Case isNumeric ( Trim(strRowWithCellText))
				'Usage is intRow, intColumnToValidate, strExpectedTxt
				Case True
					intTblRow = strRowWithCellText
					'Usage is strRowWithCellText, column, strText
				Case False
					intTblRow = eval(strObjAlias & ".GetRowWithCellText(" & Chr(34) & strRowWithCellText & Chr(34) & ")")
			End Select
		
			Execute("strPropertyValue_Actual = " & strObjAlias & ".GetCellText(" & intColumnToValidate & ", " & intTblRow & ")")
		
			'stepExpected = "List Object " & strObjAlias & "'s  col,row (" & intColumnToValidate & ", " & intTblRow &  ") = " & strPropertyvalue_expected
		
			If Not UCase(Trim(strPropertyValue_Actual)) = UCase(Trim(strPropertyvalue_expected)) Then
				strErrDescription = "Validation Failed:  Object " & strObjAlias & "'s col,row (" & intColumnToValidate & ", " & intTblRow &  ") is not equal to " & strPropertyvalue_expected
				stepResult = "Failed"
			End If
		Case Else
		
			SplitParameters = Split(strParameters, "=")
			strPropertyName = SplitParameters(0)
			strPropertyvalue_expected = SplitParameters(1)
		
			strPropertyvalue_expected = GetParameterValue(strPropertyvalue_expected)
			strPropertyValue_Actual = eval(strObjAlias & ".GetROProperty(" & Chr(34) & strPropertyName & Chr(34) & ")")
		
			'stepExpected = "Object " & strObjAlias & "'s  property " & strPropertyName & "=" & strPropertyvalue_expected
		
			If Not UCase(Trim(strPropertyValue_Actual)) = UCase(Trim(strPropertyvalue_expected)) Then
							stepErrDescription = UCASE(strAction) & " - " & "Validation Failed: Object " & strObjAlias & "'s property " & strPropertyName & " does not contain [" & strPropertyvalue_expected & "].The actual value is [" & strPropertyValue_Actual & "]." 
							stepResult = "Failed"
			Else
							stepActual = "Object " & strObjAlias & "'s property " & strPropertyName & " contains the value [" & strPropertyvalue_expected & "]."
							stepErrDescription = ""
							stepResult = "Passed"
			End If	

	End Select

End Function


Public Function Validate_Function

	strObjectClass = Eval(strObjectName & ".GetROProperty(" & Chr(34) & "micClass" & Chr(34) & ")")

	Select Case Ucase(Trim(strObjectClass))
		Case Else
			Call ValidateObjectsProperty(strObjectName, strParameters)
	End Select

End Function

'==================================================

'* Function Name -										SetObjectValue
'* Function Description -							This function sets a value to the specified object (WebEdit,WebList,WebCheckBox and WebRadioGroup).
'																		NOTE - If the object is classified as a WebTable, then the action is performed on a child item of that web table as per the properties (rowwithtext/rownum, text col, object colum, class, index).
'																						This child item is returned with an object alias name which is retrieved by the "objFromTable()" funciton.
'* Created By -											
'* Created Date -										
'* Input Parameter -									intStepNo, strItemName, strParameters
'* Output Parameter -								
'* Pre-Conditions -										Application should be opened and the object is found in the current page.
'* Post Conditions -									Retrieves the Test Object Name from the strParameters and identifies the object class and sets the value accordingly.
'==================================================

Function SetObjectValue(strTestObject, strObjValue)

	On Error Resume Next
	Err.Clear
	
	'stepExpected = "" & strTestObject & " is set with the Value " & strObjValue
	
	If Not IsObject(eval(strTestObject)) Then
		'"Object " & strTestObject & "   Not Declared"
		stepErrDescription =  UCASE(strAction) & " - " & "Object " & strTestObject & " Not Found"
		StepResult = "Failed"
		Exit Function
	End If
	
	'If  Not Eval(strTestObject & ".Exist(5)") Then
	If  Not VerifyObjExist(strTestObject) Then
	'If VerifyObjExist(strTestObject) Then
		stepErrDescription =  UCASE(strAction) & " - " & "Object " & strTestObject & " Not Present "
		StepResult = "Failed"
		Exit Function
	End If	
	
	strObjectClass = Eval(strTestObject & ".GetROProperty(" & Chr(34) & "micClass" & Chr(34) & ")")
	
	'if the above object is a child item from the table then don't initialize
	If Trim(Ucase(strTestObject)) <> "OBJECT" Then
		Execute (strTestObject & ".Init")   
		'Eval(strTestObject&".Highlight")
		'Equiv of .RefreshObject of QTP 10.0 to reinitialize
		'object in OR in order identify same obj in app

		'After browser is refreshed
	End If
	
	strObjValue = chr(34) & strObjValue & chr(34)
	
	'To handle objects other than Java objects
	
	If Trim ( Len ( strObjectClass ) ) > 0  Then
		Select Case Trim(strObjectClass)
		
			Case "WebEdit", "WinEdit", "PbEdit", "VbEdit", "TeField", "WebFile"
				Execute (strTestObject & ".Set " & strObjValue )
			
			Case "WinEditor","WinObject"
				Execute (strTestObject & ".Type " & strObjValue)
			
			Case "WebList", "WinList", "WinComboBox", "VbComboBox", "SiebPicklist", "OracleList"
				Execute (strTestObject & ".Select " & strObjValue)
			
			Case "WebCheckBox", "PbCheckBox","WinCheckBox", "SblEdit"
				strUCaseObjValue = UCase(strObjValue)
				Execute (strTestObject & ".Set " & strUCaseObjValue)
			
			Case "WebRadioGroup"
				Execute (strTestObject & ".Select " & strObjValue)
			
			Case "WinRadioButton", "PbRadioButton", "VbRadioButton"
				Execute (strTestObject & ".Set ")
			
			Case "SiebText", "SiebCalendar","SiebCalculator"
				Execute (strTestObject & ".SetText" & strObjValue)
										
			Case "OracleTree", "OracleListOfValues"
				Execute (strTestObject & ".Select " & strObjValue)
			
			Case "OracleCheckbox"
				Execute (strTestObject & ".Select")
			
			Case "OracleTextField"
				Execute (strTestObject & ".Enter " & strObjValue  )
		
		End Select
	Else 
		'To handle Java objects
		strObjectClass = eval(strTestObject & ".GetROProperty(" & Chr(34) & "Class Name" & Chr(34) & ")")
		Select Case Trim(strObjectClass)
			
			Case "JavaEdit"
				Execute (strTestObject & ".Set " & strObjValue )
				
			Case "JavaTree"
				Execute (strTestObject & ".Select " & strObjValue)
			
			Case "JavaList"
				Execute (strTestObject & ".Select " & strObjValue )
							
			Case "JavaCheckBox"
				strUCaseObjValue = UCase(strObjValue)
				Execute (strTestObject & ".Set " & strUCaseObjValue )
			
			Case "JavaRadioButton"
				Execute (strTestObject & ".Set")	'JavaRadioButton can only be checked, CANNOT UnCheck
			
			Case "JavaTable"
				Call setjavatable(strTestObject, strObjValue)
			
			Case "JavaTab"
				Execute (strTestObject & ".Select " & strObjValue)
							
		End Select
	
	End If
                
		If Err.Number <> 0 Then
					stepErrDescription = UCASE(strAction) & " - " & "Input_Function - " & strTestObject & " is not set with the value [" & strObjValue & "]." & Err.Description
					StepResult = "Failed"
					Err.Clear
					Exit Function
		Else
					stepActual = "Object with logical name " & strTestObject & " is set with the value [" & strObjValue & "]."
					stepErrDescription = ""
					StepResult = "Passed"
					'Eval(strTestObject&".Highlight")
					Exit Function
		End If


		strValueProperty = "value"
		Select Case Trim(strObjectClass)
			
			Case "WebEdit", "PbEdit"
				strValueProperty = "value"
                'stepErrDescription = "Input : " & MID(strTestObject,4) & " :- " & GetParameterValue(strParameters) &"--"&Environment.value("Temp_strMsgToBePrinted")
				stepErrDescription= UCASE(strAction) & " - " & "Input : " & strParameters & " :- " &GetParameterValue(strParameters)

			Case "WinEdit", "TeField"
				strValueProperty = "text"
			
			Case "WinEditor", "VbEdit"
				strValueProperty = "text"
			
			Case "WebList", "WinList", "PbList"
				strValueProperty = "value"
                'stepErrDescription = " Selected item : " & MID(strTestObject,4) & " :- " & GetParameterValue(strParameters)&"--"&Environment.value("Temp_strMsgToBePrinted")
				stepErrDescription= UCASE(strAction) & " - " & " Selected item : " & strParameters & " :- " &GetParameterValue(strParameters)
			
			Case "WinComboBox", "VbComboBox"
				strValueProperty = "text"	'strValueProperty = "selection"
			
			Case "WebCheckBox"
				strValueProperty = "value"
			
			Case "WebRadioGroup"
				strValueProperty = "value"
			
			Case "WinRadioButton", "PbRadioButton", "PbCheckBox", "WinCheckBox", "VbRadioButton"
				strValueProperty = "checked"
		End Select
		Reporter.Filter = rfDisableAll 				
		Execute (strTestObject & ".RefreshObject " )
		objNewValue = Eval(strTestObject & ".GetROProperty(" & Chr(34) & strValueProperty & Chr(34) & ")")
		'stepActual = "Actual: " & strTestObject & " is set with the value " & "[" & ObjValuevalueRT & "]
		'stepErrDescription = ""
		'StepResult = "Passed"
		'err.clear
		Reporter.Filter = rfEnableAll
		'Eval(strTestObject&".Highlight")
End Function

'==================================================

'* Function Name -										DeleteRunTimeTempFileIfExists
'* Function Description -							This function is used to delete the run time input sheet after test execution is complete.
'* Created By -										
'* Created Date -									
'* Input Parameter -									Runtime_TC.xls [absolute path file name string]
'* Output Parameter -								
'* Pre-Conditions -										Script execution should be complete as of the last keyword from the input sheet.
'* Post Conditions -									Runtime_TC.xls file is deleted from temporary files folder
'==================================================

Function DeleteRunTimeTempFileIfExists ( strAbsolutePathFile )

	Dim filesys
	Set filesys = CreateObject("Scripting.FileSystemObject")
	If filesys.FileExists(strAbsolutePathFile) Then
			filesys.DeleteFile strAbsolutePathFile
	End If
	Set filesys = Nothing

End Function
'==================================================
'* Function Name -										InputParam_Function
'* Function Description -							This function is used to input values into the application.
'* Created By -										
'* Created Date -										
'* Input Parameter -									
'* Output Parameter -								
'* Pre-Conditions -										
'* Post Conditions -									
'==================================================

Public Function InputParam_Function

	' If parameter has = then store the value in the variable defined berore =
	'else store the value in object name
	Dim paramName, paramValue

	paramName = strObjectName
	paramValue = strParameters

	If Instr(1,paramValue , "=") <> 0 Then
		Dim SplitParameters
		SplitParameters = Split(paramValue,"=")
		paramName = SplitParameters(0)
		paramValue = SplitParameters(1)
	End If
	'stepExpected = "Add value " & GetParameterValue(paramValue) & " to the input parameter " & paramName

	inputparameters.Add paramName,GetParameterValue(paramValue)

End Function

'==================================================
'* Function Name -										fnSetBrowserTitile
'* Function Description -							This function is used to set the Browser Title
'* Created By -												
'* Created Date -										
'* Input Parameter -									
'* Output Parameter -								
'* Pre-Conditions -										
'* Post Conditions -									
'==================================================

Function fnSetBrowserTitile(strTitle)
strBrowserTitle = ".*" & strTitle & ".*"
End Function

'==================================================

Public Function GetObjectProperties(objName,screenName)

	On Error Resume Next
	Dim strSQL,rs, propertiesScreen, propertiesObject, objType

	'Gets screen properties
	strSQL = "SELECT Properties FROM UIObjects where ScreenName = '" & screenName & "' And ObjectName = '" & screenName & "'"
	Set rs = DBConnection_Repository.Execute(strSQL)

	If Err.Number <> 0 Then   
			Reporter.ReportEvent micFail, "Executing the SQL", "Error while executing the SQL" & Err.Description 
			Err.Clear 
			GetObjectProperties = false
			Exit Function
	End If 


	If rs.RecordCount = 0 Then
		GetObjectProperties = false
		Exit Function
	End If

	propertiesScreen = rs.Fields.Item("Properties").Value 
	Set rs = nothing

	'Gets object properties
	strSQL = "SELECT ObjectClass,Properties FROM UIObjects where ScreenName = '" & screenName & "' And ObjectName = '" & objName & "'"
	Set rs = DBConnection_Repository.Execute(strSQL)

	If Err.Number <> 0 Then   
			Reporter.ReportEvent micFail, "Executing the SQL to get object properties", "Error while executing the SQL to get object properties" & Err.Description 
			Err.Clear 
			GetObjectProperties = false
			Exit Function
	End If 


	If rs.RecordCount = 0 Then
		GetObjectProperties = false
		Exit Function
	End If

	propertiesObject = rs.Fields.Item("Properties").Value 
	objClass = rs.Fields.Item("ObjectClass").Value 
	Set rs = nothing

	RepositoryObject = propertiesScreen & "." & objClass & "(" & propertiesObject & ")" 
	GetObjectProperties = True

End Function

'==================================================

'==================================================
'* Function Name - fnClearTempFilesCacheCookies
'* Function Description - This function is used to clear User Temp files, System Temp files, IE cookies and cache
'* Created By - 
'* Created Date - 
'* Input Parameter - 
'* Output Parameter - 
'* Pre-Conditions - 
'* Post Conditions - 
'==================================================

Function fnClearTempFilesCacheCookies()
'    Option Explicit
	On Error Resume Next
	Err.Clear 
	Dim objShell, objSysEnv, objUserEnv, strUserTemp, strSysTemp, userProfile, TempInternetFiles, OSType, strIEcookies, strIEcookieLow
	 
	Set objShell=CreateObject("WScript.Shell")  
	Set objSysEnv=objShell.Environment("System") 
	Set objUserEnv=objShell.Environment("User") 
	 
	strUserTemp= objShell.ExpandEnvironmentStrings(objUserEnv("TEMP")) 
	strSysTemp= objShell.ExpandEnvironmentStrings(objSysEnv("TEMP")) 
	userProfile = objShell.ExpandEnvironmentStrings("%userprofile%") 
	
	strIEcookies= userProfile&"\AppData\Roaming\Microsoft\Windows\Cookies"
	
'	DeleteTemp strUserTemp 'delete user temp files  
	DeleteTemp strSysTemp  'delete system temp files 
	DeleteTemp strIEcookies 'delete files from IE cookies folder
	
	strIEcookies= strIEcookies & "\Low"
	DeleteTemp strIEcookies ' delete files from IE cookies Low folder
	
	'delete Internet Temp files 
	'the Internet Temp files path is diffrent according to OS Type 
	OSType=FindOSType 
	
	If OSType="Windows 7" Or OSType="Windows Vista" Then 
		TempInternetFiles=userProfile & "\AppData\Local\Microsoft\Windows\Temporary Internet Files" 
	ElseIf  OSType="Windows 2003" Or OSType="Windows XP" Then 
		TempInternetFiles=userProfile & "\Local Settings\Temporary Internet Files" 
	End If 
	 
	DeleteTemp TempInternetFiles 
	'this is also to delete Content.IE5 in Internet Temp files 
	TempInternetFiles=TempInternetFiles & "\Content.IE5" 
	DeleteTemp TempInternetFiles 
	

	'Clear Browser cache
	Dim objBrowser, openurl
	WebUtil.DeleteCookies
    Err.Clear

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function FindOSType 
		'Defining Variables 
	Dim objWMI, objItem, colItems 
	Dim OSVersion, OSName 
	Dim ComputerName 
		
	ComputerName="." 
	
	'Get the WMI object and query results 
	Set objWMI = GetObject("winmgmts:\\" & ComputerName & "\root\cimv2") 
	Set colItems = objWMI.ExecQuery("Select * from Win32_OperatingSystem",,48) 
	  
	'Get the OS version number (first two) and OS product type (server or desktop)  
	For Each objItem in colItems 
		OSVersion = Left(objItem.Version,3)                  
	Next 
		 
	Select Case OSVersion 
		Case "6.1" 
				OSName = "Windows 7" 
		Case "6.0"  
				OSName = "Windows Vista" 
		Case "5.1"  
				OSName = "Windows XP" 
	End Select 
	  
	'Return the OS name 
	FindOSType = OSName 
	
	'Clear the memory 
	Set colItems = Nothing 
	Set objWMI = Nothing 
End Function
    
'"'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DeleteTemp (strTempPath) 
	On Error Resume Next 
	
	Dim objFSO, objFolder, objDir, objFile, i
	 
	Set objFSO=CreateObject("Scripting.FileSystemObject") 
	Set objFolder=objFSO.GetFolder(strTempPath) 
	 
	'delete all files 
	For Each objFile In objFolder.Files 
		objFile.delete True 
	Next 
	 
	'delete all subfolders 
	For i=0 To 10 
		For Each objDir In objFolder.SubFolders 
			objDir.Delete True 
		Next 
	Next 
	 
	'clear all objects 
	Set objFSO=Nothing 
	Set objFolder=Nothing 
	Set objDir=Nothing 
	Set objFile=Nothing 
End Sub 

'==================================================

Public Function UDFOpenURL()
	stepErrDescription =""
	arrSplitParameters = Split(valParam1, ",")
			     	strBrowser = arrSplitParameters(0)
			     	If not Instr(strBrowser,Chr(34) )>0 Then
			     		strBrowser = GetParameterValue(strBrowser)
			     	Else
			     		strBrowser =TRIM( Replace(strBrowser, Chr(34)  , ""))
			     	End If
			     
			     	strReportURL = arrSplitParameters(1)
			     	If not Instr(strReportURL,Chr(34) )>0 Then
			     		strReportURL = GetParameterValue(strReportURL)
			     	Else
			     		strReportURL =TRIM( Replace(strReportURL, Chr(34)  , ""))
			     	End If
	'Dim strWinID
	CloseAllBrowsersExceptQc()
	Dim objInitialBrowser
	if strBrowser = "ie" then
		SystemUtil.Run "iexplore.exe",""
		wait(2)
		Set objInitialBrowser = Browser("creationtime:=0")
		BrowserVersion = objInitialBrowser.GetAllROProperties().Item(3).Value
		BrowserVersion = GetBrowserVersion(strBrowser)
		objInitialBrowser.Navigate(strReportURL)
		'SystemUtil.Run "iexplore.exe",strReportURL,"",""
	Elseif strBrowser = "chrome" then
		SystemUtil.Run "chrome.exe",""
		wait(5)
		Set objInitialBrowser = Browser("creationtime:=0")
		BrowserVersion = objInitialBrowser.GetAllROProperties().Item(3).Value
		'SystemUtil.Run "chrome.exe",strReportURL,"",""
		objInitialBrowser.Navigate(strReportURL)
	ElseIf strBrowser = "edge" Then
		SystemUtil.Run "msedge.exe",""
		wait(2)
		Set objInitialBrowser = Browser("creationtime:=0")
		BrowserVersion = objInitialBrowser.GetAllROProperties().Item(3).Value
		'SystemUtil.Run "msedge.exe",strReportURL,"",""
		objInitialBrowser.Navigate(strReportURL)
	ElseIf strBrowser = "firefox" Then
		SystemUtil.Run "firefox.exe",""
		wait(2)
		Set objInitialBrowser = Browser("creationtime:=0")
		BrowserVersion = objInitialBrowser.GetAllROProperties().Item(3).Value
		'SystemUtil.Run "firefox.exe",strReportURL,"",""
		objInitialBrowser.Navigate(strReportURL)
	Else
		SystemUtil.Run "iexplore.exe",""
		wait(2)
		Set objInitialBrowser = Browser("creationtime:=0")
		BrowserVersion = objInitialBrowser.GetAllROProperties().Item(3).Value
		'SystemUtil.Run "iexplore.exe",strReportURL,"",""
		objInitialBrowser.Navigate(strReportURL)
	End if
	Browser().Page().Sync
	Wait(5)
	If not isNullisEmptyCheck( BrowserVersion) ="" Then
		stepActual = "Successfully opened " & strBrowser & " Browser with URL :- " &  strReportURL 
		stepResult = "Passed"
		stepErrDescription = ""
	End If
	If err.Number<>0 Then
		stepErrDescription = "Unable to open "  & strBrowser & " Browser with URL :- " &  strReportURL &" for error :- " & err.description 
		stepResult = "Failed"
	Else
		stepActual = "Successfully opened " & strBrowser & " Browser with URL :- " &  strReportURL 
		stepResult = "Passed"
		stepErrDescription = ""
	End If
End Function

Function PerformObjectOperation(strTestObject, strOperation)
	
	On Error Resume Next
	Err.Clear
	
	'stepExpected = "" & strOperation & " Operation Done Successfully on object " & strTestObject
	
	If Not IsObject(eval(strTestObject)) Then
		'"Object " & strTestObject & "   Not Declared"
		stepErrDescription = UCASE(strAction) & " - " & "Object " & strTestObject & " not found"
		StepResult = "Failed"
		Exit Function
	End If
	
	'If  Not Eval(strTestObject & ".Exist(5)") Then
	If  Not VerifyObjExist(strTestObject) Then
		stepErrDescription = UCASE(strAction) & " - " & "Object " & strTestObject & " not present "
		StepResult = "Failed"
		Exit Function
	End If	
	
	'If the above object is a child item from the table then don't initialize
	If eval(strObjectName & ".GetROProperty(" & Chr(34) & "micClass" & Chr(34) & ")") <> "WebTable" Then
		Execute (strTestObject & ".Init")   
		'Eval(strTestObject&".Highlight")
		'Equiv of .RefreshObject of QTP 10.0 to reinitialize
		'object in OR in order identify same obj in app

		'After browser is refreshed
	End If
	
	Select Case Ucase(Trim(strAction))
		
		Case "CLICK"
			If isnull(strParameters) or isempty(strParameters) Then
				strParameters=""
			End If
			If Len(Trim(strParameters)) = 0  Then
					Execute (strTestObject & ".Click")
					If strObjectName="lnkSignOut" Then
						If Dialog("text:=Message from webpage").Exist Then
							Dialog("text:=Message from webpage").WinButton("text:=OK").Click
						End If
					End If
			Else
					Execute (strTestObject & ".Click " & chr(34) & strParameters & chr(34))
			End If
		
		Case "OPERATION" 
			'strOperation = eval(strOperation)
			Execute (strTestObject & "." & strOperation)
		
	End Select
	
	If Err.Number <> 0 Then
		stepErrDescription = UCASE(strAction) & " - " & "Failed to click the object [" & strTestObject & "]. " & Err.Description
		StepResult = "Failed"
		Err.Clear
	Else
		StepResult = "Passed"
		stepActual = "Click operation successfully performed on the object [" & strTestObject & "]."
		stepErrDescription = ""
		'Eval(strTestObject&".Highlight")
	End If

End Function

Public Function Verify_Function()
'This function will verify test object properties aginst test data properties
'eg innertext value of one object to data  provided in test data 

	On Error Resume Next
	Err.Clear
	
	Dim SplitParameters
	
	arrSplitParameters = SplitIgnoreCommasInQuotes(strParameters)
	strObjectProp = arrSplitParameters(0)
	strFieldVal = Trim(GetParameterValue(arrSplitParameters(1)))
	If strFieldVal = arrSplitParameters(1) Then
		strFieldValSplit = Trim(GetFormFieldValue(arrSplitParameters(1)))
		strFieldValUnSplit = split(strFieldValSplit,",")
		strFieldVal = strFieldValUnSplit(0)
	End If
	Select Case Ucase(strFieldVal)
    	Case "CONFIRM"
           strFieldVal = "true"
    		
    	Case "NA"
    	   strFieldVal = "false"
    	   
    	Case "TRUE"
'    	If Not Eval(strObjectName & ".GetROProperty(" & Chr(34) & "micClass" & Chr(34) & ")") = "WebList" Then
    	  strFieldVal = "1"
'        End If 
    	
    	Case "FALSE"
'    	If Not Eval(strObjectName & ".GetROProperty(" & Chr(34) & "micClass" & Chr(34) & ")") = "WebList" Then
    	   strFieldVal = "0"
'    	End If
    	
    	Case "NULL"
    	  strFieldVal = ""
    	
    	Case "SKIP" 
    	blnSkipRslt = False
		Exit Function
    	
    End Select	
			
	
	Execute (strObjectName & ".Init")
	
	
	If Not IsObject(eval(strObjectName)) Then
		'"Object " & strTestObject & "   Not Declared"
		stepErrDescription =  UCASE(strAction) & " - " & "Object " & strObjectName & " Not Found"
		Reporter.ReportEvent micFail,"VERIFY","Objet Not Found"
		StepResult = "Failed"
		Exit Function
	End If
	
	'If  Not Eval(strTestObject & ".Exist(5)") Then
	If  Not VerifyObjExist(strObjectName) Then
		stepErrDescription =  UCASE(strAction) & " - " & "Object " & strObjectName & " Not Present "
		Reporter.ReportEvent micFail,"VERIFY","Objet Not Present"
		StepResult = "Failed"
		Exit Function
	End If	
	
	'if the above object is a child item from the table then don't initialize
	If Trim(Ucase(strObjectName)) <> "OBJECT" Then
		Execute (strObjectName & ".Init")   
		'Equiv of .RefreshObject of QTP 10.0 to reinitialize
		'object in OR in order identify same obj in app

		'After browser is refreshed
	End If
	
	
	

	strObjectProp = Eval(strObjectName & ".GetROProperty(" & Chr(34) & strObjectProp & Chr(34) & ")")
	strObjectProp = Trim(strObjectProp)
	strFieldName2 = Eval(strObjectName & ".GetROProperty(" & Chr(34) & "acc_name" & Chr(34) & ")")
    strFieldName2 = Trim(strFieldName2)

	Select Case Ucase(arrSplitParameters(0))
		
		Case "INNERTEXT","TEXT","VALUE","LABEL","CHECKED"
			If Instr(1,strObjectProp,"^") Then
				strAppFieldVal = Split(strObjectProp,"^")
				If Eval(Trim(strAppFieldVal(1)) = strFieldVal) Then
					stepActual = "Field name " & strAppFieldVal(0) & " : " & strAppFieldVal(1) & " is Equal to [" & strFieldVal & "]."
					stepErrDescription = ""
					StepResult = "Passed"
					Exit Function
					Else 
					stepErrDescription = UCASE(strAction) & " - " & strAppFieldVal(1) & " is Not Equal to [" & strFieldVal & "]." & Err.Description
					Reporter.ReportEvent micFail,"VERIFY","Not matching"
					StepResult = "Failed"
					blnFail = False
				End If	
				Elseif Eval(Trim(strObjectProp) = strFieldVal) Then
					stepActual = "Field " & strObjectName & " (" & strFieldName2 & ") current value in application [" & strObjectProp & "] is Equal to [" & strFieldVal & "] ---As Expected."
'					stepActual = "Object " & strObjectName & " with name " & strObjectProp & " is Equal to [" & strFieldVal & "]."
					stepErrDescription = ""
					StepResult = "Passed"
					Exit Function	
					Else 
					stepErrDescription = UCASE(strAction) & " - Field " & strObjectName & " (" & strFieldName2 & ") current value in application [" & strObjectProp & "] is Not Equal to [" & strFieldVal & "] ---Not Expected." & Err.Description
'					stepErrDescription = UCASE(strAction) & " - " & "Verify Function " & strObjectProp & " is Not Equal to [" & strFieldVal & "]." & Err.Description
					Reporter.ReportEvent micFail,"VERIFY","Not matching"
					StepResult = "Failed"
					blnFail = False					
			End If
								
		Case "READONLY","DISABLED"
			If Eval(strObjectProp = strFieldVal) Then
				stepActual = "Object with logical name " & strObjectName & " is Read Only."
				stepErrDescription = ""
				StepResult = "Passed"
				Exit Function
				Else
					stepErrDescription = "Object with logical name " & strObjectName & " is not Read Only."
					Reporter.ReportEvent micFail,"VERIFY","Not Read Only"
					StepResult = "Failed"
					blnFail = False	
			End If
			
		Case Else	
			If Err.Number <> 0 Then
					stepErrDescription = UCASE(strAction) & " - " & "Verify Function " & strObjectName & " is not Equal to [" & strFieldVal & "]." & Err.Description
					Reporter.ReportEvent micFail,"VERIFY","Not Read Only"
					StepResult = "Failed"
					blnFail = False
					Err.Clear
			End If
		End Select
	
End Function

Public Function CallFunction()

	On Error Resume Next

	'Use the stub "SampleFunction" to create a new function
	If Len(Trim(strObjectName)) > 0  Then
			Execute ( strObjectName & " = " & strParameters )
			inputparameters.Add strObjectName,eval(strObjectName)
	Else 
			Execute (strParameters )
			Dim splitArguments
			If Instr(strParameters,"=") > 0 Then
				splitArguments = Split(strParameters,"=")
				inputparameters.Add Trim(splitArguments(0)),eval(splitArguments(0))
				If Err.Number = 457 Then
					Err.Clear
			    End if 
			End If
	End If

	If Err.Number <> 0 Then
			stepErrDescription = UCASE(strAction) & " - [" &  strParameters & "] is failed " & Err.Description
			Err.Clear
			stepResult ="Failed"
	Else
		    stepResult ="Passed"
			stepActual = UCASE(strAction) & " - [" & strParameters & "] is passed"
	End If


End Function
Function RetrieveValue_Function()
	On Error Resume Next
	Err.Clear
	
	If Not IsObject(eval(strObjectName)) Then
		'"Object " & strTestObject & "   Not Declared"
		stepErrDescription = UCASE(strAction) & " - " & "Object " & strObjectName & " not found"
		StepResult = "Failed"
		Exit Function
	End If

	'If  Not Eval(strObjectName & ".Exist(5)") Then
	If  Not VerifyObjExist(strObjectName) Then
		stepErrDescription = UCASE(strAction) & " - " & "Object " & strObjectName & " not present "
		StepResult = "Failed"
		Exit Function
	End If    

	Dim strobjClass,SplitParameters,strRowWithCellText,intColumnToValidate,strPropertyValue_Actual
	Execute (strObjectName & ".Init")
	strobjClass = eval(strObjectName & ".GetROProperty(" & Chr(34) & "micClass" & Chr(34) & ")")

	Select Case strobjClass
	
	Case "WebTable"
	
		SplitParameters = Split(strParameters, ",")
		strRowWithCellText = SplitParameters(0) 'It can have dt_ param in the parameters sheet
		intColumnToValidate = SplitParameters(1)
	
		strRowWithCellText = GetParameterValue(strRowWithCellText)
	
		Select Case isNumeric ( Trim(strRowWithCellText))
			'Usage is intRow, intColumnToValidate, strExpectedTxt
			Case True
				intTblRow = strRowWithCellText
			'Usage is strRowWithCellText, column, strText
			Case False
				intTblRow = eval(strObjectName & ".GetRowWithCellText(" & Chr(34) & strRowWithCellText & Chr(34) & ")")
		End Select
	
		strPropertyValue_Actual = eval(strObjectName & ".GetCellData(" & intTblRow & ", " & intColumnToValidate & ")")
	
	Case Else
	
		strPropertyName = strParameters
		strPropertyValue_Actual = eval(strObjectName & ".GetROProperty(" & Chr(34) & strPropertyName & Chr(34) & ")")
			Select Case strObjectName
				Case "txtOrderHeader"
						strVar1=strPropertyValue_Actual
				Case "wbeShipmentNumber"
						strVar2=strPropertyValue_Actual
			End Select
        	
	End Select

	If Err.Number <> 0 Then
				stepErrDescription = UCASE(strAction) & " - " & "RetrieveProperty_Function - Failed to retrieve the property [" & strObjectName & "] from the object [" & testObjectName & "]." & Err.Description
				StepResult = "Failed"
				Exit Function
	Else
				stepActual= "Actual value retrieved from the field: [" & strObjectName & "] is " & strPropertyValue_Actual & "."
				StepResult = "Passed"
				stepErrDescription = ""
	End If
									
	testparameters.Add strObjectName,strPropertyValue_Actual
	'Adds variable name and value to data sheet
	DataTable.GetSheet("Action1").AddParameter strObjectName,strPropertyValue_Actual

End Function

'==================================================
'* Function Name -										EnterCustGridLines_Function()
'* Function Description -							This function can be used to fill the grild line items
'* Created By -												
'* Created Date -										
'* Input Parameter -									
'* Output Parameter -								
'* Pre-Conditions -										
'* Post Conditions -									
'==================================================

Function EnterCustGridLines_Function()

   	 On Error Resume Next 
     Err.Clear 

	Dim arrTDSColNames, arrIterations, numIterationsCount, strVal, strValue, strMessage, strRowLine
	Dim arrTempVal, iCounter, jCounter, strPropertyValue1, strPropertyValue2
	Dim strTestObject1, strTestObject2

    strParameters = Replace(strParameters, " ", "")
    strParameters = Replace(strParameters, VbCrLf, "")

	strTestObject1= Trim(split(strObjectName, ",")(0))
	strTestObject2 = Trim(split(strObjectName, ",")(1))
        
	Execute (strTestObject1 & ".Init")
	If Not IsObject(eval(strTestObject1)) Then 
			stepErrDescription = "Object [" & strTestObject1 & "] was not evaluated as an object. Please check OR and/or .txt files" 
			StepResult = "Failed" 
			Exit Function 
	End If 
	
	If  Not VerifyObjExist(strTestObject1)  Then 
			stepErrDescription = "Object [" & strTestObject1 & "] not present on the application: Please check object properties in the OR" 
			StepResult = "Failed" 
			Exit Function 
	End If         
	
	strObjectClass1 = Eval(strTestObject1 & ".GetROProperty(" & Chr(34) & "micClass" & Chr(34) & ")") 
	If  Trim(strObjectClass1) <> "WebEdit" Then
			stepErrDescription = "Object [" & strTestObject1 & "] class is not WebEdit" 
			StepResult = "Failed" 
			Exit Function 
	End If


	Execute (strTestObject2 & ".Init")
	If Not IsObject(eval(strTestObject2)) Then 
			stepErrDescription = "Object [" & strTestObject2 & "] was not evaluated as an object. Please check OR and/or .txt files" 
			StepResult = "Failed" 
			Exit Function 
	End If 
	
	If  Not VerifyObjExist(strTestObject2)Then 
			stepErrDescription = "Object [" & strTestObject2 & "] not present on the application: Please check object properties in the OR" 
			StepResult = "Failed" 
			Exit Function 
	End If         
	
	strObjectClass2 = Eval(strTestObject2 & ".GetROProperty(" & Chr(34) & "micClass" & Chr(34) & ")") 
	If Trim(strObjectClass2) <> "WebCheckBox" Then
			stepErrDescription = "Object [" & strTestObject2 & "] class is not WebCheckBox" 
			StepResult = "Failed" 
			Exit Function 
	End If

'	Set objFrame = Browser("JD Edwards EnterpriseOne").Page("JD Edwards EnterpriseOne").Frame("html id:=e1menuAppIframe")		 
    arrTDSColNames = Split(strParameters, ",")
	arrIterations = Split(testparameters(Trim(arrTDSColNames(0))),"||")
	numIterationsCount = UBound(arrIterations) 
	Set arrIterations = Nothing

	strMessage = ""
	For iCounter = 0 to numIterationsCount

         strPropertyValue2 = chr(34) & "index" & chr(34) &"," & chr(34) & iCounter & chr(34)
        Execute (strTestObject2 & ".SetToProperty "  & strPropertyValue2 )
		If Err.Number <> 0 Then 
			stepErrDescription = "[" & strTestObject2 & "] Unable to set index to " & iCounter & ": " & Err.Description 
			StepResult = "Failed" 
			Exit Function 
		End If
	    Wait(1)

		strPropertyValue2 = ""
	   If  Not VerifyObjExist(strTestObject2) Then 
			stepErrDescription = "Row " &  iCount & ", Object [" & strTestObject2 & "] not present: Please check object properties" 
			StepResult = "Failed" 
			Exit Function 
	     End If  

		Execute (strTestObject2 & ".Set " & chr(34) & "ON" & Chr(34) ) 
		If Err.Number <> 0 Then 
			stepErrDescription = "[" & strTestObject2 & "] Unable to set checkbox to  [ON]." & Err.Description 
			StepResult = "Failed" 
			Exit Function 
		End If
		Wait(1)

	    strRowLine =  "" '(iCounter + 1) & ":"
		For jCounter=0 to ubound(arrTDSColNames)
				strPropertyValue1 = chr(34) & "index" & chr(34) &"," & chr(34) & jCounter & chr(34)
				Execute (strTestObject1 & ".SetToProperty "  & strPropertyValue1 )
				If Err.Number <> 0 Then 
						stepErrDescription = "UdfEnterGridLineItems - [" & strTestObject1 & "] Unable to set index to " & jCounter & ": " & Err.Description 
						StepResult = "Failed" 
						Exit Function 
				End If
				strPropertyValue1 = ""

	            If  Not VerifyObjExist(strObject1) Then
					stepErrDescription = "UdfEnterGridLineItems - Object [" & strObject1 & "] not Present: Please check object properties" 
					StepResult = "Failed" 
					Exit Function 
				 End If  

				Wait(1)
				arrTempVal = Split(testparameters(Trim(arrTDSColNames(jCounter))), "||")
				strVal = Trim(arrTempVal(iCounter))
				strValue = chr(34) & strVal & chr(34) 
				Execute (strTestObject1 & ".Set " & strValue ) 
				If Err.Number <> 0 Then 
						stepErrDescription = " [" & strTestObject1 & "] : Unable to set the value to [" & strValue & "]." & Err.Description 
						StepResult = "Failed" 
						Exit Function 
				Else
						strRowLine = strRowLine & strVal & ","
				End If
				strVal = ""
		Next

		Wait(1)
		strMessage = strMessage & strRowLine & VbCrLf
	
	Next

		If Err.Number = 0 Then
					stepActual = "Completed" & VbCrlf & strMessage
					stepErrDescription = ""
					StepResult = "Passed"
		End If

    Set arrTempVal = Nothing
	Set arrTDSColNames = Nothing

End Function

'=======================================================================================================
'* Function Name -										DownloadTestScript_Function
'* Function Description -							This function will download an attachment from Test Plan
'* Created By -												
'* Created Date -										 
'* Input Parameter -								   N/A
'* Output Parameter -								N/A
'* Pre-Conditions -										
'* Post Conditions -									
'=======================================================================================================
Public Function DownloadTestScript_Function()

	On Error Resume Next
	Err.Clear

	' Get Parameters from Test Data sheet
	Dim arrParameters, TestID , FileName
	arrParameters = Split(strParameters,",")
	TestID = GetParameterValue(arrParameters(0))
	FileName = GetParameterValue(arrParameters(1))
			
	' QC Connection	
	Set TDConnection = TDUtil.TDConnection
	flag="false"

	' Get the object of required test
	Set TestList = TDConnection.TestFactory.NewList("SELECT * FROM TEST WHERE TS_TEST_ID  = "&TestID)
	Err.Clear
	On Error Resume Next

	' If Test Script exists in Test Plan
	If TestList.Count > 0 Then
			' Download the attachment from Test Plan
			GetAttachmentFromTest = GetAttachmentFromTestObject(TestList(1), FileName, TestID)
			If GetAttachmentFromTest="Passed" Then
					flag="true"
			Else
					flag="false"
			End If
	Else
			flag="false"
	End If

	' If attachment is not downloaded then print fail message in QC
	If flag="false" Then
			stepErrDescription = UCASE(strAction) & " - " & "Test Attachment: '" & FileName & "' not found. Please verify Test ID or Attachment name."
			StepResult = "Failed"
			Exit Function
	End If

	' If attachment is downloaded successfully then print Pass message in QC
	stepActual= "Test Attachment: '" & FileName & "' is sucessfully downloaded  to '" & Environment.Value("DOWNLOAD_PATH")   & "\' Path."
	StepResult = "Passed"
	stepErrDescription = ""
End Function 

Public Function GetAttachmentFromTestObject (TestObject, FileName, TestID)
	MyPath = GetAttachmentServerPath(TestObject, FileName, TestID)
	If StrComp(MyPath, "") = 0 Then
			GetAttachmentFromTestObject ="Failed"
			GetAttachment= ""
			stepErrDescription = UCASE(strAction) & " - " & "Test Script Attachment " & FileName & " not found"
			StepResult = "Failed"
			Exit Function
	End If
	If Right(Environment.Value("DOWNLOAD_PATH"), 1) <> "\" Then
			OutPath = Environment.Value("DOWNLOAD_PATH") & "\"
	End If
	GetAttachmentFromTestObject = "Passed"
End Function

' Download attachment from specific Test Script
Public Function GetAttachmentServerPath (TestObject, FileName, TestID)
	Set AttachmentFactory = TestObject.Attachments
	Set AttachmentList = AttachmentFactory.NewList("SELECT * FROM CROS_REF")
	For Each Attachment in AttachmentList
			If StrComp(Attachment.Name(1), FileName, 1) = False Then
					LongFileName = Attachment.Name
					Pos = Instr(1, Attachment.ServerFileName, Attachment.Name, 1)
					GetAttachmentServerPath = Left(Attachment.ServerFileName, Pos - 1)
					Set TestAttachStorage = Attachment.AttachmentStorage 
					TestAttachStorage.ClientPath=Environment.Value("DOWNLOAD_PATH")
					TestAttachStorage.ServerPath="\\Attach\"
					Attachment.Load True,Environment.Value("DOWNLOAD_PATH")
					' Rename the downloaded file
					Set FSO = CreateObject("Scripting.FileSystemObject")
					strRename= split(Attachment.Name,TestId&"_")
					If FSO.FileExists(Environment.Value("DOWNLOAD_PATH")&"\"&Attachment.Name) Then
							If FSO.FileExists(Environment.Value("DOWNLOAD_PATH")&"\"&strRename(1)) Then
									FSO.DeleteFile Environment.Value("DOWNLOAD_PATH")&"\"&strRename(1)
							End If
							FSO.MoveFile Environment.Value("DOWNLOAD_PATH")&"\"&Attachment.Name, Environment.Value("DOWNLOAD_PATH")&"\"&strRename(1)
					End If
					Set FSO = Nothing
			End If
	Next
End Function


'=======================================================================================================

'=======================================================================================================
'* Function Name -									GridRecordCount_Function
'* Function Description -							This function will Capture number of records displayed in table
'* Created By -												
'* Created Date -										 
'* Input Parameter -								N/A
'* Output Parameter -								N/A
'* Pre-Conditions -										
'* Post Conditions -									
'=======================================================================================================
Function GridRecordCount_Function()

	On Error Resume Next
	Err.Clear

	' Get Parameter value
	Dim RecordNumber

    RecordNumber = GetParameterValue(strParameters)

	' Setting Page description
	Set oPage = Browser("micclass:=Browser","title:="&strBrowserTitle).Page("title:=.*")

	' Capture no. of records
	If oPage.WebElement("innertext:=No records found.*","html tag:=B").Exist(10) Then
			Reporter.ReportEvent micPass,"No Records found","No records found."
			strValue = 0
			DataTable.AddSheet("Action1").AddParameter RecordNumber, strValue
			testparameters.Add RecordNumber, strValue
	Else
            varRecord = oPage.WebElement("innertext:=Records1 - [\d]*","html tag:=TD", "class:=gridheader").GetROProperty("innertext")
			varRecordNumber = split(varRecord,"-")
			strValue = varRecordNumber(1)
			DataTable.AddSheet("Action1").AddParameter RecordNumber, strValue
			testparameters.Add RecordNumber, strValue
	End If
	

	' Displaying Fail message to QC result
		' If value ts not captured
		If DataTable.Value(RecordNumber,"Action1")="" Then
				stepErrDescription = UCASE(strAction) & " - " & "No. of records are not captured successfully"
				StepResult = "Failed"
				Exit Function
		End If

	' If value is captured successfully display pass message in QC
    stepActual= "No. of records are captured successfully from the Grid. ' "& RecordNumber &" ' = " & strValue  & "'."
	StepResult = "Passed"
	stepErrDescription = ""

End Function


'=======================================================================================================

Function fnDateDaySelect()
    Dim Dateinnertext
'	str = Split(strParameters,",")
'	Dateinnertext = Trim(str(0))
       Dateinnertext = GetParameterValue(strParameters)
	Call fnDateDaySelectFromDatePicker(Dateinnertext)
End Function


Function fnDateDaySelectFromDatePicker(ExpectedDateInnertext)
Dim DateInnertext
Set objPage = Browser("title:=.*").Page("title:=.*").WebElement("class:=GroupHeading","innertext:=This form has ([1-9]|[1-9]\d) Errors ([0-9]|[0-9]\d) Warnings")
If objPage.Exist Then
		objPage.Highlight
		stepErrDescription = UCASE(strAction) & " - " & "CheckErrorExist_Function failed - This Form has Errors" & Err.Description
		StepResult = "Failed" 
		Exit Function 
Else
	Set oPage = Browser("title:=.*").Page("title:=.*")
	Set objWebTable = oPage.Frame("title:=.*","html tag:=IFRAME","html id:=e1menuAppIframe").WebTable("text:=.*","html tag:=TABLE","index:=105")
	Set oDesc =  Description.Create
	oDesc("micclass").value = "WebElement"
	oDesc("html tag").value= "DIV"
	oDesc("innertext").value = "([1-9]|[12][0-9]|3[0-1])"
	
	Set DateValuesCount = objWebTable.ChildObjects(oDesc)
'	msgbox DateValuesCount.count
	blnDateFound = False
	For i = 0 To DateValuesCount.count-1
		DateInnertext = DateValuesCount(i).Getroproperty("innertext")
'	Print DateInnertext
		If Trim(DateInnertext) = Trim(ExpectedDateInnertext) Then
			DateValuesCount(i).click
			blnDateFound = True
			Exit For
		End If
	Next 
		If (blnDateFound = False) Then
			Reporter.ReportEvent micFail, "The Expected " & ExpectedDateInnertext & " could not be found", "Please check Expected date " & columnName & " is not present in DatePicker"
			Exit Function
		End If
End If

Set oPage = Nothing
Set objWebTable = Nothing
Set objPage = Nothing
Set DateValuesCount = Nothing
Set oDesc = Nothing

End Function

Function ItemCheckList_Function()
	
		On Error Resume Next
		Err.Clear
		
		Dim strItemName, arrItemName, TotalItems, blnVal, strAppList, ActItemName, strFailResult
		strFailResult = ""
		strAppList = ""
'		Dim strItemName, TotalItems, ActItemName
'		Set blnItemFound = True
		
		Execute (strObjectName & ".Init")
		
		If Not IsObject(eval(strObjectName)) Then
			stepErrDescription =  UCASE(strAction) & " - " & "Object " & strObjectName & " Not Found"
			Reporter.ReportEvent micFail,"ItemCheckList","Object Not Found"
			StepResult = "Failed"
			Exit Function
		End If
		
		If  Not VerifyObjExist(strObjectName)  Then
			stepErrDescription =  UCASE(strAction) & " - " & "Object " & strObjectName & " Not Present "
			Reporter.ReportEvent micFail,"ItemCheckList","Object Not Present"
			StepResult = "Failed"
			Exit Function
		End If	
		
		If Trim(Ucase(strObjectName)) <> "OBJECT" Then
			Execute (strObjectName & ".Init")   
		End If
		
		
		strItemName = GetParameterValue(strParameters)
		arrItemName = Split(strItemName,",")
	
		TotalItems = Eval(strObjectName &".GetROProperty("  &  chr(34)  & "Items Count" &  chr(34) &  ")")
		
		For ItemVal = 0 to Ubound(arrItemName)
			blnVal = False
			strAppList = ""
			For AppItem = 1 To TotalItems
				ActItemName =  Eval(strObjectName &".GetItem(" & AppItem & ")")
				If Trim(UCase(ActItemName)) = Trim(UCase(arrItemName(ItemVal)))  Then
					blnVal = True
				End If
				strAppList = ActItemName & "," & strAppList
			Next
			
			If blnVal = False Then
				strFailResult = arrItemName(ItemVal) & "," & strFailResult
			End If
		Next
		
'		With (New RegExp)
'		    .Global = True
'		    .Pattern = "\D" 'matches all non-digits
'		    strFailResult = .Replace(strFailResult, "") 'all non-digits removed
'		    strItemName = .Replace(strItemName, "")		    
'		End With
		
		If Len(strFailResult)>1 Then
			stepErrDescription = UCASE(strAction) & " - Version(s) " & strFailResult & " not found in application list " & strAppList & "."
			Reporter.ReportEvent micFail,"ItemCheckList","Item Not Found"
			StepResult = "Failed"
			blnFail = False
			Err.Clear
			Exit Function
			Else 
				stepActual = UCASE(strAction) & " - Version(s) " & strItemName & " is found in application list " & strAppList & "." 
				stepErrDescription = ""
				StepResult = "Passed"
		End If
		
'		
'		strItemName = GetParameterValue(strParameters)
'	
'		TotalItems = Eval(strObjectName &".GetROProperty("  &  chr(34)  & "Items Count" &  chr(34) &  ")")
'		If Not TotalItems > 1  Then
'			Execute (strObjectName & ".Init")
'			TotalItems = Eval(strObjectName &".GetROProperty("  &  chr(34)  & "Items Count" &  chr(34) &  ")")
'		End If
''		Eval(strObjectName &".RefreshObject")
'		For ItemNum = 1 To TotalItems
'			ActItemName =  Eval(strObjectName &".GetItem(" & ItemNum & ")")
'			If Trim(UCase(ActItemName)) = Trim(UCase(strItemName)) Then
'				blnItemFound = True
'				stepErrDescription = UCASE(strAction) & " - " & " Item " & ActItemName & " is Equal to [" & strItemName & "]." 
'				StepResult = "Passed"
'				Exit For
'			End If
'		Next
'		If blnItemFound = False Then
'			stepErrDescription = UCASE(strAction) & " - " & strItemName & " item is not found in " & strObjectName & " Weblist " 
'			StepResult = "Failed"
'		End If
		
End Function


Function ItemSortList_Function()
	
		On Error Resume Next
		Err.Clear

Dim strFailResult
strFailResult = ""

Execute (strObjectName & ".Init")

If Not IsObject(eval(strObjectName)) Then
			stepErrDescription =  UCASE(strAction) & " - " & "Object " & strObjectName & " Not Found"
			Reporter.ReportEvent micFail,"ITEMSORTLIST","Object Not Found"
			StepResult = "Failed"
			Exit Function
		End If

If  Not VerifyObjExist(strObjectName) Then
			stepErrDescription =  UCASE(strAction) & " - " & "Object " & strObjectName & " Not Present "
			Reporter.ReportEvent micFail,"ITEMSORTLIST","Object Not Present"
			StepResult = "Failed"
			Exit Function
		End If	

If Trim(Ucase(strObjectName)) <> "OBJECT" Then
			Execute (strObjectName & ".Init")   
		End If

TotalItems = Eval(strObjectName &".GetROProperty("  &  chr(34)  & "all items" &  chr(34) &  ")")

arrArrayAsc= Split(TotalItems,";")

Set oArrayList = CreateObject("System.Collections.ArrayList")
 
For Each sElement in arrArrayAsc
    If IsNumeric(sElement) Then
      oArrayList.Add CInt(sElement)
    Else
      oArrayList.Add sElement
    End If
Next

'Sort Ascending
oArrayList.Sort

'Verify all elements in the Ascending array with the sorted oArrayList
For x = LBound(arrArrayAsc) to UBound(arrArrayAsc)
blnVal = True
   'If the array element does not match the sorted element, then the original
        'array is not in ascending order
   If Not arrArrayAsc(x) = oArrayList(x) Then
blnVal = False
stepErrDescription = "Object with logical name " & strObjectName & " is not in ascending order."
Reporter.ReportEvent micFail,"ITEMSORTLIST","The available forms list is not in ascending order"
StepResult = "Failed"
blnFail = False
Exit For
 End If
If blnVal = True Then
stepActual = "Object with logical name " & strObjectName & " is in ascending order."
stepErrDescription = ""
StepResult = "Passed"
'Reporter.ReportEvent micPass,"ITEMSORTLIST","The available forms list is in ascending order"
End If
Next

Set oArrayList = Nothing
End Function

Public Function OpenPDFclosing_Function()
'Option Explicit

Dim objIE
Dim objShell

Set objShell = CreateObject("Shell.Application")
 For Each objIE In objShell.Windows
'    WScript.Echo objIE.LocationURL
    If InstrRev(LCase(objIE.LocationName), ".pdf") > 0 Then
        objIE.Quit
        stepActual= "Open PDF is Closed Successfully"
		StepResult = "Passed"
		stepErrDescription = ""
   Else 
		stepErrDescription = UCASE(strAction) & " - " & "Open PDF is not Closed Successfully"
		StepResult = "Failed"
   End if
 next

Set objIE = Nothing
Set objShell = Nothing
End Function

Function fnDropDownMultiSelection()

str = Split(strParameters,",")
strDelimiter= GetParameterValue(str(0))
strValuesToSelect = GetParameterValue(str(1))
call DropDownMultiSelect(strDelimiter,strValuesToSelect)
End function 


Function DropDownMultiSelect(strDelimiter,strValuesToSelect)

Dim lngIndex,strValues

	On Error Resume Next
		Err.Clear
Execute (strObjectName & ".RefreshObject " )
Execute (strObjectName & ".Init")
		
		If Not IsObject(eval(strObjectName)) Then
			stepErrDescription =  UCASE(strAction) & " - " & "Object " & strObjectName & " Not Found"
			Reporter.ReportEvent micFail,"MULTISELECT","Object Not Found"
			StepResult = "Failed"
			Exit Function
		End If
		
		If  Not VerifyObjExist(strObjectName) Then
			stepErrDescription =  UCASE(strAction) & " - " & "Object " & strObjectName & " Not Present "
			Reporter.ReportEvent micFail,"MULTISELECT","Object Not Present"
			StepResult = "Failed"
			Exit Function
		End If	
		
		If Trim(Ucase(strObjectName)) <> "OBJECT" Then
			Execute (strObjectName & ".Init")   
		End If
strFieldName = Eval(strObjectName & ".GetROProperty(" & Chr(34) & "acc_name" & Chr(34) & ")")
strFieldName = Trim(strFieldName)

If Eval(strObjectName &".GetROProperty("  &  chr(34)  & "select type" &  chr(34) &  ")") = "Extended Selection" then

If strDelimiter = "ALL" Then

strValues = Eval(strObjectName &".GetROProperty("  &  chr(34)  & "all items" &  chr(34) &  ")")
strValues = Split(strValues,";")

Else
strValues = Split(strValuesToSelect, strDelimiter)

End if

For lngIndex = LBound(strValues) to UBound(strValues)
	strselectvalue = strValues(lngIndex)
   	Select Case lngIndex
    	Case 0   
'     objListBox.Select strValues(lngIndex)   
     		Execute (strObjectName & ".Select " & chr(34) & strselectvalue & chr(34))
     		StepResult = "Passed"
			stepActual = "Item selected operation successfully performed on the weblist object [" & strObjectName & " ( "  & strFieldName & " )]."
			stepErrDescription = ""
     		
    	Case Else
'     objListBox.ExtendSelect strValues(lngIndex)
	 		Execute (strObjectName & ".ExtendSelect " & chr(34) & strselectvalue & chr(34))
	 		StepResult = "Passed"
			stepActual = "MultiSelected items operation successfully performed on the weblist object [" & strObjectName & " ( "  & strFieldName & " )]."
			stepErrDescription = ""
	 		
    End Select
Next 'lngIndex
 Else
'  		msgbox "MultiSelect not supported by this control"
  			stepErrDescription =  UCASE(strAction) & " - " & "Object " & strObjectName & " MultiSelect not supported by this control "
			Reporter.ReportEvent micFail,"MULTISELECT","Multiselect control is not supported for this weblist Object"
			StepResult = "Failed"
 End If
 'Eval(strObjectName & ".Highlight")
 If Err.Number <> 0 Then
			stepErrDescription = UCASE(strAction) & " - " & "Failed to select Multiple items in the weblist object [" & strObjectName & " ( "  & strFieldName & " )]."
			Reporter.ReportEvent micFail,"MULTISELECT","DropDownMultiSelect not Performed"
			StepResult = "Failed"
			blnFail = False
			Err.Clear
		End If 
 End Function
 
 Function fnUpdateWebeditInWebtable()
str = Split(strParameters, ",")

Colname = GetParameterValue(str(0))
strtext = GetParameterValue(str(1))
Classname = GetParameterValue(str(2))
StrtexttoUpdate = GetParameterValue(str(3))

Call UpdateWebeditInWebtable(Colname,strtext,Classname,StrtexttoUpdate)
End Function


'==================================================
'* Function Name -									UpdateWebeditInWebtable(Colname,strtext,Classname,StrtexttoUpdate)
'* Function Description -							This function will update the webedit field based on column name and text
' Parameters need to passed : Colname,strtext,Classname,StrtexttoUpdate
'Colname = "Location Name"
'strtext = "Aut"
'Classname = "WebEdit"
'StrtexttoUpdate = "Aut1Update"
'==================================================

Function UpdateWebeditInWebtable(Colname,strtext,Classname,StrtexttoUpdate)
On Error Resume Next
Err.Clear
Set oPage = Browser("title:=.*").Page("title:=.*")
Set objWebTable = oPage.WebTable("html tag:=TABLE","role:=grid")
objWebTable.highlight
ColNum = fn_columnnumbycolumname(Colname)
RowNum = Fn_GetRownumwithtext(strtext,ColNum)
TotalChilditemcount = objWebTable.ChildItemCount(RowNum,ColNum,Classname)
For i  = 0 To TotalChilditemcount-1
	Set owebedit = objWebTable.ChildItem(RowNum,ColNum,Classname,i)
'	owebedit.Click
	owebedit.set StrtexttoUpdate
'	owebedit.Click
'	Call SendKeys ("TAB")
Wait(2)	
Next

If Err.Number <> 0 Then
					stepErrDescription = UCASE(strAction) & " - " & "The (" & Colname & ") " & "value " & Chr(34) & strtext & Chr(34) & " is not Updated with " & Chr(34) & StrtexttoUpdate & Chr(34) & Err.Description
					StepResult = "Failed"
					Err.Clear
					Exit Function
		Else
					stepActual =  "The (" & Colname & ") " & "value " & Chr(34) & strtext & Chr(34) & " is Updated with " & Chr(34)& StrtexttoUpdate & Chr(34)&" Successfully"
					stepErrDescription = ""
					StepResult = "Passed"
					Exit Function
		End If
	
Set oPage = Nothing
Set objWebTable = Nothing
End Function
 
  '==================================================
'* Function Name -									fn_columnnumbycolumname(StrExpectedcolumnname)
'* Function Description -							This function will get the Column number  based on column Name.
' Parameters need to passed : StrExpectedcolumnname
'StrExpectedcolumnname = "Location Name"
'==================================================
 
Function fn_columnnumbycolumname(StrExpectedcolumnname)
On Error Resume Next
Err.Clear
Dim Totalrows,Totalcolumns,blncolumnfound,columnnum,strColumnname
Set oPage = Browser("title:=.*").Page("title:=.*")
Set objWebTable = oPage.WebTable("html tag:=TABLE","role:=grid")
Totalrows = objWebTable.RowCount
Totalcolumns = objWebTable.ColumnCount(1)
blncolumnfound = False
	For i = 1 To Totalcolumns
		strColumnname = objWebTable.GetCellData(1,i)
		If UCASE(TRIM(strColumnname)) = UCASE(TRIM(StrExpectedcolumnname)) Then
			columnnum = i
			fn_columnnumbycolumname = columnnum
			blncolumnfound = True
			stepActual = "The column " & StrExpectedcolumnname & " found in webtable"
			stepErrDescription = ""
			StepResult = "Passed"
			Exit For
		End If
	Next
If (blncolumnfound = False) Then
		stepErrDescription = "The column " & StrExpectedcolumnname & " could not be found, Please check if the column " & StrExpectedcolumnname & " is present in Webtable"  & Err.Description
		StepResult = "Failed"
		Err.Clear
		Exit Function
End If

Set oPage = Nothing
Set objWebTable = Nothing
End Function

'==================================================
'* Function Name -									Fn_GetRownumwithtext(StrExpectedtext,colnum)
'* Function Description -							This function will get the rownum with text based on column number.
' Parameters need to passed : StrExpectedtext,colnum
'StrExpectedtext =  "Aut"
'colnum = 1
'==================================================

Function Fn_GetRownumwithtext(StrExpectedtext,colnum)
	
Set oPage = Browser("title:=.*").Page("title:=.*")
Set objWebTable = oPage.WebTable("html tag:=TABLE","role:=grid")
Totalrows = objWebTable.RowCount

Set PageNum=description.Create
	PageNum("micclass").value="Link"
	PageNum("html tag").value = "A"
PageNum("class").value = "page-link"
	
rownum = 0

For i = 2 To Totalrows
	StrCelltext = objWebTable.GetCellData(i,colnum)
	If TRIM(StrCelltext) = TRIM(StrExpectedtext)Then
		rownum = i
		Fn_GetRownumwithtext = rownum
		Exit for
	End If
Next

If rownum = 0 Then
	For i = 2 To 200 Step 1
		if rownum>0 then
'			Do until oPage.Link("name:=Previous","class:=page-link").GetROProperty("color") = "#999999"
'			oPage.Link("name:=Previous","class:=page-link").click
'			Wait(1)
'			Loop
			Exit For
		End if
			PageNum("innerText").value=i
			If oPage.Link(PageNum).exist(5) Then
				oPage.Link(PageNum).Click
				Totalrows = objWebTable.RowCount
				For j = 2 To Totalrows
					StrCelltext = objWebTable.GetCellData(j,colnum)
					If TRIM(StrCelltext) = TRIM(StrExpectedtext)Then
					rownum = j
					Fn_GetRownumwithtext = rownum
					Exit for
					End If
				Next
			End If
	Next
Else 
'msgbox "Row Number already fetched in 1st Page itself"
End If
Set oPage = Nothing
Set objWebTable = Nothing
End Function
 
 Function fnCheckboxInWebtable()
	strcheckboxinput = GetParameterValue(strParameters)
	call CheckboxInWebtable(strcheckboxinput)
End Function

Function CheckboxInWebtable(strChecked)
On Error Resume Next
Err.Clear
Set oPage = Browser("title:=.*").Page("title:=.*")
Set objWebTable = oPage.WebTable("html tag:=TABLE","role:=grid")
If not objWebTable.Exist Then
	stepActual =  "No records found."
	stepErrDescription = ""
	StepResult = "Passed"
	Exit Function
End If
'objWebTable.highlight
' msgbox objWebTable.RowCount
Set oChkBox = Description.Create()
oChkBox("micClass").value = "WebCheckBox"
oChkBox("disabled").value = 0

Set WbCheckboxes = objWebTable.ChildObjects(oChkBox)
strCount =  WbCheckboxes.count

For i = 1 To WbCheckboxes.count-1
	WbCheckboxes(i).set strChecked 
	'If i = strCount  Then
	Exit for
	'End If
Next

If Err.Number <> 0 Then
					stepErrDescription = UCASE(strAction) & " - " & "The checkbox is not checked successfully" & Err.Description
					StepResult = "Failed"
					Err.Clear
					Exit Function
		Else
					stepActual =  "The checkbox is checked successfully"
					stepErrDescription = ""
					StepResult = "Passed"
					Exit Function
		End If
Set oPage = Nothing
Set objWebTable = Nothing	
End Function

Function ScrollBar_Function()
 
 str = Split(strParameters,",")
 
 Select Case trim(str(0))
		
		Case "PageDown"
    	KeyNumber = 209

		Case "PageUp"
		KeyNumber = 201

End Select

 	NStrokes = trim(str(1))
 	
 	For i = 1 To NStrokes Step 1
 	Set devicereplay = CreateObject("Mercury.DeviceReplay")
 	Browser("CreationTime:=0").highlight
 	devicereplay.PressKey KeyNumber
 	Next
 	
 Set objEMCOBrw   = Nothing
 Set devicereplay = Nothing
 End Function
 
 
'==================================================
'* Function Name -									UDFInputDateAndCurrentTime
'* Function Description -							This function will update the webedit field based on Date Format
' *Parameters need to passed : 					Date format 
'*Classname = "WebEdit", "WinEdit", "PbEdit", "VbEdit", "TeField", "WebFile" 	
'*Format Supported:
'	1. MM/DD/YYYY
'	2. MM/DD/YYYY HH:MM:SS
'	3. HH:MM:SS
'	4. MM-DD-YYYY
'	5. MM-DD-YYYY HH:MM:SS
'==================================================
 Function UDFInputDateAndCurrentTime()
 	On Error Resume Next
	Err.Clear
	Dim SplitParameters
			strDate = Now()
            		currentDate = FormatDateTime(strDate, 0)
            		CurrentMonth = Split(Split(currentDate, " ")(0), "/")(0)
                	CurrentMonth = Right("00" & CurrentMonth, 2)
            		CurrentDay = Split(Split(currentDate, " ")(0), "/")(1)
               	 CurrentDay = Right("00" & CurrentDay, 2)
            		CurrentYear = Split(Split(currentDate, " ")(0), "/")(2)
            		CurrentSecond = Split(Split(FormatDateTime(strDate, 3), " ")(0), ":")(2)
                	CurrentSecond = Right("00" & CurrentSecond, 2)
            		CurrentMinute = Split(FormatDateTime(strDate, 4), ":")(1)
                	CurrentMinute = Right("00" & CurrentMinute, 2)
            		CurrentHour = Split(FormatDateTime(strDate, 4), ":")(0)
                	CurrentHour = Right("00" & CurrentHour, 2)
                	
 			Dim currentDate,strFormatedDate
 			strFormat = GetParameterValue(strParameters)
 			If strFormat="" Then
 				strFormat = "MM/DD/YYYY"
 			Else
 				strFormat = UCASE(strFormat)
 			End If
			
			Select Case strFormat
				Case "MM/DD/YYYY"
					strFormatedDate = CurrentMonth & "/" &  CurrentDay  & "/" & CurrentYear
				Case "HH:MM:SS"
					strFormatedDate =  CurrentHour & ":" & CurrentMinute & ":" & CurrentSecond
				Case "MM/DD/YYYY HH:MM:SS"
					strFormatedDate = CurrentMonth & "/" &  CurrentDay  & "/" & CurrentYear & " " & CurrentHour &  ":" & CurrentMinute & ":" & CurrentSecond
				Case "MM-DD-YYYY"
					strFormatedDate = CurrentMonth & "-" &  CurrentDay  & "-" & CurrentYear
				Case "HH:MM:SS"
					strFormatedDate =  CurrentHour &":" & CurrentMinute & ":" & CurrentSecond
				Case "MM-DD-YYYY HH:MM:SS"
					strFormatedDate = CurrentMonth & "-" &  CurrentDay  & "-" & CurrentYear & " " & CurrentHour & ":" & CurrentMinute & ":" & CurrentSecond
			End Select
			
	Execute (strObjectName & ".Init")

	strObjectClass = Eval(strObjectName & ".GetROProperty(" & Chr(34) & "micClass" & Chr(34) & ")")
	strObjectClass = Trim(strObjectClass)

	Select Case strObjectClass
		Case "WebEdit", "WinEdit", "PbEdit", "VbEdit", "TeField", "WebFile" 			
 			Call SetObjectValue (strObjectName, strFormatedDate)
		Case Else
			Reporter.ReportEvent micFail, "Enter CurrentDate in Format " & strFormat, strObjectName & " is not a Edit Field"
	End  Select
 	
 	If Err.Number <> 0 Then
		stepErrDescription = UCASE(strAction) & " - " & "Current Date Not set" & Err.Description
		StepResult = "Failed"
		Err.Clear
		Exit Function
	End  If
 End Function
 
Public  Function SplitIgnoreCommasInQuotes(text)
 	Dim arr, regex
 	
 	Set regex = New RegExp
 	regex.Pattern = ",(?=([^""]*""[^""]*"")*[^""]*$)"
 	result = Split(regex.Replace(text,";"),";")
 	
 	
 	SplitIgnoreCommasInQuotes = result
 End Function
 
 Public Function fnSynchronisePage()
	On error resume next
	
	If Not IsObject(eval(strObjectName)) Then
		'"Object " & strTestObject & "   Not Declared"
		stepErrDescription =  UCASE(strAction) & " - " & "Object " & strObjectName & " not found"
		StepResult = "Failed"
		Exit Function
	End If
	
	Execute (strObjectName & ".Init")
	Eval(strObjectName & ".Sync")
	If Err.Number <> 0 Then
				stepErrDescription = UCASE(strAction) & " - " & "Failed to Sync page [" & strObjectName & "] " & Err.Description
				StepResult = "Failed"
				Exit Function
	Else
				stepActual= "Page sync successful: [" & strObjectName & "] ."
				StepResult = "Passed"
				stepErrDescription = ""
	End If
End  Function

Public Function fnBatchNumber()
	On error resume next
	Set WSHShell = CreateObject("Wscript.Shell")
	strPythonCmd = "python C:\Users\adalamr\Downloads\pdfreader.py" 
	Set WSHExec = WSHShell.Exec(strPythonCmd)
	wait 5  
	strOutput = WSHExec.StdOut.ReadAll
	batchNumber = strOutput
	wait 5
End Function

Public Function dropDownSelectByValue()
	On Error Resume Next
	Err.Clear
	Dim value
	For Each value In strObjectName.GetTOProperties
		Print strObjectName.GetTOProperty(value)	
	Next
	strObjectClass = strObjectName.Type
	strObjectClass = Trim(strObjectClass)
	strValue = GetParameterValue(strParameters)
	If strObjectClass = "WebList" Then
		strAllItems = Split(strObjectName.GetROProperty("All Items"),";")
		For intCounter = LBound(strAllItems) To UBound(strAllItems) Step 1
			If InStr(1,strAllItems(intCounter),strFormat) > 0 Then
				Call SetObjectValue(strObjectName, strValue)
				Exit For	
			End If
		Next
	Else 
		Err.Raise 1, "Select From Drop Down", "Object is not a WebList"
	End If	
End Function

Public Function extractDataFromElement(ObjectProperty)
	On Error Resume Next
	Err.Clear
	If Not IsObject(eval(strObjectName)) Then
		'"Object " & strTestObject & "   Not Declared"
		stepErrDescription =  UCASE(strAction) & " - " & "Object " & strObjectName & " not found"
		StepResult = "Failed"
		Exit Function
	End If
	If  Not VerifyObjExist(strObjectName) Then
		stepErrDescription = UCASE(strAction) & " - " & "Object " & strObjectName & " not present "
		StepResult = "Failed"
		Exit Function
	End If
	Execute (strObjectName & ".Init")
	strObjectProp = Eval(strObjectName & ".GetROProperty(" & Chr(34) & ObjectProperty & Chr(34) & ")")
	strObjectProp = Trim(strObjectProp)
	extractDataFromElement = strObjectProp
	
End Function


Function Base64Encode(sText)
 Set oNode = CreateObject("Msxml2.DOMDocument.3.0").CreateElement("base64")
 oNode.dataType = "bin.base64"
 oNode.nodeTypedValue =Stream_StringToBinary(sText)
 Base64Encode = oNode.text
 Set oNode = Nothing
 Reporter.ReportEvent micDone, "Base64Encode", "Base64Encode completed successfully"
End Function

Function Base64Decode(ByVal vCode)
 Set oNode = CreateObject("Msxml2.DOMDocument.3.0").CreateElement("base64")
 oNode.dataType = "bin.base64"
 oNode.text = vCode
 Base64Decode = Stream_BinaryToString(oNode.nodeTypedValue)
 Set oNode = Nothing
 Reporter.ReportEvent micDone, "Base64Decode", "Base64Decode completed successfully"
End Function

Function Stream_StringToBinary(Text)
 Set BinaryStream = CreateObject("ADODB.Stream")
 BinaryStream.Type = 2
' All Format =>  utf-16le - utf-8 - utf-16le
 BinaryStream.CharSet = "us-ascii"
 BinaryStream.Open
 BinaryStream.WriteText Text
 BinaryStream.Position = 0
 BinaryStream.Type = 1
 BinaryStream.Position = 0
 Stream_StringToBinary = BinaryStream.Read
 Set BinaryStream = Nothing
End Function

Function Stream_BinaryToString(Binary)
 Set BinaryStream = CreateObject("ADODB.Stream")
 BinaryStream.Type = 1
 BinaryStream.Open
 BinaryStream.Write Binary
 BinaryStream.Position = 0
 BinaryStream.Type = 2
 ' All Format =>  utf-16le - utf-8 - utf-16le
 BinaryStream.CharSet = "utf-8"
 Stream_BinaryToString = BinaryStream.ReadText
 Set BinaryStream = Nothing
End Function


'
'Function Encrypt (data)
'    Dim objXML, objNode
'    Set objXML = CreateObject("MSXML2.DOMDocument")
'    Set objNode = objXML.createElement("base64")
'    objNode.DataType = "bin.base64"
'    objNode.nodeTypedValue = data
'    Encrypt = objNode.Text
'    Set objNode = Nothing
'    Set objXML = Nothing
'End Function
'
'Function Decrypt (encodedData)
'    Dim objXML, objNode
'    Set objXML = CreateObject("MSXML2.DOMDocument")
'    Set objNode = objXML.createElement("base64")
'    objNode.DataType = "bin.base64"
'    objNode.Text = encodedData
'    Decrypt = objNode.nodeTypedValue
'    Set objNode = Nothing
'    Set objXML = Nothing
'End Function


'Function decryptUsingCrypton(encryptedPassword)
'    ' Path to crypton.exe
'    Dim cryptonPath
'    cryptonPath = "C:\Program Files (x86)\Micro Focus\Unified Functional Testing\bin\CryptonApp.exe"
'    secondaryPath = "C:\Program Files\Micro Focus\Unified Functional Testing\bin\CryptonApp.exe"
'	If not createObject("Scripting.FileSystemObject").FileExists(cryptonPath) Then
'		cryptonPath = secondaryPath
'	End If
'    ' Command to decrypt the password
'    Dim decryptCommand
'    decryptCommand = cryptonPath & " -decrypt " & encryptedPassword
'
'    ' Run the command
'    Dim oShell, oExec, oOutput
'    Set oShell = CreateObject("WScript.Shell")
'    Set oExec = oShell.Exec(decryptCommand)
'
'    ' Capture the decrypted password from the output
'    Dim decryptedPassword
'    decryptedPassword = ""
'    Do While Not oExec.StdOut.AtEndOfStream
'        decryptedPassword = oExec.StdOut.ReadLine()
'    Loop
'
'    ' Clean up
'    Set oShell = Nothing
'    Set oExec = Nothing
'
'    ' Return the decrypted password
'    decryptUsingCrypton = decryptedPassword
'End Function

' Function to get browser version using special URLs like chrome://version, about:version, edge://version
Function GetBrowserVersion(browserName)
    Dim objBrowser, versionText, version, pageSource
    Set objBrowser = Nothing
    
    ' Initialize browser based on the browserName argument
    If LCase(browserName) = "chrome" Then
        Set objBrowser = Browser("creationtime:=0")
    ElseIf LCase(browserName) = "firefox" Then
        Set objBrowser = Browser("creationtime:=0")
    ElseIf LCase(browserName) = "ie" Then
        Set objBrowser = Browser("creationtime:=0")
    ElseIf LCase(browserName) = "edge" Then
        Set objBrowser = Browser("creationtime:=0")
    End If
    
    ' Open the special URL to get version
    If Not objBrowser Is Nothing Then
        If LCase(browserName) = "chrome" Then
            GetBrowserVersion = split(objBrowser.GetAllROProperties().Item(3).Value," ")(1)
            Exit Function
        ElseIf LCase(browserName) = "firefox" Then
            objBrowser.Navigate("about:version")
        ElseIf LCase(browserName) = "ie" Then
            objBrowser.Navigate("about:internet")
        ElseIf LCase(browserName) = "edge" Then
            objBrowser.Navigate("edge://version")
        End If
        
        ' Wait for the page to load
        objBrowser.Sync
        
        ' Extract the page source
        pageSource = objBrowser.PageSource
        
        ' Check for the version string based on the browser type
        If LCase(browserName) = "chrome" Then
            ' Extract the version from the chrome://version page
            versionText = GetTextBetween(pageSource, "Google Chrome", "<")
            version = Trim(versionText)
        ElseIf LCase(browserName) = "firefox" Then
            ' Extract the version from the about:version page
            versionText = GetTextBetween(pageSource, "Mozilla Firefox", "<")
            version = Trim(versionText)
        ElseIf LCase(browserName) = "ie" Then
            ' Extract the version from the about:internet page
            versionText = GetTextBetween(pageSource, "Internet Explorer", "<")
            version = Trim(versionText)
        ElseIf LCase(browserName) = "edge" Then
            ' Extract the version from the edge://version page
            versionText = GetTextBetween(pageSource, "Microsoft Edge", "<")
            version = Trim(versionText)
        End If
        GetBrowserVersion = version
    Else
        GetBrowserVersion = "Browser not supported."
    End If
End Function

' Helper function to extract text between two substrings
Function GetTextBetween(text, startText, endText)
    Dim startPos, endPos, result
    startPos = InStr(text, startText)
    endPos = InStr(startPos, text, endText)
    If startPos > 0 And endPos > 0 Then
        result = Mid(text, startPos + Len(startText), endPos - startPos - Len(startText))
        GetTextBetween = result
    Else
        GetTextBetween = ""
    End If
End Function

'Functional check the object exist or not- Implemented due to UFT taking longer to identify the obj and fail it
Function VerifyObjExist(strObjName)
	For Iterator = 1 To conExist Step 1
		check = Eval(strObjName & ".Exist(1)") 
		If (Not check)  and  (iterator < (conExist+1) ) Then
			Wait(1)
		else
		VerifyObjExist = true
		 Exit function
		End If
	Next
	VerifyObjExist = false
	
End Function

Public Function extractROPropertyAndSave()
	StepResult =""
	strParam = strParameters
	strObjectProperty = GetParameterValue(Split(trim(strParam),",")(0))
	strStoreObjectName = GetParameterValue(Split(trim(strParam),",")(1))
	objProperty = extractDataFromElement(strObjectProperty)
	If UCASE(StepResult)="FAILED" Then
		Exit Function
	End If
	
	inputparameters.Add strStoreObjectName,objProperty
	Print "objProperty :- " & objProperty & " STORED IN VARIABLE " & strStoreObjectName
	If Err.Number <> 0 Then
				stepErrDescription = UCASE(strAction) & " - " & "Failed to store Object property "&strObjectProperty & " in variable "& strStoreObjectName & "-" & Err.Description
				StepResult = "Failed"
				Exit Function
	Else
				stepActual=  UCASE(strAction) & " - " & "Successfully  stored Object property "&strObjectProperty & " in variable "& strStoreObjectName
				StepResult = "Passed"
				stepErrDescription = ""
	End If
End Function
