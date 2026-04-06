Public strApiTestName,strAPIIterationCount,actualAPIResult,stepAPIResult,API_Step_No
Public apifTestCaseName,apifIteration,apifRequestType,apifURLEndPoint,apifURLPath,apifRequestHeader_Accept,apifRequestHeader_Accept_Encoding,apifRequestHeader_UserName,apifRequestHeader_Password,apifAuthentication_URL,apifRequestHeader_Connection,apifRequest_PayLoad_Path,apifResponse_PayLoad_Path,apifAuthentication_ClientID, apifAuthentication_ClientSecret, apifAuthentication_GrantType, apifAuthenticationScope,apifRequest_Header_Host
Public apifTS_TestCaseName,apifTS_SNo,apifTS_StepName,apifTS_StepDescription,apifTS_ExpectedResult,apifTS_StatusVerification,apifTS_Param1,apifTS_Param2,apifTS_Key,apifTS_Value
Public apifTS_PreviousExpectedResult, apifTS_PreviousStepDescription, apifTS_PreviousStepName
Public strAPIAction, accessToken, apifApplicationType
Public xmlhttp, url, URLPath, TCStep, SkipNextAPIStep, MandatoryStep
Public strQueryParameters, RequestPayload, ResponsePayLoad
Public sc, APIExecutionStartTime, APIExecutionEndTime, StepAPIEndTime, StepAPIStartTime, stepAPIDuration
Public Function APITest()
	On Error Resume Next 
	Err.Clear
	TCStep = TCStep+1
	Environment.Value("TCStep") = "ACTStep" & TCStep
	API_Tests_Executed = True
	SkipNextAPIStep = False
	checkAPIConfig= checkSheetPresent(Environment("DRIVER_FILE"), "APIConfig")
	checkAPITestScript= checkSheetPresent(Environment("DRIVER_FILE"), "APITestScript")
	if Not CreateAPIFolderStructer Then
			Reporter.ReportEvent micFail, "Check API Requests Folder Structure", "API Requests Folder Structure does not exist"
			mstepResult = "Failed"
			stepErrDescription = "API Requests Folder Structure does not exist"
			Exit Function
	Elseif Not checkAPIConfig Then
			Reporter.ReportEvent micFail, "Check driver File for APIConfig", "APIConfig Sheet DoesNot exist in driverFile"
			mstepResult = "Failed"
			stepErrDescription = "APIConfig Sheet DoesNot exist in driverFile"
			Exit Function
	Elseif Not checkAPITestScript Then
			Reporter.ReportEvent micFail, "Check driver File for APITestScript", "APITestScript Sheet DoesNot exist in driverFile"
			mstepResult = "Failed"
			stepErrDescription = "APITestScript Sheet DoesNot exist in driverFile"
			Exit Function
	ElseIf Not InitializeAPIReport()  Then
			Reporter.ReportEvent micFail, "Initialize API Result Report", "Initialize API Result Report Failed: " & stepAPIErrDescription
			mstepResult = "Failed"
			stepErrDescription = stepAPIErrDescription
			Exit Function
	Else 
	strApiTestName = GetParameterValue(TRIM(valParam1) )
	strAPIIterationCount = GetParameterValue(TRIM(valParam2) )
	Set SC = CreateObject("ScriptControl")
	SC.Language = "JScript"
	APIExecutionStartTime = cdbl(SC.Eval("new Date().getTime();"))
	SC.Reset
	Set SC = Nothing
	'APIExecutionStartTime = Now()
	Call QTP_API_Driver()
	'Call CompleteAPIReport()
	End if 
	
		Set SC = CreateObject("ScriptControl")
		SC.Language = "JScript"
		APIExecutionEndTime = cdbl(SC.Eval("new Date().getTime();"))
		SC.Reset
		Set SC = Nothing
	
	If Err.Number <> 0 Then
		stepErrDescription =  strApiTestName&" with iteration no "& strAPIIterationCount & " could not be completed" & Err.Description
		stepResult = "Failed"
		Reporter.ReportEvent micFail ,strTestName & "---" & strApiTestName, stepErrDescription
		Err.Clear
		Set SC = CreateObject("ScriptControl")
		SC.Language = "JScript"
		APIExecutionEndTime = cdbl(SC.Eval("new Date().getTime();"))
		SC.Reset
		Set SC = Nothing
		'APIExecutionEndTime = Now()
		Exit Function
	End  If
		
End  Function

Public Function QTP_API_Driver()
	strQueryParameters = ""
	RequestPayload = ""
	'ResponsePayLoad = ""
	Call InitializeObjXmlHTTP()

	data_range_apiConfig = "[APIConfig$]"
	sqlQueryApiConfig = "Select * from " & data_range_apiConfig & "WHERE TestCaseName = '"& strApiTestName & "'"
	
	blnOneRec = false
	Set RSApiConfig= CreateObject("ADODB.Recordset")
	ExeSQL DBConnection_DriverFile, sqlQueryApiConfig, RSApiConfig, blnOneRec, numRecCnt 
	
	If strExitTest="Yes" then
		Exit function
	End if
	
	Do While not RSApiConfig.EOF
	
		apifTestCaseName = isNullisEmptyCheck(RSApiConfig.Fields.Item("TestCaseName").Value)
		apifIteration = isNullisEmptyCheck(RSApiConfig.Fields.Item("Iteration").Value)
		apifRequestType = isNullisEmptyCheck(RSApiConfig.Fields.Item("RequestType").Value)
		apifApplicationType = isNullisEmptyCheck(RSApiConfig.Fields.Item("ApplicationType").Value)
		apifURLEndPoint = isNullisEmptyCheck(RSApiConfig.Fields.Item("URLEndPoint").Value)
		apifURLPath = isNullisEmptyCheck(RSApiConfig.Fields.Item("URLPath").Value)
		apifAuthentication_URL = isNullisEmptyCheck(RSApiConfig.Fields.Item("Authentication_URL").Value)
		apifAuthentication_ClientID =  isNullisEmptyCheck(RSApiConfig.Fields.Item("Authentication_ClientID").Value)
		apifAuthentication_ClientSecret =  isNullisEmptyCheck(RSApiConfig.Fields.Item("Authentication_ClientSecret").Value)
		apifAuthentication_GrantType =  isNullisEmptyCheck(RSApiConfig.Fields.Item("Authentication_GrantType").Value)
		apifAuthenticationScope =  isNullisEmptyCheck(RSApiConfig.Fields.Item("AuthenticationScope").Value)
		RSApiConfig.MoveNext
	Loop
	Set RSApiConfig = Nothing
	
	data_range_apiScript =  "[APITestScript$]"
	sqlQueryAPITestScript = "Select * from " & data_range_apiScript & "WHERE TestCaseName = '"& strApiTestName & "'"
	
	Set RSAPITestScript= CreateObject("ADODB.Recordset")
	ExeSQL DBConnection_DriverFile, sqlQueryAPITestScript, RSAPITestScript, blnOneRec, numRecCnt 
	'RSAPITestScript_RecordCount = RSAPITestScript.RecordCount
	Do While (not RSAPITestScript.EOF)' and (RSAPITestScript_RecordCount>1)
	 
		apifTS_TestCaseName = isNullisEmptyCheck(RSAPITestScript.Fields.Item("TestCaseName").Value)
		apifTS_SNo = isNullisEmptyCheck(RSAPITestScript.Fields.Item("SNo").Value)
		apifTS_StepName = isNullisEmptyCheck(RSAPITestScript.Fields.Item("Step Name").Value)
		apifTS_StepDescription = isNullisEmptyCheck(RSAPITestScript.Fields.Item("Step Description").Value)
		apifTS_ExpectedResult = isNullisEmptyCheck(RSAPITestScript.Fields.Item("Expected Result").Value)
		apifTS_StatusVerification = isNullisEmptyCheck(RSAPITestScript.Fields.Item("Expected Status").Value)
		apifTS_Param1 = isNullisEmptyCheck(RSAPITestScript.Fields.Item("Param1").Value)
		apifTS_Param2 = isNullisEmptyCheck(RSAPITestScript.Fields.Item("Param2").Value)
		apifTS_Key = isNullisEmptyCheck(RSAPITestScript.Fields.Item("Key").Value)
		apifTS_Value = isNullisEmptyCheck(RSAPITestScript.Fields.Item("Value").Value)
	
		if not (apifTS_StepName = "" or isnull(apifTS_StepName) or isempty(apifTS_StepName)) then
			apifTS_PreviousStepName = apifTS_StepName
			'Environment("apifTS_PreviousStepName").Value = apifTS_PreviousStepName
		end if
		
		if not (apifTS_StepDescription = "" or isnull(apifTS_StepDescription) or isempty(apifTS_StepDescription)) then
			apifTS_PreviousStepDescription = apifTS_StepDescription
			'Environment("apifTS_PreviousStepDescription").Value = apifTS_PreviousStepDescription
		end if
		
		if not (apifTS_ExpectedResult = "" or isnull(apifTS_ExpectedResult) or isempty(apifTS_ExpectedResult)) then
			apifTS_PreviousExpectedResult = apifTS_ExpectedResult
			'Environment("apifTS_PreviousExpectedResult").Value = apifTS_PreviousExpectedResult
		end if
		
		Call InitializeAPIStepExecution()
		stepAPIResult = "Passed"
		strAPIAction = apifTS_StepDescription
		If not SkipNextAPIStep Then
			Call performAPIAction(strAPIAction)
			If UCASE(stepAPIResult)="FAILED" Then
			mstepResult = "Failed"
			StepResult = "Failed"
			stepErrDescription = stepAPIErrDescription
			Err.Clear
			If MandatoryStep Then
				SkipNextAPIStep = TRUE
			End If
		End If
		Else
			If UCASE( StepResult)<>"FAILED" Then
				StepResult = "Skipped"
			Else
				StepResult = "Failed"
			End If
			stepAPIResult = "Skipped"
			stepAPIErrDescription = "Mandatory Step Failed"
			stepErrDescription = "Mandatory Step Failed"
		End If
		
		'-------------------------------------------------------
		call	FinishAPIStepExecution()
		RSAPITestScript.MoveNext
	Loop
	Set RSAPITestScript = Nothing
	If  isEmpty(mstepResult) and mstepResult<>""  Then
			mstepResult = "Passed"
			stepErrDescription = ""
			stepActual = strApiTestName&" with iteration no "& strAPIIterationCount & " Completed Successfully"
	End If
	If Err.Number <> 0 Then
			Reporter.ReportEvent micFail, "Execution of API Tests", "Execution of API Tests Failed: " & err.description
			mstepResult = "Failed"
			stepErrDescription = "Execution of API Tests Failed: " & err.description
			Err.Clear
			Exit Function
	End If
End Function

Public Function isNullisEmptyCheck(strObjectVal)
	If isNull(strObjectVal) or isEmpty(strObjectVal) or strObjectVal="" Then
			strObjectVal = ""
	End If
	isNullisEmptyCheck = strObjectVal
End Function

Sub InitializeAPIStepExecution()

	On Error Resume Next
	Err.Clear

	Call ClearErrors()
	If isNullisEmptyCheck(apifTS_StepName) = "" Then
		apifTS_StepName = apifTS_PreviousStepName
	End If
	If isNullisEmptyCheck(apifTS_StepDescription) = "" Then
		apifTS_StepDescription = apifTS_PreviousStepDescription
	End If 
	If isNullisEmptyCheck(apifTS_ExpectedResult) = "" Then
		apifTS_ExpectedResult = apifTS_PreviousExpectedResult
	End If 
	
	mAPIstepResult = "Passed"
	actualAPIResult  = ""
	actualAPIResult1  = ""
	stepAPIActual=""
	stepAPIDuration = 0
	Set SC = CreateObject("ScriptControl")
	SC.Language = "JScript"
	StepAPIStartTime = cdbl(SC.Eval("new Date().getTime();"))
	SC.Reset
	Set SC = Nothing
	defaultAPIstepExpected = ""
	
End Sub


Public Function performAPIAction(strAPIAction)
upperCaseStrAPIAction = Ucase(strAPIAction)
MandatoryStep = True
Select  case upperCaseStrAPIAction
	Case Ucase("Get Authorization Token")
		Call GetAuthorizationToken()
		
	Case Ucase("Open GET Request")
		Call OpenGETRequest()
	
	Case Ucase("Open POST Request")
		Call OpenPOSTRequest()
	
	Case Ucase("Open PUT Request")
		Call OpenPUTRequest()
	
	Case Ucase("Open PATCH Request")
		Call OpenPATCHRequest()
	
	Case Ucase("Open DELETE Request")
		Call OpenDELETERequest()
		
	Case Ucase("Setup Request Header")
		Call SetUpRequestHeader()
		
	Case Ucase("Build URL End Point")
		Call BuildURLEndPoint()
	
	Case Ucase("Set Query Parameter")
		Call SetQueryParameter()
		
	Case Ucase("Load Payload")
		If ucase(apifApplicationType)="JSON" Then
			Call LoadPayload()
		ElseIf ucase(apifApplicationType)="XML" Then
			Call LoadPayLoadXML()
		End If
		
	Case Ucase("Update PayLoad")
	If ucase(apifApplicationType)="JSON"  Then
		Call UpdatePayLoad()
	ElseIf ucase(apifApplicationType)="XML" Then
		call UpdateXMLPayLoad()
	End If
		
		
	Case Ucase("Attach Payload")
		Call AttachPayload()
		
	Case Ucase("Send Request")
		Call SendRequest()
		
	
	Case Ucase("Verify Response Payload Format")
		MandatoryStep = false
		Call VerifyResponsePayloadFormat()
	
	Case Ucase("Verify Response Key And Value")
		MandatoryStep = false
		
		If ucase(apifApplicationType)="JSON"  Then
			Call VerifyResponseKeyAndValue()
		ElseIf ucase(apifApplicationType)="XML" Then
			Call VerifyResponseKeyAndValueXML()
		End If
		
	Case Ucase("Update URLPath")
		Call UpdateURLPath()
	Case else
		MandatoryStep = false
		stepAPIErrDescription =  "API Step " & chr(34) & strAPIAction & chr(34) & " is invalid"
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
	End Select
	If Err.Number <> 0 Then
		stepAPIErrDescription =  "API Step " & chr(34) & strAPIAction & chr(34) & " is invalid" & Err.Description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail, strAPIAction, stepAPIErrDescription
		Err.Clear
		Exit Function
	End  If

End  Function

Public Function GetAuthorizationToken()
 	On Error Resume Next
	Err.Clear

	Dim objHttpAccessTocken, accessToken
	tokenApiEndPoint = apifAuthentication_URL

	Set objHttpAccessTocken = CreateObject("MSXML2.XMLHTTP")
	objHttpAccessTocken.open "POST", tokenApiEndPoint, False
	objHttpAccessTocken.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objHttpAccessTocken.send "grant_type="&apifAuthentication_GrantType&"&client_id="&apifAuthentication_ClientID&"&client_secret="& apifAuthentication_ClientSecret&"&scope="&apifAuthenticationScope
	'objHttpAccessTocken.send "grant_type=client_credentials&client_id=0oabwzsjxm7r9xnYG297&client_secret=drhqguRxQo1ggWThADYHszRVjI8siZs5tUkJVcGH&scope=mule.api.read"
	If objHttpAccessTocken.Status =  200 Then

		Dim jsonDict
		Set jsonDict = CreateObject ("Scripting.Dictionary")
		Dim jsonStr
		jsonStr = objHttpAccessTocken.responseText

		accessToken = GetValueFromJSON(jsonStr, "access_token")
		Print accessToken
		stepAPIActual = "Access Tocken Generated Successfully"
		stepAPIResult = "Passed"
		stepAPIErrDescription = ""
		actualAPIResult = stepAPIActual
		Reporter.ReportEvent micPass, strAPIAction , stepAPIActual&"--"&accessToken
		parameterName = isNullisEmptyCheck(GetParameterValue(apifTS_Param1))
		if (parameterName) <> "" then
			Execute Trim(parameterName & " = accessToken" )
			If Err.Number <> 0 Then
				stepAPIErrDescription = "Unable to Save Access token in parameter: " & parameterName & " due to " & err.description
				stepAPIResult = "Failed"
				Reporter.ReportEvent micFail, "Save Access token in parameter", stepAPIErrDescription
				actualAPIResult = stepAPIErrDescription
				err.clear
				Exit Function
			End  If
		End If
	Else
		Print "Error: " & objHttpAccessTocken.Status & "-" & objHttpAccessTocken.statusText
		stepAPIErrDescription = UCASE(strAPIAction) & " - " & "Authorization Tocken Could not beFetched" & "Error: " & objHttpAccessTocken.Status & "-" & objHttpAccessTocken.statusText
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail ,strAPIAction, stepAPIErrDescription
	End If
 	
 	If Err.Number <> 0 Then
		stepAPIErrDescription = UCASE(strAPIAction) & " - " & "Authorization Tocken Could not beFetched" & Err.Description & " with response code as " & objHttpAccessTocken.Status
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail ,strAPIAction, stepAPIErrDescription
		Err.Clear
		Exit Function
	End  If
End Function

Function OpenGETRequest()
 	On Error Resume Next
	Err.Clear
	xmlhttp.Open "GET", url, False
	actualAPIResult = strAPIAction & " completed successfully"
	Reporter.ReportEvent micPass ,strAPIAction, strAPIAction & " completed successfully"
	If Err.Number <> 0 Then
		stepAPIErrDescription = UCASE(strAPIAction) & " - " & "GET method unsuccessful" & Err.Description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail ,strAPIAction, stepAPIErrDescription
		Err.Clear
		Exit Function
	End  If
 End Function
 
 Function UpdateURLPath()
 	On Error Resume Next
	Err.Clear
	URLPath = apifURLPath
	strnewValue = GetParameterValue(TRIM(apifTS_Param1 ))
	apifURLPath = strnewValue
	actualAPIResult = strAPIAction & " completed successfully" & " - " & apifURLPath
	Reporter.ReportEvent micPass ,strAPIAction, strAPIAction & " completed successfully" & " - " & apifURLPath
	If Err.Number <> 0 Then
		stepAPIErrDescription = UCASE(strAPIAction) & " - " & "Update URLpath method unsuccessful" & Err.Description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail ,strAPIAction, stepAPIErrDescription
		Err.Clear
		Exit Function
	End  If
 End Function
 

 Function OpenPOSTRequest()
 	On Error Resume Next
	Err.Clear
	xmlhttp.Open "POST", url, False
	actualAPIResult = strAPIAction & " completed successfully"
	Reporter.ReportEvent micPass ,strAPIAction, strAPIAction & " completed successfully"
	If Err.Number <> 0 Then
		stepAPIErrDescription = UCASE(strAPIAction) & " - " & "POST method unsuccessful" & Err.Description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail ,strAPIAction, stepAPIErrDescription
		Err.Clear
		Exit Function
	End  If
 End Function
 
 Function OpenPUTRequest()
 	On Error Resume Next
	Err.Clear
	xmlhttp.Open "PUT", url, False
	actualAPIResult = strAPIAction & " completed successfully"
	Reporter.ReportEvent micPass ,strAPIAction, strAPIAction & " completed successfully"
	If Err.Number <> 0 Then
		stepAPIErrDescription = UCASE(strAPIAction) & " - " & "PUT method unsuccessful" & Err.Description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail ,strAPIAction, stepAPIErrDescription
		Err.Clear
		Exit Function
	End  If
 End Function
 
  Function OpenPATCHRequest()
 	On Error Resume Next
	Err.Clear
	xmlhttp.Open "PATCH", url, False
	actualAPIResult = strAPIAction & " completed successfully"
	Reporter.ReportEvent micPass ,strAPIAction, strAPIAction & " completed successfully"
	If Err.Number <> 0 Then
		stepAPIErrDescription = UCASE(strAPIAction) & " - " & "PATCH method unsuccessful" & Err.Description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail ,strAPIAction, stepAPIErrDescription
		Err.Clear
		Exit Function
	End  If
 End Function
 
  Function OpenDELETERequest()
 	On Error Resume Next
	Err.Clear
	
	xmlhttp.Open "DELETE", url, False
	actualAPIResult = strAPIAction & " completed successfully"
	Reporter.ReportEvent micPass ,strAPIAction, strAPIAction & " completed successfully"
	If Err.Number <> 0 Then
		stepAPIErrDescription = UCASE(strAPIAction) & " - " & "DELETE method unsuccessful" & Err.Description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail ,strAPIAction, stepAPIErrDescription
		Err.Clear
		Exit Function
	End  If
 End Function
 
  Function SendPUTMethod()
 	On Error Resume Next
	Err.Clear
	xmlhttp.Open "PUT", url, False
	Reporter.ReportEvent micPass ,strAPIAction, strAPIAction & " completed successfully"
	If Err.Number <> 0 Then
		stepAPIErrDescription = UCASE(strAPIAction) & " - " & "PUT method unsuccessful" & Err.Description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail ,strAPIAction, stepAPIErrDescription
		Err.Clear
		Exit Function
	End  If
 End Function
 
 Function SetUpRequestHeader()
 	On Error Resume Next
	Err.Clear
	strKey = GetParameterValue(TRIM(apifTS_Key) )
	strValue = GetParameterValue(TRIM(apifTS_Value) )
	
	If isNullisEmptyCheck(strKey)="" or isNullisEmptyCheck(strValue)="" Then
		err.raise 9999, "key and value empty check", "either key or value empty"
	End If
	
	xmlhttp.setRequestHeader strKey , strValue
	actualAPIResult = strAPIAction & " completed successfully" 
	Reporter.ReportEvent micPass ,strAPIAction, strAPIAction & " completed successfully" 
	If Err.Number <> 0 Then
		stepAPIErrDescription = UCASE(strAPIAction) & " - " & "Header Setup Could not be completed" & Err.Description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail ,strAPIAction, stepAPIErrDescription
		Err.Clear
		Exit Function
	End  If
 End Function
 
 Function BuildURLEndPoint()
 	On Error Resume Next
	Err.Clear
	If not Right(apifURLEndPoint,1) ="/"  Then
		apifURLEndPoint = apifURLEndPoint & "/"
	End If
	
	If strQueryParameters="" Then
		url = apifURLEndPoint & apifURLPath
	Else
		url = apifURLEndPoint & apifURLPath & "?" & strQueryParameters
	End If	
	
	Reporter.ReportEvent micPass ,strAPIAction, strAPIAction & " completed successfully -- " & url
	actualAPIResult = strAPIAction & " completed successfully"
	If Err.Number <> 0 Then
		stepAPIErrDescription = UCASE(strAPIAction) & " - " & "URL End point could not be prepared" & Err.Description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail ,strAPIAction, stepAPIErrDescription
		Err.Clear
		Exit Function
	End  If
 End Function

 
 Function SetQueryParameter()
 	On Error Resume Next
	Err.Clear
	
	strKey = GetParameterValue(TRIM(apifTS_Key) )
	strValue = GetParameterValue(TRIM(apifTS_Value) )
	If not ( strQueryParameters = "" ) Then
		strQueryParameters = strQueryParameters & "&" & strKey & "=" & strValue
	Else
		strQueryParameters = strKey & "=" & strValue
	End If
	stepAPIErrDescription = ""
	stepAPIResult = "Passed"
	actualAPIResult = strAPIAction & " completed successfully "
	Reporter.ReportEvent micPass ,strAPIAction, strAPIAction & " completed successfully -- " & strQueryParameters
	If Err.Number <> 0 Then
		stepAPIErrDescription = UCASE(strAPIAction) & " - " & "Query parameter could not be set" & Err.Description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail ,strAPIAction, stepAPIErrDescription
		Err.Clear
		Exit Function
	End  If
 End Function
 
 Function LoadPayload()
 	On Error Resume Next
	Err.Clear
	Dim inputText, jsonText
	
	inputText = GetParameterValue(apifTS_Param1)
	Set oFSO = CreateObject("Scripting.FileSystemObject")
    	Set oSC = CreateObject("MSScriptControl.ScriptControl")
	oSC.Language = "JScript"
		If not InStr(inputText,"{")>0 Then
			inputText =  DownloadResourceFromQC(inputText, "REQUEST") 
			If inputText="ERROR" Then
				err.raise 5000, "Download resource from ALM", "Error occured while downloading resource from ALM" & err.description
			End If
		End If
	
	If oFSO.FileExists(inputText) Then
       		Set oFile = oFSO.OpenTextFile(inputText, 1)
       		jsonText = oFile.ReadAll
        	oFile.Close
        	
        	If Err.Number <> 0 Then
			stepAPIErrDescription = "Retrieve JSON Text from File failed :- " & err.description
			stepAPIResult = "Failed"
			actualAPIResult = stepAPIErrDescription
			Reporter.ReportEvent micFail, "Retrieve JSON Text from File ", stepAPIErrDescription
			Set oFSO = Nothing
			Set oSC = Nothing
			Err.Clear
			Exit Function
		End  If
		Reporter.ReportEvent micPass ,"Retrieve JSON Text from File ", "Retrieve JSON Text from File : " & inputText & " is successful"
   	 Else
        	jsonText = inputText
    	End If
	
	Set jsonObject = oSC.Eval("(" + jsonText + ")")
	If Err.Number <> 0 Then
		stepAPIErrDescription = "JSON Text retrieved is not correct :- " & err.description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail, "Retrieve JSON Text", stepAPIErrDescription
		Set oFSO = Nothing
		Set oSC = Nothing
		Err.Clear
		Exit Function
	End  If
	RequestPayload = replace( jsonText, vbcrlf, "")
	stepAPIErrDescription = ""
	stepAPIResult = "Passed"
	actualAPIResult = strAPIAction & " completed successfully"
	Reporter.ReportEvent micPass ,"Retrieve JSON Text ", strAPIAction & " completed successfully"
	Reporter.ReportEvent micPass ,strAPIAction, strAPIAction & " completed successfully"
	If Err.Number <> 0 Then
		stepAPIErrDescription = UCASE(strAPIAction) & " - " & "Payload could not be loaded" & Err.Description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail ,strAPIAction, stepAPIErrDescription
		Set oFSO = Nothing
		Set oSC = Nothing
		Err.Clear
		Exit Function
	End  If
 End Function
 
 Function UpdatePayLoad()
 	On Error Resume Next
	Err.Clear
	
	strKey = GetParameterValue(TRIM(apifTS_Key) )
	strValue = GetParameterValue(TRIM(apifTS_Value) )
	Dim returnJSON
	returnJSON = UpdateAndRetunJSON(RequestPayload, strKey, strValue)
	If returnJSON="" or isEmpty(returnJSON) Then
		stepAPIErrDescription = UCASE(strAPIAction) & " - " & "payload could not be updated " & Err.Description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail ,strAPIAction, stepAPIErrDescription
		Err.Clear
		Exit Function
	End If
	RequestPayload = returnJSON
	stepAPIErrDescription=""
	stepAPIResult = "Passed"
	actualAPIResult = strAPIAction & " completed successfully"
	Reporter.ReportEvent micPass ,strAPIAction, strAPIAction & " completed successfully"
	If Err.Number <> 0 Then
		stepAPIErrDescription = UCASE(strAPIAction) & " - " & "payload could not be updated " & Err.Description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail ,strAPIAction, stepAPIErrDescription
		Err.Clear
		Exit Function
	End  If
 End Function
 
 Function AttachPayload()
 	On Error Resume Next
	Err.Clear
	actualAPIResult = strAPIAction & " completed successfully"
	Reporter.ReportEvent micPass ,strAPIAction, strAPIAction & " completed successfully"
	If Err.Number <> 0 Then
		stepAPIErrDescription = UCASE(strAPIAction) & " - " & "payload could not be attached to request" & Err.Description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail ,strAPIAction, stepAPIErrDescription
		Err.Clear
		Exit Function
	End  If
 End Function
 
 Function SendRequest()
 	On Error Resume Next
	Err.Clear
	ResponsePayLoad = ""
	expectedStatus =GetParameterValue(TRIM(isNullisEmptyCheck(apifTS_StatusVerification)))
	
	If not RequestPayload = "" Then
		If ucase(apifApplicationType)="JSON" Then
			xmlhttp.Send RequestPayload
		ElseIf ucase(apifApplicationType)="XML" Then
			xmlhttp.Send replace(CreateObject("Scripting.FileSystemObject").OpenTextFile(RequestPayload).ReadAll(), vbcrlf,"")
		End If
	Else
		xmlhttp.Send
	End If
	If not expectedStatus = "" Then
		If not (expectedStatus = xmlhttp.Status & "") Then
			stepAPIErrDescription = UCASE(strAPIAction) & " - " & "Send method unsuccessful" & " Expected status code " & expectedStatus & " dont match with the actual status code " & xmlhttp.Status & vbCrLf & xmlhttp.responseText 
			stepAPIResult = "Failed"
			actualAPIResult = UCASE(strAPIAction) & " - " & "Send method unsuccessful" & " Expected status code " & expectedStatus & " dont match with the actual status code " & xmlhttp.Status
			Reporter.ReportEvent micFail ,strAPIAction, stepAPIErrDescription
		Else
			stepAPIErrDescription =""
			stepAPIResult = "Passed"
			actualAPIResult = strAPIAction & " completed successfully" & " with expected status code " & xmlhttp.Status
			Reporter.ReportEvent micPass ,strAPIAction, strAPIAction & " completed successfully" & " with expected status code " & xmlhttp.Status & vbCrLf & xmlhttp.responseText
			
		End If
	
	Else
		stepAPIErrDescription =""
		stepAPIResult = "Passed"
		actualAPIResult = strAPIAction & " completed successfully" & " with status code " 
		Reporter.ReportEvent micPass ,strAPIAction, strAPIAction & " completed successfully" & " with status code " & xmlhttp.Status & vbCrLf & xmlhttp.responseText
		
	End If
	If Err.Number <> 0 Then
		stepAPIErrDescription = UCASE(strAPIAction) & " - " & "Send method unsuccessful" & Err.Description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail ,strAPIAction, stepAPIErrDescription
		Err.Clear
		Exit Function
	End  If
	If ucase(apifApplicationType)="JSON" Then
		ResponsePayLoad = Replace(xmlhttp.responseText, vblf, "")
	ElseIf ucase(apifApplicationType)="XML" Then
		ResponsePayLoad = Environment("API_RESULT_RESPONSE_FOLDER") & "\" & apifTS_TestCaseName & "_"& Replace(FormatDateTime(Date),"/","-") & "_" & Replace(FormatDateTime(Time),":","")&".xml"
	End If
	
	strContentType = xmlhttp.getResponseHeader("Content-Type")
	If InStr(lcase(strContentType), "application/json") > 0 Then
		call CreateFile(Environment("API_RESULT_RESPONSE_FOLDER") & "\" & apifTS_TestCaseName & "_"& Replace(FormatDateTime(Date),"/","-") & "_" & Replace(FormatDateTime(Time),":","")&".json",ResponsePayLoad )
	End If
	If InStr(lcase(strContentType), "application/xml") > 0  or InStr(lcase(strContentType), "text/xml") > 0 Then
		call CreateFile(ResponsePayLoad,xmlhttp.responseText)
	End If
	If Err.Number <> 0 Then
		stepAPIErrDescription = "Could not write file to destination:-" & filePath & "->" & err.description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail ,"Write file", stepAPIErrDescription
		Err.Clear
		Exit Function
End  If
 End Function

 Function CreateFile(filePath, strText )
 On Error Resume Next
 Err.clear
Dim objFSO, objFile
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Create a new file
Set objFile = objFSO.CreateTextFile(filePath, True)

'Write the JSON content to the file
objFile.Write strText

'Close the file
objFile.Close

'Clean up
Set objFile = Nothing
Set objFSO = Nothing

If Err.Number <> 0 Then
		stepAPIErrDescription = "Could not write file to destination:-" & filePath & "->" & err.description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIResult
		Reporter.ReportEvent micFail ,"Write file", stepAPIErrDescription
		Err.Clear
		Set objFile = Nothing
		Set objFSO = Nothing
		Exit Function
End  If

End Function

Function VerifyResponsePayloadFormat()
 	On Error Resume Next
	Err.Clear
	
	inputText = GetParameterValue(apifTS_Param1)
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	Set ScriptControl = CreateObject("MSScriptControl.ScriptControl")
	ScriptControl.Language = "JScript"
	Dim jsonText, strJSONFile
	
	If isNullisEmptyCheck(inputText)="" Then
			stepAPIErrDescription = "Response Payload Format provided is Blank"
			stepAPIResult = "Failed"
			actualAPIResult = stepAPIErrDescription
			Reporter.ReportEvent micFail, strAPIAction,  stepAPIErrDescription
			Set oFSO = Nothing
			Set ScriptControl = Nothing
			Err.Clear
			Exit Function
	End If
	
	If not (InStr(inputText,"{")>0  or isNullisEmptyCheck(inputText) ="") Then
			inputText =  DownloadResourceFromQC(inputText, "RESPONSE") 
			If inputText="ERROR" Then
				err.raise 5000, "Download resource from ALM", "Error occured while downloading resource from ALM" & err.description
			End If
		End If
	
	If oFSO.FileExists(inputText) Then
       		Set oFile = oFSO.OpenTextFile(inputText, 1)
       		jsonText = oFile.ReadAll
        	oFile.Close
        	
        If Err.Number <> 0 Then
			stepAPIErrDescription = "Retrieve JSON Text from File failed :- " & err.description
			stepAPIResult = "Failed"
			actualAPIResult = stepAPIErrDescription
			Reporter.ReportEvent micFail, "Retrieve JSON Text from File ", stepAPIErrDescription
			Set oFSO = Nothing
			Set ScriptControl = Nothing
			Err.Clear
			Exit Function
		End  If
		Reporter.ReportEvent micPass ,"Retrieve JSON Text from File ", "Retrieve JSON Text from File : " & inputText & " is successful"
   	Else
        	jsonText = inputText
    End If
	
	Set jsonObject = ScriptControl.Eval("(" + jsonText + ")")
	If Err.Number <> 0 Then
		stepAPIErrDescription = "JSON Text retrieved is not correct :- " & err.description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail, "Retrieve JSON Text", stepAPIErrDescription
		Set oFSO = Nothing
		Set ScriptControl = Nothing
		Err.Clear
		Exit Function
	End  If
	strJSONFile = Replace(jsonText, vblf, "")
	Set oFSO = Nothing
	ScriptControl.ExecuteStatement("var JSON = {}; JSON.parse = function(text) { return (new Function('return ' + text))(); }; JSON.stringify = function(obj) { var t = typeof (obj); if (t != 'object' || obj === null) { if (t == 'string') obj = '\""' + obj + '\""'; return String(obj); } else { var n, v, json = [], arr = (obj && obj.constructor == Array); for (n in obj) { v = obj[n]; t = typeof(v); if (t == 'string') v = '\""' + v + '\""'; else if (t == 'object' && v !== null) v = JSON.stringify(v); json.push((arr ? '' : '\""' + n + '\"":') + String(v)); } return (arr ? '[' : '{') + String(json) + (arr ? ']' : '}'); } };")
	ScriptControl.ExecuteStatement("JSON.removeValues = function(obj) { var t = typeof (obj); if (t != 'object' || obj === null) { if (t == 'string') obj = '\""' + obj + '\""'; return String(obj); } else { var n, v, json = [], arr = (obj && obj.constructor == Array); for (n in obj) { v = obj[n]; t = typeof(v); if (t == 'string') v = '\""' + v + '\""'; else if (t == 'object' && v !== null) v = JSON.stringify(v); json.push((arr ? '' : '\""' + n + '\"":') + String('\""\""')); } return (arr ? '[' : '{') + String(json) + (arr ? ']' : '}'); } };")
	
	ScriptControl.ExecuteStatement ("var response = JSON.parse('" & Replace(ResponsePayLoad, "'", "\'") & "');")
	ScriptControl.ExecuteStatement ("var fileJson = JSON.parse('" & Replace(strJSONFile, "'", "\'") & "');")
	
	ScriptControl.ExecuteStatement (" response = JSON.removeValues(response);")
	ScriptControl.ExecuteStatement (" fileJson = JSON.removeValues(fileJson);")	
	
	ScriptControl.ExecuteStatement (" response = JSON.parse(response);")
	ScriptControl.ExecuteStatement (" fileJson = JSON.parse(fileJson);")	
	
	ScriptControl.ExecuteStatement ("response = JSON.stringify(response);")
	ScriptControl.ExecuteStatement ("fileJson = JSON.stringify(fileJson);")
	ScriptControl.ExecuteStatement ("var areStructuresEqual = (response === fileJson);")
	areStructuresEqual = ScriptControl.Eval("areStructuresEqual")
	Reporter.ReportEvent micDone, " value removed api response", ScriptControl.Eval("response")
	Reporter.ReportEvent micDone, " value removed file json", ScriptControl.Eval("fileJson")
	If areStructuresEqual Then
    		stepAPIErrDescription =""
		stepAPIResult = "Passed"
		actualAPIResult = "Response payload format matches with the expected payload format"
		Reporter.ReportEvent micPass ,strAPIAction, "Response payload format matches with the expected payload format"
	Else
    		stepAPIErrDescription =  "Response payload format do not match with the expected payload format"
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail ,strAPIAction, stepAPIErrDescription
		Err.Clear
		Set ScriptControl = Nothing
		Exit Function
	End If
	actualAPIResult = strAPIAction & " completed successfully"
	Reporter.ReportEvent micPass ,strAPIAction, strAPIAction & " completed successfully"
	If Err.Number <> 0 Then
		stepAPIErrDescription = UCASE(strAPIAction) & " - " & "Verify Response Payload Format unsuccessful" & Err.Description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail ,strAPIAction, stepAPIErrDescription
		Err.Clear
		Set ScriptControl = Nothing
		Set oFSO = Nothing
		Exit Function
	End  If
	Set ScriptControl = Nothing
 End Function
 

    Function VerifyResponseKeyAndValue()
 	On Error Resume Next
	Err.Clear
	
	strKeyPath =GetParameterValue(TRIM(isNullisEmptyCheck(apifTS_Key)))
	Dim expectedValue
	expectedValue =GetParameterValue(TRIM(isNullisEmptyCheck(apifTS_Value)))
	
	Dim actualValue
	actualValue = GetValueFromJSON(xmlhttp.responseText, strKeyPath )
	
	If actualValue = expectedValue Then
		stepAPIErrDescription = ""
		stepAPIResult = "Passed"
		actualAPIResult = strAPIAction & " completed successfully-- Validation Passed " 
		Reporter.ReportEvent micPass ,strAPIAction, strAPIAction & " completed successfully-- Validation Passed: " & strKeyPath & " value is " & actualValue  
	Else
		stepAPIErrDescription = strAPIAction & " completed successfully-- Validation Failed: " & strKeyPath & " value is " & actualValue & " Expected: " & expectedValue   
		Reporter.ReportEvent micFail ,strAPIAction, stepAPIErrDescription
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
	End If	
	Reporter.ReportEvent micPass ,strAPIAction, strAPIAction & " completed successfully"
	If Err.Number <> 0 Then
		stepAPIErrDescription = UCASE(strAPIAction) & " - " & "Verify Response Key And Value unsuccessful" & Err.Description
		stepAPIResult = "Failed"
		Reporter.ReportEvent micFail ,strAPIAction, stepAPIErrDescription
		actualAPIResult = stepAPIErrDescription
		Err.Clear
		Exit Function
	End  If
 End Function
 
 Public Function InitializeObjXmlHTTP()
 	Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
 End Function

Function ParseJSON(jsontext)
	On Error Resume Next
	Dim regexPattern 
	regexPattern = """([^""]+)"":\s*""([^""]+(?:""[^""]+)*)"""
	Dim parsedData
	Dim sc
	Set sc = CreateObject("ScriptControl")
	sc.Language  = "JScript"
	Set parsedData = sc.Eval("(" + jsontext + ")")
	If Err.Number<>0 Then
		Set parsedData = nothing
	End If
	On Error GoTo 0
	Set ParseJSON = parsedData
End Function

Function GetValueFromJSON( jsonstring, strjsonpath)
On Error Resume Next
 Set jsonobject = ParseJSON(jsonstring)
 	Set sc = CreateObject("ScriptControl")
	sc.Language  = "JScript"
 	strJavaScriptFunction = "function getJOSONPathValue (jsonobject ){return jsonobject." &strjsonpath &"; }"
	sc.AddCode strJavaScriptFunction
	call sc.Run ("getJOSONPathValue", jsonobject ) 
	If err.number<>0 Then
		stepAPIErrDescription =  "Please check for the correct Json Path - " & Err.Description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail ,"Get jsonPath Value" , stepAPIErrDescription
		Exit Function
	End  If
	If isObject(sc.Run ("getJOSONPathValue", jsonobject ) ) Then
		stepAPIErrDescription =  "Please check for the correct Json Path - " & Err.Description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail ,"Get jsonPath Value" , "Please provide full Json path , current path provide is : " & strjsonpath
		Exit Function
	Else
		returnvalue = sc.Run ("getJOSONPathValue", jsonobject )
		If isEmpty(returnvalue) Then
			stepAPIErrDescription =  "Please check for the correct Json Path"
			stepAPIResult = "Failed"
			actualAPIResult = stepAPIErrDescription
			Reporter.ReportEvent micFail ,"Get jsonPath Value" , "Please provide correct Json path , current path provide is : " & strjsonpath
			Exit Function
		End If
		stepAPIErrDescription = ""
		stepAPIResult = "Passed"
		Reporter.ReportEvent micPass, "Value of Json path ", "Value of JSON path " & strjsonpath & " is found : " & returnvalue
		GetValueFromJSON =  returnvalue
	End If
	
 End  Function

Function GetValueFromJSONObject( jsonobject, strjsonpath)
On Error Resume Next
 	Set sc = CreateObject("ScriptControl")
	sc.Language  = "JScript"
 	strJavaScriptFunction = "function getJOSONPathValue (jsonobject ){return jsonobject." &strjsonpath &"; }"
	sc.AddCode strJavaScriptFunction
	call sc.Run ("getJOSONPathValue", jsonobject ) 
	If err.number<>0 Then
		stepAPIErrDescription =  "Please check for the correct Json Path - " & Err.Description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail ,"Get jsonPath Value" , stepAPIErrDescription
		Exit Function
	End  If
	If isObject(sc.Run ("getJOSONPathValue", jsonobject ) ) Then
		stepAPIErrDescription =  "Please check for the correct Json Path - " & Err.Description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail ,"Get jsonPath Value" , "Please provide full Json path , current path provide is : " & strjsonpath
		Exit Function
	Else
		returnvalue = sc.Run ("getJOSONPathValue", jsonobject )
		If isEmpty(returnvalue) Then
			stepAPIErrDescription =  "Please check for the correct Json Path"
			stepAPIResult = "Failed"
			actualAPIResult = stepAPIErrDescription
			Reporter.ReportEvent micFail ,"Get jsonPath Value" , "Please provide correct Json path , current path provide is : " & strjsonpath
			Exit Function
		End If
		stepAPIErrDescription = ""
		stepAPIResult = "Passed"
		Reporter.ReportEvent micPass, "Value of Json path ", "Value of JSON path " & Jsonpath & " is found : " & returnvalue
		GetValueFromJSONObject =  returnvalue
		
	End If
	
 End  Function
 
 Function UpdateAndRetunJSON(jsonText, jsonPath, updateValue )
On Error Resume Next
Err.clear

Dim returnValue

Set ScriptControl = CreateObject("MSScriptControl.ScriptControl")
ScriptControl.Language = "JScript"

ScriptControl.ExecuteStatement("var JSON = {}; JSON.parse = function(text) { return (new Function('return ' + text))(); }; JSON.stringify = function(obj) { var t = typeof (obj); if (t != 'object' || obj === null) { if (t == 'string') obj = '" & chr(34) & "' + obj + '" & chr(34) & "'; return String(obj); } else { var n, v, json = [], arr = (obj && obj.constructor == Array); for (n in obj) { v = obj[n]; t = typeof(v); if (t == 'string') v = '" & chr(34) & "' + v + '" & chr(34) & "'; else if (t == 'object' && v !== null) v = JSON.stringify(v); json.push((arr ? '' : '" & chr(34) & "' + n + '" & chr(34) & ":') + String(v)); } return (arr ? '[' : '{') + String(json) + (arr ? ']' : '}'); } };")
ScriptControl.ExecuteStatement("jsonObject = JSON.parse('" & jsonText & "');")

If isEmpty(GetValueFromJSON( jsonText, jsonPath)) Then
	Exit Function
End If

ScriptControl.ExecuteStatement("jsonObject." & jsonPath & " = '" & updateValue & "';")
returnValue  = ScriptControl.Eval("JSON.stringify(jsonObject)")
If Err.Number<>0 Then
	stepAPIErrDescription = "Update JSON path: '"& jsonPath & "' with value: '" & updateValue & "' failed " & err.description
	stepAPIResult = "Failed"
	actualAPIResult = stepAPIErrDescription
	Reporter.ReportEvent micFail ,"Update JSON path: '"& jsonPath & "'",  stepAPIErrDescription
	Set ScriptControl = Nothing
	Exit Function
End If
stepAPIErrDescription = ""
stepAPIResult = "Passed"
UpdateAndRetunJSON = returnValue
Set ScriptControl = Nothing
actualAPIResult = "Update JSON path with new Value completed successfully"
Reporter.ReportEvent micPass, "Update JSON path with new Value", "Provided JSON path '" & jsonPath & "' is updated with new value " ' & updateValue & "'"
End Function

Function CreateAPIFolderStructer()
	Set fso = CreateObject("Scripting.FileSystemObject")
	strTempFolder = "C:\TCOE"
	strProjectFolder = Environment.Value("FOLDERSTRUCTURE")
	strAPIFolder = strTempFolder & "\" & strProjectFolder & "\" & "API"
	If Not (fso.FolderExists(strAPIFolder)) Then
   		fso.CreateFolder ( strAPIFolder )	
		Reporter.ReportEvent micDone, "Checking API folder", "API folder created successfully"
   	End If
   	strRequests = strAPIFolder & "\" & "Requests"
   	If Not (fso.FolderExists(strRequests)) Then
   		fso.CreateFolder ( strRequests )	
		Reporter.ReportEvent micDone, "Checking Requests folder", "Requests folder created successfully"
   	End If
   	Environment.Value("API_REQUEST_FOLDER") = strRequests
   	strResponses = strAPIFolder & "\" & "Responses"
   	If Not (fso.FolderExists(strResponses)) Then
   		fso.CreateFolder ( strResponses )	
		Reporter.ReportEvent micDone, "Checking Responses folder", "Responses folder created successfully"
   	End If
   	Environment.Value("API_RESPONSE_FOLDER") = strResponses
   	strResultResponses = Environment("CURRENT_RESULTS_FOLDER") & "\" & "Responses"
   	If Not (fso.FolderExists(strResultResponses)) Then
   		fso.CreateFolder ( strResultResponses )	
		Reporter.ReportEvent micDone, "Checking Responses folder", "Responses folder created successfully"
   	End If
   	Environment.Value("API_RESULT_RESPONSE_FOLDER") = strResultResponses
   	strResultRequests = Environment("CURRENT_RESULTS_FOLDER") & "\" & "Requests"
   	If Not (fso.FolderExists(strResultRequests)) Then
   		fso.CreateFolder ( strResultRequests )	
		Reporter.ReportEvent micDone, "Checking Requests folder", "Requests folder created successfully"
   	End If
   	Environment.Value("API_RESULT_REQUEST_FOLDER") = strResultRequests
   	If Err.Number <> 0 Then
   		Reporter.ReportEvent micFail, "Checking Folder structure", "Folder structure is incorrect " & Err.Description
   		CreateAPIFolderStructer = false
   	End If
   	
   	Set fso = nothing
   	CreateAPIFolderStructer = True
End Function

Public Function InitializeAPIReport()
	On Error Resume Next
	
	stepAPIErrDescription = ""
	stepAPIResult = "Passed"
   	InitializeAPIReport = False
	If InitializeAPIReportCount>0 Then
		InitializeAPIReport= True
		Exit Function
	End If
	Set obFSO = CreateObject("Scripting.FileSystemObject")
	ResourcesFolder = Environment.Value("RESOURCES_FOLDER")
	currentResultsFolder = Environment.Value("CURRENT_RESULTS_FOLDER")
	
	statusAPI_HTMLRESULTS_TEMPLATE = DownloadResourceFromQC (Environment("API_HTMLRESULTS_TEMPLATE"), "RESOURCES")
	If UCASE(statusAPI_HTMLRESULTS_TEMPLATE)="ERROR" Then
		Reporter.ReportEvent micWarning , "API_HTMLRESULTS_TEMPLATE Download" , "Failed to download API_HTMLRESULTS_TEMPLATE"
		If Not (Environment("GenerateHTMLResult") = "YES" or Environment("GenerateHTMLResult") = "Y" )Then
			Exit Function
		End If
	End If
	
	
	tArray = split(statusAPI_HTMLRESULTS_TEMPLATE,"\")
	TemplateFile = Environment("RESOURCES_FOLDER") & "\" & tArray(UBOUND(tArray))
	If Not obFSO.FileExists(TemplateFile) Then
		stepAPIErrDescription = "ResultsHtmFile file is not available in the path : " & TemplateFile 
		Exit Function
	End If

	ResultsAPIHtmFile = Environment.Value("CURRENT_RESULTS_FOLDER") & "\" & tArray(UBOUND(tArray))
	obFSO.CopyFile TemplateFile, Environment.Value("CURRENT_RESULTS_FOLDER") & "\"
	If Not obFSO.FileExists(ResultsAPIHtmFile) Then
		stepAPIErrDescription = "ResultsAPIHtmFile file is not available in the path : " & ResultsAPIHtmFile 
		Exit Function
	End If
	
	strSQL = "Delete * from APITestResults"
	Set rs = DBConnection_Results.Execute(strSQL)
	If Err.Number <> 0 Then   
		    stepAPIErrDescription = "Unable to execute the SQL : " & strSQL & VbCrLf & Err.Description
			Exit Function
	End If 
	
	If err.number<>0 Then
		Reporter.ReportEvent micFail, "Initialize API Reports", "Failed to initialize API Reports" & err.description
		InitializeAPIReport = False
	End If
	
	Set rs = Nothing    
	Set obFSO = Nothing
	InitializeAPIReport = True
	APIExecutionStartTime = Now
	InitializeAPIReportCount=InitializeAPIReportCount+1
End  Function

'Function VerifyResponsePayloadFormat()
'    On Error Resume Next
'    Err.Clear
'
'    inputText = GetParameterValue(apifTS_Param1)
'    Set oFSO = CreateObject("Scripting.FileSystemObject")
'    Set ScriptControl = CreateObject("MSScriptControl.ScriptControl")
'    ScriptControl.Language = "JScript"
'    Dim jsonText, strJSONFile
'
'    If isNullisEmptyCheck(inputText)="" Then
'        stepAPIErrDescription = "Response Payload Format provided is Blank"
'        stepAPIResult = "Failed"
'        Reporter.ReportEvent micFail, strAPIAction, stepAPIErrDescription
'        Set oFSO = Nothing
'        Set ScriptControl = Nothing
'        Err.Clear
'        Exit Function
'    End If
'
'    If not (InStr(inputText,"{")>0  or isNullisEmptyCheck(inputText) ="") Then
'        inputText = DownloadResourceFromQC(inputText, "RESPONSE")
'        If inputText="ERROR" Then
'            err.raise 5000, "Download resource from ALM", "Error occured while downloading resource from ALM" & err.description
'        End If
'    End If
'
'    If oFSO.FileExists(inputText) Then
'        Set oFile = oFSO.OpenTextFile(inputText, 1)
'        jsonText = oFile.ReadAll
'        oFile.Close
'
'        If Err.Number <> 0 Then
'            stepAPIErrDescription = "Retrieve JSON Text from File failed :- " & err.description
'            stepAPIResult = "Failed"
'            Reporter.ReportEvent micFail, "Retrieve JSON Text from File ", stepAPIErrDescription
'            Set oFSO = Nothing
'            Set ScriptControl = Nothing
'            Err.Clear
'            Exit Function
'        End If
'        Reporter.ReportEvent micPass, "Retrieve JSON Text from File ", "Retrieve JSON Text from File : " & inputText & " is successful"
'    Else
'        jsonText = inputText
'    End If
'
'    ScriptControl.ExecuteStatement("var helper = {};helper.parse = function(text) {	return (new Function('return ' + text))();};helper.stringify = function(obj) {	var t = typeof (obj);	if (t != 'object' || obj === null) {		if (t == 'string') obj = '\" & chr(34) & "' + obj + '\" & chr(34) & "';		return String(obj);	}	else {		var n, v, json = [], arr = (obj && obj.constructor == Array);		for (n in obj) {			v = obj[n];			t = typeof (v);			if (t == 'string')				v = '\" & chr(34) & "' + v + '\" & chr(34) & "';			else if (t == 'object' && v !== null)				v = helper.stringify(v);			json.push((arr ? '' : '\" & chr(34) & "' + n + '\" & chr(34) & ":') + String(v));		}		return (arr ? '[' : '{') + String(json) + (arr ? ']' : '}');	}}; helper.getType = function (input) {  var obj = (typeof input === 'string') ? helper.parse(input) : input;  var type = typeof obj;  if (type !== 'object' || obj === null) {    return type;  } else if (Object.prototype.toString.call(obj) === '[object Array]') {    var arrayTypes = [];    for (var i = 0; i < obj.length; i++) {      arrayTypes.push(this.getType(helper.stringify(obj[i])));    }    var uniqueTypes = [];    for (var i = 0; i < arrayTypes.length; i++) {      if (uniqueTypes.indexOf(arrayTypes[i]) === -1) {        uniqueTypes.push(arrayTypes[i]);      }    }    return uniqueTypes.length > 1 ? " & chr(34) & "(" & chr(34) & " + uniqueTypes.join(' | ') + " & chr(34) & ")[]" & chr(34) & " : uniqueTypes[0] + " & chr(34) & "[]" & chr(34) & ";  } else {    var keys = Object.keys(obj);    var types = [];    for (var i = 0; i < keys.length; i++) {      types[i] = keys[i] + " & chr(34) & ": " & chr(34) & " + this.getType(helper.stringify(obj[keys[i]]));    }    return " & chr(34) & "{ " & chr(34) & " + types.join(', ') + " & chr(34) & " }" & chr(34) & ";  }};")
'   ' ScriptControl.ExecuteStatement("var JSON = {};")
'    'ScriptControl.ExecuteStatement("var fileJSON = '';var responseJSON='';var areStructuresEqual='';")
'    
'    'ScriptControl.ExecuteStatement("JSON.parse = function(text) { return (new Function('return ' + text))(); };")
'    'ScriptControl.ExecuteStatement("JSON.stringify = function(obj) { var t = typeof (obj); if (t != 'object' || obj === null) { if (t == 'string') obj = '\""' + obj + '\""'; return String(obj); } else { var n, v, json = [], arr = (obj && obj.constructor == Array); for (n in obj) { v = obj[n]; t = typeof(v); if (t == 'string') v = '\""' + v + '\""'; else if (t == 'object' && v !== null) v = JSON.stringify(v); json.push((arr ? '' : '\""' + n + '\"":') + String(v)); } return (arr ? '[' : '{') + String(json) + (arr ? ']' : '}'); } };")
'
'    'ScriptControl.ExecuteStatement("helper.getType = function (input) {  var obj = (typeof input === 'string') ? JSON.parse(input) : input;  var type = typeof obj;  if (type !== 'object' || obj === null) {    return type;  } else if (Object.prototype.toString.call(obj) === '[object Array]') {    var arrayTypes = [];    for (var i = 0; i < obj.length; i++) {      arrayTypes.push(this.getType(JSON.stringify(obj[i])));    }    var uniqueTypes = [];    for (var i = 0; i < arrayTypes.length; i++) {      if (uniqueTypes.indexOf(arrayTypes[i]) === -1) {        uniqueTypes.push(arrayTypes[i]);      }    }    return uniqueTypes.length > 1 ? ""("" + uniqueTypes.join(' | ') + "")[]"" : uniqueTypes[0] + ""[]"";  } else {    var keys = Object.keys(obj);    var types = [];    for (var i = 0; i < keys.length; i++) {      types[i] = keys[i] + "": "" + this.getType(JSON.stringify(obj[keys[i]]));    }    return ""{ "" + types.join(', ') + "" }"";  }};")
'
'    strJSONFile = Replace(jsonText, vblf, "")
'    Set oFSO = Nothing
'
'    ScriptControl.ExecuteStatement("fileJSON = helper.getType(" & strJSONFile & ");")
'    ScriptControl.ExecuteStatement("responseJSON = helper.getType(" & ResponsePayLoad & ");")
'    ScriptControl.ExecuteStatement("areStructuresEqual = (responseJSON === fileJSON);")
'
'    areStructuresEqual = ScriptControl.Eval("areStructuresEqual")
'
'    Reporter.ReportEvent micDone, " value removed api response", ScriptControl.Eval("response")
'    Reporter.ReportEvent micDone, " value removed file json", ScriptControl.Eval("fileJson")
'
'    If areStructuresEqual Then
'        stepAPIErrDescription =""
'        stepAPIResult = "Passed"
'        Reporter.ReportEvent micPass, strAPIAction, "Response payload format matches with the expected payload format"
'    Else
'        stepAPIErrDescription = "Response payload format do not match with the expected payload format"
'        stepAPIResult = "Failed"
'        Reporter.ReportEvent micFail, strAPIAction, stepAPIErrDescription
'        Err.Clear
'        Set ScriptControl = Nothing
'        Exit Function
'    End If
'
'    Reporter.ReportEvent micPass, strAPIAction, strAPIAction & " completed successfully"
'    If Err.Number <> 0 Then
'        stepAPIErrDescription = UCASE(strAPIAction) & " - " & "Verify Response Payload Format unsuccessful" & Err.Description
'        stepAPIResult = "Failed"
'        Reporter.ReportEvent micFail, strAPIAction, stepAPIErrDescription
'        Err.Clear
'        Set ScriptControl = Nothing
'        Set oFSO = Nothing
'        Exit Function
'    End If
'    Set ScriptControl = Nothing
'End Function

Function ReportAPIToQC()

			On Error Resume Next
			Err.Clear
			If Trim(Ucase( Environment.Value("TRIGGER_FROMQC"))) = "NO" Then
				Reporter.ReportEvent micDone,"Report to QC","TRIGGER_FROMQC is set to No"
				Exit Function
			End If 

			Dim myCurentRun,myStepFactory,myStepList,nStepKey

			Set myCurentRun = QCUtil.CurrentRun
			Set myStepFactory = myCurentRun.StepFactory
			myStepFactory.AddItem(apifTS_StepName)
			Set myStepList = myStepFactory.NewList("")
			nStepKey = myStepList.Count 'This sets the step count
			myStepList.Item(nStepKey).Field("ST_STATUS") = stepAPIResult
			myStepList.Item(nStepKey).Field("ST_DESCRIPTION") = apifTS_StepDescription
			myStepList.Item(nStepKey).Field("ST_EXPECTED") = apifTS_ExpectedResult
			myStepList.Item(nStepKey).Field("ST_ACTUAL") = actualAPIResult
			myStepList.Post


			Set myStepList = Nothing
			Set myStepFactory = Nothing
			Set myCurentRun = Nothing

	If err.Number<>0 Then
		Reporter.ReportEvent micFail, "Report result to QC","Unable to report to QC"
	End If

End Function

Sub FinishAPIStepExecution()

		On Error Resume Next
		If actualAPIResult = "" Then
			If Not stepAPIErrDescription = "" Then
				actualAPIResult = stepAPIErrDescription
				Reporter.ReportEvent micFail, valAutomationStepDescription & ":-" &UCASE(strAction), UCASE(strAction) & " - " & stepAPIErrDescription
			Else
				actualAPIResult = UCASE(strAction) & " - " & stepActual
				If blnQCPrintVal = False Then
				actualAPIResult = ""
				End if
			End If
		End If

		If actualAPIResult1 = "" Then
			If Not stepAPIErrDescription = "" Then
                actualAPIResult1 = stepAPIErrDescription
			Else
            	If blnQCPrintVal = True Then
            	actualAPIResult1 = valAutomationStepDescription & ":-" & UCASE(strAction) & "-" &stepActual
            	End If
			End If
		Else
			If Not stepAPIErrDescription = "" Then
				actualAPIResult1 = actualAPIResult1& vbLf & stepAPIErrDescription
			Else
               	If blnQCPrintVal = True Then
               	actualAPIResult1 =  actualAPIResult1& vbLf & UCASE(strAction) & "-" &stepActual
                End If
			End If
		End If
		
	

		If stepAPIResult <> "Passed" Then
				mAPIstepResult = stepAPIResult
				TestResult = StepResult
				'StepResult = stepAPIResult
		End If

		Call ErrorHandler()
		If  blnExitIteration Or blnExitTestCase Then
			If strAPIAction <> "CALL_API_TEST" Then
				Set SC = CreateObject("ScriptControl")
				SC.Language = "JScript"
				StepAPIEndTime = cdbl(SC.Eval("new Date().getTime();"))
				SC.Reset
				Set SC = Nothing
				stepAPIDuration =StepAPIEndTime-StepAPIStartTime
				Call ReportAPIResult()
				Call ReportAPIToQC()
				actualAPIResult  = ""
				actualAPIResult1  = ""
				Exit Sub
			End If
		End If

		If strAPIAction <> "CALL_API_TEST"  Then
				Set SC = CreateObject("ScriptControl")
				SC.Language = "JScript"
				StepAPIEndTime = cdbl(SC.Eval("new Date().getTime();"))
				SC.Reset
				Set SC = Nothing
				stepAPIDuration =StepAPIEndTime-StepAPIStartTime
					Call ReportAPIResult( )
					actualAPIResult  = ""
					Call ReportAPIToQC()
					actualAPIResult1  = ""
				Exit Sub
		End if		
End Sub


Public Sub ReportAPIResult()

	On Error Resume Next

	Dim strSQL,rs,FieldList,ValuesList

	Const adVarWChar = 202
	Const adSingle = 4
	Const adLockOptimistic = 3
	Const adOpenDynamic = 2


	FieldList = Array("APITestCaseName", "Iteration", "TestStepName","TestStepDescription", "ExpectedResult", "StepResult","FailureDescription","Duration","TCStep")
	ValuesList = Array(apifTS_TestCaseName,strAPIIterationCount,apifTS_StepName, mid(apifTS_StepDescription,1,254), mid(apifTS_ExpectedResult,1,254),stepAPIResult,mid(actualAPIResult,1,254),cStr(stepAPIDuration),Environment("TCStep"))

	Set rs = CreateObject("ADODB.Recordset")
	rs.Open "APITestResults", DBConnection_Results, adOpenDynamic, adLockOptimistic
	rs.MoveLast
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

Public Sub AddAPITestResultToHTML()

	On Error Resume Next
	Err.clear
	Dim strSQL,APIResultSet,strFile,objFSO
	API_Count_TC = 0
	API_Count_Passed = 0
	API_Count_Failed = 0
	API_Count_NotRun = 0
	'Opens the Execution Results.htm file

	'Gets the test execution results
	
	TCResultSet = "SELECT * FROM TestResults"
	Set TCResultSet = DBConnection_Results.Execute(TCResultSet)
	strActualStep = 1
	Do While not TCResultSet.eof
	strTCStep = "ACTStep" & strActualStep
	TestCaseName = TCResultSet.Fields.Item("Param1").Value
	Iteration = TCResultSet.Fields.Item("Param2").Value
	StepResult = TCResultSet.Fields.Item("StepResult").Value
	Duration = TCResultSet.Fields.Item("Duration").Value
			API_Count_TC = API_Count_TC + 1
			If StepResult = "Passed" Then
				API_Count_Passed = API_Count_Passed + 1
			ElseIf StepResult = "Failed" Then
				API_Count_Failed = API_Count_Failed + 1
			ElseIf StepResult = "Not Run" Then
				API_Count_NotRun = API_Count_NotRun + 1
			End If
		Call AddAPITestCaseHeaderResultsToHTML( API_Count_TC, TestCaseName,Iteration,StepResult,Duration)
		Call AddAPIStepHeaderToHTML()
		
		'APIstrSQL = "SELECT * FROM APITestResults Where APITestCaseName = '"& TestCaseName &"'  ORDER BY TestStepName ASC;"
		APIstrSQL =	"SELECT * FROM APITestResults WHERE APITestCaseName = '"& TestCaseName &"'  and TCStep ='" & strTCStep  & "' ORDER BY ID ASC;"
		Set APIResultSet = DBConnection_Results.Execute(APIstrSQL)
		API_Step_No = 0
		
		Do While not APIResultSet.eof
			API_Step_No = API_Step_No+1
			
			Dim APITestCaseName,APIIteration,APITestStepName,APITestStepDescription,APIExpectedResult,APIStepResult,APIFailureDescription,APIDuration
			APITestCaseName	=	APIResultSet.Fields.Item("APITestCaseName").Value
			
			APIIteration	=	APIResultSet.Fields.Item("Iteration").Value
			APITestStepName		=	APIResultSet.Fields.Item("TestStepName").Value
			APITestStepDescription =  APIResultSet.Fields.Item("TestStepDescription").Value
			APIExpectedResult		=	APIResultSet.Fields.Item("ExpectedResult").Value
			APIStepResult 	=	APIResultSet.Fields.Item("StepResult").Value
			APIFailureDescription 	=	APIResultSet.Fields.Item("FailureDescription").Value
			APIDuration 	=	APIResultSet.Fields.Item("Duration").Value
			APITCStep 	= 	APIResultSet.Fields.Item("TCStep").Value
			Call AddAPIStepResultsToHTML(APITestStepName,APITestStepDescription,APIExpectedResult,APIStepResult,APIFailureDescription,APIDuration )
			APIResultSet.MoveNext
		Loop
		TCResultSet.MoveNext 
		strActualStep = strActualStep+1
	Loop
	
	Call ReplaceAPISummary( API_Count_Passed,API_Count_Failed,API_Count_NotRun,API_Count_TC)
End Sub

Public API_Count_TC,API_Count_Passed,API_Count_Failed,API_Count_NotRun


Public Sub ReplaceAPISummary(API_Count_Passed,API_Count_Failed,API_Count_NotRun,API_Count_TC)
					On Error Resume Next

					Dim objFSO,objTextFile,strFile,strText
					Const ForReading = 1
					Const ForWriting = 2

					'Opens the execution Results.htm file
					strFile = Environment("CURRENT_RESULTS_FOLDER") & "\APIExecutionResults.htm"
   					set objFSO = CreateObject("Scripting.FileSystemObject")					
					Set objTextFile = objFSO.OpenTextFile(strFile, ForReading)
					strText = objTextFile.ReadAll

					strText = Replace(strText,"&amp;Host&amp;", Environment.Value("LocalHostName"))
					strText = Replace(strText,"&amp;Executed By&amp;", Environment.Value("UserName"))
          				strText = Replace(strText,"&amp;OS&amp;",Environment.Value("OS"))
					strText = Replace(strText,"&amp;Start Time&amp;", ExecutionStartTime)
					strText = Replace(strText,"&amp;End Time&amp;",ExecutionEndTime)
					strText = Replace(strText,"&amp;Total TC&amp;", API_Count_TC)
					strText = Replace(strText,"&amp;Passed&amp;", API_Count_Passed)
					strText = Replace(strText,"&amp;Failed&amp;", API_Count_Failed)
					strText = Replace(strText,"&amp;Not Run&amp;", API_Count_NotRun)
					 strText = Replace(strText,"<tr id = ""TestDetails""></tr>","")
					 strText = Replace(strText,"<tr id = ""TestStepDetails""></tr>","")
					objTextFile.Close
					Set objTextFile = Nothing

					Set objTextFile = objFSO.OpenTextFile(strFile, ForWriting)
					objTextFile.Write strText
					objTextFile.Close
					Set objTextFile = Nothing

					If Err.Number <> 0 Then   
	                         Reporter.ReportEvent micFail,"ReplaceSummary - Update the pass/fail summary in the API html report","Failed to update the pass/fail summary - " & Err.Description
							 Err.Clear
							 Exit Sub
					End If
End Sub

Public Sub AddAPITestCaseHeaderResultsToHTML(SLNo,TestCaseName,Iteration,StepResult,Duration)

	On Error Resume Next

	strFile = Environment("CURRENT_RESULTS_FOLDER") & "\APIExecutionResults.htm"
   	set objFSO = CreateObject("Scripting.FileSystemObject")					
	Set objTextFile = objFSO.OpenTextFile(strFile, ForReading)
	strText = objTextFile.ReadAll

	Dim objFSO,objTextFile,strFile,strText
	Const ForReading = 1
	Const ForWriting = 2
	
	
	Dim bgColor
	bgColor = "#7E87EE"


	Dim failColor

	If StepResult = "Passed" Then
		bgColor = "#d9ead3"
	ElseIf StepResult = "Failed" Then
		bgColor = "#f4cccc"
	ElseIf StepResult = "Not Run" Then
		bgColor = "#cecece"
	End If
	Dim updateText
	'Writes out the full test execution result
	updateText = "<tr id=""outerrow" & SLNo & """style=""background-color: " & bgColor & ";"">" & vbLf &_
		"<td id=""main" & SLNo & """ style=""vertical-align: middle; text-align: center; font-family: Calibri; padding: 0px; margin: 0px; background-color: " & bgColor & ";""> " &_
			"&nbsp;<a href=""javascript:void(0)"" onclick=""toggle(" & API_Count_TC & ", 'open')"" class=""style7"">+</a>&nbsp;</td>" & vbLf &_
		 "<td style=""vertical-align: middle; text-align: center; font-family: Calibri; padding: 0px; margin: 0px; background-color: " & bgColor & ";""> " &_
			SLNo & "</td>" & vbLf &_
		 "<td style=""vertical-align: middle; text-align: left; font-family: Calibri; padding: 0px; margin: 0px; background-color: " & bgColor & ";""> " &_
			 TestCaseName & "</td>" & vbLf &_
		"<td style=""vertical-align: middle; text-align: center; font-family: Calibri; padding: 0px; margin: 0px; background-color: " & bgColor & ";""> " &_
			Iteration & "</td>" & vbLf &_
		 "<td style=""vertical-align: middle; text-align: center; font-family: Calibri; padding: 0px; margin: 0px; background-color: " & bgColor & "; color: " & failColor &  "; ""> " &_
			StepResult & "</td>" & vbLf &_
		 "<td style=""vertical-align: middle; text-align: center; font-family: Calibri; padding: 0px; margin: 0px; background-color: " & bgColor & ";""> " &_
		  Duration & "</td>" & vbLf &_
		"</tr>" & vbLf & _
		"<tr id = ""innerRow" & SLNo & """ ></tr>" & vbLf &_
		"<tr id = ""TestDetails""></tr>" 
	strText = Replace(strText,"<tr id = ""TestDetails""></tr>" , updateText)
		

	objTextFile.Close
	Set objTextFile = Nothing

	Set objTextFile = objFSO.OpenTextFile(strFile, ForWriting)
	objTextFile.Write strText
	objTextFile.Close
	Set objTextFile = Nothing

	If Err.Number <> 0 Then   
	         Reporter.ReportEvent micFail,"ReplaceTestDetails - Update the pass/fail summary in the API html report","Failed to update the pass/fail summary - " & Err.Description
			 Err.Clear
			 Exit Sub
	End If
End Sub


Public Sub AddAPIStepResultsToHTML(APITestStepName,APITestStepDescription,APIExpectedResult,APIStepResult,APIFailureDescription,APIDuration)

	On Error Resume Next

	strFile = Environment("CURRENT_RESULTS_FOLDER") & "\APIExecutionResults.htm"
   	set objFSO = CreateObject("Scripting.FileSystemObject")					
	Set objTextFile = objFSO.OpenTextFile(strFile, ForReading)
	strText = objTextFile.ReadAll

	Dim objFSO,objTextFile,strFile,strText
	Const ForReading = 1
	Const ForWriting = 2
	
	
	Dim bgColor
	If API_Step_No Mod 2 = 1 Then
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
	Dim updateText
	'Writes out the full test execution result
	updateText = "<tr class=""subItem" & API_Count_TC & """ >" & vbLf &_
		"<td></td>" & vbLf &_
		 "<td style="" width : 100px ; vertical-align: left; text-align: left; font-family: Calibri; ""> " &_
			APITestStepName & "</td>" & vbLf &_
		 "<td style=""vertical-align: middle; text-align: left; font-family: Calibri; ""> " &_
			 APITestStepDescription & "</td>" & vbLf &_
		"<td style=""vertical-align: middle; text-align: left; font-family: Calibri; ""> " &_
			APIExpectedResult & "</td>" & vbLf &_
		 "<td style=""vertical-align: middle; text-align: center; font-family: Calibri;  color: " & failColor &"; ""> " &_
			APIStepResult & "</td>" & vbLf &_
		 "<td style=""vertical-align: middle; text-align: left; font-family: Calibri; ""> " &_
		  APIFailureDescription & "</td>" & vbLf &_
		"<td style=""vertical-align: middle; text-align: center; font-family: Calibri;  ""> " &_
		  APIDuration & "</td>" & vbLf &_
		"</tr>" & vbLf & "<tr id = ""TestStepDetails""></tr>" & vbLf
		strText = Replace(strText,"<tr id = ""TestStepDetails""></tr>", updateText)
		

	objTextFile.Close
	Set objTextFile = Nothing

	Set objTextFile = objFSO.OpenTextFile(strFile, ForWriting)
	objTextFile.Write strText
	objTextFile.Close
	Set objTextFile = Nothing

	If Err.Number <> 0 Then   
	         Reporter.ReportEvent micFail,"ReplaceTestDetails - Update the pass/fail summary in the API html report","Failed to update the pass/fail summary - " & Err.Description
			 Err.Clear
			 Exit Sub
	End If
End Sub



Public Sub AddAPIStepHeaderToHTML()

	On Error Resume Next

	strFile = Environment("CURRENT_RESULTS_FOLDER") & "\APIExecutionResults.htm"
   	set objFSO = CreateObject("Scripting.FileSystemObject")					
	Set objTextFile = objFSO.OpenTextFile(strFile, ForReading)
	strText = objTextFile.ReadAll

	Dim objFSO,objTextFile,strFile,strText
	Const ForReading = 1
	Const ForWriting = 2
	
	
	Dim bgColor
	bgColor = "#7E87EE"
	'strText = replace(strText,"&nbsp;TestDetails&nbsp;","")
	updateText = "<table id = ""innerTable" & API_Count_TC &  """style=""display: none;"">" &_
	"<thead>" &_
	"<tr class=""subItem" & API_Count_TC & """; "" ></tr>"&_
		"<tr class=""subItem" & API_Count_TC & """; "" > " &_
            "<th style=""vertical-align: middle; text-align: center; font-family: Calibri; padding: 0px; margin: 0px; background-color: #007CA6; color: #FFFFFF; font-weight: bold; font-size: x-large;"" class=""style6"">"&_
             "   &nbsp;&nbsp;&nbsp;</th>" &_
            "<th style="" width : 100px ;vertical-align: middle; text-align: center; font-family: Calibri; padding: 0px; margin: 0px; background-color: #007CA6; color: #FFFFFF; font-weight: bold; font-size: large;"" class=""style5"">"& _
               " Test Step name</th>" &_
			"<th style=""vertical-align: middle; text-align: center; font-family: Calibri; padding: 0px; margin: 0px; background-color: #007CA6; color: #FFFFFF; font-weight: bold; font-size: large;"" class=""style7"">" &_
                "Test Step Description</th>" &_
          "  <th style=""vertical-align: middle; text-align: center; font-family: Calibri; padding: 0px; margin: 0px; background-color: #007CA6; color: #FFFFFF; font-weight: bold; font-size: large;"" class=""style7"">" & _
		"		Expected Result</th>" & _
            "<th style=""vertical-align: middle; text-align: center; font-family: Calibri; padding: 0px; margin: 0px; background-color: #007CA6; color: #FFFFFF; font-weight: bold; font-size: large;"" class=""style4"">" & _
             "   	Result</th>" & _
             "<th style=""vertical-align: middle; text-align: center; font-family: Calibri; padding: 0px; margin: 0px; background-color: #007CA6; color: #FFFFFF; font-weight: bold; font-size: large;"" class=""style4"">" & _
             "   	Result Description</th>" & _
           " <th style=""vertical-align: middle; text-align: center; font-family: Calibri; padding: 0px; margin: 0px; background-color: #007CA6; color: #FFFFFF; font-weight: bold; font-size: large;"">" &_
	    "       	Duration (Millis)</th>" &_
       	    " </tr> " & vbLf &_
	    "</thead>" & vbLf &_
		"<tbody>"& vbLf &_
		"<tr id = ""TestStepDetails""></tr>" & vbLf &_
		"</tbody>" & vbLf &_
       	    "</table>"
	strText = Replace(strText,"<tr id = ""TestStepDetails""></tr>", updateText)
		
	objTextFile.Close
	Set objTextFile = Nothing

	Set objTextFile = objFSO.OpenTextFile(strFile, ForWriting)
	objTextFile.Write strText
	objTextFile.Close
	Set objTextFile = Nothing

	If Err.Number <> 0 Then   
	         Reporter.ReportEvent micFail,"ReplaceTestHeaderDetails - Update the test header in the API html report","Failed to update the test header - " & Err.Description
			 Err.Clear
			 Exit Sub
	End If
End Sub




Function LoadPayloadXML()
 	On Error Resume Next
	Err.Clear
	Dim inputText, XMLText, XMLFilePath
	
	inputText = GetParameterValue(apifTS_Param1)
	Set oFSO = CreateObject("Scripting.FileSystemObject")
    	Set oXMLFile = CreateObject("Msxml2.DOMDocument")
		
		If not InStr(inputText,"?xml version=")>0 Then
			inputText =  DownloadResourceFromQC(inputText, "REQUEST") 
			If inputText="ERROR" Then
				err.raise 5000, "Download resource from ALM", "Error occured while downloading resource from ALM" & err.description
			End If
		End If
	Dim argXMLFilePath
	argXMLFilePath = Environment("API_RESULT_REQUEST_FOLDER") & "\" & apifTS_TestCaseName & "_"& Replace(FormatDateTime(Date),"/","-") & "_" & Replace(FormatDateTime(Time),":","")&".xml"
	If not oFSO.FileExists(inputText) Then
		XMLText = inputText
		Set file = oFSO.CreateTextFile(argXMLFilePath, True)
		file.Write XMLText
		file.Close
		
		If Err.Number <> 0 Then
			stepAPIErrDescription = "Write XML File failed :- " & err.description
			stepAPIResult = "Failed"
			actualAPIResult = stepAPIErrDescription
			Reporter.ReportEvent micFail, "Write XML File failed ", stepAPIErrDescription
			Set oFSO = Nothing
			Err.Clear
			Exit Function
		End  If
		inputText = argXMLFilePath
		XMLFilePath = inputText
	End If
	
	If oFSO.FileExists(inputText) Then
       		'Verify valid xml file
		oXMLFile.Load(inputText)
		oXMLFile.async = False
		If oXMLFile.parseError.reason<>"" Then
				err.raise 5000, "Verify XML File", "Malformed XML File -" & oXMLFile.parseError.reason
		End If
		
		oFSO.CopyFile inputText, argXMLFilePath
        If Err.Number <> 0 Then
			stepAPIErrDescription = "Retrieve XML File failed :- " & err.description
			stepAPIResult = "Failed"
			actualAPIResult = stepAPIErrDescription
			Reporter.ReportEvent micFail, "Retrieve JSON Text from File ", stepAPIErrDescription
			Set oFSO = Nothing
			Err.Clear
			Exit Function
		End  If
		XMLFilePath = argXMLFilePath
		Reporter.ReportEvent micPass ,"Retrieve XML File ", "Retrieve XML File : " & XMLFilePath & " is successful"	
    End If
	RequestPayload = XMLFilePath
	stepAPIErrDescription = ""
	stepAPIResult = "Passed"
	actualAPIResult = strAPIAction & " completed successfully"
	Reporter.ReportEvent micPass ,"Retrieve JSON Text ", strAPIAction & " completed successfully"
	Reporter.ReportEvent micPass ,strAPIAction, strAPIAction & " completed successfully"
	If Err.Number <> 0 Then
		stepAPIErrDescription = UCASE(strAPIAction) & " - " & "Payload could not be loaded" & Err.Description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail ,strAPIAction, stepAPIErrDescription
		Set oFSO = Nothing
		Set oSC = Nothing
		Set oXMLFile = Nothing
		Err.Clear
		Exit Function
	End  If
 End Function
 
 Function UpdateXMLPayLoad()
 	On Error Resume Next
	Err.Clear
	
	strKey = GetParameterValue(TRIM(apifTS_Key) )
	strValue = GetParameterValue(TRIM(apifTS_Value) )
	Dim returnXML
	
	Set oXMLFile = CreateObject("Msxml2.DOMDocument")
	oXMLFile.Load(RequestPayload)
	oXMLFile.async = False
	originalExpression = strKey
	Set tempXML = oXMLFile
	Set nodes = oXMLFile.selectNodes (strKey)
	If InStr(strKey,  "[" )  Then
		haveMultipleNode = TRUE
	End If
	while InStr(strKey,  "[" ) 
		InitialNodes = SPLIT(strKey, "[")
		strNode = replace(replace(InitialNodes(0),"(",""),")","")
		While not LEFT(strNode, 2) = "//"
			strNode = "/" & strNode
		Wend
		position = LEFT(InitialNodes(1),1)
		bracketPosition = instr(strKey,"[")
		partAfterBracket = Mid(strKey,bracketPosition+1)
		bracketPosition = instr(strKey,"]")
		partAfterBracket = Mid(strKey,bracketPosition+1)
		Set nodes = tempXML.selectNodes (strNode)
		If nodes.length=0 or cint(position) > nodes.length Then
			err.raise 5000, "Node Validation", "XML path is incorrect or path does not exist - " & strKey & "----" & strNode
		End  If
		Set nodes = nodes.item(position)
		strKey = partAfterBracket
		Set tempXML = nodes
	Wend
	If haveMultipleNode Then
		nodes.text = strValue
	Else
		If nodes.length=0  Then
			err.raise 5000, "Node Validation", "XML path is incorrect or path does not exist - " & strKey & "----" & strNode
		End  If
		nodes.item(0).text = strValue
	End If
	
	oXMLFile.save(RequestPayload)
	
	stepAPIErrDescription=""
	stepAPIResult = "Passed"
	actualAPIResult = strAPIAction & " completed successfully"
	Reporter.ReportEvent micPass ,strAPIAction, strAPIAction & " completed successfully"
	If Err.Number <> 0 Then
		stepAPIErrDescription = UCASE(strAPIAction) & " - " & "payload could not be updated " & Err.Description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail ,strAPIAction, stepAPIErrDescription
		Err.Clear
		Set oXMLFile = Nothing
		Exit Function
	End  If
	Set oXMLFile = nothing
 End Function
 
 Function VerifyResponseKeyAndValueXML()
 	On Error Resume Next
	Err.Clear
	
	strKey = GetParameterValue(TRIM(apifTS_Key) )
	strValue = GetParameterValue(TRIM(apifTS_Value) )
	Dim returnXML
	
	Set oXMLFile = CreateObject("Msxml2.DOMDocument")
	oXMLFile.Load(RequestPayload)
	oXMLFile.async = False
	originalExpression = strKey
	Set tempXML = oXMLFile
	while InStr(strKey,  "[" ) 
		InitialNodes = SPLIT(strKey, "[")
		strNode = replace(replace(InitialNodes(0),"(",""),")","")
		While not LEFT(strNode, 2) = "//"
			strNode = "/" & strNode
		Wend
		position = LEFT(InitialNodes(1),1)
		bracketPosition = instr(strKey,"[")
		partAfterBracket = Mid(strKey,bracketPosition+1)
		bracketPosition = instr(strKey,"]")
		partAfterBracket = Mid(strKey,bracketPosition+1)
		Set nodes = tempXML.selectNodes (strNode)
		If nodes.length=0 or cint(position) > nodes.length Then
			err.raise 5000, "Node Validation", "XML path is incorrect or path does not exist - " & strKey & "----" & strNode
		End  If
		Set nodes = nodes.item(position)
		strKey = partAfterBracket
		Set tempXML = nodes
	Wend
	
	If nodes.text = strValue Then
		stepAPIErrDescription=""
		stepAPIResult = "Passed"
		actualAPIResult = strAPIAction & " completed successfully"
		Reporter.ReportEvent micPass ,strAPIAction, strAPIAction & " completed successfully"
	Else
		stepAPIErrDescription=strAPIAction & " failed due to mismatch " & "Expected : " & strValue & " <> Actual : " &  nodes.text
		stepAPIResult = "Failed"
		actualAPIResult = strAPIAction & " failed due to mismatch "
		Reporter.ReportEvent micPass ,strAPIAction, stepAPIErrDescription
	End If
	
	
	If Err.Number <> 0 Then
		stepAPIErrDescription = UCASE(strAPIAction) & " - " & "payload could not be verified " & Err.Description
		stepAPIResult = "Failed"
		actualAPIResult = stepAPIErrDescription
		Reporter.ReportEvent micFail ,strAPIAction, stepAPIErrDescription
		Err.Clear
		Set oXMLFile = nothing
		Exit Function
	End  If
	Set oXMLFile = nothing
 End Function
