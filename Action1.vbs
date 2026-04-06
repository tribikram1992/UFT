'#*************************************************************************************************************
'#  Project Name: 
'#  Test Script Name: 
'#  Test Case Name:
'#  Test Steps Covered : 
'#  Date of Creation: 
'#  Created By: 
'#  Purpose: 
'#  Initial State: 
'#  Parameters: 
'#  Assumptions:      
'#  Dependencies: 
'#  Outputs/Effects: 
'#  Modification History: 
'#  ID  Date  Changed By  Description 
'# *************************************************************************************************************
Set oShell = CreateObject("WScript.Shell")
command = "taskkill /F /IM excel.exe"
oShell.Run command,0,true
Set oShell = Nothing
Environment.Value("DRIVER_FILE") = "C:\Users\"& Environment("UserName") & "\Documents\DriverFile.xlsx"
Driverdatafile= ""
ExecuteTestCase Driverdatafile,"1" 
