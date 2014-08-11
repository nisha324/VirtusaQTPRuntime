' Copyright 2004 ThoughtWorks, Inc. Licensed under the Apache License, Version
' 2.0 (the "License"); you may not use this file except in compliance with the
' License. You may obtain a copy of the License at
' http://www.apache.org/licenses/LICENSE-2.0 Unless required by applicable law
' or agreed to in writing, software distributed under the License is
' distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
' KIND, either express or implied. See the License for the specific language
' governing permissions and limitations under the License.


Dim xmlDoc, logRoot, commandRoot, currentCommandRoot, overallCommadStatus, tcNode, stpCommand, tcCommands, tcName,tcDes, testSuiteNode, Config
Dim executionName, ProjectPath, newvtafreportPath

Function InitializeLogging
	strPath = Environment.Value("TestDir")
	Projectpath = strPath	
	createExectionReport()
	overallCommadStatus =  "Pass"
	Set xmlDoc = CreateObject("Microsoft.XMLDOM")
	Set testSuiteNode = xmlDoc.createElement("Testsuite")
	InitializeConfigFromFile(strPath & "\vtafRuntime\configs\config.txt")
	Set fso = CreateObject("Scripting.FileSystemObject") 
	dtMyDate = Day(Date) &"_"& MonthName(Month(Date),3) &"_"& Year(Date)
	strHTMLFilePath =strPath &"\vtafRuntime\testReports\HTML_Result_Log_VTAF3.0_TestSuite_"& dtMyDate & ".txt" 
'msgbox strHTMLFilePath
	If fso.FileExists(strHTMLFilePath) Then
		fso.DeleteFile strHTMLFilePath
	End If 

	setLogger "BeforeExecute<#>"& Environment.Value("UserName") &"<#>" &  Environment.Value("LocalHostName") &"<#>"& "Android" &"<#>EN_US<#>1366x768<#>"& Date & " - " & Time
End Function

Function startTestCase(tcNametxt, tcDescriptiontxt)
	 tcName = tcNametxt
	 tcDes = tcDescriptiontxt
	overallCommadStatus = "Pass" ' Kanchana Done initiation
	Call InsertIntoHTMLReport("startOfTestCase",tcName, tcDes, "", True, Config.Item("ReportPath"),False)
    Set tcNode = xmlDoc.createElement("TC")
	testSuiteNode.appendChild tcNode

	Set tcName = xmlDoc.createElement("Name")
	tcName.Text = tcNametxt
	tcNode.appendChild tcName

	Set tcDes = xmlDoc.createElement("Description")
	tcDes.Text = tcDescriptiontxt
	tcNode.appendChild tcDes

	Set tcCommands = xmlDoc.createElement("Commands")
	tcNode.appendChild tcCommands
	Err.clear
	setLogger "BeforeTestCase<#>"& tcNametxt 
End Function


Function startCommand (commandName, objectName)

   Set stpCommand  = xmlDoc.createElement("StepCommand")
   tcCommands.appendChild stpCommand

	Set stpName = xmlDoc.createElement("Name")
	stpName.Text = commandName
	stpCommand.appendChild stpName

	Set stpObject = xmlDoc.createElement("Object")
	stpObject.Text = objectName
	stpCommand.appendChild stpObject

End Function

Function afterStep(strStepName, strStepComments,  blnStatus, dtonlytime, vtafreportImagePath  )
'msgbox blnStatus
	If blnStatus = "True" Then
			setLogger "TestStep<#>"& dtonlytime  &"<#>Success<#>" & strStepName & "<#>" & "UNKNOWN" & "<#>" & "UNKNOWN" & "<#>" & strStepComments
	ElseIf blnStatus = "False" Then
			setLogger "TestStep<#>"& dtonlytime  &"<#>Error<#>" & strStepName & "<#>" & "UNKNOWN" & "<#>" & "UNKNOWN" & "<#>" &  strStepComments & "<#>" &  "No Stacktrace" & "<#>" &  vtafreportImagePath & "<#>" &  "UNKNOWN"
	End If


 
End Function

Function endCommand(status)
   Set stpStatus = xmlDoc.createElement("Status")
   stpStatus.Text = status
   stpCommand.appendChild stpStatus

	If  status = "Fail" Then ' Kanhana removed - " and  overallCommadStatus <> "Fail" "
		overallCommadStatus = "Fail" 
	End If

End Function

Function beforeTestSuite(nametestsuite)
	setLogger "BeforeTestSuite<#>"& nametestsuite
End Function

Function afterTestSuite
	setLogger "AfterTestSuite<#>0<#>"& "Success"
End Function

Function endTestCase
	Set tcStatus = xmlDoc.createElement("OverallStatus")
	   tcStatus.Text = overallCommadStatus
	   tcNode.appendChild tcStatus
Dim vartcresult 
	If overallCommadStatus = "Pass" Then
		Call InsertIntoHTMLReport("endOfTestCase", tcName, tcDes, "", True, Config.Item("ReportPath"),False)
		vartcresult = "Success"
	Elseif overallCommadStatus = "Fail" Then
		Call InsertIntoHTMLReport("endOfTestCase", tcName,  tcDes, "", False, Config.Item("ReportPath"),False)
		vartcresult = "Failed"
	End If
    CloseApp(15)
	setLogger "AfterTestCase<#>0<#>"& vartcresult

End Function


Function WriteLog(message, logType)
	Set objRoot = xmlDoc.createElement("Log")
	logRoot.appendChild objRoot
	Set objMessage = xmlDoc.createElement("Message")
	objMessage.Text = message
	Set objLogType = xmlDoc.createElement("Type")
	objLogType.Text = logType
	Set objLogTime = xmlDoc.createElement("LogTime")
	objLogTime.Text = Now
	objRoot.appendChild objMessage
	objRoot.appendChild objLogType
	objRoot.appendChild objLogTime
End Function

Function LogMessage(message)
   WriteLog message, "Info"
End Function

Function LogError(message)
  WriteLog message, "Error"
End Function

Function FinishLogging
	Call FinishCreatingHTMLReport(Config.Item("ReportPath"))
	setLogger "AfterExecute<#>UNKNOWN<#>Success<#>"& tcFailed &"<#>0<#>"& tcPassed &"<#>"& tcFailed &"<#>0"
	generateVTAFReport
End Function

Function createObjMap()
	Dim fso, Msg, parents
	Set fso = CreateObject("Scripting.FileSystemObject") 
	Set ResultFile = fso.OpenTextFile("C:\Users\drupasinghe\Desktop\Kanchana\ObjectMap\OBM1.vbs", 2, True) 
	ResultFile.WriteLine "Dim CommandObj" & vbnewline
	ResultFile.WriteLine "Function SendObject (ByVal Obj, ByVal prm)" & vbnewline
	ResultFile.WriteLine "Select Case Obj" & vbnewline
	Set Rep = CreateObject("Mercury.ObjectRepositoryUtil")
	Rep.Load "C:\Users\drupasinghe\Desktop\Kanchana\DanushkaTestingObjects_1.tsr"
	findchilds Null, Null, Rep,ResultFile
	ResultFile.WriteLine  "End Select" & Vbnewline
	ResultFile.WriteLine  "End Function"
End Function


Function findchilds (byval node,  byval parent, byval rep, byval ResultFile )
           Set children =Rep.GetChildren(node)
                 For  i = 0 to children.count-1 
					 If  children.item(i).getToProperty("micclass") = "Browser" then
						Msg = "Case " & """" & Rep.GetLogicalName(children.item(i))& """" & vbnewline
						Msg = Msg & "Set CommandObj =  " 
                        parents= parent &children.item(i).GetToProperty("Class Name")&"("""&Rep.GetLogicalName(children.item(i))&""")"
						ResultFile.WriteLine  Msg & parents & vbnewline
					else
					Msg = "Case " & """" & Rep.GetLogicalName(children.item(i)) & """" & vbnewline
					Msg = Msg & "Set CommandObj =  " 
					parents=  parent & "."&children.item(i).GetToProperty("Class Name")&"("""&Rep.GetLogicalName(children.item(i))&""")"
				   ResultFile.WriteLine   Msg & parents & vbnewline
					end if
					
                         findchilds children.item(i),parents,rep, ResultFile
					   
                 Next
			   
  Set Rep = Nothing
End Function


Function InitializeConfigFromFile(fileName)
	DIM fso
	SET fso = CreateObject ("Scripting.FileSystemObject")
	DIM strConfigLine,fConFile,EqualSignPosition, strLen, VariableName, VariableValue
	SET fConFile = fso.OpenTextFile(fileName)
	Set Config = CreateObject("Scripting.Dictionary")

	DIM ReportPath, LogPath
	Set Config = CreateObject("Scripting.Dictionary")
	Config.Add "ReportPath", ProjectPath & "\vtafRuntime\testReports"
	Config.Add "LogPath", ProjectPath & "\vtafRuntime\testReports"

	fConFile.Close

End Function



