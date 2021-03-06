'QTP web runtime - util functions
'28-APR-2015

Dim xmlDoc, logRoot, commandRoot, currentCommandRoot, overallCommadStatus, tcNode, stpCommand, tcCommands, tcName,tcDes, testSuiteNode, Config
Dim executionName, ProjectPath, newvtafreportPath,App
Dim ErrorNO,CountPass,CountFail,CountErr,StartTime,EndTime,ExecutionTime
Dim instances,isALM


Function InitializeLogging
	ErrorNO = 0
	CountPass=0
	CountFail=0
	CountErr=0
	isALM=False
    Reporter.Filter = rfDisableAll
	Print "[INFO] Start Execution "
	strPath=Environment.Value("TestDir")
	strTestName="_"&Environment.Value("TestName")&"_"
	If NOT(isLocalProject) Then
		getResourcesFormALM()
		InitializeConfigFromFile(strPath & "\vtafRuntime\configs\"&strTestName&"config.txt")
		UpdateFileNames strPath & "\vtafRuntime\configs\",strTestName,False
	    UpdateFileNames strPath & "\vtafRuntime\configs\ReportTemplete\",strTestName,False
	    UpdateFileNames strPath & "\vtafRuntime\testData\",strTestName,False
		isALM=True
	Else
		InitializeConfigFromFile(strPath & "\vtafRuntime\configs\config.txt")
	End If
	
	CleanUpReportDirectory
	Print "[INFO] Initialize Logging"
	StartTime=Timer
	Print "[INFO] Start Time-"&Time()
	Print "[INFO] Project Path"&strPath
	Projectpath = strPath	
	createExectionReport()
	Print "[INFO] Execution Report Created"
	overallCommadStatus =  "Pass"
	Set xmlDoc = CreateObject("Microsoft.XMLDOM")
	Set testSuiteNode = xmlDoc.createElement("Testsuite")
	
	ConfigureUFT()
	Set fso = CreateObject("Scripting.FileSystemObject") 
	dtMyDate = Day(Date) &"_"& MonthName(Month(Date),3) &"_"& Year(Date)
	strHTMLFilePath =strPath &"\vtafRuntime\testReports\HTML_Result_Log_VTAF3.0_TestSuite_"& dtMyDate & ".txt" 
	'msgbox strHTMLFilePath
	Set instances = CreateObject("Scripting.Dictionary")
	If fso.FileExists(strHTMLFilePath) Then
		fso.DeleteFile strHTMLFilePath
	End If 
	ImportTsr()	
	CleanRunningProcesses()
	 CloseApp(5)
	 Print "[INFO] Cleaned Running Processes"
	setLogger "BeforeExecute<#>"& Environment.Value("UserName") &"<#>" &  Environment.Value("LocalHostName") &"<#>"& "Microsoft Windows" &"<#>EN_US<#>1366x768<#>"& Date & " - " & Time
End Function

Function ImportTsr()
    On Error Resume Next 
    Dim FolderName, vFiles, objFSOx,qtApp,count
    count = 0
    TsrFolderName = ProjectPath &"\vtafRuntime\tsr"
    Set qtApp = CreateObject("QuickTest.Application")
    Set objFSOx = CreateObject("Scripting.FileSystemObject")
    Set FolderName = objFSOx.GetFolder(TsrFolderName)

    Set vFiles =FolderName.Files
    For each vFile in vFiles

        If InStr(1,vFile,".tsr",vbTextCompare) Then
        
            qtApp.Test.Actions(1).ObjectRepositories.Add vFile
            Print "[INFO] Associate tsr file: "&vFile.name
            count = count + 1
        End If
    Next
    Print "[INFO] Associated "&count&" tsr file(s)"
    
    On Error GoTo 0 
End Function


Function startTestCase(tcNametxt, tcDescriptiontxt)

	Print "[INFO] Testcase -"&tcNametxt&" is started"
	UpdateUFTReport false,true,"Start TestCase",tcNametxt,""
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
	ErrorNO = 0
	Err.Clear
	setLogger "BeforeTestCase<#>"& tcNametxt 
	
End Function


Function startCommand (commandName, objectName)
	Print "[INFO] Start Command - "&commandName
	If Not(objectName="" or isEmpty(objectName) or isNull(objectName)) Then
		Print "[INFO] ---Object - ["&objectName&"]"
	End If

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
			setLogger "TestStep<#>"& dtonlytime  &"<#>Error<#>" & strStepName & "<#>" & "UNKNOWN" & "<#>" & "UNKNOWN" & "<#>" &  strStepComments & "<#>" &  "" & "<#>" &  vtafreportImagePath & "<#>" &  VtafreportSmallImagePath
	End If


 
End Function

Function endCommand(status)
   Print "[INFO] End Command - |"&status&"|"
   Set stpStatus = xmlDoc.createElement("Status")
   stpStatus.Text = status
   stpCommand.appendChild stpStatus

	If  status = "Fail" Then ' Kanhana removed - " and  overallCommadStatus <> "Fail" "
		CountErr=CountErr+1
		overallCommadStatus = "Fail" 
	End If
	Err.Clear
End Function

Function beforeTestSuite(nametestsuite)
	setLogger "BeforeTestSuite<#>"& nametestsuite
	UpdateUFTReport false,true,"Start Test Suite",nametestsuite,""
End Function

Function afterTestSuite
	UpdateUFTReport false,true,"End Test Suite","",""
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
		CountPass=CountPass+1
		Print "[INFO] End TestCase - |Pass|"
		UpdateUFTReport false,true,"End TestCase","",""
	Elseif overallCommadStatus = "Fail" Then
		Call InsertIntoHTMLReport("endOfTestCase", tcName,  tcDes, "", False, Config.Item("ReportPath"),False)
		vartcresult = "Failed"
		CountFail=CountFail+1
		Print "[INFO] End TestCase - |Faild|"
		UpdateUFTReport false,false,"End TestCase","",""
	End If
    CloseApp(5)
    Print "[INFO] Killed Browser Process"
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
	Print "[INFO] Finish Logging"
	Call FinishCreatingHTMLReport(Config.Item("ReportPath"))
	EndTime=Timer-StartTime
	Print "[TIME] End Time - "&Time()
	ExecutionTime=GetExecutionTime(EndTime)
	setLogger "AfterExecute<#>"&ExecutionTime&"<#>Success<#>"& CountErr &"<#>0<#>"& CountPass &"<#>"& CountFail &"<#>0"
	Print "[REPO] Execution Time : "&ExecutionTime&" | Error Count: "&CountErr&" | Pass Count : "&CountPass&" | Fail Count : "&CountFail&" | Total : "&CountPass+CountFail
	Print "[INFO] Finishing creating VTAF report"
	generateVTAFReport
		If (Config.Item("ALM")) Then
			UploadReportToALM
	    End If
	
	Print "[INFO] End of Execution"
	Print "======================================"
	
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
	Print "[INFO] Initialize Config From File"
	DIM fso
	Dim arrFileLines()
	i = 0
	SET fso = CreateObject ("Scripting.FileSystemObject")
	DIM strConfigLine,fConFile,EqualSignPosition, strLen, VariableName, VariableValue
	SET fConFile = fso.OpenTextFile(fileName)
	Do Until fConFile.AtEndOfStream
		ReDim Preserve arrFileLines(i)
		arrFileLines(i) = fConFile.ReadLine
		i = i + 1
	Loop
	Set Config = CreateObject("Scripting.Dictionary")
	For Each strLine In arrFileLines
		words = Split(strLine, "=")
		Config.Add words(0), words(1)
		Print "[info] "&words(0)&"-"&words(1)
	Next
	DIM ReportPath, LogPath
	Config.Add "imgpath", ProjectPath & "\vtafRuntime\Images\"
	Config.Add "ReportPath", ProjectPath & "\vtafRuntime\testReports"
	Config.Add "LogPath", ProjectPath & "\vtafRuntime\testReports"
	Config.Add "ScreenshotPath",ProjectPath &"\ScreenShot\"
	fConFile.Close

End Function

Function CleanRunningProcesses()

	SET WshShell = CreateObject("WScript.Shell")
	SET oExec=WshShell.Exec("taskkill /F /IM Excel.exe")
	SET oExec= Nothing
	SET WshShell =Nothing
	
End Function


Public Function GetExecutionTime(Duration)
  Duration = Int(Duration) 
  HR = Duration \ 3600 
  remainder = Duration - HR * 3600
  MIN = remainder \ 60
  remainder = remainder - MIN * 60
  SEC = remainder
  'Prepend leading zeroes if necessary
  If Len(SEC) = 1 Then SEC = "0" & SEC
  If Len(MIN) = 1 Then MIN = "0" & MIN
  If Len(HR) =1 Then HR= "0" & HR
  GetExecutionTime = HR & ":" & MIN & ":" & SEC &" - H:M:S"
End Function


Function ConfigureUFT()
 On Error Resume next
	Set App = CreateObject("QuickTest.Application")

'geting execuion mode and set execution delay
	If "fast"=Lcase(Config.Item("ExecutionSpeed")) Then
		App.Options.Run.RunMode = "Fast"
	    Print "[UFT-] Run Mode : Fast"
	ElseIf "normal"=Lcase(Config.Item("ExecutionSpeed")) Then
	App.Options.Run.RunMode = "Normal"
	Print "[UFT-] Run Mode : Normal"
		If isNumeric(Config.Item("ExecutionDelay"))Then
			max=1000
			min=-1
			exDelay=Cint(Config.Item("ExecutionDelay"))
			If min<exDelay and max>exDelay  Then 
				App.Options.Run.StepExecutionDelay = exDelay
				Print "[UFT-] Execution Delay ["&exDelay&"]"
			Else
				Print "[UFT-] Invalid Execution Delay ["&exDelay&"] : Set to UFT Default"
			End If	
		End if
	Else
		Print "[UFT-] Invalid RunMode : Set to UFT Default"
	End If
'geting object sync time out
	If isNumeric(Config.Item("ExecutionSyncTimeOut")) Then
		max=120
		min=-1
		SyncTime=Cint(Config.Item("ExecutionSyncTimeOut"))
		If min<SyncTime and max>SyncTime  Then 
			App.Test.Settings.Run.ObjectSyncTimeOut = SyncTime
			Print "[UFT-] Execution Sync Time Out ["&SyncTime&"]"
		Else
			Print "[UFT-] Sync TimeOut Out of Range ["&SyncTime&"] : Set to UFT Default"
		End IF
	Else
		Print "[UFT-] Invalid Sync TimeOut : Set to UFT Default"
	End If
'getting iterations

If isNumeric(Config.Item("Iteration")) Then

	App.Test.Settings.Run.StartIteration = cint(Config.Item("Iteration"))
	Print "[UFT-] iteration(s) ["&cstr(Config.Item("Iteration"))&"]"
Else

	Print "[UFT-] Invalid iteration(s) : Set to UFT Default"

End if
If Err.Number<>0 Then
	Print "[EROR] "&Err.Description
End If
On Error goto 0
End Function


'this function returns resources from ALM
 
Function getResourcesFormALM()
	
Dim fso, folder
  Set fso = CreateObject("Scripting.FileSystemObject")
  ExecutionDiretory=Environment.Value("TestDir")
  ExecutionName=Environment.Value("TestName")
  
  rootFolder=ExecutionDiretory&"\vtafRuntime"
  configFolder=rootFolder&"\configs"
  reportTempleteFolder=configFolder&"\ReportTemplete"
  testDataFolder=rootFolder&"\testData"
  testReportsFolder=rootFolder&"\testReports"
  vtafReportFolder=testReportsFolder&"\vtafReport"
  
  arrDirectories=Array(rootFolder,configFolder,reportTempleteFolder,testDataFolder,testReportsFolder,vtafReportFolder)
 Print "[INFO] Starting to create local directories"
For Each directory in arrDirectories
	directorypath=split(directory,ExecutionDiretory)
	If NOT(fso.FolderExists(directory)) Then
	   	Set folder = fso.CreateFolder(directory)
	   	Print "[INFO] ["&directorypath(1)&"] is created"
	Else
	 	Print "[WERN] ["&directorypath(1)&"] is already exist"
	End If

next
 
 Print "[INFO] Directory Creating is Finished"
 
Set connection = QCUtil.QCConnection
Set resourceFolderFactory = connection.QCResourceFolderFactory

If Err.Number<>0 Then
	Print "[EROR] Fail to create connection with ALM/QC Server."
	Exit Function
End If

Set resourceFolder = resourceFolderFactory.NewList("")


For Iterator = 1 To (Cint(resourceFolder.Count))
	If Lcase("_"&ExecutionName&"_testData")=Lcase(resourceFolder.Item(Iterator).Name) then
		DownloadFiles resourceFolder.Item(Iterator).ID,testDataFolder,resourceFolderFactory
	ElseIf lcase("_"&ExecutionName&"_reporttemplete")=Lcase(resourceFolder.Item(Iterator).Name) Then
		DownloadFiles resourceFolder.Item(Iterator).ID,reportTempleteFolder,resourceFolderFactory
	ElseIf lcase("_"&ExecutionName&"_configs")=Lcase(resourceFolder.Item(Iterator).Name) Then
		DownloadFiles resourceFolder.Item(Iterator).ID,configFolder,resourceFolderFactory
	End If
Next

Print "[INFO] Getting Resources from ALM is Complted."
End Function


Function DownloadFiles(byref folderID,downloadPath,byref resourceFolderFactory)
	Set currentFolder=resourceFolderFactory.Item(FolderID)
	Set currentFile=currentFolder.QCResourceFactory
	Set files=currentFile.NewList("")
	'msgbox files.Count
	directorypath=split(downloadPath,Environment.Value("TestDir"))
	For x = 1 To (cint(files.Count))
		files.Item(x).DownloadResource downloadPath, True
		Print "[INFO] Downloading File ["&files.Item(x).Name&"] to local directory ["&directorypath(1)&"]"
	Next
End Function


Function UpdateFileNames(Directory,ProjectName,UpdateControl)
	Print "[INFO] ---Geeting File in "&Directory

	Set temp = CreateObject("Scripting.Dictionary")
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set folder = fso.GetFolder(Directory)
	Set files = folder.Files
	If files.Count<>0 Then
		

		For Each file in files
		
		'IF update control is set to true. it will append project name infront of every file.
		'Else remove project name if allrady appended.
		
			IF UpdateControl=True Then
				Print "[INFO] ------File Found : "&file.Name
				IF NOT(InStr(file.Name, ProjectName)>0) Then
					file.Name=ProjectName&file.Name
					Print "[INFO] ------File Name Updated: "&file.Name
				Else
					Print "[INFO] ------File Name allready Updated: "&file.Name
				End IF
			Else
				IF InStr(file.Name, ProjectName)>0 Then
					file.Name=Replace(file.Name,ProjectName,"")
					Print "[INFO] ------File Name Reverted Back: "&file.Name
				End IF
			End IF
		Next
	
	Else
		Print "[INFO] ------No File Found to Update"
	End If
	
	Print "[INFO] ---Finished Updating File Names in "&Directory
	
	Set files=Nothing 
	Set folder=Nothing
	Set fso=Nothing
	Set temp=Nothing
	
End Function



