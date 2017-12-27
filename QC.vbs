
Dim App, Test_Path, CurrentDirectory, qtTest, Config, qtRefVbs
Dim QtLib, QtObjectmap, QtBusinessLib, QtTestPlan, QtTestSuites, strTrustPath
Dim url,domain, project, username, password,abc,NewAction
Dim objFolder, colFiles, objFile
Const E_SCRIPT_TABLE = 1
Const E_EMPTY_SCRIPT = 2
Const E_SCRIPT_NOT_VALID = 3

Set App = CreateObject("QuickTest.Application")
dim fso: set fso = CreateObject("Scripting.FileSystemObject")
Set fileCollection = CreateObject("Scripting.Dictionary")
CurrentDirectory = fso.GetAbsolutePathName(".")
wscript.echo "[INFO] Opening UFT in Backgound.." 
App.Launch
App.Visible = true

wscript.echo "[INFO] Opening Project in directory ["&CurrentDirectory&"]" 
Test_Path = CurrentDirectory
App.Open Test_Path,False

'getting Enviroment Details
ExecutionDiretory=App.Test.Environment.Value("TestDir")
ExecutionName=App.Test.Environment.Value("TestName")

wscript.echo "[INFO] Generating Actions..." 
sFolder = "ScriptFiles"

For Each oFile In fso.GetFolder(sFolder).Files

	
   
   'strLines = ""
	
	Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile("ScriptFiles\" & oFile.Name, 1)
	

	Dim strLines
	strLines = ""
	Do While objFileToRead.AtEndOfStream = False
		strLines = strLines & vbCrLf & objFileToRead.ReadLine
	loop
	objFileToRead.Close
	Set objFileToRead = Nothing	

	
	ActionContent = strLines
	ActionDescr = "A new sample action for the test."
	ActionName = oFile.Name

	'Add a new action at the begining of the test
	Set NewAction = App.Test.AddNewAction(ActionName, ActionDescr, ActionContent, False, qtAtBegining)
	'Use the Load Script function to store the content of the first new action into the script array
	script = Load_Script(NewAction.GetScript())
	
	ActionContent = Save_Script(script)
	'Set new script source to the action
	scriptError = NewAction.ValidateScript(ActionContent)
	NewAction.SetScript ActionContent
	Set NewAction=Nothing


Next
wscript.echo "[INFO] Generating Actions Finished" 


'ALM Connection
wscript.echo "[INFO] Reading Configuration..." 
InitializePropFiles CurrentDirectory&"\vtafRuntime\configs\_"&ExecutionName&"_almConfig.inf","_"&ExecutionName&"_"

url=Config.Item("url")
domain=Config.Item("domain")
project=Config.Item("project")
username=Config.Item("username")
password=Config.Item("password")

wscript.echo "[INFO] Configurations "
wscript.echo "[INFO] --- RemoteURL ["&url&"]"
wscript.echo "[INFO] --- Domain ["&domain&"]"
wscript.echo "[INFO] --- Project ["&project&"]"
wscript.echo "[INFO] --- User ["&username&"]"

if NOT(password="") then
wscript.echo "[INFO] --- Password [CONFIGURED]"
Else
wscript.echo "[INFO] --- Password [NOTSET]"
End IF

'qtqcApp.TDConnection.Connect <QC Server path>, <Domain name that contains QC project>,
'<Project Name in QC you want to connect to>, <UserName>, <Password>,
' <Whether 'password is entered in encrypted or normal. Value is True for encrypted and FALSE for normal>


wscript.echo "[INFO] Checking QC/ALM Connection"

if App.TDConnection.IsConnected Then

	wscript.echo "[INFO] --- Existing connection found"
	wscript.echo "[INFO] --- Disconnecting current Connection"
	App.TDConnection.Disconnect
	wscript.echo "[INFO] --- Establishing new QC/ALM Connection"
	App.TDConnection.Connect url,domain, project, username, password, False
	wscript.echo "[INFO] --- Reconnected with new configurations"
	
 Else
 
	wscript.echo "[INFO] --- Establishing new QC/ALM Connection"
	App.TDConnection.Connect url,domain, project, username, password, False
	
	
 End If
 
 wscript.echo "[INFO] QC/ALM Connection successful"

 Set connection =App.TDConnection.TDOTA



wscript.echo "[INFO] ---Creating Root Folder [_"&ExecutionName&"_vtafRuntime]"
	  
rootFolder=ExecutionDiretory&"\vtafRuntime"
businessLibFolder=rootFolder&"\businessLib"
configFolder=rootFolder&"\configs"
reportTempleteFolder=configFolder&"\ReportTemplete"
libFolder=rootFolder&"\lib"
objectMapsFolder=rootFolder&"\objectMaps"
testDataFolder=rootFolder&"\testData"
testReportsFolder=rootFolder&"\testReports"
vtafReportFolder=testReportsFolder&"\vtafReport"
testSuites=rootFolder&"\testSuites"
arrDirectories=Array(rootFolder,businessLibFolder,configFolder,reportTempleteFolder,libFolder,objectMapsFolder,testDataFolder,testReportsFolder,vtafReportFolder,testSuites)
	
wscript.echo "[INFO] Starting to Scan local project file directories to find Resource Files.."
'Getting File informations 	in to a globle scope dictionery 
For Each directory in arrDirectories
	'IF NOT(testDataFolder=directory or reportTempleteFolder=directory) Then
		UpdateFileNames directory,"_"&ExecutionName&"_",true
	'End IF
	getFiles directory,fileCollection
	
next
wscript.echo "[INFO] Scan completed"
'--------------------------------------------------


	


'Creating Remote Root File Directory
wscript.echo "[INFO] Starting to create directories in QC/ALM"

wscript.echo "[INFO] ---Creating Root Folder [_"&ExecutionName&"_vtafRuntime]"
Set rootFolder = connection.QCResourceFolderFactory.Root
Set newFolder = rootFolder.QCResourceFolderFactory.AddItem(null)
newFolder.ParentId = rootFolder.ID
newFolder.Name = "_"&ExecutionName&"_vtafRuntime"
newFolder.Post()

wscript.echo "[INFO] ---Folder created successfully"
'Creating Remote Directories

rootFolderID=GetFolderID("_"&ExecutionName&"_vtafRuntime",connection)

'business libs
CreateFolder "_"&ExecutionName&"_businessLib",connection,rootFolderID
businessLibID=GetFolderID("_"&ExecutionName&"_businessLib",connection)
UploadResource businessLibID,"\businesslib", connection, fileCollection

'configs
CreateFolder "_"&ExecutionName&"_configs",connection,rootFolderID
configsID=GetFolderID("_"&ExecutionName&"_configs",connection)
UploadResource configsID,"\configs", connection, fileCollection

'configs\ReportTemplete
CreateFolder "_"&ExecutionName&"_reporttemplete",connection,configsID
reporttempleteID=GetFolderID("_"&ExecutionName&"_reporttemplete",connection)
UploadResource reporttempleteID,"\configs\reporttemplete", connection, fileCollection

'lib
CreateFolder "_"&ExecutionName&"_lib",connection,rootFolderID
libID=GetFolderID("_"&ExecutionName&"_lib",connection)
UploadResource libID,"\lib", connection, fileCollection

'objectMaps
CreateFolder "_"&ExecutionName&"_objectMaps",connection,rootFolderID
objectMapsID=GetFolderID("_"&ExecutionName&"_objectMaps",connection)
UploadResource objectMapsID,"\objectmaps", connection, fileCollection

'testData
CreateFolder "_"&ExecutionName&"_testData",connection,rootFolderID
testDataID=GetFolderID("_"&ExecutionName&"_testData",connection)
UploadResource testDataID,"\testdata", connection, fileCollection

'testReports
CreateFolder "_"&ExecutionName&"_testReports",connection,rootFolderID
testReportsID=GetFolderID("_"&ExecutionName&"_testReports",connection)
UploadResource testReportsID,"\testreports", connection, fileCollection

'testReports\vtafReport
CreateFolder "_"&ExecutionName&"_vtafReport",connection,testReportsID
vtafReportID=GetFolderID("_"&ExecutionName&"_vtafReport",connection)
UploadResource vtafReportID,"\testreports\vtafreport", connection, fileCollection

'testSuites
CreateFolder "_"&ExecutionName&"_testSuites",connection,rootFolderID
testSuitesID=GetFolderID("_"&ExecutionName&"_testSuites",connection)
UploadResource testSuitesID,"\testsuites", connection, fileCollection

'removing any assoicated function libraryes before save on remote loaction

App.Test.Settings.Resources.Libraries.RemoveAll

savePath="[QualityCenter] "&Config.Item("remoteProjectPath")&"\"&ExecutionName
wscript.echo "[INFO] Saving Project in Remote Location ["&savePath&"]"
App.Test.SaveAs savePath
wscript.echo "[INFO] Closing Local Project"
'app.Test.Close

're opening project from remote locaton

wscript.echo "[INFO] Reopening Remote Project"

Test_Path = CurrentDirectory
App.Open savePath,False
Set qtTest = App.Test


'Dclaring RemoteDirectories

remoteDiretory="[QC-RESOURCE];;Resources"
remoteRootFolder=remoteDiretory&"\vtafRuntime"
remoteBusinessLibFolder=remoteRootFolder&"\businessLib"
remoteLibFolder=remoteRootFolder&"\lib"
remoteObjectMapsFolder=remoteRootFolder&"\objectMaps"
remoteTestSuites=remoteRootFolder&"\testSuites"

'Associating Function Libs

AssociateFunctionLibrary "\businesslib",remoteBusinessLibFolder,fileCollection,App
AssociateFunctionLibrary "\lib",remoteLibFolder,fileCollection,App
AssociateFunctionLibrary "\objectmaps",remoteObjectMapsFolder,fileCollection,App
AssociateFunctionLibrary "\testsuites",remoteTestSuites,fileCollection,App
wscript.echo "[INFO] Saving Project in remote"
qtTest.Save
qtTest.Close
app.Quit

wscript.echo "[INFO] Reverting File Name and Resetting local Project.."
For Each directory in arrDirectories
	
	UpdateFileNames directory,"_"&ExecutionName&"_",false
	
	
next
wscript.echo "[INFO] Resetting completed"
wscript.echo "[INFO] ALM Upload Process Finished Successfully"




Function CreateFolder(FolderName,connection,parentFolderID)
	
	Set resourceFolderFactory = connection.QCResourceFolderFactory
	Set resourceFolder = resourceFolderFactory.Item(parentFolderID)
	Set newFolder = resourceFolder.QCResourceFolderFactory.AddItem(null)
	newFolder.ParentId = resourceFolder.ID
	newFolder.Name = FolderName
	newFolder.Post()
	Set newFolder=Nothing
	Set resourceFolder=Nothing
	Set resourceFolderFactory=Nothing
	
End Function



Function GetFolderID(FolderName, connection)
	'Dim GetCurrentFolderID,resourceFolderFactory,rootFolder
	
	wscript.echo "[INFO] --- Started to get the Folder ID of ["&FolderName&"] Form Remote Server"
	
	wscript.echo "[INFO] ------Refreshing Connection"	
	if App.TDConnection.IsConnected Then
		App.TDConnection.Disconnect
		App.TDConnection.Connect url,domain, project, username, password, False
	Else
		App.TDConnection.Connect url,domain, project, username, password, False
	End If
	
	wscript.echo "[INFO] ------Refreshing Succesfull"
	Set connection =App.TDConnection.TDOTA
	
	retry=40
	counter=0
	GetFolderID=null
	wscript.echo "[INFO] ------Requesting id with in ["&retry&"] retries "
	do while(isNull(GetFolderID) and counter<retry)
	Set resourceFolderFactory = connection.QCResourceFolderFactory
	Set rootFolder = resourceFolderFactory.NewList("")
	
	
	'getting  folder id
	For Iterator = 1 To (rootFolder.Count)
	
		If lcase(FolderName)=lcase(rootFolder.Item(Iterator).Name) Then
			WScript.Sleep 200
			GetFolderID=rootFolder.Item(Iterator).ID
			wscript.echo "[INFO] ------ID reterned sucessfully"
		    WScript.Sleep 200
			Exit For
		End If
	WScript.Sleep 1000
	wscript.echo "[INFO] ---------Wating for QC/ALM Responce..."
	Next
	
	if isNull(GetFolderID) then
		wscript.echo "[EROR] ------ID is not reterned sucessfully"
	End IF 
	
	
	loop
	
	Set rootFolder=Nothing
	Set resourceFolderFactory=Nothing
	'Set connection=Nothing
	
End Function

Function UploadResource(parentFolderID,key, connection, Dictionary)

 

  If Dictionary.Exists(key) Then
    
    	Set runner=Dictionary.Item(key)  
     	k = runner.Keys 
     	
   		For i = 0 To runner.Count - 1 
	       	resourceFile=k(i)
			wscript.echo "[INFO] ---Checking connection is avalable"
			if App.TDConnection.IsConnected Then
			wscript.echo "[INFO] ---Allrady a connection is avalable"
			App.TDConnection.Disconnect
 
			App.TDConnection.Connect url,domain, project, username, password, False

			Else
			App.TDConnection.Connect url,domain, project, username, password, False
			End If
			Set connection =App.TDConnection.TDOTA
			
			wscript.echo "[INFO] Start uploading resources" 
			Set resourceFolderFactory = connection.QCResourceFolderFactory
			Set resourceFolders = resourceFolderFactory.NewList("")
			Set resourceFolder = resourceFolderFactory.Item(parentFolderID)
			
			
	       	wscript.echo "[INFO] ---Uploading File ["&k(i)&"] in Directory ["&runner.Item(k(i))&"]"
	       	
			set ResourceFactory=resourceFolder.QCResourceFactory
			Set currResourceList = ResourceFactory.NewList("")
			Set resourceItem = resourceFactory.AddItem(null)
			'resourceItem.fileName = "_"&ExecutionName&"_"&resourceFile
			resourceItem.fileName = resourceFile
				wscript.echo "[INFO] ---Uploading Res Item Name ["&resourceFile&"] "
			resourceItem.Post
			'---------------------------------------------------
			'QcResourceName = "_"&ExecutionName&"_"&resourceFile
			QcResourceName = resourceFile
			wscript.echo "[INFO] ---Qc Resource Name ["&QcResourceName&"] "
			fileName = resourceFile
			wscript.echo "[INFO] ---fileName ["&fileName&"] "
			LocalFileDirectory = runner.Item(k(i))
			Set Resources = connection.QCResourceFactory
			Set CurrentResources =Resources.NewList("")
			  wscript.echo "[INFO] ------Getting resource ["&resourceFile&"] "
			 
			resourceCount = CurrentResources.Count
			retry=40
			counter=0
			resourceFound=false
			do while(resourceFound=false and counter<retry)
				For x = 1 To resourceCount
				 selectedResource = CurrentResources.Item(x).Name
				 wscript.echo "this is the if ---- "&selectedResource&"="&QcResourceName
				   If UCase(selectedResource) = UCase(QcResourceName) then
					Set newResource = CurrentResources.Item(x)
					resourceFound = True
				   end if
				 Next
				 WScript.Sleep 1000
				 wscript.echo "[INFO] ---------Wating for QC/ALM Responce..."
			loop
			 
			If (resourceFound=True) Then
				wscript.echo "[INFO] ---------Uploading Resource ["&fileName&"]"
				newResource.Filename = fileName 
				newResource.ResourceType = "Test Resource"
				newResource.Post
				newResource.UploadResource LocalFileDirectory, true
				wscript.echo "[INFO] ------Uploaded Succesfully"
			 Else
			  	wscript.echo "[EROR] ------Resource is not found Remote Location. Fail to Update"
			 End If

    	Next 
    	
    End If
	
	
	
	Set newResource=Nothing
	Set CurrentResources=Nothing
	Set Resources=Nothing
	Set resourceItem=Nothing
	Set currResourceList=Nothing
	Set ResourceFactory=Nothing
	Set resourceFolder=Nothing
	Set resourceFolders=Nothing
	Set resourceFolderFactory=Nothing
	'Set connection=Nothing
	Set runner=Nothing
	

End Function

'Function UpdateFileNames
'@param Directory     - Strng  - File Directory
'@param ProjectName   - String - Project name need to append to the File
'@param UpdateControl - boolean - true { append ProjectName} false {Remove ProjectName}


Function UpdateFileNames(Directory,ProjectName,UpdateControl)
	wscript.echo "[INFO] ---Geeting File in "&Directory

	Set temp = CreateObject("Scripting.Dictionary")
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set folder = fso.GetFolder(Directory)
	Set files = folder.Files
	If files.Count<>0 Then
		

		For Each file in files
		
		'IF update control is set to true. it will append project name infront of every file.
		'Else remove project name if allrady appended.
		
			IF UpdateControl=True Then
				wscript.echo "[INFO] ------File Found : "&file.Name
				IF NOT(InStr(file.Name, ProjectName)>0) Then
					file.Name=ProjectName&file.Name
					wscript.echo "[INFO] ------File Name Updated: "&file.Name
				Else
					wscript.echo "[INFO] ------File Name allready Updated: "&file.Name
				End IF
			Else
				IF InStr(file.Name, ProjectName)>0 Then
					file.Name=Replace(file.Name,ProjectName,"")
					wscript.echo "[INFO] ------File Name Reverted Back: "&file.Name
				End IF
			End IF
		Next
	
	Else
		wscript.echo "[INFO] ------No File Found to Update"
	End If
	
	wscript.echo "[INFO] ---Finished Updating File Names in "&Directory
	
	Set files=Nothing 
	Set folder=Nothing
	Set fso=Nothing
	Set temp=Nothing
	
End Function

Function getFiles(Directory,byref Dictionary)
	wscript.echo "[INFO] ---Geeting File in "&Directory

	Set temp = CreateObject("Scripting.Dictionary")
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set folder = fso.GetFolder(Directory)
	Set files = folder.Files
	If files.Count<>0 Then
		

	For Each file in files
		
		Path=Split(Lcase(folder.Path),"vtafruntime")
		temp.Add file.Name,folder.Path
		key=Path(Ubound(Path))
		wscript.echo "[INFO] ------File Found : "&file.Name
		
	Next
	'wscript.echo "Current Key"&key
	Dictionary.Add key,temp

	Else
		wscript.echo "[INFO] ------No File Found"
	End If
	wscript.echo "[INFO] ---Finished Geeting File in "&Directory
	
	Set files=Nothing 
	Set folder=Nothing
	Set fso=Nothing
	Set temp=Nothing
	
End Function

Function AssociateFunctionLibrary(Key,RemotePath,Dictionary,instance)
	RemotePath=Replace(RemotePath,"\","\_"&ExecutionName&"_")	
	 If Dictionary.Exists(key) Then
    
    	Set runner=Dictionary.Item(key)  
     	k = runner.Keys 
     	
   		For i = 0 To runner.Count - 1 
	       	resourceFile=k(i)
			
			RemoteLib=RemotePath&";;\"&resourceFile
	       	wscript.echo "[INFO] Associating Function Library ["&RemoteLib&"]"
	       	instance.Test.Settings.Resources.Libraries.Add RemoteLib
	    Next
	 End IF
	 Set runner=Nothing
End Function

Function InitializePropFiles(fileName,ProjectName)
	DIM fso
	SET fso = CreateObject ("Scripting.FileSystemObject")
	DIM strConfigLine,fConFile,EqualSignPosition, strLen, VariableName, VariableValue, propattrib, propvalue
	
	IF fso.FileExists(fileName) Then
		SET fConFile = fso.OpenTextFile(fileName)
	Else
		SET fConFile = fso.OpenTextFile(Replace(fileName,ProjectName,""))
	End IF
	
	Set Config = CreateObject("Scripting.Dictionary")
	WHILE NOT fConFile.AtEndOfStream
	  strConfigLine = fConFile.ReadLine
	  strConfigLine = TRIM(strConfigLine)
	'msgbox strConfigLine
	propattrib=split(strConfigLine, "=")(0)
	propvalue=split(strConfigLine, "=")(1)

	If  propattrib = "testPlan" Then
			strTrustPath = propvalue
	Else
		Config.ADD propattrib,propvalue
		
	End If
	WEND
	fConFile.Close
	
	Set fso=Nothing
	Set fConFile=Nothing

End Function




Public Function Load_Script(src)
    If Len(src) = 0 Then
        Err.Raise E_SCRIPT_TABLE, "Load_Script", "Script is empty"
    End If
    Load_Script = Split(Trim(src), vbCrLf)
End Function

Public Function Save_Script(script)

   If Not IsArray(script) Then
       Err.Raise E_SCRIPT_TABLE, "Save_Script", "Script should be string array"
   End If
    If UBound(script) - LBound(script) = 0 Then
        Err.Raise E_SCRIPT_TABLE, "Save_Script", "Script is empty"
    End If
    Save_Script = Join(script, vbCrLf)
End Function

Public Function Find_Line(script, criteria)
    Dim rExp, I
    '********************************************************************
    'Verify that the first argument contains a string array
   If Not IsArray(script) Then
       Err.Raise E_SCRIPT_TABLE, "Find_Line", "The script should be a string array"
   End If

    Set rExp = New RegExp
    ptrn = ""
    If IsArray(criteria) Then
        ptrn = Join(criteria, " * ")
    Else
        ptrn = criteria
    End If
    rExp.Pattern = ptrn 'Set pattern string
    rExp.IgnoreCase = True ' Set case insensitivity.
    rExp.Global = True ' Set global applicability.
    I = 0
    For Each scrItem In script
        If rExp.Execute(scrItem).Count > 0 Then
            Find_Line = I
            Exit Function
        End If
        I = I + 1
    Next
    Find_Line = -1
End Function

Public Function Move_Line(script, curPos, newPos)
    '********************************************************************
    'Verify that the first argument contains a string array
   If Not IsArray(script) Then
       Err.Raise E_SCRIPT_TABLE, "Move_Line", "Script should be string array"
   End If

   scrLen = UBound(script) - LBound(script)
    If curPos = newPos Or curPos < 0 Or newPos < 0 Or scrLen < curPos Or scrLen < newPos Then
        Move_Line = script
        Exit Function
    End If
    tmpLine = script(curPos)
    If newPos > curPos Then
        For curPos = curPos + 1 To scrLen
            script(curPos - 1) = script(curPos)
            If curPos = newPos Then
                script(curPos) = tmpLine
                Exit For
            End If
        Next
    Else
        For curPos = curPos - 1 To 0 Step -1
            script(curPos + 1) = script(curPos)
            If curPos = newPos Then
                script(curPos) = tmpLine
                Exit For
            End If
        Next
    End If
    Move_Line = script
End Function

Function Insert_Line(script, lineSrc, linePos)
    '********************************************************************
    'Verify that the first argument contains a string array
   If Not IsArray(script) Then
       Err.Raise E_SCRIPT_TABLE, "Insert_Line", "Script should be string array"
   End If
   scrLen = UBound(script) - LBound(script)

   If (scrLen = 0 And linePos <> 0) Or (linePos > scrLen + 1) Then
        Insert_Line = script
        Exit Function
   End If

   newScript = Split(String(scrLen + 1, " "), " ")
   shiftIndex = 0
   For I = 0 To scrLen + 1
        If linePos = I Then
            newScript(I) = lineSrc
            shiftIndex = 1
        Else
            newScript(I) = script(I + shiftIndex)
        End If
   Next
   Insert_Line = newScript
End Function

Function Delete_Line(script, linePos)
    '********************************************************************
    'Verify that the first argument contains a string array
   If Not IsArray(script) Then
       Err.Raise E_SCRIPT_TABLE, "Delete_Line", "Script should be string array"
   End If
    scrLen = UBound(script) - LBound(script)
   If (scrLen = 0) Or (linePos > scrLen) Then
        Insert_Line = script
        Exit Function
   End If

   If scrLen = 1 Then
       Delete_Line = Array()
       Exit Function
   End If

    newScript = Split(String(scrLen - 1, " "), " ")
    shiftIndex = 0
    For I = 0 To scrLen
        If linePos = I Then
            shiftIndex = 1
        Else
            newScript(I - shiftIndex) = script(I)
        End If
    Next
    Delete_Line = newScript
End Function

Public Function Move_CallAction(script, actName, newPos)
    curPos = Find_Line(script, Array("RunAction", """" & actName & """"))
    Move_CallAction = Move_Line(script, curPos, newPos)
End Function