Function StartExecution

If isLocalProject Then
	executeTestCases
Else
	Print "[INFO] Executing ALM Project"
End If
	   
End Function


Function CleanUpReportDirectory
	
	If Cbool(Config.Item("ALM"))=true Then
		FolderPath=Environment.Value("TestDir")&"\vtafRuntime\testReports\vtafReport\"
		CleanFolder(FolderPath)
		
	End If

End Function


Function UploadReportToALM
	
	If Cbool(Config.Item("ALM"))=true Then
		Print "[INFO] Starting to Upload Report to ALM"
		UploadToALM
		Print "[INFO] Finish Uploading Report to ALM"
		
	End If
End Function



Function isLocalProject

	isLocalProject=false
		Set fso = CreateObject("Scripting.FileSystemObject")
		If (fso.FileExists(Environment.Value("TestDir") & "\vtafRuntime\configs\config.txt")) Then
  			 Print "[INFO]" & path & " exists."
   		 	 isLocalProject=true
		Else
		   	Print "[INFO]" & path & " doesn't exist."
			isLocalProject=false
		End If
		
End Function


Function UploadToALM
	Dim objFSO, objFolder, qcPath

	qcPath = Environment.Value("TestDir") & "\vtafRuntime\testReports\"
	strFolder = qcPath & "vtafReport"
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	If objFSO.FolderExists(strFolder) Then
		ArchiveFolder strFolder & ".zip",strFolder 
		SetReportAttachement "vtafReport.zip", qcPath
	Else
	
	End IF
	
End Function

Function CleanFolder(FolderPath)
		print "[INFO] Cleaning Report Folder"
		Dim objFSO, objFolder
		strFolder = FolderPath
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		
		If objFSO.FolderExists(strFolder) Then
		    print "[INFO] Folder  exist."
			    If FolderEmpty (strFolder) Then
				    print "[INFO] --Files not exist."
			    Else
			    	print "[INFO] --File exist." 
		            delete strFolder
				End If
		Else
		    print "[EROR] --Folder does not exist."
		End If
	print "[INFO] Cleaning Report Folder Finished"
End Function

' Verify that a Folder Exists
Function FolderEmpty(strFolder)

Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FolderExists(strFolder) Then
    Set objFolder = objFSO.GetFolder(strFolder)
  
    If objFolder.Files.Count = 0 And objFolder.SubFolders.Count = 0 Then
        FolderEmpty=True
    Else
        FolderEmpty=False

    End If
End If

Set objFSO = Nothing
End Function

  ' delete all 
Function delete(strFolder)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(strFolder)
            
            ' delete all files in root folder
			for each file in folder.Files
			   On Error Resume Next
			   name = file.name
			   file.Delete True
				   If Err Then
				     print "[EROR] ---Error while deleting:" & Name & " - " & Err.Description
				   Else
				     print "[INFO] ---Deleted File: " & Name
				   End If
			   On Error GoTo 0
			Next
			
			' delete all subfolders and files
			For Each subFile In folder.SubFolders
			   On Error Resume Next
			   Name = subFile.name
			   subFile.Delete True
				   If Err Then
				     print "[EROR] Error deleting:" & Name & " - " & Err.Description
				   Else
				     print "[INFO] Deleted:" & Name
				   End If
			   On Error GoTo 0
			Next

	
End Function



Function ArchiveFolder (zipFile, sFolder)

    With CreateObject("Scripting.FileSystemObject")
        zipFile = .GetAbsolutePathName(zipFile)
        sFolder = .GetAbsolutePathName(sFolder)

        With .CreateTextFile(zipFile, True)
            .Write Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, chr(0))
        End With
    End With

    With CreateObject("Shell.Application")
        .NameSpace(zipFile).CopyHere .NameSpace(sFolder).Items

        Do Until .NameSpace(zipFile).Items.Count = _
                 .NameSpace(sFolder).Items.Count
            wait 10
        Loop
    End With

End Function


Function SetReportAttachement (attachmentFile,localPath)
	Dim nowTest
	Set nowTest = QCUtil.CurrentTestSet
	Set attachmentPath = nowTest.Attachments
	Set nowAttachment = attachmentPath.AddItem(Null)
	'Replace with the path to your file:
	nowAttachment.FileName = localPath&"\"&attachmentFile
	nowAttachment.Type = 1
	nowAttachment.Post()
	
End Function
