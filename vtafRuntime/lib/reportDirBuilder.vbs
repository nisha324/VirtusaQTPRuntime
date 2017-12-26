'VTAF QTP WEB run time
'-------------------------------------------
'Change log
'-------------------------------------------
' Date          Version   
' 2014.10.20    1.0
'-------------------------------------------

Dim fso, loggertxt, varWitertoReport:
 
Function createExectionReport() 

    dtMyDate =Year(Date) &"_"& Month(Date) &"_"&Day(Date)&"_"& Hour(Time)&"_"& Minute(Time)    &"_"& Second(Time)
	executionName = "ExecutionReport" & dtMyDate

	'InitializePropFiles(strPath& "\TestConfig.properties")

	
	set fso = CreateObject("Scripting.FileSystemObject")
	'Create a new folder
	newvtafreportPath = ProjectPath & "\vtafRuntime\testReports\vtafReport\" & executionName
	'msgbox newvtafreportPath
	
	fso.CreateFolder newvtafreportPath
	fso.CreateFolder newvtafreportPath & "\LogFile" 
	fso.CreateFolder newvtafreportPath & "\images" 
	'Set loggertxt = fso.CreateTextFile newvtafreportPath & "\LogFile\vtafsupportlog.txt", True
	sOriginFolder = ProjectPath & "\vtafRuntime\configs\ReportTemplete\"

	 sDestinationFolder = newvtafreportPath
	 'msgbox sOriginFolder
	 For Each sFile In fso.GetFolder(sOriginFolder).Files
	  If Not fso.FileExists(sDestinationFolder & "\" & fso.GetFileName(sFile)) Then
	   fso.GetFile(sFile).Copy sDestinationFolder & "\" & fso.GetFileName(sFile),True
	  End If
	 Next
	  Set varWitertoReport = fso.OpenTextFile(newvtafreportPath & "\LogFile\vtafsupportlog.txt", 8, true)
   
End Function

Function setLogger(str)
	 varWitertoReport.WriteLine(str)
End Function

