'VTAF QTP WEB run time
'-------------------------------------------
'Change log
'-------------------------------------------
' Date          Version   
' 2015.04.28    1.1
'-------------------------------------------



 Dim tcTotal, tcPassed, tcFailed,VtafreportSmallImagePath
''<Procedure>
''<name> InsertIntoHTMLReport</name>
''<description> It is the main parent function that checks the availability of the  text log file for the current date and if it is present then it calls another function to update this text file, else it call function to create text file and then update it with result logs. </description>
''<param name="strStepName">[in] The calling function name or  the Step name</param>
''<param name="strStepShortDesc">[in] Short Description of Step or calling function</param>
''<param name="strExpectedResult">[in] Expected Result description</param>
''<param name="blnStatus">[in]  "True" for passed  validation, "False" for  failed validation</param>
''<param name="strLogfilePath">[in] Result folder path where the html log file needs to be created</param>
''<param name="blnImage">[in] True or False , True if SnapShot is required , False if not required </param>
''<returns> NIL</returns>
''<example>
''
''    Call InsertIntoHTMLReport( "Pearson Testing",  "Testing Pearson Scenarios", "Automation Success", True, "D:\QTP_Pearson\test_html",True)
''
''</example>
''<changelog>
''   Date                            Author                    Changes/Notes
''-----------                    ------------------                -----------------------
''     28-Dec-2010        Shalabh Dixit            Initial version.
''</changelog>
''</Procedure>
'''''*************************************************************************************************************************************************************************************************************
'''''*************************************************************************************************************************************************************************************************************


Public Sub InsertIntoHTMLReport(logType,strStepName, strStepShortDesc, strExpectedResult, blnStatus, strLogfilePath,blnImage)
    On Error Resume Next
    Dim dtMyDate
    Dim strHTMLFilePath
	Dim strLoggerPath
        dtMyDate = Day(Date) &"_"& MonthName(Month(Date),3)  &"_"& Year(Date)
           ' strHTMLFilePath =strLogfilePath &"\HTML_Result_Log_"& Environment.Value("TestName") &"_TestSuite_"& dtMyDate & ".txt"  
		   strHTMLFilePath =strLogfilePath &"\HTML_Result_Log_"& "VTAF3.0" &"_TestSuite_"& dtMyDate & ".txt"  
		   strLoggerPath = strLogfilePath &"\XML_Result_Log_"& "VTAF3.0" &"_TestSuite_DetailsTestLog"& dtMyDate & ".txt"
           
    If strLogfilePath <> "" Then  
        Dim objFSO
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        If (objFSO.FileExists(strHTMLFilePath)) Then
            Call UpdateHTMLReport(logType,strStepName, strStepShortDesc, strExpectedResult, blnStatus, strHTMLFilePath,strLogfilePath,blnImage)
        Else           
            Call CreateHTMLReport(strHTMLFilePath)
            Call UpdateHTMLReport (logType,strStepName, strStepShortDesc, strExpectedResult, blnStatus, strHTMLFilePath,strLogfilePath,blnImage)
        End If   

		If (objFSO.FileExists(strLoggerPath)) Then
            Call UpdateTestLog(logType,strStepName, strStepComments, blnStatus, strLoggerPath,strLogfilePath)
        Else           
            Call CreateTestLog(strLoggerPath)
            Call UpdateTestLog(logType,strStepName, strStepComments, blnStatus, strLoggerPath,strLogfilePath)
        End If   

    Else   
        Reporter.ReportEvent micFail, "ERROR. Log file path is not specified", ""   
    End If         
    Set objFSO = Nothing
End Sub

''''''''''''''''''UpdateHTMLReport''''''''''''''''
''''''''''''''*********************************
''<Procedure>
''<name> UpdateHTMLReport</name>
''<description> This function  updates the text log file as well as the QTP embeded Result log.</description>
''<param name="strStepName">[in] The calling function name or  the Step name</param>
''<param name="strStepDesc">[in] Short Description of Step or calling function</param>
''<param name="strExpectedResult">[in] Expected Result description</param>
''<param name="blnStatus">[in]  "True" for passed  validation, "False" for  failed validation</param>
''<param name="strHTMLFilePath">[in] The text file path which needs to be updated and transformed into html file.</param>
''<param name="strLogfilePath">[in] Result folder path where the html log file needs to be created</param>
'<param name="blnImage">[in] True or False , True if SnapShot is required , False if not required </param>
''<returns> NIL</returns>
''<example>
''    strHTMLFilePath =strLogfilePath &"\HTML_Result_Log_"& Environment.Value("TestName") &"_TestSuite_"& dtMyDate & ".txt"  
''    Call UpdateHTMLReport( "Pearson Testing",  "Testing Pearson Scenarios", "Automation Success", True, strHTMLFilePath,"D:\QTP_Pearson\test_html",True)
''
''</example>
''<changelog>
''   Date                            Author                    Changes/Notes
''-----------                    ------------------                -----------------------
''     28-Dec-2010        VTAF Team          Initial version.
''</changelog>
''</Procedure>
'''''*************************************************************************************************************************************************************************************************************
'''''*************************************************************************************************************************************************************************************************************
Public Sub UpdateHTMLReport(logType,strStepName, strStepDesc, strExpectedResult, blnStatus, strHTMLFilePath,strLogfilePath,blnImage)
On Error Resume Next
            Dim dtMyDate
            Dim qtApp
            Dim strMyStatus
            Dim strVarStatus
            Dim strColor
            Dim dtExecutionTime
            Dim a, intHour  , intMinute , intSec   
                   
                Const ForWriting = 2
                Const ForAppending = 8

                dtMyDate = Day(Date) &"_"& MonthName(Month(Date),3)  &"_"& Year(Date)

                If UCase(Trim(blnStatus)) = Ucase("True")  Then
                        strMyStatus = "Pass"
                        strColor ="GREEN"
                Elseif UCase(Trim(blnStatus)) = Ucase("False") Then
                        strMyStatus = "Fail"
                        strColor ="RED"
                End If

                Set qtApp = CreateObject("QuickTest.Application")
                strTimeStamp = Day(Date) & Month(Date) & Year(Date) &"_"& Hour(Time) & Minute(Time) & Second(Time)

                If UCase(Trim(blnImage)) =UCase("True")Then               
                ''    If blnStatus <> "True"   Then
                        Dim strStatusFileName, strStatusFilePath,strStatusStepFilePath                                               
                            strStatusFileName = "Execution_status_image_"& strTimeStamp & ".bmp"
                            strStatusFilePath = newvtafreportPath & "\images\" & strStatusFileName
							vtafreportImagePath ="\images\"  &  strStatusFileName
                            qtApp.Visible = False
                            Wait(1)
                            Desktop.CaptureBitmap strStatusFilePath, True
                            qtApp.Visible = True
                            strStatusStepFilePath = strStatusFilePath
                            
                              'Resize Error Image
		               SmallImagePath=newvtafreportPath & "\images\SmallImage_" & strStatusFileName
		               VtafreportSmallImagePath="images/SmallImage_" & strStatusFileName
		               Set oImage=DotNetFactory.CreateInstance("System.Drawing.Image","System.Drawing")
		
						vstPath=strStatusFilePath
		
						Set oBitmap=DotNetFactory.CreateInstance("System.Drawing.Bitmap","System.Drawing",oImage.FromFile(vstPath),150,100)
		
						oBitmap.Save(SmallImagePath)
		
						Set oBitmap=Nothing
						Set oImage=Nothing
		               

			
                ElseIf    UCase(Trim(blnImage)) =UCase("False") Then
                       strStatusStepFilePath  = "NA"
                Else
                        Reporter.ReportEvent micWarning, "InsertIntoHTMLReport","blnImage parameter value: "& blnImage &"  is not passed properly. Please pass boolean value."
                        Exit Sub
                End If
			  	If  logType= "endOfTestStep" Then
						dtonlytime = Hour(Time) & ":"& Minute(Time) & ":" & Second(Time)
						afterStep strStepName, strStepDesc,  blnStatus, dtonlytime, vtafreportImagePath
						
				End If
	               
                    a= "0"             
                    intHour = Hour(Time)
                    intMinute = Minute(Time)           
                    intSec = Second(Time)
                   
                    If Len(intHour) < 2  Then
                        intHour = a & intHour
                    End If
           
                    If Len(intMinute) < 2  Then
                        intMinute = a & intMinute
                    End If
           
                    If Len(intSec) < 2  Then
                        intSec = a & intSec
                    End If
                           
                dtExecutionTime =Day(Date) &"-"& MonthName(Month(Date),3)  &"-"& Year(Date)&", "& intHour  &":"& intMinute &":"& intSec           

                Set fso = CreateObject ("Scripting.FileSystemObject")   

                Set objHTMLTextFile = fso.OpenTextFile(strHTMLFilePath, 8, true)           


				Select Case logType

				Case "startOfTestCase"
					objHTMLTextFile.WriteLine("<tr>")
					objHTMLTextFile.WriteLine("<td><font color=BLACK face=ARIAL>"& strStepName &"</font></td>")
					objHTMLTextFile.WriteLine("<td><font color=BLACK face=ARIAL></font></td>")
					objHTMLTextFile.WriteLine("<td><font color=BLACK face=ARIAL>"& strStepDesc &"</font></td>")
					objHTMLTextFile.WriteLine("<td><font color=BLACK face=ARIAL>"& strExpectedResult &"</font></td>")                                
					objHTMLTextFile.WriteLine("<td ><font color=" & strColor &" face=ARIAL></font></td>")
					objHTMLTextFile.WriteLine("<td ><font color=BLACK face=ARIAL>"& dtExecutionTime &"</font></td>")
					If strStatusStepFilePath ="NA" Then   
						objHTMLTextFile.WriteLine("<td ><font color=BLACK face=ARIAL>"& strStatusStepFilePath &"</font></td>")
					Else
						objHTMLTextFile.WriteLine("<td><font color=" & strColor &" face=ARIAL><a href='file:///"& strStatusStepFilePath & "'>"& strStatusStepFilePath & "</a></font></td>")
					End If
    				objHTMLTextFile.WriteLine("</tr>")
				Case "endOfTestStep"
				
					If isEmpty(strStatusFilePath) Then
						UpdateUFTReport true,blnStatus,strStepName,strStepDesc,""
					Else
						UpdateUFTReport true,blnStatus,strStepName,strStepDesc,strStatusFilePath
					End IF

				
					objHTMLTextFile.WriteLine("<tr>")
					'objHTMLTextFile.WriteLine("<td><font color=BLACK face=ARIAL>"& Environment.Value("ActionName") &"</font></td>")
					objHTMLTextFile.WriteLine("<td><font color=BLACK face=ARIAL></font></td>")
					objHTMLTextFile.WriteLine("<td><font color=BLACK face=ARIAL>"& strStepName &"</font></td>")
					objHTMLTextFile.WriteLine("<td><font color=BLACK face=ARIAL>"& strStepDesc &"</font></td>")
					objHTMLTextFile.WriteLine("<td><font color=BLACK face=ARIAL>"& strExpectedResult &"</font></td>")                                
					objHTMLTextFile.WriteLine("<td ><font color=" & strColor &" face=ARIAL> "& strMyStatus &" </font></td>")
					objHTMLTextFile.WriteLine("<td ><font color=BLACK face=ARIAL>"& dtExecutionTime &"</font></td>")
					If strStatusStepFilePath ="NA" Then   
						objHTMLTextFile.WriteLine("<td ><font color=BLACK face=ARIAL>"& strStatusStepFilePath &"</font></td>")
					Else
						objHTMLTextFile.WriteLine("<td><font color=" & strColor &" face=ARIAL><a href='file:///"& strStatusStepFilePath & "'>"& strStatusStepFilePath & "</a></font></td>")
					End If
    				objHTMLTextFile.WriteLine("</tr>")
				Case "endOfTestCase"
					objHTMLTextFile.WriteLine("<tr>")
					objHTMLTextFile.WriteLine("<td><font color=BLACK face=ARIAL>Overall Status</font></td>")
					objHTMLTextFile.WriteLine("<td><font color=BLACK face=ARIAL></font></td>")
					objHTMLTextFile.WriteLine("<td><font color=BLACK face=ARIAL></font></td>")
					objHTMLTextFile.WriteLine("<td><font color=BLACK face=ARIAL>"& strExpectedResult &"</font></td>")                                
					objHTMLTextFile.WriteLine("<td ><font color=" & strColor &" face=ARIAL> " & strMyStatus & " </font></td>")
					objHTMLTextFile.WriteLine("<td ><font color=BLACK face=ARIAL>"& dtExecutionTime &"</font></td>")
					If strStatusStepFilePath ="NA" Then   
						objHTMLTextFile.WriteLine("<td ><font color=BLACK face=ARIAL>"& strStatusStepFilePath &"</font></td>")
					Else
						objHTMLTextFile.WriteLine("<td><font color=" & strColor &" face=ARIAL><a href='file:///"& strStatusStepFilePath & "'>"& strStatusStepFilePath & "</a></font></td>")
					End If
    				objHTMLTextFile.WriteLine("</tr>")
				End Select
               
                If strMyStatus = "Pass" Then
                    strVarStatus  = micPass
                    Reporter.ReportEvent strVarStatus , strStepName , strStepDesc & ", Expected Result :  <" & strExpectedResult &">.", strStatusStepFilePath    
                Elseif strMyStatus ="Fail" Then
                    strVarStatus  = micFail
                    Reporter.ReportEvent strVarStatus , strStepName , strStepDesc & " , Expected Result :  <" & strExpectedResult & ">.",strStatusStepFilePath    
                End If    

                objHTMLTextFile.Close
                           
                Set objHTMLTextFile = Nothing
                Set qtApp = Nothing
                Set fso = Nothing

End Sub
''<Procedure>
''<name> CreateHTMLReport</name>
''<description> This function  creates the text log file based on current date as an input for html log file.</description>
''<param name="strHTMLFilePath">[in] The text file path which needs to be created and transformed into html file.</param>

''<returns> NIL</returns>
''<example>
''    strHTMLFilePath =strLogfilePath &"\HTML_Result_Log_"& Environment.Value("TestName") &"_TestSuite_"& dtMyDate & ".txt"  
''    Call CreateHTMLReport(strHTMLFilePath)

''
''</example>
''<changelog>
''   Date                            Author                    Changes/Notes
''-----------                    ------------------                -----------------------
''     28-Dec-2010        Shalabh Dixit           Initial version.
''</changelog>
''</Procedure>
'''''*************************************************************************************************************************************************************************************************************
'''''*************************************************************************************************************************************************************************************************************

Public Sub CreateHTMLReport(strTextLogfilePath)
    On Error Resume Next
        Dim dtStartDateTime
        Dim dtEndDateTime
        Dim intHour, intMinute ,intSec , a
        Dim fso
        Const ForWriting = 2
        a= "0" 

        intHour = Hour(Time)
        intMinute = Minute(Time)           
        intSec = Second(Time)
       
        If Len(intHour) < 2  Then
            intHour = a & intHour
        End If

        If Len(intMinute) < 2  Then
            intMinute = a & intMinute
        End If

        If Len(intSec) < 2  Then
            intSec = a & intSec
        End If

        dtStartDateTime = Day(Date) &"-"& MonthName(Month(Date),3)  &"-"& Year(Date)&", "& intHour  &":"& intMinute &":"& intSec           
        Set fso = CreateObject ("Scripting.FileSystemObject")
       
        If Not (fso.FileExists(strTextLogfilePath)) Then
                    fso.CreateTextFile strTextLogfilePath, True                   
                    Set objHTMLTextFile = fso.OpenTextFile(strTextLogfilePath, 2, True)
                   
                    objHTMLTextFile.WriteLine "<html>"

                    objHTMLTextFile.WriteLine "<head>"
                    objHTMLTextFile.WriteLine "<style>"
                    objHTMLTextFile.WriteLine "td{"
                    objHTMLTextFile.WriteLine "border-color: black;"
                    objHTMLTextFile.WriteLine "font-size: 14;"
                    objHTMLTextFile.WriteLine "}"
                    objHTMLTextFile.WriteLine "thead{"
                    objHTMLTextFile.WriteLine "border-color: black;"
                    objHTMLTextFile.WriteLine "font-size: 14;"
                    objHTMLTextFile.WriteLine "font-weight: bold;"
                    objHTMLTextFile.WriteLine "text-align: center;"
                    objHTMLTextFile.WriteLine "}"
                    objHTMLTextFile.WriteLine ".resultSummary{"
                    objHTMLTextFile.WriteLine "border-color: black;"
                    objHTMLTextFile.WriteLine "font-size: 20;"
                    objHTMLTextFile.WriteLine "font-weight: bold;"
                    objHTMLTextFile.WriteLine "text-align: center;"
                    objHTMLTextFile.WriteLine "}"
                    objHTMLTextFile.WriteLine ".summaryTable{"
                    objHTMLTextFile.WriteLine "border-color: black;"
                    objHTMLTextFile.WriteLine "}"
                    objHTMLTextFile.WriteLine ".centerTd{"
                    objHTMLTextFile.WriteLine "text-align: center;"
                    objHTMLTextFile.WriteLine "}"
                    objHTMLTextFile.WriteLine "</style>"
                    objHTMLTextFile.WriteLine "</head>"

                    objHTMLTextFile.WriteLine "<body bgcolor=#99ccff >"  ''''body Color "Sky Blue"
                    objHTMLTextFile.WriteLine "<h1 style="&chr(34)&"text-align:center"&chr(34)&"><b> Automation Tests Results</b></h1>"
                    objHTMLTextFile.WriteLine "<br />"

                    objHTMLTextFile.WriteLine "<table align=center" &" " &"border=1"&" " &" class=summaryTable"&">"
                    objHTMLTextFile.WriteLine "<thead><tr><td colspan=2" &" " &"align=center " &" " &" class=resultSummary"&"> Result Summary: </td></tr></thead>"

                    objHTMLTextFile.WriteLine "<tr><td>Execution Start Date and Time : </td><td>"& dtStartDateTime &"</td></tr>"
                    objHTMLTextFile.WriteLine "<tr><td>Execution End Date and Time : </td><td> EndDateTime </td></tr>"
                    objHTMLTextFile.WriteLine "<tr><td>Executed on Machine : </td><td>"& Environment.Value("LocalHostName") &"</td></tr>"
                    objHTMLTextFile.WriteLine "<tr><td>Test Suite Name : </td><td>"& Environment.Value("TestName") &"</td></tr>" 
                    objHTMLTextFile.WriteLine "<tr><td>Executed by : </td><td>"& Environment.Value("UserName") &"</td></tr>"
                    objHTMLTextFile.WriteLine "<tr><td>Number of Steps Executed : </td><td> subTotal </td></tr>"
                    objHTMLTextFile.WriteLine "<tr><td>Number of Steps Passed : </td><td> subPassed </td></tr>"
                    objHTMLTextFile.WriteLine "<tr><td>Number of Steps Failed : </td><td> subFailed </td></tr>"

					 objHTMLTextFile.WriteLine "<tr><td>Number of TestCases Executed : </td><td> tcTotal </td></tr>"
                    objHTMLTextFile.WriteLine "<tr><td>Number of TestCases Passed : </td><td> tcPassed </td></tr>"
                    objHTMLTextFile.WriteLine "<tr><td>Number of TestCases Failed : </td><td> tcFailed </td></tr>"

                    objHTMLTextFile.WriteLine "</table>"
       
                    objHTMLTextFile.WriteLine "<hr />"
                    objHTMLTextFile.WriteLine "<h2 align=center>Result Description: </h2>"
                    objHTMLTextFile.WriteLine "<hr />"
                                   
                    objHTMLTextFile.WriteLine "<Table Border="&chr(34)&"1"&chr(34)&"cellpadding="&chr(34)&"10"&chr(34)&"bgcolor="&chr(34)&"#FFFFFF"&chr(34)&">"
                    objHTMLTextFile.WriteLine "<tr>"
                    objHTMLTextFile.WriteLine "</tr>"   

                    objHTMLTextFile.WriteLine"<thead>"
                    objHTMLTextFile.WriteLine("<tr bgcolor=#99ffff height=2>")
                    objHTMLTextFile.WriteLine("<td width=200><p align=center >Test Case Name</p></td>")                   
                    objHTMLTextFile.WriteLine("<td width=200><p align=center >Step Name</p></td>")
                    objHTMLTextFile.WriteLine("<td width=200><p align=center >Description</p></td>")
                    objHTMLTextFile.WriteLine("<td width=200><p align=center >Expected Result</p></td>")
                    objHTMLTextFile.WriteLine("<td width=136><p align=center >Status</p></td>")
                    objHTMLTextFile.WriteLine("<td width=200><p align=center >Execution Time</p></td>")
                    objHTMLTextFile.WriteLine(" <td width=200><p align=center >Status Image Location</p></td>")                                   
                    objHTMLTextFile.WriteLine("</tr>") 
                    objHTMLTextFile.WriteLine"</thead>"
                                       
                    objHTMLTextFile.Close                   
        End If   
            Set objHTMLTextFile = nothing               
            Set fso = nothing
End Sub


'''''*************************************************************************************************************************************************************************************************************
''''**************************************************************************************************************************************************************************************************************
''<Procedure>
''<name> FinishCreatingHTMLReport</name>
''<description> This function  should be called at the end of teh suite, it  converts teh text log file into html log file.</description>
''<param name="strLogfilePath">[in] Result folder path where the html log file needs to be created.</param>
''<returns> NIL</returns>
''<example>
''   
''    Call FinishCreatingHTMLReport("D:\QTP_Pearson\test_html")

''
''</example>
''<changelog>
''   Date                            Author                    Changes/Notes
''-----------                    ------------------                -----------------------
''     28-Dec-2010        Shalabh Dixit           Initial version.
''</changelog>
''</Procedure>
'''''*************************************************************************************************************************************************************************************************************
'''''*************************************************************************************************************************************************************************************************************
Public Sub FinishCreatingHTMLReport(strLogfilePath)
''Public Sub FinishCreatingHTMLReport(strTextLogfilePath,strLogfilePath)
    On Error Resume Next
            Dim fso
			Dim isTC
            Dim objHTMLTextFile
            Dim strOldContents
            Dim a, intHour, intMinute, intSec
            Dim  strMyDate, strTextLogfilePath, dtMyDate, strLineData
            Dim arrLineData, strTotal , strPassed, strFailed
            Dim  strNewContents1,strNewContents2, strNewContents3, strNewContents4, strNewContents5,strNewContents6,strNewContents7
            Const ForReading = 1

            Const ForWriting = 2
            Const ForAppending = 8

            strMyDate = Day(Date) &"_"& MonthName(Month(Date),3)  &"_"& Year(Date)
            strTextLogfilePath =strLogfilePath &"\HTML_Result_Log_"& "VTAF3.0"&"_TestSuite_"& strMyDate & ".txt"  

                dtMyDate = Day(Date) &"_"& MonthName(Month(Date),3)  &"_"& Year(Date) &"-"& Hour(Time)  &"_"& Minute(Time) &"_"& Second(Time)           
                Set fso = CreateObject ("Scripting.FileSystemObject")
                If fso.FileExists(strTextLogfilePath) Then
                    If fso.FolderExists(strLogfilePath) Then
                        a = "0"

                        intHour = Hour(Time)
                        intMinute = Minute(Time)           
                        intSec = Second(Time)
                       
                        If Len(intHour) < 2  Then
                        intHour = a & intHour
                        End If
                       
                        If Len(intMinute) < 2  Then
                        intMinute = a & intMinute
                        End If
                       
                        If Len(intSec) < 2  Then
                        intSec = a & intSec
                        End If 
                       
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            strPassed= 0
                            strFailed= 0
							tcPassed = 0
							tcFailed = 0       
                            Set objHTMLTextFile = fso.OpenTextFile(strTextLogfilePath,1)               
                            Do Until objHTMLTextFile.AtEndOfStream = "True"                               
                                    strLineData = Trim(objHTMLTextFile.ReadLine)
                                            If strLineData <> "" Then                               
                                                arrLineData = Split(strLineData," ")
                                                For i = 0 To UBound(arrLineData)
													If  InStr(1, Trim(strLineData), Trim("Overall Status")) > 0 Then
														isTC = true
													End if

                                                    If Trim(arrLineData(i)) =  "Pass" Then
															If (isTC=true) Then
																tcPassed = tcPassed+1
																isTC =  false
															Else
																strPassed = strPassed+1
															End If
													ElseIf Trim(arrLineData(i)) =  "Fail" Then
															If  (isTC=true) Then
																tcFailed = tcFailed+1   
																isTC = false
															Else
																 strFailed = strFailed+1   
															End IF
														End If
                                                    Next
										   End If
                            Loop
                    strTotal = strPassed+ strFailed
					tcTotal = tcPassed+ tcFailed
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    objHTMLTextFile.Close

                    dtEndDateTime = Day(Date) &"-"& MonthName(Month(Date),3)  &"-"& Year(Date)&", "& intHour  &":"& intMinute &":"& intSec           
                   
                    Set objHTMLTextFile = fso.OpenTextFile(strTextLogfilePath,1)
                            strOldContents = objHTMLTextFile.ReadAll
                            strNewContents1 = Replace(strOldContents, "EndDateTime",dtEndDateTime)
                            strNewContents2 =Replace(strNewContents1, "subTotal",strTotal)
                            strNewContents3 =Replace(strNewContents2, "subPassed",strPassed)
                            strNewContents4 =Replace(strNewContents3, "subFailed",strFailed)
							strNewContents5 =Replace(strNewContents4, "tcTotal",tcTotal)
                            strNewContents6 =Replace(strNewContents5, "tcPassed",tcPassed)
                            strNewContents7 =Replace(strNewContents6, "tcFailed",tcFailed)
                            objHTMLTextFile.Close
                           
                    Set objHTMLTextFile = fso.OpenTextFile(strTextLogfilePath,2, True)
                        objHTMLTextFile.Write strNewContents7
                        objHTMLTextFile.Close
                       
                    Set objHTMLTextFile = fso.OpenTextFile(strTextLogfilePath,8, True)
                    objHTMLTextFile.WriteLine "<tr>"
                    objHTMLTextFile.WriteLine "</tr>"       
                    objHTMLTextFile.WriteLine "</Table>"                   
                    objHTMLTextFile.WriteLine "</body>"
                    objHTMLTextFile.WriteLine "</html>"
                    objHTMLTextFile.Close

                        fso.MoveFile strTextLogfilePath,strLogfilePath &"\HTML_Result_Log_"& Environment.Value("TestName") &"_TestSuite_"& dtMyDate & ".html"   
                        'SystemUtil.Run(strLogfilePath &"\HTML_Result_Log_"& Environment.Value("TestName") &"_TestSuite_"& dtMyDate & ".html")
                    Else
                        Reporter.ReportEvent micWarning, "FinishCreatingHTMLReport",strLogfilePath &" : Result folder path is invalid."
                    End If
                Else
                     Reporter.ReportEvent micWarning, "FinishCreatingHTMLReport",strTextLogfilePath &": file path is invalid."
                End If

            Set fso = Nothing
End Sub
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub CreateTestLog(strTextLogfilePath)
    On Error Resume Next
        Dim dtStartDateTime
        Dim dtEndDateTime
        Dim intHour, intMinute ,intSec , a
        Dim fso
        Const ForWriting = 2
        a= "0" 

        intHour = Hour(Time)
        intMinute = Minute(Time)           
        intSec = Second(Time)
       
        If Len(intHour) < 2  Then
            intHour = a & intHour
        End If

        If Len(intMinute) < 2  Then
            intMinute = a & intMinute
        End If

        If Len(intSec) < 2  Then
            intSec = a & intSec
        End If

        dtStartDateTime = Day(Date) &"-"& MonthName(Month(Date),3)  &"-"& Year(Date)&", "& intHour  &":"& intMinute &":"& intSec           
        Set fso = CreateObject ("Scripting.FileSystemObject")
       
        If Not (fso.FileExists(strTextLogfilePath)) Then
                    fso.CreateTextFile strTextLogfilePath, True                   
                    Set objTextFile = fso.OpenTextFile(strTextLogfilePath, 2, True)
                    objTextFile.WriteLine "Starting Test Execution" & Date & " : " & Time
                    objTextFile.Close                   
        End If   
            Set objTextFile = nothing               
            Set fso = nothing
End Sub

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Public Sub UpdateTestLog(logType,strStepName, strStepComments,  blnStatus, strFilePath,strLogfilePath)
On Error Resume Next
            Dim dtMyDate
            Dim qtApp
            Dim strMyStatus
            Dim strVarStatus
            Dim strColor
            Dim dtExecutionTime
            Dim a, intHour  , intMinute , intSec   
                   
                Const ForWriting = 2
                Const ForAppending = 8

                dtMyDate = Day(Date) &"_"& MonthName(Month(Date),3)  &"_"& Year(Date)

                If UCase(Trim(blnStatus)) = Ucase("True")  Then
                        strMyStatus = "Pass"
                        'strColor ="GREEN"
                Elseif UCase(Trim(blnStatus)) = Ucase("False") Then
                        strMyStatus = "Fail"
                        'strColor ="RED"
                End If

                Set qtApp = CreateObject("QuickTest.Application")
                strTimeStamp = Day(Date) & Month(Date) & Year(Date) &"_"& Hour(Time) & Minute(Time) & Second(Time)

                    a= "0"             
                    intHour = Hour(Time)
                    intMinute = Minute(Time)           
                    intSec = Second(Time)
                   
                    If Len(intHour) < 2  Then
                        intHour = a & intHour
                    End If
           
                    If Len(intMinute) < 2  Then
                        intMinute = a & intMinute
                    End If
           
                    If Len(intSec) < 2  Then
                        intSec = a & intSec
                    End If
                           
                dtExecutionTime =Day(Date) &"-"& MonthName(Month(Date),3)  &"-"& Year(Date)&", "& intHour  &":"& intMinute &":"& intSec           
				dtonlytime = intHour  &":"& intMinute &":"& intSec       
                Set fso = CreateObject ("Scripting.FileSystemObject")   

                Set objTextFile = fso.OpenTextFile(strFilePath, 8, true)           


				Select Case logType

				Case "startOfTestCase"
					objTextFile.WriteLine "Starting Test Case : " & strStepName & " at " & dtonlytime 

				Case "endOfTestStep"
					objTextFile.WriteLine "          Execution of Test Step : " & strStepName & " at " & dtExecutionTime & " With Results "  & strMyStatus
					
				Case "endOfTestCase"
					objTextFile.WriteLine "End of Test Cas Execution "  & " at " & dtExecutionTime & " With Results "  & strMyStatus

				End Select
                           
                objTextFile.Close
                      
                Set objTextFile = Nothing
                Set qtApp = Nothing
                Set fso = Nothing

End Sub

'ALMSupport Functions


Function UpdateUFTReport(isStatment,Status,ReportStepName,ReportDetails,ImageFilePath)

Reporter.Filter=rfEnableAll

'if it is a statment

IF cbool(isStatment)=true Then

	'report pass
	IF Cbool(Status)=true Then

		IF NOT(ImageFilePath="") Then

			Reporter.ReportEvent micPass,ReportStepName,ReportDetails,ImageFilePath

		Else
			Reporter.ReportEvent micPass,ReportStepName,ReportDetails
		End IF

	ElseIF cbool(Status)=false Then

		IF NOT(ImageFilePath="") Then

			Reporter.ReportEvent micFail,ReportStepName,ReportDetails,ImageFilePath

		Else
			Reporter.ReportEvent micFail,ReportStepName,ReportDetails
		End IF
	End If
		
'if it is a event
Else

	IF NOT(ImageFilePath="") Then

	Reporter.ReportEvent micDone,ReportStepName,ReportDetails,ImageFilePath

	Else
	Reporter.ReportEvent micDone,ReportStepName,ReportDetails
	End IF

End IF

Reporter.Filter =rfDisableAll

End Function

