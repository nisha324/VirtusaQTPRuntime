'VTAF UFT-ALM Integrated WEB Runtime
'02-DEC-2015
'12-09-2016 add ExtendSelect for select command 

RegisterUserFunc "WebList", "RegexSelectDOM", "RegexSelectDOM"
Dim retryTime,arrIdentifire(10),ExcelObject,ExcelSheet,ExcelStore
retryTime = 100

Sub RegexSelectDOM(Object, sPattern)
    Dim oRegExp, oOptions, ix

    'Create RegExp Object
    Set oRegExp = New RegExp
    oRegExp.IgnoreCase = False
    oRegExp.Pattern = sPattern

    'DOM options
    Set oOptions = Object.Object.Options
	
    For ix = 0 to oOptions.Length - 1
        'If RegExp pattern matches list item, we're done!
        If oRegExp.Test(oOptions(ix).Text) Then
            oOptions(ix).selected = true
             Set oRegExp = Nothing
            Exit Sub
        End If
    Next
End Sub

'Command - Check Element Present
'@param String oObject
'@param String identifire
'@param Boolean AssertType
'@param String customErrorMessage

Public Function cCheckElementPresent (oObject, identifire , AssertType, customErrorMessage)
	On Error Resume Next
	SendObject oObject, identifire
	startCommand "CheckElementPresent  Command", oObject
	'Browser("title:=.*").Page("title:=.*").Sync
    If CommandObj.Exist(retryTime)Then
    
    If True=Cbool(Config.Item("HighLight")) Then
    	CommandObj.Highlight
    End If
	
			Call InsertIntoHTMLReport( "endOfTestStep","Check Element Present","Element ["&oObject&"] is Present" , "", True, Config.Item("ReportPath"),False)
			endCommand "Pass"
	Else 
			If  AssertType = True Then
					Call InsertIntoHTMLReport( "endOfTestStep","Check Element Present",cCustomErrorGeneration(customErrorMessage) &"Element ["&oObject&"] is not found" , "", False, Config.Item("ReportPath"),True)
					endCommand "Fail"
					ErrorNO =  8    'raise a user-defined error
			

			else 
					Call InsertIntoHTMLReport( "endOfTestStep","Check Element Present",cCustomErrorGeneration(customErrorMessage) &"Element ["&oObject&"] is not found"  , "", False, Config.Item("ReportPath"),True)
					endCommand "Fail"
					ErrorNO = 0   'raise a user-defined error
						'Err.Description = oObject & " disabled"
			End If
End if 
End Function

Function cScreenshot(FileName)
On Error Resume Next
startCommand "Screenshot Command", ""
	If True Then
		currenttime=Replace(Time,":","_")
		ScreenshotPath=Environment.Value("TestDir")&"\ScreenShot\"&FileName&"_"&currenttime&".bmp"
		
		Desktop.CaptureBitmap ScreenshotPath,false
		If Err.Number<>0 Then
			Print "[CMD-] ---Image Saving Faild"
			Call InsertIntoHTMLReport( "endOfTestStep","Screenshot", "Failed to Capture Screenshot ::"&Err.Description, "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			Err.Clear
		Else 
			Print "[CMD-] ---Image Saved:- "&ScreenshotPath
			Call InsertIntoHTMLReport( "endOfTestStep","Screenshot","Screenshot Captured ["&FileName&".bmp]" , "", True, Config.Item("ReportPath"),True)
			endCommand "Pass"
			ErrorNo = 0
		End If
		
	Else
		Call InsertIntoHTMLReport( "endOfTestStep","Screenshot", "Browser dose not Exists", "", False, Config.Item("ReportPath"),True)
		endCommand "Fail"
		ErrorNo = 8
	
	End If
	
End Function

Public Function IsCheckElementPresent (oObject, identifire, reTryTime)
	On Error Resume Next
	SendObject oObject, identifire
    If CommandObj.Exist(reTryTime)Then
		IsCheckElementPresent  = True
	Print "[CMD-] ---Element Present"
	Else
		IsCheckElementPresent  = False
	Print "[CMD-] ---Element Not Present"
    End If
End Function

Public Function cWriteToReport(comment)
On Error Resume Next
	startCommand "WriteToReport  Command", oObject
	Call InsertIntoHTMLReport("endOfTestStep","Write To Report",comment , "", True, Config.Item("ReportPath"),False)
	endCommand "Pass"
End Function

Function cCheckImagePresent(oObject, identifire, Assert,customErrorMessage)

On Error Resume Next
SendObject oObject, identifire
startCommand "CheckImagePresent  Command", oObject
Dim result,AssertType

AssertType=CBool(Assert)


 If CommandObj.Exist(retryTime)Then 
	If True=Cbool(Config.Item("HighLight")) Then
    	CommandObj.Highlight
    End If
		Call InsertIntoHTMLReport("endOfTestStep","Check Image Present","Element ["& oObject&"] is present in view" , "", True, Config.Item("ReportPath"),False)
		endCommand "Pass"

	Else
		Err.Description = "Element ["&oObject &"] is not present in view"
		If AssertType = True Then
			Call InsertIntoHTMLReport("endOfTestStep","Check Image Present",cCustomErrorGeneration(customErrorMessage)&Err.Description, "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNo = 8

		ElseIf AssertType = False Then
			Call InsertIntoHTMLReport("endOfTestStep","Check Image Present",cCustomErrorGeneration(customErrorMessage)&Err.Description , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNo = 0

		End If 

     End If
End Function

Public Function cType (oObject, identifire, sValue)
	On Error Resume Next
	
	SendObject oObject, identifire
	startCommand "Type Command", oObject
    If CommandObj.Exist(retryTime)Then
		If CommandObj.GetROProperty("disabled") = 0 Then
			If CommandObj.GetROProperty("html tag") = "SPAN" Then
				CommandObj.Click
            	CommandObj.Object.innertext = ""
            	CommandObj.Object.innertext = sValue
			Else
				CommandObj.Set sValue
			End If
			
			If Err.Number<>0 Then
				Err.Description="Command is Failed due to : "&Err.Description&" :: Element ["&oObject&"] :: Type Value ["&sValue&"]"
				Call InsertIntoHTMLReport( "endOfTestStep","Type",Err.Description , "", False, Config.Item("ReportPath"),True)
				endCommand "Fail"
				ErrorNO =  8    
				
			Else
				Call InsertIntoHTMLReport( "endOfTestStep","Type","Element ["&oObject&"] :: Type Value:-"&sValue, "", True, Config.Item("ReportPath"),False)
				endCommand "Pass"
			End If
		Else
			Err.Description = "Command is Failed due to Element ["&oObject &"] is disabled"
			Call InsertIntoHTMLReport( "endOfTestStep","Type",Err.Description , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO =  8    'raise a user-defined error
			
		End If
	Else
		Err.Description = "Command is Failed due to Element ["&oObject &"] is does not exist"
        Call InsertIntoHTMLReport( "endOfTestStep","Type",Err.Description , "", False, Config.Item("ReportPath"),True)
		endCommand "Fail"
		ErrorNO =  8    'raise a user-defined error
        
     End If
End Function

' Click Command-Drupasinghe-2012/01/03
Public Function cClick (oObject, identifire)
	On Error Resume Next
	
	SendObject oObject, identifire
	startCommand "Click  Command", oObject
    If CommandObj.Exist(retryTime)Then
	If True=Cbool(Config.Item("HighLight")) Then
    	CommandObj.Highlight
    End If
		If CommandObj.GetROProperty("disabled") = 0 Then
			CommandObj.Click
			'Browser("title:=.*").Page("title:=.*").Sync
			If Err.Number<>0 Then
			 
				Call InsertIntoHTMLReport( "endOfTestStep","Click","Command is Failed due to : "&Err.Description& " :: Element ["&oObject&"]" , "", False, Config.Item("ReportPath"),True)
				endCommand "Fail"
				ErrorNO =  8    'raise a user-defined error
				
			Else
			Call InsertIntoHTMLReport( "endOfTestStep","Click","clicked on Element ["&oObject&"]" , "", True, Config.Item("ReportPath"),False)
			endCommand "Pass"
			
			End If
		Else
			Err.Description ="Element ["&oObject&"] is disabled"
			Call InsertIntoHTMLReport( "endOfTestStep","Click","Command is Failed due to "&Err.Description , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO =  8    'raise a user-defined error
			
		End If
    Else
    	Err.Description ="Element ["&oObject&"] does not exist"
		Call InsertIntoHTMLReport( "endOfTestStep","Click","Command is Failed due to "&Err.Description , "", False, Config.Item("ReportPath"),True)
		endCommand "Fail"
		ErrorNO =  8    'raise a user-defined error
        
	End If
End Function

'RightClick
'last update on 22.04.2015
Public Function cRightClick (oObject, identifire)
	On Error Resume Next
	
	SendObject oObject, identifire
	startCommand "RightClick Command", oObject
    If CommandObj.Exist(retryTime)Then
		If CommandObj.GetROProperty("disabled") = 0 Then
			
			x = CommandObj.GetROProperty("width")/2
			y = CommandObj.GetROProperty("height")/2
			Setting.WebPackage("ReplayType") = 2
			CommandObj.FireEvent "onclick", x, y, micRightBtn
			Setting.WebPackage("ReplayType") = 1
			
			Call InsertIntoHTMLReport( "endOfTestStep","Right Click","Right Clicked on Element ["&oObject&"]" , "", True, Config.Item("ReportPath"),False)
			endCommand "Pass"
		Else
			Err.Description ="Command is Failed due to Element ["& oObject & "] is disabled"
			Call InsertIntoHTMLReport( "endOfTestStep","Right Click",Err.Description  , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO =  8    'raise a user-defined error
			
		End If
    Else
    	Err.Description = "Command is Failed due to Element ["&oObject &"] is does not exist"
		Call InsertIntoHTMLReport( "endOfTestStep","Right Click",Err.Description  , "", False, Config.Item("ReportPath"),True)
		endCommand "Fail"
		ErrorNO =  8    'raise a user-defined error
        
	End If
End Function

'Mouse Over 
'last update on 22.04.2015
Public Function cMouseOver (oObject, identifire)

	On Error Resume Next
	SendObject oObject, identifire
	startCommand "MouseOver Command", oObject
   
   If CommandObj.Exist(retryTime)Then
		
		If CommandObj.GetROProperty("disabled") = 0 Then
			Setting.WebPackage("ReplayType") = 2
			CommandObj.FireEvent "onmouseover"
			Setting.WebPackage("ReplayType") = 1
			If Err.Number<>0 Then
					Call InsertIntoHTMLReport( "endOfTestStep","Mouse Over","Failed to Mouse Over on Element ["&oObject&"] :: "&Err.Description , "", False, Config.Item("ReportPath"),True)
					endCommand "Fail"
					ErrorNO =  8    'raise a user-defined error
			Else
					Call InsertIntoHTMLReport( "endOfTestStep","Mouse Over","Mouse Overed on Element ["&oObject&"]" , "", True, Config.Item("ReportPath"),False)
					endCommand "Pass"
			End If
		Else
			Call InsertIntoHTMLReport( "endOfTestStep","Mouse Over","Failed to Mouse Over on Element ["&oObject&"] :: Element is Disabled", "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO =  8    'raise a user-defined error
		End If
    Else
		Call InsertIntoHTMLReport( "endOfTestStep","Mouse Over","Failed to Mouse Over Mouse Overed on Element ["&oObject&"] :: Element is does not exist" , "", False, Config.Item("ReportPath"),True)
		endCommand "Fail"
		ErrorNO =  8    'raise a user-defined error
	End If
End Function


'ClickAt Command
'last update on 22.04.2015
Public Function cClickAt (oObject, identifire, coordinates)
	On Error Resume Next
	
	SendObject oObject, identifire
    startCommand "ClickAt Command", oObject
    
	If CommandObj.Exist(retryTime)Then
	If True=Cbool(Config.Item("HighLight")) Then
    	CommandObj.Highlight
    End If
        X=split(coordinates, ",")(0)
        Y=split(coordinates, ",")(1)
        CommandObj.Click X,Y
		If Err.Number<>0 Then
			Call InsertIntoHTMLReport( "endOfTestStep","Click At",Err.Description&" :: Element ["&oObject&"]" , "", False, Config.Item("ReportPath"),True)
        	endCommand "Fail"
       	 	ErrorNO =  8    'raise a user-defined error
		Else
			 Call InsertIntoHTMLReport( "endOfTestStep","Click At","Clicked at "&X&"|"&Y&" on Element ["&oObject&"]" , "", True, Config.Item("ReportPath"),False)
        	endCommand "Pass"
		End If
       

    Else
     	Err.Description = "Command is Failed due to Element ["&oObject & "] does not exist"
        Call InsertIntoHTMLReport( "endOfTestStep","Click At",Err.Description , "", False, Config.Item("ReportPath"),True)
        endCommand "Fail"
        ErrorNO =  8    'raise a user-defined error 
    End If
End Function

'URL Open Command
Function cOpen (ByVal URL,ByVal identifire, ByVal WaitTime)
	On Error Resume Next
	URL=ResolveURL(URL,identifire)
	startCommand "Open Command", URL
    SystemUtil.Run Config.Item("AppType") &".exe",URL
	wait (Cint(WaitTime)/1000)
	Call InsertIntoHTMLReport( "endOfTestStep","Open","Url ["&URL&"]" , "", True, Config.Item("ReportPath"),False)
	endCommand "Pass"
End Function



'Supportive Function for open command.
'resolve the identifire in the URL
'@auther Vimukhi Hewapathirana
'Last updated on 22.04.2015
Function ResolveURL(URL,identifire)
print "[CMD-] ---URL="&Url
	If (identifire<>"") Then
	'Splitting Raw Identifires into array
		Raw_List=Split(identifire, "_PARAM,")
			'Resolving Identifires
			For i = 0 To Ubound(Raw_List) Step 1
				Identifire_List=Split(Raw_List(i),"_PARAM:")
				arrIdentifire(i)=Identifire_List(1)	
			Next
			
			Set re = CreateObject("VBScript.RegExp")
			re.Global = True   
			re.IgnoreCase = True
			re.Pattern = "<[^>]+>"
			x=0
			set match = re.Execute(Url)
			For each val in match
				Url=Replace(Url,val.value,arrIdentifire(x))
				x=x+1
			Next   	
			ResolveURL=Url
			print "[CMD-] ---RESOLVED_URL="&ResolveURL
		
	Else	
			ResolveURL=Url
	End IF
	
End Function

'Browser Close command
Function CloseApp(byVal WaitTime)
	On Error Resume Next
	startCommand "Close Command", oObject
	SystemUtil.CloseProcessByName Config.Item("AppType") &".exe"
	wait (WaitTime)
End Function


'Select Command  
'LastUpdated on 23.04.2015
Public Function cSelect( oObject, identifire, oOption)
	On Error Resume Next
    SendObject oObject, identifire
	startCommand "Select Command", oObject
	If CommandObj.Exist(retryTime)Then
		If True=Cbool(Config.Item("HighLight")) Then
    		CommandObj.Highlight
    	End If
		optionArray= split(oOption, "=")
		If  optionArray(0)="index" Then
			If CommandObj.GetROProperty("disabled") = 0 Then
			
				If optionArray(0)="index" Then
	                CommandObj.Select (CInt(optionArray(1))-1)
				Else
					CommandObj.RegexSelectDOM oOption
				End If
				
				If Err.Number<>0 Then
					Call InsertIntoHTMLReport( "endOfTestStep","Select - index","Command Failed Due to "&Err.Description&" :: Option Value = ["&optionArray(1)&"]" , "", False, Config.Item("ReportPath"),True)
					endCommand "Fail"
					ErrorNO =  8    'raise a user-defined error
				Else
					Call InsertIntoHTMLReport( "endOfTestStep","Select - index","Element ["&oObject&"] :: index ["&optionArray(1)&"] is selected." , "", True, Config.Item("ReportPath"),False)
					endCommand "Pass"
				End If
				
			Else
				Err.Description ="Command is Failed due to  Element ["&oObject & "] is disabled"
				Call InsertIntoHTMLReport( "endOfTestStep","Select - index",Err.Description , "", False, Config.Item("ReportPath"),True)
				endCommand "Fail"
				ErrorNO =  8    'raise a user-defined error
        	End If
			
		'if multiselesct option	
		Elseif optionArray(0)="ExtendSelect" Then
			selectOptionArray= split(optionArray(1), "|")
			result1 = "pass"
				For i = 0 To Ubound(selectOptionArray) Step 1
				CommandObj.ExtendSelect selectOptionArray(i)
					If Err.Number<>0 Then	  	
						Call InsertIntoHTMLReport( "endOfTestStep","Select",Err.Description&" :: Option Value = ["&selectOptionArray(i)&"]" , "", False, Config.Item("ReportPath"),True)
						endCommand "Fail"
						ErrorNO =  8    'raise a user-defined error
						result1 = "fail"
						Exit For
					 End If	
				Next
			If result1 = "pass" Then
		    Call InsertIntoHTMLReport("endOfTestStep","Select","Element ["&oObject&"] :: Option ["&optionArray(1)&"] are Selected" , "", True, Config.Item("ReportPath"),False)
			endCommand "Pass"
			End If
		'if not index value came directly
		Else
			If CommandObj.GetROProperty("disabled") = 0 Then  
				CommandObj.Select oOption
			    If Err.Number<>0 Then	  	
					Call InsertIntoHTMLReport( "endOfTestStep","Select",Err.Description&" :: Option Value = ["&oOption&"]" , "", False, Config.Item("ReportPath"),True)
					endCommand "Fail"
					ErrorNO =  8    'raise a user-defined error
				Else
				    Call InsertIntoHTMLReport( "endOfTestStep","Select","Element ["&oObject&"] :: Option ["&oOption&"] is Selected" , "", True, Config.Item("ReportPath"),False)
				    endCommand "Pass"
				 End If
			Else
				Err.Description ="Command is Failed due to ["&oObject & "] is disabled"
				Call InsertIntoHTMLReport( "endOfTestStep","Select",oObject , "", False, Config.Item("ReportPath"),True)
				endCommand "Fail"
				ErrorNO =  8    'raise a user-defined error
				
	        End If
		End If
    Else
    	Err.Description ="Command is Failed due to Element ["&oObject & "] is does not exist"
		Call InsertIntoHTMLReport( "endOfTestStep","Select",oObject , "", False, Config.Item("ReportPath"),True)
		endCommand "Fail"
		ErrorNO =  8    'raise a user-defined error
    End If
End Function


'  Drupasinghe 2013/02/07 - This is fynalyzed 
'  This can be used to verify the value of a drop down list or object properties of an object
 
' Option 1 - verifyDropDownValuesPresent

' Expected should be paased as - "Danushka;Nadie;Damith"

' Option 1 - verifyDropDownValuesPresent
' Property Name should be passes as - "verifyDropDownValuesPresent"
 ' Expected should be paased as - "Danushka;Nadie;Damith"

' Option 2- verifyDropDownValuesNotPresent
' Property Name should be passes as - "verifyDropDownValuesNotPresent"
' Expected should be paased as - "Danushka;Nadie;Damith"
 

' Option 3- Verifying a perperty of an pbject  Example - innertext 
' Property Name should be passes as - "innertextt"
' Expected should be paased as - "Danushka"

'  Sample Link - http://www.w3schools.com/tags/tryit.asp?filename=tryhtml_select  ----Testing purposes only
 '  Sample object -Browser("Browser").Page("Tryit Editor v1.6").Frame("Frame").WebList("select")  ----Testing purposes only

Function cCheckObjectProperty (oObject, identifire, byval propertyName, expectedValue, AssertType, customErrorMessage)
	On Error Resume Next
	SendObject oObject, identifire
	startCommand "CheckObjectProperty Command", oObject
	AssertType=Cbool(AssertType)
	Pattern=split(expectedValue,"regex_")
	regexarr=Ubound(Pattern)
	MatchCount=0
	
	If propertyName = "ELEMENTPRESENT" Then
		veifyElementPresent oObject,CommandObj, expectedValue,AssertType,customErrorMessage
	ElseIf CommandObj.Exist(retryTime)Then
			If propertyName = "ALLOPTIONS" Then
				veifyDropDownValuePresent oObject,CommandObj, expectedValue,AssertType,customErrorMessage
			Elseif propertyName = "MISSINGOPTION" Then
				veifyDropDownValueNotPresent oObject,CommandObj, expectedValue,AssertType,customErrorMessage
			Elseif propertyName = "SELECTEDOPTION" Then
				veifyDropDownValueSelected oObject,CommandObj, expectedValue,AssertType,customErrorMessage
			'-----------------------------------------------------------------------
			ElseIf propertyName = "PROPERTYPRESENT" Then
				veifyPropertyPresent oObject,CommandObj, expectedValue,AssertType,customErrorMessage
			'-----------------------------------------------------------------------
			ElseIf CStr(CommandObj.GetROProperty(propertyName))=Cstr(expectedValue) then
				Call InsertIntoHTMLReport( "endOfTestStep","Check Object Property","Element:-"&oObject&" ::: Expected Value "&expectedValue&" is Found in the Property "&propertyName , "", True, Config.Item("ReportPath"),False)
				endCommand "Pass"    

			ElseIf regexarr=1 Then
				Value=CommandObj.GetROProperty(propertyName)
				Dim re
				Set re = CreateObject("vbscript.regexp") 
				re.Pattern = Pattern(1)
				re.IgnoreCase = True
				re.Global = True
				    
				 For Each match In re.Execute(Value)
				 MatchCount=MatchCount+1
				 next
				
				If MatchCount>0 Then
					Call InsertIntoHTMLReport( "endOfTestStep","check Object Property","Element ["&oObject&"] :: Expected regular expression value ["&Pattern(1)&"] for Property ["&propertyName&"] is found" , "", True, Config.Item("ReportPath"),False)
					endCommand "Pass"
				Else
					If AssertType = True  Then
                        Call InsertIntoHTMLReport("endOfTestStep", "check Object Property", cCustomErrorGeneration(customErrorMessage)&"Element [" & oObject&"] :: Expected regular expression value ["&Pattern(1)&"] for Property ["&propertyName&"] is not found" , "", False, Config.Item("ReportPath"),True)
						endCommand "Fail"
						ErrorNo = errorRaiseNo    'raise a user-defined error
						Err.Description = "object property not found"
					else 
	                    Call InsertIntoHTMLReport( "endOfTestStep", "check Object Property", cCustomErrorGeneration(customErrorMessage)&"Element [" & oObject&"] :: Expected regular expression value ["&Pattern(1)&"] for Property ["&propertyName&"] is not found" , "", False, Config.Item("ReportPath"),True)
						endCommand "Fail"
						ErrorNo = 0
						End If
				End If
			Else 
					Err.Description = "Element ["&oObject&"] object property ["&propertyName&"] is not found"
					If AssertType = True  Then
						
                        Call InsertIntoHTMLReport("endOfTestStep","Check Object Property",cCustomErrorGeneration(customErrorMessage) &Err.Description , "", False, Config.Item("ReportPath"),True)
						endCommand "Fail"
						ErrorNO =  8    'raise a user-defined error
						
					else 
	                    Call InsertIntoHTMLReport( "endOfTestStep","Check Object Property",cCustomErrorGeneration(customErrorMessage) &Err.Description, "", False, Config.Item("ReportPath"),True)
						endCommand "Fail"
						ErrorNO = 0    'raise a user-defined error
						End If
			End If                             
	Else
			Err.Description = " Element ["&oObject&"] does not exist"
			Call InsertIntoHTMLReport( "endOfTestStep","Check Object Property",cCustomErrorGeneration(customErrorMessage) &Err.Description, "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO =  8    'raise a user-defined error
			
   End If
End Function


'  This can be used to verify the value of a drop down list
 '  User need to pass to object and  expected values
' expected values should be in - "Danushka;Nadie;Damith" - format
'  Sample object -Browser("Browser").Page("Tryit Editor v1.6").Frame("Frame").WebList("select")  ----Testing purposes only
Function veifyDropDownValuePresent (oObject,cmdObject, expectedValues,AssertDrpPresent,CustomErrorMsg)
	On Error Resume Next
	
	actualValues  = cmdObject.GetROProperty("all items")
	If isEmpty(actualValues) Then
	
		If  AssertDrpPresent = True Then
				Call InsertIntoHTMLReport( "endOfTestStep","Check Object Property - ALL OPTIONS","Element ["&oObject&"] is a not supported object (WebList) or No Value Found in the Select" , "", False, Config.Item("ReportPath"),True)
				endCommand "Fail"
				ErrorNO =  8    'raise a user-defined error
		Else 
				Call InsertIntoHTMLReport( "endOfTestStep","Check Object Property - ALL OPTIONS","Element ["&oObject&"] is a not supported object (WebList) or No Value Found in the Select" , "", False, Config.Item("ReportPath"),True)
				endCommand "Fail"
				ErrorNO = 0    'raise a user-defined error
			End if 
	
		
		
	Else
	
		If  actualValues = expectedValues Then
			Call InsertIntoHTMLReport( "endOfTestStep","Check Object Property - ALL OPTIONS ","Element ["&oObject&"] :: Expected Options are Found in the Select or List" , "", True, Config.Item("ReportPath"),False)
			endCommand "Pass" 
	    Else 
	    	Err.Description = "Element ["&oObject&"] Expected values are Not Found in the List or Select"
			If  AssertDrpPresent = True Then
				Call InsertIntoHTMLReport( "endOfTestStep","Check Object Property - ALL OPTIONS",Err.Description , "", False, Config.Item("ReportPath"),True)
				endCommand "Fail"
				ErrorNO =  8    'raise a user-defined error
				
			Else 
				Call InsertIntoHTMLReport( "endOfTestStep","Check Object Property - ALLOPTIONS",Err.Description , "", False, Config.Item("ReportPath"),True)
				endCommand "Fail"
				ErrorNO = 0    'raise a user-defined error
			End if 
	   End If
		
	End If
	
End Function


'  Drupasinghe 2013/02/07 - This is fynalyzed 
'  This can be used to verify the given values are not present in drop down list
 '  User need to pass to object and  expected values
' expected values should be in - "Danushka;Nadie;Damith" - format
'  Sample object -Browser("Browser").Page("Tryit Editor v1.6").Frame("Frame").WebList("select")  ----Testing purposes only


Function veifyDropDownValueNotPresent (oObject,cmdObject, expectedValues1,AssertDrpNotPresent,CustomErrorMsg)
	On Error Resume Next
	ExpectedErr = false
	actualValues1  = cmdObject.GetROProperty("all items")
	
	'If returens a Empty resutl Fail the step
	If IsEmpty(actualValues1) Then 
		
		Err.Description = "Element ["&oObject&"] - ["&cmdObject.ToString&"] is a not Supported object (WebSelect,WebList) or No value present in the Select or List"
		
		If  AssertDrpNotPresent = True Then
			Call InsertIntoHTMLReport("endOfTestStep","Check Object Property - MISSING OPTION",cCustomErrorGeneration(CustomErrorMsg)&Err.Description , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO =  8    'raise a user-defined error			
		Else 
	        Call InsertIntoHTMLReport( "endOfTestStep","Check Object Property - MISSING OPTION",cCustomErrorGeneration(CustomErrorMsg)&Err.Description , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO = 0    'clearing the user-defined error
		End If 
	
	'Result set is returend - check for not present elements
	Else
	
	'checking for the missing option
		expectedArray = Split(expectedValues1, ";")
		actualArray = Split(actualValues1, ";")
	
		For i = 0 To Ubound(expectedArray)
			For j = 0 To Ubound(actualArray)
			'if found ExpectedErr=True
				If expectedArray(i) = actualArray(j)Then
					ExpectedErr = true
					Found=expectedArray(i)
	            End If
			Next
		Next
	
		If ExpectedErr = true Then
		Err.Description = "Element ["&oObject&"] :: Missing Option ["&Found&"] is Found in the List or Select"
					If  AssertDrpNotPresent = True Then
						Call InsertIntoHTMLReport("endOfTestStep","Check Object Property - MISSING OPTION",cCustomErrorGeneration(CustomErrorMsg)&Err.Description , "", False, Config.Item("ReportPath"),True)
						endCommand "Fail"
						ErrorNO =  8    'raise a user-defined error
						
	
					else 
		                Call InsertIntoHTMLReport( "endOfTestStep","Check Object Property - MISSING OPTION",cCustomErrorGeneration(CustomErrorMsg)&Err.Description , "", False, Config.Item("ReportPath"),True)
						endCommand "Fail"
						ErrorNO = 0    'raise a user-defined error
	
					end if 
		Else
			Call InsertIntoHTMLReport("endOfTestStep","Check Object Property - MISSING OPTION","Element ["&oObject&"] :: Expected Missing Option Value(s) are Not Present" , "", True, Config.Item("ReportPath"),False)
			endCommand "Pass"
		End If
	End IF
End Function


 
Function veifyDropDownValueSelected (oObject,cmdObject, expectedValues,AssertDrpSelected,CustomErrorMsg)
	On Error Resume Next
	AssertDrpSelected=CBool(AssertDrpSelected)
	actualValues  = cmdObject.GetROProperty("selection")
	
	If IsEmpty(actualValues) Then 
		
		Err.Description = "Element ["&oObject&"]-["&cmdObject.ToString&"] is a not Supported object (WebSelect,WebList) or No value selected in the Select or List"
		
		If  AssertDrpNotPresent = True Then
			Call InsertIntoHTMLReport("endOfTestStep","Check Object Property - SELECTED OPTION",cCustomErrorGeneration(CustomErrorMsg)&Err.Description , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO =  8    'raise a user-defined error			
		Else 
	        Call InsertIntoHTMLReport( "endOfTestStep","Check Object Property - SELECTED OPTION",cCustomErrorGeneration(CustomErrorMsg)&Err.Description , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO = 0    'clearing the user-defined error
		End If 
	Else
		If  actualValues = expectedValues Then
			Call InsertIntoHTMLReport( "endOfTestStep","Check Object Property - SELECTED OPTION","Element ["&oObject&"] :: Expected value ["&expectedValues&"] is Selected" , "", True, Config.Item("ReportPath"),False)
			endCommand "Pass" 
	    Else
			Err.Description ="Element:-"&oObject &" :: Expected Selected Value is ["&expectedValues&"] Found is ["&actualValues&"]"	    
			If  AssertDrpSelected = True Then
				Call InsertIntoHTMLReport( "endOfTestStep","Check Object Property - SELECTED OPTION",cCustomErrorGeneration(CustomErrorMsg)&Err.Description, "", False, Config.Item("ReportPath"),True)
				endCommand "Fail"
				ErrorNO = 8    'raise a user-defined error
						
			Else 
				Call InsertIntoHTMLReport( "endOfTestStep","Check Object Property - SELECTED OPTION",cCustomErrorGeneration(CustomErrorMsg)&Err.Description , "", False, Config.Item("ReportPath"),True)
				endCommand "Fail"
				ErrorNO = 0    'Clear a user-defined error
			end if 
	   End If
	   End if
End Function


Function veifyElementPresent (oObject,cmdObject, expectedValue,AssertElePresent,CustomErrorMsg)
	On Error Resume Next
	
	actualValue = cmdObject.Exist(2)
	
	intCompare = StrComp(actualValue, expectedValue, vbTextCompare)
	
	If intCompare = 0 Then
		Call InsertIntoHTMLReport( "endOfTestStep","Check Object Property - Element Present","Element ["&oObject&"] :: Expected is ["&expectedValue&"] Actual is ["&actualValue&"]" , "", True, Config.Item("ReportPath"),False)
		endCommand "Pass"
	Else
		Err.Description = "Element ["&oObject&"] ::  Expected is ["&expectedValue&"] Found is ["&actualValue&"]"
		If  AssertElePresent = True Then
			Call InsertIntoHTMLReport( "endOfTestStep","Check Object Property - Element Present",cCustomErrorGeneration(CustomErrorMsg)&Err.Description, "", False, Config.Item("ReportPath"),True)
					endCommand "Fail"
					ErrorNO = 8    'raise a user-defined error
					
		else 
					Call InsertIntoHTMLReport( "endOfTestStep","Check Object Property - Element Present",cCustomErrorGeneration(CustomErrorMsg)&Err.Description , "", False, Config.Item("ReportPath"),True)
					endCommand "Fail"
					ErrorNO = 0    'raise a user-defined error
		end if 
	End If
	
End Function

Function veifyPropertyPresent (oObject,cmdObject, expectedValues,AssertPropertyPresent,CustomErrorMsg)
	On Error Resume Next
	expectedValuesArray = Split(expectedValues, "|")
	
	
	If UBound(expectedValuesArray) = 1 Then
	
		propertyName = expectedValuesArray(0)
		expectedValue = Cbool(expectedValuesArray(1))
		
		If IsEmpty(cmdObject.GetROProperty(propertyName)) Then
			actualValue=False
		Else
			actualValue = cmdObject.CheckProperty(propertyName, cmdObject.GetROProperty(propertyName))
		End If

		If expectedValue=cBool(actualValue) Then
		 	
		 	If expectedValue=True Then
				Call InsertIntoHTMLReport( "endOfTestStep","Check Object Property - PROPERTY PRESENT","Element ["&oObject&"] :: Property ["&propertyName&"] is Present", "", True, Config.Item("ReportPath"),False)
				endCommand "Pass"
			Else
				Call InsertIntoHTMLReport( "endOfTestStep","Check Object Property - PROPERTY PRESENT","Element ["&oObject&"] :: Property ["&propertyName&"] is Not Present", "", True, Config.Item("ReportPath"),False)
				endCommand "Pass"
			End If
		Else
			Err.Description ="Element ["&oObject&"] ::  Property ["&propertyName&"] Expected to be ["&expectedValue&"]. Actual is ["&actualValue&"]"
			If  AssertPropertyPresent = True Then
				Call InsertIntoHTMLReport( "endOfTestStep","Check Object Property - PROPERTY PRESENT",cCustomErrorGeneration(CustomErrorMsg)&Err.Description , "", False, Config.Item("ReportPath"),True)
				endCommand "Fail"
				ErrorNO = 8    'raise a user-defined error
				
			else 
				Call InsertIntoHTMLReport( "endOfTestStep","Check Object Property - PROPERTY PRESENT",cCustomErrorGeneration(CustomErrorMsg)&Err.Description , "", False, Config.Item("ReportPath"),True)
				endCommand "Fail"
				ErrorNO = 0    'Clear a user-defined error
			end if
		End If
		
	Else
		Err.Description ="Element ["&oObject&"] :: Invalid Input Parameters"
		If  AssertPropertyPresent = True Then
			Call InsertIntoHTMLReport( "endOfTestStep","Check Object Property - PROPERTY PRESENT",cCustomErrorGeneration(CustomErrorMsg)&Err.Description , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO = 8    'raise a user-defined error
			Exit Function
			
		Else 
			Call InsertIntoHTMLReport( "endOfTestStep","Check Object Property - PROPERTY PRESENT",cCustomErrorGeneration(CustomErrorMsg)&Err.Description , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO = 0    'Clear a user-defined error
			Exit Function
		End If
	
		
	End If
	
	
End Function



 Public Function cCheckPattern (oObject, identifire ,pattern, AssertType, customErrorMessage)
	On Error Resume Next
	SendObject oObject, identifire
	startCommand "CheckPattern Command", oObject
	'Browser("title:=.*").Page("title:=.*").Sync
    If CommandObj.Exist(retryTime)Then
    	
    	Set regEx = New RegExp
		regEx.Pattern = pattern				'	"\d{2}-\d{2}-\d{4}"
		regEx.Global = True
    	
    	checkStr = CommandObj.GetROProperty("value")

    	If checkStr <> Empty Then
    		count = regEx.Execute(checkStr).Count
    		If count = 1 Then
    			Call InsertIntoHTMLReport( "endOfTestStep","Check Pattern",oObject , "", True, Config.Item("ReportPath"),False)
				endCommand "Pass"
			Else
				If  AssertType = True Then
					Call InsertIntoHTMLReport( "endOfTestStep","Check Pattern",oObject , "", False, Config.Item("ReportPath"),True)
					endCommand "Fail"
					ErrorNO = 8    'raise a user-defined error
					Err.Description = oObject & " Wrong Regular Expression"
				Else 
					Call InsertIntoHTMLReport( "endOfTestStep","Check Pattern",oObject , "", False, Config.Item("ReportPath"),True)
					endCommand "Fail"
					ErrorNO = 0    'raise a user-defined error
				End If
			End If
    	End If
	Else
		Call InsertIntoHTMLReport( "endOfTestStep","Check Pattern",cCustomErrorGeneration(customErrorMessage) & oObject , "", False, Config.Item("ReportPath"),True)
		endCommand "Fail"
		ErrorNO =  8    'raise a user-defined error
		Err.Description = oObject & " does not exist"
	End If
End Function

'		Drupasinghe 2013/02/19 - Finalyzed
'      This is a table validation 
' 		Option 1 - Row Number Validation 
' 		Option 2- Clolumn  Number Validation 
' 		Option 3 - Table Cell Data    Validation 
' 		Option 4 - Table  Relative Cell Data    Validation 

'      Relative method
'      1. One relatice check :-  "RELATIVE", "Table cell 1,1,Table cell 2"
 '      2. Myltiple relative checks :-  "RELATIVE", "Table cell 1,1,Table cell 2#Table cell 1,3,Table cell 2"


Function cCheckTable (oObject, identifire, validationType,expectedValue,AssertType, customError)
	On Error Resume Next
	SendObject oObject, identifire
	startCommand "Check Table Command", oObject
	AssertType=Cbool(AssertType)
	If CommandObj.Exist(retryTime)Then
	If True=Cbool(Config.Item("HighLight")) Then
    	CommandObj.Highlight
    End If
		If  validationType="ROWCOUNT" AND IsNumeric(expectedValue)=true  Then
			validateTableRowCount oObject,CommandObj, expectedValue,AssertType,customError
		ElseIf validationType = "COLCOUNT" AND IsNumeric(expectedValue)=true Then
			validateTableColCount oObject,CommandObj, expectedValue,AssertType,customError
		ElseIf  validationType = "TABLECELL" Then 
			validateTableCell oObject,CommandObj,expectedValue ,AssertType,customError
		ElseIf  validationType = "RELATIVE" Then 
			validateTableOffset oObject,CommandObj, expectedValue, AssertType,customError
		ElseIf  validationType = "TABLEDATA" Then 
			validateTableData oObject,CommandObj, expectedValue, AssertType,customError		
		Else
			Call InsertIntoHTMLReport( "endOfTestStep", "Check Table", cCustomErrorGeneration(customError) &"Element ["& oObject&"]  :: Validation type not found or expected value is not a numeric value" , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO = 8    'raise a user-defined error
			Err.Description = "table validation type not found"			
		End if
	Else 
		Err.Description = "Element ["&oObject &"] does not exist"	
		Call InsertIntoHTMLReport( "endOfTestStep", "Check Table", cCustomErrorGeneration(customError) & Err.Description , "", False, Config.Item("ReportPath"),True)
		endCommand "Fail"	
		ErrorNO = 8    'raise a user-defined error
				
	End If
End Function 

'	Table Data validation 
Public function validateTableData (oObject,cmdObject,expectedValue,AssertTableCell,customError)
	On Error Resume Next
	SendObject oObject, identifire
	Dim rowCount, colCount
	Dim val
	Dim data
	Set ExcelStore=CreateObject("System.Collections.ArrayList")
	val = Split(expectedValue, ",")
    For Each text  In val
		   rowCount=cmdObject.RowCount
		       For i=1 to rowCount
                      colCount=cmdObject.ColumnCount(i)
                 For j=1 to colCount
					 If cmdObject.GetCellData(i,j)=text  Then
						 ExcelStore.add("pass")
					 Else
					     'do nothing
					 End If
                 Next
               Next     
    Next
   If UBound(val)+1=ExcelStore.Count Then
	   		Call InsertIntoHTMLReport( "endOfTestStep","Check Table - TABLEDATA","Element ["&oObject&"]"  , "", True, Config.Item("ReportPath"),False)
		    endCommand "Pass"	
   Else
		If AssertTableCell = True  Then
            Call InsertIntoHTMLReport("endOfTestStep","Check Table - TABLEDATA",customError&"Element ["&oObject&"]"  , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO = 8    'raise a user-defined error
			Err.Description = "Table Cell value error"

		Else 
	        Call InsertIntoHTMLReport( "endOfTestStep","Check Table - TABLEDATA",customError&"Element ["&oObject&"]" , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO = 8    'raise a user-defined error
			Err.Description = "Table Cell value error"
		End If
		End If		
End Function

'''' Table Row Count Validaion 
Public Function validateTableRowCount (oObject,cmdObject,expectedValue,AssertTypeRowCwnt,CustomErrorMsg)
	On Error Resume Next
	AssertTypeRowCwnt=Cbool(AssertTypeRowCwnt)
	ActualValue=cmdObject.RowCount
	
	If Err.Number<>0 Then
		If AssertTypeRowCwnt = True  Then
            Call InsertIntoHTMLReport("endOfTestStep","Check Table - ROW COUNT",cCustomErrorGeneration(CustomErrorMsg)&Err.Description&" :: Element ["&oObject&"]" , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO =  8    'raise a user-defined error
			Exit Function
			
		Else 
	        Call InsertIntoHTMLReport( "endOfTestStep","Check Table - ROW COUNT",cCustomErrorGeneration(CustomErrorMsg)&Err.Description&" :: Element ["&oObject&"]", "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO = 0    'raise a user-defined error
			Exit Function
		End If
	End If
	
	If  CInt(ActualValue)= CInt(expectedValue) Then
	
		Call InsertIntoHTMLReport( "endOfTestStep","Check Table - ROW COUNT","Element ["&oObject&"] :: Expected Row Count ["&expectedValue&"] is found" , "", True, Config.Item("ReportPath"),False)
		endCommand "Pass"		
	Else 
		Err.Description="Element ["&oObject&"] :: Expected Row Count is ["&expectedValue&"] and Found is ["&ActualValue&"]"
		If AssertTypeRowCwnt = True  Then
            Call InsertIntoHTMLReport("endOfTestStep","Check Table - ROW COUNT",cCustomErrorGeneration(CustomErrorMsg)&Err.Description , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO =  8    'raise a user-defined error
		
		else 
	        Call InsertIntoHTMLReport( "endOfTestStep","Check Table - ROW COUNT",cCustomErrorGeneration(CustomErrorMsg)&Err.Description , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO = 0    'raise a user-defined error
		End If
	
	End if 
End Function


'	Table Column Count validation 
Public Function validateTableColCount (oObject,cmdObject,expectedValue,AssertTypeColCwnt,CustomErrorMsg)
	On Error Resume Next
	ActualValue=cmdObject.Columncount(1)
	
	If Err.Number<>0 Then
		If AssertTypeRowCwnt = True  Then
            Call InsertIntoHTMLReport("endOfTestStep","Check Table - COLCOUNT",cCustomErrorGeneration(CustomErrorMsg)&Err.Description&" :: Element ["&oObject&"]", "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO =  8    'raise a user-defined error
			Exit Function
			
		Else 
	        Call InsertIntoHTMLReport( "endOfTestStep","Check Table - COLCOUNT",cCustomErrorGeneration(CustomErrorMsg)&Err.Description&" :: Element ["&oObject&"]" , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO = 0    'raise a user-defined error
			Exit Function
		End If
	End If
	
	If  Cint(ActualValue) = Cint(expectedValue) Then
		Call InsertIntoHTMLReport( "endOfTestStep","Check Table - COLCOUNT","Element ["&oObject&"] :: Expected Column Count ["&expectedValue&"] is found" , "", True, Config.Item("ReportPath"),False)
		endCommand "Pass"	
	Else 
	Err.Description="Element ["&oObject&"] :: Expected Column Count is ["&expectedValue&"] and Found is ["&ActualValue&"]"
		If AssertTypeColCwnt = True  Then
            Call InsertIntoHTMLReport("endOfTestStep","Check Table - COLCOUNT",cCustomErrorGeneration(CustomErrorMsg)&Err.Description , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO =  8    'raise a user-defined error
			Err.Description = "Column Count Error"
		else 
	        Call InsertIntoHTMLReport( "endOfTestStep","Check Table - COLCOUNT",cCustomErrorGeneration(CustomErrorMsg)&Err.Description , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO = 0    'raise a user-defined error
		End If
	End if     	
End Function

'		Table Cell validation 
Public function validateTableCell (oObject,cmdObject,expectedValue,AssertTableCell,customError)
	On Error Resume Next
	Dim row, col,a,b
	a = Split(expectedValue, "|")
	If Ubound(a)<>2 Then
			If AssertType = True  Then
                Call InsertIntoHTMLReport("endOfTestStep", "Check Table - TABLECELL", cCustomErrorGeneration(customError) &". Element ["& oObject&"] :: Invlid Format or Invalid Parameter Formation ("&expectedValue&")" , "", False, Config.Item("ReportPath"),False)
				endCommand "Fail"
				ErrorNo = errorRaiseNo    'raise a user-defined error
				Err.Description = "object property not found"
			Else 
	            Call InsertIntoHTMLReport( "endOfTestStep", "Check Table - TABLECELL", cCustomErrorGeneration(customError) &". Element ["& oObject&"] :: Invlid Format or Invalid Parameter Formation ("&expectedValue&")" , "", False, Config.Item("ReportPath"),False)
				endCommand "Fail"
				ErrorNo = 0
			End If
	End If
	row = CInt(a(0))
	col = CInt(a(1))
	expectedValue = a(2)
	b=Split(a(2),"regex_")
	Arrlength=UBound(b)
	If Arrlength=1 Then
		Value=cmdObject.GetCellData(row,col)
		Dim reg
		Set reg = CreateObject("vbscript.regexp") 
		reg.Pattern = b(1)
		reg.IgnoreCase = True
		reg.Global = True
				    
		For Each match In reg.Execute(Value)
			MatchCount=MatchCount+1
		next
				
		If MatchCount>0 Then
			Call InsertIntoHTMLReport( "endOfTestStep","Check Table - TABLECELL", "Element ["&oObject&"] :: Expected regular expression value ["&b(1)&"] is found" , "", True, Config.Item("ReportPath"),False)
			endCommand "Pass"
		Else
			If AssertType = True  Then
                Call InsertIntoHTMLReport("endOfTestStep", "Check Table - TABLECELL", cCustomErrorGeneration(customError) &" Element ["& oObject&"] :: Expected regular expression value ["&b(1)&"] is not found" , "", False, Config.Item("ReportPath"),True)
				endCommand "Fail"
				ErrorNo = errorRaiseNo    'raise a user-defined error
				Err.Description = "object property not found"
			Else 
	            Call InsertIntoHTMLReport( "endOfTestStep", "Check Table - TABLECELL", cCustomErrorGeneration(customError) &" Element ["& oObject&"] :: Expected regular expression value ["&b(1)&"] is not found" , "", False, Config.Item("ReportPath"),True)
				endCommand "Fail"
				ErrorNo = 0
			End If
		End If
	
	ElseIF  cmdObject.GetCellData(row,col) = expectedValue Then
		Call InsertIntoHTMLReport( "endOfTestStep","Check Table - TABLECELL","Element ["&oObject&"] expected value is Found In the Cell", "", True, Config.Item("ReportPath"),False)
		endCommand "Pass"	
	Else
	
		If AssertTableCell = True  Then
            Call InsertIntoHTMLReport("endOfTestStep","Check Table - TABLECELL", cCustomErrorGeneration(customError)&"Element ["&oObject&"] :: Expect Value ["&expectedValue&"] is not found in the relevant cell" , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO =  8    'raise a user-defined error
			Err.Description = "Table Cell value error"
		else 
	        Call InsertIntoHTMLReport( "endOfTestStep","Check Table - TABLECELL",cCustomErrorGeneration(customError)&"Element ["&oObject&"] :: Expect Value ["&expectedValue&"] is not found in the relevant cell" , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO = 0    'raise a user-defined error
		End If
			
	End If
End Function

'		Table Relative Data validation 
Public Function validateTableOffset(oObject,cmdObject, expectedValue,AssertTableOffSet,customError)
	On Error Resume Next
	Dim outerArray
	Dim innerArray
	Dim rd, o, ev

	outerArray = Split(expectedValue, "#")

	For i = 0 To Ubound(outerArray)
		innerArray = Split(outerArray(i),",")
		rd = innerArray(0)
		o = Cint(innerArray(1))
		ev = innerArray(2)
		validateCellOffset oObject,cmdObject, rd, o, ev,AssertTableOffSet
	Next
End Function

'		Relative Data Cell Validation
Public Function validateCellOffset(oObject,cmdObject, referenceData, offset, expectedValue, AsserValueTableOffSet)
	On Error Resume Next
	For row = 1 To cmdObject.RowCount
		For col=1 To cmdObject.ColumnCount(row)
			If cmdObject.GetCellData(row, col) = referenceData Then
				If cmdObject.GetCellData(row, col + offset) = expectedValue Then
					Call InsertIntoHTMLReport( "endOfTestStep","Check Table - RELATIVE","Element ["&oObject&"]" , "", True, Config.Item("ReportPath"),False)
					endCommand "Pass"	
				Else
					If AsserValueTableOffSet = True  Then
                        Call InsertIntoHTMLReport("endOfTestStep","Check Table - RELATIVE",customError&"Element ["&oObject&"]" , "", False, Config.Item("ReportPath"),True)
						endCommand "Fail"
						ErrorNO =  8    'raise a user-defined error
						Err.Description = "offset error"
					else 
	                    Call InsertIntoHTMLReport( "endOfTestStep","Check Table - RELATIVE",customError&"Element ["&oObject&"]" , "", False, Config.Item("ReportPath"),True)
						endCommand "Fail"
						ErrorNO = 0    'raise a user-defined error
						End If		
				End If
			End If
		Next
	Next
End Function


' Fail the test case 
' if message typed as "True:_xxxxxxxx" or "False:_xxxxxxxx" True will fail the test case and False will continue the testcase


Public Function cFail(message)
	On Error Resume Next
	value=Split(message,":_")
	max=Ubound(value)
	
	If max=0 then
		Call InsertIntoHTMLReport( "endOfTestStep","Fail", message & oObject , "", False, Config.Item("ReportPath"),True)
		endCommand "Fail"
		ErrorNO = 8
	Else
	
		value(0)=Cbool(value(0))
		If ERR.Number<>0 Then
			Call InsertIntoHTMLReport( "endOfTestStep","Fail", "Incorrect Parameters used for fail command:- "&value(0)&" - "& oObject , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO = 8
		ElseIf value(0)=True Then
			Call InsertIntoHTMLReport( "endOfTestStep","Fail", value(1) &" : "& oObject , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO = 8
		Else
			Call InsertIntoHTMLReport( "endOfTestStep","Fail", value(1)  &" : "&  oObject , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO = 0
			
		End If
		
	End If
	
    	
End Function






'Public Function cSelectWindow
Public Function cSelectWindow (oObject, identifire)
	On Error Resume Next
	SendObject oObject, identifire
	startCommand "SelectWindow Command", oObject
    If CommandObj.Exist(retryTime)Then
	If True=Cbool(Config.Item("HighLight")) Then
    	CommandObj.Highlight
    End If
			CommandObj.Activate
			'Browser("title:=.*").Page("title:=.*").Sync 'DR
			CommandObj.Maximize
			'Browser("title:=.*").Page("title:=.*").Sync 'DR
			Call InsertIntoHTMLReport( "endOfTestStep","Select Window",oObject , "", True, Config.Item("ReportPath"),False)

	Else
			Call InsertIntoHTMLReport( "endOfTestStep","Select Window",oObject , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO =  8    'raise a user-defined error
			Err.Description = oObject & " does not exist"
    End If
End Function


'--------------------------------DoubleClickAt-----------------------------
Public Function cDoubleClickAt (oObject, identifire, coordinates)
	On Error Resume Next
	
    SendObject oObject, identifire
	startCommand "DoubleClickAt Command", oObject
    If CommandObj.Exist(retryTime)Then
		If CommandObj.GetROProperty("disabled") = 0 Then
			
			x = Split(coordinates, ",")(0)
			y = Split(coordinates, ",")(1)
			Setting.WebPackage("ReplayType") = 2
			CommandObj.FireEvent "ondblclick",x,y,micLeftBtn
			Setting.WebPackage("ReplayType") = 1
			
			If Err.Number<>0 Then
				Call InsertIntoHTMLReport( "endOfTestStep","Double Click At","Element:-"&oObject&" Error Description:-"&Err.Description , "", False, Config.Item("ReportPath"),True)
				endCommand "Fail"
				ErrorNO =  8    'raise a user-defined error
			Else
				Call InsertIntoHTMLReport( "endOfTestStep","Double Click At","Double Clicked at "&x&","&y&" on "&oObject , "", True, Config.Item("ReportPath"),False)
				endCommand "Pass"
			End If
		Else
			Call InsertIntoHTMLReport( "endOfTestStep","Double Click At",oObject , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO =  8    'raise a user-defined error
			Err.Description = oObject & " disabled"
		End If
	Else
		Err.Description = oObject & " does not exist"
        Call InsertIntoHTMLReport( "endOfTestStep","Double Click At","Element :-"&oObject&" "&Err.Description, "", False, Config.Item("ReportPath"),True)
		endCommand "Fail"
		ErrorNO =  8    'raise a user-defined error
        
     End If		
End Function
'---------------------------------------------------------------

Public Function cDoubleClick (oObject, identifire)
	On Error Resume Next
	
	SendObject oObject, identifire
	startCommand "DoubleClick Command", oObject
    If CommandObj.Exist(retryTime)Then
	If True=Cbool(Config.Item("HighLight")) Then
    	CommandObj.Highlight
    End If
		If CommandObj.GetROProperty("disabled") = 0 Then
			'Setting.WebPackage("RepalyType")=2
			'CommandObj.FireEvent "ondblclick"
			'Setting.WebPackage("RepalyType")=1
			CommandObj.Drag		
			CommandObj.Drop		
			CommandObj.Drag		
			CommandObj.Drop
			wait 3
			
			If Err.Number<>0 Then
				Call InsertIntoHTMLReport( "endOfTestStep","Double Click","Element:-"&oObject&" Error Description:-"&Err.Description , "", False, Config.Item("ReportPath"),True)
				endCommand "Fail"
				ErrorNO =  8   'raise a user-defined error
			Else
				Call InsertIntoHTMLReport( "endOfTestStep","Double Click","Double clicked on Element "&oObject , "", True, Config.Item("ReportPath"),False)
				endCommand "Pass"
			End If
		Else
			Err.Description = "Element "&oObject & " is disabled"
			Call InsertIntoHTMLReport( "endOfTestStep","Double Click",Err.Description , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO =  8    'raise a user-defined error
			
		End If
	Else
		Err.Description = "Element "&oObject & " is does not exist"
        Call InsertIntoHTMLReport( "endOfTestStep","Double Click",Err.Description , "", False, Config.Item("ReportPath"),True)
		endCommand "Fail"
		ErrorNO =  8    'raise a user-defined error
        
     End If
     Err.Clear
End Function

'  	wait command
Public Function cPause (WaitTime)
	On Error Resume Next
	startCommand "Pause Command", oObject
    waitTimeInMilSeconds = CInt(WaitTime)
	waitTimeInSeconds = waitTimeInMilSeconds/1000
    wait (waitTimeInSeconds)
	Call InsertIntoHTMLReport( "endOfTestStep","Pause","Duration:- "&WaitTime&"MS" , "", True, Config.Item("ReportPath"),False)
	endCommand "Pass"
End Function

Function cKeyPress(oObject,identifire,KeyBoardInput)
   On Error Resume Next
   SendObject oObject, identifire
   startCommand "Key Press Command", oObject
   InputArr=Split(KeyBoardInput,"|")
  
    If CommandObj.Exist(retryTime)Then
	If True=Cbool(Config.Item("HighLight")) Then
    	CommandObj.Highlight
    End If
		If CommandObj.GetROProperty("disabled") = 0 Then
			CommandObj.Click
			For x = 0 To Ubound(InputArr) Step 1
				cEscapeAndFire (InputArr(x))	
			Next
			If Err.Number<>0 Then
				Call InsertIntoHTMLReport( "endOfTestStep","Key Press",Err.Description, "", False, Config.Item("ReportPath"),True)
				endCommand "Fail"
				ErrorNO =  8    'raise a user-defined error
				
			Else
				Call InsertIntoHTMLReport( "endOfTestStep","Key Press","Element ["&oObject&"] :: Key Press Values ["&KeyBoardInput&"]" , "", True, Config.Item("ReportPath"),False)
				endCommand "Pass"
				End If
		Else
			Err.Description = "Element ["&oObject & "] is disabled"
			Call InsertIntoHTMLReport( "endOfTestStep","Key Press","Element ["&oObject&"]"&Err.Description, "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO =  8    'raise a user-defined error
			
		End If
	Else
		Err.Description = "Element ["&oObject & "] is does not exist"
        Call InsertIntoHTMLReport( "endOfTestStep","Key Press","Element ["&oObject&"]"&Err.Description, "", False, Config.Item("ReportPath"),True)
		endCommand "Fail"
		ErrorNO =  8    'raise a user-defined error
        
     End If

End Function



Function cEscapeAndFire (KeyTag)
Set oShell=Createobject("Wscript.shell")


    If KeyTag = "\n" Then
        oShell.SendKeys "{ENTER}"
    ElseIf KeyTag = "\t" Then
        oShell.SendKeys "{TAB}"
    ElseIf KeyTag = "down" Then
        oShell.SendKeys "{DOWN}"
    ElseIf KeyTag = "F3" Then
        oShell.SendKeys "{F3}"
    ElseIf KeyTag = "backspace" Then
    	oShell.SendKeys "{BACKSPACE}"
    ElseIf KeyTag = "break" Then
    	oShell.SendKeys "{BREAK}"
    ElseIf KeyTag = "capslock" Then
    	oShell.SendKeys "{CAPSLOCK}"
    ElseIf KeyTag = "delete" Then
    	oShell.SendKeys "{DEL}"
    ElseIf KeyTag = "end" Then
    	oShell.SendKeys "{END}"
    ElseIf KeyTag = "esc" Then
    	oShell.SendKeys "{ESC}"
    ElseIf KeyTag = "insert" Then
    	oShell.SendKeys "{INSERT}"
    ElseIf KeyTag = "left" Then
    	oShell.SendKeys "{LEFT}"
    ElseIf KeyTag = "right" Then
    	oShell.SendKeys "{RIGHT}"
    ElseIf KeyTag = "up" Then
    	oShell.SendKeys "{UP}"
    ElseIf KeyTag = "numlock" Then
    	oShell.SendKeys "{NUMLOCK}"
    ElseIf KeyTag = "pagedown" Then
    	oShell.SendKeys "{PGDN}"
    ElseIf KeyTag = "pageup" Then
    	oShell.SendKeys "{PGUP}"
    ElseIf KeyTag = "printscreen" Then
    	oShell.SendKeys "{PRTSC}"
    ElseIf KeyTag = "scrolllock" Then
    	oShell.SendKeys "{SCROLLLOCK}"
    ElseIf KeyTag = "win" Then
    	oShell.SendKeys "^{ESC}"
    ElseIf KeyTag = "space" Then
    	oShell.SendKeys "” “"
    ElseIf KeyTag = "ctrl+F4" Then
    	oShell.SendKeys "^{f4}"
    ElseIf KeyTag = "space" Then
    	oShell.SendKeys "%f{F4}"
    Else 
	KeyTag=Replace(KeyTag,"ctrl+","^")
	KeyTag=Replace(KeyTag,"alt+","%")
	KeyTag=Replace(KeyTag,"ctrl","^")
	KeyTag=Replace(KeyTag,"alt","%")
        oShell.SendKeys KeyTag

    End If
End Function



Function fireKeyEvent(var1)
	a=Split(var1,"|")
	for each x in a
		b=Split(x,"=")
		If b(0) = "key" Then
			cEscapeAndFire b(1)
		ElseIf b(0) = "type" Then
        	cEscapeAndFire b(1)
		ElseIf b(0) = "wait" Then
			wait b(1)
		End If
	next
End Function



Public Function cFireEvent(var2,waitTime)
On Error Resume Next
	startCommand "Fire Event  Command", " "
	FireEventarr=Split(var2,"|")
	For x = 0 To Ubound(FireEventarr) Step 1
		a=Split(FireEventarr(x),"%")
		If a(0)="KEY" Then
			fireKeyEvent(a(1))
			Call InsertIntoHTMLReport( "endOfTestStep","Fire Event","Fire Event ["&a(1)&"]" , "", True, Config.Item("ReportPath"),False)
			endCommand "Pass"
	Else
		Call InsertIntoHTMLReport( "endOfTestStep","Fire Event","Invalid inputs for Fire Event" , "", False, Config.Item("ReportPath"),True)
		endCommand "Fail"
		ErrorNO =  8    'raise a user-defined error
        Err.Description = "Invalid inputs for Fire Event"
	End If
	
	'at last event wait for the wait time
	If (x=Ubound(FireEventarr)) Then
		wait (Cint(waitTime)/1000)
	End If

	Next
	
		
End Function

Public Function cSetVarProperty (oObject, identifire, sProperty)
	On Error Resume Next
	SendObject oObject, identifire
	startCommand "SetVarProperty Command", oObject
    If CommandObj.Exist(retryTime)Then
	If True=Cbool(Config.Item("HighLight")) Then
    	CommandObj.Highlight
    End If
	
		if isEmpty(CommandObj.GetROProperty(sProperty)) then 
			Call InsertIntoHTMLReport( "endOfTestStep","Set Var Property","Element:- "&oObject&", Property "&sProperty&" is not found" , "", False, Config.Item("ReportPath"),False)
			ErrorNO =  8  'raise a user-defined error
			endCommand "Fail"
		else 
			If sProperty = "textContent"Then
				cSetVarProperty=CommandObj.GetROProperty("text")
				Call InsertIntoHTMLReport( "endOfTestStep","Set Var Property","Element:- "&oObject&" Property value:-"&cSetVarProperty , "", True, Config.Item("ReportPath"),False)
				endCommand "Pass"
			Else
				cSetVarProperty=CommandObj.GetROProperty(sProperty)
				Call InsertIntoHTMLReport( "endOfTestStep","Set Var Property","Element:- "&oObject&" Property value:-"&cSetVarProperty , "", True, Config.Item("ReportPath"),False)
				endCommand "Pass"
			End If
		end if 
	Else
		Err.Description = oObject & " does not exist"
        Call InsertIntoHTMLReport( "endOfTestStep","Set Var Property",Err.Description , "", False, Config.Item("ReportPath"),False)
		endCommand "Fail"
		ErrorNO =  8    'raise a user-defined error
        
     End If
End Function




Public Function cSetTable (oObject, identifire,Action,rownumber,columnnumber,data)
	On Error Resume Next
	columnnumber = "#"& columnnumber
	SendObject oObject, identifire
	startCommand "Dable grid  Command", oObject
    If CommandObj.Exist(retryTime)Then
	If True=Cbool(Config.Item("HighLight")) Then
    	CommandObj.Highlight
    End If
		If   Action = "gridType" Then
				CommandObj.SetCellData rownumber,columnnumber,data
				Call InsertIntoHTMLReport( "endOfTestStep","Type",oObject , "", True, Config.Item("ReportPath"),False)
				endCommand "Pass"
					
	
		elseif  Action = "gridSelect" Then
				CommandObj.SelectCell rownumber,columnnumber
				Call InsertIntoHTMLReport( "endOfTestStep","Type",oObject , "", True, Config.Item("ReportPath"),False)
				endCommand "Pass"
		End If

	Else
			Call InsertIntoHTMLReport( "endOfTestStep","Type",oObject , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO =  8    'raise a user-defined error
			Err.Description = oObject & " does not exist"

	End If
End Function

'------------------------CreateDBConnection------------------------
Function cCreateDBConnection(databaseType,instanceName,url,username,password)
   On Error Resume Next
	startCommand "cCreateDBConnection", ""
	Dim isNewInstance
	If  instances.Count=0 Then
		isNewInstance=true
	ELSE
	    a = instances.Keys
		For i = 0 To instances.Count -1 
			s = s & a(i)
			s=split(s,"*")
			If s(1)=instanceName Then
				isNewInstance=false
			Else
				isNewInstance=true
			End If
		Next
	End If
	If isNewInstance Then
		If databaseType="mysql" Then
			mysql instanceName,url,username,password
		ElseIf databaseType="oracle " Then
			oracle url,username,password
		ElseIf databaseType="mssql " Then
			mssql url,username,password
		Else
			Call InsertIntoHTMLReport( "endOfTestStep","Create DB Connection","" , "", True, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO = 8   
			Err.Description = oObject & " does not exist"
		End If
	End If
End Function

Function mysql(instanceName,url,username,password)
	urlArray = split(url,"/")
	ServerArray=urlArray(2)
	Server = split(ServerArray,":")
	strConnection = "DRIVER={MySQL ODBC 5.1 Driver}; Server=" & Server(0) & "; Database="&urlArray(3)& ";Uid=" & username & ";Pwd=" & password & ";"
	Set conn = CreateObject("ADODB.Connection")
	Set Recordset = CreateObject("ADODB.Recordset")
	conn.Open strConnection
	If conn.State = 1 Then	
		conn.Close
		Set conn= nothing
	Else
		Call InsertIntoHTMLReport( "endOfTestStep","Create DB Connection","Unable to Connect Server " , "", False, Config.Item("ReportPath"),True)
		endCommand "Fail"
	End If
	instances.ADD strConnection&"*"&instanceName,conn
	Call InsertIntoHTMLReport( "endOfTestStep","Create DB Connection","Connection Created Successfully" , "", True, Config.Item("ReportPath"),False)
	endCommand "Pass"
End Function

Function oracle( instanceName,strDSNname,strDBuserName,strDBPass,strDBname)
	Dim con,rs
	Set con=createobject("adodb.connection")
	Set rs=createobject("adodb.recordset")
	con.open "Driver={Microsoft ODBC for Oracle};Server=" & url & "; Uid=" & username & ";Pwd=" & password & ";"
End Function

Function mssql( instanceName,strDSNname,strDBuserName,strDBPass,strDBname)
	Dim con,rs
	Set con=createobject("adodb.connection")
	Set rs=createobject("adodb.recordset")
	con.open"Driver={SQL Server};server=" & url & ";uid=" & username & ";pwd=" & password & ";database=pubs"
End Function


Function cgetDBTable(instanceName, query)  '  Implement only for mySql yet
	On Error Resume Next
	Dim connection
	startCommand "GetDBTable Command", ""
	Set dataList=DotnetFactory.CreateInstance("System.Collections.ArrayList")
	a = instances.Keys
	For i = 0 To instances.Count -1 
		s = s & a(i)
		s1=split(s,"*")
		If instanceName=s1(1) Then
			'==============================
			Set connection = CreateObject("ADODB.Connection")
			Set Recordset = CreateObject("ADODB.Recordset")
			connection.Open s1(0)
			'================================
		End If
	Next
	Set rs = connection.Execute(query)

	Do While not rs.EOF
		For i=0 to rs.fields.count-1
			dataList.Add(Cstr(rs.fields(i).value))
		Next 
		rs.MoveNext
	Loop
	Set cgetDBTable=dataList
End Function

'----------------CheckDBResults Command-------------------
Function cCheckDBResults(instanceName, query,expectedValue,stopOnFaliure, customErrorMessage)
	On Error Resume Next
	startCommand "cCheckDBResults", ""
	Dim objArrList
	Set inputTable=CreateObject("System.Collections.ArrayList")
	Set inputValuList=CreateObject("System.Collections.ArrayList")
	Set ary=CreateObject("System.Collections.ArrayList")
	Set objArrList=cgetDBTable (instanceName,query)

	Dim temp
	'temp = Replace(expectedValue,"\\,","\\*")
	inputValuList=Split(expectedValue,",")
	ary = Split (expectedValue,",")
	For Each x In ary
		temp = x
		' temp = Replace(x,"\\*",",")
		inputTable.Add(temp)
	Next

	Dim x
	Dim isFound 'As Boolean

	If UBound(inputTable) + 1 < cInt(objArrList.Count) Then
		x = cInt(objArrList.Count)
		'MsgBox "Equal number of Expected Values and DataBase Values"
	Else
		'MsgBox "Not equal number of Expected Values and DataBase Values"
		x =  UBound(inputTable)+1
	End If

	For x = 0 To x - 1
		If  (objArrList.Item(cInt(x)) = null) OR (inputTable(x)) = null Then   
			isFound = False
			Exit For
		End If
		If  inputTable(x) = objArrList.Item(cInt(x)) Then
			'MsgBox ("Input expected value '"&inputTable(X))&" 'and Actual Database value is '"& objArrList.Item(cInt(X))
			isFound = true
		Else
			isFound = False
			Exit For
		End If
	Next

	If 	isFound = true Then
		Call InsertIntoHTMLReport( "endOfTestStep","Check DB Results","Check DB Result Equal to expected Values" , "", True, Config.Item("ReportPath"),False)
		endCommand "Pass"	
	Else
		Dim inputTableAllString
		inputTableAllString=joinArrayList(inputTable)
		Call InsertIntoHTMLReport( "endOfTestStep","Check DB Results","inputTableAllString" & cCustomErrorGeneration(customErrorMessage), "", False, Config.Item("ReportPath"),True)
		endCommand "Fail"
		ErrorNO = 8    'raise a user-defined error
		Err.Description = oObject & " does not exist"
	End If
End Function


Function joinArrayList (aList )
	Dim var , tempVar 
	For Each  x In aList
		tempVar = cStr(x) + " "
		var = var + tempVar 
	Next
	Dim concat , a
	concat = Replace(var," "," | ")
	joinArrayList = cStr(concat)
End Function

'---------------------------GetDBResult (String, Int, Boolean)-----------------------
Function cGetStringDBResult(instanceName, query)
	On Error Resume Next
	startCommand "GetStringDBResult", ""
	Dim objArrList
	Dim val
	Set objArrList=cgetDBTable (instanceName,query)
	If objArrList.Count >= 2 Then
		Set val=CStr(objArrList.Item(CInt(0)))
		cGetStringDBResult = val
		Call InsertIntoHTMLReport( "endOfTestStep","Get String DB Result","For Query = " &  query & " Actual result contains more than one value. Return Value :- " & val, "", True, Config.Item("ReportPath"),False)
	    endCommand "Pass"	
	ElseIf objArrList.Count = 1 Then
		'Set val=CStr(objArrList.Item(CInt(0)))
		cGetStringDBResult = CStr(objArrList.Item(CInt(0)))
		Call InsertIntoHTMLReport( "endOfTestStep","Get String DB Result","For Query = " &  query & " Return Value :- " & CStr(objArrList.Item(CInt(0))), "", True, Config.Item("ReportPath"),False)
	    endCommand "Pass"
	Else
	    Call InsertIntoHTMLReport( "endOfTestStep","Get String DB Result","No Value Returned" , "", False, Config.Item("ReportPath"),True)
	    cGetStringDBResult = null
		endCommand "Fail"
		ErrorNO = 8    'raise a user-defined error
		Err.Description = "No Value Returned"	
	End If  
End Function

Function cGetIntDBResult(instanceName, query)
	On Error Resume Next
	startCommand "GetIntDBResult", ""
	Dim objArrList
	Dim val
	Set objArrList=cgetDBTable (instanceName,query)
	If objArrList.Count >= 2 Then
		If NOT(IsNumeric (objArrList.Item(CInt(0)))) Then
			Call InsertIntoHTMLReport( "endOfTestStep","Get Int DB Result","The value trying to retrive ("&CStr(objArrList.Item(CInt(0)))&" ) is not stored as an interger in the database." , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO = 8    'raise a user-defined error
			Err.Description = "The value trying to retrive ("&CStr(objArrList.Item(CInt(0)))&" ) is not stored as an interger in the database."
		Else
			val=CInt(objArrList.Item(CInt(0)))
			Call InsertIntoHTMLReport( "endOfTestStep","Get Int DB Result","For Query = " &  query & " Actual result contains more than one value. Return Value :- " & val , "", True, Config.Item("ReportPath"),False)
			endCommand "Pass"	
			cGetIntDBResult = val 
		End If
	ElseIf objArrList.Count = 1 Then
		If NOT(IsNumeric (objArrList.Item(CInt(0)))) Then
			Call InsertIntoHTMLReport( "endOfTestStep","Get Int DB Result","The value trying to retrive ("&CStr(objArrList.Item(CInt(0)))&" ) is not stored as an interger in the database." , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO = 8    'raise a user-defined error
			Err.Description = "The value trying to retrive ("&CStr(objArrList.Item(CInt(0)))&" ) is not stored as an interger in the database."
		Else
			val=CInt(objArrList.Item(CInt(0)))
			Call InsertIntoHTMLReport( "endOfTestStep","Get Int DB Result","For Query = " &  query & " Return Value :- " & val , "", True, Config.Item("ReportPath"),False)
			endCommand "Pass"	
			cGetIntDBResult = val 
		End If
	Else
		Call InsertIntoHTMLReport( "endOfTestStep","Get Int DB Result","No Value Returned" , "", False, Config.Item("ReportPath"),True)
	    cGetStringDBResult = null
		endCommand "Fail"
		ErrorNO = 8    'raise a user-defined error
		Err.Description = "No Value Returned"
	End If
End Function

Function cGetBooleanDBResult(instanceName, query)
	On Error Resume Next
	startCommand "GetBooleanDBResult", ""
	Dim objArrList
	Dim val
	Set objArrList=cgetDBTable (instanceName,query)
	If objArrList.Count >= 2 Then
		If vartype(objArrList.Item(CInt(0))) <> 11 Then
			Call InsertIntoHTMLReport( "endOfTestStep","Get Boolean DB Result","The value trying to retrive ("&CStr(objArrList.Item(CInt(0)))&" ) is not stored as an boolean in the database." , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO = 8    'raise a user-defined error
			Err.Description = "The value trying to retrive ("&CStr(objArrList.Item(CInt(0)))&" ) is not stored as an boolean in the database."
		Else
			val=CBool(objArrList.Item(CInt(0)))
			Call InsertIntoHTMLReport( "endOfTestStep","Get Boolean DB Result","For Query = " &  query & " Actual result contains more than one value. Return Value :- " & val , "", True, Config.Item("ReportPath"),False)
			endCommand "Pass"	
			cGetBooleanDBResult = val
		End If
	ElseIf objArrList.Count = 1 Then
		If vartype(objArrList.Item(CInt(0))) <> 11 Then
			Call InsertIntoHTMLReport( "endOfTestStep","Get Boolean DB Result","The value trying to retrive ("&CStr(objArrList.Item(CInt(0)))&" ) is not stored as an boolean in the database." , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO = 8    'raise a user-defined error
			Err.Description = "The value trying to retrive ("&CStr(objArrList.Item(CInt(0)))&" ) is not stored as an boolean in the database."
		Else
			val=CBool(objArrList.Item(CInt(0)))
			Call InsertIntoHTMLReport( "endOfTestStep","Get Boolean DB Result","For Query = " &  query & " Return Value :- " & val , "", True, Config.Item("ReportPath"),False)
			endCommand "Pass"	
			cGetBooleanDBResult = val
		End If
	Else
		Call InsertIntoHTMLReport( "endOfTestStep","Get Boolean DB Result","No Value Returned" , "", False, Config.Item("ReportPath"),True)
	    cGetStringDBResult = null
		endCommand "Fail"
		ErrorNO = 8    'raise a user-defined error
		Err.Description = "No Value Returned"
	End If
End Function

'--------------------------------------------------------------------------------------



Public Function cNavigateToUrl (Url, identifire,waittime)
	On Error Resume Next
	SendObject oObject, identifire
	startCommand "cNavigateToUrl  Command", oObject
	Browser("title:=.*").OpenNewTab
	Browser("title:=.*").Navigate Url
    'Browser("title:=.*").Sync
	wait (Cint(WaitTime)/1000)
	Call InsertIntoHTMLReport( "endOfTestStep","Navigate To Url","URL:-"&Url , "", True, Config.Item("ReportPath"),False)
	endCommand "Pass"
End Function

Function cCheckTextPresent(oObject, stopOnFailure, customError)
   	On Error Resume Next
	startCommand "Check Text Present", ""
	If CommandObj.Exist(retryTime)Then
	If True=Cbool(Config.Item("HighLight")) Then
    	CommandObj.Highlight
    End If
    Set oAll = Browser("x").Object.Document.All
   
   bFound = False
   For each oItem in oAll
   If oItem.OuterText = oObject.getROProperty("name") Then
      bFound = True
      Exit For
    End If
   Next

    If bFound = True Then 
	Call InsertIntoHTMLReport( "endOfTestStep","Check Text Present","" , "", True, Config.Item("ReportPath"),False)
	endCommand "Pass"
	ELSE
	       	Call InsertIntoHTMLReport( "endOfTestStep", "Check Text Present", cCustomErrorGeneration(customError) & oObject , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO = 8    'raise a user-defined error
			Err.Description = oObject & " does not exist"
	 END IF

	Else
			Call InsertIntoHTMLReport( "endOfTestStep", "Check Text Present", cCustomErrorGeneration(customError) & oObject , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO = 8    'raise a user-defined error
			Err.Description = oObject & " does not exist"
	End If
End Function

Function cGoBack (byVal WaitTime)
	On Error Resume Next
	startCommand "GoBackCommand", ""
    Browser("title:=.*").Page("title:=.*").Back
	wait (WaitTime/1000)
	Call InsertIntoHTMLReport( "endOfTestStep","GoBack","Navigate Page Back" , "", True, Config.Item("ReportPath"),False)
	endCommand "Pass"
End Function

'Check Document command
Public Function cCheckDocument (docType, filePath, pageNumberRange, verifyType, inputString, stopOnFailure, customError)
	On Error Resume Next
	startCommand "CheckDocument Command", ""
	Const ForReading = 1
	Dim arrFileLines()
	i = 0
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(filePath) Then
		If docType = "txt" Or docType = "das" Or docType = "dat" Or docType = "csv" Or docType = "ecship" Then
			
			Set objFSO = CreateObject("Scripting.FileSystemObject")
			Set objFile = objFSO.OpenTextFile(filePath, ForReading)
			
			Do Until objFile.AtEndOfStream
				ReDim Preserve arrFileLines(i)
				arrFileLines(i) = objFile.ReadLine
				i = i + 1
			Loop
			Set objFile = Nothing
			
			For Each strLine In arrFileLines
				If InStr(strLine, inputString) > 0 Then
					Call InsertIntoHTMLReport( "endOfTestStep","Check Document (" & docType & ")","Input string is available in Document", "", True, Config.Item("ReportPath"),False)
					endCommand "Pass"
					Exit Function	
				End If
			Next
			
			Call InsertIntoHTMLReport( "endOfTestStep", "Check Document (" & docType & ")", cCustomErrorGeneration(customError) & "<br />" & "KKKInput string is not available in Document" ,"", False, Config.Item("ReportPath"),True)
			ErrorNO = 8  'raise a user-defined error
			endCommand "Fail"
			Exit Function		
		
	ElseIf LCase(docType) = "excel" Then
			
			checkExcelDocument docType, filePath, pageNumberRange, verifyType, inputString, stopOnFailure, customError
	
	ElseIf Lcase(docType)="pdf" Then
						
		strPath = Environment.Value("TestDir")
			

		
		Dim acroApp, acroAVDoc, acroPDDoc, acroRect, PDTextSelect
		Dim gPDFPath, nElem, pageNo
		pageNo = CInt(pageNumberRange)-1 'first page on a PDF file
			 
	   gPDFPath = filePath

		' ** Initialize Acrobat by creating App object
		Set acroApp = CreateObject ("AcroExch.App")
		' ** show acrobatacroApp.Show()' ** Set AVDoc object
		Set acroAVDoc = CreateObject("AcroExch.AVDoc")' ** open the PDF
		If acroAVDoc.Open( gPDFPath, "Accessing PDF's") Then
			 If acroAVDoc.IsValid = False Then ExitTest()
			  acroAVDoc.BringToFront()
			 
			  Call acroAVDoc.Maximize(True)
			  Print"Current pdf title ---> "& acroAVDoc.GetTitle()
			  Set acroPDDoc = acroAVDoc.GetPDDoc()
			  Print"File Name ---> "& acroPDDoc.GetFileName()
			  Print"Number of Pages ---> "& acroPDDoc.GetNumPages()
			  
			  Set acroRect = CreateObject("AcroExch.Rect")
			  acroRect.Top = 1500
			  acroRect.Left = 10
			  acroRect.Bottom = 10
					acroRect.Right = 1000
			  ' ** Selecting page 42 ( index is 43)
			  Set PDTextSelect = acroPDDoc.CreateTextSelect( pageNo, acroRect )
			  If PDTextSelect Is Nothing Then
			   Print"Unable to Create TextSelect object."
				  ExitTest()
			  End If
			 
			  Call acroAVDoc.SetTextSelection( PDTextSelect )
			  Call acroAVDoc.ShowTextSelect()
			  Print"Selection Page Number ---> " & PDTextSelect.GetPage()
			  Print"Selection Text Elements ---> "& PDTextSelect.GetNumText()
			  ' ** Looping through text elements
			  

			Dim fullString
			Dim strSearchFor
			  
			  For nElem = 0 To PDTextSelect.GetNumText() - 1
			   'Print"Text # "& nElem &" ---> '"& PDTextSelect.GetText( nElem ) &"'"
			   fullString = fullString & PDTextSelect.GetText( nElem )
			  Next
			  
			  Dim a
			 a = replace (fullString,vbCrLf," ")
			  
			  'Print a
			  strSearchFor = inputString
			  If InStr(1, a, strSearchFor) > 0 then
			  Print "We found strSearchFor in strSearchString"
			  Call InsertIntoHTMLReport( "endOfTestStep","Check Document (" & docType & ")","Input string is available in Document", "", True, Config.Item("ReportPath"),False)
								endCommand "Pass"
								Exit Function	
			Else
			  Print "We didn't find strSearchFor in strSearchString"
			  Call InsertIntoHTMLReport( "endOfTestStep", "Check Document (" & docType & ")", cCustomErrorGeneration(customError) & " Input string is not available in Document", "", False, Config.Item("ReportPath"),True)
						ErrorNO = 8  'raise a user-defined error
						endCommand "Fail"
						Exit Function
			End If 
			  
			  
			  '  ** Destroying Text Selection
			  Call PDTextSelect.Destroy()
			 End If
			 AcroApp.CloseAllDocs()
			 AcroApp.Exit()
			Set PDTextSelect = Nothing : Set acroRect = Nothing
			Set AcroApp =  Nothing: Set AcroAVDoc =  Nothing

		ElseIf docType = "xml" Then			
			xpath = Split(inputString, "|")(0)		
			expectedValue = Split(inputString, "|")(1)
			Set xmlDoc1 = CreateObject("Microsoft.XMLDOM")
		
			xmlDoc1.Async = "False"
			xmlDoc1.Load(filePath)
			Set objNode = xmlDoc1.SelectSingleNode(xpath)
			If IsNull(objNode.text) Then
				Call InsertIntoHTMLReport( "endOfTestStep", "Check Document (" & docType & ")", cCustomErrorGeneration(customError) & "<br />" & "Invalid XPath", "", False, Config.Item("ReportPath"),True)
				ErrorNO = 8  'raise a user-defined error
				endCommand "Fail"
				Exit Function
			End If
			val = objNode.text
			If val = expectedValue Then
				Call InsertIntoHTMLReport( "endOfTestStep","Check Document (" & docType & ")","Input string is available in Document", "", True, Config.Item("ReportPath"),False)
				endCommand "Pass"
				Exit Function
			Else
				Call InsertIntoHTMLReport( "endOfTestStep", "Check Document (" & docType & ")", cCustomErrorGeneration(customError) & "<br />" & "KKKInput string is not available in Document", "", False, Config.Item("ReportPath"),True)
				ErrorNO = 8  'raise a user-defined error
				endCommand "Fail"
				Exit Function
			End If
			
			Set xmlDoc1 = Nothing
		End If
	Else
		Call InsertIntoHTMLReport( "endOfTestStep", "Check Document (" & docType & ")", cCustomErrorGeneration(customError) & "<br />" & "File not exists. Check File path", "", False, Config.Item("ReportPath"),True)
		ErrorNO = 8  'raise a user-defined error
		endCommand "Fail"
		Exit Function
	End If
	
End Function

'Store Command
Public Function cStore(key, typeOfVar, value)
startCommand "Store Command", key
If Store.Exists(key) Then
	
	store.Remove(key)
	Store.Add key,cstr(value)	
	If Err.Number<>0 Then
		Call InsertIntoHTMLReport( "endOfTestStep","Store",Err.Description ,"", False, Config.Item("ReportPath"),True)
		ErrorNO = 8  'raise a user-defined error
		endCommand "Fail"
		Err.Clear	
	Else
		Call InsertIntoHTMLReport( "endOfTestStep","Store","Key ["&key&"] is overwritten by value ["&value&"]", "", True, Config.Item("ReportPath"),False)
		endCommand "Pass"				
	End If	
Else

 Store.Add key,cstr(value)
 	If Err.Number<>0 Then	
		Call InsertIntoHTMLReport( "endOfTestStep","Store",Err.Description ,"", False, Config.Item("ReportPath"),True)
		ErrorNO = 8  'raise a user-defined error
		endCommand "Fail"
    	Err.Clear		
	Else
		Call InsertIntoHTMLReport( "endOfTestStep","Store","Key ["&key&"] is stored with value ["&value&"]", "", True, Config.Item("ReportPath"),False)
		endCommand "Pass"			
	End If 
End if

End Function



'Retrieve Command
Public Function cRetrieve (key, cVariable, typeOfVar)
startCommand "Retrieve Command", key
On Error Resume Next

If Store.Exists(key) Then

	IF typeOfVar="String" Then
		cVariable=Cstr(Store.Item(key))
	ElseIf typeOfVar="Int" Then
		cVariable=Cint(Store.Item(key))
	ElseIF typeOfVar="Boolean" Then
		cVariable=cbool(Store.Item(key))
	ElseIF typeOfVar="Double" Then
		cVariable=cdbl(Store.Item(key))
	End IF
		
		If Err.Number<>0 Then
			Call InsertIntoHTMLReport( "endOfTestStep","Retrieve",Err.Description ,"", False, Config.Item("ReportPath"),True)
			ErrorNO = 8  'raise a user-defined error
			endCommand "Fail"
		Else
			Call InsertIntoHTMLReport( "endOfTestStep","Retrieve","Key ["&key&"] returned value ["&Cstr(cVariable)&"]", "", True, Config.Item("ReportPath"),False)
			endCommand "Pass"
		
		End IF

	

Else
	    Err.Description="Invalid Key ["&key&"]"
		Call InsertIntoHTMLReport( "endOfTestStep","Retrieve",Err.Description ,"", False, Config.Item("ReportPath"),True)
		ErrorNO = 8  'raise a user-defined error
		endCommand "Fail"
		Err.Clear
	
End If
End Function



'@author Vimukthi Hewapathirana
Function cGetObjectCount(oObject,identifire)
On Error Resume Next
SendObject oObject, identifire
startCommand "GetObjectCount  Command", CommandObj.ToString

   
    cGetObjectCount = -1
    If CommandObj.Exist(0) Then
    
    	 If CommandObj.GetRoProperty("micclass")="WebTable" Then
   	  		cGetObjectCount = CommandObj.RowCount
   	  		Call InsertIntoHTMLReport( "endOfTestStep","Get Object Count","Element ["&oObject&"] :: Table Row Count ["&cGetObjectCount&"]","", True, Config.Item("ReportPath"),False)
			endCommand "Pass"		
			ErrorNO = 0
   	  		Exit Function
   		Else

	        cGetObjectCount = 1 
	        Call InsertIntoHTMLReport( "endOfTestStep","Get Object Count","Element ["&oObject&"] :: Count =[1]","", True, Config.Item("ReportPath"),False)
			endCommand "Pass"		
			ErrorNO = 0
        Exit Function
        End iF
   
   End If
 
    Dim Parent, ClassName, TOProperties, oDesc, i
    Set Parent = CommandObj.GetTOProperty("parent")
    ClassName = CommandObj.GetTOProperty("micclass")
    Set TOProperties = CommandObj.GetTOProperties()
    Set oDesc = Description.Create
 
    For i = 0 To TOProperties.Count - 1
        oDesc.Add TOProperties(i).Name, TOProperties(i).Value
    Next
    
    cGetObjectCount = Parent.ChildObjects(oDesc).Count
    
    If cGetObjectCount<1 Then
	    Call InsertIntoHTMLReport( "endOfTestStep","Get Object Count","Element ["&oObject&"] is not Present :: Count [0]","", True, Config.Item("ReportPath"),False)
		endCommand "Pass"
		ErrorNO = 0 
    Else
        Call InsertIntoHTMLReport( "endOfTestStep","Get Object Count","Element ["&oObject&"] :: Count ["&cGetObjectCount&"]" ,"", True, Config.Item("ReportPath"),False)
		endCommand "Pass"
		ErrorNO = 0 
    End IF


    
End Function


'Supportive Function
'Genarates Custom Error Message 
'@Pram msg - String
'Last updated on 22.04.2015
Function cCustomErrorGeneration(message)
	If message  = "" Then
		cCustomErrorGeneration=""
	Else
		cCustomErrorGeneration = "Error Message: " & message&" :: "
	End If
End Function



Function generateData(dataType, dataLength)
	
	Dim str
	
	If StrComp(dataType, "int", vbTextCompare) = 0 Then
		str = "1234567890"	
	ElseIf StrComp(dataType, "string", vbTextCompare) = 0 Then
		str = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
	ElseIf StrComp(dataType, "alphanumeric", vbTextCompare) = 0 Then
		str = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"
	ElseIf Instr(dataType, "date") > 0 Then
		generateData = generateRandomDate(dataType, dataLength)
		Exit Function
	Else
		Exit Function
	End If
	
	generateData = generateRandomValue(str, dataLength)
	
End Function

Function generateRandomValue(st, dataLength)
	Dim str
   ' For i = 1 to dataLength
    For i = 1 to dataLength
        str = str & Mid(st, RandomNumber(1, Len(st)), 1)
    Next
    generateRandomValue = str
End Function


Function generateRandomDate(st, dataLength)
	dataType = Split(st,"|")(0)
	format = Split(st,"|")(1)
	skipWeekend = CBool(Split(st,"|")(2))
	
	currentDate = Date
	'newDate = dateadd("d",dataLength,currentDate) 
	
	Set currDate = DotNetFactory.CreateInstance("System.DateTime")
    Set oDate = currDate.Parse(currentDate)
    newDate = oDate.ToString("MM/dd/yyyy")
    
    Set MyDate = Nothing

	generateRandomDate = newDate
	
End Function

Public Function checkExcelDocument (docType, filePath, pageNumberRange, verifyType, inputString, stopOnFailure, customError)
On Error Resume Next
If verifyType = "EXISTS" Then
	Set ExcelObject = createobject("excel.application")
			Dim xlsheet
			Dim result
			ExcelObject.Workbooks.Open filePath
			ExcelObject.Application.Visible = false
		
			sheets = Split(pageNumberRange,",")
			For Each sheet In sheets
				Set ExcelSheet = ExcelObject.ActiveWorkbook.Worksheets(sheet)

				Row = ExcelSheet.UsedRange.Row + ExcelSheet.UsedRange.Rows.Count - 1
				Col = ExcelSheet.UsedRange.Column + ExcelSheet.UsedRange.Columns.Count - 1

				For  i= 1 to Row 
					For j=1 to Col
						'MsgBox ExcelSheet.cells(i,j).value
						If ExcelSheet.cells(i,j).value = "" Then
						Else
							If InStr(ExcelSheet.cells(i,j).value, inputString) > 0 Then
							result = "Pass"	
							End If
						End If
					Next
				Next
			Next
		
		'-------------------------------------
		If result = "Pass" Then
		    Call InsertIntoHTMLReport( "endOfTestStep","CheckDocument (" & docType & ")","Input string is available in Document", "", True, Config.Item("ReportPath"),False)
			endCommand "Pass"
		Else
		    If  stopOnFailure = True Then
					Call InsertIntoHTMLReport( "endOfTestStep", "Check Document (" & docType & ")", cCustomErrorGeneration(customError) & " Input string is not available in Document" ,"", False, Config.Item("ReportPath"),True)
					endCommand "Fail"
					ErrorNO =  8    'raise a user-defined error
			

			else 
					Call InsertIntoHTMLReport( "endOfTestStep", "Check Document (" & docType & ")", cCustomErrorGeneration(customError) & " Input string is not available in Document" ,"", False, Config.Item("ReportPath"),True)
					endCommand "Fail"
					ErrorNO = 0   'raise a user-defined error
						'Err.Description = oObject & " disabled"
			End If
		
		End If
		'------------------------------------
			ExcelObject.ActiveWorkbook.Close
			ExcelObject.Application.Quit
			Set ExcelSheet = Nothing
			Set ExcelObject = Nothing
			
ElseIf verifyType = "COLCOUNT" Then

            Set ExcelObject = createobject("excel.application")
			'Dim xlsheet
			ExcelObject.Workbooks.Open filePath
			ExcelObject.Application.Visible = false

				Set ExcelSheet = ExcelObject.ActiveWorkbook.Worksheets(pageNumberRange)

				Row = ExcelSheet.UsedRange.Row + ExcelSheet.UsedRange.Rows.Count - 1
				Col = ExcelSheet.UsedRange.Column + ExcelSheet.UsedRange.Columns.Count - 1
				
				If InStr(Col, inputString) > 0 Then
					Call InsertIntoHTMLReport( "endOfTestStep","CheckDocument (" & docType & ")"," Expected COLCOUNT : "&inputString&" available in Document", "", True, Config.Item("ReportPath"),False)
					endCommand "Pass"	
				Else
					If  stopOnFailure = True Then
			       	 	Call InsertIntoHTMLReport( "endOfTestStep", "Check Document (" & docType & ")", cCustomErrorGeneration(customError) & " Expected COLCOUNT :"&inputString&" |Actual COLCOUNT : "&Col&" in Document" ,"", False, Config.Item("ReportPath"),True)
			       	 	ErrorNO = 8  'raise a user-defined error
			        	endCommand "Fail"
			        
			        else 
						Call InsertIntoHTMLReport( "endOfTestStep", "Check Document (" & docType & ")", cCustomErrorGeneration(customError) & " Expected COLCOUNT :"&inputString&" |Actual COLCOUNT : "&Col&" in Document" ,"", False, Config.Item("ReportPath"),True)
						endCommand "Fail"
						ErrorNO = 0   'raise a user-defined error
						'Err.Description = oObject & " disabled"
					End If
				End If

			ExcelObject.ActiveWorkbook.Close
			ExcelObject.Application.Quit
			Set ExcelSheet = Nothing
			Set ExcelObject = Nothing
			
ElseIf verifyType = "ROWCOUNT" Then

            Set ExcelObject = createobject("excel.application")
			
			ExcelObject.Workbooks.Open filePath
			ExcelObject.Application.Visible = false
		
				Set ExcelSheet = ExcelObject.ActiveWorkbook.Worksheets(pageNumberRange)

				Row = ExcelSheet.UsedRange.Row + ExcelSheet.UsedRange.Rows.Count - 1
				Col = ExcelSheet.UsedRange.Column + ExcelSheet.UsedRange.Columns.Count - 1
			
				
				If InStr(Row, inputString) > 0 Then
					Call InsertIntoHTMLReport( "endOfTestStep","CheckDocument (" & docType & ")"," Expected ROWCOUNT : "&inputString&" available in Document", "", True, Config.Item("ReportPath"),False)
					endCommand "Pass"	
				Else
					If  stopOnFailure = True Then
			       	 	Call InsertIntoHTMLReport( "endOfTestStep", "Check Document (" & docType & ")", cCustomErrorGeneration(customError) & " Expected ROWCOUNT :"&inputString&" |Actual ROWCOUNT : "&Row&" in Document" ,"", False, Config.Item("ReportPath"),True)
			       	 	ErrorNO = 8  'raise a user-defined error
			        	endCommand "Fail"
			        
			        else 
						Call InsertIntoHTMLReport( "endOfTestStep", "Check Document (" & docType & ")", cCustomErrorGeneration(customError) & " Expected ROWCOUNT :"&inputString&" |Actual ROWCOUNT : "&Row&" in Document" ,"", False, Config.Item("ReportPath"),True)
						endCommand "Fail"
						ErrorNO = 0   'raise a user-defined error
						'Err.Description = oObject & " disabled"
					End If
				End If

			ExcelObject.ActiveWorkbook.Close
			ExcelObject.Application.Quit
			Set ExcelSheet = Nothing
			Set ExcelObject = Nothing
			
ElseIf verifyType = "TABLEDATA" Then
	        Set ExcelObject = createobject("excel.application")
			'Dim xlsheet
			ExcelObject.Workbooks.Open filePath
			ExcelObject.Application.Visible = false
		
			sheets = Split(pageNumberRange,",")
			For Each sheet In sheets
				Set ExcelSheet = ExcelObject.ActiveWorkbook.Worksheets(sheet)

				Row = ExcelSheet.UsedRange.Row + ExcelSheet.UsedRange.Rows.Count - 1
				Col = ExcelSheet.UsedRange.Column + ExcelSheet.UsedRange.Columns.Count - 1
			
				Dim values
				Dim excelArr,checkArr
				values=""
				
				For  i= 1 to Row 
					For j=1 to Col
						
						If ExcelSheet.cells(i,j).value = "" Then
						Else
								If values="" Then
									values=ExcelSheet.cells(i,j).value
								Else
									values=values&"|#|"&ExcelSheet.cells(i,j).value
								End If								
						End If
						Next
				Next
			Next

		
		excelArr=Split(values,"|#|")
		checkArr=Split(inputString,"|#|")
		
		Dim a
		a= CompareArraysNotOrder (excelArr,checkArr)
		If a Then
			Call InsertIntoHTMLReport( "endOfTestStep","CheckDocument (" & docType & ")","Input TABLEDATA is available in Document", "", True, Config.Item("ReportPath"),False)
			endCommand "Pass"
		Else 
		    If  stopOnFailure = True Then
			       	 	Call InsertIntoHTMLReport( "endOfTestStep", "Check Document (" & docType & ")", cCustomErrorGeneration(customError) & " Input TABLEDATA Value is not available in Document" ,"", False, Config.Item("ReportPath"),True)
			       	 	ErrorNO = 8  'raise a user-defined error
			        	endCommand "Fail"
			        
			        else 
						Call InsertIntoHTMLReport( "endOfTestStep", "Check Document (" & docType & ")", cCustomErrorGeneration(customError) & " Input TABLEDATA Value is not available in Document" ,"", False, Config.Item("ReportPath"),True)
						endCommand "Fail"
						ErrorNO = 0   'raise a user-defined error
						'Err.Description = oObject & " disabled"
					End If
		End If
			

			ExcelObject.ActiveWorkbook.Close
			ExcelObject.Application.Quit
			Set ExcelSheet = Nothing
			Set ExcelObject = Nothing
		
ElseIf verifyType = "TABLECELL" Then
	        Set ExcelObject = createobject("excel.application")
			
			ExcelObject.Workbooks.Open filePath
			ExcelObject.Application.Visible = false
		
			sheets = Split(pageNumberRange,",")
			For Each sheet In sheets
				Set ExcelSheet = ExcelObject.ActiveWorkbook.Worksheets(sheet)

				Row = ExcelSheet.UsedRange.Row + ExcelSheet.UsedRange.Rows.Count - 1
				Col = ExcelSheet.UsedRange.Column + ExcelSheet.UsedRange.Columns.Count - 1
				
				checkArr=Split(inputString,"|")
				
				If Cstr(ExcelSheet.Cells(CInt(checkArr(0)),CInt(checkArr(1)))) = checkArr(2) Then
					Call InsertIntoHTMLReport( "endOfTestStep","CheckDocument (" & docType & ")","Input TABLECELL Value (" & checkArr(2) & ") is available in Document", "", True, Config.Item("ReportPath"),False)
			        endCommand "Pass"
					ExcelObject.ActiveWorkbook.Close		
					ExcelObject.Application.Quit		
					Set ExcelSheet = Nothing		
					Set ExcelObject = Nothing
			        Exit Function
				Else
					 If  stopOnFailure = True Then
			       	 	Call InsertIntoHTMLReport( "endOfTestStep", "Check Document (" & docType & ")", cCustomErrorGeneration(customError) & " Input TABLECELL Value (" & checkArr(2) & ") is not available in Document." ,"", False, Config.Item("ReportPath"),True)
			       	 	ErrorNO = 8  'raise a user-defined error
			        	endCommand "Fail"
			        
			        else 
						Call InsertIntoHTMLReport( "endOfTestStep", "Check Document (" & docType & ")", cCustomErrorGeneration(customError) & " Input TABLECELL Value (" & checkArr(2) & ") is not available in Document." ,"", False, Config.Item("ReportPath"),True)
						endCommand "Fail"
						ErrorNO = 0   'raise a user-defined error
						'Err.Description = oObject & " disabled"
					End If
				End If
				
			Next

			ExcelObject.ActiveWorkbook.Close
			ExcelObject.Application.Quit
			Set ExcelSheet = Nothing
			Set ExcelObject = Nothing
			
ElseIf verifyType = "RELATIVE" Then
	        Set ExcelObject = createobject("excel.application")
			ExcelObject.Workbooks.Open filePath
			ExcelObject.Application.Visible = false
		
			sheets = Split(pageNumberRange,",")
			For Each sheet In sheets
				Set ExcelSheet = ExcelObject.ActiveWorkbook.Worksheets(sheet)

				Row = ExcelSheet.UsedRange.Row + ExcelSheet.UsedRange.Rows.Count - 1
				Col = ExcelSheet.UsedRange.Column + ExcelSheet.UsedRange.Columns.Count - 1

				values=""
				checkArr=Split(inputString,"|")
				For  i= 1 to Row 
					For j=1 to Col
					
						If ExcelSheet.cells(i,j).value = "" Then
						Else
							If InStr(ExcelSheet.cells(i,j).value, checkArr(0)) > 0 Then

								If InStr(ExcelSheet.cells(i,j+CInt(checkArr(1))).value,checkArr(2)) > 0 Then
								    result = "Pass"
								End If
								
								
							End If
						End If
					Next
				Next
			Next
		    If result = "Pass" Then
		    Call InsertIntoHTMLReport( "endOfTestStep","CheckDocument (" & docType & ")","Input string is available in Document", "", True, Config.Item("ReportPath"),False)
			endCommand "Pass"
			Else
		    	If  stopOnFailure = True Then
					Call InsertIntoHTMLReport( "endOfTestStep", "Check Document (" & docType & ")", cCustomErrorGeneration(customError) & " Input string is not available in Document" ,"", False, Config.Item("ReportPath"),True)
					endCommand "Fail"
					ErrorNO =  8    'raise a user-defined error
			

				else 
					Call InsertIntoHTMLReport( "endOfTestStep", "Check Document (" & docType & ")", cCustomErrorGeneration(customError) & " Input string is not available in Document" ,"", False, Config.Item("ReportPath"),True)
					endCommand "Fail"
					ErrorNO = 0   'raise a user-defined error
						'Err.Description = oObject & " disabled"
				End If
			
			End If

			ExcelObject.ActiveWorkbook.Close
			ExcelObject.Application.Quit
			Set ExcelSheet = Nothing
			Set ExcelObject = Nothing				
Else
			Call InsertIntoHTMLReport( "endOfTestStep", "Check Document (" & docType & ")", cCustomErrorGeneration(customError) &" Validation type not found." , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			ErrorNO = 8    'raise a user-defined error
			Err.Description = "table validation type not found"			
					
End If
End Function

Function CompareArraysNotOrder (arrArray1, arrArray2)

   Dim intArray1, intArray2

   For intArray1 = 1 to UBound (arrArray1)
      For intArray2 = 1 to UBound (arrArray2)
         If arrArray1 (intArray1) = arrArray2 (intArray2) Then
     arrArray1 (intArray1) = "MATCHED":  arrArray2 (intArray2) = "MATCHED"
     Exit For
 End If
      Next
   Next

   CompareArraysNotOrder = True
   For intArray1 = 1 to UBound (arrArray1)
      If  arrArray1 (intArray1) <> "MATCHED" Then
 CompareArraysNotOrder = False
 Exit For
      End If
   Next

End Function
