' Copyright 2004 ThoughtWorks, Inc. Licensed under the Apache License, Version
' 2.0 (the "License"); you may not use this file except in compliance with the
' License. You may obtain a copy of the License at
' http://www.apache.org/licenses/LICENSE-2.0 Unless required by applicable law
' or agreed to in writing, software distributed under the License is
' distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
' KIND, either express or implied. See the License for the specific language
' governing permissions and limitations under the License.


RegisterUserFunc "WebList", "RegexSelectDOM", "RegexSelectDOM"
Dim retryTime
retryTime = 25

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


Public Function cCheckElementPresent (oObject, identifire , AssertType)
	On Error Resume Next
	SendObject oObject, identifire
	startCommand "CheckElimentPresent  Command", CommandObj.ToString
	Browser("title:=.*").Page("title:=.*").Sync
    If CommandObj.Exist(retryTime)Then
			Call InsertIntoHTMLReport( "endOfTestStep","CheckElimentPresent",oObject , "", True, Config.Item("ReportPath"),False)
			endCommand "Pass"
	Else 
			If  AssertType = True Then
					Call InsertIntoHTMLReport( "endOfTestStep","CheckElimentPresent",oObject , "", False, Config.Item("ReportPath"),True)
					endCommand "Fail"
					Err.Raise 8    'raise a user-defined error
					Err.Description = CommandObj.ToString & " disabled"

			else 
					Call InsertIntoHTMLReport( "endOfTestStep","CheckElimentPresent",oObject , "", False, Config.Item("ReportPath"),True)
					endCommand "Fail"
					Err.Clear   'raise a user-defined error
						'Err.Description = CommandObj.ToString & " disabled"
			End If
End if 
End Function



Public Function cType (oObject, identifire, sValue)
	On Error Resume Next
	SendObject oObject, identifire
	startCommand "Type Command", CommandObj.ToString
    If CommandObj.Exist(retryTime)Then
		If CommandObj.GetROProperty("disabled") = 0 Then
			CommandObj.Set sValue
			Browser("title:=.*").Page("title:=.*").Sync 'DR
			Call InsertIntoHTMLReport( "endOfTestStep","Type",oObject , "", True, Config.Item("ReportPath"),False)
			endCommand "Pass"
		Else
			Call InsertIntoHTMLReport( "endOfTestStep","Type",oObject , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			Err.Raise 8    'raise a user-defined error
			Err.Description = CommandObj.ToString & " disabled"
		End If
	Else
        Call InsertIntoHTMLReport( "endOfTestStep","Type",oObject , "", False, Config.Item("ReportPath"),True)
		endCommand "Fail"
		Err.Raise 8    'raise a user-defined error
        Err.Description = CommandObj.ToString & " does not exist"
     End If
End Function

' Click Command-Drupasinghe-2012/01/03
Public Function cClick (oObject, identifire)
	On Error Resume Next
	SendObject oObject, identifire
	startCommand "Click  Command", CommandObj.ToString
    If CommandObj.Exist(retryTime)Then
		If CommandObj.GetROProperty("disabled") = 0 Then
			CommandObj.Click
			Browser("title:=.*").Page("title:=.*").Sync
			Call InsertIntoHTMLReport( "endOfTestStep","Click",oObject , "", True, Config.Item("ReportPath"),False)
			endCommand "Pass"
		Else
			Call InsertIntoHTMLReport( "endOfTestStep","Click",oObject , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			Err.Raise 8    'raise a user-defined error
			Err.Description = CommandObj.ToString & " disabled"
		End If
    Else
		Call InsertIntoHTMLReport( "endOfTestStep","Click",oObject , "", False, Config.Item("ReportPath"),True)
		endCommand "Fail"
		Err.Raise 8    'raise a user-defined error
        Err.Description = CommandObj.ToString & " does not exist"
	End If
End Function

Public Function cClickAt (oObject, identifire, coordinates)
	On Error Resume Next
	SendObject oObject, identifire
    startCommand "ClickAt Command", CommandObj.ToString
    
	If CommandObj.Exist(retryTime)Then
        X=split(coordinates, ",")(0)
        Y=split(coordinates, ",")(1)
        CommandObj.Click X,Y
		Browser("title:=.*").Page("title:=.*").Sync
        Call InsertIntoHTMLReport( "endOfTestStep","ClickAt",oObject , "", True, Config.Item("ReportPath"),False)
        endCommand "Pass"

    Else
        Call InsertIntoHTMLReport( "endOfTestStep","ClickAt",oObject , "", False, Config.Item("ReportPath"),True)
        endCommand "Fail"
        Err.Raise 8    'raise a user-defined error
        Err.Description = CommandObj.ToString & " does not exist"
    End If
End Function

'URL Open Command
Function cOpen (ByVal URL,byVal WaitTime)
	On Error Resume Next
	startCommand "Open Command", URL
    SystemUtil.Run "iexplore.exe",URL
	wait (WaitTime/1000)
	Call InsertIntoHTMLReport( "endOfTestStep","Open",URL , "", True, Config.Item("ReportPath"),False)
	endCommand "Pass"
End Function

'Browser Close command
Function CloseApp(byVal WaitTime)
	On Error Resume Next
	startCommand "Close Command", CommandObj.ToString
	SystemUtil.CloseProcessByName "iexplore.exe"
	wait (WaitTime)
End Function


'Select Command
Public Function cSelect( oObject, identifire, oOption)
	On Error Resume Next
    SendObject oObject, identifire
	Dim optionArray
	Set CommandObj=oObject
     optionArray= split(oOption, "=")
	 indexIndex =split(oOption, "=")(0)
	 indexOption =split(oOption, "=")(1)
	 startCommand "Select Command", CommandObj.ToString
     If CommandObj.Exist(retryTime)Then
		If  optionArray(0)="index" Then
			  If CommandObj.GetROProperty("disabled") = 0 Then
			If indexIndex = "index" Then
                CommandObj.Select (CInt(indexOption)-1)
			Else
				CommandObj.RegexSelectDOM oOption
			End If
			Call InsertIntoHTMLReport( "endOfTestStep","Select",oObject , "", True, Config.Item("ReportPath"),False)
			endCommand "Pass"
		Else
			Call InsertIntoHTMLReport( "endOfTestStep","Select",oObject , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			Err.Raise 8    'raise a user-defined error
			Err.Description = CommandObj.ToString & " disabled"
        End If
		ELSE
		'value came directly
			  If CommandObj.GetROProperty("disabled") = 0 Then
               ' CommandObj.Select oOption
				  CommandObj.Select oOption
			    Call InsertIntoHTMLReport( "endOfTestStep","Select",oObject , "", True, Config.Item("ReportPath"),False)
			    endCommand "Pass"
		Else
			Call InsertIntoHTMLReport( "endOfTestStep","Select",oObject , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			Err.Raise 8    'raise a user-defined error
			Err.Description = CommandObj.ToString & " disabled"
        End If
		End If
    Else
		Call InsertIntoHTMLReport( "endOfTestStep","Select",oObject , "", False, Config.Item("ReportPath"),True)
		endCommand "Fail"
		Err.Raise 8    'raise a user-defined error
        Err.Description = CommandObj.ToString & " does not exist"
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

Function cCheckObjectProperty (oObject, identifire, byval propertyName, expectedValue, AssertType)
	On Error Resume Next
	SendObject oObject, identifire
	startCommand "Check Object Property  Command", CommandObj.ToString
	If CommandObj.Exist(retryTime)Then
			If propertyName = "ALLOPTIONS" Then
				veifyDropDownValuePresent CommandObj, expectedValue,AssertType
			Elseif propertyName = "MISSINGOPTION" Then
				veifyDropDownValueNotPresent CommandObj, expectedValue,AssertType
			Elseif propertyName = "SELECTEDOPTION" Then
				veifyDropDownValueSelected 	 CommandObj, expectedValue,AssertType
			Elseif propertyName = "textContent" Then
				If CommandObj.GetROProperty("text")=expectedValue then
					Call InsertIntoHTMLReport( "endOfTestStep","checkObjectProperty",oObject , "", True, Config.Item("ReportPath"),False)
					endCommand "Pass" 
				End If
			ElseIf CommandObj.GetROProperty(propertyName)=expectedValue then
				Call InsertIntoHTMLReport( "endOfTestStep","checkObjectProperty",oObject , "", True, Config.Item("ReportPath"),False)
				endCommand "Pass"     
			else 
					If AssertType = True  Then
                        Call InsertIntoHTMLReport("endOfTestStep","checkObjectProperty",oObject , "", False, Config.Item("ReportPath"),True)
						endCommand "Fail"
						Err.Raise 8    'raise a user-defined error
						Err.Description = "object property not found"
					else 
	                    Call InsertIntoHTMLReport( "endOfTestStep","CheckElimentPresent",oObject , "", False, Config.Item("ReportPath"),True)
						endCommand "Fail"
						Err.Clear    'raise a user-defined error
						End If
			End If
			Call InsertIntoHTMLReport( "endOfTestStep","checkObjectProperty",oObject , "", True, Config.Item("ReportPath"),False)
			endCommand "Pass"                                
	Else
			Call InsertIntoHTMLReport( "endOfTestStep","checkObjectProperty",oObject , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			Err.Raise 8    'raise a user-defined error
			Err.Description = CommandObj.ToString & " does not exist"
   End If
End Function







'  Drupasinghe 2013/02/07 - This is fynalyzed 
'  This can be used to verify the value of a drop down list
 '  User need to pass to object and  expected values
' expected values should be in - "Danushka;Nadie;Damith" - format
'  Sample object -Browser("Browser").Page("Tryit Editor v1.6").Frame("Frame").WebList("select")  ----Testing purposes only
Function veifyDropDownValuePresent (cmdObject, expectedValues,AssertDrpPresent)
	On Error Resume Next
	actualValues  = cmdObject.GetROProperty("all items")
	If  actualValues = expectedValues Then
		Call InsertIntoHTMLReport( "endOfTestStep","checkObjectProperty",oObject , "", True, Config.Item("ReportPath"),False)
		endCommand "Pass" 
    Else 
		If  AssertDrpPresent = True Then
			Call InsertIntoHTMLReport( "endOfTestStep","CheckElimentPresent",oObject , "", False, Config.Item("ReportPath"),True)
					endCommand "Fail"
					Err.Raise 8    'raise a user-defined error
					Err.Description = CommandObj.ToString & " Wrong Expected values"
		else 
					Call InsertIntoHTMLReport( "endOfTestStep","CheckElimentPresent",oObject , "", False, Config.Item("ReportPath"),True)
					endCommand "Fail"
					Err.Clear    'raise a user-defined error
		end if 
   End If
End Function


'  Drupasinghe 2013/02/07 - This is fynalyzed 
'  This can be used to verify the given values are not present in drop down list
 '  User need to pass to object and  expected values
' expected values should be in - "Danushka;Nadie;Damith" - format
'  Sample object -Browser("Browser").Page("Tryit Editor v1.6").Frame("Frame").WebList("select")  ----Testing purposes only
Function veifyDropDownValueNotPresent (cmdObject, expectedValues1,AssertDrpNotPresent)
	On Error Resume Next
	ExpectedErr = false
	actualValues1  = cmdObject.GetROProperty("all items")
	expectedArray = Split(expectedValues1, ";")
	actualArray = Split(actualValues1, ";")

	For i = 0 To Ubound(expectedArray)
		For j = 0 To Ubound(actualArray)
			If expectedArray(i) = actualArray(j)Then
				ExpectedErr = true
            End If
		Next
	Next

	If ExpectedErr = true Then
				If  AssertDrpNotPresent = True Then
					Call InsertIntoHTMLReport("endOfTestStep","checkObjectProperty",oObject , "", False, Config.Item("ReportPath"),True)
					endCommand "Fail"
					Err.Raise 8    'raise a user-defined error
					Err.Description = "object property not found"

				else 
	                Call InsertIntoHTMLReport( "endOfTestStep","CheckElimentPresent",oObject , "", False, Config.Item("ReportPath"),True)
					endCommand "Fail"
					Err.Clear    'raise a user-defined error

				end if 
	Else
		Call InsertIntoHTMLReport("endOfTestStep","checkObjectProperty",oObject , "", True, Config.Item("ReportPath"),False)
		endCommand "Pass"
	End If
End Function


 
 Function veifyDropDownValueSelected (cmdObject, expectedValues,AssertDrpSelected)
	On Error Resume Next
	actualValues  = cmdObject.GetROProperty("selection")
	If  actualValues = expectedValues Then
		Call InsertIntoHTMLReport( "endOfTestStep","checkObjectProperty",oObject , "", True, Config.Item("ReportPath"),False)
		endCommand "Pass" 
    Else 
		If  AssertDrpSelected = True Then
			Call InsertIntoHTMLReport( "endOfTestStep","CheckElimentPresent",oObject , "", False, Config.Item("ReportPath"),True)
					endCommand "Fail"
					Err.Raise 8    'raise a user-defined error
					Err.Description = CommandObj.ToString & " Wrong Expected values"
		else 
					Call InsertIntoHTMLReport( "endOfTestStep","CheckElimentPresent",oObject , "", False, Config.Item("ReportPath"),True)
					endCommand "Fail"
					Err.Clear    'raise a user-defined error
		end if 
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


Function cCheckTable (oObject, identifire, validationType,expectedValue,AssertType)
	On Error Resume Next
	SendObject oObject, identifire
	If CommandObj.Exist(retryTime)Then
		If  validationType="ROWCOUNT"  Then
			validateTableRowCount CommandObj, expectedvale,AssertType
		ElseIf validationType = "COLCOUNT" Then
			validateTableColCount CommandObj, expectedValue,AssertType
		ElseIf  validationType = "TABLECELL" Then 
			validateTableCell CommandObj,expectedValue ,AssertType
		ElseIf  validationType = "RELATIVE" Then 
			validateTableOffset CommandObj, expectedValue, AssertType
		Else
			Call InsertIntoHTMLReport( "endOfTestStep","Validation type not found",oObject , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			Err.Raise 8    'raise a user-defined error
			Err.Description = "table validation type not found"			
		End if
	Else 
		Call InsertIntoHTMLReport( "endOfTestStep"," Object not found ",oObject , "", False, Config.Item("ReportPath"),True)
		endCommand "Fail"	
		Err.Raise 8    'raise a user-defined error
		Err.Description = CommandObj.ToString & " does not exist"			
	End If
End Function 


'''' Table Row Count Validaion 
Public Function validateTableRowCount (cmdObject,expectedValue,AssertTypeRowCwnt)
	On Error Resume Next
	If  cmdObject.RowCount = CLng(expectedValue)  Then
		Call InsertIntoHTMLReport( "endOfTestStep","validateTableRowCount ",oObject , "", True, Config.Item("ReportPath"),False)
		endCommand "Pass"		
	Else 
	
		If AssertTypeRowCwnt = True  Then
            Call InsertIntoHTMLReport("endOfTestStep","checkObjectProperty",oObject , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			Err.Raise 8    'raise a user-defined error
			Err.Description = "Row Count Error"
		else 
	        Call InsertIntoHTMLReport( "endOfTestStep","CheckElimentPresent",oObject , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			Err.Clear    'raise a user-defined error
		End If
	
	End if 
End Function




'	Table Column Count validation 
Public Function validateTableColCount (cmdObject,expectedValue,AssertTypeColCwnt)
	On Error Resume Next
	If  cmdObject.Columncount = CLng(expectedValue)  Then
		Call InsertIntoHTMLReport( "endOfTestStep","validateTableColumCount ",oObject , "", True, Config.Item("ReportPath"),False)
		endCommand "Pass"	
	Else 
		If AssertTypeColCwnt = True  Then
            Call InsertIntoHTMLReport("endOfTestStep","checkObjectProperty",oObject , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			Err.Raise 8    'raise a user-defined error
			Err.Description = "Column Count Error"
		else 
	        Call InsertIntoHTMLReport( "endOfTestStep","CheckElimentPresent",oObject , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			Err.Clear    'raise a user-defined error
		End If
	End if     	
End Function


'		Table Cell validation 
Public function validateTableCell (cmdObject,expectedValue,AssertTableCell)
	On Error Resume Next
	Dim row, col
	Dim a
	a = Split(expectedValue, ",")
	row = CInt(a(0))
	col = CInt(a(1))
	expectedValue = a(2)
	If  cmdObject.GetCellData(row,col) = expectedValue Then
		Call InsertIntoHTMLReport( "endOfTestStep","validateTableCell ",oObject , "", True, Config.Item("ReportPath"),False)
		endCommand "Pass"	
	Else
	
		If AssertTableCell = True  Then
            Call InsertIntoHTMLReport("endOfTestStep","checkObjectProperty",oObject , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			Err.Raise 8    'raise a user-defined error
			Err.Description = "Table Cell value error"
		else 
	        Call InsertIntoHTMLReport( "endOfTestStep","CheckElimentPresent",oObject , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			Err.Clear    'raise a user-defined error
		End If
			
	End If
End Function



'		Table Relative Data validation 
Public Function validateTableOffset(cmdObject, expectedValue,AssertTableOffSet)
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
		validateCellOffset cmdObject, rd, o, ev
	Next
End Function

'		Relative Data Cell Validation
Public Function validateCellOffset(cmdObject, referenceData, offset, expectedValue, AsserValueTableOffSet)
	On Error Resume Next
	For row = 1 To cmdObject.RowCount
		For col=1 To cmdObject.ColumnCount(row)
			If cmdObject.GetCellData(row, col) = referenceData Then
				If cmdObject.GetCellData(row, col + offset) = expectedValue Then
					Call InsertIntoHTMLReport( "endOfTestStep","validateCellOffset ",oObject , "", True, Config.Item("ReportPath"),False)
					endCommand "Pass"	
				Else
					If AsserValueTableOffSet = True  Then
                        Call InsertIntoHTMLReport("endOfTestStep","checkObjectProperty",oObject , "", False, Config.Item("ReportPath"),True)
						endCommand "Fail"
						Err.Raise 8    'raise a user-defined error
						Err.Description = "offset error"
					else 
	                    Call InsertIntoHTMLReport( "endOfTestStep","CheckElimentPresent",oObject , "", False, Config.Item("ReportPath"),True)
						endCommand "Fail"
						Err.Clear    'raise a user-defined error
						End If		
				End If
			End If
		Next
	Next
End Function




' 	Dummy Command -not impliemenetd yet this will implimenetd when ever required - Drupasinghe 2012/02/19
Public Function cFail()
	On Error Resume Next
    Call InsertIntoHTMLReport( "endOfTestStep","Fail Command is not yet Implemented ",oObject , "", True, Config.Item("ReportPath"),True)
	endCommand "Fail"
	Err.Raise 8    'raise a user-defined error
	Err.Description = "Fail Command is not yet Implimented"	
End Function


' 	Dummy Command -not impliemenetd yet this will implimenetd when ever required - Drupasinghe 2012/02/19
Public Function cGoBack
	On Error Resume Next
    Call InsertIntoHTMLReport( "endOfTestStep","Goback Command is not yet Implemented ",oObject , "", True, Config.Item("ReportPath"),True)
	endCommand "Fail"	
	Err.Raise 8    'raise a user-defined error
	Err.Description = "Goback Command is not yet Implemented"		
End Function



'Public Function cSelectWindow
Public Function cSelectWindow (oObject, identifire)
	On Error Resume Next
	SendObject oObject, identifire
	startCommand "SelectWindow Command", CommandObj.ToString
    If CommandObj.Exist(retryTime)Then
			CommandObj.Activate
			Browser("title:=.*").Page("title:=.*").Sync 'DR
			CommandObj.Maximize
			Browser("title:=.*").Page("title:=.*").Sync 'DR
			Call InsertIntoHTMLReport( "endOfTestStep","SelectWindow",oObject , "", True, Config.Item("ReportPath"),False)

	Else
			Call InsertIntoHTMLReport( "endOfTestStep","SelectWindow",oObject , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			Err.Raise 8    'raise a user-defined error
			Err.Description = CommandObj.ToString & " does not exist"
    End If
End Function

' 	Dummy Command -not impliemenetd yet this will implimenetd when ever required - Drupasinghe 2012/02/19
Public Function cGetObjectCount
	On Error Resume Next
    Call InsertIntoHTMLReport( "endOfTestStep","GetObjectCount Command is not yet Implemented ",oObject , "", True, Config.Item("ReportPath"),True)
	endCommand "Fail"
	Err.Raise 8    'raise a user-defined error
	Err.Description = "GetObjectCount Command is not yet Implemented"		
End Function

' 	Dummy Command -not impliemenetd yet this will implimenetd when ever required - Drupasinghe 2012/02/19
Public Function cDoubleClickAt
	On Error Resume Next
    Call InsertIntoHTMLReport( "endOfTestStep","DoubleClickAt Command is not yet Implimented ",oObject , "", True, Config.Item("ReportPath"),True)
	endCommand "Fail"	
	Err.Raise 8    'raise a user-defined error
	Err.Description = "DoubleClickAt Command is not yet Implemented"		
End Function

Public Function cDoubleClick (oObject, identifire)
	On Error Resume Next
	SendObject oObject, identifire
	startCommand "DoubleClick Command", CommandObj.ToString
    If CommandObj.Exist(retryTime)Then
		If CommandObj.GetROProperty("disabled") = 0 Then
			CommandObj.FireEvent "dblClick"
			Browser("title:=.*").Page("title:=.*").Sync 'DR
			Call InsertIntoHTMLReport( "endOfTestStep","DoubleClick",oObject , "", True, Config.Item("ReportPath"),False)
			endCommand "Pass"
		Else
			Call InsertIntoHTMLReport( "endOfTestStep","DoubleClick",oObject , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			Err.Raise 8    'raise a user-defined error
			Err.Description = CommandObj.ToString & " disabled"
		End If
	Else
        Call InsertIntoHTMLReport( "endOfTestStep","DoubleClick",oObject , "", False, Config.Item("ReportPath"),True)
		endCommand "Fail"
		Err.Raise 8    'raise a user-defined error
        Err.Description = CommandObj.ToString & " does not exist"
     End If
End Function

'  	wait command
Public Function cPause (WaitTime)
	On Error Resume Next
    waitTimeInMilSeconds = CInt(WaitTime)
	waitTimeInSeconds = waitTimeInMilSeconds/1000
    wait (waitTimeInSeconds)
	Call InsertIntoHTMLReport( "endOfTestStep","Pause",WaitTime , "", True, Config.Item("ReportPath"),False)
End Function

Function cKeyPress(oObject,identifire,KeyBoardInput)
   On Error Resume Next
   SendObject oObject, identifire
   startCommand "KeyPress Command", CommandObj.ToString

	   If KeyBoardInput="\t" Then
	   KeyHit = micTab

	   Else KeyHit =KeyBoardInput
	   end if

    If CommandObj.Exist(retryTime)Then
		If CommandObj.GetROProperty("disabled") = 0 Then
			CommandObj.Type KeyHit
			Call InsertIntoHTMLReport( "endOfTestStep","Type",oObject , "", True, Config.Item("ReportPath"),False)
			endCommand "Pass"
			Else
			Call InsertIntoHTMLReport( "endOfTestStep","KeyPress",oObject , "", False, Config.Item("ReportPath"),True)
			endCommand "Fail"
			Err.Raise 8    'raise a user-defined error
			Err.Description = CommandObj.ToString & " disabled"
		End If
	Else
        Call InsertIntoHTMLReport( "endOfTestStep","KeyPress",oObject , "", False, Config.Item("ReportPath"),True)
		endCommand "Fail"
		Err.Raise 8    'raise a user-defined error
        Err.Description = CommandObj.ToString & " does not exist"
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
    End If
End Function


Function fireKeyEvent(var1)
	a=Split(var1,"|")
	for each x in a
		b=Split(x,"=")
		If b(0) = "key" Then
			cEscapeAndFire b(1)
		ElseIf b(0) = "wait" Then
			wait b(1)
		End If
	next
End Function



Public Function cFireEvent(var2,waitTime)
On Error Resume Next
	startCommand "Fire Event  Command", " "
	a=Split(var2,"%")
	If a(0)="KEY" Then
		fireKeyEvent(a(1))
		wait waitTime
		Call InsertIntoHTMLReport( "endOfTestStep","Fire Event",oObject , "", True, Config.Item("ReportPath"),False)
	endCommand "Pass"
	End If
End Function

Public Function cSetVarProperty (oObject, identifire, sProperty)
	On Error Resume Next
	SendObject oObject, identifire
	startCommand "SetVarProperty Command", CommandObj.ToString
    If CommandObj.Exist(retryTime)Then
	
		if CommandObj.GetROProperty(sProperty) = Empty  then 
			Call InsertIntoHTMLReport( "endOfTestStep","SetVarProperty",oObject , "", False, Config.Item("ReportPath"),True)
			Err.Raise 8  'raise a user-defined error
			endCommand "Fail"
		else 
			If sProperty = "textContent"Then
				cSetVarProperty=CommandObj.GetROProperty("text")
				Call InsertIntoHTMLReport( "endOfTestStep","SetVarProperty",oObject , "", True, Config.Item("ReportPath"),False)
				endCommand "Pass"
			Else
				cSetVarProperty=CommandObj.GetROProperty(sProperty)
				Call InsertIntoHTMLReport( "endOfTestStep","SetVarProperty",oObject , "", True, Config.Item("ReportPath"),False)
				endCommand "Pass"
			End If
		end if 
	Else
        Call InsertIntoHTMLReport( "endOfTestStep","SetVarProperty",oObject , "", False, Config.Item("ReportPath"),True)
		endCommand "Fail"
		Err.Raise 8    'raise a user-defined error
        Err.Description = CommandObj.ToString & " does not exist"
     End If
End Function


'
'Public Function cSetVarProperty (oObject, identifire, sProperty)
'	On Error Resume Next
'	SendObject oObject, identifire
'	startCommand "SetVarProperty Command", CommandObj.ToString
'    If CommandObj.Exist(retryTime)Then
'			cSetVarProperty=CommandObj.GetROProperty(sProperty)
'			Call InsertIntoHTMLReport( "endOfTestStep","SetVarProperty",oObject , "", True, Config.Item("ReportPath"),False)
'			endCommand "Pass"
'	Else
'        Call InsertIntoHTMLReport( "endOfTestStep","SetVarProperty",oObject , "", False, Config.Item("ReportPath"),True)
'		endCommand "Fail"
'		Err.Raise 8    'raise a user-defined error
'        Err.Description = CommandObj.ToString & " does not exist"
'     End If
'End Function

'cTableCell CommandObj,"","gridType"
'cTableCell CommandObj,"","gridSelect",4,5,""
Public Function cSetTable (oObject, identifire,Action,rownumber,columnnumber,data)
	On Error Resume Next
	columnnumber = "#"& columnnumber
	SendObject oObject, identifire
	startCommand "Dable grid  Command", CommandObj.ToString
    If CommandObj.Exist(retryTime)Then
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
			Err.Raise 8    'raise a user-defined error
			Err.Description = CommandObj.ToString & " does not exist"

	End If
End Function

