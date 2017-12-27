'*************************************************************************************************************
'Description:
'
'This example starts QuickTest, opens a new test, and adds a new action, which calls a second action.
'Then it edits the first action's script to move the call to the second action to a new position in the script,
'validates the syntax of the new script, defines some action parameters, and uploads the
'modified action script.

'************************************************************************************************************************
Set qtApp = CreateObject("QuickTest.Application") ' Create the application object
qtApp.Launch
qtApp.Visible = True

sFolder = "ScriptFiles"
Set oFSO = CreateObject("Scripting.FileSystemObject")


For Each oFile In oFSO.GetFolder(sFolder).Files
   
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
	Set NewAction = qtApp.Test.AddNewAction(ActionName, ActionDescr, ActionContent, False, qtAtBegining)
	'Use the Load Script function to store the content of the first new action into the script array
	script = Load_Script(NewAction.GetScript())
	ActionContent = Save_Script(script)
	'Set new script source to the action
	scriptError = NewAction.ValidateScript(ActionContent)
	NewAction.SetScript ActionContent

Next

Set oFSO = Nothing


Const E_SCRIPT_TABLE = 1
Const E_EMPTY_SCRIPT = 2
Const E_SCRIPT_NOT_VALID = 3


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