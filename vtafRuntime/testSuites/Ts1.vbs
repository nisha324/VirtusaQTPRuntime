

Function Tc1
	startTestCase "Tc1","" 	
 		
	endTestCase	
End Function

Function dataValue(iterator,colName)
	rowIndex=Int(iterator)    
	column_count = mySheet.Usedrange.Columns.Count
	dataValue = ""
	For colIndex=1 to column_count
		If(colName = mySheet.cells(1,colIndex).value) Then
			dataValue = mySheet.cells(rowIndex,colIndex).value
			Exit For
		End If
	Next
End Function

Function dataTableDir()
   strPath = Environment.Value("TestDir")
   dataTableDir = strPath & "\vtafRuntime\testData\" 
End Function
