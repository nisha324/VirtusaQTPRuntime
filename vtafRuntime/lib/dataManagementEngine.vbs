'VTAF QTP WEB run time
'-------------------------------------------
'Change log
'-------------------------------------------
' Date          Version   
' 2014.10.20    1.0
'-------------------------------------------

Dim Datatable
set Datatable = CreateObject("scripting.dictionary")


'@Function - GetDataTables
'@param - arrTables : string array contains table names
'@descrption - read data tables and load into a runtime map call @Datatable
'@auther - Vimukthi Hewapathirana 



Function GetDataTables(arrTables)
On Error Resume Next
	
Set xml = CreateObject("Microsoft.XMLDOM")
xmlFilePath=Environment.Value("TestDir")&"\vtafRuntime\testData\DataTables.xml"
'Load the XML file
xml.Load(xmlFilePath)
	
For Each table in arrTables
		
	Set tableNode=xml.selectNodes("//DataTables/Table[@name='"&table&"']")
	'@headerList - <HEADERLIST>
	Set headerList = CreateObject("System.Collections.ArrayList")
		
	Set columnNodes=xml.selectNodes("//DataTables/Table[@name='"&table&"']/Header/Column")
		
	For Each column In  columnNodes
		'@rows - <ROWLIST>
		Set rowlist = CreateObject("System.Collections.ArrayList")
						
		columnName=table&"_"&column.getAttribute("name")
		'@Datatable- MAP<KEY:[TABLENAME_HEADER]><VALUE:[<ROWLIST>]>
		Datatable.Add columnName,rowlist 
		headerList.Add columnName  	
		Set rowlist = Nothing
	Next
	
	size=(columnNodes.length)
	Set columnNodes = Nothing
				
	Set rowNodes=xml.selectNodes("//DataTables/Table[@name='"&table&"']/Row")
	Set valueNodes=xml.selectNodes("//DataTables/Table[@name='"&table&"']/Row/Value")
	numberOfRows=rowNodes.length		
	counter=0
	
		
	For Each value In valueNodes
		Datatable(headerList(counter)).Add value.text
		counter=counter+1
		If (counter=size) Then
			counter=0
		End If
	Next
	Set	rowNodes = Nothing
	Set valueNodes = Nothing
	Set headerList = Nothing
	Set tableNode = Nothing
Next

If Err.Number<>0 Then
	
	Print "[EROR] Error while reading data xml. Execution may Fail or become inconsistant"
End If

GetDataTables=(numberOfRows-1)
	
End Function







