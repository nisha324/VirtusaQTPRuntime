' Copyright 2004 ThoughtWorks, Inc. Licensed under the Apache License, Version
' 2.0 (the "License"); you may not use this file except in compliance with the
' License. You may obtain a copy of the License at
' http://www.apache.org/licenses/LICENSE-2.0 Unless required by applicable law
' or agreed to in writing, software distributed under the License is
' distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
' KIND, either express or implied. See the License for the specific language
' governing permissions and limitations under the License.


Dim ridcount, taginfo, node
Dim actExecution, actStaticTS, actTS, actUpperTC, paramsnode, actLowerTC, stepMsg, tagDataRow, itemtag, metainfotag
Dim objReport, fsoreport, readlogfile, fName, reportpath

Function generateVTAFReport()
executionName = executionName
reportpath = ProjectPath & "\vtafRuntime\testReports\vtafReport\" & executionName & "\" 
'Executions Start from Here
Set xmlDoc = CreateObject("Microsoft.XMLDOM")  
  
Set objReport = xmlDoc.createElement("report") 
xmlDoc.appendChild objReport  
setElementtoXml objReport, actExecution, "activity"
ridcount =0
'------------------------------------------------

Set fsoreport = CreateObject("Scripting.FileSystemObject")
Set readlogfile = fsoreport.OpenTextFile(reportpath & "LogFile\vtafsupportlog.txt")
do while not readlogfile.AtEndOfStream 
    fName =  readlogfile.ReadLine()
    detectTag fName
loop

End Function


Function detectTag(linesyntax)
	taginfo = split(linesyntax, "<#>")
	Dim att
	If taginfo(0) = "BeforeExecute" Then
		reportBeforeExecution att, taginfo
		
	ElseIf taginfo(0) = "BeforeTestSuite"  Then
		reportBeforeTestSuite att, taginfo
		
	ElseIf taginfo(0) = "BeforeTestCase"  Then
		reportBeforeTestCase att, taginfo

	ElseIf taginfo(0) = "TestStep"  Then
		reportAfterTestStep att, taginfo
	
	ElseIf taginfo(0) = "AfterTestCase"  Then
		reportAfterTestCase att, taginfo
		
	ElseIf taginfo(0) = "AfterTestSuite"  Then
		reportAfterTestSuite att, taginfo
	
	ElseIf taginfo(0) = "AfterExecute"  Then
		reportAfterExecution att, taginfo
		
	End If
	
	If taginfo(0) = "BeforeExecute" Then
		Set tagDetail = xmlDoc.createElement("detail")  
		tagDetail.Text = "Test Execution Report"
		actExecution.appendChild tagDetail  
	End If
End Function


Function reportBeforeExecution(att, taginfo)
		att = Array("user", "host", "osversion", "language", "screenresolution", "timestamp")
		setAtttoElement actExecution, "type", "root"
		recusiveAddAttributes actExecution, att
	
		initatt = Array("result","duration","totalerrorcount","totalwarningcount","totalsuccesscount","totalfailedcount","totalblockedcount")
		initvals = Array("","","","","","","")
		recusiveAddAttributeswithValue actExecution, initatt, initvals

End Function

Function reportBeforeTestSuite(att, taginfo)
		ridcount = ridcount +1
		setElementtoXml actExecution, actStaticTS, "activity"
		att = Array("testsuitename","runconfigname","runlabel","maxchildren","result","duration","type","rid")
		vals = Array("VTAF Test Execution Report","","","0","Success","UNKNOWN","test suite","a106eb7a56abd88")
		recusiveAddAttributeswithValue actStaticTS, att, vals
		
		setElementtoXml actStaticTS, actTS, "activity"
		setAtttoElement actTS, "type", "folder"
		setAtttoElement actTS, "rid", CStr(ridcount)
			
		att = Array("foldername")
		recusiveAddAttributes actTS, att	

		initatt = Array("result","duration")
		initvals = Array("","")
		recusiveAddAttributeswithValue actTS, initatt, initvals

		setElementtoXml actTS, paramsnode, "params"		
		
End Function

Function reportBeforeTestCase(att, taginfo)
		ridcount = ridcount +1
		setElementtoXml actTS, actUpperTC, "activity"
		att = Array("iterationcount","maxchildren","type","datasource")
		vals = Array("1","0","test case","")
		recusiveAddAttributeswithValue actUpperTC, att, vals
		
		setAtttoElement actUpperTC, "testcasename", taginfo(1)
		setAtttoElement actUpperTC, "rid", CStr(ridcount)
		
		initatt = Array("result","duration")
		initvals = Array("","")
		recusiveAddAttributeswithValue actUpperTC, initatt, initvals
		
		setElementtoXml actUpperTC, actLowerTC, "activity"
		att = Array("modulename")
		recusiveAddAttributes actLowerTC, att
		setAtttoElement actLowerTC, "moduletype", "UserCode"
		setAtttoElement actLowerTC, "rid", CStr(ridcount)
		setAtttoElement actLowerTC, "type", "test module"
			
		initatt = Array("result","duration")
		initvals = Array("","")
		recusiveAddAttributeswithValue actLowerTC, initatt, initvals
		
		setElementtoXml actLowerTC, tagDataRow, "datarow"

End Function

Function reportAfterExecution(att, taginfo)
		att = Array("duration", "result", "totalerrorcount", "totalwarningcount", "totalsuccesscount", "totalfailedcount", "totalblockedcount")
		recusiveAddAttributes actExecution, att
	Set objIntro = xmlDoc.createProcessingInstruction ("xml","version='1.0'")  
	xmlDoc.insertBefore _
	  objIntro,xmlDoc.childNodes(0)  
	xmlDoc.Save reportpath & "\report.html.data"  
End Function

Function reportAfterTestSuite(att, taginfo)
		att = Array("duration","result")
		recusiveAddAttributes actTS, att

End Function

Function reportAfterTestCase(att, taginfo)
		att = Array("duration","result")
		recusiveAddAttributes actLowerTC, att
		recusiveAddAttributes actUpperTC, att

End Function

Function reportAfterTestStep(att, taginfo)
		setElementtoXml actLowerTC, itemtag, "item"
		att = Array("time","level", "category")
		recusiveAddAttributes itemtag, att
			
		setElementtoXml itemtag, stepMsg, "message"
		setTexttoElement stepMsg, taginfo(6)
		setElementtoXml itemtag, metainfotag, "metainfo"
		setAtttoElement metainfotag, "codefile", taginfo(4)
		setAtttoElement metainfotag, "codeline", taginfo(5)
		setAtttoElement metainfotag, "loglvl", taginfo(2)
		
		If taginfo(2) = "Error" Then
			setAtttoElement metainfotag, "stacktrace", taginfo(7)
			setAtttoElement itemtag, "errimg", taginfo(8)
			setAtttoElement itemtag, "errthumb", taginfo(9)
		End If

End Function




'set Element and add in to parent node
'@param parentnde parent node
'@param parentnde child node
'@param parentnde element name
'@ none
Function setElementtoXml(parentnode, childnode, elename)
Set childnode = xmlDoc.createElement(elename) 
parentnode.appendChild childnode 
End Function

'set Text to the Element
'@param nodename node
'@param txtvalue value
'@ none
Function setTexttoElement(nodename, txtvalue)
	nodename.Text = txtvalue
End Function

'set Attribute value of the Element node
'@param node child node
'@param attname attribute name
'@param value value of the attribute
'@ none
Function setAtttoElement(node, attname, nodevalue)
	node.setAttribute attname,nodevalue
End Function

Function recusiveAddAttributes(nodeele, att)
		endloop = UBound(att)
		For i = 0 To endloop
			setAtttoElement nodeele, att(i), taginfo(i+1)
		Next
End Function

Function recusiveAddAttributeswithValue(nodename, att, vals)
		endloop = UBound(att)
		For i = 0 To endloop
			setAtttoElement nodename, att(i), vals(i)
		Next
End Function

