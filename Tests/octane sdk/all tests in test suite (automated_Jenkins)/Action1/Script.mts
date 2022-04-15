Option Explicit

Dim MyMsgBox
Set MyMsgBox = DotNetFactory.CreateInstance("System.Windows.Forms.MessageBox", "System.Windows.Forms")

Dim clientId, clientSecret, octaneUrl
Dim sharedSpaceId, workspaceId, runId, suiteId, suiteRunId

clientId = Parameter("aClientId")
clientSecret = Parameter("aClientSecret")
octaneUrl = Parameter("aOctaneUrl")
sharedSpaceId = Parameter("aOctaneSpaceId")
workspaceId = Parameter("aOctaneWorkspaceId")
runId = Parameter("aRunId")
suiteId = Parameter("aSuiteId")
suiteRunId = Parameter("aSuiteRunId")

Dim restConnector, connectionInfo, isConnected
Set restConnector = DotNetFactory.CreateInstance("MicroFocus.Adm.Octane.Api.Core.Connector.RestConnector", "MicroFocus.Adm.Octane.Api.Core")
Set connectionInfo = DotNetFactory.CreateInstance("MicroFocus.Adm.Octane.Api.Core.Connector.UserPassConnectionInfo", "MicroFocus.Adm.Octane.Api.Core", clientId, clientSecret)
isConnected = restConnector.Connect(octaneUrl, connectionInfo)
'MyMsgBox.Show  isConnected, "Is Connected"

Dim context, entityService
Set context = DotNetFactory.CreateInstance("MicroFocus.Adm.Octane.Api.Core.Services.RequestContext.WorkspaceContext", "MicroFocus.Adm.Octane.Api.Core", sharedSpaceId, workspaceId)
Set entityService = DotNetFactory.CreateInstance("MicroFocus.Adm.Octane.Api.Core.Services.NonGenericsEntityService", "MicroFocus.Adm.Octane.Api.Core", restConnector)


Dim entType, entId, entFields, query, testsList
entType = "run"
entId = runId
entFields = Array("id", "name", "test_name", "test", "run_by", "started", "native_status", "parent_suite")

query = "(test_suite={id=" + suiteId + "};!test={null})"
Set testsList = entityService.Get(context, "test_suite_link_to_tests", query, Array("id", "subtype", "test{id,name}"))

Dim i, element, testsNames
testsNames = ""
For i = 0 To testsList.BaseEntities.Count - 1
	Set element = testsList.BaseEntities.Item(CInt(i))
	If (Len(testsNames) > 0) Then
		testsNames = testsNames + ", "
	End If
	testsNames = testsNames + element.GetValue("test").Id + " " + element.GetValue("test").Name
Next
'MyMsgBox.Show testsNames, "tests"


'Write results to file
Dim run, FSO, outfile
Set run = entityService.GetById(context, entType, entId, entFields)
Set FSO = CreateObject("Scripting.FileSystemObject")
Set outFile = FSO.CreateTextFile("C:\Downloads\all tests from TS (automated, Jenkins).txt",True)
outFile.WriteLine "Test Suite ID: " + suiteId
outFile.WriteLine vbCrLf & "Tests: " + testsNames
'outFile.WriteLine "Run by: " + run.GetValue("run_by").Name
'outFile.WriteLine "Started: " + run.GetValue("started")
'outFile.WriteLine "Run Status: " + run.GetValue("native_status").Id
'outFile.WriteLine vbCrLf & "Test ID: " + run.GetValue("test").Id
'outFile.WriteLine "Test Name: " + run.GetValue("test_name")
'outFile.WriteLine vbCrLf & "Suite Run ID: " + run.GetValue ("parent_suite").Id
'outFile.WriteLine vbCrLf & 
outFile.Close
