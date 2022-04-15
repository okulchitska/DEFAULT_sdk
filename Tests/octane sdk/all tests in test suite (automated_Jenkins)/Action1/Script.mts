Option Explicit

Dim clientId, clientSecret, octaneUrl
Dim sharedSpaceId, workspaceId, runId, suiteId, suiteRunId

'These parameters should be added as inputs on the Start step and used as default values on the Action step
'The values for these parameters can be received from the Jenkins server
'octane_apiUser and octane_apiSecret should be added as additional parameters on Jenkins with providing API Access values
clientId = Parameter("aClientId")
clientSecret = Parameter("aClientSecret")
octaneUrl = Parameter("aOctaneUrl")
sharedSpaceId = Parameter("aOctaneSpaceId")
workspaceId = Parameter("aOctaneWorkspaceId")
runId = Parameter("aRunId")
suiteId = Parameter("aSuiteId")
suiteRunId = Parameter("aSuiteRunId")


'Connect to Octane
Dim restConnector, connectionInfo, isConnected
Set restConnector = DotNetFactory.CreateInstance("MicroFocus.Adm.Octane.Api.Core.Connector.RestConnector", "MicroFocus.Adm.Octane.Api.Core")
Set connectionInfo = DotNetFactory.CreateInstance("MicroFocus.Adm.Octane.Api.Core.Connector.UserPassConnectionInfo", "MicroFocus.Adm.Octane.Api.Core", clientId, clientSecret)
isConnected = restConnector.Connect(octaneUrl, connectionInfo)

Dim context, entityService
Set context = DotNetFactory.CreateInstance("MicroFocus.Adm.Octane.Api.Core.Services.RequestContext.WorkspaceContext", "MicroFocus.Adm.Octane.Api.Core", sharedSpaceId, workspaceId)
Set entityService = DotNetFactory.CreateInstance("MicroFocus.Adm.Octane.Api.Core.Services.NonGenericsEntityService", "MicroFocus.Adm.Octane.Api.Core", restConnector)


'Get a list of test IDs using the Test Suite ID (suiteId) received from Jenkins
Dim query, testsList
query = "(test_suite={id=" + suiteId + "};!test={null})"
Set testsList = entityService.Get(context, "test_suite_link_to_tests", query, Array("id", "subtype", "test{id,name}"))


'Get Tests Names
Dim i, element, testsNames
testsNames = ""
For i = 0 To testsList.BaseEntities.Count - 1
	Set element = testsList.BaseEntities.Item(CInt(i))
	If (Len(testsNames) > 0) Then
		testsNames = testsNames + ", "
	End If
	testsNames = testsNames + element.GetValue("test").Id + " " + element.GetValue("test").Name
Next


'Write results to file
Dim FSO, outfile
Set FSO = CreateObject("Scripting.FileSystemObject")
Set outFile = FSO.CreateTextFile("C:\Downloads\all tests from TS (automated, Jenkins).txt",True)
outFile.WriteLine "Test Suite ID: " + suiteId
outFile.WriteLine vbCrLf & "Tests: "
outFile.WriteLine + testsNames
outFile.Close
