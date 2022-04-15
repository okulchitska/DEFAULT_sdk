Option Explicit

'Dim MyMsgBox
'Set MyMsgBox = DotNetFactory.CreateInstance("System.Windows.Forms.MessageBox", "System.Windows.Forms")

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

Dim entType, entId, entFields, entFieldsAttach, run
entType = "run"
entId = runId
entFields = Array("id", "test")
entFieldsAttach = Array("id", "name", "author")
Set run = entityService.GetById(context, entType, entId, entFields)

Dim testType, testId, testFields, mtId, atest, script
testType = "test_automated" 
testId = run.GetValue("test").Id
testFields = Array("id", "subtype", "name", "author", "owner", "test_runner", "covered_manual_test")
Set atest = entityService.GetById(context, "test", testId, testFields)
mtId = atest.GetValue("covered_manual_test").Id

Set script = entityService.GetTestScript(context, mtId)

'MyMsgBox.Show script.Script


'Write results to file
Dim test, FSO, outfile
Set test = entityService.GetById(context, testType, testId, testFields)
Set FSO = CreateObject("Scripting.FileSystemObject")
Set outFile = FSO.CreateTextFile("C:\Downloads\test_automated (Jenkins).txt",True)
outFile.WriteLine "Script: "
outFile.WriteLine vbCrLf & script.Script
outFile.Close
