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


'Get Test ID from Automated Run ID (runId received from Jenkins)
Dim entType, entId, entFields, test
entType = "run"
entId = runId
entFields = Array("id", "test")
Set test = entityService.GetById(context, entType, entId, entFields)

'Get Manual Test ID linked to the Automated Test (using covered_manual_test parameter)
Dim testType, testId, testFields, manualTestId, automatedTest, script
testType = "test" 
testId = test.GetValue("test").Id
testFields = Array("id", "subtype", "name", "covered_manual_test")
Set automatedTest = entityService.GetById(context, "test", testId, testFields)
manualTestId = automatedTest.GetValue("covered_manual_test").Id

'Get Manual Test Script
Set script = entityService.GetTestScript(context, manualTestId)


'Write results to file
Dim FSO, outfile
Set FSO = CreateObject("Scripting.FileSystemObject")
Set outFile = FSO.CreateTextFile("C:\Downloads\script(Jenkins).txt",True)
outFile.WriteLine "Script: "
outFile.WriteLine vbCrLf & script.Script
outFile.Close
