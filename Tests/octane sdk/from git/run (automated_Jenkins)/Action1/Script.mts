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
'suiteId = Parameter("aSuiteId")
'suiteRunId = Parameter("aSuiteRunId")


'Connect to Octane
Dim restConnector, connectionInfo, isConnected
Set restConnector = DotNetFactory.CreateInstance("MicroFocus.Adm.Octane.Api.Core.Connector.RestConnector", "MicroFocus.Adm.Octane.Api.Core")
Set connectionInfo = DotNetFactory.CreateInstance("MicroFocus.Adm.Octane.Api.Core.Connector.UserPassConnectionInfo", "MicroFocus.Adm.Octane.Api.Core", clientId, clientSecret)
isConnected = restConnector.Connect(octaneUrl, connectionInfo)

Dim context, entityService
Set context = DotNetFactory.CreateInstance("MicroFocus.Adm.Octane.Api.Core.Services.RequestContext.WorkspaceContext", "MicroFocus.Adm.Octane.Api.Core", sharedSpaceId, workspaceId)
Set entityService = DotNetFactory.CreateInstance("MicroFocus.Adm.Octane.Api.Core.Services.NonGenericsEntityService", "MicroFocus.Adm.Octane.Api.Core", restConnector)


'Get ZRun's fields values from Octane by Run ID
Dim entType, entId, entFields, entFieldsAttach
entType = "run"
entId = runId
entFields = Array("id", "name", "test_name", "test", "run_by", "started", "native_status", "parent_suite")
entFieldsAttach = Array("id", "name")


'Get attachments and download
Dim attachmentsList, attachmentsList1, attachmentsList2, attachmentsName, orderBy, limit, offset
orderBy = "id"
limit = CInt(2)
offset = CInt(0)
Set attachmentsList = entityService.Get(context, "attachments", "(owner_test={id=" + entId + "})", entFieldsAttach)
'Set attachmentsList1 = entityService.Get(context, "attachments", "(owner_test={id=" + entId + "})", entFields, 1, 1)
'Set attachmentsList2 = entityService.Get(context, "attachments", "(owner_test={id=" + entId + "})", entFields, "id", 1, 0)

attachmentsName = ""
Dim i, element
For i = 0 To attachmentsList.BaseEntities.Count - 1
	Set element = attachmentsList.BaseEntities.Item(CInt(i))
	If (Len(attachmentsName) > 0) Then
		attachmentsName = attachmentsName + ", "
	End If
	attachmentsName = attachmentsName + element.Name
	entityService.DownloadAttachment "/api/shared_spaces/" +sharedSpaceId+ "/workspaces/" +workspaceId+ "/attachments/" +element.Id+ "/" + element.Name, "C:\\Downloads\\" +element.Name
Next


'Write results to text file
Dim run, FSO, outfile
Set run = entityService.GetById(context, entType, entId, entFields)
Set FSO = CreateObject("Scripting.FileSystemObject")
Set outFile = FSO.CreateTextFile("C:\Downloads\run (automated, Jenkins).txt",True)
outFile.WriteLine "Run ID: " + entId
outFile.WriteLine "Run Name: " + run.GetValue("name")
outFile.WriteLine "Run by: " + run.GetValue("run_by").Name
outFile.WriteLine "Started: " + run.GetValue("started")
outFile.WriteLine "Run Status: " + run.GetValue("native_status").Id
outFile.WriteLine vbCrLf & "Test ID: " + run.GetValue("test").Id
outFile.WriteLine "Test Name: " + run.GetValue("test_name")
outFile.WriteLine vbCrLf & "Suite Run ID: " + run.GetValue ("parent_suite").Id
outFile.WriteLine vbCrLf & "Attachments: " + attachmentsName
outFile.Close
