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
'runId = Parameter("aRunId")
suiteId = Parameter("aSuiteId")
'suiteRunId = Parameter("aSuiteRunId")


'Connect to Octane
Dim restConnector, connectionInfo, isConnected
Set restConnector = DotNetFactory.CreateInstance("MicroFocus.Adm.Octane.Api.Core.Connector.RestConnector", "MicroFocus.Adm.Octane.Api.Core")
Set connectionInfo = DotNetFactory.CreateInstance("MicroFocus.Adm.Octane.Api.Core.Connector.UserPassConnectionInfo", "MicroFocus.Adm.Octane.Api.Core", clientId, clientSecret)
isConnected = restConnector.Connect(octaneUrl, connectionInfo)

Dim context, entityService
Set context = DotNetFactory.CreateInstance("MicroFocus.Adm.Octane.Api.Core.Services.RequestContext.WorkspaceContext", "MicroFocus.Adm.Octane.Api.Core", sharedSpaceId, workspaceId)
Set entityService = DotNetFactory.CreateInstance("MicroFocus.Adm.Octane.Api.Core.Services.NonGenericsEntityService", "MicroFocus.Adm.Octane.Api.Core", restConnector)


'Get Test Suite's fields values from Octane by Suite ID (received from Jenkins)
Dim entType, entId, entFields, entFieldsAttach
entType = "test_suite"
entId = suiteId
entFields = Array("id", "name", "source_id_udf", "author", "creation_time")
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
Dim testSuite, FSO, outfile
Set testSuite = entityService.GetById(context, entType, entId, entFields)
Set FSO = CreateObject("Scripting.FileSystemObject")
Set outFile = FSO.CreateTextFile("C:\Downloads\test_suite (automated, Jenkins).txt",True)
outFile.WriteLine "Test Suite ID: " + testSuite.Id
outFile.WriteLine vbCrLf & "Test Suite Name: " + testSuite.Name
outFile.WriteLine "ALM QC ID: " + testSuite.GetValue("source_id_udf")
outFile.WriteLine "Author: " + testSuite.GetValue("author").Name
outFile.WriteLine vbCrLf & "Creatin Time: " + testSuite.GetValue("creation_time")
outFile.WriteLine vbCrLf & "Attachments: " + attachmentsName
outFile.Close
