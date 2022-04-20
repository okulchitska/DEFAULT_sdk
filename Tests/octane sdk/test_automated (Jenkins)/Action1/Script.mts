Option Explicit

Dim MyMsgBox
Set MyMsgBox = DotNetFactory.CreateInstance("System.Windows.Forms.MessageBox", "System.Windows.Forms")

Dim clientId, clientSecret, octaneUrl
Dim sharedSpaceId, workspaceId, runId

clientId = Parameter("aClientId")
clientSecret = Parameter("aClientSecret")
octaneUrl = Parameter("aOctaneUrl")
sharedSpaceId = Parameter("aOctaneSpaceId")
workspaceId = Parameter("aOctaneWorkspaceId")
runId = Parameter("aRunId")
'suiteId = Parameter("aSuiteId")
'suiteRunId = Parameter("aSuiteRunId")

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

Dim testType, testId, testFields, testFieldsAttach
testType = "test_automated" 'run.GetValue("test").Type
testId = run.GetValue("test").Id
testFields = Array("id", "subtype", "name", "author", "owner", "test_runner")
testFieldsAttach = Array("id", "name")
'MyMsgBox.Show testType + " " + testId

Dim attachmentsList, attachmentsList1, attachmentsList2, attachmentsName, orderBy, limit, offset
orderBy = "id"
limit = CInt(2)
offset = CInt(0)
Set attachmentsList = entityService.Get(context, "attachments", "(owner_test={id=" + testId + "})", testFieldsAttach)
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
'MyMsgBox.Show "Attachments: " + attachmentsName, "Attachments"

'Write results to file
Dim test, FSO, outfile
Set test = entityService.GetById(context, testType, testId, testFields)
Set FSO = CreateObject("Scripting.FileSystemObject")
Set outFile = FSO.CreateTextFile("C:\Downloads\test_automated (Jenkins).txt",True)
outFile.WriteLine "Test ID: " + test.Id
outFile.WriteLine "Test Name: " + test.Name
outFile.WriteLine vbCrLf & "Test Type: " + test.Subtype
outFile.WriteLine "Author: " + test.GetValue("author").Name
outFile.WriteLine vbCrLf & "Owner: " + test.GetValue ("owner").Name
outFile.WriteLine "UFT test runner: #" + test.GetValue("test_runner").id + ", " + test.GetValue("test_runner").Name
outFile.WriteLine vbCrLf & "Attachments: " + attachmentsName
outFile.Close
