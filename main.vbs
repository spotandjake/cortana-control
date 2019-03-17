Dim Arg, root, dir, strText, objFSO, objFolder, objShell, objTextFile, objFile, LocalAppData, oFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell") 
Set oFSO = CreateObject("Scripting.FileSystemObject") 
objShell.CurrentDirectory = oFSO.GetParentFolderName(WScript.ScriptFullName) 
LocalAppData=objShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\"
Set objShell = Nothing
Set oFSO = Nothing
Set Arg = WScript.Arguments
root = Arg(0)
If fso.FileExists(LocalAppData & "\Cortana-control" & "\keyboard.txt") Then
    'exist
else
    Set objFolder = objFSO.CreateFolder(LocalAppData & "\Cortana-control")
    Set objFolder = objFSO.GetFolder(LocalAppData & "\Cortana-control")
    Set objFile = objFSO.CreateTextFile(LocalAppData & "\Cortana-control" & "\keyboard.txt")
    Set objFile = objFSO.CreateTextFile(LocalAppData & "\Cortana-control" & "\x.txt")
    Set objFile = objFSO.CreateTextFile(LocalAppData & "\Cortana-control" & "\y.txt")
    Set objFile = objFSO.CreateTextFile(LocalAppData & "\Cortana-control" & "\click.txt")
end if
if root = "key" or root = "special" then 
strText = Arg(1)
set objFile = nothing
set objFolder = nothing
Set objTextFile = objFSO.OpenTextFile _
(LocalAppData & "\Cortana-control" & "\keyboard.txt", 2, True)
objTextFile.WriteLine(strText)
objTextFile.Close
ElseIf root = "move" then
dir = Arg(1)
if dir = "xy" or dir = "x" then 
strText = Arg(2)
set objFile = nothing
set objFolder = nothing
Set objTextFile = objFSO.OpenTextFile _
(LocalAppData & "\Cortana-control" & "\x.txt", 2, True)
objTextFile.WriteLine(strText)
objTextFile.Close
ElseIf dir = "xy" or dir = "y" then 
strText = Arg(3)
set objFile = nothing
set objFolder = nothing
Set objTextFile = objFSO.OpenTextFile _
(LocalAppData & "\Cortana-control" & "\y.txt", 2, True)
objTextFile.WriteLine(strText)
objTextFile.Close
End If
ElseIf root = "click" then
strText = Arg(1)
set objFile = nothing
set objFolder = nothing
Set objTextFile = objFSO.OpenTextFile _
(LocalAppData & "\Cortana-control" & "\click.txt", 2, True)
objTextFile.WriteLine(strText)
objTextFile.Close
End If
If objFSO.FileExists(LocalAppData & "\Cortana-control" & "\root.txt") Then
Set objFolder = objFSO.GetFolder(LocalAppData & "\Cortana-control")
Else
Set objFile = objFSO.CreateTextFile(LocalAppData & "\Cortana-control" & "\root.txt")
End If
set objFile = nothing
set objFolder = nothing
Set objTextFile = objFSO.OpenTextFile _
(LocalAppData & "\Cortana-control" & "\root.txt", 2, True)
objTextFile.WriteLine(root)
objTextFile.Close
WScript.Quit