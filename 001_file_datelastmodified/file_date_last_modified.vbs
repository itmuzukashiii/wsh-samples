Option Explicit

Dim oFso: Set oFso = WScript.CreateObject("Scripting.FileSystemObject")
Dim oFile: Set oFile = oFso.GetFile("�Ă���.txt")
WScript.Echo(oFile.DateLastModified)
