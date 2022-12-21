Option Explicit

Dim oFso: Set oFso = WScript.CreateObject("Scripting.FileSystemObject")
Dim oFile: Set oFile = oFso.GetFile("‚Ä‚·‚Æ.txt")
WScript.Echo(oFile.DateLastModified)
