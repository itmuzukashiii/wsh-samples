var oFso = WScript.CreateObject("Scripting.FileSystemObject");
var oFile = oFso.GetFile("‚Ä‚·‚Æ.txt");
WScript.Echo(oFile.DateLastModified);

