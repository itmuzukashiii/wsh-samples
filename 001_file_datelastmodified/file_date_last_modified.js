var oFso = WScript.CreateObject("Scripting.FileSystemObject");
var oFile = oFso.GetFile("�Ă���.txt");
WScript.Echo(oFile.DateLastModified);

