'***********************************************************************
' UNIX to DOS file format converter.
'
' Takes files and/or folder paths and recursively converts all files
' to DOS format.
' Eg: Converts all LF (Line Feed) to CRLF (Carriage Return + Line Feed).
'
' Usage: Simply drag however many files and folders into this script.
'
' @attention: Requires UNIX2DOS.EXE, and must be in the same folder.
' @attention: Requires unix2dos_folder.vbs, and must be in the same folder.
'***********************************************************************

Set objArgs = Wscript.Arguments
Set objFso = CreateObject("Scripting.FileSystemObject")
Set objShell = Wscript.CreateObject("Wscript.Shell")
' Get the path of the scripts
Set objThisScript = objFso.GetFile(Wscript.ScriptFullName)
strParentFolder = objFso.GetParentFolderName(objThisScript)

' Globals
Dim strSubscript
Dim strUNIX2DOS

strSubscript = objFso.BuildPath(strParentFolder, "unix2dos_folder.vbs")
strUNIX2DOS = objFso.BuildPath(strParentFolder, "UNIX2DOS.EXE")

' Iterate through all the arguments passed
For i = 0 to objArgs.count
    on error resume next

    ' Try and treat the argument like a folder
    Set folder = objFso.GetFolder(objArgs(i))

    ' If we get an error, we know it is a file
    If err.number <> 0 then
        ' This is not a folder, treat as file
        ProcessFile(objArgs(i))
    Else
        ' No error? This is a folder, process accordingly
        ProcessFolder(objArgs(i))
    End if
    On Error Goto 0
Next

' Executes UNIX2DOS.EXE directly on the file.
Function ProcessFile(strFilePath)
    strCommand = strUNIX2DOS & " """ & strFilePath & """"
    ExecCommand(strCommand)
End Function

' Executes subscript "unix2dos_folder.vbs", passing it the folder argument.
Function ProcessFolder(strFolderPath)
    strCommand = "Cscript /Nologo """ & strSubscript & """ """ & strFolderPath & """"
    ExecCommand(strCommand)
End Function

' Executes a command (string) and echos the result
Function ExecCommand(strCommand)
    strResult = ""
'    WScript.Echo "strCommand = " & strCommand
    Set objExec = objShell.Exec(strCommand)
    Do While Not objExec.StdOut.AtEndOfStream
        strResult = strResult & objExec.StdOut.ReadLine() & vbCrlf
    Loop
    
    Do While Not objExec.StdErr.AtEndOfStream
        strResult = strResult & objExec.StdErr.ReadLine() & vbCrlf
    Loop
    Wscript.Echo strResult
End Function

