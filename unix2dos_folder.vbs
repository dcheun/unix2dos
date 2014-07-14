'***********************************************************************
' UNIX to DOS file format converter.
'
' Takes a folder path as the first argument and recursively converts
' all files found to DOS format.
' Eg: Converts all LF (Line Feed) to CRLF (Carriage Return + Line Feed).
'
' Usage: Cscript unix2dos_folder.vbs <absolute_folder_path>
'
' @attention: Requires UNIX2DOS.EXE, and must be in the same folder.
'***********************************************************************

objStartFolder = Wscript.Arguments.Item(0)

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = Wscript.CreateObject("Wscript.Shell")
' Get the path of the scripts
Set objThisScript = objFso.GetFile(Wscript.ScriptFullName)
strParentFolder = objFso.GetParentFolderName(objThisScript)

' Globals
Dim strUNIX2DOS

strUNIX2DOS = objFso.BuildPath(strParentFolder, "UNIX2DOS.EXE")
Set objFolder = objFSO.GetFolder(objStartFolder)

'Set colFiles = objFolder.Files
' Process top level directory files.
For Each objFile in objFolder.Files
    strCommand = strUNIX2DOS & " """ & objFile.Path & """"
    ExecCommand(strCommand)
Next
Wscript.Echo

ProcessSubfolders objFSO.GetFolder(objStartFolder)

' Recursively processes sub-folders.
Sub ProcessSubFolders(Folder)
    For Each Subfolder in Folder.SubFolders
        Set objFolder = objFSO.GetFolder(Subfolder.Path)
        
        For Each objFile in objFolder.Files
            strCommand = strUNIX2DOS & " """ & objFile.Path & """"
            ExecCommand(strCommand)
        Next
        Wscript.Echo
        ' Recursively call ProcessSubFolders routine.
        ProcessSubFolders Subfolder
    Next
End Sub

' Executes a command (string) and echos the result
Function ExecCommand(strCommand)
    strResult = ""
'    WScript.Echo "strCommand = " & strCommand
    Set objExec = objShell.Exec(strCommand)
    Do While Not objExec.StdOut.AtEndOfStream
        strResult = strResult & objExec.StdOut.ReadLine()
    Loop
    
    Do While Not objExec.StdErr.AtEndOfStream
        strResult = strResult & objExec.StdErr.ReadLine()
    Loop
    Wscript.Echo strResult
End Function

