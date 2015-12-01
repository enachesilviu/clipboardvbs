Dim filesys, clipboard, folders, files

Set filesys = CreateObject("Scripting.FileSystemObject")
Set folders = CreateObject("System.Collections.ArrayList")
Set files   = CreateObject("System.Collections.ArrayList")
clipboard   = ""

' We need to sort folders and files separately.
' Otherwise they are sorted together and files could jump ahead of folders in the list.

' Grab folders
For i = 0 To (WScript.Arguments.Count - 1)
   If (filesys.FolderExists(WScript.Arguments.item(i))) Then
       folders.Add WScript.Arguments.item(i)
   End If
Next

' Sort folders
folders.Sort()

' Grab files
For i = 0 To (WScript.Arguments.Count - 1)
   If (filesys.FileExists(WScript.Arguments.item(i))) Then
       files.Add WScript.Arguments.item(i)
   End If
Next

' Sort folders
files.Sort()

' Add folder names to clipboard
For i = 0 To (folders.Count - 1)
   If (i > 0 ) Then
       clipboard = clipboard & vbCr & vbLf
   End If
   clipboard = clipboard & filesys.GetFileName(folders(i))
Next

'Check if there are both folders and files so we start the files list on new line.

If (folders.Count > 0 And files.Count > 0) Then
   clipboard = clipboard & vbCr & vbLf
End If

' Add file names to clipboard
For i = 0 To (files.Count - 1)
   If (i > 0 ) Then
       clipboard = clipboard & vbCr & vbLf
   End If
   clipboard = clipboard & filesys.GetFileName(files(i))
Next

Set WshShell = CreateObject("WScript.Shell")
Set oExec    = WshShell.Exec("clip")
Set oIn      = oExec.stdIn
oIn.WriteLine clipboard