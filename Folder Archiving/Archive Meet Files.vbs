'Prompt for name of the archive folder
archive_folder = UserInput( "Enter archive folder (yyyy-mm-dd - name):" )

'Current year folder
dir = ""

'Base Finish Lynx Directory
LYNX_DIR = "D:\Finish Lynx\"

'Define the destination archive folder
archive_folder = LYNX_DIR & dir  & archive_folder

'Move active folder and rename
Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.MoveFolder LYNX_DIR & "Active Meet" , archive_folder

'Recreate active folder
ParentFolder = LYNX_DIR 
set objShell = CreateObject("Shell.Application")
set objFolder = objShell.NameSpace(ParentFolder) 
objFolder.NewFolder "Active Meet"

'Display archived confirmation
WScript.Echo "Archived to: " & archive_folder

Function UserInput( myPrompt )
' This function prompts the user for some input.
' When the script runs in CSCRIPT.EXE, StdIn is used,
' otherwise the VBScript InputBox( ) function is used.
' myPrompt is the the text used to prompt the user for input.
' The function returns the input typed either on StdIn or in InputBox( ).
' Written by Rob van der Woude http://www.robvanderwoude.com
    ' Check if the script runs in CSCRIPT.EXE
    If UCase( Right( WScript.FullName, 12 ) ) = "\CSCRIPT.EXE" Then
        ' If so, use StdIn and StdOut
        WScript.StdOut.Write myPrompt & " "
        UserInput = WScript.StdIn.ReadLine
    Else
        ' If not, use InputBox( )
        UserInput = InputBox( myPrompt )
    End If
End Function