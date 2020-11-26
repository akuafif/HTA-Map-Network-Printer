'---------------------------------------------------------------------------	
' A script to create shortcut of the file based on the location of
' the map printer folder located
'---------------------------------------------------------------------------
createLoginFile
createMapFile

'---------------------------------------------------------------------------
' createLoginFile
'---------------------------------------------------------------------------
Sub createLoginFile
	Dim fso, link, objShell
	
	Set objShell = CreateObject("WScript.Shell")
	
	Set link = objShell.CreateShortcut("C:\Users\Public\Desktop\Login to Print Server.lnk")
	link.WindowStyle = 3
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	If (fso.FileExists(fso.GetParentFolderName(Wscript.ScriptFullName) + "\login.ico")) Then
		link.IconLocation = fso.GetParentFolderName(Wscript.ScriptFullName) + "\login.ico,0"
	End If

	link.TargetPath = fso.GetParentFolderName(Wscript.ScriptFullName) + "\Login_to_Print_Server.hta"
	link.WorkingDirectory = fso.GetParentFolderName(Wscript.ScriptFullName) 

	link.Save
end sub

'---------------------------------------------------------------------------
' createMapFile
'---------------------------------------------------------------------------
Sub createMapFile
	Dim fso, link, objShell
	
	Set objShell = CreateObject("WScript.Shell")

	Set link = objShell.CreateShortcut("C:\Users\Public\Desktop\Map Printer.lnk")
	link.WindowStyle = 3
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	If (fso.FileExists(fso.GetParentFolderName(Wscript.ScriptFullName) + "\map.ico")) Then
		link.IconLocation = fso.GetParentFolderName(Wscript.ScriptFullName) + "\map.ico,0"
	End If

	link.TargetPath = fso.GetParentFolderName(Wscript.ScriptFullName) + "\Map_Printer.hta"
	link.WorkingDirectory = fso.GetParentFolderName(Wscript.ScriptFullName)

	link.Save
end sub

