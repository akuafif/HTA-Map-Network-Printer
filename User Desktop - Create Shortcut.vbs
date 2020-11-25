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
	Dim fso, link, objShell, curDir, DesktopPath
	
	Set objShell = CreateObject("WScript.Shell")
	DesktopPath = objShell.SpecialFolders("Desktop")
	
	Set link = objShell.CreateShortcut(DesktopPath & "\Login to Print Server.lnk")
	link.WindowStyle = 3
	
	Set fso = CreateObject("Scripting.FileSystemObject")

	
	curDir = fso.GetParentFolderName(Wscript.ScriptFullName) 

	If (fso.FileExists(curDir + "\login.ico")) Then
		link.IconLocation = curDir + "\login.ico,0"
	End If

	link.TargetPath = curDir + "\Login_to_Print_Server.hta"
	link.WorkingDirectory = curDir

	link.Save
end sub

'---------------------------------------------------------------------------
' createMapFile
'---------------------------------------------------------------------------
Sub createMapFile
	Dim fso, link, objShell, curDir, DesktopPath
	
	Set objShell = CreateObject("WScript.Shell")
	DesktopPath = objShell.SpecialFolders("Desktop")
	
	Set link = objShell.CreateShortcut(DesktopPath & "\Map Printer.lnk")
	link.WindowStyle = 3
	
	Set fso = CreateObject("Scripting.FileSystemObject")

	curDir = fso.GetParentFolderName(Wscript.ScriptFullName) 

	If (fso.FileExists(curDir + "\map.ico")) Then
		link.IconLocation = curDir + "\map.ico,0"
	End If

	link.TargetPath = curDir + "\Map_Printer.hta"
	link.WorkingDirectory = curDir

	link.Save
end sub

