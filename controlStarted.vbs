Option explicit

On Error Resume Next

Dim WshShell, objWMIService, objProcess, colProcesses
Dim sScriptPath, isStarted, state

Set WshShell = WScript.CreateObject("WScript.Shell")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}")
sScriptPath  = Replace(WScript.ScriptFullName, WScript.ScriptName, "")

isStarted = False

Set colProcesses = nothing
Set colProcesses = objWMIService.ExecQuery("Select * From Win32_Process")
For Each objProcess in colProcesses
	if not IsNull(objProcess.CommandLine) then
		if InStr(1, objProcess.CommandLine, "AvatarWebAPIClient.exe") <> 0 then
			isStarted = True
		end if
	end if
Next

if isStarted = False then
	state = WshShell.Run(sScriptPath & "AvatarWebAPIClient.exe", 1, False)
end if

'-- Destroy objects
Set colProcesses = nothing
Set objWMIService = nothing
Set WshShell = nothing

WScript.Quit(0)