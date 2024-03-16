' CoPilot: This script checks the current system theme and switches it to the opposite one
' It uses the registry key Personalize\AppsUseLightTheme to determine the theme
' 0 means dark mode, 1 means light mode
' Source: [3](https://github.com/aulas4you/how-turn-on-dark-mode-windows-vbs-script)

Dim WshShell, regKey, theme
Set WshShell = CreateObject("WScript.Shell")
regKey = "HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme"
theme = WshShell.RegRead(regKey) ' Read the current theme value
If theme = 0 Then ' If dark mode, switch to light mode
    WshShell.RegWrite regKey, 1, "REG_DWORD"
Else ' If light mode, switch to dark mode
    WshShell.RegWrite regKey, 0, "REG_DWORD"
End If
'ModAM
regKey = "HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\SystemUsesLightTheme"
'theme = WshShell.RegRead(regKey) ' Read the current theme value
If theme = 0 Then ' If dark mode, switch to light mode
    WshShell.RegWrite regKey, 1, "REG_DWORD"
Else ' If light mode, switch to dark mode
    WshShell.RegWrite regKey, 0, "REG_DWORD"
End If
'RefreshDesktop
KillProc "Explorer.exe"
'Set objShell = CreateObject("Wscript.Shell") 
'objShell.Run "explorer.exe" 
'Set objShell = Nothing
'x=msgbox("Your Text Here" ,0, "Your Title Here")


Sub KillProc( myProcess )
 'Authors: Denis St-Pierre and Rob van der Woude
 'Purpose: Kills a process and waits until it is truly dead

     Dim blnRunning, colProcesses, objProcess
     blnRunning = False

     Set colProcesses = GetObject( _
                        "winmgmts:{impersonationLevel=impersonate}" _
                        ).ExecQuery( "Select * From Win32_Process", , 48 )
     For Each objProcess in colProcesses
         If LCase( myProcess ) = LCase( objProcess.Name ) Then
             ' Confirm that the process was actually running
             blnRunning = True
             ' Get exact case for the actual process name
             myProcess  = objProcess.Name
             ' Kill all instances of the process
             objProcess.Terminate()
         End If
     Next

     'If blnRunning Then
         ' Wait and make sure the process is terminated.
         ' Routine written by Denis St-Pierre.
      '   Do Until Not blnRunning
       '      Set colProcesses = GetObject( _
       '                         "winmgmts:{impersonationLevel=impersonate}" _
        '                        ).ExecQuery( "Select * From Win32_Process Where Name = '" _
       '                       & myProcess & "'" )
       '      WScript.Sleep 100 'Wait for 100 MilliSeconds
       '      If colProcesses.Count = 0 Then 'If no more processes are running, exit loop
      '           blnRunning = False
     '        End If
    '     Loop
         ' Display a message
         ' WScript.Echo myProcess & " was terminated"
  '   Else
    '     WScript.Echo "Process """ & myProcess & """ not found"
   '  End If
 End Sub