Attribute VB_Name = "Module3"
Sub prompt()
'Sets Text to white - Back to black
SetConsoleTextAttribute hConsoleOut, FOREGROUND_WHITE Or FOREGROUND_INTENSITY
ConsoleWrite "[/] $ "
'Sets Text to Yellow - Back to black
SetConsoleTextAttribute hConsoleOut, FOREGROUND_RED Or FOREGROUND_GREEN Or FOREGROUND_INTENSITY
szCommand = ConsoleReadLine()
'Sets Text to white - Back to black
SetConsoleTextAttribute hConsoleOut, FOREGROUND_WHITE Or FOREGROUND_INTENSITY
LCase (szCommand)
Select Case szCommand
Case Is = "ver"
    ConsoleWrite "-Ex"
    'Sets Text to red - Back to black
    SetConsoleTextAttribute hConsoleOut, FOREGROUND_RED Or FOREGROUND_INTENSITY
    ConsoleWrite "Shell"
    'Sets Text to white - Back to black
    SetConsoleTextAttribute hConsoleOut, FOREGROUND_WHITE Or FOREGROUND_INTENSITY
    ConsoleWrite "- ver 1.0.2" & vbCrLf
    prompt
Case Is = "help"
    ConsoleWriteLine vbCrLf & "ExShell Help:"
    ConsoleWriteLine "Commands:"
    ConsoleWriteLine "help - Displays this Screen"
    ConsoleWriteLine "time - Displays Time"
    ConsoleWriteLine "date - Displays Date"
    ConsoleWriteLine "whoami - Displays Current User Logged in"
    ConsoleWriteLine "login - Change user logged in"
    ConsoleWriteLine "logout - Logout to Main Screen"
    ConsoleWriteLine "exit - Closes shell" & vbCrLf
    prompt
Case Is = "time"
    ConsoleWriteLine vbCrLf & "It is Currently " & Time & "." & vbCrLf
    prompt
Case Is = "date"
    ConsoleWriteLine vbCrLf & "Today is " & Date & "." & vbCrLf
    prompt
Case Is = "whoami"
    ConsoleWriteLine vbCrLf & "You are Currently Logged in as " & szUsername & "." & vbCrLf
    prompt
Case Is = "login"
    ConsoleWriteLine vbCrLf
    login
Case Is = "logout"
    ConsoleWriteLine vbCrLf
    szUsername = ""
    szPassword = ""
    Shell
Case Is = "exit"
    EndProgram
Case Else
    ConsoleWriteLine "Bad Command." & vbCrLf
    prompt
End Select
End Sub
