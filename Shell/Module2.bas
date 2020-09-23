Attribute VB_Name = "Module2"

Function Shell()
'Sets Text to white - Back to black
SetConsoleTextAttribute hConsoleOut, FOREGROUND_WHITE Or FOREGROUND_INTENSITY
ConsoleWrite "Ex"
'Sets text to Red - Back to black
SetConsoleTextAttribute hConsoleOut, FOREGROUND_RED Or FOREGROUND_INTENSITY
ConsoleWrite "Shell"
ConsoleWrite vbCrLf
SHELL_LOGIN:
'###############GET USERNAME###################
'Sets Text to white - Back to black
SetConsoleTextAttribute hConsoleOut, FOREGROUND_WHITE Or FOREGROUND_INTENSITY
ConsoleWrite "UserName: "
'Sets Text to Yellow - Back to black
SetConsoleTextAttribute hConsoleOut, FOREGROUND_RED Or FOREGROUND_GREEN Or FOREGROUND_INTENSITY
szUsername = ConsoleReadLine()
If szUsername = "#quit" Then
    EndProgram
End If
'##############################################

'##############GET PASSWORD####################
'Sets Text to white - Back to black
SetConsoleTextAttribute hConsoleOut, FOREGROUND_WHITE Or FOREGROUND_INTENSITY
ConsoleWrite "Password: "
'Sets Text to Yellow - Back to black
SetConsoleTextAttribute hConsoleOut, FOREGROUND_RED Or FOREGROUND_GREEN Or FOREGROUND_INTENSITY
szPassword = ConsoleReadLine()
'Sets Text to white - Back to black
SetConsoleTextAttribute hConsoleOut, FOREGROUND_WHITE Or FOREGROUND_INTENSITY
'##############################################
Select Case szUsername
Case Is = "root"
    If szPassword = "r00t" Then
        
    Else
        'Sets Text to red - Back to black
        SetConsoleTextAttribute hConsoleOut, FOREGROUND_RED Or FOREGROUND_INTENSITY
        ConsoleWriteLine "Invalid Password."
        'Sets Text to white - Back to black
        SetConsoleTextAttribute hConsoleOut, FOREGROUND_WHITE Or FOREGROUND_INTENSITY
        GoTo SHELL_LOGIN
    End If
Case Is = "Ex"
    If szPassword = "r00t" Then
        
    Else
        'Sets Text to red - Back to black
        SetConsoleTextAttribute hConsoleOut, FOREGROUND_RED Or FOREGROUND_INTENSITY
        ConsoleWriteLine "Invalid Password."
        'Sets Text to white - Back to black
        SetConsoleTextAttribute hConsoleOut, FOREGROUND_WHITE Or FOREGROUND_INTENSITY
        GoTo SHELL_LOGIN
    End If
Case Else
    'Sets Text to red - Back to black
        SetConsoleTextAttribute hConsoleOut, FOREGROUND_RED Or FOREGROUND_INTENSITY
        ConsoleWriteLine "Invalid Username."
        'Sets Text to white - Back to black
        SetConsoleTextAttribute hConsoleOut, FOREGROUND_WHITE Or FOREGROUND_INTENSITY
        GoTo SHELL_LOGIN
End Select
ConsoleWriteLine "You are now logged in as " & szUsername & "."
prompt
End Function
