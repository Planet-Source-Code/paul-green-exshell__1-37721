Attribute VB_Name = "Module1"
Global Const FOREGROUND_BLUE = &H1
Global Const FOREGROUND_GREEN = &H2
Global Const FOREGROUND_RED = &H4
Global Const FOREGROUND_WHITE = &H7
Global Const BACKGROUND_BLUE = &H10
Global Const BACKGROUND_GREEN = &H20
Global Const BACKGROUND_RED = &H40
Global Const BACKGROUND_INTENSITY = &H80&
Global Const BACKGROUND_SEARCH = &H20&
Global Const FOREGROUND_INTENSITY = &H8&
Global Const FOREGROUND_SEARCH = (&H10&)
Global Const ENABLE_LINE_INPUT = &H2&
Global Const ENABLE_ECHO_INPUT = &H4&
Global Const ENABLE_MOUSE_INPUT = &H10&
Global Const ENABLE_PROCESSED_INPUT = &H1&
Global Const ENABLE_WINDOW_INPUT = &H8&
Global Const ENABLE_PROCESSED_OUTPUT = &H1&
Global Const ENABLE_WRAP_AT_EOL_OUTPUT = &H2&
Global Const STD_OUTPUT_HANDLE = -11&
Global Const STD_INPUT_HANDLE = -10&
Global Const STD_ERROR_HANDLE = -12&
Global Const INVALID_HANDLE_VALUE = -1&
Declare Function AllocConsole Lib "kernel32" () As Long
Declare Function FreeConsole Lib "kernel32" () As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Declare Function ReadConsole Lib "kernel32" Alias "ReadConsoleA" (ByVal hConsoleInput As Long, ByVal lpBuffer As String, ByVal nNumberOfCharsToRead As Long, lpNumberOfCharsRead As Long, lpReserved As Any) As Long
Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
Declare Function SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
Global hConsoleOut As Long, hConsoleIn As Long, hConsoleErr As Long
Global szCommand As String, szUsername As String, szPassword As String

Sub KillConsole()
    CloseHandle hConsoleOut
    CloseHandle hConsoleIn
    FreeConsole
End Sub
Sub ConsoleWriteLine(sInput As String)
     ConsoleWrite sInput + vbCrLf
End Sub
Sub ConsoleWrite(sInput As String)
     Dim cWritten As Long
     WriteConsole hConsoleOut, ByVal sInput, Len(sInput), cWritten, ByVal 0&
End Sub
Function ConsoleReadLine() As String
    Dim ZeroPos As Long
    'Create a buffer
    ConsoleReadLine = String(500, 0)
    'Read the input
    ReadConsole hConsoleIn, ConsoleReadLine, Len(ConsoleReadLine), vbNull, vbNull
    'Strip off trailing vbCrLf and Chr$(0)'s
    ZeroPos = InStr(ConsoleReadLine, Chr$(0))
    If ZeroPos > 0 Then ConsoleReadLine = Left$(ConsoleReadLine, ZeroPos - 3)
End Function

Sub main()
NewConsole
SetConsoleTitle "ExShell"
Shell
End Sub

Sub NewConsole()
    If AllocConsole() Then
        hConsoleOut = GetStdHandle(STD_OUTPUT_HANDLE)
        If hConsoleOut = INVALID_HANDLE_VALUE Then MsgBox "Unable to get STDOUT"
        hConsoleIn = GetStdHandle(STD_INPUT_HANDLE)
        If hConsoleOut = INVALID_HANDLE_VALUE Then MsgBox "Unable to get STDIN"
    Else
        MsgBox "Couldn't allocate console"
    End If
End Sub

Sub EndProgram()
KillConsole
End
End Sub
