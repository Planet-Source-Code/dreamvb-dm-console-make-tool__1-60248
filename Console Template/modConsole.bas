Attribute VB_Name = "WinMain"
Dim Console As New ClsConsole
' Use this module as your console project template

Public Sub Main()
Dim sInput As String, i As Integer

    'First we check that we can use our new console
    If Console.cInit = 0 Then
        MsgBox "Unable to start console session.", vbExclamation, "error"
        Exit Sub
    End If
    
    Console.cBeep 'Make the console beep
    Console.Title = "My Console App" 'add new console title
    'Next we write a line to the console note this also adds vbcrlf
    ' if you note want this simply use Console.cWrite
    Console.TextAttribute = FOREGROUND_RED + FOREGROUND_GREEN + FOREGROUND_INTENSITY 'Yellow
    Console.cWriteLine "Hello World: This is a console app"
    
    Console.cSetCursorPosition 5, 5
    Console.cWrite "What is your name: "
    sInput = Console.ReadLine 'Read the current line
    Console.cWriteLine ' this just add's a blank line
    Console.TextAttribute = FOREGROUND_RED + FOREGROUND_GREEN + FOREGROUND_BLUE + FOREGROUND_INTENSITY ' restore text color to normal
    Console.cWriteLine "Hello " & sInput & " Pleased to meet you"
    
    'umm silly little look to fill the screen up
    For i = 0 To 10
        Console.cWriteLine ("Value of i is " & Str(i))
    Next
    
    Console.cWriteLine "Press a key to clear the screen"
    Console.ReadLine
    Console.cCls 'Clear the screen
    Console.cPuase ' This just makes the console pauses untill a key is pressed
    Console.cFree 'Close the console
    
    Set Console = Nothing 'destory the console object
    
End Sub


