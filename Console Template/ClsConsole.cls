VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'API used for reading and writeing to the console
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
'API used for console
Private Declare Function SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
Private Declare Function FreeConsole Lib "kernel32" () As Long
Private Declare Function FillConsoleOutputCharacter Lib "kernel32.dll" Alias "FillConsoleOutputCharacterA" (ByVal hConsoleOutput As Long, ByVal cCharacter As Byte, ByVal nLength As Long, dwWriteCoord As Long, lpNumberOfCharsWritten As Long) As Long

Private Declare Function GetConsoleScreenBufferInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) As Long
Private Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function SetConsoleCursorPosition Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal dwCursorPosition As Long) As Long

'Console hangle consts
Private Const STD_INPUT_HANDLE = -10&
Private Const STD_OUTPUT_HANDLE = -11&

Private Type COORD
    x As Integer
    y As Integer
End Type

Private Type SMALL_RECT
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
End Type

Private Type CONSOLE_SCREEN_BUFFER_INFO
    dwSize As COORD
    dwCursorPosition As COORD
    wAttributes As Integer
    srWindow As SMALL_RECT
    dwMaximumWindowSize As COORD
End Type

Enum TextAttr
    'forecolors
    FOREGROUND_RED = &H4
    FOREGROUND_GREEN = &H2
    FOREGROUND_BLUE = &H1
    FOREGROUND_INTENSITY = &H8
    'backcolors
    BACKGROUND_RED = &H40
    BACKGROUND_GREEN = &H20
    BACKGROUND_BLUE = &H10
    BACKGROUND_INTENSITY = &H80
End Enum

Private hOutput As Long
Private hInput As Long

Private m_ConsoleTitle As String ' used to hold console title
Private m_ConsoleTextAttr As Long
Private cBuff As CONSOLE_SCREEN_BUFFER_INFO

Private Function GetConsoleBuffInfo() As Boolean
    ' Get Console screen buffer info
    GetConsoleBuffInfo = False
    GetConsoleBuffInfo = GetConsoleScreenBufferInfo(hOutput, cBuff) <> 0
End Function

Public Sub cSetCursorPosition(x As Integer, y As Integer)
Dim mCoord As Long
Dim mCoordPrt As Long
    'Sets the position of the cursor on the console
    mCoordPrt& = VarPtr(mCoord) 'Get the address of mCoord
    CopyMemory mCoord, x, 2
    CopyMemory ByVal mCoordPrt& + 2, y, 2
    SetConsoleCursorPosition hOutput, mCoord
End Sub

Public Sub cCls()
Dim m_Width As Long, m_Height As Long
    'Clears the console. by filling in the console area with the space char
    If GetConsoleBuffInfo Then
        m_Width = (cBuff.srWindow.Right - cBuff.srWindow.Left + 1)
        m_Height = (cBuff.srWindow.Bottom - cBuff.srWindow.Top + 1)
        FillConsoleOutputCharacter hOutput, 32, m_Width * m_Height, ByVal 0, 0
    End If
    
    m_Width = 0: m_Height = 0
    
End Sub

Public Sub cWrite(lpStr As String)
    ' Write a new stream of text to the console
    WriteFile hOutput, ByVal lpStr, Len(lpStr), 0, 0
End Sub

Public Sub cWriteLine(Optional lpStr As String = vbCrLf)
    ' Write a new stream of text to the console with a line break
    cWrite lpStr & vbCrLf
End Sub

Public Sub cBeep()
    'Send the bell excape key to make the console beep
    cWrite Chr(7)
End Sub

Public Sub cFree()
    'Free the current console session
    FreeConsole
End Sub

Public Sub cPuase()
   cWriteLine "Press any key to continue . . ."
   ReadLine
End Sub

Public Function ReadLine() As String
On Error Resume Next
    Dim sText As String * 256
    'Reads input from the console
    ReadFile hInput, ByVal sText, Len(sText), vbNull, vbNull
    ReadLine = Left(sText, InStr(1, sText, Chr(0)) - 3)
    sText = ""
End Function

Public Function cInit() As Long
    ' Get the hangles to the console
    hInput = GetStdHandle(STD_INPUT_HANDLE)
    hOutput = GetStdHandle(STD_OUTPUT_HANDLE)
    cInit = (hInput And hOutput)
End Function

Public Property Get Title() As String
    'Return the console title
    Title = m_ConsoleTitle
End Property

Public Property Let Title(ByVal vNewTitle As String)
    'Set new console title
    m_ConsoleTitle = vNewTitle
    SetConsoleTitle m_ConsoleTitle
End Property

Public Property Let TextAttribute(ByVal vNewTxtAttr As TextAttr)
    ' used to set the TextAttribute of the console see TextAttr enum
    m_ConsoleTextAttr = vNewTxtAttr
    SetConsoleTextAttribute hOutput, vNewTxtAttr
End Property
