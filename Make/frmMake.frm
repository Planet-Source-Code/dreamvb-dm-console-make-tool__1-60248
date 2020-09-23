VERSION 5.00
Begin VB.Form frmMake 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DM Console Make"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5955
   Icon            =   "frmMake.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdmake 
      Caption         =   "&Make"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1725
      TabIndex        =   6
      Top             =   2610
      Width           =   1215
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   375
      Left            =   3135
      TabIndex        =   7
      Top             =   2610
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   4515
      TabIndex        =   8
      Top             =   2610
      Width           =   1215
   End
   Begin VB.CheckBox chkBackup 
      Caption         =   "Create Backup"
      Height          =   195
      Left            =   270
      TabIndex        =   3
      Top             =   1095
      Width           =   5070
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Win32 Console Application"
      Height          =   225
      Left            =   255
      TabIndex        =   5
      Top             =   1995
      Value           =   -1  'True
      Width           =   2955
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Win32 GUI Application"
      Height          =   225
      Left            =   255
      TabIndex        =   4
      Top             =   1710
      Width           =   2955
   End
   Begin VB.CommandButton cmdopen 
      Caption         =   "...."
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   585
      Width           =   465
   End
   Begin VB.TextBox txtInput 
      Height          =   375
      Left            =   225
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   570
      Width           =   4605
   End
   Begin VB.Line Line1 
      X1              =   225
      X2              =   5775
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Label lblA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Switch To:"
      Height          =   195
      Left            =   300
      TabIndex        =   0
      Top             =   1395
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Input application:"
      Height          =   195
      Left            =   255
      TabIndex        =   9
      Top             =   300
      Width           =   1710
   End
End
Attribute VB_Name = "frmMake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Switch_Type As Integer

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" _
(ByVal lpFileName As String, _
ByVal dwDesiredAccess As Long, _
ByVal dwShareMode As Long, ByVal _
lpSecurityAttributes As Long, _
ByVal dwCreationDisposition As Long, _
ByVal dwFlagsAndAttributes As Long, _
ByVal hTemplateFile As Long) As Long

Private Declare Function SetFilePointer Lib "kernel32" (ByVal FilePtr As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function ReadFile Lib "kernel32" (ByVal FilePtr As Long, lplpBuffer As Any, _
ByVal nNumberOfBytesToRead As Long, _
lpNumberOfBytesRead As Long, _
ByVal lpOverlapped As Long) As Long

Private Declare Function WriteFile Lib "kernel32" (ByVal FilePtr As Long, lplpBuffer As Any, _
ByVal nNumberOfBytesToWrite As Long, _
lpNumberOfBytesWritten As Long, _
ByVal lpOverlapped As Long) As Long

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const flags = GENERIC_READ Or GENERIC_WRITE

Sub Abort(code As Integer, fp As Long)
    'error messages
    If fp <> 0 Then CloseHandle fp ' close file
    
    Select Case code
        Case 0: MsgBox "Unable to open file.", vbCritical, frmMake.Caption
        Case 1: MsgBox "Unable to read file.", vbCritical, frmMake.Caption
        Case 2: MsgBox "Unable to detect executable format.", vbCritical, frmMake.Caption
        Case 3: MsgBox "Unable detect a vaild executable signature.", vbCritical, frmMake.Caption
        Case 4: MsgBox "Unable to write to file.", vbCritical, frmMake.Caption
    End Select
    
End Sub

Function SwitFilePtrType(lpFile As String) As Integer
Dim lpBuffer(0 To 1024) As Byte
Dim Offset As Long
Dim FilePtr As Long

    SwitFilePtrType = 0
    
    FilePtr = CreateFile(lpFile, flags, 0, 0, 3, 0, 0) 'Get the pointer to the file
    
    If FilePtr = -1 Then Abort 0, FilePtr: Exit Function
    
    ' Read in the file into lpBuffer
    If ReadFile(FilePtr, lpBuffer(0), 64, 0, 0) = 0 Then
        ' Error found so we need to stop and clean up
        Abort 1, FilePtr
        Erase lpBuffer
        Exit Function
    End If
    
    If Not lpBuffer(0) And lpBuffer(1) = 72 Then
        ' Error found so we need to stop and clean up
        Abort 2, FilePtr
        Erase lpBuffer
        Exit Function
    End If
    
    CopyMemory Offset, lpBuffer(60), 4  'Move offset from lpBuffer to Offset
    
    If Offset = 0 Then
        ' Error found so we need to stop and clean up
        Abort 2, FilePtr
        Erase lpBuffer
        Exit Function
    Else
        SetFilePointer FilePtr, Offset, 0, 0  'Move into the file upon the offset
        
        If Not ReadFile(FilePtr, lpBuffer(0), 2, 0, 0) = 1 Then Abort 1, FilePtr: Exit Function
        
        If Not (lpBuffer(0) And lpBuffer(1)) = 64 Then ' check for a vaild file Sig
            ' Error found so we need to stop and clean up
            Abort 3, FilePtr
            Erase lpBuffer
            Exit Function
        Else
            SetFilePointer FilePtr, Offset + 92, 0, 0 'Move into the file upon the offset + 92
            lpBuffer(0) = Switch_Type ' Here we add the new appliaction type
            If WriteFile(FilePtr, lpBuffer(0), 1, 0, 0) <> 1 Then ' Update the file
                ' Error found so we need to stop and clean up
                Abort 4, FilePtr
                Erase lpBuffer
                Exit Function
            Else
                'Clean up and return good result of 1
                SwitFilePtrType = 1
                Offset = 0
                Erase lpBuffer
                CloseHandle FilePtr 'Close the file
            End If
        End If
    End If
    
End Function

Function GetFileExt(lpFileName As String) As String
Dim e_Pos As Integer
    ' functions that returns a file's ext
    ' ex GetFileExt(c:\this.txt") return txt
    e_Pos = InStrRev(lpFileName, ".", Len(lpFileName), vbBinaryCompare)
    If e_Pos <> 0 Then
        GetFileExt = LCase(Mid(lpFileName, e_Pos + 1, Len(lpFileName)))
    End If
End Function

Sub RemoveFileExt(lpFileName As String)
Dim e_Pos As Integer
    ' returns a file's ext
    'ex RemoveFileExt("C:\this.txt") returns c:\this.
    e_Pos = InStrRev(lpFileName, ".", Len(lpFileName), vbBinaryCompare)
    If e_Pos <> 0 Then lpFileName = Mid(lpFileName, 1, e_Pos)
End Sub

Private Sub cmdAbout_Click()
    MsgBox "DM Console Make Tool for VB6", vbInformation, frmMake.Caption
End Sub

Private Sub cmdExit_Click()
    Unload frmMake
End Sub

Private Sub cmdmake_Click()
Dim lzBackup As String, lzFile As String
    
    lzFile = "": lzBackup = ""
    
    lzFile = txtInput.Text
    If chkBackup Then ' check for backup option
        lzBackup = lzFile
        RemoveFileExt lzBackup 'Remove file ext
        lzBackup = lzBackup & "bak" 'New Backup filename
        FileCopy lzFile, lzBackup
        If SwitFilePtrType(lzFile) <> 1 Then Exit Sub
        MsgBox "The Appliaction has been changed and saved to:" _
        & vbCrLf & lzFile & vbCrLf & vbCrLf & "A backup of the file was saved to:" _
        & vbCrLf & lzBackup, vbInformation, frmMake.Caption
        lzFile = "": lzBackup = ""
        cmdExit_Click
        Exit Sub
    Else
        ' no backup needed so we just switch the app around
        If SwitFilePtrType(lzFile) <> 1 Then
            lzFile = ""
            cmdExit_Click
            Exit Sub
        Else
            MsgBox "The Appliaction has been changed and saved to:" _
            & vbCrLf & lzFile & vbCrLf, vbInformation, frmMake.Caption
            lzFile = ""
            cmdExit_Click
        End If
    End If
End Sub

Private Sub cmdopen_Click()
Dim myDlg As CDialog

    Set myDlg = New CDialog
    myDlg.DlgHwnd = frmMake.hWnd
    myDlg.hInst = App.hInstance
    myDlg.DialogTitle = "Open Program"
    myDlg.Filter = "Program Files(*.exe)" + Chr$(0) + "*.exe"
    myDlg.flags = 0
    myDlg.ShowOpen
    
    If myDlg.CancelError = False Then Exit Sub
    If GetFileExt(myDlg.FileName) <> "exe" Then
        txtInput.Text = ""
        cmdmake.Enabled = False
        MsgBox "This is not a vaild Win32 Appliaction.", vbInformation, frmMake.Caption
        Exit Sub
    Else
        txtInput.Text = myDlg.FileName
        txtInput.Enabled = True
        cmdmake.Enabled = True
    End If
    
End Sub

Private Sub Form_Load()
    Option2_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call Form_Unload(Cancel)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMake = Nothing
    End
End Sub

Private Sub Option1_Click()
    Switch_Type = 2 'GUI App
End Sub

Private Sub Option2_Click()
    Switch_Type = 3 'Console App
End Sub
