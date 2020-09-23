VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This is a small class file to inplant dialog support for my API Window example
' Supports Open and Save Dialogs

Private lpHwnd As Long

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type ChooseColor
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long


Private m_CustomColors(15) As Long
Private Const CC_RGBINIT   As Long = &H1
Private Const CC_FULLOPEN  As Long = &H2
Private Const CC_ANYCOLOR  As Long = &H100
Private Const CC_NORMAL    As Long = CC_ANYCOLOR Or CC_RGBINIT
Private Const CC_EXTENDED  As Long = CC_ANYCOLOR Or CC_RGBINIT Or CC_FULLOPEN

Private OpenFileStruct As OPENFILENAME
Private lpFilter As String
Private lpFileTitle As String
Private lpDlgTitle As String
Private lpInitialDir As String
Private lpFlags As Long
Private lpFileName As String
Private lpIndex As Integer
Private CanError As Boolean
Private lphInst As Long
' color
Private m_color As Long

Public Sub ShowColor()
Dim RetCol As ChooseColor, cReturn As Long
    CanError = False
    RetCol.lStructSize = Len(RetCol) 'Size of the color type
    RetCol.hwndOwner = lpHwnd ' Parent hangle to show the color dialog
    RetCol.rgbResult = 0
    RetCol.lpCustColors = VarPtr(m_CustomColors(0))
    RetCol.flags = CC_NORMAL Or CC_EXTENDED
    
    For I = 0 To 15
        'Build a table of custom colours
        m_CustomColors(I) = RGB(I * 16, I * 16, I * 16)
    Next
    
    cReturn = ChooseColor(RetCol)
    
    If Not cReturn <> 0 Then
        CanError = -1
    Else
        m_color = RetCol.rgbResult 'return the color from the dialog
    End If
    
End Sub

Public Sub ShowSave()
Dim lpRet As Long
Dim e_Pos As Integer, sTmp As String
    ' This function is used to show the SaveFile Dialog Box
    Dim v As Variant
    
    If lpHwnd = 0 Then lpFileName = "": Exit Sub
    With OpenFileStruct
        .lStructSize = Len(OpenFileStruct) ' size of OpenFileStruct
        .hwndOwner = lpHwnd ' Parent hangle for the dialog
        .hInstance = lphInst
        .lpstrFilter = lpFilter ' Set the dialogs Filter
        .lpstrFile = Space(254) ' Set some space to hold the filename
        .nMaxFile = 255
        .lpstrFileTitle = lpFileTitle 'set the dialogs filetitle
        .nMaxFileTitle = 255
        .nFilterIndex = lpIndex
        .lpstrInitialDir = lpInitialDir
        .lpstrTitle = lpDlgTitle
        .flags = lpFlags
    End With
   
    
    lpRet = GetSaveFileName(OpenFileStruct)
    CanError = lpRet
    If lpRet <> 0 Then
        lpIndex = OpenFileStruct.nFilterIndex
        sTmp = Trim(OpenFileStruct.lpstrFile)
        lpFileName = Left(sTmp, InStr(1, sTmp, Chr(0), vbBinaryCompare) - 1)
        sTmp = ""

        e_Pos = InStrRev(lpFileName, "\", Len(lpFileName), vbBinaryCompare)
        If e_Pos = 0 Then Exit Sub
        lpFileTitle = Mid(lpFileName, e_Pos + 1, Len(lpFileName))
        Exit Sub
    Else
        lpFileTitle = ""
        lpFileName = ""
    End If
    
End Sub
Public Sub ShowOpen()
Dim lpRet As Long
Dim e_Pos As Integer, sTmp As String
    ' This function is used to show the OpenFile Dialog Box
    
    If lpHwnd = 0 Then lpFileName = "": Exit Sub
    With OpenFileStruct
        .lStructSize = Len(OpenFileStruct) ' size of OpenFileStruct
        .hwndOwner = lpHwnd ' Parent hangle for the dialog
        .hInstance = lphInst
        .lpstrFilter = lpFilter ' Set the dialogs Filter
        .lpstrFile = Space(254) ' Set some space to hold the filename
        .nMaxFile = 255
        .lpstrFileTitle = lpFileTitle 'set the dialogs filetitle
        .nMaxFileTitle = 255
        .lpstrInitialDir = lpInitialDir
        .lpstrTitle = lpDlgTitle
        .flags = lpFlags
    End With
    
    lpRet = GetOpenFileName(OpenFileStruct)
    CanError = lpRet
    
    If lpRet <> 0 Then
        lpIndex = OpenFileStruct.nFilterIndex
        sTmp = Trim(OpenFileStruct.lpstrFile)
        lpFileName = Left(sTmp, InStr(1, sTmp, Chr(0), vbBinaryCompare) - 1)
        sTmp = ""
        e_Pos = InStrRev(lpFileName, "\", Len(lpFileName), vbBinaryCompare)
        If e_Pos = 0 Then Exit Sub
        lpFileTitle = Mid(lpFileName, e_Pos + 1, Len(lpFileName))
        
        Exit Sub
    Else
    
        lpFileTitle = ""
        lpFileName = ""
    End If
    
End Sub

Public Property Get DlgHwnd() As Long
    DlgHwnd = lpHwnd
End Property

Public Property Let DlgHwnd(ByVal vHwnd As Long)
    ' Hangle of the window to add the dialog to
    lpHwnd = vHwnd
End Property

Public Property Get Filter() As String
    ' Dialogs Filter Property
    Filter = lpFilter
End Property

Public Property Let Filter(ByVal vNewFilter As String)
    ' set a new Filter for the dialog
    lpFilter = vNewFilter
End Property

Public Property Get FileTitle() As String
    FileTitle = lpFileTitle ' get the dialogs filetitle
End Property

Public Property Let FileTitle(ByVal vNewFileTitle As String)
    ' set the dialogs file title
    lpFileTitle = vNewFileTitle
End Property

Public Property Get DialogTitle() As String
    DialogTitle = lpDlgTitle ' Get the dialogs title
End Property

Public Property Let DialogTitle(ByVal vNewTitle As String)
    lpDlgTitle = vNewTitle 'set the dialogs title
End Property

Public Property Get InitialDir() As String
    'InitialDir of were the dialogs opens from
    InitialDir = lpInitialDir
End Property

Public Property Let InitialDir(ByVal vInitialDir As String)
    ' Set dialogs InitialDir
    lpInitialDir = vInitialDir
End Property

Public Property Get flags() As Long
    ' Flags for the dialog
    flags = lpFlags
End Property

Public Property Let flags(ByVal vFlags As Long)
    ' set any flags for the dialog
    lpFlags = vFlags
End Property

Public Property Get FileName() As String
    FileName = lpFileName
End Property

Public Property Let FileName(ByVal vFileName As String)
    FileName = vFileName
End Property

Public Property Get FilterIndex() As Integer
    FilterIndex = lpIndex
End Property

Public Property Let FilterIndex(ByVal vNewIndex As Integer)
    lpIndex = vNewIndex
End Property

Public Property Get CancelError() As Boolean
    CancelError = CanError
End Property

Public Property Get hInst() As Long
    hInst = lphInst
End Property

Public Property Let hInst(ByVal vhInst As Long)
    lphInst = vhInst
End Property

Public Property Get Color() As Long
    Color = m_color
End Property

