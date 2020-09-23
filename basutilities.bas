Attribute VB_Name = "basUtilities"
Option Explicit

Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer

Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Const HEX_BMP_KEY As String = "424D"
Public Const HEX_GIF_KEY As String = "4749"
Public Const HEX_JPG_KEY As String = "4A464946"
Public Const HEX_BYTE_SIZE As Long = 9
Public Const BLOCKSIZE = 32768



Sub CenterForm(frmForm As Form)
On Error GoTo CenterFormError

frmForm.Left = Screen.Width / 2 - frmForm.Width / 2
frmForm.Top = Screen.Height / 2 - frmForm.Height / 2

CenterFormContinue:
    Exit Sub
CenterFormError:
    MsgBox Error$, vbExclamation
    Resume CenterFormContinue
End Sub


Function GetFileFromPath(sFilePath) As String
On Error GoTo GetFileFromPathError

Dim sFileTitle As String * 1024
Dim lRetVal As Long

lRetVal = GetFileTitle(sFilePath, sFileTitle, Len(sFileTitle))
GetFileFromPath = Mid(sFileTitle, 1, InStr(sFileTitle, Chr(0)) - 1)

GetFileFromPathContinue:
    Exit Function
GetFileFromPathError:
    MsgBox Error$, vbExclamation
    Resume GetFileFromPathContinue
End Function

Function SystemLogonName() As Variant
On Error GoTo SystemLogonNameError

Dim sUserName As String
Dim lRetVal As Long

sUserName = String(2048, 32)
lRetVal = GetUserName(sUserName, Len(sUserName) - 1)
'See if there is no one logged in.
If InStr(sUserName, Chr(0)) > 0 Then
    SystemLogonName = Mid(sUserName, 1, InStr(sUserName, Chr(0)) - 1)
Else
    SystemLogonName = "Unknown"
End If

SystemLogonNameContinue:
    Exit Function
SystemLogonNameError:
    MsgBox Error$, vbExclamation
    Resume SystemLogonNameContinue
End Function

