Attribute VB_Name = "comGetPath"
Option Explicit

' API声明
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" ( _
    ByVal hwndOwner As Long, _
    ByVal nFolder As Long, _
    pidl As Long _
) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" ( _
    ByVal pidl As Long, _
    ByVal pszPath As String _
) As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

' 常量
Private Const CSIDL_PROGRAM_FILES = &H26     ' Program Files
Private Const MAX_PATH = 260

' 获取特殊文件夹路径
Public Function GetSpecialFolderPath(ByVal nFolder As Long) As String
On Error Resume Next
    Dim pidl As Long
    Dim sPath As String * MAX_PATH
    
    If SHGetSpecialFolderLocation(0, nFolder, pidl) = 0 Then
        If SHGetPathFromIDList(pidl, sPath) <> 0 Then
            GetSpecialFolderPath = Left$(sPath, InStr(sPath, vbNullChar) - 1)
        End If
        CoTaskMemFree pidl
    End If
    On Error GoTo 0
End Function

