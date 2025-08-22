Attribute VB_Name = "comCopy"
Option Explicit

' API声明
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" ( _
    ByVal hModule As Long, _
    ByVal lpFileName As String, _
    ByVal nSize As Long) As Long

Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" ( _
    ByVal lpExistingFileName As String, _
    ByVal lpNewFileName As String, _
    ByVal bFailIfExists As Long) As Long

Private Declare Function Sleep Lib "kernel32" ( _
    ByVal dwMilliseconds As Long) As Long

' 复制程序自身和配置文件
Public Sub CopySelfAsUninstaller(ByVal sInstallDir As String)
    On Error Resume Next
    
    Dim sExePath As String, sNewPath As String
    Dim sInfPath As String, sNewInfPath As String
    
    ' 确保目录以反斜杠结尾
    If Right(sInstallDir, 1) <> "\" Then sInstallDir = sInstallDir & "\"
    
    ' 创建目标目录（如果不存在）
    If Dir(sInstallDir, vbDirectory) = "" Then
        MkDir sInstallDir
    End If
    
    ' 获取当前程序路径
    sExePath = String(260, 0)
    GetModuleFileName 0, sExePath, 260
    sExePath = Left(sExePath, InStr(sExePath, vbNullChar) - 1)
    
    ' 设置新路径
    sNewPath = sInstallDir & "\Uninstall.exe"
    
    ' 复制程序文件（重命名为Uninstall.exe）
    If FileExists(sNewPath) Then
        ' 如果已存在，先尝试删除
        SetAttr sNewPath, vbNormal
        Kill sNewPath
        Sleep 500 ' 等待系统释放文件
    End If
    
    CopyFile sExePath, sNewPath, 0
    
    ' 复制Setup.inf文件（假设与程序同目录）
    sInfPath = App.Path & "\Setup.inf"
    sNewInfPath = sInstallDir & "Setup.inf"
    
    If FileExists(sInfPath) Then
        If FileExists(sNewInfPath) Then
            SetAttr sNewInfPath, vbNormal
            Kill sNewInfPath
        End If
        CopyFile sInfPath, sNewInfPath, 0
    End If
    On Error GoTo 0
End Sub

' 检查文件是否存在
Private Function FileExists(ByVal sFile As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(sFile) <> "")
    On Error GoTo 0
End Function

