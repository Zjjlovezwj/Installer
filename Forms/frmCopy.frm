VERSION 5.00
Begin VB.Form frmCopy 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   Enabled         =   0   'False
   Icon            =   "frmCopy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 在模块中添加以下代码
Option Explicit

' API声明
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
    
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" _
    (ByVal pszPath As String) As Long

' 常量定义
Private Const FO_COPY = &H2
Private Const FOF_SILENT = &H4
Private Const FOF_NOCONFIRMATION = &H10
Private Const FOF_NOCONFIRMMKDIR = &H200
Private Const FOF_NOERRORUI = &H400

Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Public Function CopySystemFiles() As Boolean
    On Error Resume Next
    
    Dim op As SHFILEOPSTRUCT
    Dim result As Long
    Dim sourceFiles As String
    Dim targetDir As String
    Dim configPath As String
    
    configPath = App.path & "\Setup.inf"
    
    ' 1. 构建目标目录路径
    targetDir = frmFolder.Text1.Text & GetSetting(configPath, ID_InstallLocation)
    
    ' 确保路径以反斜杠结尾
    If Right(targetDir, 1) <> "\" Then targetDir = targetDir & "\"
    
    ' 2. 确保目标目录存在
    If Not PathExists(targetDir) Then
        MkDir targetDir
    End If
    
    ' 3. 构建要复制的文件列表（多个文件用vbNullChar分隔，最后双vbNullChar结尾）
    sourceFiles = App.path & "\msvbvm60.dll" & vbNullChar & _
                 App.path & "\COMCTL32.OCX" & vbNullChar & vbNullChar
    
    ' 4. 设置文件操作结构
    With op
        .wFunc = FO_COPY
        .pFrom = sourceFiles
        .pTo = targetDir & vbNullChar & vbNullChar
        .fFlags = FOF_SILENT Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR Or FOF_NOERRORUI
    End With
    
    ' 5. 执行复制操作
    result = SHFileOperation(op)
    
    ' 6. 检查结果
    If result = 0 And op.fAnyOperationsAborted = 0 Then
        ' 验证文件是否复制成功
        If PathExists(targetDir & "msvbvm60.dll") And _
           PathExists(targetDir & "COMCTL32.OCX") Then
            CopySystemFiles = True
        Else
            CopySystemFiles = False
        End If
    Else
        CopySystemFiles = False
    End If
    
    Exit Function
    
On Error GoTo 0
End Function

' 检查文件/目录是否存在
Private Function PathExists(ByVal path As String) As Boolean
On Error Resume Next
    PathExists = (PathFileExists(path) = 1)
    On Error GoTo 0
End Function

' 使用Shell文件操作API复制文件（支持通配符）
Public Function CopyFilesWithProgress(ByVal SourcePath As String, ByVal DestPath As String) As Boolean
On Error Resume Next
    Dim op As SHFILEOPSTRUCT
    Dim strFrom As String
    Dim strTo As String
    
    ' 确保路径以双空字符结尾
    strFrom = SourcePath & "\*.*" & vbNullChar & vbNullChar
    strTo = DestPath & vbNullChar & vbNullChar
    
    With op
        .wFunc = FO_COPY
        .pFrom = strFrom
        .pTo = strTo
        .fFlags = FOF_SILENT Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR Or FOF_NOERRORUI Or &H800
    End With
    
    ' 执行复制操作
    If SHFileOperation(op) <> 0 Then
        CopyFilesWithProgress = False
    Else
        CopyFilesWithProgress = True
    End If
    Exit Function
    
On Error GoTo 0
End Function

Private Sub Form_Load()
On Error Resume Next
    ' 获取源路径和目标路径
    Dim SourcePath As String
    Dim DestPath As String
    Dim configPath As String
    
    configPath = App.path & "\Setup.inf"
    SourcePath = App.path & "\Files"
    DestPath = frmFolder.Text1.Text
    Call CopyFilesWithProgress(SourcePath, DestPath)
    
    ' 执行复制操作
    CopySelfAsUninstaller frmFolder.Text1.Text & GetSetting(configPath, ID_InstallLocation)
    CopySystemFiles
    On Error GoTo 0
End Sub
