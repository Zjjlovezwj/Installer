Attribute VB_Name = "comBrowseForFolder"
Option Explicit

'--- API 声明 ---
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

'--- 文件夹选择对话框结构 ---
Private Type BROWSEINFO
    hOwner As Long          ' 父窗口句柄
    pidlRoot As Long        ' 根目录（通常设为0表示桌面）
    pszDisplayName As String ' 返回的文件夹名称
    lpszTitle As String     ' 对话框标题
    ulFlags As Long         ' 对话框选项
    lpfn As Long            ' 回调函数（通常设为0）
    lParam As Long          ' 回调参数（通常设为0）
    iImage As Long          ' 系统图标索引（通常不用）
End Type

'--- 常量 ---
Private Const BIF_RETURNONLYFSDIRS = &H1  ' 仅返回文件系统目录
Private Const MAX_PATH = 260              ' 最大路径长度

'--- 函数：显示文件夹选择对话框 ---
Public Function BrowseForFolder(ByVal hwndOwner As Long, ByVal sTitle As String) As String
    Dim bi As BROWSEINFO
    Dim pidl As Long
    Dim sPath As String
    Dim pos As Integer
    On Error Resume Next

    ' 设置对话框属性
    With bi
        .hOwner = hwndOwner
        .lpszTitle = sTitle
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    ' 显示文件夹选择对话框
    pidl = SHBrowseForFolder(bi)
    If pidl Then
        sPath = String$(MAX_PATH, 0)
        If SHGetPathFromIDList(pidl, sPath) Then
            ' 去除字符串末尾的空字符
            pos = InStr(sPath, vbNullChar)
            If pos Then sPath = Left$(sPath, pos - 1)
        End If
        CoTaskMemFree pidl  ' 释放内存
    End If

    BrowseForFolder = sPath
    On Error GoTo 0
End Function
