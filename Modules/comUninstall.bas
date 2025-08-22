Attribute VB_Name = "comUninstall"
' 删除快捷方式主函数
Public Sub DeleteShortcuts(ByVal sShortcutName As String, Optional ByVal bSilent As Boolean = False)
    On Error Resume Next
    
    Dim WSHShell As Object
    Set WSHShell = CreateObject("WScript.Shell")
    
    ' 记录操作日志
    Dim sLog As String
    
    ' 1. 删除桌面快捷方式
    Dim sDesktopPath As String
    sDesktopPath = WSHShell.SpecialFolders("Desktop") & "\" & sShortcutName & ".lnk"
    
    If FileExists(sDesktopPath) Then
        Kill sDesktopPath
    End If
    
    ' 2. 删除开始菜单快捷方式
    Dim sStartMenuPath As String
    sStartMenuPath = WSHShell.SpecialFolders("StartMenu") & "\Programs\" & sShortcutName & ".lnk"
    
    If FileExists(sStartMenuPath) Then
        Kill sStartMenuPath
        
        ' 检查并删除空目录
        Dim sProgramsPath As String
        sProgramsPath = WSHShell.SpecialFolders("StartMenu") & "\Programs"
        If IsFolderEmpty(sProgramsPath) Then
            RmDir sProgramsPath
        End If
    End If

    Dim sQuickLaunchPath As String
    sQuickLaunchPath = WSHShell.SpecialFolders("QuickLaunch") & "\" & sShortcutName & ".lnk"
    
    If FileExists(sQuickLaunchPath) Then
        Kill sQuickLaunchPath
    End If
On Error GoTo 0
End Sub

' 辅助函数：检查文件是否存在
Private Function FileExists(ByVal sPath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(sPath, vbNormal Or vbHidden Or vbSystem Or vbArchive) <> "")
    On Error GoTo 0
End Function

' 辅助函数：检查文件夹是否为空
Private Function IsFolderEmpty(ByVal sFolderPath As String) As Boolean
    On Error Resume Next
    
    If Dir(sFolderPath, vbDirectory) = "" Then
        IsFolderEmpty = True
        Exit Function
    End If
    
    Dim sFile As String
    sFile = Dir(sFolderPath & "\*.*", vbNormal Or vbHidden Or vbSystem Or vbArchive)
    
    IsFolderEmpty = (sFile = "")
    Exit Function
    
On Error GoTo 0
End Function

