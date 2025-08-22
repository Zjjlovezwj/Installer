Attribute VB_Name = "comInstall"
Option Explicit

' API声明
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
    ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
    ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
    ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

' 常量定义
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const KEY_ALL_ACCESS = &H3F
Private Const REG_SZ = 1
Private Const REG_DWORD = 4

Public Sub CreateUninstallEntry( _
    ByVal AppName As String, _
    ByVal DisplayName As String, _
    ByVal Publisher As String, _
    ByVal DisplayVersion As String, _
    ByVal InstallLocation As String, _
    ByVal UninstallString As String, _
    ByVal DisplayIcon As String, _
    Optional ByVal EstimatedSize As Long = 0, _
    Optional ByVal NoModify As Long = 1, _
    Optional ByVal NoRepair As Long = 1, _
    Optional ByVal URLInfoAbout As String = "", _
    Optional ByVal HelpLink As String = "" _
)
    On Error Resume Next
    
    Dim hKey As Long, lResult As Long, lDisposition As Long
    Dim sKeyPath As String
    Dim AppId As String
    
    Dim configPath As String
    configPath = App.Path & "\Setup.inf"
    
    AppId = GetSetting(configPath, ID_AppId)
    
    ' 构建注册表路径
    sKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & AppId
    
    ' 创建注册表键
    lResult = RegCreateKeyEx(HKEY_LOCAL_MACHINE, sKeyPath, 0, vbNullString, 0, _
                            KEY_ALL_ACCESS, 0, hKey, lDisposition)
    
    If lResult <> 0 Then Exit Sub
    
    ' 写入基本卸载信息
    Call RegSetValueExStr(hKey, "DisplayName", DisplayName)
    Call RegSetValueExStr(hKey, "Publisher", Publisher)
    Call RegSetValueExStr(hKey, "DisplayVersion", DisplayVersion)
    Call RegSetValueExStr(hKey, "InstallLocation", InstallLocation)
    Call RegSetValueExStr(hKey, "UninstallString", UninstallString)
    Call RegSetValueExStr(hKey, "DisplayIcon", DisplayIcon)
    Call RegSetValueExDWord(hKey, "EstimatedSize", EstimatedSize)
    Call RegSetValueExDWord(hKey, "NoModify", NoModify)
    Call RegSetValueExDWord(hKey, "NoRepair", NoRepair)
    Call RegSetValueExStr(hKey, "URLInfoAbout", URLInfoAbout)
    Call RegSetValueExStr(hKey, "HelpLink", HelpLink)
    
    ' 关闭注册表键
    RegCloseKey hKey
    On Error GoTo 0
End Sub

' 辅助函数 - 写入字符串值
Private Sub RegSetValueExStr(ByVal hKey As Long, ByVal ValueName As String, ByVal ValueData As String)
On Error Resume Next
    RegSetValueEx hKey, ValueName, 0, REG_SZ, ByVal ValueData, Len(ValueData)
    On Error GoTo 0
End Sub

' 辅助函数 - 写入DWORD值
Private Sub RegSetValueExDWord(ByVal hKey As Long, ByVal ValueName As String, ByVal ValueData As Long)
On Error Resume Next
    RegSetValueEx hKey, ValueName, 0, REG_DWORD, ValueData, 4
    On Error GoTo 0
End Sub

Public Sub DeleteUninstallEntry(ByVal AppId As String)
On Error Resume Next
    Dim WSHShell As Object
    Set WSHShell = CreateObject("WScript.Shell")
    
    On Error Resume Next
    
    Dim sKeyPath As String
    sKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & AppId
    
    ' 删除所有值
    WSHShell.RegDelete sKeyPath & "\DisplayName"
    WSHShell.RegDelete sKeyPath & "\Publisher"
    WSHShell.RegDelete sKeyPath & "\DisplayVersion"
    WSHShell.RegDelete sKeyPath & "\InstallLocation"
    WSHShell.RegDelete sKeyPath & "\UninstallString"
    WSHShell.RegDelete sKeyPath & "\DisplayIcon"
    WSHShell.RegDelete sKeyPath & "\EstimatedSize"
    WSHShell.RegDelete sKeyPath & "\NoModify"
    WSHShell.RegDelete sKeyPath & "\NoRepair"
    WSHShell.RegDelete sKeyPath & "\URLInfoAbout"
    WSHShell.RegDelete sKeyPath & "\HelpLink"
    
    ' 删除键本身
    WSHShell.RegDelete sKeyPath
    
    Set WSHShell = Nothing
    On Error GoTo 0
End Sub

