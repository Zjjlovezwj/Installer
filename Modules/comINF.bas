Attribute VB_Name = "comINF"
Option Explicit

' API声明
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String _
) As Long

' 配置项ID常量
Public Enum SettingID
    ID_AppName = 1
    ID_DisplayName
    ID_Publisher
    ID_DisplayVersion
    ID_InstallLocation
    ID_UninstallString
    ID_MainAppPath
    ID_UninstallPath
    ID_EstimatedSize
    ID_AppId
    ID_URLInfoAbout
    ID_HelpLink
End Enum

' 获取配置值的公共接口
Public Function GetSetting(ByVal filePath As String, ByVal SettingID As SettingID) As String
On Error Resume Next
    Dim Section As String
    Dim Key As String
    
    ' 确定要读取的键名
    Select Case SettingID
        Case ID_AppId:            Key = "AppId"
        Case ID_AppName:         Key = "AppName"
        Case ID_DisplayName:      Key = "DisplayName"
        Case ID_Publisher:        Key = "Publisher"
        Case ID_DisplayVersion:   Key = "DisplayVersion"
        Case ID_InstallLocation:  Key = "InstallLocation"
        Case ID_UninstallString:  Key = "UninstallString"
        Case ID_MainAppPath:      Key = "MainAppPath"
        Case ID_UninstallPath:    Key = "UninstallPath"
        Case ID_EstimatedSize:    Key = "EstimatedSize"
        Case ID_URLInfoAbout:  Key = "URLInfoAbout"
        Case ID_HelpLink:    Key = "HelpLink"
        Case Else:               Exit Function
    End Select
    
    Section = "Settings"
    GetSetting = ReadINFValue(filePath, Section, Key)
    On Error GoTo 0
End Function

' 读取数字型配置
Public Function GetNumberSetting(ByVal filePath As String, ByVal SettingID As SettingID) As Long
On Error Resume Next
    Dim strValue As String
    strValue = GetSetting(filePath, SettingID)
    If IsNumeric(strValue) Then
        GetNumberSetting = CLng(strValue)
    Else
        GetNumberSetting = 0
    End If
    On Error GoTo 0
End Function

' 内部读取函数
Private Function ReadINFValue( _
    ByVal filePath As String, _
    ByVal Section As String, _
    ByVal Key As String, _
    Optional ByVal DefaultValue As String = "" _
) As String
On Error Resume Next
    Const BUFFER_SIZE As Long = 255
    Dim Buffer As String
    Dim Length As Long
    
    Buffer = String$(BUFFER_SIZE, vbNullChar)
    Length = GetPrivateProfileString(Section, Key, DefaultValue, Buffer, Len(Buffer), filePath)
    
    If Length > 0 Then
        ReadINFValue = Left$(Buffer, Length)
    Else
        ReadINFValue = DefaultValue
    End If
    On Error GoTo 0
End Function

