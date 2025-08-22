Attribute VB_Name = "comMain"
Option Explicit

Sub Main()
    On Error Resume Next
    '预先创建语言文件（需修改comPreCreate模块）
    'CreateLanguageFiles
    Select Case Command()
        Case "/o"
            frmCmd.Show
        Case "/i"
            frmLang.Show
        Case "/q"
            LoadLanguageFile "Language\English.lang"
            frmFolder.Show
            frmFolder.Visible = False
            frmFolder.Command2 = True
        Case "/a"
            About
        Case "/u"
            frmLang2.Show
        Case Else
            MsgBox "Error" & vbCrLf & _
                  "Can not start installer", _
                  vbCritical + vbOKOnly
    End Select
    On Error GoTo 0
End Sub

Public Sub About()
On Error Resume Next
    MsgBox "AZ Studio Installer" & vbCrLf & _
           "Version 2.0 by Zjjlovelfl" & vbCrLf & _
           "Copyright (c) 2000-2025 AZ Studio. All rights reserved.", _
           vbInformation + vbOKOnly, _
           "About"
    On Error GoTo 0
End Sub

