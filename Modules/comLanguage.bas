Attribute VB_Name = "comLanguage"

Public LanguageData As Object

Public Sub LoadLanguageFile(filePath As String)
    On Error Resume Next
    Dim fileNum As Integer
    Dim line As String
    Dim parts() As String
    
    Set LanguageData = CreateObject("Scripting.Dictionary")
    
    fileNum = FreeFile
    Open filePath For Input As fileNum
    Do Until EOF(fileNum)
        Line Input #fileNum, line
        parts = Split(line, "=")
        If UBound(parts) = 1 Then
            LanguageData(parts(0)) = parts(1)
        End If
    Loop
    Close fileNum
On Error GoTo 0
End Sub

Public Function GetStringByID(id As String) As String
    On Error Resume Next
    GetStringByID = LanguageData(id)
    On Error GoTo 0
End Function

