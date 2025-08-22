Attribute VB_Name = "comRemove"

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

Private Const GWL_STYLE = (-16)
Private Const WS_MINIMIZEBOX = &H20000

Public Sub RemoveMinimizeButton(frm As Form)
On Error Resume Next
    Dim lStyle As Long
    ' 获取当前窗口样式
    lStyle = GetWindowLong(frm.hwnd, GWL_STYLE)
    ' 移除最小化按钮样式
    lStyle = lStyle And Not WS_MINIMIZEBOX
    ' 设置新样式
    SetWindowLong frm.hwnd, GWL_STYLE, lStyle
    ' 重绘菜单栏
    DrawMenuBar frm.hwnd
    On Error GoTo 0
End Sub
