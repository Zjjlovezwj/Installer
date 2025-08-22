Attribute VB_Name = "comLockForm"
Option Explicit

' API函数声明
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

' 点坐标结构
Private Type POINTAPI
    X As Long  ' X坐标
    Y As Long  ' Y坐标
End Type

' 窗口大小限制信息结构
Private Type MINMAXINFO
    ptReserved As POINTAPI      ' 保留字段
    ptMaxSize As POINTAPI       ' 最大化尺寸
    ptMaxPosition As POINTAPI   ' 最大化位置
    ptMinTrackSize As POINTAPI  ' 最小可调整尺寸
    ptMaxTrackSize As POINTAPI  ' 最大可调整尺寸
End Type

' 内存拷贝函数声明
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)

' Windows消息常量
Private Const WM_GETMINMAXINFO = &H24  ' 获取窗口大小限制信息
Private Const GWL_WNDPROC = (-4)       ' 窗口过程指针索引

' 全局变量
Dim OldProc&      ' 保存原始窗口过程地址
Dim mhWnd&        ' 窗口句柄
Dim mkd As Long   ' 窗口宽度(像素)
Dim mgd As Long   ' 窗口高度(像素)

' 锁定窗体边界(禁止调整大小)
' 参数: nhWnd - 要锁定的窗体对象
Sub lockform(nhWnd As Form)
    On Error Resume Next  ' 启用错误处理
    
    ' 如果已经锁定则退出
    If OldProc <> 0 Then Exit Sub
    
    ' 保存窗口句柄和尺寸(转换为像素)
    mhWnd& = nhWnd.hwnd
    mgd = nhWnd.Height / Screen.TwipsPerPixelY
    mkd = nhWnd.Width / Screen.TwipsPerPixelX
    
    ' 替换窗口过程为自定义过程
    OldProc = SetWindowLong(mhWnd&, GWL_WNDPROC, AddressOf WinProc)
    
    On Error GoTo 0  ' 关闭错误处理
End Sub

' 自定义窗口过程
Function WinProc&(ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam&)
    On Error Resume Next
    
    ' 处理特定消息
    Select Case wMsg&
        Case WM_GETMINMAXINFO  ' 窗口查询大小限制信息
            Dim MinMax As MINMAXINFO
            
            ' 从lParam获取原始MINMAXINFO结构
            CopyMemory MinMax, ByVal lParam, Len(MinMax)
            
            ' 设置最小和最大跟踪尺寸(锁定窗口大小)
            MinMax.ptMinTrackSize.X = mkd  ' 最小宽度
            MinMax.ptMinTrackSize.Y = mgd  ' 最小高度
            MinMax.ptMaxTrackSize.X = mkd  ' 最大宽度
            MinMax.ptMaxTrackSize.Y = mgd  ' 最大高度
            
            ' 将修改后的结构拷贝回lParam
            CopyMemory ByVal lParam, MinMax, Len(MinMax)
            
            WinProc& = 1  ' 表示已处理此消息
            Exit Function
    End Select
    
    ' 其他消息传递给原始窗口过程
    WinProc& = CallWindowProc(OldProc, hwnd, wMsg, wParam, lParam)
    
    On Error GoTo 0
End Function
