VERSION 5.00
Begin VB.Form frmComplete2 
   Caption         =   "AZ Studio Installer"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7215
   Icon            =   "frmComplete2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7215
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command3 
      Caption         =   "Back(&B)"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Finish(&F)"
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5760
      TabIndex        =   0
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "AZ Studio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Click the ""Finish"" button to exit the Setup."
      Height          =   735
      Left            =   2760
      TabIndex        =   4
      Top             =   1920
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Completing the Setup"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   3
      Top             =   360
      Width           =   4095
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000A&
      X1              =   0
      X2              =   7200
      Y1              =   4750
      Y2              =   4750
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   7200
      X2              =   0
      Y1              =   4750
      Y2              =   4750
   End
   Begin VB.Image Image1 
      Height          =   4740
      Left            =   0
      Picture         =   "frmComplete2.frx":2502
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7500
   End
End
Attribute VB_Name = "frmComplete2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim frmWidth
Dim frmHeight
' 使用您提供的API声明
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

' 常量定义
Private Const MF_BYPOSITION = &H400&
Private Const MF_BYCOMMAND = &H0&

' 系统菜单命令ID
Private Const SC_SIZE = &HF000
Private Const SC_MAXIMIZE = &HF030
Private Const SC_MINIMIZE = &HF020
Private Const SC_RESTORE = &HF120
Private Const SC_MOVE = &HF010
Private Const SC_CLOSE = &HF060
' 在模块中添加声明
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpCommandLine As String, _
    ByVal lpProcessAttributes As Long, _
    ByVal lpThreadAttributes As Long, _
    ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, _
    ByVal lpCurrentDirectory As String, _
    lpStartupInfo As STARTUPINFO, _
    lpProcessInformation As PROCESS_INFORMATION _
) As Long

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

' 彻底移除系统菜单中的大小调整选项
Public Sub CompletelyRemoveSizeMenuItems(hwnd As Long)
On Error Resume Next
    Dim hSysMenu As Long
    Dim i As Integer
    
    ' 先重置系统菜单（确保获取干净菜单）
    hSysMenu = GetSystemMenu(hwnd, 1)
    
    ' 获取系统菜单
    hSysMenu = GetSystemMenu(hwnd, 0)
    
    If hSysMenu Then
        ' 方法1：按命令ID删除（更可靠）
        DeleteMenu hSysMenu, SC_SIZE, MF_BYCOMMAND
        
        ' 方法2：按位置删除（双重保障）
        ' 遍历前10个菜单项（通常足够）
        For i = 0 To 9
        Next i
        
        ' 刷新菜单栏
        DrawMenuBar hwnd
    End If
    On Error GoTo 0
End Sub

Private Sub Command2_Click()
On Error Resume Next
    End
    On Error GoTo 0
End Sub

Private Sub Form_Load()
On Error Resume Next
    '窗体启动时，记录下窗体的宽度和高度
    frmWidth = Me.Width
    frmHeight = Me.Height
    
    lockform Me
    ' 设置您想要的BorderStyle（例如2-固定对话框）
    Me.BorderStyle = 2
    
    ' 彻底移除系统菜单中的调整选项
    CompletelyRemoveSizeMenuItems Me.hwnd
    
    Dim hMenu As Long, hID As Long
    hMenu = GetSystemMenu(Me.hwnd, 0)
    
    InsertMenu hMenu, &HFFFFFFFF, MF_BYCOMMAND + MF_SEPARATOR, 0&, vbNullString
    
    InsertMenu hMenu, &HFFFFFFFF, MF_BYPOSITION, IDM.a, "About Installer(&A)"
    
    DrawMenuBar hMenu
    
    procOld = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
    
    On Error GoTo 0
End Sub

Private Sub Form_Resize()
    '用户改变窗体大小时，强制窗体大小固定为原始尺寸
    '从而达到窗体大小固定目的！
    On Error Resume Next
    Me.Width = frmWidth
    Me.Height = frmHeight
    Me.WindowState = 0
    On Error GoTo 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
    End
    On Error GoTo 0
End Sub
