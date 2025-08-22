VERSION 5.00
Begin VB.Form frmCancel 
   Caption         =   "[ProductName] [Setup]"
   ClientHeight    =   2055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5535
   Icon            =   "frmCancel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5535
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "是(&Y)"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "否(&N)"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000A&
      X1              =   0
      X2              =   5520
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   360
      Picture         =   "frmCancel.frx":2502
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "您确实要取消 [ProductName] 安装吗？"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
End
Attribute VB_Name = "frmCancel"
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

Private Sub Form_Resize()
    '用户改变窗体大小时，强制窗体大小固定为原始尺寸
    '从而达到窗体大小固定目的！
    On Error Resume Next
    Me.Width = frmWidth
    Me.Height = frmHeight
    Me.WindowState = 0
    On Error GoTo 0
End Sub


Private Sub Command1_Click()
On Error Resume Next
    Unload Me
    On Error GoTo 0
End Sub

Private Sub Command2_Click()
On Error Resume Next
    Me.Visible = False
    Unload Me
    frmFolder.Visible = False
    Unload frmFolder
    frmInitialize.Visible = False
    Unload frmInitialize
    frmLang.Visible = False
    Unload frmLang
    frmLisence.Visible = False
    Unload frmLisence
    frmError.Show
On Error GoTo 0
End Sub

Private Sub Form_Load()
On Error Resume Next
    
    Me.Caption = GetStringByID("ID_4")
    Label1.Caption = GetStringByID("ID_10")
    Command1.Caption = GetStringByID("ID_12")
    Command2.Caption = GetStringByID("ID_11")
    
    '窗体启动时，记录下窗体的宽度和高度
    frmWidth = Me.Width
    frmHeight = Me.Height
    
    lockform Me
    ' 设置您想要的BorderStyle（例如2-固定对话框）
    Me.BorderStyle = 2
    
    ' 彻底移除系统菜单中的调整选项
    CompletelyRemoveSizeMenuItems Me.hwnd

    
    On Error GoTo 0
End Sub
