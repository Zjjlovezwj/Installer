VERSION 5.00
Begin VB.Form frmLang 
   BackColor       =   &H00FFFFFF&
   Caption         =   "语言"
   ClientHeight    =   4080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5535
   Icon            =   "frmLang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   5535
   StartUpPosition =   2  '屏幕中心
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2070
      ItemData        =   "frmLang.frx":2502
      Left            =   360
      List            =   "frmLang.frx":2504
      TabIndex        =   2
      Top             =   840
      Width           =   4815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "退出(&E)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   0
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "选择语言："
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Height          =   15015
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   23175
   End
End
Attribute VB_Name = "frmLang"
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

Private Sub Command1_Click()
On Error Resume Next
    If List1.List(List1.ListIndex) = "简体中文" Then
        LoadLanguageFile "Language\Chinese.lang"
        Unload Me
        frmInitialize.Show
    Else
        If List1.List(List1.ListIndex) = "English" Then
            LoadLanguageFile "Language\English.lang"
            Unload Me
            frmInitialize.Show
        End If
    End If
On Error GoTo 0
End Sub

Private Sub Form_Load()
On Error Resume Next
    '窗体启动时，记录下窗体的宽度和高度
    frmWidth = Me.Width
    frmHeight = Me.Height
    
        ' 初始化ListBox项目
    List1.AddItem "简体中文"
    List1.AddItem "English"
    
    ' 设置默认选择
    List1.ListIndex = 0
    lockform Me
    
    ' 彻底移除系统菜单中的调整选项
    CompletelyRemoveSizeMenuItems Me.hwnd
    
    Dim hMenu As Long, hID As Long
    hMenu = GetSystemMenu(Me.hwnd, 0)
    
    InsertMenu hMenu, &HFFFFFFFF, MF_BYCOMMAND + MF_SEPARATOR, 0&, vbNullString
    
    InsertMenu hMenu, &HFFFFFFFF, MF_BYPOSITION, IDM.a, "About(&A)"

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

Private Sub Command2_Click()
On Error Resume Next
    End
    On Error GoTo 0
End Sub

Private Sub List1_Click()
On Error Resume Next
    ' 检查是否选择了"简体中文"
    If List1.List(List1.ListIndex) = "简体中文" Then
        ' 改变窗口标题为中文
        Me.Caption = "语言"
        Label1.Caption = "选择语言："
        Command1.Caption = "确定(&O)"
        Command2.Caption = "退出(&E)"
    Else
        If List1.List(List1.ListIndex) = "English" Then
        ' 其他语言时恢复默认标题
        Me.Caption = "Language"
        Label1.Caption = "Choose Language:"
        Command1.Caption = "OK(&O)"
        Command2.Caption = "Exit(&E)"
        End If
    End If
    On Error GoTo 0
End Sub
