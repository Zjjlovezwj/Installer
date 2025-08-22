VERSION 5.00
Begin VB.Form frmFolder 
   Caption         =   "[ProductName] [Setup]"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7215
   Icon            =   "frmFolder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7215
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command4 
      Caption         =   "浏览(&O)..."
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2520
      Width           =   4815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "上一步(&B)"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "下一步(&N)"
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消"
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
      TabIndex        =   9
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "文件夹(&F):"
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "要安装到此文件夹，请单击 ""[Text_Next]""，要安装到不同文件夹，请在下面输入或单击 ""浏览"" 按钮。"
      Height          =   735
      Left            =   480
      TabIndex        =   7
      Top             =   1440
      Width           =   6135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "这是将被安装 [ProductName] 的文件夹。"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "选择安装文件夹"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   3615
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
      Height          =   885
      Left            =   -240
      Picture         =   "frmFolder.frx":2502
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7500
   End
End
Attribute VB_Name = "frmFolder"
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
' API声明
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" ( _
    ByVal hwndOwner As Long, _
    ByVal nFolder As Long, _
    pidl As Long _
) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" ( _
    ByVal pidl As Long, _
    ByVal pszPath As String _
) As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

' 常量
Private Const CSIDL_PROGRAM_FILES = &H26     ' Program Files
Private Const MAX_PATH = 260

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

' 备用方法：通过环境变量获取
Private Function GetProgramFilesPath() As String
    On Error Resume Next
    GetProgramFilesPath = Environ("ProgramFiles")
    If Right(GetProgramFilesPath, 1) <> "\" Then
        GetProgramFilesPath = GetProgramFilesPath & "\"
    End If
    On Error GoTo 0
End Function

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
    frmCancel.Show 1
    On Error GoTo 0
End Sub

Private Sub Command2_Click()
On Error Resume Next
    Dim configPath As String
    configPath = App.path & "\Setup.inf"
    Me.Hide
    frmInstall.Show
    frmInstall.Label6 = Me.Text1.Text & GetSetting(configPath, ID_MainAppPath)
    On Error GoTo 0
End Sub

Private Sub Command3_Click()
On Error Resume Next
    Me.Hide
    frmLisence.Show
    On Error GoTo 0
End Sub

Private Sub Command4_Click()
On Error Resume Next
    Dim sFolder As String
    sFolder = BrowseForFolder(Me.hwnd, "请选择文件夹")
    If sFolder <> "" Then
        Text1.Text = sFolder  ' 自动填写到 Text1
    End If
    On Error GoTo 0
End Sub

Private Sub Form_Load()
On Error Resume Next
    Timer1.Interval = 1
    Timer1.Enabled = True
    
    Me.Caption = GetStringByID("ID_4")
    Label1.Caption = GetStringByID("ID_18")
    Label2.Caption = GetStringByID("ID_19")
    Label3.Caption = GetStringByID("ID_20")
    Label4.Caption = GetStringByID("ID_21")
    Command1.Caption = GetStringByID("ID_7")
    Command2.Caption = GetStringByID("ID_8")
    Command3.Caption = GetStringByID("ID_9")
    Command4.Caption = GetStringByID("ID_22")
    
    '窗体启动时，记录下窗体的宽度和高度
    frmWidth = Me.Width
    frmHeight = Me.Height
    
    lockform Me

    Dim sPath As String
    
    ' 方法1：用API获取
    sPath = GetSpecialFolderPath(CSIDL_PROGRAM_FILES)
    
    ' 方法2：如果API失败，改用环境变量
    If sPath = "" Then
        sPath = Environ("ProgramFiles")
    End If
    
    ' 方法3：如果还是失败，使用默认值
    If sPath = "" Then
        sPath = "C:\Program Files"
    End If
    
    ' 显示路径并确保以反斜杠结尾
    Text1.Text = sPath
    If Right(Text1.Text, 1) <> "\" Then
        Text1.Text = Text1.Text & "\"
    End If
    
    ' 设置您想要的BorderStyle（例如2-固定对话框）
    Me.BorderStyle = 2
    
    ' 彻底移除系统菜单中的调整选项
    CompletelyRemoveSizeMenuItems Me.hwnd
    
    Dim hMenu As Long, hID As Long
    hMenu = GetSystemMenu(Me.hwnd, 0)
    
    InsertMenu hMenu, &HFFFFFFFF, MF_BYCOMMAND + MF_SEPARATOR, 0&, vbNullString
    
    InsertMenu hMenu, &HFFFFFFFF, MF_BYPOSITION, IDM.a, GetStringByID("ID_0")

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
    Select Case UnloadMode
        Case vbFormControlMenu  ' 用户点击关闭按钮或Alt+F4 (0)
            Cancel = True  ' 可以阻止关闭
            frmCancel.Show 1
   
        Case vbFormCode         ' 代码调用 Unload (1)
            Cancel = False
            
        Case vbAppWindows       ' Windows 退出 (2)
            Unload Me
            frmFolder.Visible = False
            Unload frmFolder
            frmInitialize.Visible = False
            Unload frmInitialize
            frmLang.Visible = False
            Unload frmLang
            frmInstall.Visible = False
            Unload frmInstall
            frmLisence.Visible = False
            Unload frmLisence
            frmError.Show
        Case vbAppTaskManager   ' 任务管理器结束进程 (3)
            Unload Me
            frmFolder.Visible = False
            Unload frmFolder
            frmInitialize.Visible = False
            Unload frmInitialize
            frmLang.Visible = False
            Unload frmLang
            frmInstall.Visible = False
            Unload frmInstall
            frmLisence.Visible = False
            Unload frmLisence
            frmError.Show
    End Select
On Error GoTo 0
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
    If Text1.Text = "" Then
        Command2.Enabled = False
    Else
        Command2.Enabled = True
    End If
    On Error GoTo 0
End Sub
