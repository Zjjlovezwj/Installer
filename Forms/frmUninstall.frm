VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.ocx"
Begin VB.Form frmUninstall 
   Caption         =   "AZ Studio Installer"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7215
   Icon            =   "frmUninstall.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7215
   StartUpPosition =   2  '屏幕中心
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   2640
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
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
      Caption         =   "Next(&N)"
      Enabled         =   0   'False
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
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   15
      Left            =   480
      TabIndex        =   10
      Top             =   3720
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   3360
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   15
      Left            =   600
      TabIndex        =   8
      Top             =   3600
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   2400
      Width           =   5295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Progress:"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label4 
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
      Caption         =   $"frmUninstall.frx":2502
      Height          =   735
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   6135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Uninstall"
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
      Picture         =   "frmUninstall.frx":2551
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7500
   End
End
Attribute VB_Name = "frmUninstall"
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
Private Const MAX_PATH = 260
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const INVALID_HANDLE_VALUE = -1

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As Currency
    ftLastAccessTime As Currency
    ftLastWriteTime As Currency
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As Any) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long

' 全局变量
Private bCancel As Boolean
Private lTotalFiles As Long
Private lCopiedFiles As Long
Private lSkippedFiles As Long
Private sLog As String
Private bCounting As Boolean ' 新增：标记是否正在计数文件

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


Private Sub Form_Load()
On Error Resume Next
    Dim configPath As String
    configPath = App.path & "\Setup.inf"
    Label6.Visible = False
    Label7.Visible = False
    Label8.Visible = False
    Label6.Caption = GetSetting(configPath, ID_AppId)
    
    '窗体启动时，记录下窗体的宽度和高度
    frmWidth = Me.Width
    frmHeight = Me.Height
    
    DeleteShortcuts Label8.Caption, True
    
    lockform Me
    
    DeleteUninstallEntry Label6.Caption
    
    Timer1.Interval = 10
    Timer1.Enabled = True
    
    Label7.Caption = GetSetting(configPath, ID_AppName)
    Label8.Caption = GetSetting(configPath, ID_DisplayName)
    
    ' 设置您想要的BorderStyle（例如2-固定对话框）
    Me.BorderStyle = 2
    
    ' 彻底移除系统菜单中的调整选项
    CompletelyRemoveSizeMenuItems Me.hwnd
    
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

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Cancel = True
    On Error GoTo 0
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
    Label5.Caption = "Uninstalling"
    If ProgressBar1.Value = 100 Then
        Me.Visible = False
        Me.Hide
        frmComplete2.Show
        Timer1.Enabled = False
        Timer1.Interval = 0
    Else
        ProgressBar1.Value = ProgressBar1.Value + 1
    End If
On Error GoTo 0
End Sub
