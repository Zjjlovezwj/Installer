VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.ocx"
Begin VB.Form frmInstall 
   Caption         =   "[ProductName] [Setup]"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7215
   Icon            =   "frmInstall.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7215
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer4 
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   2640
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "上一步(&B)"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "下一步(&N)"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消"
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
      TabIndex        =   11
      Top             =   3720
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   3360
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   15
      Left            =   600
      TabIndex        =   9
      Top             =   3600
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   2400
      Width           =   5295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "状态:"
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
      Caption         =   "请稍候，[Wizard]正在[Progress2] [ProductName]。可能需要几分钟。"
      Height          =   735
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   6135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "[Progress1] [ProductName]"
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
      Picture         =   "frmInstall.frx":2502
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7500
   End
End
Attribute VB_Name = "frmInstall"
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

Public Sub CreateShortcut(ByVal sTargetPath As String, ByVal sShortcutName As String, ByVal sDescription As String)
On Error Resume Next
    Dim WSHShell As Object
    Set WSHShell = CreateObject("WScript.Shell")
    
    ' 桌面快捷方式
    Dim sDesktopPath As String
    sDesktopPath = WSHShell.SpecialFolders("Desktop")
    
    Dim oShortcut As Object
    Set oShortcut = WSHShell.CreateShortcut(sDesktopPath & "\" & sShortcutName & ".lnk")
    
    With oShortcut
        .TargetPath = sTargetPath
        .WorkingDirectory = Left(sTargetPath, InStrRev(sTargetPath, "\") - 1)
        .Description = sDescription
        .Save
    End With
    
    ' 开始菜单快捷方式
    Dim sStartMenuPath As String
    sStartMenuPath = WSHShell.SpecialFolders("StartMenu") & "\Programs"
    
    If Not Dir(sStartMenuPath, vbDirectory) <> "" Then
        MkDir sStartMenuPath
    End If
    
    Set oShortcut = WSHShell.CreateShortcut(sStartMenuPath & "\" & sShortcutName & ".lnk")
    
    With oShortcut
        .TargetPath = sTargetPath
        .WorkingDirectory = Left(sTargetPath, InStrRev(sTargetPath, "\") - 1)
        .Description = sDescription
        .Save
    End With
    
    Set WSHShell = Nothing
    On Error GoTo 0
End Sub

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
    
    frmCopy.Show
    frmCopy.Visible = False
    
    Me.Caption = GetStringByID("ID_4")
    Label1.Caption = GetStringByID("ID_23")
    Label2.Caption = GetStringByID("ID_24")
    Label3.Caption = GetStringByID("ID_25")
    Command1.Caption = GetStringByID("ID_7")
    Command2.Caption = GetStringByID("ID_8")
    Command3.Caption = GetStringByID("ID_9")
    
    '窗体启动时，记录下窗体的宽度和高度
    frmWidth = Me.Width
    frmHeight = Me.Height
    
    lockform Me
    
    ' 自动开始复制
    StartCopying
    
    Timer1.Interval = 1
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
    Label5.Caption = GetStringByID("ID_32")
    If ProgressBar1.Value = 40 Then
        Timer2.Enabled = True
        Timer2.Interval = 20
        Timer1.Enabled = False
        Timer1.Interval = 0
        Label5.Caption = GetStringByID("ID_33")
    Else
        ProgressBar1.Value = ProgressBar1.Value + 2
    End If
On Error GoTo 0
End Sub

Private Sub Timer2_Timer()
On Error GoTo ErrorHandler
    DoEvents
    ' 自动开始复制
    StartCopying
    
    Label5.Caption = GetStringByID("ID_34")
    Timer3.Enabled = True
    Timer3.Interval = 200
    Timer2.Enabled = False
    Timer2.Interval = 0
    Exit Sub
    
    If ProgressBar1.Value = 80 Then
        Timer3.Enabled = True
        Timer3.Interval = 200
        Timer2.Enabled = False
        Timer2.Interval = 0
        Label5.Caption = GetStringByID("ID_34")
    Else
        ProgressBar1.Value = ProgressBar1.Value + 1
    End If
    
ErrorHandler:
    Me.Hide
    Unload Me
    MsgBox GetStringByID("ID_35"), vbCritical, GetStringByID("ID_4")
    frmError.Show
    Timer2.Enabled = False
    Timer2.Interval = 0
End Sub

' 注册到控制面板
Private Sub RegisterInControlPanel(ByVal sInstallPath As String)
    On Error GoTo ErrorHandler
    Dim configPath As String
    Dim UninstallStr As String
    Dim sDisplayIcon As String
    
    ' 确保安装路径以反斜杠结尾
    If Right(sInstallPath, 1) <> "\" Then sInstallPath = sInstallPath & "\"
    configPath = App.path & "\Setup.inf"
    
    ' 修正卸载路径 - 使用ID_UninstallPath而不是ID_InstallLocation
    UninstallStr = """" & sInstallPath & GetSetting(configPath, ID_UninstallPath) & """ /u"
    
    ' 修正显示图标路径
    sDisplayIcon = sInstallPath & GetSetting(configPath, ID_MainAppPath)
    
    ' 创建注册表项
    CreateUninstallEntry _
        AppName:=GetSetting(configPath, ID_AppName), _
        DisplayName:=GetSetting(configPath, ID_DisplayName), _
        Publisher:=GetSetting(configPath, ID_Publisher), _
        DisplayVersion:=GetSetting(configPath, ID_DisplayVersion), _
        InstallLocation:=sInstallPath & GetSetting(configPath, ID_InstallLocation), _
        UninstallString:=UninstallStr, _
        DisplayIcon:=sDisplayIcon, _
        EstimatedSize:=GetNumberSetting(configPath, ID_EstimatedSize), _
        URLInfoAbout:=GetSetting(configPath, ID_URLInfoAbout), _
        HelpLink:=GetSetting(configPath, ID_HelpLink)
    
    Exit Sub
    
ErrorHandler:
    Me.Hide
    Unload Me
    MsgBox GetStringByID("ID_36"), vbCritical, GetStringByID("ID_4")
    frmError.Show
    Timer3.Enabled = False
    Timer3.Interval = 0
End Sub

Private Sub Timer3_Timer()
    On Error GoTo ErrorHandler
    Dim configPath As String
    Dim sInstallPath As String
    
    ' 禁用定时器防止重复触发
    Timer3.Enabled = False
    
    ' 获取并验证安装路径
    sInstallPath = frmFolder.Text1.Text
    If sInstallPath = "" Then
        Me.Hide
        Unload Me
        MsgBox GetStringByID("ID_36"), vbCritical, GetStringByID("ID_4")
        frmError.Show
    End If
    
    ' 确保路径以反斜杠结尾
    If Right(sInstallPath, 1) <> "\" Then sInstallPath = sInstallPath & "\"
    
    ' 检查配置文件是否存在
    configPath = App.path & "\Setup.inf"
    If Dir(configPath) = "" Then
        Me.Hide
        Unload Me
        MsgBox GetStringByID("ID_36"), vbCritical, GetStringByID("ID_4")
        frmError.Show
    End If
    
    ' 先注册控制面板项
    Call RegisterInControlPanel(sInstallPath)
    Call CreateShortcut(Label6.Caption, Label7.Caption, Label8.Caption)
    
    ' 更新进度
    If ProgressBar1.Value = 100 Then
        Me.Hide
        Unload Me
        frmComplete.Show
    Else
        ProgressBar1.Value = ProgressBar1.Value + 1
        Timer3.Enabled = True
    End If
    
    Exit Sub
    
ErrorHandler:
    Me.Hide
    Unload Me
    MsgBox GetStringByID("ID_36"), vbCritical, GetStringByID("ID_4")
    frmError.Show
End Sub

Private Sub StartCopying()
On Error Resume Next
    Dim SourcePath As String
    Dim DestPath As String
    
    DoEvents
    
    ' 设置路径
    SourcePath = App.path & "\Files\"
    DestPath = frmFolder.Text1.Text
    
    ' 验证路径
    If Not ValidatePaths(SourcePath, DestPath) Then Exit Sub
    
    ' 先计算总文件数
    bCounting = True
    DoEvents
    lTotalFiles = CountFiles(SourcePath)
    bCounting = False
    
    If lTotalFiles = 0 Then
        MsgBox GetStringByID("ID_33"), vbExclamation
        Exit Sub
    End If
    
    ' 使用Timer方式避免界面卡死
    Timer4.Interval = 1
    Timer4.Enabled = True
    On Error GoTo 0
End Sub

Private Function ValidatePaths(SourcePath As String, DestPath As String) As Boolean
On Error Resume Next
    ' 检查源路径
    If Dir(SourcePath, vbDirectory) = "" Then
        MsgBox GetStringByID("ID_35") & SourcePath, vbExclamation, GetStringByID("ID_4")
        Timer4.Enabled = False
        Timer4.Interval = 0
        Me.Hide
        frmError.Show
        Exit Function
    End If
    
    ' 检查目标路径
    If DestPath = "" Then
        MsgBox GetStringByID("ID_35"), vbExclamation, GetStringByID("ID_4")
        Timer4.Enabled = False
        Timer4.Interval = 0
        Me.Hide
        frmError.Show
        Exit Function
    End If
    
    ' 添加路径分隔符
    If Right(DestPath, 1) <> "\" Then DestPath = DestPath & "\"
    
    ' 检查是否相同路径
    If UCase(SourcePath) = UCase(DestPath) Then
        MsgBox GetStringByID("ID_38"), vbExclamation, GetStringByID("ID_4")
        Exit Function
    End If
    
    ValidatePaths = True
    On Error GoTo 0
End Function

Private Function CountFiles(sPath As String) As Long
On Error Resume Next
    Dim hFind As Long
    Dim WFD As WIN32_FIND_DATA
    Dim sFile As String
    Dim lCount As Long
    
    hFind = FindFirstFile(sPath & "*.*", WFD)
    If hFind = INVALID_HANDLE_VALUE Then Exit Function
    
    Do
        sFile = TrimNull(WFD.cFileName)
        
        If sFile <> "." And sFile <> ".." Then
            If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                lCount = lCount + CountFiles(sPath & sFile & "\")
            Else
                lCount = lCount + 1
            End If
        End If
        
        ' 更新UI
        If bCounting Then
            DoEvents
        End If
    Loop While FindNextFile(hFind, WFD) And Not bCancel
    
    FindClose hFind
    CountFiles = lCount
    On Error GoTo 0
End Function
Private Sub Timer4_Timer()
    On Error Resume Next
    
    Static hFind As Long
    Static WFD As WIN32_FIND_DATA
    Static sCurrentSourcePath As String
    Static sCurrentDestPath As String
    Static sCurrentDir As String
    Static sCurrentFile As String
    Static bProcessingDir As Boolean
    Static bFirstRun As Boolean
    
    ' 初始化（首次运行时）
    If hFind = 0 Then
        sCurrentSourcePath = App.path & "\Files\"
        sCurrentDestPath = frmFolder.Text1.Text
        If Right(sCurrentDestPath, 1) <> "\" Then sCurrentDestPath = sCurrentDestPath & "\"
        sCurrentDir = sCurrentSourcePath
        hFind = FindFirstFile(sCurrentDir & "*.*", WFD)
        If hFind = INVALID_HANDLE_VALUE Then
            Timer4.Enabled = False
            Exit Sub
        End If
        bFirstRun = True
        bProcessingDir = False
    End If
    
    ' 如果是第一次运行或者是处理新找到的文件
    If bFirstRun Or sCurrentFile = "" Then
        sCurrentFile = TrimNull(WFD.cFileName)
        
        ' 跳过 "." 和 ".."
        Do While sCurrentFile = "." Or sCurrentFile = ".."
            If FindNextFile(hFind, WFD) = 0 Then
                ' 当前目录处理完毕，准备返回上级
                GoTo ProcessNextDir
            End If
            sCurrentFile = TrimNull(WFD.cFileName)
        Loop
        bFirstRun = False
    End If
    
    ' 构建完整路径
    Dim sSrcPath As String, sDstPath As String
    sSrcPath = sCurrentDir & sCurrentFile
    sDstPath = sCurrentDestPath & Mid(sCurrentDir, Len(sCurrentSourcePath)) & sCurrentFile
    
    ' 判断是目录还是文件
    If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
        ' 处理目录
        If Not bProcessingDir Then
            ' 第一次处理这个目录
            If GetFileAttributes(sDstPath) = -1 Then
                If CreateDirectory(sDstPath, ByVal 0&) = 0 Then
                    lSkippedFiles = lSkippedFiles + 1
                End If
            End If
            
            ' 进入子目录
            FindClose hFind
            sCurrentDir = sSrcPath & "\"
            hFind = FindFirstFile(sCurrentDir & "*.*", WFD)
            bProcessingDir = True
            sCurrentFile = ""
            bFirstRun = True
        Else
            ' 已经处理过目录创建，继续下一个文件
            bProcessingDir = False
            If FindNextFile(hFind, WFD) = 0 Then
ProcessNextDir:
                ' 当前目录处理完毕，返回上级
                FindClose hFind
                If sCurrentDir = sCurrentSourcePath Then
                
                    Exit Sub
                Else
                    ' 返回上级目录
                    sCurrentDir = Left(sCurrentDir, InStrRev(sCurrentDir, "\", Len(sCurrentDir) - 2)) & "\"
                    hFind = FindFirstFile(sCurrentDir & "*.*", WFD)
                    bFirstRun = True
                    sCurrentFile = ""
                End If
            Else
                sCurrentFile = TrimNull(WFD.cFileName)
            End If
        End If
    Else
        ' 处理文件
        If FileExists(sDstPath) Then
            lSkippedFiles = lSkippedFiles + 1
        Else
            If CopyFile(sSrcPath, sDstPath, 0) = 0 Then
                lSkippedFiles = lSkippedFiles + 1
            Else
                lCopiedFiles = lCopiedFiles + 1
            End If
        End If
        
        ' 更新进度
        If lTotalFiles > 0 Then
            Dim iPercent As Integer
            iPercent = ((lCopiedFiles + lSkippedFiles) / lTotalFiles) * 100
            ' 更新进度显示代码
            DoEvents
        End If
        
        ' 准备处理下一个文件
        If FindNextFile(hFind, WFD) = 0 Then
            sCurrentFile = ""
            GoTo ProcessNextDir
        Else
            sCurrentFile = TrimNull(WFD.cFileName)
        End If
    End If
End Sub

' 辅助函数：检查文件是否存在
Private Function FileExists(ByVal sPath As String) As Boolean
    On Error Resume Next
    FileExists = (GetFileAttributes(sPath) <> -1) And _
                ((GetFileAttributes(sPath) And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY)
End Function

Private Sub ReturnToParentDirectory(ByRef hFind As Long, ByRef sCurrentDir As String, _
                                   ByVal sRootDir As String, WFD As WIN32_FIND_DATA)
    On Error Resume Next
    ' 关闭当前句柄
    FindClose hFind
    hFind = 0
    
    ' 如果不是根目录，返回上级
    If sCurrentDir <> sRootDir Then
        sCurrentDir = Left(sCurrentDir, InStrRev(sCurrentDir, "\", Len(sCurrentDir) - 1))
        hFind = FindFirstFile(sCurrentDir & "*.*", WFD)
    End If
    On Error GoTo 0
End Sub

Private Function TrimNull(sItem As String) As String
On Error Resume Next
    Dim lPos As Long
    lPos = InStr(sItem, vbNullChar)
    If lPos > 0 Then
        TrimNull = Left$(sItem, lPos - 1)
    Else
        TrimNull = sItem
    End If
    On Error GoTo 0
End Function



