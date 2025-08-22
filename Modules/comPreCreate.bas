Attribute VB_Name = "comPreCreate"
Option Explicit

' 创建语言文件和目录的模块
Public Sub CreateLanguageFiles()
    Dim sAppPath As String
    Dim sLangPath As String
    Dim iFile As Integer
    On Error Resume Next
    
    ' 获取应用程序路径
    sAppPath = App.path
    If Right(sAppPath, 1) <> "\" Then sAppPath = sAppPath & "\"
    
    ' 创建Language目录
    sLangPath = sAppPath & "Language\"
    On Error Resume Next
    MkDir sLangPath
    On Error GoTo 0
    
    ' 检查目录是否创建成功
    If Dir(sLangPath, vbDirectory) = "" Then
        MsgBox "无法创建Language目录！", vbCritical
        Exit Sub
    End If
    
    ' 写入Chinese.lang文件
    iFile = FreeFile
    Open sLangPath & "Chinese.lang" For Output As #iFile
    Print #iFile, "ID_0=关于安装程序(&A)"
    Print #iFile, "ID_1=[ProductName] 安装程序"
    Print #iFile, "ID_2=Copyright (c) 2000-2025 AZ Studio. All rights reserved. 保留所有权利。"
    Print #iFile, "ID_3=警告: 本程序受版权保护，禁止未经允许反编译或者分发程序。"
    Print #iFile, "ID_4=[ProductName] 安装程序"
    Print #iFile, "ID_5=欢迎使用 [ProductName] 安装程序"
    Print #iFile, "ID_6=安装程序将在您的计算机上安装 [ProductName]。单击 ""下一步"" 继续，或单击 ""取消"" 退出安装程序。"
    Print #iFile, "ID_7=取消"
    Print #iFile, "ID_8=下一步(&N)"
    Print #iFile, "ID_9=上一步(&B)"
    Print #iFile, "ID_10=您确实要取消 [ProductName] 安装吗？"
    Print #iFile, "ID_11=是(&Y)"
    Print #iFile, "ID_12=否(&N)"
    Print #iFile, "ID_13=我接受许可协议中的条款(&A)"
    Print #iFile, "ID_14=我不接受许可协议中的条款(&D)"
    Print #iFile, "ID_15=最终用户许可协议"
    Print #iFile, "ID_16=请仔细阅读下列许可协议"
    Print #iFile, "ID_17=[Lisence]"
    Print #iFile, "ID_18=选择安装文件夹"
    Print #iFile, "ID_19=这是将被安装 [ProductName] 的文件夹。"
    Print #iFile, "ID_20=要安装到此文件夹，请单击 ""下一步""，要安装到不同文件夹，请在下面输入或单击 ""浏览"" 按钮。"
    Print #iFile, "ID_21=文件夹(&F):"
    Print #iFile, "ID_22=浏览(&O)..."
    Print #iFile, "ID_23=安装 [ProductName]"
    Print #iFile, "ID_24=请稍候，安装程序正在安装 [ProductName]。可能需要几分钟。"
    Print #iFile, "ID_25=状态:"
    Print #iFile, "ID_26=[ProductName] 安装程序过早地结束"
    Print #iFile, "ID_27=[ProductName] 的安装被中断。您的系统没有被修改。要稍后安装此程序，请再运行安装程序。"
    Print #iFile, "ID_28=单击 ""完成"" 按钮退出安装程序。"
    Print #iFile, "ID_29=完成(&F)"
    Print #iFile, "ID_30=正在完成 [ProductName] 安装程序"
    Print #iFile, "ID_31=启动 [ProductName]"
    Print #iFile, "ID_32=正在准备安装..."
    Print #iFile, "ID_33=正在复制文件..."
    Print #iFile, "ID_34=正在注册 [ProductName]..."
    Print #iFile, "ID_35=复制文件时出错"
    Print #iFile, "ID_36=安装过程中出错"
    Print #iFile, "ID_37=无法打开程序"
    Print #iFile, "ID_38=警告：复制文件时出现错误"
    Close #iFile
    
    ' 写入English.lang文件
    iFile = FreeFile
    Open sLangPath & "English.lang" For Output As #iFile
    Print #iFile, "ID_0=About Installer(&A)"
    Print #iFile, "ID_1=[ProductName] Setup"
    Print #iFile, "ID_2=Copyright (c) 2000-2025 AZ Studio. All rights reserved."
    Print #iFile, "ID_3=Warning: This program is protected by copyright law. Unauthorized decompilation or distribution is prohibited."
    Print #iFile, "ID_4=[ProductName] Setup"
    Print #iFile, "ID_5=Welcome to the [ProductName] Setup"
    Print #iFile, "ID_6=The Setup will install [ProductName] on your computer. Click ""Next"" to continue, or click ""Cancel"" to exit the Setup."
    Print #iFile, "ID_7=Cancel"
    Print #iFile, "ID_8=Next(&N)"
    Print #iFile, "ID_9=Back(&B)"
    Print #iFile, "ID_10=Are you sure you want to cancel the [ProductName] installation?"
    Print #iFile, "ID_11=Yes(&Y)"
    Print #iFile, "ID_12=No(&N)"
    Print #iFile, "ID_13=I accept the terms in the License Agreement(&A)"
    Print #iFile, "ID_14=I do not accept the terms in the License Agreement(&D)"
    Print #iFile, "ID_15=End-User License Agreement"
    Print #iFile, "ID_16=Please read the following License Agreement carefully"
    Print #iFile, "ID_17=[License]"
    Print #iFile, "ID_18=Select Installation Folder"
    Print #iFile, "ID_19=This is the folder where [ProductName] will be installed."
    Print #iFile, "ID_20=To install to this folder, click ""Next"". To install to a different folder, enter the path below or click the ""Browse"" button."
    Print #iFile, "ID_21=Folder(&F):"
    Print #iFile, "ID_22=Browse(&O)..."
    Print #iFile, "ID_23=Install [ProductName]"
    Print #iFile, "ID_24=Please wait while the Setup is installing [ProductName]. This may take a few minutes."
    Print #iFile, "ID_25=Status:"
    Print #iFile, "ID_26=[ProductName] Setup Ended Prematurely"
    Print #iFile, "ID_27=The installation of [ProductName] was interrupted. Your system has not been modified. To install this program later, run the setup again."
    Print #iFile, "ID_28=Click the ""Finish"" button to exit the Setup."
    Print #iFile, "ID_29=Finish(&F)"
    Print #iFile, "ID_30=Completing the [ProductName] Setup"
    Print #iFile, "ID_31=Launch [ProductName]"
    Print #iFile, "ID_32=Preparing to Install..."
    Print #iFile, "ID_33=Copying Files..."
    Print #iFile, "ID_34=Registering [ProductName]..."
    Print #iFile, "ID_35=Error Copying Files"
    Print #iFile, "ID_36=Error During Installation"
    Print #iFile, "ID_37=Failed to Open Program"
    Print #iFile, "ID_38=Warning:Error Copying Files"
    Close #iFile
    On Error GoTo 0
End Sub
