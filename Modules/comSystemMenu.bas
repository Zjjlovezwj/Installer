Attribute VB_Name = "comSystemMenu"
Option Explicit

Public Enum IDM
    a = 128
End Enum
Public procOld As Long
Private Declare Function CallWindowProc& Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc&, ByVal hwnd&, ByVal Msg&, ByVal wParam&, ByVal lParam&)
Dim m_transparencyKey As Long

Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Public Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd&, ByVal nIndex&, ByVal dwNewLong&)
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Public Const SC_CLOSE As Long = &HF060&

Public Const GWL_WNDPROC As Long = (-4&)

Public Const MF_BYCOMMAND As Long = &H0&
Public Const MF_BYPOSITION As Long = &H400&
Public Const MF_SEPARATOR As Long = &H800&
Public Const MF_CHECKED As Long = &H8&
Public Const MF_GRAYED As Long = &H1&
Public Const MF_BITMAP = &H4&
Public Const WM_SYSCOMMAND = &H112

Public Function WindowProc(ByVal hwnd As Long, _
ByVal iMsg As Long, ByVal wParam As Long, _
ByVal lParam As Long) As Long
On Error Resume Next

   Select Case iMsg
      Case WM_SYSCOMMAND
         Select Case wParam
         Case IDM.a
                MsgBox "AZ Studio Installer" & vbCrLf & _
                "Version 2.0 by Zjjlovelfl" & vbCrLf & _
                "Copyright (c) 2000-2025 AZ Studio. All rights reserved.", _
                vbInformation + vbOKOnly, _
                "About Installer"
         End Select
   End Select

    WindowProc = CallWindowProc(procOld, hwnd, iMsg, wParam, lParam)
    On Error GoTo 0
End Function
