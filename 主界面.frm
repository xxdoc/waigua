VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "全科系统辅助"
   ClientHeight    =   2235
   ClientLeft      =   -45
   ClientTop       =   -375
   ClientWidth     =   2595
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   2595
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   240
      Top             =   1440
   End
   Begin VB.Timer TimRun 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   1440
   End
   Begin VB.Timer TimHandle 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   960
      Top             =   1440
   End
   Begin VB.CommandButton cmdRun 
      BackColor       =   &H00FFFFFF&
      Caption         =   "开始"
      Enabled         =   0   'False
      Height          =   480
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2040
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Sub SetWindowPos _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal hWndInsertAfter As Long, _
                              ByVal X As Long, _
                              ByVal Y As Long, _
                              ByVal cx As Long, _
                              ByVal cy As Long, _
                              ByVal wFlags As Long)

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function ShowWindow _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal nCmdShow As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const HWND_TOPMOST = -1

Private Const SW_HIDE = 0 '隐藏窗口

Private Const SW_SHOW = 5 '显示窗口

Private Const HWND_NOTOPMOST = -2

Private Const SWP_NOSIZE = &H1

Private Const SWP_NOMOVE = &H2

Private Const SWP_NOACTIVATE = &H10

Private Const SWP_SHOWWINDOW = &H40

Private Declare Function FindWindow _
                Lib "user32" _
                Alias "FindWindowA" (ByVal lpClassName As String, _
                                     ByVal lpWindowName As String) As Long

Private Declare Function FindWindowEx _
                Lib "user32" _
                Alias "FindWindowExA" (ByVal hWnd1 As Long, _
                                       ByVal hWnd2 As Long, _
                                       ByVal lpsz1 As String, _
                                       ByVal lpsz2 As String) As Long

Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

Private Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const BM_CLICK As Long = &HF5

Private Declare Function PostMessage _
                Lib "user32" _
                Alias "PostMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long

Private Const WM_LBUTTONDOWN = &H201 '左键按下

Private Const WM_LBUTTONUP = &H202 '左键弹起

Private Const WM_CLOSE As Long = &H10&

Dim n                  As Long '设置时间间隔

Dim windowHandle       As Long '窗体句柄

Dim buttonHandle       As Long '按钮句柄

Dim popwindowsHandle   As Long '弹出窗口句柄

Dim windowName         As String '窗体名称

Dim popwindowName      As String '弹出窗口名称

Dim buttonName         As String '按钮名称

Dim exePach            As String '启动主程序路径

Dim formLeft           As Long '记录窗体位置

Dim formtop            As Long

Dim intTime            As Integer

Dim onTop              As Boolean '设置是否置顶

'拖动无边框窗体
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function ReleaseDC _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal hdc As Long) As Long

Const WM_NCLBUTTONDOWN = &HA1

Const HTCAPTION = 2

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        Call ReleaseCapture
        Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)

    End If

End Sub

Private Sub cmdRun_Click()

    If FindwindowHandle(windowName) > 0 Then
        ' ShowWindow windowHandle, SW_SHOW '恢复窗体
 
        If FindpopwindowHandle(popwindowName) > 0 Then '点击按钮后弹出了对话框
      
            SendMessage popwindowsHandle, WM_CLOSE, 0&, 0& '关闭窗口
            '发送消息关闭窗口

        End If
                
        cmdRun.Enabled = True
    Else
        cmdRun.Enabled = fasle
        cmdRun.Caption = "开始"
        TimHandle.Enabled = True '启动查找

    End If

    Select Case cmdRun.Caption

        Case "开始"

            If cmdRun.Enabled = False Then
                cmdRun.Caption = "开始"
                Form1.BackColor = vbWhite
            Else
                cmdRun.Caption = "停止"
                TimRun.Enabled = True '启动外挂
                Form1.BackColor = vbGreen

            End If

            If FindwindowHandle(windowName) > 0 Then
                ' ShowWindow windowHandle, SW_HIDE '隐藏窗体

            End If
           
        Case "停止"
            TimRun.Enabled = False '停止外挂
            cmdRun.Caption = "开始"
            Form1.BackColor = vbRed

    End Select

End Sub

Private Sub Form_Load()

    If App.LogMode Then SetWindowIcon Me.hwnd, "AAA"
    '读取ini设置
    mdlIni.FileName = Replace(App.Path + "\", "\\", "\") + "set.ini" '设置ini路径
    onTop = mdlIni.ReadData("set", "onTop", 1) '是否置顶
    intTime = mdlIni.ReadData("set", "Time", 1)

    If onTop = True Then
        SetWindowPos Me.hwnd, IIf(onTop, -1, -2), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

        '置顶
    End If

    formLeft = mdlIni.ReadData("set", "Left", (Screen.Width - Me.Width) / 2)
    formtop = mdlIni.ReadData("set", "Top", (Screen.Height - Me.ScaleHeight) / 2)
    
    Move formLeft, formtop '移动窗体到指定位置
    
    windowName = mdlIni.ReadData("set", "windowName", "门诊医生工作站(全科诊室)")
    popwindowName = mdlIni.ReadData("set", "popwindowName", "选取病人")
    buttonName = mdlIni.ReadData("set", "buttonName", "选病人")
    exePach = mdlIni.ReadData("set", "exePach", "")

    If exePach <> "" And FindwindowHandle(windowName) <= 0 Then  '如果设置了路径,并且程序没有启动的话.
        Set oShell = CreateObject("WSCript.shell")
        ' oShell.run "cmd /C " & exePach & sCommand, 0, True  ' wintWindowStyle = 0, so cmd will be hidden
        oShell.run "cmd /C " & exePach, 0, True   ' wintWindowStyle = 0, so cmd will be hidden
        ' Shell 启动程序 并且等待进程运行
        
    End If
    
    TimHandle.Enabled = True  ' 启动查找主程序窗口工作

End Sub

Private Function FindwindowHandle(winName As String) As Long '查找主窗口句柄
    ' winName = windowName
    FindwindowHandle = FindWindow("FNWND390", winName)
   
    windowHandle = FindwindowHandle

End Function

Private Function FindbuttonHandle(bonName As String) As Long '查找按钮句柄

    On Error Resume Next

    Dim WinWnd As Long

    Dim Ret    As Long

    Dim ErrNum As Integer
    
    'Find window Handle
    WinWnd = FindWindow("FNWND390", windowName)
    
    If WinWnd <> 0 Then
        'Show the form
        'AppActivate "email poster"
        'Find button handle by going through every child control in the form
        EnumChildWindows WinWnd, AddressOf EnumChildProc, ByVal 0&
        
        If BtnHwnd <> 0 Then
            FindbuttonHandle = BtnHwnd
            BtnHwnd = 0

        End If

    End If

    buttonHandle = FindbuttonHandle

End Function

Private Function FindpopwindowHandle(popwinName As String) As Long '查找弹出窗口句柄

    'popwindowName = popwinName
    '父窗口类名数组
    FindpopwindowHandle = FindWindowLike(0, popwinName, "FNWNS390")
    
    ' FindpopwindowHandle = FindWindow("FNWNS390", popwinName)
    Debug.Print FindpopwindowHandle
    '举例: Dim hLastWin as Long
    '      hLastWin = MyFindWindow()

    popwindowsHandle = FindpopwindowHandle

End Function

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then

        ' PopupMenu MNUexit
        If MsgBox("是:退出,否:不退出", vbOKCancel, "退出") = vbOK Then
            Unload Me
        Else

        End If

    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If FindwindowHandle(windowName) > 0 Then
        ShowWindow windowHandle, SW_SHOW '恢复窗体
 
        If FindpopwindowHandle(popwindowName) > 0 Then '点击按钮后弹出了对话框
      
            SendMessage popwindowsHandle, WM_CLOSE, 0&, 0& '关闭窗口
            '发送消息关闭窗口

        End If

    End If

    mdlIni.WriteData "set", "Left", Me.Left
    mdlIni.WriteData "set", "Top", Me.Top
 
End Sub

Private Sub subExit_Click()
    Unload Me

End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False

    If buttonHandle > 0 Then
        If FindpopwindowHandle(popwindowName) > 0 Then  '点击按钮后弹出了对话框
            ' SetActiveWindow popwindowsHandle
            'hwnd 为需要关闭的窗口程序的窗口句柄；
            ' ShowWindow popwindowsHandle, SW_HIDE
            'SetWindowPos popwindowsHandle, 1, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE '置后
            'SetWindowPos popwindowsHandle, -2, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE '正常显示
            SendMessage popwindowsHandle, WM_CLOSE, 0&, 0& '关闭窗口
            '发送的消息可以用两种选择1，wMsg为WM_CLOSE，wParam，lParam为0；2，wMsg为WM_SYSCOMMAND，wParam为CS_CLOSE，lParam为0。
     
            '发送消息关闭窗口1
            TimRun.Enabled = True

        End If

    End If

End Sub

Private Sub TimHandle_Timer() '查找主程序是否启动 '
 
    If FindwindowHandle(windowName) > 0 Then
        TimHandle.Enabled = False  '停止查找.
        Debug.Print "主程序已经启动"
        cmdRun.Enabled = True '开始按钮可用
        'ShowWindow windowHandle, SW_HIDE
    Else
        cmdRun.Enabled = False '开始按钮不可用

    End If
  
End Sub

Private Sub TimRun_Timer() '操作
    n = n + 1000 '

    If n >= intTime * 30000 Then '半分钟
       
        If FindpopwindowHandle(popwindowName) > 0 Then '点击按钮后弹出了对话框
      
            SendMessage popwindowsHandle, WM_CLOSE, 0&, 0& '关闭窗口
            '发送消息关闭窗口

        End If

        '找到按钮句柄
        If FindbuttonHandle(buttonName) > 0 Then
            SendMessage buttonHandle, BM_CLICK, 0, 0
            'sedmessage '发送按键点击按钮
            '        SetActiveWindow windowHandle
            '        PostMessage buttonHandle, WM_LBUTTONDOWN, 0, 0
            '        PostMessage buttonHandle, WM_LBUTTONUP, 0, 0
            '等待 窗口弹出
            ' SendClick buttonHandle, 130, 60
            Delay 3000
      
            Timer1.Enabled = True
            TimRun.Enabled = False
            n = 0
        Else
            TimRun.Enabled = False
            TimHandle.Enabled = True '查找
      
        End If
          
    End If

End Sub

Public Function SendClick(hwnd As Long, mX As Long, mY As Long)

    Dim i As Long
      
    i = PostMessage(hwnd, WM_LBUTTONDOWN, 0, (mX And &HFFFF) + (mY And &HFFFF) * &H10000)
 
    i = PostMessage(hwnd, WM_LBUTTONUP, 0, (mX And &HFFFF) + (mY And &HFFFF) * &H10000)

End Function

Private Sub Delay(Msec As Long)

    Dim EndTime As Long

    EndTime = GetTickCount + Msec
    Do
        Sleep 1
        DoEvents
    Loop While GetTickCount < EndTime

End Sub
