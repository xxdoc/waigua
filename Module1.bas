Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function FindWindow _
                Lib "user32" _
                Alias "FindWindowA" (ByVal lpClassName As Any, _
                                     ByVal lpWindowName As Any) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function GetWindow _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal wCmd As Long) As Long

Private Declare Function GetWindowText _
                Lib "user32" _
                Alias "GetWindowTextA" (ByVal hwnd As Long, _
                                        ByVal lpString As String, _
                                        ByVal cch As Long) As Long

Private Declare Function GetClassName _
                Lib "user32" _
                Alias "GetClassNameA" (ByVal hwnd As Long, _
                                       ByVal lpClassName As String, _
                                       ByVal nMaxCount As Long) As Long

'Private Const GW_HWNDFIRST = 0
'Private Const GW_HWNDLAST = 1
'Private Const GW_HWNDPREV = 3
'Private Const GW_OWNER = 4
Private Const GW_CHILD = 5

Private Const GW_HWNDNEXT = 2

Private Declare Function IsChild _
                Lib "user32" (ByVal hWndParent As Long, _
                              ByVal hwnd As Long) As Long

Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

Private Declare Function ShowWindow _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal nCmdShow As Long) As Long

Private Declare Function SendMessageByString _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As String) As Long

Private Const WM_SETTEXT = &HC

Private Const BM_CLICK = &HF5

Const WM_GETTEXT = &HD

Const WM_GETTEXTLENGTH = &HE

Dim NextHwnd(512) As Long

Dim ChildCount    As Integer

Public Function FindWindowLike(ByVal hWndStart As Long, _
                               WindowText As String, _
                               Classname As String) As Long

    Dim hwnd        As Long

    Dim sWindowText As String

    Dim sClassname  As String

    Dim r           As Long

    Dim sSave       As String

    Dim nSize       As Long

    'Hold the level of recursion and
    'hold the number of matching windows
    Static level    As Integer
 
    'Initialize if necessary. This is only executed when level = 0
    'and hWndStart = 0, normally only on the first call to the routine.
    If level = 0 Then
        If hWndStart = 0 Then hWndStart = GetDesktopWindow()

    End If
   
    'Increase recursion counter
    level = level + 1
   
    'Get first child window
    hwnd = GetWindow(hWndStart, GW_CHILD)

    Do Until hwnd = 0
        'Search children by recursion
        Call FindWindowLike(hwnd, WindowText, Classname)
       
        'Get the windowtext length
        nSize = SendMessageByString(hwnd, WM_GETTEXTLENGTH, 0, 0)
        sSave = Space$(nSize + 1)
        'get the window text
        SendMessageByString hwnd, WM_GETTEXT, nSize + 1, sSave
        'GetWindowText hWnd, sSave, Len(sSave)
        'remove the last Chr$(0)
        sSave = Left$(sSave, Len(sSave) - 1)
        'Get the window text and class name
        '        sWindowText = Space$(255)
        '        r = GetWindowText(hwnd, sWindowText, 255)
        '        sWindowText = Left(sWindowText, InStr(1, sWindowText, vbNullChar) - 1)
        sWindowText = sSave
        sClassname = Space$(255)
        r = GetClassName(hwnd, sClassname, 255)
        sClassname = Left(sClassname, InStr(1, sClassname, vbNullChar) - 1)
       
        'Check if window found matches the search parameters
        If (sWindowText Like WindowText) And (sClassname Like Classname) Then

            ' is it this form? if yes dont ADD to list
            '            If Me.hwnd = hwnd Then GoTo skip
            '            If IsChild(Me.hwnd, hwnd) <> 1 Then ' not part of me SO ADD
            NextHwnd(ChildCount) = hwnd
            ChildCount = ChildCount + 1

            '            End If
        
            FindWindowLike = hwnd
       
            Exit Do '> uncomment this line return ONLY the first matching window!

        End If

        'Get next child window
        'skip:
        hwnd = GetWindow(hwnd, GW_HWNDNEXT)
    Loop
 
    'Reduce the recursion counter
    level = level - 1

End Function

'Private Sub Command1_Click()
'
'    Dim i        As Integer
'
'    Dim NextBox  As Integer
'
'    Dim wnd1     As Long
'
'    Dim ChldHwnd As Long
'
'    ' find the program. (class name, titlebar caption)
'    wnd1 = FindWindow("ThunderRT6FormDC", "<----Titlebar Caption----->")
'
'    'did we find it?
'    If wnd1 = 0 Then
'        MsgBox "Program not found!"
'        Exit Sub
'
'    End If
'
'    ' reset our child hwnd count & clear array of the child Hwnds.
'    ChildCount = 0
'    Erase NextHwnd()
'
'    ' Find text boxes!
'    Call FindWindowLike(0, "*", "ThunderRT6TextBox") 'Call FindWindowLike(0, "Text1", "ThunderRT6TextBox")
'
'    For i = 0 To ChildCount
'        ChldHwnd = NextHwnd(i)
'
'        If IsChild(wnd1, ChldHwnd) Then
'            NextBox = NextBox + 1
'
'            ' 1st box
'            If NextBox = 1 Then SendMessageByString ChldHwnd, WM_SETTEXT, 0, Text1.Text
'            Exit For
'
'        End If
'
'    End If
'
'Next i
'
'' reset child hwnd count
'ChildCount = 0
'Erase NextHwnd()
'
'' Find the button.
'Call FindWindowLike(0, "Command1", "ThunderRT6CommandButton")
'
'' Click the button if there is one
'For i = 0 To ChildCount
'    ChldHwnd = NextHwnd(i)
'
'    If IsChild(wnd1, ChldHwnd) Then
'        SendMessage ChldHwnd, BM_CLICK, 0, 0
'        Exit For
'
'    End If
'
'Next i
'
'End Sub
'
