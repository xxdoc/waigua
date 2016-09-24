Attribute VB_Name = "Module2"
Declare Function GetWindowText _
        Lib "user32" _
        Alias "GetWindowTextA" (ByVal hWnd As Long, _
                                ByVal lpString As String, _
                                ByVal cch As Long) As Long
Declare Function GetWindowTextLength _
        Lib "user32" _
        Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long

Public BtnHwnd As Long

Private Const LB_SETTABSTOPS = &H192

Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hWnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

Private Declare Function SendMessageByString _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hWnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As String) As Long

Dim retlength As Long

Dim retstring As String

Const WM_GETTEXT = &HD

Const WM_GETTEXTLENGTH = &HE

Declare Function EnumChildWindows _
        Lib "user32" (ByVal hWndParent As Long, _
                      ByVal lpEnumFunc As Long, _
                      ByVal lParam As Long) As Long

Public Sub SetOneTabStopInListBox(lst As ListBox, ColWidth As Long)

    Dim ColWidths(0 To 0) As Long

    '
    ColWidths(0) = ColWidth
    SetTabStopsInListbox lst, ColWidths()

End Sub

Private Sub SetTabStopsInListbox(lst As ListBox, ColWidths() As Long)

    ' A character is approximately 4 "width" units.
    ' ColWidths() is in character widths of some standard character.
    ' ColWidths() can be either zero or one based.
    ' SADLY: This does not work for listboxes with the style set to checkbox.
    Dim NumCols As Long

    '
    NumCols = UBound(ColWidths) - LBound(ColWidths) + 1 ' Calculate the number of columns.
    SendMessage lst.hWnd, LB_SETTABSTOPS, 0&, ByVal 0& ' Clear any existing tabs.
    SendMessage lst.hWnd, LB_SETTABSTOPS, NumCols, ColWidths(LBound(ColWidths)) ' Set new tabs.

End Sub

Public Function EnumChildProc(ByVal hWnd As Long, ByVal lParam As Long) As Long

    Dim sSave As String

    Dim nSize As Long

    'Get the windowtext length
    nSize = SendMessageByString(hWnd, WM_GETTEXTLENGTH, 0, 0)
    sSave = Space$(nSize + 1)
    'get the window text
    SendMessageByString hWnd, WM_GETTEXT, nSize + 1, sSave
    'GetWindowText hWnd, sSave, Len(sSave)
    'remove the last Chr$(0)
    sSave = Left$(sSave, Len(sSave) - 1)
    
    If sSave = "—°≤°»À" Then
        BtnHwnd = hWnd
        EnumChildProc = 0
        Exit Function

    End If

    'continue enumeration
    EnumChildProc = 1

End Function

