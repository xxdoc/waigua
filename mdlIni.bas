Attribute VB_Name = "mdlIni"
Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Const IMAGE_ICON            As Long = 1
Private Const ICON_BIG              As Long = 1
Private Const ICON_SMALL            As Long = 0
Private Const WM_SETICON            As Long = &H80
Private Const LR_COPYFROMRESOURCE   As Long = &H4000
Private Const SW_SHOWNORMAL         As Long = 1

Private mFilename As String
Public Sub SetWindowIcon(hwnd As Long, ByVal sIconName As String)
    Dim hIcon As Long
    hIcon = LoadImage(App.hInstance, sIconName, IMAGE_ICON, 32, 32, LR_COPYFROMRESOURCE)
    If hIcon Then Call SendMessage(hwnd, WM_SETICON, ICON_BIG, ByVal hIcon)
    hIcon = LoadImage(App.hInstance, sIconName, IMAGE_ICON, 16, 16, LR_COPYFROMRESOURCE)
    If hIcon Then Call SendMessage(hwnd, WM_SETICON, ICON_SMALL, ByVal hIcon)
End Sub


Public Property Let FileName(ByVal sFilename As String)
    mFilename = sFilename
End Property

Public Property Get FileName() As String
    FileName = mFilename
End Property

Public Function ReadData(ByVal sSection As String, ByVal sKeyName As String, Optional sDefault As String = vbNullString) As String
    Dim sBufer As String * 256, lRet As Long
    lRet = GetPrivateProfileString(sSection, sKeyName, sDefault, sBufer, Len(sBufer), mFilename)
   ' ReadData = StrConv(Left$(StrConv(sBufer, vbFromUnicode), lRet), vbUnicode) '×ª»»Îª×Ö·û´®
    ReadData = Left$(sBufer, InStr(1, sBufer, vbNullChar) - 1)
End Function

Public Function WriteData(ByVal sSection As String, ByVal sKeyName As String, ByVal sValue As String) As Boolean
      WriteData = WritePrivateProfileString(sSection, sKeyName, sValue, mFilename)
End Function

Public Function SaveFormPosAndSize(Frm As Form)
    With Frm
        If .WindowState <> vbMinimized Then
            WriteData .Name, "Left", .Left
            WriteData .Name, "Top", .Top
            WriteData .Name, "Width", .Width
            WriteData .Name, "Height", .Height
            WriteData .Name, "WindowState", .WindowState
        End If
    End With
End Function

Public Function ReadFormPosAndSize(Frm As Form)

    With Frm
    
        If ReadData(.Name, "WindowState", vbNormal) Then
            .WindowState = vbMaximized
        Else
            .Left = ReadData(.Name, "Left", .Left)
            .Top = ReadData(.Name, "Top", .Top)
            .Width = ReadData(.Name, "Width", .Width)
            .Height = ReadData(.Name, "Height", .Height)
        End If
        
    End With

End Function
