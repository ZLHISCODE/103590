Attribute VB_Name = "Rais"
Option Explicit
Public Const TOGGLE_HIDEWINDOW = &H80
Public Const TOGGLE_UNHIDEWINDOW = &H40

Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
'设置窗体显示区域
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
'设置PictureBox的显示状态
Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'最小化窗体
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long

Public gConnect As New ADODB.Connection  '公共链接
Public gLngFormID As Long '主窗体的名柄
Public CollMap As Collection
Public RsAllPic As New ADODB.Recordset
Public BlnExistBill As Boolean

Public Function InitPicToRead()
    '编制人:朱玉宝
    '编制日期:2000-12-12
    '为提高速度,先把图片从库中读出,产生图片文件列表

    Dim StrFilePath As String
    If BlnExistBill = False Then Exit Function
    With RsAllPic
        Do While Not .EOF
            StrFilePath = Rec.DownloadPicture(RsAllPic, "图片")
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
End Function
