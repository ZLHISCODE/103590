Attribute VB_Name = "Rais"
Option Explicit
Public Const TOGGLE_HIDEWINDOW = &H80
Public Const TOGGLE_UNHIDEWINDOW = &H40

Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
'���ô�����ʾ����
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
'����PictureBox����ʾ״̬
Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'��С������
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long

Public gConnect As New ADODB.Connection  '��������
Public gLngFormID As Long '�����������
Public CollMap As Collection
Public RsAllPic As New ADODB.Recordset
Public BlnExistBill As Boolean

Public Function InitPicToRead()
    '������:����
    '��������:2000-12-12
    'Ϊ����ٶ�,�Ȱ�ͼƬ�ӿ��ж���,����ͼƬ�ļ��б�

    Dim StrFilePath As String
    If BlnExistBill = False Then Exit Function
    With RsAllPic
        Do While Not .EOF
            StrFilePath = Rec.DownloadPicture(RsAllPic, "ͼƬ")
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
End Function
