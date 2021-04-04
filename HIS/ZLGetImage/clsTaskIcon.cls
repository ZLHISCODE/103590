VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTaskIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'使用方法：1.在程序中加入一个新窗体
'          2.定义窗体的ScaleMode=3
'          3.在窗体中定义该类对象
'          4.在向任务条上加图标前初始化对象的以下属性:
'              1).frmHwnd:窗体的Hwnd属性(必须)
'              2).Icon:IPictureDsp的Handle属性(必须)
'              3).Message:提示文本(可选)
'          5.调用对象的AddIcon方法
'          6.可修改对象的Icon和Message属性改变图标和提示
'          7.调用DelIcon方法删除图标
'          8.为了激活对象的事件,在窗体的的MouseMove事件中调用MouseState方法
'              参数:MouseMove事件的X参数
'注意事项：1.程序中必须有一个窗体来接收消息
'          2.必须预先设置对象的frmHwnd和Icon属性

Option Explicit
'定义消息
Private Const NIM_ADD = &H0     '添加图标到任务栏提示区。
Private Const NIM_MODIFY = &H1  '图标特性已改变的消息。
Private Const NIM_DELETE = &H2  '删除图标。
'定义标识位
Private Const NIF_MESSAGE = &H1  '消息合法。
Private Const NIF_ICON = &H2     '图标句柄合法。
Private Const NIF_TIP = &H4      '提示合法。
Private Const WM_MOUSEMOVE = &H200 '定义MouseMove消息，该消息将被发送到窗体的MouseMove事件处理函数中处理。
'定义鼠标消息常数
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private IconData As NOTIFYICONDATA
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Long

Public Event MouseLeftDown()
Public Event MouseLeftUp()
Public Event MouseLeftDBClick()
Public Event MouseRightDown()
Public Event MouseRightUp()
Public Event MouseRightDBClick()
Private blnAdd As Boolean

Private Sub Class_Initialize()
    blnAdd = False
End Sub

Public Sub MouseState(XMessage As Single)
    Select Case CLng(XMessage / Screen.TwipsPerPixelX)
        Case WM_LBUTTONDBLCLK
            RaiseEvent MouseLeftDBClick
        Case WM_LBUTTONDOWN
            RaiseEvent MouseLeftDown
        Case WM_LBUTTONUP
            RaiseEvent MouseLeftUp
        Case WM_RBUTTONDBLCLK
            RaiseEvent MouseRightDBClick
        Case WM_RBUTTONDOWN
            RaiseEvent MouseRightDown
        Case WM_RBUTTONUP
            RaiseEvent MouseRightUp
    End Select
End Sub

'返回:0=成功;1=未设置窗体Hwnd;2=未设置图标;3=失败
Public Function AddIcon() As Byte
    With IconData
        If .hwnd = 0 Then AddIcon = 1: Exit Function
        If .hIcon = 0 Then AddIcon = 2: Exit Function
        .cbSize = Len(IconData)
        .uCallbackMessage = WM_MOUSEMOVE
        .uID = &H1
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        If Shell_NotifyIcon(NIM_ADD, IconData) <> 0 Then
            AddIcon = 0: blnAdd = True
        Else
            AddIcon = 3
        End If
    End With
End Function

Private Function ModiIcon() As Boolean
    If Not blnAdd Then ModiIcon = False: Exit Function
    If Shell_NotifyIcon(NIM_MODIFY, IconData) <> 0 Then
        ModiIcon = True
    Else
        ModiIcon = False
    End If
End Function

Public Function DelIcon() As Boolean
    If Not blnAdd Then DelIcon = False: Exit Function
    If Shell_NotifyIcon(NIM_DELETE, IconData) <> 0 Then
        DelIcon = True
    Else
        DelIcon = False
    End If
End Function

'图标(Handle)
Public Property Let Icon(lngHandle As Long)
    IconData.hIcon = lngHandle
    Call ModiIcon
End Property

Public Property Get Icon() As Long
    Icon = IconData.hIcon
End Property

'提示
Public Property Let Message(strMessage As String)
    IconData.szTip = strMessage
    IconData.szTip = Trim(IconData.szTip) & Chr$(0)
    Call ModiIcon
End Property

Public Property Get Message() As String
    Message = IconData.szTip
End Property

'窗体Hwnd
Public Property Let frmHwnd(lngHwnd As Long)
    IconData.hwnd = lngHwnd
End Property

Public Property Get frmHwnd() As Long
    frmHwnd = IconData.hwnd
End Property

