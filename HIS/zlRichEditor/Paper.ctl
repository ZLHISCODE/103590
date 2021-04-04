VERSION 5.00
Object = "{09B13292-AC31-4C5D-B44A-C83E7AAD70E6}#1.1#0"; "zlSubclass.ocx"
Begin VB.UserControl Paper 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   2025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4755
   LockControls    =   -1  'True
   MouseIcon       =   "Paper.ctx":0000
   ScaleHeight     =   2025
   ScaleWidth      =   4755
   ToolboxBitmap   =   "Paper.ctx":0152
   Begin zlSubclass.Subclass Subclass1 
      Left            =   3825
      Top             =   945
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.PictureBox picPaper 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1770
      Left            =   360
      ScaleHeight     =   1740
      ScaleWidth      =   3270
      TabIndex        =   0
      Top             =   225
      Width           =   3300
      Begin VB.Label lblPageNumber 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H009F9F9F&
         Height          =   240
         Left            =   2160
         TabIndex        =   2
         Top             =   1440
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1050
      End
   End
   Begin VB.Label lblThis 
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   3930
   End
End
Attribute VB_Name = "Paper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'######################################################################################
'##模 块 名：Paper.ctl
'##创 建 人：吴庆伟
'##日    期：2005年5月1日
'##修 改 人：
'##日    期：
'##描    述：页面视图的基础编辑控件。
'##版    本：
'######################################################################################

Option Explicit

'#############################################################################################################
'##     局部变量
'#############################################################################################################

Private m_hWnd As Long              '控件的 hWnd
Private m_hWndRTB As Long           'RTB的 hWnd
Private m_hWndParent  As Long       '父窗体的 hWnd

Private m_bSubClassing As Boolean   '是否继承子类

'#############################################################################################################
'##     独立属性
'#############################################################################################################

Private mvarRowStart   As Long          '起始行 独有属性
Private mvarRowEnd     As Long          '终止行 独有属性
Private mvarCharStart  As Long          '起始字符位置
Private mvarCharEnd    As Long          '终止字符位置
Private mvarPageNumber As Long          '当前页码 独有属性
Private mvarRequestLine As Boolean      '是否允许触发RequestLine事件

'#############################################################################################################
'##     事件声明
'#############################################################################################################

Public Event Change()   '内容改变！
Public Event MouseWheel(bBackDirection As Boolean, Shift As Integer, X As Single, Y As Single, Value As Single)    '鼠标滚轮事件
Public Event Zoom(NewFactor As Double)    '用户通过Ctrl＋鼠标来改变了缩放比例！
Public Event Resize()    '控件尺寸改变
Public Event RequestLine()              '请求行数改变
Public Event SelChange(ByVal lStart As Long, ByVal lEnd As Long)   '选择区域改变
Public Event LinkEvent(ByVal iType As LinkEventTypeEnum, ByVal lStart As Long, ByVal lEnd As Long)      '链接事件
Public Event ModifyProtected(ByRef bAllowDoIt As Boolean, ByVal lStart As Long, ByVal lEnd As Long, KeyAscii As Integer, Shift As Integer)     '试图编辑受保护区域
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event RTBMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event RTBMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event RTBMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event RequestRightMenu(ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Public Event Click()        '单击
Public Event DblClick()     '双击
Public Event HeadFootChanged()

'#############################################################################################################
'##     公共属性
'#############################################################################################################

Public Property Get objPaper() As Object
    '返回图片对象，用于打印预览
    Set objPaper = picPaper
End Property

Public Property Get hDC() As Long
    hDC = picPaper.hDC
End Property

Private Sub pAttachMessages()
'消息捕获绑定
    Subclass1.Hwnd = UserControl.Hwnd
    Subclass1.Messages(WM_MOUSEWHEEL) = True
    m_bSubClassing = True
End Sub

Private Sub pDetachMessages()
'取消消息捕获
    m_bSubClassing = False
End Sub

Private Sub pInitialise()
'句柄初始化
    pTerminate
    If (UserControl.Ambient.UserMode) Then
        m_hWnd = UserControl.Hwnd
        m_hWndParent = UserControl.Parent.Hwnd
        
        Call pAttachMessages     '消息绑定
    End If
End Sub

Private Function pTerminate()
'销毁句柄
    Call pDetachMessages          '取消消息绑定
End Function

Public Property Get PageNumber() As Long
    PageNumber = mvarPageNumber
End Property

Public Property Let PageNumber(vData As Long)
    mvarPageNumber = vData
    lblPageNumber = "- " & vData & " -"
    PropertyChanged "PageNumber"
End Property

Public Sub DrawBorder()
'绘制边界符号 （按照缩放比例）
    On Error Resume Next
    Dim X1 As Long, Y1 As Long, X2 As Long, Y2 As Long
'    picPaper.Cls
    '左上角
    X1 = PubInfo.MarginLeft * PubInfo.ZoomFactor - Screen.TwipsPerPixelX
    Y1 = PubInfo.MarginTop * PubInfo.ZoomFactor - Screen.TwipsPerPixelY
    X2 = -360 * PubInfo.ZoomFactor
    Y2 = 0
    picPaper.Line (X1, Y1)-Step(X2, Y2), RGB(166, 166, 166)
    X2 = 0
    Y2 = -360 * PubInfo.ZoomFactor
    picPaper.Line (X1, Y1)-Step(X2, Y2), RGB(166, 166, 166)
    '右上角
    X1 = ScaleWidth - PubInfo.MarginRight * PubInfo.ZoomFactor + Screen.TwipsPerPixelX * 4
    Y1 = PubInfo.MarginTop * PubInfo.ZoomFactor - Screen.TwipsPerPixelY
    X2 = 360 * PubInfo.ZoomFactor
    Y2 = 0
    picPaper.Line (X1, Y1)-Step(X2, Y2), RGB(166, 166, 166)
    X2 = 0
    Y2 = -360 * PubInfo.ZoomFactor
    picPaper.Line (X1, Y1)-Step(X2, Y2), RGB(166, 166, 166)
    '左下角
    X1 = PubInfo.MarginLeft * PubInfo.ZoomFactor - Screen.TwipsPerPixelX
    Y1 = ScaleHeight - PubInfo.MarginBottom * PubInfo.ZoomFactor + Screen.TwipsPerPixelY * 4
    X2 = -360 * PubInfo.ZoomFactor
    Y2 = 0
    picPaper.Line (X1, Y1)-Step(X2, Y2), RGB(166, 166, 166)
    X2 = 0
    Y2 = 360 * PubInfo.ZoomFactor
    picPaper.Line (X1, Y1)-Step(X2, Y2), RGB(166, 166, 166)
    '右下角
    X1 = ScaleWidth - PubInfo.MarginRight * PubInfo.ZoomFactor + Screen.TwipsPerPixelX * 4
    Y1 = ScaleHeight - PubInfo.MarginBottom * PubInfo.ZoomFactor + Screen.TwipsPerPixelY * 4
    X2 = 360 * PubInfo.ZoomFactor
    Y2 = 0
    picPaper.Line (X1, Y1)-Step(X2, Y2), RGB(166, 166, 166)
    X2 = 0
    Y2 = 360 * PubInfo.ZoomFactor
    picPaper.Line (X1, Y1)-Step(X2, Y2), RGB(166, 166, 166)
End Sub

'#############################################################################################################
'##     内部控件事件
'#############################################################################################################

Private Sub UserControl_Initialize()
'在程序创建控件及运行时时发生
    lblThis.Caption = "(页面视图)"
    picPaper.Move 0, 0
End Sub

Private Sub UserControl_InitProperties()
    ShowPageNumber = PubInfo.ShowPageNumber
    PageNumber = 1
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If X < PubInfo.MarginLeft * PubInfo.ZoomFactor Then
        '反向鼠标指针
        UserControl.MousePointer = 99
    Else
        UserControl.MousePointer = vbIbeam
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    PageNumber = PropBag.ReadProperty("PageNumber", 1)
    ShowPageNumber = PropBag.ReadProperty("ShowPageNumber", False)
    If Ambient.UserMode Then
        Call pInitialise
    End If
End Sub

Private Sub UserControl_Resize()
    lblPageNumber.Move PubInfo.PaperWidth - 800 * PubInfo.ZoomFactor, _
        PubInfo.PaperHeight - 400 * PubInfo.ZoomFactor
    picPaper.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    RaiseEvent Resize
End Sub

Public Property Get Enabled() As Boolean
    Enabled = PubInfo.Enabled
End Property

Public Property Let Enabled(ByVal vData As Boolean)
    PubInfo.Enabled = vData
    UserControl.Enabled = vData
    PropertyChanged "Enabled"
End Property

Private Sub UserControl_Show()
    If Ambient.UserMode Then
        lblThis.Visible = False
        picPaper.Visible = True
    Else
        lblThis.Visible = False
        picPaper.Visible = True
    End If
End Sub

Private Sub UserControl_Terminate()
'销毁控件时发生
    Call pTerminate
End Sub

Public Property Get ShowPageNumber() As Boolean
    ShowPageNumber = PubInfo.ShowPageNumber
End Property

Public Property Let ShowPageNumber(vData As Boolean)
    PubInfo.ShowPageNumber = vData
    If Ambient.UserMode Then lblPageNumber.Visible = vData
    PropertyChanged "ShowPageNumber"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "PageNumber", PageNumber, 1
    PropBag.WriteProperty "ShowPageNumber", ShowPageNumber, False
    PropertyChanged "PageNumber"
    PropertyChanged "ShowPageNumber"
End Sub

Private Sub Subclass1_WndProc(Msg As Long, wParam As Long, lParam As Long, Result As Long)
    If Msg = WM_MOUSEWHEEL Then
        Dim tP As POINTAPI
        Dim wzDelta, wKeys As Integer
        'wzDelta传递滚轮滚动的快慢，该值小于零表示滚轮向后滚动（朝用户方向），
        '大于零表示滚轮向前滚动（朝显示器方向）
        wzDelta = HIWORD(wParam)
        'wKeys指出是否有CTRL=8、SHIFT=4、鼠标键(左=2、中=16、右=2、附加)按下，允许复合
        wKeys = LOWORD(wParam)
        tP.X = LOWORD(lParam)    'pt鼠标的坐标
        tP.Y = HIWORD(lParam)
        '--------------------------------------------------
        If wzDelta < 0 Then  '朝用户方向
           bWay = True
        Else                 '朝显示器方向
           bWay = False
        End If
        '--------------------------------------------------
        '将屏幕坐标转换为Form1.窗口坐标
        ScreenToClient Hwnd, tP
        sngX = tP.X
        sngY = tP.Y
        intShift = wKeys
        bMouseFlag = True  '置滚动标志
        If bMouseFlag = True Then
            bMouseFlag = False
            RaiseEvent MouseWheel(bWay, intShift, sngX, sngY, CLng(wzDelta)) '激活事件
        End If
    End If
End Sub
