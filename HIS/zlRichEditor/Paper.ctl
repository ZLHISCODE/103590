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
'##ģ �� ����Paper.ctl
'##�� �� �ˣ�����ΰ
'##��    �ڣ�2005��5��1��
'##�� �� �ˣ�
'##��    �ڣ�
'##��    ����ҳ����ͼ�Ļ����༭�ؼ���
'##��    ����
'######################################################################################

Option Explicit

'#############################################################################################################
'##     �ֲ�����
'#############################################################################################################

Private m_hWnd As Long              '�ؼ��� hWnd
Private m_hWndRTB As Long           'RTB�� hWnd
Private m_hWndParent  As Long       '������� hWnd

Private m_bSubClassing As Boolean   '�Ƿ�̳�����

'#############################################################################################################
'##     ��������
'#############################################################################################################

Private mvarRowStart   As Long          '��ʼ�� ��������
Private mvarRowEnd     As Long          '��ֹ�� ��������
Private mvarCharStart  As Long          '��ʼ�ַ�λ��
Private mvarCharEnd    As Long          '��ֹ�ַ�λ��
Private mvarPageNumber As Long          '��ǰҳ�� ��������
Private mvarRequestLine As Boolean      '�Ƿ�������RequestLine�¼�

'#############################################################################################################
'##     �¼�����
'#############################################################################################################

Public Event Change()   '���ݸı䣡
Public Event MouseWheel(bBackDirection As Boolean, Shift As Integer, X As Single, Y As Single, Value As Single)    '�������¼�
Public Event Zoom(NewFactor As Double)    '�û�ͨ��Ctrl��������ı������ű�����
Public Event Resize()    '�ؼ��ߴ�ı�
Public Event RequestLine()              '���������ı�
Public Event SelChange(ByVal lStart As Long, ByVal lEnd As Long)   'ѡ������ı�
Public Event LinkEvent(ByVal iType As LinkEventTypeEnum, ByVal lStart As Long, ByVal lEnd As Long)      '�����¼�
Public Event ModifyProtected(ByRef bAllowDoIt As Boolean, ByVal lStart As Long, ByVal lEnd As Long, KeyAscii As Integer, Shift As Integer)     '��ͼ�༭�ܱ�������
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
Public Event Click()        '����
Public Event DblClick()     '˫��
Public Event HeadFootChanged()

'#############################################################################################################
'##     ��������
'#############################################################################################################

Public Property Get objPaper() As Object
    '����ͼƬ�������ڴ�ӡԤ��
    Set objPaper = picPaper
End Property

Public Property Get hDC() As Long
    hDC = picPaper.hDC
End Property

Private Sub pAttachMessages()
'��Ϣ�����
    Subclass1.Hwnd = UserControl.Hwnd
    Subclass1.Messages(WM_MOUSEWHEEL) = True
    m_bSubClassing = True
End Sub

Private Sub pDetachMessages()
'ȡ����Ϣ����
    m_bSubClassing = False
End Sub

Private Sub pInitialise()
'�����ʼ��
    pTerminate
    If (UserControl.Ambient.UserMode) Then
        m_hWnd = UserControl.Hwnd
        m_hWndParent = UserControl.Parent.Hwnd
        
        Call pAttachMessages     '��Ϣ��
    End If
End Sub

Private Function pTerminate()
'���پ��
    Call pDetachMessages          'ȡ����Ϣ��
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
'���Ʊ߽���� ���������ű�����
    On Error Resume Next
    Dim X1 As Long, Y1 As Long, X2 As Long, Y2 As Long
'    picPaper.Cls
    '���Ͻ�
    X1 = PubInfo.MarginLeft * PubInfo.ZoomFactor - Screen.TwipsPerPixelX
    Y1 = PubInfo.MarginTop * PubInfo.ZoomFactor - Screen.TwipsPerPixelY
    X2 = -360 * PubInfo.ZoomFactor
    Y2 = 0
    picPaper.Line (X1, Y1)-Step(X2, Y2), RGB(166, 166, 166)
    X2 = 0
    Y2 = -360 * PubInfo.ZoomFactor
    picPaper.Line (X1, Y1)-Step(X2, Y2), RGB(166, 166, 166)
    '���Ͻ�
    X1 = ScaleWidth - PubInfo.MarginRight * PubInfo.ZoomFactor + Screen.TwipsPerPixelX * 4
    Y1 = PubInfo.MarginTop * PubInfo.ZoomFactor - Screen.TwipsPerPixelY
    X2 = 360 * PubInfo.ZoomFactor
    Y2 = 0
    picPaper.Line (X1, Y1)-Step(X2, Y2), RGB(166, 166, 166)
    X2 = 0
    Y2 = -360 * PubInfo.ZoomFactor
    picPaper.Line (X1, Y1)-Step(X2, Y2), RGB(166, 166, 166)
    '���½�
    X1 = PubInfo.MarginLeft * PubInfo.ZoomFactor - Screen.TwipsPerPixelX
    Y1 = ScaleHeight - PubInfo.MarginBottom * PubInfo.ZoomFactor + Screen.TwipsPerPixelY * 4
    X2 = -360 * PubInfo.ZoomFactor
    Y2 = 0
    picPaper.Line (X1, Y1)-Step(X2, Y2), RGB(166, 166, 166)
    X2 = 0
    Y2 = 360 * PubInfo.ZoomFactor
    picPaper.Line (X1, Y1)-Step(X2, Y2), RGB(166, 166, 166)
    '���½�
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
'##     �ڲ��ؼ��¼�
'#############################################################################################################

Private Sub UserControl_Initialize()
'�ڳ��򴴽��ؼ�������ʱʱ����
    lblThis.Caption = "(ҳ����ͼ)"
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
        '�������ָ��
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
'���ٿؼ�ʱ����
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
        'wzDelta���ݹ��ֹ����Ŀ�������ֵС�����ʾ���������������û����򣩣�
        '�������ʾ������ǰ����������ʾ������
        wzDelta = HIWORD(wParam)
        'wKeysָ���Ƿ���CTRL=8��SHIFT=4������(��=2����=16����=2������)���£�������
        wKeys = LOWORD(wParam)
        tP.X = LOWORD(lParam)    'pt��������
        tP.Y = HIWORD(lParam)
        '--------------------------------------------------
        If wzDelta < 0 Then  '���û�����
           bWay = True
        Else                 '����ʾ������
           bWay = False
        End If
        '--------------------------------------------------
        '����Ļ����ת��ΪForm1.��������
        ScreenToClient Hwnd, tP
        sngX = tP.X
        sngY = tP.Y
        intShift = wKeys
        bMouseFlag = True  '�ù�����־
        If bMouseFlag = True Then
            bMouseFlag = False
            RaiseEvent MouseWheel(bWay, intShift, sngX, sngY, CLng(wzDelta)) '�����¼�
        End If
    End If
End Sub
