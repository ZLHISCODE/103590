VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNotify 
   BorderStyle     =   0  'None
   Caption         =   "ҽ������"
   ClientHeight    =   8070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4530
   Icon            =   "frmNotify.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmNotify"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Time_Flash 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1590
      Top             =   60
   End
   Begin VB.Timer TimNotify 
      Interval        =   500
      Left            =   2010
      Top             =   60
   End
   Begin MSComctlLib.ImageList imgPati 
      Left            =   3630
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNotify.frx":000C
            Key             =   "Pati"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNotify.frx":05A6
            Key             =   "Notify"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNotify.frx":0B40
            Key             =   "warn"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNotify.frx":0EDA
            Key             =   "Out"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNotify.frx":1274
            Key             =   "Change"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EAFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   30
      ScaleHeight     =   345
      ScaleWidth      =   4305
      TabIndex        =   1
      Top             =   30
      Visible         =   0   'False
      Width           =   4305
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ������"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   90
         Width           =   3735
      End
      Begin VB.Image imgShow 
         Height          =   360
         Left            =   0
         Picture         =   "frmNotify.frx":160E
         Stretch         =   -1  'True
         Top             =   0
         Width           =   360
      End
   End
   Begin VB.PictureBox picForm 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EAFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8040
      Left            =   0
      ScaleHeight     =   8040
      ScaleWidth      =   4485
      TabIndex        =   0
      Top             =   0
      Width           =   4490
      Begin VB.PictureBox picNotify 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7560
         Left            =   15
         ScaleHeight     =   7560
         ScaleWidth      =   4485
         TabIndex        =   4
         Top             =   480
         Width           =   4490
         Begin XtremeReportControl.ReportControl rptNotify 
            Height          =   7290
            Left            =   0
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   255
            Width           =   4470
            _Version        =   589884
            _ExtentX        =   7885
            _ExtentY        =   12859
            _StockProps     =   0
            BorderStyle     =   1
            MultipleSelection=   0   'False
            EditOnClick     =   0   'False
            AutoColumnSizing=   0   'False
         End
         Begin VB.OptionButton optNotify 
            Appearance      =   0  'Flat
            BackColor       =   &H00EAFFFF&
            Caption         =   "ȫ����"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   2250
            TabIndex        =   7
            Top             =   15
            Value           =   -1  'True
            Width           =   870
         End
         Begin VB.OptionButton optNotify 
            Appearance      =   0  'Flat
            BackColor       =   &H00EAFFFF&
            Caption         =   "���˸���"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   1005
            TabIndex        =   6
            Top             =   15
            Width           =   1155
         End
         Begin VB.Label lblNotify 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00EAFFFF&
            Caption         =   "���ѷ�Χ��"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   135
            TabIndex        =   8
            Top             =   0
            Width           =   900
         End
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   3
         Top             =   120
         Width           =   720
      End
      Begin VB.Image imgHide 
         Height          =   360
         Left            =   30
         Picture         =   "frmNotify.frx":1D10
         Stretch         =   -1  'True
         Top             =   30
         Width           =   360
      End
   End
   Begin XtremeCommandBars.ImageManager imgIcons 
      Left            =   1350
      Top             =   150
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmNotify.frx":2412
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   1350
      Top             =   3480
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngLeft_Form As Long       '���������
Private mlngTop_Form As Long        '���������
Private mlngLeft_Mouse As Long      '�����ʱ������
Private mlngTop_Mouse As Long       '�����ʱ������

Private mintCount As Integer
Private mstrPreNotify As String
Private mlngPreID As Long
Private mblnFirstIn As Boolean      '�Ƿ��һ�ν���ҽ�����ѽ�������

Public mblnExecCollapse As Boolean  'ִ������ʱ���Ѵ����Ƿ��Զ��۵�
Public mblnNormal As Boolean        'TRUE-���;FALSE-��С��
Public mblnOrientation As Boolean   'TRUE-����;False-����
Public mblnFirst As Boolean         '��һ����������ˢ��,���л�����ʱǿ��ˢ��
Public mlng����ID As Long
Public mstrScope As String
Public mdtOutBegin As Date, mdtOutEnd As Date
Public mintNotify As Integer 'ҽ�������Զ�ˢ�¼��(����)
Public mintNotifyDay As Integer '���Ѷ������ڵ�ҽ��
Public mstrNotifyAdvice As String '���ѵ�ҽ������
Public mstrRelatedUnitID As String '���廤����ID
Public mbln���廤����Ϣ As Boolean '�����Ƿ���ʾ���廤����Ϣ

Private mstrBlankTime As String     '����ҽ��ʱ��
Private mclsMsg As clsCISMsg
Private mrsMsg As ADODB.Recordset
Private mbln��Ϣ���� As Boolean


Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML

Private mclsPublicAdvice As zlPublicAdvice.clsPublicAdvice

Private Enum NOTIFYREPORT_COLUMN
    c_ͼ�� = 0
    C_����ID = 1
    C_��ҳID = 2
    c_���� = 3
    c_סԺ�� = 4
    c_���� = 5
    C_״̬ = 6
    
    '������
    C_��Ϣ = 7
    C_��� = 8
    C_���� = 9
    C_ҵ�� = 10
    C_���ﲡ�� = 11
    C_Ψһ��ʶ = 12 '����������Ϣ��Ψһ��
End Enum

Private Enum EFun_ҽ������
    EУ�� = 0
    Eֹͣ = 1
End Enum

Private Enum Msg_Type '��Ϣ�������
    m�¿� = 1
    m��ͣ = 2
    m�·� = 3
    m���� = 4
    mΣ��ֵ = 5
    m��Һ�ܾ� = 6
    m�������� = 7
    mȡѪ֪ͨ = 8
    m��Ѫ��� = 9
End Enum

Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Const GWL_EXSTYLE = (-20)

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll " (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

'����
Private Const lngWidth_Normal As Long = 4530
Private Const lngHeight_Normal As Long = 8070
Private Const lngWidthH_Collapse As Long = 4020
Private Const lngHeightH_Collapse As Long = 410
'����
Private Const lngWidthV_Collapse As Long = 410
Private Const lngHeightV_Collapse As Long = 4020

Private Const conMenu_��С�� As Long = 15
Private Const conMenu_���� As Long = 14
Private Const conMenu_���� As Long = 13
Private Const conMenu_�۵� As Long = 12
Private Const conMenu_չ�� As Long = 11
Private mobjMenu As CommandBarPopup

Private mbytFontSize As Byte

Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С ��Form_Load֮�����
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-20 15:15:00
    '����:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize
End Sub

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-20 15:15:00
    '����:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim bytSize As Byte
    bytSize = IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))
    
    Me.FontSize = mbytFontSize
    Me.FontName = "����"
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("Label")
            objCtrl.FontSize = mbytFontSize
            If UCase(objCtrl.Name) = UCase("lblInfo") Then
                If mblnOrientation = True Then
                    objCtrl.Height = TextHeight("��") + 20
                Else
                    objCtrl.Width = TextWidth("��") + 20
                End If
            Else
                objCtrl.Height = TextHeight("��") + 20
            End If
        Case UCase("OptionButton")
            objCtrl.FontSize = mbytFontSize
            objCtrl.Width = TextWidth(objCtrl.Caption & "����")
            objCtrl.Height = TextHeight("��") + 20
            
        Case UCase("ReportControl")
            Set CtlFont = objCtrl.PaintManager.CaptionFont
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
            
            Set CtlFont = objCtrl.PaintManager.TextFont
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.TextFont = CtlFont
            objCtrl.Redraw
        Case UCase("CommandBars")
            Set CtlFont = objCtrl.Options.Font
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = mbytFontSize
            Set objCtrl.Options.Font = CtlFont
        End Select
    Next
    optNotify(0).Left = lblNotify.Left + lblNotify.Width
    optNotify(1).Left = optNotify(0).Left + optNotify(0).Width + 100
    Call Form_Resize
End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    mlngLeft_Form = Me.Left
    mlngTop_Form = Me.Top
    
    Select Case Control.ID
    Case conMenu_��С��
        mblnExecCollapse = mblnExecCollapse Xor True
    Case conMenu_����
        Call AdjustInfo
        mblnOrientation = False
    Case conMenu_����
        Call AdjustInfo
        mblnOrientation = True
    Case conMenu_�۵�
        mblnNormal = False
    Case conMenu_չ��
        mblnNormal = True
    Case conMenu_View_Refresh
        mblnFirst = True
        Exit Sub
    End Select
    Call SetMode
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_��С��
        Control.Checked = mblnExecCollapse
    Case conMenu_����
        Control.Checked = Not mblnOrientation
    Case conMenu_����
        Control.Checked = mblnOrientation
    Case conMenu_�۵�
        Control.Checked = Not mblnNormal
    Case conMenu_չ��
        Control.Checked = mblnNormal
    End Select
End Sub

'ʾ��:
'����dwFlags��LWA_ALPHA��LWA_COLORKEY
'LWA_ALPHA�����õĻ�,ͨ��bAlpha����͸����.
'LWA_COLORKEY�����õĻ� , ��ָ����͸��������ɫΪcrKey, ������ɫ��������ʾ
'���ֻҪ����LWA_COLORKEY��crKey�����ҽ����屳��ɫ�Ϳؼ���ɫ��Ϊ   ��ͬ����ɫ���Ϳ�������¥����Ҫ�󣬾�ʵ�ʲ��Կ���
'Dim rtn As Long
'rtn = GetWindowLong(hwnd, GWL_EXSTYLE) 'ȡ�Ĵ���ԭ�ȵ���ʽ
'rtn = rtn Or WS_EX_LAYERED 'ʹ����������µ���ʽWS_EX_LAYERED
'SetWindowLong hwnd, GWL_EXSTYLE, rtn '���µ���ʽ��������
'SetLayeredWindowAttributes hwnd, Ҫ͸������ɫ, 0, LWA_COLORKEY
'SetLayeredWindowAttributes hwnd, 0, ͸����, LWA_ALPHA


'��������Ҳ�С�ڴ�����ʱ,�۵�ʱ��������ʾ

Private Sub Form_Load()
    Dim strCoord As String
    Dim objCol As ReportColumn
    Dim RectState As RECT
    Dim blnStateLR As Boolean  '״̬���Ƿ�������ʾ
    
    mintCount = 0
    mblnFirst = True
    mblnFirstIn = True
    mstrPreNotify = ""
    mlngPreID = 0
    mstrBlankTime = Format(Now, "MM-dd")
    
    '��ȡ״̬����λ��
    On Error Resume Next
    GetWindowRect FindWindow("Shell_TrayWnd", vbNullString), RectState
    blnStateLR = ((RectState.Bottom - RectState.Top) * Screen.TwipsPerPixelY = Screen.Height)
    err.Clear
    On Error GoTo ErrHand
    '������
    Set mclsPublicAdvice = New zlPublicAdvice.clsPublicAdvice
    Call mclsPublicAdvice.InitCommon(gcnOracle, glngSys)
    '������Ϣ����
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, 1265, gstrPrivs)
    Call AddMipModule(mclsMipModule)
    Set mclsXML = New zl9ComLib.clsXML
    
    mblnExecCollapse = (GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name, "AutoCollapse", "1") = 1)
    'ȡ���ڷ���
    mblnOrientation = (GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name, "Orient", "1") = 1)
    '����Ǻ����򲻴���,��������:��Ϊ��,��Ϊ��
    If Not mblnOrientation Then
        Call AdjustInfo
    End If
    
    '�����ô����Сǰ��ȡ�ϴ��˳�ʱ�����λ��
    strCoord = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name, "Coord", "-8000|300")
    mlngLeft_Form = Split(strCoord, "|")(0)
    mlngTop_Form = Split(strCoord, "|")(1)
    
    '79247
    If mlngLeft_Form + Me.Width < 0 Then  'ȱʡ��ʾ�����Ͻ�(��Ҫ��Ϊ����ʹ�õ��û�)
        mlngLeft_Form = Screen.Width - Me.Width - IIf(blnStateLR = True, IIf(RectState.Left = 0, 0, ((RectState.Right - RectState.Left) * Screen.TwipsPerPixelX)), 0)
    ElseIf mlngLeft_Form < IIf(blnStateLR = True, IIf(RectState.Left = 0, ((RectState.Right - RectState.Left) * Screen.TwipsPerPixelX), 0), 0) Then
        mlngLeft_Form = IIf(blnStateLR = True, IIf(RectState.Left = 0, ((RectState.Right - RectState.Left) * Screen.TwipsPerPixelX), 0), 0)
    ElseIf mlngLeft_Form + IIf(mblnOrientation = True, lngWidthH_Collapse, lngWidthV_Collapse + 100) > Screen.Width - IIf(blnStateLR = True, IIf(RectState.Left = 0, 0, ((RectState.Right - RectState.Left) * Screen.TwipsPerPixelX)), 0) Then
        mlngLeft_Form = Screen.Width - IIf(mblnOrientation = True, lngWidthH_Collapse, lngWidthV_Collapse + 100) - IIf(blnStateLR = True, IIf(RectState.Left = 0, 0, ((RectState.Right - RectState.Left) * Screen.TwipsPerPixelX)), 0)
    End If
    
    If mlngTop_Form < IIf(blnStateLR = False, IIf(RectState.Top = 0, ((RectState.Bottom - RectState.Top) * Screen.TwipsPerPixelY), 0), 0) Then
        mlngTop_Form = IIf(blnStateLR = False, IIf(RectState.Top = 0, ((RectState.Bottom - RectState.Top) * Screen.TwipsPerPixelY), 300), 300)
    ElseIf mlngTop_Form + IIf(mblnOrientation = True, lngHeightH_Collapse, lngHeightV_Collapse) > Screen.Height - IIf(blnStateLR = False, IIf(RectState.Top = 0, 0, ((RectState.Bottom - RectState.Top) * Screen.TwipsPerPixelY)), 0) Then
        mlngTop_Form = Screen.Height - IIf(mblnOrientation = True, lngHeightH_Collapse, lngHeightV_Collapse) - IIf(blnStateLR = False, IIf(RectState.Top = 0, 0, ((RectState.Bottom - RectState.Top) * Screen.TwipsPerPixelY)), 0)
    End If
    '���ô���Ĵ�С����ʾ״̬
    Call SetMode
    
    If Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name, "NotifyRang", "0")) = 0 Then
        optNotify(1).Value = True
    Else
        optNotify(0).Value = True
    End If
    
    '90278:���մ�λ����
    With rptNotify
        Set objCol = .Columns.Add(c_ͼ��, "", 18, False): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(C_����ID, "����ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_��ҳID, "��ҳID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(c_����, "����", 60, True)
        Set objCol = .Columns.Add(c_סԺ��, "סԺ��", 70, True)
        Set objCol = .Columns.Add(c_����, "����", 60, True)
        Set objCol = .Columns.Add(C_״̬, "״̬", 150, True)
        
        Set objCol = .Columns.Add(C_��Ϣ, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_���, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_����, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_ҵ��, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_���ﲡ��, "", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_Ψһ��ʶ, "", 0, False): objCol.Visible = False
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
            objCol.Sortable = False
        Next
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .HideSelection = True
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û����������..."
        End With
        .PreviewMode = False
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .SetImageList Me.imgPati
        
        '����
        '95547:����ҽ��������ʾ
        .SortOrder.Add .Columns.Find(C_���)
        .SortOrder(0).SortAscending = False
        .SortOrder.Add .Columns.Find(c_����)
        .SortOrder(1).SortAscending = True
        .SortOrder.Add .Columns.Find(C_����)
        .SortOrder(2).SortAscending = False
    End With

    Call MainDefCommandBar
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub AdjustInfo()
    Dim lngTmp As Long
    
    lngTmp = picInfo.Width
    picInfo.Width = picInfo.Height
    picInfo.Height = lngTmp
    
    lngTmp = lblInfo.Width
    lblInfo.Width = lblInfo.Height
    lblInfo.Height = lngTmp
    lngTmp = lblInfo.Top
    lblInfo.Top = lblInfo.Left
    lblInfo.Left = lngTmp
End Sub

Private Sub MainDefCommandBar()
    Dim objControl As CommandBarControl

    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    cbsMain.Icons = imgIcons.Icons

    '�˵�����
    '-----------------------------------------------------
    Set mobjMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 1, "����(&O)", -1, False) '����
    mobjMenu.ID = 1 '��xtpControlPopup���͵�����ID�����¸�ֵ
    mobjMenu.Visible = False
    With mobjMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_����, "����")   '��ӡ��ͷ��
        Set objControl = .Add(xtpControlButton, conMenu_����, "����")
        Set objControl = .Add(xtpControlButton, conMenu_չ��, "չ��"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_�۵�, "�۵�")
        Set objControl = .Add(xtpControlButton, conMenu_��С��, "ִ������ʱ�����۵�"): objControl.BeginGroup = True
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picNotify.ZOrder 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '����ʱ�����ϴε�λ��
    If mblnNormal Then
        mlngLeft_Form = Me.Left
        mlngTop_Form = Me.Top
    End If
    
    If Not (mclsPublicAdvice Is Nothing) Then
        Set mclsPublicAdvice = Nothing
    End If
    'ж����Ϣ����
    If Not (mclsMipModule Is Nothing) Then
        Call mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    If Not (mclsXML Is Nothing) Then
        Set mclsXML = Nothing
    End If

    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name, "AutoCollapse", IIf(mblnExecCollapse, "1", "0"))
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name, "Orient", IIf(mblnOrientation, "1", "0"))
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name, "Coord", mlngLeft_Form & "|" & mlngTop_Form)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name, "NotifyRang", IIf(optNotify(0).Value = True, 1, 0))
    Set mclsMsg = Nothing
    Set mrsMsg = Nothing
End Sub

Private Sub imgHide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    '����ʱ�����ϴε�λ��
    mlngLeft_Form = Me.Left
    mlngTop_Form = Me.Top
    
    '���ô�����ʾģʽ
    mblnNormal = False
    Call SetMode
End Sub

Private Sub imgShow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picinfo_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub imgShow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picinfo_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picinfo_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lblInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picinfo_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picForm_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picForm_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub mclsMipModule_ReceiveMessage(ByVal strMsgItemIdentity As String, ByVal strMsgContent As String)
    Dim blnRecToLis As Boolean '�Ƿ���ص������б���
    Dim rsMsg As ADODB.Recordset
    On Error GoTo ErrHand
    
    If mlng����ID = 0 Then Exit Sub
    
    If strMsgItemIdentity = "ZLHIS_TRANSFUSION_001" And Mid(mstrNotifyAdvice, 6, 1) = "1" Then
        blnRecToLis = True
    ElseIf strMsgItemIdentity = "ZLHIS_CIS_001" And Mid(mstrNotifyAdvice, 1, 1) = "1" Then
        blnRecToLis = True
    ElseIf strMsgItemIdentity = "ZLHIS_CIS_002" And Mid(mstrNotifyAdvice, 2, 1) = "1" Then
        blnRecToLis = True
    ElseIf strMsgItemIdentity = "ZLHIS_CIS_003" And Mid(mstrNotifyAdvice, 3, 1) = "1" Then
        blnRecToLis = True
    ElseIf strMsgItemIdentity = "ZLHIS_CHARGE_001" And Mid(mstrNotifyAdvice, 7, 1) = "1" Then
        blnRecToLis = True
    ElseIf InStr(",ZLHIS_OPER_001,ZLHIS_CIS_005,ZLHIS_CIS_015,", "," & strMsgItemIdentity & ",") > 0 And Mid(mstrNotifyAdvice, 4, 1) = "1" Then
        blnRecToLis = True
    ElseIf InStr(",ZLHIS_LIS_003,ZLHIS_PACS_005,", "," & strMsgItemIdentity & ",") > 0 And Mid(mstrNotifyAdvice, 5, 1) = "1" Then
        blnRecToLis = True
    End If
    
    If blnRecToLis Then
        Set rsMsg = zlDatabase.ParseXMLToRecord(strMsgItemIdentity, strMsgContent)
        If rsMsg Is Nothing Then Exit Sub
        Call AddMsgToLis(rsMsg)
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub optNotify_Click(Index As Integer)
    '������غ�ŵ����ˢ��, �����ʼ�����ø��¼�
    If picNotify.Visible = True Then mblnFirst = True
End Sub

Private Sub picForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picForm.Tag = 1
        mlngLeft_Mouse = X
        mlngTop_Mouse = Y
    Else
        Call AddMenus
    End If
End Sub

Private Sub AddMenus()
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    
    '��װ�Ҽ��˵�
    Set cbrPopupBar = cbsMain.Add("�����˵�", xtpBarPopup)
    Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_����, "����"): cbrPopupItem.Visible = True
    Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_����, "����"): cbrPopupItem.Visible = True
    Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_չ��, "չ��"): cbrPopupItem.Visible = True: cbrPopupItem.BeginGroup = True
    Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_�۵�, "�۵�"): cbrPopupItem.Visible = True
    Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_��С��, "ִ������ʱ�����۵�"): cbrPopupItem.BeginGroup = True
    Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_View_Refresh, "ˢ��ҽ��"): cbrPopupItem.BeginGroup = True

    cbrPopupBar.ShowPopup
End Sub

Private Sub picForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '�ƶ������λ��
    If Button <> 1 Then Exit Sub
    If Val(picForm.Tag) <> 1 Then Exit Sub
    
    Call MoveWindow(Me.hwnd, (Me.Left + X - mlngLeft_Mouse) / Screen.TwipsPerPixelX, (Me.Top + Y - mlngTop_Mouse) / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, True)
End Sub

Private Sub picForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mlngLeft_Mouse = 0
    mlngTop_Mouse = 0
    picForm.Tag = 0
End Sub

Private Sub SetMode()
    If mblnNormal Then
        Me.Width = lngWidth_Normal
        Me.Height = lngHeight_Normal
        picInfo.Visible = False
        
        picNotify.Visible = True
        picForm.Height = lngHeight_Normal
        picForm.Width = lngWidth_Normal
        picForm.ZOrder 0
        Call SetWindowPos(Me.hwnd, HWND_TOPMOST, mlngLeft_Form / Screen.TwipsPerPixelX, mlngTop_Form / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, SWP_SHOWWINDOW)
    Else
        Me.picNotify.Visible = False
        If mblnOrientation Then
            '����
            Me.Width = lngWidthH_Collapse
            Me.Height = lngHeightH_Collapse
            
            picInfo.Visible = True
            picInfo.Width = Me.Width - 60
            picInfo.ZOrder 0
            picInfo.Refresh
            Call SetWindowPos(Me.hwnd, HWND_TOPMOST, mlngLeft_Form / Screen.TwipsPerPixelX, mlngTop_Form / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, SWP_SHOWWINDOW)
        Else
            '����
            Me.Width = lngWidthV_Collapse
            Me.Height = lngHeightV_Collapse
            
            picInfo.Visible = True
            picInfo.Height = Me.Height - 60
            picInfo.ZOrder 0
            picInfo.Refresh
            Call SetWindowPos(Me.hwnd, HWND_TOPMOST, mlngLeft_Form / Screen.TwipsPerPixelX, mlngTop_Form / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, SWP_SHOWWINDOW)
        End If
        picForm.Height = Me.Height
        picForm.Width = Me.Width
        
    End If
    
    '���ô����͸��Ч��
    Call zlControl.PicShowFlat(picForm, 2)
    Call SetTransparence(Not mblnNormal)
    Me.Refresh
End Sub

Private Sub SetTransparence(Optional ByVal blnTransp As Boolean = True)
    Dim rtn As Long
    
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE) 'ȡ�Ĵ���ԭ�ȵ���ʽ
    rtn = rtn Or WS_EX_LAYERED 'ʹ����������µ���ʽWS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn '���µ���ʽ��������
    
    '38595,������,2012-09-10,�޸�ҽ�����ѿ������ɵ�������(������Ϊ��͸��)
    'ȡ�����汳����ɫΪ&HEAFFFF������Ϊ͸����δ��룬ͳһ�ĳ�չ����͸�����۵�͸��
    'SetLayeredWindowAttributes hwnd, &HEAFFFF, 0, LWA_COLORKEY
    SetLayeredWindowAttributes hwnd, 0, IIf(blnTransp, 180, 255), LWA_ALPHA
End Sub

Private Sub picinfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '���ô�����ʾģʽ
    If Button = 1 Then
        mblnNormal = True
        Call SetMode
    Else
        Call AddMenus
    End If
End Sub

Private Sub picinfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim MouseOver As Boolean
    On Error Resume Next
    
    If mblnNormal Then Exit Sub
    
    '--�жϵ�ǰ���λ���Ƿ��ڲ˵���--
    MouseOver = (0 <= X) And (X <= picInfo.Width) And (0 <= Y) And (Y <= picInfo.Height)
    If MouseOver Then
        Call SetCapture(picInfo.hwnd)
        If mblnFirstIn Then
            Call SetTransparence(False)
            mblnFirstIn = False
        End If
    Else
        Call ReleaseCapture
        Call SetTransparence(True)
        mblnFirstIn = True
    End If
End Sub

Private Sub Time_Flash_Timer()
    '������Ϣ����˸����ֹͣ
    Time_Flash.Enabled = False
    mintCount = mintCount + 1
    
    If mintCount Mod 2 = 0 Then
        lblInfo.ForeColor = 0
        lblTitle.ForeColor = 0
        If Not mblnNormal Then Call SetTransparence(True)
    Else
        lblInfo.ForeColor = 255
        lblTitle.ForeColor = 255
        If Not mblnNormal Then Call SetTransparence(False)
    End If

    Time_Flash.Enabled = True
    If mintCount = 10 Then
        mintCount = 0
        '49547,������,2012-09-05,�������ҽ����Ҫһֱ����
        'Time_Flash.Enabled = False
    End If
End Sub

Private Sub timNotify_Timer()
    Static strPreTime1 As String
    Static strPreTime2 As String
    Dim curTime As Date
    
    curTime = Now
    If gbln����Ӱ����ϢϵͳԤԼ Or mbln���廤����Ϣ Then
        If strPreTime2 = "" Then
            strPreTime2 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
        End If
        If DateDiff("s", CDate(strPreTime2), curTime) > 300 Then
            strPreTime2 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
            If gbln����Ӱ����ϢϵͳԤԼ = True Then
                If mclsPublicAdvice.GetMsgRISReady(mlng����ID) Then
                    Call LoadNotify
                Else
                    If mbln���廤����Ϣ = True Then GoTo NurseMsg
                End If
            Else
NurseMsg:
                Call LoadNurseIntegrateMsg
                Call SetNotifyState
            End If
        End If
    End If
    
    If mbln��Ϣ���� Then
        If Not mrsMsg Is Nothing Then
            If mrsMsg.RecordCount > 0 Then
                TimNotify.Enabled = False
                Call mclsMsg.PlayMsgSound(mrsMsg)
                Set mrsMsg = Nothing
                TimNotify.Enabled = True
            End If
        End If
    End If
    
    'ˢ�²����������
    If mintNotify > 0 Then
        If strPreTime1 = "" Then
            strPreTime1 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
        End If
        If DateDiff("s", CDate(strPreTime1), curTime) > mintNotify * CLng(60) Or mblnFirst Then
            strPreTime1 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
            '������Ϣƽ̨�򲻰��չ̶�ˢ��ʱ��Ϊ׼
            If mclsMipModule.IsConnect = False Or mblnFirst Then
                strPreTime2 = "" '��Ϣˢ�����¼�ʱ
                Call LoadNotify
            End If
            mblnFirst = False
        End If
     Else
        If mblnFirst = True Then
            strPreTime2 = "" '��Ϣˢ�����¼�ʱ
            Call LoadNotify
            mblnFirst = False
        End If
    End If
    
    '�����Ƿ�ͬ��ˢ��ʱʹ��
'    If Right(Format(Now, "mm:ss"), 2) Mod 5 = 0 Then
'        Time_Flash.Enabled = True
'    End If
End Sub

Private Function LoadNotify() As Boolean
    Dim rsTmp As New ADODB.Recordset, rsOut As New ADODB.Recordset
    Dim objOut As Collection, intType As Integer
    Dim strTmp As String, strSQL As String, strTmpRIS As String
    Dim i As Long, blnOk As Boolean
    
    lblTitle.Caption = IIf(mbln���廤����Ϣ = True, "��Ϣ���ѣ�", "ҽ�����ѣ�")
    lblInfo.Caption = lblTitle.Caption
    
    Screen.MousePointer = 11
    On Error GoTo errH
    blnOk = mclsPublicAdvice.GetAdviceRemind(rsTmp, mlng����ID, IIf(optNotify(0).Value = True, UserInfo.����, ""))
    Screen.MousePointer = 0
    If blnOk = False Then Exit Function
    rptNotify.Records.DeleteAll
    If rsTmp Is Nothing Then GoTo GOEND
    If rsTmp.State = adStateClosed Then GoTo GOEND
    
    '90256:��Ժ��ת��ҽ��Ҫ����ʾͼ��(ֻ����¿���ҽ��)
    strTmp = ""
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    For i = 1 To rsTmp.RecordCount
        If rsTmp!���ͱ��� = "ZLHIS_CIS_001" And InStr("," & strTmp & ",", "," & rsTmp!����ID & ":" & rsTmp!��ҳID & ",") = 0 Then
            strTmp = strTmp & "," & rsTmp!����ID & ":" & rsTmp!��ҳID
        End If
        rsTmp.MoveNext
    Next
    If Left(strTmp, 1) = "," Then strTmp = Mid(strTmp, 2)
    Set objOut = New Collection
    If strTmp <> "" Then
        '92088:�����˺�Ӥ��ͬʱ���ڳ�Ժҽ����ʱ����ͬ����SQL�����˳���������,�ʼ�Distinct(objOut.Add ��Ҳ�����ж�)
        strSQL = "Select /*+ RULE*/ " & vbNewLine & _
            " Distinct a.����id, a.��ҳid, First_Value(b.��������) Over(Partition By a.����id, a.��ҳid Order By a.����ʱ�� Desc) As ҽ������" & vbNewLine & _
            " From ����ҽ����¼ a, ������ĿĿ¼ b, Table(Cast(f_Str2list2([1]) As Zltools.t_Strlist2)) c" & vbNewLine & _
            " Where a.������Ŀid + 0 = b.Id And a.������� = 'Z' And Instr(',3,5,11,', ',' || b.�������� || ',') > 0 And a.ҽ��״̬ = 1 And" & vbNewLine & _
            "      a.����id = c.C1 And a.��ҳid = c.C2"
        Set rsOut = zlDatabase.OpenSQLRecord(strSQL, "��ȡҽ����Ϣ", strTmp)
        For i = 1 To rsOut.RecordCount
            '�����ڲ����
            If GetOutType(objOut, rsOut!����ID & "_" & rsOut!��ҳID) = 0 Then
                objOut.Add Decode(Val(NVL(rsOut!ҽ������, 0)), 3, 4, 5, 3, 11, 3, 0), rsOut!����ID & "_" & rsOut!��ҳID
            End If
            rsOut.MoveNext
        Next i
    End If
    strTmp = ","
    strTmpRIS = ","
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    For i = 1 To rsTmp.RecordCount
        Select Case rsTmp!���ͱ��� & ""
        Case "ZLHIS_PACS_006", "ZLHIS_PACS_007"
            'ZLHIS_PACS_006 ZLHIS_PACS_007 ��ϢΪһ����Ϣ����ҽ����ĿΪ��λ����ʾһ��
            If InStr(strTmpRIS, "," & rsTmp!���ͱ��� & "," & rsTmp!ҵ���ʶ & ",") = 0 Then
                strTmpRIS = strTmpRIS & rsTmp!���ͱ��� & "," & rsTmp!ҵ���ʶ & ","
                intType = 0
                Call AddReportRow(intType, rsTmp!����ID & "," & rsTmp!��ҳID, rsTmp!����ID, rsTmp!��ҳID, NVL(rsTmp!����), NVL(rsTmp!סԺ��), NVL(rsTmp!����), NVL(rsTmp!��Ϣ����), _
                    rsTmp!���ͱ��� & "", rsTmp!���ȳ̶� & "", Format(rsTmp!�Ǽ�ʱ�� & "", "yyyy-MM-dd HH:mm:ss"), rsTmp!ҵ���ʶ & "", rsTmp!������Դ & "", NVL(rsTmp!����, 0), NVL(rsTmp!���ﲡ��id, 0), rsTmp!���ͱ��� & "," & rsTmp!ҵ���ʶ)
            End If
        Case Else
            If InStr(strTmp, "," & rsTmp!����ID & "," & rsTmp!��ҳID & "," & rsTmp!���ͱ��� & ",") = 0 Then
                strTmp = strTmp & rsTmp!����ID & "," & rsTmp!��ҳID & "," & rsTmp!���ͱ��� & ","
                intType = 0
                If rsTmp!���ͱ��� = "ZLHIS_CIS_001" Then intType = GetOutType(objOut, rsTmp!����ID & "_" & rsTmp!��ҳID)
                Call AddReportRow(intType, rsTmp!����ID & "," & rsTmp!��ҳID, rsTmp!����ID, rsTmp!��ҳID, NVL(rsTmp!����), NVL(rsTmp!סԺ��), NVL(rsTmp!����), NVL(rsTmp!��Ϣ����), _
                    rsTmp!���ͱ��� & "", rsTmp!���ȳ̶� & "", Format(rsTmp!�Ǽ�ʱ�� & "", "yyyy-MM-dd HH:mm:ss"), rsTmp!ҵ���ʶ & "", rsTmp!������Դ & "", NVL(rsTmp!����, 0), NVL(rsTmp!���ﲡ��id, 0), rsTmp!����ID & "," & rsTmp!��ҳID & "," & rsTmp!���ͱ���)
            End If
        End Select
        rsTmp.MoveNext
    Next
    
GOEND:
    Call LoadNurseIntegrateMsg 'ˢ���ƶ�������Ϣ
    
    Call SetNotifyState
    
    LoadNotify = True
    mbln��Ϣ���� = Val(zlDatabase.GetPara("����������ʾ", glngSys, pסԺ��ʿվ)) = 1
    If mbln��Ϣ���� Then
        If mclsMsg Is Nothing Then
            Set mclsMsg = New clsCISMsg
            Call mclsMsg.InitCISMsg(2)
        End If
        If Not rsTmp Is Nothing Then
            If Not rsTmp.State = adStateClosed Then
                If rsTmp.RecordCount > 0 Then
                    rsTmp.MoveFirst
                    Set mrsMsg = rsTmp
                End If
            End If
        End If
    End If
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetOutType(ByVal objOut As Collection, ByVal strKey As String) As Integer
    Dim intType As Integer
    On Error Resume Next
    intType = Val(objOut(strKey))
    If err <> 0 Then err.Clear
    GetOutType = intType
End Function

Private Sub AddMsgToLis(ByVal rsMsg As ADODB.Recordset)
'���ܣ������յ�����Ϣ���������б���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim i As Long
    
    On Error GoTo errH
    
    If Mid(rsMsg!���ѳ���, 3, 1) <> "1" Then Exit Sub
    
    If InStr("," & rsMsg!����IDs & ",", "," & mlng����ID & ",") > 0 Or _
        InStr("," & rsMsg!������Ա & ",", "," & UserInfo.���� & ",") > 0 Then
        
        '�ж��б��Ƿ��Ѿ���������Ϣ�ˣ����� AddReportRow ���жϣ��������ܻ����һ��SQL��ѯ
        For i = 0 To rptNotify.Rows.Count - 1
            If Not rptNotify.Rows(i).GroupRow Then
                If rptNotify.Rows(i).Record(C_��Ϣ).Value = rsMsg!���ͱ��� And rptNotify.Rows(i).Record.Tag = CStr(rsMsg!����ID & "," & rsMsg!����id) Then
                    Exit Sub
                End If
            End If
        Next
        
        strSQL = "Select a.סԺ��, a.����, a.�Ա�, a.����, a.��ǰ���� As ����, a.����,B.��ǰ����ID ���ﲡ��id From ������Ϣ A,������ҳ B Where A.����ID=B.����ID  And  B.����id =[1] and B.��ҳID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsMsg!����ID), Val(rsMsg!����id))
        
        If Not rsTmp.EOF Then
            Call AddReportRow(rsMsg!����ID & "," & rsMsg!����id, rsMsg!����ID, rsMsg!����id, NVL(rsTmp!����), NVL(rsTmp!סԺ��), NVL(rsTmp!����), NVL(rsMsg!��Ϣ����), _
                 rsMsg!���ͱ��� & "", rsMsg!���ȳ̶� & "", Format(rsMsg!�Ǽ�ʱ�� & "", "yyyy-MM-dd HH:mm:ss"), rsMsg!ҵ���ʶ & "", rsMsg!������Դ & "", NVL(rsTmp!����, 0), NVL(rsTmp!���ﲡ��id, 0), rsMsg!����ID & "," & rsMsg!����id & "," & rsMsg!���ͱ���)
            Call SetNotifyState
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub AddReportRow(ByVal intType As Integer, ParamArray arrInput() As Variant)
'���ܣ�����Ϣ�����б�������һ��
'intType:�¿�ҽ��������(3:��Ժ,4-ת��)

    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim strRowID As String '�����б��е�Ψһ��ʶ��"����id,��ҳid,��Ϣ����"
    Dim strNO As String
    Dim strҵ�� As String
    Dim str������Դ As String
    Dim int���ȼ� As Integer
    Dim int���� As Integer
    Dim Index As Integer
    Dim objItemIcon As ReportRecordItem
    
    On Error GoTo errH
    
    Set objRecord = Me.rptNotify.Records.Add()
    objRecord.Tag = arrInput(Index): Index = Index + 1         'Tagֵ
    Set objItem = objRecord.AddItem(""): objItem.Icon = 1
    Set objItemIcon = objItem
    
    objRecord.AddItem Val(arrInput(Index)): Index = Index + 1  '����id
    objRecord.AddItem Val(arrInput(Index)): Index = Index + 1  '����id
    Set objItem = objRecord.AddItem(CStr(arrInput(Index))): Index = Index + 1 '����
    If intType = 3 Or intType = 4 Then objItem.Icon = intType 'ͼ�����:3-��Ժ,4-ת��
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index))) 'סԺ��
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    Set objItem = objRecord.AddItem(zlStr.Lpad(CStr(arrInput(Index)), 10, " ")) '����
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    Set objItem = objRecord.AddItem(CStr(arrInput(Index)))     '״̬������
    objItem.Caption = CStr(arrInput(Index)): Index = Index + 1
    
    strNO = arrInput(Index)                            '��Ϣ���
    objRecord.AddItem strNO: Index = Index + 1
    
    int���ȼ� = Val(arrInput(Index))                     '���
    objRecord.AddItem int���ȼ�: Index = Index + 1
    objRecord.AddItem CStr(arrInput(Index)): Index = Index + 1  '����
    
    strҵ�� = arrInput(Index): Index = Index + 1              'ҵ���ʶ
    str������Դ = arrInput(Index): Index = Index + 1          '������Դ
    int���� = arrInput(Index)
    
    If InStr(",ZLHIS_PACS_005,ZLHIS_LIS_003,", "," & strNO & ",") > 0 Then 'Σ��ֵ��Ϣ���⴦���Ķ�ʱ������Ϣ
        objRecord.AddItem strҵ�� & "," & Val(str������Դ)
    Else
        objRecord.AddItem strҵ��
    End If
    
    Index = Index + 1
    objRecord.AddItem Val(arrInput(Index))    '����ID
    Index = Index + 1
    objRecord.AddItem CStr(arrInput(Index))    '��ϢΨһ��ʶ
    
    If int���ȼ� > 1 Then
        For Index = 0 To rptNotify.Columns.Count - 1
            If int���ȼ� = 3 Then
                objRecord.Item(Index).ForeColor = &HC0&
            End If
            objRecord.Item(Index).Bold = True
        Next
        If (strNO = "ZLHIS_CIS_001" Or strNO = "ZLHIS_CIS_002") And int���ȼ� = 2 Then objItemIcon.Icon = 2
    End If
    '���ղ����ú�ɫ��ʾ
    If int���� > 0 And int���ȼ� <> 3 Then
        For Index = 0 To rptNotify.Columns.Count - 1
            objRecord.Item(Index).ForeColor = &HC0&
        Next
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetNotifyState()
    rptNotify.Populate 'ȱʡ��ѡ���κ���
    rptNotify.TabStop = rptNotify.Rows.Count > 0
    mstrPreNotify = IIf(rptNotify.Rows.Count = 0, "", rptNotify.Rows.Count)
    If rptNotify.Rows.Count > 0 Then
        lblTitle.Caption = IIf(mbln���廤����Ϣ = True, "��Ϣ���ѣ�", "ҽ�����ѣ�") & "����" & rptNotify.Rows.Count & "����Ϣ��Ҫ����"
        lblInfo.Caption = lblTitle.Caption
        Time_Flash.Enabled = ((mstrPreNotify <> "") Or (mlngPreID <> mlng����ID))
    Else 'û��ҽ����Ϣ��ֹͣ��˸
        lblTitle.Caption = IIf(mbln���廤����Ϣ = True, "��Ϣ���ѣ�", "ҽ�����ѣ�")
        lblInfo.Caption = lblTitle.Caption
        Time_Flash.Enabled = False
        lblInfo.ForeColor = 0
        lblTitle.ForeColor = 0
    End If
    mlngPreID = mlng����ID
End Sub

Private Sub rptNotify_KeyUp(KeyCode As Integer, Shift As Integer)
'���ܣ��Զ�����ҽ��У�ԡ�ȷ��ֹͣ��ִ�н���
    Dim blnExecute As Boolean
    Dim intFunc As Integer
    Dim strҵ�� As String
    Dim strSQL As String
    Dim objControl As CommandBarControl
    Dim strPrivs As String, strPatis As String
    Dim blnOnePati As Boolean
    Dim lng����ID As Long, lng��ҳID As Long
    Dim blnCollateAutoFind As Boolean
    Dim blnTmp As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim strNoteKey As String '��ϢΨһ��ʶ
    Dim blnNurseIntegrate As Boolean '�Ƿ������廤����Ϣ
    
    On Error GoTo ErrHand
    blnCollateAutoFind = (Val(zlDatabase.GetPara("ҽ��������Զ���λ��ҽ��ҳ��", glngSys, 1265, 0)) = 1)
    strNoteKey = ""
    intFunc = -1
    If KeyCode = vbKeyReturn Then
        If rptNotify.SelectedRows.Count > 0 Then
            With rptNotify.SelectedRows(0).Record
                strҵ�� = .Item(C_ҵ��).Value
                lng����ID = Val(.Item(C_����ID).Value)
                lng��ҳID = Val(.Item(C_��ҳID).Value)
                strNoteKey = .Item(C_Ψһ��ʶ).Value
                If InStr(",ZLHIS_CIS_001,ZLHIS_CIS_002,", .Item(C_��Ϣ).Value) > 0 Then
                    strPrivs = GetInsidePrivs(pסԺҽ������)
                    If .Item(C_��Ϣ).Value = "ZLHIS_CIS_001" Then
                        If Val(zlDatabase.GetPara("����ǰ�Զ�У��", glngSys, pסԺҽ������, 0)) = 1 Then
                            intFunc = 0
                        Else
                            intFunc = 1
                        End If
                    ElseIf .Item(C_��Ϣ).Value = "ZLHIS_CIS_002" Then
                        intFunc = 2
                    End If
                Else
                    strTmp = ""
                    '55430:������,2013-02-27,˫������ҽ����λ�����������ҽ��ҳ��,��ʿվ���ܴ���Σ��ֵ��Ϣ
                    Select Case .Item(C_��Ϣ).Value
                        Case "ZLHIS_BLOOD_003", "ZLHIS_BLOOD_001", "ZLHIS_BLOOD_007" 'ȡѪ����,��Ѫ�������,Ѫ����������
                            intFunc = 3
                        Case "ZLHIS_CIS_003" '����ҽ��
                            intFunc = 3
                        Case "ZLHIS_OPER_001,ZLHIS_CIS_005,ZLHIS_CIS_015" '��������
                            intFunc = -1
                        Case "ZLHIS_TRANSFUSION_001" '��Һ���δͨ��
                            intFunc = 11
                        Case "ZLHIS_CHARGE_001" '������������
                            intFunc = 12
                        Case "ZLHIS_LIS_003" '����Σ��ֵ
                            'strTmp = "ZLHIS_CIS_014"
                            Exit Sub
                        Case "ZLHIS_PACS_005" '���Σ��ֵ
                            'strTmp = "ZLHIS_CIS_025"
                            Exit Sub
                        Case "ZLHIS_NURSE_INTEGRATE" '���廤����Ϣ
                            blnNurseIntegrate = True
                    End Select
                    If strTmp <> "" And blnNurseIntegrate = False Then
                        If Not (mclsMipModule Is Nothing) Then
                            If mclsMipModule.IsConnect Then
                                strSQL = "select ��Ժ����ID,��ǰ����ID from ������ҳ where ����ID=[1] and ��ҳID=[2]"
                                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.Item(C_����ID).Value), Val(.Item(C_��ҳID).Value))
                                Call ZLHIS_CIS_MsgReadAfter(mclsMipModule, strTmp, .Item(C_����ID).Value, .Item(c_����).Value, .Item(c_סԺ��).Value, , Val(Split(strҵ��, ",")(1)), _
                                        .Item(C_��ҳID).Value, Val(rsTmp!��ǰ����ID & ""), Val(rsTmp!��Ժ����ID & ""), .Item(c_����).Value, Val(Split(strҵ��, ",")(0)))
                            End If
                        End If
                    End If
                End If
                strSQL = ""
                If blnNurseIntegrate = False Then
                    '������Ϣ�Ķ�״̬(ҵ����Ϣ��ر�����Ϊ�������֣����õ�����Ȩ)
                    If .Item(C_��Ϣ).Value = "ZLHIS_PACS_006" Or .Item(C_��Ϣ).Value = "ZLHIS_PACS_007" Then
                        strSQL = "Zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng��ҳID & ",'" & .Item(C_��Ϣ).Value & "',3,'" & UserInfo.���� & "'," & mlng����ID & ",null,null,'" & .Item(C_ҵ��).Value & "')"
                    ElseIf .Item(C_��Ϣ).Value = "ZLHIS_BLOOD_007" And gblnѪ��ϵͳ Then     'δ����ǰ��������Ϊ�Ѷ�
                        If gobjPublicBlood Is Nothing And gblnѪ��ϵͳ Then InitObjBlood
                        If gobjPublicBlood.zlIsBloodMessageDone(1, lng����ID, lng��ҳID, 3, mlng����ID) Then
                            If strNoteKey <> "" Then
                                Call ReMoveItemByKey(strNoteKey)
                                Call SetNotifyState
                            End If
                            If intFunc > -1 Then
                                Call frmSublimeInNurseStation.ExecFuncs(intFunc)
                            End If
                        End If
                        Exit Sub
                    Else
                        strSQL = "Zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng��ҳID & ",'" & .Item(C_��Ϣ).Value & "',3,'" & UserInfo.���� & "'," & mlng����ID & ")"
                    End If
                Else
                    Call ReadNurseIntegrateMsg(strNoteKey)
                    Exit Sub
                End If
            End With
        End If
    End If
    
    If intFunc > -1 Then
        If mblnExecCollapse Then Call imgHide_MouseDown(1, 0, 0, 0)
    End If
    
    Select Case intFunc
    Case 0, 1
        If Not HaveOperateAdvice(lng����ID, lng��ҳID, 0) Then
            If strSQL <> "" Then Call zlDatabase.ExecuteProcedure(strSQL, "Zl_ҵ����Ϣ�嵥_Read")
            Call ReMoveItemByKey(strNoteKey)
            Call SetNotifyState
        Else
            If intFunc = 0 Then
                If InStr(strPrivs, ";����ҩ������;") > 0 Or InStr(strPrivs, ";����ҩ�Ƴ���;") > 0 Or InStr(strPrivs, ";������������;") > 0 Or InStr(strPrivs, ";������������;") > 0 Then
                    Call mclsPublicAdvice.AdviceSend(Me, mlng����ID, lng����ID, lng��ҳID, gstrPrivs, mclsMipModule)
                    
                    If Not HaveOperateAdvice(lng����ID, lng��ҳID, 0) Then
                        If strSQL <> "" Then Call zlDatabase.ExecuteProcedure(strSQL, "Zl_ҵ����Ϣ�嵥_Read")
                        Call ReMoveItemByKey(strNoteKey)
                        Call SetNotifyState
                    End If
                    If blnCollateAutoFind Then Call frmSublimeInNurseStation.ExecFuncs(3)
                End If
            ElseIf intFunc = 1 Then
                If InStr(strPrivs, ";ҽ��У�Դ���;") > 0 Then
                    blnOnePati = Val(zlDatabase.GetPara("����ҽ��У��", glngSys, pסԺҽ������)) = 0
                    blnTmp = mclsPublicAdvice.AdviceOperate(Me, gstrPrivs, 3, lng����ID, lng��ҳID, mlng����ID, Val(strҵ��), mclsMipModule, strPatis, blnOnePati)
                    If strPatis <> "" And blnTmp Then Call BatchRemove(strPatis)
                    If Not blnTmp Then
                        If Not HaveOperateAdvice(lng����ID, lng��ҳID, 0) Then
                            If strSQL <> "" Then Call zlDatabase.ExecuteProcedure(strSQL, "Zl_ҵ����Ϣ�嵥_Read")
                            Call ReMoveItemByKey(strNoteKey)
                            Call SetNotifyState
                        End If
                    End If
                    If blnCollateAutoFind Then Call frmSublimeInNurseStation.ExecFuncs(3)
                End If
            End If
        End If
    Case 2
        If InStr(strPrivs, ";ҽ��ȷ��ֹͣ;") > 0 Then
            If Not HaveOperateAdvice(lng����ID, lng��ҳID, 1) Then
                If strSQL <> "" Then Call zlDatabase.ExecuteProcedure(strSQL, "Zl_ҵ����Ϣ�嵥_Read")
                Call ReMoveItemByKey(strNoteKey)
                Call SetNotifyState
            Else
                blnTmp = mclsPublicAdvice.AdviceOperate(Me, gstrPrivs, 2, lng����ID, lng��ҳID, mlng����ID, Val(strҵ��), mclsMipModule, strPatis, True)
                If strPatis <> "" And blnTmp Then
                    Call ReMoveItemByKey(strNoteKey)
                    Call SetNotifyState
                End If
                If Not blnTmp Then
                    If Not HaveOperateAdvice(lng����ID, lng��ҳID, 1) Then
                        If strSQL <> "" Then Call zlDatabase.ExecuteProcedure(strSQL, "Zl_ҵ����Ϣ�嵥_Read")
                        Call ReMoveItemByKey(strNoteKey)
                        Call SetNotifyState
                    End If
                End If
                If blnCollateAutoFind Then Call frmSublimeInNurseStation.ExecFuncs(3)
            End If
        End If
    Case Else
        If strSQL <> "" Then Call zlDatabase.ExecuteProcedure(strSQL, "Zl_ҵ����Ϣ�嵥_Read")
        If strNoteKey <> "" Then
            Call ReMoveItemByKey(strNoteKey)
            Call SetNotifyState
        End If
        If intFunc > -1 Then
            Call frmSublimeInNurseStation.ExecFuncs(intFunc)
        End If
    End Select
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub rptNotify_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call rptNotify_KeyUp(vbKeyReturn, 0)
End Sub

Private Sub rptNotify_SelectionChanged()
    Dim strBed As String, strKey As String, strNoteKey As String
    Dim strNO As String
    Dim lng���ﲡ��ID As Long
    Dim lng����ID As Long
    Dim lng��ҳID As Long
    
    If rptNotify.SelectedRows.Count = 0 Then Exit Sub  '���������
    
    strNO = CStr(Trim(rptNotify.SelectedRows(0).Record.Item(C_��Ϣ).Value))
    lng����ID = Val(Trim(rptNotify.SelectedRows(0).Record.Item(C_����ID).Value))
    lng��ҳID = Val(Trim(rptNotify.SelectedRows(0).Record.Item(C_��ҳID).Value))
    lng���ﲡ��ID = Val(Trim(rptNotify.SelectedRows(0).Record.Item(C_���ﲡ��).Value))
    
    strBed = Trim(rptNotify.SelectedRows(0).Record.Item(c_����).Value)
    strKey = Trim(rptNotify.SelectedRows(0).Record.Item(C_����ID).Value) & "|" & Trim(rptNotify.SelectedRows(0).Record.Item(C_��ҳID).Value)
    strNoteKey = Trim(rptNotify.SelectedRows(0).Record.Item(C_Ψһ��ʶ).Value)
    
    If ReadAndSendMsg(strNO, lng����ID, lng��ҳID, lng���ﲡ��ID) Then
        Call ReMoveItemByKey(strNoteKey)
        If rptNotify.Records.Count = 0 Then
            Call imgHide_MouseDown(1, 0, 0, 0)
        End If
        Call SetNotifyState
        Exit Sub
    End If
    
    Call frmSublimeInNurseStation.SelPatiCard(strBed, strKey)
    
    rptNotify.SetFocus
End Sub


Private Function ReadAndSendMsg(ByVal strNO As String, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng���ﲡ��ID As Long) As Boolean
    '���ܣ��¿���Ϣʱ������Ϣ�Ĳ����Ѿ����ٵ�ǰ���������Ƚ���Ϣ��Ϊ�Ѷ��������·�����Ϣ
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim arrSQL() As String
    Dim lng��ǰ����ID As Long
    Dim lng��ǰ����ID As Long
    Dim blnTrans As Boolean
    
    On Error GoTo errH
    
    strSQL = "select nvl(A.��ǰ����ID,0) as ��ǰ����ID, nvl(A.��ǰ����ID,0) as ��ǰ����ID from ������Ϣ A where A.����ID = [1] and ��ҳID = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)

    If rsTmp.EOF Then Exit Function
    
    lng��ǰ����ID = Val(rsTmp!��ǰ����id)
    lng��ǰ����ID = Val(rsTmp!��ǰ����ID)
    
    If lng���ﲡ��ID <> lng��ǰ����ID And lng��ǰ����ID <> 0 Then
        If strNO <> "ZLHIS_CIS_001" Then Exit Function
        If Not HaveOperateAdvice(lng����ID, lng��ҳID, 0) Then
            '������ϢΪ�Ѷ�
            strSQL = "Zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng��ҳID & ",'" & strNO & "','0010','" & _
            UserInfo.���� & "'," & lng���ﲡ��ID & ")"
            gcnOracle.BeginTrans: blnTrans = True
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            gcnOracle.CommitTrans: blnTrans = False
        Else
            strSQL = "select A.��Ϣ����, A.���ѳ���,A.���ͱ���,A.ҵ���ʶ,A.���ȳ̶� From ҵ����Ϣ�嵥 A Where a.����id=[1] And a.����id=[2] And a.���ͱ��� =[3] and a.���ﲡ��ID =[4]  And a.�Ƿ�����=0 And Rownum<2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID, strNO, lng���ﲡ��ID)
            If rsTmp.RecordCount > 0 Then
                For i = 0 To rsTmp.RecordCount - 1
                    ReDim Preserve arrSQL(i)
                    arrSQL(UBound(arrSQL)) = "Zl_ҵ����Ϣ�嵥_Insert(" & lng����ID & "," & lng��ҳID & "," & lng��ǰ����ID & "," & lng��ǰ����ID & ",2,'" & rsTmp!��Ϣ���� & "','" & rsTmp!���ѳ��� & "','" & rsTmp!���ͱ��� & "','" & rsTmp!ҵ���ʶ & "'," & rsTmp!���ȳ̶� & ",0,null," & lng��ǰ����ID & ",null)"
                Next
            End If
            
            '������ϢΪ�Ѷ�
            strSQL = "Zl_ҵ����Ϣ�嵥_Read(" & lng����ID & "," & lng��ҳID & ",'" & strNO & "','0010','" & _
            UserInfo.���� & "'," & lng���ﲡ��ID & ")"
            
            gcnOracle.BeginTrans: blnTrans = True
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            '���·�����Ϣ
            If UBound(arrSQL) <> -1 Then
                For i = 0 To UBound(arrSQL)
                    zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
                Next
            End If
            gcnOracle.CommitTrans: blnTrans = False
        End If
        ReadAndSendMsg = True
    End If
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub BatchRemove(ByVal strPatis As String, Optional ByVal blnNurseIntegrateMsg As Boolean = False)
    '���ܣ������Ƴ�ҽ����ʾ��Ϣ(�Բ���Ϊ��λ)�������廤����Ϣ����blnNurseIntegrateMsg=True ʱ����ҽ�����廤����Ϣ
    Dim objRow As ReportRow
    Dim strTmp As String
    Dim strIndexs As String
    Dim arrTmp As Variant
    Dim i As Integer
    
    For Each objRow In rptNotify.Rows
        If objRow.GroupRow Then objRow.Expanded = True
        If Not objRow.GroupRow And objRow.Childs.Count = 0 Then
            If blnNurseIntegrateMsg = False Then
                If InStr(";" & strPatis & ";", ";" & objRow.Record.Tag & ";") > 0 And objRow.Record(C_��Ϣ).Value = "ZLHIS_CIS_001" Then
                    strIndexs = strIndexs & "," & objRow.Record.Index
                End If
            Else
                If objRow.Record(C_��Ϣ).Value = "ZLHIS_NURSE_INTEGRATE" Then
                    strIndexs = strIndexs & "," & objRow.Record.Index
                End If
            End If
        End If
    Next
    If strIndexs <> "" Then
        strIndexs = Mid(strIndexs, 2)
        arrTmp = Split(strIndexs, ",")
        For i = UBound(arrTmp) To 0 Step -1
            Call rptNotify.Records.RemoveAt(Val(arrTmp(i)))
        Next
        Call SetNotifyState
    End If
End Sub

Private Sub ReMoveItemByKey(ByVal strNoteKey As String)
'���ܣ�������Ϣ�б�Ψһ��ʶ�����ݣ��Ƴ���Ӧ��Ϣ
    Dim objRow As ReportRow
    If strNoteKey = "" Then Exit Sub
    For Each objRow In rptNotify.Rows
        If objRow.GroupRow Then objRow.Expanded = True
        If Not objRow.GroupRow And objRow.Childs.Count = 0 Then
            'Ψһ��ʶ�п϶�ֻ��һ��
            If objRow.Record(C_Ψһ��ʶ).Value = strNoteKey Then
                Call rptNotify.Records.RemoveAt(objRow.Record.Index)
                Exit Sub
            End If
        End If
    Next
End Sub

Private Sub LoadNurseIntegrateMsg()
'���ܣ���ȡ���廤����Ϣ�б�
    Dim strMsg As String, strErrMsg As String
    Dim objXML As New DOMDocument
    Dim objNodeList As IXMLDOMNodeList
    Dim i As Integer
    
    '��Ϣ�ڵ���������
    Dim strID As String, strPatientID As String
    Dim lng����ID As Long, lng��ҳID As Long
    Dim strName As String, strPatiNo As String, strBedNo As String, strContent As String, lng���ﲡ��ID As Long, int���� As Integer
    Dim strCreateTime As String, strToUser As String, strRetrun As String
    Dim objPati As Collection
    Dim blnAdd As Boolean
    
    If mbln���廤����Ϣ = True Then
        If InitNurseIntegrate = True Then
            If gobjNurseIntegrate.GetMsg(mstrRelatedUnitID, strMsg, strErrMsg) = True Then
                '���ǰ��ɾ�����廤�����Ϣ��Ϣ
                Call BatchRemove("", True)
                If objXML.loadXML(strMsg) = False Then Exit Sub
                Set objNodeList = objXML.selectNodes(".//List//Msg")
                'XML���ظ�ʽ
                Set objPati = New Collection
                For i = 0 To objNodeList.length - 1
                    strID = objNodeList.Item(i).childNodes(0).Text
                    strName = objNodeList.Item(i).childNodes(1).Text
                    strBedNo = objNodeList.Item(i).childNodes(4).Text
                    strContent = objNodeList.Item(i).childNodes(5).Text
                    strCreateTime = objNodeList.Item(i).childNodes(7).Text
                    strPatientID = objNodeList.Item(i).childNodes(9).Text
                    lng����ID = Val(objNodeList.Item(i).childNodes(10).Text)
                    lng��ҳID = Val(objNodeList.Item(i).childNodes(11).Text)
                    strToUser = objNodeList.Item(i).childNodes(13).Text
                    
                    blnAdd = False
                    If optNotify(0).Value = True And UCase(strToUser) = UCase(UserInfo.�û���) Then
                        blnAdd = True
                    Else
                        blnAdd = True
                    End If
                    If blnAdd = True Then
                        '���ݲ���ID��ȡ�ƶ��������
                        strRetrun = GetPatiData(lng����ID, lng��ҳID, objPati)
                        If strRetrun <> "" Then
                            strPatiNo = Split(strRetrun, "'")(3)
                            lng���ﲡ��ID = Val(Split(strRetrun, "'")(4))
                            int���� = Val(Split(strRetrun, "'")(6))
                            Call AddReportRow(0, lng����ID & "," & lng��ҳID, lng����ID, lng��ҳID, strName, strPatiNo, strBedNo, strContent, _
                                "ZLHIS_NURSE_INTEGRATE" & "", 1 & "", Format(strCreateTime & "", "yyyy-MM-dd HH:mm:ss"), strID & "", 2 & "", int����, lng���ﲡ��ID, strID)
                        Else
                            Call AddReportRow(0, lng����ID & "," & lng��ҳID, lng����ID, lng��ҳID, strName, strPatiNo, strBedNo, strContent, _
                                "ZLHIS_NURSE_INTEGRATE" & "", 1 & "", Format(strCreateTime & "", "yyyy-MM-dd HH:mm:ss"), strID & "", 2 & "", 0, mlng����ID, strID)
                        End If
                    End If
                Next i
            Else
                MsgBox "��ȡ���廤����ID�ӿڵ���ʧ�ܣ�" & vbCrLf & "��ϸ��Ϣ��" & strErrMsg, vbInformation, gstrSysName
            End If
        End If
    Else
        '������廤����Ϣ
        Call BatchRemove("", True)
    End If
End Sub

Private Function GetPatiData(ByVal lng����ID As Long, ByVal lng��ҳID As Long, objPati As Collection) As String
    Dim rsTemp As New ADODB.Recordset
    Dim strRetrun As String, strKey As String
    Dim strSQL As String
    
    On Error Resume Next
    strKey = lng����ID & "_" & lng��ҳID
    strRetrun = objPati(strKey)
    If err <> 0 Then
        err.Clear
        On Error GoTo ErrHand
        strSQL = "Select ����,�Ա�,����,סԺ��,��ǰ����ID,��Ժ����,���� From ������ҳ where ����ID=[1] And ��ҳID=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", lng����ID, lng��ҳID)
        If Not rsTemp.EOF Then
            strRetrun = rsTemp!���� & "'" & rsTemp!�Ա� & "'" & rsTemp!���� & "'" & rsTemp!סԺ�� & "'" & NVL(rsTemp!��ǰ����ID, 0) & "'" & rsTemp!��Ժ���� & "'" & NVL(rsTemp!����, 0)
            objPati.Add strRetrun, strKey
        End If
    End If
    GetPatiData = strRetrun
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub ReadNurseIntegrateMsg(ByVal strID As String)
    Dim strErrMsg As String
    If InitNurseIntegrate = True Then
        If gobjNurseIntegrate.ReplyMsg(strID, strErrMsg) = True Then
            Call ReMoveItemByKey(strID)
            Call SetNotifyState
        Else
            MsgBox strErrMsg
        End If
    End If
End Sub
