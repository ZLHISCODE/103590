VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmBatchAction 
   ClientHeight    =   7860
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11100
   Icon            =   "frmBatchAction.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7860
   ScaleWidth      =   11100
   StartUpPosition =   1  '����������
   Visible         =   0   'False
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3945
      Left            =   6540
      ScaleHeight     =   3945
      ScaleWidth      =   3825
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   900
      Width           =   3825
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   2265
         Left            =   60
         TabIndex        =   33
         Top             =   300
         Width           =   3525
         _Version        =   589884
         _ExtentX        =   6218
         _ExtentY        =   3995
         _StockProps     =   0
         BorderStyle     =   2
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         ShowItemsInGroups=   -1  'True
         AutoColumnSizing=   0   'False
      End
      Begin VB.CheckBox chkfilter 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDD6C6&
         Caption         =   "��ͨ"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   840
         TabIndex        =   32
         ToolTipText     =   "סԺ�걾"
         Top             =   30
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkfilter 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDD6C6&
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   31
         ToolTipText     =   "�����ֱ�ӵǼǱ걾"
         Top             =   30
         Value           =   1  'Checked
         Width           =   675
      End
   End
   Begin RichTextLib.RichTextBox RtfTxt 
      Height          =   585
      Left            =   6270
      TabIndex        =   25
      Top             =   5490
      Visible         =   0   'False
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   1032
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmBatchAction.frx":6852
   End
   Begin VB.PictureBox picWhere 
      BorderStyle     =   0  'None
      Height          =   7065
      Left            =   210
      ScaleHeight     =   7065
      ScaleWidth      =   5535
      TabIndex        =   0
      Top             =   330
      Width           =   5535
      Begin XtremeReportControl.ReportControl rptMachine 
         Height          =   3105
         Left            =   60
         TabIndex        =   17
         Top             =   3000
         Width           =   4965
         _Version        =   589884
         _ExtentX        =   8758
         _ExtentY        =   5477
         _StockProps     =   0
         BorderStyle     =   2
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         ShowItemsInGroups=   -1  'True
         AutoColumnSizing=   0   'False
      End
      Begin VB.OptionButton optSort 
         Caption         =   "����"
         Height          =   180
         Index           =   2
         Left            =   4620
         TabIndex        =   39
         Top             =   1440
         Width           =   765
      End
      Begin VB.OptionButton optSort 
         Caption         =   "����"
         Height          =   180
         Index           =   1
         Left            =   3930
         TabIndex        =   38
         Top             =   1440
         Width           =   705
      End
      Begin VB.OptionButton optSort 
         Caption         =   "�걾"
         Height          =   180
         Index           =   0
         Left            =   3270
         TabIndex        =   36
         Top             =   1440
         Width           =   705
      End
      Begin VB.CheckBox chkAbnormal 
         Caption         =   "����ʾ�쳣����걾"
         Height          =   180
         Left            =   60
         TabIndex        =   35
         Top             =   1740
         Width           =   2325
      End
      Begin VB.ComboBox cboDate 
         Height          =   300
         ItemData        =   "frmBatchAction.frx":68EF
         Left            =   90
         List            =   "frmBatchAction.frx":68F1
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   90
         Width           =   1095
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "������ӡ���ı걾"
         Height          =   180
         Left            =   3270
         TabIndex        =   29
         Top             =   1740
         Width           =   1755
      End
      Begin VB.CheckBox chkPatient 
         Caption         =   "ͬһ�����˺ϲ�Ϊһ�����浥��ӡ"
         Height          =   195
         Left            =   60
         TabIndex        =   28
         Top             =   1980
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   3660
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   720
         Width           =   1605
      End
      Begin VB.ComboBox cbo��Դ 
         Height          =   300
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   720
         Width           =   1425
      End
      Begin VB.TextBox TxtModify 
         Height          =   285
         Left            =   1050
         MaxLength       =   15
         TabIndex        =   23
         Top             =   2385
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Frame Frame1 
         Height          =   105
         Left            =   -45
         TabIndex        =   21
         Top             =   2205
         Width           =   5415
      End
      Begin VB.CheckBox chkUnion 
         Caption         =   "����ӡ���ϲ��ı걾"
         Height          =   255
         Left            =   3270
         TabIndex        =   20
         Top             =   1980
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.ComboBox cboExeDept 
         Height          =   300
         Left            =   3660
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1050
         Width           =   1605
      End
      Begin VB.ComboBox cboRequisitionDept 
         Height          =   300
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1050
         Width           =   1425
      End
      Begin VB.ComboBox cboVerifyMan 
         Height          =   300
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1395
         Width           =   1425
      End
      Begin VB.TextBox TxtSample 
         Height          =   285
         Left            =   3060
         TabIndex        =   7
         Top             =   405
         Width           =   2205
      End
      Begin VB.TextBox txtBatchNum 
         Height          =   285
         Left            =   1230
         TabIndex        =   6
         Top             =   405
         Width           =   915
      End
      Begin MSComCtl2.DTPicker DtpBegin 
         Height          =   285
         Left            =   1230
         TabIndex        =   2
         Top             =   90
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   149094403
         CurrentDate     =   39198
      End
      Begin MSComCtl2.DTPicker DtpEnd 
         Height          =   285
         Left            =   3330
         TabIndex        =   4
         Top             =   90
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   149094403
         CurrentDate     =   39198
      End
      Begin VB.Label lblLisSort 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   2850
         TabIndex        =   37
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "���˲���"
         Height          =   180
         Left            =   2850
         TabIndex        =   27
         Top             =   765
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "�� �� �� Դ"
         Height          =   180
         Left            =   150
         TabIndex        =   26
         Top             =   795
         Width           =   990
      End
      Begin VB.Label LabModify 
         AutoSize        =   -1  'True
         Caption         =   "(ȷ���޸ı걾�ŵĿ�ʼ����)"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   1
         Left            =   2475
         TabIndex        =   24
         Top             =   2430
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.Label LabModify 
         AutoSize        =   -1  'True
         Caption         =   "�޸ı걾��"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Top             =   2445
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "�걾���"
         Height          =   180
         Left            =   2265
         TabIndex        =   18
         Top             =   450
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "��������(��ѡ)"
         Height          =   180
         Left            =   90
         TabIndex        =   16
         Top             =   2760
         Width           =   1260
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "ִ�п���"
         Height          =   180
         Left            =   2850
         TabIndex        =   14
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "�� �� �� ��"
         Height          =   180
         Left            =   150
         TabIndex        =   12
         Top             =   1110
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "��  ��   ��"
         Height          =   180
         Left            =   150
         TabIndex        =   10
         Top             =   1455
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "��       ��"
         Height          =   180
         Left            =   150
         TabIndex        =   5
         Top             =   450
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Left            =   3090
         TabIndex        =   3
         Top             =   135
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��"
         Height          =   180
         Left            =   90
         TabIndex        =   1
         Top             =   135
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   7485
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBatchAction.frx":68F3
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14499
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   135
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchAction.frx":7185
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchAction.frx":71F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchAction.frx":778B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchAction.frx":7D25
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchAction.frx":82BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchAction.frx":EB21
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchAction.frx":15383
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchAction.frx":1BBE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchAction.frx":22447
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchAction.frx":28CA9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   5820
      Top             =   990
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmBatchAction.frx":2F50B
      Left            =   5670
      Top             =   1680
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmBatchAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngLeftWidth As Long           '��ߵĿ��
Private Const Dkp_ID_Left As Integer = 101
Private Const Dkp_ID_Right As Integer = 102
Private mintEditType As Integer         '�������� (1=��ӡ 2=��� 3=����ɾ��)
Private mlngMachine As Long             '����ID
Private mstrPrivs As String             'Ȩ��
Private mstrAuditingMan As String       '�����
Private mstrAuditingManID As String     '�����ID
Private mintAuditing As Integer         'ʱ������
Private mDateAuditing As Date           '��˿�ʼʱ��
Private mDeptID As Long                 'ִ�п���ID
Private mintUnion As Integer            '�Ƿ���������������ʾ 0=������ 1=����
Private mMakeNoRule As String           '�걾������ɵ����ڹ���
Private mstrPrintDepts As String        '���Դ�ӡ�Ŀ���
Private mblnExec As Boolean              '�Ƿ�����ִ��

Private Enum mMCol          '����
    ID
    ѡ��
    ����
    ����
    ����
End Enum

Private Enum mCol           '�б�
    ѡ�� = 0
    ����
    ִ��״̬
    �걾��
    �걾����
    ����ʱ��
    ������
    ������
    ����ʱ��
    ������
    �������
    ��������
    ִ�п���
    ҽ��id
    ���ͺ�
    ת��
    �걾id
    ����ID
    �Ƿ����
    ��������
    ���ID
    �걾���
    ����id
    ������Դ
    Ӥ��
    ��������ID
    ������
    ��ҳID
    ������
    ����ʱ��
End Enum

Private mclsUnzip As New cUnzip
Private mclsZip As New cZip


Private Sub cbo��Դ_Click()
    If cbo��Դ.ListIndex = 2 Or cbo��Դ.ListIndex = 0 Then
        cbo����.Enabled = True
    Else
        cbo����.ListIndex = 0
        cbo����.Enabled = False
    End If
End Sub

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.ID
        Case conMenu_File_PrintSet                                                              '��ӡ����
            Call zlPrintSet
        Case conMenu_File_Exit                                                                  '�˳�
            Unload Me
        '---------------------------------------------------------------
        Case conMenu_Manage_ThingModi                                                           'ȫѡ
            Call RptSelect(Me.rptList.Records, True)
            Me.rptList.Populate
        Case conMenu_Manage_ThingDel                                                            'ȫ��
            Call RptSelect(Me.rptList.Records, False)
            Me.rptList.Populate
        Case conMenu_File_Print                                                                 '�����ӡ
            Call SaveData
        Case conMenu_Edit_Audit                                                                 '���
            Call SaveData
        Case conMenu_Edit_Delete                                                                'ɾ��
            Call SaveData
        Case conMenu_Manage_Reset                                                               '�����޸ı걾��
            Call ModifySampleNumber
        
        '---------------------------------------------------------------
        Case conMenu_View_ToolBar                                                               '������
        Case conMenu_View_ToolBar_Button                                                        '��׼��ť
            Me.cbrthis(2).Visible = Not Me.cbrthis(2).Visible
            Me.cbrthis.RecalcLayout
        Case conMenu_View_ToolBar_Text                                                          '�ı���ǩ
            Dim cbrControl As CommandBarControl
            For Each cbrControl In Me.cbrthis(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            Me.cbrthis.RecalcLayout
        Case conMenu_View_ToolBar_Size                                                          '��ͼ��
            Me.cbrthis.Options.LargeIcons = Not Me.cbrthis.Options.LargeIcons
            Me.cbrthis.RecalcLayout
        Case conMenu_View_StatusBar                                                             '״̬��
            Me.stbThis.Visible = Not Me.stbThis.Visible
            Me.cbrthis.RecalcLayout
        Case conMenu_View_Refresh                                                               'ˢ��
            Call RefreshData
        '---------------------------------------------------------------
        Case conMenu_Help_Help                                                                  '��������
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web                                                                   'WEB�ϵ�����
            Call zlHomePage(hWnd)
        Case conMenu_Help_Web_Home                                                              '��ҳ
            Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail                                                              '���ͷ���
            Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About                                                                 '����
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    End Select
End Sub

Private Sub cbrthis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbrthis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
        
    Err = 0: On Error Resume Next
    Select Case Control.ID
        Case conMenu_View_ToolBar_Button:                                                                   '��ť
            Control.Checked = Me.cbrthis(2).Visible
        
        Case conMenu_View_ToolBar_Text:                                                                     '��ť����
            Control.Checked = Not (Me.cbrthis(2).Controls(1).Style = xtpButtonIcon)
        
        Case conMenu_View_ToolBar_Size:                                                                     '��ͼ��
            Control.Checked = Me.cbrthis.Options.LargeIcons
        
        Case conMenu_View_StatusBar:                                                                        '״̬��
            Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub chkfilter_Click(Index As Integer)
    Dim Record  As ReportRecord
    For Each Record In Me.rptList.Records
        If Record.Item(mCol.����).Icon = 2 Then
            Record.Item(mCol.ѡ��).Checked = (chkfilter(0).Value = 1)
        Else
            Record.Item(mCol.ѡ��).Checked = (chkfilter(1).Value = 1)
        End If
    Next
    Me.rptList.Populate
End Sub

Private Sub dkpMain_Resize()
    Me.cbrthis.RecalcLayout
End Sub

Private Sub DtpEnd_Validate(Cancel As Boolean)
    '10765
    If DtpEnd.Value < DtpBegin.Value Then
        MsgBox "�������ڲ���С�ڿ�ʼ���ڣ�", vbInformation, gstrSysName
        Cancel = True
    End If
End Sub

Private Sub DTPBegin_Validate(Cancel As Boolean)
    '10765
    If DtpBegin.Value > DtpEnd.Value Then
        MsgBox "��ʼ���ڲ��ܴ��ڽ������ڣ�", vbInformation, gstrSysName
        Cancel = True
    End If
End Sub

Private Sub Form_Load()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    Dim intSort As Integer
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrthis.VisualTheme = xtpThemeOffice2003
    Set Me.cbrthis.Icons = zlCommFun.GetPubIcons
    With Me.cbrthis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbrthis.EnableCustomization False

    '-----------------------------------------------------
    '�˵�����
    Me.cbrthis.ActiveMenuBar.Title = "�˵�"
'    Me.cbrthis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "����(&A)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ThingModi, "ȫѡ(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ThingDel, "ȫ��(&R)")
        If mintEditType = 1 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "�����ӡ(&P)"): cbrControl.BeginGroup = True
        End If
        If mintEditType = 2 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "���(&A)"): cbrControl.BeginGroup = True
        End If
        If mintEditType = 3 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)"): cbrControl.BeginGroup = True
        End If
        If mintEditType = 4 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Reset, "�޸�(&M)"): cbrControl.BeginGroup = True
        End If
    End With

    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False)
            cbrPopControl.Checked = True
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False)
            cbrPopControl.Checked = True
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False)
            cbrPopControl.Checked = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)"): cbrControl.Checked = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&F)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With

    '�����
    With Me.cbrthis.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_Select
        .Add FCONTROL, Asc("Z"), conMenu_Edit_DeSelect
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F2, conMenu_Edit_Audit
        .Add FCONTROL, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F5, conMenu_View_Refresh
    End With

    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbrthis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ThingModi, "ȫѡ"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ThingDel, "ȫ��")
        If mintEditType = 1 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "�����ӡ"): cbrControl.BeginGroup = True
        End If
        If mintEditType = 2 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "���"): cbrControl.BeginGroup = True
        End If
        If mintEditType = 3 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��"): cbrControl.BeginGroup = True
        End If
        If mintEditType = 4 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Reset, "�޸�"): cbrControl.BeginGroup = True
        End If
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '�ָ�����
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane, Pane4 As Pane
    
    dkpMain.Options.HideClient = True
    mlngLeftWidth = Me.picWhere.Width - 250
    
    Set Pane1 = dkpMain.CreatePane(Dkp_ID_Left, 200, 150, DockLeftOf, Nothing)
    Pane1.Title = "����ѡ��"
    Pane1.Handle = Me.picWhere.hWnd
    Pane1.Options = PaneNoCaption

    Set Pane2 = dkpMain.CreatePane(Dkp_ID_Right, 200, 600, DockRightOf, Nothing)
    Pane2.Title = "�б�"
    Pane2.Handle = Me.picList.hWnd
    Pane2.Options = PaneNoCaption
    
    Pane1.Select
    
    '��ʼ��
    Me.DtpBegin = Now: Me.DtpEnd = Now
    'ʱ��
    With Me.cboDate
        .Clear
        .AddItem "����ʱ��"
        .AddItem "����ʱ��"
        .ListIndex = 0
    End With
    
    
    Dim rsTmp As New ADODB.Recordset
    '���˲���
    With Me.cbo����
        .Clear
        .AddItem "���в���"
        .ItemData(.NewIndex) = 0
    End With
    gstrSql = " select Distinct B.����,B.����,A.����id from �������Ҷ�Ӧ A,���ű� B where A.����id=B.id Order by B.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Do Until rsTmp.EOF
        Me.cbo����.AddItem "" & rsTmp("����")
        Me.cbo����.ItemData(Me.cbo����.NewIndex) = rsTmp("����ID")
        rsTmp.MoveNext
    Loop
    If Me.cbo����.ListCount > 0 Then Me.cbo����.ListIndex = 0
    '������Դ
    With Me.cbo��Դ
        .Clear
        .AddItem "���в���"
        .AddItem "����"
        .AddItem "סԺ"
        .AddItem "����"
        .AddItem "���"
    End With
    If Me.cbo��Դ.ListCount > 0 Then Me.cbo��Դ.ListIndex = 0
    '�������
    With Me.cboRequisitionDept
        .Clear
        .AddItem "���п���"
        .ItemData(.NewIndex) = 0
    End With
    gstrSql = "SELECT A.���� as ����,ID FROM ���ű� A,��������˵�� B " & _
               " WHERE A.ID=B.����id AND B.��������='�ٴ�' ORDER BY A.���� "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Do Until rsTmp.EOF
        With Me.cboRequisitionDept
            .AddItem rsTmp("����")
            .ItemData(.NewIndex) = rsTmp("Id")
        End With
        rsTmp.MoveNext
    Loop
    If Me.cboRequisitionDept.ListCount > 0 Then Me.cboRequisitionDept.ListIndex = 0
    
    'ִ�п���
    With Me.cboExeDept
        .Clear
        .AddItem "���п���"
        .ItemData(.NewIndex) = 0
    End With
    gstrSql = " SELECT A.���� as ����,ID FROM ���ű� A,��������˵�� B " & _
              " WHERE A.ID=B.����id AND b.�������� = '����'  ORDER BY A.����  "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Do Until rsTmp.EOF
        With Me.cboExeDept
            .AddItem rsTmp("����")
            .ItemData(.NewIndex) = rsTmp("Id")
        End With
        rsTmp.MoveNext
    Loop
    If Me.cboExeDept.ListCount > 0 Then Me.cboExeDept.ListIndex = 0
        
    With Me.cboExeDept
        Dim lngIndex As Long
        If mDeptID > 0 And .ListCount > 0 Then
            For lngIndex = 0 To .ListCount - 1
                If mDeptID = .ItemData(lngIndex) Then
                    .ListIndex = lngIndex
                    Exit For
                End If
            Next
        End If
    End With
    '������
    With Me.cboVerifyMan
        .Clear
        .AddItem "������Ա"
        .ItemData(.NewIndex) = 0
    End With
    gstrSql = "Select Distinct ���,���� As ����, a.Id" & vbNewLine & _
            " From ��Ա�� a, ������Ա b, ��������˵�� c" & vbNewLine & _
            " Where a.Id = b.��Աid And b.����id = c.����id And c.�������� = '����'" & vbNewLine & _
            " And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & vbNewLine & _
            " Order By ��� "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Do Until rsTmp.EOF
        With Me.cboVerifyMan
            .AddItem rsTmp("����")
            .ItemData(.NewIndex) = rsTmp("Id")
        End With
        rsTmp.MoveNext
    Loop
    If Me.cboVerifyMan.ListCount > 0 Then Me.cboVerifyMan.ListIndex = 0
    
    '��ʹ���б���
    Dim Column As ReportColumn
    Dim intLoop As Integer
    Dim Record As ReportRecord
    
    With Me.rptMachine.Columns
        
        rptMachine.AllowColumnRemove = False
        rptMachine.ShowItemsInGroups = False
        
        With rptMachine.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
        rptMachine.SetImageList ImgList
        
        Set Column = .Add(mMCol.ID, "ID", 18, False): Column.Visible = False
        Set Column = .Add(mMCol.ѡ��, "ѡ��", 18, False): Column.Icon = 0
        Set Column = .Add(mMCol.����, "����", 65, True)
        Set Column = .Add(mMCol.����, "����", 120, True)
        Set Column = .Add(mMCol.����, "����", 85, True)
        Me.rptMachine.Populate
    End With
    
    gstrSql = "select ID, ����,����,�������� from �������� "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    
    With Me.rptMachine
        .Records.DeleteAll
        .Populate
        Set Record = .Records.Add
        For intLoop = 0 To .Columns.Count
            Record.AddItem ""
        Next
        Record.Item(mMCol.ID).Value = 0
        Record.Item(mMCol.ѡ��).HasCheckbox = True: Record.Item(mMCol.ѡ��).Checked = False
        Record.Item(mMCol.����).Value = "�ֹ�"
    End With
    
    Do Until rsTmp.EOF
        With Me.rptMachine
            Set Record = .Records.Add
            For intLoop = 0 To .Columns.Count
                Record.AddItem ""
            Next
            
            Record.Item(mMCol.ID).Value = Nvl(rsTmp("ID"))
            Record.Item(mMCol.ѡ��).HasCheckbox = True
            If mlngMachine = Nvl(rsTmp("ID")) Then
                Record.Item(mMCol.ѡ��).Checked = True
            Else
                Record.Item(mMCol.ѡ��).Checked = False
            End If
            Record.Item(mMCol.����).Value = Nvl(rsTmp("����"))
            Record.Item(mMCol.����).Value = Nvl(rsTmp("����"))
            Record.Item(mMCol.����).Value = Nvl(rsTmp("��������"))
                        
        End With
        rsTmp.MoveNext
    Loop
    Me.rptMachine.Populate
    
    With Me.rptList.Columns
        
        rptList.AllowColumnRemove = False
        rptList.ShowItemsInGroups = False
        
        With rptList.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
        rptList.SetImageList ImgList
        
        
        Set Column = .Add(mCol.ѡ��, "ѡ��", 18, False): Column.Icon = 0
        Set Column = .Add(mCol.����, "����", 18, False): Column.Icon = 1
        Set Column = .Add(mCol.ִ��״̬, "״̬", 18, False): Column.Icon = 4
        
        Set Column = .Add(mCol.�걾��, "�걾��", 80, True)
        Set Column = .Add(mCol.��������, "��������", 65, True)
        Set Column = .Add(mCol.�걾����, "�걾����", 65, True)
        Set Column = .Add(mCol.����ʱ��, "����ʱ��", 80, True)
        Set Column = .Add(mCol.������, "������", 65, True)
        Set Column = .Add(mCol.������, "������", 65, True)
        Set Column = .Add(mCol.����ʱ��, "����ʱ��", 65, True)
        Set Column = .Add(mCol.������, "������", 65, True)
        Set Column = .Add(mCol.�������, "�������", 65, True)
        Set Column = .Add(mCol.��������, "��������", 65, True)
        Set Column = .Add(mCol.ִ�п���, "ִ�п���", 65, True)
        Set Column = .Add(mCol.ҽ��id, "ҽ��id", 65, True): Column.Visible = False
        Set Column = .Add(mCol.���ͺ�, "���ͺ�", 65, True): Column.Visible = False
        Set Column = .Add(mCol.ת��, "ת��", 65, True): Column.Visible = False
        Set Column = .Add(mCol.�걾id, "�걾ID", 65, True): Column.Visible = False
        Set Column = .Add(mCol.����ID, "����ID", 65, True): Column.Visible = False
        Set Column = .Add(mCol.�Ƿ����, "�Ƿ����", 65, True): Column.Visible = False
        Set Column = .Add(mCol.���ID, "���Id", 65, True): Column.Visible = False
        Set Column = .Add(mCol.�걾���, "�걾���", 65, True): Column.Visible = False
        Set Column = .Add(mCol.����id, "����ID", 65, True): Column.Visible = False
        Set Column = .Add(mCol.Ӥ��, "Ӥ��", 65, True): Column.Visible = False
        Set Column = .Add(mCol.��������ID, "��������ID", 65, True): Column.Visible = False
        Set Column = .Add(mCol.������, "������", 65, True): Column.Visible = False
        Set Column = .Add(mCol.��ҳID, "��ҳID", 65, True): Column.Visible = False
        Set Column = .Add(mCol.������, "������", 65, True): Column.Visible = False
        Set Column = .Add(mCol.����ʱ��, "����ʱ��", 80, True): Column.Visible = False
        
        Me.rptMachine.Populate
    End With
   
    Select Case mintEditType
        Case 1
            Me.Caption = "������ӡ"
        Case 2
            Me.Caption = "�������"
        Case 3
            Me.Caption = "����ɾ������"
        Case 4
            Me.Caption = "�����޸ı걾��"
    End Select
    If mintEditType = 4 Then
        Me.Label8.Top = Me.LabModify(0).Top + Me.LabModify(0).Height + 100
        chkPrint.Top = cboVerifyMan.Top
        chkPrint.Left = Label7.Left
        
        For intLoop = 0 To 2
            optSort(intLoop).Visible = False
        Next
    Else
        Me.Label8.Top = Frame1.Top + Frame1.Height + 100
    End If
    Me.rptMachine.Top = Me.Label8.Top + Me.Label8.Height + 100
    
    Me.chkUnion.Value = zlDatabase.GetPara("frmBatchAction_����ӡ���ϲ��걾", 100, 1208, 1)
    Me.chkUnion.Value = zlDatabase.GetPara("frmBatchAction_ͬһ�����˺ϲ�Ϊһ�����浥��ӡ", 100, 1208, 0)
    mintUnion = zlDatabase.GetPara("������������ʾ������Ŀ", 100, 1208, 0)
    mMakeNoRule = zlDatabase.GetPara("�걾������ɹ���", 100, 1208, "��  ��")
    mstrPrintDepts = zlDatabase.GetPara("ֻ��ָ�����ұ��浥", 100, 1208, "")
    cboDate.ListIndex = Val(zlDatabase.GetPara("������ӡʱ������", 100, 1208, 0))
    
    intSort = zlDatabase.GetPara("������ӡ����������", 100, 1208, 0)
    If intSort >= 0 And intSort <= 2 Then
        optSort(intSort).Value = True
    Else
        optSort(0).Value = True
    End If
    
    Me.chkAbnormal.Value = zlDatabase.GetPara("frmBatchAction_�����쳣����걾��ʾ", 100, 1208, 0)
    
    Call RestoreWinState(Me, App.ProductName)                   '����ָ�
End Sub

Private Sub Form_Resize()
    Dim Pane1 As Pane
    On Error Resume Next

    If Me.WindowState = 1 Then Exit Sub

    Set Pane1 = Me.dkpMain.FindPane(Dkp_ID_Left)
    Pane1.MinTrackSize.SetSize mlngLeftWidth / Screen.TwipsPerPixelX, Pane1.MaxTrackSize.Height
    Pane1.MaxTrackSize.SetSize mlngLeftWidth / Screen.TwipsPerPixelX, Pane1.MaxTrackSize.Height
    
    
    Me.dkpMain.RecalcLayout
    Me.dkpMain.NormalizeSplitters
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strTmp As String
    Dim intLoop As Integer
    If mblnExec = True Then
        If mintEditType = 1 Then
            strTmp = "��ӡ"
        ElseIf mintEditType = 2 Then
            strTmp = "���"
        Else
            strTmp = "����ɾ��"
        End If
        MsgBox "����ִ��<" & strTmp & ">���������˳�!", vbInformation, Me.Caption
        Cancel = True
        Exit Sub
    End If
        
    zlDatabase.SetPara "frmBatchAction_����ӡ���ϲ��걾", Me.chkUnion.Value, 100, 1208
    zlDatabase.SetPara "frmBatchAction_ͬһ�����˺ϲ�Ϊһ�����浥��ӡ", Me.chkPatient.Value, 100, 1208
    zlDatabase.SetPara "������ӡʱ������", Me.cboDate.ListIndex, 100, 1208

    For intLoop = 0 To 2
        If Me.optSort(intLoop).Value = True Then
            Exit For
        End If
    Next
    zlDatabase.SetPara "������ӡ����������", intLoop, 100, 1208
    zlDatabase.SetPara "frmBatchAction_�����쳣����걾��ʾ", Me.chkAbnormal.Value, 100, 1208
    frmLabMain.zlRefreshData
End Sub

Private Sub Option1_Click()

End Sub

Private Sub picList_Resize()
    With Me.rptList
        .Left = 20
        .Width = Me.picList.ScaleWidth - 40
        .Height = Me.picList.ScaleHeight - .Top - 20
    End With
End Sub

Private Sub picWhere_Resize()
    Me.rptMachine.Height = Me.picWhere.ScaleHeight - Me.rptMachine.Top - 100
End Sub
Private Sub RefreshData()
    '����                   'ˢ������
    Dim rsTmp As New ADODB.Recordset
    Dim intLoop As Integer, lngloop As Long
    Dim Record As ReportRecord
    Dim varItem As Variant
    Dim varBetween As Variant
    Dim blnMoved As Boolean
    Dim strSQLbak As String
    Dim strTmp As String
    Dim str���� As String, str����Where As String, str������ʾ As String, lng������Դ As Long
    Dim strSQL As String, i As Integer, rsTmp1 As ADODB.Recordset
    Dim intCount As Integer
    Dim blnShow As Boolean
    
    On Error GoTo errH
    
    With Me.rptList
        .Records.DeleteAll
        .Populate
    End With
    '10765  ���鿪ʼʱ��ͽ���ʱ��
    If DtpBegin.Value > DtpEnd.Value Then
        MsgBox "��ʼ���ڲ��ܴ��ڽ������ڣ�", vbInformation, gstrSysName
        DtpBegin.SetFocus
        Exit Sub
    End If
    blnMoved = MovedByDate(Me.DtpBegin.Value)
    '����
    If Me.cbo����.ItemData(Me.cbo����.ListIndex) > 0 Then
        str���� = " ,(Select K.���� As ����,J.����id,J.����ID From �������Ҷ�Ӧ J,���ű� K Where J.����id=K.ID) I   "
        str����Where = " And B.���˿���id=I.����id And I.����id= [8] "
        str������ʾ = ",����ID    "
    Else
        If optSort(2).Value = True Then
            str���� = " ,(Select K.���� As ����,J.����id,J.����ID From �������Ҷ�Ӧ J,���ű� K Where J.����id=K.ID) I   "
            str����Where = " And B.���˿���id=I.����id  "
            str������ʾ = ",����ID    "
        Else
            str���� = ""
            str����Where = ""
        End If
    End If
    gstrSql = "select /*+ RULE */ DISTINCT B.���ID AS ID,A.ҽ��id,F.���ͺ�,0 AS ѡ��," & _
                      " Decode(A.����id, Null, " & vbCrLf & _
                        " to_Char(Trunc(A.�걾���/10000)+1,'0000')|| '-'||to_Char(MOD(A.�걾���,10000),'0000'), A.�걾���) As �걾��, " & _
                      "A.�걾����," & _
                      "TO_CHAR(A.����ʱ��,'MM-DD HH24:MI') AS ����ʱ��," & _
                      "A.������," & _
                      "A.������," & _
                      "lpad(A.�걾���,8,'0') as ����," & _
                      "TO_CHAR(B.����ʱ��,'MM-DD HH24:MI') AS ����ʱ��," & _
                      "B.����ҽ�� AS ������," & _
                      "C.���� AS �������," & _
                      "E.���� AS ִ�п���," & _
                      "A.id as �걾ID, a.������,a.����ʱ��, " & _
                      "B.����id, " & _
                      "D.���� AS ��������,0 As ת��,Decode(A.�걾���,1,'��','') As ����, " & _
                      "decode(a.���ʱ��,Null,'��','��') as �Ƿ����, " & _
                      "Decode(a.����״̬, 1, '������', 2, '�Ѽ���') As ִ��״̬, " & _
                      "Decode(a.�Ƿ���, 1, '', '����ʧ��') As ����, a.��ӡ����,a.΢����걾, " & _
                      "a.����,a.�걾���,a.����ID,a.������Դ,a.Ӥ��,b.��������ID,a.������,b.��ҳID  " & str������ʾ & _
                 "from ����걾��¼ A, ����ҽ����¼ B, ���ű� C, �������� D,���ű� E,����ҽ������ F,������Ϣ G, " & _
                 " (Select * From Table(Cast(f_str2list([6]) As zltools.t_strlist))) H " & _
                  str���� & _
                 " WHERE A.ҽ��ID = B.���ID(+) AND B.��������ID = C.ID(+) AND B.ID=F.ҽ��id(+) AND " & _
                      "A.����ID = D.ID(+) AND B.ִ�п���id = E.ID(+) AND A.����״̬ IN (1,2) AND a.����ID = G.����ID(+)  " & _
                      "  " & str����Where
                      
    '����ʹ�ú��ջ��Ǳ���ʱ��
    If cboDate.Text = "����ʱ��" Then
        gstrSql = gstrSql & " and ����ʱ�� between [1] and [2] "
    Else
        gstrSql = gstrSql & " and ����ʱ�� between [1] and [2] "
    End If
                      
    Select Case mintEditType
        Case 1
            '-------- ��ú����
            If Me.chkPrint.Value <> 1 Then
                gstrSql = gstrSql & " And nvl(a.��ӡ����,0) = 0 "
            End If
            '-------- ��ú����

            gstrSql = gstrSql & " and a.����״̬ in (1,2)  and a.���� is not null " & _
                                IIf(Me.chkUnion.Value = 1, " and nvl(a.�ϲ�ID,0) = 0 ", "")
            If InStr(mstrPrivs, "δ��˴�ӡ") <= 0 Then
                gstrSql = gstrSql & " And ����״̬ = 2 "
            End If
        Case 2
            gstrSql = gstrSql & " and a.����״̬ = 1  and a.����  is not null"
        Case 3
            gstrSql = gstrSql & " and a.���� is null  and nvl(�Ƿ��ʿ�Ʒ,0) = 0  "
    End Select
    
    '�������
    If Me.cboRequisitionDept.ItemData(Me.cboRequisitionDept.ListIndex) > 0 Then
        gstrSql = gstrSql & " And b.��������ID = [3] "
    End If
    
    'ִ�п���
    If Me.cboExeDept.ItemData(Me.cboExeDept.ListIndex) > 0 Then
        gstrSql = gstrSql & " and a.ִ�п���Id = [4] "
    End If
    
    '������
    If Me.cboVerifyMan.ItemData(Me.cboVerifyMan.ListIndex) > 0 Then
        gstrSql = gstrSql & " and a.������ = [5] "
        
    End If
    
    '������Դ
    If Me.cbo��Դ.ListIndex > 0 Then
        gstrSql = gstrSql & " And a.������Դ=[9] "
        
        Select Case Me.cbo��Դ.List(Me.cbo��Դ.ListIndex)
        Case "����": lng������Դ = 1
        Case "סԺ": lng������Դ = 2
        Case "����": lng������Դ = 3
        Case "���": lng������Դ = 4
        End Select
    End If
    
    'ֻ��ʾ����ҽ��
    
    
    '����걾��
    If Trim(TxtSample.Text) <> "" Then
        TxtSample.Text = Replace(Replace(TxtSample.Text, "��", "~"), "-", "~")
        If Check_Sample = False Then Exit Sub '10861
        varItem = Split(Trim(TxtSample.Text), ",")
        For lngloop = 0 To UBound(varItem)
            varBetween = Split(varItem(lngloop), "~")
            If UBound(varBetween) > 0 Then
                strTmp = strTmp & "  OR lpad(A.�걾���,8,'0') BETWEEN lpad(" & IIf(Trim(Me.txtBatchNum) <> "", TransSampleNO(Val(Me.txtBatchNum) & "-" & Val(varBetween(0))), Val(varBetween(0))) & _
                        ",8,'0') AND lpad(" & IIf(Trim(Me.txtBatchNum) <> "", TransSampleNO(Val(Me.txtBatchNum) & "-" & Val(varBetween(1))), Val(varBetween(1))) & ",8,'0')"
            Else
                strTmp = strTmp & " OR A.�걾���='" & IIf(Trim(Me.txtBatchNum) <> "", TransSampleNO(Val(Me.txtBatchNum) & "-" & Val(varItem(lngloop))), Val(varItem(lngloop))) & "'"
            End If
        Next
            
    Else
        'ֻ������ʱѡ������
        If Trim(Me.txtBatchNum) <> "" Then
            strTmp = strTmp & " or a.�걾��� between " & TransSampleNO(Val(Me.txtBatchNum) & "-0001") & " And " & TransSampleNO(Val(Me.txtBatchNum) & "-9999")
        End If
    End If
                              
    If strTmp <> "" Then gstrSql = gstrSql & " AND (1=2 " & strTmp & ")"
    strTmp = ""
    
    With Me.rptMachine
        For intLoop = 0 To .Records.Count - 1
            If .Records(intLoop).Item(mMCol.ѡ��).Checked = True Then
                strTmp = strTmp & "," & .Records(intLoop).Item(mMCol.ID).Value
            End If
        Next
    End With
    
    If strTmp = "" Then
        MsgBox "��ѡ��һ���豸��", vbInformation
        Exit Sub
    Else
        strTmp = Mid(strTmp, 2)
    End If
    
    gstrSql = gstrSql & " And nvl(a.����ID,0) = h.Column_Value  "
    
    
    If InStr(mstrPrivs, "�������") > 0 And mintEditType = 2 Then
        '--- 20007-08-30 10783 �������ʱ������˲��ܺͼ�������ͬ
        gstrSql = gstrSql & " And A.������ <> [7] "
    End If
    
    zlCommFun.ShowFlash "����ˢ���������Ժ�..."
    Me.MousePointer = 11
    
    
'    If blnMoved Then
'        strSQLBak = gstrSql
'        strSQLBak = Replace(strSQLBak, "0 As ת��", "1 As ת��")
'        strSQLBak = Replace(strSQLBak, "����ҽ����¼", "H����ҽ����¼")
'        strSQLBak = Replace(strSQLBak, "����ҽ������", "H����ҽ������")
'        strSQLBak = Replace(strSQLBak, "����걾��¼", "H����걾��¼")
'        gstrSql = gstrSql & " Union ALL " & strSQLBak
'    End If
    '��������
    
    If optSort(0).Value = True Then
        gstrSql = gstrSql & "  Order by " & " ���� "
    ElseIf optSort(1).Value = True Then
        gstrSql = gstrSql & "  Order by " & "����ID,���� "
    ElseIf optSort(2).Value = True Then
        gstrSql = gstrSql & "  Order by " & "����ID,���� "
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CDate(Format(Me.DtpBegin, "yyyy-mm-dd 00:00:00")), _
                                         CDate(Format(Me.DtpEnd, "yyyy-mm-dd 23:23:59")), _
                                         CLng(Me.cboRequisitionDept.ItemData(Me.cboRequisitionDept.ListIndex)), _
                                         CLng(Me.cboExeDept.ItemData(Me.cboExeDept.ListIndex)), _
                                         CStr(Me.cboVerifyMan.Text), strTmp, IIf(mstrAuditingMan = "", UserInfo.����, mstrAuditingMan), _
                                         CLng(Me.cbo����.ItemData(Me.cbo����.ListIndex)), lng������Դ)
                                         
    With Me.rptList
        .Records.DeleteAll
        .Populate
    End With
    
    Do Until rsTmp.EOF
        With Me.rptList
            If chkAbnormal.Visible = True And chkAbnormal.Value = 1 Then
                blnShow = Not CheckAbnormal(rsTmp("�걾ID"))
            Else
                 blnShow = True
            End If
            
            If blnShow = True Then
        
                Set Record = .Records.Add
                For intLoop = 0 To .Columns.Count
                    Record.AddItem ""
                Next
                
                Record.Item(mCol.ѡ��).HasCheckbox = True
                Record.Item(mCol.���ID).Value = Nvl(rsTmp("ID"))
                Record.Item(mCol.�걾id).Value = Nvl(rsTmp("�걾ID"))
                Record.Item(mCol.�걾��).Value = Val(Nvl(rsTmp("�걾���")))
                Record.Item(mCol.�걾��).Caption = Trim(Nvl(rsTmp("�걾��")))
                '-----------------------------------------------------��ú��
                If CInt(Nvl(rsTmp("��ӡ����"), "0")) > 0 Then
                    Record.Item(mCol.ִ��״̬).Value = "�Ѵ�ӡ"
                    Record.Item(mCol.ִ��״̬).Icon = 7
                ElseIf Nvl(rsTmp("ִ��״̬")) = "�Ѽ���" Then
                    Record.Item(mCol.ִ��״̬).Value = "�Ѽ���"
                    Record.Item(mCol.ִ��״̬).Icon = 6
                ElseIf Nvl(rsTmp("����")) = "" Then
                    Record.Item(mCol.ִ��״̬).Value = "�Ѵ���"
                    Record.Item(mCol.ִ��״̬).Icon = 5
                End If
                
                If Val("" & rsTmp!΢����걾) = 0 Then
                    strSQL = "Select Count(A.ID) - Sum(Decode(A.������, Null, 0, 1)) As �޽����¼,Count(A.ID) as ����� " & vbNewLine & _
                            "From ������ͨ��� A" & vbNewLine & _
                            "Where A.����걾id = [1]"
                    Set rsTmp1 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val("" & rsTmp("�걾ID")))
                    If rsTmp1.EOF Then
                        For i = 0 To rptList.Columns.Count - 1
                            Record.Item(i).BackColor = vbWhite
                        Next
                    Else
                        If Val("" & rsTmp1.Fields("�޽����¼")) = 0 And Val("" & rsTmp1.Fields("�����")) > 0 Then
                            For i = 0 To rptList.Columns.Count - 1
                                Record.Item(i).BackColor = &HFDD6C6
                            Next
                        Else
                            For i = 0 To rptList.Columns.Count - 1
                                Record.Item(i).BackColor = vbWhite
                            Next
    
                        End If
                    End If
                Else
                    For i = 0 To rptList.Columns.Count - 1
                        Record.Item(i).BackColor = vbWhite
                    Next
                End If
                
                '-----------------------------------------------------
                
                Record.Item(mCol.�걾����).Value = Nvl(rsTmp("�걾����"))
                Record.Item(mCol.����ID).Value = Nvl(rsTmp("����ID"))
                Record.Item(mCol.��������).Value = Nvl(rsTmp("����"))
                Record.Item(mCol.���ͺ�).Value = Nvl(rsTmp("���ͺ�"))
                Record.Item(mCol.������).Value = Nvl(rsTmp("������"))
                Record.Item(mCol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
                Record.Item(mCol.����).Icon = IIf(Nvl(rsTmp("����")) = "��", 2, -1)
                Record.Item(mCol.������).Value = Nvl(rsTmp("������"))
                Record.Item(mCol.��������).Value = Nvl(rsTmp("��������"))
                Record.Item(mCol.�������).Value = Nvl(rsTmp("�������"))
                Record.Item(mCol.������).Value = Nvl(rsTmp("������"))
                Record.Item(mCol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
                Record.Item(mCol.�Ƿ����).Value = Nvl(rsTmp("�Ƿ����"))
                Record.Item(mCol.ҽ��id).Value = Nvl(rsTmp("ҽ��ID"))
                Record.Item(mCol.ִ�п���).Value = Nvl(rsTmp("ִ�п���"))
                Record.Item(mCol.ת��).Value = Nvl(rsTmp("ת��"))
                Record.Item(mCol.�걾���).Value = Nvl(rsTmp("�걾���"))
                Record.Item(mCol.����id).Value = Nvl(rsTmp("����ID"))
                Record.Item(mCol.������Դ).Value = Nvl(rsTmp("������Դ"), 3)
                Record.Item(mCol.Ӥ��).Value = Nvl(rsTmp("Ӥ��"), 0)
                Record.Item(mCol.��������ID).Value = Nvl(rsTmp("��������ID"))
                Record.Item(mCol.������).Value = Nvl(rsTmp("������"))
                Record.Item(mCol.��ҳID).Value = Nvl(rsTmp("��ҳID"))
                Record.Item(mCol.������).Value = Nvl(rsTmp("������"))
                Record.Item(mCol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
            Else
                intCount = intCount + 1
            End If
        End With
        rsTmp.MoveNext
    Loop
    Me.rptList.Populate
    zlCommFun.StopFlash
    stbThis.Panels(2).Text = "��ǰ���ҵ�" & rsTmp.RecordCount & "����¼��"
    Call chkfilter_Click(0)
    Me.MousePointer = 0
    Exit Sub
errH:
    zlCommFun.StopFlash
    Me.MousePointer = 0
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub



Private Function CheckAbnormal(lngSample As Long) As Boolean
    '����             �����Ƿ�걾���쳣�ģ���ʾ���޺����� �����־=5��6��
    '����             ���쳣ʱ����Ϊ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "select count(id) id from ������ͨ��� where ����걾id = [1]  and �����־ in (5,6) "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���걾�Ƿ����쳣����", lngSample)
    If rsTmp.RecordCount > 0 Then
        If Val(rsTmp("id") & "") > 0 Then
            CheckAbnormal = True
            Exit Function
        End If
    End If
    
End Function
Private Sub RptSelect(Records As ReportRecords, blTrue As Boolean)
    '����                           ѡ���ȡ��ѡ��
    '����                           Records = �б����
    '                               blTrue  True = ѡ�� False = ȡ��ѡ��
    Dim intLoop As Integer
    Me.chkfilter(0).Value = IIf(blTrue, 1, 0)
    Me.chkfilter(1).Value = IIf(blTrue, 1, 0)
    For intLoop = 0 To Records.Count - 1
        Records(intLoop).Item(mCol.ѡ��).Checked = blTrue
    Next
End Sub


Private Sub SaveData(Optional blnPrintNoAuditing As Boolean = False)
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim blPrint As Boolean
    Dim blnAutoPrint As Boolean
    Dim lngloop As Long
    Dim bln���ͨ��  As Boolean
    Dim strMsg As String '��ʾδ���ͨ������ʾ��Ϣ
    Dim strErrInfo As String
    Dim rsTmp As ADODB.Recordset
    Dim intLoop As Integer
    Dim strChart(1 To 9) As String
    Dim intPrivacy As Integer                               '���ͱ��浥��ҽ��վʱ�Ƿ���ʾ��˽��Ŀ
    Dim blnCheckExesState As Boolean                        '��Ժ���˴��ڼ��ʻ��۷���
    Dim lngAdvice As Long                                   'ҽ��ID
    Dim intPrintCount As Integer                            '��ӡ����
    Dim intOnPrint As Integer                               '���м���δ��˵ı걾δ��ӡ
    Dim lngPatient As Long                                  '��¼����ID
    Dim strҽ��ID As String                                 'ҽ��ID��","�ָ�
    Dim str�걾ID As String                                 '�걾ID��","�ָ�
    Dim intItem As Integer                                  '��ʱ��¼
    Dim astrItem() As String                                '�������ڼ�¼ID
    Dim strPrintCode As String                              '��ӡ���ݱ���
    Dim lngҽ��ID As Long                                   'ҽ��ID
    Dim lng�걾ID As Long                                   '�걾ID
    Dim lng����ID As Long                                   '����ID
    Dim lng�������ID As Long                               '�������Id
    Dim blnRollBack As Boolean                              '�Ƿ�ع�
    Dim astrSQL() As String                                 'Ҫִ�е�����
    Dim strTmp() As String
    Dim blngǿ����� As Boolean                             'ǿ�����ͨ��
    Dim strDate As String                                   '�ɼ�ʱ�䲻��ͨ��
    
    On Error GoTo ErrHand
    
    ReDim astrSQL(0)
    Me.MousePointer = 11
    mblnExec = True
    blnAutoPrint = zlDatabase.GetPara("��˴�ӡ", 100, 1208, 0)
    '��д���µĵ��Ӳ�����
    intPrivacy = zlDatabase.GetPara("���浥�Ƿ���ʾ��˽��Ŀ", 100, 1208, 0)
    intPrintCount = 0
    intOnPrint = 0
    blPrint = blnPrintNoAuditing
    
    If Me.chkPatient.Value = 1 Then
        Me.rptList.SortOrder.DeleteAll
        Me.rptList.SortOrder.Add Me.rptList.Columns(mCol.����ID)
        Me.rptList.Populate
    End If
    
    With Me.rptList
        For lngloop = 0 To .Records.Count - 1
            
            If .Records(lngloop).Item(mCol.ѡ��).Checked = True And Val(.Records(lngloop).Item(mCol.���ID).Value) > 0 Then
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                lng�������ID = Val(.Records(lngloop).Item(mCol.��������ID).Value)
                If mintEditType = 1 And InStr("," & mstrPrintDepts & ",", lng�������ID) > 0 Then
                    If Me.chkPatient.Value = 0 Then
                        '==����ǰ�걾���д�ӡ
                        '����ͼ�ι��Զ��屨�����
                        gstrSql = "select id from ����ͼ���� where �걾id = [1] order by id"
                        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(.Records(lngloop).Item(mCol.�걾id).Value))
                        
                        For intLoop = 1 To 9
                            strChart(intLoop) = ""
                        Next
                        intLoop = 1
                        Do Until rsTmp.EOF
                            If intLoop > 9 Then Exit For
                            strChart(intLoop) = App.path & "\" & rsTmp("ID") & ".cht"
                            Debug.Print strChart(intLoop)
                            Call LoadImageData(App.path, rsTmp("ID"))
                            
                            intLoop = intLoop + 1
                            rsTmp.MoveNext
                        Loop
                        If mintEditType = 1 Then '��ӡ
                            zlCommFun.ShowFlash "���ڴ�ӡ����,�����(" & lngloop + 1 & "/" & .Records.Count & ")"
                            If GetReportCode(Val(.Records(lngloop).Item(mCol.ҽ��id).Value), Val(.Records(lngloop).Item(mCol.���ͺ�).Value), strReportCode, strReportParaNo, bytReportParaMode, _
                                Val(.Records(lngloop).Item(mCol.ת��).Value) = 1) Then
                                
                                If .Records(lngloop).Item(mCol.�Ƿ����).Value = "��" And blPrint = False Then
                                    intOnPrint = intOnPrint + 1
                                Else
                                    If .Records(lngloop).Item(mCol.�Ƿ����).Value = "��" Or InStr(mstrPrivs, "δ��˴�ӡ") > 0 Then
                                        If intPrintCount = 0 Then Call ReportTaskBegin
                                        Call ReportOpen(gcnOracle, glngSys, strReportCode, Me, "NO=" & strReportParaNo, "����=" & bytReportParaMode, _
                                            "ҽ��ID=" & Val(.Records(lngloop).Item(mCol.ҽ��id).Value), _
                                            "����ID=" & Val(.Records(lngloop).Item(mCol.����ID).Value), _
                                            "�걾ID=" & Val(.Records(lngloop).Item(mCol.�걾id).Value), _
                                            "ͼ��1=" & strChart(1), "ͼ��2=" & strChart(2), "ͼ��3=" & strChart(3), "ͼ��4=" & strChart(4), _
                                            "ͼ��5=" & strChart(5), "ͼ��6=" & strChart(6), "ͼ��7=" & strChart(7), "ͼ��8=" & strChart(8), _
                                            "ͼ��9=" & strChart(9), 2)
                                            intPrintCount = intPrintCount + 1
                                        .Records(lngloop).Item(mCol.ѡ��).Checked = False
                                        .Populate
                                    End If
                                End If
                                '������˵ı걾����ӡ��־
                                If .Records(lngloop).Item(mCol.�Ƿ����).Value = "��" Then
                                    If mintUnion = 0 Then
                                        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                        astrSQL(UBound(astrSQL)) = "ZL_����걾��¼_�걾�ʿ�(" & Val(.Records(lngloop).Item(mCol.�걾id).Value) & ",'',1)"
                                    Else
                                        gstrSql = "select ID from ����걾��¼ where ҽ��ID = [1] "
                                        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(.Records(lngloop).Item(mCol.ҽ��id).Value))
                                        Do Until rsTmp.EOF
                                            ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                            astrSQL(UBound(astrSQL)) = "ZL_����걾��¼_�걾�ʿ�(" & rsTmp("ID") & ",'',1)"
                                            rsTmp.MoveNext
                                        Loop
                                    End If
                                End If
                            End If
                        End If
                    Else
                        '==�����˽��д�ӡ
                        If lngPatient <> Val(.Records(lngloop).Item(mCol.����ID).Value) Then
                            If str�걾ID <> "" Then
                                intLoop = 1
                                For intLoop = 1 To 9
                                    strChart(intLoop) = ""
                                Next
                                str�걾ID = Mid(str�걾ID, 2)
                                strҽ��ID = Mid(strҽ��ID, 2)
                                lngҽ��ID = Split(strҽ��ID, ",")(0)
                                lng�걾ID = Split(str�걾ID, ",")(0)
                                If strPrintCode = "" Then
                                    '�ж����ʽʱ�õ���ʽ
                                    frmLabMainPrintFormat.ShowMe Me, strҽ��ID, strPrintCode
                                End If
                                astrItem = Split(str�걾ID, ",")
                                For intItem = 0 To UBound(astrItem)
                                    gstrSql = "select id from ����ͼ���� where �걾id = [1] order by id"
                                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(astrItem(intItem)))
                                    Do Until rsTmp.EOF
                                        If intLoop > 9 Then Exit For
                                        strChart(intLoop) = App.path & "\" & rsTmp("ID") & ".cht"
                                        Call LoadImageData(App.path, rsTmp("ID"))
                                        intLoop = intLoop + 1
                                        rsTmp.MoveNext
                                    Loop
                                Next
                                If intPrintCount = 0 Then Call ReportTaskBegin
                                Call ReportOpen(gcnOracle, glngSys, strPrintCode, Me, "NO=" & strReportParaNo, "����=" & bytReportParaMode, "ҽ��ID=" & strҽ��ID, _
                                        "����ID=" & lngPatient, "�걾ID=" & str�걾ID, "���ҽ��=" & strҽ��ID, "����걾=" & str�걾ID, _
                                        "ͼ��1=" & strChart(1), "ͼ��2=" & strChart(2), "ͼ��3=" & strChart(3), "ͼ��4=" & strChart(4), _
                                        "ͼ��5=" & strChart(5), "ͼ��6=" & strChart(6), "ͼ��7=" & strChart(7), "ͼ��8=" & strChart(8), _
                                        "ͼ��9=" & strChart(9), 2)
                                intPrintCount = intPrintCount + 1
                                For intItem = 0 To UBound(astrItem)
                                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                    astrSQL(UBound(astrSQL)) = "ZL_����걾��¼_�걾�ʿ�(" & astrItem(intItem) & ",'',1)"
                                Next
                                str�걾ID = "": strҽ��ID = ""
                            End If
                        End If
                        str�걾ID = str�걾ID & "," & Val(.Records(lngloop).Item(mCol.�걾id).Value)
                        strҽ��ID = strҽ��ID & "," & Val(.Records(lngloop).Item(mCol.ҽ��id).Value)
                        lngPatient = Val(.Records(lngloop).Item(mCol.����ID).Value)
                                    
                    End If
                End If
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If mintEditType = 2 Then '���
                    strDate = ""
                    If .Records(lngloop).Item(mCol.������).Value <> "" Then
                        If .Records(lngloop).Item(mCol.����ʱ��).Value <> "" Then
                            If CDate(.Records(lngloop).Item(mCol.����ʱ��).Value) > zlDatabase.Currentdate Then
                                strDate = "UN"
                                strMsg = strMsg & vbNewLine & .Records(lngloop).Item(mCol.�걾��).Value & " " & .Records(lngloop).Item(mCol.��������).Value & " ����ʱ�䣬���ڵ�ǰʱ�䣬���ܽ�����ˣ�"
                            End If
                        End If
                    End If
                    If strDate <> "UN" Then
                        bln���ͨ�� = False
                        zlCommFun.ShowFlash "�����������,�����(" & lngloop + 1 & "/" & .Records.Count & ")"
                        
                        '21137 �ѹ鵵���治��ȡ��
                        gstrSql = "Select Decode(����״̬, 1, '1-�ȴ����', 2, '2-�ܾ����', 3, '3-�������', 4, '4-��鷴��', 5, '5-���鵵') As ����״̬" & vbNewLine & _
                        "From ����걾��¼ A, ������ҳ B ,�����ύ��¼ C" & vbNewLine & _
                        "Where A.����id = B.����id And A.��ҳid = B.��ҳid And A.������Դ = 2 And Nvl(B.����״̬, 0) >= 1 and A.ID=[1] " & vbNewLine & _
                        " And b.����id = c.����Id and B.��ҳid = C.��ҳID "
                        
                        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(.Records(lngloop).Item(mCol.�걾id).Value))
                        If rsTmp.EOF = False Then
                            strMsg = strMsg & vbNewLine & .Records(lngloop).Item(mCol.�걾��).Value & " " & .Records(lngloop).Item(mCol.��������).Value & " ����סԺ�Ĳ������ύ��飬������ˣ�"
                        Else
                            '-------------------------------------------------------------------------------------------
                            If VerifyAuditingRule(Val(.Records(lngloop).Item(mCol.�걾id).Value), strErrInfo, 2) = 1 Then
                                strErrInfo = ""
                                
                                blngǿ����� = (InStr(mstrPrivs, "����ǿ����˹���") > 0)
                                If blngǿ����� = True Then
                                    strMsg = strMsg & vbNewLine & .Records(lngloop).Item(mCol.�걾��).Value & " " & .Records(lngloop).Item(mCol.��������).Value & " �ļ�����ʹ��ǿ�����Ȩ��ͨ����ˣ�" & vbNewLine & strErrInfo
                                Else
                                    strMsg = strMsg & vbNewLine & .Records(lngloop).Item(mCol.�걾��).Value & " " & .Records(lngloop).Item(mCol.��������).Value & " �ļ�����δͨ��ˣ�" & vbNewLine & strErrInfo
                                End If
                            Else
                                blngǿ����� = True
                            End If
                            If blngǿ����� = True Then
                                If CheckChargeState(Val(.Records(lngloop).Item(mCol.���ID).Value), False) = False Then
                                    blnCheckExesState = CheckExesState(Val(.Records(lngloop).Item(mCol.�걾id).Value))
                                    If mintUnion = 0 Then
                                        'δ�շ�
                                        If InStr(mstrPrivs, "δ�շ����") > 0 And blnCheckExesState = True Then
                                            'ǩ�����ɹ�ʱ�˳�
                                            ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                            astrSQL(UBound(astrSQL)) = "Signature;" & Val(.Records(lngloop).Item(mCol.�걾id).Value) & ";" & mstrAuditingManID
                                            
                                            ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                            astrSQL(UBound(astrSQL)) = "ZL_����걾��¼_�������(" & Val(.Records(lngloop).Item(mCol.�걾id).Value) & ",'" & _
                                                                         IIf(mstrAuditingMan = "", UserInfo.����, mstrAuditingMan) & "','" & _
                                                                         UserInfo.��� & "','" & UserInfo.���� & "')"
                                            
                                            bln���ͨ�� = True
                                            
                                            ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                            astrSQL(UBound(astrSQL)) = "Zl_���鱨�浥_Update(" & Val(.Records(lngloop).Item(mCol.�걾id).Value) & "," & intPrivacy & ",'" & gstrUnitName & "')"           '��˺��������浥
                                            

                                            
                                        Else
                                            strMsg = strMsg & vbNewLine & .Records(lngloop).Item(mCol.�걾��).Value & " " & .Records(lngloop).Item(mCol.��������).Value & _
                                                        IIf(blnCheckExesState, " δ�շѣ�", " ��Ժ���˼��ʻ��۷��ò������")
                                        End If
                                    Else
                                        gstrSql = "select ID from ����걾��¼ where ҽ��ID = [1] "
                                        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(.Records(lngloop).Item(mCol.ҽ��id).Value))
                                        Do Until rsTmp.EOF
                                            'δ�շ�
                                            If InStr(mstrPrivs, "δ�շ����") > 0 And blnCheckExesState = True Then
                                                'ǩ�����ɹ�ʱ�˳�
                                                ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                                astrSQL(UBound(astrSQL)) = "Signature;" & Val(.Records(lngloop).Item(mCol.�걾id).Value) & ";" & mstrAuditingManID
                                                
                                                ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                                astrSQL(UBound(astrSQL)) = "ZL_����걾��¼_�������(" & rsTmp("ID") & ",'" & _
                                                                             IIf(mstrAuditingMan = "", UserInfo.����, mstrAuditingMan) & "','" & _
                                                                             UserInfo.��� & "','" & UserInfo.���� & "')"

                                                bln���ͨ�� = True
                                                ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                                astrSQL(UBound(astrSQL)) = "Zl_���鱨�浥_Update(" & rsTmp("ID") & "," & intPrivacy & ",'" & gstrUnitName & "')"            '��˺��������浥

                                                
                                            Else
                                                strMsg = strMsg & vbNewLine & .Records(lngloop).Item(mCol.�걾��).Value & " " & .Records(lngloop).Item(mCol.��������).Value & _
                                                            IIf(blnCheckExesState, " δ�շѣ�", " ��Ժ���˼��ʻ��۷��ò������")
                                            End If
                                            rsTmp.MoveNext
                                        Loop
                                    End If 'δ�շ� End
                                Else  '����շ�״̬
                                    If mintUnion = 0 Then
                                        'ǩ�����ɹ�ʱ�˳�
                                        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                        astrSQL(UBound(astrSQL)) = "Signature;" & Val(.Records(lngloop).Item(mCol.�걾id).Value) & ";" & mstrAuditingManID
                                        
                                        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                        astrSQL(UBound(astrSQL)) = "ZL_����걾��¼_�������(" & Val(.Records(lngloop).Item(mCol.�걾id).Value) & ",'" & _
                                                                         IIf(mstrAuditingMan = "", UserInfo.����, mstrAuditingMan) & "','" & _
                                                                         UserInfo.��� & "','" & UserInfo.���� & "')"

                                        bln���ͨ�� = True
                                        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                        astrSQL(UBound(astrSQL)) = "Zl_���鱨�浥_Update(" & Val(.Records(lngloop).Item(mCol.�걾id).Value) & "," & intPrivacy & ",'" & gstrUnitName & "')"           '��˺��������浥

                                        
                                    Else
                                        gstrSql = "select ID from ����걾��¼ where ҽ��ID = [1] "
                                        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(.Records(lngloop).Item(mCol.ҽ��id).Value))
                                        Do Until rsTmp.EOF
                                           'ǩ�����ɹ�ʱ�˳�
                                            ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                            astrSQL(UBound(astrSQL)) = "Signature;" & Val(.Records(lngloop).Item(mCol.�걾id).Value) & ";" & mstrAuditingManID
                                            
                                            ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                            astrSQL(UBound(astrSQL)) = "ZL_����걾��¼_�������(" & rsTmp("ID") & ",'" & _
                                                                         IIf(mstrAuditingMan = "", UserInfo.����, mstrAuditingMan) & "','" & _
                                                                         UserInfo.��� & "','" & UserInfo.���� & "')"
 
                                            bln���ͨ�� = True
                                            ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                            astrSQL(UBound(astrSQL)) = "Zl_���鱨�浥_Update(" & rsTmp("ID") & "," & intPrivacy & ",'" & gstrUnitName & "')"         '��˺��������浥

                                            
                                            rsTmp.MoveNext
                                        Loop
                                    End If  '�Ƿ���������������ʾ End
                                End If '����շ�״̬ End
                            End If '��˹��� End
                        End If
                        '-------------------------------------------------------------------------------------------
                    End If
                    If blnAutoPrint And bln���ͨ�� And InStr("," & mstrPrintDepts & ",", "," & lng�������ID & ",") > 0 Then

                        If GetReportCode(Val(.Records(lngloop).Item(mCol.ҽ��id).Value), Val(.Records(lngloop).Item(mCol.���ͺ�).Value), strReportCode, strReportParaNo, bytReportParaMode, _
                             False) Then
                            '����ͼ�ι��Զ��屨�����
                            'frmLabMainImage.zlRefresh .Records(lngLoop).Item(mCol.�걾ID).Value, True
'                            frmLabMain.ReadImageData .Records(lngLoop).Item(mCol.�걾ID).Value, True
                            gstrSql = "select id from ����ͼ���� where �걾id = [1] order by id "
                            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(.Records(lngloop).Item(mCol.�걾id).Value))
                            
                            For intLoop = 1 To 9
                                strChart(intLoop) = ""
                            Next
                            intLoop = 1
                            Do Until rsTmp.EOF
                                If intLoop > 9 Then Exit For
                                strChart(intLoop) = App.path & "\" & rsTmp("ID") & ".cht"
                                Call LoadImageData(App.path, rsTmp("ID"))
                                intLoop = intLoop + 1
                                rsTmp.MoveNext
                            Loop
                            ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                            astrSQL(UBound(astrSQL)) = "��ӡ�ִ�;" & strReportCode & ";;" & strReportParaNo & ";" & bytReportParaMode & ";" & Val(.Records(lngloop).Item(mCol.ҽ��id).Value) & ";" & _
                                            Val(.Records(lngloop).Item(mCol.����ID).Value) & ";" & Val(.Records(lngloop).Item(mCol.�걾id).Value) & _
                                            ";" & strChart(1) & ";" & strChart(2) & ";" & strChart(3) & ";" & strChart(4) & ";" & strChart(5) & ";" & strChart(6) & ";" & strChart(7) & _
                                            ";" & strChart(8) & ";" & strChart(9)
'                            If intPrintCount = 0 Then Call ReportTaskBegin
'                            Call ReportOpen(gcnOracle, glngSys, strReportCode, Me, "NO=" & strReportParaNo, "����=" & bytReportParaMode, _
'                                "ҽ��ID=" & Val(.Records(lngLoop).Item(mCol.ҽ��id).Value), _
'                                "����ID=" & Val(.Records(lngLoop).Item(mCol.����ID).Value), _
'                                "�걾ID=" & Val(.Records(lngLoop).Item(mCol.�걾id).Value), _
'                                "ͼ��1=" & strChart(1), "ͼ��2=" & strChart(2), "ͼ��3=" & strChart(3), "ͼ��4=" & strChart(4), _
'                                "ͼ��5=" & strChart(5), "ͼ��6=" & strChart(6), "ͼ��7=" & strChart(7), "ͼ��8=" & strChart(8), _
'                                "ͼ��9=" & strChart(9), 2)
'                                intPrintCount = intPrintCount + 1
                            '��Ǵ�ӡ
                            If mintUnion = 0 Then
                                ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                astrSQL(UBound(astrSQL)) = "ZL_����걾��¼_�걾�ʿ�(" & Val(.Records(lngloop).Item(mCol.�걾id).Value) & ",'',1)"
                            Else
                                gstrSql = "select ID from ����걾��¼ where ҽ��ID = [1] "
                                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(.Records(lngloop).Item(mCol.ҽ��id).Value))
                                Do Until rsTmp.EOF
                                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                                    astrSQL(UBound(astrSQL)) = "ZL_����걾��¼_�걾�ʿ�(" & rsTmp("ID") & ",'',1)"
                                    rsTmp.MoveNext
                                Loop
                            End If
                        End If
                    End If
                    
                    .Records(lngloop).Item(mCol.ѡ��).Checked = False
                    .Populate
                End If
            End If
            'ɾ������
            If .Records(lngloop).Item(mCol.ѡ��).Checked = True And mintEditType = 3 Then
                ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                astrSQL(UBound(astrSQL)) = "ZL_����걾��¼_�걾ɾ��(" & Val(.Records(lngloop).Item(mCol.�걾id).Value) & ")"
                .Records(lngloop).Item(mCol.ѡ��).Checked = False
                .Populate
            End If
            '10�ű��浥��һ�������д�ӡ
            If intPrintCount >= 10 Then intPrintCount = 0:   Call ReportTaskEnd
            DoEvents
        Next
        
        '===================����ϲ���ӡ�е����һ���걾====================
        If Me.chkPatient.Value = 1 And str�걾ID <> "" And InStr("," & mstrPrintDepts & ",", lng�������ID) > 0 Then
            intLoop = 1
            For intLoop = 1 To 9
                strChart(intLoop) = ""
            Next
            str�걾ID = Mid(str�걾ID, 2)
            strҽ��ID = Mid(strҽ��ID, 2)
            lngҽ��ID = Split(strҽ��ID, ",")(0)
            lng�걾ID = Split(str�걾ID, ",")(0)
            If strPrintCode = "" Then
                '�ж����ʽʱ�õ���ʽ
                frmLabMainPrintFormat.ShowMe Me, strҽ��ID, strPrintCode
            End If
            astrItem = Split(str�걾ID, ",")
            For intItem = 0 To UBound(astrItem)
                gstrSql = "select id from ����ͼ���� where �걾id = [1] order by id"
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(astrItem(intItem)))
                Do Until rsTmp.EOF
                    If intLoop > 9 Then Exit For
                    strChart(intLoop) = App.path & "\" & rsTmp("ID") & ".cht"
                    Call LoadImageData(App.path, rsTmp("ID"))
                    intLoop = intLoop + 1
                    rsTmp.MoveNext
                Loop
            Next
            If intPrintCount = 0 Then Call ReportTaskBegin
            Call ReportOpen(gcnOracle, glngSys, strPrintCode, Me, "NO=" & strReportParaNo, "����=" & bytReportParaMode, "ҽ��ID=" & strҽ��ID, _
                    "����ID=" & lngPatient, "�걾ID=" & str�걾ID, "���ҽ��=" & strҽ��ID, "����걾=" & str�걾ID, _
                    "ͼ��1=" & strChart(1), "ͼ��2=" & strChart(2), "ͼ��3=" & strChart(3), "ͼ��4=" & strChart(4), _
                    "ͼ��5=" & strChart(5), "ͼ��6=" & strChart(6), "ͼ��7=" & strChart(7), "ͼ��8=" & strChart(8), _
                    "ͼ��9=" & strChart(9), 2)
            intPrintCount = intPrintCount + 1
            For intItem = 0 To UBound(astrItem)
                ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                astrSQL(UBound(astrSQL)) = "ZL_����걾��¼_�걾�ʿ�(" & astrItem(intItem) & ",'',1)"
            Next
            str�걾ID = "": strҽ��ID = ""
        End If
        '==================================================================================
        If intPrintCount > 0 Then Call ReportTaskEnd
    End With
    
'    gcnOracle.BeginTrans
'    blnRollBack = True
    
    For lngloop = 1 To UBound(astrSQL)
        If Trim(astrSQL(lngloop)) <> "" Then
            If UCase(Mid(astrSQL(lngloop), 1, 3)) = "ZL_" Then
                zlDatabase.ExecuteProcedure astrSQL(lngloop), Me.Caption
            ElseIf UCase(Mid(astrSQL(lngloop), 1, 4)) = "��ӡ�ִ�" Then
                strTmp = Split(astrSQL(lngloop), ";")
                If intPrintCount = 0 Then Call ReportTaskBegin
                Call ReportOpen(gcnOracle, glngSys, strTmp(1), Me, "NO=" & strTmp(3), "����=" & strTmp(4), "ҽ��ID=" & strTmp(5), "����ID=" & strTmp(6), _
                "�걾ID=" & strTmp(7), "ͼ��1=" & strTmp(8), "ͼ��2=" & strTmp(9), "ͼ��3=" & strTmp(10), _
                , "ͼ��4=" & strTmp(11), "ͼ��5=" & strTmp(12), "ͼ��6=" & strTmp(13), "ͼ��7=" & strTmp(14), "ͼ��8=" & strTmp(15), "ͼ��9=" & strTmp(16), 2)
                intPrintCount = intPrintCount + 1
                
                '10�ű��浥��һ�������д�ӡ
                If intPrintCount >= 10 Then intPrintCount = 0:   Call ReportTaskEnd
            Else
                'ǩ�����ɹ�ʱ�˳�
                If Signature(Val(Split(astrSQL(lngloop), ";")(1)), mstrAuditingManID) = False Then
'                    gcnOracle.RollbackTrans
'                    blnRollBack = False
                     zlCommFun.StopFlash
                    mblnExec = False
                    Exit Sub
                End If
            End If
        End If
    Next
    If intPrintCount > 0 Then Call ReportTaskEnd
    
    blnRollBack = False
'    gcnOracle.CommitTrans
    
    zlCommFun.StopFlash
    Me.MousePointer = 0
    
    If strMsg <> "" Then
        MsgBox "���¼�¼δͨ����ˣ�" & strMsg, vbInformation, Me.Caption
    End If
    
    If intOnPrint > 0 Then
        If MsgBox("����δ��˵ı��浥" & intOnPrint & "�ţ��Ƿ��ӡ?", vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
            Call SaveData(True)
        End If
    End If
    Call RefreshData
    On Error Resume Next
    'ɾ��ͼ���ļ�
    For intLoop = 1 To 9
        If strChart(intLoop) <> "" Then
            Kill strChart(intLoop)
        End If
    Next
    mblnExec = False
    Exit Sub
    
ErrHand:
    If blnRollBack = True Then gcnOracle.RollbackTrans: blnRollBack = False
    zlCommFun.StopFlash
    Me.MousePointer = 0
    mblnExec = False
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub ShowMe(objfrm As Object, intEditType As Integer, Optional lngMachine As Long, Optional strPrivs As String, Optional strAuditingMan As String, _
                  Optional intAuditing As Integer, Optional DateAuditing As Date, Optional DeptID As Long, Optional strAuditingManID As String)
    '----------------------------------------------------------------------------------------------------------------
    '���ܣ���ʾ���༭����
    '������             lngMachine = ����ID strprivs = Ȩ�� strAuditingMan = ������ intAuditing = ʱ������
    '                   DataAuditing = ��ʼʱ��  objfrm =  ������, intEditType = ��������(1=��ӡ 2=��� 3=����ɾ�� 4=�����޸ı걾��)
    '���أ�
    '-----------------------------------------------------------------------------------------------------------------
    
    mintEditType = intEditType
    mlngMachine = lngMachine
    mstrPrivs = strPrivs
    mstrAuditingMan = strAuditingMan
    mstrAuditingManID = strAuditingManID
    mintAuditing = intAuditing
    mDateAuditing = DateAuditing
    mDeptID = DeptID
    stbThis.Panels(2).Text = "׼����"
    If mintEditType = 1 Then
        Me.chkUnion.Visible = True
        Me.chkPatient.Visible = True
    End If
    If mintEditType = 4 Then
        Frame1.Top = Me.chkPatient.Top - 50
        LabModify(0).Top = Frame1.Top + Frame1.Height + 100
        LabModify(1).Top = LabModify(0).Top
        TxtModify.Top = LabModify(0).Top - 50
        Label8.Top = LabModify(0).Top + LabModify(0).Height + 100
        rptMachine.Top = Me.Label8.Top + Me.Label8.Height + 100
        Me.LabModify(0).Visible = True: Me.LabModify(1).Visible = True: Me.TxtModify.Visible = True
    End If
    Me.Show , objfrm
    
End Sub

Private Sub rptList_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim hitColumn As ReportColumn
    Dim Record As ReportRecord
    Dim blSelect As Boolean

    With Me.rptList
        Set hitColumn = .HitTest(X, Y).Column
        If Not hitColumn Is Nothing Then
            If hitColumn.Caption = "ѡ��" And .HitTest(X, Y).ht = xtpHitTestHeader Then
                If .Records.Count > 0 Then blSelect = Not .Records(0).Item(mCol.ѡ��).Checked
                For Each Record In .Records
                    Record.Item(mCol.ѡ��).Checked = blSelect
                Next
            End If
        End If
        .Populate
    End With
End Sub

Private Sub rptMachine_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim hitColumn As ReportColumn
    Dim Record As ReportRecord
    Dim blSelect As Boolean

    With Me.rptMachine
        Set hitColumn = .HitTest(X, Y).Column
        If Not hitColumn Is Nothing Then
            If hitColumn.Caption = "ѡ��" And .HitTest(X, Y).ht = xtpHitTestHeader Then
                If .Records.Count > 0 Then blSelect = Not .Records(0).Item(mMCol.ѡ��).Checked
                For Each Record In .Records
                    Record.Item(mMCol.ѡ��).Checked = blSelect
                Next
            End If
        End If
        .Populate
    End With
End Sub

Private Sub txtSample_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call RefreshData: TxtSample.SelStart = 0: TxtSample.SelLength = Len(TxtSample)
End Sub

Private Sub TxtSample_Validate(Cancel As Boolean)
    '   10861 �������Ƿ�Ӧ����%,?���ַ�������
    If Check_Sample = False Then
        Cancel = True
    End If
End Sub

Private Function Check_Sample() As Boolean
    '   10861 �������Ƿ�Ӧ����%,?���ַ�������
    Dim i As Long, str�ַ� As String
    str�ַ� = ""
    If Len(TxtSample) > 0 Then
        For i = 1 To Len(TxtSample)
            If InStr("0123456789,~", Mid(TxtSample, i, 1)) <= 0 Then
                str�ַ� = str�ַ� & Mid(TxtSample, i, 1)
            End If
        Next
    End If

    If str�ַ� <> "" Then
        MsgBox "��������" & str�ַ�, vbQuestion, gstrSysName
        Check_Sample = False
    Else
        Check_Sample = True
    End If

End Function
Private Function ModifySampleNumber() As Boolean
    '����               �����޸ı걾��
    '����               intModifyNumber   ��ʼ�޸ĵı걾��
    Dim Record As ReportRecord
    Dim intLoop As Integer
    Dim astrSQL() As String
    Dim intCount As Integer
    Dim lngMachine As Long
    Dim rsTmp  As New ADODB.Recordset
    Dim blnUnChecked As Boolean, strMsg As String
    Dim strStartDate As String
    Dim strEndDate As String
    Dim blnBegin As Boolean
    
    On Error GoTo errH
    Me.MousePointer = 11
    zlCommFun.ShowFlash "����׼���޸ı걾����ȴ�..."
    
    If Trim(Me.TxtModify.Text) = "" Then
        strMsg = "�������޸ı걾�ŵĿ�ʼ����!"
        blnUnChecked = True
    ElseIf Not IsNumeric(Trim(Me.TxtModify.Text)) Then
        '11484 ��ʼ����Ϊ������ʱ������
        strMsg = "��ѿ�ʼ�����Ϊ����!"
        blnUnChecked = True
    End If
        
    If blnUnChecked Then
        MsgBox strMsg, vbQuestion, gstrSysName
        zlCommFun.StopFlash
        Me.MousePointer = 0
        Me.TxtModify.SetFocus
        Exit Function
    End If
    
    With Me.rptList
        For intLoop = 0 To .Records.Count - 1
            If .Records(intLoop).Item(mCol.ѡ��).Checked = True Then
                intCount = intCount + 1
                ReDim Preserve astrSQL(1 To intCount)
                astrSQL(intCount) = "ZL_����걾��¼_�걾���(" & .Records(intLoop).Item(mCol.�걾id).Value & ",'" & intCount + TxtModify.Text - 1 & _
                "',null,null,to_date('" & Now & "','yyyy-mm-dd hh24:mi:ss')," & "to_date('" & Now & "','yyyy-mm-dd hh24:mi:ss'))"
                lngMachine = Val(.Records(intLoop).Item(mCol.����id).Value)
                
                strStartDate = GetDateTime(mMakeNoRule, 1, .Records(intLoop).Item(mCol.����ʱ��).Value)
                strEndDate = GetDateTime(mMakeNoRule, 2, .Records(intLoop).Item(mCol.����ʱ��).Value)
                
                gstrSql = "Select Id From ����걾��¼ Where �걾��� = [1] " & IIf(lngMachine > 0, " And ����id = [2] ", "") & " And " & _
                          " ����ʱ�� Between [3] And [4] And ID <> [5] "
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, intCount + TxtModify.Text - 1, lngMachine, CDate(strStartDate), _
                            CDate(strEndDate), CStr(.Records(intLoop).Item(mCol.�걾id).Value))
                
                zlCommFun.ShowFlash "���ڸ��±걾" & .Records(intLoop).Item(mCol.�걾���).Value
                If rsTmp.EOF = False Then
                    zlCommFun.StopFlash
                    Me.MousePointer = 0
                    MsgBox "�걾��" & intCount + TxtModify.Text - 1 & "�ѱ�ʹ�ã�������������޸ı걾��!", vbInformation, Me.Caption
                    Exit Function
                End If
            End If
        Next
        If intCount = 0 Then
            zlCommFun.StopFlash
            Me.MousePointer = 0
            Exit Function
        End If
        gcnOracle.BeginTrans
        blnBegin = True
        For intLoop = 1 To UBound(astrSQL)
            zlDatabase.ExecuteProcedure astrSQL(intLoop), Me.Caption
        Next
        gcnOracle.CommitTrans
    End With
    zlCommFun.StopFlash
    Me.MousePointer = 0
    RefreshData
Exit Function
errH:
    If blnBegin = True Then
        gcnOracle.RollbackTrans
    End If
    RefreshData
    zlCommFun.StopFlash
    Me.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function


