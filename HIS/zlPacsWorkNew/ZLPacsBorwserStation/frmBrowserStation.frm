VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmBrowserStation 
   Caption         =   "Ӱ���������վ"
   ClientHeight    =   7305
   ClientLeft      =   10185
   ClientTop       =   345
   ClientWidth     =   11325
   Icon            =   "frmBrowserStation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   11325
   Begin ZLPacsBrowserStation.ucReadCard ucLocate 
      Height          =   330
      Left            =   6720
      TabIndex        =   16
      Top             =   1320
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   582
      Picture         =   "frmBrowserStation.frx":0E42
   End
   Begin VB.PictureBox PicWindow 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   1725
      ScaleHeight     =   4215
      ScaleWidth      =   9510
      TabIndex        =   1
      Top             =   2670
      Width           =   9510
      Begin VB.PictureBox picInfo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   120
         ScaleHeight     =   1575
         ScaleWidth      =   9465
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   120
         Width           =   9465
         Begin VB.Frame fraRegist 
            Height          =   810
            Left            =   0
            TabIndex        =   7
            Top             =   -75
            Width           =   8700
            Begin VB.CommandButton cmdReportView 
               Appearance      =   0  'Flat
               Height          =   615
               Left            =   8040
               Picture         =   "frmBrowserStation.frx":1194
               Style           =   1  'Graphical
               TabIndex        =   15
               ToolTipText     =   "����鿴"
               Top             =   120
               Width           =   615
            End
            Begin VB.ComboBox cboTimes 
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   260
               Width           =   6315
            End
            Begin VB.Label lblRegist 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����¼(&G)��"
               Height          =   180
               Left            =   105
               TabIndex        =   9
               Top             =   320
               Width           =   1170
            End
         End
         Begin VB.Frame fraInfo 
            Height          =   700
            Left            =   0
            TabIndex        =   4
            Top             =   600
            Width           =   7410
            Begin VB.Label lblCash 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "��"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   21.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   540
               Left            =   6840
               TabIndex        =   10
               Top             =   120
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.Label lbl�����Ϣ 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�����Ϣ"
               ForeColor       =   &H00C00000&
               Height          =   180
               Left            =   90
               TabIndex        =   6
               Top             =   400
               Width           =   720
            End
            Begin VB.Label lbl������Ϣ 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������Ϣ"
               ForeColor       =   &H00C00000&
               Height          =   180
               Left            =   90
               TabIndex        =   5
               Top             =   150
               Width           =   720
            End
         End
      End
      Begin XtremeSuiteControls.TabControl TabWindow 
         Height          =   2415
         Left            =   240
         TabIndex        =   2
         Top             =   1800
         Width           =   4125
         _Version        =   589884
         _ExtentX        =   7276
         _ExtentY        =   4260
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4275
      Left            =   30
      ScaleHeight     =   4275
      ScaleWidth      =   4500
      TabIndex        =   11
      Top             =   495
      Width           =   4495
      Begin VB.TextBox txtAppend 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDD6C6&
         BorderStyle     =   0  'None
         Height          =   2100
         Left            =   630
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   1605
         Width           =   2010
      End
      Begin VSFlex8Ctl.VSFlexGrid vsList 
         Height          =   2685
         Left            =   450
         TabIndex        =   13
         Top             =   435
         Width           =   3360
         _cx             =   5927
         _cy             =   4736
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   7
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin VB.CommandButton cmdInfo 
            Caption         =   "��"
            Height          =   240
            Left            =   2730
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "ѡ����Ŀ(*)"
            Top             =   270
            Visible         =   0   'False
            Width           =   270
         End
      End
      Begin ZLPacsBrowserStation.ucReadCard ucFilter 
         Height          =   330
         Left            =   360
         TabIndex        =   17
         Top             =   120
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   582
         Picture         =   "frmBrowserStation.frx":1DD6
      End
      Begin XtremeCommandBars.CommandBars cbrdock 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.Timer TimerRefresh 
      Enabled         =   0   'False
      Left            =   7200
      Top             =   600
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Bindings        =   "frmBrowserStation.frx":2128
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6945
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBrowserStation.frx":213C
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7805
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.ImageList Imglist 
      Left            =   6570
      Top             =   75
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowserStation.frx":29D0
            Key             =   "����"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowserStation.frx":2F6A
            Key             =   "סԺ"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowserStation.frx":3844
            Key             =   "����"
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowserStation.frx":399E
            Key             =   "Ӱ��"
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowserStation.frx":4118
            Key             =   "�ѽ�"
            Object.Tag             =   "5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowserStation.frx":44B2
            Key             =   "��ɫͨ��"
            Object.Tag             =   "6"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   5940
      Top             =   75
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowserStation.frx":460C
            Key             =   "��ѡ����"
            Object.Tag             =   "90000"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowserStation.frx":4BA6
            Key             =   "��ѡѡ��"
            Object.Tag             =   "90001"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   840
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmBrowserStation.frx":5140
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmBrowserStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mintCurҵ������ As Integer = 1 '��ǰϵͳ������ҵ������

Private Const ConstrCol = "����;300|��Դ;400|����;300|����;300|����;1200|����;1400|������;800|�Ա�;450|����;450" & _
                        "|��ʶ��;1400|ҽ������;2400|��λ����;1400|ִ�м�;600|����ʱ��;1800|����ʱ��;1800|����ҽ��;800" & _
                        "|���;450|����;450|Ӥ��;450|�Ǽ���;800|������;800|�����;800|��ӡ��Ƭ;800|�������;800" & _
                        "|��ɫͨ��;0|�����ӡ;800|������;800|������;800|��鼼ʦ;800|��ͼʱ��;1800|�������;2400" & _
                        "|Ӱ�����;0|����ID;0|��ҳID;0|�Һŵ�;0|���˿���ID;0|ҽ��ID;1200|���ͺ�;0|���UID;0" & _
                        "|���״̬;0|NO;0|��¼����;0|ת��;0|����;0|��ǰ����ID;0|���淢��;800|��Ϸ���;800" & _
                        "|ִ�п���ID;0|����ID;0|���˿���;800|���￨��;800|���ݺ�;800|���֤��;800"
Private mstrCol As String   '�б�˳�������ʱ��ȡע�������ֵ��ConstrColΪĬ��ֵ

'ID_���ҷ�ʽ+100֮����7������Ϊ���ҷ�ʽѡ���
'ID_Ӱ�����֮����40��������ΪӰ����𣬴�4021-4060
Private Enum FilterID
    ID_���� = 4001: ID_סԺ = 4002: ID_��� = 4003: ID_���� = 4004
    ID_���� = 4005: ID_�ѽ� = 4006: ID_δ�� = 4007: ID_�Ǽ� = 4008
    ID_���� = 4009: ID_��� = 4010: ID_���� = 4011: ID_��� = 4012: ID_��� = 4013
    ID_���ҷ�ʽ = 4014: ID_����ֵ = 4015: ID_��ʼ���� = 4016: ID_����סԺ = 4017
    ID_Ӱ����� = 4020
End Enum

Private mblncmd���� As Boolean, mblncmdסԺ As Boolean, mblncmd��� As Boolean, mblncmd���� As Boolean, mblncmd�ѽ� As Boolean, mblncmdδ�� As Boolean
Private mblncmd�Ǽ� As Boolean, mblncmd���� As Boolean, mblncmd��� As Boolean, mblncmd���� As Boolean, mblncmd��� As Boolean, mblncmd��� As Boolean
Private mblncmd���� As Boolean
Private mintcmdӰ����� As Integer      '0��ʾû��ѡ��Ӱ������������ֱ�ʾѡ���Ӱ����������
Private mblncmdӰ�����() As Boolean    '���浱ǰѡ���Ӱ������Ƿ�ѡ��



Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mobjICCard As Object
Private Enum IDKinds
    C0��������￨ = 0
    C1ҽ���� = 1
    C2���֤�� = 2
    C3IC���� = 3
End Enum

'�Ӵ������
Private mfrmPACSImg As frmPACSImg       'Ӱ���Ӵ���
Public mobjRichEPR As New zlRichEPR.cRichEPR
Private WithEvents mobjPacsCore As zl9PacsCore.clsViewer           '��Ƭվ����
Attribute mobjPacsCore.VB_VarHelpID = -1

'���ڱ���
Private mlngAdviceID As Long
'Private mlngCur����ID As Long                               '��ǰ����ID
Private mstrҽ���������� As String                            '��ǰ���ÿ���  ID_����-����
Private mstrҽ����������IDs As String
'Private mstr��ǰҽ������ As String                            '��ǰ���� ����-����
Private mblnInitOk As Boolean, mblnvsRefresh As Boolean     '��ʼ�����,װ�ر��
Private mstrPrivs As String, mlngModul As Long              'ģ��ţ���ģ��Ȩ��
Private mblnAllDepts As Boolean                             '�Ƿ�ѡ��ȫ������
Private mlngSortCol As Long                                 '�����б��У���ǰ�����������
Private mintSortOrder As Integer                            '�����б��У���ǰ��������ķ�ʽ

'���̿��Ʊ���
Private mblnShowImgAtReport As Boolean                      '�򿪱���ʱ�򿪹�Ƭվ
Private mBeforeDays As Integer                              'Ĭ�ϲ�ѯ������
Private mlngRefreshInterval As Long                         '�����б��Զ�ˢ�¼��
Private mblnRelatingPatient As Boolean                      '�Ƿ����ù�������
Private mblnMoved As Boolean                                '��ǰʱ������Ƿ�ת�ƹ�

Private mblnUse3D As Boolean                                '�Ƿ�������ά�ؽ�����
Private mstr3DExeDir As String                              '��ά�ؽ�����·��
Private mstr3DPara As String                                '��ά�ؽ�����
Private mstr3DFunctions As String                           '��ά�ؽ�����

'������������
Private Type Type_SQLCondition
    ��ʼʱ�� As Date
    ����ʱ�� As Date
    ʱ������ As Integer                                 'ʱ���ѯ��ʽ 1=������ʱ�䣨����ҽ������.����ʱ�䣩��2=������ʱ�䣨����ҽ������.�״�ʱ�䣩��3=��ͼʱ�䣨Ӱ�����¼.�������ڣ�
    ���ݺ� As String
    ����� As Double
    סԺ�� As Double
    ���￨ As String
    ���� As String
    �Ա� As String
    ��ʼ���� As Long
    �������� As Long
    �������� As String
    ���� As Double
    ���֤  As String
    IC�� As String
    ���˿��� As Long
    �걾��λ As String
    ���ҽ�� As String
    ���ҽ�� As String
    ������� As String
    �������� As String
    ������� As Integer
    Ӱ������ As String
    ��鼼ʦ As String
    ������ As String
    Ӱ����� As String
    ������� As String
    ������ As String
    ���� As String
    ��� As String
    ����ID As Long
End Type
Private SQLCondition As Type_SQLCondition

Private mlngHSendNo As Long
Private mstrHStudyUID As String
Private mlngExecuteStep As Long '���ִ�й���
Private mblnHMoved As Boolean


Private Sub OpenReportPreview(ByVal lngAdviceID As Long)
    If mobjRichEPR Is Nothing Then Exit Sub
    
    On Error GoTo errHandle
        
        Dim strSQL As String
        Dim lngExecuteStep As Long
        Dim rsReport As ADODB.Recordset
        Dim blnCanPrint As Boolean
        
        strSQL = "select ִ�й���,����ID from ����ҽ������ A,����ҽ������ R where R.ҽ��ID=A.ҽ��ID and  A.ҽ��ID=[1]"
        Set rsReport = zlDatabase.OpenSQLRecord(strSQL, "��ѯ����ID", lngAdviceID)
        
        If rsReport.EOF Then
            MsgBoxD Me, "û���ҵ���ǰ����Ӧ�Ĳ�����Ϣ�����飡", vbInformation, gstrSysName
            Exit Sub
        End If
        
        lngExecuteStep = rsReport!ִ�й���
        
        'If lngExecuteStep <> 5 And rsReport!ִ�й��� <> 6 Then
        If rsReport!ִ�й��� <> 6 Then
            MsgBoxD Me, "������δ��ɣ����ܽ��в鿴��", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If InStr(mstrPrivs, "PACS�����ӡ") > 0 Then
            blnCanPrint = True
        Else
            blnCanPrint = False
        End If
        
        Call mobjRichEPR.ViewDocument(Me, rsReport!����Id, blnCanPrint)
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume Next
    End If
End Sub


Private Sub Menu_Help_About_click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Menu_Help_Help_click()
    '���ܣ����ð�������
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub Menu_Help_Web_Forum_click()
    Call zlWebForum(Me.Hwnd)
End Sub


Private Sub Menu_Help_Web_Mail_click()
    zlMailTo Hwnd
End Sub


Private Sub Menu_Manage_��Ƭ()
    If TabWindow.Selected.Tag <> "Ӱ��ͼ��" Then '��ˢ��ͼ������
        Call mfrmPACSImg.zlRefresh(mlngAdviceID, mlngHSendNo, mstrPrivs, mblnHMoved)
    End If
    
    Call mfrmPACSImg.zlMenuClick("Ӱ����")
End Sub


Private Sub Menu_Manage_�Աȹ�Ƭ()
    If TabWindow.Selected.Tag <> "Ӱ��ͼ��" Then '��ˢ��ͼ������
        Call mfrmPACSImg.zlRefresh(mlngAdviceID, mlngHSendNo, mstrPrivs, mblnHMoved)
    End If
    
    Call mfrmPACSImg.zlMenuClick("Ӱ��Ա�")
End Sub


Private Sub Menu_View_StatusBar_click(ByVal control As XtremeCommandBars.ICommandBarControl)
    Me.stbThis.Visible = Not Me.stbThis.Visible
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_View_ToolBar_Button_click(ByVal control As XtremeCommandBars.ICommandBarControl)
Dim i As Integer
    For i = 2 To cbrMain.Count
        Me.cbrMain(i).Visible = Not Me.cbrMain(i).Visible
    Next

    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_View_ToolBar_Size_click(ByVal control As XtremeCommandBars.ICommandBarControl)
    Me.cbrMain.Options.LargeIcons = Not Me.cbrMain.Options.LargeIcons
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_View_ToolBar_Text_click(ByVal control As XtremeCommandBars.ICommandBarControl)
Dim i As Integer, cbrControl As CommandBarControl
    For i = 2 To cbrMain.Count
        For Each cbrControl In Me.cbrMain(i).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
    Next
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub cboTimes_Click()
    If cboTimes.ListCount <= 1 Then Exit Sub
    If cboTimes.Tag = "" Then Exit Sub '��ʱcbotime��Ŀδ������ɣ���listindex��ֵ����
    
    On Error GoTo errHandle
    
    mlngAdviceID = cboTimes.ItemData(cboTimes.ListIndex)
    'If mlngAdviceID = vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")) Then Call vsList_RowColChange: Exit Sub '�����뵱ǰѡ��ҽ��ID��ͬʱ���ɱ���������
    
    If mlngAdviceID = vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")) Then
      Call FillTxtInfor  '������Ϸ����˻�����Ϣ
      Call FillTxtAppend '������½�ҽ������
    
      Call RefreshTabWindow 'ˢ���Ӵ���
        
    Else
      '�����������̵������Ⱥ�˳�������
      Call FillTxtInfor(mlngAdviceID)  '������Ϸ����˻�����Ϣ
      Call FillTxtAppend(mlngAdviceID) '������½�ҽ������
    
      Call RefreshTabWindow(mlngAdviceID) 'ˢ���Ӵ���
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cboTimes_DropDown()
    Call SendMessage(cboTimes.Hwnd, &H160, 500, 0)
End Sub

Private Sub cbrdock_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Dim objControl As CommandBarControl
    Dim i As Integer
    Dim strTemp As String
    Dim strCardName As String
    Dim strCardText As String
    Dim lngPatientID As Long
    
    Select Case control.ID
        Case ID_����
            mblncmd���� = Not control.Checked
        Case ID_סԺ
            mblncmdסԺ = Not control.Checked
        Case ID_����
            mblncmd���� = Not control.Checked
        Case ID_���
            mblncmd��� = Not control.Checked
        Case ID_�ѽ�
            mblncmd�ѽ� = Not control.Checked
            If mblncmd�ѽ� Then mblncmdδ�� = False
        Case ID_δ��
            mblncmdδ�� = Not control.Checked
            If mblncmdδ�� Then mblncmd�ѽ� = False
        Case ID_Ӱ����� + 1 To ID_Ӱ����� + 40
            control.Checked = Not control.Checked
            mblncmdӰ�����(control.ID - ID_Ӱ����� - 1) = control.Checked
            If control.Checked = True Then
                mintcmdӰ����� = mintcmdӰ����� + 1
            Else
                mintcmdӰ����� = mintcmdӰ����� - 1
            End If
            Set objControl = cbrdock.FindControl(, ID_Ӱ�����)
            If mintcmdӰ����� = 0 Then
                strTemp = "Ӱ�����"
            Else
                strTemp = ""
                For i = 1 To objControl.CommandBar.Controls.Count
                    If objControl.CommandBar.FindControl(, ID_Ӱ����� + i).Checked = True Then
                        strTemp = IIf(strTemp = "", objControl.CommandBar.FindControl(, ID_Ӱ����� + i).Caption, strTemp & "," & objControl.CommandBar.FindControl(, ID_Ӱ����� + i).Caption)
                    End If
                Next i
            End If
            objControl.Caption = strTemp
        Case ID_�Ǽ�
            mblncmd�Ǽ� = Not control.Checked
        Case ID_����
            mblncmd���� = Not control.Checked
        Case ID_���
            mblncmd��� = Not control.Checked
        Case ID_����
            mblncmd���� = Not control.Checked
        Case ID_���
            mblncmd��� = Not control.Checked
        Case ID_���
            mblncmd��� = Not control.Checked
        Case ID_����סԺ
            control.Checked = Not control.Checked
            mblncmd���� = Not mblncmd����
        Case ID_��ʼ����
            Call ucFilter.GetCardValue(strCardName, strCardText, lngPatientID)
            Call subRefreshFilterCondition(strCardName, strCardText, lngPatientID)
    End Select
    
    cbrdock.RecalcLayout
    Call RefreshList
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub subRefreshFilterCondition(ByVal strCardName As String, ByVal strCardText As String, ByVal lngPatientID As Long)
'------------------------------------------------
'���ܣ���txtFilter�ؼ������ݸ��¹�������
'������ strFilter --- ��������
'���أ���
'------------------------------------------------

On Error GoTo err
    Dim strFilter As String
    
    strFilter = strCardText
    
    With SQLCondition
        .���� = ""
        .���￨ = ""
        .����� = 0
        .סԺ�� = 0
        .���ݺ� = ""
        .���� = 0
        .���֤ = ""
        .IC�� = ""
        .����ID = 0
        
        Select Case strCardName
            Case "����", "��  ��", "��   ��" '��������ǰ��ʽ����
                .���� = Trim(strFilter)
                
            Case "���￨"
                .���￨ = Trim(strFilter)
                
            Case "�����"   '��ݷ�ʽ�ǡ�*+���֡�,VAL��ȡǰ����*��Ҫ���⴦��
                If Left(strCardText, 1) = "*" Then
                    strFilter = Mid(strFilter, 2)
                End If
                .����� = Val(strFilter)
                
            Case "סԺ��"   '��ݷ�ʽ�ǡ�++���֡�
                .סԺ�� = Val(strFilter)
                
            Case "���ݺ�"
                .���ݺ� = Trim(strFilter)
                
            Case "����"
                .���� = Val(strFilter)
                
            Case "���֤��", "���֤"
                .���֤ = Trim(strFilter)
                
            Case "IC����", "IC��"
                .IC�� = Trim(strFilter)
                
            Case Else
                .����ID = lngPatientID
        End Select
    End With
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbrdock_Resize()
Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    On Error Resume Next
    Call Me.cbrdock.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    vsList.Top = lngTop: vsList.Left = lngLeft
    vsList.Width = picList.Width
    vsList.Height = picList.Height - lngTop - txtAppend.Height - 100


    txtAppend.Top = vsList.Top + vsList.Height + 100: txtAppend.Left = lngLeft + 100
    txtAppend.Width = picList.Width - 200
End Sub

Private Sub cbrdock_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    Select Case control.ID
        Case ID_����
            control.Checked = mblncmd����
            control.IconId = IIf(mblncmd����, 90001, 90000)
        Case ID_סԺ
            control.Checked = mblncmdסԺ
            control.IconId = IIf(mblncmdסԺ, 90001, 90000)
        Case ID_����
            control.Checked = mblncmd����
            control.IconId = IIf(mblncmd����, 90001, 90000)
        Case ID_���
            control.Checked = mblncmd���
            control.IconId = IIf(mblncmd���, 90001, 90000)
        Case ID_����
            control.Checked = mblncmd�ѽ� Xor mblncmdδ��
            control.Caption = IIf(mblncmd�ѽ� Xor mblncmdδ��, IIf(mblncmd�ѽ�, " �ѽɷ�", " δ�ɷ�"), " ��  ��")
        Case ID_�ѽ�
            control.Checked = mblncmd�ѽ�
            control.IconId = IIf(mblncmd�ѽ�, 90001, 90000)
        Case ID_δ��
            control.Checked = mblncmdδ��
            control.IconId = IIf(mblncmdδ��, 90001, 90000)
        Case ID_Ӱ�����
            control.IconId = IIf(mintcmdӰ����� = 0, 90000, 90001)
        Case ID_Ӱ����� + 1 To ID_Ӱ����� + 40
            control.Checked = mblncmdӰ�����(control.ID - ID_Ӱ����� - 1)
            control.IconId = IIf(control.Checked, 90001, 90000)
        Case ID_�Ǽ�
            control.Checked = mblncmd�Ǽ�
            control.IconId = IIf(mblncmd�Ǽ�, 90001, 90000)
        Case ID_����
            control.Checked = mblncmd����
            control.IconId = IIf(mblncmd����, 90001, 90000)
        Case ID_���
            control.Checked = mblncmd���
            control.IconId = IIf(mblncmd���, 90001, 90000)
        Case ID_����
            control.Checked = mblncmd����
            control.IconId = IIf(mblncmd����, 90001, 90000)
        Case ID_���
            control.Checked = mblncmd���
            control.IconId = IIf(mblncmd���, 90001, 90000)
        Case ID_���
            control.Checked = mblncmd���
            control.IconId = IIf(mblncmd���, 90001, 90000)
        Case ID_����סԺ
            control.IconId = IIf(control.Checked, 90001, 90000)
    End Select
End Sub
Private Sub cbrMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = stbThis.Height
End Sub



Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    
    If control.ID <> 0 Then
        If cbrMain.FindControl(, control.ID, True, True) Is Nothing Then Exit Sub
    End If
    
    cbrMain.RecalcLayout
    Select Case control.ID
    
'--------------------------�ļ�------------------
        
        Case conMenu_Manage_Change_In   '�����б�
            If dkpMain.Panes(1).Hidden = False Then
                dkpMain.Panes(1).Hide
            Else
                dkpMain.ShowPane (1)
            End If

        Case conMenu_File_Exit '�˳�
            Unload Me
            
'---------------------------���-----------------
        Case conMenu_Manage_RequestPrint * 10# + 1 To conMenu_Manage_RequestPrint * 10# + 9 '��ӡ���Ƶ���
            Call FuncBillPrint(control)
'
        Case conMenu_Img_Look                         '��Ƭ
            Call Menu_Manage_��Ƭ
        
        Case conMenu_Img_Contrast                     '�Աȹ�Ƭ
            Call Menu_Manage_�Աȹ�Ƭ
        
        Case conMenu_Img_3D_MMPR                    '��ά�ؽ���MMPR
            Call sub��ά�ؽ�("MMPR")
        Case conMenu_Img_3D_MPR                     '��ά�ؽ���MPR
            Call sub��ά�ؽ�("MPR")
        Case conMenu_Img_3D_PF                     '��ά�ؽ�,��ע����
            Call sub��ά�ؽ�("PF")
        Case conMenu_Img_3D_SA                     '��ά�ؽ��������ؽ�
            Call sub��ά�ؽ�("SA")
        Case conMenu_Img_3D_VA                     '��ά�ؽ����ݻ��ؽ�
            Call sub��ά�ؽ�("VA")
        Case conMenu_Img_3D_VE                     '��ά�ؽ��������ڿ���
            Call sub��ά�ؽ�("VE")

        Case conMenu_File_Preview ', conMenu_File_Print       '����Ԥ���ʹ�ӡ
            Dim i As Integer
            'û���治�ܴ�ӡ��Ԥ��
            If vsList.TextMatrix(vsList.Row, GetCN("������")) = "" Then
                MsgBoxD Me, "��ǰ����û�м�鱨�棬���ܲ��������飡", vbInformation, gstrSysName
                Exit Sub
            End If
            
            Call OpenReportPreview(mlngAdviceID)

'---------------------------�鿴----------------
        Case conMenu_View_ToolBar_Button                        '������
            Call Menu_View_ToolBar_Button_click(control)
        Case conMenu_View_ToolBar_Text                          '��ť����
            Call Menu_View_ToolBar_Text_click(control)
        Case conMenu_View_ToolBar_Size                          '��ͼ��
            Call Menu_View_ToolBar_Size_click(control)
        Case conMenu_View_StatusBar                             '״̬��
            Call Menu_View_StatusBar_click(control)
        Case conMenu_View_Filter                                '����
            Call Menu_View_Filter_click
        Case conMenu_View_Refresh                               'ˢ��
            Call RefreshList
        Case conMenu_Help_Help                                  '����
            Call Menu_Help_Help_click
        Case conMenu_Help_Web_Forum                             '��������
'            Case zlWebForum(Me.Hwnd)
        Case conMenu_Help_Web_Home                              '��������
            Call zlHomePage(Me.Hwnd)
        Case conMenu_Help_Web_Mail                              '��������
            Call zlMailTo(Me.Hwnd)
        Case conMenu_Help_About                                 '����
            Call Menu_Help_About_click
        Case conMenu_View_Filter * 100# To conMenu_View_Filter * 100# + UBound(Split(mstrҽ����������, "|")) + 1 '���ĵ�ǰ����
            Call Menu_Dept_Select(control)
    End Select
End Sub


Private Sub Menu_Dept_Select(ByVal control As XtremeCommandBars.ICommandBarControl)
    '�������ң������µ����������¹��˲���
    '���ѡ�����ȫ�����ң��� mlngCur����ID ���ı�
    '���ѡ�����ĳ��������ң���ı� mlngCur����ID
    If glngDeptId <> control.DescriptionText Or (control.DescriptionText <> 0 And mblnAllDepts = True) Then
        'ѡ���˾�����ң��Ÿı䵱ǰ���ҵ�����
        If control.DescriptionText = 0 Then
            mblnAllDepts = True
        Else
            mblnAllDepts = False
            glngDeptId = control.DescriptionText
            gstrDeptName = Split(control.Caption, "(")(0)
            
        End If
        
        Call cbrMain.RecalcLayout
        Call RefreshList
    End If
End Sub


Private Sub Menu_View_Filter_click()
    On Error GoTo errHandle
    
    With frmPACSFilter
        .mlngModul = mlngModul
        .mBeforeDays = mBeforeDays
'        .mDept = mlngCur����ID '��ǰ����
        .Show 1, Me
        If Not .mblnOK Then Exit Sub 'û�з�������
        
        '��ʹ��ʱ������ʱ����չ̶�����
        ucFilter.CardText = ""
        SQLCondition.���� = ""
        SQLCondition.���￨ = ""
        SQLCondition.����� = 0
        SQLCondition.סԺ�� = 0
        SQLCondition.���ݺ� = ""
        SQLCondition.���� = 0
        SQLCondition.���֤ = ""
        SQLCondition.IC�� = ""
        
        SQLCondition.��ʼʱ�� = Format(.dtpBegin.value, "yyyy-MM-dd HH:mm:00")
        SQLCondition.����ʱ�� = Format(.dtpEnd.value, "yyyy-MM-dd HH:mm:59")
'        If Format(.dtpEnd.value, "yyyy-MM-dd HH:mm") = Format(.dtpEnd.Tag, "yyyy-MM-dd HH:mm") Then
'            SQLCondition.����ʱ�� = CDate(0) '��ʾȡ��ǰʱ��
'        Else
'            SQLCondition.����ʱ�� = Format(.dtpEnd.value, "yyyy-MM-dd HH:mm:59")
'        End If
        
        mblnMoved = MovedByDate(SQLCondition.��ʼʱ��)
        
        If .optFindType(1).value = True Then 'ʱ���ѯ��ʽ 1=������ʱ�䣨����ҽ������.����ʱ�䣩��2=������ʱ�䣨����ҽ������.�״�ʱ�䣩��3=��ͼʱ�䣨Ӱ�����¼.�������ڣ�
            SQLCondition.ʱ������ = 1
        ElseIf .optFindType(2).value = True Then
            SQLCondition.ʱ������ = 2
        Else
            SQLCondition.ʱ������ = 3
        End If
        
        If .cboPart.ListIndex <> 0 Then '���걾��λ
            SQLCondition.�걾��λ = NeedName(.cboPart.Text)
        Else
            SQLCondition.�걾��λ = ""
        End If
        
        '�����Ա�
        If NeedName(.cboSex.Text) = "ȫ��" Then
            SQLCondition.�Ա� = ""
        Else
            SQLCondition.�Ա� = NeedName(.cboSex.Text)
        End If
        
        '��������
        Select Case NeedName(.cboAgeType.Text)
            Case "��"
                SQLCondition.��ʼ���� = Val(.txtBeginAge.Text) * 365
                SQLCondition.�������� = Val(.txtEndAge.Text) * 365
            Case "��"
                SQLCondition.��ʼ���� = Val(.txtBeginAge.Text) * 30
                SQLCondition.�������� = Val(.txtEndAge.Text) * 30
            Case "��"
                SQLCondition.��ʼ���� = Val(.txtBeginAge.Text) * 7
                SQLCondition.�������� = Val(.txtEndAge.Text) * 7
            Case "��"
                SQLCondition.��ʼ���� = Val(.txtBeginAge.Text) * 1
                SQLCondition.�������� = Val(.txtEndAge.Text) * 1
        End Select
        
        If Trim(.txtBeginAge.Text) = "" Then SQLCondition.��ʼ���� = -1
        If Trim(.txtEndAge.Text) = "" Then SQLCondition.�������� = -1
        
        SQLCondition.�������� = Trim(.cboAgeWhere.Text)
        
        If .cboDept.ListIndex <> 0 Then '���˿���
            SQLCondition.���˿��� = .cboDept.ItemData(.cboDept.ListIndex)
        Else
            SQLCondition.���˿��� = 0
        End If

        If .cbodiagdoc.ListIndex <> 0 Then '���ҽ��
            SQLCondition.���ҽ�� = NeedName(.cbodiagdoc.Text)
        Else
            SQLCondition.���ҽ�� = ""
        End If
        
        If .cboAuditing.ListIndex <> 0 Then '���ҽ��
            SQLCondition.���ҽ�� = NeedName(.cboAuditing.Text)
        Else
            SQLCondition.���ҽ�� = ""
        End If
        
        
'        If .cboCheckStep.ListIndex <> 0 Then '������
'            SQLCondition.������ = .cboCheckStep.Text
'        Else
'            SQLCondition.������ = ""
'        End If
        
        
        If .cboModality.ListIndex <> 0 Then 'Ӱ�����
            SQLCondition.Ӱ����� = Split(.cboModality.Text, "--")(1)
        Else
            SQLCondition.Ӱ����� = ""
        End If
        
        
        If Trim(.TxtӰ�����) <> "" Then 'Ӱ�����
            SQLCondition.������� = Trim(.TxtӰ�����)
        Else
            SQLCondition.������� = ""
        End If
        
        If Trim(.txt��������) <> "" Then '��������
            SQLCondition.�������� = Trim(.txt��������)
        Else
            SQLCondition.�������� = ""
        End If
        
        If NeedName(.cboYinYangXing.Text) = "����" Then
            SQLCondition.������� = 1
        ElseIf NeedName(.cboYinYangXing.Text) = "����" Then
            SQLCondition.������� = 0
        Else
            SQLCondition.������� = -1
        End If
        
        If .cbo����.ListIndex = 0 Then
            SQLCondition.Ӱ������ = ""
        Else
            SQLCondition.Ӱ������ = NeedName(.cbo����.Text)
        End If
        
        If .cbo��鼼ʦ.ListIndex = 0 Then
            SQLCondition.��鼼ʦ = ""
        Else
            SQLCondition.��鼼ʦ = NeedName(.cbo��鼼ʦ.Text)
        End If
        
        
        If Trim(.txtPacsRpt(0)) <> "" Then 'PACS�������
            SQLCondition.������� = Trim(.txtPacsRpt(0))
        Else
            SQLCondition.������� = ""
        End If
        
        If Trim(.txtPacsRpt(1)) <> "" Then
            SQLCondition.������ = Trim(.txtPacsRpt(1))
        Else
            SQLCondition.������ = ""
        End If
        
        If Trim(.txtPacsRpt(2)) <> "" Then
            SQLCondition.���� = Trim(.txtPacsRpt(2))
        Else
            SQLCondition.���� = ""
        End If
        
        If Trim(.txt���.Text) <> "" Then
            SQLCondition.��� = Trim(.txt���.Text)
        Else
            SQLCondition.��� = ""
        End If
        
        Call RefreshList '����ˢ��
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub cbrMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim objControl As CommandBarControl, i As Integer
    
    If CommandBar.Parent Is Nothing Then Exit Sub
    Select Case CommandBar.Parent.ID
        Case conMenu_View_Filter * 10#
            With CommandBar.Controls
                If .Count = 0 Then
                    '�����ȫ������
                    Set objControl = .Add(xtpControlButton, conMenu_View_Filter * 100#, "ȫ������")
                    objControl.BeginGroup = True
                    objControl.Category = "Main"
                    objControl.DescriptionText = 0
                    If mblnAllDepts = True Then objControl.Checked = True
                    
                    '�����ÿһ���������
                    For i = 0 To UBound(Split(mstrҽ����������, "|"))  'mstrҽ����������=id_����-����|id_����-����
                        Set objControl = .Add(xtpControlButton, conMenu_View_Filter * 100# + i + 1, Split(Split(mstrҽ����������, "|")(i), "_")(1) & "(&" & i & ")")
                        objControl.Category = "Main"
                        objControl.DescriptionText = Split(Split(mstrҽ����������, "|")(i), "_")(0)
                        If mblnAllDepts = False And glngDeptId = objControl.DescriptionText Then objControl.Checked = True
                    Next
                End If
            End With
    End Select
End Sub


Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim blnNoRecord As Boolean, intState As Integer, blnCancel As Boolean
    If Not mblnInitOk Then Exit Sub

    blnNoRecord = Val(vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID"))) = 0
    control.Style = xtpButtonIconAndCaption
    If Not blnNoRecord Then
        intState = Val(vsList.TextMatrix(vsList.Row, GetCN("���״̬")))
        blnCancel = vsList.TextMatrix(vsList.Row, GetCN("������")) = "�Ѿܾ�"
    End If

    Select Case control.ID
        Case conMenu_Manage_LocateValue
            control.Enabled = Not blnNoRecord
        Case conMenu_View_Filter * 10#
            control.Caption = "��ǰ����:" & IIf(mblnAllDepts = True, "ȫ������", gstrDeptName)
        Case conMenu_View_Filter * 100# To conMenu_View_Filter * 100# + UBound(Split(mstrҽ����������, "|")) + 1
            If mblnAllDepts = True Then
                control.Checked = (control.DescriptionText = 0)
            Else
                control.Checked = (control.DescriptionText = glngDeptId)
            End If
        Case conMenu_View_ToolBar_Button '������
            If cbrMain.Count >= 2 Then
                control.Checked = Me.cbrMain(2).Visible
            End If
        Case conMenu_View_ToolBar_Text 'ͼ������
            If cbrMain.Count >= 2 Then
                control.Checked = Not (Me.cbrMain(2).Controls(1).Style = xtpButtonIcon)
            End If
        Case conMenu_View_ToolBar_Size '��ͼ��
            control.Checked = Me.cbrMain.Options.LargeIcons
        Case conMenu_View_StatusBar '״̬��
            control.Checked = Me.stbThis.Visible
        Case conMenu_View_Filter   '����

        Case conMenu_View_Refresh  'ˢ��

        Case conMenu_Manage_RequestPrint
            control.Enabled = control.CommandBar.Controls.Count > 0 And Not blnNoRecord

        Case conMenu_Img_Contrast, conMenu_Img_Look     'Ӱ��Ա�,Ӱ���Ƭ
            If blnNoRecord Then control.Enabled = False: Exit Sub

            control.Enabled = mstrHStudyUID <> ""
                        
            'If control.Parent.Type <> xtpControlPopup Then control.Visible = control.Enabled
        Case conMenu_Img_3D     '��ά�ؽ�
            If InStr(mstrPrivs, "��ά�ؽ�����") <> 0 And mblnUse3D = True Then
                control.Visible = True
            Else
                control.Visible = False
            End If
            If control.Visible = True Then
                If blnNoRecord Then control.Enabled = False: Exit Sub
                If control.Parent.Type <> xtpControlPopup Then
                    control.Visible = vsList.TextMatrix(vsList.Row, GetCN("���UID")) <> ""
                    control.Enabled = control.Visible
                Else
                    control.Enabled = vsList.TextMatrix(vsList.Row, GetCN("���UID")) <> ""
                End If
            End If

'        Case conMenu_File_PrintSet     '��ӡ����(&S)
        Case conMenu_File_Preview, conMenu_File_Print '����Ԥ��(&V) �����ӡ(&P)
            control.Enabled = Not blnNoRecord And (mlngExecuteStep = 5 Or mlngExecuteStep = 6)
'        Case conMenu_File_Excel         '�嵥��ӡ(&L)
'            control.Enabled = Not blnNoRecord
        Case conMenu_ReportPopup, conMenu_ReportPopup * 100# + 1 To conMenu_ReportPopup * 100# + 99 '����
        Case conMenu_FilePopup, conMenu_ManagePopup, conMenu_ViewPopup, conMenu_HelpPopup
        Case conMenu_Help_Help, conMenu_Help_About  '����
        Case conMenu_Help_Web, conMenu_Help_Web_Forum, conMenu_Help_Web_Home, conMenu_Help_Web_Mail '����WEB
        Case conMenu_File_Exit      '�˳�
        Case conMenu_View_ToolBar   '������
        Case conMenu_Cap_DevSet     'Ӱ���豸����
        Case conMenu_Manage_Change_In   '�����б�
    End Select
End Sub

Private Sub InitMvar(intType As Integer)
'����:��ʼ��ģ�鼶����,������ء����Ҹı�ʱ����
'������intType---0�Ӳ˵�����FormLoad�������Ҹı䣬ˢ�²��˹��˿�ʼʱ�䣻intType---1�Ӳ����б������Ҹı䣬������ˢ�¹��˿�ʼʱ��

    On Error GoTo err
    
    '��ȡ��������ص����̹������
    mBeforeDays = 1 'Val(GetDeptPara(mlngCur����ID, "Ĭ�Ϲ�������", 2)) '                   'Ĭ�Ϲ�������
    If mBeforeDays > 15 Or mBeforeDays <= 0 Then
        mBeforeDays = 2
    End If


    If intType = 0 Then    '�Ӳ˵�����FormLoad�������Ҹı䣬ˢ�²��˹��˿�ʼʱ��
        SQLCondition.��ʼʱ�� = CDate(Format(zlDatabase.Currentdate - mBeforeDays, "yyyy-mm-dd 00:00"))
        mblnMoved = MovedByDate(SQLCondition.��ʼʱ��)
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub


Private Sub cmdInfo_Click()
    On Error GoTo errHandle
    frmDegreeCard.ShowMe Val(vsList.TextMatrix(vsList.Row, GetCN("����ID"))), Val(vsList.TextMatrix(vsList.Row, GetCN("��ҳID")))
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdReportView_Click()
            'û���治�ܴ�ӡ��Ԥ��
            If vsList.TextMatrix(vsList.Row, GetCN("������")) = "" Then
                MsgBoxD Me, "��ǰ����û�м�鱨�棬���ܲ��������飡", vbInformation, gstrSysName
                Exit Sub
            End If
            
            Call OpenReportPreview(mlngAdviceID)
End Sub

Private Sub DkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = picList.Hwnd
    ElseIf Item.ID = 2 Then
        Item.Handle = PicWindow.Hwnd
    End If
End Sub
Private Sub Form_Load()
    mstrPrivs = gstrPrivs           'Ȩ��
    mlngModul = glngModul           'ģ���
'    mlngCur����ID = 0
    mstrҽ���������� = ""
'    mblnAllDepts = False            'Ĭ�ϲ�ѡ��ȫ������
    mlngSortCol = 0
    mintSortOrder = 0
    
    mblnInitOk = False  '��ʼ����,��ʼ�����֮ǰ���������ݵ���ȡ
    mblnvsRefresh = False
    
    ucFilter.CardNames = "����;���￨;�����;סԺ��;���ݺ�;����;���֤��;������;IC����;"
    Call ucFilter.InitCardType(glngSys, mlngModul, UserInfo.����, gcnOracle)
    
    ucLocate.CardNames = "����;���￨;�����;סԺ��;���ݺ�;����;���֤��;������;IC����;"
    Call ucLocate.InitCardType(glngSys, mlngModul, UserInfo.����, gcnOracle)
    
    Call InitLocalPars '����ע������
    If Not InitDepts Then Unload Me: Exit Sub '��ʼ��ҽ������
    
    ReDim gConnectedShardDir(0) As String   '��ʼ������Ŀ¼���Ӵ�
    
    Call InitMvar(0) '��ʼ��ģ�鼶����
    '��ʼ�Ӵ���
    Set mfrmPACSImg = New frmPACSImg
    
    Call mobjRichEPR.InitRichEPR(gcnOracle, Me, glngSys, False)
    
    Set mobjPacsCore = New zl9PacsCore.clsViewer

    Call InitFilterCmd
    Call InitCommandBars
    Call InitSubForm
    Call InitFaceScheme
    Call InitList

    Set mfrmPACSImg.pobjPacsCore = mobjPacsCore
    
    mblnInitOk = True '��ʼ�����
    
    Call RestoreWinState(Me, App.ProductName)
    '���ܱ�restorewinstate���������д�����
    Call RefreshList
    
    ClearCacheFolder App.Path & "\TmpImage\"    '����ʱĿ¼���ˣ�����ո�Ŀ¼
    Me.stbThis.Panels(3).Text = "����ҽ����" & UserInfo.����
End Sub

Private Function InitDepts() As Boolean
'���ܣ���ʼ��ҽ����������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str����IDs As String
    
    On Error GoTo errH
    
    strSQL = "select distinct A.ID, A.����,A.���� from ���ű� A, ������Ա B Where a.ID = b.����ID And b.��Աid = [1]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, glngUserId)
    
    If rsTmp.EOF Then
        MsgBoxD Me, "û�з�������������Ϣ,���ȵ����Ź��������á�", vbInformation, gstrSysName
        Exit Function
    Else
        Do Until rsTmp.EOF
            mstrҽ���������� = mstrҽ���������� & "|" & rsTmp!ID & "_" & rsTmp!���� & "-" & rsTmp!����
            mstrҽ����������IDs = mstrҽ����������IDs & "," & rsTmp!ID
            
            rsTmp.MoveNext
        Loop
        
        mstrҽ���������� = Mid(mstrҽ����������, 2)
        mstrҽ����������IDs = Mid(mstrҽ����������IDs, 2)
        

        InitDepts = True
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Dim strTemp As String
    Dim i As Integer
    
    On Error Resume Next
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���ﲡ��", IIf(mblncmd����, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "סԺ����", IIf(mblncmdסԺ, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���ﲡ��", IIf(mblncmd����, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��첡��", IIf(mblncmd���, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�����ѽ�", IIf(mblncmd�ѽ�, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����δ��", IIf(mblncmdδ��, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�Ǽǲ���", IIf(mblncmd�Ǽ�, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��������", IIf(mblncmd����, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��鲡��", IIf(mblncmd���, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���没��", IIf(mblncmd����, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��˲���", IIf(mblncmd���, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��ɲ���", IIf(mblncmd���, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���˷�ʽ", ucFilter.CurCardName
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��λ��ʽ", ucLocate.CurCardName
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����סԺ", IIf(mblncmd����, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "������", mlngSortCol
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "������", mintSortOrder
    
    If UBound(mblncmdӰ�����) >= 0 Then
        strTemp = mblncmdӰ�����(0)
    End If
    For i = 1 To UBound(mblncmdӰ�����)
        strTemp = strTemp & "," & mblncmdӰ�����(i)
    Next i
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "Ӱ��������", strTemp
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
    Call SaveWinState(Me, App.ProductName)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(vsList), vsList.Name, mstrCol)

    
    Unload mfrmPACSImg


    If Not mobjPacsCore Is Nothing Then mobjPacsCore.Closefrom
    
    Set mobjIDCard = Nothing
    Set mobjPacsCore = Nothing
    Set mobjRichEPR = Nothing
End Sub

Private Function GetCN(ByVal Col As String) As Integer
Dim arrCol As Variant, i As Integer
    If mstrCol = "" Then mstrCol = ConstrCol
    arrCol = Split(mstrCol, "|")
    For i = 0 To UBound(arrCol)
        'If InStr(arrCol(i), Col) > 0 Then GetCN = i: Exit Function
        If Split(arrCol(i), ";")(0) = Col Then GetCN = i: Exit Function
    Next
    GetCN = 0
End Function

Private Function GetCW(ByVal Col As String) As Long
    Dim arrCol As Variant, i As Integer
    arrCol = Split(mstrCol, "|")
    For i = 0 To UBound(arrCol)
        'If InStr(arrCol(i), Col) > 0 Then GetCW = Split(arrCol(i), ";")(1): Exit Function
        If Split(arrCol(i), ";")(0) = Col Then GetCW = Split(arrCol(i), ";")(1): Exit Function
    Next
    GetCW = 0
End Function

Private Sub InitLocalPars()
'��ʼ����ʱ���ز������Ը������ã�ע������Ϊ��,������أ��������õȵ���
    Dim strTemp As String
    Dim strTempArry() As String
    Dim i As Integer
    
    On Error GoTo err
    mblncmd���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���ﲡ��", 1))
    mblncmdסԺ = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "סԺ����", 1))
    mblncmd���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���ﲡ��", 1))
    mblncmd��� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��첡��", 1))
    mblncmd�ѽ� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�����ѽ�", 0))
    mblncmdδ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����δ��", 0))
    mblncmd�Ǽ� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�Ǽǲ���", 1))
    mblncmd���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��������", 1))
    mblncmd��� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��鲡��", 1))
    mblncmd���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���没��", 1))
    mblncmd��� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��˲���", 1))
    mblncmd��� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��ɲ���", 1))

    ucFilter.CurCardName = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���˷�ʽ", "����")
    ucLocate.CurCardName = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��λ��ʽ", "����")
    mblncmd���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����סԺ", "0"))
    mlngSortCol = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "������", 0))
    mintSortOrder = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "������", 0))
    
    strTemp = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "Ӱ��������", "")
    ReDim strTempArry(0)
    ReDim mblncmdӰ�����(0)
    On Error Resume Next
    strTempArry = Split(strTemp, ",")
    If UBound(strTempArry) >= 0 Then ReDim mblncmdӰ�����(UBound(strTempArry))
    For i = 0 To UBound(strTempArry)
        mblncmdӰ�����(i) = IIf(UCase(strTempArry(i)) = "TRUE", True, False)
    Next i
    
    On Error GoTo err
    
    '��ȡ��ά�ؽ�����
    mblnUse3D = Val(zlDatabase.GetPara("������ά�ؽ�", glngSys, mlngModul, 0))
    mstr3DExeDir = zlDatabase.GetPara("3D����·��", glngSys, mlngModul, "")
    mstr3DPara = zlDatabase.GetPara("3D����", glngSys, mlngModul, "")
    mstr3DFunctions = zlDatabase.GetPara("3D����", glngSys, mlngModul, "")

    With SQLCondition '------------------------ '����������ʼ
        'ʱ���ѯ��ʽ 1=������ʱ�䣨����ҽ������.����ʱ�䣩��2=������ʱ�䣨����ҽ������.�״�ʱ�䣩��3=��ͼʱ�䣨Ӱ�����¼.�������ڣ�
        .ʱ������ = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "����ʱ������", 1))
        .���ݺ� = ""
        .����� = 0
        .סԺ�� = 0
        .���￨ = ""
        .���� = ""
        .�Ա� = ""
        .��ʼ���� = -1
        .�������� = -1
        .�������� = "="
        .���� = 0
        .���֤ = ""
        .IC�� = ""
        .���˿��� = 0
        .�걾��λ = ""
        .���ҽ�� = ""
        .���ҽ�� = ""
        .������� = ""
        .�������� = ""
        .������� = -1
        .Ӱ������ = ""
        .��鼼ʦ = ""
        .������ = ""
        .Ӱ����� = ""
        .������� = ""
        .������ = ""
        .���� = ""
        .��� = ""
    End With
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitFaceScheme()
    '��ʼ���沼��
    Dim Pane1 As Pane, Pane2 As Pane
    With Me.dkpMain
        .SetCommandBars cbrMain
        .Options.HideClient = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    Set Pane1 = dkpMain.CreatePane(1, 240, 250, DockLeftOf, Nothing)
    Pane1.Title = "����б�"
    Pane1.Handle = picList.Hwnd
    Pane1.Options = PaneNoCloseable Or PaneNoFloatable
    
    Set Pane2 = dkpMain.CreatePane(2, 700, 250, DockRightOf, Nothing)
    Pane2.Title = "�Ӵ���"
    Pane2.Handle = PicWindow.Hwnd
    Pane2.Options = PaneNoCaption Or PaneNoCloseable
    dkpMain.LoadStateFromString GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, "")
End Sub

Private Sub InitFilterCmd()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl, cbrPopControl As CommandBarControl
    Dim objPopbar As CommandBarPopup, objCusControl As CommandBarControlCustom
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    Dim i As Integer

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbrdock.VisualTheme = xtpThemeOfficeXP
    With Me.cbrdock.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
        .UseSharedImageList = False 'ImageList��ʽʱ,��ͬһApp�й���,��AddImageList֮ǰ����ΪFalse
    End With
    cbrdock.AddImageList img16 '��VB.ImageList��Tag��ID���й���
    cbrdock.EnableCustomization False
    cbrdock.ActiveMenuBar.Visible = False
    
    Set objBar = cbrdock.Add("��Դ", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, ID_����, "����")
            objControl.ToolTipText = "��ʾ���ﲡ��"
        Set objControl = .Add(xtpControlButton, ID_סԺ, "סԺ")
            objControl.ToolTipText = "��ʾסԺ����"
        Set objControl = .Add(xtpControlButton, ID_����, "����")
            objControl.ToolTipText = "��ʾ���ﲡ��"
        Set objControl = .Add(xtpControlButton, ID_���, "���")
            objControl.ToolTipText = "��ʾ��첡��"
        Set objControl = .Add(xtpControlButtonPopup, ID_����, " ��  ��")
            objControl.ToolTipText = "��ʾ�����ѽ�/δ�ɲ���"
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_δ��, "δ��")
            cbrPopControl.ToolTipText = "��ʾ����δ�ɲ���"
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_�ѽ�, "�ѽ�")
            cbrPopControl.ToolTipText = "��ʾ�����ѽɲ���"
        
        '�������Ӱ�����
        Set objControl = .Add(xtpControlButtonPopup, ID_Ӱ�����, "Ӱ�����")
        objControl.ToolTipText = "��ʾӰ�����"
        strSQL = "select ����,���� from Ӱ�������"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "Ӱ�������")
        i = 1
        mintcmdӰ����� = 0
        strTemp = ""
        ReDim Preserve mblncmdӰ�����(rsTemp.RecordCount - 1)
        While rsTemp.EOF = False
            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_Ӱ����� + i, rsTemp("����"))
            cbrPopControl.DescriptionText = rsTemp("����")
            cbrPopControl.Style = xtpButtonIconAndCaption
            cbrPopControl.Checked = mblncmdӰ�����(i - 1)
            cbrPopControl.CloseSubMenuOnClick = False
            If mblncmdӰ�����(i - 1) = True Then
                mintcmdӰ����� = mintcmdӰ����� + 1
                strTemp = IIf(strTemp = "", cbrPopControl.Caption, strTemp & "," & cbrPopControl.Caption)
            End If
            rsTemp.MoveNext
            i = i + 1
        Wend
        If strTemp <> "" Then objControl.Caption = strTemp
    End With
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    Set objBar = cbrdock.Add("״̬", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, ID_�Ǽ�, "�Ǽ�")
            objControl.ToolTipText = "��ʾ�ѵǼǲ���"
        Set objControl = .Add(xtpControlButton, ID_����, "����")
            objControl.ToolTipText = "��ʾ�ѱ�������"
        Set objControl = .Add(xtpControlButton, ID_���, "���")
            objControl.ToolTipText = "��ʾ�Ѽ�鲡��"
        Set objControl = .Add(xtpControlButton, ID_����, "����")
            objControl.ToolTipText = "��ʾ�ѱ��没��"
        Set objControl = .Add(xtpControlButton, ID_���, "���")
            objControl.ToolTipText = "��ʾ����˲���"
        Set objControl = .Add(xtpControlButton, ID_���, "���")
            objControl.ToolTipText = "��ʾ����ɲ���"
    End With
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    Set objBar = cbrdock.Add("����", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
        
    Set objCusControl = objBar.Controls.Add(xtpControlCustom, ID_����ֵ, "����ֵ")
        objCusControl.Handle = ucFilter.Handle
        objCusControl.Flags = xtpFlagRightAlign
        
    Set objControl = objBar.Controls.Add(xtpControlButton, ID_��ʼ����, "��ʼ����")
        objControl.Style = xtpButtonIconAndCaption
        objControl.IconId = conMenu_View_Filter
        
    Set objControl = objBar.Controls.Add(xtpControlButton, ID_����סԺ, "����")
    objControl.ToolTipText = "ֻ��ʾ����סԺ����¼"
    objControl.Style = xtpButtonIconAndCaption
    objControl.IconId = conMenu_View_Filter
    
    With cbrdock.KeyBindings
        .Add FCONTROL, vbKey0, ID_����
        .Add FCONTROL, vbKey1, ID_סԺ
        .Add FCONTROL, vbKey2, ID_����
        .Add FCONTROL, vbKey3, ID_���
        
        .Add FCONTROL, vbKey4, ID_�Ǽ�
        .Add FCONTROL, vbKey5, ID_����
        .Add FCONTROL, vbKey6, ID_���
        .Add FCONTROL, vbKey7, ID_����
        .Add FCONTROL, vbKey8, ID_���
        .Add FCONTROL, vbKey9, ID_���
        .Add FCONTROL, Asc("G"), ID_��ʼ����
    End With
    cbrdock.RecalcLayout
End Sub

Private Sub InitCommandBars()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    Dim str3DFuncs() As String
    Dim i As Integer
    Dim i3DFunc As Integer

    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    Set Me.cbrMain.Icons = zlCommFun.GetPubIcons
    
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    

'�˵�����
'Begin------------------------�ļ��˵�--------------------------------------Ĭ�Ͽɼ�
    Me.cbrMain.ActiveMenuBar.Title = "�˵�"
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)"): cbrControl.IconId = 181
        
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "�����ӡ(&P)"): cbrControl.IconId = 103
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�嵥��ӡ(&L)"): cbrControl.BeginGroup = True: cbrControl.IconId = 103
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Change_In, "�����б�")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"):: cbrControl.IconId = 191: cbrControl.BeginGroup = True
    End With


'Begin----------------------���˵�--------------------------------------Ĭ�Ͽɼ�
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "���(&S)", -1, False)
    cbrMenuBar.ID = conMenu_ManagePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "����鿴(&V)"): cbrControl.IconId = 102:  cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Look, "Ӱ���Ƭ(&S)"): cbrControl.IconId = 8111
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Contrast, "��Ƭ�Ա�(&E)"): cbrControl.IconId = 8112
        
        '���������ά�ؽ����ܣ��򴴽���Ӧ�˵�
        If mblnUse3D = True Then
            Set cbrControl = .Add(xtpControlPopup, conMenu_Img_3D, "��ά�ؽ�"): cbrControl.ID = conMenu_Img_3D
                If mstr3DFunctions <> "" Then
                    str3DFuncs = Split(mstr3DFunctions, ",")
                    For i = 1 To UBound(str3DFuncs)
                        i3DFunc = Val(str3DFuncs(i))
                        If i3DFunc >= 1 And i3DFunc <= 6 Then
                            Select Case i3DFunc
                                Case 1
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_VA, "�ݻ��ؽ�")
                                Case 2
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_MPR, "MPR")
                                Case 3
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_MMPR, "MMPR")
                                Case 4
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_VE, "�����ڿ���")
                                Case 5
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_SA, "�����ؽ�")
                                Case 6
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_PF, "��ע����")
                            End Select
                        End If
                    Next i
                End If
        End If
        
    End With
    
    
'Begin----------------------�鿴�˵�--------------------------------------
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.ID = conMenu_View_ToolBar
            With cbrControl.CommandBar.Controls '�����˵�
                Set cbrPopControl = .Add(xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False): cbrPopControl.Checked = True
                Set cbrPopControl = .Add(xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False): cbrPopControl.Checked = True
                Set cbrPopControl = .Add(xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False): cbrPopControl.Checked = True
            End With
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)"): cbrControl.Checked = True: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_Manage_LocateType, "��λ��ʽ(&G)"): cbrControl.ID = conMenu_Manage_LocateType
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_View_Filter * 10#, "��ǰ����"): cbrControl.ID = conMenu_View_Filter * 10#
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "���ٹ���(&K)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&F)")
    End With


'Begin----------------------�����˵�--------------------------------------Ĭ�Ͽɼ�
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������", -1, False)
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_Help_Web, "WEB�ϵ�����(&E)")
            With cbrControl.CommandBar.Controls
                Set cbrPopControl = .Add(xtpControlButton, conMenu_Help_Web_Forum, "������̳(&F)", -1, False)
                Set cbrPopControl = .Add(xtpControlButton, conMenu_Help_Web_Home, "������ҳ(&H)", -1, False)
                Set cbrPopControl = .Add(xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False)
            End With
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With
    
    '��ȡ��������ģ��ı���(��������ģ���)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(cbrMain, glngSys, mlngModul, mstrPrivs)
    
'----------------------�����------------------------------------------
    With Me.cbrMain.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print '��ӡ------------------Ctrl+P
        
        .Add 0, VK_F1, conMenu_Help_Help              '����-------------F1
        .Add 0, VK_F5, conMenu_View_Refresh           'ˢ��-------------F5
        .Add FCONTROL, Asc("G"), conMenu_Manage_LocateType    '��λ��ʽ---------Ctrl+F
        .Add 0, VK_F3, conMenu_View_Filter            '����-------------F3
    End With
    
'---------------------�������Ͻǵ�ǰ����----------------------------------
        Set cbrControl = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_Filter * 10#, "��ǰ����")
            cbrControl.ID = conMenu_View_Filter * 10#
            cbrControl.Flags = xtpFlagRightAlign
            cbrControl.Category = "Main"
            
        Set cbrCustom = cbrMain.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_Manage_LocateValue, "��λ����")
            cbrCustom.Handle = ucLocate.Handle
            cbrCustom.Flags = xtpFlagRightAlign
            cbrCustom.Style = xtpButtonIconAndCaption
            cbrCustom.Category = "Main"
    
'---------------------����������------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True
'    cbrToolBar.EnableDocking xtpFlagStretched '+ xtpFlagHideWrap
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "����"): cbrControl.IconId = 102: cbrControl.ToolTipText = "����鿴"
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ"): cbrControl.IconId = 103: cbrControl.ToolTipText = "�����ӡ"
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Look, "��Ƭ"): cbrControl.ToolTipText = "Ӱ���Ƭ"
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Contrast, "�Ա�"): cbrControl.IconId = 8112: cbrControl.ToolTipText = "��Ƭ�Ա�"
        '���������ά�ؽ����ܣ��򴴽���Ӧ�˵�
        If mblnUse3D = True Then
            Set cbrControl = .Add(xtpControlPopup, conMenu_Img_3D, "��ά"): cbrControl.ID = conMenu_Img_3D: cbrControl.ToolTipText = "��ά�ؽ�"
                If mstr3DFunctions <> "" Then
                    str3DFuncs = Split(mstr3DFunctions, ",")
                    For i = 1 To UBound(str3DFuncs)
                        i3DFunc = Val(str3DFuncs(i))
                        If i3DFunc >= 1 And i3DFunc <= 6 Then
                            Select Case i3DFunc
                                Case 1
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_VA, "�ݻ��ؽ�")
                                Case 2
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_MPR, "MPR")
                                Case 3
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_MMPR, "MMPR")
                                Case 4
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_VE, "�����ڿ���")
                                Case 5
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_SA, "�����ؽ�")
                                Case 6
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_PF, "��ע����")
                            End Select
                        End If
                    Next i
                End If
        End If
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
        
        
    End With

End Sub

Private Sub InitSubForm()
Dim i As Integer
    With TabWindow
        .RemoveAll
        Set .Icons = zlCommFun.GetPubIcons
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTaskPanelHighlightNone
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .PaintManager.ColorSet.ButtonSelected = &HFFC0C0
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True
        .RemoveAll
        
        .InsertItem 0, "Ӱ���¼", mfrmPACSImg.Hwnd, conMenu_Img_Look
        .Item(TabWindow.ItemCount - 1).Tag = "Ӱ��ͼ��"
        
        
    End With

End Sub


Private Sub InitList()
'��ʼ�����
Dim C���� As Long, C��Դ As Long, C���� As Long, C���� As Long, C���� As Long, C���� As Long, C������ As Long, C�Ա� As Long, C���� As Long
Dim C��ʶ�� As Long, Cҽ������ As Long, C��λ���� As Long, Cִ�м� As Long, C����ʱ�� As Long, C����ʱ�� As Long, C����ҽ�� As Long
Dim C��� As Long, C���� As Long, CӤ�� As Long, C�Ǽ��� As Long, C������ As Long, C����� As Long, C��ӡ��Ƭ As Long, C������� As Long
Dim C��ɫͨ�� As Long, C�����ӡ As Long, C������ As Long, C������ As Long, C��鼼ʦ As Long, C��ͼʱ�� As Long, C������� As Long
Dim CӰ����� As Long, C����ID As Long, C��ҳID As Long, C�Һŵ� As Long, C���˿���ID As Long, Cҽ��ID As Long, C���ͺ� As Long, C���UID As Long
Dim C���״̬ As Long, CNO As Long, C��¼���� As Long, Cת�� As Long, C���� As Long, C��ǰ����ID As Long, C���淢�� As Long
Dim C��Ϸ��� As Long, Cִ�п���ID As Long, C����ID As Long, C���˿��� As Long, C���￨�� As Long, C���ݺ� As Long, C���֤�� As Long
 
    If mstrCol = "" Then
        mstrCol = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(vsList), vsList.Name, ConstrCol)
        '�ж��Ƿ��޸Ĺ���ʾ������������޸Ĺ������ȡĬ��ֵ�������Ƕ�ȡע���
        If UBound(Split(mstrCol, "|")) <> UBound(Split(ConstrCol, "|")) Then
            mstrCol = ConstrCol
        End If
    End If
    With vsList
        .Clear
        .FixedRows = 1
        .Rows = 2
        .Cols = 53
        '��ȡ����
        C���� = GetCN("����"):           C��Դ = GetCN("��Դ"):          C���� = GetCN("����")
        C���� = GetCN("����"):          C���� = GetCN("����"):          C���� = GetCN("����")
        C������ = GetCN("������"):  C�Ա� = GetCN("�Ա�"):          C���� = GetCN("����")
        C��ʶ�� = GetCN("��ʶ��"):      Cҽ������ = GetCN("ҽ������"):  C��λ���� = GetCN("��λ����")
        Cִ�м� = GetCN("ִ�м�"):      C����ʱ�� = GetCN("����ʱ��"):  C����ʱ�� = GetCN("����ʱ��")
        C����ҽ�� = GetCN("����ҽ��"):   C��� = GetCN("���"):          C���� = GetCN("����")
        CӤ�� = GetCN("Ӥ��"):          C�Ǽ��� = GetCN("�Ǽ���"):      C������ = GetCN("������")
        C����� = GetCN("�����"):      C��ӡ��Ƭ = GetCN("��ӡ��Ƭ"):  C������� = GetCN("�������")
        C��ɫͨ�� = GetCN("��ɫͨ��"):  C�����ӡ = GetCN("�����ӡ"):  C������ = GetCN("������")
        C������ = GetCN("������"):      C��鼼ʦ = GetCN("��鼼ʦ"):  C��ͼʱ�� = GetCN("��ͼʱ��")
        C������� = GetCN("�������"):  CӰ����� = GetCN("Ӱ�����"):  C����ID = GetCN("����ID")
        C��ҳID = GetCN("��ҳID"):      C�Һŵ� = GetCN("�Һŵ�"):      Cҽ��ID = GetCN("ҽ��ID")
        C���ͺ� = GetCN("���ͺ�"):      C���˿���ID = GetCN("���˿���ID"): C���UID = GetCN("���UID")
        C���״̬ = GetCN("���״̬"):  CNO = GetCN("NO"):              C��¼���� = GetCN("��¼����")
        Cת�� = GetCN("ת��"):          C���� = GetCN("����"):          C��ǰ����ID = GetCN("��ǰ����ID")
        C���淢�� = GetCN("���淢��"):  C��Ϸ��� = GetCN("��Ϸ���"):  Cִ�п���ID = GetCN("ִ�п���ID")
        C����ID = GetCN("����ID"):      C���˿��� = GetCN("���˿���"):  C���￨�� = GetCN("���￨��")
        C���ݺ� = GetCN("���ݺ�"):      C���֤�� = GetCN("���֤��")
        '��ȡ��ָ���п�
        .ColWidth(C����) = GetCW("����"):           .ColWidth(C��Դ) = GetCW("��Դ"):           .ColWidth(C����) = GetCW("����")
        .ColWidth(C����) = GetCW("����"):           .ColWidth(C����) = GetCW("����"):           .ColWidth(C����) = GetCW("����")
        .ColWidth(C������) = GetCW("������"):   .ColWidth(C�Ա�) = GetCW("�Ա�"):           .ColWidth(C����) = GetCW("����")
        .ColWidth(C��ʶ��) = GetCW("��ʶ��"):       .ColWidth(Cҽ������) = GetCW("ҽ������"):   .ColWidth(C��λ����) = GetCW("��λ����")
        .ColWidth(Cִ�м�) = GetCW("ִ�м�"):       .ColWidth(C����ʱ��) = GetCW("����ʱ��"):   .ColWidth(C����ʱ��) = GetCW("����ʱ��")
        .ColWidth(C����ҽ��) = GetCW("����ҽ��"):   .ColWidth(C���) = GetCW("���"):           .ColWidth(C����) = GetCW("����")
        .ColWidth(CӤ��) = GetCW("Ӥ��"):           .ColWidth(C�Ǽ���) = GetCW("�Ǽ���"):       .ColWidth(C������) = GetCW("������")
        .ColWidth(C�����) = GetCW("�����"):       .ColWidth(C��ӡ��Ƭ) = GetCW("��ӡ��Ƭ"):   .ColWidth(C�������) = GetCW("�������")
        .ColWidth(C��ɫͨ��) = GetCW("��ɫͨ��"):   .ColWidth(C�����ӡ) = GetCW("�����ӡ"):   .ColWidth(C������) = GetCW("������")
        .ColWidth(C������) = GetCW("������"):       .ColWidth(C��鼼ʦ) = GetCW("��鼼ʦ"):   .ColWidth(C��ͼʱ��) = GetCW("��ͼʱ��")
        .ColWidth(C�������) = GetCW("�������"):   .ColWidth(CӰ�����) = GetCW("Ӱ�����"):   .ColWidth(C����ID) = GetCW("����ID")
        .ColWidth(C��ҳID) = GetCW("��ҳID"):       .ColWidth(C�Һŵ�) = GetCW("�Һŵ�"):       .ColWidth(Cҽ��ID) = GetCW("ҽ��ID")
        .ColWidth(C���ͺ�) = GetCW("���ͺ�"):       .ColWidth(C���˿���ID) = GetCW("���˿���ID"): .ColWidth(C���UID) = GetCW("���UID")
        .ColWidth(C���״̬) = GetCW("���״̬"):   .ColWidth(CNO) = GetCW("NO"):               .ColWidth(C��¼����) = GetCW("��¼����")
        .ColWidth(Cת��) = GetCW("ת��"):           .ColWidth(C����) = GetCW("����"):           .ColWidth(C��ǰ����ID) = GetCW("��ǰ����ID")
        .ColWidth(C���淢��) = GetCW("���淢��"):   .ColWidth(C��Ϸ���) = GetCW("��Ϸ���"):   .ColWidth(Cִ�п���ID) = GetCW("ִ�п���ID")
        .ColWidth(C����ID) = GetCW("����ID"):       .ColWidth(C���˿���) = GetCW("���˿���"):   .ColWidth(C���￨��) = GetCW("���￨��")
        .ColWidth(C���ݺ�) = GetCW("���ݺ�"):       .ColWidth(C���֤��) = GetCW("���֤��")
        
        '������
        .Cell(flexcpData, 0, C����) = "����":               .Cell(flexcpData, 0, C��Դ) = "��Դ":               .Cell(flexcpData, 0, C����) = "����"
        .Cell(flexcpData, 0, C����) = "����":               .Cell(flexcpData, 0, C����) = "����":               .Cell(flexcpData, 0, C����) = "����"
        .Cell(flexcpData, 0, C������) = "������":       .Cell(flexcpData, 0, C�Ա�) = "�Ա�":               .Cell(flexcpData, 0, C����) = "����"
        .Cell(flexcpData, 0, C��ʶ��) = "��ʶ��":           .Cell(flexcpData, 0, Cҽ������) = "ҽ������":       .Cell(flexcpData, 0, C��λ����) = "��λ����"
        .Cell(flexcpData, 0, Cִ�м�) = "ִ�м�":           .Cell(flexcpData, 0, C����ʱ��) = "����ʱ��":       .Cell(flexcpData, 0, C����ʱ��) = "����ʱ��"
        .Cell(flexcpData, 0, C����ҽ��) = "����ҽ��":       .Cell(flexcpData, 0, C���) = "���":               .Cell(flexcpData, 0, C����) = "����"
        .Cell(flexcpData, 0, CӤ��) = "Ӥ��":               .Cell(flexcpData, 0, C�Ǽ���) = "�Ǽ���":           .Cell(flexcpData, 0, C������) = "������"
        .Cell(flexcpData, 0, C�����) = "�����":           .Cell(flexcpData, 0, C��ӡ��Ƭ) = "��ӡ��Ƭ":       .Cell(flexcpData, 0, C�������) = "�������"
        .Cell(flexcpData, 0, C��ɫͨ��) = "��ɫͨ��":       .Cell(flexcpData, 0, C�����ӡ) = "�����ӡ":       .Cell(flexcpData, 0, C������) = "������"
        .Cell(flexcpData, 0, C������) = "������":           .Cell(flexcpData, 0, C��鼼ʦ) = "��鼼ʦ":       .Cell(flexcpData, 0, C��ͼʱ��) = "��ͼʱ��"
        .Cell(flexcpData, 0, C�������) = "�������":       .Cell(flexcpData, 0, CӰ�����) = "Ӱ�����":       .Cell(flexcpData, 0, C����ID) = "����ID"
        .Cell(flexcpData, 0, C��ҳID) = "��ҳID":           .Cell(flexcpData, 0, C�Һŵ�) = "�Һŵ�":           .Cell(flexcpData, 0, C���˿���ID) = "���˿���ID"
        .Cell(flexcpData, 0, Cҽ��ID) = "ҽ��ID":           .Cell(flexcpData, 0, C���ͺ�) = "���ͺ�":           .Cell(flexcpData, 0, C���UID) = "���UID"
        .Cell(flexcpData, 0, C���״̬) = "���״̬":       .Cell(flexcpData, 0, CNO) = "NO":                   .Cell(flexcpData, 0, C��¼����) = "��¼����"
        .Cell(flexcpData, 0, Cת��) = "ת��":               .Cell(flexcpData, 0, C����) = "����":               .Cell(flexcpData, 0, C��ǰ����ID) = "��ǰ����ID"
        .Cell(flexcpData, 0, C���淢��) = "���淢��":       .Cell(flexcpData, 0, C��Ϸ���) = "��Ϸ���":       .Cell(flexcpData, 0, Cִ�п���ID) = "ִ�п���ID"
        .Cell(flexcpData, 0, C����ID) = "����ID":           .Cell(flexcpData, 0, C���˿���) = "���˿���":       .Cell(flexcpData, 0, C���￨��) = "���￨��"
        .Cell(flexcpData, 0, C���ݺ�) = "���ݺ�":           .Cell(flexcpData, 0, C���֤��) = "���֤��"
        
        '��ʾ������
        Set .Cell(flexcpPicture, 0, C����) = Imglist.ListImages("����").Picture
        Set .Cell(flexcpPicture, 0, C��Դ) = Imglist.ListImages("סԺ").Picture
        Set .Cell(flexcpPicture, 0, C����) = Imglist.ListImages("����").Picture
        .TextMatrix(0, C����) = "��":               .TextMatrix(0, C����) = "����":              .TextMatrix(0, C����) = "����"
        .TextMatrix(0, C������) = "������":     .TextMatrix(0, C�Ա�) = "�Ա�":             .TextMatrix(0, C����) = "����"
        .TextMatrix(0, C��ʶ��) = "��ʶ��":         .TextMatrix(0, Cҽ������) = "ҽ������":     .TextMatrix(0, C��λ����) = "��λ����"
        .TextMatrix(0, Cִ�м�) = "ִ�м�":         .TextMatrix(0, C����ʱ��) = "����ʱ��":     .TextMatrix(0, C����ʱ��) = "����ʱ��"
        .TextMatrix(0, C����ҽ��) = "����ҽ��":     .TextMatrix(0, C���) = "���":             .TextMatrix(0, C����) = "����"
        .TextMatrix(0, CӤ��) = "Ӥ��":             .TextMatrix(0, C�Ǽ���) = "�Ǽ���":         .TextMatrix(0, C������) = "������"
        .TextMatrix(0, C�����) = "�����":         .TextMatrix(0, C��ӡ��Ƭ) = "��ӡ��Ƭ":     .TextMatrix(0, C�������) = "�������"
        .TextMatrix(0, C��ɫͨ��) = "��ɫͨ��":     .TextMatrix(0, C�����ӡ) = "�����ӡ":     .TextMatrix(0, C������) = "������"
        .TextMatrix(0, C������) = "������":         .TextMatrix(0, C��鼼ʦ) = "��鼼ʦ":     .TextMatrix(0, C��ͼʱ��) = "��ͼʱ��"
        .TextMatrix(0, C�������) = "�������":     .TextMatrix(0, CӰ�����) = "Ӱ�����":     .TextMatrix(0, C����ID) = "����ID"
        .TextMatrix(0, C��ҳID) = "��ҳID":         .TextMatrix(0, C�Һŵ�) = "�Һŵ�":         .TextMatrix(0, C���˿���ID) = "���˿���ID"
        .TextMatrix(0, Cҽ��ID) = "ҽ��ID":         .TextMatrix(0, C���ͺ�) = "���ͺ�":         .TextMatrix(0, C���UID) = "���UID"
        .TextMatrix(0, C���״̬) = "���״̬":     .TextMatrix(0, CNO) = "NO":                 .TextMatrix(0, C��¼����) = "��¼����"
        .TextMatrix(0, Cת��) = "ת��":             .TextMatrix(0, C����) = "����":             .TextMatrix(0, C��ǰ����ID) = "��ǰ����ID"
        .TextMatrix(0, C���淢��) = "���淢��":     .TextMatrix(0, C��Ϸ���) = "��Ϸ���":     .TextMatrix(0, Cִ�п���ID) = "ִ�п���ID"
        .TextMatrix(0, C����ID) = "����ID":         .TextMatrix(0, C���˿���) = "���˿���":     .TextMatrix(0, C���￨��) = "���￨��"
        .TextMatrix(0, C���ݺ�) = "���ݺ�":         .TextMatrix(0, C���֤��) = "���֤��"
        
        Dim i As Integer
        For i = 0 To .Cols - 1
            .ColAlignment(i) = flexAlignLeftCenter
        Next
        
        '��ȡ�����ò����б������
        .FontName = zlDatabase.GetPara("�����б���������", glngSys, mlngModul, "����")
        .FontSize = Val(zlDatabase.GetPara("�����б������ֺ�", glngSys, mlngModul, 9))
        .FontBold = zlDatabase.GetPara("�����б����ݴ���", glngSys, mlngModul, 0) = 1
        .FontItalic = zlDatabase.GetPara("�����б�����б��", glngSys, mlngModul, 0) = 1
        .Cell(flexcpFontName, 0, 0, 0, .Cols - 1) = zlDatabase.GetPara("�����б��ͷ����", glngSys, mlngModul, "����")
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = Val(zlDatabase.GetPara("�����б��ͷ�ֺ�", glngSys, mlngModul, 9))
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = zlDatabase.GetPara("�����б��ͷ����", glngSys, mlngModul, 0) = 1
        .Cell(flexcpFontItalic, 0, 0, 0, .Cols - 1) = zlDatabase.GetPara("�����б��ͷб��", glngSys, mlngModul, 0) = 1
        .Editable = flexEDNone
    End With
End Sub




Private Sub OpenViewerWithReport()
'���ݲ����򿪱����ͬʱ�򿪹�Ƭվ���ж��Ƿ�򿪹�Ƭվ
    Dim lngOrderID As Long
    
    On Error GoTo err
    
    lngOrderID = vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID"))
    
    If mblnShowImgAtReport And vsList.TextMatrix(vsList.Row, GetCN("���UID")) <> "" Then
        Dim intImageInverval As Integer
        
        intImageInverval = Val(mfrmPACSImg.cbrMain.FindControl(, conMenu_Manage_ImageInterval, , True).Text)
        Call OpenViewer(mobjPacsCore, lngOrderID, False, Me, , , False, intImageInverval)
    End If
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Function ShowBillList(objPopup As CommandBarPopup) As Boolean
'���ܣ���ʾ��ǰִ��ҽ�����Դ�ӡ�����Ƶ����ڲ˵���
    Dim rsTmp As New ADODB.Recordset
    Dim objControl As CommandBarControl
        
    On Error GoTo errH
    
    objPopup.CommandBar.Controls.DeleteAll
    With vsList
        gstrSQL = "Select Distinct C.���,C.����,C.˵��" & _
            " From ����ҽ����¼ A,��������Ӧ�� B,�����ļ��б� C" & _
            " Where A.ID=[1] And A.���ID IS NULL" & _
            " And A.������ĿID=B.������ĿID" & _
            " And B.Ӧ�ó���=[2] And B.�����ļ�ID=C.ID And C.����=7" & _
            " Order by C.���"
        If .TextMatrix(.Row, GetCN("ת��")) = 1 Then
            gstrSQL = Replace(gstrSQL, "����ҽ����¼", "H����ҽ����¼")
            gstrSQL = Replace(gstrSQL, "����ҽ������", "H����ҽ������")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(.TextMatrix(.Row, GetCN("ҽ��ID"))), CLng(Decode(.TextMatrix(.Row, GetCN("��Դ")), "��", 1, "ס", 2, "��", 3, 4)))
    End With
    
    If Not rsTmp.EOF Then
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Manage_RequestPrint * 10# + 1, rsTmp!���� & "(&0)")
            objControl.Parameter = "ZLCISBILL" & Format(rsTmp!���, "00000") & "-1" '��Ӧ���Զ��屨����
        End With
        cbrMain.KeyBindings.Add 0, vbKeyF10, conMenu_Manage_RequestPrint * 10# + 1
    End If
    
    ShowBillList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncBillPrint(objControl As CommandBarControl)
'���ܣ���ӡ���Ƶ���
    On Error GoTo errH
    If objControl.Parameter = "" Then '��֣�ֱ�Ӱ�F10ʱ����һ���յ�Control
        Set objControl = cbrMain.FindControl(, conMenu_Manage_RequestPrint * 10# + 1, , True)
        If objControl Is Nothing Then Exit Sub
    End If
    If objControl.Parameter = "" Then Exit Sub
    

    If ReportPrintSet(gcnOracle, glngSys, objControl.Parameter, Me) Then
        Call ReportOpen(gcnOracle, glngSys, objControl.Parameter, Me, "NO=" & vsList.TextMatrix(vsList.Row, GetCN("NO")), "����=" & vsList.TextMatrix(vsList.Row, GetCN("��¼����")), 1)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub RefreshList(Optional ByVal lngAdviceID As Long = 0)
Dim i As Integer, lngcurҽ��ID As Long, lngRow As Long, lngTopRow As Long
    With vsList
        If lngAdviceID <> 0 Then
            lngcurҽ��ID = lngAdviceID
        Else
            lngcurҽ��ID = Val(.TextMatrix(.Row, GetCN("ҽ��ID"))) '��ǰ��ҽ��ID
            lngRow = .Row: lngTopRow = .TopRow               '��ǰ�кͶ���֮��Ĳ��
        End If
        
        Call LoadPatiList
        If lngcurҽ��ID = 0 Then
            Call .Select(1, GetCN("����"))
            Exit Sub
        End If
        
        '�м�¼ʱҪ���¶�λ��֮ǰ��¼
        On Error Resume Next
        lngcurҽ��ID = .FindRow(CStr(lngcurҽ��ID), , GetCN("ҽ��ID"))
        If lngcurҽ��ID <> -1 Then
            lngRow = Abs(lngRow - lngTopRow)
            If .Row = lngcurҽ��ID Then '��ͬʱ���ᴥ��CHANGE�¼�
                Call vsList_RowColChange 'ǿ��ˢ���ұ��Ӵ���
            Else
                .Row = lngcurҽ��ID
            End If
            .TopRow = .Row - lngRow
        Else
            If .Row <> 1 Then
                .Row = 1
            Else
                Call vsList_RowColChange 'ǿ��ˢ���ұ��Ӵ���
            End If
        End If
        err.Clear
    End With
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    fraRegist.Left = 0
    fraRegist.Top = -75
    fraRegist.Width = picInfo.ScaleWidth
    cboTimes.Left = lblRegist.Width + 30
    cboTimes.Width = fraRegist.Width - lblRegist.Width - cmdReportView.Width - 80
    
    cmdReportView.Left = fraRegist.Width - cmdReportView.Width - 20
    
    fraInfo.Top = fraRegist.Height - 20
    fraInfo.Left = 0 'fraRegist.Left + fraRegist.Width
    fraInfo.Width = picInfo.ScaleWidth '- fraInfo.Left
    
    
    lblCash.Top = 120 '(fraInfo.Height - lblCash.Height) / 2 ' (picInfo.ScaleHeight - lblCash.Height) / 2 - fraInfo.Top
    lblCash.Left = fraInfo.Width - lblCash.Width - 100

    lbl������Ϣ.Width = lblCash.Left
    lbl�����Ϣ.Width = lblCash.Left
End Sub

Private Sub LoadPatiList()
'���ܣ���ȡ��ǰҽ�����ҵ�ִ��ҽ��(����)�嵥
Dim strSQL As String, strSQLBak As String, i As Long, rsList As ADODB.Recordset
Dim str��Դ As String
Dim strFilter As String
Dim strModalitys As String
Dim blnUseTime As Boolean       '�Ƿ�ʹ��ʱ������

    If Not mblnInitOk Then Exit Sub      '��ʼ��δ���
    mblnvsRefresh = True
    On Error GoTo errHandle
    With SQLCondition
        blnUseTime = False  'Ĭ�ϲ�ʹ��ʱ������
        '�������������ʹ��ʱ������
        If .����� <> 0 Then
            strFilter = " And C.�����=[1]"
        ElseIf .סԺ�� <> 0 Then
            strFilter = " And C.סԺ��=[2]"
        ElseIf .���￨ <> "" Then
            strFilter = " And C.���￨��=[3]"
        ElseIf .���� <> "" And InStr(.����, "*") = 0 Then   '�������⴦����*�ű�ʾģ����ѯ
            strFilter = " And C.����=[4]"
        ElseIf .���֤ <> "" Then
            strFilter = " And C.���֤��=[5]"
        ElseIf .IC�� <> "" Then
            strFilter = " And C.IC��=[6]"
        ElseIf .���ݺ� <> "" Then
            strFilter = " And A.NO=[7] "
        ElseIf .���� <> 0 Then
            strFilter = " And H.����=[8] "
        Else
        '����������ѯ��ʹ��ʱ������
            blnUseTime = True
            '��д����ʱ������
            'ʱ���ѯ��ʽ 1=������ʱ�䣨����ҽ������.����ʱ�䣩��2=������ʱ�䣨����ҽ������.�״�ʱ�䣩��3=��ͼʱ�䣨Ӱ�����¼.�������ڣ�
            If .ʱ������ = 1 Then       '������ʱ��
                strFilter = " And A.����ʱ�� Between [9] and "
            ElseIf .ʱ������ = 2 Then   '������ʱ��
                strFilter = " And A.�״�ʱ�� Between [9] and "
            Else                        '��ͼʱ��
                strFilter = " And H.�������� Between [9] and "
            End If
            If .����ʱ�� <> CDate(0) Then
                strFilter = strFilter & " [10] "
            Else
                strFilter = strFilter & " Sysdate+1/(24*3600) "
            End If
            
            '�ȴ��������д�*�ŵģ����д�ʱ��������ģ����ѯ
            If .���� <> "" And InStr(.����, "*") <> 0 Then
                .���� = Replace(.����, "*", "%")
                strFilter = strFilter & " And C.���� like [4]"
            End If
            
            If .�Ա� <> "" Then
                strFilter = strFilter & " And Nvl(H.�Ա�,C.�Ա�)=[29]"
            End If
        
        
            '��������-��ʼ����(ֻ�е�����ʹ�á����������ڶ�������֮��ʱ����ʹ�ÿ�ʼ����)
            If .��ʼ���� <> -1 Then
                If .�������� = "~" Then
                    strFilter = strFilter & " And ZL_AgeToDays(C.����)>=[30]"
                End If
            End If
            
            '��������-��������
            If .�������� <> -1 Then
                If .�������� = "~" Then
                    strFilter = strFilter & " And ZL_AgeToDays(C.����)<=[31]"
                Else
                    strFilter = strFilter & " And ZL_AgeToDays(C.����)" & .�������� & "[31]"
                End If
            End If
            
            If .���˿��� <> 0 Then
                strFilter = strFilter & " And B.���˿���ID+0=[11] "
            End If
        
            If .�걾��λ <> "" Then
                strFilter = strFilter & " And instr(B.ҽ������,[12])>0"
            End If
            
            If .������� <> -1 Then
                strFilter = strFilter & " And Nvl(A.�������, 0)=[32]"
            End If
            
            If .���ҽ�� <> "" Then
                strFilter = strFilter & " And H.������=[13] "
            End If
            
            If .���ҽ�� <> "" Then
                strFilter = strFilter & " And H.������=[14] "
            End If
            
            If .Ӱ������ <> "" Then
                strFilter = strFilter & " And H.Ӱ������=[15]"
            End If
            
            If .��鼼ʦ <> "" Then
                strFilter = strFilter & " And H.��鼼ʦ=[16]"
            End If
            
            'Ӱ������������ط�������������ѡ�񣬹��˴��ں����������棬���������е�Ϊ��
            If mintcmdӰ����� > 0 Then
                Dim objControl As CommandBarControl
                
                Set objControl = cbrdock.FindControl(, ID_Ӱ�����)
                For i = 1 To objControl.CommandBar.Controls.Count
                    If objControl.CommandBar.FindControl(, ID_Ӱ����� + i).Checked = True Then
                        strModalitys = strModalitys & "," & objControl.CommandBar.FindControl(, ID_Ӱ����� + i).DescriptionText
                    End If
                Next i
                If strModalitys <> "" Then
                    strFilter = strFilter & " And instr([26],H.Ӱ�����)>0 "
                End If
            Else
                If .Ӱ����� <> "" Then
                    strFilter = strFilter & " And H.Ӱ�����=[17] "
                End If
            End If
            
            
            
            If .��� <> "" Then
                strFilter = strFilter & " And  Instr(H.�������, [18]) > 0 "
            End If
            
            If .������� <> "" Then
                strFilter = strFilter & " And B.ID IN ( Select t.ҽ��id From ����ҽ������ t Where t.����id In " & _
                                                                    " (Select Distinct A.ID  " & _
                                                                        "From ���Ӳ�����¼ A,���Ӳ������� B " & _
                                                                        "Where A.����ʱ��>[9] AND A.Id=B.�ļ�ID  " & _
                                                                            "And B.��������=7 And instr(B.��������,'52;')>0 And instr(B.�����ı�,[19])>0))"
            End If
            
            Dim strSubFilter As String '����PACS�����������
            If .������� <> "" Then
                strSubFilter = " (b.�����ı� ='�������' And Instr(c.�����ı�, [20]) > 0)"
            End If
            
            If .������ <> "" Then
                If strSubFilter = "" Then
                    strSubFilter = " (b.�����ı� ='������' And Instr(c.�����ı�, [21]) > 0)"
                Else
                    strSubFilter = strSubFilter & " or (b.�����ı� ='������' And Instr(c.�����ı�, [21]) > 0)"
                End If
            End If
            
            If .���� <> "" Then
                If strSubFilter = "" Then
                    strSubFilter = " (b.�����ı� ='����' And Instr(c.�����ı�, [22]) > 0)"
                Else
                    strSubFilter = strSubFilter & " or (b.�����ı� ='����' And Instr(c.�����ı�, [22]) > 0)"
                End If
            End If
            
            If strSubFilter <> "" Then
                strSubFilter = " (" & strSubFilter & ")"
                
                strFilter = strFilter & " And B.ID IN ( Select t.ҽ��id From ����ҽ������ t Where t.����id In  " _
                    & " (Select Distinct a.ID From ���Ӳ�����¼ a, ���Ӳ������� b,���Ӳ������� c " _
                    & " Where a.����ʱ�� > [9] And a.Id = b.�ļ�id And b.Id = C.��ID And b.�������� = 3 And c.�������� = 2 And c.��ֹ�� = 0 and " _
                    & strSubFilter & "))"
            End If
           
'            If .������ <> "" Then
'                If .������ = "ȫ��" Then
'
'                ElseIf .������ = "�ѵǼ�" Then
'                    strFilter = strFilter & " And (A.ִ�й��� =0 or A.ִ�й���=1 Or A.ִ�й��� Is Null) "
'                ElseIf .������ = "�ѱ���" Then
'                    strFilter = strFilter & " And (A.ִ�й��� = 2 and h.������ is null) "
'                ElseIf .������ = "�Ѽ��" Then
'                    strFilter = strFilter & " And (A.ִ�й��� = 3 and h.������ is null) "
'                ElseIf .������ = "������" Then
'                    strFilter = strFilter & " And (not h.������� is null) "
'                ElseIf .������ = "������" Then
'                    strFilter = strFilter & " And ((A.ִ�й��� =2 or A.ִ�й���=3) and not h.������ is null and h.������� is null) "
'                ElseIf .������ = "�ѱ���" Then
'                    strFilter = strFilter & " And (A.ִ�й���=4 and h.������ is null) "
'                ElseIf .������ = "�����" Then
'                    strFilter = strFilter & " And (A.ִ�й���=4 and not h.������ is null) "
'                ElseIf .������ = "�����" Then
'                    strFilter = strFilter & " And A.ִ�й���=5 "
'                ElseIf .������ = "�����" Then
'                    strFilter = strFilter & " And A.ִ�й���=6 "
'                End If
'            End If
        End If
        
        '�����˴��ڡ��͡�������ҡ������������������������ʹ��ʱ����������������Ϊ��������
        
        '������Դ (1-����,2-סԺ,3-����,4-���)
        '���������Դ��ѡ���ˣ���ʾ�������в��ˣ�����Ӳ�����Դ�Ĳ�ѯ����
        If mblncmd���� And mblncmdסԺ And mblncmd��� And mblncmd���� Then
        
        Else
            If mblncmd���� Then str��Դ = "1,"
            If mblncmdסԺ Then str��Դ = str��Դ & "2,"
            If mblncmd���� Then str��Դ = str��Դ & "3,"
            If mblncmd��� Then str��Դ = str��Դ & "4,"
            If str��Դ <> "" Then   'str��ԴΪ�գ���ʾû��ѡ���κ���Դ������Ӳ�����Դ�Ĳ�ѯ����
                str��Դ = Mid(str��Դ, 1, Len(str��Դ) - 1)
                strFilter = strFilter & " And Instr([23],B.������Դ)> 0"
            End If
        End If
        
'        If mstrRoom <> "" Then  'ֻ��ʾִ�м䷶Χ�ڵ�
'            If Not mblncmd�Ǽ� Then
'                strFilter = strFilter & " And Instr([24],','|| A.ִ�м� || ',' )>0"
'            Else
'                strFilter = strFilter & " And (Instr([24],','|| A.ִ�м� || ',' )>0 And Nvl(A.ִ�й���,0)>1 OR Nvl(A.ִ�й���,0)<2)"
'            End If
'        End If
    
'        If mblnNoShowCancel Then '����ʾȡ���Ǽǵļ��
'            strFilter = strFilter & " And A.ִ��״̬<>2 "
'        End If
        
        If mblncmd���� Then        'ֻ��ʾ����סԺ��¼
            strFilter = strFilter & vbNewLine & " And (B.������Դ=2 And B.��ҳID=C.סԺ���� Or Nvl(B.������Դ,0)<>2)"
        End If

        '�Ƿ�ѡ����ȫ������
        If mblnAllDepts = True Then
            strFilter = strFilter & " And (Instr( [27], B.ִ�п���ID ) >0  or Instr( [27], B.��������ID ) > 0) "
        Else
            strFilter = strFilter & " AND (B.ִ�п���ID + 0 =[25] or B.��������ID + 0 = [25])"
        End If


        
         
        '������������
        If .�������� <> "" Then
            strFilter = strFilter & " And B.id IN ( Select t.ҽ��id From ����ҽ������ t Where t.����id In " & _
                                                                    " (Select Distinct A.ID " & _
                                                                    " From ���Ӳ�����¼ A,���Ӳ������� B " & _
                                                                    " Where A.����ʱ��>[9] AND A.Id=B.�ļ�ID " & _
                                                                    " And B.��������=2 And instr(B.�����ı�,[28])>0 And B.��ֹ�� = 0)) "
        End If
        
        gstrSQL = "Select /*+ RULE */ Distinct" & vbNewLine & _
                    "       A.ҽ��ID,A.���ͺ�,A.�״�ʱ�� ����ʱ��,A.����ʱ�� ����ʱ��,A.ִ��״̬,nvl(A.ִ�й���,0) ������,A.ִ�м�,A.������� ����," & vbNewLine & _
                    "       B.����ID,B.��ҳID,B.�Һŵ�,B.���˿���ID,Decode(B.������Դ, 1, '��', 2, 'ס', 3, '��', 4, '��') ��Դ,B.ҽ������,B.�걾��λ," & vbNewLine & _
                    "       Nvl(B.������־, 0) ������־, Nvl(B.Ӥ��, 0) Ӥ��,B.����ҽ��,A.NO,C.��ǰ����,C.��ǰ����ID,Decode(B.������Դ,2,C.סԺ��,C.�����) ��ʶ��," & vbNewLine & _
                    "       Nvl(H.����,C.����) ����,H.Ӱ�����,H.����,Nvl(H.�Ա�,C.�Ա�) �Ա�,Nvl(H.����,C.����) ����,H.���,H.����,H.Ӱ������," & vbNewLine & _
                    "       Decode(B.������Դ,3,B.����ҽ��,A.������) �Ǽ���,H.������,H.���淢��,H.����ID, " & vbNewLine & _
                    "       H.�����,H.�Ƿ��ӡ,H.�������,H.��ɫͨ��,H.�����ӡ,H.������,H.������,H.��鼼ʦ,H.�������� ��ͼʱ��," & vbNewLine & _
                    "       H.�������,H.��Ϸ���,H.���UID,A.ִ�в���ID as ִ�п���ID,0 as ת��,F.���� AS ���˿���, " & vbNewLine & _
                    "       C.���￨��,A.NO as ���ݺ�,C.���֤�� " & vbNewLine & _
                    " From ����ҽ������ A,����ҽ����¼ B,������Ϣ C,Ӱ�����¼ H,Ӱ������Ŀ G,���ű� F " & vbNewLine & _
                    " Where B.���ID is NULL And A.ҽ��ID=B.ID And A.ҽ��ID=H.ҽ��ID(+) And A.���ͺ�=H.���ͺ�(+) " & vbNewLine & _
                    " And B.������ĿID=G.������ĿID And B.����ID=C.����ID And B.���˿���id=F.ID "
        gstrSQL = gstrSQL & vbNewLine & strFilter
        
        If mblncmd�ѽ� Xor mblncmdδ�� Then '����ѡ��
            '���ݲ�����Դ�Ĺ��������жϲ�ѯ��Щ���ñ�,
            'ͬʱ��ѯ�����סԺ���ñ��������ĸ���Դ��ѡ���ĸ���Դ����ѡ��ѡ����סԺͬʱ��ѡ���������κ�һ����Դ
            'ֻ��סԺ��������ֻѡ����סԺ
            'ֻ�������������ֻѡ������������죬��������Դ
            
            If mblncmdסԺ = True And mblncmd���� = False And mblncmd���� = False And mblncmd��� = False Then
                'ֻ��ѯסԺ��
                strFilter = "Select Distinct NO From סԺ���ü�¼ D Where A.NO = D.NO And A.��¼���� = D.��¼���� And D.��¼״̬ = 1"
            ElseIf mblncmdסԺ = False And (mblncmd���� = True Or mblncmd���� = True Or mblncmd��� = True) Then
                'ֻ�������
                strFilter = "Select Distinct NO From ������ü�¼ E Where A.NO = E.NO And A.��¼���� = E.��¼���� And E.��¼״̬ = 1"
            Else    '�������ͬʱ��������
                strFilter = "Select Distinct NO From סԺ���ü�¼ D Where A.NO = D.NO And A.��¼���� = D.��¼���� And D.��¼״̬ = 1" & vbNewLine & _
                            "Union" & vbNewLine & _
                            "Select Distinct NO From ������ü�¼ E Where A.NO = E.NO And A.��¼���� = E.��¼���� And E.��¼״̬ = 1"
            End If
            
            gstrSQL = gstrSQL & vbNewLine & IIf(mblncmd�ѽ�, " And Exists ", " And Not Exists") & "(" & strFilter & ")"
        End If
        
        '��ʹ�ü��Ų���ʱһ���Ǳ������ģ�Ӱ�����¼���м�¼����ʱȡ�������ӱ���ȫ��ɨ��
        'ʹ�òɼ�ʱ����ˣ�Ӱ�����¼���м�¼
        If .���� <> 0 Or (blnUseTime = True And SQLCondition.ʱ������ = 3) Then
            gstrSQL = Replace(Replace(gstrSQL, "H.ҽ��ID(+)", "H.ҽ��ID"), "H.���ͺ�(+)", "H.���ͺ�")
        End If
        
        '���������ת����Ҫ�����󱸱�
        If mblnMoved Then
            strSQLBak = gstrSQL
            strSQLBak = Replace(strSQLBak, "����ҽ����¼", "H����ҽ����¼")
            strSQLBak = Replace(strSQLBak, "����ҽ������", "H����ҽ������")
            strSQLBak = Replace(strSQLBak, "Ӱ�����¼", "HӰ�����¼")
            strSQLBak = Replace(strSQLBak, "���Ӳ�����¼", "H���Ӳ�����¼")
            strSQLBak = Replace(strSQLBak, "���Ӳ�������", "H���Ӳ�������")
            strSQLBak = Replace(strSQLBak, "0 as ת��", "1 as ת��")
            strSQL = strSQL & " Union ALL " & strSQLBak
        End If
        gstrSQL = "Select * From (" & vbNewLine & gstrSQL & vbNewLine & ") Order by ������,����ʱ��,����ʱ��"
    
        Set rsList = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����б�", .�����, .סԺ��, .���￨, .����, .���֤, .IC��, .���ݺ�, _
                                            .����, .��ʼʱ��, .����ʱ��, .���˿���, .�걾��λ, .���ҽ��, .���ҽ��, .Ӱ������, _
                                            .��鼼ʦ, .Ӱ�����, .���, .�������, .�������, .������, .����, str��Դ, "", _
                                           glngDeptId, strModalitys, mstrҽ����������IDs, .��������, .�Ա�, .��ʼ����, .��������, .�������)
    End With

    strFilter = ""
    If mblncmd�Ǽ� Then strFilter = "������=0 or ������=1 or "
    If mblncmd���� Then strFilter = IIf(strFilter <> "", strFilter & "������=2 or ", "������=2 or ")
    If mblncmd��� Then strFilter = IIf(strFilter <> "", strFilter & "������=3 or ", "������=3 or ")
    If mblncmd���� Then strFilter = IIf(strFilter <> "", strFilter & "������=4 or ", "������=4 or ")
    If mblncmd��� Then strFilter = IIf(strFilter <> "", strFilter & "������=5 or ", "������=5 or ")
    If mblncmd��� Then strFilter = IIf(strFilter <> "", strFilter & "������=6 or ", "������=6 or ")
    
    If mblncmd�Ǽ� And mblncmd���� And mblncmd��� And mblncmd���� Then ' And mblncmd��� And mblncmd��� Then
        strFilter = ""
    End If

    If strFilter <> "" Then
        strFilter = Mid(strFilter, 1, Len(strFilter) - 4)
        rsList.Filter = strFilter
    End If
    
    Call FillList(rsList)
    mblnvsRefresh = False
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub FillList(ByVal rsTemp As ADODB.Recordset)
Dim rsBaby As ADODB.Recordset
    On Error GoTo errHandle
    Call InitList
    If rsTemp.EOF Then stbThis.Panels(2).Text = "û���ҵ��κ�ƥ��ļ�¼": Exit Sub
    
    With vsList
        .Rows = rsTemp.RecordCount + 1
        Do Until rsTemp.EOF
            .Row = rsTemp.AbsolutePosition
            .Cell(flexcpData, .Row, GetCN("����")) = Val(rsTemp!������־)
            If rsTemp!������־ <> 0 Then
                Set .Cell(flexcpPicture, .Row, GetCN("����")) = Imglist.ListImages("����").Picture
            End If
            If rsTemp!��Դ = "ס" Then
                Set .Cell(flexcpPicture, .Row, GetCN("��Դ")) = Imglist.ListImages("סԺ").Picture
            End If
            .TextMatrix(.Row, GetCN("��Դ")) = rsTemp!��Դ
            .Cell(flexcpData, .Row, GetCN("��Դ")) = Decode(rsTemp!��Դ, "��", 1, "ס", 2, "��", 3, 4)
            
            If Nvl(rsTemp!����, 0) <> 0 Then
                Set .Cell(flexcpPicture, .Row, GetCN("����")) = Imglist.ListImages("����").Picture
            End If
            
            If Nvl(rsTemp!��ɫͨ��, 0) <> 0 Then
                Set .Cell(flexcpPicture, .Row, GetCN("����")) = Imglist.ListImages("��ɫͨ��").Picture
            End If
            
            If Nvl(rsTemp!���uid) <> "" Then
                Set .Cell(flexcpPicture, .Row, GetCN("����")) = Imglist.ListImages("Ӱ��").Picture
            End If
            
            .TextMatrix(.Row, GetCN("����")) = Nvl(rsTemp!Ӱ������)
            .TextMatrix(.Row, GetCN("����")) = Nvl(rsTemp!����)
            .TextMatrix(.Row, GetCN("����")) = Nvl(rsTemp!����)
            .TextMatrix(.Row, GetCN("������")) = IIf(rsTemp!ִ��״̬ = 2, "�Ѿܾ�", Decode(Nvl(rsTemp!������, 0), 0, "�ѵǼ�", 1, "�ѵǼ�", _
                                                                                        2, IIf(Nvl(rsTemp!�������) <> "", "������", _
                                                                                                IIf(Nvl(rsTemp!������) = "", "�ѱ���", "������")), _
                                                                                        3, IIf(Nvl(rsTemp!�������) <> "", "������", _
                                                                                                IIf(Nvl(rsTemp!������) = "", "�Ѽ��", "������")), _
                                                                                        4, IIf(Nvl(rsTemp!�������) <> "", "������", _
                                                                                                IIf(Nvl(rsTemp!������) <> "", "�����", "�ѱ���")), _
                                                                                        5, "�����", "�����"))
            .TextMatrix(.Row, GetCN("�Ա�")) = Nvl(rsTemp!�Ա�)
            .TextMatrix(.Row, GetCN("����")) = Nvl(rsTemp!����)
            If InStr(Nvl(rsTemp!ҽ������), ":") > 0 Then '�µ�ģʽ������ҽ����������Ϣ�� ����,ִ�б��:��λ(����,����),��λ---
                .TextMatrix(.Row, GetCN("ҽ������")) = Split(rsTemp!ҽ������, ":")(0)
                .TextMatrix(.Row, GetCN("��λ����")) = Split(rsTemp!ҽ������, ":")(1)
            Else
                .TextMatrix(.Row, GetCN("ҽ������")) = Nvl(rsTemp!ҽ������)
            End If
            .TextMatrix(.Row, GetCN("ִ�м�")) = Nvl(rsTemp!ִ�м�)
            .TextMatrix(.Row, GetCN("����ʱ��")) = Nvl(rsTemp!����ʱ��)
            .TextMatrix(.Row, GetCN("����ʱ��")) = Nvl(rsTemp!����ʱ��)
            .TextMatrix(.Row, GetCN("����ҽ��")) = Nvl(rsTemp!����ҽ��)
            .TextMatrix(.Row, GetCN("���")) = Nvl(rsTemp!���)
            .TextMatrix(.Row, GetCN("����")) = Nvl(rsTemp!����)
            .TextMatrix(.Row, GetCN("Ӥ��")) = Nvl(rsTemp!Ӥ��)
            .TextMatrix(.Row, GetCN("�Ǽ���")) = Nvl(rsTemp!�Ǽ���)
            .TextMatrix(.Row, GetCN("������")) = Nvl(rsTemp!������)
            .TextMatrix(.Row, GetCN("�����")) = Nvl(rsTemp!�����)
            .TextMatrix(.Row, GetCN("��ӡ��Ƭ")) = IIf(Nvl(rsTemp!�Ƿ��ӡ) = 1, "�Ѵ�ӡ", "δ��ӡ")
            .TextMatrix(.Row, GetCN("�������")) = Nvl(rsTemp!�������)
            .TextMatrix(.Row, GetCN("��ɫͨ��")) = Nvl(rsTemp!��ɫͨ��)
            .TextMatrix(.Row, GetCN("�����ӡ")) = IIf(Nvl(rsTemp!�����ӡ) = 1, "�Ѵ�ӡ", "δ��ӡ")
            .TextMatrix(.Row, GetCN("������")) = Nvl(rsTemp!������)
            .TextMatrix(.Row, GetCN("������")) = Nvl(rsTemp!������)
            .TextMatrix(.Row, GetCN("��鼼ʦ")) = Nvl(rsTemp!��鼼ʦ)
            .TextMatrix(.Row, GetCN("��ͼʱ��")) = Nvl(rsTemp!��ͼʱ��)
            .TextMatrix(.Row, GetCN("Ӱ�����")) = Nvl(rsTemp!Ӱ�����)
            .TextMatrix(.Row, GetCN("����ID")) = Nvl(rsTemp!����ID, 0)
            .TextMatrix(.Row, GetCN("��ҳID")) = Nvl(rsTemp!��ҳID, 0)
            .TextMatrix(.Row, GetCN("�Һŵ�")) = Nvl(rsTemp!�Һŵ�)
            .TextMatrix(.Row, GetCN("���˿���ID")) = Nvl(rsTemp!���˿���ID, 0)
            .TextMatrix(.Row, GetCN("ҽ��ID")) = Nvl(rsTemp!ҽ��id)
            .TextMatrix(.Row, GetCN("���ͺ�")) = Nvl(rsTemp!���ͺ�)
            .TextMatrix(.Row, GetCN("���UID")) = Nvl(rsTemp!���uid)
            .TextMatrix(.Row, GetCN("���״̬")) = Nvl(rsTemp!������)
            .TextMatrix(.Row, GetCN("�������")) = Nvl(rsTemp!�������)
            .TextMatrix(.Row, GetCN("NO")) = Nvl(rsTemp!no)
            .TextMatrix(.Row, GetCN("ת��")) = Nvl(rsTemp!ת��)
            .TextMatrix(.Row, GetCN("����")) = Nvl(rsTemp!��ǰ����)
            .TextMatrix(.Row, GetCN("��ǰ����ID")) = Nvl(rsTemp!��ǰ����ID, 0)
            .TextMatrix(.Row, GetCN("��ʶ��")) = Nvl(rsTemp!��ʶ��)
            .TextMatrix(.Row, GetCN("���淢��")) = IIf(Nvl(rsTemp!���淢��, 0) = 0, "δ����", "�ѷ���")
            .TextMatrix(.Row, GetCN("��Ϸ���")) = Nvl(rsTemp!��Ϸ���)
            .TextMatrix(.Row, GetCN("ִ�п���ID")) = Nvl(rsTemp!ִ�п���ID)
            .TextMatrix(.Row, GetCN("����ID")) = Nvl(rsTemp!����ID, 0)
            .TextMatrix(.Row, GetCN("���˿���")) = Nvl(rsTemp!���˿���)
            .TextMatrix(.Row, GetCN("���￨��")) = Nvl(rsTemp!���￨��)
            .TextMatrix(.Row, GetCN("���ݺ�")) = Nvl(rsTemp!���ݺ�)
            .TextMatrix(.Row, GetCN("���֤��")) = Nvl(rsTemp!���֤��)
            
            If Nvl(rsTemp!Ӥ��) <> 0 Then
                gstrSQL = "Select Nvl(A.Ӥ������, B.���� || '֮��' || Trim(To_Char(A.���, '9'))) As Ӥ������, Ӥ���Ա�, ����ʱ��" & vbNewLine & _
                            "From ������������¼ A, ������Ϣ B" & vbNewLine & _
                            "Where A.����id = [1] And A.��ҳid = [2] And A.����id = B.����id And A.��� = [3]"

                Set rsBaby = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡӤ����Ϣ", CLng(rsTemp!����ID), CLng(Nvl(rsTemp!��ҳID, 0)), CLng(rsTemp!Ӥ��))
                If Not rsBaby.EOF Then
                    .TextMatrix(.Row, GetCN("����")) = rsBaby!Ӥ������
                    .TextMatrix(.Row, GetCN("�Ա�")) = Nvl(rsBaby!Ӥ���Ա�)
                    .TextMatrix(.Row, GetCN("����")) = Nvl(rsBaby!����ʱ��)
                End If
            End If
            
            If .TextMatrix(.Row, GetCN("������")) = "�Ѿܾ�" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor�Ѿܾ�
            If .TextMatrix(.Row, GetCN("������")) = "�����" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor�����
            If .TextMatrix(.Row, GetCN("������")) = "�ѱ���" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor�ѱ���
            If .TextMatrix(.Row, GetCN("������")) = "�ѵǼ�" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor�ѵǼ�
            If .TextMatrix(.Row, GetCN("������")) = "�Ѽ��" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor�Ѽ��
            If .TextMatrix(.Row, GetCN("������")) = "�����" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor�����
            If .TextMatrix(.Row, GetCN("������")) = "������" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor������
            If .TextMatrix(.Row, GetCN("������")) = "������" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor������
            If .TextMatrix(.Row, GetCN("������")) = "�����" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor�����
            If .TextMatrix(.Row, GetCN("������")) = "�ѱ���" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor�ѱ���
            
            rsTemp.MoveNext
        Loop
    End With
    
    '�ָ�����
    If mlngSortCol <> 0 And mintSortOrder <> 0 Then
        If mlngSortCol < vsList.Cols Then
            vsList.Col = mlngSortCol
            vsList.Sort = mintSortOrder
        End If
    End If
    
    stbThis.Panels(2).Text = "�� " & vsList.Rows - 1 & " ����¼": stbThis.Panels(2).Alignment = sbrCenter
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub PicWindow_Resize()
    On Error Resume Next
    With picInfo
        .Top = 0
        .Left = 0
        .Width = PicWindow.ScaleWidth
    End With
        
    With TabWindow
        .Top = picInfo.ScaleHeight
        .Left = 0
        .Width = PicWindow.ScaleWidth
        .Height = PicWindow.ScaleHeight - picInfo.ScaleHeight
    End With
End Sub


Private Sub TimerRefresh_Timer()
    'ˢ�²����б�
    Call RefreshList
End Sub


Private Sub SeekNextPati(ByVal blnFirst As Boolean, ByVal strCardName As String, _
    ByVal strCardText As String, ByVal lngPatientID As Long)
'------------------------------------------------
'���ܣ��ڲ����б��ж�λָ���ļ�¼
'������ blnFirst -- �Ƿ��һ�β���
'���أ��ޣ�ֱ���ڲ����б��ж�λ
'------------------------------------------------
    Dim blnOk As Boolean, lngCount As Long, intB As Integer
    Dim lngRow As Long

    '���û�м�¼�����˳�
    If Val(vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID"))) = 0 Then Exit Sub

    intB = 0
    
    On Error GoTo err
    
    If Not blnFirst Then
        intB = vsList.Row + 1
        If intB >= vsList.Rows Then intB = 1
    End If

    blnOk = False
    For lngCount = intB To vsList.Rows - 1 '�ڵ�ǰ״̬�в���
        Select Case strCardName
            Case "��ʶ��", "סԺ��", "�����"
                If UCase(Nvl(vsList.TextMatrix(lngCount, GetCN("��ʶ��")), 0)) Like UCase(strCardText) & "*" Then blnOk = True
            Case "���￨", "IC����", "IC��"
                If UCase(Nvl(vsList.TextMatrix(lngCount, GetCN("���￨��")), 0)) Like UCase(strCardText) & "*" Then blnOk = True
            Case "���ݺ�"
                If UCase(Nvl(vsList.TextMatrix(lngCount, GetCN("NO")), 0)) Like UCase(strCardText) & "*" Then blnOk = True
            Case "����"
                If UCase(Nvl(vsList.TextMatrix(lngCount, GetCN("����")), 0)) Like UCase(strCardText) & "*" Then blnOk = True
            Case "����", "�� ��", "��  ��", "��   ��"
                If UCase(Nvl(vsList.TextMatrix(lngCount, GetCN("����")), "")) Like UCase(strCardText) & "*" Then blnOk = True
                If zlCommFun.SpellCode(Nvl(vsList.TextMatrix(lngCount, GetCN("����")), "")) Like UCase(strCardText) & "*" Then blnOk = True
            Case "���֤��", "���֤"
                If UCase(Nvl(vsList.TextMatrix(lngCount, GetCN("���֤��")), 0)) Like UCase(strCardText) & "*" Then blnOk = True
            Case Else
                If UCase(Nvl(vsList.TextMatrix(lngCount, GetCN("����ID")), 0)) = UCase(lngPatientID) Then blnOk = True
        End Select

        If blnOk Then
            ucLocate.Tag = ucLocate.CardText
            On Error Resume Next
            '���㵱ǰ�кͶ���֮��Ĳ��
            lngRow = Abs(vsList.Row - vsList.TopRow)

            vsList.Row = lngCount
            vsList.TopRow = vsList.Row - lngRow

            Exit Sub
        End If
    Next
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub sub3DProcess(strCommand As String, strImageDir As String)
    Dim str3DCommand As String
    
    '��֯��ά�ؽ����
    str3DCommand = mstr3DExeDir & " " & mstr3DPara & " " & strCommand & " " & strImageDir
    On Error Resume Next
    Shell str3DCommand
End Sub

Private Sub sub��ά�ؽ�(strCommand As String)
    Dim strImageDir As String
    
    If TabWindow.Selected.Tag <> "Ӱ��ͼ��" Then '��ˢ��ͼ������
        Call mfrmPACSImg.zlRefresh(vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")), vsList.TextMatrix(vsList.Row, GetCN("���ͺ�")), mstrPrivs, vsList.TextMatrix(vsList.Row, GetCN("ת��")) = 1)
    End If
    
    '��֯��ά�ؽ���Ҫ��ͼ��
    strImageDir = mfrmPACSImg.ZLfun3DImgProcess
    If strImageDir <> "" Then
        Call sub3DProcess(strCommand, strImageDir)
    End If
End Sub


Private Sub ucFilter_OnClick(ByVal strCardName As String, ByVal strCardText As String, ByVal lngKindId As Long, ByVal lngCardLen As Long, ByVal lngSwipingType As Long, ByVal blnIsPwdInput As Boolean)
'�������ÿؼ�ʱ�������
On Error GoTo errHandle
    Dim lngPatientID As Long
    
    '���Ϊ1�������
    If lngSwipingType = 1 Then ucFilter.CardText = ucFilter.ReadCard(lngPatientID)
    
    Call subRefreshFilterCondition(strCardName, ucFilter.CardText, lngPatientID)
    Call RefreshList
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ucFilter_OnRead(ByVal strCardName As String, ByVal strCardText As String, ByVal lngPatientID As Long)
'��ʼ��������
On Error GoTo errHandle
    Call subRefreshFilterCondition(strCardName, strCardText, lngPatientID)
    Call RefreshList
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ucFilter_OnResize()
On Error Resume Next
    Call cbrdock.RecalcLayout
err.Clear
End Sub

Private Sub ucLocate_OnClick(ByVal strCardName As String, ByVal strCardText As String, ByVal lngKindId As Long, ByVal lngCardLen As Long, ByVal lngSwipingType As Long, ByVal blnIsPwdInput As Boolean)
'�������ÿؼ�ʱ�������
On Error GoTo errHandle
    Dim lngPatientID As Long
    
    '���Ϊ1�������
    If lngSwipingType = 1 Then ucLocate.CardText = ucLocate.ReadCard(lngPatientID)
    
    Call SeekNextPati(ucLocate.Tag <> ucLocate.CardText, strCardName, ucLocate.CardText, lngPatientID)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ucLocate_OnRead(ByVal strCardName As String, ByVal strCardText As String, ByVal lngPatientID As Long)

    
On Error GoTo errHandle
     Call SeekNextPati(ucLocate.Tag <> ucLocate.CardText, strCardName, strCardText, lngPatientID)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub vsList_AfterMoveColumn(ByVal Col As Long, Position As Long)
Dim i As Integer, strCol As String
    For i = 0 To vsList.Cols - 1
        strCol = strCol & "|" & vsList.Cell(flexcpData, 0, i) & ";" & vsList.ColWidth(i)
    Next
    mstrCol = Mid(strCol, 2)
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'����: ��ʾ���˿�Ƭ��ť
    If vsList.TextMatrix(NewRow, GetCN("ҽ��ID")) = "" Then
        cmdInfo.Visible = False
    Else
        If vsList.LeftCol > GetCN("����") Then
            cmdInfo.Visible = False
        Else
            cmdInfo.Left = vsList.Cell(flexcpLeft, NewRow, GetCN("����")) + vsList.Cell(flexcpWidth, NewRow, GetCN("����")) - cmdInfo.Width - 15
            cmdInfo.Top = vsList.Cell(flexcpTop, vsList.Row, GetCN("����")) + 15
            cmdInfo.Visible = True
        End If
    End If
End Sub
Private Sub vsList_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
'����:��ʾ���˿�Ƭ��ť
    If vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")) = "" Then
        cmdInfo.Visible = False
    Else
        If NewLeftCol > GetCN("����") Then
            cmdInfo.Visible = False
        Else
            cmdInfo.Left = vsList.Cell(flexcpLeft, vsList.Row, GetCN("����")) + vsList.Cell(flexcpWidth, vsList.Row, GetCN("����")) - cmdInfo.Width - 15
            cmdInfo.Top = vsList.Cell(flexcpTop, vsList.Row, GetCN("����")) + 15
            cmdInfo.Visible = True
        End If
    End If
End Sub

Private Sub vsList_AfterSort(ByVal Col As Long, Order As Integer)
    mlngSortCol = Col
    mintSortOrder = Order
End Sub

Private Sub vsList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
'����:��ʾ���˿�Ƭ��ť
    If vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")) = "" Then
        cmdInfo.Visible = False
    Else
        If vsList.LeftCol > GetCN("����") Then
            cmdInfo.Visible = False
        Else
            cmdInfo.Left = vsList.Cell(flexcpLeft, vsList.Row, GetCN("����")) + vsList.Cell(flexcpWidth, vsList.Row, GetCN("����")) - cmdInfo.Width - 15
            cmdInfo.Top = vsList.Cell(flexcpTop, vsList.Row, GetCN("����")) + 15
            cmdInfo.Visible = True
        End If
    End If
    
    Dim i As Integer, strCol As String
    For i = 0 To vsList.Cols - 1 '�ݴ������п�����ر�ʱ����ע���
        strCol = strCol & "|" & vsList.Cell(flexcpData, 0, i) & ";" & vsList.ColWidth(i)
    Next
    mstrCol = Mid(strCol, 2)
End Sub

Private Sub vsList_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < GetCN("����") Then Cancel = True
End Sub

Private Sub vsList_DblClick()

    'û���治�ܴ�ӡ��Ԥ��
    If vsList.TextMatrix(vsList.Row, GetCN("������")) = "" Then
        MsgBoxD Me, "��ǰ����û�м�鱨�棬���ܲ��������飡", vbInformation, gstrSysName
        Exit Sub
    End If
            
            
    If vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")) <> "" Then
        mlngAdviceID = vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID"))
        Call OpenReportPreview(mlngAdviceID)
    End If
    
End Sub

Private Sub vsList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Dim control As CommandBarControl, Menucontrol As CommandBarControl
        Dim Popup As CommandBar
        Set Popup = cbrMain.Add("�Ҽ��˵�", xtpBarPopup)
        For Each Menucontrol In cbrMain.ActiveMenuBar.Controls
'            If Menucontrol.Parent.BarID = conMenu_ManagePopup Then
            If (Menucontrol.ID <> conMenu_FilePopup And Menucontrol.ID <> conMenu_ToolPopup _
                And Menucontrol.ID <> conMenu_ViewPopup And Menucontrol.ID <> conMenu_HelpPopup) And Menucontrol.Type = xtpControlPopup Then
                
                For Each control In Menucontrol.CommandBar.Controls
                    control.Copy Popup
                Next
            End If
        Next
        Popup.ShowPopup
    End If
End Sub

Private Sub vsList_RowColChange()
    On Error GoTo errHandle
'    mblnIsHistory = False

    If mlngAdviceID = Val(vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID"))) Then Exit Sub

    mlngAdviceID = Val(vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID")))
    mstrHStudyUID = vsList.TextMatrix(vsList.Row, GetCN("���UID"))
    
    If Val(vsList.TextMatrix(vsList.Row, GetCN("ҽ��ID"))) = 0 Then  '�޼�¼ʱ����
        Call RefreshTabWindow(0, True)
        cboTimes.Clear
        txtAppend = ""
        lbl������Ϣ.Caption = "��  ��:" & Space(12) & "��  ��:" & Space(13) & "��  ��:" & Space(10) & "��ʶ��:" & Space(12) & "��  ��:" & Space(10)
        lbl�����Ϣ.Caption = "����:" & Space(12) & "���˿���:" & Space(11) & "����ҽ��:" & Space(8) & "�����Ŀ:"
        lblCash.Visible = False
    Else
        
        Call FillHistory '������μ���¼
        Call FillTxtInfor '������Ϸ����˻�����Ϣ
        Call FillTxtAppend '������½�ҽ������

        Call RefreshTabWindow

    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub FillTxtInfor(Optional lngAdviceID As Long = 0)
'������Ϸ����˻�����Ϣ
Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    
    With vsList
        lbl������Ϣ.Caption = "��  ��:" & Rpad(.TextMatrix(.Row, GetCN("����")), 12, " ") & "��  ��:" & Rpad(.TextMatrix(.Row, GetCN("�Ա�")), 13, " ") & _
                          "��  ��:" & Rpad(.TextMatrix(.Row, GetCN("����")), 10, " ") & "��ʶ��:" & Rpad(.TextMatrix(.Row, GetCN("��ʶ��")), 12, " ") & _
                          "��  ��:" & Rpad(.TextMatrix(.Row, GetCN("����")) & "", 10, " ")
                          
        If lngAdviceID = 0 Then '---------------------------�����μ��ֱ�����б��м�¼���
            gstrSQL = "Select ���� From ���ű� Where ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���˿���", CLng(.TextMatrix(.Row, GetCN("���˿���ID"))))
            lbl�����Ϣ.Caption = "����:" & Rpad(.TextMatrix(.Row, GetCN("����")), 12, " ") & "���˿���:" & Rpad(rsTemp!����, 11, " ") & _
                                  "����ҽ��:" & Rpad(.TextMatrix(.Row, GetCN("����ҽ��")), 8, " ") & "�����Ŀ:" & .TextMatrix(.Row, GetCN("ҽ������"))
            If .TextMatrix(.Row, GetCN("��λ����")) <> "" Then lbl�����Ϣ.Caption = lbl�����Ϣ.Caption & "(" & .TextMatrix(.Row, GetCN("��λ����")) & ")"
            
            mlngHSendNo = Nvl(.TextMatrix(.Row, GetCN("���ͺ�")), 0)
            mstrHStudyUID = Nvl(.TextMatrix(.Row, GetCN("���UID")))
            mlngExecuteStep = Decode(.TextMatrix(.Row, GetCN("������")), "�����", 6, "�����", 5, 0)
            mblnHMoved = IIf(.TextMatrix(.Row, GetCN("ת��")) = 1, True, False)
            
            lblCash.Caption = "��"
            lblCash.Visible = False
            lblCash.Visible = CheckChargeState(.TextMatrix(.Row, GetCN("ҽ��ID")), CLng(Decode(.TextMatrix(.Row, GetCN("��Դ")), "��", 1, "ס", 2, "��", 3, 4))) = 1
        Else
            Dim strSQLBak As String
            gstrSQL = "Select A.ID, A.���˿���id, A.����ҽ��,A.������Դ, A.ҽ������, Nvl(A.Ӥ��, 0) Ӥ��,A.����id, " & vbNewLine & _
                        " A.��ҳid, A.�Һŵ�, B.����, B.���uid, C.����, D.ִ�й���, D.���ͺ�,D.ִ��״̬,0 as ת��,A.ִ�п���ID " & vbNewLine & _
                        "From ����ҽ����¼ A, Ӱ�����¼ B, ���ű� C, ����ҽ������ D" & vbNewLine & _
                        "Where A.ID = [1] And A.ID = B.ҽ��id And A.���˿���id = C.ID And A.ID = D.ҽ��id"
            strSQLBak = gstrSQL
            strSQLBak = Replace(strSQLBak, "����ҽ����¼", "H����ҽ����¼")
            strSQLBak = Replace(strSQLBak, "����ҽ������", "H����ҽ������")
            strSQLBak = Replace(strSQLBak, "Ӱ�����¼", "HӰ�����¼")
            strSQLBak = Replace(strSQLBak, "0 as ת��", "1 as ת��")
            gstrSQL = gstrSQL & vbNewLine & " Union ALL " & strSQLBak
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�����μ�¼��Ϣ", lngAdviceID)
            If Not rsTemp.EOF Then

                mlngHSendNo = Nvl(rsTemp!���ͺ�, 0)
                mstrHStudyUID = Nvl(rsTemp!���uid)
                mlngExecuteStep = Nvl(rsTemp!ִ�й���, 0)
                mblnHMoved = IIf(rsTemp!ת�� = 1, True, False)
                
                fraInfo.Tag = rsTemp!����ID & "|" & rsTemp!��ҳID & "|" & rsTemp!ID & "|" & rsTemp!���ͺ� _
                            & "|" & rsTemp!���˿���ID & "|" & rsTemp!�Һŵ� & "|" & Nvl(rsTemp!������Դ, 3) _
                            & "|" & rsTemp!���uid & "|" & rsTemp!ת�� & "|" & rsTemp!ִ��״̬ & "|" & rsTemp!ִ�п���ID
                            
                lbl�����Ϣ.Caption = "����:" & Rpad(Nvl(rsTemp!����), 12, " ") & "���˿���:" & Rpad(rsTemp!����, 11, " ") & _
                                      "����ҽ��:" & Rpad(rsTemp!����ҽ��, 8, " ") & "�����Ŀ:" & rsTemp!ҽ������
                If rsTemp!Ӥ�� <> 0 Then
                    Dim lngBaby As Integer, lngPatID As Long, lngPageID As Long
                    lngBaby = rsTemp!Ӥ��: lngPatID = rsTemp!����ID: lngPageID = Nvl(rsTemp!��ҳID, 0)
                    gstrSQL = "Select Nvl(A.Ӥ������, B.���� || '֮��' || Trim(To_Char(A.���, '9'))) As Ӥ������, Ӥ���Ա�, ����ʱ��" & vbNewLine & _
                            "From ������������¼ A, ������Ϣ B" & vbNewLine & _
                            "Where A.����id = [1] And A.��ҳid = [2] And A.����id = B.����id And A.��� = [3]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡӤ����Ϣ", lngPatID, lngPageID, lngBaby)
                    If Not rsTemp.EOF Then
                        lbl������Ϣ.Caption = "��  ��:" & Rpad(rsTemp!Ӥ������, 12, " ") & "��  ��:" & Rpad(rsTemp!Ӥ���Ա�, 13, " ") & _
                                            "��  ��:" & Rpad(rsTemp!����ʱ��, 10, " ") & "��ʶ��:" & Rpad(.TextMatrix(.Row, GetCN("��ʶ��")), 12, " ") & _
                                            "��  ��:" & Rpad(.TextMatrix(.Row, GetCN("����")) & "", 10, " ")
                    End If
                End If
            Else
                lbl�����Ϣ.Caption = "����:" & Space(12) & "���˿���:" & Space(11) & "����ҽ��:" & Space(8) & "�����Ŀ:"
            End If
            lblCash.Caption = "��": lblCash.Visible = True
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub FillTxtAppend(Optional lngAdviceIDtmp As Long = 0)
'������½�ҽ������
Dim lngAdviceID As Long, strAppend As String, rsTemp As ADODB.Recordset, i As Integer
    On Error GoTo errHandle
    With vsList
        If lngAdviceIDtmp = 0 Then
            lngAdviceID = Val(.TextMatrix(.Row, GetCN("ҽ��ID")))
        Else
            lngAdviceID = lngAdviceIDtmp
        End If
        
        If lngAdviceIDtmp = 0 Then '-------------------------------------------�б�ѡ�����
            If .TextMatrix(.Row, GetCN("��λ����")) <> "" Then
                For i = 0 To UBound(Split(.TextMatrix(.Row, GetCN("��λ����")), "),"))
                    If i = 0 Then
                        txtAppend = "��鲿λ:" & vbCrLf & Space(2) & "1:" & Split(.TextMatrix(.Row, GetCN("��λ����")), "),")(i) & ")"
                    Else
                        txtAppend = txtAppend & vbCrLf & Space(2) & i + 1 & ":" & Split(.TextMatrix(.Row, GetCN("��λ����")), "),")(i) & ")"
                    End If
                Next
                If Trim(txtAppend) <> "" Then txtAppend = Mid(txtAppend, 1, Len(txtAppend) - 1) 'ȡ����������
            Else
                txtAppend = "��鲿λ:" & .TextMatrix(.Row, GetCN("ҽ������"))
            End If
            gstrSQL = "Select ��Ŀ,���� From ����ҽ������ Where ҽ��ID=[1] Order By ����"
            If .TextMatrix(.Row, GetCN("ת��")) = 1 Then gstrSQL = Replace(gstrSQL, "����ҽ������", "H����ҽ������")
        Else                    '-------------------------------------------���μ�¼ѡ�����
            Dim strTemp As String
            txtAppend = ""
            strTemp = Mid(lbl�����Ϣ.Caption, InStr(lbl�����Ϣ.Caption, "�����Ŀ:") + 5)
            If strTemp <> "" Then
                If InStr(strTemp, ":") > 0 Then
                    strTemp = Split(strTemp, ":")(1)
                    For i = 0 To UBound(Split(strTemp, "),"))
                        If i = 0 Then
                            txtAppend = "��鲿λ:" & vbCrLf & Space(2) & "1:" & Split(strTemp, "),")(i) & ")"
                        Else
                            txtAppend = txtAppend & vbCrLf & Space(2) & i + 1 & ":" & Split(strTemp, "),")(i) & ")"
                        End If
                    Next
                    If Trim(txtAppend) <> "" Then txtAppend = Mid(txtAppend, 1, Len(txtAppend) - 1) 'ȡ����������
                Else
                    txtAppend = strTemp
                End If
            End If
            gstrSQL = "Select ��Ŀ,���� From ����ҽ������ Where ҽ��ID=[1] Order By ����" '�������μ�¼�Ƿ�ת���жϲ���ʷ��
            If Split(fraInfo.Tag, "|")(8) = 1 Then gstrSQL = Replace(gstrSQL, "����ҽ������", "H����ҽ������")
        End If
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���˸���", lngAdviceID)
        Do Until rsTemp.EOF
            strAppend = strAppend & rsTemp!��Ŀ & ":" & Nvl(rsTemp!����) & vbCrLf
            rsTemp.MoveNext
        Loop
        
        txtAppend = txtAppend & vbCrLf & vbCrLf & strAppend
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub FillHistory()
'������μ���¼
Dim rsTemp As ADODB.Recordset, strTemp As String
    On Error GoTo errHandle
    With vsList
        cboTimes.Tag = "" 'cbotime����ʱ�õ�������������"������Ŀ"ʱ��������"���cbotimes"����
        gstrSQL = "Select A.ID ҽ��ID,A.����ʱ��  ����ʱ��,A.ҽ������ " & _
                   " From ����ҽ����¼ A,����ҽ������ B,Ӱ�����¼ C" & _
                   " Where A.����id = [1] And A.���id Is Null And B.ҽ��ID=A.ID " & _
                   "" & IIf(.TextMatrix(.Row, GetCN("������")) = "�Ѿܾ�", "", " And B.ִ��״̬<>2 ") & _
                   " AND A.ID=C.ҽ��ID"

        gstrSQL = gstrSQL & " And (A.ִ�п���id+0 =[2] or A.��������ID+0=[2])"


        '���ù������ˣ��Ų�ѯ����ID
        If mblnRelatingPatient = True And .TextMatrix(.Row, GetCN("����ID")) <> 0 Then
            gstrSQL = gstrSQL & " union select A.ID ҽ��ID,A.����ʱ��  ����ʱ��,A.ҽ������ " & _
                    " From ����ҽ����¼ A " & _
                    " Where A.id in (Select ҽ��ID from Ӱ�����¼ Where ����ID =[4]) "
        End If

        strTemp = Replace(gstrSQL, "����ҽ����¼", "H����ҽ����¼")
        strTemp = Replace(strTemp, "����ҽ������", "H����ҽ������")
        strTemp = Replace(strTemp, "Ӱ�����¼", "HӰ�����¼")
        gstrSQL = gstrSQL & vbNewLine & " Union ALL " & vbNewLine & strTemp
        gstrSQL = "Select * From (" & vbNewLine & gstrSQL & vbNewLine & ") Order By ����ʱ�� Asc"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "", CLng(.TextMatrix(.Row, GetCN("����ID"))), _
                glngDeptId, "", CLng(.TextMatrix(.Row, GetCN("����ID"))))

        cboTimes.Clear
        Do Until rsTemp.EOF
           cboTimes.AddItem "��" & rsTemp.AbsolutePosition & "��(" & Format(rsTemp!����ʱ��, "yyyy-mm-dd") & ")  " & Trim(rsTemp!ҽ������)
           cboTimes.ItemData(cboTimes.NewIndex) = rsTemp!ҽ��id
           If rsTemp!ҽ��id = .TextMatrix(.Row, GetCN("ҽ��ID")) Then cboTimes.ListIndex = cboTimes.NewIndex
           rsTemp.MoveNext
        Loop
        cboTimes.Tag = "���"
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub RefreshTabWindow(Optional lngAdviceIDtmp As Long = 0, Optional blnClear As Boolean = False, Optional blnRefresh As Boolean = False)
'lngAdviceIDtmp���μ�¼ʱ���� , ������0, blnclear��յ�ǰ�б�, blnRefreshǿ��ˢ��
'ˢ�µ�ǰҳ��,���ã��б�ѡ�����μ�¼ѡ���Ӵ���ѡ��
'���μ�¼ʱfraInfo.Tag = 0����ID|1��ҳID|2ҽ��ID|3���ͺ�|4���˿���ID|5�Һŵ�|6������Դ|7���UID|8ת��|9ִ��״̬
Dim lngAdviceID As Long, lngSendNO As Long, lngPatID As Long, lngPageID As Long, blnCanPrint As Boolean, blnIsInsidePatient As Boolean
Dim lngUnit As Long, lngPatDept As Long, strRegNo As String, intMoved As Boolean, intState As Integer, i As Integer, intPatientForm As Integer

    On Error GoTo errHandle
    If lngAdviceIDtmp = 0 Then '-----------------------�б�ѡ�����
        If blnClear Then       '�޼�¼ʱ��������Ӵ���
            lngAdviceID = 0: lngSendNO = 0: lngPatID = 0: lngPageID = 0
            lngPatDept = 0: strRegNo = "": intMoved = 0: intState = 0: lngUnit = 0: blnCanPrint = False
        Else
            With vsList
                lngAdviceID = .TextMatrix(.Row, GetCN("ҽ��ID")): lngSendNO = .TextMatrix(.Row, GetCN("���ͺ�"))
                lngPatID = .TextMatrix(.Row, GetCN("����ID")): lngPageID = Val(.TextMatrix(.Row, GetCN("��ҳID")))
                lngPatDept = .TextMatrix(.Row, GetCN("���˿���ID")): strRegNo = .TextMatrix(.Row, GetCN("�Һŵ�"))
                intMoved = .TextMatrix(.Row, GetCN("ת��"))
                intState = IIf(.TextMatrix(.Row, GetCN("������")) = "�Ѿܾ�", 2, IIf(.TextMatrix(.Row, GetCN("������")) = "�����", 1, 3))
                lngUnit = Val(.TextMatrix(.Row, GetCN("��ǰ����ID")))
'                blnCanPrint = IIf(mblnCanPrint, IIf(.Cell(flexcpData, .Row, GetCN("����")) = 1, .TextMatrix(.Row, GetCN("������")) <> "", .TextMatrix(.Row, GetCN("������")) <> ""), True)
                intPatientForm = Decode(.TextMatrix(.Row, GetCN("��Դ")), "��", 1, "ס", 2, "��", 3, 4)
            End With
        End If
    Else                       '----------------------���μ�¼ѡ�����
        lngAdviceID = lngAdviceIDtmp: lngSendNO = Split(fraInfo.Tag, "|")(3)
        lngPatID = Split(fraInfo.Tag, "|")(0): lngPageID = Val(Split(fraInfo.Tag, "|")(1))
        lngPatDept = Split(fraInfo.Tag, "|")(4): strRegNo = Split(fraInfo.Tag, "|")(5)
        intMoved = Split(fraInfo.Tag, "|")(8): intState = Split(fraInfo.Tag, "|")(9)
        lngUnit = lngPatDept
'        blnCanPrint = True
        intPatientForm = Split(fraInfo.Tag, "|")(6)
    End If
    
    blnIsInsidePatient = (intPatientForm = 1) Or (intPatientForm = 2)
    
    mfrmPACSImg.zlRefresh lngAdviceID, lngSendNO, mstrPrivs, intMoved = 1, blnRefresh
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub subTriggleRefreshTimer(blnEnable As Boolean)
    '�������߹ر��Զ�ˢ�µ�Timer
    If blnEnable = False Then
        TimerRefresh.Enabled = False
    Else
        TimerRefresh.Enabled = mlngRefreshInterval > 0
    End If
End Sub

Private Function GetDeptName(lngDeptID As Long, strDeptStrings As String) As String
'ͨ�����õĿ��Ҵ�����ȡָ������ID�Ŀ�������
    Dim strDepts() As String
    Dim i As Integer
    
    On Error GoTo errHandle
    
    strDepts = Split(strDeptStrings, "|")
    For i = 0 To UBound(strDepts)
        If Split(strDepts(i), "_")(0) = lngDeptID Then
            GetDeptName = Split(strDepts(i), "_")(1)
            Exit For
        End If
    Next i
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
