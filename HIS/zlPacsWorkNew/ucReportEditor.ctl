VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "*\Azl9PacsControl\zl9PacsControl.vbp"
Begin VB.UserControl ucReportEditor 
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10380
   ScaleHeight     =   10815
   ScaleWidth      =   10380
   Begin VB.Timer timerTmp 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   7800
      Top             =   120
   End
   Begin VB.PictureBox picState 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   535
      Left            =   120
      ScaleHeight     =   540
      ScaleWidth      =   9975
      TabIndex        =   12
      Top             =   9840
      Width           =   9975
      Begin VB.Label labEditState 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   9600
         TabIndex        =   15
         Top             =   160
         Width           =   240
      End
      Begin VB.Label labFmt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   375
         TabIndex        =   29
         Top             =   0
         Width           =   9120
      End
      Begin VB.Label lab���� 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   270
         Left            =   0
         TabIndex        =   28
         ToolTipText     =   "������"
         Top             =   240
         Width           =   270
      End
      Begin VB.Label labΣ�� 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "!"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   270
         Left            =   0
         TabIndex        =   27
         ToolTipText     =   "Σ��״̬"
         Top             =   0
         Width           =   270
      End
      Begin VB.Label labSignTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "ǩ��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   375
         TabIndex        =   14
         Top             =   240
         Width           =   840
      End
      Begin VB.Label labSign 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   240
         Width           =   7695
      End
   End
   Begin VB.PictureBox picChar 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   960
      ScaleHeight     =   315
      ScaleWidth      =   5295
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   5295
      Begin XtremeCommandBars.CommandBars cbrChar 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.PictureBox picContainer 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   120
      ScaleHeight     =   9015
      ScaleWidth      =   9975
      TabIndex        =   4
      Top             =   720
      Width           =   9975
      Begin VB.PictureBox picImageBack 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   120
         ScaleHeight     =   2895
         ScaleWidth      =   9735
         TabIndex        =   8
         Top             =   480
         Width           =   9735
         Begin zl9PacsControl.ucSplitter ucSplitter1 
            Bindings        =   "ucReportEditor.ctx":0000
            Height          =   2895
            Left            =   5625
            TabIndex        =   9
            Top             =   0
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   5106
            SplitLevel      =   3
            Con1MinSize     =   2000
            Con2MinSize     =   1000
            Control1Name    =   "dcmReportImg"
            Control2Name    =   "dcmMarkImage"
         End
         Begin VB.PictureBox picMarkImgOper 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   5880
            ScaleHeight     =   375
            ScaleWidth      =   1815
            TabIndex        =   21
            Top             =   120
            Visible         =   0   'False
            Width           =   1815
            Begin VB.CommandButton cmdOper 
               BackColor       =   &H0080FF80&
               Caption         =   "4"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   7
               Left            =   1440
               Style           =   1  'Graphical
               TabIndex        =   26
               ToolTipText     =   "����ƶ�����ͼ��"
               Top             =   0
               Width           =   375
            End
            Begin VB.CommandButton cmdOper 
               BackColor       =   &H00FFFFFF&
               Caption         =   "AU"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   3
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   24
               ToolTipText     =   "ɾ������ͼ��"
               Top             =   0
               Width           =   375
            End
            Begin VB.CommandButton cmdOper 
               BackColor       =   &H0080C0FF&
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   4
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   23
               ToolTipText     =   "��ǰ�ƶ�����ͼ��"
               Top             =   0
               Width           =   375
            End
            Begin VB.CommandButton cmdOper 
               BackColor       =   &H00FF80FF&
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   5
               Left            =   720
               Style           =   1  'Graphical
               TabIndex        =   22
               ToolTipText     =   "����ƶ�����ͼ��"
               Top             =   0
               Width           =   375
            End
            Begin VB.CommandButton cmdOper 
               BackColor       =   &H008080FF&
               Caption         =   "3"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   6
               Left            =   1080
               Style           =   1  'Graphical
               TabIndex        =   25
               ToolTipText     =   "����ƶ�����ͼ��"
               Top             =   0
               Width           =   375
            End
         End
         Begin VB.PictureBox picReportImgOper 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   1095
            TabIndex        =   17
            Top             =   0
            Visible         =   0   'False
            Width           =   1095
            Begin VB.CommandButton cmdOper 
               Height          =   375
               Index           =   2
               Left            =   720
               Picture         =   "ucReportEditor.ctx":0014
               Style           =   1  'Graphical
               TabIndex        =   18
               ToolTipText     =   "����ƶ�����ͼ��"
               Top             =   0
               Width           =   375
            End
            Begin VB.CommandButton cmdOper 
               Height          =   375
               Index           =   1
               Left            =   360
               Picture         =   "ucReportEditor.ctx":0716
               Style           =   1  'Graphical
               TabIndex        =   19
               ToolTipText     =   "��ǰ�ƶ�����ͼ��"
               Top             =   0
               Width           =   375
            End
            Begin VB.CommandButton cmdOper 
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   0
               Picture         =   "ucReportEditor.ctx":0E18
               Style           =   1  'Graphical
               TabIndex        =   20
               ToolTipText     =   "ɾ������ͼ��"
               Top             =   0
               Width           =   375
            End
         End
         Begin DicomObjects.DicomViewer dcmMarkImage 
            Height          =   2895
            Left            =   5760
            TabIndex        =   10
            Top             =   0
            Width           =   3975
            _Version        =   262147
            _ExtentX        =   7011
            _ExtentY        =   5106
            _StockProps     =   35
            BackColor       =   4210752
            CellSpacing     =   2
         End
         Begin DicomObjects.DicomViewer dcmReportImg 
            Height          =   2895
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   5625
            _Version        =   262147
            _ExtentX        =   9922
            _ExtentY        =   5106
            _StockProps     =   35
            BackColor       =   4210752
            CellSpacing     =   2
         End
      End
      Begin VB.PictureBox picDesc 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   120
         ScaleHeight     =   1815
         ScaleWidth      =   9735
         TabIndex        =   7
         Top             =   3720
         Visible         =   0   'False
         Width           =   9735
         Begin RichTextLib.RichTextBox rtb���� 
            Height          =   1695
            Left            =   0
            TabIndex        =   0
            Top             =   0
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   2990
            _Version        =   393217
            BorderStyle     =   0
            Enabled         =   -1  'True
            HideSelection   =   0   'False
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"ucReportEditor.ctx":115A
         End
      End
      Begin VB.PictureBox picOpin 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   120
         ScaleHeight     =   1695
         ScaleWidth      =   9735
         TabIndex        =   6
         Top             =   5760
         Visible         =   0   'False
         Width           =   9735
         Begin RichTextLib.RichTextBox rtb��� 
            Height          =   1575
            Left            =   0
            TabIndex        =   1
            Top             =   0
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   2778
            _Version        =   393217
            BorderStyle     =   0
            Enabled         =   -1  'True
            HideSelection   =   0   'False
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"ucReportEditor.ctx":11F7
         End
      End
      Begin VB.PictureBox picAdvi 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   9735
         TabIndex        =   5
         Top             =   7680
         Visible         =   0   'False
         Width           =   9735
         Begin RichTextLib.RichTextBox rtb���� 
            Height          =   975
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   1720
            _Version        =   393217
            BorderStyle     =   0
            Enabled         =   -1  'True
            HideSelection   =   0   'False
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"ucReportEditor.ctx":1294
         End
      End
      Begin XtremeDockingPane.DockingPane dkpMain 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   450
         _ExtentY        =   423
         _StockProps     =   0
      End
   End
   Begin MSComctlLib.ImageList listCur 
      Left            =   240
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucReportEditor.ctx":1331
            Key             =   "pen"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtxtSaveElement 
      Height          =   375
      Left            =   6960
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393217
      BorderStyle     =   0
      Appearance      =   0
      TextRTF         =   $"ucReportEditor.ctx":200B
   End
   Begin VB.Menu menuReport 
      Caption         =   "����ͼ"
      Begin VB.Menu menuReport_Del 
         Caption         =   "ɾ��(&D)"
      End
      Begin VB.Menu menuReport_Split 
         Caption         =   "-"
      End
      Begin VB.Menu menuReport_Last 
         Caption         =   "ǰ��(&L)"
      End
      Begin VB.Menu menuReport_Next 
         Caption         =   "����(&N)"
      End
   End
   Begin VB.Menu menuLab 
      Caption         =   "��ע"
      Visible         =   0   'False
      Begin VB.Menu menuLab_Del 
         Caption         =   "ɾ��(&D)"
      End
   End
End
Attribute VB_Name = "ucReportEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const Report_Element_����ǩ�� = "����ǩ��"

'ǩ��״̬
Private Enum EPRSignLevelEnum
    cprSL_�հ� = 0              'δǩ��
    cprSL_���� = 1              '����ҽʦǩ��
    cprSL_���� = 2              '����ҽʦǩ��
    cprSL_���� = 3              '����ҽʦǩ��
    cprSL_���� = 4              '���ߣ�ǩ�����𲻰�����ֻ��ʾ��Ա��������ְ�ƣ��Ա���������ҽʦ
End Enum

Private Type TReportInfo
    �������� As Date      '��������
    �����û� As String    '������
    ����û� As String    '�����
    ǩ������ As EPRSignLevelEnum
    ���汾 As Long
    Ŀ��汾 As Long
    
End Type


Private Enum TReportFmtFrom
    rffTemplate = 0 '����ģ��
    rffSample = 1   '���Է���
    rffReport = 2   '���Ա���
End Enum


Private Type paneInfo
    title As String
    ID As Long
    hwnd As Long
    hidden As Boolean
    iconid As Long
    options As PaneOptions
    tag As Long
End Type


Private mlngModule As Long
Private mlngDeptID As Long
Private mObjNotify As IEventNotify


Private mlngAdviceId As Long        'ҽ��ID
Private mlngFileID As Long          '��ʽ�ļ�ID
Private mstrEprFmtName As String    '�����ʽ����
Private mstrPrintFmts As String     '�����ӡ��ʽ
Private mlngSampleId As Long        '����ID
Private mlngReportID As Long        '����ID
Private mblnIsMoved As Boolean      '�Ƿ�ת��
Private mblnIsLoadData As Boolean   '�Ƿ��Ѿ�����������
Private mstrReportImgPath As String '����ͼ·��
Private mftpConTag As TFtpConTag    'ftp���ӱ��

Private mintEditFontSize As Integer '�༭�������С
Private mrtbActive As RichTextBox   '��ǰ�༭��
Private mlngSelReportImgIndex As Long   'ѡ��ı���ͼ����
Private mblnIsInit As Boolean       '�Ƿ��ʼ��
Private mstrPrivs As String         'ģ��Ȩ��

Private mobjSpePlugin As Object     'ר�Ʊ�����
Private mblnIsSpeState As Boolean   '�Ƿ�ר�Ʊ���༭ģʽ״̬


Private mblnIsLockingEdit As Boolean   '�Ƿ������༭��
Private mlngSignCount As Long           'ǩ������
Private mlngSignLevel As TReportSignLevel   'ǩ������
Private mstrFirstSignUser As String     '�״�ǩ���û�
Private mstrFinalSignUser As String     '����ǩ���û�
Private mintTargetVer As Integer        'Ŀ��汾
Private mintSourceVer As Integer
Private mstrCreateUser As String        '������
Private mstrSaveUser As String          '��󱣴���
Private mlngCreateDeptId As Long        '��������ID

'��Ҫ�Ӳ��������ж�ȡ
Private mblnTechReptSame As Boolean 'ֻ����д�Լ����ı���
Private mlngSignPassType As Long        'ǩ������ '������֤����ϵͳ������ 0-���룻1�����֣�2�����߽Կ�

Private mblnUseImgSign As Boolean   '�Ƿ�ʹ��ͼ��ǩ��
Private mblnVisibleSpecialty As Boolean '�Ƿ���ʾר�Ʊ���
Private mblnCheckPrintPara As Boolean   'ƽ����Ҫ��˲��ܴ�ӡ
 
Private mblnReportWithResult As Boolean '��Ӱ�����Ϊ����
Private mblnReportDefaultPositive As Boolean '���Ĭ������
Private mblnIgnoreResult As Boolean     '���Խ��������
Private mblnIsEditWithReportImage As Boolean    '��ͼ�����д����

Private mstrDescTitle As String         '��������
Private mstrOpinTitle As String         '�������
Private mstrAdviTitle As String         '�������


Private mblnIsEditable As Boolean   '���������Ƿ��ܹ��༭
Private mblnIsReadOnly As Boolean   '���Ʒ�������صĹ��ܣ�����ˣ����ˣ�Ԥ����ӡ�ȣ���û�ж�ӦȨ��ʱ�ᴦ��true״̬
Private mblnIsComplete As Boolean   '�Ƿ���ɣ���readonly״̬ʱ����鲻һ��Ϊ���״̬�����Խ��б���ɾ����ز���

Private mblnIsModifyText As Boolean
Private mblnIsModifyImage As Boolean
Private mblnIsModifyMarks As Boolean
 
Private mlngMarkType As TImgMarkType
Private mstrMarkText As String

Private WithEvents mobjMarkProcessV2 As frmImageProcessV2
Attribute mobjMarkProcessV2.VB_VarHelpID = -1

Public Event OnOutlineChange(ByVal lngSelOutline As TOutlineType)
Public Event OnStateChange()
Public Event OnDelRepImg(ByVal strImgKey As String) 'ɾ������ͼ�¼�


'��ǰ���
Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

'ҽ��ID
Property Get AdviceId() As Long
    AdviceId = mlngAdviceId
End Property

'����ID
Property Get ReportID() As Long
    ReportID = mlngReportID
End Property

'�Ƿ�ת��
Property Get IsMoved() As Boolean
    IsMoved = mblnIsMoved
End Property

'�Ƿ���ר�Ʊ���
Property Get HasSpeReport() As Boolean
    HasSpeReport = IIf(mobjSpePlugin Is Nothing, False, True)
End Property

'�Ƿ�ר�Ʊ���״̬
Property Get IsSpeState() As Boolean
    IsSpeState = mblnIsSpeState
End Property

Property Let IsSpeState(ByVal value As Boolean)
    Call ChangeSepState(value, False)
End Property

'����ID
Property Get SampleId() As Long
    SampleId = mlngSampleId
End Property

'�������ݸ�ʽ����
Property Get EPRFmtName() As String
    EPRFmtName = mstrEprFmtName
End Property

'��ǩ���汾
Property Get SourceVer() As Long
    SourceVer = mintSourceVer
End Property


'Ŀ��汾
Property Get TargetVer() As Integer
    TargetVer = mintTargetVer
End Property
 

'������
Property Get CreateUser() As String
    CreateUser = mstrCreateUser
End Property
 

'������
Property Get SaveUser() As String
    SaveUser = mstrSaveUser
End Property
 
'��������ID
Property Get CreateDeptId() As Long
    CreateDeptId = mlngCreateDeptId
End Property
 
'Ȩ�޴�
Property Get Privs() As String
    Privs = mstrPrivs
End Property

Property Let Privs(ByVal value As String)
    mstrPrivs = value
End Property


'�༭�����С
Property Get EditFontSize() As Integer
    EditFontSize = mintEditFontSize
End Property

Property Let EditFontSize(ByVal value As Integer)
    mintEditFontSize = value
    
    If mintEditFontSize <> 0 Then
        Call SetContextFont(mintEditFontSize)
    Else
        Call SetContextFont(gbytFontSize + 3)
    End If
End Property

Public Sub SetContextFont(ByVal intFontSize As Integer)
    rtb����.Font.Size = intFontSize
    rtb���.Font.Size = intFontSize
    rtb����.Font.Size = intFontSize

    rtb����.SelFontSize = intFontSize
    rtb���.SelFontSize = intFontSize
    rtb����.SelFontSize = intFontSize
    
End Sub


'ǩ������
Property Get SignCount() As Long
    SignCount = Val(labSign.tag)
End Property

'ǩ������
Property Get SignPassType() As Long
    SignPassType = mlngSignPassType
End Property

Property Let SignPassType(ByVal value As Long)
    mlngSignPassType = value
End Property


'�Ƿ����
Property Get IsComplete() As Boolean
    IsComplete = mblnIsComplete
End Property

'ֻ������
Property Get IsReadOnly() As Boolean
    IsReadOnly = mblnIsReadOnly
End Property

'Property Let IsReadOnly(ByVal value As Boolean)
'    mblnIsReadOnly = value
'End Property

'�ɱ༭����
Property Get IsEditable() As Boolean
    IsEditable = mblnIsEditable
End Property

Property Let IsEditable(ByVal value As Boolean)
    mblnIsEditable = value
End Property

Property Get IsModify() As Boolean
'�жϱ����Ƿ����޸�
    IsModify = mblnIsModifyText Or mblnIsModifyImage Or mblnIsModifyMarks
    
    If Not mobjSpePlugin Is Nothing Then
        IsModify = IsModify Or mobjSpePlugin.pModified
    End If
End Property


'�Ƿ���ı����ݸ����޸�
Property Get IsModifyText() As Boolean
    IsModifyText = mblnIsModifyText
End Property

'�Ƿ�Ա���ͼ������޸�
Property Get IsModifyImage() As Boolean
    IsModifyImage = mblnIsModifyImage
End Property

'ͼ�����Ƿ��޸�
Property Get IsModifyMarks() As Boolean
    IsModifyMarks = mblnIsModifyMarks
End Property


'ר��
Property Get VisibleSpecialty() As Boolean
    VisibleSpecialty = mblnVisibleSpecialty
End Property

Property Let VisibleSpecialty(ByVal value As Boolean)
    mblnVisibleSpecialty = value
End Property





'�����������-----------------------
Property Get DescTitle() As String
    DescTitle = mstrDescTitle
End Property

Property Let DescTitle(ByVal value As String)
    mstrDescTitle = value
End Property

'��Ͻ������-----------------------
Property Get AdviTitle() As String
    AdviTitle = mstrAdviTitle
End Property

Property Let AdviTitle(ByVal value As String)
    mstrAdviTitle = value
End Property

'����������-----------------------
Property Get OpinTitle() As String
    OpinTitle = mstrOpinTitle
End Property

Property Let OpinTitle(ByVal value As String)
    mstrOpinTitle = value
End Property


Property Get DescContext() As String
'�����������
    DescContext = rtb����.Text
End Property

Property Get OpinContext() As String
'����������
    OpinContext = rtb���.Text
End Property

Property Get AdviContext() As String
'��������
    AdviContext = rtb����.Text
End Property


'����ͼ
Property Get RepImageCount() As Long
    RepImageCount = dcmReportImg.Images.Count
End Property


Property Get RepImage(ByVal lngIndex As Long) As Object
On Error GoTo errhandle
    Set RepImage = dcmReportImg.Images(lngIndex)
Exit Sub
errhandle:
    Set RepImage = Nothing
End Property

'���ͼ
Property Get MarkImageCount() As Long
    MarkImageCount = dcmMarkImage.Images.Count
End Property

Property Get MarkImage() As Object
On Error GoTo errhandle
    Set MarkImage = dcmMarkImage.Images(0)
Exit Property
errhandle:
    Set MarkImage = Nothing
End Property


Property Get CurOutlineType() As TOutlineType
    CurOutlineType = otNone
    
    If mrtbActive Is Nothing Then Exit Property
    
    If mrtbActive Is rtb���� Then
        CurOutlineType = otDesc
    End If
    
    If mrtbActive Is rtb��� Then
        CurOutlineType = otOpin
    End If
    
    If mrtbActive Is rtb���� Then
        CurOutlineType = otAdvi
    End If
End Property

Public Sub SetFontSize(ByVal intFontSize As Integer)
    Dim objCapFont As New StdFont
    
    FontSize = intFontSize
    
    objCapFont.Name = FontName
    objCapFont.Size = intFontSize + 3
  
    Set dkpMain.PaintManager.CaptionFont = objCapFont
        
    Set cbrChar.options.Font = objCapFont
    
    picChar.FontSize = FontSize
    
    dkpMain.RecalcLayout
End Sub


Public Sub ChangeSepState(ByVal blnState As Boolean, ByVal blnIsForceRefresh As Boolean)
    Dim i As Long
    Dim objPane As Pane
    Dim Left As Long, Right As Long
    Dim Top As Long, Bottom As Long
    Dim strErr As String
    
    mblnIsSpeState = False
    
    If mobjSpePlugin Is Nothing Then Exit Sub
    
    Call dkpMain.GetClientRect(Left, Top, Right, Bottom)
    
    If blnState Then
        If dkpMain.PanesCount < 5 Then
            'ר�Ʊ���¼��
            If dkpMain.Panes(1).Closed = False Then
                Set objPane = dkpMain.CreatePane(5, 0, 1000 - (picImageBack.Height / (Height - 3000)) * 1000, DockBottomOf, dkpMain.Panes(1))
            Else
                Set objPane = dkpMain.CreatePane(5, 0, 1000 - (picImageBack.Height / (Height - 3000)) * 1000, DockBottomOf)
            End If
            
            objPane.title = "ר��¼��"
            objPane.Handle = mobjSpePlugin.hwnd
            objPane.tag = 4
            objPane.options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
            objPane.Closed = True
        End If
    End If
    
    For i = 1 To dkpMain.PanesCount
        If dkpMain.Panes(i).tag <> 0 And dkpMain.Panes(i).tag <> 4 Then
             '������Ч�Ž�����ʾ
             dkpMain.Panes(i).Closed = IIf(blnState, True, dkpMain.Panes(i).iconid = 0)
            
        ElseIf dkpMain.Panes(i).tag = 4 Then
            dkpMain.Panes(i).Closed = IIf(blnState, False, True)
            
            mblnIsSpeState = blnState
            
        End If
    Next
    
    If mblnIsSpeState Then
        
        If dkpMain.Panes(5).ID = mlngAdviceId And blnIsForceRefresh = False Then
            picChar.Visible = False
            Exit Sub
        End If
        
        On Error GoTo errhandle
            mobjSpePlugin.Refresh mlngAdviceId, mlngReportID, mblnIsEditable And Not mblnIsReadOnly, mblnIsMoved
errhandle:
        strErr = err.Description
        If err.Number <> 0 Then MsgboxH GetRootHwnd, "ר�Ʊ�����ˢ�´���:" & strErr, vbOKOnly, "��ʾ"
        
        dkpMain.Panes(5).ID = mlngAdviceId
         
    End If
    
    picChar.Visible = False
End Sub


Private Sub cbrChar_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strErr As String
On Error GoTo errhandle
    Dim objWordCharCfg As frmWordCharCfgV2
    
    Select Case Control.ID
        Case 1  '���ó��ôʾ�
            Set objWordCharCfg = New frmWordCharCfgV2
            If objWordCharCfg.zlShowWordCharCfg(mlngModule, mObjNotify.Owner) Then
                Call InitReportChar
                
                Call mObjNotify.Broadcast(BM_REPORT_EVENT_REFWCHR, , Parent.hwnd, glngSys, mlngModule)
            End If
            
        Case Else   'д��ѡ��Ĵʾ�
            If mblnIsEditable = False Then Exit Sub
            
            If mrtbActive Is Nothing Then Exit Sub
            mrtbActive.SelText = Control.Caption
    End Select
Exit Sub
errhandle:
    strErr = err.Description
    
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "��ʾ"
End Sub

'Private Sub chkCritical_Click()
'On Error GoTo errhandle
'
'    If chkCritical.value = 0 Then
'        chkCritical.ForeColor = &H404040
'    Else
'        chkCritical.ForeColor = ColorConstants.vbRed
'    End If
'
'    If mblnIsLoadData = False Or mblnIsEditable = False Then Exit Sub
'
'    mblnIsModifyText = True
'Exit Sub
'errhandle:
'    Debug.Print "chkPositive_Click:" & err.Description
'End Sub
'
'Private Sub chkPositive_Click()
'On Error GoTo errhandle
'    If chkPositive.value = 0 Then
'        chkPositive.ForeColor = &H404040
'    Else
'        chkPositive.ForeColor = ColorConstants.vbRed
'    End If
'
'    If mblnIsLoadData = False Or mblnIsEditable = False Then Exit Sub
'
'    mblnIsModifyText = True
'Exit Sub
'errhandle:
'    Debug.Print "chkPositive_Click:" & err.Description
'End Sub


Public Sub Init(objNotify As IEventNotify, ByVal lngModuleNo As Long, ByVal lngDeptId As Long, _
    ByVal strPrivs As String, lngSignPassType As Long, Optional ByVal blnIsForce As Boolean = False)
'ģ���ʼ��
    mlngModule = lngModuleNo
    mlngDeptID = lngDeptId
    mlngSignPassType = lngSignPassType
    
    Set mObjNotify = objNotify
    
    mstrPrivs = strPrivs
    
    If mblnIsInit And blnIsForce = False Then Exit Sub
    
    Call InitPar
    
    Call InitReportChar
    
    Call Relayout
    
    mblnIsInit = True
End Sub


Public Sub InitReportChar()
    Dim cbrToolBar As CommandBar
    Dim strWord As String
    Dim aryWord() As String
    Dim i As Long
    Dim blnIsSetGroup As Boolean
    Dim lngWordLen As Long
    
    cbrChar.DeleteAll
    
    
    With cbrChar.options
        .UpdatePeriod = 800
        
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .UseDisabledIcons = False
        .LargeIcons = False
    End With
    
    Set cbrToolBar = cbrChar.Add("�����ַ�", xtpBarTop)
     
    With cbrToolBar
        .Position = xtpBarTop
        .Customizable = False
        .ShowTextBelowIcons = True
        .Closeable = False
        .EnableDocking xtpFlagHideWrap
    End With
    
    strWord = zlDatabase.GetPara("���泣�ôʾ�", glngSys, mlngModule)
    aryWord = Split(strWord & vbCrLf, vbCrLf)
    
    blnIsSetGroup = False
    With cbrToolBar.Controls
        .Add(xtpControlButton, 1, "��").ToolTipText = "�����ַ�����"
        
        For i = 0 To UBound(aryWord)
            If Len(aryWord(i)) > 0 Then
                lngWordLen = TextWidth(aryWord(i))
                
                If lngWordLen <= picChar.Width - 2000 Then
                    If blnIsSetGroup = False Then
                        .Add(xtpControlButton, i + 2, aryWord(i)).BeginGroup = True
                        blnIsSetGroup = True
                    Else
                        .Add xtpControlButton, i + 2, aryWord(i)
                    End If
                    
                    aryWord(i) = ""
                End If
            End If
        Next
        
        For i = 0 To UBound(aryWord)
            If Len(aryWord(i)) > 0 Then
                If blnIsSetGroup = False Then
                    .Add(xtpControlButton, i + 2, aryWord(i)).BeginGroup = True
                    blnIsSetGroup = True
                Else
                    .Add xtpControlButton, i + 2, aryWord(i)
                End If
            End If
        Next
    End With
End Sub


Private Sub InitPar()
'��ʼ������
    mblnIgnoreResult = Val(GetDeptPara(mlngDeptID, "���Խ��������", 0)) <> 0 '        '���Խ��������
    mblnReportDefaultPositive = Val(GetDeptPara(mlngDeptID, "��Ͻ��Ĭ������", 0)) <> 0
    mblnTechReptSame = Val(GetDeptPara(mlngDeptID, "ֻ����д�Լ����ı���", 0)) <> 0
    mblnReportWithResult = Val(GetDeptPara(mlngDeptID, "��Ӱ�����Ϊ����", 0)) <> 0 '  '��Ӱ�����Ϊ����
    mblnVisibleSpecialty = Val(GetDeptPara(mlngDeptID, "��ʾר�Ʊ���", 0)) <> 0
    mblnUseImgSign = Val(GetDeptPara(mlngDeptID, "ͼ��ǩ����֤")) <> 0
    mblnIsEditWithReportImage = Val(GetDeptPara(mlngDeptID, "��ͼ�����д����", 0)) <> 0
    
    mstrDescTitle = GetDeptPara(mlngDeptID, "�����������", "�������")
    mstrOpinTitle = GetDeptPara(mlngDeptID, "����������", "������")
    mstrAdviTitle = GetDeptPara(mlngDeptID, "��������", "��Ͻ���")
End Sub


Public Sub Refresh(ByVal lngAdviceId As Long, ByVal lngFileId As Long, ByVal lngSampleId As Long, ByVal lngReportID As Long, _
    Optional ByVal blnIsMoved As Boolean = False, Optional ByVal blnIsForce As Boolean = False)
    Dim lngSelStart As Long
    
    '���������ͬ���ҷ�ǿ��ˢ�£���ֱ���˳�
    If lngFileId = mlngFileID _
        And lngSampleId = mlngSampleId _
        And lngReportID = mlngReportID _
        And Not blnIsForce Then Exit Sub
    
    mblnIsLoadData = False
    
    picReportImgOper.Visible = False
    picMarkImgOper.Visible = False
    
    '���ҽ��ID�����в�ͬ��˵���ǲ�ͬ���ı��棬�򱨸�༭���㲻��Ҫ���б���
    If lngAdviceId <> mlngAdviceId Then Set mrtbActive = Nothing
    
    mlngAdviceId = lngAdviceId
    mblnIsMoved = blnIsMoved
    mlngFileID = lngFileId
    mlngSampleId = lngSampleId
    mlngReportID = lngReportID ' 0 '
    mlngSelReportImgIndex = 0
    
    mblnIsModifyMarks = False
    mblnIsModifyImage = False
    mblnIsModifyText = False
    
    mblnIsEditable = False
    
    mlngMarkType = imtAuto ' imtNormal
    
    mftpConTag.Ip = ""
    
    mblnIsLockingEdit = False
    
    
    '�������û�м��أ��򲻽�����ʾ
'    If Extender.Visible = False Then Exit Sub     'And Not blnIsForce ��Ҫ����Ԥ�����ӡ

    '����֮ǰ�Ĺ�������ı���λ��
    If Not mrtbActive Is Nothing Then lngSelStart = mrtbActive.SelStart
    
    Call ResetContext
    
    If mintEditFontSize <> 0 Then
        Call SetContextFont(mintEditFontSize)
    Else
        Call SetContextFont(gbytFontSize)
    End If
    
    mstrReportImgPath = GetReportImgPath(lngAdviceId, blnIsMoved)
    
    '���뱨��
    Call LoadReport
    
    '�ָ��ı���Ĺ��λ��
    If Not mrtbActive Is Nothing And lngSelStart > 0 Then mrtbActive.SelStart = lngSelStart
    
    '����ר�Ʊ���
    If Not mobjSpePlugin Is Nothing Then
        
        If mblnIsSpeState Then
            '�ָ���ר����ʾ����
            Call ChangeSepState(True, blnIsForce)
        Else
            '�����ǿ��ˢ�£������¶�id�������ã��Ա�����л���ר�Ʊ���ʱ�ܹ�����ˢ�²���
            If blnIsForce And Not dkpMain.Panes(5) Is Nothing Then dkpMain.Panes(5).ID = -5
            
        End If
    End If
    
    Call ShowPrintFormat(mstrPrintFmts)
    
    mblnIsLoadData = True
End Sub


Public Sub ResetContext()
    mlngCreateDeptId = mlngDeptID ' 0
    
    mstrCreateUser = UserInfo.���� ' ""
    mstrSaveUser = UserInfo.���� ' ""
     
    mlngSignCount = 0
    mlngSignLevel = cprSL_�հ�
    mstrFirstSignUser = ""
    mstrFinalSignUser = ""
    mintTargetVer = 1
    mintSourceVer = 0
    
    labEditState.Caption = ""
    
    dcmMarkImage.Images.Clear
    dcmReportImg.Images.Clear
    picChar.Visible = False
    
    rtb����.Text = ""
    rtb���.Text = ""
    rtb����.Text = ""
    
    labSign.Caption = ""
    labSign.tag = ""
    
'    chkPositive.value = 0
'    chkCritical.value = 0
    
    mblnIsModifyImage = False
    mblnIsModifyMarks = False
    mblnIsModifyText = False
    
'    If mblnIgnoreResult = False Then
'        '��������������ԣ������������Ե�Ĭ��ֵ
'        chkPositive.value = Abs(CLng(mblnReportDefaultPositive))
'    End If
    
'    mblnTechReptSame = False
End Sub

Private Sub LoadReport()
'���뱨��
    Dim strSQL As String
    Dim strPicSql As String
    Dim strContextSql As String
    Dim rsData As ADODB.Recordset
    Dim lngFileId As Long
    Dim lngDataFrom As TReportFmtFrom
    Dim strTmp As String
    Dim blnHas���� As Boolean
    Dim blnHas��� As Boolean
    Dim blnHas���� As Boolean
    Dim strTitle As String
    Dim blnForceRead As Boolean
    Dim blnReportVisible As Boolean
    Dim blnMarkVisible As Boolean
    Dim i As Long
    Dim blnReadyRepImg As Boolean
    Dim strFile As String
    
    '������ͬ����ʱ��������Ϊnothing
'    Set mrtbActive = Nothing
    
    mstrEprFmtName = ""
    If mlngFileID <> 0 Then
        strSQL = "Select ���� From �����ļ��б� where ID=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ��������", mlngFileID)
        
        If rsData.RecordCount > 0 Then mstrEprFmtName = nvl(rsData!����)
    End If
    
    '����ͼ��ѯ...
    If mlngReportID <> 0 Then
        lngFileId = mlngReportID
        lngDataFrom = rffReport
        
        '�ӵ��Ӳ��������в�ѯ����
        strSQL = "Select  Id As ���Id From ���Ӳ�������" & _
                    " Where �ļ�id = [1] And �������� = 3 And Substr(��������, Instr(��������, ';', 1, 18) + 1, 1) = '2' " & _
                    " Order By �������"
                    
        strPicSql = "select ID,�ļ�ID,��ID,��ʼ��,������,��������,�����д� from ���Ӳ������� where  �ļ�ID=[1] and ��ID=[2] and ��������=5 order by ������"
        
        strContextSql = "Select a.�����ı� As ����, b.��������, b.�����ı� As ����" & vbNewLine & _
                 " From ���Ӳ������� a,���Ӳ������� b " & _
                 " Where a.�ļ�id = [1] And (a.�������� = 3) And a.Id = b.��ID And b.�������� = 2 And b.��ֹ�� = 0"
        '(a.�������� = 3 or a.��������=1 ) order by a.������� '֧�ֲ�ʹ��1*1�ı��
        
        If mblnIsMoved Then
            strSQL = Replace(strSQL, "���Ӳ�������", "H���Ӳ�������")
            strPicSql = Replace(strPicSql, "���Ӳ�������", "H���Ӳ�������")
            strContextSql = Replace(strContextSql, "���Ӳ�������", "H���Ӳ�������")
        End If
        
    Else
        If mlngSampleId <> 0 Then
            lngDataFrom = rffSample
            lngFileId = mlngSampleId
            
            '�ӷ����в�ѯ��ʽ����
            strSQL = "Select  Id As ���Id From ������������ a " & _
                        " Where �ļ�id = [1] And �������� = 3 And Substr(��������, Instr(��������, ';', 1, 18) + 1, 1) = '2' " & _
                        " Order By �������"
                        
            strPicSql = "select ID,�ļ�ID,��ID,1 as ��ʼ��,������,��������,�����д� from ������������ where  �ļ�ID=[1] and ��ID=[2] and ��������=5 order by ������"
            
            strContextSql = "Select a.�����ı� As ����, b.��������, b.�����ı� As ����" & vbNewLine & _
                    " From ������������ a, ������������ b" & vbNewLine & _
                    " Where a.�ļ�id = [1] And (a.�������� = 3 ) And a.Id = b.��id And b.�������� = 2"
            '(a.�������� = 3 or a.��������=1 ) order by a.������� '֧�ֲ�ʹ��1*1�ı��
        Else
            lngDataFrom = rffTemplate
            lngFileId = mlngFileID
            
            '�Ӳ��������в�ѯ��ʽ����
            strSQL = "Select  Id As ���Id From �����ļ��ṹ" & _
                        " Where �ļ�id = [1] And �������� = 3 And Substr(��������, Instr(��������, ';', 1, 18) + 1, 1) = '2' " & _
                        " Order By �������"
                        
            strPicSql = "select ID,�ļ�ID,��ID,1 as ��ʼ��,������,��������,�����д� from �����ļ��ṹ where  �ļ�ID=[1] and ��ID=[2] and ��������=5 order by ������"
            
            strContextSql = "Select a.�����ı� As ����, b.��������, b.�����ı� As ���� " & _
                     " From �����ļ��ṹ a, �����ļ��ṹ b" & _
                     " Where a.�ļ�id = [1] And (a.�������� = 3 ) And a.Id = b.��id And b.�������� = 2 "
                     
            '(a.�������� = 3 or a.��������=1 ) order by a.������� '֧�ֲ�ʹ��1*1�ı��
        End If
    End If
    
    '��ȡ����ͼ��Ϣ****************************************
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ����ͼ��", lngFileId)
    
    dcmReportImg.MultiColumns = 1
    dcmReportImg.MultiRows = 1
    
    dcmReportImg.Visible = False
    dcmMarkImage.Visible = False
    
    blnReportVisible = False
    blnMarkVisible = False
    
    dcmReportImg.Images.Clear
    dcmMarkImage.Images.Clear

    If rsData.RecordCount > 0 Then
        '��ȡ���ͼ������ͼ
        blnReportVisible = True
        dcmReportImg.Visible = True
        'ͼ������ѯ
        dcmReportImg.tag = Val(nvl(rsData!���ID))
        
        Set rsData = zlDatabase.OpenSQLRecord(strPicSql, "��ѯ����ͼƬ", lngFileId, Val(nvl(rsData!���ID)))
        If rsData.RecordCount > 0 Then
            
            Call ParshReportImgData(rsData, lngDataFrom)
            
            If dcmMarkImage.Images.Count > 0 Then blnMarkVisible = True
        End If
        
        '��ȡԤ�����õı���ͼ
        'ֻ�гɹ����ص����ص�ͼ�񣬲����Զ���ӵı���ͼ��
        strPicSql = "select ͼ��UID from Ӱ����ͼ�� a, Ӱ�������� b, Ӱ�����¼ c where  a.����UID=b.����UID and b.���UID=c.���UID and c.ҽ��ID=[1] and a.����ͼ>=0 order by ͼ��ʱ��"
        If mblnIsMoved Then
            strPicSql = Replace(strPicSql, "Ӱ����ͼ��", "HӰ����ͼ��")
            strPicSql = Replace(strPicSql, "Ӱ��������", "HӰ��������")
            strPicSql = Replace(strPicSql, "Ӱ�����¼", "HӰ�����¼")
        End If
        
        Set rsData = zlDatabase.OpenSQLRecord(strPicSql, "��ѯԤ�豨��ͼ", mlngAdviceId)
        If rsData.RecordCount > 0 Then
            While Not rsData.EOF
                blnReadyRepImg = True
                For i = 1 To dcmReportImg.Images.Count
                    If nvl(rsData!ͼ��UID) = dcmReportImg.Images(i).InstanceUID Then
                        blnReadyRepImg = False
                        Exit For
                    End If
                Next
                
                If blnReadyRepImg Then
                    strFile = FormatFilePath(mstrReportImgPath & "\" & nvl(rsData!ͼ��UID))
                    
                    If FileExists(strFile) Then
                        '���ڶ�Ӧ��ͼ��
                        Call AddRepImgFile(strFile, , , True)
                    End If
                End If
                
                Call rsData.MoveNext
            Wend
        End If
    End If
    
    
    If blnReportVisible = False And blnMarkVisible = False Then
        '�رձ���ͼ
        dkpMain.Panes(1).Closed = True
    Else
        '�򿪱���ͼ
        dkpMain.Panes(1).Closed = False
        
        If blnMarkVisible = False Then
            dcmReportImg.Width = picImageBack.Width
            ucSplitter1.Visible = False
        Else
            If dcmReportImg.Width = picImageBack.Width Then
                If picImageBack.Width - dcmMarkImage.Width < 0 Then
                    dcmMarkImage.Width = 0.34 * picImageBack.Width
                End If
                
                dcmReportImg.Width = picImageBack.Width - dcmMarkImage.Width
                ucSplitter1.Left = dcmReportImg.Width
                ucSplitter1.RePaint
            End If
            
            ucSplitter1.Visible = True
        End If
    End If

    '��ȡ�����ı�����****************************************
    Set rsData = zlDatabase.OpenSQLRecord(strContextSql, "��ѯ�����ı�", lngFileId)
    
    rtb����.Text = ""
    rtb���.Text = ""
    rtb����.Text = ""
    
    blnHas���� = False
    blnHas��� = False
    blnHas���� = False
    
    While rsData.EOF = False
        strTmp = nvl(rsData!��������)
        strTitle = nvl(rsData!����)
        
'        If strTitle = "�������" Or InStr(strTitle, "����") >= 1 Then
'            ReadReport nvl(rsData!����), strTmp, rtb����
'            blnHas���� = True
'
'        ElseIf strTitle = "������" Or InStr(strTitle, "���") >= 1 Then
'            ReadReport nvl(rsData!����), strTmp, rtb���
'            blnHas��� = True
'
'        ElseIf strTitle = "����" Or InStr(strTitle, "����") >= 1 Then
'            ReadReport nvl(rsData!����), strTmp, rtb����
'            blnHas���� = True
'        End If
        

        Select Case nvl(rsData!����)
            Case "�������"
                ReadReport nvl(rsData!����), strTmp, rtb����
                blnHas���� = True

            Case "������"
                ReadReport nvl(rsData!����), strTmp, rtb���
                blnHas��� = True

            Case "����"
                ReadReport nvl(rsData!����), strTmp, rtb����
                blnHas���� = True

        End Select
        
        rsData.MoveNext
    Wend

    blnForceRead = False
    If blnHas���� = False And blnHas��� = False And blnHas���� = False Then
        blnForceRead = True
        
        If lngFileId <> 0 Then
            labEditState.Caption = "�޶�Ӧ��ٹ���"
            picChar.Visible = False
            MsgboxH GetRootHwnd, "��Ч�ı����ʽ���á�", vbOKOnly, "��ʾ"
        End If
        
        dkpMain.Panes(2).Closed = False
        dkpMain.Panes(3).Closed = False
        dkpMain.Panes(4).Closed = False
        
        dkpMain.Panes(2).iconid = 0
        dkpMain.Panes(3).iconid = 0
        dkpMain.Panes(4).iconid = 0
    Else
        dkpMain.Panes(2).Closed = Not blnHas����
        dkpMain.Panes(3).Closed = Not blnHas���
        dkpMain.Panes(4).Closed = Not blnHas����
        
        dkpMain.Panes(2).iconid = IIf(blnHas����, 2, 0)
        dkpMain.Panes(3).iconid = IIf(blnHas���, 3, 0)
        dkpMain.Panes(4).iconid = IIf(blnHas����, 4, 0)
    End If
        
    rtb����.Enabled = blnHas����
    rtb���.Enabled = blnHas���
    rtb����.Enabled = blnHas����
    
    '��ȡ����ǩ���������Ϣ****************************************
    lab����.ForeColor = &H808080
    labΣ��.ForeColor = &H808080
    
    If mlngReportID <> 0 Then
        Call ReadResultTag(mlngAdviceId, mblnIsMoved, 0) '֧�ֶ౨������£�����Ҫ���ݱ���ID
        
        Call ReadVersion(mlngReportID)
        Call ReadSigns(mlngReportID)
    End If

    Call ConfigFaceState(blnForceRead)
    
End Sub

Public Sub ReadResultTag(ByVal lngAdviceId As Long, ByVal blnMoved As Boolean, Optional ByVal lngReportID As Long = 0)
'��ȡ������
'�����ǰ��������ԣ�Σ��״̬����Ϣ
'lngReportID�౨������£���ͨ���ò���ָ��ĳ�ݱ���
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
On Error GoTo errhandle
    lab����.ForeColor = &H808080
    labΣ��.ForeColor = &H808080
    
    If lngReportID = 0 Then
        strSQL = "select �������, Σ��״̬ " & _
                " from Ӱ�����¼ A, ����ҽ������ B " & _
                " where A.ҽ��ID=B.ҽ��ID and A.���ͺ�=B.���ͺ� And A.ҽ��id=[1]"
    Else
        'TODO:�౨������£����ݱ���ID���в�ѯ...
    End If
    
    If blnMoved Then
        strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ������״̬", lngAdviceId)
    If rsData.RecordCount <= 0 Then Exit Sub
    
    If IsNull(rsData!�������) = False Then
        If Val(nvl(rsData!�������)) <> 0 Then
            lab����.ForeColor = vbRed
        Else
            lab����.ForeColor = &H808080
        End If
    End If
    
    If Val(nvl(rsData!Σ��״̬)) <> 0 Then
        labΣ��.ForeColor = vbRed
    Else
        labΣ��.ForeColor = &H808080
    End If
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub ReadVersion(ByVal lngReportID As Long)
'��ȡ����汾
'ǩ��ʱĿ��汾����Ҫ����1
On Error GoTo errhandle
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    strSQL = "Select ���汾,ǩ������,������,������,����ID From ���Ӳ�����¼  Where Id =[1]"
    If mblnIsMoved = True Then
        strSQL = Replace(strSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ����ǩ���汾", lngReportID)
    
    If rsTemp.RecordCount > 0 Then
        mlngCreateDeptId = Val(nvl(rsTemp!����ID))
        mstrCreateUser = nvl(rsTemp!������)
        mstrSaveUser = nvl(rsTemp!������)
        mlngSignLevel = nvl(rsTemp!ǩ������, cprSL_�հ�)
        mintTargetVer = nvl(rsTemp!���汾, 1)
'    Else
'        'û�б���ʱ�ĸ�ֵ����
'        mlngCreateDeptId = 0
'        mstrCreateUser = ""
'        mstrSaveUser = ""
    End If
    
    If mlngSignLevel = cprSL_�հ� Then
        mintSourceVer = 0
    Else
        mintSourceVer = mintTargetVer
    End If
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub ReadSigns(ByVal lngReportID As Long)
'------------------------------------------------
'���ܣ���ȡǩ������ɾ������ǩ���Ķ������´����ݿ��ȡ��ȷ��ǩ����������ݸ����ݿ��һ�£�ǩ������ˢ��֮����ñ�����
'������ ��
'���أ� ��
'-----------------------------------------------
On Error GoTo errH
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim reportSignInfo As TReportSignInfo
    Dim strSigns As String
    Dim strSignName As String
    
    mlngSignCount = 0
    
    strSQL = "Select Id,������ From ���Ӳ������� Where �ļ�id= [1] And ��������=8 Order By ������"
    
    If mblnIsMoved = True Then
        strSQL = Replace(strSQL, "���Ӳ�������", "H���Ӳ�������")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ����ǩ����¼", lngReportID)
    While rsTemp.EOF = False
        If GetReportSignInfo(Val(rsTemp!ID), reportSignInfo, mblnIsMoved) Then
            reportSignInfo.Key = nvl(rsTemp!������, 0)
            
            If Len(strSigns) > 0 Then strSigns = strSigns & "  "
            
            If InStr(reportSignInfo.����, M_STR_TAG_SIGNWITHIMG) > 0 Then
                strSignName = Mid(reportSignInfo.����, 1, InStr(reportSignInfo.����, M_STR_TAG_SIGNWITHIMG) - 1)
            Else
                strSignName = reportSignInfo.����
            End If
            
            strSigns = strSigns & reportSignInfo.ǰ������ & strSignName
            
            mstrFinalSignUser = strSignName
            
            If mstrFirstSignUser = "" Then
                mstrFirstSignUser = strSignName
            End If
        End If
        
        rsTemp.MoveNext
    Wend
     
    mlngSignCount = rsTemp.RecordCount
    
    '��дǩ���ı���
    labSign.Caption = strSigns
    labSign.tag = mlngSignCount    '����ǩ������
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub ReadReport(ByVal strtext As String, ByVal strPros As String, rText As RichTextBox)
    'intType---0 ���������1 ��������2 ����
    Dim lngCount As Long
    Dim lngSelStart As Long
    Dim lngPosStart As Long
    Dim lngPosEnd As Long
    Dim aryTextPros() As String
    
    On Error GoTo err
    
    lngSelStart = rText.SelStart
    aryTextPros = Split(strPros, "|")
    
    rText.tag = strPros
    
    rText.SelLength = 0
    rText.SelText = strtext
    '������ɫ
    rText.SelStart = lngSelStart
    rText.SelLength = Len(strtext)
    rText.SelColor = vbBlack
    
    On Error Resume Next
    'rText.Tag �ǵ��Ӳ�����ʽ�Ķ������ԣ��á�|���ָ����ܹ�26��Ԫ��
    rText.SelStart = 0
    rText.SelLength = Len(rText.Text)
    rText.SelFontName = aryTextPros(15)     '  rText.SelFontName
    
    If mintEditFontSize <> 0 Then
        rText.SelFontSize = mintEditFontSize
    Else
        rText.SelFontSize = gbytFontSize
    End If
        
    rText.SelBold = aryTextPros(17)     'rText.SelBold
    rText.SelItalic = aryTextPros(18)   'rText.SelItalic
    
    On Error GoTo 0
    
    '������ǰ��������֣��Ƿ���Ҫ�أ������������ɫ��ʾ����
    '�Ȳ��ѡҪ��
    For lngCount = 1 To Len(strtext)
        lngPosStart = InStr(lngCount, strtext, "{{")
        lngPosEnd = InStr(lngCount, strtext, "}}")
        If lngPosStart <> 0 And lngPosEnd <> 0 And lngPosEnd > lngPosStart Then
            '���ҵ�Ҫ�أ����Ҫ������ɫ��ʾ
            rText.SelStart = lngSelStart + lngPosStart - 1
            rText.SelLength = lngPosEnd - lngPosStart + 2
            rText.SelColor = vbBlue
            lngCount = lngPosEnd
        Else
            Exit For
        End If
    Next lngCount
    
    '�ٲ鵥ѡҪ��
    For lngCount = 1 To Len(strtext)
        lngPosStart = InStr(lngCount, strtext, "{<")
        lngPosEnd = InStr(lngCount, strtext, ">}")
        If lngPosStart <> 0 And lngPosEnd <> 0 And lngPosEnd > lngPosStart Then
            '���ҵ�Ҫ�أ����Ҫ������ɫ��ʾ
            rText.SelStart = lngSelStart + lngPosStart - 1
            rText.SelLength = lngPosEnd - lngPosStart + 2
            rText.SelColor = vbBlue
            lngCount = lngPosEnd
        Else
            Exit For
        End If
    Next lngCount
    
    rText.SelStart = lngSelStart + Len(strtext)
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ParshReportImgData(rsData As ADODB.Recordset, ByVal lngDataFrom As TReportFmtFrom)
'��������ͼ������
    Dim aryImgPro() As String
    Dim reportImgTag As TReportImgTag
    Dim result As ftpResult
    Dim blnIsAbort As Boolean
    Dim objDcmImg As DicomImage
    
 
    If rsData Is Nothing Then Exit Sub
    
    rsData.MoveFirst
    blnIsAbort = False
    
    While Not rsData.EOF
        '��һ������˵����0��ͨͼ��1���ͼ��2����ͼ��
        aryImgPro = Split(nvl(rsData!��������) & ";;;;;;;;;;;;;;;;;;;;", ";")
        
        reportImgTag.lngFileId = Val(rsData!�ļ�ID)
        reportImgTag.lngTableId = Val(rsData!��ID)
        reportImgTag.strObjectTag = Val(rsData!������)
        reportImgTag.strPros = nvl(rsData!��������)
        reportImgTag.lngStartVer = Val(rsData!��ʼ��)
        reportImgTag.strKey = Val(rsData!ID)
        reportImgTag.strImgMarks = ""
        
        If Val(aryImgPro(0)) = 1 Then '���ͼ
            reportImgTag.lngImgType = ritMark
            
            Call ReadMarkImage(dcmMarkImage.Images, lngDataFrom, reportImgTag)
            
            dcmMarkImage.Visible = True
        End If
        
        If Val(aryImgPro(0)) = 2 Then '����ͼ
            reportImgTag.lngImgType = ritReport
            reportImgTag.lngFromAdvice = Val(GetReportImagePro(reportImgTag.strPros, "ADVICEID"))
            
            If blnIsAbort = False Then
                result = ReadReportImage(dcmReportImg.Images, reportImgTag)
            Else
                '�����滻ͼ��
                Set objDcmImg = dcmReportImg.Images.AddNew
                dcmReportImg.Images(dcmReportImg.Images.Count).tag = reportImgTag
                
                Call DrawBorder(objDcmImg, 0)
                Call DrawErrorText(objDcmImg, "�ѱ���ֹ")
                
            End If
            
            Call CalcImgView
            
            If result = frAbort Then
                '��������쳣����ѡ����ֹ���أ����˳�ͼ����ش���
                blnIsAbort = True
            End If
        End If
        
        Call rsData.MoveNext
    Wend
End Sub

Private Function GetRootHwnd() As Long
    Dim lngCurHwnd As Long
    
    lngCurHwnd = GetAncestor(hwnd, GA_ROOT)
    
On Error GoTo errhandle
    '�ڴ��ڵ�queryunload�¼��е��ø÷���ʱ����parent���κη��ʶ�����ʾ�ͻ��˲����ô���
    If Parent.hwnd = lngCurHwnd Then
        If Parent.Visible = False Then
            lngCurHwnd = MainForm.hwnd
        End If
    End If
errhandle:
    If err.Number <> 0 Then
        lngCurHwnd = MainForm.hwnd
    End If
    
    GetRootHwnd = lngCurHwnd
End Function

Private Function ReadMarkImage(objImages As DicomImages, _
    ByVal lngDataFrom As TReportFmtFrom, reportImgTag As TReportImgTag) As Boolean
'��ȡ���ͼ��
    Dim strFile As String
    Dim strSQL As String
    Dim lngAction As Long
    Dim rsTemp As ADODB.Recordset
    Dim objPicMarks As clsPicMarks
    Dim dblMarkZoom As Double
    Dim strError As String
    Dim strTableName As String
    Dim objDcmImg As DicomImage
    Dim strFileName As String
    
    ReadMarkImage = False
    
    Select Case lngDataFrom
        Case rffReport
            lngAction = 6
            strTableName = "���Ӳ�������"
            
            If mblnIsMoved Then strTableName = "H���Ӳ�������"
            
        Case rffSample
            lngAction = 4
            strTableName = "������������"
            
        Case rffTemplate
            lngAction = 2
            strTableName = "�����ļ��ṹ"
            
    End Select
    
    strFileName = "MarkImage_" & reportImgTag.lngFileId & "_" & reportImgTag.strKey & ".JPG"
    strFile = mstrReportImgPath & strFileName
                
    If DirExists(mstrReportImgPath) = False Then Call MkLocalDir(mstrReportImgPath)
    
    If FileExists(strFile) = False Then
        Call Sys.ReadLob(glngSys, lngAction, reportImgTag.strKey, strFile)
    End If
    
    If FileExists(strFile) Then
        Set objDcmImg = ReadDicomFile(strFile, strError)
        
        If objDcmImg Is Nothing Then
            MsgboxH GetRootHwnd, "���ͼ��ȡʧ��:" & strError, vbOKOnly, "��ʾ"
        Else
            objDcmImg.tag = reportImgTag
            
            '���Ʊ߿�
            Call DrawBorder(objDcmImg, 0)
            
            Call objImages.Add(objDcmImg)
         
        
            '��ȡ���picMarks...
            strSQL = "Select �����ı� " & _
                " From " & strTableName & _
                " Where �ļ�ID = [1] And ��id=[2] And ��������=6 " & _
                " Order By �����д�"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ������ͼ���", reportImgTag.lngFileId, reportImgTag.strKey)
        
            While Not rsTemp.EOF
                reportImgTag.strImgMarks = reportImgTag.strImgMarks & nvl(rsTemp!�����ı�)
                
                Call rsTemp.MoveNext
            Wend
            
            reportImgTag.strImgFile = strFileName
            objImages(1).tag = reportImgTag
            
            Set objPicMarks = New clsPicMarks
            
            objPicMarks.�������� = reportImgTag.strImgMarks
            
            dblMarkZoom = objImages(1).SizeX / Val(GetReportImagePro(reportImgTag.strPros, "width")) * Screen.TwipsPerPixelX
            
            '���Ʊ��
            Call DrawMarks(objImages(1), objPicMarks, dblMarkZoom)
            
            ReadMarkImage = True
        End If
    Else
        MsgboxH GetRootHwnd, "���ͼ��ȡʧ��", vbOKOnly, "��ʾ"
    End If
End Function
 
 
Private Function DownLoadFtpFile(ByVal lngAdviceId As Long, ByVal strFtpFile As String, ByVal strLocalFile As String, ByVal blnMoved As Boolean) As ftpResult
'����ftp�ļ�
    DownLoadFtpFile = frNormal
    If Len(mftpConTag.Ip) <= 0 Or Val(mftpConTag.tag) <> lngAdviceId Then
        mftpConTag = GetReportDevice(lngAdviceId, blnMoved)
        mftpConTag.tag = lngAdviceId
        
        If Len(mftpConTag.Ip) <= 0 Then
            DownLoadFtpFile = frAbort
            Exit Function
        End If
    End If
    
    DownLoadFtpFile = FtpDownload(mftpConTag, strFtpFile, strLocalFile)
End Function


Private Function UpLoadFtpFile(ByVal lngAdviceId As Long, ByVal strFtpFile As String, ByVal strLocalFile As String, ByVal blnMoved As Boolean) As ftpResult
'�ϴ�ftp�ļ�
    UpLoadFtpFile = frNormal
    If Len(mftpConTag.Ip) <= 0 Or Val(mftpConTag.tag) <> lngAdviceId Then
        mftpConTag = GetReportDevice(lngAdviceId, blnMoved)
        mftpConTag.tag = lngAdviceId
        
        If Len(mftpConTag.Ip) <= 0 Then
            UpLoadFtpFile = frAbort
            Exit Function
        End If
    End If
    
    UpLoadFtpFile = FtpUpload(mftpConTag, strFtpFile, strLocalFile)
End Function

Private Function GetReportDevice(ByVal lngAdviceId As Long, ByVal blnMoved As Boolean) As TFtpConTag
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
On Error GoTo errhandle
    strSQL = "select NVl(���ID, ID) as ID from ����ҽ����¼ where ID=[1]"
    If blnMoved Then strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ��ҽ��ID", lngAdviceId)
    
    If rsData.RecordCount <= 0 Then
        Call MsgboxH(GetRootHwnd, "ҽ������У��ʧ�ܣ�δ�ҵ��������ҽ����Ϣ��", vbOKOnly, "��ʾ")
        Exit Function
    End If
    
    strSQL = " Select Decode(A.��������,Null,'',to_Char(A.��������,'YYYYMMDD')||'/') ||A.���UID||'/' As URL," & _
            " B.�豸�� as �豸��1, B.�豸�� As �豸��1, B.FTP�û��� As User1,B.FTP���� As Pwd1, B.IP��ַ As Host1, " & _
                    " decode(B.FtpĿ¼, null, '/', '/'||B.FtpĿ¼||'/') As Root1,B.����Ŀ¼ as ����Ŀ¼1,B.����Ŀ¼�û��� as ����Ŀ¼�û���1,B.����Ŀ¼���� as ����Ŀ¼����1 " & _
            " From  Ӱ�����¼ A,Ӱ���豸Ŀ¼ B " & _
            " Where A.ҽ��ID=[1] And nvl(A.λ��һ, A.λ�ö�)=B.�豸��(+)  "
    If blnMoved Then
        strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ����ͼ��洢", Val(rsData!ID))
            
    If rsData.RecordCount <= 0 Then
        Call MsgboxH(GetRootHwnd, "δ�ҵ�����ͼ��Ӧ�Ĵ洢�豸�����������Ƿ���ȷ��", vbOKOnly, "��ʾ")
        Exit Function
    End If
    
    If nvl(rsData!Host1) <> "" Then
        GetReportDevice = FtpTagInstance(rsData!Host1, rsData!User1, rsData!Pwd1, rsData!Root1 & rsData!Url)
    End If
Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ReadReportImage(objImages As DicomImages, reportImgTag As TReportImgTag) As ftpResult
'��ȡ����ͼ
    Dim strFile As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim objPicMarks As clsPicMarks
    Dim dblMarkZoom As Double
    Dim strError As String
    Dim objDcmImg As DicomImage
    Dim strFileName As String
    Dim blnImgReadState As Boolean
    Dim lngAdviceId As Long
    
    
    ReadReportImage = frNormal
    blnImgReadState = True
    
    strFileName = GetReportImagePro(reportImgTag.strPros, "PicName")
    If Len(strFileName) > 0 Then
        
        strFile = FormatFilePath(mstrReportImgPath & "\" & strFileName)
        
        '��ftp����ͼ��
        If FileExists(strFile) = False Then
            lngAdviceId = GetReportImagePro(reportImgTag.strPros, "ADVICEID")
            If lngAdviceId = mlngAdviceId Then
                ReadReportImage = DownLoadFtpFile(lngAdviceId, strFileName, strFile, mblnIsMoved)
            Else
                '������ҽ�������ر���ͼ��
                strSQL = "Select ID From ����ҽ����¼ where Id=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯҽ����¼", lngAdviceId)
                
                If rsTemp.RecordCount > 0 Then
                    ReadReportImage = DownLoadFtpFile(lngAdviceId, strFileName, strFile, False)
                Else
                    ReadReportImage = DownLoadFtpFile(lngAdviceId, strFileName, strFile, True)
                End If
            End If
            
            
            If ReadReportImage <> frNormal Then
                blnImgReadState = False
            End If
        End If
    Else
        strFile = FormatFilePath(mstrReportImgPath & "\����ͼ_" & reportImgTag.strKey & ".JPG")
        
        '�����ݿ��ȡͼ��
        If FileExists(strFile) = False Then
            Call Sys.ReadLob(glngSys, 6, reportImgTag.strKey, strFile)
        End If
    End If
    
    If FileExists(strFile) = False Then
        If Len(strError) <= 0 Then strError = "δ�ҵ�����ͼ���ļ� [" & strFile & "]"
        blnImgReadState = False
    End If
    
    If blnImgReadState Then
        'ͼ���ȡ�ɹ��Ĵ���
        Set objDcmImg = ReadDicomFile(strFile, strError)
        
        If Not objDcmImg Is Nothing Then
            reportImgTag.strImgFile = strFileName
            
            objDcmImg.InstanceUID = Replace(strFileName, ".JPG", "")
            objDcmImg.tag = reportImgTag
            
            Call objImages.Add(objDcmImg)
            Call DrawBorder(objDcmImg, 0)
        Else
            blnImgReadState = False
        End If
    End If
    
    If blnImgReadState = False Then
        '����ʧ�ܵ�ͼ��
        
        Set objDcmImg = objImages.AddNew
        
        objImages(objImages.Count).tag = reportImgTag
        
        Call DrawBorder(objDcmImg, 0)
        Call DrawErrorText(objDcmImg, strError)
        
        If ReadReportImage = frNormal Then Call MsgboxH(GetRootHwnd, "ͼ���ȡʧ�ܡ�" & vbCrLf & strError, vbOKOnly, "��ʾ")
    End If
End Function

'Private Function GetReportImgFiles() As String
''��ȡ����ͼ�ļ�
'End Function
'
'Private Function GetMarkImgFile(objMarks As cPicMarks) As String
''��ȡ���ͼ�ļ�
'End Function

Public Sub AutoSave()
'TODO:�Զ�����

End Sub

Public Function PromptSave(ByVal lngNewAdviceId As Long, ByVal lngNewReportId As Long, Optional ByVal blnIsForceHint As Boolean = False) As Boolean
'������ʾ
    Dim blnIsHint As Boolean
    
    PromptSave = False
    '���û���޸ģ���ֱ���˳�
    If mblnIsEditable = False Or IsModify = False Then Exit Function
    
    blnIsHint = False
     
    If lngNewAdviceId <> mlngAdviceId Then
        blnIsHint = True
    End If
    
    If lngNewReportId <> mlngReportID Then
        blnIsHint = True
    End If
    
    If blnIsHint Or blnIsForceHint Then
        If MsgboxH(GetRootHwnd, "�����ѱ��޸ģ��Ƿ񱣴�", vbYesNo Or vbDefaultButton1, "��ʾ") = vbNo Then
            If mlngReportID = 0 Then
                '�������״̬
                Call UpdateReporter(mlngAdviceId, "")
            End If
            
            mblnIsModifyImage = False
            mblnIsModifyMarks = False
            mblnIsModifyText = False
         
            Exit Function
        End If
    End If
    
    PromptSave = SaveReport
    
End Function


Private Function CreateReport() As Long
'��������
    Dim iType As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    
On Error GoTo errhandle
    CreateReport = 0
    
    ' iType��0-�Ӳ����ļ��б�������, 1-�Ӳ�������Ŀ¼��������
    iType = 0
    If mlngSampleId <> 0 Then iType = 1
    
    '�������Ӳ�������
    strSQL = "ZL_Ӱ�񱨸�����_����(" & mlngAdviceId & "," & mlngFileID & "," & mlngSampleId & "," & iType & ")"
    zlDatabase.ExecuteProcedure strSQL, "���������ʽ"
    
    '�´����ı��棬�����ݿ��ж�ȡ��������ID
    strSQL = "Select ����ID From ����ҽ������ Where ҽ��ID= [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�½�����ID", mlngAdviceId)
    
    If rsTemp.EOF = True Then
        MsgboxH GetRootHwnd, "������������ȷ���޷����ҵ���������ID��", vbOKOnly, "��ʾ"
        Exit Function
    End If
    
    CreateReport = Val(nvl(rsTemp!����Id))
Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function DelReportData(ByVal lngReportID As Long, Optional ByVal blnIsErrHint As Boolean = False) As Boolean
'ɾ������
    Dim strSQL As String
On Error GoTo errhandle
    DelReportData = False
    
    If Not mobjSpePlugin Is Nothing Then
        If PluginAction(lngReportID, 1) = False Then Exit Function   'ɾ������
    End If
    
    strSQL = "Zl_���Ӳ�����¼_Delete(" & lngReportID & ")"
    
    zlDatabase.ExecuteProcedure strSQL, "ɾ������"
    
    DelReportData = True
Exit Function
errhandle:
    If blnIsErrHint Then
        If ErrCenter() = 1 Then Resume
    End If
    
    Call SaveErrLog
End Function


Private Sub ResetRtbTag(rText As RichTextBox)
    Dim strItem() As String
    Dim i As Integer
    Dim intCnt As Integer
    
    
    '�޸ĸ��ı����TAG,���TAGΪ�գ�����ʱ����¼
    If rText.tag <> "" Then
        strItem = Split(rText.tag, "|")
        
        strItem(15) = nvl(rText.SelFontName, "����")     'FontName
        strItem(17) = nvl(rText.SelBold, "False")    'FontBold
        strItem(18) = nvl(rText.SelItalic, "False")    'FontItalic
        
        rText.tag = ""
        For i = 0 To UBound(strItem()) - 1
            rText.tag = rText.tag & strItem(i) & "|"
        Next i
                
    End If
End Sub

Private Function GetSpecialtyContext() As String
'��ȡר�Ʊ�������
    Dim strSpeModifyContext As String
    GetSpecialtyContext = ""
    
    If mobjSpePlugin Is Nothing Then Exit Function
    If mobjSpePlugin.pModified = False Then Exit Function
    
    strSpeModifyContext = mobjSpePlugin.getElementString
    
    '��������޸�״̬����ר�Ʊ�������Ϊ�գ�˵����ɾ��������ר�Ʊ��������
    If Len(strSpeModifyContext) <= 0 Then
        strSpeModifyContext = "[[@]]ר�Ʊ���[[;]]"
    End If
    
    'TODO:�������ݲ���Ϊ������֤����
    GetSpecialtyContext = strSpeModifyContext & _
                "[[@]]��������[[;]]����һ�α��潨�������" & vbCrLf & "����ֱ����ר�Ʊ�������д���." & _
                "[[@]]δ����Ҫ��[[;]]���Ҫ����û�ж����" & _
                "" '"[[@]]�������[[;]]����һ�μ���������������"
End Function

Public Sub WriteContext(ByVal lngReportID As Long, ByRef arrSQL() As String)
    Dim strReport As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strElements As String
    Dim blnInTrans As Boolean
    Dim i As Integer
    Dim intLevel As Integer 'ǩ������
    Dim strSQLLevel As String 'ǩ����ѯ
    Dim rsTempLevel As ADODB.Recordset 'ǩ����ѯ���
    Dim strUnitName As String
    Dim strSpecialtyContext As String


    On Error GoTo errhandle
    
    If mobjSpePlugin Is Nothing Then
        If mblnIsModifyText = False Then Exit Sub
    Else
        If mblnIsModifyText = False And mobjSpePlugin.pModified = False Then Exit Sub
    End If


    ReDim Preserve arrSQL(UBound(arrSQL) + 1)


    '�޸ı���ǩ��Ҫ�أ����������滻Ϊ�� ��
    strElements = SPLITER_REPORT & Report_Element_����ǩ�� & SPLITER_ELEMENT & " "
    '��֯ר�Ʊ�������
    strSpecialtyContext = GetSpecialtyContext
    strElements = strElements & strSpecialtyContext
    
    '�ж�ר�Ʊ������Ƿ�����������������
    If Len(strSpecialtyContext) > 0 Then
        If InStr(strSpecialtyContext, "[[@]]�������[[;]]") > 0 Then
            rtb����.Text = ParseSpecialtyElement(strSpecialtyContext, "�������")
        End If
        
        If InStr(strSpecialtyContext, "[[@]]������[[;]]") > 0 Then
            rtb���.Text = ParseSpecialtyElement(strSpecialtyContext, "������")
        End If
        
        If InStr(strSpecialtyContext, "[[@]]����[[;]]") > 0 Then
            rtb����.Text = ParseSpecialtyElement(strSpecialtyContext, "����")
        End If
    End If

    '��֯���ı��εĶ�������,���TagΪ�գ�������ݿ��ȡĬ��ֵ
    If rtb����.tag = "" Or rtb���.tag = "" Or rtb����.tag = "" Then
        strSQL = "Select a.�����ı� As ����, b.�������� " & _
                " From ���Ӳ������� a,���Ӳ������� b " & _
                " Where a.�ļ�id = [1] And a.�������� = 3 And a.Id = b.��ID And b.�������� = 2 And b.��ֹ�� = 0"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�����ı�����", lngReportID)

        While rsTemp.EOF = False
            Select Case rsTemp!����
                Case "�������"
                    If rtb����.tag = "" Then
                        rtb����.tag = rsTemp!��������
                    End If
                Case "������"
                    If rtb���.tag = "" Then
                        rtb���.tag = rsTemp!��������
                    End If
                Case "����"
                    If rtb����.tag = "" Then
                        rtb����.tag = rsTemp!��������
                    End If
            End Select
            
            rsTemp.MoveNext
        Wend
    End If
    
    ResetRtbTag rtb����
    ResetRtbTag rtb���
    ResetRtbTag rtb����
    

    '��󱣴���ı������ݣ���ʱ��������ݿ����ݣ��Զ����±����е�Ҫ��
    strReport = SPLITER_REPORT & _
        "1" & rtb����.tag & SPLITER_ELEMENT & rtb����.Text & SPLITER_REPORT & _
        "2" & rtb���.tag & SPLITER_ELEMENT & rtb���.Text & SPLITER_REPORT & _
        "3" & rtb����.tag & SPLITER_ELEMENT & rtb����.Text

    '����ţ�80185
    'ʹ���������ǩ������
    '�������ݵ�ʱ�򣬱����ǩ������ʼ����0���������ǩ������ͨ��ǩ���Ĺ���������
    strSQLLevel = " Select id as ����id,ǩ������ " & _
                " From ���Ӳ�����¼ Where id = [1] "
    Set rsTempLevel = zlDatabase.OpenSQLRecord(strSQLLevel, "��ȡ�Ƿ�ǩ��", lngReportID)
    
    If rsTempLevel.EOF = True Then
        intLevel = 0
    Else
        intLevel = nvl(rsTempLevel!ǩ������)
    End If

    strUnitName = zlRegInfo("��λ����")

    strSQL = "ZL_Ӱ�񱨸�����_update(" & mlngAdviceId & "," & _
                                        lngReportID & _
                                        ",'" & Replace(strReport, "'", "��") & _
                                        " ','" & strElements & "'," & _
                                        mintTargetVer & "," & _
                                        intLevel & _
                                        ",'" & strUnitName & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = strSQL
    
'    If chkCritical.value <> 0 Then
'        'Σ��״̬
'        strSQL = "zl_Ӱ����_Σ������(" & mlngAdviceId & ",1)"
'
'        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
'        arrSQL(UBound(arrSQL)) = strSQL
'    End If
'
'    If mblnIgnoreResult = False And chkPositive.value <> 0 Then
'        'û�к���������
'        strSQL = "ZL_Ӱ����_���(" & mlngAdviceId & ",1)"
'
'        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
'        arrSQL(UBound(arrSQL)) = strSQL
'    End If
    
    Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub WritePicMarks(ByVal lngReportID As Long, ByVal blnCreate As Boolean, _
    ByRef arrSQL() As String)
'д��ͼ����
On Error GoTo errhandle
    Dim i As Integer
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim reportImgTag As TReportImgTag
    
    'û�б��ͼ���˳�
    If mblnIsModifyMarks = False Then Exit Sub
    If dcmMarkImage.Visible = False Then Exit Sub
    If dcmMarkImage.Images.Count <= 0 Then Exit Sub
    
    reportImgTag = dcmMarkImage.Images(1).tag
    
    'û�б����ֱ���˳�
    If Len(reportImgTag.strImgMarks) <= 0 Then
        'ֱ�����±�����ͼƬ
        strSQL = "ZL_Ӱ�񱨸��ע_����(" & reportImgTag.strKey & ",'')"
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
    
        Exit Sub
    End If

    If blnCreate = True Then
        '�´����ı��棬�ӵ��Ӳ��������ж�ȡ���ͼID
        strSQL = "Select Id From ���Ӳ������� Where �ļ�ID=[1] And  ��������= 5 And substr(��������,1,1)='1' "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�½�����ID", lngReportID)
        
        If rsTemp.EOF = False Then  '�б��ͼ
            reportImgTag.strKey = Val(rsTemp!ID)
        Else    'û�б��ͼ
            reportImgTag.strKey = 0
        End If
        
        '���±��ͼ����
        dcmMarkImage.Images(1).tag = reportImgTag
    End If
    
    strSQL = "ZL_Ӱ�񱨸��ע_����(" & reportImgTag.strKey & ",'" & reportImgTag.strImgMarks & "')"
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = strSQL
Exit Sub
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function WriteImages(ByVal lngReportID As Long, ByVal blnCreate As Boolean, _
    ByRef arrSQL() As String, Optional ByRef strReportImgs As String = "") As Boolean
'д�뱨��ͼ
'ֻ�з�ת��״̬�ļ�����ִ�е��˹���
    Dim lngTableId  As Double
    Dim reportImgTag As TReportImgTag
    Dim ftpResult As ftpResult
    Dim strLocalFile As String
    Dim iImgCount As Integer
    Dim strSQL  As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim strPicAttrs As String
    Dim strBufferDir As String
    Dim strRepImgName As String
    Dim lngImageAdiceId As Long
    
    On Error GoTo errhandle
    
    WriteImages = True
     
    If mblnIsModifyImage = False Then Exit Function
    If dcmReportImg.Visible = False Then Exit Function
       
    
    lngTableId = Val(dcmReportImg.tag)
    
    
    If blnCreate = True Then
        strSQL = "Select Id As ���Id From ���Ӳ�������" & vbNewLine & _
            " Where �ļ�id = [1] And �������� = 3 And Substr(��������, Instr(��������, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
            " Order By �������"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ����ͼ��ID", lngReportID)
        If rsTemp.RecordCount > 0 Then
            lngTableId = Val(nvl(rsTemp!���ID))
            dcmReportImg.tag = lngTableId
        End If
        
        '������½����棬��û�б���ͼ�����˳���������
        If dcmReportImg.Images.Count <= 0 Then Exit Function
    Else
        '���½�����£��������ͼ����Ϊ0������Ҫ�������ͼ
        If dcmReportImg.Images.Count <= 0 Then
            strSQL = "ZL_Ӱ�񱨸�ͼ��_����(" & lngTableId & ",'')"
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSQL
            
            Exit Function
        End If
    End If

  
    strBufferDir = mstrReportImgPath
    
    '�ж�Ŀ¼�Ƿ����
    Call MkLocalDir(strBufferDir)
    
    '�����ͱ���ÿһ��ͼ����
    iImgCount = dcmReportImg.Images.Count
    strPicAttrs = ""
    
    lngImageAdiceId = mlngAdviceId
    
    For i = 1 To iImgCount
        reportImgTag = dcmReportImg.Images(i).tag
       
        strLocalFile = FormatFilePath(strBufferDir & "\" & reportImgTag.strImgFile)
        If FileExists(strLocalFile) = False Then
            '��������ļ������ڣ����dicomͼ���е���
            'dcmReportImg.Images(i).FileExport strLocalFile, "JPG"
            dcmReportImg.Images(i).FileExport strLocalFile, "BMP"   '�����ļ���ʽͷ����...
        End If
        
        '˵������ͼ���Ǵ����������������ȡ����Ҫ������ͼ��洢����Ӧ�ļ���豸��
        If reportImgTag.lngFromAdvice <> 0 Then
            lngImageAdiceId = reportImgTag.lngFromAdvice
        End If
        
        ftpResult = UpLoadFtpFile(lngImageAdiceId, reportImgTag.strImgFile, strLocalFile, False)
                        
        If ftpResult <> frNormal Then
            WriteImages = False
            Exit Function
        End If
         
        'ֻ��ҽ����ͬ��ʱ�򣬲���Ҫ���±���ͼ״̬��ʾ
        If lngImageAdiceId = mlngAdviceId Then
            strReportImgs = strReportImgs & ";" & reportImgTag.strKey
        End If
        
        strRepImgName = GetReportImagePro(reportImgTag.strPros, "picname")
        If Len(strRepImgName) <= 0 Then strRepImgName = reportImgTag.strImgFile
        
        strPicAttrs = strPicAttrs & ";" & strRepImgName & "," & lngImageAdiceId
        
        If Len(reportImgTag.strPros) <= 0 Then
            '����������ı���ͼ����û��strPros����
            reportImgTag.strPros = strPicAttrs
            dcmReportImg.Images(i).tag = reportImgTag
        End If
    Next
 
    strSQL = "ZL_Ӱ�񱨸�ͼ��_����(" & lngTableId & ",'" & strPicAttrs & "')"
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = strSQL
Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ParseSpecialtyElement(ByVal strSpecialtyContext As String, ByVal strElementName As String) As String
'����ר�Ʊ����п��ܰ�����Ҫ�أ�
'Ҫ�ظ�ʽΪ����ʽ[[@]]Ҫ������[[;]]ר�Ʊ�������
    Dim lngStartIndex As Long
    Dim strElementContext As String
    
    ParseSpecialtyElement = " "
    
    If Len(strSpecialtyContext) <= 0 Then Exit Function
    
    lngStartIndex = InStr(strSpecialtyContext, strElementName & "[[;]]")
    
    If lngStartIndex <= 0 Then Exit Function
    
    strElementContext = Mid(strSpecialtyContext, lngStartIndex + Len(strElementName & "[[;]]"))
    
    lngStartIndex = InStr(strElementContext, "[[@]]")
    
    If lngStartIndex > 0 Then
        ParseSpecialtyElement = Mid(strElementContext, 1, lngStartIndex - 1)
    Else
        ParseSpecialtyElement = strElementContext
    End If
End Function

Private Function WriteRtfFormat(ByVal lngReportID As Long, ByRef arrSQL() As String) As Boolean
'------------------------------------------------
'���ܣ����汨���ʽRTF�ļ����Ա������ǩ�����߻���
'������     OneSign -- ��Ϊ�գ����ʾ����ǩ�����߻��ˣ�Ϊ�գ���ʾֻ�Ǳ����ʽ��������ǩ��
'           blnAddSign ���ӻ��߻���ǩ����True--����ǩ��,OneSignΪ�ձ�ʾ���汨���ʽ��False--����ǩ��
'���أ� �ޣ�ֱ�ӱ���RTF�����ʽ�ĵ����Ա���ǩ�����߻���
'-----------------------------------------------
On Error GoTo errhandle
    Dim strZipFile As String
    Dim strTemp As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strReport As String
    Dim lngSignPos As Long
    Dim strReportFormatFile As String
    Dim strErrCount As String
    Dim strElementContext As String
    Dim strSpecialtyContext As String
    
    strErrCount = ""
    WriteRtfFormat = False
    
reLoad:
    strReportFormatFile = FormatFilePath(SysRootPath & "\ReportFmt" & strErrCount)
    
    '�ȸ��Ʊ����ʽ
    If Dir(strReportFormatFile) <> "" Then Call RemoveFile(strReportFormatFile)
    
    '�����ݿ��ȡRTF�����ʽ�ĵ�
    strZipFile = Sys.ReadLob(glngSys, 5, lngReportID, strReportFormatFile)
    
    '��ѹ���ļ�
    strTemp = zlFileUnzip(strZipFile)
    
    If strTemp <> "" Then
        If FileExists(strTemp) = False Then
            If MsgboxH(GetRootHwnd, "����IDΪ[" & lngReportID & "]��RTF��ʽ�ļ���ȡʧ�ܣ��Ƿ����ԣ�", vbYesNo) = vbYes Then
                strErrCount = CStr(Val(strErrCount) + 1)
                GoTo reLoad
            End If
            
            Exit Function
        End If
        '�����ļ������ݱ������ݣ��޸�����Ҫ������
        '��ȡRTF�ļ�����
        rtxtSaveElement.Filename = strTemp
        strReport = rtxtSaveElement.TextRTF
        
        strSpecialtyContext = ""
        If Not mobjSpePlugin Is Nothing Then
            strSpecialtyContext = GetSpecialtyContext
        End If
       
        '��ȡ���ݿ��е�Ҫ�أ��Ѹ���Ҫ��������д����ʽ��
        strSQL = "Select ������,�����ı�,Ҫ������ From ���Ӳ������� Where �ļ�ID= [1] And �������� = 4 And ��ֹ��=0 and �������� =0 order by ������ "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ����Ҫ��", lngReportID)
        
        While (rsTemp.EOF = False)
            strElementContext = nvl(rsTemp!�����ı�, "")
            
            If Len(strElementContext) <= 0 Then
                '�ж�ר�Ʊ������Ƿ��������
                '��ʽ[[@]]Ҫ������[[;]]ר�Ʊ�������
                
                strElementContext = ParseSpecialtyElement(strSpecialtyContext, nvl(rsTemp!Ҫ������))
            End If
            
            
            UpdateReportElement strReport, "E", rsTemp!������, strElementContext
            rsTemp.MoveNext
        Wend
        
        '����RTF�ļ�
        rtxtSaveElement.TextRTF = strReport
        rtxtSaveElement.SaveFile strTemp
            
        'ѹ���ļ�
        strZipFile = zlFileZip(strTemp)
        
        '�����ʽ
        zlSaveLob 5, lngReportID, strZipFile, arrSQL
        
        WriteRtfFormat = True
    
        'ɾ����ʱzip�ļ�
        Call RemoveFile(strZipFile)
    Else
        If MsgboxH(GetRootHwnd, "�޷���ȡ���߽�ѹ�����ʽ" & strReportFormatFile & vbCrLf & "��ʹ�á������༭���ķ������༭�˱�������Զ�ȡ���Ƿ����ԣ�", vbYesNo) = vbYes Then
            If Dir(strReportFormatFile) <> "" Then Kill strReportFormatFile
            
            strErrCount = CStr(Val(strErrCount) + 1)
            GoTo reLoad
        End If
    End If
    
    Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function SaveReport(Optional ByRef strReportImages As String = "") As Boolean
'���汨��
    Dim i As Long
    Dim arySql() As String
    Dim blnIsNew As Boolean
    Dim blnIsSaveImg As Boolean
    Dim blnInTrans As Boolean
    
On Error GoTo errhandle
    
    '���汾ӦΪǩ������+1�����û��ǩ���������汾Ϊ1
'    If blnIsSignSave Then 'If mlngSignLevel <> cprSL_�հ� Then
'        mintTargetVer = mintSourceVer + 1   '�����ǩ�����棬����Ҫ��Ŀ��汾��һ
'    End If
    
    If Not IsHaveContent() Then
        MsgBoxD Me, "û����Ч�ı������ݣ��������档", vbInformation, gstrSysName
        Exit Function
    End If
            
    mintTargetVer = mlngSignCount + 1   'Ŀ��汾ֱ�Ӻ�ǩ�������������ǩ������+1����û��ǩ������ʱ��Ŀ��汾Ϊ1�����ǩ������Ϊ1����Ŀ��汾Ϊ2
   
    SaveReport = False
    
    If Not mblnIsEditable Then
    '�Ǳ༭״̬��������
        MsgboxH GetRootHwnd, "�Ǳ༭״̬�²������档", vbOKOnly, "��ʾ"
        Exit Function
    End If
    
    'δ�Ա�������޸�ʱ�������б���
    If IsModify = False Then Exit Function
    
    '�жϱ����ı��γ����Ƿ񳬹�2000���ַ����������������ʾ�����˳�
    If Len(rtb����.Text) > 2000 Or Len(rtb���.Text) > 2000 Or Len(rtb����.Text) > 2000 Then
        MsgboxH GetRootHwnd, "����������������������������2000����ɾ���������ֺ󱣴档", vbOKOnly, "��ʾ"
        Exit Function
    End If
    
    blnIsNew = False
    If mlngReportID = 0 Then
        '�½�����
        blnIsNew = True
        mlngReportID = CreateReport()
    End If
    
    If mlngReportID = 0 Then
        MsgboxH GetRootHwnd, "δȡ����Ч�ı���ID���ݣ����ܼ����˲�����", vbOKOnly, "��ʾ"
        Exit Function
    End If
    
    ReDim arySql(0)
    
    Call WriteContext(mlngReportID, arySql)
    
    Call WritePicMarks(mlngReportID, blnIsNew, arySql)
    
    
    blnIsSaveImg = WriteImages(mlngReportID, blnIsNew, arySql, strReportImages)
    
    
    '������½����棬����û�б��汨���������ݣ�Ҳ��Ҫд��rtf�����ʽ
    If blnIsNew Or UBound(arySql) > 0 Then Call WriteRtfFormat(mlngReportID, arySql)
    
    If blnIsSaveImg = False Then
        Call MsgboxH(GetRootHwnd, "����ͼ�ϴ�ʧ��,δ�ܱ��棬���Ժ����ԡ�", vbOKOnly, "��ʾ")
    End If
   
    gcnOracle.BeginTrans        '----------���汨������
    blnInTrans = True
    For i = 0 To UBound(arySql)
        If Trim(arySql(i)) <> "" Then
            Call zlDatabase.ExecuteProcedure(CStr(arySql(i)), "���汨������[" & i & "]")
        End If
    Next i
    gcnOracle.CommitTrans
    blnInTrans = False
    
    mblnIsModifyImage = False
    mblnIsModifyMarks = False
    mblnIsModifyText = False
    
    If Not mobjSpePlugin Is Nothing Then
        Call PluginAction(mlngReportID, 0)   '���汨��
        mobjSpePlugin.pModified = False
    End If
    
'    If mlngCreateDeptId <= 0 Then mlngCreateDeptId = mlngDeptID
'    If Len(mstrCreateUser) <= 0 Then mstrCreateUser = UserInfo.����
'    If Len(mstrSaveUser) <= 0 Then mstrSaveUser = UserInfo.����
    
    SaveReport = True
Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
    If blnIsNew Then
        'ɾ���½��ı���
        Call DelReportData(mlngReportID)
        
        mlngReportID = 0
    End If
    
    If blnInTrans Then gcnOracle.RollbackTrans
End Function


Private Function PluginAction(ByVal lngReportID As Long, ByVal lngActionType As Long) As Boolean
    Dim strErr As String
On Error GoTo errhandle
    PluginAction = True
    
    If mobjSpePlugin Is Nothing Then Exit Function
    
    PluginAction = mobjSpePlugin.zlReportAction(lngReportID, lngActionType)
Exit Function
errhandle:
    strErr = err.Description
     
    If lngActionType = 1 Then
        If MsgboxH(GetRootHwnd, "ר�Ʊ�����ִ���쳣��" & strErr & vbCrLf & "�Ƿ�ǿ��ɾ���÷ݱ��棿", vbYesNo, "��ʾ") = vbNo Then
            PluginAction = False
        Else
            PluginAction = True
        End If
    Else
        Call MsgboxH(GetRootHwnd, "ר�Ʊ�����ִ���쳣��" & strErr, vbOKOnly, "��ʾ")
        
        PluginAction = False
    End If
 
End Function

Public Sub ShowPrintFormat(ByVal strFmtName As String)
On Error GoTo errhandle
    mstrPrintFmts = strFmtName
    
    labFmt.Caption = IIf(mstrEprFmtName <> "", mstrEprFmtName & "��", "") & mstrPrintFmts
Exit Sub
errhandle:

End Sub

Public Function ChangeReportFormat(ByVal lngFmtId As Long) As Boolean
'���ı����ʽ
    Dim strSQL As String
    Dim strPicSql As String
    Dim strContextSql As String
    Dim rsData As ADODB.Recordset
    Dim strTmp As String
    Dim lngDataFrom As Long
    Dim lngFileId As Long
    Dim blnHas���� As Boolean
    Dim blnHas��� As Boolean
    Dim blnHas���� As Boolean
    
    Dim strSource���� As String
    Dim strSource��� As String
    Dim strSource���� As String
    Dim blnReportVisible As Boolean
    Dim blnMarkVisible As Boolean
     
 
    ChangeReportFormat = False
    
    If mlngReportID <> 0 Or IsModify Then
        If MsgboxH(GetRootHwnd, "���ĸ�ʽ���Ḳ�ǵ�ǰ�������ݣ��Ƿ������", vbYesNo, "��ʾ") = vbNo Then Exit Function
    End If
     
    If lngFmtId = 0 Then
        '��׼��ʽ���ȴӲ����ļ��ṹ�ж�ȡ����
        
        lngDataFrom = rffTemplate
        lngFileId = mlngFileID
        
        '�Ӳ��������в�ѯ��ʽ����
        strSQL = "Select  Id As ���Id From �����ļ��ṹ" & _
                    " Where �ļ�id = [1] And �������� = 3 And Substr(��������, Instr(��������, ';', 1, 18) + 1, 1) = '2' " & _
                    " Order By �������"
                    
        strPicSql = "select ID,�ļ�ID,��ID,1 as ��ʼ��,������,��������,�����д� from �����ļ��ṹ where  �ļ�ID=[1] and ��ID=[2] and ��������=5 order by ������"
        
        strContextSql = "Select a.�����ı� As ����, b.��������, b.�����ı� As ���� " & _
                 " From �����ļ��ṹ a, �����ļ��ṹ b" & _
                 " Where a.�ļ�id = [1] And a.�������� = 3 And a.Id = b.��id And b.�������� = 2 "
    Else
        
        lngDataFrom = rffSample
        lngFileId = lngFmtId
            
        '�ӷ����в�ѯ��ʽ����
        strSQL = "Select  Id As ���Id From ������������ a " & _
                    " Where �ļ�id = [1] And �������� = 3 And Substr(��������, Instr(��������, ';', 1, 18) + 1, 1) = '2' " & _
                    " Order By �������"
                    
        strPicSql = "select ID,�ļ�ID,��ID,1 as ��ʼ��,������,��������,�����д� from ������������ where  �ļ�ID=[1] and ��ID=[2] and ��������=5 order by ������"
        
        strContextSql = "Select a.�����ı� As ����, b.��������, b.�����ı� As ����" & vbNewLine & _
                " From ������������ a, ������������ b" & vbNewLine & _
                " Where a.�ļ�id = [1] And a.�������� = 3 And a.Id = b.��id And b.�������� = 2"
    End If
    
    
    '��ȡ�����ı�����****************************************
    Set rsData = zlDatabase.OpenSQLRecord(strContextSql, "��ѯ�����ı�", lngFileId)
    
    strSource���� = rtb����.Text
    strSource��� = rtb���.Text
    strSource���� = rtb����.Text
    
    rtb����.Text = ""
    rtb���.Text = ""
    rtb����.Text = ""
    
    blnHas���� = False
    blnHas��� = False
    blnHas���� = False
    
    While rsData.EOF = False
        strTmp = nvl(rsData!��������)
                           
        Select Case nvl(rsData!����)
            Case "�������"
                ReadReport nvl(rsData!����), strTmp, rtb����
                blnHas���� = True
                
            Case "������"
                ReadReport nvl(rsData!����), strTmp, rtb���
                blnHas��� = True
                
            Case "����"
                ReadReport nvl(rsData!����), strTmp, rtb����
                blnHas���� = True
                
        End Select
        
        rsData.MoveNext
    Wend
    
    If blnHas���� = False And blnHas��� = False And blnHas���� = False Then
        picChar.Visible = False
        MsgboxH GetRootHwnd, "��Ч�ı����ʽ����,���ܽ����л���", vbOKOnly, "��ʾ"
        
        '�ָ����л�ǰ���ı�����
        rtb����.Text = strSource����
        rtb���.Text = strSource���
        rtb����.Text = strSource����
        
        Exit Function
    Else
        dkpMain.Panes(2).Closed = Not blnHas����
        dkpMain.Panes(3).Closed = Not blnHas���
        dkpMain.Panes(4).Closed = Not blnHas����
    End If
    
    mlngSampleId = lngFmtId
    
    rtb����.Enabled = blnHas����
    rtb���.Enabled = blnHas���
    rtb����.Enabled = blnHas����
    
    
    
    '��ȡ����ͼ��Ϣ****************************************
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ����ͼ��", lngFileId)
    
    dcmReportImg.Visible = False
    dcmMarkImage.Visible = False
    
    blnReportVisible = False
    blnMarkVisible = False
    
    dcmReportImg.Images.Clear
    dcmMarkImage.Images.Clear

    If rsData.RecordCount > 0 Then
        '��ȡ���ͼ������ͼ
        
        blnReportVisible = True
        dcmReportImg.Visible = True
        
        'ͼ������ѯ
        dcmReportImg.tag = Val(nvl(rsData!���ID))
        
        Set rsData = zlDatabase.OpenSQLRecord(strPicSql, "��ѯ����ͼƬ", lngFileId, Val(nvl(rsData!���ID)))
        If rsData.RecordCount > 0 Then
            
            Call ParshReportImgData(rsData, lngDataFrom)
            
            If dcmMarkImage.Images.Count > 0 Then blnMarkVisible = True
            
            mblnIsModifyMarks = True
        End If
    End If
    
    
    If blnReportVisible = False And blnMarkVisible = False Then
        '�رձ���ͼ
        dkpMain.Panes(1).Closed = True
    Else
        '�򿪱���ͼ
        dkpMain.Panes(1).Closed = False
        
        If blnMarkVisible = False Then
            dcmReportImg.Width = picImageBack.Width
            ucSplitter1.Visible = False
        Else
            If dcmReportImg.Width = picImageBack.Width Then
                If picImageBack.Width - dcmMarkImage.Width < 0 Then
                    dcmMarkImage.Width = 0.34 * picImageBack.Width
                End If
                
                dcmReportImg.Width = picImageBack.Width - dcmMarkImage.Width
                
                ucSplitter1.Left = dcmReportImg.Width
                ucSplitter1.RePaint
            End If
            
            ucSplitter1.Visible = True
        End If
    End If
    
    
    mblnIsModifyText = True
    mlngReportID = 0
    
    ChangeReportFormat = True
End Function

Public Function SignVerifiy(ByVal lngSignVer As Long) As Boolean
'ǩ����֤
'------------------------------------------------
'���ܣ�У���鱨��ĵ���ǩ��(�ɶ���ת�Ƶ�����),У��汾Ϊintǩ���汾 ��ǩ��
'������ intǩ���汾 -- ������Ҫ��֤��ǩ���İ汾
'       blnMoved -- �����Ƿ�Ǩ��
'���أ�
'-----------------------------------------------
    Dim strSource As String
    Dim dblǩ��ID  As Double                  'ǩ�����ڵ��е�ID
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errhandle
    
    SignVerifiy = False
    
    '���ݱ���ID��ǩ���汾����ǩ������
    strSQL = "Select Id , ��ʼ�� From ���Ӳ������� Where �ļ�ID = [1] And �������� = 8 and ��ʼ�� =[2] "
    If mblnIsMoved Then
        strSQL = Replace(strSQL, "���Ӳ�������", "H���Ӳ�������")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���ǩ���汾", mlngReportID, lngSignVer)
    If rsTemp.RecordCount = 0 Then
        MsgboxH GetRootHwnd, "���α���û�а汾Ϊ" & lngSignVer & "��ǩ�����޷�������ǩ������֤��", vbInformation, gstrSysName
        Exit Function
    End If
    
    dblǩ��ID = Val(rsTemp!ID)
    
    '��ȡԴ��
    strSource = GetSignSource(mlngReportID, lngSignVer, mblnIsMoved)
    
    '������صĹ���=0����ʾ��ȡԴ��ʧ��
    If Len(strSource) = 0 Then
        MsgboxH GetRootHwnd, "���α���汾Ϊ" & lngSignVer & "��ǩ��Դ����ȡʧ�ܣ��޷�������ǩ������֤��", vbOK, "��ʾ"
        Exit Function
    End If
    
    '����ǩ�����󣬶�Դ�Ľ���ǩ����֤
    On Error Resume Next
    If gobjESign Is Nothing Then
        Set gobjESign = Interaction.GetObject(, "zl9ESign.clsESign")
        If gobjESign Is Nothing Then Set gobjESign = DynamicCreate("zl9ESign.clsESign", "����ǩ��")
        If err <> 0 Then err = 0
        
        If Not gobjESign Is Nothing Then
            Call gobjESign.Initialize(gcnOracle, glngSys)
        End If
    End If
        
    On Error GoTo errhandle
        
    If Not gobjESign Is Nothing Then
        'ǩ����֤
        Call gobjESign.VerifySignature(strSource, dblǩ��ID, 2)
        
        SignVerifiy = True
    End If
Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Private Function GetSignSource(ByVal lngReportID As Long, ByVal intǩ���汾 As Integer, ByVal blnMoved As Boolean) As String
'------------------------------------------------
'���ܣ���ȡ���ڵ���ǩ����ǩ����֤�ı���Դ������
'������ int��ȡ���� -- 1��ǩ��ʱ��ȡԴ�ģ�2��ǩ����֤ʱ��ȡԴ��
'       lngReportID -- ����ID�����Ӳ�����¼ID
'       intǩ���汾 -- ����ǩ��/��֤ǩ����ȡԴ�ĵİ汾��
'       blnMoved --- ���������Ƿ��Ѿ�ת��
'       thisSign --- ǩ������ǩ����ʱ����˶�����֤ǩ����ʱ����nothing
'       strSourceOut -- �����ء�ǩ��Դ��
'���أ� ǩ��/��֤ǩ����Դ�����ɹ���
'-----------------------------------------------
    Dim intRule As Integer
    Dim lngǩ��ID  As Long                  'ǩ�����ڵ��е�ID
    Dim strSQL As String
    Dim rs������¼ As ADODB.Recordset
    Dim rs�������� As ADODB.Recordset
    Dim rsǩ����¼ As ADODB.Recordset
    Dim rsData As ADODB.Recordset
    Dim strǩ��ʱ�� As String
    Dim arr��������() As String
    Dim strSignName As String
    Dim strSignImgBase64 As String
    Dim strImgFiles As String
    Dim strSourceOut As String
    Dim lngImgAdviceId As Long
    
    'Դ����ȡ����
    'intRule = 1ʱ����ȡ ID������ID��Ӥ���������ˣ�����ʱ�䣬ҽ��������ǩ������ǩ��ʱ��,���������������������
    '��֤ǩ����ʱ��ҽ��������ǩ������ǩ��ʱ���ǩ����¼�л�ȡ���ֱ���ҽ������= �������ı�����ǩ������=��Ҫ�ر�ʾ����ǩ��ʱ�� =���������ԣ�5����
    'ǩ����ʱ��ҽ��������ǩ������ǩ��ʱ�� ��ǩ�������л�ȡ
    On Error GoTo err
    
    If lngReportID = 0 Or intǩ���汾 = 0 Then Exit Function
    
    strSourceOut = ""
     
    '�ӵ��Ӳ�����¼����ȡ����Դ�ĵĻ�����Ϣ
    strSQL = "Select ID,����ID,Ӥ��,������,����ʱ�� From ���Ӳ�����¼ Where Id = [1]"
    If blnMoved Then strSQL = Replace(strSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
    
    Set rs������¼ = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����Դ�Ļ�����Ϣ", lngReportID)
    
    If rs������¼.RecordCount <= 0 Then Exit Function
    
    '�ӵ��Ӳ�����������ȡ����Դ�ĵ�������Ϣ
    strSQL = "Select a.�����ı� As ����, b.��������, b.�����ı� As ����,b.��ʼ�� as �汾 From ���Ӳ������� a,���Ӳ������� b " & _
             " Where a.�ļ�id = [1] And a.�������� = 3 And a.Id = b.��ID And b.�������� = 2 and b.��ʼ�� = [2]  "
    If blnMoved Then strSQL = Replace(strSQL, "���Ӳ�������", "H���Ӳ�������")
    
    Set rs�������� = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����Դ��������Ϣ", lngReportID, intǩ���汾)
    
    If rs��������.RecordCount = 0 Then Exit Function
    
     
    '��֤ǩ������ǩ����¼����ȡҽ��������ǩ������ǩ��ʱ����Ϣ,ǩ������
    strSQL = "Select �����ı� as ҽ������ ,Ҫ�ر�ʾ  as ǩ������ ,�������� From ���Ӳ������� Where �ļ�ID = [1] And �������� = 8 and ��ʼ�� =[2] "
    If blnMoved Then strSQL = Replace(strSQL, "���Ӳ�������", "H���Ӳ�������")
    
    Set rsǩ����¼ = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��󱨸�Դ��ǩ����Ϣ", lngReportID, intǩ���汾)
    
    If rsǩ����¼.RecordCount = 0 Then Exit Function
    
    '��ȡ��ʽ����ǩ��ʱ�䣬ǩ������
    arr�������� = Split(rsǩ����¼!��������, ";")
    If UBound(arr��������) >= 5 Then
        intRule = Val(arr��������(1))
        strǩ��ʱ�� = Format(arr��������(4), "yyyy-MM-dd HH:mm:ss")
    End If
    If intRule = 0 Then Exit Function
    
    '���ݹ�����֯����Դ�ģ� ID������ID��Ӥ���������ˣ�����ʱ�䣬ҽ��������ǩ������ǩ��ʱ��,���������������������
    If intRule = 1 Then
        'Դ�Ļ�����Ϣ
        strSourceOut = rs������¼!ID
        strSourceOut = strSourceOut & vbTab & nvl(rs������¼!����ID)
        strSourceOut = strSourceOut & vbTab & nvl(rs������¼!Ӥ��)
        strSourceOut = strSourceOut & vbTab & nvl(rs������¼!������)
        strSourceOut = strSourceOut & vbTab & nvl(rs������¼!����ʱ��)
 
        '��֤ǩ���������ݿ�ǩ����¼��ȡ
        strSignName = nvl(rsǩ����¼!ҽ������)
        If InStr(strSignName, M_STR_TAG_SIGNWITHIMG) > 0 Then
            strImgFiles = Split(strSignName, M_STR_TAG_SIGNWITHIMG)(1)
            strSignName = Split(strSignName, M_STR_TAG_SIGNWITHIMG)(0)
        End If

        strSourceOut = strSourceOut & vbTab & strSignName   '����
        strSourceOut = strSourceOut & vbTab & nvl(rsǩ����¼!ǩ������)
        strSourceOut = strSourceOut & vbTab & strǩ��ʱ��
  
        
        'Դ�ı�������
        rs��������.Filter = "���� ='" & ReportViewType_������� & "'"
        If rs��������.RecordCount = 0 Then
            strSourceOut = strSourceOut & vbTab
        Else
            strSourceOut = strSourceOut & vbTab & nvl(rs��������!����)
        End If
        
        rs��������.Filter = "���� ='" & ReportViewType_������ & "'"
        If rs��������.RecordCount = 0 Then
            strSourceOut = strSourceOut & vbTab
        Else
            strSourceOut = strSourceOut & vbTab & nvl(rs��������!����)
        End If
        
        rs��������.Filter = "���� ='" & ReportViewType_���� & "'"
        If rs��������.RecordCount = 0 Then
            strSourceOut = strSourceOut & vbTab
        Else
            strSourceOut = strSourceOut & vbTab & nvl(rs��������!����)
        End If
        
        'Դ��ǩ��ͼ����Ϣ
        If mblnUseImgSign Then
            '�����ݿ�ǩ����¼��ȡ
            lngImgAdviceId = Val(Split(strImgFiles & "[ADV]", "[ADV]")(1))
            
            If lngImgAdviceId <= 0 Then
                strSignImgBase64 = GetSignedImgB64(strImgFiles, mlngAdviceId, blnMoved)
            Else
                '�ж�ͼ���Ӧ��ҽ���Ƿ��Ѿ���ת��
                strSQL = "select ID From ����ҽ����¼ where Id=[1]"
                Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯҽ����¼", lngImgAdviceId)
                
                If rsData.RecordCount > 0 Then
                    strSignImgBase64 = GetSignedImgB64(Split(strImgFiles, "[ADV]")(0), lngImgAdviceId, False)
                Else
                    strSignImgBase64 = GetSignedImgB64(Split(strImgFiles, "[ADV]")(0), lngImgAdviceId, True)
                End If
            End If
            If Len(strSignImgBase64) <= 0 Then Exit Function
            
            strSourceOut = strSourceOut & vbTab & strSignImgBase64
        End If
    End If
    
    GetSignSource = strSourceOut
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function GetSignedImgB64(ByVal strSignImgFiles As String, ByVal lngImgAdviceId As Long, ByVal blnMoved As Boolean) As String
'��ȡͼ���base64
    Dim i As Long
    Dim aryFile() As String
    Dim strLocalFile As String
    Dim strCurB64 As String
    
    GetSignedImgB64 = ""
    aryFile = Split(strSignImgFiles & ";", ";")
    For i = 0 To UBound(aryFile)
        If Len(aryFile(i)) > 0 Then
            strLocalFile = mstrReportImgPath & aryFile(i)
            If FileExists(strLocalFile) = False Then
                If DownLoadFtpFile(lngImgAdviceId, aryFile(i), strLocalFile, blnMoved) <> frNormal Then
                    Exit Function
                End If
            End If
            
            If FileExists(strLocalFile) = False Then
                GetSignedImgB64 = ""
                MsgboxH GetRootHwnd, "δ��ȡ���ð汾��Ӧ��ǩ��ͼ�񣬲�����֤��", vbOKOnly, "��ʾ"
                Exit Function
            End If
            
            strCurB64 = zlStr.EncodeBase64_File(strLocalFile)
            
            If GetSignedImgB64 <> "" Then GetSignedImgB64 = GetSignedImgB64 & ";"
            GetSignedImgB64 = GetSignedImgB64 & strCurB64
        End If
    Next

End Function


Public Function SignUntread() As Boolean
'ǩ������
    Dim signInfo As TReportSignInfo
    Dim strSQL As String
    Dim arrSQL() As String
    Dim blIsUntread As Boolean
    Dim intRobackType As Integer '����ǩ������
    Dim i As Long
    
    SignUntread = False
 
    
    If Val(labSign.tag) = 1 Then  'ֻ��һ��ǩ������ʾ��ǰ����дģʽ�µĻ���
        signInfo = frmEPRUntread.ShowUntread(mlngReportID, cprET_�������༭, Me)
    Else
        signInfo = frmEPRUntread.ShowUntread(mlngReportID, cprET_���������, Me)
    End If
    
    If signInfo.ID <= 0 Then Exit Function
 
    If MsgboxH(GetRootHwnd, "ע�⣺���˲��������ɻָ����Ƿ������", vbYesNo + vbDefaultButton2 + vbQuestion, "��ʾ") = vbNo Then Exit Function

    mblnIsLoadData = False '�ڵ���loadreport����ǰ����Ҫ���ô˱���Ϊfalse�������������ݺ��ı����ݵ��޸�״̬Ϊtrue
    
    '�������ֻ��˷�ʽ
    If signInfo.Key > 0 Then
        ReDim arrSQL(1)

        '���ǩ��,�������ʽ
        If SaveSignFormat(mlngAdviceId, mlngReportID, signInfo, "", True) = False Then Exit Function
    ElseIf signInfo.ǩ���汾 > 1 Then  '�����޶�
        'ֱ���޸����ݿ����ݾͿ�����  '�ѻ����޶����浽���ݿ�
        strSQL = "ZL_Ӱ�񱨸����(0," & mlngReportID & "," & signInfo.ǩ���汾 & ")"
        zlDatabase.ExecuteProcedure strSQL, "���˱���ǩ��"
    End If
    
    Call ResetContext
    
    Call LoadReport
     
    '����ר�Ʊ���
    If Not mobjSpePlugin Is Nothing Then
        
        If mblnIsSpeState Then
            '�ָ���ר����ʾ����
            Call ChangeSepState(True, True)
            
        End If
    End If
    
    mblnIsLoadData = True 'loadreport����������ɺ�����Ϊtrue
    
    SignUntread = True
End Function


Public Function Sign() As Long
'ǩ�� 0-δǩ����1-���ǩ����2-���ǩ��
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim objSignForm As frmEPRSign
    Dim strImgBase64Code As String
    Dim strImgFiles As String
    Dim signInfo As TReportSignInfo
    Dim strRtfFile As String
    Dim lngImgAdviceId As Long
    
On Error GoTo errhandle
    
    
    Sign = 0
    If Not IsHaveContent() Then
        MsgBoxD Me, "û����Ч�ı������ݣ�������ǩ����", vbInformation, gstrSysName
        Exit Function
    End If
    
    'ɾ��rtf��ʽ�ļ�
    strRtfFile = FormatFilePath(SysRootPath & "\TMP.RTF")
    If Dir(strRtfFile) <> "" Then Call RemoveFile(strRtfFile)
    
    'ǩ��֮ǰ�ȱ��汨��
    If IsModify Then
        If SaveReport() = False Then Exit Function
    Else
        mintTargetVer = mlngSignCount + 1
    End If
    
    If mlngReportID = 0 Then
        MsgboxH GetRootHwnd, "δ�ҵ���Ӧ������Ϣ�����ܽ���ǩ����", vbOKOnly, "��ʾ"
        Exit Function
    End If
    
    If mintTargetVer >= 16 Then
        MsgboxH GetRootHwnd, "Ŀǰϵͳ֧�ֵ����ǩ���汾��Ϊ16������˻�����������", vbOKOnly, "��ʾ"
        Exit Function
    End If
    
    '��ȡǩ��ͼ��
    If mblnUseImgSign Then
        '�м��ͼ������²�������ж�ͼ�����ǩ��
        If dcmReportImg.Images.Count > 0 Then
            lngImgAdviceId = 0
            
            If GetSignImgEncode(mlngReportID, mintTargetVer, strImgFiles, strImgBase64Code, lngImgAdviceId) = False Then Exit Function
            If lngImgAdviceId = 0 Then lngImgAdviceId = mlngAdviceId
            
            If StorageSignImg(lngImgAdviceId, mlngReportID, mintTargetVer) = False Then Exit Function
        Else
            '���������Ա��ͼ����ǩ��
            If dcmMarkImage.Images.Count > 0 Then
                If MsgboxH(GetRootHwnd, "��ǰ����û�б���ͼ�񣬲��ܵ����Ա��ͼ����ǩ�����Ƿ������", vbYesNo, "��ʾ") = vbNo Then Exit Function
            End If
        End If
    End If
    
    If mstrFinalSignUser = UserInfo.���� Then
        If MsgboxH(GetRootHwnd, "[" & mstrFinalSignUser & "] �û��ѽ���ǩ�������Ƿ������", vbYesNo, "��ʾ") = vbNo Then Exit Function
    End If
    
    Set objSignForm = New frmEPRSign
    
    signInfo = objSignForm.ShowSign(UserControl.Parent, mlngSignPassType, mlngReportID, _
                                    mstrPrivs, mlngSignLevel, mstrFirstSignUser, mintTargetVer, _
                                    strImgBase64Code, mlngAdviceId)
    '���ǩ����Ϣ��ȡʧ�ܣ����˳�ǩ��
    If signInfo.ǩ����ʽ = 0 Or (signInfo.ǩ����ʽ = 2 And Len(signInfo.ǩ����Ϣ) <= 0) Then Exit Function
    
    
    'ǩ����ʽ����ɹ�����Ҫ�����汾��
    If SaveSignFormat(mlngAdviceId, mlngReportID, signInfo, strImgFiles) Then
    
        Sign = IIf(signInfo.ǩ������ > 1, 2, 1)
        
        mstrFinalSignUser = UserInfo.����
        If mstrFirstSignUser = "" Then mstrFirstSignUser = UserInfo.����
        
        mlngSignCount = mlngSignCount + 1
        mlngSignLevel = signInfo.ǩ������
        mintSourceVer = mintSourceVer + 1
        
        labSign.Caption = labSign.Caption & "  " & UserInfo.����
        
        'ǩ����������1
        labSign.tag = mlngSignCount
        
        Call ConfigFaceState
    End If
    
Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function CreateEditor() As Object
'����Editor�ؼ�
    Dim objEditor As Object
    
    Set objEditor = Controls.Add("zlRichEditor.Editor", "Editor")
    objEditor.Visible = False 'ʹ�ؼ��ɼ�
    
    Set CreateEditor = objEditor
End Function

Private Function RemoveEditor(objEditor As Object)
'�Ƴ�Editor�ؼ�
    Controls.Remove objEditor
    Set objEditor = Nothing
End Function



Private Function SaveSignFormat(ByVal lngAdviceId As Long, ByVal lngReportID As Long, signInfo As TReportSignInfo, _
    ByVal strSignFiles As String, Optional ByVal blnIsUntread As Boolean = False) As Boolean
'------------------------------------------------
'���ܣ����汨���ʽRTF�ļ����Ա������ǩ�����߻���
'������     OneSign -- ��Ϊ�գ����ʾ����ǩ�����߻��ˣ�Ϊ�գ���ʾֻ�Ǳ����ʽ��������ǩ��
'           blnAddSign ���ӻ��߻���ǩ����True--����ǩ��,OneSignΪ�ձ�ʾ���汨���ʽ��False--����ǩ��
'���أ� �ޣ�ֱ�ӱ���RTF�����ʽ�ĵ����Ա���ǩ�����߻���
'-----------------------------------------------
On Error GoTo errH
    Dim strZipFile As String
    Dim i As Long
    Dim strTemp As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strReport As String
    Dim lngSignPos As Long
    Dim strReportFormatFile As String
    Dim strErrCount As String
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long
    Dim bFinded As Boolean, sKeyType As String, bNeeded As Boolean
    Dim objEditor As Object
    Dim arySql() As String
    Dim blnInTrans As Boolean

    SaveSignFormat = False
    
    strErrCount = ""
    If signInfo.ǩ����ʽ = 0 Or (signInfo.ǩ����ʽ = 2 And Len(signInfo.ǩ����Ϣ) <= 0) Then Exit Function
    
reLoad:
    strReportFormatFile = FormatFilePath(SysRootPath & "\TMP.RTF")
    If Dir(strReportFormatFile) <> "" Then
        '���ش��ڱ��汣��ʱ��rtf��ʽ�ļ�ʱ������Ҫ���¶�ȡ
        strTemp = strReportFormatFile
    Else
        strReportFormatFile = FormatFilePath(SysRootPath & "\SignFmt" & strErrCount)
        
        '�ȸ��Ʊ����ʽ
        If Dir(strReportFormatFile) <> "" Then Kill strReportFormatFile
        
        '�����ݿ��ȡRTF�����ʽ�ĵ�
        strZipFile = Sys.ReadLob(glngSys, 5, lngReportID, strReportFormatFile)
        
        '��ѹ���ļ�
        strTemp = zlFileUnzip(strZipFile)
    End If
    

    ReDim arySql(0)
    
    If strTemp <> "" Then
        If FileExists(strTemp) = False Then
            If MsgboxH(GetRootHwnd, "����IDΪ[" & lngReportID & "]��RTF��ʽ�ļ���ȡʧ�ܣ��Ƿ����ԣ�", vbYesNo) = vbYes Then
                strErrCount = CStr(Val(strErrCount) + 1)
                GoTo reLoad
            End If
            
            Exit Function
        End If
         
        Set objEditor = CreateEditor()
        objEditor.OpenDoc strTemp

        If blnIsUntread Then
            '����ǩ��
            Call DeleteFromEditor(objEditor, signInfo)
            
            '�ѻ���ǩ�����浽���ݿ�
            strSQL = "ZL_Ӱ�񱨸����(" & signInfo.ID & "," & lngReportID & ",0)"
            
            ReDim Preserve arySql(UBound(arySql) + 1)
            arySql(UBound(arySql)) = strSQL
        Else
            '����д��ǩ����λ��
            strSQL = "Select ������ From ���Ӳ������� Where �ļ�ID= [1] And �������� = 4 And Ҫ������ ='����ǩ��' "
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ����ǩ��λ", lngReportID)
            lngSignPos = -1
            If rsTemp.EOF = False Then
                bFinded = FindKey(objEditor, "E", nvl(rsTemp!������, 0), lKSS, lKSE, lKES, lKEE, bNeeded)
                If bFinded = True Then lngSignPos = lKEE
            End If
            
            '��ָ��λ��д��ǩ��
            InsertIntoEditor objEditor, signInfo, lngSignPos
            
            '��ǩ�����浽���ݿ�
            strSQL = "ZL_Ӱ�񱨸�ǩ��_����(" & lngReportID & "," & _
                    signInfo.��ʼ�� & "," & signInfo.��ֹ�� & " ,'" & signInfo.�������� & "','" & _
                    signInfo.���� & strSignFiles & "','" & signInfo.ǰ������ & "','" & signInfo.ʱ��� & "'," & signInfo.ǩ������ & ",'" & signInfo.ǩ����Ϣ & "')"
            
            ReDim Preserve arySql(UBound(arySql) + 1)
            arySql(UBound(arySql)) = strSQL
            
            
            '���µ��Ӳ�����¼�е����汾
            strSQL = "ZL_Ӱ�񱨸�����_update(" & lngAdviceId & "," & _
                                                lngReportID & ",'',''," & _
                                                signInfo.��ʼ�� & "," & _
                                                signInfo.ǩ������ & ")"
                                                
            ReDim Preserve arySql(UBound(arySql) + 1)
            arySql(UBound(arySql)) = strSQL
        End If
        
        
        '�������ʱ�ļ�
        objEditor.SaveDoc strTemp
        
        'ѹ���ļ�
        strZipFile = zlFileZip(strTemp)
        
        '�����ʽ
        zlSaveLob 5, lngReportID, strZipFile, arySql
    
        'ɾ����ʱzip�ļ�
        Kill strZipFile
        
        RemoveEditor objEditor
        
        '����д���ʽ
        gcnOracle.BeginTrans        '----------���汨������
        blnInTrans = True
        For i = 0 To UBound(arySql)
            If Trim(arySql(i)) <> "" Then
                Call zlDatabase.ExecuteProcedure(CStr(arySql(i)), IIf(blnIsUntread, "���˱���ǩ��", "���汨��ǩ��[" & i & "]"))
            End If
        Next i
        gcnOracle.CommitTrans
        
        SaveSignFormat = True
    Else
        If MsgboxH(GetRootHwnd, "�޷���ȡ���߽�ѹ�����ʽ" & strReportFormatFile & vbCrLf & "��ʹ�á������༭���ķ������༭�˱�������Զ�ȡ���Ƿ����ԣ�", vbYesNo) = vbYes Then
            If Dir(strReportFormatFile) <> "" Then Kill strReportFormatFile
            
            strErrCount = CStr(Val(strErrCount) + 1)
            GoTo reLoad
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
    If blnInTrans Then gcnOracle.RollbackTrans
End Function


Private Function GetSignMarkImgName(ByVal lngReportID As Long, ByVal intSignVer As Integer) As String
    GetSignMarkImgName = "MarkImage_" & lngReportID & "_" & intSignVer & ".JPG"
End Function

Private Function StorageSignImg(ByVal lngImgAdviceId As Long, ByVal lngReportID As Long, ByVal intSignVer As Integer) As Boolean
'����ǩ��ͼ��
    Dim strFile As String
    Dim strFileName As String
    Dim reportImgTag As TReportImgTag
    
    StorageSignImg = True
    If dcmMarkImage.Images.Count <= 0 Then Exit Function
    
    strFileName = GetSignMarkImgName(lngReportID, intSignVer)
    strFile = mstrReportImgPath & strFileName
    
    If FileExists(strFile) = False Then
        Call dcmMarkImage.Images(1).FileExport(strFile, "JPG")
    End If
    
    reportImgTag = dcmMarkImage.Images(1).tag
    
    If UpLoadFtpFile(lngImgAdviceId, strFileName, strFile, False) <> frNormal Then
        StorageSignImg = False
    End If
    
End Function


Private Function GetSignImgEncode(ByVal lngReportID As Long, ByVal intSignVer As Integer, _
    ByRef strImgFiles As String, ByRef strBase64Code As String, ByRef lngImgAdviceId As Long) As Boolean
'����ͼ��ǩ����Ϣ
'���ظ�ʽΪ
    Dim strErr As String
On Error GoTo errhandle
    Dim i As Integer
    Dim strFile As String
    Dim strFileName As String
    Dim strResult As String
    Dim strCurB64 As String
    Dim reportImgTag As TReportImgTag
    
    GetSignImgEncode = True
    
    strImgFiles = ""
    strBase64Code = ""
    
    If dcmMarkImage.Images.Count <= 0 And dcmReportImg.Images.Count <= 0 Then Exit Function

    strResult = ""
    
    '������ͼ
    If dcmMarkImage.Images.Count > 0 Then
        '"���ͼ_" & reportImgTag.lngFileId & "_" & reportImgTag.strKey & ".JPG"
        strFileName = GetSignMarkImgName(lngReportID, intSignVer)
        strFile = mstrReportImgPath & strFileName
        Call dcmMarkImage.Images(1).FileExport(strFile, "JPG")
        
        strCurB64 = zlStr.EncodeBase64_File(strFile)
        If Len(strCurB64) > 0 Then
            strResult = M_STR_TAG_SIGNWITHIMG & strFileName
            
            If strBase64Code <> "" Then strBase64Code = strBase64Code & ";"
            strBase64Code = strBase64Code & strCurB64
        Else
            GetSignImgEncode = False
            MsgboxH GetRootHwnd, "���ͼתBase64ʧ�ܣ����ܽ���ͼ��ǩ����", vbOKOnly, "��ʾ"
            Exit Function
        End If
    End If
    
    lngImgAdviceId = 0
    
    '������ͼ
    For i = 1 To dcmReportImg.Images.Count
        reportImgTag = dcmReportImg.Images(i).tag
        
        If lngImgAdviceId = 0 Then lngImgAdviceId = reportImgTag.lngFromAdvice
        
        strFileName = reportImgTag.strImgFile
        
        strFile = mstrReportImgPath & strFileName
        
        If FileExists(strFile) = False Then
            Call dcmReportImg.Images(i).FileExport(strFile, "JPG")
        End If
        
        strCurB64 = zlStr.EncodeBase64_File(strFile)
        If Len(strCurB64) > 0 Then
            If strResult = "" Then
                strResult = M_STR_TAG_SIGNWITHIMG
            Else
                strResult = strResult & ";"
            End If
        
            strResult = strResult & strFileName
            If strBase64Code <> "" Then strBase64Code = strBase64Code & ";"
            strBase64Code = strBase64Code & strCurB64
        Else
            GetSignImgEncode = False
            MsgboxH GetRootHwnd, "����ͼתBase64ʧ�ܣ����ܽ���ͼ��ǩ����", vbOKOnly, "��ʾ"
            Exit Function
        End If
    Next
    
    If lngImgAdviceId <> 0 Then
        strResult = strResult & "[ADV]" & lngImgAdviceId
    End If
    
    strImgFiles = strResult
    
    Exit Function:
errhandle:
    GetSignImgEncode = False
    strErr = err.Description
    
    MsgboxH GetRootHwnd, "ͼ��Base64ת�����󣬲��ܽ���ǩ����" & vbCrLf & strErr, vbOKOnly, "��ʾ"
End Function

Public Sub ReportPreview(ByVal strReportNo As String, ByVal strPrintFmts As String)
'����Ԥ��
    Call PrintReport(False, strReportNo, strPrintFmts)
End Sub

Public Function ReportPrint(ByVal strReportNo As String, ByVal strPrintFmts As String, Optional ByVal blnIsBat As Boolean = False) As Boolean
'�����ӡ
    ReportPrint = PrintReport(True, strReportNo, strPrintFmts, blnIsBat)
End Function


Private Function GetReportId(ByVal lngAdviceId As Long, ByVal blnIsMoved As Boolean, ByRef lngFileFmtId As Long, Optional ByVal lngSpecifyReportId As Long = 0) As Long
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    GetReportId = 0
    
    If lngSpecifyReportId = 0 Then
        strSQL = "select a.����ID, b.�ļ�ID from ����ҽ������ a, ���Ӳ�����¼ b where a.����id=b.id and a.ҽ��ID=[1]"
        If blnIsMoved Then strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���ҽ������", lngAdviceId)
    Else
        strSQL = "select a.����ID, b.�ļ�ID from ����ҽ������ a, ���Ӳ�����¼ b where a.����id=b.id and a.����ID=[1]"
        If blnIsMoved Then strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���ҽ������", lngSpecifyReportId)
    End If
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    lngFileFmtId = Val(nvl(rsData!�ļ�ID))
    GetReportId = Val(nvl(rsData!����Id))
End Function


Public Function ReportPreviewEx(ByVal lngAdviceId As Long, ByVal blnIsMoved As Boolean, _
    Optional ByVal lngSpecifyReportId As Long = 0, Optional ByVal blnIsOneFmt As Boolean = False) As Boolean
    Dim lngReportID As Long
    Dim lngFmtId As Long
    Dim strReportNo As String
    Dim strPrintFmt As String
     
    ReportPreviewEx = False
    
    lngReportID = GetReportId(lngAdviceId, blnIsMoved, lngFmtId, lngSpecifyReportId)
    
    If lngReportID <= 0 Then
        MsgboxH GetRootHwnd, "δ�ҵ��ɴ�ӡ�ļ�鱨�棬��ȷ�ϱ����Ƿ���д��", vbOKOnly, "��ʾ"
        Exit Function
    End If
    
    Call Refresh(lngAdviceId, lngFmtId, 0, lngReportID, blnIsMoved)
    
    If GetPrintFormat(mObjNotify.Owner, lngFmtId, strReportNo, strPrintFmt, blnIsOneFmt) = False Then Exit Function
    
    Call ReportPreview(strReportNo, strPrintFmt)
End Function


Private Function GetPrintFormat(Owner As Object, ByVal lngFileId As Long, _
    ByRef strReportNo As String, ByRef strPrintFmt As String, Optional ByVal blnIsOneFmt As Boolean = False) As Boolean
'��ʼ�������ӡ��ʽ
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strRegReportNo As String
    Dim blnCancel As Boolean
      
    strReportNo = ""
'    strPrintFmt = ""
    GetPrintFormat = True
    
    '���ж��Ƿ�ʹ���Զ��屨��
    strSQL = "Select ͨ��,��� From �����ļ��б�  Where Id =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�����ӡ��ʽ", lngFileId)
    
    If rsTemp.RecordCount <= 0 Then Exit Function
    If nvl(rsTemp!ͨ��) <> 2 Then Exit Function
    
    strReportNo = "ZLCISBILL" & Format(nvl(rsTemp!���), "00000") & "-2"
    
    '��������˸�ʽ��˵��ֻ��Ҫ��ȡ������
    If Len(strPrintFmt) > 0 Then
        If Split(strPrintFmt, ":")(0) = strReportNo Then
            strPrintFmt = Split(strPrintFmt & ":", ":")(1)
            Exit Function
        End If
    End If
        
    strSQL = "Select b.��� as ID, a.���, b.˵�� as ���� From zlreports a,zlrptfmts b Where a.Id=b.����ID And a.���=[1] Order By ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�Զ��屨���ʽ", strReportNo)
    
    If rsTemp.RecordCount <= 1 Then Exit Function
    
    MainForm.SetFocus
    
'    Call SetActiveWindow(GetRootHwnd)
    
    '�жϸ�ʽ�Ƿ������ѡ
    If blnIsOneFmt Then
        Set rsTemp = zlDatabase.ShowSQLSelect(Parent, strSQL, 0, "��ʽѡ��", True, "ID", "��ѡ����Ҫ��ӡ�ĸ�ʽ...", False, False, False, _
                                        Screen.Width / 2 - 3000, Screen.Height / 2 - 2000, 2000, blnCancel, True, False, strReportNo)
    Else
        '�����ʽ����һ�����򵯳���ʽѡ����
        Set rsTemp = zlDatabase.ShowSQLMultiSelect(Parent, strSQL, 0, "��ʽѡ��", True, "ID", "��ѡ����Ҫ��ӡ�ĸ�ʽ...", False, False, False, _
                                        Screen.Width / 2 - 3000, Screen.Height / 2 - 2000, 2000, blnCancel, True, False, strReportNo)
    End If
    
    If blnCancel Or rsTemp Is Nothing Then
        GetPrintFormat = False
        Exit Function
    End If
    
    If rsTemp.RecordCount <= 0 Then
        GetPrintFormat = False
        Exit Function
    End If
    
    While Not rsTemp.EOF
        strPrintFmt = strPrintFmt & Val(nvl(rsTemp!ID)) & ","
        Call rsTemp.MoveNext
    Wend
    
End Function

Public Function ReportPrintEx(ByVal lngAdviceId As Long, ByVal blnIsMoved As Boolean, _
    Optional ByVal lngSpecifyReportId As Long = 0, Optional ByVal blnIsOneFmt As Boolean = False, Optional ByVal strPrintFmts As String = "") As Boolean
    Dim lngReportID As Long
    Dim lngFmtId As Long
    Dim strReportNo As String
    Dim strPrintFmt As String
    
    ReportPrintEx = False
    
    lngReportID = GetReportId(lngAdviceId, blnIsMoved, lngFmtId, lngSpecifyReportId)
    
    If lngReportID <= 0 Then
        MsgboxH GetRootHwnd, "δ�ҵ��ɴ�ӡ�ļ�鱨�棬��ȷ�ϱ����Ƿ���д��", vbOKOnly, "��ʾ"
        Exit Function
    End If
    
    Call Refresh(lngAdviceId, lngFmtId, 0, lngReportID, blnIsMoved)
    
    If Len(strPrintFmts) > 0 Then strPrintFmt = strPrintFmts
    If GetPrintFormat(mObjNotify.Owner, lngFmtId, strReportNo, strPrintFmt, blnIsOneFmt) = False Then Exit Function
    
    
    ReportPrintEx = ReportPrint(strReportNo, strPrintFmt)
End Function


'Private Function IsAllowPrint() As Boolean
''�ж��Ƿ��������ӡ
'On Error GoTo errH
'    Dim strSQL As String
'    Dim rsTemp As ADODB.Recordset
'
'    strSQL = "Select a.������,a.������,b.������־ ,b.Id From Ӱ�����¼ a ,����ҽ����¼ b Where a.ҽ��id = b.Id And b.Id = [1] "
'    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��֤�Ƿ���Դ�ӡ", mlngAdviceID)
'
'    If rsTemp.EOF = False Then
'        IsAllowPrint = IIf(nvl(rsTemp!������־, 0) = 1, nvl(rsTemp!������) <> "", nvl(rsTemp!������) <> "")
'    Else
'        IsAllowPrint = False
'    End If
'
'    Exit Function
'errH:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Function


Private Function PrintReport(ByVal blnIsPrint As Boolean, _
    Optional ByVal strReportNo As String, Optional ByVal strPrintFmts As String, _
    Optional ByVal blnSilent As Boolean = False) As Boolean
'blnIsPrint:�Ƿ���˺��Զ���ӡ����
'blnSilent������ӡ����ʱ������Ҫ���ݴ˲���

On Error GoTo errH
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim blnUseCustomReport As Boolean
    Dim objRichEpr As Object
    Dim blnNoAsk As Boolean
    Dim objCusReport As Object
    
    PrintReport = False
    
    '�жϱ����Ƿ���Դ�ӡ
    '�������˺��ӡ���棬��ʱ���ݿ⻹δ�������ݣ����õ���chkPrintState�ж�
'    If mblnCheckPrintPara = True And blnIsPrint Then
'        If IsAllowPrint = False Then
'            MsgboxH GetRootHwnd, "��ǰ���治�����ӡ��", vbOKOnly, "��ʾ"
'            Exit Function
'        End If
'    End If
    
    If mlngReportID = 0 Then
        MsgboxH GetRootHwnd, "δ�ҵ����������Ϣ�����ȱ��汨�档", vbOKOnly, "��ʾ"
        Exit Function
    End If
    
    '��ӡԤ��ǰ����Ҫ�ж��Ƿ񱣴汨��
    If IsModify Then Call SaveReport

 
    '��ӡ�������Ԥ������
    If Len(strReportNo) > 0 Then
        '�Ƿ�Ĭ��ӡ
        blnNoAsk = (zlDatabase.GetPara("NoAsk", glngSys, 1070, 0) = "1")
        If blnSilent = True Then blnNoAsk = True
    
        Set objCusReport = DynamicCreate("zl9Report.clsReport", "�Զ��屨��")
        If objCusReport Is Nothing Then Exit Function
        
'        If Not blnNoAsk Then
            
            '��û�����ô�ӡ��ʱ���ᵯ����ӡ�����ô��ڣ������Ҫ����һ��Ĭ�ϵı����ʽ
            objCusReport.SetReportPrintSet gcnOracle, glngSys, strReportNo, "Format", Split(strPrintFmts & "-", "-")(0)
            
'            If objCusReport.ReportPrintSet(gcnOracle, glngSys, strReportNo) = False Then
'                '�˴�ˢ�»���ɽ������
'                Exit Function
'            End If
'
'            strPrintFmts = objCusReport.GetReportPrintSet(gcnOracle, glngSys, strReportNo, UserInfo.�û���, 1, , "Format")
'        End If

        
        objCusReport.InitOracle gcnOracle
        
        PrintReport = CustomReportPrint(objCusReport, mlngReportID, strPrintFmts, strReportNo, blnIsPrint)
        
        Call objCusReport.CloseWindows
        Set objCusReport = Nothing
        
    Else        'ʹ�ñ༭ģʽ��ӡ�����ò����Ĵ�ӡ����
        Set objRichEpr = DynamicCreate("zlRichEPR.cRichEPR", "���Ӳ���")
        If objRichEpr Is Nothing Then Exit Function
    
        objRichEpr.InitRichEPR gcnOracle, mObjNotify.Owner, glngSys, False
        Call objRichEpr.PrintOrPreviewDoc(mObjNotify.Owner, cpr���Ʊ���, mlngReportID, blnIsPrint, True)
        
        PrintReport = True
        
        Call objRichEpr.CloseWindows
        Set objRichEpr = Nothing
    End If
    
   
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function




Private Function CustomReportPrint(objCusReport As Object, ByVal lngReportID As Long, _
    ByVal strSelFmts As String, ByVal strReportNo As String, _
    ByVal blnPrint As Boolean) As Boolean
'ʹ���Զ��屨���ӡ��Ԥ������
'������ blnPrint---True��ӡ��FalseԤ��
'       blnSilent ---ǿ�ƾ�Ĭ��ӡ��������ӡʱ��Ҫ
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strExseNo As String, intExseKind As Integer
    Dim intPCount As Integer
    Dim i As Integer, j As Integer, intParaCount As Integer
    Dim strPicFile As String
    Dim aryRepPara(19) As String, aryMarkPara(1) As String     '����ͼ�е�ͼ���¼
    Dim aryPrintPara(19) As String, strFlagString As String 'ʵ�ʴ����Զ��屨�������
    Dim dcmMarkImages As New DicomImages
    Dim dcmRepImages As New DicomImages
    Dim dcmResultImage As DicomImage
    Dim arr�����ʽ() As String
    Dim int��ʽ�� As Integer
    Dim intRows As Integer, intCols As Integer
    Dim blnIsImageReport As Boolean
    Dim strPicSql As String
    Dim aryImgPro() As String
    Dim reportImgTag As TReportImgTag
    Dim lngReportBoxCount As Long
    Dim lngReportImgCount As Long
    Dim strFirstFile As String

On Error GoTo errhandle
 
    CustomReportPrint = False
    
    '��ȡ����ļ�¼���ʺ�No
    strSQL = "Select ��¼����, No From ����ҽ������ Where ҽ��id = [1]"
    If mblnIsMoved = True Then strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ǰ��¼���ʺ�No", mlngAdviceId)
    If rsTemp.RecordCount = 0 Then Exit Function

    strExseNo = "" & rsTemp!no
    intExseKind = Val("" & rsTemp!��¼����)


    '��ȡ����ͼ�񣨰������ͼ�����ɱ����ļ�
    'һ���������п������ж������ͼ
    strSQL = "Select Id As ���Id From ���Ӳ�������" & vbNewLine & _
        "       Where �ļ�id = [1] And �������� = 3 And Substr(��������, Instr(��������, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
        "       Order By �������"
    If mblnIsMoved = True Then strSQL = Replace(strSQL, "���Ӳ�������", "H���Ӳ�������")

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡͼ��", lngReportID)

    If rsTemp.RecordCount > 0 Then
        strPicSql = "select ID,�ļ�ID,��ID,��ʼ��,������,��������,�����д� from ���Ӳ������� where  �ļ�ID=[1] and ��ID=[2] and ��������=5 order by ������"
        Set rsTemp = zlDatabase.OpenSQLRecord(strPicSql, "��ѯԤ����ӡͼƬ", lngReportID, Val(nvl(rsTemp!���ID)))
        
        If rsTemp.RecordCount > 0 Then
            intPCount = 0
            Do While Not rsTemp.EOF
                aryImgPro = Split(nvl(rsTemp!��������) & ";;;;;;;;;;;;;;;;;;;;", ";")
                
                reportImgTag.lngFileId = Val(rsTemp!�ļ�ID)
                reportImgTag.lngTableId = Val(rsTemp!��ID)
                reportImgTag.strObjectTag = Val(rsTemp!������)
                reportImgTag.strPros = nvl(rsTemp!��������)
                reportImgTag.lngStartVer = Val(rsTemp!��ʼ��)
                reportImgTag.strKey = Val(rsTemp!ID)
                reportImgTag.strImgMarks = ""
            
                If Val(aryImgPro(0)) = 1 Then '���ͼ
                    reportImgTag.lngImgType = ritMark
                    
                    strPicFile = mstrReportImgPath & GetSignMarkImgName(mlngReportID, 0)
                    If ReadMarkImage(dcmMarkImages, rffReport, reportImgTag) = False Then
                        Exit Function
                    End If
                    
                    If dcmMarkImages.Count > 0 Then
                        dcmMarkImages(1).FileExport strPicFile, "BMP"
                    End If
                    
                    aryMarkPara(0) = strPicFile
                End If
                
                If Val(aryImgPro(0)) = 2 Then '����ͼ
                    reportImgTag.lngImgType = ritReport
                    
                    '��ȡ����ͼ�ļ�����
                    strPicFile = GetReportImagePro(reportImgTag.strPros, "PicName")
                    If Len(strPicFile) > 0 Then
                        strPicFile = FormatFilePath(mstrReportImgPath & "\" & strPicFile)
                    Else
                        '���ͼƬ�洢�����ݿ��У���û��picname����
                        strPicFile = FormatFilePath(mstrReportImgPath & "\����ͼ_" & reportImgTag.strKey & ".JPG")
                    End If
                     
                    If ReadReportImage(dcmRepImages, reportImgTag) <> frNormal Then Exit Function
                    
                    aryRepPara(intPCount) = strPicFile
                    
                    intPCount = intPCount + 1
                    If intPCount > UBound(aryRepPara) Then Exit Do
                End If
                
                Call rsTemp.MoveNext
            Loop
        End If
    End If


    '����ѡ����Զ��屨���ʽ����֯ͼ��
    '���ֻѡ����һ�ָ�ʽ�������Ƿ�ֻ��һ��ͼ���,ֻ��һ��ͼ����ʱ���Զ����ͼ��
    '���ѡ����2�����ϵĸ�ʽ�����ֻ��һ��ͼ������������Զ����
    arr�����ʽ = Split(strSelFmts, ",")

    '����û��ѡ���ʽ�����
    If Trim(strSelFmts) = "" Then
        ReDim arr�����ʽ(0)
        arr�����ʽ(0) = "1-1"
    End If
 

    '��ȡͼ�񣬵��ñ���
    lngReportImgCount = intPCount
    
    blnIsImageReport = False
    intPCount = 0       '��¼ͼ�������
    
    If lngReportImgCount > 0 Then strFirstFile = aryRepPara(0)
    
    For i = 0 To UBound(arr�����ʽ)
        If arr�����ʽ(i) <> "" Then
            int��ʽ�� = Split(arr�����ʽ(i), "-")(0)
    
            strSQL = "Select b.����,b.W,b.H From zlReports a, zlRptItems b" & vbNewLine & _
            "       Where a.Id = b.����id And a.��� = [1] And Nvl(b.����, 0) = 1 And b.���� = 11 And b.��ʽ�� = [2]" & vbNewLine & _
            "       Order By b.����" 'Trunc(b.y/567),Trunc(b.x/567)
            
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡͼ���", strReportNo, int��ʽ��)
            
            lngReportBoxCount = rsTemp.RecordCount
            
            rsTemp.Filter = "���� like '%���%'"
            lngReportBoxCount = lngReportBoxCount - rsTemp.RecordCount
            
            rsTemp.Filter = ""
            
            '����ͼ��ֻ��һ����������ͼ�ж��ʱ����Ҫ���ͼ��
            If lngReportImgCount > 1 Then
                aryRepPara(0) = strFirstFile
                
                If lngReportBoxCount = 1 Then
                    '���ͼ��
                    ResizeRegion lngReportImgCount, rsTemp("W"), rsTemp("H"), intRows, intCols
                    Set dcmResultImage = AssembleImage(dcmRepImages, intRows, intCols, rsTemp("H"), rsTemp("W"))
                    
                    aryRepPara(0) = Replace(Right(aryRepPara(0), Len(aryRepPara(0)) - InStr(aryRepPara(0), "=")), ".JPG", "") & "_GRP.JPG"
                    
                    dcmResultImage.FileExport aryRepPara(0), "JPEG"
                End If
            End If
    
            'װ��ͼ������
            intParaCount = 0
            Do While Not rsTemp.EOF
                blnIsImageReport = True
    
                '�ֱ�װ�ڱ��ͼ�ͱ���ͼ
                If InStr(rsTemp!����, "���") <> 0 Then '���ͼ
                    If aryMarkPara(0) <> "" Then strFlagString = rsTemp!���� & "=" & aryMarkPara(0)
                Else    '����ͼ
                    If intPCount > UBound(aryRepPara) Then Exit Do      '�������ı����е�ͼ����������ʵ�ʱ���ͼ���������˳�
                    If aryRepPara(intPCount) <> "" Then          '�����е�ͼ���ȱ����еĶ࣬�˳�
                        aryPrintPara(intParaCount) = rsTemp!���� & "=" & aryRepPara(intPCount)
                        intParaCount = intParaCount + 1
                    End If
                    
                    If lngReportBoxCount <> 1 Then intPCount = intPCount + 1
                    
                End If
                rsTemp.MoveNext
            Loop
    
            '��������ͼ�αȱ������ٵ����
            For j = intParaCount To UBound(aryPrintPara)
                If aryPrintPara(j) Like "*=*" Then aryPrintPara(j) = ""
            Next j
    
            '����Ǳ���Ԥ������ͼʱ���򲻽�����ʾ
            If blnIsImageReport And blnPrint Then
                If Trim(aryPrintPara(0)) = "" _
                    And Trim(aryPrintPara(1)) = "" _
                    And Trim(aryPrintPara(2)) = "" _
                    And Trim(aryPrintPara(3)) = "" _
                    And Trim(aryPrintPara(4)) = "" _
                    And Trim(aryPrintPara(5)) = "" _
                    And Trim(aryPrintPara(6)) = "" _
                    And Trim(aryPrintPara(7)) = "" _
                    And Trim(aryPrintPara(8)) = "" _
                    And Trim(aryPrintPara(9)) = "" Then
                    
                    If MsgboxH(GetRootHwnd, "δ���ִ���ӡ�ı���ͼ���Ƿ������ӡ��", vbYesNo, "��ʾ") = vbNo Then
                        Exit Function
                    End If
                End If
            End If
    
            '���ñ���
            Call objCusReport.ReportOpen(gcnOracle, glngSys, strReportNo, Nothing, _
                "NO=" & strExseNo, "����=" & intExseKind, "ҽ��ID=" & mlngAdviceId, strFlagString, _
                aryPrintPara(0), aryPrintPara(1), aryPrintPara(2), aryPrintPara(3), aryPrintPara(4), aryPrintPara(5), _
                aryPrintPara(6), aryPrintPara(7), aryPrintPara(8), aryPrintPara(9), aryPrintPara(10), aryPrintPara(11), _
                aryPrintPara(12), aryPrintPara(13), aryPrintPara(14), aryPrintPara(15), aryPrintPara(16), aryPrintPara(17), _
                aryPrintPara(18), aryPrintPara(19), "ReportFormat=" & int��ʽ��, IIf(blnPrint, 2, 1))
                
            CustomReportPrint = True
        End If
    Next i

Exit Function
errhandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ReportReject() As Boolean
'���沵��
Dim objfrmRj As frmReject
Dim i As Long
Dim lngAdviceColIndex As Long
Dim lngProcedureColIndex As Long
Dim lngRowIndex As Long
    
On Error GoTo errFree
    If mlngReportID <= 0 Then
        MsgboxH GetRootHwnd, "��ǰ���û�б��棬���ܽ��в��ز�����", vbOKOnly, "��ʾ"
        Exit Function
    End If
    
    Set objfrmRj = New frmReject
    
    ReportReject = objfrmRj.ShowRejectWindow(mlngAdviceId, mlngReportID, mObjNotify.Owner)
    
errFree:
    Unload objfrmRj
    Set objfrmRj = Nothing
End Function


Public Sub RejectHistory()
'��ʾ������ʷ
Dim frmRj As frmReject
    
On Error GoTo errFree
    If mlngReportID <= 0 Then
        MsgboxH GetRootHwnd, "��ǰ���û�б��棬�����ڲ�����ʷ��¼��", vbInformation, "��ʾ"
        Exit Sub
    End If
    
    Set frmRj = New frmReject
    
    Call frmRj.ShowRejectHistory(mlngAdviceId, mlngReportID, mObjNotify.Owner)
errFree:
    Unload frmRj
    Set frmRj = Nothing
End Sub


Public Sub RevisionHistory()
    Dim objHistory As New frmReportHistory
    
    Call objHistory.ZlShowMe(mObjNotify.Owner, mlngAdviceId, mlngReportID, mblnIsMoved)
End Sub

Public Sub ClearMark(Optional ByVal blnIsTriggerModify As Boolean = False)
'������
    Dim reportImgTag As TReportImgTag
    
    If dcmMarkImage.Images.Count > 0 Then
        dcmMarkImage.Images(1).Labels.Clear
        dcmMarkImage.Refresh
        
        reportImgTag = dcmMarkImage.Images(1).tag
        reportImgTag.strImgMarks = ""
        
        dcmMarkImage.Images(1).tag = reportImgTag
        
        If blnIsTriggerModify Then Call EnterModify(, , True)
    End If
End Sub

Public Sub ClearReportImg()
'�������ͼ
    dcmReportImg.Images.Clear
End Sub




Public Sub ClearInfo()
'���ǩ����Ϣ
    mlngCreateDeptId = mlngDeptID ' 0
    mstrCreateUser = UserInfo.���� ' ""
    mstrSaveUser = UserInfo.���� ' ""
    
    mblnIsLockingEdit = False
    mlngReportID = 0
    
    mlngSignLevel = cprSL_�հ�
    mstrFirstSignUser = ""
    mstrFinalSignUser = ""
    mintTargetVer = 1
    mintSourceVer = 0
    
    labEditState.Caption = ""
    labSign.Caption = ""
    
'    chkPositive.value = 0
'    chkCritical.value = 0
    
    mblnIsModifyImage = False
    mblnIsModifyMarks = False
    mblnIsModifyText = False
End Sub

Public Function LockEditor(Optional ByRef strErrMsg As String) As Boolean
'�����༭��
    'ʹ��ȫ����ʱ����в�������
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
On Error GoTo errhandle
    
    LockEditor = False
    
    If mblnIsLockingEdit Then
        LockEditor = True
        Exit Function
    End If
    
    strSQL = "select ������� from Ӱ�����¼ where ҽ��ID=[1]"
    
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���������", mlngAdviceId)
    If rsData.RecordCount > 0 Then
        If nvl(rsData!�������) <> "" Then
            '�ѱ��������ж��������Ƿ���ͬ
            If nvl(rsData!�������) = UserInfo.���� Then
                LockEditor = True
                mblnIsLockingEdit = True
            Else
                strErrMsg = "�����ѱ� [" & nvl(rsData!�������) & "] �༭����."
            End If
            
            Exit Function
        End If
    End If
        
        
    'û���������������������
    Call UpdateReporter(mlngAdviceId, UserInfo.����)
     
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���������", mlngAdviceId)
    If rsData.RecordCount <= 0 Then
        '����ʧ��
        strErrMsg = "��������ʧ��."
        Exit Function
    End If
    
    If nvl(rsData!�������) = UserInfo.���� Then
        LockEditor = True
        mblnIsLockingEdit = True
    Else
        strErrMsg = "��������ʧ��,�ѱ� [" & nvl(rsData!�������) & "] �༭����."
    End If
    
Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub UnlockEditor()
'����༭��
    Call UpdateReporter(mlngAdviceId, "")
End Sub

Public Function IsLockEditor(Optional ByRef strEditor As String = "")
'�����Ƿ������༭
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
On Error GoTo errhandle
    
    IsLockEditor = False
    
    strSQL = "Select ������� From Ӱ�����¼ Where ҽ��ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��¼", mlngAdviceId)
     
    If rsTemp.RecordCount > 0 Then
        If nvl(rsTemp!�������) <> "" And nvl(rsTemp!�������) <> UserInfo.���� Then
            strEditor = nvl(rsTemp!�������)
            IsLockEditor = True
        End If
    End If
    
    Exit Function
    
errhandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ResetEditState(Optional ByVal blnForceRead As Boolean = False)
'�Ƿ�ֻ��
    'ֻ��״ֻ̬�ܽ��в鿴��Ԥ������ӡ�Ȳ���
    
    '�Ѿ�ת���ı���Ϊֻ��
    '����ɵļ�飬��û�в�¼Ȩ��Ϊֻ��
    '�ѳ�Ժ�ҹ鵵�ı���Ϊֻ��
    
    
'�Ƿ�༭
    '�Ǳ༭״̬�ɽ��л��ˣ����أ���˵Ȳ���
    
    '���û���޸����˱����Ȩ�ޣ��ұ��洴����Ϊ�������ܱ༭
    '��������Ѿ���ˣ����ܱ༭�����Ǵ����˺��������ͬ
    '���û�б���༭Ȩ�ޣ����ܱ༭
    '��������Ѿ��������û������༭ʱ�����ܼ������б༭
    
'ֻ��״̬�µı��棬�϶����ܽ��б༭
    
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim lngAdviceState As Long
    Dim lngCurSignLevel As Long
    Dim strEditor As String
    
    If blnForceRead Then
    'ǿ�ƶ�״̬
        mblnIsEditable = False
        mblnIsReadOnly = True
        Exit Sub
    End If
    
    mblnIsEditable = Not mblnIsMoved    '�Ѿ�ת���ı��治�ܱ༭
    mblnIsReadOnly = mblnIsMoved
    mblnIsComplete = mblnIsMoved
   
    If mblnIsReadOnly Then Exit Sub '����Ѿ���ֻ��������Ҫ�����ж�
    
    '*****************************
    '��ѯҽ��ִ��״̬,���Ƿ��Ժ�鵵
    strSQL = "Select b.ִ�п���ID, a.ִ�й���,c.��Ժ����,c.����״̬,c.���ʱ�� " & _
        " From ����ҽ������ a,����ҽ����¼ b,������ҳ c  " & _
        " Where a.ҽ��ID = b.Id And  b.����ID = c.����ID(+) And b.��ҳID = c.��ҳID(+) And a.ҽ��ID= [1] "
    If mblnIsMoved Then
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
        
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯҽ������״̬", mlngAdviceId)
    If rsTemp.RecordCount > 0 Then
        lngAdviceState = nvl(rsTemp!ִ�й���, 0)
        
        mblnIsComplete = IIf(lngAdviceState = 6, True, False)
        
        If mlngReportID = 0 And lngAdviceState = 6 Then '��¼����ֻ�ܲ�¼ҽ����Ӧִ�п����µı��棬���ܿ���Ҳ�¼����
            '���û�б����Ҽ�������ʱ��ֻ�о߱�����¼���桱Ȩ��ʱ�����ܱ༭
            If CheckPopedom(mstrPrivs, "��¼����") And Val(nvl(rsTemp!ִ�п���ID)) = mlngDeptID Then
                labEditState.Caption = "��������"
                labEditState.ForeColor = vbBlue
                
                Exit Sub
            End If
        End If
    
        
        '����ɵı��棬Ϊֻ��״̬
        mblnIsReadOnly = IIf(lngAdviceState = 6 Or lngAdviceState = 0, True, False)
        If mblnIsReadOnly Then
            labEditState.Caption = IIf(lngAdviceState = 0, "δ����...", "���������")
            labEditState.ForeColor = vbBlue
            
            mblnIsEditable = False
            Exit Sub
        End If
        
        '��Ժ�ҹ鵵�󣬱��治�ɲ���,����״̬Ϊ5��ʾ���鵵
        mblnIsReadOnly = IIf(nvl(rsTemp!��Ժ����) <> "" And (nvl(rsTemp!����״̬, 0) = 5 Or nvl(rsTemp!���ʱ��, "") <> ""), True, False)
        If mblnIsReadOnly Then
            mblnIsComplete = True '�ѹ鵵�����ʾ�����
            labEditState.Caption = "�����ѹ鵵"
            labEditState.ForeColor = vbBlue
            
            mblnIsEditable = False
            Exit Sub
        End If
    End If
    
    '*****************************
    '�ͼ����ҽ�������޶��߼���ҽ���ı��棬�򿪱���󣬱���Ϊֻ���ġ�
    '�������ֻ���ڱ����Ѿ�ǩ������ȥ���ǣ�����ǩ������<>0���޸ĺ�δǩ���ģ��ں�����chkEditState�д���
    If mintSourceVer > 0 Then
        '�Լ���д�ı��棬Ӧ���ǿ��Ի��˵�
        '��ȡ��ǰ�û���ǩ������
        lngCurSignLevel = GetUserSignLevel(UserInfo.ID)
        If lngCurSignLevel < mlngSignLevel Then
            If mstrFirstSignUser = mstrFinalSignUser And mstrSaveUser = mstrFinalSignUser And mstrFinalSignUser = UserInfo.���� Then
                '�Լ�������ǩ���ı��棬�п��ܺ��汻�������û�ǩ������
                
            Else
                labEditState.Caption = "��������Ȩ�༭"
                labEditState.ForeColor = vbRed
                
                mblnIsReadOnly = True
                mblnIsEditable = False
                
                Exit Sub
            End If
        End If
    End If
    
    '*****************************
    '����������ж�
    If IsLockEditor(strEditor) Then
        labEditState.Caption = "��������[" & strEditor & "]�༭"
        labEditState.ForeColor = vbRed
        
        mblnIsReadOnly = True
        mblnIsEditable = False
        Exit Sub
    End If
    
    
    '*****************************
    '�жϴ����˺͵�ǰ�û��Ƿ���ͬ�������ͬ��������༭
    If mintSourceVer = 0 And CheckPopedom(mstrPrivs, "PACS������д") Then
        '�б�����дȨ��
        If mstrCreateUser = UserInfo.���� Then
            mblnIsEditable = True
        ElseIf CheckPopedom(mstrPrivs, "PACS���˱���") And (mlngCreateDeptId = mlngDeptID Or IsContainDept(UserInfo.ID, mlngCreateDeptId)) Then   '�����˱���Ȩ�޵ģ�������д�����ҵı���
            mblnIsEditable = True
        Else
            labEditState.Caption = "��Ȩ�༭[" & mstrCreateUser & "]�ı���"
            labEditState.ForeColor = vbRed
            
            mblnIsReadOnly = True '�����˱���Ȩ��ʱ��������û��ǩ������£�����������κβ���
            mblnIsEditable = False
            Exit Sub
        End If
    ElseIf mintSourceVer > 0 Then   '�Ѿ�ǩ���ı��棬������ֱ��ɾ���������Ƚ��л��˴���
        If (CheckPopedom(mstrPrivs, "PACS�����޶�")) And (mlngCreateDeptId = mlngDeptID Or IsContainDept(UserInfo.ID, mlngCreateDeptId)) Then  ' Or CheckPopedom(mstrPrivs, "PACS���˱���")
            '�б����޶�Ȩ��
            '�ڱ����޶���״̬�£��б����޶�Ȩ�޵��ˣ�������д�����ҵı��档
            'mstrCreateUser = UserInfo.���� And mstrSaveUser <> UserInfo.������ʾ�������Լ��������ѱ������޶������
            If (mstrSaveUser = UserInfo.����) Or (Not (mstrCreateUser = UserInfo.���� And mstrSaveUser <> UserInfo.����)) Then     '����������Լ���󱣴�ģ�����ǰ����޸����Ѿ�ǩ��
                mblnIsEditable = True
            Else
                '�Ѿ��������޶��������,�޸��Ѿ����棬����û��ǩ�������治�ɱ༭����¼�޶�������
                labEditState.Caption = "�ѱ�[" & mstrSaveUser & "]�޶�"
                mblnIsEditable = False
                Exit Sub
            End If
        ElseIf mstrFirstSignUser = UserInfo.���� And mstrFinalSignUser = UserInfo.���� Then '���û���޶������˱���Ȩ�ޣ����ж��Ƿ����ǩ���͵�ǰ�û���ͬ
            '���߱����˱���Ȩ�ޣ��ڶ����˴����ı������ǩ�����״�ǩ���˺�����ǩ��������ͬ��
            mblnIsEditable = True
        Else
            'ֻ�о߱���д���޶�����˻����˱����Ҽ�����ڵ��ڵ�ǰǩ������Ȩ�޲��ܽ��л���,���ֻ�߱�������дȨ�ޣ���û������������������л��˵�
            If Not (CheckPopedom(mstrPrivs, "PACS�������") Or CheckPopedom(mstrPrivs, "PACS��������")) Then
                '���޶�������ˣ�������Ȩ��ʱΪֻ�������ܽ��л��˺�ɾ������
                mblnIsReadOnly = True
            End If
            
            If mstrCreateUser = UserInfo.���� And mstrFinalSignUser <> UserInfo.���� Then
                labEditState.Caption = "�����ѱ�[" & mstrFinalSignUser & "]�޶�"
            Else
                labEditState.Caption = "��Ȩ�޶�[" & mstrCreateUser & "]�ı���"
            End If
            labEditState.ForeColor = vbRed
            
            mblnIsEditable = False
            Exit Sub
        End If
    Else
        If mintSourceVer <= 0 Then  'û�н����κ�ǩ��,δǩ�����治����������
            mblnIsReadOnly = True
        End If
        
        If mstrCreateUser <> UserInfo.���� Then
            labEditState.Caption = "��Ȩ�༭[" & mstrCreateUser & "]�ı���"
        Else
            labEditState.Caption = "�ޱ�����дȨ��"
        End If
        
        labEditState.ForeColor = vbRed
        
        mblnIsEditable = False
        Exit Sub
    End If
     
    '*****************************
    'ֻ����д��鼼ʦΪ�Լ��ı���
    If mblnTechReptSame Then
        strSQL = " select ��鼼ʦ from Ӱ�����¼ where ҽ��id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ��鼼ʦ", mlngAdviceId)
        
        If rsTemp.RecordCount < 1 Then
            labEditState.Caption = "��Ч������ݽ�ֹ�༭"
            labEditState.ForeColor = vbRed
            
            mblnIsEditable = False
            Exit Sub
        End If
        
        If nvl(rsTemp!��鼼ʦ) <> UserInfo.���� Then
            
            labEditState.Caption = "ֻ����д�Լ����ı��棬��ǰ���û��ȷ����鼼ʦ��"
            If nvl(rsTemp!��鼼ʦ, "") <> "" Then
                labEditState.Caption = labEditState.Caption & "����Ȩ��[" & nvl(rsTemp!��鼼ʦ) & "]�ļ����д"
            End If
            labEditState.ForeColor = vbRed
            
            mblnIsEditable = False
            Exit Sub
        Else
            mblnIsEditable = True
        End If
    
    End If

    '��ͼ�������д����
    If mblnIsEditWithReportImage Then
        strSQL = " select ���UID from Ӱ�����¼ where ҽ��id = [1]"
        If mblnIsMoved Then strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
        
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���UID", mlngAdviceId)
        
        If rsTemp.RecordCount < 1 Then
            labEditState.Caption = "��Ч������ݽ�ֹ�༭"
            labEditState.ForeColor = vbRed

            mblnIsEditable = False
            
            Exit Sub
        End If
        
        If Len(nvl(rsTemp!���UID)) <= 0 Then
            labEditState.Caption = "�޼��ͼ���ֹ��д"
            labEditState.ForeColor = vbRed

            mblnIsEditable = False
            Exit Sub
        End If
    End If
    
    '��ʾ���浱ǰ״̬
    If mlngReportID = 0 Then
        labEditState.Caption = "��ʼ�༭..."
        labEditState.ForeColor = vbBlue
    Else
        If mintSourceVer >= 1 Then
            labEditState.Caption = "�����޶���..."
        Else
            labEditState.Caption = "������д��..."
        End If
        
        labEditState.ForeColor = vbBlue
    End If
    
End Sub

Public Sub ConfigFaceState(Optional ByVal blnForceRead As Boolean = False, Optional ByVal strEditState As String = "")
'���ý���״̬
    Call ResetEditState(blnForceRead)
    
    rtb����.Locked = Not mblnIsEditable Or mblnIsReadOnly
    rtb���.Locked = Not mblnIsEditable Or mblnIsReadOnly
    rtb����.Locked = Not mblnIsEditable Or mblnIsReadOnly
    
'    chkPositive.Enabled = mblnIsEditable Or Not mblnIsReadOnly
'    chkCritical.Enabled = mblnIsEditable Or Not mblnIsReadOnly
'    picState.Enabled = mblnIsEditable Or Not mblnIsReadOnly
    picImageBack.Enabled = mblnIsEditable Or Not mblnIsReadOnly
    
    If Not mblnIsEditable Or mblnIsReadOnly Then
        rtb����.BackColor = UserControl.BackColor
        rtb���.BackColor = UserControl.BackColor
        rtb����.BackColor = UserControl.BackColor
    Else
        rtb����.BackColor = ColorConstants.vbWhite
        rtb���.BackColor = ColorConstants.vbWhite
        rtb����.BackColor = ColorConstants.vbWhite
    End If
    
    If Len(strEditState) > 0 Then labEditState.Caption = strEditState
End Sub


Public Sub AddRepImage(objDcmImg As Object, _
    Optional ByVal lngReleationImageAdviceId As Long = 0, Optional ByVal strFileName As String = "", _
    Optional ByVal blnForceAdd As Boolean = False)
'��ӱ���ͼ
'lngReleationImageAdviceId:����ͼ���ҽ��id���п��ܸü��鿴��ͼ���Ǵӹ�������������д򿪵�

    Dim objCurDicom As DicomImage
    Dim reportTag As TReportImgTag
    
    
    '�Ǳ༭״̬�£�������Ա���ͼ���в���
    If mblnIsEditable = False And blnForceAdd = False Then Exit Sub
    
    Set objCurDicom = objDcmImg
    
    reportTag.strKey = objCurDicom.InstanceUID
    
    If strFileName = "" Then
        If lngReleationImageAdviceId = 0 Then
            reportTag.strImgFile = objCurDicom.InstanceUID & ".JPG"
        Else
            reportTag.strImgFile = objCurDicom.InstanceUID & "_" & lngReleationImageAdviceId & ".JPG"
        End If
    Else
        reportTag.strImgFile = strFileName
    End If
    
    reportTag.lngFromAdvice = lngReleationImageAdviceId
    
    objCurDicom.tag = reportTag
 
    Call DrawBorder(objCurDicom, 0)
    
    Call dcmReportImg.Images.Add(objCurDicom)
    Call CalcImgView
    
    If blnForceAdd = False Then Call EnterModify(, True)
End Sub

Public Sub AddRepImgFile(ByVal strFile As String, _
    Optional ByVal lngImageAdviceId As Long = 0, Optional ByVal strFileName As String = "", _
    Optional ByVal blnForceAdd As Boolean = False)
'��ӱ���ͼ�ļ�
    Dim objCurDicom As DicomImage
    Dim strError As String
    
    Set objCurDicom = ReadDicomFile(strFile, strError, False)
    
    Call AddRepImage(objCurDicom, lngImageAdviceId, strFileName, blnForceAdd)
End Sub

Private Sub DelRepImage()
'ɾ����ǰ����ͼ
    Dim strImgKey As String
    Dim strSQL As String
    
    If mlngSelReportImgIndex <= 0 Or mlngSelReportImgIndex > dcmReportImg.Images.Count Then Exit Sub
    
    strImgKey = dcmReportImg.Images(mlngSelReportImgIndex).InstanceUID
    
    '�������ݿ�
    strSQL = "Zl_Ӱ����_���ñ���ͼ('" & strImgKey & "',2)"
    Call zlDatabase.ExecuteProcedure(strSQL, "ɾ������ͼ")
    
    Call dcmReportImg.Images.Remove(mlngSelReportImgIndex)
    
    Call CalcImgView
    
    mlngSelReportImgIndex = 0
     
    Call EnterModify(, True)
    
    RaiseEvent OnDelRepImg(strImgKey)
End Sub


Public Sub Mark(ByVal MarkType As TImgMarkType, Optional ByVal strMark As String = "")
'ָ���ı����
    mlngMarkType = MarkType
    mstrMarkText = strMark
End Sub


Public Sub InputWord(ByVal strFreeText As String, _
    ByVal str���� As String, ByVal str��� As String, ByVal str���� As String)
'д��ʾ�
    Dim blnIsUseSpecialty As Boolean
    Dim lngStartSel As Long
    
    blnIsUseSpecialty = False
    If Not mobjSpePlugin Is Nothing Then
        blnIsUseSpecialty = mblnIsSpeState
    End If
    
    If blnIsUseSpecialty = False Then
        If Len(strFreeText) > 0 Then
            If Not mrtbActive Is Nothing Then
                If mrtbActive.Enabled And mrtbActive.Locked = False And mrtbActive.Visible Then
                    lngStartSel = mrtbActive.SelStart
                    
                    mrtbActive.SelText = strFreeText
                    Call SetWordStyle(mrtbActive)
                    
                    mrtbActive.SelStart = lngStartSel + Len(strFreeText)
                    mrtbActive.SetFocus
                End If
            End If
        End If
        
        If Len(str����) > 0 Then
            If rtb����.Enabled And rtb����.Locked = False And rtb����.Visible Then
                lngStartSel = rtb����.SelStart
                rtb����.SelText = str����
                
                Call SetWordStyle(rtb����)
                
                rtb����.SelStart = lngStartSel + Len(str����)
                
                rtb����.SetFocus
            Else
                If rtb����.Visible = False Then MsgboxH GetRootHwnd, "�ôʾ����ݽ������� [����] ������١�", vbOKOnly, "��ʾ"
            End If
        End If
        
        If Len(str���) > 0 Then
            If rtb���.Enabled And rtb���.Locked = False And rtb���.Visible Then
                lngStartSel = rtb���.SelStart
                rtb���.SelText = str���
                
                Call SetWordStyle(rtb���)
                
                rtb����.SelStart = lngStartSel + Len(str���)
                rtb����.SetFocus
            Else
                If rtb���.Visible = False Then MsgboxH GetRootHwnd, "�ôʾ����ݽ������� [���] ������١�", vbOKOnly, "��ʾ"
            End If
        End If
        
        If Len(str����) > 0 Then
            If rtb����.Enabled And rtb����.Locked = False And rtb����.Visible Then
                lngStartSel = rtb����.SelStart
                rtb����.SelText = str����
                
                Call SetWordStyle(rtb����)
                
                rtb����.SelStart = lngStartSel + Len(str����)
                rtb����.SetFocus
            Else
                If rtb����.Visible = False Then MsgboxH GetRootHwnd, "�ôʾ����ݽ������� [����] ������١�", vbOKOnly, "��ʾ"
            End If
        End If
    Else
        Call InputWordToSpecialty(strFreeText, str����, str���, str����)
    End If
End Sub

Private Sub InputWordToSpecialty(ByVal strFreeText As String, _
    ByVal str���� As String, ByVal str��� As String, ByVal str���� As String)
On Error GoTo errhandle
    If mobjSpePlugin Is Nothing Then Exit Sub
    
    Call mobjSpePlugin.InputWord(strFreeText, str����, str���, str����)
Exit Sub
errhandle:
    
End Sub


Public Sub GetReportContext(ByRef str���� As String, ByRef str��� As String, ByRef str���� As String, _
    Optional ByRef strSelText As String = "")
'��ȡ��������
    str���� = rtb����.Text
    str��� = rtb���.Text
    str���� = rtb����.Text
    
    If Not mrtbActive Is Nothing Then
        strSelText = mrtbActive.SelText
    End If
End Sub



Private Function GetDkpStateString(ByVal strSourceStateString As String, ByVal strCurStateString As String) As String
    Dim strSourceFmt As String
    Dim arySourcePaneInfo() As String
 
    Dim strCurFmt As String
    Dim aryCurPaneInfo() As String

    Dim i As Long
    Dim strNewPaneFmt As String
    Dim strTitle As String
    Dim strSourcePaneFmt As String
    Dim lngPaneInfoCount As Long
    
    GetDkpStateString = ""

    strSourceFmt = strSourceStateString
    strSourceFmt = Mid(strSourceFmt, InStr(strSourceFmt, "<Pane-1"), 4096)
    strSourceFmt = Mid(strSourceFmt, 1, InStr(strSourceFmt, "</Common>") - 1)
 
    strCurFmt = strCurStateString
    strCurFmt = Mid(strCurFmt, InStr(strCurFmt, "<Pane-1"), 4096)
    strCurFmt = Mid(strCurFmt, 1, InStr(strCurFmt, "</Common>") - 1)

    arySourcePaneInfo = Split(strSourceFmt, "<Pane-")
    aryCurPaneInfo = Split(strCurFmt, "<Pane-")

    lngPaneInfoCount = UBound(arySourcePaneInfo)

    For i = 1 To lngPaneInfoCount
        strSourcePaneFmt = arySourcePaneInfo(i)
        strSourcePaneFmt = Mid(strSourcePaneFmt, InStr(strSourcePaneFmt, "Type="), 255)
        
        strTitle = GetDkpTitleValue(strSourcePaneFmt)
        
        If InStr(strSourcePaneFmt, "Type=""2""") > 0 Then            '
            strNewPaneFmt = strNewPaneFmt & "<Pane-" & i & " " & strSourcePaneFmt
            
        ElseIf InStr(strSourcePaneFmt, "Type=""1""") > 0 Then
            strNewPaneFmt = strNewPaneFmt & "<Pane-" & i & " " & GetDkpReleationFmt(i, i + 1, strSourceFmt, strCurFmt, arySourcePaneInfo, aryCurPaneInfo)
            
        ElseIf InStr(strSourcePaneFmt, "Type=""0""") > 0 Then
            strNewPaneFmt = strNewPaneFmt & "<Pane-" & i & " " & GetDkpNewFmt(strTitle, strCurFmt, i - 1)
            
        Else
            strNewPaneFmt = strNewPaneFmt & "<Pane-" & i & " " & strSourcePaneFmt
            
        End If
    Next
    
    strNewPaneFmt = "<Layout><Common CompactMode=""1"">" & GetDkpSummaryInfo(strSourceStateString) & strNewPaneFmt & "</Common></Layout>"
    
    GetDkpStateString = strNewPaneFmt
End Function

Private Function GetDkpSummaryInfo(ByVal strCurFmt As String) As String
    Dim strTmp As String
    
    strTmp = Mid(strCurFmt, InStr(strCurFmt, "<Summary"), 4096)
    strTmp = Mid(strTmp, 1, InStr(strTmp, "/>") - 1) & "/>"
    
    GetDkpSummaryInfo = strTmp
End Function

Private Function GetDkpReleationFmt(ByVal lngPaneIndex As Long, ByVal lngBindIndex As Long, _
    ByVal strSourceFmt As String, ByVal strCurFmt As String, _
    arySource() As String, aryCur() As String) As String
'����ָ����Դpane��������ȡ�µĸ�ʽ����
    
    Dim strReleationTitle As String
    Dim lngReleationIndex As Long
    Dim strFmt As String
    Dim strTmp As String
    Dim strSourcePaneFmt As String
    
    
    GetDkpReleationFmt = ""
    strSourcePaneFmt = arySource(lngPaneIndex)
    
    strFmt = strSourcePaneFmt
    strFmt = Mid(strFmt, InStr(strFmt, "Pane-1=""") + 8, 4096)
    strTmp = Mid(strFmt, 1, InStr(strFmt, """/>") - 1)
    
    lngReleationIndex = Val(strTmp)
    
    
    strSourcePaneFmt = arySource(lngReleationIndex)
    strFmt = Mid(strSourcePaneFmt, InStr(strSourcePaneFmt, "Title=""") + 7, 4096)
    strReleationTitle = Mid(strFmt, 1, InStr(strFmt, """") - 1)
     
    
    
    strFmt = Mid(strCurFmt, InStr(strCurFmt, strReleationTitle), 4096)
    strFmt = Mid(strFmt, 1, InStr(strFmt, "/>") - 1)
    strFmt = Mid(strFmt, InStr(strFmt, "LastHolder=""") + 12, 4096)
    lngReleationIndex = Val(Mid(strFmt, 1, InStr(strFmt, """") - 1))
    
    
    
    strFmt = Mid(aryCur(lngReleationIndex), 3, 4096)
    
    If InStr(strFmt, "Selected=") > 0 Then
        strFmt = Mid(strFmt, 1, InStr(strFmt, "Selected=") - 1)
        
        GetDkpReleationFmt = strFmt & "Selected=""" & lngBindIndex & """ Pane-1=""" & lngBindIndex & """/>"
    Else
        GetDkpReleationFmt = strFmt
    End If
End Function

Private Function GetDkpTitleValue(ByVal strPaneFmt As String) As String
'�Ӹ�ʽ�л�ȡpane����
    Dim strTmp As String

    GetDkpTitleValue = ""
    If InStr(strPaneFmt, "Title=") <= 0 Then Exit Function
    
    strTmp = Mid(strPaneFmt, InStr(strPaneFmt, "Title=") + 6, 4096)
    
    GetDkpTitleValue = Mid(strTmp, 1, InStr(strTmp, " ID=") - 1)
End Function


Private Function GetDkpNewFmt(ByVal strTitle As String, ByVal strCurFmt As String, ByVal lngHolderIndex As Long) As String
'����ָ�������ȡ�µ�pane��ʽ����
    Dim strTmp As String
    
    strTmp = Mid(strCurFmt, InStrRev(strCurFmt, "Pane-", InStr(strCurFmt, strTitle)), 4096)
    strTmp = Mid(strTmp, 1, InStr(strTmp, "/>") - 1)
    strTmp = Mid(strTmp, InStr(strTmp, "Type="), 4096)
    
    strTmp = Mid(strTmp, 1, InStr(strTmp, "DockingHolder=") - 1)
    
    strTmp = strTmp & "DockingHolder=""" & lngHolderIndex & """ LastHolder=""" & lngHolderIndex & """/>"
    
    GetDkpNewFmt = strTmp
End Function


Public Function GetLayoutStr() As String
'���ظ�ʽ�ַ���[Key=TESTNAME@picturebox1.width:20;picturebox1.height:30;]
    If dkpMain.PanesCount >= 5 Then
        Call dkpMain.DestroyPane(dkpMain.Panes(5))
    End If
    
    GetLayoutStr = "[KEY=EDITOR@" & _
                                        GetProFmt("DKPMAINSTATESTR", GetDkpStateString(dkpMain.tag, dkpMain.SaveStateToString())) & _
                                        GetProFmt("REPORTIMG.WIDTH", dcmReportImg.Width) & _
                                        GetProFmt("MARKIMG.WIDTH", dcmMarkImage.Width) & _
                                 "]"
                                  
End Function

Public Function GetFaceKey() As String
    Dim strKeyTag As String

    strKeyTag = IIf(dkpMain.Panes(1).Closed, "0", "1")

    strKeyTag = strKeyTag & IIf(dkpMain.Panes(2).Closed, "0", "1")

    strKeyTag = strKeyTag & IIf(dkpMain.Panes(3).Closed, "0", "1")

    strKeyTag = strKeyTag & IIf(dkpMain.Panes(4).Closed, "0", "1")
    
    GetFaceKey = strKeyTag
End Function

Public Sub SetLayout(ByVal strLayout As String)
    Dim strPros As String
    Dim strPro As String
    Dim arySourcePane() As paneInfo
    Dim i As Long
    Dim objPane As Pane

    If Len(strLayout) <= 0 Then Exit Sub

    strPros = GetPros(strLayout, "EDITOR")

    strPro = GetProValue(strPros, "DKPMAINSTATESTR")
    If Len(strPro) > 0 Then
        ReDim arySourcePane(dkpMain.PanesCount)
        
        For i = 1 To dkpMain.PanesCount
            arySourcePane(i - 1).ID = dkpMain.Panes(i).ID
            arySourcePane(i - 1).hwnd = dkpMain.Panes(i).Handle
    
            arySourcePane(i - 1).hidden = dkpMain.Panes(i).hidden
            arySourcePane(i - 1).iconid = dkpMain.Panes(i).iconid
            arySourcePane(i - 1).options = dkpMain.Panes(i).options
            arySourcePane(i - 1).tag = dkpMain.Panes(i).tag
            arySourcePane(i - 1).title = dkpMain.Panes(i).title
        Next
  
        Call dkpMain.LoadStateFromString(strPro)
        
        For i = dkpMain.PanesCount To 1 Step -1
    
            dkpMain.Panes(i).ID = arySourcePane(i - 1).ID
            dkpMain.Panes(i).Handle = arySourcePane(i - 1).hwnd
            dkpMain.Panes(i).hidden = arySourcePane(i - 1).hidden
            dkpMain.Panes(i).iconid = arySourcePane(i - 1).iconid
            dkpMain.Panes(i).tag = arySourcePane(i - 1).tag
            dkpMain.Panes(i).title = arySourcePane(i - 1).title
            dkpMain.Panes(i).options = arySourcePane(i - 1).options
        Next
        
'        If Not mobjSpePlugin Is Nothing Then
'            '����ר�Ʊ���
'            If dkpMain.PanesCount < 5 Then
'                Set objPane = dkpMain.CreatePane(5, 0, 700, DockBottomOf, dkpMain.Panes(1))
'                objPane.title = "ר��¼��"
'                objPane.Handle = mobjSpePlugin.hwnd
'                objPane.tag = 4
'                objPane.options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
'                objPane.Closed = True
'            Else
'                If dkpMain.Panes(5).Handle = 0 Then
'                    dkpMain.Panes(5).title = "ר��¼��"
'                    dkpMain.Panes(5).Handle = mobjSpePlugin.hwnd
'                    dkpMain.Panes(5).tag = 4
'                    dkpMain.Panes(5).options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
'                    dkpMain.Panes(5).Closed = True
'                End If
'            End If
'        End If
      
    End If
    
    If (dcmReportImg.Visible Or Val(dcmReportImg.tag) <> 0) And (dcmMarkImage.Visible Or dcmMarkImage.Images.Count > 0) Then
        strPro = GetProValue(strPros, "REPORTIMG.WIDTH")
        If Val(strPro) > 0 Then dcmReportImg.Width = Val(strPro)
         
        strPro = GetProValue(strPros, "MARKIMG.WIDTH")
        If Val(strPro) > 0 Then dcmMarkImage.Width = Val(strPro)
    End If
End Sub

Public Sub Relayout()
    Dim i As Long
    
    '��ʼ���沼��
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane, pane4 As Pane, Pane5 As Pane
    
'    If dkpMain.PanesCount > 0 Then
'        For i = 1 To dkpMain.PanesCount
'            SetParent dkpMain.Panes(i).Handle, hwnd
'            dkpMain.Panes(i).Handle = 0
'        Next
'    End If

    If dkpMain.PanesCount <= 0 Then
    
        With dkpMain
            .CloseAll
            .DestroyAll
            .options.HideClient = True
            .options.UseSplitterTracker = False 'ʵʱ�϶�
            .options.ThemedFloatingFrames = True
            .options.AlphaDockingContext = True
        End With
        
        'ͼ��
        Set Pane1 = dkpMain.CreatePane(1, 0, 200, DockTopOf, Nothing)
        Pane1.title = "����ͼ��"
        Pane1.Handle = picImageBack.hwnd
        Pane1.tag = 0 '"REPIMG"
        Pane1.options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
        
        '����
        Set Pane2 = dkpMain.CreatePane(2, 0, 400, DockBottomOf, Nothing)   'Pane1
        Pane2.title = mstrDescTitle
        Pane2.Handle = picDesc.hwnd
        Pane2.tag = 1 ' "DESC"
        Pane2.options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        
        '���
        Set Pane3 = dkpMain.CreatePane(3, 0, 300, DockBottomOf, Pane2)
        Pane3.title = mstrOpinTitle
        Pane3.Handle = picOpin.hwnd
        Pane3.tag = 2 '"OPIN"
        Pane3.options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        
        '����
        Set pane4 = dkpMain.CreatePane(4, 0, 100, DockBottomOf, Pane3)
        pane4.title = mstrAdviTitle
        pane4.Handle = picAdvi.hwnd
        pane4.tag = 3 '"ADVI"
        pane4.options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        
        
        dkpMain.tag = dkpMain.SaveStateToString
    End If
    
    '����ר��
    If mblnVisibleSpecialty Then
        Call LoadSpecialtyPlugin
        
        If mobjSpePlugin Is Nothing Then
            mblnVisibleSpecialty = False
'        Else
'            'ר�Ʊ���¼��
'            Set Pane5 = dkpMain.CreatePane(5, 0, 700, DockBottomOf, Pane1)
'            Pane5.title = "ר��¼��"
'            Pane5.Handle = mobjSpePlugin.hwnd
'            Pane5.tag = 4
'            Pane5.options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
'            Pane5.Closed = True
        End If
    Else
        '���������ר�ƣ�������֮ǰ������ʾ��ר��¼�����
        If mblnIsSpeState Then
            Call ChangeSepState(False, True)
        End If
        
        If Not mobjSpePlugin Is Nothing Then SetParent mobjSpePlugin.hwnd, 0
        
        Set mobjSpePlugin = Nothing
    End If
'
'    Call dkpMain.RecalcLayout
'
    
    
'    chkPositive.Visible = Not mblnIgnoreResult
End Sub


Public Sub LocateEditBox()
    Dim objActive As Object

    
    If dkpMain.PanesCount <= 0 Then Exit Sub
    
    'ר�Ʊ���¼��ʱ��������Ա༭����ж�λ
    If mblnIsSpeState Then
        '...
        Exit Sub
    End If
    
    Set objActive = UserControl.ActiveControl
    
    If objActive Is Nothing Then
        Set objActive = mrtbActive
    Else
        If Not (TypeOf objActive Is RichTextBox) Then Set objActive = mrtbActive
    End If


    If Not objActive Is Nothing Then
        If TypeOf objActive Is RichTextBox Then
            If objActive.Visible And objActive.Locked = False Then objActive.SetFocus
            Exit Sub
        End If
    End If

    If rtb����.Visible Then
        If rtb����.Enabled Then
            rtb����.SetFocus
            rtb����.SelStart = Len(rtb����.Text)
        End If
        Exit Sub
    End If

    If rtb���.Visible Then
        If rtb���.Enabled Then
            rtb���.SetFocus
            rtb���.SelStart = Len(rtb���.Text)
        End If
        
        Exit Sub
    End If

    If rtb����.Visible Then
        If rtb����.Enabled Then
            rtb����.SetFocus
            rtb����.SelStart = Len(rtb����.Text)
        End If
        Exit Sub
    End If
End Sub

Public Sub GetReport(ByRef str���� As String, ByRef str��� As String, ByRef str���� As String)
'���֧�ַ�������ȡ��ǰ�༭��¼��ı�������
    str���� = rtb����.Text
    str��� = rtb���.Text
    str���� = rtb����.Text
End Sub

Public Sub ClearReport(ByVal blnClearDesc As Boolean, ByVal blnClearOpin As Boolean, ByVal blnClearAdvi As Boolean)
'���֧�ַ����������ǰ�����е��ı�����
    If blnClearDesc Then rtb����.Text = ""
    If blnClearOpin Then rtb���.Text = ""
    If blnClearAdvi Then rtb����.Text = ""
End Sub

Public Sub SendReport(ByVal str���� As String, ByVal str��� As String, ByVal str���� As String)
'���֧�ַ���������ר�Ʊ�����¼����ı�����
    rtb����.Text = str����
    rtb���.Text = str���
    rtb����.Text = str����
End Sub

Private Function LoadSpecialtyPlugin() As Boolean
'����ר�Ʊ�����
    Dim objParent As Object
    Dim strErr As String
    
On Error GoTo errhandle
    LoadSpecialtyPlugin = False
    If mblnVisibleSpecialty = False Then Exit Function
    
    Set mobjSpePlugin = DynamicCreate("ZLPacsProReport.clsZLPacsProReport", "ר��¼��")
     
     If mobjSpePlugin Is Nothing Then Exit Function
     
    Set objParent = UserControl.Extender
    Call mobjSpePlugin.InitPlugin(gcnOracle, objParent)
    
    LoadSpecialtyPlugin = True
Exit Function
errhandle:
    mblnVisibleSpecialty = False
    strErr = err.Description
    
    MsgboxH GetRootHwnd, "ר�Ʊ����ʼ��ʧ�ܣ�" & strErr, vbOKOnly, "��ʾ"
End Function

Private Sub ResizeEdit(rtbEdit As RichTextBox, picParent As PictureBox)
    rtbEdit.Left = 0
    rtbEdit.Top = 0
    rtbEdit.Width = picParent.Width
    rtbEdit.Height = picParent.Height
End Sub


Private Sub dcmMarkImage_DblClick()
'TASK:��ʱ��֧�ֱ��ͼ�������߼�����
''�򿪱��ͼ����
'    If dcmMarkImage.Images.Count <> 1 Then Exit Sub
'
'
'    Call ShowMarkImgProcess
End Sub

'Public Sub ShowMarkImgProcess()
'    Dim i As Long
'    Dim objDcmImg As DicomImage
'    Dim aryNull() As Object
'
'    If mobjMarkProcessV2 Is Nothing Then Set mobjMarkProcessV2 = New frmImageProcessV2
'
'    Set objDcmImg = dcmMarkImage.Images(1)
'
'    '��û���κδ����£�2����Զ��رմ�ͼԤ��
'    Call mobjMarkProcessV2.ZlShowMe(mObjNotify.Owner, mlngAdviceID, objDcmImg, aryNull, ptMark, 2, False)
'End Sub


Private Sub dcmMarkImage_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim lngFrame As Long
    
    If dcmMarkImage.Images.Count <= 0 Then Exit Sub
    If mlngMarkType = imtNormal Then
        '�����б�Ǵ���
        picImageBack.MousePointer = 0
        picImageBack.MouseIcon = Nothing
        
        Exit Sub
    End If
    
    lngFrame = 2
    
    '�������
    If dcmMarkImage.ImageXPosition(X, Y) > lngFrame And dcmMarkImage.ImageXPosition(X, Y) < dcmMarkImage.Images(1).SizeX - lngFrame _
       And dcmMarkImage.ImageYPosition(X, Y) > lngFrame And dcmMarkImage.ImageYPosition(X, Y) < dcmMarkImage.Images(1).SizeY - lngFrame Then
        picImageBack.MousePointer = 99
        picImageBack.MouseIcon = listCur.ListImages("pen").Picture
        
        SetCapture dcmMarkImage.hwnd
    Else
        ReleaseCapture
        
        picImageBack.MousePointer = 0
        picImageBack.MouseIcon = Nothing
    End If
End Sub

Public Sub AddNumber()
'���ı������ǰ�����������
'mintReportViewType 0-�������CheckView��1-������Result��2-����Advice

    Dim rText As RichTextBox
    Dim strtext As String
    Dim iCount As Integer
    Dim iStart As Integer
    
On Error GoTo err
 
    Set rText = mrtbActive
    
    If rText Is Nothing Then
        MsgboxH GetRootHwnd, "��ѡ����Ҫ��ŵı���༭��", vbOKOnly, "��Ϣ��ʾ"
        Exit Sub
    End If
    
    strtext = rText.Text
    
    '���ж��ı����Ƿ�����
    If rText.Locked = True Then
        MsgboxH GetRootHwnd, "�ı��α�����������༭��", vbOKOnly, "��ʾ"
        Exit Sub
    End If
    
    '���жϸ��ı����е�һ���ַ��Ƿ�����1������ǣ�����ʾ�Ѿ������ֱ�ţ��Ƿ�Ҫ���
    If Left(strtext, 1) = "1" Then
        If MsgboxH(GetRootHwnd, "�����ı����Ѿ��������ֱ�ţ��Ƿ�Ҫ������ֱ�ţ�", vbOKCancel, "��ʾ") = vbCancel Then
            Exit Sub
        End If
    End If
    
    '��ʼ������ֱ��,ÿһ���س�֮��������ǿո񣬾�������
    iStart = 1
    
    '��һ��Ҳ��Ҫ�ж��Ƿ��������
    If Left(strtext, 1) <> " " Then
        iCount = 1
        strtext = iCount & ". " & strtext
    Else
        iCount = 0
    End If
    iStart = InStr(iStart, strtext, vbCrLf)
    
    While (iStart <> 0)
        If Mid(strtext, iStart + 2, 1) <> " " And Mid(strtext, iStart + 2, 2) <> vbCrLf And Mid(strtext, iStart + 2, 1) <> "" Then
            iCount = iCount + 1
            strtext = Left(strtext, iStart + 1) & iCount & ". " & Right(strtext, Len(strtext) - iStart - 1)
        End If
        iStart = InStr(iStart + 1, strtext, vbCrLf)
    Wend
    
    rText.Text = strtext
    
    Exit Sub
err:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub AddCurMarks(ByVal X As Long, ByVal Y As Long)
    Dim objPicMarks As New clsPicMarks
    Dim lTemp As DicomLabel
    Dim lngBound As Long
    Dim dblMarkZoom As Double
    Dim ImgTag As TReportImgTag
    Dim strMaxOrder As String
    
    objPicMarks.�������� = dcmMarkImage.Images(1).tag.strImgMarks
    
    strMaxOrder = 0
    If objPicMarks.Count > 0 Then
        strMaxOrder = objPicMarks.Item(objPicMarks.Count).����
    End If
    
    '����ע
    '�������͵ı�ע��һ����ֱ���Զ���ţ���һ�����ֹ����
    lngBound = objPicMarks.Count + 1
    
    objPicMarks.Add lngBound
    objPicMarks(lngBound).Selected = False
    
    If IsNumeric(mstrMarkText) Or Len(mstrMarkText) <= 0 Then
        objPicMarks(lngBound).���� = 6     'Բ�α��
    Else
        objPicMarks(lngBound).���� = 0      '0-��ʾ�ı�
    End If
        
    If mlngMarkType = imtAuto Then
        objPicMarks(lngBound).���� = Val(strMaxOrder) + 1
    ElseIf mlngMarkType = imtSpecify Then
        objPicMarks(lngBound).���� = mstrMarkText
    Else
        Exit Sub
    End If
    
    '�㼯û������
    Set lTemp = New DicomLabel
    lTemp.Left = X
    lTemp.Top = Y
    lTemp.Width = 20
    lTemp.Height = 20
    lTemp.ImageTied = True
    lTemp.Rescale dcmMarkImage.Images(1)
    
    dblMarkZoom = dcmMarkImage.Images(1).SizeX / Val(GetReportImagePro(dcmMarkImage.Images(1).tag.strPros, "width")) * Screen.TwipsPerPixelX
    
    objPicMarks(lngBound).X1 = lTemp.Left / dblMarkZoom
    objPicMarks(lngBound).Y1 = lTemp.Top / dblMarkZoom
    objPicMarks(lngBound).X2 = objPicMarks(lngBound).X1
    objPicMarks(lngBound).Y2 = objPicMarks(lngBound).Y1
    objPicMarks(lngBound).���ɫ = glngColor(lngBound Mod 9 + 1)
    objPicMarks(lngBound).��䷽ʽ = -2
    '����ɫ���գ�����ɫ����
    objPicMarks(lngBound).���� = 1
    objPicMarks(lngBound).�߿� = 1
    
    Set objPicMarks(lngBound).���� = New StdFont '  "����"
    
    Call DrawMarks(dcmMarkImage.Images(1), objPicMarks, dblMarkZoom)
    
    ImgTag = dcmMarkImage.Images(1).tag
    ImgTag.strImgMarks = objPicMarks.��������
    
    dcmMarkImage.Images(1).tag = ImgTag

    Call EnterModify(, , True)
End Sub

Private Sub dcmMarkImage_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim lTemp As DicomLabel
    Dim strNum As Integer
    Dim objDcmLabs As DicomLabels
    Dim strErr As String

On Error GoTo errhandle
    
    If dcmMarkImage.Images.Count <= 0 Then Exit Sub
    If mblnIsEditable = False Then Exit Sub
    
    If Button = 2 Then
        Set objDcmLabs = dcmMarkImage.LabelHits(X, Y, False, False, True)
        If objDcmLabs.Count > 0 Then
            menuLab.tag = dcmMarkImage.Images(1).Labels.IndexOf(objDcmLabs.Item(objDcmLabs.Count))
            PopupMenu menuLab, 2
            
        End If
        
        Exit Sub
    End If

    If mlngMarkType = imtNormal Then Exit Sub
    If Button = 1 And picImageBack.MousePointer = 99 Then
        '����ע
        Call AddCurMarks(X, Y)
    End If
Exit Sub
errhandle:
    strErr = err.Description
    
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "��ʾ"
End Sub


Private Sub dcmReportImg_Click()
    Dim strErr As String
On Error GoTo errhandle
    If dcmReportImg.Images.Count <= 0 Then Exit Sub
    If mlngSelReportImgIndex <= 0 Then Exit Sub

    picReportImgOper.Left = (dcmReportImg.Width - picReportImgOper.Width) / 2
    picReportImgOper.Top = dcmReportImg.Height - picReportImgOper.Height

    picReportImgOper.Visible = mblnIsEditable
      
Exit Sub
errhandle:
    strErr = err.Description
    
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "��ʾ"
End Sub


Private Sub cmdOper_Click(Index As Integer)
    Dim strErr As String
On Error GoTo errhandle
    Select Case Index
        Case 0  'ɾ������
            Call DelRepImage
            
        Case 1  'ǰ��
            If mlngSelReportImgIndex <= 1 Then
'                MsgboxH GetRootHwnd, "������ǰ�ƶ���", vbOKOnly, "��ʾ"
                Exit Sub
            End If
            
            dcmReportImg.Images.Move mlngSelReportImgIndex, mlngSelReportImgIndex - 1
            
            DrawBorder dcmReportImg.Images(mlngSelReportImgIndex - 1), 0
            
            mblnIsModifyImage = True
        Case 2  '����
            If mlngSelReportImgIndex >= dcmReportImg.Images.Count Then
'                MsgboxH GetRootHwnd, "��������ƶ���", vbOKOnly, "��ʾ"
                Exit Sub
            End If
            
            dcmReportImg.Images.Move mlngSelReportImgIndex, mlngSelReportImgIndex + 1
            
            DrawBorder dcmReportImg.Images(mlngSelReportImgIndex + 1), 0
            
            mblnIsModifyImage = True
        Case 3  '�Զ�
            Call Mark(imtAuto)
            
        Case 4, 5, 6, 7 '���1,2,3,4
            Call Mark(imtSpecify, Index - 3)
            
    End Select
    
    mlngSelReportImgIndex = 0
    
    picReportImgOper.Visible = False
'    picMarkImgOper.Visible = False
Exit Sub
errhandle:
    strErr = err.Description
    
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "��ʾ"
End Sub
 
 

Private Sub dcmReportImg_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim blnIsReportImgArea As Boolean
    Dim lngBound As Long
    
    If dcmReportImg.Images.Count <= 0 Then Exit Sub
    If mlngSelReportImgIndex <= 0 Then Exit Sub

    blnIsReportImgArea = False
    lngBound = 135
    
    '�ж��Ƿ���Ҫ��ʾͼ��
    If (lngBound <= X * Screen.TwipsPerPixelX) And (X * Screen.TwipsPerPixelX <= dcmReportImg.Width - lngBound) And _
       (lngBound <= Y * Screen.TwipsPerPixelY) And (Y * Screen.TwipsPerPixelY <= dcmReportImg.Height - lngBound) Then
        blnIsReportImgArea = True
    End If

    picReportImgOper.Visible = blnIsReportImgArea And mblnIsEditable
    
    'ע�⣺���¼��У����ܶ����������������߻������ʾ�ı���ͼ��ش���ť���ܲ���
    
End Sub

Private Sub dcmReportImg_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim i As Integer
    Dim strErr As String
On Error GoTo errhandle
    If Button = 2 Then
        '����Ҽ�
        If mblnIsEditable Then
            If dcmReportImg.Images.Count <= 0 Then Exit Sub
            PopupMenu menuReport, 2
        End If
    Else
        mlngSelReportImgIndex = dcmReportImg.ImageIndex(X, Y)
        
        If mlngSelReportImgIndex <= 0 Or mlngSelReportImgIndex > dcmReportImg.Images.Count Then Exit Sub
        
        For i = 1 To dcmReportImg.Images.Count
            Call DrawBorder(dcmReportImg.Images(i), 0)
        Next
            
        Call DrawBorder(dcmReportImg.Images(mlngSelReportImgIndex), ColorConstants.vbRed, True)
    End If
    
'    RaiseEvent OnMouseUp(Button, Shift, x, y)
Exit Sub
errhandle:
    strErr = err.Description
    
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "��ʾ"
End Sub

Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
On Error Resume Next
    Call HideCharInput
End Sub

Private Sub dkpMain_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
On Error Resume Next
    Call HideCharInput
End Sub

Private Sub HideCharInput()
'�����ַ�¼��
    Dim blnHide As Boolean
    
    If mrtbActive Is Nothing Then Exit Sub

    blnHide = False
    If Not ActiveControl Is Nothing Then
        Select Case ActiveControl.hwnd
            Case picDesc.hwnd
                If rtb����.Visible And rtb����.Locked = False Then
                    rtb����.SetFocus
                Else
                    blnHide = True
                End If
            Case picOpin.hwnd
                If rtb���.Visible And rtb���.Locked = False Then
                    rtb���.SetFocus
                Else
                    blnHide = True
                End If
            Case picAdvi.hwnd
                If rtb����.Visible And rtb����.Locked = False Then
                    rtb����.SetFocus
                Else
                    blnHide = True
                End If
            Case picImageBack.hwnd
                blnHide = True
        End Select
    Else
        blnHide = mrtbActive.Locked
    End If
    
    If blnHide Then picChar.Visible = False
End Sub

Private Sub menuLab_Del_Click()
'ɾ����ע\
    Dim strErr As String
On Error GoTo errhandle
    Dim objLab As DicomLabel
    Dim strDelTag As String
    Dim strMarks As String
    Dim objReportImgTag As TReportImgTag
    Dim i As Long
    Dim aryMark() As String
    Dim strRemoveMarks  As String
    Dim objLinkLab As DicomLabel
    
    If menuLab.tag = "" Then Exit Sub
    If dcmMarkImage.Images.Count <= 0 Then Exit Sub
    
    Set objLab = dcmMarkImage.Images(1).Labels(menuLab.tag)
    If objLab Is Nothing Then Exit Sub
    
    strDelTag = ""
    
    If Not objLab.TagObject Is Nothing Then
        strDelTag = objLab.TagObject.Text
        Set objLinkLab = objLab.TagObject
    End If
    
    If strDelTag = "" Then strDelTag = objLab.Text
    
    Call dcmMarkImage.Images(1).Labels.Remove(menuLab.tag)
    
    If Not objLinkLab Is Nothing Then
        Call dcmMarkImage.Images(1).Labels.Remove(dcmMarkImage.Images(1).Labels.IndexOf(objLinkLab))
    End If
    
    strMarks = dcmMarkImage.Images(1).tag.strImgMarks
    
    strRemoveMarks = ""
    If Len(strMarks) > 0 Then
        aryMark = Split(strMarks, "0|6|")
        For i = 0 To UBound(aryMark)
            If aryMark(i) <> "" Then
                If Val(Mid(aryMark(i), 1, 2)) = Val(strDelTag) Then
                    strRemoveMarks = "0|6|" & aryMark(i)
                    Exit For
                End If
            End If
        Next
    End If
    
    objReportImgTag = dcmMarkImage.Images(1).tag
    strMarks = Replace(strMarks, strRemoveMarks, "")
    
    If Len(strMarks) > 0 Then
        If Right(strMarks, 2) = "||" Then
            strMarks = Mid(strMarks, 1, Len(strMarks) - 2)
        End If
    End If
    objReportImgTag.strImgMarks = strMarks
    
  
    
    dcmMarkImage.Images(1).tag = objReportImgTag
    
    Call dcmMarkImage.Images(1).Refresh(False)
    
    menuLab.tag = ""
    
    Call EnterModify(, , True)
Exit Sub
errhandle:
    strErr = err.Description
    
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "��ʾ"
End Sub

Private Sub menuReport_Del_Click()
    Call cmdOper_Click(0)
End Sub

Private Sub menuReport_Last_Click()
    Call cmdOper_Click(1)
End Sub

Private Sub menuReport_Next_Click()
    Call cmdOper_Click(2)
End Sub

Private Sub mobjMarkProcessV2_OnSaveImage(ByVal emImageType As TImageType, dcmImage As DicomObjects.DicomImage)
'    Dim reportImgTag As TReportImgTag
'
'    If emImageType <> mtTagImage Then Exit Sub
'    If dcmMarkImage.Images.Count <= 0 Then Exit Sub
'
'    reportImgTag = dcmMarkImage.Images(1).tag
'
'    dcmMarkImage.Images.Clear
'
'    reportImgTag.strImgMarks = ""
'    dcmImage.tag = reportImgTag
'
'    Call DrawBorder(dcmImage, 0)
'
'    dcmMarkImage.Images.Add dcmImage
'    dcmMarkImage.Images(1).tag = reportImgTag
'
'    mblnIsModifyMarks = True
End Sub

Private Sub picDesc_Resize()
On Error Resume Next
    Call ResizeEdit(rtb����, picDesc)
    
    If rtb����.Visible And rtb����.Locked = False Then
        Call SyncWordChar
    End If
End Sub

Private Sub picOpin_Resize()
On Error Resume Next
    Call ResizeEdit(rtb���, picOpin)
    
    If rtb���.Visible And rtb���.Locked = False Then
        Call SyncWordChar
    End If
End Sub


Private Sub picAdvi_Resize()
On Error Resume Next
    Call ResizeEdit(rtb����, picAdvi)
    
    If rtb����.Visible And rtb����.Locked = False Then
        Call SyncWordChar
    End If
End Sub


Private Sub SyncWordChar()
'��ʾ�ַ�¼����
    Dim p As POINTAPI
    Dim p2rect As RECT
    Dim objPic As PictureBox
    Dim strOutlineTitle As String
    
    If mrtbActive Is Nothing Then Exit Sub
    
    Select Case mrtbActive.Name
        Case "rtb����"
            Set objPic = picDesc
            strOutlineTitle = mstrDescTitle
        Case "rtb���"
            Set objPic = picOpin
            strOutlineTitle = mstrOpinTitle
        Case "rtb����"
            Set objPic = picAdvi
            strOutlineTitle = mstrAdviTitle
    End Select
    
    GetWindowRect objPic.hwnd, p2rect
    
    p.X = 0
    p.Y = p2rect.Top
    
    ScreenToClient UserControl.hwnd, p
    p.Y = ScaleY(p.Y, vbPixels, vbTwips)
    
    picChar.Left = Len(strOutlineTitle) * (TextWidth("��") + 90)
    picChar.Top = p.Y - objPic.Top + 10 - 15
    picChar.Width = rtb����.Width - picChar.Left
    picChar.Height = TextHeight("��") + 120
    
    picChar.Visible = True
End Sub

Private Sub picImageBack_Resize()

On Error Resume Next
    
    If ucSplitter1.Visible Then
        Call ucSplitter1.RePaint(False)
    Else
        If dcmReportImg.Visible Or Val(dcmReportImg.tag) <> 0 Then
            dcmReportImg.Width = picImageBack.Width
            dcmReportImg.Height = picImageBack.Height
        End If
    End If
  
    Call CalcImgView
Exit Sub
errhandle:
'    MsgboxH GetRootHwnd, err.Description, vbOKOnly, "��ʾ"
 
End Sub


Private Sub CalcImgView()
    Dim iCols As Integer, iRows As Integer
    
    If dcmReportImg.Images.Count = 1 Then
        dcmReportImg.MultiColumns = 1
        dcmReportImg.MultiRows = 1
    
        Exit Sub
    End If
    
On Error Resume Next
      
    '����ͼ����ʾ����
    ResizeRegion dcmReportImg.Images.Count, dcmReportImg.Width, dcmReportImg.Height, iRows, iCols

    dcmReportImg.MultiColumns = iCols
    dcmReportImg.MultiRows = iRows
    
    If dcmReportImg.Images.Count > 0 Then
        dcmReportImg.CurrentIndex = 1
    Else
        dcmReportImg.CurrentIndex = 0
    End If
End Sub

Private Sub EnterModify(Optional ByVal blnIsTextModify As Boolean = False, _
    Optional ByVal blnIsImageModify As Boolean = False, Optional ByVal blnIsMarkModify As Boolean = False)
'�����޸�״̬
    Dim strMsg As String
    
    If mlngReportID = 0 And (blnIsTextModify Or blnIsImageModify Or blnIsMarkModify) Then
        If LockEditor(strMsg) = False Then
            '����ʧ�ܵĴ���
            ResetContext
            'ResetEditState
            Call ConfigFaceState
            
            MsgboxH GetRootHwnd, strMsg, vbOKOnly, "��ʾ"
            
            
            
            Exit Sub
        End If
    End If
    
    If blnIsTextModify Then mblnIsModifyText = blnIsTextModify
    If blnIsImageModify Then mblnIsModifyImage = blnIsImageModify
    If blnIsMarkModify Then mblnIsModifyMarks = blnIsMarkModify
    

End Sub


Private Sub rtb����_Change()
    Dim strErr As String
On Error GoTo errhandle
    If mblnIsLoadData = False Then Exit Sub
    If mblnIsEditable = False Then Exit Sub
    
    Call EnterModify(True)
Exit Sub
errhandle:
    strErr = err.Description
    
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "��ʾ"
End Sub
 

Private Sub rtb����_DblClick()
    Dim strErr As String
On Error GoTo errhandle
    If mblnIsEditable = False Then Exit Sub
    If Trim(rtb����.Text) = "" Then Exit Sub
    
    timerTmp.Enabled = True
'    Call richTextBoxShowElements(rtb����, Parent)
Exit Sub
errhandle:
    strErr = err.Description
    
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "��ʾ"
End Sub

Private Sub rtb����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyV And Shift = 2 Then
        Call ClipbrdFormat
    End If
End Sub

Private Sub rtb����_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Call ClipbrdFormat
    End If
End Sub

Private Sub rtb����_Change()
    Dim strErr As String
On Error GoTo errhandle
    If mblnIsLoadData = False Then Exit Sub
    If mblnIsEditable = False Then Exit Sub
    
    Call EnterModify(True)
Exit Sub
errhandle:
    strErr = err.Description
    
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "��ʾ"
End Sub

Private Sub rtb����_DblClick()
    Dim strErr As String
On Error GoTo errhandle
    If mblnIsEditable = False Then Exit Sub
    If Trim(rtb����.Text) = "" Then Exit Sub
    
    timerTmp.Enabled = True
'    Call richTextBoxShowElements(rtb����, Parent)
Exit Sub
errhandle:
    strErr = err.Description
    
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "��ʾ"
End Sub

Private Sub rtb����_GotFocus()
On Error Resume Next
    Set mrtbActive = rtb����
    
    If mrtbActive.Visible And mrtbActive.Locked = False Then
        Call SyncWordChar
    Else
        picChar.Visible = False
    End If
    
    RaiseEvent OnOutlineChange(otDesc)
End Sub
 

Private Sub rtb����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyV And Shift = 2 Then
        Call ClipbrdFormat
    End If
End Sub

Private Sub rtb����_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Call ClipbrdFormat
    End If
End Sub

Private Sub rtb���_Change()
    Dim strErr As String
On Error GoTo errhandle
    If mblnIsLoadData = False Then Exit Sub
    If mblnIsEditable = False Then Exit Sub
    
    Call EnterModify(True)
Exit Sub
errhandle:
    strErr = err.Description
    
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "��ʾ"
End Sub

Private Sub rtb���_DblClick()
    Dim strErr As String
On Error GoTo errhandle
    If mblnIsEditable = False Then Exit Sub
    If Trim(rtb���.Text) = "" Then Exit Sub
    
    timerTmp.Enabled = True
'    Call richTextBoxShowElements(rtb���, Parent)
Exit Sub
errhandle:
    strErr = err.Description
    
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "��ʾ"
End Sub

Private Sub rtb���_GotFocus()
On Error Resume Next
    Set mrtbActive = rtb���
    
    If mrtbActive.Visible And mrtbActive.Locked = False Then
        Call SyncWordChar
    Else
        picChar.Visible = False
    End If
    
    RaiseEvent OnOutlineChange(otOpin)
End Sub

Private Sub rtb����_GotFocus()
On Error Resume Next
    Set mrtbActive = rtb����
    
    If mrtbActive.Visible And mrtbActive.Locked = False Then
        Call SyncWordChar
    Else
        picChar.Visible = False
    End If
    
    RaiseEvent OnOutlineChange(otAdvi)
End Sub

Private Sub rtb���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyV And Shift = 2 Then
        Call ClipbrdFormat
    End If
End Sub

Private Sub rtb���_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Call ClipbrdFormat
    End If
End Sub

Private Sub timerTmp_Timer()
On Error GoTo errhandle
    timerTmp.Enabled = False
    If mrtbActive Is Nothing Then Exit Sub
    
    Call richTextBoxShowElements(mrtbActive, Parent)
    
Exit Sub
errhandle:
    timerTmp.Enabled = False
End Sub

Private Sub UserControl_Initialize()
    mblnIsLoadData = False
     
    mblnVisibleSpecialty = True     '����
    mblnUseImgSign = False
    
    mstrDescTitle = "�������"
    mstrOpinTitle = "������"
    mstrAdviTitle = "��    ��"
    
    mlngSelReportImgIndex = 0
    mintEditFontSize = 0
    mblnTechReptSame = False
    
    Set mobjSpePlugin = Nothing
End Sub

 
Private Sub UserControl_InitProperties()
On Error GoTo errhandle
    Call Relayout
Exit Sub
errhandle:

End Sub

Private Sub UserControl_Paint()
On Error GoTo errhandle
    If mblnIsLoadData = False Then
        'TODO:��������...
    End If
Exit Sub
errhandle:

End Sub

Private Sub UserControl_Resize()
On Error GoTo errhandle
    picContainer.Left = 0
    picContainer.Top = 0
    picContainer.Width = ScaleWidth
    picContainer.Height = ScaleHeight - picState.Height
    
    picState.Left = 0
    picState.Top = picContainer.Height
    picState.Width = ScaleWidth
    
    labFmt.Width = picState.ScaleWidth - labFmt.Left
    
    labSign.Width = picState.ScaleWidth - labSign.Left
    labEditState.Left = picState.ScaleWidth - labEditState.Width - 50
Exit Sub
errhandle:
    Debug.Print "UserControl_Resize:" & err.Description
End Sub

Public Sub Destory()

    ucSplitter1.Destory
    
    If Not mobjMarkProcessV2 Is Nothing Then
        Unload mobjMarkProcessV2
    End If
    
    Set mobjMarkProcessV2 = Nothing
    Set mobjSpePlugin = Nothing
    Set mObjNotify = Nothing
    Set mrtbActive = Nothing
End Sub


Private Sub UserControl_Terminate()
On Error GoTo errhandle

    SetParent picImageBack.hwnd, 0
    SetParent picDesc.hwnd, 0
    SetParent picOpin.hwnd, 0
    SetParent picAdvi.hwnd, 0
    
    If Not mobjSpePlugin Is Nothing Then SetParent mobjSpePlugin.hwnd, 0
    
    dkpMain.CloseAll
    
    Call Destory
Exit Sub
errhandle:
    Debug.Print "ucReportEditor_Terminate Err:" & err.Description
End Sub

Private Function IsHaveContent() As Boolean
'����:�����Ƿ��������
On Error GoTo errH
    IsHaveContent = True
    If rtb����.Text = "" And rtb����.Text = "" And rtb���.Text = "" Then IsHaveContent = False
    
    Exit Function
errH:
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
End Function
 
