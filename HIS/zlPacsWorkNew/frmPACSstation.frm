VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#1.0#0"; "zlIDKind.ocx"
Begin VB.Form frmPACSStation 
   Caption         =   "Ӱ��ҽ������վ"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11325
   Icon            =   "frmPACSstation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   11325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox PicWindow 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   1725
      ScaleHeight     =   3735
      ScaleWidth      =   9510
      TabIndex        =   1
      Top             =   2670
      Width           =   9510
      Begin VB.PictureBox picInfo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   625
         Left            =   0
         ScaleHeight     =   630
         ScaleWidth      =   9465
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   15
         Width           =   9465
         Begin VB.Frame fraRegist 
            Height          =   700
            Left            =   0
            TabIndex        =   7
            Top             =   -75
            Width           =   1980
            Begin VB.ComboBox cboTimes 
               Height          =   300
               Left            =   60
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   325
               Width           =   1875
            End
            Begin VB.Label lblRegist 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����¼(&G)"
               Height          =   180
               Left            =   105
               TabIndex        =   9
               Top             =   150
               Width           =   990
            End
         End
         Begin VB.Frame fraInfo 
            Height          =   700
            Left            =   1980
            TabIndex        =   4
            Top             =   -75
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
               Left            =   6825
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
         Left            =   0
         TabIndex        =   2
         Top             =   360
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
      TabIndex        =   12
      Top             =   495
      Width           =   4495
      Begin VB.TextBox txtFilter 
         Appearance      =   0  'Flat
         Height          =   250
         Left            =   870
         TabIndex        =   14
         ToolTipText     =   "*����ţ�+סԺ�ţ�����ѡ���ҷ�ʽ��������ɺ�ֱ�ӻس���ʼ����"
         Top             =   45
         Width           =   1485
      End
      Begin VB.TextBox txtAppend 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDD6C6&
         BorderStyle     =   0  'None
         Height          =   2100
         Left            =   630
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   1605
         Width           =   2010
      End
      Begin VSFlex8Ctl.VSFlexGrid vsList 
         Height          =   2685
         Left            =   450
         TabIndex        =   15
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
         ExplorerBar     =   0
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
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "ѡ����Ŀ(*)"
            Top             =   270
            Visible         =   0   'False
            Width           =   270
         End
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
   Begin zlIDKind.IDKind IDKind 
      Bindings        =   "frmPACSstation.frx":1CFA
      Height          =   360
      Left            =   4815
      TabIndex        =   11
      Top             =   225
      Visible         =   0   'False
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   635
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
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
            Picture         =   "frmPACSstation.frx":1D0E
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7832
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
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
            Picture         =   "frmPACSstation.frx":25A2
            Key             =   "����"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSstation.frx":2B3C
            Key             =   "סԺ"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSstation.frx":3416
            Key             =   "����"
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSstation.frx":3570
            Key             =   "Ӱ��"
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSstation.frx":3CEA
            Key             =   "�ѽ�"
            Object.Tag             =   "5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSstation.frx":4084
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
            Picture         =   "frmPACSstation.frx":41DE
            Key             =   "��ѡ����"
            Object.Tag             =   "90000"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSstation.frx":4778
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
      Bindings        =   "frmPACSstation.frx":4D12
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmPACSStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mcol
    Col���� = 0: Col��Դ: Col����: Col����: Col����: Col����: Col������: Col�Ա�: Col����: Col��ʶ��: Colҽ������: Col��λ����: Colִ�м�: Col���ʱ��: Col����ʱ��: Col����ҽ��
    
    Col��� = 16: Col����: ColӤ��: Col�Ǽ���: Col������: Col�����: Col��ӡ��Ƭ: Col�������: Col��ɫͨ��: Col�����ӡ: Col������: Col������: Col��鼼ʦ: Col��ͼʱ��: Col�������
    
    ColӰ����� = 31: Col����ID: Col��ҳID: Col�Һŵ�: Col���˿���ID: Colҽ��ID: Col���ͺ�: Col���UID: Col���״̬: ColNO: Col��¼����: Colת��: Col����: Col��ǰ����ID '��31�п�ʼ����ʾ
End Enum

Private Enum FilterID
    ID_���� = 4001: ID_סԺ = 4002: ID_��� = 4003: ID_���� = 4004
    ID_���� = 4005: ID_�ѽ� = 4006: ID_δ�� = 4007: ID_�Ǽ� = 4008
    ID_���� = 4009: ID_���� = 4010: ID_��� = 4011: ID_��� = 4012
    ID_���ҷ�ʽ = 4013: ID_����ֵ = 4014: ID_��ʼ���� = 4015: ID_����סԺ = 4016
End Enum

Private mblncmd���� As Boolean, mblncmdסԺ As Boolean, mblncmd��� As Boolean, mblncmd���� As Boolean, mblncmd�ѽ� As Boolean, mblncmdδ�� As Boolean
Private mblncmd�Ǽ� As Boolean, mblncmd���� As Boolean, mblncmd���� As Boolean, mblncmd��� As Boolean, mblncmd��� As Boolean, mblncmd���� As Boolean
Private mstrFirstTab As String '�״���ʾ��ҳ��

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
Private WithEvents mfrmPacsReport As frmReport     'PACS����༭����Ƕ��������Ĵ���
Attribute mfrmPacsReport.VB_VarHelpID = -1
Private WithEvents mfrmPacsReportDock As frmReport     'PACS����༭��,��������
Attribute mfrmPacsReportDock.VB_VarHelpID = -1
Private WithEvents mobjReport As zlRichEPR.cDockReport  '�������
Attribute mobjReport.VB_VarHelpID = -1
Private mobjExpense As zlCISKernel.clsDockExpense       '���ö���
Private WithEvents mobjInAdvice As zlCISKernel.clsDockInAdvices    'סԺҽ������
Attribute mobjInAdvice.VB_VarHelpID = -1
Private WithEvents mobjOutAdvice As zlCISKernel.clsDockOutAdvices  '����ҽ������
Attribute mobjOutAdvice.VB_VarHelpID = -1
Private mobjInEPRs As zlRichEPR.cDockInEPRs             'סԺ��������
Private mobjOutEPRs As zlRichEPR.cDockOutEPRs           '���ﲡ������
Private WithEvents mobjPacsCore As zl9PacsCore.clsViewer           '��Ƭվ����
Attribute mobjPacsCore.VB_VarHelpID = -1
Private mobjPacsReportArry() As frmReport                   'PACS����༭������
Private mobjQueue As zl9QueueManage.clsQueueManage          '�Ŷӽк�
'���ڱ���
Private mlngCur����ID As Long                               '��ǰ����ID
Private mstrCur���� As String                               '��ǰ���� ����-����
Private mstrCanUse���� As String                            '��ǰ���ÿ���  ID_����-����
Private mstrCurFindtype As String                           '��������
Private mblnInitOk As Boolean                               '��ʼ�����
Private mstrPrivs As String, mlngModul As Long              'ģ��ţ���ģ��Ȩ��
'���̿��Ʊ���
Private mblnFinishCommit As Boolean                         '�ޱ��������,�Ƿ������ٴ�ȷ��
Private mblnCompleteCommit As Boolean                       '��˺������ٴ�ȷ��
Private mblnIgnoreResult As Boolean                         '���������� '=true ����
Private mintResultInput As Integer                          '��ʾ���������Ժ�Ӱ������
Private mblnShowImgAtReport As Boolean                      '�򿪱���ʱ�򿪹�Ƭվ
Private mblnReportWithImage As Boolean                      '��ͼ�����д���棬��ͼ�񲻿�д����
Private mblnReportWithResult As Boolean                     '��Ӱ�����Ϊ����
Private mblnLocalizerBackward As Boolean                    '��λƬ����
Private mblnPacsReport As Boolean                           '�Ƿ�ʹ��PACS����༭����Fasleʱʹ�õ��Ӳ����༭��
Private mblnPrintCommit As Boolean                          '��ӡ��ֱ�����
Private mBeforeDays As Integer                              'Ĭ�ϲ�ѯ������
Private mlngRefreshInterval As Long                         '�����б��Զ�ˢ�¼��
Private mblnUseQueue As Boolean         '�Ƿ������Ŷӽк�
Private mAstr��������() As String       '�������ƣ�ִ�м������
'�������ز���
Private mstrRoom As String                                  'ֻ����ִ�м��ڵĲ���
Private mblnPatTrack As Boolean                             '�Ƿ�Խ����˽��и���
Private mblnֱ�Ӽ�� As Boolean                             '�ǼǺ�ֱ�ӽ�����
Private mblnNoShowCancel As Boolean                         '����ʾȡ���ļ��
Private mblnMoved As Boolean                                '��ǰʱ������Ƿ�ת�ƹ�
Private mblnOpenReport As Boolean                           '��ʼ����Զ��򿪱���
Private mblnTechReptSame As Boolean                         'ֻ����д�Լ����ı���
Private mblnUse3D As Boolean                                '�Ƿ�������ά�ؽ�����
Private mstr3DExeDir As String                              '��ά�ؽ�����·��
Private mstr3DPara As String                                '��ά�ؽ�����
Private mstr3DFunctions As String                           '��ά�ؽ�����
'������������
Private Type Type_SQLCondition
    ��ʼʱ�� As Date
    ����ʱ�� As Date
    ʱ������ As Integer                                 'ʱ���ѯ��ʽ 1=�����ʱ�䡢2=������ʱ��
    ���ݺ� As String
    ����� As Long
    סԺ�� As Long
    ���￨ As String
    ���� As String
    ���� As Long
    ���֤  As String
    IC�� As String
    ���˿��� As Long
    �걾��λ As String
    ���ҽ�� As String
    ���ҽ�� As String
    ������� As String
    ������� As Boolean
    Ӱ������ As String
    ��鼼ʦ As String
    ������ As String
    Ӱ����� As String
    ������� As String
    ������ As String
    ���� As String
    ��� As String
End Type
Private SQLCondition As Type_SQLCondition

'��ʷ��¼����ʾ
Private mblnIsHistory As Boolean
Private mlngHOrderID As Long
Private mlngHSendNo As Long
Private mstrHStudyUID As String
Private mblnHMoved As Boolean


Private Sub Menu_File_Excel_click()
Dim bytMode As Byte
    On Error GoTo ErrHandle
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL

    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vsList
    objPrint.Title.Text = "��鲡���嵥"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    bytMode = zlPrintAsk(objPrint)
    If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_File_BatPrint()
Dim cbrControl As CommandBarControl, strReturn As String, l As Integer
Dim objReportPrint As New zlRichEPR.cDockReport
    Set cbrControl = Me.cbrMain(2).FindControl(, conMenu_File_Print)
    If Not cbrControl Is Nothing Then
        cbrControl.ID = conMenu_File_BatPrint
    Else
        Exit Sub
    End If

    'ѡ����
    strReturn = frmDocPrintPatiList.Showfrm(vsList, Me)
    'ѭ������
    For l = 0 To UBound(Split(strReturn, "|"))
        objReportPrint.zlRefresh CLng(Split(strReturn, "|")(l)), mlngCur����ID
        Call objReportPrint.zlExecuteCommandBars(cbrControl)
        Call AfterPrinted(CLng(Split(strReturn, "|")(l)))
    Next
    cbrControl.ID = conMenu_File_Print
    Unload objReportPrint.zlGetForm
End Sub
Private Sub Menu_RichEPR(ByVal cbrID As Long)
    Dim cbrControl As CommandBarControl, i As Integer
    
    '����ҳ�治�ɼ�ʱ��ִ���κβ���
    If TabWindow.Selected.Tag <> "������д" Then
        For i = 0 To TabWindow.ItemCount - 1 'ѭ�����˲Ŵ���
            If TabWindow(i).Tag = "������д" And TabWindow(i).Visible = True Then TabWindow(i).Selected = True
        Next
        If TabWindow.Selected.Tag <> "������д" Then Exit Sub
    Else
        If TabWindow.Selected.Visible = False Then Exit Sub
    End If
    
    'ˢ��Ƕ��ҳ������
    If mblnPacsReport = True Then
        Call mfrmPacsReport.zlRefresh(Val(vsList.TextMatrix(vsList.Row, Colҽ��ID)), Nvl(vsList.TextMatrix(vsList.Row, Col���ͺ�), 0), mlngCur����ID, mstrPrivs, mlngModul, Me, vsList.TextMatrix(vsList.Row, Colת��) = 1)
    Else
        Call mobjReport.zlRefresh(Val(vsList.TextMatrix(vsList.Row, Colҽ��ID)), mlngCur����ID, True)
    End If
    
    '�жϰ���������
    Set cbrControl = Me.cbrMain.FindControl(, IIf(mblnPacsReport, conMenu_PacsReport_Open, cbrID))
    If cbrControl Is Nothing Then Exit Sub
    Call cbrMain_Update(cbrControl)
    If cbrControl.Enabled = False Then Exit Sub
        
    Call cbrMain_Execute(cbrControl)
End Sub
Private Sub Menu_File_Parmeter_click()
    With frmTechnicSetup
        .mlngModul = mlngModul
        .mlng����ID = mlngCur����ID
        .mstrPrivs = mstrPrivs
        .Show 1, Me
        If .mblnOK Then
            InitLocalPars
            Call RefreshList
        End If
    End With
End Sub

Private Sub Menu_Help_About_click()
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
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

Private Sub Menu_Manage_ȡ������()
'ȡ��������������ǣ�ÿ��ȡ��������ͼ��ȫ���������б���ɢ��N����ʱ��¼
Dim strFilter As String, rsTmp As ADODB.Recordset, lngAdviceID As Long, lngSendNO As Long
    On Error GoTo ErrHandle
    '��ʾ����ѡ�񴰿�
    With vsList
        lngAdviceID = Nvl(.TextMatrix(.Row, Colҽ��ID), 0)
        lngSendNO = Nvl(.TextMatrix(.Row, Col���ͺ�), 0)
    End With
    
    gstrSQL = "select 0 as ѡ��,B.����UID as ID ,B.���к�,B.��������,SUM(1) AS ͼ���� from Ӱ�����¼ A ," & _
            "Ӱ�������� B, Ӱ����ͼ�� C Where a.���UID = B.���UID And B.����UID = C.����UID" & _
            " And a.ҽ��ID = [1] and A.���ͺ�= [2] group by B.����UID,B.���к�,B.��������"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngAdviceID, lngSendNO)
    
    frmSelectMuli.ShowSelect rsTmp, "ID,3000,0,1;���к�,800,0,1;��������,2000,0,1;ͼ����,800,0,1", 0, 0, 14000, 10000, "ȡ������"
    
    If frmSelectMuli.mblnOK = True Then
        strFilter = frmSelectMuli.strFilter
        rsTmp.Filter = strFilter
        '�����ѡ�����У�����ÿһ�����е�ȡ��
        While Not rsTmp.EOF
            subCancelSeriesRelate lngAdviceID, lngSendNO, rsTmp!ID
            rsTmp.MoveNext
        Wend
        
        '����Ӱ����״̬�������ǰҽ���Ѿ�û��ͼ�񣬶��Ҽ�����Ϊ3�����޸�Ϊ2
        If vsList.TextMatrix(vsList.Row, Col���״̬) = 3 Then
            gstrSQL = "Select ���uid From Ӱ�����¼ Where  ҽ��ID=[1] And ���ͺ�=[2]"
            Set rsTmp = OpenSQLRecord(gstrSQL, Me.Caption, lngAdviceID, lngSendNO)
            If IsNull(rsTmp!���UID) Then
                gstrSQL = "Zl_Ӱ����_State(" & lngAdviceID & "," & lngSendNO & ",2)"
                zlDatabase.ExecuteProcedure gstrSQL, "ȡ������"
            End If
        End If
        
        mfrmPACSImg.zlRefresh 0, 0, mstrPrivs
        Call RefreshList '����ȡ��������ȷ����ˢ��
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_�ޱ������()
'ֻ�н����еı�����Բ����ò˵�,��Ϊ��ʱ��û��ǩ��

        On Error GoTo ErrHandle
        With vsList
            If .TextMatrix(.Row, Col������) <> "" Or .TextMatrix(.Row, Col�������) <> "" Then
                If MsgBoxD(Me, "�Ƿ��ޱ���ֱ�����,ֱ����ɽ�ɾ������д�ı���!", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
            
            If mblnFinishCommit And InStr(mstrPrivs, "������") > 0 Then '�ޱ�����ɺ������ٴ�ȷ�����,����Ҫ�м����ɵ�Ȩ��
                '�˹���,��״̬=6,���ұ���ID��Ϊ�ս�ɾ�����Ӳ�����¼
                If zlDatabase.GetPara(81, glngSys) = 1 And Not bln������Ժ(.TextMatrix(.Row, Col����ID), .TextMatrix(.Row, Col��ҳID)) And bln����δ�󻮼۵�(.TextMatrix(.Row, Colҽ��ID)) Then 'ִ�к��Զ���˻��۵���Ч�����Ҳ����ѳ�Ժ������δ��˵Ļ��۵�
                    MsgBoxD Me, "�ò����ѳ�Ժ������δ��˵Ļ��۵�������ɣ�", vbExclamation, gstrSysName
                Else
                    gstrSQL = "ZL_Ӱ����_STATE(" & .TextMatrix(.Row, Colҽ��ID) & "," & .TextMatrix(.Row, Col���ͺ�) & ",6,1)"
                End If
            Else
                gstrSQL = "ZL_Ӱ����_STATE(" & .TextMatrix(.Row, Colҽ��ID) & "," & .TextMatrix(.Row, Col���ͺ�) & ",5,1)"
            End If
        End With
        
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�ı������")
        
        If mblnPatTrack Then
            If mblnFinishCommit Then
                Call StateCheck(6)
            Else
                Call StateCheck(5)
            End If
        Else
            Call RefreshList
        End If
        Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Edit_�ޱ������()
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    If MsgBoxD(Me, "ȷ��Ҫ���˸�������", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    With vsList
            '�����ͼ������˵����Ѽ�顱��������˵����ѱ�����
            gstrSQL = "Select ���UID From Ӱ�����¼ Where ҽ��ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ���ͼ��", CLng(.TextMatrix(.Row, Colҽ��ID)))
            
            gstrSQL = "ZL_Ӱ����_STATE(" & .TextMatrix(.Row, Colҽ��ID) & "," & .TextMatrix(.Row, Col���ͺ�) & "," & IIf(Nvl(rsTemp!���UID) = "", 2, 3) & ")"
            zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    End With
    
    If mblnPatTrack Then
        Call StateCheck(2)
    Else
        Call RefreshList
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Menu_Manage_����������(Optional lngҽ��ID As Long = 0, Optional blnRefresh As Boolean = True)
'�������������̵��ã���ʱ������ҽ��ID������ҪȨ���ж�
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    If lngҽ��ID = 0 Then
        lngҽ��ID = vsList.TextMatrix(vsList.Row, Colҽ��ID)
    End If
    
    If InStr(mstrPrivs, "������") <= 0 Then Exit Sub
    
    gstrSQL = "Select a.���ͺ�,b.����ID,b.��ҳID From ����ҽ������ a,����ҽ����¼ b Where a.ҽ��id = [1] And a.ҽ��ID=b.Id"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����������", lngҽ��ID)
    
    If rsTemp.EOF = True Then Exit Sub
    
    If zlDatabase.GetPara(81, glngSys) = 1 And Not bln������Ժ(rsTemp!����ID, Nvl(rsTemp!��ҳID, 0)) And bln����δ�󻮼۵�(Nvl(lngҽ��ID)) Then 'ִ�к��Զ���˻��۵���Ч�����Ҳ����ѳ�Ժ������δ��˵Ļ��۵�
        MsgBoxD Me, "�ò����ѳ�Ժ������δ��˵Ļ��۵���������ɣ�", vbExclamation, gstrSysName
    Else
        gstrSQL = "ZL_Ӱ����_STATE(" & lngҽ��ID & "," & rsTemp!���ͺ� & ",6)"
    End If

    Call zlDatabase.ExecuteProcedure(gstrSQL, "�ı������")

    If blnRefresh Then Call StateCheck(6)
    Exit Sub

ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_ȡ��������()
    On Error GoTo ErrHandle
    With vsList
            If .TextMatrix(.Row, Colת��) = 1 Then MsgBox "�ò��˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������", vbInformation, gstrSysName: Exit Sub
            gstrSQL = "ZL_Ӱ����_STATE(" & .TextMatrix(.Row, Colҽ��ID) & "," & .TextMatrix(.Row, Col���ͺ�) & ",5)"
            zlDatabase.ExecuteProcedure gstrSQL, "ȡ��������"
    End With
    
    Call StateCheck(5)
    Exit Sub

ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_�������(ByVal lngID As Long)
    Dim iresult As Integer

    On Error GoTo ErrHandle
    Select Case lngID
        Case conMenu_Manage_Negative
            iresult = 1
        Case conMenu_Manage_Positive
            iresult = 0
    End Select
    With vsList
        gstrSQL = "ZL_Ӱ����_���(" & .TextMatrix(.Row, Colҽ��ID) & "," & iresult & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���������")
        
        If iresult = 1 Then
            Set .Cell(flexcpPicture, .Row, Col����) = imgList.ListImages("����").Picture
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_��ɫͨ��(ByVal lngID As Long)
    Dim intResult As Integer

    On Error GoTo ErrHandle
    Select Case lngID
        Case conMenu_Manage_GChannelOk
            intResult = "1"
        Case conMenu_Manage_GChannelCancel
            intResult = "0"
    End Select
    With vsList
        gstrSQL = "Zl_��ɫͨ��_Update(" & .TextMatrix(.Row, Colҽ��ID) & ",'" & intResult & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "��ɫͨ��")
        .TextMatrix(.Row, Col��ɫͨ��) = intResult
        If intResult = 1 Then
            Set .Cell(flexcpPicture, .Row, Col����) = imgList.ListImages("��ɫͨ��").Picture
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Menu_Manage_Ӱ������(ByVal lngID As Long)
    Dim strResult As String

    On Error GoTo ErrHandle
    Select Case lngID
        Case conMenu_Manage_First
            strResult = "��"
        Case conMenu_Manage_Second
            strResult = "��"
    End Select
    With vsList
        gstrSQL = "Zl_Ӱ������_Update(" & .TextMatrix(.Row, Colҽ��ID) & ",'" & strResult & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "Ӱ������")
        .TextMatrix(.Row, Col����) = strResult
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_�޸�()
    
    With frmRISRequest
        .mlngModul = mlngModul
        .mlngSendNo = vsList.TextMatrix(vsList.Row, Col���ͺ�)
        .mlngAdviceID = vsList.TextMatrix(vsList.Row, Colҽ��ID)
        .mintEditMode = IIf(vsList.TextMatrix(vsList.Row, Col���״̬) > 1, 3, 1) '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
        .mlngCurDeptId = mlngCur����ID
        .InitMvar
        .RefreshPatiInfor False 'ˢ�²���
        .mblnOK = False
        .zlShowMe Me
        If .mblnOK Then RefreshList '�ɹ�����
    End With
End Sub
Private Sub Menu_Manage_���ƵǼ�()
    With frmRISRequest
        .mlngModul = mlngModul
        .mlngSendNo = 0
        .mlngAdviceID = 0
        .mintEditMode = 0 '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
        .mlngCurDeptId = mlngCur����ID
        .mblnOK = False
        .InitMvar
        .CopyCheck vsList.TextMatrix(vsList.Row, Colҽ��ID), vsList.TextMatrix(vsList.Row, Col���ͺ�) 'ˢ�²���
        .zlShowMe Me
        If .mblnOK Then '�ɹ�����
            If mblnֱ�Ӽ�� Then
                Call StateCheck(2)
            Else
                Call RefreshList
            End If
        End If
    End With
End Sub
Private Sub Menu_Manage_�Ǽ�()
    With frmRISRequest
        .mlngModul = mlngModul
        .mlngSendNo = 0
        .mlngAdviceID = 0
        .mintEditMode = 0 '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
        .mlngCurDeptId = mlngCur����ID
        .mblnOK = False
        .InitMvar
        .zlShowMe Me
        If .mblnOK Then '�ɹ�����
            If mblnֱ�Ӽ�� Then
                Call StateCheck(2)
            Else
                Call RefreshList
            End If
        End If
    End With
End Sub
Private Sub Menu_Manage_ȡ���Ǽ�()
    On Error GoTo ErrHandle
    
    If MsgBoxD(Me, "ȷ��Ҫȡ����ǰ������" & Chr(10) & Chr(13) & "����ȡ�������Ӧ��ҽ�����ܾ�ִ�У�", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    gstrSQL = "ZL_����ҽ��ִ��_�ܾ�ִ��(" & vsList.TextMatrix(vsList.Row, Colҽ��ID) & "," & vsList.TextMatrix(vsList.Row, Col���ͺ�) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�����Ǽ�")
    Call RefreshList
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_�ٻ�ȡ��()
'���ܣ��ٻر�ȡ���ĵǼ�
    On Error GoTo errH
    
    If MsgBoxD(Me, "ȷʵҪ�ٻر�ȡ���Ǽǵ���Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    gstrSQL = "ZL_����ҽ��ִ��_ȡ���ܾ�(" & vsList.TextMatrix(vsList.Row, Colҽ��ID) & "," & vsList.TextMatrix(vsList.Row, Col���ͺ�) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Call RefreshList
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub Menu_Manage_����()
Dim blnFocusFind As Boolean
Dim rsTemp As ADODB.Recordset

    blnFocusFind = (Me.ActiveControl.Name = "txtFilter")
    With frmRISRequest
        .mlngModul = mlngModul
        .mlngSendNo = vsList.TextMatrix(vsList.Row, Col���ͺ�)
        .mlngAdviceID = vsList.TextMatrix(vsList.Row, Colҽ��ID)
        .mintEditMode = 2 '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
        .mlngCurDeptId = mlngCur����ID
        .InitMvar
        .RefreshPatiInfor True 'ˢ�²���
        .mblnOK = False
        .zlShowMe Me
        If .mblnOK Then  '�ɹ�����
            Call StateCheck(2)
            If mblnOpenReport Then Call Menu_RichEPR(conMenu_Edit_Modify)              '��ʼ����Զ��򿪱���
        End If
        If blnFocusFind Then txtFilter.SetFocus '�Զ���λ����λ��
    End With
End Sub
Private Sub Menu_Manage_ȡ������()
Dim rsTemp As ADODB.Recordset, lngAdviceID As Long
    
    On Error GoTo ErrHandle
    With vsList
        If .TextMatrix(.Row, Col���״̬) <= 1 Then Call Menu_Manage_ȡ���Ǽ�: Exit Sub '����������
        '------------------------------------��ǩ������Ҫ�Ȼ���ǩ�����ٳ���
        lngAdviceID = .TextMatrix(.Row, Colҽ��ID)
        gstrSQL = "Select Distinct B.���ʱ�� From ����ҽ������ A, ���Ӳ�����¼ B Where A.����ID=B.Id And A.ҽ��ID=[1]"
        Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡ�Ƿ�ǩ��", lngAdviceID)
        If Not rsTemp.EOF Then
            If Nvl(rsTemp!���ʱ��, "") <> "" Then 'ǩ������
                MsgBoxD Me, "��ǰ���˵ļ�鱨���Ѿ�ǩ��,����ȡ�����,���Ȼ���ǩ��!", vbExclamation, gstrSysName
                Exit Sub
            End If
        End If

        If MsgBoxD(Me, "ȡ�����μ�齫ɾ����Ӧ�ļ��ͼ��ͼ�鱨�棬�Ƿ������", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
        If .TextMatrix(.Row, Col���UID) <> "" And InStr(mstrPrivs, "���ͼ��") <= 0 Then
            MsgBoxD Me, "��û��������ͼ��Ȩ��,�������ͼ��,���в���ȡ��������!", vbInformation, gstrSysName
            Exit Sub
        End If
        
        'ȡ���Ŷ���Ϣ
        If mblnUseQueue = True Then
            Call mobjQueue.zlDelQueue(Split(mstrCur����, "-")(1) & .TextMatrix(.Row, Colִ�м�), lngAdviceID)
        End If
        
        gstrSQL = "ZL_Ӱ����_CANCEL(" & lngAdviceID & "," & .TextMatrix(.Row, Col���ͺ�) & ",1)"
        ExecuteProc gstrSQL, Me.Caption
        'ɾ��Ӱ���ļ���Ŀ¼
        RemoveCheckImages lngAdviceID, .TextMatrix(.Row, Col���ͺ�)
    End With
    
    Call StateCheck(1)
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_��Ƭ()
    If TabWindow.Selected.Tag <> "Ӱ��ͼ��" Then '��ˢ��ͼ������
        If mblnIsHistory = True Then
            Call mfrmPACSImg.zlRefresh(mlngHOrderID, mlngHSendNo, mstrPrivs, mblnHMoved)
        Else
            Call mfrmPACSImg.zlRefresh(vsList.TextMatrix(vsList.Row, Colҽ��ID), vsList.TextMatrix(vsList.Row, Col���ͺ�), mstrPrivs, vsList.TextMatrix(vsList.Row, Colת��) = 1)
        End If
    End If
    Call mfrmPACSImg.zlMenuClick("Ӱ����")
End Sub
Private Sub Menu_Manage_�Աȹ�Ƭ()
    If TabWindow.Selected.Tag <> "Ӱ��ͼ��" Then '��ˢ��ͼ������
        If mblnIsHistory = True Then
            Call mfrmPACSImg.zlRefresh(mlngHOrderID, mlngHSendNo, mstrPrivs, mblnHMoved)
        Else
            Call mfrmPACSImg.zlRefresh(vsList.TextMatrix(vsList.Row, Colҽ��ID), vsList.TextMatrix(vsList.Row, Col���ͺ�), mstrPrivs, vsList.TextMatrix(vsList.Row, Colת��) = 1)
        End If
    End If
    Call mfrmPACSImg.zlMenuClick("Ӱ��Ա�")
End Sub
            
Private Sub Menu_Manage_ͼ��ɾ��()
Dim rsTemp As ADODB.Recordset, lngAdviceID As Long, lngSendNO As Long
    
    On Error GoTo ErrHandle
    With vsList
        lngAdviceID = .TextMatrix(.Row, Colҽ��ID)
        lngSendNO = .TextMatrix(.Row, Col���ͺ�)
    End With
    
    If TabWindow.Selected.Tag <> "Ӱ��ͼ��" Then '��ˢ��ͼ������
        Call mfrmPACSImg.zlRefresh(lngAdviceID, lngSendNO, mstrPrivs, vsList.TextMatrix(vsList.Row, Colת��) = 1)
    End If
    
    gstrSQL = "select ���UID from Ӱ�����¼ where ҽ��ID =[1] and  ���ͺ� = [2]"
    Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡ���UID", lngAdviceID, lngSendNO)
    If rsTemp.EOF Then Exit Sub
    
    If MsgBoxD(Me, "�Ƿ�ȷ��Ҫɾ���ü�������Ӱ��", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    'ɾ��Ӱ���ļ���Ŀ¼
    RemoveCheckImages lngAdviceID, lngSendNO
    gstrSQL = "ZL_Ӱ����_PhotoDelete(" & lngAdviceID & "," & lngSendNO & ")"
    ExecuteProc gstrSQL, Me.Caption
    
    '����Ӱ����״̬�����������Ϊ3�����޸�Ϊ2
    If vsList.TextMatrix(vsList.Row, Col���״̬) = 3 Then
        gstrSQL = "Zl_Ӱ����_State(" & lngAdviceID & "," & lngSendNO & ",2)"
        zlDatabase.ExecuteProcedure gstrSQL, "ɾ��ͼ��"
    End If
    
    mfrmPACSImg.zlRefresh 0, 0, mstrPrivs
    Call RefreshList
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
        
Private Sub Menu_Manage_��ȡͼ��()
Dim strImageDeviceNumber As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    If TabWindow.Selected.Tag <> "Ӱ��ͼ��" Then '��ˢ��ͼ������
        Call mfrmPACSImg.zlRefresh(vsList.TextMatrix(vsList.Row, Colҽ��ID), vsList.TextMatrix(vsList.Row, Col���ͺ�), mstrPrivs, vsList.TextMatrix(vsList.Row, Colת��) = 1)
    End If
    
    strImageDeviceNumber = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPACSImageDeviceSetup", "Ĭ��Ӱ���豸", "")
    
    'û��Ĭ���豸ʱ����
    If strImageDeviceNumber = "" Then
        If MsgBoxD(Me, "û������Ĭ��Ӱ�����豸���Ƿ��������ã�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        Else
            frmPACSImageDeviceSetup.Show vbModal, Me
            strImageDeviceNumber = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPACSImageDeviceSetup", "Ĭ��Ӱ���豸", "")
            If strImageDeviceNumber = "" Then Exit Sub
        End If
    End If
    
    gstrSQL = "select �豸��,�豸��, IP��ַ,�˿ں�,����AE,�豸AE from Ӱ���豸Ŀ¼ where �豸�� = [1] "
    Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption, Mid(strImageDeviceNumber, 2))
    
    '��Ĭ���豸��ɾ������������
    If rsTemp.EOF = True Then
        MsgBoxD Me, "Ĭ���豸�ѱ�ɾ�������������ã�", vbInformation, gstrSysName
        frmPACSImageDeviceSetup.Show vbModal, Me
        Exit Sub
    End If
        
    frmPACSGetDeviceImage.ShowMe Me, rsTemp("IP��ַ"), rsTemp("�˿ں�"), rsTemp("�豸��"), Nvl(rsTemp("����AE")), Nvl(rsTemp("�豸AE")), vsList.TextMatrix(vsList.Row, Colҽ��ID)
    Call RefreshList
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_����Ӱ��()
Dim strSQL As String
Dim rsTemp As ADODB.Recordset, lngAdviceID As Long, lngSendNO As Long
    
    On Error GoTo ErrHandle
    With vsList
        lngAdviceID = .TextMatrix(.Row, Colҽ��ID)
        lngSendNO = .TextMatrix(.Row, Col���ͺ�)

        Call funRelateSeries(lngAdviceID, lngSendNO)
        '����Ӱ����״̬�����ԭ����״̬���ѱ��������޸ĳ��Ѽ�飬
        If .TextMatrix(.Row, Col���״̬) < 3 Then
            '��������Ѿ���ͼ�����޸ĳ��Ѽ��
            strSQL = "Select ���UID From Ӱ�����¼ Where ҽ��ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ���ͼ��", lngAdviceID)
            
            If Not IsNull(rsTemp!���UID) Then
                gstrSQL = "Zl_Ӱ����_State(" & lngAdviceID & "," & lngSendNO & ",3)"
                zlDatabase.ExecuteProcedure gstrSQL, "����Ӱ��"
            End If
        End If
    End With
    mfrmPACSImg.zlRefresh 0, 0, mstrPrivs
    Call RefreshList
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Dept_Select(ByVal control As XtremeCommandBars.ICommandBarControl)
    If mlngCur����ID <> control.DescriptionText Then
        mlngCur����ID = control.DescriptionText
        mstrCur���� = Split(control.Caption, "(")(0)
        Call cbrMain.RecalcLayout
        Call InitMvar
        Call InitSubForm
        Call RefreshList
    End If
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
    
    On Error GoTo ErrHandle
    Dim lngAdviceID As Long
    lngAdviceID = cboTimes.ItemData(cboTimes.ListIndex)
    If lngAdviceID = vsList.TextMatrix(vsList.Row, Colҽ��ID) Then Call vsList_RowColChange: Exit Sub '�����뵱ǰѡ��ҽ��ID��ͬʱ���ɱ���������

    mblnIsHistory = True: mlngHOrderID = lngAdviceID '�����������̵������Ⱥ�˳�������
    Call FillTxtInfor(mlngHOrderID)  '������Ϸ����˻�����Ϣ
    Call FillTxtAppend(mlngHOrderID) '������½�ҽ������
    Call ShowTab(mlngHOrderID)  '���ݲ����ṩ��ͬѡ�
    Call RefreshTabWindow(mlngHOrderID) 'ˢ���Ӵ���
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cboTimes_DropDown()
    Call SendMessage(cboTimes.Hwnd, &H160, 500, 0)
End Sub

Private Sub cbrdock_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
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
        Case ID_�Ǽ�
            mblncmd�Ǽ� = Not control.Checked
        Case ID_����
            mblncmd���� = Not control.Checked
        Case ID_����
            mblncmd���� = Not control.Checked
        Case ID_���
            mblncmd��� = Not control.Checked
        Case ID_���
            mblncmd��� = Not control.Checked
        Case ID_����סԺ
            control.Checked = Not control.Checked
            mblncmd���� = Not mblncmd����
        Case ID_���ҷ�ʽ * 100# To ID_���ҷ�ʽ * 100# + 7
            mstrCurFindtype = Split(control.Caption, "(")(0)
            If InStr(mstrCurFindtype, "�ɣÿ�") > 0 Then
                If mobjICCard Is Nothing Then
                    Set mobjICCard = CreateObject("zlICCard.clsICCard")
                End If
                txtFilter.Text = mobjICCard.Read_Card(Me)
            End If
            txtFilter.SetFocus
            cbrdock.RecalcLayout
            Exit Sub
        Case ID_��ʼ����
            With SQLCondition
                .���� = ""
                .���￨ = ""
                .����� = 0
                .סԺ�� = 0
                .���ݺ� = ""
                .���� = 0
                .���֤ = ""
                .IC�� = ""
                Select Case mstrCurFindtype
                    Case "��  ��"
                        .���� = Trim(txtFilter)
                    Case "���￨"
                        .���￨ = Trim(txtFilter)
                    Case "�����"
                        .����� = Val(txtFilter)
                    Case "סԺ��"
                        .סԺ�� = Val(txtFilter)
                    Case "���ݺ�"
                        .���ݺ� = Trim(txtFilter)
                    Case "����"
                        .���� = Val(txtFilter)
                    Case "���֤"
                        .���֤ = Trim(txtFilter)
                    Case "�ɣÿ�"
                        .IC�� = Trim(txtFilter)
                End Select
            End With
    End Select
cbrdock.RecalcLayout
Call RefreshList
End Sub

Private Sub cbrdock_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim objControl As CommandBarControl, i As Integer
    If CommandBar.Parent Is Nothing Then Exit Sub
    If CommandBar.Parent.ID = ID_���ҷ�ʽ Then
        With CommandBar.Controls
            If .Count = 0 Then '��̬�Ӳ˵�,��1λ
                Set objControl = .Add(xtpControlButton, ID_���ҷ�ʽ * 100# + 0, "�����(&1)")
                Set objControl = .Add(xtpControlButton, ID_���ҷ�ʽ * 100# + 1, "סԺ��(&2)")
                Set objControl = .Add(xtpControlButton, ID_���ҷ�ʽ * 100# + 2, "���￨(&3)")
                Set objControl = .Add(xtpControlButton, ID_���ҷ�ʽ * 100# + 3, "��  ��(&4)")
                Set objControl = .Add(xtpControlButton, ID_���ҷ�ʽ * 100# + 4, "���ݺ�(&5)")
                Set objControl = .Add(xtpControlButton, ID_���ҷ�ʽ * 100# + 5, "����(&6)")
                Set objControl = .Add(xtpControlButton, ID_���ҷ�ʽ * 100# + 6, "���֤(&7)")
                Set objControl = .Add(xtpControlButton, ID_���ҷ�ʽ * 100# + 7, "�ɣÿ�(&8)")
            End If
        End With
    End If
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
        Case ID_�Ǽ�
            control.Checked = mblncmd�Ǽ�
            control.IconId = IIf(mblncmd�Ǽ�, 90001, 90000)
        Case ID_����
            control.Checked = mblncmd����
            control.IconId = IIf(mblncmd����, 90001, 90000)
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
        Case ID_���ҷ�ʽ
            control.Caption = mstrCurFindtype
        Case ID_���ҷ�ʽ * 100# To ID_���ҷ�ʽ * 100# + 7
            control.Checked = (InStr(control.Caption, mstrCurFindtype) > 0)
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
        Case conMenu_File_PrintSet '��ӡ����
            Call zlPrintSet
           
        Case conMenu_File_Excel '�嵥��ӡ
            Call Menu_File_Excel_click
            
        Case conMenu_File_BatPrint '������ӡ
            Call Menu_File_BatPrint
            
        Case conMenu_File_Parameter '��������
            Call Menu_File_Parmeter_click
            
        Case conMenu_Cap_DevSet 'Ӱ���豸����
            frmPACSImageDeviceSetup.Show vbModal, Me
            
        Case conMenu_File_SendImg '����ͼ��
            frmPacsSendImage.ShowMe Me
            
        Case conMenu_File_Exit '�˳�
            Unload Me
            
'---------------------------���-----------------
        Case conMenu_Manage_RequestPrint * 10# + 1 To conMenu_Manage_RequestPrint * 10# + 9 '��ӡ���Ƶ���
            Call FuncBillPrint(control)
            
        Case conMenu_Manage_Regist                          '�Ǽ�
            Call Menu_Manage_�Ǽ�
            
        Case conMenu_Manage_CopyCheck                       '���ƵǼ�
            Call Menu_Manage_���ƵǼ�
            
        Case conMenu_Manage_Receive                         '����
            Call Menu_Manage_����
            
        Case conMenu_Manage_Redo                            'ȡ���Ǽ�
            Call Menu_Manage_ȡ���Ǽ�
            
        Case conMenu_Manage_ReGet                           '�ٻ�ȡ��
            Call Menu_Manage_�ٻ�ȡ��
        
        Case conMenu_Manage_ThingModi                       '�޸ĵǼ�
            Call Menu_Manage_�޸�
            
        Case conMenu_Manage_Logout                          'ȡ������
            Call Menu_Manage_ȡ������
            
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
            
        Case conMenu_Img_Delete                       'ͼ��ɾ��
            Call Menu_Manage_ͼ��ɾ��
        
        Case conMenu_Img_Query                        '���豸��ȡͼ��
            Call Menu_Manage_��ȡͼ��
        
        Case conMenu_Manage_Transfer                        '����Ӱ��
            Call Menu_Manage_����Ӱ��
            
        Case conMenu_Manage_Cancel                          'ȡ������
            Call Menu_Manage_ȡ������
        
        Case conMenu_Manage_Review                          '���
            Call Menu_Manage_���
        
        Case conMenu_Manage_Negative, conMenu_Manage_Positive                  '���������
            Call Menu_Manage_�������(control.ID)
            
        Case conMenu_Manage_First, conMenu_Manage_Second
            Call Menu_Manage_Ӱ������(control.ID)
            
        Case conMenu_Manage_GChannelOk, conMenu_Manage_GChannelCancel
            Call Menu_Manage_��ɫͨ��(control.ID)
            
        Case conMenu_Manage_ClearUp                           '�ޱ������
            Call Menu_Edit_�ޱ������
                    
        Case conMenu_Manage_Finish                          '�ޱ���ֱ�����
            Call Menu_Manage_�ޱ������
            
        Case conMenu_Manage_Complete                        '������
                Call Menu_Manage_����������
                
        Case conMenu_Manage_Undone                          'ȡ��������
            Call Menu_Manage_ȡ��������
            
        Case conMenu_Manage_ChangeDevice                    '��������豸
            Call Menu_Manage_��������豸
            
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
        Case conMenu_View_Filter * 100# To conMenu_View_Filter * 100# + UBound(Split(mstrCanUse����, "|")) '���ĵ�ǰ����
            Call Menu_Dept_Select(control)
        Case conMenu_ReportPopup * 100# + 1 To conMenu_ReportPopup * 100# + 99
            If control.Parameter <> "" Then 'ִ�з�������ǰģ��ı���
                With vsList
                    If .TextMatrix(.Row, Colҽ��ID) <> "" Then
                        Call ReportOpen(gcnOracle, Split(control.Parameter, ",")(0), Split(control.Parameter, ",")(1), Me, _
                            "NO=" & .TextMatrix(.Row, ColNO), "����=" & .TextMatrix(.Row, Col��¼����), "ҽ��id=" & .TextMatrix(.Row, Colҽ��ID), 1)
                    Else
                        Call ReportOpen(gcnOracle, Split(control.Parameter, ",")(0), Split(control.Parameter, ",")(1), Me, "", 1)
                    End If
                End With
            End If
        Case Else
            If vsList.TextMatrix(vsList.Row, Colҽ��ID) = "" Then Exit Sub
            
            Select Case TabWindow.Selected.Tag
                Case "������д"
                    'û���治�ܴ�ӡ��Ԥ��
                    If (vsList.TextMatrix(vsList.Row, Col������) = "" Or vsList.TextMatrix(vsList.Row, Col�������) = "") And (control.ID = conMenu_File_Preview Or control.ID = conMenu_File_Print) Then
                        MsgBoxD Me, "��ǰ����û�м�鱨�棬���ܲ��������飡", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    '���汻ĳ�˴򿪺��ٱ��������˱༭���޶�
                    If control.ID = conMenu_Edit_Audit Or control.ID = conMenu_Edit_Modify Or control.ID = conMenu_PacsReport_Open Or control.ID = conMenu_Edit_Delete Then
                        If CheckConcurrentReport(vsList.TextMatrix(vsList.Row, Colҽ��ID)) = False Then Exit Sub
                    End If
                    
                    '���� ֻ����д�Լ����ı���,'��������д���޶���ɾ��
                    If mblnTechReptSame = True _
                        And (control.ID = conMenu_Edit_Modify Or control.ID = conMenu_Edit_Audit Or control.ID = conMenu_Edit_Delete) _
                        And Nvl(vsList.TextMatrix(vsList.Row, Col��鼼ʦ)) <> "" _
                        And Nvl(vsList.TextMatrix(vsList.Row, Col��鼼ʦ)) <> UserInfo.���� Then
                        MsgBoxD Me, "�㲻��������ߵļ�鼼ʦ���޷�������ݱ�", vbInformation, gstrSysName
                    Else
                        If mblnPacsReport = True Then
                            If control.ID = conMenu_PacsReport_Open Then   '�򿪱��洰��
                                Call Menu_Manage_PACS����
                            Else
                                mfrmPacsReport.zlExecuteCommandBars control
                            End If
                        Else
                            mobjReport.zlExecuteCommandBars control
                        End If
                    End If
                Case "�������"
                    mobjExpense.zlExecuteCommandBars control
                Case "סԺҽ��"
                    mobjInAdvice.zlExecuteCommandBars control
                Case "����ҽ��"
                    mobjOutAdvice.zlExecuteCommandBars control
                Case "סԺ����"
                    mobjInEPRs.zlExecuteCommandBars control
                Case "���ﲡ��"
                    mobjOutEPRs.zlExecuteCommandBars control
                Case "�Ŷӽк�"
                    mobjQueue.zlExecuteCommandBars control
            End Select
    End Select
End Sub

Private Sub Menu_View_Filter_click()
    On Error GoTo ErrHandle
    
    With frmPACSFilter
        .mlngModul = mlngModul
        .mBeforeDays = mBeforeDays
        .mDept = mlngCur����ID '��ǰ����
        .Show 1, Me
        If Not .mblnOK Then Exit Sub 'û�з�������
        
        '��ʹ��ʱ������ʱ����չ̶�����
        txtFilter.Text = ""
        SQLCondition.���� = ""
        SQLCondition.���￨ = ""
        SQLCondition.����� = 0
        SQLCondition.סԺ�� = 0
        SQLCondition.���ݺ� = ""
        SQLCondition.���� = 0
        SQLCondition.���֤ = ""
        SQLCondition.IC�� = ""
        
        SQLCondition.��ʼʱ�� = Format(.DTPBegin.Value, "yyyy-MM-dd HH:mm:00")
        If Format(.DTPEnd.Value, "yyyy-MM-dd HH:mm") = Format(.DTPEnd.Tag, "yyyy-MM-dd HH:mm") Then
            SQLCondition.����ʱ�� = CDate(0) '��ʾȡ��ǰʱ��
        Else
            SQLCondition.����ʱ�� = Format(.DTPEnd.Value, "yyyy-MM-dd HH:mm:59")
        End If
        
        mblnMoved = MovedByDate(SQLCondition.��ʼʱ��)
        
        If .optFindType(0).Value = True Then 'ʱ����ҷ�ʽ��1�������ʱ�䡢2��������ʱ��
            SQLCondition.ʱ������ = 1
        Else
            SQLCondition.ʱ������ = 2
        End If
        
        If .cboPart.ListIndex <> 0 Then '���걾��λ
            SQLCondition.�걾��λ = .cboPart.Text
        Else
            SQLCondition.�걾��λ = ""
        End If
        
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
        
        
        If .cboCheckStep.ListIndex <> 0 Then '������
            SQLCondition.������ = .cboCheckStep.Text
        Else
            SQLCondition.������ = ""
        End If
        
        
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
        
        If .chk�������.Value = 1 Then
            SQLCondition.������� = True
        Else
            SQLCondition.������� = False
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
ErrHandle:
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
                    For i = 0 To UBound(Split(mstrCanUse����, "|")) 'mstrCanUse����=id_����-����|id_����-����
                        Set objControl = .Add(xtpControlButton, conMenu_View_Filter * 100# + i, Split(Split(mstrCanUse����, "|")(i), "_")(1) & "(&" & i & ")")
                        objControl.Category = "Main"
                        objControl.DescriptionText = Split(Split(mstrCanUse����, "|")(i), "_")(0)
                        If mlngCur����ID = objControl.DescriptionText Then objControl.Checked = True
                    Next
                End If
            End With
        Case Else
            Select Case Me.TabWindow.Selected.Tag
                Case "סԺҽ��"
                    mobjInAdvice.zlPopupCommandBars CommandBar
                Case "����ҽ��" '����
                    mobjOutAdvice.zlPopupCommandBars CommandBar
                Case "�������"
                    mobjExpense.zlPopupCommandBars CommandBar
            End Select
    End Select
End Sub


Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
Dim blnNoRecord As Boolean, intState As Integer, blnCancel As Boolean
    If Not mblnInitOk Then Exit Sub
    
    blnNoRecord = Val(vsList.TextMatrix(vsList.Row, Colҽ��ID)) = 0
    control.Style = xtpButtonIconAndCaption
    If Not blnNoRecord Then
        intState = Val(vsList.TextMatrix(vsList.Row, Col���״̬))
        blnCancel = vsList.TextMatrix(vsList.Row, Col������) = "�Ѿܾ�"
    End If
    
    Select Case control.ID
        Case conMenu_View_Filter * 10#
            control.Caption = "��ǰ����:" & mstrCur����
        Case conMenu_View_Filter * 100# To conMenu_View_Filter * 100# + UBound(Split(mstrCanUse����, "|"))
            control.Checked = (control.DescriptionText = mlngCur����ID)
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
                
        Case conMenu_Manage_Regist   '���Ǽ�(&I)
            If InStr(mstrPrivs, "���Ǽ�") <= 0 Then
                control.Visible = False
            End If
        Case conMenu_Manage_CopyCheck '�ٴεǼ�
            If InStr(mstrPrivs, "���Ǽ�") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = True
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Redo   'ȡ���Ǽ�(&R)
            If InStr(mstrPrivs, "���Ǽ�") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = intState <= 1 And Not blnCancel
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_ReGet   '�ٻ�ȡ��
            If Not blnNoRecord Then
                control.Enabled = blnCancel
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_ThingModi   '�޸���Ϣ(&M)
            If InStr(mstrPrivs, "���Ǽ�") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = intState <= 3 And Not blnCancel
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Receive   '��鱨��(&L)
            If InStr(mstrPrivs, "��鱨��") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = intState <= 1 And Not blnCancel
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Logout   'ȡ������(&D)
            If blnNoRecord Then
                control.Enabled = False
            ElseIf control.Parent.Type = xtpControlPopup Then
                If InStr(mstrPrivs, "ȡ������") <= 0 Then
                    control.Visible = False
                Else
                    control.Visible = True
                    control.ToolTipText = "ȡ������"
                    control.Caption = "ȡ������(&D)"
                    control.Enabled = (intState = 2 Or intState = 3)
                End If
            Else ' �������е���ȡ��������ȡ���Ǽ�,ͬһ�������ȡ���ǼǺ�ȡ����鹦��
                control.Visible = IIf(intState <= 1, InStr(mstrPrivs, "���Ǽ�") > 0, InStr(mstrPrivs, "ȡ������") > 0)
                control.Enabled = (intState = 2 Or intState = 3) Or (intState <= 1 And Not blnCancel) '���ܾ��Ĳ��ܱ��ٴξܾ�
                control.ToolTipText = IIf(intState <= 1, "ȡ���Ǽ�", "ȡ������")
                control.Caption = "ȡ��"
            End If
        Case conMenu_Manage_Transfer   '����Ӱ��(&C)
            If InStr(mstrPrivs, "ͼ�����") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = intState >= 2 And intState <= 5 '��2---5֮�����
            End If
        Case conMenu_Manage_Cancel   'ȡ������(&B)
            If InStr(mstrPrivs, "ͼ�����") <= 0 Then
                control.Visible = False
            ElseIf intState >= 2 And intState <= 5 Then
                control.Enabled = vsList.TextMatrix(vsList.Row, Col���UID) <> ""
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_First, conMenu_Manage_Second, conMenu_Manage_Quality
            If InStr(mstrPrivs, "Ӱ���ʿ�") <= 0 Then
                control.Visible = False
            ElseIf intState >= 2 And intState <= 5 Then
                control.Enabled = vsList.TextMatrix(vsList.Row, Col���UID) <> ""
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Review  '���
            If InStr(mstrPrivs, "���") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = True
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Result, conMenu_Manage_Negative, conMenu_Manage_Positive '���������(&X)
            If (InStr(GetInsidePrivs(p���Ʊ������), "������д") <= 0 And InStr(GetInsidePrivs(p���Ʊ������), "�����޶�") <= 0) Or _
                mblnIgnoreResult Then
                control.Visible = False
            Else
                control.Enabled = intState >= 2 And intState <= 5 '��2---5֮�����
            End If
        Case conMenu_Manage_GChannel, conMenu_Manage_GChannelOk, conMenu_Manage_GChannelCancel '��ɫͨ�����/ȡ��
            If InStr(mstrPrivs, "��ɫͨ��") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = intState >= 2 And intState <= 5 '��2---5֮�����
            End If
        Case conMenu_Manage_Finish   '�ޱ������(&F)
            If InStr(mstrPrivs, "�ޱ������") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = intState = 2 Or intState = 3
            End If
        Case conMenu_Manage_ClearUp   '�ޱ������(&U)
            If InStr(mstrPrivs, "�ޱ������") <= 0 Then
                control.Visible = False
            ElseIf intState = 5 Then
                control.Enabled = vsList.TextMatrix(vsList.Row, Col������) = ""
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Complete   '������(&E)
            If InStr(mstrPrivs, "������") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = (intState = 4 Or intState = 5)
            End If
        Case conMenu_Manage_Undone   'ȡ�����(&U)
            If InStr(mstrPrivs, "ȡ��������") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = intState = 6
            End If
        Case conMenu_File_SendImg  '����ͼ��
            If InStr(mstrPrivs, "�ļ�����") <= 0 Then control.Visible = False
        Case conMenu_Img_Contrast, conMenu_Img_Look     'Ӱ��Ա�,Ӱ���Ƭ
            If blnNoRecord Then control.Enabled = False: Exit Sub
            If mblnIsHistory = True Then
                control.Enabled = mstrHStudyUID <> ""
            Else
                control.Enabled = vsList.TextMatrix(vsList.Row, Col���UID) <> ""
            End If
            If control.Parent.Type <> xtpControlPopup Then control.Visible = control.Enabled
        Case conMenu_Img_3D     '��ά�ؽ�
            If InStr(mstrPrivs, "��ά�ؽ�����") <> 0 And mblnUse3D = True Then
                control.Visible = True
            Else
                control.Visible = False
            End If
            If control.Visible = True Then
                If blnNoRecord Then control.Enabled = False: Exit Sub
                If control.Parent.Type <> xtpControlPopup Then
                    control.Visible = vsList.TextMatrix(vsList.Row, Col���UID) <> ""
                    control.Enabled = control.Visible
                Else
                    control.Enabled = vsList.TextMatrix(vsList.Row, Col���UID) <> ""
                End If
            End If
        Case conMenu_Img_Delete '���ͼ��
            If InStr(mstrPrivs, "���ͼ��") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = vsList.TextMatrix(vsList.Row, Col���UID) <> ""
            Else
                control.Enabled = False
            End If
        Case conMenu_Img_Query ',��ȡͼ��
            If InStr(mstrPrivs, "���ͼ��") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = intState > 1
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_ChangeDevice    '����Ӱ���豸
            If blnNoRecord = True Then
                control.Enabled = False
            Else
                If UCase(vsList.TextMatrix(vsList.Row, ColӰ�����)) = "CR" Or _
                    UCase(vsList.TextMatrix(vsList.Row, ColӰ�����)) = "DR" Or _
                    UCase(vsList.TextMatrix(vsList.Row, ColӰ�����)) = "DX" Or _
                    UCase(vsList.TextMatrix(vsList.Row, ColӰ�����)) = "RF" Then
                    control.Enabled = True
                Else
                    control.Enabled = False
                End If
            End If
        Case conMenu_File_PrintSet     '��ӡ����(&S)
        Case conMenu_File_Preview, conMenu_File_Print '����Ԥ��(&V) �����ӡ(&P)
            control.Enabled = Not blnNoRecord
        Case conMenu_File_Excel         '�嵥��ӡ(&L)
            control.Enabled = Not blnNoRecord
        Case conMenu_File_BatPrint    ' ������ӡ(&B)
            control.Enabled = Not blnNoRecord
        Case conMenu_File_Parameter     '��������(&O)
        Case conMenu_ReportPopup, conMenu_ReportPopup * 100# + 1 To conMenu_ReportPopup * 100# + 99 '����
        Case conMenu_FilePopup, conMenu_ManagePopup, conMenu_ViewPopup, conMenu_HelpPopup
        Case conMenu_File_Exit
        Case conMenu_View_ToolBar
        Case Else
            If blnNoRecord Then control.Enabled = False: Exit Sub
            Select Case TabWindow.Selected.Tag
                Case "������д"
                    If mblnPacsReport = True Then
                        mfrmPacsReport.zlUpdateCommandBars control
                    Else
                        mobjReport.zlUpdateCommandBars control
                    End If
                Case "�������"
                    mobjExpense.zlUpdateCommandBars control
                Case "סԺҽ��"
                    mobjInAdvice.zlUpdateCommandBars control
                Case "����ҽ��"
                    mobjOutAdvice.zlUpdateCommandBars control
                Case "סԺ����"
                    mobjInEPRs.zlUpdateCommandBars control
                Case "���ﲡ��"
                    mobjOutEPRs.zlUpdateCommandBars control
            End Select

            If Not blnNoRecord Then
                'ɾ��ֻ�����ѱ���ͽ����п���
                If control.ID = conMenu_Edit_Delete And Val(vsList.TextMatrix(vsList.Row, Col���״̬)) >= 4 Then
                    control.Enabled = False
                End If
                '��ǰ�鿴�������μ�¼��˵���������
                If cboTimes.ListIndex <> -1 Then
                    If vsList.TextMatrix(vsList.Row, Colҽ��ID) <> cboTimes.ItemData(cboTimes.ListIndex) Then control.Enabled = False
                End If
                '����ɳ�����,�Լ�ҽ���б���鿴��ӡ����Ƭ�˵����������
                If Val(vsList.TextMatrix(vsList.Row, Col���״̬)) = 6 Then
                    Select Case control.ID
                        Case conMenu_Edit_MarkMap, conMenu_File_Open, conMenu_Edit_Compend, conMenu_Edit_Compend * 10# + 1 To conMenu_Edit_Compend * 10# + 3
                            control.Enabled = True
                        Case Else
                            control.Enabled = False
                    End Select
                End If
            End If
    End Select
End Sub
Private Sub InitMvar()
'����:��ʼ��ģ�鼶����,���������ʱ����һ��
    On Error GoTo err
    
    mblnIgnoreResult = GetDeptPara(mlngCur����ID, "���Խ��������", 0) = 1 '        '���Խ��������
    mblnFinishCommit = GetDeptPara(mlngCur����ID, "�ޱ�����ɺ�ֱ�����", 0) = 1 '  '�ޱ�����ɺ�ֱ�����
    mblnReportWithImage = GetDeptPara(mlngCur����ID, "��ͼ�����д����", 0) = 1 '   '��ͼ�����д����
    mblnReportWithResult = GetDeptPara(mlngCur����ID, "��Ӱ�����Ϊ����", 0) = 1 '  '��Ӱ�����Ϊ����
    mblnLocalizerBackward = GetDeptPara(mlngCur����ID, "��λƬ����", 0) = 1 '       '��λƬ����
    mblnCompleteCommit = GetDeptPara(mlngCur����ID, "��˺�ֱ�����", 0) = 1 '      '��˺�ֱ�����
    mBeforeDays = GetDeptPara(mlngCur����ID, "Ĭ�Ϲ�������", 2) '                   'Ĭ�Ϲ�������
    mblnTechReptSame = GetDeptPara(mlngCur����ID, "ֻ����д�Լ����ı���", 0) = 1  'ֻ����д�Լ����ı���
    mblnPacsReport = GetDeptPara(mlngCur����ID, "����༭��", 0) = 1 '              '����༭��
    mintResultInput = GetDeptPara(mlngCur����ID, "��ʾ������", 1)    '              '��ʾ������
    mblnPrintCommit = GetDeptPara(mlngCur����ID, "��ӡ��ֱ�����", 0) = 1 '         '��ӡ��ֱ�����
    If InStr(mstrPrivs, "�Ŷӽк�") > 0 Then                                        '��Ȩ��ʹ�òŸ��ݲ�������
        mblnUseQueue = GetDeptPara(mlngCur����ID, "�����Ŷӽк�", 0) = 1 '          'Ĭ�ϲ������Ŷӽк�
    End If
    mlngRefreshInterval = GetDeptPara(mlngCur����ID, "�Զ�ˢ�¼��", 0) = 1 '      '�Զ�ˢ�¼��,Ĭ�ϲ��Զ�ˢ��
    If mlngRefreshInterval > 0 Then
        If mlngRefreshInterval > 65 Then mlngRefreshInterval = 65
        TimerRefresh.Interval = mlngRefreshInterval * 1000
        TimerRefresh.Enabled = True
    End If

    SQLCondition.��ʼʱ�� = CDate(Format(zlDatabase.Currentdate - mBeforeDays, "yyyy-mm-dd 00:00"))
    mblnMoved = MovedByDate(SQLCondition.��ʼʱ��)
    
    '��ʼ�����������б�
    Dim iCount As Integer, rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    iCount = 1
    gstrSQL = "Select ִ�м�,����豸 From ҽ��ִ�з��� where ����id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡִ�м�����", mlngCur����ID)
    If rsTemp.EOF <> True Then
        ReDim mAstr��������(rsTemp.RecordCount) As String
    While rsTemp.EOF = False
        mAstr��������(iCount) = Split(mstrCur����, "-")(1) & Nvl(rsTemp!ִ�м�)
        iCount = iCount + 1
        rsTemp.MoveNext
    Wend
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub Menu_Manage_PACS����()
    Dim i As Integer, lngAdviceID As Long, lngSendNO As Long
    
    With vsList
        lngAdviceID = .TextMatrix(.Row, Colҽ��ID)
        lngSendNO = .TextMatrix(.Row, Col���ͺ�)
    End With
    
    If Not mfrmPacsReportDock Is Nothing Then
        '���жϵ�ǰ�����Ƿ�����Ҫ�򿪵Ĵ��壬������ǣ�����Ҵ�������
        If lngAdviceID = mfrmPacsReportDock.mlngAdviceID Then
            '��ǰmfrmPacsReportDockָ��Ĵ��壬������Ҫ�򿪵Ĵ���
            mfrmPacsReportDock.WindowState = 0  'normal
            mfrmPacsReportDock.ZOrder
            Exit Sub
        End If
    End If
    
    '���Ҵ�������,�ҵ���Ҫ�򿪵Ĵ��壬��ͨ��Zorder�Ѵ�����ʾ����ǰ��
    If SafeArrayGetDim(mobjPacsReportArry) <> 0 Then
        For i = 1 To UBound(mobjPacsReportArry)
            If lngAdviceID = mobjPacsReportArry(i).mlngAdviceID Then
                Set mfrmPacsReportDock = mobjPacsReportArry(i)
                mfrmPacsReportDock.WindowState = 0  'normal
                mfrmPacsReportDock.ZOrder
                Exit Sub
            End If
        Next i
    End If
    
    'û���ҵ���Ҫ�򿪵Ĵ��壬�Ҵ��´���,����¼��ǰ����
    Set mfrmPacsReportDock = New frmReport
    mfrmPacsReportDock.zlEditReport lngAdviceID, lngSendNO, mlngCur����ID, Me, mstrPrivs, mlngModul, vsList.TextMatrix(vsList.Row, Colת��) = 1
    
    If SafeArrayGetDim(mobjPacsReportArry) = 0 Then
        ReDim mobjPacsReportArry(1) As frmReport
    Else
        ReDim Preserve mobjPacsReportArry(UBound(mobjPacsReportArry) + 1) As frmReport
    End If
    Set mobjPacsReportArry(UBound(mobjPacsReportArry)) = mfrmPacsReportDock
End Sub

Private Sub cmdInfo_Click()
    On Error GoTo ErrHandle
    frmDegreeCard.ShowMe Val(vsList.TextMatrix(vsList.Row, Col����ID)), Val(vsList.TextMatrix(vsList.Row, Col��ҳID))
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
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
    mlngCur����ID = 0
    mstrCur���� = ""
    mstrCanUse���� = ""
    mblnInitOk = False  '��ʼ����,��ʼ�����֮ǰ���������ݵ���ȡ
    Call InitLocalPars '����ע������
    If Not InitDepts Then Unload Me: Exit Sub '��ʼ��ҽ������
    Call InitMvar '��ʼ��ģ�鼶����
    '��ʼ�Ӵ���
    Set mfrmPACSImg = New frmPACSImg
    Set mfrmPacsReport = New frmReport  'PACS����
    Set mobjReport = New zlRichEPR.cDockReport
    Set mobjExpense = New zlCISKernel.clsDockExpense
    Set mobjInAdvice = New zlCISKernel.clsDockInAdvices
    Set mobjOutAdvice = New zlCISKernel.clsDockOutAdvices
    Set mobjInEPRs = New zlRichEPR.cDockInEPRs
    Set mobjOutEPRs = New zlRichEPR.cDockOutEPRs
    Set mobjPacsCore = New zl9PacsCore.clsViewer
    Set mobjQueue = New zl9QueueManage.clsQueueManage      '�Ŷӽк�
    
    Call InitFilterCmd
    Call InitCommandBars
    Call InitSubForm
    Call InitFaceScheme
    Call InitList

    Set mfrmPACSImg.pobjPacsCore = mobjPacsCore
    'ȥ��PACS���洰��Ŀ��ƿ�
    FormSetCaption mfrmPacsReport, False, False
    mblnInitOk = True '��ʼ�����
    Call RefreshList
    
    Call RestoreWinState(Me, App.ProductName)
    
    ClearCacheFolder App.Path & "\TmpImage\"    '����ʱĿ¼���ˣ�����ո�Ŀ¼
    
    
    Me.stbThis.Panels(3).Text = "����ҽ����" & UserInfo.����
    ReDim mobjPacsReportArry(0) As frmReport
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���ﲡ��", IIf(mblncmd����, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "סԺ����", IIf(mblncmdסԺ, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���ﲡ��", IIf(mblncmd����, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��첡��", IIf(mblncmd���, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�����ѽ�", IIf(mblncmd�ѽ�, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����δ��", IIf(mblncmdδ��, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�Ǽǲ���", IIf(mblncmd�Ǽ�, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��������", IIf(mblncmd����, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���没��", IIf(mblncmd����, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��˲���", IIf(mblncmd���, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��ɲ���", IIf(mblncmd���, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��λ��ʽ", mstrCurFindtype
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����סԺ", IIf(mblncmd����, 1, 0)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
    Call SaveWinState(Me, App.ProductName)
    
    '�ж�Ƕ��ʽ����༭���еı����Ƿ�û�б���
    If mblnPacsReport = True Then    'ʹ��PACS����༭��
        Call mfrmPacsReport.PromptModify
    End If
    
    Unload mfrmPACSImg
    Unload mfrmPacsReport
    Unload mobjReport.zlGetForm
    Unload mobjExpense.zlGetForm
    Unload mobjInAdvice.zlGetForm
    Unload mobjOutAdvice.zlGetForm
    Unload mobjInEPRs.zlGetForm
    Unload mobjOutEPRs.zlGetForm
    If Not mobjPacsCore Is Nothing Then mobjPacsCore.Closefrom
    
    Set mobjIDCard = Nothing
    Set mfrmPacsReport = Nothing
    Set mobjReport = Nothing
    Set mobjExpense = Nothing
    Set mobjInAdvice = Nothing
    Set mobjOutAdvice = Nothing
    Set mobjInEPRs = Nothing
    Set mobjOutEPRs = Nothing
    Set mobjPacsCore = Nothing
    
     On Error Resume Next
     
    '�ر��ŶӽкŴ���
    Unload mobjQueue.zlGetForm
    Set mobjQueue = Nothing
    
    '�������ά�ؽ����ر���ά�ؽ��Ĵ���
    If mblnUse3D = True Then
        Call sub3DProcess("EXIT")
    End If
End Sub

Private Sub InitLocalPars()
'��ʼ����ʱ���ز������Ը������ã�ע������Ϊ��,������أ��������õȵ���
    On Error GoTo err
    mblncmd���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���ﲡ��", 1))
    mblncmdסԺ = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "סԺ����", 1))
    mblncmd���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���ﲡ��", 1))
    mblncmd��� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��첡��", 1))
    mblncmd�ѽ� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�����ѽ�", 0))
    mblncmdδ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����δ��", 0))
    mblncmd�Ǽ� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�Ǽǲ���", 1))
    mblncmd���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��������", 1))
    mblncmd���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���没��", 1))
    mblncmd��� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��˲���", 1))
    mblncmd��� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��ɲ���", 1))
    mstrCurFindtype = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��λ��ʽ", "����")
    mblncmd���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����סԺ", "0"))
    
    mstrFirstTab = zlDatabase.GetPara("������ҳ", glngSys, mlngModul, "") 'Ϊ�ձ�ʾ��ʹ�ö��ƹ�����ҳ����
    mblnֱ�Ӽ�� = (Val(zlDatabase.GetPara("�Ǽ�ֱ�Ӽ��", glngSys, mlngModul, 0)) = 1)
    mblnOpenReport = (Val(zlDatabase.GetPara("��ʼ����Զ��򿪱���", glngSys, mlngModul, 0)) = 1)
    mblnShowImgAtReport = (Val(zlDatabase.GetPara("����ʱ��Ƭ", glngSys, mlngModul, 0)) = 1)
    mblnNoShowCancel = (Val(zlDatabase.GetPara("����ʾ��ȡ���ĵǼ�", glngSys, mlngModul, 0)) = 1)
    mblnPatTrack = (Val(zlDatabase.GetPara("���˸���", glngSys, mlngModul, 0)) = 1)
    mstrRoom = zlDatabase.GetPara("ִ�м䷶Χ", glngSys, mlngModul, "")
    If mstrRoom <> "" Then mstrRoom = "'," & Replace(mstrRoom, "|", ",") & ",'"
    
    '��ȡ��ά�ؽ�����
    mblnUse3D = Val(zlDatabase.GetPara("������ά�ؽ�", glngSys, mlngModul, 0))
    mstr3DExeDir = zlDatabase.GetPara("3D����·��", glngSys, mlngModul, "")
    mstr3DPara = zlDatabase.GetPara("3D����", glngSys, mlngModul, "")
    mstr3DFunctions = zlDatabase.GetPara("3D����", glngSys, mlngModul, "")

    With SQLCondition '------------------------ '����������ʼ
        .ʱ������ = 1                           'ʱ���ѯ��ʽ 1=�����ʱ�䡢2=������ʱ��
        .���ݺ� = ""
        .����� = 0
        .סԺ�� = 0
        .���￨ = ""
        .���� = ""
        .���� = 0
        .���֤ = ""
        .IC�� = ""
        .���˿��� = 0
        .�걾��λ = ""
        .���ҽ�� = ""
        .���ҽ�� = ""
        .������� = ""
        .������� = False
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

Private Function InitDepts() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str����IDs As String, str��Դ As String
    
    On Error GoTo errH
    
    str��Դ = "1,2,3"
    If InStr(mstrPrivs, "���п���") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where B.����ID = A.ID " & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " And instr([1],','||B.�������||',')> 0 And B.�������� IN('���')" & _
            " Order by A.����"
    Else
        strSQL = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B,������Ա C " & _
            " Where B.����ID = A.ID And A.ID=C.����ID And C.��ԱID=" & UserInfo.ID & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " And instr([1],','||B.�������||',')>0  And B.�������� IN('���')" & _
            " Order by A.����"
    End If
   

    
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, "," & str��Դ & ",")
    
    If rsTmp.EOF Then
        MsgBoxD Me, "û�з���ҽ��������Ϣ,���ȵ����Ź��������á�", vbInformation, gstrSysName
        Exit Function
    Else
        str����IDs = GetUser����IDs
        Do Until rsTmp.EOF
            mstrCanUse���� = mstrCanUse���� & "|" & rsTmp!ID & "_" & rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ID = UserInfo.����ID Then mlngCur����ID = rsTmp!ID: mstrCur���� = rsTmp!���� & "-" & rsTmp!���� '��ȡĬ�Ͽ���
            If InStr("," & str����IDs & ",", "," & rsTmp!ID & ",") > 0 And mlngCur����ID = 0 Then mlngCur����ID = rsTmp!ID: mstrCur���� = rsTmp!���� & "-" & rsTmp!���� 'û��Ĭ�Ͽ���ȡ��һ������������
            rsTmp.MoveNext
        Loop
        mstrCanUse���� = Mid(mstrCanUse����, 2)
        If InStr(mstrPrivs, "���п���") > 0 And mlngCur����ID = 0 Then
            mlngCur����ID = Split(Split(mstrCanUse����, "|")(0), "_")(0)
            mstrCur���� = Split(Split(mstrCanUse����, "|")(0), "_")(1)
        End If
        
        If mlngCur����ID = 0 And InStr(mstrPrivs, "���п���") <= 0 Then 'û�����п��Ҳ���Ȩ��,���Ҳ����߿��Ҳ����ڼ�������
            MsgBoxD Me, "û�з�������������,����ʹ��ҽ������վ��", vbInformation, gstrSysName
            Exit Function
        End If
        InitDepts = True
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

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
    Set objPopbar = objBar.Controls.Add(xtpControlPopup, ID_���ҷ�ʽ, "���ҷ�ʽ")
        objPopbar.ID = ID_���ҷ�ʽ
        objPopbar.Flags = xtpFlagRightAlign
        
    Set objCusControl = objBar.Controls.Add(xtpControlCustom, ID_����ֵ, "����ֵ")
        objCusControl.Handle = txtFilter.Hwnd
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
        .Add FCONTROL, vbKey4, ID_����
        .Add FCONTROL, vbKey5, ID_�Ǽ�
        .Add FCONTROL, vbKey6, ID_����
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
    Me.cbrMain.Icons = frmPubIcons.imgPublic.Icons
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        '.SetIconSize False, 16, 16
    End With
    Me.cbrMain.EnableCustomization False
    

'�˵�����
'Begin------------------------�ļ��˵�--------------------------------------Ĭ�Ͽɼ�
    Me.cbrMain.ActiveMenuBar.Title = "�˵�"
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)"): cbrControl.IconId = 181
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "����Ԥ��(&V)"): cbrControl.IconId = 102
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "�����ӡ(&P)"): cbrControl.IconId = 103
        Set cbrControl = .Add(xtpControlButton, conMenu_File_BatPrint, "������ӡ(&B)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�嵥��ӡ(&L)"): cbrControl.BeginGroup = True: cbrControl.IconId = 103
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&O)"):: cbrControl.IconId = 181
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_DevSet, "Ӱ���豸����(&D)"):: cbrControl.IconId = 181
        Set cbrControl = .Add(xtpControlButton, conMenu_File_SendImg, "����ͼ��(&T)"): cbrControl.IconId = 3061
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"):: cbrControl.IconId = 191: cbrControl.BeginGroup = True
    End With


'Begin----------------------���˵�--------------------------------------Ĭ�Ͽɼ�
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "���(&S)", -1, False)
    cbrMenuBar.ID = conMenu_ManagePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_Manage_RequestPrint, "��ӡ���뵥��(&J)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Regist, "���Ǽ�(&I)"): cbrControl.IconId = 211: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_CopyCheck, "���ƵǼ�(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Redo, "ȡ���Ǽ�(&R)"): cbrControl.IconId = 742
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ReGet, "�ٻ�ȡ��(&G)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ThingModi, "�޸���Ϣ(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Receive, "��鱨��(&L)"):  cbrControl.BeginGroup = True: cbrControl.IconId = 744
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Logout, "ȡ������(&D)"): cbrControl.IconId = 743
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Look, "Ӱ���Ƭ(&S)"): cbrControl.IconId = 8111:  cbrControl.BeginGroup = True
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
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Delete, "Ӱ��ɾ��(&K)"): cbrControl.IconId = 8113
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Query, "Q/R��ȡͼ��(&Q)"): cbrControl.IconId = 8111
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Transfer, "����Ӱ��(&C)"):  cbrControl.BeginGroup = True: cbrControl.IconId = 505: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Cancel, "ȡ������(&B)"): cbrControl.IconId = 506
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Review, "���(&R)"):  cbrControl.BeginGroup = True: cbrControl.IconId = 232
        Set cbrControl = .Add(xtpControlPopup, conMenu_Manage_Result, "�����(&X)"): cbrControl.ID = conMenu_Manage_Result
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Negative, "����(&X)"): cbrPopControl.IconId = 3506
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Positive, "����(&Y)"): cbrPopControl.IconId = 3507
        Set cbrControl = .Add(xtpControlPopup, conMenu_Manage_Quality, "Ӱ������(&Y)"): cbrControl.ID = conMenu_Manage_Quality: cbrControl.IconId = 3061
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_First, "�׼�(&J)"): cbrPopControl.IconId = 3587
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Second, "�Ҽ�(&Y)"): cbrPopControl.IconId = 3010
        Set cbrControl = .Add(xtpControlPopup, conMenu_Manage_GChannel, "��ɫͨ��(&G)"): cbrControl.ID = conMenu_Manage_GChannel
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_GChannelOk, "���(&J)")
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_GChannelCancel, "ȡ��(&Y)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Finish, "�ޱ������(&F)"): cbrControl.IconId = 216: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ClearUp, "�ޱ������(&U)"):  cbrControl.IconId = 3012
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Complete, "������(&E)"): cbrControl.IconId = 225
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Undone, "ȡ�����(&U)"): cbrControl.IconId = 219
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ChangeDevice, "�����豸"): cbrControl.IconId = 3203
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
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_View_Filter * 10#, "������"): cbrControl.ID = conMenu_View_Filter * 10#
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "���ٹ���(&K)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&F)")
    End With


'Begin----------------------�����˵�--------------------------------------Ĭ�Ͽɼ�
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������", -1, False)
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "WEB�ϵ�����(&E)")
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
        .Add 0, VK_F12, conMenu_File_Parameter      '��������--------------F12
        
        .Add 0, VK_F2, conMenu_Manage_Regist       '�Ǽ�-----------------F2
        .Add 0, VK_F7, conMenu_Manage_CopyCheck    '���ƵǼ�-------------F7
        .Add 0, VK_F4, conMenu_Manage_Receive       '����-----------------F4
        .Add 0, VK_F9, conMenu_Manage_ClearUp       '���ر���------------F9
        .Add 0, VK_F6, conMenu_Manage_Complete         '��˱���----------F6
        
        
        .Add 0, VK_F1, conMenu_Help_Help              '����-------------F1
        .Add 0, VK_F5, conMenu_View_Refresh           'ˢ��-------------F5
        .Add 0, VK_F3, conMenu_View_Filter            '����-------------F3
    End With
    
'---------------------�������Ͻǵ�ǰ����----------------------------------
        Set cbrControl = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_Filter * 10#, "������")
            cbrControl.ID = conMenu_View_Filter * 10#: cbrControl.Flags = xtpFlagRightAlign: cbrControl.Category = "Main"
    
    
'---------------------����������------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��"): cbrControl.IconId = 102: cbrControl.ToolTipText = "����Ԥ��"
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ"): cbrControl.IconId = 103: cbrControl.ToolTipText = "�����ӡ"
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Regist, "�Ǽ�"): cbrControl.BeginGroup = True: cbrControl.IconId = 211
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Receive, "����"): cbrControl.IconId = 744
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Logout, "ȡ��"): cbrControl.IconId = 743: cbrControl.ToolTipText = "ȡ������"
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
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Review, "���"):  cbrControl.BeginGroup = True: cbrControl.IconId = 232
        Set cbrControl = .Add(xtpControlPopup, conMenu_Manage_Result, "���"): cbrControl.ID = conMenu_Manage_Result: cbrControl.IconId = 3506: cbrControl.ToolTipText = "�����������"
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Negative, "����(&X)"): cbrPopControl.IconId = 3506
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Positive, "����(&Y)"): cbrPopControl.IconId = 3507
        Set cbrControl = .Add(xtpControlPopup, conMenu_Manage_Quality, "����"): cbrControl.ID = conMenu_Manage_Quality: cbrControl.IconId = 3061: cbrControl.ToolTipText = "Ӱ������"
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_First, "�׼�(&J)"): cbrPopControl.IconId = 3587
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Second, "�Ҽ�(&Y)"): cbrPopControl.IconId = 3010
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Complete, "���"): cbrControl.IconId = 225: cbrControl.ToolTipText = "����������"
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
        .Icons = frmPubIcons.imgPublic.Icons
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .PaintManager.ColorSet.ButtonSelected = &HFFC0C0
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True
        .RemoveAll
        
        .InsertItem 0, "Ӱ���¼", mfrmPACSImg.Hwnd, conMenu_Img_Look
            .Item(TabWindow.ItemCount - 1).Tag = "Ӱ��ͼ��"
        
        If GetInsidePrivs(p���Ʊ������, True) <> "" Then
            If mblnPacsReport = True Then
                .InsertItem 1, "Ӱ�񱨸�", mfrmPacsReport.Hwnd, conMenu_Edit_Compend
            Else
                .InsertItem 1, "Ӱ�񱨸�", mobjReport.zlGetForm.Hwnd, conMenu_Edit_Compend
            End If
            .Item(TabWindow.ItemCount - 1).Tag = "������д"
        End If
        
        If GetInsidePrivs(pҽ�����ѹ���, True) <> "" Then
            .InsertItem 2, "���ü�¼", mobjExpense.zlGetForm.Hwnd, conMenu_Manage_Request
            .Item(TabWindow.ItemCount - 1).Tag = "�������"
        End If
        
        If GetInsidePrivs(pסԺҽ���´�, True) <> "" Then
            .InsertItem 3, "ҽ����¼", mobjInAdvice.zlGetForm.Hwnd, conMenu_Edit_NewItem
            .Item(TabWindow.ItemCount - 1).Tag = "סԺҽ��"
        End If
        
        If GetInsidePrivs(p����ҽ���´�, True) <> "" Then
            .InsertItem 4, "ҽ����¼", mobjOutAdvice.zlGetForm.Hwnd, conMenu_Edit_NewItem
            .Item(TabWindow.ItemCount - 1).Tag = "����ҽ��": .Item(TabWindow.ItemCount - 1).Visible = False
        End If
        
        If GetInsidePrivs(pסԺ��������, True) <> "" Then
            .InsertItem 5, "������¼", mobjInEPRs.zlGetForm.Hwnd, conMenu_Edit_Archive
            .Item(TabWindow.ItemCount - 1).Tag = "סԺ����"
        End If
        
        If GetInsidePrivs(p���ﲡ������, True) <> "" Then
            .InsertItem 6, "������¼", mobjOutEPRs.zlGetForm.Hwnd, conMenu_Edit_Archive
            .Item(TabWindow.ItemCount - 1).Tag = "���ﲡ��": .Item(TabWindow.ItemCount - 1).Visible = False
        End If
        
        '����Ŷӽк�ҳ��
        If mblnUseQueue = True Then
            .InsertItem 7, "�Ŷӽк�", mobjQueue.zlGetForm.Hwnd, conMenu_File_PrintSingleBill
            .Item(TabWindow.ItemCount - 1).Tag = "�Ŷӽк�"
        End If
        
        If mstrFirstTab <> "" Then
            For i = 0 To .ItemCount - 1
                If InStr(.Item(i).Tag, mstrFirstTab) > 0 And .Item(i).Visible Then
                    .Item(i).Selected = True: Exit For
                End If
            Next
        End If
    End With

End Sub


Private Sub InitList()
    With vsList
        .Clear
        .FixedRows = 1
        .Rows = 2
        .Cols = 45
        
        .ColWidth(Col����) = 300: .ColWidth(Col��Դ) = 400: .ColWidth(Col����) = 300: .ColWidth(Col����) = 300: .ColWidth(Col����) = 1200
        .ColWidth(Col����) = 1400: .ColWidth(Col������) = 800: .ColWidth(Col�Ա�) = 450: .ColWidth(Col����) = 450: .ColWidth(Col��ʶ��) = 1400: .ColWidth(Colҽ������) = 2400
        .ColWidth(Col��λ����) = 1400: .ColWidth(Colִ�м�) = 600: .ColWidth(Col���ʱ��) = 1700: .ColWidth(Col����ʱ��) = 1700: .ColWidth(Col����ҽ��) = 800
        .ColWidth(Col���) = 450: .ColWidth(Col����) = 450: .ColWidth(ColӤ��) = 450: .ColWidth(Col�Ǽ���) = 800: .ColWidth(Col������) = 800
        .ColWidth(Col�����) = 800: .ColWidth(Col��ӡ��Ƭ) = 800: .ColWidth(Col�������) = 800: .ColWidth(Col��ɫͨ��) = 0: .ColWidth(Col�����ӡ) = 800
        .ColWidth(Col������) = 800: .ColWidth(Col������) = 800: .ColWidth(Col��鼼ʦ) = 800: .ColWidth(Col��ͼʱ��) = 1700: .ColWidth(Col�������) = 2400
        
        .ColWidth(ColӰ�����) = 0: .ColWidth(Col����ID) = 0: .ColWidth(Col��ҳID) = 0: .ColWidth(Col�Һŵ�) = 0: .ColWidth(Colҽ��ID) = 1200: .ColWidth(Col���ͺ�) = 0
        .ColWidth(Col���˿���ID) = 0: .ColWidth(Col���UID) = 0: .ColWidth(Col���״̬) = 0: .ColWidth(ColNO) = 0: .ColWidth(Col��¼����) = 0: .ColWidth(Colת��) = 0
        .ColWidth(Col����) = 0: .ColWidth(Col��ǰ����ID) = 0
        
        Set .Cell(flexcpPicture, 0, Col����) = imgList.ListImages("����").Picture
        Set .Cell(flexcpPicture, 0, Col��Դ) = imgList.ListImages("סԺ").Picture
        Set .Cell(flexcpPicture, 0, Col����) = imgList.ListImages("����").Picture
        
        .TextMatrix(0, Col����) = "��": .TextMatrix(0, Col����) = "����": .TextMatrix(0, Col����) = "����": .TextMatrix(0, Col������) = "������"
        .TextMatrix(0, Col�Ա�) = "�Ա�": .TextMatrix(0, Col����) = "����": .TextMatrix(0, Col��ʶ��) = "��ʶ��": .TextMatrix(0, Colҽ������) = "ҽ������": .TextMatrix(0, Col��λ����) = "��λ����"
        .TextMatrix(0, Colִ�м�) = "ִ�м�": .TextMatrix(0, Col���ʱ��) = "���ʱ��": .TextMatrix(0, Col����ʱ��) = "����ʱ��": .TextMatrix(0, Col����ҽ��) = "����ҽ��"
        .TextMatrix(0, Col���) = "���": .TextMatrix(0, Col����) = "����": .TextMatrix(0, ColӤ��) = "Ӥ��": .TextMatrix(0, Col�Ǽ���) = "�Ǽ���"
        .TextMatrix(0, Col������) = "������": .TextMatrix(0, Col�����) = "�����": .TextMatrix(0, Col��ӡ��Ƭ) = "��ӡ��Ƭ": .TextMatrix(0, Col�������) = "�������"
        .TextMatrix(0, Col��ɫͨ��) = "��ɫͨ��": .TextMatrix(0, Col�����ӡ) = "�����ӡ": .TextMatrix(0, Col������) = "������": .TextMatrix(0, Col������) = "������"
        .TextMatrix(0, Col��鼼ʦ) = "��鼼ʦ": .TextMatrix(0, Col��ͼʱ��) = "��ͼʱ��": .TextMatrix(0, Col�������) = "�������"
        
        .TextMatrix(0, ColӰ�����) = "Ӱ�����": .TextMatrix(0, Col����ID) = "����ID": .TextMatrix(0, Col��ҳID) = "��ҳID": .TextMatrix(0, Col�Һŵ�) = "�Һŵ�"
        .TextMatrix(0, Col���˿���ID) = "���˿���ID": .TextMatrix(0, Colҽ��ID) = "ҽ��ID": .TextMatrix(0, Col���ͺ�) = "���ͺ�": .TextMatrix(0, Col���UID) = "���UID"
        .TextMatrix(0, Col���״̬) = "���״̬": .TextMatrix(0, ColNO) = "NO": .TextMatrix(0, Col��¼����) = "��¼����": .TextMatrix(0, Colת��) = "ת��"
        .TextMatrix(0, Col����) = "����": .TextMatrix(0, Col��ǰ����ID) = "��ǰ����ID"

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


Private Sub mfrmPacsReport_AfterClosed(ByVal lngOrderID As Long)
    Call EditorClosed(lngOrderID)
End Sub

Private Sub mfrmPacsReport_AfterDeleted(ByVal lngOrderID As Long)
    AfterDeleted lngOrderID
End Sub

Private Sub mfrmPacsReport_AfterPrinted(ByVal lngOrderID As Long)
    Call AfterPrinted(lngOrderID)
End Sub

Private Sub mfrmPacsReport_AfterSaved(ByVal lngOrderID As Long)
    Call AfterReportSaved(lngOrderID)
End Sub

Private Sub mfrmPacsReport_BeforeEdit()
Dim lngOrderID As Long

    On Error GoTo ErrHandle
    lngOrderID = vsList.TextMatrix(vsList.Row, Colҽ��ID)
    If CheckConcurrentReport(lngOrderID) Then '����Ƿ��������ڲ�������
        Call UpdateReporter(lngOrderID, UserInfo.����)
    Else
        Call mfrmPacsReport.PromptModify(True)
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mfrmPacsReportDock_AfterOpen()
    Call AfterReportOpen
End Sub

Private Sub mfrmPacsReportDock_AfterPrinted(ByVal lngOrderID As Long)
    Call AfterPrinted(lngOrderID)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    If txtFilter.Text = "" And Me.ActiveControl Is txtFilter Then
        IDKind.IDKind = IDKinds.C2���֤��
        mstrCurFindtype = "���֤"
        txtFilter = strID
        Call txtFilter_KeyDown(vbKeyReturn, 0)
    End If
End Sub

Private Sub mobjInAdvice_ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean)
Dim cbrControl As CommandBarControl, lngҽ��ID As Long, rsTemp As ADODB.Recordset
    gstrSQL = "select ҽ��ID FROM ����ҽ������ where ����ID=[1]"
    Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡҽ��ID", CLng(����ID))
    If rsTemp.EOF Then Exit Sub
    
    lngҽ��ID = Nvl(rsTemp!ҽ��ID, 0)
    mobjReport.zlRefresh lngҽ��ID, mlngCur����ID, False '�Բ���Edit��ʽˢ�¶���
    
    Set cbrControl = cbrMain(2).Controls.Find(, conMenu_Help_Help)
    cbrControl.ID = conMenu_File_Open
    mobjReport.zlExecuteCommandBars cbrControl '���ò��ı���
    cbrControl.ID = conMenu_Help_Help
End Sub

Private Sub mobjInAdvice_ViewPACSImage(ByVal ҽ��ID As Long)
    '����100��ͼ������У�Ĭ��ÿ��5�Ŵ�һ��
    Call OpenViewer(mobjPacsCore, ҽ��ID, False, Me, , , mblnLocalizerBackward, 5)
End Sub

Private Sub mobjOutAdvice_ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean)
Dim cbrControl As CommandBarControl, lngҽ��ID As Long, rsTemp As ADODB.Recordset
    gstrSQL = "select ҽ��ID FROM ����ҽ������ where ����ID=[1]"
    Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡҽ��ID", CLng(����ID))
    If rsTemp.EOF Then Exit Sub
    
    lngҽ��ID = Nvl(rsTemp!ҽ��ID, 0)
    mobjReport.zlRefresh lngҽ��ID, mlngCur����ID, False '�Բ���Edit��ʽˢ�¶���
    
    Set cbrControl = cbrMain(2).Controls.Find(, conMenu_Help_Help)
    cbrControl.ID = conMenu_File_Open
    mobjReport.zlExecuteCommandBars cbrControl '���ò��ı���
    cbrControl.ID = conMenu_Help_Help
End Sub

Private Sub mobjOutAdvice_ViewPACSImage(ByVal ҽ��ID As Long)
    '����100��ͼ������У�Ĭ��ÿ��5�Ŵ�һ��
    Call OpenViewer(mobjPacsCore, ҽ��ID, False, Me, , , mblnLocalizerBackward, 5)
End Sub

Private Sub mobjPacsCore_AfterSaveReportImage(strStudyUID As String)
    If mblnPacsReport = True Then
        mfrmPacsReport.RefPacsPic 'ˢ��ͼƬ
        If Not mfrmPacsReportDock Is Nothing Then
            mfrmPacsReportDock.RefPacsPic 'ˢ��ͼƬ
        End If
    Else
        mobjReport.RefPacsPic 'ˢ��ͼƬ
    End If
End Sub

Private Sub mobjReport_AfterClosed(ByVal lngOrderID As Long)
    Call EditorClosed(lngOrderID)
End Sub
Public Sub EditorClosed(ByVal lngOrderID As Long)
    Dim i As Integer
    Dim j As Integer
    
    Call UpdateReporter(lngOrderID, "")
    '����PACS����༭���Ĵ�������
    On Error Resume Next
    If mblnPacsReport = True Then
        '���Ҵ������飬�ҵ���Ӧ�Ĵ��ڲ�ɾ��
        If SafeArrayGetDim(mobjPacsReportArry) <> 0 Then
            For i = 1 To UBound(mobjPacsReportArry)
                If mobjPacsReportArry(i).mlngAdviceID = lngOrderID Then
                    '��������ɾ��
                    For j = i To UBound(mobjPacsReportArry)
                        Set mobjPacsReportArry(j) = mobjPacsReportArry(j + 1)
                    Next j
                    ReDim Preserve mobjPacsReportArry(UBound(mobjPacsReportArry) - 1) As frmReport
                    Exit For
                End If
            Next i
        End If
        
        If Not mfrmPacsReportDock Is Nothing Then
            If lngOrderID = mfrmPacsReportDock.mlngAdviceID Then
                '�رյ�ǰ���洰�ڣ�����ǰ�������óɿ�
                Set mfrmPacsReportDock = Nothing
            End If
        End If
    End If
End Sub

Private Sub mobjReport_AfterDeleted(ByVal lngOrderID As Long)
    AfterDeleted lngOrderID
End Sub

Private Sub AfterDeleted(ByVal lngOrderID As Long)
    On Error GoTo ErrHandle
    gstrSQL = "ZL_Ӱ�񱨸���_Clear(" & lngOrderID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, "��ձ��"
    Call RefreshList
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub mobjReport_AfterOpen(ByVal intEditType As zlRichEPR.EditTypeEnum)
    Call AfterReportOpen
End Sub

Private Sub AfterReportOpen()
Dim lngOrderID As Long
    On Error GoTo ErrHandle
    lngOrderID = vsList.TextMatrix(vsList.Row, Colҽ��ID)
    
    Call UpdateReporter(lngOrderID, UserInfo.����)
    
    If mblnShowImgAtReport And vsList.TextMatrix(vsList.Row, Col���UID) <> "" Then
        Dim intImageInverval As Integer
        
        intImageInverval = Val(mfrmPACSImg.cbrMain.FindControl(, conMenu_Manage_ImageInterval, , True).Text)
        Call OpenViewer(mobjPacsCore, lngOrderID, False, Me, , , mblnLocalizerBackward, intImageInverval)
    End If
    Exit Sub
    
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mobjReport_AfterPrinted(ByVal lngOrderID As Long)
    Call AfterPrinted(lngOrderID)
End Sub
Public Sub AfterPrinted(lngOrderID As Long)
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    gstrSQL = "ZL_Ӱ�񱨸��ӡ_Update(" & lngOrderID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, "���´�ӡ���"
    
    If Not mblnIgnoreResult And mintResultInput = 2 Then
        strSQL = "Select �������  From  ����ҽ������ Where ҽ��id= [1]"
        Set rsTemp = OpenSQLRecord(strSQL, "��ȡ�������", lngOrderID)
        
        If IsNull(rsTemp!�������) Then  '�ڱ���ʱ��ʾ���������
            Call PromptResult(lngOrderID, mlngModul, Me)
        End If
    End If
    
    If mblnPrintCommit = True Then
        Call Menu_Manage_����������(lngOrderID, False)
    End If
    
    Call RefreshList
    Exit Sub
    
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub mobjReport_AfterSaved(ByVal lngOrderID As Long)
    Call AfterReportSaved(lngOrderID)
End Sub

Public Sub AfterReportSaved(lngOrderID As Long)
    Dim rsTemp As ADODB.Recordset, i As Integer, intState As Integer, lngSendId As Long
    If mblnPacsReport = True Then
'        mfrmPacsReport.zlRefresh 0, 0, 0
    Else
        mobjReport.zlRefresh 0, mlngCur����ID, False
    End If

    gstrSQL = "Select Distinct A.ҽ��id, B.ID,B.������,B.������,B.ǩ������, B.���ʱ��, B.���汾, C.���ͺ�,C.�������, D.���UID " & vbNewLine & _
                "From ����ҽ������ A, ���Ӳ�����¼ B, ����ҽ������ C,Ӱ�����¼ D " & vbNewLine & _
                "Where A.ҽ��id =[1] And A.����id = B.ID And A.ҽ��id = C.ҽ��id AND D.ҽ��id = C.ҽ��id"
    Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡ�Ƿ�ǩ��", CLng(lngOrderID))
    If rsTemp.EOF Then Exit Sub
    lngSendId = rsTemp!���ͺ�
    
    If Nvl(rsTemp!���ʱ��, "") = "" And rsTemp!���汾 = 1 Then 'δǩ������ �����һ��ҽʦ��ǩ
        gstrSQL = "Zl_Ӱ����_State(" & lngOrderID & "," & lngSendId & "," & IIf(Nvl(rsTemp!���UID) = "", 2, 3) & ")"
        zlDatabase.ExecuteProcedure gstrSQL, "��Ϊ����ʱ"
        gstrSQL = "ZL_Ӱ�񱨸汣��_Update(" & lngOrderID & ",'" & Nvl(rsTemp!������, rsTemp!������) & "','')"
        zlDatabase.ExecuteProcedure gstrSQL, "���汨����"
        intState = IIf(Nvl(rsTemp!���UID) = "", 2, 3)
    Else
        If rsTemp!ǩ������ < 2 Then '���һ��ǩ��Ϊҽʦ,�п��ܵ���� 1-ҽʦ��N��ǩ�� 2-���μ������һ����ǩ 3-�޶�ģʽ�±���(ǩ������=0)
            gstrSQL = "Zl_Ӱ����_State(" & lngOrderID & "," & lngSendId & ",4)"
            zlDatabase.ExecuteProcedure gstrSQL, "��Ϊ����ʱ"
            
            intState = 4
        Else                        '���μ����ϼ���ǩ��
            gstrSQL = "Zl_Ӱ����_State(" & lngOrderID & "," & lngSendId & ",5)"
            zlDatabase.ExecuteProcedure gstrSQL, "��Ϊ���ʱ"

            intState = 5
            If mblnCompleteCommit Then
                intState = 6
                Call Menu_Manage_����������(lngOrderID, False)
            End If
        End If
        
        gstrSQL = "ZL_Ӱ�񱨸汣��_Update(" & lngOrderID & ",'" & IIf(rsTemp!ǩ������ = 1, Nvl(rsTemp!������), IIf(rsTemp!���汾 = 1, Nvl(rsTemp!������), "")) & "','" & IIf(rsTemp!ǩ������ = 1, "", Nvl(rsTemp!������)) & "')"
        zlDatabase.ExecuteProcedure gstrSQL, "���渴����" 'ǩ�����𣽣���ʾ��ҽ��ǩ��,�����ǵ�N�Σ���ʱ����������Ҫ���棬��������Ҫ���;������������˴��գ��������д�������������ֵ
    
        If Not mblnIgnoreResult And IsNull(rsTemp!�������) Then  '�ڱ���ʱ��ʾ���������
            If mblnReportWithResult Then '��Ӱ�����Ϊ����  -����ʾ�Զ����
                gstrSQL = "ZL_Ӱ����_���(" & lngOrderID & ",0)"
                zlDatabase.ExecuteProcedure gstrSQL, "���������"
            ElseIf mintResultInput = 1 Then
                Call PromptResult(lngOrderID, mlngModul, Me)
            End If
        End If
    End If

    Call StateCheck(intState)
End Sub

Private Sub StateCheck(ByVal intState As Integer)
Dim cbrControl As CommandBarControl
    Select Case intState '���ݲ�����״̬ȷ����״̬�����Ƿ�ѡ��
        Case 0, 1
            If Not mblncmd�Ǽ� Then Set cbrControl = Me.cbrdock.FindControl(, ID_�Ǽ�)
        Case 2, 3
            If Not mblncmd���� Then Set cbrControl = Me.cbrdock.FindControl(, ID_����)
        Case 4
            If Not mblncmd���� Then Set cbrControl = Me.cbrdock.FindControl(, ID_����)
        Case 5
            If Not mblncmd��� Then Set cbrControl = Me.cbrdock.FindControl(, ID_���)
        Case 6
            If Not mblncmd��� Then Set cbrControl = Me.cbrdock.FindControl(, ID_���)
    End Select
    If mblnPatTrack Then
        If Not cbrControl Is Nothing Then '����ѡ��,ѡ�д����б�ˢ��ͬʱʵ�ָ���
            cbrdock_Execute cbrControl
        Else
            Call RefreshList
        End If
    Else '������ֻˢ���б�
        Call RefreshList
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
        If .TextMatrix(.Row, Colת��) = 1 Then
            gstrSQL = Replace(gstrSQL, "����ҽ����¼", "H����ҽ����¼")
            gstrSQL = Replace(gstrSQL, "����ҽ������", "H����ҽ������")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(.TextMatrix(.Row, Colҽ��ID)), CLng(Decode(.TextMatrix(.Row, Col��Դ), "��", 1, "ס", 2, "��", 3, 4)))
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
        Call ReportOpen(gcnOracle, glngSys, objControl.Parameter, Me, "NO=" & vsList.TextMatrix(vsList.Row, ColNO), "����=" & vsList.TextMatrix(vsList.Row, Col��¼����), 1)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub RefreshList()
Dim i As Integer, lngcurҽ��ID As Long
    With vsList
        lngcurҽ��ID = Val(.TextMatrix(.Row, Colҽ��ID)) '��ǰ��ҽ��ID
        
        Call LoadPatiList
        If lngcurҽ��ID = 0 Then
            Call .Select(1, Col����)
            Exit Sub
        End If
        
        '�м�¼ʱҪ���¶�λ��֮ǰ��¼
        For i = 1 To .Rows - 1
            If lngcurҽ��ID = Val(.TextMatrix(i, Colҽ��ID)) Then
                Call .Select(i, Col����)
                Exit Sub
            End If
        Next
        'û�ܶ�λ֮ǰ�ļ�¼����λ����1��
        Call .Select(1, Col����)
    End With
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    fraRegist.Left = 0
    fraRegist.Top = -75
    fraInfo.Top = -75
    fraInfo.Left = fraRegist.Left + fraRegist.Width
    fraInfo.Width = picInfo.ScaleWidth - fraInfo.Left
    
    lblCash.Top = (picInfo.ScaleHeight - lblCash.Height) / 2 - fraInfo.Top
    lblCash.Left = fraInfo.Width - lblCash.Width - 100

    lbl������Ϣ.Width = lblCash.Left
    lbl�����Ϣ.Width = lblCash.Left
End Sub

Private Sub LoadPatiList()
'���ܣ���ȡ��ǰҽ�����ҵ�ִ��ҽ��(����)�嵥
Dim strSQL As String, strSQLBak As String, i As Long, rsList As ADODB.Recordset
Dim str��Դ As String
Dim strFilter As String

    If Not mblnInitOk Then Exit Sub      '��ʼ��δ���
    
    On Error GoTo ErrHandle
    With SQLCondition
        If Not (.����� = 0 And .סԺ�� = 0 And .���￨ = "" And .���� = "" And .���֤ = "" And .IC�� = "") Then '�������������ʹ��ʱ������
            If .����� <> 0 Then
                strFilter = " And C.�����=[1]"
            ElseIf .סԺ�� <> 0 Then
                strFilter = " And C.סԺ��=[2]"
            ElseIf .���￨ <> "" Then
                strFilter = " And C.���￨=[3]"
            ElseIf .���� <> "" Then
                strFilter = " And C.����=[4]"
            ElseIf .���֤ <> "" Then
                strFilter = " And C.���֤=[5]"
            ElseIf .IC�� <> "" Then
                strFilter = " And C.IC��=[6]"
            End If
        ElseIf .���ݺ� <> "" Then
            strFilter = " And A.NO=[7] "
        ElseIf .���� <> 0 Then
            strFilter = " And H.����=[8] "
        Else
            If .����ʱ�� <> CDate(0) Then
                strFilter = " And " & IIf(mblncmd�Ǽ�, "A.����ʱ��", IIf(.ʱ������ = 2, "A.����ʱ��", "A.�״�ʱ��")) & " Between [9] and [10] "
            Else 'ȱʡ��ѯ����
                strFilter = " And " & IIf(mblncmd�Ǽ�, "A.����ʱ��", IIf(.ʱ������ = 2, "A.����ʱ��", "A.�״�ʱ��")) & " Between [9] and Sysdate+1/(24*3600) "
            End If
            
            If .���˿��� <> 0 Then
                strFilter = strFilter & " And B.���˿���ID+0=[11] "
            End If
        
            If .�걾��λ <> "" Then
                strFilter = strFilter & " And instr(B.ҽ������,[12])>0"
            End If
            
            If .������� Then
                strFilter = strFilter & " And Nvl(A.�������, 0)=1"
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
            
            If .Ӱ����� <> "" Then
                strFilter = strFilter & " And H.Ӱ�����=[17] "
            End If
            
            If .��� <> "" Then
                strFilter = strFilter & " And  Instr(H.�������, [18]) > 0 "
            End If
            
            If .������� <> "" Then '-------------------------------------------------------------------------��Ҫ��
                strFilter = strFilter & " And (B.����ID,B.��ҳID) IN(Select Distinct A.����Id,A.��ҳID  " & _
                                                                        "From ���Ӳ�����¼ A,���Ӳ������� B " & _
                                                                        "Where A.����ʱ��>[1] AND A.Id=B.�ļ�ID  " & _
                                                                            "And B.��������=7 And instr(B.��������,'52;')>0 And instr(B.�����ı�,[19])>0)"
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
                strFilter = strFilter & " And (B.����ID,B.��ҳID) IN(Select Distinct a.����Id,a.��ҳID From ���Ӳ�����¼ a, ���Ӳ������� b,���Ӳ������� c " _
                    & " Where a.����ʱ�� > [9] And a.Id = b.�ļ�id And b.Id = C.��ID And b.�������� = 3 And c.�������� = 2 And c.��ֹ�� = 0 and " _
                    & strSubFilter & ")"
            End If
           
            If .������ <> "" Then
                If .������ = "ȫ��" Then
                
                ElseIf .������ = "�ѵǼ�" Then
                    strFilter = strFilter & " And (A.ִ�й��� =0 or A.ִ�й���=1 Or A.ִ�й��� Is Null) "
                ElseIf .������ = "�ѱ���" Then
                    strFilter = strFilter & " And (A.ִ�й��� = 2 and h.������ is null) "
                ElseIf .������ = "�Ѽ��" Then
                    strFilter = strFilter & " And (A.ִ�й��� = 3 and h.������ is null) "
                ElseIf .������ = "������" Then
                    strFilter = strFilter & " And (not h.������� is null) "
                ElseIf .������ = "������" Then
                    strFilter = strFilter & " And ((A.ִ�й��� =2 or A.ִ�й���=3) and not h.������ is null and h.������� is null) "
                ElseIf .������ = "�ѱ���" Then
                    strFilter = strFilter & " And (A.ִ�й���=4 and h.������ is null) "
                ElseIf .������ = "�����" Then
                    strFilter = strFilter & " And (A.ִ�й���=4 and not h.������ is null) "
                ElseIf .������ = "�����" Then
                    strFilter = strFilter & " And A.ִ�й���=5 "
                ElseIf .������ = "�����" Then
                    strFilter = strFilter & " And A.ִ�й���=6 "
                End If
            End If
        End If
        
        '�����˴��ڡ��͡�������ҡ������������������������ʹ��ʱ����������������Ϊ��������
        
        '������Դ (1-����,2-סԺ,3-����,4-���)
        If mblncmd���� Then str��Դ = "1,"
        If mblncmdסԺ Then str��Դ = str��Դ & "2,"
        If mblncmd���� Then str��Դ = str��Դ & "3,"
        If mblncmd��� Then str��Դ = str��Դ & "4,"
        If str��Դ <> "" Then
            str��Դ = Mid(str��Դ, 1, Len(str��Դ) - 1)
            strFilter = strFilter & " And Instr([23],B.������Դ)> 0"
        End If
        

            
        If mstrRoom <> "" Then  'ֻ��ʾִ�м䷶Χ�ڵ�
            If Not mblncmd�Ǽ� Then
                strFilter = strFilter & " And Instr([24],','|| A.ִ�м� || ',' )>0"
            Else
                strFilter = strFilter & " And (Instr([24],','|| A.ִ�м� || ',' )>0 And Nvl(A.ִ�й���,0)>1 OR Nvl(A.ִ�й���,0)<2)"
            End If
        End If
    
        If mblnNoShowCancel Then '����ʾȡ���Ǽǵļ��
            strFilter = strFilter & " And A.ִ��״̬<>2 "
        End If
        
        If mblncmd���� Then        'ֻ��ʾ����סԺ��¼
            strFilter = strFilter & vbNewLine & " And (B.������Դ=2 And B.��ҳID=C.סԺ���� Or Nvl(B.������Դ,0)<>2)"
        End If

        gstrSQL = "Select Distinct" & vbNewLine & _
                    "       A.ҽ��ID,A.���ͺ�,A.�״�ʱ�� ���ʱ��,A.����ʱ�� ����ʱ��,A.ִ��״̬,nvl(A.ִ�й���,0) ������,A.ִ�м�,A.������� ����," & vbNewLine & _
                    "       B.����ID,B.��ҳID,B.�Һŵ�,B.���˿���ID,Decode(B.������Դ, 1, '��', 2, 'ס', 3, '��', 4, '��') ��Դ,B.ҽ������,B.�걾��λ," & vbNewLine & _
                    "       Nvl(B.������־, 0) ������־, Nvl(B.Ӥ��, 0) Ӥ��,B.����ҽ��,A.NO,C.��ǰ����,C.��ǰ����ID,Decode(B.������Դ,2,C.סԺ��,C.�����) ��ʶ��," & vbNewLine & _
                    "       Nvl(H.����,C.����) ����,H.Ӱ�����,H.����,Nvl(H.�Ա�,C.�Ա�) �Ա�,Nvl(H.����,C.����) ����,H.���,H.����,H.Ӱ������," & vbNewLine & _
                    "       Decode(B.������Դ,3,B.����ҽ��,A.������) �Ǽ���,H.������," & vbNewLine & _
                    "       H.�����,H.�Ƿ��ӡ,H.�������,H.��ɫͨ��,H.�����ӡ,H.������,H.������,H.��鼼ʦ,H.�������� ��ͼʱ��,H.�������,H.���UID,0 as ת��" & vbNewLine & _
                    " From ����ҽ������ A,����ҽ����¼ B,������Ϣ C,Ӱ�����¼ H,Ӱ������Ŀ G" & vbNewLine & _
                    " Where B.���ID is NULL And A.ҽ��ID=B.ID And A.ҽ��ID=H.ҽ��ID(+) And A.���ͺ�=H.���ͺ�(+) And B.������ĿID=G.������ĿID And B.����ID=C.����ID"
        gstrSQL = gstrSQL & vbNewLine & strFilter & " And A.ִ�в���ID+0=[25]"
        
        If mblncmd�ѽ� Xor mblncmdδ�� Then '����ѡ��
            strFilter = "(Select Distinct No From ���˷��ü�¼ D Where A.NO = D.NO And A.��¼����=D.��¼���� And D.��¼״̬ = 1)"
            gstrSQL = gstrSQL & vbNewLine & IIf(mblncmd�ѽ�, " And Exists ", " And Not Exists") & strFilter
        End If
        
        If .���� <> 0 Then                        '��ʹ�ü��Ų���ʱһ���Ǳ������ģ�Ӱ�����¼���м�¼����ʱȡ�������ӱ���ȫ��ɨ��
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
        gstrSQL = "Select * From (" & vbNewLine & gstrSQL & vbNewLine & ") Order by ������,���ʱ��,����ʱ��"
    
        Set rsList = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����б�", .�����, .סԺ��, .���￨, .����, .���֤, .IC��, .���ݺ�, .����, .��ʼʱ��, .����ʱ��, _
                                            .���˿���, .�걾��λ, .���ҽ��, .���ҽ��, .Ӱ������, .��鼼ʦ, .Ӱ�����, .���, _
                                            .�������, .�������, .������, .����, str��Դ, mstrRoom, mlngCur����ID)
    End With
    
    strFilter = ""
    If mblncmd�Ǽ� Then strFilter = "������=0 or ������=1 or "
    If mblncmd���� Then strFilter = IIf(strFilter <> "", strFilter & "������=2 or ������=3 or ", "������=2 or ������=3 or ")
    If mblncmd���� Then strFilter = IIf(strFilter <> "", strFilter & "������=4 or ", "������=4 or ")
    If mblncmd��� Then strFilter = IIf(strFilter <> "", strFilter & "������=5 or ", "������=5 or ")
    If mblncmd��� Then strFilter = IIf(strFilter <> "", strFilter & "������=6 or ", "������=6 or ")
    If mblncmd�Ǽ� And mblncmd���� And mblncmd���� And mblncmd��� And mblncmd��� Then
        strFilter = ""
    End If

    If strFilter <> "" Then
        strFilter = Mid(strFilter, 1, Len(strFilter) - 4)
        rsList.Filter = strFilter
    End If
    
    Call FillList(rsList)
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub FillList(ByVal rsTemp As ADODB.Recordset)
Dim rsBaby As ADODB.Recordset
    On Error GoTo ErrHandle
    Call InitList
    If rsTemp.EOF Then stbThis.Panels(2).Text = "û���ҵ��κ�ƥ��ļ�¼": Exit Sub
    
    With vsList
        .Rows = rsTemp.RecordCount + 1
        Do Until rsTemp.EOF
            .Row = rsTemp.AbsolutePosition
            If rsTemp!������־ <> 0 Then
                Set .Cell(flexcpPicture, .Row, Col����) = imgList.ListImages("����").Picture
            End If
            If rsTemp!��Դ = "ס" Then
                Set .Cell(flexcpPicture, .Row, Col��Դ) = imgList.ListImages("סԺ").Picture
            Else
                .TextMatrix(.Row, Col��Դ) = rsTemp!��Դ
            End If
            If Nvl(rsTemp!����, 0) <> 0 Then
                Set .Cell(flexcpPicture, .Row, Col����) = imgList.ListImages("����").Picture
            End If
            
            If Nvl(rsTemp!��ɫͨ��, 0) <> 0 Then
                Set .Cell(flexcpPicture, .Row, Col����) = imgList.ListImages("��ɫͨ��").Picture
            End If
            
            If Nvl(rsTemp!���UID) <> "" Then
                Set .Cell(flexcpPicture, .Row, Col����) = imgList.ListImages("Ӱ��").Picture
            End If
            
            .TextMatrix(.Row, Col����) = Nvl(rsTemp!Ӱ������)
            .TextMatrix(.Row, Col����) = Nvl(rsTemp!����)
            .TextMatrix(.Row, Col����) = Nvl(rsTemp!����)
            .TextMatrix(.Row, Col������) = IIf(rsTemp!ִ��״̬ = 2, "�Ѿܾ�", Decode(Nvl(rsTemp!������, 0), 0, "�ѵǼ�", 1, "�ѵǼ�", _
                                                                                        2, IIf(Nvl(rsTemp!�������) <> "", "������", _
                                                                                                IIf(Nvl(rsTemp!������) = "", "�ѱ���", "������")), _
                                                                                        3, IIf(Nvl(rsTemp!�������) <> "", "������", _
                                                                                                IIf(Nvl(rsTemp!������) = "", "�Ѽ��", "������")), _
                                                                                        4, IIf(Nvl(rsTemp!�������) <> "", "������", _
                                                                                                IIf(Nvl(rsTemp!������) <> "", "�����", "�ѱ���")), _
                                                                                        5, "�����", "�����"))
            .TextMatrix(.Row, Col�Ա�) = Nvl(rsTemp!�Ա�)
            .TextMatrix(.Row, Col����) = Nvl(rsTemp!����)
            If InStr(Nvl(rsTemp!ҽ������), ":") > 0 Then '�µ�ģʽ������ҽ����������Ϣ�� ����,ִ�б��:��λ(����,����),��λ---
                .TextMatrix(.Row, Colҽ������) = Split(rsTemp!ҽ������, ":")(0)
                .TextMatrix(.Row, Col��λ����) = Split(rsTemp!ҽ������, ":")(1)
            Else
                .TextMatrix(.Row, Colҽ������) = Nvl(rsTemp!ҽ������)
            End If
            .TextMatrix(.Row, Colִ�м�) = Nvl(rsTemp!ִ�м�)
            .TextMatrix(.Row, Col���ʱ��) = Nvl(rsTemp!���ʱ��)
            .TextMatrix(.Row, Col����ʱ��) = Nvl(rsTemp!����ʱ��)
            .TextMatrix(.Row, Col����ҽ��) = Nvl(rsTemp!����ҽ��)
            .TextMatrix(.Row, Col���) = Nvl(rsTemp!���)
            .TextMatrix(.Row, Col����) = Nvl(rsTemp!����)
            .TextMatrix(.Row, ColӤ��) = Nvl(rsTemp!Ӥ��)
            .TextMatrix(.Row, Col�Ǽ���) = Nvl(rsTemp!�Ǽ���)
            .TextMatrix(.Row, Col������) = Nvl(rsTemp!������)
            .TextMatrix(.Row, Col�����) = Nvl(rsTemp!�����)
            .TextMatrix(.Row, Col��ӡ��Ƭ) = Nvl(rsTemp!�Ƿ��ӡ)
            .TextMatrix(.Row, Col�������) = Nvl(rsTemp!�������)
            .TextMatrix(.Row, Col��ɫͨ��) = Nvl(rsTemp!��ɫͨ��)
            .TextMatrix(.Row, Col�����ӡ) = Nvl(rsTemp!�����ӡ)
            .TextMatrix(.Row, Col������) = Nvl(rsTemp!������)
            .TextMatrix(.Row, Col������) = Nvl(rsTemp!������)
            .TextMatrix(.Row, Col��鼼ʦ) = Nvl(rsTemp!��鼼ʦ)
            .TextMatrix(.Row, Col��ͼʱ��) = Nvl(rsTemp!��ͼʱ��)
            .TextMatrix(.Row, ColӰ�����) = Nvl(rsTemp!Ӱ�����)
            .TextMatrix(.Row, Col����ID) = Nvl(rsTemp!����ID)
            .TextMatrix(.Row, Col��ҳID) = Nvl(rsTemp!��ҳID)
            .TextMatrix(.Row, Col�Һŵ�) = Nvl(rsTemp!�Һŵ�)
            .TextMatrix(.Row, Col���˿���ID) = Nvl(rsTemp!���˿���ID)
            .TextMatrix(.Row, Colҽ��ID) = Nvl(rsTemp!ҽ��ID)
            .TextMatrix(.Row, Col���ͺ�) = Nvl(rsTemp!���ͺ�)
            .TextMatrix(.Row, Col���UID) = Nvl(rsTemp!���UID)
            .TextMatrix(.Row, Col���״̬) = Nvl(rsTemp!������)
            .TextMatrix(.Row, Col�������) = Nvl(rsTemp!�������)
            .TextMatrix(.Row, ColNO) = Nvl(rsTemp!NO)
            .TextMatrix(.Row, Colת��) = Nvl(rsTemp!ת��)
            .TextMatrix(.Row, Col����) = Nvl(rsTemp!��ǰ����)
            .TextMatrix(.Row, Col��ǰ����ID) = Nvl(rsTemp!��ǰ����ID)
            .TextMatrix(.Row, Col��ʶ��) = Nvl(rsTemp!��ʶ��)
            
            If Nvl(rsTemp!Ӥ��) <> 0 Then
                gstrSQL = "Select Nvl(A.Ӥ������, B.���� || '֮��' || Trim(To_Char(A.���, '9'))) As Ӥ������, Ӥ���Ա�, ����ʱ��" & vbNewLine & _
                            "From ������������¼ A, ������Ϣ B" & vbNewLine & _
                            "Where A.����id = [1] And A.��ҳid = [2] And A.����id = B.����id And A.��� = [3]"

                Set rsBaby = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡӤ����Ϣ", CLng(rsTemp!����ID), CLng(Nvl(rsTemp!��ҳID, 0)), CLng(rsTemp!Ӥ��))
                If Not rsBaby.EOF Then
                    .TextMatrix(.Row, Col����) = rsBaby!Ӥ������
                    .TextMatrix(.Row, Col�Ա�) = Nvl(rsBaby!Ӥ���Ա�)
                    .TextMatrix(.Row, Col����) = Nvl(rsBaby!����ʱ��)
                End If
            End If
            
            If .TextMatrix(.Row, Col������) = "�Ѿܾ�" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = &HFFFF&
            If .TextMatrix(.Row, Col������) = "�����" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = &HFF00&
            
            rsTemp.MoveNext
        Loop
    End With
    stbThis.Panels(2).Text = "�� " & vsList.Rows - 1 & " ����¼": stbThis.Panels(2).Alignment = sbrCenter
    Exit Sub
ErrHandle:
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
Private Sub TabWindow_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Not mblnInitOk Then Exit Sub

    On Error GoTo ErrHandle
    If mblnIsHistory Then
        RefreshTabWindow mlngHOrderID
    ElseIf Val(vsList.TextMatrix(vsList.Row, Colҽ��ID)) = 0 Then
        RefreshTabWindow 0, True
    Else
        RefreshTabWindow 0, False, True
        If vsList.TextMatrix(vsList.Row, Col���UID) = "" And mfrmPACSImg.lvwSeq.ListItems.Count > 0 Then '�������ˢ�º�����ͼ�ˣ���ˢ�²����б�Ŀ����Ϊ���ù�Ƭ�Ȱ�������
            vsList.TextMatrix(vsList.Row, Col���UID) = mfrmPACSImg.lvwSeq.Tag
        End If
    End If
    
    'ɾ�����ڵĹ������������˵���
    Call LockWindowUpdate(Me.Hwnd)
    Dim lngCount As Long
    For lngCount = cbrMain.ActiveMenuBar.Controls.Count To 1 Step -1
        cbrMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbrMain.Count To 2 Step -1
        cbrMain(lngCount).Delete
    Next
    Call InitCommandBars
    
    Select Case Item.Tag
        Case "������д"
            If mblnPacsReport = True Then    'ʹ��PACS����༭��
                mfrmPacsReport.zlDefCommandBars Me.cbrMain
            Else
                mobjReport.zlDefCommandBars Me.cbrMain
            End If
        Case "�������"
            mobjExpense.zlDefCommandBars Me, Me.cbrMain
        Case "סԺҽ��"
            mobjInAdvice.zlDefCommandBars Me, Me.cbrMain, 2
        Case "����ҽ��"
            mobjOutAdvice.zlDefCommandBars Me, Me.cbrMain, 2
        Case "סԺ����"
            mobjInEPRs.zlDefCommandBars cbrMain
        Case "���ﲡ��"
            mobjOutEPRs.zlDefCommandBars cbrMain
        Case "�Ŷӽк�"
            mobjQueue.zlDefCommandBars cbrMain
    End Select
    Call LockWindowUpdate(0)
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub TimerRefresh_Timer()
    'ˢ�²����б�
    Call RefreshList
End Sub

Private Sub txtFilter_Change()
    If Not mobjIDCard Is Nothing Then
        mobjIDCard.SetEnabled (txtFilter.Text = "" And Me.ActiveControl Is txtFilter)
    End If
    If txtFilter.Text = "" Then txtFilter.Tag = ""
End Sub

Private Sub txtFilter_GotFocus()
    If mobjIDCard Is Nothing Then Set mobjIDCard = New clsIDCard         '���֤ʶ�����
    
    If txtFilter.Text <> "" Then Call zlControl.TxtSelAll(txtFilter)
    If InStr(mstrCurFindtype, "��  ��") > 0 Then
        Call zlCommFun.OpenIme(True)
    End If
    
    If Not mobjIDCard Is Nothing And txtFilter.Text = "" Then '�������֤�����豸
        mobjIDCard.SetEnabled (True)
    End If
End Sub
Private Sub txtFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtFilter_Validate(False)
        Call zlControl.TxtSelAll(txtFilter)
    End If
End Sub

Private Sub txtFilter_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        Select Case mstrCurFindtype
            Case "�����", "סԺ��"
                If InStr("*+0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case "���￨"
                Dim blnCard As Boolean
    
                'ȥ���ſ��������������ַ�
                If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
                
                blnCard = InputIsCard(Me.txtFilter, KeyAscii)
                
                'ˢ����ɻ�ȷ������
                If blnCard And Len(Me.txtFilter.Text) = Val(gbytCardLen) - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Me.txtFilter.Text <> "" Then
                    If KeyAscii <> 13 Then
                        Me.txtFilter.Text = Me.txtFilter.Text & Chr(KeyAscii)
                        Me.txtFilter.SelStart = Len(Me.txtFilter.Text)
                    End If
                    KeyAscii = 0
                    Me.txtFilter.Text = UCase(Me.txtFilter)
                    Me.txtFilter.SetFocus
                End If
            Case "���ݺ�"
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                If Not (txtFilter.Text = "" Or txtFilter.SelLength = Len(txtFilter.Text)) _
                    And InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            Case "����"
            
        End Select
    Else
        If Trim(txtFilter.Text) <> "" Then
            If Mid(txtFilter.Text, 1, 1) = "*" Then mstrCurFindtype = "�����"
            If Mid(txtFilter.Text, 1, 1) = "+" Then mstrCurFindtype = "סԺ��"
        End If
        Dim cbrControl As CommandBarControl
        Set cbrControl = cbrdock.FindControl(, ID_��ʼ����)
        If Not cbrControl Is Nothing Then
            cbrdock_Execute cbrControl
        End If
    End If
End Sub

Private Sub txtFilter_LostFocus()
    Call zlCommFun.OpenIme
    If Not mobjIDCard Is Nothing Then
        mobjIDCard.SetEnabled (False)
    End If
End Sub

Private Sub txtFilter_Validate(Cancel As Boolean)
    If InStr(mstrCurFindtype, "���ݺ�") > 0 Then
        If IsNumeric(txtFilter.Text) Then
            txtFilter.Text = GetFullNO(txtFilter.Text, 0)
        End If
    End If
End Sub
Private Sub Menu_Manage_��������豸()
Dim strModality As String
Dim rResult As VbMsgBoxResult
Dim strSQL As String
    
    frmChangeDevice.ShowMe UCase(vsList.TextMatrix(vsList.Row, ColӰ�����)), Me
    strModality = frmChangeDevice.strDeviceType
    
     If strModality <> "" Then
         strSQL = "Zl_Ӱ����_Ӱ�����(" & vsList.TextMatrix(vsList.Row, Colҽ��ID) & "," & vsList.TextMatrix(vsList.Row, Col���ͺ�) & ",'" & strModality & "')"
         ExecuteProc strSQL, Me.Caption
     End If
     
     'ˢ�²����б�
     Call RefreshList
End Sub

Private Sub sub3DProcess(strCommand As String)
    Dim str3DCommand As String
    Dim str3DImgDir As String

    str3DImgDir = App.Path & "\TmpImage\3D\"

    '��֯��ά�ؽ����
    str3DCommand = mstr3DExeDir & " " & mstr3DPara & " " & strCommand & " " & str3DImgDir
    On Error Resume Next
    Shell str3DCommand
End Sub

Private Sub sub��ά�ؽ�(strCommand As String)

    If TabWindow.Selected.Tag <> "Ӱ��ͼ��" Then '��ˢ��ͼ������
        Call mfrmPACSImg.zlRefresh(vsList.TextMatrix(vsList.Row, Colҽ��ID), vsList.TextMatrix(vsList.Row, Col���ͺ�), mstrPrivs, vsList.TextMatrix(vsList.Row, Colת��) = 1)
    End If
    
    '��֯��ά�ؽ���Ҫ��ͼ��
    If mfrmPACSImg.fun3DImgProcess = True Then Call sub3DProcess(strCommand)
End Sub


Private Sub Menu_Manage_���()
Dim strReview As String

    On Error GoTo ErrHandle
    
    If frmReview.ShowMe(vsList.TextMatrix(vsList.Row, Colҽ��ID), vsList.TextMatrix(vsList.Row, Col���ͺ�), Me, strReview) = True Then
        vsList.TextMatrix(vsList.Row, Col�������) = strReview
    End If

    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function zlInQueue(str�������� As String, lngҵ��ID As Long, lng����ID As Long, _
        str�������� As String, str���� As String, strҽ������ As String, Optional str�Ŷӱ�� As String = "", Optional lng�ŶӺ��� As Long) As Boolean
        
        On Error GoTo err
        
        If mblnUseQueue = True Then
            mobjQueue.zlInQueue str��������, lngҵ��ID, lng����ID, str��������, str����, strҽ������, str�Ŷӱ��, lng�ŶӺ���
        End If
        zlInQueue = True
        Exit Function
err:
    If ErrCenter = 1 Then
        Resume
    End If
        
End Function

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'����:��ʾ���˿�Ƭ��ť
    If vsList.TextMatrix(NewRow, Colҽ��ID) = "" Then
        cmdInfo.Visible = False
    Else
        cmdInfo.Height = vsList.CellHeight - 30
        cmdInfo.Left = vsList.Cell(flexcpLeft, NewRow, Col����) + vsList.Cell(flexcpWidth, NewRow, Col����) - cmdInfo.Width - 15
        cmdInfo.Top = vsList.CellTop + 15
        cmdInfo.Visible = True
    End If
End Sub

Private Sub vsList_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
'����:��ʾ���˿�Ƭ��ť
    If vsList.TextMatrix(vsList.Row, Colҽ��ID) = "" Then
        cmdInfo.Visible = False
    Else
        If NewLeftCol > Col���� Then
            cmdInfo.Visible = False
        Else
            cmdInfo.Height = vsList.CellHeight - 30
            cmdInfo.Left = vsList.Cell(flexcpLeft, vsList.Row, Col����) + vsList.Cell(flexcpWidth, vsList.Row, Col����) - cmdInfo.Width - 15
            cmdInfo.Top = vsList.CellTop + 15
            cmdInfo.Visible = True
        End If
    End If
End Sub

Private Sub vsList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
'����:��ʾ���˿�Ƭ��ť
    If vsList.TextMatrix(vsList.Row, Colҽ��ID) = "" Then
        cmdInfo.Visible = False
    Else
        cmdInfo.Height = vsList.CellHeight - 30
        cmdInfo.Left = vsList.Cell(flexcpLeft, vsList.Row, Col����) + vsList.Cell(flexcpWidth, vsList.Row, Col����) - cmdInfo.Width - 15
        cmdInfo.Top = vsList.CellTop + 15
        cmdInfo.Visible = True
    End If
End Sub

Private Sub vsList_DblClick()
    If vsList.TextMatrix(vsList.Row, Colҽ��ID) <> "" Then
        Select Case vsList.TextMatrix(vsList.Row, Col���״̬)
            Case 1, 0
                Call Menu_Manage_����
            Case 2, 3               '˫������д����,�����ʱ�����趨�Ƿ�򿪹�Ƭվ
                Call Menu_RichEPR(conMenu_Edit_Modify)
            Case 4, 5               '˫���޶�����,�����ʱ�����趨�Ƿ�򿪹�Ƭվ
                Call Menu_RichEPR(conMenu_Edit_Audit)
            Case 6                  '����
                Call Menu_RichEPR(conMenu_File_Open)
        End Select
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
    On Error GoTo ErrHandle
    mblnIsHistory = False
    '�ж�Ƕ��ʽ����༭���еı����Ƿ�û�б���
    If mblnPacsReport = True Then    'ʹ��PACS����༭��
        Call mfrmPacsReport.PromptModify
    End If
    
    If Val(vsList.TextMatrix(vsList.Row, Colҽ��ID)) = 0 Then '�޼�¼ʱ����
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
        Call ShowTab '���ݲ����ṩ��ͬѡ�
        
        If mstrFirstTab <> "" Then '��Ϊ�ձ�ʾ��������ҳ��ʾ,��TabWindow����ˢ��
            Dim i As Integer
            For i = 0 To TabWindow.ItemCount - 1
                If InStr(TabWindow.Item(i).Tag, mstrFirstTab) > 0 And TabWindow.Item(i).Visible Then
                    If TabWindow.Item(i).Selected Then
                        Call RefreshTabWindow
                    Else
                        TabWindow.Item(i).Selected = True
                    End If
                    Exit Sub
                End If
            Next
            If i = TabWindow.ItemCount Then TabWindow(0).Selected = True 'ûѭ�����˴�����0��tab
        Else
            Call RefreshTabWindow
        End If
        
        Call ShowBillList(cbrMain.FindControl(, conMenu_Manage_RequestPrint, , True))  '��ʾ�ɴ�ӡ�����Ƶ���:֮���Լ�ʱ����,��Ϊ��ʹ��F2�ȼ�
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub FillTxtInfor(Optional lngAdviceID As Long = 0)
'������Ϸ����˻�����Ϣ
Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrHandle
    
    With vsList
        lbl������Ϣ.Caption = "��  ��:" & Rpad(.TextMatrix(.Row, Col����), 12, " ") & "��  ��:" & Rpad(.TextMatrix(.Row, Col�Ա�), 13, " ") & _
                          "��  ��:" & Rpad(.TextMatrix(.Row, Col����), 10, " ") & "��ʶ��:" & Rpad(.TextMatrix(.Row, Col��ʶ��), 12, " ") & _
                          "��  ��:" & Rpad(.TextMatrix(.Row, Col����) & "", 10, " ")
                          
        If lngAdviceID = 0 Then '---------------------------�����μ��ֱ�����б��м�¼���
            gstrSQL = "Select ���� From ���ű� Where ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���˿���", CLng(.TextMatrix(.Row, Col���˿���ID)))
            lbl�����Ϣ.Caption = "����:" & Rpad(.TextMatrix(.Row, Col����), 12, " ") & "���˿���:" & Rpad(rsTemp!����, 11, " ") & _
                                  "����ҽ��:" & Rpad(.TextMatrix(.Row, Col����ҽ��), 8, " ") & "�����Ŀ:" & .TextMatrix(.Row, Colҽ������)
            If .TextMatrix(.Row, Col��λ����) <> "" Then lbl�����Ϣ.Caption = lbl�����Ϣ.Caption & "(" & .TextMatrix(.Row, Col��λ����) & ")"
            lblCash.Caption = "��": lblCash.Visible = False
            lblCash.Visible = CheckChargeState(.TextMatrix(.Row, Colҽ��ID)) = 1
        Else
            Dim strSQLBak As String
            gstrSQL = "Select A.ID, A.���˿���id, A.����ҽ��,A.������Դ, A.ҽ������, Nvl(A.Ӥ��, 0) Ӥ��, A.����id, A.��ҳid, A.�Һŵ�, B.����, B.���uid, C.����, D.���ͺ�,D.ִ��״̬,0 as ת��" & vbNewLine & _
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
                mlngHOrderID = lngAdviceID
                mlngHSendNo = Nvl(rsTemp!���ͺ�, 0)
                mstrHStudyUID = Nvl(rsTemp!���UID)
                mblnHMoved = IIf(rsTemp!ת�� = 1, True, False)
                fraInfo.Tag = rsTemp!����ID & "|" & rsTemp!��ҳID & "|" & rsTemp!ID & "|" & rsTemp!���ͺ� & "|" & rsTemp!���˿���ID & "|" & rsTemp!�Һŵ� & "|" & Nvl(rsTemp!������Դ, 3) & "|" & rsTemp!���UID & "|" & rsTemp!ת�� & "|" & rsTemp!ִ��״̬
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
                                            "��  ��:" & Rpad(rsTemp!����ʱ��, 10, " ") & "��ʶ��:" & Rpad(.TextMatrix(.Row, Col��ʶ��), 12, " ") & _
                                            "��  ��:" & Rpad(.TextMatrix(.Row, Col����) & "", 10, " ")
                    End If
                End If
            Else
                lbl�����Ϣ.Caption = "����:" & Space(12) & "���˿���:" & Space(11) & "����ҽ��:" & Space(8) & "�����Ŀ:"
            End If
            lblCash.Caption = "��": lblCash.Visible = True
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub FillTxtAppend(Optional lngAdviceIDtmp As Long = 0)
'������½�ҽ������
Dim lngAdviceID As Long, strAppend As String, rsTemp As ADODB.Recordset, i As Integer
    On Error GoTo ErrHandle
    With vsList
        If lngAdviceIDtmp = 0 Then
            lngAdviceID = Val(.TextMatrix(.Row, Colҽ��ID))
        Else
            lngAdviceID = lngAdviceIDtmp
        End If
        
        If lngAdviceIDtmp = 0 Then '-------------------------------------------�б�ѡ�����
            If .TextMatrix(.Row, Col��λ����) <> "" Then
                For i = 0 To UBound(Split(.TextMatrix(.Row, Col��λ����), "),"))
                    If i = 0 Then
                        txtAppend = "��鲿λ:" & vbCrLf & Space(2) & "1:" & Split(.TextMatrix(.Row, Col��λ����), "),")(i) & ")"
                    Else
                        txtAppend = txtAppend & vbCrLf & Space(2) & i + 1 & ":" & Split(.TextMatrix(.Row, Col��λ����), "),")(i) & ")"
                    End If
                Next
                If Trim(txtAppend) <> "" Then txtAppend = Mid(txtAppend, 1, Len(txtAppend) - 1) 'ȡ����������
            Else
                txtAppend = "��鲿λ:" & .TextMatrix(.Row, Colҽ������)
            End If
            gstrSQL = "Select ��Ŀ,���� From ����ҽ������ Where ҽ��ID=[1] Order By ����"
            If .TextMatrix(.Row, Colת��) = 1 Then gstrSQL = Replace(gstrSQL, "����ҽ������", "H����ҽ������")
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
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub FillHistory()
'������μ���¼
Dim rsTemp As ADODB.Recordset, strTemp As String
    On Error GoTo ErrHandle
    With vsList
        cboTimes.Tag = "" 'cbotime����ʱ�õ�������������"������Ŀ"ʱ��������"���cbotimes"����
        gstrSQL = "Select A.ID ҽ��ID,A.����ʱ��  ����ʱ��,A.ҽ������ " & _
                   " From ����ҽ����¼ A,����ҽ������ B,Ӱ�����¼ C" & _
                   " Where A.����id = [1] And A.���id Is Null And A.ִ�п���id+0 =[2] And B.ҽ��ID=A.ID " & _
                   "" & IIf(.TextMatrix(.Row, Col������) = "�Ѿܾ�", "", " And B.ִ��״̬<>2 ") & _
                   " AND A.ID=C.ҽ��ID"
        strTemp = Replace(gstrSQL, "����ҽ����¼", "H����ҽ����¼")
        strTemp = Replace(strTemp, "����ҽ������", "H����ҽ������")
        strTemp = Replace(strTemp, "Ӱ�����¼", "HӰ�����¼")
        gstrSQL = gstrSQL & vbNewLine & " Union ALL " & vbNewLine & strTemp
        gstrSQL = "Select * From (" & vbNewLine & gstrSQL & vbNewLine & ") Order By ����ʱ�� Asc"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "", CLng(.TextMatrix(.Row, Col����ID)), mlngCur����ID)
        
        cboTimes.Clear
        Do Until rsTemp.EOF
           cboTimes.AddItem "��" & rsTemp.AbsolutePosition & "��(" & Format(rsTemp!����ʱ��, "yyyy-mm-dd") & ")  " & Trim(rsTemp!ҽ������)
           cboTimes.ItemData(cboTimes.NewIndex) = rsTemp!ҽ��ID
           If rsTemp!ҽ��ID = .TextMatrix(.Row, Colҽ��ID) Then cboTimes.ListIndex = cboTimes.NewIndex
           rsTemp.MoveNext
        Loop
        cboTimes.Tag = "���"
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub ShowTab(Optional lngAdviceID As Long = 0)
'���ݲ�����Դ���Ʋ�����ҽ��ѡ�
Dim int��Դ As Integer, i As Integer
    On Error GoTo ErrHandle
    
    If lngAdviceID = 0 Then '-------------------------------------------�б�ѡ�����
        int��Դ = Val(vsList.TextMatrix(vsList.Row, Col��Դ))
        Dim blnShowReport As Boolean
        '�ж� ��ͼ����д����
        blnShowReport = True
        If mblnReportWithImage = True Then
            If vsList.TextMatrix(vsList.Row, Col���UID) = "" Then blnShowReport = False
        End If
    Else                    '-------------------------------------------���μ�¼ѡ�����
        '���μ�¼ʱfraInfo.Tag = 0����ID|1��ҳID|2ҽ��ID|3���ͺ�|4���˿���ID|5�Һŵ�|6������Դ|7���UID|8ת��
        int��Դ = Split(fraInfo.Tag, "|")(6)
    End If
    
    If int��Դ <> 2 Then '���ݲ�����Դ���Ʋ�����ҽ��ѡ�
        For i = 0 To TabWindow.ItemCount - 1
            Select Case TabWindow(i).Tag
                Case "���ﲡ��", "����ҽ��"
                    TabWindow(i).Visible = True
                Case "סԺ����", "סԺҽ��"
                    TabWindow(i).Visible = False
                Case "Ӱ��ͼ��"
                    TabWindow(i).Visible = True
                Case "������д"
                    TabWindow(i).Visible = IIf(lngAdviceID = 0, vsList.TextMatrix(vsList.Row, Col���״̬) > 1 And blnShowReport, True)
            End Select
        Next
    Else
        For i = 0 To TabWindow.ItemCount - 1
            Select Case TabWindow(i).Tag
                Case "���ﲡ��", "����ҽ��"
                    TabWindow(i).Visible = False
                Case "סԺ����", "סԺҽ��"
                    TabWindow(i).Visible = True
                Case "Ӱ��ͼ��"
                    TabWindow(i).Visible = True
                Case "������д"
                    TabWindow(i).Visible = IIf(lngAdviceID = 0, vsList.TextMatrix(vsList.Row, Col���״̬) > 1 And blnShowReport, True)
            End Select
        Next
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub RefreshTabWindow(Optional lngAdviceIDtmp As Long = 0, Optional blnClear As Boolean = False, Optional blnRefresh As Boolean = False)
'lngAdviceIDtmp���μ�¼ʱ���� , ������0, blnclear��յ�ǰ�б�, blnRefreshǿ��ˢ��
'ˢ�µ�ǰҳ��,���ã��б�ѡ�����μ�¼ѡ���Ӵ���ѡ��
'���μ�¼ʱfraInfo.Tag = 0����ID|1��ҳID|2ҽ��ID|3���ͺ�|4���˿���ID|5�Һŵ�|6������Դ|7���UID|8ת��|9ִ��״̬
Dim lngAdviceID As Long, lngSendNO As Long, lngPatID As Long, lngPageID As Long
Dim lngUnit As Long, lngPatDept As Long, strRegNo As String, intMoved As Boolean, intState As Integer, i As Integer

    On Error GoTo ErrHandle
    If lngAdviceIDtmp = 0 Then '-----------------------�б�ѡ�����
        If blnClear Then       '�޼�¼ʱ��������Ӵ���
            lngAdviceID = 0: lngSendNO = 0: lngPatID = 0: lngPageID = 0
            lngPatDept = 0: strRegNo = "": intMoved = 0: intState = 0: lngUnit = 0
        Else
            With vsList
                lngAdviceID = .TextMatrix(.Row, Colҽ��ID): lngSendNO = .TextMatrix(.Row, Col���ͺ�)
                lngPatID = .TextMatrix(.Row, Col����ID): lngPageID = Val(.TextMatrix(.Row, Col��ҳID))
                lngPatDept = .TextMatrix(.Row, Col���˿���ID): strRegNo = .TextMatrix(.Row, Col�Һŵ�)
                intMoved = .TextMatrix(.Row, Colת��): intState = .TextMatrix(.Row, Col���״̬)
                lngUnit = Val(.TextMatrix(.Row, Col��ǰ����ID))
            End With
        End If
    Else                       '----------------------���μ�¼ѡ�����
        lngAdviceID = lngAdviceIDtmp: lngSendNO = Split(fraInfo.Tag, "|")(3)
        lngPatID = Split(fraInfo.Tag, "|")(0): lngPageID = Val(Split(fraInfo.Tag, "|")(1))
        lngPatDept = Split(fraInfo.Tag, "|")(4): strRegNo = Split(fraInfo.Tag, "|")(5)
        intMoved = Split(fraInfo.Tag, "|")(8): intState = Split(fraInfo.Tag, "|")(9)
        lngUnit = lngPatDept
    End If
    
    mfrmPACSImg.zlRefresh lngAdviceID, lngSendNO, mstrPrivs, intMoved = 1, blnRefresh
    
    For i = 0 To TabWindow.ItemCount - 1 'ѭ�����˲�ˢ��,����û����Ȩ��ģ��
        If TabWindow(i).Tag = "������д" Then
            If mblnPacsReport = True Then
                mfrmPacsReport.zlRefresh lngAdviceID, lngSendNO, mlngCur����ID, mstrPrivs, mlngModul, Me, intMoved = 1
            Else
                mobjReport.zlRefresh lngAdviceID, mlngCur����ID, Not mblnIsHistory, intMoved = 1
            End If
        End If
    Next
    
    Select Case TabWindow(TabWindow.Selected.Index).Tag
        Case "�������"
            mobjExpense.zlRefresh mlngCur����ID, lngAdviceID, lngSendNO, intMoved = 1
        Case "�Ŷӽк�"
            If Not mblnIsHistory Then
                mobjQueue.zlRefresh gcnOracle, mAstr��������, Split(mstrCur����, "-")(1) & vsList.TextMatrix(vsList.Row, Colִ�м�), lngAdviceID
            End If
        Case "סԺҽ��"
            If TabWindow.Selected.Visible Then '������סԺ��¼ת�����������¼,��ʱ����û����Ȩ����ҽ��Ȩ��
                mobjInAdvice.zlRefresh lngPatID, lngPageID, lngUnit, lngPatDept, 0, intMoved = 1, lngAdviceID, intState
            Else
                For i = 0 To TabWindow.ItemCount - 1 'ѭ�����˲Ŵ���
                    If TabWindow(i).Tag = "����ҽ��" Then
                        If strRegNo = "" Then   '���еǼǵĲ���û�йҺŵ���
                            mobjOutAdvice.zlRefresh lngPatID, "", False
                        Else
                            mobjOutAdvice.zlRefresh lngPatID, strRegNo, Not mblnIsHistory, intMoved = 1, lngAdviceID
                        End If
                    End If
                Next
            End If
        Case "����ҽ��"
            If TabWindow.Selected.Visible Then '�����������¼ת������סԺ��¼,��ʱ����û����ȨסԺҽ��Ȩ��
                If strRegNo = "" Then   '���еǼǵĲ���û�йҺŵ���
                    mobjOutAdvice.zlRefresh lngPatID, "", False
                Else
                    mobjOutAdvice.zlRefresh lngPatID, strRegNo, Not mblnIsHistory, intMoved = 1, lngAdviceID
                End If
            Else
                For i = 0 To TabWindow.ItemCount - 1 'ѭ�����˲Ŵ���
                    If TabWindow(i).Tag = "סԺҽ��" Then
                      mobjInAdvice.zlRefresh lngPatID, lngPageID, lngUnit, lngPatDept, 0, intMoved = 1, lngAdviceID, intState
                    End If
                Next
            End If
        Case "סԺ����"
            If TabWindow.Selected.Visible Then '������סԺ��¼ת�����������¼,��ʱ����û����Ȩ���ﲡ��Ȩ��
                mobjInEPRs.zlRefresh lngPatID, lngPageID, mlngCur����ID, Not mblnIsHistory, intMoved = 1
            Else
                For i = 0 To TabWindow.ItemCount - 1 'ѭ�����˲Ŵ���
                    If TabWindow(i).Tag = "���ﲡ��" Then
                       mobjOutEPRs.zlRefresh lngPatID, lngPageID, mlngCur����ID, Not mblnIsHistory, intMoved = 1
                    End If
                Next
            End If
        Case "���ﲡ��"
            If TabWindow.Selected.Visible Then '�����������¼ת������סԺ��¼,��ʱ����û����ȨסԺ����Ȩ��
                mobjOutEPRs.zlRefresh lngPatID, lngPageID, mlngCur����ID, Not mblnIsHistory, intMoved = 1
            Else
                For i = 0 To TabWindow.ItemCount - 1 'ѭ�����˲Ŵ���
                    If TabWindow(i).Tag = "סԺ����" Then
                        mobjInEPRs.zlRefresh lngPatID, lngPageID, mlngCur����ID, Not mblnIsHistory, intMoved = 1
                    End If
                Next
            End If
    End Select
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
