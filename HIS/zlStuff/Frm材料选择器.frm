VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm����ѡ���� 
   Caption         =   "��������ѡ����"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9465
   Icon            =   "Frm����ѡ����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   9465
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picSplit02_S 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   40
      Left            =   2625
      ScaleHeight     =   45
      ScaleWidth      =   2535
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6240
      Width           =   2535
   End
   Begin VB.CheckBox chkContinue 
      Caption         =   "����ѡ��(&M)"
      Height          =   180
      Left            =   2760
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox picѡ���� 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   2625
      ScaleHeight     =   2535
      ScaleWidth      =   4815
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6360
      Width           =   4815
      Begin VB.PictureBox picUpDown01 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   220
         Left            =   3600
         Picture         =   "Frm����ѡ����.frx":0E42
         ScaleHeight     =   225
         ScaleWidth      =   270
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   270
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfѡ�� 
         Height          =   2085
         Left            =   0
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   4275
         _cx             =   7541
         _cy             =   3678
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   15724527
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   32
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"Frm����ѡ����.frx":1184
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
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
      End
      Begin VB.Label lblѡ�� 
         BackColor       =   &H00FFEDDD&
         Caption         =   "ѡ������"
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   3885
      End
   End
   Begin VB.CommandButton Cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8250
      TabIndex        =   4
      Top             =   5850
      Width           =   1100
   End
   Begin VB.CommandButton Cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   350
      Left            =   7020
      TabIndex        =   3
      Top             =   5850
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msf���Ϲ�� 
      Height          =   3675
      Left            =   2640
      TabIndex        =   1
      Top             =   405
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   6482
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ImageList ImgTvw 
      Left            =   2010
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm����ѡ����.frx":15ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm����ѡ����.frx":2C47
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   5715
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   10081
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImgTvw"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImgLvwSmall 
      Left            =   8820
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm����ѡ����.frx":4951
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf���� 
      Height          =   1620
      Left            =   2625
      TabIndex        =   2
      Top             =   4155
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   2858
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ImageList imgsMain 
      Left            =   240
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm����ѡ����.frx":665B
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm����ѡ����.frx":69AD
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.Image ImgLeftRight_S 
      Height          =   4485
      Left            =   2580
      MousePointer    =   9  'Size W E
      Top             =   1290
      Width           =   45
   End
   Begin VB.Image ImgUpDown_S 
      Height          =   45
      Left            =   2640
      MousePointer    =   7  'Size N S
      Top             =   4080
      Width           =   6765
   End
End
Attribute VB_Name = "Frm����ѡ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--�������--
Private mintEditState As Integer                 '�༭״̬(1-���;2-����)
Private mlngԴ�ⷿID As Long                     'Դ�ⷿID
Private mlngĿ�ⷿID As Long                     'Ŀ�ⷿID
Private mlngʹ�ò���ID As Long                   'ʹ�ò���ID
Private mlng��Ӧ��ID As Long                     '��Ӧ��ID
Private mobjOut As Form                          'ʹ�ñ�����Ĵ��壨�����ṩһ��������¼�������Է��أ�

Private mblnStartUp As Boolean                   '�����ɹ�
Private mblnFirstStart As Boolean                '��һ������
Private mrsUnit As New ADODB.Recordset           '��λ
Private mstrUnit As String                       '��λ����
Private mstrUnitString As String                 'SQL�ִ�
Private mintStockCheck As Integer                '�����
Private mbln�̵㵥 As Boolean                    '�̵㵥�ݱ�־
Private mbln������ As Boolean                    '�Ƿ����ӿ����ι�����
Private mblnCheck As Boolean                     '�Ƿ�����(�̵㡢���á�������)
Private mblnPrice As Boolean                     '�Ƿ�����ʱ�ۻ��������������
Private mblnTrackUsing As Boolean                '�������ò���

'������ʹ�ü�¼��
Private mrsData As New ADODB.Recordset           '������;����
Private mrsCard As New ADODB.Recordset           '���Ŀ�Ƭ
Private mrsStock As New ADODB.Recordset          '���Ĺ��

'���ؼ�¼��
Private mrsReturn As ADODB.Recordset            '���ؼ�¼��(������Ϣ������,����Ŀ¼������,���Ŀ��������)
Private mint�ⷿ As Integer                      '1-���Ŀ�;2-���ϲ���;3-�Ƽ���
Private mint���� As Integer                      '0-������;1-�ⷿ����;2-���÷���;3-���Ŀ����÷���
Private mblnֻ��ʾ���ٲ��� As Boolean
Private mblnʱ�� As Boolean                      'ʱ��
Private mblnStock  As Boolean
Private mstrCardSortBy As String                 '���Ŀ�Ƭ������
Private mstrPhysicSortBy As String               '���Ĺ��������
Private mlngCardRow As Long
Private mlngPhysicRow As Long
Private mlngLastSelect����ID As Long             '�ϴ�ѡ��Ĳ���ID�������Ƿ�ˢ�£�
Private mbln����ʾ������� As Boolean
Private mbln���޴洢�ⷿ���� As Boolean
Private mblnCostView As Boolean                 '�鿴�ɱ��� true-����鿴 false-������鿴
Private mblnProvider As Boolean                 '�鿴�ϴι�Ӧ�������Ϣ true-����鿴 false-������鿴
Private mbln��ʾ���� As Boolean                 'true-��ʾ�����б�false-����ʾ�����б�
Private mstrPrivs As String                    '����ԱȨ��
Private mlngModule As Long
Private mbln�Ƿ���� As Boolean                '�Ƿ��ɹ��˴�

'����get���ÿ��󣬷��صĿ���������ʵ��������ʵ�ʽ�ʵ�ʲ��
Private msin�������� As Single
Private msinʵ������ As Single
Private msinʵ�ʽ�� As Single
Private msinʵ�ʲ�� As Single
Private mblnɢװ��λ As Boolean
Private mstr�̵�ʱ�� As String
Private Enum mCol
    ����id = 0
    ����ID
    ����id
    ����
    ��������
    ��Ʒ��
    ���
    ����
    ��׼�ĺ�
    ע��֤��
    �ϴι�Ӧ��
    �ۼ�
    ���³ɱ���
    ɢװ��λ
    ����ϵ��
    ��װ��λ
    ��������
    �������
    �����
    �����
    ��Ч��
    ���Ч��
    ���ʧЧ��
    һ���Բ���
    �޾��Բ���
    �ⷿ����
    ���÷���
    ʱ��
    ָ��������
    ָ�������
    �ⷿ��λ
End Enum


'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------


'--����--
Private Const mCols = 31            '������

Public Property Get In_�༭״̬() As Integer
    In_�༭״̬ = mintEditState
End Property

Public Property Let In_�༭״̬(ByVal vNewValue As Integer)
    mintEditState = vNewValue
End Property

Public Property Get In_Դ�ⷿ() As Long
    In_Դ�ⷿ = mlngԴ�ⷿID
End Property

Public Property Let In_Դ�ⷿ(ByVal vNewValue As Long)
    mlngԴ�ⷿID = vNewValue
End Property

Public Property Get In_Ŀ�ⷿ() As Long
    In_Ŀ�ⷿ = mlngĿ�ⷿID
End Property

Public Property Let In_Ŀ�ⷿ(ByVal vNewValue As Long)
    mlngĿ�ⷿID = vNewValue
End Property

Public Property Get In_����() As Long
    In_���� = mlngʹ�ò���ID
End Property

Public Property Let In_����(ByVal vNewValue As Long)
    mlngʹ�ò���ID = vNewValue
End Property

Public Property Let In_MainFrm(ByVal vNewValue As Form)
    Set mobjOut = vNewValue
End Property

Private Sub SetFormat(Optional ByVal IntMain As Integer = 1, Optional ByVal BlnSetHeader As Boolean = False)
    Dim intCol As Integer
    
    '���ø��б�ؼ��ĸ�ʽ
    Select Case IntMain
    Case 1
        With msf���Ϲ��
            
            If BlnSetHeader Then
                .Cols = mCols
                .TextMatrix(0, mCol.����id) = "����ID"
                .TextMatrix(0, mCol.����ID) = "����ID"
                .TextMatrix(0, mCol.����id) = "����ID"
                .TextMatrix(0, mCol.����) = "����"
                .TextMatrix(0, mCol.��������) = "��������"
                .TextMatrix(0, mCol.��Ʒ��) = "��Ʒ��"
                .TextMatrix(0, mCol.���) = "���"
                .TextMatrix(0, mCol.����) = "����"
                .TextMatrix(0, mCol.�ۼ�) = "�ۼ�"
                .TextMatrix(0, mCol.ɢװ��λ) = "ɢװ��λ"
                .TextMatrix(0, mCol.����ϵ��) = "����ϵ��"
                .TextMatrix(0, mCol.��װ��λ) = "��װ��λ"
                .TextMatrix(0, mCol.��������) = "��������"
                .TextMatrix(0, mCol.�������) = "�������"
                .TextMatrix(0, mCol.�����) = "�����"
                .TextMatrix(0, mCol.�����) = "�����"
                .TextMatrix(0, mCol.��Ч��) = "��Ч��"
                .TextMatrix(0, mCol.�ⷿ����) = "�ⷿ����"
                .TextMatrix(0, mCol.���÷���) = "���÷���"
                .TextMatrix(0, mCol.һ���Բ���) = "һ���Բ���"
                .TextMatrix(0, mCol.�޾��Բ���) = "�޾��Բ���"
                .TextMatrix(0, mCol.���Ч��) = "���Ч��"
                .TextMatrix(0, mCol.���ʧЧ��) = "���ʧЧ��"
                .TextMatrix(0, mCol.ʱ��) = "ʱ��"
                .TextMatrix(0, mCol.ָ��������) = "ָ��������"
                .TextMatrix(0, mCol.ָ�������) = "ָ�������"
                .TextMatrix(0, mCol.�ⷿ��λ) = "�ⷿ��λ"
                .TextMatrix(0, mCol.���³ɱ���) = "���³ɱ���"
                
            End If
            
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
                If intCol >= mCol.�������� And intCol <= mCol.����� Or intCol >= mCol.ʱ�� And intCol <= mCol.ָ������� Or intCol = mCol.�ۼ� Or intCol = mCol.����ϵ�� Or intCol = mCol.���³ɱ��� Then
                    .ColAlignment(intCol) = 7
                ElseIf intCol = mCol.ɢװ��λ Or intCol = mCol.��װ��λ Or intCol = mCol.���Ч�� Or intCol = mCol.һ���Բ��� Or intCol = mCol.�޾��Բ��� Then
                    .ColAlignment(intCol) = 4
                Else
                    .ColAlignment(intCol) = 1
                End If
            Next
            
            If mblnStartUp = False Then
                .ColWidth(mCol.����id) = 0
                .ColWidth(mCol.����ID) = 0
                .ColWidth(mCol.����id) = 0
                .ColWidth(mCol.����) = 800
                .ColWidth(mCol.��������) = 2000
                .ColWidth(mCol.��Ʒ��) = 2000
                .ColWidth(mCol.���) = 1600
                .ColWidth(mCol.����) = 1500
                .ColWidth(mCol.�ۼ�) = 1000
                .ColWidth(mCol.ɢװ��λ) = 800
                .ColWidth(mCol.����ϵ��) = 800
                .ColWidth(mCol.��װ��λ) = 800
                .ColWidth(mCol.��������) = 1000
                .ColWidth(mCol.�������) = 1000
                .ColWidth(mCol.�����) = 1000
                .ColWidth(mCol.��Ч��) = 1000
                .ColWidth(mCol.���ʧЧ��) = 1000
                .ColWidth(mCol.���Ч��) = 0
                .ColWidth(mCol.���Ч��) = 0
                .ColWidth(mCol.һ���Բ���) = 800
                .ColWidth(mCol.�޾��Բ���) = 800
                .ColWidth(mCol.�ⷿ����) = 800
                .ColWidth(mCol.���÷���) = 800
                .ColWidth(mCol.ʱ��) = 1000
                .ColWidth(mCol.ָ�������) = 0
                .ColWidth(mCol.�ⷿ��λ) = 1000
                .Row = 1
                RestoreFlexState msf���Ϲ��, Me.Name
                .ColWidth(mCol.�����) = IIf(mblnCostView = False, 0, 1000)
                .ColWidth(mCol.���³ɱ���) = IIf(mblnCostView = False, 0, 1000)
                .ColWidth(mCol.ָ��������) = IIf(mblnCostView = False, 0, 1000)
                If mlngModule = 1725 Or .ColWidth(mCol.�ϴι�Ӧ��) = 0 Then .ColWidth(mCol.�ϴι�Ӧ��) = IIf(mblnProvider = False, 0, 1300)
            End If
        End With
    Case 0
        With Msf����
            
            If BlnSetHeader Then
                .Cols = 17
                .TextMatrix(0, 0) = ""
                .TextMatrix(0, 1) = "�ⷿ"
                .TextMatrix(0, 2) = "����"
                .TextMatrix(0, 3) = "����"
                .TextMatrix(0, 4) = "ʧЧ��"
                .TextMatrix(0, 5) = "��������"
                .TextMatrix(0, 6) = "�������"
                .TextMatrix(0, 7) = "�����"
                .TextMatrix(0, 8) = "�����"
                .TextMatrix(0, 9) = "���ʧЧ��"
                .TextMatrix(0, 10) = "�ۼ�"
                .TextMatrix(0, 11) = "�ɱ���"
                .TextMatrix(0, 12) = "�ϴι���"
                .TextMatrix(0, 13) = "����"
                .TextMatrix(0, 14) = "��������"
                .TextMatrix(0, 15) = "�ϴι�Ӧ��ID"
                .TextMatrix(0, 16) = "��׼�ĺ�"
                
                
            End If
            
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
                .ColAlignment(intCol) = 1
            Next
            .ColWidth(0) = 0
            .ColWidth(15) = 0
            
            .ColAlignment(5) = 7
            .ColAlignment(6) = 7
            .ColAlignment(7) = 7
            .ColAlignment(8) = 7
            .ColAlignment(10) = 7
            .ColAlignment(11) = 7
            .ColAlignment(12) = 7
            
            If mblnStartUp = False Then
                .ColWidth(0) = 0
                .ColWidth(1) = 1200
                .ColWidth(2) = 0
                .ColWidth(3) = 1000
                .ColWidth(4) = 1000
                .ColWidth(5) = 1200
                .ColWidth(6) = 1200
                .ColWidth(7) = 1200
                .ColWidth(9) = 1200
                .ColWidth(10) = 1200
                .ColWidth(13) = 1200
                .ColWidth(14) = 1200
                .ColWidth(16) = 1200
                                
                .Row = 1
                Call RestoreFlexState(Msf����, Me.Name)
                .ColWidth(8) = IIf(mblnCostView = False, 0, 1200)
                .ColWidth(11) = IIf(mblnCostView = False, 0, 1200)
                .ColWidth(12) = IIf(mblnCostView = False, 0, 1200)
            End If
        End With
    End Select
End Sub

Private Sub chkContinue_Click()
    Dim blnState As Boolean

    If vsfѡ��.Rows > 2 And chkContinue.Value = 0 Then
        If MsgBox("�Ѿ���ѡ�����Ĵ��ڣ�ȡ��������ѡ�񡱽������ѡ�������ģ���ȷ����" _
            , vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            vsfѡ��.Rows = 1
            vsfѡ��.Rows = 2
            lblѡ��.Caption = "ѡ��ҩƷ"
        Else
            chkContinue.Value = 1
            Exit Sub
        End If
        
    End If

    picѡ����.Visible = chkContinue.Value = 1
    picSplit02_S.Visible = chkContinue.Value = 1
    Form_Resize
    
    
    If chkContinue.Value = 0 Then
        picѡ����.Tag = "չ��"
        Set picUpDown01.Picture = imgsMain.ListImages(2).Picture
        picSplit02_S.MousePointer = 0
    End If
    
    '�ж�ȷ�ϰ�ť�Ƿ����
    If In_�༭״̬ = 1 Then cmdȷ��.Enabled = True: Exit Sub
    
    blnState = ((mint���� = 3 And mint�ⷿ <> 3) Or (mint���� = 1 And mint�ⷿ = 1) Or (mint���� = 2 And mint�ⷿ = 2)) And Not mrsStock.EOF

    If In_�༭״̬ = 2 And ((mint���� = 3 And mint�ⷿ <> 3) Or (mint���� = 1 And mint�ⷿ = 1) Or (mint���� = 2 And mint�ⷿ = 2)) And mblnPrice Then
        If mbln��ʾ���� = False Then
            cmdȷ��.Enabled = True
        Else
            cmdȷ��.Enabled = blnState
        End If
    Else
        cmdȷ��.Enabled = True
    End If
    
    If chkContinue.Value = 1 Then cmdȷ��.Enabled = True
    
End Sub

Private Sub Cmdȡ��_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdȷ��_Click()
    Dim blnValid As Boolean
    
    If chkContinue.Value = 0 Then '���ɶ�ѡʱ
        If In_�༭״̬ = 2 Then If CheckData = False Then Exit Sub
        
        '�������������������Ƿ�һ��
        If In_�༭״̬ = 2 Then
            blnValid = ���������(mlngԴ�ⷿID, mlngLastSelect����ID)
        Else
            blnValid = ���������(mlngĿ�ⷿID, mlngLastSelect����ID)
        End If
        
        If Not blnValid Then
            ShowMsgBox "���ָ������ڵ�ǰ�ⷿ�еĿ���¼���ڴ��󣨿����ǻ���������" & vbCrLf & "�ô������鵱ǰ�ⷿ�Ĳ������ʼ������ĵķ������ԣ���"
            Exit Sub
        End If
        '��װ��¼��
        If CombinateRec = False Then Exit Sub
        Unload Me
        Exit Sub
    Else '�ɶ�ѡ����
        If CombinateRec = False Then Exit Sub
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    RestoreWinState Me
    mblnStartUp = False
    mblnFirstStart = False
    mblnֻ��ʾ���ٲ��� = False
    'ȡ�ۼ۵�λ
    mstrUnit = ""
    mstrUnitString = ""
    mintStockCheck = 0
    mlngLastSelect����ID = 0
    
    chkContinue.Visible = mbln�Ƿ���� = False
    
    Msf����.Visible = (In_�༭״̬ = 2)
    picѡ����.Visible = False
    picSplit02_S.Visible = False
    picѡ����.Tag = "չ��"
    
    On Error GoTo ErrHandle

    '��ʼ����¼��
    InitRec
    
    If mobjOut Is Nothing Then
        ShowMsgBox "��ָ�������壡"
        Exit Sub
    End If
    
    '��ʼ��������������������
    If LoadTvwData() = False Then Exit Sub
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    If mlngModule = 1725 Then
        mblnProvider = zlStr.IsHavePrivs(mstrPrivs, "�鿴��Ӧ��")
    Else
        mblnProvider = True
    End If
    
    '��ȡ��ǰ�����Ʋ���
    gstrSQL = "Select Nvl(��鷽ʽ,0) ����� From ���ϳ����� Where �ⷿID=[1]"
    Set mrsUnit = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngԴ�ⷿID)
        
    With mrsUnit
        If Not mrsUnit.EOF Then
            mintStockCheck = mrsUnit!�����
        End If
    End With
        
    '���Դ�ⷿ�Ƿ�Ϊ���Ŀ�
    If mlngԴ�ⷿID <> 0 Then
        mint�ⷿ = 3
        
        gstrSQL = "select ����ID from ��������˵�� where (�������� like '���ϲ���' Or �������� like '%�Ƽ���') And ����id=[1]"
        Set mrsUnit = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngԴ�ⷿID)
        
        If mrsUnit.EOF Then
            gstrSQL = "select ����ID from ��������˵�� where �������� In('���Ŀ�','����ⷿ') And ����id=[1]"
            Set mrsUnit = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngԴ�ⷿID)
            If Not mrsUnit.EOF Then mint�ⷿ = 1
        Else
            mint�ⷿ = 2
        End If
    End If
    
    '������ʹ�õĵ�λ����
    If mblnɢװ��λ Then
        mstrUnitString = "/1"
    Else
        mstrUnitString = "/nvl(����ϵ��,1)"
    End If
  
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(IIf(mblnɢװ��λ, 0, 1), g_�ɱ���)
        .FM_��� = GetFmtString(IIf(mblnɢװ��λ, 0, 1), g_���)
        .FM_���ۼ� = GetFmtString(IIf(mblnɢװ��λ, 0, 1), g_�ۼ�)
        .FM_���� = GetFmtString(IIf(mblnɢװ��λ, 0, 1), g_����)
    End With
    With mOraFMT
        .FM_�ɱ��� = GetFmtString(IIf(mblnɢװ��λ, 0, 1), g_�ɱ���, True)
        .FM_��� = GetFmtString(IIf(mblnɢװ��λ, 0, 1), g_���, True)
        .FM_���ۼ� = GetFmtString(IIf(mblnɢװ��λ, 0, 1), g_�ۼ�, True)
        .FM_���� = GetFmtString(IIf(mblnɢװ��λ, 0, 1), g_����, True)
    End With
    
    
    tvwClass_NodeClick tvwClass.SelectedItem
    mblnStartUp = True

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function LoadTvwData() As Boolean
    Dim NodeThis As Node, ItemThis As ListItem
    
    Dim Intĩ�� As Integer
    Dim rs���ʷ��� As New ADODB.Recordset
    
    '������;�����Ƿ�������
    On Error GoTo ErrHandle
    LoadTvwData = False
    
    With mrsData
        
        gstrSQL = "" & _
            "   Select ID,�ϼ�ID,����,1 as ĩ�� " & _
            "   From ���Ʒ���Ŀ¼ where ����=7" & _
            "   Start With �ϼ�ID IS NULL Connect By Prior ID=�ϼ�ID " & _
            "   Order by level,ID"
        
        zlDatabase.OpenRecordset mrsData, gstrSQL, Me.Caption
        
        If .EOF Then
            ShowMsgBox "���ʼ�����ķ��ࣨ����Ŀ¼������"
            Exit Function
        End If
        
        
        '��������;��������װ��
        tvwClass.Nodes.Clear
        tvwClass.Nodes.Add , 4, "Root", "������������", 1, 1
        
        Do While Not .EOF
            
            If IsNull(!�ϼ�ID) Then
                Set NodeThis = tvwClass.Nodes.Add("Root", 4, "K_" & !Id, !����, 2, 2)
            Else
                Set NodeThis = tvwClass.Nodes.Add("K_" & !�ϼ�ID, 4, "K_" & !Id, !����, 2, 2)
            End If
            .MoveNext
        Loop
    End With
    
    With tvwClass
        .Nodes(1).Selected = True
    End With
    
    LoadTvwData = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    mblnFirstStart = True
    If Me.Height < 5835 Then Me.Height = 5835
    If Me.Width < 8415 Then Me.Width = 8415
    
    With ImgLeftRight_S
        .Top = 0
        .Height = Me.ScaleHeight - 200 - Cmdȡ��.Height - .Top
    End With
    
    With tvwClass
        .Top = 0
        .Height = ImgLeftRight_S.Height
        .Width = ImgLeftRight_S.Left
    End With
    
    With ImgUpDown_S
        .Left = ImgLeftRight_S.Left + ImgLeftRight_S.Width
        .Width = Me.ScaleWidth - .Left
    End With
    
    With msf���Ϲ��
        .Left = ImgUpDown_S.Left
        .Top = ImgLeftRight_S.Top + (chkContinue.Height + 2 * chkContinue.Top)
        .Width = ImgUpDown_S.Width
    End With
    
    With Msf����
        If .Visible Then
            .Top = ImgUpDown_S.Top + ImgUpDown_S.Height
            .Height = ImgLeftRight_S.Top + ImgLeftRight_S.Height - .Top
            .Left = msf���Ϲ��.Left
            .Width = msf���Ϲ��.Width
        End If
    End With
    
    With Cmdȡ��
        .Top = tvwClass.Top + tvwClass.Height + 150
        .Left = Me.ScaleWidth - .Width - 150
    End With
    With cmdȷ��
        .Top = Cmdȡ��.Top
        .Left = Cmdȡ��.Left - .Width - 100
    End With
    
    With msf���Ϲ��
        .Height = IIf(Msf����.Visible = False, tvwClass.Top + tvwClass.Height - .Top, Msf����.Top - 45 - .Top)
        If .RowIsVisible(.Row) = False Then .TopRow = .TopRow + IIf(.Row > .TopRow, 1, -1)
    End With
    
    If picSplit02_S.Visible Then
        '���÷ֽ��ߵ�top
        If Msf����.Visible Then '���οɼ�
            Msf����.Height = Msf����.Height - (lblѡ��.Height + picSplit02_S.Height)
            picSplit02_S.Top = Msf����.Top + Msf����.Height
        Else
            msf���Ϲ��.Height = msf���Ϲ��.Height - (lblѡ��.Height + picSplit02_S.Height)
            picSplit02_S.Top = msf���Ϲ��.Top + msf���Ϲ��.Height
        End If
        
        picSplit02_S.Width = msf���Ϲ��.Width
    End If
    
    If picѡ����.Visible Then
        picѡ����.Width = msf���Ϲ��.Width
        picѡ����.Height = lblѡ��.Height
        
        picѡ����.Top = picSplit02_S.Top + picSplit02_S.Height
        
        With lblѡ��
            .Top = 0
            .Left = 0
            .Width = picѡ����.Width
        End With
        With picUpDown01
            .Left = picѡ����.Width - .Width
            .Top = 0
        End With

        If picѡ����.Tag = "����" Then
            picѡ����.Tag = "չ��"
            Set picUpDown01.Picture = imgsMain.ListImages(2).Picture
            picSplit02_S.MousePointer = 0
        End If
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me
    SaveFlexState msf���Ϲ��, Me.Name
    SaveFlexState Msf����, Me.Name
End Sub

Private Sub ImgLeftRight_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    With ImgLeftRight_S
        If .Left + x < 2500 Then Exit Sub
        If .Left + x > Me.ScaleWidth - 4500 Then Exit Sub
        
        .Move .Left + x
    End With
    
    Form_Resize
End Sub

Private Sub ImgUpDown_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    With ImgUpDown_S
        If .Top + y < 2500 Then Exit Sub
        If .Top + y > Me.ScaleHeight - 2500 Then Exit Sub
        
        .Move .Left, .Top + y
    End With
    
    Form_Resize
End Sub
Private Sub Msf����_Click()
    Dim StrHeader As String
    Dim intCol As Integer
    Dim i As Integer
    'ʵ��������
    With Msf����
        If .MouseRow <> 0 Then Exit Sub
        If mrsStock.EOF Then Exit Sub
        
        StrHeader = .TextMatrix(0, .MouseCol)
        Set .DataSource = Nothing
        If Mid(mstrPhysicSortBy, 2) = StrHeader Then
            mstrPhysicSortBy = IIf(Mid(mstrPhysicSortBy, 1, 1) = "A", "D", "A") & .TextMatrix(0, .MouseCol)
            mrsStock.Sort = .TextMatrix(0, .MouseCol) & IIf(Mid(mstrPhysicSortBy, 1, 1) = "A", " Desc", " Asc")
        Else
            mstrPhysicSortBy = "A" & .TextMatrix(0, .MouseCol)
            mrsStock.Sort = .TextMatrix(0, .MouseCol) & " Asc"
        End If
        Set .DataSource = mrsStock

        For intCol = 0 To .Cols - 1
            .ColAlignmentFixed(intCol) = 4
        Next
        
        Call SetFormat(0, False)
        
    End With
    
    With Msf����
        If glngModul = 1717 Then
            If InStr(1, gstrPrivs, "��ʾ�Է����") = 0 Then
                For i = 1 To .Rows - 1
                    If Val(.TextMatrix(i, 6)) > 0 Then
                        .TextMatrix(i, 6) = "��"
                    Else
                        .TextMatrix(i, 6) = "��"
                    End If
                    .TextMatrix(i, 7) = ""
                    .TextMatrix(i, 8) = ""
                Next
            End If
        End If
    End With
End Sub

Private Sub Msf����_DblClick()
    On Error Resume Next
    If cmdȷ��.Enabled = False Then Exit Sub
    
    With mrsStock
        If .RecordCount <> 0 Then .MoveFirst
        If .EOF Then Exit Sub
        If .RecordCount = 0 Then Exit Sub
    End With
    
    If chkContinue.Value = 1 Then
        FillVSFѡ��
        Exit Sub
    End If
    
    Call cmdȷ��_Click
End Sub

Private Sub Msf����_EnterCell()
    Dim intCol As Integer, LngSelectRow As Long
    Dim recGetPrice As New ADODB.Recordset
    Dim lng�շ�ϸĿID As Long
    On Error Resume Next
    
    With Msf����
        .Redraw = False
        
        LngSelectRow = .Row     '���浱ǰѡ����
        If mlngPhysicRow <> 0 Then
            .Row = mlngPhysicRow       '����ϴ�ѡ����
            For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
        
        mlngPhysicRow = LngSelectRow
        .Row = mlngPhysicRow     '���õ�ǰѡ����
        For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
        Next
        .Col = 0
        
        .Redraw = True
    End With
End Sub

Private Sub Msf����_GotFocus()
    With Msf����
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
    End With
End Sub

Private Sub Msf����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Msf����_DblClick
End Sub

Private Sub Msf����_LostFocus()
    With Msf����
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
    End With
End Sub

Private Sub Msf���Ϲ��_Click()
    Dim StrHeader As String
    Dim intCol As Integer
    Dim i As Integer
    
    'ʵ��������
    With msf���Ϲ��
        If .MouseRow <> 0 Then Exit Sub
        If mrsCard.EOF Then Exit Sub
        
        StrHeader = .TextMatrix(0, .MouseCol)
        Set .DataSource = Nothing
        If Mid(mstrCardSortBy, 2) = StrHeader Then
            mstrCardSortBy = IIf(Mid(mstrCardSortBy, 1, 1) = "A", "D", "A") & .TextMatrix(0, .MouseCol)
            mrsCard.Sort = .TextMatrix(0, .MouseCol) & IIf(Mid(mstrCardSortBy, 1, 1) = "A", " Desc", " Asc")
        Else
            mstrCardSortBy = "A" & .TextMatrix(0, .MouseCol)
            mrsCard.Sort = .TextMatrix(0, .MouseCol) & " Asc"
        End If
        Set .DataSource = mrsCard

        For intCol = 0 To .Cols - 1
            .ColAlignmentFixed(intCol) = 4
        Next
        
        Call SetFormat(1, False)
    End With
    
    With msf���Ϲ��
        If glngModul = 1717 Then
            If InStr(1, gstrPrivs, "��ʾ�Է����") = 0 Then
                For i = 1 To .Rows - 1
                    If Val(.TextMatrix(i, mCol.�������)) > 0 Then
                        .TextMatrix(i, mCol.�������) = "��"
                    Else
                        .TextMatrix(i, mCol.�������) = "��"
                    End If
                    .TextMatrix(i, mCol.�����) = ""
                    .TextMatrix(i, mCol.�����) = ""
                Next
            End If
        End If
    End With
End Sub

Private Sub Msf���Ϲ��_DblClick()
    If mrsCard.EOF Then Exit Sub
    If mrsCard.RecordCount = 0 Then Exit Sub
    
    If chkContinue.Value = 1 Then
        FillVSFѡ��
        Exit Sub
    End If
    
    If cmdȷ��.Enabled Then
        cmdȷ��_Click
    Else
        MsgBox "������û�п�棬���ܼ���������", vbInformation, gstrSysName
    End If
End Sub

Private Sub FillVSFѡ��()
    Dim blnEof As Boolean         '�Ƿ�������ο��
    Dim i As Integer
    Dim blnValid    As Boolean
    
    '���ҩƷ�ظ�
    If chkContinue.Value = 1 Then
        For i = 1 To vsfѡ��.Rows - 2
            If Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("����ID"))) = Val(msf���Ϲ��.TextMatrix(msf���Ϲ��.Row, mCol.����ID)) Then
                If Msf����.Visible Then
                    If vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("����")) = Msf����.TextMatrix(Msf����.Row, 2) Then
                        Exit Sub
                    End If
                Else
                    Exit Sub
                End If
            End If
        Next
    End If
    
    If In_�༭״̬ = 2 Then If CheckData = False Then Exit Sub
        
    '�������������������Ƿ�һ��
    If In_�༭״̬ = 2 Then
        blnValid = ���������(mlngԴ�ⷿID, mlngLastSelect����ID)
    Else
        blnValid = ���������(mlngĿ�ⷿID, mlngLastSelect����ID)
    End If
    
    If Not blnValid Then
        ShowMsgBox "���ָ������ڵ�ǰ�ⷿ�еĿ���¼���ڴ��󣨿����ǻ���������" & vbCrLf & "�ô������鵱ǰ�ⷿ�Ĳ������ʼ������ĵķ������ԣ���"
        Exit Sub
    End If
    
    
    With mrsCard
        If .RecordCount <> 0 Then .MoveFirst
        .Find "����ID=" & Val(msf���Ϲ��.TextMatrix(msf���Ϲ��.Row, mCol.����ID))
        
        If .EOF Then
            MsgBox "�����ڲ�����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mbln��ʾ���� = True Then 'ֻ����ʾ���ε�����²���Ҫ�����²���
            If ((mint���� = 3 And mint�ⷿ <> 3) Or (mint���� = 1 And mint�ⷿ = 1) Or (mint���� = 2 And mint�ⷿ = 2)) And In_�༭״̬ = 2 Then
                With mrsStock
                    If .RecordCount <> 0 Then .MoveFirst
                    .Find "����=" & Val(Msf����.TextMatrix(Msf����.Row, 2))
                    If .EOF Then
                        blnEof = True
                        If mblnPrice Then
                            MsgBox "�޿�����ݣ�", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End With
            End If
        End If
    End With
    
    'װ����д���¼��������������ʹ��
    With vsfѡ��
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 2, .ColIndex("����ID")) = mrsCard!����ID
        .TextMatrix(.Rows - 2, .ColIndex("����id")) = mrsCard!����id
        .TextMatrix(.Rows - 2, .ColIndex("����id")) = mrsCard!����id
        
        .TextMatrix(.Rows - 2, .ColIndex("����")) = mrsCard!����
        .TextMatrix(.Rows - 2, .ColIndex("����")) = zlStr.Nvl(mrsCard!ͨ������)
        .TextMatrix(.Rows - 2, .ColIndex("��Ʒ��")) = zlStr.Nvl(mrsCard!��Ʒ��)
        .TextMatrix(.Rows - 2, .ColIndex("���")) = mrsCard!���
        .TextMatrix(.Rows - 2, .ColIndex("����")) = "" & mrsCard!����
        .TextMatrix(.Rows - 2, .ColIndex("�ۼ�")) = zlStr.Nvl(mrsCard!�ۼ�, 0)
        .TextMatrix(.Rows - 2, .ColIndex("ɢװ��λ")) = mrsCard!ɢװ��λ
        .TextMatrix(.Rows - 2, .ColIndex("����ϵ��")) = mrsCard!����ϵ��
        .TextMatrix(.Rows - 2, .ColIndex("��װ��λ")) = mrsCard!��װ��λ
        .TextMatrix(.Rows - 2, .ColIndex("���Ч��")) = "" & mrsCard!��Ч��
        .TextMatrix(.Rows - 2, .ColIndex("���Ч��")) = "" & mrsCard!���Ч��
        .TextMatrix(.Rows - 2, .ColIndex("���ʧЧ��")) = "" & mrsCard!���ʧЧ��
        .TextMatrix(.Rows - 2, .ColIndex("һ���Բ���")) = IIf(mrsCard!һ���Բ��� = "��", 1, 0)
        .TextMatrix(.Rows - 2, .ColIndex("�޾��Բ���")) = IIf(mrsCard!�޾��Բ��� = "��", 1, 0)
        .TextMatrix(.Rows - 2, .ColIndex("�ⷿ����")) = IIf(mrsCard!�ⷿ���� = "��", 1, 0)
        .TextMatrix(.Rows - 2, .ColIndex("���÷���")) = IIf(mrsCard!���÷��� = "��", 1, 0)
        
        .TextMatrix(.Rows - 2, .ColIndex("ʱ��")) = IIf(mrsCard!ʱ�� = "��", 1, 0)
        
        '�����ҷ���
        If In_�༭״̬ = 2 And ((mint���� = 3 And mint�ⷿ <> 3) Or (mint���� = 1 And mint�ⷿ = 1) Or (mint���� = 2 And mint�ⷿ = 2)) Then
            If mbln��ʾ���� = True Then 'ֻ����ʾ��������²���Ҫ�����²���������������
                If Msf����.TextMatrix(Msf����.Row, 3) = "����������������" Then
                    .TextMatrix(.Rows - 2, .ColIndex("����")) = -1
                Else
                    If Not blnEof Then
                        .TextMatrix(.Rows - 2, .ColIndex("����")) = Val(mrsStock!����)
                        .TextMatrix(.Rows - 2, .ColIndex("����")) = "" & mrsStock!����
                        .TextMatrix(.Rows - 2, .ColIndex("Ч��")) = "" & mrsStock!ʧЧ��
                        .TextMatrix(.Rows - 2, .ColIndex("���ʧЧ��")) = "" & mrsStock!���ʧЧ��
                        .TextMatrix(.Rows - 2, .ColIndex("����")) = "" & mrsStock!����
                        .TextMatrix(.Rows - 2, .ColIndex("��������")) = "" & mrsStock!��������
                        .TextMatrix(.Rows - 2, .ColIndex("��׼�ĺ�")) = "" & mrsStock!��׼�ĺ�
                        .TextMatrix(.Rows - 2, .ColIndex("��ҩ��λID")) = "" & mrsStock!�ϴι�Ӧ��id
                        .TextMatrix(.Rows - 2, .ColIndex("��������")) = IIf(IsNull(mrsStock!��������), 0, mrsStock!��������)
                        .TextMatrix(.Rows - 2, .ColIndex("ʵ������")) = IIf(IsNull(mrsStock!�������), 0, mrsStock!�������)
                        .TextMatrix(.Rows - 2, .ColIndex("ʵ�ʽ��")) = IIf(IsNull(mrsStock!�����), 0, mrsStock!�����)
                        .TextMatrix(.Rows - 2, .ColIndex("ʵ�ʲ��")) = IIf(IsNull(mrsStock!�����), 0, mrsStock!�����)
                        If Not mblnStock Then Call Get���ÿ��(.TextMatrix(.Rows - 2, .ColIndex("����ID")), .TextMatrix(.Rows - 2, .ColIndex("����")))
                    End If
                End If
            Else
                If Not mblnStock Then Call Get���ÿ��(mrsCard!����ID, 0)
            End If
        Else
        '���򲻷���
            .TextMatrix(.Rows - 2, .ColIndex("��������")) = IIf(IsNull(mrsCard!��������), 0, mrsCard!��������)
            .TextMatrix(.Rows - 2, .ColIndex("ʵ������")) = IIf(IsNull(mrsCard!�������), 0, mrsCard!�������)
            .TextMatrix(.Rows - 2, .ColIndex("ʵ�ʽ��")) = IIf(IsNull(mrsCard!�����), 0, mrsCard!�����)
            .TextMatrix(.Rows - 2, .ColIndex("ʵ�ʲ��")) = IIf(IsNull(mrsCard!�����), 0, mrsCard!�����)
            If In_�༭״̬ = 1 Then
                .TextMatrix(.Rows - 2, .ColIndex("��׼�ĺ�")) = "" & mrsCard!��׼�ĺ�
            Else
                If mrsStock.RecordCount > 0 Then
                    mrsStock.MoveFirst
                    .TextMatrix(.Rows - 2, .ColIndex("��׼�ĺ�")) = zlStr.Nvl(mrsStock!��׼�ĺ�)
                Else
                    .TextMatrix(.Rows - 2, .ColIndex("��׼�ĺ�")) = ""
                End If
            End If
            
            If Not mblnStock Then Call Get���ÿ��(.TextMatrix(.Rows - 2, .ColIndex("����ID")), 0)
        End If
        
        '�������ʾ�Է��ⷿ�Ŀ�棬��������ȡ������
        If Not mblnStock Then
            .TextMatrix(.Rows - 2, .ColIndex("msin��������")) = msin��������
            .TextMatrix(.Rows - 2, .ColIndex("msinʵ������")) = msinʵ������
            .TextMatrix(.Rows - 2, .ColIndex("msinʵ�ʽ��")) = msinʵ�ʽ��
            .TextMatrix(.Rows - 2, .ColIndex("msinʵ�ʲ��")) = msinʵ�ʲ��
        End If
        .TextMatrix(.Rows - 2, .ColIndex("ָ��������")) = mrsCard!ָ��������
        .TextMatrix(.Rows - 2, .ColIndex("ָ�������")) = mrsCard!ָ�������
    End With
    
    lblѡ��.Caption = "ѡ�����ģ�" & vsfѡ��.Rows - 2 & "����"
End Sub

Private Sub Msf���Ϲ��_EnterCell()
    Dim lng�շ�ϸĿID As Long, intCol As Integer, LngSelectRow As Long
    Dim strTmp As String, recGetPrice As New ADODB.Recordset
    Dim strKc As String
    Dim i As Integer
    
'    On Error Resume Next
    On Error GoTo ErrHandle

    With msf���Ϲ��
        .Redraw = False
        
        LngSelectRow = .Row     '���浱ǰѡ����
        If mlngCardRow <> 0 Then
            If mlngCardRow <= .Rows - 1 Then
                .Row = mlngCardRow       '����ϴ�ѡ����
            End If
            For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H80000005
                .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
        
        mlngCardRow = LngSelectRow
        .Row = mlngCardRow       '���õ�ǰѡ����
        For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
        Next
        .Col = 0
        .Redraw = True
        
        '����ù�����ĵļ۸�ִ��ʱ�仹δִ��,�򴥷�
        lng�շ�ϸĿID = Val(.TextMatrix(.Row, mCol.����ID))
        
        If lng�շ�ϸĿID = 0 Then
            If Msf����.Visible Then
                Msf����.Clear
                Msf����.Rows = 2
                Call SetFormat(0, True)
                Msf����_EnterCell
            Else
                Call SetFormat(0, True)
            End If
            mlngLastSelect����ID = 0
            Exit Sub
        End If
        
        If mlngLastSelect����ID = lng�շ�ϸĿID Then Exit Sub
        mlngLastSelect����ID = lng�շ�ϸĿID
        
        
        '����ѵ�ִ�����ڶ��۸�δִ�У�ִ�м������
        
        gstrSQL = "Select ID From �շѼ�Ŀ " & _
                  "Where �շ�ϸĿID=[1] And �䶯ԭ��=0" & GetPriceClassString("")
        Set recGetPrice = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng�շ�ϸĿID)
        
        With recGetPrice
            If Not recGetPrice.EOF Then
                If Not IsNull(recGetPrice!Id) Then
                    lng�շ�ϸĿID = recGetPrice!Id
                    gstrSQL = "zl_�����շ���¼_Adjust(" & lng�շ�ϸĿID & ")"
                    
                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-�������ϼ۸������¼")
                End If
            End If
        End With
    End With
    
    If In_�༭״̬ = 2 Then
        Msf����.Visible = False
        '���������Ĺ�������е��������ο����Ϣ
        
        mblnʱ�� = (msf���Ϲ��.TextMatrix(msf���Ϲ��.Row, mCol.ʱ��) = "��")
        mint���� = 0
        If msf���Ϲ��.TextMatrix(msf���Ϲ��.Row, mCol.�ⷿ����) = "��" Or msf���Ϲ��.TextMatrix(msf���Ϲ��.Row, mCol.���÷���) = "��" Then
            If msf���Ϲ��.TextMatrix(msf���Ϲ��.Row, mCol.�ⷿ����) = "��" And msf���Ϲ��.TextMatrix(msf���Ϲ��.Row, mCol.���÷���) = "��" Then
                mint���� = 3
            ElseIf msf���Ϲ��.TextMatrix(msf���Ϲ��.Row, mCol.�ⷿ����) = "��" Then
                mint���� = 1
            Else
                mint���� = 2
            End If
        End If
        'mint�ⷿ 1-���Ŀ�;2-���ϲ���;3-�Ƽ���
        'mint���� 0-������;1-�ⷿ����;2-���÷���;3-���Ŀ����÷���
        If Not ((mint���� = 3 And mint�ⷿ <> 3) Or (mint���� = 1 And mint�ⷿ = 1) Or (mint���� = 2 And mint�ⷿ = 2)) Then '��������Ĳ�����
            Msf����.Visible = False
            Form_Resize
        Else
            If Msf����.Visible = False Then
                If mbln��ʾ���� = True Then '�˲��������ܲ�����ʾ�����б������첻��ȷ����ģʽ
                    Msf����.Visible = True
                End If
            Else
                If mbln��ʾ���� = False Then '�˲��������ܲ�����ʾ�����б������첻��ȷ����ģʽ
                    Msf����.Visible = False
                End If
            End If
        End If
        Form_Resize
        
        gstrSQL = ""
        
        If mbln������ Then
            gstrSQL = "" & _
                "   Select " & IIf(mstr�̵�ʱ�� <> "", "/*+ Rule*/", "") & " 1 RID,���� �ⷿ,0 ����,'����������������' ����,sysdate ʧЧ��,to_char(0," & gOraFmt_Max.FM_���� & ") ��������,to_char(0," & gOraFmt_Max.FM_���� & ") �������,to_char(0," & gOraFmt_Max.FM_��� & ") �����" & _
                "           ,to_char(0," & gOraFmt_Max.FM_��� & ") �����,sysdate as ���ʧЧ��,to_char(0," & gOraFmt_Max.FM_���ۼ� & ") �ۼ�,'' As �ɱ���,to_char(0," & gOraFmt_Max.FM_�ɱ��� & ") �ϴι���,'' ���� , Sysdate As ��������, 0 As �ϴι�Ӧ��id,'' ��׼�ĺ�" & _
                "   From ���ű�" & _
                "   Where ID=[1]" & _
                "   Union "
        End If
        
        gstrSQL = gstrSQL & " Select " & IIf(mstr�̵�ʱ�� <> "", "/*+ Rule*/", "") & " 2 RID,P.���� �ⷿ,K.����,K.�ϴ����� ����,K.Ч�� ʧЧ��,"
        
        If mblnStock Then
            If mblnɢװ��λ Then
                strTmp = " to_char(K.��������," & gOraFmt_Max.FM_���� & ") ��������," & _
                         " to_char(K.ʵ������," & gOraFmt_Max.FM_���� & ") as �������,"
            Else
                strTmp = " to_char(K.��������," & gOraFmt_Max.FM_���� & ") ��������," & _
                         " to_char(K.ʵ������," & gOraFmt_Max.FM_���� & ") as �������,"
            End If
        Else
            strTmp = "to_char(0," & gOraFmt_Max.FM_���� & ") ��������,to_char(0," & gOraFmt_Max.FM_���� & ") �������,"
        End If
                 
        'ȡ���
        '20060731:���˺���룬��Ҫ����̵�ʱ��Ŀ��
        strKc = "" & _
            "   SELECT a.�ⷿid, a.ҩƷid, NVL (a.����, 0) AS ����,a.�ϴι�Ӧ��ID, a.�ϴβɹ���," & _
            "           a.ʵ������,a.ʵ�ʽ��, a.ʵ�ʲ��, a.��������,A.���ۼ�,ƽ���ɱ���,a.�ϴ�����,a.�ϴβ���,a.Ч��,a.���Ч��,a.�ϴ���������,a.��׼�ĺ� " & _
            "   FROM ҩƷ��� a " & _
            "   Where a.ҩƷid=[4]" & _
            "       AND a.����=1 " & _
            "       AND a.�ⷿid+0 = "
        If mlngԴ�ⷿID <> 0 Or mlngĿ�ⷿID <> 0 Then
            strKc = strKc & IIf(mlngԴ�ⷿID = 0, "[1]", "[2]")
        End If
        
        If mstr�̵�ʱ�� <> "" Then
            strKc = strKc & _
                "   UNION ALL " & _
                "   SELECT a.�ⷿid, a.ҩƷid, NVL (a.����, 0) AS ����, a.��ҩ��λID �ϴι�Ӧ��ID,max(a.�ɱ���) �ϴβɹ���, " & _
                "           -SUM (DECODE (a.���ϵ��, 1, a.ʵ������*a.����, -a.ʵ������*a.����)) AS ʵ������, " & _
                "           -SUM (DECODE (a.���ϵ��, 1, a.���۽��, -a.���۽��)) AS ʵ�ʽ��," & _
                "           -SUM (DECODE (a.���ϵ��, 1, a.���, -a.���)) AS ʵ�ʲ��, " & _
                "           -SUM (DECODE (a.���ϵ��, 1, a.��д����*a.����, -a.��д����*a.����)) AS ��������, " & _
                "           Max(���ۼ�) as ���ۼ�,0 as ƽ���ɱ���,a.����,a.���� , A.Ч��,a.���Ч��,a.��������,a.��׼�ĺ�" & _
                "   FROM ҩƷ�շ���¼ a " & _
                "   Where  a.ҩƷid+0=[4]  " & _
                "           AND a.�ⷿid + 0 ="
            If mlngԴ�ⷿID <> 0 Or mlngĿ�ⷿID <> 0 Then
                strKc = strKc & IIf(mlngԴ�ⷿID = 0, "[1]", "[2]")
            End If
            strKc = strKc & " AND a.������� >[5] " & _
                " GROUP BY A.�ⷿid, a.ҩƷid,a.��ҩ��λid, A.����, A.����, A.����, A.Ч��, A.���Ч��,a.��������,a.��׼�ĺ� "
        End If
        
        strKc = "" & _
            "   Select �ⷿid,ҩƷid,nvl(����,0) ����,max(�ϴ�����) �ϴ�����,min(���Ч��) as ���ʧЧ��,max(�ϴι�Ӧ��ID) �ϴι�Ӧ��ID, " & _
            "       Sum(nvl(��������,0)) ��������," & _
            "       Sum(ʵ������) ʵ������," & _
            "       Sum(ʵ�ʽ��) ʵ�ʽ��," & _
            "       Sum(ʵ�ʲ��) ʵ�ʲ��," & _
            "       max(�ϴβɹ���) �ϴβɹ���,Max(���ۼ�) as ���ۼ�,max(ƽ���ɱ���) as ƽ���ɱ���, " & _
            "        Min(���Ч��) ���Ч��,Min(Ч��) Ч��,max(�ϴβ���) �ϴβ��� ,max(�ϴ���������) �ϴ���������,max(��׼�ĺ�) as ��׼�ĺ�,1 As ����" & _
            "   From (" & strKc & ")" & _
            "   Group by �ⷿid,ҩƷid,nvl(����,0) "
                 
                 
        '1.ʵ��:�������:����е����ۼ� ,����Ϊʵ�ʽ��/ʵ������*����ϵ��=�ۼ�
        '2.����:�շѼ�Ŀ�е��ּ�=�ۼ�
        '3.�ɱ���:
        '       a.��������ϴι���,�����ϴι���Ϊ׼
        '       b.����������ϴι���,���ԣ������-�����)/�������Ϊ׼.
        
        gstrSQL = gstrSQL & strTmp & _
                 IIf(mblnStock, " to_char(K.ʵ�ʽ��," & gOraFmt_Max.FM_��� & ") as �����,", "to_char(''," & gOraFmt_Max.FM_��� & ") �����,") & _
                 IIf(mblnStock, " to_char(K.ʵ�ʲ��," & gOraFmt_Max.FM_��� & ") as �����", "to_char(''," & gOraFmt_Max.FM_��� & ") �����") & ",K.���Ч�� ���ʧЧ��," & _
                 IIf(mblnStock, "to_char(Decode(nvl(M.�Ƿ���,0),0,G.�ּ�,decode(nvl(K.���ۼ�,0),0,nvl(K.ʵ�ʽ��,0)/decode(K.ʵ������,null,1,0,1,K.ʵ������),K.���ۼ�))" & IIf(mblnɢװ��λ, "", "*nvl(D.����ϵ��,1)") & "," & gOraFmt_Max.FM_���ۼ� & ") �ۼ�,", "to_char(0," & gOraFmt_Max.FM_���ۼ� & ") �ۼ�,") & _
        " to_char(k.ƽ���ɱ���," & gOraFmt_Max.FM_�ɱ��� & ") as �ɱ���, " & _
                 IIf(mblnStock, "to_char(decode(nvl(K.�ϴβɹ���,0),0,(nvl(K.ʵ�ʽ��,0)-nvl(K.ʵ�ʲ��,0))/decode(K.ʵ������,null,1,0,1,K.ʵ������),K.�ϴβɹ���)" & IIf(mblnɢװ��λ, "", "*nvl(D.����ϵ��,1)") & "," & gOraFmt_Max.FM_�ɱ��� & ") �ϴι���", "to_char(0," & gOraFmt_Max.FM_�ɱ��� & ") �ϴι���") & _
        "       ,K.�ϴβ��� ����,k.�ϴ��������� �������� ,k.�ϴι�Ӧ��ID,k.��׼�ĺ� " & _
        " From ���ű� P, �������� D, " & IIf(mstr�̵�ʱ�� <> "", "(" & strKc & ")", " ҩƷ���") & " K,�շ���ĿĿ¼ M,�շѼ�Ŀ G " & _
        " Where     K.�ⷿID = P.ID And D.����ID = K.ҩƷID " & _
        " And K.�ⷿID " & IIf(mstr�̵�ʱ�� <> "", "+0=", "=") & IIf(mlngԴ�ⷿID = 0, "[1]", "[2]") & _
        " And K.ҩƷID " & IIf(mstr�̵�ʱ�� <> "", "+0=", "=") & "[4] And K.����=1 " & _
        " And D.����id=G.�շ�ϸĿID(+) " & _
        " And D.����ID=M.ID And (M.վ��=[6] or M.վ�� is null) " & _
        " And m.Id = g.�շ�ϸĿid And (Sysdate Between g.ִ������ And Nvl(g.��ֹ����, Sysdate)) " & _
        GetPriceClassString("G")
                 
        If mbln�̵㵥 Then
            gstrSQL = gstrSQL & " And (K.ʵ������<>0 Or K.ʵ�ʽ��<>0 Or K.ʵ�ʲ��<>0)"
        Else
            gstrSQL = gstrSQL & " And K.ʵ������<>0 "
        End If
        
'        If mlng��Ӧ��ID <> 0 Then gstrSQL = gstrSQL & " And K.�ϴι�Ӧ��ID=[3]"
        
        If gSystem_Para.P156_�����㷨 = 0 Then
            gstrSQL = gstrSQL & " Order by RID,����"
        Else
            gstrSQL = gstrSQL & " Order by RID,ʧЧ��,����"
        End If
        
        Dim dtDate As Date
        If mstr�̵�ʱ�� <> "" Then
            dtDate = CDate(mstr�̵�ʱ��)
        Else
            dtDate = Now
        End If
        
        Set mrsStock = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngĿ�ⷿID, mlngԴ�ⷿID, mlng��Ӧ��ID, mlngLastSelect����ID, dtDate, gstrNodeNo, gstrPriceClass)
          
        Dim blnState As Boolean
        With Msf����
            If Not mrsStock.EOF Then
                Set .DataSource = mrsStock
                .ColWidth(0) = 0
            Else
                .Clear
                .Rows = 2
            End If
            
            Call SetFormat(0, mrsStock.EOF)
            If mbln������ And mrsStock.RecordCount <> 0 Then
                If .Row > 2 Then
                    .Row = 2
                End If
            End If
            blnState = ((mint���� = 3 And mint�ⷿ <> 3) Or (mint���� = 1 And mint�ⷿ = 1) Or (mint���� = 2 And mint�ⷿ = 2)) And Not mrsStock.EOF
            If mbln��ʾ���� = True And blnState = True Then 'ֻ��������ʾ�����б���ٲ�ѯ����������Ϣ����������
                .Visible = True
            Else
                .Visible = False
            End If

            Msf����_EnterCell
        End With
        Form_Resize
    End If
    
    With Msf����
        If glngModul = 1717 Then
            If InStr(1, gstrPrivs, "��ʾ�Է����") = 0 And Msf����.Visible = True Then
                For i = 1 To .Rows - 1
                    If Val(.TextMatrix(i, 6)) > 0 Then
                        .TextMatrix(i, 6) = "��"
                    Else
                        .TextMatrix(i, 6) = "��"
                    End If
                    .TextMatrix(i, 7) = ""
                    .TextMatrix(i, 8) = ""
                Next
            End If
        End If
    End With
    
    '���ð�ť״̬
    With mrsCard
        If .RecordCount <> 0 Then .MoveFirst
        .Find "����ID=" & Val(msf���Ϲ��.TextMatrix(msf���Ϲ��.Row, mCol.����ID))
        If .EOF Then
            MsgBox "�����ڲ�����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If In_�༭״̬ = 2 And ((mint���� = 3 And mint�ⷿ <> 3) Or (mint���� = 1 And mint�ⷿ = 1) Or (mint���� = 2 And mint�ⷿ = 2)) And mblnPrice Then
            If mbln��ʾ���� = False Then
                cmdȷ��.Enabled = True
            Else
                cmdȷ��.Enabled = blnState
            End If
        Else
            cmdȷ��.Enabled = True
        End If
        If chkContinue.Value = 1 Then cmdȷ��.Enabled = True
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Msf���Ϲ��_GotFocus()
    With msf���Ϲ��
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
    End With
End Sub

Private Sub Msf���Ϲ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Msf���Ϲ��_DblClick
End Sub

Private Sub Msf���Ϲ��_LostFocus()
    With msf���Ϲ��
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
    End With
End Sub

Private Sub picSplit02_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    With picSplit02_S
        If .Top + y < msf���Ϲ��.Top + 1000 Then Exit Sub
        If .Top + y > tvwClass.Height - 1500 Then Exit Sub
        .Move .Left, .Top + y
    End With
    
    picѡ����.Move picѡ����.Left, picѡ����.Top + y, picѡ����.Width, picѡ����.Height - y
    
End Sub

Private Sub picUpDown01_Click()
    If picѡ����.Tag = "չ��" Then
        picѡ����.Tag = "����"
        Set picUpDown01.Picture = imgsMain.ListImages(1).Picture
        picSplit02_S.MousePointer = 7

        
        picSplit02_S.Top = Me.tvwClass.Height / 2
        picѡ����.Top = picSplit02_S.Top + picSplit02_S.Height
        picѡ����.Height = tvwClass.Height - picѡ����.Top
        
    Else
        picѡ����.Tag = "չ��"
        Set picUpDown01.Picture = imgsMain.ListImages(2).Picture
        picSplit02_S.MousePointer = 0
        
        Form_Resize
    End If
End Sub

Private Sub picѡ����_Resize()
    lblѡ��.Width = picѡ����.Width
    
    With vsfѡ��
        .Top = lblѡ��.Height
        .Left = 0
        .Width = lblѡ��.Width
        .Height = picѡ����.Height - lblѡ��.Height
    End With
End Sub


Private Sub tvwClass_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim strTmp As String, StrGroupBy As String
    Dim strKc As String
    Dim rsTmp As ADODB.Recordset
    Dim blnVirtualStock As Boolean
    Dim i As Integer
    Dim int������ȡֵ��ʽ As Integer
    
    '������������;����Ĺ������
    '    ���Ŀ��ⷿ����ȷ����������������ã������Ƽ��ң����������Ĳ���
    '    ���Ŀ��ⷿ����ȷ����������������ã��������Ŀ⡢�Ƽ��ң������Ʒ������
    '    ���Ŀ��ⷿ�Ƿ��������ﲡ�ˣ���������ҩ���Խ��룻
    '    ���Ŀ��ⷿ�Ƿ�����סԺ���ˣ���סԺ��ҩ���Խ��룻
    
    On Error GoTo ErrHandle

    If mlngModule = 1712 Or mlngModule = 1714 Then
        int������ȡֵ��ʽ = Val(zlDatabase.GetPara(268, glngSys))
    End If
    
    If mlngĿ�ⷿID <> 0 Then
        mblnֻ��ʾ���ٲ��� = �ж�ֻ�߱����ϲ���(mlngĿ�ⷿID)
        If mblnֻ��ʾ���ٲ��� = False Then
            mblnֻ��ʾ���ٲ��� = �ж�ֻ�߱����ϲ���(mlngԴ�ⷿID)
        End If
    End If
    
    '�ж�����ⷿ
    gstrSQL = "select count(*) rec from ��������˵�� where ��������='����ⷿ' and ����id=[1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж�����ⷿ", mlngĿ�ⷿID)
    If rsTmp!rec = 1 And (mobjOut.Name = "frmPurchaseCard" Or mobjOut.Name = "frmOtherInputCard") Then
        blnVirtualStock = True
    End If
    
    '����ͷ��˳��
    gstrSQL = "" & _
        " Select " & IIf(mstr�̵�ʱ�� <> "", "/*+ Rule*/", "") & " D.����id,D.����id,D.����ID,D.����,D.ͨ������,D.��Ʒ��,D.���,D.����,d.��׼�ĺ�,d.ע��֤��,x.���� As �ϴι�Ӧ�� ,to_char(D.�ۼ�," & gOraFmt_Max.FM_���ۼ� & ") �ۼ�,to_char(d.�ɱ���," & gOraFmt_Max.FM_�ɱ��� & ") as ���³ɱ���,D.ɢװ��λ,D.����ϵ��,D.��װ��λ," & _
          IIf(mblnStock, "  to_char(S.�������� " & IIf(mblnɢװ��λ, "", "/D.����ϵ��") & "," & gOraFmt_Max.FM_���� & ") ��������, to_char(S.������� " & _
          IIf(mblnɢװ��λ, "", "/D.����ϵ��") & "," & gOraFmt_Max.FM_���� & ") ������� ,to_char(S.�����," & gOraFmt_Max.FM_��� & ") �����,to_char(S.�����," & gOraFmt_Max.FM_��� & ") �����,", "to_char(''," & gOraFmt_Max.FM_���� & ") ��������,to_char(''," & gOraFmt_Max.FM_���� & ") �������,to_char(''," & gOraFmt_Max.FM_��� & ") �����,to_char(''," & gOraFmt_Max.FM_��� & ") �����,") & _
        "     D.���Ч�� ��Ч��,D.���Ч��,S.���ʧЧ��,D.һ���Բ���,D.�޾��Բ���,D.�ⷿ����,D.���÷���,D.ʱ��,to_char(D.ָ��������," & gOraFmt_Max.FM_���ۼ� & ") ָ��������,D.ָ�������,E.�ⷿ��λ" & _
        " From "

    '������Ϣ������Ŀ¼
    If mblnֻ��ʾ���ٲ��� Then
        gstrSQL = gstrSQL & _
                "     (Select Distinct u.����id,u.����id,H.����ID,V.����,V.���� As ͨ������,B.���� As ��Ʒ��,V.���," & IIf(int������ȡֵ��ʽ = 0, "decode(u.�ϴβ���,null,v.����,u.�ϴβ���)", "decode(v.����,null,u.�ϴβ���,v.����)") & " as ����,u.��׼�ĺ�,u.ע��֤��,V.���㵥λ as ɢװ��λ,U.��װ��λ," & _
                "          To_Char(U.����ϵ��," & GFM_XS & " ) ����ϵ��,nvl(To_Char(U.���Ч��,'9999990'),0) ���Ч��,nvl(To_Char(U.���Ч��,'9999990'),0) ���Ч��," & _
                "          Decode(U.�ⷿ����,1,'��','��') �ⷿ����,Decode(U.���÷���,1,'��','��') ���÷���,Decode(U.һ���Բ���,1,'��','��')  һ���Բ���,Decode(U.�޾��Բ���,1,'��','��') �޾��Բ���,Decode(V.�Ƿ���,1,'��','��') ʱ��," & _
                "          U.ָ�������� ,To_Char(U.ָ�������," & GFM_CJL & " ) ָ�������,�ּ� as �ۼ�,u.�ɱ���,Nvl(u.�ϴι�Ӧ��id, 0) As �ϴι�Ӧ��id " & _
                "      From �������� U,�շ���ĿĿ¼ V,������ĿĿ¼ H," & _
                "       (SELECT �շ�ϸĿid, ִ�п���id FROM �շ�ִ�п��� WHERE ִ�п���ID" & IIf(mlngԴ�ⷿID <> 0, "+0=[1]", IIf(mlngĿ�ⷿID <> 0, "+0=[2]", " Is Not NULL")) & ") K," & _
                "       (Select �շ�ϸĿID, ִ�п���ID From �շ�ִ�п��� Where ִ�п���ID" & IIf(mlngĿ�ⷿID <> 0, "+0=[2]", IIf(mlngԴ�ⷿID <> 0, "+0=[1]", " Is Not NULL")) & " ) i," & _
                "       �շ���Ŀ���� B, �շѼ�Ŀ P " & _
                "      Where U.����id=v.id And (v.վ��=[5] or v.վ�� is null) And U.����id=H.id  And V.ID = B.�շ�ϸĿid(+) And B.����(+) = 3 " & _
                "          AND U.����id=K.�շ�ϸĿID " & IIf(mbln���޴洢�ⷿ����, "(+)", "") & _
                "          AND U.����id=i.�շ�ϸĿID " & IIf(mbln���޴洢�ⷿ����, "(+)", "") & _
                           IIf(mblnֻ��ʾ���ٲ���, " and  U.�������� =1 ", IIf(mblnTrackUsing = True, " and  U.�������� =0 ", "")) & " And v.Id = p.�շ�ϸĿid And (Sysdate Between p.ִ������ And Nvl(p.��ֹ����, Sysdate)) " & _
                           GetPriceClassString("P")
    Else
        gstrSQL = gstrSQL & _
                "     (Select Distinct u.����id,u.����id,H.����ID,V.����,V.���� As ͨ������,B.���� As ��Ʒ��,V.���," & IIf(int������ȡֵ��ʽ = 0, "decode(u.�ϴβ���,null,v.����,u.�ϴβ���)", "decode(v.����,null,u.�ϴβ���,v.����)") & " as ����,u.��׼�ĺ�,u.ע��֤��,V.���㵥λ as ɢװ��λ,U.��װ��λ," & _
                "          To_Char(U.����ϵ��," & GFM_XS & " ) ����ϵ��,nvl(To_Char(U.���Ч��,'9999990'),0) ���Ч��,nvl(To_Char(U.���Ч��,'9999990'),0) ���Ч��," & _
                "          Decode(U.�ⷿ����,1,'��','��') �ⷿ����,Decode(U.���÷���,1,'��','��') ���÷���,Decode(U.һ���Բ���,1,'��','��')  һ���Բ���,Decode(U.�޾��Բ���,1,'��','��') �޾��Բ���,Decode(V.�Ƿ���,1,'��','��') ʱ��," & _
                "          U.ָ��������,To_Char(U.ָ�������," & GFM_CJL & " ) ָ�������,�ּ� �ۼ�,u.�ɱ���,Nvl(u.�ϴι�Ӧ��id, 0) As �ϴι�Ӧ��id " & _
                "      From �������� U,�շ���ĿĿ¼ V,������ĿĿ¼ H," & _
                "     (Select ִ�п���ID,�շ�ϸĿID From �շ�ִ�п��� Where ִ�п���ID" & IIf(mlngԴ�ⷿID <> 0, "=[1]", IIf(mlngĿ�ⷿID <> 0, "=[2]", " Is Not NULL")) & " ) K," & _
                "     (Select ִ�п���ID,�շ�ϸĿID From �շ�ִ�п��� Where ִ�п���ID" & IIf(mlngĿ�ⷿID <> 0, "+0=[2]", IIf(mlngԴ�ⷿID <> 0, "+0=[1]", " Is Not NULL")) & " ) i," & _
                "      �շ���Ŀ���� B, �շѼ�Ŀ P  " & _
                "      Where U.����id=v.id And (v.վ��=[5] or v.վ�� is null) And U.����id=H.id  And V.ID = B.�շ�ϸĿid(+) And B.����(+) = 3 " & _
                "          AND U.����id=K.�շ�ϸĿID " & IIf(mbln���޴洢�ⷿ����, "(+)", "") & _
                "          AND U.����id=i.�շ�ϸĿID " & IIf(mbln���޴洢�ⷿ����, "(+)", "") & _
                           IIf(mblnֻ��ʾ���ٲ���, " and  U.�������� =1 ", IIf(mblnTrackUsing = True, " and  U.�������� =0 ", "")) & " And v.Id = p.�շ�ϸĿid And (Sysdate Between p.ִ������ And Nvl(p.��ֹ����, Sysdate)) " & _
                           GetPriceClassString("P")
    End If
    
    If mlngĿ�ⷿID > 0 Then
        gstrSQL = gstrSQL & " And" & _
            "     ( exists(select 1 from ��������˵�� where �������� In('�Ƽ���','���Ŀ�','���ϲ���','����ⷿ')  and ����id=[2])  " & _
            "       or v.�������=(select distinct '1' from ��������˵�� where �������� like '���ϲ���' and ����id=[2] and ������� in(1,3))" & _
            "       or v.�������=(select distinct '2' from ��������˵�� where �������� like '���ϲ���' and ����id=[2] and ������� in(2,3)))"
    End If
    
    '����ָ��������;����Ĺ�����
    If Not (Node.Key Like "Root") Then
        gstrSQL = gstrSQL & " And H.����ID IN (Select ID from ���Ʒ���Ŀ¼ where ����=7 Start With ID=" & Mid(Node.Key, 3) & " Connect By Prior ID=�ϼ�ID)"
    Else
        gstrSQL = gstrSQL & " "
    End If
    
    'ֻ����δͣ�õĹ������
    If mstr�̵�ʱ�� <> "" Then      '���̵�ʱ����˵������̵�ʱ��С��ͣ�õ�ʱ��ҲӦ����ʾ����
        gstrSQL = gstrSQL & " And (V.����ʱ�� Is Null Or V.����ʱ��>[4])"
    Else
        gstrSQL = gstrSQL & " And (V.����ʱ�� Is Null Or To_char(V.����ʱ��,'yyyy-MM-dd')='3000-01-01')"
    End If
    
    'ֻ����ָ�����ʷ���Ĺ������
    gstrSQL = gstrSQL & IIf(blnVirtualStock, " and nvl(u.��ֵ����,0)=1 and nvl(u.���ٲ���,0)=1 and nvl(u.��������,0)=1 and nvl(u.���÷���,0)=1", "") & " ) D,"

    'ȡ���
    '20060731:���˺���룬��Ҫ����̵�ʱ��Ŀ��
    strKc = "" & _
        "   SELECT a.�ⷿid, a.ҩƷid, NVL (a.����, 0) AS ����,a.�ϴι�Ӧ��ID," & _
        "           a.ʵ������,a.ʵ�ʽ��, a.ʵ�ʲ��, a.��������,a.�ϴ�����,a.�ϴβ���,a.Ч��,a.���Ч�� " & _
        "   FROM ҩƷ��� a " & _
        "   Where a.����=1 AND a.�ⷿid = "
    If mlngԴ�ⷿID <> 0 Or mlngĿ�ⷿID <> 0 Then
        strKc = strKc & IIf(mlngԴ�ⷿID = 0, "[2]", "[1]")
    End If
    
    '�̵�ʱ�����̵�ʱ������̵�ʱ�䵽��ǰʱ��ķ�����
    If mstr�̵�ʱ�� <> "" Then
        strKc = strKc & _
            "   UNION ALL " & _
            "   SELECT a.�ⷿid, a.ҩƷid, NVL (a.����, 0) AS ����, a.��ҩ��λID �ϴι�Ӧ��ID, " & _
            "           -SUM (DECODE (a.���ϵ��, 1, a.ʵ������*a.����, -a.ʵ������*a.����)) AS ʵ������, " & _
            "           -SUM (DECODE (a.���ϵ��, 1, a.���۽��, -a.���۽��)) AS ʵ�ʽ��," & _
            "           -SUM (DECODE (a.���ϵ��, 1, a.���, -a.���)) AS ʵ�ʲ��,-SUM (DECODE (a.���ϵ��, 1, a.ʵ������*a.����, -a.ʵ������*a.����)) AS ��������,a.����,a.���� , A.Ч��,a.���Ч��" & _
            "   FROM ҩƷ�շ���¼ a " & _
            "   Where a.�ⷿid + 0 ="
        If mlngԴ�ⷿID <> 0 Or mlngĿ�ⷿID <> 0 Then
            strKc = strKc & IIf(mlngԴ�ⷿID = 0, "[2]", "[1]")
        End If
        strKc = strKc & " AND a.������� >[4] " & _
            " GROUP BY A.�ⷿid, a.ҩƷid,a.��ҩ��λid, A.����, A.����, A.����, A.Ч��, A.���Ч�� "
    End If
    
    If mblnStock Then
        gstrSQL = gstrSQL & " (Select ҩƷid as ����id,min(���Ч��) as ���ʧЧ�� , Sum(nvl(��������,0)) ��������," & _
                " Sum(nvl(ʵ������,0)) �������," & _
                " Sum(nvl(ʵ�ʽ��,0)) �����," & _
                " Sum(nvl(ʵ�ʲ��,0)) �����"
    Else
        gstrSQL = gstrSQL & " (Select ҩƷid as ����id,min(���Ч��) as ���ʧЧ��, 0 ��������," & _
                " 0 �������,0 �����,0 �����"
    End If
    If mstr�̵�ʱ�� <> "" Then
         gstrSQL = gstrSQL & " From (" & strKc & ") where 1=1 "
    Else
         gstrSQL = gstrSQL & " From ҩƷ��� Where ����=1 "
    End If
    
    
    'If mlng��Ӧ��ID <> 0 Then gstrSQL = gstrSQL & " And �ϴι�Ӧ��ID=[3]"
    
    If mlngԴ�ⷿID <> 0 Or mlngĿ�ⷿID <> 0 Then
        gstrSQL = gstrSQL & " And �ⷿID" & IIf(mstr�̵�ʱ�� <> "", "+0=", "=") & IIf(mlngԴ�ⷿID = 0, "[2]", "[1]") & "  Group By ҩƷid) S"
    Else
        gstrSQL = gstrSQL & " Group By ҩƷid) S"
    End If
    
    gstrSQL = gstrSQL & ",(Select ����id,�ⷿID,�ⷿ��λ From ���ϴ����޶�" & _
              " Where �ⷿID=" & IIf(mintEditState = 2, "[1]", "[2]") & ") E,��Ӧ�� X"
    
    '������
    gstrSQL = gstrSQL & " Where D.����ID=S.����ID"
    
    If mbln����ʾ������� And mblnStock Then
        gstrSQL = gstrSQL & " And S.��������<>0"
    Else
        '��ϵͳ���������ϳ������顱Ϊ�����ֹʱ��������Ϊ��
        If Not (mintStockCheck = 2 And In_�༭״̬ = 2) Or mbln�̵㵥 Or Not mblnCheck Then gstrSQL = gstrSQL & "(+) "
        'If In_�༭״̬ = 2 Then gstrSQL = gstrSQL & " And S.��������<>0"
    End If
    
    gstrSQL = gstrSQL & " And D.����ID=E.����ID(+) And d.�ϴι�Ӧ��id = x.Id(+) Order By D.����"
    Dim dtDate As Date
    If mstr�̵�ʱ�� <> "" Then
        dtDate = CDate(mstr�̵�ʱ��)
    Else
        dtDate = Now
    End If
    Set mrsCard = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngԴ�ⷿID, mlngĿ�ⷿID, mlng��Ӧ��ID, dtDate, gstrNodeNo)
    
    With msf���Ϲ��
        If Not mrsCard.EOF Then
            Set .DataSource = mrsCard
        Else
            .Clear
            .Rows = 2
        End If
        Call SetFormat(1, mrsCard.EOF)
    End With
    cmdȷ��.Enabled = (mrsCard.EOF <> True)
    
    Call Msf���Ϲ��_EnterCell
    
    With msf���Ϲ��
        If glngModul = 1717 Then
            If InStr(1, gstrPrivs, "��ʾ�Է����") = 0 Then
                For i = 1 To .Rows - 1
                    If Val(.TextMatrix(i, mCol.�������)) > 0 Then
                        .TextMatrix(i, mCol.�������) = "��"
                    Else
                        .TextMatrix(i, mCol.�������) = "��"
                    End If
                    .TextMatrix(i, mCol.�����) = ""
                    .TextMatrix(i, mCol.�����) = ""
                Next
            End If
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Function InitRec()
    '----------------------------------------------------------------------------------------
    '����:�����������ݼ��ṹ
    '----------------------------------------------------------------------------------------
        Set mrsReturn = New ADODB.Recordset
        With mrsReturn
            If .State = 1 Then .Close
            .Fields.Append "����ID", adDouble, 18, adFldIsNullable
            .Fields.Append "����ID", adDouble, 18, adFldIsNullable
            .Fields.Append "����ID", adDouble, 18, adFldIsNullable
            .Fields.Append "��ҩ��λID", adDouble, 18, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 80, adFldIsNullable
            .Fields.Append "��Ʒ��", adLongVarChar, 80, adFldIsNullable
            .Fields.Append "���", adLongVarChar, 82, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 40, adFldIsNullable
            .Fields.Append "�ۼ�", adDouble, 18, adFldIsNullable
            .Fields.Append "ɢװ��λ", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "����ϵ��", adDouble, 11, adFldIsNullable
            .Fields.Append "��װ��λ", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "���Ч��", adDouble, 5, adFldIsNullable
            .Fields.Append "���Ч��", adDouble, 5, adFldIsNullable
            .Fields.Append "һ���Բ���", adDouble, 2, adFldIsNullable
            .Fields.Append "�޾��Բ���", adDouble, 2, adFldIsNullable
            .Fields.Append "�ⷿ����", adDouble, 2, adFldIsNullable
            .Fields.Append "���÷���", adDouble, 2, adFldIsNullable
            .Fields.Append "��׼�ĺ�", adLongVarChar, 50, adFldIsNullable
            
            .Fields.Append "ʱ��", adDouble, 2, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "����", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "Ч��", adDate, , adFldIsNullable
            .Fields.Append "���ʧЧ��", adDate, , adFldIsNullable
            .Fields.Append "��������", adDate, , adFldIsNullable
            .Fields.Append "��������", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "ʵ������", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "ʵ�ʽ��", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "ʵ�ʲ��", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "ָ��������", adDouble, 11, adFldIsNullable
            .Fields.Append "ָ�������", adDouble, 11, adFldIsNullable
            
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
End Function

Private Function CombinateRec() As Boolean
    '��װ��¼��
    '��λ��¼��
    Dim blnEof As Boolean               '�Ƿ�������ο��
    Dim i As Integer
    
    CombinateRec = False
    
    On Error GoTo ErrHandle:
    
    If chkContinue.Value = 0 Then '��װһ������
        With mrsCard
            If .RecordCount <> 0 Then .MoveFirst
            .Find "����ID=" & Val(msf���Ϲ��.TextMatrix(msf���Ϲ��.Row, mCol.����ID))
            
            If .EOF Then
                MsgBox "�����ڲ�����", vbInformation, gstrSysName
                Exit Function
            End If
            
            If mbln��ʾ���� = True Then 'ֻ����ʾ���ε�����²���Ҫ�����²���
                If ((mint���� = 3 And mint�ⷿ <> 3) Or (mint���� = 1 And mint�ⷿ = 1) Or (mint���� = 2 And mint�ⷿ = 2)) And In_�༭״̬ = 2 Then
                    With mrsStock
                        If .RecordCount <> 0 Then .MoveFirst
                        .Find "����=" & Val(Msf����.TextMatrix(Msf����.Row, 2))
                        If .EOF Then
                            blnEof = True
                            If mblnPrice Then
                                MsgBox "�޿�����ݣ�", vbInformation, gstrSysName
                                Exit Function
                            End If
                        End If
                    End With
                End If
            End If
        End With
        
        'װ����д���¼��������������ʹ��
        With mrsReturn
            If .EOF Then .AddNew
            !����ID = mrsCard!����ID
            !����id = mrsCard!����id
            !����id = mrsCard!����id
            
            !���� = mrsCard!����
            !���� = zlStr.Nvl(mrsCard!ͨ������)
            !��Ʒ�� = zlStr.Nvl(mrsCard!��Ʒ��)
            !��� = mrsCard!���
            !���� = mrsCard!����
            !�ۼ� = zlStr.Nvl(mrsCard!�ۼ�, 0)
            !ɢװ��λ = mrsCard!ɢװ��λ
            !����ϵ�� = mrsCard!����ϵ��
            !��װ��λ = mrsCard!��װ��λ
            !���Ч�� = mrsCard!��Ч��
            !���Ч�� = mrsCard!���Ч��
            !���ʧЧ�� = mrsCard!���ʧЧ��
            !һ���Բ��� = IIf(mrsCard!һ���Բ��� = "��", 1, 0)
            !�޾��Բ��� = IIf(mrsCard!�޾��Բ��� = "��", 1, 0)
            !�ⷿ���� = IIf(mrsCard!�ⷿ���� = "��", 1, 0)
            !���÷��� = IIf(mrsCard!���÷��� = "��", 1, 0)
            
            !ʱ�� = IIf(mrsCard!ʱ�� = "��", 1, 0)
            
            '�����ҷ���
            If In_�༭״̬ = 2 And ((mint���� = 3 And mint�ⷿ <> 3) Or (mint���� = 1 And mint�ⷿ = 1) Or (mint���� = 2 And mint�ⷿ = 2)) Then
                If mbln��ʾ���� = True Then 'ֻ����ʾ��������²���Ҫ�����²���������������
                    If Msf����.TextMatrix(Msf����.Row, 3) = "����������������" Then
                        !���� = -1
                    Else
                        If Not blnEof Then
                            !���� = Val(mrsStock!����)
                            !���� = mrsStock!����
                            !Ч�� = mrsStock!ʧЧ��
                            !���ʧЧ�� = mrsStock!���ʧЧ��
                            !���� = mrsStock!����
                            !�������� = mrsStock!��������
                            !��׼�ĺ� = mrsStock!��׼�ĺ�
                            !��ҩ��λID = mrsStock!�ϴι�Ӧ��id
                            !�������� = IIf(IsNull(mrsStock!��������), 0, mrsStock!��������)
                            !ʵ������ = IIf(IsNull(mrsStock!�������), 0, mrsStock!�������)
                            !ʵ�ʽ�� = IIf(IsNull(mrsStock!�����), 0, mrsStock!�����)
                            !ʵ�ʲ�� = IIf(IsNull(mrsStock!�����), 0, mrsStock!�����)
                            If Not mblnStock Then Call Get���ÿ��(!����ID, !����)
                        End If
                    End If
                Else
                    If Not mblnStock Then Call Get���ÿ��(mrsCard!����ID, 0)
                End If
            Else
            '���򲻷���
                !�������� = IIf(IsNull(mrsCard!��������), 0, mrsCard!��������)
                !ʵ������ = IIf(IsNull(mrsCard!�������), 0, mrsCard!�������)
                !ʵ�ʽ�� = IIf(IsNull(mrsCard!�����), 0, mrsCard!�����)
                !ʵ�ʲ�� = IIf(IsNull(mrsCard!�����), 0, mrsCard!�����)
                If In_�༭״̬ = 1 Then
                    !��׼�ĺ� = mrsCard!��׼�ĺ�
                Else
                    If mrsStock.RecordCount > 0 Then
                        mrsStock.MoveFirst
                        !��׼�ĺ� = zlStr.Nvl(mrsStock!��׼�ĺ�)
                    Else
                        !��׼�ĺ� = ""
                    End If
                End If
                
                If Not mblnStock Then Call Get���ÿ��(!����ID, 0)
            End If
            
            '�������ʾ�Է��ⷿ�Ŀ�棬��������ȡ������
            If Not mblnStock Then
                !�������� = msin��������
                !ʵ������ = msinʵ������
                !ʵ�ʽ�� = msinʵ�ʽ��
                !ʵ�ʲ�� = msinʵ�ʲ��
            End If
            !ָ�������� = mrsCard!ָ��������
            !ָ������� = mrsCard!ָ�������
            .Update
        End With
    Else '��װ��������
        With mrsReturn
            For i = 1 To vsfѡ��.Rows - 2
                .AddNew
                !����ID = Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("����ID")))
                !����id = Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("����id")))
                !����id = Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("����id")))
                
                !���� = vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("����"))
                !���� = vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("����"))
                !��Ʒ�� = vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("��Ʒ��"))
                !��� = vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("���"))
                !���� = vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("����"))
                !�ۼ� = Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("�ۼ�")))
                !ɢװ��λ = vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("ɢװ��λ"))
                !����ϵ�� = Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("����ϵ��")))
                !��װ��λ = vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("��װ��λ"))
                !���Ч�� = IIf(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("���Ч��")) = "", Null, vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("���Ч��")))
                !���Ч�� = Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("���Ч��")))
                !���ʧЧ�� = IIf(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("���ʧЧ��")) = "", Null, vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("���ʧЧ��")))
                !һ���Բ��� = Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("һ���Բ���")))
                !�޾��Բ��� = vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("�޾��Բ���"))
                !�ⷿ���� = Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("�ⷿ����")))
                !���÷��� = Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("���÷���")))
                !ʱ�� = Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("ʱ��")))
                !���� = Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("����")))
                !���� = vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("����"))
                !Ч�� = IIf(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("Ч��")) = "", Null, vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("Ч��")))
                !�������� = IIf(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("��������")) = "", Null, vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("��������")))
                !��׼�ĺ� = vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("��׼�ĺ�"))
                !��ҩ��λID = IIf(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("��ҩ��λID")) = "", Null, vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("��ҩ��λID")))
                !�������� = IIf(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("��������")) = "", Null, vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("��������")))
                !ʵ������ = IIf(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("ʵ������")) = "", Null, vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("ʵ������")))
                !ʵ�ʽ�� = IIf(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("ʵ�ʽ��")) = "", Null, vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("ʵ�ʽ��")))
                !ʵ�ʲ�� = IIf(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("ʵ�ʲ��")) = "", Null, vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("ʵ�ʲ��")))
                !ָ�������� = Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("ָ��������")))
                !ָ������� = Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("ָ�������")))
                
                .Update
            Next
        End With
    End If
    
    CombinateRec = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckData() As Boolean
    Dim DblCurStock As Double       '��ǰ�����
    Dim intCol As Integer
    '����Ƿ�����ѡ��
    CheckData = False
    
    If cmdȷ��.Enabled = False Then Exit Function
    
    If mbln��ʾ���� = False Then
        CheckData = True
        Exit Function '����ǲ���ʾ����ģʽ��ֱ�Ӳ������
    End If
    
    If Msf����.Visible Then
        'lng��Ӧ��ID��Ϊ�㣬��ʾ�˻����޿��ʱ��׼����
        If mlng��Ӧ��ID <> 0 Then
            intCol = GetCol(Msf����, "�ϴι�Ӧ��ID")
            If intCol < 0 Then Exit Function
            If Val(Msf����.TextMatrix(Msf����.Row, intCol)) <> 0 And mlng��Ӧ��ID <> Val(Msf����.TextMatrix(Msf����.Row, intCol)) Then
                MsgBox "��ѡ����˻��̲��Ǹ��������ϵĹ�Ӧ�̣����ܼ���������", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If mblnStock Then
            DblCurStock = Val(Msf����.TextMatrix(Msf����.Row, 5))
        Else
            DblCurStock = Get���ÿ��(Val(msf���Ϲ��.TextMatrix(msf���Ϲ��.Row, mCol.����ID)), Val(Msf����.TextMatrix(Msf����.Row, 2)))
        End If
    Else
        If Not mrsCard.EOF Then
            If mblnStock Then
                DblCurStock = Val(msf���Ϲ��.TextMatrix(msf���Ϲ��.Row, mCol.��������))
            Else
                DblCurStock = Get���ÿ��(Val(msf���Ϲ��.TextMatrix(msf���Ϲ��.Row, mCol.����ID)))
            End If
        End If
    End If
    
    If DblCurStock > 0 Then
        CheckData = True
        Exit Function
    End If
    
    '���Դ�ⷿ��Ŀ�ⷿΪ�գ�������ǲ���Ŀ¼�Լ��ڽ��г������ã����ж�
    If (mlngԴ�ⷿID = 0 And mlngĿ�ⷿID = 0) Then
        CheckData = True
        Exit Function
    End If
    
    '������̵㵥���ò���ѡ�����������жϣ�ֱ���˳�
    If mbln�̵㵥 Then
        CheckData = True
        Exit Function
    End If
    If Msf����.Visible Or mblnʱ�� Then
        If (DblCurStock <> 0) Or Not mblnPrice Or Msf����.TextMatrix(Msf����.Row, 3) = "����������������" Then CheckData = True: Exit Function
        MsgBox "��" & IIf(mblnʱ��, "ʱ��", "����") & "�����Ѿ�û�п�棬���ܼ���������", vbInformation, gstrSysName
        Exit Function
    Else
        If mblnCheck = False Then
           CheckData = True
           Exit Function
        End If
    End If
    
    'mlng��Ӧ��ID��Ϊ�㣬��ʾ�˻����޿��ʱ��׼����
    If mlng��Ӧ��ID <> 0 Then
        MsgBox "�������Ѿ�û�п�棬���ܼ���������", vbInformation, gstrSysName
        Exit Function
    End If
    
    Select Case mintStockCheck
    Case 1
        If MsgBox("�������Ѿ�û�п�棬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Case 2
        MsgBox "�������Ѿ�û�п�棬���ܼ���������", vbInformation, gstrSysName
        Exit Function
    End Select
    CheckData = True
End Function

Public Function ShowMe(ByVal frmMain As Form, ByVal �༭ģʽ As Integer, Optional ByVal Դ�ⷿ As Long, _
                    Optional ByVal Ŀ�ⷿ As Long = 0, Optional ByVal ʹ�ò��� As Long = 0, Optional ByVal Bln����� As Boolean = True, _
                    Optional ByVal bln������λ�ʱ�� As Boolean = True, Optional ByVal mbln�̵㵥�� As Boolean = False, Optional ByVal bln���ӿ����� As Boolean = False, _
                    Optional ByVal bln��ʾ��� As Boolean = True, Optional ByVal lng��Ӧ�� As Long = 0, Optional ByVal blnɢװ��λ As Boolean = True, _
                    Optional blnֻ��ʾ���ٲ��� As Boolean = False, _
                    Optional str�̵�ʱ�� As String = "", _
                    Optional bln����ʾ������� As Boolean = False, _
                    Optional lngModule As Long = 0, _
                    Optional ByVal bln���޴洢�ⷿ���� As Boolean = False, _
                    Optional ByVal strPrivs As String = "", _
                    Optional ByVal bln��ʾ���� As Boolean = True, Optional bln�Ƿ���� As Boolean = True) As ADODB.Recordset
    '------------------------------------------------------------------------------------------------------
    '����:��ʾѡ����
    '����:  bln�����-�����������ļ�ʱ���������治׼����ԭ�򣬿�ǿ������not (���� or ʱ��) ���ĳ���
    '       bln������λ�ʱ��:����������������ļ�ʱ�����ĳ���
    '       mlng��Ӧ��ID:��Ϊ���ʾ�˻�
    '����:��ѡ��ļ�¼��
    '------------------------------------------------------------------------------------------------------
    
    mblnɢװ��λ = blnɢװ��λ
    mlngModule = lngModule
    If lngModule = 1717 Then    '1717:��������
        mblnTrackUsing = IIf(Val(zlDatabase.GetPara("��������", glngSys, lngModule, "0")) = 1, True, False)
    Else
        mblnTrackUsing = False
    End If
    
    With Me
        .In_�༭״̬ = �༭ģʽ
        .In_Դ�ⷿ = Դ�ⷿ
        .In_Ŀ�ⷿ = Ŀ�ⷿ
        .In_���� = ʹ�ò���
        .In_MainFrm = frmMain
        mbln�̵㵥 = mbln�̵㵥��
        mbln������ = bln���ӿ�����
        mblnCheck = Bln�����
        mblnPrice = bln������λ�ʱ��
        mblnStock = bln��ʾ���
        mlng��Ӧ��ID = lng��Ӧ��
        mblnֻ��ʾ���ٲ��� = blnֻ��ʾ���ٲ���
        mstr�̵�ʱ�� = str�̵�ʱ��
        '�޸�:���˺�   Bug:12792    ����:2008-05-08 15:03:47
        mbln����ʾ������� = bln����ʾ�������
        mbln���޴洢�ⷿ���� = bln���޴洢�ⷿ����
        mstrPrivs = strPrivs
        mbln��ʾ���� = bln��ʾ����
        mbln�Ƿ���� = bln�Ƿ����
        .Show 1, frmMain
    End With
    Set ShowMe = mrsReturn.Clone
End Function

Public Function Get���ÿ��(ByVal lng����ID As Long, Optional ByVal lng���� As Long = 0) As Single
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "" & _
        " Select Sum(A.��������" & mstrUnitString & ") ��������,Sum(A.ʵ������" & mstrUnitString & ") ʵ������,sum(A.ʵ�ʽ��) ʵ�ʽ��,sum(A.ʵ�ʲ��) ʵ�ʲ�� " & _
              " From ҩƷ��� A,�������� B " & _
              " Where A.ҩƷID=B.����ID and A.����=1  And A.ҩƷID=[1]" & IIf(lng���� = 0, "", " And Nvl(A.����,0)=[2]")
    
    If mlngԴ�ⷿID <> 0 Or mlngĿ�ⷿID <> 0 Then
        gstrSQL = gstrSQL & " And A.�ⷿID=" & IIf(mlngԴ�ⷿID = 0, "[4]", "[3]")
    End If
    
    gstrSQL = gstrSQL & " Group By A.ҩƷid"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID, lng����, mlngԴ�ⷿID, mlngĿ�ⷿID)
    
    msin�������� = 0
    msinʵ�ʲ�� = 0
    msinʵ�ʽ�� = 0
    msinʵ������ = 0
    If Not rsTemp.EOF Then
        msin�������� = IIf(IsNull(rsTemp!��������), 0, rsTemp!��������)
        msinʵ�ʲ�� = IIf(IsNull(rsTemp!ʵ�ʲ��), 0, rsTemp!ʵ�ʲ��)
        msinʵ�ʽ�� = IIf(IsNull(rsTemp!ʵ�ʽ��), 0, rsTemp!ʵ�ʽ��)
        msinʵ������ = IIf(IsNull(rsTemp!ʵ������), 0, rsTemp!ʵ������)
    End If
    Get���ÿ�� = msin��������
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsfѡ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If vsfѡ��.Rows > 2 Then
            If vsfѡ��.Row <> vsfѡ��.Rows - 1 Then vsfѡ��.RemoveItem vsfѡ��.Row
            If vsfѡ��.Rows = 2 Then
                lblѡ��.Caption = "ѡ������"
            Else
                lblѡ��.Caption = "ѡ�����ģ�" & vsfѡ��.Rows - 2 & "����"
            End If
        End If
    End If
End Sub
