VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmStuffRequestPlanCard 
   Caption         =   "�����깺���༭"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmStuffRequestPlanCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11400
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   13
      Top             =   5137
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   12
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   11
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   8820
      TabIndex        =   9
      Top             =   5085
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   10140
      TabIndex        =   10
      Top             =   5085
      Width           =   1100
   End
   Begin VB.PictureBox Pic���� 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   -15
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   14
      Top             =   45
      Width           =   11715
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   9510
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   570
         Width           =   1710
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   1290
         TabIndex        =   1
         Text            =   "cboStock"
         Top             =   570
         Width           =   2055
      End
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   9930
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   180
         Width           =   1425
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   6
         Top             =   950
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   4948
         Appearance      =   0
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Active          =   -1  'True
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483634
      End
      Begin VB.TextBox txtժҪ 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   8
         Top             =   4080
         Width           =   10410
      End
      Begin VB.ComboBox cboEnterStock 
         Height          =   300
         ItemData        =   "frmStuffRequestPlanCard.frx":014A
         Left            =   4890
         List            =   "frmStuffRequestPlanCard.frx":014C
         TabIndex        =   3
         Text            =   "cboEnterStock"
         Top             =   570
         Width           =   2115
      End
      Begin VB.Label LblEnterStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���깺�ⷿ(&I)"
         Height          =   180
         Left            =   3510
         TabIndex        =   2
         Top             =   630
         Width           =   1170
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�깺����(&S)"
         Height          =   180
         Left            =   210
         TabIndex        =   0
         Top             =   630
         Width           =   990
      End
      Begin VB.Label txt�ƻ����� 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1080
         TabIndex        =   27
         Top             =   660
         Width           =   1845
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "���ϼƣ�"
         Height          =   180
         Left            =   240
         TabIndex        =   26
         Top             =   3840
         Width           =   900
      End
      Begin VB.Label Txt����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7950
         TabIndex        =   24
         Top             =   4440
         Width           =   1005
      End
      Begin VB.Label Txt������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10050
         TabIndex        =   23
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt�������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   22
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt������ 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   21
         Top             =   4440
         Width           =   1005
      End
      Begin VB.Label LblNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NO."
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   9480
         TabIndex        =   20
         Top             =   195
         Width           =   480
      End
      Begin VB.Label lblժҪ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ժҪ(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   7
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "�����깺��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   30
         TabIndex        =   19
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label Lbl�ƻ����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ƻ�����:"
         Height          =   180
         Left            =   8550
         TabIndex        =   4
         Top             =   630
         Width           =   810
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   300
         TabIndex        =   18
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl�������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   2160
         TabIndex        =   17
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label Lbl����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   7365
         TabIndex        =   16
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   9240
         TabIndex        =   15
         Top             =   4500
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList imghot 
      Left            =   840
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":014E
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":0368
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":0582
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":079C
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":09B6
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":0BD0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":0DEA
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":1004
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgcold 
      Left            =   120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":121E
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":1438
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":1652
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":186C
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":1A86
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":1CA0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":1EBA
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRequestPlanCard.frx":20D4
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   28
      Top             =   6615
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmStuffRequestPlanCard.frx":22EE
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13758
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmStuffRequestPlanCard.frx":2B82
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmStuffRequestPlanCard.frx":3084
            Key             =   "WB"
            Object.ToolTipText     =   "���(F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf��Ӧ��ѡ�� 
      Height          =   2565
      Left            =   5850
      TabIndex        =   29
      Top             =   1890
      Visible         =   0   'False
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   4524
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblCode 
      Caption         =   "����"
      Height          =   255
      Left            =   3240
      TabIndex        =   25
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmStuffRequestPlanCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mint�༭״̬ As Integer             '1.������2���޸ģ�3�����գ�4���鿴��5
Private mstr���ݺ� As String                '����ĵ��ݺ�;
Private mint��¼״̬ As Integer             '1:������¼;2-������¼;3-�Ѿ�������ԭ��¼
Private mblnSuccess As Boolean              'ֻҪ��һ�ųɹ�����ΪTrue������ΪFalse
Private mblnFirst As Boolean                '��һ����ʾ
Private mblnSave As Boolean                 '�Ƿ���̺����   TURE���ɹ���
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '�Ƿ�����޸�
Private mblnChange As Boolean               '�Ƿ���й��༭
Private mintErrMsg As Integer       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼
Private mintUnit As Integer            '0-ɢװ��λ,1-��װ��λ
Private mbln���� As Boolean                 '����ȡ���ڴ������޵�ҩƷ
Private mint���� As Integer
Private mint���� As Integer

Private mlng�ƻ�ID As Long
Private mlng�ⷿid As Long
Private mint�ƻ����� As Integer
Private mint���Ʒ��� As Integer
Private mstr������ID As String      '��id�ָ�
Private mbln�б굥λ As Boolean '�����б깩����,Ҫ��mstr������λһ��������.
Private mstr�ڼ�  As String                  '������λ��ʾ,������λ��ʾ,������λ��ʾ
Dim mstrPrivs As String                     'Ȩ��
Private Const mlngModule = 1725
Private mstrLike As String
Private mblnCostView As Boolean                 '�鿴�ɱ��� true-����鿴 false-������鿴
Private mblnProvider As Boolean                 '�鿴�ϴι�Ӧ�������Ϣ true-����鿴 false-������鿴
Private Const mstrCaption As String = "�����깺���༭"
Private mstr�ظ����� As String '��¼�ظ�������

'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------


'=========================================================================================
Private Enum mHeadCol
    ��� = 1
    ���� = 2
    ��� = 3
    ���� = 4
    ��λ = 5
    ����ϵ�� = 6
    �б���� = 7
    �빺���� = 8
    �ƻ����� = 9
    ���� = 10
    ��� = 11
    �ϴι�Ӧ�� = 12
End Enum

Private Const mconIntColS  As Integer = 13     '������

'=========================================================================================

Public Sub ShowCard(frmMain As Form, ByVal str���ݺ� As String, _
        ByVal int�༭״̬ As Integer, ByVal strPrivs As String, Optional blnSuccess As Boolean = False)
    '----------------------------------------------------------------------------------------------------------------
    '����:�깺�ƻ��༭���
    '����:frmMain-���õĸ�����
    '     str���ݺ�-���ݺ�
    '     int�༭״̬-1.������2���޸ģ�3�����գ�4���鿴��5
    '     strPrivs-Ȩ�޴�
    '     blnSuccess-�༭�ɹ�,����true,���򷵻�False
    '----------------------------------------------------------------------------------------------------------------
    mblnSave = False
    mblnSuccess = False
    mstr���ݺ� = str���ݺ�
    mint�༭״̬ = int�༭״̬
    mintErrMsg = 1
    mstrPrivs = strPrivs

    mblnSuccess = blnSuccess
    mblnChange = False
    mblnFirst = True

    Set mfrmMain = frmMain
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    mblnProvider = zlStr.IsHavePrivs(mstrPrivs, "�鿴��Ӧ��")
    
    If Not GetDepend(mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)) Then Exit Sub
    
    If mint�༭״̬ = 1 Then
        mblnEdit = True
    ElseIf mint�༭״̬ = 2 Then
        mblnEdit = True
    ElseIf mint�༭״̬ = 3 Then
        mblnEdit = False
        CmdSave.Caption = "���(&V)"
    ElseIf mint�༭״̬ = 4 Then
        mblnEdit = False
        CmdSave.Caption = "��ӡ(&P)"
        If InStr(mstrPrivs, "���ݴ�ӡ") = 0 Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
    End If
    
    Me.Show vbModal, frmMain
    blnSuccess = mblnSuccess
    str���ݺ� = mstr���ݺ�
End Sub

Private Sub cboEnterStock_Change()
    mblnChange = True
End Sub

Private Sub cboEnterStock_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strTmp As String, strվ������ As String
    Dim vRect As RECT, blnCancel As Boolean
    Dim intIdx As Integer
    
    On Local Error Resume Next
    
    If Visible Then mblnChange = True
    strվ������ = GetDeptStationNode(cboStock.ItemData(cboStock.ListIndex))
    
    If cboEnterStock.ItemData(cboEnterStock.ListIndex) = -1 And Visible Then
        strSQL = "" & _
            "   SELECT DISTINCT a.id,a.����,a.����||'-'||a.����  as ����" & _
            "   FROM ��������˵�� c, �������ʷ��� b, ���ű� a " & _
            "   Where c.�������� = b.���� " & _
            IIf(strվ������ <> "", " And (a.վ�� = [1] or a.վ�� is null) ", "") & _
            "     And b.���� In('V','K') " & _
            "     AND a.id = c.����id " & _
            "     AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'"
        vRect = zlControl.GetControlRect(cboEnterStock.hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "���깺�ⷿ", False, "", "", False, False, True, _
                    vRect.Left, vRect.Top, cboEnterStock.Height, blnCancel, False, True, strվ������)
        If Not rsTmp Is Nothing Then
            intIdx = cbo.FindIndex(cboEnterStock, rsTmp!Id)
            If intIdx <> -1 Then
                cboEnterStock.ListIndex = intIdx
'            Else
'                cboEnterStock.AddItem rsTmp!����, cboEnterStock.ListCount - 1
'                cboEnterStock.ItemData(cboEnterStock.NewIndex) = rsTmp!Id
'                cboEnterStock.ListIndex = cboEnterStock.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "û�б��깺�ⷿ���ݣ����ȵ����Ź��������á�", vbInformation, gstrSysName
            End If

            intIdx = cbo.FindIndex(cboEnterStock, cboEnterStock.Tag)
            Call cbo.SetIndex(cboEnterStock.hwnd, intIdx)
        End If
    Else
        cboEnterStock.Tag = cboEnterStock.Text
    End If
End Sub

Private Sub cboEnterStock_GotFocus()
    Call zlControl.TxtSelAll(cboEnterStock)
End Sub

Private Sub cboEnterStock_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If cboEnterStock.Style = 2 And cboEnterStock.ListIndex <> -1 Then
            cboEnterStock.ListIndex = -1
        End If
    End If
End Sub

Private Sub cboEnterStock_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call OS.PressKey(vbKeyTab)
    ElseIf KeyAscii >= 32 Then
        If Not cboEnterStock.Locked And cboEnterStock.Style = 2 Then
            lngIdx = cbo.MatchIndex(cboEnterStock.hwnd, KeyAscii)
            If lngIdx = -1 And cboEnterStock.ListCount > 0 Then lngIdx = 0
            cboEnterStock.ListIndex = lngIdx
        End If
    End If
End Sub

Private Sub cboEnterStock_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String, strվ������ As String
    Dim vRect As RECT, blnCancel As Boolean
        
    If cboEnterStock.ListIndex <> -1 Then Exit Sub '��ѡ��
    If cboEnterStock.Text = "" Then cboEnterStock.Tag = "": Exit Sub '������
    
    strվ������ = GetDeptStationNode(cboStock.ItemData(cboStock.ListIndex))
    strInput = UCase(NeedName(cboEnterStock.Text))
    strSQL = " SELECT DISTINCT a.id,a.����,a.����||'-'||a.����  as ����" & _
            "  FROM ��������˵�� c, �������ʷ��� b, ���ű� a " & _
            "  Where c.�������� = b.���� " & IIf(strվ������ <> "", " And (a.վ�� = [3] or a.վ�� is null) ", "") & _
            "    And b.���� In('V','K') " & _
            "    AND a.id = c.����id " & _
            "    AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" & _
            "    And (Upper(a.����) Like [1] Or Upper(a.����) Like [2] Or Upper(a.����) Like [2]) " & _
            "  Order by ���� "
    
    On Error GoTo errH
    vRect = zlControl.GetControlRect(cboEnterStock.hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "���깺�ⷿ", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cboEnterStock.Height, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%", strվ������)
    If Not rsTmp Is Nothing Then
        intIdx = cbo.FindIndex(cboEnterStock, rsTmp!Id)
        If intIdx <> -1 Then
            cboEnterStock.ListIndex = intIdx
        Else
            cboEnterStock.AddItem rsTmp!����, cboEnterStock.ListCount - 1
            cboEnterStock.ItemData(cboEnterStock.NewIndex) = rsTmp!Id
            cboEnterStock.ListIndex = cboEnterStock.NewIndex
        End If
        mlng�ⷿid = cboEnterStock.ItemData(cboEnterStock.ListIndex)
    Else
        If Not blnCancel Then
            MsgBox "δ�ҵ���Ӧ�ı��깺�ⷿ��", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboStock_Change()
    mblnChange = True
End Sub

Private Sub cboStock_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strTmp As String, strInput As String
    Dim vRect As RECT, blnCancel As Boolean
    Dim intIdx As Integer
    
    On Local Error Resume Next
    
    If Visible Then mblnChange = True
    
    strInput = UCase(NeedName(cboStock.Text))
    
    If cboStock.ItemData(cboStock.ListIndex) = -1 And Visible Then
        strSQL = "" & _
            "   SELECT DISTINCT a.id,a.����,a.����||'-'||a.���� as ����" & _
            "   FROM ���ű� a  " & _
            "   where (a.����ʱ�� is null or TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01') And (a.վ��=[2] or a.վ�� is null) " & _
            IIf(InStr(1, mstrPrivs, ";���в���;") > 0, "", " and  id in (Select ����id from ������Ա where ��Աid =[1])") & _
            "   Order by ����"
        vRect = zlControl.GetControlRect(cboStock.hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "�깺����", False, "", "", False, False, True, _
                    vRect.Left, vRect.Top, cboStock.Height, blnCancel, False, True, strInput, gstrNodeNo)
        If Not rsTmp Is Nothing Then
            intIdx = cbo.FindIndex(cboStock, rsTmp!Id)
            If intIdx <> -1 Then
                cboStock.ListIndex = intIdx
            Else
                cboStock.AddItem rsTmp!����, cboStock.ListCount - 1
                cboStock.ItemData(cboStock.NewIndex) = rsTmp!Id
                cboStock.ListIndex = cboStock.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "û���깺�������ݣ����ȵ����Ź��������á�", vbInformation, gstrSysName
            End If

            intIdx = cbo.FindIndex(cboStock, cboStock.Tag)
            Call cbo.SetIndex(cboStock.hwnd, intIdx)
        End If
    Else
        cboStock.Tag = cboStock.Text
        'ˢ��cboEnterStock
        SetEnterStock cboStock.ItemData(cboStock.ListIndex)
    End If
End Sub

Private Sub cboStock_GotFocus()
    Call zlControl.TxtSelAll(cboStock)
End Sub

Private Sub cboStock_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If cboStock.Style = 2 And cboStock.ListIndex <> -1 Then
            cboStock.ListIndex = -1
        End If
    End If
End Sub

Private Sub cboStock_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call OS.PressKey(vbKeyTab)
    ElseIf KeyAscii >= 32 Then
        If Not cboStock.Locked And cboStock.Style = 2 Then
            lngIdx = cbo.MatchIndex(cboStock.hwnd, KeyAscii)
            If lngIdx = -1 And cboStock.ListCount > 0 Then lngIdx = 0
            cboStock.ListIndex = lngIdx
        End If
    End If
End Sub

Private Sub cboStock_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
        
    If cboStock.ListIndex <> -1 Then Exit Sub '��ѡ��
    If cboStock.Text = "" Then cboStock.Tag = "": Exit Sub '������
    
    strInput = UCase(NeedName(cboStock.Text))
    
    strSQL = "" & _
        "   SELECT DISTINCT a.id,a.����,a.����||'-'||a.���� as ����" & _
        "   FROM ���ű� a  " & _
        "   where (a.����ʱ�� is null or TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01') And (a.վ��=[3] or a.վ�� is null) " & _
        IIf(InStr(1, mstrPrivs, ";���в���;") > 0, "", " and  id in (Select ����id from ������Ա where ��Աid =[1])") & _
        " And (Upper(����) Like [1] Or Upper(����) Like [2] Or Upper(����) Like [2]) " & _
        " Order by ����"
    
    On Error GoTo errH
    vRect = zlControl.GetControlRect(cboStock.hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "�깺����", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cboStock.Height, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%", gstrNodeNo)
    If Not rsTmp Is Nothing Then
        intIdx = cbo.FindIndex(cboStock, rsTmp!Id)
        If intIdx <> -1 Then
            cboStock.ListIndex = intIdx
        Else
            cboStock.AddItem rsTmp!����, cboStock.ListCount - 1
            cboStock.ItemData(cboStock.NewIndex) = rsTmp!Id
            cboStock.ListIndex = cboStock.NewIndex
        End If
    Else
        If Not blnCancel Then
            MsgBox "δ�ҵ���Ӧ�Ĳ���Ա��", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo����_Change()
    mblnChange = True
End Sub

Private Sub cbo����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'����
Private Sub cmdFind_Click()
    If lblCode.Visible = False Then
        lblCode.Visible = True
        txtCode.Visible = True
        txtCode.SetFocus
    Else
        FindData mshBill, mHeadCol.����, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            cmdFind_Click
        End If
    ElseIf KeyCode = vbKeyF3 Then
        FindData mshBill, mHeadCol.����, txtCode.Text, False
    ElseIf KeyCode = vbKeyEscape Then
        If Msf��Ӧ��ѡ��.Visible Then
            Msf��Ӧ��ѡ��.ZOrder 1
            Msf��Ӧ��ѡ��.Visible = False
            Exit Sub
        End If
        Call cmdCancel_Click
    ElseIf KeyCode = vbKeyF7 Then
        If stbThis.Panels("PY").Bevel = sbrRaised Then
            Logogram stbThis, 0
        Else
            Logogram stbThis, 1
        End If
    End If
End Sub

Private Sub CmdSave_Click()
    Dim blnSuccess As Boolean
    
    If mint�༭״̬ = 4 Then    '�鿴
        '��ӡ
        Call FrmBillPrint.ShowMe(Me, glngSys, "zl1_bill_1725", 0, mintUnit, 1725, "�����깺��", txtNO.Tag)
        '�˳�
        Unload Me
        Exit Sub
    End If

    If mint�༭״̬ = 3 Then        '���
        If SaveCheck = True Then
            If IIf(Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModule, "0")) = 1, 1, 0) = 1 Then
                '��ӡ
                If InStr(mstrPrivs, "���ݴ�ӡ") <> 0 Then
                    ReportOpen gcnOracle, glngSys, "zl1_bill_1725", Me, "���ݱ��=" & txtNO.Tag, "��λ=" & mintUnit, 2
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If

    If ValidData = False Then Exit Sub
    blnSuccess = SaveCard

    If blnSuccess = True Then

        If IIf(Val(zlDatabase.GetPara("���̴�ӡ", glngSys, mlngModule, "0")) = 1, 1, 0) = 1 Then
            '��ӡ
            If InStr(mstrPrivs, "���ݴ�ӡ") <> 0 Then
                ReportOpen gcnOracle, glngSys, "zl1_bill_1725", Me, "���ݱ��=" & txtNO.Tag, "��λ=" & mintUnit, 2
            End If
        End If
        If mint�༭״̬ = 2 Then   '�޸�
            Unload Me
            Exit Sub
        End If
    Else
        Exit Sub
    End If

    mblnSave = False
    mblnEdit = True
    mshBill.ClearBill
    txtժҪ.Text = ""
    mblnChange = False
    If txtNO.Tag <> "" Then Me.stbThis.Panels(2).Text = "��һ�ŵ��ݵ�NO�ţ�" & txtNO.Tag
End Sub

Private Sub Form_Activate()
    Dim intMonth As Integer
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    If mint�༭״̬ = 1 Then
        If cboEnterStock.Enabled Then cboEnterStock.SetFocus
        If cboStock.Enabled Then cboStock.SetFocus
    Else
'        mblnChange = False
        Select Case mintErrMsg
            Case 1
                '����
            Case 2
                '�����ѱ�ɾ��
                MsgBox "�õ����ѱ�ɾ�������飡", vbOKOnly, gstrSysName
                Unload Me
                Exit Sub
            Case 3
                '�޸ĵĵ����ѱ����
                MsgBox "�õ����ѱ���������ˣ����飡", vbOKOnly, gstrSysName
                Unload Me
                Exit Sub
        End Select
    End If
    '��ʼ�����뷽ʽ
    If (mint�༭״̬ = 1 Or mint�༭״̬ = 2) And gbytSimpleCodeTrans = 1 Then
        stbThis.Panels("PY").Visible = True
        stbThis.Panels("WB").Visible = True
        gSystem_Para.int���뷽ʽ = Val(zlDatabase.GetPara("���뷽ʽ", , , 0))    'Ĭ��ƴ������
        Logogram stbThis, gSystem_Para.int���뷽ʽ
    Else
        stbThis.Panels("PY").Visible = False
        stbThis.Panels("WB").Visible = False
    End If
End Sub
 

Private Sub Form_Load()
    mFMT.FM_��� = GetDigit

    Me.cboStock.Enabled = True
    mintUnit = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngModule, "0"))
    mstrLike = IIf(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
    
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
    End With
    
    txtNO = mstr���ݺ�
    txtNO.Tag = txtNO
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    initCard
    RestoreWinState Me, App.ProductName, mstrCaption
    With mshBill
        .ColWidth(mHeadCol.����) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(mHeadCol.���) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(mHeadCol.�ϴι�Ӧ��) = IIf(mblnProvider = True, 1000, 0)
    End With
End Sub
Private Sub init����()
    With cbo���� '
        .Clear
        .AddItem "�¶ȼƻ�"
        .AddItem "���ȼƻ�"
        .AddItem "��ȼƻ�"
        .ListIndex = 0
    End With
End Sub
Private Sub initCard()
    Dim i As Integer
    Dim rsInitCard As New Recordset
    Dim strUnit As String
    Dim strStock As String
    Dim intRow As Integer
    Dim intRecordCount As Integer
    Dim str��λ As String
    Dim strOrder As String, strCompare As String
    Dim blnNO�ⷿ As Boolean
    
    On Error GoTo ErrHandle
    strOrder = zlDatabase.GetPara("��������", glngSys, mlngModule, "00")
    strCompare = Mid(strOrder, 1, 1)
    
    Call init����
    
    If mint�༭״̬ <> 4 Then
        With mfrmMain.cboStock
            cboStock.Clear
            If InStr(1, gstrPrivs, ";���в���;") > 0 Then
                For i = 1 To .ListCount - 1
                    If .List(i) <> "���в���" Then
                        cboStock.AddItem .List(i)
                        cboStock.ItemData(cboStock.NewIndex) = .ItemData(i)
                    End If
                Next
                mintcboIndex = .ListIndex - 1
                cboStock.ListIndex = .ListIndex - 1
            Else
                For i = 0 To .ListCount - 1
                    cboStock.AddItem .List(i)
                    cboStock.ItemData(cboStock.NewIndex) = .ItemData(i)
                Next
                mintcboIndex = .ListIndex
                cboStock.ListIndex = .ListIndex
            End If
       End With
    End If
   
'   strStock = " And b.���� In('V','K','W','12') "
'
'   gstrSQL = "SELECT DISTINCT a.id,a.����||'-'||a.����  as ���� " & _
'             "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " & _
'             "Where c.�������� = b.���� " & _
'             IIf(strվ������ <> "", " And a.վ�� = [2] ", "") & _
'             strStock & _
'             "  AND a.id = c.����id " & _
'             "  AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' "
'
'    Set rsInitCard = zldatabase.OpenSQLRecord(gstrSQL, mstrCaption, UserInfo.Id, strվ������)
'
'    With cboEnterStock
'        .Clear
'        Do While Not rsInitCard.EOF
'            .AddItem NVL(rsInitCard!����)
'            .ItemData(.NewIndex) = Val(NVL(rsInitCard!Id))
'            rsInitCard.MoveNext
'        Loop
'    End With
    '��ʼ��cboEnterStock�ؼ�
    SetEnterStock mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)

    '�ⷿ
    Select Case mint�༭״̬
        Case 1
            Txt������ = gstrUserName
            Txt�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        Case 2, 3, 4
            strUnit = "��װ��λ"
            Select Case mintUnit
            Case 0
                str��λ = ",j.���㵥λ ��λ,1 ����ϵ��"
            Case Else
                str��λ = ",m.��װ��λ ��λ,m.����ϵ�� ����ϵ��"
            End Select
            
            initGrid
            
            gstrSQL = "" & _
                "   Select a.�ⷿid,a.����ID,b.����||'-'||b.���� as �ⷿ,c.����||'-'||c.���� as ����" & _
                "   From  ���ϲɹ��ƻ� a,���ű� b,���ű� C " & _
                "   where a.�ⷿid=b.id(+) and a.����id=c.id(+) and  a.����=1 and  a.NO=[1] and rownum=1 "
            Set rsInitCard = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr���ݺ�)
            If rsInitCard.EOF Then
                mintErrMsg = 2
                Exit Sub
            End If
            mlng�ⷿid = Val(zlStr.Nvl(rsInitCard!����ID))
            With cboStock
                blnNO�ⷿ = True
                For i = 0 To .ListCount - 1
                    If .ItemData(i) = mlng�ⷿid Then
                        blnNO�ⷿ = False
                        .ListIndex = i: Exit For
                    End If
                Next
                If blnNO�ⷿ Then
                    If mlng�ⷿid <> 0 Then
                        .AddItem zlStr.Nvl(rsInitCard!����)
                        .ListIndex = .NewIndex
                    Else
                        .ListIndex = 0
                    End If
                End If
            End With
            
            
            mlng�ⷿid = Val(zlStr.Nvl(rsInitCard!�ⷿID))
            With cboEnterStock
                blnNO�ⷿ = True
                For i = 0 To .ListCount - 1
                    If .ItemData(i) = mlng�ⷿid Then
                        blnNO�ⷿ = False
                        .ListIndex = i: Exit For
                    End If
                Next
                If blnNO�ⷿ Then
                    If mlng�ⷿid <> 0 Then
                        .AddItem zlStr.Nvl(rsInitCard!�ⷿ)
                        .ListIndex = .NewIndex
                    Else
                        .ListIndex = 0
                    End If
                End If
            End With
            
            gstrSQL = "" & _
                "   SELECT a.id,nvl(a.�ⷿid,0) as �ⷿid,nvl(c.����,'ȫԺ') AS �ⷿ,a.no, a.�ƻ�����,a.�ڼ�, a.���Ʒ���, a.������," & _
                "           TO_CHAR (a.��������, 'yyyy-mm-dd HH24:MI:SS') AS ��������, a.�����," & _
                "           TO_CHAR (a.�������, 'yyyy-mm-dd HH24:MI:SS') AS �������,a.����˵��," & _
                "           b.���,b.����id ҩƷid,m.�б����,J.����,J.���� ͨ������, J.���" & str��λ & _
                "          ,b.�빺����,b.�ƻ�����, b.����, b.���, b.�ϴι�Ӧ��,b.�ϴ������� " & _
                "   FROM ���ϲɹ��ƻ� a, ���ϼƻ����� b,���ű� c,�������� M,�շ���ĿĿ¼ J " & _
                "   Where a.id = b.�ƻ�id and nvl(a.�ⷿid,0)=c.id(+) " & _
                "          and b.����id=m.����id and m.����id=J.id and nvl(a.����,0)=1 AND a.no = [1]" & _
                "   Order by " & IIf(strCompare = "0", "���", IIf(strCompare = "1", "����", "ͨ������")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            Set rsInitCard = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr���ݺ�)
            
            If rsInitCard.EOF Then
                mintErrMsg = 2
                Exit Sub
            End If

            intRecordCount = rsInitCard.RecordCount

            Txt������ = rsInitCard!������
            If mint�༭״̬ = 2 Then
                Txt������ = gstrUserName
            End If
            Txt�������� = Format(rsInitCard!��������, "yyyy-mm-dd hh:mm:ss")

            Txt����� = IIf(IsNull(rsInitCard!�����), "", rsInitCard!�����)
            Txt������� = IIf(IsNull(rsInitCard!�������), "", Format(rsInitCard!�������, "yyyy-mm-dd hh:mm:ss"))
            txtժҪ.Text = IIf(IsNull(rsInitCard!����˵��), "", rsInitCard!����˵��)
            mint�ƻ����� = rsInitCard!�ƻ�����
            mint���Ʒ��� = rsInitCard!���Ʒ���
            mlng�ⷿid = rsInitCard!�ⷿID
            mlng�ƻ�ID = rsInitCard!Id
            
            mstr�ڼ� = zlStr.Nvl(rsInitCard!�ڼ�)
            If mint�ƻ����� >= 1 And mint�ƻ����� <= 3 Then
                cbo����.ListIndex = mint�ƻ����� - 1
            End If
            
            If (mint�༭״̬ = 2 Or mint�༭״̬ = 3) And Txt����� <> "" Then
                mintErrMsg = 3
                Exit Sub
            End If

            With mshBill
                For intRow = 1 To intRecordCount

                    .TextMatrix(intRow, 0) = rsInitCard!ҩƷid
                    .TextMatrix(intRow, mHeadCol.����) = "[" & rsInitCard!���� & "]" & rsInitCard!ͨ������
                    .TextMatrix(intRow, mHeadCol.���) = IIf(IsNull(rsInitCard!���), "", rsInitCard!���)
                    .TextMatrix(intRow, mHeadCol.�ϴι�Ӧ��) = IIf(IsNull(rsInitCard!�ϴι�Ӧ��), "", rsInitCard!�ϴι�Ӧ��)
                    .TextMatrix(intRow, mHeadCol.����) = IIf(IsNull(rsInitCard!�ϴ�������), "", rsInitCard!�ϴ�������)
                    .TextMatrix(intRow, mHeadCol.��λ) = rsInitCard!��λ
                    .TextMatrix(intRow, mHeadCol.����ϵ��) = rsInitCard!����ϵ��
                    .TextMatrix(intRow, mHeadCol.�б����) = IIf(Val(zlStr.Nvl(rsInitCard!�б����)) = 1, "��", "")
                    .TextMatrix(intRow, mHeadCol.�빺����) = IIf(Format(Val(zlStr.Nvl(rsInitCard!�빺����)), mFMT.FM_����) = 0, "", Format(rsInitCard!�빺���� / rsInitCard!����ϵ��, mFMT.FM_����))
                    If mint�༭״̬ = 3 Then
                        .TextMatrix(intRow, mHeadCol.�ƻ�����) = .TextMatrix(intRow, mHeadCol.�빺����)
                    Else
                        .TextMatrix(intRow, mHeadCol.�ƻ�����) = IIf(Format(Val(zlStr.Nvl(rsInitCard!�ƻ�����)), mFMT.FM_����) = 0, "", Format(rsInitCard!�ƻ����� / rsInitCard!����ϵ��, mFMT.FM_����))
                    End If
                    .TextMatrix(intRow, mHeadCol.����) = Format(Val(zlStr.Nvl(rsInitCard!����)) * rsInitCard!����ϵ��, mFMT.FM_�ɱ���)
                    .TextMatrix(intRow, mHeadCol.���) = IIf(Format(Val(zlStr.Nvl(rsInitCard!���)), mFMT.FM_���) = 0, "", Format(Val(zlStr.Nvl(rsInitCard!���)), mFMT.FM_���))
                    If intRow = .Rows - 1 Then .Rows = .Rows + 1
                    rsInitCard.MoveNext
                Next
            End With
            rsInitCard.Close
    End Select
    Call SetEdit
    Call RefreshRowNO(mshBill, mHeadCol.���, 1)
    Call ��ʾ�ϼƽ��
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'��ʼ���༭�ؼ�
Private Sub initGrid()
    Dim intCol As Integer

    With mshBill
        .Active = True
        .Cols = mconIntColS
        .MsfObj.FixedCols = 2

        .TextMatrix(0, mHeadCol.���) = "���"
        .TextMatrix(0, mHeadCol.����) = "�������������"
        .TextMatrix(0, mHeadCol.���) = "���"
        .TextMatrix(0, mHeadCol.����) = "����"
        .TextMatrix(0, mHeadCol.��λ) = "��λ"
        .TextMatrix(0, mHeadCol.����ϵ��) = "����ϵ��"
        .TextMatrix(0, mHeadCol.�б����) = "�б����"
        .TextMatrix(0, mHeadCol.�빺����) = "�빺����"
        .TextMatrix(0, mHeadCol.�ƻ�����) = "��������"
        .TextMatrix(0, mHeadCol.����) = "�ɱ���"
        .TextMatrix(0, mHeadCol.���) = "�ɱ����"
        .TextMatrix(0, mHeadCol.�ϴι�Ӧ��) = "�ϴι�Ӧ��"
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mHeadCol.���) = "1"

        .ColWidth(mHeadCol.���) = 500
        .ColWidth(mHeadCol.����) = 2000
        .ColWidth(mHeadCol.���) = 900
        .ColWidth(mHeadCol.����) = 800
        .ColWidth(mHeadCol.��λ) = 500
        .ColWidth(mHeadCol.�б����) = 800
        .ColWidth(mHeadCol.�빺����) = 1000
        .ColWidth(mHeadCol.�ƻ�����) = 1000
        .ColWidth(mHeadCol.����) = IIf(mblnCostView = False, 0, 1000)
        .ColWidth(mHeadCol.���) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mHeadCol.�ϴι�Ӧ��) = IIf(mblnProvider = False, 0, 1000)
        .ColWidth(mHeadCol.����ϵ��) = 0
        .ColWidth(0) = 0

        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��
        For intCol = 0 To .Cols - 1
            .ColData(intCol) = 5
        Next

        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
            txtժҪ.Enabled = True
            .ColData(mHeadCol.����) = 1
            .ColData(mHeadCol.�빺����) = 4
            .ColData(mHeadCol.����) = 4
            .ColData(mHeadCol.����) = 4
            .ColData(mHeadCol.�ϴι�Ӧ��) = 1
        ElseIf mint�༭״̬ = 3 Then
            txtժҪ.Enabled = False
            .ColData(mHeadCol.�ƻ�����) = 4
        ElseIf mint�༭״̬ = 4 Then
            txtժҪ.Enabled = False
            .ColData(mHeadCol.�ƻ�����) = 0
        End If

        .ColAlignment(mHeadCol.����) = flexAlignLeftCenter
        .ColAlignment(mHeadCol.���) = flexAlignLeftCenter
        .ColAlignment(mHeadCol.����) = flexAlignLeftCenter
        .ColAlignment(mHeadCol.��λ) = flexAlignCenterCenter
        .ColAlignment(mHeadCol.�빺����) = flexAlignRightCenter
        .ColAlignment(mHeadCol.�ƻ�����) = flexAlignRightCenter
        .ColAlignment(mHeadCol.����) = flexAlignRightCenter
        .ColAlignment(mHeadCol.���) = flexAlignRightCenter
        .ColAlignment(mHeadCol.�ϴι�Ӧ��) = flexAlignLeftCenter
        .ColAlignment(mHeadCol.�б����) = 4
        If mint�༭״̬ = 3 Then
            .PrimaryCol = mHeadCol.����
            .LocateCol = mHeadCol.�ƻ�����
        Else
            .PrimaryCol = mHeadCol.����
            .LocateCol = mHeadCol.����
        End If
        If InStr(1, "34", mint�༭״̬) <> 0 Then .ColData(mHeadCol.����) = 0
    End With

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    With Pic����
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top - 100 - CmdCancel.Height - 200
    End With

    With LblTitle
        .Left = 0
        .Top = 150
        .Width = Pic����.Width
    End With


    With mshBill
        .Left = 200
        .Width = Pic����.Width - .Left * 2
    End With
    
    With txtNO
        .Left = mshBill.Left + mshBill.Width - .Width
        LblNo.Left = .Left - LblNo.Width - 100
        .Top = LblTitle.Top
        LblNo.Top = .Top
    End With
    
    cbo����.Left = mshBill.Left + mshBill.Width - cbo����.Width
    Lbl�ƻ�����.Left = cbo����.Left - Lbl�ƻ�����.Width - 50


    LblStock.Left = mshBill.Left
    cboStock.Left = LblStock.Left + LblStock.Width + 50

    LblEnterStock.Left = cboStock.Left + cboStock.Width + cboStock.Width * 0.3
    cboEnterStock.Left = LblEnterStock.Left + LblEnterStock.Width + 50

    With Lbl������
        .Top = Pic����.Height - 200 - .Height
        .Left = mshBill.Left + 100
    End With

    With Txt������
        .Top = Lbl������.Top - 80
        .Left = Lbl������.Left + Lbl������.Width + 100
    End With

    With Lbl��������
        .Top = Lbl������.Top
        .Left = Txt������.Left + Txt������.Width + 250
    End With

    With Txt��������
        .Top = Lbl��������.Top - 80
        .Left = Lbl��������.Left + Lbl��������.Width + 100
    End With

    With Txt�������
        .Top = Lbl������.Top - 80
        .Left = mshBill.Left + mshBill.Width - .Width
    End With

    With Lbl�������
        .Top = Lbl������.Top
        .Left = Txt�������.Left - 100 - .Width
    End With

    With Txt�����
        .Top = Lbl������.Top - 80
        .Left = Lbl�������.Left - 200 - .Width
    End With

    With Lbl�����
        .Top = Lbl������.Top
        .Left = Txt�����.Left - 100 - .Width
    End With

    With txtժҪ
        .Top = Lbl������.Top - 140 - .Height
        .Left = Txt������.Left
        .Width = mshBill.Left + mshBill.Width - .Left
    End With

    With lblժҪ
        .Top = txtժҪ.Top + 50
        .Left = txtժҪ.Left - .Width - 100
    End With

    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = txtժҪ.Top - 60 - .Height
        .Width = mshBill.Width
    End With

    With mshBill
        .Height = lblPurchasePrice.Top - .Top - 60
    End With

    With CmdCancel
        .Left = Pic����.Left + mshBill.Left + mshBill.Width - .Width
        .Top = Pic����.Top + Pic����.Height + 100
    End With

    With CmdSave
        .Left = CmdCancel.Left - .Width - 100
        .Top = CmdCancel.Top
    End With

    With cmdHelp
        .Left = Pic����.Left + mshBill.Left
        .Top = CmdCancel.Top
    End With

    With cmdFind
        .Top = CmdCancel.Top
    End With

    With lblCode
        .Top = CmdCancel.Top + 50
    End With
    With txtCode
        .Top = CmdCancel.Top + 30
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Or mint�༭״̬ = 4 Or mint�༭״̬ = 3 Then
        SaveWinState Me, App.ProductName, mstrCaption
        Exit Sub
    End If
    If MsgBox("���ݿ����Ѹı䣬��δ���̣���Ҫ�˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        SaveWinState Me, App.ProductName, mstrCaption
    End If

End Sub
Private Function SaveCheck() As Boolean
    Dim str����� As String, intRow As Integer, lng��� As Long, ����ID_IN As Long
    Dim dbl����_IN As Double, dbl���_IN As Double, dbl�빺����_IN As Double, dbl�ƻ�����_IN As Double
    Dim �ϴι�Ӧ��_IN As String, �ϴ�������_IN As String
    Dim cllProc As New Collection
    mblnSave = False
    SaveCheck = False

    str����� = gstrUserName
    'Zl_���ϼƻ�����_Delete
    gstrSQL = "Zl_���ϼƻ�����_Delete("
    '  Id_In       In ���ϲɹ��ƻ�.ID%Type,
    gstrSQL = gstrSQL & "" & mlng�ƻ�ID & ","
    '  ɾ����ϸ_In Integer:=0
    '  --1-ֻɾ����ϸ,����ȫɾ��
    gstrSQL = gstrSQL & "1)"
    cllProc.Add gstrSQL
    '������ϸ
    With mshBill
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng��� = .TextMatrix(intRow, mHeadCol.���)
                ����ID_IN = .TextMatrix(intRow, 0)
                dbl����_IN = Round(Val(.TextMatrix(intRow, mHeadCol.����)) / Val(.TextMatrix(intRow, mHeadCol.����ϵ��)), g_С��λ��.obj_ɢװС��.�ɱ���С��)
                dbl���_IN = Round(Val(.TextMatrix(intRow, mHeadCol.���)), g_С��λ��.obj_ɢװС��.���С��)
                dbl�빺����_IN = Round(Val(.TextMatrix(intRow, mHeadCol.�빺����)) * Val(.TextMatrix(intRow, mHeadCol.����ϵ��)), g_С��λ��.obj_ɢװС��.����С��)
                dbl�ƻ�����_IN = Round(Val(.TextMatrix(intRow, mHeadCol.�ƻ�����)) * Val(.TextMatrix(intRow, mHeadCol.����ϵ��)), g_С��λ��.obj_ɢװС��.����С��)
                �ϴι�Ӧ��_IN = .TextMatrix(intRow, mHeadCol.�ϴι�Ӧ��)
                �ϴ�������_IN = .TextMatrix(intRow, mHeadCol.����)
                'Zl_���ϼƻ�����α�_Insert
                gstrSQL = "Zl_���ϼƻ�����α�_Insert("
                '  �ƻ�id_In     In ���ϼƻ�����.�ƻ�id%Type,
                gstrSQL = gstrSQL & "" & mlng�ƻ�ID & ","
                '  ����id_In     In ���ϼƻ�����.����id%Type,
                gstrSQL = gstrSQL & "" & ����ID_IN & ","
                '  ���_In       In ���ϼƻ�����.���%Type,
                gstrSQL = gstrSQL & "" & lng��� & ","
                '  �빺����_In   In ���ϼƻ�����.�빺����%Type,
                gstrSQL = gstrSQL & "" & dbl�빺����_IN & ","
                '  �ƻ�����_IN   In ���ϼƻ�����.�ƻ�����%Type,
                gstrSQL = gstrSQL & "" & dbl�ƻ�����_IN & ","
                '  ����_IN       In ���ϼƻ�����.����%Type,
                gstrSQL = gstrSQL & "" & dbl����_IN & ","
                '  ���_IN       In ���ϼƻ�����.���%Type,
                gstrSQL = gstrSQL & "" & dbl���_IN & ","
                '  ǰ������_In   In ���ϼƻ�����.ǰ������%Type := Null,
                gstrSQL = gstrSQL & "" & 0 & ","
                '  ��������_In   In ���ϼƻ�����.��������%Type := Null,
                gstrSQL = gstrSQL & "" & 0 & ","
                '  �������_In   In ���ϼƻ�����.�������%Type := Null,
                gstrSQL = gstrSQL & "" & 0 & ","
                '  �ϴι�Ӧ��_In In ���ϼƻ�����.�ϴι�Ӧ��%Type := Null,
                gstrSQL = gstrSQL & "'" & �ϴι�Ӧ��_IN & "',"
                '  �ϴ�������_In In ���ϼƻ�����.�ϴ�������%Type := Null
                gstrSQL = gstrSQL & "'" & �ϴ�������_IN & "')"
                cllProc.Add gstrSQL
            End If
        Next
    End With
    'zl_���ϼƻ�����_VERIFY( /*ID_IN*/, /*�����_IN*/ );
    gstrSQL = "zl_���ϼƻ�����_VERIFY('" & mlng�ƻ�ID & "','" & str����� & "')"
    cllProc.Add gstrSQL
    
    err = 0: On Error GoTo ErrHandle
    ExecuteProcedureArrAy cllProc, mstrCaption
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Msf��Ӧ��ѡ��_DblClick()
    Dim blnCancel As Boolean
    With mshBill
        .Text = Msf��Ӧ��ѡ��.TextMatrix(Msf��Ӧ��ѡ��.Row, 2)
        .TextMatrix(.Row, mHeadCol.�ϴι�Ӧ��) = Msf��Ӧ��ѡ��.TextMatrix(Msf��Ӧ��ѡ��.Row, 2)
    End With
    Msf��Ӧ��ѡ��.Visible = False
    mshBill.SetFocus
    Call SendKeys("{ENTER}")
End Sub

Private Sub Msf��Ӧ��ѡ��_GotFocus()
    If Msf��Ӧ��ѡ��.Rows - 1 = 1 Then Call Msf��Ӧ��ѡ��_DblClick
End Sub

Private Sub Msf��Ӧ��ѡ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Msf��Ӧ��ѡ��_DblClick
    End If
End Sub

Private Sub Msf��Ӧ��ѡ��_LostFocus()
    Msf��Ӧ��ѡ��.ZOrder 1
    Msf��Ӧ��ѡ��.Visible = False
End Sub

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mHeadCol.���, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call RefreshRowNO(mshBill, mHeadCol.���, mshBill.Row)
    Call ��ʾ�ϼƽ��
End Sub

Private Sub mshBill_BeforeAddRow(Row As Long)
    If mshBill.ColData(mHeadCol.����) = 0 Then
        Exit Sub
    End If
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If InStr(1, "34", mint�༭״̬) <> 0 Then
        Cancel = True
        Exit Sub
    End If
    With mshBill
        If .TextMatrix(.Row, 0) <> "" Then
            If MsgBox("��ȷʵҪɾ����������������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub mshbill_CommandClick()
    Dim sngLeft As Single, sngTop As Single
    Dim RecReturn As Recordset
    Dim strUnit As String
    Dim i As Integer
    Dim int����� As Integer
    
    int����� = mshBill.Row
    
    On Error GoTo ErrHandle
    If mshBill.Col = mHeadCol.���� Then
        
        If cboEnterStock.ListIndex = -1 Then
            MsgBox "��ѡ���깺�Ŀⷿ��", vbInformation + vbOKOnly, gstrSysName
            cboEnterStock.SetFocus
            Exit Sub
        End If
        
        Set RecReturn = Frm����ѡ����.ShowMe(Me, 1, Val(cboEnterStock.ItemData(cboEnterStock.ListIndex)), Val(cboEnterStock.ItemData(cboEnterStock.ListIndex)), , , , , , , , , , , , mlngModule, , mstrPrivs, , False)
        If RecReturn.RecordCount > 0 Then
            mblnChange = True
            RecReturn.MoveFirst
            
            If mintUnit = 0 Then
                strUnit = "ɢװ��λ"
            Else
                strUnit = "��װ��λ"
            End If
            
            For i = 1 To RecReturn.RecordCount
                If SetStuffRows(RecReturn!����ID, "[" & RecReturn!���� & "]" & RecReturn!����, _
                            IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                            Switch(strUnit = "ɢװ��λ", zlStr.Nvl(RecReturn!ɢװ��λ), strUnit = "��װ��λ", zlStr.Nvl(RecReturn!��װ��λ)), Val(zlStr.Nvl(RecReturn!ָ��������)), _
                            Switch(strUnit = "ɢװ��λ", 1, strUnit = "��װ��λ", Val(zlStr.Nvl(RecReturn!����ϵ��)))) Then
                    
                    If mshBill.Row = mshBill.Rows - 1 Then mshBill.Rows = mshBill.Rows + 1 'ֻ�е�ǰ�������һ��ʱ��������
                    mshBill.Row = mshBill.Row + 1
                End If
            
                RecReturn.MoveNext
            Next
            
            mshBill.Row = int�����
            
            If mstr�ظ����� <> "" Then
                MsgBox mstr�ظ����� & "�б����Ѿ������ˣ�" & vbCrLf & "�������Ĳ�����ӣ�", vbInformation + vbOKOnly, gstrSysName
                mstr�ظ����� = ""
            End If
            
'            If RecReturn.RecordCount = 1 Then
'                If mintUnit = 0 Then
'                    strUnit = "ɢװ��λ"
'                Else
'                    strUnit = "��װ��λ"
'                End If
'                SetStuffRows RecReturn!����ID, "[" & RecReturn!���� & "]" & RecReturn!����, _
'                            IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
'                            Switch(strUnit = "ɢװ��λ", zlStr.Nvl(RecReturn!ɢװ��λ), strUnit = "��װ��λ", zlStr.Nvl(RecReturn!��װ��λ)), Val(zlStr.Nvl(RecReturn!ָ��������)), _
'                            Switch(strUnit = "ɢװ��λ", 1, strUnit = "��װ��λ", Val(zlStr.Nvl(RecReturn!����ϵ��)))
'            End If
            RecReturn.Close
        End If
    Else
        'ҩƷ��Ӧ�̵�ѡ��
        sngLeft = mshBill.Left + mshBill.MsfObj.CellLeft
        sngTop = mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
        If sngLeft + Msf��Ӧ��ѡ��.Width > Me.ScaleWidth Then sngLeft = Me.ScaleWidth - Msf��Ӧ��ѡ��.Width - 100
        Set RecReturn = New ADODB.Recordset
        gstrSQL = "Select ID,����,����,���� From ��Ӧ�� " & _
                  "Where ĩ��=1 And (substr(����,5,1)=1 And (վ��=[1] or վ�� is null) Or Nvl(ĩ��,0)=0) " & _
                  "    And (To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' or ����ʱ�� is null) Order By ���� "
        Set RecReturn = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�������Ϲ�Ӧ��", gstrNodeNo)
        If RecReturn.RecordCount = 0 Then
            MsgBox "���ȳ�ʼ���������Ϲ�Ӧ�̣�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        With Msf��Ӧ��ѡ��
            .Clear
            Set .DataSource = RecReturn
            .ColWidth(0) = 0
            .ColWidth(1) = 800
            .ColWidth(2) = 3000
            .ColWidth(3) = 800
            .Row = 1
            .ColSel = .Cols - 1
        End With
        With Msf��Ӧ��ѡ��
            .Left = sngLeft
            .Top = sngTop
            .Visible = True
            .ZOrder 0
            .SetFocus
        End With
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub

Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If

        Select Case .Col
            Case mHeadCol.����
                .TxtCheck = False
                .MaxLength = 80
                'ֻ��ҩ���в���ʾ�ϼ���Ϣ�Ϳ����
                Call ��ʾ�ϼƽ��
            Case mHeadCol.����
                .TxtCheck = False
                .MaxLength = 40
            Case mHeadCol.�ϴι�Ӧ��
                .MaxLength = 40
                .TxtCheck = False
            Case mHeadCol.�ƻ�����
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
            Case mHeadCol.�빺����
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
            Case mHeadCol.����
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"

        End Select

    End With
End Sub
Private Function Get����(Optional strKey As String = "") As Boolean
    '����:��ȡ��Ա��Ϣ
    Dim rsTemp  As ADODB.Recordset
    Dim blnCancel  As Boolean
    Dim strSearch As String
    Dim vRect As RECT
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
     
     If strKey <> "" Then
        strSearch = GetMatchingSting(strKey)
        gstrSQL = "" & _
            "   Select ���� as id ,a.���� ,a.���� ,a.����,a.������ҵ���֤ " & _
            "   From ���������� a " & _
            "   Where (���� like [1] or ���� like [1] or ���� like [1])  " & _
            "   order by ����"
         vRect = zlControl.GetControlRect(mshBill.TxtHwnd)
         Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "����ѡ��", False, "", "", False, False, True, vRect.Left - 15, vRect.Top, mshBill.RowHeight(mshBill.Row) - 50, blnCancel, False, False, strSearch)
     Else
        gstrSQL = "" & _
            "   Select ���� as id,a.���� ,a.���� ,a.����,a.������ҵ���֤" & _
            "   From ���������� a " & _
            "   order by ����"
        Set rsTemp = zlDatabase.ShowSelect(Me, gstrSQL, 0, "����ѡ����", True, "", "��ѡ����صĲ���", True, False, , , , , blnCancel)
    End If
    If rsTemp Is Nothing Then Exit Function
    If blnCancel = True Then Exit Function
    
    With mshBill '
        .TextMatrix(.Row, mHeadCol.����) = zlStr.Nvl(rsTemp!����)
        .Text = .TextMatrix(.Row, mHeadCol.����)
    End With
    Get���� = True
End Function

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsStuff As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    
    Dim rsTemp As Recordset
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim i As Integer
    Dim int����� As Integer
    
    int����� = mshBill.Row
    
    On Error GoTo ErrHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    With mshBill
        If .Col = mHeadCol.���� Then
            .Text = Trim(.Text)
        Else
            .Text = Trim(.Text)
        End If
        strKey = .Text

        If Mid(strKey, 1, 1) = "[" Then
            If InStr(2, strKey, "]") <> 0 Then
                strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
            Else
                strKey = Mid(strKey, 2)
            End If
        End If
        Select Case .Col

            Case mHeadCol.����
                If strKey <> "" Then
                    If cboEnterStock.ListIndex = -1 Then
                        MsgBox "��ѡ���깺�Ŀⷿ��", vbInformation + vbOKOnly, gstrSysName
                        cboEnterStock.SetFocus
                        Exit Sub
                    End If
        
                    sngLeft = Me.Left + Pic����.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic����.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 4530
                    End If

                    Set rsTemp = FrmMulitSel.ShowSelect(Me, 1, Val(cboEnterStock.ItemData(cboEnterStock.ListIndex)), Val(cboEnterStock.ItemData(cboEnterStock.ListIndex)), , strKey, sngLeft, sngTop, mshBill.MsfObj.CellWidth, mshBill.MsfObj.CellHeight, , , , , , , , , , mlngModule, , mstrPrivs, , False)
                    
                    If rsTemp.RecordCount <= 0 Then
                        Cancel = True
                        Exit Sub
                    End If
                    
                    If mintUnit = 0 Then
                        strUnit = "ɢװ��λ"
                    Else
                        strUnit = "��װ��λ"
                    End If
                    
                    rsTemp.MoveFirst
                    For i = 1 To rsTemp.RecordCount
                        If SetStuffRows(rsTemp!����ID, "[" & rsTemp!���� & "]" & rsTemp!����, _
                            IIf(IsNull(rsTemp!���), "", rsTemp!���), IIf(IsNull(rsTemp!����), "", rsTemp!����), _
                            Switch(strUnit = "ɢװ��λ", zlStr.Nvl(rsTemp!ɢװ��λ), strUnit = "��װ��λ", zlStr.Nvl(rsTemp!��װ��λ)), Val(zlStr.Nvl(rsTemp!ָ��������)), _
                            Switch(strUnit = "ɢװ��λ", 1, strUnit = "��װ��λ", Val(zlStr.Nvl(rsTemp!����ϵ��)))) Then
                            
                            If .Row = .Rows - 1 Then .Rows = .Rows + 1 'ֻ�е�ǰ�������һ��ʱ��������
                            .Row = .Row + 1
                            
                            .Text = .TextMatrix(.Row, .Col)
                        Else
                            Cancel = True
                        End If
                        
                        rsTemp.MoveNext
                    Next
                    
                    mshBill.Row = int�����
                    
                    If mstr�ظ����� <> "" Then
                        MsgBox mstr�ظ����� & "�б����Ѿ������ˣ�" & vbCrLf & "�������Ĳ�����ӣ�", vbInformation + vbOKOnly, gstrSysName
                        mstr�ظ����� = ""
                    End If
                    
'                    If rsTemp.RecordCount = 1 Then
'                        If mintUnit = 0 Then
'                            strUnit = "ɢװ��λ"
'                        Else
'                            strUnit = "��װ��λ"
'                        End If
'                        If SetStuffRows(rsTemp!����ID, "[" & rsTemp!���� & "]" & rsTemp!����, _
'                            IIf(IsNull(rsTemp!���), "", rsTemp!���), IIf(IsNull(rsTemp!����), "", rsTemp!����), _
'                            Switch(strUnit = "ɢװ��λ", zlStr.NVL(rsTemp!ɢװ��λ), strUnit = "��װ��λ", zlStr.NVL(rsTemp!��װ��λ)), Val(zlStr.NVL(rsTemp!ָ��������)), _
'                            Switch(strUnit = "ɢװ��λ", 1, strUnit = "��װ��λ", Val(zlStr.NVL(rsTemp!����ϵ��)))) = False Then
'                            Cancel = True
'                            Exit Sub
'                        End If
'                        .Text = .TextMatrix(.Row, .Col)
'                    Else
'
'                        Cancel = True
'                    End If
                End If
            Case mHeadCol.�ƻ�����
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "������������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Val(strKey) > 99999999 Or Val(strKey) < 0 Then
                    MsgBox "��������������(0~99999999)��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If .Text = "" Then
                    Cancel = True
                    Exit Sub
                End If
                If strKey <> "" Then
                    strKey = Format(strKey, mFMT.FM_����)
                    .Text = strKey
                    If .TextMatrix(.Row, mHeadCol.����) <> "" Then
                        .TextMatrix(.Row, mHeadCol.���) = Format(.TextMatrix(.Row, mHeadCol.����) * strKey, mFMT.FM_���)
                    End If
                End If
                Call ��ʾ�ϼƽ��
            Case mHeadCol.�빺����
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "�빺��������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Val(strKey) > 99999999 Or Val(strKey) < 0 Then
                    MsgBox "�빺����������(0~99999999)��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If .Text = "" Then
                    Cancel = True
                    Exit Sub
                End If
                If strKey <> "" Then
                    strKey = Format(strKey, mFMT.FM_����)
                    .Text = strKey
                    If .TextMatrix(.Row, mHeadCol.����) <> "" Then
                        .TextMatrix(.Row, mHeadCol.���) = Format(.TextMatrix(.Row, mHeadCol.����) * strKey, mFMT.FM_���)
                    End If
                End If
                Call ��ʾ�ϼƽ��
            Case mHeadCol.����
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "�ɱ��۱���Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Val(strKey) > 99999999 Or Val(strKey) < 0 Then
                    MsgBox "�ɱ��۱�����(0~99999999)��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mHeadCol.����) = " "
                        .Text = " "
                    End If
                End If
                If strKey <> "" Then
                    strKey = Format(strKey, mFMT.FM_����)
                    .Text = strKey
                    .TextMatrix(.Row, mHeadCol.����) = strKey
                End If
                .TextMatrix(.Row, mHeadCol.���) = Format(Val(.TextMatrix(.Row, mHeadCol.����)) * Val(.TextMatrix(.Row, mHeadCol.�빺����)), mFMT.FM_���)
                Call ��ʾ�ϼƽ��
                
            Case mHeadCol.����
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mHeadCol.����) = ""
                    End If
                    .Col = mHeadCol.�빺����
                    Cancel = True
                    Exit Sub
                Else
                    If strKey <> "" Then
                        If Get����(strKey) = False Then
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                End If
                OS.OpenIme False
            Case mHeadCol.�ϴι�Ӧ��
'                If .TxtVisible = False Then Exit Sub
                If strKey = "" And .TextMatrix(.Row, mHeadCol.�ϴι�Ӧ��) = "" Then
                    strKey = " "
                    .Text = strKey
                    .TextMatrix(.Row, mHeadCol.�ϴι�Ӧ��) = strKey
                Else
                    If .TxtVisible = False Then Exit Sub
                    If StrIsValid(strKey, 40) = False Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strKey = UCase(strKey)
                    sngLeft = mshBill.Left + mshBill.MsfObj.CellLeft
                    sngTop = mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngLeft + Msf��Ӧ��ѡ��.Width > Me.ScaleWidth Then sngLeft = Me.ScaleWidth - Msf��Ӧ��ѡ��.Width - 100
            
                    Set rsTemp = New ADODB.Recordset
                    gstrSQL = "" & _
                        "   Select ID,����,����,���� " & _
                        "   From ��Ӧ�� " & _
                        "   Where ĩ��=1 And (substr(����,5,1)=1 And (վ��=[2] or վ�� is null) Or Nvl(ĩ��,0)=0) " & _
                        "       And (To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' or ����ʱ�� is null) " & _
                        "       And (upper(����) Like [1] Or Upper(����) Like [1] Or Upper(����) Like [1])" & _
                        "   Order By ���� "
                    
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�������Ϲ�Ӧ��", strKey & "%", gstrNodeNo)
                    
                    If rsTemp.RecordCount = 0 Then
                        MsgBox "û���ҵ����������Ĺ�Ӧ�̣�", vbInformation, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    ElseIf rsTemp.RecordCount = 1 Then
                        .Text = rsTemp!����
                        Exit Sub
                    End If
                    
                    With Msf��Ӧ��ѡ��
                        .Clear
                        Set .DataSource = rsTemp
                        .ColWidth(0) = 0
                        .ColWidth(1) = 800
                        .ColWidth(2) = 3000
                        .ColWidth(3) = 800
            
                        .Row = 1
                        .ColSel = .Cols - 1
                    End With
                    With Msf��Ӧ��ѡ��
                        .Left = sngLeft
                        .Top = sngTop
                        .Visible = True
                        .ZOrder 0
                        .SetFocus
                    End With
                    Cancel = True
                End If
        End Select
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "PY" And stbThis.Tag <> "PY" Then
        Logogram stbThis, 0
        stbThis.Tag = Panel.Key
    ElseIf Panel.Key = "WB" And stbThis.Tag <> "WB" Then
        Logogram stbThis, 1
        stbThis.Tag = Panel.Key
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    If KeyAscii = 13 Then
        cmdFind_Click
    End If
End Sub

Private Function ValidData() As Boolean
    ValidData = False
    Dim intLop As Integer

    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '�����з�����
            
            If cboEnterStock.ListIndex = -1 Then
                ShowMsgBox "���깺�ⷿ����Ϊ�գ�"
                cboEnterStock.SetFocus
                Exit Function
            End If

            If LenB(StrConv(txtժҪ.Text, vbFromUnicode)) > 40 Then
                MsgBox "ժҪ����,���������20�����ֻ�40���ַ�!", vbInformation + vbOKOnly, gstrSysName
                txtժҪ.SetFocus
                Exit Function
            End If

            For intLop = 1 To .Rows - 1
                If Trim(.TextMatrix(intLop, mHeadCol.����)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mHeadCol.�ƻ�����))) <> "" Then
                        If Not IsNumeric(.TextMatrix(intLop, mHeadCol.�ƻ�����)) Then
                            MsgBox "��" & intLop & "���������ϵ�����������Ϊ�����ͣ����飡", vbInformation, gstrSysName
                            mshBill.SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            .Col = mHeadCol.�ƻ�����
                            Exit Function
                        End If
                    End If
                    
                    If Val(.TextMatrix(intLop, mHeadCol.�ƻ�����)) > 9999999999# Then
                        MsgBox "��" & intLop & "���������ϵ������������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mHeadCol.�ƻ�����
                        Exit Function
                    End If
                    If mint�༭״̬ <> 3 Then
                        If Trim(Trim(.TextMatrix(intLop, mHeadCol.�빺����))) <> "" Then
                            If Not IsNumeric(.TextMatrix(intLop, mHeadCol.�빺����)) Then
                                MsgBox "��" & intLop & "���������ϵ��빺������Ϊ�����ͣ����飡", vbInformation, gstrSysName
                                mshBill.SetFocus
                                .Row = intLop
                                .MsfObj.TopRow = intLop
                                .Col = mHeadCol.�빺����
                                Exit Function
                            End If
                        End If
                        
                        If Val(.TextMatrix(intLop, mHeadCol.�빺����)) > 9999999999# Then
                            MsgBox "��" & intLop & "���������ϵ��빺�������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                            mshBill.SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            .Col = mHeadCol.�빺����
                            Exit Function
                        End If
                    End If
                    If Val(.TextMatrix(intLop, mHeadCol.����)) > 9999999999# Then
                        MsgBox "��" & intLop & "���������ϵĳɱ��۴��������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mHeadCol.����
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mHeadCol.���)) > 9999999999999# Then
                        MsgBox "��" & intLop & "���������ϵĳɱ������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        If mint�༭״̬ = 3 Then
                            .Col = mHeadCol.�ƻ�����
                        Else
                            .Col = mHeadCol.�빺����
                        End If
                        Exit Function
                    End If
                End If
            Next
        Else
            Exit Function
        End If
    End With

    ValidData = True
End Function


Private Function SaveCard() As Boolean
    Dim lng��� As Long
    Dim ID_IN As Long
    Dim NO_IN As Variant
    Dim �ƻ�����_IN As Integer
    Dim �ڼ�_IN As String
    Dim �ⷿID_IN As Long
    Dim ���Ʒ���_IN As Integer
    Dim ������_IN As String
    Dim ��������_IN As String
    Dim ����˵��_IN As String

    Dim ����ID_IN As Long
    Dim dbl�ƻ�����_IN As Double
    Dim dbl����_IN As Double
    Dim dbl���_IN As Double
    Dim dbl�빺����_IN As Double
    Dim ��������_IN As Double
    Dim �������_IN As Double
    Dim �ϴι�Ӧ��_IN As String
    Dim �ϴ�������_IN As String, intMonth As Integer
    Dim lng����ID As Long
    Dim intRow As Integer
    Dim cllTemp As New Collection
    SaveCard = False
    Select Case cbo����.ListIndex + 1
        Case 1       '�¼ƻ�
            mstr�ڼ� = Format(DateAdd("m", 1, sys.Currentdate), "yyyyMM")
        Case 2       '���ƻ�
            intMonth = Month(DateAdd("Q", 1, sys.Currentdate))
            mstr�ڼ� = Format(DateAdd("Q", 1, sys.Currentdate), "yyyy") & IIf(intMonth <= 3, 1, IIf(intMonth >= 10, 4, IIf(intMonth <= 9 And intMonth >= 7, 3, 2)))
        Case Else    '��ƻ�
            mstr�ڼ� = Format(DateAdd("yyyy", 1, sys.Currentdate), "yyyy")
    End Select
            
    With mshBill
        ID_IN = sys.NextId("���ϲɹ��ƻ�")
        NO_IN = Trim(txtNO)
        
        If NO_IN = "" Then NO_IN = sys.GetNextNo(85, mlng�ⷿid)
        If IsNull(NO_IN) Then Exit Function
        Me.txtNO.Tag = NO_IN
        
        �ƻ�����_IN = cbo����.ListIndex + 1
        ���Ʒ���_IN = mint���Ʒ���
        If cboEnterStock.ListIndex < 0 Then
            �ⷿID_IN = 0
        Else
            �ⷿID_IN = cboEnterStock.ItemData(cboEnterStock.ListIndex)
        End If
        lng����ID = cboStock.ItemData(cboStock.ListIndex)
        ������_IN = gstrUserName
        ��������_IN = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        ����˵��_IN = Trim(txtժҪ.Text)
        �ڼ�_IN = mstr�ڼ�

        If mint�༭״̬ = 2 Then        '�޸�
            gstrSQL = "zl_���ϼƻ�����_DELETE('" & mlng�ƻ�ID & "')"
            cllTemp.Add gstrSQL
        End If
        'Zl_���ϼƻ���������_Insert
        gstrSQL = "Zl_���ϼƻ���������_Insert("
        '  Id_In       In ���ϲɹ��ƻ�.ID%Type,
        gstrSQL = gstrSQL & "" & ID_IN & ","
        '  ����_In     In ���ϲɹ��ƻ�.����%Type,
        gstrSQL = gstrSQL & "" & 1 & ","
        '  No_In       In ���ϲɹ��ƻ�.NO%Type,
        gstrSQL = gstrSQL & "'" & NO_IN & "',"
        '  �ƻ�����_In In ���ϲɹ��ƻ�.�ƻ�����%Type,
        gstrSQL = gstrSQL & "" & �ƻ�����_IN & ","
        '  �ڼ�_In     In ���ϲɹ��ƻ�.�ڼ�%Type,
        gstrSQL = gstrSQL & "'" & �ڼ�_IN & "',"
        '  �ⷿid_In   In ���ϲɹ��ƻ�.�ⷿid%Type,
        gstrSQL = gstrSQL & "" & IIf(�ⷿID_IN = 0, "NULL", �ⷿID_IN) & ","
        '  ����id_In   In ���ϲɹ��ƻ�.����id%Type,
        gstrSQL = gstrSQL & "" & lng����ID & ","
        '  ���Ʒ���_In In ���ϲɹ��ƻ�.���Ʒ���%Type,
        gstrSQL = gstrSQL & "" & ���Ʒ���_IN & ","
        '  ������_In   In ���ϲɹ��ƻ�.������%Type,
        gstrSQL = gstrSQL & "'" & ������_IN & "',"
        '  ��������_In In ���ϲɹ��ƻ�.��������%Type,
        gstrSQL = gstrSQL & "to_date('" & ��������_IN & "','yyyy-mm-dd HH24:MI:SS'),"
        '  ����˵��_In In ���ϲɹ��ƻ�.����˵��%Type := Null
        gstrSQL = gstrSQL & "'" & ����˵��_IN & "')"
        cllTemp.Add gstrSQL
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng��� = .TextMatrix(intRow, mHeadCol.���)
                ����ID_IN = .TextMatrix(intRow, 0)
                dbl����_IN = Round(Val(.TextMatrix(intRow, mHeadCol.����)) / Val(.TextMatrix(intRow, mHeadCol.����ϵ��)), g_С��λ��.obj_���С��.�ɱ���С��)
                dbl���_IN = Round(Val(.TextMatrix(intRow, mHeadCol.���)), g_С��λ��.obj_���С��.���С��)
                dbl�빺����_IN = Round(Val(.TextMatrix(intRow, mHeadCol.�빺����)) * Val(.TextMatrix(intRow, mHeadCol.����ϵ��)), g_С��λ��.obj_���С��.����С��)
                dbl�ƻ�����_IN = Round(Val(.TextMatrix(intRow, mHeadCol.�ƻ�����)) * Val(.TextMatrix(intRow, mHeadCol.����ϵ��)), g_С��λ��.obj_���С��.����С��)
                �ϴι�Ӧ��_IN = .TextMatrix(intRow, mHeadCol.�ϴι�Ӧ��)
                �ϴ�������_IN = .TextMatrix(intRow, mHeadCol.����)
                'Zl_���ϼƻ�����α�_Insert
                gstrSQL = "Zl_���ϼƻ�����α�_Insert("
                '  �ƻ�id_In     In ���ϼƻ�����.�ƻ�id%Type,
                gstrSQL = gstrSQL & "" & ID_IN & ","
                '  ����id_In     In ���ϼƻ�����.����id%Type,
                gstrSQL = gstrSQL & "" & ����ID_IN & ","
                '  ���_In       In ���ϼƻ�����.���%Type,
                gstrSQL = gstrSQL & "" & lng��� & ","
                '  �빺����_In   In ���ϼƻ�����.�빺����%Type,
                gstrSQL = gstrSQL & "" & dbl�빺����_IN & ","
                '  �ƻ�����_IN   In ���ϼƻ�����.�ƻ�����%Type,
                gstrSQL = gstrSQL & "" & dbl�ƻ�����_IN & ","
                '  ����_IN       In ���ϼƻ�����.����%Type,
                gstrSQL = gstrSQL & "" & dbl����_IN & ","
                '  ���_IN       In ���ϼƻ�����.���%Type,
                gstrSQL = gstrSQL & "" & dbl���_IN & ","
                '  ǰ������_In   In ���ϼƻ�����.ǰ������%Type := Null,
                gstrSQL = gstrSQL & "" & 0 & ","
                '  ��������_In   In ���ϼƻ�����.��������%Type := Null,
                gstrSQL = gstrSQL & "" & 0 & ","
                '  �������_In   In ���ϼƻ�����.�������%Type := Null,
                gstrSQL = gstrSQL & "" & 0 & ","
                '  �ϴι�Ӧ��_In In ���ϼƻ�����.�ϴι�Ӧ��%Type := Null,
                gstrSQL = gstrSQL & "'" & �ϴι�Ӧ��_IN & "',"
                '  �ϴ�������_In In ���ϼƻ�����.�ϴ�������%Type := Null
                gstrSQL = gstrSQL & "'" & �ϴ�������_IN & "')"
                cllTemp.Add gstrSQL
            End If
        Next
    End With
    On Error GoTo ErrHandle
    ExecuteProcedureArrAy cllTemp, mstrCaption
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    SaveCard = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ��ʾ�ϼƽ��()
    Dim Dbl��� As Double
    Dim intLop As Integer

    Dbl��� = 0

    With mshBill
        For intLop = 1 To .Rows - 1
            If .TextMatrix(intLop, 0) <> "" Then
                Dbl��� = Dbl��� + Val(.TextMatrix(intLop, mHeadCol.���))
            End If
        Next
    End With

    lblPurchasePrice.Caption = "���ϼƣ�" & Format(Dbl���, mFMT.FM_���)
End Sub


Private Sub txtժҪ_Change()
    mblnChange = True
End Sub

Private Sub txtժҪ_GotFocus()
    OS.OpenIme (True)
    With txtժҪ
        .SelStart = 0
        .SelLength = Len(txtժҪ.Text)
    End With
End Sub

Private Sub txtժҪ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey (vbKeyTab)
        KeyCode = 0
    End If
End Sub

Private Sub txtժҪ_LostFocus()
    OS.OpenIme False
End Sub

Private Function SetStuffRows(ByVal lng����ID As Long, ByVal str���� As String, _
        ByVal str��� As String, ByVal str���� As String, ByVal str��λ As String, _
        ByVal dblָ�������� As Double, ByVal dbl����ϵ�� As Double) As Boolean
    Dim rsData As New Recordset
    Dim intCount As Integer
    Dim intRow As Integer
    Dim intCol As Integer

    Dim lng���� As Long
    Dim dbl������� As Double
    Dim dbl�ɱ��� As Double

    On Error GoTo errH
    SetStuffRows = False

    With mshBill
        intRow = .Row
        For intCount = 1 To .Rows - 1
            If intCount <> intRow And .TextMatrix(intCount, 0) <> "" Then
                If .TextMatrix(intCount, 0) = lng����ID Then
                    If UBound(Split(mstr�ظ�����, "��")) < 3 Then mstr�ظ����� = mstr�ظ����� & str���� & "��"  '����¼�����ظ�������
                    'MsgBox "�������ϡ�" & str���� & "�������ˣ��������䣡", vbOKOnly + vbExclamation, gstrSysName
                    Exit Function
                End If
            End If
        Next

        For intCol = 0 To .Cols - 1
            .TextMatrix(intRow, intCol) = ""
        Next
    End With
    If cboEnterStock.ListIndex < 0 Then
        mlng�ⷿid = 0
    Else
        mlng�ⷿid = cboEnterStock.ItemData(cboEnterStock.ListIndex)
    End If
    With mshBill
        .TextMatrix(.Row, mHeadCol.���) = .Row
        .TextMatrix(.Row, mHeadCol.����) = str����
        .TextMatrix(.Row, 0) = lng����ID
        .TextMatrix(.Row, mHeadCol.����ϵ��) = dbl����ϵ��
        
        'ȡƽ���ɱ��ۣ����û�����ã���ȡָ�������ۣ�
        gstrSQL = "Select �ɱ���,ָ��������,�б���� From  �������� Where ����ID=[1]"
        
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�ɱ���", lng����ID)
        
        dbl�ɱ��� = zlStr.Nvl(rsData!�ɱ���, 0)
        If dbl�ɱ��� = 0 Then dbl�ɱ��� = zlStr.Nvl(rsData!ָ��������, 0)
        .TextMatrix(.Row, mHeadCol.�б����) = IIf(Val(zlStr.Nvl(rsData!�б����)) = "1", "��", "")
        
        gstrSQL = "Select a.�ϴβ���, b.���� As ��Ӧ�� From �������� A, ��Ӧ�� B Where a.�ϴι�Ӧ��id = b.Id And a.����id = [1]"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�ϴι�Ӧ�̼�������Ϣ", lng����ID)
            
        If Not rsData.EOF Then
            .TextMatrix(.Row, mHeadCol.�ϴι�Ӧ��) = IIf(IsNull(rsData!��Ӧ��), "", rsData!��Ӧ��)
            .TextMatrix(.Row, mHeadCol.����) = IIf(IsNull(rsData!�ϴβ���), str����, rsData!�ϴβ���)
        End If
        .TextMatrix(.Row, mHeadCol.����) = str����
        .TextMatrix(.Row, mHeadCol.���) = str���
        .TextMatrix(.Row, mHeadCol.��λ) = str��λ
        .TextMatrix(.Row, mHeadCol.����) = Format(dbl�ɱ��� * dbl����ϵ��, mFMT.FM_�ɱ���)
        
    End With
    rsData.Close
    SetStuffRows = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
'����ַ����Ƿ��зǷ��ַ�������ṩ���ȣ��Գ��ȵĺϷ���Ҳ����⡣
    If InStr(strInput, "'") > 0 Then
        MsgBox "���������ݺ��зǷ��ַ���", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "���������ݲ��ܳ���" & Int(intMax / 2) & "������" & "��" & intMax & "����ĸ��", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    StrIsValid = True
End Function

'�����룬���ƣ���������ĳһ��
Private Function FindData(ByVal mshBill As BillEdit, ByVal int�Ƚ��� As Integer, _
    ByVal str�Ƚ�ֵ As String, ByVal blnFirst As Boolean) As Boolean
    Dim intStartRow As Integer
    Dim intRow As Integer
    Dim strSpell As String
    Dim strCode As String
    Dim rsCode As New Recordset
    Dim strKey As String
    FindData = True
    
    On Error GoTo ErrHandle
    With mshBill
        If .Rows = 2 Then Exit Function
        If str�Ƚ�ֵ = "" Then Exit Function
        
        If blnFirst = True Then
            intStartRow = 0
        Else
            intStartRow = .Row
        End If
        If intStartRow = .Rows - 1 Then
            intStartRow = 1
        Else
            intStartRow = intStartRow + 1
        End If
        
        For intRow = intStartRow To .Rows - 1
            If .TextMatrix(intRow, int�Ƚ���) <> "" Then
                strCode = .TextMatrix(intRow, int�Ƚ���)
                If InStr(1, UCase(strCode), UCase(str�Ƚ�ֵ)) <> 0 Then
                    .SetFocus
                    .Row = intRow
                    .Col = int�Ƚ���
                    .MsfObj.TopRow = .Row
                    Exit Function
                End If
            End If
        Next
        
        gstrSQL = " SELECT DISTINCT b.���� " & _
                  " FROM (SELECT DISTINCT A.�շ�ϸĿid " & _
                  "       FROM �շ���Ŀ���� A" & _
                  "       Where A.���� LIKE [1]) a, �շ���ĿĿ¼ B " & _
                  " Where a.�շ�ϸĿid = b.ID And (b.վ��=[2] or b.վ�� is null) "
        
        strKey = IIf(gstrMatchMethod = "0", "%", "") & str�Ƚ�ֵ & "%"
        Set rsCode = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, strKey, gstrNodeNo)
                  
        If rsCode.EOF Then
            FindData = False
            Exit Function
        End If
        
        For intRow = intStartRow To .Rows - 1
            If .TextMatrix(intRow, int�Ƚ���) <> "" Then
                strCode = .TextMatrix(intRow, int�Ƚ���)
                rsCode.MoveFirst
                Do While Not rsCode.EOF
                    If InStr(1, UCase(strCode), UCase(rsCode!����)) <> 0 Then
                        .SetFocus
                        .Row = intRow
                        .Col = int�Ƚ���
                        .MsfObj.TopRow = .Row
                        rsCode.Close
                        Exit Function
                    End If
                    rsCode.MoveNext
                Loop
            
            End If
        Next
        rsCode.Close
    End With
    FindData = False
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Sub SetEdit()
    Dim intCol As Integer
    
    With mshBill
        If mblnEdit = False Then
            If mint�༭״̬ <> 3 Then
                For intCol = 0 To .Cols - 1
                    .ColData(intCol) = 0
                Next
            End If
            cboStock.Enabled = False
            cboEnterStock.Enabled = False
            txtժҪ.Enabled = False
            cbo����.Enabled = -False
        Else
            cboStock.Enabled = True
            cboEnterStock.Enabled = True
            cbo����.Enabled = True
            txtժҪ.Enabled = True
        End If
    End With
End Sub

Private Function GetDepend(ByVal lngStockID As Long) As Boolean
    Dim rsSQL As ADODB.Recordset
    Dim strվ������ As String
    
    On Error GoTo ErrHandle
    strվ������ = GetDeptStationNode(lngStockID)
    gstrSQL = "SELECT DISTINCT a.id,a.����||'-'||a.����  as ���� " & _
              "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " & _
              "Where c.�������� = b.���� " & _
              IIf(strվ������ <> "", " And (a.վ�� = [2] or a.վ�� is null) ", "") & _
              "  And b.���� In('V','K') " & _
              "  AND a.id = c.����id " & _
              "  AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' " & _
              "Order by ���� "
    Set rsSQL = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, UserInfo.Id, strվ������)
    
    If rsSQL.EOF Then
        MsgBox "û���κοⷿ�����깺�����ڲ��Ź��������ö�Ӧ���ŵĹ�������Ϊ[���Ŀ�]��[�Ƽ���]��", vbInformation, gstrSysName
        Exit Function
    End If
    
    GetDepend = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Sub SetEnterStock(ByVal lngStockID As Long)
    Dim rsSQL As ADODB.Recordset
    Dim strվ������ As String
    
    On Error GoTo ErrHandle
    strվ������ = GetDeptStationNode(lngStockID)
    gstrSQL = "SELECT DISTINCT a.id,a.����||'-'||a.����  as ���� " & _
              "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " & _
              "Where c.�������� = b.���� " & _
              IIf(strվ������ <> "", " And (a.վ�� = [2] or a.վ�� is null) ", "") & _
              "  And b.���� In('V','K') " & _
              "  AND a.id = c.����id " & _
              "  AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' " & _
              "Order by ���� "
    Set rsSQL = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, UserInfo.Id, strվ������)
    
    With cboEnterStock
        .Clear
        Do While Not rsSQL.EOF
            .AddItem zlStr.Nvl(rsSQL!����)
            .ItemData(.NewIndex) = Val(zlStr.Nvl(rsSQL!Id))
            rsSQL.MoveNext
        Loop
        rsSQL.Close
    End With

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Sub
