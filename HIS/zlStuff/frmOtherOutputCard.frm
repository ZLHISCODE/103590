VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmOtherOutputCard 
   Caption         =   "�����������ⵥ"
   ClientHeight    =   6960
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmOtherOutputCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   11400
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "ȫ��(&L)"
      Height          =   350
      Left            =   7560
      TabIndex        =   31
      Top             =   5460
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "ȫ��(&A)"
      Height          =   350
      Left            =   6240
      TabIndex        =   30
      Top             =   5460
      Visible         =   0   'False
      Width           =   1100
   End
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
      Left            =   6240
      TabIndex        =   9
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   10
      Top             =   5040
      Width           =   1100
   End
   Begin VB.PictureBox Pic���� 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   14
      Top             =   0
      Width           =   11715
      Begin VB.ComboBox cbo������λ 
         Height          =   300
         Left            =   7890
         TabIndex        =   5
         Text            =   "cbo������λ"
         Top             =   600
         Visible         =   0   'False
         Width           =   3435
      End
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   9960
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   165
         Width           =   1425
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   2115
      End
      Begin VB.TextBox txtժҪ 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   8
         Top             =   4080
         Width           =   10410
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   1515
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
      Begin VB.Label lblOther 
         AutoSize        =   -1  'True
         Caption         =   "�����ϼ�:"
         Height          =   180
         Left            =   6600
         TabIndex        =   33
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label lbl������λ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������λ(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6840
         TabIndex        =   4
         Top             =   660
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "��ۺϼ�:"
         Height          =   180
         Left            =   4920
         TabIndex        =   28
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ۼ۽��ϼ�:"
         Height          =   180
         Left            =   2040
         TabIndex        =   27
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ɱ����ϼ�:"
         Height          =   180
         Left            =   240
         TabIndex        =   26
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label Txt����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7950
         TabIndex        =   24
         Top             =   4440
         Width           =   915
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
         Width           =   915
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
         Caption         =   "���������������ⵥ"
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
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ⷿ(&S)"
         Height          =   180
         Left            =   240
         TabIndex        =   0
         Top             =   660
         Width           =   630
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
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������(&T)"
         Height          =   180
         Left            =   3480
         TabIndex        =   2
         Top             =   660
         Width           =   990
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
            Picture         =   "frmOtherOutputCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":1000
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
            Picture         =   "frmOtherOutputCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   29
      Top             =   6600
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
            Picture         =   "frmOtherOutputCard.frx":22EA
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
            Picture         =   "frmOtherOutputCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmOtherOutputCard.frx":3080
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
Attribute VB_Name = "frmOtherOutputCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbln��������    As Boolean          '����ʱ���ݺ��ۼ�1
Private mintUnit  As Integer                '��ʾ��λ:0-ɢװ��λ,1-��װ��λ
Private mblnFirst As Boolean

Private mint�༭״̬ As Integer             '1.������2���޸ģ�3�����գ�4���鿴��5
Private mstr���ݺ� As String                '����ĵ��ݺ�;
Private mint��¼״̬ As Integer             '1:������¼;2-������¼;3-�Ѿ�������ԭ��¼
Private mblnSuccess As Boolean              'ֻҪ��һ�ųɹ�����ΪTrue������ΪFalse
Private mblnSave As Boolean                 '�Ƿ���̺����   TURE���ɹ���
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '�Ƿ�����޸�
Private mblnChange As Boolean               '�Ƿ���й��༭
Private mintParallelRecord As Integer       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼
Private mint����� As Integer             '��ʾ���ĳ���ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ
Private mcolUsedCount As Collection         '��ʹ�õ���������
Dim mstrPrivs As String                     'Ȩ��

'���˺�:2007/06/10:����10813
Private mstrTime_Start As String            '���뵥�ݱ༭�ĵ���ʱ�� ,��Ҫ�ж��Ƿ񵥾ݱ����˸��Ĺ�,����༭��,���ܽ������
Private mstrTime_End As String
Private Const mlngModule = 1718
Private mblnCostView As Boolean                 '�鿴�ɱ��� true-����鿴 false-������鿴
Private mblnUpdate As Boolean               '��ʾ�Ƿ��Ѹ������¼۸���µ�������

Private mstrLike As String
Private Const mstrCaption As String = "�����������ⵥ"
Private mstr�ظ����� As String '��¼�ظ�������

Private recSort As ADODB.Recordset          '��ҩƷID�����ר�ü�¼��
'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------


'=========================================================================================
Private Const mconIntCol�к� As Integer = 1
Private Const mconIntCol���� As Integer = 2
Private Const mconIntCol��� As Integer = 3
Private Const mconIntCol��� As Integer = 4
Private Const mconIntCol�������� As Integer = 5
Private Const mconIntColָ������� As Integer = 6
Private Const mconIntColʵ�ʽ�� As Integer = 7
Private Const mconIntColʵ�ʲ�� As Integer = 8
Private Const mconIntCol����ϵ�� As Integer = 9
Private Const mconIntCol���� As Integer = 10
Private Const mconIntCol���� As Integer = 11
Private Const mconIntCol��׼�ĺ� As Integer = 12
Private Const mconIntCol��λ As Integer = 13
Private Const mconIntCol���� As Integer = 14
Private Const mconIntColЧ�� As Integer = 15
Private Const mconIntCol���ʧЧ�� As Integer = 16
Private Const mconIntCol���� As Integer = 17
Private Const mconIntCol�������� As Integer = 18
Private Const mconIntCol�ɹ��� As Integer = 19
Private Const mconIntCol�ɹ���� As Integer = 20
Private Const mconIntCol�ۼ� As Integer = 21
Private Const mconIntCol�ۼ۽�� As Integer = 22
Private Const mconintCol��� As Integer = 23
Private Const mconintCol������ As Integer = 24
Private Const mconintCol������� As Integer = 25
Private Const mconintCol��ֵ˰�� As Integer = 26
Private Const mconintCol˰�� As Integer = 27
Private Const mconIntColS  As Integer = 28              '������


'=========================================================================================


'�������������
Private Function GetDepend() As Boolean
    Dim rsTemp As New Recordset
    
    On Error GoTo ErrHandle
    GetDepend = False
    gstrSQL = "" & _
        "   SELECT B.Id " & _
        "   FROM ҩƷ�������� A, ҩƷ������ B " & _
        "   Where A.���id = B.ID AND A.���� = 36"
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "������������"
    If rsTemp.EOF Then
        ShowMsgBox "û������������������ĳ����������������������ã�"
        rsTemp.Close
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


Public Sub ShowCard(frmMain As Form, ByVal str���ݺ� As String, ByVal int�༭״̬ As Integer, _
    Optional int��¼״̬ As Integer = 1, Optional strPrivs As String, Optional blnSuccess As Boolean = False)
    Dim strReg As String
    
    mblnSave = False
    mblnSuccess = False
    mstr���ݺ� = str���ݺ�
    mint�༭״̬ = int�༭״̬
    mint��¼״̬ = int��¼״̬
    mblnSuccess = blnSuccess
    mblnChange = False
    mintParallelRecord = 1
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
    If Not GetDepend Then Exit Sub
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
        
    Call GetRegInFor(g˽��ģ��, "���������������", "���ݺ��ۼ�", strReg)
    mbln�������� = IIf(strReg = "", True, Val(strReg) = 1)
         
     
    If mint�༭״̬ = 1 Then
'        If mbln�������� Then
'            mstr���ݺ� = NextNo(74)
'        End If
        mblnEdit = True

        txtNO.Locked = True
        txtNO.TabStop = True

        txtNO = mstr���ݺ�
        txtNO.Tag = txtNO.Text
        
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
    ElseIf mint�༭״̬ = 6 Then
        CmdSave.Caption = "����(&O)"
        cmdAllCls.Visible = True
        cmdAllSel.Visible = True
    End If
      
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    Me.Show vbModal, frmMain
    blnSuccess = mblnSuccess
    str���ݺ� = mstr���ݺ�
End Sub

Private Sub cboStock_Change()
    mblnChange = True
End Sub
Private Sub cboStock_Click()
    mint����� = Get������(cboStock.ItemData(cboStock.ListIndex))
End Sub

Private Sub cboStock_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboStock_Validate False
        OS.PressKey (vbKeyTab)
    End If
End Sub

Private Sub cboStock_Validate(Cancel As Boolean)
    Dim i As Integer
    
    With cboStock
        If .ListIndex <> mintcboIndex Then
            For i = 1 To mshBill.Rows - 1
                If mshBill.TextMatrix(i, 0) <> "" Then
                    Exit For
                End If
            Next
            If i <> mshBill.Rows Then
                If MsgBox("����ı�ⷿ���п���Ҫ�ı���Ӧ���ĵĵ�λ��" & vbCrLf & "��Ҫ������е������ݣ����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '�������ĵ�λ�ı�
                    mintcboIndex = .ListIndex
                    mshBill.ClearBill
                            
                Else
                    .ListIndex = mintcboIndex
                End If
            Else
                mintcboIndex = .ListIndex
            End If
        End If
        
    End With
End Sub

Private Sub cboType_Click()
    Me.lbl������λ.Visible = False
    Me.cbo������λ.Visible = False
    
    mshBill.ColData(mconintCol������) = 5
    mshBill.ColWidth(mconintCol������) = 0
    mshBill.ColWidth(mconintCol�������) = 0
    mshBill.ColWidth(mconintCol��ֵ˰��) = 0
    mshBill.ColWidth(mconintCol˰��) = 0
        
    If cboType.Text = "��������" Then
        Me.lbl������λ.Visible = True
        Me.cbo������λ.Visible = True
        
        mshBill.ColWidth(mconintCol������) = 1000
        mshBill.ColWidth(mconintCol�������) = 1000
        mshBill.ColWidth(mconintCol��ֵ˰��) = 1000
        mshBill.ColWidth(mconintCol˰��) = 1000
        cbo������λ.Enabled = (mint�༭״̬ = 1 Or mint�༭״̬ = 2)
        mshBill.ColData(mconintCol������) = IIf(cbo������λ.Enabled, 4, 5)
    End If
End Sub

Private Sub cboType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub cbo������λ_GotFocus()
    If cbo������λ.Style = 0 Then
        Call zlControl.TxtSelAll(cbo������λ)
    End If
End Sub

Private Sub cbo������λ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If cbo������λ.Style = 2 And cbo������λ.ListIndex <> -1 Then
            cbo������λ.ListIndex = -1
        End If
    End If
End Sub


Private Sub cbo������λ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call OS.PressKey(vbKeyTab)
    ElseIf KeyAscii >= 32 Then
        If Not cbo������λ.Locked And cbo������λ.Style = 2 Then
            lngIdx = cbo.MatchIndex(cbo������λ.hwnd, KeyAscii)
            If lngIdx = -1 And cbo������λ.ListCount > 0 Then lngIdx = 0
            cbo������λ.ListIndex = lngIdx
        End If
    End If
End Sub


Private Sub cbo������λ_Validate(Cancel As Boolean)
    '���ܣ��������������,�Զ�ƥ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
        
    If cbo������λ.ListIndex <> -1 Then Exit Sub '��ѡ��
    If cbo������λ.Text = "" Then cbo������λ.Tag = "": Exit Sub '������
    
    strInput = UCase(NeedName(cbo������λ.Text))
    strSQL = "Select Rownum As id,����,����,���� From ����������λ Where Upper(����) Like [1] Or Upper(����) Like [2] Or Upper(����) Like [2] Order By ����"
        
    On Error GoTo errH
    vRect = zlControl.GetControlRect(cbo������λ.hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "�����λ", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cbo������λ.Height, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
    If Not rsTmp Is Nothing Then
        intIdx = cbo.FindIndex(cbo������λ, zlStr.Nvl(rsTmp!����) & "-" & Chr(13) & rsTmp!����)
        If intIdx <> -1 Then
            cbo������λ.ListIndex = intIdx
        Else
            cbo������λ.AddItem zlStr.Nvl(rsTmp!����) & "-" & Chr(13) & rsTmp!����, cbo������λ.ListCount - 1
            cbo������λ.ListIndex = cbo������λ.NewIndex
        End If
    Else
        If Not blnCancel Then
            MsgBox "δ�ҵ���Ӧ�������λ��", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function NeedName(strList As String) As String
    If InStr(strList, Chr(13)) > 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, Chr(13)) + 1))
    ElseIf InStr(strList, "]") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, "]") + 1))
    ElseIf InStr(strList, ")") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, ")") + 1))
    Else
        NeedName = LTrim(Mid(strList, InStr(strList, "-") + 1))
    End If
End Function

Private Sub cmdAllCls_Click()
    Dim intRow As Integer
    
    With mshBill
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                .TextMatrix(intRow, mconIntCol��������) = Format(0, mFMT.FM_����)
                .TextMatrix(intRow, mconIntCol�ɹ����) = Format(0, mFMT.FM_���)
                .TextMatrix(intRow, mconIntCol�ۼ۽��) = Format(0, mFMT.FM_���)
                .TextMatrix(intRow, mconintCol���) = Format(0, mFMT.FM_���)
            End If
        Next
    End With
    Call ��ʾ�ϼƽ��
End Sub

Private Sub cmdAllSel_Click()
    Dim intRow As Integer
    
    With mshBill
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                .TextMatrix(intRow, mconIntCol��������) = .TextMatrix(intRow, mconIntCol����)
                .TextMatrix(intRow, mconIntCol�ɹ����) = Format(.TextMatrix(intRow, mconIntCol����) * .TextMatrix(intRow, mconIntCol�ɹ���), mFMT.FM_���)
                .TextMatrix(intRow, mconIntCol�ۼ۽��) = Format(.TextMatrix(intRow, mconIntCol����) * .TextMatrix(intRow, mconIntCol�ۼ�), mFMT.FM_���)
                .TextMatrix(intRow, mconintCol���) = Format(.TextMatrix(intRow, mconIntCol�ۼ۽��) - .TextMatrix(intRow, mconIntCol�ɹ����), mFMT.FM_���)
            End If
        Next
    End With
    Call ��ʾ�ϼƽ��
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
        FindRownew mshBill, mconIntCol����, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
'    mblnChange = False
    Select Case mintParallelRecord
        Case 1
            '����
        Case 2
            If mint�༭״̬ = 6 Then
                ShowMsgBox "�õ�����û�п��Գ��������ģ����飡"
            Else
                '�����ѱ�ɾ��
                ShowMsgBox "�õ����ѱ�ɾ�������飡"
            End If
            Unload Me
            Exit Sub
        Case 3
            '�޸ĵĵ����ѱ����
            ShowMsgBox "�õ����ѱ���������ˣ����飡"
            Unload Me
            Exit Sub
    End Select
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            cmdFind_Click
        End If
    ElseIf KeyCode = vbKeyF3 Then
        FindRownew mshBill, mconIntCol����, txtCode.Text, False
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
    Dim strReg As String
    
    '�����������ݼ�
    Call SetSortRecord
    
    If mint�༭״̬ = 4 Then    '�鿴
        '��ӡ
        printbill
        '�˳�
        Unload Me
        Exit Sub
    End If
    
    If mint�༭״̬ = 3 Then        '���
        
        If Not ���ϵ������(Txt������.Caption) Then Exit Sub
        
        '���˺�:2007/06/10:����10813
        mstrTime_End = GetBillInfo(21, txtNO.Tag)
        If mstrTime_End = "" Then
            MsgBox "ע��:" & vbCrLf & "  �õ����Ѿ�����������Աɾ��,���ܼ�����", vbInformation, gstrSysName
            Exit Sub
        End If
        If mstrTime_End <> mstrTime_Start Then
            If MsgBox("ע��:" & vbCrLf & "  �õ����Ѿ�����������Ա�༭�����ܼ���!" & vbCrLf & "  �Ƿ�����ˢ�µ���?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                Call initCard
            End If
            Exit Sub
        End If
        
        If Not ��鵥��(21, txtNO.Tag, False) And Not mblnUpdate Then
            '�����µļ۸���µ����壬�˳���Ŀ�������û���һ�����յĵ���
            MsgBox "�м�¼δʹ�����¼۸񣬳����Զ���ɸ��£��ۼۡ��ɱ��ۡ��ۼ۽��ɱ�����ۣ������º����飡", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
        
        '������ʱ�޸��˵��ݣ����������ɵ��ݱ���
        If mblnChange Then
            If Not SaveCard(True) Then
                gcnOracle.RollbackTrans: Exit Sub
            End If
        End If
        
        If SaveCheck = True Then
            strReg = IIf(Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModule, "0")) = 1, 1, 0)
            If Val(strReg) = 1 Then
                '��ӡ
                If InStr(mstrPrivs, "���ݴ�ӡ") <> 0 Then
                    printbill
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If
    
    If mint�༭״̬ = 6 Then '����
        If SaveStrike Then Unload Me
        Exit Sub
    End If
            
    If ValidData = False Then Exit Sub
    blnSuccess = SaveCard
        
    If blnSuccess = True Then
        strReg = IIf(Val(zlDatabase.GetPara("���̴�ӡ", glngSys, mlngModule, "0")) = 1, 1, 0)
        If Val(strReg) = 1 Then
            '��ӡ
            If InStr(mstrPrivs, "���ݴ�ӡ") <> 0 Then
                printbill
            End If
        End If
        If mint�༭״̬ = 2 Then   '�޸�
            Unload Me
            Exit Sub
        End If
    Else
        Exit Sub
    End If
'    If mbln�������� Then
'        mstr���ݺ� = NextNo(74)
'        txtNO = mstr���ݺ�
'    End If
    
    mblnSave = False
    mblnEdit = True
    mshBill.ClearBill
    Call RefreshRowNO(mshBill, mconIntCol�к�, 1)

    txtժҪ.Text = ""
    If cboType.Enabled Then cboType.SetFocus
    mblnChange = False
    If txtNO.Tag <> "" Then Me.stbThis.Panels(2).Text = "��һ�ŵ��ݵ�NO�ţ�" & txtNO.Tag
End Sub

Private Sub RefreshBill()
    '�����¼۸����µ���������ݣ����ڵ������ʱ
    Dim lngRow As Long, lngRows As Long, lng����ID As Long
    Dim dbl���� As Double, dbl�ɱ��� As Double, dbl�ɱ���� As Double, dbl���ۼ� As Double, dbl���۽�� As Double, dbl��� As Double
    Dim rsprice As New ADODB.Recordset
    Dim rsStock As ADODB.Recordset
    Dim blnAdj As Boolean
    
    On Error GoTo ErrHandle
    
    gstrSQL = " Select '�ۼ�' As ����, a.���, a.ҩƷid As ����id, Nvl(a.����, 0) As ����, b.�ּ�" & _
            " From ҩƷ�շ���¼ A," & _
                 " (Select �շ�ϸĿid, Nvl(�ּ�, 0) �ּ�, ִ������" & _
                   " From �շѼ�Ŀ" & _
                   " Where (��ֹ���� Is Null Or Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd')))" & _
                   GetPriceClassString("") & ") B, �շ���ĿĿ¼ C" & _
            " Where a.���� = 21 And a.No = [1] And a.ҩƷid = b.�շ�ϸĿid And c.Id = b.�շ�ϸĿid And Round(a.���ۼ�," & g_С��λ��.obj_ɢװС��.���ۼ�С�� & ") <> Round(b.�ּ�, " & g_С��λ��.obj_ɢװС��.���ۼ�С�� & ") And" & _
              "    NVL(c.�Ƿ���, 0) = 0" & _
            " Union All" & _
            " Select '�ۼ�' As ����, a.���, a.ҩƷid As ����id, Nvl(a.����, 0) As ����, decode(nvl(b.����,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�) As �ּ�" & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B, �շ���ĿĿ¼ C" & _
            " Where a.���� = 21 And a.No = [1] And c.Id = a.ҩƷid And Round(a.���ۼ�," & g_С��λ��.obj_ɢװС��.���ۼ�С�� & ") <> Round(decode(nvl(b.����,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�), " & g_С��λ��.obj_ɢװС��.���ۼ�С�� & ") And Nvl(c.�Ƿ���, 0) = 1 And" & _
                  " b.���� = 1 And b.�ⷿid = a.�ⷿid And b.ҩƷid = a.ҩƷid And NVL(b.����, 0) = NVL(a.����, 0) And NVL(b.ʵ������, 0) <> 0 And a.���ϵ�� = -1" & _
            " Union All" & _
            " Select '�ɱ���' As ����, a.���, a.ҩƷid As ����id, Nvl(a.����, 0) As ����, b.ƽ���ɱ��� As �ּ�" & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B" & _
            " Where a.���� = 21 And a.No = [1] And a.ҩƷid = b.ҩƷid And Nvl(a.����, 0) = Nvl(b.����, 0) and round(a.�ɱ���," & g_С��λ��.obj_ɢװС��.�ɱ���С�� & ")<>round(b.ƽ���ɱ���," & g_С��λ��.obj_ɢװС��.�ɱ���С�� & ") And a.�ⷿid = b.�ⷿid and a.���ϵ��=-1 and b.����=1" & _
            " Order By ����, ����id, ���"

    Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[ȡ��ǰ�۸�]", CStr(Me.txtNO.Text))
    
    If rsprice.EOF Then Exit Sub
    
    lngRows = mshBill.Rows - 1
    For lngRow = 1 To lngRows
        blnAdj = False
        lng����ID = Val(mshBill.TextMatrix(lngRow, 0))
        dbl���� = Val(mshBill.TextMatrix(lngRow, mconIntCol����))
        dbl�ɱ��� = Val(mshBill.TextMatrix(lngRow, mconIntCol�ɹ���))
        dbl���ۼ� = Val(mshBill.TextMatrix(lngRow, mconIntCol�ۼ�))
        dbl�ɱ���� = dbl�ɱ��� * dbl����
        dbl���۽�� = dbl���ۼ� * dbl����
        dbl��� = dbl���۽�� - dbl�ɱ����
'
        If lng����ID <> 0 Then
            rsprice.Filter = "����='�ۼ�' And ����id=" & lng����ID & " And ����=" & Val(mshBill.TextMatrix(lngRow, mconIntCol����))
            If rsprice.RecordCount > 0 Then
                blnAdj = True
                dbl���ۼ� = Val(Format(rsprice!�ּ� * Val(mshBill.TextMatrix(lngRow, mconIntCol����ϵ��)), mFMT.FM_���ۼ�))
                dbl���۽�� = Val(Format(dbl���ۼ� * dbl����, mFMT.FM_���))
                dbl��� = Val(Format(dbl���۽�� - dbl�ɱ����, mFMT.FM_���))
            End If

            rsprice.Filter = "����='�ɱ���' And ����id=" & lng����ID & " And ����=" & Val(mshBill.TextMatrix(lngRow, mconIntCol����))
            If rsprice.RecordCount > 0 Then
                blnAdj = True
                dbl���۽�� = Val(Format(dbl���ۼ� * dbl����, mFMT.FM_���))
                dbl�ɱ��� = Val(Format(rsprice!�ּ� * Val(mshBill.TextMatrix(lngRow, mconIntCol����ϵ��)), mFMT.FM_���))
                dbl�ɱ���� = Val(Format(dbl�ɱ��� * dbl����, mFMT.FM_���))
                dbl��� = Val(Format(dbl���۽�� - dbl�ɱ����, mFMT.FM_���))
            End If

            If blnAdj = True Then
                '�Ե�ǰ���¼۸����µ���������ݣ��ۼۡ��ɱ��ۡ����۽��ɱ�����ۣ�
                mshBill.TextMatrix(lngRow, mconIntCol�ۼ�) = Format(dbl���ۼ�, mFMT.FM_���ۼ�)
                mshBill.TextMatrix(lngRow, mconIntCol�ۼ۽��) = Format(dbl���۽��, mFMT.FM_���)
                mshBill.TextMatrix(lngRow, mconIntCol�ɹ���) = Format(dbl�ɱ���, mFMT.FM_�ɱ���)
                mshBill.TextMatrix(lngRow, mconIntCol�ɹ����) = Format(dbl�ɱ����, mFMT.FM_���)
                mshBill.TextMatrix(lngRow, mconintCol���) = Format(dbl���, mFMT.FM_���)
            End If
        End If
    Next
    rsprice.Filter = 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strReg As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    mblnUpdate = False
    strReg = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngModule, "0"))
    mstrLike = IIf(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
    mintUnit = Val(strReg)
    
    mblnFirst = True
  
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
    End With
    
    initGrid
    
    With cboType
        .Clear
        gstrSQL = "" & _
            "   SELECT b.Id,b.���� " & _
            "   FROM ҩƷ�������� A, ҩƷ������ B " & _
            "   Where A.���id = B.ID AND A.���� = 36 "
        zlDatabase.OpenRecordset rsTemp, gstrSQL, mstrCaption
        Do While Not rsTemp.EOF
            .AddItem rsTemp.Fields(1)
            .ItemData(.NewIndex) = rsTemp.Fields(0)
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        .ListIndex = 0
    End With
    
    With cbo������λ
        .Clear
        gstrSQL = "Select Rownum As id,����,����,���� From ����������λ Order By ����"
        Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "��ȡ������λ")
        
        .AddItem ""
        Do While Not rsTemp.EOF
            .AddItem rsTemp!���� & "-" & rsTemp!����
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        .ListIndex = 0
    End With
    
    txtNO = mstr���ݺ�
    txtNO.Tag = txtNO.Text
    Call initCard
    
    '�ָ����Ի���������
    RestoreWinState Me, App.ProductName, mstrCaption
    '�ָ����Ի��������ú󣬻���Ҫ��Ȩ�޿��Ƶ��н�һ������
    With mshBill
        .ColWidth(mconIntCol��������) = IIf(mint�༭״̬ = 6, 800, 0)
        .ColWidth(mconintCol������) = IIf(cboType.Text = "��������", 1000, 0)
        .ColWidth(mconintCol�������) = IIf(cboType.Text = "��������", 1000, 0)
        .ColWidth(mconintCol��ֵ˰��) = IIf(cboType.Text = "��������", 1000, 0)
        .ColWidth(mconintCol˰��) = IIf(cboType.Text = "��������", 1000, 0)
        
        .ColWidth(mconIntCol�ɹ���) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(mconIntCol�ɹ����) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(mconintCol���) = IIf(mblnCostView = True, 900, 0)
    End With
    mblnChange = False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsTemp As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim numUseAbleCount As Double
    Dim vardrug As Variant
    Dim strOrder As String, strCompare As String
    Dim str���� As String, strArray As String
    
    '�ⷿ
    On Error GoTo ErrHandle
    strOrder = zlDatabase.GetPara("��������", glngSys, mlngModule, "00")
    strOrder = IIf(strOrder = "", "00", strOrder)
    
    strCompare = Mid(strOrder, 1, 1)
    If mint�༭״̬ <> 4 Then
        With mfrmMain.cboStock
            cboStock.Clear
            For i = 0 To .ListCount - 1
                cboStock.AddItem .List(i)
                cboStock.ItemData(cboStock.NewIndex) = .ItemData(i)
            Next
            mintcboIndex = .ListIndex
            cboStock.ListIndex = .ListIndex
            cboStock.Enabled = .Enabled
        End With
    End If
    
    Select Case mint�༭״̬
        Case 1
            Txt������ = UserInfo.�û���
            Txt�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        Case 2, 3, 4, 6
            initGrid
            
            If mint�༭״̬ = 4 Then
                gstrSQL = "" & _
                    "   Select b.id,b.���� " & _
                    "   From ҩƷ�շ���¼ a,���ű� b " & _
                    "   Where a.�ⷿid=b.id and A.���� = 21 and a.no=[1]"
                
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr���ݺ�)
                
                If rsTemp.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If

                With cboStock
                    .AddItem rsTemp!����
                    .ItemData(.NewIndex) = rsTemp!Id
                    .ListIndex = 0
                End With
                rsTemp.Close
            End If
            Select Case mintUnit
                Case 0
                    strUnitQuantity = "c.���㵥λ AS ��λ, A.��д����,a.ʵ������,a.�ɱ���,a.���ۼ�,nvl(a.����,0) As ������,'1' as ����ϵ��,"
                Case Else
                    strUnitQuantity = "B.��װ��λ AS ��λ,(A.��д���� / B.����ϵ��) AS ��д����,(A.ʵ������ / B.����ϵ��) AS ʵ������,a.�ɱ���*B.����ϵ�� as �ɱ���,a.���ۼ�*B.����ϵ�� as ���ۼ�,nvl(a.����,0)*B.����ϵ�� As ������,B.����ϵ�� as ����ϵ��,"
            End Select
            
            If mint�༭״̬ <> 6 Then
                    gstrSQL = "" & _
                    "   Select w.*,z.��������/w.����ϵ�� ��������,z.ʵ�ʽ��,z.ʵ�ʲ�� " & _
                    "   From (  SELECT distinct a.ҩƷid ����id,A.���,('[' || c.���� || ']' ||c.����) AS ������Ϣ," & _
                    "                   zlSpellCode(c.����) ����,c.���,c.���� as ԭ����,A.����,A.��׼�ĺ�, A.����,a.����,b.ָ�������,a.Ч��," & _
                                        strUnitQuantity & _
                    "                   A.�ɱ����,A.���۽��, A.���, " & _
                    "                   a.ժҪ,������,a.��������,a.�����,�������,a.�ⷿid,a.������id,c.�Ƿ���,b.���÷���,d.���� AS ������λ, To_Number(Trim(To_Char(Nvl(A.Ƶ��, '0'), '999999999999.0000'))) As ��ֵ˰�� " & _
                    "           FROM ҩƷ�շ���¼ A, �������� B,�շ���ĿĿ¼ c,����������λ D " & _
                    "           Where A.ҩƷid = B.����id and a.ҩƷid=c.id  " & _
                    "                   AND A.��¼״̬ =[3] " & _
                    "                   AND A.���� = 21 AND A.No = [1] And A.��ҩ����=D.����(+) " & _
                    "           ) w,(   Select  ҩƷid ����id,Nvl(����,0) ����,��������,ʵ�ʽ��,ʵ�ʲ�� " & _
                    "                   From ҩƷ��� where �ⷿid=[2]  and ����=1)  z " & _
                    "   Where w.����id=z.����id(+) and nvl(w.����,0)=nvl(z.����(+),0) " & _
                    " ORDER BY " & IIf(strCompare = "0", "���", IIf(strCompare = "1", "������Ϣ", "����")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            Else
                    gstrSQL = "" & _
                    "   Select w.*,z.��������/w.����ϵ�� ��������,z.ʵ�ʽ��,z.ʵ�ʲ�� " & _
                    "   From (  SELECT distinct a.����id,A.���,('[' || c.���� || ']' || c.����) AS ������Ϣ," & _
                    "                   zlSpellCode(c.����) ����,c.���,c.���� as ԭ����,A.����,A.��׼�ĺ�, A.����,a.����,b.ָ�������,a.Ч��," & _
                                        strUnitQuantity & _
                    "                   A.�ɱ����,0 ���۽��,0 ���, " & _
                    "                   a.ժҪ,a.�ⷿid,a.������id,c.�Ƿ���,b.���÷���,d.���� AS ������λ,A.��ֵ˰�� " & _
                    "           FROM (  Select min(id) as id, sum(ʵ������) as ��д����,0 ʵ������,sum(�ɱ����) as �ɱ����,ҩƷid ����ID,���,����,��׼�ĺ�, ����,Ч��," & _
                    "           Nvl(����,0) ����,����,�ɱ���,���ۼ�,ժҪ,�ⷿID,������ID,����,��ҩ����, To_Number(Trim(To_Char(Nvl(Ƶ��, '0'), '999999999999.0000'))) As ��ֵ˰��" & _
                    "                   From ҩƷ�շ���¼ x " & _
                    "                   WHERE NO=[1] AND ����=21  " & _
                    "                   Group by ҩƷID,���,����,��׼�ĺ�,����,Ч��,Nvl(����,0),����,�ɱ���,���ۼ�,ժҪ,�ⷿID,�Է�����ID,������ID,����,��ҩ����, To_Number(Trim(To_Char(Nvl(Ƶ��, '0'), '999999999999.0000'))) " & _
                    "                   having sum(��д����)<>0 " & _
                    "               ) A, �������� B,�շ���ĿĿ¼ c,����������λ d " & _
                    "           Where A.����id = B.����id and a.����id=c.id And A.��ҩ����=d.����(+) " & _
                    "       ) w,(Select  ҩƷid ����id,Nvl(����,0) ����,��������,ʵ�ʽ��,ʵ�ʲ�� " & _
                    "            From ҩƷ��� " & _
                    "            Where �ⷿid=[2]  and ����=1)  z " & _
                    "   Where w.����id=z.����id(+) and nvl(w.����,0)=nvl(z.����(+),0) " & _
                    "   ORDER BY " & IIf(strCompare = "0", "���", IIf(strCompare = "1", "������Ϣ", "����")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
                    
            End If
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr���ݺ�, cboStock.ItemData(cboStock.ListIndex), mint��¼״̬)
            If rsTemp.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            '���˺�:2007/06/10:����10813
            mstrTime_Start = GetBillInfo(21, mstr���ݺ�)
            
            Dim intCount As Integer
            With cboType
                For intCount = 0 To .ListCount - 1
                    If .ItemData(intCount) = rsTemp!������ID Then
                        .ListIndex = intCount
                        Exit For
                    End If
                Next
                
                If .Text = "��������" Then
                    Me.cbo������λ.Visible = True
                    
                    '��λ������λ
                    If Not IsNull(rsTemp!������λ) Then
                        For i = 1 To cbo������λ.ListCount - 1
                            If Mid(cbo������λ.List(i), InStr(1, cbo������λ.List(i), "-") + 1) = rsTemp!������λ Then
                                cbo������λ.ListIndex = i
                                Exit For
                            End If
                        Next
                    End If
                End If
            End With
            
            Select Case mint�༭״̬
            Case 2, 6
                Txt������ = UserInfo.�û���
                Txt�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                If mint�༭״̬ = 6 Then
                    Txt����� = UserInfo.�û���
                    Txt������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                End If
            Case Else
                Txt������ = rsTemp!������
                Txt�������� = Format(rsTemp!��������, "yyyy-mm-dd hh:mm:ss")
                Txt����� = IIf(IsNull(rsTemp!�����), "", rsTemp!�����)
                Txt������� = IIf(IsNull(rsTemp!�������), "", Format(rsTemp!�������, "yyyy-mm-dd hh:mm:ss"))
            End Select
            txtժҪ.Text = IIf(IsNull(rsTemp!ժҪ), "", rsTemp!ժҪ)
            
            If (mint�༭״̬ = 2 Or mint�༭״̬ = 3) And Txt����� <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            If mint�༭״̬ = 2 Or mint�༭״̬ = 3 Then
                Set mcolUsedCount = New Collection
            End If
            
            intRow = 0
            With mshBill
                Do While Not rsTemp.EOF
                    
                    intRow = intRow + 1
                    .Rows = intRow + 1
                    .TextMatrix(intRow, 0) = rsTemp.Fields(0)
                    .TextMatrix(intRow, mconIntCol����) = rsTemp!������Ϣ
                    .TextMatrix(intRow, mconIntCol���) = rsTemp!���
                    .TextMatrix(intRow, mconIntCol���) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                    .TextMatrix(intRow, mconIntCol��׼�ĺ�) = IIf(IsNull(rsTemp!��׼�ĺ�), "", rsTemp!��׼�ĺ�)
                    .TextMatrix(intRow, mconIntCol��λ) = rsTemp!��λ
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                    .TextMatrix(intRow, mconIntColЧ��) = IIf(IsNull(rsTemp!Ч��), "", Format(rsTemp!Ч��, "yyyy-mm-dd"))
                    .TextMatrix(intRow, mconIntCol����) = Format(rsTemp!��д����, mFMT.FM_����)
                    .TextMatrix(intRow, mconIntCol�ɹ���) = Format(rsTemp!�ɱ���, mFMT.FM_�ɱ���)
                    .TextMatrix(intRow, mconIntCol�ɹ����) = Format(IIf(mint�༭״̬ = 6, 0, rsTemp!�ɱ����), mFMT.FM_���)
                    .TextMatrix(intRow, mconIntCol�ۼ�) = Format(rsTemp!���ۼ�, mFMT.FM_���ۼ�)
                    .TextMatrix(intRow, mconIntCol�ۼ۽��) = Format(rsTemp!���۽��, mFMT.FM_���)
                    .TextMatrix(intRow, mconintCol���) = Format(rsTemp!���, mFMT.FM_���)
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsTemp!����), "0", rsTemp!����)
                    .TextMatrix(intRow, mconIntCol����ϵ��) = rsTemp!����ϵ��
                    .TextMatrix(intRow, mconIntColָ�������) = rsTemp!ָ������� & "||" & rsTemp!�Ƿ��� & "||" & rsTemp!���÷���
                    .TextMatrix(intRow, mconIntCol��������) = IIf(IsNull(rsTemp!��������), "0", rsTemp!��������)
                    .TextMatrix(intRow, mconIntColʵ�ʲ��) = IIf(IsNull(rsTemp!ʵ�ʲ��), "0", rsTemp!ʵ�ʲ��)
                    .TextMatrix(intRow, mconIntColʵ�ʽ��) = IIf(IsNull(rsTemp!ʵ�ʽ��), "0", rsTemp!ʵ�ʽ��)
                    
                    .TextMatrix(intRow, mconintCol������) = Format(rsTemp!������, mFMT.FM_���ۼ�)
                    .TextMatrix(intRow, mconintCol��ֵ˰��) = GetFormat(IIf(IsNull(rsTemp!��ֵ˰��), "0", rsTemp!��ֵ˰��), 2)
                    
                    If mint�༭״̬ = 6 Then
                        .TextMatrix(intRow, mconintCol�������) = Format(0, mFMT.FM_���)
                        .TextMatrix(intRow, mconintCol˰��) = Format(0, mFMT.FM_���)
                    Else
                        .TextMatrix(intRow, mconintCol�������) = Format(rsTemp!������ * rsTemp!��д����, mFMT.FM_���)
                        .TextMatrix(intRow, mconintCol˰��) = Format(rsTemp!������ * rsTemp!��д���� * (Val(.TextMatrix(intRow, mconintCol��ֵ˰��)) / 100 / (1 + Val(.TextMatrix(intRow, mconintCol��ֵ˰��)) / 100)), mFMT.FM_���)
                    End If
                    
                    If mint�༭״̬ = 2 Or mint�༭״̬ = 3 Then
                        numUseAbleCount = 0
                        For Each vardrug In mcolUsedCount
                            If vardrug(0) = CStr(rsTemp!����ID & IIf(IsNull(rsTemp!����), "0", rsTemp!����)) Then
                                numUseAbleCount = vardrug(1)
                                mcolUsedCount.Remove vardrug(0)
                                Exit For
                            End If
                        Next
                        str���� = rsTemp!����ID & IIf(IsNull(rsTemp!����), "0", rsTemp!����)
                        If mint�༭״̬ = 2 Then
                            strArray = numUseAbleCount + IIf(IsNull(rsTemp!��д����), "0", rsTemp!��д����)
                        Else
                            strArray = numUseAbleCount + IIf(IsNull(rsTemp!ʵ������), "0", rsTemp!ʵ������)
                        End If
                        mcolUsedCount.Add Array(str����, strArray), str����
                    End If
                    
                    rsTemp.MoveNext
                Loop
                .Rows = intRow + 2
            End With
            rsTemp.Close
    End Select
    Call RefreshRowNO(mshBill, mconIntCol�к�, 1)
    Call ��ʾ�ϼƽ��
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = mconIntColS
        
        .MsfObj.FixedCols = 1
        .TextMatrix(0, mconIntCol�к�) = ""
        .TextMatrix(0, mconIntCol����) = "���������"
        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntCol��׼�ĺ�) = "��׼�ĺ�"
        .TextMatrix(0, mconIntCol��λ) = "��λ"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntColЧ��) = "ʧЧ��"
        .TextMatrix(0, mconIntCol���ʧЧ��) = "���ʧЧ��"
        
        .TextMatrix(0, mconIntCol����) = IIf(mint�༭״̬ = 6, "����", "��д����")
        .TextMatrix(0, mconIntCol��������) = "��������"
        .TextMatrix(0, mconIntCol�ɹ���) = "�ɱ���"
        .TextMatrix(0, mconIntCol�ɹ����) = "�ɱ����"
        .TextMatrix(0, mconIntCol�ۼ�) = "�ۼ�"
        .TextMatrix(0, mconIntCol�ۼ۽��) = "�ۼ۽��"
        .TextMatrix(0, mconintCol���) = "���"
        .TextMatrix(0, mconIntCol��������) = "��������"
        .TextMatrix(0, mconIntColʵ�ʲ��) = "ʵ�ʲ��"
        .TextMatrix(0, mconIntColʵ�ʽ��) = "ʵ�ʽ��"
        .TextMatrix(0, mconIntColָ�������) = "ָ�������"
        .TextMatrix(0, mconIntCol����ϵ��) = "����ϵ��"
        .TextMatrix(0, mconIntCol����) = "����"
        
        .TextMatrix(0, mconintCol������) = "������"
        .TextMatrix(0, mconintCol�������) = "�������"
        .TextMatrix(0, mconintCol��ֵ˰��) = "��ֵ˰��%"
        .TextMatrix(0, mconintCol˰��) = "˰��"
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mconIntCol�к�) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mconIntCol�к�) = 300
        .ColWidth(mconIntCol����) = 2000
        .ColWidth(mconIntCol���) = 0
        .ColWidth(mconIntCol���) = 900
        .ColWidth(mconIntCol����) = 800
        .ColWidth(mconIntCol��׼�ĺ�) = 1000
        .ColWidth(mconIntCol��λ) = 500
        .ColWidth(mconIntCol����) = 800
        .ColWidth(mconIntColЧ��) = 1000
        .ColWidth(mconIntCol���ʧЧ��) = 1000
        .ColWidth(mconIntCol����) = 800
        .ColWidth(mconIntCol��������) = 0
        .ColWidth(mconIntCol�ɹ���) = IIf(mblnCostView = False, 0, 800)
        .ColWidth(mconIntCol�ɹ����) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mconIntCol�ۼ�) = 800
        .ColWidth(mconIntCol�ۼ۽��) = 900
        .ColWidth(mconintCol���) = IIf(mblnCostView = False, 0, 800)
        
        .ColWidth(mconIntCol��������) = 0
        
        .ColWidth(mconIntColʵ�ʲ��) = 0
        .ColWidth(mconIntColʵ�ʽ��) = 0
        .ColWidth(mconIntColָ�������) = 0
        .ColWidth(mconIntCol����ϵ��) = 0
        .ColWidth(mconIntCol����) = 0
         
        .ColWidth(mconintCol������) = 0
        .ColWidth(mconintCol�������) = 0
        .ColWidth(mconintCol��ֵ˰��) = 0
        .ColWidth(mconintCol˰��) = 0
        
        
        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��

        .ColData(0) = 5
        .ColData(mconIntCol�к�) = 5
        .ColData(mconIntCol���) = 5
        .ColData(mconIntCol���) = 5
        .ColData(mconIntCol����) = 5
        .ColData(mconIntCol��׼�ĺ�) = 5
        .ColData(mconIntCol��λ) = 5
        .ColData(mconIntCol����) = 5
        .ColData(mconIntColЧ��) = 5
        .ColData(mconIntCol���ʧЧ��) = 5
        
        .ColData(mconintCol�������) = 5
        .ColData(mconintCol��ֵ˰��) = 5
        .ColData(mconintCol˰��) = 5
        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
            cboType.Enabled = True
            txtժҪ.Enabled = True
            cboStock.Enabled = True
            .ColData(mconIntCol����) = 1
            .ColData(mconIntCol����) = 4
            .ColData(mconIntCol��������) = 5
            
            .ColData(mconintCol������) = IIf(Me.cbo������λ.Visible, 4, 5)
        ElseIf mint�༭״̬ = 3 Or mint�༭״̬ = 6 Then
            cboStock.Enabled = False
            
            cboType.Enabled = False
            txtժҪ.Enabled = False
            
            .ColData(mconIntCol����) = 5
            .ColData(mconIntCol��������) = 4
        ElseIf mint�༭״̬ = 4 Then
            cboStock.Enabled = False
            
            cboType.Enabled = False
            txtժҪ.Enabled = False
            
            .ColData(mconIntCol����) = 5
            .ColData(mconIntCol��������) = 5
            
        End If
        
        .ColData(mconIntCol�ɹ���) = 5
        .ColData(mconIntCol�ɹ����) = 5
        .ColData(mconIntCol�ۼ�) = 5
        .ColData(mconIntCol�ۼ۽��) = 5
        .ColData(mconintCol���) = 5
        
        .ColData(mconIntCol��������) = 5
        
        .ColData(mconIntColʵ�ʲ��) = 5
        .ColData(mconIntColʵ�ʽ��) = 5
        .ColData(mconIntColָ�������) = 5
        .ColData(mconIntCol����ϵ��) = 5
        .ColData(mconIntCol����) = 5
        
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntCol���) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��λ) = flexAlignCenterCenter
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntColЧ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����) = flexAlignRightCenter
        .ColAlignment(mconIntCol��������) = flexAlignRightCenter
        
        .ColAlignment(mconIntCol�ɹ���) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ɹ����) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ�) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ۽��) = flexAlignRightCenter
        .ColAlignment(mconintCol���) = flexAlignRightCenter
        .ColAlignment(mconIntCol���ʧЧ��) = flexAlignCenterCenter
        
        .PrimaryCol = mconIntCol����
        .LocateCol = mconIntCol����
        If InStr(1, "34", mint�༭״̬) <> 0 Then .ColData(mconIntCol����) = 0
    End With
    txtժҪ.MaxLength = sys.FieldsLength("ҩƷ�շ���¼", "ժҪ")
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
        lblNO.Left = .Left - lblNO.Width - 100
        .Top = LblTitle.Top
        lblNO.Top = .Top
    End With
    
    
    LblStock.Left = mshBill.Left
    cboStock.Left = LblStock.Left + LblStock.Width + 100
    
    cbo������λ.Left = mshBill.Left + mshBill.Width - cbo������λ.Width
    lbl������λ.Left = cbo������λ.Left - lbl������λ.Width - 100
    
    lblType.Left = cboStock.Left + cboStock.Width + (lbl������λ.Left - cboStock.Left - cboStock.Width - (lblType.Width + cboType.Width + 100)) / 2
    cboType.Left = lblType.Left + lblType.Width + 100
    
'    cboType.Left = mshBill.Left + mshBill.Width - cboType.Width
'    lblType.Left = cboType.Left - lblType.Width - 100
    
    
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
    
    With lbl�����
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
        lblSalePrice.Top = .Top
        lblDifference.Top = .Top
        lblOther.Top = .Top
    End With
    If mblnCostView = False Then
        lblPurchasePrice.Visible = False
    End If
    
    With lblSalePrice
        .Left = lblPurchasePrice.Left + mshBill.Width / 4
    End With
    With lblDifference
        .Left = lblPurchasePrice.Left + mshBill.Width / 4 * 2
    End With
    If mblnCostView = False Then
        lblDifference.Visible = False
    End If
    
    With lblOther
        .Left = lblPurchasePrice.Left + mshBill.Width / 4 * 3
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
    
    With cmdAllCls
        .Left = CmdSave.Left - .Width - 500
        .Top = CmdCancel.Top
    End With
    
    With cmdAllSel
        .Left = cmdAllCls.Left - .Width - 100
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
    Dim intRow As Integer
    Dim strNo As String
    Dim lng�ⷿID As Long
    Dim str����� As String
    Dim dat������� As String
    
    Dim int��� As Integer
    Dim lng����ID As Long
    Dim lng���� As Long
    Dim dbl���� As Double
    Dim dbl�ɱ��� As Double
    Dim dbl�ɱ���� As Double
    Dim dbl���۽�� As Double
    Dim dbl��� As Double
    Dim lng������ID As Long
    Dim n As Long
    Dim arrSQL As Variant
    
    
    arrSQL = Array()
    
    mblnSave = False
    SaveCheck = False
    
    lng�ⷿID = cboStock.ItemData(cboStock.ListIndex)
    lng������ID = cboType.ItemData(cboType.ListIndex)
    str����� = UserInfo.�û���
    strNo = txtNO.Tag
    
    dat������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
    With mshBill
        On Error GoTo ErrHandle
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
'                If Val(.TextMatrix(intRow, mconIntColʵ�ʽ��)) = 0 Then
'                   .TextMatrix(intRow, mconintCol���) = Format(Val(.TextMatrix(intRow, mconIntCol�ۼ۽��)) * Split(.TextMatrix(.Row, mconIntColָ�������), "||")(0) / 100, mFMT.FM_���)
'                Else
'                   .TextMatrix(intRow, mconintCol���) = Format(Val(.TextMatrix(intRow, mconIntCol�ۼ۽��)) * (Val(.TextMatrix(intRow, mconIntColʵ�ʲ��)) / Val(.TextMatrix(intRow, mconIntColʵ�ʽ��))), mFMT.FM_���)
'                End If
'
'                If Val(.TextMatrix(intRow, mconIntCol����)) = 0 Then
'                    .TextMatrix(intRow, mconIntCol�ɹ���) = 0
'                Else
'                    .TextMatrix(intRow, mconIntCol�ɹ���) = Format((Val(.TextMatrix(intRow, mconIntCol�ۼ۽��)) - Val(.TextMatrix(intRow, mconintCol���))) / (Val(.TextMatrix(intRow, mconIntCol����))), mFMT.FM_�ɱ���)
'                End If
'
'                .TextMatrix(intRow, mconIntCol�ɹ����) = Format(Val(.TextMatrix(intRow, mconIntCol�ɹ���)) * Val(.TextMatrix(intRow, mconIntCol����)), mFMT.FM_���)
                
                lng����ID = Val(.TextMatrix(intRow, 0))
                lng���� = Val(.TextMatrix(intRow, mconIntCol����))
                dbl���� = Round(Val(.TextMatrix(intRow, mconIntCol����)) * Val(.TextMatrix(intRow, mconIntCol����ϵ��)), g_С��λ��.obj_ɢװС��.����С��)
                dbl�ɱ��� = Round(Val(.TextMatrix(intRow, mconIntCol�ɹ���)) / Val(.TextMatrix(intRow, mconIntCol����ϵ��)), g_С��λ��.obj_ɢװС��.�ɱ���С��)
                dbl�ɱ���� = Round(Val(.TextMatrix(intRow, mconIntCol�ɹ����)), g_С��λ��.obj_ɢװС��.���С��)
                dbl���۽�� = Round(Val(.TextMatrix(intRow, mconIntCol�ۼ۽��)), g_С��λ��.obj_ɢװС��.���С��)
                dbl��� = Round(Val(.TextMatrix(intRow, mconintCol���)), g_С��λ��.obj_ɢװС��.���С��)
                int��� = Val(.TextMatrix(intRow, mconIntCol���))
                         
                'zl_������������_VERIFY( /*NO_IN*/, /*�ⷿID_IN*/, /*ҩƷID_IN*/, /*����_IN*/,
                    '/*ʵ������_IN*/, /*�ɱ���_IN*/, /*�ɱ����_IN*/, /*���۽��_IN*/,
                    '/*���_IN*/, /*������ID_IN*/, /*�����_IN*/, /*�������_IN*/ );
                         
                gstrSQL = "zl_������������_Verify(" & _
                    int��� & ",'" & _
                    strNo & "'," & _
                    lng�ⷿID & "," & _
                    lng����ID & "," & _
                    lng���� & "," & _
                    dbl���� & "," & _
                    dbl�ɱ��� & "," & _
                    dbl�ɱ���� & "," & _
                    dbl���۽�� & "," & _
                    dbl��� & "," & _
                    lng������ID & ",'" & _
                    str����� & "',to_date('" & _
                    dat������� & "','yyyy-mm-dd HH24:MI:SS'),1)"
                    
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = CStr(lng����ID) & ";" & vbCrLf & gstrSQL
            End If
            
            recSort.MoveNext
        Next
    End With
    
    If Not ExecuteSql(arrSQL, mstrCaption, False) Then Exit Function
'    If Not ��鵥��(21, txtNO.Tag) Then
'        gcnOracle.RollbackTrans
'        Exit Function
'    End If
    gcnOracle.CommitTrans
    
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function

Private Function SaveStrike() As Boolean
    Dim �д�_IN As Integer
    Dim ԭ��¼״̬_IN As Integer
    Dim NO_IN As String
    Dim ���_IN As Integer
    Dim ����ID_IN As Long
    Dim ��������_IN As Double
    Dim ������_IN As String
    Dim ��������_IN  As String
    Dim intRow As Integer
    Dim n As Long
    
    SaveStrike = False
    With mshBill
        '����������������С����
        For intRow = 1 To .Rows - 1
            If Val(.TextMatrix(intRow, mconIntCol��������)) <> 0 Then
                If Not ��ͬ����(Val(.TextMatrix(intRow, mconIntCol����)), Val(.TextMatrix(intRow, mconIntCol��������))) Then
                    ShowMsgBox "������Ϸ��ĳ�����������" & intRow & "�У���"
                    Exit Function
                End If
            End If
        Next
        
        NO_IN = Trim(txtNO.Tag)
        ������_IN = UserInfo.�û���
        ��������_IN = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        ԭ��¼״̬_IN = mint��¼״̬
        
        On Error GoTo ErrHandle
        gcnOracle.BeginTrans
        
        �д�_IN = 0
        Dim blnȫ�� As Boolean, dblʵ������ As Double
        
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" And Val(.TextMatrix(intRow, mconIntCol��������)) <> 0 Then
                �д�_IN = �д�_IN + 1
                
                ����ID_IN = .TextMatrix(intRow, 0)
                ��������_IN = Format(.TextMatrix(intRow, mconIntCol��������) * .TextMatrix(intRow, mconIntCol����ϵ��), mFMT.FM_����)
                ���_IN = .TextMatrix(intRow, mconIntCol���)
                dblʵ������ = Val(Format(Val(.TextMatrix(intRow, mconIntCol����)) * Val(.TextMatrix(intRow, mconIntCol����ϵ��)), mFMT.FM_����))
                blnȫ�� = (��������_IN = dblʵ������)
                
                'ZL_������������_STRIKE(
                '    �д�_In       In Integer,
                '    ԭ��¼״̬_In In ҩƷ�շ���¼.��¼״̬%Type,
                '    No_In         In ҩƷ�շ���¼.NO%Type,
                '    ���_In       In ҩƷ�շ���¼.���%Type,
                '    ����id_In     In ҩƷ�շ���¼.ҩƷid%Type,
                '    ��������_In   In ҩƷ�շ���¼.ʵ������%Type,
                '    ������_In     In ҩƷ�շ���¼.������%Type,
                '    ��������_In   In ҩƷ�շ���¼.��������%Type,
                '    ȫ������_In   In ҩƷ�շ���¼.ʵ������%Type := 0 --1-ȫ������,0-���ֳ���
                
                gstrSQL = "ZL_������������_STRIKE(" & _
                    �д�_IN & "," & _
                    ԭ��¼״̬_IN & ",'" & _
                    NO_IN & "'," & _
                    ���_IN & "," & _
                    ����ID_IN & "," & _
                    ��������_IN & ",'" & _
                    ������_IN & "',to_date('" & _
                    Format(��������_IN, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS')," & IIf(blnȫ��, 1, 0) & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
            End If
            
            recSort.MoveNext
        Next
        
        gcnOracle.CommitTrans
        
        If �д�_IN = 0 Then
            ShowMsgBox "û��ѡ��һ�����������������ܳ��������飡"
            Exit Function
        End If
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveStrike = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Function

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mconIntCol�к�, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call RefreshRowNO(mshBill, mconIntCol�к�, mshBill.Row)
End Sub

Private Sub mshBill_BeforeAddRow(Row As Long)
    If mshBill.ColData(mconIntCol����) = 0 Then
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
            If MsgBox("��ȷʵҪɾ���������ģ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub mshbill_CommandClick()
    Dim RecReturn As Recordset
    Dim i As Integer
    Dim int����� As Integer
    
    int����� = mshBill.Row
    
    
    Set RecReturn = Frm����ѡ����.ShowMe(Me, 2, cboStock.ItemData(cboStock.ListIndex), , cboStock.ItemData(cboStock.ListIndex), , , , , , , , , , , , , mstrPrivs, , False)
    If RecReturn.RecordCount > 0 Then
    
        With mshBill
            mblnChange = True
            RecReturn.MoveFirst
            For i = 1 To RecReturn.RecordCount
                If SetColValue(.Row, RecReturn!����ID, "[" & RecReturn!���� & "]" & RecReturn!����, _
                    IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                    IIf(mintUnit = 0, RecReturn!ɢװ��λ, RecReturn!��װ��λ), _
                    IIf(IsNull(RecReturn!�ۼ�), 0, RecReturn!�ۼ�), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                    IIf(IsNull(RecReturn!Ч��), "", Format(RecReturn!Ч��, "yyyy-MM-dd")), _
                    IIf(IsNull(RecReturn!���ʧЧ��), "", Format(RecReturn!���ʧЧ��, "yyyy-MM-dd")), _
                    IIf(IsNull(RecReturn!��������), "0", RecReturn!��������), _
                    IIf(IsNull(RecReturn!ʵ�ʽ��), "0", RecReturn!ʵ�ʽ��), _
                    IIf(IsNull(RecReturn!ʵ�ʲ��), "0", RecReturn!ʵ�ʲ��), _
                    IIf(IsNull(RecReturn!ָ�������), "0", RecReturn!ָ�������), _
                    IIf(mintUnit = 0, 1, RecReturn!����ϵ��), IIf(IsNull(RecReturn!����), 0, RecReturn!����), RecReturn!ʱ��, RecReturn!���÷���, IIf(IsNull(RecReturn!��׼�ĺ�), "", RecReturn!��׼�ĺ�)) Then
                    
                    If .Row = .Rows - 1 Then .Rows = .Rows + 1 'ֻ�е�ǰ�������һ��ʱ��������
                    .Row = .Row + 1
                    
                End If
                .Col = mconIntCol����
                RecReturn.MoveNext
            Next
            
            mshBill.Row = int�����
            
            If mstr�ظ����� <> "" Then
                MsgBox mstr�ظ����� & "�б����Ѿ������ˣ�" & vbCrLf & "�������Ĳ�����ӣ�", vbInformation + vbOKOnly, gstrSysName
                mstr�ظ����� = ""
            End If
                
'            If RecReturn.RecordCount = 1 Then
'                SetColValue .Row, RecReturn!����ID, "[" & RecReturn!���� & "]" & RecReturn!����, _
'                    IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
'                    IIf(mintUnit = 0, RecReturn!ɢװ��λ, RecReturn!��װ��λ), _
'                    IIf(IsNull(RecReturn!�ۼ�), 0, RecReturn!�ۼ�), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
'                    IIf(IsNull(RecReturn!Ч��), "", Format(RecReturn!Ч��, "yyyy-MM-dd")), _
'                    IIf(IsNull(RecReturn!���ʧЧ��), "", Format(RecReturn!���ʧЧ��, "yyyy-MM-dd")), _
'                    IIf(IsNull(RecReturn!��������), "0", RecReturn!��������), _
'                    IIf(IsNull(RecReturn!ʵ�ʽ��), "0", RecReturn!ʵ�ʽ��), _
'                    IIf(IsNull(RecReturn!ʵ�ʲ��), "0", RecReturn!ʵ�ʲ��), _
'                    IIf(IsNull(RecReturn!ָ�������), "0", RecReturn!ָ�������), _
'                    IIf(mintUnit = 0, 1, RecReturn!����ϵ��), IIf(IsNull(RecReturn!����), 0, RecReturn!����), RecReturn!ʱ��, RecReturn!���÷���, IIf(IsNull(RecReturn!��׼�ĺ�), "", RecReturn!��׼�ĺ�)
'                .Col = mconIntCol����
'            End If
        End With
        RecReturn.Close
    End If
End Sub

Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub

Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer
    
    With mshBill
        If .Col = mconIntCol���� Or .Col = mconIntCol�������� Then
            strKey = .Text
            If strKey = "" Then
                strKey = .TextMatrix(.Row, .Col)
            End If
            Select Case .Col
                Case mconIntCol����, mconIntCol��������
                    intDigit = IIf(mintUnit = 1, g_С��λ��.obj_��װС��.����С��, g_С��λ��.obj_ɢװС��.����С��)
            End Select
            
            If InStr(strKey, ".") <> 0 And Chr(KeyAscii) = "." Then   'ֻ�ܴ���һ��С����
                KeyAscii = 0
                Exit Sub
            End If
            
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
                If .SelLength = Len(strKey) Then Exit Sub
                If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= intDigit And strKey Like "*.*" Then
                    KeyAscii = 0
                    Exit Sub
                Else
                    Exit Sub
                End If
            End If
        End If
    End With
End Sub

Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If
        
        Select Case .Col
            Case mconIntCol����
                .TxtCheck = False
                .MaxLength = 80
                'ֻ��ҩ���в���ʾ�ϼ���Ϣ�Ϳ����
                Call ��ʾ�ϼƽ��
                Call ��ʾ�����
            Case mconIntCol����, mconIntCol��������
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = "-.1234567890"
        End Select
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsDrug As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim i As Integer
    Dim int����� As Integer
    
    int����� = mshBill.Row
    
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With mshBill
        .Text = UCase(Trim(.Text))
        strKey = UCase(Trim(.Text))
        
        If Mid(strKey, 1, 1) = "[" Then
            If InStr(2, strKey, "]") <> 0 Then
                strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
            Else
                strKey = Mid(strKey, 2)
            End If
        End If
        Select Case .Col
            
            Case mconIntCol����
                If strKey <> "" Then
                    Dim RecReturn As Recordset
                    Dim sngLeft As Single
                    Dim sngTop As Single
                    
                    sngLeft = Me.Left + Pic����.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic����.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 4530
                    End If
                    
                    Set RecReturn = FrmMulitSel.ShowSelect(Me, 2, cboStock.ItemData(cboStock.ListIndex), , cboStock.ItemData(cboStock.ListIndex), strKey, sngLeft, sngTop, mshBill.MsfObj.CellWidth, mshBill.MsfObj.CellHeight, , , , , , , , , , , , mstrPrivs, , False)
                    
                    If RecReturn.RecordCount <= 0 Then
                        Cancel = True
                        Exit Sub
                    End If
                    
                    RecReturn.MoveFirst
                    For i = 1 To RecReturn.RecordCount
                        If SetColValue(.Row, RecReturn!����ID, "[" & RecReturn!���� & "]" & RecReturn!����, _
                                IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                                IIf(mintUnit = 0, RecReturn!ɢװ��λ, RecReturn!��װ��λ), _
                                IIf(IsNull(RecReturn!�ۼ�), 0, RecReturn!�ۼ�), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                                IIf(IsNull(RecReturn!Ч��), "", Format(RecReturn!Ч��, "yyyy-MM-dd")), _
                                IIf(IsNull(RecReturn!���ʧЧ��), "", Format(RecReturn!���ʧЧ��, "yyyy-MM-dd")), _
                                IIf(IsNull(RecReturn!��������), "0", RecReturn!��������), _
                                IIf(IsNull(RecReturn!ʵ�ʽ��), "0", RecReturn!ʵ�ʽ��), _
                                IIf(IsNull(RecReturn!ʵ�ʲ��), "0", RecReturn!ʵ�ʲ��), _
                                IIf(IsNull(RecReturn!ָ�������), "0", RecReturn!ָ�������), _
                                IIf(mintUnit = 0, 1, RecReturn!����ϵ��), IIf(IsNull(RecReturn!����), 0, RecReturn!����), RecReturn!ʱ��, RecReturn!���÷���, IIf(IsNull(RecReturn!��׼�ĺ�), "", RecReturn!��׼�ĺ�)) Then
                            
                            If .Row = .Rows - 1 Then .Rows = .Rows + 1 'ֻ�е�ǰ�������һ��ʱ��������
                            .Row = .Row + 1
                            
                            .Text = .TextMatrix(.Row, .Col)
                        Else
                            Cancel = True
                        End If
                        
                        RecReturn.MoveNext
                    Next
                    
                    mshBill.Row = int�����
                    
                    If mstr�ظ����� <> "" Then
                        MsgBox mstr�ظ����� & "�б����Ѿ������ˣ�" & vbCrLf & "�������Ĳ�����ӣ�", vbInformation + vbOKOnly, gstrSysName
                        mstr�ظ����� = ""
                    End If
                    
'                    If RecReturn.RecordCount = 1 Then
'                        If SetColValue(.Row, RecReturn!����ID, "[" & RecReturn!���� & "]" & RecReturn!����, _
'                                IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
'                                IIf(mintUnit = 0, RecReturn!ɢװ��λ, RecReturn!��װ��λ), _
'                                IIf(IsNull(RecReturn!�ۼ�), 0, RecReturn!�ۼ�), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
'                                IIf(IsNull(RecReturn!Ч��), "", Format(RecReturn!Ч��, "yyyy-MM-dd")), _
'                                IIf(IsNull(RecReturn!���ʧЧ��), "", Format(RecReturn!���ʧЧ��, "yyyy-MM-dd")), _
'                                IIf(IsNull(RecReturn!��������), "0", RecReturn!��������), _
'                                IIf(IsNull(RecReturn!ʵ�ʽ��), "0", RecReturn!ʵ�ʽ��), _
'                                IIf(IsNull(RecReturn!ʵ�ʲ��), "0", RecReturn!ʵ�ʲ��), _
'                                IIf(IsNull(RecReturn!ָ�������), "0", RecReturn!ָ�������), _
'                                IIf(mintUnit = 0, 1, RecReturn!����ϵ��), IIf(IsNull(RecReturn!����), 0, RecReturn!����), RecReturn!ʱ��, RecReturn!���÷���, IIf(IsNull(RecReturn!��׼�ĺ�), "", RecReturn!��׼�ĺ�)) = False Then
'                            Cancel = True
'                            Exit Sub
'                        End If
'                        .Text = .TextMatrix(.Row, .Col)
'                    Else
'                        Cancel = True
'                    End If
                    Call ��ʾ�����
                End If
            
            Case mconIntCol����, mconIntCol��������
                If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
                    MsgBox "�����������룡", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "��������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) = 0 And mint�༭״̬ <> 3 Then
                        MsgBox "��������Ϊ��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If mint�༭״̬ = 6 Then
                        If Val(strKey) > Val(.TextMatrix(.Row, mconIntCol����)) Then
                            MsgBox "�����������ܴ���������", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    If .TextMatrix(.Row, 0) = "" Then Exit Sub
                    If Not CompareUsableQuantity(.Row, strKey) Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '�ɱ��۵Ĺ�ʽ��     ������=����*�ۼ�
                    '                  ������=������*��ʵ�ʲ��/ʵ�ʽ�
                    '                  if ʵ�ʽ��=0 then  ������=������*ָ�������
                    '                  ���ۣ��ɱ��ۣ�=��������-�����ۣ�/����
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    
                    strKey = Format(strKey, mFMT.FM_����)
                    .Text = strKey
                    
                    If .TextMatrix(.Row, mconIntCol�ۼ�) <> "" Then
                        .TextMatrix(.Row, mconIntCol�ۼ۽��) = Format(.TextMatrix(.Row, mconIntCol�ۼ�) * strKey, mFMT.FM_���)
                    End If
                    
                    If mint�༭״̬ <> 6 Then
'                        Dim dbl��� As Double, dbl���� As Double, dbl�ɱ���� As Double
'
'                        Call ��֤�����ۼ���(cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol����)), Val(.TextMatrix(.Row, mconIntCol����ϵ��)), Val(.TextMatrix(.Row, mconIntColʵ�ʲ��)), Val(.TextMatrix(.Row, mconIntColʵ�ʽ��)), Val(Split(.TextMatrix(.Row, mconIntColָ�������), "||")(0)) / 100, Val(strKey), Val(.TextMatrix(.Row, mconIntCol�ۼ۽��)), dbl���, dbl����, dbl�ɱ����)
'                        .TextMatrix(.Row, mconintCol���) = Format(dbl���, mFMT.FM_���)
                        .TextMatrix(.Row, mconIntCol�ɹ���) = Format(Get�ɱ���(Val(.TextMatrix(.Row, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, mconIntCol����))) * Val(.TextMatrix(.Row, mconIntCol����ϵ��)), mFMT.FM_�ɱ���)
'                        .TextMatrix(.Row, mconIntCol�ɹ����) = Format(dbl�ɱ����, mFMT.FM_���)
'                    Else
'                        .TextMatrix(.Row, mconIntCol�ɹ����) = Format(Val(.TextMatrix(.Row, mconIntCol�ɹ���)) * strKey, mFMT.FM_���)
'                        .TextMatrix(.Row, mconintCol���) = Format(Val(.TextMatrix(.Row, mconIntCol�ۼ۽��)) - Val(.TextMatrix(.Row, mconIntCol�ɹ����)), mFMT.FM_���)
                    End If
                    
                    .TextMatrix(.Row, mconIntCol�ɹ����) = Format(Val(.TextMatrix(.Row, mconIntCol�ɹ���)) * strKey, mFMT.FM_���)
                    .TextMatrix(.Row, mconintCol���) = Format(Val(.TextMatrix(.Row, mconIntCol�ۼ۽��)) - Val(.TextMatrix(.Row, mconIntCol�ɹ����)), mFMT.FM_���)
                    
                    If .Col = mconIntCol���� Then
                        .TextMatrix(.Row, mconIntCol��������) = strKey
                    End If
                    
                    .TextMatrix(.Row, mconintCol�������) = Format(Val(.TextMatrix(.Row, mconintCol������)) * Val(strKey), mFMT.FM_���)
                    
                    '˰��=�������*��ֵ˰��
                    .TextMatrix(.Row, mconintCol˰��) = Format(Val(.TextMatrix(.Row, mconintCol������)) * Val(strKey) * (Val(.TextMatrix(.Row, mconintCol��ֵ˰��)) / 100 / (1 + Val(.TextMatrix(.Row, mconintCol��ֵ˰��)) / 100)), mFMT.FM_���)
                End If
                ��ʾ�ϼƽ��
            Case mconintCol������
                If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "�����۱���Ϊ�����ͣ������䣡", vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If strKey <> "" Then
                    If Val(strKey) < 0.001 Then
                        MsgBox "�Բ��������۱������0.001,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "�����۱���С��" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    .Text = Format(strKey, mFMT.FM_���ۼ�)
                    .TextMatrix(.Row, .Col) = .Text
                    
                    '����������
                    .TextMatrix(.Row, mconintCol�������) = Format(Val(.TextMatrix(.Row, mconintCol������)) * Val(.TextMatrix(.Row, mconIntCol����)), mFMT.FM_���)
                    
                    '����˰��
                    .TextMatrix(.Row, mconintCol˰��) = Format(Val(.TextMatrix(.Row, mconintCol������)) * Val(.TextMatrix(.Row, mconIntCol����)) * (Val(.TextMatrix(.Row, mconintCol��ֵ˰��)) / 100 / (1 + Val(.TextMatrix(.Row, mconintCol��ֵ˰��)) / 100)), mFMT.FM_���)
                End If
        End Select
    End With
End Sub

'�Ӳ���������ȡֵ��������Ӧ����
Private Function SetColValue(ByVal intRow As Integer, ByVal lng����ID As Long, _
        ByVal str���� As String, ByVal str��� As String, ByVal str���� As String, _
        ByVal str��λ As String, ByVal num�ۼ� As Double, ByVal str���� As String, _
        ByVal strЧ�� As String, ByVal str���ʧЧ�� As String, ByVal num�������� As Double, ByVal numʵ�ʽ�� As Double, _
        ByVal numʵ�ʲ�� As Double, ByVal numָ������� As Double, _
        ByVal num����ϵ�� As Double, ByVal lng���� As Long, _
        ByVal int�Ƿ��� As Integer, ByVal int���÷��� As Integer, ByVal str��׼�ĺ� As String) As Boolean
    
        Dim intCount As Integer
        Dim intCol As Integer
        Dim dblPrice As Double
        Dim rsTemp As New Recordset
    
    On Error GoTo ErrHandle
    SetColValue = False
    If Format(str���ʧЧ��, "yyyy-mm-dd") < Format(sys.Currentdate, "yyyy-mm-dd") And Trim(str���ʧЧ��) <> "" Then
       If MsgBox("���ϡ�" & str���� & "(" & lng���� & ")���Ѿ��������ʧЧ��,�Ƿ�Ҫ���ã�", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) <> vbYes Then
            Exit Function
       End If
    End If
    
    With mshBill
        Dim lngRow As Long
        For lngRow = 1 To .Rows - 1
            If lngRow <> intRow And .TextMatrix(lngRow, 0) <> "" Then
                If .TextMatrix(lngRow, 0) = lng����ID And Val(.TextMatrix(lngRow, mconIntCol����)) = lng���� Then
                    If UBound(Split(mstr�ظ�����, "��")) < 3 Then mstr�ظ����� = mstr�ظ����� & str���� & "��"  '����¼�����ظ�������
                    'Call MsgBox("�������ϡ�" & str���� & "(" & lng���� & ")���Ѿ����ڣ���ϲ��������ӣ�", vbOKOnly + vbInformation + vbDefaultButton2, gstrSysName)
                    Exit Function
                End If
            End If
        Next
                
        If int�Ƿ��� = 1 Then
            gstrSQL = "" & _
                "   Select nvl(���ۼ�,0)*" & num����ϵ�� & " as  �����ۼ�,ʵ�ʽ��/ʵ������* " & num����ϵ�� & " as ƽ�����ۼ�" & _
                "   From ҩƷ��� " & _
                "   Where �ⷿid=[1]" & _
                "           and ҩƷid=[2]" & _
                "           and ����=1 and ʵ������>0 and " & _
                "           nvl(����,0)=[3]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex), lng����ID, lng����)
                        
            If rsTemp.EOF Then
                MsgBox "ʱ������û�п�棬���ܳ��⣬���飡", vbOKOnly, gstrSysName
                Exit Function
            End If
            
            If lng���� = 0 Then
                dblPrice = rsTemp!ƽ�����ۼ�
            Else
                dblPrice = rsTemp!�����ۼ�
            End If
        End If
        
        For intCol = 0 To .Cols - 1
            If intCol <> mconIntCol�к� Then .TextMatrix(intRow, intCol) = ""
        Next
        
        .TextMatrix(intRow, mconIntCol�к�) = intRow
        .TextMatrix(intRow, 0) = lng����ID
        .TextMatrix(intRow, mconIntCol����) = str����
        .TextMatrix(intRow, mconIntCol���) = str���
        .TextMatrix(intRow, mconIntCol����) = str����
        .TextMatrix(intRow, mconIntCol��׼�ĺ�) = str��׼�ĺ�
        .TextMatrix(intRow, mconIntCol��λ) = str��λ
        .TextMatrix(intRow, mconIntCol����) = str����
        .TextMatrix(intRow, mconIntColЧ��) = Format(strЧ��, "yyyy-mm-dd")
        .TextMatrix(intRow, mconIntCol���ʧЧ��) = Format(str���ʧЧ��, "yyyy-mm-dd")
    
        .TextMatrix(intRow, mconIntCol�ۼ�) = Format(num�ۼ� * num����ϵ��, mFMT.FM_���ۼ�)
        .TextMatrix(intRow, mconIntCol��������) = Format(num��������, mFMT.FM_����)
        .TextMatrix(intRow, mconIntColʵ�ʲ��) = numʵ�ʲ��
        .TextMatrix(intRow, mconIntColʵ�ʽ��) = numʵ�ʽ��
        .TextMatrix(intRow, mconIntColָ�������) = numָ������� & "||" & int�Ƿ��� & "||" & int���÷���
        .TextMatrix(intRow, mconIntCol����ϵ��) = num����ϵ��
        .TextMatrix(intRow, mconIntCol����) = lng����
        If int�Ƿ��� = 1 Then .TextMatrix(intRow, mconIntCol�ۼ�) = Format(dblPrice, mFMT.FM_���ۼ�)
        Call CheckLapse(strЧ��)
        
        '������Ĭ��Ϊ�ɹ���=�����/����
        gstrSQL = "Select A.ָ��������, A.��ֵ˰��, Nvl(B.�ɹ���,0) As �ɹ��� " & _
            " From �������� A, " & _
            " (Select ҩƷid, �ϴβɹ��� / Nvl(�ϴο���, 100) * 100 As �ɹ��� " & _
            " From ҩƷ��� " & _
            " Where ���� = 1 And �ⷿid + 0 = [1] And ҩƷid = [2] And Nvl(����, 0) = [3]) B " & _
            " Where A.����id = B.ҩƷid(+) And A.����id = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҩƷ������Ϣ", Val(cboStock.ItemData(cboStock.ListIndex)), lng����ID, lng����)
        
        If Not rsTemp.EOF Then
            .TextMatrix(intRow, mconintCol��ֵ˰��) = zlStr.FormatEx(rsTemp!��ֵ˰��, 2)
            
            If rsTemp!�ɹ��� > 0 Then
                .TextMatrix(intRow, mconintCol������) = Format(rsTemp!�ɹ��� * num����ϵ��, mFMT.FM_���ۼ�)
            Else
                .TextMatrix(intRow, mconintCol������) = Format(rsTemp!ָ�������� * num����ϵ��, mFMT.FM_���ۼ�)
            End If
        End If
    End With
    Call ��ʾ�����
    SetColValue = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

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
    
    If txtNO.Locked = False Then
        If Trim(txtNO.Text) = "" Then
            ShowMsgBox "���ݺŲ���Ϊ��"
            Exit Function
        End If
        
        If InStr(1, txtNO.Text, "'") <> 0 Then
            ShowMsgBox "���ݺ��в��ܺ��зǷ��ַ�"
            Exit Function
        End If
        
        If InStr(1, txtNO.Text, ";") <> 0 Then
            ShowMsgBox "���ݺ��в��ܺ��зǷ��ַ�"
            Exit Function
        End If
        
        If LenB(StrConv(txtNO.Text, vbFromUnicode)) > txtNO.MaxLength Then
            ShowMsgBox "���ݺų���,���������" & CInt(txtNO.MaxLength / 2) & "�����֣���ò�Ҫ���֣���" & txtNO.MaxLength & "���ַ�!"
            txtNO.SetFocus
            Exit Function
        End If
    End If
    
    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '�����з�����
            
            If LenB(StrConv(txtժҪ.Text, vbFromUnicode)) > txtժҪ.MaxLength Then
                ShowMsgBox "ժҪ����,���������" & CInt(txtժҪ.MaxLength / 2) & "�����ֻ�" & txtժҪ.MaxLength & "���ַ�!"
                txtժҪ.SetFocus
                Exit Function
            End If
            If InStr(1, txtժҪ.Text, ";") <> 0 Then
                ShowMsgBox "��ժҪ�в�������ֺ�!"
                txtժҪ.SetFocus
                Exit Function
            End If
            For intLop = 1 To .Rows - 1
                If Trim(.TextMatrix(intLop, mconIntCol����)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mconIntCol����))) = "" Then
                        ShowMsgBox "��" & intLop & "�����ĵ�����Ϊ���ˣ����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol����
                        Exit Function
                    End If
                    
                    If Trim(Trim(.TextMatrix(intLop, mconIntCol��������))) = "" And mint�༭״̬ = 6 Then
                        ShowMsgBox "��" & intLop & "�����ĵ�����Ϊ���ˣ����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol��������
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconIntCol����)) > 9999999999# Then
                        ShowMsgBox "��" & intLop & "�����ĵ���д�������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol����
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol��������)) > 9999999999# Then
                        ShowMsgBox "��" & intLop & "�����ĵ�ʵ���������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol��������
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol�ɹ����)) > 9999999999999# Then
                        ShowMsgBox "��" & intLop & "�����ĵĳɱ������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = IIf(.ColData(mconIntCol����) = 4, mconIntCol����, mconIntCol��������)
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconIntCol�ۼ۽��)) > 9999999999999# Then
                        ShowMsgBox "��" & intLop & "�����ĵ��ۼ۽����������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = IIf(.ColData(mconIntCol����) = 4, mconIntCol����, mconIntCol��������)
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


Private Function SaveCard(Optional ByVal blnǿ�Ʊ��� As Boolean = False) As Boolean
    Dim lng������ID As Long
    Dim chrNo As Variant
    Dim lng��� As Long
    Dim lng�ⷿID As Long
    Dim lngTypeID As Long
    Dim lng����ID As Long
    Dim str���� As String
    Dim lng���� As Long
    Dim str���� As String
    Dim strЧ�� As String
    Dim dbl���� As Double
    Dim dbl�ɱ��� As Double
    Dim dbl�ɱ���� As Double
    Dim dbl���ۼ� As Double
    Dim dbl���۽�� As Double
    Dim dbl��� As Double
    Dim strժҪ As String
    Dim str������ As String
    Dim str�������� As String
    Dim str����� As String
    Dim datAssessDate As String
    Dim str���Ч�� As String
    Dim arrSQL As Variant
    Dim intRow As Integer
    
    Dim dblOutPrice As Double   '�����
    Dim strOutUnit As String    '�����λ
    Dim dbl��ֵ˰�� As Double
    Dim n As Long
    
    SaveCard = False
    arrSQL = Array()
    
    
    '����������������ID����Ҫ���������Ķ�Ҫ����
    
    
    With mshBill
        chrNo = Trim(txtNO)
        lng�ⷿID = cboStock.ItemData(cboStock.ListIndex)
        
        If mint�༭״̬ = 1 Then   'mbln�������� Or
            If chrNo <> "" Then
                If CheckNOExists(74, chrNo) Then Exit Function
            End If
        
            If chrNo = "" Then chrNo = sys.GetNextNo(74, lng�ⷿID)
            If IsNull(chrNo) Then Exit Function
        End If
        txtNO.Tag = chrNo
        
        lng������ID = cboType.ItemData(cboType.ListIndex)
        strժҪ = Trim(txtժҪ.Text)
        str������ = Txt������
        str�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        str����� = Txt�����
        
        If cboType.Text = "��������" Then
            strOutUnit = Mid(cbo������λ.Text, 1, InStr(1, cbo������λ.Text, "-") - 1)
        Else
            strOutUnit = ""
        End If
        
        On Error GoTo ErrHandle
        If mint�༭״̬ = 2 Or blnǿ�Ʊ��� = True Then       '�޸�
            gstrSQL = "zl_������������_Delete('" & mstr���ݺ� & "')"
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "0" & ";" & vbCrLf & gstrSQL
        End If
            
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng����ID = .TextMatrix(intRow, 0)
                str���� = .TextMatrix(intRow, mconIntCol����)
                str���� = .TextMatrix(intRow, mconIntCol����)
                lng���� = .TextMatrix(intRow, mconIntCol����)
                strЧ�� = IIf(.TextMatrix(intRow, mconIntColЧ��) = "", "", .TextMatrix(intRow, mconIntColЧ��))
                dbl���� = Round(Val(.TextMatrix(intRow, mconIntCol����)) * Val(.TextMatrix(intRow, mconIntCol����ϵ��)), g_С��λ��.obj_���С��.����С��)
                dbl�ɱ��� = Round(Val(.TextMatrix(intRow, mconIntCol�ɹ���)) / Val(.TextMatrix(intRow, mconIntCol����ϵ��)), g_С��λ��.obj_���С��.�ɱ���С��)
                dbl�ɱ���� = Round(Val(.TextMatrix(intRow, mconIntCol�ɹ����)), g_С��λ��.obj_���С��.���С��)
                dbl���ۼ� = Round(Val(.TextMatrix(intRow, mconIntCol�ۼ�)) / Val(.TextMatrix(intRow, mconIntCol����ϵ��)), g_С��λ��.obj_���С��.���ۼ�С��)
                str���Ч�� = IIf(.TextMatrix(intRow, mconIntCol���ʧЧ��) = "", "", .TextMatrix(intRow, mconIntCol���ʧЧ��))
                
                dbl���۽�� = Round(Val(.TextMatrix(intRow, mconIntCol�ۼ۽��)), g_С��λ��.obj_���С��.���С��)
                dbl��� = Round(Val(.TextMatrix(intRow, mconintCol���)), g_С��λ��.obj_���С��.���С��)
                lng��� = intRow
                
                If cboType.Text = "��������" Then
                    dblOutPrice = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol������)) / Val(.TextMatrix(intRow, mconIntCol����ϵ��)), g_С��λ��.obj_���С��.���ۼ�С��)
                End If
                
                dbl��ֵ˰�� = Val(.TextMatrix(intRow, mconintCol��ֵ˰��))
                
                'zl_������������_INSERT( /*������ID_IN*/, /*NO_IN*/, /*���_IN*/,
                    '/*�ⷿID_IN*/, /*����ID_IN*/, /*����_IN*/, /*��д����_IN*/,
                    '/*�ɱ���_IN*/, /*�ɱ����_IN*/, /*���ۼ�_IN*/, /*���۽��_IN*/,
                    '/*���_IN*/, /*������_IN*/, /*��������_IN*/, /*����_IN*/,
                    '/*����_IN*/, /*Ч��_IN*/���Ч��/, /*ժҪ_IN*/ );
                
                gstrSQL = "zl_������������_INSERT(" & _
                    lng������ID & ",'" & _
                    chrNo & "'," & _
                    lng��� & "," & _
                    lng�ⷿID & "," & lng����ID & "," & lng���� & "," & dbl���� & "," & _
                    dbl�ɱ��� & "," & dbl�ɱ���� & "," & dbl���ۼ� & "," & dbl���۽�� & "," & _
                    dbl��� & ",'" & str������ & "',to_date('" & str�������� & "','yyyy-mm-dd HH24:MI:SS'),'" & str���� & "','" & _
                    str���� & "'," & _
                    IIf(strЧ�� = "", "Null", "to_date('" & Format(strЧ��, "yyyy-mm-dd") & "','yyyy-mm-dd')") & "," & _
                    IIf(str���Ч�� = "", "Null", "to_date('" & Format(str���Ч��, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ",'" & _
                    strժҪ & "'," & _
                    dblOutPrice & ",'" & _
                    strOutUnit & "'," & _
                    dbl��ֵ˰�� & ",1)"
                    
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = CStr(lng����ID) & ";" & vbCrLf & gstrSQL

            End If
            
            recSort.MoveNext
        Next
        
        If Not ExecuteSql(arrSQL, mstrCaption, False) Then Exit Function
        If Not ��鵥��(21, txtNO.Tag) Then
            gcnOracle.RollbackTrans
            Exit Function
        End If
        gcnOracle.CommitTrans
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    
    SaveCard = True
    Exit Function
ErrHandle:
    
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog

End Function


Private Sub ��ʾ�ϼƽ��()
    Dim curTotal As Double, Cur���ʽ�� As Double, Cur���ʲ�� As Double, Cur������� As Double
    Dim intLop As Integer
    
    curTotal = 0: Cur���ʽ�� = 0: Cur���ʲ�� = 0:
    
    With mshBill
        For intLop = 1 To .Rows - 1
            curTotal = curTotal + Val(.TextMatrix(intLop, mconIntCol�ɹ����))
            Cur���ʽ�� = Cur���ʽ�� + Val(.TextMatrix(intLop, mconIntCol�ۼ۽��))
            Cur������� = Cur������� + Val(.TextMatrix(intLop, mconintCol�������))
        Next
    End With
    
    Cur���ʲ�� = Cur���ʽ�� - curTotal
    lblPurchasePrice.Caption = "�ɱ����ϼƣ�" & Format(curTotal, mFMT.FM_���)
    lblSalePrice.Caption = "�ۼ۽��ϼƣ�" & Format(Cur���ʽ��, mFMT.FM_���)
    lblDifference.Caption = "��ۺϼƣ�" & Format(Cur���ʲ��, mFMT.FM_���)
    lblOther.Caption = "�����ϼƣ�" & Format(Cur�������, mFMT.FM_���)
End Sub

Private Sub ��ʾ�����()
    Dim rsTemp As New Recordset
    
    On Error GoTo ErrHandle
    With mshBill
        If .TextMatrix(.Row, mconIntCol����) = "" Then
            stbThis.Panels(2).Text = ""
            Exit Sub
        End If
        If .TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
        gstrSQL = "" & _
            "   Select ��������/" & .TextMatrix(.Row, mconIntCol����ϵ��) & " as  �������� " & _
            "   From ҩƷ��� " & _
            "   Where �ⷿid=[1]" & _
            "           and ҩƷid=[2]" & _
            "           and ����=1 and " & _
            "           nvl(����,0)=[3]"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol����)))
        
        If rsTemp.EOF Then
            .TextMatrix(.Row, mconIntCol��������) = 0
        Else
            .TextMatrix(.Row, mconIntCol��������) = IIf(IsNull(rsTemp.Fields(0)), 0, rsTemp.Fields(0))
        End If
        rsTemp.Close
        stbThis.Panels(2).Text = "�����ĵ�ǰ�����Ϊ[" & Format(.TextMatrix(.Row, mconIntCol��������), mFMT.FM_����) & "]" & .TextMatrix(.Row, mconIntCol��λ)
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboType_LostFocus()
    If cboType.Text = "" Then
        cboType.Tag = "0"
        Exit Sub
    End If
End Sub

Private Sub cboType_Validate(Cancel As Boolean)
    If cboType.Text = "" Then
        cboType.Tag = "0"
        Exit Sub
    End If
End Sub

Private Sub TxtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub txtժҪ_Change()
    mblnChange = True
End Sub

Private Sub txtժҪ_GotFocus()
    ImeLanguage True
    
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

Private Sub txtժҪ_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtժҪ, KeyAscii, m�ı�ʽ
    If KeyAscii = Asc(";") Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtժҪ_LostFocus()
    ImeLanguage False
End Sub

'������������бȽ�
Private Function CompareUsableQuantity(ByVal intRow As Integer, ByVal dbl��д���� As Double) As Boolean
    Dim dblUsableQuantity As Double      'ʵ��������Ӧ���������
    Dim numUsedCount As Double, dbltotal As Double
    Dim vardrug As Variant, intLop As Integer
    
    'mint�����: 0-�����;1-��飬�������ѣ�2-��飬�����ֹ
    
    CompareUsableQuantity = False

    With mshBill
        If .TextMatrix(intRow, 0) = "" Then Exit Function
        dblUsableQuantity = Format(.TextMatrix(intRow, mconIntCol��������), mFMT.FM_����)
        
        If mint����� = 0 Then
            '0-�����
        ElseIf mint����� = 1 Then
            '1-��飬��������
            If mint�༭״̬ = 1 Then
                If dbl��д���� > dblUsableQuantity Then
                    If MsgBox("�������������" & dbl��д���� & "�������˸����ĵĿ��ÿ��������" & dblUsableQuantity & "�����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
            ElseIf mint�༭״̬ = 2 Or mint�༭״̬ = 3 Then
                numUsedCount = 0
                For Each vardrug In mcolUsedCount
                    If vardrug(0) = .TextMatrix(intRow, 0) & .TextMatrix(intRow, mconIntCol����) Then
                        numUsedCount = vardrug(1)
                        Exit For
                    End If
                Next
                
                If gSystem_Para.para_������¿��ÿ�� = False Then
                    '���û��Ԥ��������������������ԭʼ����
                    numUsedCount = 0
                End If
                
                If dbl��д���� > dblUsableQuantity + numUsedCount Then
                    If MsgBox("�������������" & dbl��д���� & "�������˸����ĵĿ��ÿ��������" & dblUsableQuantity + numUsedCount & "�����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
            End If
            
        ElseIf mint����� = 2 Then
            '2-��飬�����ֹ
            If mint�༭״̬ = 1 Then
                If dbl��д���� > dblUsableQuantity Then
                    MsgBox "�������������" & dbl��д���� & "�������˸����ĵĿ��ÿ��������" & dblUsableQuantity & "���������䣡", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            ElseIf mint�༭״̬ = 2 Or mint�༭״̬ = 3 Then
                numUsedCount = 0
                For Each vardrug In mcolUsedCount
                    If vardrug(0) = .TextMatrix(intRow, 0) & .TextMatrix(intRow, mconIntCol����) Then
                        numUsedCount = vardrug(1)
                        Exit For
                    End If
                Next
                
                If gSystem_Para.para_������¿��ÿ�� = False Then
                    '���û��Ԥ��������������������ԭʼ����
                    numUsedCount = 0
                End If
                
                If dbl��д���� > dblUsableQuantity + numUsedCount Then
                    MsgBox "�������������" & dbl��д���� & "�������˸����ĵĿ��ÿ��������" & dblUsableQuantity + numUsedCount & "���������䣡", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
        End If
            
    End With
    
    CompareUsableQuantity = True
    
End Function

'��ӡ����
Private Sub printbill()
    Dim strNo As String
    strNo = txtNO.Tag
    FrmBillPrint.ShowMe Me, glngSys, "zl1_bill_1718", mint��¼״̬, mintUnit, 1718, "������������", strNo
End Sub

Private Sub SetSortRecord()
    Dim n As Integer
    
    If mshBill.Rows < 2 Then Exit Sub
    If mshBill.TextMatrix(1, 0) = "" Then Exit Sub
    
    Set recSort = New ADODB.Recordset
    With recSort
        If .State = 1 Then .Close
        .Fields.Append "�к�", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        For n = 1 To mshBill.Rows - 1
            If mshBill.TextMatrix(n, 0) <> "" Then
                .AddNew
                !�к� = n
                !��� = IIf(Val(mshBill.TextMatrix(n, mconIntCol���)) = 0, n, Val(mshBill.TextMatrix(n, mconIntCol���)))
                !ҩƷid = Val(mshBill.TextMatrix(n, 0))
                !���� = Val(mshBill.TextMatrix(n, mconIntCol����))
                
                .Update
            End If
        Next
        
    End With
End Sub

