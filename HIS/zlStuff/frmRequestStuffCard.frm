VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmRequestStuffCard 
   Caption         =   "�������쵥"
   ClientHeight    =   6855
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11760
   Icon            =   "frmRequestStuffCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   11760
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdRequest 
      Caption         =   "���깺������(&R)"
      Height          =   350
      Left            =   3840
      TabIndex        =   29
      Top             =   5880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "ȫ��(&A)"
      Height          =   350
      Left            =   6345
      TabIndex        =   28
      Top             =   5535
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "ȫ��(&L)"
      Height          =   350
      Left            =   7665
      TabIndex        =   27
      Top             =   5535
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   9
      Top             =   5137
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   8
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   7
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   6
      Top             =   5040
      Width           =   1100
   End
   Begin VB.PictureBox Pic���� 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   10
      Top             =   0
      Width           =   11715
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   2
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
         TabIndex        =   4
         Top             =   4080
         Width           =   10410
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   557
         Width           =   1515
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "��ۺϼ�:"
         Height          =   180
         Left            =   4920
         TabIndex        =   25
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ۼ۽��ϼ�:"
         Height          =   180
         Left            =   2040
         TabIndex        =   24
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ɱ����ϼ�:"
         Height          =   180
         Left            =   240
         TabIndex        =   23
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label Txt����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7950
         TabIndex        =   21
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Txt������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10050
         TabIndex        =   20
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt�������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   19
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt������ 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   18
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   9960
         TabIndex        =   17
         Top             =   550
         Width           =   1425
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
         TabIndex        =   16
         Top             =   587
         Width           =   480
      End
      Begin VB.Label lblժҪ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ժҪ(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   3
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "�����������쵥"
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
         TabIndex        =   15
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���Ͽⷿ(&S)"
         Height          =   180
         Left            =   240
         TabIndex        =   0
         Top             =   615
         Width           =   990
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   300
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
            Picture         =   "frmRequestStuffCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":1000
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
            Picture         =   "frmRequestStuffCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestStuffCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   26
      Top             =   6495
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmRequestStuffCard.frx":22EA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14393
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmRequestStuffCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmRequestStuffCard.frx":3080
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
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   5
      Top             =   5040
      Width           =   1100
   End
   Begin VB.Label lblCode 
      Caption         =   "����"
      Height          =   255
      Left            =   3255
      TabIndex        =   22
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmRequestStuffCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint�༭״̬ As Integer             '1.������2���޸ģ�3�����գ�4���鿴��5��ͨ����������6�����ܣ����պ��¼���յǼ��ˣ�����ȡ������Ľ��գ���7������
Private mstr���ݺ� As String                '����ĵ��ݺ�;
Private mint��¼״̬ As Integer             '1:������¼;2-������¼;3-�Ѿ�������ԭ��¼
Private mblnSuccess As Boolean              'ֻҪ��һ�ųɹ�����ΪTrue������ΪFalse
Private mblnFirst As Boolean
Private mblnSave As Boolean                 '�Ƿ���̺����   TURE���ɹ���
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '�Ƿ�����޸�
Private mblnChange As Boolean               '�Ƿ���й��༭

Private mint��ȷ���� As Integer             '��ʾ����д���쵥ʱ���Ƿ���ȷ���ĵ�����
Private mint����� As Integer             '��ʾ���ĳ���ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ
Private mcolUsedCount As Collection         '��ʹ�õ���������
Private mstrPrivs As String                     'Ȩ��
Private mlngStockID As Long                 '��ǰ�û���ѡ�ķ��ϲ���ID
Private rsDepend As New ADODB.Recordset

Private mintParallelRecord As Integer       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼
Private Const mlngModule = 1722
Private mint����ʾ�п������ As Boolean
Private mblnCostView As Boolean                 '�鿴�ɱ��� true-����鿴 false-������鿴
Private mblnUpdate As Boolean               '��ʾ�Ƿ��Ѹ������¼۸���µ�������
Private Const mstrCaption As String = "�������쵥"
Private mbln����˲� As Boolean             '�ƿ��Ƿ���Ҫ�˲�,true-��Ҫ��false-����Ҫ
Private mint����ʽ As Integer             '����ʱ��0������������1�������������뵥��
Private mstr�ظ����� As String '��¼�ظ�������

Private mstrRequestNO As String     '���깺���ƿ�NO ���մ��������깺����ʽ���죬�������깺������
'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
Private recSort As ADODB.Recordset          '��ҩƷID�����ר�ü�¼��

Private mstrTime_Start As String                        '���뵥�ݱ༭����ʱ�����༭���ݵ�����޸�ʱ��
Private mstrTime_End As String                        '�˿̸ñ༭���ݵ�����޸�ʱ��

'=========================================================================================
Private Enum mBillCol
    C_�к� = 1
    C_���� = 2
    c_��� = 3
    C_��� = 4
    C_�ⷿ���� = 5
    C_���Ч�� = 6
    C_�������� = 7
    C_ָ������� = 8
    C_ʵ�ʽ�� = 9
    C_ʵ�ʲ�� = 10
    C_����ϵ�� = 11
    c_���� = 12
    C_���� = 13
    C_��׼�ĺ� = 14
    c_��λ = 15
    c_���� = 16
    C_Ч�� = 17
    C_���ʧЧ�� = 18
    C_��ǰ��� = 19
    C_�Է���� = 20
    C_��д���� = 21
    C_ʵ������ = 22
    c_ԭʼ���� = 23
    C_�ɹ��� = 24
    C_�ɹ���� = 25
    C_�ۼ� = 26
    C_�ۼ۽�� = 27
    C_��� = 28
End Enum

Private Const mBillCols  As Integer = 29              '������
'=========================================================================================
Private mintUnit  As Integer                '��ʾ��λ:0-ɢװ��λ,1-��װ��λ


'�������������
Private Function GetDepend() As Boolean
    Dim strMsg As String
    
    On Error GoTo ErrHandle
    GetDepend = False
    
    '���ҩƷ�������Ƿ�����
    strMsg = "û�����������ƿ����⼰�����������������������ã�"
    
    gstrSQL = "" & _
        "   SELECT B.Id,B.ϵ�� " & _
        "   FROM ҩƷ�������� A, ҩƷ������ B " & _
        "   Where A.���id = B.ID AND A.���� = 34"
    
    zlDatabase.OpenRecordset rsDepend, gstrSQL, "�����ƿ����"
        
    With rsDepend
        If .RecordCount = 0 Then GoTo ErrHand
        .Filter = "ϵ��=1"
        If .RecordCount = 0 Then
            .Filter = 0
            strMsg = "û�����������ƿ������������������������ã�"
            GoTo ErrHand
        End If
        .Filter = "ϵ��=-1"
        If .RecordCount = 0 Then
            .Filter = 0
            strMsg = "û�����������ƿ�ĳ����������������������ã�"
            GoTo ErrHand
        End If
        .Filter = 0
        .Close
    End With
    
    Set rsDepend = ReturnSQL(mlngStockID, "�����ƿ����", False, , 1722)
    rsDepend.Filter = "ID<>" & mlngStockID
    With rsDepend
        strMsg = "û���κοⷿ�������죬����[���Ĳ�������]���������������ã�"
        If .RecordCount = 0 Then GoTo ErrHand
    End With
    GetDepend = True
    Exit Function
ErrHand:
    MsgBox strMsg, vbInformation, gstrSysName
    rsDepend.Close
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub ShowCard(frmMain As Form, ByVal str���ݺ� As String, ByVal int�༭״̬ As Integer, Optional int��¼״̬ As Integer = 1, Optional strPrivs As String, Optional blnSuccess As Boolean = False, Optional lngStockID As Long = 0, Optional int����ʽ As Integer = 0)
    Dim strSQL As String
    Dim rsPara As New ADODB.Recordset
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
    mlngStockID = IIf(lngStockID = 0, glngDeptId, lngStockID)
    mint����ʽ = int����ʽ

    Set mfrmMain = frmMain
    If Not GetDepend Then Exit Sub
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    
    strReg = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngModule, "0"))
    
    mintUnit = Val(strReg)
    mint����� = Get������(mlngStockID)
    
    mint��ȷ���� = IIf(IS��������, 1, 0)

    If mint��ȷ���� = 0 Then mint����� = 0
    
    If mint�༭״̬ = 1 Or mint�༭״̬ = 5 Then
        mblnEdit = True
        mblnFirst = True
    ElseIf mint�༭״̬ = 2 Then
        mblnEdit = True
        mblnFirst = True
    ElseIf mint�༭״̬ = 3 Then
        CmdSave.Caption = "�˲�(&C)"
        Lbl������.Caption = "�˲���"
        Lbl��������.Caption = "�˲�����"
    ElseIf mint�༭״̬ = 4 Then
        mblnFirst = True
        mblnEdit = False
        CmdSave.Caption = "��ӡ(&P)"
        If InStr(mstrPrivs, "���ݴ�ӡ") = 0 Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
    ElseIf mint�༭״̬ = 7 Then
        mblnEdit = False
        mblnFirst = True
        CmdSave.Caption = "����(&O)"
        cmdAllSel.Visible = True
        cmdAllCls.Visible = True
        
        If mint����ʽ = 1 Then
            CmdSave.Caption = "�������(&O)"
            CmdSave.Width = CmdSave.Width + 200
        Else
            CmdSave.Caption = "����(&O)"
            CmdSave.Width = cmdCancel.Width
        End If

    End If
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    Me.Show vbModal, frmMain
    blnSuccess = mblnSuccess
    str���ݺ� = mstr���ݺ�
    
End Sub

Private Sub cboStock_Change()
    mblnChange = True
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
                If MsgBox("����ı�ⷿ���п���Ҫ�ı���Ӧ���ĵĵ�λ����Ҫ������е������ݣ����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
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
        
        mint����� = Get������(cboStock.ItemData(cboStock.ListIndex))
        If mint��ȷ���� = 0 Then mint����� = 0
        
    End With
End Sub

Private Sub cmdAllCls_Click()
    Dim intRow As Integer
    
    With mshBill
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                .TextMatrix(intRow, mBillCol.C_ʵ������) = Format(0, mFMT.FM_����)
                .TextMatrix(intRow, mBillCol.C_�ɹ����) = Format(0, mFMT.FM_���)
                .TextMatrix(intRow, mBillCol.C_�ۼ۽��) = Format(0, mFMT.FM_���)
                .TextMatrix(intRow, mBillCol.C_���) = Format(0, mFMT.FM_���)
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
                .TextMatrix(intRow, mBillCol.C_ʵ������) = .TextMatrix(intRow, mBillCol.C_��д����)
                .TextMatrix(intRow, mBillCol.C_�ɹ����) = Format(.TextMatrix(intRow, mBillCol.C_��д����) * .TextMatrix(intRow, mBillCol.C_�ɹ���), mFMT.FM_���)
                .TextMatrix(intRow, mBillCol.C_�ۼ۽��) = Format(.TextMatrix(intRow, mBillCol.C_��д����) * .TextMatrix(intRow, mBillCol.C_�ۼ�), mFMT.FM_���)
                .TextMatrix(intRow, mBillCol.C_���) = Format(.TextMatrix(intRow, mBillCol.C_�ۼ۽��) - .TextMatrix(intRow, mBillCol.C_�ɹ����), mFMT.FM_���)
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
        
        cmdRequest.Left = txtCode.Left + txtCode.Width + (cmdFind.Left - cmdHelp.Left - cmdHelp.Width)
    Else
        FindRownew mshBill, mBillCol.C_����, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
        
        cmdRequest.Left = cmdFind.Left + cmdFind.Width + (cmdFind.Left - cmdHelp.Left - cmdHelp.Width)
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub



Private Sub cmdRequest_Click()
    Dim rsTemp As ADODB.Recordset
    Dim lngRow As Long
    Dim blnDo As Boolean
    Dim str���Ч�� As String
    Dim blnҩ�� As Boolean
    Dim dblPrice As Double
    Dim strЧ�� As String
    Dim dbl���� As Double
    Dim lng����ID As Long
    Dim bln���� As Boolean
    Dim dbl�깺����  As Double
    Dim dbl�ѵ����� As Double
    
    If mlngStockID = 0 Then  '������ⷿ
        MsgBox "����ⷿ����Ϊ�գ�", vbInformation, gstrSysName
        Exit Sub
    End If


    mstrRequestNO = frmDrawCondition.ShowMe(Me, mintUnit, cboStock.Text, Val(cboStock.ItemData(cboStock.ListIndex)), mfrmMain.cboStock.Text, mlngStockID)
    If mstrRequestNO <> "" Then
        blnDo = False
        mstrRequestNO = Mid(mstrRequestNO, 1, LenB(StrConv(mstrRequestNO, vbFromUnicode)) - 1)

        blnҩ�� = True
        gstrSQL = "Select Distinct 0 " & _
                                    "From ��������˵�� " & _
                                    "Where ((�������� Like '���ϲ���') Or (�������� Like '�Ƽ���')) And ����id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex))
        If rsTemp.RecordCount = 0 Then
            blnҩ�� = False
        End If

        gstrSQL = "Select a.Id as ����id, d.���� as �ƻ�����,a.����,a.���� ,a.���,c.�ּ� as �ۼ�,a.���㵥λ as ɢװ��λ,a.�Ƿ��� as ʱ��,b.��װ��λ,b.����ϵ��,b.ָ�������,b.���Ч��,b.һ���Բ���" & vbNewLine & _
                    ",e.�ϴβ��� as ����,e.�ϴ����� as ����,nvl(e.����,0) as ����,e.Ч��,e.���Ч��,e.��������,nvl(e.ʵ������,0) as ʵ������,e.ʵ�ʽ��,e.ʵ�ʲ��,e.���ۼ�,e.ƽ���ɱ���,e.��׼�ĺ�,b.�ⷿ����,b.���÷���, nvl(b.���ٲ���,0) as ���ٲ���" & vbNewLine & _
                    "From �շ���ĿĿ¼ A, �������� B, �շѼ�Ŀ C," & vbNewLine & _
                    "     (Select  b.����id, Sum(b.�ƻ�����) As ����" & vbNewLine & _
                    "       From ���ϲɹ��ƻ� A, ���ϼƻ����� B" & vbNewLine & _
                    "       Where a.Id = b.�ƻ�id  and a.����=1 And a.No In (Select * From Table(Cast(f_Str2list([1]) As Zltools.t_Strlist)))" & vbNewLine & _
                    "       Group By b.����id) D,ҩƷ��� e" & vbNewLine & _
                    "Where a.Id = b.����id And b.����id = c.�շ�ϸĿid And a.Id = d.����id and b.����id=e.ҩƷid(+)  and e.�ⷿid=[2] and e.ʵ������>0 and e.����=1 And Sysdate Between c.ִ������ And c.��ֹ����"

        If gSystem_Para.P156_�����㷨 = 0 Then '���λ���Ч�������ȳ���
            gstrSQL = gstrSQL & " Order by a.id,Nvl(e.����, 0)"
        Else
            gstrSQL = gstrSQL & " Order by a.id,e.Ч��,Nvl(e.����, 0)"
        End If

        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "cmdRequestTransfer_Click", mstrRequestNO, cboStock.ItemData(cboStock.ListIndex))

        Do While Not rsTemp.EOF
            With mshBill
                For lngRow = 1 To .Rows - 1
                    If Val(.TextMatrix(lngRow, 0)) <> 0 Then
                        If Val(.TextMatrix(lngRow, 0)) = rsTemp!����ID And Val(.TextMatrix(lngRow, mBillCol.c_����)) = rsTemp!���� Then
                            blnDo = True
                            MsgBox "�ظ�����" & "[" & rsTemp!���� & "-" & rsTemp!���� & "]" & "������ӣ�", vbInformation, gstrSysName
                            Exit For
                        End If
                    End If
                Next

                If Val(.TextMatrix(.Rows - 1, 0)) = 0 Then
                    lngRow = .Rows - 1
                Else
                    .Rows = .Rows + 1
                    lngRow = .Rows - 1
                End If

                str���Ч�� = IIf(IsNull(rsTemp!���Ч��), "", Format(rsTemp!���Ч��, "yyyy-MM-dd"))
                If Format(str���Ч��, "yyyy-mm-dd") < Format(zlDatabase.Currentdate, "yyyy-mm-dd") And Trim(str���Ч��) <> "" Then
                   If MsgBox("[" & rsTemp!���� & "-" & rsTemp!���� & "]" & "�����Ѿ��������Ч��,�Ƿ�Ҫ���ã�", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) <> vbYes Then
                        blnDo = True
                   End If
                End If

'                strЧ�� = IIf(IsNull(rsTemp!Ч��), "", Format(rsTemp!Ч��, "yyyy-MM-dd"))
'                If IsDate(strЧ��) Then
'                    If Format(strЧ��, "yyyy-MM-dd") < Format(zldatabase.Currentdate, "yyyy-MM-dd") Then
'                        MsgBox "[" & rsTemp!���� & "-" & rsTemp!���� & "]" & "���������Ѿ�ʧЧ�ˣ�", vbInformation, gstrSysName
'                    End If
'                End If

                'ȡ�ۼ�
                If rsTemp!ʱ�� = 1 Then
                    If rsTemp!���÷��� = 0 Then
                        If rsTemp!�ⷿ���� = 1 And blnҩ�� = False Then
                            bln���� = True
                        Else
                            bln���� = False
                        End If
                    Else
                        bln���� = True
                    End If

                    If bln���� = True Then
                        If IsNull(rsTemp!���ۼ�) Then
                            If rsTemp!ʵ������ = 0 Then
                                dblPrice = 0
                            Else
                                dblPrice = rsTemp!ʵ�ʽ�� / rsTemp!ʵ������
                            End If
                        Else
                            dblPrice = rsTemp!���ۼ�
                        End If
                    Else
                        If rsTemp!ʵ������ = 0 Then
                            dblPrice = 0
                        Else
                            dblPrice = rsTemp!ʵ�ʽ�� / rsTemp!ʵ������
                        End If
                    End If
                Else
                    dblPrice = IIf(IsNull(rsTemp!�ۼ�), 0, rsTemp!�ۼ�)
                End If

                If lng����ID = rsTemp!����ID Then
                    If rsTemp!�������� + dbl�ѵ����� > rsTemp!�ƻ����� Then
                        If rsTemp!�ƻ����� - dbl�ѵ����� <> 0 Then
                            dbl���� = rsTemp!�ƻ����� - dbl�ѵ�����
                            dbl�ѵ����� = dbl�ѵ����� + dbl����
                        Else
                            blnDo = True
                        End If
                    Else
                        dbl���� = rsTemp!��������
                        dbl�ѵ����� = dbl�ѵ����� + dbl����
                    End If
                Else
                    If rsTemp!�������� > rsTemp!�ƻ����� Then
                        dbl���� = rsTemp!�ƻ�����
                    Else
                        dbl���� = rsTemp!��������
                    End If
                    dbl�ѵ����� = dbl����
                End If
                lng����ID = rsTemp!����ID

                If dbl���� = 0 Then
                    blnDo = True
                End If

                'ֻ�в��ظ��Ĳ���ӵ������ȥ
                If blnDo = False Then
                
                    SetRequestColValue lngRow, rsTemp!����ID, "[" & rsTemp!���� & "]" & rsTemp!����, _
                    IIf(IsNull(rsTemp!���), "", rsTemp!���), IIf(IsNull(rsTemp!����), "", rsTemp!����), _
                    IIf(mintUnit = 0, rsTemp!ɢװ��λ, rsTemp!��װ��λ), _
                    IIf(IsNull(rsTemp!�ۼ�), 0, rsTemp!�ۼ�), IIf(IsNull(rsTemp!����), "", rsTemp!����), _
                    IIf(IsNull(rsTemp!Ч��), "", rsTemp!Ч��), _
                    IIf(IsNull(rsTemp!���Ч��), "0", rsTemp!���Ч��), _
                    IIf(rsTemp!һ���Բ��� = 1, True, False), _
                    IIf(IsNull(rsTemp!���Ч��), "", rsTemp!���Ч��), _
                    rsTemp!�ⷿ����, _
                    IIf(IsNull(rsTemp!��������), "0", rsTemp!��������), _
                    IIf(IsNull(rsTemp!ʵ�ʽ��), "0", rsTemp!ʵ�ʽ��), _
                    IIf(IsNull(rsTemp!ʵ�ʲ��), "0", rsTemp!ʵ�ʲ��), _
                    IIf(IsNull(rsTemp!ָ�������), "0", rsTemp!ָ�������), _
                    IIf(mintUnit = 0, 1, rsTemp!����ϵ��), IIf(IsNull(rsTemp!����), 0, rsTemp!����), rsTemp!ʱ��, rsTemp!���÷���, IIf(IsNull(rsTemp!��׼�ĺ�), "", rsTemp!��׼�ĺ�)
                    
                    With mshBill
                        .Row = lngRow
                        .TextMatrix(lngRow, mBillCol.C_�к�) = lngRow
                    
                        .TextMatrix(lngRow, mBillCol.C_��д����) = Format(dbl���� / IIf(mintUnit = 0, 1, rsTemp!����ϵ��), mFMT.FM_����)

                        If .TextMatrix(lngRow, mBillCol.C_�ۼ�) <> "" Then
                            .TextMatrix(lngRow, mBillCol.C_�ۼ۽��) = Format(.TextMatrix(lngRow, mBillCol.C_�ۼ�) * .TextMatrix(lngRow, mBillCol.C_��д����), mFMT.FM_���)
                        End If
                        
                        Dim dbl��� As Double, dbl���� As Double, dbl�ɱ���� As Double
                        
                        Call ��֤�����ۼ���(cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(lngRow, 0)), Val(.TextMatrix(lngRow, mBillCol.c_����)), _
                            Val(.TextMatrix(lngRow, mBillCol.C_����ϵ��)), Val(.TextMatrix(lngRow, mBillCol.C_ʵ�ʲ��)), Val(.TextMatrix(lngRow, mBillCol.C_ʵ�ʽ��)), _
                            Val(Split(.TextMatrix(lngRow, mBillCol.C_ָ�������), "||")(0)) / 100, Val(.TextMatrix(lngRow, mBillCol.C_��д����)), Val(.TextMatrix(lngRow, mBillCol.C_�ۼ۽��)), dbl���, dbl����, dbl�ɱ����)
                        
                        .TextMatrix(lngRow, mBillCol.C_���) = Format(dbl���, mFMT.FM_���)
                        .TextMatrix(lngRow, mBillCol.C_�ɹ���) = Format(dbl����, mFMT.FM_�ɱ���)
                        .TextMatrix(lngRow, mBillCol.C_�ɹ����) = Format(dbl�ɱ����, mFMT.FM_���)

                        .TextMatrix(lngRow, mBillCol.C_ʵ������) = Format(dbl���� / IIf(mintUnit = 0, 1, rsTemp!����ϵ��), mFMT.FM_����)
                    End With

                End If

                blnDo = False
                rsTemp.MoveNext
            End With
        Loop
    End If
End Sub



Private Sub Form_Activate()
    If mblnFirst = False Then
        If mshBill.Rows > 50 Then
            Call AviShow(Me) '��ʾ�û����ڲ�ѯ����
        End If
        Call get�������    'Ϊ��ǰ��������ͶԷ���������и�ֵ
        If mshBill.Rows > 50 Then
            Call AviShow(Me, False)
        End If
        Exit Sub
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
    
    mblnFirst = False
    If mint�༭״̬ = 5 Then
        
        If Not frmRequestNavigation.ShowNavigation(Me, mlngStockID) = True Then
            Unload Me
            Exit Sub
        End If
        mint����� = Get������(cboStock.ItemData(cboStock.ListIndex))
        If mint��ȷ���� = 0 Then mint����� = 0
        mshBill.SetFocus
    End If
'    mblnChange = False
    Select Case mintParallelRecord
        Case 1
            '����
        Case 2
            '�����ѱ�ɾ��
            If mint�༭״̬ = 7 Then
                MsgBox "�õ�����û�п��Գ����Ĳ��ϣ����飡", vbOKOnly, gstrSysName
            Else
                '�����ѱ�ɾ��
                MsgBox "�õ����ѱ�ɾ�������飡", vbOKOnly, gstrSysName
            End If
            Unload Me
            Exit Sub
        Case 3
            '�޸ĵĵ����ѱ����
            MsgBox "�õ����ѱ���������ˣ����飡", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            cmdFind_Click
        End If
    ElseIf KeyCode = vbKeyF3 Then
        FindRownew mshBill, mBillCol.C_����, txtCode.Text, False
    ElseIf KeyCode = vbKeyF7 Then
        If stbThis.Panels("PY").Bevel = sbrRaised Then
            Logogram stbThis, 0
        Else
            Logogram stbThis, 1
        End If
    End If
End Sub

Private Function DeleteNo() As Boolean
    'ɾ������
    If txtNO.Caption <> "" Then
        gstrSQL = "zl_��������_delete('" & txtNO.Caption & "')"
        
        zlDatabase.ExecuteProcedure gstrSQL, "ɾ������"
        DeleteNo = True
        Exit Function
    End If
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub CmdSave_Click()
    Dim blnSuccess As Boolean
    Dim strReg As String
    
    '�����������ݼ�
    Call SetSortRecord
    
    If mint�༭״̬ = 3 Then
        '�˲�
        Call SaveCard
    End If
    
    If mint�༭״̬ = 4 Then    '�鿴
        '��ӡ
        printbill
        '�˳�
        Unload Me
        Exit Sub
    End If

    If mint�༭״̬ = 6 Then       '���
        If Not ���ϵ������(Txt������.Caption) Then Exit Sub
        
        If Not ��鵥��(19, txtNO.Tag, False) And Not mblnUpdate Then
            MsgBox "�м�¼δʹ�����¼۸񣬳����Զ���ɸ��£��ۼۡ��ɱ��ۡ��ۼ۽��ɱ�����ۣ������º����飡", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
        
        If SaveCheck() = True Then
            If IIf(Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModule, "0")) = 1, 1, 0) = 1 Then
                '��ӡ
                If InStr(mstrPrivs, "���ݴ�ӡ") <> 0 Then
                    printbill
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If
    
    
    If mint�༭״̬ = 7 Then '����
        If SaveStrike Then Unload Me
        Exit Sub
    End If
    
    
'    If mint�༭״̬ = 6 Or mint�༭״̬ = 7 Then '���ܣ����½�����
'        gstrSQL = "ZL_�����ƿ�_RECEIVE('" & txtNo.Caption & "'," & IIf(mint�༭״̬ = 6, "'" & gstrUserName & "'", "NULL") & ")"
'        Call ExecuteProcedure("���ܻ���տⷿ�����ĵ���")
'        mblnSuccess = True
'        Unload Me
'        Exit Sub
'    End If
    
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
    
'    mstr���ݺ� = NextNo(72)
    txtNO = ""
    mblnSave = False
    mblnEdit = True
    mshBill.ClearBill
    Call RefreshRowNO(mshBill, mBillCol.C_�к�, 1)
    
    txtժҪ.Text = ""
    If cboStock.Enabled Then cboStock.SetFocus
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
            " Where a.���� = 19 And a.No = [1] And a.ҩƷid = b.�շ�ϸĿid And c.Id = b.�շ�ϸĿid And Round(a.���ۼ�," & g_С��λ��.obj_ɢװС��.���ۼ�С�� & ") <> Round(b.�ּ�, " & g_С��λ��.obj_ɢװС��.���ۼ�С�� & ") And" & _
              "    NVL(c.�Ƿ���, 0) = 0" & _
            " Union All" & _
            " Select '�ۼ�' As ����, a.���, a.ҩƷid As ����id, Nvl(a.����, 0) As ����, decode(nvl(b.����,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�) As �ּ�" & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B, �շ���ĿĿ¼ C" & _
            " Where a.���� = 19 And a.No = [1] And c.Id = a.ҩƷid And Round(a.���ۼ�," & g_С��λ��.obj_ɢװС��.���ۼ�С�� & ") <> Round(decode(nvl(b.����,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�), " & g_С��λ��.obj_ɢװС��.���ۼ�С�� & ") And Nvl(c.�Ƿ���, 0) = 1 And" & _
                  " b.���� = 1 And b.�ⷿid = a.�ⷿid And b.ҩƷid = a.ҩƷid And NVL(b.����, 0) = NVL(a.����, 0) And NVL(b.ʵ������, 0) <> 0 And a.���ϵ�� = -1" & _
            " Union All" & _
            " Select '�ɱ���' As ����, a.���, a.ҩƷid As ����id, Nvl(a.����, 0) As ����, b.ƽ���ɱ��� As �ּ�" & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B" & _
            " Where a.���� = 19 And a.No = [1] And a.ҩƷid = b.ҩƷid And Nvl(a.����, 0) = Nvl(b.����, 0) and round(a.�ɱ���," & g_С��λ��.obj_ɢװС��.�ɱ���С�� & ")<>round(b.ƽ���ɱ���," & g_С��λ��.obj_ɢװС��.�ɱ���С�� & ") And a.�ⷿid = b.�ⷿid and a.���ϵ��=-1 and b.����=1" & _
            " Order By ����, ����id, ���"

    Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[ȡ��ǰ�۸�]", CStr(Me.txtNO.Caption))
    
    If rsprice.EOF Then Exit Sub
    
    lngRows = mshBill.Rows - 1
    For lngRow = 1 To lngRows
        blnAdj = False
        lng����ID = Val(mshBill.TextMatrix(lngRow, 0))
        dbl���� = Val(mshBill.TextMatrix(lngRow, mBillCol.C_ʵ������))
        dbl�ɱ��� = Val(mshBill.TextMatrix(lngRow, mBillCol.C_�ɹ���))
        dbl���ۼ� = Val(mshBill.TextMatrix(lngRow, mBillCol.C_�ۼ�))
        dbl�ɱ���� = dbl�ɱ��� * dbl����
        dbl���۽�� = dbl���ۼ� * dbl����
        dbl��� = dbl���۽�� - dbl�ɱ����
'
        If lng����ID <> 0 Then
            rsprice.Filter = "����='�ۼ�' And ����id=" & lng����ID & " And ����=" & Val(mshBill.TextMatrix(lngRow, mBillCol.c_����))
            If rsprice.RecordCount > 0 Then
                blnAdj = True
                dbl���ۼ� = Val(Format(rsprice!�ּ� * Val(mshBill.TextMatrix(lngRow, mBillCol.C_����ϵ��)), mFMT.FM_���ۼ�))
                dbl���۽�� = Val(Format(dbl���ۼ� * dbl����, mFMT.FM_���))
                dbl��� = Val(Format(dbl���۽�� - dbl�ɱ����, mFMT.FM_���))
            End If

            rsprice.Filter = "����='�ɱ���' And ����id=" & lng����ID & " And ����=" & Val(mshBill.TextMatrix(lngRow, mBillCol.c_����))
            If rsprice.RecordCount > 0 Then
                blnAdj = True
                dbl���۽�� = Val(Format(dbl���ۼ� * dbl����, mFMT.FM_���))
                dbl�ɱ��� = Val(Format(rsprice!�ּ� * Val(mshBill.TextMatrix(lngRow, mBillCol.C_����ϵ��)), mFMT.FM_���))
                dbl�ɱ���� = Val(Format(dbl�ɱ��� * dbl����, mFMT.FM_���))
                dbl��� = Val(Format(dbl���۽�� - dbl�ɱ����, mFMT.FM_���))
            End If

            If blnAdj = True Then
                '�Ե�ǰ���¼۸����µ���������ݣ��ۼۡ��ɱ��ۡ����۽��ɱ�����ۣ�
                mshBill.TextMatrix(lngRow, mBillCol.C_�ۼ�) = Format(dbl���ۼ�, mFMT.FM_���ۼ�)
                mshBill.TextMatrix(lngRow, mBillCol.C_�ۼ۽��) = Format(dbl���۽��, mFMT.FM_���)
                mshBill.TextMatrix(lngRow, mBillCol.C_�ɹ���) = Format(dbl�ɱ���, mFMT.FM_�ɱ���)
                mshBill.TextMatrix(lngRow, mBillCol.C_�ɹ����) = Format(dbl�ɱ����, mFMT.FM_���)
                mshBill.TextMatrix(lngRow, mBillCol.C_���) = Format(dbl���, mFMT.FM_���)
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


Private Sub Form_Load()
    Dim strStock As String
    Dim rsStock As New Recordset
    Dim strReg As String
    
    mblnUpdate = False
    strReg = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngModule, "0"))
    mintUnit = Val(strReg)
    mbln����˲� = IIf((zlDatabase.GetPara("������Ҫ�˲������ƿ�", glngSys, mlngModule, "0")) = 0, False, True)
  
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
    End With
    
    
    txtNO = mstr���ݺ�
    txtNO.Tag = mstr���ݺ�
    With cboStock
        .Clear
        Do While Not rsDepend.EOF
            If mlngStockID <> rsDepend!Id Then
                .AddItem rsDepend!����
                .ItemData(.NewIndex) = rsDepend!Id
            End If
            rsDepend.MoveNext
        Loop
        .ListIndex = 0
    End With
    mstrTime_Start = GetBillInfo(19, mstr���ݺ�)
    
    Call initCard
    '�ָ����Ի���������
    RestoreWinState Me, App.ProductName, mstrCaption
    '�ָ����Ի��������ú󣬻���Ҫ��Ȩ�޿��Ƶ��н�һ������
    With mshBill
        .ColWidth(mBillCol.C_�ɹ���) = IIf(mblnCostView = True, 1000, 0)
        .ColWidth(mBillCol.C_�ɹ����) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(mBillCol.C_���) = IIf(mblnCostView = True, 800, 0)
    End With
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsInitCard As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim strUnitQuantity_Stock As String
    Dim intRow As Integer
    Dim varStuff As Variant
    Dim numUseAbleCount As Double
    Dim lngStockID  As Long
    Dim intCount As Integer
    '�ⷿ
    On Error GoTo ErrHandle
   With cboStock
        If Not (mint�༭״̬ = 1 Or mint�༭״̬ = 5) Then
            'ȡָ�����ݵĳ���ⷿ�����ⷿ
            gstrSQL = " Select �ⷿID,�Է�����ID From ҩƷ�շ���¼" & _
                      " Where NO=[1] And ����=19 And ���ϵ��=-1 And Rownum<2"
            Set rsInitCard = zlDatabase.OpenSQLRecord(gstrSQL, "ȡָ�����ݵĳ���ⷿ�����ⷿ", mstr���ݺ�)
                      
            If rsInitCard.RecordCount <> 0 Then
                lngStockID = rsInitCard!�ⷿID
            End If
        End If
        For intCount = 0 To .ListCount - 1
            If .ItemData(intCount) = lngStockID Then
                .ListIndex = intCount: Exit For
            End If
        Next
        mintcboIndex = .ListIndex
    End With
    
    
    Select Case mint�༭״̬
        Case 1, 5
            Txt������ = gstrUserName
            Txt�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        Case 2, 3, 4, 6, 7
            initGrid

                        
            Select Case mintUnit
                Case 0
                    strUnitQuantity = "D.���㵥λ AS ��λ, A.��д����,a.ʵ������,a.�ɱ���,a.���ۼ�,'1' as ����ϵ��,"
                    strUnitQuantity_Stock = "Z.��������,Z.ʵ�ʽ��,Z.ʵ�ʲ��"
                Case Else
                    strUnitQuantity = "B.��װ��λ AS ��λ,(A.��д���� / B.����ϵ��) AS ��д����,(A.ʵ������ / B.����ϵ��) AS ʵ������,a.�ɱ���*B.����ϵ�� as �ɱ���,a.���ۼ�*B.����ϵ�� as ���ۼ�,B.����ϵ�� as ����ϵ��,"
                    strUnitQuantity_Stock = "Z.��������/B.����ϵ�� As ��������,Z.ʵ�ʽ��,Z.ʵ�ʲ��"
            End Select
            
            
            
            If mint�༭״̬ <> 7 Then
                gstrSQL = "" & _
                    "   SELECT DISTINCT A.ҩƷID ����id,A.���,('['||D.����||']'||D.����) AS ������Ϣ," & _
                    "                   B.������Դ,D.���,D.���� AS ԭ����,A.����,A.��׼�ĺ�,A.����,A.����,B.ָ�������,B.�ⷿ���� ," & _
                    "                   B.���Ч��,A.Ч��,A.���Ч��,A.��д���� as ԭʼ����," & strUnitQuantity & _
                    "                   A.�ɱ����,A.���۽��, A.���, " & strUnitQuantity_Stock & _
                    "                   ,A.ժҪ,������,��������,�����,�������,A.�ⷿID,A.�Է�����ID,D.�Ƿ���,B.���÷��� " & _
                    "   FROM ҩƷ�շ���¼ A, �������� B,�շ���ĿĿ¼ D, " & _
                    "       (   SELECT ҩƷID ����ID,NVL(����,0) ����,��������,ʵ�ʽ��,ʵ�ʲ�� " & _
                    "           FROM ҩƷ��� WHERE �ⷿID=[2] AND ����=1) Z " & _
                    " WHERE A.ҩƷID = B.����ID AND b.����ID=D.ID " & _
                    "       AND A.���� = 19 AND A.���ϵ��=-1 AND A.NO =[1] AND A.��¼״̬ =[3]" & _
                    "       AND A.ҩƷID=Z.����ID(+) AND NVL(A.����,0)=Z.����(+) " & _
                    " ORDER BY A.��� "
            Else
                gstrSQL = "" & _
                    "   SELECT W.*,Z.��������/W.����ϵ�� AS  ��������,Z.ʵ�ʽ��,Z.ʵ�ʲ�� " & _
                    "   FROM (" & _
                    "           SELECT DISTINCT A.ҩƷID ����ID,A.���,('['||D.����||']'||D.����) AS ������Ϣ," & _
                    "                   B.������Դ,D.���,D.���� AS ԭ����,A.����,A.��׼�ĺ�, A.����,A.����,B.ָ�������,B.�ⷿ���� ," & _
                    "                   B.���Ч��,A.Ч��,A.���Ч��,A.��д���� as ԭʼ����," & strUnitQuantity & _
                    "                   0 �ɱ����,0 ���۽��, 0 ���,A.ժҪ,A.�ⷿID,A.�Է�����ID,D.�Ƿ���,B.���÷���" & _
                    "           FROM ( " & _
                    "                   SELECT MIN(ID) AS ID, SUM(ʵ������) AS ��д����,0 ʵ������,SUM(�ɱ����) AS �ɱ����," & _
                    "                           ҩƷID,���,����,��׼�ĺ�, ����,Ч��,���Ч��,NVL(����,0) ����,����,�ɱ���,���ۼ�,ժҪ,�ⷿID,�Է�����ID,������ID" & _
                    "                   FROM ҩƷ�շ���¼ X " & _
                    "                   WHERE NO=[1] AND ����=19 AND ���ϵ��=-1 " & _
                    "                   GROUP BY ҩƷID,���,����,��׼�ĺ�,����,Ч��,���Ч��,NVL(����,0),����,�ɱ���,���ۼ�,ժҪ,�ⷿID,�Է�����ID,������ID" & _
                    "                   Having SUM(ʵ������)<>0 ) A," & _
                    "               �������� B,�շ���ĿĿ¼ D" & _
                    "           WHERE A.ҩƷID = B.����ID AND B.����ID=D.ID ) W," & _
                    "           (   SELECT  ҩƷID,NVL(����,0) ����,��������,ʵ�ʽ��,ʵ�ʲ�� " & _
                    "               FROM ҩƷ��� WHERE �ⷿID=[2] AND ����=1) Z " & _
                    "   WHERE W.����ID=Z.ҩƷID(+) AND NVL(W.����,0)=Z.����(+) " & _
                    "   ORDER BY ���"
            End If
            
            Set rsInitCard = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr���ݺ�, lngStockID, mint��¼״̬)
               
            
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            
            
            If mint�༭״̬ = 7 Then
                Txt������ = gstrUserName
                Txt�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                Txt����� = gstrUserName
                Txt������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            Else
                Txt������ = rsInitCard!������
                If mint�༭״̬ = 2 Then
                    Txt������ = gstrUserName
                End If
                Txt�������� = Format(rsInitCard!��������, "yyyy-mm-dd hh:mm:ss")
                Txt����� = IIf(IsNull(rsInitCard!�����), "", rsInitCard!�����)
                Txt������� = IIf(IsNull(rsInitCard!�������), "", Format(rsInitCard!�������, "yyyy-mm-dd hh:mm:ss"))
            End If
            txtժҪ.Text = IIf(IsNull(rsInitCard!ժҪ), "", rsInitCard!ժҪ)
            
            If (mint�༭״̬ = 2 Or mint�༭״̬ = 3) And Txt����� <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            
            If mint�༭״̬ = 2 Then
                Set mcolUsedCount = New Collection
            End If
            
            With mshBill
                Do While Not rsInitCard.EOF
                    intRow = rsInitCard.AbsolutePosition
                    'IntRow = rsInitCard!���
                    .Rows = intRow + 1
                    .TextMatrix(intRow, 0) = rsInitCard.Fields(0)
                    .TextMatrix(intRow, mBillCol.C_����) = rsInitCard!������Ϣ
                    .TextMatrix(intRow, mBillCol.c_���) = IIf(IsNull(rsInitCard!���), "", rsInitCard!���)
                    .TextMatrix(intRow, mBillCol.C_���) = zlStr.NVL(rsInitCard!���)
                    
                    .TextMatrix(intRow, mBillCol.C_����) = IIf(IsNull(rsInitCard!����), "", rsInitCard!����)
                    .TextMatrix(intRow, mBillCol.C_��׼�ĺ�) = IIf(IsNull(rsInitCard!��׼�ĺ�), "", rsInitCard!��׼�ĺ�)
                    .TextMatrix(intRow, mBillCol.c_��λ) = rsInitCard!��λ
                    .TextMatrix(intRow, mBillCol.c_����) = IIf(IsNull(rsInitCard!����), "", rsInitCard!����)
                    .TextMatrix(intRow, mBillCol.C_Ч��) = IIf(IsNull(rsInitCard!Ч��), "", Format(rsInitCard!Ч��, "yyyy-mm-dd"))
                    .TextMatrix(intRow, mBillCol.C_���ʧЧ��) = IIf(IsNull(rsInitCard!���Ч��), "", Format(rsInitCard!���Ч��, "yyyy-mm-dd"))
                                
                    .TextMatrix(intRow, mBillCol.C_��д����) = Format(rsInitCard!��д����, mFMT.FM_����)
                    .TextMatrix(intRow, mBillCol.C_ʵ������) = Format(rsInitCard!ʵ������, mFMT.FM_����)
                                
                    .TextMatrix(intRow, mBillCol.C_�ɹ���) = Format(rsInitCard!�ɱ���, mFMT.FM_�ɱ���)
                    .TextMatrix(intRow, mBillCol.C_�ɹ����) = Format(rsInitCard!�ɱ����, mFMT.FM_���)
                    .TextMatrix(intRow, mBillCol.C_�ۼ�) = Format(rsInitCard!���ۼ�, mFMT.FM_���ۼ�)
                    .TextMatrix(intRow, mBillCol.C_�ۼ۽��) = Format(rsInitCard!���۽��, mFMT.FM_���)
                    .TextMatrix(intRow, mBillCol.C_���) = Format(rsInitCard!���, mFMT.FM_���)
                    
                    .TextMatrix(intRow, mBillCol.C_���Ч��) = IIf(IsNull(rsInitCard!���Ч��), "0", rsInitCard!���Ч��) & "||" & rsInitCard!�Ƿ��� & "||" & rsInitCard!���÷���
                    .TextMatrix(intRow, mBillCol.c_����) = IIf(IsNull(rsInitCard!����), "0", rsInitCard!����)
                    .TextMatrix(intRow, mBillCol.C_����ϵ��) = rsInitCard!����ϵ��
                    .TextMatrix(intRow, mBillCol.C_ָ�������) = rsInitCard!ָ�������
                    .TextMatrix(intRow, mBillCol.C_�ⷿ����) = IIf(IsNull(rsInitCard!�ⷿ����), "0", rsInitCard!�ⷿ����)
                    .TextMatrix(intRow, mBillCol.C_��������) = IIf(IsNull(rsInitCard!��������), "0", rsInitCard!��������)
                    .TextMatrix(intRow, mBillCol.C_ʵ�ʲ��) = IIf(IsNull(rsInitCard!ʵ�ʲ��), "0", rsInitCard!ʵ�ʲ��)
                    .TextMatrix(intRow, mBillCol.C_ʵ�ʽ��) = IIf(IsNull(rsInitCard!ʵ�ʽ��), "0", rsInitCard!ʵ�ʽ��)
                    .TextMatrix(intRow, mBillCol.c_ԭʼ����) = Val(zlStr.NVL(rsInitCard!ԭʼ����))
                    
                    If mint�༭״̬ = 2 Then
                        numUseAbleCount = 0
                        For Each varStuff In mcolUsedCount
                            If varStuff(0) = CStr(rsInitCard!����ID & IIf(IsNull(rsInitCard!����), "0", rsInitCard!����)) Then
                                numUseAbleCount = varStuff(1)
                                mcolUsedCount.Remove varStuff(0)
                                Exit For
                            End If
                        Next
                        mcolUsedCount.Add Array(CStr(rsInitCard!����ID & IIf(IsNull(rsInitCard!����), "0", rsInitCard!����)), CStr(numUseAbleCount + IIf(IsNull(rsInitCard!��д����), "0", rsInitCard!��д����))), CStr(rsInitCard!����ID) & CStr(IIf(IsNull(rsInitCard!����), "0", rsInitCard!����))
                        
                    End If
                    
                    rsInitCard.MoveNext
                Loop
            End With
            rsInitCard.Close
    End Select
    
    Call get�������
    Call RefreshRowNO(mshBill, mBillCol.C_�к�, 1)
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
        .Cols = mBillCols
        
        .MsfObj.FixedCols = 1
        
        .TextMatrix(0, mBillCol.C_�к�) = ""
        .TextMatrix(0, mBillCol.C_���) = "���"
                
        .TextMatrix(0, mBillCol.C_����) = "���������"
        .TextMatrix(0, mBillCol.c_���) = "���"
        .TextMatrix(0, mBillCol.C_����) = "����"
        .TextMatrix(0, mBillCol.C_��׼�ĺ�) = "��׼�ĺ�"
        .TextMatrix(0, mBillCol.c_��λ) = "��λ"
        .TextMatrix(0, mBillCol.c_����) = "����"
        .TextMatrix(0, mBillCol.C_Ч��) = "Ч��"
        .TextMatrix(0, mBillCol.C_���ʧЧ��) = "���ʧЧ��"
        
        .TextMatrix(0, mBillCol.C_��ǰ���) = "��ǰ���"
        .TextMatrix(0, mBillCol.C_�Է����) = "�Է����"
        
        .TextMatrix(0, mBillCol.C_��д����) = IIf(mint�༭״̬ = 7, "����", "��д����")
        .TextMatrix(0, mBillCol.C_ʵ������) = IIf(mint�༭״̬ = 7, "��������", "ʵ������")
        .TextMatrix(0, mBillCol.c_ԭʼ����) = "ԭʼ����"
    
        .TextMatrix(0, mBillCol.C_�ɹ���) = "�ɱ���"
        .TextMatrix(0, mBillCol.C_�ɹ����) = "�ɱ����"
        .TextMatrix(0, mBillCol.C_�ۼ�) = "�ۼ�"
        .TextMatrix(0, mBillCol.C_�ۼ۽��) = "�ۼ۽��"
        .TextMatrix(0, mBillCol.C_���) = "���"
        
        .TextMatrix(0, mBillCol.C_��������) = "��������"
        .TextMatrix(0, mBillCol.C_�ⷿ����) = "�ⷿ����"
        .TextMatrix(0, mBillCol.C_���Ч��) = "���Ч��"
        .TextMatrix(0, mBillCol.C_ʵ�ʲ��) = "ʵ�ʲ��"
        .TextMatrix(0, mBillCol.C_ʵ�ʽ��) = "ʵ�ʽ��"
        .TextMatrix(0, mBillCol.C_ָ�������) = "ָ�������"
        .TextMatrix(0, mBillCol.C_����ϵ��) = "����ϵ��"
        .TextMatrix(0, mBillCol.c_����) = "����"
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mBillCol.C_�к�) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mBillCol.C_���) = 0
        
        .ColWidth(mBillCol.C_�к�) = 300
        .ColWidth(mBillCol.C_����) = 2200
        .ColWidth(mBillCol.c_���) = 900
        .ColWidth(mBillCol.C_����) = 800
        .ColWidth(mBillCol.C_��׼�ĺ�) = 1000
        .ColWidth(mBillCol.c_��λ) = 400
        .ColWidth(mBillCol.c_����) = 800
        .ColWidth(mBillCol.C_Ч��) = 1000
        .ColWidth(mBillCol.C_���ʧЧ��) = 1000
        .ColWidth(mBillCol.C_��ǰ���) = 1100
        .ColWidth(mBillCol.C_�Է����) = 1100
        .ColWidth(mBillCol.C_��д����) = 1100
        .ColWidth(mBillCol.C_ʵ������) = 1100
        .ColWidth(mBillCol.C_�ɹ���) = IIf(mblnCostView = False, 0, 1000)
        .ColWidth(mBillCol.C_�ɹ����) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mBillCol.C_�ۼ�) = 1000
        .ColWidth(mBillCol.C_�ۼ۽��) = 900
        .ColWidth(mBillCol.C_���) = IIf(mblnCostView = False, 0, 800)
        .ColWidth(mBillCol.c_ԭʼ����) = 0
        
        .ColWidth(mBillCol.C_�ⷿ����) = 0
        .ColWidth(mBillCol.C_��������) = 0
        .ColWidth(mBillCol.C_���Ч��) = 0
        .ColWidth(mBillCol.C_ʵ�ʲ��) = 0
        .ColWidth(mBillCol.C_ʵ�ʽ��) = 0
        .ColWidth(mBillCol.C_ָ�������) = 0
        .ColWidth(mBillCol.C_����ϵ��) = 0
        .ColWidth(mBillCol.c_����) = 0
        
        
        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��

        .ColData(0) = 5
        .ColData(mBillCol.C_���) = 5
        .ColData(mBillCol.C_�к�) = 5
        .ColData(mBillCol.c_���) = 5
        .ColData(mBillCol.C_����) = 5
        .ColData(mBillCol.C_��׼�ĺ�) = 5
        .ColData(mBillCol.c_��λ) = 5
        .ColData(mBillCol.c_����) = 5
        .ColData(mBillCol.C_Ч��) = 5
        .ColData(mBillCol.C_���ʧЧ��) = 5
        .ColData(mBillCol.c_ԭʼ����) = 5
        .ColData(mBillCol.C_��ǰ���) = 5
        .ColData(mBillCol.C_�Է����) = 5
        
        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Or mint�༭״̬ = 5 Then
            cboStock.Enabled = True
            txtժҪ.Enabled = True
            .ColData(mBillCol.C_����) = 1
            .ColData(mBillCol.C_��д����) = 4
            .ColData(mBillCol.C_ʵ������) = 5
        ElseIf mint�༭״̬ = 3 Or mint�༭״̬ = 4 Or mint�༭״̬ = 6 Or mint�༭״̬ = 7 Then
            cboStock.Enabled = False
            txtժҪ.Enabled = False
            .ColData(mBillCol.C_��д����) = 5
            .ColData(mBillCol.C_ʵ������) = IIf(mint�༭״̬ <> 6, 4, 5)
            .ColData(mBillCol.C_����) = 0
        End If
        
        
        .ColData(mBillCol.C_�ɹ���) = 5
        .ColData(mBillCol.C_�ɹ����) = 5
        .ColData(mBillCol.C_�ۼ�) = 5
        .ColData(mBillCol.C_�ۼ۽��) = 5
        .ColData(mBillCol.C_���) = 5
        
        .ColData(mBillCol.C_�ⷿ����) = 5
        .ColData(mBillCol.C_��������) = 5
        .ColData(mBillCol.C_���Ч��) = 5
        .ColData(mBillCol.C_ʵ�ʲ��) = 5
        .ColData(mBillCol.C_ʵ�ʽ��) = 5
        .ColData(mBillCol.C_ָ�������) = 5
        .ColData(mBillCol.C_����ϵ��) = 5
        .ColData(mBillCol.c_����) = 5
        
        .ColAlignment(mBillCol.C_����) = flexAlignLeftCenter
        .ColAlignment(mBillCol.c_���) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_����) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_��׼�ĺ�) = flexAlignLeftCenter
        .ColAlignment(mBillCol.c_��λ) = flexAlignCenterCenter
        .ColAlignment(mBillCol.c_����) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_Ч��) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_��ǰ���) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_�Է����) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_��д����) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_ʵ������) = flexAlignRightCenter
        
        .ColAlignment(mBillCol.C_�ɹ���) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_�ɹ����) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_�ۼ�) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_�ۼ۽��) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_���) = flexAlignRightCenter
        
        .PrimaryCol = mBillCol.C_����
        .LocateCol = mBillCol.C_����
        If InStr(1, "34", mint�༭״̬) <> 0 Then .ColData(mBillCol.C_����) = 0
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
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top - 100 - cmdCancel.Height - 200
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
    End With
    
    
    LblStock.Left = mshBill.Left
    cboStock.Left = LblStock.Left + LblStock.Width + 100
    
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
        '.Width = .Left - .Left
        Debug.Print .Width
    End With
        
    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = txtժҪ.Top - 60 - .Height
        .Width = mshBill.Width
        lblSalePrice.Top = .Top
        lblDifference.Top = .Top
    End With
    If mblnCostView = False Then
        lblPurchasePrice.Visible = False
    End If
    
    With lblSalePrice
        .Left = lblPurchasePrice.Left + mshBill.Width / 3
    End With
    With lblDifference
        .Left = lblPurchasePrice.Left + mshBill.Width / 3 * 2
    End With
    If mblnCostView = False Then
        lblDifference.Visible = False
    End If
    
    With mshBill
        .Height = lblPurchasePrice.Top - .Top - 60
    End With
    
    With cmdCancel
        .Left = Pic����.Left + mshBill.Left + mshBill.Width - .Width
        .Top = Pic����.Top + Pic����.Height + 100
    End With
    
    With CmdSave
        .Left = cmdCancel.Left - .Width - 100
        .Top = cmdCancel.Top
    End With
    
    With cmdHelp
        .Left = Pic����.Left + mshBill.Left
        .Top = cmdCancel.Top
    End With
    
    With cmdAllCls
        .Left = CmdSave.Left - .Width - 500
        .Top = cmdCancel.Top
    End With
    
    With cmdAllSel
        .Left = cmdAllCls.Left - .Width - 100
        .Top = cmdCancel.Top
    End With
        
    With cmdFind
        .Top = cmdCancel.Top
    End With
    
    With cmdRequest
        .Top = cmdFind.Top
        
        .Visible = (mint�༭״̬ = 1 Or mint�༭״̬ = 2) '�������޸Ĳſɼ�
        
    End With
    
    With lblCode
        .Top = cmdCancel.Top + 50
    End With
    With txtCode
        .Top = cmdCancel.Top + 30
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    
    If mblnChange = False Or mint�༭״̬ = 4 Then
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

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mBillCol.C_�к�, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call ��ʾ�ϼƽ��
    Call RefreshRowNO(mshBill, mBillCol.C_�к�, mshBill.Row)
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
    
    mint����ʾ�п������ = gSystem_Para.para_������¿��ÿ�� And mint����� = 2
    
    Set RecReturn = Frm����ѡ����.ShowMe(Me, 2, cboStock.ItemData(cboStock.ListIndex), _
        mlngStockID, mlngStockID, IIf(mint����� = 0, False, IIf(mint��ȷ���� = 0, False, True)), IIf(mint��ȷ���� = 0, False, True), _
        False, False, (InStr(1, mstrPrivs, "��ʾ�Է����")), , , , , mint����ʾ�п������, , , mstrPrivs, IIf(mint��ȷ���� = 0, False, True), False)
    If RecReturn.RecordCount > 0 Then
    
        With mshBill
            RecReturn.MoveFirst
            For i = 1 To RecReturn.RecordCount
                mblnChange = True
                
                If SetColValue(.Row, RecReturn!����ID, "[" & RecReturn!���� & "]" & RecReturn!����, _
                    IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                    IIf(mintUnit = 0, RecReturn!ɢװ��λ, RecReturn!��װ��λ), _
                    IIf(IsNull(RecReturn!�ۼ�), 0, RecReturn!�ۼ�), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                    IIf(IsNull(RecReturn!Ч��), "", RecReturn!Ч��), _
                    IIf(IsNull(RecReturn!���Ч��), "0", RecReturn!���Ч��), _
                    IIf(RecReturn!һ���Բ��� = 1, True, False), _
                    IIf(IsNull(RecReturn!���ʧЧ��), "", RecReturn!���ʧЧ��), _
                    RecReturn!�ⷿ����, _
                    IIf(IsNull(RecReturn!��������), "0", RecReturn!��������), _
                    IIf(IsNull(RecReturn!ʵ�ʽ��), "0", RecReturn!ʵ�ʽ��), _
                    IIf(IsNull(RecReturn!ʵ�ʲ��), "0", RecReturn!ʵ�ʲ��), _
                    IIf(IsNull(RecReturn!ָ�������), "0", RecReturn!ָ�������), _
                    IIf(mintUnit = 0, 1, RecReturn!����ϵ��), IIf(IsNull(RecReturn!����), 0, RecReturn!����), RecReturn!ʱ��, RecReturn!���÷���, IIf(IsNull(RecReturn!��׼�ĺ�), "", RecReturn!��׼�ĺ�)) Then
                
                    If .Row = .Rows - 1 Then .Rows = .Rows + 1 'ֻ�е�ǰ�������һ��ʱ��������
                    .Row = .Row + 1
                End If
                
                .Col = mBillCol.C_��д����
                RecReturn.MoveNext
            Next
            
            mshBill.Row = int�����
            
            If mstr�ظ����� <> "" Then
                MsgBox mstr�ظ����� & "�б����Ѿ������ˣ�" & vbCrLf & "�������Ĳ�����ӣ�", vbInformation + vbOKOnly, gstrSysName
                mstr�ظ����� = ""
            End If
        
'            If RecReturn.RecordCount = 1 Then
'                mblnChange = True
'
'                SetColValue .Row, RecReturn!����ID, "[" & RecReturn!���� & "]" & RecReturn!����, _
'                    IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
'                    IIf(mintUnit = 0, RecReturn!ɢװ��λ, RecReturn!��װ��λ), _
'                    IIf(IsNull(RecReturn!�ۼ�), 0, RecReturn!�ۼ�), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
'                    IIf(IsNull(RecReturn!Ч��), "", RecReturn!Ч��), _
'                    IIf(IsNull(RecReturn!���Ч��), "0", RecReturn!���Ч��), _
'                    IIf(RecReturn!һ���Բ��� = 1, True, False), _
'                    IIf(IsNull(RecReturn!���ʧЧ��), "", RecReturn!���ʧЧ��), _
'                    RecReturn!�ⷿ����, _
'                    IIf(IsNull(RecReturn!��������), "0", RecReturn!��������), _
'                    IIf(IsNull(RecReturn!ʵ�ʽ��), "0", RecReturn!ʵ�ʽ��), _
'                    IIf(IsNull(RecReturn!ʵ�ʲ��), "0", RecReturn!ʵ�ʲ��), _
'                    IIf(IsNull(RecReturn!ָ�������), "0", RecReturn!ָ�������), _
'                    IIf(mintUnit = 0, 1, RecReturn!����ϵ��), IIf(IsNull(RecReturn!����), 0, RecReturn!����), RecReturn!ʱ��, RecReturn!���÷���, IIf(IsNull(RecReturn!��׼�ĺ�), "", RecReturn!��׼�ĺ�)
'                .Col = mBillCol.C_��д����
'            End If
        End With
        RecReturn.Close
    End If
End Sub


Private Sub mshbill_EditChange(curText As String)
    With mshBill
        If .Col <> mBillCol.C_���� Then
            mshBill.Text = UCase(curText)
            mshBill.SelStart = Len(mshBill.Text)
        End If
    End With
    mblnChange = True
End Sub


Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer
    
    With mshBill
        If .Col = mBillCol.C_��д���� Or .Col = mBillCol.C_ʵ������ Then
            strKey = .Text
            If strKey = "" Then
                strKey = .TextMatrix(.Row, .Col)
            End If
            Select Case .Col
                Case mBillCol.C_��д����, mBillCol.C_ʵ������
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
            Case mBillCol.C_����
                .TxtCheck = False
                .MaxLength = 80
                'ֻ��ҩ���в���ʾ�ϼ���Ϣ�Ϳ����
                Call ��ʾ�ϼƽ��
'                Call ��ʾ�����
                
            Case mBillCol.c_����
                .TxtCheck = True
                .TextMask = "1234567890"
                .MaxLength = 8
            
            Case mBillCol.C_Ч��
                .TxtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                If .TextMatrix(.Row, mBillCol.c_����) <> "" And .ColData(.Col) = 2 Then
                    Dim strxq As String
                    
                    If IsNumeric(.TextMatrix(.Row, mBillCol.c_����)) And .TextMatrix(.Row, mBillCol.C_���Ч��) <> "" Then
                        If Split(.TextMatrix(.Row, mBillCol.C_���Ч��), "||")(0) <> 0 Then
                            strxq = .TextMatrix(.Row, mBillCol.c_����)
                            strxq = TranNumToDate(strxq)
                            If strxq = "" Then Exit Sub
                            
                            .TextMatrix(.Row, mBillCol.C_Ч��) = Format(DateAdd("M", Split(.TextMatrix(.Row, mBillCol.C_���Ч��), "||")(0), strxq), "yyyy-mm-dd")
                        End If
                    End If
                End If
            Case mBillCol.C_��д����, mBillCol.C_ʵ������
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = "-.1234567890"
                
        End Select
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsStuff As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim i As Integer
    Dim int����� As Integer
    
    int����� = mshBill.Row
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With mshBill
        strKey = UCase(Trim(.Text))
        
        If Mid(strKey, 1, 1) = "[" Then
            If InStr(2, strKey, "]") <> 0 Then
                strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
            Else
                strKey = Mid(strKey, 2)
            End If
        End If
        Select Case .Col
            
            Case mBillCol.C_����
                If strKey <> "" Then
                    Dim RecReturn As Recordset
                    Dim sngLeft As Single
                    Dim sngTop As Single
                    

                    sngLeft = Me.Left + Pic����.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic����.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 3630
                    End If
                    
                    mint����ʾ�п������ = gSystem_Para.para_������¿��ÿ�� And mint����� = 2
                    Set RecReturn = FrmMulitSel.ShowSelect(Me, 2, cboStock.ItemData(cboStock.ListIndex), _
                        mlngStockID, mlngStockID, strKey, sngLeft, sngTop, mshBill.MsfObj.CellWidth, mshBill.MsfObj.CellHeight, IIf(mint����� = 0, False, IIf(mint��ȷ���� = 0, False, True)), _
                        IIf(mint��ȷ���� = 0, False, True), False, False, (InStr(1, mstrPrivs, "��ʾ�Է����")), , , , mint����ʾ�п������, , , mstrPrivs, IIf(mint��ȷ���� = 0, False, True), False)
                        
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
                                IIf(IsNull(RecReturn!Ч��), "", RecReturn!Ч��), _
                                IIf(IsNull(RecReturn!���Ч��), "0", RecReturn!���Ч��), _
                                IIf(zlStr.NVL(RecReturn!һ���Բ���, 0) = 1, True, False), _
                                IIf(zlStr.NVL(RecReturn!���ʧЧ��) = "", "", Format(RecReturn!���ʧЧ��, "yyyy-mm-dd")), _
                                RecReturn!�ⷿ����, _
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
'
'                    If RecReturn.RecordCount = 1 Then
'                        If SetColValue(.Row, RecReturn!����ID, "[" & RecReturn!���� & "]" & RecReturn!����, _
'                                IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
'                                IIf(mintUnit = 0, RecReturn!ɢװ��λ, RecReturn!��װ��λ), _
'                                IIf(IsNull(RecReturn!�ۼ�), 0, RecReturn!�ۼ�), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
'                                IIf(IsNull(RecReturn!Ч��), "", RecReturn!Ч��), _
'                                IIf(IsNull(RecReturn!���Ч��), "0", RecReturn!���Ч��), _
'                                IIf(zlStr.NVL(RecReturn!һ���Բ���, 0) = 1, True, False), _
'                                IIf(zlStr.NVL(RecReturn!���ʧЧ��) = "", "", Format(RecReturn!���ʧЧ��, "yyyy-mm-dd")), _
'                                RecReturn!�ⷿ����, _
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
'                    Call ��ʾ�����
                End If
            Case mBillCol.c_����
                '�޴���
                If strKey = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mBillCol.c_����) = ""
                    End If
                    If .ColData(mBillCol.C_Ч��) = 2 Then
                        .Col = mBillCol.C_Ч��
                    Else
                        .Col = mBillCol.C_��д����
                    End If
                    
                    
                    Cancel = True
                    Exit Sub
                End If
                
                If Len(strKey) < 8 Then
                    MsgBox "���ų��Ȳ���������Ϊ8λ,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
            Case mBillCol.C_Ч��
                '�д���
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        If strKey = "" Then
                            MsgBox "Ч�ڱ���Ϊ�����ͣ�", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            Exit Sub
                        End If
                        .Text = strKey
                        Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "Ч�ڱ���Ϊ��������(2000-10-10) ��20001010��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                ElseIf strKey = "" And strKey <> .TextMatrix(.Row, mBillCol.C_Ч��) Then
                
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    End If
                    
                    Exit Sub
                End If
            
            Case mBillCol.C_��д����, mBillCol.C_ʵ������
                If .TextMatrix(.Row, 0) = "" Then .Text = "": .TextMatrix(.Row, mBillCol.C_��д����) = "": Exit Sub
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
                    If Val(strKey) = 0 Then
                        MsgBox "��������Ϊ��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) < 0.001 Then
                        MsgBox "�����������0.001,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "��������С��" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
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
                    
                    If .TextMatrix(.Row, mBillCol.C_�ۼ�) <> "" Then
                        .TextMatrix(.Row, mBillCol.C_�ۼ۽��) = Format(.TextMatrix(.Row, mBillCol.C_�ۼ�) * strKey, mFMT.FM_���)
                    End If
                    
                    If mint�༭״̬ <> 7 Then
                        Dim dbl��� As Double, dbl���� As Double, dbl�ɱ���� As Double
                        'cboStock.ItemData(cboStock.ListIndex)
                        
                        Call ��֤�����ۼ���(cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mBillCol.c_����)), _
                            Val(.TextMatrix(.Row, mBillCol.C_����ϵ��)), Val(.TextMatrix(.Row, mBillCol.C_ʵ�ʲ��)), Val(.TextMatrix(.Row, mBillCol.C_ʵ�ʽ��)), _
                            Val(Split(.TextMatrix(.Row, mBillCol.C_ָ�������), "||")(0)) / 100, Val(strKey), Val(.TextMatrix(.Row, mBillCol.C_�ۼ۽��)), dbl���, dbl����, dbl�ɱ����)
                        .TextMatrix(.Row, mBillCol.C_���) = Format(dbl���, mFMT.FM_���)
                        .TextMatrix(.Row, mBillCol.C_�ɹ���) = Format(dbl����, mFMT.FM_�ɱ���)
                        .TextMatrix(.Row, mBillCol.C_�ɹ����) = Format(dbl�ɱ����, mFMT.FM_���)
                    Else
                        .TextMatrix(.Row, mBillCol.C_�ɹ����) = Format(Val(.TextMatrix(.Row, mBillCol.C_�ɹ���)) * strKey, mFMT.FM_���)
                        .TextMatrix(.Row, mBillCol.C_���) = Format(Val(.TextMatrix(.Row, mBillCol.C_�ۼ۽��)) - Val(.TextMatrix(.Row, mBillCol.C_�ɹ����)), mFMT.FM_���)
                    End If
                    
                    If .Col = mBillCol.C_��д���� Then
                        .TextMatrix(.Row, mBillCol.C_ʵ������) = strKey
                    End If
                End If
                ��ʾ�ϼƽ��
            
        End Select
    End With
End Sub

'�Ӳ���Ŀ¼��ȡֵ��������Ӧ����
Private Function SetColValue(ByVal intRow As Integer, ByVal lng����ID As Long, _
    ByVal str���� As String, ByVal str��� As String, ByVal str���� As String, _
    ByVal str��λ As String, ByVal num�ۼ� As Double, ByVal str���� As String, _
    ByVal strЧ�� As String, ByVal int���Ч�� As Integer, ByVal blnһ���Բ��� As Boolean, _
    ByVal str���ʧЧ�� As String, ByVal int�ⷿ���� As Integer, _
    ByVal num�������� As Double, ByVal numʵ�ʽ�� As Double, ByVal numʵ�ʲ�� As Double, _
    ByVal numָ������� As Double, ByVal num����ϵ�� As Double, ByVal lng���� As Long, _
    ByVal int�Ƿ��� As Integer, ByVal int���÷��� As Integer, ByVal str��׼�ĺ� As String) As Boolean
    
    Dim intCount As Integer
    Dim intCol As Integer
    Dim dblPrice As Double
    Dim rsprice As New Recordset
    Dim bln���� As Boolean
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    If blnһ���Բ��� = True Then
        If Format(str���ʧЧ��, "yyyy-mm-dd") < Format(sys.Currentdate, "yyyy-mm-dd") And Trim(str���ʧЧ��) <> "" Then
           If MsgBox("���ġ�" & str���� & "(" & lng���� & ")���Ѿ��������ʧЧ��,�Ƿ�Ҫ���죡", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) <> vbYes Then
                Exit Function
           End If
        End If
    End If
    
    SetColValue = False
    With mshBill
        Dim lngRow As Long
        For lngRow = 1 To .Rows - 1
            If lngRow <> intRow And .TextMatrix(lngRow, 0) <> "" Then
                If .TextMatrix(lngRow, 0) = lng����ID And Val(.TextMatrix(lngRow, mBillCol.c_����)) = lng���� Then
                    If UBound(Split(mstr�ظ�����, "��")) < 3 Then mstr�ظ����� = mstr�ظ����� & str���� & "��"  '����¼�����ظ�������
                    'Call MsgBox("�������ϡ�" & str���� & "(" & lng���� & ")���Ѿ����ڣ���ϲ��������ӣ�", vbOKOnly + vbInformation + vbDefaultButton2, gstrSysName)
                    Exit Function
                End If
            End If
        Next
        
        If int�Ƿ��� = 1 Then
            If int���÷��� = 0 Then
                If int�ⷿ���� = 1 Then
                    gstrSQL = "Select Distinct 0 " & _
                            "From ��������˵�� " & _
                            "Where ((�������� Like '���ϲ���') Or (�������� Like '�Ƽ���')) And ����id = [1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex))
                    If rsTemp.RecordCount = 0 Then
                        bln���� = True
                    End If
                End If
            Else
                bln���� = True
            End If
        
            gstrSQL = "" & _
                "   Select nvl(���ۼ�,0)*" & num����ϵ�� & " as  �����ۼ�,ʵ�ʽ��/ʵ������* " & num����ϵ�� & " as ƽ�����ۼ�" & _
                "   From ҩƷ��� " & _
                "   Where �ⷿid=[1]" & _
                "       and ҩƷid=[2]" & _
                "       and ����=1 and ʵ������>0 and " & _
                "       nvl(����,0)=[3]"
            
            Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex), lng����ID, lng����)
            If rsprice.EOF Then
                If mint��ȷ���� = 1 Then
                    MsgBox "ʱ������û�п�棬���ܳ��⣬���飡", vbOKOnly, gstrSysName
                    Exit Function
                Else
                    dblPrice = num�ۼ� * num����ϵ��
                End If
            Else
                If bln���� = True Then
                    dblPrice = rsprice!�����ۼ�
                Else
                    dblPrice = rsprice!ƽ�����ۼ�
                End If
            End If
        End If
        
        For intCol = 0 To .Cols - 1
            If intCol <> mBillCol.C_�к� Then .TextMatrix(intRow, intCol) = ""
        Next
        
        .TextMatrix(intRow, mBillCol.C_�к�) = intRow
        .TextMatrix(intRow, 0) = lng����ID
        .TextMatrix(intRow, mBillCol.C_����) = str����
        .TextMatrix(intRow, mBillCol.c_���) = str���
        .TextMatrix(intRow, mBillCol.C_����) = str����
        .TextMatrix(intRow, mBillCol.C_��׼�ĺ�) = str��׼�ĺ�
        .TextMatrix(intRow, mBillCol.c_��λ) = str��λ
        .TextMatrix(intRow, mBillCol.C_�ۼ�) = Format(num�ۼ� * num����ϵ��, mFMT.FM_���ۼ�)
        .TextMatrix(intRow, mBillCol.C_�ⷿ����) = int�ⷿ����
        .TextMatrix(intRow, mBillCol.C_��������) = Format(num�������� / num����ϵ��, mFMT.FM_����)
        .TextMatrix(intRow, mBillCol.C_���Ч��) = int���Ч�� & "||" & int�Ƿ��� & "||" & int���÷���
        .TextMatrix(intRow, mBillCol.C_ʵ�ʲ��) = numʵ�ʲ��
        .TextMatrix(intRow, mBillCol.C_ʵ�ʽ��) = numʵ�ʽ��
        .TextMatrix(intRow, mBillCol.C_ָ�������) = numָ�������
        .TextMatrix(intRow, mBillCol.C_����ϵ��) = num����ϵ��
        If mint��ȷ���� = 1 Then
            .TextMatrix(intRow, mBillCol.c_����) = lng����
            .TextMatrix(intRow, mBillCol.c_����) = str����
            .TextMatrix(intRow, mBillCol.C_Ч��) = Format(strЧ��, "yyyy-mm-dd")
            .TextMatrix(intRow, mBillCol.C_���ʧЧ��) = Format(str���ʧЧ��, "yyyy-mm-dd")
        Else
            .TextMatrix(intRow, mBillCol.c_����) = lng���� '�ֹ��������ģ����δ���Ϊ0�����깺����ȡ�ᴫ��������
            .TextMatrix(intRow, mBillCol.c_����) = ""
            .TextMatrix(intRow, mBillCol.C_Ч��) = ""
            .TextMatrix(intRow, mBillCol.C_���ʧЧ��) = ""
        End If
        '��Ҫ����ʱ�۷����Ͳ��������
        If int�Ƿ��� = 1 Then .TextMatrix(intRow, mBillCol.C_�ۼ�) = Format(dblPrice, mFMT.FM_���ۼ�)
        Call CheckLapse(strЧ��)
        
        Call get�������(intRow)
    End With
'    Call ��ʾ�����
    SetColValue = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'�Ӳ���Ŀ¼��ȡֵ��������Ӧ����
Private Function SetRequestColValue(ByVal intRow As Integer, ByVal lng����ID As Long, _
    ByVal str���� As String, ByVal str��� As String, ByVal str���� As String, _
    ByVal str��λ As String, ByVal num�ۼ� As Double, ByVal str���� As String, _
    ByVal strЧ�� As String, ByVal int���Ч�� As Integer, ByVal blnһ���Բ��� As Boolean, _
    ByVal str���ʧЧ�� As String, ByVal int�ⷿ���� As Integer, _
    ByVal num�������� As Double, ByVal numʵ�ʽ�� As Double, ByVal numʵ�ʲ�� As Double, _
    ByVal numָ������� As Double, ByVal num����ϵ�� As Double, ByVal lng���� As Long, _
    ByVal int�Ƿ��� As Integer, ByVal int���÷��� As Integer, ByVal str��׼�ĺ� As String) As Boolean
    
    Dim intCount As Integer
    Dim intCol As Integer
    Dim dblPrice As Double
    Dim rsprice As New Recordset
    Dim bln���� As Boolean
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    If blnһ���Բ��� = True Then
        If Format(str���ʧЧ��, "yyyy-mm-dd") < Format(sys.Currentdate, "yyyy-mm-dd") And Trim(str���ʧЧ��) <> "" Then
           If MsgBox("���ġ�" & str���� & "(" & lng���� & ")���Ѿ��������ʧЧ��,�Ƿ�Ҫ���죡", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) <> vbYes Then
                Exit Function
           End If
        End If
    End If
    
    SetRequestColValue = False
    With mshBill
        Dim lngRow As Long
        For lngRow = 1 To .Rows - 1
            If lngRow <> intRow And .TextMatrix(lngRow, 0) <> "" Then
                If .TextMatrix(lngRow, 0) = lng����ID And Val(.TextMatrix(lngRow, mBillCol.c_����)) = lng���� Then
                    If UBound(Split(mstr�ظ�����, "��")) < 3 Then mstr�ظ����� = mstr�ظ����� & str���� & "��"  '����¼�����ظ�������
                    'Call MsgBox("�������ϡ�" & str���� & "(" & lng���� & ")���Ѿ����ڣ���ϲ��������ӣ�", vbOKOnly + vbInformation + vbDefaultButton2, gstrSysName)
                    Exit Function
                End If
            End If
        Next
        
        If int�Ƿ��� = 1 Then
            If int���÷��� = 0 Then
                If int�ⷿ���� = 1 Then
                    gstrSQL = "Select Distinct 0 " & _
                            "From ��������˵�� " & _
                            "Where ((�������� Like '���ϲ���') Or (�������� Like '�Ƽ���')) And ����id = [1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex))
                    If rsTemp.RecordCount = 0 Then
                        bln���� = True
                    End If
                End If
            Else
                bln���� = True
            End If
        
            gstrSQL = "" & _
                "   Select nvl(���ۼ�,0)*" & num����ϵ�� & " as  �����ۼ�,ʵ�ʽ��/ʵ������* " & num����ϵ�� & " as ƽ�����ۼ�" & _
                "   From ҩƷ��� " & _
                "   Where �ⷿid=[1]" & _
                "       and ҩƷid=[2]" & _
                "       and ����=1 and ʵ������>0 and " & _
                "       nvl(����,0)=[3]"
            
            Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex), lng����ID, lng����)
            If rsprice.EOF Then
                If mint��ȷ���� = 1 Then
                    MsgBox "ʱ������û�п�棬���ܳ��⣬���飡", vbOKOnly, gstrSysName
                    Exit Function
                Else
                    dblPrice = num�ۼ� * num����ϵ��
                End If
            Else
                If bln���� = True Then
                    dblPrice = rsprice!�����ۼ�
                Else
                    dblPrice = rsprice!ƽ�����ۼ�
                End If
            End If
        End If
        
        For intCol = 0 To .Cols - 1
            If intCol <> mBillCol.C_�к� Then .TextMatrix(intRow, intCol) = ""
        Next
        
        .TextMatrix(intRow, mBillCol.C_�к�) = intRow
        .TextMatrix(intRow, 0) = lng����ID
        .TextMatrix(intRow, mBillCol.C_����) = str����
        .TextMatrix(intRow, mBillCol.c_���) = str���
        .TextMatrix(intRow, mBillCol.C_����) = str����
        .TextMatrix(intRow, mBillCol.C_��׼�ĺ�) = str��׼�ĺ�
        .TextMatrix(intRow, mBillCol.c_��λ) = str��λ
        .TextMatrix(intRow, mBillCol.C_�ۼ�) = Format(num�ۼ� * num����ϵ��, mFMT.FM_���ۼ�)
        .TextMatrix(intRow, mBillCol.C_�ⷿ����) = int�ⷿ����
        .TextMatrix(intRow, mBillCol.C_��������) = Format(num��������, mFMT.FM_����)
        .TextMatrix(intRow, mBillCol.C_���Ч��) = int���Ч�� & "||" & int�Ƿ��� & "||" & int���÷���
        .TextMatrix(intRow, mBillCol.C_ʵ�ʲ��) = numʵ�ʲ��
        .TextMatrix(intRow, mBillCol.C_ʵ�ʽ��) = numʵ�ʽ��
        .TextMatrix(intRow, mBillCol.C_ָ�������) = numָ�������
        .TextMatrix(intRow, mBillCol.C_����ϵ��) = num����ϵ��
        '���깺����������ȷ���εģ������ж��Ƿ���������
        .TextMatrix(intRow, mBillCol.c_����) = lng����
        .TextMatrix(intRow, mBillCol.c_����) = str����
        .TextMatrix(intRow, mBillCol.C_Ч��) = Format(strЧ��, "yyyy-mm-dd")
        .TextMatrix(intRow, mBillCol.C_���ʧЧ��) = Format(str���ʧЧ��, "yyyy-mm-dd")
       
        '��Ҫ����ʱ�۷����Ͳ��������
        If int�Ƿ��� = 1 Then .TextMatrix(intRow, mBillCol.C_�ۼ�) = Format(dblPrice, mFMT.FM_���ۼ�)
        Call CheckLapse(strЧ��)
        
        Call get�������(intRow)
    End With
'    Call ��ʾ�����
    SetRequestColValue = True
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
  
    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '�����з�����
            
            If LenB(StrConv(txtժҪ.Text, vbFromUnicode)) > txtժҪ.MaxLength Then
                MsgBox "ժҪ����,���������" & CInt(txtժҪ.MaxLength / 2) & "�����ֻ�" & txtժҪ.MaxLength & "���ַ�!", vbInformation + vbOKOnly, gstrSysName
                txtժҪ.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .Rows - 1
                If Val(.TextMatrix(intLop, 0)) > 0 And Trim(.TextMatrix(intLop, mBillCol.C_����)) = "" Then
                    MsgBox "��" & intLop & "�����ĵ�����Ϊ���ˣ����飡", vbInformation, gstrSysName
                    mshBill.SetFocus
                    .Row = intLop
                    .MsfObj.TopRow = intLop
                    .Col = mBillCol.C_����
                    Exit Function
                End If
                
                If Trim(.TextMatrix(intLop, mBillCol.C_����)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mBillCol.C_��д����))) = "" Then
                        MsgBox "��" & intLop & "�����ĵ�����Ϊ���ˣ����飡", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mBillCol.C_��д����
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mBillCol.C_��д����)) > 9999999999# Then
                        MsgBox "��" & intLop & "�����ĵ���д�������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mBillCol.C_��д����
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mBillCol.C_ʵ������)) > 9999999999# Then
                        MsgBox "��" & intLop & "�����ĵ�ʵ���������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mBillCol.C_ʵ������
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mBillCol.C_�ɹ����)) > 9999999999999# Then
                        MsgBox "��" & intLop & "�����ĵĳɱ������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = IIf(.ColData(mBillCol.C_��д����) = 4, mBillCol.C_��д����, mBillCol.C_ʵ������)
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mBillCol.C_�ۼ۽��)) > 9999999999999# Then
                        MsgBox "��" & intLop & "�����ĵ��ۼ۽����������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = IIf(.ColData(mBillCol.C_��д����) = 4, mBillCol.C_��д����, mBillCol.C_ʵ������)
                        Exit Function
                    End If
                               
                    '����/�޸�ʱ��������������ֹ����
                    If Not CompareUsableQuantity(intLop, Val(Trim(.TextMatrix(intLop, mBillCol.C_��д����))), True) Then
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mBillCol.C_��д����
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
    Dim chrNo As Variant
    Dim lng��� As Long
    Dim lng�ⷿID As Long
    Dim lng����ID As Long
    Dim lng����ID As Long
    Dim str���� As String
    Dim lng���� As Long
    Dim str���� As String
    Dim strЧ�� As String
    Dim dbl��д���� As Double
    Dim dbl�ɱ��� As Double
    Dim dbl�ɱ���� As Double
    Dim dbl���ۼ� As Double
    Dim dbl���۽�� As Double
    Dim dblʵ������ As Double
    Dim dbl��� As Double
    Dim strժҪ As String
    Dim str������ As String
    Dim str�������� As String
    Dim str�˲����� As String
    Dim str����� As String
    Dim datAssessDate As String
    Dim str���Ч�� As String
    Dim str�˲��� As String
    Dim n As Long
    
    Dim intRow As Integer
    Dim arrSQL As Variant
    
    '�Զ��ֽ������¼ʱʹ��
    Dim blnAuto As Boolean              '�Ƿ���Ҫ�Զ��ֽ�
    Dim dbl��д����_Cur As Double
    Dim rsStock As New ADODB.Recordset
    
    SaveCard = False
    arrSQL = Array()
    
    With mshBill
        chrNo = Trim(txtNO)
        lng�ⷿID = cboStock.ItemData(cboStock.ListIndex)
        If chrNo <> "" Then
            If CheckNOExists(72, chrNo) Then Exit Function
        End If
        
        If chrNo = "" Then chrNo = sys.GetNextNo(72, lng�ⷿID)
        If IsNull(chrNo) Then Exit Function
        txtNO.Tag = chrNo
        
        lng����ID = mlngStockID
        strժҪ = Trim(txtժҪ.Text)
        str������ = Txt������
        str�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        
        If mbln����˲� = True And mint�༭״̬ = 3 Then
            str�������� = Txt��������
        End If
        str�˲����� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        str����� = Txt�����
        
        If mbln����˲� = True And mint�༭״̬ = 3 Then
            str�˲��� = Txt������
        End If
        
        On Error GoTo ErrHandle
        
        If mint�༭״̬ = 2 Or mint�༭״̬ = 3 Then      '�޸ĺͺ˲�
            gstrSQL = "zl_��������_Delete('" & mstr���ݺ� & "')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "0;" & vbCrLf & gstrSQL
        End If
        
        Dim intTmp As Integer
        lng��� = -1
        
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                '�����ǰ�������Ĳ������Զ�ȡ�������ε����ģ�������������¼
                lng����ID = .TextMatrix(intRow, 0)
                str���� = .TextMatrix(intRow, mBillCol.C_����)
                str���� = .TextMatrix(intRow, mBillCol.c_����)
                intTmp = Val(.TextMatrix(intRow, mBillCol.C_����ϵ��))
                lng���� = Val(.TextMatrix(intRow, mBillCol.c_����))
                strЧ�� = IIf(.TextMatrix(intRow, mBillCol.C_Ч��) = "", "", .TextMatrix(intRow, mBillCol.C_Ч��))
                str���Ч�� = IIf(.TextMatrix(intRow, mBillCol.C_���ʧЧ��) = "", "", .TextMatrix(intRow, mBillCol.C_���ʧЧ��))
                dbl��д���� = Round(Val(.TextMatrix(intRow, mBillCol.C_��д����)) * intTmp, g_С��λ��.obj_���С��.����С��)
                dblʵ������ = Round(Val(.TextMatrix(intRow, mBillCol.C_ʵ������)) * intTmp, g_С��λ��.obj_���С��.����С��)
                dbl�ɱ��� = Round(Val(.TextMatrix(intRow, mBillCol.C_�ɹ���)) / IIf(intTmp = 0, 1, intTmp), g_С��λ��.obj_���С��.�ɱ���С��)
                dbl�ɱ���� = Round(Val(.TextMatrix(intRow, mBillCol.C_�ɹ����)), g_С��λ��.obj_���С��.���С��)
                dbl���ۼ� = Round(Val(.TextMatrix(intRow, mBillCol.C_�ۼ�)) / IIf(intTmp = 0, 1, intTmp), g_С��λ��.obj_���С��.���ۼ�С��)
                dbl���۽�� = Round(Val(.TextMatrix(intRow, mBillCol.C_�ۼ۽��)), g_С��λ��.obj_���С��.���С��)
                dbl��� = Round(Val(.TextMatrix(intRow, mBillCol.C_���)), g_С��λ��.obj_���С��.���С��)
                lng��� = lng��� + 2  '����������ʽΪ��2n-1;�������Ϊż��
                'zl_�����ƿ�_INSERT( /*NO_IN*/, /*���_IN*/, /*�ⷿID_IN*/,
                '/*�Է�����ID_IN*/, /*����ID_IN*/, /*����_IN*/, /*��д����_IN*/ʵ������/,
                '/*�ɱ���_IN*/, /*�ɱ����_IN*/, /*���ۼ�_IN*/, /*���۽��_IN*/,
                '/*���_IN*/, /*������_IN*/, /*����_IN*/, /*����_IN*/, /*Ч��_IN*/,
                '/*ժҪ_IN*/��������_in );
                gstrSQL = "zl_��������_INSERT('" & _
                    chrNo & "'," & _
                    lng��� & "," & _
                    lng�ⷿID & "," & _
                    lng����ID & "," & _
                    lng����ID & "," & _
                    lng���� & "," & _
                    dbl��д���� & "," & _
                    dblʵ������ & "," & _
                    dbl�ɱ��� & "," & _
                    dbl�ɱ���� & "," & _
                    dbl���ۼ� & "," & _
                    dbl���۽�� & "," & _
                    dbl��� & ",'" & _
                    str������ & "','" & _
                    str���� & "','" & _
                    str���� & "'," & _
                    IIf(strЧ�� = "", "Null", "to_date('" & strЧ�� & "','yyyy-mm-dd')") & "," & _
                    IIf(str���Ч�� = "", "Null", "to_date('" & str���Ч�� & "','yyyy-mm-dd')") & ",'" & _
                    strժҪ & "',to_date('" & _
                    str�������� & "','yyyy-mm-dd HH24:MI:SS')," & _
                    IIf(str�˲��� <> "", "'" & str�˲��� & "'", "Null") & "," & _
                    IIf(str�˲��� <> "", "to_date('" & str�˲����� & "','yyyy-mm-dd')", "null") & ")"
                    
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = CStr(lng����ID) & ";" & vbCrLf & gstrSQL
            End If
            recSort.MoveNext
        Next
        If Not ExecuteSql(arrSQL, mstrCaption) Then Exit Function
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub ��ʾ�ϼƽ��()
    Dim curTotal As Double, Cur���ʽ�� As Double, Cur���ʲ�� As Double
    Dim intLop As Integer
    
    curTotal = 0: Cur���ʽ�� = 0: Cur���ʲ�� = 0:
    
    With mshBill
        For intLop = 1 To .Rows - 1
            curTotal = curTotal + Val(.TextMatrix(intLop, mBillCol.C_�ɹ����))
            Cur���ʽ�� = Cur���ʽ�� + Val(.TextMatrix(intLop, mBillCol.C_�ۼ۽��))
        Next
    End With
    
    Cur���ʲ�� = Cur���ʽ�� - curTotal
    lblPurchasePrice.Caption = "�ɱ����ϼƣ�" & Format(curTotal, mFMT.FM_���)
    lblSalePrice.Caption = "�ۼ۽��ϼƣ�" & Format(Cur���ʽ��, mFMT.FM_���)
    lblDifference.Caption = "��ۺϼƣ�" & Format(Cur���ʲ��, mFMT.FM_���)
End Sub

Private Sub ��ʾ�����()
    Dim rsUseCount As New Recordset
    Dim dblStock As Double
    Dim blnIs��ʾ�Է���� As Boolean
    Dim str�Է������ As String
    
    On Error GoTo ErrHandle
    With mshBill
        If .TextMatrix(.Row, mBillCol.C_����) = "" Then
            stbThis.Panels(2).Text = ""
            Exit Sub
        End If
        If .TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
        
        
        '�������ĵ�ǰ���ĵĿ�������
        If mint��ȷ���� = 1 Then
            gstrSQL = " Select ��������/" & .TextMatrix(.Row, mBillCol.C_����ϵ��) & " as �������� from ҩƷ��� " & _
                      " Where �ⷿid=[1]" & _
                      " And ҩƷid=[2] And ����=1 " & _
                      " And Nvl(����,0)=[3]"
        Else
            gstrSQL = " Select Sum(��������)/" & .TextMatrix(.Row, mBillCol.C_����ϵ��) & " as �������� from ҩƷ��� " & _
                      " Where �ⷿid=[1]" & _
                      " And ҩƷid=[2] And ����=1 "
        End If
        
        Set rsUseCount = zlDatabase.OpenSQLRecord(gstrSQL, "�����ⷿ��������", cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mBillCol.c_����)))
        
        If rsUseCount.EOF Then
            .TextMatrix(.Row, mBillCol.C_��������) = 0
        Else
            .TextMatrix(.Row, mBillCol.C_��������) = IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0))
        End If
        rsUseCount.Close
        
        '��ǰ���ϲ��ŵĿ�������,����ⷿ����ʼ��Ϊ�ÿⷿ�������ο��
'        If mint��ȷ���� = 1 Then
'            gstrSQL = " Select Sum(��������/" & .TextMatrix(.Row, mBillCol.C_����ϵ��) & ") as �������� from ҩƷ��� where �ⷿid=[1]" & _
'                      " And ҩƷid=[2] And ����=1 " & _
'                      " And nvl(����,0)=[3]"
'        Else
            gstrSQL = " Select Sum(��������/" & .TextMatrix(.Row, mBillCol.C_����ϵ��) & ") as �������� from ҩƷ��� where �ⷿid=[1]" & _
                      " And ҩƷid=[2] And ����=1 "
'        End If
        
        Set rsUseCount = zlDatabase.OpenSQLRecord(gstrSQL, "��ǰ���õĿ�������", mlngStockID, Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mBillCol.c_����)))
        
        If rsUseCount.EOF Then
            dblStock = 0
        Else
            dblStock = IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0))
        End If
'        stbThis.Panels(2).Text = "�����ĵ�ǰ�����Ϊ[" & Format(dblStock, mFMT.FM_����) & "]" & .TextMatrix(.Row, mBillCol.C_��λ)
    
        
        blnIs��ʾ�Է���� = zlStr.IsHavePrivs(mstrPrivs, "��ʾ�Է����")
        str�Է������ = "��" & Me.cboStock.Text & "�����Ϊ[" & Format(.TextMatrix(.Row, mBillCol.C_��������), mFMT.FM_����) & "]" & .TextMatrix(.Row, mBillCol.c_��λ)
        
        stbThis.Panels(2).Text = "������" & mfrmMain.cboStock.Text & "�����Ϊ[" & Format(dblStock, mFMT.FM_����) & "]" & .TextMatrix(.Row, mBillCol.c_��λ) _
            & IIf(blnIs��ʾ�Է����, str�Է������, "")
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub get�������(Optional ByVal intRow As Integer = 0)
'''''''''''''''''''''''''''''''''''''
'��ȡ��������ķ���
'''''''''''''''''''''''''''''''''''''
    Dim rsUseCount As New Recordset
    Dim dblStock As Double
    Dim blnIs��ʾ�Է���� As Boolean
    Dim intStart As Integer, intEnd As Integer
    Dim i As Integer
    
    On Error GoTo ErrHandle
    
    blnIs��ʾ�Է���� = zlStr.IsHavePrivs(mstrPrivs, "��ʾ�Է����")
    
    If intRow > 0 Then
        intStart = intRow
        intEnd = intRow
    Else
        intStart = 1
        intEnd = mshBill.Rows - 1
    End If
    
    With mshBill
        For i = intStart To intEnd
            If .TextMatrix(i, 0) = "" Then Exit Sub

            If blnIs��ʾ�Է���� Then
                If Val(.TextMatrix(i, c_����)) > 0 Then
                    gstrSQL = " Select Nvl(��������,0)/" & .TextMatrix(i, C_����ϵ��) & " as ��������, Nvl(ʵ������,0)/" & .TextMatrix(i, C_����ϵ��) & " as ʵ������ from ҩƷ��� " & _
                              " Where �ⷿid=[1] " & _
                              " And ҩƷid=[2] And ����=1 " & _
                              " And Nvl(����,0)=[3] "
                Else
                    If Get��������(Val(cboStock.ItemData(cboStock.ListIndex)), Val(.TextMatrix(i, 0))) = 1 Then
                        '�������ⷿ�Ƿ�������ͳ���������εĺϼ�����
                        gstrSQL = " Select Sum(Nvl(��������,0))/" & .TextMatrix(i, C_����ϵ��) & " as ��������, Sum(Nvl(ʵ������,0))/" & .TextMatrix(i, C_����ϵ��) & " as ʵ������ from ҩƷ��� " & _
                              " Where �ⷿid=[1] " & _
                              " And ҩƷid=[2] And ����=1 And Nvl(����,0)>0 "
                    Else
                        '�������ⷿ�ǲ������ģ���ͳ���ܵ�����
                        gstrSQL = " Select Sum(Nvl(��������,0))/" & .TextMatrix(i, C_����ϵ��) & " as ��������, Sum(Nvl(ʵ������,0))/" & .TextMatrix(i, C_����ϵ��) & " as ʵ������ from ҩƷ��� " & _
                              " Where �ⷿid=[1] " & _
                              " And ҩƷid=[2] And ����=1 "
                    End If
                End If
                Set rsUseCount = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[�����ⷿ����]", cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(i, 0)), Val(.TextMatrix(i, c_����)))
                
                If rsUseCount.EOF Then
                    dblStock = 0
                Else
                    If mint�༭״̬ = 6 Then
                        '����(���)ʱ��ʾʵ������
                        dblStock = NVL(rsUseCount!ʵ������, 0)
                    Else
                        '����״̬ʱ��ʾ��������
                        dblStock = NVL(rsUseCount!��������, 0)
                    End If
                End If
                .TextMatrix(i, C_�Է����) = Format(dblStock, mFMT.FM_����)
                rsUseCount.Close
            End If
                
            '���ϲ���ʼ����ʾ��������
            gstrSQL = " Select Sum(Nvl(��������,0))/" & .TextMatrix(i, C_����ϵ��) & " as ��������, Sum(Nvl(ʵ������,0))/" & .TextMatrix(i, C_����ϵ��) & " as ʵ������ from ҩƷ��� where �ⷿid=[1] " & _
                      " And ҩƷid=[2] And ����=1 "
            Set rsUseCount = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[���첿������]", mlngStockID, Val(.TextMatrix(i, 0)), Val(.TextMatrix(i, c_����)))
            
            If rsUseCount.EOF Then
                dblStock = 0
            Else
                If mint�༭״̬ = 6 Then
                    '����(���)ʱ��ʾʵ������
                    dblStock = NVL(rsUseCount!ʵ������, 0)
                Else
                    '����״̬ʱ��ʾ��������
                    dblStock = NVL(rsUseCount!��������, 0)
                End If
            End If
            .TextMatrix(i, C_��ǰ���) = Format(dblStock, mFMT.FM_����)
       Next
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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

Private Sub txtժҪ_LostFocus()
    ImeLanguage False
End Sub

'ת����ֵΪ����
Private Function TranNumToDate(ByVal strNum As Long) As String
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    Dim strDate As String
    
    TranNumToDate = ""
    strYear = Mid(strNum, 1, 4)
    strMonth = Mid(strNum, 5, 2)
    strDay = Mid(strNum, 7, 2)
        
    If strYear < 2000 Or strYear > 5000 Then Exit Function
    If strMonth = "" Then strMonth = "01"
    If strDay = "" Then strDay = "01"
    
    If strMonth > 12 Or strMonth < 1 Then Exit Function
    strDate = strYear & "-" & strMonth & "-" & strDay
        
    If Not IsDate(strDate) Then Exit Function
    
    strDate = Format(strDate, "yyyy-mm-dd")
    TranNumToDate = strDate
    
End Function

'������������бȽ�
Private Function CompareUsableQuantity(ByVal intRow As Integer, ByVal dbl��д���� As Double, Optional ByVal blnSave As Boolean = False) As Boolean
    Dim dblUsableQuantity As Double      'ʵ��������Ӧ���������
    Dim numUsedCount As Double
    Dim varStuff As Variant
    Dim rsCheck As ADODB.Recordset
    Dim strSaveCheck As String
    
    'mint�����: 0-�����;1-��飬�������ѣ�2-��飬�����ֹ
    'ֻҪ�Ƿ������ģ���������ȵ�ǰ���δ�������������Զ��ֽ⣬��������ʱ���������ԵĲ�����
    CompareUsableQuantity = False
    If mint��ȷ���� = 0 Then CompareUsableQuantity = True: Exit Function
    
    With mshBill
        If .TextMatrix(intRow, 0) = "" Then Exit Function
        
        If Not blnSave Then
            dblUsableQuantity = Format(.TextMatrix(intRow, mBillCol.C_��������), mFMT.FM_����)
        Else
            '�������޸ı���ʱ����ȡ���ݿ��еĿ�����������Ҫ��ֹ���������ͬʱ��Կ�������ȡֵ��Ӱ��
            gstrSQL = "Select Nvl(��������, 0) �������� From ҩƷ��� Where ���� = 1 And �ⷿid = [1] And ҩƷid = [2] And Nvl(����, 0) = [3] "
            Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "CompareUsableQuantity", Val(cboStock.ItemData(cboStock.ListIndex)), Val(.TextMatrix(intRow, 0)), Val(.TextMatrix(intRow, mBillCol.c_����)))
            
            If rsCheck.EOF Then
                dblUsableQuantity = 0
            Else
                dblUsableQuantity = Val(Format(rsCheck!�������� / Val(.TextMatrix(intRow, mBillCol.C_����ϵ��)), mFMT.FM_����))
                
                If dblUsableQuantity <> Val(Format(.TextMatrix(intRow, mBillCol.C_��������), mFMT.FM_����)) Then
                    .TextMatrix(intRow, mBillCol.C_��������) = dblUsableQuantity
                End If
                
                strSaveCheck = "����������������Ա�ռ����"
            End If
        End If
        
        If mint����� = 0 Then
            '0-�����
        ElseIf mint����� = 1 Then
            '1-��飬��������
            If mint�༭״̬ = 1 Or mint�༭״̬ = 5 Then
                If dbl��д���� > dblUsableQuantity Then
                    If MsgBox("�������������" & dbl��д���� & "�������˸����ĵĿ��ÿ��������" & dblUsableQuantity & "��" & strSaveCheck & "���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
            ElseIf mint�༭״̬ = 2 Then
                numUsedCount = 0
                For Each varStuff In mcolUsedCount
                    If varStuff(0) = .TextMatrix(intRow, 0) & .TextMatrix(intRow, mBillCol.c_����) Then
                        numUsedCount = varStuff(1)
                        Exit For
                    End If
                Next
                
                If gSystem_Para.para_������¿��ÿ�� = False Then
                    '���û��Ԥ��������������������ԭʼ����
                    numUsedCount = 0
                End If
                
                If dbl��д���� > dblUsableQuantity + numUsedCount Then
                    If MsgBox("�������������" & dbl��д���� & "�������˸����ĵĿ��ÿ��������" & dblUsableQuantity + numUsedCount & "��" & strSaveCheck & "���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
            End If
            
        ElseIf mint����� = 2 Then
            '2-��飬�����ֹ
            If mint�༭״̬ = 1 Or mint�༭״̬ = 5 Then
                If dbl��д���� > dblUsableQuantity Then
                    MsgBox "�������������" & dbl��д���� & "�������˸����ĵĿ��ÿ��������" & dblUsableQuantity & "��" & strSaveCheck & "�������䣡", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            ElseIf mint�༭״̬ = 2 Then
                numUsedCount = 0
                For Each varStuff In mcolUsedCount
                    If varStuff(0) = .TextMatrix(intRow, 0) & .TextMatrix(intRow, mBillCol.c_����) Then
                        numUsedCount = varStuff(1)
                        Exit For
                    End If
                Next
                
                If gSystem_Para.para_������¿��ÿ�� = False Then
                    '���û��Ԥ��������������������ԭʼ����
                    numUsedCount = 0
                End If
                
                If dbl��д���� > dblUsableQuantity + numUsedCount Then
                    MsgBox "�������������" & dbl��д���� & "�������˸����ĵĿ��ÿ��������" & dblUsableQuantity + numUsedCount & "��" & strSaveCheck & "�������䣡", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
        End If
            
    End With
    
    CompareUsableQuantity = True
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Function ExecuteSql(ByRef arrSQL As Variant, strTitle As String, Optional ByVal blnǿ�Ʊ��� As Boolean = False) As Boolean
    Dim strTmp As Variant
    Dim i As Integer, j As Integer

    ExecuteSql = False
    If UBound(arrSQL) >= 0 Then
        '��SQL���в���ID��������
        For i = 0 To UBound(arrSQL) - 1
            For j = i + 1 To UBound(arrSQL)
                If CLng(Split(arrSQL(j), ";")(0)) < CLng(Split(arrSQL(i), ";")(0)) Then
                    strTmp = CStr(arrSQL(j))
                    arrSQL(j) = arrSQL(i)
                    arrSQL(i) = strTmp
                End If
            Next
        Next
        
        'ִ��SQL���
        On Error GoTo errH
        If Not blnǿ�Ʊ��� Then gcnOracle.BeginTrans
        For i = 0 To UBound(arrSQL)
            zlDatabase.ExecuteProcedure CStr(Split(arrSQL(i), ";")(1)), mstrCaption
                        
'            Call SQLTest(App.ProductName, strTitle, CStr(Split(arrSql(i), ";")(1)))
'            Debug.Print CStr(Split(arrSql(i), ";")(1))
'            gcnOracle.Execute CStr(Split(arrSql(i), ";")(1)), , adCmdStoredProc
'            Call SQLTest
        Next
        If Not blnǿ�Ʊ��� Then gcnOracle.CommitTrans
        ExecuteSql = True
    End If
    Exit Function
errH:
    If Not blnǿ�Ʊ��� Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


'��ӡ����
Private Sub printbill()
    
    Dim strNo As String
    strNo = txtNO.Tag
    FrmBillPrint.ShowMe Me, glngSys, "zl1_bill_1722", mint��¼״̬, mintUnit, 1722, "�������쵥", strNo
End Sub


Private Function SaveCheck() As Boolean
    Dim rsTemp As New Recordset
    Dim intRow As Integer
    
    Dim strNo As String
    Dim lng�ⷿID As Long
    Dim lng�Է�����id As Long
    Dim str����� As String
    
    Dim lng����ID As Long
    Dim str���� As String
    Dim lng������ As Long
    Dim dbl��д���� As Double
    Dim dblʵ������ As Double
    Dim dbl�ɱ��� As Double
    Dim dbl�ɱ���� As Double
    Dim dbl�ۼ� As Double
    Dim dbl���۽�� As Double
    Dim dbl��� As Double
    Dim lng�����id As Long
    Dim lng�����id As Long
    Dim str���� As String
    Dim strЧ�� As String
    Dim str���Ч�� As String
    Dim str������� As String
    Dim int���к� As Integer
    Dim n As Long
    
    Dim arrSQL As Variant
    
    On Error GoTo ErrHandle
    arrSQL = Array()
    mblnSave = False
    SaveCheck = False
    
    '���õ����Ƿ��ڽ���༭����󣬱���������Ա�޸�
    mstrTime_End = GetBillInfo(19, mstr���ݺ�)
    If mstrTime_End = "" Then
        MsgBox "�õ����Ѿ�����������Աɾ����", vbInformation, gstrSysName
        Exit Function
    End If
    If mstrTime_End > mstrTime_Start Then
        MsgBox "�õ����Ѿ�����������Ա�༭�����˳������ԣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    
    '���õ����Ƿ���������
    gstrSQL = " Select ��ҩ���� From ҩƷ�շ���¼ " & _
            " Where ����=19 And NO=[1] And Rownum<2"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���õ����Ƿ���������", Me.txtNO.Tag)
    
    If IsNull(rsTemp!��ҩ����) Then
        MsgBox "�õ��ݱ���������Աȡ�����ͣ���������գ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    lng�ⷿID = cboStock.ItemData(cboStock.ListIndex)
    lng�Է�����id = mlngStockID
    str����� = gstrUserName
    strNo = txtNO.Tag
    
    
    gstrSQL = "" & _
        "   SELECT b.ϵ��,b.id AS ���id " & _
        "   FROM ҩƷ�������� a, ҩƷ������ b " & _
        "   Where a.���id = b.ID AND a.���� = 34 "
    
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "�����ƿ����")
    
    If rsTemp.EOF Then
        MsgBox "����������಻ȫ������!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If rsTemp.RecordCount < 2 Then
        MsgBox "����������಻ȫ������!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        If rsTemp!ϵ�� = 1 Then
            lng�����id = rsTemp!���ID
        Else
            lng�����id = rsTemp!���ID
        End If
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    str������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
    With mshBill
        On Error GoTo ErrHandle
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng����ID = .TextMatrix(intRow, 0)
                str���� = .TextMatrix(intRow, mBillCol.C_����)
                lng������ = .TextMatrix(intRow, mBillCol.c_����)
                dbl��д���� = Round(Val(.TextMatrix(intRow, mBillCol.C_��д����)) * .TextMatrix(intRow, mBillCol.C_����ϵ��), g_С��λ��.obj_ɢװС��.����С��)
                dblʵ������ = Round(Val(.TextMatrix(intRow, mBillCol.C_ʵ������)) * .TextMatrix(intRow, mBillCol.C_����ϵ��), g_С��λ��.obj_ɢװС��.����С��)
                dbl�ɱ��� = Round(Val(.TextMatrix(intRow, mBillCol.C_�ɹ���)) / .TextMatrix(intRow, mBillCol.C_����ϵ��), g_С��λ��.obj_ɢװС��.�ɱ���С��)
                dbl�ɱ���� = Round(Val(.TextMatrix(intRow, mBillCol.C_�ɹ����)), g_С��λ��.obj_ɢװС��.���С��)
                dbl�ۼ� = Round(Val(.TextMatrix(intRow, mBillCol.C_�ۼ�)) / .TextMatrix(intRow, mBillCol.C_����ϵ��), g_С��λ��.obj_ɢװС��.���ۼ�С��)
                dbl���۽�� = Round(Val(.TextMatrix(intRow, mBillCol.C_�ۼ۽��)), g_С��λ��.obj_ɢװС��.���С��)
                dbl��� = Round(Val(.TextMatrix(intRow, mBillCol.C_���)), g_С��λ��.obj_ɢװС��.���С��)
                str���� = .TextMatrix(intRow, mBillCol.c_����)
                strЧ�� = IIf(.TextMatrix(intRow, mBillCol.C_Ч��) = "", "Null", "to_date('" & .TextMatrix(intRow, mBillCol.C_Ч��) & "','yyyy-mm-dd')")
                str���Ч�� = IIf(.TextMatrix(intRow, mBillCol.C_���ʧЧ��) = "", "Null", "to_date('" & .TextMatrix(intRow, mBillCol.C_���ʧЧ��) & "','yyyy-mm-dd')")

                int���к� = Val(.TextMatrix(intRow, mBillCol.C_���))
                
                'zl_�����ƿ�_VERIFY( /*�ⷿID_IN*/, /*�Է�����ID_IN*/, /*ҩƷID_IN*/,
                    '����_IN*/, /*������_IN*/, /*��д����_IN*/, /*ʵ������_IN*/, /*�ɱ���_IN*/,
                    '/*�ɱ����_IN*/, /*���۽��_IN*/, /*���_IN*/, /*�����ID_IN*/, /*�����ID_IN*/,
                    '/*NO_IN*/, /*�����_IN*/, /*����_IN*/, /*Ч��_IN*/���Ч��_IN );
                        
                gstrSQL = "zl_�����ƿ�_Verify(" & int���к� & "," & lng�ⷿID & "," & lng�Է�����id & "," & _
                     lng����ID & ",'" & str���� & "'," & lng������ & "," & dbl��д���� & "," & _
                     dblʵ������ & "," & dbl�ɱ��� & "," & dbl�ɱ���� & "," & dbl���۽�� & "," & _
                     dbl��� & "," & lng�����id & "," & lng�����id & ",'" & _
                     strNo & "','" & str����� & "','" & str���� & "'," & strЧ�� & "," & str���Ч�� & ",to_date('" & str������� & "','yyyy-mm-dd HH24:MI:SS')" & _
                    ",1," & dbl�ۼ� & " )"
                    
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = CStr(lng����ID) & ";" & gstrSQL
            End If
            recSort.MoveNext
        Next
    End With
    
    gcnOracle.BeginTrans
    If Not ExecuteSql(arrSQL, mstrCaption, True) Then
        gcnOracle.RollbackTrans
        Exit Function
    End If
'    If Not ��鵥��(19, txtNo.Tag) Then
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
    gcnOracle.RollbackTrans
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
    Dim rsTemp As New ADODB.Recordset
    Dim n As Long
    
    SaveStrike = False
    
    With mshBill
        '����������������С����
        For intRow = 1 To .Rows - 1
            If Val(.TextMatrix(intRow, mBillCol.C_ʵ������)) <> 0 Then
                If Not ��ͬ����(Val(.TextMatrix(intRow, mBillCol.C_��д����)), Val(.TextMatrix(intRow, mBillCol.C_ʵ������))) Then
                    MsgBox "������Ϸ��ĳ�����������" & intRow & "�У���", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Next
        
        NO_IN = Trim(txtNO.Tag)
        ������_IN = gstrUserName
        ��������_IN = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        ԭ��¼״̬_IN = mint��¼״̬
        
        err = 0: On Error GoTo ErrHandle
        
        gcnOracle.BeginTrans
        
        �д�_IN = 0
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" And Val(.TextMatrix(intRow, mBillCol.C_ʵ������)) <> 0 Then
                �д�_IN = �д�_IN + 1
                
                ����ID_IN = .TextMatrix(intRow, 0)
                ��������_IN = Round(Val(.TextMatrix(intRow, mBillCol.C_ʵ������)) * Val(.TextMatrix(intRow, mBillCol.C_����ϵ��)), g_С��λ��.obj_ɢװС��.����С��)
                If Val(.TextMatrix(intRow, mBillCol.C_ʵ������)) = Val(.TextMatrix(intRow, mBillCol.C_��д����)) Then
                    ��������_IN = Val(.TextMatrix(intRow, mBillCol.c_ԭʼ����))
                End If
                
                ���_IN = .TextMatrix(intRow, mBillCol.C_���)
                
                'ZL_�����ƿ�_STRIKE(/*�д�_IN*/,/*ԭ��¼״̬_IN*/,/*NO_IN*/,/*���_IN*/, /*����ID_IN*/,
                '/*��������_IN*/,/*������_IN*/, /*��������_IN*/);
                gstrSQL = "" & _
                    "   ZL_�����ƿ�_STRIKE(" & _
                            �д�_IN & "," & _
                            ԭ��¼״̬_IN & ",'" & _
                            NO_IN & "'," & _
                            ���_IN & "," & _
                            ����ID_IN & "," & _
                            ��������_IN & ",'" & _
                            ������_IN & "',to_date('" & _
                            Format(��������_IN, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS')," & _
                            mint����ʽ & ")"
                zlDatabase.ExecuteProcedure gstrSQL, mstrCaption
                
            End If
            recSort.MoveNext
        Next
        gcnOracle.CommitTrans
        
        If �д�_IN = 0 Then
            MsgBox "û��ѡ��һ�в��������������ܳ��������飡", vbInformation + vbOKOnly, gstrSysName
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
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


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
                !��� = IIf(Val(mshBill.TextMatrix(n, mBillCol.C_���)) = 0, n, Val(mshBill.TextMatrix(n, mBillCol.C_���)))
                !ҩƷid = Val(mshBill.TextMatrix(n, 0))
                !���� = Val(mshBill.TextMatrix(n, mBillCol.c_����))
                
                .Update
            End If
        Next
        
    End With
End Sub

