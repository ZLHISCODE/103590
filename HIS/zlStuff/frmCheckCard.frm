VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmCheckCard 
   Caption         =   "�����̵��"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmCheckCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11400
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmd�̶��� 
      Caption         =   "�̶���(&L)"
      Height          =   350
      Left            =   6090
      TabIndex        =   27
      Top             =   5040
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   8
      Top             =   5137
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   7
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   6
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7425
      TabIndex        =   4
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8730
      TabIndex        =   5
      Top             =   5040
      Width           =   1100
   End
   Begin VB.PictureBox Pic���� 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   9
      Top             =   0
      Width           =   11715
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   9960
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   165
         Width           =   1425
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   210
         TabIndex        =   1
         Top             =   945
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
         TabIndex        =   3
         Top             =   4080
         Width           =   10410
      End
      Begin VB.Label lblCheckCostSum 
         AutoSize        =   -1  'True
         Caption         =   "�̵�ɱ����ϼƣ�"
         Height          =   180
         Left            =   3960
         TabIndex        =   29
         Top             =   3840
         Width           =   1620
      End
      Begin VB.Label lblCheckSum 
         AutoSize        =   -1  'True
         Caption         =   "�̵���ϼƣ�"
         Height          =   180
         Left            =   1920
         TabIndex        =   26
         Top             =   3840
         Width           =   1260
      End
      Begin VB.Label lblCheckDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�̵�ʱ��"
         Height          =   180
         Left            =   8640
         TabIndex        =   24
         Top             =   660
         Width           =   720
      End
      Begin VB.Label txtCheckDate 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   9600
         TabIndex        =   23
         Top             =   600
         Width           =   1875
      End
      Begin VB.Label txtStock 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1080
         TabIndex        =   22
         Top             =   600
         Width           =   1845
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "����ϼƣ�"
         Height          =   180
         Left            =   240
         TabIndex        =   21
         Top             =   3840
         Width           =   1080
      End
      Begin VB.Label Txt����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7950
         TabIndex        =   19
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Txt������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10050
         TabIndex        =   18
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt�������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   17
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt������ 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   2
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "���������̵��"
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
         TabIndex        =   14
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�̵�ⷿ"
         Height          =   180
         Left            =   270
         TabIndex        =   0
         Top             =   660
         Width           =   720
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   300
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
            Picture         =   "frmCheckCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":1000
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
            Picture         =   "frmCheckCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   25
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
            Picture         =   "frmCheckCard.frx":22EA
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
            Picture         =   "frmCheckCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmCheckCard.frx":3080
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
      TabIndex        =   20
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Menu PopMenu 
      Caption         =   "�̶���"
      Visible         =   0   'False
      Begin VB.Menu mnuFirst 
         Caption         =   "�Ӳ�����Ϣ����λ��(&1)"
      End
      Begin VB.Menu mnuSecond 
         Caption         =   "�Ӳ�����Ϣ��Ч����(&2)"
      End
      Begin VB.Menu mnuSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDefault 
         Caption         =   "�ָ�(&D)"
      End
   End
End
Attribute VB_Name = "frmCheckCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mint�༭״̬ As Integer             '1.������2���޸ģ�3�����գ�4���鿴��5�������̵��¼��,�����̵��;6��ȫ����Ϊ��
Private mstr���ݺ� As String                '����ĵ��ݺ�;
Private mint��¼״̬ As Integer             '1:������¼;2-������¼;3-�Ѿ�������ԭ��¼
Private mblnSuccess As Boolean              'ֻҪ��һ�ųɹ�����ΪTrue������ΪFalse
Private mblnFirst As Boolean                '��һ����ʾ
Private mblnSave As Boolean                 '�Ƿ���̺����   TURE���ɹ���
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '�Ƿ�����޸�
Private mblnChange As Boolean               '�Ƿ���й��༭
Private mintParallelRecord As Integer       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼
Private mintBatchNoLen As Integer           '���ݿ������Ŷ��峤��

Private mint����� As Integer             '��ʾ�������ϳ���ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ
Dim mstrPrivs As String                     'Ȩ��
Private Const mstrCaption As String = "�����̵��"
Private mstr�ظ����� As String '��¼�ظ�������

Private recSort As ADODB.Recordset          '��ҩƷID�����������ר�ü�¼��

'���˺�:2007/06/10
Private mstrTime_Start As String            '���뵥�ݱ༭�ĵ���ʱ�� ,��Ҫ�ж��Ƿ񵥾ݱ����˸��Ĺ�,����༭��,���ܽ������
Private mstrTime_End As String
Private Const mlngModule = 1719
Private mblnCostView As Boolean                 '�鿴�ɱ��� true-����鿴 false-������鿴

'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------

Private mbln��������    As Boolean          '����ʱ���ݺ��ۼ�1
Private mintUnit  As Integer                '��ʾ��λ:0-ɢװ��λ,1-��װ��λ
Private mstr�̵㵥�� As String  '��NO,NOΪ�ָ�
Private mblnֻͳ���̵㵥���� As Boolean
Private mblnɾ���̵㵥 As Boolean
Private mbln���޴洢�ⷿ���� As Boolean
Private mbln�����������Ų��ؿ��� As Boolean  '�Ƿ�������������Ų����Ƿ�¼��

'=========================================================================================
Private Enum mBillCol
     C_�к� = 1
     C_���� = 2
     C_��� = 3
     c_��� = 4
     C_���� = 5
     C_�������� = 6
     c_����ϵ�� = 7
     C_ָ������� = 8
     C_ʵ�ʲ�� = 9
     C_ʵ�ʽ�� = 10
     C_���� = 11
     C_��׼�ĺ� = 12
     C_�ⷿ��λ = 13
     c_��λ = 14
     c_���� = 15
     C_Ч�� = 16
     C_�������� = 17
     C_ʵ������ = 18
     C_��־ = 19
     C_������ = 20
     C_�ɱ��� = 21
     C_�ۼ� = 22
     c_���� = 23
     c_��۲� = 24
     C_�̵��� = 25
     C_�̵�ɱ���� = 26
     C_�̵�ɱ����� = 27
     c_������ = 28
     c_���ű༭ = 29
     c_���ر༭ = 30
     C_Cols = 31               '������
End Enum

'=========================================================================================


'�������������
Private Function GetDepend() As Boolean
    Dim rsTemp As New Recordset
    
    On Error GoTo errHandle
    GetDepend = False
    
    gstrSQL = "" & _
        "   SELECT B.Id " & _
        "   FROM ҩƷ�������� A, ҩƷ������ B " & _
        "   Where A.���id = B.ID " & _
        "           AND A.���� = [1]  and b.ϵ��=[2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���������̵����", 37, 1)
    
    If rsTemp.EOF Then
        ShowMsgBox "û���������������̵�������������������������ã�"
        rsTemp.Close
        Exit Function
    End If
    rsTemp.Close
    
    
    gstrSQL = "" & _
        "   SELECT B.Id " & _
        "   FROM ҩƷ�������� A, ҩƷ������ B " & _
        "   Where A.���id = B.ID " & _
        "           AND A.���� = [1]  and b.ϵ��=[2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���������̵����", 37, -1)

    If rsTemp.EOF Then
        ShowMsgBox "û���������������̵��ĳ����������������������ã�"
        rsTemp.Close
        Exit Function
    End If
    GetDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Sub ShowCard(FrmMain As Form, ByVal str���ݺ� As String, ByVal int�༭״̬ As Integer, Optional int��¼״̬ As Integer = 1, _
    Optional strPrivs As String, Optional blnSuccess As Boolean = False)
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�༭���ݻ���ʾ����,�ǵ��ݵ�Ψһ���
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim strReg As String
    
    mblnSave = False
    mblnSuccess = False
    mstr���ݺ� = str���ݺ�
    mint�༭״̬ = int�༭״̬
    mint��¼״̬ = int��¼״̬
    mblnSuccess = blnSuccess
    mblnChange = False
    mblnFirst = True
    mintParallelRecord = 1
    mstrPrivs = strPrivs
    
    Set mfrmMain = FrmMain
    If Not GetDepend Then Exit Sub
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    
    Call GetRegInFor(g˽��ģ��, "�����̵����", "���ݺ��ۼ�", strReg)
    mbln�������� = IIf(strReg = "", True, Val(strReg) = 1)
    
    If mint�༭״̬ = 1 Or mint�༭״̬ = 5 Or mint�༭״̬ = 6 Then
        mblnEdit = True
        If mbln�������� Then
            'mstr���ݺ� = NextNo(75)
        End If

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
        If mint�༭״̬ = 4 Then
        If InStr(mstrPrivs, "���ݴ�ӡ") = 0 Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
    End If
    End If
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    
    Me.Show vbModal, FrmMain
    blnSuccess = mblnSuccess
    str���ݺ� = mstr���ݺ�
    
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
        FindRownew mshBill, mBillCol.C_����, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub cmd�̶���_Click()
    Call PopupMenu(PopMenu, 2)
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


Private Sub CmdSave_Click()
    Dim blnSuccess As Boolean
    Dim strReg As String
    
    '�����������ݼ�
    Call SetSortRecord
    
    If mint�༭״̬ = 5 Then    '���ܲ����̵��
        If ValidData = False Then Exit Sub
        blnSuccess = SaveCard
        
        If blnSuccess Then
            '�����������
'            If SaveCheck Then
'                strReg = IIf(Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModule, "0")) = 1, 1, 0)
'                If Val(strReg) = 1 Then
'                    '��ӡ
'                    If InStr(mstrPrivs, "���ݴ�ӡ") <> 0 Then
'                        printbill
'                    End If
'                End If
'            End If
            strReg = IIf(Val(zlDatabase.GetPara("���̴�ӡ", glngSys, mlngModule, "0")) = 1, 1, 0)
            If Val(strReg) = 1 Then
                '��ӡ
                If InStr(mstrPrivs, "���ݴ�ӡ") <> 0 Then
                    printbill
                End If
            End If
        End If
        
        Unload Me
        Exit Sub
    End If
    If mint�༭״̬ = 4 Then    '�鿴
        '��ӡ
        printbill
        '�˳�
        Unload Me
        Exit Sub
    End If
    
    If mint�༭״̬ = 3 Then        '���
        
        mstrTime_End = GetBillInfo(22, mstr���ݺ�)
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
        
        If Not ���ϵ������(Txt������.Caption) Then Exit Sub
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
'
'    If mbln�������� Then
'        'mstr���ݺ� = NextNo(75)
'        txtNO = mstr���ݺ�
'    End If
    mblnSave = False
    mblnEdit = True
    mshBill.ClearBill
    Call RefreshRowNO(mshBill, mBillCol.C_�к�, 1)
    txtժҪ.Text = ""
    mblnChange = False
    If txtNO.Tag <> "" Then Me.stbThis.Panels(2).Text = "��һ�ŵ��ݵ�NO�ţ�" & txtNO.Tag
End Sub

Private Sub Form_Activate()
    
    Dim str����ID As String, lng�ⷿid As Long, int�̵㷽ʽ As Integer, str�̵�ʱ�� As String, str�ⷿ��λ As String
    Dim int���޿����� As Integer, bln�̵����������н�� As Boolean
    
    If mblnFirst = False Then Exit Sub
        
    mint����� = Get������(lng�ⷿid)
    mintBatchNoLen = GetBatchNoLen()
    
    If mintParallelRecord <> 1 Then mblnChange = False
    
    Select Case mintParallelRecord
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
        Case 5
            MsgBox "������δ��˵��������ϵ��ݣ���ȫ����˺����ԣ�", vbOKOnly, gstrSysName
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
    
    mblnFirst = False
    '��ʼ������
    str����ID = ""
    
    If mint�༭״̬ = 1 Then
        '�Զ��������ֹ������̵��
        mshBill.ClearBill
        Call RefreshRowNO(mshBill, mBillCol.C_�к�, 1)
        
        If frmCheckCondition.GetCondition(mfrmMain, str����ID, lng�ⷿid, int�̵㷽ʽ, str�̵�ʱ��, int���޿�����, bln�̵����������н��, str�ⷿ��λ) = True Then
            If str����ID <> "" Then
                If str����ID = "������������" Then
                    str����ID = ""
                End If
                Call SearchData(str����ID, lng�ⷿid, int�̵㷽ʽ, str�̵�ʱ��, int���޿�����, bln�̵����������н��, str�ⷿ��λ)
            End If
        Else
            Unload Me
            Exit Sub
        End If
        
        If CmdCancel.Enabled = False Then
            CmdCancel.Enabled = True
        End If
        If CmdSave.Enabled = False Then
            CmdSave.Enabled = True
        End If
        
        If mshBill.Visible = True Then
            mshBill.SetFocus
        End If
    ElseIf mint�༭״̬ = 5 Then
        '�����̵������ָ��ʱ�̵��̵��¼����ָ��ʱ�̵Ŀ�棩
        mshBill.ClearBill
        Call RefreshRowNO(mshBill, mBillCol.C_�к�, 1)
        
        If FrmCheckCourseCondition.GetCondition(mfrmMain, lng�ⷿid, str�̵�ʱ��, mstr�̵㵥��, mblnֻͳ���̵㵥����, mblnɾ���̵㵥) = True Then
            Call SearchTableData(lng�ⷿid, str�̵�ʱ��)
        
        Else
            Unload Me
            Exit Sub
        End If
        If CmdCancel.Enabled = False Then
            CmdCancel.Enabled = True
        End If
        If CmdSave.Enabled = False Then
            CmdSave.Enabled = True
        End If
        
        If mshBill.Visible = True Then
            mshBill.SetFocus
        End If
    ElseIf mint�༭״̬ = 6 Then
        'ȫ����Ϊ��
        str�̵�ʱ�� = Format(sys.Currentdate, "yyyy-MM-dd HH:mm:ss")
        txtCheckDate = str�̵�ʱ��
        txtStock.Caption = mfrmMain.cboStock.Text
        lng�ⷿid = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
        txtStock.Tag = lng�ⷿid
        
        mshBill.ClearBill
        Call RefreshRowNO(mshBill, mBillCol.C_�к�, 1)
        
        Call SearchTableData(lng�ⷿid, str�̵�ʱ��)
        If CmdCancel.Enabled = False Then
            CmdCancel.Enabled = True
        End If
        If CmdSave.Enabled = False Then
            CmdSave.Enabled = True
        End If
        If mshBill.Visible = True Then
            mshBill.SetFocus
        End If
    End If
End Sub

Private Sub SearchData(ByVal str����ID As String, ByVal lng�ⷿid As Long, _
    ByVal int�̵㷽ʽ As Integer, ByVal str�̵�ʱ�� As String, ByVal int���޿����� As Integer, _
    ByVal bln�̵����������н�� As Boolean, ByVal str�ⷿ��λ As String)
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:������������ȡ�������
    '--�����:str����ID-����ID(1,2)
    '         lng�ⷿID-�ⷿid
    '         int�̵㷽ʽ:����,����...
    '         str�̵�ʱ��-�̵�����
    '         int���޿�����-�����̵��޿�������Ĳ���
    '         bln�̵����������н��-�����̵��޿���������н�����������
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------

    
    Dim rsData As ADODB.Recordset '����¼��
    Dim rsTemp As ADODB.Recordset
    
    Dim strPhysic As String, i As Long
    Dim sngLevel As Single
    Dim lngRecordCount As Long
    Dim dbl�ɱ��� As Double
    Dim bln�ⷿ As Boolean
    Dim rsprice As New Recordset
    Dim strMoneyDigit As String
    Dim dbl����, dbl��۲� As Double
    
'    On Error Resume Next
    On Error GoTo errHandle
    '���ý�����ʾ����
    Select Case int�̵㷽ʽ
        Case 1
            stbThis.Panels(2).Text = "���ڶ�" & txtStock & "���������Ͻ������̵�"
        Case 2
            stbThis.Panels(2).Text = "���ڶ�" & txtStock & "���������Ͻ������̵�"
        Case 3
            stbThis.Panels(2).Text = "���ڶ�" & txtStock & "���������Ͻ������̵�"
        Case 4
            stbThis.Panels(2).Text = "���ڶ�" & txtStock & "���������Ͻ��м����̵�"
        Case 5
            stbThis.Panels(2).Text = "���ڶ�" & txtStock & "���������Ͻ��к����̵㷽ʽ�̵�"
    End Select
    
  
    Call FS.ShowFlash("���ڼ����������Ͽ������,���Ժ� ...", Me)

    DoEvents    ': Me.Refresh
    Set rsData = GetDateStock(str�̵�ʱ��, lng�ⷿid, int�̵㷽ʽ, IIf(int���޿����� = 0, False, True), , str����ID, , bln�̵����������н��, str�ⷿ��λ)
    
    Call FS.StopFlash    ': Me.Refresh
    
    lngRecordCount = rsData.RecordCount
    If lngRecordCount = 0 Then
        If mint�༭״̬ = 6 Then
            ShowMsgBox "δ����ȷ��ȡ�������Ͽ������,�����ԣ�": Exit Sub
        Else
            ShowMsgBox "δ����ȷ��ȡ�������Ͽ������,�����Ի��ֹ������������ϣ�": Exit Sub
        End If
    End If
    
    Call FS.ShowFlash("����װ��������������,���Ժ� ...", Me)
    DoEvents: 'Me.Refresh
    mshBill.Redraw = False
    
    rsData.MoveFirst
    i = 1
    bln�ⷿ = CheckPartProp(lng�ⷿid)
    
    With mshBill
        Do While Not rsData.EOF
            If i > 1 Then .Rows = .Rows + 1
            .TextMatrix(i, 0) = rsData!����ID
            
            'ȡ�ò��ϵĳɱ��ۣ����������Ϊ������ʱ�۲���ʱ�����ڼ����ۣ�
'            gstrSQL = "Select Nvl(�ɱ���,0) �ɱ��� From �������� Where ����ID=[1]"
'            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "--ȡ���������ϵĳɱ���", Val(NVL(rsData!����ID)))
'
            dbl�ɱ��� = Val(zlStr.NVL(rsData!������)) ' rsTemp!�ɱ���
            
            'ʱ�۲��������ۼ�
            If rsData!�Ƿ��� = 1 Then
                .TextMatrix(i, mBillCol.C_�ۼ�) = Format(Get���ۼ�(Val(zlStr.NVL(rsData!����ID)), Val(txtStock.Tag), Val(zlStr.NVL(rsData!����)), rsData!����ϵ��), mFMT.FM_���ۼ�)
            Else
                .TextMatrix(i, mBillCol.C_�ۼ�) = Format(IIf(IsNull(rsData!�ۼ�), 0, rsData!�ۼ�), mFMT.FM_���ۼ�)
            End If
           
            .TextMatrix(i, mBillCol.C_����) = "[" & rsData!���� & "]" & rsData!��Ʒ����
            .TextMatrix(i, mBillCol.c_���) = IIf(IsNull(rsData!���), "", rsData!���)
            .TextMatrix(i, mBillCol.C_����) = IIf(IsNull(rsData!����), "", rsData!����)
            .TextMatrix(i, mBillCol.C_��׼�ĺ�) = IIf(IsNull(rsData!��׼�ĺ�), "", rsData!��׼�ĺ�)
            .TextMatrix(i, mBillCol.C_�ⷿ��λ) = IIf(IsNull(rsData!�ⷿ��λ), "", rsData!�ⷿ��λ)
            .TextMatrix(i, mBillCol.c_��λ) = IIf(IsNull(rsData!��λ), "", rsData!��λ)
            .TextMatrix(i, mBillCol.c_����) = IIf(IsNull(rsData!����), "", rsData!����)
            
            '����Ƿ������ϣ������θ���Ϊ-1����ʾ��������
            .TextMatrix(i, mBillCol.C_����) = IIf(IsNull(rsData!����), "", rsData!����)
            
            If Val(.TextMatrix(i, mBillCol.C_����)) <> 0 Then
                .TextMatrix(i, mBillCol.c_���ű༭) = rsData!���ű༭
                .TextMatrix(i, mBillCol.c_���ر༭) = rsData!���ر༭
            End If
            
            If CheckPhysicBatch(bln�ⷿ, rsData!�ⷿ����, rsData!���÷���) And Val(.TextMatrix(i, mBillCol.C_����)) = 0 Then
                .TextMatrix(i, mBillCol.C_����) = -1
            End If
            If Val(.TextMatrix(i, mBillCol.C_����)) = -1 Then
                .TextMatrix(i, mBillCol.C_�ɱ���) = Format(dbl�ɱ���, mFMT.FM_�ɱ���)
            Else
                .TextMatrix(i, mBillCol.C_�ɱ���) = Format(Val(zlStr.NVL(rsData!�ɱ���)), mFMT.FM_�ɱ���)
            End If
            .TextMatrix(i, mBillCol.C_Ч��) = IIf(IsNull(rsData!Ч��), "", Format(rsData!Ч��, "yyyy-MM-dd"))
            .TextMatrix(i, mBillCol.C_��������) = Format(Val(zlStr.NVL(rsData!��������)), mFMT.FM_����)
            .TextMatrix(i, mBillCol.C_ʵ������) = .TextMatrix(i, mBillCol.C_��������)
            .TextMatrix(i, mBillCol.C_�̵���) = Format(Val(.TextMatrix(i, mBillCol.C_ʵ������)) * Val(.TextMatrix(i, mBillCol.C_�ۼ�)), mFMT.FM_���)
            .TextMatrix(i, mBillCol.C_��������) = rsData!��������
            .TextMatrix(i, mBillCol.C_ʵ�ʽ��) = rsData!ʵ�ʽ��
            .TextMatrix(i, mBillCol.C_ʵ�ʲ��) = rsData!ʵ�ʲ��
            .TextMatrix(i, mBillCol.c_����ϵ��) = rsData!����ϵ��
            
            .TextMatrix(i, mBillCol.C_ָ�������) = rsData!ָ������� & "||" & rsData!�Ƿ��� & "||" & rsData!���÷���
            .TextMatrix(i, mBillCol.C_��־) = "ƽ"
            .TextMatrix(i, mBillCol.C_������) = Format("0", mFMT.FM_����)
            
            If Val(.TextMatrix(i, mBillCol.C_��������)) = 0 Then
                strMoneyDigit = "#0.00000"
            Else
                strMoneyDigit = mFMT.FM_���
            End If
             
             '����=��ǰ�ۼ�*ʵ������-ʵ�ʽ��
             '��۲�=����*iif(ʵ�ʽ��<=0,ָ�������,(ʵ�ʲ��/ʵ�ʽ��))
            .TextMatrix(i, mBillCol.c_����) = Format(Val(.TextMatrix(i, mBillCol.C_�ۼ�)) * Val(.TextMatrix(i, mBillCol.C_ʵ������)) - Val(.TextMatrix(i, mBillCol.C_ʵ�ʽ��)), strMoneyDigit)
            .TextMatrix(i, mBillCol.c_��۲�) = Format((Val(.TextMatrix(i, mBillCol.C_�ۼ�)) - Val(.TextMatrix(i, mBillCol.C_�ɱ���))) * Val(.TextMatrix(i, mBillCol.C_ʵ������)) - Val(.TextMatrix(i, mBillCol.C_ʵ�ʲ��)), strMoneyDigit)
            dbl���� = Val(.TextMatrix(i, mBillCol.c_����))
            dbl��۲� = Val(.TextMatrix(i, mBillCol.c_��۲�))
            
            .TextMatrix(i, mBillCol.C_�̵�ɱ����) = Format(Val(.TextMatrix(i, mBillCol.C_ʵ�ʽ��)) + dbl���� - (Val(.TextMatrix(i, mBillCol.C_ʵ�ʲ��)) + dbl��۲�), mFMT.FM_���)
            .TextMatrix(i, mBillCol.C_�̵�ɱ�����) = Format(Val(.TextMatrix(i, mBillCol.c_����)) - Val(.TextMatrix(i, mBillCol.c_��۲�)), mFMT.FM_���)
            Call ShowPercent(i / lngRecordCount)
            i = i + 1
nextloop:
            rsData.MoveNext
        Loop
        Call RefreshRowNO(mshBill, mBillCol.C_�к�, 1)
        .Redraw = True
    End With
    Call FS.StopFlash
    stbThis.Panels(2).Text = ""
    mshBill.Row = 1: mshBill.Col = mBillCol.C_ʵ������
    If Me.Visible = True Then
        mshBill.SetFocus
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SearchTableData(ByVal lng�ⷿid As Long, ByVal str�̵�ʱ�� As String)
    Dim rsData As ADODB.Recordset '�������Ͽ���¼��
    Dim rsTemp As ADODB.Recordset
    Dim strPhysic As String, i As Long
    Dim sngLevel As Single
    Dim lngRecordCount As Long
    Dim sinPrice As Single
    Dim dbl�ɱ��� As Double
    Dim lngPhysic As Long
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim rsprice As New Recordset
    Dim str�̵㵥NO�� As String
    Dim strMoneyDigit As String
    Dim dbl����, dbl��۲� As Double
    
'    On Error Resume Next
    On Error GoTo errHandle
    
    Call FS.ShowFlash("���ڼ����������Ͽ������,���Ժ� ...", Me)

    DoEvents
    
    If mint�༭״̬ = 5 Then
        Set rsData = Get���ܼ�¼��(lng�ⷿid, str�̵�ʱ��)
    Else 'mint�༭״̬ = 6��ֻ��5��6�ŵ����˸ù��̣�
        Set rsData = GetDateStock(str�̵�ʱ��, lng�ⷿid, 0, False, IIf(mint�༭״̬ = 5, True, False))
    End If
    Call FS.StopFlash
    
    lngRecordCount = rsData.RecordCount
    If lngRecordCount = 0 Then
        If mint�༭״̬ = 6 Then
            ShowMsgBox "δ����ȷ��ȡ�������Ͽ������,�����ԣ�": Exit Sub
        Else
            ShowMsgBox "δ����ȷ��ȡ�������Ͽ������,�����Ի��ֹ�������ϣ�": Exit Sub
        End If
    End If
    
    Call FS.ShowFlash("����װ���������,���Ժ� ...", Me)
    DoEvents
    mshBill.Redraw = False
    
    rsData.MoveFirst
    i = 1: lngPhysic = 0
    With mshBill
        Do While Not rsData.EOF
            If i > 1 Then .Rows = .Rows + 1
            '�������ID��ͬ�������ʱ�۲��ϣ���ȡʵ�ʵ����ۼ�
            .TextMatrix(i, 0) = rsData!����ID
            lngPhysic = rsData!����ID
            sinPrice = IIf(rsData!�Ƿ��� = 1, 0, IIf(IsNull(rsData!�ۼ�), 0, rsData!�ۼ�))
            
            
            dbl�ɱ��� = Val(zlStr.NVL(rsData!������))
            
            '�����ʱ�۲��ϣ��������ۼ�
            If rsData!�Ƿ��� = 1 Then
                sinPrice = Get���ۼ�(Val(zlStr.NVL(rsData!����ID)), lng�ⷿid, Val(zlStr.NVL(rsData!����)), rsData!����ϵ��)
                .TextMatrix(i, mBillCol.C_�ۼ�) = Format(sinPrice, mFMT.FM_���ۼ�)
            Else
                .TextMatrix(i, mBillCol.C_�ۼ�) = Format(sinPrice, mFMT.FM_���ۼ�)
            End If
            
            If (rsData!���� = -1) Then
                '��ʾ��ʼ�����������Դ򿪼�¼��
                Select Case mintUnit
                    Case 0
                        strUnitQuantity = ",Sum(A.����) AS �̵�����"
                    Case Else
                        strUnitQuantity = ",Sum(A.����/b.����ϵ��) AS �̵�����"
                End Select
                
                str�̵㵥NO�� = Replace(mstr�̵㵥��, "'", "")
                
                gstrSQL = "" & _
                    "   Select /*+rule*/ Nvl(A.����,0) ����,A.����,A.Ч��,A.����,A.���� �ɱ���" & strUnitQuantity & _
                    "   From ҩƷ�շ���¼ A,�������� B,Table(Cast(f_Str2list([3]) As zlTools.t_Strlist)) C" & _
                    "   Where A.ҩƷID+0=[1]" & " And Nvl(A.����,0)=-1 " & _
                    "           And A.NO=C.Column_Value And A.����=23 And A.ҩƷID=B.����ID" & _
                    "   Group By Nvl(����,0),����,Ч��,����,���,A.����"
                
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "--��ȡ��¼����������", lngPhysic, str�̵�ʱ��, str�̵㵥NO��)
                
                Do While Not rsTemp.EOF
'                    If i > 1 Then .Rows = .Rows + 1
                    If rsTemp.AbsolutePosition > 1 Then .Rows = .Rows + 1 '���ص�һ������Ҫ.rows+1
                    .TextMatrix(i, 0) = rsData!����ID
                    .TextMatrix(i, mBillCol.C_�ۼ�) = Format(sinPrice, mFMT.FM_���ۼ�)
                    .TextMatrix(i, mBillCol.C_����) = "[" & rsData!���� & "]" & rsData!��Ʒ����
                    .TextMatrix(i, mBillCol.c_���) = IIf(IsNull(rsData!���), "", rsData!���)
                    .TextMatrix(i, mBillCol.C_����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                    .TextMatrix(i, mBillCol.C_��׼�ĺ�) = IIf(IsNull(rsData!��׼�ĺ�), "", rsData!��׼�ĺ�)
                    .TextMatrix(i, mBillCol.C_�ⷿ��λ) = IIf(IsNull(rsData!�ⷿ��λ), "", rsData!�ⷿ��λ)
                    .TextMatrix(i, mBillCol.c_��λ) = IIf(IsNull(rsData!��λ), "", rsData!��λ)
                    .TextMatrix(i, mBillCol.c_����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                    .TextMatrix(i, mBillCol.C_����) = IIf(IsNull(rsData!����), "", rsData!����)
                    .TextMatrix(i, mBillCol.C_Ч��) = IIf(IsNull(rsTemp!Ч��), "", Format(rsTemp!Ч��, "yyyy-MM-dd"))
                    .TextMatrix(i, mBillCol.C_��������) = Format(Val(zlStr.NVL(rsData!��������)), mFMT.FM_����)
                    .TextMatrix(i, mBillCol.C_ʵ������) = Format(IIf(IsNull(rsTemp!�̵�����), 0, rsTemp!�̵�����), mFMT.FM_����)
                    .TextMatrix(i, mBillCol.C_�̵���) = Format(Val(.TextMatrix(i, mBillCol.C_ʵ������)) * Val(.TextMatrix(i, mBillCol.C_�ۼ�)), mFMT.FM_���)
                    .TextMatrix(i, mBillCol.C_��������) = rsData!��������
                    .TextMatrix(i, mBillCol.C_ʵ�ʽ��) = rsData!ʵ�ʽ��
                    .TextMatrix(i, mBillCol.C_ʵ�ʲ��) = rsData!ʵ�ʲ��
                    .TextMatrix(i, mBillCol.c_����ϵ��) = rsData!����ϵ��
                    .TextMatrix(i, mBillCol.C_ָ�������) = rsData!ָ������� & "||" & rsData!�Ƿ��� & "||" & rsData!���÷���
                    .TextMatrix(i, mBillCol.C_�ɱ���) = Format(Val(rsTemp!�ɱ���) * Val(rsData!����ϵ��), mFMT.FM_�ɱ���)
                    
                    If Val(.TextMatrix(i, mBillCol.C_��������)) > Val(.TextMatrix(i, mBillCol.C_ʵ������)) Then
                        .TextMatrix(i, mBillCol.C_��־) = "��"
                    ElseIf Val(.TextMatrix(i, mBillCol.C_��������)) < Val(.TextMatrix(i, mBillCol.C_ʵ������)) Then
                        .TextMatrix(i, mBillCol.C_��־) = "ӯ"
                    Else
                        .TextMatrix(i, mBillCol.C_��־) = "ƽ"
                    End If
                    .TextMatrix(i, mBillCol.C_������) = Format(Abs(Val(.TextMatrix(i, mBillCol.C_ʵ������)) - Val(.TextMatrix(i, mBillCol.C_��������))), mFMT.FM_����)
                    
                    If Val(.TextMatrix(i, mBillCol.C_��������)) = 0 Then
                        strMoneyDigit = "#0.00000"
                    Else
                        strMoneyDigit = mFMT.FM_���
                    End If
                     '����=��ǰ�ۼ�*ʵ������-ʵ�ʽ��
                    '��۲�=����*iif(ʵ�ʽ��<=0,ָ�������,(ʵ�ʲ��/ʵ�ʽ��))
                    .TextMatrix(i, mBillCol.c_����) = Format(Val(.TextMatrix(i, mBillCol.C_�ۼ�)) * Val(.TextMatrix(i, mBillCol.C_ʵ������)) - Val(.TextMatrix(i, mBillCol.C_ʵ�ʽ��)), strMoneyDigit)
'                    If rsData!�Ƿ��� = 1 And Val(.TextMatrix(i, mBillCol.C_��������)) = 0 Then
'                        .TextMatrix(i, mBillCol.C_��۲�) = Format(Val(.TextMatrix(i, mBillCol.C_������)) * (Val(.TextMatrix(i, mBillCol.C_�ۼ�)) - dbl�ɱ��� * rsData!����ϵ��), strMoneyDigit)
'                    Else
                        .TextMatrix(i, mBillCol.c_��۲�) = Format(Val(.TextMatrix(i, mBillCol.C_ʵ������)) * (Val(.TextMatrix(i, mBillCol.C_�ۼ�)) - Val(.TextMatrix(i, mBillCol.C_�ɱ���))) - Val(.TextMatrix(i, mBillCol.C_ʵ�ʲ��)), strMoneyDigit)
'                    End If
                    
                    dbl���� = .TextMatrix(i, mBillCol.c_����)
                    dbl��۲� = .TextMatrix(i, mBillCol.c_��۲�)
                    
                    If .TextMatrix(i, mBillCol.C_��־) = "��" Then
                        '��֤ʵ�ʽ���������ͬ�ķ��ţ���Ϊ���ϵ��Ϊ-1���������ܱ�֤��ȫ����Ϊ�㣩
                        If Not ��ͬ����(Val(.TextMatrix(i, mBillCol.c_����)), Val(.TextMatrix(i, mBillCol.C_ʵ�ʽ��))) Then
                            .TextMatrix(i, mBillCol.c_����) = Format(-1 * Val(.TextMatrix(i, mBillCol.c_����)), strMoneyDigit)
                        End If
                        If Not ��ͬ����(Val(.TextMatrix(i, mBillCol.c_��۲�)), Val(.TextMatrix(i, mBillCol.C_ʵ�ʲ��))) Then
                            .TextMatrix(i, mBillCol.c_��۲�) = Format(-1 * Val(.TextMatrix(i, mBillCol.c_��۲�)), strMoneyDigit)
                        End If
                    End If
                    .TextMatrix(i, mBillCol.C_�̵�ɱ����) = Format(Val(.TextMatrix(i, mBillCol.C_ʵ�ʽ��)) + dbl���� - (Val(.TextMatrix(i, mBillCol.C_ʵ�ʲ��)) + dbl��۲�), mFMT.FM_���)
                    .TextMatrix(i, mBillCol.C_�̵�ɱ�����) = Format(Val(.TextMatrix(i, mBillCol.c_����)) - Val(.TextMatrix(i, mBillCol.c_��۲�)), mFMT.FM_���)
                    
                    i = i + 1
                    rsTemp.MoveNext
                Loop
                i = i - 1
            Else
                .TextMatrix(i, mBillCol.C_����) = "[" & rsData!���� & "]" & rsData!��Ʒ����
                .TextMatrix(i, mBillCol.c_���) = IIf(IsNull(rsData!���), "", rsData!���)
                .TextMatrix(i, mBillCol.C_����) = IIf(IsNull(rsData!����), "", rsData!����)
                .TextMatrix(i, mBillCol.C_��׼�ĺ�) = IIf(IsNull(rsData!��׼�ĺ�), "", rsData!��׼�ĺ�)
                .TextMatrix(i, mBillCol.C_�ⷿ��λ) = IIf(IsNull(rsData!�ⷿ��λ), "", rsData!�ⷿ��λ)
                .TextMatrix(i, mBillCol.c_��λ) = IIf(IsNull(rsData!��λ), "", rsData!��λ)
                .TextMatrix(i, mBillCol.c_����) = IIf(IsNull(rsData!����), "", rsData!����)
                .TextMatrix(i, mBillCol.C_����) = IIf(IsNull(rsData!����), "", rsData!����)
                
                If Val(.TextMatrix(i, mBillCol.C_����)) <> 0 Then
                    .TextMatrix(i, mBillCol.c_���ű༭) = rsData!���ű༭
                    .TextMatrix(i, mBillCol.c_���ر༭) = rsData!���ر༭
                End If
                
                .TextMatrix(i, mBillCol.C_Ч��) = IIf(IsNull(rsData!Ч��), "", Format(rsData!Ч��, "yyyy-MM-dd"))
                .TextMatrix(i, mBillCol.C_��������) = Format(IIf(IsNull(rsData!��������), 0, rsData!��������), mFMT.FM_����)
                If mint�༭״̬ = 5 Then
                    .TextMatrix(i, mBillCol.C_ʵ������) = Format(IIf(IsNull(rsData!�̵�����), 0, rsData!�̵�����), mFMT.FM_����)
                Else
                    .TextMatrix(i, mBillCol.C_ʵ������) = Format(0, mFMT.FM_����)
                End If
                .TextMatrix(i, mBillCol.C_�̵���) = Format(Val(.TextMatrix(i, mBillCol.C_ʵ������)) * Val(.TextMatrix(i, mBillCol.C_�ۼ�)), mFMT.FM_���)
                .TextMatrix(i, mBillCol.C_��������) = rsData!��������
                .TextMatrix(i, mBillCol.C_ʵ�ʽ��) = rsData!ʵ�ʽ��
                .TextMatrix(i, mBillCol.C_ʵ�ʲ��) = rsData!ʵ�ʲ��
                .TextMatrix(i, mBillCol.c_����ϵ��) = rsData!����ϵ��
                .TextMatrix(i, mBillCol.C_�ɱ���) = Format(Val(zlStr.NVL(rsData!�ɱ���)), mFMT.FM_�ɱ���)
                
                .TextMatrix(i, mBillCol.C_ָ�������) = rsData!ָ������� & "||" & rsData!�Ƿ��� & "||" & rsData!���÷���
                If Val(.TextMatrix(i, mBillCol.C_��������)) > Val(.TextMatrix(i, mBillCol.C_ʵ������)) Then
                    .TextMatrix(i, mBillCol.C_��־) = "��"
                ElseIf Val(.TextMatrix(i, mBillCol.C_��������)) < Val(.TextMatrix(i, mBillCol.C_ʵ������)) Then
                    .TextMatrix(i, mBillCol.C_��־) = "ӯ"
                Else
                    .TextMatrix(i, mBillCol.C_��־) = "ƽ"
                End If
                .TextMatrix(i, mBillCol.C_������) = Format(Abs(Val(.TextMatrix(i, mBillCol.C_ʵ������)) - Val(.TextMatrix(i, mBillCol.C_��������))), mFMT.FM_����)
                
                
                If Val(.TextMatrix(i, mBillCol.C_��������)) = 0 Then
                    strMoneyDigit = "#0.00000"
                Else
                    strMoneyDigit = mFMT.FM_���
                End If
                 '����=��ǰ�ۼ�*ʵ������-ʵ�ʽ��
                '��۲�=����*iif(ʵ�ʽ��<=0,ָ�������,(ʵ�ʲ��/ʵ�ʽ��))
                .TextMatrix(i, mBillCol.c_����) = Format(Val(.TextMatrix(i, mBillCol.C_�ۼ�)) * Val(.TextMatrix(i, mBillCol.C_ʵ������)) - Val(.TextMatrix(i, mBillCol.C_ʵ�ʽ��)), strMoneyDigit)
                If rsData!�Ƿ��� = 1 And Val(.TextMatrix(i, mBillCol.C_��������)) = 0 Then
                    .TextMatrix(i, mBillCol.c_��۲�) = Format(Val(.TextMatrix(i, mBillCol.C_������)) * (Val(.TextMatrix(i, mBillCol.C_�ۼ�)) - dbl�ɱ���) - Val(.TextMatrix(i, mBillCol.C_ʵ�ʲ��)), strMoneyDigit)
                Else
                    .TextMatrix(i, mBillCol.c_��۲�) = Format(Val(.TextMatrix(i, mBillCol.C_ʵ������)) * (Val(.TextMatrix(i, mBillCol.C_�ۼ�)) - Val(.TextMatrix(i, mBillCol.C_�ɱ���))) - Val(.TextMatrix(i, mBillCol.C_ʵ�ʲ��)), strMoneyDigit)
                End If
                dbl���� = .TextMatrix(i, mBillCol.c_����)
                dbl��۲� = .TextMatrix(i, mBillCol.c_��۲�)
                
                If .TextMatrix(i, mBillCol.C_��־) = "��" Then
                    '��֤ʵ�ʽ���������ͬ�ķ��ţ���Ϊ���ϵ��Ϊ-1���������ܱ�֤��ȫ����Ϊ�㣩
                    If Not ��ͬ����(Val(.TextMatrix(i, mBillCol.c_����)), Val(.TextMatrix(i, mBillCol.C_ʵ�ʽ��))) Then
                        .TextMatrix(i, mBillCol.c_����) = Format(-1 * Val(.TextMatrix(i, mBillCol.c_����)), strMoneyDigit)
                    End If
                    If Not ��ͬ����(Val(.TextMatrix(i, mBillCol.c_��۲�)), Val(.TextMatrix(i, mBillCol.C_ʵ�ʲ��))) Then
                        .TextMatrix(i, mBillCol.c_��۲�) = Format(-1 * Val(.TextMatrix(i, mBillCol.c_��۲�)), strMoneyDigit)
                    End If
                End If
            End If
            
            .TextMatrix(i, mBillCol.C_�̵�ɱ����) = Format(Val(.TextMatrix(i, mBillCol.C_ʵ�ʽ��)) + dbl���� - (Val(.TextMatrix(i, mBillCol.C_ʵ�ʲ��)) + dbl��۲�), mFMT.FM_���)
            .TextMatrix(i, mBillCol.C_�̵�ɱ�����) = Format(Val(.TextMatrix(i, mBillCol.c_����)) - Val(.TextMatrix(i, mBillCol.c_��۲�)), mFMT.FM_���)
            
            Call ShowPercent(i / lngRecordCount)
            i = i + 1
nextloop:
            rsData.MoveNext
        Loop
        Call RefreshRowNO(mshBill, mBillCol.C_�к�, 1)
        .Redraw = True
    End With
    Call FS.StopFlash
    Call ��ʾ�ϼƽ��
    stbThis.Panels(2).Text = ""
    mshBill.Row = 1: mshBill.Col = mBillCol.C_ʵ������
    If Me.Visible = True Then
        mshBill.SetFocus
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ShowPercent(sngPercent As Single)
    '����:��״̬���ϸ��ݰٷֱ���ʾ��ǰ�������(��)
    Dim intAll As Integer
    intAll = stbThis.Panels(2).Width / TextWidth("��") - 4
    stbThis.Panels(2).Text = Format(sngPercent, "0% ") & String(intAll * sngPercent, "��")
End Sub



Private Function GetDateStock(str�̴�ʱ�� As String, lng�ⷿid As Long, int�̵㷽ʽ As Integer, _
    Optional blnZero As Boolean = False, Optional ByVal bln���� As Boolean = False, Optional str����ID As String = "", _
    Optional lng����ID As Long = 0, Optional bln�̵����������н�� As Boolean = False, Optional ByVal str�ⷿ��λ As String = "����") As ADODB.Recordset
    '���ܣ���ȡָ������������ָ��ʱ���Ŀ�漰�����Ϣ
    '������str�̴�ʱ��=Ҫ����YYYY-MM-DD HH24:MI:SSΪ��ʽ��ʱ���ַ���
    '      int�̵㷽ʽ: ��0-�Զ������̵��1-ÿ�� ;2-ÿ�� ;3-ÿ�� ;4-ÿ���� ;5-�����̵㷽ʽ����0-��ʾ���Զ������̵��
    '      bln����ID:Ϊstr������������ʾ������ID���й���
    '      blnZero=�Ƿ��ȡ��������Ϊ0�Ĳ���,ȱʡ��.��ǿ������ò���ʱ,����Ϊ�ǡ�
    Dim rsTmp As New ADODB.Recordset
    Dim strUnitQuantity As String
    Dim strUnit As String
    Dim blnStock As Boolean
    Dim strOrder As String, strCompare As String
    Dim strRule As String
    Dim str�������� As String
    
    On Error GoTo errH
    
    '������ϲ�ѯ����(�������� B)
    str�������� = " And (c.����ʱ��>[2] Or c.����ʱ�� is NULL)"
    
    If int�̵㷽ʽ <> 5 And int�̵㷽ʽ <> 0 Then '�����̵㷽ʽ
        str�������� = str�������� & " And Substr(E.�̵�����," & int�̵㷽ʽ & ",1)='1' "
    End If

    strOrder = zlDatabase.GetPara("��������", glngSys, mlngModule, "00")
    strOrder = IIf(strOrder = "", "00", strOrder)
    strCompare = Mid(strOrder, 1, 1)
    
    blnStock = CheckPartProp(lng�ⷿid)

    If str�ⷿ��λ = "" Then
        str�ⷿ��λ = "����"
    ElseIf str�ⷿ��λ <> "����" Then
        str�ⷿ��λ = Replace(str�ⷿ��λ, "'", "")
        str�ⷿ��λ = "," & str�ⷿ��λ & ","
    End If
    
    If int�̵㷽ʽ <> 0 Then '�Զ������̵��ʱ���� ���ⷿ��λ��
        str�������� = str�������� & IIf(str�ⷿ��λ = "����", "", " and (Instr([6], ',' || e.�ⷿ��λ || ',') > 0) ")
    End If
    
    If lng����ID > 0 Then '����Ĳ��ϣ����ϲ�ѯ������д
        str�������� = " And B.����ID=[4] "
    End If
    
    'ȡ�õ�ǰ���
    gstrSQL = "" & _
        "   SELECT a.�ⷿid, b.����id, NVL (a.����, 0) AS ����, a.ʵ������,0 �̵�����,a.ʵ�ʽ��, a.ʵ�ʲ��, a.��������, a.ƽ���ɱ��� �ɱ���,a.�ϴ����� AS ����,a.�ϴβ��� AS ����,a.��׼�ĺ�,a.Ч��, e.�ⷿ��λ " & _
        "   FROM ҩƷ��� a, �������� b,�շ���ĿĿ¼ c ,������ĿĿ¼ D, ���ϴ����޶� E" & _
        "   Where a.ҩƷid = b.����id and a.ҩƷid=c.id and b.����id=d.id  " & _
        "           and a.�ⷿid = e.�ⷿid" & IIf(int�̵㷽ʽ = 5, "(+)", "") & " And a.ҩƷid = e.����id" & IIf(int�̵㷽ʽ = 5, "(+)", "") & _
        "           AND a.����=1 " & _
        "           AND a.�ⷿid =[1] " & str�������� & IIf(str����ID = "", "", " and D.����id in (select /*+cardinality(X,10)*/ * from Table(Cast(f_Num2List([5]) As zlTools.t_NumList)) X) ")
    
    gstrSQL = gstrSQL & _
        "   UNION ALL " & _
        "   SELECT a.�ⷿid, b.����id, NVL (a.����, 0) AS ����, " & _
        "           -SUM (DECODE (a.���ϵ��, 1, a.ʵ������*a.����, -a.ʵ������*a.����)) AS ʵ������,0 �̵�����, " & _
        "           -SUM (DECODE (a.���ϵ��, 1, a.���۽��, -a.���۽��)) AS ʵ�ʽ��," & _
        "           -SUM (DECODE (a.���ϵ��, 1, a.���, -a.���)) AS ʵ�ʲ��,0 AS ��������,max(decode(a.����,22,a.����,23,0,a.�ɱ���)) as �ɱ���,a.����,a.��׼�ĺ�,a.����,a.Ч��,Max(e.�ⷿ��λ) As �ⷿ��λ " & _
        "   FROM ҩƷ�շ���¼ a,  �������� b,�շ���ĿĿ¼ c ,������ĿĿ¼ D,���ϴ����޶� E,�շ�ִ�п��� G " & _
        "   Where a.ҩƷid + 0 = b.����id and a.ҩƷid +0 =c.id and b.����id=d.id " & _
        "           and a.�ⷿid + 0 = e.�ⷿid" & IIf(int�̵㷽ʽ = 5, "(+)", "") & " And a.ҩƷid + 0 = e.����id" & IIf(int�̵㷽ʽ = 5, "(+)", "") & _
        "           AND a.�ⷿid + 0 =[1] " & _
        "           and b.����id=g.�շ�ϸĿid " & IIf(mbln���޴洢�ⷿ����, "(+)", "") & _
        "           and G.ִ�п���id" & IIf(mbln���޴洢�ⷿ����, "(+)", "") & "=[1] " & _
        "           AND a.������� >[2] " & str�������� & IIf(str����ID = "", "", " and D.����id in (select /*+cardinality(X,10)*/ * from Table(Cast(f_Num2List([5]) As zlTools.t_NumList)) X) ") & _
        " GROUP BY a.�ⷿid, b.����id, a.����,a.����,a.����,a.��׼�ĺ�,a.Ч�� "
    
    If bln���� Then
        gstrSQL = gstrSQL & _
            "   UNION ALL" & _
            "   SELECT A.�ⷿID,B.����ID,NVL(A.����, 0) AS ����,0 AS ʵ������,SUM(A.����) �̵�����," & _
            "           0 AS ʵ�ʽ��,0 AS ʵ�ʲ��,0 AS ��������,0 as �ɱ���,A.����,A.����,A.��׼�ĺ�,A.Ч��,e.�ⷿ��λ " & _
            "   FROM ҩƷ�շ���¼ A, �������� b,�շ���ĿĿ¼ c ,������ĿĿ¼ D,���ϴ����޶� E" & _
            "   Where A.ҩƷID+0 = B.����ID And A.���� = 23 and a.ҩƷid+0=c.id and b.����id=d.id" & _
            "           AND a.�ⷿid + 0 = e.�ⷿid" & IIf(int�̵㷽ʽ = 5, "(+)", "") & " And a.ҩƷid + 0 = e.����id" & IIf(int�̵㷽ʽ = 5, "(+)", "") & " AND A.�ⷿID + 0 =[1] " & _
            "           AND A.Ƶ�� =[3] " & str�������� & IIf(str����ID = "", "", " and D.����id in (select /*+cardinality(X,10)*/ * from Table(Cast(f_Num2List([5]) As zlTools.t_NumList)) X) ") & _
            "           AND (c.����ʱ�� >[2] OR c.����ʱ�� IS NULL)" & _
            " GROUP BY A.�ⷿID,B.����ID,A.����,A.����,A.����,A.��׼�ĺ�,A.Ч��,e.�ⷿ��λ"
    End If
    
    'ȡ���̵�ʱ����һ�̵���������
    gstrSQL = "" & _
        "   SELECT �ⷿid, ����id, ����, SUM (ʵ������) AS ��������,SUM (�̵�����) AS �̵�����," & _
        "           SUM (ʵ�ʽ��) AS ʵ�ʽ��, SUM (ʵ�ʲ��) AS ʵ�ʲ��, " & _
        "           SUM(��������) As ��������,max(�ɱ���) as �ɱ���,max(����) as ����,max(����) as ���� ,max(��׼�ĺ�) as ��׼�ĺ�,max(Ч��) as Ч��,Max(�ⷿ��λ) As �ⷿ��λ " & _
        "   FROM ( " & gstrSQL & ") " & _
        "   GROUP BY �ⷿid, ����id, ���� " & _
       IIf(bln�̵����������н��, "   Having sum(ʵ������)=0 and (sum(ʵ�ʽ��)<>0 or sum(ʵ�ʲ��)<>0 )", "")
    
    
    Select Case mintUnit
        Case 0
            strUnitQuantity = "c.���㵥λ AS ��λ, nvl(a.��������,0) AS ��������, nvl(a.�̵�����,0) AS �̵�����,nvl(a.��������,0) AS ��������, '1' as ����ϵ��," & _
             " f.�ּ� as �ۼ�,decode(nvl(a.�ɱ���,0),0,B.�ɱ���,A.�ɱ���) as �ɱ���,b.�ɱ��� as ������,"
        Case Else
            strUnitQuantity = "b.��װ��λ AS ��λ, (nvl(a.��������,0) / b.����ϵ��) AS ��������, (nvl(a.�̵�����,0) / b.����ϵ��) AS �̵�����,(nvl(a.��������,0) / b.����ϵ��) AS ��������,b.����ϵ�� as ����ϵ��," & _
             "f.�ּ�*b.����ϵ�� as �ۼ�,decode(nvl(a.�ɱ���,0),0,B.�ɱ���,A.�ɱ���)*b.����ϵ�� as �ɱ���,b.�ɱ���*b.����ϵ�� as ������,"
    End Select
    '�ɱ��ۼ���ķ�ʽ:
    'a.���������
    '          1.�������:�ɱ���=(�����-�����)/�������,
    '          2.�޿������:�ϴγɱ���:��ȡ�������Եĳɱ���
    'b.�������
    '          1.�п��,ȡ�����ϴβɹ���,
    '          2.�޿������:�ϴγɱ���:��ȡ�������Եĳɱ���
    gstrSQL = "" & _
        "   SELECT  DISTINCT b.����id, c.����, c.���� AS ��Ʒ����," & _
        "           zlSpellCode(c.����) ����,c.���, Decode(a.����, Null, decode(b.�ϴβ���,null,c.����,b.�ϴβ���), a.����) As ����,A.��׼�ĺ�,a.�ⷿ��λ,nvl(a.����,0) ����, a.����, a.Ч��," & strUnitQuantity & _
        "           nvl(a.ʵ�ʽ��,0) as ʵ�ʽ�� ,nvl(a.ʵ�ʲ��,0) as ʵ�ʲ��, b.ָ�������,c.�Ƿ���,b.�ⷿ����,b.���÷���,nvl(b.���Ч��,0) ���Ч��,decode(a.����,null,1,0) ���ű༭,decode(a.����,null,1,0) ���ر༭ " & _
        "   From (" & gstrSQL & ") A , �������� b,�շ���ĿĿ¼ c ,�շ�ִ�п��� G,�շѼ�Ŀ F "
        
    gstrSQL = gstrSQL & IIf(blnZero And str����ID <> "", ", (select D.ID From ������ĿĿ¼ D  where 1=1 " & IIf(str����ID = "", "", " and D.����id in (select /*+cardinality(X,10)*/ * from Table(Cast(f_Num2List([5]) As zlTools.t_NumList)) X) ") & ") D ", "") & _
        "   Where " & IIf(blnZero = False, "a.����id = b.����id and a.����id=c.id  ", " b.����id = a.����id(+) and b.����id=c.id(+) " & IIf(str����ID <> "", " and b.����id=d.id", "")) & _
        "           AND b.����id=f.�շ�ϸĿid " & _
        "           and b.����id=g.�շ�ϸĿid " & IIf(mbln���޴洢�ⷿ����, "(+)", "") & _
        "           and G.ִ�п���id" & IIf(mbln���޴洢�ⷿ����, "(+)", "") & "=[1] " & _
        "           and ((SYSDATE BETWEEN f.ִ������ AND f.��ֹ����) OR (SYSDATE >= f.ִ������ AND f.��ֹ���� IS NULL)) " & _
        GetPriceClassString("F") & _
                    IIf(blnZero = False, " AND (a.��������<>0 or nvl(a.ʵ�ʽ��,0)<>0 or nvl(a.ʵ�ʲ��,0)<>0 Or nvl(a.�̵�����,0)<>0)", "") & _
                    IIf(lng����ID > 0, " And B.����ID=[4] ", "") & _
        " ORDER BY " & IIf(strCompare = "0", "c.����", IIf(strCompare = "1", "c.����", IIf(strCompare = "2", "c.����", "a.�ⷿ��λ"))) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
        
    Screen.MousePointer = 11
    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "���������̵����", lng�ⷿid, CDate(str�̴�ʱ��), str�̴�ʱ��, lng����ID, str����ID, str�ⷿ��λ)
    
    Set GetDateStock = rsTmp
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Me.Refresh
        Resume
    End If
    Call SaveErrLog

End Function

Private Function Get���ܼ�¼��(ByVal lng�ⷿid As Long, ByVal str�̵�ʱ�� As String) As ADODB.Recordset
    '--------------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ����̵��¼�����л���ͳ��
    '������lng�ⷿID-�ⷿID
    '      str�̵�ʱ�� -�̵�ʱ��:��ʽyyyy-mm-dd hh24:mi:ss
    '���أ����ط��������ļ�¼
    '--------------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strUnitQuantity As String
    Dim strUnit As String
    Dim str�̵㵥 As String
    Dim blnStock As Boolean
    Dim strOrder As String, strCompare As String
    Dim str�̵㵥NO�� As String
    
    On Error GoTo errH
    strOrder = zlDatabase.GetPara("��������", glngSys, mlngModule, "00")
    strOrder = IIf(strOrder = "", "00", strOrder)
    strCompare = Mid(strOrder, 1, 1)
    
    str�̵㵥NO�� = Replace(mstr�̵㵥��, "'", "")
    
    blnStock = CheckPartProp(lng�ⷿid)
    
    'ȡ�õ�ǰ���
    gstrSQL = "" & _
        "   SELECT a.�ⷿid, b.����id, NVL (a.����,0) AS ����,a.ʵ������,0 �̵�����,a.ʵ�ʽ��, a.ʵ�ʲ��, a.��������,A.ƽ���ɱ��� as �ɱ���,a.�ϴ����� AS ����,a.�ϴβ��� AS ����,A.��׼�ĺ�,a.Ч�� " & _
        "   FROM ҩƷ��� a, �������� b " & _
        "   Where a.ҩƷid = b.����id " & _
        "           AND a.����=1 " & _
        "           AND a.�ⷿid =[1] "
    gstrSQL = gstrSQL & _
        "   UNION ALL " & _
        "   SELECT a.�ⷿid, b.����id, NVL (a.����, 0) AS ����, " & _
        "           -SUM (DECODE (a.���ϵ��, 1, a.ʵ������*a.����, -a.ʵ������*a.����)) AS ʵ������,0 �̵�����, " & _
        "           -SUM (DECODE (a.���ϵ��, 1, a.���۽��, -a.���۽��)) AS ʵ�ʽ��," & _
        "           -SUM (DECODE (a.���ϵ��, 1, a.���, -a.���)) AS ʵ�ʲ��,0 AS ��������,max(decode(a.����,22,a.����,23,0,a.�ɱ���)) as �ɱ���,a.����,a.����,a.��׼�ĺ�,a.Ч�� " & _
        "   FROM ҩƷ�շ���¼ a,  �������� b" & _
        "   Where a.ҩƷid = b.����id" & _
        "           AND a.�ⷿid + 0 =[1] " & _
        "           AND a.������� >[2] " & _
        " GROUP BY a.�ⷿid, b.����id, a.����,a.����,a.����,a.��׼�ĺ�,a.Ч�� "
        
    str�̵㵥 = "" & _
            "   SELECT A.�ⷿID,B.����ID,NVL(A.����, 0) AS ����,0 AS ʵ������,SUM(A.����) �̵�����," & _
            "           0 AS ʵ�ʽ��,0 AS ʵ�ʲ��,0 AS ��������,a.���� as �ɱ���,A.����,A.����,A.��׼�ĺ�,A.Ч��" & _
            "   FROM ҩƷ�շ���¼ A,�������� b " & _
            "   Where A.ҩƷID = B.����ID And A.���� = 23 AND A.�ⷿID + 0 =[1] " & _
            "           AND A.No in (select * from Table(Cast(f_Str2list([3]) As zlTools.t_Strlist))) " & _
            " GROUP BY A.�ⷿID,B.����ID,A.����,A.����,A.����,a.��׼�ĺ�,A.Ч��,a.����"
    
    gstrSQL = gstrSQL & _
        "   UNION ALL" & vbCrLf & str�̵㵥
    
    
    'ȡ���̵�ʱ����һ�̵���������
    gstrSQL = "" & _
        "Select �ⷿID,����ID,����,max(a.�ɱ���) as �ɱ���,max(����) ����,max(����) ���� ,max(��׼�ĺ�) as ��׼�ĺ�,max(Ч��) Ч��," & _
        "       sum(nvl(��������,0)) ��������," & _
        "       sum(nvl(ʵ������,0)) ��������," & _
        "       sum(nvl(�̵�����,0)) �̵�����," & _
        "       sum(nvl(ʵ�ʽ��,0)) ʵ�ʽ��," & _
        "       sum(nvl(ʵ�ʲ��,0)) ʵ�ʲ��" & _
        "   From (" & gstrSQL & ") a" & _
        IIf(mblnֻͳ���̵㵥����, _
            " where  Exists (Select 1 from ҩƷ�շ���¼ T1 " & _
            "                where T1.NO in (select * from Table(Cast(f_Str2list([3]) as zlTools.t_Strlist))) and T1.����=23 and a.����ID=T1.ҩƷid+0 ) ", _
            "") & _
        "  Group by �ⷿID,����ID,����"
    
    Select Case mintUnit
        Case 0
            strUnitQuantity = "c.���㵥λ AS ��λ, nvl(a.��������,0) AS ��������, nvl(a.�̵�����,0) AS �̵�����,nvl(a.��������,0) AS ��������, '1' as ����ϵ��," & _
             " f.�ۼ�,decode(a.�ɱ���,null,B.�ɱ���,A.�ɱ���) as �ɱ���,b.�ɱ��� as ������,"
        Case Else
            strUnitQuantity = "b.��װ��λ AS ��λ, (nvl(a.��������,0) / b.����ϵ��) AS ��������, (nvl(a.�̵�����,0) / b.����ϵ��) AS �̵�����,(nvl(a.��������,0) / b.����ϵ��) AS ��������,b.����ϵ�� as ����ϵ��," & _
             "f.�ۼ�*b.����ϵ�� as �ۼ�,decode(a.�ɱ���,null,B.�ɱ���,A.�ɱ���)*b.����ϵ�� as �ɱ���,b.�ɱ���*b.����ϵ�� as ������, "
    End Select
    
    gstrSQL = "" & _
        "   SELECT  DISTINCT b.����id, c.����, c.���� AS ��Ʒ����," & _
        "           zlSpellCode(c.����) ����,c.���, a.����,a.��׼�ĺ�,e.�ⷿ��λ," & _
        "           nvl(a.����,0) ����, a.����, a.Ч��," & strUnitQuantity & _
        "           nvl(a.ʵ�ʽ��,0) as ʵ�ʽ�� ,nvl(a.ʵ�ʲ��,0) as ʵ�ʲ��, b.ָ�������,c.�Ƿ���,b.�ⷿ����,b.���÷���,decode(a.����,null,1,0) ���ű༭,decode(a.����,null,1,0) ���ر༭ " & _
        "   From (" & gstrSQL & ") A , �������� b,�շ���ĿĿ¼ c ,���ϴ����޶� e, " & _
        "   (select �շ�ϸĿid,ִ�п���id from �շ�ִ�п��� where ִ�п���id=[1]) G, " & _
        "        (SELECT �շ�ϸĿid, �ּ� as �ۼ� From �շѼ�Ŀ  WHERE ((SYSDATE BETWEEN ִ������ AND ��ֹ����) OR (SYSDATE >= ִ������ AND ��ֹ���� IS NULL))" & _
        GetPriceClassString("") & ") f " & _
        "   Where a.����id = b.����id and a.����id=c.id   AND b.����id=f.�շ�ϸĿid " & _
        "         and (c.����ʱ�� is null or c.����ʱ��>[2]) " & _
        "           and A.�ⷿid=E.�ⷿid(+) and A.����id=E.����id(+) and b.����id=g.�շ�ϸĿid " & IIf(mbln���޴洢�ⷿ����, "(+)", "") & _
        " ORDER BY " & IIf(strCompare = "0", "c.����", IIf(strCompare = "1", "c.����", IIf(strCompare = "2", "c.����", "e.�ⷿ��λ"))) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
        
    Screen.MousePointer = 11
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�������������̵��¼", lng�ⷿid, CDate(str�̵�ʱ��), str�̵㵥NO��)
    
    Set Get���ܼ�¼�� = rsTemp
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Me.Refresh
        Resume
    End If
    Call SaveErrLog

End Function


Private Sub Form_Load()

    Dim strReg As String
    mintUnit = Val(zlDatabase.GetPara("�̵��λ", glngSys, mlngModule, "0"))
    mbln���޴洢�ⷿ���� = Val(zlDatabase.GetPara("�洢�ⷿ", glngSys, mlngModule, "0"))
    
    mbln�����������Ų��ؿ��� = Val(zlDatabase.GetPara(305, glngSys, 0)) = 1
    
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
    End With
    
    mblnFirst = True
    
    txtNO = mstr���ݺ�
    txtNO.Tag = txtNO.Text
    initCard
    '�ָ����Ի���������
    RestoreWinState Me, App.ProductName, mstrCaption
    '�ָ����Ի��������ú󣬻���Ҫ��Ȩ�޿��Ƶ��н�һ������
    With mshBill
        .ColWidth(mBillCol.C_�̵�ɱ����) = IIf(mblnCostView = True, 1400, 0)
        .ColWidth(mBillCol.C_�̵�ɱ�����) = IIf(mblnCostView = True, 1400, 0)
        .ColWidth(mBillCol.C_�ɱ���) = IIf(mblnCostView = True, 800, 0)
        .ColWidth(mBillCol.c_��۲�) = IIf(mblnCostView = True, 900, 0)
    End With
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsTemp As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim lngRow As Long
    Dim strOrder As String, strCompare As String
    Dim dbl���� As Double
    Dim dbl��۲� As Double
    Dim strMoneyDigit As String
    '�ⷿ
    
    On Error GoTo errHandle
    strOrder = zlDatabase.GetPara("��������", glngSys, mlngModule, "00")
    strOrder = IIf(strOrder = "", "00", strOrder)
    
    strCompare = Mid(strOrder, 1, 1)
    Select Case mint�༭״̬
        Case 1, 5, 6
            Txt������ = UserInfo.�û���
            Txt�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
            
            '�����ȫ����Ϊ�㣬�����Ƿ����δ��˵��̵㵥
'            If mint�༭״̬ = 6 Then
'                gstrSQL = "" & _
'                    "    Select Count(*) Records " & _
'                    "    From ҩƷ�շ���¼" & _
'                    "    Where ����<>23 And ����� Is NULL And �ⷿID=[1]"
'
'                Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "����Ƿ����δ���������ϵ���", mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex))
'                If Not rsTemp.EOF Then
'                    If Not IsNull(rsTemp!Records) Then
'                        If rsTemp!Records <> 0 Then
'                            mintParallelRecord = 5
'                            Exit Sub
'                        End If
'                    End If
'                End If
'            End If
            
            cmd�̶���.Visible = (mint�༭״̬ = 1)
        Case 2, 3, 4
            initGrid
            If mint�༭״̬ <> 4 Then
                txtStock = mfrmMain.cboStock.Text
                txtStock.Tag = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
            Else
                gstrSQL = "" & _
                    "   Select distinct b.id,b.���� " & _
                    "   From ҩƷ�շ���¼ a,���ű� b " & _
                    "   Where a.�ⷿid=b.id " & _
                    "           and A.���� = 22 and a.no=[1]"
                    
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", mstr���ݺ�)
                If rsTemp.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If
                
                txtStock = rsTemp!����
                txtStock.Tag = rsTemp!Id
                rsTemp.Close
            End If
            
            
            Select Case mintUnit
                Case 0
                    strUnitQuantity = "D.���㵥λ AS ��λ, A.��д���� AS ��������,A.���� AS ʵ������, A.ʵ������ AS ������,'1' as ����ϵ��,a.���ۼ� as �ۼ�,A.���� as �ɱ���,"
                Case Else
                    strUnitQuantity = "B.��װ��λ AS ��λ,(A.��д����/ B.����ϵ��) AS ��������,(A.����/ B.����ϵ��) AS ʵ������, (A.ʵ������ / B.����ϵ��) AS ������,B.����ϵ�� as ����ϵ��,a.���ۼ�*B.����ϵ�� as �ۼ�,a.����*B.����ϵ�� as �ɱ���,"
            End Select
            
            gstrSQL = "" & _
                "   Select * " & _
                "   From (  SELECT distinct a.ҩƷid ����id,A.���,('[' || D.���� || ']' || D.����) AS ������Ϣ," & _
                "                   zlSpellCode(D.����) ����,A.���ϵ��,D.���,A.����,A.��׼�ĺ�,C.�ⷿ��λ, A.����,a.Ч��,a.����," & strUnitQuantity & _
                "                   A.���۽�� as ����,A.��� as ��۲�, " & _
                "                   a.ժҪ,������,��������,�����,�������,a.Ƶ�� as �̵�ʱ��,A.����,a.�ɱ��� as �����,a.�ɱ���� as �����,b.ָ�������,d.�Ƿ���,b.���÷���,nvl(a.��ҩ��ʽ,0) as ������,decode(E.�ϴ�����,null,1,0) ���ű༭,decode(E.�ϴβ���,null,1,0) ���ر༭ " & _
                "           FROM ҩƷ�շ���¼ A, �������� b,�շ���ĿĿ¼ D,���ϴ����޶� C,ҩƷ��� E " & _
                "           Where A.ҩƷid = B.����id and a.ҩƷid=d.id  " & _
                "                   And A.ҩƷID=C.����ID(+) And A.�ⷿID=C.�ⷿID(+) AND A.��¼״̬ =[3]" & _
                "                   And A.ҩƷID=E.ҩƷID(+) And A.�ⷿID=E.�ⷿID(+) And nvl(A.����,0) = nvl(E.����(+),0) AND A.���� =[1] AND A.No =[2]" & _
                "       ) " & _
                "   ORDER BY " & IIf(strCompare = "0", "���", IIf(strCompare = "1", "������Ϣ", IIf(strCompare = "2", "����", "�ⷿ��λ"))) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            
            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, 22, mstr���ݺ�, mint��¼״̬)
            mstrTime_Start = GetBillInfo(22, mstr���ݺ�)
            If rsTemp.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            Txt������ = rsTemp!������
            If mint�༭״̬ = 2 Then
                Txt������ = UserInfo.�û���
            End If
            Txt�������� = Format(rsTemp!��������, "yyyy-mm-dd hh:mm:ss")
            
            Txt����� = IIf(IsNull(rsTemp!�����), "", rsTemp!�����)
            Txt������� = IIf(IsNull(rsTemp!�������), "", Format(rsTemp!�������, "yyyy-mm-dd hh:mm:ss"))
            txtժҪ.Text = IIf(IsNull(rsTemp!ժҪ), "", rsTemp!ժҪ)
            txtCheckDate.Caption = rsTemp!�̵�ʱ��
            
            If (mint�༭״̬ = 2 Or mint�༭״̬ = 3) And Txt����� <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            lngRow = 0
            With mshBill
                Do While Not rsTemp.EOF
                    
                    lngRow = lngRow + 1
                    .Rows = lngRow + 1
                    .TextMatrix(lngRow, 0) = rsTemp.Fields(0)
                    .TextMatrix(lngRow, mBillCol.C_����) = rsTemp!������Ϣ
                    .TextMatrix(lngRow, mBillCol.C_���) = rsTemp!���
                    .TextMatrix(lngRow, mBillCol.c_���) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
                    .TextMatrix(lngRow, mBillCol.C_����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                    .TextMatrix(lngRow, mBillCol.C_��׼�ĺ�) = IIf(IsNull(rsTemp!��׼�ĺ�), "", rsTemp!��׼�ĺ�)
                    .TextMatrix(lngRow, mBillCol.C_�ⷿ��λ) = IIf(IsNull(rsTemp!�ⷿ��λ), "", rsTemp!�ⷿ��λ)
                    .TextMatrix(lngRow, mBillCol.c_��λ) = rsTemp!��λ
                    .TextMatrix(lngRow, mBillCol.c_����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                    .TextMatrix(lngRow, mBillCol.C_Ч��) = IIf(IsNull(rsTemp!Ч��), "", Format(rsTemp!Ч��, "yyyy-mm-dd"))
                    
                    .TextMatrix(lngRow, mBillCol.C_ʵ�ʲ��) = Format(rsTemp!�����, mFMT.FM_���)
                    .TextMatrix(lngRow, mBillCol.C_ʵ�ʽ��) = Format(rsTemp!�����, mFMT.FM_���)
                    .TextMatrix(lngRow, mBillCol.C_ָ�������) = Format(rsTemp!ָ�������, mFMT.FM_���) & "||" & rsTemp!�Ƿ��� & "||" & rsTemp!���÷���
                    .TextMatrix(lngRow, mBillCol.C_����) = IIf(IsNull(rsTemp!����), "0", rsTemp!����)
                    .TextMatrix(lngRow, mBillCol.c_������) = rsTemp!������
                    .TextMatrix(lngRow, mBillCol.c_����ϵ��) = rsTemp!����ϵ��
                    
                    If Val(.TextMatrix(lngRow, mBillCol.C_����)) <> 0 Then '��������
                        .TextMatrix(lngRow, mBillCol.c_���ű༭) = rsTemp!���ű༭
                        .TextMatrix(lngRow, mBillCol.c_���ر༭) = rsTemp!���ر༭
                    End If
                    
                    .TextMatrix(lngRow, mBillCol.C_��������) = Format(rsTemp!��������, mFMT.FM_����)
                    .TextMatrix(lngRow, mBillCol.C_ʵ������) = Format(rsTemp!ʵ������, mFMT.FM_����)
                    .TextMatrix(lngRow, mBillCol.C_������) = Format(rsTemp!������, mFMT.FM_����)
                    If rsTemp!ʵ������ > rsTemp!�������� Then
                        .TextMatrix(lngRow, mBillCol.C_��־) = "ӯ"
                    ElseIf rsTemp!ʵ������ < rsTemp!�������� Then
                        .TextMatrix(lngRow, mBillCol.C_��־) = "��"
                    Else
                        .TextMatrix(lngRow, mBillCol.C_��־) = "ƽ"
                    End If
                    
                    If Val(.TextMatrix(lngRow, mBillCol.C_��������)) = 0 Then
                        strMoneyDigit = "#0.00000"
                    Else
                        strMoneyDigit = mFMT.FM_���
                    End If
                    
                    .TextMatrix(lngRow, mBillCol.c_����) = Format(rsTemp!����, strMoneyDigit)
                    .TextMatrix(lngRow, mBillCol.c_��۲�) = Format(rsTemp!��۲�, strMoneyDigit)
                    
                    .TextMatrix(lngRow, mBillCol.C_�ۼ�) = Format(rsTemp!�ۼ�, mFMT.FM_���ۼ�)
                    .TextMatrix(lngRow, mBillCol.C_�ɱ���) = Format(zlStr.NVL(rsTemp!�ɱ���, 0), mFMT.FM_�ɱ���)
                    .TextMatrix(lngRow, mBillCol.C_�̵���) = Format(Val(.TextMatrix(lngRow, mBillCol.C_ʵ������)) * Val(.TextMatrix(lngRow, mBillCol.C_�ۼ�)), mFMT.FM_���)
                    '���������������Ͳ�۲��㷨һ��
                    dbl���� = Val(.TextMatrix(lngRow, mBillCol.c_����)) * rsTemp!���ϵ�� * IIf(mint��¼״̬ = 1, 1, IIf(mint��¼״̬ Mod 3 = 0, 1, -1))
                    dbl��۲� = Val(.TextMatrix(lngRow, mBillCol.c_��۲�)) * rsTemp!���ϵ�� * IIf(mint��¼״̬ = 1, 1, IIf(mint��¼״̬ Mod 3 = 0, 1, -1))
                    '�ɱ����=�ɱ���*ʵ������=(������+����) -(������+��۲�) �ú�����Ϊ�˿��Ʊ������������̵㵥�ܶ���
                    .TextMatrix(lngRow, mBillCol.C_�̵�ɱ����) = Format((zlStr.NVL(rsTemp!�����, 0) + dbl����) - (zlStr.NVL(rsTemp!�����, 0) + dbl��۲�), mFMT.FM_���)
                    .TextMatrix(lngRow, mBillCol.C_�̵�ɱ�����) = Format(Val(.TextMatrix(lngRow, mBillCol.c_����)) - Val(.TextMatrix(lngRow, mBillCol.c_��۲�)), mFMT.FM_���)
                    
                    rsTemp.MoveNext
                Loop
            End With
            rsTemp.Close
    End Select
    Call RefreshRowNO(mshBill, mBillCol.C_�к�, 1)
    Call ��ʾ�ϼƽ��
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'��ʼ���༭�ؼ�
Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = mBillCol.C_Cols
        .ClearBill
        .MsfObj.FixedCols = 1
        
        .TextMatrix(0, mBillCol.C_�к�) = ""
        .TextMatrix(0, mBillCol.C_����) = "���������"
        .TextMatrix(0, mBillCol.C_���) = "���"
        .TextMatrix(0, mBillCol.c_���) = "���"
        .TextMatrix(0, mBillCol.C_����) = "����"
        .TextMatrix(0, mBillCol.C_��׼�ĺ�) = "��׼�ĺ�"
        .TextMatrix(0, mBillCol.C_�ⷿ��λ) = "�ⷿ��λ"
        .TextMatrix(0, mBillCol.c_��λ) = "��λ"
        .TextMatrix(0, mBillCol.c_����) = "����"
        .TextMatrix(0, mBillCol.C_Ч��) = "ʧЧ��"
        .TextMatrix(0, mBillCol.C_����) = "����"
        .TextMatrix(0, mBillCol.C_��������) = "��������"
        .TextMatrix(0, mBillCol.c_����ϵ��) = "����ϵ��"
        .TextMatrix(0, mBillCol.C_ָ�������) = "ָ�������"
        .TextMatrix(0, mBillCol.C_ʵ�ʲ��) = "ʵ�ʲ��"
        .TextMatrix(0, mBillCol.C_ʵ�ʽ��) = "ʵ�ʽ��"
        .TextMatrix(0, mBillCol.C_��������) = "��������"
        .TextMatrix(0, mBillCol.C_ʵ������) = "ʵ������"
        .TextMatrix(0, mBillCol.C_��־) = "��־"
        .TextMatrix(0, mBillCol.C_������) = "������"
        .TextMatrix(0, mBillCol.C_�ɱ���) = "�ɱ���"
        .TextMatrix(0, mBillCol.C_�ۼ�) = "�ۼ�"
        .TextMatrix(0, mBillCol.c_����) = "����"
        .TextMatrix(0, mBillCol.c_��۲�) = "��۲�"
        .TextMatrix(0, mBillCol.C_�̵���) = "�̵���"
        .TextMatrix(0, mBillCol.C_�̵�ɱ����) = "�̵�ɱ����"
        .TextMatrix(0, mBillCol.C_�̵�ɱ�����) = "�̵�ɱ�����"
        .TextMatrix(0, mBillCol.c_������) = "������"
        
        .TextMatrix(0, mBillCol.c_���ű༭) = "���ű༭"
        .TextMatrix(0, mBillCol.c_���ر༭) = "���ر༭"
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mBillCol.C_�к�) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mBillCol.C_�к�) = 300
        .ColWidth(mBillCol.C_����) = 0
        .ColWidth(mBillCol.C_���) = 0
        .ColWidth(mBillCol.C_��������) = 0
        .ColWidth(mBillCol.c_����ϵ��) = 0
        .ColWidth(mBillCol.C_ָ�������) = 0
        .ColWidth(mBillCol.C_ʵ�ʲ��) = 0
        .ColWidth(mBillCol.C_ʵ�ʽ��) = 0
        .ColWidth(mBillCol.C_����) = 2000
        .ColWidth(mBillCol.c_���) = 900
        .ColWidth(mBillCol.C_����) = 800
        .ColWidth(mBillCol.C_��׼�ĺ�) = 1000
        .ColWidth(mBillCol.C_�ⷿ��λ) = 2000
        .ColWidth(mBillCol.c_��λ) = 500
        .ColWidth(mBillCol.c_����) = 800
        .ColWidth(mBillCol.C_Ч��) = 1000
        .ColWidth(mBillCol.C_��������) = 800
        .ColWidth(mBillCol.C_ʵ������) = 800
        .ColWidth(mBillCol.C_��־) = 500
        .ColWidth(mBillCol.C_������) = 800
        .ColWidth(mBillCol.C_�ɱ���) = IIf(mblnCostView = False, 0, 800)
        .ColWidth(mBillCol.C_�ۼ�) = 800
        .ColWidth(mBillCol.c_����) = 900
        .ColWidth(mBillCol.c_��۲�) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mBillCol.C_�̵���) = 900
        .ColWidth(mBillCol.C_�̵�ɱ����) = IIf(mblnCostView = False, 0, 1400)
        .ColWidth(mBillCol.C_�̵�ɱ�����) = IIf(mblnCostView = False, 0, 1500)
        .ColWidth(mBillCol.c_������) = 0
        .ColWidth(mBillCol.c_���ű༭) = 0
        .ColWidth(mBillCol.c_���ر༭) = 0
        
        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��

        .ColData(0) = 5
        .ColData(mBillCol.C_�к�) = 5
        .ColData(mBillCol.c_���) = 5
        .ColData(mBillCol.C_���) = 5
        .ColData(mBillCol.C_����) = 5
        .ColData(mBillCol.C_��׼�ĺ�) = 5
        .ColData(mBillCol.C_�ⷿ��λ) = 5
        .ColData(mBillCol.c_��λ) = 5
        .ColData(mBillCol.c_����) = 5
        .ColData(mBillCol.C_Ч��) = 5
        .ColData(mBillCol.C_����) = 5
        .ColData(mBillCol.C_��������) = 5
        .ColData(mBillCol.c_����ϵ��) = 5
        .ColData(mBillCol.C_ָ�������) = 5
        .ColData(mBillCol.C_ʵ�ʲ��) = 5
        .ColData(mBillCol.C_ʵ�ʽ��) = 5
        .ColData(mBillCol.C_��������) = 5
        
        .ColData(mBillCol.C_��־) = 5
        .ColData(mBillCol.C_������) = 5
        .ColData(mBillCol.C_�ɱ���) = 5
        .ColData(mBillCol.C_�ۼ�) = 5
        .ColData(mBillCol.c_����) = 5
        .ColData(mBillCol.c_��۲�) = 5
        .ColData(mBillCol.C_�̵���) = 5
        .ColData(mBillCol.C_�̵�ɱ����) = 5
        .ColData(mBillCol.C_�̵�ɱ�����) = 5
        .ColData(mBillCol.c_������) = 5
        .ColData(mBillCol.c_���ű༭) = 5
        .ColData(mBillCol.c_���ر༭) = 5
                
        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
            txtժҪ.Enabled = True
            .ColData(mBillCol.C_����) = 1
            .ColData(mBillCol.C_ʵ������) = 4
        ElseIf mint�༭״̬ = 3 Or mint�༭״̬ = 4 Then
            txtժҪ.Enabled = False
            .ColData(mBillCol.C_ʵ������) = 5
        ElseIf mint�༭״̬ = 5 Or mint�༭״̬ = 6 Then
'            .Active = False
            txtժҪ.Enabled = True
            .ColData(mBillCol.C_ʵ������) = 5
        End If
        
        .ColAlignment(mBillCol.C_����) = flexAlignLeftCenter
        .ColAlignment(mBillCol.c_���) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_����) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_��׼�ĺ�) = flexAlignLeftCenter
        .ColAlignment(mBillCol.c_��λ) = flexAlignCenterCenter
        .ColAlignment(mBillCol.c_����) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_Ч��) = flexAlignLeftCenter
        
        .ColAlignment(mBillCol.C_��������) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_ʵ������) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_��־) = flexAlignCenterCenter
        .ColAlignment(mBillCol.C_������) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_�ɱ���) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_�ۼ�) = flexAlignRightCenter
        .ColAlignment(mBillCol.c_����) = flexAlignRightCenter
        .ColAlignment(mBillCol.c_��۲�) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_�̵���) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_�̵�ɱ����) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_�̵�ɱ�����) = flexAlignRightCenter
        .ColAlignment(mBillCol.c_������) = flexAlignRightCenter
        .ColAlignment(mBillCol.c_���ű༭) = flexAlignRightCenter
        .ColAlignment(mBillCol.c_���ر༭) = flexAlignRightCenter
        
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
    
    txtCheckDate.Left = mshBill.Left + mshBill.Width - txtCheckDate.Width
    lblCheckDate.Left = txtCheckDate.Left - lblCheckDate.Width - 100
    
    LblStock.Left = mshBill.Left
    txtStock.Left = LblStock.Left + LblStock.Width + 100
    
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
        .Width = Pic����.TextWidth(.Caption) + 200
        
        lblCheckSum.Left = .Left + .Width + 100
        lblCheckSum.Top = .Top
        lblCheckSum.Width = Pic����.TextWidth(lblCheckSum.Caption) + 200
    End With
    
    With lblCheckSum
        lblCheckCostSum.Left = .Left + .Width + 100
        lblCheckCostSum.Top = .Top
    End With
    
    If mblnCostView = False Then
        lblCheckCostSum.Visible = False
    End If
    
    
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
    
    With cmd�̶���
        .Left = CmdSave.Left - .Width - 150
        .Top = CmdSave.Top
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
    Dim strNo As String
    Dim str����� As String
    
    mblnSave = False
    SaveCheck = False
    
    str����� = UserInfo.�û���
    strNo = txtNO.Tag
    On Error GoTo errHandle
    
    gstrSQL = "zl_�����̵�_Verify('" & strNo & "','" & str����� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
        
        
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function



Private Sub mnuDefault_Click()
    mshBill.MsfObj.FixedCols = 1
    mshBill.ColData(mBillCol.C_����) = 1
    mshBill.LocateCol = mBillCol.C_����
End Sub

Private Sub mnuFirst_Click()
    mshBill.Redraw = False
    mshBill.ColData(mBillCol.C_����) = 5
    mshBill.MsfObj.FixedCols = 14
    mshBill.LocateCol = 17
    
    '���ö��뷽ʽ
    mshBill.MsfObj.ColAlignmentFixed(mBillCol.c_���) = 1
    mshBill.MsfObj.ColAlignmentFixed(mBillCol.c_��λ) = 4
    mshBill.MsfObj.ColAlignmentFixed(mBillCol.c_����) = 1
    mshBill.MsfObj.ColAlignmentFixed(mBillCol.C_Ч��) = 1
    mshBill.Refresh
    mshBill.Redraw = True
End Sub

Private Sub mnuSecond_Click()
    mshBill.Redraw = False
    mshBill.ColData(mBillCol.C_����) = 5
    mshBill.MsfObj.FixedCols = 16
    mshBill.LocateCol = 17
    
    '���ö��뷽ʽ
    mshBill.MsfObj.ColAlignmentFixed(mBillCol.c_���) = 1
    mshBill.MsfObj.ColAlignmentFixed(mBillCol.c_��λ) = 4
    mshBill.MsfObj.ColAlignmentFixed(mBillCol.c_����) = 1
    mshBill.MsfObj.ColAlignmentFixed(mBillCol.C_Ч��) = 1
    mshBill.Refresh
    mshBill.Redraw = True
End Sub

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mBillCol.C_�к�, Row)
    If mshBill.MsfObj.FixedCols > mBillCol.C_���� Then
        mshBill.PrimaryCol = mBillCol.C_ʵ������
        mshBill.Col = mBillCol.C_ʵ������
        mshBill.PrimaryCol = mBillCol.C_����
    End If
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call ��ʾ�ϼƽ��
    Call RefreshRowNO(mshBill, mBillCol.C_�к�, mshBill.Row)
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If InStr(1, "3456", mint�༭״̬) <> 0 Then
        Cancel = True
        Exit Sub
    End If
    With mshBill
        If .TextMatrix(.Row, 0) <> "" Then
            If MsgBox("��ȷʵҪɾ�������������ϣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub mshbill_CommandClick()
    Dim RecReturn As Recordset
    Dim i As Integer
    Dim int����� As Integer
    
    On Error GoTo errHandle
    
    int����� = mshBill.Row
    
    If mshBill.Col = mBillCol.C_���� Then
        Set RecReturn = Frm����ѡ����.ShowMe(Me, 2, txtStock.Tag, txtStock.Tag, txtStock.Tag, False, True, True, True, , , , , txtCheckDate.Caption, , , mbln���޴洢�ⷿ����, mstrPrivs, , False)
        If RecReturn.RecordCount > 0 Then
            mblnChange = True
            
            With mshBill
                RecReturn.MoveFirst
                For i = 1 To RecReturn.RecordCount
                    
                    If SetPhiscRows(RecReturn!����ID, IIf(IsNull(RecReturn!����), 0, RecReturn!����)) Then
    
                        If .Row = .Rows - 1 Then .Rows = .Rows + 1 'ֻ�е�ǰ�������һ��ʱ��������
                        .Row = .Row + 1
                    End If
                    
                    .Col = mBillCol.C_ʵ������
                    RecReturn.MoveNext
                Next
                
                mshBill.Row = int�����
                
                If mstr�ظ����� <> "" Then
                    MsgBox mstr�ظ����� & "�б����Ѿ������ˣ�" & vbCrLf & "�������Ĳ�����ӣ�", vbInformation + vbOKOnly, gstrSysName
                    mstr�ظ����� = ""
                End If
                
    '            If RecReturn.RecordCount = 1 Then
    '                Call SetPhiscRows(RecReturn!����ID, IIf(IsNull(RecReturn!����), 0, RecReturn!����))
    '                .Col = mBillCol.C_ʵ������
    '            End If
            End With
            RecReturn.Close
        End If
    Else
        gstrSQL = "Select rownum as id,null as �ϼ�id,����,����,����,1 as ĩ�� From ���������� "
        Set RecReturn = zlDatabase.ShowSelect(Me, gstrSQL, 1, "����������ѡ��", True, , "ѡ���������������̻���")
  
        If RecReturn Is Nothing Then Exit Sub
        If RecReturn.State <> 1 Then Exit Sub
        
        With RecReturn
            If CheckQualifications(mlngModule, 1, CStr(NVL(!����))) = False Then Exit Sub
            mshBill.TextMatrix(mshBill.Row, mBillCol.C_����) = NVL(!����)
        End With
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Sub

Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub

Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer
    
    With mshBill
        If .Col = mBillCol.C_�������� Or .Col = mBillCol.C_ʵ������ Or .Col = mBillCol.C_�ɱ��� Then
            strKey = .Text
            If strKey = "" Then
                strKey = .TextMatrix(.Row, .Col)
            End If
            Select Case .Col
                Case mBillCol.C_��������, mBillCol.C_ʵ������
                    intDigit = IIf(mintUnit = 1, g_С��λ��.obj_��װС��.����С��, g_С��λ��.obj_ɢװС��.����С��)
                Case mBillCol.C_�ɱ���
                   intDigit = IIf(mintUnit = 1, g_С��λ��.obj_��װС��.�ɱ���С��, g_С��λ��.obj_ɢװС��.�ɱ���С��)
                    intDigit = IIf(mintUnit = 1, g_С��λ��.obj_��װС��.���ۼ�С��, g_С��λ��.obj_ɢװС��.���ۼ�С��)
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
    Dim lng����  As Long
    Dim lng������  As Long
    
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
                Call ��ʾ�����
                If mshBill.MsfObj.FixedCols > mBillCol.C_���� Then
                    mshBill.PrimaryCol = mBillCol.C_ʵ������
                    mshBill.Col = mBillCol.C_ʵ������
                    mshBill.PrimaryCol = mBillCol.C_����
                End If
            Case mBillCol.c_����
                .TxtCheck = False
                .MaxLength = mintBatchNoLen
            
            Case mBillCol.C_Ч��
                .TxtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                If .ColData(mBillCol.C_Ч��) = 2 Then
                    If .TextMatrix(.Row, mBillCol.c_����) <> "" And Len(Trim(.TextMatrix(.Row, mBillCol.c_����))) = 8 Then
                        Dim strxq As String
                        
                        If IsNumeric(.TextMatrix(.Row, mBillCol.c_����)) Then
                            strxq = UCase(.TextMatrix(.Row, mBillCol.c_����))
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                strxq = TranNumToDate(strxq)
                                If strxq <> "" Then .TextMatrix(.Row, mBillCol.C_Ч��) = Format(DateAdd("M", .RowData(.Row), strxq), "yyyy-mm-dd")
                            End If
                        End If
                    End If
                End If
            Case mBillCol.C_ʵ������
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
        End Select
        
        lng���� = Val(.TextMatrix(.Row, mBillCol.C_����))
        If mint�༭״̬ = 1 Then
            .ColData(mBillCol.C_����) = IIf(lng���� = -1 Or Val(.TextMatrix(.Row, mBillCol.c_���ر༭)) = 1, 1, 5)
            .ColData(mBillCol.c_����) = IIf(lng���� = -1 Or Val(.TextMatrix(.Row, mBillCol.c_���ű༭)) = 1, 4, 5)
            .ColData(mBillCol.C_Ч��) = IIf(lng���� = -1, 2, 5)
        End If
        
        If mint�༭״̬ = 2 Or mint�༭״̬ = 5 Or mint�༭״̬ = 6 Then
            lng������ = Val(.TextMatrix(.Row, mBillCol.c_������))
            .ColData(mBillCol.C_����) = IIf(lng������ = 1 Or lng���� = -1 Or Val(.TextMatrix(.Row, mBillCol.c_���ر༭)) = 1, 1, 5)
            .ColData(mBillCol.c_����) = IIf(lng������ = 1 Or lng���� = -1 Or Val(.TextMatrix(.Row, mBillCol.c_���ű༭)) = 1, 4, 5)
            .ColData(mBillCol.C_Ч��) = IIf(lng������ = 1 Or lng���� = -1, 2, 5)
        End If
        
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsDrug As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim dbl����, dbl��۲� As Double
    Dim i As Integer
    Dim int����� As Integer
    Dim strMoneyDigit As String
    int����� = mshBill.Row
    
    On Error GoTo errHandle
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
            
            Case mBillCol.C_����
                If strKey <> "" Then
                    Dim RecReturn As Recordset
                    Dim sngLeft As Single
                    Dim sngTop As Single
                    
                    sngLeft = Me.Left + Pic����.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic����.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 4530
                    End If
                    
                    Set RecReturn = FrmMulitSel.ShowSelect(Me, 2, txtStock.Tag, txtStock.Tag, txtStock.Tag, strKey, sngLeft, sngTop, mshBill.MsfObj.CellWidth, mshBill.MsfObj.CellHeight, False, True, True, True, , , , Me.txtCheckDate.Caption, , , mbln���޴洢�ⷿ����, mstrPrivs, , False)
                    
                    If RecReturn.RecordCount <= 0 Then
                        Cancel = True
                        Exit Sub
                    End If
                        
                    RecReturn.MoveFirst
                    For i = 1 To RecReturn.RecordCount
                        If SetPhiscRows(RecReturn!����ID, IIf(IsNull(RecReturn!����), 0, RecReturn!����)) Then
                            
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
'                        If Not SetPhiscRows(RecReturn!����ID, IIf(IsNull(RecReturn!����), 0, RecReturn!����)) Then
'                            Cancel = True
'                            Exit Sub
'                        End If
'                        .Text = .TextMatrix(.Row, .Col)
'                    Else
'                        Cancel = True
'                    End If
                    
                    Call ��ʾ�����
                End If
            Case mBillCol.C_����
                If strKey = "" Then Exit Sub
                If SelectAndNotAddItem(Me, mshBill, strKey, "����������", "����������ѡ����", True, True, , zl_��ȡվ������(True)) = True Then
                    .Text = .TextMatrix(.Row, .Col)
                Else
                    .Text = ""
                    .Col = mBillCol.C_����
                    Cancel = True
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
                        .Col = mBillCol.C_ʵ������
                    End If
                    
                    Cancel = True
                    Exit Sub
                End If
            Case mBillCol.C_Ч��
                '�д���
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        If strKey = "" Then
                            ShowMsgBox "ʧЧ�ڱ���Ϊ�����ͣ�"
                            Cancel = True
                            Exit Sub
                        End If
                        .Text = strKey
                        Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        ShowMsgBox "ʧЧ�ڱ���Ϊ��������(2000-10-10) ��20001010��,�����䣡"
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
            Case mBillCol.C_ʵ������
                Dim dbl�ɱ��� As Double
                Dim rsTemp As New ADODB.Recordset
                
                If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
                    ShowMsgBox "ʵ�������������룡"
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Not IsNumeric(strKey) And strKey <> "" Then
                    ShowMsgBox "ʵ����������Ϊ������,�����䣡"
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" And .TextMatrix(.Row, 0) <> "" Then
                    strKey = Format(strKey, mFMT.FM_����)
                    .Text = strKey
                    .TextMatrix(.Row, mBillCol.C_������) = Format(Abs(Val(strKey) - Val(.TextMatrix(.Row, mBillCol.C_��������))), mFMT.FM_����)
                    If Val(strKey) > Val(.TextMatrix(.Row, mBillCol.C_��������)) Then
                        .TextMatrix(.Row, mBillCol.C_��־) = "ӯ"
                    ElseIf Val(strKey) < Val(.TextMatrix(.Row, mBillCol.C_��������)) Then
                        .TextMatrix(.Row, mBillCol.C_��־) = "��"
                    Else
                        .TextMatrix(.Row, mBillCol.C_��־) = "ƽ"
                    End If
                    
                    If Val(.TextMatrix(.Row, mBillCol.C_��������)) = 0 Then
                        strMoneyDigit = "#0.00000"
                    Else
                        strMoneyDigit = mFMT.FM_���
                    End If
                    
                    '����=��ǰ�ۼ�*ʵ������-ʵ�ʽ��
                    '��۲�=����*iif(ʵ�ʽ��=0,ָ�������,(ʵ�ʲ��/ʵ�ʽ��))
                    .TextMatrix(.Row, mBillCol.c_����) = Format(Val(.TextMatrix(.Row, mBillCol.C_�ۼ�)) * Val(strKey) - Val(.TextMatrix(.Row, mBillCol.C_ʵ�ʽ��)), strMoneyDigit)
                    .TextMatrix(.Row, mBillCol.c_��۲�) = Format(Val(strKey) * (Val(.TextMatrix(.Row, mBillCol.C_�ۼ�)) - Val(.TextMatrix(.Row, mBillCol.C_�ɱ���))) - Val(.TextMatrix(.Row, mBillCol.C_ʵ�ʲ��)), strMoneyDigit)
                    
                    dbl���� = .TextMatrix(.Row, mBillCol.c_����)
                    dbl��۲� = .TextMatrix(.Row, mBillCol.c_��۲�)
                    
                    If .TextMatrix(.Row, mBillCol.C_��־) = "��" Then    '���������¼�еĽ����۲�ķ���һ��
                        If Val(.TextMatrix(.Row, mBillCol.C_ʵ�ʽ��)) >= 0 Then
                            .TextMatrix(.Row, mBillCol.c_����) = Format(Abs(.TextMatrix(.Row, mBillCol.c_����)), strMoneyDigit)
                        Else
                            .TextMatrix(.Row, mBillCol.c_����) = Format(Abs(.TextMatrix(.Row, mBillCol.c_����)) * -1, strMoneyDigit)
                        End If
                        If Val(.TextMatrix(.Row, mBillCol.C_ʵ�ʲ��)) >= 0 Then
                            .TextMatrix(.Row, mBillCol.c_��۲�) = Format(Abs(.TextMatrix(.Row, mBillCol.c_��۲�)), strMoneyDigit)
                        Else
                            .TextMatrix(.Row, mBillCol.c_��۲�) = Format(Abs(.TextMatrix(.Row, mBillCol.c_��۲�)) * -1, strMoneyDigit)
                        End If
                    End If
                    .TextMatrix(.Row, mBillCol.C_�̵���) = Format(Val(.TextMatrix(.Row, mBillCol.C_�ۼ�)) * Val(strKey), mFMT.FM_���)
                    .TextMatrix(.Row, mBillCol.C_�̵�ɱ����) = Format(Val(.TextMatrix(.Row, mBillCol.C_ʵ�ʽ��)) + dbl���� - (Val(.TextMatrix(.Row, mBillCol.C_ʵ�ʲ��)) + dbl��۲�), mFMT.FM_���)
                    .TextMatrix(.Row, mBillCol.C_�̵�ɱ�����) = Format(Val(.TextMatrix(.Row, mBillCol.c_����)) - Val(.TextMatrix(.Row, mBillCol.c_��۲�)), mFMT.FM_���)
                    
                End If
                Call ��ʾ�ϼƽ��
        End Select
    End With
    Exit Sub
errHandle:
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
    Dim lngLop As Long
    Dim lngЧ�� As Long
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    If txtNO.Locked = False Then
        If Trim(txtNO.Text) = "" Then
            ShowMsgBox "���ݺŲ���Ϊ��"
            Exit Function
        End If
        If LenB(StrConv(txtNO.Text, vbFromUnicode)) > txtNO.MaxLength Then
            ShowMsgBox "���ݺų���,���������" & CInt(txtNO.MaxLength / 2) & "�����֣���ò�Ҫ���֣���" & txtNO.MaxLength & "���ַ�!"
            txtNO.SetFocus
            Exit Function
        End If
        If InStr(1, txtNO.Text, "'") <> 0 Then
            ShowMsgBox "���ݺ��в��ܺ��зǷ��ַ�"
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
        
            For lngLop = 1 To .Rows - 1
                If Trim(.TextMatrix(lngLop, mBillCol.C_����)) <> "" Then
                    If Trim(Trim(.TextMatrix(lngLop, mBillCol.C_ʵ������))) = "" Then
                        ShowMsgBox "��" & lngLop & "���������ϵ�ʵ������Ϊ���ˣ����飡"
                        mshBill.SetFocus
                        .Row = lngLop
                        .MsfObj.TopRow = lngLop
                        .Col = mBillCol.C_ʵ������
                        Exit Function
                    End If
                    If Val(.TextMatrix(lngLop, mBillCol.C_ʵ������)) > 9999999999# Then
                        ShowMsgBox "��" & lngLop & "���������ϵ�ʵ���������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡"
                        mshBill.SetFocus
                        .Row = lngLop
                        .MsfObj.TopRow = lngLop
                        .Col = mBillCol.C_ʵ������
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(lngLop, mBillCol.c_����)) > 9999999999999# Then
                        ShowMsgBox "��" & lngLop & "���������ϵĽ�����������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡"
                        mshBill.SetFocus
                        .Row = lngLop
                        .MsfObj.TopRow = lngLop
                        .Col = mBillCol.C_ʵ������
                        Exit Function
                    End If
                    If Val(.TextMatrix(lngLop, mBillCol.C_������)) > 9999999999999# Then
                        ShowMsgBox "��" & lngLop & "���������ϵ���������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡"
                        mshBill.SetFocus
                        .Row = lngLop
                        .MsfObj.TopRow = lngLop
                        .Col = mBillCol.C_ʵ������
                        Exit Function
                    End If
                
                    If Val(.TextMatrix(lngLop, mBillCol.C_����)) = -1 Or Val(.TextMatrix(lngLop, mBillCol.c_������)) = 1 Then '�������ϱ���¼�������Ϣ
                        If LenB(StrConv(Trim(Trim(.TextMatrix(lngLop, mBillCol.c_����))), vbFromUnicode)) > mintBatchNoLen Then
                            ShowMsgBox "��" & lngLop & "���������ϵ����ų���,���������" & Int(mintBatchNoLen / 2) & "�����ֻ�" & mintBatchNoLen & "���ַ�!"
                            .SetFocus
                            .Row = lngLop
                            .MsfObj.TopRow = lngLop
                            .Col = mBillCol.c_����
                            Exit Function
                        End If
                        
                        '�ж��Ƿ�ΪЧ����������
                        gstrSQL = "Select Nvl(���Ч��,0) Ч�� From �������� Where ����ID=[1]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ�ΪЧ����������", Val(.TextMatrix(lngLop, 0)))
                        
                        lngЧ�� = rsTemp!Ч��
                        If lngЧ�� <> 0 Then
                            If Trim(.TextMatrix(lngLop, mBillCol.c_����)) = "" Or Trim(.TextMatrix(lngLop, mBillCol.C_Ч��)) = "" Then
                                ShowMsgBox "��" & lngLop & "�е�����������Ч�ڲ���,����������ż�Ч��" & vbCrLf & "��Ϣ�������뵥���У�"
                                mshBill.SetFocus
                                .Row = lngLop
                                .MsfObj.TopRow = lngLop
                                If .TextMatrix(lngLop, mBillCol.c_����) = "" Then
                                    .Col = mBillCol.c_����
                                Else
                                    .Col = mBillCol.C_Ч��
                                End If
                                Exit Function
                            End If
                        End If
                        
                        '�жϲ��غ������Ƿ�Ϊ��
                        If mbln�����������Ų��ؿ��� = True Then
                            If Trim(.TextMatrix(lngLop, mBillCol.C_����)) = "" Then  '���ر�������
                                ShowMsgBox "��" & lngLop & "�����������Ƿ������ϣ���¼����أ�"
                                mshBill.SetFocus
                                .Row = lngLop
                                .MsfObj.TopRow = lngLop
                                .Col = mBillCol.C_����
                                Exit Function
                            End If
                            If Trim(.TextMatrix(lngLop, mBillCol.c_����)) = "" Then  '���ر�������
                                ShowMsgBox "��" & lngLop & "�����������Ƿ������ϣ���¼�����ţ�"
                                mshBill.SetFocus
                                .Row = lngLop
                                .MsfObj.TopRow = lngLop
                                .Col = mBillCol.c_����
                                Exit Function
                            End If
                        End If
                        
                    End If
                    
                    If Val(.TextMatrix(lngLop, mBillCol.C_����)) > 0 Then '��������
                        '�жϲ��غ������Ƿ�Ϊ��
                        If mbln�����������Ų��ؿ��� = True Then
                            If Trim(.TextMatrix(lngLop, mBillCol.C_����)) = "" Then  '���ر�������
                                ShowMsgBox "��" & lngLop & "�����������Ƿ������ϣ���¼����أ�"
                                mshBill.SetFocus
                                .Row = lngLop
                                .MsfObj.TopRow = lngLop
                                .Col = mBillCol.C_����
                                Exit Function
                            End If
                            If Trim(.TextMatrix(lngLop, mBillCol.c_����)) = "" Then  '���ر�������
                                ShowMsgBox "��" & lngLop & "�����������Ƿ������ϣ���¼�����ţ�"
                                mshBill.SetFocus
                                .Row = lngLop
                                .MsfObj.TopRow = lngLop
                                .Col = mBillCol.c_����
                                Exit Function
                            End If
                        End If
                    End If
                    
                End If
            Next
        Else
            Exit Function
        End If
    End With
    
    ValidData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function SaveCard() As Boolean
    Dim lng������ID As Long
    Dim int���ϵ�� As Integer
    Dim lng������ID As Integer
    Dim lng�������ID As Integer
    
    Dim chrNo As Variant
    Dim lng��� As Long
    Dim lng�ⷿid As Long
    Dim lng����ID As Long
    Dim str���� As String
    Dim lng����ID As Long
    Dim str���� As String
    Dim datЧ�� As String
    Dim dbl�������� As Double
    Dim dblʵ������ As Double
    Dim dbl������ As Double
    Dim dbl�ۼ� As Double
    Dim dbl�ɱ���  As Double
    Dim dbl���� As Double
    Dim dbl��۲� As Double
    Dim strժҪ As String
    Dim str������ As String
    Dim dat�������� As String
    Dim str�̵�ʱ�� As String
    Dim dbl����� As Double
    Dim dbl����� As Double
    Dim rsTemp As New Recordset
    Dim lngRow As Long
    Dim strArr As Variant
    Dim i As Long
    Dim cllSQL As Collection
    Dim int������ As Integer
    Dim n As Long
    
    On Error GoTo errHandle
    SaveCard = False
    '����������������ID����Ҫ�����в��϶�Ҫ����
    gstrSQL = "" & _
        "   SELECT b.ϵ��,b.id AS ���id " & _
        "   FROM ҩƷ�������� a, ҩƷ������ b " & _
        "   Where a.���id = b.ID " & _
        "       AND a.���� =[1] "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, 37)
    If rsTemp.EOF Then
        ShowMsgBox "û���������������̵����������������������������!"
        Exit Function
    End If
    
    lng������ID = 0
    lng�������ID = 0
    
    rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        If rsTemp!ϵ�� = 1 Then
            lng������ID = rsTemp!���ID
        Else
            lng�������ID = rsTemp!���ID
        End If
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    If lng������ID = 0 Then
        ShowMsgBox "û���������������̵����������������������������!"
        Exit Function
    End If
    If lng�������ID = 0 Then
        ShowMsgBox "û���������������̵����ĳ�����������������������!"
        Exit Function
    End If
    
    Set cllSQL = New Collection
    With mshBill
        lng�ⷿid = txtStock.Tag
        
        chrNo = Trim(txtNO)
        If mint�༭״̬ = 1 Or mint�༭״̬ = 5 Or mint�༭״̬ = 6 Then 'mbln�������� Or
            If chrNo <> "" Then
                If CheckNOExists(75, chrNo) Then Exit Function
            End If
            If chrNo = "" Then chrNo = sys.GetNextNo(75, lng�ⷿid)
            If IsNull(chrNo) Then Exit Function
        End If
        
        txtNO.Tag = chrNo
        
        strժҪ = Trim(txtժҪ.Text)
        str������ = Txt������
        dat�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        str�̵�ʱ�� = txtCheckDate.Caption
        
        If mint�༭״̬ = 2 Then        '�޸�
            gstrSQL = "zl_�����̵�_Delete('" & mstr���ݺ� & "')"
            AddArray cllSQL, gstrSQL
        End If
        
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            lngRow = recSort!�к�
'        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, 0) <> "" Then
                lng����ID = .TextMatrix(lngRow, 0)
                str���� = .TextMatrix(lngRow, mBillCol.C_����)
                str���� = .TextMatrix(lngRow, mBillCol.c_����)
                lng����ID = IIf(.TextMatrix(lngRow, mBillCol.C_����) = "", 0, .TextMatrix(lngRow, mBillCol.C_����))
    
                int������ = 0
                If Val(.TextMatrix(lngRow, mBillCol.C_����)) = -1 Or Val(.TextMatrix(lngRow, mBillCol.c_������)) = 1 Then
                    int������ = 1
                End If
                
                datЧ�� = IIf(.TextMatrix(lngRow, mBillCol.C_Ч��) = "", "", .TextMatrix(lngRow, mBillCol.C_Ч��))
                datЧ�� = IIf(.TextMatrix(lngRow, mBillCol.C_Ч��) = "", "", .TextMatrix(lngRow, mBillCol.C_Ч��))
                
                dbl�������� = Round(Val(.TextMatrix(lngRow, mBillCol.C_��������)) * Val(.TextMatrix(lngRow, mBillCol.c_����ϵ��)), g_С��λ��.obj_���С��.����С��)
                dblʵ������ = Round(Val(.TextMatrix(lngRow, mBillCol.C_ʵ������)) * Val(.TextMatrix(lngRow, mBillCol.c_����ϵ��)), g_С��λ��.obj_���С��.����С��)
                dbl������ = Round(Val(.TextMatrix(lngRow, mBillCol.C_������)) * Val(.TextMatrix(lngRow, mBillCol.c_����ϵ��)), g_С��λ��.obj_���С��.����С��)
                dbl�ɱ��� = Round(.TextMatrix(lngRow, mBillCol.C_�ɱ���) / Val(.TextMatrix(lngRow, mBillCol.c_����ϵ��)), g_С��λ��.obj_���С��.�ɱ���С��)
                dbl�ۼ� = Round(Val(.TextMatrix(lngRow, mBillCol.C_�ۼ�)) / Val(.TextMatrix(lngRow, mBillCol.c_����ϵ��)), g_С��λ��.obj_���С��.���ۼ�С��)
                
                If dblʵ������ = 0 Then
                    dbl���� = Round(Val(.TextMatrix(lngRow, mBillCol.c_����)), g_С��λ��.obj_���С��.���С��)
                    dbl��۲� = Round(Val(.TextMatrix(lngRow, mBillCol.c_��۲�)), g_С��λ��.obj_���С��.���С��)
                    dbl����� = Round(Val(.TextMatrix(lngRow, mBillCol.C_ʵ�ʽ��)), g_С��λ��.obj_���С��.���С��)
                    dbl����� = Round(Val(.TextMatrix(lngRow, mBillCol.C_ʵ�ʲ��)), g_С��λ��.obj_���С��.���С��)
                Else
                    dbl���� = Round(Val(.TextMatrix(lngRow, mBillCol.c_����)), g_С��λ��.obj_���С��.���С��)
                    dbl��۲� = Round(Val(.TextMatrix(lngRow, mBillCol.c_��۲�)), g_С��λ��.obj_���С��.���С��)
                    dbl����� = Round(Val(.TextMatrix(lngRow, mBillCol.C_ʵ�ʽ��)), g_С��λ��.obj_���С��.���С��)
                    dbl����� = Round(Val(.TextMatrix(lngRow, mBillCol.C_ʵ�ʲ��)), g_С��λ��.obj_���С��.���С��)
                End If
                If dbl�������� <= dblʵ������ Then
                    lng������ID = lng������ID
                    int���ϵ�� = 1
                Else
                    lng������ID = lng�������ID
                    int���ϵ�� = -1
                End If
                 
                lng��� = lngRow
                'zl_�����̵�_INSERT
                '    No_In         In ҩƷ�շ���¼.NO%Type,
                '    ���_In       In ҩƷ�շ���¼.���%Type,
                '    �ⷿid_In     In ҩƷ�շ���¼.�ⷿid%Type,
                '    ����_In       In ҩƷ�շ���¼.����%Type,
                '    ������id_In In ҩƷ�շ���¼.������id%Type,
                '    ���ϵ��_In   In ҩƷ�շ���¼.���ϵ��%Type,
                '    ����id_In     In ҩƷ�շ���¼.ҩƷid%Type,
                '    ��������_In   In ҩƷ�շ���¼.��д����%Type,
                '    ʵ������_In   In ҩƷ�շ���¼.����%Type,
                '    ������_In     In ҩƷ�շ���¼.ʵ������%Type,
                '    �ɱ���_In     In ҩƷ�շ���¼.����%Type,
                '    �ۼ�_In       In ҩƷ�շ���¼.���ۼ�%Type,
                '    ����_In     In ҩƷ�շ���¼.���۽��%Type,
                '    ��۲�_In     In ҩƷ�շ���¼.���%Type,
                '    ������_In     In ҩƷ�շ���¼.������%Type,
                '    ��������_In   In ҩƷ�շ���¼.��������%Type,
                '    ժҪ_In       In ҩƷ�շ���¼.ժҪ%Type := Null,
                '    ����_In       In ҩƷ�շ���¼.����%Type := Null,
                '    ����_In       In ҩƷ�շ���¼.����%Type := Null,
                '    Ч��_In       In ҩƷ�շ���¼.Ч��%Type := Null,
                '    �̵�ʱ��_In   In ҩƷ�շ���¼.Ƶ��%Type := Null,
                '    �����_In   In ҩƷ�շ���¼.�ɱ���%Type := Null,
                '    �����_In   In ҩƷ�շ���¼.�ɱ����%Type := Null
                '    ������_In     In Number := 0
                gstrSQL = "zl_�����̵�_INSERT('" & _
                    chrNo & "'," & _
                    lng��� & "," & _
                    lng�ⷿid & "," & _
                    lng����ID & "," & _
                    lng������ID & "," & _
                    int���ϵ�� & "," & _
                    lng����ID & "," & _
                    dbl�������� & "," & _
                    dblʵ������ & "," & _
                    dbl������ & "," & _
                    dbl�ɱ��� & "," & _
                    dbl�ۼ� & "," & _
                    dbl���� & "," & _
                    dbl��۲� & ",'" & _
                    str������ & "',to_date('" & _
                    dat�������� & "','yyyy-mm-dd HH24:MI:SS'),'" & _
                    strժҪ & "','" & _
                    str���� & "','" & _
                    str���� & "'," & _
                    IIf(datЧ�� = "", "Null", "to_date('" & Format(datЧ��, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ",'" & _
                    str�̵�ʱ�� & "'," & _
                    dbl����� & "," & _
                    dbl����� & "," & _
                    int������ & ")"
                AddArray cllSQL, gstrSQL
            End If
            
            recSort.MoveNext
        Next
        
        If mint�༭״̬ = 5 Then
            '���˺�:20060801
            'ɾ��������̴�����е��̵��¼��
            strArr = Split(mstr�̵㵥��, ",")
            
            For i = 0 To UBound(strArr)
                
                If mblnɾ���̵㵥 Then
                    'Zl_�����̵��¼��_DELETE:
                    '   NO_IN
                    gstrSQL = "Zl_�����̵��¼��_DELETE(" & strArr(i) & ")"
                Else
                    'Zl_�����̵��¼��_Update:
                    '   NO_IN
                    gstrSQL = "Zl_�����̵��¼��_Update(" & strArr(i) & ")"
                End If
                AddArray cllSQL, gstrSQL
            Next
        End If
        
    End With
        
    'ִ�����SQL
    Call ExecuteProcedureArrAy(cllSQL, mstrCaption)
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    SaveCard = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ��ʾ�ϼƽ��()
    Dim dbl���� As Double
    Dim dbl�̵��� As Double
    Dim dbl�ɱ��̵��� As Double
    Dim lngLop As Long
    
    dbl���� = 0
    dbl�̵��� = 0
    dbl�ɱ��̵��� = 0
    
    With mshBill
        For lngLop = 1 To .Rows - 1
            If .TextMatrix(lngLop, 0) <> "" Then
                
                dbl���� = dbl���� + Val(.TextMatrix(lngLop, mBillCol.c_����)) * IIf(.TextMatrix(lngLop, mBillCol.C_��־) = "��", -1, 1)
                dbl�̵��� = dbl�̵��� + Val(.TextMatrix(lngLop, mBillCol.C_�̵���))
                dbl�ɱ��̵��� = dbl�ɱ��̵��� + Val(.TextMatrix(lngLop, mBillCol.C_�̵�ɱ����))
            End If
        Next
    End With
    
    lblPurchasePrice.Caption = "����ϼƣ�" & Format(dbl����, mFMT.FM_���)
    lblPurchasePrice.Width = Pic����.TextWidth(lblPurchasePrice.Caption)
    lblCheckSum.Left = lblPurchasePrice.Left + lblPurchasePrice.Width + 200
    
    lblCheckSum.Caption = "�̵���ϼƣ�" & Format(dbl�̵���, mFMT.FM_���)
    lblCheckSum.Width = Pic����.TextWidth(lblCheckSum.Caption)
    
    lblCheckCostSum.Top = lblCheckSum.Top
    lblCheckCostSum.Left = lblCheckSum.Left + lblCheckSum.Width + 200
    lblCheckCostSum.Caption = "�̵�ɱ����ϼƣ�" & Format(dbl�ɱ��̵���, mFMT.FM_���)
    lblCheckCostSum.Width = Pic����.TextWidth(lblCheckCostSum.Caption)
    
End Sub

Private Sub ��ʾ�����()
    Dim rsTemp As New Recordset
    Dim strKc As String
    
    On Error GoTo errHandle
    'ȡ���
    '20060731:���˺���룬��Ҫ����̵�ʱ��Ŀ��
    strKc = "" & _
        "   SELECT " & _
        "           nvl(a.��������,0)/[5] ��������,nvl(a.ʵ������,0)/[5] ʵ������,a.ʵ�ʽ��, a.ʵ�ʲ��" & _
        "   FROM ҩƷ��� a" & _
        "   Where a.ҩƷid=[2] and nvl(a.����,0)=[3] " & _
        "           AND a.����=1 " & _
        "           AND a.�ⷿid =[1] "
           
    
    With mshBill
        If .TextMatrix(.Row, mBillCol.C_����) = "" Then
            stbThis.Panels(2).Text = ""
            Exit Sub
        End If
        If .TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
        
       ' gstrSQL = "" & _
            "   Select ��������/" & .TextMatrix(.Row, mBillCol.C_����ϵ��) & " as  �������� " & _
            "   From ҩƷ��� where �ⷿid=[1]" & _
            "       and ҩƷid=[2]" & _
            "       and ����=1 " & _
            "       and  nvl(����,0)=[3]"
        gstrSQL = strKc
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ʾ�����", Val(txtStock.Tag), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mBillCol.C_����)), CDate(txtCheckDate.Caption), Val(.TextMatrix(.Row, mBillCol.c_����ϵ��)))
        
        If rsTemp.EOF Then
            .TextMatrix(.Row, mBillCol.C_��������) = 0
        Else
            .TextMatrix(.Row, mBillCol.C_��������) = IIf(IsNull(rsTemp.Fields(1)), 0, rsTemp.Fields(1))
        End If
        rsTemp.Close
        
        stbThis.Panels(2).Text = "���������ϵ�ǰ�����Ϊ[" & Format(.TextMatrix(.Row, mBillCol.C_��������), mFMT.FM_����) & "]" & .TextMatrix(.Row, mBillCol.c_��λ)
    End With
    Exit Sub
errHandle:
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

Private Function SetPhiscRows(ByVal lngId As Long, ByVal lng���� As Long) As Boolean
'���ܣ����ݲ���ID���̴������ʾ������ò��ϵĳ�ʼ�̴���Ϣ
'˵����
'   1.����Ƿǿⷿ����ҩ,���Ѿ�������,����ʾ���˳���
'   2.����ǿⷿ����ҩ����ֱ����ҩ��δ����ĸ����ο���С�
    Dim i As Integer
    Dim rsData As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim blnModi As Boolean, sngLevel As Single
    Dim lngRecordCount As Long
    Dim lngRow As Long
    Dim bln�ⷿ As Boolean
    Dim dbl�ɱ��� As Double
    Dim dblָ������� As Double
    Dim lngBatch As Long
    Dim rsprice As New Recordset
    Dim lngTmp As Long
    Dim dbl����, dbl��۲� As Double
    Dim strMoneyDigit As String
    
    On Error GoTo errH
    
    SetPhiscRows = False
    Set rsData = GetDateStock(txtCheckDate.Caption, txtStock.Tag, 0, True, , , lngId)
    lngRecordCount = rsData.RecordCount
    If lngRecordCount = 0 Then Exit Function
    
    bln�ⷿ = CheckPartProp(Val(txtStock.Tag))
    '��������ҩƷ
    If lng���� <> -1 Then
        rsData.MoveFirst
        rsData.Find "����=" & lng����
        If rsData.EOF Then Exit Function
    End If
    
    With mshBill
        '��鵥���Ƿ���ڶ�Ӧ����
        If lng���� <> -1 Then
            For lngRow = 1 To .Rows - 1
                If .TextMatrix(lngRow, 0) <> "" Then
                    If .TextMatrix(lngRow, 0) = rsData!����ID And IIf(.TextMatrix(lngRow, mBillCol.C_����) = "", "0", .TextMatrix(lngRow, mBillCol.C_����)) = lng���� Then
                        If UBound(Split(mstr�ظ�����, "��")) < 3 Then mstr�ظ����� = mstr�ظ����� & .TextMatrix(lngRow, mBillCol.C_����) & "��"  '����¼�����ظ�������
    '                    MsgBox "�����������ϡ�" & .TextMatrix(lngRow, mBillCol.C_����) & "(" & lng���� & ")����������ӣ�", vbOKOnly, gstrSysName
                        Exit Function
                    End If
                End If
            Next
        End If
        
        mshBill.Redraw = False
        lngRow = .Row
        .TextMatrix(lngRow, 0) = rsData!����ID
        
        'ȡ���ò��ϵĳɱ���
        'gstrSQL = "Select Nvl(�ɱ���,0) �ɱ���,nvl(ָ�������,0) From �������� Where ����ID=[1]"
        'Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�����������ϵĳɱ���", Val(NVL(rsData!����ID)))
                
        dbl�ɱ��� = Val(zlStr.NVL(rsData!������))
        dblָ������� = Val(zlStr.NVL(rsData!ָ�������))
            
        .TextMatrix(lngRow, mBillCol.C_����) = "[" & rsData!���� & "]" & rsData!��Ʒ����
        .TextMatrix(lngRow, mBillCol.c_���) = IIf(IsNull(rsData!���), "", rsData!���)
        .TextMatrix(lngRow, mBillCol.C_����) = IIf(IsNull(rsData!����), "", rsData!����)
        .TextMatrix(lngRow, mBillCol.C_��׼�ĺ�) = IIf(IsNull(rsData!��׼�ĺ�), "", rsData!��׼�ĺ�)
        .TextMatrix(lngRow, mBillCol.C_�ⷿ��λ) = IIf(IsNull(rsData!�ⷿ��λ), "", rsData!�ⷿ��λ)
        .TextMatrix(lngRow, mBillCol.c_��λ) = IIf(IsNull(rsData!��λ), "", rsData!��λ)
        .TextMatrix(lngRow, mBillCol.C_����) = IIf(IsNull(rsData!����), "0", rsData!����)
        
        If Val(.TextMatrix(lngRow, mBillCol.C_����)) <> 0 Then
            .TextMatrix(lngRow, mBillCol.c_���ű༭) = rsData!���ű༭
            .TextMatrix(lngRow, mBillCol.c_���ر༭) = rsData!���ر༭
        End If
            
        If CheckPhysicBatch(bln�ⷿ, rsData!�ⷿ����, rsData!���÷���) And Val(.TextMatrix(lngRow, mBillCol.C_����)) = 0 Then
            .TextMatrix(lngRow, mBillCol.C_����) = -1
        End If
        
        If lng���� = -1 Then
            .TextMatrix(lngRow, mBillCol.C_����) = lng����
            .TextMatrix(lngRow, mBillCol.c_����) = ""
            .TextMatrix(lngRow, mBillCol.C_Ч��) = ""
            .TextMatrix(lngRow, mBillCol.C_��������) = Format(0, mFMT.FM_����)
            .TextMatrix(lngRow, mBillCol.C_ʵ������) = .TextMatrix(lngRow, mBillCol.C_��������)
            .TextMatrix(lngRow, mBillCol.C_�̵���) = Format(0, mFMT.FM_���)
            .TextMatrix(lngRow, mBillCol.C_��������) = 0
            .TextMatrix(lngRow, mBillCol.C_ʵ�ʽ��) = 0
            .TextMatrix(lngRow, mBillCol.C_ʵ�ʲ��) = 0
            .TextMatrix(lngRow, mBillCol.C_�ۼ�) = Format(IIf(IsNull(rsData!�ۼ�), 0, rsData!�ۼ�), mFMT.FM_���ۼ�)
            .TextMatrix(lngRow, mBillCol.C_�ɱ���) = Format(dbl�ɱ���, mFMT.FM_�ɱ���)
            .ColData(mBillCol.c_����) = 4
            .ColData(mBillCol.C_Ч��) = 2
        Else
            lngBatch = Val(.TextMatrix(lngRow, mBillCol.C_����))
            .ColData(mBillCol.c_����) = IIf(lngBatch = -1, 4, 5)
            .ColData(mBillCol.C_Ч��) = IIf(lngBatch = -1, 2, 5)
            
            .TextMatrix(lngRow, mBillCol.c_����) = IIf(IsNull(rsData!����), "", rsData!����)
            .TextMatrix(lngRow, mBillCol.C_Ч��) = IIf(IsNull(rsData!Ч��), "", Format(rsData!Ч��, "yyyy-MM-dd"))
            .TextMatrix(lngRow, mBillCol.C_��������) = Format(IIf(IsNull(rsData!��������), 0, rsData!��������), mFMT.FM_����)
            .TextMatrix(lngRow, mBillCol.C_ʵ������) = .TextMatrix(lngRow, mBillCol.C_��������)
            .TextMatrix(lngRow, mBillCol.C_�ۼ�) = Format(IIf(IsNull(rsData!�ۼ�), 0, rsData!�ۼ�), mFMT.FM_���ۼ�)
            .TextMatrix(lngRow, mBillCol.C_�ɱ���) = Format(Val(zlStr.NVL(rsData!�ɱ���)), mFMT.FM_�ɱ���)
            .TextMatrix(lngRow, mBillCol.C_�̵���) = Format(Val(.TextMatrix(lngRow, mBillCol.C_ʵ������)) * Val(.TextMatrix(lngRow, mBillCol.C_�ۼ�)), mFMT.FM_���)
            
            .TextMatrix(lngRow, mBillCol.C_��������) = rsData!��������
            .TextMatrix(lngRow, mBillCol.C_ʵ�ʽ��) = rsData!ʵ�ʽ��
            .TextMatrix(lngRow, mBillCol.C_ʵ�ʲ��) = rsData!ʵ�ʲ��
            .TextMatrix(lngRow, mBillCol.C_�ɱ���) = Format(Val(zlStr.NVL(rsData!�ɱ���)), mFMT.FM_�ɱ���)
        End If
        
        .TextMatrix(lngRow, mBillCol.c_����ϵ��) = rsData!����ϵ��
        .TextMatrix(lngRow, mBillCol.C_ָ�������) = rsData!ָ������� & "||" & rsData!�Ƿ��� & "||" & rsData!���÷���
        
        .TextMatrix(lngRow, mBillCol.C_��־) = "ƽ"
        .TextMatrix(lngRow, mBillCol.C_������) = Format("0", mFMT.FM_���)
            
        If rsData!�Ƿ��� = 1 Then
            .TextMatrix(lngRow, mBillCol.C_�ۼ�) = Format(Get���ۼ�(Val(zlStr.NVL(rsData!����ID)), Val(txtStock.Tag), Val(zlStr.NVL(rsData!����)), rsData!����ϵ��), mFMT.FM_���ۼ�)
        End If
        
        .RowData(lngRow) = IIf(IsNull(rsData!���Ч��), 0, rsData!���Ч��)
        
        If Val(.TextMatrix(lngRow, mBillCol.C_��������)) = 0 Then
            strMoneyDigit = "#0.00000"
        Else
            strMoneyDigit = mFMT.FM_���
        End If
                    
        '����=��ǰ�ۼ�*ʵ������-ʵ�ʽ��
        '��۲�=����*iif(ʵ�ʽ��=0,ָ�������,(ʵ�ʲ��/ʵ�ʽ��))
        .TextMatrix(lngRow, mBillCol.c_����) = Format(Val(.TextMatrix(lngRow, mBillCol.C_�ۼ�)) * Val(.TextMatrix(lngRow, mBillCol.C_ʵ������)) - Val(.TextMatrix(lngRow, mBillCol.C_ʵ�ʽ��)), strMoneyDigit)
            
        If rsData!�Ƿ��� = 1 And Val(.TextMatrix(lngRow, mBillCol.C_��������)) = 0 Then
            .TextMatrix(lngRow, mBillCol.c_��۲�) = Format(Val(.TextMatrix(lngRow, mBillCol.C_������)) * (Val(.TextMatrix(lngRow, mBillCol.C_�ۼ�)) - dbl�ɱ���) - Val(.TextMatrix(lngRow, mBillCol.C_ʵ�ʲ��)), strMoneyDigit)
        Else
            .TextMatrix(lngRow, mBillCol.c_��۲�) = Format(Val(.TextMatrix(lngRow, mBillCol.C_ʵ������)) * (Val(.TextMatrix(lngRow, mBillCol.C_�ۼ�)) - Val(.TextMatrix(lngRow, mBillCol.C_�ɱ���))) - Val(.TextMatrix(lngRow, mBillCol.C_ʵ�ʲ��)), strMoneyDigit)
        End If
        
        dbl���� = .TextMatrix(lngRow, mBillCol.c_����)
        dbl��۲� = .TextMatrix(lngRow, mBillCol.c_��۲�)
        
        .TextMatrix(lngRow, mBillCol.C_�̵�ɱ����) = Format(Val(.TextMatrix(lngRow, mBillCol.C_ʵ�ʽ��)) + dbl���� - (Val(.TextMatrix(lngRow, mBillCol.C_ʵ�ʲ��)) + dbl��۲�), mFMT.FM_���)
        .TextMatrix(lngRow, mBillCol.C_�̵�ɱ�����) = Format(Val(.TextMatrix(lngRow, mBillCol.c_����)) - Val(.TextMatrix(lngRow, mBillCol.c_��۲�)), mFMT.FM_���)
        
        Call RefreshRowNO(mshBill, mBillCol.C_�к�, 1)
        mshBill.Redraw = True
    End With
    Call ��ʾ�����
    rsData.Close
    SetPhiscRows = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'��һ���в���
Private Sub InsertRow(ByVal intRow As Integer, ByVal intRecordCount As Integer)
    Dim blnHaveData As Boolean
    Dim lngOldRows As Long
    Dim lngLop As Long
    Dim lngExchange As Long
    Dim intCol As Integer
    
    With mshBill
        blnHaveData = False
        lngOldRows = .Rows - 1
        .Rows = .Rows + intRecordCount
        For lngLop = intRow + 1 To intRecordCount
            If .TextMatrix(lngLop, 0) <> "" Then
                blnHaveData = True
                Exit For
            End If
        Next
        If blnHaveData = True Then
            For lngExchange = .Rows - 1 To lngOldRows Step -1
                For intCol = 0 To .Cols - 1
                    .TextMatrix(lngExchange, intCol) = .TextMatrix(lngExchange - intRecordCount, intCol)
                    .TextMatrix(lngExchange - intRecordCount, intCol) = ""
                Next
            Next
        End If
    End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

'��ӡ����
Private Sub printbill()
    Dim strNo As String
    strNo = txtNO.Tag
    Call FrmBillPrint.ShowMe(Me, glngSys, "zl1_bill_1719", mint��¼״̬, mintUnit, 1719, "���������̵��", strNo)
End Sub

Private Function CheckPartProp(ByVal lng�ⷿid As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    '���ⷿ���ԣ�����ǿⷿ��������
    gstrSQL = "" & _
        "   SELECT count(*)" & _
        "   From ��������˵�� " & _
        "   WHERE ((�������� LIKE '���ϲ���') OR (�������� LIKE '�Ƽ���')) " & _
        "           AND ����id =[1]"
        
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng�ⷿid)
    
    If rsTemp.Fields(0) > 0 Then
        CheckPartProp = False
    Else
        CheckPartProp = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckPhysicBatch(ByVal bln�ⷿ As Boolean, ByVal int�ⷿ���� As Integer, ByVal int���÷��� As Integer) As Boolean
    '���ظò����Ƿ�����ı�ʶ
    CheckPhysicBatch = (bln�ⷿ And (int�ⷿ���� = 1)) Or (Not bln�ⷿ And (int���÷��� = 1))
End Function

'ȡ���ݿ������ŵĳ��ȣ������������е����ų��������ݿ��б���һ����
Private Function GetBatchNoLen() As Integer
    Dim rsBatchNolen As New Recordset
    
    gstrSQL = "select ���� from ҩƷ�շ���¼ where rownum<1 "
    Call zlDatabase.OpenRecordset(rsBatchNolen, gstrSQL, "ȡ�ֶγ���")
    GetBatchNoLen = rsBatchNolen.Fields(0).DefinedSize
    rsBatchNolen.Close
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
                !���� = Val(mshBill.TextMatrix(n, mBillCol.C_����))
                
                .Update
            End If
        Next
        
    End With
End Sub
