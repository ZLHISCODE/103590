VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCheckCourseCard 
   Caption         =   "�����̵��¼��"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmCheckCourseCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11400
   StartUpPosition =   2  '��Ļ����
   Begin MSMask.MaskEdBox TxtCheckDate 
      Height          =   315
      Left            =   9510
      TabIndex        =   6
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   19
      Format          =   "yyyy-MM-dd HH:mm:ss"
      Mask            =   "####-##-## ##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   25
      Top             =   5137
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   23
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   22
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   20
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   21
      Top             =   5040
      Width           =   1100
   End
   Begin VB.PictureBox Pic���� 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11715
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   9945
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   135
         Width           =   1425
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   7
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
         TabIndex        =   11
         Top             =   4080
         Width           =   10410
      End
      Begin VB.Label lblCheckSum 
         AutoSize        =   -1  'True
         Caption         =   "�̵���ϼƣ�"
         Height          =   180
         Left            =   1920
         TabIndex        =   9
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
         TabIndex        =   5
         Top             =   660
         Width           =   720
      End
      Begin VB.Label txtStock 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1080
         TabIndex        =   4
         Top             =   600
         Width           =   1845
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "�̵�ɱ����ϼƣ�"
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   3840
         Width           =   1620
      End
      Begin VB.Label Txt����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7950
         TabIndex        =   17
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Txt������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10050
         TabIndex        =   19
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt�������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   15
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt������ 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   13
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
         TabIndex        =   2
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
         TabIndex        =   10
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "�����̵��¼��"
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
         TabIndex        =   1
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
         TabIndex        =   3
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
         TabIndex        =   12
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
         TabIndex        =   14
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
         TabIndex        =   18
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
            Picture         =   "frmCheckCourseCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":1000
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
            Picture         =   "frmCheckCourseCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCourseCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   26
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
            Picture         =   "frmCheckCourseCard.frx":22EA
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
            Picture         =   "frmCheckCourseCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmCheckCourseCard.frx":3080
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
      TabIndex        =   24
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmCheckCourseCard"
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
Private mintParallelRecord As Integer       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼
Private mintBatchNoLen As Integer           '���ݿ������Ŷ��峤��
Private mintDefault As Integer              'ȱʡ��λ
Private mint����� As Integer             '��ʾ�������ϳ���ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ
Dim mstrPrivs As String                     'Ȩ��
Private mbln���޴洢�ⷿ���� As Boolean
Private Const mstrCaption As String = "�����̵��¼��"
Private mstr�ظ����� As String '��¼�ظ�������
Private mbln�����������Ų��ؿ��� As Boolean  '�Ƿ�������������Ų����Ƿ�¼��


'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
Private Const mlngModule = 1719

Private mbln��������    As Boolean          '����ʱ���ݺ��ۼ�1
Private mintUnit  As Integer                '��ʾ��λ:0-ɢװ��λ,1-��װ��λ

'=========================================================================================
Private Const mconIntCol�к� As Integer = 1
Private Const mconIntCol���� As Integer = 2
Private Const mconIntCol��� As Integer = 3
Private Const mconIntCol��� As Integer = 4
Private Const mconIntCol���� As Integer = 5
Private Const mconIntCol�������� As Integer = 6
Private Const mconIntCol����ϵ�� As Integer = 7
Private Const mconIntColָ������� As Integer = 8
Private Const mconIntColʵ�ʲ�� As Integer = 9
Private Const mconIntColʵ�ʽ�� As Integer = 10
Private Const mconIntCol���� As Integer = 11
Private Const mconIntCol�ⷿ��λ As Integer = 12
Private Const mconIntCol��λ As Integer = 13
Private Const mconIntCol���� As Integer = 14
Private Const mconIntColЧ�� As Integer = 15
Private Const mconIntCol���Ч�� As Integer = 16
Private Const mconintCol�������� As Integer = 17
Private Const mconintColʵ������ As Integer = 18
Private Const mconintCol��־ As Integer = 19
Private Const mconintCol������ As Integer = 20
Private Const mconIntCol�ɱ��� As Integer = 21
Private Const mconIntCol�ɱ���� As Integer = 22
Private Const mconIntCol�ۼ� As Integer = 23
Private Const mconIntCol�ۼ۽�� As Integer = 24
Private Const mconintCol���� As Integer = 25
Private Const mconintCol��۲� As Integer = 26
Private Const mconintCol�̵��� As Integer = 27
Private Const mconintCol���ű༭ As Integer = 28
Private Const mconintCol���ر༭ As Integer = 29
Private Const mconIntColS  As Integer = 30             '������
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
        "           AND A.���� = 39  and b.ϵ��=1 "
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "���������̵����"
    If rsTemp.EOF Then
        ShowMsgBox "û���������������̵��¼��������������������������ã�"
        rsTemp.Close
        Exit Function
    End If
    rsTemp.Close
    GetDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub ShowCard(FrmMain As Form, ByVal str���ݺ� As String, ByVal int�༭״̬ As Integer, _
                Optional int��¼״̬ As Integer = 1, Optional strPrivs As String, Optional blnSuccess As Boolean = False)
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ʾ��༭����
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

    Call GetRegInFor(g˽��ģ��, "�����̵����", "���ݺ��ۼ�", strReg)
    mbln�������� = IIf(strReg = "", True, Val(strReg) = 1)

 

    
    If mint�༭״̬ = 1 Then
'        If mbln�������� Then
'            mstr���ݺ� = NextNo(76)
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
        CmdSave.Visible = False
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
        FindRownew mshBill, mconIntCol����, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
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
    
    If mint�༭״̬ = 4 Then    '�鿴
        '��ӡ
        printbill
        '�˳�
        Unload Me
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
'        mstr���ݺ� = NextNo(76)
'        txtNO = mstr���ݺ�
'    End If
    
    mblnSave = False
    mblnEdit = True
    mshBill.ClearBill
    Call RefreshRowNO(mshBill, mconIntCol�к�, 1)
    txtժҪ.Text = ""
    mblnChange = False
    If txtNO.Tag <> "" Then Me.stbThis.Panels(2).Text = "��һ�ŵ��ݵ�NO�ţ�" & txtNO.Tag
End Sub

Private Sub Form_Activate()
    
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    If mint�༭״̬ = 1 Then
        mshBill.ClearBill
        Call RefreshRowNO(mshBill, mconIntCol�к�, 1)
    Else
'        mblnChange = False
        Select Case mintParallelRecord
            Case 1
                '����
            Case 2
                '�����ѱ�ɾ��
                ShowMsgBox "�õ����ѱ�ɾ�������飡"
                Unload Me
                Exit Sub
            Case 3
                '�޸ĵĵ����ѱ����
                ShowMsgBox "�õ����ѱ���������ˣ����飡"
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

Private Function GetDateStock(str�̴�ʱ�� As String, lng�ⷿid As Long, str�������� As String, Optional blnZero As Boolean = False, Optional ByVal bln���� As Boolean = False, Optional lng����ID As Long = 0) As ADODB.Recordset
    '���ܣ���ȡָ������������ָ��ʱ���Ŀ�漰�����Ϣ
    '������str�̴�ʱ��=Ҫ����YYYY-MM-DD HH24:MI:SSΪ��ʽ��ʱ���ַ���
    '      str��������=" And B.����ID=... And ..."
    '      blnZero=�Ƿ��ȡ��������Ϊ0�Ĳ���,ȱʡ��.��ǿ������ò���ʱ,����Ϊ�ǡ�
    Dim rsTemp As New ADODB.Recordset
    Dim strUnitQuantity As String
    Dim blnStock As Boolean
    Dim strOrder As String, strCompare As String
    
    On Error GoTo errH
    strOrder = zlDatabase.GetPara("��������", glngSys, mlngModule, "00")
    strOrder = IIf(strOrder = "", "00", strOrder)
    strCompare = Mid(strOrder, 1, 1)
    
    gstrSQL = "" & _
        "   SELECT count(*)" & _
        "   From ��������˵�� " & _
        "   WHERE ((�������� LIKE '���ϲ���') OR (�������� LIKE '�Ƽ���')) " & _
        "       AND ����id =[1]"
        
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng�ⷿid)
    
    If rsTemp.Fields(0) > 0 Then
        blnStock = False
    Else
        blnStock = True
    End If
    
    If lng����ID <> 0 Then
        str�������� = " And B.����ID=[3]"
    End If
    'ȡ�õ�ǰ���
    gstrSQL = "" & _
        "   SELECT a.�ⷿid, b.����id, NVL (����, 0) AS ����, a.ʵ������,a.ʵ�ʽ��, a.ʵ�ʲ��, a.��������,a.�ϴ����� AS ����,a.�ϴβ��� AS ����,a.Ч��,a.ƽ���ɱ��� " & _
        "   FROM (Select �ⷿid,ҩƷid,����,ʵ������,ʵ�ʽ��,ʵ�ʲ��,��������,�ϴ�����,�ϴβ���,Ч��,ƽ���ɱ��� From ҩƷ��� Where ����=1 And �ⷿID+0=[1]) a, �������� b,(Select �ⷿid, ����id, ����, ����, �̵�����, �ⷿ��λ From ���ϴ����޶� Where �ⷿID+0=[1] )e " & _
        "   Where a.ҩƷid(+) = b.����id " & _
        "           and b.����id=e.����id(+) " & str��������
    'ȡ���̵�ʱ���ľ�������
    gstrSQL = gstrSQL & _
        "   UNION ALL " & _
        "       SELECT a.�ⷿid, b.����id, NVL (a.����, 0) AS ����, " & _
        "               -SUM (DECODE (a.���ϵ��, 1, a.ʵ������*a.����, -a.ʵ������*a.����)) AS ʵ������, " & _
        "               -SUM (DECODE (a.���ϵ��, 1, a.���۽��, -a.���۽��)) AS ʵ�ʽ��," & _
        "               -SUM (DECODE (a.���ϵ��, 1, a.���, -a.���)) AS ʵ�ʲ��,0 AS ��������,a.����,a.����,a.Ч��,null ƽ���ɱ��� " & _
        "       FROM ҩƷ�շ���¼ a, �������� b,(Select �ⷿid, ����id, ����, ����, �̵�����, �ⷿ��λ From ���ϴ����޶� Where �ⷿID+0=[1] )e " & _
        "       Where a.ҩƷid = b.����id " & _
        "               and b.����id=e.����id(+) " & _
        "               AND a.�ⷿid + 0 =[1]" & _
        "               AND a.������� > [2] " & str�������� & _
        "       GROUP BY a.�ⷿid, b.����id, a.����,a.����,a.����,a.Ч�� "
    
    'ȡ���̵�ʱ����һ�̵���������
    gstrSQL = "" & _
        "   SELECT �ⷿid, ����id, ����, SUM (ʵ������) AS ��������," & _
        "           SUM (ʵ�ʽ��) AS ʵ�ʽ��, SUM (ʵ�ʲ��) AS ʵ�ʲ��, " & _
        "           SUM(��������) As ��������,max(����) as ����,max(����) as ���� ,max(Ч��) as Ч��, Max(ƽ���ɱ���) As ƽ���ɱ��� " & _
        "   FROM ( " & gstrSQL & ") " & _
        "   GROUP BY �ⷿid, ����id, ����,ƽ���ɱ��� "
    
    '(nvl(a.��������,0) / b.סԺ��װ) AS סԺ��������,(nvl(a.��������,0) / b.סԺ��װ) AS סԺ��������,
    If mintUnit = 0 Then
        strUnitQuantity = "c.���㵥λ as ��λ,'1' as �ۼ�ϵ��,f.�ۼ� �ۼ�,"
    Else
        strUnitQuantity = "b.��װ��λ as ��λ,b.����ϵ��,f.�ۼ�*b.����ϵ�� as  �ۼ�,"
    End If
    
    gstrSQL = "" & _
        "   SELECT DISTINCT b.����id, c.����, c.���� AS ��Ʒ����,b.����ϵ��," & _
        "           zlSpellCode(c.����) ����,nvl(b.���Ч��,0) ���Ч��,c.���,Decode(a.����, Null, decode(b.�ϴβ���,null,c.����,b.�ϴβ���), a.����) As ����,e.�ⷿ��λ,a.����, a.����, a.Ч��," & strUnitQuantity & _
        "           nvl(a.ʵ�ʽ��,0) as ʵ�ʽ�� ,nvl(a.ʵ�ʲ��,0) as ʵ�ʲ��, b.ָ�������,c.�Ƿ���,b.�ⷿ����,b.���÷���,decode(a.ƽ���ɱ���,null,b.�ɱ���,a.ƽ���ɱ���) �ɱ���,decode(a.����,null,1,0) ���ű༭,decode(a.����,null,1,0) ���ر༭ " & _
        "   From (" & gstrSQL & ") A ,�������� b,�շ���ĿĿ¼ c,(Select �ⷿid, ����id, ����, ����, �̵�����, �ⷿ��λ From ���ϴ����޶� Where �ⷿID+0=[1] )e, " & _
        "        (  SELECT �շ�ϸĿid, �ּ� as �ۼ� From �շѼ�Ŀ WHERE ((SYSDATE BETWEEN ִ������ AND ��ֹ����) OR (SYSDATE >= ִ������ AND ��ֹ���� IS NULL))" & _
        GetPriceClassString("") & ") f " & _
        "   Where " & IIf(blnZero = False, "a.����id = b.����id and a.����id=c.id ", " b.����id = a.����id(+)and b.����id=c.id ") & _
        "           AND b.����id=f.�շ�ϸĿid " & _
        "           and b.����id=e.����id(+) " & _
            IIf(blnZero, IIf(blnStock, " And (Nvl(b.�ⷿ����,0)=1 ", " And (Nvl(b.���÷���,0)=1 "), "") & _
            IIf(blnZero, IIf(blnStock, " or Nvl(b.�ⷿ����,0)=0)", " Or Nvl(b.���÷���,0)=0)"), "") & _
            IIf(blnZero = False, " AND (a.��������<>0 or nvl(a.ʵ�ʽ��,0)<>0 or nvl(a.ʵ�ʲ��,0)<>0)", str��������) & _
        "   ORDER BY " & IIf(strCompare = "0", "c.����", IIf(strCompare = "1", "c.����", IIf(strCompare = "2", "c.����", "e.�ⷿ��λ"))) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
    
    Screen.MousePointer = 11
        
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng�ⷿid, CDate(str�̴�ʱ��), lng����ID)
     
    Set GetDateStock = rsTemp
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
    
     
    mintUnit = IIf(Val(zlDatabase.GetPara("��¼����λ", glngSys, mlngModule, "0")) = 1, 1, 0)
    mbln�����������Ų��ؿ��� = Val(zlDatabase.GetPara(305, glngSys, 0)) = 1
    
    mblnFirst = True
  
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
    End With
    
    
    mintBatchNoLen = GetBatchNoLen()
    txtNO = mstr���ݺ�
    txtNO.Tag = txtNO.Text
    Call initCard
    RestoreWinState Me, App.ProductName, mstrCaption
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsTemp As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim strOrder As String, strCompare As String
    '�ⷿ
    
    On Error GoTo errHandle
    mbln���޴洢�ⷿ���� = Val(zlDatabase.GetPara("�洢�ⷿ", glngSys, mlngModule, "0"))
    strOrder = zlDatabase.GetPara("��������", glngSys, mlngModule, "00")
    strCompare = Mid(strOrder, 1, 1)
    Select Case mint�༭״̬
        Case 1
            Txt������ = UserInfo.�û���
            Txt�������� = Format(sys.Currentdate, "yyyy-mm-dd HH:mm:ss")
            TxtCheckDate.Text = Txt��������.Caption
            txtStock = mfrmMain.cboStock.Text
            txtStock.Tag = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
            initGrid
        Case 2, 3, 4
            initGrid
            If mint�༭״̬ <> 4 Then
                txtStock = mfrmMain.cboStock.Text
                txtStock.Tag = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
            Else
                gstrSQL = "" & _
                    "   Select distinct b.id,b.���� " & _
                    "   From ҩƷ�շ���¼ a,���ű� b " & _
                    "   where a.�ⷿid=b.id  and A.���� = 23 and a.no=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr���ݺ�)
                
                If rsTemp.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If
                
                txtStock = rsTemp!����
                txtStock.Tag = rsTemp!Id
                rsTemp.Close
            End If
            
            '(Nvl(A.��д����,0)/ B.�����װ) AS ������������,(A.����/ B.�����װ) AS ����ʵ������, (Nvl(A.ʵ������,0) / B.�����װ) AS ����������,
            Select Case mintUnit
            Case 0
                    strUnitQuantity = "A.���� ʵ������,A.ʵ������,A.��д���� ��������,A.ʵ������ ������,d.���㵥λ AS ��λ,'1' as ����ϵ��,a.���ۼ� as �ۼ��ۼ�,"
            Case Else
                    strUnitQuantity = "A.����/b.����ϵ�� as ʵ������,A.��д����/b.����ϵ�� ��������,A.ʵ������/b.����ϵ�� ������,B.��װ��λ AS ��λ,b.����ϵ��,a.���ۼ�*b.����ϵ�� as �ۼ��ۼ�,"
            End Select
            
            gstrSQL = "" & _
                "   Select * " & _
                "   From (  SELECT distinct a.ҩƷid ����id,A.���,('[' || d.���� || ']' || d.����) AS ������Ϣ," & _
                "                   zlSpellCode(d.����) ����,Nvl(B.���Ч��,0) ���Ч��,d.���,A.����,C.�ⷿ��λ, A.����,a.Ч��,a.����," & strUnitQuantity & _
                "                   A.���۽�� as ����,A.��� as ��۲�, " & _
                "                   a.ժҪ,������,��������,�����,�������,a.Ƶ�� as �̵�ʱ��,a.�ɱ��� as �����,a.�ɱ���� as �����,b.ָ�������,D.�Ƿ���,b.���÷���,A.���ۼ�,A.���� As �ɱ���,decode(E.�ϴ�����,null,1,0) ���ű༭,decode(E.�ϴβ���,null,1,0) ���ر༭ " & _
                "           FROM ҩƷ�շ���¼ A, �������� B,�շ���ĿĿ¼ D,���ϴ����޶� C,ҩƷ��� E " & _
                "           Where A.ҩƷid = B.����id and a.ҩƷid=D.id " & _
                "                   And A.ҩƷID=C.����ID(+) And A.�ⷿID=C.�ⷿID(+) AND A.��¼״̬ =[2]" & _
                "                   And A.ҩƷID=E.ҩƷID(+) And A.�ⷿID=E.�ⷿID(+) And nvl(A.����,0) = nvl(E.����(+),0) AND A.���� =23 AND A.No = [1] " & _
                "   ) " & _
                "   ORDER BY " & IIf(strCompare = "0", "���", IIf(strCompare = "1", "������Ϣ", IIf(strCompare = "2", "����", "�ⷿ��λ"))) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr���ݺ�, mint��¼״̬)
            
            If rsTemp.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            Txt������ = rsTemp!������
            If mint�༭״̬ = 2 Then
                Txt������ = UserInfo.�û���
            End If
            Txt�������� = Format(rsTemp!��������, "yyyy-mm-dd HH:mm:ss")
            
            Txt����� = IIf(IsNull(rsTemp!�����), "", rsTemp!�����)
            Txt������� = IIf(IsNull(rsTemp!�������), "", Format(rsTemp!�������, "yyyy-mm-dd HH:mm:ss"))
            txtժҪ.Text = IIf(IsNull(rsTemp!ժҪ), "", rsTemp!ժҪ)
            TxtCheckDate.Text = rsTemp!�̵�ʱ��
            
            If (mint�༭״̬ = 2 Or mint�༭״̬ = 3) And Txt����� <> "" Then
                mintParallelRecord = 3
                Exit Sub
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
                    .TextMatrix(intRow, mconIntCol�ⷿ��λ) = IIf(IsNull(rsTemp!�ⷿ��λ), "", rsTemp!�ⷿ��λ)
                    .TextMatrix(intRow, mconIntCol��λ) = zlStr.NVL(rsTemp!��λ)
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                    .TextMatrix(intRow, mconIntColЧ��) = IIf(IsNull(rsTemp!Ч��), "", Format(rsTemp!Ч��, "yyyy-mm-dd"))
                    .TextMatrix(intRow, mconIntColָ�������) = Format(rsTemp!ָ�������, mFMT.FM_���) & "||" & rsTemp!�Ƿ��� & "||" & rsTemp!���÷���
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsTemp!����), "0", rsTemp!����)
                    .TextMatrix(intRow, mconIntCol����ϵ��) = zlStr.NVL(rsTemp!����ϵ��)
                    .TextMatrix(intRow, mconintColʵ������) = Format(rsTemp.Fields("ʵ������").Value, mFMT.FM_����)
                    
                    If Val(.TextMatrix(intRow, mconIntCol����)) <> 0 Then '��������
                        .TextMatrix(intRow, mconintCol���ű༭) = rsTemp!���ű༭
                        .TextMatrix(intRow, mconintCol���ر༭) = rsTemp!���ر༭
                    End If
                    
                    .TextMatrix(intRow, mconIntCol�ɱ���) = Format(rsTemp!�ɱ��� * IIf(mintUnit = 0, 1, Val(.TextMatrix(intRow, mconIntCol����ϵ��))), mFMT.FM_�ɱ���)
                    .TextMatrix(intRow, mconIntCol�ۼ�) = Format(rsTemp!���ۼ� * IIf(mintUnit = 0, 1, Val(.TextMatrix(intRow, mconIntCol����ϵ��))), mFMT.FM_���ۼ�)
                    .TextMatrix(intRow, mconIntCol�ɱ����) = Format(Val(.TextMatrix(intRow, mconIntCol�ɱ���)) * Val(.TextMatrix(intRow, mconintColʵ������)), mFMT.FM_���)
                    .TextMatrix(intRow, mconIntCol�ۼ۽��) = Format(Val(.TextMatrix(intRow, mconIntCol�ۼ�)) * Val(.TextMatrix(intRow, mconintColʵ������)), mFMT.FM_���)
                    
                    .RowData(intRow) = IIf(IsNull(rsTemp!���Ч��), 0, rsTemp!���Ч��)
                    rsTemp.MoveNext
                Loop
            End With
            rsTemp.Close
    End Select
    Call RefreshRowNO(mshBill, mconIntCol�к�, 1)
    Call ��ʾ�ϼƽ��
    mint����� = Get������(Val(txtStock.Tag))
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
        .Cols = mconIntColS
        .ClearBill
        .MsfObj.FixedCols = 1
        
        .TextMatrix(0, mconIntCol�к�) = ""
        .TextMatrix(0, mconIntCol����) = "���������"
        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntCol�ⷿ��λ) = "�ⷿ��λ"
        .TextMatrix(0, mconIntCol��λ) = "��λ"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntColЧ��) = "ʧЧ��"
        .TextMatrix(0, mconIntCol���Ч��) = "���Ч��"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntCol��������) = "��������"
        .TextMatrix(0, mconIntCol����ϵ��) = "����ϵ��"
        .TextMatrix(0, mconIntColָ�������) = "ָ�������"
        .TextMatrix(0, mconIntColʵ�ʲ��) = "ʵ�ʲ��"
        .TextMatrix(0, mconIntColʵ�ʽ��) = "ʵ�ʽ��"
        .TextMatrix(0, mconintCol��������) = "��������"
        .TextMatrix(0, mconintColʵ������) = "ʵ������"
        .TextMatrix(0, mconintCol��־) = "��־"
        .TextMatrix(0, mconintCol������) = "������"
        .TextMatrix(0, mconIntCol�ɱ���) = "�ɱ���"
        .TextMatrix(0, mconIntCol�ɱ����) = "�ɱ����"
        .TextMatrix(0, mconIntCol�ۼ�) = "�ۼ�"
        .TextMatrix(0, mconIntCol�ۼ۽��) = "�ۼ۽��"
        .TextMatrix(0, mconintCol����) = "����"
        .TextMatrix(0, mconintCol��۲�) = "��۲�"
        .TextMatrix(0, mconintCol�̵���) = "�̵���"
        .TextMatrix(0, mconintCol���ű༭) = "���ű༭"
        .TextMatrix(0, mconintCol���ر༭) = "���ر༭"
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mconIntCol�к�) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mconIntCol�к�) = 300
        .ColWidth(mconIntCol����) = 0
        .ColWidth(mconIntCol���) = 0
        .ColWidth(mconIntCol��������) = 0
        .ColWidth(mconIntCol����ϵ��) = 0
        .ColWidth(mconIntColָ�������) = 0
        .ColWidth(mconIntColʵ�ʲ��) = 0
        .ColWidth(mconIntColʵ�ʽ��) = 0
        
        .ColWidth(mconIntCol����) = 2000
        .ColWidth(mconIntCol���) = 900
        .ColWidth(mconIntCol����) = 800
        .ColWidth(mconIntCol�ⷿ��λ) = 2000
        .ColWidth(mconIntCol��λ) = 0
        .ColWidth(mconIntCol����) = 800
        .ColWidth(mconIntColЧ��) = 1000
        .ColWidth(mconintCol��������) = 0
        .ColWidth(mconintColʵ������) = 1000
        
        .ColWidth(mconintCol��־) = 0
        .ColWidth(mconintCol������) = 0
        .ColWidth(mconIntCol�ɱ���) = 800
        .ColWidth(mconIntCol�ɱ����) = 1200
        .ColWidth(mconIntCol�ۼ�) = 800
        .ColWidth(mconIntCol�ۼ۽��) = 1200
        .ColWidth(mconintCol����) = 0
        .ColWidth(mconintCol��۲�) = 0
        .ColWidth(mconintCol�̵���) = 0
        .ColWidth(mconIntCol���Ч��) = 0
        
        .ColWidth(mconintCol���ű༭) = 0
        .ColWidth(mconintCol���ر༭) = 0
        
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
        .ColData(mconIntCol�ⷿ��λ) = 5
        .ColData(mconIntCol��λ) = 5
        .ColData(mconIntCol����) = 5
        .ColData(mconIntColЧ��) = 5
        .ColData(mconIntCol����) = 5
        .ColData(mconIntCol��������) = 5
        .ColData(mconIntCol����ϵ��) = 5
        .ColData(mconIntColָ�������) = 5
        .ColData(mconIntColʵ�ʲ��) = 5
        .ColData(mconIntColʵ�ʽ��) = 5
        .ColData(mconintCol��������) = 5
        .ColData(mconintCol��־) = 5
        .ColData(mconintCol������) = 5
        .ColData(mconIntCol�ɱ���) = 5
        .ColData(mconIntCol�ɱ����) = 5
        .ColData(mconIntCol�ۼ�) = 5
        .ColData(mconIntCol�ۼ۽��) = 5
        .ColData(mconintCol����) = 5
        .ColData(mconintCol��۲�) = 5
        .ColData(mconintCol�̵���) = 5
        .ColData(mconIntCol���Ч��) = 5
        
        .ColData(mconintCol���ű༭) = 5
        .ColData(mconintCol���ر༭) = 5
        
        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
            txtժҪ.Enabled = True
            .ColData(mconIntCol����) = 1
            .ColData(mconintColʵ������) = 4
        ElseIf mint�༭״̬ = 3 Or mint�༭״̬ = 4 Then
            txtժҪ.Enabled = False
            
            .ColData(mconintColʵ������) = 5
        End If
        
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntCol���) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��λ) = flexAlignCenterCenter
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntColЧ��) = flexAlignLeftCenter
        .ColAlignment(mconintCol��������) = flexAlignRightCenter
        .ColAlignment(mconintCol��־) = flexAlignCenterCenter
        .ColAlignment(mconintCol������) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ɱ���) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ɱ����) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ�) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ۽��) = flexAlignRightCenter
        .ColAlignment(mconintCol����) = flexAlignRightCenter
        .ColAlignment(mconintCol��۲�) = flexAlignRightCenter
        .ColAlignment(mconintCol�̵���) = flexAlignRightCenter
        
        .ColAlignment(mconintCol���ű༭) = flexAlignRightCenter
        .ColAlignment(mconintCol���ر༭) = flexAlignRightCenter
        
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
        LblNo.Left = .Left - LblNo.Width - 100
        .Top = LblTitle.Top
        LblNo.Top = .Top
    End With
    
    TxtCheckDate.Left = mshBill.Left + mshBill.Width - TxtCheckDate.Width
    lblCheckDate.Left = TxtCheckDate.Left - lblCheckDate.Width - 100
    
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

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mconIntCol�к�, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call ��ʾ�ϼƽ��
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
    
    If mshBill.Col = mconIntCol���� Then
        If Not IsDate(TxtCheckDate) Then
            MsgBox "�̵�ʱ�䲻��,������!", vbInformation + vbDefaultButton1, gstrSysName
            If TxtCheckDate.Enabled Then TxtCheckDate.SetFocus
            Exit Sub
        End If
        Set RecReturn = Frm����ѡ����.ShowMe(Me, 2, txtStock.Tag, txtStock.Tag, txtStock.Tag, False, True, True, True, , , , , TxtCheckDate.Text, , , mbln���޴洢�ⷿ����, mstrPrivs, , False)
        If RecReturn.RecordCount > 0 Then
            mblnChange = True
            
            With mshBill
                RecReturn.MoveFirst
                For i = 1 To RecReturn.RecordCount
                    If SetPhiscRows(RecReturn!����ID, IIf(IsNull(RecReturn!����), 0, RecReturn!����)) Then
                        If .Row = .Rows - 1 Then .Rows = .Rows + 1 'ֻ�е�ǰ�������һ��ʱ��������
                        .Row = .Row + 1
                    End If
                    
                    RecReturn.MoveNext
                Next
                
                mshBill.Row = int�����
                
                If mstr�ظ����� <> "" Then
                    MsgBox mstr�ظ����� & "�б����Ѿ������ˣ�" & vbCrLf & "�������Ĳ�����ӣ�", vbInformation + vbOKOnly, gstrSysName
                    mstr�ظ����� = ""
                End If
            
    '            If RecReturn.RecordCount = 1 Then
    '                Call SetPhiscRows(RecReturn!����ID, IIf(IsNull(RecReturn!����), 0, RecReturn!����))
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
            mshBill.TextMatrix(mshBill.Row, mconIntCol����) = NVL(!����)
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
        If .Col = mconintColʵ������ Then
            strKey = .Text
            If strKey = "" Then
                strKey = .TextMatrix(.Row, .Col)
            End If
            Select Case .Col
                Case mconintColʵ������
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
        
        If .Col = mconIntCol�ɱ��� Then
            strKey = .Text
            If InStr(strKey, ".") <> 0 And Chr(KeyAscii) = "." Then   'ֻ�ܴ���һ��С����
                KeyAscii = 0
                Exit Sub
            End If
            
            If InStr("0123456789.", Chr(KeyAscii)) > 0 Or Chr(KeyAscii) = vbBack Or Chr(KeyAscii) = vbCr Then '�������������ֵ
                KeyAscii = KeyAscii
            Else
                KeyAscii = 0
            End If
            
            
        End If
    End With
End Sub

Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    Dim lng���� As Long
    
    With mshBill
        If .Active = False Then Exit Sub
        If mint�༭״̬ = 4 Then Exit Sub
        If .Row <> .LastRow Then
            
        End If
        
        Select Case .Col
            Case mconIntCol����
                .TxtCheck = False
                .MaxLength = 80
                'ֻ��ҩ���в���ʾ�ϼ���Ϣ�Ϳ����
                Call ��ʾ�ϼƽ��
                Call ��ʾ�����
            Case mconIntCol����
                .TxtCheck = False
                .MaxLength = mintBatchNoLen
            
            Case mconIntColЧ��
                .TxtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                If .ColData(mconIntColЧ��) = 2 Then
                    If .TextMatrix(.Row, mconIntCol����) <> "" And Len(Trim(.TextMatrix(.Row, mconIntCol����))) = 8 Then
                        Dim strxq As String
                        
                        If IsNumeric(.TextMatrix(.Row, mconIntCol����)) Then
                            strxq = UCase(.TextMatrix(.Row, mconIntCol����))
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                strxq = TranNumToDate(strxq)
                                If strxq <> "" Then .TextMatrix(.Row, mconIntColЧ��) = Format(DateAdd("M", .RowData(.Row), strxq), "yyyy-mm-dd")
                            End If
                        End If
                    End If
                End If
            Case mconintColʵ������
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
            Case mconIntCol�ɱ���
                If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
                       .ColData(mconIntCol�ɱ���) = IIf(Val(.TextMatrix(.Row, mconIntCol����)) = -1, 4, 5)
                End If
        End Select
        
        lng���� = Val(.TextMatrix(.Row, mconIntCol����))
        
        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
            .ColData(mconIntCol����) = IIf(lng���� = -1 Or Val(.TextMatrix(.Row, mconintCol���ر༭)) = 1, 1, 5)
            .ColData(mconIntCol����) = IIf(lng���� = -1 Or Val(.TextMatrix(.Row, mconintCol���ű༭)) = 1, 4, 5)
            .ColData(mconIntColЧ��) = IIf(lng���� = -1, 2, 5)
        End If
        
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsDrug As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim i As Integer
    Dim int����� As Integer
    
    On Error GoTo errHandle
    
    int����� = mshBill.Row
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With mshBill
        .Text = Trim(.Text)
        strKey = Trim(.Text)
        
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
                    If Not IsDate(TxtCheckDate) Then
                        MsgBox "�̵�ʱ�䲻��,������!", vbInformation + vbDefaultButton1, gstrSysName
                        If TxtCheckDate.Enabled Then TxtCheckDate.SetFocus
                        Exit Sub
                    End If
                    
                    Set RecReturn = FrmMulitSel.ShowSelect(Me, 2, txtStock.Tag, txtStock.Tag, txtStock.Tag, strKey, sngLeft, sngTop, mshBill.MsfObj.CellWidth, mshBill.MsfObj.CellHeight, False, True, True, True, , , , TxtCheckDate.Text, , , mbln���޴洢�ⷿ����, mstrPrivs, , False)
                    
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
            Case mconIntCol����
                If strKey = "" Then Exit Sub
                If SelectAndNotAddItem(Me, mshBill, strKey, "����������", "����������ѡ����", True, True, , zl_��ȡվ������(True)) = True Then
                    .Text = .TextMatrix(.Row, .Col)
                Else
                    .Text = ""
                    .Col = mconIntCol����
                    Cancel = True
                End If

            Case mconIntCol����
                '�޴���
                If strKey = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconIntCol����) = ""
                    End If
                    If .ColData(mconIntColЧ��) = 2 Then
                        .Col = mconIntColЧ��
                    Else
                        .Col = mconintColʵ������
                    End If
                    
                    Cancel = True
                    Exit Sub
                End If
            Case mconIntColЧ��
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
                ElseIf strKey = "" And strKey <> .TextMatrix(.Row, mconIntColЧ��) Then
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    End If
                    
                    Exit Sub
                End If
            Case mconintColʵ������
                If strKey <> "" Then
                    If Not IsNumeric(strKey) And strKey <> "" Then
                        ShowMsgBox "ʵ����������Ϊ������,�����䣡"
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                Else
                    .Text = IIf(.TextMatrix(.Row, .Col) = "", " ", .TextMatrix(.Row, .Col))
                    .TextMatrix(.Row, .Col) = .Text
                End If
                
                If strKey <> "" And .TextMatrix(.Row, 0) <> "" Then
                    strKey = Format(strKey, mFMT.FM_����)
                    .Text = strKey
                End If
                
                '��ʾ�ϼ�����
                If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
                If .Col = mconintColʵ������ Then
                    strKey = Val(.Text) * Val(.TextMatrix(.Row, mconIntCol����ϵ��))
                Else
                    strKey = Val(.TextMatrix(.Row, mconintColʵ������)) * Val(.TextMatrix(.Row, mconIntCol����ϵ��))
                End If
                
                .TextMatrix(.Row, mconIntCol�ɱ����) = Format(Val(.TextMatrix(.Row, mconIntCol�ɱ���)) * Val(.Text), mFMT.FM_���)
                .TextMatrix(.Row, mconIntCol�ۼ۽��) = Format(Val(.TextMatrix(.Row, mconIntCol�ۼ�)) * Val(.Text), mFMT.FM_���)
                
                Call ��ʾ�ϼƽ��
            Case mconIntCol�ɱ���
                If strKey <> "" Then
                    If Not IsNumeric(strKey) And strKey <> "" Then
                        ShowMsgBox "�ɱ��۱���Ϊ������,�����䣡"
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strKey = Format(strKey, mFMT.FM_�ɱ���)
                    .Text = strKey
                    
                    .TextMatrix(.Row, mconIntCol�ɱ����) = Format(Val(.TextMatrix(.Row, mconintColʵ������)) * Val(.Text), mFMT.FM_���)
                End If
                
                
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

Private Sub TxtCheckDate_GotFocus()
    With TxtCheckDate
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub TxtCheckDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey (vbKeyTab)
End Sub

Private Sub TxtCheckDate_Validate(Cancel As Boolean)
    
    If Not IsDate(TxtCheckDate.Text) Then
        ShowMsgBox "�����ʱ���ʽ!"
        Cancel = True
        Exit Sub
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
    Dim lngЧ�� As Long
    Dim rsTemp As New ADODB.Recordset
    
    If txtNO.Locked = False Then
        If Trim(txtNO.Text) = "" Then
            ShowMsgBox "���ݺŲ���Ϊ��"
            Exit Function
        End If
        If InStr(1, txtNO.Text, "'") <> 0 Then
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
        
            For intLop = 1 To .Rows - 1
                If Trim(.TextMatrix(intLop, mconIntCol����)) <> "" Then
                    If LenB(StrConv(Trim(Trim(.TextMatrix(intLop, mconIntCol����))), vbFromUnicode)) > mintBatchNoLen Then
                        ShowMsgBox "��" & intLop & "���������ϵ����ų���,���������" & Int(mintBatchNoLen / 2) & "�����ֻ�" & mintBatchNoLen & "���ַ�!"
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol����
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconintColʵ������)) > 9999999999# Then
                        ShowMsgBox "��" & intLop & "���������ϵ��������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconintColʵ������
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol����)) = -1 Then  '�������ļ����غ�����
                        
                        '�ж��Ƿ�ΪЧ����������
                        gstrSQL = "Select Nvl(���Ч��,0) Ч�� From �������� Where ����ID=[1]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ�ΪЧ����������", Val(.TextMatrix(intLop, 0)))
                        
                        lngЧ�� = rsTemp!Ч��
                        If lngЧ�� <> 0 Then
                            If Trim(.TextMatrix(intLop, mconIntCol����)) = "" Or Trim(.TextMatrix(intLop, mconIntColЧ��)) = "" Then
                                ShowMsgBox "��" & intLop & "�е�����������Ч�ڲ���,����������ż�Ч��" & vbCrLf & "��Ϣ�������뵥���У�"
                                mshBill.SetFocus
                                .Row = intLop
                                .MsfObj.TopRow = intLop
                                If .TextMatrix(intLop, mconIntCol����) = "" Then
                                    .Col = mconIntCol����
                                Else
                                    .Col = mconIntColЧ��
                                End If
                                Exit Function
                            End If
                        End If
                        
                        If mbln�����������Ų��ؿ��� = True Then
                            If Trim(.TextMatrix(intLop, mconIntCol����)) = "" Then '���ر�������
                                ShowMsgBox "��" & intLop & "�����������Ƿ������ϣ���¼����أ�"
                                mshBill.SetFocus
                                .Row = intLop
                                .MsfObj.TopRow = intLop
                                .Col = mconIntCol����
                                Exit Function
                            End If
                            
                            If Trim(.TextMatrix(intLop, mconIntCol����)) = "" Then  '���ر�������
                                ShowMsgBox "��" & intLop & "�����������Ƿ������ϣ���¼�����ţ�"
                                mshBill.SetFocus
                                .Row = intLop
                                .MsfObj.TopRow = intLop
                                .Col = mconIntCol����
                                Exit Function
                            End If
                        End If
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol����)) > 0 Then '��������
                        If mbln�����������Ų��ؿ��� = True Then
                            If Trim(.TextMatrix(intLop, mconIntCol����)) = "" Then '���ر�������
                                ShowMsgBox "��" & intLop & "�����������Ƿ������ϣ���¼����أ�"
                                mshBill.SetFocus
                                .Row = intLop
                                .MsfObj.TopRow = intLop
                                .Col = mconIntCol����
                                Exit Function
                            End If
                            
                            If Trim(.TextMatrix(intLop, mconIntCol����)) = "" Then  '���ر�������
                                ShowMsgBox "��" & intLop & "�����������Ƿ������ϣ���¼�����ţ�"
                                mshBill.SetFocus
                                .Row = intLop
                                .MsfObj.TopRow = intLop
                                .Col = mconIntCol����
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
    Dim dbl�ɱ��� As Double
    Dim dbl�ۼ� As Double
    Dim dbl���� As Double
    Dim dbl��۲� As Double
    Dim strժҪ As String
    Dim str������ As String
    Dim str�������� As String
    Dim str�̵�ʱ�� As String
    Dim dbl����� As Double
    Dim dbl����� As Double
    Dim rsTemp As New Recordset
    Dim intRow As Integer
    
    On Error GoTo errHandle
    SaveCard = False
    '����������������ID����Ҫ�������������϶�Ҫ����
    gstrSQL = "" & _
        "   SELECT b.ϵ��,b.id AS ���id " & _
        "   FROM ҩƷ�������� a, ҩƷ������ b " & _
        "   Where a.���id = b.ID " & _
        "           AND a.���� = 39 "
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, mstrCaption
    
    If rsTemp.EOF Then
        ShowMsgBox "û���������������̵����������������������������!"
        Exit Function
    End If
    
    lng������ID = 0
    lng�������ID = 0
    
    If rsTemp!ϵ�� = 1 Then lng������ID = rsTemp!���ID
    rsTemp.Close
    
    If lng������ID = 0 Then
        ShowMsgBox "û���������������̵��¼���������������������������!"
        Exit Function
    End If
    
    With mshBill
        lng�ⷿid = txtStock.Tag
        
        chrNo = Trim(txtNO.Text)
        If mint�༭״̬ = 1 Then
            If chrNo <> "" Then
                If CheckNOExists(76, chrNo) Then Exit Function
            End If
            If chrNo = "" Then chrNo = sys.GetNextNo(76, lng�ⷿid)
            If IsNull(chrNo) Then Exit Function
        End If
        txtNO.Tag = chrNo
        strժҪ = Trim(txtժҪ.Text)
        str������ = Txt������
        str�������� = Format(sys.Currentdate, "yyyy-mm-dd HH:mm:ss")
        str�̵�ʱ�� = TxtCheckDate.Text
        
        gcnOracle.BeginTrans
        If mint�༭״̬ = 2 Then        '�޸�
            gstrSQL = "zl_�����̵��¼��_Delete('" & mstr���ݺ� & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
        End If
            
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng����ID = .TextMatrix(intRow, 0)
                str���� = .TextMatrix(intRow, mconIntCol����)
                str���� = .TextMatrix(intRow, mconIntCol����)
                lng����ID = IIf(.TextMatrix(intRow, mconIntCol����) = "", 0, .TextMatrix(intRow, mconIntCol����))
                datЧ�� = IIf(Trim(.TextMatrix(intRow, mconIntColЧ��)) = "", "", .TextMatrix(intRow, mconIntColЧ��))
                
                dbl�������� = Round(Val(.TextMatrix(intRow, mconintCol��������)) * IIf(mintUnit = 1, Val(.TextMatrix(intRow, mconIntCol����ϵ��)), 1), g_С��λ��.obj_���С��.����С��)
                dblʵ������ = Round(Val(.TextMatrix(intRow, mconintColʵ������)) * IIf(mintUnit = 1, Val(.TextMatrix(intRow, mconIntCol����ϵ��)), 1), g_С��λ��.obj_���С��.����С��)
                
                dbl������ = 0
                dbl�ɱ��� = Round(Val(.TextMatrix(intRow, mconIntCol�ɱ���)) / IIf(mintUnit = 1, Val(.TextMatrix(intRow, mconIntCol����ϵ��)), 1), g_С��λ��.obj_���С��.�ɱ���С��)
                dbl�ۼ� = Round(Val(.TextMatrix(intRow, mconIntCol�ۼ�)) / IIf(mintUnit = 1, Val(.TextMatrix(intRow, mconIntCol����ϵ��)), 1), g_С��λ��.obj_���С��.���ۼ�С��)
                
                If Split(.TextMatrix(intRow, mconIntColָ�������), "||")(1) = 1 Then 'ʱ��
                    dbl�ۼ� = Get���ۼ�(lng����ID, lng�ⷿid, lng����ID, 1) 'ȡ�ۼ۵�λ����
                End If
                
                dbl���� = Round(Val(.TextMatrix(intRow, mconintCol����)), g_С��λ��.obj_���С��.���С��)
                dbl��۲� = Round(Val(.TextMatrix(intRow, mconintCol��۲�)), g_С��λ��.obj_���С��.���С��)
                dbl����� = Round(Val(.TextMatrix(intRow, mconIntColʵ�ʽ��)), g_С��λ��.obj_���С��.���С��)
                dbl����� = Round(Val(.TextMatrix(intRow, mconIntColʵ�ʲ��)), g_С��λ��.obj_���С��.���С��)
                
                If dbl�������� <= dblʵ������ Then
                    lng������ID = lng������ID
                    int���ϵ�� = 1
                Else
                    lng������ID = lng�������ID
                    int���ϵ�� = -1
                End If
                 
                lng��� = intRow
                
                'zl_�����̵��¼��_INSERT( /*NO_IN*/, /*���_IN*/, /*�ⷿID_IN*/, /*����_IN*/,
                    '/*������ID_IN*/, /*���ϵ��_IN*/, /*����ID_IN*/, /*��������_IN*/,
                    '/*ʵ������_IN*/, /*������_IN*/, /*�ۼ�_IN*/, /*����_IN*/, /*��۲�_IN*/,
                    '/*������_IN*/, /*��������_IN*/, /*ժҪ_IN*/, /*����_IN*/, /*����_IN*/,
                    '/*Ч��_IN*/, /*�̵�ʱ��_IN*/ );
                
                gstrSQL = "zl_�����̵��¼��_INSERT('" & _
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
                    str�������� & "','yyyy-mm-dd HH24:MI:SS'),'" & _
                    strժҪ & "','" & _
                    str���� & "','" & _
                    str���� & "'," & _
                    IIf(datЧ�� = "", "Null", "to_date('" & Format(datЧ��, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ",'" & _
                    str�̵�ʱ�� & "'," & _
                    dbl����� & "," & _
                    dbl����� & ")"
                
                Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
            End If
        Next
        gcnOracle.CommitTrans
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Function

Private Sub ��ʾ�ϼƽ��()
    Dim dbl�ɱ���� As Double
    Dim dbl�̵��� As Double
    Dim intLop As Integer

    dbl�ɱ���� = 0
    dbl�̵��� = 0

    With mshBill
        For intLop = 1 To .Rows - 1
            If .TextMatrix(intLop, 0) <> "" Then
                dbl�ɱ���� = dbl�ɱ���� + Val(.TextMatrix(intLop, mconIntCol�ɱ����))
                dbl�̵��� = dbl�̵��� + Val(.TextMatrix(intLop, mconIntCol�ۼ۽��))
            End If
        Next
    End With

    lblPurchasePrice.Caption = "�̵�ɱ����ϼƣ�" & Format(dbl�ɱ����, mFMT.FM_���)
    lblPurchasePrice.Width = Pic����.TextWidth(lblPurchasePrice.Caption)
    lblCheckSum.Left = lblPurchasePrice.Left + lblPurchasePrice.Width + 200

    lblCheckSum.Caption = "�̵���ϼƣ�" & Format(dbl�̵���, mFMT.FM_���)
    lblCheckSum.Width = Pic����.TextWidth(lblCheckSum.Caption)
'
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
        "   Where a.ҩƷid=[1] and nvl(a.����,0)=[2] " & _
        "           AND a.����=1 " & _
        "           AND a.�ⷿid =[3] "
        
    gstrSQL = strKc
    With mshBill
        If .TextMatrix(.Row, mconIntCol����) = "" Then
            stbThis.Panels(2).Text = ""
            Exit Sub
        End If
        If .TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
        
'        gstrSQL = "" & _
            "   Select ��������/" & IIf(mintUnit = 0, 1, Val(.TextMatrix(.Row, mconIntCol����ϵ��))) & " as  �������� " & _
            "   From ҩƷ��� " & _
            "   where �ⷿid=[3]" & _
            "           and ҩƷid=[1]" & _
            "           and ����=1 and " & _
            "           nvl(����,0)=[2]"
        
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol����)), Val(txtStock.Tag), CDate(TxtCheckDate.Text), IIf(mintUnit = 0, 1, Val(.TextMatrix(.Row, mconIntCol����ϵ��))))
        
        If rsTemp.EOF Then
            .TextMatrix(.Row, mconIntCol��������) = 0
        Else
            .TextMatrix(.Row, mconIntCol��������) = IIf(IsNull(rsTemp.Fields(0)), 0, rsTemp.Fields(0))
        End If
        rsTemp.Close
        stbThis.Panels(2).Text = "���������ϵ�ǰ�����Ϊ[" & Format(.TextMatrix(.Row, mconIntCol��������), mFMT.FM_����) & "]" & .TextMatrix(.Row, mconIntCol��λ)
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
    '���ܣ�������������ID���̴������ʾ��������������ϵĳ�ʼ�̴���Ϣ
    '˵����
    '   1.����Ƿǿⷿ����ҩ,���Ѿ�������,����ʾ���˳���
    '   2.����ǿⷿ����ҩ����ֱ����ҩ��δ����ĸ����ο���С�
    Dim i As Integer, lngRow As Long
    Dim rsData As ADODB.Recordset
    Dim blnModi As Boolean, sngLevel As Single
    Dim intRecordCount As Integer
    Dim intCurrentRow As Integer
    Dim intRow As Integer
    Dim rsprice As New Recordset
    
    On Error GoTo errH
    
    SetPhiscRows = False
    Set rsData = GetDateStock(TxtCheckDate.Text, txtStock.Tag, "", True, True, lngId)
    intRecordCount = rsData.RecordCount
    If intRecordCount = 0 Then Exit Function
    '����������������
    If lng���� <> -1 Then
        rsData.MoveFirst
        rsData.Find "����=" & lng����
        If rsData.EOF Then Exit Function
    End If
    
    With mshBill
        If lng���� <> -1 Then
            For intRow = 1 To .Rows - 1
                If .TextMatrix(intRow, 0) <> "" Then
                    If Val(.TextMatrix(intRow, 0)) = lngId And IIf(.TextMatrix(intRow, mconIntCol����) = "", "0", .TextMatrix(intRow, mconIntCol����)) = lng���� Then
                        If UBound(Split(mstr�ظ�����, "��")) < 3 Then mstr�ظ����� = mstr�ظ����� & .TextMatrix(intRow, mconIntCol����) & "��"  '����¼�����ظ�������
                        'ShowMsgBox "�����������ϡ�" & .TextMatrix(intRow, mconIntCol����) & "(" & lng���� & ")����������ӣ�"
                        Exit Function
                    End If
                End If
            Next
        End If
        
        mshBill.Redraw = False
        intRow = .Row
        intCurrentRow = .Row
        .TextMatrix(intRow, 0) = rsData!����ID
        .TextMatrix(intRow, mconIntCol����) = "[" & rsData!���� & "]" & rsData!��Ʒ����
        .TextMatrix(intRow, mconIntCol���) = IIf(IsNull(rsData!���), "", rsData!���)
        .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsData!����), "", rsData!����)
        .TextMatrix(intRow, mconIntCol�ⷿ��λ) = IIf(IsNull(rsData!�ⷿ��λ), "", rsData!�ⷿ��λ)
        .TextMatrix(intRow, mconIntCol��λ) = zlStr.NVL(rsData!��λ)
        
        If lng���� = -1 Then
            .TextMatrix(intRow, mconIntCol����) = lng����
            .TextMatrix(intRow, mconIntCol����) = ""
            .TextMatrix(intRow, mconIntColЧ��) = ""
        Else
            .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsData!����), "0", rsData!����)
            .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsData!����), "", rsData!����)
            .TextMatrix(intRow, mconIntColЧ��) = IIf(IsNull(rsData!Ч��), "", Format(rsData!Ч��, "yyyy-MM-dd"))
        End If
        
        If Val(.TextMatrix(intRow, mconIntCol����)) <> 0 Then
            .TextMatrix(intRow, mconintCol���ű༭) = rsData!���ű༭
            .TextMatrix(intRow, mconintCol���ر༭) = rsData!���ر༭
        End If
        
        .TextMatrix(intRow, mconIntCol����ϵ��) = IIf(mintUnit = 0, 1, zlStr.NVL(rsData!����ϵ��)) ' ��ȡ����ϵ��(rsData)
        .TextMatrix(intRow, mconIntColָ�������) = rsData!ָ������� & "||" & rsData!�Ƿ��� & "||" & rsData!���÷���
        
        .TextMatrix(intRow, mconIntCol�ɱ���) = Format(rsData!�ɱ��� * Val(.TextMatrix(intRow, mconIntCol����ϵ��)), mFMT.FM_�ɱ���)
        .TextMatrix(intRow, mconIntCol�ɱ����) = Format(Val(.TextMatrix(intRow, mconIntCol�ɱ���)) * Val(.TextMatrix(intRow, mconintColʵ������)), mFMT.FM_���)
        
        If rsData!�Ƿ��� = 1 Then
'            gstrSQL = "" & _
'                "   Select ʵ�ʽ��/ʵ������*" & IIf(mintUnit = 0, "1", zlStr.NVL(rsData!����ϵ��)) & " as  �ۼ� " & _
'                "   From ҩƷ��� " & _
'                "   Where �ⷿid=[1] " & _
'                "           and ҩƷid=[2]" & _
'                "  and ����=1 and ʵ������>0 and " & _
'                "  nvl(����,0)=[3]"
'
'            Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(txtStock.Tag), Val(zlStr.NVL(rsData!����ID)), Val(zlStr.NVL(rsData!����)))
'
'            If rsprice.EOF Then
'                .TextMatrix(intRow, mconIntCol�ۼ�) = Format(IIf(IsNull(rsData.Fields("�ۼ�").Value), 0, rsData.Fields("�ۼ�").Value), mFMT.FM_���ۼ�)
'            Else
'                .TextMatrix(intRow, mconIntCol�ۼ�) = Format(rsprice.Fields(0), mFMT.FM_���ۼ�)
'            End If
            
            .TextMatrix(intRow, mconIntCol�ۼ�) = Format(Get���ۼ�(lngId, Val(txtStock.Tag), lng����, Val(.TextMatrix(intRow, mconIntCol����ϵ��))), mFMT.FM_���ۼ�)
            
        Else '����
            gstrSQL = "SELECT  �ּ� as �ۼ� From �շѼ�Ŀ WHERE ((SYSDATE BETWEEN ִ������ AND ��ֹ����) OR (SYSDATE >= ִ������ AND ��ֹ���� IS NULL))" & _
                    GetPriceClassString("") & " And �շ�ϸĿid = [1] "
            
            Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, Val(zlStr.NVL(rsData!����ID)))
            
            .TextMatrix(intRow, mconIntCol�ۼ�) = Format(rsprice.Fields(0) * Val(.TextMatrix(intRow, mconIntCol����ϵ��)), mFMT.FM_���ۼ�)
        End If
        
        .TextMatrix(intRow, mconIntCol�ۼ۽��) = Format(Val(.TextMatrix(intRow, mconIntCol�ۼ�)) * Val(.TextMatrix(intRow, mconintColʵ������)), mFMT.FM_���)
        
        .RowData(intRow) = IIf(IsNull(rsData!���Ч��), 0, rsData!���Ч��)
        rsData.MoveNext
        
        Call RefreshRowNO(mshBill, mconIntCol�к�, 1)
        .Col = IIf(lng���� = -1, mconIntCol����, mconintColʵ������)
        mshBill.Redraw = True
    End With
    Call ��ʾ�����
    
    rsData.Close
    SetPhiscRows = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'��һ���в���
Private Sub InsertRow(ByVal intRow As Integer, ByVal intRecordCount As Integer)
    Dim blnHaveData As Boolean
    Dim intOldRows As Integer
    Dim intLop As Integer
    Dim intExchange As Integer
    Dim intCol As Integer
    
    With mshBill
        blnHaveData = False
        intOldRows = .Rows - 1
        .Rows = .Rows + intRecordCount
        For intLop = intRow + 1 To intRecordCount
            If .TextMatrix(intLop, 0) <> "" Then
                blnHaveData = True
                Exit For
            End If
        Next
        If blnHaveData = True Then
            For intExchange = .Rows - 1 To intOldRows Step -1
                For intCol = 0 To .Cols - 1
                    .TextMatrix(intExchange, intCol) = .TextMatrix(intExchange - intRecordCount, intCol)
                    .TextMatrix(intExchange - intRecordCount, intCol) = ""
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
'    Dim StrNo As String
'    StrNo = txtNO.Tag
'    Call FrmBillPrint.ShowME(Me, glngSys, "zl1_bill_1719", mint��¼״̬, mintUnit, 1719, "���������̵��¼��", StrNo)
End Sub

'ȡ���ݿ������ŵĳ��ȣ������������е����ų��������ݿ��б���һ����
Private Function GetBatchNoLen() As Integer
    Dim rsBatchNolen As New Recordset
    
    On Error GoTo errHandle
    gstrSQL = "select ���� from ҩƷ�շ���¼ where rownum<1 "
    Call zlDatabase.OpenRecordset(rsBatchNolen, gstrSQL, "ȡ�ֶγ���")
    GetBatchNoLen = rsBatchNolen.Fields(0).DefinedSize
    rsBatchNolen.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


