VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmMediHerbalItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�в�ҩƷ�ֱ༭"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8145
   Icon            =   "frmMediHerbalItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chk������ҩ 
      Caption         =   "������ҩ"
      Height          =   210
      Left            =   5400
      TabIndex        =   38
      Top             =   4680
      Width           =   1050
   End
   Begin VB.CommandButton cmdSaveAddSpec 
      Caption         =   "������������(&B)"
      Height          =   350
      Left            =   3720
      TabIndex        =   47
      Top             =   5625
      Width           =   1695
   End
   Begin VB.CommandButton cmdSaveAddItem 
      Caption         =   "���������Ʒ��(&A)"
      Height          =   350
      Left            =   1920
      TabIndex        =   46
      Top             =   5625
      Width           =   1695
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "��"
      Height          =   285
      Left            =   4800
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   248
      Width           =   285
   End
   Begin VB.CheckBox chkԭ��ҩ 
      Caption         =   "ԭ��ҩ(&M)"
      Height          =   210
      Left            =   5400
      TabIndex        =   37
      Top             =   4320
      Width           =   1155
   End
   Begin VB.CheckBox chk����Ӧ�� 
      Caption         =   "��ζʹ��(&Q)"
      Height          =   210
      Left            =   5400
      TabIndex        =   36
      Top             =   3960
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6840
      TabIndex        =   42
      Top             =   5625
      Width           =   1100
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   315
      Picture         =   "frmMediHerbalItem.frx":058A
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   5625
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "�����˳�(&O)"
      Height          =   350
      Left            =   5520
      TabIndex        =   39
      Top             =   5625
      Width           =   1215
   End
   Begin VB.TextBox txtӢ�� 
      Height          =   300
      Left            =   1395
      TabIndex        =   10
      Top             =   1815
      Width           =   3675
   End
   Begin VB.TextBox txt��� 
      Height          =   300
      Left            =   3300
      MaxLength       =   12
      TabIndex        =   8
      Top             =   1410
      Width           =   1170
   End
   Begin VB.ComboBox cboҩƷ���� 
      Height          =   300
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   1815
      Width           =   1455
   End
   Begin VB.ComboBox cboҽ��ְ�� 
      Height          =   300
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   2625
      Width           =   1455
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1395
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   1
      Top             =   240
      Width           =   3360
   End
   Begin VB.TextBox txtƴ�� 
      Height          =   300
      Left            =   1395
      MaxLength       =   12
      TabIndex        =   7
      Top             =   1425
      Width           =   1170
   End
   Begin VB.ComboBox cbo���� 
      Height          =   300
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   240
      Width           =   1455
   End
   Begin VB.ComboBox cbo��Դ 
      Height          =   300
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   1020
      Width           =   1455
   End
   Begin VB.ComboBox cbo��ֵ 
      Height          =   300
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   630
      Width           =   1455
   End
   Begin VB.ComboBox cbo�ݴ� 
      Height          =   300
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   1425
      Width           =   1455
   End
   Begin VB.TextBox txt�������� 
      Height          =   300
      Left            =   6480
      MaxLength       =   16
      TabIndex        =   33
      Text            =   "0"
      Top             =   3030
      Width           =   1455
   End
   Begin VB.ComboBox cbo��λ 
      Height          =   300
      Left            =   1395
      TabIndex        =   12
      Top             =   2220
      Width           =   1155
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1395
      MaxLength       =   40
      TabIndex        =   5
      Top             =   1020
      Width           =   3675
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1395
      MaxLength       =   13
      TabIndex        =   3
      Top             =   630
      Width           =   1935
   End
   Begin VB.ComboBox cbo����ְ�� 
      Height          =   300
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   2190
      Width           =   1455
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -120
      TabIndex        =   43
      Top             =   5265
      Width           =   8490
   End
   Begin VB.TextBox txt�ο� 
      Height          =   300
      Left            =   1395
      TabIndex        =   14
      Top             =   2610
      Width           =   3380
   End
   Begin VB.CommandButton cmd�ο� 
      Caption         =   "��"
      Height          =   285
      Left            =   4800
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2610
      Width           =   285
   End
   Begin VB.ComboBox cbo�����Ա� 
      Height          =   300
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   3465
      Width           =   1455
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   3570
      Left            =   1560
      TabIndex        =   44
      TabStop         =   0   'False
      Tag             =   "1000"
      Top             =   5880
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   6297
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgList"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   4275
      Top             =   5520
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
            Picture         =   "frmMediHerbalItem.frx":06D4
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediHerbalItem.frx":0C6E
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediHerbalItem.frx":1208
            Key             =   "ItemUse1"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediHerbalItem.frx":17A2
            Key             =   "ItemStop1"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediHerbalItem.frx":1D3C
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediHerbalItem.frx":2436
            Key             =   "ItemStop"
         EndProperty
      EndProperty
   End
   Begin ZL9BillEdit.BillEdit msf���� 
      Height          =   1815
      Left            =   300
      TabIndex        =   17
      Top             =   3345
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   3201
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "ע����Ʒ�ֽ�����2003-09-01"
      Height          =   180
      Left            =   300
      TabIndex        =   40
      Top             =   5385
      Width           =   2340
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   16
      Top             =   3090
      Width           =   720
   End
   Begin VB.Label lblӢ�� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ӣ������(&E)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   330
      TabIndex        =   9
      Top             =   1890
      Width           =   990
   End
   Begin VB.Label LblҩƷ���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҩƷ����(&K)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5445
      TabIndex        =   26
      Top             =   1875
      Width           =   990
   End
   Begin VB.Label Lblҽ��ְ�� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ��ְ��(&I)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5445
      TabIndex        =   30
      Top             =   2685
      Width           =   990
   End
   Begin VB.Label Lbl����ְ�� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����ְ��(&Z)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5445
      TabIndex        =   28
      Top             =   2250
      Width           =   990
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҩƷ����(&F)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   330
      TabIndex        =   0
      Top             =   315
      Width           =   990
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���Ƽ���(&S)              (ƴ��)               (���)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   330
      TabIndex        =   6
      Top             =   1485
      Width           =   4680
   End
   Begin VB.Label Lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�������(&X)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5445
      TabIndex        =   18
      Top             =   300
      Width           =   990
   End
   Begin VB.Label Lbl��Դ 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��Դ���(&R)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5445
      TabIndex        =   22
      Top             =   1080
      Width           =   990
   End
   Begin VB.Label Lbl��ֵ 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ֵ����(&V)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5445
      TabIndex        =   20
      Top             =   690
      Width           =   990
   End
   Begin VB.Label Lbl�ݴ� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ҩ�ݴ�(&G)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5445
      TabIndex        =   24
      Top             =   1485
      Width           =   990
   End
   Begin VB.Label Lbl�������� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������(&L)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5445
      TabIndex        =   32
      Top             =   3090
      Width           =   990
   End
   Begin VB.Label Lbl��λ 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������λ(&U)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   330
      TabIndex        =   11
      Top             =   2265
      Width           =   990
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ͨ������(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   330
      TabIndex        =   4
      Top             =   1080
      Width           =   990
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҩƷ����(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   330
      TabIndex        =   2
      Top             =   690
      Width           =   990
   End
   Begin VB.Label Label2 
      Caption         =   "�ο���Ŀ(&F)"
      Height          =   255
      Left            =   330
      TabIndex        =   13
      Top             =   2670
      Width           =   1095
   End
   Begin VB.Label lbl�����Ա� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�����Ա�(&S)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5430
      TabIndex        =   34
      Top             =   3525
      Width           =   990
   End
End
Attribute VB_Name = "frmMediHerbalItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'˵����
'   1����ǰ���ʣ���Me.tag��ţ��ֱ�Ϊ1-����ҩ��2-�г�ҩ�����ϼ�������
'   2���༭״̬����Me.cmdCancel.Tag��ţ��ֱ�Ϊ"����"��"�޸�"��"����"�����ϼ�������
'---------------------------------------------------
Public lng����id As Long        '���༭�ķ���ID���ϼ����򴫵ݽ���
Public lngҩ��id As Long        '���༭��ҩ��ID���޸ġ�����ʱ���ϼ����򴫵ݽ���
Public strPrivs As String       '��ǰ�û��Ա������Ȩ�ޣ����ϼ�����򴫵ݽ���
Private mint������� As Integer     'ҩƷƷ�ֱ����������
Private mblnOK As Boolean       '��¼ȷ����ť�Ƿ񱻵����
Private mblnCancel As Boolean   '��¼ȡ����ť�Ƿ񱻵����

Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim strTemp As String, aryTemp() As String
Dim intCount As Integer

Dim mstrMatch As String, strRefer As String '�ο�����
Private mstrID As String, mstr���� As String, mstr���� As String
Private mblnLoad As Boolean      '��¼������صĴ���
Private mstr���м�¼ As String      '��¼���������е�ֵ

Private mlng���볤�� As Long
Private mlng���볤�� As Long
Private mint���Ƴ��� As Integer
Private mintӢ�ĳ��� As Integer
Private mbln�Թ�ҩ As Boolean

Public Sub ShowMe(ByVal bln�Թ�ҩ As Boolean, ByVal frmPar As Form)
    mbln�Թ�ҩ = bln�Թ�ҩ
    Me.Show vbModal, frmPar
End Sub

Private Sub GetDefineSize()
    '���ܣ��õ����ݿ�ı��ֶεĳ���
    On Error GoTo errHandle
    Dim rsTmp As New ADODB.Recordset
    
    gstrSql = "Select ���� From �շ���ĿĿ¼ Where Rownum<0"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSql, Me.Caption)
    
    mlng���볤�� = rsTmp.Fields("����").DefinedSize
        
    txt����.MaxLength = mlng���볤��
    
    gstrSql = "Select ����,���� From ������Ŀ���� Where Rownum<0"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSql, Me.Caption)
    
    mlng���볤�� = rsTmp.Fields("����").DefinedSize
    mintӢ�ĳ��� = rsTmp.Fields("����").DefinedSize
    
    txtƴ��.MaxLength = mlng���볤��
    txt���.MaxLength = mlng���볤��
    txtӢ��.MaxLength = mintӢ�ĳ���
    
    gstrSql = "Select ���� From ������ĿĿ¼ Where Rownum<0"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSql, Me.Caption)
    
    mint���Ƴ��� = rsTmp.Fields("����").DefinedSize
    txt����.MaxLength = mint���Ƴ���
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub cbo����ְ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cbo��λ_GotFocus()
    Me.cbo��λ.SelStart = 0: Me.cbo��λ.SelLength = 100
End Sub

Private Sub cbo��λ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Or (Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9) Then KeyAscii = 0
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cbo��Դ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cbo��ֵ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cbo�����Ա�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cbo�ݴ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cboҩƷ����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub cboҽ��ְ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chk����Ӧ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub chkԭ��ҩ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub
Private Sub chk������ҩ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub
Private Sub cmdCancel_Click()
    Dim strTemp As String
    Dim i As Integer
    Dim str���� As String
    
    With msf����
        For i = 1 To .Rows - 1
            str���� = str���� & "|" & .TextMatrix(i, 1) & "," & .TextMatrix(i, 2) & "," & .TextMatrix(i, 3)
        Next
    End With
    strTemp = txt����.Text & "|" & txt����.Text & "|" & txt����.Text & "|" & txtƴ��.Text & "|" & txt���.Text & "|" & txtӢ��.Text & "|" & cbo��λ.Text & "|" & _
                txt�ο�.Text & "|" & cbo����.Text & "|" & cbo��ֵ.Text & "|" & cbo��Դ.Text & "|" & cbo�ݴ�.Text & "|" & cboҩƷ����.Text & "|" & _
                cbo����ְ��.Text & "|" & cboҽ��ְ��.Text & "|" & txt��������.Text & "|" & cbo�����Ա�.Text & "|" & chk������ҩ.Value & "|" & chk����Ӧ��.Value & "|" & chkԭ��ҩ.Value & "|" & _
                str����
    If strTemp <> mstr���м�¼ And Me.cmdCancel.Tag <> "����" Then
        mblnCancel = True
        If MsgBox("�����ݱ��޸���ȷ���˳���", vbYesNo, gstrSysName) = vbYes Then
            Unload Me
        Else
            mblnCancel = False
        End If
    Else
        Unload Me
    End If
    gblnCancel = True
    Exit Sub
End Sub

Private Sub cmdOK_Click()
    '�༭���ݼ��
    mblnOK = True
    If Trim(Me.txt����.Text) = "" Then
        MsgBox "������ҩƷ���룡", vbInformation, gstrSysName
        Me.txt����.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > mlng���볤�� Then
        MsgBox "ҩƷ����ĳ��ȳ��������" & mlng���볤�� & "���ַ�����", vbInformation, gstrSysName
        Me.txt����.SetFocus: Exit Sub
    End If
    If Trim(Me.txt����.Text) = "" Then
        MsgBox "������ͨ�����ƣ�", vbInformation, gstrSysName
        Me.txt����.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > mint���Ƴ��� Then
        MsgBox "ͨ�����Ƴ��ȳ��������" & mint���Ƴ��� & "���ַ���" & Int(mint���Ƴ��� / 2) & "�����֣���", vbInformation, gstrSysName
        Me.txt����.SetFocus: Exit Sub
    End If
    If Trim(Me.cbo��λ.Text) = "" Then
        MsgBox "�����������λ��", vbInformation, gstrSysName
        Me.cbo��λ.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.cbo��λ.Text), vbFromUnicode)) > 10 Then
        MsgBox "������λ�ĳ��ȳ��������10���ַ���5�����֣���", vbInformation, gstrSysName
        Me.cbo��λ.SetFocus: Exit Sub
    End If
    
    '�������
    strTemp = ";" & Trim(Me.txt����.Text) & ";" & Trim(Me.txtӢ��.Text)
    With Me.msf����
        For intCount = 1 To .Rows - 1
            If Trim(.TextMatrix(intCount, 1)) <> "" Then
                If InStr(1, strTemp & ";", ";" & Trim(.TextMatrix(intCount, 1)) & ";") > 0 Then
                    MsgBox "���������ظ�������ͨ�����ƺ�Ӣ�����ƣ���", vbInformation, gstrSysName
                    .SetFocus: Exit Sub
                Else
                    strTemp = strTemp & ";" & Trim(.TextMatrix(intCount, 1))
                End If
            End If
        Next
    End With
    
    '���ݱ���
    If Me.cmdCancel.Tag = "����" Then
        lngҩ��id = Sys.NextId("������ĿĿ¼")
        If zlClinicCodeRepeat(Trim(Me.txt����.Text)) = True Then Exit Sub
    Else
        If zlClinicCodeRepeat(Trim(Me.txt����.Text), lngҩ��id) = True Then Exit Sub
    End If
    gstrSql = Me.txt����.Tag & "," & lngҩ��id & ",'" & Trim(Me.txt����.Text) & "','" & Trim(Me.txt����.Text) & "'"
    gstrSql = gstrSql & ",'" & MoveSpecialChar(Trim(Me.txtƴ��.Text)) & "','" & MoveSpecialChar(Trim(Me.txt���.Text)) & "','" & MoveSpecialChar(Trim(Me.txtӢ��.Text)) & "'"
    gstrSql = gstrSql & ",'" & MoveSpecialChar(Trim(Me.cbo��λ.Text)) & "'"
    gstrSql = gstrSql & ",'" & Mid(Me.cbo����.Text, InStr(1, Me.cbo����.Text, "-") + 1) & "'"
    gstrSql = gstrSql & ",'" & Mid(Me.cbo��ֵ.Text, InStr(1, Me.cbo��ֵ.Text, "-") + 1) & "'"
    gstrSql = gstrSql & ",'" & Mid(Me.cbo��Դ.Text, InStr(1, Me.cbo��Դ.Text, "-") + 1) & "'"
    gstrSql = gstrSql & ",'" & Mid(Me.cbo�ݴ�.Text, InStr(1, Me.cbo�ݴ�.Text, "-") + 1) & "'"
    gstrSql = gstrSql & "," & Left(Me.cboҩƷ����.Text, 1) & ",'" & Left(Me.cbo����ְ��.Text, 1) & Left(Me.cboҽ��ְ��.Text, 1) & "'"
    gstrSql = gstrSql & "," & Val(Trim(Me.txt��������.Text))
    gstrSql = gstrSql & "," & Me.chk����Ӧ��.Value & "," & Me.chkԭ��ҩ.Value
    gstrSql = gstrSql & "," & Left(Me.cbo�����Ա�.Text, 1)
    gstrSql = gstrSql & "," & ZVal(Me.txt�ο�.Tag) '& "'"
    strTemp = ""
    With Me.msf����
        For intCount = 1 To .Rows - 1
            If Trim(.TextMatrix(intCount, 1)) <> "" Then
                strTemp = strTemp & "|" & Trim(.TextMatrix(intCount, 1)) & "^" & Trim(.TextMatrix(intCount, 2)) & "^" & Trim(.TextMatrix(intCount, 3))
            End If
        Next
    End With
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    
    '������������
    If LenB(strTemp) > 4000 Then
        msf����.SetFocus
        MsgBox "�����ַ���̫��������ٱ����������߱������ȡ�", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    err = 0: On Error GoTo ErrHand
    If Me.cmdCancel.Tag = "����" Then
        gstrSql = "zl_��ҩƷ��_INSERT('7'," & gstrSql & ",'" & strTemp & "'," & IIf(mbln�Թ�ҩ = True, "1", "Null") & "," & Me.chk������ҩ.Value & ")"
    Else
        gstrSql = "zl_��ҩƷ��_UPDATE(" & gstrSql & ",'" & strTemp & "'," & IIf(mbln�Թ�ҩ = True, "1", "Null") & "," & Me.chk������ҩ.Value & ")"
    End If
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    If Me.cmdCancel.Tag = "����" Then
        'Val(zldatabase.GetPara("Ʒ������ģʽ", glngSys, 1023, 0))
        Select Case ActiveControl
        Case cmdSaveAddItem  'Ʒ����������
            Call frmMediLists.zlRefRecords(lngҩ��id)
            lngҩ��id = 0
            Call Form_Activate
            Me.txt����.SetFocus
        Case cmdSaveAddSpec  '����Ʒ�ֺ����ӹ��
            Unload Me
            Call frmMediLists.zlRefRecords(lngҩ��id)
            With frmMediHerbalSpec
                .mlng����id = lng����id
                .stbSpec.Tag = "����"
                .lngҩ��id = lngҩ��id
                .lngҩƷID = 0
                .strPrivs = Me.strPrivs
                .Show 1, frmMediLists
            End With
        Case Else
            Unload Me
        End Select
    Else
        Unload Me
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSaveAddItem_Click()
    Call cmdOK_Click
End Sub

Private Sub cmdSaveAddSpec_Click()
    Call cmdOK_Click
End Sub

Private Sub cmd����_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmd�ο�_Click()
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = SelectRefer
    If Not rsTmp Is Nothing Then
        Me.txt�ο� = rsTmp("����"): Me.txt�ο�.Tag = rsTmp("ID"): strRefer = Me.txt�ο�
    End If
End Sub

Private Function SelectRefer(Optional ByVal strName As String = "") As ADODB.Recordset
    Dim strSql As String, strSQLItem As String
    Dim rsTmp As New ADODB.Recordset, iAttr As Integer
    
    On Error GoTo ErrHand
    strSql = "Select ���� From ���Ʒ���Ŀ¼ Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����id)
    
    If rsTmp.EOF Then
        iAttr = -1
    Else
        iAttr = rsTmp(0)
    End If
    If Len(strName) = 0 Then
        strSql = " Select ID,����ID,����,����,˵�� From ���Ʋο�Ŀ¼ a Where ����=" & iAttr & " Order By ����"
    Else
        strSQLItem = " From ���Ʋο�Ŀ¼ A,���Ʋο����� B" & _
            " Where A.ID=B.�ο�Ŀ¼ID And A.����=" & iAttr & _
            " And (Upper(A.����) Like '" & UCase(strName) & "%'" & _
            " Or Upper(A.����) Like '" & mstrMatch & UCase(strName) & "%'" & _
            " Or Upper(B.����) Like '" & mstrMatch & UCase(strName) & "%'" & _
            " Or Upper(B.����) Like '" & mstrMatch & UCase(strName) & "%')"

        strSql = " Select DISTINCT A.ID,A.����ID,A.����,A.����,A.˵�� " & strSQLItem & " Order By ����"
    End If
    Set SelectRefer = zlDatabase.ShowSelect(Me, strSql, 0, "�ο�", , , , , True)
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub cmd����_Click()
    Dim blnRe As Boolean
    
    gstrSql = "select ID,�ϼ�ID,����,����,����" & _
            " From ���Ʒ���Ŀ¼" & _
            " Where ���� = " & Me.Tag & _
            " start with �ϼ�ID is null" & _
            " connect by prior ID=�ϼ�ID"
    blnRe = frmTreeSel.ShowTree(gstrSql, mstrID, mstr����, mstr����, "", "ҩƷ����", "���з���", False)
    If blnRe Then
        txt����.Text = "[" & mstr���� & "]" & mstr����
        txt����.Tag = mstrID
        Me.txt����.SetFocus
        lng����id = mstrID
    End If
    mblnLoad = True
End Sub

Private Sub Form_Activate()
    Dim strCode As String
    Dim str���� As String
    Dim i As Integer
    
    
'    If Me.Tag = "1" Then
'        Me.Caption = "����ҩƷ��" & Me.cmdCancel.Tag
'    Else
'        Me.Caption = "�г�ҩƷ��" & Me.cmdCancel.Tag
'    End If
    gblnCancel = False
    
    If Me.cmdCancel.Tag <> "����" Then cmdSaveAddItem.Enabled = False: cmdSaveAddSpec.Enabled = False
    If mblnLoad = False Then
        Me.Caption = "�в�ҩƷ��" & Me.cmdCancel.Tag
        
        '�������ݼ��
        'If Me.cbo����.ListCount = 0 Then MsgBox "δ����ҩƷ���ͣ������ֵ�����н�������", vbExclamation, gstrSysName: Unload Me: Exit Sub
        If Me.cbo����.ListCount = 0 Then MsgBox "�޶���������ݣ�����ϵϵͳ����Ա", vbExclamation, gstrSysName: Unload Me: Exit Sub
        If Me.cbo��ֵ.ListCount = 0 Then MsgBox "�޼�ֵ�������ݣ�����ϵϵͳ����Ա", vbExclamation, gstrSysName: Unload Me: Exit Sub
        If Me.cbo��Դ.ListCount = 0 Then MsgBox "�޻�Դ�������ݣ�����ϵϵͳ����Ա", vbExclamation, gstrSysName: Unload Me: Exit Sub
        If Me.cbo�ݴ�.ListCount = 0 Then MsgBox "����ҩ�ݴ����ݣ�����ϵϵͳ����Ա", vbExclamation, gstrSysName: Unload Me: Exit Sub
        
        '�������޸�ʱ������Ȩ������
        If InStr(1, strPrivs, "ҽ����ҩĿ¼") = 0 Then
            Me.cboҽ��ְ��.Enabled = False
        End If
        
        '�������޸�ʱ��װ�뱾���͵ķ���
        err = 0: On Error GoTo ErrHand
        
        '����ѡ����װ��
        gstrSql = "select ID,�ϼ�ID,����,����,����" & _
                    " From ���Ʒ���Ŀ¼" & _
                    " Where ���� = [1] and id=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, lng����id)
        
        With rsTemp
            Me.txt����.Text = "[" & rsTemp!���� & "]" & rsTemp!����
            Me.txt����.Tag = lng����id
        End With
        
        '���ݱ༭״̬��������������ʾ
        If Me.cmdCancel.Tag = "����" Then
            lngҩ��id = 0
    
            If mint������� = 0 Then
'                gstrSql = "select nvl(max(����),'0000000') as ����" & _
'                        " From ������ĿĿ¼" & _
'                        " Where ��� = '7'"

                gstrSql = "Select Nvl(Max(����), '0000000') As ����" & vbNewLine & _
                                "From (Select ���� From ������ĿĿ¼ Where ��� = '7' Order By Length(����) Desc, ���� Desc, ����ʱ�� Desc)" & vbNewLine & _
                                "Where Rownum = 1  "
    
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
                
'                Me.txt����.Text = Right(String(13, "0") & Val(rsTemp!����) + 1, Len(rsTemp!����))
                Me.txt����.Text = zlCommFun.IncStr(rsTemp!����)

            Else
                strTemp = Mid(Me.txt����.Text, 2, InStr(1, Me.txt����.Text, "]") - 2)
'                gstrSql = "select nvl(max(����),'') as ����" & _
'                        " From ������ĿĿ¼" & _
'                        " Where ��� = '7' and ���� like [1] and length(����)>=[2]"

                gstrSql = "Select Nvl(Max(����), '') As ����" & vbNewLine & _
                                "From (Select ����" & vbNewLine & _
                                "       From ������ĿĿ¼" & vbNewLine & _
                                "       Where ��� ='7' And ���� Like [1] And Length(����) >=[2] " & vbNewLine & _
                                "       Order By Length(����) Desc, ���� Desc, ����ʱ�� Desc)" & vbNewLine & _
                                "Where Rownum = 1"
                                
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, "7" & strTemp & "%", Len("*" & strTemp & "**"))
                
                err = 0: On Error Resume Next
                
                strTemp = "7" & strTemp
                If Nvl(rsTemp!����) = "" Then
                    Me.txt����.Text = strTemp & "01"
                Else
                    strCode = rsTemp!����
                    strCode = Mid(strCode, Len(strTemp) + 1)
                    strCode = zlCommFun.IncStr(strCode)
                    Me.txt����.Text = strTemp & strCode
                End If
            End If
    
            Me.txt����.Text = "": Me.txtӢ��.Text = ""
            Me.lblNote.Visible = False
            Me.txt�ο� = "": Me.txt�ο�.Tag = "": strRefer = ""
        Else
            '������Ϣ��Ŀ
            gstrSql = "select I.����ID,I.����,I.����,I.���㵥λ,I.����Ӧ��,T.�Ƿ�ԭ��,T.ҩƷ����," & _
                    "        T.�������,T.��Դ���,T.��ֵ����,T.��ҩ�ݴ�," & _
                    "        nvl(T.ҩƷ����,0) as ҩƷ����,nvl(T.����ְ��,'00') as ����ְ��,nvl(T.��������,0) as ��������," & _
                    "        nvl(T.����ҩ��,0) as ����ҩ��,nvl(T.�Ƿ�ԭ��,0) as �Ƿ�ԭ��,nvl(t.������,0) as ������,nvl(T.�Ƿ���ҩ,0) as �Ƿ���ҩ,nvl(T.�Ƿ�Ƥ��,0) as �Ƿ�Ƥ��," & _
                    "        I.����ʱ��,nvl(I.����ʱ��,to_date('3000-01-01','YYYY-MM-DD')) as ����ʱ��,B.���� as �ο�����,I.�ο�Ŀ¼id,nvl(T.Ʒ��ҽ��,0) as Ʒ��ҽ��,Nvl(I.�����Ա�,0) AS �����Ա�,Nvl(t.�Ƿ�����ҩ, 0) As �Ƿ�����ҩ" & _
                    " from ������ĿĿ¼ I,ҩƷ���� T,���Ʋο�Ŀ¼ B" & _
                    " where I.ID=T.ҩ��ID and I.ID=[1] and I.�ο�Ŀ¼id=B.id(+) "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩ��id)
            
            With rsTemp
                If Not .EOF Then
                    Me.lblNote.Caption = "ע����ҩƷ������" & Format(rsTemp!����ʱ��, "YYYY-MM-DD")
                    If Format(rsTemp!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                        Me.lblNote.Caption = Me.lblNote.Caption & "����" & Format(rsTemp!����ʱ��, "YYYY-MM-DD") & "ͣ�á�"
                    End If
                    Me.txt����.Text = rsTemp!����
                    Me.txt����.Text = rsTemp!����
                    Me.cbo��λ.Text = IIf(IsNull(rsTemp!���㵥λ), "", rsTemp!���㵥λ)
                    Me.txt�ο�.Text = Nvl(rsTemp!�ο�����)
                    Me.txt�ο�.Tag = Nvl(rsTemp!�ο�Ŀ¼ID)
                    strRefer = Me.txt�ο�.Text
    '                For intCount = 0 To Me.cbo����.ListCount - 1
    '                    If Mid(Me.cbo����.List(intCount), InStr(1, Me.cbo����.List(intCount), "-") + 1) = IIf(IsNull(!ҩƷ����), "", !ҩƷ����) Then
    '                        Me.cbo����.ListIndex = intCount: Exit For
    '                    End If
    '                Next
                    For intCount = 0 To Me.cbo����.ListCount - 1
                        If Mid(Me.cbo����.List(intCount), InStr(1, Me.cbo����.List(intCount), "-") + 1) = IIf(IsNull(rsTemp!�������), "", rsTemp!�������) Then
                            Me.cbo����.ListIndex = intCount: Exit For
                        End If
                    Next
                    For intCount = 0 To Me.cbo��ֵ.ListCount - 1
                        If Mid(Me.cbo��ֵ.List(intCount), InStr(1, Me.cbo��ֵ.List(intCount), "-") + 1) = IIf(IsNull(rsTemp!��ֵ����), "", rsTemp!��ֵ����) Then
                            Me.cbo��ֵ.ListIndex = intCount: Exit For
                        End If
                    Next
                    For intCount = 0 To Me.cbo��Դ.ListCount - 1
                        If Mid(Me.cbo��Դ.List(intCount), InStr(1, Me.cbo��Դ.List(intCount), "-") + 1) = IIf(IsNull(rsTemp!��Դ���), "", rsTemp!��Դ���) Then
                            Me.cbo��Դ.ListIndex = intCount: Exit For
                        End If
                    Next
                    For intCount = 0 To Me.cbo�ݴ�.ListCount - 1
                        If Mid(Me.cbo�ݴ�.List(intCount), InStr(1, Me.cbo�ݴ�.List(intCount), "-") + 1) = IIf(IsNull(rsTemp!��ҩ�ݴ�), "", rsTemp!��ҩ�ݴ�) Then
                            Me.cbo�ݴ�.ListIndex = intCount: Exit For
                        End If
                    Next
                    Me.cboҩƷ����.ListIndex = rsTemp!ҩƷ����
                    Me.cbo����ְ��.ListIndex = IIf(CInt(Left(Format(rsTemp!����ְ��, "00"), 1)) <> 9, CInt(Left(Format(rsTemp!����ְ��, "00"), 1)), Me.cbo����ְ��.ListCount - 1)
                    Me.cboҽ��ְ��.ListIndex = IIf(CInt(Right(Format(rsTemp!����ְ��, "00"), 1)) <> 9, CInt(Right(Format(rsTemp!����ְ��, "00"), 1)), Me.cboҽ��ְ��.ListCount - 1)
                    Me.txt��������.Text = rsTemp!��������
                    Me.chk����Ӧ��.Value = IIf(rsTemp!����Ӧ�� = 0, 0, 1)
                    Me.chkԭ��ҩ.Value = IIf(rsTemp!�Ƿ�ԭ�� = 0, 0, 1)
                    'Me.chk��ҩ.Value = IIf(!�Ƿ���ҩ = 0, 0, 1)
                    'Me.chkƤ��.Value = IIf(!�Ƿ�Ƥ�� = 0, 0, 1)
                    '���˺�:2008/03/17���뿹����,��ҪӦ����Ժ��ϵͳ�Ŀ���ҩ��ļ��:12753
                    'Me.chk������.Value = IIf(!������ = 1, 1, 0)
                    'Me.chkƷ��ҽ��.Value = IIf(!Ʒ��ҽ�� = 0, 0, 1)
                    Me.cbo�����Ա�.ListIndex = rsTemp!�����Ա�
                    Me.chk������ҩ.Value = IIf(!�Ƿ�����ҩ = 0, 0, 1)
                End If
            End With
            
            '����������Ӣ����
            gstrSql = "select ����,����,����,���� from ������Ŀ���� where ���� in (1,2) and ������ĿID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩ��id)
            
            With rsTemp
                Do While Not .EOF
                    If rsTemp!���� = 1 And rsTemp!���� = 1 Then Me.txtƴ��.Text = rsTemp!����
                    If !���� = 1 And !���� = 2 Then Me.txt���.Text = rsTemp!����
                    If rsTemp!���� = 2 Then Me.txtӢ��.Text = rsTemp!����
                    .MoveNext
                Loop
            End With
                
            '��������
            gstrSql = "select N.����,P.���� as ƴ��,W.���� as ���" & _
                    " from (select distinct ���� from ������Ŀ���� where ������ĿID=[1] and ����=9) N," & _
                    "      (select ����,���� from ������Ŀ���� where ������ĿID=[1] and ����=9 and ����=1) P," & _
                    "      (select ����,���� from ������Ŀ���� where ������ĿID=[1] and ����=9 and ����=2) W" & _
                    " where N.����=P.����(+) and N.����=W.����(+)"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩ��id)
            
            With rsTemp
                Do While Not .EOF
                    If Me.msf����.Rows - 1 < rsTemp.AbsolutePosition Then Me.msf����.Rows = Me.msf����.Rows + 1
                    Me.msf����.TextMatrix(rsTemp.AbsolutePosition, 0) = .AbsolutePosition
                    Me.msf����.TextMatrix(rsTemp.AbsolutePosition, 1) = rsTemp!����
                    Me.msf����.TextMatrix(rsTemp.AbsolutePosition, 2) = IIf(IsNull(rsTemp!ƴ��), "", rsTemp!ƴ��)
                    Me.msf����.TextMatrix(rsTemp.AbsolutePosition, 3) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
                    rsTemp.MoveNext
                Loop
            End With
        End If
        
        If Me.cmdCancel.Tag = "����" Then
            '����ʱ�����ؼ��ı༭״̬
            Me.cmdOK.Visible = False
            Me.cmdCancel.Caption = "�ر�(&C)"
            Me.txt����.Enabled = False: Me.cmd����.Enabled = False
            Me.txt����.Enabled = False
            Me.txt����.Enabled = False
            Me.txtƴ��.Enabled = False: Me.txt���.Enabled = False
            Me.txtӢ��.Enabled = False
            Me.cbo��λ.Enabled = False
            Me.cbo����.Enabled = False: Me.cbo��ֵ.Enabled = False: Me.cbo��Դ.Enabled = False: Me.cbo�ݴ�.Enabled = False
            Me.cboҩƷ����.Enabled = False: Me.cbo����ְ��.Enabled = False: Me.cboҽ��ְ��.Enabled = False: Me.txt��������.Enabled = False
            Me.chk����Ӧ��.Enabled = False: Me.chkԭ��ҩ.Enabled = False
            Me.msf����.Active = False
            Me.txt�ο�.Enabled = False
            Me.cmd�ο�.Enabled = False
            Me.cbo�����Ա�.Enabled = False
            Me.chk������ҩ.Enabled = False
        End If
    End If
    If Me.cmdCancel.Tag <> "����" Then Me.txt����.SetFocus
    mstr���м�¼ = ""
    str���� = ""
    With msf����
        For i = 1 To .Rows - 1
            str���� = str���� & "|" & .TextMatrix(i, 1) & "," & .TextMatrix(i, 2) & "," & .TextMatrix(i, 3)
        Next
    End With
    mstr���м�¼ = txt����.Text & "|" & txt����.Text & "|" & txt����.Text & "|" & txtƴ��.Text & "|" & txt���.Text & "|" & txtӢ��.Text & "|" & cbo��λ.Text & "|" & _
                txt�ο�.Text & "|" & cbo����.Text & "|" & cbo��ֵ.Text & "|" & cbo��Դ.Text & "|" & cbo�ݴ�.Text & "|" & cboҩƷ����.Text & "|" & _
                cbo����ְ��.Text & "|" & cboҽ��ְ��.Text & "|" & txt��������.Text & "|" & cbo�����Ա�.Text & "|" & chk������ҩ.Value & "|" & chk����Ӧ��.Value & "|" & chkԭ��ҩ.Value & "|" & _
                str����
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmdCancel_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    'ȡҩƷƷ�ֱ���Ĳ�������
    mint������� = Val(zlDatabase.GetPara(87, glngSys))
    
    Call GetDefineSize
    
    '-------------����ѡ������װ��-----------------------
    On Error GoTo ErrHand
    
    With rsTemp
'        gstrSql = "select ����||'-'||���� from ҩƷ���� order by ����"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
'        Me.cbo����.Clear
'        Do While Not .EOF
'            Me.cbo����.AddItem .Fields(0).Value
'            .MoveNext
'        Loop
'        If Me.cbo����.ListCount > 0 Then Me.cbo����.ListIndex = 0
    
        gstrSql = "select distinct ���㵥λ from ������ĿĿ¼ where ��� in ('5','6') and ���㵥λ is not null"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
        Do While Not rsTemp.EOF
            Me.cbo��λ.AddItem rsTemp.Fields(0).Value
            rsTemp.MoveNext
        Loop
        
        gstrSql = "select ����||'-'||���� from ҩƷ������� order by ����"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
        Me.cbo����.Clear
        Do While Not rsTemp.EOF
            Me.cbo����.AddItem rsTemp.Fields(0).Value
            If InStr(1, rsTemp.Fields(0).Value, "��ͨ") > 0 Then
                Me.cbo����.ListIndex = Me.cbo����.NewIndex
            End If
            rsTemp.MoveNext
        Loop
        If Me.cbo����.ListIndex = -1 And Me.cbo����.ListCount > 0 Then Me.cbo����.ListIndex = 0
    
        gstrSql = "select ����||'-'||���� from ҩƷ��ֵ���� order by ����"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
        Me.cbo��ֵ.Clear
        Do While Not rsTemp.EOF
            Me.cbo��ֵ.AddItem rsTemp.Fields(0).Value
            rsTemp.MoveNext
        Loop
        If Me.cbo��ֵ.ListCount > 0 Then Me.cbo��ֵ.ListIndex = 0
    
        gstrSql = "select ����||'-'||���� from ҩƷ��Դ��� order by ����"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
        Me.cbo��Դ.Clear
        Do While Not rsTemp.EOF
            Me.cbo��Դ.AddItem rsTemp.Fields(0).Value
            rsTemp.MoveNext
        Loop
        If Me.cbo��Դ.ListCount > 0 Then Me.cbo��Դ.ListIndex = 0
    
        gstrSql = "select ����||'-'||���� from ҩƷ��ҩ�ݴ� order by ����"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
        Me.cbo�ݴ�.Clear
        Do While Not rsTemp.EOF
            Me.cbo�ݴ�.AddItem rsTemp.Fields(0).Value
            rsTemp.MoveNext
        Loop
        If Me.cbo�ݴ�.ListCount > 0 Then Me.cbo�ݴ�.ListIndex = 0
    End With
    
    aryTemp = Split("0-δ�趨;1-����ҩ;2-����Ǵ���ҩ;3-����Ǵ���ҩ;4-�Ǵ���ҩ;5-������ҩ", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cboҩƷ����.AddItem aryTemp(intCount)
    Next
    Me.cboҩƷ����.ListIndex = 0
    
    aryTemp = Split("0-����;1-����;2-����;3-�м�;4-����/ʦ��;5-Ա/ʿ;9-��Ƹ", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo����ְ��.AddItem aryTemp(intCount): Me.cboҽ��ְ��.AddItem aryTemp(intCount)
    Next
    Me.cbo����ְ��.ListIndex = 0: Me.cboҽ��ְ��.ListIndex = 0
    
    aryTemp = Split("0-���Ա�����;1-����;2-Ů��", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo�����Ա�.AddItem aryTemp(intCount)
    Next
    Me.cbo�����Ա�.ListIndex = 0
    
    '��ʼ�����ñ��༭
    With Me.msf����
        .Active = True
        .MsfObj.FixedCols = 1: .Rows = 2: .Cols = 4
        .TextMatrix(0, 0) = "": .TextMatrix(0, 1) = "ҩƷ����": .TextMatrix(0, 2) = "ƴ����": .TextMatrix(0, 3) = "�����"
        .colData(0) = 5: .colData(1) = 4: .colData(2) = 4: .colData(3) = 4
        .ColWidth(0) = 250: .ColWidth(1) = 2500: .ColWidth(2) = 950: .ColWidth(3) = 950
        .TextMatrix(1, 0) = "1"
        .PrimaryCol = 1: .LocateCol = 1
        .Row = 1: .Col = 1
    End With
    
    mstrMatch = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
    strRefer = ""
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strTemp As String
    Dim i As Integer
    Dim str���� As String
    
    If mblnOK = False And mblnCancel = False Then
        With msf����
            For i = 1 To .Rows - 1
                str���� = str���� & "|" & .TextMatrix(i, 1) & "," & .TextMatrix(i, 2) & "," & .TextMatrix(i, 3)
            Next
        End With
        strTemp = txt����.Text & "|" & txt����.Text & "|" & txt����.Text & "|" & txtƴ��.Text & "|" & txt���.Text & "|" & txtӢ��.Text & "|" & cbo��λ.Text & "|" & _
                    txt�ο�.Text & "|" & cbo����.Text & "|" & cbo��ֵ.Text & "|" & cbo��Դ.Text & "|" & cbo�ݴ�.Text & "|" & cboҩƷ����.Text & "|" & _
                    cbo����ְ��.Text & "|" & cboҽ��ְ��.Text & "|" & txt��������.Text & "|" & cbo�����Ա�.Text & "|" & chk������ҩ.Value & "|" & chk����Ӧ��.Value & "|" & chkԭ��ҩ.Value & "|" & _
                    str����
        If strTemp <> mstr���м�¼ And Me.cmdCancel.Tag <> "����" Then
            If MsgBox("�����ݱ��޸���ȷ���˳���", vbYesNo, gstrSysName) = vbYes Then
                mblnLoad = False
                mblnOK = False
                mblnCancel = False
                Unload Me
            Else
                Cancel = 1
            End If
        Else
            mblnLoad = False
            mblnOK = False
            mblnCancel = False
            Unload Me
        End If
    End If
    mblnLoad = False
    mblnOK = False
    mblnCancel = False
    Unload Me
End Sub

Private Sub msf����_AfterAddRow(Row As Long)
    With Me.msf����
        For intCount = Row To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msf����_AfterDeleteRow()
    With Me.msf����
        For intCount = IIf(.Row <> 1, .Row - 1, .Row) To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msf����_EditKeyPress(KeyAscii As Integer)
    If InStr(" '|^", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub msf����_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.msf����
        If .Col = 1 Then
            If .TxtVisible = False And .TextMatrix(.Row, .Col) = "" Then Call OS.PressKey(vbKeyTab): Exit Sub
            strTemp = Trim(.Text)
            If strTemp <> "" Then
                .TextMatrix(.Row, 1) = strTemp
                .TextMatrix(.Row, 2) = zlStr.GetCodeByORCL(strTemp, False, mlng���볤��)
                .TextMatrix(.Row, 3) = zlStr.GetCodeByORCL(strTemp, True, mlng���볤��)
            Else
                Call OS.PressKey(vbKeyTab): Exit Sub
            End If
        End If
    End With
End Sub

Private Sub msf����_KeyPress(KeyAscii As Integer)
    If InStr(" '|^", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 100
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Asc("-")
        If InStr(1, txt����.Text, "-") > 0 Then
            KeyAscii = 0
        End If
        Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt�ο�_GotFocus()
    Me.txt�ο�.SelStart = 0: Me.txt�ο�.SelLength = 100
End Sub

Private Sub txt�ο�_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        If Me.txt�ο� <> strRefer Then
            Set rsTmp = SelectRefer(Trim(Me.txt�ο�))
            If rsTmp Is Nothing Then
                Me.txt�ο� = strRefer
                Me.SetFocus
                Exit Sub
            Else
                Me.txt�ο� = rsTmp("����"): Me.txt�ο�.Tag = rsTmp("ID"): strRefer = Me.txt�ο�
            End If
        End If
        Call OS.PressKey(vbKeyTab)
    End If
    If InStr(" ~!@#$%^&|=`;'""?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt�ο�_LostFocus()
    If Me.txt�ο� <> strRefer Then
        Me.txt�ο� = strRefer
    End If
End Sub

Private Sub txt��������_GotFocus()
    Me.txt��������.SelStart = 0: Me.txt��������.SelLength = 100
End Sub

Private Sub txt��������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    If InStr(1, "0123456789.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 100
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt����_Change()
    Dim strTmp As String
    '���¼�����ƣ���ȥ �������ַ�
    strTmp = MoveSpecialChar(txt����.Text)
    If txt����.Text <> strTmp Then
        txt����.Text = strTmp
    End If
    Me.txtƴ��.Text = zlStr.GetCodeByORCL(strTmp, False, mlng���볤��)
    Me.txt���.Text = zlStr.GetCodeByORCL(strTmp, True, mlng���볤��)
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 100
    Call OS.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("?")
            KeyAscii = Asc("��")
        Case Asc("%")
            KeyAscii = Asc("��")
        Case Asc("_")
            KeyAscii = Asc("��")
    End Select
    If KeyAscii = vbKeyReturn Then
        Call OS.PressKey(vbKeyTab)
    Else
        If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
        Me.txtƴ��.Text = zlStr.GetCodeByORCL(Me.txt����.Text, False, mlng���볤��)
        Me.txt���.Text = zlStr.GetCodeByORCL(Me.txt����.Text, True, mlng���볤��)
    End If
    
        
End Sub

Private Sub txt����_LostFocus()
'    Me.txtƴ��.Text = zlGetSymbol(Me.txt����.Text, 0, mlng���볤��)
'    Me.txt���.Text = zlGetSymbol(Me.txt����.Text, 1, mlng���볤��)
    Call OS.OpenIme
End Sub


Private Sub txtƴ��_GotFocus()
    Me.txtƴ��.SelStart = 0: Me.txtƴ��.SelLength = 100
End Sub

Private Sub txtƴ��_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub


Private Sub txt���_GotFocus()
    Me.txt���.SelStart = 0: Me.txt���.SelLength = 100
End Sub

Private Sub txt���_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub


Private Sub txtӢ��_GotFocus()
    Me.txtӢ��.SelStart = 0: Me.txtӢ��.SelLength = 100
End Sub

Private Sub txtӢ��_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("?")
            KeyAscii = Asc("��")
        Case Asc("%")
            KeyAscii = Asc("��")
        Case Asc("_")
            KeyAscii = Asc("��")
        Case vbKeyReturn
            Call OS.PressKey(vbKeyTab)
    End Select
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


