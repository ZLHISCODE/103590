VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMediPriceNavigation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ѡ��"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4950
   Icon            =   "frmMediPriceNavigation.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCanc 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3480
      Picture         =   "frmMediPriceNavigation.frx":000C
      TabIndex        =   11
      Top             =   3600
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2400
      Picture         =   "frmMediPriceNavigation.frx":0156
      TabIndex        =   10
      Top             =   3600
      Width           =   1100
   End
   Begin VB.Frame fra����ѡ�� 
      Caption         =   "����ѡ��ɱ��۵�����أ�"
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   4695
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshProvider 
         Height          =   1695
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Visible         =   0   'False
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   2990
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   32768
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CheckBox chk�ӳ��� 
         Caption         =   "ָ���ӳ���"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   1125
         Width           =   1215
      End
      Begin VB.CheckBox chk��Ӧ�� 
         Caption         =   "ָ����Ӧ��"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chkӦ����¼ 
         Caption         =   "�����ɱ��۵��۴�����Ӧ����������¼"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox txt�ӳ��� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   270
         Left            =   1440
         TabIndex        =   7
         Text            =   "15.0000"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txt��Ӧ�� 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton cmd��Ӧ�� 
         Caption         =   "��"
         Enabled         =   0   'False
         Height          =   270
         Left            =   4080
         TabIndex        =   5
         Top             =   350
         Width           =   375
      End
      Begin VB.Label lblComment�ӳ��� 
         Caption         =   "��ָ���ӳ��ʣ���ͳһĬ�ϰ��üӳ��ʼ���ɱ��ۣ���ָ������Ĭ����ʾʵ�ʼӳ��ʣ�"
         ForeColor       =   &H00FF0000&
         Height          =   540
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   4260
      End
      Begin VB.Label lblComment��Ӧ�� 
         AutoSize        =   -1  'True
         Caption         =   "��ָ����Ӧ�̣���ֻ�����ù�Ӧ�̵Ŀ��ҩƷ�ɱ��ۣ�"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   4320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   180
         Left            =   2415
         TabIndex        =   8
         Top             =   1125
         Width           =   90
      End
   End
   Begin VB.Frame fra���� 
      Caption         =   "��������"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.OptionButton opt���� 
         Caption         =   "���ۼۼ��ɱ���"
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "�����ɱ���"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "���ۼ�"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmMediPriceNavigation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mblnSelect As Boolean
Private mint���� As Integer             '0-���ۼ�;1-���ɱ���;2-���ۼۼ��ɱ���
Private mlng��Ӧ��ID As Long
Private mdbl�ӳ��� As Double
Private mblnӦ����¼ As Boolean         '0-������Ӧ����¼;1-����Ӧ����¼
Private mstrPrivs As String
Private Sub SetForm(ByVal int���� As Integer)
    If int���� = 0 Then
        fra����ѡ��.Visible = False
        cmdOk.Top = fra����.Top + fra����.Height + 200
        cmdCanc.Top = cmdOk.Top
    Else
        fra����ѡ��.Visible = True
        cmdOk.Top = fra����ѡ��.Top + fra����ѡ��.Height + 200
        cmdCanc.Top = cmdOk.Top
    End If
    Me.Height = cmdOk.Top + cmdOk.Height + 800
    
    If InStr(1, mstrPrivs, "�ۼ۹���") = 0 Then
        opt����(0).Visible = False
        opt����(2).Visible = False
        opt����(1).Left = opt����(0).Left
    End If
End Sub

Private Sub chk��Ӧ��_Click()
    If chk��Ӧ��.Value = 1 Then
        txt��Ӧ��.Enabled = True
        cmd��Ӧ��.Enabled = True
        chkӦ����¼.Enabled = True
    Else
        txt��Ӧ��.Enabled = False
        cmd��Ӧ��.Enabled = False
        chkӦ����¼.Value = 0
        chkӦ����¼.Enabled = False
    End If
End Sub

Private Sub chk�ӳ���_Click()
    If chk�ӳ���.Value = 1 Then
        txt�ӳ���.Enabled = True
        If Val(Trim(txt�ӳ���.Text)) = 0 Then
            txt�ӳ���.Text = "15.0000"
        End If
    Else
        txt�ӳ���.Enabled = False
    End If
End Sub

Private Sub cmdCanc_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If opt����(0).Value Then
        mint���� = 0
    ElseIf opt����(1).Value Then
        mint���� = 1
    Else
        mint���� = 2
    End If
    
    If fra����ѡ��.Visible Then
        If chk��Ӧ��.Value = 1 Then
            If Val(Split(txt��Ӧ��.Tag, "|")(0)) = 0 Then
                MsgBox "��ѡ��Ӧ�̡�", vbInformation, gstrSysName
                txt��Ӧ��.SetFocus
                Exit Sub
            End If
        End If
                
        mlng��Ӧ��ID = IIf(chk��Ӧ��.Value = 1, Val(Split(txt��Ӧ��.Tag, "|")(0)), 0)
        mdbl�ӳ��� = IIf(chk�ӳ���.Value = 1, Val(Trim(txt�ӳ���.Text)), 0)
        mblnӦ����¼ = (chkӦ����¼.Enabled And chkӦ����¼.Value = 1)
    End If
    
    mblnSelect = True
    Unload Me
End Sub

Public Function GetCondition(frmMain As Form, ByVal strPrivs As String, ByRef int���� As Integer, ByRef lng��Ӧ��ID As Long, ByRef dbl�ӳ��� As Double, ByRef blnӦ����¼ As Boolean) As Boolean
    mblnSelect = False
    mstrPrivs = strPrivs
    Set mfrmMain = frmMain
    Me.Show vbModal, frmMain
    GetCondition = mblnSelect
    
    If mblnSelect = False Then Exit Function
    
    int���� = mint����
    lng��Ӧ��ID = mlng��Ӧ��ID
    dbl�ӳ��� = mdbl�ӳ���
    blnӦ����¼ = mblnӦ����¼
End Function


Private Sub cmd��Ӧ��_Click()
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSql = "Select ����,����,����,id" & _
        " From ��Ӧ��" & _
        " where ĩ��=1 And substr(����,1,1) = '1' And (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) " & _
        " Order By ���� "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "ȡ��Ӧ����Ϣ")
    If rsTemp.EOF Then
        MsgBox "���ʼ����Ӧ�̣��ֵ������", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With Me.mshProvider
        .Left = chk��Ӧ��.Left
        .Top = txt��Ӧ��.Top + txt��Ӧ��.Height
        .Clear
        Set .DataSource = rsTemp
        .ColWidth(0) = 800: .ColWidth(1) = 2500: .ColWidth(2) = 800: .ColWidth(3) = 0
        .Row = 1: .ColSel = .Cols - 1
        .ZOrder 0: .Visible = True: .SetFocus
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    Me.txt��Ӧ��.Tag = "|"
    Call SetForm(0)
End Sub

Private Sub mshProvider_DblClick()
    With Me.mshProvider
        Me.txt��Ӧ��.Text = .TextMatrix(.Row, 1)
        Me.txt��Ӧ��.Tag = .TextMatrix(.Row, 3) & "|" & .TextMatrix(.Row, 1)
        .Visible = False
    End With
    Me.txt��Ӧ��.SetFocus
End Sub


Private Sub mshProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call mshProvider_DblClick
End Sub


Private Sub mshProvider_LostFocus()
    Me.mshProvider.Visible = False
End Sub


Private Sub opt����_Click(Index As Integer)
    SetForm (Index)
End Sub


Private Sub txt��Ӧ��_GotFocus()
    Me.txt��Ӧ��.SelStart = 0: Me.txt��Ӧ��.SelLength = Len(Me.txt��Ӧ��.Text)
End Sub


Private Sub txt��Ӧ��_KeyPress(KeyAscii As Integer)
    Dim strTmp As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
        
    strTmp = UCase(Trim(Me.txt��Ӧ��.Text))
    
    If strTmp = "" Then
        Me.txt��Ӧ��.Tag = "|"
        Exit Sub
    ElseIf strTmp = Split(Me.txt��Ӧ��.Tag, "|")(1) Then
        Exit Sub
    End If
    
    gstrSql = "Select ����,����,����,id" & _
            " From ��Ӧ��" & _
            " where (���� Like [1] " & _
            "       Or ���� Like [2] " & _
            "       Or ���� Like [2])" & _
            " And ĩ��=1 And substr(����,1,1) = '1' And (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) " & _
            " Order By ���� "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, strTmp & "%", gstrMatch & strTmp & "%")
    
    With rsTemp
        If .EOF Then
            MsgBox "û���ҵ�ƥ��Ĺ�Ӧ�̣����ڹ�Ӧ�̹��������ӹ�Ӧ�̣�", vbInformation, gstrSysName
            Me.txt��Ӧ��.Text = Split(Me.txt��Ӧ��.Tag, "|")(1)
            Me.txt��Ӧ��.SelStart = 0: Me.txt��Ӧ��.SelLength = Len(Me.txt��Ӧ��.Text)
            Exit Sub
        End If
        
        If .RecordCount = 1 Then
            Me.txt��Ӧ��.Text = Trim(rsTemp!����): Me.txt��Ӧ��.Tag = rsTemp!ID & "|" & rsTemp!����
            Exit Sub
        Else
            With Me.mshProvider
                .Left = Me.chk��Ӧ��.Left
                .Top = Me.txt��Ӧ��.Top + Me.txt��Ӧ��.Height
                .Clear
                Set .DataSource = rsTemp
                .ColWidth(0) = 800: .ColWidth(1) = 2500: .ColWidth(2) = 800: .ColWidth(3) = 0
                .Row = 1: .ColSel = .Cols - 1
                .ZOrder 0: .Visible = True: .SetFocus
            End With
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub txt��Ӧ��_Validate(Cancel As Boolean)
    If Me.txt��Ӧ��.Text = "" Then
        Me.txt��Ӧ��.Tag = "|"
    ElseIf Me.txt��Ӧ��.Text <> Split(Me.txt��Ӧ��.Tag, "|")(1) Then
        txt��Ӧ��_KeyPress (vbKeyReturn)
    End If
End Sub


Private Sub txt�ӳ���_GotFocus()
    txt�ӳ���.SelStart = 0
    txt�ӳ���.SelLength = Len(txt�ӳ���)
End Sub

Private Sub txt�ӳ���_KeyPress(KeyAscii As Integer)
    If Not (Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9 Or KeyAscii = vbKeyBack Or KeyAscii = 46) Then KeyAscii = 0
End Sub


