VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Begin VB.Form frmRegistFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdDef 
      Caption         =   "ȱʡ(&D)"
      Height          =   350
      Left            =   5880
      TabIndex        =   20
      Top             =   1560
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   4005
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   5640
      Begin VB.CheckBox chkFilter 
         Caption         =   "ԤԼ���յ���ԤԼʱ����ʾ"
         Height          =   210
         Left            =   690
         TabIndex        =   17
         Top             =   3645
         Width           =   2580
      End
      Begin VB.OptionButton optRegistRecord 
         Caption         =   "�Һż��˺ż�¼"
         Height          =   315
         Index           =   2
         Left            =   3480
         TabIndex        =   16
         Top             =   3240
         Width           =   1575
      End
      Begin VB.OptionButton optRegistRecord 
         Caption         =   "�˺ż�¼"
         Height          =   315
         Index           =   1
         Left            =   2115
         TabIndex        =   15
         Top             =   3240
         Width           =   1305
      End
      Begin VB.OptionButton optRegistRecord 
         Caption         =   "�Һż�¼"
         Height          =   315
         Index           =   0
         Left            =   690
         TabIndex        =   14
         Top             =   3240
         Value           =   -1  'True
         Width           =   1365
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   300
         Left            =   960
         TabIndex        =   12
         Top             =   2880
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         Appearance      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   12
         FontName        =   "����"
         IDKind          =   -1
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         BackColor       =   -2147483633
      End
      Begin VB.ComboBox cbo�ѱ� 
         Height          =   300
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1920
         Width           =   1830
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3585
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1920
         Width           =   1830
      End
      Begin VB.TextBox txtҽ�� 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   3585
         MaxLength       =   15
         TabIndex        =   7
         Top             =   1500
         Width           =   1830
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1500
         Width           =   1830
      End
      Begin VB.ComboBox cbo����Ա 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2400
         Width           =   1830
      End
      Begin VB.TextBox txtNOBegin 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1005
         MaxLength       =   8
         TabIndex        =   2
         Top             =   675
         Width           =   1815
      End
      Begin VB.TextBox txtNOEnd 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3585
         MaxLength       =   8
         TabIndex        =   3
         Top             =   675
         Width           =   1830
      End
      Begin VB.TextBox txtFactBegin 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1005
         TabIndex        =   4
         Top             =   1095
         Width           =   1815
      End
      Begin VB.TextBox txtFactEnd 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3585
         TabIndex        =   5
         Top             =   1095
         Width           =   1830
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   3360
         TabIndex        =   1
         Top             =   270
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   136970243
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1005
         TabIndex        =   0
         Top             =   270
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   136970243
         CurrentDate     =   36588
      End
      Begin VB.TextBox txtPatient 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   13
         Top             =   2880
         Width           =   3855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   3000
         TabIndex        =   32
         Top             =   1980
         Width           =   360
      End
      Begin VB.Label lbl�ѱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ѱ�"
         Height          =   180
         Left            =   585
         TabIndex        =   31
         Top             =   1980
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   3120
         TabIndex        =   30
         Top             =   1155
         Width           =   180
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   3120
         TabIndex        =   29
         Top             =   735
         Width           =   180
      End
      Begin VB.Label lblData_ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   3120
         TabIndex        =   28
         Top             =   330
         Width           =   180
      End
      Begin VB.Label lblҽ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ��"
         Height          =   180
         Left            =   3030
         TabIndex        =   27
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   225
         TabIndex        =   11
         Top             =   2925
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ݺ�"
         Height          =   180
         Left            =   405
         TabIndex        =   26
         Top             =   735
         Width           =   540
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Һ�ʱ��"
         Height          =   180
         Left            =   225
         TabIndex        =   25
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   585
         TabIndex        =   24
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label lbl����Ա 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Һ�Ա"
         Height          =   180
         Left            =   405
         TabIndex        =   23
         Top             =   2460
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ʊ�ݺ�"
         Height          =   180
         Left            =   405
         TabIndex        =   22
         Top             =   1155
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5880
      TabIndex        =   19
      Top             =   600
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5880
      TabIndex        =   18
      Top             =   120
      Width           =   1100
   End
End
Attribute VB_Name = "frmRegistFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit    'Ҫ���������
Private mbytType As Byte     '0-�Һ��嵥����,1-ԤԼ�嵥����,2-�����嵥����
Public mlngModule As Long
Public mstrFilter As String
Public mstrSectName As String   '����ָ����ǰĬ�ϵĿ���
Public mblnDateMoved As Boolean    'Out
Private mstrCardStr As String    '�����������õĿ�
Private Const mstrIDKind = "1-����;2-���￨;3-�����;4-ҽ����;5-���֤��;6-IC����"

Private mblnNotClick As Boolean
Private mblnUnChange As Boolean
Private mrsInfo As ADODB.Recordset
Private mbln����סԺ���˹Һ� As Boolean
Private mblnOlnyBJYB As Boolean
Public mlngPrePatient As Long
Private mblnKeyReturn As Boolean
Private mblnValid As Boolean
Private Sub cbo����Ա_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo����Ա.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo����Ա.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo����Ա.ListIndex = lngIdx
    If cbo����Ա.ListIndex = -1 And cbo����Ա.ListCount <> 0 Then cbo����Ա.ListIndex = 0
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo����.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo����.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo����.ListIndex = lngIdx
    If cbo����.ListIndex = -1 And cbo����.ListCount <> 0 Then cbo����.ListIndex = 0
End Sub

Private Sub cmdCancel_Click()
    gblnOk = False
    Hide
End Sub

Private Sub cmdDef_Click()
    Form_Load
End Sub



Private Sub cmdOK_Click()
    If Not IsNull(dtpEnd.Value) Then
        If dtpEnd.Value < dtpBegin.Value Then
            MsgBox "����ʱ�䲻��С�ڿ�ʼʱ�䣡", vbInformation, gstrSysName
            dtpEnd.SetFocus: Exit Sub
        End If
    End If
    If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
        If txtNoEnd.Text < txtNOBegin.Text Then
            MsgBox "�������ݺŲ���С�ڿ�ʼ���ݺţ�", vbInformation, gstrSysName
            txtNoEnd.SetFocus: Exit Sub
        End If
    End If
    If txtFactBegin.Text <> "" And txtFactEnd.Text <> "" Then
        If txtFactEnd.Text < txtFactBegin.Text Then
            MsgBox "����Ʊ�ݺŲ���С�ڿ�ʼƱ�ݺţ�", vbInformation, gstrSysName
            txtFactEnd.SetFocus: Exit Sub
        End If
    End If
    '74237:ԤԼʱ��Ĳ�ѯ��Χ��ʾ
    If mbytType = 1 Then
        If dtpEnd.Value - dtpBegin.Value > gintԤԼ���� + 1 Then
            If MsgBox("��ǰԤԼʱ�䷶Χ����(����" & gintԤԼ���� & "��),���ܻᵼ�¶�ȡ�ͼ���ʱ�����,���Ƿ���Ҫ������ѯ?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
    Call MakeFilter

    gblnOk = True
    Hide
End Sub

Private Sub Form_Activate()
    txtNOBegin.SetFocus
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Not ActiveControl Is txtPatient Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr(1, "'[]", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = 13 And Not ActiveControl Is txtPatient Then KeyAscii = 0
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim Curdate As Date, i As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    On Error GoTo errH

    gblnOk = False

    txtNOBegin.Text = ""
    txtNoEnd.Text = ""
    txtFactBegin.Text = ""
    txtFactEnd.Text = ""
    optRegistRecord(0).Value = True
    txtPatient.Text = ""
    txtҽ��.Text = ""
    '47928
    InitIDKind
    'dtpBegin.Enabled = mbytType <> 2:����44946
    'dtpEnd.Enabled = mbytType <> 2:����44946
    txtFactBegin.Enabled = mbytType = 0
    txtFactEnd.Enabled = mbytType = 0
    txtFactBegin.BackColor = IIf(mbytType = 0, txtPatient.BackColor, Me.BackColor)
    txtFactEnd.BackColor = IIf(mbytType = 0, txtPatient.BackColor, Me.BackColor)
    optRegistRecord(0).Enabled = mbytType = 0 Or mbytType = 1
    optRegistRecord(1).Enabled = mbytType = 0 Or mbytType = 1
    optRegistRecord(2).Enabled = mbytType = 0 Or mbytType = 1
    dtpEnd.MinDate = CDate("1905-01-01")
    dtpBegin.MinDate = CDate("1905-01-01")
    Curdate = zlDatabase.Currentdate
    If mbytType = 0 Then    '�Һ�
        lblDate.Caption = "�Һ�ʱ��"
        'ȱʡʱ��Ϊ������
        dtpBegin.Value = Format(Curdate, "yyyy-MM-dd 00:00:00")
        dtpEnd.Value = Format(Curdate, "yyyy-MM-dd 23:59:59")
    ElseIf mbytType = 1 Then    'ԤԼ
        'ȱʡΪԤԼʱ��δʧЧ�ĵ���
        lblDate.Caption = "ԤԼʱ��"
        dtpBegin.Value = Format(Curdate, "yyyy-MM-dd 00:00:00")
        dtpEnd.Value = Format(Curdate + gintԤԼ����, "yyyy-MM-dd 23:59:59")
    ElseIf mbytType = 2 Then    '����
        '����Ҫ����ʱ��
        lblDate.Caption = "ԤԼʱ��"
        dtpBegin.Value = Format(Curdate, "yyyy-MM-dd 00:00:00")
        dtpEnd.Value = Format(Curdate, "yyyy-MM-dd 23:59:59")
        dtpEnd.MinDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
        dtpBegin.MinDate = dtpEnd.MinDate
    End If

    '�Һ�Ա
    cbo����Ա.Clear
    cbo����Ա.AddItem "���йҺ�Ա"
    cbo����Ա.ListIndex = 0

    Set rsTmp = GetPersonnel("����Һ�Ա", True)
    If rsTmp.RecordCount > 0 Then
        For i = 1 To rsTmp.RecordCount
            cbo����Ա.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!id = UserInfo.id Then cbo����Ա.ListIndex = cbo����Ա.NewIndex
            rsTmp.MoveNext
        Next
    End If

    '�Һſ���
    Set rsTmp = GetDepartments("'�ٴ�'", "1,3")
    cbo����.Clear
    cbo����.AddItem "���п���"
    cbo����.ListIndex = 0

    Do While Not rsTmp.EOF
        cbo����.AddItem rsTmp!���� & "-" & rsTmp!����
        cbo����.ItemData(cbo����.NewIndex) = rsTmp!id
        If mstrSectName = rsTmp!���� Then cbo����.ListIndex = cbo����.NewIndex
        rsTmp.MoveNext
    Loop

    '�ѱ�
    cbo�ѱ�.Clear
    cbo�ѱ�.AddItem "���зѱ�"
    cbo�ѱ�.ListIndex = 0
    strSQL = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From �ѱ� Where Nvl(�������,3) IN(1,3) Order by ����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo�ѱ�.AddItem rsTmp!���� & "-" & rsTmp!����
            rsTmp.MoveNext
        Next
    End If

    '����
    cbo����.Clear
    cbo����.AddItem "���к���"
    cbo����.ListIndex = 0
    strSQL = "Select ����,����,����,ȱʡ��־,˵�� From ���� Order by ����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        cbo����.AddItem rsTmp!����
        rsTmp.MoveNext
    Next
    mbln����סԺ���˹Һ� = zlDatabase.GetPara("����סԺ���˹Һ�", glngSys, mlngModule, 0) = "1"
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optRegistRecord_Click(index As Integer)
    If optRegistRecord(1).Value = True Then
        lbl����Ա.Caption = "�˺�Ա"
        lblDate.Caption = "�˺�ʱ��"
    Else
        lbl����Ա.Caption = "�Һ�Ա"
        lblDate.Caption = "�Һ�ʱ��"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbytType = 0
    Set mrsInfo = Nothing
    mlngPrePatient = 0
    IDKind.SetAutoReadCard False
End Sub

Private Sub txtFactBegin_GotFocus()
    zlControl.TxtSelAll txtFactBegin
End Sub

Private Sub txtFactBegin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFactEnd_GotFocus()
    zlControl.TxtSelAll txtFactEnd
End Sub

Private Sub txtFactEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFactBegin_Change()
    txtFactEnd.Enabled = Not (Trim(txtFactBegin.Text) = "")
    If Trim(txtFactBegin.Text = "") Then txtFactEnd.Text = ""
End Sub

Private Sub txtPatient_Change()
    txtPatient.Tag = "": mlngPrePatient = 0
    If Me.ActiveControl Is txtPatient Then
        'If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
       ' If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard txtPatient.Text = ""
    End If
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    Dim blnCard As Boolean, blnICCard As Boolean

    On Error GoTo errH
    If txtPatient.Locked Then Exit Sub
    mblnKeyReturn = KeyAscii = 13
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub

    If IDKind.GetCurCard.���� Like "����*" Then
        blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
    ElseIf IDKind.IDKind = IDKind.GetKindIndex("�����") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not (IsNumeric(Chr(KeyAscii)) Or Chr(KeyAscii) = "-") Then KeyAscii = 0: Exit Sub
        End If
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
    End If
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        Call FindPati(IDKind.GetCurCard, blnCard, txtPatient.Text)
    End If
    If Me.ActiveControl Is txtPatient And mblnKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog    '
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ҳ���
    '����:���˺�
    '����:2012-08-29 17:53:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnICCard As Boolean, blnIDCard As Boolean
   '��ȡ������Ϣ
    Call GetPatient(objCard, txtPatient.Text, blnCard)
End Sub
Private Sub txtNOBegin_Change()
    txtNoEnd.Enabled = Not (Trim(txtNOBegin.Text) = "")
    If Trim(txtNOBegin.Text = "") Then txtNoEnd.Text = ""
End Sub

Private Sub txtNOBegin_GotFocus()
    zlControl.TxtSelAll txtNOBegin
End Sub

Private Sub txtNOBegin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '46512
    zlControl.TxtCheckKeyPress txtNOBegin, KeyAscii, m�ı�ʽ
End Sub

Private Sub txtNOBegin_LostFocus()
    If txtNOBegin.Text <> "" Then txtNOBegin.Text = GetFullNO(txtNOBegin.Text, 12)
End Sub


Private Sub txtNOEnd_LostFocus()
    If txtNoEnd.Text <> "" Then txtNoEnd.Text = GetFullNO(txtNoEnd.Text, 12)
End Sub

Private Sub txtNoEnd_GotFocus()
    zlControl.TxtSelAll txtNoEnd
End Sub


Private Sub txtNoEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '46512
    zlControl.TxtCheckKeyPress txtNoEnd, KeyAscii, m�ı�ʽ
End Sub

Private Sub MakeFilter()
    Dim strSQL As String
    Dim strSQLtmp As String

    mstrFilter = " And 1=1"

    If mbytType = 0 Then    '�Һ�
        If chkFilter.Value = 0 Then
            mstrFilter = " And A.�Ǽ�ʱ�� Between [1] And [2]"
        Else
            mstrFilter = " And A.����ʱ�� Between [1] And [2]"
        End If
    ElseIf mbytType = 1 Then    'ԤԼ
        mstrFilter = " And A.����ʱ�� Between [1] And [2]"
    ElseIf mbytType = 2 Then    '����
        '����Ҫ����ʱ������
    End If

    mblnDateMoved = zlDatabase.DateMoved(Format(IIf(dtpBegin.Value < dtpEnd.Value, dtpBegin.Value, dtpEnd.Value), dtpBegin.CustomFormat), , , Me.Caption)

    If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
        mstrFilter = mstrFilter & " And A.NO Between [3] And [4]"
    ElseIf txtNOBegin.Text <> "" Then
        mstrFilter = mstrFilter & " And A.NO=[3]"
    ElseIf txtNoEnd.Text <> "" Then
        mstrFilter = mstrFilter & " And A.NO=[4]"
    End If

    If cbo����Ա.ListIndex > 0 Then mstrFilter = mstrFilter & " And A.����Ա����||''=[5]"

    If txtPatient.Text <> "" And mlngPrePatient <> 0 And Not mrsInfo Is Nothing Then
        If Val(Nvl(mrsInfo!id)) = mlngPrePatient Then
            mstrFilter = mstrFilter & " And D.����ID=[6]"
        End If
    ElseIf txtPatient.Text <> "" And mrsInfo Is Nothing Then
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(txtPatient.Text, 1))) > 0 Then
            mstrFilter = mstrFilter & " And Upper(A.����) Like [13]"
            txtPatient.Text = UCase(txtPatient.Text)
        Else
            mstrFilter = mstrFilter & " And A.���� Like [13]"
        End If
    End If


    If txtҽ��.Text <> "" Then
        mstrFilter = mstrFilter & " And A.ִ���� Like [7]"
    End If

    If (txtFactBegin.Text <> "" And txtFactEnd.Text <> "") Or (txtFactBegin.Text <> "" And txtFactEnd.Text = "") Then
        '�������Ʊ�ݺ��ж�,ֱ�Ӹ��ݵ��ݵĵǼ�ʱ���ж�
        strSQLtmp = IIf(txtFactEnd.Text = "", " =[8] ", " Between [8] And [9] ")

        If mblnDateMoved Then
            strSQL = "(Select A.NO" & _
                   " From Ʊ�ݴ�ӡ���� A,Ʊ��ʹ����ϸ B" & _
                   " Where A.��������=4 And A.ID=B.��ӡID And B.����=1" & _
                   " And B.���� " & strSQLtmp & ") Union All" & _
                   " (Select A.NO " & _
                   " From HƱ�ݴ�ӡ���� A,HƱ��ʹ����ϸ B" & _
                   " Where A.��������=4 And A.ID=B.��ӡID And B.����=1" & _
                   " And B.���� " & strSQLtmp & ")"
        Else
            strSQL = "Select A.NO" & _
                   " From Ʊ�ݴ�ӡ���� A,Ʊ��ʹ����ϸ B" & _
                   " Where A.��������=4 And A.ID=B.��ӡID And B.����=1" & _
                   " And B.���� " & strSQLtmp
        End If
    End If

    If strSQL <> "" Then mstrFilter = mstrFilter & " And A.NO IN(" & strSQL & ")"

    '�Һſ���(ִ�п���)
    If cbo����.ListIndex > 0 Then
        mstrFilter = mstrFilter & " And A.ִ�в���ID+0=[10]"
    End If

    If cbo�ѱ�.ListIndex > 0 Then
        mstrFilter = mstrFilter & " And (F.�ѱ� = [11] or F.�ѱ� is Null)"
    End If

    If cbo����.ListIndex > 0 Then
        mstrFilter = mstrFilter & " And B.���� = [12]"
    End If

End Sub

Private Sub txtPatient_GotFocus()
    Call zlControl.TxtSelAll(txtPatient)
    Call zlCommFun.OpenIme(True)
    If txtPatient.Text = "" And ActiveControl Is txtPatient Then
'        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
'        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard txtPatient.Text = ""
    End If
End Sub

Private Sub txtPatient_LostFocus()
    Call zlCommFun.OpenIme
    IDKind.SetAutoReadCard False
End Sub

 

Private Sub txtҽ��_GotFocus()
    Call zlControl.TxtSelAll(txtҽ��)
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtҽ��_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txtҽ��_Validate(Cancel As Boolean)
    Dim strDoctor As String
    strDoctor = UCase(Trim(txtҽ��.Text))
    If strDoctor <> "" Then
        If zlCommFun.IsNumOrChar(strDoctor) Then
            strDoctor = GetDoctorName(strDoctor)
            If strDoctor = "" Then Cancel = True
        End If
    End If
    txtҽ��.Text = strDoctor
End Sub

Private Function GetDoctorName(ByVal strCode As String) As String
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strIF As String, lngDept As Long, blnCancel As Boolean, vRect As RECT

    On Error GoTo Hd
    If zlCommFun.IsCharAlpha(strCode) Then
        strIF = " And A.���� Like [1]"
        strCode = strCode & "%"
    Else
        strIF = " And (A.���� = [1] Or A.��� = [1])"
    End If
    If cbo����.ListIndex > 0 Then
        strIF = strIF & " And B.����ID = [2]"
        lngDept = cbo����.ItemData(cbo����.ListIndex)
    End If
    strSQL = "Select Distinct A.Id,A.���� From ��Ա�� A, ������Ա B,��Ա����˵�� C" & vbCrLf & _
             "Where A.id=B.��Աid And A.id=C.��Աid  And C.��Ա����='ҽ��'" & strIF

    vRect = zlControl.GetControlRect(txtҽ��.Hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ѡ��ҽ��", 1, "", "��ѡ��ҽ��", False, False, True, vRect.Left, vRect.Top, txtҽ��.Height, blnCancel, False, True, strCode, lngDept)
    If Not rsTmp Is Nothing Then
        GetDoctorName = rsTmp!����
    End If
    Exit Function
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function

'��ʼ��IDKIND
Private Function InitIDKind() As Boolean
    Dim objCard As Card, rsTmp As ADODB.Recordset
    Dim lngCardID As Long, strSQL As String
    Call IDKind.zlInit(Me, glngSys, mlngModule, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    lngCardID = Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, mlngModule, 0))
    '72936:������,2014-05-13,ȱʡ�������ͱ�ͣ�ú󱨴������
    If lngCardID <> 0 Then
        strSQL = "Select 1 From ҽ�ƿ���� Where ID=[1] And Nvl(�Ƿ�����,0)=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngCardID)
        If Not rsTmp.EOF Then IDKind.DefaultCardType = lngCardID
    End If
    Set objCard = IDKind.GetfaultCard
    If IDKind.Cards.��ȱʡ������ And Not objCard Is Nothing Then
        gobjSquare.blnȱʡ�������� = objCard.�������Ĺ��� <> ""
        gobjSquare.intȱʡ���ų��� = objCard.���ų���
        Set gobjSquare.objDefaultCard = objCard

    Else
        gobjSquare.blnȱʡ�������� = IDKind.Cards.������ʾ
        gobjSquare.intȱʡ���ų��� = 100
    End If
End Function
'��ȡĬ��IDKind����
Private Function IDKindDefaultKind() As Long
    Dim lngIndex As Long
    'IDkind��Ĭ��Kind
    If IDKind.DefaultCardType = "" Then
        lngIndex = -1
    Else
        If IsNumeric(IDKind.DefaultCardType) Then
            lngIndex = IDKind.GetKindIndex(IDKind.GetfaultCard.����)
        Else
            lngIndex = IDKind.GetKindIndex(IDKind.DefaultCardType)
        End If
    End If
    IDKindDefaultKind = lngIndex
End Function


'�ؼ������Ƿ�ƥ��
Private Function IsCardType(ByVal IDKindCtl As IDKindNew, ByVal strCardName As String) As Boolean
    If IDKindCtl Is Nothing Then Exit Function
    If UCase(TypeName(IDKindCtl)) <> "IDKINDNEW" Then Exit Function
    Select Case strCardName
    Case "����", "��������￨"
        IsCardType = IDKindCtl.GetCurCard.���� Like "����*"
    Case "���֤", "���֤��", "�������֤"
        IsCardType = IDKindCtl.GetCurCard.���� Like "*���֤*"
    Case "IC����", "IC��"
        IsCardType = IDKindCtl.GetCurCard.���� Like "IC��*"
    Case "ҽ����"
        IsCardType = IDKindCtl.GetCurCard.���� = "ҽ����"
    Case "�����"
        IsCardType = IDKindCtl.GetCurCard.���� = "�����"
    Case Else
        If IDKindCtl.GetCurCard Is Nothing Then Exit Function
        If Not IsNumeric(strCardName) Or Val(strCardName) <= 0 Then Exit Function
        If IDKindCtl.GetCurCard.�ӿ���� <= 0 Then Exit Function
        IsCardType = IDKindCtl.GetCurCard.�ӿ���� = Val(strCardName)
    End Select
End Function

Private Sub IDKind_ItemClick(index As Integer, objCard As zlIDKind.Card)
    Set gobjSquare.objCurCard = objCard
    If txtPatient.Text <> "" Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
'        If mobjICCard Is Nothing Then
'            Set mobjICCard = CreateObject("zlICCard.clsICCard")
'            Set mobjICCard.gcnOracle = gcnOracle
'        End If
'        If mobjICCard Is Nothing Then Exit Sub
'        txtPatient.Text = mobjICCard.Read_Card()
'        If txtPatient.Text <> "" Then
'            Call FindPati(objCard, True, txtPatient.Text)
'        End If
        Exit Sub
    End If
    
   lng�����ID = objCard.�ӿ����
    If lng�����ID <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '����:�����ӿ�
    '    '���:frmMain-���õĸ�����
    '    '       lngModule-���õ�ģ���
    '    '       strExpand-��չ����,������
    '    '       blnOlnyCardNO-������ȡ����
    '    '����:strOutCardNO-���صĿ���
    '    '       strOutPatiInforXML-(������Ϣ����.XML��)
    '    '����:��������    True:���óɹ�,False:����ʧ��\
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModule, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then
        Call FindPati(objCard, True, txtPatient.Text)
    End If
End Sub
Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.����
    Call FindPati(objCard, True, txtPatient.Text)
End Sub
 


Private Sub GetPatient(ByVal objCard As Card, ByVal strInput As String, Optional blnCard As Boolean)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡ������Ϣ
    '��Σ�blnCard=�Ƿ���￨ˢ��
    '���ƣ����˺�
    '���ڣ�2010-07-16 14:24:14
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strTemp As String
    Dim blnSame As Boolean, blnCancel As Boolean
    Dim cur��� As Currency, curMoney As Currency
    Dim i As Integer, strPati As String
    Dim vRect As RECT, str����Ժ As String
    Dim strSQL As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim strTmp As String
    Dim lng����ID As Long, blnHavePassWord As Boolean

    On Error GoTo errH
    If Not mbln����סԺ���˹Һ� Then
        str����Ժ = " And Not Exists(Select 1 From ������ҳ Where ����ID=B.����ID And ��ҳID=B.��ҳID And Nvl(��������,0)=0 And ��Ժ���� is Null)"
    End If

    strSQL = ""
    If blnCard = True And objCard.���� Like "����*" Then    'ˢ��
        If IDKind.Cards.��ȱʡ������ And Not IDKind.GetfaultCard Is Nothing Then
            lng�����ID = IDKind.GetfaultCard.�ӿ����
        Else
            lng�����ID = "-1"
        End If
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng����ID
        blnHavePassWord = True
        strSQL = strSQL & " And B.����ID=[2] " & str����Ժ
        
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then
        '�����
        strSQL = strSQL & " And B.�����=[2]" & str����Ժ
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then
        '����ID
        strSQL = strSQL & " And B.����ID=[2]" & str����Ժ
    Else
        Select Case objCard.����
        Case "����", "��������￨"
            txtPatient.Tag = strInput
            Set mrsInfo = Nothing: Exit Sub
            zlCommFun.PressKey vbKeyTab
        Case "ҽ����"
            strInput = UCase(strInput)
            If mblnOlnyBJYB And zlCommFun.ActualLen(strInput) >= 9 Then
                '������ҽ������Ч:������:����:26982
                strSQL = strSQL & " And B.ҽ���� like [3] " & str����Ժ
                strTemp = Left(strInput, 9) & "%"
            Else
                strSQL = strSQL & " And B.ҽ����=[1]" & str����Ժ
            End If
        Case "���֤��", "���֤", "�������֤"
            strInput = UCase(strInput)
            If gobjSquare.objSquareCard.zlGetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
            strSQL = strSQL & " And B.����ID=[2]" & str����Ժ
            strInput = "-" & lng����ID
        Case "IC����", "IC��"
            strInput = UCase(strInput)
            If gobjSquare.objSquareCard.zlGetPatiID("IC��", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
            strSQL = strSQL & " And B.����ID=[2]" & str����Ժ
            strInput = "-" & lng����ID
        Case "�����"
            If Not IsNumeric(strInput) Then strInput = "0"
            strSQL = strSQL & " And B.�����=[1]" & str����Ժ
        Case Else
            '��������,��ȡ��صĲ���ID
            If Val(objCard.�ӿ����) > 0 Then
                lng�����ID = Val(objCard.�ӿ����)
                If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                If lng����ID = 0 Then lng����ID = 0
            Else
                If gobjSquare.objSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, _
                                                        strPassWord, strErrMsg) = False Then lng����ID = 0
            End If
            If lng����ID <= 0 Then lng����ID = 0
            strSQL = strSQL & " And B.����ID=[2]" & str����Ժ
            strInput = "-" & lng����ID
            blnHavePassWord = True
        End Select
    End If
    strSQL = "" & _
    "   Select distinct  B.����id As ID, Decode(sign(nvl(X.����id,0)),0,'','��') as �����˻�,  " & _
    "           B.����id,B.����, B.�Ա�, B.����, B.�����, B.��������, B.���֤��, B.��ͥ��ַ, B.������λ," & _
    "            A.���� ��������" & _
    "   From ������Ϣ B, ������� A,ҽ�ƿ���� Y,����ҽ�ƿ���Ϣ X" & _
    "   Where B.���� = A.���(+) and b.����id=X.����id(+)  " & _
    "               And X.״̬(+)=0 and  X.�����id=Y.id(+)  and Y.�Ƿ�����(+)=0 And B.ͣ��ʱ�� Is Null   " & _
                    strSQL
    On Error GoTo errH
    vRect = zlControl.GetControlRect(txtPatient.Hwnd)
    Set mrsInfo = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "���˲���", 1, "��", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput, CStr(Mid(strInput, 2)), strInput & "%", dtpBegin.Value, dtpEnd.Value)
    
    If blnCancel Or mrsInfo Is Nothing Then
        Set mrsInfo = Nothing: txtPatient.Text = "": Exit Sub
    End If
    
    If mrsInfo!id = 0 Then    'û���ҵ�������Ϣ
        Set mrsInfo = Nothing: txtPatient.Text = "": Exit Sub
    End If
    
    txtPatient.MaxLength = zlGetPatiInforMaxLen.intPatiName
    txtPatient.Text = Nvl(mrsInfo!����)
    Me.txtPatient.Tag = Nvl(mrsInfo!id)
    mlngPrePatient = Val(Nvl(mrsInfo!id))
    zlCommFun.PressKey vbKeyTab
    Exit Sub
    
NotFoundPati:
    Set mrsInfo = Nothing: txtPatient.Text = "": Exit Sub
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



'���˷�ʽ
Public Property Let bytType(ByVal vNewValue As Byte)
    mbytType = vNewValue
    chkFilter.Visible = mbytType = 0
End Property
