VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPACSFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   ControlBox      =   0   'False
   Icon            =   "frmPACSFilter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboPart 
      Height          =   300
      Left            =   1170
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   3540
      Width           =   3870
   End
   Begin VB.CommandButton cmdClear 
      Height          =   300
      Left            =   4725
      Picture         =   "frmPACSFilter.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "������������б�"
      Top             =   4305
      Width           =   300
   End
   Begin VB.ComboBox cboContent 
      Height          =   300
      Left            =   1170
      TabIndex        =   13
      Top             =   4305
      Width           =   3570
   End
   Begin VB.ComboBox cboItem 
      Height          =   300
      Left            =   1170
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   3915
      Width           =   3870
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   3
      Left            =   -15
      TabIndex        =   30
      Top             =   4665
      Width           =   5775
   End
   Begin VB.OptionButton optAdviceTime 
      Caption         =   "������ʱ�����&A(ֱ����ҽ��վ����ɵģ����ô������)"
      Height          =   180
      Left            =   75
      TabIndex        =   5
      Top             =   1995
      Width           =   5130
   End
   Begin VB.OptionButton optCheckTime 
      Caption         =   "�����ʱ�����&T(�Ƽ�)"
      Height          =   270
      Left            =   75
      TabIndex        =   4
      Top             =   1680
      Value           =   -1  'True
      Width           =   2490
   End
   Begin VB.TextBox txtChkNO 
      Height          =   300
      Left            =   1170
      MaxLength       =   10
      TabIndex        =   8
      Top             =   2655
      Width           =   795
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1170
      TabIndex        =   2
      Top             =   1215
      Width           =   1185
   End
   Begin VB.TextBox txtNO 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3420
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1215
      Width           =   1185
   End
   Begin VB.TextBox txt���￨ 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3420
      MaxLength       =   10
      TabIndex        =   1
      Top             =   855
      Width           =   1185
   End
   Begin VB.TextBox txt��ʶ�� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1170
      MaxLength       =   10
      TabIndex        =   0
      Top             =   855
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   2
      Left            =   0
      TabIndex        =   24
      Top             =   1620
      Width           =   5775
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   22
      Top             =   720
      Width           =   5775
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   0
      Left            =   -105
      TabIndex        =   21
      Top             =   3435
      Width           =   5775
   End
   Begin VB.CommandButton cmdDefault 
      Cancel          =   -1  'True
      Caption         =   "ȱʡ(&D)"
      Height          =   350
      Left            =   330
      TabIndex        =   20
      Top             =   4785
      Width           =   1100
   End
   Begin VB.CheckBox chk��Դ 
      Caption         =   "סԺ����"
      Height          =   195
      Index           =   1
      Left            =   4020
      TabIndex        =   10
      Top             =   2730
      Value           =   1  'Checked
      Width           =   1020
   End
   Begin VB.CheckBox chk��Դ 
      Caption         =   "���ﲡ��"
      Height          =   195
      Index           =   0
      Left            =   2910
      TabIndex        =   9
      Top             =   2715
      Value           =   1  'Checked
      Width           =   1020
   End
   Begin VB.ComboBox cboDept 
      Height          =   300
      Left            =   1170
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3060
      Width           =   3870
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   300
      Left            =   3210
      TabIndex        =   7
      Top             =   2295
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   84738051
      CurrentDate     =   38082
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   300
      Left            =   1170
      TabIndex        =   6
      Top             =   2295
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   84738051
      CurrentDate     =   38082
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2775
      TabIndex        =   15
      Top             =   4785
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3945
      TabIndex        =   16
      Top             =   4785
      Width           =   1100
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "��鲿λ"
      Height          =   180
      Left            =   330
      TabIndex        =   33
      Top             =   3600
      Width           =   720
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Left            =   330
      TabIndex        =   32
      Top             =   4365
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "������Ŀ"
      Height          =   180
      Left            =   330
      TabIndex        =   31
      Top             =   3975
      Width           =   720
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   285
      Left            =   510
      TabIndex        =   29
      Top             =   2745
      Width           =   705
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����(&3)"
      Height          =   180
      Left            =   510
      TabIndex        =   28
      Top             =   1275
      Width           =   630
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ݺ�(&4)"
      Height          =   180
      Left            =   2595
      TabIndex        =   27
      Top             =   1275
      Width           =   810
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���￨(&2)"
      Height          =   180
      Left            =   2610
      TabIndex        =   25
      Top             =   915
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ʶ��(&1)"
      Height          =   180
      Left            =   330
      TabIndex        =   26
      Top             =   915
      Width           =   810
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   270
      Picture         =   "frmPACSFilter.frx":0596
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    ͨ���������������Ա�׼ȷ����ִ�м�¼������ʱ�䷶Χ������ȷ���Բ��ұ�֤Ч�ʡ�"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   915
      TabIndex        =   23
      Top             =   180
      Width           =   4035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������Դ"
      Height          =   180
      Left            =   2070
      TabIndex        =   19
      Top             =   2745
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���˿���"
      Height          =   180
      Left            =   330
      TabIndex        =   18
      Top             =   3120
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ʱ�䷶Χ                      ��"
      Height          =   180
      Left            =   330
      TabIndex        =   17
      Top             =   2355
      Width           =   2880
   End
End
Attribute VB_Name = "frmPACSFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrFilter As String
Public mstrPati As String
Private mblnLoad As Boolean
Public mBeforeDays As Integer 'Ĭ�ϲ�ѯ������
Public mblnOK As Boolean
Public FindType As Integer 'ʱ����ҷ�ʽ��1�������ʱ�䡢2��������ʱ��

Private Sub cboContent_LostFocus()
    Dim i As Integer
    With cboContent
        If Len(Trim(.Text)) = 0 Then Exit Sub
        
        For i = 0 To .ListCount - 1
            If Trim(.Text) = .List(i) Then Exit For
        Next
        If i > .ListCount - 1 Then .AddItem .Text
    End With
End Sub

Private Sub chk��Դ_Click(Index As Integer)
    If chk��Դ(0).Value = 0 And chk��Դ(1).Value = 0 Then
        chk��Դ((Index + 1) Mod 2).Value = 1
    End If
    Call LoadDept
End Sub

Private Sub cmdCancel_Click()
    mstrFilter = ""
    mstrPati = ""
    mblnOK = False
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPACSFilter", "���˱�����Ŀ", cboItem.Text)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmPACSFilter", "���˲��˿���", cboDept.Text)
    Me.Hide
End Sub

Private Sub cmdClear_Click()
    cboContent.Clear
End Sub

Private Sub cmdDefault_Click()
    Me.optCheckTime.Value = True: mBeforeDays = 2
    Call Form_Load
End Sub

Private Sub cmdOK_Click()
    Call txtNO_Validate(False)
    Call MakeFilter(mstrFilter, mstrPati)
    
    mBeforeDays = dtpEnd.Value - dtpBegin.Value
    FindType = IIf(Me.optCheckTime.Value, 1, 2)
    mblnOK = True
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPACSFilter", "���˱�����Ŀ", cboItem.Text)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmPACSFilter", "���˲��˿���", cboDept.Text)
    Me.Hide
End Sub

Private Sub MakeFilter(strFilter As String, strPati As String)
'���ܣ���������(����ҽ������ A,����ҽ����¼ B)
    Dim strTmp As String
    
    '����ʱ��
    If Format(dtpEnd.Value, "yyyy-MM-dd HH:mm") = Format(dtpEnd.Tag, "yyyy-MM-dd HH:mm") Then
        strFilter = " And A.����ʱ�� Between To_Date('" & Format(dtpBegin.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS') And Sysdate"
    Else
        strFilter = " And A.����ʱ�� Between To_Date('" & Format(dtpBegin.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')" & _
            " And To_Date('" & Format(dtpEnd.Value, "yyyy-MM-dd HH:mm:59") & "','YYYY-MM-DD HH24:MI:SS')"
    End If
    
    '���ݺ�
    If txtNO.Text <> "" Then
        strFilter = strFilter & " And A.NO='" & txtNO.Text & "'"
    End If
    
    '���˿���
    If cboDept.ListIndex <> 0 Then
        strFilter = strFilter & " And B.���˿���ID+0=" & cboDept.ItemData(cboDept.ListIndex)
    End If
    
    '������Դ
    strFilter = strFilter & " And Nvl(B.������Դ,0) IN(3," & IIf(chk��Դ(0).Value, "1,4", "-1") & "," & IIf(chk��Դ(1).Value, 2, -1) & ")"
        
    '�걾��λ
    If Trim(Me.cboPart) <> "" Then
        strFilter = strFilter & " And B.�걾��λ = '" & Me.cboPart.Text & "' "
    End If
    
    '���˱�ʶ
    strPati = ""
    If txt��ʶ��.Text <> "" Then
        strPati = strPati & " And Decode(B.������Դ,1,D.�����,2,D.סԺ��,NULL)=" & txt��ʶ��.Text
    End If
    If txt���￨.Text <> "" Then
        strPati = strPati & " And D.���￨��||''='" & txt���￨.Text & "'"
    End If
    If txt����.Text <> "" Then
        strPati = strPati & " And D.����||''='" & txt����.Text & "'"
    End If
    If txtChkNO.Text <> "" Then
        strPati = strPati & " And H.����=" & txtChkNO.Text
    End If
    
    
End Sub

Private Sub Form_Activate()
    Dim curDate As Date
    
    '�����һ����ȡ�ĵ�ǰʱ��,����������ʱˢ�½��ʱ��Ϊ��ǰʱ��
    If Not mblnLoad Then
        If Format(dtpEnd.Value, "yyyy-MM-dd HH:mm") = Format(dtpEnd.Tag, "yyyy-MM-dd HH:mm") Then
            curDate = zlDatabase.Currentdate
            dtpEnd.MaxDate = curDate: dtpBegin.MaxDate = curDate
            dtpEnd.Value = Format(curDate, "yyyy-MM-dd HH:mm")
            dtpEnd.Tag = Format(dtpEnd.Value, "yyyy-MM-dd HH:mm")
        End If
    End If
    If mblnLoad Then mblnLoad = False
        
    '�Զ���λ
    dtpBegin.SetFocus
    If txtNO.Text <> "" Then
        txtNO.Text = "": txtNO.SetFocus
    End If
    If txt����.Text <> "" Then
        txt����.Text = "": txt����.SetFocus
    End If
    If txt���￨.Text <> "" Then
        txt���￨.Text = "": txt���￨.SetFocus
    End If
    If txt��ʶ��.Text <> "" Then
        txt��ʶ��.Text = "": txt��ʶ��.SetFocus
    End If
    If txtChkNO.Text <> "" Then
        txtChkNO.Text = "": txtChkNO.SetFocus
    End If
    '������Ŀ
    Call InitRptItem
    On Error Resume Next
    cboItem.Text = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPACSFilter", "���˱�����Ŀ", "")
    cboContent.Text = ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim curDate As Date
    Dim aContent() As String, i As Long
    Dim strTmp As String
    
    mblnLoad = True
    
    txtNO.Text = ""
    txt��ʶ��.Text = ""
    txt����.Text = ""
    txt���￨.Text = ""
    txt���￨.PasswordChar = IIf(gblnCardHide, "*", "")
    txtChkNO.Text = ""
    
    '��Դ��״̬
    chk��Դ(0).Value = 1
    chk��Դ(1).Value = 1
    
    '����ʱ��
    curDate = zlDatabase.Currentdate
    If mBeforeDays <= 0 Then mBeforeDays = 2 'Ĭ�ϲ�ѯ3��ǰ������
    dtpEnd.MaxDate = curDate: dtpBegin.MaxDate = curDate
    dtpBegin.Value = Format(curDate - mBeforeDays, "yyyy-MM-dd 00:00")
    dtpEnd.Value = Format(curDate, "yyyy-MM-dd HH:mm")
    dtpEnd.Tag = Format(dtpEnd.Value, "yyyy-MM-dd HH:mm")
        
    '���˿���
    Call LoadDept
    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmPACSFilter", "���˲��˿���", "")
    On Error Resume Next
    If strTmp <> "" Then Me.cboDept.Text = strTmp
    On Error GoTo 0
    '��ʼ��������ѡ��
    aContent = Split(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPACSFilter", "���˱�������", ""), "|")
    With cboContent
        .Clear: .AddItem ""
        For i = 0 To UBound(aContent)
            .AddItem aContent(i)
        Next
    End With
    
    mstrFilter = ""
    mstrPati = ""
    mblnOK = False
End Sub

Private Function LoadDept() As Boolean
'���ܣ����ݲ�����Դ��ȡ���˿���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngPre As Long
    
    If cboDept.ListIndex <> -1 Then
        lngPre = cboDept.ItemData(cboDept.ListIndex)
    End If
    strSQL = "Select Distinct A.ID,A.����,A.����,B.�������" & _
        " From ���ű� A,��������˵�� B" & _
        " Where A.ID=B.����ID And B.�������� IN('�ٴ�','����')" & _
        " And B.������� IN(3," & IIf(chk��Դ(0).Value, 1, -1) & "," & IIf(chk��Դ(1).Value, 2, -1) & ")" & _
        " And (A.����ʱ�� is NULL Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Order by A.����"
    On Error GoTo errH
    Call OpenRecord(rsTmp, strSQL, Me.Caption)
    On Error GoTo 0
    cboDept.Clear
    cboDept.AddItem "���п���"
    cboDept.ListIndex = 0
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!���� & "-" & rsTmp!����
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        If rsTmp!ID = lngPre Then cboDept.ListIndex = cboDept.NewIndex
        rsTmp.MoveNext
    Next
    LoadDept = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function InitRptItem() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    
    InitRptItem = True
    On Error GoTo errH
    strSQL = "Select Distinct D.�����ı� " & _
        "From ����ִ�п��� B, ���Ƶ���Ӧ�� C, �����ļ���� D " & _
        "Where B.������Ŀid = C.������Ŀid AND C.�����ļ�ID=D.�����ļ�ID And B.ִ�п���id = [1] AND D.��дʱ��=2"
    With frmPACStation
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, .cboDept.ItemData(.cboDept.ListIndex))
    End With
    With cboItem
        .Clear: .AddItem " "
        Do While Not rsTmp.EOF
            .AddItem Nvl(rsTmp("�����ı�"))
        
            rsTmp.MoveNext
        Loop
    End With
    
    strSQL = "Select Distinct  �걾��λ  From ������ĿĿ¼ Where ��� = 'D' And �걾��λ Is Not Null"
    zlDatabase.OpenRecordset rsTmp, strSQL, gstrSysName
    With Me.cboPart
        .Clear: .AddItem ""
        Do While Not rsTmp.EOF
            .AddItem Nvl(rsTmp("�걾��λ"))
            rsTmp.MoveNext
        Loop
    End With
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    InitRptItem = False
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer, strContent As String
    
    '���汨������
    strContent = ""
    With cboContent
        For i = 0 To .ListCount - 1
            If Len(Trim(.List(i))) > 0 Then strContent = strContent & "|" & .List(i)
        Next
    End With
    If Len(strContent) > 0 Then strContent = Mid(strContent, 2)
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPACSFilter", "���˱�������", strContent)
End Sub

Private Sub optAdviceTime_Click()
    Me.dtpBegin.SetFocus
End Sub

Private Sub optCheckTime_Click()
    Me.dtpBegin.SetFocus
End Sub

Private Sub txtChkNO_GotFocus()
    Call zlControl.TxtSelAll(txtChkNO)
End Sub

Private Sub txtChkNO_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii <> 13 Then
        If Not (txtNO.Text = "" Or txtNO.SelLength = Len(txtNO.Text)) _
            And InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtNO_Validate(Cancel As Boolean)
    If IsNumeric(txtNO.Text) Then
        txtNO.Text = GetFullNO(txtNO.Text, 0)
    End If
End Sub

Private Sub txt���￨_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txt���￨.Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txt���￨.Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt���￨_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Call SetWindowLong(txt���￨.Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txtNO_GotFocus()
    Call zlControl.TxtSelAll(txtNO)
End Sub

Private Sub txt���￨_GotFocus()
    Call zlControl.TxtSelAll(txt���￨)
End Sub

Private Sub txt���￨_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean
    
    'ȥ���ſ��������������ַ�
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
    blnCard = InputIsCard(Me.txt���￨, KeyAscii)
    
    'ˢ����ɻ�ȷ������
    If blnCard And Len(Me.txt���￨.Text) = gbytCardLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Me.txt���￨.Text <> "" Then
        If KeyAscii <> 13 Then
            Me.txt���￨.Text = Me.txt���￨.Text & Chr(KeyAscii)
            Me.txt���￨.SelStart = Len(Me.txt���￨.Text)
        End If
        KeyAscii = 0
        Me.txt���￨.Text = UCase(Me.txt���￨)
        Me.txt���￨.SetFocus
    End If
End Sub

Private Sub txt��ʶ��_GotFocus()
    Call zlControl.TxtSelAll(txt��ʶ��)
End Sub

Private Sub txt��ʶ��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
