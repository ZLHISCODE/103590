VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPatiFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ﲡ�˹���"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   Icon            =   "frmPatiFilter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame3 
      Height          =   120
      Left            =   -75
      TabIndex        =   20
      Top             =   570
      Width           =   5085
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   -45
      TabIndex        =   19
      Top             =   1530
      Width           =   5085
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -75
      TabIndex        =   18
      Top             =   2550
      Width           =   5130
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2940
      TabIndex        =   9
      Top             =   2805
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1845
      TabIndex        =   8
      Top             =   2805
      Width           =   1100
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   2955
      TabIndex        =   7
      Top             =   2190
      Width           =   1260
   End
   Begin VB.TextBox txt���￨ 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   960
      TabIndex        =   6
      Top             =   2190
      Width           =   1260
   End
   Begin VB.TextBox txt����� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2955
      TabIndex        =   5
      Top             =   1785
      Width           =   1260
   End
   Begin VB.TextBox txt�Һŵ� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   960
      TabIndex        =   4
      Top             =   1785
      Width           =   1260
   End
   Begin VB.ComboBox cboҽ�� 
      Height          =   300
      Left            =   960
      TabIndex        =   3
      Top             =   1215
      Width           =   3255
   End
   Begin VB.ComboBox cbo���� 
      Height          =   300
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   3255
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   300
      Left            =   960
      TabIndex        =   0
      Top             =   210
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   102039555
      CurrentDate     =   38004
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   300
      Left            =   2955
      TabIndex        =   1
      Top             =   210
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   102039555
      CurrentDate     =   38004
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Left            =   2535
      TabIndex        =   17
      Top             =   2250
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���￨"
      Height          =   180
      Left            =   360
      TabIndex        =   16
      Top             =   2250
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����"
      Height          =   180
      Left            =   2355
      TabIndex        =   15
      Top             =   1845
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Һŵ�"
      Height          =   180
      Left            =   360
      TabIndex        =   14
      Top             =   1845
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ҽ��"
      Height          =   180
      Left            =   180
      TabIndex        =   13
      Top             =   1275
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������"
      Height          =   180
      Left            =   180
      TabIndex        =   12
      Top             =   900
      Width           =   720
   End
   Begin VB.Label lbl��ʼʱ�� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "����ʱ��"
      Height          =   180
      Left            =   180
      TabIndex        =   11
      Top             =   255
      Width           =   720
   End
   Begin VB.Label lbl����ʱ�� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      Height          =   180
      Left            =   2490
      TabIndex        =   10
      Top             =   270
      Width           =   180
   End
End
Attribute VB_Name = "frmPatiFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mstrLike As String
Private mblnOK As Boolean
Private mlngPreDept As Long
Private mdBegin As Date, mdEnd As Date
Private mlng����ID As Long, mstrҽ�� As String
Private mstr�Һŵ� As String, mstr����� As String, mstr���￨ As String, mstr���� As String

Public Function ShowMe(frmParent As Object, dBegin As Date, dEnd As Date, _
     lng����ID As Long, strҽ�� As String, str�Һŵ� As String, _
     str����� As String, str���￨ As String, str���� As String, strPrivs As String) As Boolean
    
    mdBegin = dBegin
    mdEnd = dEnd
    mlng����ID = lng����ID
    mstrҽ�� = strҽ��
    mstr�Һŵ� = str�Һŵ�
    mstr����� = str�����
    mstr���￨ = str���￨
    mstr���� = str����
    mstrPrivs = strPrivs
    
    Me.Show 1, frmParent
    
    If mblnOK Then
        dBegin = mdBegin
        dEnd = mdEnd
        lng����ID = mlng����ID
        strҽ�� = mstrҽ��
        str�Һŵ� = mstr�Һŵ�
        str����� = mstr�����
        str���￨ = mstr���￨
        str���� = mstr����
    End If
    ShowMe = mblnOK
End Function

Private Sub cbo����_Click()
    If cbo����.ListIndex <> -1 Then
        If mlngPreDept <> cbo����.ItemData(cbo����.ListIndex) Then
            mlngPreDept = cbo����.ItemData(cbo����.ListIndex)
            Call ReadDoctor(mlngPreDept)
        End If
    ElseIf mlngPreDept <> 0 Then
        mlngPreDept = 0
        Call ReadDoctor
    End If
End Sub

Private Sub cbo����_GotFocus()
    Call zlControl.TxtSelAll(cbo����)
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call ZLCommFun.PressKey(vbKeyTab) '��SetFocus��ʽ�ἤ��Validate�¼�,������һ����vsFlexGrid�ؼ���
    Else
        If InStr(mstrPrivs, "���в���Ա") = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub cbo����_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnDo As Boolean
    
    On Error GoTo errH
        
    If cbo����.Text <> "" Then
        blnDo = True
        If cbo����.ListIndex <> -1 Then
            If cbo����.List(cbo����.ListIndex) = cbo����.Text Then blnDo = False
        End If
        If blnDo Then
            strSql = "Select B.ID,B.����,B.����" & _
                " From ���ű� B,��������˵�� C" & _
                " Where B.ID=C.����ID And C.������� In(1,3) And C.��������='�ٴ�'" & _
                " And (B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or B.����ʱ�� is Null)" & _
                " And (B.���� Like [1] Or Upper(B.����) Like [2] Or Upper(B.����) Like [2])" & _
                " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & _
                " Order by B.����"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UCase(cbo����.Text) & "%", mstrLike & UCase(cbo����.Text) & "%")
            If Not rsTmp.EOF Then
                Call Cbo.SeekIndex(cbo����, rsTmp!ID)
            ElseIf mlngPreDept <> 0 Then
                Call Cbo.SeekIndex(cbo����, mlngPreDept)
            Else
                cbo����.Text = ""
                Call cbo����_Click
            End If
        End If
    Else
        Call cbo����_Click
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboҽ��_GotFocus()
    Call zlControl.TxtSelAll(cboҽ��)
End Sub

Private Sub cboҽ��_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnDo As Boolean
    Dim lng����ID As Long
    
    On Error GoTo errH
        
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        If cboҽ��.Text <> "" Then
            blnDo = True
            If cboҽ��.ListIndex <> -1 Then
                If cboҽ��.List(cboҽ��.ListIndex) = cboҽ��.Text Then blnDo = False
            End If
            
            If blnDo Then
                If cbo����.ListIndex <> -1 Then
                    lng����ID = cbo����.ItemData(cbo����.ListIndex)
                End If
                If lng����ID <> 0 Then
                    strSql = "Select Distinct A.����,A.����" & _
                        " From ��Ա�� A,��Ա����˵�� B,������Ա C" & _
                        " Where A.ID=B.��ԱID And A.ID=C.��ԱID" & _
                        " And B.��Ա����='ҽ��' And C.����ID=[1]" & _
                        " And (A.��� Like [2] Or Upper(A.����) Like [3] Or Upper(A.����) Like [3])" & _
                        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
                        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                        " Order by A.����"
                Else
                    strSql = "Select Distinct A.����,A.����" & _
                        " From ��Ա�� A,��Ա����˵�� B" & _
                        " Where A.ID=B.��ԱID And B.��Ա����='ҽ��'" & _
                        " And (A.��� Like [2] Or Upper(A.����) Like [3] Or Upper(A.����) Like [3])" & _
                        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
                        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                        " Order by A.����"
                End If
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID, UCase(cboҽ��.Text) & "%", mstrLike & UCase(cboҽ��.Text) & "%")
                If Not rsTmp.EOF Then
                    Call Cbo.SeekIndex(cboҽ��, rsTmp!����)
                Else
                    cboҽ��.Text = ""
                End If
            End If
        End If
        Call ZLCommFun.PressKey(vbKeyTab)
    Else
        If InStr(mstrPrivs, "���в���Ա") = 0 Then KeyAscii = 0
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim curDate As Date
    
    If cbo����.ListIndex <> -1 Then
        mlng����ID = cbo����.ItemData(cbo����.ListIndex)
    Else
        mlng����ID = 0
    End If
    If cboҽ��.ListIndex <> -1 Then
        mstrҽ�� = ZLCommFun.GetNeedName(cboҽ��.Text)
    Else
        mstrҽ�� = ""
    End If
    If txt�Һŵ�.Text <> "" Then
        mstr�Һŵ� = txt�Һŵ�.Text
    Else
        mstr�Һŵ� = ""
    End If
    If txt�����.Text <> "" Then
        mstr����� = txt�����.Text
    Else
        mstr����� = ""
    End If
    If txt���￨.Text <> "" Then
        mstr���￨ = txt���￨.Text
    Else
        mstr���￨ = ""
    End If
    If txt����.Text <> "" Then
        mstr���� = txt����.Text
    Else
        mstr���� = ""
    End If
    
    mdBegin = dtpBegin.Value
    mdEnd = dtpEnd.Value
    
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 And Not (Me.ActiveControl Is cbo���� Or Me.ActiveControl Is cboҽ��) Then
        KeyAscii = 0
        Call ZLCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    mblnOK = False
    mlngPreDept = -1
    mstrLike = IIf(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "") '����ƥ�䷽ʽ
    
    Call Cbo.SeekIndex(cbo����, mlng����ID)
    Call Cbo.SeekIndex(cboҽ��, mstrҽ��)
    dtpBegin.Value = mdBegin
    dtpEnd.Value = mdEnd
    txt�Һŵ�.Text = mstr�Һŵ�
    txt�����.Text = mstr�����
    txt���￨.Text = mstr���￨
    txt���￨.PasswordChar = IIf(gblnCardHide, "*", "")
    txt����.Text = mstr����
    
    On Error GoTo errH
    
    '��ȡ�������:ȱʡΪ�޿���
    If InStr(mstrPrivs, "���в���Ա") > 0 Then
        strSql = "Select Distinct B.ID,B.����,B.����" & _
            " From ���ű� B,��������˵�� C" & _
            " Where B.ID=C.����ID And C.������� In(1,3) And C.��������='�ٴ�'" & _
            " And (B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or B.����ʱ�� is Null)" & _
            " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & _
            " Order by B.����"
    Else
        strSql = "Select Distinct B.ID,B.����,B.����" & _
            " From ���ű� B,��������˵�� C,������Ա D" & _
            " Where B.ID=C.����ID And B.ID=D.����ID And D.��ԱID=[1]" & _
            " And C.������� In(1,3) And C.��������='�ٴ�'" & _
            " And (B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or B.����ʱ�� is Null)" & _
            " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & _
            " Order by B.����"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    Do While Not rsTmp.EOF
        cbo����.AddItem rsTmp!���� & "-" & rsTmp!����
        cbo����.ItemData(cbo����.NewIndex) = rsTmp!ID
        If rsTmp!ID = mlng����ID Then
            Call Cbo.SetIndex(cbo����.hwnd, cbo����.NewIndex)
        End If
        rsTmp.MoveNext
    Loop
        
    '��ȡ����ҽ��
    Call cbo����_Click
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ReadDoctor(Optional ByVal lng����ID As Long)
'���ܣ���ȡָ��������ҵ�ҽ��
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    cboҽ��.Clear
    
    On Error GoTo errH
    
    If InStr(mstrPrivs, "���в���Ա") = 0 Then
        strSql = "Select ����,���� From ��Ա�� Where ID=[2]"
        cboҽ��.Enabled = False
    ElseIf lng����ID <> 0 Then
        strSql = "Select Distinct A.����,A.����" & _
            " From ��Ա�� A,��Ա����˵�� B,������Ա C" & _
            " Where A.ID=B.��ԱID And A.ID=C.��ԱID" & _
            " And B.��Ա����='ҽ��' And C.����ID=[1]" & _
            " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " Order by A.����"
    Else
        strSql = "Select Distinct A.����,A.����" & _
            " From ��Ա�� A,��Ա����˵�� B" & _
            " Where A.ID=B.��ԱID And B.��Ա����='ҽ��'" & _
            " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " Order by A.����"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID, UserInfo.ID)
    Do While Not rsTmp.EOF
        cboҽ��.AddItem rsTmp!���� & "-" & rsTmp!����
        If rsTmp!���� = mstrҽ�� Then
            cboҽ��.ListIndex = cboҽ��.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt�Һŵ�_Change()
    If txt�Һŵ�.Text <> "" Then
        txt���￨.Text = ""
        txt����.Text = ""
        txt�����.Text = ""
    End If
End Sub

Private Sub txt�Һŵ�_GotFocus()
    Call zlControl.TxtSelAll(txt�Һŵ�)
End Sub

Private Sub txt�Һŵ�_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt�Һŵ�_Validate(Cancel As Boolean)
    If IsNumeric(txt�Һŵ�.Text) Then
        txt�Һŵ�.Text = GetFullNO(txt�Һŵ�.Text, 12)
    End If
End Sub

Private Sub txt���￨_Change()
    If txt���￨.Text <> "" Then
        txt�����.Text = ""
        txt����.Text = ""
        txt�Һŵ�.Text = ""
    End If
End Sub

Private Sub txt���￨_GotFocus()
    Call zlControl.TxtSelAll(txt���￨)
End Sub

Private Sub txt���￨_KeyPress(KeyAscii As Integer)
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt�����_Change()
    If txt�����.Text <> "" Then
        txt���￨.Text = ""
        txt����.Text = ""
        txt�Һŵ�.Text = ""
    End If
End Sub

Private Sub txt�����_GotFocus()
    Call zlControl.TxtSelAll(txt�����)
End Sub

Private Sub txt�����_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt����_Change()
    If txt����.Text <> "" Then
        txt�����.Text = ""
        txt���￨.Text = ""
        txt�Һŵ�.Text = ""
    End If
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub
