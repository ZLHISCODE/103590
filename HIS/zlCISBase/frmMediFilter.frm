VERSION 5.00
Begin VB.Form frmMediFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   1905
      TabIndex        =   0
      Top             =   1278
      Width           =   3525
   End
   Begin VB.TextBox txtSim 
      Height          =   300
      Left            =   1905
      TabIndex        =   2
      Top             =   2055
      Width           =   3525
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1905
      TabIndex        =   1
      Top             =   1671
      Width           =   3525
   End
   Begin VB.CheckBox chk�в�ҩ 
      Caption         =   "�в�ҩ"
      Height          =   210
      Left            =   4365
      TabIndex        =   12
      Top             =   945
      Width           =   1035
   End
   Begin VB.CheckBox chk�г�ҩ 
      Caption         =   "�г�ҩ"
      Height          =   210
      Left            =   3142
      TabIndex        =   11
      Top             =   945
      Value           =   1  'Checked
      Width           =   1035
   End
   Begin VB.CheckBox chk����ҩ 
      Caption         =   "����ҩ"
      Height          =   210
      Left            =   1920
      TabIndex        =   10
      Top             =   945
      Value           =   1  'Checked
      Width           =   1035
   End
   Begin VB.CommandButton cmdFilter 
      Cancel          =   -1  'True
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   3300
      TabIndex        =   4
      Top             =   2850
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4560
      TabIndex        =   5
      Top             =   2850
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   60
      TabIndex        =   9
      Top             =   2700
      Width           =   5760
   End
   Begin VB.CheckBox chk��� 
      Alignment       =   1  'Right Justify
      Caption         =   "���˹��ҩƷ"
      Height          =   210
      Left            =   540
      TabIndex        =   3
      Top             =   2445
      Width           =   1545
   End
   Begin VB.Frame fraLine 
      Height          =   60
      Left            =   135
      TabIndex        =   7
      Top             =   630
      Width           =   5760
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "����ҩƷ����"
      Height          =   180
      Left            =   555
      TabIndex        =   15
      Top             =   1335
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "����ҩƷ����"
      Height          =   180
      Left            =   555
      TabIndex        =   14
      Top             =   2115
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "����ҩƷ����"
      Height          =   180
      Left            =   555
      TabIndex        =   13
      Top             =   1725
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ѡ��ҩƷ����"
      Height          =   180
      Left            =   585
      TabIndex        =   8
      Top             =   945
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   345
      Picture         =   "frmMediFilter.frx":0000
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lblComment 
      Caption         =   "    ����ϣ�����˵�ҩƷ�Ĳ��ʼ�ҩƷ���롢���ơ�������������롣����ڶ������򷵻ض������˽����"
      Height          =   435
      Left            =   1080
      TabIndex        =   6
      Top             =   105
      Width           =   4500
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMediFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbln��ʾͣ��ҩƷ As Boolean
Private mblnSelfMedi As Boolean  '�Թ�ҩ true-�Թ�ҩ false-�����Թ�ҩ

Public Sub ShowMe(ByVal frmParent As Object, ByVal blnͣ�� As Boolean, ByVal blnSelfMedi As Boolean)
    mbln��ʾͣ��ҩƷ = blnͣ��
    mblnSelfMedi = blnSelfMedi
    Me.Show , frmParent
End Sub

Private Sub chk���_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub chk����ҩ_Click()
    If chk����ҩ.Value = 1 Then
        chk�в�ҩ.Value = 0
    ElseIf chk�г�ҩ.Value = 0 Then
        chk�в�ҩ.Value = 1
    End If
End Sub

Private Sub chk�в�ҩ_Click()
    If chk�в�ҩ.Value = 1 Then
        chk����ҩ.Value = 0
        chk�г�ҩ.Value = 0
    Else
        chk����ҩ.Value = 1
        chk�г�ҩ.Value = 1
    End If
End Sub

Private Sub chk�г�ҩ_Click()
    If chk�г�ҩ.Value = 1 Then
        chk�в�ҩ.Value = 0
    ElseIf chk����ҩ.Value = 0 Then
        chk�в�ҩ.Value = 1
    End If
End Sub


Private Sub CmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdFilter_Click()
    Dim rs As New ADODB.Recordset
    Dim strKind As String
    Dim strKind1 As String
    Dim strResult As String
    Dim n As Long
    Dim strCondition As String
    
    If Len(Trim(Me.txtCode.Text)) > 0 Then
        If Me.chk���.Value = 0 Then
            strCondition = " AND I.���� LIKE [1] "
        Else
            strCondition = " AND D.���� LIKE [1] "
        End If
    End If
    If Len(Trim(Me.txtName.Text)) > 0 Then
        strCondition = " AND N.���� LIKE [2] "
    End If
    If Len(Trim(Me.txtSim.Text)) > 0 Then
        strCondition = " AND N.���� LIKE [3] "
    End If
    
    If Len(strCondition) = 0 Then
        MsgBox "������ҩƷ��Ϣ", vbExclamation, gstrSysName
        On Error Resume Next
        Me.Show
        Me.txtCode.SetFocus: Exit Sub
    End If
    
    If chk�в�ҩ.Value = 1 Then
        strKind = "7"
        strKind1 = ",7,"
    ElseIf chk����ҩ.Value = 1 And chk�г�ҩ.Value = 1 Then
        strKind = "5,6"
        strKind1 = ",5,6,"
    ElseIf chk����ҩ.Value = 1 Then
        strKind = "5"
        strKind1 = ",5,"
    ElseIf chk�г�ҩ.Value = 1 Then
        strKind = "6"
        strKind1 = ",6,"
    End If
    
    If Me.chk���.Value = 0 Then
        If mblnSelfMedi = True Then
            gstrSql = "Select Distinct i.����id, i.Id As ҩ��id, 0 As ҩƷid" & vbNewLine & _
                    "From ������ĿĿ¼ I, ������Ŀ���� N, ҩƷ���� A" & vbNewLine & _
                    "Where i.Id = n.������Ŀid And i.Id = a.ҩ��id And Instr([4], ',' || i.��� || ',') > 0 And a.�ٴ��Թ�ҩ = 1 " & strCondition

        Else
            gstrSql = "SELECT DISTINCT I.����ID,I.ID AS ҩ��ID,0 AS ҩƷID" & _
                    " FROM ������ĿĿ¼ I,������Ŀ���� N" & _
                    " WHERE I.ID=N.������ĿID AND Instr([4], ','||I.���||',') > 0 " & strCondition
        End If

    Else
        If mblnSelfMedi = True Then
            gstrSql = "SELECT DISTINCT I.����ID,I.ID AS ҩ��ID,D.ID AS ҩƷID  " & _
                     " FROM ������ĿĿ¼ I,ҩƷ��� T,�շ���ĿĿ¼ D,�շ���Ŀ���� N,ҩƷ���� A " & _
                     " WHERE I.ID=T.ҩ��ID And T.ҩƷID=D.ID AND T.ҩƷID=N.�շ�ϸĿID AND i.ID=A.ҩ��id and a.�ٴ��Թ�ҩ=1 AND Instr([4], ','||I.���||',') > 0 " & strCondition
        Else
            gstrSql = "SELECT DISTINCT I.����ID,I.ID AS ҩ��ID,D.ID AS ҩƷID  " & _
                     " FROM ������ĿĿ¼ I,ҩƷ��� T,�շ���ĿĿ¼ D,�շ���Ŀ���� N " & _
                     " WHERE I.ID=T.ҩ��ID And T.ҩƷID=D.ID AND T.ҩƷID=N.�շ�ϸĿID AND Instr([4], ','||I.���||',') > 0 " & strCondition
        End If
                 
    End If
    If Not mbln��ʾͣ��ҩƷ Then
        gstrSql = gstrSql & " And (I.����ʱ�� Is NULL Or to_Char(I.����ʱ��,'yyyy-MM-dd')='3000-01-01')"
        If Me.chk���.Value = 1 Then gstrSql = gstrSql & " And (D.����ʱ�� Is NULL Or to_Char(D.����ʱ��,'yyyy-MM-dd')='3000-01-01')"
    End If

    Err = 0: On Error GoTo errHand
    
    Set rs = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Trim(Me.txtCode.Text) & "%", gstrMatch & Trim(Me.txtName.Text) & "%", gstrMatch & Trim(Me.txtSim.Text) & "%", strKind1)
    
    With rs
        If .EOF Then
            MsgBox "û���ҵ�ҩƷ��Ϣ��", vbExclamation, gstrSysName
            On Error Resume Next
            Me.Show
            Me.txtCode.SetFocus
            Exit Sub
        Else
            For n = 1 To .RecordCount
                If n = 1 Then
                    strResult = Val(!ҩ��ID)
                Else
                    strResult = strResult & "," & Val(!ҩ��ID)
                End If
                .MoveNext
            Next
        End If
    End With
    
    Me.Hide
    Call frmMediLists.zlGetFilter(strKind, strResult)
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    If mblnSelfMedi = True Then
        chk�в�ҩ.Visible = False
    Else
        chk�в�ҩ.Visible = True
    End If
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub txtAlia_Change()
    
End Sub

Private Sub txtAlia_GotFocus()
    Me.txtCode.SelStart = 0
    Me.txtCode.SelLength = 100
End Sub

Private Sub txtAlia_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtcode_GotFocus()
    Me.txtCode.SelStart = 0
    Me.txtCode.SelLength = 100
End Sub


Private Sub txtcode_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub





Private Sub txtName_GotFocus()
    Me.txtCode.SelStart = 0
    Me.txtCode.SelLength = 100
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtSim_GotFocus()
    Me.txtCode.SelStart = 0
    Me.txtCode.SelLength = 100
End Sub

Private Sub txtSim_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


