VERSION 5.00
Begin VB.Form FrmStuffFilter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "����"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5745
   Icon            =   "frmStuffFilter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame fraLine 
      Height          =   60
      Left            =   75
      TabIndex        =   7
      Top             =   690
      Width           =   5760
   End
   Begin VB.CheckBox chk��� 
      Alignment       =   1  'Right Justify
      Caption         =   "�������Ĺ��"
      Height          =   210
      Left            =   480
      TabIndex        =   6
      Top             =   2025
      Width           =   1545
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   0
      TabIndex        =   5
      Top             =   2280
      Width           =   5760
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4500
      TabIndex        =   4
      Top             =   2430
      Width           =   1100
   End
   Begin VB.CommandButton cmdFilter 
      Cancel          =   -1  'True
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   3240
      TabIndex        =   3
      Top             =   2430
      Width           =   1100
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1845
      TabIndex        =   1
      Top             =   1245
      Width           =   3525
   End
   Begin VB.TextBox txtSim 
      Height          =   300
      Left            =   1845
      TabIndex        =   2
      Top             =   1635
      Width           =   3525
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   1845
      TabIndex        =   0
      Top             =   855
      Width           =   3525
   End
   Begin VB.Label lblComment 
      Caption         =   "    ����ϣ�����˵����ı��롢���ơ�������������롣����ڶ������򷵻ض������˽����"
      Height          =   435
      Left            =   1020
      TabIndex        =   11
      Top             =   165
      Width           =   4500
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   285
      Picture         =   "frmStuffFilter.frx":000C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "������������"
      Height          =   180
      Left            =   495
      TabIndex        =   10
      Top             =   1305
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "�������ļ���"
      Height          =   180
      Left            =   495
      TabIndex        =   9
      Top             =   1695
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�������ı���"
      Height          =   180
      Left            =   495
      TabIndex        =   8
      Top             =   915
      Width           =   1080
   End
End
Attribute VB_Name = "FrmStuffFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbln��ʾͣ�� As Boolean

Public Sub ShowMe(ByVal FrmParent As Object, ByVal blnͣ�� As Boolean)
    mbln��ʾͣ�� = blnͣ��
    Me.Show , FrmParent
End Sub



Private Sub chk���_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then OS.PressKey (vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdFilter_Click()
    Dim rs As New ADODB.Recordset
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
        MsgBox "������������Ϣ", vbExclamation, gstrSysName
        On Error Resume Next
        Me.Show
        Me.txtCode.SetFocus: Exit Sub
    End If

    
    If Me.chk���.Value = 0 Then
        gstrSQL = "SELECT DISTINCT I.ID AS ����ID" & _
                " FROM ������ĿĿ¼ I,������Ŀ���� N" & _
                " WHERE I.ID=N.������ĿID AND I.��� = '4' " & strCondition

    Else
        gstrSQL = "SELECT DISTINCT I.ID AS ����ID " & _
                 " FROM ������ĿĿ¼ I,�������� T,�շ���ĿĿ¼ D,�շ���Ŀ���� N " & _
                 " WHERE I.ID=T.����ID And T.����ID=D.ID AND T.����ID=N.�շ�ϸĿID AND I.��� = '4' " & strCondition
                 
    End If
    If Not mbln��ʾͣ�� Then
        gstrSQL = gstrSQL & " And (I.����ʱ�� Is NULL Or to_Char(I.����ʱ��,'yyyy-MM-dd')='3000-01-01')"
        If Me.chk���.Value = 1 Then gstrSQL = gstrSQL & " And (D.����ʱ�� Is NULL Or to_Char(D.����ʱ��,'yyyy-MM-dd')='3000-01-01')"
    End If
    
    gstrSQL = gstrSQL & "order by ����id"
    
    err = 0: On Error GoTo ErrHand
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Trim(Me.txtCode.Text) & "%", IIf(gstrMatchMethod = "0", "%", "") & Trim(Me.txtName.Text) & "%", IIf(gstrMatchMethod = "0", "%", "") & Trim(Me.txtSim.Text) & "%")
    
    With rs
        If .EOF Then
            MsgBox "û���ҵ�������Ϣ��", vbExclamation, gstrSysName
            On Error Resume Next
            Me.Show
            Me.txtCode.SetFocus
            Exit Sub
        Else
            For n = 1 To .RecordCount
                If n = 1 Then
                    strResult = Val(!����id)
                Else
                    strResult = strResult & "," & Val(!����id)
                End If
                .MoveNext
            Next
        End If
    End With
    
    Me.Hide
    Call frmStuffMgr.zlGetFilter(strResult)
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
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
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtcode_GotFocus()
    Me.txtCode.SelStart = 0
    Me.txtCode.SelLength = 100
End Sub


Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtName_GotFocus()
    Me.txtCode.SelStart = 0
    Me.txtCode.SelLength = 100
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtSim_GotFocus()
    Me.txtCode.SelStart = 0
    Me.txtCode.SelLength = 100
End Sub

Private Sub txtSim_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub




