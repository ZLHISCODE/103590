VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLabItemFormula 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ʽ"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9195
   Icon            =   "FrmLabItemFormula.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   9195
   StartUpPosition =   2  '��Ļ����
   Begin MSComctlLib.TreeView tvwItem 
      Height          =   4290
      Left            =   105
      TabIndex        =   6
      Top             =   360
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   7567
      _Version        =   393217
      Indentation     =   459
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      Appearance      =   1
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "ȡ��(&E)"
      Height          =   350
      Left            =   7815
      TabIndex        =   3
      Top             =   4275
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6315
      TabIndex        =   2
      Top             =   4275
      Width           =   1100
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "��֤(&C)"
      Height          =   350
      Left            =   3555
      TabIndex        =   1
      Top             =   4275
      Width           =   1100
   End
   Begin VB.TextBox txtFormula 
      Height          =   3825
      Left            =   3285
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   5745
   End
   Begin VB.Label lblFormula 
      Caption         =   "���磺([SD]+[CV])/100"
      Height          =   210
      Left            =   3360
      TabIndex        =   5
      Top             =   105
      Width           =   4080
   End
   Begin VB.Label lblItem 
      Caption         =   "��Ŀ"
      Height          =   210
      Left            =   180
      TabIndex        =   4
      Top             =   105
      Width           =   585
   End
End
Attribute VB_Name = "FrmLabItemFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFormula As String '����Ĺ�ʽ
Private mlngID As Long '������ĿID
Private mstrItem As String '��Ŀ�����ڼ��

Private Sub cmdCheck_Click()
    If CheckFormula(txtFormula) Then
        cmdOk.Enabled = True
    Else
        MsgBox "��ʽ����", vbExclamation, Me.Caption
    End If
    
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If CheckFormula(txtFormula) Then
        mstrFormula = txtFormula
        Unload Me
    Else
         MsgBox "��ʽ����", vbExclamation, Me.Caption
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim itemTmp As ListItem
    Dim rsGroup As ADODB.Recordset
    
    On Error GoTo ErrHandle
    txtFormula = mstrFormula
    tvwItem.Nodes.Clear
    strSQL = "Select ����||'-'||���� as ��ʾ���� ,���� ,���� From ���Ƽ�������"
    Set rsGroup = zldatabase.OpenSQLRecord(strSQL, Me.Caption)
    mstrItem = ","
    Do Until rsGroup.EOF
        tvwItem.Nodes.Add , , "" & rsGroup.Fields("����"), "" & rsGroup.Fields("��ʾ����")
        strSQL = "Select distinct A.������Ŀid, A.��д, B.������ " & vbNewLine & _
                "From ������ĿĿ¼ D, ���鱨����Ŀ C, ����������Ŀ B, ������Ŀ A" & vbNewLine & _
                "Where C.������Ŀid = D.ID And C.������Ŀid = A.������Ŀid And A.������Ŀid = B.ID And D.�������� = [1] And A.������� = 1"
        Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, "" & rsGroup.Fields("����"))
        Do Until rsTmp.EOF
            '�����������������õ���Ŀ�����Ǽ�����Ŀ��
            mstrItem = mstrItem & "[" & IIf("" & rsTmp.Fields("��д") = "", rsTmp.Fields("������ĿID"), rsTmp.Fields("��д")) & "]" & ","
            tvwItem.Nodes.Add "" & rsGroup.Fields("����"), tvwChild, "K" & rsGroup.Fields("����") & "_" & rsTmp.Fields("������ĿID"), "[" & IIf("" & rsTmp.Fields("��д") = "", rsTmp.Fields("������ĿID"), rsTmp.Fields("��д")) & "]" & rsTmp.Fields("������")
            rsTmp.MoveNext
        Loop
        rsGroup.MoveNext
    Loop
    cmdOk.Enabled = False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function DefFormula(ByVal lngID As Long, ByVal strFormula As String, ByVal frmMain As Form) As String
    'lngID :��ǰ�����ļ�����Ŀ ID
    'strFormula :ԭ���Ĺ�ʽ
    'frmMain: ���ô���
    mlngID = lngID
    mstrFormula = strFormula
    
    Me.Show vbModal, frmMain
    DefFormula = mstrFormula
End Function

Private Function CheckFormula(ByVal strFormula As String) As Boolean
    '
    Dim strTmp As String, strLine As String, i As Integer
    Dim dblValues As Double, strItem As String, lngLength As Long
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo ErrHandle
    strLine = strFormula
    strTmp = ""
    Do While strLine Like "*[[]*[]]*"
        strTmp = strTmp & Mid(strLine, 1, InStr(strLine, "[") - 1) & "(" & i & "+ 1)"
        lngLength = InStr(strLine, "]") - InStr(strLine, "[")
        strItem = Mid(strLine, InStr(strLine, "["), lngLength + 1)
        If InStr(mstrItem, "," & strItem & ",") <= 0 Then
            Exit Function
        End If
        strLine = Mid(strLine, InStr(strLine, "]") + 1)
        i = i + 1
    Loop
    strTmp = strTmp & strLine

    Set rsTmp = zldatabase.OpenSQLRecord("Select " & strTmp & " as ������ From Dual", Me.Caption)
    If Not rsTmp.EOF Then
        dblValues = rsTmp.Fields("������")
        CheckFormula = True
    End If
    
    Exit Function
ErrHandle:
    CheckFormula = False
End Function

Private Sub tvwItem_DblClick()
    If InStr(tvwItem.SelectedItem.Text, "]") > 0 Then
        txtFormula.SelText = Mid(tvwItem.SelectedItem.Text, 1, InStr(tvwItem.SelectedItem.Text, "]"))
    End If
End Sub

Private Sub txtFormula_Change()
    cmdOk.Enabled = False
End Sub
