VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLabVerifySet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10545
   Icon            =   "frmLabVerifySet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin MSComctlLib.TreeView tvwItem 
      Height          =   4995
      Left            =   75
      TabIndex        =   5
      Top             =   90
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   8811
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
      Left            =   9165
      TabIndex        =   3
      Top             =   3765
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7665
      TabIndex        =   2
      Top             =   3765
      Width           =   1100
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "��֤(&C)"
      Height          =   350
      Left            =   3555
      TabIndex        =   1
      Top             =   3765
      Width           =   1100
   End
   Begin VB.TextBox txtFormula 
      Height          =   3600
      Left            =   3210
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   90
      Width           =   7200
   End
   Begin VB.Label lbl˵�� 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmLabVerifySet.frx":000C
      ForeColor       =   &H00008000&
      Height          =   555
      Left            =   3255
      TabIndex        =   6
      Top             =   4260
      Width           =   7170
   End
   Begin VB.Label lblFormula 
      Caption         =   "���磺[��ϸ��]>2 AND [��ϸ��]<10) OR ([��ϸ��ƽ�����] >4 AND [��ϸ��ѹ��] <20) "
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   3255
      TabIndex        =   4
      Top             =   4860
      Width           =   7170
   End
End
Attribute VB_Name = "frmLabVerifySet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFormula As String '����Ĺ���
Private mlngID As Long        '������ĿID
Private mstrItem As String    '��Ŀ�����ڼ��
Private mlng����ID As Long

'----------------------------------------------------
'-- �����Ǳ�����ؼ�����
'----------------------------------------------------
Private Sub Form_Load()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim itemTmp As ListItem
    Dim rsGroup As ADODB.Recordset
    Dim strItem As String
    
    txtFormula = mstrFormula
    tvwItem.Nodes.Clear
    
    On Error GoTo errHand
    If mlngID = 0 And mlng����ID = 0 Then
        strSQL = "Select ����||'-'||���� as ��ʾ���� ,���� ,���� From ���Ƽ������� where ���� IN (" & vbNewLine & _
                        "Select D.�������� From ������Ŀ A, ����������Ŀ B, ������ĿĿ¼ D, ���鱨����Ŀ C" & vbNewLine & _
                        "Where A.������Ŀid = B.ID And B.ID = C.������Ŀid And C.������Ŀid = D.ID And D.��� = 'C'  And" & vbNewLine & _
                        "      Nvl(D.�����Ŀ, 0) = 0 )"
    Else
'        strSQL = "Select ���� || '-' || ���� As ��ʾ����, ����, ����" & vbNewLine & _
'                "From ���Ƽ�������" & vbNewLine & _
'                "Where ���� In (" & vbNewLine & _
'                "" & vbNewLine & _
'                "             Select A.��������" & vbNewLine & _
'                "             From ������Ŀ��� B, ������ĿĿ¼ A" & vbNewLine & _
'                "             Where A.ID = B.������Ŀid And A.��� = 'C' And Nvl(A.�����Ŀ, 0) = 0 And Nvl(A.����Ӧ��, 0) = 1 And B.�������id = [1]" & vbNewLine & _
'                "             Union" & vbNewLine & _
'                "" & vbNewLine & _
'                "             Select A.��������" & vbNewLine & _
'                "             From ������ĿĿ¼ A" & vbNewLine & _
'                "             Where Nvl(����Ӧ��, 0) = 1 And Nvl(A.�����Ŀ, 0) = 0 And A.��� = 'C' And A.ID = [1]" & vbNewLine & _
'                "             Union" & vbNewLine & _
'                "" & vbNewLine & _
'                "             Select D.��������" & vbNewLine & _
'                "             From ������ĿĿ¼ D, ���鱨����Ŀ C, ����������Ŀ B, ����������Ŀ A" & vbNewLine & _
'                "             Where C.������Ŀid = D.ID And B.ID = C.������Ŀid And A.��Ŀid = B.ID And D.��� = 'C' And Nvl(D.�����Ŀ, 0) = 0 And" & vbNewLine & _
'                "                   Nvl(D.����Ӧ��, 0) = 1 And A.����id = [2])
        strSQL = "Select ���� || '-' || ���� As ��ʾ����, ����, ����" & vbNewLine & _
                "From ���Ƽ�������" & vbNewLine & _
                "Where ���� In (Select ��������" & vbNewLine & _
                "             From ������ĿĿ¼" & vbNewLine & _
                "             Where ID = [1]" & vbNewLine & _
                "             Union" & vbNewLine & _
                "             Select D.��������" & vbNewLine & _
                "             From ������ĿĿ¼ D, ���鱨����Ŀ C, ����������Ŀ B, ����������Ŀ A" & vbNewLine & _
                "             Where C.������Ŀid = D.ID And B.ID = C.������Ŀid And A.��Ŀid = B.ID And D.��� = 'C' And Nvl(D.�����Ŀ, 0) = 0 And" & vbNewLine & _
                "                    A.����id = [2])"


    End If
    Set rsGroup = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID, mlng����ID)
    mstrItem = ","
    Do Until rsGroup.EOF
        tvwItem.Nodes.Add , , "" & rsGroup.Fields("����"), "" & rsGroup.Fields("��ʾ����")
        If mlngID = 0 And mlng����ID = 0 Then
            
            strSQL = "Select Distinct A.������Ŀid, A.��д, B.������,d.���� " & vbNewLine & _
                    "From ������Ŀ A, ����������Ŀ B, ������ĿĿ¼ D, ���鱨����Ŀ C" & vbNewLine & _
                    "Where A.������Ŀid = B.ID And B.ID = C.������Ŀid And C.������Ŀid = D.ID And D.��� = 'C'  And" & vbNewLine & _
                    "      Nvl(D.�����Ŀ, 0) = 0 And D.�������� = [1]"

        Else
'            strSQL = "Select E.������Ŀid, E.��д, D.������" & vbNewLine & _
'                    "From ������Ŀ E, ����������Ŀ D, ���鱨����Ŀ C, ������Ŀ��� B, ������ĿĿ¼ A" & vbNewLine & _
'                    "Where A.ID = C.������Ŀid And C.������Ŀid = D.ID And D.ID = E.������Ŀid And A.ID = B.������Ŀid And A.��� = 'C' And Nvl(A.�����Ŀ, 0) = 0 And" & vbNewLine & _
'                    "      Nvl(A.����Ӧ��, 0) = 1 And B.�������id = [2] And A.�������� = [1]" & vbNewLine & _
'                    "Union" & vbNewLine & _
'                    "" & vbNewLine & _
'                    "Select E.������Ŀid, E.��д, B.������" & vbNewLine & _
'                    "From ������Ŀ E, ���鱨����Ŀ C, ����������Ŀ B, ������ĿĿ¼ A" & vbNewLine & _
'                    "Where E.������Ŀid = B.ID And A.ID = C.������Ŀid And C.������Ŀid = E.������Ŀid And Nvl(A.����Ӧ��, 0) = 1 And Nvl(A.�����Ŀ, 0) = 0 And" & vbNewLine & _
'                    "      A.��� = 'C' And A.ID = [2] And A.�������� = [1]" & vbNewLine & _
'                    "Union" & vbNewLine & _
'                    "" & vbNewLine & _
'                    "Select E.������Ŀid, E.��д, B.������" & vbNewLine & _
'                    "From ������Ŀ E, ������ĿĿ¼ D, ���鱨����Ŀ C, ����������Ŀ B, ����������Ŀ A" & vbNewLine & _
'                    "Where E.������Ŀid = C.������Ŀid And C.������Ŀid = D.ID And B.ID = C.������Ŀid And A.��Ŀid = B.ID And D.��� = 'C' And" & vbNewLine & _
'                    "      Nvl(D.�����Ŀ, 0) = 0 And Nvl(D.����Ӧ��, 0) = 1 And A.����id = [3] And D.�������� = [1]"
            strSQL = "Select E.������Ŀid, E.��д, D.������,d.���� " & vbNewLine & _
                    "From ������Ŀ E, ����������Ŀ D, ���鱨����Ŀ C" & vbNewLine & _
                    "Where C.������Ŀid = D.ID And D.ID = E.������Ŀid And C.������Ŀid = [2]" & vbNewLine & _
                    "Union" & vbNewLine & _
                    "" & vbNewLine & _
                    "Select E.������Ŀid, E.��д, B.������,d.���� " & vbNewLine & _
                    "From ������Ŀ E, ������ĿĿ¼ D, ���鱨����Ŀ C, ����������Ŀ B, ����������Ŀ A" & vbNewLine & _
                    "Where E.������Ŀid = C.������Ŀid And C.������Ŀid = D.ID And B.ID = C.������Ŀid And A.��Ŀid = B.ID And D.��� = 'C' And" & vbNewLine & _
                    "      Nvl(D.�����Ŀ, 0) = 0  And A.����id = [3] And D.�������� = [1]"

        End If
        Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, "" & rsGroup.Fields("����"), mlngID, mlng����ID)
        Do Until rsTmp.EOF
            mstrItem = mstrItem & "[" & IIf("" & rsTmp.Fields("��д") = "", rsTmp.Fields("������ĿID"), rsTmp.Fields("����") & "_" & rsTmp.Fields("��д")) & "]" & ","
            tvwItem.Nodes.Add "" & rsGroup.Fields("����"), tvwChild, "K" & rsGroup.Fields("����") & "_" & rsTmp.Fields("������ĿID"), _
            "[" & IIf("" & rsTmp.Fields("��д") = "", rsTmp.Fields("������ĿID"), rsTmp.Fields("����") & "_" & rsTmp.Fields("��д")) & "]" & rsTmp.Fields("������")
            rsTmp.MoveNext
        Loop
        rsGroup.MoveNext
    Loop
    cmdOk.Enabled = False
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tvwItem_DblClick()
    If InStr(tvwItem.SelectedItem.Text, "]") > 0 Then
        txtFormula.SelText = Mid(tvwItem.SelectedItem.Text, 1, InStr(tvwItem.SelectedItem.Text, "]"))
    End If
End Sub

Private Sub txtFormula_Change()
    If Trim(txtFormula.Text) <> "" Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
End Sub

Private Sub cmdCheck_Click()
    If Trim(txtFormula.Text) = "" Then cmdOk.Enabled = True: Exit Sub
    If CheckRule(txtFormula, mstrItem) Then
        cmdOk.Enabled = True
    End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(txtFormula.Text) = "" Then mstrFormula = "": Unload Me: Exit Sub
    If CheckRule(txtFormula, mstrItem) Then
        mstrFormula = txtFormula
        Unload Me
    Else
         MsgBox "�������ô���", vbExclamation, Me.Caption
    End If
End Sub

'-----------------------------------------------------------------
'-- ������ �Զ������
'-----------------------------------------------------------------

Public Function DefFormula(ByVal lngID As Long, ByVal lng����ID As Long, ByVal strFormula As String, ByVal frmMain As Form) As String
    '���ܣ��������
    'lngID :��ǰ�����ļ�����Ŀ ID
    'strFormula :ԭ���Ĺ�ʽ
    'frmMain: ���ô���
    mlngID = lngID: mlng����ID = lng����ID
    mstrFormula = strFormula
    
    Me.Show vbModal, frmMain
    DefFormula = mstrFormula
End Function


