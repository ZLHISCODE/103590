VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLabVerifyEspecial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   10350
   Icon            =   "frmLabVerifyEspecial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraRule5 
      Caption         =   "���ϴν����ǱȽ�"
      Height          =   3750
      Left            =   3555
      TabIndex        =   37
      Top             =   2730
      Width           =   6630
      Begin VB.TextBox txtLastTag 
         Height          =   2300
         Left            =   3015
         TabIndex        =   39
         Top             =   510
         Width           =   3300
      End
      Begin MSComctlLib.TreeView tvwLastTag 
         Height          =   2295
         Left            =   135
         TabIndex        =   38
         Top             =   510
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   4048
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   459
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         Appearance      =   1
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"frmLabVerifyEspecial.frx":000C
         ForeColor       =   &H0000C000&
         Height          =   720
         Left            =   225
         TabIndex        =   42
         Top             =   2850
         Width           =   5400
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʽ�༭"
         Height          =   210
         Left            =   3015
         TabIndex        =   41
         Top             =   300
         Width           =   885
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ŀ�б�"
         Height          =   195
         Left            =   150
         TabIndex        =   40
         Top             =   300
         Width           =   885
      End
   End
   Begin MSComctlLib.TreeView tvwItem 
      Height          =   6795
      Left            =   165
      TabIndex        =   0
      Top             =   195
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   11986
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   617
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "��֤(&C)"
      Height          =   350
      Left            =   3630
      TabIndex        =   21
      Top             =   6615
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7155
      TabIndex        =   20
      Top             =   6600
      Width           =   1100
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "ȡ��(&E)"
      Height          =   350
      Left            =   8655
      TabIndex        =   19
      Top             =   6600
      Width           =   1100
   End
   Begin VB.TextBox txt��ʽ 
      Height          =   2000
      Left            =   3555
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   195
      Width           =   6645
   End
   Begin VB.OptionButton optOr 
      Caption         =   "OR"
      Height          =   270
      Left            =   8220
      TabIndex        =   3
      Top             =   2250
      Value           =   -1  'True
      Width           =   675
   End
   Begin VB.OptionButton optAnd 
      Caption         =   "AND"
      Height          =   270
      Left            =   7500
      TabIndex        =   2
      Top             =   2250
      Width           =   705
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "����(&A)"
      Height          =   350
      Left            =   8955
      TabIndex        =   1
      Top             =   2220
      Width           =   1100
   End
   Begin VB.Frame FraRule2 
      Caption         =   "���ϴν���Ƚ�"
      Height          =   3750
      Left            =   3570
      TabIndex        =   4
      Top             =   2730
      Width           =   6630
      Begin VB.TextBox txtLast 
         Height          =   2300
         Left            =   3120
         TabIndex        =   5
         Top             =   495
         Width           =   3300
      End
      Begin MSComctlLib.TreeView tvwLast 
         Height          =   2295
         Left            =   210
         TabIndex        =   25
         Top             =   495
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   4048
         _Version        =   393217
         Indentation     =   459
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         Appearance      =   1
      End
      Begin VB.Label Label9 
         Caption         =   "��Ŀ�б�"
         Height          =   195
         Left            =   225
         TabIndex        =   27
         Top             =   285
         Width           =   885
      End
      Begin VB.Label Label8 
         Caption         =   "��ʽ�༭"
         Height          =   210
         Left            =   3150
         TabIndex        =   26
         Top             =   285
         Width           =   885
      End
      Begin VB.Label Label7 
         Caption         =   $"frmLabVerifyEspecial.frx":00EF
         ForeColor       =   &H0000C000&
         Height          =   765
         Left            =   420
         TabIndex        =   6
         Top             =   2850
         Width           =   5970
      End
   End
   Begin VB.Frame FraRule1 
      Caption         =   "���ΪX�ĳ���N��"
      Height          =   3750
      Left            =   3540
      TabIndex        =   11
      Top             =   2730
      Width           =   6630
      Begin VB.ComboBox cbo��Ŀ���� 
         Height          =   300
         ItemData        =   "frmLabVerifyEspecial.frx":01CB
         Left            =   3345
         List            =   "frmLabVerifyEspecial.frx":01E1
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   585
         Width           =   810
      End
      Begin VB.ComboBox cbo������ 
         Height          =   300
         ItemData        =   "frmLabVerifyEspecial.frx":0213
         Left            =   195
         List            =   "frmLabVerifyEspecial.frx":022C
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3300
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.TextBox txt��Ŀ���� 
         Height          =   285
         Left            =   4140
         TabIndex        =   13
         Top             =   585
         Width           =   1080
      End
      Begin VB.TextBox txt������ 
         Height          =   285
         Left            =   1410
         TabIndex        =   12
         Top             =   585
         Width           =   960
      End
      Begin VB.Label Label12 
         Caption         =   "��ʱ����ʾ��"
         Height          =   210
         Left            =   5265
         TabIndex        =   24
         Top             =   630
         Width           =   1080
      End
      Begin VB.Label Label3 
         Caption         =   "����Ŀ����"
         Height          =   210
         Left            =   2400
         TabIndex        =   16
         Top             =   630
         Width           =   930
      End
      Begin VB.Label Label2 
         Caption         =   "����������"
         Height          =   210
         Left            =   270
         TabIndex        =   15
         Top             =   630
         Width           =   1080
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmLabVerifyEspecial.frx":0264
         ForeColor       =   &H0000C000&
         Height          =   1065
         Left            =   1485
         TabIndex        =   14
         Top             =   1770
         Width           =   5025
      End
   End
   Begin VB.Frame fraRule4 
      Caption         =   "©�������"
      Height          =   3750
      Left            =   3555
      TabIndex        =   33
      Top             =   2715
      Width           =   6630
      Begin VB.ComboBox cbo��鷽ʽ 
         Height          =   300
         ItemData        =   "frmLabVerifyEspecial.frx":02CE
         Left            =   1545
         List            =   "frmLabVerifyEspecial.frx":02DB
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   480
         Width           =   1725
      End
      Begin VB.Label Label13 
         Caption         =   "��鷽ʽ"
         Height          =   225
         Left            =   660
         TabIndex        =   36
         Top             =   525
         Width           =   810
      End
      Begin VB.Label Label16 
         Caption         =   "���磺�������صĽ��ȱ��RBCʱ������ʾ��"
         ForeColor       =   &H0000C000&
         Height          =   285
         Left            =   330
         TabIndex        =   35
         Top             =   2235
         Width           =   5040
      End
   End
   Begin VB.Frame FraRule3 
      Caption         =   "��������Ŀ�⣬���ΪX"
      Height          =   3750
      Left            =   3555
      TabIndex        =   7
      Top             =   2730
      Width           =   6630
      Begin VB.ComboBox cboNot���� 
         Height          =   300
         ItemData        =   "frmLabVerifyEspecial.frx":0301
         Left            =   2865
         List            =   "frmLabVerifyEspecial.frx":0317
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   2625
         Width           =   810
      End
      Begin VB.TextBox txtNotֵ 
         Height          =   285
         Left            =   3735
         TabIndex        =   30
         Top             =   2625
         Width           =   1080
      End
      Begin VB.TextBox txtNot��Ŀ 
         Height          =   1770
         Left            =   2865
         TabIndex        =   8
         Top             =   495
         Width           =   3570
      End
      Begin MSComctlLib.TreeView tvwNot 
         Height          =   2445
         Left            =   210
         TabIndex        =   28
         Top             =   495
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   4313
         _Version        =   393217
         Indentation     =   459
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         Appearance      =   1
      End
      Begin VB.Label Label5 
         Caption         =   "��������Ŀ�⣬������Ŀ�Ľ���У������"
         Height          =   210
         Left            =   2865
         TabIndex        =   31
         Top             =   2370
         Width           =   3540
      End
      Begin VB.Label Label11 
         Caption         =   "��Ŀ�б�"
         Height          =   195
         Left            =   225
         TabIndex        =   29
         Top             =   255
         Width           =   885
      End
      Begin VB.Label Label6 
         Caption         =   "����Ŀ������ʾ��"
         Height          =   180
         Left            =   4905
         TabIndex        =   10
         Top             =   2685
         Width           =   1470
      End
      Begin VB.Label Label4 
         Caption         =   "���磺��BEact,Beecf��,������Ŀ�Ľ���и����ģ�����ʾ��"
         ForeColor       =   &H0000C000&
         Height          =   285
         Left            =   345
         TabIndex        =   9
         Top             =   3150
         Width           =   5040
      End
   End
   Begin VB.Label Label10 
      Caption         =   "�¼��������ԭ����Ĺ�ϵ"
      Height          =   195
      Left            =   5220
      TabIndex        =   18
      Top             =   2280
      Width           =   2205
   End
End
Attribute VB_Name = "frmLabVerifyEspecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngID As Long
Private mlng����ID As Long
Private mstrFormula As String
Private mStrItem As String

Private Sub cmdAdd_Click()
    Dim strAndOr As String
    
    If InStr(txt��ʽ.Text, "{") > 0 Then strAndOr = IIf(optAnd, " AND ", " OR ")
    If FraRule1.Visible Then
        txt��ʽ.Text = txt��ʽ.Text & strAndOr & GenFormula("A", mStrItem, txt������, Gen����(cbo��Ŀ����), txt��Ŀ����)
    ElseIf FraRule2.Visible Then
        txt��ʽ.Text = txt��ʽ.Text & strAndOr & GenFormula("B", mStrItem, txtLast)
    ElseIf FraRule3.Visible Then
        txt��ʽ.Text = txt��ʽ.Text & strAndOr & GenFormula("C", Replace(mStrItem, "�ϴ�.", ""), txtNot��Ŀ, Gen����(cboNot����), txtNotֵ)
    ElseIf fraRule4.Visible Then
        txt��ʽ.Text = txt��ʽ.Text & strAndOr & GenFormula("D", mStrItem, cbo��鷽ʽ)
    ElseIf fraRule5.Visible Then
        txt��ʽ.Text = txt��ʽ.Text & strAndOr & GenFormula("E", mStrItem, txtLastTag)
    End If
End Sub

Private Sub cmdCheck_Click()
    If Trim(Me.txt��ʽ.Text) = "" Then cmdOk.Enabled = True: Exit Sub
    If CheckEspecial(txt��ʽ, mStrItem) Then
        cmdOk.Enabled = True
    End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(Me.txt��ʽ.Text) = "" Then mstrFormula = "": Unload Me: Exit Sub
    If CheckEspecial(txt��ʽ, mStrItem) Then
        mstrFormula = txt��ʽ
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim itemTmp As ListItem
    Dim rsGroup As ADODB.Recordset
    Dim strItem As String
    
    On Error GoTo ErrHandle
    txt��ʽ = mstrFormula
    tvwLast.Nodes.Clear
    tvwNot.Nodes.Clear
    tvwLastTag.Nodes.Clear
    If mlngID = 0 And mlng����ID = 0 Then
        strSQL = "Select ����||'-'||���� as ��ʾ���� ,���� ,���� From ���Ƽ������� where ���� IN (" & vbNewLine & _
                        "Select D.�������� From ������Ŀ A, ����������Ŀ B, ������ĿĿ¼ D, ���鱨����Ŀ C" & vbNewLine & _
                        "Where A.������Ŀid = B.ID And B.ID = C.������Ŀid And C.������Ŀid = D.ID And D.��� = 'C'  And" & vbNewLine & _
                        "      Nvl(D.�����Ŀ, 0) = 0 )"
    Else
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
    Set rsGroup = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID, mlng����ID)
    mStrItem = ","
    Do Until rsGroup.EOF
        tvwLast.Nodes.Add , , "" & rsGroup.Fields("����"), "" & rsGroup.Fields("��ʾ����")
        tvwNot.Nodes.Add , , "" & rsGroup.Fields("����"), "" & rsGroup.Fields("��ʾ����")
        tvwLastTag.Nodes.Add , , "" & rsGroup.Fields("����"), "" & rsGroup.Fields("��ʾ����")
        If mlngID = 0 And mlng����ID = 0 Then
            strSQL = "Select Distinct A.������Ŀid, A.��д, B.������,b.���� " & vbNewLine & _
                    "From ������Ŀ A, ����������Ŀ B, ������ĿĿ¼ D, ���鱨����Ŀ C" & vbNewLine & _
                    "Where A.������Ŀid = B.ID And B.ID = C.������Ŀid And C.������Ŀid = D.ID And D.��� = 'C'  And" & vbNewLine & _
                    "      Nvl(D.�����Ŀ, 0) = 0 And D.�������� = [1]"
        Else
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
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "" & rsGroup.Fields("����"), mlngID, mlng����ID)
        Do Until rsTmp.EOF
            mStrItem = mStrItem & "[" & IIf("" & rsTmp.Fields("��д") = "", rsTmp.Fields("������ĿID"), rsTmp.Fields("����") & "_" & rsTmp.Fields("��д")) & "]" & ","
            mStrItem = mStrItem & "[�ϴ�." & IIf("" & rsTmp.Fields("��д") = "", rsTmp.Fields("������ĿID"), rsTmp.Fields("����") & "_" & rsTmp.Fields("��д")) & "]" & ","
            mStrItem = mStrItem & "[���." & IIf("" & rsTmp.Fields("��д") = "", rsTmp.Fields("������ĿID"), rsTmp.Fields("����") & "_" & rsTmp.Fields("��д")) & "]" & ","
            
            '-- B�����Ŀ�ѡ��
            tvwLast.Nodes.Add "" & rsGroup.Fields("����"), tvwChild, "K" & rsGroup.Fields("����") & "_" & rsTmp.Fields("������ĿID"), "[" & _
                IIf("" & rsTmp.Fields("��д") = "", rsTmp.Fields("������ĿID"), rsTmp.Fields("����") & "_" & rsTmp.Fields("��д")) & "]" & rsTmp.Fields("������")
            tvwLast.Nodes.Add "" & rsGroup.Fields("����"), tvwChild, "KL" & rsGroup.Fields("����") & "_" & rsTmp.Fields("������ĿID"), "[�ϴ�." & _
                IIf("" & rsTmp.Fields("��д") = "", rsTmp.Fields("������ĿID"), rsTmp.Fields("����") & "_" & rsTmp.Fields("��д")) & "]" & rsTmp.Fields("������")
            '-- C�����Ŀ�ѡ��
            tvwNot.Nodes.Add "" & rsGroup.Fields("����"), tvwChild, "K" & rsGroup.Fields("����") & "_" & rsTmp.Fields("������ĿID"), "[" & _
                IIf("" & rsTmp.Fields("��д") = "", rsTmp.Fields("������ĿID"), rsTmp.Fields("����") & "_" & rsTmp.Fields("��д")) & "]" & rsTmp.Fields("������")
            '-- E�����Ŀ�ѡ��
            tvwLastTag.Nodes.Add "" & rsGroup.Fields("����"), tvwChild, "K" & rsGroup.Fields("����") & "_" & rsTmp.Fields("������ĿID"), "[���." & _
                 IIf("" & rsTmp.Fields("��д") = "", rsTmp.Fields("������ĿID"), rsTmp.Fields("����") & "_" & rsTmp.Fields("��д")) & "]" & rsTmp.Fields("������")
            tvwLastTag.Nodes.Add "" & rsGroup.Fields("����"), tvwChild, "KL" & rsGroup.Fields("����") & "_" & rsTmp.Fields("������ĿID"), "[�ϴ�." & _
                 IIf("" & rsTmp.Fields("��д") = "", rsTmp.Fields("������ĿID"), rsTmp.Fields("����") & "_" & rsTmp.Fields("��д")) & "]" & rsTmp.Fields("������")
            rsTmp.MoveNext
        Loop
        rsGroup.MoveNext
    Loop
    
    '----
    Dim nodX As Node
    tvwItem.Nodes.Clear
    tvwItem.LabelEdit = tvwManual
    Set nodX = tvwItem.Nodes.Add(, , "R1", "{A:X|N}�����ΪX����N��")
    Set nodX = tvwItem.Nodes.Add(, , "R2", "{B:P} ���ϴν���Ƚ�")
    Set nodX = tvwItem.Nodes.Add(, , "R3", "{C:not N|X} ��N����,���ΪX")
    Set nodX = tvwItem.Nodes.Add(, , "R4", "{D:X} ©�������")
    Set nodX = tvwItem.Nodes.Add(, , "R5", "{E:P} ���ϴν����־�Ƚ�")
    cbo������.ListIndex = 0: cbo��Ŀ����.ListIndex = 0: cboNot����.ListIndex = 0: cbo��鷽ʽ.ListIndex = 2
    Call tvwItem_NodeClick(tvwItem.Nodes("R1"))
    tvwItem.Nodes("R1").Selected = True
    
    cmdOk.Enabled = False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tvwItem_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Key = "R1" Then
        FraRule1.Visible = True
        FraRule2.Visible = False
        FraRule3.Visible = False
        fraRule4.Visible = False
        fraRule5.Visible = False
    ElseIf Node.Key = "R2" Then
        FraRule1.Visible = False
        FraRule2.Visible = True
        FraRule3.Visible = False
        fraRule4.Visible = False
        fraRule5.Visible = False
    ElseIf Node.Key = "R3" Then
        FraRule1.Visible = False
        FraRule2.Visible = False
        FraRule3.Visible = True
        fraRule4.Visible = False
        fraRule5.Visible = False
    ElseIf Node.Key = "R4" Then
        FraRule1.Visible = False
        FraRule2.Visible = False
        FraRule3.Visible = False
        fraRule4.Visible = True
        fraRule5.Visible = False
    Else
        FraRule1.Visible = False
        FraRule2.Visible = False
        FraRule3.Visible = False
        fraRule4.Visible = False
        fraRule5.Visible = True
    End If
End Sub

Private Sub tvwLast_DblClick()
    If InStr(tvwLast.SelectedItem.Text, "]") > 0 Then
        txtLast.SelText = Mid(tvwLast.SelectedItem.Text, 1, InStr(tvwLast.SelectedItem.Text, "]"))
    End If
End Sub

Private Sub tvwLastTag_DblClick()
    If InStr(tvwLastTag.SelectedItem.Text, "]") > 0 Then
        txtLastTag.SelText = Mid(tvwLastTag.SelectedItem.Text, 1, InStr(tvwLastTag.SelectedItem.Text, "]"))
    End If
End Sub

Private Sub tvwNot_DblClick()
    If InStr(tvwNot.SelectedItem.Text, "]") > 0 Then
        txtNot��Ŀ = IIf(txtNot��Ŀ = "", "", txtNot��Ŀ & ",") & Mid(tvwNot.SelectedItem.Text, 1, InStr(tvwNot.SelectedItem.Text, "]"))
    End If
End Sub

'-----------------------------
'-- �������Զ������
'------------------------------
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


Private Function Gen����(ByVal strIn As String) As String
    Select Case strIn
    Case "����": Gen���� = "="
    Case "����": Gen���� = ">"
    Case "С��": Gen���� = "<"
    Case "���ڵ���": Gen���� = ">="
    Case "С�ڵ���": Gen���� = "<="
    Case "������": Gen���� = "<>"
    Case "����": Gen���� = " Like "
    End Select
End Function





Private Sub txt��ʽ_Change()
    If Trim(txt��ʽ.Text) = "" Then
        Me.cmdOk.Enabled = True
    End If
End Sub
