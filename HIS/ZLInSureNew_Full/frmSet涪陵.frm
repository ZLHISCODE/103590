VERSION 5.00
Begin VB.Form frmSet���� 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   Icon            =   "frmSet����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmd��ȡ��Ϣ 
      Caption         =   "��"
      Height          =   285
      Left            =   3300
      TabIndex        =   5
      Top             =   630
      Width           =   285
   End
   Begin VB.TextBox txtҽԺ���� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1725
      TabIndex        =   7
      Top             =   1020
      Width           =   1845
   End
   Begin VB.TextBox txtҽ���������� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1725
      TabIndex        =   4
      Top             =   630
      Width           =   1575
   End
   Begin VB.CommandButton cmdTrans 
      Caption         =   "�ϴ�"
      Height          =   350
      Left            =   120
      TabIndex        =   16
      Top             =   2490
      Width           =   1100
   End
   Begin VB.CheckBox chk��λ 
      Caption         =   "�ϴ���λ��Ϣ"
      Height          =   210
      Left            =   2790
      TabIndex        =   12
      Top             =   1950
      Width           =   2295
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "�ϴ�������Ŀ��Ϣ"
      Height          =   210
      Left            =   420
      TabIndex        =   11
      Top             =   1950
      Width           =   2295
   End
   Begin VB.CheckBox chkҩƷ 
      Caption         =   "�ϴ�ҩƷ������Ϣ"
      Height          =   210
      Left            =   2790
      TabIndex        =   10
      Top             =   1620
      Width           =   2295
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "�ϴ�����������Ϣ"
      Height          =   210
      Left            =   420
      TabIndex        =   9
      Top             =   1620
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   1410
      Width           =   5265
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3270
      TabIndex        =   15
      Top             =   2490
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2010
      TabIndex        =   14
      Top             =   2490
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   2310
      Width           =   5265
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1725
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "1"
      Top             =   240
      Width           =   360
   End
   Begin VB.Label lblҽԺ���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҽԺ����(&H)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   630
      TabIndex        =   6
      Top             =   1080
      Width           =   990
   End
   Begin VB.Label lblҽ���������� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ����������(&Y)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   270
      TabIndex        =   3
      Top             =   690
      Width           =   1350
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "�Ŵ���"
      Height          =   180
      Index           =   4
      Left            =   2130
      TabIndex        =   2
      Top             =   300
      Width           =   540
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "��ǰ����(&D)"
      Height          =   180
      Index           =   3
      Left            =   630
      TabIndex        =   0
      Top             =   300
      Width           =   990
   End
End
Attribute VB_Name = "frmSet����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnReturn As Boolean
Private mlng���� As Long
 
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(txtEdit) = "" Then Exit Sub
    
    gcnOracle.BeginTrans
    On Error GoTo errHand
    
    'ɾ���Ѿ�����
    gstrSQL = "zl_���ղ���_Delete(" & mlng���� & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '������������
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",NULL,'�˿ں�','" & txtEdit.Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",NULL,'ҽ����������','" & txtҽ����������.Text & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",NULL,'ҽԺ����','" & txtҽԺ����.Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gcnOracle.CommitTrans
    gintComPort = txtEdit.Text
    mblnReturn = True
    
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Sub cmdTrans_Click()
    Dim rsTemp As New ADODB.Recordset, iLoop As Long, strTemp As String
    On Error GoTo errHand
'    gstrҽ���������� = "500102"
'    gstrҽԺ���� = "5001020003"
    If gstrҽ���������� = "" Then
        MsgBox "׼����ȡҽ���������룬�����ϵͳ�����˿�", vbInformation, gstrSysName
CheckCard:
        initType
        Call DebugTool("��ȡϵͳ�������ú�����getybjgbm")
        mblnReturn = fl_getybjgbm(gstrOutPara)
        Call DebugTool("�ɹ���")
        TrimType
        If mblnReturn = False Then
            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
                GoTo CheckCard
            Else
                Exit Sub
            End If
        End If
        gstrҽ���������� = gstrOutPara.out1
        gstrҽԺ���� = gstrOutPara.out2
        txtҽ����������.Text = gstrҽ����
        txtҽԺ����.Text = gstrҽԺ����
    End If
    If chk����.Value = 1 Then
        Call DebugTool("׼���ϴ����ղ���")
        gstrSQL = "Select id as ����,���� From ���ղ���"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName)
        chk����.Caption = "�ϴ�����������Ϣ(0/" & rsTemp.RecordCount & ")"
        iLoop = 0
        DoEvents
        Call DebugTool("���ղ��ּ�¼����" & rsTemp.RecordCount)
        While Not rsTemp.EOF
            initType
            Call DebugTool("��������wyyglxx����Σ�" & gstrҽ���������� & "," & gstrҽԺ���� & "," & "0" & "," & rsTemp!���� & "," & rsTemp!���� & "," & "")
            mblnReturn = fl_wyyglxx(gstrҽ����������, gstrҽԺ����, "0", rsTemp!����, rsTemp!����, "", gstrOutPara)
            Call DebugTool("���óɹ���")
            rsTemp.MoveNext
            iLoop = iLoop + 1
            chk����.Caption = "�ϴ�����������Ϣ(" & iLoop & "/" & rsTemp.RecordCount & ")"
            DoEvents
        Wend
        chk����.Value = 0
    End If
    If chkҩƷ.Value = 1 Then
        gstrSQL = " Select a.��� as ���,a.id as ����,a.���� as ����,b.ҩƷ��Դ as ҩƷ��Դ " & _
                  " From �շ�ϸĿ a,ҩƷĿ¼ b " & _
                  " Where a.��� In ('5','6','7') and a.ID=b.ҩƷID" & _
                  " And (A.����ʱ�� Is NULL Or to_char(A.����ʱ��,'yyyy-MM-dd')='3000-01-01')"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName)
        chkҩƷ.Caption = "�ϴ�ҩƷ������Ϣ(0/" & rsTemp.RecordCount & ")"
        iLoop = 0
        DoEvents
        While Not rsTemp.EOF
            initType
            mblnReturn = fl_wyyglxx(gstrҽ����������, gstrҽԺ����, "1", rsTemp!��� & "_" & rsTemp!����, rsTemp!����, IIf(rsTemp!ҩƷ��Դ = "����", "03", "02"), gstrOutPara)
            rsTemp.MoveNext
            iLoop = iLoop + 1
            chkҩƷ.Caption = "�ϴ�ҩƷ������Ϣ(" & iLoop & "/" & rsTemp.RecordCount & ")"
            DoEvents
        Wend
        chkҩƷ.Value = 0
    End If
    If chk����.Value = 1 Then
        gstrSQL = "select * from �շ�ϸĿ where ��� Not In ('J','5','6','7')" & _
            " And (����ʱ�� Is NULL Or to_char(����ʱ��,'yyyy-MM-dd')='3000-01-01')"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName)
        chk����.Caption = "�ϴ�������Ŀ��Ϣ(0/" & rsTemp.RecordCount & ")"
        iLoop = 0
        DoEvents
        While Not rsTemp.EOF
            initType
            mblnReturn = fl_wyyglxx(gstrҽ����������, gstrҽԺ����, "2", rsTemp!��� & "_" & rsTemp!ID, rsTemp!����, "", gstrOutPara)
            rsTemp.MoveNext
            iLoop = iLoop + 1
            chk����.Caption = "�ϴ�������Ŀ��Ϣ(" & iLoop & "/" & rsTemp.RecordCount & ")"
            DoEvents
        Wend
        chk����.Value = 0
    End If
    If chk��λ.Value = 1 Then
        gstrSQL = "select * from �շ�ϸĿ where ���='J'" & _
            " And (����ʱ�� Is NULL Or to_char(����ʱ��,'yyyy-MM-dd')='3000-01-01')"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName)
        chk��λ.Caption = "�ϴ���λ��Ϣ(0/" & rsTemp.RecordCount & ")"
        iLoop = 0
        DoEvents
        While Not rsTemp.EOF
            initType
            mblnReturn = fl_wyyglxx(gstrҽ����������, gstrҽԺ����, "3", rsTemp!��� & "_" & rsTemp!ID, rsTemp!����, " ", gstrOutPara)
            rsTemp.MoveNext
            iLoop = iLoop + 1
            chk��λ.Caption = "�ϴ���λ��Ϣ(" & iLoop & "/" & rsTemp.RecordCount & ")"
            DoEvents
        Wend
        chk��λ.Value = 0
    End If
    MsgBox "������Ŀ��Ϣ�ϴ����", vbInformation, gstrSysName
    Exit Sub
errHand:
    MsgBox "����ţ�" & Err.Number & vbCrLf & "������Ϣ��" & Err.Description
End Sub

Private Sub cmd��ȡ��Ϣ_Click()
    MsgBox "׼����ȡҽ���������룬�����ϵͳ�����˿�", vbInformation, gstrSysName
CheckCard:
    initType
    mblnReturn = fl_getybjgbm(gstrOutPara)
    TrimType
    If mblnReturn = False Then
        If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
            GoTo CheckCard
        Else
            Exit Sub
        End If
    End If
    gstrҽ���������� = gstrOutPara.out1
    gstrҽԺ���� = gstrOutPara.out2
    txtҽ����������.Text = gstrҽ����������
    txtҽԺ����.Text = gstrҽԺ����
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    mblnReturn = False
    
    gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����", mlng����)
    
    With rsTemp
        Do While Not .EOF
            If !������ = "�˿ں�" Then
                txtEdit.Text = Nvl(!����ֵ)
            ElseIf !������ = "ҽ����������" Then
                txtҽ����������.Text = Nvl(!����ֵ)
            Else
                txtҽԺ����.Text = Nvl(!����ֵ)
            End If
            .MoveNext
        Loop
    End With
End Sub

Public Function ShowME(ByVal lng���� As Long) As Boolean
    mlng���� = lng����
    Me.Show 1
    ShowME = mblnReturn
End Function
