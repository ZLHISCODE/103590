VERSION 5.00
Begin VB.Form frmSet��Ԫ 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4995
   Icon            =   "frmSet��Ԫ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkʵʱ�ϴ� 
      Caption         =   "ʵʱ�ϴ�������ϸ(&R)"
      Height          =   195
      Left            =   720
      TabIndex        =   8
      Top             =   1440
      Value           =   1  'Checked
      Width           =   2955
   End
   Begin VB.TextBox txtҽ���������� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1815
      TabIndex        =   4
      Top             =   630
      Width           =   1575
   End
   Begin VB.TextBox txtҽԺ���� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1815
      TabIndex        =   7
      Top             =   1020
      Width           =   1845
   End
   Begin VB.CommandButton cmd��ȡ��Ϣ 
      Caption         =   "��"
      Height          =   285
      Left            =   3390
      TabIndex        =   5
      Top             =   630
      Width           =   285
   End
   Begin VB.CommandButton cmdTrans 
      Caption         =   "�ϴ�"
      Height          =   350
      Left            =   120
      TabIndex        =   15
      Top             =   2910
      Width           =   1100
   End
   Begin VB.CheckBox chk��λ 
      Caption         =   "�ϴ���λ��Ϣ"
      Height          =   210
      Left            =   2400
      TabIndex        =   12
      Top             =   2250
      Width           =   1815
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "�ϴ�������Ŀ��Ϣ"
      Height          =   210
      Left            =   420
      TabIndex        =   11
      Top             =   2250
      Width           =   1815
   End
   Begin VB.CheckBox chkҩƷ 
      Caption         =   "�ϴ�ҩƷ������Ϣ"
      Height          =   210
      Left            =   2400
      TabIndex        =   10
      Top             =   1950
      Width           =   1815
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "�ϴ�����������Ϣ"
      Height          =   210
      Left            =   420
      TabIndex        =   9
      Top             =   1950
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Index           =   1
      Left            =   0
      TabIndex        =   17
      Top             =   1740
      Width           =   5265
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3840
      TabIndex        =   14
      Top             =   2910
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2700
      TabIndex        =   13
      Top             =   2910
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Index           =   0
      Left            =   30
      TabIndex        =   16
      Top             =   2670
      Width           =   5265
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1815
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "1"
      Top             =   240
      Width           =   360
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
      Left            =   360
      TabIndex        =   3
      Top             =   690
      Width           =   1350
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
      Left            =   720
      TabIndex        =   6
      Top             =   1080
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "�Ŵ���"
      Height          =   180
      Index           =   4
      Left            =   2220
      TabIndex        =   2
      Top             =   300
      Width           =   540
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "��ǰ����(&D)"
      Height          =   180
      Index           =   3
      Left            =   720
      TabIndex        =   0
      Top             =   300
      Width           =   990
   End
End
Attribute VB_Name = "frmSet��Ԫ"
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
    If Trim(TxtEdit) = "" Then Exit Sub
    
    gcnOracle.BeginTrans
    On Error GoTo errHand
    
    'ɾ���Ѿ�����
    gstrSQL = "zl_���ղ���_Delete(" & mlng���� & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '������������
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",NULL,'�˿ں�','" & TxtEdit.Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",NULL,'ҽ����������','" & txtҽ����������.Text & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",NULL,'ҽԺ����','" & txtҽԺ����.Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",NULL,'ʵʱ�ϴ�','" & chkʵʱ�ϴ�.Value & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gcnOracle.CommitTrans
    gintComPort = TxtEdit.Text
    mblnReturn = True
    
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Sub cmdTrans_Click()
    Dim rsTemp As New ADODB.Recordset, iLoop As Long, strTemp As String
'    gstrҽ���������� = "500102"
'    gstrҽԺ���� = "5001020003"
    If gstrҽ���������� = "" Then
        MsgBox "׼����ȡҽ���������룬�����ϵͳ�����˿�", vbInformation, gstrSysName
CheckCard:
        initType
        mblnReturn = gy_getybjgbm(gstrOutPara)
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
    End If
    If chk����.Value = 1 Then
        gstrSQL = "Select id as ����,���� From ���ղ���"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName)
        chk����.Caption = "�ϴ�����������Ϣ(0/" & rsTemp.RecordCount & ")"
        iLoop = 0
        DoEvents
        While Not rsTemp.EOF
            initType
            mblnReturn = gy_wyyglxx(gstrҽ����������, gstrҽԺ����, "0", rsTemp!����, rsTemp!����, "", gstrOutPara)
            rsTemp.MoveNext
            iLoop = iLoop + 1
            chk����.Caption = "�ϴ�����������Ϣ(" & iLoop & "/" & rsTemp.RecordCount & ")"
            DoEvents
        Wend
        chk����.Value = 0
    End If
    If chkҩƷ.Value = 1 Then
        gstrSQL = "select a.��� as ���,a.id as ����,a.���� as ����,b.ҩƷ��Դ as ҩƷ��Դ from �շ�ϸĿ a,ҩƷĿ¼ b " & _
            " Where a.��� In ('5','6','7') and a.����=b.����" & _
            " And (A.����ʱ�� Is NULL Or to_char(A.����ʱ��,'yyyy-MM-dd')='3000-01-01')"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName)
        chkҩƷ.Caption = "�ϴ�ҩƷ������Ϣ(0/" & rsTemp.RecordCount & ")"
        iLoop = 0
        DoEvents
        While Not rsTemp.EOF
            initType
            mblnReturn = gy_wyyglxx(gstrҽ����������, gstrҽԺ����, "1", rsTemp!��� & "_" & rsTemp!����, rsTemp!����, IIf(rsTemp!ҩƷ��Դ = "����", "01", "03"), gstrOutPara)
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
            mblnReturn = gy_wyyglxx(gstrҽ����������, gstrҽԺ����, "2", rsTemp!��� & "_" & rsTemp!ID, rsTemp!����, "", gstrOutPara)
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
            mblnReturn = gy_wyyglxx(gstrҽ����������, gstrҽԺ����, "3", rsTemp!��� & "_" & rsTemp!ID, rsTemp!����, " ", gstrOutPara)
            rsTemp.MoveNext
            iLoop = iLoop + 1
            chk��λ.Caption = "�ϴ���λ��Ϣ(" & iLoop & "/" & rsTemp.RecordCount & ")"
            DoEvents
        Wend
        chk��λ.Value = 0
    End If
    MsgBox "������Ŀ��Ϣ�ϴ����", vbInformation, gstrSysName
End Sub

Private Sub cmd��ȡ��Ϣ_Click()
    MsgBox "׼����ȡҽ���������룬�����ϵͳ�����˿�", vbInformation, gstrSysName
CheckCard:
    initType
    mblnReturn = gy_getybjgbm(gstrOutPara)
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
                TxtEdit.Text = Nvl(!����ֵ)
            ElseIf !������ = "ҽ����������" Then
                txtҽ����������.Text = Nvl(!����ֵ)
            ElseIf !������ = "ҽԺ����" Then
                txtҽԺ����.Text = Nvl(!����ֵ)
            Else
                chkʵʱ�ϴ�.Value = Nvl(!����ֵ, 1)
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
