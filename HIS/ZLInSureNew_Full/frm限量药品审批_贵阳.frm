VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm����ҩƷ����_���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ҩƷ����"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10875
   Icon            =   "frm����ҩƷ����_����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   10875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fma 
      Caption         =   "������"
      Height          =   3525
      Index           =   2
      Left            =   0
      TabIndex        =   22
      Top             =   2850
      Width           =   10815
      Begin VB.CommandButton cmd������־ 
         Caption         =   "������־"
         Height          =   315
         Left            =   8280
         TabIndex        =   45
         Top             =   3150
         Width           =   1065
      End
      Begin VB.CommandButton cmdȡ������ 
         Caption         =   "ȡ������"
         Enabled         =   0   'False
         Height          =   315
         Left            =   9690
         TabIndex        =   43
         Top             =   3150
         Width           =   1065
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshBill 
         Height          =   2895
         Left            =   60
         TabIndex        =   19
         Top             =   210
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   5106
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   -2147483631
         GridColorFixed  =   8421504
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame fma 
      Caption         =   "������"
      Height          =   1815
      Index           =   1
      Left            =   0
      TabIndex        =   21
      Top             =   1020
      Width           =   10815
      Begin VB.TextBox txtҩƷ���� 
         Height          =   315
         Left            =   840
         TabIndex        =   9
         Top             =   180
         Width           =   3495
      End
      Begin VB.TextBox txt��� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   6450
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   1020
         Width           =   1515
      End
      Begin VB.TextBox txt���� 
         Height          =   315
         Left            =   4560
         TabIndex        =   14
         Top             =   1020
         Width           =   975
      End
      Begin VB.TextBox txt�ۼ� 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   2190
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1020
         Width           =   1305
      End
      Begin VB.TextBox txt���� 
         Enabled         =   0   'False
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   7125
      End
      Begin VB.TextBox txt��� 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   180
         Width           =   2685
      End
      Begin VB.TextBox txt��λ 
         Enabled         =   0   'False
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1020
         Width           =   915
      End
      Begin VB.TextBox txt��ע 
         Height          =   285
         Left            =   840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   1448
         Width           =   7125
      End
      Begin VB.CommandButton cmdҩƷ 
         Caption         =   "��"
         Height          =   315
         Left            =   4350
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   180
         Width           =   315
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "ͨ������"
         Height          =   315
         Left            =   8040
         TabIndex        =   18
         Top             =   1410
         Width           =   1065
      End
      Begin MSComCtl2.DTPicker dtpЧ�� 
         Height          =   300
         Left            =   8700
         TabIndex        =   16
         Top             =   1020
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   105971715
         CurrentDate     =   36279
         MinDate         =   2
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ч��"
         Enabled         =   0   'False
         Height          =   180
         Index           =   18
         Left            =   8130
         TabIndex        =   44
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ����"
         Height          =   180
         Index           =   17
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         Enabled         =   0   'False
         Height          =   180
         Index           =   16
         Left            =   4890
         TabIndex        =   41
         Top             =   240
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Enabled         =   0   'False
         Height          =   180
         Index           =   15
         Left            =   480
         TabIndex        =   40
         Top             =   660
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "׼����"
         Height          =   180
         Index           =   14
         Left            =   3840
         TabIndex        =   39
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ۼ۵�λ"
         Enabled         =   0   'False
         Height          =   180
         Index           =   13
         Left            =   120
         TabIndex        =   38
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ۼ�"
         Enabled         =   0   'False
         Height          =   180
         Index           =   12
         Left            =   1800
         TabIndex        =   37
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ۼ۽��"
         Enabled         =   0   'False
         Height          =   180
         Index           =   11
         Left            =   5760
         TabIndex        =   36
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ע"
         Height          =   180
         Index           =   10
         Left            =   450
         TabIndex        =   35
         Top             =   1500
         Width           =   360
      End
   End
   Begin VB.CommandButton cmd�˳� 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&X)"
      Height          =   345
      Left            =   9720
      TabIndex        =   23
      Top             =   6510
      Width           =   1065
   End
   Begin VB.Frame fma 
      Caption         =   "������Ϣ"
      Height          =   945
      Index           =   0
      Left            =   0
      TabIndex        =   20
      Top             =   60
      Width           =   10815
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   9000
         TabIndex        =   8
         Top             =   540
         Width           =   1635
      End
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   5940
         TabIndex        =   7
         Top             =   540
         Width           =   1635
      End
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   3480
         TabIndex        =   6
         Top             =   540
         Width           =   1635
      End
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   1770
         TabIndex        =   5
         Top             =   540
         Width           =   705
      End
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   840
         TabIndex        =   4
         Top             =   540
         Width           =   465
      End
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   9000
         TabIndex        =   3
         Top             =   203
         Width           =   1635
      End
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   5940
         TabIndex        =   2
         Top             =   203
         Width           =   1635
      End
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3480
         TabIndex        =   1
         Top             =   203
         Width           =   1635
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   0
         Top             =   203
         Width           =   1635
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ϴξ���ʱ��"
         Enabled         =   0   'False
         Height          =   180
         Index           =   8
         Left            =   7920
         TabIndex        =   32
         Top             =   585
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ID"
         Enabled         =   0   'False
         Height          =   180
         Index           =   7
         Left            =   2910
         TabIndex        =   31
         Top             =   255
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Enabled         =   0   'False
         Height          =   180
         Index           =   6
         Left            =   450
         TabIndex        =   30
         Top             =   592
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Enabled         =   0   'False
         Height          =   180
         Index           =   5
         Left            =   1410
         TabIndex        =   29
         Top             =   585
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ����"
         Enabled         =   0   'False
         Height          =   180
         Index           =   4
         Left            =   5400
         TabIndex        =   28
         Top             =   592
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���￨��"
         Enabled         =   0   'False
         Height          =   180
         Index           =   3
         Left            =   2730
         TabIndex        =   27
         Top             =   585
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��"
         Enabled         =   0   'False
         Height          =   180
         Index           =   2
         Left            =   8460
         TabIndex        =   26
         Top             =   255
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Enabled         =   0   'False
         Height          =   180
         Index           =   1
         Left            =   5400
         TabIndex        =   25
         Top             =   255
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   0
         Left            =   450
         TabIndex        =   24
         Top             =   255
         Width           =   360
      End
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "###"
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   9
      Left            =   60
      TabIndex        =   33
      Top             =   6480
      Width           =   270
   End
End
Attribute VB_Name = "frm����ҩƷ����_����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint���� As Integer
Public Sub ShowME(ByVal intinsure As Integer)
    mint���� = intinsure
    dtpЧ��.Value = zlDatabase.Currentdate + 7
    lbl(9).Caption = "��ʾ��1��������������ͨ������-����ID:+סԺ��;*�����;/ҽ���Ż�ֱ��¼�������ķ�ʽ��ȡ������Ϣ" & vbNewLine & _
                    "       2�������漰������������׼������ʱ����ָ��������������������óɰ�������ʽ����ʱ��ϵͳ���Զ����㣡"
    Me.Show 1
End Sub

Private Sub cmd������־_Click()
     Call zl9Report.ReportOpen(gcnOracle, 0, "SYB_LOG1", Me)
End Sub

Private Sub cmdȡ������_Click()
    On Error GoTo errHand
    With mshBill
        If .Row = 0 Or .Rows = 1 Then ShowMsgbox "��ѡ��Ҫȡ�������ļ�¼��": Exit Sub
        If MsgBox("Ҫȡ����[" & .TextMatrix(.Row, 2) & "]" & .TextMatrix(.Row, 3) & "��������������Ϣ��", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
        gstrSQL = "Zl_��ҩ��������_����_Delete(" & mint���� & ",'" & .TextMatrix(.Row, 0) & "','" & .TextMatrix(.Row, 1) & "','" & UserInfo.���� & "',Sysdate)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Call Get������
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmd����_Click()
    On Error GoTo errHand
    If txtҩƷ����.Tag = "" Then ShowMsgbox "��ȷ��ҩƷ��Ϣ��": txtҩƷ����.SetFocus: Exit Sub
    If Val(txt����.Text) <= 0 Then ShowMsgbox "ҩƷ��׼�������������0��": txt����.SetFocus: Exit Sub
    If Val(txt(0).Tag) = 0 Then ShowMsgbox "����ȷ�����������Ϣ��": txt(0).SetFocus: Exit Sub
    If Format(dtpЧ��.Value, "yyyy-mm-dd") < Format(zlDatabase.Currentdate, "yyyy-mm-dd") Then
        ShowMsgbox "ҩƷ����ʹ����Ч�ڲ���С�ڽ��죡": dtpЧ��.SetFocus: Exit Sub
    End If
    gstrSQL = "Zl_��ҩ��������_����_Insert(" & mint���� & ",'" & txt(0).Tag & "','" & txtҩƷ����.Tag & "','" & txt����.Text & "'," & _
              "To_Date('" & Format(dtpЧ��.Value, "yyyy-mm-dd") & "','yyyy-mm-dd'),'" & txt��ע.Text & "','" & UserInfo.���� & "',Sysdate)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Call Get������
    ShowMsgbox "ҩƷ����������Ϣ���óɹ���"
    txtҩƷ����.Text = "": txt��ע.Text = "": txt����.Text = "": txt���.Text = ""
    txt���.Text = ""
    txt���.Text = ""
    txt��λ.Text = "": txt�ۼ�.Text = ""
    txt����.Text = 1: txtҩƷ����.Tag = "": txtҩƷ����.SetFocus
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmd�˳�_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mshBill.Cols > 3 Then Call SaveFlexState(mshBill, Me.Caption)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo errHand
    Dim strCode As String, rsTemp As New ADODB.Recordset
    Select Case Index
        Case 0
            If KeyAscii <> vbKeyReturn Then Exit Sub
                strCode = Trim(txt(Index).Text)
                If Left(strCode, 1) = "-" And IsNumeric(Mid(strCode, 2)) Then  '����ID
                   gstrSQL = "Select A.����id As ID, A.����id, A.�����, A.סԺ��, A.���￨��, A.����, A.�Ա�, A.����, B.ҽ����," & _
                           "      B.����ʱ�� As ������ʱ�� " & _
                           "From ������Ϣ A, �����ʻ� B " & _
                           "Where A.���� = B.���� And A.����ID = B.����ID And B.����=" & mint���� & " And A.����ID='" & Mid(strCode, 2) & "'"
                ElseIf Left(strCode, 1) = "+" And IsNumeric(Mid(strCode, 2)) Then  'סԺ��
                    gstrSQL = "Select A.����id As ID, A.����id, A.�����, A.סԺ��, A.���￨��, A.����, A.�Ա�, A.����, B.ҽ����," & _
                           "      B.����ʱ�� As ������ʱ�� " & _
                           "From ������Ϣ A, �����ʻ� B " & _
                           "Where A.���� = B.���� And A.����ID = B.����ID And B.����=" & mint���� & " And A.סԺ��='" & Mid(strCode, 2) & "'"
                ElseIf (Left(strCode, 1) = "*") And IsNumeric(Mid(strCode, 2)) Then      '�����
                    gstrSQL = "Select A.����id As ID, A.����id, A.�����, A.סԺ��, A.���￨��, A.����, A.�Ա�, A.����, B.ҽ����," & _
                           "      B.����ʱ�� As ������ʱ�� " & _
                           "From ������Ϣ A, �����ʻ� B " & _
                           "Where A.���� = B.���� And A.����ID = B.����ID And B.����=" & mint���� & " And A.�����='" & Mid(strCode, 2) & "'"
                ElseIf (Left(strCode, 1) = "/") And IsNumeric(Mid(strCode, 2)) Then      'ҽ����
                    gstrSQL = "Select A.����id As ID, A.����id, A.�����, A.סԺ��, A.���￨��, A.����, A.�Ա�, A.����, B.ҽ����," & _
                           "      B.����ʱ�� As ������ʱ�� " & _
                           "From ������Ϣ A, �����ʻ� B " & _
                           "Where A.���� = B.���� And A.����ID = B.����ID And B.����=" & mint���� & " And B.ҽ����='" & Mid(strCode, 2) & "'"
                Else
                    'ȫ��������
                    gstrSQL = "Select A.����id As ID, A.����id, A.�����, A.סԺ��, A.���￨��, A.����, A.�Ա�, A.����, B.ҽ����," & _
                           "      B.����ʱ�� As ������ʱ�� " & _
                           "From ������Ϣ A, �����ʻ� B " & _
                           "Where A.���� = B.���� And A.����ID = B.����ID And B.����=" & mint���� & " And A.����='" & strCode & "'"
                End If
                Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 0, "ҽ������", , txt(Index).Text)
                If Not rsTemp Is Nothing Then
                    If rsTemp.State = 1 Then
                        txt(0).Tag = Nvl(rsTemp!ID): txt(0).Text = Nvl(rsTemp!����)
                        txt(1).Text = Nvl(rsTemp!����ID): txt(2).Text = Nvl(rsTemp!�����)
                        txt(3).Text = Nvl(rsTemp!סԺ��): txt(4).Text = Nvl(rsTemp!�Ա�)
                        txt(5).Text = Nvl(rsTemp!����): txt(6).Text = Nvl(rsTemp!���￨��)
                        txt(7).Text = Nvl(rsTemp!ҽ����): txt(8).Text = Nvl(rsTemp!������ʱ��)
                        Call zlCommFun.PressKey(vbKeyTab)
                    Else
                       txt(0).Tag = "": txt(0).SetFocus
                    End If
                Else
                    txt(0).Tag = "": txt(0).SetFocus
                End If
                Call Get������
    End Select
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Get������()
   '�����е�������Ϣ
   On Error GoTo errHand
   Dim rsTemp As New ADODB.Recordset
    gstrSQL = " Select A.����id, A.ҩƷid, B.����, B.����, B.���, B.���㵥λ As �ۼ۵�λ, C.�ּ� As �ۼ�, A.׼����, A.��Ч��," & _
              "       A.��ע  From ��ҩ��������_���� A, �շ�ϸĿ B, �շѼ�Ŀ C " & _
              "Where A.ҩƷid = B.ID And B.ID = C.�շ�ϸĿid And (C.��ֹ���� Is Null Or C.��ֹ���� = To_Date('3000-01-01', 'yyyy-mm-dd')) " & _
              "      And A.����=[1] And A.����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint����, CLng(Val(txt(0).Tag)))
    Set mshBill.DataSource = rsTemp
    Call CenterTableCaption(mshBill)
    mshBill.ColWidth(0) = 0: mshBill.ColWidth(1) = 0
    cmdȡ������.Enabled = mshBill.Rows > 1 And mshBill.Row <> 0
    Call RestoreFlexState(mshBill, Me.Caption)
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub
Private Sub cmdҩƷ_Click()
    On Error GoTo errHand
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select  ID, Null As �ϼ�id, 0 As ĩ��, ���� As ����,  ����, ����, Null As ���, Null As ����," & _
              "      Null As �ۼ۵�λ, -null As �ۼ� From ҩƷ��;���� Union All " & _
              "Select B.ҩƷid As ID, A.��;����id As �ϼ�id, 1 As ĩ��, D.���� As ����, B.����, B.����, B.���, B.����, B.�ۼ۵�λ," & _
              "      C.�ּ� As �ۼ� From ҩƷ��Ϣ A, ҩƷĿ¼ B, �շѼ�Ŀ C, ҩƷ��;���� D " & _
              "Where A.��;����id = D.ID And A.ҩ��id = B.ҩ��id And B.ҩƷid = C.�շ�ϸĿid And " & _
              "     (B.����ʱ�� Is Null Or B.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And " & _
              "     (C.��ֹ���� Is Null Or C.��ֹ���� = To_Date('3000-01-01', 'yyyy-mm-dd')) "
    Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 2, "ҩƷĿ¼", , txtҩƷ����.Text)
    If Not rsTemp Is Nothing Then
        If rsTemp.State = 1 Then
            txtҩƷ����.Text = "[" & Nvl(rsTemp!����) & "]" & Nvl(rsTemp!����)
            txt���.Text = Nvl(rsTemp!���): txt����.Text = Nvl(rsTemp!����)
            txt��λ.Text = Nvl(rsTemp!�ۼ۵�λ): txt�ۼ�.Text = Format(Nvl(rsTemp!�ۼ�), "0.00000")
            txt����.Text = 1: txtҩƷ����.Tag = Nvl(rsTemp!ID)
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            txt����.Text = 0: txtҩƷ����.Tag = ""
        End If
    Else
        txt����.Text = 0: txtҩƷ����.Tag = ""
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Private Sub txt����_Change()
    txt���.Text = Format(Val(txt�ۼ�.Text) * Val(txt����.Text), "0.00000")
End Sub

Private Sub txt����_GotFocus()
     zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txt����, KeyAscii, m���ʽ)
End Sub

Private Sub txtҩƷ����_GotFocus()
    zlControl.TxtSelAll txtҩƷ����
End Sub

Private Sub txtҩƷ����_KeyPress(KeyAscii As Integer)
     On Error GoTo errHand
    If KeyAscii <> vbKeyReturn Then Exit Sub
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select Distinct B.ҩƷid As ID,Null As �ϼ�ID,D.���� As ����, B.����, B.����, B.���, B.����," & _
              "       B.�ۼ۵�λ, C.�ּ� As �ۼ� From ҩƷ��Ϣ A, ҩƷĿ¼ B, �շѼ�Ŀ C, ҩƷ��;���� D, �շѱ��� E " & _
              "Where A.��;����id = D.ID And A.ҩ��id = B.ҩ��id And B.ҩƷid = C.�շ�ϸĿid And " & _
              "     (B.����ʱ�� Is Null Or B.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And " & _
              "     (C.��ֹ���� Is Null Or C.��ֹ���� = To_Date('3000-01-01', 'yyyy-mm-dd')) And B.ҩƷid = E.�շ�ϸĿid " & _
              "     And (Upper(B.����) Like '%" & UCase(txtҩƷ����.Text) & "%' Or  Upper(B.����) Like '%" & UCase(txtҩƷ����.Text) & "%' " & _
              "     Or Upper(E.����) Like '%" & UCase(txtҩƷ����.Text) & "%')"
    Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 0, "ҩƷĿ¼", , txtҩƷ����.Text)
    If Not rsTemp Is Nothing Then
        If rsTemp.State = 1 Then
            txtҩƷ����.Text = "[" & Nvl(rsTemp!����) & "]" & Nvl(rsTemp!����)
            txt���.Text = Nvl(rsTemp!���): txt����.Text = Nvl(rsTemp!����)
            txt��λ.Text = Nvl(rsTemp!�ۼ۵�λ): txt�ۼ�.Text = Format(Nvl(rsTemp!�ۼ�), "0.00000")
            txt����.Text = 1: txtҩƷ����.Tag = Nvl(rsTemp!ID)
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            txt����.Text = 0: txtҩƷ����.Tag = ""
            txtҩƷ����.SetFocus
        End If
    Else
        txt����.Text = 0: txtҩƷ����.Tag = ""
        txtҩƷ����.SetFocus
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub


