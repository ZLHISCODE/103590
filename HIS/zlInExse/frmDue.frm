VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ӧ�տ�Ǽ�"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDue.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9255
   StartUpPosition =   1  '����������
   Begin VB.Frame fraLine 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   120
      TabIndex        =   14
      Top             =   4800
      Width           =   8925
   End
   Begin VB.TextBox txtPay 
      Appearance      =   0  'Flat
      BackColor       =   &H00E7CFBA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   7440
      MaxLength       =   10
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2880
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.TextBox txtLack 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   120
      Width           =   1995
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   120
      Width           =   1995
   End
   Begin VB.TextBox txt�ɿ� 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   9
      Text            =   "0.00"
      Top             =   4320
      Width           =   1995
   End
   Begin VB.TextBox txt�Ҳ� 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   6960
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   4320
      Width           =   1995
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5760
      TabIndex        =   10
      Top             =   5160
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7560
      TabIndex        =   11
      Top             =   5160
      Width           =   1500
   End
   Begin VB.TextBox txtPay 
      Appearance      =   0  'Flat
      BackColor       =   &H00E7CFBA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   1380
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshPay 
      Height          =   1665
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   2937
      _Version        =   393216
      Rows            =   5
      Cols            =   4
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   "^  ���㷽ʽ  |^  ������  |^     �������     |^             ��ע           "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshBalance 
      Height          =   1665
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   2937
      _Version        =   393216
      Rows            =   5
      Cols            =   7
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   "^ѡ��|^  ���ݺ�  |^  Ʊ�ݺ�  |^������ |^   ��������   |^ Ӧ�ս��  |^    ��Ӧ��  "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ӧ�պϼ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1200
   End
   Begin VB.Label lbl�ɿ� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ֽ�ɿ�(J)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   1320
   End
   Begin VB.Label lbl�Ҳ� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ֽ��Ҳ�(&B)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5520
      TabIndex        =   12
      Top             =   4440
      Width           =   1320
   End
   Begin VB.Label lblLack 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5520
      TabIndex        =   2
      Top             =   240
      Width           =   960
   End
End
Attribute VB_Name = "frmDue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Public mlng����ID As Long

Private mlngDefalt As Long 'ȱʡ���㷽ʽ��
Private mlngCash As Long    '�ֽ���㷽ʽ
Private mcurTotal As Currency   'ѡ�������ĳ�Ӧ�տ�ϼ�,С�ڵ���Ӧ�տ�ϼ�
Private mcurInsure As Currency  'ҽ���ĳ�Ӧ�տ�ϼ�,С�ڵ���Ӧ�տ�ϼ�
Private mstrBalance As String   'ҽ��������㷽ʽ
Private mcurCheckInsure As Currency

Private Enum PAYCOL
    C0��ʽ = 0
    C1��� = 1
    C2���� = 2
    C3��ע = 3
End Enum
Private Enum BALANCECOL
    C0ѡ�� = 0
    C1���ݺ� = 1
    C2Ʊ�ݺ� = 2
    C3������ = 3
    C4�������� = 4
    C5Ӧ�ս�� = 5
    C6��Ӧ�� = 6
End Enum

Private Sub cmdCancel_Click()
    gblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, strNO As String, Curdate As Currency, curInsure As Currency
    Dim arrSQL() As Variant, blnTrans As Boolean, curOther As Currency
    Dim curTemp As Currency
    
    If Val(txtLack.Text) > 0 Then
        MsgBox "�������Ҫ����Ľ��ʿ�������㹻�Ľ�����!", vbInformation, gstrSysName
        mshPay.SetFocus: Exit Sub
    ElseIf Val(txtLack.Text) < 0 Then
        MsgBox "���������Ҫ����Ľ��ʿ���������!", vbInformation, gstrSysName
        mshPay.SetFocus: Exit Sub
    End If
    If mcurTotal = 0 Then   'û����ɿҲû��ѡ��Ӧ�տ�
        MsgBox "��ѡ������Ӧ�տ�ĳ���!", vbInformation, gstrSysName
        mshBalance.SetFocus: Exit Sub
    End If
    
    With mshBalance
        For i = 1 To .Rows - 1
            If .TextMatrix(i, BALANCECOL.C0ѡ��) = "��" And Val(.TextMatrix(i, BALANCECOL.C6��Ӧ��)) <> 0 Then
                curTemp = curTemp + Val(.TextMatrix(i, BALANCECOL.C5Ӧ�ս��))
            End If
        Next
    End With
    
    With mshPay
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, PAYCOL.C1���)) = 0 Then
                If (.TextMatrix(i, PAYCOL.C2����) <> "" Or .TextMatrix(i, PAYCOL.C3��ע) <> "") Then
                    If MsgBox("ע��:��" & i & "��û��������,��������" & IIf(.TextMatrix(i, PAYCOL.C2����) <> "", "�������", "��ע") & _
                        vbCrLf & "!����Ϣ���ᱣ��!ȷ��Ҫ������?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        .SetFocus: .Row = i: .Col = PAYCOL.C1���: Exit Sub
                    End If
                End If
            End If
            If .RowData(i) = 4 Then
                curInsure = curInsure + Val(.TextMatrix(i, PAYCOL.C1���))
            Else
                curOther = curOther + Val(.TextMatrix(i, PAYCOL.C1���))
            End If
        Next
        If curInsure > mcurInsure Then
            MsgBox "ע��:�����ҽ��������(>" & Format(mcurInsure, "0.00") & "),����!", vbInformation, gstrSysName
            mshPay.SetFocus: Exit Sub
        End If
        If curOther > curTemp - mcurCheckInsure And curOther > 0 Then
            MsgBox "ע��:����ķ�ҽ��������(>" & Format(curTemp - mcurCheckInsure, "0.00") & "),����!", vbInformation, gstrSysName
            mshPay.SetFocus: Exit Sub
        End If
    End With
    
    On Error GoTo errH
    arrSQL = Array()
    strNO = zlDatabase.GetNextNo(18)
    Curdate = zlDatabase.Currentdate
    
    With mshPay
        For i = 1 To .Rows - 1
            If .TextMatrix(i, PAYCOL.C0��ʽ) <> "" And Val(.TextMatrix(i, PAYCOL.C1���)) <> 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_���˽ɿ��¼_Insert(" & mlng����ID & ",'" & strNO & "','" & .TextMatrix(i, PAYCOL.C0��ʽ) & "','" & _
                    .TextMatrix(i, PAYCOL.C2����) & "'," & .TextMatrix(i, PAYCOL.C1���) & ",'" & .TextMatrix(i, PAYCOL.C3��ע) & "'," & _
                    "To_Date('" & Format(Curdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'" & UserInfo.���� & "')"
            End If
        Next
    End With
    With mshBalance
        For i = 1 To .Rows - 1
            If .TextMatrix(i, BALANCECOL.C0ѡ��) = "��" And Val(.TextMatrix(i, BALANCECOL.C6��Ӧ��)) <> 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_���˽ɿ����_Insert('" & strNO & "'," & .RowData(i) & "," & .TextMatrix(i, BALANCECOL.C6��Ӧ��) & ")"
            End If
        Next
    End With
    
    gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next
    gcnOracle.CommitTrans: blnTrans = False
        
    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1137_1", Me, "NO=" & strNO, 2)
       
    gblnOK = True
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    mshPay.SetFocus
End Sub

Private Sub Form_Load()
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim i As Long, j As Long, strBalance() As String
    Dim curInsure As Currency
    
    gblnOK = False
    mcurTotal = 0
    mlngDefalt = 0
    mlngCash = 0
    Call RestoreWinState(Me, App.ProductName)
    
    On Error GoTo errH
    'Ӧ�տ�Ҫ��Ϊ����
    strSql = "Select A.ID , A.���ݺ�, A.Ʊ�ݺ�, A.������, To_Char(A.����ʱ��,'YYYY-MM-DD') ��������, Ӧ�ս�� - Nvl(Sum(B.���), 0) Ӧ�ս��," & vbNewLine & _
            "       Ӧ�ս�� - Nvl(Sum(B.���), 0) ��Ӧ��" & vbNewLine & _
            "From (Select A.NO ���ݺ�, A.ʵ��Ʊ�� Ʊ�ݺ�, A.����Ա���� ������, A.�շ�ʱ�� ����ʱ��, A.ID, Sum(B.��Ԥ��) Ӧ�ս��" & vbNewLine & _
            "       From ���˽��ʼ�¼ A, ����Ԥ����¼ B, ���㷽ʽ C" & vbNewLine & _
            "       Where A.����id = [1] And A.��¼״̬ =1 And A.ID = B.����id And B.���㷽ʽ = C.���� And C.Ӧ�տ� = 1" & vbNewLine & _
            "       Group By A.NO, A.ʵ��Ʊ��, A.����Ա����, A.�շ�ʱ��, A.ID) A, ���˽ɿ���� B" & vbNewLine & _
            "Where A.ID = B.����id(+) " & vbNewLine & _
            "Group By A.ID , A.���ݺ�, A.Ʊ�ݺ�, A.������, A.����ʱ��, Ӧ�ս��" & vbNewLine & _
            "Having (Ӧ�ս�� - Nvl(Sum(B.���), 0))>0 " & vbNewLine & _
            "Order By ��������,���ݺ�"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID)
    If rsTmp.RecordCount = 0 Then
        MsgBox "��ǰ����û��δ�����Ӧ�տ�!", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    With mshBalance
        .Rows = rsTmp.RecordCount + 1
        For i = 1 To rsTmp.RecordCount
            .RowData(i) = rsTmp!ID
            .TextMatrix(i, BALANCECOL.C0ѡ��) = "��"
            .TextMatrix(i, BALANCECOL.C1���ݺ�) = rsTmp!���ݺ�
            .TextMatrix(i, BALANCECOL.C2Ʊ�ݺ�) = "" & rsTmp!Ʊ�ݺ�
            .TextMatrix(i, BALANCECOL.C3������) = rsTmp!������
            .TextMatrix(i, BALANCECOL.C4��������) = rsTmp!��������
            .TextMatrix(i, BALANCECOL.C5Ӧ�ս��) = rsTmp!Ӧ�ս��
            .TextMatrix(i, BALANCECOL.C6��Ӧ��) = Val("" & rsTmp!��Ӧ��)
            rsTmp.MoveNext
        Next
    End With
    Call SetTotal
    
    strSql = "Select A.����,A.ȱʡ��־ ȱʡ,A.����,A.����" & vbNewLine & _
            "From ���㷽ʽ A, ���㷽ʽӦ�� B" & vbNewLine & _
            "Where A.���� = B.���㷽ʽ And B.Ӧ�ó��� = '����' And A.���� Not In (3, 4, 9) " & vbNewLine & _
            " Union " & _
            "Select A.����,A.ȱʡ��־ ȱʡ,A.����,A.����" & vbNewLine & _
            "From ���㷽ʽ A, ���㷽ʽӦ�� B" & vbNewLine & _
            "Where A.���� = B.���㷽ʽ And B.Ӧ�ó��� = '����' And A.���� = 4 And Nvl(A.Ӧ�տ�, 0) = 1" & vbNewLine & _
            "Order By ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    If rsTmp.RecordCount = 0 Then
        MsgBox "û�����ý��ʳ��ϵĽ��㷽ʽ,���ܽ����տ�!", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    If mstrBalance <> "" Then
        strBalance = Split(mstrBalance, "|")
    End If
    
    curInsure = mcurInsure
    
    With mshPay
        .Rows = rsTmp.RecordCount + 1
        For i = 1 To rsTmp.RecordCount
            .TextMatrix(i, PAYCOL.C0��ʽ) = rsTmp!����
            .RowData(i) = rsTmp!����
            
            If Val("" & rsTmp!ȱʡ) = 1 Then mlngDefalt = i
            If rsTmp!���� = 1 Then
                mlngCash = i
                If mlngDefalt = 0 Then mlngDefalt = i
            End If
            If rsTmp!���� = 4 Then
                If mstrBalance <> "" Then
                    For j = 0 To UBound(strBalance)
                        If .TextMatrix(i, PAYCOL.C0��ʽ) = Split(strBalance(j), ",")(0) Then
                            If curInsure > Val(Split(strBalance(j), ",")(1)) Then
                                .TextMatrix(i, PAYCOL.C1���) = Val(Split(strBalance(j), ",")(1))
                                curInsure = curInsure - Val(Split(strBalance(j), ",")(1))
                            Else
                                If curInsure <> 0 Then .TextMatrix(i, PAYCOL.C1���) = curInsure
                                curInsure = 0
                            End If
                        End If
                    Next j
                End If
            End If
            rsTmp.MoveNext
        Next
        If mlngDefalt = 0 Then mlngDefalt = 1
        If mcurTotal - mcurInsure <> 0 Then .TextMatrix(mlngDefalt, PAYCOL.C1���) = mcurTotal - mcurInsure
    End With
    
    txtLack.Text = "0.00"
    If mlngCash = 0 Then txt�ɿ�.Enabled = False
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlng����ID = 0
    Call SaveWinState(Me, App.ProductName)
End Sub

'---------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------

Private Sub mshPay_DblClick()
    If Not txtPay(0).Visible And mshPay.Row > 0 And mshPay.Col > PAYCOL.C0��ʽ Then
        Call SetTxtPay(0)
        txtPay(0).Text = mshPay.TextMatrix(mshPay.Row, mshPay.Col)
        txtPay(0).SelStart = 0: txtPay(0).SelLength = Len(txtPay(0).Text)
    End If
End Sub

Private Sub mshPay_KeyDown(KeyCode As Integer, Shift As Integer)
    If mshPay.Row <= 0 Then Exit Sub
    If KeyCode = 13 Then
        Call LocateMshpay
    ElseIf KeyCode = vbKeyDelete Then
        If mshPay.Col > PAYCOL.C0��ʽ Then
            mshPay.TextMatrix(mshPay.Row, mshPay.Col) = ""
            If mshPay.Col = PAYCOL.C1��� Then Call SetLack
        End If
    End If
End Sub

Private Sub mshPay_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        If mshPay.Row <= 0 Then Exit Sub
        If Not txtPay(0).Visible And mshPay.Col > PAYCOL.C0��ʽ Then
            If mshPay.Col = PAYCOL.C1��� Then
                'ֻ��������
                If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
            ElseIf mshPay.Col = PAYCOL.C2���� Then
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                If ZLCommFun.IsCharChinese(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
            End If
            
            Call SetTxtPay(0)
            txtPay(0).Text = Chr(KeyAscii)
            txtPay(0).SelStart = 1
        End If
    End If
End Sub

Private Sub mshPay_LeaveCell()
    txtPay(0).Visible = False
End Sub

Private Sub mshPay_Scroll()
    txtPay(0).Visible = False
End Sub

Private Sub mshPay_EnterCell()
    If mshPay.Col = PAYCOL.C3��ע Then
        txtPay(0).IMEMode = 1
        Call OpenIme(gstrIme)
    Else
        Call OpenIme
        txtPay(0).IMEMode = 3
    End If
End Sub

'--------------------------------------------------------------------------------------------

Private Sub mshBalance_DblClick()
    If mshBalance.Row <= 0 Then Exit Sub
    If Not txtPay(1).Visible And mshBalance.Col = BALANCECOL.C6��Ӧ�� Then
        Call SetTxtPay(1)
        txtPay(1).Text = mshBalance.TextMatrix(mshBalance.Row, mshBalance.Col)
        txtPay(1).SelStart = 0: txtPay(1).SelLength = Len(txtPay(1).Text)
    ElseIf mshBalance.Col = BALANCECOL.C0ѡ�� Then
        Call SetBalanceSelect
    End If
End Sub

Private Sub mshBalance_KeyDown(KeyCode As Integer, Shift As Integer)
    If mshBalance.Row <= 0 Then Exit Sub
    If KeyCode = 13 Then
        Call LocateMshBalance
    ElseIf KeyCode = vbKeyDelete Then
        If mshBalance.Col = BALANCECOL.C6��Ӧ�� Then
            mshBalance.TextMatrix(mshBalance.Row, mshBalance.Col) = ""
            Call SetTotal
            Call SetLack
        End If
    ElseIf KeyCode = vbKeySpace And mshBalance.Col = BALANCECOL.C0ѡ�� Then
        Call SetBalanceSelect
    End If
End Sub

Private Sub SetBalanceSelect()
    If mshBalance.TextMatrix(mshBalance.Row, mshBalance.Col) = "" Then
        mshBalance.TextMatrix(mshBalance.Row, mshBalance.Col) = "��"
    Else
        mshBalance.TextMatrix(mshBalance.Row, mshBalance.Col) = ""
    End If
    If mshBalance.TextMatrix(mshBalance.Row, BALANCECOL.C6��Ӧ��) = "" Then
        mshBalance.TextMatrix(mshBalance.Row, BALANCECOL.C6��Ӧ��) = mshBalance.TextMatrix(mshBalance.Row, BALANCECOL.C5Ӧ�ս��)
    End If
    
    Call AutoEquate
End Sub

Private Sub mshBalance_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        If mshBalance.Row <= 0 Then Exit Sub
        If Not txtPay(1).Visible And mshBalance.Col = BALANCECOL.C6��Ӧ�� Then
            'ֻ��������,���Ҳ��ܴ���Ӧ�տ��
            If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
            Call SetTxtPay(1)
            txtPay(1).Text = Chr(KeyAscii)
            txtPay(1).SelStart = 1
        End If
    End If
End Sub

Private Sub mshBalance_LeaveCell()
    txtPay(1).Visible = False
End Sub

Private Sub mshBalance_Scroll()
    txtPay(1).Visible = False
End Sub

Private Sub mshBalance_EnterCell()
   Call OpenIme
End Sub



Private Sub LocateMshBalance()
    With mshBalance
        If .Row < .Rows - 1 Then '��ĩ�е����һ�л���
            .Row = .Row + 1
            .Col = BALANCECOL.C6��Ӧ��
            If .Row - (.Height \ .RowHeight(0) - 2) > 1 Then
                .TopRow = .Row - (.Height \ .RowHeight(1) - 2)
            End If
            Call mshBalance_EnterCell
        Else 'ĩ�е����һ��
            Call ZLCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub

Private Sub LocateMshpay()
    With mshPay
        '��ĩ�е����һ�л���,��ĩ�н��Ϊ�㻻��
        If .Row < .Rows - 1 And (.Col = .Cols - 1 Or .TextMatrix(.Row, PAYCOL.C1���) = "" And .Col <> PAYCOL.C0��ʽ) Then
            .Row = .Row + 1
            .Col = PAYCOL.C1���
            If .Row - (.Height \ .RowHeight(0) - 2) > 1 Then
                .TopRow = .Row - (.Height \ .RowHeight(1) - 2)
            End If
            mshPay.SetFocus: Call mshPay_EnterCell
        ElseIf .Row = .Rows - 1 And (.Col = .Cols - 1 Or .TextMatrix(.Row, PAYCOL.C1���) = "" And .Col <> PAYCOL.C0��ʽ) Then
        'ĩ�е����һ��,ĩ�н��Ϊ��,Tab
            Call ZLCommFun.PressKey(vbKeyTab)
        Else
             If .RowData(.Row) = 1 And .Col = PAYCOL.C1��� Then '�ֽ�������������
                .Col = .Col + 2
            Else
                .Col = .Col + 1
            End If
            mshPay.SetFocus: Call mshPay_EnterCell
        End If
    End With
End Sub

Private Sub SetTxtPay(Index As Integer)
    Dim mshTmp As MSHFlexGrid
    Set mshTmp = IIf(Index = 0, mshPay, mshBalance)
    
    With txtPay(Index)
        If Index = 0 Then
            .MaxLength = Val("" & Choose(mshPay.Col, 10, 30, 25))   'ժҪ�50λ,�����������ֻ��25������
        Else
            .MaxLength = 10
        End If
        .Left = mshTmp.Left + mshTmp.CellLeft + 15
        .Top = mshTmp.Top + mshTmp.CellTop + (mshTmp.CellHeight - txtPay(Index).Height) / 2 - 15
        .Width = mshTmp.CellWidth - 60
        .ForeColor = mshTmp.CellForeColor
        .BackColor = mshTmp.CellBackColor
        If Index = 0 Then
            .Alignment = IIf(mshTmp.Col = PAYCOL.C1���, 1, 0)
        Else
            .Alignment = IIf(mshTmp.Col = BALANCECOL.C6��Ӧ��, 1, 0)
        End If
        .ZOrder: .Visible = True
        .SetFocus
    End With
End Sub

Private Sub txtPay_GotFocus(Index As Integer)
    If Index = 0 Then
        If mshPay.Col = PAYCOL.C3��ע Then
            txtPay(Index).IMEMode = 1
            Call OpenIme(gstrIme)
        Else
            txtPay(Index).IMEMode = 3
        End If
    End If
End Sub

Private Sub txtPay_LostFocus(Index As Integer)
    txtPay(Index).Visible = False
    If Index = 0 Then
        txtPay(Index).IMEMode = 3: Call OpenIme
    End If
End Sub

Private Sub txtPay_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtPay(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPay(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPay_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Call SetWindowLong(txtPay(Index).hWnd, GWL_WNDPROC, glngTXTProc)
End Sub

Private Sub txtPay_Validate(Index As Integer, Cancel As Boolean)
    txtPay(Index).Visible = False
End Sub

Private Sub txtPay_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim i As Long, strTmp As String
    
    If KeyAscii <> 13 Then
        If IIf(Index = 0, mshPay.Col = PAYCOL.C1���, mshBalance.Col = BALANCECOL.C6��Ӧ��) Then
            If InStr(txtPay(Index).Text, ".") > 0 And Chr(KeyAscii) = "." Then KeyAscii = 0: Exit Sub
            'ֻ��������
            If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        ElseIf Index = 0 And mshPay.Col = PAYCOL.C2���� Then '������������ַ�����
            If InStr("'|,", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Beep: Exit Sub
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        ElseIf Index = 0 And mshPay.Col = PAYCOL.C3��ע Then    '��ע
            If InStr("'|", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Beep: Exit Sub
        End If
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        strTmp = txtPay(Index).Text
        If IIf(Index = 0, mshPay.Col = PAYCOL.C1���, mshBalance.Col = BALANCECOL.C6��Ӧ��) Then
            If Not IsNumeric(strTmp) And strTmp <> "" Then
                MsgBox "��������ȷ����ֵ��", vbInformation, gstrSysName
                Call zlControl.TxtSelAll(txtPay(Index)): Exit Sub
            End If
            '��Ӧ�ղ��ܴ���Ӧ�տ��
            If Index = 1 Then
                If Val(strTmp) > Val(mshBalance.TextMatrix(mshBalance.Row, BALANCECOL.C5Ӧ�ս��)) Then
                    MsgBox "��Ӧ�ս��ܴ���Ӧ�ս��!", vbInformation, gstrSysName
                    txtPay(Index).Text = mshBalance.TextMatrix(mshBalance.Row, BALANCECOL.C5Ӧ�ս��)
                    Call zlControl.TxtSelAll(txtPay(Index)): Exit Sub
                End If
            End If
            
            strTmp = Format(strTmp, "0.00")    '���ý��зֱҴ���
            If Index = 0 Then
                mshPay.TextMatrix(mshPay.Row, mshPay.Col) = IIf(Val(strTmp) = 0, "", strTmp)
                If mshPay.Row <> mlngDefalt Then
                    Call AutoEquate(False)
                Else
                    Call SetLack
                End If
                
            Else
                mshBalance.TextMatrix(mshBalance.Row, mshBalance.Col) = IIf(Val(strTmp) = 0, "", strTmp)
                Call AutoEquate(True)
            End If
        Else
            mshPay.TextMatrix(mshPay.Row, mshPay.Col) = IIf(Index = 0 And mshPay.Col = PAYCOL.C2����, UCase(strTmp), strTmp)
        End If
        txtPay(Index).Visible = False
        
        '�����뻻��
        If Index = 0 Then
            Call LocateMshpay
        Else
            Call LocateMshBalance
        End If
    End If
End Sub

Private Sub AutoEquate(Optional blnRecalInsure As Boolean = False)
    Dim curPay As Currency, i As Long, j As Long
    Dim curInsure As Currency, blnHave As Boolean
    Dim strBalance() As String
    Dim intBalance As Integer
    '���ݽ��ʵ���ѡ�񼰽������ı仯���Զ�������ȱʡ�Ľ��㷽ʽ��
    
    Call SetTotal
    
    If blnRecalInsure Then
        curInsure = mcurInsure
        If mstrBalance <> "" Then
            strBalance = Split(mstrBalance, "|")
            For i = 1 To mshPay.Rows - 1
                blnHave = False
                For j = 0 To UBound(strBalance)
                    If mshPay.TextMatrix(i, PAYCOL.C0��ʽ) = Split(strBalance(j), ",")(0) Then
                        blnHave = True
                        intBalance = j
                    End If
                Next j
                If blnHave Then
                    If curInsure > Val(Split(strBalance(intBalance), ",")(1)) Then
                        mshPay.TextMatrix(i, PAYCOL.C1���) = Format(Val(Split(strBalance(intBalance), ",")(1)), "0.00")
                        curInsure = curInsure - Val(Split(strBalance(intBalance), ",")(1))
                    Else
                        mshPay.TextMatrix(i, PAYCOL.C1���) = Format(curInsure, "0.00")
                        curInsure = 0
                    End If
                End If
            Next
        End If
    End If
    
    For i = 1 To mshPay.Rows - 1
        If mshPay.TextMatrix(i, PAYCOL.C0��ʽ) <> "" And i <> mlngDefalt Then
            curPay = curPay + Val(mshPay.TextMatrix(i, PAYCOL.C1���))
        End If
    Next
    curPay = mcurTotal - curPay
    mshPay.TextMatrix(mlngDefalt, PAYCOL.C1���) = IIf(curPay = 0, "", Format(curPay, "0.00"))
        
    Call SetLack
End Sub


Private Sub SetTotal()
    Dim i As Long, rsTmped As ADODB.Recordset
    Dim strBalanceIDs As String, blnHave As Boolean
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    mcurTotal = 0
    mcurInsure = 0
    mstrBalance = ""
    For i = 1 To mshBalance.Rows - 1
        If mshBalance.TextMatrix(i, BALANCECOL.C0ѡ��) = "��" Then
            mcurTotal = mcurTotal + Val(mshBalance.TextMatrix(i, BALANCECOL.C6��Ӧ��))
            strBalanceIDs = strBalanceIDs & "," & mshBalance.RowData(i)
        End If
    Next
    If strBalanceIDs <> "" Then
        strBalanceIDs = Mid(strBalanceIDs, 2)
        strSql = "Select Sum(A.��Ԥ��) As ���,A.���㷽ʽ" & vbNewLine & _
                "From ����Ԥ����¼ A, ���㷽ʽ B" & vbNewLine & _
                "Where a.���㷽ʽ = b.���� And b.���� = 4 And Nvl(b.Ӧ�տ�, 0) = 1 And" & vbNewLine & _
                "      a.����id In (Select Column_Value From Table(f_Str2list([1]))) Group By A.���㷽ʽ"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strBalanceIDs)
        strSql = "Select Sum(a.���) As ���,a.���㷽ʽ" & vbNewLine & _
                "From ���˽ɿ��¼ A, ���˽ɿ���� B, ���㷽ʽ C" & vbNewLine & _
                "Where a.No = b.�ɿ And b.����id In (Select Column_Value From Table(f_Str2list([1]))) And a.���㷽ʽ = c.���� And c.���� = 4 And" & vbNewLine & _
                "      Nvl(c.Ӧ�տ�, 0) = 1 Group By a.���㷽ʽ"
        Set rsTmped = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strBalanceIDs)
        Do While Not rsTmp.EOF
            blnHave = False
            If rsTmped.RecordCount <> 0 Then
                rsTmped.MoveFirst
                Do While Not rsTmped.EOF
                    If Nvl(rsTmped!���㷽ʽ) = Nvl(rsTmp!���㷽ʽ) Then
                        blnHave = True
                        Exit Do
                    End If
                    rsTmped.MoveNext
                Loop
            End If
            If blnHave Then
                mcurInsure = mcurInsure + Val(Nvl(rsTmp!���)) - Val(Nvl(rsTmped!���))
                If Val(Nvl(rsTmp!���)) - Val(Nvl(rsTmped!���)) <> 0 Then mstrBalance = mstrBalance & "|" & Nvl(rsTmp!���㷽ʽ) & "," & Val(Nvl(rsTmp!���)) - Val(Nvl(rsTmped!���))
            Else
                mcurInsure = mcurInsure + Val(Nvl(rsTmp!���))
                mstrBalance = mstrBalance & "|" & Nvl(rsTmp!���㷽ʽ) & "," & Nvl(rsTmp!���)
            End If
            rsTmp.MoveNext
        Loop
    End If
    If mstrBalance <> "" Then mstrBalance = Mid(mstrBalance, 2)
'    If mcurInsure <> 0 Then
'        If mstrBalance <> "" Then mstrBalance = Mid(mstrBalance, 2)
'        strSql = "Select Sum(a.���) As ���" & vbNewLine & _
'                "From ���˽ɿ��¼ A, ���˽ɿ���� B, ���㷽ʽ C" & vbNewLine & _
'                "Where a.No = b.�ɿ And b.����id In (Select Column_Value From Table(f_Str2list([1]))) And a.���㷽ʽ = c.���� And c.���� = 4 And" & vbNewLine & _
'                "      Nvl(c.Ӧ�տ�, 0) = 1"
'        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strBalanceIDs)
'        If Not rsTmp.EOF Then
'            mcurInsure = mcurInsure - Val(Nvl(rsTmp!���))
'        End If
'    End If
    mcurCheckInsure = mcurInsure
    If mcurInsure > mcurTotal Then
        mcurInsure = mcurTotal
    End If
    txtTotal.Text = Format(mcurTotal, "0.00")
End Sub

Private Sub SetLack()
    Dim i As Long, curPay As Currency
    
    For i = 1 To mshPay.Rows - 1
        If mshPay.TextMatrix(i, PAYCOL.C0��ʽ) <> "" Then
            curPay = curPay + Val(mshPay.TextMatrix(i, PAYCOL.C1���))
        End If
    Next
    txtLack.Text = Format(mcurTotal - curPay, "0.00")
End Sub


'-------------------------------------------------------------------------------------------------------------
Private Sub txt�ɿ�_Change()
    If Val(txt�ɿ�.Text) = 0 Then txt�Ҳ�.Text = "0.00": Exit Sub
    
    txt�Ҳ�.Text = Format(Val(txt�ɿ�.Text) - Val(mshPay.TextMatrix(mlngCash, PAYCOL.C1���)), "0.00")
End Sub

Private Sub txt�ɿ�_GotFocus()
     Call zlControl.TxtSelAll(txt�ɿ�)
End Sub

Private Sub txt�ɿ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(txt�ɿ�.Text) <> 0 Then
            If Val(txt�Ҳ�.Text) >= 0 Then
                Call ZLCommFun.PressKey(vbKeyTab)
            Else
                MsgBox "�ɿ����,�벹��Ӧ�ɽ�", vbInformation, gstrSysName
                txt�ɿ�.SetFocus
                zlControl.TxtSelAll txt�ɿ�
            End If
        Else
            Call ZLCommFun.PressKey(vbKeyTab)
        End If
    Else
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        ElseIf KeyAscii = Asc(".") And InStr(txt�ɿ�.Text, ".") > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txt�ɿ�_Validate(Cancel As Boolean)
    txt�ɿ�.Text = Format(Trim(txt�ɿ�.Text), "0.00")
End Sub
