VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMedicareBalance 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ҽ�������շѽ���У��"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7365
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ ��(&C)"
      Height          =   435
      Left            =   5520
      TabIndex        =   16
      Top             =   4560
      Visible         =   0   'False
      Width           =   1395
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
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   3840
      Width           =   1755
   End
   Begin VB.TextBox txt�Ҳ� 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   5190
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   3840
      Width           =   1755
   End
   Begin VB.TextBox txtMargin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   5190
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "0.00"
      ToolTipText     =   "�����ı�ȱʡ���㷽ʽ�Ľ��ʱ�Ų���"
      Top             =   120
      Width           =   1755
   End
   Begin VB.TextBox txtTmp 
      Appearance      =   0  'Flat
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
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1560
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.TextBox txtԤ����� 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   5190
      MaxLength       =   10
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   600
      Width           =   1755
   End
   Begin VB.TextBox txtPay 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   600
      Width           =   1755
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   120
      Width           =   1755
   End
   Begin VB.Frame Frame2 
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
      Left            =   0
      TabIndex        =   9
      Top             =   4320
      Width           =   7365
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ ��(&O)"
      Height          =   435
      Left            =   3840
      TabIndex        =   8
      Top             =   4560
      Width           =   1395
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshBalance 
      Height          =   2505
      Left            =   390
      TabIndex        =   4
      Top             =   1200
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   4419
      _Version        =   393216
      Rows            =   5
      Cols            =   3
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
      FormatString    =   "^  ���㷽ʽ  |^   ������   |^      �������      "
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
      _Band(0).Cols   =   3
   End
   Begin VB.Label lbl�ɿ� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ֽ�ɿ�"
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
      Left            =   360
      TabIndex        =   15
      Top             =   3960
      Width           =   960
   End
   Begin VB.Label lbl�Ҳ� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ֽ��Ҳ�"
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
      Left            =   4080
      TabIndex        =   14
      Top             =   3930
      Width           =   960
   End
   Begin VB.Label lblMargin 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�շѲ��"
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
      Left            =   4080
      TabIndex        =   13
      Top             =   240
      Width           =   960
   End
   Begin VB.Label lblԤ����� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ԥ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4080
      TabIndex        =   12
      Top             =   720
      Width           =   960
   End
   Begin VB.Label lblPay 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ӧ�ɽ��"
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
      Left            =   360
      TabIndex        =   11
      Top             =   720
      Width           =   960
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ʵ�ս��"
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
      Left            =   360
      TabIndex        =   10
      Top             =   240
      Width           =   960
   End
End
Attribute VB_Name = "frmMedicareBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Private mbytInFun As Byte '0-����ģ�����,1-ҽ��ģ�����

Private mlng����ID As Long
Private mcurʵ�ս�� As Currency
Private mcur��Ԥ���� As Currency
Private mstr���ս��� As String
Private mstr�շѽ��� As String      '��ϼ�Ϊmcurʵ�ս��-mcurԤ�����-���ս���ϼ�+mcur�շ����
Private mcur�շ���� As String
Private mcurԤ����� As Currency
Private mintInsure As Integer       '�����ж��Ƿ�֧�ֱַҴ���
Private mcur�ɿ� As Currency

Private mblnOK  As Boolean
Private mintDefault As Integer 'ȱʡ���㷽ʽ��(Ϊ0��ʾû��)
Private mcurMediCare   As Currency  'ҽ������ϼ�,����[mstr���ս���]����
Private mblnClickOK As Boolean      '����ֻ�����ȷ���˳�
Private mblnCent As Boolean         'ҽ���Ƿ�֧�ֱַҴ���

'1-�ֽ���㷽ʽ,2-������ҽ������,3-ҽ�������ʻ�,4-ҽ������ͳ��,5-���տ���
Private Enum PayType
    �ֽ� = 1
    ��ҽ�����ֽ� = 2
    ҽ�������ʻ� = 3
    ҽ���������� = 4
    ���տ� = 5
End Enum

'ģ�������˽�л�
Private Const support�ֱҴ��� = 25  'ҽ�������Ƿ���ֱ�   ,��Ҫ��Ϊ�˱���ҽ����ҽԺ����
Private mstr���㷽ʽ As String
Private mstrDec As String
Private mBytMoney As Byte '�շѷֱҴ�����


Public Function ShowMeFromOut(ByRef frmParent As Object, ByVal lng����ID As Long) As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, i As Long, lng����ID As Long, strTmp As String
        
    On Error GoTo errH
    mlng����ID = lng����ID
    
    strSql = "" & _
    " Select Sum(Decode(Nvl(���ӱ�־, 0), 9, 0, ʵ�ս��)) As ʵ�ս��," & _
    "       Sum(Decode(Nvl(���ӱ�־, 0), 9, ʵ�ս��, 0)) As �����" & _
    " From ������ü�¼" & _
    " Where ��¼״̬ = 1 And ����id = [1] "
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���ս������", lng����ID)
    mcurʵ�ս�� = Val("" & rsTmp!ʵ�ս��)
    mcur�շ���� = Val("" & rsTmp!�����)
        
    strSql = "Select a.����ID,a.��¼����,a.���㷽ʽ,a.�������,b.���� ��������,a.��Ԥ�� " & _
             "   From ����Ԥ����¼ a,���㷽ʽ b " & _
             "   Where a.��¼״̬ = 1 And a.���㷽ʽ = B.���� And ����id =[1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���ս������", lng����ID)
    
    If rsTmp.RecordCount > 0 Then lng����ID = rsTmp!����ID
    
    mcur��Ԥ���� = 0
    rsTmp.Filter = "��¼����=1 or ��¼����=11"
    For i = 1 To rsTmp.RecordCount
        mcur��Ԥ���� = mcur��Ԥ���� + rsTmp!��Ԥ��
        rsTmp.MoveNext
    Next
        
    mstr�շѽ��� = "" '���㷽ʽ|������|�������||
    rsTmp.Filter = "��¼����=3 And ��������<>3 And ��������<>4"
    For i = 1 To rsTmp.RecordCount
        mstr�շѽ��� = mstr�շѽ��� & "||" & rsTmp!���㷽ʽ & "|" & rsTmp!��Ԥ�� & "|" & zlCommFun.Nvl(rsTmp!�������)
        rsTmp.MoveNext
    Next
    If mstr�շѽ��� <> "" Then mstr�շѽ��� = Mid(mstr�շѽ���, 3)
    
        
    rsTmp.Filter = 0
    strSql = "Select ���㷽ʽ,��� From ҽ���˶Ա� Where ����id =[1] And ���㷽ʽ<>'�ֽ�'"  'ҽ���ܿصĹ��̶̹�д����һ��"�ֽ�"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���ս������", lng����ID)
    mstr���ս��� = ""   '���㷽ʽ|������||
    For i = 1 To rsTmp.RecordCount
        mstr���ս��� = mstr���ս��� & "||" & rsTmp!���㷽ʽ & "|" & rsTmp!���
        rsTmp.MoveNext
    Next
    If mstr���ս��� <> "" Then mstr���ս��� = Mid(mstr���ս���, 3)
    
    
    
    mcurԤ����� = 0
    mintInsure = 0
    If lng����ID <> 0 Then
        Set rsTmp = GetMoneyInfo(lng����ID)
        If Not rsTmp.EOF Then mcurԤ����� = Val("" & rsTmp!Ԥ�����) - Val("" & rsTmp!�������)
        
        strSql = "Select nvl(����,0) as ���� From ������Ϣ Where ����ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���ս������", lng����ID)
        If Not rsTmp.EOF Then mintInsure = rsTmp!����
    End If
    
    '���ػ�ϵͳ����
    mstr���㷽ʽ = zlDatabase.GetPara("ȱʡ���㷽ʽ", glngSys, 1121)
            
    mstrDec = "0." & String(Val(zlDatabase.GetPara(9, glngSys, , 2)), "0")
    strTmp = zlDatabase.GetPara(14, glngSys, , 0)
    mBytMoney = Val(IIf(Len(strTmp) = 1, strTmp, Mid(strTmp, 2, 1)))
    
    mbytInFun = 1
    Me.Show 1, frmParent
    ShowMeFromOut = mblnOK
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function ShowMe(ByRef frmParent As Object, ByVal lng����ID As Long, ByVal curʵ�ս�� As Currency, _
        ByVal cur��Ԥ���� As Currency, ByVal str���ս��� As String, ByRef str�շѽ��� As String, _
        ByVal cur�շ���� As Currency, ByVal curԤ����� As Currency, ByVal intInsure As Integer, _
        ByVal strȱʡ���㷽ʽ As String, ByVal strȱʡ���λ�� As String, ByVal bytȱʡ�ֱҷ�ʽ As Byte, ByRef cur�ɿ� As Currency) As Boolean
        
    mlng����ID = lng����ID
    mintInsure = intInsure
    mcurʵ�ս�� = curʵ�ս��
    mcur��Ԥ���� = cur��Ԥ����
    mstr���ս��� = str���ս���
    mstr�շѽ��� = str�շѽ���
    mcur�շ���� = cur�շ����
    mcurԤ����� = curԤ�����
    
    mstr���㷽ʽ = strȱʡ���㷽ʽ
    mstrDec = strȱʡ���λ��
    mBytMoney = bytȱʡ�ֱҷ�ʽ
    mcur�ɿ� = cur�ɿ�
    
    mbytInFun = 0
    Me.Show 1, frmParent
    
    str�շѽ��� = mstr�շѽ���  '�������ڽɿ��ۼ�
    cur�ɿ� = mcur�ɿ�
    
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    mblnOK = False
    mblnClickOK = True: Unload Me
End Sub

Private Sub cmdOK_Click()
    '�������
    Dim i As Long
    
    If Val(txtMargin.Text) <> 0 Then
        If Val(txtMargin.Text) > 0 Then
            MsgBox "����֧������,�밴����ʾ�Ĳ��", vbExclamation, gstrSysName
            mshBalance.SetFocus: Exit Sub
        Else
            MsgBox "����֧��������,�밴����ʾ�Ĳ���˿", vbExclamation, gstrSysName
            mshBalance.SetFocus: Exit Sub
        End If
    End If
    
    '��������
    mstr�շѽ��� = ""
    For i = 1 To mshBalance.Rows - 1
        If Val(mshBalance.TextMatrix(i, 1)) <> 0 Then
            If mshBalance.RowData(i) <> PayType.ҽ�������ʻ� And mshBalance.RowData(i) <> PayType.ҽ���������� Then
                mstr�շѽ��� = mstr�շѽ��� & "||" & mshBalance.TextMatrix(i, 0) & "|" & Val(mshBalance.TextMatrix(i, 1)) & _
                    "|" & IIf(mshBalance.TextMatrix(i, 2) = "", " ", mshBalance.TextMatrix(i, 2))
            End If
        End If
    Next
    mstr�շѽ��� = Mid(mstr�շѽ���, 3)

    gstrSQL = "zl_�����շѽ���_Update(" & mlng����ID & ",'" & mstr�շѽ��� & "'," & mcur��Ԥ���� & ",'" & mstr���ս��� & "'," & mcur�շ���� & _
        IIf(Val(txt�ɿ�.Text) <> 0, "," & Val(txt�ɿ�.Text) & "," & Val(txt�Ҳ�.Text), "") & ",null,0)"
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    mblnOK = True
    mblnClickOK = True: Unload Me
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    mblnClickOK = True: Unload Me
End Sub

Private Sub Form_Activate()
    txt�ɿ�.SetFocus
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim rsӦ�ó��� As New ADODB.Recordset
    Dim strSql As String, i As Long, j As Long
    Dim arrPay As Variant, arrMediCare As Variant, blnExist As Boolean
    Dim curPay As Currency          '���������ϼ�
    Dim curBalance As Currency      'ʵ�պϼƼ�ҽ������ϼ�֮������
    Dim str���õ�ҽ�����㷽ʽ As String
    
    '������ʼ
    mblnClickOK = False
    mblnOK = False
    mintDefault = 0
    mcurMediCare = 0
    
    'ȷ����ȡ����ť
    If mbytInFun = 0 Then
        cmdOK.Left = cmdCancel.Left
        cmdCancel.Visible = False
    Else
        cmdCancel.Visible = True
    End If
    
    
    mblnCent = gclsInsure.GetCapability(support�ֱҴ���, , mintInsure)
    
    arrPay = Array()
    If mstr�շѽ��� <> "" Then                  '���㷽ʽ|������|�������||
        arrPay = Split(mstr�շѽ���, "||")
    End If
    arrMediCare = Array()                       '���㷽ʽ|������||
    If mstr���ս��� <> "" Then
        arrMediCare = Split(mstr���ս���, "||")
    End If
    
    On Error GoTo errH
    strSql = _
        " Select Distinct B.����,B.����,B.����,A.ȱʡ��־" & _
        " From ���㷽ʽӦ�� A,���㷽ʽ B" & _
        " Where ((A.Ӧ�ó���='�շ�' And B.����<>3 And B.����<>4) OR (B.����=3 OR B.����=4)) " & _
        "       And B.����=A.���㷽ʽ(+) And B.����<>5 And a.���ʽ Is Null" & _
        " Order by B.����,B.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    strSql = "Select Ӧ�ó���,���㷽ʽ From ���㷽ʽӦ�� Where Ӧ�ó���='�շ�' And ���ʽ Is Null"
    Set rsӦ�ó��� = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    mshBalance.ColAlignment(0) = 1
    mshBalance.ColAlignment(1) = 7
    mshBalance.ColAlignment(2) = 1
    mshBalance.Rows = rsTmp.RecordCount + 1
    i = 1
    Do While Not rsTmp.EOF
        mshBalance.RowData(i) = zlCommFun.Nvl(rsTmp!����, PayType.�ֽ�)               '�����ж��Ƿ�����޸Ľ��,�Լ��Ƿ����ֽ�
        mshBalance.TextMatrix(i, 0) = rsTmp!����
        
        'ҽ�����㷽ʽ�������޸�,���ò�ͬ��ɫ
        If mshBalance.RowData(i) = PayType.ҽ�������ʻ� Or mshBalance.RowData(i) = PayType.ҽ���������� Then
            
            '���ս���
            blnExist = False
            For j = 0 To UBound(arrMediCare)
                If Split(arrMediCare(j), "|")(0) = rsTmp!���� Then
                    blnExist = True
                    rsӦ�ó���.Filter = "���㷽ʽ='" & rsTmp!���� & "'"
                    If rsӦ�ó���.EOF Then
                        MsgBox "ע��:���㷽ʽ[" & rsTmp!���� & "]δ����Ӧ����[�շ�]����,�뵽[���㷽ʽ����]������!", vbInformation, gstrSysName
                    End If
                    
                    mshBalance.TextMatrix(i, 1) = Split(arrMediCare(j), "|")(1)
                    mshBalance.TextMatrix(i, 2) = ""    '�޽������
                    mcurMediCare = mcurMediCare + Val(mshBalance.TextMatrix(i, 1))
                    Exit For
                End If
            Next
            If blnExist Then
                For j = 0 To mshBalance.COLS - 1
                    mshBalance.Row = i: mshBalance.Col = j
                    mshBalance.CellBackColor = &HE7CFBA
                Next
                i = i + 1                                   'ҽ�����㲻�����,û�н��Ĳ���ʾ
            End If
            
            str���õ�ҽ�����㷽ʽ = str���õ�ҽ�����㷽ʽ & "," & rsTmp!����
            
        Else
            If rsTmp!���� = mstr���㷽ʽ Then mintDefault = i
            If zlCommFun.Nvl(rsTmp!ȱʡ��־, 0) = 1 And mintDefault = 0 Then mintDefault = i
            If zlCommFun.Nvl(rsTmp!����, 1) = 1 And mintDefault = 0 Then mintDefault = i
        
            '�շѽ���
            For j = 0 To UBound(arrPay)
                If Split(arrPay(j), "|")(0) = rsTmp!���� Then
                    mshBalance.TextMatrix(i, 1) = Split(arrPay(j), "|")(1)
                    mshBalance.TextMatrix(i, 2) = Trim(Split(arrPay(j), "|")(2))
                    Exit For
                End If
            Next
            i = i + 1                                      '��Ϊ�����,û�н���ҲҪ��ʾ
        End If
        rsTmp.MoveNext
    Loop
    
    mshBalance.Rows = i     '���һ�μ�1����ʹ���������б�����,���һ����ҽ���ҽ��Ϊ��,iû�м�1����ɾ��
    
    
    '�ȼ��ÿһ��ҽ�����㷽ʽ�Ƿ񶼴���
    If mstr���ս��� <> "" Then
        str���õ�ҽ�����㷽ʽ = str���õ�ҽ�����㷽ʽ & ","
        For j = 0 To UBound(arrMediCare)
            If InStr(str���õ�ҽ�����㷽ʽ, "," & Split(arrMediCare(j), "|")(0) & ",") <= 0 Then
                MsgBox "ҽ�����㷽ʽ[" & Split(arrMediCare(j), "|")(0) & "]δ����,���ȵ�[���㷽ʽ����]������!", vbInformation, gstrSysName
                cmdCancel.Visible = True
                cmdOK.Visible = False
            End If
        Next
    End If
    
    
    If mintDefault > 0 Then
        mshBalance.Row = mintDefault: mshBalance.Col = 0
        mshBalance.CellFontBold = True
        mshBalance.Col = 1
    Else        '���㷽ʽû��ȱʡֵ,�������ֽ�ʽ�����
        mshBalance.Row = 1: mshBalance.Col = 1
    End If
        
    txtԤ�����.Text = Format(mcur��Ԥ����, "0.00")
    txtԤ�����.Enabled = mcurԤ����� > 0      'Ӧ���ϼƴ����������ʹ��
    If txtԤ�����.Enabled Then txtԤ�����.Enabled = (mcurʵ�ս�� - mcurMediCare > 0)
    
    txtTotal.Text = Format(mcurʵ�ս��, mstrDec)
    txtTotal.ToolTipText = "Ԥ����ʱ,�����:" & Format(mcur�շ����, mstrDec)
    
    Call ShowMoney(True)
            
    curPay = 0
    For i = 1 To mshBalance.Rows - 1
        If mshBalance.RowData(i) <> PayType.ҽ�������ʻ� And mshBalance.RowData(i) <> PayType.ҽ���������� Then
            curPay = curPay + Val(mshBalance.TextMatrix(i, 1))
        End If
    Next
    txtPay.Text = Format(curPay, "0.00")
    txt�ɿ�.Text = Format(mcur�ɿ�, "0.00")
            
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtTmp_KeyPress(KeyAscii As Integer)
    Dim i As Long
    Dim curPay As Currency
    
    If KeyAscii <> 13 Then
        If mshBalance.Col = 1 Then
            If KeyAscii = vbKeyEscape Then txtTmp.Visible = False: mshBalance.SetFocus: Exit Sub
            If InStr(txtTmp.Text, ".") > 0 And Chr(KeyAscii) = "." Then KeyAscii = 0:  Exit Sub
            If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        Else '������������ַ�����
            If InStr("'|,", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Beep: Exit Sub
        End If
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0

        If mshBalance.Col = 2 Then
            '�����������������ַ�
            If InStr(txtTmp.Text, "'") > 0 Or InStr(txtTmp.Text, "|") > 0 Or InStr(txtTmp.Text, ",") > 0 Then
                Call Beep: Exit Sub
            End If
            mshBalance.TextMatrix(mshBalance.Row, mshBalance.Col) = txtTmp.Text
        ElseIf mshBalance.Col = 1 Then
            If Not IsNumeric(txtTmp.Text) And Trim(txtTmp.Text) <> "" Then
                MsgBox "��������ȷ����ֵ��", vbInformation, gstrSysName
                zlControl.TxtSelAll txtTmp: Exit Sub
            End If
            
            If Val(txtTmp.Text) <> 0 Then   '���ַ���valΪ��
                txtTmp.Text = Format(Val(txtTmp.Text), "0.00")
                If Val(mshBalance.RowData(mshBalance.Row)) = PayType.�ֽ� And mblnCent Then  '��������ֽ���������,����зֱҴ���
                    txtTmp.Text = Format(CentMoney(Val(txtTmp.Text)), "0.00")
                End If
                If Val(mshBalance.TextMatrix(mshBalance.Row, mshBalance.Col)) = Val(txtTmp.Text) Then txtTmp.Visible = False: mshBalance.SetFocus: Exit Sub
                If txtTmp.Text = "0.00" Then
                    mshBalance.TextMatrix(mshBalance.Row, mshBalance.Col) = ""
                Else
                    mshBalance.TextMatrix(mshBalance.Row, mshBalance.Col) = Format(Val(txtTmp.Text), "0.00")
                End If
            Else
                If mshBalance.TextMatrix(mshBalance.Row, mshBalance.Col) = "" Then txtTmp.Visible = False: mshBalance.SetFocus: Exit Sub
                mshBalance.TextMatrix(mshBalance.Row, mshBalance.Col) = ""
            End If
            
            Call ShowMoney(mintDefault <> mshBalance.Row)
        End If
        mshBalance.SetFocus
        txtTmp.Visible = False
        
        If mshBalance.Row = mshBalance.Rows - 1 Then
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            '��һ�д���
            mshBalance.Row = mshBalance.Row + 1
            mshBalance.Col = 1
            If mshBalance.Row - (mshBalance.Height \ mshBalance.RowHeight(0) - 2) > 1 Then
                mshBalance.TopRow = mshBalance.Row - (mshBalance.Height \ mshBalance.RowHeight(1) - 2)
            End If
        End If
    End If
End Sub
Private Sub ShowMoney(Optional ByVal blnAutoSet As Boolean)
    Dim curPay As Currency, curBalance As Currency
    Dim blnCent As Boolean, i As Long, bln���ڲ��� As Boolean
        
    If blnAutoSet And mintDefault > 0 Then      '���ݲ���Զ���ƽ������
        For i = 1 To mshBalance.Rows - 1
            If mshBalance.RowData(i) <> PayType.ҽ�������ʻ� And mshBalance.RowData(i) <> PayType.ҽ���������� Then
                curPay = curPay + Val(mshBalance.TextMatrix(i, 1))
            End If
        Next
        curBalance = mcurʵ�ս�� + mcur�շ���� - (curPay + mcurMediCare + mcur��Ԥ����)
    
        'ʣ�ಿ�ݳ������õ�ȱʡ���㷽ʽ��
        curBalance = Val(mshBalance.TextMatrix(mintDefault, 1)) + curBalance
        If mshBalance.RowData(mintDefault) = PayType.�ֽ� And mblnCent Then   '�ֽ�ʱҪ���зֱҴ���
            mshBalance.TextMatrix(mintDefault, 1) = Format(CentMoney(curBalance), "0.00")
        Else
            mshBalance.TextMatrix(mintDefault, 1) = Format(curBalance, "0.00")
        End If
        If Val(mshBalance.TextMatrix(mintDefault, 1)) = 0 Then mshBalance.TextMatrix(mintDefault, 1) = ""
        
        curPay = 0
        For i = 1 To mshBalance.Rows - 1
            curPay = curPay + Val(mshBalance.TextMatrix(i, 1))
        Next
        
        mcur�շ���� = curPay - mcurʵ�ս��
        txtPay.ToolTipText = "��ʽ�����,�����:" & Format(mcur�շ����, "0.00")
        
    Else
        bln���ڲ��� = True
    End If
    
    curPay = 0
    For i = 1 To mshBalance.Rows - 1
        If mshBalance.RowData(i) <> PayType.ҽ�������ʻ� And mshBalance.RowData(i) <> PayType.ҽ���������� Then
            curPay = curPay + Val(mshBalance.TextMatrix(i, 1))
        End If
    Next
    
    If bln���ڲ��� Then
       txtMargin.Text = Format(mcurʵ�ս�� + mcur�շ���� - (curPay + mcurMediCare + mcur��Ԥ����), "0.00")
    Else
        txtMargin.Text = "0.00"
    End If
    
    If Val(txt�ɿ�.Text) > 0 Then Call txt�ɿ�_Change
End Sub
Private Sub txtTmp_LostFocus()
    txtTmp.Visible = False
End Sub

Private Sub txtTmp_Validate(Cancel As Boolean)
    txtTmp.Visible = False
End Sub


Private Sub txtԤ�����_GotFocus()
    zlControl.TxtSelAll txtԤ�����
    txtԤ�����.Tag = txtԤ�����.Text  '��¼ԭֵ
End Sub

Private Sub txtԤ�����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    End If
    If InStr(txtԤ�����.Text, ".") > 0 And Chr(KeyAscii) = "." Then KeyAscii = 0:  Exit Sub
    If InStr("0123456789." & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtԤ�����_Validate(Cancel As Boolean)
    If Trim(txtԤ�����.Text) = "" Then
        txtԤ�����.Text = "0.00"
    ElseIf Not IsNumeric(txtԤ�����.Text) Then
        MsgBox "��Ч��ֵ��", vbInformation, gstrSysName
        zlControl.TxtSelAll txtԤ�����: Cancel = True: Exit Sub
        
    ElseIf Val(txtԤ�����.Text) < 0 Then
        MsgBox "Ԥ���������Ϊ����", vbInformation, gstrSysName
        zlControl.TxtSelAll txtԤ�����: Cancel = True: Exit Sub
        
    ElseIf Val(txtԤ�����.Text) > 0 And (mcurʵ�ս�� - mcurMediCare) < 0 Then
        MsgBox "����Ӧ�����Ϊ��ʱ����ʹ��Ԥ���", vbInformation, gstrSysName
        txtԤ�����.Text = "0.00"
        zlControl.TxtSelAll txtԤ�����: Cancel = True: Exit Sub
        
    ElseIf Val(txtԤ�����.Text) > mcurԤ����� Then
        MsgBox "Ԥ��������ܳ������˵�Ԥ�����:" & CStr(mcurԤ�����) & " ��", vbInformation, gstrSysName
        txtԤ�����.Text = CStr(mcurԤ�����)
        zlControl.TxtSelAll txtԤ�����: Cancel = True: Exit Sub
        
    ElseIf Val(txtԤ�����.Text) > (mcurʵ�ս�� - mcurMediCare) And Val(txtԤ�����.Text) <> 0 Then
        MsgBox "Ԥ��������ܴ���Ӧ�����:" & CStr((mcurʵ�ս�� - mcurMediCare)) & " ��", vbInformation, gstrSysName
        zlControl.TxtSelAll txtԤ�����: Cancel = True: Exit Sub
    Else
        txtԤ�����.Text = Format(txtԤ�����.Text, "0.00")
    End If

    If Val(txtԤ�����.Text) <> Val(txtԤ�����.Tag) Then
        
        mcur��Ԥ���� = Val(txtԤ�����.Text)
        Call ShowMoney(True)
        
        Dim curPay As Currency, i As Long
        curPay = 0
        For i = 1 To mshBalance.Rows - 1
            If mshBalance.RowData(i) <> PayType.ҽ�������ʻ� And mshBalance.RowData(i) <> PayType.ҽ���������� Then
                curPay = curPay + Val(mshBalance.TextMatrix(i, 1))
            End If
        Next
        txtPay.Text = Format(curPay, "0.00")
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mblnClickOK Then Cancel = 1
End Sub

Private Sub mshBalance_DblClick()
    If Not txtTmp.Visible And mshBalance.Row >= 1 And mshBalance.Col >= 1 And _
        mshBalance.RowData(mshBalance.Row) <> PayType.ҽ�������ʻ� And mshBalance.RowData(mshBalance.Row) <> PayType.ҽ���������� Then
        With txtTmp
            .MaxLength = IIf(mshBalance.Col = 2, 30, 10)
            .Left = mshBalance.Left + mshBalance.CellLeft + 15
            .Top = mshBalance.Top + mshBalance.CellTop + (mshBalance.CellHeight - txtTmp.Height) / 2 - 15
            .Width = mshBalance.CellWidth - 60
            .ForeColor = mshBalance.CellForeColor
            .BackColor = mshBalance.CellBackColor
            .Alignment = IIf(mshBalance.Col = 1, 1, 0)
            .Text = mshBalance.TextMatrix(mshBalance.Row, mshBalance.Col)
            .SelStart = 0: .SelLength = Len(.Text)
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub

Private Sub mshBalance_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If mshBalance.Col = 0 Then
            mshBalance.Col = 1
        ElseIf mshBalance.Row < mshBalance.Rows - 1 Then
            mshBalance.Row = mshBalance.Row + 1
            mshBalance.Col = 1
            If mshBalance.Row - (mshBalance.Height \ mshBalance.RowHeight(0) - 2) > 1 Then
                mshBalance.TopRow = mshBalance.Row - (mshBalance.Height \ mshBalance.RowHeight(1) - 2)
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub mshBalance_KeyPress(KeyAscii As Integer)
    If Not txtTmp.Visible And mshBalance.Row >= 1 And mshBalance.Col > 0 And KeyAscii <> 13 And KeyAscii <> vbKeyEscape And _
            mshBalance.RowData(mshBalance.Row) <> PayType.ҽ�������ʻ� And mshBalance.RowData(mshBalance.Row) <> PayType.ҽ���������� Then
        
        If mshBalance.Col = 1 Then
            If InStr("0123456789.-", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        Else '������������ַ�����
            If InStr("'|,", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Beep: Exit Sub
        End If
        
        With txtTmp
            .MaxLength = IIf(mshBalance.Col = 2, 30, 10)
            .Left = mshBalance.Left + mshBalance.CellLeft + 15
            .Top = mshBalance.Top + mshBalance.CellTop + (mshBalance.CellHeight - .Height) / 2 - 15
            .Width = mshBalance.CellWidth - 60
            .ForeColor = mshBalance.CellForeColor
            .BackColor = mshBalance.CellBackColor
            .Alignment = IIf(mshBalance.Col = 1, 1, 0)
            .Text = UCase(Chr(KeyAscii))
            .SelStart = 1
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub


Private Sub txt�ɿ�_Change()
    Dim cur�ֽ� As Currency, i As Long
    For i = 1 To mshBalance.Rows - 1
        If mshBalance.RowData(i) = PayType.�ֽ� Then
            cur�ֽ� = Val(mshBalance.TextMatrix(i, 1))
            Exit For
        End If
    Next
    mcur�ɿ� = Val(txt�ɿ�.Text)
    If mcur�ɿ� = 0 Then txt�Ҳ�.Text = "0.00": Exit Sub
    txt�Ҳ�.Text = Format(mcur�ɿ� - cur�ֽ�, "0.00")
End Sub

Private Sub txt�ɿ�_GotFocus()
    Call zlControl.TxtSelAll(txt�ɿ�)
End Sub

Private Sub txt�ɿ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(txt�ɿ�.Text) = 0 Then txt�ɿ�.Text = "0.00"
        If txt�ɿ�.Text <> "0.00" Then
            If Val(txt�Ҳ�.Text) >= 0 Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                MsgBox "�ɿ����,�벹��Ӧ�ɽ�", vbInformation, gstrSysName
                txt�ɿ�.SetFocus
                zlControl.TxtSelAll txt�ɿ�
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab) '�����ۼӽɿ�
        End If
    Else
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        ElseIf KeyAscii = Asc(".") And InStr(txt�ɿ�.Text, ".") > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txt�ɿ�_LostFocus()
    txt�ɿ�.Text = Format(Val(txt�ɿ�.Text), "0.00")
End Sub

Private Sub txt�ɿ�_Validate(Cancel As Boolean)
    txt�ɿ�.Text = Format(Val(txt�ɿ�.Text), "0.00")
End Sub


Private Function GetMoneyInfo(lng����ID As Long, Optional curModiMoney As Currency) As ADODB.Recordset
'���ܣ���ȡָ�����˵�ʣ���
    Dim strSql As String
        
    If curModiMoney = 0 Then
        strSql = "Select Nvl(�������,0) as �������,Nvl(Ԥ�����,0) as Ԥ����� From ������� Where ����=1 And ����=1 And ����ID=[1]"
    Else
        strSql = "Select Nvl(�������,0)-[2] as �������,Nvl(Ԥ�����,0) as Ԥ����� From ������� Where ����=1 And ����=1 And ����ID=[1]"
    End If
    On Error GoTo errH
    Set GetMoneyInfo = zlDatabase.OpenSQLRecord(strSql, "mdlOutExse", lng����ID, curModiMoney)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CentMoney(ByVal curMoney As Currency) As Currency
'���ܣ���ָ�����ֱҴ��������д���,���ش����Ľ��
'������curMoney=Ҫ���зֱҴ���Ľ��(ΪӦ�ɽ��,2λС��)
'      mBytMoney=
'         0.������
'         1.��ȡ�������뷨,eg:0.51=0.50;0.56=0.60
'         2.�����շ�,eg:0.51=0.60,0.56=0.60
'         3.����շ�,eg:0.51=0.50,0.56=0.50
'         4.�����������˫,eg:0.14=0.10,0.16=0.20,0.151=0.20,0.15=0.20,0.25=0.20
'           �����������˫,����ҹ���ѧ����ίԱ����ʽ�䲼�ġ�������Լ����,������vb��Round����,�������������ְ�����λ����ʱ�����Ը����ֽ���������Լ
'           �����м����뷨:���������忼�ǣ�������ͽ�һ�������㿴��ż����ǰΪżӦ��ȥ����ǰΪ��Ҫ��һ
'         5.�������塢�������,�Խǽ��д�������Ҫ�ȶԷֱҽ�������,��0.29(��)���¶�����ǣ�0.80(��)���϶����ǣ�0.3-0.79����Ϊ0.5��
'         6-��������:eg:0.15=0.10:0.16=0.2:    ����:34519
    Dim intSign As Integer, curTmp As Currency

    If mBytMoney = 0 Then
        CentMoney = Format(curMoney, "0.00")
    ElseIf mBytMoney = 1 Then
        curMoney = Format(curMoney, "0.00")    '��ȡ��λ���,�ٴ���ֱ�,��:0.248 ��0.3
        CentMoney = Format(curMoney, "0.0")
    ElseIf mBytMoney = 2 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        If Int(curMoney * 10) / 10 = curMoney Then
            CentMoney = intSign * curMoney
        Else
            CentMoney = intSign * Int(curMoney * 10 + 1) / 10
        End If
    ElseIf mBytMoney = 3 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        curMoney = Int(curMoney * 10) / 10
        CentMoney = intSign * curMoney
    ElseIf mBytMoney = 4 Then
        CentMoney = Format(Round(curMoney, 1), "0.00")
    ElseIf mBytMoney = 5 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        curTmp = Format(curMoney - Int(curMoney), "0.0")
        If curTmp >= 0.8 Then
            curTmp = 1
        ElseIf curTmp < 0.3 Then
            curTmp = 0
        Else
            curTmp = 0.5
        End If
        CentMoney = intSign * Format(Int(curMoney) + curTmp, "0.00")
    ElseIf mBytMoney = 6 Then
         '���˺� ����:34519 ��������:eg:0.15=0.10:0.16=0.2:    ����:2010-12-06 09:58:02
          CentMoney = Format(Format(curMoney - 0.01, "0.0"), "0.00")
    End If
End Function

 

 
