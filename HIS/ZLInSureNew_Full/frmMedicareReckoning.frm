VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMedicareReckoning 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ҽ�����˽���У��"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9675
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
   ScaleHeight     =   5745
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ �� (&C)"
      Height          =   420
      Left            =   8160
      TabIndex        =   17
      Top             =   5040
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtMoney 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E7CFBA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2760
      MaxLength       =   10
      TabIndex        =   12
      Top             =   1320
      Visible         =   0   'False
      Width           =   1005
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
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   5040
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
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   4350
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   5040
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
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   4470
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   240
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
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   240
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
      TabIndex        =   7
      Top             =   4680
      Width           =   9885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ ��(&O)"
      Height          =   420
      Left            =   6480
      TabIndex        =   6
      Top             =   5040
      Width           =   1395
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshMoney 
      Height          =   3345
      Left            =   5280
      TabIndex        =   3
      Top             =   960
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   5900
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
      FormatString    =   "^ ���㷽ʽ |^ ������ |^   �������  "
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDeposit 
      Height          =   3345
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   5900
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
   Begin VB.Label lblӦ�� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ӧ��:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   7680
      TabIndex        =   16
      Tag             =   "Ӧ��:"
      Top             =   4440
      Width           =   600
   End
   Begin VB.Label lblҽ��֧�� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ��֧��:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   5280
      TabIndex        =   15
      Tag             =   "ҽ��֧��:"
      Top             =   4440
      Width           =   1080
   End
   Begin VB.Label lbl��Ԥ�� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ԥ��:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   2280
      TabIndex        =   14
      Tag             =   "��Ԥ��:"
      Top             =   4440
      Width           =   840
   End
   Begin VB.Label lblԤ����� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ԥ�����:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   120
      TabIndex        =   13
      Tag             =   "Ԥ�����:"
      Top             =   4440
      Width           =   1080
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
      Left            =   240
      TabIndex        =   11
      Top             =   5160
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
      Left            =   3240
      TabIndex        =   10
      Top             =   5160
      Width           =   960
   End
   Begin VB.Label lblMargin 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ӧ�����"
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
      Left            =   3360
      TabIndex        =   9
      Top             =   360
      Width           =   960
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ʽ��"
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
      Left            =   240
      TabIndex        =   8
      Top             =   360
      Width           =   960
   End
End
Attribute VB_Name = "frmMedicareReckoning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
'���볣�����ܶ���ɹ����ģ�������ʹ�õ��ĵط��������壬�ڱ���ʱͳһ�޸�
#Const gverControl = 99  ' 0-��֧�ֶ�̬ҽ��(9.19��ǰ),1-֧�ֶ�̬ҽ���޸��Ӳ���(9.22��ǰ) , _
'    2-����������������ʽ��������һ��;����������ԭʼ��������һ��;�����շ�����������;3-������������GetNextNO();4-V10.24�����ϰ汾
'    99-���н������Ӹ��Ӳ���(���°�)

Private mbytInFun As Byte '0-����ģ�����,1-ҽ��ģ�����

Private mlng����ID As Long
Private mlng����ID As Long
Private mbln��;���� As Boolean     '��Ժ����,δ�����Ԥ�����Ҫ��Ϊ�ֽ�
Private mstr���ս��� As String
Private mstr������Ϣ As String      '�������,��������,�����ʺ�
Private mcur���ʽ�� As Currency
Private mcurԤ����� As Currency
Private mintInsure As Integer       '�����ж��Ƿ�֧�ֱַҴ���
Private mcur�ɿ� As Currency


Private mcur�շ���� As Currency
Private mblnOK  As Boolean
Private mintDefault As Integer      'ȱʡ���㷽ʽ��(Ϊ0��ʾû��)
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
Private mstrDec As String
Private mBytMoney As Byte '�շѷֱҴ�����


Public Function ShowMeFromOut(ByRef frmParent As Object, ByVal lng����ID As Long) As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, lng����ID As Long, strValue As String
    
    On Error GoTo errH
    If Not IsZLHIS10 Then
        ShowMeFromOut = frmMedicareReckoning9.ShowMeFromOut(frmParent, lng����ID)
        Exit Function
    End If
    
    mlng����ID = lng����ID
    strSQL = "Select a.����ID,a.��¼����,a.���㷽ʽ,a.�������,b.���� ��������,a.��Ԥ��,a.�ɿλ,a.��λ������,a.��λ�ʺ� " & _
             "   From ����Ԥ����¼ a,���㷽ʽ b " & _
             "   Where a.��¼״̬ = 1 And a.���㷽ʽ = B.���� And ����id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���ս������", lng����ID)
    mlng����ID = Val("" & rsTmp!����ID)
    
    mbln��;���� = True     '�޷��������ݿ���Ϣ����,Ĭ��Ϊ��õķ�ʽ:��;����,�������,����Ա����ȥ����ÿ��Ԥ���ĳ���,�Ա����ֽ�
        
    rsTmp.Filter = "(��¼����=2 And ��������=3) or (��¼����=2 And ��������=4)"
    If rsTmp.RecordCount > 0 Then mstr������Ϣ = zlCommFun.Nvl(rsTmp!�ɿλ, " ") & "," & zlCommFun.Nvl(rsTmp!��λ������, " ") & "," & zlCommFun.Nvl(rsTmp!��λ�ʺ�, " ")
       
    
    rsTmp.Filter = 0    '����ȡʵ�ս��,��Ϊ���������ٽ���ʱ,������ϸû��ʵ�ս��
    strSQL = "Select Sum(���ʽ��) As ���ʽ��" & _
             " From סԺ���ü�¼" & _
             " Where ���ӱ�־<>9 And ����id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���ս������", lng����ID)
    mcur���ʽ�� = Val("" & rsTmp!���ʽ��)
    
    
    '������Ϣ
    rsTmp.Filter = 0
    strSQL = "Select ���㷽ʽ,��� From ҽ���˶Ա� Where ����id = [1] And ���㷽ʽ<>'�ֽ�'"  'ҽ���ܿصĹ��̶̹�д����һ��"�ֽ�"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���ս������", lng����ID)
    mstr���ս��� = ""   '���㷽ʽ|������||
    For i = 1 To rsTmp.RecordCount
        mstr���ս��� = mstr���ս��� & "||" & rsTmp!���㷽ʽ & "|" & rsTmp!���
        rsTmp.MoveNext
    Next
    If mstr���ս��� <> "" Then mstr���ս��� = Mid(mstr���ս���, 3)
    
    
    mintInsure = 0
    If mlng����ID <> 0 Then
        strSQL = "Select ���� From ������ҳ Where ����id = [1]" & _
                 " And ��ҳid = (Select Max(��ҳid) From ������ҳ Where ����id = [1])"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���ս������", mlng����ID)
        If Not rsTmp.EOF Then mintInsure = zlCommFun.Nvl(rsTmp!����, 0)
    End If
    
    #If gverControl >= 4 Then
        mstrDec = "0." & String(Val(zlDatabase.GetPara(9, glngSys, , 2)), "0")
        strValue = zlDatabase.GetPara(14, glngSys, , 0)
    #Else
        mstrDec = "0." & String(Val(GetPara(9, glngSys, , , 2)), "0")
        strValue = GetPara(14, glngSys, , , 0)
    #End If
    mBytMoney = Val(IIf(Len(strValue) = 1, strValue, Mid(strValue, 3, 1)))
    
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


Public Function ShowMe(ByRef frmParent As Object, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal bln��;���� As Boolean, _
        ByVal cur���ʽ�� As Currency, ByVal str���ս��� As String, ByVal str������Ϣ As String, ByVal intInsure As Integer, _
        ByVal strȱʡ���λ�� As String, ByVal bytȱʡ�ֱҷ�ʽ As Byte, ByVal cur�ɿ� As Currency) As Boolean
    
    mlng����ID = lng����ID
    mlng����ID = lng����ID
    mbln��;���� = bln��;����
    mstr���ս��� = str���ս���
    mstr������Ϣ = str������Ϣ      '����ҽ���洢:�������,��������,�����ʺ�
    mcur���ʽ�� = cur���ʽ��
    mintInsure = intInsure
    mcur�ɿ� = cur�ɿ�
    
    mstrDec = strȱʡ���λ��
    mBytMoney = bytȱʡ�ֱҷ�ʽ
    
    mbytInFun = 0
    Me.Show 1, frmParent
    
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    mblnOK = False
    mblnClickOK = True: Unload Me
End Sub

Private Sub cmdOK_Click()
    '�������
    Dim i As Long
    Dim str���ʽ��� As String, str���NO As String, str��Ԥ�� As String
    
    If Val(txtMargin.Text) <> 0 Then
        If Val(txtMargin.Text) > 0 Then
            MsgBox "����֧������,�밴����ʾ�Ĳ��", vbExclamation, gstrSysName
            mshMoney.SetFocus: Exit Sub
        Else
            MsgBox "����֧��������,�밴����ʾ�Ĳ���˿", vbExclamation, gstrSysName
            mshMoney.SetFocus: Exit Sub
        End If
    End If
    
    '��������
    str���ʽ��� = ""
    For i = 1 To mshMoney.Rows - 1
        If Val(mshMoney.TextMatrix(i, 1)) <> 0 Then
            str���ʽ��� = str���ʽ��� & "||" & mshMoney.TextMatrix(i, 0) & "|" & Val(mshMoney.TextMatrix(i, 1)) & "|"
            
            If mshMoney.RowData(i) <> PayType.ҽ�������ʻ� And mshMoney.RowData(i) <> PayType.ҽ���������� Then
                 'Oracle���̸��ݽ�������ֶ��ж��Ƿ�ҽ��,���ԽɷѵĽ�����벻�ܺ���,��
                 '���㷽ʽ|������|�������||.....
                str���ʽ��� = str���ʽ��� & IIf(mshMoney.TextMatrix(i, 2) = "", " ", mshMoney.TextMatrix(i, 2))
            Else
                str���ʽ��� = str���ʽ��� & mstr������Ϣ
                '���㷽ʽ|������|�������,��������,�����ʺ�||.....
            End If
        End If
    Next
    str���ʽ��� = Mid(str���ʽ���, 3)
    
    For i = 1 To mshDeposit.Rows - 1
        If Val(mshDeposit.TextMatrix(i, mshDeposit.Cols - 1)) <> 0 Then     'ID|���ݺ�|���|��¼״̬||  IdΪ���ʾ��Ԥ�����(�ǵ�һ��)
            str��Ԥ�� = str��Ԥ�� & "||" & mshDeposit.TextMatrix(i, 0) & "|" & mshDeposit.TextMatrix(i, 1) & "|" & Val(mshDeposit.TextMatrix(i, mshDeposit.Cols - 1)) & "|" & Val(mshDeposit.RowData(i))
        End If
    Next
    If str��Ԥ�� <> "" Then str��Ԥ�� = Mid(str��Ԥ��, 3)
    If mcur�շ���� <> 0 Then str���NO = zlDatabase.GetNextNO(14)
    
    gstrSQL = "zl_סԺ�շѽ���_Update(" & mlng����ID & ",'" & IIf(str���ʽ��� = "", "", str���ʽ���) & "','" & IIf(str��Ԥ�� = "", "", str��Ԥ��) & "'," & _
            mcur�շ���� & ",'" & IIf(str���NO = "", "", str���NO) & "'" & _
            IIf(Val(txt�ɿ�.Text) <> 0, "," & Val(txt�ɿ�.Text) & "," & Val(txt�Ҳ�.Text), "") & ")"
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
    Dim strSQL As String, i As Long, j As Long
    Dim arrMediCare As Variant
    Dim bln������� As Boolean, blnExist As Boolean
    Dim str���õ�ҽ�����㷽ʽ As String
    
    '������ʼ
    mblnCent = gclsInsure.GetCapability(support�ֱҴ���, , mintInsure)
    mcur�շ���� = 0
    mblnOK = False
    mblnClickOK = False
    mintDefault = 0
    mcurMediCare = 0
    
    'ȷ����ȡ����ť
    If mbytInFun = 0 Then
        cmdOK.Left = cmdCancel.Left
        cmdCancel.Visible = False
    Else
        cmdCancel.Visible = True
    End If
    
    '��ʾԤ����ϸ
    Call AdjustDepost
    Set rsTmp = GetDepositBefor(mlng����ID)
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            mshDeposit.Redraw = False
            mshDeposit.Rows = rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                mshDeposit.Row = i
                mshDeposit.COL = mshDeposit.Cols - 1: mshDeposit.CellBackColor = txtMoney.BackColor
                mshDeposit.COL = mshDeposit.Cols - 2: mshDeposit.CellBackColor = 12900351
                
                mshDeposit.RowData(i) = IIf(IsNull(rsTmp!��¼״̬), 0, rsTmp!��¼״̬)
                mshDeposit.TextMatrix(i, 0) = rsTmp!ID
                mshDeposit.TextMatrix(i, 1) = rsTmp!NO

                mshDeposit.TextMatrix(i, 2) = Format(rsTmp!����, "yyyy-MM-dd")
                mshDeposit.TextMatrix(i, 3) = IIf(IsNull(rsTmp!���㷽ʽ), "", rsTmp!���㷽ʽ)
                mshDeposit.TextMatrix(i, 4) = Format(rsTmp!���, "0.00")
                mshDeposit.TextMatrix(i, 5) = Format(rsTmp!���, "0.00")
                rsTmp.MoveNext
            Next
            mshDeposit.Row = 1: mshDeposit.COL = mshDeposit.Cols - 1
            mshDeposit.Redraw = True
        End If
    End If
    
    
    '��ʾ���ս��㼰�ָ����㷽ʽ,��ʹ��֧��ʹ�ø���,Ҳ������,����ҽ���Ĳ������
    arrMediCare = Array()                   '���㷽ʽ|������||
    If mstr���ս��� <> "" Then arrMediCare = Split(mstr���ս���, "||")
    
    On Error GoTo errH
    
    strSQL = _
        " Select Distinct B.����,B.����,B.����,A.ȱʡ��־" & _
        " From ���㷽ʽӦ�� A,���㷽ʽ B" & _
        " Where ((A.Ӧ�ó���='����' And B.����<>3 And B.����<>4) OR (B.����=3 OR B.����=4)) And B.����=A.���㷽ʽ(+) " & _
        " Order by B.����,B.����"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    strSQL = "Select Ӧ�ó���,���㷽ʽ From ���㷽ʽӦ�� Where Ӧ�ó���='����'"
    Call zlDatabase.OpenRecordset(rsӦ�ó���, strSQL, Me.Caption)
    
    With mshMoney
        .ColAlignment(0) = 1    '���㷽ʽ�����
        .ColAlignment(1) = 7    '����Ҷ���
        .Redraw = False
        .Rows = rsTmp.RecordCount + 1
        i = 1
        Do While Not rsTmp.EOF
            .RowData(i) = zlCommFun.Nvl(rsTmp!����, PayType.�ֽ�)                '�����ж��Ƿ�����޸Ľ��,�Լ��Ƿ����ֽ�
            .TextMatrix(i, 0) = rsTmp!����
            .TextMatrix(i, 1) = "0.00"
            
            'ȱʡ���㷽ʽ(û�������ֽ�) ��������ҽ��
            If .RowData(i) <> PayType.ҽ�������ʻ� And .RowData(i) <> PayType.ҽ���������� Then
                If zlCommFun.Nvl(rsTmp!ȱʡ��־, 0) = 1 Then mintDefault = i
                If zlCommFun.Nvl(rsTmp!����, 1) = 1 And mintDefault = 0 Then mintDefault = i
                i = i + 1
            Else
                '���ս���
                blnExist = False
                For j = 0 To UBound(arrMediCare)
                    If Split(arrMediCare(j), "|")(0) = rsTmp!���� Then
                        blnExist = True
                        rsӦ�ó���.Filter = "���㷽ʽ='" & rsTmp!���� & "'"
                        
                        If rsӦ�ó���.EOF Then
                            MsgBox "ע��:���㷽ʽ[" & rsTmp!���� & "]δ����Ӧ����[����]����,�뵽[���㷽ʽ����]������!", vbInformation, gstrSysName
                        End If
                        
                        .TextMatrix(i, 1) = Split(arrMediCare(j), "|")(1)
                        .TextMatrix(i, 2) = ""    '�޽������
                        mcurMediCare = mcurMediCare + Val(.TextMatrix(i, 1))
                        Exit For
                    End If
                Next
                
                If blnExist Then
                     For j = 0 To .Cols - 1
                         .Row = i: .COL = j: .CellBackColor = &HE7CFBA
                     Next
                     i = i + 1
                End If
                
                str���õ�ҽ�����㷽ʽ = str���õ�ҽ�����㷽ʽ & "," & rsTmp!����
            End If
            rsTmp.MoveNext
        Loop
        
        .Rows = i
        .Redraw = True
    End With
    
    
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
    
    
    '���ʽ��
    txtTotal.Text = Format(mcur���ʽ��, mstrDec)
    
    '��Ԥ��,���ʽ���ȥҽ����ʽ���ʺ�����
    Call ShowMoney(True)
    
    If mintDefault > 0 Then
        mshMoney.Row = mintDefault: mshMoney.COL = 0
        mshMoney.CellFontBold = True
        mshMoney.COL = 1
    Else        '���㷽ʽû��ȱʡֵ,�������ֽ�ʽ�����
        mshMoney.Row = 1: mshMoney.COL = 1
    End If
    txt�ɿ�.Text = Format(mcur�ɿ�, "0.00")
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function ShowMoney(Optional ByVal blnAutoSet As Boolean) As String
'���ܣ����ú���ʾ����ĸ��ֽ��

    Dim i As Long, j As Long
    Dim cur���ʺϼ� As Currency, curMoney As Currency
    Dim curԤ���ϼ� As Currency, cur��Ԥ���ϼ� As Currency, curӦ�ɽ�� As Currency
    Dim bln���ڲ��� As Boolean  'ֻ�е�û��ȱʡ���㷽ʽ,�����޸�ȱʡ���㷽ʽ�Ľ��ʱ,����
        
    
    '�����Զ���Ԥ������Ľ�����
    '---------------------------------------------------------------------------------------------
    If blnAutoSet Then
        '���ó�Ԥ��(���ʺϼ� - ���պϼ�)
        cur���ʺϼ� = mcur���ʽ�� - mcurMediCare
        
        If mshDeposit.TextMatrix(1, 0) <> "" Then   '����û��Ԥ��,ȫ���ֿ�
            If Not mbln��;���� Then
                '��Ժ����ȫ��������(����˾����ָ�)
                For i = 1 To mshDeposit.Rows - 1
                    mshDeposit.TextMatrix(i, mshDeposit.Cols - 1) = Format(Val(mshDeposit.TextMatrix(i, mshDeposit.Cols - 2)), "0.00")
                    cur���ʺϼ� = cur���ʺϼ� - Val(mshDeposit.TextMatrix(i, mshDeposit.Cols - 1))
                Next
            Else
                '��;����ֻ���㹻��
                For i = 1 To mshDeposit.Rows - 1
                    If cur���ʺϼ� = 0 Then
                        mshDeposit.TextMatrix(i, mshDeposit.Cols - 1) = "0.00"
                    Else
                        If Val(mshDeposit.TextMatrix(i, mshDeposit.Cols - 2)) <= Format(cur���ʺϼ�, "0.00") Then
                            mshDeposit.TextMatrix(i, mshDeposit.Cols - 1) = Format(Val(mshDeposit.TextMatrix(i, mshDeposit.Cols - 2)), "0.00")
                        Else
                            mshDeposit.TextMatrix(i, mshDeposit.Cols - 1) = Format(cur���ʺϼ�, "0.00")
                        End If
                        cur���ʺϼ� = cur���ʺϼ� - Val(mshDeposit.TextMatrix(i, mshDeposit.Cols - 1))
                    End If
                Next
            End If
        End If
        
        'ʣ��Ӧ�ɲ��ݳ������õ�ȱʡ���㷽ʽ    '�ж��Ƿ�Ӧ�ý��зֱҴ���
        If mintDefault <> 0 Then
            If mshMoney.RowData(mintDefault) = PayType.�ֽ� And mblnCent Then '�ֽ�ʱҪ���зֱҴ���
                mshMoney.TextMatrix(mintDefault, 1) = Format(CentMoney(cur���ʺϼ�), "0.00")
            Else
                mshMoney.TextMatrix(mintDefault, 1) = Format(cur���ʺϼ�, "0.00")
            End If
        Else
            bln���ڲ��� = True
        End If
    
    '�޸ĳ�Ԥ����������
    Else
        cur���ʺϼ� = mcur���ʽ�� - GetSumMoney
        
        If mintDefault <> 0 And (Not Me.ActiveControl Is mshMoney Or _
                                Me.ActiveControl Is mshMoney And mintDefault <> mshMoney.Row) Then
            If mshMoney.RowData(mintDefault) = PayType.�ֽ� And mblnCent Then '�ֽ�ʱҪ���зֱҴ���
                mshMoney.TextMatrix(mintDefault, 1) = Format(Val(mshMoney.TextMatrix(mintDefault, 1)) + CentMoney(cur���ʺϼ�), "0.00")
            Else
                mshMoney.TextMatrix(mintDefault, 1) = Format(Val(mshMoney.TextMatrix(mintDefault, 1)) + cur���ʺϼ�, "0.00")
            End If
        Else
            bln���ڲ��� = True
        End If
    End If
    
        
    '��ʾ�����
    '-----------------------------------------------------------------------------------------------------
    curMoney = GetSumMoney(curԤ���ϼ�, cur��Ԥ���ϼ�, curӦ�ɽ��)
    mcur�շ���� = Format(curMoney - mcur���ʽ��, mstrDec)
    If bln���ڲ��� Then
        txtMargin.Text = Format(mcur���ʽ�� - curMoney, "0.00")
    Else
        txtMargin.Text = "0.00"
    End If
    txtMargin.ToolTipText = "�����:" & Format(mcur�շ����, mstrDec)
    
    
    lblԤ�����.Caption = lblԤ�����.Tag & Format(curԤ���ϼ�, "0.00")
    lblԤ�����.ToolTipText = "����δ��Ԥ��֮ǰ��Ԥ�����"
    lbl��Ԥ��.Caption = lbl��Ԥ��.Tag & Format(cur��Ԥ���ϼ�, "0.00")
    lblҽ��֧��.Caption = lblҽ��֧��.Tag & Format(mcurMediCare, "0.00")
    lblӦ��.Caption = lblӦ��.Tag & Format(curӦ�ɽ��, "0.00")
    
    
    lblԤ�����.Left = mshDeposit.Left
    lbl��Ԥ��.Left = lblԤ�����.Left + lblԤ�����.Width + 600
    lblҽ��֧��.Left = mshMoney.Left
    lblӦ��.Left = lblҽ��֧��.Left + lblҽ��֧��.Width + 600
End Function

Private Sub txtMoney_KeyPress(KeyAscii As Integer)
    Dim blnCent As Boolean, i As Long
    
    If KeyAscii <> 13 Then        '��������
        If InStr(txtMoney.Text, ".") > 0 And Chr(KeyAscii) = "." Then KeyAscii = 0: Beep: Exit Sub
        
        If txtMoney.Left > mshMoney.Left Then   '��������
            If mshMoney.COL = mshMoney.Cols - 1 Then    '�������,���������ڹ������ж��Ƿ���ҽ�����㷽ʽ
                If InStr("'|,", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Beep: Exit Sub
            Else
                If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
            End If
        Else    'Ԥ������
            If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
        End If
        KeyAscii = asc(UCase(Chr(KeyAscii)))
    Else
        KeyAscii = 0
         '��������ȷ��
        If txtMoney.Left > mshMoney.Left Then
            If mshMoney.COL = mshMoney.Cols - 1 Then    '��������
                If InStr(txtMoney.Text, "'") > 0 Or InStr(txtMoney.Text, "|") > 0 Or InStr(txtMoney.Text, ",") > 0 Then
                    Exit Sub
                End If
                
                mshMoney.TextMatrix(mshMoney.Row, mshMoney.COL) = Trim(txtMoney.Text)
                txtMoney.Visible = False
            Else
                If Trim(txtMoney.Text) = "" Or Not IsNumeric(Trim(txtMoney.Text)) Then
                    zlControl.TxtSelAll txtMoney: Call Beep: Exit Sub
                End If
                If mshMoney.RowData(mshMoney.Row) = PayType.�ֽ� And mblnCent Then
                    txtMoney.Text = Format(CentMoney(Val(txtMoney.Text)), "0.00")
                End If
                                
                If Val(mshMoney.TextMatrix(mshMoney.Row, mshMoney.COL)) <> Format(Val(txtMoney.Text), "0.00") Then
                    mshMoney.TextMatrix(mshMoney.Row, mshMoney.COL) = Format(Val(txtMoney.Text), "0.00")
                    txtMoney.Visible = False
                    mshMoney.SetFocus   '��������,ShowMoney���Դ��ж�
                    
                    Call ShowMoney
                Else
                    txtMoney.Visible = False
                    mshMoney.SetFocus
                End If
            End If
            
            If mshMoney.COL < mshMoney.Cols - 2 Then
                mshMoney.COL = mshMoney.COL + 1
            Else
                If mshMoney.Row = mshMoney.Rows - 1 Then
                    '��һ�ؼ�����
                    If GetӦ�� > 0 And txt�ɿ�.Visible Then
                        txt�ɿ�.SetFocus
                    ElseIf cmdOK.Visible And cmdOK.Enabled Then
                        cmdOK.SetFocus
                    End If
                Else
                    '��һ�д���
                    If mshMoney.RowData(mshMoney.Row) = PayType.��ҽ�����ֽ� Then
                       If mshMoney.COL = mshMoney.Cols - 2 Then
                            mshMoney.COL = mshMoney.Cols - 1
                       Else
                            mshMoney.Row = mshMoney.Row + 1
                            mshMoney.COL = mshMoney.Cols - 2
                       End If
                    Else
                        mshMoney.Row = mshMoney.Row + 1
                        mshMoney.COL = mshMoney.Cols - 2
                    End If
                    
                    If mshMoney.Row - (mshMoney.Height \ mshMoney.RowHeight(0) - 2) > 1 Then
                        mshMoney.TopRow = mshMoney.Row - (mshMoney.Height \ mshMoney.RowHeight(1) - 2)
                    End If
                End If
            End If
        
        'Ԥ������ȷ��
        Else
            If Trim(txtMoney.Text) = "" Or Not IsNumeric(Trim(txtMoney.Text)) Then
                zlControl.TxtSelAll txtMoney: Call Beep: Exit Sub
            End If
            
            '�޸Ĳ��ܳ�������
            If Val(txtMoney.Text) > Val(mshDeposit.TextMatrix(mshDeposit.Row, 4)) Then
                txtMoney.Text = Val(mshDeposit.TextMatrix(mshDeposit.Row, 4))
            End If
            
            If Val(mshDeposit.TextMatrix(mshDeposit.Row, mshDeposit.COL)) <> Format(Val(txtMoney.Text), "0.00") Then
                mshDeposit.TextMatrix(mshDeposit.Row, mshDeposit.COL) = Format(Val(txtMoney.Text), "0.00")
                txtMoney.Visible = False
                mshDeposit.SetFocus '��������
                
                Call ShowMoney
            Else
                txtMoney.Visible = False
                mshDeposit.SetFocus
            End If
            
            If mshDeposit.Row = mshDeposit.Rows - 1 Then
                '��һ�ؼ�����
                mshMoney.SetFocus
            Else
                '��һ�д���
                mshDeposit.Row = mshDeposit.Row + 1
                If mshDeposit.Row - (mshDeposit.Height \ mshDeposit.RowHeight(0) - 2) > 1 Then
                    mshDeposit.TopRow = mshDeposit.Row - (mshDeposit.Height \ mshDeposit.RowHeight(1) - 2)
                End If
                mshDeposit.COL = mshDeposit.Cols - 1
            End If
        End If
        
        If Val(txt�ɿ�.Text) > 0 Then Call txt�ɿ�_Change
    End If
End Sub

Private Sub txtMoney_LostFocus()
    txtMoney.Visible = False
End Sub

Private Sub txtMoney_Validate(Cancel As Boolean)
    If txtMoney.Visible Then Call txtMoney_KeyPress(13)
End Sub

Private Sub AdjustDepost()
    Dim bln As Boolean
    With mshDeposit
        bln = .Redraw
        .Redraw = False
        .Clear
        .Rows = 2: .Cols = 6
        
        .TextMatrix(0, 1) = "���ݺ�"
        .TextMatrix(0, 2) = "����"
        .TextMatrix(0, 3) = "���㷽ʽ"
        .TextMatrix(0, 4) = "���"
        .TextMatrix(0, 5) = "��Ԥ��"
        
        .ColAlignmentFixed(1) = 4: .ColAlignment(1) = 1
        .ColAlignmentFixed(2) = 4: .ColAlignment(2) = 6
        .ColAlignmentFixed(3) = 1: .ColAlignment(3) = 1
        .ColAlignmentFixed(4) = 4: .ColAlignment(4) = 7
        .ColAlignmentFixed(5) = 4: .ColAlignment(5) = 7
        
        .ColWidth(0) = 0
        
        .ColWidth(1) = 1100
        .ColWidth(2) = 1050
        .ColWidth(3) = 620
        .ColWidth(4) = 1100
        .ColWidth(5) = 1100
        
        .Row = 1: .COL = .Cols - 1
        .AllowUserResizing = flexResizeColumns
        .ScrollBars = flexScrollBarBoth
        
        .Redraw = bln
    End With
End Sub

Private Function GetDepositBefor(lng����ID As Long) As ADODB.Recordset
'���ܣ���ȡ���˱���ҽ������֮ǰ��ʣ��Ԥ������ϸ,�������γ�����Ԥ��
    
    Dim strSQL As String, strSub1 As String
    
    On Error GoTo errH
    
    '���Ӳ�ѯ��������Ԥ�����շѼ��˷�ʱ��һ��һ��,ע��ϵͳ�������ʵ�Ԥ�������Ԥ���˷�,��Ҫ���ϼ�¼״̬�ж�
    strSub1 = _
        "Select NO,Sum(Nvl(A.���,0)) as ��� From ����Ԥ����¼ A" & _
        " Where (A.����ID Is Null Or A.����ID=[1]) And Nvl(A.���, 0)<>0 And A.����ID=[2]" & _
        " Group by NO Having Sum(Nvl(A.���,0))<>0"

    strSQL = _
        "Select A.ID,A.��¼״̬,A.NO,A.�տ�ʱ�� as ����,A.���㷽ʽ,Nvl(A.���,0) as ���" & _
        " From ����Ԥ����¼ A,(" & strSub1 & ") B" & _
        " Where (A.����ID Is Null Or A.����ID=[1]) And Nvl(A.���,0)<>0" & _
        " And A.���㷽ʽ Not IN(Select ���� From ���㷽ʽ Where ����=5)" & _
        " And A.NO=B.NO And A.����ID=[2]" & _
        " Union All" & _
        " Select 0 as ID,��¼״̬,NO,�տ�ʱ�� as ����,���㷽ʽ,Sum(Nvl(���,0)-Nvl(��Ԥ��,0)) as ���" & _
        " From ����Ԥ����¼" & _
        " Where ��¼���� IN(1,11) And ����ID is Not NULL And ����ID<>[1] And Nvl(���,0)<>Nvl(��Ԥ��,0) And ����ID=[2]" & _
        " Having Sum(Nvl(���,0)-Nvl(��Ԥ��,0))<>0" & _
        " Group by ��¼״̬,NO,�տ�ʱ��,���㷽ʽ" & _
        " Order by ID,����,NO,���㷽ʽ"
    Set GetDepositBefor = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng����ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function GetSumMoney(Optional ByRef curԤ���ϼ� As Currency, Optional ByRef cur��Ԥ���ϼ� As Currency, Optional ByRef curӦ�ɽ�� As Currency) As Currency
    Dim i As Long
    Dim curMoney As Currency
    
    curԤ���ϼ� = 0: cur��Ԥ���ϼ� = 0: curӦ�ɽ�� = 0
    
    If mshDeposit.TextMatrix(1, 0) <> "" Then
        For i = 1 To mshDeposit.Rows - 1
            curMoney = curMoney + Val(mshDeposit.TextMatrix(i, mshDeposit.Cols - 1))
            curԤ���ϼ� = curԤ���ϼ� + Val(mshDeposit.TextMatrix(i, mshDeposit.Cols - 2))
            cur��Ԥ���ϼ� = cur��Ԥ���ϼ� + Val(mshDeposit.TextMatrix(i, mshDeposit.Cols - 1))
        Next
    End If
    For i = 1 To mshMoney.Rows - 1
        If IsNumeric(mshMoney.TextMatrix(i, 1)) Then
            curMoney = curMoney + Val(mshMoney.TextMatrix(i, 1))
            If mshMoney.RowData(i) <> PayType.ҽ�������ʻ� And mshMoney.RowData(i) <> PayType.ҽ���������� Then
                curӦ�ɽ�� = curӦ�ɽ�� + Val(mshMoney.TextMatrix(i, 1))
            End If
        End If
    Next
    
    GetSumMoney = curMoney
End Function

Private Sub Form_Unload(Cancel As Integer)
    If Not mblnClickOK Then Cancel = 1
End Sub

Private Sub mshDeposit_DblClick()
    If Not txtMoney.Visible And mshDeposit.Row >= 1 And mshDeposit.COL = mshDeposit.Cols - 1 Then
        With txtMoney
            .Left = mshDeposit.Left + mshDeposit.CellLeft + 15
            .Top = mshDeposit.Top + mshDeposit.CellTop + (mshDeposit.CellHeight - txtMoney.Height) / 2 - 15
            .Width = mshDeposit.CellWidth - 60
            .ForeColor = mshDeposit.CellForeColor
            .BackColor = mshDeposit.CellBackColor
            .Alignment = 1
            .Text = mshDeposit.TextMatrix(mshDeposit.Row, mshDeposit.COL)
            .SelStart = 0: .SelLength = Len(.Text)
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub

Private Sub mshDeposit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If mshDeposit.COL = 0 Then
            mshDeposit.COL = mshDeposit.COL + 1
        ElseIf mshDeposit.Row < mshDeposit.Rows - 1 Then
            mshDeposit.Row = mshDeposit.Row + 1
            mshDeposit.COL = mshDeposit.Cols - 1
            If mshDeposit.Row - (mshDeposit.Height \ mshDeposit.RowHeight(0) - 2) > 1 Then
                mshDeposit.TopRow = mshDeposit.Row - (mshDeposit.Height \ mshDeposit.RowHeight(1) - 2)
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub mshDeposit_KeyPress(KeyAscii As Integer)
    If Not txtMoney.Visible And KeyAscii <> 13 And KeyAscii <> vbKeyEscape Then
        If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
        With txtMoney
            .Left = mshDeposit.Left + mshDeposit.CellLeft + 15
            .Top = mshDeposit.Top + mshDeposit.CellTop + (mshDeposit.CellHeight - txtMoney.Height) / 2 - 15
            .Width = mshDeposit.CellWidth - 60
            .ForeColor = mshDeposit.CellForeColor
            .BackColor = mshDeposit.CellBackColor
            .Alignment = 1
            .Text = Chr(KeyAscii)
            .SelStart = 1
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub

Private Sub mshMoney_DblClick()
    If Not txtMoney.Visible And mshMoney.Row >= 1 And mshMoney.COL > 0 And _
        mshMoney.RowData(mshMoney.Row) <> PayType.ҽ�������ʻ� And mshMoney.RowData(mshMoney.Row) <> PayType.ҽ���������� Then
        
        With txtMoney
            .MaxLength = IIf(mshMoney.COL = 2, 30, 10)
            .Left = mshMoney.Left + mshMoney.CellLeft + 15
            .Top = mshMoney.Top + mshMoney.CellTop + (mshMoney.CellHeight - txtMoney.Height) / 2 - 15
            .Width = mshMoney.CellWidth - 60
            .ForeColor = mshMoney.CellForeColor
            .BackColor = mshMoney.CellBackColor
            .Alignment = IIf(mshMoney.COL = 2, 0, 1)
            .Text = mshMoney.TextMatrix(mshMoney.Row, mshMoney.COL)
            .SelStart = 0: .SelLength = Len(.Text)
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub

Private Sub mshMoney_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And mshMoney.Row >= 1 Then
        If mshMoney.COL = 0 Then
            mshMoney.COL = mshMoney.COL + 1
        Else
            If mshMoney.Row < mshMoney.Rows - 1 Then
                
                If mshMoney.RowData(mshMoney.Row) = PayType.��ҽ�����ֽ� Then
                   If mshMoney.COL = mshMoney.Cols - 2 Then
                        mshMoney.COL = mshMoney.Cols - 1
                   Else
                        mshMoney.Row = mshMoney.Row + 1
                        mshMoney.COL = mshMoney.Cols - 2
                   End If
                Else
                    mshMoney.Row = mshMoney.Row + 1
                    mshMoney.COL = mshMoney.Cols - 2
                End If
                If mshMoney.Row - (mshMoney.Height \ mshMoney.RowHeight(0) - 2) > 1 Then
                    mshMoney.TopRow = mshMoney.Row - (mshMoney.Height \ mshMoney.RowHeight(1) - 2)
                End If
            Else
                If GetӦ�� > 0 Then
                    txt�ɿ�.SetFocus
                Else
                    cmdOK.SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Sub mshMoney_KeyPress(KeyAscii As Integer)
    If Not txtMoney.Visible And mshMoney.Row >= 1 And mshMoney.COL > 0 And KeyAscii <> 13 And KeyAscii <> vbKeyEscape And _
         mshMoney.RowData(mshMoney.Row) <> PayType.ҽ�������ʻ� And mshMoney.RowData(mshMoney.Row) <> PayType.ҽ���������� Then
                        
        If mshMoney.COL = 1 Then
            If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
        Else '������������ַ�����,���������ڹ������ж��Ƿ���ҽ�����㷽ʽ
            If InStr("'||,", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Beep: Exit Sub
        End If
        
        With txtMoney
            .MaxLength = IIf(mshMoney.COL = 2, 30, 10)
            .Left = mshMoney.Left + mshMoney.CellLeft + 15
            .Top = mshMoney.Top + mshMoney.CellTop + (mshMoney.CellHeight - txtMoney.Height) / 2 - 15
            .Width = mshMoney.CellWidth - 60
            .ForeColor = mshMoney.CellForeColor
            .BackColor = mshMoney.CellBackColor
            .Alignment = IIf(mshMoney.COL = 2, 0, 1)
            .Text = UCase(Chr(KeyAscii))
            .SelStart = 1
            .ZOrder: .Visible = True
            .SetFocus
        End With
    End If
End Sub

Private Function GetӦ��() As Currency
    Dim i As Long
    
    For i = 1 To mshMoney.Rows - 1
        If mshMoney.RowData(i) = PayType.�ֽ� Then
            GetӦ�� = Val(mshMoney.TextMatrix(i, 1))
            Exit Function
        End If
    Next
End Function

Private Sub txt�ɿ�_Change()
    Dim cur�ֽ� As Currency, i As Long
    For i = 1 To mshMoney.Rows - 1
        If mshMoney.RowData(i) = PayType.�ֽ� Then
            cur�ֽ� = Val(mshMoney.TextMatrix(i, 1))
            Exit For
        End If
    Next
    If Val(txt�ɿ�.Text) = 0 Then txt�Ҳ�.Text = "0.00": Exit Sub
    txt�Ҳ�.Text = Format(Val(txt�ɿ�.Text) - cur�ֽ�, "0.00")
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
        ElseIf KeyAscii = asc(".") And InStr(txt�ɿ�.Text, ".") > 0 Then
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
        curTmp = curMoney - Int(curMoney)
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

