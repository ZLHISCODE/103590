VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmChangeGroup 
   Caption         =   "����תҽ��С��"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   Icon            =   "frmChangeGroup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   6135
   StartUpPosition =   1  '����������
   Begin VB.Frame fraGroup 
      Height          =   975
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   1560
      Width           =   5940
      Begin VB.ComboBox cboסԺҽʦ 
         Height          =   300
         Left            =   3900
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   180
         Width           =   1830
      End
      Begin VB.ComboBox cboҽ��С�� 
         Height          =   300
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   180
         Width           =   1890
      End
      Begin VB.ComboBox cbo����ҽʦ 
         Height          =   300
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   555
         Width           =   1890
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   300
         Left            =   3900
         TabIndex        =   3
         Top             =   555
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   19
         Format          =   "yyyy-MM-dd HH:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Чʱ��"
         Height          =   180
         Left            =   3120
         TabIndex        =   26
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺҽʦ"
         Height          =   180
         Left            =   3120
         TabIndex        =   25
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ҽʦ"
         Height          =   180
         Left            =   210
         TabIndex        =   24
         Top             =   615
         Width           =   720
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ҽ��С��"
         Height          =   180
         Left            =   30
         TabIndex        =   23
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3495
      TabIndex        =   4
      Top             =   2670
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4755
      TabIndex        =   5
      Top             =   2670
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   5940
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1005
         Width           =   1890
      End
      Begin VB.TextBox txtPre 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1005
         Width           =   1830
      End
      Begin VB.TextBox txtסԺ�� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   630
         Width           =   1890
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   690
      End
      Begin VB.TextBox txt�Ա� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   675
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   1890
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   630
         Width           =   1830
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ԭҽ��С��"
         Height          =   180
         Left            =   2940
         TabIndex        =   19
         Top             =   1065
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ��λ"
         Height          =   180
         Left            =   195
         TabIndex        =   18
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   4605
         TabIndex        =   11
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   3000
         TabIndex        =   9
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   540
         TabIndex        =   7
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "סԺ��"
         Height          =   180
         Left            =   375
         TabIndex        =   17
         Top             =   690
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ����"
         Height          =   180
         Left            =   3120
         TabIndex        =   16
         Top             =   690
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   315
      TabIndex        =   6
      Top             =   2670
      Width           =   1100
   End
End
Attribute VB_Name = "frmChangeGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mstrPirvs As String
Private mlngUnit As Long

Private mstrUnit As String
Private mrsPatiInfo As ADODB.Recordset
Private mfrmParent As Object

Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML

Private Sub cboҽ��С��_Click()
    Dim strSQL As String, strSQLҽ��С�� As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    Dim lngҽʦ As Long
    
    On Error GoTo errHandle:
    
    '���Ϊ����ָ����ҽ��С�飬��"סԺҽʦ������ҽʦ"���Ӷ�Ӧҽ��С���е�ҽ����ѡ��
    strSQLҽ��С�� = "Select Distinct A.ID, A.���, A.����, A.����" & vbNewLine & _
                    " From ��Ա�� A, ��Ա����˵�� B, ������Ա C, ҽ��С����Ա D" & vbNewLine & _
                    " Where A.ID = B.��Աid And A.ID = C.��Աid And a.id = d.��Աid And B.��Ա���� = 'ҽ��' And d.С��id = [1] And" & vbNewLine & _
                    "   (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) And" & vbNewLine & _
                    "   (Instr(',' || [2] || ',', ',' || C.����id || ',') > 0 Or a.����=[3]) And Instr(',' || [4] || ',', ',' || A.רҵ����ְ�� || ',') > 0" & vbNewLine & _
                    "   And (A.վ��=[5] Or A.վ�� is Null)" & vbNewLine & _
                    " Order By A.����"
    strSQL = "Select Distinct A.ID, A.���, A.����, A.����" & vbNewLine & _
                        " From ��Ա�� A, ��Ա����˵�� B, ������Ա C" & vbNewLine & _
                        " Where A.ID = B.��Աid And A.ID = C.��Աid And B.��Ա���� = 'ҽ��' And" & vbNewLine & _
                        "      (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) And" & vbNewLine & _
                        "      (Instr(',' || [1] || ',', ',' || C.����id || ',') > 0 Or A.����=[2]) And Instr(',' || [3] || ',', ',' || A.רҵ����ְ�� || ',') > 0" & vbNewLine & _
                        "      And (A.վ��=[4] Or A.վ�� is Null)" & _
                        " Order By A.����"
    If cboҽ��С��.ListIndex <> -1 And cboҽ��С��.ListIndex <> cboҽ��С��.ListCount - 1 Then
        If Val(cboҽ��С��.ItemData(cboҽ��С��.ListIndex)) > 0 Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQLҽ��С��, Me.Caption, Val(cboҽ��С��.ItemData(cboҽ��С��.ListIndex)), mstrUnit, CStr("" & mrsPatiInfo!סԺҽʦ), "����ҽʦ,������ҽʦ,����ҽʦ,ҽʦ,ҽʿ", gstrNodeNo)
            If Not rsTmp.RecordCount > 0 Then
                '���С��δ����ҽ�����򱣳���ǰ�Ŀ���ѡ��Χ
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrUnit, CStr("" & mrsPatiInfo!סԺҽʦ), "����ҽʦ,������ҽʦ,����ҽʦ,ҽʦ,ҽʿ", gstrNodeNo)
            End If
            
            If cboסԺҽʦ.ListIndex <> -1 Then
                lngҽʦ = cboסԺҽʦ.ItemData(cboסԺҽʦ.ListIndex)
            Else
                lngҽʦ = 0
            End If
            cboסԺҽʦ.Clear
            Do Until rsTmp.EOF
                cboסԺҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����
                cboסԺҽʦ.ItemData(cboסԺҽʦ.NewIndex) = rsTmp!ID
                rsTmp.MoveNext
            Loop
            '105133:��סԺҽʦ����ѡҽ��С��ʱ���ı�סԺҽʦ
            If lngҽʦ <> 0 Then Call cbo.SetIndex(cboסԺҽʦ.hWnd, cbo.FindIndex(cboסԺҽʦ, lngҽʦ))
        
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQLҽ��С��, Me.Caption, Val(cboҽ��С��.ItemData(cboҽ��С��.ListIndex)), mstrUnit, CStr("" & mrsPatiInfo!����ҽʦ), "����ҽʦ,������ҽʦ,����ҽʦ", gstrNodeNo)
            
            If Not rsTmp.RecordCount > 0 Then
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrUnit, CStr("" & mrsPatiInfo!����ҽʦ), "����ҽʦ,������ҽʦ,����ҽʦ", gstrNodeNo)
            End If
            If cbo����ҽʦ.ListIndex <> -1 Then
                lngҽʦ = cbo����ҽʦ.ItemData(cbo����ҽʦ.ListIndex)
            Else
                lngҽʦ = 0
            End If
            cbo����ҽʦ.Clear
            Do Until rsTmp.EOF
                cbo����ҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����
                cbo����ҽʦ.ItemData(cbo����ҽʦ.NewIndex) = rsTmp!ID
                rsTmp.MoveNext
            Loop
             '105133:������ҽʦ����ѡҽ��С��ʱ���ı�����ҽʦ
            If lngҽʦ <> 0 Then Call cbo.SetIndex(cbo����ҽʦ.hWnd, cbo.FindIndex(cbo����ҽʦ, lngҽʦ))
        End If
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrUnit, CStr("" & mrsPatiInfo!סԺҽʦ), "����ҽʦ,������ҽʦ,����ҽʦ,ҽʦ,ҽʿ", gstrNodeNo)
        cboסԺҽʦ.Clear
        Do Until rsTmp.EOF
            cboסԺҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����
            cboסԺҽʦ.ItemData(cboסԺҽʦ.NewIndex) = rsTmp!ID
            rsTmp.MoveNext
        Loop
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrUnit, CStr("" & mrsPatiInfo!����ҽʦ), "����ҽʦ,������ҽʦ,����ҽʦ", gstrNodeNo)
        cbo����ҽʦ.Clear
        Do Until rsTmp.EOF
            cbo����ҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����
            cbo����ҽʦ.ItemData(cbo����ҽʦ.NewIndex) = rsTmp!ID
            rsTmp.MoveNext
        Loop
    End If
    cboסԺҽʦ.AddItem "����..."
    cbo����ҽʦ.AddItem "����..."
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo����ҽʦ_Click()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    Dim strSQL As String
    On Error GoTo errHandle:
    
    If cbo����ҽʦ.Text = "����..." Then
        Set rsTmp = GetSelectPersonal("ҽ��", "����ҽʦ,������ҽʦ,����ҽʦ", Me)
        If Not rsTmp Is Nothing Then
            For i = 0 To cbo����ҽʦ.ListCount - 1
                If cbo����ҽʦ.List(i) = rsTmp!���� & "-" & rsTmp!���� Then
                    cbo����ҽʦ.ListIndex = i: Exit Sub
                End If
            Next
            cbo����ҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����, cbo����ҽʦ.ListCount - 1
            cbo����ҽʦ.ListIndex = cbo����ҽʦ.NewIndex
            cbo����ҽʦ.ItemData(cbo����ҽʦ.NewIndex) = rsTmp!ID
        Else
            cbo����ҽʦ.ListIndex = -1
        End If
    Else
        '����ҽʦѡ��ʱҽ��С����סԺҽʦ����
        '105133:�����������֮ǰ��Ӧ�ø���סԺҽʦ������ҽ��С��
        If cboҽ��С��.ListCount <= 1 Or Not Me.Visible Then Exit Sub
        strSQL = "Select ID,����,˵�� From �ٴ�ҽ��С�� A, ҽ��С����Ա B " & _
                "Where a.id=b.С��id And b.��Աid=[1] And a.����id=[2] And (����ʱ�� Is NULL Or Trunc(����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD')) Order By ID"
        If cboסԺҽʦ.ListIndex <> -1 Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val("" & cboסԺҽʦ.ItemData(cboסԺҽʦ.ListIndex)), Val("" & txt����.Tag))
            Do While Not rsTmp.EOF
                If cboҽ��С��.Text = Nvl(rsTmp!ID) & "-" & Nvl(rsTmp!����) Then Exit Sub
            rsTmp.MoveNext
            Loop
            If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                Call cbo.SetIndex(cboҽ��С��.hWnd, cbo.FindIndex(cboҽ��С��, Nvl(rsTmp!����), True))
                Exit Sub
            End If
        End If
        If cbo����ҽʦ.ListIndex <> -1 Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val("" & cbo����ҽʦ.ItemData(cbo����ҽʦ.ListIndex)), Val("" & txt����.Tag))
            Do While Not rsTmp.EOF
                If cboҽ��С��.Text = Nvl(rsTmp!ID) & "-" & Nvl(rsTmp!����) Then Exit Sub
            rsTmp.MoveNext
            Loop
            If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                Call cbo.SetIndex(cboҽ��С��.hWnd, cbo.FindIndex(cboҽ��С��, Nvl(rsTmp!����), True))
            Else
                Call cbo.SetIndex(cboҽ��С��.hWnd, cboҽ��С��.ListCount - 1)
            End If
        Else
            Call cbo.SetIndex(cboҽ��С��.hWnd, cboҽ��С��.ListCount - 1)
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboסԺҽʦ_Click()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    Dim strSQL As String
    
    On Error GoTo errHandle:
    
    If cboסԺҽʦ.Text = "����..." Then
        Set rsTmp = GetSelectPersonal("ҽ��", "����ҽʦ,������ҽʦ,����ҽʦ,ҽʦ,ҽʿ", Me)
        If Not rsTmp Is Nothing Then
            For i = 0 To cboסԺҽʦ.ListCount - 1
                If cboסԺҽʦ.List(i) = rsTmp!���� & "-" & rsTmp!���� Then
                    cboסԺҽʦ.ListIndex = i: Exit Sub
                End If
            Next
            cboסԺҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����, cboסԺҽʦ.ListCount - 1
            cboסԺҽʦ.ListIndex = cboסԺҽʦ.NewIndex
            cboסԺҽʦ.ItemData(cboסԺҽʦ.NewIndex) = rsTmp!�ϼ�ID
        Else
            cboסԺҽʦ.ListIndex = -1
        End If
    Else
        '105133:�����������֮ǰ��Ӧ�ø���סԺҽʦ������ҽ��С��
        If cboҽ��С��.ListCount <= 1 Or Not Me.Visible Then Exit Sub
        
        strSQL = "Select ID,����,˵�� From �ٴ�ҽ��С�� A, ҽ��С����Ա B " & _
                "Where a.id=b.С��id And b.��Աid=[1] And a.����id=[2] And (����ʱ�� Is NULL Or Trunc(����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD')) Order By ID"
        If cboסԺҽʦ.ListIndex <> -1 Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val("" & cboסԺҽʦ.ItemData(cboסԺҽʦ.ListIndex)), Val("" & txt����.Tag))
            Do While Not rsTmp.EOF
                If cboҽ��С��.Text = Nvl(rsTmp!ID) & "-" & Nvl(rsTmp!����) Then Exit Sub
            rsTmp.MoveNext
            Loop
            If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                Call cbo.SetIndex(cboҽ��С��.hWnd, cbo.FindIndex(cboҽ��С��, Nvl(rsTmp!����), True))
                Exit Sub
            End If
        End If
        If cbo����ҽʦ.ListIndex <> -1 Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val("" & cbo����ҽʦ.ItemData(cbo����ҽʦ.ListIndex)), Val("" & txt����.Tag))
            Do While Not rsTmp.EOF
                If cboҽ��С��.Text = Nvl(rsTmp!ID) & "-" & Nvl(rsTmp!����) Then Exit Sub
            rsTmp.MoveNext
            Loop
            If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                Call cbo.SetIndex(cboҽ��С��.hWnd, cbo.FindIndex(cboҽ��С��, Nvl(rsTmp!����), True))
            Else
                Call cbo.SetIndex(cboҽ��С��.hWnd, cboҽ��С��.ListCount - 1)
            End If
        Else
            Call cbo.SetIndex(cboҽ��С��.hWnd, cboҽ��С��.ListCount - 1)
        End If
    End If
    Exit Sub
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboסԺҽʦ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cboסԺҽʦ.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = cbo.MatchIndex(cboסԺҽʦ.hWnd, KeyAscii, 0.5)
    If lngIdx <> -2 Then cboסԺҽʦ.ListIndex = lngIdx
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strSQLҽ��С�� As String
    Dim str���� As String
    Dim i As Integer, lngLevel As Long
    
    
    gblnOK = False
    
    strSQL = "Select NVl(A.����,D.����) ����,NVL(A.�Ա�,D.�Ա�) �Ա�, A.����,To_Char(A.��Ժ����, 'YYYY-MM-DD HH24:MI:SS') As ��Ժʱ��,E.���� as ��ǰ����,A.��Ժ����id as ��ǰ����ID,H.���� ��ǰ����,A.��ǰ����Id, A.ҽ��С��id, g.���� as ҽ��С��, " & vbNewLine & _
            "A.סԺ��,A.���λ�ʿ, A.����ҽʦ, A.סԺҽʦ, B.��Ϣֵ ����ҽʦ, C.��Ϣֵ ����ҽʦ, A.�ѱ�, A.����״��, A.ѧ��," & vbNewLine & _
            "       A.ְҵ, A.��ǰ����, A.��λ��ַ, A.��λ�ʱ�, A.��λ�绰, A.��ͥ��ַ, A.��ͥ�绰, A.��ϵ�˵�ַ," & vbNewLine & _
            "       A.��ϵ�˵绰, A.��ϵ������, A.��ϵ�˹�ϵ, A.����Ժ, A.��������, A.����, D.���֤��, D.����, D.�����ص�," & vbNewLine & _
            "       D.��������, A.��Ժ����,D.��ͬ��λid, F.���� As ����ȼ�,Nvl(A.��������,Decode(A.����,Null,'��ͨ����','ҽ������')) ��������,A.��Ժ��ʽ,A.��ע,A.�Ƿ����" & vbNewLine & _
            "From ������ҳ A, ������ҳ�ӱ� B, ������ҳ�ӱ� C, ������Ϣ D,���ű� E,���ű� H,�շ���ĿĿ¼ F, �ٴ�ҽ��С�� G " & vbNewLine & _
            "Where A.����id = [1] And A.��ҳid = [2] And A.����id = B.����id(+) And A.��ҳid = B.��ҳid(+) And A.����id = C.����id(+) And" & vbNewLine & _
            "      A.��ҳid = C.��ҳid(+) And A.ҽ��С��id = G.id(+) And B.��Ϣ��(+) = '����ҽʦ' And C.��Ϣ��(+) = '����ҽʦ' And A.����id = D.����id And A.��Ժ����id = E.id And A.��ǰ����Id=h.id(+)" & vbNewLine & _
            " And A.����ȼ�id = F.ID(+)"
    Set mrsPatiInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    
    gstrSQL = "Select ���� as ҽ��С�� From �ٴ�ҽ��С�� Where ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val("" & mrsPatiInfo!ҽ��С��id))
    If Not rsTmp.EOF Then
        txtPre.Text = Nvl(rsTmp!ҽ��С��)
    End If
    Set rsTmp = GetPatiBeds(mlng����ID)
    If rsTmp.RecordCount = 0 Then
        str���� = "��ͥ����"
    Else
        Do While Not rsTmp.EOF
            str���� = str���� & "," & rsTmp!����
            rsTmp.MoveNext
        Loop
        str���� = Mid(str����, 2)
    End If
    txt����.Text = str����
    
    With mrsPatiInfo
       txt����.Text = !����
       txt�Ա�.Text = "" & !�Ա�
       txt����.Text = "" & !����
       txtסԺ��.Text = "" & !סԺ��
       txt����.Text = "" & !��ǰ����
       txt����.Tag = "" & !��ǰ����id
    End With

    txtDate.Text = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    mstrUnit = Get����IDs(mlngUnit) & "," & mlngUnit
    
    '��ʼ��ҽ��С��
    strSQL = "Select ID,����,˵��,����ʱ��,����ʱ�� From �ٴ�ҽ��С�� Where ����id=[1] " & _
            " And (����ʱ�� Is NULL Or Trunc(����ʱ��) = To_Date('3000-01-01','YYYY-MM-DD')) Order By Id "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val("" & txt����.Tag))
    
    If Not rsTmp.EOF Then
        cboҽ��С��.Clear
        Do Until rsTmp.EOF
            cboҽ��С��.AddItem rsTmp!ID & "-" & rsTmp!����
            cboҽ��С��.ItemData(cboҽ��С��.NewIndex) = rsTmp!ID
            rsTmp.MoveNext
        Loop
        cboҽ��С��.AddItem "": cboҽ��С��.ItemData(cboҽ��С��.NewIndex) = 0: cboҽ��С��.ListIndex = -1
    Else
        MsgBox "�ÿ���δ����ҽ��С��,���ȵ����ٴ�ҽ��С�顿�����ã�", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    '��ʼ��סԺҽʦ������ҽʦ
'    strSQLҽ��С�� = "Select Distinct A.ID, A.���, A.����, A.����" & vbNewLine & _
'                    " From ��Ա�� A, ��Ա����˵�� B, ������Ա C, ҽ��С����Ա D" & vbNewLine & _
'                    " Where A.ID = B.��Աid And A.ID = C.��Աid And a.id = d.��Աid And B.��Ա���� = 'ҽ��' And d.С��id = [1] And" & vbNewLine & _
'                    "   (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) And" & vbNewLine & _
'                    "   (Instr(',' || [2] || ',', ',' || C.����id || ',') > 0 Or a.����=[3]) And Instr(',' || [4] || ',', ',' || A.רҵ����ְ�� || ',') > 0" & vbNewLine & _
'                    "   And (A.վ��=[5] Or A.վ�� is Null)" & vbNewLine & _
'                    " Order By A.����"
    strSQL = "Select Distinct A.ID, A.���, A.����, A.����" & vbNewLine & _
                        " From ��Ա�� A, ��Ա����˵�� B, ������Ա C" & vbNewLine & _
                        " Where A.ID = B.��Աid And A.ID = C.��Աid And B.��Ա���� = 'ҽ��' And" & vbNewLine & _
                        "      (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) And" & vbNewLine & _
                        "      (Instr(',' || [1] || ',', ',' || C.����id || ',') > 0 Or A.����=[2]) And Instr(',' || [3] || ',', ',' || A.רҵ����ְ�� || ',') > 0" & vbNewLine & _
                        "      And (A.վ��=[4] Or A.վ�� is Null)" & _
                        " Order By A.����"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrUnit & "," & mlngUnit, CStr("" & mrsPatiInfo!סԺҽʦ), "����ҽʦ,������ҽʦ,����ҽʦ,ҽʦ,ҽʿ", gstrNodeNo)
    cboסԺҽʦ.Clear
    Do Until rsTmp.EOF
        cboסԺҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����
        cboסԺҽʦ.ItemData(cboסԺҽʦ.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Loop
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrUnit & "," & mlngUnit, CStr("" & mrsPatiInfo!����ҽʦ), "����ҽʦ,������ҽʦ,����ҽʦ", gstrNodeNo)
    cbo����ҽʦ.Clear
    Do Until rsTmp.EOF
        cbo����ҽʦ.AddItem rsTmp!���� & "-" & rsTmp!����
        cbo����ҽʦ.ItemData(cbo����ҽʦ.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Loop
    
    cboסԺҽʦ.AddItem "����..."
    cbo����ҽʦ.AddItem "����..."
    
    '��궨λ
    cboҽ��С��.ListIndex = cbo.FindIndex(cboҽ��С��, IIf(IsNull(mrsPatiInfo!ҽ��С��), "", mrsPatiInfo!ҽ��С��), True)
    cboסԺҽʦ.ListIndex = cbo.FindIndex(cboסԺҽʦ, IIf(IsNull(mrsPatiInfo!סԺҽʦ), "", mrsPatiInfo!סԺҽʦ), True)
    cbo����ҽʦ.ListIndex = cbo.FindIndex(cbo����ҽʦ, IIf(IsNull(mrsPatiInfo!����ҽʦ), "", mrsPatiInfo!����ҽʦ), True)
    
    '������Ϣ����
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, 1132, mstrPirvs, mfrmParent.hWnd)
    Call AddMipModule(mclsMipModule)
    Set mclsXML = New zl9ComLib.clsXML
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'ж����Ϣ����
    If Not (mclsMipModule Is Nothing) Then
        Call mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    If Not (mclsXML Is Nothing) Then
        Set mclsXML = Nothing
    End If
End Sub

Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_LostFocus()
    If Not IsDate(txtDate.Text) Then txtDate.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim dMax As Date, strSQL As String
    Dim Curdate As Date
    Dim blnTrue As Boolean
    
    If cboҽ��С��.ListIndex = -1 Then
        MsgBox "��ѡ���µ�ҽ��С�飡", vbInformation, gstrSysName
        cboҽ��С��.SetFocus: Exit Sub
    End If
    
    blnTrue = (Val(zlDatabase.GetPara("��סָ��ҽ��С��", glngSys, glngModul, 0)) = 1)
    If cboҽ��С��.ItemData(cboҽ��С��.ListIndex) = 0 And blnTrue = True Then
        MsgBox "���ڹ�ѡ�˲���[��סָ��ҽ��С��],����ѡ��һ��ҽ��С�飬��ѡ��", vbInformation, gstrSysName
        If cboҽ��С��.Enabled And cboҽ��С��.Visible Then cboҽ��С��.SetFocus
        Exit Sub
    End If
    
    If cboסԺҽʦ.ListIndex = -1 Then
        MsgBox "��ѡ��סԺҽʦ��", vbInformation, gstrSysName
        cboҽ��С��.SetFocus: Exit Sub
    End If
    
    If cbo����ҽʦ.ListIndex = -1 Then
        MsgBox "��ѡ������ҽʦ��", vbInformation, gstrSysName
        cboҽ��С��.SetFocus: Exit Sub
    End If
    
    If Not IsDate(txtDate.Text) Then
        MsgBox "������Ϸ�����Чʱ�䣡", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    
    dMax = GetMaxDate(mlng����ID, mlng��ҳID)
    If CDate(txtDate.Text) <= dMax Then
        MsgBox "��Чʱ�������ڸò����ϴα䶯ʱ�� " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & " ��", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    
    'ʱ�䲻�ܳ�����ǰʱ��̫��(һ����)
    Curdate = zlDatabase.Currentdate
    If CDate(txtDate.Text) > Curdate Then
        If CDate(txtDate.Text) - Curdate > 30 Then
            MsgBox "��Чʱ��ȵ�ǰʱ���ù���,���飡", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        If MsgBox("��Чʱ������˵�ǰϵͳʱ��,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txtDate.SetFocus: Exit Sub
        End If
    End If
        
    strSQL = "zl_���˱䶯��¼_ChangeGroup(" & mlng����ID & "," & mlng��ҳID & "," & _
            cboҽ��С��.ItemData(cboҽ��С��.ListIndex) & ",'" & zlCommFun.GetNeedName(cboסԺҽʦ.Text) & "','" & zlCommFun.GetNeedName(cbo����ҽʦ.Text) & "',To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS')," & _
            "'" & UserInfo.��� & "','" & UserInfo.���� & "')"
    
    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    gblnOK = True
    
    On Error Resume Next
     'סԺҽʦ�䶯�󴥷���Ϣ
    If zlCommFun.GetNeedName(cboסԺҽʦ.Text) <> Nvl(mrsPatiInfo!סԺҽʦ) Or zlCommFun.GetNeedName(cbo����ҽʦ.Text) <> Nvl(mrsPatiInfo!����ҽʦ) Then
        If mclsMipModule.IsConnect = True Then
            mclsXML.ClearXmlText '��������е�XML
            '--������Ϣ��װ
            '������Ϣ
            mclsXML.AppendNode "in_patient"
            'patient_id      ����id  1   N
            mclsXML.appendData "patient_id", mlng����ID, xsNumber  '����ID
            'page_id     ��ҳid  1   N
            mclsXML.appendData "page_id", mlng��ҳID, xsNumber '��ҳID
            'patient_name        ����    1   S
            mclsXML.appendData "patient_name", txt����.Text, xsString '����
            'patient_sex     �Ա�    0..1    S
            mclsXML.appendData "patient_sex", txt�Ա�.Text, xsString '�Ա�
            'in_number       סԺ��  1   S
            mclsXML.appendData "in_number", txtסԺ��.Text, xsString  'סԺ��
            mclsXML.AppendNode "in_patient", True
            
            '��ǰ���
            'current_state       ��ǰ���    1
            mclsXML.AppendNode "current_state"
            'current_area_id     ��ǰ����id  0..1    N
            mclsXML.appendData "current_area_id", Val(Nvl(mrsPatiInfo!��ǰ����ID)), xsNumber
            'current_area_title      ��ǰ����    0..1    S
            mclsXML.appendData "current_area_title", Nvl(mrsPatiInfo!��ǰ����), xsString
            'current_dept_id     ��ǰ����id  1   N
            mclsXML.appendData "current_dept_id", Val(txt����.Tag), xsNumber
            'current_dept_title      ��ǰ����    1   S
            mclsXML.appendData "current_dept_title", txt����.Text, xsString
            'curren_in_doctor        סԺҽʦ    1   S
            mclsXML.appendData "curren_in_doctor", Nvl(mrsPatiInfo!סԺҽʦ), xsString
            'curren_director_doctor      ����ҽʦ    1   S
            mclsXML.appendData "curren_director_doctor", Nvl(mrsPatiInfo!����ҽʦ), xsString
            'curren_treat_doctor     ����ҽʦ    1   S
            mclsXML.appendData "curren_treat_doctor", Nvl(mrsPatiInfo!����ҽʦ), xsString
            'curren_duty_nurse       ���λ�ʿ    1   S
            mclsXML.appendData "curren_duty_nurse", Nvl(mrsPatiInfo!���λ�ʿ), xsString
            mclsXML.AppendNode "current_state", True
            
            strSQL = "Select ID �䶯id,��ʼʱ�� �䶯ʱ�� From ���˱䶯��¼ Where ����ID=[1] And ��ҳId=[2] And ��ʼԭ��=[3] And ��ʼʱ��=[4] And NVL(���Ӵ�λ,0)=0"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���˱䶯��¼", mlng����ID, mlng��ҳID, 14, CDate(Format(txtDate.Text, "YYYY-MM-DD HH:mm:ss")))
            '�����Ϣ
            'change_state        �����Ϣ    1
            mclsXML.AppendNode "change_state"
            'change_id       ���id  1   N
            mclsXML.appendData "change_id", rsTmp!�䶯ID, xsNumber
            'change_date     ���ʱ��    1   S
            mclsXML.appendData "change_date", Format(Nvl(rsTmp!�䶯ʱ��), "YYYY-MM-DD HH:mm:ss"), xsString
            'change_in_doctor        סԺҽʦ    1   S
            mclsXML.appendData "change_in_doctor", zlCommFun.GetNeedName(cboסԺҽʦ.Text), xsString
            'change_director_doctor      ����ҽʦ    1   S
            mclsXML.appendData "change_director_doctor", Nvl(mrsPatiInfo!����ҽʦ), xsString
            'change_treat_doctor     ����ҽʦ    1   S
            mclsXML.appendData "change_treat_doctor", zlCommFun.GetNeedName(cbo����ҽʦ.Text), xsString
            'change_duty_nurse       ���λ�ʿ    1   S
            mclsXML.appendData "change_duty_nurse", Nvl(mrsPatiInfo!���λ�ʿ), xsString
            'change_operator         ����Ա      1   S
            mclsXML.appendData "change_operator", UserInfo.����, xsString
            mclsXML.AppendNode "change_state", True
    
            mclsMipModule.CommitMessage "ZLHIS_PATIENT_007", mclsXML.XmlText
        End If
    End If
    
    If Err <> 0 Then Err.Clear
    
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function ShowMe(frmParent As Object, ByVal lngUnit As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strPrivs As String) As Boolean
    On Error Resume Next
    Set mfrmParent = frmParent
    mstrPirvs = strPrivs
    mlngUnit = lngUnit
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    
    
    Me.Show 1, frmParent
    ShowMe = gblnOK
End Function
