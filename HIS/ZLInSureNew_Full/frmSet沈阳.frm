VERSION 5.00
Begin VB.Form frmSet���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���в�������"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   Icon            =   "frmSet����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cbo���õ��� 
      Height          =   300
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   4140
      Width           =   1785
   End
   Begin VB.CheckBox chk�Ƿ�������Ժ���˽�������ҵ�� 
      Caption         =   "�Ƿ�������Ժ���˽�������ҵ��"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   3810
      Width           =   2895
   End
   Begin VB.CheckBox chk�Һ� 
      Alignment       =   1  'Right Justify
      Caption         =   "�Һ�(&R)"
      Height          =   195
      Left            =   300
      TabIndex        =   13
      Top             =   2940
      Width           =   945
   End
   Begin VB.Frame fra�Һ� 
      Enabled         =   0   'False
      Height          =   705
      Left            =   210
      TabIndex        =   14
      Top             =   2970
      Width           =   3165
      Begin VB.ComboBox cbo������Ŀ 
         Height          =   300
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label lbl�����ʻ�֧�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����ʻ�֧��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   210
         TabIndex        =   15
         Top             =   300
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3540
      TabIndex        =   21
      Top             =   840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3540
      TabIndex        =   20
      Top             =   360
      Width           =   1100
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   3
      Left            =   2220
      MaxLength       =   2
      TabIndex        =   12
      Top             =   2640
      Width           =   855
   End
   Begin VB.Frame fraҽ�������� 
      Caption         =   "ҽ��������"
      Height          =   2385
      Left            =   210
      TabIndex        =   0
      Top             =   120
      Width           =   3165
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   1260
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   1890
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   8
         Top             =   1500
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   2
         Top             =   330
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   4
         Top             =   720
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   6
         Top             =   1110
         Width           =   1635
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����(&W)"
         Height          =   180
         Index           =   5
         Left            =   570
         TabIndex        =   9
         Top             =   1950
         Width           =   630
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����Ա(&U)"
         Height          =   180
         Index           =   4
         Left            =   390
         TabIndex        =   7
         Top             =   1560
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��ַ(&A)"
         Height          =   180
         Index           =   0
         Left            =   570
         TabIndex        =   1
         Top             =   390
         Width           =   630
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�˿ں�(&P)"
         Height          =   180
         Index           =   1
         Left            =   390
         TabIndex        =   3
         Top             =   780
         Width           =   810
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��ڳ���(&S)"
         Height          =   180
         Index           =   2
         Left            =   210
         TabIndex        =   5
         Top             =   1170
         Width           =   990
      End
   End
   Begin VB.Label lbl���õ��� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���õ���(&K)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   18
      Top             =   4200
      Width           =   990
   End
   Begin VB.Label lblEdit 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "IC����д�豸�˿ں�(&I)"
      Height          =   180
      Index           =   3
      Left            =   270
      TabIndex        =   11
      Top             =   2700
      Width           =   1890
   End
End
Attribute VB_Name = "frmSet����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Modified By ���� ��������ɳ ԭ�������������˵�¼����Ա����¼����
Private Enum ����
    ��ַ = 0
    �˿ں�
    ��ڳ���
    IC�豸�˿�
    ��¼����Ա
    ��¼����
End Enum
Private mblnReturn As Boolean

Private Sub chk�Һ�_Click()
    On Error Resume Next
    fra�Һ�.Enabled = (chk�Һ�.Value = 1)
    If fra�Һ�.Enabled Then
        cbo������Ŀ.SetFocus
    Else
        cbo������Ŀ.ListIndex = 0
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHand
    
    If Not Valid Then Exit Sub
    gcnOracle.BeginTrans
    
    'ɾ���Ѿ�����
    gstrSQL = "zl_���ղ���_Delete(" & TYPE_������ & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '������������
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_������ & ",NULL,'��������ַ','" & txtEdit(����.��ַ).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_������ & ",NULL,'�������˿ں�','" & txtEdit(����.�˿ں�).Text & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_������ & ",NULL,'��������ڳ���','" & txtEdit(����.��ڳ���).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_������ & ",NULL,'��¼����Ա','" & txtEdit(����.��¼����Ա).Text & "',4)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_������ & ",NULL,'��¼����','" & txtEdit(����.��¼����).Text & "',5)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_������ & ",NULL,'�����ʻ�֧��(�Һ�)','" & cbo������Ŀ.ItemData(cbo������Ŀ.ListIndex) & "',6)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_������ & ",NULL,'��������ҵ��','" & chk�Ƿ�������Ժ���˽�������ҵ��.Value & "',7)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
'    gstrSQL = "zl_���ղ���_Insert(" & TYPE_������ & ",NULL,'��������ҩ','" & chk������ҩ.Value & "',8)"
'    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_������ & ",NULL,'���õ���','" & cbo���õ���.ListIndex & "',9)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gcnOracle.CommitTrans
    mblnReturn = True
    
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName, "IC�豸�˿�", txtEdit(����.IC�豸�˿�).Text)
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Function Valid() As Boolean
    Dim strPara As String, arrPara
    Dim intDO As Integer, intUbound As Integer
    '����Ƿ��������Ĳ�����������
    strPara = "��������ַ||�������˿ں�||��������ڳ���||IC�豸�˿ں�||��¼����Ա||����"
    arrPara = Split(strPara, "||")
    
    intUbound = txtEdit.Count - 1
    For intDO = 0 To intUbound
        If Trim(txtEdit(intDO)) = "" Then
            MsgBox arrPara(intDO) & "����Ϊ�գ�", vbInformation, gstrSysName
            txtEdit(intDO).SetFocus
            Exit Function
        End If
    Next
    
    Valid = True
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim intDO As Integer, blnFind As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    '��ȡ������Ŀ
    gstrSQL = "Select ID,���� From ������Ŀ Order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ŀ")
    
    cbo������Ŀ.Clear
    cbo������Ŀ.AddItem ""
    Do While Not rsTemp.EOF
        cbo������Ŀ.AddItem Nvl(rsTemp!����)
        cbo������Ŀ.ItemData(cbo������Ŀ.NewIndex) = rsTemp!ID
        rsTemp.MoveNext
    Loop
    cbo������Ŀ.ListIndex = 0
    
    With cbo���õ���
        .Clear
        .AddItem "��������"
        .AddItem "��������"
        .AddItem "��������"
        .ListIndex = 0
    End With
    
    '��ȡ��������ַ���˿ڼ��������('��������ַ','�������˿ں�','��������ڳ���')
    gstrSQL = " Select ������,����ֵ From ���ղ���" & _
              " Where ����=[1] And ������ Like '������%'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������ַ���˿ڼ��������", TYPE_������)
    
    With rsTemp
        Do While Not .EOF
            Select Case !������
            Case "��������ַ"
                txtEdit(����.��ַ).Text = Nvl(!����ֵ)
            Case "�������˿ں�"
                txtEdit(����.�˿ں�).Text = Nvl(!����ֵ)
            Case "��������ڳ���"
                txtEdit(����.��ڳ���).Text = Nvl(!����ֵ)
            Case "��¼����Ա"
                txtEdit(����.��¼����Ա).Text = Nvl(!����ֵ)
            Case "��¼����"
                txtEdit(����.��¼����).Text = Nvl(!����ֵ)
            Case "�����ʻ�֧��(�Һ�)"
                For intDO = 1 To cbo������Ŀ.ListCount
                    cbo������Ŀ.ListIndex = intDO - 1
                    If cbo������Ŀ.ItemData(cbo������Ŀ.ListIndex) = Nvl(!����ֵ, 0) Then
                        blnFind = True
                        Exit For
                    End If
                    If Not blnFind Then cbo������Ŀ.ListIndex = 0
                Next
            Case "��������ҵ��"
                chk�Ƿ�������Ժ���˽�������ҵ��.Value = Nvl(!����ֵ, 0)
'            Case "��������ҩ"
'                chk������ҩ.Value = NVL(!����ֵ, 0)
            Case "���õ���"
                cbo���õ���.ListIndex = Nvl(!����ֵ, 0)
            End Select
            .MoveNext
        Loop
    End With
    txtEdit(����.IC�豸�˿�).Text = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "IC�豸�˿�", 1)
    
    If cbo������Ŀ.ItemData(cbo������Ŀ.ListIndex) <> 0 Then chk�Һ�.Value = 1
End Sub

Public Function ShowME() As Boolean
    mblnReturn = False
    Me.Show 1
    ShowME = mblnReturn
End Function
