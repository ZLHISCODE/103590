VERSION 5.00
Begin VB.Form frmStPathEdit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��׼·���޸�"
   ClientHeight    =   6105
   ClientLeft      =   8400
   ClientTop       =   4605
   ClientWidth     =   5415
   Icon            =   "frmStPathEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboType 
      Enabled         =   0   'False
      Height          =   300
      ItemData        =   "frmStPathEdit.frx":076A
      Left            =   3600
      List            =   "frmStPathEdit.frx":0774
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   975
      Width           =   1695
   End
   Begin VB.TextBox txtPathName 
      Height          =   300
      Left            =   1200
      MaxLength       =   80
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "ѡ������(&M)"
      Height          =   350
      Index           =   1
      Left            =   3945
      TabIndex        =   11
      Top             =   3240
      Width           =   1350
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "ѡ�񼲲�(&D)"
      Height          =   350
      Index           =   0
      Left            =   3945
      TabIndex        =   8
      Top             =   1320
      Width           =   1350
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   5415
      TabIndex        =   17
      Top             =   5490
      Width           =   5415
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   2760
         TabIndex        =   14
         Top             =   160
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   4200
         TabIndex        =   15
         Top             =   160
         Width           =   1100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   10000
         Y1              =   45
         Y2              =   45
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   10000
         Y1              =   30
         Y2              =   30
      End
   End
   Begin VB.TextBox txtSuitCode 
      Height          =   1335
      Index           =   1
      Left            =   120
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   3720
      Width           =   5175
   End
   Begin VB.TextBox txtSuitCode 
      Height          =   1335
      Index           =   0
      Left            =   120
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   1800
      Width           =   5175
   End
   Begin VB.ComboBox cboVersion 
      Height          =   300
      Left            =   3600
      TabIndex        =   5
      Top             =   520
      Width           =   1695
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   7
      Top             =   980
      Width           =   1695
   End
   Begin VB.ComboBox cboDept 
      Height          =   300
      ItemData        =   "frmStPathEdit.frx":0788
      Left            =   1200
      List            =   "frmStPathEdit.frx":078A
      TabIndex        =   3
      Top             =   520
      Width           =   1695
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      Caption         =   "����(&T)"
      Height          =   180
      Left            =   3000
      TabIndex        =   19
      Top             =   1035
      Width           =   630
   End
   Begin VB.Label lblPathName 
      AutoSize        =   -1  'True
      Caption         =   "·������(&N)"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   990
   End
   Begin VB.Label lblAttention 
      Caption         =   "��ʾ�������Ŀ�Զ��ŷָ�"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   5160
      Width           =   5175
   End
   Begin VB.Label lblOprCode 
      Caption         =   "��������(ICD-9-CM3��������)(&G)"
      Height          =   180
      Left            =   120
      TabIndex        =   12
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label lblDiseaseCode 
      Caption         =   "���ü���(ICD-10��������)(&F)"
      Height          =   180
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label lblCode 
      Caption         =   "��    ��(&S)"
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   1040
      Width           =   990
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "�汾(&V)"
      Height          =   180
      Left            =   3000
      TabIndex        =   4
      Top             =   585
      Width           =   630
   End
   Begin VB.Label lblDept 
      AutoSize        =   -1  'True
      Caption         =   "��������(&K)"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   580
      Width           =   990
   End
End
Attribute VB_Name = "frmStPathEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintMode As Integer  '�������޸ģ�0-������1-�޸�
Private mlngStPathID As Long 'Ҫ�޸ĵı�׼·��ID
Private mblnOK As Boolean
Private mstr�������� As String
Private mstr�汾     As String
Private mstr·������ As String
Private mstr����     As String
Private mstr��������s As String
Private mstr��������s As String
Private mbytType As Byte

Private Enum LenMax '�ֶ������ݿ��еĳ���
    LM_���� = 8
    LM_�������� = 100
    LM_·������ = 80
    LM_�汾˵�� = 20
    LM_�������� = 100
    LM_�������� = 100
End Enum
Public Function ShowMe(FrmParent As Object, ByVal intMode As Integer, Optional ByRef lngStPathID As Long, Optional ByVal str·������ As String, _
            Optional ByVal str���� As String, Optional ByVal str�������� As String, Optional ByVal str�汾 As String, _
                Optional ByVal str��������s As String, Optional ByVal str��������s As String, Optional ByVal bytType As Byte) As Boolean
'˵����·��ά�������еĸ��±�׼·������ӱ�׼·��ʱ����
'   intMode:'0-������׼·��
'            1-�޸ĸ��±�׼·�����׼·����Ӧ��������
'   lngStPathID,���±�׼·��ʱ����,��������׼·��ʱ���������ı�׼·��ID
'   str·������,str����,str��������,str�汾,str��������s,str��������s:���±�׼·��ʱ����
'   bytType��0-��ҽ,1-��ҽ
    
    mintMode = intMode
    mlngStPathID = lngStPathID
    mstr·������ = str·������
    mstr���� = str����
    mstr�������� = str��������
    mstr�汾 = str�汾
    mstr��������s = str��������s
    mstr��������s = str��������s
    mbytType = bytType
    
    Me.Show 1, FrmParent
    ShowMe = mblnOK
    lngStPathID = mlngStPathID
End Function


Private Sub cboVersion_Change()
    Call CheckInput(False)
End Sub

Private Sub cboDept_Change()
    Call CheckInput(False)
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub LoadCboData()
'���ܣ����������б�����
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSql = "Select  ��������,Rownum ID  From (Select Distinct �������� From ��׼·��Ŀ¼ where NVl(���,0)=[1] Order By ��������)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mbytType)
    Call zlControl.CboAddData(cboDept, rsTmp, True)
    strSql = "Select �汾˵��,Rownum ID From (Select Distinct �汾˵�� From ��׼·��Ŀ¼ where NVl(���,0)=[1] Order By �汾˵��)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mbytType)
    Call zlControl.CboAddData(cboVersion, rsTmp, True)
    '����
    Call zlControl.CboLocate(cboType, mbytType, True)
    
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Load����()
'���ܣ��������������뼲������
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSql = "Select ��������, �������� From ��׼·������ Where ��׼·��id =[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngStPathID)
    If rsTmp.RecordCount > 0 Then
        txtSuitCode(0).Text = Nvl(rsTmp!��������)
        txtSuitCode(1).Text = Nvl(rsTmp!��������)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    If Not SaveData Then Exit Sub
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdSel_Click(Index As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim str����s As String
    
    On Error GoTo errH
    'D:ICD-10�������� S:ICD-9-CM3�������� B:��ҽ��������
    Set rsTmp = zlDatabase.ShowILLSelect(Me, IIf(Index = 0, IIf(mbytType = 0, "D", "B,D"), "S"), 0, , True, , IIf(Trim(txtSuitCode(Index).Text) = "", "", "," & txtSuitCode(Index).Text & ","))
    
    If rsTmp Is Nothing Then Exit Sub
    
    If rsTmp.RecordCount <> 0 Then
            str����s = ""
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                str����s = str����s & "," & rsTmp!����
                rsTmp.MoveNext
            Loop
            str����s = Mid(str����s, 2)
            txtSuitCode(Index).Text = str����s
    End If
    Call CheckInput(False)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
'���ܣ���ʼ����������
    Call LoadCboData
    
    If mintMode = 0 Then
        Me.Caption = "������׼·��"
        cboDept.Text = mstr��������
    Else
        Call Load����
        Me.Caption = "�޸ı�׼·��"
        cboVersion.Text = mstr�汾
        cboDept.Text = mstr��������
        txtPathName.Text = mstr·������
        txtCode.Text = mstr����
        txtSuitCode(0).Text = mstr��������s
        txtSuitCode(1).Text = mstr��������s
    End If
    
    If mbytType = 0 Then
        lblDiseaseCode.Caption = "���ü���(ICD-10��������)(&F)"
    Else
        lblDiseaseCode.Caption = "���ü���(TCD�����ICD-10��������)(&F)"
    End If

End Sub



Private Function SaveData() As Boolean
'���ܣ����ݺ����Լ�鲢����
    Dim rsTmp As ADODB.Recordset
    Dim str��������s As String, str��������s As String
    Dim strSql As String
    
    If CheckInput(True) = False Then
        SaveData = False: Exit Function
    End If
    
    str��������s = Replace(Trim(txtSuitCode(0).Text), "��", ",")
    str��������s = Replace(Trim(txtSuitCode(1).Text), "��", ",")
    
    On Error GoTo errH
    '����
    If mintMode = 0 Then
        '��ȡ����·����ID
        strSql = "Select ��׼·��Ŀ¼_Id.Nextval ID From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        mlngStPathID = rsTmp!ID
        
        strSql = "Zl_��׼·��Ŀ¼_Insert(" & mlngStPathID & ",'" & Trim(cboDept.Text) & "','" & Trim(txtCode.Text) & "','" & Trim(txtPathName.Text) & "','" & Trim(cboVersion.Text) & "','" & _
                str��������s & "','" & str��������s & "'," & mbytType & ")"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    Else '�޸�
        strSql = "Zl_��׼·��Ŀ¼_Update(" & mlngStPathID & ",'" & Trim(cboDept.Text) & "','" & Trim(txtCode.Text) & "','" & Trim(txtPathName.Text) & "','" & Trim(cboVersion.Text) & "','" & _
                str��������s & "','" & str��������s & "')"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    End If
    SaveData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckInput(ByVal blnCheckNull As Boolean) As Boolean
'���ܣ���������Ϸ��Լ��
'������blnCheckNull���Ƿ���п�ֵ����

    Dim strMsg As String
    '��ֵ���
    If blnCheckNull Then
        If Trim(txtPathName.Text) = "" Then
            MsgBox "����δ�����׼·������", vbInformation, gstrSysName
            txtPathName.SetFocus: Exit Function
        End If
        If Trim(txtCode.Text) = "" Then
            MsgBox "����δ�����׼·������", vbInformation, gstrSysName
            txtCode.SetFocus: Exit Function
        End If
        If Trim(cboVersion.Text) = "" Then
            MsgBox "����δ�����׼·������", vbInformation, gstrSysName
            cboVersion.SetFocus: Exit Function
        End If
        If Trim(cboDept.Text) = "" Then
            MsgBox "����δ�����׼·������", vbInformation, gstrSysName
            cboDept.SetFocus: Exit Function
        End If
    End If
    '���ȼ��
    If LenB(StrConv(txtPathName.Text, vbFromUnicode)) > LM_·������ Then
        strMsg = "�������·�����Ƴ������������󳤶�" & LM_·������ & "(" & LM_·������ \ 2 & "�����ĵĳ��Ȼ�" & LM_·������ & "����ĸ�����ֵĳ���)"
        MsgBox strMsg, vbInformation, gstrSysName
        txtPathName.SetFocus
        CheckInput = False
        Exit Function
    End If
    
    If LenB(StrConv(cboDept.Text, vbFromUnicode)) > LM_�������� Then
        strMsg = "������Ŀ������Ƴ������������󳤶�" & LM_�������� & "(" & LM_�������� \ 2 & "�����ĵĳ��Ȼ�" & LM_�������� & "����ĸ�����ֵĳ���)"
        MsgBox strMsg, vbInformation, gstrSysName
        cboDept.SetFocus
        CheckInput = False
        Exit Function
    End If
    
    If LenB(StrConv(cboVersion.Text, vbFromUnicode)) > LM_�汾˵�� Then
        strMsg = "������İ汾�������������󳤶�" & LM_�汾˵�� & "(" & LM_�汾˵�� \ 2 & "�����ĵĳ��Ȼ�" & LM_�汾˵�� & "����ĸ�����ֵĳ���)"
        MsgBox strMsg, vbInformation, gstrSysName
        cboVersion.SetFocus
        CheckInput = False
        Exit Function
    End If
    
    If LenB(StrConv(txtCode.Text, vbFromUnicode)) > LM_���� Then
        strMsg = "������ı��볬�����������󳤶�" & LM_���� & "(" & LM_���� \ 2 & "�����ĵĳ��Ȼ�" & LM_���� & "����ĸ�����ֵĳ���)"
        MsgBox strMsg, vbInformation, gstrSysName
        txtCode.SetFocus
        CheckInput = False
        Exit Function
    End If
    
    If LenB(StrConv(txtSuitCode(0).Text, vbFromUnicode)) > LM_�������� Then
        strMsg = "������ļ������볬�����������󳤶�" & LM_�������� & "(" & LM_�������� \ 2 & "�����ĵĳ��Ȼ�" & LM_�������� & "����ĸ�����ֵĳ���)"
        MsgBox strMsg, vbInformation, gstrSysName
        txtSuitCode(0).SetFocus
        CheckInput = False
        Exit Function
    End If
    
    If LenB(StrConv(txtSuitCode(1).Text, vbFromUnicode)) > LM_�������� Then
        strMsg = "��������������볬�����������󳤶�" & LM_�������� & "(" & LM_�������� \ 2 & "�����ĵĳ��Ȼ�" & LM_�������� & "����ĸ�����ֵĳ���)"
        MsgBox strMsg, vbInformation, gstrSysName
        txtSuitCode(1).SetFocus
        CheckInput = False
        Exit Function
    End If
    
    CheckInput = True
    
End Function

Private Sub Form_Resize()
    
    If Me.WindowState = vbMaximized Or Me.WindowState = vbMinimized Then Exit Sub
    '������ı䴰���С
    If Me.Width < 5500 Then Me.Width = 5500
    If Me.Height < 6500 Then Me.Height = 6500
    
End Sub

Private Sub txtCode_Change()
    Call CheckInput(False)
End Sub

Private Sub txtcode_GotFocus()
'���ܣ���ý�����ѡ���ı�
    txtCode.SelStart = 0
    txtCode.SelLength = Len(txtCode.Text)
End Sub

Private Sub txtPathName_Change()
    Call CheckInput(False)
End Sub

Private Sub txtPathName_GotFocus()
'���ܣ���ý�����ѡ���ı�
    txtPathName.SelStart = 0
    txtPathName.SelLength = Len(txtPathName.Text)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'���ܣ��س���λ��һ���ؼ�
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub
