VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPACSReg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ʼӰ����"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   Icon            =   "frmPacsReg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkUnicode 
      Caption         =   "ͬһ���ߵļ����ڱ�����ͳһ���"
      Height          =   210
      Left            =   120
      TabIndex        =   24
      Top             =   4080
      Width           =   4080
   End
   Begin VB.Frame fraMatch 
      Caption         =   "������ǰ���еļ�飬��������Ŀƥ����ͼ��"
      Height          =   645
      Left            =   -30
      TabIndex        =   43
      Top             =   4395
      Width           =   6135
      Begin VB.OptionButton optMatch 
         Caption         =   "����"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   25
         ToolTipText     =   "�����Ž����˺ͽ��յ�Ӱ�����ƥ��"
         Top             =   300
         Width           =   1395
      End
      Begin VB.OptionButton optMatch 
         Caption         =   "����/סԺ��"
         Height          =   195
         Index           =   1
         Left            =   2550
         TabIndex        =   26
         ToolTipText     =   "�����˱�ʶ�Ž����˺ͽ��յ�Ӱ�����ƥ��"
         Top             =   300
         Width           =   1395
      End
      Begin VB.OptionButton optMatch 
         Caption         =   "����ʶ��"
         Height          =   195
         Index           =   2
         Left            =   4620
         TabIndex        =   27
         ToolTipText     =   "������ʶ�Ž����˺ͽ��յ�Ӱ�����ƥ��"
         Top             =   300
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Height          =   945
      Left            =   -100
      TabIndex        =   32
      Top             =   720
      Width           =   6250
      Begin VB.TextBox txtDept 
         Height          =   300
         Left            =   3780
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   510
         Width           =   2085
      End
      Begin VB.TextBox txtPatID 
         Height          =   300
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   510
         Width           =   1455
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   0
         Left            =   1275
         MaxLength       =   10
         TabIndex        =   1
         Top             =   120
         Width           =   4590
      End
      Begin VB.Label Label7 
         Caption         =   "�������"
         Height          =   255
         Left            =   2940
         TabIndex        =   40
         Top             =   570
         Width           =   765
      End
      Begin VB.Label lblPatID 
         Caption         =   "�����"
         Height          =   225
         Left            =   420
         TabIndex        =   38
         Top             =   570
         Width           =   555
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "������(&T)"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   6255
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4795
      TabIndex        =   29
      Top             =   5145
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3540
      TabIndex        =   28
      Top             =   5145
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   150
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5130
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Height          =   2350
      Left            =   -100
      TabIndex        =   34
      Top             =   1680
      Width           =   6250
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   8
         Left            =   1275
         MaxLength       =   20
         TabIndex        =   13
         Top             =   1230
         Width           =   1455
      End
      Begin VB.ComboBox cboRoom 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   60
         Width           =   4620
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   7
         Left            =   1275
         MaxLength       =   30
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   450
         Width           =   1455
      End
      Begin VB.ComboBox cboSex 
         Height          =   300
         ItemData        =   "frmPacsReg.frx":000C
         Left            =   3795
         List            =   "frmPacsReg.frx":0019
         TabIndex        =   15
         Text            =   "cboSex"
         Top             =   1230
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DTBirth 
         Height          =   300
         Left            =   1275
         TabIndex        =   42
         Top             =   1230
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   80478211
         CurrentDate     =   38156
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   1
         Left            =   3795
         MaxLength       =   30
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   450
         Width           =   2055
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   4
         Left            =   5280
         MaxLength       =   4
         TabIndex        =   17
         Top             =   1230
         Width           =   570
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   2
         Left            =   1275
         MaxLength       =   20
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   3
         Left            =   3795
         MaxLength       =   30
         TabIndex        =   11
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   6
         Left            =   3795
         MaxLength       =   3
         TabIndex        =   21
         Top             =   1620
         Width           =   1815
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   5
         Left            =   1275
         MaxLength       =   3
         TabIndex        =   19
         Top             =   1620
         Width           =   1215
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "������(&C)"
         Height          =   255
         Left            =   1275
         TabIndex        =   22
         Top             =   1995
         Width           =   1455
      End
      Begin VB.CheckBox chk��Ƭ 
         Caption         =   "���Ž�Ƭ(&F)"
         Height          =   255
         Left            =   3795
         TabIndex        =   23
         Top             =   1995
         Width           =   1335
      End
      Begin VB.Label lblRoom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ִ�м�(&R)"
         Height          =   180
         Left            =   420
         TabIndex        =   2
         Top             =   120
         Width           =   810
      End
      Begin VB.Label Label6 
         Caption         =   "����(&U)"
         Height          =   255
         Left            =   420
         TabIndex        =   4
         Top             =   510
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "KG"
         Height          =   180
         Left            =   5680
         TabIndex        =   37
         Top             =   1680
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "CM"
         Height          =   180
         Left            =   2520
         TabIndex        =   36
         Top             =   1680
         Width           =   180
      End
      Begin VB.Label Label3 
         Height          =   135
         Left            =   0
         TabIndex        =   35
         Top             =   -20
         Width           =   6255
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "����豸(&D)"
         Height          =   180
         Index           =   8
         Left            =   2760
         TabIndex        =   6
         Top             =   510
         Width           =   990
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "����(&A)"
         Height          =   180
         Index           =   6
         Left            =   4600
         TabIndex        =   16
         Top             =   1290
         Width           =   630
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   1
         Left            =   600
         TabIndex        =   8
         Top             =   900
         Width           =   630
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "Ӣ����(&E)"
         Height          =   180
         Index           =   4
         Left            =   2940
         TabIndex        =   10
         Top             =   900
         Width           =   810
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�(&S)"
         Height          =   180
         Index           =   5
         Left            =   3120
         TabIndex        =   14
         Top             =   1290
         Width           =   630
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "�绰(&B)"
         Height          =   180
         Index           =   2
         Left            =   600
         TabIndex        =   12
         Top             =   1290
         Width           =   630
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "����(&W)"
         Height          =   180
         Index           =   3
         Left            =   3120
         TabIndex        =   20
         Top             =   1680
         Width           =   630
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "���(&H)"
         Height          =   180
         Index           =   7
         Left            =   600
         TabIndex        =   18
         Top             =   1680
         Width           =   630
      End
   End
   Begin VB.Label Label1 
      Caption         =   $"frmPacsReg.frx":0028
      Height          =   615
      Left            =   840
      TabIndex        =   31
      Top             =   120
      Width           =   4815
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "frmPacsReg.frx":00BA
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmPACSReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private AdviceID As Long, SendNO As Long, mlngPatientID As Long
Private iReturn As Integer, blnModi As Boolean
Private aDevices() As String

Public Function ShowMe(objParent As Object, ByVal lngAdviceID As Long, ByVal lngSendNO As Long) As Integer
    '���أ�0��ȡ����1=��ʼ��顢2���޸ļ����Ϣ
    AdviceID = lngAdviceID: SendNO = lngSendNO
    
    blnModi = False
    Me.Show vbModal, objParent
    ShowMe = iReturn
End Function

Private Sub cboRoom_Click()
    On Error Resume Next
    txtItem(1) = aDevices(cboRoom.ListIndex)
End Sub

Private Sub cboRoom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboSex_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkUnicode_Click()
    If Not blnModi Then
        txtItem(7).Text = Next����(Me.txtItem(0), mlngPatientID, AdviceID, SendNO, chkUnicode.Value = 1)
    End If
End Sub

Private Sub chk����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk��Ƭ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo DBError
    If Len(Trim(txtItem(2))) = 0 Then
        MsgBox "������������", vbInformation, gstrSysName
        txtItem(2).SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(txtItem(2).Text), vbFromUnicode)) > txtItem(2).MaxLength Then
        MsgBox "�������������" & txtItem(2).MaxLength & "���ַ���" & CInt(txtItem(2).MaxLength / 2) & "�����֣���", vbInformation, gstrSysName
        txtItem(2).SetFocus: Exit Sub
    End If
    If Len(Trim(txtItem(3))) = 0 Then
        MsgBox "������Ӣ������", vbInformation, gstrSysName
        txtItem(3).SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(txtItem(3).Text), vbFromUnicode)) > txtItem(3).MaxLength Then
        MsgBox "Ӣ�������������" & txtItem(3).MaxLength & "���ַ���" & CInt(txtItem(3).MaxLength / 2) & "�����֣���", vbInformation, gstrSysName
        txtItem(3).SetFocus: Exit Sub
    End If
    
    '�жϼ����Ƿ��ظ�
    strSQL = "Select ����,�Ա�,���� From Ӱ�����¼ Where Ӱ�����=[1] And ����=[2]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, txtItem(0), txtItem(7))
    If Not rsTmp.EOF Then
        If MsgBox("��ǰ���������л����ظ����Ƿ������" & Chr(10) & Chr(13) & "������Ϣ��" & Nvl(rsTmp(0)) & " " & Nvl(rsTmp(1)) & " " & Nvl(rsTmp(2)), vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            txtItem(7).SetFocus: Exit Sub
        End If
    End If
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��ǰִ�м�", Me.cboRoom.Text
    strSQL = ""
    For i = 0 To cboRoom.ListCount - 1
        strSQL = strSQL & "||" & cboRoom.List(i) & "|" & aDevices(i)
    Next
    If Len(strSQL) > 0 Then strSQL = Mid(strSQL, 3)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "����豸", strSQL
    For i = 0 To optMatch.Count - 1
        If optMatch(i).Value Then Exit For
    Next
    If i > optMatch.Count - 1 Then i = 0
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\��ʼ���", "Ӱ��ƥ�䷽ʽ", i
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\��ʼ���", "��ִ�п��ұ��", chkUnicode.Value
    
    gcnOracle.BeginTrans
    strSQL = "ZL_Ӱ����_BEGIN('" & cboRoom.Text & "'," & txtItem(7).Text & "," & AdviceID & "," & SendNO & ",'" & txtItem(0) & "','" & _
        Trim(txtItem(2)) & "','" & Trim(txtItem(3)) & "','" & Trim(cboSex.Text) & "','" & _
        txtItem(4) & "'," & IIf(IsNull(DTBirth.Value), "Null", "to_Date('" & Format(DTBirth.Value, "yyyy-MM-dd") & "','YYYY-MM-DD')") & ",'" & txtItem(5) & "','" & txtItem(6) & "'," & _
        Me.chk����.Value & "," & Me.chk��Ƭ.Value & ",'" & Trim(txtItem(1)) & "'," & _
        IIf(blnModi, 1, 0) & ",'" & txtItem(8) & "')"
    ExecuteProc strSQL, Me.Caption
        
    '������ǰ���еļ��
    strSQL = "Select A.���UID As ID From Ӱ����ʱ��¼ a " & _
        " Where a.����=[1] And a.Ӱ�����=[2]"
    If optMatch(0).Value Then '����
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, txtItem(7).Text, txtItem(0).Text)
    End If
    If optMatch(1).Value Then '����/סԺ��
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, txtPatID.Text, txtItem(0).Text)
    End If
    If optMatch(2).Value Then '����ʶ�ţ�ҽ��ID��
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, AdviceID, txtItem(0).Text)
    End If
    If rsTmp.RecordCount = 1 Then '��ͼ��ͼ���Զ�ƥ��
        strSQL = "ZL_Ӱ����_SET(" & AdviceID & "," & SendNO & ",'" & _
            rsTmp("ID") & "')"
        ExecuteProc strSQL, Me.Caption
    End If
    gcnOracle.CommitTrans
    
    iReturn = IIf(blnModi, 2, 1)
    Unload Me
    Exit Sub
DBError:
    gcnOracle.RollbackTrans
    Call ErrCenter
    txtItem(7).SetFocus
    Call SaveErrLog
End Sub

Private Sub DTBirth_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Activate()
    If Me.Tag = "Unload" Then
        Me.Tag = ""
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    Call cmdCancel_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strExeRoom As String
    
    iReturn = 0
    
    On Error GoTo DBError
    
    strSQL = "Select Nvl(A.ִ�в���ID,0) As ִ�в���ID,Nvl(E.����,C.����) As ����,Nvl(E.����,C.����) As ����," & _
        "Nvl(E.�Ա�,C.�Ա�) As �Ա�,Nvl(E.��������,C.��������) As ��������," & _
        "Nvl(D.Ӱ�����,' ') As Ӱ�����,E.������,E.���Ž�Ƭ," & _
        "Nvl(D.���в���,0) As ���в���,Nvl(D.�ɷ���Ƭ,0) As �ɷ���Ƭ," & _
        "E.����,E.Ӣ����,E.���,E.����,E.����豸,A.ִ�м�,Nvl(A.ִ��״̬,0) As ִ��״̬,B.����ID," & _
        "Nvl(E.��ϵ�绰,C.��ϵ�˵绰) As ��ϵ�˵绰,B.������Դ,Decode(B.������Դ,2,C.סԺ��,C.�����) As ��ʶ��,F.���� As ������� " & _
        "From ����ҽ������ A,����ҽ����¼ B,������Ϣ C,Ӱ������Ŀ D,Ӱ�����¼ E,���ű� F " & _
        "Where A.ҽ��ID=B.ID And B.����ID=C.����ID And B.������ĿID=D.������ĿID(+) " & _
        "And A.ҽ��ID=E.ҽ��ID(+) And A.���ͺ�=E.���ͺ�(+) And B.��������ID=F.ID " & _
        "And A.ҽ��ID= [1] And A.���ͺ�=[2] "
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, AdviceID, SendNO)
    
    If rsTmp.EOF Then
        MsgBox "������ȷ��ȡִ����Ŀ��Ϣ��", vbInformation, gstrSysName
        Me.Tag = "Unload": Exit Sub
    End If
    If rsTmp("ִ��״̬") = 1 Or rsTmp("ִ��״̬") = 2 Then
        MsgBox "�ü���ѱ�������" & IIf(rsTmp("ִ��״̬") = 1, "ִ����ɡ�", "�ܾ�ִ�С�"), vbInformation, gstrSysName
        Me.Tag = "Unload": Exit Sub
    End If
    
    mlngPatientID = Nvl(rsTmp!����ID, 0)
    Me.txtItem(0) = rsTmp("Ӱ�����")
    Me.lblPatID.Caption = IIf(Nvl(rsTmp("������Դ"), 0) = 2, "סԺ��", "�����")
    Me.txtPatID = Nvl(rsTmp("��ʶ��"))
    Me.txtDept = Nvl(rsTmp("�������"))
    Me.txtItem(2) = rsTmp("����")
    Me.txtItem(4) = Nvl(rsTmp("����")): Me.cboSex.Text = Nvl(rsTmp("�Ա�"), " ")
    If IsNull(rsTmp!��������) Then
        DTBirth.Value = Empty
    Else
        DTBirth.Value = rsTmp!��������
    End If
    Me.txtItem(8) = Nvl(rsTmp("��ϵ�˵绰"))
    chk����.Value = Nvl(rsTmp!������, 0)
    Select Case rsTmp("���в���")
        Case 0, 1
            chk����.Value = rsTmp("���в���"): chk����.Enabled = False
        Case Else
            chk����.Enabled = True
    End Select
    
    chk��Ƭ.Value = Nvl(rsTmp!���Ž�Ƭ, 0)
    Select Case rsTmp("�ɷ���Ƭ")
        Case 0, 1
            chk��Ƭ.Value = rsTmp("�ɷ���Ƭ"): chk��Ƭ.Enabled = False
        Case Else
            chk��Ƭ.Enabled = True
    End Select
    txtItem(1).Text = Nvl(rsTmp!����豸)
    txtItem(3).Text = Nvl(rsTmp!Ӣ����, UCase(Replace(zlCommFun.mGetFullPY(Trim(txtItem(2))), vbCrLf, "")))
    txtItem(5).Text = Nvl(rsTmp!���)
    txtItem(6).Text = Nvl(rsTmp!����)
    If Not IsNull(rsTmp!����) Then
        txtItem(7).Text = rsTmp!����
        blnModi = True
    Else
        txtItem(7).Text = Next����(rsTmp!Ӱ�����, Nvl(rsTmp!����ID, 0), AdviceID, SendNO, chkUnicode.Value = 1)
    End If
    
    'ִ�м�����
    strExeRoom = Nvl(rsTmp("ִ�м�"))
    If Len(Trim(strExeRoom)) = 0 Then 'ȡĬ�ϱ���ִ�м�
        strExeRoom = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName, "��ǰִ�м�")
    End If
    If rsTmp("ִ�в���ID") = 0 Then
        strSQL = "Select * From ҽ��ִ�з���"
    Else
        strSQL = "Select * From ҽ��ִ�з��� Where ����ID=[1]"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Val(rsTmp("ִ�в���ID")))
    cboRoom.Clear
    If rsTmp.EOF Then
        cboRoom.AddItem "": cboRoom.ListIndex = 0
    Else
        Do While Not rsTmp.EOF
            cboRoom.AddItem rsTmp!ִ�м�
            rsTmp.MoveNext
        Loop
    End If
    InitDevice
    
    'Ӱ��ƥ������
    i = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\��ʼ���", "Ӱ��ƥ�䷽ʽ", 0))
    optMatch(i).Value = True
    chkUnicode.Value = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\��ʼ���", "��ִ�п��ұ��", 0))
    
    If blnModi Then
        Me.Caption = "�޸�Ӱ����Ϣ"
    Else
        Me.Caption = "��ʼӰ����"
    End If
    On Error Resume Next
    cboRoom.ListIndex = 0
    cboRoom.Text = strExeRoom
    On Error GoTo DBError
    
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optMatch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtItem_GotFocus(Index As Integer)
    With Me.txtItem(Index)
        .SelStart = 0: .SelLength = .MaxLength
    End With
    Select Case Index
        Case 1, 2
            Call zlCommFun.OpenIme(True)
        Case Else
            Call zlCommFun.OpenIme(False)
    End Select
End Sub

Private Sub txtItem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
        '�õ�ƴ��
        If Trim(txtItem(3)) = "" Then
            txtItem(3).Text = UCase(Replace(zlCommFun.mGetFullPY(Trim(txtItem(2).Text)), vbCrLf, ""))
        End If
    End If
End Sub

Private Sub txtItem_KeyPress(Index As Integer, KeyAscii As Integer)
    If ifEditKey(KeyAscii, False) Then Exit Sub
    
    If LenB(StrConv(Trim(txtItem(Index).Text), vbFromUnicode)) >= txtItem(Index).MaxLength Then
        KeyAscii = 0
        Exit Sub
    End If
    Select Case Index
        Case 5, 6, 7
            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then KeyAscii = 0
    End Select
End Sub

Private Sub txtItem_LostFocus(Index As Integer)
    Select Case Index
        Case 1, 2
            Call zlCommFun.OpenIme(False)
            If Index = 1 Then aDevices(cboRoom.ListIndex) = txtItem(1)
    End Select
End Sub

'�ж��Ƿ�Ϊ�༭��
Private Function ifEditKey(ByVal KeyAscii As Integer, Optional ByVal AllowSubtract As Boolean = True) As Boolean
    If KeyAscii = vbKeyBack Or (KeyAscii = vbKeyInsert And AllowSubtract) Or KeyAscii = vbKeyDelete Or _
      KeyAscii = vbKeyHome Or KeyAscii = vbKeyEnd Or KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or _
      KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Then
        ifEditKey = True
    Else
        ifEditKey = False
    End If
End Function

Private Function Next����(str��� As String, ByVal lngPatientID As Long, Optional ByVal lngAdviceID As Long = 0, Optional ByVal lngSendNO As Long = 0, Optional ByVal blnUnicode As Boolean = False) As Double
    Dim rsCtrl As New ADODB.Recordset
    Dim strSQL As String, lngNO As Double
    Dim lngExeDept As Long

ReStart:
    Err = 0
    On Error GoTo errH
    
    If Not blnUnicode Then '�������
        strSQL = "Select ���� From Ӱ�����¼ A,����ҽ����¼ B" & _
            " Where A.ҽ��ID=B.ID And B.����ID=[1] And A.Ӱ�����=[2] Order By B.ͣ��ʱ�� Desc"
        Set rsCtrl = OpenSQLRecord(strSQL, Me.Caption, lngPatientID, str���)
        If Not rsCtrl.EOF Then
            lngNO = Val(Nvl(rsCtrl("����"), 0))
        Else
            strSQL = "Select * From Ӱ������� Where ����=[1]"
            Set rsCtrl = OpenSQLRecord(strSQL, Me.Caption, str���)
            If rsCtrl.EOF Then Exit Function
            
            lngNO = Val(Nvl(rsCtrl("������"), 0)) + 1
        End If
    Else '��ִ�п��ұ��
        strSQL = "Select A.ִ�в���ID,B.����ID From ����ҽ������ A,����ҽ����¼ B" & _
            " Where A.ҽ��ID=B.ID And A.ҽ��ID=[1] And A.���ͺ�=[2]"
        Set rsCtrl = OpenSQLRecord(strSQL, Me.Caption, lngAdviceID, lngSendNO)
        If rsCtrl.EOF Then
            Next���� = 0
            Exit Function
        End If
        
        lngExeDept = Nvl(rsCtrl(0), 0)
        strSQL = "Select A.���� From Ӱ�����¼ A,����ҽ����¼ B" & _
            " Where A.ҽ��ID=B.ID+0 And B.ִ�п���ID+0=[1] And B.����ID=[2] Order By B.ͣ��ʱ�� Desc"
        Set rsCtrl = OpenSQLRecord(strSQL, Me.Caption, lngExeDept, lngPatientID)
        If rsCtrl.EOF Then 'ȡ����������
'            strSQL = "Select * From Ӱ������� Where ����=[1]"
'            Set rsCtrl = OpenSQLRecord(strSQL, Me.Caption, str���)
            strSQL = "SELECT DISTINCT C.����,Nvl(C.������,0) FROM Ӱ������Ŀ A,����ִ�п��� B,Ӱ������� C" & _
                " WHERE A.������ĿID=B.������ĿID+0 AND A.Ӱ�����||''=C.���� AND B.ִ�п���ID=[1] ORDER BY Nvl(C.������,0) DESC"
            Set rsCtrl = OpenSQLRecord(strSQL, Me.Caption, lngExeDept)
            If rsCtrl.EOF Then
                Next���� = 0
                Exit Function
            End If
            
            lngNO = Val(Nvl(rsCtrl(1), 0)) + 1
        Else
            lngNO = Val(Nvl(rsCtrl("����"), 0))
        End If
    End If
    
    Next���� = lngNO
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitDevice()
    Dim i As Integer, iPos As Integer
    Dim strDevices As String, aTmpArray() As String, aTmpArray1() As String
    On Error Resume Next
    
    strDevices = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "����豸")
    If Len(Trim(strDevices)) = 0 Then
        ReDim aTmpArray(1, 0) As String
    Else
        aTmpArray1 = Split(strDevices, "||")
        ReDim aTmpArray(1, UBound(aTmpArray1)) As String
        For i = 0 To UBound(aTmpArray1)
            iPos = InStr(aTmpArray1(i), "|")
            If iPos = 0 Then
                aTmpArray(0, i) = ""
                aTmpArray(1, i) = aTmpArray1(i)
            Else
                aTmpArray(0, i) = Mid(aTmpArray1(i), 1, iPos - 1)
                aTmpArray(1, i) = Mid(aTmpArray1(i), iPos + 1)
            End If
        Next
    End If
    
    ReDim aDevices(cboRoom.ListCount - 1) As String
    For i = 0 To cboRoom.ListCount - 1
        iPos = GetIndex(aTmpArray, cboRoom.List(i))
        aDevices(i) = aTmpArray(1, iPos)
    Next
End Sub

Private Function GetIndex(aSeekArray() As String, ByVal vSeekValue As Variant) As Integer
    Dim i As Integer
    For i = 0 To UBound(aSeekArray, 2)
        If aSeekArray(0, i) = vSeekValue Then Exit For
    Next
    If i > UBound(aSeekArray, 2) Then
        GetIndex = 0
    Else
        GetIndex = i
    End If
End Function
