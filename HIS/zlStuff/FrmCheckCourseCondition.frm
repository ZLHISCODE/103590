VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmCheckCourseCondition 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����̵��"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "FrmCheckCourseCondition.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox Cbo�̵�ʱ�� 
      Height          =   300
      Left            =   1185
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2085
      Width           =   2910
   End
   Begin VB.CheckBox chkֻ����̵㵥�е����Ľ����̵� 
      Caption         =   "ֻ����̵㵥�е����Ľ����̵�(&Z)"
      Height          =   225
      Left            =   1185
      TabIndex        =   9
      Top             =   2490
      Width           =   3240
   End
   Begin VB.CheckBox chk�Զ�ɾ�����ܺ���̵㵥 
      Caption         =   "�Զ�ɾ�����ܺ���̵㵥(&S)"
      Height          =   225
      Left            =   1185
      TabIndex        =   10
      Top             =   2745
      Width           =   2895
   End
   Begin VB.ComboBox cbo�ⷿ 
      Height          =   300
      Left            =   1185
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1020
      Width           =   2895
   End
   Begin VB.Frame fra 
      Height          =   105
      Index           =   1
      Left            =   -270
      TabIndex        =   15
      Top             =   3015
      Width           =   6660
   End
   Begin VB.Frame fra 
      Height          =   105
      Index           =   0
      Left            =   -75
      TabIndex        =   14
      Top             =   765
      Width           =   6285
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   120
      TabIndex        =   13
      Top             =   3240
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3945
      TabIndex        =   12
      Top             =   3240
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2790
      TabIndex        =   11
      Top             =   3240
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   300
      Left            =   1185
      TabIndex        =   4
      Top             =   1365
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   114491395
      CurrentDate     =   38552
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   300
      Left            =   1185
      TabIndex        =   6
      Top             =   1695
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   114491395
      CurrentDate     =   38552
   End
   Begin VB.Label lblInfor 
      Caption         =   "�����̵��¼�������̴�����������еġ���ʼʱ�䡢����ʱ�䡱��Ҫ���ڹ��˳��̵�ʱ�䡣"
      ForeColor       =   &H8000000D&
      Height          =   525
      Left            =   795
      TabIndex        =   0
      Top             =   390
      Width           =   4215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "����ʱ��(&J)"
      Height          =   180
      Left            =   135
      TabIndex        =   5
      Top             =   1785
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��ʼʱ��(&K)"
      Height          =   180
      Left            =   135
      TabIndex        =   3
      Top             =   1425
      Width           =   990
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�̵�ʱ��(&P)"
      Height          =   180
      Left            =   135
      TabIndex        =   7
      Top             =   2130
      Width           =   990
   End
   Begin VB.Label lbl�ⷿ 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ⷿ(&D)"
      Height          =   180
      Left            =   495
      TabIndex        =   1
      Top             =   1080
      Width           =   630
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   135
      Picture         =   "FrmCheckCourseCondition.frx":000C
      Top             =   195
      Width           =   480
   End
End
Attribute VB_Name = "FrmCheckCourseCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnSelect As Boolean
Private mlng�ⷿid As Long
Private mstr�̵�ʱ�� As String
Private mfrmMain As Form
Private mcllNO As Collection
Private mbln�̵㵥 As Boolean 'ֻ����̵㵥�е�ҩƷ�����̵�
Private mstr�̵㵥�� As String
Private mblnɾ���̵㵥 As Boolean
Private Const mlngModule = 1719


Private Sub GetCheckCard()
    Dim rsTemp As New ADODB.Recordset
    Dim str�̵�ʱ�� As String
    Dim strNo As String
    Dim n As Long
    
    On Error GoTo ErrHandle
    CmdSave.Enabled = False
    gstrSQL = "" & _
        "   Select Distinct Ƶ�� �̵�ʱ��,No From ҩƷ�շ���¼" & _
        "   Where ���� = 23 And NVL(���,' ')<>'1' And �ⷿID+0=[1] " & _
        "           And �������� Between [2] And [3] " & _
        "   Order by Ƶ��"
        
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ�̵�ʱ��]", cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), dtpStartDate.Value, dtpEndDate.Value)
    
    Cbo�̵�ʱ��.Clear
    Set mcllNO = New Collection
    With rsTemp
        Do While Not .EOF
            If Format(CDate(!�̵�ʱ��), "yyyy-mm-dd hh:mm:ss") = str�̵�ʱ�� Then
                strNo = IIf(strNo = "", "'" & !NO & "'", strNo & "," & "'" & !NO & "'")
                mcllNO.Remove (str�̵�ʱ��)
                mcllNO.Add strNo, str�̵�ʱ��
            Else
                str�̵�ʱ�� = Format(CDate(!�̵�ʱ��), "yyyy-mm-dd hh:mm:ss")
                strNo = "'" & !NO & "'"
                mcllNO.Add strNo, str�̵�ʱ��
                Cbo�̵�ʱ��.AddItem str�̵�ʱ��
            End If
            .MoveNext
        Loop
    End With
    
    If Cbo�̵�ʱ��.ListCount <> 0 Then
        CmdSave.Enabled = True
        Cbo�̵�ʱ��.ListIndex = 0
    Else
        CmdSave.Enabled = False
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub cbo�ⷿ_Click()
'    Dim rsTmp As New ADODB.Recordset
'    CmdSave.Enabled = False
'    'װ���̵�ʱ��
'    gstrSQL = "" & _
'        "   Select Distinct Ƶ�� �̵�ʱ�� From ҩƷ�շ���¼" & _
'        "   Where ���� = 23 And �ⷿID=[1]" & _
'        "       And Ƶ�� Not In (" & _
'        "               Select Distinct Ƶ�� From ҩƷ�շ���¼" & _
'        "               Where ����=22 And �ⷿID=[1]" & _
'        "               And ����� Is Not Null And Mod(��¼״̬,3)=1 And Ƶ�� Is Not Null)" & _
'        "    Order by Ƶ��"
'
'    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "-��ȡ�̵�ʱ��", cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex))
'
'    With rsTmp
'        Cbo�̵�ʱ��.Clear
'        Do While Not .EOF
'            Cbo�̵�ʱ��.AddItem !�̵�ʱ��
'            .MoveNext
'        Loop
'    End With
'
'    rsTmp.Close
'    If Cbo�̵�ʱ��.ListCount <> 0 Then
'        CmdSave.Enabled = True
'        Cbo�̵�ʱ��.ListIndex = 0
'    End If

    Call GetCheckCard
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub CmdSave_Click()
    mlng�ⷿid = cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)
    mstr�̵�ʱ�� = Cbo�̵�ʱ��
    
    mbln�̵㵥 = (chkֻ����̵㵥�е����Ľ����̵�.Value = 1)
    mblnɾ���̵㵥 = (chk�Զ�ɾ�����ܺ���̵㵥.Value = 1)
    mstr�̵㵥�� = mcllNO.Item(Cbo�̵�ʱ��.Text)
    
    frmCheckCard.txtStock.Caption = cbo�ⷿ.Text
    frmCheckCard.txtStock.Tag = mlng�ⷿid
    frmCheckCard.txtCheckDate = mstr�̵�ʱ��
    frmCheckCard.CmdSave.Enabled = False
    frmCheckCard.CmdCancel.Enabled = False
    mblnSelect = True
    Unload Me
End Sub

 

Private Sub dtpEndDate_Change()
    If dtpEndDate.Value < dtpStartDate.Value Then
        dtpStartDate.Value = dtpEndDate.Value
    End If
    
    Call GetCheckCard
End Sub


Private Sub dtpStartDate_Change()
    If dtpEndDate.Value < dtpStartDate.Value Then
        dtpEndDate.Value = Format(dtpStartDate.Value, "yyyy-mm-dd") & " 23:59:59"
    End If
    Call GetCheckCard
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    Dim mblnSelectStock As String, mintLoop As Integer
    
    mblnSelectStock = IIf(Val(zlDatabase.GetPara("�Ƿ�ѡ��ⷿ", glngSys, mlngModule, "0")) = 1, 1, 0)
    dtpStartDate.CustomFormat = "yyyy-MM-dd HH:mm:ss"
    dtpEndDate.CustomFormat = "yyyy-MM-dd HH:mm:ss"
    dtpEndDate.Value = Format(Sys.Currentdate, "yyyy-mm-dd 23:59:59")
    dtpStartDate.Value = Format(Sys.Currentdate, "yyyy-mm-dd") & " 00:00:00"
    
    
    'װ��ⷿ
    With mfrmMain.cboStock
        cbo�ⷿ.Clear
        For mintLoop = 0 To .ListCount - 1
            cbo�ⷿ.AddItem .List(mintLoop)
            cbo�ⷿ.ItemData(cbo�ⷿ.NewIndex) = .ItemData(mintLoop)
        Next
        cbo�ⷿ.ListIndex = .ListIndex
    End With
    If InStr(1, mfrmMain.mstrPrivs, "���пⷿ") <> 0 Then
        If mblnSelectStock = "0" Then
            cbo�ⷿ.Enabled = False
        Else
            cbo�ⷿ.Enabled = True
        End If
    Else
        cbo�ⷿ.Enabled = False
    End If
End Sub

Public Function GetCondition(frmMain As Form, _
            ByRef lng�ⷿID As Long, ByRef str�̵�ʱ�� As String, _
            ByRef str�̵㵥�� As String, ByRef blnֻͳ���̵㵥���� As Boolean, _
            ByRef blnɾ���̵㵥 As Boolean) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���̵��¼�����ܵ��������
    '���:frmMain-������
    '����:
    '       lng�ⷿID-�ⷿID
    '       str�̵�ʱ��-�̵�ʱ���ʽΪyyyy-mm-dd hh24:mi:ss
    '       str�̵㵥��-�̵㵥�ݺ�,��'NO','NO'�ָ�
    '       blnֻͳ���̵㵥����-ֻͳ���̵��д��ڵĵ�����
    '       blnɾ���̵㵥-ɾ�����������ܵ��̵㵥�е�����,��������str�̵㵥���е��̵��¼��
    '����:����ȷ��Ϊtrue,����ΪFalse
    '------------------------------------------------------------------------------------------------------------------------------
    mblnSelect = False
    Set mfrmMain = frmMain
    Me.Show vbModal, frmMain
    GetCondition = mblnSelect
    If mblnSelect = False Then Exit Function
    lng�ⷿID = mlng�ⷿid
    str�̵�ʱ�� = mstr�̵�ʱ��
    str�̵㵥�� = mstr�̵㵥��
    
    blnֻͳ���̵㵥���� = mbln�̵㵥
    blnɾ���̵㵥 = mblnɾ���̵㵥
End Function

 
