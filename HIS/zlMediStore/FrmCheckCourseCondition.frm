VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCheckCourseCondition 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����̵��"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4035
   Icon            =   "FrmCheckCourseCondition.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2730
      TabIndex        =   7
      Top             =   2775
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1560
      TabIndex        =   6
      Top             =   2775
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   90
      TabIndex        =   8
      Top             =   2775
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "����"
      Height          =   2535
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   3765
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   300
         Left            =   960
         TabIndex        =   12
         Top             =   735
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   159514627
         CurrentDate     =   38552
      End
      Begin VB.CheckBox chk�Զ�ɾ�����ܺ���̵㵥 
         Caption         =   "�Զ�ɾ�����ܺ���̵㵥"
         Height          =   225
         Left            =   690
         TabIndex        =   9
         Top             =   2205
         Width           =   2895
      End
      Begin VB.CheckBox chkֻ����̵㵥�е�ҩƷ�����̵� 
         Caption         =   "ֻ����̵㵥�е�ҩƷ�����̵�"
         Height          =   225
         Left            =   690
         TabIndex        =   5
         Top             =   1950
         Width           =   2895
      End
      Begin VB.ComboBox Cbo�̵�ʱ�� 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1545
         Width           =   2475
      End
      Begin VB.ComboBox cbo�ⷿ 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   330
         Width           =   2475
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   300
         Left            =   960
         TabIndex        =   13
         Top             =   1140
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   159514627
         CurrentDate     =   38552
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��ʼʱ��"
         Height          =   180
         Left            =   180
         TabIndex        =   11
         Top             =   795
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��"
         Height          =   180
         Left            =   180
         TabIndex        =   10
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label lbl�ⷿ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ⷿ"
         Height          =   180
         Left            =   540
         TabIndex        =   1
         Top             =   390
         Width           =   360
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�̵�ʱ��"
         Height          =   180
         Left            =   165
         TabIndex        =   3
         Top             =   1605
         Width           =   720
      End
   End
End
Attribute VB_Name = "FrmCheckCourseCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnSelect As Boolean
Private mlng�ⷿID As Long
Private mstr�̵�ʱ�� As String
Private mbln�̵㵥 As Boolean 'ֻ����̵㵥�е�ҩƷ�����̵�
Private mfrmMain As Form
Private mstr�̵㵥�� As String
Private mblnɾ���̵㵥 As Boolean
Private mcolCheckCourseCard As Collection     '��¼ÿ���̵�ʱ���Ӧ���̵㵥��
Private Sub GetCheckCard()
    Dim rsTmp As New ADODB.Recordset
    Dim str�̵�ʱ�� As String
    Dim strNo As String
    Dim n As Long
    
    On Error GoTo errHandle
    CmdSave.Enabled = False
    'װ���̵�ʱ��
    gstrSQL = "Select Distinct Ƶ�� �̵�ʱ��,No From ҩƷ�շ���¼" & _
              " Where ���� = 14 And NVL(���,' ')<>'1' And �ⷿID+0=[1] " & _
              " And �������� Between [2] And [3] " & _
              " Order by Ƶ��"
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ�̵�ʱ��]", cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), dtpStartDate.Value, dtpEndDate.Value)
    
    Set mcolCheckCourseCard = New Collection
    
    Cbo�̵�ʱ��.Clear
    With rsTmp
        Do While Not .EOF
            If Format(CDate(!�̵�ʱ��), "yyyy-mm-dd hh:mm:ss") = str�̵�ʱ�� Then
                strNo = IIf(strNo = "", "'" & !NO & "'", strNo & "," & "'" & !NO & "'")
                mcolCheckCourseCard.Remove (str�̵�ʱ��)
                mcolCheckCourseCard.Add strNo, str�̵�ʱ��
            Else
                str�̵�ʱ�� = Format(CDate(!�̵�ʱ��), "yyyy-mm-dd hh:mm:ss")
                strNo = "'" & !NO & "'"
                mcolCheckCourseCard.Add strNo, str�̵�ʱ��
                Cbo�̵�ʱ��.AddItem str�̵�ʱ��
            End If
            .MoveNext
        Loop
    End With
    
    rsTmp.Close
    If Cbo�̵�ʱ��.ListCount <> 0 Then
        CmdSave.Enabled = True
        Cbo�̵�ʱ��.ListIndex = 0
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo�ⷿ_Click()
    Call GetCheckCard
End Sub






Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    mlng�ⷿID = cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)
    mstr�̵�ʱ�� = Cbo�̵�ʱ��
    mbln�̵㵥 = (chkֻ����̵㵥�е�ҩƷ�����̵�.Value = 1)
    mblnɾ���̵㵥 = (chk�Զ�ɾ�����ܺ���̵㵥.Value = 1)
    mstr�̵㵥�� = mcolCheckCourseCard.Item(Cbo�̵�ʱ��.Text)
    frmNewCheckCard.txtStock.Caption = cbo�ⷿ.Text
    frmNewCheckCard.txtStock.Tag = mlng�ⷿID
    frmNewCheckCard.txtCheckDate = mstr�̵�ʱ��
'    frmCheckCard.CmdSave.Enabled = False
'    frmCheckCard.CmdCancel.Enabled = False
    
    mblnSelect = True
    Unload Me
End Sub

Private Sub dtpEndDate_Change()
    Call GetCheckCard
End Sub


Private Sub dtpStartDate_Change()
    Call GetCheckCard
End Sub


Private Sub Form_Load()
    Dim mblnSelectStock As String, mintLoop As Integer
    
    mblnSelectStock = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ�̵����", "�ⷿ", "0")
    
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

Public Function GetCondition(FrmMain As Form, ByRef lng�ⷿID As Long, ByRef str�̵㵥�� As String, ByRef bln�̵㵥 As Boolean, ByRef blnɾ���̵㵥) As Boolean
    mblnSelect = False
    Set mfrmMain = FrmMain
    Me.Show vbModal, FrmMain
    GetCondition = mblnSelect
    
    If mblnSelect = False Then Exit Function
    lng�ⷿID = mlng�ⷿID
    bln�̵㵥 = mbln�̵㵥
    str�̵㵥�� = mstr�̵㵥��
    blnɾ���̵㵥 = mblnɾ���̵㵥
End Function
