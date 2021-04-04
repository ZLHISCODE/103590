VERSION 5.00
Begin VB.Form frmLabSamplingSelect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "������Ŀѡ��"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboDept 
      Height          =   300
      Left            =   930
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   30
      Width           =   5025
   End
   Begin VB.CommandButton cmdOK 
      Height          =   315
      Left            =   4800
      Picture         =   "frmLabSamplingSelect.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "ȷ��(F2)"
      Top             =   5040
      Width           =   450
   End
   Begin VB.ComboBox cboSample 
      Height          =   300
      Left            =   510
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   5040
      Width           =   1305
   End
   Begin VB.OptionButton OptUnionItem 
      Caption         =   "����ָ��"
      Height          =   225
      Index           =   1
      Left            =   3360
      TabIndex        =   5
      Top             =   5070
      Width           =   1125
   End
   Begin VB.OptionButton OptUnionItem 
      Caption         =   "�����Ŀ"
      Height          =   225
      Index           =   0
      Left            =   2010
      TabIndex        =   4
      Top             =   5070
      Value           =   -1  'True
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   5490
      Picture         =   "frmLabSamplingSelect.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "ȡ��(Esc)"
      Top             =   5040
      Width           =   450
   End
   Begin VB.Frame Frame1 
      Height          =   4635
      Left            =   30
      TabIndex        =   0
      Top             =   360
      Width           =   5925
      Begin VB.ListBox listGoal 
         Height          =   4380
         Left            =   3210
         TabIndex        =   9
         Top             =   180
         Width           =   2625
      End
      Begin VB.ListBox listSoure 
         Height          =   4380
         Left            =   90
         TabIndex        =   8
         Top             =   180
         Width           =   2625
      End
      Begin VB.CommandButton cmdLeft 
         Caption         =   "<"
         Height          =   375
         Left            =   2760
         TabIndex        =   2
         Top             =   2535
         Width           =   405
      End
      Begin VB.CommandButton cmdRight 
         Caption         =   ">"
         Height          =   375
         Left            =   2760
         TabIndex        =   1
         Top             =   1440
         Width           =   405
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ִ�п���"
      Height          =   180
      Left            =   120
      TabIndex        =   12
      Top             =   90
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�걾"
      Height          =   180
      Left            =   90
      TabIndex        =   7
      Top             =   5100
      Width           =   360
   End
End
Attribute VB_Name = "frmLabSamplingSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mlngDeptID As Long
Dim mStrRecord As String
Public Function ShowMe(objfrm As Object, lngDeptID As Long) As String
    mlngDeptID = lngDeptID
    frmLabSamplingSelect.Show vbModal, objfrm
    ShowMe = mStrRecord
End Function
Private Sub LoadListData()
    '����           �����б�����
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    Me.listSoure.Clear
    If Me.cboSample.ListCount = 0 Then Exit Sub
    If Me.cboDept.ListIndex = -1 Then Exit Sub
       
    strSQL = "Select Distinct A.ID, A.����, A.���� " & vbCrLf & _
             " From ������ĿĿ¼ A, ����ִ�п��� B, ���鱨����Ŀ C " & vbCrLf & _
             " Where A.��� = 'C'  And A.ID = B.������Ŀid And A.ID = C.������Ŀid And " & vbCrLf & _
             " B.ִ�п���ID = [1] " & IIf(Trim(Me.cboSample.Text) = "���б걾", "", " And A.�걾��λ = [2] ") & _
             " And �����Ŀ = [3] order by a.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, CLng(Me.cboDept.ItemData(Me.cboDept.ListIndex)), _
                CStr(cboSample.Text), IIf(Me.OptUnionItem(0).Value, 1, 0))
                
    Do Until rsTmp.EOF
        If CheckItemRepeat(rsTmp("ID")) = False Then
            With Me.listSoure
                .AddItem rsTmp("����")
                .ItemData(.NewIndex) = rsTmp("ID")
            End With
        End If
        rsTmp.MoveNext
    Loop
End Sub

Private Sub LoadcboSample()
    '����           ����걾����
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    strSQL = "Select Distinct c.���� ,  A.�걾��λ " & _
             " From ������ĿĿ¼ A, ����ִ�п��� B , ���Ƽ���걾 c " & _
             " Where A.ID = B.������Ŀid And A.��� = 'C' And A.�������� Is Not Null " & _
             " And A.�걾��λ Is Not Null And B.ִ�п���id = [1] and " & _
             " a.�걾��λ = c.���� Order By c.���� "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, mlngDeptID)
    With Me.cboSample
        .AddItem "���б걾"
        .ItemData(.NewIndex) = 0
        Do Until rsTmp.EOF
            Me.cboSample.AddItem rsTmp("�걾��λ")
            rsTmp.MoveNext
        Loop
    End With
    
    strSQL = "  Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B,������Ա C " & _
            " Where B.����ID = A.ID And A.ID=C.����ID " & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " And B.������� IN(1,2,3,4) And B.�������� IN('����')"
            
    strSQL = strSQL & " Order by A.����"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    
    cboDept.Clear
    Do Until rsTmp.EOF
        cboDept.AddItem rsTmp("����") & "-" & rsTmp("����")
        cboDept.ItemData(cboDept.NewIndex) = rsTmp("ID")
        If rsTmp("id") = IIf(mlngDeptID = 0, UserInfo.����ID, mlngDeptID) Then
            cboDept.ListIndex = cboDept.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    
    If Me.cboSample.ListCount > 0 Then
        Me.cboSample.ListIndex = 0
    End If
    
    If cboDept.Text = "" Then
        cboDept.ListIndex = 0
    End If
End Sub

Private Function CheckItemRepeat(lngItemID As Long) As Boolean
    '�����Ƿ��ظ�
    '����                       ������ĿĿ¼_ID
    '����                       =True �ظ� =False ���ظ�
    Dim intLoop As Integer
    
    If Me.listGoal.ListCount = 0 Then Exit Function
    
    For intLoop = 0 To Me.listGoal.ListCount - 1
        If lngItemID = Me.listGoal.ItemData(intLoop) Then
            CheckItemRepeat = True
            Exit Function
        End If
    Next
    
End Function

Private Sub cboDept_Click()
    Me.listGoal.Clear
    LoadListData
End Sub

Private Sub cboSample_Click()
    Me.listGoal.Clear
    LoadListData
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdLeft_Click()
    listGoal_DblClick
    Me.listGoal.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim intLoop As Integer
    mStrRecord = ""
    If Me.listGoal.ListCount > 0 Then
        For intLoop = 0 To Me.listGoal.ListCount - 1
            mStrRecord = mStrRecord & "," & Me.listGoal.ItemData(intLoop)
        Next
    End If
    If Len(mStrRecord) > 0 Then
        mStrRecord = Mid(mStrRecord, 2) & ";" & Me.cboSample.Text
    End If
    If mStrRecord <> "" Then
        Unload Me
    End If
End Sub

Private Sub cmdRight_Click()
    Call listSoure_DblClick
    Me.listSoure.SetFocus
End Sub

Private Sub Form_Load()
    LoadcboSample
    LoadListData
End Sub

Private Sub listGoal_DblClick()
    If Me.listGoal.ListCount > 0 Then
        If Me.listGoal.ListIndex = -1 Then Me.listGoal.ListIndex = 0
        Me.listGoal.RemoveItem (Me.listGoal.ListIndex)
        LoadListData
    End If
End Sub

Private Sub listSoure_DblClick()
    If Me.listSoure.ListCount > 0 Then
        If Me.listSoure.ListIndex = -1 Then Me.listSoure.ListIndex = 0
        Me.listGoal.AddItem Me.listSoure.List(Me.listSoure.ListIndex)
        Me.listGoal.ItemData(Me.listGoal.NewIndex) = Me.listSoure.ItemData(Me.listSoure.ListIndex)
        Me.listSoure.RemoveItem (Me.listSoure.ListIndex)
    End If
End Sub

Private Sub OptUnionItem_Click(Index As Integer)
    Me.listGoal.Clear
    LoadListData
End Sub
