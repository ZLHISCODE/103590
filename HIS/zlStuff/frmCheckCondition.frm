VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCheckCondition 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�̵���������"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin TabDlg.SSTab sst 
      Height          =   5895
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   10398
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "����(&1)"
      TabPicture(0)   =   "frmCheckCondition.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Lbl����"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDate"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Lbl�̵㷽ʽ"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl�ⷿ"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tvw����"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "dtpDate"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkNoNum"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkNum"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Cbo�̵㷽ʽ"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cbo�ⷿ"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "�ⷿ��λ(&2)"
      TabPicture(1)   =   "frmCheckCondition.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "tvw��λ"
      Tab(1).Control(2)=   "chk��λ"
      Tab(1).ControlCount=   3
      Begin VB.CheckBox chk��λ 
         Caption         =   "����ʾ��ǰ�ⷿ�ѷ���Ļ�λ"
         Height          =   255
         Left            =   -73080
         TabIndex        =   14
         Top             =   480
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.ComboBox cbo�ⷿ 
         Height          =   300
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   3045
      End
      Begin VB.ComboBox Cbo�̵㷽ʽ 
         Height          =   300
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   4185
         Width           =   3045
      End
      Begin VB.CheckBox chkNum 
         Caption         =   "���޿�����"
         Height          =   255
         Left            =   930
         TabIndex        =   5
         Top             =   4965
         Width           =   1935
      End
      Begin VB.CheckBox chkNoNum 
         Caption         =   "���������������п������۵Ĳ���"
         Height          =   255
         Left            =   930
         TabIndex        =   4
         Top             =   5250
         Width           =   3585
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Left            =   930
         TabIndex        =   6
         Top             =   4575
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
         Format          =   121700355
         CurrentDate     =   36901
      End
      Begin MSComctlLib.TreeView tvw���� 
         Height          =   3000
         Left            =   240
         TabIndex        =   9
         Top             =   1125
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   5292
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "img16"
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView tvw��λ 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   15
         Top             =   750
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   8493
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   476
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "img16"
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���Ͽⷿ��λ(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   -74760
         TabIndex        =   16
         Top             =   510
         Width           =   1350
      End
      Begin VB.Label lbl�ⷿ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ⷿ(&S)"
         Height          =   180
         Left            =   285
         TabIndex        =   13
         Top             =   540
         Width           =   630
      End
      Begin VB.Label Lbl�̵㷽ʽ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʽ(&F)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   270
         TabIndex        =   12
         Top             =   4245
         Width           =   630
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʱ��(&T)"
         Height          =   180
         Left            =   270
         TabIndex        =   11
         Top             =   4635
         Width           =   630
      End
      Begin VB.Label Lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&G)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   285
         TabIndex        =   10
         Top             =   870
         Width           =   630
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   5055
      TabIndex        =   2
      Top             =   4980
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5040
      TabIndex        =   1
      Top             =   960
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5040
      TabIndex        =   0
      Top             =   525
      Width           =   1100
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCondition.frx":0038
            Key             =   "Item"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCondition.frx":0F12
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCheckCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnSelect As Boolean
Private mblnBootUp As Boolean
Private mstr����ID  As String
Private mlng�ⷿid As Long
Private mlng�̵㷽ʽ As Integer
Private mstr�̵�ʱ�� As String
Private mint���޿����� As Integer
Private mbln�̵����������н�� As Boolean
Private mfrmMain As Form
Private Const mlngModule = 1719
Private mstr�ⷿ��λ  As String
Public Function GetCondition(frmMain As Form, ByRef str����ID As String, ByRef lng�ⷿID As Long, _
        ByRef �̵㷽ʽ As Integer, ByRef str�̵�ʱ��, ByRef int���޿����� As Integer, _
        ByRef bln�̵����������н�� As Boolean, ByRef str�ⷿ��λ As String) As Boolean
    
    mstr����ID = ""
    mlng�ⷿid = 0
    mlng�̵㷽ʽ = 0
    mstr�̵�ʱ�� = ""
    mint���޿����� = 0
    mblnSelect = False
    mbln�̵����������н�� = False
    mstr�ⷿ��λ = "����"
    
    Set mfrmMain = frmMain
    Me.Show vbModal, frmMain
    GetCondition = mblnSelect
    
    str����ID = mstr����ID
    lng�ⷿID = mlng�ⷿid
    �̵㷽ʽ = mlng�̵㷽ʽ
    str�̵�ʱ�� = mstr�̵�ʱ��
    int���޿����� = mint���޿�����
    bln�̵����������н�� = mbln�̵����������н��
    str�ⷿ��λ = mstr�ⷿ��λ
End Function

Private Sub chkNoNum_Click()
    chkNum.Enabled = chkNoNum.Value <> 1
    If chkNum.Enabled = False Then
        chkNum.Value = False
    End If
End Sub
Private Sub chkNum_Click()
    chkNoNum.Enabled = chkNum.Value <> 1
    If chkNoNum.Enabled = False Then
        chkNoNum.Value = 0
    End If
End Sub

 
Private Sub chk��λ_Click()
    Load�ⷿ��λ
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim intItem As Integer, intItems As Integer
    Dim i As Long
    
    mstr����ID = ""
    If tvw����.Nodes("Root").Checked Then
        '��������
        mstr����ID = "������������"
    Else
        For i = 1 To tvw����.Nodes.Count
            If tvw����.Nodes(i).Key <> "Root" And _
                tvw����.Nodes(i).Checked Then
                mstr����ID = mstr����ID & "," & Mid(tvw����.Nodes(i).Key, 2)
            End If
        Next
        If mstr����ID <> "" Then
                mstr����ID = Mid(mstr����ID, 2)
        End If
    End If
    
    'ȡ�ÿⷿ��λ����ѡ�ⷿ��ʾ���пⷿ��
    mstr�ⷿ��λ = ""
    For intItem = 1 To tvw��λ.Nodes.Count
        If tvw��λ.Nodes(intItem).Key <> "Root" Then
            If tvw��λ.Nodes(intItem).Checked Then
                mstr�ⷿ��λ = mstr�ⷿ��λ & "," & tvw��λ.Nodes(intItem).Text
            End If
        End If
        
'        If tvw��λ.Nodes(intItem).Key = "Root" And tvw��λ.Nodes(intItem).Checked = True Then
'            mstr�ⷿ��λ = ""
'            Exit For
'        ElseIf tvw��λ.Nodes(intItem).Checked Then
'            mstr�ⷿ��λ = mstr�ⷿ��λ & "," & tvw��λ.Nodes(intItem).Text
'        End If
    Next
    
    If mstr�ⷿ��λ <> "" Then
        mstr�ⷿ��λ = Mid(mstr�ⷿ��λ, 2)
    End If
    
    mlng�ⷿid = cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)
    mlng�̵㷽ʽ = Cbo�̵㷽ʽ.ItemData(Cbo�̵㷽ʽ.ListIndex)
    mstr�̵�ʱ�� = Format(dtpDate.Value, "yyyy-MM-dd hh:mm:ss")
    mblnSelect = True
    mint���޿����� = chkNum.Value
    mbln�̵����������н�� = (chkNoNum.Value = 1)
    
    frmCheckCard.txtStock.Caption = cbo�ⷿ.Text
    frmCheckCard.txtStock.Tag = mlng�ⷿid
    frmCheckCard.txtCheckDate = mstr�̵�ʱ��
    frmCheckCard.CmdSave.Enabled = False
    frmCheckCard.CmdCancel.Enabled = False
    
    Unload Me
End Sub

Private Sub Command1_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub Form_Activate()
    If mblnBootUp = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim strSelectStock As String
    
    On Error GoTo ErrHandle
    strSelectStock = IIf(Val(zlDatabase.GetPara("�Ƿ�ѡ��ⷿ", glngSys, mlngModule, "0")) = 1, 1, 0)
    '���Ĳ���Ȩ�޿���
    
    dtpDate.Value = Format(sys.Currentdate, dtpDate.CustomFormat)
    dtpDate.MaxDate = dtpDate.Value
    
    mblnBootUp = False

    With Cbo�̵㷽ʽ
        .Clear
        .AddItem "ÿ��"
        .ItemData(.NewIndex) = 1
        .AddItem "ÿ��"
        .ItemData(.NewIndex) = 2
        .AddItem "ÿ��"
        .ItemData(.NewIndex) = 3
        .AddItem "ÿ����"
        .ItemData(.NewIndex) = 4
        .AddItem "�����̵㷽ʽ"
        .ItemData(.NewIndex) = 5
        .ListIndex = 0
    End With
    
    With mfrmMain.cboStock
        cbo�ⷿ.Clear
        For i = 0 To .ListCount - 1
            cbo�ⷿ.AddItem .List(i)
            cbo�ⷿ.ItemData(cbo�ⷿ.NewIndex) = .ItemData(i)
        Next
        cbo�ⷿ.ListIndex = .ListIndex
    End With
        
    If InStr(1, gstrPrivs, "���пⷿ") <> 0 Then
        If strSelectStock = "0" Then
            cbo�ⷿ.Enabled = False
        Else
            cbo�ⷿ.Enabled = True
        End If
    Else
        cbo�ⷿ.Enabled = False
    End If
    
    With rsTemp
        gstrSQL = "Select ����,���� From ���Ʒ���Ŀ¼ where ����=7 order by ���� "
        Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "���ķ���")
        
        If .EOF Then
            MsgBox "���ķ��಻������", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    
'��;
    gstrSQL = "" & _
        "   Select Level as ��,ID,�ϼ�ID,���� From ���Ʒ���Ŀ¼ where ����=7" & _
        "   Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
        "   Order by ��"
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    If rsTemp.RecordCount = 0 Then
        MsgBox "���ķ��಻������", vbInformation, gstrSysName
        Exit Sub
    End If

    Dim objNode As Node
    Set objNode = tvw����.Nodes.Add(, , "Root", "�������ķ���", "Item")
    
    Do While Not rsTemp.EOF
        If rsTemp!�� = 1 Then
            Set objNode = tvw����.Nodes.Add("Root", 4, "_" & rsTemp!Id, rsTemp!����, "Item")
        Else
            Set objNode = tvw����.Nodes.Add("_" & rsTemp!�ϼ�ID, 4, "_" & rsTemp!Id, rsTemp!����, "Item")
        End If
        rsTemp.MoveNext
    Loop
    tvw����.Nodes("Root").Selected = True
    tvw����.Nodes("Root").Expanded = True
    mblnBootUp = True
    
    '�ⷿ��λ
    Load�ⷿ��λ
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckCount() As Integer
    Dim i As Integer
    For i = 1 To tvw����.Nodes.Count
        If tvw����.Nodes(i).Checked Then CheckCount = CheckCount + 1
    Next
End Function

Private Sub Load�ⷿ��λ()
    Dim rsTemp As ADODB.Recordset
    Dim objNode As Node
    
    On Error GoTo ErrHandle
    '��ȡ���в��Ͽⷿ��λ
    If chk��λ.Value = 1 Then
        gstrSQL = "Select Distinct B.����, B.����" & _
            " From ���ϴ����޶� A, ���Ͽⷿ��λ B " & _
            " Where A.�ⷿ��λ = B.���� And A.�ⷿid = [1] " & _
            " Order By B.����"
    Else
        gstrSQL = "Select ����,���� From ���Ͽⷿ��λ Order By ���� "
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���в��Ͽⷿ��λ", Val(cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)))
    
    tvw��λ.Nodes.Clear
    Set objNode = tvw��λ.Nodes.Add(, , "Root", "���пⷿ", 1)
    Do While Not rsTemp.EOF
        Set objNode = tvw��λ.Nodes.Add("Root", 4, "_" & rsTemp!����, rsTemp!����, 1)

        rsTemp.MoveNext
    Loop
    tvw��λ.Nodes("Root").Selected = True
    tvw��λ.Nodes("Root").Expanded = True
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub tvw����_NodeCheck(ByVal Node As MSComctlLib.Node)
    CheckNode Node, Node.Checked
    SetParentNode tvw����, Node, Node.Checked
End Sub

Private Sub tvw��λ_NodeCheck(ByVal Node As MSComctlLib.Node)
    CheckNode Node, Node.Checked
    SetParentNode tvw��λ, Node, Node.Checked
End Sub


Private Function CheckNode(ByVal Node As Object, blnCheck As Boolean)
    Dim intIdx As Integer

    If Node.Children > 0 Then
        Set Node = Node.Child
        Do While Not Node Is Nothing
            Node.Checked = blnCheck
            If Node.Children > 0 Then
                CheckNode Node, blnCheck
            End If
            Set Node = Node.Next
        Loop
    Else
        Node.Checked = blnCheck
    End If
End Function

Private Sub SetParentNode(ByVal objMyTreeView As TreeView, ByVal Node As MSComctlLib.Node, blnCheck As Boolean)
    Dim intIdx As Integer
    
    If Not Node.Parent Is Nothing Then
        If blnCheck = True Then
            '���Ƿ������ֵܽӵ��Ƿ�Ҳȫ��TRUE�����ǣ������丸�ڵ�ҲΪTRUE�����򣬲���
            intIdx = Node.FirstSibling.Index
            Do While intIdx <> Node.LastSibling.Index
                If objMyTreeView.Nodes(intIdx).Checked = False Then
                    Node.Parent.Checked = False
                    Exit Do
                End If
                intIdx = objMyTreeView.Nodes(intIdx).Next.Index
            Loop
            If intIdx = Node.LastSibling.Index Then
                If objMyTreeView.Nodes(intIdx).Checked = True Then
                    Node.Parent.Checked = True
                End If
            End If
        Else
            Node.Parent.Checked = False
        End If
        
        Set Node = Node.Parent
        If Not Node Is Nothing Then
            SetParentNode objMyTreeView, Node, blnCheck
        End If
    End If
End Sub
