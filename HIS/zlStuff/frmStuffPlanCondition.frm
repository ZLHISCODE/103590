VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmStuffPlanCondition 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   4920
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7560
   Icon            =   "frmStuffPlanCondition.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin MSComctlLib.ImageList img16 
      Left            =   6360
      Top             =   1245
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCondition.frx":000C
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCondition.frx":0EE6
            Key             =   "Folder1"
            Object.Tag             =   "Folder1"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCondition.frx":1338
            Key             =   "Card"
            Object.Tag             =   "Card"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPlanCondition.frx":178A
            Key             =   "Folder"
            Object.Tag             =   "Folder"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   6360
      TabIndex        =   22
      Top             =   4440
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6330
      TabIndex        =   21
      Top             =   855
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6330
      TabIndex        =   20
      Top             =   435
      Width           =   1100
   End
   Begin TabDlg.SSTab stb 
      Height          =   4725
      Left            =   45
      TabIndex        =   24
      Top             =   105
      Width           =   6045
      _ExtentX        =   10668
      _ExtentY        =   8340
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "�ƻ�(&J)"
      TabPicture(0)   =   "frmStuffPlanCondition.frx":1BDC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl�ⷿ"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra��������"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra�ƻ�����"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra�ƻ�����"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cbo�ⷿ"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Chk����ȡ��ȡ���޵Ĳ���"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Chk�������ƻ�����"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "��;����(Y)"
      TabPicture(1)   =   "frmStuffPlanCondition.frx":1BF8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tvw��;"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "��Ӧ��(&F)"
      TabPicture(2)   =   "frmStuffPlanCondition.frx":1C14
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chk�б굥λ"
      Tab(2).Control(1)=   "tvw������λ"
      Tab(2).ControlCount=   2
      Begin VB.CheckBox Chk�������ƻ����� 
         Appearance      =   0  'Flat
         Caption         =   "�������ƻ�����"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   192
         TabIndex        =   25
         Top             =   600
         Width           =   1560
      End
      Begin VB.CheckBox chk�б굥λ 
         Caption         =   "���ϴι�Ӧ�����б굥λΪ׼(&W)"
         Enabled         =   0   'False
         Height          =   240
         Left            =   -74520
         TabIndex        =   19
         Top             =   4305
         Width           =   2985
      End
      Begin VB.CheckBox Chk����ȡ��ȡ���޵Ĳ��� 
         Appearance      =   0  'Flat
         Caption         =   "����ȡ�������޵Ĳ���(&Q)"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   924
         TabIndex        =   16
         Top             =   4245
         Width           =   2895
      End
      Begin VB.ComboBox cbo�ⷿ 
         Height          =   276
         Left            =   924
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   3795
         Width           =   4848
      End
      Begin VB.Frame fra�ƻ����� 
         Caption         =   "���Ʒ���"
         Height          =   1695
         Left            =   192
         TabIndex        =   4
         Top             =   1875
         Width           =   2640
         Begin VB.OptionButton opt���� 
            Caption         =   "�����깺���շ�(&5)"
            Height          =   270
            Index           =   4
            Left            =   240
            TabIndex        =   9
            Top             =   1230
            Width           =   2370
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "���������������շ�(&4)"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   8
            Top             =   1020
            Width           =   2190
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "���ϴ���������շ�(&3)"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   7
            Top             =   795
            Width           =   2190
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "�ٽ��ڼ�ƽ�����շ�(&2)"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   6
            Top             =   570
            Width           =   2190
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "����ͬ�����Բ��շ�(&1)"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   5
            Top             =   330
            Value           =   -1  'True
            Width           =   2190
         End
      End
      Begin VB.Frame fra�ƻ����� 
         Caption         =   "�ƻ�����"
         Height          =   765
         Left            =   192
         TabIndex        =   0
         Top             =   960
         Width           =   5580
         Begin VB.OptionButton opt�ƻ� 
            Caption         =   "�ܶȼƻ�(&W)"
            Height          =   210
            Index           =   3
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Value           =   -1  'True
            Width           =   1296
         End
         Begin VB.OptionButton opt�ƻ� 
            Caption         =   "��ȼƻ�(&Y)"
            Enabled         =   0   'False
            Height          =   210
            Index           =   2
            Left            =   4116
            TabIndex        =   3
            Top             =   375
            Width           =   1290
         End
         Begin VB.OptionButton opt�ƻ� 
            Caption         =   "���ȼƻ�(&B)"
            Height          =   210
            Index           =   1
            Left            =   2786
            TabIndex        =   2
            Top             =   375
            Width           =   1290
         End
         Begin VB.OptionButton opt�ƻ� 
            Caption         =   "�¶ȼƻ�(&A)"
            Height          =   210
            Index           =   0
            Left            =   1456
            TabIndex        =   1
            Top             =   375
            Width           =   1290
         End
      End
      Begin VB.Frame fra�������� 
         Caption         =   "��������"
         Enabled         =   0   'False
         Height          =   1680
         Left            =   3387
         TabIndex        =   23
         Top             =   1875
         Width           =   2385
         Begin VB.TextBox txt�������� 
            Height          =   300
            Left            =   1185
            TabIndex        =   13
            Top             =   885
            Width           =   900
         End
         Begin VB.TextBox txt�������� 
            Height          =   300
            Left            =   1185
            TabIndex        =   11
            Top             =   525
            Width           =   900
         End
         Begin VB.Label lbl�������� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��������(&T)"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   150
            TabIndex        =   12
            Top             =   945
            Width           =   990
         End
         Begin VB.Label lbl�������� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��������(&X)"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   150
            TabIndex        =   10
            Top             =   585
            Width           =   990
         End
      End
      Begin MSComctlLib.TreeView tvw������λ 
         Height          =   3780
         Left            =   -74925
         TabIndex        =   18
         Top             =   465
         Width           =   5805
         _ExtentX        =   10245
         _ExtentY        =   6668
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
      Begin MSComctlLib.TreeView tvw��; 
         Height          =   4230
         Left            =   -74925
         TabIndex        =   17
         Top             =   420
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   7451
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
      Begin VB.Label lbl�ⷿ 
         AutoSize        =   -1  'True
         Caption         =   "�ⷿ(&K)"
         Height          =   180
         Left            =   192
         TabIndex        =   14
         Top             =   3840
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmStuffPlanCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnSelect As Boolean

Private mstr����ID As String
Private mstr������ID As String
Private mbln�б굥λ As Boolean
Private mlng�ⷿid As Long
Private mint�ƻ����� As Integer
Private mint���Ʒ��� As Integer
Private mbln���� As Boolean
Private mint���� As Integer
Private mint���� As Integer
Private mfrmMain As Form
Private mbln�ƻ����� As Boolean
Private Const mlngModule = 1724

Public Function GetCondition(frmMain As Form, ByRef str����ID, ByRef lng�ⷿID As Long, _
    ByRef int�ƻ����� As Integer, ByRef int���Ʒ��� As Integer, ByRef bln���� As Boolean, _
    ByRef int���� As Integer, ByRef int���� As Integer, _
    ByRef str������ID As String, ByRef bln�б굥λ As Boolean, ByRef bln�ƻ����� As Boolean) As Boolean
    
    mstr����ID = ""
    mlng�ⷿid = 0
    mint�ƻ����� = 0
    mint���Ʒ��� = 0
    mblnSelect = False
    
    Set mfrmMain = frmMain
    Me.Show vbModal, frmMain
    GetCondition = mblnSelect
    
    bln�б굥λ = mbln�б굥λ
    str����ID = mstr����ID
    str������ID = mstr������ID
    lng�ⷿID = mlng�ⷿid
    int�ƻ����� = mint�ƻ�����
    int���Ʒ��� = mint���Ʒ���
    bln���� = mbln����
    int���� = mint����
    int���� = mint����
    bln�ƻ����� = mbln�ƻ�����
End Function

Private Sub cbo�ⷿ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub Chk����ȡ��ȡ���޵Ĳ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        stb.Tab = 1
        If tvw��;.Enabled And tvw��;.Visible Then
            tvw��;.SetFocus
        Else
            OS.PressKey vbKeyTab
        End If
    End If
End Sub

Private Sub chk�б굥λ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnSelect = False
    Hide
    Unload Me
End Sub


Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name)
End Sub
Private Function ISValid() As Boolean
    '��֤����
    ISValid = False
    
    If opt����(3).Value Then
        '���������������С�ڿ����������
        '���������������������������Ϊ��
        If Trim(txt��������.Text) = "" Then
            MsgBox "������������������", vbInformation, gstrSysName
            txt��������.SetFocus
            Exit Function
        End If
        If Trim(txt��������.Text) = "" Then
            MsgBox "������������������", vbInformation, gstrSysName
            txt��������.SetFocus
            Exit Function
        End If
        If Not IsNumeric(txt��������.Text) Then
            MsgBox "������������к��зǷ��ַ���", vbInformation, gstrSysName
            txt��������.SetFocus
            Exit Function
        End If
        If Not IsNumeric(txt��������.Text) Then
            MsgBox "������������к��зǷ��ַ���", vbInformation, gstrSysName
            txt��������.SetFocus
            Exit Function
        End If
        If Val(txt��������.Text) <= 0 Then
            MsgBox "���������������С���㣡", vbInformation, gstrSysName
            txt��������.SetFocus
            Exit Function
        End If
        If Val(txt��������.Text) <= 0 Then
            MsgBox "���������������С���㣡", vbInformation, gstrSysName
            txt��������.SetFocus
            Exit Function
        End If
        If Val(txt��������.Text) < Val(txt��������.Text) Then
            MsgBox "���������������С�ڿ������������", vbInformation, gstrSysName
            txt��������.SetFocus
            Exit Function
        End If
        If Val(txt��������.Text) > 300 Then
            MsgBox "��������������ܴ���300�죡", vbInformation, gstrSysName
            txt��������.SetFocus
            Exit Function
        End If
    End If
    ISValid = True
    
End Function
Private Sub cmdOK_Click()
    Dim intIndex As Integer
    Dim i As Integer
    Dim Str�ڼ� As String
    Dim intMonth As Integer
    
    If ISValid() = False Then Exit Sub
    
    mstr����ID = ""
    For i = 1 To tvw��;.Nodes.Count
        If tvw��;.Nodes(i).Key <> "Root" And _
            tvw��;.Nodes(i).Checked Then
            mstr����ID = mstr����ID & "," & Mid(tvw��;.Nodes(i).Key, 2)
        End If
    Next
    mstr������ID = ""
    For i = 1 To tvw������λ.Nodes.Count
        If tvw������λ.Nodes(i).Key <> "Root" And _
            tvw������λ.Nodes(i).Checked Then
            If tvw������λ.Nodes(i).Tag = "1" Then
                mstr������ID = mstr������ID & "," & Mid(tvw������λ.Nodes(i).Key, 2)
            End If
        End If
    Next
    mint���� = Val(txt��������.Text)
    mint���� = Val(txt��������.Text)
    
    If mstr������ID <> "" Then mstr������ID = Mid(mstr������ID, 2)
    If chk�б굥λ.Value = 1 And chk�б굥λ.Enabled Then
        If mstr������ID = "" Then
            ShowMsgBox "��û��ѡ�񹩻���λʱ������ѡ�����ϴι�Ӧ�����б굥λΪ׼��"
            Me.stb.Tab = 2
            If tvw������λ.Enabled Then tvw������λ.SetFocus
            Exit Sub
        End If
    End If
    
    mbln�б굥λ = chk�б굥λ.Value = 1
    
    If mbln�б굥λ And mstr������ID = "" Then
        ShowMsgBox "δѡ��Ӧ��,�������á����ϴι�Ӧ�����б굥λΪ׼��"
        stb.Tab = 2
        If chk�б굥λ.Enabled And chk�б굥λ.Visible Then chk�б굥λ.SetFocus
        Exit Sub
    End If
    
    If mstr����ID <> "" Then
        mstr����ID = Mid(mstr����ID, 2)
    End If
    
    mlng�ⷿid = cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)
    frmStuffPlanCard.LblTitle.Tag = cbo�ⷿ.Text
    
    For i = 0 To opt�ƻ�.Count - 1
       If opt�ƻ�(i).Value Then
           frmStuffPlanCard.txt�ƻ�����.Caption = Mid(opt�ƻ�(i).Caption, 1, InStr(1, opt�ƻ�(i).Caption, "(") - 1)
           mint�ƻ����� = i + 1
           Exit For
       End If
    Next
    
    For i = 0 To opt����.Count - 1
       If opt����(i).Value Then
           frmStuffPlanCard.txt���Ʒ���.Caption = Mid(opt����(i).Caption, 1, InStr(1, opt����(i).Caption, "(") - 1)
           mint���Ʒ��� = i + 1
           Exit For
       End If
    Next
    mbln���� = (Chk����ȡ��ȡ���޵Ĳ���.Value = 1)
    mbln�ƻ����� = (Chk�������ƻ�����.Value <> 1)
    
    mblnSelect = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If tvw��;.Visible Then
            tvw��;.Visible = False
        Else
            cmdCancel_Click
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim strReg As String
    
    Dim rsTemp As New ADODB.Recordset
    Dim objNode As Node
    Dim strIco As String
    Dim i As Integer
    Dim strSelectStock As String
    
    On Error GoTo errH
    strReg = IIf(Val(zlDatabase.GetPara("�Ƿ�ѡ��ⷿ", glngSys, mlngModule, "0")) = 1, 1, 0)
    strSelectStock = Val(strReg)
    
    stb.Tab = 0
    Call opt����_Click(0)
    
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
        

    gstrSQL = "" & _
        "   Select Level as ��,ID,�ϼ�ID,���� " & _
        "   From ���Ʒ���Ŀ¼" & _
        "   where ����=7 " & _
        "   Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
        "   Order by Level"
        
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    Set objNode = tvw��;.Nodes.Add(, , "Root", "���в��Ϸ���", "Item")
    
    Do While Not rsTemp.EOF
        If rsTemp!�� = 1 Then
            Set objNode = tvw��;.Nodes.Add("Root", 4, "_" & rsTemp!Id, rsTemp!����, "Item")
        Else
            Set objNode = tvw��;.Nodes.Add("_" & rsTemp!�ϼ�ID, 4, "_" & rsTemp!Id, rsTemp!����, "Item")
        End If
        rsTemp.MoveNext
    Loop
    tvw��;.Nodes("Root").Selected = True
    tvw��;.Nodes("Root").Expanded = True
    
    gstrSQL = "" & _
        "   Select Level as ��,ID,�ϼ�ID,����||'-'||���� ����,ĩ�� " & _
        "   From ��Ӧ��" & _
        "   where (substr(����,5,1)=1 and (վ��=[1] or վ�� is null) Or Nvl(ĩ��,0)=0) And (To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' or ����ʱ�� is null)" & _
        "   Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
        "   Order by Level"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, gstrNodeNo)
    
    tvw������λ.Nodes.Clear
    Set objNode = tvw������λ.Nodes.Add(, , "Root", "�������Ĺ�����", "Folder")
    objNode.Sorted = True
    Do While Not rsTemp.EOF
        strIco = IIf(Val(NVL(rsTemp!ĩ��)) = 1, "Card", "Folder")
        If rsTemp!�� = 1 Then
            Set objNode = tvw������λ.Nodes.Add("Root", 4, "_" & rsTemp!Id, rsTemp!����, strIco)
        Else
            Set objNode = tvw������λ.Nodes.Add("_" & rsTemp!�ϼ�ID, 4, "_" & rsTemp!Id, rsTemp!����, strIco)
        End If
        If strIco = "Card" Then
            objNode.Tag = "1"
        End If
        objNode.Sorted = True
        rsTemp.MoveNext
    Loop
    tvw������λ.Nodes("Root").Selected = True
    tvw������λ.Nodes("Root").Expanded = True
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Sub
Private Sub opt����_Click(Index As Integer)
    fra��������.Enabled = False
    Chk����ȡ��ȡ���޵Ĳ���.Enabled = True
    tvw������λ.Enabled = True
    chk�б굥λ.Enabled = True
    If Index = 0 Then
        If opt�ƻ�(2).Value = True Then
            opt�ƻ�(2).Value = False
            opt�ƻ�(0).Value = True
        End If
        If opt�ƻ�(3).Value = True Then
            opt�ƻ�(3).Value = False
            opt�ƻ�(0).Value = True
        End If
        opt�ƻ�(2).Enabled = False
        opt�ƻ�(3).Enabled = False
    ElseIf Index = 3 Then
        fra��������.Enabled = True
'        opt�ƻ�(0).Value = True
'        opt�ƻ�(1).Value = False
'        opt�ƻ�(2).Value = False
        opt�ƻ�(2).Enabled = False
        opt�ƻ�(3).Enabled = False
        If opt�ƻ�(2).Value = True Then
            opt�ƻ�(2).Value = False
            opt�ƻ�(0).Value = True
        End If
        If opt�ƻ�(1).Value = True Then
            opt�ƻ�(1).Value = False
            opt�ƻ�(0).Value = True
        End If
        If opt�ƻ�(3).Value = True Then
            opt�ƻ�(3).Value = False
            opt�ƻ�(0).Value = True
        End If
        
    ElseIf Index = 4 Then
        '���ݲ����깺���Ƽƻ�
        fra��������.Enabled = False
        Chk����ȡ��ȡ���޵Ĳ���.Enabled = False
        tvw������λ.Enabled = False
        chk�б굥λ.Enabled = False
        opt�ƻ�(2).Enabled = True
        opt�ƻ�(3).Enabled = False
        If opt�ƻ�(3).Value = True Then
            opt�ƻ�(3).Value = False
            opt�ƻ�(0).Value = True
        End If
    Else
        opt�ƻ�(2).Enabled = True
        opt�ƻ�(3).Enabled = True
    End If
End Sub

Private Sub opt����_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub opt�ƻ�_Click(Index As Integer)
    If opt�ƻ�(0).Value = False Then
        If opt����(3).Value Then
            opt����(0).Value = True
        End If
    End If
End Sub

Private Sub opt�ƻ�_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub



Private Sub stb_Click(PreviousTab As Integer)
    Select Case stb.Tab
    Case 1
       If tvw��;.Visible And tvw��;.Enabled Then tvw��;.SetFocus
    Case 2
       If tvw������λ.Visible And tvw������λ.Enabled Then tvw������λ.SetFocus
    End Select
    
End Sub
 
Private Sub tvw������λ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub tvw������λ_NodeCheck(ByVal Node As MSComctlLib.Node)
    CheckNode Node, Node.Checked
    SetParentNode Node, Node.Checked, False
    chk�б굥λ.Enabled = IsHavingCheck(Node)
End Sub
Private Function IsHavingCheck(ByVal objNode As Node) As Boolean
    '����:����Ƿ����Node��ѡ���˵�
    Dim objNode1 As Node
    If Not objNode Is Nothing Then
        If objNode.Checked = True Then IsHavingCheck = True: Exit Function
    End If
    For Each objNode1 In tvw������λ.Nodes
        If objNode1.Checked Then IsHavingCheck = True: Exit Function
    Next
End Function
Private Sub tvw��;_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        stb.Tab = 2
        If tvw������λ.Enabled And tvw������λ.Visible Then
            tvw������λ.SetFocus
        Else
            OS.PressKey vbKeyTab
        End If

    End If
End Sub

Private Sub tvw��;_NodeCheck(ByVal Node As MSComctlLib.Node)
    CheckNode Node, Node.Checked
    SetParentNode Node, Node.Checked
End Sub

Private Sub SetParentNode(ByVal Node As MSComctlLib.Node, blnCheck As Boolean, Optional blnTvw��; As Boolean = True)
    Dim intIdx As Integer
    
    If Not Node.Parent Is Nothing Then
        If blnCheck = True Then
            '���Ƿ������ֵܽӵ��Ƿ�Ҳȫ��TRUE�����ǣ������丸�ڵ�ҲΪTRUE�����򣬲���
            intIdx = Node.FirstSibling.Index
            Do While intIdx <> Node.LastSibling.Index
                If blnTvw��; = True Then
                    If tvw��;.Nodes(intIdx).Checked = False Then
                        Node.Parent.Checked = False
                        Exit Do
                    End If
                    intIdx = tvw��;.Nodes(intIdx).Next.Index
                Else
                    If tvw������λ.Nodes(intIdx).Checked = False Then
                        Node.Parent.Checked = False
                        Exit Do
                    End If
                    intIdx = tvw������λ.Nodes(intIdx).Next.Index
                End If
            Loop
            If intIdx = Node.LastSibling.Index Then
                If blnTvw��; = True Then
                       If tvw������λ.Nodes(intIdx).Checked = True Then
                           Node.Parent.Checked = True
                       End If
                Else
                       If tvw������λ.Nodes(intIdx).Checked = True Then
                           Node.Parent.Checked = True
                       End If
                End If
            End If
        Else
            Node.Parent.Checked = False
        End If
        
        Set Node = Node.Parent
        If Not Node Is Nothing Then
            SetParentNode Node, blnCheck, blnTvw��;
        End If
    End If
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

Private Function CheckCount() As Integer
    Dim i As Integer
    For i = 1 To tvw��;.Nodes.Count
        If tvw��;.Nodes(i).Checked Then CheckCount = CheckCount + 1
    Next
End Function
Private Sub txt��������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txt��������_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt��������, KeyAscii, m����ʽ
End Sub
Private Sub txt��������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txt��������_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt��������, KeyAscii, m����ʽ
End Sub
