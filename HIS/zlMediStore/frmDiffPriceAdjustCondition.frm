VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDiffPriceAdjustCondition 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�Զ���������"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command1 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4920
      TabIndex        =   16
      Top             =   4200
      Width           =   1100
   End
   Begin VB.Frame fraRangeSelect 
      Caption         =   " ���� "
      Height          =   5190
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   4680
      Begin VB.CheckBox Chk���� 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "ȫѡ"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3600
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1080
         Width           =   675
      End
      Begin MSComCtl2.UpDown updRate 
         Height          =   300
         Left            =   3225
         TabIndex        =   11
         Top             =   3825
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtRate"
         BuddyDispid     =   196613
         OrigLeft        =   3720
         OrigTop         =   4200
         OrigRight       =   3960
         OrigBottom      =   4575
         Max             =   100
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtRate 
         Height          =   300
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   10
         Text            =   "10"
         Top             =   3825
         Width           =   1935
      End
      Begin VB.CommandButton Cmd��; 
         Caption         =   "��"
         Height          =   300
         Left            =   3885
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   675
         Width           =   285
      End
      Begin VB.TextBox Txt��; 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1110
         MaxLength       =   50
         TabIndex        =   4
         Top             =   675
         Width           =   2775
      End
      Begin VB.ComboBox cbo�ⷿ 
         Height          =   300
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   285
         Width           =   3045
      End
      Begin MSComctlLib.ListView Lvw���� 
         Height          =   2430
         Left            =   345
         TabIndex        =   8
         Top             =   1290
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   4286
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&G)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   405
         TabIndex        =   6
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "  ˵����ʵ�ʲ����ʵ�ʽ��֮�ȴ��ڻ�С��ָ������ʵİٷֵ�Ϊ������ʵ���ЩҩƷ�ų�����"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   330
         TabIndex        =   13
         Top             =   4245
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3570
         TabIndex        =   12
         Top             =   3885
         Width           =   255
      End
      Begin VB.Label Lbl�̵㷽ʽ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   345
         TabIndex        =   9
         Top             =   3885
         Width           =   900
      End
      Begin VB.Label Lbl��;���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��;����"
         Height          =   180
         Left            =   315
         TabIndex        =   3
         Top             =   735
         Width           =   720
      End
      Begin VB.Label lbl�ⷿ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ⷿ"
         Height          =   180
         Left            =   675
         TabIndex        =   1
         Top             =   345
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4920
      TabIndex        =   15
      Top             =   720
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4920
      TabIndex        =   14
      Top             =   285
      Width           =   1100
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   5280
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCondition.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCondition.frx":0E52
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCondition.frx":2B5C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView Tvw��;���� 
      Height          =   2700
      Left            =   510
      TabIndex        =   17
      Top             =   1095
      Visible         =   0   'False
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   4763
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      BorderStyle     =   1
      Appearance      =   0
   End
End
Attribute VB_Name = "frmDiffPriceAdjustCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnSelect As Boolean
Private mblnBootUp As Boolean
Private mblnFirstUp As Boolean

Private mstr��;ID As String
Private mstr���� As String
Private mlng�ⷿID As Long
Private mintRate As Integer
Private mbln��ҩ�ⷿ As Boolean
Private mfrmMain As Form

Public Function GetCondition(FrmMain As Form, ByRef str��;ID, ByRef str���� As String, _
    ByRef lng�ⷿID As Long, ByRef int������ As Integer) As Boolean
    
    mstr��;ID = ""
    mstr���� = ""
    mlng�ⷿID = 0
    mintRate = int������
    mblnSelect = False
    
    Set mfrmMain = FrmMain
    Me.Show vbModal, FrmMain
    GetCondition = mblnSelect
    str��;ID = mstr��;ID
    str���� = mstr����
    lng�ⷿID = mlng�ⷿID
    int������ = mintRate
End Function

Private Sub cbo�ⷿ_Click()
    Dim rsTemp As New ADODB.Recordset
    '��ȡ�ÿⷿ���м��ͣ����û�ѡ��
    mbln��ҩ�ⷿ = False
    On Error GoTo errHandle
    
    gstrSQL = "Select 1 From ��������˵�� " & _
             " Where �������� Like '��ҩ%' And ����ID=[1]"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[��鲿������]", Me.cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex))
             
    If Not rsTemp.EOF Then mbln��ҩ�ⷿ = True
    
    gstrSQL = "Select Distinct J.����,J.���� " & _
             " From ����ִ�п��� A,ҩƷ���� B,ҩƷ���� J " & _
             " Where A.������ĿID=B.ҩ��ID And B.ҩƷ����=J.���� " & _
             "     And A.ִ�п���ID=[1]"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ�ÿⷿ���ڼ���]", Me.cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex))
    
    lvw����.ListItems.Clear
    
    With rsTemp
        Do While Not .EOF
            lvw����.ListItems.Add , "K" & !����, !����, , 1
            .MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Chk����_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnSelect = False
    
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim intItem As Integer, intItems As Integer
    Dim rsTmp As ADODB.Recordset
    Dim strsql As String
    
    On Error GoTo errHandle
    If Tvw��;����.SelectedItem.Key <> "R" Then
        Select Case Tvw��;����.SelectedItem.Key
            Case "R_�г�ҩ", "R_�в�ҩ", "R_����ҩ"
                mstr��;ID = "'" & Tvw��;����.SelectedItem & "'"
            Case Else
                strsql = "Select ID From ���Ʒ���Ŀ¼ " & _
                         "Start With ID=[1] " & _
                         "Connect by Prior ID=�ϼ�ID "
                Set rsTmp = zlDataBase.OpenSQLRecord(strsql, Me.Caption & "-���Ʒ���ID", Mid(Tvw��;����.SelectedItem.Key, 3))
                With rsTmp
                    mstr��;ID = ""
                    Do While Not .EOF
                        mstr��;ID = mstr��;ID & !id & ","
                        .MoveNext
                    Loop
                End With
                mstr��;ID = Mid(mstr��;ID, 1, Len(mstr��;ID) - 1)
        End Select
    End If
    
    mstr���� = ""
    intItems = Me.lvw����.ListItems.count
    For intItem = 1 To intItems
        If lvw����.ListItems(intItem).Checked Then
            mstr���� = mstr���� & "," & lvw����.ListItems(intItem).Text
        End If
    Next
    If mbln��ҩ�ⷿ Then mstr���� = mstr���� & "," & "����"
    If mstr���� <> "" Then mstr���� = Mid(mstr����, 2)
    
    mlng�ⷿID = cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)
    mintRate = Val(txtRate.Text)
    
    mblnSelect = True
    frmDiffPriceAdjustCard.txtStock.Caption = cbo�ⷿ.Text
    frmDiffPriceAdjustCard.txtStock.Tag = mlng�ⷿID
    
    frmDiffPriceAdjustCard.CmdSave.Enabled = False
    frmDiffPriceAdjustCard.cmdCancel.Enabled = False
    
    Hide
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Cmd��;_Click()
    '��ҩƷ��;����װ��TREEVIEW
    Tvw��;����.Visible = Tvw��;����.Visible Xor True
    If Tvw��;����.Visible Then
        Tvw��;����.Top = Txt��;.Top + Txt��;.Height + fraRangeSelect.Top
        Tvw��;����.Left = Txt��;.Left + fraRangeSelect.Left
        Tvw��;����.ZOrder 0
        Tvw��;����.SetFocus
    End If
End Sub


Private Sub Command1_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub Form_Click()
    If Tvw��;����.Visible = True Then
        Tvw��;����.Visible = False
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim rs��;���� As New Recordset
    Dim rs���� As New Recordset
    Dim rs���ʷ��� As New Recordset
    Dim Str���� As String
    
    Dim blnSelectStock As String
    On Error GoTo errHandle
    blnSelectStock = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & mfrmMain.Caption, "�ⷿ", "0")
    
    mblnBootUp = False
    mblnFirstUp = True
    
    With mfrmMain.cboStock
        cbo�ⷿ.Clear
        For i = 0 To .ListCount - 1
            cbo�ⷿ.AddItem .List(i)
            cbo�ⷿ.ItemData(cbo�ⷿ.NewIndex) = .ItemData(i)
        Next
        cbo�ⷿ.ListIndex = .ListIndex
    End With
        
    If zlStr.IsHavePrivs(gstrprivs, "���пⷿ") Then
        If blnSelectStock = "0" Then
            cbo�ⷿ.Enabled = False
        Else
            cbo�ⷿ.Enabled = True
        End If
    Else
        cbo�ⷿ.Enabled = False
    End If
        
    'ҩƷ����Ȩ�޿���
    Str���� = ""
    If UserInfo.strMaterial <> "" Then      'Ϊ�ձ�ʾ�����пⷿȨ��
        If InStr(1, UserInfo.strMaterial, "�г�ҩ") <> 0 Then Str���� = Str���� & IIf(Str���� = "", "", ",") & "2"
        If InStr(1, UserInfo.strMaterial, "����ҩ") <> 0 Then Str���� = Str���� & IIf(Str���� = "", "", ",") & "1"
        If InStr(1, UserInfo.strMaterial, "�в�ҩ") <> 0 Then Str���� = Str���� & IIf(Str���� = "", "", ",") & "3"
        If Str���� = "" Then
            MsgBox "�Բ�����һ�����ʷ���Ȩ�޶�û�У�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
            Exit Sub
        End If
    Else
        Str���� = "1,2,3"
    End If
    
    gstrSQL = " SELECT a.ID,a.�ϼ�ID,a.����,1 AS ĩ��,DECODE(a.����,1,'����ҩ',2,'�г�ҩ','�в�ҩ') AS ����  " & _
              " FROM ���Ʒ���Ŀ¼ a " & _
              IIf(Str���� = "", "", " WHERE a.���� in (select * from Table(Cast(f_Num2List([1]) As zlTools.t_NumList))) ") & _
              " START WITH a.�ϼ�ID IS NULL CONNECT BY PRIOR a.ID =a.�ϼ�ID ORDER BY LEVEL,a.ID "
    Set rs��;���� = zlDataBase.OpenSQLRecord(gstrSQL, "��ȡ��;����", Str����)
    
    With rs��;����
        If .EOF Then
            MsgBox "ҩƷ��;���಻������", vbInformation, gstrSysName
            Exit Sub
        End If
        
        With rs��;����
            Tvw��;����.Nodes.Clear
            Tvw��;����.Nodes.Add , , "R", "������;����", 1, 1
            Txt��;.Text = "������;����"
            
            
            gstrSQL = "Select ���� From ������Ŀ��� Where ���� IN ('5','6','7')"
            Set rs���ʷ��� = zlDataBase.OpenSQLRecord(gstrSQL, "Form_Load")
            
            With rs���ʷ���
                Do While Not .EOF
                    Tvw��;����.Nodes.Add "R", tvwChild, "R_" & !����, !����, 2, 2
                    .MoveNext
                Loop
                .Close
            End With
            
            .MoveFirst
            Do While Not .EOF
                If IsNull(!�ϼ�ID) Then
                    If !ĩ�� = 1 Then
                        Tvw��;����.Nodes.Add "R_" & !����, tvwChild, "K_" & !id, !����, 3, 3
                    Else
                        Tvw��;����.Nodes.Add "R_" & !����, tvwChild, "K_" & !id, !����, 2, 2
                    End If
                Else
                    If !ĩ�� = 1 Then
                        Tvw��;����.Nodes.Add "K_" & !�ϼ�ID, tvwChild, "K_" & !id, !����, 3, 3
                    Else
                        Tvw��;����.Nodes.Add "K_" & !�ϼ�ID, tvwChild, "K_" & !id, !����, 2, 2
                    End If
                End If
                Tvw��;����.Nodes("K_" & !id).Tag = !ĩ��
                .MoveNext
            Loop
        End With
    
        Tvw��;����.Nodes("R").Selected = True
        Tvw��;����.Nodes("R").Expanded = True
    End With
    
    mblnBootUp = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     Select Case UnloadMode
        Case vbFormControlMenu, vbAppWindows, vbAppTaskManager, vbFormOwner
            Me.Hide
        Case vbFormCode
            If Tvw��;����.Visible Then
                Tvw��;����.Visible = False
                Cmd��;.SetFocus
                Cancel = 1
                Exit Sub
            End If
    End Select
End Sub

Private Sub fraRangeSelect_Click()
    If Tvw��;����.Visible = True Then
        Tvw��;����.Visible = False
    End If
End Sub

Private Sub lst����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
    
End Sub

Private Sub Tvw��;����_DblClick()
    Me.Txt��;.Text = Tvw��;����.SelectedItem.Text
    Tvw��;����.Visible = False
    lvw����.SetFocus
End Sub

Private Sub Tvw��;����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Tvw��;����_DblClick
    End If
End Sub

Private Sub Tvw��;����_LostFocus()
    Tvw��;����.Visible = False
End Sub

Private Sub txtRate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyAdd
            If Val(txtRate.Text) < 100 Then
                txtRate.Text = Val(txtRate.Text) + 1
            End If
        Case vbKeySubtract
            If Val(txtRate.Text) > 1 Then
                txtRate.Text = Val(txtRate.Text) - 1
            End If
    End Select
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
        
        Case 48 To 57
            If IsNumeric(txtRate.Text) Then
                If txtRate.SelLength <> Len(txtRate.Text) Then
                    If Val(txtRate.Text & Chr(KeyAscii)) > 100 Then
                        KeyAscii = 0
                    End If
                End If
            End If
        Case 8          '�˸��
            
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtRate_Validate(Cancel As Boolean)
    If Trim(txtRate.Text) = "" Or Trim(txtRate.Text) = "0" Then
        Cancel = True
    End If
End Sub

Private Sub Chk����_Click()
    If Chk����.Value = 2 Then Exit Sub
    Call SetSelect(lvw����, Chk����.Value)
End Sub

Private Sub Lvw����_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call ItemCheck(lvw����, Item)
End Sub

Private Sub SetSelect(ByVal lvwObj As Object, Optional ByVal BlnSelect As Boolean = True)
    Dim intSelect As Integer
    With lvwObj
        For intSelect = 1 To .ListItems.count
            .ListItems(intSelect).Checked = BlnSelect
        Next
    End With
End Sub

Private Sub ItemCheck(ByVal lvwObj As Object, ByVal Item As MSComctlLib.ListItem)
    Dim lngCheck As Long, blnCheck As Boolean, intCount As Integer
    
    intCount = 0
    With lvwObj
        For lngCheck = 1 To .ListItems.count
            If .ListItems(lngCheck).Checked = True Then
                intCount = intCount + 1
            End If
        Next
        
        If intCount = lvwObj.ListItems.count Then
            Chk����.Value = 1
        ElseIf intCount > 0 Then
            Chk����.Value = 2
        Else
            Chk����.Value = 0
        End If
    End With
End Sub
