VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm���ղ��ֱ༭ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ղ��ֱ༭"
   ClientHeight    =   5280
   ClientLeft      =   2760
   ClientTop       =   3645
   ClientWidth     =   7950
   Icon            =   "frm���ղ��ֱ༭.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CmdClear 
      Caption         =   "ȫ��(&A)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   6660
      TabIndex        =   20
      Top             =   4110
      Width           =   1100
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "���(&D)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   5460
      TabIndex        =   19
      Top             =   4110
      Width           =   1100
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   360
      Left            =   3930
      TabIndex        =   13
      Top             =   240
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   635
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "��׼����(&0)"
      TabPicture(0)   =   "frm���ղ��ֱ༭.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "��׼��ϸ(&1)"
      TabPicture(1)   =   "frm���ղ��ֱ༭.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "ѡ��(&S)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4110
      TabIndex        =   18
      Top             =   4110
      Width           =   1100
   End
   Begin VB.Frame Fra��� 
      Caption         =   "���"
      Height          =   1545
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   3675
      Begin VB.OptionButton opt��� 
         Caption         =   "���Բ�(&M)"
         Height          =   180
         Index           =   1
         Left            =   300
         TabIndex        =   11
         Top             =   780
         Width           =   1155
      End
      Begin VB.OptionButton opt��� 
         Caption         =   "��ͨ��(&G)"
         Height          =   180
         Index           =   0
         Left            =   300
         TabIndex        =   10
         Top             =   420
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton opt��� 
         Caption         =   "���ֲ�(&T)"
         Height          =   180
         Index           =   2
         Left            =   300
         TabIndex        =   12
         Top             =   1140
         Width           =   1155
      End
   End
   Begin VB.Frame fra���� 
      Caption         =   "����"
      Height          =   2835
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   3675
      Begin VB.TextBox txt�ⶥ�� 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2070
         TabIndex        =   8
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CheckBox chk�ⶥ�� 
         Caption         =   "ʹ������ⶥ��(&T)"
         Height          =   315
         Left            =   180
         TabIndex        =   7
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   825
         MaxLength       =   6
         TabIndex        =   2
         Top             =   390
         Width           =   1095
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   825
         MaxLength       =   50
         TabIndex        =   4
         Top             =   780
         Width           =   2715
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   825
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1185
         Width           =   1095
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   3
         Top             =   840
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&E)"
         Height          =   180
         Index           =   2
         Left            =   150
         TabIndex        =   5
         Top             =   1245
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&U)"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   450
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   195
      TabIndex        =   23
      Top             =   4800
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5460
      TabIndex        =   21
      Top             =   4770
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6660
      TabIndex        =   22
      Top             =   4770
      Width           =   1100
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2940
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ղ��ֱ༭.frx":0044
            Key             =   "Fix"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ղ��ֱ༭.frx":035E
            Key             =   "FixD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ղ��ֱ༭.frx":0678
            Key             =   "Common"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ղ��ֱ༭.frx":0992
            Key             =   "CommonD"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ղ��ֱ༭.frx":0CAC
            Key             =   "Disease"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ղ��ֱ༭.frx":1246
            Key             =   "Limit"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra���� 
      Height          =   4245
      Left            =   3930
      TabIndex        =   14
      Top             =   420
      Width           =   3885
      Begin MSComctlLib.ListView lvw���� 
         Height          =   3300
         Left            =   90
         TabIndex        =   15
         Top             =   240
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   5821
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ils16"
         SmallIcons      =   "ils16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "����"
            Text            =   "����"
            Object.Width           =   3933
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "����"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin VB.Frame Fra��ϸ 
      Height          =   4245
      Left            =   3930
      TabIndex        =   16
      Top             =   420
      Width           =   3885
      Begin MSComctlLib.ListView Lvw��ϸ 
         Height          =   3300
         Left            =   90
         TabIndex        =   17
         Top             =   240
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   5821
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ils16"
         SmallIcons      =   "ils16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "����"
            Text            =   "����"
            Object.Width           =   3933
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "���"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "����"
            Object.Width           =   1764
         EndProperty
      End
   End
End
Attribute VB_Name = "frm���ղ��ֱ༭"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enum�༭
    text���� = 0
    Text���� = 1
    Text���� = 2
End Enum

Dim mlng���� As Long
Dim mstrID As String         '��ǰ�༭��ҽ������ID
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '�Ƿ�ı���

Private Sub chk�ⶥ��_Click()
    mblnChange = True
    If chk�ⶥ��.Value = 1 Then
        txt�ⶥ��.Enabled = True
        txt�ⶥ��.BackColor = txtEdit(1).BackColor
    Else
        txt�ⶥ��.Text = ""
        txt�ⶥ��.Enabled = False
        txt�ⶥ��.BackColor = fra����.BackColor
    End If
End Sub

Private Sub chk�ⶥ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0 'ʹ֮����
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub CmdClear_Click()
    Dim objLvw As ListView
    
    If SSTab.Tab = 0 Then
        Set objLvw = lvw����
    Else
        Set objLvw = Lvw��ϸ
    End If
    
    objLvw.ListItems.Clear
    
    CmdDel.Enabled = (objLvw.ListItems.Count <> 0)
    CmdClear.Enabled = CmdDel.Enabled
End Sub

Private Sub CmdDel_Click()
    Dim lngItem As Long
    Dim objLvw As ListView
    
    If SSTab.Tab = 0 Then
        Set objLvw = lvw����
    Else
        Set objLvw = Lvw��ϸ
    End If
    
    With objLvw
        If .ListItems.Count = 0 Then Exit Sub
        If .SelectedItem Is Nothing Then Exit Sub
        
        For lngItem = 1 To .ListItems.Count
            If lngItem > .ListItems.Count Then Exit For
            If .ListItems(lngItem).Selected Then
                .ListItems.Remove .ListItems(lngItem).Key
                lngItem = lngItem - 1
            End If
        Next
        
        If .ListItems.Count <> 0 Then .ListItems(1).Selected = True
    End With
    
    CmdDel.Enabled = (objLvw.ListItems.Count <> 0)
    CmdClear.Enabled = CmdDel.Enabled
End Sub

Private Sub Form_Load()
    If mlng���� = TYPE_���������� Then
        Load opt���(3)
        opt���(3).Top = opt���(0).Top
        opt���(3).Left = opt���(0).Left + opt���(0).Width + 150
        opt���(3).Visible = True
        
        Load opt���(4)
        opt���(4).Top = opt���(1).Top
        opt���(4).Left = opt���(1).Left + opt���(1).Width + 150
        opt���(4).Visible = True
        
        '�޸�����
        opt���(1).Caption = "��ͨ��"
        opt���(1).Caption = "���ⲡ"
        opt���(2).Caption = "���ﲡ"
        opt���(3).Caption = "��������"
        opt���(4).Caption = "����"
        opt���(0).Value = True
    End If
End Sub

Private Sub lvw����_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvw����, ColumnHeader.Index)
End Sub

Private Sub SSTab_Click(PreviousTab As Integer)
    Dim objLvw As ListView
    Select Case SSTab.Tab
    Case 0
        Fra��ϸ.Visible = True
        Set objLvw = lvw����
        fra����.ZOrder
    Case 1
        fra����.Visible = True
        Set objLvw = Lvw��ϸ
        Fra��ϸ.ZOrder
    End Select
    
    SSTab.ZOrder
    cmdADD.ZOrder
    CmdClear.ZOrder
    CmdDel.ZOrder
    
    CmdDel.Enabled = (objLvw.ListItems.Count <> 0)
    CmdClear.Enabled = CmdDel.Enabled
End Sub

Private Sub cmdADD_Click()
    With frm��׼��Ŀѡ��
        .lng���� = mlng����
        .bln��ϸ = (SSTab.Tab = 1)
        Set .frmParent = Me
        .Show 1, Me
    End With
    Call SSTab_Click(SSTab.Tab)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngIndex As Long
    
    If IsValid() = False Then Exit Sub
    If Save��Ŀ() = False Then Exit Sub
    
    If mstrID = "" Then
        '��������
        'Modified by ���� 20031218 ����������
        If mlng���� = TYPE_�������� Or mlng���� = TYPE_����ʡ Or mlng���� = TYPE_������ Or mlng���� = TYPE_��ƽ�� Then
            txtEdit(text����).Text = GetMaxCode
        Else
            txtEdit(text����).Text = zlDatabase.GetMax("���ղ���", "����", 6, " where ����=" & mlng����)
        End If
        For lngIndex = Text���� To Text����
            txtEdit(lngIndex).Text = ""
        Next
        lvw����.ListItems.Clear
        Lvw��ϸ.ListItems.Clear
        
        mblnChange = False
        txtEdit(text����).SetFocus
    Else
        mblnChange = False
        Unload Me
    End If
End Sub

Private Function Save��Ŀ() As Boolean
    Dim lngID As Long, lng��� As Long, lng����ID As Long
    Dim lngIndex As Long, lst As ListItem
    Dim strCode As String
    Dim rsTmp As New ADODB.Recordset
    
    For lngIndex = opt���.LBound To opt���.UBound
        If opt���(lngIndex).Value = True Then
            lng��� = lngIndex
            Exit For
        End If
    Next
    
    On Error GoTo errHandle
    gcnOracle.BeginTrans
    If mstrID = "" Then
        '����
        lngID = zlDatabase.GetNextId("���ղ���")
        'Modified by ���� 20031218 ����������
        If mlng���� = TYPE_�������� Or mlng���� = TYPE_����ʡ Or mlng���� = TYPE_������ Or mlng���� = TYPE_��ƽ�� Then
            If CheckCode(txtEdit(text����)) = False Then Exit Function
            '��ȡ���ձ���
            strCode = zlDatabase.GetMax("���ղ���", "����", 6, " Where ����=" & mlng����)
            gstrSQL = "zl_���ղ���_INSERT(" & lngID & "," & mlng���� & ",'" & strCode & "','" & _
                    Trim(txtEdit(text����).Text) & "@@" & Trim(txtEdit(Text����).Text) & "','" & Trim(txtEdit(Text����).Text) & "'," & lng��� & ",null,null)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Else
            lng����ID = lngID
            gstrSQL = "zl_���ղ���_INSERT(" & lngID & "," & mlng���� & ",'" & Trim(txtEdit(text����).Text) & "','" & _
                    Trim(txtEdit(Text����).Text) & "','" & Trim(txtEdit(Text����).Text) & "'," & lng��� & _
                    "," & chk�ⶥ��.Value & "," & IIf(chk�ⶥ��.Value = 0, "null", IIf(txt�ⶥ��.Text = "", "null", txt�ⶥ��.Text)) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
    Else
        'Modified by ���� 20031218 ����������
        If mlng���� = TYPE_�������� Or mlng���� = TYPE_����ʡ Or mlng���� = TYPE_������ Or mlng���� = TYPE_��ƽ�� Then
            If CheckCode(txtEdit(text����), False) = False Then Exit Function
            '��ȡ���ձ���
            gstrSQL = "Select ���� From ���ղ��� Where ����=" & mlng���� & " And ID=" & mstrID
            Call OpenRecordset(rsTmp, "��ȡ��ǰ���ղ��ֵı���")
            strCode = rsTmp!����
            
            gstrSQL = "zl_���ղ���_Update(" & mstrID & ",'" & strCode & "','" & _
                    Trim(txtEdit(text����).Text) & "@@" & Trim(txtEdit(Text����).Text) & "','" & Trim(txtEdit(Text����).Text) & "'," & lng��� & ",null,null)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Else
            lng����ID = mstrID
            gstrSQL = "zl_���ղ���_Update(" & mstrID & ",'" & Trim(txtEdit(text����).Text) & "','" & _
                    Trim(txtEdit(Text����).Text) & "','" & Trim(txtEdit(Text����).Text) & "'," & lng��� & _
                    "," & chk�ⶥ��.Value & "," & IIf(chk�ⶥ��.Value = 0, "null", IIf(txt�ⶥ��.Text = "", "null", txt�ⶥ��.Text)) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
        gstrSQL = "zl_������׼��Ŀ_INSERT(" & mstrID & ",NULL,0,0,1)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    
    '������׼��Ŀ�����ࣩ
    For Each lst In lvw����.ListItems
        gstrSQL = "zl_������׼��Ŀ_INSERT(" & lng����ID & "," & Mid(lst.Key, 2) & ",1," & Mid(lst.SubItems(1), 1, 1) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    
    '������׼��Ŀ����ϸ��
    For Each lst In Lvw��ϸ.ListItems
        gstrSQL = "zl_������׼��Ŀ_INSERT(" & lng����ID & "," & Mid(lst.Key, 2) & ",0," & Mid(lst.SubItems(2), 1, 1) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    gcnOracle.CommitTrans
    
    '����������
    If mstrID = "" Then
        Set lst = frm���ղ���.lvwItem.ListItems.Add(, "K" & lngID, txtEdit(text����), "Disease", "Disease")
    Else
        Set lst = frm���ղ���.lvwItem.SelectedItem
    End If
    lst.SubItems(1) = Trim(txtEdit(Text����).Text)
    lst.SubItems(2) = Trim(txtEdit(Text����).Text)
    '��������ҽ�������� 204-03-31
    If mlng���� = TYPE_���������� Then
        lst.SubItems(3) = IIf(lng��� = 1, "���ⲡ", IIf(lng��� = 2, "���ﲡ", IIf(lng��� = 3, "��������", IIf(lng��� = 4, "����", "��ͨ��"))))
    Else
        lst.SubItems(3) = IIf(lng��� = 0, "��ͨ��", IIf(lng��� = 1, "���Բ�", "���ֲ�"))
    End If
    lst.SubItems(4) = IIf(chk�ⶥ��.Value = 1, "��", "")
    lst.SubItems(5) = IIf(chk�ⶥ��.Value = 0, "", IIf(txt�ⶥ��.Text = "", "�޷ⶥ��", txt�ⶥ��.Text))
    
    Save��Ŀ = True
    mblnOK = True
    Exit Function

errHandle:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Function

Private Function IsValid() As Boolean
'����:���������й�ҽ�����������Ƿ���Ч
'����:
'����ֵ:��Ч����True,����ΪFalse
    Dim lngIndex As Integer
    For lngIndex = text���� To Text����
        If zlCommFun.StrIsValid(Trim(txtEdit(lngIndex).Text), txtEdit(lngIndex).MaxLength) = False Then
            txtEdit(lngIndex).SetFocus
            zlControl.TxtSelAll txtEdit(lngIndex)
            Exit Function
        End If
        
        If lngIndex = text���� Or lngIndex = Text���� Then
            If Len(Trim(txtEdit(lngIndex).Text)) = 0 Then
                txtEdit(lngIndex).Text = ""
                MsgBox "��������ƶ�����Ϊ�ա�", vbExclamation, gstrSysName
                txtEdit(lngIndex).SetFocus
                Exit Function
            End If
        End If
    Next
    If txt�ⶥ��.Enabled = True Then
        If txt�ⶥ��.Text <> "" Then
            If zlCommFun.IntIsValid(txt�ⶥ��.Text, 8, True, True, txt�ⶥ��.hwnd, "�ⶥ�߽��") = False Then
                Exit Function
            End If
        End If
    End If
    
    IsValid = True
End Function

Private Sub opt���_Click(Index As Integer)
    mblnChange = True
End Sub

Private Sub opt���_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = Text���� Then
        txtEdit(Text����).Text = zlCommFun.SpellCode(txtEdit(Text����).Text)
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    Select Case Index
        Case Text����
            zlCommFun.OpenIme True
        Case Else
            zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0 'ʹ֮����
        zlCommFun.PressKey (vbKeyTab)
    Else
        If Index = text���� Then
            zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m�ı�ʽ
'            KeyAscii = Asc(UCase(Chr(KeyAscii)))
'            If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
End Sub

Public Function �༭����(ByVal lng���� As Long, ByVal strID As String) As Boolean
'����:��������õ�ҽ���������ڽ���ͨѶ�ĳ���
'����:str���           ��ǰ�༭��ҽ�����ĵ����
'����ֵ:�༭�ɹ�����True,����ΪFalse
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer, lst As ListItem
    
    mblnOK = False
    mlng���� = lng����
    mstrID = strID
    
    rsTemp.CursorLocation = adUseClient
    If mstrID <> "" Then
        '�޸�ҽ������
        'Modified by ���� 20031218 ����������
        If mlng���� = TYPE_�������� Or mlng���� = TYPE_����ʡ Or mlng���� = TYPE_������ Or mlng���� = TYPE_��ƽ�� Then
            'Modified by ���� 20031218 ����������
            txtEdit(text����).MaxLength = 20
            gstrSQL = "select substr(����,1,instr(����,'@@')-1) ����,substr(����,instr(����,'@@')+2) ����,����,nvl(���,'0') as ���,����ⶥ��,�ⶥ�߽�� from ���ղ��� where ID=" & mstrID
        Else
            gstrSQL = "select ����,����,����,nvl(���,'0') as ���,����ⶥ��,�ⶥ�߽�� from ���ղ��� where ID=" & mstrID
        End If
        Call OpenRecordset(rsTemp, Me.Caption)
        
        txtEdit(text����).Text = rsTemp("����")
        txtEdit(Text����).Text = rsTemp("����")
        txtEdit(Text����).Text = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
        opt���(rsTemp("���")).Value = True
        
        If rsTemp("����ⶥ��") = 1 Then
            chk�ⶥ��.Value = 1
            If Not IsNull(rsTemp("�ⶥ�߽��")) Then
                txt�ⶥ��.Text = rsTemp("�ⶥ�߽��")
            End If
        End If
        
        '�޸���׼����
        gstrSQL = "select A.ID,A.����,A.����,decode(B.����,1,'1-����',2,'2-�ų�','0-����') ���� from ����֧������ A,������׼��Ŀ B where A.ID=B.�շ�ϸĿID and B.����=1 and B.����ID=" & mstrID
        Call OpenRecordset(rsTemp, Me.Caption)
        
        Do Until rsTemp.EOF
            Set lst = lvw����.ListItems.Add(, "K" & rsTemp("ID"), "[" & rsTemp!���� & "]" & rsTemp("����"), "Limit", "Limit")
            lst.SubItems(1) = rsTemp("����")
            rsTemp.MoveNext
        Loop
    
        '�޸���׼��Ŀ
        gstrSQL = "select A.ID,A.����,A.����,A.���,decode(B.����,1,'1-����',2,'2-�ų�','0-����') ���� from �շ�ϸĿ A,������׼��Ŀ B where A.ID=B.�շ�ϸĿID and B.����=0 and B.����ID=" & mstrID
        Call OpenRecordset(rsTemp, Me.Caption)
        
        Do Until rsTemp.EOF
            Set lst = Lvw��ϸ.ListItems.Add(, "K" & rsTemp("ID"), "[" & rsTemp!���� & "]" & rsTemp("����"), "Fix", "Fix")
            lst.SubItems(1) = Nvl(rsTemp("���"))
            lst.SubItems(2) = Nvl(rsTemp("����"))
            rsTemp.MoveNext
        Loop
    
    Else
        '����ҽ������
        txtEdit(text����).Text = zlDatabase.GetMax("���ղ���", "����", 6, " where ����=" & mlng����)
    End If
    
    '��������ҽ�������� 204-03-31
    'Modified by ���� 20031218 ����������
    If mlng���� = TYPE_�������� Or mlng���� = TYPE_����ʡ Or mlng���� = TYPE_������ Or mlng���� = TYPE_��ƽ�� Or mlng���� = TYPE_���������� Then
        'Modified by ���� 20031218 ����������
        txtEdit(text����).MaxLength = 20
        cmdADD.Enabled = False
        CmdClear.Enabled = False
        CmdDel.Enabled = False
    End If
    
    mblnChange = False
    frm���ղ��ֱ༭.Show vbModal, frm���ղ���
    �༭���� = mblnOK
End Function

Private Sub txt�ⶥ��_Change()
    mblnChange = True
End Sub

Private Sub txt�ⶥ��_GotFocus()
    zlControl.TxtSelAll txt�ⶥ��
End Sub

Private Sub txt�ⶥ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0 'ʹ֮����
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Function GetMaxCode() As String
'���ܣ���ȡָ����ı�����������ֵ
'���أ��ɹ����� �¼�������; ���߷��� 0
    Dim rsTemp As New ADODB.Recordset
    Dim varTemp As Variant
    Dim lngLengh As Long
    
    On Error GoTo ErrHand
    With rsTemp
        gstrSQL = "SELECT max(length(substr(����,1,instr(����,'@@')-1))) as �ֵ FROM ���ղ��� where ����=" & mlng����
        Call OpenRecordset(rsTemp, Me.Caption)
        If rsTemp.EOF = True Then
            GetMaxCode = "1"
            Exit Function
        Else
            lngLengh = Nvl(rsTemp("�ֵ"), "1")
        End If
        
        gstrSQL = "SELECT MAX(LPAD(substr(����,1,instr(����,'@@')-1)," & lngLengh & ",' ')) as ���ֵ FROM ���ղ��� where ����=" & mlng����
        Call OpenRecordset(rsTemp, Me.Caption)
        If rsTemp.EOF Then
            GetMaxCode = Format(1, String(lngLengh, "0"))
            Exit Function
        End If
        
        varTemp = Nvl(rsTemp("���ֵ"), "0")
        If IsNumeric(varTemp) Then
            GetMaxCode = CStr(Val(varTemp) + 1)
            GetMaxCode = Format(GetMaxCode, String(lngLengh, "0"))
        Else
            GetMaxCode = Mid(varTemp, 1, Len(varTemp) - 1) & Chr(asc(Right(varTemp, 1)) + 1)
            GetMaxCode = Trim(GetMaxCode)
        End If
        .Close
    End With
    Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
End Function

Private Function CheckCode(ByVal strCode As String, Optional ByVal blnNew As Boolean = True) As Boolean
    Dim rsCode As New ADODB.Recordset
    '��Ϊ���볬����ֻ�н����������Ʊ����������У���������ʵ�ʱ�����Ǽ�¼�������û��޸ı���ʱ����Ҫ�жϱ����Ƿ��ظ�
    
    CheckCode = False
    gstrSQL = "Select 1 From ���ղ��� Where ����=" & mlng���� & " And substr(����,1,instr(����,'@@')-1)='" & strCode & "'" & IIf(blnNew, "", " And ID<>" & mstrID)
    Call OpenRecordset(rsCode, "�жϱ����Ƿ��ظ�")
    
    If Not rsCode.EOF Then
        MsgBox "���ղ��ֱ����ظ���", vbInformation, gstrSysName
        txtEdit(text����).SetFocus
        Exit Function
    End If
    CheckCode = True
End Function
