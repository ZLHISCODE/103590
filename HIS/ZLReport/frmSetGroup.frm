VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSetGroup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ñ�����"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   Icon            =   "frmSetGroup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdUp 
      Caption         =   "��"
      Height          =   435
      Left            =   6450
      TabIndex        =   9
      ToolTipText     =   "������"
      Top             =   1230
      Width           =   345
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "��"
      Height          =   435
      Left            =   6450
      TabIndex        =   8
      ToolTipText     =   "������"
      Top             =   1800
      Width           =   345
   End
   Begin MSComctlLib.ImageList Img 
      Left            =   2880
      Top             =   210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetGroup.frx":020A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetGroup.frx":0524
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdSend 
      Caption         =   "ȫ��"
      Height          =   350
      Index           =   3
      Left            =   2730
      TabIndex        =   4
      Top             =   2460
      Width           =   885
   End
   Begin VB.CommandButton CmdSend 
      Caption         =   "���"
      Height          =   350
      Index           =   2
      Left            =   2730
      TabIndex        =   3
      Top             =   1980
      Width           =   885
   End
   Begin VB.CommandButton CmdSend 
      Caption         =   "ѡ��"
      Height          =   350
      Index           =   1
      Left            =   2730
      TabIndex        =   2
      Top             =   1320
      Width           =   885
   End
   Begin VB.CommandButton CmdSend 
      Caption         =   "ȫѡ"
      Height          =   350
      Index           =   0
      Left            =   2730
      TabIndex        =   1
      Top             =   840
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6990
      TabIndex        =   6
      Top             =   90
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6990
      TabIndex        =   7
      Top             =   510
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwFrom 
      Height          =   3390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   5980
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "Img"
      SmallIcons      =   "Img"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "������������"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "˵��"
         Object.Width           =   4410
      EndProperty
   End
   Begin MSComctlLib.ListView lvwSet 
      Height          =   3390
      Left            =   3750
      TabIndex        =   5
      Top             =   0
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   5980
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "Img"
      SmallIcons      =   "Img"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "���"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "������������"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "˵��"
         Object.Width           =   4410
      EndProperty
   End
End
Attribute VB_Name = "frmSetGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private BlnSave As Boolean
Private ItemThis As ListItem
Private IntItems As Integer
Private mrsLoad As New ADODB.Recordset
Public LngGroupID As Long
Public strCaption As String

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdDown_Click()
    Dim ItemThis As ListItem, ItemUpper As ListItem, TmpKey As String
    Dim ItemText As String, ItemTag As String, ItemKey As String
    Dim ItemSub1 As String, ItemSub2 As String, SubTag1 As String
    
    Set ItemUpper = lvwSet.ListItems(lvwSet.SelectedItem.Index + 1)
    Set ItemThis = lvwSet.SelectedItem
    ItemText = ItemUpper.Text
    ItemSub1 = ItemUpper.SubItems(1)
    ItemSub2 = ItemUpper.SubItems(2)
    SubTag1 = ItemUpper.ListSubItems(1).Tag
    ItemKey = ItemUpper.Key
    TmpKey = ItemThis.Key
    ItemUpper.Key = "_8888"
    ItemThis.Key = "_9999"
    
    ItemUpper.Text = ItemThis.Text
    ItemUpper.SubItems(1) = ItemThis.SubItems(1)
    ItemUpper.SubItems(2) = ItemThis.SubItems(2)
    ItemUpper.ListSubItems(1).Tag = ItemThis.ListSubItems(1).Tag
    ItemUpper.Key = TmpKey
    
    ItemThis.Text = ItemText
    ItemThis.SubItems(1) = ItemSub1
    ItemThis.SubItems(2) = ItemSub2
    ItemThis.ListSubItems(1).Tag = SubTag1
    ItemThis.Key = ItemKey
    
    Set lvwSet.SelectedItem = lvwSet.ListItems(lvwSet.SelectedItem.Index + 1)
    Call WriteOrder
    
    cmdUp.Enabled = False
    cmdDown.Enabled = False
    If Not lvwSet.SelectedItem Is Nothing Then
        cmdDown.Enabled = (lvwSet.SelectedItem.Index < lvwSet.ListItems.count)
        cmdUp.Enabled = (lvwSet.SelectedItem.Index > 1)
    End If
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer, j As Integer
    
    On Error GoTo ErrHand
    
    For i = 1 To lvwSet.ListItems.count - 1
        For j = i + 1 To lvwSet.ListItems.count
            If lvwSet.ListItems(j).ListSubItems(1).Tag = lvwSet.ListItems(i).ListSubItems(1).Tag Then
                MsgBox "�������д������ƶ�Ϊ""" & lvwSet.ListItems(i).ListSubItems(1).Tag & """�Ķ��ű���ͬһ�������е��ӱ����Ʋ�����ͬ��", vbInformation, App.Title
                Exit Sub
            End If
        Next
    Next
    
    
    '��������
    gcnOracle.BeginTrans
    gcnOracle.Execute "Delete zlRPTSubs Where ��ID=" & LngGroupID
    For IntItems = 1 To lvwSet.ListItems.count
        gcnOracle.Execute "Insert Into zlRPTSubs(��ID,����ID,���,����) Values(" & _
            LngGroupID & "," & Mid(lvwSet.ListItems(IntItems).Key, 2) & "," & lvwSet.ListItems(IntItems).Text & ",'" & lvwSet.ListItems(IntItems).ListSubItems(1).Tag & "')"
    Next
    gcnOracle.CommitTrans
    
    BlnSave = True
    Unload Me
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Sub CmdSend_Click(Index As Integer)
    Select Case Index
    Case 0  'ȫѡ
        For IntItems = 1 To lvwFrom.ListItems.count
            Set ItemThis = lvwFrom.ListItems(IntItems)
            Call lvwSet.ListItems.Add(, ItemThis.Key, 1, ItemThis.Icon, ItemThis.SmallIcon)
            lvwSet.ListItems(ItemThis.Key).SubItems(1) = ItemThis.Text
            lvwSet.ListItems(ItemThis.Key).SubItems(2) = ItemThis.SubItems(1)
            lvwSet.ListItems(ItemThis.Key).ListSubItems(1).Tag = ItemThis.ListSubItems(1).Tag
        Next
        lvwFrom.ListItems.Clear
    Case 1  '��ѡ
        Set ItemThis = lvwFrom.SelectedItem
        Call lvwSet.ListItems.Add(, ItemThis.Key, 1, ItemThis.Icon, ItemThis.SmallIcon)
        lvwSet.ListItems(ItemThis.Key).SubItems(1) = ItemThis.Text
        lvwSet.ListItems(ItemThis.Key).SubItems(2) = ItemThis.SubItems(1)
        lvwSet.ListItems(ItemThis.Key).ListSubItems(1).Tag = ItemThis.ListSubItems(1).Tag
        lvwFrom.ListItems.Remove ItemThis.Key
    Case 2  '���
        Set ItemThis = lvwSet.SelectedItem
        Call lvwFrom.ListItems.Add(, ItemThis.Key, ItemThis.SubItems(1), ItemThis.Icon, ItemThis.SmallIcon)
        lvwFrom.ListItems(ItemThis.Key).SubItems(1) = ItemThis.SubItems(2)
        lvwFrom.ListItems(ItemThis.Key).ListSubItems(1).Tag = ItemThis.ListSubItems(1).Tag
        lvwSet.ListItems.Remove ItemThis.Key
    Case 3  'ȫ��
        For IntItems = 1 To lvwSet.ListItems.count
            Set ItemThis = lvwSet.ListItems(IntItems)
            Call lvwFrom.ListItems.Add(, ItemThis.Key, ItemThis.SubItems(1), ItemThis.Icon, ItemThis.SmallIcon)
            lvwFrom.ListItems(ItemThis.Key).SubItems(1) = ItemThis.SubItems(2)
            lvwFrom.ListItems(ItemThis.Key).ListSubItems(1).Tag = ItemThis.ListSubItems(1).Tag
        Next
        lvwSet.ListItems.Clear
    End Select
    BlnSave = False
    Call SetCmdState
    Call WriteOrder
End Sub

Private Sub cmdUp_Click()
    Dim ItemThis As ListItem, ItemUpper As ListItem, TmpKey As String
    Dim ItemText As String, ItemTag As String, ItemKey As String
    Dim ItemSub1 As String, ItemSub2 As String, SubTag1 As String
    
    Set ItemUpper = lvwSet.ListItems(lvwSet.SelectedItem.Index - 1)
    Set ItemThis = lvwSet.SelectedItem
    ItemText = ItemUpper.Text
    ItemSub1 = ItemUpper.SubItems(1)
    ItemSub2 = ItemUpper.SubItems(2)
    SubTag1 = ItemUpper.ListSubItems(1).Tag
    ItemKey = ItemUpper.Key
    TmpKey = ItemThis.Key
    ItemUpper.Key = "_8888"
    ItemThis.Key = "_9999"
    
    ItemUpper.Text = ItemThis.Text
    ItemUpper.SubItems(1) = ItemThis.SubItems(1)
    ItemUpper.SubItems(2) = ItemThis.SubItems(2)
    ItemUpper.ListSubItems(1).Tag = ItemThis.ListSubItems(1).Tag
    ItemUpper.Key = TmpKey
    
    ItemThis.Text = ItemText
    ItemThis.SubItems(1) = ItemSub1
    ItemThis.SubItems(2) = ItemSub2
    ItemThis.ListSubItems(1).Tag = SubTag1
    ItemThis.Key = ItemKey
    
    Set lvwSet.SelectedItem = lvwSet.ListItems(lvwSet.SelectedItem.Index - 1)
    Call WriteOrder
    
    cmdUp.Enabled = False
    cmdDown.Enabled = False
    If Not lvwSet.SelectedItem Is Nothing Then
        cmdDown.Enabled = (lvwSet.SelectedItem.Index < lvwSet.ListItems.count)
        cmdUp.Enabled = (lvwSet.SelectedItem.Index > 1)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub Form_Load()
    BlnSave = True
    Caption = strCaption
    Call LoadAnother
    Call LoadHold
    Call SetCmdState
End Sub

Private Function SetCmdState()
    '���ø���ť��״̬
    CmdSend(0).Enabled = (lvwFrom.ListItems.count <> 0)
    CmdSend(1).Enabled = (Not lvwFrom.SelectedItem Is Nothing)
    CmdSend(2).Enabled = (Not lvwSet.SelectedItem Is Nothing)
    CmdSend(3).Enabled = (lvwSet.ListItems.count <> 0)
    cmdUp.Enabled = False
    cmdDown.Enabled = False
    If Not lvwSet.SelectedItem Is Nothing Then
        cmdDown.Enabled = (lvwSet.SelectedItem.Index < lvwSet.ListItems.count)
        cmdUp.Enabled = (lvwSet.SelectedItem.Index > 1)
    End If
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If BlnSave = False Then
        If MsgBox("��ȷ��Ҫ�˳��𣿣���δ���棩", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Cancel = 1: Exit Sub
    End If
End Sub

Private Sub lvwFrom_DblClick()
    If CmdSend(1).Enabled Then CmdSend_Click (1)
End Sub

Private Sub lvwFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then lvwFrom_DblClick
End Sub

Private Sub lvwSet_DblClick()
    If CmdSend(2).Enabled Then CmdSend_Click (2)
End Sub

Private Sub lvwSet_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmdUp.Enabled = False
    cmdDown.Enabled = False
    If Not lvwSet.SelectedItem Is Nothing Then
        cmdDown.Enabled = (lvwSet.SelectedItem.Index < lvwSet.ListItems.count)
        cmdUp.Enabled = (lvwSet.SelectedItem.Index > 1)
    End If
End Sub

Private Sub lvwSet_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then lvwSet_DblClick
End Sub

Private Function LoadAnother()
    Dim strSQL As String, intIcon As Integer
    
    'װ�뱾ϵͳ�����б���(֧���������б���)
    With frmMain.cboSys
        strSQL = "Select ID,���,����,˵��,����ID From zlReports" & _
              " Where " & IIF(.ItemData(.ListIndex) = 0, " ϵͳ Is Null", " ϵͳ=[1]") & _
              " And ID Not In (Select ����ID From zlRPTSubs Where ��ID=[2])"
        Set mrsLoad = OpenSQLRecord(strSQL, Me.Caption, .ItemData(.ListIndex), LngGroupID)
    End With
    lvwFrom.ListItems.Clear
    Do While Not mrsLoad.EOF
        intIcon = IIF(mrsLoad!����ID = 0, 1, 2)
        lvwFrom.ListItems.Add , "_" & mrsLoad!id, "[" & mrsLoad!��� & "]" & mrsLoad!����, intIcon, intIcon
        lvwFrom.ListItems("_" & mrsLoad!id).SubItems(1) = Nvl(mrsLoad!˵��)
        lvwFrom.ListItems("_" & mrsLoad!id).ListSubItems(1).Tag = mrsLoad!����
        mrsLoad.MoveNext
    Loop
End Function

Private Function LoadHold()
    'װ�뱾����������б���
    Dim strSQL As String, intIcon As Integer
    
    strSQL = "Select A.ID,A.���,A.����,A.˵��,A.����ID,B.��� From zlReports A,zlRPTSubs B Where B.����ID=A.ID And B.��ID=[1] Order by B.���"
    Set mrsLoad = OpenSQLRecord(strSQL, Me.Caption, LngGroupID)
    lvwSet.ListItems.Clear
    Do While Not mrsLoad.EOF
        intIcon = IIF(mrsLoad!����ID = 0, 1, 2)
        lvwSet.ListItems.Add , "_" & mrsLoad!id, mrsLoad!���, intIcon, intIcon
        lvwSet.ListItems("_" & mrsLoad!id).SubItems(1) = "[" & mrsLoad!��� & "]" & mrsLoad!����
        lvwSet.ListItems("_" & mrsLoad!id).SubItems(2) = Nvl(mrsLoad!˵��)
        lvwSet.ListItems("_" & mrsLoad!id).ListSubItems(1).Tag = mrsLoad!����
        mrsLoad.MoveNext
    Loop
End Function

Private Function WriteOrder()
    Dim intLoop As Integer, ItemChange As ListItem
    '�������
    
    For intLoop = 1 To lvwSet.ListItems.count
        Set ItemChange = lvwSet.ListItems(intLoop)
        ItemChange.Text = intLoop
    Next
End Function
