VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmCaseNarSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������Ŀ����"
   ClientHeight    =   4440
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6900
   Icon            =   "frmCaseNarSet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ɾ��(&D)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   3330
      TabIndex        =   18
      Top             =   3930
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "����Ŀ(&N)"
      Height          =   350
      Left            =   2205
      TabIndex        =   17
      Top             =   3930
      Visible         =   0   'False
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3525
      Left            =   30
      TabIndex        =   1
      Top             =   225
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   6218
      SortKey         =   3
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "��Ŀ����"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "��λ"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "�ַ�"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "��¼����"
         Object.Width           =   1852
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   6585
      Top             =   5070
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraData 
      Caption         =   "��Ŀ����"
      Height          =   2265
      Left            =   4005
      TabIndex        =   23
      Top             =   135
      Width           =   2835
      Begin MSComCtl2.UpDown udMin 
         Height          =   300
         Left            =   2326
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1380
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   10
         BuddyControl    =   "txtMin"
         BuddyDispid     =   196615
         OrigLeft        =   3330
         OrigTop         =   1350
         OrigRight       =   3570
         OrigBottom      =   1560
         Increment       =   20
         Max             =   300
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMax 
         Height          =   300
         Left            =   2326
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1800
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   10
         BuddyControl    =   "txtMax"
         BuddyDispid     =   196614
         OrigLeft        =   5565
         OrigTop         =   480
         OrigRight       =   5805
         OrigBottom      =   690
         Increment       =   20
         Max             =   300
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.ComboBox cboType 
         Enabled         =   0   'False
         Height          =   300
         ItemData        =   "frmCaseNarSet.frx":000C
         Left            =   1230
         List            =   "frmCaseNarSet.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1005
         Width           =   1095
      End
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   1230
         MaxLength       =   10
         TabIndex        =   3
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtMax 
         Height          =   300
         Left            =   1230
         MaxLength       =   12
         TabIndex        =   11
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtMin 
         Height          =   300
         Left            =   1230
         MaxLength       =   12
         TabIndex        =   9
         Top             =   1380
         Width           =   1095
      End
      Begin VB.TextBox txtUnit 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1230
         MaxLength       =   6
         TabIndex        =   5
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "��¼����(&T)"
         Height          =   240
         Left            =   165
         TabIndex        =   6
         Top             =   1050
         Width           =   1035
      End
      Begin VB.Label Label7 
         Caption         =   "��Ŀ����(&M)"
         Height          =   240
         Left            =   165
         TabIndex        =   2
         Top             =   330
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "���ֵ(&A)"
         Height          =   195
         Left            =   165
         TabIndex        =   10
         Top             =   1860
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "��Сֵ(&I)"
         Height          =   195
         Left            =   165
         TabIndex        =   8
         Top             =   1455
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "��λ(&U)"
         Height          =   240
         Left            =   165
         TabIndex        =   4
         Top             =   690
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&X)"
      Height          =   350
      Left            =   5730
      TabIndex        =   19
      Top             =   3915
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "����(&S)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4500
      TabIndex        =   16
      Top             =   3930
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   45
      TabIndex        =   20
      Top             =   3915
      Width           =   1100
   End
   Begin VB.Frame fraDisplay 
      Caption         =   "��ʾЧ��"
      Height          =   1155
      Left            =   4005
      TabIndex        =   21
      Top             =   2595
      Width           =   2850
      Begin VB.ComboBox cboChar 
         Height          =   300
         Left            =   1245
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   300
         Width           =   1155
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "��"
         Height          =   270
         Left            =   2370
         TabIndex        =   15
         Top             =   720
         Width           =   270
      End
      Begin VB.Label lblColor 
         BackColor       =   &H80000007&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   1245
         TabIndex        =   22
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label3 
         Caption         =   "��¼ɫ(&L)"
         Height          =   210
         Left            =   390
         TabIndex        =   14
         Top             =   735
         Width           =   930
      End
      Begin VB.Label Label2 
         Caption         =   "��¼��(&R)"
         Height          =   225
         Left            =   390
         TabIndex        =   12
         Top             =   345
         Width           =   990
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1230
      Top             =   15
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
            Picture         =   "frmCaseNarSet.frx":0010
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseNarSet.frx":0468
            Key             =   "NewItem"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      Caption         =   "������Ŀ(&G)"
      Height          =   180
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   1005
   End
End
Attribute VB_Name = "frmCaseNarSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public blnOK As Boolean
Private strSQL As String
Private rsTmp As New ADODB.Recordset

Private svrItem As MSComctlLib.ListItem
'��־����2���б�־Ϊ��0��ǰ    1����
'      ��3���б�־Ϊ��0���κα仯  1����   2ɾ��

Private Sub ItemDel(objLvw As ListView)
'��־Ϊɾ��
Dim i As Long
Dim lngMaxIndex As Long
Dim objList As ListItem

With objLvw
    If .SelectedItem Is Nothing Then Exit Sub
    lngMaxIndex = .SelectedItem.Index
    '����ɾ����ɫ
    For i = 1 To .ColumnHeaders.Count
        If i = 1 Then
            .SelectedItem.ForeColor = RGB(255, 0, 0)
        Else
            .SelectedItem.ListSubItems(i - 1).ForeColor = RGB(255, 0, 0)
        End If
    Next
    If lvw.SelectedItem.ListSubItems(2).Tag = 0 Then
        Me.cmdDelete.Caption = "��ɾ��(&D)"
    Else
        Me.cmdDelete.Caption = "ɾ��(&D)"
    End If
    '����ɾ����־
    If .SelectedItem.ListSubItems(2).Tag = 0 Then
        .SelectedItem.ListSubItems(3).Tag = 2
    Else
        .ListItems.Remove lngMaxIndex
    End If
    '������һѡ����
    On Error Resume Next
    Err.Clear
    Set objList = objLvw.ListItems(lngMaxIndex)
    If Not (objList Is Nothing) Then
        objList.Selected = True
        objList.EnsureVisible
        lvw_ItemClick lvw.SelectedItem
    ElseIf Err <> 0 Then
        Set objList = objLvw.ListItems(lngMaxIndex - 1)
        If Err <> 0 Or Not (objList Is Nothing) Then
            objList.Selected = True
            objList.EnsureVisible
            lvw_ItemClick lvw.SelectedItem
        Else
            Err.Clear
        End If
    End If
End With
End Sub

Private Sub UNItemDel(objLvw As ListView)
'ȡ��ɾ����־
Dim i As Long
Dim lngMaxIndex As Long
Dim objList As ListItem

With objLvw
    If .SelectedItem Is Nothing Then Exit Sub
    lngMaxIndex = .SelectedItem.Index
    'ȡ��ɾ����־
    For i = 1 To .ColumnHeaders.Count
        If i = 1 Then
            .SelectedItem.ForeColor = 0
        Else
            .SelectedItem.ListSubItems(i - 1).ForeColor = 0
        End If
    Next
    Me.cmdDelete.Caption = "ɾ��(&D)"
    .SelectedItem.ListSubItems(3).Tag = 1
End With
End Sub

Private Sub WriteTag(ByVal bytTag As Byte)
'д���־
'��־����2���б�־Ϊ��0��ǰ    1����
'      ��3���б�־Ϊ��0���κα仯  1����   2ɾ��
    If lvw.SelectedItem Is Nothing Then Exit Sub
    lvw.SelectedItem.ListSubItems(3).Tag = bytTag
End Sub

Private Sub cboChar_Change()
    cmdOK.Enabled = True
    If Not (svrItem Is Nothing) Then WriteBack svrItem
End Sub

Private Sub cboChar_Click()
    cmdOK.Enabled = True
    If Not (svrItem Is Nothing) Then WriteBack svrItem
End Sub

Private Sub cboChar_GotFocus()
    zlControl.TxtSelAll cboChar
End Sub

Private Sub cboChar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    If Check�Ƿ����(UCase(Chr(KeyAscii)), "'") = True Then
        KeyAscii = 0
    End If
    If InStr(UCase(Chr(KeyAscii)), ";") > 0 Then
        KeyAscii = 0
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub

Private Sub cboChar_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(cboChar.Text, 2)
    If Cancel = False Then
        If InStr(cboChar.Text, ";") > 0 Then
            MsgBox "�����зǷ��ַ���", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

Private Sub cboType_Click()
    If Not (lvw.SelectedItem Is Nothing) Then CustomEnabled cboType.Text
    cmdOK.Enabled = True
    If Not (svrItem Is Nothing) Then WriteBack svrItem
End Sub

Private Sub cboType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cmdCancel_Click()
    If cmdOK.Enabled = True Then
        If MsgBox("��ȷ�Ͼ�������������˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    Unload Me
End Sub

Private Sub cmdColor_Click()
    
    With dlg
        .Color = lblColor.BackColor
        .ShowColor
        If lblColor.BackColor <> .Color Then
            lblColor.BackColor = .Color
            cmdOK.Enabled = True
            If Not (svrItem Is Nothing) Then WriteBack svrItem
        End If
    End With
End Sub

Private Sub cmdDelete_Click()
'��־����2���б�־Ϊ��0��ǰ    1����
'      ��3���б�־Ϊ��0���κα仯  1����   2ɾ��
    If lvw.SelectedItem Is Nothing Then Exit Sub
    If lvw.SelectedItem.ListSubItems(3).Tag = 2 Then
        UNItemDel lvw
    Else
        ItemDel lvw
    End If
    If lvw.SelectedItem Is Nothing Then
        cmdDelete.Enabled = False
    Else
        cmdDelete.Enabled = True
        lvw_ItemClick lvw.SelectedItem
    End If
    cmdOK.Enabled = True
End Sub

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdNew_Click()
    Dim itmx As ListItem
    Dim strType As String
On Error GoTo ErrHandle
    
    If lvw.ListItems.Count >= 12 Then
        MsgBox "������Ŀ����ע��Ŀ���̶���Ŀ�������ܳ���12", vbInformation, gstrSysName
        Exit Sub
    End If
    strType = "������Ŀ"
    If Not (lvw.SelectedItem Is Nothing) Then
        If lvw.SelectedItem.SubItems(3) <> "�̶���Ŀ" Then strType = lvw.SelectedItem.SubItems(3)
    End If
    
    
    Set itmx = lvw.ListItems.Add(, , "����Ŀ", "NewItem", "NewItem")
    itmx.SubItems(1) = ""
    itmx.SubItems(2) = ""
    itmx.SubItems(3) = strType
    itmx.SubItems(4) = ""
    itmx.SubItems(5) = ""
    itmx.ListSubItems(1).Tag = "10;300;0"
    '��־����2���б�־Ϊ��0��ǰ    1����
    '      ��3���б�־Ϊ��0���κα仯  1����   2ɾ��
    itmx.ListSubItems(2).Tag = 1
    itmx.ListSubItems(3).Tag = 0
    itmx.Selected = True
    itmx.EnsureVisible
    lvw_ItemClick itmx
    '��־����2���б�־Ϊ��0��ǰ    1����
    '      ��3���б�־Ϊ��0���κα仯  1����   2ɾ��
    itmx.ListSubItems(2).Tag = 1
    itmx.ListSubItems(3).Tag = 0
    cmdDelete.Enabled = True
    cmdOK.Enabled = True
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, j As Long
    Dim itmx As ListItem
    Dim lng��� As Long
    Dim strErr As String
    Dim strTmp As String
    
    On Error GoTo ErrHand
    
    strTmp = ""
    For i = 1 To lvw.ListItems.Count
        Set itmx = lvw.ListItems(i)
        If Trim(itmx.Text) = "˵��" Or Trim(itmx.Text) = "����" Or Trim(itmx.Text) = "�����" Or Trim(itmx.Text) = "��Һ" Or Trim(itmx.Text) = "��ҩ" Or Trim(itmx.Text) = "��Һ" Then
            strTmp = "�� " & i & " �е���Ŀ������ϵͳ�����ظ�������������"
            lvw.ListItems(i).Selected = True
            lvw.ListItems(i).EnsureVisible
            lvw_ItemClick lvw.ListItems(i)
            If Me.txtName.Enabled And Me.txtName.Visible Then Me.txtName.SetFocus
            Exit For
        End If
        If Trim(itmx.Text) = "" Then
            strTmp = "�� " & i & " �е���Ŀ���Ʋ���Ϊ�գ�"
            lvw.ListItems(i).Selected = True
            lvw.ListItems(i).EnsureVisible
            lvw_ItemClick lvw.ListItems(i)
            If Me.txtName.Enabled And Me.txtName.Visible Then Me.txtName.SetFocus
            Exit For
        End If
        For j = 1 To lvw.ListItems.Count
            If Trim(lvw.ListItems(i).Text) = Trim(lvw.ListItems(j).Text) And i <> j Then
                lvw.ListItems(j).Selected = True
                lvw.ListItems(j).EnsureVisible
                lvw_ItemClick lvw.ListItems(j)
                MsgBox "��Ŀ��" & lvw.ListItems(j).Text & "�����ظ���", vbOKOnly + vbInformation, gstrSysName
                If Me.txtName.Enabled And Me.txtName.Visible Then Me.txtName.SetFocus
                Exit Sub
            End If
        Next
        If Trim(itmx.SubItems(2)) = "" Then
            strTmp = "��Ŀ��" & itmx.Text & "���ļ�¼������Ϊ�գ�"
            lvw.ListItems(i).Selected = True
            lvw.ListItems(i).EnsureVisible
            lvw_ItemClick lvw.ListItems(i)
            If Me.cboChar.Enabled And Me.cboChar.Visible Then Me.cboChar.SetFocus
            Exit For
        End If
                        
        If itmx.SubItems(3) = "������Ŀ" And Val(Split(itmx.ListSubItems(1).Tag, ";")(1)) <= (Val(Split(itmx.ListSubItems(1).Tag, ";")(0))) Then
            strTmp = "�ڡ�" & i & "������������Ŀ���ֵ���������Сֵ��"
            lvw.ListItems(i).Selected = True
            lvw.ListItems(i).EnsureVisible
            lvw_ItemClick lvw.ListItems(i)
            If Me.txtMax.Enabled And Me.txtMax.Visible Then Me.txtMax.SetFocus
            Exit For
        End If
    Next
    If strTmp <> "" Then
        MsgBox strTmp, vbOKOnly + vbInformation, gstrSysName
        Exit Sub
    End If
    
    gcnOracle.BeginTrans
    
    For i = 1 To lvw.ListItems.Count
        Set itmx = lvw.ListItems(i)
        If itmx.SubItems(3) = "������Ŀ" Or itmx.SubItems(3) = "��ע��Ŀ" Then
'            If itmx.ListSubItems(2).Tag = 1 Then    '����Ŀ
                '��Ŀ����===��λ===�ַ�===��¼����===���
                '����_IN����λ_IN����¼��_IN����¼��_IN����¼ɫ_IN�����ֵ_IN����Сֵ_IN
                'itmx.ListSubItems(1).Tag = ��Сֵ;���ֵ;��ɫ
'                gstrSql = "ZL_�����¼��Ŀ_INSERT('" & Trim(itmx.Text) & "','" & _
'                        Trim(itmx.SubItems(1)) & "'," & IIf(itmx.SubItems(3) = "������Ŀ", 1, 2) & ",'" & _
'                        Trim(itmx.SubItems(2)) & "'," & Split(itmx.ListSubItems(1).Tag, ";")(2) & "," & _
'                        Split(itmx.ListSubItems(1).Tag, ";")(1) & "," & Split(itmx.ListSubItems(1).Tag, ";")(0) & ")"
'
'                Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
'            Else
            If itmx.ListSubItems(2).Tag = 0 Then    '����Ŀ
                If itmx.ListSubItems(3).Tag = 1 Then
                    '�ؼ���_IN       �������¼��.���%TYPE,
                    '����_IN         �������¼��.��¼��%TYPE,
                    '��¼��_IN       �������¼��.��¼��%TYPE,
                    '��¼ɫ_IN       �������¼��.��¼ɫ%TYPE,
                    '���ֵ_IN       �������¼��.���ֵ%TYPE,
                    '��Сֵ_IN       �������¼��.��Сֵ%TYPE
                    '��Ŀ����,1500,0,1;��λ,900,0,2;�ַ�,600,0,2;��¼����,1200,0,2
                    'itmx.ListSubItems(1).Tag = ��Сֵ;���ֵ;��ɫ
                    gstrSql = "ZL_�����¼��Ŀ_UPDATE(" & itmx.Tag & ",'" & Trim(itmx.Text) & "','" & _
                            Trim(itmx.SubItems(2)) & "'," & Split(itmx.ListSubItems(1).Tag, ";")(2) & "," & _
                            Split(itmx.ListSubItems(1).Tag, ";")(1) & "," & Split(itmx.ListSubItems(1).Tag, ";")(0) & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
'                ElseIf itmx.ListSubItems(3).Tag = 2 Then
'                     '��ĿID_IN���ؼ���_IN
'                    gstrSql = "ZL_�����¼��Ŀ_DELETE(" & itmx.Tag & "," & Trim(itmx.SubItems(4)) & ")"
'                    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
                End If
            End If
'            End If
        End If
    Next
    
    gcnOracle.CommitTrans
    blnOK = True
    cmdOK.Enabled = False
    Unload Me
    Exit Sub
ErrHand:
    strErr = Err.Description
    If InStr(1, strErr, "[ZLSOFT]") > 0 Then
        On Error Resume Next
        strErr = Split(strErr, "[ZLSOFT]")(1)
    End If
    gcnOracle.RollbackTrans
    Call SaveErrLog
    If strErr <> "" Then MsgBox strErr, vbExclamation, gstrSysName
End Sub

Private Sub Form_Load()
    blnOK = False
    Set svrItem = Nothing
    zlControl.LvwSelectColumns lvw, "��Ŀ����,1500,0,1;��λ,900,0,2;�ַ�,600,0,2;��¼����,1200,0,2;���,0,0,2;��,0,0,2", True
    zlControl.LvwFlatColumnHeader lvw
    cboType.AddItem "������Ŀ"
    cboType.AddItem "��ע��Ŀ"
    lvw.Sorted = False
    
    Init
    If lvw.ListItems.Count > 0 Then
        lvw.ListItems(1).Selected = True
        cmdDelete.Enabled = True
    End If
    lvw_ItemClick lvw.SelectedItem
    cmdOK.Enabled = False
End Sub

Private Sub Init()
    Dim i As Long
    Dim itmx As ListItem
    Dim rsTmp As New ADODB.Recordset
On Error GoTo ErrHandle
    
    With rsTmp
        lvw.ListItems.Clear
        '����̶���Ŀ
        strSQL = _
            "SELECT nvl(a.id,0) ��Ŀid,b.��� ��Ŀ��, " & vbCrLf & _
            "    b.��¼�� ��Ŀ��,nvl(a.��λ,'') ��λ, " & vbCrLf & _
            "    b.���ֵ,b.��Сֵ,b.��¼��,b.��¼ɫ " & vbCrLf & _
            " FROM ����������Ŀ a,�������¼�� b " & vbCrLf & _
            " WHERE b.��Ŀid=a.id(+) AND b.����=2 AND b.��¼�� is null and b.��¼�� in ('��  ��','��  ��')" & vbCrLf & _
            " ORDER BY b.���"
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        If .BOF = False Then
            While Not .EOF
                Set itmx = lvw.ListItems.Add(, "K" & rsTmp!��Ŀ��, rsTmp!��Ŀ��, 1, 1)
                itmx.Tag = rsTmp!��Ŀ��
                itmx.SubItems(1) = IIf(IsNull(rsTmp!��λ), "", rsTmp!��λ)
                itmx.SubItems(2) = IIf(IsNull(rsTmp!��¼��), "", rsTmp!��¼��)
                itmx.SubItems(3) = "�̶���Ŀ"
                itmx.SubItems(4) = CStr(rsTmp!��Ŀ��)
                itmx.SubItems(5) = ""
                itmx.ListSubItems(1).Tag = IIf(IsNull(rsTmp!��Сֵ), "10", rsTmp!��Сֵ) & ";" & IIf(IsNull(rsTmp!���ֵ), "300", rsTmp!���ֵ) & ";" & IIf(IsNull(rsTmp!��¼ɫ), 0, rsTmp!��¼ɫ)
                '��־����2���б�־Ϊ��0��ǰ    1����
                '      ��3���б�־Ϊ��0���κα仯  1����   2ɾ��
                itmx.ListSubItems(2).Tag = 0
                itmx.ListSubItems(3).Tag = 0
                .MoveNext
            Wend
        End If
        '��������
        strSQL = _
            "SELECT nvl(a.id,0) ��Ŀid,b.��� ��Ŀ��, " & vbCrLf & _
            "    b.��¼�� ��Ŀ��,nvl(a.��λ,'') ��λ, " & vbCrLf & _
            "    b.���ֵ,b.��Сֵ,b.��¼��,b.��¼ɫ " & vbCrLf & _
            " FROM ����������Ŀ a,�������¼�� b " & vbCrLf & _
            " WHERE b.��Ŀid=a.id(+) AND b.����=2 and  b.��¼�� is  null and not b.��¼�� in ('��  ��','��  ��')" & vbCrLf & _
            " ORDER BY b.���"
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        If .BOF = False Then
            While Not .EOF
                Set itmx = lvw.ListItems.Add(, "K" & rsTmp!��Ŀ��, rsTmp!��Ŀ��, 1, 1)
                itmx.Tag = rsTmp!��Ŀ��
                itmx.SubItems(1) = IIf(IsNull(rsTmp!��λ), "", rsTmp!��λ)
                itmx.SubItems(2) = IIf(IsNull(rsTmp!��¼��), "", rsTmp!��¼��)
                itmx.SubItems(3) = "������Ŀ"
                itmx.SubItems(4) = CStr(rsTmp!��Ŀ��)
                itmx.SubItems(5) = ""
                itmx.ListSubItems(1).Tag = IIf(IsNull(rsTmp!��Сֵ), "10", rsTmp!��Сֵ) & ";" & IIf(IsNull(rsTmp!���ֵ), "300", rsTmp!���ֵ) & ";" & IIf(IsNull(rsTmp!��¼ɫ), 0, rsTmp!��¼ɫ)
                '��־����2���б�־Ϊ��0��ǰ    1����
                '      ��3���б�־Ϊ��0���κα仯  1����   2ɾ��
                itmx.ListSubItems(2).Tag = 0
                itmx.ListSubItems(3).Tag = 0
                .MoveNext
            Wend
        End If
        
        'û�з�������Ŀ
'        '����������
''        strSQL = _
''            "SELECT a.id ��Ŀid,a.С�� ��Ŀ��," & vbCrLf & _
''            "   a.������ ��Ŀ��,a.��λ," & vbCrLf & _
''            "    b.���ֵ,b.��Сֵ,b.��¼��,b.��¼ɫ" & vbCrLf & _
''            "FROM ����������Ŀ a,�������¼�� b" & vbCrLf & _
''            "WHERE b.��Ŀid=a.id AND b.����=2 AND b.��¼��=2" & vbCrLf & _
''            "ORDER BY a.С��"
'        strSQL = _
'            "SELECT nvl(a.id,0) ��Ŀid,b.��� ��Ŀ��, " & vbCrLf & _
'            "    b.��¼�� ��Ŀ��,nvl(a.��λ,'') ��λ, " & vbCrLf & _
'            "    b.���ֵ,b.��Сֵ,b.��¼��,b.��¼ɫ " & vbCrLf & _
'            " FROM ����������Ŀ a,�������¼�� b " & vbCrLf & _
'            " WHERE b.��Ŀid=a.id(+) AND b.����=2 and  b.��¼�� is  null and not b.��¼�� in ('��  ��','��  ��')" & vbCrLf & _
'            " ORDER BY b.���"
'        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
'        If .BOF = False Then
'            While Not .EOF
'                Set itmx = lvw.ListItems.Add(, "K" & rsTmp!��Ŀ��, rsTmp!��Ŀ��, 1, 1)
'                itmx.Tag = rsTmp!��ĿID
'                itmx.SubItems(1) = IIf(IsNull(rsTmp!��λ), "", rsTmp!��λ)
'                itmx.SubItems(2) = IIf(IsNull(rsTmp!��¼��), "", rsTmp!��¼��)
'                itmx.SubItems(3) = "��ע��Ŀ"
'                itmx.SubItems(4) = CStr(rsTmp!��Ŀ��)
'                itmx.SubItems(5) = ""
'                itmx.ListSubItems(1).Tag = IIf(IsNull(rsTmp!��Сֵ), "10", rsTmp!��Сֵ) & ";" & IIf(IsNull(rsTmp!���ֵ), "300", rsTmp!���ֵ) & ";" & IIf(IsNull(rsTmp!��¼ɫ), 0, rsTmp!��¼ɫ)
'                '��־����2���б�־Ϊ��0��ǰ    1����
'                '      ��3���б�־Ϊ��0���κα仯  1����   2ɾ��
'                itmx.ListSubItems(2).Tag = 0
'                itmx.ListSubItems(3).Tag = 0
'                .MoveNext
'            Wend
'        End If
        
        For i = 65 To 90
            cboChar.AddItem Chr(i)
        Next
        cboChar.AddItem "��"
        cboChar.AddItem "��"
        cboChar.AddItem "+"
        cboChar.AddItem "*"
        cboChar.AddItem "��"
        cboChar.Text = "A"
        If cboType.ListCount > 0 Then cboType.ListIndex = 0
        Me.txtMax.Text = "20"
        Me.txtMin.Text = "0"
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    RefreshItem Item
    If Item Is Nothing Then
        CustomEnabled "��ע��Ŀ"
    Else
        CustomEnabled Item.SubItems(3)
    End If
End Sub

Private Sub RefreshItem(ByVal Item As MSComctlLib.ListItem)
    Dim svrOK As Boolean
On Error GoTo ErrHandle
    
    If Item Is Nothing Then Exit Sub
    Set svrItem = Nothing
    svrOK = cmdOK.Enabled
    cboType.Clear
    txtName.Text = Item.Text
    txtUnit.Text = Item.SubItems(1)
    cboChar.Text = Item.SubItems(2)
    If Item.SubItems(3) = "�̶���Ŀ" Then
        cboType.AddItem "�̶���Ŀ"
        Me.cmdDelete.Caption = "ɾ��(&D)"
    Else
        cboType.AddItem "������Ŀ"
        cboType.AddItem "��ע��Ŀ"
        If Item.ListSubItems(3).Tag = 2 Then
            Me.cmdDelete.Caption = "��ɾ��(&D)"
        Else
            Me.cmdDelete.Caption = "ɾ��(&D)"
        End If
    End If
    cboType.Text = Item.SubItems(3)
    
    txtMin.Text = Split(Item.ListSubItems(1).Tag, ";")(0)
    txtMax.Text = Split(Item.ListSubItems(1).Tag, ";")(1)
    udMin.Value = IIf(Val(txtMin.Text) < 0, 0, Val(txtMin.Text))
    udMax.Value = IIf(Val(txtMax.Text) < 0, 0, Val(txtMax.Text))
    lblColor.BackColor = Val(Split(Item.ListSubItems(1).Tag, ";")(2))
    Set svrItem = Item
    cmdOK.Enabled = svrOK
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub WriteBack(ByVal Item As MSComctlLib.ListItem)
    Item.Text = txtName.Text
    Item.SubItems(1) = txtUnit.Text
    Item.SubItems(2) = cboChar.Text
    Item.SubItems(3) = cboType.Text
    Item.ListSubItems(1).Tag = txtMin.Text & ";" & txtMax.Text & ";" & lblColor.BackColor
    Call WriteTag(1)
End Sub

Private Sub CustomEnabled(ByVal strFlag As String)
    If lvw.SelectedItem Is Nothing Then
        fraData.Enabled = False
            '��Ŀ����
            Me.Label7.Enabled = False
            '��λ
            Me.Label1.Enabled = False
            '��¼����
            Me.Label8.Enabled = False
            '���ֵ
            Me.Label5.Enabled = False
            '��Сֵ
            Me.Label4.Enabled = False
        fraDisplay.Enabled = False
            '��¼��
            Me.Label2.Enabled = False
            '��¼ɫ
            Me.Label3.Enabled = False
        txtName.Enabled = False
        txtUnit.Enabled = False
        txtMin.Enabled = False
        txtMax.Enabled = False
        cboType.Enabled = False
        udMax.Enabled = False
        udMin.Enabled = False
        cmdColor.Enabled = False
        cboChar.Enabled = False
        cmdDelete.Enabled = False
        Exit Sub
    End If
    fraData.Enabled = True
        '��Ŀ����
        Me.Label7.Enabled = True
        '��λ
        Me.Label1.Enabled = False
        '��¼����
        Me.Label8.Enabled = False
        '���ֵ
        Me.Label5.Enabled = True
        '��Сֵ
        Me.Label4.Enabled = True
    fraDisplay.Enabled = True
        '��¼��
        Me.Label2.Enabled = True
        '��¼ɫ
        Me.Label3.Enabled = True
    txtName.Enabled = True
    txtUnit.Enabled = False
    txtMin.Enabled = True
    txtMax.Enabled = True
    '��¼����
    cboType.Enabled = False
    udMax.Enabled = True
    udMin.Enabled = True
    cmdColor.Enabled = True
    cboChar.Enabled = True
    cmdDelete.Enabled = False
    
    If lvw.SelectedItem.ListSubItems(3).Tag = 2 Then
        fraData.Enabled = False
            '��Ŀ����
            Me.Label7.Enabled = False
            '��λ
            Me.Label1.Enabled = False
            '��¼����
            Me.Label8.Enabled = False
            '���ֵ
            Me.Label5.Enabled = False
            '��Сֵ
            Me.Label4.Enabled = False
        fraDisplay.Enabled = False
            '��¼��
            Me.Label2.Enabled = False
            '��¼ɫ
            Me.Label3.Enabled = False
        txtName.Enabled = False
        txtUnit.Enabled = False
        txtMin.Enabled = False
        txtMax.Enabled = False
        cboType.Enabled = False
        udMax.Enabled = False
        udMin.Enabled = False
        cmdColor.Enabled = False
        cboChar.Enabled = False
    Else
        If strFlag = "�̶���Ŀ" Then
            fraData.Enabled = False
                '��Ŀ����
                Me.Label7.Enabled = False
                '��λ
                Me.Label1.Enabled = False
                '��¼����
                Me.Label8.Enabled = False
                '���ֵ
                Me.Label5.Enabled = False
                '��Сֵ
                Me.Label4.Enabled = False
                txtName.Enabled = False
                txtUnit.Enabled = False
                txtMin.Enabled = False
                txtMax.Enabled = False
                cboType.Enabled = False
                udMax.Enabled = False
                udMin.Enabled = False
                cmdDelete.Enabled = False
        ElseIf strFlag = "��ע��Ŀ" Then
            txtUnit.Enabled = False
            txtMin.Enabled = False
            txtMax.Enabled = False
            udMax.Enabled = False
            udMin.Enabled = False
            '���ֵ
            Me.Label5.Enabled = False
            '��Сֵ
            Me.Label4.Enabled = False
        End If
    End If
End Sub

Private Sub txtMax_Change()
    cmdOK.Enabled = True
    If Not (svrItem Is Nothing) Then WriteBack svrItem
    If Val(txtMax.Text) > 300 Then txtMax.Text = "300": txtMax.SelStart = Len(txtMax.Text)
    udMin.Max = Val(txtMax.Text)
End Sub

Private Sub txtMax_GotFocus()
    zlControl.TxtSelAll txtMax
End Sub

Private Sub txtMax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    If Check�Ƿ����(UCase(Chr(KeyAscii)), "������") = True Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtMax_Validate(Cancel As Boolean)
On Error GoTo ErrHandle
    If IsNumeric(txtMax.Text) = False Then
        MsgBox "��������ȷ���֣�", vbInformation, gstrSysName
        Cancel = True
    End If
    If Val(txtMin.Text) >= Val(txtMax.Text) Then
        MsgBox "����ֵ��Ч����СֵӦС�����ֵ��", vbInformation, gstrSysName
        Cancel = True
    End If
    If Val(txtMax.Text) > 300 Or Val(txtMax.Text) < 20 Then
        MsgBox "����ֵ��Ч��ֻ������20��300֮�������", vbInformation, gstrSysName
        Cancel = True
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtMin_Change()
    cmdOK.Enabled = True
    If Not (svrItem Is Nothing) Then WriteBack svrItem
    If Val(txtMin.Text) > 300 Then txtMin.Text = "300": txtMin.SelStart = Len(txtMin.Text)
    udMax.Min = Format(txtMin.Text)
End Sub

Private Sub txtMin_GotFocus()
    zlControl.TxtSelAll txtMin
End Sub

Private Sub txtMin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    If Check�Ƿ����(UCase(Chr(KeyAscii)), "������") = True Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtMin_Validate(Cancel As Boolean)
    If IsNumeric(txtMin.Text) = False Then
        MsgBox "��������ȷ���֣�", vbInformation, gstrSysName
        Cancel = True
    End If
    If Val(txtMin.Text) >= Val(txtMax.Text) Then
        MsgBox "����ֵ��Ч����СֵӦС�����ֵ��", vbInformation, gstrSysName
        Cancel = True
    End If
    If Val(txtMin.Text) > 300 Then
        MsgBox "����ֵ��Ч��ֻ������0��300֮�������", vbInformation, gstrSysName
        Cancel = True
    End If
End Sub

Private Sub txtName_Change()
    cmdOK.Enabled = True
    If Not (svrItem Is Nothing) Then WriteBack svrItem
End Sub

Private Sub txtName_GotFocus()
    zlControl.TxtSelAll txtName
    zlCommFun.OpenIme True
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    If Check�Ƿ����(UCase(Chr(KeyAscii)), "'") = True Then
        KeyAscii = 0
    End If
    If Check�Ƿ����(UCase(Chr(KeyAscii)), ";") = True Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtName_LostFocus()
    zlCommFun.OpenIme
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txtName.Text, txtName.MaxLength)
    If Cancel = False Then
        If InStr(txtName.Text, ";") > 0 Then
            MsgBox "�����зǷ��ַ���", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

Private Sub txtUnit_Change()
    cmdOK.Enabled = True
    If Not (svrItem Is Nothing) Then WriteBack svrItem
End Sub

Private Sub txtUnit_GotFocus()
    zlControl.TxtSelAll txtUnit
End Sub

Private Sub txtUnit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    If Check�Ƿ����(UCase(Chr(KeyAscii)), "'") = True Then
        KeyAscii = 0
    End If
    If Check�Ƿ����(UCase(Chr(KeyAscii)), ";") = True Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtUnit_Validate(Cancel As Boolean)
On Error GoTo ErrHandle
    Cancel = Not StrIsValid(txtUnit.Text, txtUnit.MaxLength)
    If Cancel = False Then
        If InStr(txtUnit.Text, ";") > 0 Then
            MsgBox "�����зǷ��ַ���", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

