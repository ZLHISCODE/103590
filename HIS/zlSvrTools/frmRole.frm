VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRole 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "��ɫ��Ȩ����"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9735
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmRole.frx":0000
   ScaleHeight     =   6240
   ScaleWidth      =   9735
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdRolesReset 
      Caption         =   "�������н�ɫ"
      Height          =   350
      Left            =   7575
      TabIndex        =   24
      Top             =   3120
      Width           =   1875
   End
   Begin VB.CheckBox chkOnlyShowNOSystem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ֻ��ʾδ��ϵͳ��ɫ"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   7575
      TabIndex        =   23
      Top             =   225
      Width           =   2040
   End
   Begin VB.CommandButton cmdSystemMove 
      Caption         =   "�Ƶ�ϵͳ(&S)"
      Height          =   350
      Left            =   7575
      TabIndex        =   22
      Top             =   2565
      Width           =   1875
   End
   Begin VB.CommandButton cmdRoleMove 
      Caption         =   "�Ƶ�����(&M)"
      Height          =   350
      Left            =   7575
      TabIndex        =   21
      Top             =   2205
      Width           =   1875
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "ɾ����ɫ(&D)"
      Height          =   350
      Left            =   7575
      TabIndex        =   20
      Top             =   1560
      Width           =   1875
   End
   Begin VB.CommandButton cmdDeleteGroup 
      Caption         =   "ɾ����(&R)"
      Height          =   350
      Left            =   2325
      TabIndex        =   19
      Top             =   3900
      Width           =   1100
   End
   Begin VB.CommandButton cmdModifyGroup 
      Caption         =   "�޸���(&E)"
      Height          =   350
      Left            =   1230
      TabIndex        =   18
      Top             =   3900
      Width           =   1100
   End
   Begin VB.PictureBox picHLine 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   2070
      MousePointer    =   7  'Size N S
      ScaleHeight     =   90
      ScaleWidth      =   5835
      TabIndex        =   16
      Top             =   4350
      Width           =   5835
   End
   Begin VB.CommandButton cmdNewGroup 
      Caption         =   "�½���(&N)"
      Height          =   350
      Left            =   135
      TabIndex        =   5
      Top             =   3900
      Width           =   1100
   End
   Begin VB.CommandButton cmdGrantAll 
      Caption         =   "�ָ����н�ɫ��Ȩ��"
      Height          =   350
      Left            =   7575
      TabIndex        =   8
      Top             =   3465
      Width           =   1875
   End
   Begin VB.CommandButton cmdGrant 
      Caption         =   "��ɫ��Ȩ(&G)"
      Height          =   350
      Left            =   7575
      TabIndex        =   7
      Top             =   1215
      Width           =   1875
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "���ƽ�ɫ(&C)"
      Height          =   350
      Left            =   7575
      TabIndex        =   9
      Top             =   1875
      Width           =   1875
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "���ӽ�ɫ(&A)"
      Height          =   350
      Left            =   7575
      TabIndex        =   6
      Top             =   870
      Width           =   1875
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   7560
      TabIndex        =   3
      Top             =   555
      Width           =   1875
   End
   Begin MSComctlLib.TreeView tvwGroups 
      Height          =   2955
      Left            =   150
      TabIndex        =   1
      Top             =   870
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   5212
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   295
      Style           =   7
      ImageList       =   "img16"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   5880
      Top             =   3915
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
            Picture         =   "frmRole.frx":803A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdUser 
      Cancel          =   -1  'True
      Caption         =   "�޸Ľ�ɫ����Ȩ�û�"
      Height          =   350
      Left            =   7575
      TabIndex        =   14
      Top             =   4155
      Width           =   1875
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "�޸�ģ���ʹ��Ȩ��"
      Height          =   350
      Left            =   7575
      TabIndex        =   13
      Top             =   3810
      Width           =   1875
   End
   Begin VB.ComboBox cmbSystem 
      Height          =   300
      Left            =   4755
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   555
      Width           =   3915
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   4620
      Top             =   3915
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRole.frx":85D4
            Key             =   "Role"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRole.frx":92AE
            Key             =   "Role_Moved"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwRole 
      Height          =   2970
      Left            =   3510
      TabIndex        =   4
      Top             =   870
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   5239
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils32"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "_��ɫ"
         Object.Tag             =   "��ɫ"
         Text            =   "��ɫ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Grantee"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Admin_Option"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Group"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "System"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "SystemName"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvwModule 
      Height          =   1605
      Left            =   135
      TabIndex        =   15
      Top             =   4575
      Width           =   8580
      _ExtentX        =   15134
      _ExtentY        =   2831
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ɫ��Ȩ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   105
      TabIndex        =   17
      Top             =   150
      Width           =   1440
   End
   Begin VB.Label lblRoleGroup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "��ɫ����Ϣ"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   135
      TabIndex        =   0
      Top             =   630
      Width           =   1530
   End
   Begin VB.Label lblSearch 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   7110
      TabIndex        =   2
      Top             =   615
      Width           =   360
   End
   Begin VB.Label lblModule 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����Ȩ�������"
      Height          =   180
      Left            =   135
      TabIndex        =   10
      Top             =   4320
      Width           =   1440
   End
   Begin VB.Label lblSystem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ɫ����ϵͳ"
      Height          =   180
      Left            =   3510
      TabIndex        =   11
      Top             =   615
      Width           =   1200
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "�����˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuAddGroups 
         Caption         =   "�½���(&N)"
      End
      Begin VB.Menu mnuPopuModify 
         Caption         =   "�޸���(&M)"
      End
      Begin VB.Menu mnuPopuDeleteGroups 
         Caption         =   "ɾ����(&D)"
      End
   End
   Begin VB.Menu mnuPopuRole 
      Caption         =   "�����˵���ɫ"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuRoleAdd 
         Caption         =   "���ӽ�ɫ(&N)"
      End
      Begin VB.Menu mnuPopuRoleDelete 
         Caption         =   "ɾ����ɫ(&M)"
      End
      Begin VB.Menu mnuPopuRoleMove1 
         Caption         =   "��ɫ�Ƶ�����(&M)��"
         Begin VB.Menu mnuPopuRoleMoveGroups 
            Caption         =   "��1"
            Index           =   0
         End
      End
      Begin VB.Menu mnuPopuRoleMove2 
         Caption         =   "��ɫ�Ƶ�ϵͳ(&S)��"
         Begin VB.Menu mnuPopuRoleMoveSystems 
            Caption         =   "��2"
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "frmRole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsRole As ADODB.Recordset
Private mblnFirst As Boolean
Private mblnMoveTop As Boolean
Private msngPreHeigt As Single
Private mobjTip  As clsTipSwap           '������ʾ�����
Private mfrmGrant As frmRoleGrant
Private mstrSystemsName As String        '��ɫ�ƶ�ϵͳʱ��Ŀ��ϵͳ������չʾ��ʾ����

Private Enum lvwModuleHeader
    LH_ϵͳ = 0
    LH_��� = 1
    LH_���ܻ���� = 2
    LH_˵�� = 3
    LH_��Ȩ���� = 4
    LH_ϵͳ�� = 5
End Enum

Private Enum LvwRoleHeader
    LRH_��ɫ = 0
    LRH_Grantee = 1
    LRH_Admin_Option = 2
    LRH_Group = 3
    LRH_System = 4
    LRH_SystemName = 5
End Enum

Private Sub chkOnlyShowNOSystem_Click()
    If tvwGroups.SelectedItem Is Nothing Then Exit Sub
    tvwGroups_NodeClick tvwGroups.SelectedItem
End Sub

Private Sub cmdAdd_Click()
    Dim cnTemp As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim strRoleName As String
    Dim blnLimited As Boolean
    Dim lst As ListItem
    Dim str������() As String
    Dim strSQL As String, rsTmp As ADODB.Recordset, lngCount As Long
    Dim strUserName As String
    
    On Error GoTo ErrHandle
    If cmbSystem.ItemData(cmbSystem.ListIndex) <> -1 And glngSysNo = -1 Then
        MsgBox "���ڱ�ϵͳ�´����˽�ɫ��" & vbNewLine & "��ô�ý�ɫ������ϵͳ��������", vbInformation, gstrSysName
    End If
    '���û�ӵ�еĽ�ɫ�����ﵽ148��ʱ���û���¼ʱ����ʾ����
    gstrSQL = "Select Count(*) as ���� From DBA_Role_Privs Where Grantee='" & gstrUserName & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open gstrSQL, gcnOracle, adOpenKeyset, adLockReadOnly
    If Nvl(rsTemp!����, 0) >= 148 Then
        If Not gblnOwner Then
            MsgBox "��ɫ�����Ѵﵽ������ƣ����������ӡ�", vbInformation, gstrSysName
            Exit Sub
        Else
            '�����߽�ɫ�����ﵽ����ʱ������Systeme�û�����
            'SYSTEM�������Ľ�ɫ��������������
            gstrSQL = "Select Count(*) as ���� From DBA_Role_Privs Where Grantee='SYSTEM'"
            Set rsTemp = New ADODB.Recordset
            rsTemp.CursorLocation = adUseClient
            rsTemp.Open gstrSQL, gcnOracle, adOpenKeyset, adLockReadOnly
            If Nvl(rsTemp!����, 0) >= 148 Then
                MsgBox "��ɫ�����Ѵﵽ������ƣ����������ӡ�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        blnLimited = True
    End If
    
    strRoleName = frmNameEdit.GetName(name��ɫ)
    If strRoleName = "" Then Exit Sub
    strRoleName = "ZL_" & UCase(Trim(strRoleName))
    strUserName = gstrUserName
    Set cnTemp = gcnOracle
    If blnLimited And gblnOwner Then
        Set gcnSystem = GetConnection("SYSTEM")
        If gcnSystem Is Nothing Then Exit Sub
        strUserName = "SYSTEM"
        Set cnTemp = gcnSystem
    End If
    
    On Error Resume Next
    cnTemp.Execute "Create Role " & strRoleName & " Not Identified"
    
    If err <> 0 Then
        MsgBox "��������������������߽�ɫ�����������ݿ�Ĳ�������" & vbCrLf & _
                "(���޸����ݿ���������������ɫ��Ŀ)�����½�ɫ����ʧ�ܡ�", vbExclamation, gstrSysName
        Call SetEnable
    Else
        On Error GoTo ErrHandle
        '����ɫ��Ϣͬ�����뵽Zlroles����
        gstrSQL = "zltools.Zl_Zlroles_Edit(1,'" & strRoleName & "'" & IIf(cmbSystem.ItemData(cmbSystem.ListIndex) = -1, "", "," & cmbSystem.ItemData(cmbSystem.ListIndex)) & ")"
        ExecuteProcedure gstrSQL, "������ɫ��ϵͳ�Ķ�Ӧ��ϵ"
        
        '������Ҫ������־
        Call SaveAuditLog(1, "���ӽ�ɫ", Split(strRoleName, "_")(1))
        
        strSQL = "Select Distinct s.������ From All_Tables t, Zlsystems s Where t.Table_Name = '���ű�' And t.Owner = s.������"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, App.Title)
        ReDim str������(0 To rsTmp.RecordCount)
        Do While Not rsTmp.EOF
            str������(lngCount) = rsTmp!������
            lngCount = lngCount + 1
            rsTmp.MoveNext
        Loop
        Call GrantSpecialToRole(cnTemp, strRoleName, False, str������, True)
        If tvwGroups.SelectedItem Is Nothing Then
        ElseIf tvwGroups.SelectedItem.Key <> "Root" And tvwGroups.SelectedItem.Key <> "unGroup" Then
            '���˺�:20070615����
            '���̲���:zlTools.b_Rolegroupmgr.RoletoRolegroup
            '        ����_In In ZlRolegroups.����%Type,
            '        ��ɫ_In In ZlRolegroups.��ɫ%Type := Null
            gstrSQL = "zlTools.b_Rolegroupmgr.RoleToRoleGroup("
            gstrSQL = gstrSQL & "'" & Mid(tvwGroups.SelectedItem.Key, 2) & "',"
            gstrSQL = gstrSQL & "'" & strRoleName & "')"
            ExecuteProcedure gstrSQL, Me.Caption
        End If
        Set lst = lvwRole.ListItems.Add(, strRoleName, Mid(strRoleName, 4), "Role", "Role")
        If Not lst Is Nothing Then
            lst.SubItems(LRH_Grantee) = strUserName
            lst.SubItems(LRH_Admin_Option) = "YES"
        End If
        lst.Selected = True
        Call InitRoleData
        Call lvwRole_ItemClick(lst)
    End If
    Exit Sub
ErrHandle:
    MsgBox "����" & err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
End Sub

Private Sub DeleteRole()
    '------------------------------------------------------------------------------------------------------------------------------------------------
    '����:ɾ����ɫ
    '����:���˺�
    '����:2007/06/15
    '------------------------------------------------------------------------------------------------------------------------------------------------
    Dim strRoleName As String
    Dim intIndex As Integer
    Dim strSQL As String, strUserList As String
    Dim i As Long
    Dim rsTemp As ADODB.Recordset
    Dim strRemarks As String
    
    On Error GoTo ErrHandle
    
    If lvwRole.SelectedItem Is Nothing Then Exit Sub
    
    strRoleName = lvwRole.SelectedItem.Key
    intIndex = lvwRole.SelectedItem.Index
    
    If MsgBox("���Ҫɾ����ɫ��" & lvwRole.SelectedItem.Text & "����", vbDefaultButton2 Or vbQuestion Or vbYesNo, gstrSysName) = vbNo Then Exit Sub
    '�жϸý�ɫ�Ƿ����ڱ�ʹ��
    strSQL = "Select Grantee �û��� From Dba_Role_Privs Where Granted_Role = [1]"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, strRoleName)
    If rsTemp.RecordCount > 1 Then
        '˵���ý�ɫ���ڱ�ʹ�ã����ܱ�ɾ��
        For i = 1 To rsTemp.RecordCount
            If i > 3 Then Exit For
            strUserList = strUserList & "��" & rsTemp!�û��� & "��" & vbNewLine
            rsTemp.MoveNext
        Next
        If rsTemp.RecordCount > 3 Then
            MsgBox "�ý�ɫ���ڱ�" & vbNewLine & strUserList & "��" & rsTemp.RecordCount & _
            "���û�ʹ�ã�Ҫɾ���ý�ɫ�����޸������û��ĵĽ�ɫ��", vbInformation, gstrSysName
        Else
            MsgBox "�ý�ɫ���ڱ��û�" & vbNewLine & strUserList & _
            "ʹ�ã�Ҫɾ���ý�ɫ�����޸������û��ĵĽ�ɫ��", vbInformation, gstrSysName
        End If
        Exit Sub
    End If
    '��֤��ݲ��������˵��
    strRemarks = "ɾ����ɫ��" & lvwRole.SelectedItem.Text
    If Not CheckAuditStatus("0401", "ɾ����ɫ", strRemarks) Then Exit Sub
    Screen.MousePointer = 11
    If lvwRole.SelectedItem.SubItems(LRH_Grantee) = UCase(gstrUserName) _
        And lvwRole.SelectedItem.SubItems(LRH_Admin_Option) = "YES" Then
        gcnOracle.Execute "Drop Role " & strRoleName
    Else
        Set gcnSystem = GetConnection("SYSTEM")
        If gcnSystem Is Nothing Then Exit Sub
        gcnSystem.Execute "Drop Role " & strRoleName
    End If
    
    gstrSQL = "zlTools.b_Rolegroupmgr.Role_Delete('" & UCase(strRoleName) & "')"
    ExecuteProcedure gstrSQL, Me.Caption
    
    gstrSQL = "zltools.Zl_Zlroles_Edit(3,'" & strRoleName & "')"
    ExecuteProcedure gstrSQL, "ɾ����ɫ��ϵͳ�Ķ�Ӧ��ϵ"
    
    '������Ҫ������־
    Call SaveAuditLog(3, "ɾ����ɫ", lvwRole.SelectedItem.Text, strRemarks)
    lvwRole.ListItems.Remove intIndex
    If lvwRole.ListItems.Count > 0 Then
        If intIndex > lvwRole.ListItems.Count Then
            intIndex = lvwRole.ListItems.Count
        End If
        lvwRole.ListItems(intIndex).Selected = True
    End If
    Call InitRoleData
    Call FillModule
    Call SetEnable
    Screen.MousePointer = 0
    Exit Sub
ErrHandle:
    Screen.MousePointer = 0
    MsgBox "����" & err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
End Sub

Private Sub cmdCopy_Click()
    '------------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ƽ�ɫ
    '����:���˺�
    '����:2007/06/15
    '------------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSourceRole As String
    Dim cnTemp As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim strRoleName As String
    Dim blnLimited As Boolean
    Dim objItem As ListItem
    Dim strUserName As String
    Dim strKey As String
    
    On Error GoTo ErrHandle
    
    If lvwRole.SelectedItem Is Nothing Then Exit Sub
    If cmbSystem.ItemData(cmbSystem.ListIndex) <> -1 And glngSysNo = -1 Then
        MsgBox "���ڱ�ϵͳ�´����˽�ɫ��" & vbNewLine & "��ô�ý�ɫ������ϵͳ��������", vbInformation, gstrSysName
    End If
    '���û�ӵ�еĽ�ɫ�����ﵽ148��ʱ���û���¼ʱ����ʾ����
    gstrSQL = "Select Count(*) as ���� From DBA_Role_Privs Where Grantee='" & gstrUserName & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open gstrSQL, gcnOracle, adOpenKeyset, adLockReadOnly
    If Nvl(rsTemp!����, 0) >= 148 Then
        If Not gblnOwner Then
            MsgBox "��ɫ�����Ѵﵽ������ƣ����������ӡ�", vbInformation, gstrSysName
            Exit Sub
        Else
            If Not CheckRushHours("0401", "���ƽ�ɫ") Then
                Exit Sub
            End If
            '�����߽�ɫ�����ﵽ����ʱ������Systeme�û�����
            'SYSTEM�������Ľ�ɫ��������������
            gstrSQL = "Select Count(*) as ���� From DBA_Role_Privs Where Grantee='SYSTEM'"
            Set rsTemp = New ADODB.Recordset
            rsTemp.CursorLocation = adUseClient
            rsTemp.Open gstrSQL, gcnOracle, adOpenKeyset, adLockReadOnly
            If Nvl(rsTemp!����, 0) >= 148 Then
                MsgBox "��ɫ�����Ѵﵽ������ƣ����������ӡ�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        blnLimited = True
    End If
    
    strSourceRole = lvwRole.SelectedItem.Key
    strRoleName = frmNameEdit.GetName(name��ɫ)
    If strRoleName = "" Then Exit Sub
    strRoleName = "ZL_" & UCase(Trim(strRoleName))
    strUserName = gstrUserName
    Set cnTemp = gcnOracle
    If blnLimited And gblnOwner Then
        Set gcnSystem = GetConnection("SYSTEM")
        If gcnSystem Is Nothing Then Exit Sub
        strUserName = "SYSTEM"
        Set cnTemp = gcnSystem
    End If
 
    If Not CopyRole(cnTemp, strSourceRole, strRoleName) Then Exit Sub
    
    '������Ȩ
    Call RoleGrant(strRoleName)
    '������Ҫ������־
    Call SaveAuditLog(1, "���ƽ�ɫ", "��" & Split(strSourceRole, "_")(1) & "���Ƶõ�" & Split(strRoleName, "_")(1))
    Set objItem = lvwRole.ListItems.Add(, strRoleName, Mid(strRoleName, 4), "Role", "Role")
    If Not objItem Is Nothing Then
        objItem.SubItems(LRH_Grantee) = strUserName
        objItem.SubItems(LRH_Admin_Option) = "YES"
    End If
    Call InitRoleData
    
    strKey = lvwRole.SelectedItem.Key
    err = 0: On Error Resume Next
    objItem.Selected = True
    If err = 0 Then
        lvwRole.ListItems(strKey).Selected = False
    End If
    Call lvwRole_ItemClick(lvwRole.SelectedItem)
    
    Exit Sub
ErrHandle:
    MsgBox "����" & err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
End Sub

Private Function CopyRole(cnTemp As ADODB.Connection, ByVal strSourceRole As String, ByVal strTargetRole As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ɫȨ�޳��µĽ�ɫȨ��
    '����:strSourceRole-Դ��ɫ
    '     strTargetRole-Ŀ���ɫ
    '����:���Ƴɹ�,����true,����False
    '����:���˺�
    '����:2007/07/18
    '------------------------------------------------------------------------------------------------------------------------------------------------
    Dim str������() As String
    Dim strSQL As String, rsTmp As ADODB.Recordset, lngCount As Long
    
    err = 0: On Error Resume Next
    cnTemp.Execute "Create Role " & strTargetRole & " Not Identified"
    If err <> 0 Then
        MsgBox "���������������������Ŀ���ɫ�����������ݿ�Ĳ�������" & vbCrLf & _
                "(���޸����ݿ���������������ɫ��Ŀ)�����½�ɫ����ʧ�ܡ�", vbExclamation, gstrSysName
        Exit Function
    End If
    err = 0: On Error GoTo errHand:
    strSQL = "Select Distinct s.������ From All_Tables t, Zlsystems s Where t.Table_Name = '���ű�' And t.Owner = s.������"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, App.Title)
    ReDim str������(0 To rsTmp.RecordCount)
    Do While Not rsTmp.EOF
        str������(lngCount) = rsTmp!������
        lngCount = lngCount + 1
        rsTmp.MoveNext
    Loop
    Call GrantSpecialToRole(cnTemp, strTargetRole, False, str������, True)
    '����:zlTools.b_Rolegroupmgr.Role_Copy
    '    Դ��ɫ_In   In zlRoleGrant.��ɫ%Type,
    '    Ŀ���ɫ_In In zlRoleGrant.��ɫ%Type
    gstrSQL = "zlTools.b_Rolegroupmgr.Role_Copy("
    gstrSQL = gstrSQL & "'" & UCase(strSourceRole) & "',"
    gstrSQL = gstrSQL & "'" & UCase(strTargetRole) & "')"
    ExecuteProcedure gstrSQL, Me.Caption
    '����ɫ��Ϣͬ�����뵽Zlroles����
    gstrSQL = "zltools.Zl_Zlroles_Edit(1,'" & strTargetRole & "'" & IIf(cmbSystem.ItemData(cmbSystem.ListIndex) = -1, "", "," & cmbSystem.ItemData(cmbSystem.ListIndex)) & ")"
    ExecuteProcedure gstrSQL, "������ɫ��ϵͳ�Ķ�Ӧ��ϵ"
    CopyRole = True
    Exit Function
errHand:
    Call ShowErrHand
End Function

Private Function RoleGrant(ByVal str��ɫ As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ָ���Ľ�ɫ������Ȩ
    '����:str��ɫ-��ɫ
    '����:���Ƴɹ�,����true,����False
    '����:���˺�
    '����:2007/07/18
    '------------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsUser As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    
    Dim str������() As String
    Dim lngCount As Long
    
    Me.MousePointer = vbHourglass
    On Error GoTo ErrHandle
    
    
    '�õ��������н�����Ȩ����������
    rsUser.CursorLocation = adUseClient
    gstrSQL = "select distinct S.������ from all_tables T,zlsystems S where T.table_name='���ű�' And T.OWNER=S.������"
    rsUser.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    
    ReDim str������(0 To rsUser.RecordCount)
    Do Until rsUser.EOF
        str������(lngCount) = rsUser("������")
        lngCount = lngCount + 1
        rsUser.MoveNext
    Loop
    rsUser.Close

    '����Ȩ�ޱ�����д��Ȩ��
    Dim objclsPrivilege As New clsPrivilege
    Call objclsPrivilege.InitOracle(gcnOracle)
    Call objclsPrivilege.ReGrantPrivilege(str��ɫ, str������)
    Set objclsPrivilege = Nothing
    
    MousePointer = 0
    RoleGrant = True
    Exit Function
ErrHandle:
    MousePointer = 0
    MsgBox "��ǰ�û���Ȩ�޲�����ɱ�������", vbInformation, gstrSysName
End Function

Private Sub cmdDel_Click()
    'ɾ����ɫ
    If cmdAdd.Enabled = False Then Exit Sub
    Call DeleteRole
End Sub

Private Sub cmdDeleteGroup_Click()
    'ɾ����
    If tvwGroups.SelectedItem Is Nothing Then Exit Sub
    If tvwGroups.SelectedItem.Key <> "Root" And tvwGroups.SelectedItem.Key <> "unGroup" Then
        Call DeleteRoleGroups
        Call LoadMenus
    End If
End Sub

Private Sub cmdGrant_Click()
    If Not CheckRushHours("0401", "��ɫ��Ȩ") Then
        Exit Sub
    End If
    If mfrmGrant Is Nothing Then
        Set mfrmGrant = New frmRoleGrant
    End If
    If mfrmGrant.GrantToRole(lvwRole.SelectedItem.Key) = True Then
        Call FillModule
    End If
End Sub

Private Sub cmdGrantAll_Click()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lst As ListItem, lngCount As Long
    Dim str������() As String
    
    If MsgBox("�����ܸ���Ӧ��ϵͳ��������н�ɫ�������ݿ��м�鲢���䴴����ɫ������Ӧ��ϵͳ�Ĺ�����������Ȩ��,�Լ�����������ݿ����ķ���Ȩ�ޡ�" & vbCrLf & _
          "�������ݿ���ɾ���˽�ɫ�����߰��û�ģʽ�ָ�����ʱ��ִ�д˲�����������һ�µ����ݣ��Լ���Ӧ��ϵͳ��ɫ�Ķ���������Ȩ��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    If Not CheckRushHours("0401", "�ָ����н�ɫ��Ȩ��") Then
        Exit Sub
    End If
    Me.MousePointer = vbHourglass
    On Error GoTo errH
    strSQL = "Select Distinct s.������ From All_Tables t, Zlsystems s Where t.Table_Name = '���ű�' And t.Owner = s.������"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, App.Title)
    ReDim str������(0 To rsTmp.RecordCount)
    Do While Not rsTmp.EOF
        str������(lngCount) = rsTmp!������
        lngCount = lngCount + 1
        rsTmp.MoveNext
    Loop
    On Error Resume Next
    '���ȴ��������ڵĽ�ɫ
    strSQL = "Select Distinct r.��ɫ" & vbNewLine & _
            "From Zlsystems s, Zlrolegrant r" & vbNewLine & _
            "Where s.��� = r.ϵͳ And s.������ = User And r.��ɫ Not In (Select Granted_Role From User_Role_Privs)"

    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, App.Title)
    If err.Number <> 0 Then err.Clear
    If Not rsTmp Is Nothing Then
        Do While Not rsTmp.EOF
            '���������ڵĽ�ɫ
            gcnOracle.Execute "Create Role " & rsTmp!��ɫ & " Not Identified"
            If err.Number = 0 Then
                '�����ɹ����������ӵ��б���
                Set lst = lvwRole.ListItems.Add(, rsTmp!��ɫ & "", Mid(rsTmp!��ɫ & "", 4), "Role", "Role")
                Call GrantSpecialToRole(gcnOracle, rsTmp!��ɫ, False, str������, True)
                '����ɫ��Ϣͬ�����뵽Zlroles����
                gstrSQL = "zltools.Zl_Zlroles_Edit(1,'" & rsTmp!��ɫ & "')"
                ExecuteProcedure gstrSQL, "������ɫ��ϵͳ�Ķ�Ӧ��ϵ"
            Else
                err.Clear
            End If
            rsTmp.MoveNext
        Loop
    End If
    If err.Number <> 0 Then err.Clear
    On Error GoTo errH
    '��ʼ��Ȩ
    Call ReGrantToRole(gcnOracle, "", True, str������)
    '������Ҫ������־
    Call SaveAuditLog(2, "�ָ����н�ɫ��Ȩ��", "�ָ����н�ɫ��Ȩ��")
    '��ʾ��Ȩ�嵥
    If Not lst Is Nothing Then
        Call InitRoleData
        lst.Selected = True
        Call lvwRole_ItemClick(lst)
    End If
    MsgBox "���н�ɫ������Ȩ��ɣ�", vbInformation, gstrSysName
    MousePointer = 0
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MousePointer = 0
    MsgBox "��ǰ�û���Ȩ�޲�����ɱ�������", vbInformation, gstrSysName
End Sub

Private Sub cmdModify_Click()
    frmProgPriv.ProgPriv
End Sub

Private Sub cmdModifyGroup_Click()
    '����
    If tvwGroups.SelectedItem Is Nothing Then Exit Sub
    If tvwGroups.SelectedItem.Key <> "Root" And tvwGroups.SelectedItem.Key <> "unGroup" Then
        tvwGroups.SelectedItem.Text = Mid(tvwGroups.SelectedItem.Text, 1, Len(tvwGroups.SelectedItem.Text) - Len(tvwGroups.SelectedItem.Tag))
        tvwGroups.Tag = 1
        Call tvwGroups.StartLabelEdit
    End If
    tvwGroups.Tag = 0
    tvwGroups.SelectedItem.Text = tvwGroups.SelectedItem.Text & tvwGroups.SelectedItem.Tag
End Sub

Private Sub cmdNewGroup_Click()
    Dim strGroupsName As String
    Dim lst As ListItem
    Dim objNode As Node
ReDo:
    strGroupsName = frmNameEdit.GetName(name����)
    
    If strGroupsName = "" Then Exit Sub
    If strGroupsName = "δ����" Then
        MsgBox "��δ���顱Ϊ�������,�����ظ����,����", vbInformation, gstrSysName
        Exit Sub
    End If
    If ActualLen(strGroupsName) > 30 Then
        MsgBox "������Ľ�ɫ���Ʋ��ܴ���30���ַ���15������,����!", vbDefaultButton1 + vbInformation, gstrSysName
        GoTo ReDo:
    End If
    strGroupsName = UCase(Trim(strGroupsName))
    
    '���˺�:20070615����
    '���̲���:zlTools.b_Rolegroupmgr.Rolegroup_Add(����_In In ZlRolegroups.����%Type)
    err = 0: On Error GoTo errHand:
    gstrSQL = "zlTools.b_Rolegroupmgr.Rolegroup_Add("
    gstrSQL = gstrSQL & "'" & UCase(strGroupsName) & "')"
    ExecuteProcedure gstrSQL, Me.Caption
        
    Set objNode = tvwGroups.Nodes.Add("Root", 4, "K" & strGroupsName, strGroupsName & "(0)", 1, 1)
    objNode.Tag = "(0)"
    'objNode.Selected = True
    Call LoadMenu(strGroupsName, strGroupsName)
    
    Call FillModule
    Call SetEnable

'     '������
'     '������:��Tvw������һ���µĽ�ɫ����
'     Dim objNode As Node
'     Dim int��� As Integer
'     Dim strTargetGroup As String
'ReDo:
'    Err = 0: On Error Resume Next
'    int��� = int��� + 1
'    strTargetGroup = "�½���:" & int���
'     Set objNode = tvwGroups.Nodes.Add(, "Root", strTargetGroup, strTargetGroup)
'     If Err <> 0 Then
'        Err.Clear: On Error GoTo 0
'        GoTo ReDo
'     End If
'     Err = 0
'    objNode.Tag = "1"
'    objNode.Selected = True
'    tvwGroups.SetFocus
'    tvwGroups.LabelEdit
    Exit Sub
errHand:
        Call ShowErrHand
End Sub
Private Sub ShowErrHand()
    '------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '------------------------------------------------------------------------------------------
    Dim strNote As String, lngErrNum As Long
    If gcnOracle.Errors.Count <> 0 Then
        strNote = gcnOracle.Errors(0).Description
        If InStr(UCase(strNote), "[ZLSOFT]") > 0 Then
            '��־����
            lngErrNum = gcnOracle.Errors(0).NativeError
            MsgBox Split(strNote, "[ZLSOFT]")(1), vbExclamation, App.Title
            Exit Sub
        Else
            MsgBox "ע��:" & vbCrLf & "�����������´���:" & vbCrLf & err.Description, vbExclamation, App.Title
        End If
    Else
        MsgBox "ע��:" & vbCrLf & "�����������´���:" & vbCrLf & err.Description, vbExclamation, App.Title
    End If
End Sub

Private Sub cmdRoleMove_Click()
    '��һ����ɫ��һ��ϵͳ�ƶ�����һ������
    Dim i As Long
    
    '���ݵ�ǰѡ���ɫ����ϵͳ����飬�����ڵ����˵��û�
    If Not lvwRole.SelectedItem Is Nothing Then
        For i = 1 To mnuPopuRoleMoveGroups.UBound
            If lvwRole.SelectedItem.SubItems(LRH_Group) = mnuPopuRoleMoveGroups(i).Tag Then
                mnuPopuRoleMoveGroups(i).Enabled = False
            End If
        Next
    End If
    PopupMenu mnuPopuRoleMove1
    mnuPopuRoleMove1.Visible = True
End Sub

Private Sub cmdRolesReset_Click()
    On Error GoTo errH
    If MsgBox("�����ܽ��������Ʒ��������н�ɫ�������û������ݿ���ʵ��ӵ�еĽ�ɫ���²������н�ɫ���ݡ�" & vbCrLf & _
                "���û���Ӧ��ϵͳ�еĽ�ɫ�����ݿ���ʵ�ʵĽ�ɫ��һ��ʱ��ִ�д˲�����������һ�µ����ݡ�" & vbCrLf & _
                "��ȷ��Ҫ������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    Call ExecuteProcedure("Zltools.Zl_Checkrolesdiff()", "��鲢����zlroles")
    Call FillRoleData(tvwGroups.SelectedItem.Key, True)
    '������Ҫ������־
    Call SaveAuditLog(2, "�������н�ɫ", "�������н�ɫ��������ɡ�")
    MsgBox "�������н�ɫ��������ɡ�", vbInformation, gstrSysName
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub cmdSystemMove_Click()
    '��һ��ϵͳ�µĽ�ɫ�ƶ�����һ��ϵͳ��
    Dim i As Long
    
    '���ݵ�ǰѡ���ɫ����ϵͳ����飬�����ڵ����˵��û�
    If Not lvwRole.SelectedItem Is Nothing Then
        For i = 1 To mnuPopuRoleMoveSystems.UBound
            If lvwRole.SelectedItem.SubItems(LRH_System) = mnuPopuRoleMoveSystems(i).Tag Then
                mnuPopuRoleMoveSystems(i).Enabled = False
            End If
        Next
    End If
    PopupMenu mnuPopuRoleMove2
    mnuPopuRoleMove2.Visible = True
End Sub

Private Sub cmdUser_Click()
    If lvwRole.SelectedItem Is Nothing Then Exit Sub
    Call frmRoleUser.ShowEdit(lvwRole.SelectedItem.Text)
End Sub


Private Sub Form_Activate()
    Dim lngTop As Long
    mblnMoveTop = False
    If mblnFirst = False Then Exit Sub
    '���Ի�����
    lngTop = Val(GetSetting("ZLSOFT", "����ģ��\������������\��ɫ����", "PicHLine_TOP", "4500"))
    '������û�ԭ���ù�������ܣ���ôע�����һ���ܲ鵽���ݣ������4500С��������ʾ�ͻ���ѿ�
    If lngTop < 4500 Then lngTop = 4500
    picHLine.Top = lngTop
    mblnMoveTop = True
    Call Form_Resize
    mblnMoveTop = False
    mblnFirst = False
End Sub

Private Sub Form_Load()
   Dim rsTemp As New ADODB.Recordset
   Dim lngTop As Long
   
    '�жϸ��û��ܷ񴴽���ɫ
    gstrSQL = _
        " Select 1 From User_Sys_Privs Where Privilege='CREATE ROLE'" & _
        " Union" & _
        " Select 1 From Role_Sys_Privs Where Privilege='CREATE ROLE'"
    Call OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    cmdAdd.Enabled = rsTemp.RecordCount > 0
    
    If glngSysNo <> -1 Then
        cmdRolesReset.Visible = False
        cmdGrantAll.Visible = False
        cmdSystemMove.Visible = False
    End If
    Call Getע����
    Call FillRollGroup
    Call FillSystem
    mblnFirst = True
End Sub

Private Sub cmbSystem_Click()
    cmbSystem.Tag = cmbSystem.ItemData(cmbSystem.ListIndex)
    If cmbSystem.ItemData(cmbSystem.ListIndex) = -1 Then
        chkOnlyShowNOSystem.Visible = True
    Else
        chkOnlyShowNOSystem.Visible = False
    End If
    lvwRole.Tag = 1
    Call tvwGroups_NodeClick(tvwGroups.SelectedItem)
    lvwRole.Tag = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '���Ի�����
    SaveSetting "ZLSOFT", "����ģ��\������������\��ɫ����", "PicHLine_TOP", picHLine.Top
    If Not mfrmGrant Is Nothing Then
        Set mfrmGrant = Nothing
    End If
    Set mrsRole = Nothing
    Set mobjTip = Nothing
End Sub

Private Sub lvwModule_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwModule.SortKey = ColumnHeader.Index - 1
    lvwModule.SortOrder = Abs(Not lvwModule.SortOrder = 1)
    lvwModule.Sorted = True
End Sub

Private Sub lvwRole_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call FillModule
    Call SetEnable
End Sub

Private Sub Form_Resize()
    Dim sngTemp As Single
    On Error Resume Next
    
    Me.lvwRole.Height = Me.picHLine.Top - Me.lvwRole.Top
    Me.tvwGroups.Height = Me.lvwRole.Height - cmdNewGroup.Height - 20
    Me.lblSystem.Left = lvwRole.Left
    lblModule.Top = picHLine.Top + picHLine.Height + 50
    
    With Me.lvwModule
        .Top = lblModule.Top + lblModule.Height + 50 ' cmdModify.Top + cmdModify.Height + 50
        If Me.ScaleHeight - .Top < 0 Then
             .Height = 0
        Else
            .Height = Me.ScaleHeight - .Top - 50
        End If
        .Width = ScaleWidth - 50 - .Left
        .ColumnHeaders(5).Width = .Width - .ColumnHeaders(1).Width - _
                                .ColumnHeaders(2).Width - .ColumnHeaders(3).Width - _
                                .ColumnHeaders(4).Width - .ColumnHeaders(6).Width
    End With
    
    With cmdAdd
        .Left = ScaleWidth - .Width - 50
    End With
    
    cmdGrant.Left = cmdAdd.Left
    cmdCopy.Left = cmdAdd.Left
    cmdDel.Left = cmdAdd.Left
    cmdRoleMove.Left = cmdAdd.Left
    cmdModify.Left = cmdAdd.Left
    cmdSystemMove.Left = cmdAdd.Left
    chkOnlyShowNOSystem.Left = cmdAdd.Left
    chkOnlyShowNOSystem.Top = txtSearch.Top
    
    With cmdUser
        .Top = lvwRole.Top + lvwRole.Height - .Height
        .Left = cmdAdd.Left
    End With
    
    With cmdModify
        .Top = cmdUser.Top - .Height
        .Left = cmdAdd.Left
    End With
    
    With cmdGrantAll
        .Top = cmdModify.Top - .Height
        .Left = cmdAdd.Left
    End With
    
    With cmdRolesReset
        .Top = cmdGrantAll.Top - .Height
        .Left = cmdAdd.Left
    End With
    
    With lvwRole
        If cmdAdd.Left - 50 - .Left < 0 Then
            .Width = 0
        Else
            .Width = cmdAdd.Left - 50 - .Left
        End If
    End With
    
    With txtSearch
        .Left = lvwRole.Width + lvwRole.Left - .Width
    End With
    
    With lblSearch
        .Left = txtSearch.Left - 50 - .Width
    End With
    
    cmdNewGroup.Top = tvwGroups.Top + tvwGroups.Height + 20
    cmdModifyGroup.Top = cmdNewGroup.Top
    cmdDeleteGroup.Top = cmdNewGroup.Top
    Me.picHLine.Left = 0: Me.picHLine.Width = Me.ScaleWidth
    msngPreHeigt = Me.ScaleHeight - picHLine.Top
End Sub

'Private Sub FillRole()
'    Dim rsTemp As New ADODB.Recordset
'
'    rsTemp.CursorLocation = adUseClient
'
'    '�жϸ��û��ܷ񴴽���ɫ
'    gstrSQL = "Select 1 from User_Sys_privs where privilege='CREATE ROLE' " & _
'        "union Select 1 from Role_Sys_privs where privilege='CREATE ROLE'"
'
'    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
'    cmdAdd.Enabled = rsTemp.RecordCount > 0
'    cmdDelete.Enabled = cmdAdd.Enabled
'    rsTemp.Close
'
'
'
'    '��ʾ���Խ�����Ȩ�Ľ�ɫ
'    If gblnDBA = True Then
'        gstrSQL = "select * from DBA_Roles where Upper(Role) Like 'ZL_%'"
'    Else
'        gstrSQL = "select Granted_Role as Role from user_Role_privs " & _
'            "where Granted_Role Like 'ZL_%'" 'ADMIN_OPTION='YES'ѡ����Բ���
'    End If
'    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
'    lvwRole.ListItems.Clear
'    Do Until rsTemp.EOF
'        lvwRole.ListItems.Add , rsTemp("Role"), Mid(rsTemp("Role"), 4), "Role", "Role"
'        rsTemp.MoveNext
'    Loop
'    If lvwRole.ListItems.Count > 0 Then
'        lvwRole.ListItems(1).Selected = True
'    Else
'        cmdGrant.Enabled = False
'    End If
'    rsTemp.Close
'    Call SetEnable
'End Sub

Private Sub FillSystem()
    Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrHandle
    
    '��Ϊ��ϵͳ��¼������ʾϵͳ
    If glngSysNo <> -1 Then
        lblSystem.Visible = False
        cmbSystem.Visible = False
        cmbSystem.addItem "��ϵͳ��¼"
        cmbSystem.ItemData(cmbSystem.NewIndex) = glngSysNo
        cmbSystem.ListIndex = 0
        chkOnlyShowNOSystem.Visible = False
    Else
        '��ʾ�������е�ϵͳ
        Set rsTemp = New ADODB.Recordset
        rsTemp.CursorLocation = adUseClient
        Set rsTemp = zlGetRegSystems
        cmbSystem.Clear
        cmbSystem.addItem "����ϵͳ"
        cmbSystem.ItemData(cmbSystem.NewIndex) = -1
        Call LoadMenu("����ϵͳ", "-1", False)
        Do Until rsTemp.EOF
            cmbSystem.addItem RPAD(rsTemp("����") & "��" & rsTemp("���") & "��", 25) & " v" & rsTemp("�汾��")
            cmbSystem.ItemData(cmbSystem.NewIndex) = rsTemp("���")
            Call LoadMenu(rsTemp("����"), rsTemp("���"), False)
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        If cmbSystem.ListIndex < 0 Then cmbSystem.ListIndex = 0
    End If
    Exit Sub
ErrHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub FillModule()
    Dim rsTemp As ADODB.Recordset
    Dim lst As ListItem
    Dim strRole As String
    Dim strSystem As String
    
    If cmbSystem.ListCount = 0 And glngSysNo = -1 Then Exit Sub
    
    '��ѡ���ϵͳΪ����ϵͳ,����ʾ�ý�ɫ������ϵͳ�µ�Ȩ��
    '��ѡ���ϵͳΪ�����ϵͳ�������ʾ�ý�ɫ�ڸ�ϵͳ�µ�Ȩ��
    LockWindowUpdate lvwModule.hwnd
    
    lvwModule.ColumnHeaders.Clear
    lvwModule.ListItems.Clear
    If Not lvwRole.SelectedItem Is Nothing Then
        strRole = lvwRole.SelectedItem.Key
    End If
    '�����б���
    With lvwModule.ColumnHeaders
        .Add , , "ϵͳ", "2000"
        .Add , , "���", "800"
        .Add , , "���ܻ����", "1500"
        .Add , , "˵��", "5000"
        .Add , , "��Ȩ����", "1500"
        .Add , , "ϵͳ��", "0"
    End With
    
    If strRole = "" Then
        '��ɫΪ�գ��˳�
        LockWindowUpdate 0
        Exit Sub
    End If
    
    '��ʾ�ý�ɫ�ܷ��ʵĻ�����
    gstrSQL = "Select t.ϵͳ, t.���, t.����, t.˵��" & vbNewLine & _
            "From (Select s.���� || '��' || s.��� || '��' As ϵͳ, s.���, s.������, b.����, b.˵��" & vbNewLine & _
            "       From Zlsystems s, Zlbasecode b" & vbNewLine & _
            "       Where b.ϵͳ = s.���" & IIf(cmbSystem.ItemData(cmbSystem.ListIndex) = -1, "", " And s.��� = [2]") & ") t, User_Tab_Privs r" & vbNewLine & _
            "Where t.������ = r.Owner And t.���� = r.Table_Name And r.Grantee = [1] And" & vbNewLine & _
            "      r.Privilege In ('SELECT', 'INSERT', 'UPDATE', 'DELETE')" & vbNewLine & _
            "Group By t.ϵͳ, t.���, t.����, t.˵��" & vbNewLine & _
            "Having Count(r.Privilege) = 4"

    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, "��ȡ����������Ϣ", strRole, cmbSystem.ItemData(cmbSystem.ListIndex))
    Do Until rsTemp.EOF
        Set lst = lvwModule.ListItems.Add(, , rsTemp("ϵͳ"))
        lst.SubItems(LH_���ܻ����) = rsTemp("����")
        lst.SubItems(LH_˵��) = IIf(IsNull(rsTemp("˵��")), "", rsTemp("˵��"))
        lst.SubItems(LH_ϵͳ��) = Lpad(rsTemp("���"), 4)
        rsTemp.MoveNext
    Loop
    '��ʾ�ý�ɫ�ܷ��ʵĻ�����
    gstrSQL = "Select s.���� || '��' || s.��� || '��' As ϵͳ, s.���, s.������, f.������, f.������, f.˵��" & vbNewLine & _
            "From Zlsystems s, Zlfunctions f, User_Tab_Privs r" & vbNewLine & _
            "Where f.ϵͳ = s.��� And s.������ = r.Owner And Upper(f.������) = r.Table_Name And r.Grantee = [1] And" & vbNewLine & _
            "      r.Privilege = 'EXECUTE'" & IIf(cmbSystem.ItemData(cmbSystem.ListIndex) = -1, "", " And s.��� = [2]")

    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, "��ȡȡ��������Ϣ", strRole, cmbSystem.ItemData(cmbSystem.ListIndex))
    Do Until rsTemp.EOF
        Set lst = lvwModule.ListItems.Add(, , rsTemp("ϵͳ"))
        lst.SubItems(LH_���ܻ����) = rsTemp("������") & "(" & rsTemp("������") & ")"
        lst.SubItems(LH_˵��) = IIf(IsNull(rsTemp("˵��")), "", rsTemp("˵��"))
        lst.SubItems(LH_ϵͳ��) = Lpad(rsTemp("���"), 4)
        rsTemp.MoveNext
    Loop
    '��ʾ�ý�ɫ�ܷ��ʵĻ�������
    gstrSQL = "select P.���,P.����,P.˵��,R.���� from " & _
            "zlRoleGrant R,zlPrograms P " & _
            "where R.ϵͳ is Null And P.���=R.��� And R.��ɫ=[1]" & _
            " And P.ϵͳ is Null And P.���<100 And P.���� is Null " & _
            " Order By P.���"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, "��ȡ����������Ϣ", strRole)
    On Error Resume Next
    Do Until rsTemp.EOF
        Set lst = lvwModule.ListItems.Add(, "C" & rsTemp("���"), "��������")
        lst.SubItems(LH_ϵͳ��) = "9999"
        If IsNull(rsTemp("���")) Then
            If rsTemp("����") <> "����" Then
                Set lst = lvwModule.ListItems("C" & rsTemp("���"))
                lst.SubItems(LH_��Ȩ����) = IIf(lst.SubItems(LH_��Ȩ����) = "", "", lst.SubItems(LH_��Ȩ����) & ",") & rsTemp("����")
            End If
        Else
            lst.SubItems(LH_���) = Lpad(rsTemp("���"), 6)
            lst.SubItems(LH_���ܻ����) = rsTemp("����")
            lst.SubItems(LH_˵��) = IIf(IsNull(rsTemp("˵��")), "", rsTemp("˵��"))
            If rsTemp("����") <> "����" Then
                lst.SubItems(LH_��Ȩ����) = rsTemp("����")
            End If
        End If
        rsTemp.MoveNext
    Loop
        
    '��ʾ�ý�ɫ�ܷ��ʵ�ģ���Լ��Զ��屨��
    gstrSQL = "select S.����||'��'||S.���||'��' ϵͳ,S.���,P.���,P.����,P.˵��,R.���� from " & _
            "zlRoleGrant R,zlPrograms P ,zlSystems S " & _
            "where nvl(R.ϵͳ,0)=nvl(P.ϵͳ,0) And P.���=R.��� And P.���>=100 And R.��ɫ=[1] And P.ϵͳ = S.���(+)" & _
            IIf(cmbSystem.ItemData(cmbSystem.ListIndex) = -1, "", " And Nvl(P.ϵͳ, [2]) = [2]") & _
            " Order By P.���"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, "��ȡ����ϵͳ����Ϣ", strRole, cmbSystem.ItemData(cmbSystem.ListIndex))
    On Error Resume Next
    Do Until rsTemp.EOF
        strSystem = rsTemp("ϵͳ")
        If rsTemp("ϵͳ") = "����" Then
            Set lst = lvwModule.ListItems.Add(, "C" & rsTemp("���"), "�Զ��屨��")
            lst.SubItems(LH_ϵͳ��) = "8888"
        Else
            Set lst = lvwModule.ListItems.Add(, "C" & rsTemp("���"), rsTemp("ϵͳ"))
            lst.SubItems(LH_ϵͳ��) = Lpad(rsTemp("���"), 4)
        End If
        If err <> 0 Then
            err.Clear
            If rsTemp("����") <> "����" Then
                Set lst = lvwModule.ListItems("C" & rsTemp("���"))
                lst.SubItems(LH_��Ȩ����) = IIf(lst.SubItems(LH_��Ȩ����) = "", "", lst.SubItems(LH_��Ȩ����) & ",") & rsTemp("����")
            End If
        Else
            lst.SubItems(LH_���) = Lpad(rsTemp("���"), 6)
            lst.SubItems(LH_���ܻ����) = rsTemp("����")
            lst.SubItems(LH_˵��) = IIf(IsNull(rsTemp("˵��")), "", rsTemp("˵��"))
            If rsTemp("����") <> "����" Then
                lst.SubItems(LH_��Ȩ����) = rsTemp("����")
            End If
        End If
        rsTemp.MoveNext
    Loop
    
    lvwModule.SortKey = LH_ϵͳ��
    lvwModule.SortOrder = 0
    lvwModule.Sorted = True
    
    LockWindowUpdate 0
End Sub

Private Sub SetEnable()
'���ø�����ť��Enable����
    Dim blnHave As Boolean
    Dim i As Long
    Dim lstItem As ListItem
    blnHave = Not lvwRole.SelectedItem Is Nothing
    mnuPopuModify.Enabled = tvwGroups.SelectedItem.Key <> "Root" And tvwGroups.SelectedItem.Key <> "unGroup"
    mnuPopuDeleteGroups.Enabled = mnuPopuModify.Enabled
    cmdModifyGroup.Enabled = mnuPopuModify.Enabled
    cmdDeleteGroup.Enabled = mnuPopuModify.Enabled
    'cmdDelete.Enabled = cmdAdd.Enabled And blnHave
    cmdGrant.Enabled = blnHave
    cmdUser.Enabled = blnHave
    cmdCopy.Enabled = blnHave
    cmdDel.Enabled = blnHave
    cmdRoleMove.Enabled = blnHave
    cmdGrantAll.Enabled = (gblnOwner = True)
    mnuPopuRoleDelete.Enabled = blnHave
    blnHave = False
    For Each lstItem In lvwRole.ListItems
        If lstItem.Selected = True Then
            blnHave = True
            Exit For
        End If
    Next
    For i = 1 To mnuPopuRoleMoveGroups.UBound
        mnuPopuRoleMoveGroups(i).Enabled = blnHave
        If UCase(mnuPopuRoleMoveGroups(i).Tag) = UCase(tvwGroups.SelectedItem.Key) Then
            mnuPopuRoleMoveGroups(i).Enabled = False
        End If
    Next
    For i = 1 To mnuPopuRoleMoveSystems.UBound
        mnuPopuRoleMoveSystems(i).Enabled = blnHave
        If mnuPopuRoleMoveSystems(i).Tag = cmbSystem.ItemData(cmbSystem.ListIndex) Then
            mnuPopuRoleMoveSystems(i).Enabled = False
        End If
    Next
    If glngSysNo <> -1 Then
        mnuPopuRoleMove2.Visible = False
    End If
End Sub


Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = True
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�

'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objPrint As zlPrintLvw
    Dim rsTmp As ADODB.Recordset
    
    If lvwRole.SelectedItem Is Nothing Then Exit Sub
    Set objPrint = New zlPrintLvw
    objPrint.Title.Text = "��ɫȨ��"
    Set objPrint.Body.objData = lvwModule
    objPrint.UnderAppItems.Add "��ɫ��" & lvwRole.SelectedItem.Text
    If glngSysNo <> -1 Then
        gstrSQL = "Select ���� From Zlsystems Where ��� = [1]"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, "��ѯϵͳ����", cmbSystem.ItemData(cmbSystem.ListIndex))
        objPrint.UnderAppItems.Add "��Ȩϵͳ��" & rsTmp!����
    Else
        objPrint.UnderAppItems.Add "��Ȩϵͳ��" & cmbSystem.Text
    End If
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Format(CurrentDate, "yyyy��MM��dd��")
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewLvw objPrint, 1
          Case 2
              zlPrintOrViewLvw objPrint, 2
          Case 3
              zlPrintOrViewLvw objPrint, 3
      End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If

End Sub

Private Sub lvwRole_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDelete
        Call mnuPopuRoleDelete_Click
    End Select
End Sub

Private Sub lvwRole_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objItem As ListItem
    Dim strTip As String, strTitle As String
    
    If Button = 1 Then
        '����ͼ��
        If lvwRole.SelectedItem Is Nothing Then Exit Sub
        Set lvwRole.DragIcon = lvwRole.SelectedItem.CreateDragImage
        lvwRole.Drag 1
    End If
    
    'ֻ�е��ǵ�ϵͳ��¼���ҵ�ǰѡ��ϵͳΪ����ϵͳʱ���ŵ�����ʾ��
    If cmbSystem.ItemData(cmbSystem.ListIndex) = -1 And glngSysNo = -1 And Button = 0 Then
        If mobjTip Is Nothing Then
            Call InitTips
        End If
        Set objItem = lvwRole.HitTest(x, y)
        If Not objItem Is Nothing Then
            If objItem.SubItems(LRH_SystemName) <> "����ϵͳ" Then
                strTip = objItem.SubItems(LRH_SystemName)
                strTitle = "����ϵͳ"
            Else
                strTip = ""
                strTitle = ""
            End If
            mobjTip.TipText = strTip
            mobjTip.Title = strTitle
        Else
            mobjTip.TipText = ""
            mobjTip.Title = ""
        End If
    End If
End Sub

Private Sub InitTips()
    Set mobjTip = New clsTipSwap
    Set mobjTip.ParentControl = lvwRole
    mobjTip.Icon = TTIconInfo
    mobjTip.Style = TTBalloon
    mobjTip.Create
End Sub

Private Sub lvwRole_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long

    If Button = 1 Then Exit Sub
    If Not lvwRole.SelectedItem Is Nothing Then
        For i = 1 To mnuPopuRoleMoveGroups.UBound
            mnuPopuRoleMoveGroups(i).Enabled = True
            If lvwRole.SelectedItem.SubItems(LRH_Group) = mnuPopuRoleMoveGroups(i).Tag Then
                mnuPopuRoleMoveGroups(i).Enabled = False
            End If
        Next
        For i = 1 To mnuPopuRoleMoveSystems.UBound
            mnuPopuRoleMoveSystems(i).Enabled = True
            If lvwRole.SelectedItem.SubItems(LRH_System) = mnuPopuRoleMoveSystems(i).Tag Then
                mnuPopuRoleMoveSystems(i).Enabled = False
            End If
        Next
    End If
    PopupMenu mnuPopuRole
End Sub

Private Sub mnuPopuAddGroups_Click()
    Call cmdNewGroup_Click
End Sub

Private Sub mnuPopuDeleteGroups_Click()
    Call cmdDeleteGroup_Click
End Sub

Private Sub mnuPopuModify_Click()
    Call cmdModifyGroup_Click
End Sub

Private Sub mnuPopuRoleAdd_Click()
    Call cmdAdd_Click
End Sub

Private Sub mnuPopuRoleDelete_Click()
    Call cmdDel_Click
End Sub

Private Sub mnuPopuRoleMoveGroups_Click(Index As Integer)
    Dim strTargetGroup As String
    If mnuPopuRoleMoveGroups(Index).Tag = "" Then Exit Sub
    strTargetGroup = Mid(mnuPopuRoleMoveGroups(Index).Tag, 2)
    
    If strTargetGroup = "NGROUP" Then
        If strTargetGroup = Mid(tvwGroups.SelectedItem.Key, 2) Then Exit Sub
        If MsgBox("����Ҫ����ɫ��" & lvwRole.SelectedItem.Text & "...�� �Ƴ�������?", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) <> vbYes Then Exit Sub
        strTargetGroup = ""
    Else
        If strTargetGroup = Mid(tvwGroups.SelectedItem.Key, 2) Then Exit Sub
        If MsgBox("����Ҫ����ɫ��" & lvwRole.SelectedItem.Text & "...�� �ƶ����顰" & strTargetGroup & "������?", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) <> vbYes Then Exit Sub
    End If
    If MoveToGroups(strTargetGroup, True) = False Then Exit Sub
    Call SetEnable
    
End Sub

Private Sub mnuPopuRoleMoveSystems_Click(Index As Integer)
    Dim strSystemsNo As String
    
    If mnuPopuRoleMoveSystems(Index).Tag = "" Then Exit Sub
    mstrSystemsName = mnuPopuRoleMoveSystems(Index).Caption
    strSystemsNo = mnuPopuRoleMoveSystems(Index).Tag
    
    If mstrSystemsName = "����ϵͳ" Then
        If MsgBox("����Ҫ����ɫ��" & lvwRole.SelectedItem.Text & "...�� �Ƴ���ϵͳ��?", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) <> vbYes Then Exit Sub
        mstrSystemsName = ""
    Else
        If MsgBox("����Ҫ����ɫ��" & lvwRole.SelectedItem.Text & "...�� �ƶ���ϵͳ��" & mstrSystemsName & "������?", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) <> vbYes Then Exit Sub
    End If
    If MoveToGroups(strSystemsNo, False) = False Then Exit Sub
    Call SetEnable
End Sub

Private Sub picHLine_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Me.picHLine.BackColor = &H8000000F: Me.picHLine.Top = Me.picHLine.Top + y
End Sub

Private Sub picHLine_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.picHLine.BackColor = Me.BackColor
    If Me.picHLine.Top < 4500 Then Me.picHLine.Top = 4500
    If Me.picHLine.Top > Me.ScaleHeight - 1500 Then Me.picHLine.Top = Me.ScaleHeight - 1500
    mblnMoveTop = True
    Call Form_Resize
    mblnMoveTop = False
End Sub
 
Private Sub tvwGroups_AfterLabelEdit(Cancel As Integer, NewString As String)
    err = 0: On Error GoTo errHand:
    Dim strKey As String
    Dim strTag As String
    
    strKey = UCase(Mid(tvwGroups.SelectedItem.Key, 2))
    If strKey = NewString Then
        NewString = UCase(NewString) & tvwGroups.SelectedItem.Tag
        Exit Sub
    End If
    If NewString = "δ����" Then
        MsgBox "����Ϊ:δ�����Ѿ�����,�������Ӵ���,����", vbExclamation, "������������"
        NewString = tvwGroups.SelectedItem.Text
    Else
        '����:zlTools.b_Rolegroupmgr.Rolegroup_Delete(
        '    ����_Old_In In ZlRolegroups.����%Type,
        '    ����_New_In In ZlRolegroups.����%Type
        gstrSQL = "zlTools.b_Rolegroupmgr.Rolegroup_Rename("
        gstrSQL = gstrSQL & "'" & strKey & "',"
        gstrSQL = gstrSQL & "'" & UCase(NewString) & "')"
        ExecuteProcedure gstrSQL, Me.Caption
        tvwGroups.SelectedItem.Key = "K" & NewString
        NewString = UCase(NewString) & tvwGroups.SelectedItem.Tag
        strTag = tvwGroups.SelectedItem.Tag
        Call InitRoleData
        tvwGroups.SelectedItem.Tag = strTag
    End If
    Exit Sub
errHand:
    Cancel = True
    Call ShowErrHand
End Sub
Private Sub tvwGroups_BeforeLabelEdit(Cancel As Integer)
    'tvwGroups.Tag = 0 ��ʾ����������ť��˫���ſ��޸��������Ǵ��෽���������޸�
    If Me.tvwGroups.SelectedItem.Key = "Root" Or Me.tvwGroups.SelectedItem.Key = "unGroup" Or Val(tvwGroups.Tag) = 0 Then
        Cancel = True
    End If
End Sub
Private Sub DeleteRoleGroups()
    '---------------------------------------------------------------------------------------------------------
    '����:ɾ����
    '����:���˺�
    '����:2007/06/15
    '---------------------------------------------------------------------------------------------------------
    Dim strRoleGroupName As String
    Dim intIndex As Integer
    Dim rsTmp As ADODB.Recordset
    
    strRoleGroupName = Mid(tvwGroups.SelectedItem.Text, 1, Len(tvwGroups.SelectedItem.Text) - Len(tvwGroups.SelectedItem.Tag))
    intIndex = tvwGroups.SelectedItem.Index
    
    If MsgBox("���Ҫɾ����" & strRoleGroupName & "���Ľ�ɫ����", vbDefaultButton2 Or vbQuestion Or vbYesNo, gstrSysName) = vbNo Then Exit Sub
    
    '����ǰѡ���ϵͳΪ������ϵͳ���������жϡ�
    '�����ж�����ϵͳ�и÷������Ƿ��н�ɫ�����У�����ɾ��
    If cmbSystem.ItemData(cmbSystem.ListIndex) <> "-1" Then
        '���жϵ�ǰϵͳ�и÷������Ƿ��н�ɫ�����ȼ��lvwRole����û����Ŀ
        If lvwRole.ListItems.Count > 0 Then
            MsgBox "�÷����»��н�ɫ���ʲ���ɾ���÷��顣" & vbNewLine & "��һ��Ҫɾ���÷��飬�ɽ�ϵͳ�л�Ϊ������ϵͳ�����ٽ���ɾ�����������", vbInformation, gstrSysName
            Exit Sub
        Else
            gstrSQL = "Select Count(1) ���� From Zlrolegroups b Where ���� = [1] And ��ɫ Is Not Null"
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, "��������ϵͳ�и÷����µĽ�ɫ", strRoleGroupName)
            If rsTmp!���� > 0 Then
                MsgBox "�÷����»��н�ɫ���ʲ���ɾ���÷��顣" & vbNewLine & "��һ��Ҫɾ���÷��飬�ɽ�ϵͳ�л�Ϊ������ϵͳ�����ٽ���ɾ�����������", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    
    err = 0: On Error GoTo errHand:
    '����:zlTools.b_Rolegroupmgr.Rolegroup_Delete(����_In In ZlRolegroups.����%Type)
    gstrSQL = "zlTools.b_Rolegroupmgr.Rolegroup_Delete("
    gstrSQL = gstrSQL & "'" & UCase(strRoleGroupName) & "')"
    ExecuteProcedure gstrSQL, Me.Caption
    
    tvwGroups.Nodes.Remove intIndex
    If tvwGroups.Nodes.Count > 0 Then
        If intIndex > tvwGroups.Nodes.Count Then intIndex = tvwGroups.Nodes.Count
        tvwGroups.Nodes(intIndex).Selected = True
        tvwGroups.Nodes(intIndex).EnsureVisible
    End If
    Call FillRoleData(tvwGroups.SelectedItem.Key, IIf(cmbSystem.ItemData(cmbSystem.ListIndex) = -1, True, False))
    Exit Sub
errHand:
    Call ShowErrHand
End Sub

Private Sub tvwGroups_DblClick()
    Call cmdModifyGroup_Click
End Sub

Private Sub tvwGroups_DragDrop(Source As Control, x As Single, y As Single)
    Dim strTargetGroup As String, str��ɫ As String, intIndex As Integer
    Dim lstItem As ListItem
    Dim strKeys As String
    Dim arrVar As Variant
    Dim i As Long
    
    err = 0: On Error GoTo errHand:
    If Source Is lvwRole And Not tvwGroups.DropHighlight Is Nothing Then
        intIndex = -1
        strTargetGroup = Mid(tvwGroups.DropHighlight.Key, 2)
        Set tvwGroups.DropHighlight = Nothing
        tvwGroups.DropHighlight = tvwGroups.SelectedItem

        If strTargetGroup = "oot" Or strTargetGroup = "nGroup" Then
            If MsgBox("����Ҫ����ɫ��" & Source.SelectedItem.Text & "...�� �Ƴ�������?", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
            strTargetGroup = ""
        Else
            If MsgBox("����Ҫ����ɫ��" & Source.SelectedItem.Text & "...�� �ƶ����顰" & strTargetGroup & "������?", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
        End If

        gcnOracle.BeginTrans
        strKeys = ""
        For Each lstItem In lvwRole.ListItems
            If lstItem.Selected Then
                If intIndex < 0 Then
                    intIndex = lstItem.Index
                End If
                str��ɫ = lstItem.Key
                strKeys = strKeys & "'" & lstItem.Key

                If MoveToRoleGroup(strTargetGroup, str��ɫ) = False Then
                    gcnOracle.RollbackTrans
                    Exit Sub
                End If
            End If
        Next
        gcnOracle.CommitTrans
        If strKeys <> "" Then strKeys = Mid(strKeys, 2)
        '����ǰ�ڵ��������н�ɫ�У����������
        Call InitRoleData
        If tvwGroups.SelectedItem.Key <> "Root" Then
            arrVar = Split(strKeys, "'")
            For i = 0 To UBound(arrVar)
                lvwRole.ListItems.Remove arrVar(i)
            Next

            If lvwRole.ListItems.Count > 0 Then
                If intIndex > lvwRole.ListItems.Count Then intIndex = lvwRole.ListItems.Count
                lvwRole.ListItems(intIndex).Selected = True
            End If
            Call FillModule
        Else
            lvwRole.SelectedItem.SubItems(LRH_Group) = IIf(strTargetGroup = "", "UNGROUP", "K" & strTargetGroup)
        End If
    End If
    Call SetEnable
    tvwGroups.Refresh
     
    Set tvwGroups.DropHighlight = Nothing
    Exit Sub
errHand:
    Set tvwGroups.DropHighlight = Nothing
    Call ShowErrHand
End Sub
Private Function MoveToGroups(ByVal strTargetGroup As String, Optional ByVal blnType As Boolean = True) As Boolean
    '-------------------------------------------------------------------------------------------------------------------
    '����:��ָ����ɫ�ƶ�������
    '����:
    '     strTargetGroup-�Ƶ�Ŀ����������������ƶ���ϵͳ����Ϊϵͳ�ţ������ƶ������飬��Ϊ����
    '     blnType-��������
    '         blnType = True  ����ɫ����
    '         blnType = False ��ϵͳ����
    '�ƶ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/06/19
    '-------------------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    Dim strKeys  As String
    Dim lstItem As ListItem
    Dim str��ɫ As String
    Dim arrVar As Variant
    Dim i As Long
    MoveToGroups = False
    intIndex = -1
    gcnOracle.BeginTrans
    strKeys = ""
    For Each lstItem In lvwRole.ListItems
        If lstItem.Selected Then
            If intIndex < 0 Then
                intIndex = lstItem.Index
            End If
            str��ɫ = lstItem.Key
            strKeys = strKeys & "'" & lstItem.Key
            If blnType = True Then
                If MoveToRoleGroup(strTargetGroup, str��ɫ) = False Then
                    gcnOracle.RollbackTrans
                    Exit Function
                End If
            Else
                If MoveToSystemGroup(strTargetGroup, str��ɫ) = False Then
                    gcnOracle.RollbackTrans
                    Exit Function
                End If
            End If
        End If
    Next
    gcnOracle.CommitTrans
    If strKeys <> "" Then strKeys = Mid(strKeys, 2)
    '����ǰ�ڵ��������н�ɫ������ϵͳ�У��������������Ҫ�����ƶ��ɹ�����ʾ
    Call InitRoleData
    If (blnType = True And tvwGroups.SelectedItem.Key <> "Root") Or (blnType = False And (cmbSystem.ItemData(cmbSystem.ListIndex) <> -1 Or chkOnlyShowNOSystem.value = 1)) Then
        arrVar = Split(strKeys, "'")
        For i = 0 To UBound(arrVar)
            lvwRole.ListItems.Remove arrVar(i)
        Next
        If lvwRole.ListItems.Count > 0 Then
            If intIndex > lvwRole.ListItems.Count Then intIndex = lvwRole.ListItems.Count
            lvwRole.ListItems(intIndex).Selected = True
        End If
        FillModule
    Else
        If blnType Then
            '����ɫ����
            lvwRole.SelectedItem.SubItems(LRH_Group) = IIf(strTargetGroup = "", "UNGROUP", "K" & strTargetGroup)
        Else
            '��ϵͳ����
            lvwRole.SelectedItem.SubItems(LRH_System) = strTargetGroup
            lvwRole.SelectedItem.SubItems(LRH_SystemName) = mstrSystemsName
            If strTargetGroup = -1 Then
                lvwRole.SelectedItem.Icon = "Role"
            Else
                lvwRole.SelectedItem.Icon = "Role_Moved"
        End If
        End If
    End If
    MoveToGroups = True
End Function
Private Function MoveToRoleGroup(ByVal str�� As String, str��ɫ As String) As Boolean
    '-------------------------------------------------------------------------------------------------------------------
    '����:��ָ���Ľ�ɫ�Ƶ�����
    '����:str��-�Ƶ��������
    '     str��ɫ-ָ���Ľ�ɫ
    '�ƶ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/06/19
    '-------------------------------------------------------------------------------------------------------------------
    err = 0: On Error GoTo errHand:
    
    '�жϸý�ɫ�Ƿ��ڸ������Ѵ���
    If str�� = "" Then
        mrsRole.Filter = "���� = null And Role = '" & str��ɫ & "'"
    Else
        mrsRole.Filter = "���� = '" & str�� & "' And Role = '" & str��ɫ & "'"
    End If
    If mrsRole.RecordCount = 1 Then
        MsgBox "��ɫ��" & str��ɫ & "���Ѵ�����" & vbNewLine & "���顰" & IIf(str�� = "", "δ����", str��) & "���У������ٴ��ƶ���", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���˺�:20070615����
    '���̲���:zlTools.b_Rolegroupmgr.RoletoRolegroup
    '        ����_In In ZlRolegroups.����%Type,
    '        ��ɫ_In In ZlRolegroups.��ɫ%Type := Null
    gstrSQL = "zlTools.b_Rolegroupmgr.RoletoRolegroup("
    gstrSQL = gstrSQL & IIf(str�� = "", "Null", "'" & UCase(str��) & "'") & ","
    gstrSQL = gstrSQL & "'" & UCase(str��ɫ) & "')"
    ExecuteProcedure gstrSQL, Me.Caption
    MoveToRoleGroup = True
    Exit Function
errHand:
    Call ShowErrHand
End Function

Private Function MoveToSystemGroup(ByVal lngGroup As Long, ByVal strRoleName As String) As Boolean
    '���ܣ���ָ����ɫ�ƶ���ϵͳ������
    '������
    '     lngGroup-��Ҫ�Ƶ���ϵͳ�ı�ţ���Ϊ-1����ʾҪ����ɫ�ƶ���������ϵͳ��
    '     strRoleName-��Ҫ�ƶ��Ľ�ɫ������
    
    On Error GoTo errHand:
    
    '�жϸý�ɫ�Ƿ��ڸ�ϵͳ���Ѵ���
    mrsRole.Filter = "ϵͳ = " & lngGroup & " And Role = '" & strRoleName & "'"
    If mrsRole.RecordCount = 1 Then
        MsgBox "��ɫ��" & strRoleName & "���Ѵ�����" & vbNewLine & "��ϵͳ�У������ٴ��ƶ���", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "zltools.Zl_Zlroles_Edit(2,'" & strRoleName & "'" & IIf(lngGroup = -1, "", "," & lngGroup) & ")"
    ExecuteProcedure gstrSQL, "�޸Ľ�ɫ����ϵͳ"
    MoveToSystemGroup = True
    Exit Function
errHand:
    Call ShowErrHand
End Function

Private Sub tvwGroups_DragOver(Source As Control, x As Single, y As Single, State As Integer)
      Dim objOver As Node
      If Source Is lvwRole Then
           Set objOver = tvwGroups.HitTest(x, y)
            If Not objOver Is Nothing Then
                '�жϵ�ǰ��ѡ��ɫ�Ƿ����ڵ�ǰ��ѡ������
                If lvwRole.SelectedItem.SubItems(LRH_Group) = UCase(objOver.Key) Or objOver.Key = "Root" Then
                    Set tvwGroups.DropHighlight = Nothing
                    lvwRole.DragIcon = Nothing
                Else
                    Set tvwGroups.DropHighlight = objOver
                    tvwGroups.DropHighlight.EnsureVisible
                    lvwRole.DragIcon = ils32.ListImages(lvwRole.SelectedItem.Icon).Picture
                    
                End If
            Else
                Set tvwGroups.DropHighlight = Nothing
            End If
      End If
End Sub

Private Sub tvwGroups_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDelete
        mnuPopuDeleteGroups_Click
    End Select
End Sub

Private Sub tvwGroups_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Exit Sub
    PopupMenu mnuPopu
End Sub

Private Sub tvwGroups_NodeClick(ByVal Node As MSComctlLib.Node)
    '---------------------------------------------------------------------------------------------------------
    '��ȡ��Ӧ�Ľ�ɫȨ��
    '---------------------------------------------------------------------------------------------------------
    Call FillRoleData(Node.Key, IIf(Val(lvwRole.Tag) = 1, True, False))
End Sub

Private Function FillRoleData(ByVal strTargetGroup As String, Optional ByVal blnRefreshData As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------
    '����:��������,��ȡ��ɫ��Ϣ,����䵽lvw��
    '����:
    '     strTargetGroup:<>""ָ������,=""��ʾ���н�ɫ
    '     blnRefreshData:����Ƿ���Ҫˢ�½�ɫ����
    '����:���سɹ�,����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------
    Dim rsGroups As ADODB.Recordset
    Dim objItem As ListItem
    Dim strSearch  As String, strFilter As String
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    
    strSearch = UCase(Trim(txtSearch.Text))
    If strTargetGroup <> "Root" And strTargetGroup <> "unGroup" Then
        strTargetGroup = Mid(strTargetGroup, 2)
    End If
     
    '���鿴����ʱ����������Դ����ˢ��
    '�������������޸ĺ󣬲�����ˢ������Դ
    If blnRefreshData Then Call InitRoleData
    If mrsRole Is Nothing Then Call InitRoleData
    
    If strTargetGroup = "Root" Then
        strFilter = ""
    ElseIf strTargetGroup = "unGroup" Then
        strFilter = "���� = null"
    Else
        strFilter = "���� = '" & strTargetGroup & "'"
    End If
    
    '�ж��Ƿ����ʾδ��ϵͳ�Ľ�ɫ
    If chkOnlyShowNOSystem.Visible = True And chkOnlyShowNOSystem.value = 1 Then
        If strFilter = "" Then
            strFilter = "ϵͳ = null"
        Else
            strFilter = strFilter & " And ϵͳ = null"
        End If
    End If
    
    If strSearch <> "" Then
        If strFilter = "" Then
            strFilter = "RoleName Like '" & strSearch & "%' or ���� Like '" & strSearch & "%'"
        Else
            strFilter = "(" & strFilter & " And RoleName Like '" & strSearch & "%') or (" & strFilter & " And ���� Like '" & strSearch & "%')"
        End If
    End If
    mrsRole.Filter = strFilter
    lvwRole.ListItems.Clear
    
    With mrsRole
        Do Until .EOF
            If IsNull(!ϵͳ) Or cmbSystem.ItemData(cmbSystem.ListIndex) <> -1 Then
            Set objItem = lvwRole.ListItems.Add(, Nvl(!Role), Nvl(!RoleName), "Role", "Role")
            Else
                Set objItem = lvwRole.ListItems.Add(, Nvl(!Role), Nvl(!RoleName), "Role_Moved", "Role_Moved")
            End If
            If Not objItem Is Nothing Then
                objItem.SubItems(LRH_Grantee) = Nvl(!Grantee)
                objItem.SubItems(LRH_Admin_Option) = Nvl(!Admin_Option)
                objItem.SubItems(LRH_Group) = IIf(IsNull(!����), "UNGROUP", "K" & !����)
                objItem.SubItems(LRH_System) = IIf(IsNull(!ϵͳ), "-1", !ϵͳ)
                objItem.SubItems(LRH_SystemName) = IIf(IsNull(!ϵͳ����), "����ϵͳ", !ϵͳ����)
            End If
            mrsRole.MoveNext
        Loop
    End With
    If lvwRole.ListItems.Count > 0 Then
        lvwRole.ListItems(LRH_Grantee).Selected = True
        Call lvwRole_ItemClick(lvwRole.SelectedItem)
    Else
        cmdGrant.Enabled = False
        Call SetEnable
    End If
    mrsRole.Filter = 0
    
    FillRoleData = True
    
    Exit Function
ErrHandle:
    MsgBox "����" & err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
    If 1 = 0 Then
        Resume
    End If
End Function

Private Sub InitRoleData()
    '��ʼ����¼�������ÿ�������еĽ�ɫ����
    Dim Node As Node
    
    On Error GoTo errH
    gstrSQL = _
        " Select User as Grantee,'YES' as Admin_Option,Substr(A.����,4) as RoleName," & _
        " A.���� Role,zlSpellCode(Substr(A.����,4)) as ����, b.����, a.ϵͳ, c.���� ϵͳ����" & _
        " From zlTools.Zlroles A,zlTools.zlRoleGroups B, zlSystems C" & _
        " Where A.����=B.��ɫ(+) and A.ϵͳ = C.���(+)" & _
        IIf(cmbSystem.ItemData(cmbSystem.ListIndex) = -1, "", " And A.ϵͳ = [1]") & _
        " Order by A.����"
    Set mrsRole = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, Me.Caption, cmbSystem.ItemData(cmbSystem.ListIndex))
    
    For Each Node In tvwGroups.Nodes
        If Node.Key <> "Root" Then
            mrsRole.Filter = IIf(Node.Key = "unGroup", "���� = null", "���� = '" & Mid(Node.Text, 1, Len(Node.Text) - Len(Node.Tag)) & "'") & ""
        Else
            mrsRole.Filter = ""
        End If
        If Node.Tag = "" Then
            Node.Text = Node.Text & "(" & mrsRole.RecordCount & ")"
        Else
            Node.Text = Mid(Node.Text, 1, Len(Node.Text) - Len(Node.Tag)) & "(" & mrsRole.RecordCount & ")"
        End If
        Node.Tag = "(" & mrsRole.RecordCount & ")"
    Next
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Function SearchRole(ByVal strFilter As String) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------
    '����:���ǳ���Ӧ�Ľ�ɫ
    '����:strFilter-���˴�
    '����:�ɹ�,����ture,���򷵻�False
    '----------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    SearchRole = True
    If mrsRole Is Nothing Then Exit Function
    If mrsRole.State <> 1 Then Exit Function
    If mrsRole.RecordCount = 0 Then Exit Function
    
    strFilter = UCase(strFilter)
    SearchRole = False
    If strFilter = "" Then
    Else
        mrsRole.Filter = "RoleName Like '" & strFilter & "%' or ���� Like '" & strFilter & "%'"
    End If
    lvwRole.ListItems.Clear
    Do Until mrsRole.EOF
        lvwRole.ListItems.Add , Nvl(mrsRole!Role), Nvl(mrsRole!RoleName), "Role", "Role"
        mrsRole.MoveNext
    Loop
    If lvwRole.ListItems.Count > 0 Then
        lvwRole.ListItems(LRH_Grantee).Selected = True
    Else
        cmdGrant.Enabled = False
    End If
    Call SetEnable
    mrsRole.Filter = 0
    SearchRole = True
End Function

Private Sub txtSearch_Change()
    Call SearchRole(Trim(txtSearch.Text))
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Or KeyAscii = Asc("*") Or KeyAscii = Asc("_") Then
        KeyAscii = 0
    End If
End Sub
Private Sub FillRollGroup()
    '--------------------------------------------------------------------------------------------
    '����:���ؽ�ɫ��
    '����:���˺�
    '����:2007/06/15
    '--------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    Dim objNode As Node
    gstrSQL = "Select distinct ���� From zlRoleGroups"
    Call OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    With tvwGroups
        .Nodes.Clear
        Set objNode = tvwGroups.Nodes.Add(, 4, "Root", "���н�ɫ", 1, 1)
        objNode.Selected = True
        objNode.Expanded = True
        Set objNode = tvwGroups.Nodes.Add("Root", tvwChild, "unGroup", "δ����", 1, 1)
        objNode.Sorted = True
        Call LoadMenu("δ����", "unGroup")
        Do While Not rsTemp.EOF
            Set objNode = tvwGroups.Nodes.Add("Root", 4, "K" & Nvl(rsTemp!����), Nvl(rsTemp!����), 1, 1)
            objNode.Sorted = True
            Call LoadMenu(Nvl(rsTemp!����), Nvl(rsTemp!����))
            rsTemp.MoveNext
        Loop
    End With
End Sub
Private Sub LoadMenu(ByVal strTittle As String, ByVal strTag As String, Optional ByVal blnType As Boolean = True)
'���ء���ɫ�ƶ������顱����ɫ�ƶ����˵��������б�
'blnType = true : ���ط���˵�
'blnType = false : ����ϵͳ�˵�

        Dim varMenu As Variant
        Dim intCount As Integer
        
        If blnType = True Then
            Set varMenu = mnuPopuRoleMoveGroups
            intCount = varMenu.Count
            Load varMenu(intCount)
            varMenu(intCount).Caption = strTittle
            If strTag = "unGroup" Then
                varMenu(intCount).Tag = UCase(strTag)
            Else
                varMenu(intCount).Tag = UCase("K" & strTag)
            End If
            varMenu(intCount).Visible = True
            mnuPopuRoleMove1.Visible = True
            varMenu(0).Visible = False
        Else
            Set varMenu = mnuPopuRoleMoveSystems
            intCount = varMenu.Count
            Load varMenu(intCount)
            varMenu(intCount).Caption = strTittle
            varMenu(intCount).Tag = strTag
            varMenu(intCount).Visible = True
            mnuPopuRoleMove2.Visible = True
            varMenu(0).Visible = False
        End If
End Sub
Private Sub LoadMenus()
    Dim objNode As Node
    Call UnLoadMenus
    For Each objNode In tvwGroups.Nodes
        If objNode.Key <> "Root" Then
            Call LoadMenu(Mid(objNode.Text, 1, Len(objNode.Text) - Len(objNode.Tag)), IIf(objNode.Key = "unGroup", objNode.Key, Mid(objNode.Key, 2)))
        End If
    Next
End Sub
Private Sub UnLoadMenus()
    '����:��ж�˵�
        Dim varMenu As Variant
        Dim intCount As Integer
        Set varMenu = mnuPopuRoleMoveGroups
        mnuPopuRoleMoveGroups(0).Visible = True
        mnuPopuRoleMove1.Visible = True
        For intCount = 1 To mnuPopuRoleMoveGroups.UBound
            Unload varMenu(intCount)
        Next
        
End Sub

