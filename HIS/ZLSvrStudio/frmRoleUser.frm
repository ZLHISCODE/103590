VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRoleUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ɫ��Ȩ�û�"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   Icon            =   "frmRoleUser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdClear 
      Caption         =   "ȫ��(&R)"
      Height          =   345
      Left            =   7320
      TabIndex        =   8
      ToolTipText     =   "Ctrl + R"
      Top             =   2055
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelAll 
      Caption         =   "ȫѡ(&A)"
      Height          =   345
      Left            =   7320
      TabIndex        =   7
      ToolTipText     =   "Ctrl + A"
      Top             =   1695
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7320
      TabIndex        =   5
      Top             =   600
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7320
      TabIndex        =   6
      Top             =   960
      Width           =   1100
   End
   Begin MSComctlLib.ImageList Img��ͼ�� 
      Left            =   7395
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoleUser.frx":000C
            Key             =   "User"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgСͼ�� 
      Left            =   7410
      Top             =   2850
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
            Picture         =   "frmRoleUser.frx":0326
            Key             =   "User"
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboSystem 
      Height          =   300
      Left            =   4020
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   135
      Width           =   3120
   End
   Begin MSComctlLib.ListView lvwUser 
      Height          =   4395
      Left            =   105
      TabIndex        =   3
      Top             =   480
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   7752
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "Img��ͼ��"
      SmallIcons      =   "ImgСͼ��"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "�û���"
         Object.Width           =   3422
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "��Ա���"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "��Ա����"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "��������"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Key             =   "Admin"
         Text            =   "����ת��"
         Object.Width           =   1587
      EndProperty
   End
   Begin VB.CheckBox chkOnlyGranted 
      Caption         =   "ֻ��ʾ��ɫ����Ȩ�û�(&O)"
      BeginProperty DataFormat 
         Type            =   4
         Format          =   "H:mm:ss"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   8
      EndProperty
      Height          =   195
      Left            =   165
      TabIndex        =   4
      Top             =   4950
      Width           =   2445
   End
   Begin VB.Label lblRole 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ǰ��ɫ��"
      ForeColor       =   &H00000080&
      Height          =   180
      Left            =   165
      TabIndex        =   0
      Top             =   210
      Width           =   900
   End
   Begin VB.Label lblSys 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ӧ��ϵͳ"
      Height          =   180
      Left            =   3240
      TabIndex        =   1
      Top             =   195
      Width           =   720
   End
End
Attribute VB_Name = "frmRoleUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==============================================================
'==ģ�����
'==============================================================
Private mstrRole As String
Private Enum Cols
    Col_��� = 1
    Col_���� = 2
    Col_���� = 3
    Col_ת�� = 4
End Enum
Private mrsSystem As New ADODB.Recordset
Private mintColumn As Integer '
Private mblnAdminCol As Boolean
Private mstrAllOwner As String
Private mlngSysNo As Long
'==============================================================
'==�����ӿ�
'==============================================================

Public Function ShowEdit(ByVal strRole As String) As Boolean
    mstrRole = strRole
    Set mrsSystem = gclsBase.GetSystems()
    If mrsSystem.RecordCount = 0 Then
        MsgBox "δ��װӦ��ϵͳ������ѡ����Ȩ�û���", vbInformation, gstrSysName
        Exit Function
    End If
    '����Ϊֻ���ؾ��в�����Ա����������ӵ�е�ϵͳ
    Set mrsSystem = gclsBase.GetMenSystems()
    If mrsSystem.RecordCount = 0 Then
        MsgBox "��װ��Ӧ��ϵͳ�в����ڲ�����Ա��������ѡ����Ȩ�û���", vbInformation, gstrSysName
        Exit Function
    End If
    mstrAllOwner = ""
    frmRoleUser.Show vbModal, frmMDIMain
End Function

'==============================================================
'==�ؼ��¼�
'==============================================================
Private Sub chkOnlyGranted_Click()
    Call FillUser
End Sub

Private Sub chkOnlyGranted_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cboSystem_Click()
    cboSystem.Tag = cboSystem.ListIndex
    mlngSysNo = cboSystem.ItemData(cboSystem.ListIndex)
    Call FillUser
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Call SelItem
End Sub

Private Sub cmdOK_Click()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim cnTmp As ADODB.Connection
    Dim lst As ListItem
    Dim strMsg As String
    
    On Error GoTo errH
    '��ǰ�û��Ƿ���Ȩ�޴���û������SYSTEM�û�
    strSQL = "Select 1 From User_Role_Privs Where Admin_Option = 'YES' And Granted_Role = '" & mstrRole & "'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        Set cnTmp = gcnOracle
    Else
        Set gcnSystem = GetConnection("SYSTEM")
        If gcnSystem Is Nothing Then Exit Sub
        Set cnTmp = gcnSystem
    End If
    On Error Resume Next
    Screen.MousePointer = 11
    For Each lst In lvwUser.ListItems
        '��Ȩ�仯��ת�ڷ����仯�ұ�ѡ��
        If lst.Tag <> CStr(lst.Checked) Or lst.SubItems(Col_ת��) <> lst.ListSubItems(Col_ת��).Tag And lst.Checked Then
            '��Ҫ�ɷ�ת�ڱ�Ϊת�ڣ���Ҫ��ȡ����Ȩ
            If lst.SubItems(Col_ת��) <> lst.ListSubItems(Col_ת��).Tag And lst.Checked And lst.Tag = CStr(lst.Checked) Then
                Call gclsBase.ExecuteCmdText("Revoke " & mstrRole & " From " & lst.Text, Me.Caption, cnTmp)
                Call ExecuteProcedure("Zl_Zluserroles_Del('" & lst.Text & "','" & mstrRole & "')", Me.Caption, cnTmp)
            End If
            '��Ȩ������仯
            If lst.Checked Then
                Call gclsBase.ExecuteCmdText("Grant " & mstrRole & " To " & lst.Text & IIf(lst.SubItems(Col_ת��) <> "", " With Admin Option", ""), Me.Caption, cnTmp)
                Call ExecuteProcedure("Zl_Zluserroles_Add('" & lst.Text & "','" & mstrRole & "'," & IIf(lst.SubItems(Col_ת��) <> "", 1, 0) & ")", Me.Caption, cnTmp)
            Else
                Call gclsBase.ExecuteCmdText("Revoke " & mstrRole & " From " & lst.Text, Me.Caption, cnTmp)
                Call ExecuteProcedure("Zl_Zluserroles_Del('" & lst.Text & "','" & mstrRole & "')", Me.Caption, cnTmp)
            End If
        End If
    Next
    Screen.MousePointer = 0
    If err.Number <> 0 Then
        MsgBox "���û�������Ȩʱ���ִ�����Щ�û�δ�ɹ��޸���Ȩ��" & vbNewLine & err.Description, vbInformation, gstrSysName
        strMsg = "���û�������Ȩʱ���ִ�����Щ�û�δ�ɹ��޸���Ȩ��"
        err.Clear
    Else
        MsgBox "�޸��û���Ȩ�ɹ���", vbInformation, gstrSysName
        strMsg = "�޸��û���Ȩ�ɹ���"
    End If
    '������Ҫ������־
    Call SaveAuditLog(2, "�޸Ľ�ɫ����Ȩ�û�", strMsg)
    Unload Me
    Exit Sub
errH:
    Screen.MousePointer = 0
    MsgBox "����" & err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
End Sub

Private Sub cmdSelAll_Click()
    Call SelItem(1)
End Sub

Private Sub Form_Activate()
    Call lvwUser.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyA Then
        Call cmdSelAll_Click
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyR Then
        Call cmdClear_Click
    End If
End Sub

Private Sub Form_Load()
    lblRole.Caption = "��ǰ��ɫ��" & mstrRole
    mstrRole = "ZL_" & mstrRole
    mlngSysNo = glngSysNo
    Call FillSystem
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mrsSystem Is Nothing Then Set mrsSystem = Nothing
End Sub

Private Sub lvwUser_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwUser.Sorted = True
    If mintColumn = ColumnHeader.Index - 1 Then '���Ǹղ�����
        lvwUser.SortOrder = IIf(lvwUser.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwUser.SortKey = mintColumn
        lvwUser.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwUser_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        '����ʱ���ָ�ԭʼ��ת��״̬
        Item.SubItems(4) = Item.ListSubItems(4).Tag
    Else
        Item.SubItems(4) = ""
    End If
End Sub

Private Sub lvwUser_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objItem As ListItem
    
    mblnAdminCol = False
    Set objItem = lvwUser.HitTest(x, y)
    If Not objItem Is Nothing Then
        If x >= lvwUser.ColumnHeaders("Admin").Left Then
            mblnAdminCol = True
        End If
    End If
End Sub

Private Sub lvwUser_DblClick()
    If mblnAdminCol Then
        Call SelItem(-1)
    End If
    mblnAdminCol = False
End Sub

'==============================================================
'==˽�з���
'==============================================================
Private Sub FillSystem()
'���ܣ�����Ӧ��ϵͳ
    If glngSysNo <> -1 Then
        lblSys.Visible = False
        cboSystem.Visible = False
        Call FillUser
        Exit Sub
    End If
    cboSystem.Clear: cboSystem.Tag = ""
    Do While Not mrsSystem.EOF
        cboSystem.AddItem mrsSystem!���� & " v" & mrsSystem!�汾�� & "��" & mrsSystem!��� & "��"
        cboSystem.ItemData(cboSystem.NewIndex) = mrsSystem!���
        If mrsSystem!������ & "" = UCase(gstrUserName) And cboSystem.Tag = "" Then
            cboSystem.Tag = cboSystem.NewIndex
        End If
        mrsSystem.MoveNext
    Loop
    cboSystem.ListIndex = Val(cboSystem.Tag)
End Sub

Private Sub FillUser()
'���ܣ������û�
    Dim strTmp As String, strOwner As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lst As ListItem, blnOnlyGrant As Boolean
    
    On Error GoTo errH
    If mstrAllOwner = "" Then
        mrsSystem.Filter = "": mrsSystem.Sort = "�����,���"
        Do While Not mrsSystem.EOF
            If strTmp <> mrsSystem!������ Then
                strTmp = mrsSystem!������
                mstrAllOwner = mstrAllOwner & ",'" & strTmp & "'"
            End If
            mrsSystem.MoveNext
        Loop
        strSQL = "Select Upper(������) ������ From Zlbakspaces Where Db���� Is Null Order by ������"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
        Do While Not rsTmp.EOF
            If strTmp <> rsTmp!������ Then
                strTmp = rsTmp!������
                mstrAllOwner = mstrAllOwner & ",'" & strTmp & "'"
            End If
            rsTmp.MoveNext
        Loop
    End If
    '��ʾ���Խ��е�ǰϵͳ���û����Ӧ����Ա
    mrsSystem.Filter = "���=" & mlngSysNo
    strOwner = mrsSystem!������
    strSQL = "Select u.Username, r.���, r.����, r.��������, p.Granted_Role, p.Admin_Option" & vbNewLine & _
                    "From All_Users u, Dba_Role_Privs p," & vbNewLine & _
                    "     (Select c.�û���, p.���, p.����, d.���� As ��������" & vbNewLine & _
                    "       From " & strOwner & ".��Ա�� p, " & strOwner & ".���ű� d, " & strOwner & ".������Ա b, " & strOwner & ".�ϻ���Ա�� c" & vbNewLine & _
                    "       Where p.Id = c.��Աid And c.��Աid = b.��Աid And d.Id = b.����id And" & vbNewLine & _
                    "             (p.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or p.����ʱ�� Is Null) And b.ȱʡ = 1) r" & vbNewLine & _
                    "Where u.Username = r.�û���(+) And (u.Username Not In (" & G_STR_USERS & mstrAllOwner & ")) And u.Username = p.Grantee(+) And" & vbNewLine & _
                    "      p.Granted_Role(+) = '" & mstrRole & "'" & vbNewLine & _
                    "Order By u.Username"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    lvwUser.ListItems.Clear: blnOnlyGrant = chkOnlyGranted.value = 1
    Do While Not rsTmp.EOF
        If blnOnlyGrant And IsNull(rsTmp!GRANTED_ROLE) Then
        Else
            Set lst = lvwUser.ListItems.Add(, "K" & rsTmp!USERNAME, rsTmp!USERNAME, "User", "User")
            lst.SubItems(Col_���) = rsTmp!��� & ""
            lst.SubItems(Col_����) = rsTmp!���� & ""
            lst.SubItems(Col_����) = rsTmp!�������� & ""
            lst.SubItems(Col_ת��) = IIf(Nvl(rsTmp!Admin_Option) = "YES", "��", "")
            lst.Checked = Not IsNull(rsTmp!GRANTED_ROLE)
            lst.Tag = CStr(lst.Checked) '��¼ԭʼ�Ƿ�����Ȩ
            lst.ListSubItems(Col_ת��).Tag = lst.SubItems(Col_ת��) '��¼ԭʼ�Ƿ�����ת��
        End If
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub SelItem(Optional ByVal intSel As Integer)
'���ܣ�ѡ���û�
'blnSel=0:ȫ��ȡ�� ��1-ȫ��ѡ��-1����ת��Ȩ��ѡ���з�ѡ
'intIndex>0:��ĳ�н��з�ѡ
    Dim objItem As ListItem, blnSel As Boolean
    If intSel <> -1 Then
        blnSel = intSel = 1
        For Each objItem In lvwUser.ListItems
            If objItem.Checked Then
                objItem.Checked = blnSel
                Call lvwUser_ItemCheck(objItem)
            End If
        Next
        lvwUser.SetFocus
    Else
        Set objItem = lvwUser.SelectedItem
        If Not objItem Is Nothing Then
            objItem.SubItems(Col_ת��) = IIf(objItem.SubItems(Col_ת��) = "", "��", "")
            If objItem.SubItems(Col_ת��) <> "" Then
                objItem.Checked = True
            End If
        End If
    End If
End Sub

