VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPresRoleBat 
   Caption         =   "������ɫ����"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7710
   Icon            =   "frmDeptRole.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   7710
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancel 
      Caption         =   "�˳�"
      Height          =   350
      Left            =   3600
      TabIndex        =   8
      Top             =   5640
      Width           =   900
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "����Ȩ��"
      Height          =   350
      Left            =   3600
      TabIndex        =   7
      Top             =   4680
      Width           =   900
   End
   Begin VB.CommandButton cmdGrant 
      Caption         =   "����Ȩ��"
      Height          =   350
      Left            =   3600
      TabIndex        =   6
      Top             =   3720
      Width           =   900
   End
   Begin VB.ListBox lstRole 
      Height          =   1680
      Left            =   600
      TabIndex        =   5
      Top             =   960
      Width           =   2895
   End
   Begin VB.PictureBox picPerson 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   7710
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   7710
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "���ң�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   4
         Top             =   180
         Width           =   585
      End
   End
   Begin VB.ComboBox cboSystem 
      Height          =   300
      Left            =   4710
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   2895
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Left            =   570
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "B.����"
      Top             =   600
      Width           =   2895
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   480
      TabIndex        =   0
      Top             =   2880
      Width           =   7800
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   120
      Top             =   7320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptRole.frx":6852
            Key             =   "Role"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptRole.frx":752C
            Key             =   "NO"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptRole.frx":DD8E
            Key             =   "YES"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1560
      Top             =   7560
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
            Picture         =   "frmDeptRole.frx":145F0
            Key             =   "YES"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptRole.frx":1AE52
            Key             =   "NO"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwModule 
      Height          =   1725
      Left            =   4710
      TabIndex        =   9
      Top             =   960
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3043
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ϵͳ"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ģ��"
         Object.Width           =   6174
      EndProperty
   End
   Begin MSComctlLib.ListView lvwGrant 
      Height          =   4005
      Left            =   570
      TabIndex        =   10
      Top             =   3240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   7064
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ColHdrIcons     =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   2471
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "�û���"
         Object.Width           =   2293
      EndProperty
   End
   Begin MSComctlLib.ListView lvwRemove 
      Height          =   4005
      Left            =   4710
      TabIndex        =   11
      Top             =   3240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   7064
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ColHdrIcons     =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "�û���"
         Object.Width           =   2293
      EndProperty
   End
   Begin VB.Label lblRole 
      Caption         =   "��ɫ"
      Height          =   255
      Left            =   98
      TabIndex        =   17
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblNO��Ա 
      Caption         =   "δ����ý�ɫ����Ա"
      Height          =   315
      Left            =   570
      TabIndex        =   16
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label lbl��Ա 
      Caption         =   "������ý�ɫ����Ա"
      Height          =   195
      Left            =   4710
      TabIndex        =   15
      Top             =   3000
      Width           =   1785
   End
   Begin VB.Label lbl����Ȩ 
      Caption         =   "ģ���嵥"
      Height          =   255
      Left            =   3825
      TabIndex        =   14
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblContent 
      AutoSize        =   -1  'True
      Caption         =   "��Ȩ����"
      Height          =   180
      Left            =   3825
      TabIndex        =   13
      Top             =   660
      Width           =   735
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   2
      Left            =   98
      TabIndex        =   12
      Top             =   660
      Width           =   360
   End
End
Attribute VB_Name = "frmPresRoleBat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngDept As Long

Private Sub FillRole()
    Const STR_ICON = "Role"
    Dim rsTmp As ADODB.Recordset
    Dim lstTmp As ListItem
    Dim i As Long

    On Error GoTo errHandle
    gstrSQL = "Select a.��ɫ, Decode(b.��ɫ, Null, 0, 1) As Ӧ��" & vbNewLine & _
            "  From (Select Substr(Granted_Role, 4) ��ɫ" & vbNewLine & _
            "           From Dba_Role_Privs" & vbNewLine & _
            "          Where Granted_Role Like 'ZL_%' And Admin_Option = 'YES' And Grantee = User) a, (Select Distinct Substr(B1.Granted_Role, 4) ��ɫ" & vbNewLine & _
            "           From Dba_Role_Privs B1" & vbNewLine & _
            "          Where B1.Granted_Role Like" & vbNewLine & _
            "                'ZL_%') b" & vbNewLine & _
            " Where a.��ɫ = b.��ɫ(+)" & vbNewLine & _
            " Order By a.��ɫ"

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Caption & "-�û���ɫ")
    With Me.lstRole
        For i = 1 To rsTmp.RecordCount
            .AddItem rsTmp!��ɫ
            rsTmp.MoveNext
        Next
        rsTmp.Close
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub ShowMe(ByVal lngDept As Long, ByVal strDeptName As String)
    mlngDept = lngDept
    lblName.Caption = "���ң�" & strDeptName
    Me.Show
End Sub

Private Sub cboRole_Click()
'    LoadInfo
    LoadPreson
End Sub

Private Sub cboSystem_Click()
    FillModule
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGrant_Click()
    Dim i As Integer
    Dim strErr As String
    
    On Error Resume Next
    
    For i = 1 To Me.lvwGrant.ListItems.Count
        If Me.lvwGrant.ListItems(i).Checked Then
            gstrSQL = "Grant ZL_" & Me.lstRole.Text & "  to " & Me.lvwGrant.ListItems(i).SubItems(1)
            gcnOracle.Execute gstrSQL
            If Err <> 0 Then
                strErr = vbCrLf & Err.Description
'                MsgBox "Ȩ�޲��㣬�����ɫʧ�ܡ�" & vbCrLf & "������Ϣ����:" & vbCrLf & Err.Description, vbExclamation, gstrSysName
            End If
        End If
    Next
    
    If strErr <> "" Then
        MsgBox "Ȩ�޲��㣬�����ɫʧ�ܡ�" & vbCrLf & "������Ϣ����:" & strErr, vbExclamation, gstrSysName
    End If
    
    LoadPreson
End Sub

Private Sub cmdRemove_Click()
    Dim i As Integer
    Dim strErr As String
    On Error Resume Next
    
    For i = 1 To Me.lvwRemove.ListItems.Count
        If Me.lvwRemove.ListItems(i).Checked Then
            gstrSQL = "revoke ZL_" & Me.lstRole.Text & " from " & Me.lvwRemove.ListItems(i).SubItems(1)
            gcnOracle.Execute gstrSQL
            If Err <> 0 Then
                strErr = vbCrLf & Err.Description
'                MsgBox "Ȩ�޲��㣬�����ɫʧ�ܡ�" & vbCrLf & "������Ϣ����:" & vbCrLf & Err.Description, vbExclamation, gstrSysName
            End If
        End If
    Next
    
    If strErr <> "" Then
        MsgBox "Ȩ�޲��㣬�����ɫʧ�ܡ�" & vbCrLf & "������Ϣ����:" & strErr, vbExclamation, gstrSysName
    End If
    LoadPreson
End Sub

Private Sub Form_Load()
    FillRole
    
    FillSystem
    
End Sub

'Private Sub LoadInfo()
'    Dim rsTemp As Recordset
'    Dim lstTmp As ListItem
'
'    On Error GoTo errHandle
'    gstrSQL = "Select Distinct a.����, b.����,B.���" & vbNewLine & _
'    "From Zlsystems a, Zlprograms b, Zlrolegrant c" & vbNewLine & _
'    "Where a.��� = c.ϵͳ And b.ϵͳ = c.ϵͳ And b.��� = c.��� And c.��ɫ=[1]"
'    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "LoadInfo", "ZL_" & Me.lstRole.Text)
'
'    With Me.lvwRole
'        .ListItems.Clear
'        Do While Not rsTemp.EOF
'            Set lstTmp = .ListItems.Add(, "A" & rsTemp!���, rsTemp!����)
'            lstTmp.SubItems(1) = rsTemp!����
'            rsTemp.MoveNext
'        Loop
'    End With
'    Exit Sub
'errHandle:
'    If ErrCenter = 1 Then Resume
'    Call SaveErrLog
'End Sub

Private Sub LoadPreson()
    Dim rsTmp As Recordset
    Dim lstTmp As ListItem
    
    On Error GoTo errHandle
    gstrSQL = "Select a.�û���, d.����,decode(nvl(e.Granted_Role,''),'',0,1) Ȩ��" & vbNewLine & _
            "From �ϻ���Ա�� a, ������Ա b, ���ű� c, ��Ա�� d, Dba_Role_Privs e" & vbNewLine & _
            "Where a.��Աid = b.��Աid And b.����id = c.Id And d.Id = a.��Աid And A.�û���=e.Grantee(+) And b.����id =[1] And e.Granted_Role(+) = [2] " & _
            "And a.�û���<>[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Caption & "-�û���ɫ", mlngDept, "ZL_" & Me.lstRole.Text, gstrDbUser)
    
    Me.lvwGrant.ListItems.Clear
    Me.lvwRemove.ListItems.Clear
    Do While Not rsTmp.EOF
        If rsTmp!Ȩ�� = 1 Then
            Set lstTmp = lvwRemove.ListItems.Add(, rsTmp!�û���, rsTmp!����, , "YES")
            lstTmp.SubItems(1) = rsTmp!�û���
        Else
            Set lstTmp = lvwGrant.ListItems.Add(, rsTmp!�û���, rsTmp!����, , "NO")
            lstTmp.SubItems(1) = rsTmp!�û���
        End If
        lstTmp.Checked = True
        rsTmp.MoveNext
    Loop
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub FillSystem()
    Dim rsTemp As New ADODB.Recordset
    Dim strSystem As String
    Dim i As Integer
    
    On Error GoTo errHandle
    rsTemp.CursorLocation = adUseClient
    
    '��ʾ�������е�ϵͳ
'    gstrSQL = "Select ���, ����, �����, ������, ��װ����, ������װ, �汾�� From zlSystems order by ���"
'    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Caption & "-�Ѱ�װϵͳ")
    cboSystem.Clear
    
    For i = 1 To lstRole.ListCount
        gstrSQL = "Select distinct M.���, M.����, M.�����, M.������, M.��װ����,M.������װ, M.�汾�� " & _
                  "  From zlRoleGrant R, zlPrograms P,zlsystems M" & _
                  "  Where Nvl(r.ϵͳ, 0) = Nvl(p.ϵͳ, 0) And p.��� = r.��� And p.��� >= 100 And Substr(r.��ɫ, 4) = [1] and m.���=p.ϵͳ" & _
                    "  Order By m.���"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Caption & "-�Ѱ�װϵͳ", lstRole.List(i))
        Do Until rsTemp.EOF
            If InStr(1, strSystem, rsTemp!��� & "|") = 0 Then
                cboSystem.AddItem rsTemp("����") & " v" & rsTemp("�汾��") & "��" & rsTemp("���") & "��"
                strSystem = strSystem & "|" & rsTemp!��� & "|"
                cboSystem.ItemData(cboSystem.NewIndex) = rsTemp("���")
                If rsTemp("������") = UCase(gstrUserName) And cboSystem.ListIndex < 0 Then
                    cboSystem.ListIndex = cboSystem.NewIndex
                End If
            End If
            rsTemp.MoveNext
        Loop
    Next
    
    '������ϵͳ�ǳ���̶���
    If (zlRegTool And 2) = 2 Then cboSystem.AddItem "�Զ��屨��"
    cboSystem.AddItem "��������"
    cboSystem.AddItem "ȡ������"
    cboSystem.AddItem "��������"
    If cboSystem.ListIndex < 0 Then cboSystem.ListIndex = 0
    Exit Sub

errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
'    MsgBox Err.Description, vbCritical, Me.Caption
End Sub


Private Sub lstRole_Click()
    LoadPreson
    FillModule
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    
    If KeyAscii <> 13 Then Exit Sub
    
    If Me.txtEdit.Text = "" Then Exit Sub
    
    zlControl.TxtSelAll txtEdit
    
    For i = 0 To Me.lstRole.ListCount
        If InStr(1, Me.lstRole.List(i), Me.txtEdit.Text) > 0 Then
            Me.lstRole.Selected(i) = True
            Exit Sub
        End If
    Next
End Sub

Private Sub FillModule()
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem
    Dim strRole As String
    
'    LockWindowUpdate lvwModule.hwnd
    On Error GoTo errHandle
    lvwModule.ColumnHeaders.Clear
    lvwModule.ListItems.Clear
    If lstRole.ListIndex <> -1 Then
        strRole = lstRole.List(lstRole.ListIndex)
    End If
    '�����б���
    With lvwModule.ColumnHeaders
        If cboSystem.Text = "��������" Then
'            lblModule.Caption = "�ɹ���ı����"
            .Add , , "�����", "1200"
            .Add , , "����ϵͳ", "2100"
            .Add , , "˵��", "2500"
        ElseIf cboSystem.Text = "ȡ������" Then
'            lblModule.Caption = "�ɵ��õĺ���"
            .Add , , "������", "1200"
            .Add , , "������", "1500"
            .Add , , "����ϵͳ", "2100"
            .Add , , "˵��", "2500"
        ElseIf cboSystem.Text = "��������" Then
'            lblModule.Caption = "����Ȩ�Ļ�������"
            .Add , , "���", "600"
            .Add , , "����", "1800"
            .Add , , "˵��", "3000"
'            .Add , , "��Ȩ����", "5000"
        Else
'            lblModule.Caption = "����Ȩģ��"
            .Add , , "���", "600"
            .Add , , "����", "1800"
            .Add , , "˵��", "3000"
'            .Add , , "��Ȩ����", "5000"
        End If
    End With
'    lnModuel.X1 = lblModule.Left + lblModule.Width
    
    If strRole = "" Then
        '��ɫΪ�գ��˳�
'        LockWindowUpdate 0
        Exit Sub
    End If

    If cboSystem.Text = "��������" Then
        '��ʾ�ý�ɫ�ܷ��ʵĻ�����
        gstrSQL = "select T.ϵͳ,T.����,T.˵�� from " & _
                "(SELECT S.����||'��'||S.���||'��' as ϵͳ,S.������,B.����,B.˵�� FROM zlSystems S,zlBaseCode B where B.ϵͳ=S.���) T,USER_TAB_PRIVS R " & _
                "where T.������=R.OWNER AND T.����=R.TABLE_NAME AND R.GRANTEE='" & strRole & _
                "' and R.PRIVILEGE in ('SELECT','INSERT','UPDATE','DELETE') " & _
                "GROUP BY T.ϵͳ,T.����,T.˵�� " & _
                "Having Count(R.PRIVILEGE) = 4"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "FillModule")
        Do Until rsTemp.EOF
            Set lst = lvwModule.ListItems.Add(, , rsTemp("����"))
            lst.SubItems(1) = rsTemp("ϵͳ")
            lst.SubItems(2) = IIF(IsNull(rsTemp("˵��")), "", rsTemp("˵��"))
            rsTemp.MoveNext
        Loop
    ElseIf cboSystem.Text = "ȡ������" Then
        '��ʾ�ý�ɫ�ܷ��ʵĻ�����
        gstrSQL = "select S.����||'��'||S.���||'��' as ϵͳ,S.������,F.������,F.������,F.˵�� " & _
                  " from zlSystems S,zlFunctions F,USER_TAB_PRIVS R " & _
                  " where  F.ϵͳ=S.��� and S.������=R.OWNER AND upper(F.������)=R.TABLE_NAME AND R.GRANTEE='" & strRole & "' and R.PRIVILEGE ='EXECUTE'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "FillModule")
        Do Until rsTemp.EOF
            Set lst = lvwModule.ListItems.Add(, , rsTemp("������"))
            lst.SubItems(1) = rsTemp("������")
            lst.SubItems(2) = rsTemp("ϵͳ")
            lst.SubItems(3) = IIF(IsNull(rsTemp("˵��")), "", rsTemp("˵��"))
            rsTemp.MoveNext
        Loop
    ElseIf cboSystem.Text = "��������" Then
        '��ʾ�ý�ɫ�ܷ��ʵĻ�������
        gstrSQL = "select P.���,P.����,P.˵��,R.���� from " & _
                "zlRoleGrant R,zlPrograms P " & _
                "where R.ϵͳ is null and P.���=R.��� AND substr(R.��ɫ,4)='" & strRole & _
                "'  AND P.ϵͳ is null and P.���<100 and P.���� is null " & _
                " Order By P.���"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "FillModule")
        
        On Error Resume Next
        Do Until rsTemp.EOF
            Set lst = lvwModule.ListItems.Add(, "C" & rsTemp("���"), rsTemp("���"))
            If Err <> 0 Then
                Err.Clear
                If rsTemp("����") <> "����" Then
                    Set lst = lvwModule.ListItems("C" & rsTemp("���"))
'                    lst.SubItems(3) = IIF(lst.SubItems(3) = "", "", lst.SubItems(3) & ",") & rsTemp("����")
                End If
            Else
                lst.SubItems(1) = rsTemp("����")
                lst.SubItems(2) = IIF(IsNull(rsTemp("˵��")), "", rsTemp("˵��"))
'                If rsTemp("����") <> "����" Then
'                    lst.SubItems(3) = rsTemp("����")
'                End If
            End If
            rsTemp.MoveNext
        Loop
    Else
        '��ʾ�ý�ɫ�ܷ��ʵ�ģ��
        gstrSQL = "select P.���,P.����,P.˵��,R.���� from " & _
                "zlRoleGrant R,zlPrograms P " & _
                "where nvl(R.ϵͳ,0)=nvl(P.ϵͳ,0) and P.���=R.��� and P.���>=100 AND substr(R.��ɫ,4)='" & strRole & "'  AND " & _
                IIF(cboSystem.Text = "�Զ��屨��", " P.ϵͳ is null ", " P.ϵͳ=" & cboSystem.ItemData(cboSystem.ListIndex)) & _
                " Order By P.���"
         Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "FillModule")
        On Error Resume Next
        Do Until rsTemp.EOF
            Set lst = lvwModule.ListItems.Add(, "C" & rsTemp("���"), rsTemp("���"))
            If Err <> 0 Then
                Err.Clear
                If rsTemp("����") <> "����" Then
                    Set lst = lvwModule.ListItems("C" & rsTemp("���"))
'                    lst.SubItems(3) = IIF(lst.SubItems(3) = "", "", lst.SubItems(3) & ",") & rsTemp("����")
                End If
            Else
                lst.SubItems(1) = rsTemp("����")
                lst.SubItems(2) = IIF(IsNull(rsTemp("˵��")), "", rsTemp("˵��"))
'                If rsTemp("����") <> "����" Then
'                    lst.SubItems(3) = rsTemp("����")
'                End If
            End If
            rsTemp.MoveNext
        Loop
    End If
    
'    LockWindowUpdate 0
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub
