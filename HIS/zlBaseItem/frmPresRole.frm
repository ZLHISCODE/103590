VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPresRole 
   Caption         =   "��Ա��ɫ����"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9000
   Icon            =   "frmPresRole.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   9000
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picSplit 
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   5640
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4455
      ScaleWidth      =   45
      TabIndex        =   16
      Top             =   3480
      Width           =   50
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   7560
      TabIndex        =   15
      ToolTipText     =   "�����������Ľ�ɫ�����в��ң�"
      Top             =   540
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfUnGrantedFuncs 
      Height          =   2055
      Left            =   5760
      TabIndex        =   13
      Top             =   5880
      Width           =   3135
      _cx             =   5530
      _cy             =   3625
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPresRole.frx":000C
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfGrantedFuncs 
      Height          =   1815
      Left            =   5760
      TabIndex        =   12
      Top             =   3720
      Width           =   3135
      _cx             =   5530
      _cy             =   3201
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPresRole.frx":0053
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   5280
      Top             =   2880
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
            Picture         =   "frmPresRole.frx":009A
            Key             =   "Role"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7800
      TabIndex        =   8
      Top             =   8040
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6600
      TabIndex        =   7
      Top             =   8040
      Width           =   1100
   End
   Begin VB.ComboBox cboSystem 
      Height          =   300
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3120
      Width           =   3255
   End
   Begin VB.PictureBox picPerson 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   9000
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9000
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "������"
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
         TabIndex        =   1
         Top             =   180
         Width           =   585
      End
   End
   Begin MSComctlLib.ListView lvwModule 
      Height          =   4245
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   7488
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "��������"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "ȱʡ"
         Object.Width           =   970
      EndProperty
   End
   Begin MSComctlLib.ListView lvwRole 
      Height          =   2205
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   3889
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "��������"
         Object.Width           =   6174
      EndProperty
   End
   Begin VB.Label lblFind 
      Caption         =   "����(&F)"
      Height          =   195
      Left            =   6720
      TabIndex        =   14
      Top             =   593
      Width           =   705
   End
   Begin VB.Label lblGrantedFuncs 
      Caption         =   "����Ȩ����"
      Height          =   200
      Left            =   5760
      TabIndex        =   11
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label lblUnGrantedFuncs 
      Caption         =   "δ��Ȩ����"
      Height          =   200
      Left            =   5760
      TabIndex        =   10
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label lblModule 
      Caption         =   "ģ���嵥"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label lblSystem 
      AutoSize        =   -1  'True
      Caption         =   "��Ȩ����(&T)"
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   3165
      Width           =   990
   End
   Begin VB.Label lblRole 
      AutoSize        =   -1  'True
      Caption         =   "��Ա��ɫ(&R)"
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   990
   End
End
Attribute VB_Name = "frmPresRole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngPersonID    As Long        '��ԱID
Private mstrUser        As String      '��Ա��Ӧ�û�
Private mblnOk          As Boolean
Private mlngSysIdx      As Long         '��ǰѡ���ϵͳIndex
Private mlngRoleIdx     As String       '��ǰѡ��Ľ�ɫIndex
Private mlngModuleIdx   As Long         '��ǰѡ���ģ��Index
Private mblnLoad        As Boolean
Private Enum FuncCols
    Col_��� = 0
    Col_���� = 1
End Enum

Public Function ShowMe(ByVal frmParent As Object, ByVal lngPersonID As Long) As Boolean
'���ܣ���ں���
    mblnOk = False
    mlngPersonID = lngPersonID
    Me.Show vbModal, frmParent
    ShowMe = mblnOk
End Function

Private Sub cboSystem_Click()
    If mblnLoad Then Exit Sub
    If mlngSysIdx = cboSystem.ListIndex Then Exit Sub
    mlngSysIdx = cboSystem.ListIndex
    mlngModuleIdx = -1
    Call FillModule
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strRole As String
    Dim i As Long
    
    On Error Resume Next
    For i = 1 To lvwRole.ListItems.Count
        If lvwRole.ListItems(i).Checked = True Then
            strRole = strRole & "ZL_" & lvwRole.ListItems(i).Text & ","
        End If
    Next

    If strRole <> "" Then
        gstrSQL = "Grant " & Mid(strRole, 1, Len(strRole) - 1) & " to " & mstrUser
        gcnOracle.Execute gstrSQL, , adCmdText
        If Err <> 0 Then
            MsgBox "Ȩ�޲��㣬�����ɫʧ�ܡ�" & vbCrLf & "������Ϣ����:" & vbCrLf & Err.Description, vbExclamation, gstrSysName
            If lvwRole.Enabled = True Then lvwRole.SetFocus
            Exit Sub
        End If
    End If
    '��Ҫ�ջ�Ȩ��
    For i = 1 To lvwRole.ListItems.Count
        If lvwRole.ListItems(i).Checked = False Then
            gstrSQL = "revoke ZL_" & lvwRole.ListItems(i).Text & " from " & mstrUser
            gcnOracle.Execute gstrSQL, , adCmdText
        End If
    Next
    Call zlDatabase.ExecuteProcedure("Zl_Zluserroles_Add('" & mstrUser & "')", Me.Caption)
    If Err.Number <> 0 Then Err.Clear
    Unload Me
    mblnOk = True
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo ErrH
    strSQL = "Select ����, b.�û���" & vbNewLine & _
            "From ��Ա�� a, �ϻ���Ա�� b" & vbNewLine & _
            "Where Id = [1]" & vbNewLine & _
            "And a.Id = b.��Աid"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption & "-��Ա����", mlngPersonID)
    If Not rsTmp.EOF Then
        lblName.Caption = "������ " & rsTmp!����
        mstrUser = rsTmp!�û��� & ""
    End If
    '��ʼ��
    mlngModuleIdx = -1
    mlngSysIdx = -1
    mlngModuleIdx = -1
    mblnLoad = True
    lvwRole.Icons = ils32
    Call FillRoleAndSystem
    mblnLoad = False
    Call cboSystem_Click
    Exit Sub
ErrH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FillRoleAndSystem()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Integer
    Dim lstTmp As ListItem
    
    On Error GoTo ErrH
    strSQL = "Select a.��ɫ, Decode(d.��ɫ, Null, 0, 1) As ����Ȩ" & vbNewLine & _
            "From (Select Substr(Granted_Role, 4) ��ɫ" & vbNewLine & _
            "       From Dba_Role_Privs" & vbNewLine & _
            "       Where Granted_Role Like 'ZL_%'" & vbNewLine & _
            "       And Admin_Option = 'YES'" & vbNewLine & _
            "       And Grantee = User) a," & vbNewLine & _
            "     (Select Distinct Substr(b.Granted_Role, 4) ��ɫ" & vbNewLine & _
            "       From Dba_Role_Privs b, �ϻ���Ա�� c" & vbNewLine & _
            "       Where b.Grantee = c.�û���" & vbNewLine & _
            "       And b.Granted_Role Like 'ZL_%'" & vbNewLine & _
            "       And c.��Աid = [1]) d" & vbNewLine & _
            "Where a.��ɫ = d.��ɫ(+)" & vbNewLine & _
            "Order By a.��ɫ"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Caption & "-�û���ɫ", mlngPersonID)
    With lvwRole
        .ListItems.Clear
        For i = 1 To rsTmp.RecordCount
            Set lstTmp = .ListItems.Add(, "R" & Format(i, "00000"), rsTmp!��ɫ, "Role")
            lstTmp.Checked = rsTmp!����Ȩ = 1
            rsTmp.MoveNext
        Next
        cmdOK.Enabled = .ListItems.Count > 0
        If .ListItems.Count > 0 Then
            .ListItems(1).Selected = True
            mlngRoleIdx = .SelectedItem.Index
        End If
        
    End With
    strSQL = "Select Distinct m.���, m.����, m.�����, m.������, m.��װ����, m.������װ, m.�汾��" & vbNewLine & _
            "From (Select Distinct ��ɫ, ϵͳ From Zlrolegrant Where ��� >= 100) r, Zlsystems m," & vbNewLine & _
            "     (Select Granted_Role" & vbNewLine & _
            "       From Dba_Role_Privs" & vbNewLine & _
            "       Where Granted_Role Like 'ZL_%'" & vbNewLine & _
            "       And Admin_Option = 'YES'" & vbNewLine & _
            "       And Grantee = User) n" & vbNewLine & _
            "Where r.ϵͳ = m.���" & vbNewLine & _
            "And r.��ɫ = n.Granted_Role" & vbNewLine & _
            "Order By m.���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption & "-�Ѱ�װϵͳ")
    
    With cboSystem
        .Clear
        Do While Not rsTmp.EOF
            .AddItem rsTmp!���� & " v" & rsTmp!�汾�� & "��" & rsTmp!��� & "��"
            .ItemData(cboSystem.NewIndex) = rsTmp!���
            If rsTmp!������ = UCase(gstrUserName) And .ListIndex < 0 Then
                .ListIndex = .NewIndex
            End If
            rsTmp.MoveNext
        Loop
        '������ϵͳ�ǳ���̶���
        If (zlRegTool And 2) = 2 Then .AddItem "�Զ��屨��"
        .AddItem "��������"
        .AddItem "ȡ������"
        .AddItem "��������"
        If .ListIndex < 0 Then .ListIndex = 0
        mlngSysIdx = .ListIndex
    End With
 
    Exit Sub
ErrH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    txtFind.Move Me.ScaleWidth - 200 - txtFind.Width, lblRole.Top
    lblFind.Move txtFind.Left - lblFind.Width - 50, txtFind.Top + 30
    lvwRole.Move 50, txtFind.Top + txtFind.Height + 30, Me.ScaleWidth - 200, Me.ScaleHeight / 4
    lblSystem.Move 50, lvwRole.Top + lvwRole.Height + 150
    cboSystem.Move lblSystem.Width + lblSystem.Left + 50, lblSystem.Top - 50
    lblModule.Move lblSystem.Left, cboSystem.Top + cboSystem.Height + 100
    
    lvwModule.Move 50, lblModule.Top + lblModule.Height, (Me.ScaleWidth / 3) * 2, Me.ScaleHeight - lblModule.Top - lblModule.Height - 100 - cmdOK.Height - 100
    picSplit.Move lvwModule.Left + lvwModule.Width, lvwModule.Top, 50, lvwModule.Height
    vsfGrantedFuncs.Move lvwModule.Left + lvwModule.Width + 50, lvwModule.Top, Me.ScaleWidth - lvwModule.Width - 200, lvwModule.Height / 2 - 200
    lblUnGrantedFuncs.Move vsfGrantedFuncs.Left + 100, vsfGrantedFuncs.Top + vsfGrantedFuncs.Height + 100
    vsfUnGrantedFuncs.Move vsfGrantedFuncs.Left, lblUnGrantedFuncs.Top + lblUnGrantedFuncs.Height + 50, vsfGrantedFuncs.Width, (lvwModule.Height / 2) - 150
    cmdCancel.Move Me.ScaleWidth - cmdCancel.Width - 200, vsfUnGrantedFuncs.Top + vsfUnGrantedFuncs.Height + 120
    cmdOK.Move cmdCancel.Left - cmdOK.Width - 200, cmdCancel.Top
    lblGrantedFuncs.Move lblUnGrantedFuncs.Left, lblModule.Top
End Sub

Private Sub lvwModule_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim lngModule As Long, strRole As String, lngSys As Long
    Dim vsTmp As VSFlexGrid
    
    If mblnLoad Then Exit Sub
    If mlngModuleIdx = Item.Index Then Exit Sub
    mlngModuleIdx = Item.Index
    On Error GoTo ErrH
    If Not lvwRole.SelectedItem Is Nothing Then
        strRole = "ZL_" & lvwRole.SelectedItem.Text
    End If
    If strRole = "" Then Exit Sub
    If cboSystem.Text = "��������" Or cboSystem.Text = "ȡ������" Then Exit Sub
    lngModule = Val(Mid(Item.Key, 2))
    lngSys = cboSystem.ItemData(cboSystem.ListIndex)
    
    strSQL = "Select a.����, Decode(b.����, Null, 0, 1) ��Ȩ" & vbNewLine & _
            "From (Select ����" & vbNewLine & _
            "       From Zlprogfuncs" & vbNewLine & _
            "       Where Nvl(ϵͳ,0) = [1]" & vbNewLine & _
            "       And ��� = [2]) a," & vbNewLine & _
            "     (Select ����" & vbNewLine & _
            "       From Zlrolegrant" & vbNewLine & _
            "       Where Nvl(ϵͳ,0) = [1]" & vbNewLine & _
            "       And ��� = [2]" & vbNewLine & _
            "       And ��ɫ = [3]) b" & vbNewLine & _
            "Where a.���� = b.����(+)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngSys, lngModule, strRole)
    vsfGrantedFuncs.Rows = vsfGrantedFuncs.FixedRows
    vsfUnGrantedFuncs.Rows = vsfUnGrantedFuncs.FixedRows
    Do While Not rsTmp.EOF
        If rsTmp!��Ȩ = 1 Then
            Set vsTmp = vsfGrantedFuncs
        Else
            Set vsTmp = vsfUnGrantedFuncs
        End If
        vsTmp.Rows = vsTmp.Rows + 1
        vsTmp.TextMatrix(vsTmp.Rows - 1, Col_���) = vsTmp.Rows - 1
        vsTmp.TextMatrix(vsTmp.Rows - 1, Col_����) = rsTmp!���� & ""
        rsTmp.MoveNext
    Loop
    Exit Sub
ErrH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwRole_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If mblnLoad Then Exit Sub
    If mlngRoleIdx = Item.Index Then Exit Sub
    mlngRoleIdx = Item.Index
    mlngModuleIdx = -1
    Call FillModule
End Sub

Private Sub FillModule()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim lst As ListItem
    Dim strRole As String, strPre As String
    
    On Error GoTo ErrH
    lvwModule.ListItems.Clear
    vsfGrantedFuncs.Rows = vsfGrantedFuncs.FixedRows
    vsfUnGrantedFuncs.Rows = vsfUnGrantedFuncs.FixedRows
    '�����б���
    With lvwModule.ColumnHeaders
        .Clear
        If cboSystem.Text = "��������" Then
            .Add , , "�����", "1200"
            .Add , , "����ϵͳ", "2100"
            .Add , , "˵��", "2500"
        ElseIf cboSystem.Text = "ȡ������" Then
            .Add , , "������", "1200"
            .Add , , "������", "1500"
            .Add , , "����ϵͳ", "2100"
            .Add , , "˵��", "2500"
        ElseIf cboSystem.Text = "��������" Then
            .Add , , "���", "600"
            .Add , , "����", "1800"
            .Add , , "˵��", "3000"
        Else
            .Add , , "���", "600"
            .Add , , "����", "1800"
            .Add , , "˵��", "3000"
        End If
    End With
    If Not lvwRole.SelectedItem Is Nothing Then
        strRole = "ZL_" & lvwRole.SelectedItem.Text
    End If
    If strRole = "" Then Exit Sub
    If cboSystem.Text = "��������" Then '��ʾ�ý�ɫ�ܷ��ʵĻ�����
        strSQL = "Select t.ϵͳ, t.����, t.˵��" & vbNewLine & _
                    "From (Select s.���� || '��' || s.��� || '��' As ϵͳ, s.������, b.����, b.˵��" & vbNewLine & _
                    "       From Zlsystems s, Zlbasecode b" & vbNewLine & _
                    "       Where b.ϵͳ = s.���) t, User_Tab_Privs r" & vbNewLine & _
                    "Where t.������ = r.Owner" & vbNewLine & _
                    "And t.���� = r.Table_Name" & vbNewLine & _
                    "And r.Grantee =[1]" & vbNewLine & _
                    "And r.Privilege In ('SELECT', 'INSERT', 'UPDATE', 'DELETE')" & vbNewLine & _
                    "Group By t.ϵͳ, t.����, t.˵��" & vbNewLine & _
                    "Having Count(r.Privilege) = 4"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "FillModule", strRole)
        Do While Not rsTmp.EOF
            Set lst = lvwModule.ListItems.Add(, , rsTmp!����)
            lst.SubItems(1) = rsTmp!ϵͳ
            lst.SubItems(2) = rsTmp!˵�� & ""
            rsTmp.MoveNext
        Loop
    ElseIf cboSystem.Text = "ȡ������" Then '��ʾ�ý�ɫ�ܷ��ʵ�ȡ������
        strSQL = "Select s.���� || '��' || s.��� || '��' As ϵͳ, s.������, f.������, f.������, f.˵��" & vbNewLine & _
                "From Zlsystems s, Zlfunctions f, User_Tab_Privs r" & vbNewLine & _
                "Where f.ϵͳ = s.���" & vbNewLine & _
                "And s.������ = r.Owner" & vbNewLine & _
                "And Upper(f.������) = r.Table_Name" & vbNewLine & _
                "And r.Grantee =[1]" & vbNewLine & _
                "And r.Privilege = 'EXECUTE'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "FillModule", strRole)
        Do While Not rsTmp.EOF
            Set lst = lvwModule.ListItems.Add(, , rsTmp!������)
            lst.SubItems(1) = rsTmp!������
            lst.SubItems(2) = rsTmp!ϵͳ
            lst.SubItems(3) = rsTmp!˵�� & ""
            rsTmp.MoveNext
        Loop
    Else
        If cboSystem.Text = "��������" Then '��ʾ�ý�ɫ�ܷ��ʵĻ�������
            strSQL = "Select p.���, p.����, p.˵��, r.����" & vbNewLine & _
                    "From Zlrolegrant r, Zlprograms p" & vbNewLine & _
                    "Where r.ϵͳ Is Null" & vbNewLine & _
                    "And p.��� = r.���" & vbNewLine & _
                    "And r.��ɫ =[1]" & vbNewLine & _
                    "And p.ϵͳ Is Null" & vbNewLine & _
                    "And p.��� < 100" & vbNewLine & _
                    "Order By p.���"
            
        Else '��ʾ�ý�ɫ�ܷ��ʵ�ģ��
            strSQL = "Select p.���, p.����, p.˵��, r.����" & vbNewLine & _
                    "From Zlrolegrant r, Zlprograms p" & vbNewLine & _
                    "Where Nvl(r.ϵͳ, 0) = Nvl(p.ϵͳ, 0)" & vbNewLine & _
                    "And p.��� = r.���" & vbNewLine & _
                    "And p.��� >= 100" & vbNewLine & _
                    "And r.��ɫ = [1] And " & vbNewLine & _
                    IIF(cboSystem.Text = "�Զ��屨��", " P.ϵͳ is null", " (P.ϵͳ=" & cboSystem.ItemData(cboSystem.ListIndex) & " OR P.��� Between 10000 And 19999)") & vbNewLine & _
                    "Order By p.���"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "FillModule", strRole)
        Do While Not rsTmp.EOF
            If strPre <> rsTmp!��� & "" Then
                Set lst = lvwModule.ListItems.Add(, "K" & rsTmp!���, rsTmp!���)
                lst.SubItems(1) = rsTmp!����
                lst.SubItems(2) = rsTmp!˵�� & ""
                strPre = rsTmp!���
            End If
            rsTmp.MoveNext
        Loop
    End If
    Exit Sub
ErrH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then  '������
        If lvwModule.Width + X > 300 And picSplit.Left + X < Me.Width - 300 Then
            lvwModule.Move lvwModule.Left, lvwModule.Top, lvwModule.Width + X
            picSplit.Left = picSplit.Left + X
            lblGrantedFuncs.Left = picSplit.Left + picSplit.Width + 100
            vsfGrantedFuncs.Left = picSplit.Left + picSplit.Width
            vsfGrantedFuncs.Width = vsfGrantedFuncs.Width - X
            lblUnGrantedFuncs.Left = lblGrantedFuncs.Left
            vsfUnGrantedFuncs.Left = picSplit.Left + picSplit.Width
            vsfUnGrantedFuncs.Width = vsfUnGrantedFuncs.Width - X
        End If
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim litem As ListItem
    Dim lsItem As ListItem
    
    If KeyAscii = vbKeyReturn Then
        For Each litem In lvwRole.ListItems
            If litem.Text = txtFind.Text Or litem.Text = UCase(txtFind.Text) Then
                Set lsItem = litem
            Else
                litem.Selected = False
            End If
        Next
        If Not lsItem Is Nothing Then
            lsItem.Selected = True
            txtFind.SetFocus
            txtFind.SelStart = 0
            txtFind.SelLength = Len(txtFind.Text)
        Else
            MsgBox "δ��ѯ������Ҫ�Ľ�ɫ�����������룡", vbInformation, gstrSysName
            txtFind.SetFocus
            txtFind.SelStart = 0
            txtFind.SelLength = Len(txtFind.Text)
        End If
    End If
End Sub
