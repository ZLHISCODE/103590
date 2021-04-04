VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm参数上传 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数上传"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   Icon            =   "frm参数上传.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ListView lvw方案号 
      Height          =   2610
      Left            =   1110
      TabIndex        =   30
      Top             =   3735
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4604
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ilt16"
      SmallIcons      =   "ilt16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "方案号"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "方案名称"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "上传用户"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "上传站点"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "方案描述"
         Object.Width           =   4304
      EndProperty
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   5055
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5070
      TabIndex        =   26
      Top             =   195
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5070
      TabIndex        =   27
      Top             =   600
      Width           =   1100
   End
   Begin MSComctlLib.ImageList ilt32 
      Left            =   0
      Top             =   0
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
            Picture         =   "frm参数上传.frx":000C
            Key             =   "User"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm参数上传.frx":0326
            Key             =   "Client"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm参数上传.frx":0DF0
            Key             =   "Scheame"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilt16 
      Left            =   2900
      Top             =   1275
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
            Picture         =   "frm参数上传.frx":110A
            Key             =   "User"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm参数上传.frx":1424
            Key             =   "Client"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm参数上传.frx":1EEE
            Key             =   "Scheame"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      Caption         =   "上传设置"
      Height          =   3570
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   4830
      Begin VB.ComboBox cbo方式 
         Height          =   300
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2745
         Width           =   3600
      End
      Begin VB.ComboBox cbo用户 
         Height          =   300
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3120
         Width           =   3600
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   930
         MaxLength       =   18
         TabIndex        =   2
         Tag             =   "方案号"
         Top             =   360
         Width           =   1755
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   930
         MaxLength       =   50
         TabIndex        =   4
         Tag             =   "方案名称"
         Top             =   720
         Width           =   3570
      End
      Begin VB.TextBox txtEdit 
         Height          =   1590
         Index           =   3
         Left            =   930
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Tag             =   "方案描述"
         Top             =   1095
         Width           =   3570
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "…"
         Height          =   285
         Left            =   4470
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   720
         Width           =   300
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "上传方式"
         Height          =   180
         Index           =   3
         Left            =   165
         TabIndex        =   8
         Top             =   2805
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "参数用户"
         Height          =   180
         Index           =   4
         Left            =   165
         TabIndex        =   10
         Top             =   3180
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "方案号"
         Height          =   180
         Index           =   0
         Left            =   345
         TabIndex        =   1
         Top             =   420
         Width           =   540
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "方案名称"
         Height          =   180
         Index           =   1
         Left            =   165
         TabIndex        =   3
         Top             =   780
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "方案描述"
         Height          =   180
         Index           =   2
         Left            =   165
         TabIndex        =   6
         Top             =   1095
         Width           =   720
      End
   End
   Begin VB.Frame fra 
      Caption         =   "注册信息备份与恢复"
      Height          =   3555
      Index           =   2
      Left            =   90
      TabIndex        =   29
      Top             =   135
      Width           =   4815
      Begin VB.CommandButton cmdSearch 
         Caption         =   "…"
         Height          =   330
         Left            =   4335
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1785
         Width           =   300
      End
      Begin VB.CommandButton cmdBakup 
         Caption         =   "备份(&B)"
         Height          =   350
         Left            =   2475
         TabIndex        =   17
         Top             =   2400
         Width           =   1100
      End
      Begin VB.CommandButton cmdRestore 
         Caption         =   "恢复(&R)"
         Height          =   350
         Left            =   3660
         TabIndex        =   18
         Top             =   2400
         Width           =   1100
      End
      Begin VB.TextBox txtFile 
         Height          =   350
         Left            =   825
         MaxLength       =   500
         TabIndex        =   15
         Top             =   1770
         Width           =   3840
      End
      Begin VB.Label lblFile 
         AutoSize        =   -1  'True
         Caption         =   "注册文件"
         Height          =   180
         Left            =   105
         TabIndex        =   14
         Top             =   1860
         Width           =   720
      End
      Begin VB.Label lblReg 
         Caption         =   "    参数备份主要是指将注册表中ZLSOFT下的所有参数设置备份成指定的reg文件,以备恢复。"
         Height          =   450
         Index           =   0
         Left            =   915
         TabIndex        =   12
         Top             =   390
         Width           =   3795
      End
      Begin VB.Image img 
         Height          =   480
         Left            =   165
         Picture         =   "frm参数上传.frx":2208
         Top             =   615
         Width           =   480
      End
      Begin VB.Label lblReg 
         Caption         =   "    参数恢复主要是指将备份的Reg文件进行注册表信息的恢复。"
         Height          =   435
         Index           =   1
         Left            =   885
         TabIndex        =   13
         Top             =   855
         Width           =   3795
      End
   End
   Begin VB.Frame fra 
      Caption         =   "参数恢复"
      Height          =   3555
      Index           =   1
      Left            =   90
      TabIndex        =   19
      Top             =   150
      Width           =   4830
      Begin VB.Frame Frame3 
         Caption         =   "用户选择"
         Height          =   2220
         Left            =   105
         TabIndex        =   28
         Top             =   690
         Width           =   4575
         Begin VB.CheckBox chkAllUser 
            Caption         =   "所有用户"
            Height          =   240
            Left            =   3390
            TabIndex        =   22
            Top             =   0
            Width           =   1035
         End
         Begin MSComctlLib.ListView lvw用户 
            Height          =   1935
            Left            =   90
            TabIndex        =   23
            Top             =   225
            Width           =   4425
            _ExtentX        =   7805
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "ilt32"
            SmallIcons      =   "ilt16"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "用户名"
               Object.Width           =   4304
            EndProperty
         End
      End
      Begin VB.ComboBox cbo下载 
         Height          =   300
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   300
         Width           =   3720
      End
      Begin VB.Label lblInforLst 
         Caption         =   "方案信息：上传站点[lxh],上传用户(zlhis)"
         ForeColor       =   &H80000001&
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   3270
         Width           =   4680
      End
      Begin VB.Label lblInfor 
         Caption         =   "恢复方案：[123456789]lxh配置新方案"
         ForeColor       =   &H80000001&
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   3030
         Width           =   4680
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "恢复方式"
         Height          =   180
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   720
      End
   End
End
Attribute VB_Name = "frm参数上传"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrPrvs As String      '权限串
Dim mblnFirst As Boolean
Dim mbytParaType As Byte    '0-上传,1-下载,2-本地注册表备份与恢复

Public Sub ShowEdit(ByVal frmMain As Object, ByVal bytParaType As Byte)
    '----------------------------------------------------------------------------------------------------------------
    '--功能:显示上传窗体.
    '--参数:frmMain-父窗口
    '       bytParaType-参数类型(0-上传,1-下载)
    '       strPrivs-权限串
    '----------------------------------------------------------------------------------------------------------------
    mstrPrvs = GetPrivFunc(0, 工具清单.本地参数管理)
    mbytParaType = bytParaType
    
    Me.Show vbModal, frmMain
End Sub
Private Sub setCtlShowMode()
    '功能:设置控件的显示模式
    fra(0).Visible = False
    fra(1).Visible = False
    fra(2).Visible = False
    If mbytParaType = 0 Then
        fra(0).Visible = True
    ElseIf mbytParaType = 1 Then
        fra(1).Visible = True
        cmdSave.Enabled = True
    Else
        fra(2).Visible = True
        cmdSave.Enabled = True
    End If
End Sub
Private Sub cbo方式_Click()
    Dim bytType As Byte, i As Integer
    If mbytParaType <> 0 Then Exit Sub
    
    '初始化参数
    bytType = Val(Split(Me.cbo方式.Text, "-")(0))
    Select Case bytType
    Case 1          '公用
        Me.cbo用户.Enabled = False      '不用选用户
    Case Else         '所有,私有
        '要指定用户
        If InStr(1, mstrPrvs, "参数上传") <> 0 Then
            Me.cbo用户.Enabled = True
        Else
            '指定是当前用户
            Me.cbo用户.ListIndex = -1
            For i = 0 To Me.cbo用户.ListCount - 1
                If Me.cbo用户.List(i) = UCase(gstrDbUser) Then
                    Me.cbo用户.ListIndex = i
                    Exit For
                End If
            Next
            Me.cbo用户.Enabled = False
        End If
    End Select
End Sub

Private Sub cbo方式_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If

End Sub

Private Sub cbo下载_Click()
  Dim bytType As Byte, i As Integer
  If mbytParaType = 0 Then Exit Sub
    '初始化参数
    bytType = Val(Split(Me.cbo下载.Text, "-")(0))
    Select Case bytType
    Case 1          '公用
        Me.lvw用户.Enabled = False      '不用选用户
        Me.lvw用户.BackColor = Me.BackColor
        Me.chkAllUser.Enabled = False
        
    Case Else         '所有,私有
        '要指定用户
        If InStr(1, mstrPrvs, "参数下载") <> 0 Then
            Me.lvw用户.Enabled = True
            Me.lvw用户.BackColor = Me.cbo下载.BackColor
            Me.chkAllUser.Enabled = True
        Else
            '指定是当前用户
            lvw用户.Enabled = False
            Me.lvw用户.BackColor = Me.BackColor
            Me.chkAllUser.Enabled = False
        End If
    End Select
End Sub
Private Sub cbo下载_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo用户_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chkAllUser_Click()
    Dim lstItem As ListItem
    Dim blnCheck As Boolean
    If chkAllUser.Value = 2 Then Exit Sub
    blnCheck = chkAllUser.Value = 1
    For Each lstItem In lvw用户.ListItems
        lstItem.Checked = blnCheck
    Next
End Sub
Private Sub chkAllUser_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdBakup_Click()
    '备份
    Dim strFile As String
    Dim strPath As String
    Dim strPathAndFile As String
    Dim objFile As New FileSystemObject
    
    If Trim(txtFile.Text) = "" Then
        MsgBox "请选择或输入要备份的文件!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    strPathAndFile = Trim(txtFile.Text)
    strFile = objFile.GetFileName(strPathAndFile)
    
    strPath = objFile.GetParentFolderName(strPathAndFile)
    If strFile = "" Then
        MsgBox "不存在文件名,请重输!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    If FindFile(strPath) = False Then
        MsgBox "不存在该路径,请重新设置文件路径!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    
    '备份
    If ExportReg(strPathAndFile, """HKEY_CURRENT_USER\SOFTWARE\VB AND VBA PROGRAM SETTINGS\ZLSOFT""") = False Then
        MsgBox "导出失败!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    MsgBox "导出成功!", vbInformation + vbDefaultButton1, gstrSysName
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRestore_Click()
    '备份
    Dim strFile As String
    Dim strPath As String
    Dim strPathAndFile As String
    Dim objFile As New FileSystemObject
    
    If Trim(txtFile.Text) = "" Then
        MsgBox "请选择或输入要恢复的文件!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    strPathAndFile = Trim(txtFile.Text)
    strFile = objFile.GetFileName(strPathAndFile)
    strPath = objFile.GetParentFolderName(strPathAndFile)
    If strFile = "" Then
        MsgBox "不存文件名,请重输!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    If FindFile(strPath) = False Then
        MsgBox "不存在该路径,请重新设置文件路径!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    
    If FindFile(strPathAndFile) = False Then
        MsgBox "不存在该注册文件,请重新设置文件!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    
    '备份
    If RestoreWndowsReg(strPathAndFile) = False Then
        MsgBox "导入失败!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    MsgBox "导入成功!", vbInformation + vbDefaultButton1, gstrSysName
'    Unload Me
End Sub

Private Sub cmdSave_Click()
   
    If mbytParaType = 0 Then
        '上传
        If SaveUPData = False Then Exit Sub
        Unload Me
        Exit Sub
    ElseIf mbytParaType = 1 Then
        '下载
        If DataBaseToRestoreReg = False Then Exit Sub
        Unload Me
        Exit Sub
    Else
        Unload Me
        Exit Sub
    End If

End Sub
Private Function DataBaseToRestoreReg() As Boolean
    '功能:从设置的参数中恢复到本机注册表中
        Dim bytType As Byte, lng方案号 As Long
        Dim cllUser As New Collection, lstItem As ListItem
        DataBaseToRestoreReg = False
        Err = 0: On Error GoTo ErrHand:
        
        bytType = Val(Split(cbo下载.Text, "-")(0))
        For Each lstItem In lvw用户.ListItems
            If lstItem.Checked Then
                cllUser.Add lstItem.Text
            End If
        Next
        If cllUser.Count = 0 And (bytType = 0 Or bytType = 2) Then
            MsgBox "未选择恢复的用户名,不能继续!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
        lng方案号 = Val(lblInfor.Tag)
        If RegRestore(bytType, lng方案号, cllUser) = False Then
            Exit Function
        End If
        MsgBox "恢复成功!", vbInformation + vbDefaultButton1, gstrSysName
        DataBaseToRestoreReg = True
        Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function SaveUPData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------
    '功能:上传数据保存
    '参数:
    '返回:成功返回true,否则返回false
    '---------------------------------------------------------------------------------------------------------------------
    SaveUPData = False
   '判断相关设置值是否正确
    Dim bytType As Byte
    Dim strSQL As String
    Dim strTemp As String
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    For i = 1 To txtEdit.UBound
        txtEdit(i).Text = Trim(txtEdit(i).Text)
        strTemp = txtEdit(i).Text
        If i <= 2 Then
            If strTemp = "" Then
                MsgBox txtEdit(i).Tag & "必需输入!", vbInformation + vbDefaultButton1, gstrSysName
                If txtEdit(i).Enabled Then txtEdit(i).SetFocus
                Exit Function
            End If
        End If
        If zlCommFun.ActualLen(strTemp) > txtEdit(i).MaxLength Then
            MsgBox txtEdit(i).Tag & "输入过长,必需小于等于" & txtEdit(i).MaxLength & "个符或" & Int(txtEdit(i).MaxLength / 2) & "个汉字!", vbInformation + vbDefaultButton1, gstrSysName
            If txtEdit(i).Enabled Then txtEdit(i).SetFocus
            Exit Function
        End If
        If InStr(1, strTemp, "'") > 0 Then
            MsgBox txtEdit(i).Tag & "输入了非法字符（'）!", vbInformation + vbDefaultButton1, gstrSysName
            If txtEdit(i).Enabled Then txtEdit(i).SetFocus
            Exit Function
        End If
    Next
    bytType = Val(Split(cbo方式.Text, "-")(0))
    If bytType = 0 Or bytType = 2 Then
        '需要指定相关用户
        If cbo用户.Text = "" Or cbo用户.ListIndex < 0 Then
            MsgBox "未指定上传用户，请重新指定!", vbInformation + vbDefaultButton1, gstrSysName
            If cbo用户.Enabled Then cbo用户.SetFocus
            Exit Function
        End If
        If InStr(1, mstrPrvs, "参数上传") = 0 Then
            If cbo用户.Text <> UCase(gstrDbUser) Then
                MsgBox "你没有权限上传" & cbo用户.Text & "用户的" & vbCrLf & "相关私有参数，请重新指定!", vbInformation + vbDefaultButton1, gstrSysName
                If cbo用户.Enabled Then cbo用户.SetFocus
                Exit Function
            End If
        End If
    End If
    strSQL = "Select 方案号 ,用户名 from zlClientScheme where 方案号=[1]"
    Set rsTemp = zlDatabase.OpensqlRecord(strSQL, Me.Caption, Val(txtEdit(1).Text))
    If Not rsTemp.EOF Then
                
        '判断用户名是否相同
        If zlCommFun.Nvl(rsTemp!用户名) <> UCase(gstrDbUser) And InStr(1, mstrPrvs, "参数上传") = 0 Then
            '该参数号已经被他人设置，系统将默认一下个新号!"
            MsgBox "该参数号已经被他人设置，系统将默认一个新号!", vbInformation + vbDefaultButton1, gstrSysName
            txtEdit(1).Text = Max方案号
            If txtEdit(1).Enabled Then txtEdit(1).SetFocus
            Exit Function
        Else
            If MsgBox("该方案号已经存在,是否覆盖该方案?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                txtEdit(1).Text = Max方案号
                If txtEdit(1).Enabled Then txtEdit(1).SetFocus
                Exit Function
            End If
        End If
    End If
    
    Dim cllData As New Collection
    If ExportParasToCollection(bytType, UCase(cbo用户.Text), cllData) = False Then Exit Function
    If cllData Is Nothing Then
        MsgBox "无满足条件的参数信息,请检查!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    If cllData.Count = 0 Then
        MsgBox "无满足条件的参数信息,请检查!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    
    Err = 0: On Error GoTo ErrHand:
    Call SaveClientParaToDataBase(cllData)
    MsgBox "上传成功!", vbInformation + vbDefaultButton1, gstrSysName
    SaveUPData = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Function
Private Sub SaveClientParaToDataBase(ByVal cllData As Collection)
    '--------------------------------------------------------------------------------------------------------------------------------------------
    '功能:将客户端参数保存到数据库中
    '参数:cllData-参数集
    '--------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    Dim lng参数号 As Long
    Dim str参数名 As String
    Dim str参数说明 As String
    Dim strKey As String
    Dim strTemp As String
    Dim strData As String
    lng参数号 = Val(txtEdit(1).Text)
    str参数名 = Trim(txtEdit(2).Text)
    str参数说明 = Trim(txtEdit(3).Text)
    Dim rsTemp As New Recordset
    Dim lng目录 As Long, lng键名 As Long, lng键值 As Long

    strSQL = "select 目录,键名,键值 from zlClientparaList where rownum<1 "
    zlDatabase.OpenRecordset rsTemp, strSQL, "取字段长度"
    lng目录 = rsTemp.Fields("目录").DefinedSize
    lng键名 = rsTemp.Fields("键名").DefinedSize
    lng键值 = rsTemp.Fields("键值").DefinedSize

    gcnOracle.BeginTrans
    '先删除该方案
    strSQL = "delete zlClientScheme where 方案号=" & lng参数号
    gcnOracle.Execute strSQL
    
    strSQL = " insert into zlClientScheme(方案号,方案名称, 方案描述,工作站,用户名) values(" & lng参数号 & ","
    strSQL = strSQL & "'" & str参数名 & "',"
    strSQL = strSQL & "'" & str参数说明 & "',"
    strSQL = strSQL & "'" & AnalyseComputer & "',"
    strSQL = strSQL & "'" & UCase(gstrDbUser) & "')"
    gcnOracle.Execute strSQL
    For i = 1 To cllData.Count
        '加上传参数
        strSQL = "insert into zlClientparaList(方案号,序号,类别,目录,键名,键值,参数来源,参数说明) values ("
        strSQL = strSQL & "" & lng参数号 & ","
        strSQL = strSQL & "" & i & ","
        strKey = cllData(i)(0)
        '    If InStr(1, strSection, "私有模块") > 0 Then
        '    cllData.Add Array(strSection, strKey, strData), "S0" & i
        '    ElseIf InStr(1, strSection, "私有全局") > 0 Then
        '    cllData.Add Array(strSection, strKey, strData), "S1" & i
        '    ElseIf InStr(1, strSection, "公共模块") > 0 Then
        '    cllData.Add Array(strSection, strKey, strData), "G0" & i
        '    ElseIf InStr(1, strSection, "公共全局") > 0 Then
        '    cllData.Add Array(strSection, strKey, strData), "G1" & i
        '    End If
        If strKey = "私有模块" Then
            strSQL = strSQL & "'私有模块',"
            strTemp = Replace(cllData(i)(1), "私有模块\" & UCase(Trim(cbo用户.Text)) & "\", "")
            strTemp = Replace(strTemp, "私有模块\" & UCase(Trim(cbo用户.Text)), "")
        ElseIf strKey = "私有全局" Then
            strSQL = strSQL & "'私有全局',"
            strTemp = Replace(cllData(i)(1), "私有全局\" & UCase(Trim(cbo用户.Text)) & "\", "")
            strTemp = Replace(strTemp, "私有全局\" & UCase(Trim(cbo用户.Text)), "")
        ElseIf strKey = "公共模块" Then
            strSQL = strSQL & "'公共模块',"
            strTemp = Replace(cllData(i)(1), "公共模块\", "")
            strTemp = Replace(strTemp, "公共模块", "")
        Else
            strSQL = strSQL & "'公共全局',"
            strTemp = Replace(cllData(i)(1), "公共全局\", "")
            strTemp = Replace(strTemp, "公共全局", "")
        End If
        
        strSQL = strSQL & "'" & strTemp & "',"
        strData = Replace(cllData(i)(3), "'", "''")             '键值
        strData = Replace(strData, "\\", "\")
              
        strSQL = strSQL & "'" & cllData(i)(2) & "',"
        strSQL = strSQL & "'" & strData & "',"
        strSQL = strSQL & "0,NULL)"
        If zlCommFun.ActualLen(strTemp) > lng目录 Or zlCommFun.ActualLen(cllData(i)(2)) > lng键名 Or zlCommFun.ActualLen(strData) > lng键值 Then
            '大于数据库的存储范围,就不保存了
            Debug.Print "fds"
        Else
            gcnOracle.Execute strSQL
        End If
    Next
    gcnOracle.CommitTrans
End Sub

Private Sub initData()
    '--------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '--------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset

    
    '--初始化上传方式
    Me.cbo方式.Clear
    Me.cbo方式.AddItem "0-所有参数(公共+私有)"
    Me.cbo方式.AddItem "1-公共参数(公共全局+公共模块)"
    Me.cbo方式.AddItem "2-私有参数(私有全局+私有模块)"
    Me.cbo方式.ListIndex = 0
    Me.cbo下载.Clear
    Me.cbo下载.AddItem "0-所有参数(公共+私有)"
    Me.cbo下载.AddItem "1-公共参数(公共全局+公共模块)"
    Me.cbo下载.AddItem "2-私有参数(私有全局+私有模块)"
    Me.cbo下载.ListIndex = 0
    
    If mbytParaType = 0 Then
        '初始化上传的用户名
        strSQL = "Select distinct upper(用户名) 用户名 From 上机人员表 "
        zlDatabase.OpenRecordset rsTemp, strSQL, Me.Caption
        With rsTemp
            Me.cbo用户.Clear
            Do While Not .EOF
                Me.cbo用户.AddItem zlCommFun.Nvl(!用户名)
                If zlCommFun.Nvl(!用户名) = gstrDbUser Then
                    Me.cbo用户.ListIndex = Me.cbo用户.NewIndex
                End If
                .MoveNext
            Loop
            If Me.cbo用户.ListCount = 0 Or Me.cbo用户.ListIndex < 0 Then
                If gstrDbUser <> "" Then
                    '加入当前用户
                    Me.cbo用户.AddItem gstrDbUser
                    Me.cbo用户.ListIndex = Me.cbo用户.NewIndex
                End If
            End If
        End With
        
        '查找出当前用户的最后一次制定的方案号
        strSQL = "Select 方案号,方案名称, 方案描述,工作站 from zlClientScheme where 方案号 =(Select max(方案号) from zlClientScheme where 用户名 =[1])"
        Set rsTemp = zlDatabase.OpensqlRecord(strSQL, Me.Caption, gstrDbUser)
        If Not rsTemp.EOF Then
            '加载数据
            txtEdit(1).Text = zlCommFun.Nvl(rsTemp!方案号)
            txtEdit(2).Text = zlCommFun.Nvl(rsTemp!方案名称)
            txtEdit(2).Tag = zlCommFun.Nvl(rsTemp!方案号)
            txtEdit(3).Text = zlCommFun.Nvl(rsTemp!方案描述)
        Else
            txtEdit(1).Text = Max方案号
        End If
        Exit Sub
    End If
    If mbytParaType = 2 Then
        Exit Sub
    End If
    lvw用户.ListItems.Clear
    Dim strComputerName As String
    
    strComputerName = AnalyseComputer
    strSQL = "Select distinct a.方案号,a.工作站  ,a.用户名,b.方案名称,b.方案描述,b.工作站 as 上传站点,b.用户名 as 上传用户名" & _
            " From Zlclientparaset a,zlClientScheme b" & _
            " where a.方案号=b.方案号 and (a.工作站=[1] or (a.工作站 is null and a.用户名 is not null))"
            
    Set rsTemp = zlDatabase.OpensqlRecord(strSQL, Me.Caption, strComputerName)
    If rsTemp.EOF Then
        MsgBox "本站点没设置任何需要恢复的参数方案，" & vbCrLf & "不能进行恢复[服务器管理工具中的“参数配置”]！", vbInformation + vbDefaultButton1, gstrSysName
        Unload Me
        Exit Sub
    End If
    Dim lst As ListItem
    Dim bln公共 As Boolean '存在共公参数
    bln公共 = False
    lblInfor.Caption = "恢复方案：[" & zlCommFun.Nvl(rsTemp!方案号) & "]" & zlCommFun.Nvl(rsTemp!方案名称)
    lblInforLst.Caption = "方案信息：上传站点[" & zlCommFun.Nvl(rsTemp!上传站点) & "],上传用户[" & zlCommFun.Nvl(rsTemp!上传用户名) & "]"
    lblInfor.Tag = zlCommFun.Nvl(rsTemp!方案号)
    Err = 0: On Error Resume Next
    With rsTemp
        Do While Not .EOF
            If Not IsNull(!工作站) And IsNull(!用户名) Then
                '存在公用部分
                bln公共 = True
            End If
            If zlCommFun.Nvl(rsTemp!用户名) <> "" Then
                If InStr(1, mstrPrvs, "参数下载") <> 0 Then
                    Set lst = lvw用户.ListItems.Add(, "K" & zlCommFun.Nvl(rsTemp!用户名), zlCommFun.Nvl(rsTemp!用户名), "User", "User")
                ElseIf zlCommFun.Nvl(rsTemp!用户名) = UCase(gstrDbUser) Then
                    Set lst = lvw用户.ListItems.Add(, "K" & zlCommFun.Nvl(rsTemp!用户名), zlCommFun.Nvl(rsTemp!用户名), "User", "User")
                End If
                lst.Checked = True
            End If
            .MoveNext
        Loop
    End With
    If bln公共 = False And lvw用户.ListItems.Count = 0 Then
        MsgBox "本站点不存在任何恢复参数方案，不能进行恢复！", vbInformation + vbDefaultButton1, gstrSysName
        Unload Me
        Exit Sub
    End If
    Dim i As Long
    For i = 0 To cbo下载.ListCount - 1
        Select Case Split(cbo下载.List(i), "-")(0)
        Case 0 '所有
            If bln公共 = False Or lvw用户.ListItems.Count = 0 Then
                cbo下载.RemoveItem i
            End If
        Case 1 '共公
            If bln公共 = False Then
                cbo下载.RemoveItem i
            End If
        Case 2 '私有
            If lvw用户.ListItems.Count = 0 Then
                cbo下载.RemoveItem i
            End If
        End Select
    Next
    If cbo下载.ListCount = 0 Then
        MsgBox "本站点不存在任何恢复参数方案，不能进行恢复！", vbInformation + vbDefaultButton1, gstrSysName
        Unload Me
        Exit Sub
    End If
    cbo下载.ListIndex = 0
End Sub
Private Function Max方案号() As Long
    '功能:获取最大参数号
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    strSQL = "Select nvl(Max(方案号),0)+1 as 方案号 from zlClientScheme"
    zlDatabase.OpenRecordset rsTemp, strSQL, Me.Caption
    If rsTemp.EOF Then
        Max方案号 = 1
    Else
        Max方案号 = Val(zlCommFun.Nvl(rsTemp!方案号))
    End If
End Function

Private Sub cmdSearch_Click()
    Dim strFile As String
    
    Err = 0
    On Error Resume Next
    With Dlg
        .Filter = "注册文件(*.reg)|*.reg"
        .Flags = cdlOFNFileMustExist Or cdlOFNLongNames
        .ShowOpen
        If Err <> 0 Then Exit Sub
        strFile = .FileName
    End With
    Err = 0
    txtFile.Text = strFile
End Sub

Private Sub Command3_Click()

End Sub

Private Sub cmdSel_Click()
        Call SelectScreme("")
End Sub
Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    '初始化参数
    Call initData
    Call setCtlShowMode
    If mbytParaType = 0 Then
        Me.Caption = "参数上传"
        If Me.txtEdit(1).Enabled And Me.txtEdit(1).Visible Then Me.txtEdit(1).SetFocus
    ElseIf mbytParaType = 1 Then
        Me.Caption = "参数恢复"
        If Me.cbo下载.Enabled And Me.cbo下载.Visible Then Me.cbo下载.SetFocus
    Else
        Me.Caption = "本地注册信息备份与恢复"
        Me.cmdSave.Caption = "退出(&X)"
        Me.cmdCancel.Visible = False
        If Me.txtFile.Enabled And Me.txtFile.Visible Then Me.txtFile.SetFocus
    End If
End Sub

Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Sub lvw用户_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    
    Me.chkAllUser.Value = 2
End Sub

Private Sub lvw用户_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtEdit_Change(Index As Integer)
    Call SetSaveCtlEnable
    If Index = 2 Then
        txtEdit(Index).Tag = ""
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
'    If Index <> 1 Then
'        '打开输入法
'        ImeLanguage True
'    End If
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Index = 3 Then
            '取掉回车符
            KeyCode = 0
        Else
            If Index = 2 And Trim(txtEdit(Index)) <> "" And Trim(txtEdit(2).Tag) = "" Then
                Call SelectScreme(Trim(txtEdit(Index)))
            Else
                zlCommFun.PressKey vbKeyTab
            End If
        End If
    End If
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m数字式
    Else
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m文本式
    End If
    If KeyAscii = vbKeyReturn And Index = 3 Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
Private Sub SetSaveCtlEnable()
    '功能:设置Save控件的Enable属性
    If mbytParaType = 0 Then
        Me.cmdSave.Enabled = Trim(txtEdit(1).Text) <> "" And Trim(txtEdit(2).Text) <> ""
    Else
        Me.cmdSave.Enabled = True
    End If
End Sub

Private Sub txtFile_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub SelectScreme(ByVal strKey As String)
    '功能:选择方案
    '参数:  strKey -过滤条件
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim lstItem As ListItem
    Err = 0: On Error GoTo ErrHand:
    strSQL = "" & _
        "   Select  方案号,方案名称,方案描述,工作站 as 上传站点,用户名 as 上传用户名" & _
        " From zlClientScheme"
    If InStr(1, mstrPrvs, "参数上传") = 0 Then
            strSQL = strSQL & " where 用户名='" & UCase(gstrDbUser) & "'"
            If strKey <> "" Then
                strSQL = strSQL & " and ( 方案号 like '" & strKey & "%' or 方案名称 like '" & strKey & "%' or 用户名 like '" & strKey & "%')"
            End If
    Else
        If strKey <> "" Then
            strSQL = strSQL & " where 方案号 like '" & strKey & "%' or 方案名称 like '" & strKey & "%' or 用户名 like '" & strKey & "%'"
        End If
    End If
    strSQL = strSQL & " order by 方案号"
    zlDatabase.OpenRecordset rsTemp, strSQL, Me.Caption
    With rsTemp
        If .RecordCount = 0 Then
            MsgBox "不存在符合条件的方案!", vbInformation + vbDefaultButton1, gstrSysName
            If txtEdit(2).Enabled Then txtEdit(2).SetFocus
            Exit Sub
        End If
        If .RecordCount = 1 Then
            txtEdit(1).Text = zlCommFun.Nvl(!方案号)
            txtEdit(2).Text = zlCommFun.Nvl(!方案名称)
            txtEdit(2).Tag = zlCommFun.Nvl(!方案号)
            txtEdit(3).Text = zlCommFun.Nvl(!方案描述)
            If cbo方式.Enabled And fra(0).Visible Then cbo方式.SetFocus
            Exit Sub
        End If
        Me.lvw方案号.ListItems.Clear
        Do While Not .EOF
            Set lstItem = lvw方案号.ListItems.Add(, "K" & zlCommFun.Nvl(!方案号), zlCommFun.Nvl(!方案号), "Scheame", "Scheame")
            lstItem.SubItems(1) = zlCommFun.Nvl(!方案名称)
            lstItem.SubItems(2) = zlCommFun.Nvl(!上传用户名)
            lstItem.SubItems(3) = zlCommFun.Nvl(!上传站点)
            lstItem.SubItems(4) = zlCommFun.Nvl(!方案描述)
            If .AbsolutePosition = 1 Then lstItem.Selected = True
            .MoveNext
        Loop
    End With
    With lvw方案号
        .Top = fra(0).Top + txtEdit(2).Top + txtEdit(2).Height
        .Left = fra(0).Left + txtEdit(2).Left
        .Visible = True
        .SetFocus
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub lvw方案号_DblClick()
    Call lvw方案号_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub lvw方案号_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If lvw方案号.ListItems.Count = 0 Then Exit Sub
    If lvw方案号.SelectedItem Is Nothing Then Exit Sub
    
    txtEdit(1).Text = lvw方案号.SelectedItem.Text
    txtEdit(2).Text = lvw方案号.SelectedItem.SubItems(1)
    txtEdit(2).Tag = lvw方案号.SelectedItem.Text
    txtEdit(3).Text = lvw方案号.SelectedItem.SubItems(4)
    
    lvw方案号.Visible = False
    If cbo方式.Enabled And fra(0).Visible Then cbo方式.SetFocus
End Sub
Private Sub lvw方案号_LostFocus()
    lvw方案号.Visible = False
End Sub


