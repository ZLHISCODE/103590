VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUser 
   BackColor       =   &H80000005&
   Caption         =   "用户授权管理"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8550
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmUser.frx":0000
   ScaleHeight     =   5940
   ScaleWidth      =   8550
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picHLine 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   945
      MousePointer    =   7  'Size N S
      ScaleHeight     =   90
      ScaleWidth      =   5835
      TabIndex        =   18
      Top             =   3420
      Width           =   5835
   End
   Begin VB.CommandButton cmdWhole 
      Caption         =   "批量创建用户"
      Enabled         =   0   'False
      Height          =   350
      Index           =   0
      Left            =   6015
      TabIndex        =   17
      Top             =   1350
      Width           =   1440
   End
   Begin VB.CommandButton cmdUpdatePWD 
      Caption         =   "修改密码(&P)"
      Height          =   350
      Left            =   6015
      TabIndex        =   16
      Top             =   3030
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picSel 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5625
      ScaleHeight     =   285
      ScaleWidth      =   1065
      TabIndex        =   15
      Top             =   968
      Width           =   1065
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   6870
      TabIndex        =   14
      Top             =   960
      Width           =   2610
   End
   Begin VB.CommandButton cmdUnDoLock 
      Caption         =   "用户解锁(&J)"
      Height          =   350
      Left            =   6015
      TabIndex        =   13
      Top             =   2700
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除用户(&D)"
      Height          =   350
      Left            =   6015
      TabIndex        =   11
      Top             =   2370
      Width           =   1440
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "修改用户(&M)"
      Height          =   350
      Left            =   6015
      TabIndex        =   10
      Top             =   2040
      Width           =   1440
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "新增用户(&A)"
      Height          =   350
      Left            =   6015
      TabIndex        =   12
      Top             =   1695
      Width           =   1440
   End
   Begin VB.Frame fraFuncs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   915
      TabIndex        =   6
      Top             =   5055
      Width           =   6990
      Begin VB.CommandButton cmdWhole 
         Caption         =   "恢复所有用户角色"
         Height          =   350
         Index           =   3
         Left            =   2160
         TabIndex        =   9
         Top             =   0
         Width           =   2160
      End
      Begin VB.CommandButton cmdWhole 
         Caption         =   "重整所有用户角色"
         Height          =   350
         Index           =   2
         Left            =   4320
         TabIndex        =   8
         Top             =   0
         Width           =   2160
      End
      Begin VB.CommandButton cmdWhole 
         Caption         =   "根据上机人员恢复用户"
         Enabled         =   0   'False
         Height          =   350
         Index           =   1
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   2160
      End
   End
   Begin MSComctlLib.ImageList Img大图标 
      Left            =   165
      Top             =   2490
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":04F9
            Key             =   "User"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":228B
            Key             =   "Role_Dba"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":480D
            Key             =   "Role_User"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":6D8F
            Key             =   "Role"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":7A69
            Key             =   "User1"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":E2CB
            Key             =   "UserInfor"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":14B2D
            Key             =   "UserLock"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Img小图标 
      Left            =   135
      Top             =   1710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":1B38F
            Key             =   "User"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":1D121
            Key             =   "Role_Dba"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":1F6A3
            Key             =   "Role_User"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":21C25
            Key             =   "Role"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":228FF
            Key             =   "User1"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":29161
            Key             =   "UserInfor"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":2F9C3
            Key             =   "UserLock"
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboSystem 
      Height          =   300
      Left            =   1710
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   3645
   End
   Begin MSComctlLib.ListView lvwRole 
      Height          =   1185
      Left            =   945
      TabIndex        =   2
      Top             =   3750
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   2090
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "Img大图标"
      SmallIcons      =   "Img小图标"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "标记"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwUser 
      Height          =   2070
      Left            =   945
      TabIndex        =   3
      Top             =   1320
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   3651
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "Img大图标"
      SmallIcons      =   "Img小图标"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Settlement"
         Text            =   "用户名"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "人员编号"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "人员姓名"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "所属部门"
         Object.Width           =   3087
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "用户状态"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   240
      Picture         =   "frmUser.frx":36225
      Top             =   690
      Width           =   480
   End
   Begin VB.Label lblRole 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "授权角色"
      Height          =   180
      Left            =   945
      TabIndex        =   5
      Top             =   3525
      Width           =   720
   End
   Begin VB.Label lblSys 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "应用系统"
      Height          =   180
      Left            =   945
      TabIndex        =   4
      Top             =   1020
      Width           =   720
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户授权管理"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   0
      Top             =   105
      Width           =   1440
   End
   Begin VB.Menu mnuPopuMenu 
      Caption         =   "弹出菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuMenuSearch 
         Caption         =   "按部门过滤(&B)"
         Index           =   0
      End
      Begin VB.Menu mnuPopuMenuSearch 
         Caption         =   "按用户过滤(&U)"
         Index           =   1
      End
      Begin VB.Menu mnuPopuMenuSearch 
         Caption         =   "按人员过滤(&P)"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==============================================================
'==模块变量
'==============================================================
'过滤菜单枚举
Private Enum menuEnum
    ME_部门 = 0
    ME_用户 = 1
    ME_人员 = 2
End Enum
'用户列表列，从1开始，第0列没有
Private Enum UserCol
    Col_人员编号 = 1
    Col_人员姓名 = 2
    Col_所属部门 = 3
    Col_用户状态 = 4
End Enum

Private Enum WholeEnum
    WE_CreateAllUser = 0 '所有人员设为用户
    WE_RestoreAllUser = 1 '恢复所有上机人员
    WE_RecUserRoles = 2 '记录所有用户角色
    WE_RestoreUserRoles = 3 '恢复所有用户角色
End Enum
Private mrsSystem As New ADODB.Recordset
Private mstrBakOwner As String '所有系统历史库所有者字符串
Private mstrAllSysOwner As String '所有系统所有者
Private mstr所有者 As String '保存当前系统的所有者名
Private mintColumn As Integer '

Private mbytSearch As Byte      '0-按所属部搜索,1-按用户搜索,2-按人员搜索
Private mrsUsers As ADODB.Recordset
Private mLastIndex As Long '上次选中的用户

Private mobjTip  As clsTipSwap           '悬浮提示框对象

'==============================================================
'==公共接口
'==============================================================
Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = True
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objPrint As zlPrintLvw
    Dim rsTmp As ADODB.Recordset
    
    Set objPrint = New zlPrintLvw
    objPrint.Title.Text = "系统用户"
    Set objPrint.Body.objData = lvwUser
    objPrint.UnderAppItems.Add "应用系统：" & cboSystem.Text
    objPrint.BelowAppItems.Add "打印时间：" & Format(CurrentDate, "yyyy年MM月dd日")
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
'==============================================================
'==控件事件
'==============================================================
Private Sub cboSystem_Click()
    Call FillUser
End Sub

Private Sub cmdAdd_Click()
'新增用户
    Dim blnSucced As Boolean
    If frmUserEdit.UserEdit(mstr所有者) Then
        Set mrsUsers = Nothing
        Call cboSystem_Click
    End If
End Sub

Private Sub CmdDelete_Click()
'删除相应用户
    Dim strUser As String, intIndex As Integer
    Dim strRemarks As String
        
    If gblnMustRIS And Not gblnRIS And UCase(gstrSTOwner) = UCase(mstr所有者) Then
        MsgBox "RIS接口创建失败，不能继续当前操作。可能是接口文件安装或注册不正常，请与系统管理员联系。", vbInformation, gstrSysName
        Exit Sub
    End If
    strUser = lvwUser.SelectedItem.Text
    intIndex = lvwUser.SelectedItem.Index
    If UCase(strUser) = "ZLYB" Then
        MsgBox "这是一些特殊用户，不能使用本程序删除。", vbInformation, gstrSysName
        Exit Sub
    End If
    If UCase(strUser) = "ZLDOC" Then
        MsgBox "这是资料文档定义的用户，不能使用本程序删除。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mrsSystem.Filter = "所有者='" & strUser & "'"
    If mrsSystem.RecordCount > 0 Then
        MsgBox "用户" & strUser & "是系统《" & mrsSystem("名称") & "》的所有者，不能删除。" & _
            vbCrLf & "如果你确实要删除该用户，请使用装卸管理程序。", vbExclamation, gstrSysName
        Exit Sub
    End If
    If MsgBox("你确实要删除用户" & strUser & "吗？" & vbCrLf & _
        "这会把该用户下的所有数据库对象一并删除。", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    If Mid(lvwUser.SelectedItem.Tag, 3) = "" Then
        If MsgBox("该用户可能不是你创建的,你确实要删除用户" & strUser & "吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    '验证身份并输入操作说明
    strRemarks = "删除用户：" & strUser
    If Not CheckAuditStatus("0402", "删除用户", strRemarks) Then Exit Sub
    On Error Resume Next
    MousePointer = 11
    DoEvents
    gcnOracle.Execute "drop user " & strUser & " cascade"
    If err.Number <> 0 Then
        MsgBox "该用户可能不是由你创建的，删除失败。", vbExclamation, gstrSysName
        err.Clear: MousePointer = 0
        Exit Sub
    End If
    gcnOracle.Execute "delete from " & mstr所有者 & ".上机人员表 where 用户名='" & strUser & "'"
    Call ExecuteProcedure("Zl_Zluserroles_Del('" & strUser & "')", Me.Caption)
    
    '插入重要操作日志
    Call SaveAuditLog(3, "删除用户", strUser, strRemarks)
    
    If UCase(gstrSTOwner) = UCase(mstr所有者) And gblnRIS And gblnMustRIS Then  '是标准版的所有者
        '通知新网该用户已经被删除
        If Not gobjRIS.UserEdit(3, strUser) Then
            MsgBox "当前启用了影像信息系统接口， 但由于影像信息系统接口(UserEdit)未调用成功，请联系管理员！", vbInformation, gstrSysName
        End If
    End If
    
    lvwUser.ListItems.Remove intIndex
    If lvwUser.ListItems.Count > 0 Then
        If intIndex > lvwUser.ListItems.Count Then intIndex = lvwUser.ListItems.Count
        lvwUser.ListItems(intIndex).Selected = True
        Call lvwUser_ItemClick(lvwUser.ListItems(intIndex))
    End If
    MousePointer = 0
    Call SetEnable
End Sub

Private Sub cmdModify_Click()
    '修改用户
    Dim strItem As String, arrTmp As Variant
    
    If lvwUser.SelectedItem Is Nothing Then Exit Sub
    If gblnMustRIS And Not gblnRIS And UCase(gstrSTOwner) = UCase(mstr所有者) Then
        MsgBox "RIS接口创建失败，不能继续当前操作。可能是接口文件安装或注册不正常，请与系统管理员联系。", vbInformation, gstrSysName
        Exit Sub
    End If
    If frmUserEdit.UserEdit(mstr所有者, lvwUser.SelectedItem.Text, strItem) Then
        If strItem = "" Then
            Set mrsUsers = Nothing
            Call cboSystem_Click
        ElseIf mLastIndex > 0 And mLastIndex < lvwUser.ListItems.Count Then
            arrTmp = Split(strItem, "|")
            lvwUser.ListItems(mLastIndex).SubItems(Col_人员编号) = arrTmp(0)
            lvwUser.ListItems(mLastIndex).SubItems(Col_人员姓名) = arrTmp(1)
            lvwUser.ListItems(mLastIndex).SubItems(Col_所属部门) = arrTmp(2)
            lvwUser.ListItems(mLastIndex).Selected = True
            Call lvwUser_ItemClick(lvwUser.ListItems(mLastIndex))
        End If
    End If
End Sub

Private Sub cmdUnDoLock_Click()
    '功能:对用户进行解锁
    Dim strKey As String, blnLock As Boolean
    
    If lvwUser.SelectedItem Is Nothing Then Exit Sub
    blnLock = Mid(lvwUser.SelectedItem.Tag, 1, 1) <> "1"
    strKey = lvwUser.SelectedItem.Key
    If MsgBox("确定要" & IIf(blnLock, "禁用", "启用") & "用户：“" & Mid(strKey, 2) & "”吗？", vbInformation + vbOKCancel + vbDefaultButton1) = vbCancel Then Exit Sub
    If LockUser(lvwUser.SelectedItem.Text, blnLock) = False Then Exit Sub
    Call FillUser
    err = 0: On Error Resume Next
    lvwUser.ListItems(strKey).Selected = True
    lvwUser.ListItems(strKey).EnsureVisible
    Call lvwUser_ItemClick(lvwUser.ListItems(strKey))
    Call SetEnable
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub cmdUpdatePWD_Click()
    Dim strUserName As String, strPassword As String
    Dim strError As String
    
    If lvwUser.SelectedItem Is Nothing Then Exit Sub
        
    strUserName = lvwUser.SelectedItem.Text
    strPassword = InputBox("请输入新的密码", "修改" & strUserName & "的密码", "123")
    
    If strPassword = "" Then Exit Sub
    
    If gobjRegister.UpdateUserPassword(gcnOracle, strUserName, strPassword, True, strError) Then
        MsgBox "修改" & strUserName & "的密码成功。", vbInformation + vbOKOnly, "提示"
        '插入重要操作日志
        Call SaveAuditLog(2, "修改密码", "修改用户：" & strUserName & "的密码")
    Else
        MsgBox "修改" & strUserName & "的密码失败。" & vbCrLf & strError, vbExclamation, "提示"
    End If
    
    If gstrUserName = strUserName Then
        MsgBox "修改当前用户的密码之后需要重新登录", vbInformation, "提示"
        frmUserLogin.Show 1
        If gcnOracle.State = adStateClosed Then
            End
        End If
    End If
End Sub

Private Sub cmdWhole_Click(Index As Integer)
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strDept As String, strError As String, strPrompt As String
    Dim strKey As String, strUserName As String
    Dim blnHaveRis As Boolean
    Dim blnMsgRis As Boolean
    Dim i As Long
    
    On Error GoTo errH
    Select Case Index
        Case WE_CreateAllUser  '批量创建用户(&1)
            If UCase(gstrSTOwner) = UCase(mstr所有者) Then   '是标准版的所有者
                blnHaveRis = gblnRIS
                If gblnMustRIS And Not gblnRIS Then
                    MsgBox "RIS接口创建失败，不能继续当前操作。可能是接口文件安装或注册不正常，请与系统管理员联系。", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            strDept = frmUserBatCreate.ShowMe(mstr所有者)
            '1、如果人员编号首位为英文字母，则以人员编号作为用户名
            '2、如果人员编号首位为数字，则以“U+人员编号”作为用户名
            '3、用户密码和用户名一致。
            If strDept = "" Then Exit Sub
            strSQL = "Select /*+Rule */" & vbNewLine & _
                        " a.Id, a.编号, a.姓名, a.简码" & vbNewLine & _
                        "From " & mstr所有者 & ".人员表 a," & mstr所有者 & ".部门人员 b, Table(Cast(f_Num2list('" & strDept & "') As Zltools.t_Numlist)) c" & vbNewLine & _
                        "Where a.Id = b.人员id And b.部门id = c.Column_Value And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And" & vbNewLine & _
                        "      Id Not In (Select 人员id From " & mstr所有者 & ".上机人员表)" & vbNewLine & _
                        "Order By a.编号"
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
            If rsTmp Is Nothing Then Exit Sub
            On Error Resume Next
            With rsTmp
                Call ShowFlash("正在创建用户，请稍候！", 0)
                For i = 1 To .RecordCount
                    If UCase(Left(!编号, 1)) >= "A" And UCase(Left(!编号, 1)) <= "Z" Then
                        strUserName = !编号
                    Else
                        strUserName = "U" & !编号
                    End If
                    frmMDIMain.stbThis.Panels(2).Text = "正在创建用户:" & strUserName
                    Call ShowFlash("正在创建用户:【" & strUserName & "】", i / .RecordCount)
                    strError = ""
                    Call gobjRegister.CreateUser(gcnOracle, strUserName, strUserName, strError)
                    If strError = "" Then
                        gcnOracle.Execute "Grant Connect,Alter Session,Create Session,Create Synonym,Create Table,Create View,Create Sequence,Create Database Link,Create Cluster to " & strUserName
                        gcnOracle.Execute "insert into " & mstr所有者 & ".上机人员表(用户名,人员id) values ('" & strUserName & "'," & !id & ")"
                        Call AlterUserTableSpaces(gcnOracle, strUserName)
                        '通知新网该用户被创建
                        If blnHaveRis Then
                            If Not gobjRIS.UserEdit(1, strUserName) Then
                                blnMsgRis = True
                            End If
                        End If
                    Else
                        strPrompt = strPrompt & vbCrLf & "[" & !strUserName & "]" & !姓名 & ":" & strError
                    End If
                    .MoveNext
                Next
                Call ShowFlash("")
                If strPrompt = "" Then
                    strPrompt = "全部人员正确设置为上机用户！"
                Else
                    strPrompt = "以下人员未正常设置为上机用户：" & strPrompt
                End If
                '插入重要操作日志
                Call SaveAuditLog(2, "批量创建用户", strPrompt)
                If blnMsgRis Then
                    strPrompt = strPrompt & vbNewLine & "当前启用了影像信息系统接口， 但由于影像信息系统接口(UserEdit)未调用成功，请联系管理员！"
                End If
                MsgBox strPrompt, vbInformation, gstrSysName
            End With
            '由于用户角色未发生变更，尽管新增用户，但是用户尚未对应角色，因此不更新用户角色记录
        Case WE_RestoreAllUser '根据上机人员恢复用户(&2)
            If MsgBox("本功能用户于“按用户恢复数据”模式下，恢复数据之后创建以前的用户并授权，将以用户名作为初始密码。" & vbCrLf _
                    & "你确定要继续吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            strSQL = "Select 用户名 From " & mstr所有者 & ".上机人员表 Where 用户名 Not In (Select Username From All_Users)"
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
            With rsTmp
                On Error Resume Next
                Call ShowFlash("正在创建用户，请稍候！", 0)
                For i = 1 To .RecordCount
                    frmMDIMain.stbThis.Panels(2).Text = "正在创建用户:" & !用户名
                    Call ShowFlash("正在创建用户【" & !用户名 & "】", i / .RecordCount)
                    Call gobjRegister.CreateUser(gcnOracle, !用户名, !用户名, strError)
                    If strError = "" Then
                        gcnOracle.Execute "Grant Connect,Alter Session,Create Session,Create Synonym,Create Table,Create View,Create Sequence,Create Database Link,Create Cluster to " & !用户名
                        Call AlterUserTableSpaces(gcnOracle, Nvl(!用户名))
                    Else
                        strPrompt = strPrompt & vbCrLf & !用户名 & ":" & strError
                    End If
                    .MoveNext
                Next
                Call ShowFlash("")
                If strPrompt = "" Then
                    strPrompt = "上机用户恢复完毕！"
                Else
                    strPrompt = "以下上机用户没有恢复：" & strPrompt
                End If
                MsgBox strPrompt, vbExclamation, gstrSysName
            End With
            '由于用户角色未发生变更，尽管新增用户，但是用户尚未对应角色，因此不更新用户角色记录
            '插入重要操作日志
            Call SaveAuditLog(2, "根据上机人员恢复用户", strPrompt)
        Case WE_RecUserRoles '重整所有用户角色(&3)
            If MsgBox("本功能将清除本系统保存的所有用户的角色，根据用户在数据库中实际拥有的角色重新产生本系统的所有用户的角色数据。" & vbCrLf & _
                        "当用户在应用系统中的角色与数据库中实际的角色不一致时，执行此操作可修正不一致的数据。" & vbCrLf & _
                        "你确定要继续吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            If Not CheckRushHours("0402", "重整所有用户角色") Then
                Exit Sub
            End If
            Call ExecuteProcedure("Zl_Zluserroles_Add()", Me.Caption)
            Call FillRole
            MsgBox "重整所有用户角色，操作完成。", vbInformation, gstrSysName
            '插入重要操作日志
            Call SaveAuditLog(2, "重整所有用户角色", "重整所有用户角色，操作完成")
        Case WE_RestoreUserRoles '恢复所有用户角色(&4)
            If MsgBox("本功能将以用户在应用系统中的记录角色重新进行角色授权。一般用于“按用户恢复数据”模式下，恢复角色和用户之后，重建用户的角色。" & vbCrLf & _
                        "你确定要继续吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            If Not CheckRushHours("0402", "恢复所有用户角色") Then
                Exit Sub
            End If
            strSQL = "Select 用户, 角色, 管理 From Zltools.Zluserroles Where 用户 In (Select Username From All_Users)"
            Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
            With rsTmp
                On Error Resume Next
                Call ShowFlash("正在授予用户角色，请稍候！", 0)
                For i = 1 To .RecordCount
                    frmMDIMain.stbThis.Panels(2).Text = "正在授予用户 " & !用户 & " " & !角色
                    Call ShowFlash("正在授予用户【" & !用户 & "】 " & !角色, i / .RecordCount)
                    gcnOracle.Execute "Grant " & !角色 & " to " & !用户 & IIf(!管理 = 1, " With Admin Option", "")
                    If err.Number <> 0 Then strPrompt = strPrompt & vbCrLf & !角色 & "授予" & !用户 & "失败": err.Clear
                    .MoveNext
                Next
                Call ShowFlash("")
                If strPrompt = "" Then
                    strPrompt = "用户角色恢复完毕"
                Else
                    strPrompt = "以下用户角色没有恢复：" & strPrompt
                End If
                MsgBox strPrompt, vbExclamation, gstrSysName
                Call FillRole
                frmMDIMain.stbThis.Panels(2).Text = "正在授予用户 " & !用户 & " " & !角色
            End With
            '插入重要操作日志
            Call SaveAuditLog(2, "恢复所有用户角色", strPrompt)
    End Select
    
    frmMDIMain.stbThis.Panels(2).Text = ""
    '重新加载用户，并恢复原始选择
    If Index = WE_CreateAllUser Or Index = WE_RestoreAllUser Then
        If Not lvwUser.SelectedItem Is Nothing Then strKey = lvwUser.SelectedItem.Key
        On Error GoTo errH
        Call FillUser
        err = 0: On Error Resume Next
        lvwUser.ListItems(strKey).Selected = True
        Call lvwUser_ItemClick(lvwUser.ListItems(strKey))
        Call SetEnable
        If err.Number <> 0 Then err.Clear
    End If
    Exit Sub
errH:
    frmMDIMain.stbThis.Panels(2).Text = ""
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub Form_Load()
    Dim strTmp As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    picHLine.Top = Val(GetSetting("ZLSOFT", "公共模块\服务器管理工具\用户管理", "PicHLine_TOP", "6500"))
    gblnMustRIS = Val(gclsBase.GetPara(255, 100, 0, , , "0")) = 1
    If gblnMustRIS Then
        Call CheckAndAdjustMustTable("zlParameters", "部门") '高版本程序，连接低版本数据库会出错
        gblnRIS = GetRIS
        If gblnRIS Then
            Call gobjRIS.InitConn(gcnOracle)
        End If
    Else
        gblnRIS = False
    End If
    mbytSearch = ME_部门: mnuPopuMenuSearch(ME_部门).Checked = True: txtSearch.Tag = "按部门过滤"
    Call PrintSearch("按部门过滤", vbBlue, False)
    If gstrSTOwner = "" Then
        gstrSTOwner = GetOwnerName(100, gcnOracle)
    End If
    '用户状态列
    lvwUser.ColumnHeaders(Col_所属部门 + 1).Width = lvwUser.ColumnHeaders(Col_所属部门 + 1).Width + IIf(gblnDBA, 0, 1000)
    lvwUser.ColumnHeaders(Col_用户状态 + 1).Width = IIf(gblnDBA, 1000, 0)
    cmdUnDoLock.Visible = gblnDBA
    cmdUpdatePWD.Visible = gblnDBA
    
    mstrBakOwner = ""
    On Error GoTo errH
    strSQL = "Select Upper(所有者) 所有者 From Zlbakspaces Where Db连接 Is Null Order by 所有者"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        If strTmp <> rsTmp!所有者 Then
            strTmp = rsTmp!所有者
            mstrBakOwner = mstrBakOwner & ",'" & strTmp & "'"
        End If
        rsTmp.MoveNext
    Loop
    mstrAllSysOwner = ""
    Call FillSystem
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, Me.Caption
End Sub

Private Sub Form_Resize()
    Dim lngTemp As Long
    err = 0: On Error Resume Next
    Me.cmdAdd.Left = Me.ScaleWidth - 200 - Me.cmdAdd.Width
    Me.cmdDelete.Left = Me.cmdAdd.Left
    Me.cmdModify.Left = Me.cmdAdd.Left
    Me.cmdUnDoLock.Left = Me.cmdAdd.Left
    Me.cmdUpdatePWD.Left = Me.cmdAdd.Left
    Me.cmdWhole(WE_CreateAllUser).Left = Me.cmdAdd.Left
    Me.lvwUser.Width = Me.cmdAdd.Left - 90 - Me.lvwUser.Left
    Me.lvwRole.Width = Me.ScaleWidth - Me.lvwRole.Left - 200
    fraFuncs.Top = Me.ScaleHeight - fraFuncs.Height - 100
    picHLine.Width = lvwRole.Width
    lvwUser.Height = picHLine.Top - lvwUser.Top
    lblRole.Top = picHLine.Top + picHLine.Height
    lvwRole.Top = lblRole.Top + lblRole.Height + 50
    lvwRole.Height = fraFuncs.Top - lvwRole.Top - 100
    
    txtSearch.Left = lvwUser.Left + lvwUser.Width - txtSearch.Width
    picSel.Left = txtSearch.Left - picSel.Width - 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mrsSystem.State = 1 Then mrsSystem.Close
    SaveSetting "ZLSOFT", "公共模块\服务器管理工具\用户管理", "PicHLine_TOP", picHLine.Top
    Set mrsSystem = Nothing
    Set mobjTip = Nothing
    mstr所有者 = ""
End Sub

Private Sub lvwRole_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objItem As ListItem
    Dim strTip As String, strTitle As String
    
    If mobjTip Is Nothing Then
        Call InitTips
    End If
    Set objItem = lvwRole.HitTest(x, y)
    If Not objItem Is Nothing Then
        If objItem.SubItems(1) = 1 Then
            If glngSysNo <> -1 Then
                strTip = "该角色授权存在问题，请先以多系统方式登录管理工具" & vbNewLine & "执行“重整所有用户角色”操作！"
            Else
                strTip = "该角色授权存在问题，请先执行“重整所有用户角色”操作！"
            End If
            strTitle = objItem.Text
        ElseIf objItem.SubItems(1) = 2 Then
            If glngSysNo <> -1 Then
                strTip = "该角色授权存在问题，请先以多系统方式登录管理工具" & vbNewLine & "执行“恢复所有用户角色”操作！"
            Else
                strTip = "该角色授权存在问题，请先执行“恢复所有用户角色”操作！"
            End If
            strTitle = objItem.Text
        ElseIf objItem.SubItems(1) = 4 Then
            If glngSysNo <> -1 Then
                strTip = "该角色不存在，请先以多系统方式登录管理工具" & vbNewLine & "执行“重整所有用户角色”操作！"
            Else
                strTip = "该角色不存在，请先执行“重整所有用户角色”操作！"
            End If
            strTitle = objItem.Text
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
End Sub

Private Sub InitTips()
    Set mobjTip = New clsTipSwap
    Set mobjTip.ParentControl = lvwRole
    mobjTip.Icon = TTIconInfo
    mobjTip.Style = TTBalloon
    mobjTip.Create
End Sub

Private Sub lvwUser_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '仍是刚才那列
        lvwUser.SortOrder = IIf(lvwUser.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwUser.SortKey = mintColumn
        lvwUser.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwUser_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call FillRole
    mLastIndex = Item.Index
End Sub

Private Sub mnuPopuMenuSearch_Click(Index As Integer)
    Dim i As Integer
    mbytSearch = Index
    For i = ME_部门 To ME_人员
        mnuPopuMenuSearch(i).Checked = i = Index
    Next
    txtSearch.Tag = Split(mnuPopuMenuSearch(Index).Caption, "(")(0)
    Call PrintSearch(txtSearch.Tag, vbBlue, False)
    If txtSearch.Enabled Then txtSearch.SetFocus
    If txtSearch.Text <> "" Then
        Call FillUser(True)
    End If
End Sub

Private Sub picHLine_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Me.picHLine.BackColor = &H8000000F: Me.picHLine.Top = Me.picHLine.Top + y
End Sub

Private Sub picHLine_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.picHLine.BackColor = Me.BackColor
    If Me.picHLine.Top < 4000 Then Me.picHLine.Top = 4000
    If Me.picHLine.Top > Me.ScaleHeight - 3000 Then Me.picHLine.Top = Me.ScaleHeight - 3000
    Call Form_Resize
End Sub

Private Sub picSel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If picSel.Tag = "In" Then
        If x < 0 Or y < 0 Or x > picSel.Width Or y > picSel.Height Then
            ReleaseCapture
            picSel.Tag = ""
            PrintSearch Me.txtSearch.Tag, vbBlue, False
        End If
    Else
        picSel.Tag = "In"
        SetCapture picSel.hwnd
        MousePointer = 99
        PrintSearch Me.txtSearch.Tag, vbRed, True
    End If
End Sub

Private Sub picSel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    PopupMenu Me.mnuPopuMenu, vbPopupMenuRightAlign, Me.picSel.Left + 600, Me.picSel.Top + Me.picSel.Height
    Call PrintSearch(Me.txtSearch.Tag, vbBlue, False)
    picSel.Tag = ""
End Sub


Private Sub txtSearch_Change()
    Call FillUser(True)
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Or KeyAscii = Asc("*") Or KeyAscii = Asc("_") Then
        KeyAscii = 0
    End If
End Sub


Private Sub PrintSearch(ByVal strTittle As String, ByVal lngColor As Long, ByVal blnBoderStyle As Boolean)
    '----------------------------------------------------------------------------------------
    '功能:打印指定的索引条件
    '参数:strTittle-标题
    '     lngColor-颜色值
    '     lngBoderStyl-是否加边框线
    '----------------------------------------------------------------------------------------
    '功能:打印时间范围
    With picSel
        picSel.Width = 980
        .Left = txtSearch.Left - .Width - 50
        .Cls
        '.FontUnderline = blnBoderStyle ' IIf(blnBoderStyle, 1, 0)
        '.ScaleWidth = .TextWidth(strTittle)
        .ForeColor = lngColor
         .FontUnderline = True
        .CurrentX = 10 '(.ScaleWidth - .TextWidth(strTittle))
        .CurrentY = (.ScaleHeight - .TextHeight(strTittle)) / 2
        picSel.Print strTittle
        .ZOrder 1
    End With
End Sub

Private Sub FillSystem()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strTmp As String
    
    '判断该用户能否创建用户
    On Error GoTo errH
    strSQL = "Select 1" & vbNewLine & _
                    "From User_Sys_Privs" & vbNewLine & _
                    "Where Privilege = 'CREATE USER'" & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select 1" & vbNewLine & _
                    "From Role_Sys_Privs" & vbNewLine & _
                    "Where Privilege = 'CREATE USER'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    cmdAdd.Enabled = rsTmp.RecordCount > 0
    '没有系统时添加删除修改不可用
    cmdAdd.Enabled = rsTmp.RecordCount > 0
    cmdDelete.Enabled = cmdAdd.Enabled
    cmdModify.Enabled = cmdAdd.Enabled
    cmdWhole(WE_CreateAllUser).Enabled = cmdAdd.Enabled
    cmdWhole(WE_RestoreAllUser).Enabled = cmdAdd.Enabled
    
    '显示系统所有者具有部门人员管理的系统
    If glngSysNo <> -1 Then
        cmdWhole(WE_RecUserRoles).Visible = False
        cmdWhole(WE_RestoreUserRoles).Visible = False
        Set mrsSystem = gclsBase.GetMenSystems(True)
        mrsSystem.Filter = "编号 = " & glngSysNo
    Else
        Set mrsSystem = gclsBase.GetMenSystems(True, True)
    End If
    If mrsSystem.RecordCount <= 0 Then Exit Sub
    Do While Not mrsSystem.EOF
        If strTmp <> mrsSystem!所有者 Then
            strTmp = mrsSystem!所有者
            mstrAllSysOwner = mstrAllSysOwner & "," & strTmp
        End If
        mrsSystem.MoveNext
    Loop
    mstrAllSysOwner = mstrAllSysOwner & ","
    '加载系统，最后触发系统选择
    '记录集过滤，空值默认优先
    If mrsSystem.RecordCount = 1 Then
        lblSys.Visible = False
        cboSystem.Visible = False
    Else
        mrsSystem.Filter = "人员管理=1": mrsSystem.Sort = "共享号,编号"
    End If
    cboSystem.Clear: cboSystem.Tag = ""
    mrsSystem.MoveFirst
    Do While Not mrsSystem.EOF
        cboSystem.addItem mrsSystem!名称 & " v" & mrsSystem!版本号 & "（" & mrsSystem!编号 & "）"
        cboSystem.ItemData(cboSystem.NewIndex) = mrsSystem!编号
        If mrsSystem!所有者 & "" = UCase(gstrUserName) And cboSystem.Tag = "" Then
            cboSystem.Tag = cboSystem.NewIndex
        End If
        mrsSystem.MoveNext
    Loop
    cboSystem.ListIndex = Val(cboSystem.Tag) '触发Click事件，加载用户
    Exit Sub
errH:
    MsgBox err.Description, vbCritical, Me.Caption
    If 1 = 0 Then
        Resume
    End If
End Sub

Private Sub FillUser(Optional blnFilter As Boolean = False)
'功能：填充用户
'blnFilter=是否过滤模式
    Dim strTmp As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strSearch As String, strIco As String
    Dim lst As ListItem, blnLock As Boolean
    Dim blnOwner As Boolean, bln过期 As Boolean
    
    On Error GoTo errH
    '显示可以进行当前系统的用户与对应的人员
    mrsSystem.Filter = "编号=" & cboSystem.ItemData(cboSystem.ListIndex)
    mstr所有者 = mrsSystem!所有者
    If blnFilter And Not mrsUsers Is Nothing Then
    Else
        '历史数据空间不应加入用户管理
        '其它系统的所有者不加入，不允许对其它系统的所有者授权，因为一个所有者的对象可能和其它系统的公共同义词冲突
        If gblnDBA Then
            strSQL = "Select u.Username, 编号, 姓名, 人员简码, 部门编码, 部门名称, 部门简码, m.Account_Status" & vbNewLine & _
                            "From All_Users u," & vbNewLine & _
                            "     (Select c.用户名, p.编号, p.姓名, p.简码 As 人员简码, d.编码 As 部门编码, d.名称 As 部门名称, d.简码 As 部门简码" & vbNewLine & _
                            "       From " & mstr所有者 & ".人员表 p, " & mstr所有者 & ".部门表 d, " & mstr所有者 & ".部门人员 b, " & mstr所有者 & ".上机人员表 c" & vbNewLine & _
                            "       Where p.Id = c.人员id And c.人员id = b.人员id And d.Id = b.部门id And" & vbNewLine & _
                            "             (p.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or p.撤档时间 Is Null) And b.缺省 = 1) r, Dba_Users m" & vbNewLine & _
                            "Where u.Username = r.用户名(+) And u.Username Not In (" & G_STR_USERS & mstrBakOwner & ") And u.User_Id = m.User_Id And" & vbNewLine & _
                            "      Not m.Default_Tablespace In ('SYSTEM', 'DRSYS') And u.Username Not Like 'ZLBAK%' And u.Username Not Like 'ZLHD%'"
        Else
            strSQL = "Select u.Username, 编号, 姓名, 人员简码, 部门编码, 部门名称, 部门简码, 'OPEN' As Account_Status" & vbNewLine & _
                            "From All_Users u," & vbNewLine & _
                            "     (Select c.用户名, p.编号, p.姓名, p.简码 As 人员简码, d.编码 As 部门编码, d.名称 As 部门名称, d.简码 As 部门简码" & vbNewLine & _
                            "       From " & mstr所有者 & ".人员表 p, " & mstr所有者 & ".部门表 d, " & mstr所有者 & ".部门人员 b, " & mstr所有者 & ".上机人员表 c" & vbNewLine & _
                            "       Where p.Id = c.人员id And c.人员id = b.人员id And d.Id = b.部门id And" & vbNewLine & _
                            "             (p.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or p.撤档时间 Is Null) And b.缺省 = 1) r" & vbNewLine & _
                            "Where u.Username = r.用户名(+) And u.Username Not In (" & G_STR_USERS & mstrBakOwner & ") And u.Username Not Like 'ZLBAK%' And u.Username Not Like 'ZLHD%'"
        End If
        Set mrsUsers = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    End If
    
    '数据过滤
    strSearch = Replace(Trim(UCase(txtSearch.Text)), "'", "")
    lvwUser.ListItems.Clear
    If strSearch = "" Then
        mrsUsers.Filter = 0
    Else
        Select Case mbytSearch
            Case ME_人员 '按人员
                mrsUsers.Filter = "编号 like '" & strSearch & "%' or 姓名 like '" & strSearch & "%' or 人员简码 like '" & strSearch & "%'"
            Case ME_用户 '按用户
                mrsUsers.Filter = "USERNAME like '" & strSearch & "%'"
            Case Else
                '按所属部门
                mrsUsers.Filter = "部门编码 like '" & strSearch & "%' or 部门名称 like '" & strSearch & "%' or 部门简码 like '" & strSearch & "%'"
        End Select
    End If
    '数据加载
    With mrsUsers
        Do While Not .EOF
            blnOwner = InStr(mstrAllSysOwner, "," & !USERNAME & ",") > 0
            If Not blnOwner Or gstrUserName = !USERNAME Then
                strIco = "User": blnLock = UCase(!ACCOUNT_STATUS & "") <> "OPEN"
                bln过期 = UCase(!ACCOUNT_STATUS & "") = "EXPIRED"
                If blnLock Then
                    strIco = "UserLock"
                ElseIf IsNull(!姓名) And Not blnOwner Then
                    strIco = "UserInfor"
                End If
                Set lst = lvwUser.ListItems.Add(, "K" & !USERNAME, !USERNAME, strIco, strIco)
                lst.SubItems(Col_人员编号) = !编号 & ""
                lst.SubItems(Col_人员姓名) = !姓名 & ""
                lst.SubItems(Col_所属部门) = !部门名称 & ""
                lst.SubItems(Col_用户状态) = IIf(blnLock, IIf(bln过期, "密码过期", "已锁"), "")
                lst.Tag = IIf(blnLock And Not bln过期, "1", "0") & IIf(blnOwner, 1, 0) & !姓名
            End If
            mrsUsers.MoveNext
        Loop
    End With
    
    If lvwUser.ListItems.Count > 0 Then
        If mLastIndex > 0 And mLastIndex < lvwUser.ListItems.Count Then
            lvwUser.ListItems(mLastIndex).Selected = True
        Else
            lvwUser.ListItems(1).Selected = True
        End If
        Call FillRole
    End If
    Call SetEnable
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub FillRole()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strUser As String
    Dim lst As ListItem
    
    On Error GoTo errH
    lvwRole.ListItems.Clear
    If lvwUser.SelectedItem Is Nothing Then
        Exit Sub
    Else
        strUser = lvwUser.SelectedItem.Text
    End If
    '显示该用户具有的角色
    '标记为1表示Dba_Role_Privs中有角色但Zluserroles中没有
    '标记为2表示Dba_Role_Privs中没有角色但Zluserroles中有，而且角色是存在的
    '标记为3表示两个表中共有
    '标记为4表示Dba_Role_Privs中没有角色但Zluserroles中有，而且角色是不存在的
    strSQL = "Select 角色, Sum(标记) 标记" & vbNewLine & _
            "From (Select Granted_Role 角色, 1 标记" & vbNewLine & _
            "       From Dba_Role_Privs" & vbNewLine & _
            "       Where Grantee = [1] And Granted_Role Like 'ZL_%'" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select a.角色, Decode(b.名称, Null, 4, 2) 标记" & vbNewLine & _
            "       From Zluserroles a, Zlroles b" & vbNewLine & _
            "       Where a.用户 = [1] And a.角色 = b.名称(+))" & vbNewLine & _
            IIf(glngSysNo = -1, "", " a Where Exists (Select 1 From Zlroles b Where a.角色 = b.名称 And b.系统 = [2])") & vbNewLine & _
            "Group By 角色"

    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, strUser, cboSystem.ItemData(cboSystem.ListIndex))
    Do While Not rsTmp.EOF
        If rsTmp!标记 = 1 Then
            Set lst = lvwRole.ListItems.Add(, , Mid(rsTmp!角色 & "", 4), "Role_User")
        ElseIf rsTmp!标记 = 2 Or rsTmp!标记 = 4 Then
            Set lst = lvwRole.ListItems.Add(, , Mid(rsTmp!角色 & "", 4), "Role_Dba")
        Else
            Set lst = lvwRole.ListItems.Add(, , Mid(rsTmp!角色 & "", 4), "Role")
        End If
        lst.SubItems(1) = rsTmp!标记
        rsTmp.MoveNext
    Loop
    Call SetEnable
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub SetEnable()
'设置各个按钮的Enable属性
    Dim blnHave As Boolean
    Dim blnLock As Boolean
    Dim blnOwner As Boolean '所有者
    
    blnHave = Not lvwUser.SelectedItem Is Nothing
    blnOwner = False
    If blnHave Then
        blnLock = Mid(lvwUser.SelectedItem.Tag, 1, 1) = "1"
        blnOwner = Mid(lvwUser.SelectedItem.Tag, 2, 1) = "1"
    End If
    cmdDelete.Enabled = cmdAdd.Enabled And blnHave And Not blnLock And blnOwner = False
    If cmdDelete.Enabled = True Then
        If lvwUser.SelectedItem.Text = "ZLTOOLS" Then cmdDelete.Enabled = False
    End If
    cmdModify.Enabled = blnHave And Not blnLock
    If blnLock = True Then
        cmdUnDoLock.Caption = "启用用户(&S)"
    Else
        cmdUnDoLock.Caption = "禁用用户(&J)"
    End If
End Sub

Private Function LockUser(ByVal strUser As String, Optional ByVal blnLock As Boolean = True) As Boolean
'功能:针对指定用户进行加锁或解锁
'参数:strUser-用户名
'     blnLock-true:加锁;false-解锁
'成功:加解锁成功,返回true,否则返回false
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    If blnLock Then
        '需要判断是否用其他用户进行连接了的.
        strSQL = "Select Osuser, Machine, Terminal As 终端, Program From gV$session Where Username = [1]"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, strUser)
        If Not rsTmp.EOF Then
            If MsgBox("警告: " & vbCrLf & "   用户" & strUser & "正连接在数据库上,禁用对已经登陆的用户将无效,是否还要对该用户进行禁用?", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
    End If
    strSQL = "alter user " & strUser & " ACCOUNT " & IIf(blnLock, "LOCK", "unlock ")
    '解锁和加锁
    err = 0: On Error Resume Next
    gcnOracle.Execute strSQL
    If err.Number <> 0 Then
        MsgBox "针对用户[" & strUser & "]的" & IIf(blnLock, "加锁", "解锁") & "失败,请稍后再继续!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        err.Clear
        Exit Function
    End If
    '插入重要操作日志
    Call SaveAuditLog(2, "启停用户", IIf(blnLock, "禁用", "启用") & "用户：" & strUser)
    LockUser = True
End Function

