VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConnectionsManager 
   BackColor       =   &H80000005&
   Caption         =   "数据连接"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9795
   Icon            =   "frmConnectionsManager.frx":0000
   LinkTopic       =   "form3"
   LockControls    =   -1  'True
   Picture         =   "frmConnectionsManager.frx":6852
   ScaleHeight     =   5715
   ScaleWidth      =   9795
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&E)"
      Height          =   350
      Left            =   7755
      TabIndex        =   8
      Top             =   5205
      Width           =   1100
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "测试(&T)"
      Height          =   350
      Left            =   6645
      TabIndex        =   4
      Top             =   5205
      Width           =   1100
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除(&D)"
      Height          =   350
      Left            =   5535
      TabIndex        =   3
      Top             =   5205
      Width           =   1100
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "修改(&M)"
      Height          =   350
      Left            =   4440
      TabIndex        =   2
      Top             =   5205
      Width           =   1100
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "新增(&A)"
      Height          =   350
      Left            =   3375
      TabIndex        =   1
      Top             =   5205
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   2985
      Left            =   180
      TabIndex        =   5
      Top             =   1605
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   5265
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
      Appearance      =   0
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "编号"
         Object.Tag             =   "编号"
         Text            =   "编号"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "连接名称"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "用户名"
         Object.Tag             =   "用户名"
         Text            =   "用户名"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "密码"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "IP"
         Object.Tag             =   "IP"
         Text            =   "IP"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Key             =   "端口号"
         Object.Tag             =   "端口号"
         Text            =   "端口号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "服务器"
         Object.Tag             =   "服务器"
         Text            =   "实例名"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "说明"
         Object.Tag             =   "说明"
         Text            =   "说明"
         Object.Width           =   8819
      EndProperty
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   195
      Picture         =   "frmConnectionsManager.frx":6D4B
      Top             =   600
      Width           =   480
   End
   Begin VB.Label lblIntroduce 
      BackColor       =   &H80000005&
      Caption         =   "数据连接用于在服务器预先存储连接到其他数据库的配置信息，提供给客户端连接其他数据库来查询数据，例如：在报表中查询其他数据库的数据。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   885
      TabIndex        =   7
      Top             =   645
      Width           =   7605
   End
   Begin VB.Label lblExplain 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "已创建的数据连接"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   180
      TabIndex        =   6
      Top             =   1305
      Width           =   1680
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "数据连接"
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
      Top             =   135
      Width           =   960
   End
End
Attribute VB_Name = "frmConnectionsManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnIsChange As Boolean  '记录界面是否发生改变
Private mblnTransferMode As Boolean  '记录界面的是否为frmConnManagerParent，是：True，否：False
Private Declare Function ImageList_Create Lib "COMCTL32" (ByVal MinCx As Long, ByVal MinCy As Long, ByVal Flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_MAXIMIZE = &HF030&
Private Const SC_MINIMIZE = &HF020&
Private Const SC_RESTORE = &HF120&
Private Const LVM_FIRST = &H1000
Private Const LVM_SETIMAGELIST = (LVM_FIRST + 3)
Private Const LVSIL_SMALL = 1
Private Const LVM_UPDATE = (LVM_FIRST + 42)
Private hImageList As Long
Private Enum lvwMainCol
    LC_连接名称 = 1
    LC_用户名 = 2
    LC_密码 = 3
    LC_IP地址 = 4
    LC_端口号 = 5
    LC_实例名 = 6
    LC_说明 = 7
End Enum

Private Sub cmdAdd_Click()
    Dim lngNumber As Long
    Dim strUser As String, strPasswd As String, strIp As String, strDatabase As String, strNotes As String, strPort As String, strLinkName As String
    
    If frmConnectionEdit.ShowEdit(lngNumber, strLinkName, strUser, strPasswd, strIp, strPort, strDatabase, strNotes) Then
        Call EditData("_" & lngNumber, strLinkName, strUser, strPasswd, strIp, strPort, strDatabase, strNotes)
        Call SetEnabled
        Call lvwMain_ItemClick(lvwMain.SelectedItem)
        lvwMain.ListItems("_" & lngNumber).EnsureVisible
    End If
End Sub

Private Sub CmdDelete_Click()
    Dim strKey As String, strSQL As String, strReportList As String
    Dim lngIndex As Long, i As Long, lngNumber As Long
    Dim rsTemp As ADODB.Recordset
    Dim strRemarks As String
    
    On Error GoTo errH:
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    If MsgBox("确定要删除编号为【" & lvwMain.SelectedItem.Text & "】的连接吗？", vbOKCancel + vbInformation, gstrSysName) = vbOK Then
        strKey = lvwMain.SelectedItem.Key
        lngNumber = Split(strKey, "_")(1)
        strSQL = "Select 名称 From Zlreports a Where Exists (Select 1 From Zlrptdatas b Where 数据连接编号 = [1] And a.Id = b.报表id)"

        Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "查询正在用此连接的报表", lngNumber)
        If rsTemp.RecordCount > 0 Then
            For i = 1 To rsTemp.RecordCount
                If i > 3 Then Exit For
                strReportList = strReportList & "【" & rsTemp!名称 & "】" & vbNewLine
                rsTemp.MoveNext
            Next
            If rsTemp.RecordCount > 3 Then
                MsgBox "该连接正在被" & vbNewLine & strReportList & "等" & rsTemp.RecordCount & _
                "个报表使用，要删除该连接请先将以上报表的数据源设为其它连接！", vbInformation, gstrSysName
            Else
                MsgBox "该连接正在被报表" & vbNewLine & strReportList & _
                "使用，要删除该连接请先将以上报表的数据源设为其它连接！", vbInformation, gstrSysName
            End If
            Exit Sub
        End If
        '验证身份并输入操作说明
        If Not CheckAuditStatus("0207", "删除", strRemarks) Then Exit Sub
        strSQL = "Zl_Zlconnections_Edit(2,'" & lngNumber & "')"
        Call ExecuteProcedure(strSQL, Me.Caption)
        '插入重要操作日志
        Call SaveAuditLog(3, "删除", "删除连接“" & lvwMain.SelectedItem.SubItems(LC_连接名称) & "”", strRemarks)
        lvwMain.Tag = ""
        With lvwMain
            lngIndex = .SelectedItem.Index
            .ListItems.Remove strKey
            Call SetEnabled
            If .ListItems.Count > 0 Then
                lngIndex = IIf(.ListItems.Count > lngIndex, lngIndex, .ListItems.Count)
                Call lvwMain_ItemClick(.ListItems(lngIndex))
            End If
        End With
        mblnIsChange = True
    End If
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdTest_Click()
    Dim strServerName As String
    Dim rsPasswd As ADODB.Recordset
    Dim cnOracle As ADODB.Connection
    Dim clsCiph As clsCipher
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    strServerName = lvwMain.SelectedItem.SubItems(LC_IP地址) & ":" & lvwMain.SelectedItem.SubItems(LC_端口号) & "/" & lvwMain.SelectedItem.SubItems(LC_实例名)
    
    Set clsCiph = New clsCipher
    Set cnOracle = gobjRegister.GetConnection(strServerName, lvwMain.SelectedItem.SubItems(LC_用户名), _
                    clsCiph.Decipher(MSTR_DBLINK_KEY, lvwMain.SelectedItem.SubItems(LC_密码)), False, OraOLEDB, , False)
    If cnOracle.State = adStateOpen Then
        MsgBox "连接成功！", vbInformation, gstrSysName
    End If

End Sub

Private Sub cmdUpdate_Click()
    Dim strKey As String
    Dim strUser As String, strPasswd As String, strIp As String, strDatabase As String
    Dim strNotes As String, strPort As String, strLinkName As String, strSQL As String, strReportList As String
    Dim lngNumber As Long, i As Long
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errH:
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    
    strLinkName = lvwMain.SelectedItem.SubItems(LC_连接名称)
    strUser = lvwMain.SelectedItem.SubItems(LC_用户名)
    strIp = lvwMain.SelectedItem.SubItems(LC_IP地址)
    strPort = lvwMain.SelectedItem.SubItems(LC_端口号)
    strDatabase = lvwMain.SelectedItem.SubItems(LC_实例名)
    strNotes = lvwMain.SelectedItem.SubItems(LC_说明)
    strPasswd = lvwMain.SelectedItem.SubItems(LC_密码)
    
    strKey = lvwMain.SelectedItem.Key
    lngNumber = Split(strKey, "_")(1)
    strSQL = "Select 名称 From Zlreports a Where Exists (Select 1 From Zlrptdatas b Where 数据连接编号 = [1] And a.Id = b.报表id)"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "查询正在用此连接的报表", lngNumber)
    
    If rsTemp.RecordCount > 0 Then
        For i = 1 To rsTemp.RecordCount
            If i > 3 Then Exit For
            strReportList = strReportList & "【" & rsTemp!名称 & "】" & vbNewLine
            rsTemp.MoveNext
        Next
        If rsTemp.RecordCount > 3 Then
            MsgBox "该连接正在被" & vbNewLine & strReportList & "等" & rsTemp.RecordCount & _
            "个报表使用，在修改后请检查以上报表是否可以正常使用！", vbInformation, gstrSysName
        Else
            MsgBox "该连接正在被报表" & vbNewLine & strReportList & "使用，在修改后请检查以上报表是否可以正常使用！", vbInformation, gstrSysName
        End If
    End If
    
    If frmConnectionEdit.ShowEdit(lngNumber, strLinkName, strUser, strPasswd, strIp, strPort, strDatabase, strNotes) Then
        Call EditData(strKey, strLinkName, strUser, strPasswd, strIp, strPort, strDatabase, strNotes)
    End If
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

Private Sub Form_Load()
    Call LoadConnInfor
    Call SetEnabled
    SetListViewRowHeight lvwMain.hwnd, 15
End Sub

Private Sub LoadConnInfor()
    '加载数据
    Dim strSQL As String
    Dim i As Long
    Dim objItem As ListItem
    Dim rsConnections As ADODB.Recordset
    
    On Error GoTo errH:
    strSQL = "Select 编号, 名称, 用户名, 密码, Ip, 端口, 实例名, 说明 From zlConnections Order By 编号"
    Set rsConnections = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "查询连接信息")
    lvwMain.ListItems.Clear
    With rsConnections
        Do While Not .EOF
            Set objItem = lvwMain.ListItems.Add(, "_" & !编号, !编号)
            objItem.SubItems(LC_连接名称) = !名称
            objItem.SubItems(LC_用户名) = !用户名
            objItem.SubItems(LC_密码) = !密码
            objItem.SubItems(LC_IP地址) = !IP
            objItem.SubItems(LC_端口号) = !端口
            objItem.SubItems(LC_实例名) = !实例名
            objItem.SubItems(LC_说明) = "" & !说明
            .MoveNext
        Loop
    End With
    If rsConnections.RecordCount <> 0 Then
        Call lvwMain_ItemClick(lvwMain.SelectedItem)
    End If
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

Private Sub SetEnabled()
    '设置按钮是否可用
    If lvwMain.SelectedItem Is Nothing Then
        cmdUpdate.Enabled = False
        cmdDelete.Enabled = False
        cmdTest.Enabled = False
    Else
        cmdUpdate.Enabled = True
        cmdDelete.Enabled = True
        cmdTest.Enabled = True
    End If
End Sub

Private Sub Form_Resize()
    Dim i As Long
    
    On Error Resume Next
    If mblnTransferMode = True Then
        cmdExit.Visible = False
        cmdTest.Left = Me.ScaleWidth - cmdTest.Width - 150
        cmdTest.Top = Me.ScaleHeight - cmdTest.Height - 150
    Else
        cmdExit.Left = Me.ScaleWidth - cmdExit.Width - 150
        cmdExit.Top = Me.ScaleHeight - cmdExit.Height - 150
        cmdTest.Left = cmdExit.Left - cmdTest.Width
        cmdTest.Top = cmdExit.Top
    End If
    cmdDelete.Left = cmdTest.Left - cmdDelete.Width
    cmdDelete.Top = cmdTest.Top
    cmdUpdate.Left = cmdDelete.Left - cmdUpdate.Width
    cmdUpdate.Top = cmdTest.Top
    cmdAdd.Left = cmdUpdate.Left - cmdAdd.Width
    cmdAdd.Top = cmdTest.Top
    lvwMain.Width = Me.ScaleWidth - lvwMain.Left - 150
    lvwMain.Height = cmdTest.Top - lvwMain.Top - 150
    lvwMain.ColumnHeaders(8).Width = lvwMain.Width - lvwMain.ColumnHeaders(1).Width - _
                                        lvwMain.ColumnHeaders(2).Width - lvwMain.ColumnHeaders(3).Width - _
                                        lvwMain.ColumnHeaders(4).Width - lvwMain.ColumnHeaders(5).Width - _
                                        lvwMain.ColumnHeaders(6).Width - lvwMain.ColumnHeaders(7).Width
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub EditData(ByVal strKey As String, ByVal strLinkName As String, ByVal strUser As String, _
                    ByVal strPasswd As String, ByVal strIp As String, ByVal strPort As String, ByVal strDatabase As String, ByVal strNotes As String)
    Dim objItem As ListItem
    
    '更新本界面数据
    On Error Resume Next
    Set objItem = lvwMain.ListItems(strKey)
    If err.Number <> 0 Then
        Set objItem = lvwMain.ListItems.Add(, strKey, Split(strKey, "_")(1))
        err.Clear
    End If
    objItem.SubItems(LC_连接名称) = strLinkName
    objItem.SubItems(LC_用户名) = strUser
    objItem.SubItems(LC_密码) = strPasswd
    objItem.SubItems(LC_IP地址) = strIp
    objItem.SubItems(LC_端口号) = strPort
    objItem.SubItems(LC_实例名) = strDatabase
    objItem.SubItems(LC_说明) = strNotes
    mblnIsChange = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetListViewRowHeight_Destroy
End Sub

Private Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Item.Selected = True
    Item.EnsureVisible
    If lvwMain.Tag <> "" Then
        Call SetSelItemBold(lvwMain.ListItems(lvwMain.Tag), False)
    End If
    Call SetSelItemBold(Item, True)
    lvwMain.Tag = Item.Key
End Sub

Private Sub SetSelItemBold(ByVal itm As ListItem, ByVal blnBold As Boolean)
    Dim i As Integer
        
    '设置是否加粗
    itm.Bold = blnBold
    For i = 1 To itm.ListSubItems.Count
        itm.ListSubItems(i).Bold = blnBold
    Next
End Sub

Public Function ShowMe(ByRef frmParent As Object, ByRef blnIsChange As Boolean) As Boolean
    '--------------------------------------------
    '显示窗体
    'frmParent为该窗体的父窗体
    'blnIsChange为判断该窗体是否发生改变
    '--------------------------------------------
    Dim strUnit As String, strHaveProg As String, strSQL As String
    Dim strDest() As Byte, StrJiemi() As Byte
    Dim blnGrantMgr As Boolean
    Dim rsTemp As ADODB.Recordset
    
    '检查权限
    If Not gblnDBA And Not gblnOwner Then
        '获取解密秘钥
        strUnit = gobjRegister.zlRegInfo("单位名称", False, 0)
        If strUnit = "" Then End

        '检查是否有操作“连接管理”的权限
        strSQL = "select 功能 from zltools.Zlmgrgrant Where 用户名='" & gstrLoginUserName & "'"
        Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "管理工具授权用户")
        If rsTemp.RecordCount > 0 Then
            strHaveProg = rsTemp!功能 & ""
            If strHaveProg <> "" Then
                ReDim Preserve strDest(0): ReDim Preserve StrJiemi(0)
                Call Func16CodeToByte(strHaveProg, strDest)
                Call DES_Decode(strDest, StrJiemi, strUnit)
                strHaveProg = Replace(StrConv(StrJiemi, vbUnicode), Chr(0), "")
                blnGrantMgr = True
            End If
        End If
        If Not blnGrantMgr Then
            ShowMe = False
            MsgBox "您没有连接管理的使用权限，请联系管理员。", vbInformation, gstrSysName
            Exit Function
        ElseIf strHaveProg = "" Then
            ShowMe = False
            MsgBox "您的连接管理的使用权限丢失，请联系管理员重新授权。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If frmParent.Name = "frmConnManagerParent" Then
        mblnTransferMode = True
        Call FormSetCaption(Me, False, False)
        SetParent Me.hwnd, frmParent.hwnd
        Call SendMessage(frmParent.hwnd, WM_SYSCOMMAND, SC_RESTORE, 0)
        Call SendMessage(frmParent.hwnd, WM_SYSCOMMAND, SC_MAXIMIZE, 0)
        ShowWindow Me.hwnd, 3
    Else
        mblnTransferMode = False
        Me.Show vbModal, frmParent
    End If
    blnIsChange = mblnIsChange
    ShowMe = True
End Function

'调整listview行高
Private Sub SetListViewRowHeight(ByVal listViewHwnd As Long, ByVal rowHeight As Long)
    Call SetListViewRowHeight_Destroy
    hImageList = ImageList_Create(1, rowHeight, 1, 0, 0)
    SendMessage listViewHwnd, LVM_SETIMAGELIST, LVSIL_SMALL, ByVal hImageList
    SendMessage listViewHwnd, LVM_UPDATE, 0, ByVal 0
End Sub

Private Sub SetListViewRowHeight_Destroy()
    If hImageList <> 0 Then ImageList_Destroy hImageList
End Sub

