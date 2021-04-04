VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmClientUpgradeSeverConfigure 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "文件服务器配置"
   ClientHeight    =   6255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12750
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   12750
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraOption 
      BorderStyle     =   0  'None
      Height          =   252
      Left            =   120
      TabIndex        =   11
      Top             =   5880
      Width           =   4932
      Begin VB.CheckBox chkSampleServer 
         Caption         =   "下载前不检查文件是否存在（适用于简易FTP工具）"
         Height          =   180
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   4932
      End
   End
   Begin VB.PictureBox picBtn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   120
      ScaleHeight     =   330
      ScaleWidth      =   5385
      TabIndex        =   10
      Top             =   60
      Width           =   5385
      Begin VB.CommandButton cmdAdd 
         Caption         =   "新增(&A)"
         Height          =   300
         Left            =   0
         TabIndex        =   0
         ToolTipText     =   "新增一个升级或收集服务器"
         Top             =   0
         Width           =   900
      End
      Begin VB.CommandButton cmdCheck 
         Caption         =   "服务器可用性检测(&X)"
         Height          =   300
         Left            =   3240
         TabIndex        =   3
         ToolTipText     =   "测试校验服务器是否能连接成功"
         Top             =   0
         Width           =   2000
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除(&D)"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2160
         TabIndex        =   2
         ToolTipText     =   "删除一个服务器信息"
         Top             =   0
         Width           =   900
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "修改(&S)"
         Enabled         =   0   'False
         Height          =   300
         Left            =   1080
         TabIndex        =   1
         ToolTipText     =   "修改一个服务器信息"
         Top             =   0
         Width           =   900
      End
   End
   Begin VB.PictureBox picFilter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   165
      ScaleHeight     =   255
      ScaleWidth      =   4005
      TabIndex        =   9
      Top             =   5610
      Width           =   4000
      Begin VB.OptionButton optFilter 
         Caption         =   "停用"
         Height          =   240
         Index           =   2
         Left            =   3195
         TabIndex        =   7
         ToolTipText     =   "显示停用的服务器"
         Top             =   0
         Width           =   720
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "启用"
         Height          =   240
         Index           =   1
         Left            =   2190
         TabIndex        =   6
         ToolTipText     =   "显示启用的服务器"
         Top             =   0
         Width           =   720
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "全部"
         Height          =   240
         Index           =   0
         Left            =   1080
         TabIndex        =   5
         ToolTipText     =   "显示所有服务器"
         Top             =   0
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Label lblFilter 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "服务器列表"
         Height          =   180
         Left            =   0
         TabIndex        =   4
         Top             =   15
         Width           =   900
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfMain 
      Height          =   4995
      Left            =   90
      TabIndex        =   8
      Top             =   465
      Width           =   12495
      _cx             =   22040
      _cy             =   8811
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   7000
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmClientUpgradeSeverConfigure.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
End
Attribute VB_Name = "frmClientUpgradeSeverConfigure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'=================================================================
'模块变量
'=================================================================
Private Enum ServerListCols
    Col_编号 = 0
    Col_类型 = 1
    Col_服务器状态 = 2 '启用 or 停用
    Col_服务器路径 = 3
    Col_用户名 = 4
    Col_密码 = 5
    Col_端口 = 6
    Col_是否升级 = 7
    Col_是否缺省 = 8
    Col_是否收集 = 9
    Col_收集类型 = 10
    Col_检测结果 = 11
End Enum
Private mblnHaveDefault As Boolean '是否存在默认服务器
Private mblnAllowEdit As Boolean '标记当前界面是否允许编辑
'=================================================================
'公共接口
'=================================================================
Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口
End Sub

Public Sub SetMenu()
'功能：设置状态栏数据
    frmMDIMain.stbThis.Panels(2).Text = "列表中共显示有" & vsfMain.Rows - 1 & "行数据。"
End Sub
'
Public Sub RefreshData()
'功能：主窗体调用的刷新数据接口
    Call LoadSeverListData
End Sub

'=================================================================
'私有方法
'=================================================================
Private Sub chkSampleServer_Click()
    If chkSampleServer.Tag <> "" Then
        Call gclsBase.UpdateZLReginfo("FTP不检查文件存在", chkSampleServer.value)
    End If
End Sub

Private Sub cmdAdd_Click()
    Dim frmEdit As New frmClientUpgradeSeverEdit
    If frmEdit.ShowMe(0, mblnHaveDefault) Then
        Call LoadSeverListData
    End If
End Sub

Private Sub cmdCheck_Click()
    Dim i As Long, objConn As clsConnect
    Dim strErr As String
    
    With vsfMain
        If .Rows < .FixedRows Then Exit Sub
        For i = .FixedRows To .Rows - 1
            ShowFlash "正在检测" & .TextMatrix(i, Col_编号) & "号: " & .TextMatrix(i, Col_服务器路径), (i - 1) / (.Rows - 1), Me, True
            DoEvents
            Set objConn = New clsConnect
            strErr = ""
            If Not objConn.ToConnect(IIf(Trim(.TextMatrix(i, Col_类型)) = "FTP", SCT_FTP, SCT_Share), .TextMatrix(i, Col_服务器路径), .TextMatrix(i, Col_用户名), .Cell(flexcpData, i, Col_密码), Val(.TextMatrix(i, Col_端口)), "", False, strErr) Then
                .TextMatrix(i, Col_检测结果) = "不可用：" & strErr
            Else
                .TextMatrix(i, Col_检测结果) = "可用"
            End If
            ShowFlash "正在检测" & .TextMatrix(i, Col_编号) & "号: " & .TextMatrix(i, Col_服务器路径), i / (.Rows - 1), Me, True
            Call objConn.CloseConnect
        Next
        Call ShowFlash("")
    End With
End Sub

Private Sub cmdDel_Click()
    Dim strSQL As String
    Dim strRemarks As String
    
    If vsfMain.TextMatrix(vsfMain.Row, Col_是否缺省) <> "" Then
        MsgBox vsfMain.TextMatrix(vsfMain.Row, Col_编号) & " 号" & vsfMain.TextMatrix(vsfMain.Row, Col_类型) & "服务器为缺省服务器不能删除，请切换缺省服务器后删除！", vbInformation, gstrSysName
        Exit Sub
    End If
    If MsgBox("确定要删除 " & vsfMain.TextMatrix(vsfMain.Row, Col_编号) & " 号" & vsfMain.TextMatrix(vsfMain.Row, Col_类型) & "服务器？", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
        '验证身份并输入操作说明
        If Not CheckAuditStatus("0307", "文件服务器配置-删除", strRemarks) Then Exit Sub
        strSQL = "Zl_Zlupgradeserver_Update(2," & vsfMain.TextMatrix(vsfMain.Row, Col_编号) & ")"
        Call ExecuteProcedure(strSQL, Me.Caption, gcnOracle)
        '插入重要操作日志
        Call SaveAuditLog(3, "文件服务器配置-删除", "删除编号为" & vsfMain.TextMatrix(vsfMain.Row, Col_编号) & "的文件服务器", strRemarks)
        Call LoadSeverListData
    End If
End Sub

Private Sub cmdModify_Click()
    Dim frmEdit As New frmClientUpgradeSeverEdit
    If frmEdit.ShowMe(Val(vsfMain.TextMatrix(vsfMain.Row, Col_编号)), mblnHaveDefault) Then
        Call LoadSeverListData
    End If
End Sub

Private Sub Form_Load()
    mblnAllowEdit = True
    Call TransOldData
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fraOption.Top = Me.ScaleHeight - fraOption.Height - 60
    picFilter.Top = fraOption.Top - picFilter.Height - 90
    vsfMain.Height = picFilter.Top - 90 - vsfMain.Top
    vsfMain.Width = Me.ScaleWidth - 120
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub optFilter_Click(Index As Integer)
    Dim i As Integer
    
    With vsfMain
        If .Rows < .FixedRows Then Exit Sub
        For i = 1 To .Rows - 1
            .RowHidden(i) = Not ((Index = .Cell(flexcpData, i, Col_服务器状态)) Or (Index = 0))
        Next
    End With
End Sub

Private Sub vsfMain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnAllowEdit = False Then Exit Sub
    cmdModify.Enabled = NewRow >= vsfMain.FixedRows
    cmdDel.Enabled = NewRow >= vsfMain.FixedRows
    cmdCheck.Enabled = NewRow >= vsfMain.FixedRows
End Sub

Private Sub vsfMain_DblClick()
    Dim intUpdate       As Integer, intDefault As Integer, intCollect As Integer
    Dim strFilesType    As String, strSQL      As String
    
    If mblnAllowEdit = False Then Exit Sub
    With vsfMain
        If .MouseRow <> .Row Then Exit Sub
        intUpdate = IIf(.TextMatrix(.Row, Col_是否升级) = "√", 1, 0)
        intDefault = IIf(.TextMatrix(.Row, Col_是否缺省) = "√", 1, 0)
        intCollect = IIf(.TextMatrix(.Row, Col_是否收集) = "√", 1, 0)
        strFilesType = .TextMatrix(.Row, Col_收集类型)
        If intDefault = 1 And (.ColSel = Col_是否升级 Or .ColSel = Col_是否缺省 Or .ColSel = Col_是否收集) Then
            Call MsgBox("选中编号 " & .TextMatrix(.Row, Col_编号) & " 服务器为缺省服务器，可以先将其他服务器切换为缺省以请保证有一个缺省服务器。", vbInformation, gstrSysName)
            Exit Sub
        ElseIf Not (.ColSel = Col_是否升级 Or .ColSel = Col_是否缺省 Or .ColSel = Col_是否收集) Then
            Exit Sub
        End If
        On Error GoTo ErrH
        Select Case .ColSel
            Case Col_是否升级
                If intCollect = 1 Then
                    If MsgBox("选中编号 " & .TextMatrix(.Row, Col_编号) & " 服务器为收集服务器，是否要切换为升级服务器？ ", vbInformation + vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                        Exit Sub
                    End If
                ElseIf intUpdate = 1 Then
                    If MsgBox("是否要取消该升级服务器，取消后将会清空已设置过该服务器为升级服务器的客户端", vbInformation + vbOKCancel, gstrSysName) = vbCancel Then
                        Exit Sub
                    End If
                End If
                strFilesType = ""
                intUpdate = IIf(intUpdate = 1, 0, 1)
                intCollect = 0
                '当设置为升级服务器，且没有缺省服务器，则自动缺省
                If intUpdate = 1 And Not mblnHaveDefault Then intDefault = 1
            Case Col_是否缺省
                If intCollect = 1 Then
                    If MsgBox("选中编号 " & .TextMatrix(.Row, Col_编号) & " 服务器为收集服务器，是否要切换为升级服务器并设置为缺省服务器？ ", vbInformation + vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                        Exit Sub
                    End If
                ElseIf intUpdate = 0 Then
                    If MsgBox("选中编号 " & .TextMatrix(.Row, Col_编号) & " 服务器为停用状态，是否要启用该服务器并设置为缺省服务器？ ", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                        Exit Sub
                    End If
                End If
                '用户升级并则设置缺省
                strFilesType = ""
                intUpdate = 1
                intCollect = 0
                intDefault = 1
            Case Col_是否收集
                If intUpdate = 0 Then
                    If MsgBox("选中编号 " & .TextMatrix(.Row, Col_编号) & " 服务器为升级服务器，是否要切换为收集服务器？ ", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                        Exit Sub
                    End If
                End If
                strFilesType = ""
                intUpdate = 0
                intCollect = IIf(intCollect = 1, 0, 1)
                intDefault = 0
        End Select
        strSQL = "Zl_Zlupgradeserver_Update(11," & .TextMatrix(.Row, Col_编号) & "," & IIf(.TextMatrix(.Row, Col_类型) = "共享", 0, 1) & ",'" & Trim(.TextMatrix(.Row, Col_服务器路径)) & "','" & Trim(.TextMatrix(.Row, Col_用户名)) & "'," & SQLAdjust(Cipher(Trim(.Cell(flexcpData, .Row, Col_密码)))) & ",'" & Trim(.TextMatrix(.Row, Col_端口)) & "'," & intUpdate & "," & intDefault & "," & intCollect & "," & SQLAdjust(strFilesType) & "," & SQLAdjust(Trim(.Cell(flexcpData, .Row, Col_密码))) & ")"
        Call ExecuteProcedure(strSQL, Me.Caption, gcnOracle)
        Call LoadSeverListData
        optFilter.Item(0).value = True
    End With
    Exit Sub
ErrH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

Private Sub vsfMain_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Row < vsfMain.FixedRows
End Sub

Public Sub LoadSeverListData()
'功能：加载升级文件服务器清单
    Dim lngRow  As Long
    Dim strSQL  As String, rsTmp As ADODB.Recordset

    On Error GoTo ErrH
    mblnHaveDefault = False
    '加载使用简易FTP设置
    strSQL = "Select 内容 As 使用简易ftp工具 From Zlreginfo Where 项目 =[1]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, "FTP不检查文件存在")
    If rsTmp.EOF Then
        chkSampleServer.value = 0
        Call gclsBase.UpdateZLReginfo("FTP不检查文件存在", 0, 1)
    Else
        chkSampleServer.value = Val(rsTmp!使用简易ftp工具 & "")
    End If
    chkSampleServer.Tag = "数据已经加载"
    strSQL = "Select 编号, 类型, 位置, 用户名, 密码, 端口, 是否升级, 是否缺省 From ZLTOOLS.Zlupgradeserver Order By 编号"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    With vsfMain
        .Rows = .FixedRows
        Do While Not rsTmp.EOF
            .Rows = .Rows + 1
            lngRow = .Rows - 1
            .TextMatrix(lngRow, Col_编号) = rsTmp!编号 & ""
            .TextMatrix(lngRow, Col_类型) = IIf(Val(rsTmp!类型 & "") = 1, "FTP", "共享")
            .TextMatrix(lngRow, Col_服务器路径) = rsTmp!位置 & ""
            .TextMatrix(lngRow, Col_用户名) = rsTmp!用户名 & ""
            .TextMatrix(lngRow, Col_密码) = "***"
            .Cell(flexcpData, lngRow, Col_密码) = Decipher(rsTmp!密码 & "")
            .TextMatrix(lngRow, Col_端口) = rsTmp!端口 & ""
            .Cell(flexcpBackColor, lngRow, Col_是否升级, lngRow, Col_是否收集) = RGB(210, 240, 255)
            .TextMatrix(lngRow, Col_是否升级) = IIf(Val(rsTmp!是否升级 & "") = 1, "√", "")
            .TextMatrix(lngRow, Col_是否缺省) = IIf(Val(rsTmp!是否缺省 & "") = 1, "√", "")
            If .TextMatrix(lngRow, Col_是否缺省) = "√" Then
                mblnHaveDefault = True
                .Cell(flexcpForeColor, lngRow, .FixedCols, lngRow, .Cols - 1) = vbBlue
            End If
            .TextMatrix(lngRow, Col_检测结果) = ""
            If .TextMatrix(lngRow, Col_是否升级) = "" And .TextMatrix(lngRow, Col_是否缺省) = "" And .TextMatrix(lngRow, Col_是否收集) = "" Then
                .TextMatrix(lngRow, Col_服务器状态) = "停用"
                .Cell(flexcpForeColor, lngRow, .FixedCols, lngRow, .Cols - 1) = vbGrayText
                .Cell(flexcpData, lngRow, Col_服务器状态) = 2
            Else
                .TextMatrix(lngRow, Col_服务器状态) = "启用"
                .Cell(flexcpData, lngRow, Col_服务器状态) = 1
            End If
            rsTmp.MoveNext
        Loop
        If lngRow > .FixedRows Then
            .Row = .FixedRows
        End If
        Call SetMenu
    End With
    Exit Sub
ErrH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox "服务器列表加载错误,信息:" & err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

Private Function TransOldData() As Boolean
'功能：讲旧设置防盗新的存储方案中
    Dim strSQL          As String, rsTmp        As ADODB.Recordset, rsNum As ADODB.Recordset
    Dim intClientUpType As Integer, strFileType As String
    Dim lngServerNO     As Integer, strTmp      As String
    Dim strUser         As String, strPwd       As String, strPort  As String, strPath  As String
    Dim blnSetDefault   As Boolean
    
    On Error GoTo ErrH
    '先判断有无数据
    strSQL = "Select 1 From Zlupgradeserver Where Rownum < 2"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    If Not rsTmp.EOF Then TransOldData = True: Exit Function
    '获取默认升级类型
    strSQL = "Select Max(内容) As 升级类型 From Zlreginfo Where 项目 = [1]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, "升级类型")
    intClientUpType = Val(rsTmp!升级类型 & "")
    '先转移FTP升级类型的服务器
    strSQL = "Select 项目, 内容" & vbNewLine & _
            "From Zlreginfo" & vbNewLine & _
            "Where (项目 Like 'FTP服务器%' Or 项目 Like 'FTP用户%' Or 项目 Like 'FTP密码%' Or 项目 Like 'FTP端口%')" & vbNewLine & _
            "And 内容 Is Not Null"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    '1.先处理FTP服务器
    rsTmp.Filter = "项目 Like 'FTP服务器*'"
    Set rsNum = CopyNewRec(rsTmp)
    rsNum.Sort = "项目"
    Do While Not rsNum.EOF
        strTmp = Mid(rsNum!项目, Len("FTP服务器") + 1)
        strUser = "": strPwd = "": strPort = "": strPath = ""
        strPath = rsNum!内容 & ""
        rsTmp.Filter = "项目='FTP用户" & strTmp & "'"
        If Not rsTmp.EOF Then strUser = rsTmp!内容 & ""
        rsTmp.Filter = "项目='FTP密码" & strTmp & "'"
        If Not rsTmp.EOF Then strPwd = rsTmp!内容 & ""
        rsTmp.Filter = "项目='FTP端口" & strTmp & "'"
        If Not rsTmp.EOF Then strPort = rsTmp!内容 & ""
        lngServerNO = lngServerNO + 1
        strSQL = "Zl_Zlupgradeserver_Update(0," & lngServerNO & ",1," & SQLAdjust(strPath) & "," & SQLAdjust(strUser) & "," & SQLAdjust(Cipher(strPwd)) & "," & Val(strPort) & ",1," & IIf(intClientUpType = 1 And Not blnSetDefault, 1, 0) & ",0,NULL," & SQLAdjust(strPwd) & ")"
        Call ExecuteProcedure(strSQL, Me.Caption, gcnOracle)
        If Not blnSetDefault Then
            blnSetDefault = intClientUpType = 1
        End If
        If intClientUpType = 1 Then
            strSQL = "Update Zltools.Zlclients Set 升级文件服务器 = " & lngServerNO & " Where Ftp服务器 " & IIf(strTmp = "", "Is Null", "=" & strTmp)
            Call gclsBase.ExecuteCmdText(strSQL, Me.Caption, gcnOracle)
        End If
        rsNum.MoveNext
    Loop
    '再转移共享升级类型的服务器
    strSQL = "Select 项目, 内容" & vbNewLine & _
            "From Zlreginfo" & vbNewLine & _
            "Where (项目 Like '服务器目录%' Or 项目 Like '访问用户%' Or 项目 Like '访问密码%')" & vbNewLine & _
            "And 内容 Is Not Null"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)

    rsTmp.Filter = "项目 Like '服务器目录*'"
    Set rsNum = CopyNewRec(rsTmp)
    rsNum.Sort = "项目"
    Do While Not rsNum.EOF
        strTmp = Mid(rsNum!项目, Len("服务器目录") + 1)
        strUser = "": strPwd = "": strPath = ""
        strPath = rsNum!内容 & ""
        rsTmp.Filter = "项目='访问用户" & strTmp & "'"
        If Not rsTmp.EOF Then strUser = rsTmp!内容 & ""
        rsTmp.Filter = "项目='访问密码" & strTmp & "'"
        If Not rsTmp.EOF Then strPwd = rsTmp!内容 & ""
        lngServerNO = lngServerNO + 1
        strSQL = "Zl_Zlupgradeserver_Update(0," & lngServerNO & ",0," & SQLAdjust(strPath) & "," & SQLAdjust(strUser) & "," & SQLAdjust(Cipher(strPwd)) & ",NULL,1," & IIf(intClientUpType = 0 And Not blnSetDefault, 1, 0) & ",0,NULL," & SQLAdjust(strPwd) & ")"
        Call ExecuteProcedure(strSQL, Me.Caption, gcnOracle)
        If Not blnSetDefault Then
            blnSetDefault = intClientUpType = 0
        End If
        If intClientUpType = 0 Then
            strSQL = "Update Zltools.Zlclients Set 升级文件服务器 = " & lngServerNO & " Where 升级服务器 " & IIf(strTmp = "", "Is Null", "=" & strTmp)
            Call gclsBase.ExecuteCmdText(strSQL, Me.Caption, gcnOracle)
        End If
        rsNum.MoveNext
    Loop
    If lngServerNO > 0 Then
        '清空客户端曾经配置的升级服务器
        strSQL = "Update Zlclients Set 升级服务器 = Null, Ftp服务器 = Null"
        Call gclsBase.ExecuteCmdText(strSQL, Me.Caption, gcnOracle)
    End If
    '获取收集类型与收集方式
    strSQL = "Select Max(内容) As 收集方式 From Zlreginfo Where 项目 = [1]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, "收集方式")
    intClientUpType = Val(rsTmp!收集方式 & "")
    strSQL = "Select Max(内容) As 收集类型 From Zlreginfo Where 项目 = [1]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, "收集类型")
    strFileType = rsTmp!收集类型 & ""
    '处理FTP收集服务器
    strSQL = "Select 项目, 内容" & vbNewLine & _
        "From Zlreginfo" & vbNewLine & _
        "Where 项目 In ('收集目录S', '访问用户S', '访问密码S', '收集目录F', '访问用户F', '访问密码F', '访问端口F')" & vbNewLine & _
        "And 内容 Is Not Null"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    rsTmp.Filter = "项目 = '收集目录F'"
    If Not rsTmp.EOF Then
        strUser = "": strPwd = "": strPort = "": strPath = ""
        strPath = rsNum!内容 & ""
        rsTmp.Filter = "项目='访问用户F'"
        If Not rsTmp.EOF Then strUser = rsTmp!内容 & ""
        rsTmp.Filter = "项目='访问密码F'"
        If Not rsTmp.EOF Then strPwd = rsTmp!内容 & ""
        rsTmp.Filter = "项目='访问端口F'"
        If Not rsTmp.EOF Then strPort = rsTmp!内容 & ""
        lngServerNO = lngServerNO + 1
        strSQL = "Zl_Zlupgradeserver_Update(0," & lngServerNO & ",1," & SQLAdjust(strPath) & "," & SQLAdjust(strUser) & "," & SQLAdjust(Cipher(strPwd)) & "," & Val(strPort) & ",0,0,1," & IIf(intClientUpType = 1, SQLAdjust(strFileType), "NULL") & "," & SQLAdjust(strPwd) & ")"
        Call ExecuteProcedure(strSQL, Me.Caption, gcnOracle)
    End If
    '处理共享收集
    rsTmp.Filter = "项目 = '收集目录F'"
    If Not rsTmp.EOF Then
        strTmp = Mid(rsNum!项目, Len("服务器目录") + 1)
        strUser = "": strPwd = "": strPath = ""
        strPath = rsNum!内容 & ""
        rsTmp.Filter = "项目='访问用户" & strTmp & "'"
        If Not rsTmp.EOF Then strUser = rsTmp!内容 & ""
        rsTmp.Filter = "项目='访问密码" & strTmp & "'"
        If Not rsTmp.EOF Then strPwd = rsTmp!内容 & ""
        lngServerNO = lngServerNO + 1
        strSQL = "Zl_Zlupgradeserver_Update(0," & lngServerNO & ",0," & SQLAdjust(strPath) & "," & SQLAdjust(strUser) & "," & SQLAdjust(Cipher(strPwd)) & ",NULL,0,0,1," & IIf(intClientUpType = 0, SQLAdjust(strFileType), "NULL") & "," & SQLAdjust(strPwd) & ")"
        Call ExecuteProcedure(strSQL, Me.Caption, gcnOracle)
    End If
    If lngServerNO > 0 Then
        '清理旧数据
        '1-清理FTP数据
        strSQL = "Delete From Zlreginfo" & vbNewLine & _
            "Where (项目 Like 'FTP服务器%' And 项目 > 'FTP服务器0')" & vbNewLine & _
            "Or (项目 Like 'FTP用户%' And 项目 > 'FTP用户0')" & vbNewLine & _
            "Or (项目 Like 'FTP密码%' And 项目 > 'FTP密码0')" & vbNewLine & _
            "Or (项目 Like 'FTP端口%' And 项目 > 'FTP端口0')"
        Call gclsBase.ExecuteCmdText(strSQL, Me.Caption, gcnOracle)
        '2-清理共享数据
        strSQL = "Delete From Zlreginfo" & vbNewLine & _
            "Where (项目 Like '服务器目录%' And 项目 > '服务器目录0')" & vbNewLine & _
            "Or (项目 Like '访问用户%' And 项目 > '访问用户0')" & vbNewLine & _
            "Or (项目 Like '访问密码%' And 项目 > '访问密码0')"
        Call gclsBase.ExecuteCmdText(strSQL, Me.Caption, gcnOracle)
        '3-清理收集数据
        strSQL = "Delete From Zlreginfo" & vbNewLine & _
                    "Where 项目 In ('收集方式'," & vbNewLine & _
                    "             '收集类型'," & vbNewLine & _
                    "             '收集目录S'," & vbNewLine & _
                    "             '访问用户S'," & vbNewLine & _
                    "             '访问密码S'," & vbNewLine & _
                    "             '收集目录F'," & vbNewLine & _
                    "             '访问用户F'," & vbNewLine & _
                    "             '访问密码F'," & vbNewLine & _
                    "             '访问端口F')"
        Call gclsBase.ExecuteCmdText(strSQL, Me.Caption, gcnOracle)
        '4-清理ZLClients升级设置
        strSQL = "Update Zlclients Set 升级服务器 = Null, Ftp服务器 = Null Where 升级文件服务器 Is Null"
        Call gclsBase.ExecuteCmdText(strSQL, Me.Caption, gcnOracle)
    End If
    TransOldData = True
    Exit Function
ErrH:
    TransOldData = False
    If 0 = 1 Then
        Resume
    End If
    MsgBox "旧版本服务器数据转换失败, 请联系开发人员!信息：" & err.Description, vbInformation, gstrSysName
    err.Clear
End Function

Public Sub SetControlEnable(ByVal strProgFunc As String)
'根据权限字符串设置控件状态
'strProgFunc:权限字符串
    Dim arrFunc() As String
    Dim i As Long
    
    mblnAllowEdit = False
    arrFunc = Split(strProgFunc, "|")
    For i = 0 To UBound(arrFunc)
        If arrFunc(i) = "文件服务器配置" Then
            mblnAllowEdit = True
        End If
    Next
    '若没有权限，则将一些控件设为不可用
    If mblnAllowEdit = False Then
        cmdAdd.Enabled = False
        cmdModify.Enabled = False
        cmdDel.Enabled = False
        chkSampleServer.Enabled = False
        vsfMain.Editable = flexEDNone
    End If
End Sub
