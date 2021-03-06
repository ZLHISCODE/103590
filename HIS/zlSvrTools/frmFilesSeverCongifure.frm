VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmFilesSeverConfigure 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "文件服务器配置"
   ClientHeight    =   5220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraOption 
      BorderStyle     =   0  'None
      Height          =   252
      Left            =   120
      TabIndex        =   11
      Top             =   4200
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
      Left            =   135
      ScaleHeight     =   336
      ScaleWidth      =   5388
      TabIndex        =   10
      Top             =   180
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
      ScaleHeight     =   252
      ScaleWidth      =   3996
      TabIndex        =   9
      Top             =   3930
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
      Height          =   2835
      Left            =   120
      TabIndex        =   8
      Top             =   700
      Width           =   6870
      _cx             =   12118
      _cy             =   5001
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
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFilesSeverCongifure.frx":0000
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
      ExplorerBar     =   5
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
End
Attribute VB_Name = "frmFilesSeverConfigure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrSelectSeverNum As String '定位行服务器编号
Private mblnFirstAdd As Boolean '第一次添加服务器需要设置为默认服务器
Public blnRefreshData As Boolean '界面切换刷新判断标志

Private Enum ServerListCols
    Col_编号 = 0 '状态值 0-正常 1-升级部件缺失(本地文件必不存在) 2-本地文件不存在 3-无需更新 4-警告但可以上传 5-已经上传
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
    Col_服务器列表列数 = 12
End Enum

Private Const SS_停用 = "停用"
Private Const SS_启用 = "启用"
Private Const ST_FTP = "FTP"
Private Const ST_共享 = "共享"

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口
End Sub


Private Sub chkSampleServer_Click()
    Dim strSQL As String
    
    If chkSampleServer.Tag <> "" Then
        strSQL = "Update Zlreginfo Set 内容 = '" & chkSampleServer.value & "' Where 项目 = 'FTP不检查文件存在'"
        gcnOracle.Execute strSQL, , adCmdText
    End If
End Sub

Private Sub cmdAdd_Click()
    Dim frmEdit As New frmFilesSeverEdit
    
    If frmEdit.ShowMe(0, mblnFirstAdd) Then
        LoadSeverListData
        vsfMain.Row = vsfMain.Rows - 1
        vsfMain.SetFocus
    End If
End Sub

Private Sub cmdCheck_Click()
    Dim strSeverAddress As String
    Dim strUser As String
    Dim strPassword As String
    Dim strPort As String
    Dim lngRowsCount As Long
    Dim strInformation As String
    Dim i As Long
    
    With vsfMain
        If .Rows < .FixedRows Then Exit Sub
        lngRowsCount = .Rows - 1
        .Row = 0
        cmdDel.Enabled = False: cmdModify.Enabled = False
        For i = .FixedRows To lngRowsCount
            ShowFlash "正在检测" & .TextMatrix(i, Col_编号) & "号: " & .TextMatrix(i, Col_服务器路径), i / (lngRowsCount), Me, True
            DoEvents
            If .TextMatrix(i, Col_服务器状态) = SS_停用 Then
                .TextMatrix(i, Col_检测结果) = "不可用：" & "该服务器处于停用状态，请启用后尝试连接校验。"
            Else
                strSeverAddress = Trim(.TextMatrix(i, Col_服务器路径))
                strUser = Trim(.TextMatrix(i, Col_用户名))
                strPassword = Trim(.Cell(flexcpData, i, Col_密码))
                strPort = Trim(.TextMatrix(i, Col_端口))
                
                If Trim(.TextMatrix(i, Col_类型)) = ST_FTP Then
                    If CheckFTPServer(strSeverAddress, strUser, strPassword, strPort, strInformation) = False Then
                        .TextMatrix(i, Col_检测结果) = "不可用：" & strInformation
                    Else
                        .TextMatrix(i, Col_检测结果) = "可用"
                    End If
                Else
                    If CheckFileServer(strSeverAddress, strUser, strPassword, strInformation) = False Then
                        .TextMatrix(i, Col_检测结果) = "不可用：" & strInformation
                    Else
                        .TextMatrix(i, Col_检测结果) = "可用"
                    End If
                End If
            End If
        Next
        ShowFlash ("")
    End With
End Sub

Private Sub cmdDel_Click()
    Dim strSQL As String
    If vsfMain.TextMatrix(vsfMain.Row, Col_是否缺省) <> "" Then
        MsgBox vsfMain.TextMatrix(vsfMain.Row, Col_编号) & " 号" & vsfMain.TextMatrix(vsfMain.Row, Col_类型) & "服务器为缺省服务器不能删除，请切换缺省服务器后删除", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("确定要删除 " & vsfMain.TextMatrix(vsfMain.Row, Col_编号) & " 号" & vsfMain.TextMatrix(vsfMain.Row, Col_类型) & "服务器？", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
        strSQL = "Zl_Zlupgradeserver_Delete('" & vsfMain.TextMatrix(vsfMain.Row, Col_编号) & "')"
        Call ExecuteProcedure(strSQL, Me.Caption)
        strSQL = "update ZLClients set 升级文件服务器 = null where 升级文件服务器 = " & vsfMain.TextMatrix(vsfMain.Row, Col_编号)
        gcnOracle.Execute strSQL
'        Load frmUpgradeManage
        
        Call LoadSeverListData
        vsfMain.SetFocus
    End If
End Sub

Private Sub cmdModify_Click()
    Dim frmEdit As New frmFilesSeverEdit

    If frmEdit.ShowMe(1, mblnFirstAdd, Nvl(vsfMain.TextMatrix(vsfMain.Row, 0), "")) Then
        LoadSeverListData
    End If
    
End Sub

Private Sub Form_Load()
    '加载服务器数据信息
    If TransData = False Then MsgBox "旧版本服务器数据转换失败！请联系开发人员！", vbInformation, gstrSysName
'    Call RefreshData
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fraOption.Top = Me.ScaleHeight - fraOption.Height - 60
    vsfMain.Move 50, 650, Me.ScaleWidth - 120, fraOption.Top - vsfMain.Top - 120
    picBtn.Move 50, 210
    If Me.ScaleWidth < 8000 Then picFilter.Visible = False
    If Me.ScaleWidth >= 8000 Then picFilter.Visible = True
    picFilter.Move Me.ScaleWidth - picFilter.Width, 280
End Sub

Public Sub SetMenu()
    frmMDIMain.stbThis.Panels(2).Text = "列表中共显示有" & vsfMain.Rows - 1 & "行数据。"
End Sub

Public Sub LoadSeverListData(Optional ByVal strFilter As String, Optional ByVal strLocationName As String)
    Dim i, j As Long
    Dim strSQL       As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngLocationRow As Long
    
    On Error GoTo errH

    mblnFirstAdd = True
    strSQL = "Select 内容 As 使用简易ftp工具 From Zlreginfo Where 项目 =[1]"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, "FTP不检查文件存在")
    If rsTemp.EOF Then
        chkSampleServer.value = 0
        strSQL = "Insert Into Zlreginfo (项目, 内容) Values ('FTP不检查文件存在', '0')"
        gcnOracle.Execute strSQL, , adCmdText
    Else
        chkSampleServer.value = Val(rsTemp!使用简易ftp工具 & "")
    End If
    chkSampleServer.Tag = "数据已经加载"
    With vsfMain

        If .Row < .FixedRows Then .Row = 0
        lngLocationRow = .Row
    
        .Redraw = flexRDNone
        .Rows = .FixedRows
'        .Clear
'        .Cols = Col_服务器列表列数

'        strSQL = "select 编号,类型,位置,用户名,密码,端口,是否升级,是否缺省,是否收集,收集类型 from ZLUpgradeServer order by 编号"
        strSQL = "select 编号,类型,位置,用户名,密码,端口,是否升级,是否缺省 from ZLUpgradeServer order by 编号"
        Call OpenRecordset(rsTemp, strSQL, Me.Caption)

        '数据填入
        .Rows = rsTemp.RecordCount + 1
        i = .FixedRows
        Do Until rsTemp.EOF
        
            .TextMatrix(i, Col_编号) = Nvl(rsTemp.Fields("编号"), "")
            .TextMatrix(i, Col_类型) = IIf(Nvl(rsTemp.Fields("类型"), "") = "1", ST_FTP, ST_共享)
            .TextMatrix(i, Col_服务器路径) = Nvl(rsTemp.Fields("位置"), "")
            .TextMatrix(i, Col_用户名) = Nvl(rsTemp.Fields("用户名"), "")
            .TextMatrix(i, Col_密码) = "***"
            .Cell(flexcpData, i, Col_密码) = Decipher(Nvl(rsTemp.Fields("密码"), ""))
            .TextMatrix(i, Col_端口) = Nvl(rsTemp.Fields("端口"), "")
            .Cell(flexcpBackColor, i, Col_是否升级, i, Col_是否收集) = RGB(210, 240, 255) 'RGB(247, 247, 247)
            .TextMatrix(i, Col_是否升级) = IIf(Nvl(rsTemp.Fields("是否升级"), "") = "1", "√", "")
            .TextMatrix(i, Col_是否缺省) = IIf(Nvl(rsTemp.Fields("是否缺省"), "") = "1", "√", "")
            If .TextMatrix(i, Col_是否缺省) = "√" Then
                mblnFirstAdd = False
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlue
            End If
'            .TextMatrix(i, Col_是否收集) = IIf(Nvl(rsTemp.Fields("是否收集"), "") = "1", "√", "")
            .TextMatrix(i, Col_检测结果) = ""
'            .TextMatrix(i, Col_收集类型) = Nvl(rsTemp.Fields("收集类型"), "")
            
            If .TextMatrix(i, Col_是否升级) = "" And .TextMatrix(i, Col_是否缺省) = "" And .TextMatrix(i, Col_是否收集) = "" Then
                .TextMatrix(i, Col_服务器状态) = SS_停用
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbGrayText
            Else
                .Cell(flexcpText, i, Col_服务器状态) = SS_启用
            End If

            rsTemp.MoveNext
            i = i + 1
        Loop
        
        '选中框风格
        .FocusRect = flexFocusSolid
        '最后一列自动列宽
        .ExtendLastCol = True
        '滚动画面跟随
        .ScrollTrack = True
        '自动换行
        .WordWrap = True
        '行高设置
        .RowHeightMin = 300
        .RowHeightMax = 300
        '最大宽度设置
        .ColWidthMax = 7000
        '自动适应行高、列宽
        .AutoSizeMode = flexAutoSizeRowHeight
        .SelectionMode = flexSelectionListBox
        .AllowBigSelection = False
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False
        
        If lngLocationRow > .Rows - 1 Then lngLocationRow = .Rows - 1
        .Row = lngLocationRow
        .Redraw = flexRDBuffered
        
        Call SetMenu
        
    End With
    Exit Sub
errH:
    Call MsgBox("服务器列表加载错误", vbInformation, gstrSysName)
    If False Then
        Resume
    End If
End Sub

Private Sub optFilter_Click(Index As Integer)
Dim i As Long
    With vsfMain
        If .Rows < 1 Then Exit Sub
        .Redraw = flexRDNone
        For i = 1 To .Rows - 1
            Select Case Index
            Case 0
                .RowHidden(i) = False
            Case 1
                .RowHidden(i) = .TextMatrix(i, Col_服务器状态) = SS_停用
            Case 2
                .RowHidden(i) = .TextMatrix(i, Col_服务器状态) = SS_启用
            End Select
        Next
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub vsfMain_AfterSort(ByVal Col As Long, Order As Integer)
    vsfMain.Row = vsfMain.FindRow(mstrSelectSeverNum, , Col_编号)
End Sub

Private Sub vsfMain_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        If Row = 0 Then Cancel = True
End Sub

Private Sub vsfMain_DblClick()
    Dim strIsUpgrade As String
    Dim strIsCheck As String
    Dim strIsCollect As String
    Dim strFilesType As String
    Dim strSQL As String
    On Error GoTo errHand
    
    With vsfMain
        If .MouseRow <> .Row Then Exit Sub
        strIsUpgrade = IIf(.TextMatrix(.Row, Col_是否升级) = "√", "1", "0")
        strIsCheck = IIf(.TextMatrix(.Row, Col_是否缺省) = "√", "1", "0")
        strIsCollect = IIf(.TextMatrix(.Row, Col_是否收集) = "√", "1", "0")
        strFilesType = .TextMatrix(.Row, Col_收集类型)

        Select Case .ColSel
        Case Col_是否升级
            If strIsCheck = "1" Then
                Call MsgBox("选中编号 " & .TextMatrix(.Row, Col_编号) & " 服务器为缺省服务器，请保证至少有一个缺省服务器？ ", vbInformation, gstrSysName)
                Exit Sub
            ElseIf strIsCollect = "1" Then
                If MsgBox("选中编号 " & .TextMatrix(.Row, Col_编号) & " 服务器为收集服务器，是否要切换为升级服务器？ ", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                    If mblnFirstAdd = True Then
                        strFilesType = ""
                        strIsUpgrade = "1"
                        strIsCheck = "1"
                        strIsCollect = "0"
                    Else
                        strFilesType = ""
                        strIsUpgrade = "1"
                        strIsCollect = "0"
                    End If
                End If
            ElseIf mblnFirstAdd = True Then
                strIsUpgrade = "1"
                strIsCheck = "1"
            Else
                If strIsUpgrade = "1" Then
                    If MsgBox("是否要取消该升级服务器，取消后将会清空已设置过该服务器为升级服务器的客户端", vbOKCancel, gstrSysName) = vbOK Then
                        strIsUpgrade = "0"
                    End If
                Else
                    strIsUpgrade = "1"
                End If
            End If
        Case Col_是否缺省
            If strIsCheck = "1" Then
                Call MsgBox("选中编号 " & .TextMatrix(.Row, Col_编号) & " 服务器为默认服务器，请保证至少有一个默认服务器 ", vbInformation, gstrSysName)
                Exit Sub
            ElseIf strIsCollect = "1" Then
                If MsgBox("选中编号 " & .TextMatrix(.Row, Col_编号) & " 服务器为收集服务器，是否要切换为升级服务器并设置为缺省服务器？ ", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                    strFilesType = ""
                    strIsUpgrade = "1"
                    strIsCheck = "1"
                    strIsCollect = "0"
                End If
            ElseIf strIsUpgrade = "0" Then
                If MsgBox("选中编号 " & .TextMatrix(.Row, Col_编号) & " 服务器为停用状态，是否要启用该服务器并设置为缺省服务器？ ", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                    strFilesType = ""
                    strIsUpgrade = "1"
                    strIsCheck = "1"
                    strIsCollect = "0"
                End If
            Else
                strIsCheck = IIf(strIsCheck = "0", "1", "0")
            End If
        Case Col_是否收集
            If strIsCheck = "1" Then
                Call MsgBox("选中编号 " & .TextMatrix(.Row, Col_编号) & " 服务器为缺省默认升级服务器,不能切换为收集服务器！", vbInformation, gstrSysName)
                Exit Sub
            ElseIf strIsUpgrade = "1" Then
                If MsgBox("选中编号 " & .TextMatrix(.Row, Col_编号) & " 服务器为升级服务器，是否要切换为收集服务器？ ", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                    If .TextMatrix(.Row, Col_收集类型) = "" Then strFilesType = "Log"
                    strIsUpgrade = "0"
                    strIsCheck = "0"
                    strIsCollect = "1"
                End If
            Else
                strIsCollect = IIf(strIsCollect = "1", "0", "1")
            End If
        Case Else
            Exit Sub
        End Select

        strSQL = "Zl_Zlupgradeserver_Update('" & .TextMatrix(.Row, Col_编号) & "','','','','','','" & strIsUpgrade & "','" & strIsCheck & "','" & strIsCollect & "','" & strFilesType & "','" & 1 & "')"
        Call ExecuteProcedure(strSQL, Me.Caption)
        
        If strIsUpgrade = "0" Then
            strSQL = "update ZLClients set 升级文件服务器 = null where 升级文件服务器 = " & .TextMatrix(.Row, Col_编号)
            gcnOracle.Execute strSQL
        End If
        
        If strIsCheck = "1" Then
            strSQL = "ZLReginfo_DefaultServer('" & IIf(.TextMatrix(.Row, Col_类型) = ST_共享, "0", "1") & "','" & Trim(.TextMatrix(.Row, Col_服务器路径)) & "','" & Trim(.TextMatrix(.Row, Col_用户名)) & "','" & Trim(.Cell(flexcpData, .Row, Col_密码)) & "','" & Trim(.TextMatrix(.Row, Col_端口)) & "')"
            Call ExecuteProcedure(strSQL, Me.Caption)
        End If

        LoadSeverListData
        optFilter.Item(0).value = True
    End With
    
    Exit Sub
errHand:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub vsfMain_RowColChange()
    mstrSelectSeverNum = vsfMain.TextMatrix(vsfMain.Row, Col_编号)
    If vsfMain.Row > 0 Then cmdModify.Enabled = True: cmdDel.Enabled = True: cmdCheck.Enabled = True
End Sub

Private Function CheckFTPServer(ByVal strIp As String, ByVal strUser As String, ByVal strPass As String, ByVal strPort As String, Optional ByRef strError As String) As Boolean
    '-----------------------------------------------------------------------------
    '功能:检查当前的FTP服务器是否正确
    '返回:当前的文件服务器的各项正确,返回true,否则返回False
    '编制:陈振原
    '日期:2016/07/05
    'strIp - FTP地址
    'strUser - 用户名
    'strPass - 密码
    'strPort - 端口
    '-----------------------------------------------------------------------------
    On Error GoTo errHand:
    
    If strIp = "" Or strUser = "" Or strPass = "" Or strPort = "" Then
        CheckFTPServer = False
        Exit Function
    End If
    
    If IsFtpServer(Trim(strIp), Trim(strUser), Trim(strPass), Trim(strPort)) Then
        CheckFTPServer = True
        strError = "连接成功"
    Else
        CheckFTPServer = False
        strError = "不能连接升级服务器，请检查FTP服务器配置"
    End If
    CancelFtpServer
    Exit Function
    
errHand:
        MsgBox err.Description, vbInformation, gstrSysName
End Function


Private Function CheckFileServer(ByVal strAddress As String, ByVal strUser As String, ByVal strPass As String, Optional ByRef strError As String) As Boolean
    '-----------------------------------------------------------------------------
    '功能:检查当前的文件服务器是否正确
    '返回:当前的文件服务器的各项正确,返回true,否则返回False
    '编制:陈振原
    '日期:2016/07/05
    'strAddress - 地址
    'strUser - 用户
    'strPass - 密码
    '-----------------------------------------------------------------------------
    Dim typOfStruct As OFSTRUCT

    On Error GoTo errHand:
    
    If strAddress = "" Or strUser = "" Or strPass = "" Then
        CheckFileServer = False
        Exit Function
    End If
    
    If IsNetServer(Trim(strAddress), Trim(strUser), Trim(strPass)) = False Then
        strError = "升级文件的指定目录不存在,请重新设置"
        CheckFileServer = False
    Else
        strError = "连接成功"
        CheckFileServer = True
    End If
    Call CancelNetServer(Trim(strAddress))
    
    Exit Function
errHand:
        MsgBox err.Description, vbInformation, gstrSysName
End Function

Private Function FindFile(ByVal strFileName As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------
    '--功能:查找指定的文件或文夹是否存在
    '--返回: 如果存在此文件为True,否则为Flase
    '------------------------------------------------------------------------------------------------------------------------------------
    Dim typOfStruct As OFSTRUCT
    
    On Error Resume Next
    FindFile = False
    If Len(strFileName) > 0 Then
        apiOpenFile strFileName, typOfStruct, OF_EXIST
        FindFile = typOfStruct.nErrCode <> 2
    End If
End Function

Private Function TransData() As Boolean
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim strCollectType As String
    Dim intNoNum As Integer
    Dim intSeverNum As Integer
    Dim intUpType As Integer
    Dim i As Long
    
    On Error GoTo errH
    strSQL = "select * from zlupgradeserver"
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
    If Not rsTemp.EOF Then TransData = True: Exit Function
'    If MsgBox("是否将旧版本设置过的服务器数据转换至新版本服务器设置？", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then TransData = True: Exit Function
    
    '清空客户端曾经配置的升级服务器
    strSQL = "update zlclients set 升级文件服务器 = null,升级服务器 = null,FTP服务器=null"
    gcnOracle.Execute strSQL
    
    '处理FTP 0-N 服务器数据
    strSQL = "select max(nvl(replace(项目,'FTP服务器',''),'-1')) as 服务器数 from zlreginfo where 项目 like 'FTP服务器%'"
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
        
    intSeverNum = Nvl(rsTemp.Fields("服务器数"), -1)
        
    If intSeverNum >= 0 Then
        For i = 0 To intSeverNum
            strSQL = "Select Max(Decode(项目, 'FTP端口', 内容, '')) FTP端口, Max(Decode(项目, 'FTP服务器', 内容, '')) FTP服务器," & vbNewLine & _
                        "       Max(Decode(项目, 'FTP密码', 内容, '')) FTP密码, Max(Decode(项目, 'FTP用户', 内容, '')) FTP用户" & vbNewLine & _
                        "From (Select Substr(项目, Length(项目), 1) ID, Substr(项目, 1, Length(项目) - 1) 项目, 内容" & vbNewLine & _
                        "       From zlRegInfo" & vbNewLine & _
                        "       Where 项目 = 'FTP服务器" & i & "' or 项目 = 'FTP用户" & i & "' or 项目 = 'FTP密码" & i & "' or 项目 = 'FTP端口" & i & "')" & vbNewLine & _
                        "Group By ID"
            Call OpenRecordset(rsTemp, strSQL, Me.Caption)
                
            If rsTemp.EOF = False Then
                If Nvl(rsTemp.Fields("FTP服务器"), "") <> "" And Nvl(rsTemp.Fields("FTP用户"), "") <> "" And Nvl(rsTemp.Fields("FTP密码"), "") <> "" And Nvl(rsTemp.Fields("FTP端口"), "") <> "" Then
                    intNoNum = intNoNum + 1
                    strSQL = "Zl_Zlupgradeserver_Insert('" & intNoNum & "','" & 1 & "','" & rsTemp.Fields("FTP服务器") & "','" & rsTemp.Fields("FTP用户") & "','" & Cipher(rsTemp.Fields("FTP密码")) & "','" & rsTemp.Fields("FTP端口") & "','" & 1 & "','" & 0 & "','" & 0 & "','')"
                    Call ExecuteProcedure(strSQL, Me.Caption)
                End If
            Else
                
            End If
        Next
    End If
        
    '处理FTP服务器老数据
    strSQL = "Select Max(Decode(项目, 'FTP端口', 内容, '')) FTP端口, Max(Decode(项目, 'FTP服务器', 内容, '')) FTP服务器," & vbNewLine & _
                "       Max(Decode(项目, 'FTP密码', 内容, '')) FTP密码, Max(Decode(项目, 'FTP用户', 内容, '')) FTP用户" & vbNewLine & _
                "From (Select Substr(项目, Length(项目), 1) ID, Substr(项目, 1, Length(项目) - 1) 项目, 内容" & vbNewLine & _
                "       From zlRegInfo" & vbNewLine & _
                "       Where (项目 = 'FTP服务器' or 项目 = 'FTP用户' or 项目 = 'FTP密码' or 项目 = 'FTP端口') And Not 内容 Is Null)" & vbNewLine & _
                "Group By ID"
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
            
    If rsTemp.EOF = False Then
        If Nvl(rsTemp.Fields("FTP服务器"), "") <> "" And Nvl(rsTemp.Fields("FTP用户"), "") <> "" And Nvl(rsTemp.Fields("FTP密码"), "") <> "" And Nvl(rsTemp.Fields("FTP端口"), "") <> "" Then
            intNoNum = intNoNum + 1
            strSQL = "Zl_Zlupgradeserver_Insert('" & intNoNum & "','" & 1 & "','" & rsTemp.Fields("FTP服务器") & "','" & rsTemp.Fields("FTP用户") & "','" & Cipher(rsTemp.Fields("FTP密码")) & "','" & rsTemp.Fields("FTP端口") & "','" & 1 & "','" & 0 & "','" & 0 & "','')"
            Call ExecuteProcedure(strSQL, Me.Caption)
        End If
    Else
    End If

    '处理 共享 0-N 服务器数据
    strSQL = "select max(nvl(replace(项目,'服务器目录',''),'-1')) as 服务器数 from zlreginfo where 项目 like '服务器目录%'"
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
    intSeverNum = Nvl(rsTemp.Fields("服务器数"), -1)
        
    If intSeverNum >= 0 Then
        For i = 0 To intSeverNum
            strSQL = "Select Max(Decode(项目, '服务器目录', 内容, '')) 服务器目录, Max(Decode(项目, '访问用户', 内容, '')) 访问用户," & vbNewLine & _
                        "       Max(Decode(项目, '访问密码', 内容, '')) 访问密码" & vbNewLine & _
                        "From (Select Substr(项目, Length(项目), 1) ID, Substr(项目, 1, Length(项目) - 1) 项目, 内容" & vbNewLine & _
                        "       From zlRegInfo" & vbNewLine & _
                        "       Where 项目 = '服务器目录" & i & "' Or 项目 = '访问用户" & i & "' Or 项目 = '访问密码" & i & "')" & vbNewLine & _
                        "Group By ID"
            Call OpenRecordset(rsTemp, strSQL, Me.Caption)
                
            If rsTemp.EOF = False Then
                If Nvl(rsTemp.Fields("服务器目录"), "") <> "" And Nvl(rsTemp.Fields("访问用户"), "") <> "" And Nvl(rsTemp.Fields("访问密码"), "") <> "" Then
                    intNoNum = intNoNum + 1
                    strSQL = "Zl_Zlupgradeserver_Insert('" & intNoNum & "','" & 0 & "','" & rsTemp.Fields("服务器目录") & "','" & rsTemp.Fields("访问用户") & "','" & Cipher(rsTemp.Fields("访问密码")) & "','','" & 1 & "','" & 0 & "','" & 0 & "','')"
                    Call ExecuteProcedure(strSQL, Me.Caption)
                End If
            End If
        Next
    End If
    
    '处理 共享 服务器老数据
    strSQL = "Select Max(Decode(项目, '服务器目录', 内容, '')) 服务器目录, Max(Decode(项目, '访问用户', 内容, '')) 访问用户," & vbNewLine & _
                "       Max(Decode(项目, '访问密码', 内容, '')) 访问密码" & vbNewLine & _
                "From (Select 1 ID, 项目, 内容" & vbNewLine & _
                "       From zlRegInfo" & vbNewLine & _
                "       Where (项目 = '服务器目录' Or 项目 = '访问用户' Or 项目 = '访问密码') And Not 内容 Is Null)" & vbNewLine & _
                "Group By ID"
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
            
    If rsTemp.EOF = False Then
        If Nvl(rsTemp.Fields("服务器目录"), "") <> "" And Nvl(rsTemp.Fields("访问用户"), "") <> "" And Nvl(rsTemp.Fields("访问密码"), "") <> "" Then
            intNoNum = intNoNum + 1
            strSQL = "Zl_Zlupgradeserver_Insert('" & intNoNum & "','" & 0 & "','" & rsTemp.Fields("服务器目录") & "','" & rsTemp.Fields("访问用户") & "','" & Cipher(rsTemp.Fields("访问密码")) & "','','" & 1 & "','" & 0 & "','" & 0 & "','')"
            Call ExecuteProcedure(strSQL, Me.Caption)
        End If
    End If

    '收集服务器处理
'    strSQL = "select 内容 as 收集类型 from zlreginfo where 项目 = '收集类型'"
'    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
'    strCollectType = Nvl(rsTemp.Fields("收集类型"), "")
'
'    strSQL = "Select Max(Decode(项目, '收集目录S', 内容, '')) 收集目录, Max(Decode(项目, '访问用户S', 内容, '')) 访问用户," & vbNewLine & _
'                "       Max(Decode(项目, '访问密码S', 内容, '')) 访问密码, Max(Decode(项目, '收集类型', 内容, '')) 收集类型" & vbNewLine & _
'                "From (Select 1 As ID, 项目, 内容" & vbNewLine & _
'                "       From zlRegInfo" & vbNewLine & _
'                "       Where (项目 = '收集目录S' Or 项目 = '访问用户S' Or 项目 = '访问密码S') And Not 内容 Is Null)" & vbNewLine & _
'                "Group By ID"
'    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
'
'    If rsTemp.EOF = False Then
'        intNoNum = intNoNum + 1
'        strSQL = "Zl_Zlupgradeserver_Insert('" & intNoNum & "','" & 0 & "','" & Nvl(rsTemp.Fields("访问用户"), "无效") & "','" & Nvl(rsTemp.Fields("访问密码")) & "','" & Cipher(Nvl(rsTemp.Fields("访问密码"), "无效")) & "','','" & 0 & "','" & 0 & "','" & 1 & "','" & strCollectType & "')"
'        Call ExecuteProcedure(strSQL, Me.Caption)
'    Else
'    End If
'
'    strSQL = "Select Max(Decode(项目, '收集目录F', 内容, '')) 收集目录, Max(Decode(项目, '访问用户F', 内容, '')) 访问用户," & vbNewLine & _
'                "       Max(Decode(项目, '访问密码F', 内容, '')) 访问密码, Max(Decode(项目, '收集类型', 内容, '')) 收集类型" & vbNewLine & _
'                "From (Select 1 As ID, 项目, 内容" & vbNewLine & _
'                "       From zlRegInfo" & vbNewLine & _
'                "       Where (项目 = '收集目录F' Or 项目 = '访问用户F' Or 项目 = '访问密码F') And Not 内容 Is Null)" & vbNewLine & _
'                "Group By ID"
'    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
'
'    If rsTemp.EOF = False Then
'        intNoNum = intNoNum + 1
'        strSQL = "Zl_Zlupgradeserver_Insert('" & intNoNum & "','" & 1 & "','" & Nvl(rsTemp.Fields("访问用户"), "无效") & "','" & Nvl(rsTemp.Fields("访问密码")) & "','" & Cipher(Nvl(rsTemp.Fields("访问密码"), "无效")) & "','','" & 0 & "','" & 0 & "','" & 1 & "','" & strCollectType & "')"
'        Call ExecuteProcedure(strSQL, Me.Caption)
'    Else
'    End If
        
    '删除旧数据
'    strSQL = "delete from zlreginfo where 项目 like 'FTP%'or 项目 like '访问%' or 项目 like '服务器目录%'or 项目 like '收集目录%'"
'    gcnOracle.Execute strSQL
        
    '设置默认服务器
    If intNoNum > 0 Then
        strSQL = "select max(内容) as 升级类型 from zlreginfo where 项目 = '升级类型'"
        Call OpenRecordset(rsTemp, strSQL, Me.Caption)
        
        intUpType = Nvl(rsTemp.Fields("升级类型"), 0)

        strSQL = "select 编号,类型,位置,用户名,密码,端口 from zlupgradeserver where 编号 = (select min(编号) from (select 编号 from zlupgradeserver where 类型 = " & intUpType & "))"
        Call OpenRecordset(rsTemp, strSQL, Me.Caption)
            
        If rsTemp.EOF Then
            strSQL = "select 编号,类型,位置,用户名,密码,端口 from zlupgradeserver where 编号 = 1"
            Call OpenRecordset(rsTemp, strSQL, Me.Caption)
        End If
        
        strSQL = "Zl_Zlupgradeserver_Update('" & rsTemp.Fields("编号") & "','','','','','','" & 1 & "','" & 1 & "','" & 0 & "','','" & 1 & "')"
        Call ExecuteProcedure(strSQL, Me.Caption)
        strSQL = "ZLReginfo_DefaultServer('" & rsTemp.Fields("类型") & "','" & rsTemp.Fields("位置") & "','" & rsTemp.Fields("用户名") & "','" & Decipher(rsTemp.Fields("密码")) & "','" & rsTemp.Fields("端口") & "')"
        Call ExecuteProcedure(strSQL, Me.Caption)
    End If
    strSQL = "Select 内容 As 使用简易ftp工具 From Zlreginfo Where 项目 =[1]"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, "FTP不检查文件存在")
    If rsTemp.EOF Then
        chkSampleServer.value = 0
        strSQL = "Insert Into Zlreginfo (项目, 内容) Values ('FTP不检查文件存在', '0')"
        gcnOracle.Execute strSQL, , adCmdText
    Else
        chkSampleServer.value = Val(rsTemp!使用简易ftp工具 & "")
    End If
    chkSampleServer.Tag = "数据已经加载"
    TransData = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    TransData = False
    If False Then
        Resume
    End If
End Function

Public Sub RefreshData()
    Call LoadSeverListData
End Sub

