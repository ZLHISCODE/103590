VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmClientsUpgrade 
   BackColor       =   &H80000005&
   Caption         =   "站点部件升级"
   ClientHeight    =   5544
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   10704
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmClientsUpgrade.frx":0000
   ScaleHeight     =   5550
   ScaleMode       =   0  'User
   ScaleWidth      =   10704
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboUpResult 
      Appearance      =   0  'Flat
      Height          =   276
      Left            =   6315
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   1182
      Width           =   1000
   End
   Begin VB.CommandButton cmdClearUpLog 
      Caption         =   "清除所有升级日志"
      Height          =   330
      Left            =   7320
      TabIndex        =   18
      ToolTipText     =   "新一次升级时,可以重新设置各站点的升级状态为""未升级"""
      Top             =   1155
      Width           =   1668
   End
   Begin VB.PictureBox Piccmb 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   2520
      ScaleHeight     =   252
      ScaleWidth      =   924
      TabIndex        =   17
      Top             =   1185
      Width           =   945
      Begin VB.ComboBox cboFind 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -30
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   -30
         Width           =   1000
      End
   End
   Begin VB.CommandButton cmdUpdateS 
      Caption         =   "重置升级状态"
      Height          =   330
      Left            =   9000
      TabIndex        =   5
      ToolTipText     =   "新一次升级时,可以重新设置各站点的升级状态为""未升级"""
      Top             =   1155
      Width           =   1425
   End
   Begin VB.CommandButton cmd用户密码设置 
      Caption         =   "管理设置(&J)"
      Height          =   330
      Left            =   6696
      TabIndex        =   10
      ToolTipText     =   "客户端为User权限,升级时使用的管理员用户、密码设置"
      Top             =   5125
      Width           =   1200
   End
   Begin VB.CommandButton cmd预升级设置 
      Caption         =   "预升级设置(&K)"
      Height          =   330
      Left            =   7872
      TabIndex        =   11
      Top             =   5125
      Width           =   1300
   End
   Begin VB.OptionButton OptType 
      BackColor       =   &H80000005&
      Caption         =   "FTP"
      Height          =   180
      Index           =   1
      Left            =   4725
      TabIndex        =   16
      Top             =   5200
      Width           =   810
   End
   Begin VB.OptionButton OptType 
      BackColor       =   &H80000005&
      Caption         =   "文件共享"
      Height          =   180
      Index           =   0
      Left            =   3705
      TabIndex        =   9
      Top             =   5200
      Value           =   -1  'True
      Width           =   1065
   End
   Begin VB.CommandButton cmd应用 
      Caption         =   "应用于本列(&P)…"
      Height          =   350
      Left            =   900
      TabIndex        =   7
      Top             =   5115
      Width           =   1575
   End
   Begin VB.CommandButton cmd保存 
      Caption         =   "保存设置(&O)"
      Height          =   350
      Left            =   2475
      TabIndex        =   8
      Top             =   5115
      Width           =   1200
   End
   Begin VSFlex8Ctl.VSFlexGrid vsClients 
      Height          =   3396
      Left            =   156
      TabIndex        =   0
      Top             =   1512
      Width           =   10260
      _cx             =   18098
      _cy             =   5990
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
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483643
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmClientsUpgrade.frx":04F9
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
      ExplorerBar     =   7
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
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   810
      TabIndex        =   2
      Text            =   "255.255.255.255"
      Top             =   1185
      Width           =   1680
   End
   Begin VB.PictureBox picSel 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3705
      ScaleHeight     =   288
      ScaleWidth      =   1200
      TabIndex        =   15
      Top             =   75
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "刷新(&R)"
      Height          =   350
      Left            =   105
      TabIndex        =   6
      Top             =   5115
      Width           =   795
   End
   Begin MSComctlLib.ImageList ilsIcon 
      Left            =   5565
      Top             =   -210
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientsUpgrade.frx":06C5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "部件配置(&L)"
      Height          =   330
      Left            =   9180
      TabIndex        =   12
      Top             =   5125
      Width           =   1200
   End
   Begin VB.CheckBox chkAllSel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "当前全部站点升级(&A)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3480
      TabIndex        =   4
      Top             =   1230
      Width           =   2040
   End
   Begin VB.Label lblUpResult 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "升级状态"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5520
      TabIndex        =   19
      Top             =   1230
      Width           =   720
   End
   Begin VB.Label lblFind 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "查找(&Z)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   1
      Top             =   1230
      Width           =   630
   End
   Begin VB.Image imgMain 
      Height          =   384
      Left            =   156
      Picture         =   "frmClientsUpgrade.frx":118F
      Top             =   612
      Width           =   384
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   2925
      Top             =   0
      _Version        =   589884
      _ExtentX        =   508
      _ExtentY        =   508
      _StockProps     =   0
   End
   Begin VB.Label lblNote 
      BackStyle       =   0  'Transparent
      Caption         =   "对站点部件升级进行设置和收集文件服务系统的最新部件信息。可通过双击客户端查看升级情况。"
      Height          =   348
      Left            =   828
      TabIndex        =   14
      Top             =   648
      Width           =   5112
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "站点部件升级"
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
      TabIndex        =   13
      Top             =   105
      Width           =   1440
   End
End
Attribute VB_Name = "frmClientsUpgrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Const mMenu_Popu = 1
Private Const mMenu_Popu_ClientName = 11
Private Const mMenu_Popu_ClientIP = 12
Private Const mMenu_Popu_ClientDept = 13
Private Const mMenu_Popu_ClientUser = 14

Dim mintColumn As Integer
Private mintType As Integer     '11-按站点名进行过滤,12-按IP过滤,13-按部门过滤,14-按用途过滤
Private mrsClients As ADODB.Recordset
Private mrsFileServer As ADODB.Recordset
Private mrsFilePreUpgrade As ADODB.Recordset '预升级记录集
Private mblnChange As Boolean '发生了改变
Private mblnTypeChange As Boolean '升级方式发生改变
Private mintUpType     As Integer  '0 共享方式 1 FTP方式'
Private mblnLoad       As Boolean '是否已经加载完毕
Private Enum UpgradeState
    US_未升级 = 0
    US_成功 = 1
    US_失败 = 2
    US_升级中 = 3
    US_所有 = 4
End Enum
Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口
End Sub

Private Sub cboFind_Click()
    Select Case cboFind.ListIndex
        Case 0
            mintType = 11
            txtSearch.Tag = "请输入工作站名称"
        Case 1
            mintType = 13
            txtSearch.Tag = "请输入部门名称"
        Case 2
            mintType = 12
            txtSearch.Tag = "请输入IP地址"
        Case 3
            mintType = 14
            txtSearch.Tag = "请输入用途"
    End Select
    txtSearch.Text = txtSearch.Tag
End Sub

Private Sub cboUpResult_Click()
    If mblnLoad Then
        LoadClientsInfor
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
        Dim objControl As CommandBarControl
        Dim objPopu As CommandBarPopup
        
        Select Case Control.Id
        Case mMenu_Popu_ClientName '站点名称
            mintType = Control.Id
            Set objPopu = cbsMain.FindControl(, mMenu_Popu)
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientName).Checked = False
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientIP).Checked = False
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientDept).Checked = False
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientUser).Checked = False
            Control.Checked = True
            txtSearch.Tag = Split(Control.Caption, "(")(0)
            Call PrintSearch(txtSearch.Tag, vbBlue, False)
            If txtSearch.Enabled Then txtSearch.SetFocus
            Call LoadClientsInfor
        Case mMenu_Popu_ClientIP   'IP
            mintType = Control.Id
            Set objPopu = cbsMain.FindControl(, mMenu_Popu)
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientName).Checked = False
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientIP).Checked = False
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientDept).Checked = False
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientUser).Checked = False
            Control.Checked = True
            txtSearch.Tag = Split(Control.Caption, "(")(0)
            Call PrintSearch(txtSearch.Tag, vbBlue, False)
            If txtSearch.Enabled Then txtSearch.SetFocus
            Call LoadClientsInfor
        Case mMenu_Popu_ClientDept '部门名称
            mintType = Control.Id
            Set objPopu = cbsMain.FindControl(, mMenu_Popu)
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientName).Checked = False
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientIP).Checked = False
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientDept).Checked = False
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientUser).Checked = False
            Control.Checked = True
            txtSearch.Tag = Split(Control.Caption, "(")(0)
            Call PrintSearch(txtSearch.Tag, vbBlue, False)
            If txtSearch.Enabled Then txtSearch.SetFocus
            Call LoadClientsInfor
        Case mMenu_Popu_ClientUser  '用途
            mintType = Control.Id
            Set objPopu = cbsMain.FindControl(, mMenu_Popu)
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientName).Checked = False
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientIP).Checked = False
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientDept).Checked = False
            objPopu.CommandBar.FindControl(, mMenu_Popu_ClientUser).Checked = False
            Control.Checked = True
            txtSearch.Tag = Split(Control.Caption, "(")(0)
            Call PrintSearch(txtSearch.Tag, vbBlue, False)
            If txtSearch.Enabled Then txtSearch.SetFocus
            Call LoadClientsInfor
        End Select
End Sub

Private Sub chkAllSel_Click()
    Dim i As Long
    If chkAllSel.Tag = "T" Then chkAllSel.Tag = "": Exit Sub
    With vsClients
        .Cell(flexcpChecked, 1, .ColIndex("升级"), .Rows - 1, .ColIndex("升级")) = IIf(Me.chkAllSel.value = 1, flexChecked, flexUnchecked)
    End With
    mblnChange = True
    Call SetCtlEnabled
End Sub

Private Sub cmdClearUpLog_Click()
    Dim strSQL As String
    If MsgBox("确认要清除所有升级日志吗?", vbYesNo + vbInformation, gstrSysName) = vbYes Then
        '删除升级日志
        strSQL = "delete zltools.zlClientUpdatelog"
        gcnOracle.Execute strSQL
    End If
End Sub

Private Sub cmdFile_Click()
    Dim blnReturn As Boolean
    If OptType(0).value Then
        Call frmFilesSet.ShowEdit(Me, blnReturn)
        If blnReturn = False Then Exit Sub
        '暂无其他操作
        gstrSQL = "Select 项目,内容 From zlRegInfo where  项目 like '服务器目录%' or 项目 like '访问用户%' or 项目 like '访问密码%'"
        Call OpenRecordset(mrsFileServer, gstrSQL, Me.Caption)
        Call initVsGrid
        LoadClientsInfor (True)
    Else
        Call frmFilesFTPSet.ShowEdit(Me, blnReturn)
        If blnReturn = False Then Exit Sub
        gstrSQL = "Select 项目,内容 From zlRegInfo where  项目 like 'FTP服务器%' or 项目 like 'FTP用户%' or 项目 like 'FTP密码%'"
        Call OpenRecordset(mrsFileServer, gstrSQL, Me.Caption)
        Call initVsGrid
        LoadClientsInfor (True)
    End If
End Sub
Private Sub cmdRefresh_Click()
    '初始化信息
    Call LoadClientsInfor(True)
End Sub

Private Sub cmdUpdateS_Click()
    Dim lngRet As Long
    Dim i As Long
    Dim strName As String
    Dim strSQL As String
    
    lngRet = MsgBox("新的一次升级时,可以重新设置各站点的升级状态为[未升级]" & vbNewLine & "确定要重置选择站点的升级状态吗?", vbYesNo + vbInformation, "提示")
    If lngRet = vbYes Then
        With vsClients
            For i = .Row To .RowSel
                '根据光标选中行来
                'If Val(.TextMatrix(i, .ColIndex("升级"))) = -1 Then
                    strName = .TextMatrix(i, .ColIndex("工作站"))
                    strSQL = "Zl_Zlclients_Control(6,'" & strName & "')"
                    Call ExecuteProcedure(strSQL, Me.Caption)
                    
                    '删除升级日志
                    strSQL = "delete zltools.zlClientUpdatelog where 工作站='" & UCase(strName) & "'"
                    gcnOracle.Execute strSQL
                'End If
            Next
            Call LoadClientsInfor(True)  '刷新列表
        End With
    End If
End Sub

Private Sub Cmd保存_Click()
    If mblnChange Then
        If SaveData = False Then
            MsgBox "升级站点配置失败!", vbInformation, gstrSysName
            Exit Sub
        Else
            MsgBox "升级站点配置成功!", vbInformation, gstrSysName
        End If
    End If
    
    If mblnTypeChange Then
        Call SaveUpType
        
        If mintUpType = 0 Then
            
            gstrSQL = "Select 项目,内容 From zlRegInfo where  项目 like '服务器目录%' or 项目 like '访问用户%' or 项目 like '访问密码%'"
            Call OpenRecordset(mrsFileServer, gstrSQL, Me.Caption)
            mrsFileServer.Filter = ""
            initVsGrid
            
        Else
            gstrSQL = "Select 项目,内容 From zlRegInfo where  项目 like 'FTP服务器%' or 项目 like 'FTP用户%' or 项目 like 'FTP密码%'"
            Call OpenRecordset(mrsFileServer, gstrSQL, Me.Caption)
            mrsFileServer.Filter = ""
            
            initVsGrid
        End If
    End If
    Call LoadClientsInfor(mblnTypeChange Or mblnChange)
    mblnTypeChange = False
    mblnChange = False
    Call SetCtlEnabled
End Sub

Private Sub cmd应用_Click()
    
    Dim i As Long
    Dim strKey As String
    With vsClients
        
    
        If .Col = .ColIndex("服务器") Then
            .Redraw = flexRDNone
            strKey = Trim(.TextMatrix(.Row, .Col))
            For i = 1 To .Rows - 1
                .TextMatrix(i, .Col) = strKey
            Next
            .Redraw = flexRDBuffered
        End If
        
        If .Col = .ColIndex("预升时点") Then
            .Redraw = flexRDNone
            strKey = Trim(.TextMatrix(.Row, .Col))
            For i = .Row To .RowSel
                .TextMatrix(i, .Col) = strKey
            Next
            
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("升级")) = -1 Then
                    .TextMatrix(i, .Col) = strKey
                End If
            Next
            
            .Redraw = flexRDBuffered
        End If
        
        If .Col = .ColIndex("预升完成") Then
            .Redraw = flexRDNone
            strKey = Trim(.TextMatrix(.Row, .Col))
            For i = .Row To .RowSel
                .TextMatrix(i, .Col) = strKey
                If strKey = "" Or strKey = "未完成" Then
                    .Cell(flexcpForeColor, i, .Col, i, .Col) = 0
                Else
                    .Cell(flexcpForeColor, i, .Col, i, .Col) = vbRed
                End If
            Next
            
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("升级")) = -1 Then
                    .TextMatrix(i, .Col) = strKey
                    
                    If strKey = "" Or strKey = "未完成" Then
                        .Cell(flexcpForeColor, i, .Col, i, .Col) = 0
                    Else
                        .Cell(flexcpForeColor, i, .Col, i, .Col) = vbRed
                    End If
                End If
            Next
            .Redraw = flexRDBuffered
        End If
        
    End With
End Sub

Private Sub cmd用户密码设置_Click()
    Load frmFilesUpgradeAdmin
    frmFilesUpgradeAdmin.Show 1, frmMDIMain
    If frmFilesUpgradeAdmin.mblnOk Then
    End If
    Exit Sub
End Sub

Private Sub cmd预升级设置_Click()
    Load frmFilesUpgradeTime
    frmFilesUpgradeTime.Show 1, frmMDIMain
    If frmFilesUpgradeTime.mblnOk Then
        '设置预升级时间点分配
        On Error GoTo errHandle
        Call ExecuteProcedure("Zl_Zlclients_Control(1)", Me.Caption)
        gstrSQL = "Select 项目,内容 From zlRegInfo where  项目 = '客户端预升级时间点'"
        Set mrsFilePreUpgrade = New ADODB.Recordset
        Call OpenRecordset(mrsFilePreUpgrade, gstrSQL, Me.Caption)
    End If
    Call initVsGrid
    LoadClientsInfor (True)
    Exit Sub
errHandle:
    MsgBox "保存失败。" & vbNewLine & err.Description, vbExclamation, gstrSysName
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        cmdRefresh_Click
    End If
End Sub

Private Sub Form_Load()
    mblnLoad = False
    mintType = mMenu_Popu_ClientName
'    Call PrintSearch("按工作站搜索", vbBlue, False)
'    txtSearch.Tag = "按工作站搜索"
    '初始升级方式
    Call InitUpType
    
    'mintUpType =0 共享方式
    If mintUpType = 0 Then
        gstrSQL = "Select 项目,内容 From zlRegInfo where  项目 like '服务器目录%' or 项目 like '访问用户%' or 项目 like '访问密码%'"
    Else
        gstrSQL = "Select 项目,内容 From zlRegInfo where  项目 like 'FTP服务器%' or 项目 like 'FTP用户%' or 项目 like 'FTP密码%'"
    End If
    Set mrsFileServer = New ADODB.Recordset
    Call OpenRecordset(mrsFileServer, gstrSQL, Me.Caption)
    
    
    gstrSQL = "Select 项目,内容 From zlRegInfo where  项目 = '客户端预升级时间点'"
    Set mrsFilePreUpgrade = New ADODB.Recordset
    Call OpenRecordset(mrsFilePreUpgrade, gstrSQL, Me.Caption)
    '查找功能初始化
    cboFind.AddItem "工作站", 0
    cboFind.AddItem "部门", 1
    cboFind.AddItem "IP", 2
    cboFind.AddItem "用途", 3
    cboFind.ListIndex = 0
    
    cboUpResult.AddItem "未升级", US_未升级
    cboUpResult.AddItem "成功", US_成功
    cboUpResult.AddItem "失败", US_失败
    cboUpResult.AddItem "升级中", US_升级中
    cboUpResult.AddItem "所有", US_所有
    cboUpResult.ListIndex = US_所有
    
    txtSearch.ForeColor = vbGrayText
    '初始菜单
    Call InitCommandBar
    '初始化网格的相关属性
    Call initVsGrid
    mblnLoad = True
    '初始化信息
    Call LoadClientsInfor(True)
    Call RestoreGridSet

    mblnChange = False
End Sub

Private Sub RestoreGridSet()
    '---------------------------------------------------------------------------------
    '功能:恢复个性化设置
    '编制:刘兴宏
    '日期:2007/09/10
    '---------------------------------------------------------------------------------
    Dim i As Long
    Dim strColumns As String
    Dim arrColumn As Variant
    Dim arrValue As Variant
    err = 0: On Error GoTo errHand:
    '恢复个性化设置
    strColumns = ""
    strColumns = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Caption, "站点", "")
    
    If strColumns <> "" Then
        arrColumn = Split(strColumns, "|")
        With vsClients
            For i = 0 To UBound(arrColumn)
                arrValue = Split(arrColumn(i), ",")
                .ColWidth(.ColIndex(arrValue(0))) = Val(arrValue(1))
                .ColPosition(.ColIndex(arrValue(0))) = i
            Next
        End With
    End If
errHand:
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.lblNote.Width = ScaleWidth - Me.lblNote.Left
    
    With cmdUpdateS
         .Left = ScaleWidth - .Width - 50
    End With
    
    With cmdClearUpLog
         .Left = cmdUpdateS.Left - .Width - 5
    End With
    
    Call SetCtrlPosOnLine(False, chkAllSel, 30, lblUpResult, 15, cboUpResult)
    
    With cmdRefresh
        .Top = ScaleHeight - .Height - 50
        cmd应用.Top = .Top
        cmd保存.Top = .Top
    End With
    
    With cmdFile
        .Top = cmdRefresh.Top
        .Left = ScaleWidth - .Width - 50
    End With
    
    With cmd预升级设置
        .Top = cmdFile.Top
        .Left = cmdFile.Left - .Width - 5
    End With
    
    With cmd用户密码设置
        .Top = cmdFile.Top
        .Left = cmd预升级设置.Left - .Width - 5
    End With
    
    With vsClients
        .Width = ScaleWidth - .Left - 50
        .Height = cmdRefresh.Top - .Top - 50
    End With
    With picSel
        .Left = vsClients.Left
    End With
    
    With OptType(0)
        .Left = cmd保存.Left + cmd保存.Width + 200
        .Top = cmd保存.Top + 75
    End With
    
    With OptType(1)
        .Left = OptType(0).Left + OptType(0).Width + 50
        .Top = cmd保存.Top + 75
    End With
End Sub

Private Sub SetCtlEnabled()
    '---------------------------------------------------------------------------------------------
    '功能：设置控件的相关属性
    '参数：
    '返回：
    '编制：刘兴宏
    '日期：2007/09/07
    '---------------------------------------------------------------------------------------------
    
    Dim blnNoClients As Boolean '没有站点
    Dim i As Long, bln应用 As Boolean
    blnNoClients = True
    With vsClients
        bln应用 = (.Col = .ColIndex("服务器")) Or (.Col = .ColIndex("预升时点")) Or (.Col = .ColIndex("预升完成"))
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("工作站"))) <> "" Then
                blnNoClients = False
                Exit For
            End If
        Next
    End With
    chkAllSel.Enabled = Not blnNoClients
    cmd应用.Enabled = bln应用
    cmd保存.Enabled = mblnChange
End Sub

Private Sub LoadClientsInfor(Optional blnRefresh As Boolean)
    '---------------------------------------------------------------------------------------------
    '功能：加载站点信息
    '参数：blnFilter-是否通过现有的记录进行过滤
    '返回：
    '编制：刘兴宏
    '日期:2007/08/20
    '---------------------------------------------------------------------------------------------
    Dim strSQL As String, strFilter As String, strKey As String
    Dim i As Long
    Dim StrDate As String
    Dim lngColore   As Long
    
    err = 0: On Error GoTo errHand:
    
    strSQL = "" & _
    "   Select 工作站, ip, cpu, 内存, 硬盘, 操作系统, 部门,zlspellcode(部门) as 部门简码, 用途,Decode(升级情况,0,'未升级',1,'成功',2,'失败',3,'升级中') as 升级情况,升级情况 as 升级状态, 说明, " & _
    "           升级标志, 收集标志, 禁止使用, 连接数, 升级服务器,FTP服务器,预升时点,预升完成" & _
    "   From zlClients" & _
    "   Order by IP"
    
    
    If blnRefresh = True Or mrsClients Is Nothing Then
        Set mrsClients = New ADODB.Recordset
        Call OpenRecordset(mrsClients, strSQL, Me.Caption)
        'Set rsClients = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Client", "")
    ElseIf mrsClients.State <> 1 Then
        Call OpenRecordset(mrsClients, strSQL, Me.Caption)
    End If
    strKey = txtSearch.Text
    If strKey <> "" And strKey <> txtSearch.Tag Or cboUpResult.ListIndex <> US_所有 Then
        If strKey <> "" And strKey <> txtSearch.Tag Then
            Select Case mintType
                Case 12     '-按IP过滤
                    strFilter = "IP like '" & strKey & "%'"
                Case 13     '-按部门过滤
                    strFilter = "部门 like '" & strKey & "%' OR 部门简码 like '" & UCase(strKey) & "%'"
                Case 14     '按用途过滤
                    strFilter = "用途 like '" & strKey & "%'"
                Case Else           ' 11-按站点名进行过滤
                    strFilter = "工作站 like '" & UCase(strKey) & "%'"
            End Select
        End If
        
        If cboUpResult.ListIndex <> US_所有 Then
            If strFilter <> "" Then
                If mintType = 13 Then
                    strFilter = "(部门 like '" & strKey & "%' And 升级状态=" & cboUpResult.ListIndex & ") OR (部门简码 like '" & UCase(strKey) & "%' And 升级状态=" & cboUpResult.ListIndex & ")"
                Else
                    strFilter = strFilter & " And 升级状态=" & cboUpResult.ListIndex
                End If
            Else
                strFilter = "升级状态=" & cboUpResult.ListIndex
            End If
        End If
        mrsClients.Filter = strFilter
    Else
        mrsClients.Filter = 0
    End If
    
    With vsClients
        .Redraw = flexRDNone
        .Rows = IIf(mrsClients.RecordCount = 0, 1, mrsClients.RecordCount) + 1
        If mrsClients.RecordCount <> 0 Then mrsClients.MoveFirst
        If mrsClients.EOF Then
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
            Next
            .Redraw = flexRDBuffered
            SetCtlEnabled
            mrsClients.Filter = 0
            Exit Sub
        End If
        '有数据
        i = 1
        Do While Not mrsClients.EOF
            lngColore = 0
            If mintUpType = 0 Then
                strKey = Val(Nvl(mrsClients!升级服务器))
            Else
                strKey = Val(Nvl(mrsClients!FTP服务器))
            End If
'            strKey = IIf(Val(strKey) = 0, "", strKey)
            If mintUpType = 0 Then
                mrsFileServer.Find "项目='服务器目录" & strKey & "'", , , 1
            Else
                If strKey = "" Then strKey = "0"
                mrsFileServer.Find "项目='FTP服务器" & strKey & "'", , , 1
            End If
            If mrsFileServer.EOF = False Then
                .TextMatrix(i, .ColIndex("服务器")) = Val(strKey) & ":" & Nvl(mrsFileServer!内容)
            Else
                .TextMatrix(i, .ColIndex("服务器")) = Val(strKey) & ":"
            End If
            .Cell(flexcpData, i, .ColIndex("服务器")) = Val(strKey)
            
            If mrsFilePreUpgrade.EOF = False Then
                If Nvl(mrsClients!预升时点) = "" Then
                    StrDate = ""
                Else
                    StrDate = Format(Nvl(mrsClients!预升时点), "HH:mm")
                End If
                .TextMatrix(i, .ColIndex("预升时点")) = StrDate
            Else
                .TextMatrix(i, .ColIndex("预升时点")) = ""
            End If
            .TextMatrix(i, .ColIndex("工作站")) = Nvl(mrsClients!工作站)
            .TextMatrix(i, .ColIndex("IP")) = Nvl(mrsClients!IP)
            .TextMatrix(i, .ColIndex("CPU")) = Nvl(mrsClients!cpu)
            .TextMatrix(i, .ColIndex("内存")) = Nvl(mrsClients!内存)
            .TextMatrix(i, .ColIndex("硬盘")) = Nvl(mrsClients!硬盘)
            .TextMatrix(i, .ColIndex("操作系统")) = Nvl(mrsClients!操作系统)
            .TextMatrix(i, .ColIndex("部门")) = Nvl(mrsClients!部门)
            .TextMatrix(i, .ColIndex("用途")) = Nvl(mrsClients!用途)
            .TextMatrix(i, .ColIndex("升级情况")) = Nvl(mrsClients!升级情况)
            If Nvl(mrsClients!升级状态, 0) = 3 Then
                lngColore = vbGreen '绿色
            ElseIf Nvl(mrsClients!升级状态, 0) = 2 Then
                lngColore = vbRed '红色
            ElseIf Nvl(mrsClients!升级状态, 0) = 1 Then
                lngColore = vbBlue '蓝色
            End If
            .Cell(flexcpForeColor, i, .ColIndex("工作站"), i, .ColIndex("IP")) = lngColore
            '使用颜色标识预升是否完成!
            If Nvl(mrsClients!预升完成, 0) = 1 Then
                .TextMatrix(i, .ColIndex("预升完成")) = "完成"
                .Cell(flexcpForeColor, i, .ColIndex("预升完成"), i, .ColIndex("预升完成")) = vbRed
            Else
                .TextMatrix(i, .ColIndex("预升完成")) = "未完成"
                .Cell(flexcpForeColor, i, .ColIndex("预升完成"), i, .ColIndex("预升完成")) = 0
            End If
            
            .TextMatrix(i, .ColIndex("说明")) = Nvl(mrsClients!说明)
            If Val(Nvl(mrsClients!升级标志)) = 1 Then
                .Cell(flexcpChecked, i, .ColIndex("升级")) = flexChecked
            Else
                .Cell(flexcpChecked, i, .ColIndex("升级")) = flexUnchecked
            End If
            i = i + 1
            mrsClients.MoveNext
        Loop
        .Redraw = flexRDBuffered
    End With
    mrsClients.Filter = 0
    SetCtlEnabled
    Exit Sub
errHand:
   ' Resume
    vsClients.Redraw = flexRDBuffered
    MsgBox "系统出现错误,错误为:" & err.Description, vbInformation + vbDefaultButton1, gstrSysName
    SetCtlEnabled
    Exit Sub
End Sub

Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------
    '功能:保存升级信息
    '参数:str工作站-工作站
    '     bln升级标志
    '     str服务器号
    '返回:成功返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/09/07
    '---------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim i As Long
    Dim str工作站 As String, int升级 As Integer, str服务器号 As String
    Dim str预升时点 As String
    Dim int预升完成 As Integer
    Dim strIp As String
    err = 0: On Error GoTo errHand:
    
    With vsClients
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("工作站"))) <> "" Then
                str工作站 = Trim(.TextMatrix(i, .ColIndex("工作站")))
                int升级 = IIf(.Cell(flexcpChecked, i, .ColIndex("升级")) = flexChecked, 1, 0)
                strIp = Trim(.TextMatrix(i, .ColIndex("IP")))
                If Trim(.TextMatrix(i, .ColIndex("服务器"))) = "" Then
                    str服务器号 = "0"
                Else
                    str服务器号 = Val(Split(Trim(.TextMatrix(i, .ColIndex("服务器"))), ":")(0))
                End If
                
                If Trim(.TextMatrix(i, .ColIndex("预升时点"))) = "" Then
                    str预升时点 = "NULL"
                Else
                    str预升时点 = Trim(.TextMatrix(i, .ColIndex("预升时点")))
                    str预升时点 = Format(Now(), "yyyy-MM-dd") & " " & Format(str预升时点, "hh:mm:00")
                    str预升时点 = "to_date('" & str预升时点 & "','YYYY-MM-DD HH24:MI:SS')"
                End If
                
                If Trim(.TextMatrix(i, .ColIndex("预升完成"))) = "" Then
                    int预升完成 = 0
                Else
                    If Trim(.TextMatrix(i, .ColIndex("预升完成"))) = "未完成" Then
                        int预升完成 = 0
                    Else
                        int预升完成 = 1
                    End If
                End If
                
                If mintUpType = 0 Then
                    strSQL = "Zl_Zlclients_Control(2,Null,'" & strIp & "'," & int升级 & "," & str服务器号 & "," & str预升时点 & "," & int预升完成 & ")"
                Else
                    strSQL = "Zl_Zlclients_Control(2,Null,'" & strIp & "'," & int升级 & ",Null," & str预升时点 & "," & int预升完成 & "," & str服务器号 & ")"
                End If
                
                Call ExecuteProcedure(strSQL, Me.Caption)
            End If
        Next
    End With
    SaveData = True
    Exit Function
errHand:
    MsgBox "保存升级信息时出错,错误信息如下:" & vbCrLf & "错误号:" & err.Number & vbCrLf & "错误描述:" & err.Description, vbInformation, gstrSysName
'''    Resume
End Function

Private Sub Form_Unload(Cancel As Integer)
    '保存个性化设置
    Dim i As Long
    Dim strColumns As String
    strColumns = ""
    With vsClients
        For i = 0 To .Cols - 1
            strColumns = strColumns & "|" & .ColKey(i) & "," & .ColWidth(i)
        Next
    End With
    If strColumns <> "" Then strColumns = Mid(strColumns, 2)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Caption, "站点", strColumns
End Sub

Private Sub OptType_Click(Index As Integer)
    mblnTypeChange = True
    cmd保存.Enabled = True
    mintUpType = Index
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
    Dim objPopup As CommandBarPopup
    Set objPopup = cbsMain.FindControl(xtpControlPopup, mMenu_Popu, , True)
    If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
    
    Call PrintSearch(Me.txtSearch.Tag, vbBlue, False)
    picSel.Tag = ""
End Sub

Private Sub PrintSearch(ByVal strTittle As String, ByVal lngColor As Long, ByVal blnBoderStyle As Boolean)
    '----------------------------------------------------------------------------------------
    '功能:打印指定的索引条件
    '参数:strTittle-标题
    '     lngColor-颜色值
    '     lngBoderStyl-是否加边框线
    '----------------------------------------------------------------------------------------
 
    With picSel
        .Cls
        .ForeColor = lngColor
        .FontUnderline = True
        .CurrentX = 30 '(.ScaleWidth - .TextWidth(strTittle))
        .CurrentY = (.ScaleHeight - .TextHeight(strTittle)) / 2
        picSel.Print strTittle
        .ZOrder 1
    End With
End Sub

Private Sub txtSearch_Change()
    If mblnLoad Then
        If txtSearch.Text = txtSearch.Tag Then Exit Sub
        If mblnChange = True Then
            If MsgBox("站点升级信息被你编辑过,是否保存你编辑的信息?", vbQuestion + vbYesNo + vbQuestion) = vbYes Then
                Call SaveData
            End If
            mblnChange = False
        End If
        LoadClientsInfor
    End If
End Sub

Private Sub txtSearch_GotFocus()
    If txtSearch.ForeColor = vbGrayText Then
        txtSearch.Text = ""
        txtSearch.ForeColor = vbBlack
    End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub InitCommandBar()
    '-------------------------------------------------------------------------------------------
    '功能:初始化菜单
    '参数:
    '返回:
    '编制:刘兴宏
    '日期:2007/08/07
    '-------------------------------------------------------------------------------------------
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objDeptBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    
    Set cbsMain.Icons = frmPubIcons.imgPublic.Icons
    
    '菜单定义:包括公共部份
    '    请对xtpControlPopup类型的命令ID重新赋值
    '-----------------------------------------------------
    cbsMain.ActiveMenuBar.Title = "弹出菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mMenu_Popu, "弹出菜单(&P)", -1, False)
    objMenu.Id = mMenu_Popu
    objMenu.Visible = False
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, mMenu_Popu_ClientName, "按工作站搜索(&0)"): objControl.Id = mMenu_Popu_ClientName: objControl.IconId = 102: objControl.Checked = True
        Set objControl = .Add(xtpControlButton, mMenu_Popu_ClientIP, "    按IP搜索(&1)"): objControl.Id = mMenu_Popu_ClientIP: objControl.IconId = 102
        Set objControl = .Add(xtpControlButton, mMenu_Popu_ClientDept, "  按部门搜索(&2)"): objControl.Id = mMenu_Popu_ClientDept: objControl.IconId = 102
        Set objControl = .Add(xtpControlButton, mMenu_Popu_ClientUser, "  按用途搜索(&3)"): objControl.Id = mMenu_Popu_ClientUser: objControl.IconId = 102
    End With
 End Sub
 Private Sub initVsGrid()
    '----------------------------------------------------------------------------------------
    '功能:初始化站点网格的相关设置
    '----------------------------------------------------------------------------------------
    With vsClients
        .Editable = flexEDKbdMouse
        .ColComboList(.ColIndex("服务器")) = Get服务器
        .ColComboList(.ColIndex("预升时点")) = Get预升时点
        If .ColComboList(.ColIndex("预升时点")) = "" Then
            .ColHidden(.ColIndex("预升时点")) = True
        Else
            .ColHidden(.ColIndex("预升时点")) = False
        End If
        .ColComboList(.ColIndex("预升完成")) = Get预升完成
    End With
 End Sub
 
 Private Function Get服务器() As String
    Dim strCombox As String
    Dim strTemp As String
    strCombox = ""
    With mrsFileServer
        If mintUpType = 0 Then
            .Filter = "项目 like '服务器目录%'"
            Do While Not .EOF
                strTemp = Replace(Nvl(!项目), "服务器目录", "")
                strCombox = strCombox & "|" & Val(strTemp) & ":" & Nvl(!内容)
                .MoveNext
            Loop
        Else
            .Filter = "项目 like 'FTP服务器%'"
            Do While Not .EOF
                strTemp = Replace(Nvl(!项目), "FTP服务器", "")
                strCombox = strCombox & "|" & Val(strTemp) & ":" & Nvl(!内容)
                .MoveNext
            Loop
        End If

    End With
    If strCombox <> "" Then strCombox = Mid(strCombox, 2)
    Get服务器 = strCombox
 End Function
 
 Private Function Get预升时点() As String
    Dim strTemp As String
    If mrsFilePreUpgrade.RecordCount = 1 Then
        mrsFilePreUpgrade.MoveFirst
        strTemp = Replace(Nvl(mrsFilePreUpgrade!内容), ",", "|")
    Else
        strTemp = ""
    End If
    
    Get预升时点 = strTemp
 End Function
 
 Private Function Get预升完成() As String
    Get预升完成 = "未完成|完成"
 End Function
 
Private Sub txtSearch_LostFocus()
    If txtSearch.Text = "" Then
        txtSearch.Text = txtSearch.Tag
        txtSearch.ForeColor = vbGrayText
    End If
End Sub

Private Sub vsClients_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
      With vsClients
        Select Case Col
        Case vsClients.ColIndex("服务器")
            mblnChange = True
            Call SetCtlEnabled
        Case vsClients.ColIndex("预升时点")
            mblnChange = True
            Call SetCtlEnabled
        Case vsClients.ColIndex("预升完成")
            mblnChange = True
            If vsClients.TextMatrix(Row, Col) = "" Or vsClients.TextMatrix(Row, Col) = "未完成" Then
              vsClients.Cell(flexcpForeColor, Row, Col, Row, Col) = 0
            Else
              vsClients.Cell(flexcpForeColor, Row, Col, Row, Col) = vbRed
            End If
            Call SetCtlEnabled
        Case vsClients.ColIndex("升级")
            chkAllSel.Tag = "T"
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, .ColIndex("升级")) = flexChecked Then
'                    MsgBox "第" & i & "行"
                    Exit For
                End If
            Next
            If i = .Rows Then
                chkAllSel.value = 0
            Else
                chkAllSel.value = 2
            End If
            mblnChange = True
            Call SetCtlEnabled
        Case Else
        End Select
    End With
End Sub

Private Sub vsClients_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case vsClients.ColIndex("服务器"), vsClients.ColIndex("升级"), vsClients.ColIndex("预升时点"), vsClients.ColIndex("预升完成")
        '只有服务器列和升级列才能更改
    Case Else
        '其他列不能更改
        Cancel = True
    End Select
End Sub
  
Private Sub vsClients_DblClick()
    '查看升级情况
    Dim strName As String
    If vsClients.Row > 0 Then
        If vsClients.TextMatrix(vsClients.Row, vsClients.ColIndex("升级情况")) <> "未升级" Then
            strName = vsClients.TextMatrix(vsClients.Row, vsClients.ColIndex("工作站"))
            frmFilesUpgradeLogView.mstrName = strName
            Load frmFilesUpgradeLogView
            frmFilesUpgradeLogView.Show 1, frmMDIMain
        End If
    End If
    
    Exit Sub
errHandle:
        MsgBox "保存失败。" & vbNewLine & err.Description, vbExclamation, gstrSysName
End Sub

Private Sub vsClients_RowColChange()
    Call SetCtlEnabled
End Sub

Private Sub SaveUpType()
'----------------------------------------------------------------------------------------
'功能:修改升级方式信息
'----------------------------------------------------------------------------------------
    On Error GoTo errH
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim str项目 As String '项目
    Dim str内容 As String '内容
    Dim strSQLTemp As String
    str项目 = "升级类型"
    If OptType(0).value Then
        str内容 = "0"
    Else
        str内容 = "1"
    End If
    strSQL = " Select 项目,内容 From zlregInfo where 项目= '升级类型'"
    
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
    If rsTemp.EOF = True Then
'        gcnOracle.BeginTrans
        strSQLTemp = "insert into zlregInfo(项目,内容) values ('" & str项目 & "','" & str内容 & "')"
        gcnOracle.Execute strSQLTemp
'        gcnOracle.CommitTrans
    Else
'        gcnOracle.BeginTrans
        strSQLTemp = "delete zlRegInfo where 项目='" & str项目 & "'"
        gcnOracle.Execute strSQLTemp
        strSQLTemp = "insert into zlregInfo(项目,内容) values ('" & str项目 & "','" & str内容 & "')"
        gcnOracle.Execute strSQLTemp
'        gcnOracle.CommitTrans
    End If
    
    Exit Sub
errH:
    If err Then
        MsgBox "保存升级类型信息时出错,错误信息如下:" & vbCrLf & "错误号:" & err.Number & vbCrLf & "错误描述:" & err.Description, vbInformation, gstrSysName
    End If
End Sub

Private Sub InitUpType()
'----------------------------------------------------------------------------------------
'功能:初始升级方式信息
'----------------------------------------------------------------------------------------
    On Error GoTo errH
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    strSQL = " Select 项目,内容 From zlregInfo where 项目= '升级类型'"
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)

    If rsTemp.EOF = False Then
        strTemp = Nvl(rsTemp!内容, "0")
        If strTemp = "1" Then
             OptType(1).value = True
             mintUpType = 1
        Else
             OptType(0).value = True
             mintUpType = 0
        End If
    Else
        OptType(0).value = True
        mintUpType = 0
    End If
    Exit Sub
errH:
    If err Then
        MsgBox "初始化升级方式出错,错误信息如下:" & vbCrLf & "错误号:" & err.Number & vbCrLf & "错误描述:" & err.Description, vbInformation, gstrSysName
    End If
End Sub

