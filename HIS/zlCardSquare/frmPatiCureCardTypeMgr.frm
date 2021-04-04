VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPatiCureCardTypeMgr 
   Caption         =   "医疗卡类别管理"
   ClientHeight    =   7605
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9930
   Icon            =   "frmPatiCureCardTypeMgr.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   9930
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7245
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatiCureCardTypeMgr.frx":08CA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12435
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   1350
      Top             =   1890
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
            Picture         =   "frmPatiCureCardTypeMgr.frx":115E
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiCureCardTypeMgr.frx":1A38
            Key             =   "Stop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1440
      Top             =   2760
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
            Picture         =   "frmPatiCureCardTypeMgr.frx":2312
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiCureCardTypeMgr.frx":28AC
            Key             =   "Stop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   2235
      Left            =   1950
      TabIndex        =   1
      Top             =   1440
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   3942
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   225
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmPatiCureCardTypeMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar, mcbrComboxToolBar As CommandBar
Private mlngModule As Long, mstrPrivs As String
Private Const mstrLvw As String = "名称,2174.74,0,1;编码,799.9371,0,2;短名,629.8583,2,0;前缀文本,929.7639,2,0;" & _
    "卡号长度,989.8583,0,1;缺省,630,2,1;固定项,760,2,0;严格控制,1000,2,0;卡类别,800,2,0;存在帐户,1000,2,0;" & _
    "全退,1400,2,0;部件,1500,0,0;医疗卡费,1500,0,0;结算方式,1620,0,0;卡号密文,1000,0,0;启用,600,2,0;备注,2000,0,0;" & _
    "模糊查找,1000,2,0;是否制卡,1000,2,0;是否发卡,1000,2,0;是否写卡,1000,2,0;转帐及代扣,1200,2,0;刷卡,800,2,0;" & _
    "扫描卡,800,2,0;接触式读卡,1200,2,0;非接触读卡,1200,2,0;键盘类型,1000,2,0;必须持卡消费,1300,2,0;" & _
    "发送调用接口,1300,2,0;是否退款验卡,1300,2,0;证件,800,2,0;启用回车,1000,2,0;是否缺省退现,1000,2,0" '问题号:56508
Private mintColumn As Integer
Private mblnShowStop As Boolean

Private Sub LoadData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据
    '编制:刘兴洪
    '日期:2011-06-27 20:52:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim objItem As ListItem, lngCol As Long, varValue As Variant
    Dim str停用 As String, strKey As String
    
    On Error GoTo errHandle
    
    '"           nvl(密码长度,10) as 密码长度,nvl(密码长度限制,0) as 密码长度限制,nvl(密码规则,0) as 密码规则," & _
    '"名称,2000,0,1;编码,800,0,2;短名,500,0,0;前缀文本,800,0,0;卡号长度,800,0,1;缺省标志,400,0,1;是否固定,400,0,0;是否严格控制,1000,0,0;是否刷卡,1000,0,0;是否自制,1000,0,0;是否存在帐户,1400,0,0;是否全退,1400,0,0;部件,1500,0,0;特定项目,1000,0,0;结算方式,1000,0,0;卡号密文,1000,0,0;是否启用,1000,0,0;备注,2000,0,0"
    '问题号:56508
    '103310,李南春,2016/12/7:卡号后增加回车符位
    '90875:李南春,2016/1/21,增加医疗卡证件类型
    '77872,李南春,2014/9/15:是否支持转帐及代扣
    strSQL = "" & _
    "   Select A.Id, A.名称, A.编码, A.短名, A.前缀文本, A.卡号长度, decode(nvl(A.缺省标志,0),1,'√','') as 缺省,  " & _
    "           decode(nvl(A.是否固定,0),1,'√','') as 固定项, decode(nvl(A.是否严格控制,0),1,'√','') as  严格控制, " & _
    "           decode(nvl(A.是否自制,0),1,'院内卡','院外卡') as    卡类别," & _
    "           decode(nvl(A.是否存在帐户,0),1,'√','') as     存在帐户, decode(nvl(A.是否全退,0),1,'√','') as    全退," & _
    "           A.部件,C.名称 as 医疗卡费, A.结算方式,A.卡号密文, decode(nvl(A.是否启用,0),1,'√','') as   启用, A.备注,  " & _
    "           decode(nvl(A.是否模糊查找,0),1,'√','')  as 模糊查找," & _
    "           decode(nvl(A.是否制卡,0),1,'√','') as   是否制卡,decode(nvl(A.是否发卡,0),1,'√','') as   是否发卡,decode(nvl(A.是否写卡,0),1,'√','') as   是否写卡," & _
    "           decode(nvl(A.是否转帐及代扣,0),1,'√','') as   转帐及代扣, decode(nvl(A.是否证件,0),1,'√','') as   证件," & _
    "           decode(substr(nvl(A.读卡性质,'0000'),1,1),1,'√','') as   刷卡," & _
    "           decode(substr(nvl(A.读卡性质,'0000'),2,1),1,'√','') as   扫描卡," & _
    "           decode(substr(nvl(A.读卡性质,'0000'),3,1),1,'√','') as   接触式读卡," & _
    "           decode(substr(nvl(A.读卡性质,'0000'),4,1),1,'√','') as   非接触读卡," & _
    "           decode(nvl(A.键盘控制方式,0),0,'禁用',1,'数字',2,'字符','禁用') as  键盘类型, " & _
    "           decode(nvl(A.是否持卡消费,0),1,'√','') as 必须持卡消费, " & _
    "           decode(nvl(A.发送调用接口,0),1,'√','') as 发送调用接口, " & _
    "           Decode(Nvl(a.是否退款验卡,0),1,'√','') As 是否退款验卡, " & _
    "           decode(A.设备是否启用回车,1,'√','') as   启用回车, " & _
    "           decode(A.是否缺省退现,1,'√','') as   是否缺省退现 " & _
    "    From 医疗卡类别 A ,收费特定项目 B,收费项目目录 C" & _
    "    Where   A.特定项目=B.特定项目(+) and B.收费细目ID=C.ID(+) " & _
            IIf(mblnShowStop, "", " and Nvl(是否启用,0)=1") & _
    "    Order by A.编码"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not lvwMain.SelectedItem Is Nothing Then
        strKey = lvwMain.SelectedItem.Key
    End If
    lvwMain.ListItems.Clear
    Do While Not rsTemp.EOF
        If NVL(rsTemp!启用) = "√" Then
            str停用 = "Start"
        Else
            str停用 = "Stop"
        End If
        Set objItem = lvwMain.ListItems.Add(, "K" & rsTemp!id, rsTemp!名称, str停用, str停用)
        If str停用 = "Stop" Then objItem.ForeColor = RGB(255, 0, 0)
        objItem.Tag = IIf(NVL(rsTemp!启用) = "√", 0, 1) & "-" & IIf(NVL(rsTemp!固定项) = "√", 1, 0) & "-" & IIf(NVL(rsTemp!证件) = "√", 1, 0)
        '根据ListView的列名从数据库取数
        For lngCol = 2 To lvwMain.ColumnHeaders.count
            varValue = rsTemp(lvwMain.ColumnHeaders(lngCol).Text).value
            objItem.SubItems(lngCol - 1) = IIf(IsNull(varValue), "", varValue)
            If str停用 = "Stop" Then objItem.ListSubItems(lngCol - 1).ForeColor = RGB(255, 0, 0)
        Next
        rsTemp.MoveNext
    Loop
    
    If lvwMain.ListItems.count > 0 Then
        On Error Resume Next
        Set objItem = lvwMain.ListItems(strKey)
        If Err <> 0 Then
            Err.Clear
            Set objItem = lvwMain.ListItems(1)
            objItem.Selected = True
            objItem.EnsureVisible
        Else
            objItem.Selected = True
            objItem.EnsureVisible
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub InitLvwHead()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查单据中输入的负数数量及退回科室是否正确
    '编制:刘兴洪
    '日期:2011-06-28 00:48:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strText As String, i As Integer
    On Error GoTo Errhand
    lvwMain.Tag = "可变化的"
    '如果ListView的还未被设置，比如第一次使用，那就调用缺省的初始化
    If lvwMain.ColumnHeaders.count = 0 Then
        zlControl.LvwSelectColumns lvwMain, mstrLvw, True
    End If
    strText = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\zl9CardSquare\" & Me.Name & "\ListView", lvwMain.Name & "名称")
    For i = 1 To lvwMain.ColumnHeaders.count
        '如果添加了列，则不恢复个性化
        If InStr(strText, lvwMain.ColumnHeaders(i).Text) = 0 Then lvwMain.Tag = "": Exit For
        '如果减少了列，也不恢复个性化
        strText = Replace(strText, lvwMain.ColumnHeaders(i).Text, "")
    Next
    strText = Replace(strText, ",", "")
    If strText <> "" Then lvwMain.Tag = ""
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub zlCardStopAndResume(Optional blnStop As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:卡类别停用或启用
    '编制:刘兴洪
    '日期:2011-06-27 20:56:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngTypeId As Long, lngColor As Long, i As Long
    Dim strSQL As String, intIndex As Integer
    Dim varTemp As Variant
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    varTemp = Split(lvwMain.SelectedItem.Tag & "-", "-")
    'If Val(varTemp(1)) = 1 Then Exit Sub     '系统固定的,不允许停用和启用
    Err = 0: On Error GoTo Errhand
    lngTypeId = Val(Mid(Me.lvwMain.SelectedItem.Key, 2))
    With lvwMain
         If MsgBox("你确认要" & IIf(blnStop, "停用", "启用") & "医疗卡""" & lvwMain.SelectedItem.Text & """吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            strSQL = "Zl_医疗卡类别_Stopandstart(" & lngTypeId & "," & IIf(blnStop, 1, 0) & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            If mblnShowStop = False And blnStop Then
                    intIndex = .SelectedItem.Index
                    .ListItems.Remove .SelectedItem.Key
                    If .ListItems.count > 0 Then
                            intIndex = IIf(.ListItems.count > intIndex, intIndex, .ListItems.count)
                            .ListItems(intIndex).Selected = True
                            .ListItems(intIndex).EnsureVisible
                    Else
                        Call lvwMain_GotFocus
                    End If
            Else
                If blnStop Then
                    .SelectedItem.Icon = "Stop": .SelectedItem.SmallIcon = "Stop"
                    lngColor = vbRed
                Else
                    .SelectedItem.Icon = "Start": .SelectedItem.SmallIcon = "Start"
                    lngColor = RGB(0, 0, 0)
                End If
                .SelectedItem.ForeColor = lngColor
                For i = 1 To .ColumnHeaders.count
                    If i < .ColumnHeaders.count Then
                        .SelectedItem.ListSubItems(i).ForeColor = lngColor
                    End If
                    If .ColumnHeaders(i).Text = "是否启用" Then
                        .SelectedItem.SubItems(i - 1) = IIf(blnStop, "", "√")
                    End If
                Next
                .SelectedItem.Tag = IIf(blnStop, 1, 0) & "-" & varTemp(1) & "-" & varTemp(2)
            End If
        End If
    End With
    zlCtlSetFocus lvwMain, True
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
 
Private Sub ModifyData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:修改数据
    '编制:刘兴洪
    '日期:2011-06-28 00:59:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngTypeId As Long, varTemp As Variant
    
    If Me.lvwMain.SelectedItem Is Nothing Then Exit Sub
    varTemp = Split(lvwMain.SelectedItem.Tag & "-", "-")
    If Val(varTemp(0)) = 1 Or Val(varTemp(2)) = 1 Then Exit Sub
'    If Val(varTemp(1)) = 1 Then
'        MsgBox "系统固定项,不能修改,请检查!", vbOKOnly + vbInformation, gstrSysName
'        Exit Sub
'    End If
    lngTypeId = Val(Mid(lvwMain.SelectedItem.Key, 2))
    If frmPatiCureCardTypeEdit.zlEditCard(Me, mlngModule, mstrPrivs, edt_修改, lngTypeId) = False Then Exit Sub
    If MsgBox("数据已经发生改变,是否重新刷新数据?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Call LoadData
End Sub
Private Sub DeleteData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:修改数据
    '编制:刘兴洪
    '日期:2011-06-28 00:59:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngTypeId As Long, varTemp As Variant
    Dim intIndex As Integer, strSQL As String
    Err = 0: On Error GoTo Errhand:
    If Me.lvwMain.SelectedItem Is Nothing Then Exit Sub
    varTemp = Split(lvwMain.SelectedItem.Tag & "-", "-")
    If Val(varTemp(0)) = 1 Or Val(varTemp(2)) = 1 Then Exit Sub
    If Val(varTemp(1)) = 1 Then
        MsgBox "系统固定项,不能删除,请检查!", vbOKOnly + vbInformation, gstrSysName
        Exit Sub
    End If
    If MsgBox("你确认要删除名称为“" & lvwMain.SelectedItem.Text & "”的医疗卡类别吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    lngTypeId = Val(Mid(lvwMain.SelectedItem.Key, 2))
        'Zl_医疗卡类别_Delete(Id_In In 医疗卡类别.ID%Type) Is
      strSQL = "Zl_医疗卡类别_Delete(" & lngTypeId & ")"
      
      Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    With lvwMain
        intIndex = .SelectedItem.Index
        .ListItems.Remove .SelectedItem.Key
        If .ListItems.count > 0 Then
            intIndex = IIf(.ListItems.count > intIndex, intIndex, .ListItems.count)
            .ListItems(intIndex).Selected = True
            .ListItems(intIndex).EnsureVisible
        End If
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub ViewData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查询数据
    '编制:刘兴洪
    '日期:2011-06-28 00:59:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngTypeId As Long
    If Me.lvwMain.SelectedItem Is Nothing Then Exit Sub
    lngTypeId = Val(Mid(lvwMain.SelectedItem.Key, 2))
    Call frmPatiCureCardTypeEdit.zlEditCard(Me, mlngModule, mstrPrivs, dt_查看, lngTypeId)
End Sub
Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '入参:
    '出参:
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-11-18 16:53:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup
    Dim objComBar As CommandBarComboBox
        
      
    Err = 0: On Error GoTo Errhand:
    '-----------------------------------------------------
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    
    cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    mcbrMenuBar.id = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): mcbrControl.BeginGroup = True
    End With


    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    mcbrMenuBar.id = conMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "增加(&A)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)"):
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "启用(&R)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "停用(&S)")
    End With

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    mcbrMenuBar.id = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_LargeICO, "大图标(&L)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_MinICO, "小图标(&M)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_ListICO, "列表(&M)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_DetailsICO, "详细资料(&D)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_ShowStoped, "显示停用项目(&S)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    mcbrMenuBar.id = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): mcbrControl.BeginGroup = True
    End With
    
    '快键绑定
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("D"), conMenu_Edit_Delete
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F12, conMenu_File_Parameter
    End With
    
    '设置不常用菜单
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    '-----------------------------------------------------
    '工具栏定义
    Set mcbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "增加"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "启用"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "停用")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
     zlDefCommandBars = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------
Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRow As Long, lngID As Long
    Dim ctrCombox As CommandBarComboBox
    '------------------------------------
        

    Select Case Control.id
    'bytMode=1 打印;2 预览;3 输出到EXCEL
    Case conMenu_File_Preview: Call zlRptPrint(2)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_Edit_NewItem
        If frmPatiCureCardTypeEdit.zlEditCard(Me, mlngModule, mstrPrivs, edT_增加) = False Then Exit Sub
        If MsgBox("数据已经发生改变,是否重新刷新数据?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Call LoadData
    Case conMenu_Edit_Modify
        '修改
        Call ModifyData
    Case conMenu_Edit_Delete    '删除
        Call DeleteData
    Case conMenu_Edit_Reuse  '启用
          Call zlCardStopAndResume(False)
    Case conMenu_Edit_Stop '停用
          Call zlCardStopAndResume(True)
    Case conMenu_View_ShowStoped '显示停用项止
            mblnShowStop = Not mblnShowStop
            Call LoadData
    Case conMenu_View_LargeICO  '大图标
         lvwMain.View = lvwIcon
    Case conMenu_View_MinICO    '小图标
         lvwMain.View = lvwSmallIcon
    Case conMenu_View_ListICO   '列表
         lvwMain.View = lvwList
    Case conMenu_View_DetailsICO    '详细资料
         lvwMain.View = lvwReport
    Case conMenu_View_Refresh   '刷新
        '重新刷新数据
        Call LoadData
    Case Else
        If (Control.id >= conMenu_ReportPopup * 100# + 1 And Control.id <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Call zl_OpenReport(Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1))
        End If
    End Select
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub
 
Private Function IsModifyOrDelete() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否允许修改或删除
    '编制:刘兴洪
    '日期:2011-06-28 11:54:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant
    If lvwMain.SelectedItem Is Nothing Then Exit Function
    varTemp = Split(lvwMain.SelectedItem.Tag & "-", "-")
    IsModifyOrDelete = Val(varTemp(0)) = 0 And Val(varTemp(1)) = 0 And Val(varTemp(2)) = 0
End Function
Private Function IsModify() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否允许修改或删除
    '编制:刘兴洪
    '日期:2011-06-28 11:54:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant
    If lvwMain.SelectedItem Is Nothing Then Exit Function
    ' 是否启用-是否固定
    varTemp = Split(lvwMain.SelectedItem.Tag & "-", "-")
    IsModify = Val(varTemp(0)) = 0 And Val(varTemp(2)) = 0
End Function
Private Function IsStop() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否允许停用
    '编制:刘兴洪
    '日期:2011-06-28 11:54:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant
    If lvwMain.SelectedItem Is Nothing Then Exit Function
    varTemp = Split(lvwMain.SelectedItem.Tag & "-", "-")
    IsStop = Val(varTemp(0)) = 0 ' And Val(varTemp(1)) = 0
End Function
Private Function IsStart() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否允许启用
    '编制:刘兴洪
    '日期:2011-06-28 11:54:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant
    If lvwMain.SelectedItem Is Nothing Then Exit Function
    varTemp = Split(lvwMain.SelectedItem.Tag & "-", "-")
    IsStart = Val(varTemp(0)) = 1 ' And Val(varTemp(1)) = 0
End Function

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngID As Long, blnEnabled As Boolean
     
    If Me.Visible = False Then Exit Sub
    If Control.type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.id
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = lvwMain.ListItems.count > 0
    Case conMenu_Edit_NewItem
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "增加")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Modify
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "修改")
        Control.Enabled = Control.Visible And IsModify
    Case conMenu_Edit_Delete
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "删除")
        Control.Enabled = Control.Visible And IsModifyOrDelete
    Case conMenu_Edit_Reuse
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "启用")
        Control.Enabled = Control.Visible And IsStart
    Case conMenu_Edit_Stop
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "停用")
        Control.Enabled = Control.Visible And IsStop
    Case conMenu_View_ShowStoped '显示停用项止
        Control.Checked = mblnShowStop
    Case conMenu_View_LargeICO  '大图标
        Control.Checked = lvwMain.View = lvwIcon
    Case conMenu_View_MinICO    '小图标
        Control.Checked = lvwMain.View = lvwSmallIcon
    Case conMenu_View_ListICO   '列表
        Control.Checked = lvwMain.View = lvwList
    Case conMenu_View_DetailsICO    '详细资料
        Control.Checked = lvwMain.View = lvwReport
    Case conMenu_View_Refresh   '刷新
    Case Else
        If (Control.id >= conMenu_ReportPopup * 100# + 1 And Control.id <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Control.Visible = Split(Control.Parameter, ",")(1) <> "ZL" & glngSys \ 100 & "_INSIDE_1503_1" And Split(Control.Parameter, ",")(1) <> "ZL" & glngSys \ 100 & "_INSIDE_1107_2"
        End If
    End Select
End Sub
 
'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    '------------------------------------
    Select Case Control.id
        Case conMenu_File_Exit: Unload Me
        Case conMenu_File_PrintSet: Call zlPrintSet
        Case conMenu_View_StatusBar
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Button
            cbsThis(2).Visible = Not cbsThis(2).Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text
            For Each mcbrControl In cbsThis(2).Controls
                mcbrControl.Style = IIf(mcbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size
            cbsThis.Options.LargeIcons = Not cbsThis.Options.LargeIcons
            cbsThis.RecalcLayout
        Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case Else   '其他操作功能调用
            Call zlExecuteCommandBars(Control)
        End Select
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub
Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub
Private Sub cbsThis_Resize()
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    cbsThis.GetClientRect Left, Top, Right, Bottom
    On Error Resume Next
   With lvwMain
        .Left = Left
        .Top = Top
        .Width = Right - Left
        .Height = Bottom - Top
   End With
End Sub
Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean
    If Me.Visible = False Then Exit Sub
    If Control.type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.id
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = stbThis.Visible
    Case Else
        Call zlUpdateCommandBars(Control)
    End Select
End Sub

Private Sub Form_Load()
    mlngModule = glngModul
    mstrPrivs = gstrPrivs
    Call zlDefCommandBars
    Call InitLvwHead
     
    '76905,冉俊明,2014-8-21,第一次进入医疗卡类别管理中,当停用类别是默认显示时,未显示停用类别
    Call InitPara
    Call LoadData
    RestoreWinState Me, App.ProductName
    lvwMain.Tag = "可变化的"
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
End Sub
Private Sub InitPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关数据
    '编制:刘兴洪
    '日期:2011-06-28 11:20:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    mblnShowStop = Val(zlDatabase.GetPara("显示停用类别", glngSys, mlngModule, "1", , , InStr(1, mstrPrivs, ";参数设置;") > 0)) = 1
    i = Val(zlDatabase.GetPara("图标显示方式", glngSys, mlngModule, "3", , , InStr(1, mstrPrivs, ";参数设置;") > 0))
    If i < 0 Or i > 3 Then i = 3
    lvwMain.View = i
End Sub
 

Private Sub Form_Unload(Cancel As Integer)
    Call zlDatabase.SetPara("显示停用类别", IIf(mblnShowStop, 1, 0), glngSys, mlngModule, InStr(1, mstrPrivs, ";参数设置;") > 0)
    Call zlDatabase.SetPara("图标显示方式", lvwMain.View, glngSys, mlngModule, InStr(1, mstrPrivs, ";参数设置;") > 0)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '仍是刚才那列
        lvwMain.SortOrder = IIf(lvwMain.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwMain.SortKey = mintColumn
        lvwMain.SortOrder = lvwAscending
    End If
 
End Sub

Private Sub lvwMain_DblClick()
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    If lvwMain.SelectedItem.Tag Like "1*" Or lvwMain.SelectedItem.Tag Like "*-1" Or InStr(1, mstrPrivs, ";修改;") = 0 Then
        Call ViewData
    Else
        Call ModifyData
    End If
End Sub
Private Sub lvwMain_GotFocus()
    With lvwMain
        stbThis.Panels(2).Text = "共有" & .ListItems.count & "种医疗卡类别。"
    End With
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行打印,预览和输出到EXCEL
    '入参:bytFunc=1 打印;2 预览;3 输出到EXCEL
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-11-20 15:34:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPrint As New zlPrintLvw
    objPrint.Title.Text = gstrUnitName & "医疗类别清单"
    Set objPrint.Body.objData = lvwMain
    objPrint.BelowAppItems.Add "打印人：" & UserInfo.姓名
    objPrint.BelowAppItems.Add "打印时间：" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
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
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
End Sub
Private Sub zl_OpenReport(ByVal lngSys As Long, ByVal strReportCode As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开指定报表
    '入参:lngSys-系统号
    '     strReportCode报表编号
    '编制:刘兴洪
    '日期:2009-11-19 14:15:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID  As Long
    With lvwMain
        If Not .SelectedItem Is Nothing Then
            lngID = Val(Mid(.SelectedItem.Key, 2))
        End If
    End With
    Call ReportOpen(gcnOracle, lngSys, strReportCode, Me, "ID=" & lngID)
End Sub

Private Sub zlPopuMenus()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置弹出菜单
    '编制:刘兴洪
    '日期:2011-06-28 12:18:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cbrPopupBar As CommandBar, cbrPopupItem As CommandBarControl
    
    Err = 0: On Error Resume Next
    If Not (Me.cbsThis.ActiveMenuBar.Controls(2).Visible Or Me.cbsThis.ActiveMenuBar.Controls(3).Visible) Then Exit Sub
    Set cbrPopupBar = Me.cbsThis.Add("弹出菜单", xtpBarPopup)
    If Me.cbsThis.ActiveMenuBar.Controls(2).Visible Then
        Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(2)
        For Each mcbrControl In mcbrMenuBar.CommandBar.Controls
            If mcbrControl.id <> conMenu_View_ToolBar Then
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, mcbrControl.id, mcbrControl.Caption)
            cbrPopupItem.BeginGroup = mcbrControl.BeginGroup
            End If
        Next
    End If
    If Me.cbsThis.ActiveMenuBar.Controls(3).Visible Then
        Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(3)
        For Each mcbrControl In mcbrMenuBar.CommandBar.Controls
            Select Case mcbrControl.id
            Case conMenu_View_ToolBar
            Case Else
                If mcbrControl.Caption Like "工具栏*" Then
                Else
                    Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, mcbrControl.id, mcbrControl.Caption)
                    cbrPopupItem.BeginGroup = mcbrControl.BeginGroup
                End If
            End Select
        Next
    End If
    cbrPopupBar.ShowPopup
End Sub

Private Sub lvwMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Exit Sub
    Call zlPopuMenus
End Sub
