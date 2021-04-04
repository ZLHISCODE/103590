VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~4.OCX"
Begin VB.Form frmTaskSend 
   Caption         =   "发送体检结果"
   ClientHeight    =   6990
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11265
   Icon            =   "frmTaskSend.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picContainer 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   4725
      ScaleHeight     =   2055
      ScaleWidth      =   2790
      TabIndex        =   1
      Top             =   855
      Width           =   2790
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1320
         Left            =   270
         TabIndex        =   2
         Top             =   390
         Width           =   3135
         _cx             =   5530
         _cy             =   2328
         Appearance      =   1
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
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
         Begin VB.Line lnX 
            Index           =   0
            Visible         =   0   'False
            X1              =   -555
            X2              =   1230
            Y1              =   555
            Y2              =   555
         End
         Begin VB.Line lnY 
            Index           =   0
            Visible         =   0   'False
            X1              =   270
            X2              =   270
            Y1              =   420
            Y2              =   1635
         End
      End
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3060
      Left            =   60
      TabIndex        =   0
      Top             =   1035
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   5398
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "任务包号"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "任务包名"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "发送状态"
         Object.Width           =   1587
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1155
      Top             =   5190
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
            Picture         =   "frmTaskSend.frx":076A
            Key             =   "package"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskSend.frx":6FCC
            Key             =   "package_ok"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenuHot 
      Left            =   7995
      Top             =   1020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskSend.frx":D82E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskSend.frx":DA4E
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskSend.frx":DC6E
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskSend.frx":E3E8
            Key             =   "Send"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   4500
      Top             =   4650
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   6630
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTaskSend.frx":EB62
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14790
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
      DesignerControls=   "frmTaskSend.frx":F3F6
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   600
      Top             =   105
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmTaskSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean                          '窗体启动标志
Private mlngLoop As Long
Private mstrKey As String
Private mfrmMain As Object
Private mvarParam As Variant
Private mstrSQL As String
Private mstrPrive As String
Private mblnShowAll As Boolean

Private Enum mCol
    姓名
    门诊号
    报到
    进度
    总检
    完成
    组别
End Enum

Private Function InitMenuBar() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：初始化菜单、工具栏
    '------------------------------------------------------------------------------------------------------------------
    Dim cbrMenuBar As Object
    Dim obj As CommandBarControl
    Dim cbrControl As Object
    Dim cbrToolBar As CommandBar
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Me.cbsThis.Icons = frmPubIcons.imgPublic.Icons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = True
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单栏"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数(&P)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Task_Send, "发送(&S)")
        
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
        cbrControl.BeginGroup = True
    End With

        
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Expend_CurExpend, "所有人员(&A)")
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
        cbrControl.BeginGroup = True
        
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)..."): cbrControl.BeginGroup = True
    End With
    
     '快键绑定
    With cbsThis.KeyBindings
        
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    

    '工具栏定义
    Set cbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Task_Send, "发送")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
End Function

Private Function InitClient() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：初始化窗格
    '------------------------------------------------------------------------------------------------------------------
    Dim panTab As Pane
    
    Set panTab = dkpMan.CreatePane(1, 200, 500, DockLeftOf, Nothing)
    panTab.Title = ""
    panTab.Options = PaneNoCaption
    
    Set panTab = dkpMan.CreatePane(2, 500, 200, DockRightOf, Nothing)
    panTab.Title = ""
    panTab.Options = PaneNoCaption
    
    dkpMan.SetCommandBars cbsThis
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.HideClient = True
        
End Function

Private Function zlClearData(Optional ByVal strItem As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：清除指定区域的显示数据
    '返回：True
    '------------------------------------------------------------------------------------------------------------------
    
    lvw.ListItems.Clear
    
    vsf.Rows = 2
    vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
    vsf.RowData(1) = 0

    Call AppendRows(vsf, lnX, lnY)
    
    zlClearData = True
    
End Function

Private Function zlMenuClick(ByVal strMenuItem As String, Optional ByVal strParam As String) As Boolean
   
    On Error GoTo errHand
    
    Select Case strMenuItem
    Case "读取体检单"
        
        Call zlClearData
        
        If ReadBill Then
            If Not (lvw.SelectedItem Is Nothing) Then
                Call zlMenuClick("读取概况")
            End If
        End If
       
    Case "读取概况"
        
        frmWait.OpenWait Me, "读取任务包数据"
        frmWait.WaitInfo = "正在读取任务包对应的体检人员"
    
        vsf.Rows = 2
        vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
        vsf.RowData(1) = 0
        
        Call ReadBillState
        
        frmWait.CloseWait
        
    Case "发送结果包"
        
        If ConnectAccess(strParam) Then
        
            DoEvents
            
            If SendPackage Then
                ShowSimpleMsg "体检结果包已被成功发送！"
                lvw.SelectedItem.SubItems(2) = "已发送"
            End If
            
        End If
        
        If gcnAccess.State = adStateOpen Then gcnAccess.Close
                        
    End Select
    
    zlMenuClick = True
    
    Exit Function
    
errHand:
    frmWait.CloseWait
    ShowSimpleMsg Err.Description
        
End Function

Private Function ReadBill() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:读取任务
    '返回:读取成功返回True；否则返回False
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strStart As String
    Dim strEnd As String
    Dim objItem As ListItem
    
    On Error GoTo errHand
    
    gstrSQL = "Select Decode(B.发送状态,1,'package_ok','package') As Icon,A.ID,B.任务包号,B.任务包名,Decode(B.发送状态,1,'已发送','未发送') As 发送状态 From 体检登记记录 A,体检登记记录_干保 B Where A.ID=B.登记id And A.体检状态>2"
    
    strStart = GetDateTime(GetSetting("ZLSOFT", "公共全局\干保接口", "体检时间", "今  天"), 1)
    strEnd = GetDateTime(GetSetting("ZLSOFT", "公共全局\干保接口", "体检时间", "今  天"), 2)
    If strStart = "" Then strStart = GetDateTime("今  天", 1)
    If strEnd = "" Then strEnd = GetDateTime("今  天", 2)
            
    gstrSQL = gstrSQL & "AND A.体检时间 BETWEEN TO_DATE('" & strStart & "','yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & strEnd & "','yyyy-mm-dd hh24:mi:ss')"

    Call OpenRecord(rs, gstrSQL, Me.Caption)
    If rs.BOF = False Then
        Do While Not rs.EOF
            Set objItem = lvw.ListItems.Add(, "_" & rs("ID").Value, NVL(rs("任务包号")), NVL(rs("Icon")), NVL(rs("Icon")))
            objItem.SubItems(1) = NVL(rs("任务包名"))
            objItem.SubItems(2) = NVL(rs("发送状态"))
            rs.MoveNext
        Loop
    End If
    
    ReadBill = True
    
    Exit Function
    
errHand:
    ShowSimpleMsg Err.Description

End Function

Private Function ReadBillState() As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------------------
    '
    '
    '
    '------------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngKey As Long
    Dim lngRow As Long
    Dim rs As New ADODB.Recordset
    Dim lngLoop As Long
    Dim lngCount1 As Long
    Dim lngCount0 As Long
    Dim lngCount2 As Long
    
    On Error GoTo errHand
    
    If lvw.SelectedItem Is Nothing Then Exit Function
    
    lngKey = Val(Mid(lvw.SelectedItem.Key, 2))
    
    gstrSQL = "SELECT A.组别名称 AS 组别,A.病人id AS ID,A.姓名,B.门诊号," & _
                      "A.体检报到 AS 报到," & _
                      "DECODE(C.体检进度,NULL,NULL,TRIM(TO_CHAR(C.体检进度,'9990.0'))||'%') AS 进度," & _
                      "DECODE(A.体检病历ID, Null, 0, 1) As 总检, " & _
                      "DECODE(A.体检状态, 5, 1, 0) As 完成 " & _
                 "FROM 体检人员档案 A," & _
                      "病人信息 B," & _
                      "(SELECT 病人id,DECODE(COUNT(*), NULL, NULL, 100 * SUM(是否已检) / COUNT(*)) AS 体检进度 " & _
                         "FROM (SELECT C.病人id," & _
                                      "(select DECODE(S.报告id, NULL, 0, 1) " & _
                                         "FROM 病人医嘱记录 M, 病人医嘱发送 S " & _
                                        "Where (M.ID = C.医嘱ID Or M.相关id = C.医嘱ID) AND M.ID = S.医嘱ID AND S.报告id > 0 AND ROWNUM < 2) AS 是否已检 " & _
                                 "FROM 体检项目医嘱 C," & _
                                      "(SELECT A.ID, B.病人id " & _
                                         "FROM 体检项目清单 A, 体检人员档案 B " & _
                                        "WHERE A.登记ID = B.登记id AND A.组别名称 = B.组别名称 AND A.登记ID = " & lngKey & " " & _
                                       "Union All " & _
                                         "SELECT A.ID, B.病人id " & _
                                           "FROM 体检项目清单 A, 体检人员档案 B " & _
                                          "WHERE A.登记ID = B.登记id AND A.病人id = B.病人id AND A.登记ID = " & lngKey & " " & _
                                       ") D " & _
                                "WHERE C.清单ID = D.ID AND C.病人ID = D.病人id) " & _
                        "GROUP BY 病人id) C " & _
                "WHERE A.病人ID = B.病人ID(+) AND A.病人ID = C.病人id(+) AND A.登记ID = " & lngKey
                
    If mblnShowAll = False Then
        gstrSQL = gstrSQL & " And A.体检报到=1 "
    End If
    gstrSQL = gstrSQL & " ORDER BY A.组别名称,B.门诊号 "
    
    Call OpenRecord(rs, gstrSQL, Me.Caption)
    If rs.BOF = False Then

        Call LoadGrid(vsf, rs)
        
        Call AppendRows(vsf, lnX, lnY)
        
        '统计已完人数、未完人数及未报到人数
        For lngLoop = 1 To vsf.Rows - 1
            
            '未报到统计
            If Abs(Val(vsf.TextMatrix(lngLoop, mCol.报到))) <> 1 Then
                lngCount0 = lngCount0 + 1
            Else
                '已完统计
                If Abs(Val(vsf.TextMatrix(lngLoop, mCol.完成))) = 1 Then
                    lngCount1 = lngCount1 + 1
                Else
                    '未完统计
                    lngCount2 = lngCount2 + 1
                End If
            End If
        Next
        
        stbThis.Panels(2).Text = "应到:" & lngCount0 + lngCount1 + lngCount2 & "人;实到:" & lngCount1 + lngCount2 & "人(完成:" & lngCount1 & "人;未完:" & lngCount2 & "人);未到:" & lngCount0 & "人"
                    
    End If
    
    
    ReadBillState = True
    
    Exit Function
    
errHand:
    
    ShowSimpleMsg Err.Description
End Function

Private Function SendPackage() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:接受任务
    '返回:接受成功返回True；否则返回False
    '------------------------------------------------------------------------------------------------------------------
    Dim rsSQL As ADODB.Recordset
    
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim lngLoop As Long
    Dim blnTran As Boolean
    
    Dim str体检号 As String
    Dim strSvr组合科室 As String
    Dim strSvr组合编码 As String
    Dim strSvr组合名称 As String
    Dim strSvr体检医生  As String
    Dim strSvr审核日期 As String
    Dim str科室小结 As String
    Dim str科室项目编码 As String
    Dim str科室项目小结 As String
    
    On Error GoTo errHand
    
    
    If lvw.SelectedItem.SubItems(2) = "已发送" Then
        If MsgBox("此任务的体检结果已经发送，是否需要重新发送？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
        
    blnTran = True
    gcnAccess.BeginTrans

    frmWait.OpenWait frmMain, "发送结果包"
    frmWait.WaitInfo = "正在删除原有数据..."
    
    gstrSQL = "Delete From hdatadeptest_分科项目结果"
    gcnAccess.Execute gstrSQL
    
    gstrSQL = "Delete From hcheckmemb_已检人员"
    gcnAccess.Execute gstrSQL
    
    gstrSQL = "Delete From hdatadep_分科小结"
    gcnAccess.Execute gstrSQL
    
    gstrSQL = "Delete From hdatadepunion_体检组合结果"
    gcnAccess.Execute gstrSQL
    
    gstrSQL = "Delete From hdatadepdiag_分科诊断结果"
    gcnAccess.Execute gstrSQL
    
    gstrSQL = "Delete From hdatadiag_主检诊断结果"
    gcnAccess.Execute gstrSQL
    
    gstrSQL = "Delete From hdatarep_主检报告"
    gcnAccess.Execute gstrSQL
    
    frmWait.ShowProgress = True
    
    For mlngLoop = 1 To vsf.Rows - 1
        
        frmWait.WaitInfo = "正在发送体检结果..."
        frmWait.WaitProgress = Format(100 * mlngLoop / (vsf.Rows - 1), "0.00")
            
        If Abs(Val(vsf.TextMatrix(mlngLoop, mCol.报到))) = 1 Then
                        
            mstrSQL = GetPublicSQL(SQL.人员基本资料)
                                                    
            Set rs = OpenSQLRecord(mstrSQL, Me.Caption, lvw.SelectedItem.Text, Val(vsf.RowData(mlngLoop)))
'            Call OpenRecord(rs, mstrSQL, Me.Caption)
            If rs.BOF = False Then
                
                str体检号 = NVL(rs("任务包号")) & NVL(rs("人员序号"))
                
                '1.上传已检人员，hcheckmemb_已检人员------------------------------------------------------------------
                
                gstrSQL = "Delete From hcheckmemb_已检人员 Where checkcode='" & str体检号 & "'"
                gcnAccess.Execute gstrSQL
                
                gstrSQL = "Insert Into hcheckmemb_已检人员(checkcode,ifdel,taskcode,taskseq,seq,ifprinted,checkstatus,iffinished,b0110,b0105,b0160,membcode,membtype,a0101,a0107,age,a6405,a0704,checkdate,asmcode,asmseq,asmname,ifasmdep,asmdepstr,checkfee,ifplus,checkfeeplus,remark,tele,email,accesscode,sendwayid,ifsend,"
                gstrSQL = gstrSQL & "fee01,fee02,fee03,fee04,fee05,fee06,fee07,fee08,fee09,fee10,fee11,fee12,fee13,fee14,fee15,fee16,fee17,fee18,fee19,fee20,"
                gstrSQL = gstrSQL & "feesum,ifcard,workunit,undofee,discountfee,pis_01,bseq,ifad,adclass,tasktype) Values ("
                
                gstrSQL = gstrSQL & "'" & str体检号 & "','0','" & NVL(rs("任务包号")) & "','" & NVL(rs("人员序号")) & "','" & NVL(rs("人员序号")) & "',"
                gstrSQL = gstrSQL & "'0','0','1',"
                gstrSQL = gstrSQL & "'" & NVL(rs("单位编码")) & "',"
                gstrSQL = gstrSQL & "'" & NVL(rs("单位名称")) & "',"
                gstrSQL = gstrSQL & "NULL,"
                gstrSQL = gstrSQL & "'" & NVL(rs("健康号")) & "',"
                gstrSQL = gstrSQL & "'01',"
                gstrSQL = gstrSQL & "'" & NVL(rs("姓名")) & "',"
                gstrSQL = gstrSQL & "'" & NVL(rs("性别")) & "',"
                gstrSQL = gstrSQL & "'" & NVL(rs("年龄")) & "',"
                gstrSQL = gstrSQL & "'" & NVL(rs("在职情况")) & "',"
                gstrSQL = gstrSQL & "'" & NVL(rs("任职级别")) & "',"
                gstrSQL = gstrSQL & "'" & Format(NVL(rs("体检时间")), "yyyy-MM-dd") & "',"
                gstrSQL = gstrSQL & "'" & NVL(rs("套餐编码")) & "',"
                gstrSQL = gstrSQL & "'" & NVL(rs("套餐序号")) & "',"
                gstrSQL = gstrSQL & "'" & NVL(rs("套餐名称")) & "',"
                gstrSQL = gstrSQL & "'0',NULL,'0','0','0',NULL,NULL,NULL,NULL,NULL,'0',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,"
                gstrSQL = gstrSQL & "'0',NULL,NULL,'0','0',NULL,NULL,'0','0',NULL"
                gstrSQL = gstrSQL & ")"
                gcnAccess.Execute gstrSQL
                
                                    
                '2.上传分科项目结果，hdatadeptest_分科项目结果------------------------------------------------------------------
                '在医院增加的项目不上传，上传的都是任务包里指定了的项目,如果没有对码的话，也不上传
                
                mstrSQL = GetPublicSQL(SQL.分科项目结果)
                Set rsTmp = OpenSQLRecord(mstrSQL, Me.Caption, Val(NVL(rs("登记id"))), Val(vsf.RowData(mlngLoop)), lvw.SelectedItem.Text)
                If rsTmp.BOF = False Then
                    Do While Not rsTmp.EOF
                        
                        gstrSQL = "Delete From hdatadeptest_分科项目结果 " & _
                                    "Where checkcode='" & str体检号 & "' and deptcode='" & NVL(rsTmp("组合科室")) & "' and testcode='" & NVL(rsTmp("项目编码")) & "'"
                        gcnAccess.Execute gstrSQL
                        
                        gstrSQL = "Insert Into hdatadeptest_分科项目结果(gb2260,checkcode,deptcode,testcode,wayid,taskcode,membcode,unioncode,htestcode,testname,testresult,teststatus,testsign,testunit,testrange,testlower,testhigher,warncode,remark) values ("
                        
                        gstrSQL = gstrSQL & "'5000',"                                                                   'gb2260,默认5000
                        gstrSQL = gstrSQL & "'" & str体检号 & "',"                                                      'checkcode,体检号
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("组合科室")) & "',"                                            'deptcode,体检科室编码
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("项目编码")) & "',"                                            'testcode,项目编码
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("项目方法")) & "',"                                            'wayid,方法
                        gstrSQL = gstrSQL & "'" & NVL(rs("任务包号")) & "',"                                            'taskcode,任务包号
                        gstrSQL = gstrSQL & "'" & NVL(rs("健康号")) & "',"                                              'membcode,保健号
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("组合编码")) & "',"                                            'unioncode,组合编码
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("项目分支")) & "',"                                            'htestcode,项目所有大类
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("项目名称")) & "',"                                            'testname,项目名称
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("结果")) & "',"                                             'testresult,项目结果
                        gstrSQL = gstrSQL & "NULL,"                                                                     'teststatus,项目状态 null，偏低 偏高
                        gstrSQL = gstrSQL & "'" & Left(NVL(rsTmp("标志"), "0"), 1) & "',"                                           'testsign,0：正常 1：偏低 2：偏高 3：阴性 4：阳性 5：弱阳 9：异常
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("单位")) & "',"                                             'testunit,项目单位
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("参考")) & "',"                                             'testrange,范围
                        gstrSQL = gstrSQL & "NULL,"                                                                     'testlower,默认null
                        gstrSQL = gstrSQL & "NULL,"                                                                     'testhigher,默认null
                        gstrSQL = gstrSQL & "NULL,"                                                                     'warncode,默认null
                        gstrSQL = gstrSQL & "NULL"                                                                     'remark,默认null
                        gstrSQL = gstrSQL & ")"
                        
                        gcnAccess.Execute gstrSQL
                        
                        rsTmp.MoveNext
                    Loop
                End If
                
                '上传小结和诊断数据
                mstrSQL = GetPublicSQL(SQL.分科项目结论)
                'Call OpenRecord(rsTmp, mstrSQL, Me.Caption)
                Set rsTmp = OpenSQLRecord(mstrSQL, Me.Caption, Val(NVL(rs("登记id"))), Val(vsf.RowData(mlngLoop)), lvw.SelectedItem.Text)
                If rsTmp.BOF = False Then
                    Do While Not rsTmp.EOF
                        
                        strSvr组合科室 = NVL(rsTmp("组合科室"))
                        strSvr组合编码 = NVL(rsTmp("组合编码"))
                        strSvr组合名称 = NVL(rsTmp("组合名称"))
                        strSvr体检医生 = NVL(rsTmp("书写人"))
                        strSvr审核日期 = Format(NVL(rsTmp("审阅日期")), "yyyy-MM-dd")
                        
                        str科室小结 = str科室小结 & NVL(rsTmp("结论描述")) & vbCrLf
                        str科室项目编码 = str科室项目编码 & ";" & NVL(rsTmp("组合编码"))

                        str科室项目小结 = str科室项目小结 & NVL(rsTmp("结论描述")) & vbCrLf
                              
                        rsTmp.MoveNext
                        
                        If rsTmp.EOF Then
                            GoTo 科室小结
                        Else
                            If strSvr组合科室 <> NVL(rsTmp("组合科室")) Then
                                GoTo 科室小结
                            ElseIf strSvr组合编码 <> NVL(rsTmp("组合编码")) Then
                                GoTo 科室项目小结
                            End If
                        End If
                        
                        GoTo OverPoint
科室小结:
                        If str科室小结 <> "" Then
                                
                            If str科室项目编码 <> "" Then str科室项目编码 = Mid(str科室项目编码, 2)
                                
                            '3.上传科室分科小结，hdatadep_分科小结------------------------------------------------------------------
                            
                            gstrSQL = "Delete From hdatadep_分科小结 " & _
                                    "Where checkcode='" & str体检号 & "' and deptcode='" & strSvr组合科室 & "'"
                            gcnAccess.Execute gstrSQL
                        
                            gstrSQL = "Insert Into hdatadep_分科小结(gb2260,checkcode,deptcode,unioncode,taskcode,seq,membcode,membtype,initday,checkstatus,sampleno,depresult,checkdate,checkdoc,reviewdoc,iffinished,iflock,ifdata,unionfee,ifplus,checklevel,tag,remark,depsignstr,depdiagstr,depopsstr,oper,ifad,adclass,deptseq) values ("
                            
                            gstrSQL = gstrSQL & "'5000',"                                                                           'gb2260,默认5000
                            gstrSQL = gstrSQL & "'" & str体检号 & "',"                                                              'checkcode,体检号
                            gstrSQL = gstrSQL & "'" & strSvr组合科室 & "',"                                                         'deptcode,体检科室编码
                            gstrSQL = gstrSQL & "'" & str科室项目编码 & "',"                                                        'unioncode,体检组合编码
                            gstrSQL = gstrSQL & "'" & NVL(rs("任务包号")) & "',"                                                    'taskcode,任务包号
                            gstrSQL = gstrSQL & "'" & NVL(rs("人员序号")) & "',"                                            'seq,任务包序号
                            gstrSQL = gstrSQL & "'" & NVL(rs("健康号")) & "',"                                              'membcode,保健号
                            gstrSQL = gstrSQL & "'01',"                                                                     'membtype,默认01
                            gstrSQL = gstrSQL & "'" & Format(NVL(rs("体检时间")), "yyyy-MM-dd") & "',"                      'initday,体检日期
                            gstrSQL = gstrSQL & "'5',"                                                                      'checkstatus,默认5
                            gstrSQL = gstrSQL & "NULL,"                                                                     'sampleno,默认null
                            gstrSQL = gstrSQL & "'" & str科室小结 & "',"                                                  'depresult,分科小结
                            gstrSQL = gstrSQL & "NULL,"                                                                     'checkdate,最后审核日期
                            gstrSQL = gstrSQL & "NULL,"                                                                     'checkdoc,审核医生
                            gstrSQL = gstrSQL & "NULL,"                                                                     'reviewdoc,复查医生,默认null
                            gstrSQL = gstrSQL & "'1',"                                                                      'iffinished,默认1
                            gstrSQL = gstrSQL & "'0',"                                                                      'iflock,默认0
                            gstrSQL = gstrSQL & "'0',"                                                                      'ifdata,默认0
                            gstrSQL = gstrSQL & "NULL,"                                                                     'unionfee,默认null
                            gstrSQL = gstrSQL & "'0',"                                                                      'ifplus,默认0
                            gstrSQL = gstrSQL & "'0',"                                                                      'checklevel,默认0
                            gstrSQL = gstrSQL & "NULL,"                                                                     'tag,默认null
                            gstrSQL = gstrSQL & "NULL,"                                                                     'remark,默认null
                            gstrSQL = gstrSQL & "NULL,"                                                                     'depsignstr,默认null
                            gstrSQL = gstrSQL & "NULL,"                                                                     'depdiagstr,默认null
                            gstrSQL = gstrSQL & "NULL,"                                                                     'depopsstr,默认null
                            gstrSQL = gstrSQL & "NULL,"                                                                     'oper,默认null
                            gstrSQL = gstrSQL & "'0',"                                                                      'ifad,默认0
                            gstrSQL = gstrSQL & "'DEP',"                                                                    'adclass,默认DEP
                            gstrSQL = gstrSQL & "'0'"                                                                      'deptseq默认0
                            
                            gstrSQL = gstrSQL & ")"
                            gcnAccess.Execute gstrSQL
                        End If
                        
                        str科室小结 = ""
                        str科室项目编码 = ""
                                                    
科室项目小结:
                        If str科室项目小结 <> "" Then
                                
                            '上传体检组合结果
                            
                            gstrSQL = "Delete From hdatadepunion_体检组合结果 " & _
                                    "Where checkcode='" & str体检号 & "' and deptcode='" & strSvr组合科室 & "' and unioncode='" & strSvr组合编码 & "'"
                            gcnAccess.Execute gstrSQL
                            
                            gstrSQL = "Insert Into hdatadepunion_体检组合结果(gb2260,checkcode,deptcode,unioncode,taskcode,seq,membcode,membtype,initday,checkstatus,sampleno,depresult,checkdate,checkdoc,reviewdoc,iffinished,iflock,ifdata,testsignstr,unionfee,ifplus,tag,ifsettle,rackno,rackoper,racktime,uname,deptseq,rackbatch,uniondesc,regstatus,checklevel,settlecode) values ("
                                'work
                            
                            gstrSQL = gstrSQL & "'5000',"                                                                  'gb2260,默认5000
                            gstrSQL = gstrSQL & "'" & str体检号 & "',"                                                      'checkcode,体检号
                            gstrSQL = gstrSQL & "'" & strSvr组合科室 & "',"                                             'deptcode,体检科室编码
                            gstrSQL = gstrSQL & "'" & strSvr组合编码 & "',"                                     'unioncode,组合编码
                            gstrSQL = gstrSQL & "'" & NVL(rs("任务包号")) & "',"                                        'taskcode,任务包号
                            gstrSQL = gstrSQL & "'" & NVL(rs("人员序号")) & "',"                                        'seq,任务包序号
                            gstrSQL = gstrSQL & "'" & NVL(rs("健康号")) & "',"                                          'membcode,保健号
                            gstrSQL = gstrSQL & "'01',"                                                                  'membtype,默认01
                            gstrSQL = gstrSQL & "'" & Format(NVL(rs("体检时间")), "yyyy-MM-dd") & "',"                  'initday,体检日期,日期/时间
                            gstrSQL = gstrSQL & "'0',"                                                                  'checkstatus,默认0
                            gstrSQL = gstrSQL & "NULL,"                                                                 'sampleno,默认null
                            gstrSQL = gstrSQL & "'" & str科室项目小结 & "',"                                                 'depresult,组合结果
                            gstrSQL = gstrSQL & "'" & strSvr审核日期 & "',"                                         'checkdate,审核日期,日期/时间
                            gstrSQL = gstrSQL & "'" & strSvr体检医生 & "',"                                        'checkdoc,体检医生
                            gstrSQL = gstrSQL & "NULL,"                                                                 'reviewdoc,复审医生，可选
                            gstrSQL = gstrSQL & "'1',"                                                                  'iffinished,默认1
                            gstrSQL = gstrSQL & "'0',"                                                                  'iflock,默认0
                            gstrSQL = gstrSQL & "'0',"                                                                  'ifdata,默认0
                            gstrSQL = gstrSQL & "'0',"                                                                  'testsignstr,默认0
                            gstrSQL = gstrSQL & "'0',"                                                                  'unionfee,组合费用
                            gstrSQL = gstrSQL & "0,"                                                                  'ifplus,是否加项 0：否 1：是
                            gstrSQL = gstrSQL & "NULL,"                                                                 'tag,默认null
                            gstrSQL = gstrSQL & "'0',"                                                                  'ifsettle,默认0
                            gstrSQL = gstrSQL & "NULL,"                                                                 'rackno,默认null
                            gstrSQL = gstrSQL & "NULL,"                                                                 'rackoper,默认null
                            gstrSQL = gstrSQL & "NULL,"                                                                 'racktime,默认null
                            gstrSQL = gstrSQL & "'" & strSvr组合名称 & "',"                                             'uname,组合名称
                            gstrSQL = gstrSQL & "0,"                                                                  'deptseq,默认0
                            'gstrSQL = gstrSQL & "NULL,"                                                                 'work,默认null
                            gstrSQL = gstrSQL & "NULL,"                                                                 'rackbatch,默认null
                            gstrSQL = gstrSQL & "NULL,"                                                                 'uniondesc,默认null
                            gstrSQL = gstrSQL & "0,"                                                                  'regstatus,默认0
                            gstrSQL = gstrSQL & "'0',"                                                                  'checklevel,默认0
                            gstrSQL = gstrSQL & "NULL"                                                                 'settlecode,默认空字符串
                            gstrSQL = gstrSQL & ")"
                            gcnAccess.Execute gstrSQL
                        End If
                        
                        str科室项目小结 = ""
                                                    
OverPoint:
                    Loop
                End If
                
                '上传诊断数据
                mstrSQL = GetPublicSQL(SQL.分科项目诊断)
                Set rsTmp = OpenSQLRecord(mstrSQL, Me.Caption, Val(NVL(rs("登记id"))), Val(vsf.RowData(mlngLoop)), lvw.SelectedItem.Text)
                If rsTmp.BOF = False Then
                    Do While Not rsTmp.EOF
                        
                        '4.上传分科诊断结果，hdatadepdiag_分科诊断结果------------------------------------------------------------------
                        
                        gstrSQL = "Delete From hdatadepdiag_分科诊断结果 " & _
                                "Where checkcode='" & str体检号 & "' and deptcode='" & NVL(rsTmp("组合科室")) & "' and signcode='" & NVL(rsTmp("诊断编码")) & "'"
                        gcnAccess.Execute gstrSQL
                        
                        gstrSQL = "Insert Into hdatadepdiag_分科诊断结果(gb2260,checkcode,deptcode,unioncode,taskcode,membcode,ifdel,iftag,ifnewsign,ifwhy,signtype,signcode,signclass,signstatus,diagdeptcode,seq,htestcode,diagwhere,diagdegree,diagclass,stdcode,diagcode,mcode,diagname,diagviewform,checkdate,checkdoc,testinfoch,diaginfoch,remark,tag,ifguide,ifimpt,ifdoubt) values ("
                        
                        gstrSQL = gstrSQL & "'5000',"                                                           'gb2260,默认5000
                        gstrSQL = gstrSQL & "'" & str体检号 & "',"                                              'checkcode,体检号
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("组合科室")) & "',"                                 'deptcode,体检科室编码
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("组合编码")) & "',"                                 'unioncode,组合编码
                        gstrSQL = gstrSQL & "'" & NVL(rs("任务包号")) & "',"                                    'taskcode,任务包号
                        gstrSQL = gstrSQL & "'" & NVL(rs("健康号")) & "',"                                      'membcode,保健号
                        gstrSQL = gstrSQL & "'0',"                                                              'ifdel,默认0
                        gstrSQL = gstrSQL & "'0',"                                                              'iftag,?
                        gstrSQL = gstrSQL & "'0',"                                                              'ifnewsign ,是否新发现诊断,默认0
                        gstrSQL = gstrSQL & "'0',"                                                              'ifwhy,是否可疑诊断,默认0
                        gstrSQL = gstrSQL & "'1',"                                                              'signtype,诊断类型 0:阳性诊断 1:疾病诊断
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("诊断编码")) & "',"                                 'signcode,诊断编码
                        gstrSQL = gstrSQL & "NULL,"                                                             'signclass,默认null
                        gstrSQL = gstrSQL & "NULL,"                                                             'signstatus,默认null
                        gstrSQL = gstrSQL & "NULL,"                                                             'diagdeptcode ,诊断科室编码
                        gstrSQL = gstrSQL & "NULL,"                                                             'seq,默认null
                        gstrSQL = gstrSQL & "'',"                                 'htestcode,项目大类编码
                        gstrSQL = gstrSQL & "NULL,"                                                             'diagwhere,诊断方位，默认null
                        gstrSQL = gstrSQL & "NULL,"                                                             'diagdegree,诊断程度，默认null
                        gstrSQL = gstrSQL & "NULL,"                                                             'diagclass,诊断级别，默认null
                        gstrSQL = gstrSQL & "'ICD-10',"                                                         'stdcode,标准编码集名称
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("疾病编码")) & "',"                                 'diagcode,标准编码
                        gstrSQL = gstrSQL & "NULL,"                                                             'mcode,默认null
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("诊断名称")) & "',"                                 'diagname,诊断名称
                        gstrSQL = gstrSQL & "NULL,"                                                             'diagviewform,默认null
                        gstrSQL = gstrSQL & "'" & Format(NVL(rs("体检时间")), "yyyy-MM-dd") & "',"              'checkdate,体检日期
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("书写人")) & "',"                                   'checkdoc,体检医生
                        gstrSQL = gstrSQL & "NULL,"                                                             'testinfoch,默认null
                        gstrSQL = gstrSQL & "NULL,"                                                             'diaginfoch,默认null
                        gstrSQL = gstrSQL & "'" & NVL(rsTmp("诊断建议")) & "',"                                 'remark,诊断建议
                        gstrSQL = gstrSQL & "NULL,"                                                             'Tag,默认null
                        gstrSQL = gstrSQL & "'0',"                                                              'ifguide,默认0
                        gstrSQL = gstrSQL & "'0',"                                                              'ifimpt,默认0
                        gstrSQL = gstrSQL & "'0'"                                                               'ifdoubt,默认0
                        gstrSQL = gstrSQL & ")"
                        gcnAccess.Execute gstrSQL
                        
                        rsTmp.MoveNext
                    Loop
                End If
                                    
                
                '7.上传主检报告，hdatarep_主检报告------------------------------------------------------------------
                mstrSQL = GetPublicSQL(SQL.总检报告建议)
                Set rsTmp = OpenSQLRecord(mstrSQL, Me.Caption, Val(NVL(rs("登记id"))), Val(vsf.RowData(mlngLoop)))
                If rsTmp.BOF = False Then
                    
                    gstrSQL = "Delete From hdatarep_主检报告 Where checkcode='" & str体检号 & "'"
                    gcnAccess.Execute gstrSQL
                    
                    gstrSQL = "Insert Into hdatarep_主检报告(gb2260,checkcode,taskcode,seq,membcode,membtype,initday,checkstatus,iffinished,iflock,hresult,hresultother,hadvice,checkdoc,reviewdoc,checkdate,remark,workunit,iftrace,ifprint,ifad,adclass) values ("

                    gstrSQL = gstrSQL & "'5000',"                                                              'gb2260,默认5000
                    gstrSQL = gstrSQL & "'" & str体检号 & "',"                                                  'checkcode,体检号
                    gstrSQL = gstrSQL & "'" & NVL(rs("任务包号")) & "',"                                        'taskcode,任务包号
                    gstrSQL = gstrSQL & "'" & NVL(rs("人员序号")) & "',"                                        'seq,任务包序号
                    gstrSQL = gstrSQL & "'" & NVL(rs("健康号")) & "',"                                          'membcode,保健号
                    gstrSQL = gstrSQL & "'01',"                                                                 'membtype,默认01
                    gstrSQL = gstrSQL & "'" & Format(NVL(rs("体检时间")), "yyyy-MM-dd") & "',"                  'initday,体检日期
                    gstrSQL = gstrSQL & "'3',"                                                              'checkstatus,默认3
                    gstrSQL = gstrSQL & "'1',"                                                              'iffinished,默认1
                    gstrSQL = gstrSQL & "'0',"                                                              'iflock,默认0
                    gstrSQL = gstrSQL & "'" & NVL(rsTmp("报告头")) & "',"                                      'hresult,体检报告头
                    gstrSQL = gstrSQL & "NULL,"                                                             'hresultother,默认null
                    gstrSQL = gstrSQL & "'" & NVL(rsTmp("健康指导")) & "',"                                      'hadvice,综合健康指导
                    gstrSQL = gstrSQL & "'" & NVL(rsTmp("书写人")) & "',"                               'checkdoc,主检医生
                    gstrSQL = gstrSQL & "NULL,"                                                         'reviewdoc,复审医生 可选
                    gstrSQL = gstrSQL & "'" & Format(NVL(rsTmp("书写日期")), "yyyy-MM-dd") & "',"              'checkdate,主检日期
                    gstrSQL = gstrSQL & "NULL,"                                                             'remark,默认null
                    gstrSQL = gstrSQL & "NULL,"                                                             'workunit,默认null
                    gstrSQL = gstrSQL & "'0',"                                                              'iftrace,默认0
                    gstrSQL = gstrSQL & "'0',"                                                              'ifprint,默认0
                    gstrSQL = gstrSQL & "'0',"                                                              'ifad,默认0
                    gstrSQL = gstrSQL & "'REP'"                                                              'adclass,默认REP

                    gstrSQL = gstrSQL & ")"
                    gcnAccess.Execute gstrSQL
                    
                End If
                
                
                '8.上传主检诊断结果，hdatadiag_主检诊断结果------------------------------------------------------------------
                mstrSQL = GetPublicSQL(SQL.主检诊断结果)
                Set rsTmp = OpenSQLRecord(mstrSQL, Me.Caption, Val(NVL(rs("登记id"))), Val(vsf.RowData(mlngLoop)), lvw.SelectedItem.Text)
                If rsTmp.BOF = False Then
                    
                    'gb2260,checkcode,deptcode,unioncode,taskcode,membcode,ifdel,iftag,ifnewsign,ifwhy,signtype,signcode,signclass,signstatus,diagdeptcode,seq,htestcode,diagwhere,diagdegree,diagclass,stdcode,diagcode,mcode,diagname,diagviewform,checkdate,checkdoc,testinfoch,diaginfoch,remark,tag,ifguide,ifimpt,ifdoubt
                    gstrSQL = "Delete From hdatadiag_主检诊断结果 Where checkcode='" & str体检号 & "' and signcode='" & NVL(rsTmp("诊断编码")) & "'"
                    gcnAccess.Execute gstrSQL
                    
                    gstrSQL = "Insert Into hdatadiag_主检诊断结果(gb2260,checkcode,deptcode,unioncode,taskcode,membcode,ifdel,iftag,ifnewsign,ifwhy,signtype,signcode,signclass,signstatus,diagdeptcode,seq,htestcode,diagwhere,diagdegree,diagclass,stdcode,diagcode,mcode,diagname,diagviewform,checkdate,checkdoc,testinfoch,diaginfoch,remark,tag,ifguide,ifimpt,ifdoubt) values ("

                    gstrSQL = gstrSQL & "'5000',"                                                           'gb2260,默认5000
                    gstrSQL = gstrSQL & "'" & str体检号 & "',"                                              'checkcode,体检号
                    gstrSQL = gstrSQL & "'',"                                 'deptcode,体检科室编码
                    gstrSQL = gstrSQL & "'',"                                 'unioncode,组合编码
                    gstrSQL = gstrSQL & "'" & NVL(rs("任务包号")) & "',"                                    'taskcode,任务包号
                    gstrSQL = gstrSQL & "'" & NVL(rs("健康号")) & "',"                                      'membcode,保健号
                    gstrSQL = gstrSQL & "'0',"                                                              'ifdel,默认0
                    gstrSQL = gstrSQL & "'0',"                                                              'iftag,?
                    gstrSQL = gstrSQL & "'0',"                                                              'ifnewsign ,是否新发现诊断,默认0
                    gstrSQL = gstrSQL & "'0',"                                                              'ifwhy,是否可疑诊断,默认0
                    gstrSQL = gstrSQL & "'1',"                                                              'signtype,诊断类型 0:阳性诊断 1:疾病诊断
                    gstrSQL = gstrSQL & "'" & NVL(rsTmp("诊断编码")) & "',"                                 'signcode,诊断编码
                    gstrSQL = gstrSQL & "NULL,"                                                             'signclass,默认null
                    gstrSQL = gstrSQL & "NULL,"                                                             'signstatus,默认null
                    gstrSQL = gstrSQL & "NULL,"                                                             'diagdeptcode ,诊断科室编码
                    gstrSQL = gstrSQL & "NULL,"                                                             'seq,默认null
                    gstrSQL = gstrSQL & "'',"                                 'htestcode,项目大类编码
                    gstrSQL = gstrSQL & "NULL,"                                                             'diagwhere,诊断方位，默认null
                    gstrSQL = gstrSQL & "NULL,"                                                             'diagdegree,诊断程度，默认null
                    gstrSQL = gstrSQL & "NULL,"                                                             'diagclass,诊断级别，默认null
                    gstrSQL = gstrSQL & "'ICD-10',"                                                         'stdcode,标准编码集名称
                    gstrSQL = gstrSQL & "'" & NVL(rsTmp("疾病编码")) & "',"                                 'diagcode,标准编码
                    gstrSQL = gstrSQL & "NULL,"                                                             'mcode,默认null
                    gstrSQL = gstrSQL & "'" & NVL(rsTmp("诊断名称")) & "',"                                 'diagname,诊断名称
                    gstrSQL = gstrSQL & "NULL,"                                                             'diagviewform,默认null
                    gstrSQL = gstrSQL & "'" & Format(NVL(rs("体检时间")), "yyyy-MM-dd") & "',"              'checkdate,体检日期
                    gstrSQL = gstrSQL & "'" & NVL(rsTmp("书写人")) & "',"                                   'checkdoc,体检医生
                    gstrSQL = gstrSQL & "NULL,"                                                             'testinfoch,默认null
                    gstrSQL = gstrSQL & "NULL,"                                                             'diaginfoch,默认null
                    gstrSQL = gstrSQL & "'" & NVL(rsTmp("诊断建议")) & "',"                                 'remark,诊断建议
                    gstrSQL = gstrSQL & "NULL,"                                                             'Tag,默认null
                    gstrSQL = gstrSQL & "'0',"                                                              'ifguide,默认0
                    gstrSQL = gstrSQL & "'0',"                                                              'ifimpt,默认0
                    gstrSQL = gstrSQL & "'0'"                                                               'ifdoubt,默认0
                    gstrSQL = gstrSQL & ")"

                    gcnAccess.Execute gstrSQL
                    
                End If
                
                gstrSQL = "Update 体检登记记录_干保 Set 发送状态=1 Where 任务包号='" & NVL(rs("任务包号")) & "'"
                gcnOracle.Execute gstrSQL
            
            End If
        End If
    Next
    
    gcnAccess.CommitTrans
    frmWait.CloseWait
    blnTran = False
        
    SendPackage = True
    
    Exit Function
    
errHand:
    Dim strError As String
    
    strError = Err.Description
    If blnTran Then gcnAccess.RollbackTrans
    
    frmWait.CloseWait
    ShowSimpleMsg strError
    
'    Resume
End Function

Private Function InitData() As Boolean
    
    Dim strVsf As String
    
    strVsf = "姓名,900,1,1,1,;门诊号,1080,7,1,1,;报到,600,4,1,1,;进度,810,7,1,1,;总检,600,4,1,1,;完成,600,4,1,1,;组别,2100,1,1,1,"
    Call CreateVsf(vsf, strVsf)
    vsf.Cols = vsf.Cols + 1
    
    vsf.ColWidth(vsf.Cols - 1) = 15
    vsf.ColDataType(mCol.报到) = flexDTBoolean
    vsf.ColDataType(mCol.总检) = flexDTBoolean
    vsf.ColDataType(mCol.完成) = flexDTBoolean
    Call AppendRows(vsf, lnX, lnY)
    
    mblnShowAll = False
    
    InitData = True
    
End Function

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As Object

    On Error GoTo errHand
    
    Select Case Control.ID
                
        Case conMenu_File_Parameter
            
            If frmTaskSendFilter.ShowFilter(Me) Then
                Call zlMenuClick("读取体检单")
                If Not (lvw.SelectedItem Is Nothing) Then Call zlMenuClick("读取概况")
            End If
            
        Case conMenu_Task_Send
            
            If MsgBox("确定现在要发送结果包吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
   
            dlg.Flags = &H4 Or &H200000 Or &H800 & &H1000
            dlg.Filter = "体检结果包|结果包.mdb"
            dlg.FilterIndex = 0
            
            dlg.DialogTitle = "打开体检结果包"
            dlg.FileName = ""
            dlg.ShowOpen
            If dlg.FileName <> "" Then Call zlMenuClick("发送结果包", dlg.FileName)
            
            
        Case conMenu_View_ToolBar_Button
        
            cbsThis(2).Visible = Not cbsThis(2).Visible
            cbsThis.RecalcLayout
        
        Case conMenu_View_ToolBar_Text
        
            For Each cbrControl In cbsThis(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            
            cbsThis.RecalcLayout
            
        Case conMenu_View_StatusBar
        
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
        
        Case conMenu_View_Expend_CurExpend
            
            mblnShowAll = Not mblnShowAll
            If Not (lvw.SelectedItem Is Nothing) Then Call zlMenuClick("读取概况")
            
        Case conMenu_View_Refresh
            
            Call zlMenuClick("读取体检单")
            If Not (lvw.SelectedItem Is Nothing) Then Call zlMenuClick("读取概况")
                        
        Case conMenu_Help_Help
        
            Call ShowHelp(Me.hWnd, Me.Name)
        
        Case conMenu_Help_About
            
            frmAbout.Show 1, Me
            
        Case conMenu_File_Exit
        
            Unload Me
            Exit Sub
            
    End Select
    
    
    cbsThis.RecalcLayout
    
    Exit Sub
    
errHand:
    
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub


Private Sub cbsThis_Resize()
    
    Call AppendRows(vsf, lnX, lnY)
    
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.ID
    Case conMenu_Task_Send
            
        Control.Visible = (InStr(mstrPrive, ";发送结果;") > 0)
        Control.Enabled = (lvw.ListItems.Count > 0)
        
    Case conMenu_View_ToolBar_Button
        Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text
        Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_StatusBar
        Control.Checked = Me.stbThis.Visible
    Case conMenu_View_Expend_CurExpend
    
        Control.Checked = mblnShowAll
        
    End Select
    
    
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    On Error Resume Next
    
    Select Case Item.ID
    Case 1
        
        
        Item.Handle = lvw.hWnd
        
    Case 2
        
       Item.Handle = picContainer.hWnd
    End Select
End Sub

Private Sub Form_Activate()
    
    If mblnStartUp = False Then Exit Sub
        
    If InitData = False Then
        Unload Me
        Exit Sub
    End If
    
    DoEvents
    mblnStartUp = False
    
    Call zlMenuClick("读取体检单")
    If Not (lvw.SelectedItem Is Nothing) Then Call zlMenuClick("读取概况")
    
End Sub

Private Sub Form_Load()
    
    mblnStartUp = True
    
    Call InitMenuBar
    Call InitClient
    
    Call RestoreWinState(Me, App.ProductName)
    mstrPrive = gstrPrive
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If mstrKey <> Item.Key Then
        
        mstrKey = Item.Key
        
        Call zlMenuClick("读取概况")
        
    End If
End Sub

Private Sub picContainer_Resize()
    On Error Resume Next
    
    vsf.Left = 0
    vsf.Top = 0
    vsf.Width = picContainer.Width
    vsf.Height = picContainer.Height
    
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

