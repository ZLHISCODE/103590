VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmQCAddData1 
   Caption         =   "质控数据录入"
   ClientHeight    =   8580
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9900
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmQCAddData2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cbo仪器 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4905
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   0
      Width           =   3200
   End
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2760
      Left            =   120
      TabIndex        =   0
      Top             =   585
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   4868
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   8.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   119341058
      CurrentDate     =   40246
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   8205
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmQCAddData2.frx":000C
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12383
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
   Begin VSFlex8Ctl.VSFlexGrid vfgQCControl 
      Height          =   2595
      Left            =   3150
      TabIndex        =   3
      Top             =   570
      Width           =   6240
      _cx             =   11007
      _cy             =   4577
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
      BackColorFixed  =   15790320
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16635590
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
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
   Begin VSFlex8Ctl.VSFlexGrid vfgQCdata 
      Height          =   3105
      Left            =   150
      TabIndex        =   4
      Top             =   3525
      Width           =   9390
      _cx             =   16563
      _cy             =   5477
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
      BackColorFixed  =   15790320
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16635590
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   270
      Top             =   90
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmQCAddData1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPriv As String
Private mrsQCData As Recordset
Private mintFmtNum As Integer

'-----------------------------------------------------------------------------
'--- 界面逻辑部分
'-----------------------------------------------------------------------------

Private Sub cbo仪器_Click()
    Dim lng仪器id As Long, dateValue As Date
    
    If Me.cbo仪器.ListIndex >= 0 Then
        lng仪器id = Val(Me.cbo仪器.ItemData(Me.cbo仪器.ListIndex))
        dateValue = Me.mvDate.Value
        Call GetQCControlData(Me.vfgQCControl, lng仪器id, dateValue)
        Call vfgQCControl_RowColChange
    End If
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer
    Dim objControl As CommandBarControl
    Select Case Control.ID
    
    Case conMenu_Edit_Modify
        
        Me.vfgQCdata.Editable = flexEDKbdMouse
        Me.vfgQCdata.SelectionMode = flexSelectionFree
        Me.cbo仪器.Enabled = False
        Me.vfgQCControl.Enabled = False
        Me.mvDate.Enabled = False
    Case conMenu_Edit_Untread
        '
        Me.vfgQCdata.Editable = flexEDNone
        Me.vfgQCdata.SelectionMode = flexSelectionByRow
        Me.cbo仪器.Enabled = True
        Me.vfgQCControl.Enabled = True
        Me.mvDate.Enabled = True
        Call RefreshData
        
    
    Case conMenu_Edit_Save
        Dim lng仪器id As Long, lngQCID As Long, dateCurr As Date, strGetQCVal As String
        lng仪器id = Val(Me.cbo仪器.ItemData(Me.cbo仪器.ListIndex))
        dateCurr = Me.mvDate.Value
        lngQCID = Val("" & vfgQCControl.TextMatrix(vfgQCControl.Row, 0))
        strGetQCVal = "" & vfgQCControl.TextMatrix(vfgQCControl.Row, 8)
        Call SaveQcData(vfgQCdata, lng仪器id, lngQCID, dateCurr, strGetQCVal)
        
        Me.vfgQCdata.Editable = flexEDNone
        Me.vfgQCdata.SelectionMode = flexSelectionByRow
        Me.cbo仪器.Enabled = True
        Me.vfgQCControl.Enabled = True
        Me.mvDate.Enabled = True
        
        Call RefreshData
        
    Case conMenu_View_Refresh
        Call RefreshData
    Case conMenu_File_Exit
        Unload Me
    Case conMenu_View_ToolBar_Button '工具栏
        For i = 2 To cbsThis.Count
            Me.cbsThis(i).Visible = Not Me.cbsThis(i).Visible
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text '按钮文字
        For i = 2 To cbsThis.Count
            For Each objControl In Me.cbsThis(i).Controls
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size '大图标
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar '状态栏
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    
    End Select
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_Modify
        Control.Enabled = Not (Me.vfgQCdata.Editable = flexEDKbdMouse)
    Case conMenu_Edit_Save, conMenu_Edit_Untread
        Control.Enabled = (Me.vfgQCdata.Editable = flexEDKbdMouse)
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
        
    End Select
End Sub

Private Sub mvDate_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)
    Call RefreshData
End Sub

Private Sub vfgQCControl_RowColChange()
    Call RefreshData
End Sub

Private Sub cbsThis_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    With Me.mvDate
        .Top = lngTop + 45
        .Left = lngLeft + 45
    End With
    With Me.vfgQCControl
        .Top = Me.mvDate.Top
        .Left = Me.mvDate.Left + Me.mvDate.Width + 45
        .Width = lngRight - .Left - 45
        .Height = Me.mvDate.Height
    End With
    With Me.vfgQCdata
        .Left = Me.mvDate.Left
        .Width = (lngRight - lngLeft) - .Left - 45
        .Top = Me.mvDate.Top + Me.mvDate.Height + 45
        
        .Height = lngBottom - .Top - Me.stbThis.Height - 45
    End With

End Sub
Public Sub ShowMe(ByVal strPrivate As String, ByVal frmMain As Form)
    mstrPriv = strPrivate
    
    Me.Show vbModal, frmMain
End Sub

Private Sub Form_Load()
    
    '初始始化控件公共部分
    '菜单,工具栏
    Call initCbsThis(cbsThis)
    mstrPriv = gstrPrivs
    '状态栏
    'Call InitStatusBar
    
    '初始化控件
    Me.mvDate.Value = Now()
    
    '装入检验仪器数据
    Call LoadInstruments(Me.cbo仪器)
    
End Sub

Private Function initCbsThis(cbsMain As CommandBars) As Boolean
    '作为子窗体处理菜单的基准
    '功能：主窗口菜单定义部份
    '说明：
    '1.其中固有的菜单和按钮必须有，
    '2.其他命令根据主窗体业务的不同，可能不同
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    '菜单定义
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)  '固有
    objMenu.ID = conMenu_FilePopup '对xtpControlPopup类型的命令ID需重新赋值
    With objMenu.CommandBar.Controls
        'Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")  '固有
        'Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        'Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        'Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): objControl.BeginGroup = True '固有
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "放弃(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存(&P)")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False) '固有
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)") '固有
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False '固有
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False '固有
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False '固有
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): objControl.BeginGroup = True '固有

    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False) '固有
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)") '固有
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName) '固有
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False '固有
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False '固有
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False '固有
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): objControl.BeginGroup = True '固有
    End With

    '查找项特殊处理
    '-----------------------------------------------------
'    主菜单右侧的仪器选择
    With cbsMain.ActiveMenuBar.Controls
        Set objControl = .Add(xtpControlLabel, conMenu_View_Dept, "仪器")
        objControl.ID = conMenu_View_Dept
        objControl.Flags = xtpFlagRightAlign
        
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Dept + 1, "")
        objCustom.Handle = cbo仪器.hwnd
        objCustom.Flags = xtpFlagRightAlign
                
    End With

    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        'Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印") '固有
        'Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览") '固有

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "放弃"):
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存")

        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): objControl.BeginGroup = True '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出") '固有
        
    End With

    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        '.Add FCONTROL, vbKeyP, conMenu_File_Print           '打印
        .Add 0, vbKeyF1, conMenu_Help_Help                  '帮助
    End With

    '设置一些公共的不常用命令
    '-----------------------------------------------------
    With cbsMain.Options
       ' .AddHiddenCommand conMenu_File_PrintSet         '打印设置
       ' .AddHiddenCommand conMenu_File_Excel            '输出到Excel
    End With
    
End Function

Private Sub Form_Resize()

   ' On Error Resume Next
    
    Call cbsThis_Resize

End Sub

Private Sub vfgQCdata_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strLists As String, strValue As String
    Dim lngCount As Long
    
    With Me.vfgQCdata
    
        If Col = 0 Then Exit Sub
        If Trim(.TextMatrix(Row, Col)) = "" Then Exit Sub
        
        strLists = Trim(.TextMatrix(Row, 14)) '序列
        strValue = Trim(.TextMatrix(Row, Col))
        
        If strLists = "" Then

            If InStr(strValue, "E+") > 0 And Val(strValue) > 0 Then
                .TextMatrix(Row, Col) = strValue
            Else
                mintFmtNum = Val("" & .TextMatrix(Row, 17))
                If mintFmtNum > 0 Then
                    .TextMatrix(Row, Col) = Format(Val(strValue), "0." & String(mintFmtNum, "0"))
                Else
                    .TextMatrix(Row, Col) = Format(Val(strValue), "0")
                End If
            End If
            
            Exit Sub
        End If
        For lngCount = 0 To UBound(Split(strLists, ";"))
            If .TextMatrix(Row, Col) = Split(strLists, ";")(lngCount) Then Exit Sub
        Next
'        .TextMatrix(Row, Col) = ""
    End With
'    strValue = "该项目为半定量项目，需符合取值序列(" & strLists & ")要求！"
'    MsgBox strValue, vbInformation, gstrSysName
End Sub

Private Sub vfgQCdata_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vfgQCdata
        If Not .TextMatrix(.FixedRows - 1, Col) Like "第?次" Then Cancel = True
    End With
End Sub

Private Sub vfgQCdata_DblClick()
    Me.vfgQCdata.Editable = flexEDKbdMouse
    Me.vfgQCdata.SelectionMode = flexSelectionFree
    Me.cbo仪器.Enabled = False
    Me.vfgQCControl.Enabled = False
    Me.mvDate.Enabled = False
End Sub

Private Sub vfgQCdata_KeyDown(KeyCode As Integer, Shift As Integer)
    With vfgQCdata
        If .Editable <> flexEDNone Then
            If KeyCode = vbKeyReturn Then
                KeyCode = 0
                If .TextMatrix(.FixedRows - 1, .Col) Like "第?次" Then
                    If .Row < .Rows - 1 Then
                        .Select .Row + 1, .Col
                    ElseIf .Col < .Cols - 1 Then
                        If .TextMatrix(.FixedRows - 1, .Col + 1) Like "第?次" Then .Select .FixedRows, .Col + 1
                    End If
                End If
            End If
        End If
    End With

End Sub

Private Sub vfgQCdata_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    With vfgQCdata
        If .Editable <> flexEDNone Then
            If KeyCode = vbKeyReturn Then
                If .TextMatrix(.FixedRows - 1, .Col) Like "第?次" Then
                    If .Row < .Rows - 1 Then
                        .Select .Row + 1, .Col
                    ElseIf .Col < .Cols - 1 Then
                        If .TextMatrix(.FixedRows - 1, .Col + 1) Like "第?次" Then .Select .FixedRows, .Col + 1
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub RefreshData()
    Dim lngQCID As Long, dateValue As Date, strGetQCVal As String
    dateValue = Me.mvDate.Value
    With vfgQCControl
        lngQCID = Val("" & .TextMatrix(.Row, 0))
        strGetQCVal = "" & vfgQCControl.TextMatrix(vfgQCControl.Row, 8)
       
        Call GetQCData(vfgQCdata, lngQCID, dateValue, strGetQCVal)
    End With
End Sub


'-----------------------------------------------------------------------------
'--- 数据处理部分
'-----------------------------------------------------------------------------
Private Function SaveQcData(ByRef vsGrid As VSFlexGrid, ByVal lngDeviceID As Long, ByVal lngQCID As Long, ByVal dateWhy As Date, ByVal strGetQCVal) As Boolean
    '保存数据
    Dim strsql(9) As String, intRow As Integer, lng项目ID As Long
    Dim str结果(9) As String, intCol As Integer, bln有非空结果(9) As Boolean
    Dim strSampleNO(9) As String, lng标本ID(9) As Long, dBegin As Date, dEnd As Date
    Dim rsTemp As ADODB.Recordset, strTmp As String, rsNo As ADODB.Recordset
    Dim blnBegin As Boolean
    Dim strQCVal As String
    On Error GoTo hErr
    
    dBegin = Format(dateWhy, "yyyy-MM-dd 00:00:00")
    dEnd = Format(dateWhy, "yyyy-MM-dd 23:59:59")
    
    With vsGrid
        .Select .FixedRows - 1, 3
        For intRow = .FixedRows To .Rows - 1
            lng项目ID = Val("" & .TextMatrix(intRow, 13))
            If lng项目ID > 0 Then
                
                For intCol = 3 To 11
                    If strGetQCVal = "[SCO]" Then
                        strQCVal = "^^^" & Trim("" & .TextMatrix(intRow, intCol))
                    ElseIf strGetQCVal = "[OD]" Then
                        strQCVal = "^" & Trim("" & .TextMatrix(intRow, intCol)) & "^^"
                    Else
                         strQCVal = Trim("" & .TextMatrix(intRow, intCol))
                    End If
                
                    str结果(intCol - 3) = str结果(intCol - 3) & "|" & lng项目ID & "^" & strQCVal
                    
                    If Trim("" & .TextMatrix(intRow, intCol)) <> "" Then
                        bln有非空结果(intCol - 3) = True
                    End If
                Next
            End If
        Next
    End With
    '取每次的标本号
    For intCol = LBound(strSampleNO) To UBound(strSampleNO)
        
        
        lng标本ID(intCol) = 0
        strSampleNO(intCol) = ""
        Call GetSampleIDNO(lngDeviceID, lngQCID, dBegin, dEnd, intCol + 1, lng标本ID(intCol), strSampleNO(intCol))
        
'        gcnOracle.BeginTrans
'        blnBegin = True
        If lng标本ID(intCol) <= 0 Then
            '无对应标本记录，要增加,但是没有录入数据，全空的不加
            If bln有非空结果(intCol) = True Then
                lng标本ID(intCol) = zlDatabase.GetNextId("检验标本记录")
                gstrSql = "ZL_检验标本记录_INSERT(" & lng标本ID(intCol) & ",NULL,'" & _
                    strSampleNO(intCol) & "',NULL,NULL," & lngDeviceID & ",NULL," & _
                    "To_Date('" & Format(dateWhy, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),NULL," & _
                    "To_Date('" & Format(dateWhy, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),'" & UserInfo.姓名 & "'," & _
                    "Null,To_Date('" & Format(dateWhy, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),'" & gstrUserName & "','0',Null,0,0)"
                zlDatabase.ExecuteProcedure gstrSql, "插入检验临时记录"
                
            End If
        
        End If
        
        If lng标本ID(intCol) > 0 Then
            gstrSql = "ZL_检验普通结果_BATCHUPDATE(" & lng标本ID(intCol) & "," & _
                lngDeviceID & ",Null,Null,Null,'" & Mid(str结果(intCol), 2) & "')"
            zlDatabase.ExecuteProcedure gstrSql, "检验结果报告"
            
            gstrSql = "ZL_检验质控记录_EDIT(1," & lng标本ID(intCol) & "," & lngQCID & ",Null,Null,Null,Null,Null,Null," & intCol + 1 & ")"
            zlDatabase.ExecuteProcedure gstrSql, "保存为质控品"
        End If
'        gcnOracle.CommitTrans
        blnBegin = False
        
    Next
    
    Exit Function
hErr:
    If blnBegin Then gcnOracle.RollbackTrans
    Call ErrCenter
End Function

Private Function GetSampleIDNO(ByVal lngDevId As Long, ByVal lngQC As Long, ByVal dBegin As Date, dEnd As Date, ByVal intC As Integer, ByRef lngSampleID As Long, ByRef strSampleNO As String)
    Dim strTmp As String, rsTemp As ADODB.Recordset, rsSampleNO As ADODB.Recordset
    
    On Error GoTo errH
    strTmp = "Select a.标本id, a.标本序号,b.名称, b.标本号, b.水平" & vbNewLine & _
            "From 检验质控记录 A, 检验质控品 B" & vbNewLine & _
            "Where 质控品id(+) = b.Id And b.Id = [1] And a.检验时间(+) between [2] and [3] And a.测试次数(+) = [4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strTmp, Me.Caption, lngQC, dBegin, dEnd, intC)
    Do Until rsTemp.EOF
        lngSampleID = Val("" & rsTemp!标本ID)
        strSampleNO = IIf(lngSampleID <= 0, Trim("" & rsTemp!标本号), Trim("" & rsTemp!标本序号))
        If strSampleNO = "" Or strSampleNO = "0" Then strSampleNO = rsTemp!名称 & "-" & (intC - 1)
        If lngSampleID <= 0 Then
            
            Call GenNo(lngDevId, intC - 1, dBegin, dEnd, rsTemp!名称, strSampleNO)
            
        End If
        rsTemp.MoveNext
    Loop
            
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub GenNo(ByVal lngDevId As Long, intC As Integer, dBegin As Date, dEnd As Date, strName As String, strSampleNO As String)
    Dim strTmp As String, rsTemp As ADODB.Recordset, rsSampleNO As ADODB.Recordset
    
    strTmp = "Select 测试次数 from 检验质控记录 where 仪器ID=[1] and 检验时间 between [2] and [3] And 标本序号=[4] "
    Set rsSampleNO = zlDatabase.OpenSQLRecord(strTmp, Me.Caption, lngDevId, dBegin, dEnd, strSampleNO)
    If Not rsSampleNO.EOF Then
        strSampleNO = strName & "-" & intC + 1
        Call GenNo(lngDevId, intC + 1, dBegin, dEnd, strName, strSampleNO)
    End If
End Sub

Private Sub LoadInstruments(ctrCbo As ComboBox, Optional intIndex As Integer)
    ' 取检验仪器数据到Cbo控件
    Dim strsql As String, rsTemp As ADODB.Recordset
    Dim lngMachineID As Long, lngIndex As Long
    On Error GoTo hErr
    
    lngMachineID = Val(zlDatabase.GetPara("仪器", glngSys, 1209, 0))
    If intIndex <> 0 Then lngIndex = intIndex
    
    If InStr(1, mstrPriv, "所有科室") > 0 Then
        strsql = " Select Distinct  a.id,a.编码 , a.名称  From 检验仪器 a ,部门表 b,检验质控品 c " & _
                  "Where a.使用小组ID = b.ID and a.id = c.仪器id"
        Set rsTemp = zlDatabase.OpenSQLRecord(strsql, gstrSysName)
        
    Else
        strsql = " Select Distinct a.id,a.编码 , a.名称  From 部门人员 D,检验仪器 a ,部门表 b , 检验质控品 c " & _
                  " Where a.使用小组ID = b.ID and a.使用小组id=D.部门id and D.人员id = [1]  " & _
                  " and a.id = c.仪器Id "
        Set rsTemp = zlDatabase.OpenSQLRecord(strsql, gstrSysName, UserInfo.ID)
    End If
    
    ctrCbo.Clear
    Do Until rsTemp.EOF
        ctrCbo.AddItem "" & rsTemp!编码 & " " & rsTemp!名称
        ctrCbo.ItemData(ctrCbo.NewIndex) = rsTemp!ID
        If lngMachineID = rsTemp!ID Then lngIndex = ctrCbo.NewIndex
        rsTemp.MoveNext
    Loop
    
    If ctrCbo.ListCount > 0 Then
        If lngIndex >= 0 And lngIndex < ctrCbo.ListCount Then
            ctrCbo.ListIndex = lngIndex
        Else
            ctrCbo.ListIndex = 0
        End If
    End If
    Exit Sub
hErr:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function GetQCControlData(ByRef vsGrid As VSFlexGrid, ByVal lng仪器id As Long, ByVal dateWhy As Date) As Boolean
    '取QCControl控件的数据
    Dim strsql As String, rsTemp As ADODB.Recordset
    On Error GoTo hErr
    
    strsql = "Select distinct ID,标本号,水平, 名称, 批号, 浓度,  To_Char(开始日期, 'yyyy-MM-dd') As 开始日期, To_Char(结束日期, 'yyyy-MM-dd') As 结束日期,b.质控取值   " & vbNewLine & _
            "From 检验质控品 a,检验质控品项目 b " & vbNewLine & _
            "Where a.id = b.质控品id and [1] Between a.开始日期 And a.结束日期 And a.仪器id = [2]" & vbNewLine & _
            "Order By 开始日期 Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, dateWhy, lng仪器id)
    
    With vsGrid
        .Clear
        .Rows = 2: .Cols = 9
        Set .DataSource = rsTemp
        
        If .Cols > 1 Then
            .ColWidth(0) = 0
            .ColHidden(0) = True
            .ColHidden(1) = True
            .ColHidden(8) = True
            If .Rows > 1 Then
                .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
            End If
            
        End If
        If Not rsTemp.EOF Then .AutoSize 2, .Cols - 1
            
      '  .Select .FixedRows, 1
    End With
    
    Exit Function
hErr:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetQCData(ByRef vsGrid As VSFlexGrid, ByVal lngQCID As Long, ByVal dateWhy As Date, strGetQCVal As String) As Boolean
    '取QC数据
    Dim strsql As String, rsTemp As ADODB.Recordset
    Dim dBegin As Date, dEnd As Date, iCol As Integer, iRow As Integer
    
    On Error GoTo hErr
    
    dBegin = Format(dateWhy, "yyyy-MM-dd 00:00:00")
    dEnd = Format(dateWhy, "yyyy-MM-dd 23:59:59")
    
    strsql = "Select Distinct  F.编码, F.中文名, E.缩写, '' as 第一次,'' as 第二次,'' as 第三次,'' as 第四次,'' as 第五次,'' as 第六次, '' as 第七次,'' as 第八次,'' as 第九次" & vbNewLine & _
            "       ,A.质控品id, A.项目id, A.取值序列, A.序列值, E.结果类型,Nvl(G.小数位数,2) as 小数位数, '' as 标本id,'' as 标本序号,'' as 检验人,'' 弃用记录,'' as 标记" & vbNewLine & _
            "From 检验质控品项目 A, 检验项目 E, 诊治所见项目 F,检验仪器项目 G" & vbNewLine & _
            "Where A.项目id = E.诊治项目id And A.项目id = F.ID And A.质控品id = [1] and A.项目id= G.项目ID and G.仪器ID=[2]" & vbNewLine & _
            "Order By F.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, lngQCID, Val(cbo仪器.ItemData(cbo仪器.ListIndex)))
    With vsGrid
        .Clear
        .Rows = 2: .Cols = 23
        
        Set .DataSource = rsTemp
        For iCol = 11 To .Cols - 1
            .ColHidden(iCol) = True
        Next
        
        If Not rsTemp.EOF Then .AutoSize 0, 11
 
        '取数据
        strsql = "Select f.编码, f.中文名, e.缩写, d.检验结果,d.od,d.sco, t.标记, e.结果类型, nvl(i.小数位数,2) as 小数位数, a.*" & vbNewLine & _
                "From (Select a.质控品id, a.项目id, c.标本序号, b.标本id, b.检验时间, a.取值序列, a.序列值, b.测试次数, b.检验人, b.弃用记录, b.仪器id" & vbNewLine & _
                "       From 检验质控品项目 A, 检验质控记录 B, 检验标本记录 C" & vbNewLine & _
                "       Where b.标本id = c.Id And a.质控品id = b.质控品id And a.质控品id = [1] And" & vbNewLine & _
                "             b.检验时间 Between [2] And [3]) A, 检验普通结果 D, 检验项目 E, 诊治所见项目 F, 检验质控报告 T, 检验仪器项目 I" & vbNewLine & _
                "Where d.Id = t.结果id(+) And a.标本id = d.检验标本id And a.项目id = d.检验项目id And a.项目id = e.诊治项目id And a.项目id = f.Id And" & vbNewLine & _
                "      a.仪器id = i.仪器id And a.项目id = i.项目id"

        Set mrsQCData = zlDatabase.OpenSQLRecord(strsql, Me.Caption, lngQCID, dBegin, dEnd)
        Do Until mrsQCData.EOF
            For iRow = .FixedRows To .Rows - 1
                If .TextMatrix(iRow, 0) = "" & mrsQCData!编码 Then
                     
                    If strGetQCVal = "[SCO]" Then
                        .TextMatrix(iRow, 2 + Val("" & mrsQCData!测试次数)) = Trim("" & mrsQCData!sco)
                        If Val("" & mrsQCData!标记) = 2 Then '失控(红)
                            .Cell(flexcpForeColor, iRow, 2 + Val("" & mrsQCData!测试次数)) = vbRed
                        ElseIf Val("" & mrsQCData!标记) = 0 Then '正常
                            .Cell(flexcpForeColor, iRow, 2 + Val("" & mrsQCData!测试次数)) = .ForeColor
                        Else  '警告(洋红)
                            .Cell(flexcpForeColor, iRow, 2 + Val("" & mrsQCData!测试次数)) = vbMagenta
                        End If
                        Exit For
                    ElseIf strGetQCVal = "[OD]" Then
                        .TextMatrix(iRow, 2 + Val("" & mrsQCData!测试次数)) = Trim("" & mrsQCData!od)
                        If Val("" & mrsQCData!标记) = 2 Then '失控(红)
                            .Cell(flexcpForeColor, iRow, 2 + Val("" & mrsQCData!测试次数)) = vbRed
                        ElseIf Val("" & mrsQCData!标记) = 0 Then '正常
                            .Cell(flexcpForeColor, iRow, 2 + Val("" & mrsQCData!测试次数)) = .ForeColor
                        Else  '警告(洋红)
                            .Cell(flexcpForeColor, iRow, 2 + Val("" & mrsQCData!测试次数)) = vbMagenta
                        End If
                        Exit For
                    Else
                        .TextMatrix(iRow, 2 + Val("" & mrsQCData!测试次数)) = Trim("" & mrsQCData!检验结果)
                        If Val("" & mrsQCData!标记) = 2 Then '失控(红)
                            .Cell(flexcpForeColor, iRow, 2 + Val("" & mrsQCData!测试次数)) = vbRed
                        ElseIf Val("" & mrsQCData!标记) = 0 Then '正常
                            .Cell(flexcpForeColor, iRow, 2 + Val("" & mrsQCData!测试次数)) = .ForeColor
                        Else  '警告(洋红)
                            .Cell(flexcpForeColor, iRow, 2 + Val("" & mrsQCData!测试次数)) = vbMagenta
                        End If
                        Exit For
                    End If
                    
                End If
            Next
            mrsQCData.MoveNext
        Loop
    End With
    Exit Function
hErr:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function






