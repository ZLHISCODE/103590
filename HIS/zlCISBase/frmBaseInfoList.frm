VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmBaseInfoList 
   Caption         =   "基础数据管理"
   ClientHeight    =   7530
   ClientLeft      =   165
   ClientTop       =   870
   ClientWidth     =   11670
   Icon            =   "frmBaseInfoList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   11670
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picDesc 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   5160
      ScaleHeight     =   2295
      ScaleWidth      =   3135
      TabIndex        =   5
      Top             =   3720
      Width           =   3135
      Begin VB.Image imgNote 
         Height          =   240
         Left            =   120
         Picture         =   "frmBaseInfoList.frx":058A
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"frmBaseInfoList.frx":0B14
         ForeColor       =   &H00008000&
         Height          =   9000
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   2460
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picType 
      BackColor       =   &H00FFEBD7&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   5160
      ScaleHeight     =   3375
      ScaleWidth      =   3135
      TabIndex        =   3
      Top             =   360
      Width           =   3135
      Begin XtremeSuiteControls.ShortcutBar sbType 
         Height          =   3255
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   3015
         _Version        =   589884
         _ExtentX        =   5318
         _ExtentY        =   5741
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picList 
      BackColor       =   &H00FFEBD7&
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   720
      ScaleHeight     =   5295
      ScaleWidth      =   4425
      TabIndex        =   0
      Top             =   360
      Width           =   4425
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   4410
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   4395
         _Version        =   589884
         _ExtentX        =   7752
         _ExtentY        =   7779
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   0
         Top             =   4680
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBaseInfoList.frx":0F10
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   7155
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBaseInfoList.frx":14AA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15505
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
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
      Left            =   0
      Top             =   480
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
            Picture         =   "frmBaseInfoList.frx":1D3C
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseInfoList.frx":2194
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseInfoList.frx":25E6
            Key             =   "Read"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   0
      Top             =   1140
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
            Picture         =   "frmBaseInfoList.frx":2900
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseInfoList.frx":2D58
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBaseInfoList.frx":31AA
            Key             =   "Read"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   900
      Left            =   720
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5760
      Visible         =   0   'False
      Width           =   1080
      _cx             =   1905
      _cy             =   1587
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmBaseInfoList.frx":34C4
      Left            =   945
      Top             =   105
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmBaseInfoList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const conPane_Type = 201
Const conPane_List = 202
Const conPane_Edit = 203
Const conPane_Desc = 204
'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mstrPrivs As String     '当前使用者权限串
Private mfrmEdit As frmBaseInfoEdit

Private mintEditState As Integer    '当前编辑状态：0-非编辑状态,1-编辑状态
Private mstr编码 As String

'-----------------------------------------------------
'临时变量
'-----------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar

Dim rptCol As ReportColumn
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow

Dim lngCount As Long

'-----------------------------------------------------
'以下为内部公共程序
'-----------------------------------------------------
Public Function zlRefList(strItemName As String) As Long
    '功能：刷新装入指定种类的病历文件清单，并定位到指定的文件上
    Dim i As Integer
    Dim rsTemp As New ADODB.Recordset
        
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select * from " & strItemName & " order by to_number(编码)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Me.rptList.Records.DeleteAll
    With rsTemp
        Do While Not .EOF
            Set rptRcd = Me.rptList.Records.Add()
            For i = 0 To .Fields.Count - 1
                If i = 0 Then
                    Set rptItem = rptRcd.AddItem(CStr(Nvl(.Fields(i)))): rptItem.Icon = 0
                End If
                If .Fields(i).Name = "缺省标志" Then
                    Set rptItem = rptRcd.AddItem(IIf(CStr(Nvl(.Fields(i))) = 1, "√", "")): rptItem.SortPriority = Val(("" & Nvl(.Fields(i))))
                Else
                    Set rptItem = rptRcd.AddItem(CStr(Nvl(.Fields(i))))   ': rptItem.SortPriority = Val(("" & Nvl(.Fields(i))))
                End If
            Next
            .MoveNext
        Loop
    End With
    With Me.rptList
        .GroupsOrder.DeleteAll
        .Populate
    End With

    If mstr编码 <> "" Then
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow = False Then
                If Val(rptRow.Record(1).Value) = mstr编码 Then
                    Set Me.rptList.FocusedRow = rptRow
                    Exit For
                End If
            End If
        Next
    End If
    If Me.rptList.FocusedRow Is Nothing And Me.rptList.Rows.Count > 0 Then
        If Me.rptList.Rows(0).GroupRow Then
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0).Childs(0)
        Else
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
        End If
    End If
    Call rptList_SelectionChanged

    zlRefList = Me.rptList.Records.Count
    Me.stbThis.Panels(2).Text = "【" & strItemName & "】共有" & Me.rptList.Records.Count & "条记录"
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefList = Me.rptList.Records.Count
End Function

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    If Me.rptList.Records.Count = 0 Then Exit Sub

    '-------------------------------------------------
    '复制数据表格
    If zlControl.RPTCopyToVSF(Me.rptList, Me.vfgList) Is Nothing Then Exit Sub
    '-------------------------------------------------
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow

    Set objPrint.Body = Me.vfgList
    objPrint.Title.Text = "检验质控规则"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)

    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim str编码 As String
    Dim lngRetuId As Long
    Dim panThis As Pane
    '------------------------------------
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_Exit: Unload Me: Unload frmMedTech: Unload frmMedTreat

    Case conMenu_Edit_Save:
        str编码 = mfrmEdit.zlEditSave(gstrItemName)
        If str编码 <> "" Then
            ShowEdit False
            mstr编码 = str编码: Call zlRefList(gstrItemName)
            mintEditState = 0: Me.picList.Enabled = True: Me.rptList.SetFocus
        End If
    Case conMenu_Edit_Untread:
        ShowEdit False
        Call mfrmEdit.zlEditCancel
        mintEditState = 0: Me.picList.Enabled = True: Me.rptList.SetFocus
        
    Case conMenu_Edit_NewItem
        mfrmEdit.fraEdit.BackColor = vbWhite
        ShowEdit True
        
        If mstr编码 = "" Then Exit Sub
        If mfrmEdit.zlEditStart(True, gstrItemName, mstr编码) = False Then Exit Sub
        mintEditState = 1: Me.picList.Enabled = False
        
    Case conMenu_Edit_Modify
        mfrmEdit.fraEdit.BackColor = vbWhite
        ShowEdit True
        
        If mstr编码 = "" Then Exit Sub
        If mfrmEdit.zlEditStart(False, gstrItemName, mstr编码) = False Then Exit Sub
        mintEditState = 1: Me.picList.Enabled = False
        
    Case conMenu_Edit_Delete
        Dim strMsg As String
        With Me.rptList
            strMsg = "真的删除该项目记录吗？" & vbCrLf & "——" & .FocusedRow.Record(2).Value
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

            gstrSql = "zl_" & gstrItemName & "_Edit(3,NULL,'" & mstr编码 & "')"

            Err = 0: On Error GoTo ErrHand
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)

            Err = 0: On Error GoTo 0
            mstr编码 = "": lngRetuId = .FocusedRow.Index
            If .Rows.Count > lngRetuId + 1 Then
                lngRetuId = lngRetuId + 1
            ElseIf lngRetuId > 0 Then
                lngRetuId = lngRetuId - 1
            End If
            If .Rows(lngRetuId).GroupRow = False Then mstr编码 = .Rows(lngRetuId).Record(1).Value
            Call Me.zlRefList(gstrItemName)
        End With
'
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_Refresh
        Call zlRefList(gstrItemName)

    Case conMenu_Help_Help:     Call ShowHelp(gstrLisHelp, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    End Select

    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If

    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (Me.rptList.Records.Count <> 0 And mintEditState = 0)
    Case conMenu_Edit_Save, conMenu_Edit_Untread
        Control.Enabled = (mintEditState <> 0)
    Case conMenu_Edit_NewItem
        Control.Enabled = (InStr(1, mstrPrivs, "增删改") > 0 And mintEditState = 0)
    Case conMenu_Edit_Modify, conMenu_Edit_Delete
        Control.Enabled = (InStr(1, mstrPrivs, "增删改") > 0 And mintEditState = 0 And Me.rptList.Rows.Count)
        If Control.Enabled Then Control.Enabled = mstr编码 <> ""
        If Control.Enabled Then Control.Enabled = Not Me.rptList.FocusedRow.GroupRow
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    Case conMenu_View_Find, conMenu_View_Refresh, conMenu_View_Option: Control.Enabled = (mintEditState = 0)
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_Type
        Item.Handle = picType.hWnd
    Case conPane_List
        Item.Handle = Me.picList.hWnd
    Case conPane_Edit
        If mfrmEdit Is Nothing Then Set mfrmEdit = New frmBaseInfoEdit
        Item.Handle = mfrmEdit.fraEdit.hWnd
    Case conPane_Desc
        Item.Handle = Me.picDesc.hWnd
    End Select
End Sub

Private Sub Form_Load()
    '-----------------------------------------------------
    '权限限制串复制，避免同时进入其他模块而导致gstrPrivs变化，导致控制无效
    mstrPrivs = gstrPrivs
    
    mintEditState = 0
    mstr编码 = 0
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '菜单定义
    Me.cbsThis.ActiveMenuBar.Title = "菜单"
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With

    '快键绑定
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("Z"), conMenu_Edit_Untread
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FSHIFT, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With

    '设置不常用菜单
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next

    '-----------------------------------------------------
    '设置词句显示停靠窗格
    Dim panType As Pane, panList As Pane, panEdit As Pane, panDesc As Pane
    If mfrmEdit Is Nothing Then Set mfrmEdit = New frmBaseInfoEdit

    Set panType = dkpMan.CreatePane(conPane_Type, 160, 1000, DockLeftOf, Nothing)
    panType.Title = "基础信息类型"
    panType.Options = PaneNoCaption Or PaneNoHideable Or PaneNoCloseable Or PaneNoFloatable
    
    Set panList = dkpMan.CreatePane(conPane_List, 800, 800, DockRightOf, panType)
    panList.Title = "基础信息列表"
    panList.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set panEdit = dkpMan.CreatePane(conPane_Edit, 800, 200, DockBottomOf, panList)
    panEdit.Title = "基础信息编辑"
    panEdit.Options = PaneNoCaption
    panEdit.Close

    Set panDesc = dkpMan.CreatePane(conPane_Desc, 200, 1000, DockRightOf, Nothing)
    panDesc.Title = "分类说明"
    panDesc.Options = PaneNoCloseable

    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    Call CreateShortCutBar
    
    Call DrawRpt(gstrItemName)          '动态加载 rptControl
    Call zlRefList(gstrItemName)        '数据装入
    Call LoadControl(gstrItemName)      '动态加载编辑控件

    '界面恢复
    '-----------------------------------------------------
    Call RestoreWinState(Me, App.ProductName)
    Exit Sub
ErrHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmEdit
    Set mfrmEdit = Nothing
    Call SaveWinState(Me, App.ProductName)
    Unload Me
End Sub

Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With Me.rptList
        .Left = Me.picList.ScaleLeft: .Width = Me.picList.ScaleWidth - .Left
        .Top = Me.picList.ScaleTop: .Height = Me.picList.ScaleHeight - .Top
    End With
    mfrmEdit.fraEdit.Width = mfrmEdit.Width
    mfrmEdit.fraEdit.Height = Me.ScaleHeight - Me.picList.ScaleHeight
End Sub

Private Sub picType_Resize()
    Err = 0: On Error Resume Next
    With Me.sbType
        .Left = Me.picType.ScaleLeft: .Width = Me.picType.ScaleWidth - .Left
        .Top = Me.picType.ScaleTop: .Height = Me.picType.ScaleHeight - .Top
    End With
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.rptList.Visible = False Then Exit Sub
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    If Me.rptList.FocusedRow.GroupRow Then Exit Sub
    Call rptList_RowDblClick(Me.rptList.FocusedRow, Me.rptList.FocusedRow.Record.Item(0))
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl

    If Button <> vbRightButton Then Exit Sub
    If Me.cbsThis.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = Me.cbsThis.Add("弹出菜单", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
    Next
    cbrPopupBar.ShowPopup
End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If mstr编码 = 0 Then Exit Sub

    Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Modify)
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)

End Sub

Private Sub rptList_SelectionChanged()
    With Me.rptList
        If .FocusedRow Is Nothing Then
            mstr编码 = 0
        ElseIf .FocusedRow.GroupRow = True Then
            mstr编码 = 0
        Else
            mstr编码 = Trim(.FocusedRow.Record.Item(0).Value)
        End If
        Call mfrmEdit.zlRefresh(gstrItemName, mstr编码)
    End With
End Sub

Private Sub CreateShortCutBar()
    Dim objItem As ShortcutBarItem
    Dim objItemMain As ShortcutBarItem
      
    Set objItemMain = sbType.AddItem(1, "医技工作", frmMedTreat.hWnd)
    Set objItem = sbType.AddItem(2, "医疗工作", frmMedTech.hWnd)
    
    sbType.Selected = objItemMain
    sbType.ExpandedLinesCount = sbType.ItemCount
End Sub

Private Sub DrawRpt(strItemName As String)
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    
    On Err GoTo ErrHand:
    
    gstrSql = "select * from " & strItemName & " where rownum = 1 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)

    rptList.Columns.DeleteAll
    rptList.AutoColumnSizing = (Screen.Width / Screen.TwipsPerPixelX > 800)   '必须在列设置之前设置，才能生效
    
    Set rptCol = rptList.Columns.Add(0, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: _
                rptCol.Alignment = xtpAlignmentCenter
                
    For i = 0 To rsTemp.Fields.Count - 1
        Set rptCol = rptList.Columns.Add(i + 1, "" & rsTemp.Fields(i).Name, 85, True): rptCol.Editable = False: rptCol.Groupable = False
    Next
    

    rptList.SetImageList Me.ils16
    rptList.AllowColumnRemove = False
    rptList.MultipleSelection = False
    rptList.ShowItemsInGroups = False
    With rptList.PaintManager
        .ColumnStyle = xtpColumnFlat
        .GridLineColor = RGB(225, 225, 225)
        .NoGroupByText = "拖动列标题到这里,按该列性质..."
        .NoItemsText = "没有可显示的项目..."
        .VerticalGridStyle = xtpGridSolid
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub ShowItemInfo(strItemName As String)
    gstrItemName = strItemName
    Call DrawRpt(gstrItemName)
    Call zlRefList(gstrItemName)
    Call LoadControl(gstrItemName)
End Sub

Private Sub LoadControl(strItemName As String)
    Dim objControl As Object
    Dim int间隔 As Integer: int间隔 = 570
    Dim intLblAndTxt As Integer: intLblAndTxt = 210

    '将所有隐藏的控件全部显示
    For Each objControl In mfrmEdit.Controls
        If objControl.Visible = False Then
            objControl.Visible = True
        End If
    Next

    '恢复控件原有大小
    mfrmEdit.txt编码.Left = 1050
    mfrmEdit.txt编码.Top = 360
    mfrmEdit.txt编码.Width = 1380

    mfrmEdit.txt名称.Left = 3570
    mfrmEdit.txt名称.Top = 360
    mfrmEdit.txt名称.Width = 2235

    mfrmEdit.cbo适用性别.Left = 7035
    mfrmEdit.cbo适用性别.Top = 360
    mfrmEdit.cbo适用性别.Width = 2235

    mfrmEdit.txt简码.Left = 7035
    mfrmEdit.txt简码.Top = 360
    mfrmEdit.txt简码.Width = 1215

    mfrmEdit.txt说明.Left = 1050
    mfrmEdit.txt说明.Top = 840
    mfrmEdit.txt说明.Width = 3285
    mfrmEdit.txt说明.Height = 720

    mfrmEdit.txt管码.Left = 1050
    mfrmEdit.txt管码.Top = 840
    mfrmEdit.txt管码.Width = 615

    mfrmEdit.chk缺省标志.Left = 4980
    mfrmEdit.chk缺省标志.Top = 840
    mfrmEdit.chk缺省标志.Width = 3255

    mfrmEdit.cbo分类.Left = 10290
    mfrmEdit.cbo分类.Top = 360
    mfrmEdit.cbo分类.Width = 1395
    
    mfrmEdit.cbo分类1.Left = 10290
    mfrmEdit.cbo分类1.Top = 360
    mfrmEdit.cbo分类1.Width = 1395
    mfrmEdit.cbo分类1.Visible = False
    
    mfrmEdit.lbl1.Left = 6375
    mfrmEdit.lbl1.Top = 420
    mfrmEdit.lbl说明.Caption = "说明"

    '根据不同的选项重新加载界面控件
    With mfrmEdit
        Select Case Trim(strItemName)
        Case "诊疗检验标本"
            .txt说明.Visible = False
            .txt管码.Visible = False
            .chk缺省标志.Visible = False
            .lbl说明.Visible = False
            .cbo分类.Visible = False
            .lbl1.Caption = "简码"
            .txt简码.Left = .cbo适用性别.Left
            .txt简码.Top = .cbo适用性别.Top
            .lbl分类.Left = .txt简码.Left + .txt简码.Width + int间隔
            .lbl分类.Caption = "适用性别"
            .lbl分类.Width = 2 * .lbl编码.Width
            .cbo适用性别.Left = .lbl分类.Left + .lbl分类.Width + intLblAndTxt
            .cbo适用性别.Top = .cbo分类.Top
            .cbo适用性别.Width = 800
            
            .txt编码.MaxLength = 2
            .txt名称.MaxLength = 20
            .txt简码.MaxLength = 8
            
        Case "诊疗检验类型"
            .txt说明.Visible = False
            .lbl分类.Visible = False
            .cbo分类.Visible = False
            .cbo适用性别.Visible = False
            .lbl1.Caption = "简码"
            .txt简码.Left = .cbo适用性别.Left
            .txt简码.Top = .cbo适用性别.Top
            .lbl说明.Caption = "管码"
            .txt管码.Left = .txt说明.Left
            .txt管码.Top = .txt说明.Top
            .chk缺省标志.Left = .lbl名称.Left
            
            .txt编码.MaxLength = 2
            .txt名称.MaxLength = 20
            .txt简码.MaxLength = 8
            .txt管码.MaxLength = 2
            
        Case "检验备注文字"
            .cbo适用性别.Visible = False
            .txt管码.Visible = False
            .chk缺省标志.Visible = False
            .lbl1.Caption = "简码"
            .lbl说明.Caption = "说明"
            .txt简码.Left = .cbo适用性别.Left
            .txt简码.Top = .cbo适用性别.Top
            .lbl分类.Caption = "分类"
            .cbo分类.Left = .lbl分类.Left + .lbl分类.Width + intLblAndTxt
            
            .txt编码.MaxLength = 10
            .txt名称.MaxLength = 100
            .txt简码.MaxLength = 10
            .txt说明.MaxLength = 80

        Case "检验培养文字"
            .txt管码.Visible = False
            .chk缺省标志.Visible = False
            .lbl分类.Visible = False
            .cbo分类.Visible = False
            .cbo适用性别.Visible = False
            .txt名称.Width = .txt名称.Width * 2
            .lbl1.Caption = "简码"
            .lbl1.Left = .txt名称.Left + .txt名称.Width + int间隔
            .txt简码.Left = .lbl1.Left + .lbl1.Width + intLblAndTxt
            .txt简码.Top = .cbo适用性别.Top
            .lbl说明.Caption = "说明"
            
            .txt编码.MaxLength = 10
            .txt名称.MaxLength = 100
            .txt简码.MaxLength = 10
            .txt说明.MaxLength = 80

        Case "检验评语文字"
            .cbo适用性别.Visible = False
            .txt管码.Visible = False
            .chk缺省标志.Visible = False
            .lbl1.Caption = "简码"
            .lbl分类.Caption = "分类"
            .txt简码.Left = .cbo适用性别.Left
            .txt简码.Top = .cbo适用性别.Top
            .lbl说明.Caption = "说明"
            .cbo分类.Left = .lbl分类.Left + .lbl分类.Width + intLblAndTxt
            
            .txt编码.MaxLength = 3
            .txt名称.MaxLength = 50
            .txt简码.MaxLength = 10
            .txt说明.MaxLength = 80
            
        Case "检验标本形态"
            .cbo适用性别.Visible = False
            .lbl1.Visible = False
            .txt简码.Visible = False
            .txt管码.Visible = False
            .cbo分类.Visible = False
            .chk缺省标志.Visible = False
            .lbl分类.Visible = False
            
            .txt编码.MaxLength = 10
            .txt名称.MaxLength = 50
            .txt说明.MaxLength = 100

        Case "检验分析用途"
            .txt说明.Visible = False
            .lbl1.Visible = False
            .cbo适用性别.Visible = False
            .txt简码.Visible = False
            .txt管码.Visible = False
            .cbo分类.Visible = False
            .chk缺省标志.Visible = False
            .lbl说明.Visible = False
            .lbl分类.Visible = False
            .txt名称.Width = .txt名称.Width * 2
            
            .txt编码.MaxLength = 10
            .txt名称.MaxLength = 200
            
        Case "检验拒收理由"
            .lbl1.Visible = False
            .cbo适用性别.Visible = False
            .txt简码.Visible = False
            .txt管码.Visible = False
            .cbo分类.Visible = False
            .chk缺省标志.Visible = False
            .lbl分类.Visible = False
            .lbl名称.Visible = False
            .txt名称.Visible = False
            .lbl说明.Caption = "名称"
            
            .txt编码.MaxLength = 10
            .txt说明.MaxLength = 200
            
        Case "检验审核类别"
            .cbo适用性别.Visible = False
            .lbl说明.Visible = False
            .txt说明.Visible = False
            .txt管码.Visible = False
            .lbl分类.Visible = False
            .cbo分类.Visible = False
            .lbl1.Caption = "简码"
            .txt简码.Left = .cbo适用性别.Left
            .txt简码.Top = .cbo适用性别.Top
            .chk缺省标志.Left = .lbl说明.Left
            .chk缺省标志.Top = .lbl说明.Top
            
            .txt编码.MaxLength = 2
            .txt名称.MaxLength = 20
            .txt简码.MaxLength = 8

        Case "检验细菌类别"
            .cbo适用性别.Visible = False
            .lbl说明.Visible = False
            .txt说明.Visible = False
            .txt管码.Visible = False
            .lbl分类.Visible = False
            .cbo分类.Visible = False
            .lbl1.Caption = "简码"
            .txt简码.Left = .cbo适用性别.Left
            .txt简码.Top = .cbo适用性别.Top
            .chk缺省标志.Left = .lbl说明.Left
            .chk缺省标志.Top = .lbl说明.Top
            
            .txt编码.MaxLength = 8
            .txt名称.MaxLength = 30
            .txt简码.MaxLength = 20

        Case "检验细菌菌属"
            .cbo适用性别.Visible = False
            .lbl说明.Visible = False
            .txt说明.Visible = False
            .txt管码.Visible = False
            .chk缺省标志.Visible = False
            .lbl分类.Visible = False
            .cbo分类.Visible = False
            .lbl1.Caption = "简码"
            .txt简码.Left = .cbo适用性别.Left
            .txt简码.Top = .cbo适用性别.Top
            
            .txt编码.MaxLength = 8
            .txt名称.MaxLength = 30
            .txt简码.MaxLength = 20

        Case "革兰染色分类"
            .cbo适用性别.Visible = False
            .lbl说明.Visible = False
            .txt说明.Visible = False
            .txt管码.Visible = False
            .lbl分类.Visible = False
            .cbo分类.Visible = False
            .lbl1.Caption = "简码"
            .txt简码.Left = .cbo适用性别.Left
            .txt简码.Top = .cbo适用性别.Top
            .chk缺省标志.Left = .lbl说明.Left
            .chk缺省标志.Top = .lbl说明.Top
            
            .txt编码.MaxLength = 8
            .txt名称.MaxLength = 30
            .txt简码.MaxLength = 20

        Case "质控报告词句"
            .txt说明.Visible = False
            .txt管码.Visible = False
            .chk缺省标志.Visible = False
            .cbo分类.Visible = False
            .txt名称.Width = .txt名称.Width * 3
            .lbl说明.Caption = "简码"
            .txt简码.Left = .txt说明.Left
            .txt简码.Top = .txt说明.Top
            .lbl1.Caption = "分组"
            .lbl1.Left = .lbl名称.Left
            .lbl1.Top = .lbl说明.Top
            .cbo适用性别.Left = .lbl1.Left + .lbl1.Width + intLblAndTxt
            .cbo适用性别.Top = .lbl1.Top
            .cbo适用性别.Width = 1000
            .cbo分类1.Left = .cbo适用性别.Left
            .cbo分类1.Top = .cbo适用性别.Top
            .cbo适用性别.Width = .cbo适用性别.Width
            .cbo分类1.Visible = True
            .cbo适用性别.Visible = False
            
            .txt编码.MaxLength = 3
            .txt名称.MaxLength = 80
            .txt简码.MaxLength = 10
            
        Case "质控检验方法"
            .cbo适用性别.Visible = False
            .lbl说明.Visible = False
            .txt说明.Visible = False
            .txt管码.Visible = False
            .chk缺省标志.Visible = False
            .lbl分类.Visible = False
            .cbo分类.Visible = False
            .lbl1.Caption = "简码"
            .txt简码.Left = .cbo适用性别.Left
            .txt简码.Top = .cbo适用性别.Top
            
            .txt编码.MaxLength = 6
            .txt名称.MaxLength = 30
            .txt简码.MaxLength = 10

        Case "质控试剂来源"
            .cbo适用性别.Visible = False
            .txt说明.Visible = False
            .chk缺省标志.Visible = False
            .lbl说明.Visible = False
            .cbo分类.Visible = False
            .lbl1.Caption = "简码"
            .txt简码.Left = .cbo适用性别.Left
            .txt简码.Top = .cbo适用性别.Top
            .lbl分类.Caption = "QC编码"
            .txt管码.Left = .lbl分类.Left + .lbl分类.Width + intLblAndTxt
            .txt管码.Top = .cbo分类.Top
            .txt管码.Width = 800
            
            .txt编码.MaxLength = 6
            .txt名称.MaxLength = 30
            .txt简码.MaxLength = 10
            .txt管码.MaxLength = 8

        Case "检验结果描述"
            .txt说明.Visible = False
            .txt管码.Visible = False
            .chk缺省标志.Visible = False
            .lbl说明.Visible = False
            .cbo适用性别.Visible = False
            .lbl1.Caption = "简码"
            .txt简码.Left = .cbo适用性别.Left
            .txt简码.Top = .cbo适用性别.Top
            .cbo分类.Width = 1000
            .cbo分类1.Width = .cbo分类.Width
            .cbo分类.Left = .lbl分类.Left + .lbl分类.Width + intLblAndTxt
            .cbo分类1.Left = .cbo分类.Left
            .cbo分类.Visible = False
            .cbo分类1.Visible = True
            .lbl分类.Caption = "分类"
            
            .txt编码.MaxLength = 3
            .txt名称.MaxLength = 200
            .txt简码.MaxLength = 20
            
            
        Case "细菌检测方法"
            .cbo适用性别.Visible = False
            .lbl说明.Visible = False
            .txt说明.Visible = False
            .txt管码.Visible = False
            .chk缺省标志.Visible = False
            .lbl分类.Visible = False
            .cbo分类.Visible = False
            .lbl1.Caption = "简码"
            .txt简码.Left = .cbo适用性别.Left
            .txt简码.Top = .cbo适用性别.Top
            
            .txt编码.MaxLength = 2
            .txt名称.MaxLength = 20
            .txt简码.MaxLength = 10
            
        Case "细菌耐药机制"
            .cbo适用性别.Visible = False
            .lbl说明.Visible = False
            .txt说明.Visible = False
            .txt管码.Visible = False
            .chk缺省标志.Visible = False
            .lbl分类.Visible = False
            .cbo分类.Visible = False
            .lbl1.Caption = "简码"
            .txt简码.Left = .cbo适用性别.Left
            .txt简码.Top = .cbo适用性别.Top
            
            .txt编码.MaxLength = 4
            .txt名称.MaxLength = 100
            .txt简码.MaxLength = 20
            
        End Select
    End With
End Sub

Private Sub sbType_ExpandButtonDown(CancelMenu As Boolean)
    CancelMenu = True
End Sub

Private Sub sbType_SelectedChanged(ByVal Item As XtremeSuiteControls.IShortcutBarItem)
    Dim i As Integer
    Select Case Item.ID
        Case "1"
            For i = 1 To frmMedTreat.tplFunc.Groups(1).Items.Count
                If frmMedTreat.tplFunc.Groups(1).Items(i).Selected Then
                    Call ShowItemInfo(frmMedTreat.tplFunc.Groups(1).Items(i).Caption)
                    Exit For
                End If
            Next
            
        Case "2"
            For i = 1 To frmMedTech.tplFunc.Groups(1).Items.Count - 1
                If frmMedTech.tplFunc.Groups(1).Items(i).Selected Then
                    Call ShowItemInfo(frmMedTech.tplFunc.Groups(1).Items(i).Caption)
                    Exit For
                End If
            Next
    End Select
End Sub

Private Sub ShowEdit(blnShow As Boolean)
    '功能       是否显示登记窗体
    Dim Pane1 As Pane
    Set Pane1 = dkpMan.FindPane(conPane_Edit)
    If blnShow = True Then
        Pane1.Select
    Else
        Pane1.Close
    End If
    dkpMan.RecalcLayout
End Sub




