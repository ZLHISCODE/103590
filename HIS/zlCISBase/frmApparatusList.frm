VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmApparatusList 
   Caption         =   "检验仪器管理"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   11295
   Icon            =   "frmApparatusList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7260
   ScaleWidth      =   11295
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picList 
      BackColor       =   &H00FFEBD7&
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   90
      ScaleHeight     =   5295
      ScaleWidth      =   4425
      TabIndex        =   2
      Top             =   450
      Width           =   4425
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   4560
         Left            =   15
         TabIndex        =   3
         Top             =   15
         Width           =   4395
         _Version        =   589884
         _ExtentX        =   7752
         _ExtentY        =   8043
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   1845
         Top             =   4800
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
               Picture         =   "frmApparatusList.frx":058A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmApparatusList.frx":0B24
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6885
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmApparatusList.frx":10BE
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14843
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
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   900
      Left            =   180
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5625
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
      Bindings        =   "frmApparatusList.frx":1950
      Left            =   945
      Top             =   105
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmApparatusList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    图标 = 0: ID: 编码: 名称: 简码: 仪器类型: 使用部门: 微生物: 转换仪器
End Enum

Const conPane_List = 201
Const conPane_Info = 202
Const conPane_Item = 203
Const conPane_Rule = 204

'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mstrPrivs As String     '当前使用者权限串

Private mfrmInfo As frmApparatusInfo
Attribute mfrmInfo.VB_VarHelpID = -1
Private mfrmItem As frmApparatusItem
Private mfrmGerm As frmApparatusGerm

Private mintEditState As Integer    '当前编辑状态：0-非编辑状态,1-基本信息编辑,2-项目通道编辑,3-细菌通道编辑
Private mLngAptId As Long, mbln微生物 As Boolean
Private mLngEditWidth As Long       '为适应大字体情况下窗体变大.先读入窗体大小.

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

Private mLngCount As Long
Private mstr仪器数量 As String '注册码中限制的检验仪器数量
Private mstr可使用仪器 As String '仪器数量<>""情况下有效

'-----------------------------------------------------
'以下为内部公共程序
'-----------------------------------------------------
Public Function zlRefList(Optional lngAptId As Long) As Long
    '功能：刷新装入指定种类的病历文件清单，并定位到指定的文件上
    Dim rsTemp As New ADODB.Recordset
    Dim blnAdd As Boolean
    
    gstrSql = "Select A.ID, A.编码, A.名称, A.简码, A.仪器类型, D.名称 As 使用部门, A.微生物, A.转换日期, T.名称 As 转换仪器," & vbNewLine & _
            "       Count(S.项目id) As 是否失控" & vbNewLine & _
            "From 检验仪器 A, 部门表 D, 检验仪器 T, 检验仪器状态 S" & vbNewLine & _
            "Where A.使用小组id = D.ID(+) And A.转换仪器id = T.ID(+) And A.ID = S.仪器id(+)" & vbNewLine & _
            "Group By A.ID, A.编码, A.名称, A.简码, A.仪器类型, D.名称, A.微生物, A.转换日期, T.名称"
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Me.rptList.Records.DeleteAll
    mLngCount = 0
    With rsTemp
        Do While Not .EOF
            blnAdd = True
            If mstr仪器数量 <> "" Then
                If Val(mstr仪器数量) > 0 Then
                    If InStr(mstr可使用仪器 & ",", "," & !ID & ",") <= 0 Then
                        blnAdd = False
                    End If
                Else
                    blnAdd = False
                End If
            End If
                    
            If blnAdd Then
                Set rptRcd = Me.rptList.Records.Add()
                If Val("" & !是否失控) = 0 Then
                    Set rptItem = rptRcd.AddItem("0"): rptItem.Icon = 0
                Else
                    Set rptItem = rptRcd.AddItem("1"): rptItem.Icon = 1
                End If
                rptRcd.AddItem CStr(!ID)
                rptRcd.AddItem CStr(!编码)
                rptRcd.AddItem CStr(!名称)
                rptRcd.AddItem CStr("" & !简码)
                rptRcd.AddItem CStr("" & !仪器类型)
                rptRcd.AddItem CStr("" & !使用部门)
                rptRcd.AddItem CStr(Val("" & !微生物))
                If IsNull(!转换日期) Then
                    rptRcd.AddItem CStr("")
                Else
                    rptRcd.AddItem Format(!转换日期, "MM-dd hh:mm") & "→" & CStr(!转换仪器)
                End If

            End If
            .MoveNext
        Loop
    End With
    With Me.rptList
        .Populate
    End With
    
    If lngAptId <> 0 Then
        For Each rptRow In Me.rptList.Rows
            If Val(rptRow.Record(mCol.ID).Value) = lngAptId Then
                Set Me.rptList.FocusedRow = rptRow: Exit For
            End If
        Next
    End If
    If Me.rptList.Rows.Count > 0 And (Me.rptList.FocusedRow Is Nothing) Then
        Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
    End If
    Call rptList_SelectionChanged
    zlRefList = Me.rptList.Records.Count
    Me.stbThis.Panels(2).Text = "共有" & Me.rptList.Records.Count & "台仪器"
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
    objPrint.Title.Text = "检验设备清单"
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
    Dim lngRetuId As Long
    
    '------------------------------------
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_Exit: Unload Me
    
    Case conMenu_Edit_Save
        lngRetuId = 0
        Select Case mintEditState   '0-非编辑状态,1-基本信息编辑,2-项目通道编辑,3-细菌通道编辑
        Case 1
            lngRetuId = mfrmInfo.zlEditSave()
            If lngRetuId <> 0 Then mLngAptId = lngRetuId
        Case 2:
            lngRetuId = mfrmItem.zlEditSave()
        Case 3:
            lngRetuId = mfrmGerm.zlEditSave()
        End Select
        If lngRetuId <> 0 Then Call zlRefList(mLngAptId): mintEditState = 0: Me.picList.Enabled = True: Me.rptList.SetFocus
    Case conMenu_Edit_Untread:
        Select Case mintEditState   '0-非编辑状态,1-基本信息编辑,2-项目通道编辑,3-细菌通道编辑
        Case 1: Call mfrmInfo.zlEditCancel
        Case 2: Call mfrmItem.zlEditCancel
        Case 3: Call mfrmGerm.zlEditCancel
        End Select
        mintEditState = 0: Me.picList.Enabled = True: Me.rptList.SetFocus
    Case conMenu_Edit_NewItem
        If mfrmInfo.zlEditStart(True, mLngAptId) Then mintEditState = 1: Me.picList.Enabled = False
        Me.dkpMan.FindPane(conPane_Info).Select
    Case conMenu_Edit_Modify
        If mLngAptId = 0 Then Exit Sub
        If mfrmInfo.zlEditStart(False, mLngAptId) Then mintEditState = 1: Me.picList.Enabled = False
        Me.dkpMan.FindPane(conPane_Info).Select
    Case conMenu_Edit_Delete
        If mLngAptId = 0 Then Exit Sub
        With Me.rptList
            If MsgBox("真的删除该检验仪器吗？" & vbCrLf & "——" & .FocusedRow.Record(mCol.名称).Value, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                gstrSql = "Zl_检验仪器_Delete(" & mLngAptId & ")"
                Err = 0: On Error GoTo ErrHand
                Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
                
                Err = 0: On Error GoTo 0
                mLngAptId = 0: lngRetuId = .FocusedRow.Index
                If .Rows.Count > lngRetuId + 1 Then
                    If .Rows(lngRetuId + 1).GroupRow = False Then mLngAptId = .Rows(lngRetuId + 1).Record(mCol.ID).Value
                ElseIf lngRetuId > 0 Then
                    If .Rows(lngRetuId - 1).GroupRow = False Then mLngAptId = .Rows(lngRetuId - 1).Record(mCol.ID).Value
                End If
                Call Me.zlRefList(mLngAptId)
            End If
        End With
        Exit Sub
    
    Case conMenu_Edit_Compend
        If mbln微生物 Then
            If mfrmGerm.zlEditStart Then mintEditState = 3: Me.picList.Enabled = False
        Else
            If mfrmItem.zlEditStart Then mintEditState = 2: Me.picList.Enabled = False
        End If
        Me.dkpMan.FindPane(conPane_Item).Select
    
    Case conMenu_Edit_Pause
        Call frmApparatusTran.ShowMe(Me, mLngAptId)
        Call Me.zlRefList(mLngAptId)
    
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
        Call zlRefList(mLngAptId)
    
    Case conMenu_Help_Help:     Call ShowHelp(gstrLisHelp, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
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
    
    Dim lngAptId As Long
    If Me.rptList.FocusedRow Is Nothing Then
        lngAptId = 0
    ElseIf Me.rptList.FocusedRow.GroupRow = True Then
        lngAptId = 0
    Else
        lngAptId = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
    End If
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel: Control.Enabled = (Me.rptList.Records.Count <> 0 And mintEditState = 0)
    Case conMenu_Edit_Save, conMenu_Edit_Untread: Control.Enabled = (mintEditState <> 0)
    
    Case conMenu_Edit_NewItem
    
     Control.Enabled = (InStr(1, mstrPrivs, "增删改") > 0 And mintEditState = 0)
     If Control.Enabled = True Then
        If mstr仪器数量 <> "" Then
            If Val(mstr仪器数量) <= 0 Then
                Control.Enabled = False
            ElseIf Me.rptList.Records.Count >= Val(mstr仪器数量) Then
                Control.Enabled = False
            End If
        End If
     End If
    Case conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_Compend, conMenu_Edit_Pause
        Control.Enabled = (InStr(1, mstrPrivs, "增删改") > 0 And mintEditState = 0 And lngAptId <> 0)
        If Control.Enabled = True And Control.ID <> conMenu_Edit_Delete Then
            If mstr仪器数量 <> "" Then
                If Val(mstr仪器数量) <= 0 Then
                    Control.Enabled = False
                ElseIf InStr(mstr可使用仪器 & ",", "," & lngAptId & ",") <= 0 Then
                    Control.Enabled = False
                End If
            End If
        End If
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_List
        Item.Handle = Me.picList.hWnd
    Case conPane_Info
        If mfrmInfo Is Nothing Then Set mfrmInfo = New frmApparatusInfo
        Item.Handle = mfrmInfo.hWnd
    Case conPane_Item
        If mbln微生物 Then
            If mfrmGerm Is Nothing Then Set mfrmGerm = New frmApparatusGerm
            Item.Handle = mfrmGerm.hWnd
        Else
            If mfrmItem Is Nothing Then Set mfrmItem = New frmApparatusItem
            Item.Handle = mfrmItem.hWnd
        End If
    End Select
End Sub

Private Sub Form_Load()
    '-----------------------------------------------------
    '权限限制串复制，避免同时进入其他模块而导致gstrPrivs变化，导致控制无效
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    mLngEditWidth = frmApparatusInfo.ScaleWidth
    
    mstrPrivs = gstrPrivs
    
    mintEditState = 0
    mLngAptId = 0: mbln微生物 = False
    Err = 0: On Error GoTo ErrHand
    mstr仪器数量 = zlRegInfo("检验仪器数量")
    mstr可使用仪器 = ""
    If mstr仪器数量 <> "" Then
        If Val(mstr仪器数量) > 0 Then
            strSql = "select * From (select ID From 检验仪器 Order by ID) Where Rownum<=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(mstr仪器数量))
            Do Until rsTmp.EOF
                mstr可使用仪器 = mstr可使用仪器 & "," & rsTmp!ID
                rsTmp.MoveNext
            Loop
        End If
        Me.Caption = Me.Caption & "(" & mstr仪器数量 & "台仪器)"
    End If
    
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Compend, "项目通道(&W)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Pause, "仪器转换(&T)"): cbrControl.BeginGroup = True
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
        .Add FCONTROL, Asc("W"), conMenu_Edit_Compend
        .Add FCONTROL, Asc("T"), conMenu_Edit_Pause
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Compend, "项目通道"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '设置词句显示停靠窗格
    Dim panThis As Pane, panSub As Pane
    
    If mfrmInfo Is Nothing Then Set mfrmInfo = New frmApparatusInfo
    If mfrmItem Is Nothing Then Set mfrmItem = New frmApparatusItem
    If mfrmGerm Is Nothing Then Set mfrmGerm = New frmApparatusGerm
    
    Set panThis = dkpMan.CreatePane(conPane_List, 450, 580, DockLeftOf, Nothing)
    panThis.Title = "检验仪器列表"
    panThis.Options = PaneNoCaption
    
    Set panThis = dkpMan.CreatePane(conPane_Info, 550, 580, DockRightOf, Nothing)
    panThis.Title = "仪器基本信息"
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable

    Set panSub = dkpMan.CreatePane(conPane_Item, 550, 800, DockBottomOf, panThis)
    panSub.Title = "仪器项目通道"
    panSub.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable

    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    
    '-----------------------------------------------------
    With Me.rptList
        .AutoColumnSizing = (Screen.Width / Screen.TwipsPerPixelX > 1024)   '必须在列设置之前设置，才能生效
        Set rptCol = .Columns.Add(mCol.图标, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False
        rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.编码, "编码", 60, False): rptCol.Editable = False: rptCol.Groupable = False
        .SortOrder.Add rptCol: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.名称, "名称", 150, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.简码, "简码", 100, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.仪器类型, "仪器类型", 100, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.使用部门, "使用部门", 90, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.微生物, "微生物", 0, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.转换仪器, "转换仪器", 180, True): rptCol.Editable = False: rptCol.Groupable = False
        
        .SetImageList Me.imgList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With
    
    '-----------------------------------------------------
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)
    '-----------------------------------------------------
    '数据装入
    Call zlRefList
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Dim panBase As Pane
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Width - mLngEditWidth / Screen.TwipsPerPixelX - picList.Width <= 9000 Then Exit Sub
    
    Set panBase = Me.dkpMan.FindPane(conPane_Info)
    panBase.MinTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, 280
    panBase.MaxTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, 280
    
    
    Me.dkpMan.RecalcLayout
    Me.dkpMan.NormalizeSplitters

    panBase.MinTrackSize.SetSize 0, 0
    panBase.MaxTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, 280
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmInfo
    Unload mfrmItem
    Unload mfrmGerm
    Set mfrmInfo = Nothing
    Set mfrmItem = Nothing
    Set mfrmGerm = Nothing
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub optScope_Click(Index As Integer)
    Dim lngAptId As Long
    
    If Me.rptList.FocusedRow Is Nothing Then
        lngAptId = 0
    ElseIf Me.rptList.FocusedRow.GroupRow = True Then
        lngAptId = 0
    Else
        lngAptId = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
    End If
    
    Call Me.zlRefList(lngAptId)
End Sub

Private Sub picList_Resize()
    With Me.rptList
        .Left = Me.picList.ScaleLeft: .Width = Me.picList.ScaleWidth - .Left
        .Top = Me.picList.ScaleTop: .Height = Me.picList.ScaleHeight - .Top
    End With
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.rptList.Visible = False Then Exit Sub
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    If Me.rptList.FocusedRow.GroupRow Then Exit Sub
    Call rptList_RowDblClick(Me.rptList.FocusedRow, Me.rptList.FocusedRow.Record.Item(mCol.ID))
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
    Dim lngAptId As Long
    If Me.rptList.FocusedRow Is Nothing Then
        lngAptId = 0
    ElseIf Me.rptList.FocusedRow.GroupRow = True Then
        lngAptId = 0
    Else
        lngAptId = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
    End If
    If lngAptId = 0 Then Exit Sub
    
    Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Modify)
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)

End Sub

Private Sub rptList_SelectionChanged()
    If Me.rptList.FocusedRow Is Nothing Then
        mLngAptId = 0
    Else
        mLngAptId = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
        mbln微生物 = (Me.rptList.FocusedRow.Record.Item(mCol.微生物).Value = "1")
    End If
    
    Dim panThis As Pane
    Set panThis = Me.dkpMan.FindPane(conPane_Item)
    If mbln微生物 Then
        If panThis.Handle <> mfrmGerm.hWnd Then
            panThis.Handle = mfrmGerm.hWnd
            mfrmItem.Visible = False
            Me.dkpMan.RecalcLayout
        End If
    Else
        If panThis.Handle <> mfrmItem.hWnd Then
            mfrmItem.Visible = True
            panThis.Handle = mfrmItem.hWnd
            Me.dkpMan.RecalcLayout
        End If
    End If
    
    Call mfrmInfo.zlRefresh(mLngAptId)
    If mbln微生物 Then
        Call mfrmGerm.zlRefresh(mLngAptId)
    Else
        Call mfrmItem.zlRefresh(mLngAptId)
    End If
End Sub

