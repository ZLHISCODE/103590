VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.Form frmSegmentList 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   4350
      Left            =   90
      TabIndex        =   0
      Top             =   465
      Width           =   3405
      _Version        =   589884
      _ExtentX        =   6006
      _ExtentY        =   7673
      _StockProps     =   0
      BorderStyle     =   2
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
   End
   Begin VB.PictureBox picView 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   2145
      Left            =   90
      ScaleHeight     =   2145
      ScaleWidth      =   3450
      TabIndex        =   1
      Top             =   4920
      Width           =   3450
      Begin XtremeReportControl.ReportControl rptView 
         Height          =   2025
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   3240
         _Version        =   589884
         _ExtentX        =   5715
         _ExtentY        =   3572
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2820
      Top             =   165
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
            Picture         =   "frmSegmentList.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSegmentList.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSegmentList.frx":0B34
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   15
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmSegmentList.frx":0ECE
      Left            =   645
      Top             =   75
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmSegmentList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    图标 = 0: ID: 编号: 名称
End Enum
Const con_UnDefine = -999
Const conPane_View = 201

'---------------------------------
'公共事件
Public Event RowDblClick(ByVal ROW As XtremeReportControl.IReportRow)   '双击一行或在行上按回车
Public Event ModifiedOrDeleted(Action As Integer)                       '修改或删除示范时

'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mlngDemoId As Long          '当前示范id

Private mfrmParent As Form          '父窗体
Private mlngFileID As Long          '定义文件id
Private mlngPatient As Long         '病人id，在病人病历编辑时，用来确定条件示范是否满足
Private mlngVisit As Long           '主页id或挂号单ID
Private mlngAdvice As Long          '医嘱ID

Private mintPower As Integer        '示范管理权范围
'    mintPower=con_UnDefine，未定义;
'    mintPower=-1，不具备示范管理权;
'    mintPower=0，全院，这时显示所有的示范，也可以更改;
'    mintPower=1，科室，这时显示全院通用示范(科室id is null)和所在科室公有或部门内人员私有的示范，但不能更改全院通用示范;
'    mintPower=2，个人，这时显示全院通用示范(科室id is null)和所在科室通用示范(人员id is null)和个人示范，仅个人示范可更改

'-----------------------------------------------------
'临时变量
'-----------------------------------------------------


Dim lngCount As Long

'-----------------------------------------------------
'以下为外部公共程序
'-----------------------------------------------------
Public Function zlRefresh(ByVal frmParent As Form) As Long
    '功能：根据指定文件，刷新列表
    '参数：
    If frmParent.Name <> "frmMain" Then zlRefresh = 0: Exit Function
    
    Err = 0: On Error Resume Next
    With frmParent.Document
        mlngFileID = .EPRFileInfo.ID
        mlngPatient = .EPRPatiRecInfo.病人ID
        mlngVisit = .EPRPatiRecInfo.主页ID
        mlngAdvice = .EPRPatiRecInfo.医嘱id
    End With
    Set mfrmParent = frmParent
    
    Err = 0: On Error GoTo 0
    zlRefresh = zlSubRefList(mlngDemoId)
End Function

'-----------------------------------------------------
'以下为内部公共程序
'-----------------------------------------------------
Private Function zlGetPower() As Integer
    '功能：获得当前用户的示范管理的权限
    '返回：示范管理权限数值
    If mintPower = con_UnDefine Then
        If InStr(1, gstrPrivsEpr, "全院病历范文") <> 0 Then
            mintPower = 0
        ElseIf InStr(1, gstrPrivsEpr, "科室病历范文") <> 0 Then
            mintPower = 1
        ElseIf InStr(1, gstrPrivsEpr, "个人病历范文") <> 0 Then
            mintPower = 2
        Else
            mintPower = -1
        End If
    End If
    zlGetPower = mintPower
End Function

Private Function zlSubRefList(Optional lngID As Long) As Long
    '功能：刷新装入清单，并定位到指定的记录上
Dim rsTemp As New ADODB.Recordset
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow

    gstrSQL = "Select l.Id, l.编号, l.名称, l.简码, l.通用级" & vbNewLine & _
            "From 病历范文目录 l, Table(Cast(f_Segment_Usable([1], [2], [3], [4]) As " & gstrDbOwner & ".t_Dic_Rowset)) u" & vbNewLine & _
            "Where l.文件id = [1] And Nvl(l.性质, 0) = [5] And l.Id = To_Number(u.编码)"
    Select Case mintPower
    Case 0
    Case 1
        gstrSQL = gstrSQL & " And" & vbNewLine & _
                "      (Nvl(L.通用级, 0) = 0 Or" & vbNewLine & _
                "      L.通用级 In (1, 2) And" & vbNewLine & _
                "      L.科室id In (Select R.部门id From 部门人员 R, 上机人员表 U Where R.人员id = U.人员id And U.用户名 = User))"

    Case Else
        gstrSQL = gstrSQL & " And" & vbNewLine & _
                "      (Nvl(L.通用级, 0) = 0 Or" & vbNewLine & _
                "      L.通用级 = 1 And" & vbNewLine & _
                "      L.科室id In (Select R.部门id From 部门人员 R, 上机人员表 U Where R.人员id = U.人员id And U.用户名 = User) Or" & vbNewLine & _
                "      L.通用级 = 2 And L.人员id In (Select U.人员id From 上机人员表 U Where U.用户名 = User))"
    End Select
    
    gstrSQL = gstrSQL & " Order By L.通用级 Desc, Lpad(L.编号,13,'0') "
    
    Err = 0: On Error GoTo errHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "frmSegmentList", mlngFileID, mlngPatient, mlngVisit, mlngAdvice, 1)
    
    Me.rptList.Records.DeleteAll
    With rsTemp
        Do While Not .EOF
            Set rptRcd = Me.rptList.Records.Add()
            Set rptItem = rptRcd.AddItem(CInt(Val("" & !通用级))): rptItem.Icon = rptItem.Value
            rptRcd.AddItem CLng(!ID)
            rptRcd.AddItem CStr("" & !编号)
            rptRcd.AddItem CStr("" & !名称)
            .MoveNext
        Loop
    End With
    Me.rptList.Populate
    If lngID <> 0 Then
        For Each rptRow In Me.rptList.Rows
            If Val(rptRow.Record(mCol.ID).Value) = lngID Then
                Set Me.rptList.FocusedRow = rptRow: Exit For
            End If
        Next
    End If
    If Me.rptList.Rows.Count > 0 And (Me.rptList.FocusedRow Is Nothing) Then
        Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
    End If
    Call rptList_SelectionChanged
    zlSubRefList = Me.rptList.Records.Count
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlSubRefList = Me.rptList.Records.Count
End Function

'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRetuId As Long, strTemp As String
    
    '------------------------------------
    Select Case Control.ID
    Case conMenu_Edit_Modify
        lngRetuId = frmEPRModelEdit.ShowMe(mfrmParent, False, CByte(mintPower), mlngFileID, mlngDemoId)
        If lngRetuId = 0 Then Exit Sub
        Call zlSubRefList(lngRetuId)
        RaiseEvent ModifiedOrDeleted(1)
    Case conMenu_Edit_Delete
        Err = 0: On Error GoTo errHand
        strTemp = "真的删除该示范吗？" & vbCrLf & "——" & Me.rptList.FocusedRow.Record(mCol.名称).Value
        If MsgBox(strTemp, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        gstrSQL = "zl_病历范文目录_delete(" & mlngDemoId & ")"
        zlDatabase.ExecuteProcedure gstrSQL, "词句列表"
        With Me.rptList
            mlngDemoId = 0: lngRetuId = .FocusedRow.Index
            If .Rows.Count > lngRetuId + 1 Then
                mlngDemoId = .Rows(lngRetuId + 1).Record(mCol.ID).Value
            ElseIf lngRetuId > 0 Then
                mlngDemoId = .Rows(lngRetuId - 1).Record(mCol.ID).Value
            End If
        End With
        Call zlSubRefList(mlngDemoId)
        RaiseEvent ModifiedOrDeleted(2)
    Case conMenu_Edit_Request
        If frmEPRModelRequest.ShowMe(Me, mlngDemoId, mintPower) = True Then Call zlSubRefList(mlngDemoId)
    Case conMenu_View_Refresh
        Call zlSubRefList(mlngDemoId)
    End Select
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub

Private Sub cbsThis_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    Err = 0: On Error Resume Next
    With Me.rptList
        .Left = Left: .Width = Right
        .Top = Top: .Height = Bottom - Top
    End With
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_Request
        Control.Visible = (mintPower >= 0)
        Control.Enabled = (mlngDemoId <> 0)
        If Control.Enabled Then Control.Enabled = (Me.rptList.FocusedRow.Record(mCol.图标).Value >= mintPower)
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_View: Item.Handle = Me.picView.hwnd
    End Select
End Sub

Private Sub Form_Load()
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim rptCol As ReportColumn
    '-----------------------------------------------------
    '权限限制串复制，避免同时进入其他模块而导致gmstrPrivs变化，导致控制无效
    mintPower = con_UnDefine
    mintPower = zlGetPower
    mlngDemoId = 0
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
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
    Me.cbsThis.ActiveMenuBar.Title = "菜单": Me.cbsThis.ActiveMenuBar.Visible = False
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Request, "条件(&Q)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&V)"): cbrControl.BeginGroup = True
    End With
    
    '-----------------------------------------------------
    '设置示范显示停靠窗格
    Dim panThis As Pane
    Set panThis = dkpMan.CreatePane(conPane_View, 450, 150, DockBottomOf, Nothing)
    panThis.Title = "示范内容"
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = False
    
    '-----------------------------------------------------
    With Me.rptList
        Set rptCol = .Columns.Add(mCol.图标, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False
        rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.编号, "编号", 50, False): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.名称, "名称", 120, True): rptCol.Editable = False: rptCol.Groupable = False
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
    
    With Me.rptView
        Set rptCol = .Columns.Add(0, "提纲", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .ShowHeader = False
        .PreviewMode = True
        With .PaintManager
            .NoItemsText = "没有可显示的内容..."
            .SetPreviewIndent 18, 0, 8, 6
        End With
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mfrmParent = Nothing
    imgList.ListImages.Clear
    ImageList_Destroy imgList.hImageList
End Sub

Private Sub picView_Resize()
    Err = 0: On Error Resume Next
    With Me.rptView
        .Left = 0: .Width = Me.picView.ScaleWidth
        .Top = 0: .Height = Me.picView.ScaleHeight
    End With
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.rptList
        If .Visible = False Then Exit Sub
        If .FocusedRow Is Nothing Then Exit Sub
        If .FocusedRow.GroupRow Then Exit Sub
        Call rptList_RowDblClick(.FocusedRow, .FocusedRow.Record.Item(mCol.ID))
    End With
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
Dim cbrPopupBar As CommandBar
Dim cbrPopupItem As CommandBarControl
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
    
    If Button <> vbRightButton Then Exit Sub
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.FindControl(xtpControlPopup, conMenu_EditPopup)
    If cbrMenuBar Is Nothing Then Exit Sub
    If cbrMenuBar.Visible = False Then Exit Sub
    
    Set cbrPopupBar = Me.cbsThis.Add("弹出菜单", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
    Next
    cbrPopupBar.ShowPopup
End Sub

Private Sub rptList_RowDblClick(ByVal ROW As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Me.rptList.FocusedRow Is Nothing Then
        mlngDemoId = 0
    ElseIf Me.rptList.FocusedRow.GroupRow = True Then
        mlngDemoId = 0
    Else
        mlngDemoId = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
    End If
    If mlngDemoId = 0 Then Exit Sub
    RaiseEvent RowDblClick(Me.rptList.FocusedRow)
End Sub

Private Sub rptList_SelectionChanged()
Dim rsTemp As New ADODB.Recordset
Dim rsView As New ADODB.Recordset, strVSql As String, strView As String
Dim rptRcd As ReportRecord
    
    If Me.rptList.FocusedRow Is Nothing Then
        mlngDemoId = 0
    ElseIf Me.rptList.FocusedRow.GroupRow = True Then
        mlngDemoId = 0
    Else
        mlngDemoId = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
    End If
    If Me.Visible = False Then Exit Sub

    '刷新示范内容
    gstrSQL = "Select Id, 内容文本 From 病历范文内容 Where 文件id = [1] And 对象类型 = 1 Order By 对象序号"
    strVSql = "Select Id, 对象类型, 内容文本, 是否换行, 要素名称" & vbNewLine & _
            "From 病历范文内容" & vbNewLine & _
            "Where 文件id = [1] And 父id = [2]" & vbNewLine & _
            "Order By 对象序号"
    
    Err = 0: On Error GoTo errHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "frmSegmentList", mlngDemoId)
    Me.rptView.Records.DeleteAll
    Do While Not rsTemp.EOF
        Set rptRcd = Me.rptView.Records.Add()
        rptRcd.AddItem CStr(rsTemp!内容文本 & ":")
        strView = ""
        Set rsView = zlDatabase.OpenSQLRecord(strVSql, "frmSegmentList", mlngDemoId, CLng(rsTemp!ID))
        Do While Not rsView.EOF
            Select Case rsView!对象类型
            Case 2: strView = strView & rsView!内容文本
            Case 3, 5: strView = strView & vbCrLf & "□" & vbCrLf
            Case 4: strView = strView & "[" & IIf(Trim("" & rsView!内容文本) = "", rsView!要素名称, rsView!内容文本) & "]"
            Case 7: strView = strView & "<" & rsView!内容文本 & ">"
            End Select
            strView = strView & IIf(Val("" & rsView!是否换行) = 1, vbCrLf, "")
            rsView.MoveNext
        Loop
        rptRcd.PreviewText = Trim(strView)
        rsTemp.MoveNext
    Loop
    Me.rptView.Populate
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

