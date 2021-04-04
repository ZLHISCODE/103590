VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTendItemMan 
   Caption         =   "护理记录项目管理"
   ClientHeight    =   6750
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10305
   Icon            =   "frmTendItem.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   10305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   2040
      Left            =   435
      TabIndex        =   0
      Top             =   2295
      Width           =   1995
      _Version        =   589884
      _ExtentX        =   3519
      _ExtentY        =   3598
      _StockProps     =   0
      BorderStyle     =   2
      ShowGroupBox    =   -1  'True
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VB.PictureBox picColorItem 
      BackColor       =   &H00E4E8EA&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   6225
      ScaleHeight     =   255
      ScaleWidth      =   2520
      TabIndex        =   4
      Top             =   4470
      Width           =   2520
      Begin VB.Label lblColor 
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   735
         TabIndex        =   6
         Top             =   45
         Width           =   765
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "记录色："
         Height          =   180
         Left            =   30
         TabIndex        =   5
         Top             =   45
         Width           =   720
      End
   End
   Begin XtremeSuiteControls.TaskPanel tkp 
      Height          =   3030
      Left            =   6630
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1290
      Width           =   3045
      _Version        =   589884
      _ExtentX        =   5371
      _ExtentY        =   5345
      _StockProps     =   64
      VisualTheme     =   5
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   6375
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTendItem.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15266
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
   Begin MSComctlLib.ImageList ilsList 
      Left            =   8715
      Top             =   5280
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
            Picture         =   "frmTendItem.frx":0E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendItem.frx":767E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendItem.frx":77D8
            Key             =   "User"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfPrint 
      Height          =   540
      Left            =   6885
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5610
      Visible         =   0   'False
      Width           =   1095
      _cx             =   1931
      _cy             =   952
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
      DesignerControls=   "frmTendItem.frx":7932
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   690
      Top             =   60
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmTendItemMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'窗口级变量申明########################################################################################################

Private Enum mCol
    图标 = 0
    项目序号
    分组名称
    项目名称
    项目类型
    项目长度
    项目小数
    项目单位
    项目表示
    项目值域
    最低护理
    体温项目
    保留项目
    适用病人
    应用方式
    项目性质
    应用场合
    项目id
    说明
End Enum

Private mstrPrivs As String      '当前使用者权限串
Private mblnStartUp As Boolean
Private mstrSQL As String
Private mblnOK As Boolean
'Private mblnShowStop As Boolean

'自定义过程/函数申明###################################################################################################

Public Sub zlRptPrint(ByVal bytMode As Byte)
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    If rptList.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '复制数据表格
    If zlReportToVSFlexGrid(vsfPrint, rptList) = False Then Exit Sub
    
    '-------------------------------------------------
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    
    Set objPrint.Body = vsfPrint
    
    objPrint.Title.Text = "护理记录项目清单"
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

Private Sub RefreshStateInfo()
    '------------------------------------------------------------------------------------------------------------------
    '功能：刷新状态栏显示信息
    '------------------------------------------------------------------------------------------------------------------
    
    stbThis.Panels(2).Text = "共有 " & rptList.Records.Count & " 个护理记录项目！"
    
End Sub

Private Function InitGrid() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：初始化报表控件
    '------------------------------------------------------------------------------------------------------------------
    Dim rptCol As ReportColumn
    
    With rptList
        
        Set rptCol = .Columns.Add(mCol.图标, "", 20, False)
        rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        
        Set rptCol = .Columns.Add(mCol.项目序号, "序号", 49, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.分组名称, "分组名称", 100, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        
        Set rptCol = .Columns.Add(mCol.项目名称, "名称", 100, True): rptCol.Editable = False: rptCol.Groupable = False
        
        Set rptCol = .Columns.Add(mCol.项目类型, "类型", 49, False): rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(mCol.项目长度, "长度", 49, False): rptCol.Editable = False: rptCol.Groupable = False
        rptCol.Alignment = xtpAlignmentRight
        
        Set rptCol = .Columns.Add(mCol.项目小数, "小数", 49, False): rptCol.Editable = False: rptCol.Groupable = False
        rptCol.Alignment = xtpAlignmentRight
        
        Set rptCol = .Columns.Add(mCol.项目单位, "单位", 49, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.项目表示, "表示", 49, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.项目值域, "值域", 160, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.最低护理, "最低护理", 80, False): rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(mCol.体温项目, "体温", 49, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.保留项目, "保留", 49, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.适用病人, "适用病人", 75, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.应用方式, "应用方式", 75, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.项目性质, "项目性质", 75, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.应用场合, "应用场合", 75, False): rptCol.Editable = False: rptCol.Groupable = False
        
        rptCol.Alignment = xtpAlignmentCenter
        
        Set rptCol = .Columns.Add(mCol.项目id, "项目id", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        
        Set rptCol = .Columns.Add(mCol.说明, "项目说明", 200, True): rptCol.Editable = False: rptCol.Groupable = False
        
        .SetImageList ilsList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GridLineColor = RGB(225, 225, 225)
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
        End With
        .PreviewMode = True
        
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns.Find(mCol.分组名称)
        .GroupsOrder(0).SortAscending = True
        .SortOrder.Add .Columns.Find(mCol.项目序号)
    End With
    
    InitGrid = True
    
End Function

Private Function CreateToolBox() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    Dim objGrp As TaskPanelGroup
    Dim objItem As TaskPanelGroupItem
    Dim objIlsItem As Object
    
    Call tkp.SetImageList(ilsList)

    Set objGrp = tkp.Groups.Add(0, "要素属性")
    objGrp.Expandable = False
    
    Set objItem = objGrp.Items.Add(0, "编  码：", xtpTaskItemTypeText)
    Call objGrp.Items.Add(0, "中文名：", xtpTaskItemTypeText)
    Call objGrp.Items.Add(0, "英文名：", xtpTaskItemTypeText)
    Call objGrp.Items.Add(0, "类  型：", xtpTaskItemTypeText)
             
    Set objGrp = tkp.Groups.Add(1, "临床意义")
    objGrp.Expandable = False
    Call objGrp.Items.Add(1, "", xtpTaskItemTypeText)
    
    Call tkp.SetImageList(ilsList)
    Set objGrp = tkp.Groups.Add(2, "体温属性")
    objGrp.Expandable = False
    Call objGrp.Items.Add(2, "排列号：", xtpTaskItemTypeText)
    Call objGrp.Items.Add(2, "记录名：", xtpTaskItemTypeText)
    Call objGrp.Items.Add(2, "记录法：", xtpTaskItemTypeText)
    Call objGrp.Items.Add(2, "记录符：", xtpTaskItemTypeText)
    
    'Set objItem = objGrp1.Items.Add(0, rs("记录名").Value & "(" & rs("记录符").Value & ")", xtpTaskItemTypeLink, ils16.ListImages("K" & NVL(rs("记录色"))).Index)

    Call objGrp.Items.Add(2, "记录色：", xtpTaskItemTypeControl)
    Call objGrp.Items.Add(2, "最小值：", xtpTaskItemTypeText)
    Call objGrp.Items.Add(2, "最大值：", xtpTaskItemTypeText)
    Call objGrp.Items.Add(2, "单位值：", xtpTaskItemTypeText)
    Call objGrp.Items.Add(2, "最高行：", xtpTaskItemTypeText)
    Call objGrp.Items.Add(2, "记录频次：", xtpTaskItemTypeText)
      
    Set tkp.Groups(3).Items(5).Control = picColorItem
    
    Set objGrp = tkp.Groups.Add(3, "适用科室")
    objGrp.Expandable = False
    Call objGrp.Items.Add(3, "内科，外科，麻麻科，尼科，屁屁科", xtpTaskItemTypeText)
    
    tkp.Animation = xtpTaskPanelAnimationNo
    tkp.Behaviour = xtpTaskPanelBehaviourExplorer
    tkp.HotTrackStyle = xtpTaskPanelHighlightItem
    
    tkp.SetGroupInnerMargins 0, 1, 1, 1
    
    tkp.AllowDrag = False
    tkp.SelectItemOnFocus = False

    tkp.Groups(1).Expanded = True
    
    
    CreateToolBox = True
    
End Function


Private Function zlMenuClick(ByVal strMenuItem As String, Optional ByVal strParam As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：功能处理
    '------------------------------------------------------------------------------------------------------------------
    Dim lngKey As Long
    Dim rptRcd As ReportRecord
    Dim rptItem As ReportRecordItem
    Dim objItem As Object
    Dim strTmp As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand

    
    Select Case strMenuItem
    Case "读取数据"
                
        mstrSQL = " SELECT A.项目序号," & _
                          "A.分组名 As 分组名称," & _
                          "A.项目名称," & _
                          "Decode(A.项目类型,3,'逻辑',2,'日期',1,'文字','数值') As 项目类型," & _
                          "A.项目长度," & _
                          "A.项目小数," & _
                          "A.项目单位,Decode(A.项目性质,1,'固定项目','活动项目') As 项目性质," & _
                          "Decode(A.项目表示,1,'上下',2,'单选',3,'复选',4,'汇总',5,'选择','文本') As 项目表示," & _
                          "A.项目值域," & _
                          "Decode(A.护理等级,1,'一级护理',2,'二级护理',3,'三级护理','特级护理') As 最低护理," & _
                          "Decode(C.项目序号,Null,'','√') As 体温项目," & _
                          "A.保留项目,Decode(A.适用病人,0,'所有',1,'病人',2,'婴儿') As 适用病人,Decode(A.应用方式,0,'禁用使用',1,'单独使用',2,'与脉搏共用','') As 应用方式," & _
                          "Decode(A.应用场合,1,'体温单',2,'记录单','通用') As 应用场合," & _
                          "A.项目id,A.说明 " & _
                     "FROM 护理记录项目 A,体温记录项目 C WHERE C.项目序号(+)=A.项目序号 Order By A.分组名,A.项目序号"
        
        Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption)
        If rs.BOF = False Then
            rptList.Records.DeleteAll
            
            Do While Not rs.EOF
                
                Set rptRcd = rptList.Records.Add()

                Set rptItem = rptRcd.AddItem("")
                rptItem.Icon = IIf(Val(NVL(rs("项目id"))) > 0, 1, 0)
                
                rptRcd.AddItem Zero(NVL(rs("项目序号")))
                rptRcd.AddItem NVL(rs("分组名称"))
                rptRcd.AddItem NVL(rs("项目名称"))
                rptRcd.AddItem NVL(rs("项目类型"))
                
                rptRcd.AddItem Zero(NVL(rs("项目长度")))
                rptRcd.AddItem Zero(NVL(rs("项目小数")))
                rptRcd.AddItem NVL(rs("项目单位"))
                rptRcd.AddItem NVL(rs("项目表示"))
                rptRcd.AddItem NVL(rs("项目值域"))
                rptRcd.AddItem NVL(rs("最低护理"))
                rptRcd.AddItem NVL(rs("体温项目"))
                rptRcd.AddItem IIf(NVL(rs("保留项目")) = 1, "√", "")
                rptRcd.AddItem NVL(rs("适用病人"))
                rptRcd.AddItem NVL(rs("应用方式"))
                rptRcd.AddItem NVL(rs("项目性质"))
                rptRcd.AddItem NVL(rs("应用场合"))
                rptRcd.AddItem Zero(NVL(rs("项目id")))
                rptRcd.AddItem NVL(rs("说明"))
                rs.MoveNext
            Loop
            
            rptList.Populate
            
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "读取要素"
               
        tkp.Groups(1).Items(1).Caption = "编  码："
        tkp.Groups(1).Items(2).Caption = "中文名："
        tkp.Groups(1).Items(3).Caption = "英文名："
        tkp.Groups(1).Items(4).Caption = "类  型："
        
        tkp.Groups(2).Items(1).Caption = ""
        
        If Not (rptList.FocusedRow Is Nothing) Then
            If Not (rptList.FocusedRow.Record Is Nothing) Then lngKey = Val(rptList.FocusedRow.Record.Item(mCol.项目id).Value)
        End If
                
        mstrSQL = "Select 编码,中文名,英文名,Decode(类型,1,'文字',2,'日期',3,'逻辑','数值') As 类型,临床意义 From 诊治所见项目 Where ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, lngKey)
        If rs.BOF = False Then
            tkp.Groups(1).Items(1).Caption = "编  码：" & zlCommFun.NVL(rs("编码"))
            tkp.Groups(1).Items(2).Caption = "中文名：" & zlCommFun.NVL(rs("中文名"))
            tkp.Groups(1).Items(3).Caption = "英文名：" & zlCommFun.NVL(rs("英文名"))
            tkp.Groups(1).Items(4).Caption = "类  型：" & zlCommFun.NVL(rs("类型"))
            tkp.Groups(2).Items(1).Caption = zlCommFun.NVL(rs("临床意义"))
                        
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "读取适用科室"
        
        If Not (rptList.FocusedRow Is Nothing) Then
            If Not (rptList.FocusedRow.Record Is Nothing) Then lngKey = Val(rptList.FocusedRow.Record.Item(mCol.项目序号).Value)
        End If
        
        tkp.Groups(4).Items(1).Caption = " "
        mstrSQL = "Select 适用科室 From 护理记录项目 Where 项目序号=[1]"
        Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, lngKey)
        If rs.BOF = False Then
            Select Case zlCommFun.NVL(rs("适用科室"), 0)
            Case 0
                tkp.Groups(4).Items(1).Caption = "该项目暂时不使用"
            Case 1
                tkp.Groups(4).Items(1).Caption = "该项目全院通用"
            Case 2
                mstrSQL = "Select b.名称 From 护理适用科室 a,部门表 b Where a.项目序号=[1] And a.科室id=b.ID"
                Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, lngKey)
                If rs.BOF = False Then
                    strTmp = ""
                    Do While Not rs.EOF
                        strTmp = strTmp & "、" & zlCommFun.NVL(rs("名称"))
                        rs.MoveNext
                    Loop
                    If strTmp <> "" Then strTmp = Mid(strTmp, 2)
                    tkp.Groups(4).Items(1).Caption = strTmp
                End If
                
            End Select
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "读取体温"
    
        tkp.Groups(3).Items(1).Caption = "排列号："
        tkp.Groups(3).Items(2).Caption = "记录名："
        tkp.Groups(3).Items(3).Caption = "记录法："
        tkp.Groups(3).Items(4).Caption = "记录符："
        tkp.Groups(3).Items(5).Caption = "记录色："
        tkp.Groups(3).Items(6).Caption = "最小值："
        tkp.Groups(3).Items(7).Caption = "最大值："
        tkp.Groups(3).Items(8).Caption = "单位值："
        tkp.Groups(3).Items(9).Caption = "最高行："
        tkp.Groups(3).Items(10).Caption = "记录频次："
        lblColor.BackStyle = 0
        
        If Not (rptList.FocusedRow Is Nothing) Then
            If Not (rptList.FocusedRow.Record Is Nothing) Then lngKey = Val(rptList.FocusedRow.Record.Item(mCol.项目序号).Value)
        End If
        
        mstrSQL = "Select 排列序号,记录名,Decode(记录法,1,'曲线',2,'表格') As 记录法,记录符,记录色,最小值,最大值,单位值,最高行,记录频次 From 体温记录项目 Where 项目序号=[1]"
        Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, lngKey)
        If rs.BOF = False Then
        
            tkp.Groups(3).Items(1).Caption = "排列号：" & zlCommFun.NVL(rs("排列序号"))
            tkp.Groups(3).Items(2).Caption = "记录名：" & zlCommFun.NVL(rs("记录名"))
            tkp.Groups(3).Items(3).Caption = "记录法：" & zlCommFun.NVL(rs("记录法"))
            
            If lngKey = 1 Then
                tkp.Groups(3).Items(4).Caption = "记录符：" & zlCommFun.NVL(rs("记录符").Value, "·,×,○")
            Else
                tkp.Groups(3).Items(4).Caption = "记录符：" & zlCommFun.NVL(rs("记录符").Value)
            End If
            
            '产生颜色
            On Error Resume Next
            Set objItem = Nothing
            Set objItem = ilsList.ListImages("K" & NVL(rs("记录色"), 0))
            If objItem Is Nothing Then Call SetColorIcon(Me, "K" & NVL(rs("记录色"), 0), NVL(rs("记录色"), 0), ilsList)
            On Error GoTo 0
            
            
            tkp.Groups(3).Items(5).Caption = "记录色：" & zlCommFun.NVL(rs("记录色"))
'            If zlCommFun.NVL(rs("记录色"), -1) = -1 Then
'                lblColor.BackStyle = 0
'            Else
                lblColor.BackStyle = 1
                lblColor.BackColor = zlCommFun.NVL(rs("记录色"), 0)
'            End If

            tkp.Groups(3).Items(6).Caption = "最小值：" & zlCommFun.NVL(rs("最小值"))
                        
            tkp.Groups(3).Items(7).Caption = "最大值：" & zlCommFun.NVL(rs("最大值"))
            tkp.Groups(3).Items(8).Caption = "单位值：" & Format(zlCommFun.NVL(rs("单位值")), "0.0")
            tkp.Groups(3).Items(9).Caption = "最高行：" & zlCommFun.NVL(rs("最高行"))
            tkp.Groups(3).Items(10).Caption = "记录频次：" & zlCommFun.NVL(rs("记录频次"))
            
        End If
    End Select
    
    cbsThis.RecalcLayout
    Call RefreshStateInfo
    
    zlMenuClick = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Public Function EditRefresh(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：数据新增/修改后数据重显处理
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long

    rptList.Records.DeleteAll
    
    Call zlMenuClick("读取数据")
    
    '恢复
    rptList.Populate
    
    For lngLoop = 0 To rptList.Rows.Count - 1
        If Not (rptList.Rows(lngLoop).Record Is Nothing) Then
            If Val(rptList.Rows(lngLoop).Record.Item(mCol.项目序号).Value) = lngKey Then
                Set rptList.FocusedRow = rptList.Rows(lngLoop)
                Call rptList_SelectionChanged
                Exit For
            End If
        End If
    Next

End Function

Private Function InitMenuBar() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：初始化菜单、工具栏
    '------------------------------------------------------------------------------------------------------------------
    Dim cbrMenuBar As Object
    Dim obj As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    
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
    cbsThis.ActiveMenuBar.Title = "菜单栏"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
        cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyTo, "适用科室(&T)"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "体温排列(&J)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Compend, "饮入代换(&Q)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "体温重叠(&K)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Request, "护理模板(&T)")
        
        '新版护士工作站新增功能
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CollectMan, "汇总项目(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AnimalPart, "体温部位(&P)"): cbrControl.IconId = 2612
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "记录频次(&L)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_WavyMan, "波动项目(&B)")
        '47964:刘鹏飞,2013-01-21,添加体温曲线同步设置功能
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_WaveSynchro, "体温同步(&S)")
    End With

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        
'        Set cbrControl = .Add(xtpControlButton, conMenu_View_Show, "显示停用(&A)"): cbrControl.BeginGroup = True: cbrControl.IconId = 1
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False  '固有
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)..."): cbrControl.BeginGroup = True
    End With
    
     '快键绑定
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    

    '工具栏定义
    Set cbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增")
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyTo, "适用科室"): cbrControl.BeginGroup = True
                
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.STYLE = xtpButtonIconAndCaption
    Next
    
    '读取发布到该模块的报表:因为是一次性读取,全局变量可用
    '---------------------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
End Function

'控件事件##############################################################################################################

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strKey As String
    Dim lngLoop As Long
    Dim strSQL() As String
    Dim blnTran As Boolean
    Dim lngIndex As Long
    Dim cbrControl As Object
    Dim lngKey As Long
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
        
    Select Case Control.ID
        Case conMenu_File_PrintSet
            
            Call zlPrintSet
                    
        Case conMenu_File_Preview
            
            Call zlRptPrint(2)
        
        Case conMenu_File_Print
        
            Call zlRptPrint(1)
        
        Case conMenu_File_Excel
        
            Call zlRptPrint(3)
    
        Case conMenu_View_ToolBar_Button
        
            cbsThis(2).Visible = Not cbsThis(2).Visible
            cbsThis.RecalcLayout
        
        Case conMenu_View_ToolBar_Text
        
            For Each cbrControl In cbsThis(2).Controls
                cbrControl.STYLE = IIf(cbrControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            
            cbsThis.RecalcLayout
            
        Case conMenu_View_StatusBar
        
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
            
        Case conMenu_Edit_NewItem
            '新增项目
            
            lngKey = 0
            
            
            If Not (rptList.FocusedRow Is Nothing) Then
                If Not (rptList.FocusedRow.Record Is Nothing) Then lngKey = Val(rptList.FocusedRow.Record.Item(mCol.项目序号).Value)
            End If
                
            If frmTendEdit.ShowEdit(Me, 0, lngKey) Then
                mblnOK = True
                rptList.SetFocus
            End If
    
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Modify
            
            If rptList.FocusedRow Is Nothing Then Exit Sub
            If rptList.FocusedRow.Record Is Nothing Then Exit Sub
            
            '修改项目
            If frmTendEdit.ShowEdit(Me, Val(rptList.FocusedRow.Record.Item(mCol.项目序号).Value)) Then
                mblnOK = True
                rptList.SetFocus
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Delete
            '删除项目
            If rptList.FocusedRow Is Nothing Then Exit Sub
            If rptList.FocusedRow.Record Is Nothing Then Exit Sub
            If CheckItemExistData(3, Val(rptList.FocusedRow.Record(mCol.项目序号).Value), rptList.FocusedRow.Record(mCol.项目名称).Value) = True Then Exit Sub
            If MsgBox("你真的要删除“" & rptList.FocusedRow.Record(mCol.项目名称).Value & "”？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            strSQL(ReDimArray(strSQL)) = "ZL_护理记录项目_DELETE(" & Val(rptList.FocusedRow.Record(mCol.项目序号).Value) & ")"
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_CollectMan    '汇总项目管理
            On Error Resume Next
            frmCollectMan.Show 1, Me
        
        Case conMenu_Edit_AnimalPart    '体温部位管理
            frmAnimalPartMan.Show 1, Me
            
        Case conMenu_Edit_Reuse         '记录频次管理
            frmItemRecordMan.Show 1, Me
        Case conMenu_Edit_WavyMan  '波动项目管理
            frmItemWaveMan.Show 1, Me
        '47964:刘鹏飞,2013-01-21,添加体温曲线同步设置功能
        Case conMenu_Edit_WaveSynchro '体温同步设置
            FrmTendWaveDataSet.Show 1, Me
        Case conMenu_Edit_ApplyTo
            If rptList.FocusedRow Is Nothing Then Exit Sub
            If rptList.FocusedRow.Record Is Nothing Then Exit Sub
            
            If frmTendItemDept.ShowMe(Me, Val(rptList.FocusedRow.Record.Item(mCol.项目序号).Value)) Then
                Call rptList_SelectionChanged
            End If
            
'        Case conMenu_Edit_Stop
'            '停用项目
'            If rptList.FocusedRow Is Nothing Then Exit Sub
'            If rptList.FocusedRow.Record Is Nothing Then Exit Sub
'
'            If MsgBox("你真的要停用“" & rptList.FocusedRow.Record(mCol.项目名称).Value & "”？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
'            strSQL(ReDimArray(strSQL)) = "ZL_护理记录项目_Stop(" & Val(rptList.FocusedRow.Record(mCol.项目序号).Value) & ")"
'
'        Case conMenu_Edit_Reuse
'            '启用项目
'            If rptList.FocusedRow Is Nothing Then Exit Sub
'            If rptList.FocusedRow.Record Is Nothing Then Exit Sub
'
'            If MsgBox("你真的要启用“" & rptList.FocusedRow.Record(mCol.项目名称).Value & "”？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
'            strSQL(ReDimArray(strSQL)) = "ZL_护理记录项目_Reuse(" & Val(rptList.FocusedRow.Record(mCol.项目序号).Value) & ")"
        
        Case conMenu_Edit_Adjust
            
            If frmTendBodyArrage.ShowEdit(Me) Then
                Call rptList_SelectionChanged
            End If
        
        Case conMenu_Edit_Compend
            
            '设置药品饮入代换关系
            If frmTendDrink.ShowEdit(Me) Then
                Call rptList_SelectionChanged
            End If
            
        Case conMenu_Edit_MarkMap

            Call frmTendBlanket.ShowEdit(Me, mstrPrivs)
        
        Case conMenu_Edit_Request   '护理项目模板
            Call frmTendItemTemplate.ShowMe(Me, mstrPrivs)
        
        Case conMenu_View_Refresh
                            
            '保存
            If Not (rptList.FocusedRow Is Nothing) Then
                If Not (rptList.FocusedRow.Record Is Nothing) Then strKey = Val(rptList.FocusedRow.Record(mCol.项目序号).Value)
            End If

            rptList.Records.DeleteAll
            
            Call zlMenuClick("读取数据")

            '恢复
            For lngLoop = 0 To rptList.Rows.Count - 1
                If Not (rptList.Rows(lngLoop).Record Is Nothing) Then
                    If Val(rptList.Rows(lngLoop).Record.Item(mCol.项目序号).Value) = Val(strKey) Then
                        Set rptList.FocusedRow = rptList.Rows(lngLoop)
                        Call rptList_SelectionChanged
                        Exit For
                    End If
                End If
            Next
            
        Case conMenu_Help_Help
        
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        
        Case conMenu_Help_About
            
            Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
            
        Case conMenu_Help_Web_Home
            
            Call zlHomePage(Me.hWnd)
            
        Case conMenu_Help_Web_Forum '中联论坛
            Call zlWebForum(Me.hWnd)

        Case conMenu_Help_Web_Mail
            
            Call zlMailTo(Me.hWnd)
            
        Case conMenu_File_Exit
            Unload Me
            Exit Sub
        Case Else
            '执行发布到当前模块的报表
'            Dim lng项目序号 As Long, str项目名称 As String
'            If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
'                If rptList.SelectedRows.Count > 0 Then
'                    If Not rptList.SelectedRows(0).GroupRow Then
'                        lng项目序号 = Val(rptList.SelectedRows(0).Record(mCol.项目序号).Value)
'                        str项目名称 = rptList.SelectedRows(0).Record(mCol.项目名称).Value
'                    End If
'                End If
'                If str项目名称 <> "" Then
'                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "项目序号=" & lng项目序号, "项目名称=" & str项目名称)
'                Else
'                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
'                End If
'            End If
            Exit Sub
    End Select
    
    blnTran = True
    gcnOracle.BeginTrans
    
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    
    gcnOracle.CommitTrans
    blnTran = False
    
    Select Case Control.ID
        Case conMenu_Edit_Delete
            '删除行
            
            lngIndex = rptList.FocusedRow.Index
            rptList.Records.RemoveAt (rptList.FocusedRow.Record.Index)
            rptList.Populate
            
            If rptList.Records.Count > 0 Then
                lngIndex = IIf(rptList.Records.Count - 1 > lngIndex, lngIndex, rptList.Records.Count - 1)
                rptList.Rows(lngIndex).Selected = True
                Set rptList.FocusedRow = rptList.Rows(lngIndex)
            End If
            rptList.SetFocus
            Call rptList_SelectionChanged
            mblnOK = True
            
'        Case conMenu_Edit_Stop
'            '填写撤档时间或删除此行
'
'            If mblnShowStop Then
'                rptList.FocusedRow.Record(mCol.撤档时间).Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
'            Else
'                lngIndex = rptList.FocusedRow.Index
'                rptList.Records.RemoveAt (rptList.FocusedRow.Record.Index)
'                rptList.Populate
'
'                If rptList.Records.Count > 0 Then
'                    lngIndex = IIf(rptList.Records.Count - 1 > lngIndex, lngIndex, rptList.Records.Count - 1)
'                    rptList.Rows(lngIndex).Selected = True
'                    Set rptList.FocusedRow = rptList.Rows(lngIndex)
'                End If
'                rptList.SetFocus
'                Call rptList_SelectionChanged
'
'            End If
'
'        Case conMenu_Edit_Reuse
'            '更改撤档时间为空
'            If rptList.FocusedRow Is Nothing Then Exit Sub
'            If rptList.FocusedRow.Record Is Nothing Then Exit Sub
'
'            rptList.FocusedRow.Record(mCol.撤档时间).Value = ""
    End Select
    
    cbsThis.RecalcLayout
    Call RefreshStateInfo
    
    Exit Sub
    
errHand:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)

    If stbThis.Visible Then Bottom = stbThis.Height
    
End Sub

Private Sub cbsThis_Resize()
    
    On Error Resume Next
    
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long  '客户区域的大小

    Call cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    
    With rptList
        .Left = lngLeft
        .Width = lngRight - lngLeft - tkp.Width - 45
        .Top = lngTop
        .Height = lngBottom - lngTop
    End With
    
    With tkp
        .Left = rptList.Left + rptList.Width + 45
        .Top = rptList.Top
        .Height = rptList.Height
    End With
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup
        End Select
    End If
    
    Err = 0: On Error Resume Next
    
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (rptList.Records.Count > 0)
    Case conMenu_Edit_NewItem, conMenu_Edit_Adjust
        Control.Enabled = (InStr(1, mstrPrivs, "增删改") > 0)
    Case conMenu_Edit_Request
        Control.Enabled = (InStr(1, mstrPrivs, "护理模板") > 0)
    Case conMenu_Edit_CollectMan
        Control.Enabled = (InStr(1, mstrPrivs, "汇总项目") > 0)
    Case conMenu_Edit_AnimalPart
        Control.Enabled = (InStr(1, mstrPrivs, "体温部位") > 0)
    Case conMenu_Edit_WavyMan
        Control.Enabled = (InStr(1, mstrPrivs, "护理波动项目") > 0)
    '47964:刘鹏飞,2013-01-21,添加体温曲线同步设置功能
    Case conMenu_Edit_WaveSynchro '体温同步设置
        Control.Enabled = (InStr(1, mstrPrivs, "体温同步项目") > 0)
    Case conMenu_Edit_Reuse
        Control.Enabled = (InStr(1, mstrPrivs, "增删改") > 0)
    Case conMenu_Edit_Modify
        If Not (rptList.FocusedRow Is Nothing) Then
            If Not (rptList.FocusedRow.Record Is Nothing) Then
                Control.Enabled = (InStr(1, mstrPrivs, "增删改") > 0)
            Else
                Control.Enabled = False
            End If
        Else
            Control.Enabled = False
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete
        If Not (rptList.FocusedRow Is Nothing) Then
            If Not (rptList.FocusedRow.Record Is Nothing) Then
                Control.Enabled = (InStr(1, mstrPrivs, "增删改") > 0)
            Else
                Control.Enabled = False
            End If
        Else
            Control.Enabled = False
        End If
        If Control.Enabled Then Control.Enabled = (rptList.FocusedRow.Record.Item(mCol.保留项目).Value <> "√")
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_ApplyTo

        Control.Visible = (InStr(1, mstrPrivs, "增删改") > 0)

        If Not (rptList.FocusedRow Is Nothing) Then
            If Not (rptList.FocusedRow.Record Is Nothing) Then
                Control.Enabled = (Control.Visible And rptList.FocusedRow.Record.Item(mCol.项目序号).Value > 2)
            Else
                Control.Enabled = False
            End If
        Else
            Control.Enabled = False
        End If
    Case conMenu_View_ToolBar_Button
        Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text
        Control.Checked = Not (Me.cbsThis(2).Controls(1).STYLE = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size
        Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar
        Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub Form_Load()
    mstrPrivs = gstrPrivs
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    
    mblnStartUp = True
'    mblnShowStop = False
    
    Call InitCommonControls
        
    Call InitMenuBar
    
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitGrid
    Call CreateToolBox
    
    mblnStartUp = False
    
    Call zlMenuClick("读取数据")
    
    On Error Resume Next
    
    If rptList.Records.Count > 0 Then Set rptList.FocusedRow = rptList.Rows(0)
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call SaveWinState(Me, App.ProductName)
    
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Not (rptList.FocusedRow Is Nothing) Then
            Call rptList_RowDblClick(rptList.FocusedRow, rptList.FocusedRow.Record.Item(mCol.项目名称))
        End If
    End If
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As Object
    
    If Button <> 2 Then Exit Sub
    
    If cbsThis.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = cbsThis.Add("弹出菜单", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
    Next
    cbrPopupBar.ShowPopup
End Sub

Private Sub rptList_RowDblClick(ByVal ROW As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    
    If Not (rptList.FocusedRow Is Nothing) Then
          Call cbsThis_Execute(cbsThis.FindControl(, conMenu_Edit_Modify))
    End If
End Sub

Private Sub rptList_SelectionChanged()

    If mblnStartUp Then Exit Sub
        
    Call zlMenuClick("读取要素")
    Call zlMenuClick("读取体温")
    Call zlMenuClick("读取适用科室")
        
End Sub


Public Function CheckItemExistData(ByVal bytType As Byte, ParamArray arrInput() As Variant) As Boolean
'功能:检查对应的护理项目是否在体温单或记录单已经存在数据
'bytType:1：只检查该项目是否已经产生了护理业务数据。2：只检查该项目是否已经绑定护理记录单:其它：存1否2
    Dim rsTemp As New ADODB.Recordset
    Dim strInfo As String
    Dim strSQL1 As String, strSQL2 As String
    On Error GoTo errHand
    CheckItemExistData = True
    strSQL1 = "Select Id" & vbNewLine & _
        " From (Select Id" & vbNewLine & _
        "       From 病人护理明细" & vbNewLine & _
        "       Where 项目序号 = [1] And Rownum < 2" & vbNewLine & _
        "       Union All" & vbNewLine & _
        "       Select Id" & vbNewLine & _
        "       From 病人护理内容" & vbNewLine & _
        "       Where 项目序号 = [1] And Rownum < 2)"
    strSQL2 = " Select a.名称 " & vbNewLine & _
        " From 病历文件结构 d, 病历文件结构 p, 病历文件列表 a" & vbNewLine & _
        " Where p.Id = d.父id And p.对象类型 = 1 And p.内容文本 = '表列集合' And d.要素名称 = [1] And p.文件id = a.Id And a.种类 = 3 And 保留 <> -1 And Rownum <2"
    
    If bytType = 1 Then
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL1, "检查项目是否已经存在护理数据", Val(arrInput(0)))
        If rsTemp.RecordCount > 0 Then
            MsgBox "该项目已经产生了护理数据，不允许删除或修改！", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf bytType = 2 Then
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL2, "检查项目是否已绑定护理记录单", CStr(arrInput(0)))
        If rsTemp.RecordCount > 0 Then
            MsgBox "该项目已与护理文件绑定，不允许进行删除或修改名称！", vbInformation, gstrSysName
            Exit Function
        End If
    Else
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL1, "检查项目是否已经存在护理数据", Val(arrInput(0)))
        If rsTemp.RecordCount > 0 Then
            MsgBox "该项目已经产生了护理数据，不允许删除！", vbInformation, gstrSysName
            Exit Function
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL2, "检查项目是否已绑定护理记录单", CStr(arrInput(1)))
        If rsTemp.RecordCount > 0 Then
            MsgBox "该项目已与护理文件绑定，不允许进行删除或修改名称！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CheckItemExistData = False
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


