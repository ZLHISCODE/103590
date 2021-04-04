VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmModelExportOrImport 
   Caption         =   "范文批量导出列表"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9990
   Icon            =   "frmModelExportOrImport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   9990
   StartUpPosition =   1  '所有者中心
   Begin zlRichEditor.Editor Editor1 
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
   End
   Begin VB.ComboBox cboList 
      Height          =   300
      Left            =   4920
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox PicBtn 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   9990
      TabIndex        =   0
      Top             =   6360
      Width           =   9990
      Begin MSComctlLib.ProgressBar progBar 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   3720
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex8Ctl.VSFlexGrid vsgrid 
      Height          =   1860
      Left            =   720
      TabIndex        =   2
      Top             =   2040
      Width           =   2655
      _cx             =   4683
      _cy             =   3281
      Appearance      =   2
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
      BackColor       =   -2147483639
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16761024
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483639
      GridColor       =   12632256
      GridColorFixed  =   12632256
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
   Begin MSComctlLib.ImageList img16 
      Left            =   4440
      Top             =   360
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
            Picture         =   "frmModelExportOrImport.frx":6852
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmModelExportOrImport.frx":6DEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmModelExportOrImport.frx":7386
            Key             =   "签名"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox RTbFootText 
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmModelExportOrImport.frx":76D8
   End
   Begin RichTextLib.RichTextBox RTbHeadText 
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmModelExportOrImport.frx":7775
   End
   Begin RichTextLib.RichTextBox RTbContext 
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmModelExportOrImport.frx":7812
   End
   Begin XtremeCommandBars.ImageManager imgManager 
      Left            =   5280
      Top             =   360
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmModelExportOrImport.frx":78AF
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
End
Attribute VB_Name = "frmModelExportOrImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum menu_this
    menu_Cover = 1                  '导入覆盖文件
    menu_Add = 2                    '导入新增文件
    menu_RemoveRow = 3              '移除选中行
    menu_Clear = 4                  '清空列表
    menu_Export = 10                '导出
    menu_Import = 11                '导入
    menu_IcheckAll = 12             '导入全选
    menu_IclearAll = 13             '导入全清
    menu_CheckThis = 14             '选择当前行
    menu_AddFile = 15               '导入新增文件
    menu_Unload = 16                '退出
    menu_EcheckAll = 17             '导出全选
    menu_EclearAll = 18             '导出全清
    menu_ExportOne = 101            '导出一个XML
    menu_ExportMore = 102           '导出多个XML
    menu_CheckHave = 121            '全选已存在文件
    menu_ClearHave = 131            '全清已存在文件
    menu_ImportOption = 110         '导入设置
End Enum
Private Enum mCols
    col_种类 = 0
    col_文件ID = 1
    col_文件名称 = 2
    col_ID = 3
    col_范文ID = 4
    col_编号 = 5
    col_范文名称 = 6
    col_简码 = 7
    col_分类 = 8
    col_性质 = 9
    col_说明 = 11
    col_通用级 = 10
    col_部门 = 12
    col_人员 = 13
End Enum
'################################################################################################################
'## 功能：  显示导出/导入窗体
'## 参数：  lngType     :显示内容（0-导入列表  1-导出列表）
'##         objParent   :父窗体
'################################################################################################################
Public Sub ShowMe(ByVal objParent As Object, ByVal lngType As Long)
    If lngType = 1 Then
        Me.Caption = "范文批量导出列表"
        Me.vsgrid.Tag = "Export"
        Me.PicBtn.Visible = True
        If Not ExportList Then InitVsGrid ("暂时没有可以导出的数据！")
        Me.Show 1, objParent
    Else
        Me.Caption = "范文批量导入列表"
        Me.vsgrid.Tag = "Import"
        Me.PicBtn.Visible = True
        If Not ImportList Then Exit Sub
        Me.Show 1, objParent
    End If
End Sub

Private Sub cboList_Click()
    Dim i As Integer, strTempName As String
    With vsgrid
        If .Row < 1 Then Exit Sub
        strTempName = .TextMatrix(.Row + 1, 2)
        .Cell(flexcpData, .Row, 3, .GetNodeRow(.Row, flexNTLastChild), 3) = Me.cboList.ItemData(cboList.ListIndex)
        .Cell(flexcpText, .Row, 2, .GetNodeRow(.Row, flexNTLastChild), 2) = Me.cboList.List(cboList.ListIndex)
        .Cell(flexcpText, .Row, 2, .Row, 5) = Me.cboList.List(cboList.ListIndex) & "(原所属文件：" & .Cell(flexcpData, .Row, 4) & ")"
        .Cell(flexcpForeColor, .Row, 2, .Row, 5) = vbBlue
    End With
End Sub

Private Sub cboList_GotFocus()
    cboList_Click
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
           Case menu_RemoveRow
                If vsgrid.RowOutlineLevel(vsgrid.Row) = 0 Then MsgBox "该文件节点行不能移除！", vbInformation, gstrSysName:   Exit Sub
                vsgrid.RemoveItem (vsgrid.Row)
                If vsgrid.Rows = 1 Then InitVsGrid ("请先添加需要导入的范文文件 ！")
                Me.Tag = Val(Me.Tag) - 1
           Case menu_Clear
                Call InitVsGrid("请先添加需要导入的范文文件 ！")
           Case menu_Export
                StartExportToXMLs
           Case menu_Import
                StartImportFromXML
           Case menu_AddFile
                ImportList
           Case menu_EcheckAll
                CheckAllOrClearAll True
           Case menu_EclearAll
                CheckAllOrClearAll False
           Case menu_Unload
                Unload Me
    End Select
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
           Case menu_Import
                Control.Visible = IIf(vsgrid.Tag = "Import", True, False)
                Control.Enabled = IIf(progBar.Visible Or vsgrid.Cols = 1, False, True)
           Case menu_Export
                Control.Visible = IIf(vsgrid.Tag = "Export", True, False)
                Control.Enabled = IIf(progBar.Visible Or vsgrid.Cols = 1, False, True)
           Case menu_AddFile
                Control.Visible = IIf(vsgrid.Cols = 1 And vsgrid.Tag = "Import", True, False)
           Case menu_EcheckAll, menu_EclearAll
                Control.Enabled = IIf(vsgrid.Cols = 1 Or progBar.Visible, False, True)
           Case menu_Unload
                Control.Enabled = IIf(progBar.Visible, False, True)
    End Select
End Sub

Private Sub Form_Load()
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Set objBar = Me.cbsThis.Add("Tools", xtpBarTop)
    objBar.ContextMenuPresent = False           '工具栏上点击鼠标右键时不弹出设置菜单
    objBar.ShowTextBelowIcons = False           '工具栏中的按钮文字显示在图标右侧
    objBar.EnableDocking xtpFlagStretched
    Me.cbsThis.Icons = Me.imgManager.Icons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True                 '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, menu_Export, "导出"): objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, menu_Import, "导入"): objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, menu_AddFile, "添加"): objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, menu_EcheckAll, "全选"): objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, menu_EclearAll, "全清"): objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, menu_Unload, "退出"): objControl.STYLE = xtpButtonIconAndCaption
        objControl.STYLE = xtpButtonIconAndCaption  '同时显示图标和文字
    End With
    Me.cbsThis.ActiveMenuBar.Visible = False
End Sub

Private Sub Form_Resize()
    Me.vsgrid.Move 0, 500, Me.ScaleWidth, Me.ScaleHeight - Me.PicBtn.Height - 500
    Me.PicBtn.Move 0, vsgrid.Height + 500, Me.ScaleWidth, PicBtn.Height
    Me.progBar.Move 0, 60, PicBtn.Width, progBar.Height
    If vsgrid.Rows = 1 Or vsgrid.Cols = 1 Then
        vsgrid.ROWHEIGHT(0) = Me.ScaleHeight
    End If
End Sub
'初始化Vsgrid
Private Function InitVsGrid(ByVal strMsg As String)
    vsgrid.Tag = ""
    With vsgrid
        .Clear: .Cols = 1: .Rows = 1: .FixedRows = 1
        .ROWHEIGHT(0) = vsgrid.Height
        .TextMatrix(0, 0) = strMsg
        .Cell(flexcpFontSize, 0, 0) = 20
        .Cell(flexcpAlignment, 0, 0) = flexAlignCenterCenter
        .Cell(flexcpBackColor, 0, 0) = vbWhite
    End With
    vsgrid.ROWHEIGHT(0) = Me.ScaleHeight
End Function
'导出列表加载
Private Function ExportList() As Boolean
    Dim strType As String, strFileName As String, i As Long, j As Long, k As Long, lngRow As Long
    Dim rsFiles As ADODB.Recordset              '文件数据集
    Dim rsModels As New ADODB.Recordset         '范文数据集
    On Error GoTo errHand
    gstrSQL = "(Select decode(f.种类,1,'1-门诊病历',2,'2-住院病历',4,'4-护理病历',5,'5-疾病证明报告',6,'6-知情文件',7,'诊疗单据') as 种类," & _
              "  f.id as 文件ID,f.名称 as 文件名称 From 病历范文目录 l, 部门表 d," & _
              "  人员表 p,病历文件列表 f Where l.科室id = d.Id and l.文件id=f.id  And l.人员id = p.Id and decode(l.性质,null,0,0,0)=0 group by f.种类,f.id, f.名称" & _
              "  Union All" & _
              "  Select decode(f.种类,1,'1-门诊病历',2,'2-住院病历',4,'4-护理病历',5,'5-疾病证明报告',6,'6-知情文件',7,'诊疗单据') as 种类," & _
              "  0,null  From 病历范文目录 l, 部门表 d," & _
              "  人员表 p,病历文件列表 f Where l.科室id = d.Id and l.文件id=f.id  And l.人员id = p.Id and decode(l.性质,null,0,0,0)=0 group by 种类) order by 种类,文件ID"
        Set rsFiles = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
         '基础设置
        If rsFiles.RecordCount < 1 Then
             Call InitVsGrid("暂时没有可以导出的病历文件 ！"):  rsFiles.Close
             Exit Function
        End If
        With vsgrid
            .Clear: .FixedRows = 1: .Cols = 14: .ROWHEIGHT(0) = 10
            '设置分组
            .OutlineCol = 0: .OutlineBar = flexOutlineBarCompleteLeaf
            .BackColorSel = vbWhite
            .TextMatrix(0, col_编号) = "编号": .TextMatrix(0, col_范文名称) = "名称": .TextMatrix(0, col_分类) = "分类": .TextMatrix(0, col_通用级) = "通用级": .TextMatrix(0, col_说明) = "说明":
            For i = 1 To rsFiles.RecordCount
                '加载种类节点
                If strType <> rsFiles!种类 Then
                    .AddItem ""
                    Me.Tag = Val(Me.Tag) + 1
                    lngRow = Val(Me.Tag)
                    For k = 2 To .Cols - 1
                        .TextMatrix(lngRow, k) = NVL(rsFiles("种类").Value)
                        .ColAlignment(k) = flexAlignLeftCenter
                        .ColWidth(k) = 300
                    Next k
                    .Cell(flexcpBackColor, lngRow, 0, lngRow, 13) = &HFFC0C0
                    .IsSubtotal(lngRow) = True
                    .Cell(flexcpData, lngRow, 1) = 1
                    .MergeCells = flexMergeFree
                    .MergeRow(lngRow) = True '是否左右行合并
                     strType = rsFiles!种类
                Else
                    '加载文件节点
                    .AddItem ""
                    Me.Tag = Val(Me.Tag) + 1
                    lngRow = Val(Me.Tag)
                    For k = 3 To .Cols - 1
                        .TextMatrix(lngRow, k) = NVL(rsFiles("文件名称").Value)
                    Next k
                    .Cell(flexcpData, lngRow, 2) = NVL(rsFiles("文件ID").Value)
                    .Cell(flexcpBackColor, lngRow, 2, lngRow, 12) = &H80000016
                    .IsSubtotal(lngRow) = True
                    .RowOutlineLevel(lngRow) = 1
                    .MergeCells = flexMergeFree
                    .MergeRow(lngRow) = True '是否左右行合并
                    
                    gstrSQL = "Select l.Id,l.文件id,l.编号, l.名称,l.简码,Nvl(l.分类, '未分类') As 分类,l.性质,l.说明, l.通用级,d.名称 As 部门," & _
                              "p.姓名 As 人员,Decode(l.分类, Null, 1, 2) As 排序 From 病历范文目录 l, 部门表 d, 人员表 p Where l.科室id = d.Id " & _
                              "And l.人员id = p.Id and l.文件id=" & rsFiles("文件ID").Value & " Order By Decode(l.分类, Null, 1, 2), l.分类, l.编号"
                    Set rsModels = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                    rsModels.MoveFirst
                    '加载范文节点
                    Do While Not rsModels.EOF
                    '0-全院通用;1-部门通用;2-个人使用
                        .AddItem ""
                        Me.Tag = Val(Me.Tag) + 1
                        lngRow = Val(Me.Tag)
                        .Cell(flexcpData, lngRow, col_编号) = NVL(rsModels!ID)
                        .TextMatrix(lngRow, col_编号) = NVL(rsModels!编号)
                        .TextMatrix(lngRow, col_范文名称) = NVL(rsModels!名称)
                        .TextMatrix(lngRow, col_简码) = NVL(rsModels!简码)
                        .TextMatrix(lngRow, col_分类) = NVL(rsModels!分类)
                        .TextMatrix(lngRow, col_性质) = NVL(rsModels!性质)
                        .TextMatrix(lngRow, col_说明) = NVL(rsModels!说明)
                        .Cell(flexcpData, lngRow, col_通用级) = NVL(rsModels!通用级)
                        .Cell(flexcpBackColor, lngRow, 5, lngRow, 13) = &HE0E0E0
                        Select Case Val(rsModels!通用级)
                               Case 0
                                .TextMatrix(lngRow, col_通用级) = "全院通用"
                               Case 1
                               .TextMatrix(lngRow, col_通用级) = "部门通用"
                               Case 2
                               .TextMatrix(lngRow, col_通用级) = "个人使用"
                        End Select
                        .TextMatrix(lngRow, col_人员) = NVL(rsModels!人员)
                        .RowOutlineLevel(lngRow) = 2
                        rsModels.MoveNext
                    Loop
                End If
                rsFiles.MoveNext
           Next i
            .ColWidth(0) = 400: .ColWidth(1) = 270: .ColWidth(2) = 0: .ColWidth(4) = 270
            .ColWidth(col_范文名称) = 1500: .ColWidth(col_分类) = 1000: .ColWidth(col_说明) = 1000: .ColWidth(col_通用级) = 1000: .ColWidth(col_编号) = 700
            .ColWidth(col_性质) = 0: .ColWidth(col_简码) = 0: .ColWidth(col_部门) = 0: .ColWidth(col_人员) = 0
           '清除列表空余部分的边框线
           For i = 1 To vsgrid.Rows - 1
                If .IsSubtotal(i) = True Then
                    .GetNode(i).Expanded = True
                End If
                If .RowOutlineLevel(i) = 2 Then
                    .Cell(flexcpPicture, i, 4) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, i, 4) = 1
                    .CellBorderRange i, 0, i, 3, vbWhite, 1, 0, 0, 1, 1, 1
                End If
                If .RowOutlineLevel(i) = 1 Then
                    .Cell(flexcpPicture, i, 1) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, i, 1) = 1
                    .CellBorderRange i, 0, i, 13, &H80000016, 1, 1, 1, 1, 1, 1
                End If
                If .RowOutlineLevel(i) = 0 Then
                    .Cell(flexcpPicture, i, 1) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, i, 1) = 1
                    .CellBorderRange i, 0, i, 13, &H80000016, 1, 1, 1, 1, 1, 1
                End If
           Next i
           '----------------------------------------
           vsgrid.RemoveItem (lngRow + 1)
           If vsgrid.Rows > 1 Then vsgrid.Row = 2
    End With
    rsFiles.Close
    ExportList = True
    Exit Function
errHand:
    ExportList = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function ImportList() As Boolean
    Dim i As Integer, j As Integer, k As Integer, l As Integer, lngDemoId As Long, lngRow As Long
    Dim strXMLPath As String                'XML路径
    Dim strTempPath As String               '临时路径
    Dim strItems As String                  '基础信息
    Dim strFileName As String               '病历文件名称
    Dim strItemsArr As Variant              '范文基础信息数组
    Dim strArrXml As Variant                '文件地址数组
    Dim oDoc As New DOMDocument             'Xml文档
    Dim cDoc As New cEPRDocument            '文档对象
    Dim oRoot As IXMLDOMElement             '根节点
    Dim oFileList As IXMLDOMNodeList        '文件节点集合
    Dim oDemoList As IXMLDOMNodeList        '范文节点集合
    Dim oSubNode As IXMLDOMElement          '子节点
    Dim oSubNode1 As IXMLDOMElement         '子节点
    Dim rsTemp As New ADODB.Recordset
    
    On Error Resume Next
    dlgThis.MaxFileSize = 32767
    dlgThis.Filter = "*.ZIP|*.zip"
    dlgThis.DialogTitle = "打开"
    dlgThis.CancelError = True
    dlgThis.flags = &H10& Or &H80000
    dlgThis.ShowOpen
    If Err.Number = 32755 Then Err.Clear: ImportList = False: Exit Function
    On Error GoTo errHand
    With vsgrid
            '重新设置VsGrid
            If Val(Me.Tag) < 1 Then
                .Clear
                .FixedRows = 1: .ExplorerBar = flexExSortShow
                .OutlineCol = 0: .OutlineBar = flexOutlineBarComplete
                .Cols = 6: .ColWidth(0) = 200: .Rows = 1: .ColAlignment(1) = flexAlignLeftCenter: .ROWHEIGHT(0) = Me.cboList.Height: .ColAlignment(0) = flexAlignRightCenter
                .ColWidth(1) = 270: .ColWidth(2) = 1500: .ColWidth(3) = 2500: .ColWidth(4) = 2500: .ColWidth(5) = 6000
                .TextMatrix(0, 1) = "选择": .TextMatrix(0, 2) = "所属文件": .TextMatrix(0, 3) = "范文名称": .TextMatrix(0, 4) = "导出单位": .TextMatrix(0, 5) = "文件位置"
            End If
            '加载文件选择下拉框列表数据
            gstrSQL = "select ID,名称 from 病历文件列表 "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            If rsTemp.RecordCount = 0 Then MsgBox "当前系统不存在病历文件，请先添加病历文件！", vbInformation, gstrSysName: Exit Function
            rsTemp.MoveFirst
            Do While Not rsTemp.EOF
                Me.cboList.AddItem (NVL(rsTemp!名称, ""))
                Me.cboList.ItemData(i) = NVL(rsTemp!ID, 0)
                 i = i + 1
                Me.cboList.ListIndex = 0
                rsTemp.MoveNext
            Loop
            '开始循环加载导入范文列表
            strXMLPath = dlgThis.Filename
            Me.cboList.Tag = zlFilesUnZip(strXMLPath)         '保存加载文件路径
            If Me.cboList.Tag = "" Then MsgBox "你选择的文件数据格式不正确或已被损坏!", vbInformation, gstrSysName: Exit Function
            oDoc.Load Me.cboList.Tag
            '删除临时文件
            gobjFSO.DeleteFile (Me.cboList.Tag)
            '如果该路径下文件已被加载则不再加载
            For l = 1 To vsgrid.Rows - 1
                If strXMLPath = Trim(vsgrid.TextMatrix(l, 5)) Then
                     MsgBox strArrXml(i) & ",已经被打开，请勿重复打开 ！", vbInformation, gstrSysName: Exit Function
                End If
            Next l
            '读取XML文件根节点
            Set oRoot = oDoc.selectSingleNode("EPRDemosList")
            If oRoot Is Nothing Then MsgBox "你选择的文件数据格式不正确或已被损坏!", vbInformation, gstrSysName: Exit Function
            '读取XML文件中病历文件集合
            Set oFileList = oRoot.selectNodes("/EPRDemosList/Kind/File")
            If oFileList.Item(0) Is Nothing Then MsgBox "你选择的文件数据格式不正确或已被损坏!", vbInformation, gstrSysName: Exit Function
            '开始循环遍历文件集合
            For Each oSubNode In oFileList
                '绑定文件节点
                Set oDemoList = oSubNode.selectNodes("Demo")
                Me.Tag = Val(Me.Tag) + 1   '将VsGrid的Rows保存
                lngRow = Val(Me.Tag)       '取出VsGrid的Rows
                .AddItem ""
                strFileName = NVL(oSubNode.getAttribute("FileName"))
                gstrSQL = "select ID from 病历文件列表 where 名称='" & strFileName & "'"
                '判断文件是否存在
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                    .Cell(flexcpForeColor, lngRow, 2, lngRow, 5) = IIf(rsTemp.RecordCount > 0, vbBlue, vbMagenta)
                If rsTemp.RecordCount > 0 Then
                   .Cell(flexcpData, lngRow, 3) = Val(rsTemp!ID)                        '绑定文件ID
                   .Cell(flexcpData, lngRow, 1) = 1                                     '设置选中列的值
                   .Cell(flexcpPicture, lngRow, 1) = img16.ListImages("Check").Picture  '设置选中图片
                End If
                .Cell(flexcpData, lngRow, 2) = rsTemp.RecordCount  '此列值作为记号
                .Cell(flexcpData, lngRow, 4) = strFileName
                For k = 2 To 5
                    .TextMatrix(lngRow, k) = IIf(rsTemp.RecordCount > 0, strFileName, strFileName & "(该病历文件在当前数据库不存在，请单击此处选择病例文件...)")
                    .ColAlignment(k) = flexAlignLeftCenter
                Next k
                .IsSubtotal(lngRow) = True
                .ROWHEIGHT(lngRow) = Me.cboList.Height
                .MergeCells = flexMergeFree
                .MergeRow(lngRow) = True '是否左右行合并
                '绑定范文节点
                For Each oSubNode1 In oDemoList
                     Me.Tag = Val(Me.Tag) + 1
                     lngRow = Val(Me.Tag)
                    .AddItem ""
                    If rsTemp.RecordCount > 0 Then
                        .Cell(flexcpData, lngRow, 1) = 1
                        .Cell(flexcpPicture, lngRow, 1) = img16.ListImages("Check").Picture
                    End If
                     strItems = oSubNode1.getAttribute("Items")
                     lngDemoId = Val(oSubNode1.getAttribute("ID"))
                     strItemsArr = Split(strItems, "|")
                    .ROWHEIGHT(lngRow) = Me.cboList.Height
                    .TextMatrix(lngRow, 2) = strFileName                    '绑定文件名称
                    .TextMatrix(lngRow, 3) = strItemsArr(0)                 '绑定范文名称
                    .TextMatrix(lngRow, 4) = oRoot.getAttribute("UnitName") '绑定导出单位
                    .TextMatrix(lngRow, 5) = strXMLPath                                                    '绑定文件路径
                    .Cell(flexcpData, lngRow, 3) = .Cell(flexcpData, .GetNodeRow(lngRow, flexNTParent), 3) '设置选中列的值
                    .Cell(flexcpForeColor, lngRow, 3) = IIf(rsTemp.RecordCount > 0, vbBlue, vbMagenta)     '行颜色设置
                    .Cell(flexcpData, lngRow, 4) = lngDemoId                '绑定范文ID
                    .Cell(flexcpData, lngRow, 5) = strItems                 '绑定范文基础数据
                    .RowOutlineLevel(lngRow) = 1
                Next
            Next
    End With
    ImportList = True: Exit Function
errHand:
    ImportList = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub Form_Unload(Cancel As Integer)
    If Me.cboList.Tag <> "" Then
        gobjFSO.DeleteFile (gobjFSO.GetParentFolderName(Replace(Me.cboList.Tag, "_范文列表.xml", "_范文信息.xml")))
    End If
End Sub

'单击事件
Private Sub vsgrid_Click()
    If Not vsgrid.MouseIcon Is Nothing And vsgrid.MouseRow > 0 Then
         CheckItems vsgrid.Row
    End If
    If vsgrid.Tag = "Import" Then
        If vsgrid.IsSubtotal(vsgrid.Row) And (vsgrid.MouseCol = 2 Or vsgrid.MouseCol = 3 Or vsgrid.MouseCol = 4) And vsgrid.Cell(flexcpData, vsgrid.Row, 2) <> 1 Then
            Me.cboList.Visible = True
            Me.cboList.Move vsgrid.Cell(flexcpLeft, vsgrid.Row, 2), vsgrid.Cell(flexcpTop, vsgrid.Row, 1) + vsgrid.ROWHEIGHT(vsgrid.Row) * 2 - 100, vsgrid.ColWidth(1) + vsgrid.ColWidth(3)
        Else
            Me.cboList.Visible = False
        End If
    End If
End Sub
'###############################################################
'# 方法： 选中Vsgrid的某行
'# 参数： lngRow :行号
'###############################################################
Private Sub CheckItems(ByVal lngRow As Long)
    Dim i As Long
    With vsgrid
        If .Tag = "Export" Then
            Select Case .RowOutlineLevel(lngRow)
                   Case 0  '一级
                     .Cell(flexcpData, lngRow, 1) = IIf(.Cell(flexcpData, lngRow, 1) = 1, 0, 1)
                     .Cell(flexcpPicture, lngRow, 1) = IIf(.Cell(flexcpData, lngRow, 1) = 1, img16.ListImages("Check").Picture, Nothing)
                     For i = lngRow To .GetNodeRow(lngRow, flexNTLastChild)
                        If .RowOutlineLevel(i) = 2 Then
                            .Cell(flexcpData, i, 4) = .Cell(flexcpData, lngRow, 1)
                            .Cell(flexcpPicture, i, 4) = IIf(.Cell(flexcpData, i, 4) = 1, img16.ListImages("Check").Picture, Nothing)
                        ElseIf .RowOutlineLevel(i) = 1 Then
                            .Cell(flexcpData, i, 1) = .Cell(flexcpData, lngRow, 1)
                            .Cell(flexcpPicture, i, 1) = IIf(.Cell(flexcpData, lngRow, 1) = 1, img16.ListImages("Check").Picture, Nothing)
                        End If
                     Next i
                   Case 1 '二级
                     .Cell(flexcpData, lngRow, 1) = IIf(.Cell(flexcpData, lngRow, 1) = 1, 0, 1)
                     .Cell(flexcpPicture, lngRow, 1) = IIf(.Cell(flexcpData, lngRow, 1) = 1, img16.ListImages("Check").Picture, Nothing)
                     For i = lngRow To .GetNodeRow(lngRow, flexNTLastChild)
                        If .RowOutlineLevel(i) = 2 Then
                            .Cell(flexcpData, i, 4) = .Cell(flexcpData, lngRow, 1)
                            .Cell(flexcpPicture, i, 4) = IIf(.Cell(flexcpData, i, 4) = 1, img16.ListImages("Check").Picture, Nothing)
                        End If
                     Next i
                   Case 2  '三级
                        .Cell(flexcpData, lngRow, 4) = IIf(.Cell(flexcpData, lngRow, 4) = 1, 0, 1)
                        .Cell(flexcpPicture, lngRow, 4) = IIf(.Cell(flexcpData, lngRow, 4) = 1, img16.ListImages("Check").Picture, Nothing)
            End Select
        Else
            If Val(vsgrid.Cell(flexcpData, lngRow, 3)) = 0 Then
                MsgBox "所属病例文件不存在，请先选择病例文件！", vbInformation, gstrSysName: Exit Sub
            End If
            If vsgrid.IsSubtotal(lngRow) Then
              For i = lngRow To .GetNodeRow(lngRow, flexNTLastChild)
                .Cell(flexcpData, i, 1) = IIf(.Cell(flexcpData, lngRow, 1) = 1, 0, 1)
                .Cell(flexcpPicture, i, 1) = IIf(.Cell(flexcpData, lngRow, 1) = 1, img16.ListImages("Check").Picture, Nothing)
                .Cell(flexcpData, i, 1) = IIf(.Cell(flexcpData, lngRow, 1) = 1, 1, 0)
              Next i
            Else
            .Cell(flexcpData, lngRow, 1) = IIf(.Cell(flexcpData, lngRow, 1) = 1, 0, 1)
            .Cell(flexcpPicture, lngRow, 1) = IIf(.Cell(flexcpData, lngRow, 1) = 1, img16.ListImages("Check").Picture, Nothing)
            End If
        End If
    End With
End Sub
'全选/全清
Private Sub CheckAllOrClearAll(ByVal blnCheck As Boolean)
    Dim i As Long
    If vsgrid.Tag = "Export" Then
        For i = 1 To vsgrid.Rows - 1
            If vsgrid.RowOutlineLevel(i) = 0 Then
                vsgrid.Cell(flexcpData, i, 1) = IIf(blnCheck, 0, 1)
                CheckItems i
            End If
        Next i
    Else
       For i = 1 To vsgrid.Rows - 1
            If vsgrid.RowOutlineLevel(i) = 0 And vsgrid.Cell(flexcpData, i, 3) <> Empty Then
                vsgrid.Cell(flexcpData, i, 1) = IIf(blnCheck, 0, 1)
                vsgrid.Cell(flexcpPicture, i, 1) = IIf(blnCheck, img16.ListImages("Check").Picture, Nothing)
                CheckItems i
           End If
       Next i
    End If
End Sub
'获取范文的行数据集字符串
Private Function GetRowsData(ByVal lngRow As Long) As String
    Dim strRowData As String, intModelID As Long
    Dim rsTemp As New ADODB.Recordset                                                                                            '行字符串内容及索引：
    With vsgrid
            strRowData = strRowData & .Cell(flexcpText, .GetNodeRow(.GetNodeRow(lngRow, flexNTParent), flexNTParent), 2) & "|"   '----种类名称  0
            strRowData = strRowData & .Cell(flexcpData, .GetNodeRow(lngRow, flexNTParent), 2) & "|"                              '----文件ID    1
            strRowData = strRowData & .TextMatrix(.GetNodeRow(lngRow, flexNTParent), 3) & "|"                                    '----文件名称  2
            strRowData = strRowData & .Cell(flexcpData, lngRow, col_编号) & "|"                                                  '---- 范文ID   3
            strRowData = strRowData & .TextMatrix(lngRow, col_编号) & "|"                                                        '---- 编号     4
            strRowData = strRowData & .TextMatrix(lngRow, col_范文名称) & "|"                                                    '---- 范文名称 5
            strRowData = strRowData & .TextMatrix(lngRow, col_简码) & "|"                                                        '---- 简码     6
            strRowData = strRowData & .TextMatrix(lngRow, col_性质) & "|"                                                        '---- 性质     7
            strRowData = strRowData & .TextMatrix(lngRow, col_说明) & "|"                                                        '---- 说明     8
            strRowData = strRowData & .Cell(flexcpData, lngRow, col_通用级) & "|"                                                '---- 通用级   9
            strRowData = strRowData & .TextMatrix(lngRow, col_分类) & "|"                                                        '---- 分类     10
            strRowData = strRowData & glngDeptId & "|"                                                                           '---- 部门ID   11
            strRowData = strRowData & glngUserId & "|"                                                                           '---- 操作员ID 12
            intModelID = Val(.Cell(flexcpData, lngRow, col_编号))
            gstrSQL = "Select 名称 As 条件项, 简码 As 条件值 From Table(Cast(f_Segment_条件项('" & intModelID & "') As ZLHIS.t_Dic_Rowset)) where 简码 is not null"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            '如果无条件去掉末尾的"|"
            If rsTemp.EOF Then strRowData = Mid(strRowData, 1, Len(strRowData) - 1)
            '添加范文条件值到字符串中
            Do While Not rsTemp.EOF
                strRowData = strRowData & rsTemp!条件项 & ":" & rsTemp!条件值 & ";"
                rsTemp.MoveNext
            Loop
            strRowData = Mid(strRowData, 1, Len(strRowData) - 1)
    End With
    GetRowsData = strRowData
End Function
Private Sub vsgrid_DblClick()
    If vsgrid.MouseRow < 1 Then Exit Sub
    CheckItems (vsgrid.Row)
End Sub

Private Sub vsgrid_KeyDown(KeyCode As Integer, Shift As Integer)
     With vsgrid
        If .IsSubtotal(.Row) Then
            Select Case KeyCode
              Case vbKeyLeft    '←键收缩
                  .GetNode(.Row).Expanded = False
              Case vbKeySpace   '空格选中
                   CheckItems .Row
              Case vbKeyRight   '→键展开
                  .GetNode(.Row).Expanded = True
              Case vbKeyReturn  '回车展开/收缩
                .GetNode(.Row).Expanded = Not .GetNode(.Row).Expanded
              Case vbKeyA       'CTRL+A 全选
                If Shift = 2 Then CheckAllOrClearAll (True)
              Case vbKeyZ       'CTRL+Z 全清
                If Shift = 2 Then CheckAllOrClearAll (False)
            End Select
        End If
   End With
End Sub

Private Sub vsgrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If vsgrid.Cols = 1 Or vsgrid.Rows = 1 Then Exit Sub
     If Button = 2 Then
        If vsgrid.Tag = "Import" And vsgrid.MouseCol > 0 And vsgrid.MouseRow > 0 Then
                Dim Popup As CommandBar
                Dim objControl As CommandBarControl
                Set Popup = cbsThis.Add("Popup", xtpBarPopup)
                With Popup.Controls
                    .Add xtpControlButton, menu_RemoveRow, "从列表中移除(&D)"
                    .Add xtpControlButton, menu_Clear, "清空列表(&C)"
                End With
                Popup.ShowPopup
        End If
      End If
      
End Sub

Private Sub vsgrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intX As Integer, intW As Integer, intX2 As Integer, intW2 As Integer
    If vsgrid.Cols = 1 Then Exit Sub
    If vsgrid.MouseRow = -1 Then
        vsgrid.MousePointer = flexDefault
        Set vsgrid.MouseIcon = Nothing: Exit Sub
    End If
    If vsgrid.Tag = "Import" Then
         If vsgrid.MouseCol = 1 And vsgrid.MouseRow <> 0 Then
                 vsgrid.MousePointer = flexCustom
                 Set vsgrid.MouseIcon = Me.img16.ListImages(1).Picture
         Else
                 vsgrid.MousePointer = flexDefault
                 Set vsgrid.MouseIcon = Nothing
         End If
    Else
        intX = CSng(vsgrid.ColWidth(0)): intW = CSng(vsgrid.ColWidth(0) + vsgrid.ColWidth(1))
        intX2 = vsgrid.ColWidth(0) + vsgrid.ColWidth(1) + vsgrid.ColWidth(2) + vsgrid.ColWidth(3)
        intW2 = intX2 + vsgrid.ColWidth(4)
        If (X > intX And X < intW And vsgrid.Cell(flexcpData, vsgrid.MouseRow, 1) <> "") Or (X > intX2 And X < intW2 And vsgrid.Cell(flexcpData, vsgrid.MouseRow, 4) <> "") Then
         vsgrid.MousePointer = flexCustom
         Set vsgrid.MouseIcon = Me.img16.ListImages(1).Picture
        Else
             vsgrid.MousePointer = flexDefault
             Set vsgrid.MouseIcon = Nothing
        End If
    End If
   
End Sub
''病历文件全部范文导出
Private Function StartExportToXMLs() As Boolean
    Dim strListPath As String, strInfoPath As String, strPath As String, strRowData As String, strRows As String, strPathZip As String
    Dim i As Long, j As Long, lngRecId As Long, lngDemoId As Long, lngTime As Long, lngRow As Long
    Dim strItemArr As Variant                '行字符串数据的数组
    Dim strRowArr As Variant                 '选中的行号数组
    Dim oDocDemosList As New DOMDocument     'Demo列表文档
    Dim oDocDemosInfo As New DOMDocument     'Demo信息文档
    Dim oRootDemosList As IXMLDOMElement     'Demo列表根节点
    Dim oRootInfo As IXMLDOMElement          'Demo信息根节点
    Dim cEPRDoc As New cEPRDocument          '文档对象
    Dim oKind As IXMLDOMElement              '种类节点
    Dim oFile As IXMLDOMElement              '文件节点
    Dim oDemo As IXMLDOMElement              '范文节点
    Dim oTempNode As IXMLDOMElement          '临时节点
        '普通住院病历
        On Error Resume Next
        strPath = zl9ComLib.OS.OpenDir(Me.hWnd, "指定导出目录")
        If strPath = "" Then Exit Function
        strPathZip = strPath & "\" & zl9ComLib.GetUnitName & "_范文.ZIP"
        If gobjFSO.FileExists(strPathZip) Then
            If MsgBox("该文件已经存在，是否替换？", vbOKCancel + vbQuestion, gstrSysName) = vbOK Then
                gobjFSO.DeleteFile (strPathZip)
            Else
             Exit Function
            End If
        End If
        strListPath = strPath & "\" & zl9ComLib.GetUnitName & "_范文列表.xml"
        strInfoPath = strPath & "\" & zl9ComLib.GetUnitName & "_范文信息.xml"
        '创建Demo信息根节点
        If oRootDemosList Is Nothing Then
            Set oRootDemosList = oDocDemosList.createElement("EPRDemosList")
            Call oRootDemosList.setAttribute("UnitName", zl9ComLib.GetUnitName)
            Set oDocDemosList.documentElement = oRootDemosList   '设置为根节点
        End If
        '创建Demo列表根节点
        If oRootInfo Is Nothing Then
            Set oRootInfo = oDocDemosInfo.createElement("EPRDemosInfo")
            Call oRootInfo.setAttribute("UnitName", zl9ComLib.GetUnitName)
            Set oDocDemosInfo.documentElement = oRootInfo        '设置为根节点
        End If
        On Error GoTo errHand
            EnableControlBar Me, False    '禁用窗体最大/小化、关闭功能
            lngTime = GetTickCount
            '计算导出范文的个数
            For i = 0 To vsgrid.Rows - 1
                If Not vsgrid.Cell(flexcpPicture, i, 4) Is Nothing Then
                    strRows = strRows & "," & i
                End If
            Next i
            strRowArr = Split(Mid(strRows, 2, Len(strRows)), ",")
            '开始循环导出
            For lngRow = 0 To UBound(strRowArr)
                DoEvents
                i = strRowArr(lngRow)
                If Not vsgrid.Cell(flexcpPicture, i, 4) Is Nothing Then
                    '获取该行Demo字符串数据
                    strRowData = GetRowsData(i)
                    If strRowData = "" Then MsgBox vsgrid.TextMatrix(lngRow, col_范文名称) & ": 内容格式不正确或该数据已损坏 ！", vbInformation, gstrSysName: Exit Function
                    '将该行字符串拆分为数组
                    strItemArr = Split(strRowData, "|")
                    lngDemoId = Val(strItemArr(3))
                    If gobjFSO.FileExists(strListPath) Then
                        If MsgBox("该文件已经存在，是否覆盖？", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then Exit Function
                    End If
                    Debug.Print vsgrid.TextMatrix(lngRow, col_范文名称)
                    If Val(strItemArr(7)) <> 2 Then  '不等于表格格式
                        '创建种类节点
                        Set oTempNode = oRootDemosList.selectNodes("/EPRDemosList/Kind[@KindName='" & strItemArr(0) & "']")(0)
                        '判断是否已经存在
                        If oTempNode Is Nothing Then
                            Set oKind = CreateNode(1, oRootDemosList, "Kind", NODE_ELEMENT, "")
                            Call oKind.setAttribute("KindName", strItemArr(0))
                        End If
                        '创建文件节点
                        Set oTempNode = oRootDemosList.selectNodes("/EPRDemosList/Kind[@KindName='" & strItemArr(0) & "']/File[@FileName='" & strItemArr(2) & "']")(0)
                        If oTempNode Is Nothing Then
                            Set oFile = CreateNode(1, oKind, "File", NODE_ELEMENT, "")
                            Call oFile.setAttribute("FileName", strItemArr(2))
                            Set oTempNode = oFile
                        End If
                        progBar.Visible = True
                        Call ExportDemosToXML(strRowData, lngDemoId, oTempNode, oRootInfo)
                        progBar.Value = IIf(progBar.Value + progBar.Max / (UBound(strRowArr) + 1) > progBar.Max, progBar.Max, progBar.Value + progBar.Max / (UBound(strRowArr) + 1))
                    End If
                End If
            Next lngRow
        oDocDemosList.Save strListPath
        oDocDemosInfo.Save strInfoPath
        '压缩XML文件
        Call zlFilesZip(strListPath & "," & strInfoPath, strPathZip)
        MsgBox "导出完成！", vbOKOnly + vbInformation, gstrSysName
        Unload Me
        StartExportToXMLs = True
        Exit Function
errHand:
    progBar.Value = 0
    progBar.Visible = False
    EnableControlBar Me, True '恢复窗体最大/小化、关闭功能
    StartExportToXMLs = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
'开始导入XML文件
Private Function StartImportFromXML() As Boolean
    Dim oDoc As New DOMDocument                 'xml文档
    Dim oRoot  As IXMLDOMElement            '根节点
    Dim oDemoNodeList As IXMLDOMNodeList    '临时文件节点集合
    Dim oDemoNode As IXMLDOMElement         '范文节点
    Dim strCheckItems As String, strNumber As String, strRows As String
    Dim i As Long, j As Long, k As Long, lngDemoId As Long, lngFileID As Long, lngRow As Long, lngTime As Long
    Dim strTermsArr As Variant, strItemArr As Variant, strSQLArr As Variant, strRowArr As Variant
    On Error GoTo errHand
    With vsgrid
        '计算导入范文的个数
        For i = 1 To .Rows - 1
            If Not .Cell(flexcpPicture, i, 1) Is Nothing And .RowOutlineLevel(i) = 1 Then
                strRows = strRows & "," & i
            End If
        Next i
        If strRows = "" Then MsgBox "请选择需要导入的范文！", vbInformation, gstrSysName: Exit Function
        progBar.Visible = True
        EnableControlBar Me, False  '禁用窗体最大/小化、关闭功能
        oDoc.Load Replace(Me.cboList.Tag, "_范文列表.xml", "_范文信息.xml")    '替换路径为范文信息文件
        gobjFSO.DeleteFolder (gobjFSO.GetParentFolderName(Me.cboList.Tag))     '删除临时文件夹
        Me.cboList.Tag = "" '清空.Tag
        '读取文件根节点
        Set oRoot = oDoc.selectSingleNode("EPRDemosInfo")
        If oRoot Is Nothing Then MsgBox "此文件内容格式不正确，不能在此处导入该文件！", vbInformation, gstrSysName: Exit Function
        '获取范文节点集合
        Set oDemoNodeList = oRoot.selectNodes("/EPRDemosInfo/Demo")
        If oDemoNodeList.Length < 1 Then MsgBox "此文件数据内容可能为空，不能在此处导入该文件！", vbInformation, gstrSysName: Exit Function
        '开始循环导入
        strRowArr = Split(Mid(strRows, 2, Len(strRows)), ",")
        For lngRow = 0 To UBound(strRowArr)
                DoEvents
                i = strRowArr(lngRow)
                strTermsArr = Split(.Cell(flexcpData, i, 5), "|")
                lngFileID = Val(.Cell(flexcpData, i, 3))                                    '文件ID
                lngDemoId = zlDatabase.GetNextId("病历范文目录")                            '获取新增ID
                strNumber = GetMax("病历范文目录", "编号", 5, " Where 文件id=" & lngFileID) '新增编号
                gstrSQL = lngDemoId & "," & lngFileID & ",'" & strNumber & "','" & strTermsArr(0) & "','" & strTermsArr(1) & "'," & 0
                gstrSQL = gstrSQL & ",'" & strTermsArr(2) & "'," & strTermsArr(3) & "," & glngDeptId & "," & glngUserId & ",'" & strTermsArr(4) & "'"
                gstrSQL = "Zl_病历范文目录_Insert(" & gstrSQL & ")"
                If strTermsArr(5) <> "0" Then
                    strTermsArr = Split(strTermsArr(5), ";")
                    For j = 0 To UBound(strTermsArr)
                         gstrSQL = "Zl_病历范文条件_Edit(' " & lngDemoId & " ','" & Split(strTermsArr(j), ":")(0) & "','" & Split(strTermsArr(j), ":")(1) & "')"
                          Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                    Next j
                End If
                Set oDemoNode = oRoot.selectSingleNode("/EPRDemosInfo/Demo[@ID='" & .Cell(flexcpData, i, 4) & "']")
                ImportDemosFromXML oDemoNode, lngDemoId, strTermsArr(0), .Cell(flexcpData, i, 4), gstrSQL
                k = k + 1
                progBar.Value = IIf(progBar.Value + progBar.Max / (UBound(strRowArr) + 1) > progBar.Max, progBar.Max, progBar.Value + progBar.Max / (UBound(strRowArr) + 1))
        Next lngRow
        Dim strMsg As String
        MsgBox "导入完成！", vbOKOnly + vbInformation, gstrSysName
        Unload Me
    End With
    StartImportFromXML = True
    Exit Function
errHand:
    progBar.Value = 0
    progBar.Visible = False
    StartImportFromXML = False
    EnableControlBar Me, True  '禁用窗体最大/小化、关闭功能
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
'################################################################################################################
'## 功能：  将范文文件批量导出到一个XML文档中
'##
'## 参数：  strModel     :   范文数据字符串
'##         lngDemoId    :   范文ID
'##         oFileNode    :   文件节点，（属于用于存储范文列表DOC）
'##         oContextNode :   范文内容节点，（属于用于存储范文信息的DOC）
'## 返回：  保存成功，返回Ture；否则返回False。
'################################################################################################################
Private Function ExportDemosToXML(ByVal strModel As String, ByVal lngDemoId As Long, ByRef oFileNode As IXMLDOMElement, ByRef oContextNode As IXMLDOMElement) As Boolean
    Dim i As Long, j As Long, k As Long
    Dim oDoc As New DOMDocument
    Dim oDemoRoot As IXMLDOMElement     '范文节点
    Dim oRootDemo As IXMLDOMElement     '根节点
    Dim oNode As IXMLDOMElement         '父节点
    Dim CompendsoNode As IXMLDOMNode    '提纲节点
    Dim ElementsNode As IXMLDOMNode     '要素节点
    Dim PicturesNode As IXMLDOMNode     '图片节点
    Dim DiagnosisesNode As IXMLDOMNode  '诊断
    Dim TablesNode As IXMLDOMNode       '表格节点
    Dim TableCells As IXMLDOMNode       '表格中文本集合节点
    Dim TableElements As IXMLDOMNode    '表格中要素集合节点
    Dim TablePictures As IXMLDOMNode    '表格中图片集合节点
    Dim CellNode As IXMLDOMNode         '单元格节点
    Dim ContentNode As IXMLDOMNode      '内容节点
    Dim oSubNode1 As IXMLDOMNode        '子节点
    Dim oSubNode2 As IXMLDOMNode        '节点
    Dim oSubNode3 As IXMLDOMNode        '节点
    Dim oSubNode4 As IXMLDOMNode        '节点
    Dim oSubNode5 As IXMLDOMNode        '节点
    Dim oStream As New ADODB.Stream     '流对象
    Dim strPath As String               '临时文件目录
    Dim strPic As String                '临时图片文件
    Dim TempPic As New StdPicture, strTempPic As String
    Dim strObjArr As Variant
    Dim strItemArr As Variant, strTermsArr As Variant
    Dim strContextFile As String, strTemp As String
    Dim rs As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    '----------------------------------------------------------------------------
    On Error GoTo errHand:
    strItemArr = Split(strModel, "|")
    strTermsArr = Split(strItemArr(UBound(strItemArr)), ";")
    strPath = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))
    '存储到范文列表DOC中
    Set oRootDemo = CreateNode(1, oFileNode, "Demo", NODE_ELEMENT, "")
    Call oRootDemo.setAttribute("EditType", 1)
    Call oRootDemo.setAttribute("ID", lngDemoId)
    Call oRootDemo.setAttribute("Items", strItemArr(5) & "|" & strItemArr(6) & "|" & strItemArr(8) & "|" & strItemArr(9) & "|" & strItemArr(10) & "|" & IIf(UBound(strTermsArr) > 0, strItemArr(UBound(strItemArr)), 0))
    '存储到范文信息DOC中
    Set oDemoRoot = CreateNode(1, oContextNode, "Demo", NODE_ELEMENT, "")
    Call oDemoRoot.setAttribute("EditType", 1)
    Call oDemoRoot.setAttribute("ID", lngDemoId)
    '导出内容RTF文本
    strContextFile = zlBlobRead(3, lngDemoId)
    If strContextFile <> "" Then
       strTemp = zlFileUnzip(strContextFile)
       Me.RTbContext.LoadFile strTemp
       gobjFSO.DeleteFile strTemp
       gobjFSO.DeleteFile strContextFile, True
    End If
    '读取范文结构
    gstrSQL = "Select Level, ID, 文件id, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id,定义提纲ID, 复用提纲, 使用时机," & vbNewLine & _
            "       诊治要素id, 替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域" & vbNewLine & _
            "From (Select ID, 文件id, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行,预制提纲id,定义提纲ID,复用提纲,使用时机," & vbNewLine & _
            "              诊治要素id, 替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域" & vbNewLine & _
            "       From 病历范文内容" & vbNewLine & _
            "       Where 文件id = [1] And 对象序号 <> 0)" & vbNewLine & _
            "Start With 父id Is Null" & vbNewLine & _
            "Connect By Prior ID = 父id" & vbNewLine & _
            "Order By 对象序号, 内容行次"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngDemoId)
    Do While Not rs.EOF
        Select Case NVL(rs("对象类型"), 2)
            Case 1  '提纲
                If CompendsoNode Is Nothing Then Set CompendsoNode = CreateNode(1, oDemoRoot, "Compends", NODE_ELEMENT, "")
                 Set oSubNode1 = CreateNode(2, CompendsoNode, "Compend", NODE_ELEMENT, "")
                    CreateNode 3, oSubNode1, "Key", , NVL(rs!对象标记, "")
                    CreateNode 3, oSubNode1, "ID", , rs!ID
                    CreateNode 3, oSubNode1, "文件ID", , NVL(rs!文件ID, 0)
                    CreateNode 3, oSubNode1, "父ID", , 0
                    CreateNode 3, oSubNode1, "对象序号", , NVL(rs!对象序号, 0)
                    CreateNode 3, oSubNode1, "保留对象", , IIf(NVL(rs!保留对象, 0) = 0, False, True)
                    CreateNode 3, oSubNode1, "名称", , NVL(rs!内容文本)
                    CreateNode 3, oSubNode1, "说明", , NVL(rs!对象属性)
                    CreateNode 3, oSubNode1, "预制提纲ID", , NVL(rs!预制提纲ID, 0)
                    CreateNode 3, oSubNode1, "定义提纲ID", , NVL(rs!定义提纲ID, 0)
                    CreateNode 3, oSubNode1, "复用提纲", , IIf(NVL(rs!复用提纲, 0) = 0, False, True)
                    CreateNode 3, oSubNode1, "使用时机", , NVL(rs!使用时机)
                    CreateNode 3, oSubNode1, "Level", , NVL(rs!Level, 0)
                    CreateNode 3, oSubNode1, "内部序号", , NVL(rs!对象序号, 0)
            Case 3  '表格
                If TablesNode Is Nothing Then Set TablesNode = CreateNode(1, oDemoRoot, "Tables", NODE_ELEMENT, "")
                Set oSubNode1 = CreateNode(2, TablesNode, "Table", NODE_ELEMENT, "")
                    CreateNode 3, oSubNode1, "Key", , NVL(rs!对象标记, "")
                    CreateNode 3, oSubNode1, "ID", , NVL(rs!ID, 0)
                    CreateNode 3, oSubNode1, "文件ID", , NVL(rs!文件ID, 0)
                    CreateNode 3, oSubNode1, "父ID", , NVL(rs!父ID, 0)
                    CreateNode 3, oSubNode1, "对象序号", , NVL(rs!对象序号, 0)
                    CreateNode 3, oSubNode1, "保留对象", , IIf(NVL(rs!保留对象, 0) = 0, False, True)
                    CreateNode 3, oSubNode1, "是否换行", , IIf(NVL(rs!是否换行, 0) = 0, False, True)
                    CreateNode 3, oSubNode1, "预制提纲ID", , NVL(rs!预制提纲ID, 0)
                    CreateNode 3, oSubNode1, "对象属性", , NVL(rs!对象属性)
                    gstrSQL = "Select Level, ID, 文件id, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id," & _
                              "诊治要素ID , 替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域 " & _
                              "From (Select ID, 文件id, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id, 诊治要素id, " & _
                              "替换域 , 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域 From 病历范文内容 " & _
                              "Where 文件id = " & NVL(rs!文件ID, 0) & " ) Start With 父id + 0 = " & NVL(rs!ID, 0) & " Connect By Prior ID = 父id + 0 Order By 对象标记, 对象序号, 内容行次"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                    Set TableCells = CreateNode(2, oSubNode1, "Cells", NODE_ELEMENT, "")
                    Set TableElements = CreateNode(2, oSubNode1, "Elements", NODE_ELEMENT, "")
                    Set TablePictures = CreateNode(2, oSubNode1, "Pictures", NODE_ELEMENT, "")
                    Do While Not rsTemp.EOF
                        Select Case NVL(rsTemp!对象类型, 0)
                            Case 2  '文本
                                Set CellNode = CreateNode(3, TableCells, "Cell", NODE_ELEMENT, "")
                                    CreateNode 4, CellNode, "Key", , NVL(rsTemp!对象标记, "")
                                    CreateNode 4, CellNode, "ID", , NVL(rsTemp!ID, 0)
                                    CreateNode 4, CellNode, "文件ID", , NVL(rsTemp!文件ID, 0)
                                    CreateNode 4, CellNode, "父ID", , NVL(rsTemp!父ID, 0)
                                    CreateNode 4, CellNode, "对象序号", , NVL(rsTemp!对象序号, 0)
                                    CreateNode 4, CellNode, "内容文本", , NVL(rsTemp!内容文本, "")
                                    CreateNode 4, CellNode, "保留对象", , IIf(NVL(rsTemp!保留对象, 0) = 0, False, True)
                                    CreateNode 4, CellNode, "对象属性", , NVL(rsTemp!对象属性)
                            Case 4  '诊治要素
                                 '先增加单元 Cell 到 Cells集合节点中
                                 Set CellNode = CreateNode(3, TableCells, "Cell", NODE_ELEMENT, "")
                                    CreateNode 4, CellNode, "Key", , NVL(rsTemp!对象标记, "")
                                    CreateNode 4, CellNode, "ID", , NVL(rsTemp!ID, 0)
                                    CreateNode 4, CellNode, "文件ID", , NVL(rsTemp!文件ID, 0)
                                    CreateNode 4, CellNode, "父ID", , NVL(rsTemp!父ID, 0)
                                    CreateNode 4, CellNode, "对象序号", , NVL(rsTemp!对象序号, 0)
                                    CreateNode 4, CellNode, "内容文本", , NVL(rsTemp!内容文本, "")
                                    CreateNode 4, CellNode, "保留对象", , IIf(NVL(rsTemp!保留对象, 0) = 0, False, True)
                                    CreateNode 4, CellNode, "对象属性", , NVL(rsTemp!对象属性)
                                 '再增加诊治要素 Element 到 Elements集合节点中
                                Set oSubNode3 = CreateNode(3, TableElements, "Element", NODE_ELEMENT, "")
                                    CreateNode 4, oSubNode3, "Key", , NVL(rsTemp!对象标记, "")
                                    CreateNode 4, oSubNode3, "ID", , rsTemp!ID
                                    CreateNode 4, oSubNode3, "文件ID", , NVL(rsTemp!文件ID, 0)
                                    CreateNode 4, oSubNode3, "父ID", , NVL(rsTemp!父ID, 0)
                                    CreateNode 4, oSubNode3, "对象序号", , NVL(rsTemp!对象序号, 0)
                                    CreateNode 4, oSubNode3, "保留对象", , IIf(NVL(rsTemp!保留对象, 0) = 0, False, True)
                                    CreateNode 4, oSubNode3, "内容文本", , NVL(rsTemp!内容文本)
                                    CreateNode 4, oSubNode3, "是否换行", , IIf(NVL(rsTemp!是否换行, 0) = 0, False, True)
                                    CreateNode 4, oSubNode3, "诊治要素ID", , NVL(rsTemp!诊治要素ID, 0)
                                    CreateNode 4, oSubNode3, "替换域", , NVL(rsTemp!替换域, 0)
                                    CreateNode 4, oSubNode3, "要素名称", , NVL(rsTemp!要素名称)
                                    CreateNode 4, oSubNode3, "要素类型", , NVL(rsTemp!要素类型, 0)
                                    CreateNode 4, oSubNode3, "要素长度", , NVL(rsTemp!要素长度, 0)
                                    CreateNode 4, oSubNode3, "要素小数", , NVL(rsTemp!要素小数, 0)
                                    CreateNode 4, oSubNode3, "要素单位", , NVL(rsTemp!要素单位)
                                    CreateNode 4, oSubNode3, "要素表示", , NVL(rsTemp!要素表示, 0)
                                    CreateNode 4, oSubNode3, "输入形态", , NVL(rsTemp!输入形态, 0)
                                    CreateNode 4, oSubNode3, "要素值域", , NVL(rsTemp!要素值域)
                                    CreateNode 4, oSubNode3, "对象属性", , NVL(rsTemp!对象属性)
                            Case 5  '图片
                                 Set oSubNode3 = CreateNode(3, TablePictures, "Picture", NODE_ELEMENT, "")
                                    CreateNode 4, oSubNode3, "Key", , NVL(rsTemp!对象标记, "")
                                    CreateNode 4, oSubNode3, "ID", , NVL(rsTemp!ID, 0)
                                    CreateNode 4, oSubNode3, "文件ID", , NVL(rsTemp!文件ID, 0)
                                    CreateNode 4, oSubNode3, "父ID", , NVL(rsTemp!父ID, 0)
                                    CreateNode 4, oSubNode3, "对象序号", , NVL(rsTemp!对象序号, 0)
                                    CreateNode 4, oSubNode3, "保留对象", , IIf(NVL(rsTemp!保留对象, 0) = 0, False, True)
                                    CreateNode 4, oSubNode3, "内容文本", , NVL(rsTemp!内容文本, "")
                                    CreateNode 4, oSubNode3, "是否换行", , IIf(NVL(rsTemp!是否换行, 0) = 0, False, True)
                                    CreateNode 4, oSubNode3, "对象属性", , NVL(rsTemp!对象属性, "")
                                    '存储图片对象
                                    strTempPic = zlBlobRead(4, rsTemp!ID)
                                    Set TempPic = LoadPicture(strTempPic)
                                    gobjFSO.DeleteFile strTempPic, True      '删除临时文件
                                    oStream.Type = adTypeBinary
                                    oStream.Open
                                    strPic = strPath & "\XMLPIC" & App.hInstance & ".jpg"
                                    SavePicture TempPic, strPic
                                    oStream.LoadFromFile strPic
                                    Set oSubNode4 = oDoc.createElement("OrigPic")
                                    oSubNode4.datatype = "bin.base64"
                                    oSubNode4.nodeTypedValue = oStream.Read
                                    oSubNode3.appendChild oSubNode4
                                    oStream.Close
                                    '删除临时文件
                                    If gobjFSO.FileExists(strPic) Then gobjFSO.DeleteFile strPic, True
                        End Select
                        rsTemp.MoveNext
                    Loop
            Case 4  '要素
                 If ElementsNode Is Nothing Then Set ElementsNode = CreateNode(1, oDemoRoot, "Elements", NODE_ELEMENT, "")
                 Set oSubNode1 = CreateNode(2, ElementsNode, "Element", NODE_ELEMENT, "")
                    CreateNode 3, oSubNode1, "Key", , NVL(rs!对象标记, "")
                    CreateNode 3, oSubNode1, "ID", , rs!ID
                    CreateNode 3, oSubNode1, "文件ID", , NVL(rs!文件ID, 0)
                    CreateNode 3, oSubNode1, "父ID", , NVL(rs!父ID, 0)
                    CreateNode 3, oSubNode1, "对象序号", , NVL(rs!对象序号, 0)
                    CreateNode 3, oSubNode1, "保留对象", , IIf(NVL(rs!保留对象, 0) = 0, False, True)
                    CreateNode 3, oSubNode1, "内容文本", , NVL(rs!内容文本)
                    CreateNode 3, oSubNode1, "是否换行", , IIf(NVL(rs!是否换行, 0) = 0, False, True)
                    CreateNode 3, oSubNode1, "诊治要素ID", , NVL(rs!诊治要素ID, 0)
                    CreateNode 3, oSubNode1, "替换域", , NVL(rs!替换域, 0)
                    CreateNode 3, oSubNode1, "要素名称", , NVL(rs!要素名称)
                    CreateNode 3, oSubNode1, "要素类型", , NVL(rs!要素类型, 0)
                    CreateNode 3, oSubNode1, "要素长度", , NVL(rs!要素长度, 0)
                    CreateNode 3, oSubNode1, "要素小数", , NVL(rs!要素小数, 0)
                    CreateNode 3, oSubNode1, "要素单位", , NVL(rs!要素单位)
                    CreateNode 3, oSubNode1, "要素表示", , NVL(rs!要素表示, 0)
                    CreateNode 3, oSubNode1, "输入形态", , NVL(rs!输入形态, 0)
                    CreateNode 3, oSubNode1, "要素值域", , NVL(rs!要素值域)
                    CreateNode 3, oSubNode1, "对象属性", , NVL(rs!对象属性)
            Case 5  '图片
                 If PicturesNode Is Nothing Then Set PicturesNode = CreateNode(1, oDemoRoot, "Pictures", NODE_ELEMENT, "")
                 Set oSubNode1 = CreateNode(2, PicturesNode, "Picture", NODE_ELEMENT, "")
                    CreateNode 3, oSubNode1, "Key", , NVL(rs!对象标记, "")
                    CreateNode 3, oSubNode1, "ID", , NVL(rs!ID, 0)
                    CreateNode 3, oSubNode1, "文件ID", , NVL(rs!文件ID, 0)
                    CreateNode 3, oSubNode1, "父ID", , NVL(rs!父ID, 0)
                    CreateNode 3, oSubNode1, "对象序号", , NVL(rs!对象序号, 0)
                    CreateNode 3, oSubNode1, "保留对象", , IIf(NVL(rs!保留对象, 0) = 0, False, True)
                    CreateNode 3, oSubNode1, "内容文本", , NVL(rs!内容文本)
                    CreateNode 3, oSubNode1, "是否换行", , IIf(NVL(rs!是否换行, 0) = 0, False, True)
                    CreateNode 3, oSubNode1, "对象属性", , NVL(rs!对象属性, "")
                    '存储图片对象
                    strTempPic = zlBlobRead(4, rs!ID)
                    If strTempPic <> "" Then
                        Set TempPic = LoadPicture(strTempPic)
                        gobjFSO.DeleteFile strTempPic, True      '删除临时文件
                        oStream.Type = adTypeBinary
                        oStream.Open
                        strPic = strPath & "\XMLPIC" & App.hInstance & ".jpg"
                        SavePicture TempPic, strPic
                        oStream.LoadFromFile strPic
                        Set oSubNode2 = oDoc.createElement("OrigPic")
                        oSubNode2.datatype = "bin.base64"
                        oSubNode2.nodeTypedValue = oStream.Read
                        oSubNode1.appendChild oSubNode2
                        oStream.Close
                        '删除临时文件
                        If gobjFSO.FileExists(strPic) Then gobjFSO.DeleteFile strPic, True
                    End If
            Case 7  '诊断
                If DiagnosisesNode Is Nothing Then Set DiagnosisesNode = CreateNode(1, oDemoRoot, "Diagnosises", NODE_ELEMENT, "")
                Set oSubNode1 = CreateNode(2, DiagnosisesNode, "Diagnosis", NODE_ELEMENT, "")
                CreateNode 3, oSubNode1, "Key", , NVL(rs!对象标记, "")
                CreateNode 3, oSubNode1, "ID", , NVL(rs!ID, 0)
                CreateNode 3, oSubNode1, "文件ID", , NVL(rs!文件ID, 0)
                CreateNode 3, oSubNode1, "父ID", , NVL(rs!父ID, 0)
                CreateNode 3, oSubNode1, "对象序号", , NVL(rs!对象序号, 0)
                CreateNode 3, oSubNode1, "描述", , NVL(rs!内容文本, "")
                CreateNode 3, oSubNode1, "对象属性", , NVL(rs!对象属性, "")
        End Select
        rs.MoveNext
    Loop
    'RTF文本
    Set oNode = CreateNode(1, oDemoRoot, "Content", NODE_ELEMENT, "")
    Set oSubNode1 = CreateNode(2, oNode, "RTF", NODE_ELEMENT, "")
    CreateNode 3, oSubNode1, "RTFText", NODE_CDATA_SECTION, Replace(Me.RTbContext.TextRTF, "]]>", "]] >")
    ExportDemosToXML = True
    Exit Function
errHand:
    ExportDemosToXML = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
'################################################################################################################
'## 功能：  批量导入范文
'## 参数：  oDemoNode       :范文节点
'##         lngFileID       :文件ID
'##         strFileName     :文件名称
'##         lngOldId        :就范文ID
'##         strSql          :新增范文SQL语句
'## 返回：  保存成功，返回Ture；否则返回False。
'################################################################################################################
Private Function ImportDemosFromXML(ByVal oDemoNode As IXMLDOMElement, ByVal lngFileID As Long, ByVal strFileName As String, ByVal lngOldId As Long, ByVal strSQL As String) As Boolean
    Dim oNodeList As IXMLDOMNodeList    '节点集合
    Dim oNode As IXMLDOMNode            '子节点
    Dim oSubNode1 As IXMLDOMNode        '子节点
    Dim oSubNode2 As IXMLDOMNode        '子节点
    Dim EPRFileInfoNode As IXMLDOMNode  '基础信息节点
    Dim Compends As IXMLDOMNodeList     '提纲节点
    Dim Elements As IXMLDOMNodeList     '要素节点
    Dim Pictures As IXMLDOMNodeList     '图片节点
    Dim Diagnosises As IXMLDOMNodeList
    Dim Tables As IXMLDOMNodeList       '表格节点
    Dim Cells As IXMLDOMNodeList        '表格中文本集合节点
    Dim TableElements As IXMLDOMNodeList    '表格中要素集合节点
    Dim TablePictures As IXMLDOMNodeList    '表格中图片集合节点
    Dim CellNode As IXMLDOMNode         '单元格节点
    Dim ContentNode As IXMLDOMNode      '内容节点
    Dim oStream As New ADODB.Stream     '流对象
    Dim strPath As String               '临时文件目录
    Dim strTemp As String               '临时文件
    Dim strPic As String                '临时图片文件
    Dim strContextFile As String        '临时内容文
    Dim strArrNames As Variant          '文件名称数组
    Dim ArraySQL() As String            'SQL数组
    Dim lngID As Long, lng行次 As Long
    Dim TempPic As New StdPicture, strTempPic As String
    '------------------------------------
    Dim GpInput As GdiplusStartupInput
    Dim m_GDIpToken         As Long         ' 用于关闭 GDI+
    Dim oDIB As New cDIB
    Dim DIBDither As New cDIBDither
    Dim DIBPal As New cDIBPal
    '-------------------------------------------------------------------------
    '从XML提取提纲信息
    On Error GoTo errHand
    ReDim ArraySQL(1 To 2) As String
    ArraySQL(1) = strSQL
    If Not oDemoNode.selectSingleNode("Compends") Is Nothing Then Set Compends = oDemoNode.selectSingleNode("Compends").selectNodes("Compend")
    If Not Compends Is Nothing Then
        For Each oNode In Compends
            lngID = zlDatabase.GetNextId("病历范文目录")
            ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
            ArraySQL(UBound(ArraySQL)) = "Zl_病历范文内容_Update(" & lngID & "," & lngFileID & "," & IIf(oNode.selectSingleNode("父ID").Text = 0, "NULL", oNode.selectSingleNode("父ID").Text) & "," & _
                oNode.selectSingleNode("对象序号").Text & ",1," & oNode.selectSingleNode("Key").Text & "," & IIf(oNode.selectSingleNode("保留对象").Text, 1, 0) & ",'" & oNode.selectSingleNode("说明").Text & "',NULL,'" & oNode.selectSingleNode("名称").Text & "',NULL," & _
                IIf(oNode.selectSingleNode("定义提纲ID").Text = 0, "NULL", oNode.selectSingleNode("定义提纲ID").Text) & "," & IIf(oNode.selectSingleNode("预制提纲ID").Text = 0, "NULL", oNode.selectSingleNode("预制提纲ID").Text) & "," & IIf(oNode.selectSingleNode("复用提纲").Text, 1, 0) & ",'" & oNode.selectSingleNode("使用时机").Text & "')"
            '改变所有子项父ID
            Set oNodeList = oDemoNode.selectNodes("/EPRDemosInfo/Demo[@ID='" & lngOldId & "']//父ID[text()=" & oNode.selectSingleNode("ID").Text & "]")
            For Each oSubNode1 In oNodeList
                oSubNode1.Text = lngID
            Next
        Next
    End If
    Debug.Print lngID
    Debug.Print '--------------------------------'
    '从XML提取要素信息
    If Not oDemoNode.selectSingleNode("Elements") Is Nothing Then Set Elements = oDemoNode.selectSingleNode("Elements").selectNodes("Element")
    If Not Elements Is Nothing Then
        For Each oNode In Elements
            ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
            ArraySQL(UBound(ArraySQL)) = "Zl_病历范文内容_Update(" & zlDatabase.GetNextId("病历范文目录") & "," & lngFileID & "," & IIf(oNode.selectSingleNode("父ID").Text = 0, "NULL", oNode.selectSingleNode("父ID").Text) & "," & _
                oNode.selectSingleNode("对象序号").Text & ",4," & oNode.selectSingleNode("Key").Text & "," & IIf(oNode.selectSingleNode("保留对象").Text, 1, 0) & ",'" & oNode.selectSingleNode("对象属性").Text & "',NULL,'" & _
                Replace(oNode.selectSingleNode("内容文本").Text, "'", "' || chr(39) || '") & "'," & IIf(oNode.selectSingleNode("是否换行").Text, 1, 0) & ",NULL,NULL,NULL," & _
                 "NULL," & IIf(CheckValid(oNode.selectSingleNode("诊治要素ID").Text, oNode.selectSingleNode("要素名称").Text), oNode.selectSingleNode("诊治要素ID").Text, "NULL") & "," & _
                oNode.selectSingleNode("替换域").Text & ",'" & oNode.selectSingleNode("要素名称").Text & "'," & oNode.selectSingleNode("要素类型").Text & "," & oNode.selectSingleNode("要素长度").Text & "," & _
                oNode.selectSingleNode("要素小数").Text & ",'" & oNode.selectSingleNode("要素单位").Text & "'," & oNode.selectSingleNode("要素表示").Text & "," & oNode.selectSingleNode("输入形态").Text & ",'" & oNode.selectSingleNode("要素值域").Text & "')"
        Next
    End If
    '从XML提取表格信息
    If Not oDemoNode.selectSingleNode("Tables") Is Nothing Then Set Tables = oDemoNode.selectSingleNode("Tables").selectNodes("Table")
    If Not Tables Is Nothing Then
        For Each oNode In Tables
            lngID = zlDatabase.GetNextId("病历范文目录")
            '保存表格结构SQL语句
            ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
            ArraySQL(UBound(ArraySQL)) = "Zl_病历范文内容_Update(" & lngID & "," & lngFileID & "," & IIf(oNode.selectSingleNode("父ID").Text = 0, "NULL", oNode.selectSingleNode("父ID").Text) & "," & _
            oNode.selectSingleNode("对象序号").Text & ",3," & oNode.selectSingleNode("Key").Text & "," & IIf(oNode.selectSingleNode("保留对象").Text, 1, 0) & ",'" & oNode.selectSingleNode("对象属性").Text & "',NULL,'" & "" & "'," & IIf(oNode.selectSingleNode("是否换行").Text, 1, 0) & _
            "," & IIf(oNode.selectSingleNode("预制提纲ID").Text = 0, "NULL", oNode.selectSingleNode("预制提纲ID").Text) & ")"
            '更改所有子项的父ID
            For Each oSubNode1 In oNode.selectNodes("/EPRDemosInfo/Demo[@ID='" & lngOldId & "']//父ID[text()=" & oNode.selectSingleNode("ID").Text & "]")
                oSubNode1.Text = lngID
            Next
            '获取表格中的元素
            If Not oNode.selectSingleNode("Cells") Is Nothing Then Set Cells = oNode.selectSingleNode("Cells").selectNodes("Cell")
            If Not oNode.selectSingleNode("Elements") Is Nothing Then Set TableElements = oNode.selectSingleNode("Elements").selectNodes("Element")
            If Not oNode.selectSingleNode("Pictures") Is Nothing Then Set TablePictures = oNode.selectSingleNode("Pictures").selectNodes("Picture")
            '单元格文本及要素
            If Not Cells Is Nothing Then
                lng行次 = 1
                For Each oSubNode1 In Cells
                    Dim lngElementKey As Long
                    lngElementKey = Split(oSubNode1.selectSingleNode("对象属性").Text, "|")(0)
                    If lngElementKey > 0 Then    '要素处理
                        Set oSubNode2 = oNode.selectSingleNode("Elements").selectSingleNode("*[Key=" & oSubNode1.selectSingleNode("Key").Text & " ]")
                        If Not oSubNode2 Is Nothing Then
                            ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
                            ArraySQL(UBound(ArraySQL)) = "Zl_病历范文内容_Update(" & zlDatabase.GetNextId("病历范文目录") & "," & lngFileID & "," & IIf(oSubNode2.selectSingleNode("父ID").Text = 0, "NULL", oSubNode2.selectSingleNode("父ID").Text) & "," & _
                            IIf(oSubNode2.selectSingleNode("对象序号").Text = 0, "NULL", oSubNode2.selectSingleNode("对象序号").Text) & ",4," & oSubNode2.selectSingleNode("Key").Text & "," & IIf(oSubNode2.selectSingleNode("保留对象").Text, 1, 0) & ",'" & _
                            oSubNode2.selectSingleNode("对象属性").Text & "'," & lng行次 & ",'" & Replace(oSubNode2.selectSingleNode("内容文本").Text, "'", "' || chr(39) || '") & "'," & IIf(oSubNode2.selectSingleNode("是否换行").Text, 1, 0) & ",NULL,NULL,NULL," & _
                             "NULL," & IIf(CheckValid(oSubNode2.selectSingleNode("诊治要素ID").Text, oSubNode2.selectSingleNode("要素名称").Text), oSubNode2.selectSingleNode("诊治要素ID").Text, "NULL") & "," & _
                            oSubNode2.selectSingleNode("替换域").Text & ",'" & oSubNode2.selectSingleNode("要素名称").Text & "'," & oSubNode2.selectSingleNode("要素类型").Text & "," & oSubNode2.selectSingleNode("要素长度").Text & "," & _
                            oSubNode2.selectSingleNode("要素小数").Text & ",'" & oSubNode2.selectSingleNode("要素单位").Text & "'," & oSubNode2.selectSingleNode("要素表示").Text & "," & oSubNode2.selectSingleNode("输入形态").Text & ",'" & oSubNode2.selectSingleNode("要素值域").Text & "')"
                        End If
                    Else '文本
                        ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
                         ArraySQL(UBound(ArraySQL)) = "Zl_病历范文内容_Update(" & zlDatabase.GetNextId("病历范文目录") & "," & lngFileID & "," & oSubNode1.selectSingleNode("父ID").Text & ",NULL," & _
                        "2," & oSubNode1.selectSingleNode("Key").Text & ",NULL,'" & oSubNode1.selectSingleNode("对象属性").Text & "'," & lng行次 & ",'" & Replace(oSubNode1.selectSingleNode("内容文本").Text, "'", "' || chr(39) || '") & "')"
                    End If
                    lng行次 = lng行次 + 1
                Next
            End If
            '图片处理
            If Not TablePictures Is Nothing Then
                For Each oSubNode1 In TablePictures
                        ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
                        lngID = zlDatabase.GetNextId("病历范文目录")
                        ArraySQL(UBound(ArraySQL)) = "Zl_病历范文内容_Update(" & lngID & "," & lngFileID & "," & IIf(oSubNode1.selectSingleNode("父ID").Text = 0, "NULL", oSubNode1.selectSingleNode("父ID").Text) & "," & _
                        oSubNode1.selectSingleNode("对象序号").Text & ",5," & oSubNode1.selectSingleNode("Key").Text & "," & IIf(oSubNode1.selectSingleNode("保留对象").Text, 1, 0) & ",'" & oSubNode1.selectSingleNode("对象属性").Text & "'," & _
                         lng行次 & ",'" & oSubNode1.selectSingleNode("内容文本").Text & "'," & IIf(oSubNode1.selectSingleNode("是否换行").Text, 1, 0) & ")"
                        oStream.Type = adTypeBinary
                        oStream.Open
                        oStream.Write oSubNode1.selectSingleNode("OrigPic").nodeTypedValue
                        strPic = App.Path & "\OrigPic" & Timer & ".jpg"
                        oStream.SaveToFile strPic, adSaveCreateOverWrite
                        '-- 调入 GDI+ Dll
                        GpInput.GdiplusVersion = 1
                        If (mGdIpEx.GdiplusStartup(m_GDIpToken, GpInput) <> [OK]) Then
                            '按照BMP格式保存！会增大图片体积
                            SavePicture TempPic, strPic       '保存格式为BMP格式
                        Else
                            '采用JPEG压缩格式保存
                            Call oDIB.CreateFromStdPicture(TempPic, DIBPal, DIBDither)
                            '压缩存储
                            Call mGdIpEx.SaveDIB(oDIB, strFileName, [ImageJPEG], 100)          '90%的JPEG图片压缩质量
                        End If
                        'Unload the GDI+ Dll
                        Call mGdIpEx.GdiplusShutdown(m_GDIpToken)
                        gstrSQL = "select 对象ID from 病历范文图形 where 对象ID=[1]"
                        Call zlBlobSql(4, lngID, strPic, ArraySQL)
                        oStream.Close
                Next
            End If
        Next
    End If
    '从XML提取诊断信息
     If Not oDemoNode.selectSingleNode("Diagnosises") Is Nothing Then Set Diagnosises = oDemoNode.selectSingleNode("Diagnosises").selectNodes("Diagnosis")
     If Not Diagnosises Is Nothing Then
        For Each oNode In Diagnosises
            ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
            lngID = zlDatabase.GetNextId("病历范文目录")
            ArraySQL(UBound(ArraySQL)) = "Zl_病历范文内容_Update(" & lngID & "," & lngFileID & "," & _
            IIf(oNode.selectSingleNode("父ID").Text = 0, "NULL", oNode.selectSingleNode("父ID").Text) & "," & oNode.selectSingleNode("对象序号").Text & ",7," & _
            oNode.selectSingleNode("Key").Text & ",0,'" & oNode.selectSingleNode("对象属性").Text & "',NULL,'" & oNode.selectSingleNode("描述").Text & "')"
        Next
    End If
    '从XML提取内容图片信息
    If Not oDemoNode.selectSingleNode("Pictures") Is Nothing Then Set Pictures = oDemoNode.selectSingleNode("Pictures").selectNodes("Picture")
    If Not Pictures Is Nothing Then
        For Each oNode In Pictures
            ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
            lngID = zlDatabase.GetNextId("病历范文目录")
            ArraySQL(UBound(ArraySQL)) = "Zl_病历范文内容_Update(" & lngID & "," & lngFileID & "," & IIf(oNode.selectSingleNode("父ID").Text = 0, "NULL", oNode.selectSingleNode("父ID").Text) & "," & _
            oNode.selectSingleNode("对象序号").Text & ",5," & oNode.selectSingleNode("Key").Text & "," & IIf(oNode.selectSingleNode("保留对象").Text, 1, 0) & ",'" & oNode.selectSingleNode("对象属性").Text & "'," & _
            "NULL" & ",'" & oNode.selectSingleNode("内容文本").Text & "'," & IIf(oNode.selectSingleNode("是否换行").Text, 1, 0) & ")"
            oStream.Type = adTypeBinary
            oStream.Open
            oStream.Write oNode.selectSingleNode("OrigPic").nodeTypedValue
            strPic = App.Path & "\OrigPic" & Timer & ".jpg"
            oStream.SaveToFile strPic, adSaveCreateOverWrite
            '-- 调入 GDI+ Dll
            GpInput.GdiplusVersion = 1
            If (mGdIpEx.GdiplusStartup(m_GDIpToken, GpInput) <> [OK]) Then
                '按照BMP格式保存！会增大图片体积
                SavePicture TempPic, strPic       '保存格式为BMP格式
            Else
                '采用JPEG压缩格式保存
                Call oDIB.CreateFromStdPicture(TempPic, DIBPal, DIBDither)
                '压缩存储
                Call mGdIpEx.SaveDIB(oDIB, strFileName, [ImageJPEG], 100)          '90%的JPEG图片压缩质量
            End If
            'Unload the GDI+ Dll
            Call mGdIpEx.GdiplusShutdown(m_GDIpToken)
            gstrSQL = "select 对象ID from 病历范文图形 where 对象ID=[1]"
            Call zlBlobSql(4, lngID, strPic, ArraySQL)
            oStream.Close
        Next
    End If
    '后期处理
     ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
     gstrSQL = "zl_病历范文内容_commit(" & lngFileID & ")"
     ArraySQL(UBound(ArraySQL)) = gstrSQL
    '=========================================================================================
    '保存RTFText的Sql
    '=========================================================================================
    If Not oDemoNode.selectSingleNode("Content") Is Nothing Then
        Set ContentNode = oDemoNode.selectSingleNode("Content")
        Me.RTbContext.TextRTF = ContentNode.selectSingleNode("RTF").Text
        If gobjFSO.FileExists(App.Path & "\TMP.rtf") Then gobjFSO.DeleteFile App.Path & "\TMP.rtf", True    '保存为临时文件
        Me.RTbContext.SaveFile App.Path & "\TMP.rtf"
        strTemp = zlFileZip(App.Path & "\TMP.rtf")
        If gobjFSO.FileExists(App.Path & "\TMP.rtf") Then gobjFSO.DeleteFile App.Path & "\TMP.rtf", True
        If gobjFSO.FileExists(strTemp) Then
            Call zlBlobSql(3, lngFileID, strTemp, ArraySQL)
            gobjFSO.DeleteFile strTemp, True      '删除临时文件
        End If
    End If
bb: If Not BeginTrans(ArraySQL) Then gcnOracle.RollbackTrans: Err.Clear: GoTo errHand
    ImportDemosFromXML = True: Exit Function
errHand:
    ImportDemosFromXML = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
'启动事务执行SQL
Private Function BeginTrans(ByVal ArraySQL As Variant) As Boolean
    On Error GoTo errHand
    Dim i As Long
    gcnOracle.BeginTrans
    For i = 1 To UBound(ArraySQL)
        gstrSQL = ArraySQL(i)
        If Trim(gstrSQL) <> "" Then
            Call zlDatabase.ExecuteProcedure(gstrSQL, "cEPRCompends")
        End If
    Next
    gcnOracle.CommitTrans
    BeginTrans = True
    Exit Function
errHand:
    BeginTrans = False
End Function
'################################################################################################################
'## 功能：  检查诊治要素的原始定义是否存在（用于XML导入时的验证）
'################################################################################################################
Public Function CheckValid(ByVal ID As Long, ByVal Name As String) As Boolean
    Dim rs As New Recordset
    gstrSQL = "Select ID From 诊治所见项目 Where ID = [1] And 中文名 = [2]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, ID, Name)
    If rs.EOF Then
        CheckValid = False
    Else
        CheckValid = (rs!ID > 0)
    End If
End Function
