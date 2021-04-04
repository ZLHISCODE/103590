VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFileExportOrImport 
   Caption         =   "病历文件导出列表"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8775
   Icon            =   "frmFileExportOrImport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   8775
   StartUpPosition =   1  '所有者中心
   Begin zlRichEditor.Editor Editor1 
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   3960
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex8Ctl.VSFlexGrid vsgrid 
      Height          =   1860
      Left            =   720
      TabIndex        =   0
      Top             =   720
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
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
   Begin VB.PictureBox PicBtn 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   8775
      TabIndex        =   1
      Top             =   5160
      Width           =   8775
      Begin MSComctlLib.ProgressBar progBar 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   5520
      Top             =   1080
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
            Picture         =   "frmFileExportOrImport.frx":6852
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileExportOrImport.frx":6DEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileExportOrImport.frx":7386
            Key             =   "签名"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox RTbFootText 
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      _Version        =   393217
      TextRTF         =   $"frmFileExportOrImport.frx":76D8
   End
   Begin RichTextLib.RichTextBox RTbHeadText 
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      _Version        =   393217
      TextRTF         =   $"frmFileExportOrImport.frx":7775
   End
   Begin RichTextLib.RichTextBox RTbContext 
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      _Version        =   393217
      TextRTF         =   $"frmFileExportOrImport.frx":7812
   End
   Begin XtremeCommandBars.ImageManager imgManager 
      Left            =   6360
      Top             =   1200
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmFileExportOrImport.frx":78AF
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
Attribute VB_Name = "frmFileExportOrImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Private mdoc As DOMDocument         'Xml文档
Private mIntClosed As Integer       '控制窗体是否可以关闭
Private mstrPath As String          '路径
Private mblnInit As Boolean         '工具栏Enable状态
Private mlngType As Long            '当前窗体处于导出/导入状态（1为导出，2为导入）
Private Type mDocType
    mDocXML As New DOMDocument
    mXmlPath As String
End Type
Private Enum mExportCols
    Range = 0: Choose: cType: cID: cNull: cName
End Enum
Private Enum mImportCols
    Choose = 0: cName: cImportType: cTip: cUnit: cPath
End Enum
Private mDocArr() As mDocType
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
Public Function ShowMe(ByVal objParent As Object, ByVal lngType As Long)
    If lngType = 1 Then
        Call ExportList
        Me.Caption = "病历文件导出列表"
        Me.vsgrid.Tag = "Export"
        Me.Tag = ""
        Me.ShowControl(Me.cbsThis, 11, True).Visible = False
        Me.Show 1, objParent
    Else
        If ImportList Then
        Me.Caption = "病历文件导入列表"
        Me.vsgrid.Tag = "Import"
        Me.ShowControl(Me.cbsThis, 10, True).Visible = False
        Me.Show 1, objParent
        End If
    End If
End Function
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case menu_Cover
             If vsgrid.Row < 1 Then Exit Sub
             vsgrid.Cell(flexcpForeColor, vsgrid.Row, 0, vsgrid.Row, 3) = vbMagenta
             vsgrid.TextMatrix(vsgrid.Row, 3) = "已存在,导入将覆盖原有文件！"
             vsgrid.TextMatrix(vsgrid.Row, 2) = Split(vsgrid.TextMatrix(vsgrid.Row, 2), "_")(0) & "_1"
             vsgrid.Cell(flexcpData, vsgrid.Row, 3) = "存在"
             vsgrid_DblClick
        Case menu_Add
             If vsgrid.Row < 1 Then Exit Sub
             vsgrid.Cell(flexcpForeColor, vsgrid.Row, 0, vsgrid.Row, 3) = vbBlack
             vsgrid.TextMatrix(vsgrid.Row, 3) = "已存在,导入将新增此文件！"
             vsgrid.TextMatrix(vsgrid.Row, 2) = Split(vsgrid.TextMatrix(vsgrid.Row, 2), "_")(0) & "_2"
             vsgrid.Cell(flexcpData, vsgrid.Row, 3) = ""
             If vsgrid.Cell(flexcpPicture, vsgrid.Row, 0) Is Nothing Then vsgrid_DblClick
        Case menu_RemoveRow
             vsgrid.RemoveItem (vsgrid.Row)
             If vsgrid.Rows = 1 Then InitVsGrid ("请先添加需要导入的病历文件 ！")
             Me.Tag = Val(Me.Tag) - 1
        Case menu_Clear
            Call InitVsGrid("请先添加需要导入的病历文件 ！")
        Case menu_Export
            Call Export
        Case menu_ExportOne
             If vsgrid.Cols > 1 Then
                Control.Checked = True
                ShowControl(Me.cbsThis, 102, True).Checked = False
             End If
        Case menu_ExportMore
             If vsgrid.Cols > 1 Then
                Control.Checked = True
                ShowControl(Me.cbsThis, 101, True).Checked = False
             End If
        Case menu_Import
            Call Import
        Case menu_IcheckAll, menu_EcheckAll
            Call CheckItems(True)
        Case menu_CheckHave
            Call CheckHave(True)
        Case menu_ClearHave
            Call CheckHave(False)
        Case menu_IclearAll, menu_EclearAll
            Call CheckItems(False)
        Case menu_CheckThis
            Call vsgridClick
        Case menu_AddFile
            Call ImportList
            Me.vsgrid.Tag = "Import"
        Case menu_Unload
            Unload Me
         Exit Sub
    End Select
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If mblnInit Then
        Control.Enabled = False
    Else
       Select Case Control.ID
           Case menu_Cover
                Control.Enabled = IIf(vsgrid.TextMatrix(vsgrid.Row, 3) = "", False, True)
                Control.Checked = IIf(vsgrid.TextMatrix(vsgrid.Row, 3) = "已存在,导入将覆盖原有文件！", True, False)
           Case menu_Add
                Control.Enabled = IIf(vsgrid.TextMatrix(vsgrid.Row, 3) = "", False, True)
                Control.Checked = IIf(vsgrid.TextMatrix(vsgrid.Row, 3) = "已存在,导入将新增此文件！", True, False)
           Case menu_Export, menu_Import
                Control.Enabled = IIf(vsgrid.Tag = "" Or vsgrid.Cols = 1, False, True)
           Case menu_ExportOne, menu_ExportMore
                Control.Enabled = IIf(vsgrid.Tag = "Import", False, True)
           Case menu_IcheckAll, menu_IclearAll, menu_ImportOption
                Control.Enabled = IIf(vsgrid.Rows = 1, False, True)
                Control.Visible = IIf(vsgrid.Tag = "Export", False, True)
           Case menu_CheckThis
                Control.Visible = False
           Case menu_AddFile
                Control.Visible = IIf(vsgrid.Tag = "Export", False, True)
           Case menu_EcheckAll, menu_EclearAll
                Control.Visible = IIf(vsgrid.Tag = "Import", False, True)
                Control.Enabled = IIf(vsgrid.Rows = 1, False, True)
        End Select
    End If
End Sub
Private Sub CheckItems(ByVal blnOn As Boolean)
    Dim i As Integer, intCol As Integer, j As Integer
    intCol = IIf(vsgrid.Tag = "Import", 0, 1)
    For i = 1 To vsgrid.Rows - 1
         If intCol = 0 Then
            vsgrid.Cell(flexcpPicture, i, 0) = IIf(blnOn And Not IsHave(i), img16.ListImages("Check").Picture, Nothing)
            vsgrid.Cell(flexcpData, i, 0) = IIf(blnOn And Not IsHave(i), 1, 0)
         Else
            vsgrid.Cell(flexcpPicture, i, 1) = IIf(blnOn, img16.ListImages("Check").Picture, Nothing)
            vsgrid.Cell(flexcpData, i, 1) = IIf(blnOn, 1, 0)
         End If
    Next i
End Sub
Private Sub Form_Load()
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim blnChecked As Boolean
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
        Set objPopup = .Add(xtpControlSplitButtonPopup, 10, "导出"): objPopup.STYLE = xtpButtonIconAndCaption
            objPopup.BeginGroup = True
            objPopup.ID = 10                    'Popup的ID需重新赋值才能生效
            objPopup.IconId = 10                'Popup的IconId需重新赋值才能生效
        objPopup.CommandBar.Width = 100
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, 101, "导出为一个XML文件(&Q)"
            .Add xtpControlButton, 102, "导出为多个XML文件(&W)"
        End With
        Set objControl = .Add(xtpControlButton, 11, "导入"): objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, 15, "添加"): objControl.STYLE = xtpButtonIconAndCaption
        Set objPopup = .Add(xtpControlButtonPopup, 110, "导入选项"): objPopup.STYLE = xtpButtonIconAndCaption
            objPopup.BeginGroup = True
            objPopup.ID = 110
            objPopup.IconId = 110
        objPopup.CommandBar.Width = 60
        With objPopup.CommandBar.Controls
        .Add xtpControlButton, 1, "导入将覆盖原有文件"
        .Add xtpControlButton, 2, "导入将新增该文件"
        End With
        Set objControl = .Add(xtpControlButton, 17, "全选"): objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, 18, "全清"): objControl.STYLE = xtpButtonIconAndCaption
        Set objPopup = .Add(xtpControlSplitButtonPopup, 12, "全选"): objPopup.STYLE = xtpButtonIconAndCaption
            objPopup.BeginGroup = True
            objPopup.ID = 12
            objPopup.IconId = 12
        objPopup.CommandBar.Width = 60
        With objPopup.CommandBar.Controls
        .Add xtpControlButton, 121, "全选已存在文件"
        End With
        Set objPopup = .Add(xtpControlSplitButtonPopup, 13, "全清"): objPopup.STYLE = xtpButtonIconAndCaption
            objPopup.BeginGroup = True
            objPopup.ID = 13
            objPopup.IconId = 13
        objPopup.CommandBar.Width = 60
        With objPopup.CommandBar.Controls
        .Add xtpControlButton, 131, "全清已存在文件"
        End With
        Set objControl = .Add(xtpControlButton, 14, "选择"): objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, 16, "退出"): objControl.STYLE = xtpButtonIconAndCaption
        objControl.STYLE = xtpButtonIconAndCaption  '同时显示图标和文字
    End With
    blnChecked = IIf(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ExportType", App.Path) = "More", False, True)
    Me.cbsThis.ActiveMenuBar.Visible = False
    Me.ShowControl(Me.cbsThis, 101, False).Checked = blnChecked
    Me.ShowControl(Me.cbsThis, 102, False).Checked = Not blnChecked
End Sub
Private Sub Export()
    Dim i As Integer, lngArrFile As Variant, strItems As String
    '得到所有选中项
     For i = 1 To vsgrid.Rows - 1
         If Not vsgrid.Cell(flexcpPicture, i, mExportCols.Choose) Is Nothing And vsgrid.GetNode(i).Children < 1 Then
             strItems = strItems & vsgrid.TextMatrix(i, mExportCols.cID) & "_" & vsgrid.TextMatrix(i, mExportCols.cName) & "_" & vsgrid.TextMatrix(i, mExportCols.cType) & ","
         End If
     Next i
     If strItems = "" Then
         MsgBox "请选择需要导出的文件！", vbInformation, gstrSysName
         Exit Sub
     End If
     strItems = Mid(strItems, 1, Len(strItems) - 1)
     '指定保存的文件路径
     mstrPath = zl9ComLib.OS.OpenDir(Me.hWnd, "指定导出目录")
     If mstrPath = "" Then Exit Sub
     On Error Resume Next
     mstrPath = mstrPath & "\" & zl9ComLib.GetUnitName
     gobjFSO.CreateFolder (mstrPath)
     If Err.Number = 32755 Then Err.Clear: Exit Sub
     On Error GoTo errHand
     lngArrFile = Split(strItems, ",")
     mIntClosed = 1: EnableControlBar Me, False: mblnInit = True
     Call StartExportToXMLFile(cprEM_修改, cprET_病历文件定义, lngArrFile, IIf(ShowControl(Me.cbsThis, 101, True).Checked, 1, 2))
     MsgBox "数据已导出到目标文件：" & mstrPath, vbApplicationModal + vbInformation, "提醒"
     mIntClosed = 0: mblnInit = False
     Unload Me
     Exit Sub
errHand:
    mIntClosed = 0: mblnInit = False
    EnableControlBar Me, True
    If ErrCenter() = 1 Then
         Resume
    End If
    Call SaveErrLog
End Sub
Private Sub Import()
    Dim i As Long, j As Long
    On Error GoTo errHand
    '开始循环导入
    '---------------
    Me.progBar.Visible = True
    mIntClosed = 1: EnableControlBar Me, False: mblnInit = True
    ReDim Preserve mDocArr(1 To 1) As mDocType
    For i = 1 To vsgrid.Rows - 1
        Set mdoc = Nothing
        If Not vsgrid.Cell(flexcpPicture, i, mImportCols.Choose) Is Nothing Then
            If UBound(mDocArr) > 1 Then
                For j = 1 To UBound(mDocArr)
                    If vsgrid.TextMatrix(i, mImportCols.cPath) = mDocArr(j).mXmlPath Then Set mdoc = mDocArr(j).mDocXML: Exit For
                Next j
            End If
            If mdoc Is Nothing Then
                ReDim Preserve mDocArr(1 To UBound(mDocArr) + 1) As mDocType
                mDocArr(UBound(mDocArr)).mDocXML.Load vsgrid.TextMatrix(i, mImportCols.cPath)
                Set mdoc = mDocArr(UBound(mDocArr)).mDocXML
                mDocArr(UBound(mDocArr)).mXmlPath = vsgrid.TextMatrix(i, mImportCols.cPath)
            End If
            DoEvents
            Me.Refresh
            Call ImportFromXml(vsgrid.TextMatrix(i, mImportCols.cPath), vsgrid.TextMatrix(i, mImportCols.cImportType))
        End If
        progBar.Value = IIf(progBar.Value + progBar.Max / (vsgrid.Rows - 1) > progBar.Max, progBar.Max, progBar.Value + progBar.Max / (vsgrid.Rows - 1))
    Next i
    '----------------
    MsgBox "导入完成 ！", vbInformation, gstrSysName
    mIntClosed = 0: mblnInit = False
    Unload Me
    Exit Sub
errHand:
    mIntClosed = 0: mblnInit = False
    EnableControlBar Me, True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
'加载病历文件列表
Public Sub ExportList()
    Dim strType As String, i As Long, j As Long, k As Long
    Dim rsTemp As ADODB.Recordset
    gstrSQL = "select distinct ID,种类,编号,名称 from (" & _
              "  Select l.Id, decode(l.种类,1,'1-门诊病历',2,'2-住院病历',4,'4-护理病历',5,'5-疾病证明报告',6,'6-知情文件') as 种类, l.编号, l.名称" & _
              "  From 病历文件列表 l where l.保留<>2 and l.种类 in (1,2,4,5,6)" & _
              "  Union All Select 0, decode(l.种类,1,'1-门诊病历',2,'2-住院病历',4,'4-护理病历',5,'5-疾病证明报告',6,'6-知情文件') as 种类, null, null" & _
              "  From 病历文件列表 l where l.保留<>2 and l.种类 in (1,2,4,5,6)) order by 种类,ID"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        'vsgridClick
        With vsgrid
            '基础设置
            If rsTemp.RecordCount < 1 Then
                 Call InitVsGrid("暂时没有可以导出的病历文件 ！"):  rsTemp.Close
                 Exit Sub
            End If
            .Clear: .Cols = 6: .Rows = 1: .FixedRows = 1: .ROWHEIGHT(0) = 10
            .ColWidth(mExportCols.Range) = 400: .ColWidth(mExportCols.Choose) = 270: .ColWidth(mExportCols.cType) = 0: .ColWidth(mExportCols.cNull) = 0: .ColWidth(mExportCols.cID) = 0: .ColWidth(mExportCols.cName) = 2500
            .Cell(flexcpData, 0, mExportCols.Choose) = 1: .TextMatrix(0, mExportCols.cName) = "名称": .TextMatrix(0, mExportCols.Choose) = "选择": .ColAlignment(mExportCols.cName) = flexAlignLeftCenter
            '设置分组
            .OutlineCol = 0: .OutlineBar = flexOutlineBarCompleteLeaf
            For i = 1 To rsTemp.RecordCount
                If strType <> rsTemp!种类 Then
                    .AddItem ""
                    For k = 2 To .Cols - 1
                        .TextMatrix(i, k) = NVL(rsTemp("种类").Value)
                    Next k
                    .Cell(flexcpBackColor, i, 0, i, 5) = &HFFC0C0
                    .IsSubtotal(i) = True
                    .Cell(flexcpData, i, mExportCols.Choose) = 1
                    .MergeCells = flexMergeFree
                    .MergeRow(1) = True '是否左右行合并
                    strType = rsTemp!种类
                Else
                    .AddItem ""
                    .Cell(flexcpData, i, mExportCols.Choose) = 0
                    .TextMatrix(i, mExportCols.cType) = NVL(rsTemp("种类").Value)
                    .TextMatrix(i, mExportCols.cID) = NVL(rsTemp("ID").Value)
                    .TextMatrix(i, mExportCols.cName) = NVL(rsTemp("名称").Value)
                    .IsSubtotal(i) = True
                    .RowOutlineLevel(i) = 1
                End If
                .Cell(flexcpPicture, i, mExportCols.Choose) = img16.ListImages("Check").Picture
                .Cell(flexcpData, i, mExportCols.Choose) = 1
                rsTemp.MoveNext
           Next i
           For i = 1 To vsgrid.Rows - 1
                If .IsSubtotal(i) = True Then
                    .GetNode(i).Expanded = True
                End If
           Next i
           If vsgrid.Rows > 1 Then vsgrid.Row = 2
    End With
    rsTemp.Close
End Sub
'选中已存在文件
Private Sub CheckHave(ByVal blnOn As Boolean)
    Dim i As Integer
    For i = 1 To vsgrid.Rows - 1
       If vsgrid.Cell(flexcpForeColor, i, 1, i, 3) = vbMagenta Then
          vsgrid.Cell(flexcpPicture, i, 0) = IIf(blnOn And Not IsHave(i), img16.ListImages("Check").Picture, Nothing)
       End If
    Next i
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Cancel = mIntClosed
End Sub

Private Sub Form_Resize()
    Me.vsgrid.Move 0, 500, Me.ScaleWidth, Me.ScaleHeight - Me.PicBtn.Height - 500
    Me.PicBtn.Move 0, vsgrid.Height + 500, Me.ScaleWidth, PicBtn.Height
    Me.progBar.Move 0, 60, PicBtn.Width, progBar.Height
    If vsgrid.Rows = 1 Or vsgrid.Cols = 1 Then
        vsgrid.ROWHEIGHT(0) = Me.ScaleHeight
    End If
End Sub

'单击VsGrid
Private Sub vsgridClick()
    Dim i As Long, j As Long, strItems As String
    Dim intX As Integer, intW As Integer
    If vsgrid.Cols = 1 Then Exit Sub
    If vsgrid.Tag = "Import" Then
       If vsgrid.Row > 0 Then
         If IsHave(vsgrid.Row) And vsgrid.Cell(flexcpData, vsgrid.Row, 0) = 0 And vsgrid.Cell(flexcpData, vsgrid.Row, 3) = "存在" Then
            MsgBox "不能同时选择两个名称相同的病历文件覆盖原有文件！", vbInformation, gstrSysName: Exit Sub
         End If
         vsgrid.Cell(flexcpData, vsgrid.Row, 0) = IIf(vsgrid.Cell(flexcpData, vsgrid.Row, 0) = 1, 0, 1)
         vsgrid.Cell(flexcpPicture, vsgrid.Row, 0) = IIf(vsgrid.Cell(flexcpData, vsgrid.Row, 0) = 1, img16.ListImages("Check").Picture, Nothing)
       End If
    Else
        '选中节点下所有子项
        If vsgrid.MouseRow < 0 Then Exit Sub
        vsgrid.Cell(flexcpData, vsgrid.Row, 1) = IIf(vsgrid.Cell(flexcpData, vsgrid.Row, 1) = 1, 0, 1)
        For i = vsgrid.Row To vsgrid.GetNode(vsgrid.Row).Children + vsgrid.Row
             vsgrid.Cell(flexcpData, vsgrid.Row, 0) = IIf(vsgrid.Cell(flexcpData, vsgrid.Row, 0) = 1, 0, 1)
             vsgrid.Cell(flexcpPicture, i, 1) = IIf(vsgrid.Cell(flexcpData, vsgrid.Row, 1) = 1, img16.ListImages("Check").Picture, Nothing)
        Next i
    End If
End Sub
'
Private Function IsHave(ByVal intRow As Long) As Boolean
    Dim i As Long
    For i = 1 To vsgrid.Rows - 1
        If intRow <> i And vsgrid.Cell(flexcpData, i, 3) = "存在" And Not vsgrid.Cell(flexcpPicture, i, 0) Is Nothing Then
            If vsgrid.TextMatrix(intRow, 1) = vsgrid.TextMatrix(i, 1) And vsgrid.Cell(flexcpForeColor, intRow, 1, intRow, 3) = vbMagenta Then
            IsHave = True: Exit Function
            End If
        End If
    Next i
    IsHave = False
End Function

Private Sub Form_Unload(Cancel As Integer)
  Dim ExportType As String
  ExportType = IIf(ShowControl(Me.cbsThis, 101, True).Checked, "One", "More")
  SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ExportType", ExportType
  If Not mdoc Is Nothing Then Set mdoc = Nothing
End Sub

Private Sub vsgrid_Click()
    If vsgrid.Cols = 1 Then Exit Sub
    If vsgrid.MouseIcon Is Nothing Then Exit Sub
    If vsgrid.MouseIcon = Me.img16.ListImages(1).Picture Then
            vsgridClick
    End If
End Sub

'双击事件
Private Sub vsgrid_DblClick()
     If vsgrid.Cols = 1 Then Exit Sub
     If vsgrid.MouseIcon Is Nothing And vsgrid.MouseRow > 1 Then
        If vsgrid.Tag = "Export" Then
            If vsgrid.GetNode(vsgrid.Row).Children > 1 Then
                vsgrid.GetNode(vsgrid.Row).Expanded = Not vsgrid.GetNode(vsgrid.Row).Expanded: Exit Sub
            End If
        End If
        vsgridClick
        Exit Sub
     End If
End Sub
'按下键盘事件
Private Sub vsgrid_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsgrid
        If .IsSubtotal(.Row) Then
            Select Case KeyCode
              Case vbKeyLeft
                  .GetNode(.Row).Expanded = False
              Case vbKeySpace
                   vsgridClick
              Case vbKeyRight
                  .GetNode(.Row).Expanded = True
              Case 13
                .GetNode(.Row).Expanded = Not .GetNode(.Row).Expanded
              Case vbKeyA
                If Shift = 2 Then CheckItems (True)
              Case vbKeyZ
                If Shift = 2 Then CheckItems (False)
            End Select
        ElseIf vsgrid.Tag = "Import" Then
            If KeyCode = 13 Then
              Call vsgridClick
            ElseIf KeyCode = vbKeySpace Then
                vsgridClick
            ElseIf KeyCode = vbKeyA Then
              If Shift = 2 Then CheckItems (True)
            ElseIf KeyCode = vbKeyZ Then
              If Shift = 2 Then CheckItems (False)
            End If
        End If
   End With
End Sub

Private Sub vsgrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngH As Long, lngY As Long
     If vsgrid.Cols = 1 Or vsgrid.Rows = 1 Then Exit Sub
     lngH = vsgrid.Row * 255: lngY = 255 * (vsgrid.Row + 1)
     If Button = 2 Then
        If vsgrid.Tag = "Import" And Y > lngH And Y < lngY Then
                Dim Popup As CommandBar
                Dim objControl As CommandBarControl
                Set Popup = cbsThis.Add("Popup", xtpBarPopup)
                With Popup.Controls
                    .Add xtpControlButton, 1, "导入时覆盖该病历(&F)"
                    .Add xtpControlButton, 2, "导入时新增该病历(&A)"
                    .Add xtpControlButton, 3, "从列表中移除(&D)"
                    .Add xtpControlButton, 4, "清空列表(&C)"
                End With
                Popup.ShowPopup
        End If
      End If
End Sub

Private Sub vsgrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intX As Integer, intW As Integer
    If vsgrid.Cols = 1 Then Exit Sub
    If vsgrid.Tag = "Import" Then intX = 0: intW = CSng(vsgrid.ColWidth(0)) Else intX = CSng(vsgrid.ColWidth(0)): intW = CSng(vsgrid.ColWidth(0) + vsgrid.ColWidth(1))
    If X > intX And X < intW And Y > 255 And Y < CSng(vsgrid.Rows * 255) And vsgrid.MouseRow > -1 Then
         vsgrid.MousePointer = flexCustom
         Set vsgrid.MouseIcon = Me.img16.ListImages(1).Picture
    Else
         vsgrid.MousePointer = flexDefault
         Set vsgrid.MouseIcon = Nothing
    End If
End Sub
'################################################################################################################
'## 功能：  将病历文件从XML文件中导出到病历文件列表中
'##
'##
'## 返回：  保存成功，返回Ture；否则返回False。
'################################################################################################################
Public Function ImportList() As Boolean
    '从XML文件导入
    Dim strXML As String, strArrXml As Variant, strTempName As String
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    Dim oFileRootList As IXMLDOMNodeList      '文件节点
    Dim oFileRoot As IXMLDOMElement, strTemp As String
    Dim oRoot  As IXMLDOMElement        '根节点
    Dim rsTemp As ADODB.Recordset
    Dim strUnitName As String
    Static intRow As Long
    On Error Resume Next
    dlgThis.MaxFileSize = 32767
    dlgThis.Filter = "*.XML|*.xml"
    dlgThis.DialogTitle = "打开(支持多选)"
    dlgThis.CancelError = True
    dlgThis.flags = &H10& Or &H200& Or &H80000
    dlgThis.ShowOpen

'    dlgThis.Action = 1
    
    If Err.Number = 32755 Then Err.Clear: ImportList = False: Exit Function
    If dlgThis.FileTitle = "" Then
        strTempName = Split(dlgThis.Filename, "\")(UBound(Split(dlgThis.Filename, "\")))
        strArrXml = Split(Trim(strTempName), Chr(0)) '处理文件名称字符串
        strArrXml(0) = Replace(dlgThis.Filename, strTempName, strArrXml(0) & "\")
    Else
       strTempName = Split(dlgThis.Filename, "\")(UBound(Split(dlgThis.Filename, "\")))
       strTempName = Replace(dlgThis.Filename, strTempName, "") & "," & Split(dlgThis.Filename, "\")(UBound(Split(dlgThis.Filename, "\")))
       strArrXml = Split(strTempName, ",")
    End If
    With vsgrid
        If vsgrid.Tag = "Export" Or vsgrid.Tag = "" Then
            .Clear
            .FixedRows = 1: .ExplorerBar = flexExSortShow
            .Cols = 6: .ColWidth(mImportCols.Choose) = 270: .Rows = 1: .ColAlignment(mImportCols.cName) = flexAlignLeftCenter
            .ColWidth(mImportCols.cImportType) = 0: .ColWidth(mImportCols.cName) = 1500: .ROWHEIGHT(mImportCols.Choose) = 50: .ColWidth(mImportCols.cTip) = 2500: .ColWidth(mImportCols.cUnit) = 2500: .ColWidth(mImportCols.cPath) = 6000
            .TextMatrix(0, mImportCols.Choose) = "选择": .TextMatrix(0, mImportCols.cName) = "名称": .TextMatrix(0, mImportCols.cTip) = "提示": .TextMatrix(0, mImportCols.cUnit) = "导出单位": .TextMatrix(0, mImportCols.cPath) = "文件位置"
            intRow = 0
        End If
    End With
    For k = 1 To UBound(strArrXml)
        strXML = strArrXml(0) & strArrXml(k)
        Set mdoc = New DOMDocument
        mdoc.Load strXML
        '如果该路径下文件已被加载则不再加载
        For l = 1 To vsgrid.Rows - 1
            If strXML = Trim(vsgrid.TextMatrix(l, mImportCols.cPath)) Then
                 MsgBox strArrXml(k) & ",已经被打开，请勿重复打开 ！", vbInformation, gstrSysName
                 GoTo a
            End If
        Next l
        '如果不包含任何元素，则退出
        If mdoc.documentElement Is Nothing Then
           MsgBox "你选择的XML文件不是该软件导出的正确XML格式的文件!", vbInformation, gstrSysName: Exit Function
        End If
        '读取文件结构
        Set oRoot = mdoc.selectSingleNode("Document")       'oRoot置为根节点
        Set oFileRootList = oRoot.selectNodes("File")
        If oRoot Is Nothing Then
            MsgBox "你选择的XML文件不是该软件导出的正确XML格式的文件!", vbInformation, gstrSysName: Exit Function
        ElseIf Not oRoot.selectSingleNode("EPRFileInfo") Is Nothing Then
             strTemp = strTemp & oRoot.selectSingleNode("EPRFileInfo").selectSingleNode("ID").Text & "_" & oRoot.selectSingleNode("EPRFileInfo").selectSingleNode("名称").Text & "_" & oRoot.selectSingleNode("EPRFileInfo").selectSingleNode("种类").Text & ","
        ElseIf oFileRootList.Item(0) Is Nothing Then
            MsgBox "你选择的XML文件数据可能为空!", vbInformation, gstrSysName: Exit Function
        Else
            For Each oFileRoot In oFileRootList
                strTemp = strTemp & GetNodeValue(oFileRoot, "ID", 0) & "_" & GetNodeValue(oFileRoot, "名称", 0) & "_" & GetNodeValue(oFileRoot, "种类", 0) & "_" & oRoot.getAttribute("UnitName") & ","
            Next
        End If
        strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
        gstrSQL = "  select distinct ID,种类,编号,名称 from (" & _
                  "  Select l.Id, decode(l.种类,1,'1-门诊病历',2,'2-住院病历',4,'4-护理病历',5,'5-疾病证明报告',6,'6-知情文件') as 种类, l.编号, l.名称" & _
                  "  From 病历文件列表 l where l.保留<>2 and l.种类 in (1,2,4,5,6)" & _
                  "  Union All Select 0, decode(l.种类,1,'1-门诊病历',2,'2-住院病历',4,'4-护理病历',5,'5-疾病证明报告',6,'6-知情文件') as 种类, null, null" & _
                  "  From 病历文件列表 l where l.保留<>2 and l.种类 in (1,2,4,5,6)) order by 种类,ID"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        With vsgrid
            For i = 1 To UBound(Split(strTemp, ",")) + 1
                   Me.Tag = Val(Me.Tag) + 1
                   intRow = Me.Tag
                  .AddItem ""
                  .TextMatrix(intRow, mImportCols.cName) = Split(Split(strTemp, ",")(i - 1), "_")(1)
                  .TextMatrix(intRow, mImportCols.cImportType) = Split(Split(strTemp, ",")(i - 1), "_")(1) & "_2"
                  .TextMatrix(intRow, mImportCols.cTip) = "不存在,导入将新增该文件！"
                  .TextMatrix(intRow, mImportCols.cPath) = strXML
                  .Cell(flexcpData, intRow, mImportCols.cName) = Split(Split(strTemp, ",")(i - 1), "_")(0)
                  .TextMatrix(intRow, mImportCols.cUnit) = Split(Split(strTemp, ",")(i - 1), "_")(3)
                  If Not rsTemp Is Nothing Then
                      Do While Not rsTemp.EOF
                          If Trim(NVL(rsTemp!名称, "")) = Trim(Split(Split(strTemp, ",")(i - 1), "_")(1)) And Val(NVL(rsTemp!种类, "")) = Split(Split(strTemp, ",")(i - 1), "_")(2) Then
                              .Cell(flexcpForeColor, intRow, mImportCols.cName, intRow, mImportCols.cTip) = vbMagenta
                              .Cell(flexcpData, intRow, mImportCols.cTip) = "存在"
                              .TextMatrix(intRow, mImportCols.cImportType) = Split(Split(strTemp, ",")(i - 1), "_")(1) & "_1"
                              .TextMatrix(intRow, mImportCols.cTip) = "已存在,导入将覆盖原有文件！"
                              j = 1
                          End If
                          rsTemp.MoveNext
                      Loop
                  End If
                  '判断文件列表中是否有相同名称文件，相同则不选择
                  If Not IsHave(intRow) Then
                    .Cell(flexcpPicture, intRow, mImportCols.Choose) = img16.ListImages("Check").Picture
                    .Cell(flexcpData, i, mImportCols.Choose) = 1
                  End If
                  vsgrid.Cell(flexcpData, intRow, mImportCols.Choose) = 1
                  rsTemp.MoveFirst
            Next i
        End With
        strTemp = ""
        Me.Tag = vsgrid.Rows - 1
        
a:
    Next k
    If j = 1 Then Me.ShowControl(Me.cbsThis, 110, True).Visible = True
    If vsgrid.Rows > 1 Then vsgrid.Row = 1
    dlgThis.Filename = ""
    ImportList = True
    Exit Function
errHand:
    ImportList = False
     If ErrCenter() = 1 Then Resume
       Call SaveErrLog
End Function
'################################################################################################################
'## 功能：  开始将病历文件导出到XML文档中
'##
'## 参数：  eEdtMode    :当前编辑模式（新增、修改）
'##         eEdtType    :当前编辑方式（文件定义、示范编辑、单病历编辑、单病历审核）
'##         lngArrFile  :当前选择项的ID、名称集合
'##       lngExportType :导出类型(1,表示导出单个文件，2表示导出多个文件)
'## 返回：  保存成功，返回Ture；否则返回False。
'################################################################################################################
Public Function StartExportToXMLFile(ByVal eEdtMode As EditModeEnum, ByVal eEdtType As EditTypeEnum, ByVal lngArrFile, ByVal lngExportType As Long) As Boolean
    Dim i As Long, lngFileID As Long, strFileName As String, strFileType As String
    Dim cDoc As New DOMDocument              'xml文档
    Dim Result As VbMsgBoxResult
    Dim pi As IXMLDOMProcessingInstruction  '版本信息
    Dim oRootNew As IXMLDOMElement
    Dim oRoot As IXMLDOMElement         '根节点
    '------------------------------------------------
    
    If gobjFSO.FileExists(dlgThis.Filename) Then
        DoEvents
        If MsgBox(dlgThis.Filename & "文件已经存在，是否覆盖？", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then Exit Function
    End If
    cDoc.appendChild cDoc.createComment(gstrSysName & "操作员:" & gstrUserName & "，部门:" & gstrDeptName & "，时间:" & Format(Now(), "YYYY年MM月DD日"))
    Me.progBar.Visible = True
    For i = 0 To UBound(lngArrFile)
        DoEvents
        Me.Refresh
        strFileName = Split(lngArrFile(i), "_")(1)
        lngFileID = Split(lngArrFile(i), "_")(0)
        strFileType = Split(Split(lngArrFile(i), "_")(2), "-")(0)
        '普通住院病历
        ZLCommFun.ShowFlash "正在导出文件，请稍候..."
        Screen.MousePointer = vbHourglass
        If lngExportType = 1 Then
             Call ExportToXml(cprEM_修改, cprET_病历文件定义, lngFileID, cDoc, oRoot)
        ElseIf lngExportType = 2 Then
             Set cDoc = New DOMDocument
             cDoc.appendChild cDoc.createComment(gstrSysName & "操作员:" & gstrUserName & "，部门:" & gstrDeptName & "，时间:" & Format(Now(), "YYYY年MM月DD日"))
             Set oRootNew = Nothing
             Call ExportToXml(cprEM_修改, cprET_病历文件定义, lngFileID, cDoc, oRootNew)
             Set pi = cDoc.createProcessingInstruction("xml", "version='1.0' encoding='gb2312'")
             Call cDoc.insertBefore(pi, cDoc.childNodes(0))
             cDoc.Save mstrPath & "/" & "定义_" & strFileName & ".XML"
        End If
        Me.progBar.Value = IIf(Me.progBar.Value + Me.progBar.Max / (UBound(lngArrFile) + 1) > progBar.Max, progBar.Max, Me.progBar.Value + Me.progBar.Max / (UBound(lngArrFile) + 1))
    Next i
    If lngExportType <> 2 Then
        Set pi = cDoc.createProcessingInstruction("xml", "version='1.0' encoding='gb2312'")
        Call cDoc.insertBefore(pi, cDoc.childNodes(0))
        cDoc.Save mstrPath & "/" & zl9ComLib.GetUnitName & "_病历文件列表.XML"
        Set cDoc = Nothing
    End If
    Screen.MousePointer = vbDefault
    Me.progBar.Value = Me.progBar.Max
    Me.progBar.Visible = False
    Me.progBar.Value = 0
End Function

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
    Me.Tag = ""
    vsgrid.ROWHEIGHT(0) = Me.ScaleHeight
End Function
'------------------------------------------------
   '功能： 设置CommandBars菜单及工具栏的显示状态
   '参数： CommandBars控件，toolId ,bolOn开关
   '返回： 该按钮对象
'------------------------------------------------
Public Function ShowControl(cbrObj As CommandBars, toolId As Long, blnOn As Boolean) As CommandBarControl
    Dim Control As CommandBarControl
    Dim ControlMenu As CommandBarControl
    '   工具栏
    Set Control = cbrObj.FindControl(, toolId, , True)
    If Not Control Is Nothing Then
        Control.Enabled = blnOn
    End If
  Set ShowControl = Control
End Function
'传入ID字符串设置工具拉显示状态
Public Function ShowControlEnabled(ByVal strControlID As String, ByVal blnOn As Boolean)
    Dim strArrId As Variant, i As Integer
    strArrId = Split(strControlID, ",")
    For i = 0 To UBound(strArrId)
        Call ShowControl(Me.cbsThis, Val(strArrId(i)), blnOn)
    Next i
End Function
'获取所有加载的XML文件路劲(不包括重复的)
Public Function getXmlPath() As String
    Dim i As Integer, j As Integer, strResult As String, strArr As Variant
    For i = 1 To vsgrid.Rows - 1
        If Not vsgrid.Cell(flexcpPicture, i, 0) Is Nothing Then
            strResult = strResult & vsgrid.TextMatrix(i, 5) & ","
        End If
    Next i
    strArr = Split(Mid(strResult, 1, Len(strResult) - 1), ",")
    strResult = ""
    For i = 0 To UBound(strArr)
        If InStr(strResult, strArr(i)) = 0 Then
        strResult = strResult & "," & strArr(i)
        End If
    Next i
    strResult = Mid(strResult, 2)
    getXmlPath = strResult
End Function
'################################################################################################################
'## 功能：  将数据库数据写入到XML
'##
'## 参数：  eEdtMode    :当前编辑模式（新增、修改）
'##         eEdtType    :当前编辑方式（文件定义、示范编辑、单病历编辑、单病历审核）
'##         lngFileID   :文件ID（根据编辑方式的不同，可以表示文件定义ID、范文ID或者病人病历ID）
'##         oDoc        :XML对象
'##         oRoot       :XML根节点
'################################################################################################################
Public Function ExportToXml(ByVal eEdtMode As EditModeEnum, ByVal eEdtType As EditTypeEnum, _
ByVal lngFileID As Long, ByRef oDoc As DOMDocument, ByRef oRoot As IXMLDOMElement) As Boolean
    Dim rs As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim oFileRoot As IXMLDOMElement      '文件节点
    Dim oNode As IXMLDOMNode            '父节点
    Dim oSubNode1 As IXMLDOMNode        '子节点
    Dim oSubNode2 As IXMLDOMNode        '子节点
    Dim oSubNode3 As IXMLDOMNode        '子节点
    Dim oSubNode4 As IXMLDOMNode        '子节点
    Dim EPRFileInfoNode As IXMLDOMNode  '基础信息节点
    Dim CompendsoNode As IXMLDOMNode    '提纲节点
    Dim ElementsNode As IXMLDOMNode     '要素节点
    Dim PicturesNode As IXMLDOMNode     '图片节点
    Dim TablesNode As IXMLDOMNode       '表格节点
    Dim TableCells As IXMLDOMNode       '表格中文本集合节点
    Dim TableElements As IXMLDOMNode    '表格中要素集合节点
    Dim TablePictures As IXMLDOMNode    '表格中图片集合节点
    Dim CellNode As IXMLDOMNode         '单元格节点
    Dim ContentNode As IXMLDOMNode      '内容节点
    Dim oStream As New ADODB.Stream     '流对象
    Dim strPath As String               '临时文件目录
    Dim strTemp As String               '临时文件
    Dim strPic As String                '临时图片文件
    Dim strHeadRtfFile As String        '临时页眉文件
    Dim strFootRtfFile As String        '临时页脚文件
    Dim strContextFile As String        '临时内容文件
    Dim TempPic As New StdPicture, strTempPic As String
    Dim strObjArr As Variant
    On Error GoTo errHand
    strPath = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))
    '判断是否写入到同一个XML中
    If oRoot Is Nothing Then
        Set oRoot = oDoc.createElement("Document")
        Set oDoc.documentElement = oRoot    '设置为根节点
        Call oRoot.setAttribute("UnitName", zl9ComLib.GetUnitName)
    End If
    '设置病历文件节点
    Set oFileRoot = CreateNode(1, oRoot, "File", NODE_ELEMENT, "")
    Call oFileRoot.setAttribute("EditType", eEdtType)

    '从数据库提取病历文件基础信息
    gstrSQL = "Select a.ID, a.种类, a.编号, a.名称, a.说明, a.页面, a.保留, a.通用, b.名称 As 页面名称, b.报表, b.格式, b.页眉, b.页脚 " & _
                " From 病历文件列表 a, 病历页面格式 b " & _
                " Where a.页面 = b.编号 And a.种类 = b.种类 And a.Id = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
    Call oFileRoot.setAttribute("Name", NVL(rs("名称")))
    'EPRFileInfo
    Set EPRFileInfoNode = CreateNode(1, oFileRoot, "EPRFileInfo", NODE_ELEMENT, "")
    CreateNode 2, EPRFileInfoNode, "ID", , lngFileID      '子节点
    CreateNode 2, EPRFileInfoNode, "种类", , NVL(rs("种类"), 1)  '1-门诊病历;2-住院病历;3-护理记录;4-护理病历;5-疾病证明报告;6-知情文件;7-诊疗报告;8-诊疗申请
    CreateNode 2, EPRFileInfoNode, "编号", , NVL(rs("编号"), 0)
    CreateNode 2, EPRFileInfoNode, "名称", , NVL(rs("名称"))
    CreateNode 2, EPRFileInfoNode, "说明", , NVL(rs("说明"))
    CreateNode 2, EPRFileInfoNode, "页面", , NVL(rs("页面"))
    CreateNode 2, EPRFileInfoNode, "保留", , NVL(rs("保留"), 0)
    CreateNode 2, EPRFileInfoNode, "通用", , NVL(rs("通用"), 0)
    CreateNode 2, EPRFileInfoNode, "报表", , NVL(rs("报表"), 0)
    CreateNode 2, EPRFileInfoNode, "页面名称", , NVL(rs("页面名称"))
    CreateNode 2, EPRFileInfoNode, "格式", , NVL(rs("格式"))
    CreateNode 2, EPRFileInfoNode, "页眉", , NVL(rs("页眉"))
    CreateNode 2, EPRFileInfoNode, "页脚", , NVL(rs("页脚"))
    '读取病历内容RTF
    strContextFile = zlBlobRead(1, lngFileID)
    If strContextFile <> "" Then
       strTemp = zlFileUnzip(strContextFile)
       Me.RTbContext.LoadFile strTemp
       gobjFSO.DeleteFile strTemp
       gobjFSO.DeleteFile strContextFile, True
    End If
    '读取页眉文件（.RTF）
    strHeadRtfFile = zlBlobRead(12, NVL(rs("种类"), 1) & "-" & NVL(rs("页面")), App.Path & "\Head.rtf")
    If gobjFSO.FileExists(strHeadRtfFile) Then
        Me.RTbHeadText.LoadFile strHeadRtfFile             '读取文件
        gobjFSO.DeleteFile strHeadRtfFile, True            '删除临时文件
    End If
    CreateNode 2, EPRFileInfoNode, "页眉文件", , Replace(Me.RTbHeadText.TextRTF, "]]>", "]] >")
    '读取页脚文件（.RTF）
    strFootRtfFile = zlBlobRead(13, NVL(rs("种类"), 1) & "-" & NVL(rs("页面")), App.Path & "\Foot.rtf")
    If gobjFSO.FileExists(strFootRtfFile) Then
        Me.RTbHeadText.LoadFile strFootRtfFile              '读取文件
        gobjFSO.DeleteFile strFootRtfFile, True             '删除临时文件
    End If
    CreateNode 2, EPRFileInfoNode, "页脚文件", , Replace(Me.RTbHeadText.TextRTF, "]]>", "]] >")
    '读取页眉图片对象
    strTempPic = zlBlobRead(7, NVL(rs("种类"), 1) & "-" & NVL(rs("页面")))
    If gobjFSO.FileExists(strTempPic) Then
        Set TempPic = LoadPicture(strTempPic)
        gobjFSO.DeleteFile strTempPic, True      '删除临时文件
        If Not TempPic Is Nothing Then
            oStream.Type = adTypeBinary
            oStream.Open
            strPic = strPath & "\XMLPIC" & App.hInstance & ".jpg"
            SavePicture TempPic, strPic
            oStream.LoadFromFile strPic
            Set oSubNode1 = oDoc.createElement("OrigPic")
            oSubNode1.datatype = "bin.base64"
            oSubNode1.nodeTypedValue = oStream.Read
            EPRFileInfoNode.appendChild oSubNode1
            oStream.Close
            If gobjFSO.FileExists(strPic) Then gobjFSO.DeleteFile strPic, True
        End If
    End If
    '读取元素集合
    gstrSQL = "Select Level, ID, 文件id, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id,定义提纲ID, 复用提纲, 使用时机," & vbNewLine & _
                "       诊治要素id, 替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域" & vbNewLine & _
                "From (Select ID, 文件id, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行,预制提纲id,ID 定义提纲ID,复用提纲,使用时机," & vbNewLine & _
                "              诊治要素id, 替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域" & vbNewLine & _
                "       From 病历文件结构" & vbNewLine & _
                "       Where 文件id = [1] And 对象序号 > 0)" & vbNewLine & _
                "Start With 父id Is Null" & vbNewLine & _
                "Connect By Prior ID = 父id" & vbNewLine & _
                "Order By 对象序号, 内容行次"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
     Do While Not rs.EOF
        Select Case NVL(rs("对象类型"), 2)
            Case 1  'Compends 提纲节点
                 If CompendsoNode Is Nothing Then Set CompendsoNode = CreateNode(1, oFileRoot, "Compends", NODE_ELEMENT, "")
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
                If TablesNode Is Nothing Then Set TablesNode = CreateNode(1, oFileRoot, "Tables", NODE_ELEMENT, "")
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
                              "替换域 , 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域 From 病历文件结构 " & _
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
                                    strTempPic = zlBlobRead(2, rsTemp!ID)
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
                 If ElementsNode Is Nothing Then Set ElementsNode = CreateNode(1, oFileRoot, "Elements", NODE_ELEMENT, "")
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
                 If PicturesNode Is Nothing Then Set PicturesNode = CreateNode(1, oFileRoot, "Pictures", NODE_ELEMENT, "")
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
                    strTempPic = zlBlobRead(2, rs!ID)
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
          End Select
        rs.MoveNext
    Loop
     'RTF文本
    Set oNode = CreateNode(1, oFileRoot, "Content", NODE_ELEMENT, "")
    Set oSubNode1 = CreateNode(2, oNode, "RTF", NODE_ELEMENT, "")
    CreateNode 3, oSubNode1, "RTFText", NODE_CDATA_SECTION, Replace(Me.RTbContext.TextRTF, "]]>", "]] >")
    Exit Function
errHand:
    ExportToXml = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ExportToXml = True
End Function
'################################################################################################################
'## 功能：  将XML数据导入到数据库
'##
'## 参数：  strFilePath    :XML文件的路径
'##         strFileName    :导入病历文件的名称（不是XML的名称）
'################################################################################################################
Private Function ImportFromXml(ByVal strFilePath As String, ByVal strFileName As String) As Boolean
    '---------------------------------------------------
    Dim oDoc As New DOMDocument
    Dim oRoot As IXMLDOMElement         '根节点
    Dim oFileRoot As IXMLDOMElement     '文件节点
    Dim oNodeList As IXMLDOMNodeList    '节点集合
    Dim oNode As IXMLDOMNode            '子节点
    Dim oSubNode1 As IXMLDOMNode        '子节点
    Dim oSubNode2 As IXMLDOMNode        '子节点
    Dim oSubNode3 As IXMLDOMNode        '子节点
    Dim oSubNode4 As IXMLDOMNode        '子节点
    Dim EPRFileInfoNode As IXMLDOMNode  '基础信息节点
    Dim Compends As IXMLDOMNodeList     '提纲节点
    Dim Elements As IXMLDOMNodeList     '要素节点
    Dim Pictures As IXMLDOMNodeList     '图片节点
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
    Dim strHeadRtfFile As String        '临时页眉文件
    Dim strFootRtfFile As String        '临时页脚文件
    Dim strContextFile As String        '临时内容文
    Dim strArrNames As Variant          '文件名称数组
    Dim ArraySQL() As String            'SQL数组
    '-------------------------------------------------------------
    Dim strTempName As String
    
    Dim lngTempID As Long
    Dim lngCompendID As Long, lngID As Long, lng行次 As Long
    '-------------------------------------------------------------
    Dim GpInput As GdiplusStartupInput
    Dim m_GDIpToken         As Long         ' 用于关闭 GDI+
    Dim oDIB As New cDIB
    Dim DIBDither As New cDIBDither
    Dim DIBPal As New cDIBPal
    '-------------------------------------------------------------
    Dim TempPic As New StdPicture, strTempPic As String
    Dim rsTemp As New ADODB.Recordset, rs As New ADODB.Recordset
    Dim Result As VbMsgBoxResult
    '-------------------------------------------------------------
    On Error GoTo errHand
    'oDoc.Load strFilePath
    Set oDoc = mdoc
    Set oRoot = oDoc.selectSingleNode("Document")
    If oRoot Is Nothing Then GoTo errMsg
    strArrNames = Split(strFileName, "_")
    Set oFileRoot = oRoot.selectSingleNode("/Document/File[@Name='" & strArrNames(0) & "']")
    If oFileRoot Is Nothing Then
        If Not oRoot.selectSingleNode("EPRFileInfo") Is Nothing Then
            Set oFileRoot = oRoot
        End If
    End If
    ReDim ArraySQL(1 To 1) As String
    '导入方式判断
    Set EPRFileInfoNode = oFileRoot.selectSingleNode("EPRFileInfo")
    strTempName = EPRFileInfoNode.selectSingleNode("名称").Text
    lngTempID = NVL(EPRFileInfoNode.selectSingleNode("ID").Text, 0)
    Dim strPageName As String
    If Val(strArrNames(1)) = 2 Then '导入方式（1.覆盖 、 2.新增）
         Dim strSearchName As String
         gstrSQL = "select 名称 from 病历文件列表 where 名称 like '" & strTempName & "%'"
         Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "")
         If rsTemp.RecordCount = 0 Then
             EPRFileInfoNode.selectSingleNode("名称").Text = strTempName
             strPageName = strTempName
         ElseIf rsTemp.RecordCount = 1 Then
             If rsTemp!名称 = EPRFileInfoNode.selectSingleNode("名称").Text Then
                EPRFileInfoNode.selectSingleNode("名称").Text = strTempName & "-1"
                strPageName = strTempName & "-1"
             End If
         Else
             gstrSQL = "select '" & strTempName & "-'|| max(to_number(replace(名称,'" & strTempName & "-',''))+1) as 名称 " & _
             " from 病历文件列表 where 名称 like '" & strTempName & "-%' and instr(replace(名称,'" & strTempName & "-',''),'-')<1"
             Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "")
             EPRFileInfoNode.selectSingleNode("名称").Text = rsTemp!名称
             strPageName = rsTemp!名称
         End If
         Set rsTemp = zlDatabase.OpenSQLRecord("select LPad(nvl(max(编号),0)+1, 3,0 ) as 编号 from 病历文件列表 Where 种类 = [1]", "新增", EPRFileInfoNode.selectSingleNode("种类").Text)
         EPRFileInfoNode.selectSingleNode("页面").Text = rsTemp!编号
         Set rsTemp = zlDatabase.OpenSQLRecord("Select nvl(nvl(max(ID),0)+1,'000') as ID From 病历文件列表", "ID")
         lngTempID = Val(rsTemp!ID)   '新增ID
         gstrSQL = "Zl_病历文件列表_Insert('" & lngTempID & "','" & EPRFileInfoNode.selectSingleNode("种类").Text & "','" & EPRFileInfoNode.selectSingleNode("页面").Text & "','" & EPRFileInfoNode.selectSingleNode("名称").Text & "','" & EPRFileInfoNode.selectSingleNode("说明").Text & "','" & EPRFileInfoNode.selectSingleNode("页面").Text & "','" & strPageName & "','" & EPRFileInfoNode.selectSingleNode("报表").Text & "')"
         ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
         ArraySQL(UBound(ArraySQL)) = gstrSQL
    Else
         gstrSQL = "select a.ID,a.名称,a.编号,a.页面 ,b.名称 as 共享页面 from 病历文件列表 a , 病历文件列表 b " & _
              " Where a.种类 = " & EPRFileInfoNode.selectSingleNode("种类").Text & " and b.种类=" & EPRFileInfoNode.selectSingleNode("种类").Text & " And b.编号 = a.页面 And a.名称 ='" & EPRFileInfoNode.selectSingleNode("名称").Text & "'"
         Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "")
         lngTempID = Val(rsTemp!ID)
         EPRFileInfoNode.selectSingleNode("ID").Text = Val(rsTemp!ID)
         EPRFileInfoNode.selectSingleNode("页面").Text = rsTemp!页面
         strPageName = rsTemp!共享页面
    End If
    '从XML提取文件基础信息
    Set EPRFileInfoNode = oFileRoot.selectSingleNode("EPRFileInfo")
    ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
    ArraySQL(UBound(ArraySQL)) = "Zl_病历页面格式_Update(" & EPRFileInfoNode.selectSingleNode("种类").Text & ",'" & EPRFileInfoNode.selectSingleNode("页面").Text & "','" & strPageName & "'," & _
    EPRFileInfoNode.selectSingleNode("报表").Text & ",'" & EPRFileInfoNode.selectSingleNode("格式").Text & "'," & _
    "'" & EPRFileInfoNode.selectSingleNode("页眉").Text & "','" & EPRFileInfoNode.selectSingleNode("页脚").Text & "')"
    '从XML读取页眉文件
    If Not EPRFileInfoNode.selectSingleNode("页眉文件") Is Nothing Then Me.RTbHeadText.TextRTF = EPRFileInfoNode.selectSingleNode("页眉文件").Text
    If Me.RTbHeadText.TextRTF <> "" Then
        Me.RTbHeadText.SaveFile App.Path & "\Head.rtf"
        Call zlBlobSql(12, EPRFileInfoNode.selectSingleNode("种类").Text & "-" & EPRFileInfoNode.selectSingleNode("页面").Text, App.Path & "\Head.rtf", ArraySQL)
        gobjFSO.DeleteFile App.Path & "\Head.rtf", True
    End If
    '从XML读取页脚文件
    If Not EPRFileInfoNode.selectSingleNode("页脚文件") Is Nothing Then Me.RTbFootText.TextRTF = EPRFileInfoNode.selectSingleNode("页脚文件").Text
    If Me.RTbFootText.TextRTF <> "" Then
        Me.RTbFootText.SaveFile App.Path & "\Foot.rtf"
        Call zlBlobSql(13, EPRFileInfoNode.selectSingleNode("种类").Text & "-" & EPRFileInfoNode.selectSingleNode("页面").Text, App.Path & "\Foot.rtf", ArraySQL)
        gobjFSO.DeleteFile App.Path & "\Foot.rtf", True
    End If
    '从XML读取页眉图片
    If Not EPRFileInfoNode.selectSingleNode("OrigPic") Is Nothing Then
        oStream.Type = adTypeBinary
        oStream.Open
        oStream.Write EPRFileInfoNode.selectSingleNode("OrigPic").nodeTypedValue
        strPic = App.Path & "\XML2JPG" & App.hInstance & ".JPG"
        oStream.SaveToFile strPic, adSaveCreateOverWrite
        oStream.Close
        Call zlBlobSql(7, EPRFileInfoNode.selectSingleNode("种类").Text & "-" & EPRFileInfoNode.selectSingleNode("页面").Text, strPic, ArraySQL)
        gobjFSO.DeleteFile strPic, True      '删除临时文件
    End If
    '从XML提取提纲信息
    If Not oFileRoot.selectSingleNode("Compends") Is Nothing Then Set Compends = oFileRoot.selectSingleNode("Compends").selectNodes("Compend")
    If Not Compends Is Nothing Then
        For Each oNode In Compends
            lngCompendID = zlDatabase.GetNextId("病历文件结构")
            ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
            ArraySQL(UBound(ArraySQL)) = "Zl_病历文件结构_Update(" & lngCompendID & "," & lngTempID & "," & IIf(oNode.selectSingleNode("父ID").Text = 0, "NULL", oNode.selectSingleNode("父ID").Text) & "," & _
                oNode.selectSingleNode("对象序号").Text & ",1," & oNode.selectSingleNode("Key").Text & "," & IIf(oNode.selectSingleNode("保留对象").Text, 1, 0) & ",'" & oNode.selectSingleNode("说明").Text & "',NULL,'" & oNode.selectSingleNode("名称").Text & "',NULL," & _
                IIf(oNode.selectSingleNode("预制提纲ID").Text = 0, "NULL", oNode.selectSingleNode("预制提纲ID").Text) & "," & IIf(oNode.selectSingleNode("复用提纲").Text, 1, 0) & ",'" & oNode.selectSingleNode("使用时机").Text & "')"
            Set oNodeList = oFileRoot.selectNodes("//*[父ID=" & oNode.selectSingleNode("ID").Text & " ]")
            For Each oSubNode1 In oNodeList
                oSubNode1.selectSingleNode("父ID").Text = lngCompendID
            Next
        Next
       
    End If
    '从XML提取要素信息
    If Not oFileRoot.selectSingleNode("Elements") Is Nothing Then Set Elements = oFileRoot.selectSingleNode("Elements").selectNodes("Element")
    If Not Elements Is Nothing Then
        For Each oNode In Elements
            ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
            ArraySQL(UBound(ArraySQL)) = "Zl_病历文件结构_Update(" & zlDatabase.GetNextId("病历文件结构") & "," & lngTempID & "," & IIf(oNode.selectSingleNode("父ID").Text = 0, "NULL", oNode.selectSingleNode("父ID").Text) & "," & _
                oNode.selectSingleNode("对象序号").Text & ",4," & oNode.selectSingleNode("Key").Text & "," & IIf(oNode.selectSingleNode("保留对象").Text, 1, 0) & ",'" & oNode.selectSingleNode("对象属性").Text & "',NULL,'" & _
                Replace(oNode.selectSingleNode("内容文本").Text, "'", "' || chr(39) || '") & "'," & IIf(oNode.selectSingleNode("是否换行").Text, 1, 0) & ",NULL,NULL,NULL," & _
                IIf(CheckValid(oNode.selectSingleNode("诊治要素ID").Text, oNode.selectSingleNode("要素名称").Text), oNode.selectSingleNode("诊治要素ID").Text, "NULL") & "," & _
                oNode.selectSingleNode("替换域").Text & ",'" & oNode.selectSingleNode("要素名称").Text & "'," & oNode.selectSingleNode("要素类型").Text & "," & oNode.selectSingleNode("要素长度").Text & "," & _
                oNode.selectSingleNode("要素小数").Text & ",'" & oNode.selectSingleNode("要素单位").Text & "'," & oNode.selectSingleNode("要素表示").Text & "," & oNode.selectSingleNode("输入形态").Text & ",'" & oNode.selectSingleNode("要素值域").Text & "')"
        Next
    End If
    '从XML提取表格信息
    If Not oFileRoot.selectSingleNode("Tables") Is Nothing Then Set Tables = oFileRoot.selectSingleNode("Tables").selectNodes("Table")
    If Not Tables Is Nothing Then
        For Each oNode In Tables
            lngID = zlDatabase.GetNextId("病历文件结构")
            '保存表格结构SQL语句
            ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
            ArraySQL(UBound(ArraySQL)) = "Zl_病历文件结构_Update(" & lngID & "," & lngTempID & "," & IIf(oNode.selectSingleNode("父ID").Text = 0, "NULL", oNode.selectSingleNode("父ID").Text) & "," & _
            oNode.selectSingleNode("对象序号").Text & ",3," & oNode.selectSingleNode("Key").Text & "," & IIf(oNode.selectSingleNode("保留对象").Text, 1, 0) & ",'" & oNode.selectSingleNode("对象属性").Text & "',NULL,'" & "" & "'," & IIf(oNode.selectSingleNode("是否换行").Text, 1, 0) & _
            "," & IIf(oNode.selectSingleNode("预制提纲ID").Text = 0, "NULL", oNode.selectSingleNode("预制提纲ID").Text) & ")"
            '更改所有子项的父ID
            Set oNodeList = oNode.selectNodes("//*[父ID=" & oNode.selectSingleNode("ID").Text & " ]")
            For Each oSubNode1 In oNodeList
                oSubNode1.selectSingleNode("父ID").Text = lngID
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
                            ArraySQL(UBound(ArraySQL)) = "Zl_病历文件结构_Update(" & zlDatabase.GetNextId("病历文件结构") & "," & lngTempID & "," & IIf(oSubNode2.selectSingleNode("父ID").Text = 0, "NULL", oSubNode2.selectSingleNode("父ID").Text) & "," & _
                            IIf(oSubNode2.selectSingleNode("对象序号").Text = 0, "NULL", oSubNode2.selectSingleNode("对象序号").Text) & ",4," & oSubNode2.selectSingleNode("Key").Text & "," & IIf(oSubNode2.selectSingleNode("保留对象").Text, 1, 0) & ",'" & _
                            oSubNode2.selectSingleNode("对象属性").Text & "'," & lng行次 & ",'" & Replace(oSubNode2.selectSingleNode("内容文本").Text, "'", "' || chr(39) || '") & "'," & IIf(oSubNode2.selectSingleNode("是否换行").Text, 1, 0) & ",NULL,NULL,NULL," & _
                            IIf(CheckValid(oSubNode2.selectSingleNode("诊治要素ID").Text, oSubNode2.selectSingleNode("要素名称").Text), oSubNode2.selectSingleNode("诊治要素ID").Text, "NULL") & "," & _
                            oSubNode2.selectSingleNode("替换域").Text & ",'" & oSubNode2.selectSingleNode("要素名称").Text & "'," & oSubNode2.selectSingleNode("要素类型").Text & "," & oSubNode2.selectSingleNode("要素长度").Text & "," & _
                            oSubNode2.selectSingleNode("要素小数").Text & ",'" & oSubNode2.selectSingleNode("要素单位").Text & "'," & oSubNode2.selectSingleNode("要素表示").Text & "," & oSubNode2.selectSingleNode("输入形态").Text & ",'" & oSubNode2.selectSingleNode("要素值域").Text & "')"
                        End If
                    Else '文本
                        ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
                         ArraySQL(UBound(ArraySQL)) = "Zl_病历文件结构_Update(" & zlDatabase.GetNextId("病历文件结构") & "," & lngTempID & "," & oSubNode1.selectSingleNode("父ID").Text & ",NULL," & _
                        "2," & oSubNode1.selectSingleNode("Key").Text & ",NULL,'" & oSubNode1.selectSingleNode("对象属性").Text & "'," & lng行次 & ",'" & Replace(oSubNode1.selectSingleNode("内容文本").Text, "'", "' || chr(39) || '") & "')"
                    End If
                    lng行次 = lng行次 + 1
                Next
            End If
            '图片处理
            If Not TablePictures Is Nothing Then
                For Each oSubNode1 In TablePictures
                        ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
                        lngID = zlDatabase.GetNextId("病历文件结构")
                        ArraySQL(UBound(ArraySQL)) = "Zl_病历文件结构_Update(" & lngID & "," & lngTempID & "," & IIf(oSubNode1.selectSingleNode("父ID").Text = 0, "NULL", oSubNode1.selectSingleNode("父ID").Text) & "," & _
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
                        gstrSQL = "select 对象ID from 病历文件图形 where 对象ID=[1]"
                        Call zlBlobSql(2, lngID, strPic, ArraySQL)
                        oStream.Close
                Next
            End If
        Next
    End If
    '从XML提取内容图片信息
    If Not oFileRoot.selectSingleNode("Pictures") Is Nothing Then Set Pictures = oFileRoot.selectSingleNode("Pictures").selectNodes("Picture")
    If Not Pictures Is Nothing Then
        For Each oNode In Pictures
            ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
            lngID = zlDatabase.GetNextId("病历文件结构")
            ArraySQL(UBound(ArraySQL)) = "Zl_病历文件结构_Update(" & lngID & "," & lngTempID & "," & IIf(oNode.selectSingleNode("父ID").Text = 0, "NULL", oNode.selectSingleNode("父ID").Text) & "," & _
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
            gstrSQL = "select 对象ID from 病历文件图形 where 对象ID=[1]"
            Call zlBlobSql(2, lngID, strPic, ArraySQL)
            oStream.Close
        Next
    End If
    '后期处理
     ReDim Preserve ArraySQL(1 To UBound(ArraySQL) + 1) As String
     gstrSQL = "Zl_病历文件结构_Commit(" & lngTempID & ")"
     ArraySQL(UBound(ArraySQL)) = gstrSQL
    '=========================================================================================
    '保存RTFText的Sql
    '=========================================================================================
    If Not oFileRoot.selectSingleNode("Content") Is Nothing Then
        Set ContentNode = oFileRoot.selectSingleNode("Content")
        Me.RTbContext.TextRTF = ContentNode.selectSingleNode("RTF").Text
        If gobjFSO.FileExists(App.Path & "\TMP.rtf") Then gobjFSO.DeleteFile App.Path & "\TMP.rtf", True    '保存为临时文件
        Me.RTbContext.SaveFile App.Path & "\TMP.rtf"
        strTemp = zlFileZip(App.Path & "\TMP.rtf")
        If gobjFSO.FileExists(App.Path & "\TMP.rtf") Then gobjFSO.DeleteFile App.Path & "\TMP.rtf", True
        If gobjFSO.FileExists(strTemp) Then
            Call zlBlobSql(1, lngTempID, strTemp, ArraySQL)
            gobjFSO.DeleteFile strTemp, True      '删除临时文件
        End If
    End If
    '#########################################################################################
    '启动事务
    '=========================================================================================
bb:    If Not BeginTrans(ArraySQL) Then gcnOracle.RollbackTrans: Err.Clear: GoTo errMsg
       ImportFromXml = True
       Exit Function
errHand:
    ImportFromXml = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ImportFromXml = True
    Exit Function
errMsg:
     Result = ZLCommFun.ShowMsgBox("病历文件导入", strTempName & ",导入数据格式不正确或已被损坏  ！", "重试(&A),忽略(&O)", Nothing)
     If Result = "重试" Then GoTo bb
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











