VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmImportFile 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "导入项目"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10425
   ForeColor       =   &H00FF0000&
   Icon            =   "frmImportFile.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   10425
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   120
      ScaleHeight     =   2865
      ScaleWidth      =   5145
      TabIndex        =   6
      Top             =   900
      Width           =   5175
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   2565
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   4905
         _cx             =   8652
         _cy             =   4524
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
         BackColorSel    =   4227072
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmImportFile.frx":6852
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
         ExplorerBar     =   1
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
         VirtualData     =   0   'False
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
   End
   Begin VB.TextBox txtFile 
      Enabled         =   0   'False
      Height          =   300
      Left            =   960
      TabIndex        =   2
      Top             =   480
      Width           =   3495
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "…"
      Height          =   300
      Left            =   4440
      TabIndex        =   1
      Top             =   480
      Width           =   280
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   250
      Left            =   300
      MousePointer    =   7  'Size N S
      ScaleHeight     =   354.167
      ScaleMode       =   0  'User
      ScaleWidth      =   4215
      TabIndex        =   0
      Top             =   3990
      Width           =   4215
      Begin VB.Label lblCollect 
         BackColor       =   &H80000005&
         Caption         =   "列的输入提示"
         Height          =   180
         Left            =   45
         TabIndex        =   8
         Top             =   45
         Width           =   1080
      End
   End
   Begin MSComctlLib.ImageList imgError 
      Left            =   5145
      Top             =   75
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
            Picture         =   "frmImportFile.frx":68C7
            Key             =   "error"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportFile.frx":D129
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   6090
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.xls|*.xls|*.xlsx|*.xlsx"
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfError 
      Height          =   765
      Left            =   75
      TabIndex        =   3
      Top             =   4425
      Width           =   5055
      _cx             =   8916
      _cy             =   1349
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
      BackColorSel    =   8454016
      ForeColorSel    =   16744576
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmImportFile.frx":1398B
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
      ExplorerBar     =   1
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
      VirtualData     =   0   'False
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
   Begin XtremeSuiteControls.TabControl TabControl 
      Height          =   750
      Left            =   7305
      TabIndex        =   5
      Top             =   1110
      Width           =   1500
      _Version        =   589884
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   64
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "文  件(&F)"
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   525
      Width           =   810
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgPicture 
      Bindings        =   "frmImportFile.frx":13A00
      Left            =   840
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmImportFile.frx":13A14
   End
End
Attribute VB_Name = "frmImportFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbyt导入方式, mbyt分类类别, mbyt上级名称, mbyt编码, mbyt名称, mbytNumAndKind, mbytNameAndKind As Byte
Private mbyt药品类别, mbyt药品分类, mbyt品种编码, mbyt品种名称, mbyt规格编码, mbyt药品规格, mbyt产地, mbyt剂型, mbyt单位, mbyt换算系数, mbyt变价, mbyt价格, mbyt效期, mbyt收入项目, mbyt分零, mbyt服务对象, mbyt分批属性, mbyt供应商, mbyt日期, mbyt品种唯一, mbyt规格唯一 As Byte
'表格列(列名,0必选1可选,0显示1隐藏|...)
Private Const MSTRMEDICAL      As String = "类别,0,0|上级名称,0,0|编码,0,0|名称,0,0||类别,0,0|分类,0,0|品种编码,0,0|品种名称,0,0|规格编码,0,0|药品规格,0,0|生产商,0,0|剂型,0,0|剂量单位,0,0|售价单位,0,0|售价换算系数,0,0|门诊单位,0,0|门诊换算系数,0,0|住院单位,0,0|住院换算系数,0,0|药库单位,0,0|药库换算系数,0,0|" & _
                                        "是否变价,0,0|成本价,0,0|售价,0,0|收入项目,0,0|住院可否分零,0,0|门诊可否分零,0,0|服务对象,1,0|药库分批,1,0|药房分批,1,0|效期(月),1,0|供应商名称,1,0|供应商许可证号,1,0|供应商许可证效期,1,0|"
'列的输入提示(列名;提示|...)
Private Const MSTRCOMMENT      As String = "类别;类别只能是西成药、中成药、中草药，不能为空|上级名称;上级分类参照表中已有的数据：诊疗分类目录.名称，格式：用\分隔各个级别 例：一级卫材\其他，为空表示没有上级|编码;编码不能为空，长度不能超过数据库字段长度|名称;名称不能含有非法字符，如：单引号，不能为空，长度不能超过数据库字段长度||" & _
                                        "类别;类别只能是西成药、中成药、中草药，不能为空|分类;分类参照表中已有的数据：诊疗分类目录.名称 格式：用\分隔各个级别 例：一级卫材\其他，不能为空|品种编码;品种编码不能为空，长度不能超过数据库字段长度|品种名称;品种名称不能含有非法字符，如：单引号，不能为空，长度不能超过数据库字段长度|" & _
                                        "规格编码;规格编码不能为空，长度不能超过数据库字段长度|药品规格;药品规格不能含有非法字符，如：单引号，不能为空，长度不能超过数据库字段长度|生产商;生产商字段长度不能超过数据库字段设计长度，不能含有非法字符，如：单引号，可以为空|剂型;剂型参照数据库表中已有数据，不能为空，不能含有非法字符，长度不能超过数据库字段设计长度|" & _
                                        "剂量单位;单位不能为空，长度不能超过数据库字段设计长度|售价单位;|售价换算系数;换算系数不能为空，单位换算系数合理且都>0，单位相同换算系数必须相同且都是数字|门诊单位;|门诊换算系数;|住院单位;|住院换算系数;|药库单位;|药库换算系数;|是否变价;为空默认为定价，“√”表示时价|成本价;价格字段只能是数字型，精度不能超过最大设置精度|" & _
                                        "售价;|收入项目;收入项目不能为空，只能是数据库已有收入项目|住院可否分零;分零方式：0-可以分零,1-不可分零,2-一次性使用,3-分零后一天内有效,4-分零后两天内有效,5-分零后三天内有效|门诊可否分零;|服务对象;服务对象：0和空-不服务于病人，1-门诊，2-住院，3-门诊和住院|" & _
                                        "药库分批;为空表示不分批，“√”表示分批|药房分批;药库分批时药房才能分批|效期(月);效期必须是数字且只能是不小于0的整数|供应商名称;供应商参照数据库表中已有数据，可以为空|供应商许可证号;|供应商许可证效期;录入日期的必须符合日期格式：2015-10-10或者2015/10/10或者2015.10.10|"

'药品明细下拉框(列名1|选项1,选项2,选项3;列名2|选项1,选项2,选项3)
Private Const mstr药品明细 As String = "类别|西成药,中成药,中草药;是否变价|√;住院可否分零|0-可以分零,1-不可分零,2-一次性使用,3-分零后一天内有效,4-分零后两天内有效,5-分零后三天内有效;" & _
                                                            "门诊可否分零|0-可以分零,1-不可分零,2-一次性使用,3-分零后一天内有效,4-分零后两天内有效,5-分零后三天内有效;服务对象|0-不服务于病人,1-门诊,2-住院,3-门诊和住院;" & _
                                                            "药库分批|√;药房分批|√"
Private Const MCONTOOLMODE     As Integer = 100  'Excel样本
Private Const MCONTOOLOUTPUT   As Integer = 101  '导出Excel
Private Const MCONTOOLCHECK    As Integer = 102  '校验
Private Const MCONTOOLSAVE     As Integer = 103  '保存
Private Const MCONTOOLEXIT     As Integer = 104  '退出
Private Const MCONTOOLCHECKSET As Integer = 107  '检查设置
Private Const MCONTOOLCOLSET   As Integer = 109  '列设置
Private mstrType               As String         '显示的分类列头
Private mstrMedi               As String         '显示的明细列头
Private mstrTypeMsg            As String         '分类表格中所有信息
Private mstrMediMsg            As String         '明细表格中所有信息
Private mintType               As Integer        '导入文件类型 1-费用，2-药品，3-卫材
Private mlngModule             As Long           '模块号
Private mobjXLS As Object
Private mobjWB As Object
Private mobjWS As Object


Private Sub InitComandbar()
    '初始化工具栏
    Dim cbrControlMain As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim cbrToolPopup As CommandBarPopup
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsMain.VisualTheme = xtpThemeOffice2003

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
    Me.cbsMain.EnableCustomization False
    Me.cbsMain.Icons = Me.imgPicture.Icons
    
    '工具栏定义
    Set cbrToolBar = Me.cbsMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched Or xtpFlagFloating Or xtpFlagAlignAny
        
    With cbrToolBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, MCONTOOLMODE, "生成Excel样本")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        Set cbrControlMain = .Add(xtpControlButton, MCONTOOLOUTPUT, "导出Excel")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        Set cbrControlMain = .Add(xtpControlButton, MCONTOOLCHECKSET, "检查设置")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, MCONTOOLCOLSET, "列设置")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        Set cbrControlMain = .Add(xtpControlButton, MCONTOOLCHECK, "校验")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, MCONTOOLSAVE, "保存")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        Set cbrControlMain = .Add(xtpControlButton, MCONTOOLEXIT, "退出")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
    End With
    cbsMain.Item(1).Delete
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case MCONTOOLMODE     'Excel样本
            Call ProduceStyleBook
        Case MCONTOOLOUTPUT   '导出Excel
            Call OutPutFile
        Case MCONTOOLCHECKSET '条件设置
            frmImportFileCondition.ShowMe Me, mlngModule
            If vsfList.Rows > 1 Then
                If TabControl.Selected.Caption = "分类" Then Call CheckKind
                If TabControl.Selected.Caption = "明细" Then Call Check品种: Call Check规格
            End If
        Case MCONTOOLCOLSET   '列设置
            Call SetCols
        Case MCONTOOLCHECK    '校验
            Call FS.ShowFlash("正在校验数据,请稍候 ...", Me)
            Me.MousePointer = vbHourglass
            If TabControl.Selected.Caption = "分类" Then
                Call CheckKind
                Call GetColumns("分类")
                Call CheckKind
                Call GetColumns("分类")
            Else
                Call Check品种
                Call Check规格
                Call GetColumns("明细")
                Call Check品种
                Call Check规格
                Call GetColumns("明细")
            End If
            Me.MousePointer = vbDefault
            Call FS.StopFlash
        Case MCONTOOLSAVE     '保存
            Call SaveCard
        Case MCONTOOLEXIT     '退出
            Unload Me
    End Select
End Sub

Private Sub OutPutFile()
    '导出表格文件
    Dim strFileName As String
    Dim i As Long
    Dim j As Long
    Dim arrType As Variant
    Dim arrMedi As Variant
    Dim intNum As Integer
    Dim blnFinished As Boolean
    
    On Error GoTo ErrHand
    
    arrType = Split(mstrTypeMsg, "|")
    arrMedi = Split(mstrMediMsg, "|")
    
    If mobjXLS Is Nothing Then Call InitExcel
    mobjXLS.SheetsInNewWorkbook = 1  '将新建的工作薄数量设为1
    mobjXLS.Workbooks.Add          '增加一个工作薄
    mobjXLS.Sheets(mobjXLS.Sheets.Count).Name = "药品分类"  '修改工作薄名称
    mobjXLS.Sheets.Add , mobjXLS.Sheets("药品分类") '增加第二个工作薄在第一个之后
    mobjXLS.Sheets(mobjXLS.Sheets.Count).Name = "药品明细"
    
    mobjXLS.Sheets("药品分类").Select     '选中工作薄<药品分类>
    mobjXLS.Columns("A:L").NumberFormatLocal = "@"   '设置文本格式
    '循环写入数据
    For i = LBound(arrType) To UBound(arrType) - 1
        For j = 0 To UBound(Split(arrType(i), ";")) - 1
            mobjXLS.Cells(i + 1, j + 1) = Split(arrType(i), ";")(j)
            '设置Excel批注
            If i = 0 Then
                For intNum = 0 To UBound(Split(Split(MSTRCOMMENT, "||")(0) & "|", "|")) - 1
                    If Split(Split(Split(MSTRCOMMENT, "||")(0) & "|", "|")(intNum), ";")(0) = Split(arrType(i), ";")(j) Then
                        If Split(Split(Split(MSTRCOMMENT, "||")(0) & "|", "|")(intNum), ";")(1) <> "" Then
                            mobjXLS.ActiveSheet.Cells(i + 1, j + 1).AddComment Split(Split(Split(MSTRCOMMENT, "||")(0) & "|", "|")(intNum), ";")(1)
                        End If
                        Exit For
                    End If
                Next
            End If
        Next
    Next
    Call SetExcel("分类", j, i)
    
    mobjXLS.Sheets("药品明细").Select     '选中工作薄<药品明细>
    mobjXLS.Columns("A:AE").NumberFormatLocal = "@"   '设置文本格式
    '循环写入数据
    For i = LBound(arrMedi) To UBound(arrMedi) - 1
        For j = 0 To UBound(Split(arrMedi(i), ";")) - 1
            mobjXLS.Cells(i + 1, j + 1) = Split(arrMedi(i), ";")(j)
            '设置Excel批注
            If i = 0 Then
                For intNum = 0 To UBound(Split(Split(MSTRCOMMENT, "||")(1), "|")) - 1
                    If Split(Split(Split(MSTRCOMMENT, "||")(1), "|")(intNum), ";")(0) = Split(arrMedi(i), ";")(j) Then
                        If Split(Split(Split(MSTRCOMMENT, "||")(1), "|")(intNum), ";")(1) <> "" Then
                            mobjXLS.ActiveSheet.Cells(i + 1, j + 1).AddComment Split(Split(Split(MSTRCOMMENT, "||")(1), "|")(intNum), ";")(1)
                        End If
                        Exit For
                    End If
                Next
            End If
        Next
    Next
    Call SetExcel("明细", j, i)
    mobjXLS.Sheets("药品分类").Select
    
    With dlgOpenFile
        .CancelError = True
        .FileName = ""
        .Filter = "*.xlsx|*.xlsx|*.xls|*.xls"
        .ShowSave
        strFileName = .FileName
        If Trim(strFileName) <> "" Then
            mobjXLS.ActiveWorkbook.SaveAs strFileName
            blnFinished = True
        End If
    End With
    
ErrHand:
    mobjXLS.Quit
    If blnFinished Then
        MsgBox "导出成功！", vbInformation, gstrSysName
    Else
        MsgBox "导出失败！", vbInformation, gstrSysName
    End If
End Sub

Private Sub cmdFile_Click()
'获取文件，提取数据
    On Error GoTo ErrHand
    
    dlgOpenFile.FileName = ""
    dlgOpenFile.Filter = "*.xlsx|*.xlsx|*.xls|*.xls"
    dlgOpenFile.ShowOpen
    If dlgOpenFile.FileName <> "" Then
        txtFile.Text = dlgOpenFile.FileName
    Else
        GoTo ErrHand
    End If
    
    If txtFile.Text <> "" Then
        DoEvents
        Call FS.ShowFlash("正在加载数据,请稍候 ...", Me)
        Me.MousePointer = vbHourglass
        '提取数据
        Call GetExcelData
        Me.MousePointer = vbDefault
        Call FS.StopFlash
    End If
    
    Exit Sub
ErrHand:
    Exit Sub
End Sub

Private Sub ParseParameter()
'解析参数，获取校验方式
    Dim arryPara As Variant
    Dim strPara  As String
    
    '导入方式/类别|上级名称|编码|名称|编码和类别唯一检查|名称、类别、上级分类唯一检查|类别|分类|品种编码|品种名称|规格编码|药品规格|产地合法性检查|剂型|各级单位检查|各级单位换算检查|变价检查|价格检查|效期|收入项目|门诊/住院分零|服务对象|分批属性|供应商|日期格式|品种唯一性检查|规格唯一性检查
    '(0错误提示1错误禁止/0提示1禁止|....)
    strPara = zlDatabase.GetPara("导入文件检查方式", glngSys, mlngModule, "0/0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0")
    mbyt导入方式 = Mid(strPara, 1, 1)
    strPara = Mid(strPara, 3)
    arryPara = Split(strPara, "|")
    '分类
    mbyt分类类别 = arryPara(0)
    mbyt上级名称 = arryPara(1)
    mbyt编码 = arryPara(2)
    mbyt名称 = arryPara(3)
    mbytNumAndKind = arryPara(4)
    mbytNameAndKind = arryPara(5)
    '明细
    mbyt药品类别 = arryPara(6)
    mbyt药品分类 = arryPara(7)
    mbyt品种编码 = arryPara(8)
    mbyt品种名称 = arryPara(9)
    mbyt规格编码 = arryPara(10)
    mbyt药品规格 = arryPara(11)
    mbyt产地 = arryPara(12)
    mbyt剂型 = arryPara(13)
    mbyt单位 = arryPara(14)
    mbyt换算系数 = arryPara(15)
    mbyt变价 = arryPara(16)
    mbyt价格 = arryPara(17)
    mbyt效期 = arryPara(18)
    mbyt收入项目 = arryPara(19)
    mbyt分零 = arryPara(20)
    mbyt服务对象 = arryPara(21)
    mbyt分批属性 = arryPara(22)
    mbyt供应商 = arryPara(23)
    mbyt日期 = arryPara(24)
    mbyt品种唯一 = arryPara(25)
    mbyt规格唯一 = arryPara(26)
End Sub

Private Sub ProduceStyleBook()
'生成导入外部文件的标准XLS文件样本
    Dim arrTypeCols As Variant
    Dim arrMediCols As Variant
    Dim blnFinished As Boolean
    Dim strFileName As String
    Dim strMedi     As String
    Dim intNum      As Integer
    Dim i           As Integer
    
    On Error GoTo ErrHand
    
    strMedi = zlDatabase.GetPara("列的显示隐藏", glngSys, mlngModule, MSTRMEDICAL)
    arrTypeCols = Split(Split(strMedi, "||")(0) & "|", "|")
    arrMediCols = Split(Split(strMedi, "||")(1), "|")
    
    If mobjXLS Is Nothing Then Call InitExcel
    mobjXLS.SheetsInNewWorkbook = 1  '将新建的工作薄数量设为1
    mobjXLS.Workbooks.Add          '增加一个工作薄
    mobjXLS.Sheets(mobjXLS.Sheets.Count).Name = "药品分类"  '修改工作薄名称
    mobjXLS.Sheets.Add , mobjXLS.Sheets("药品分类") '增加第二个工作薄在第一个之后
    mobjXLS.Sheets(mobjXLS.Sheets.Count).Name = "药品明细"
    
    mobjXLS.Sheets("药品分类").Select     '选中工作薄<药品分类>
    '循环写入数据
    For i = LBound(arrTypeCols) To UBound(arrTypeCols) - 1
        mobjXLS.Cells(1, i + 1) = Split(arrTypeCols(i), ",")(0)
        '设置Excel批注
        For intNum = 0 To UBound(Split(Split(MSTRCOMMENT, "||")(0) & "|", "|")) - 1
            If Split(Split(Split(MSTRCOMMENT, "||")(0) & "|", "|")(intNum), ";")(0) = Split(arrTypeCols(i), ",")(0) Then
                If Split(Split(Split(MSTRCOMMENT, "||")(0) & "|", "|")(intNum), ";")(1) <> "" Then
                    mobjXLS.ActiveSheet.Cells(1, i + 1).AddComment Split(Split(Split(MSTRCOMMENT, "||")(0) & "|", "|")(intNum), ";")(1)
                End If
                Exit For
            End If
        Next
    Next
    Call SetExcel("分类", i, 1)
    
    mobjXLS.Sheets("药品明细").Select     '选中工作薄<药品明细>
    '循环写入数据
    For i = LBound(arrMediCols) To UBound(arrMediCols) - 1
        mobjXLS.Cells(1, i + 1) = Split(arrMediCols(i), ",")(0)
        For intNum = 0 To UBound(Split(Split(MSTRCOMMENT, "||")(1), "|")) - 1
            If Split(Split(Split(MSTRCOMMENT, "||")(1), "|")(intNum), ";")(0) = Split(arrMediCols(i), ",")(0) Then
                If Split(Split(Split(MSTRCOMMENT, "||")(1), "|")(intNum), ";")(1) <> "" Then
                    mobjXLS.ActiveSheet.Cells(1, i + 1).AddComment Split(Split(Split(MSTRCOMMENT, "||")(1), "|")(intNum), ";")(1)
                End If
                Exit For
            End If
        Next
    Next
    Call SetExcel("明细", i, 1)
    mobjXLS.Sheets("药品分类").Select
    
    With dlgOpenFile
        .CancelError = False
        .FileName = ""
        .Filter = "*.xlsx|*.xlsx|*.xls|*.xls"
        .ShowSave
        strFileName = .FileName
        If Trim(strFileName) <> "" Then
            mobjXLS.ActiveWorkbook.SaveAs strFileName
            blnFinished = True
        End If
    End With
ErrHand:
    mobjXLS.Quit
    If blnFinished Then
        MsgBox "标准文件样本已经生成！", vbInformation, gstrSysName
    Else
        MsgBox "标准文件样本生成失败！", vbInformation, gstrSysName
    End If
End Sub

Private Sub SetExcel(ByVal strType As String, ByVal intCol As Integer, ByVal intRow As Integer)
'设置Excel属性
    Dim intCount As Integer
    Dim strMedi  As String
    Dim strStart As String
    Dim strEnd   As String
    Dim strFileColumn As String
    Dim lngCol As Long
    Dim strArr明细下拉框() As String
    Dim strArr明细列名() As String
    Dim strArr分类列名() As String
    Dim arrMediCols() As String
    Dim i As Integer, n As Integer
    
    On Error GoTo ErrHand
    
    strMedi = zlDatabase.GetPara("列的显示隐藏", glngSys, mlngModule, MSTRMEDICAL)
    
    With mobjXLS
        If strType = "分类" Then
            For intCount = 0 To intCol - 1
                .ActiveCell(1, intCount + 1).HorizontalAlignment = 3  '列头文本居中对齐
                If Split(Split(Split(strMedi, "||")(0), "|")(intCount), ",")(1) = 0 Then
                    .ActiveSheet.Cells(1, intCount + 1).Interior.Color = &H80FF80  '绿色为必须显示项目
                End If
            Next
            '设置边框
            If intCount < 27 Then
                strEnd = Chr(intCount - 1 + 65) & intRow
            Else
                strEnd = "A" & Chr(intCount - 27 + 65) & intRow
            End If
            .Range("A1", strEnd).Borders.Weight = 2
        Else
            For intCount = 0 To intCol - 1
                .ActiveCell(1, intCount + 1).HorizontalAlignment = 3  '列头文本居中对齐
                If Split(Split(Split(strMedi, "||")(1), "|")(intCount), ",")(1) = 0 Then
                    .ActiveSheet.Cells(1, intCount + 1).Interior.Color = &H80FF80  '绿色为必须显示项目
                End If
            Next
            '设置边框
            If intCount < 27 Then
                strEnd = Chr(intCount - 1 + 65) & intRow
            Else
                strEnd = "A" & Chr(intCount - 27 + 65) & intRow
            End If
            .Range("A1", strEnd).Borders.Weight = 2
        End If
        
        .Rows("1:1").Select           '选中第一行
        .Selection.Font.Bold = True   '设为粗体
        .Selection.Font.Size = 11     '设置字体大小
        .Columns.ColumnWidth = 16     '列宽
        .ActiveWindow.SplitRow = 1    '固定行
        .ActiveWindow.FreezePanes = True
        .ActiveSheet.Rows(1).RowHeight = 25  '行高
        .ActiveSheet.Rows(1).Insert   '插入一行
        .Cells(1).Value = " 说明：绿色为必须显示项目，填表时请注意查看批注。"
'        .Range("A1:C1").Select        '合并
'        .Range("A1:C1").Merge
        .Range("A3").Select
        
        '药品分类下拉框
        If strType = "分类" Then
            .Sheets("药品分类").Select
            .Columns("A:L").NumberFormatLocal = "@"   '设置文本格式
            arrMediCols = Split(Split(strMedi, "||")(0), "|")
            strFileColumn = ""
            For i = LBound(arrMediCols) To UBound(arrMediCols)
                strFileColumn = strFileColumn & "|" & Trim(Split(arrMediCols(i), ",")(0)) & "," & i + 1
            Next
            strFileColumn = Mid(strFileColumn, 2)
            strArr分类列名 = Split(strFileColumn, "|")
            For i = LBound(strArr分类列名) To UBound(strArr分类列名)
                If Split(strArr分类列名(i), ",")(0) = "类别" Then
                    lngCol = Split(strArr分类列名(i), ",")(1)
                    .Columns(lngCol).Select
                    With mobjXLS.Selection.Validation
                        .Delete
                        .Add 3, 1, 1, "西成药,中成药,中草药"
                        .IgnoreBlank = True
                        .InCellDropdown = True
                        .InputTitle = ""
                        .ErrorTitle = "输入提醒"
                        .InputMessage = ""
                        .ErrorMessage = "必须从下拉列表中选择"
                        .IMEMode = 0
                        .ShowInput = True
                        .ShowError = True
                    End With

                    .Rows("1:2").Select
                    With mobjXLS.Selection.Validation
                        .Delete
                    End With
                End If
            Next
        End If
        '药品明细下拉框
        If strType = "明细" Then
            arrMediCols = Split(Split(strMedi, "||")(1), "|")
            strFileColumn = ""
            For i = LBound(arrMediCols) To UBound(arrMediCols) - 1
                strFileColumn = strFileColumn & "|" & Trim(Split(arrMediCols(i), ",")(0)) & "," & i + 1
            Next
            strFileColumn = Mid(strFileColumn, 2)
            strArr明细列名 = Split(strFileColumn, "|")
            strArr明细下拉框 = Split(mstr药品明细, ";")
            For i = LBound(strArr明细列名) To UBound(strArr明细列名)
                For n = LBound(strArr明细下拉框) To UBound(strArr明细下拉框)
                    If Split(strArr明细列名(i), ",")(0) = Split(strArr明细下拉框(n), "|")(0) Then
                        .Sheets("药品明细").Select
                        .Columns("A:AE").NumberFormatLocal = "@"   '设置文本格式
                        lngCol = Split(strArr明细列名(i), ",")(1)
                        .Columns(lngCol).Select
                        With mobjXLS.Selection.Validation
                            .Delete
                            .Add 3, 1, 1, Split(strArr明细下拉框(n), "|")(1)
                            .IgnoreBlank = True
                            .InCellDropdown = True
                            .InputTitle = ""
                            .ErrorTitle = "输入提醒"
                            .InputMessage = ""
                            .ErrorMessage = "必须从下拉列表中选择"
                            .IMEMode = 0
                            .ShowInput = True
                            .ShowError = True
                        End With
                        
                        .Rows("1:2").Select
                        With mobjXLS.Selection.Validation
                            .Delete
                        End With
                    End If
                Next
            Next
        End If
        .Range("A3").Select
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SaveCard()
    '保存数据
    Dim cbrControl As CommandBarControl
    Dim lngRow As Long
    Dim lngCol As Long
    
    On Error GoTo ErrHand
    
    If vsfList.Rows = 1 Then Exit Sub
    '判断导入方式
    With vsfError
        If .Rows > 1 Then
            For lngRow = 1 To .Rows - 1
                If mbyt导入方式 = 1 Then
                    If .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(2).Picture Or .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(1).Picture Then
                        MsgBox "不能存在任何不合格的数据，请修正！", vbInformation, gstrSysName
                        Exit Sub
                    End If
                Else
                    If .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(2).Picture Or .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(1).Picture Then
                        If MsgBox("还存在不合格数据，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Exit Sub
                        Else
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
    End With
    '保存
    If TabControl.Selected.Caption = "分类" Then
        Call SaveType
    Else
        Call SaveMedi
    End If
    
    Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLSAVE, , True)
    cbrControl.Enabled = False
    
    Exit Sub
ErrHand:
    If Not mobjWB Is Nothing Then
        mobjWB.Close
    End If
    Set mobjWB = Nothing
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim cbrControl As CommandBarControl
    Dim rsTemp     As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    mlngModule = glngModul
    Me.Height = 600 * 15
    Me.Width = 800 * 15
    lblCollect.Caption = ""
    
    Call InitComandbar
    Call InitTabControl
    Call GetColumnHead
    Call InitVsf
        
    If vsfList.Rows <= 1 Then
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLOUTPUT, , True)
        cbrControl.Enabled = False
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLCHECK, , True)
        cbrControl.Enabled = False
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLSAVE, , True)
        cbrControl.Enabled = False
    End If
    
    Exit Sub
ErrHand:
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub InitVsf()
    '初始化vsf表格控件
    With vsfList
        .Rows = 1
        .Cols = 16
        .Editable = flexEDNone
        .ExplorerBar = flexExNone   '列不支持排序和拖动
    End With
    
    With vsfError
        .Rows = 1
        .Cols = 4
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 300
        .TextMatrix(0, 1) = "错误位置"
        .ColWidth(1) = 2000
        .TextMatrix(0, 2) = "错误类型"
        .ColWidth(2) = 2000
        .TextMatrix(0, 3) = "错误原因"
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .ScrollBars = flexScrollBarBoth
        '.ExtendLastCol = True '最后一列填充满
        .ExplorerBar = flexExNone   '列不支持排序和拖动
        .AllowUserResizing = flexResizeNone
    End With
End Sub

Private Sub InitExcel()
    '初始化Excel表格
    Set mobjXLS = CreateObject("Excel.Application")
    mobjXLS.DisplayAlerts = False
End Sub

Private Function GetExcelData() As Boolean
    '获取excel表格数据，并将其以字符串形式保存到
    '返回true-有错误 返回false-没有错误
    Dim strFileColumn As String    '文件中列名称
    Dim blnNotNullRow As Boolean   '检查该行是不是空行
    Dim cbrControl    As CommandBarControl
    Dim lngRow        As Long
    Dim lngCol        As Long
    Dim rsTemp        As Recordset
    Dim strSql        As String
    Dim i             As Integer
    
    On Error GoTo ErrHand
    
    vsfList.Clear
    lblCollect.Caption = ""
    vsfError.Rows = 1

    If txtFile.Text = "" Then Exit Function
    
    If mobjXLS Is Nothing Then Call InitExcel
    Set mobjWB = mobjXLS.Workbooks.Open(txtFile.Text)
    
    '检查分类列头
    Set mobjWS = mobjWB.Sheets(1)
    If mobjWS Is Nothing Then Exit Function
    For lngCol = 1 To mobjWS.UsedRange.Columns.Count
        strFileColumn = strFileColumn & Trim(mobjWS.UsedRange.Cells(2, lngCol)) & "|"
    Next
    For lngCol = 0 To UBound(Split(mstrType, "|")) - 1
        If InStr(1, "|" & strFileColumn, "|" & Split(mstrType, "|")(lngCol) & "|") = 0 Then
            vsfError.Rows = vsfError.Rows + 1
            vsfError.TextMatrix(vsfError.Rows - 1, 1) = "分类页"
            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "列头错误"
            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "列头【" & Split(mstrType, "|")(lngCol) & "】为显示列，然导入文件中不存在该列，请修正要导入的Excel文件！"
            GetExcelData = True
        End If
    Next
    '检查明细列头
    strFileColumn = ""
    Set mobjWS = Nothing
    Set mobjWS = mobjWB.Sheets(2)
    If mobjWS Is Nothing Then Exit Function
    For lngCol = 1 To mobjWS.UsedRange.Columns.Count
        strFileColumn = strFileColumn & Trim(mobjWS.UsedRange.Cells(2, lngCol)) & "|"
    Next
    For lngCol = 0 To UBound(Split(mstrMedi, "|")) - 1
        If InStr(1, "|" & strFileColumn, "|" & Split(mstrMedi, "|")(lngCol) & "|") = 0 Then
            vsfError.Rows = vsfError.Rows + 1
            vsfError.TextMatrix(vsfError.Rows - 1, 1) = "明细页"
            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "列头错误"
            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "列头【" & Split(mstrMedi, "|")(lngCol) & "】为显示列，然导入文件中不存在该列，请修正要导入的Excel文件！"
            GetExcelData = True
        End If
    Next
    
    If GetExcelData = True Then Exit Function
    
    '添加分类数据
    Set mobjWS = Nothing
    Set mobjWS = mobjWB.Sheets(1)
    If mobjWS Is Nothing Then Exit Function
    With mobjWS.UsedRange
        vsfList.Redraw = flexRDNone
        vsfList.Cols = UBound(Split(mstrType, "|")) + 1
        vsfList.Rows = 1
        
        For lngCol = 1 To UBound(Split(mstrType, "|"))
            vsfList.ColKey(lngCol) = Split(mstrType, "|")(lngCol - 1)
            vsfList.TextMatrix(0, lngCol) = Split(mstrType, "|")(lngCol - 1)
        Next
        
        For i = 1 To vsfList.Cols - 1
            For lngCol = 1 To .Columns.Count
                If vsfList.ColKey(i) = Trim(.Cells(2, lngCol)) Then
                    For lngRow = 3 To .Rows.Count
                        If i = 1 Then
                            vsfList.Rows = vsfList.Rows + 1
                        End If
                        vsfList.TextMatrix(lngRow - 2, i) = Trim(.Cells(lngRow, lngCol))
                    Next
                    Exit For
                End If
            Next
        Next
    End With
    Call GetColumns("分类")
    
    '添加明细数据
    vsfList.Clear
    Set mobjWS = Nothing
    Set mobjWS = mobjWB.Sheets(2)
    If mobjWS Is Nothing Then Exit Function
    With mobjWS.UsedRange
        vsfList.Redraw = flexRDNone
        vsfList.Cols = UBound(Split(mstrMedi, "|")) + 1
        vsfList.Rows = 1
        
        For lngCol = 1 To UBound(Split(mstrMedi, "|"))
            vsfList.ColKey(lngCol) = Split(mstrMedi, "|")(lngCol - 1)
            vsfList.TextMatrix(0, lngCol) = Split(mstrMedi, "|")(lngCol - 1)
        Next
        
        For i = 1 To vsfList.Cols - 1
            For lngCol = 1 To .Columns.Count
                If vsfList.ColKey(i) = Trim(.Cells(2, lngCol)) Then
                    For lngRow = 3 To .Rows.Count
                        If i = 1 Then
                            vsfList.Rows = vsfList.Rows + 1
                        End If
                        vsfList.TextMatrix(lngRow - 2, i) = Trim(.Cells(lngRow, lngCol))
                    Next
                    Exit For
                End If
            Next
        Next
    End With
    Call GetColumns("明细")
    
    If mstrTypeMsg <> "" And TabControl.Selected.Caption = "分类" Then
        Call SetColumns("分类")
        If vsfList.Rows > 1 Then Call CheckKind
    ElseIf mstrMediMsg <> "" And TabControl.Selected.Caption = "明细" Then
        Call SetColumns("明细")
        If vsfList.Rows > 1 Then Call Check品种: Call Check规格
    End If
    
    Set mobjWS = Nothing
    Set mobjWB = Nothing
    mobjXLS.Quit
    
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetColumns(ByVal strType As String)
    '将药品信息显示到界面表格中
    Dim blnNotNullRow As Boolean
    Dim cbrControl    As CommandBarControl
    Dim rsTemp        As Recordset
    Dim lngRow        As Long
    Dim lngCol        As Long
    Dim strSql        As String
    
    With vsfList
        .Tag = "1"
        .Clear
        .Redraw = flexRDDirect
        If strType = "分类" Then
            .Rows = UBound(Split(mstrTypeMsg, "|"))
            .Cols = UBound(Split(Split(mstrTypeMsg, "|")(0), ";")) + 1
            For lngRow = 0 To .Rows - 1
                For lngCol = 1 To .Cols - 1
                    If lngRow = 0 Then
                        .ColKey(lngCol) = Split(Split(mstrTypeMsg, "|")(lngRow), ";")(lngCol - 1)
                    End If
                    .TextMatrix(lngRow, lngCol) = Split(Split(mstrTypeMsg, "|")(lngRow), ";")(lngCol - 1)
                Next
            Next
        ElseIf strType = "明细" Then
            strSql = "select 内容,精度 from 药品卫材精度 where 类别=1 and 内容 in(1, 2) and 单位=1"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "")
            .Rows = UBound(Split(mstrMediMsg, "|"))
            .Cols = UBound(Split(Split(mstrMediMsg, "|")(0), ";")) + 1
            For lngRow = 0 To .Rows - 1
                For lngCol = 1 To .Cols - 1
                    If lngRow = 0 Then
                        .ColKey(lngCol) = Split(Split(mstrMediMsg, "|")(lngRow), ";")(lngCol - 1)
                    End If
                    Select Case .ColKey(lngCol)
                        Case "成本价", "售价"
                            rsTemp.Filter = ""
                            rsTemp.Filter = "内容=" & IIf(.ColKey(lngCol) = "成本价", 1, 2)
                            If IsNumeric(Split(Split(mstrMediMsg, "|")(lngRow), ";")(lngCol - 1)) Then
                                .TextMatrix(lngRow, lngCol) = zlStr.FormatEx(Split(Split(mstrMediMsg, "|")(lngRow), ";")(lngCol - 1), Val(rsTemp!精度), , True)
                            Else
                                .TextMatrix(lngRow, lngCol) = Split(Split(mstrMediMsg, "|")(lngRow), ";")(lngCol - 1)
                            End If
                        Case "供应商许可证效期"
                            If IsNumeric(Split(Split(mstrMediMsg, "|")(lngRow), ";")(lngCol - 1)) Then
                                .TextMatrix(lngRow, lngCol) = TranNumToDate(Split(Split(mstrMediMsg, "|")(lngRow), ";")(lngCol - 1))
                            Else
                                .TextMatrix(lngRow, lngCol) = FormatDate(Split(Split(mstrMediMsg, "|")(lngRow), ";")(lngCol - 1))
                            End If
                        Case "服务对象"
                            If IsNumeric(Split(Split(mstrMediMsg, "|")(lngRow), ";")(lngCol - 1)) Then
                                Select Case Split(Split(mstrMediMsg, "|")(lngRow), ";")(lngCol - 1)
                                    Case 0
                                        .TextMatrix(lngRow, lngCol) = "0-不服务于病人"
                                    Case 1
                                        .TextMatrix(lngRow, lngCol) = "1-门诊"
                                    Case 2
                                        .TextMatrix(lngRow, lngCol) = "2-住院"
                                    Case 3
                                        .TextMatrix(lngRow, lngCol) = "3-门诊和住院"
                                End Select
                            Else
                                .TextMatrix(lngRow, lngCol) = Split(Split(mstrMediMsg, "|")(lngRow), ";")(lngCol - 1)
                            End If
                        Case Else
                            .TextMatrix(lngRow, lngCol) = Split(Split(mstrMediMsg, "|")(lngRow), ";")(lngCol - 1)
                    End Select
                Next
            Next
        End If
        
        '将空行删除
        blnNotNullRow = True
        For lngRow = .Rows - 1 To 1 Step -1
            For lngCol = 1 To .Cols - 1
                If .TextMatrix(lngRow, lngCol) <> "" Then
                    blnNotNullRow = False
                End If
            Next
            '如果是空行将其删除
            If blnNotNullRow = True Then vsfList.RemoveItem lngRow
        Next
        
        If .Rows <= 1 Then
            Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLOUTPUT, , True)
            cbrControl.Enabled = False
            Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLCHECK, , True)
            cbrControl.Enabled = False
            Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLSAVE, , True)
            cbrControl.Enabled = False
        End If
        
        Call setVSF
        .Tag = "0"
    End With
End Sub

Private Function FormatDate(ByVal StrDate As String) As String
    '功能：格式化日期，返回用分号(-)分隔的日期格式
    Dim strYear  As String
    Dim strMonth As String
    Dim strDay   As String
    
    If LenB(StrConv(StrDate, vbFromUnicode)) >= 8 Then
        If InStr(1, StrDate, ".") > 0 Or InStr(1, StrDate, "/") > 0 Or InStr(1, StrDate, "-") > 0 Then
            StrDate = Replace(StrDate, ".", "")
            StrDate = Replace(StrDate, "/", "")
            StrDate = Replace(StrDate, "-", "")
        End If
        strYear = Mid(StrDate, 1, 4)
        If LenB(StrConv(StrDate, vbFromUnicode)) < 8 Then
            strMonth = Mid(StrDate, 5, 1)
        Else
            strMonth = Mid(StrDate, 5, 2)
        End If
        If LenB(StrConv(StrDate, vbFromUnicode)) < 8 Then
            strDay = Mid(StrDate, 6, 1)
        Else
            strDay = Mid(StrDate, 7, 2)
        End If
        If IsNumeric(strYear) = True And IsNumeric(strMonth) = True And IsNumeric(strDay) = True Then
            FormatDate = strYear & "-" & IIf(strMonth < 10, "0" & strMonth, strMonth) & "-" & IIf(strDay < 10, "0" & strDay, strDay)
        Else
            FormatDate = StrDate
        End If
    Else
        FormatDate = StrDate
    End If
End Function

Public Function TranNumToDate(ByVal strNum As String, Optional ByVal blnDec As Boolean = False) As String
    '转换数值为日期
    Dim strYear  As String
    Dim strMonth As String
    Dim strDay   As String
    Dim StrDate  As String
    
    TranNumToDate = ""
    If LenB(StrConv(strNum, vbFromUnicode)) < 4 Or LenB(StrConv(strNum, vbFromUnicode)) > 8 Then Exit Function
    
    strYear = Mid(strNum, 1, 4)
    strMonth = Mid(strNum, 5, 2)
    strDay = Mid(strNum, 7, 2)
        
    If strYear < 1000 Or strYear > 5000 Then Exit Function
    If strMonth = "" Then strMonth = "01"
    If strDay = "" Then strDay = "01"
    
    If strMonth > 12 Or strMonth < 1 Then Exit Function
    StrDate = strYear & "-" & strMonth & "-" & strDay
        
    If Not IsDate(StrDate) Then Exit Function
    
    StrDate = Format(StrDate, "yyyy-mm-dd")
    If blnDec Then StrDate = DateAdd("d", -1, Format(StrDate, "yyyy-mm-dd"))
    TranNumToDate = StrDate
End Function

Private Function GetColumnPostation(ByVal strColumn As String) As Integer
    '获取列位置和判断是否存在
    '参数 strcolumn-传入的列名
    '返回值 :返回传入列位置 0-没有找到 >0找到了
    Dim lngRow As Long
    Dim lngCol As Long
    
    With vsfList
        For lngCol = 1 To .Cols - 1
            If .TextMatrix(0, lngCol) = strColumn Then
                GetColumnPostation = lngCol
                Exit Function
            End If
        Next
        GetColumnPostation = 0
    End With
End Function

Private Sub CheckKind()
'检查分类数据合法性
    Dim cbrControl As CommandBarControl
    Dim rsTemp  As Recordset
    Dim lngRow  As Long
    Dim lngCol  As Long
    Dim strTemp As String
    Dim strSql  As String
    Dim j       As Integer
    Dim rs名称 As Recordset
    Dim strSqls As String
    
    On Error GoTo ErrHand
    
    Call ParseParameter
    
    strSql = "Select ID,Decode(Substr(编码, 1, 1), '0', Substr(编码, 2), 编码) As 编码,名称,上级ID,类型 From 诊疗分类目录 Where 类型 in (1,2,3)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "诊疗分类目录")
    
    vsfError.Rows = 1
    With vsfList
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = vbBlack '先设置成黑色
        For lngRow = 1 To .Rows - 1
            .TextMatrix(lngRow, 0) = lngRow '添加行标
            '类别
            If GetColumnPostation("类别") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("类别"))) <> "" Then
                    If Trim(.TextMatrix(lngRow, .ColIndex("类别"))) <> "西成药" And Trim(.TextMatrix(lngRow, .ColIndex("类别"))) <> "中成药" And Trim(.TextMatrix(lngRow, .ColIndex("类别"))) <> "中草药" Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("类别"), lngRow, .ColIndex("类别")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt分类类别 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【类别】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【类别】列只能是西成药、中成药、中草药！"
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("类别"), lngRow, .ColIndex("类别")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt分类类别 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【类别】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【类别】列不能为空！"
                End If
            End If
            '上级名称
            If GetColumnPostation("上级名称") > 0 And GetColumnPostation("类别") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("上级名称"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("上级名称"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("上级名称"), lngRow, .ColIndex("上级名称")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt上级名称 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【上级名称】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【上级名称】列不能有非法字符！"
                    Else
                        If GetTypeID(.TextMatrix(lngRow, .ColIndex("上级名称")), .TextMatrix(lngRow, .ColIndex("类别"))) = 0 Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("上级名称"), lngRow, .ColIndex("上级名称")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt上级名称 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【上级名称】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "诊疗分类目录中不存在类别为【" & .TextMatrix(lngRow, .ColIndex("类别")) & "】、名称为【" & .TextMatrix(lngRow, .ColIndex("上级名称")) & "】的项！"
                        End If
                    End If
                End If
            End If
            '编码
            If GetColumnPostation("编码") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("编码"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("编码"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("编码"), lngRow, .ColIndex("编码")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt编码 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【编码】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【编码】列不能有非法字符！"
                    Else
                        If LenB(StrConv(Trim(.TextMatrix(lngRow, .ColIndex("编码"))), vbFromUnicode)) > rsTemp("编码").DefinedSize Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("编码"), lngRow, .ColIndex("编码")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt编码 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【编码】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【编码】列字段长度不能超过数据库字段长度“" & rsTemp("编码").DefinedSize & "”！"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("编码"), lngRow, .ColIndex("编码")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt编码 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【编码】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【编码】列不能为空！"
                End If
            End If
            '名称
            If GetColumnPostation("名称") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("名称"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("名称"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("名称"), lngRow, .ColIndex("名称")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt名称 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【名称】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【名称】列不能有非法字符！"
                    Else
                        If LenB(StrConv(Trim(.TextMatrix(lngRow, .ColIndex("名称"))), vbFromUnicode)) > rsTemp("名称").DefinedSize Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("名称"), lngRow, .ColIndex("名称")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt名称 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【名称】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【名称】列字段长度不能超过数据库字段长度“" & rsTemp("名称").DefinedSize & "”！"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("名称"), lngRow, .ColIndex("名称")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt名称 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【名称】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【名称】列不能为空！"
                End If
            End If
            '编码和类别唯一
            If GetColumnPostation("类别") > 0 And GetColumnPostation("编码") > 0 Then
                If lngRow > 1 Then
                    For j = lngRow - 1 To 1 Step -1
                        If .TextMatrix(lngRow, .ColIndex("类别")) = .TextMatrix(j, .ColIndex("类别")) And .TextMatrix(lngRow, .ColIndex("编码")) = .TextMatrix(j, .ColIndex("编码")) And .TextMatrix(lngRow, .ColIndex("编码")) <> "" And .TextMatrix(lngRow, .ColIndex("类别")) <> "" Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("编码"), lngRow, .ColIndex("编码")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytNumAndKind = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【编码】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "唯一性错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "该条数据前面已存在类别为【" & Trim(.TextMatrix(lngRow, .ColIndex("类别"))) & "】、编码为【" & Trim(.TextMatrix(lngRow, .ColIndex("编码"))) & "】的数据，请检查！"
                        End If
                    Next
                End If
                If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("编码"))), "'") = 0 And Trim(.TextMatrix(lngRow, .ColIndex("类别"))) <> "" Then
                    rsTemp.Filter = ""
                    rsTemp.Filter = "类型=" & Switch(Trim(.TextMatrix(lngRow, .ColIndex("类别"))) = "西成药", 1, Trim(.TextMatrix(lngRow, .ColIndex("类别"))) = "中成药", 2, Trim(.TextMatrix(lngRow, .ColIndex("类别"))) = "中草药", 3) & " and 编码='" & IIf(Mid(Trim(.TextMatrix(lngRow, .ColIndex("编码"))), 1, 1) = 0, Mid(Trim(.TextMatrix(lngRow, .ColIndex("编码"))), 2), Trim(.TextMatrix(lngRow, .ColIndex("编码")))) & "'"
                    If rsTemp.RecordCount > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("编码"), lngRow, .ColIndex("编码")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytNumAndKind = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【编码】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "唯一性错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "导入项目与数据库中已有数据冲突！类别【" & Trim(.TextMatrix(lngRow, .ColIndex("类别"))) & "】下已存在编码【" & Trim(.TextMatrix(lngRow, .ColIndex("编码"))) & "】"
                    End If
                End If
            End If
            '名称、类别、上级名称唯一
            If GetColumnPostation("类别") > 0 And GetColumnPostation("上级名称") > 0 And GetColumnPostation("名称") > 0 Then
                If lngRow > 1 Then
                    For j = lngRow - 1 To 1 Step -1
                        If .TextMatrix(lngRow, .ColIndex("类别")) = .TextMatrix(j, .ColIndex("类别")) And .TextMatrix(lngRow, .ColIndex("上级名称")) = .TextMatrix(j, .ColIndex("上级名称")) And .TextMatrix(lngRow, .ColIndex("名称")) = .TextMatrix(j, .ColIndex("名称")) And .TextMatrix(lngRow, .ColIndex("类别")) <> "" And .TextMatrix(lngRow, .ColIndex("名称")) <> "" Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("名称"), lngRow, .ColIndex("名称")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytNumAndKind = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【名称】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "唯一性错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "该条数据前面已存在类别【" & Trim(.TextMatrix(lngRow, .ColIndex("类别"))) & "】、上级名称【" & Trim(.TextMatrix(lngRow, .ColIndex("上级名称"))) & "】、名称【" & Trim(.TextMatrix(lngRow, .ColIndex("名称"))) & "】的数据，请检查！"
                        End If
                    Next
                End If
                If GetTypeID(.TextMatrix(lngRow, .ColIndex("上级名称")), .TextMatrix(lngRow, .ColIndex("类别"))) >= 0 And InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("上级名称"))), "'") = 0 And InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("名称"))), "'") = 0 Then
                    strSqls = "Select ID,Decode(Substr(编码, 1, 1), '0', Substr(编码, 2), 编码) As 编码,名称,上级ID,类型 From 诊疗分类目录 Where 类型 in (1,2,3) and 名称=[1] and 上级ID" & IIf(GetTypeID(.TextMatrix(lngRow, .ColIndex("上级名称")), .TextMatrix(lngRow, .ColIndex("类别"))) = 0, " is null", "=" & GetTypeID(.TextMatrix(lngRow, .ColIndex("上级名称")), .TextMatrix(lngRow, .ColIndex("类别"))))
                    Set rs名称 = zlDatabase.OpenSQLRecord(strSqls, "诊疗项目目录", Trim(.TextMatrix(lngRow, .ColIndex("名称"))))
                    If rs名称.RecordCount > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("名称"), lngRow, .ColIndex("名称")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytNameAndKind = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【名称】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "唯一性错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "导入项目与数据库中已有数据冲突！类别【" & Trim(.TextMatrix(lngRow, .ColIndex("类别"))) & "】和上级名称【" & Trim(.TextMatrix(lngRow, .ColIndex("上级名称"))) & "】下已存在名称【" & Trim(.TextMatrix(lngRow, .ColIndex("名称"))) & "】"
                    End If
                End If
            End If
        Next
    End With
    
    Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLSAVE, , True)
    cbrControl.Enabled = True
    With vsfError
        If .Rows > 1 Then
            If mbyt导入方式 = 0 Then
                For lngRow = 1 To .Rows - 1
                    If .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(1).Picture Then
                        cbrControl.Enabled = False
                        Exit For
                    Else
                        cbrControl.Enabled = True
                    End If
                Next
            Else
                For lngRow = 1 To .Rows - 1
                    If .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(1).Picture Or .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(2).Picture Then
                        cbrControl.Enabled = False
                        Exit For
                    End If
                Next
            End If
        End If
    End With
    
    
    If vsfList.Rows > 1 Then
        vsfList.Row = 1: vsfList.Col = 1
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLOUTPUT, , True)
        cbrControl.Enabled = True
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLCHECK, , True)
        cbrControl.Enabled = True
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Check品种()
'检查明细数据（品种）合法性
    Dim cbrControl As CommandBarControl
    Dim rsTemp  As Recordset
    Dim rs剂型  As Recordset
    Dim lngRow  As Long
    Dim lngCol  As Long
    Dim strTemp As String
    Dim strSql  As String
    Dim j       As Integer
    Dim rs名称 As Recordset
    Dim strSqls As String

    On Error GoTo ErrHand
    
    Call ParseParameter
    
    strSql = "Select 类别,分类ID,ID,编码,名称,计算单位 From 诊疗项目目录 Where 类别 In ('5', '6', '7')"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "诊疗项目目录")
    Set rs剂型 = zlDatabase.OpenSQLRecord("select 名称 from 药品剂型", "药品剂型")
    
    vsfError.Rows = 1
    With vsfList
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = vbBlack '先设置成黑色
        For lngRow = 1 To .Rows - 1
            .TextMatrix(lngRow, 0) = lngRow '添加行标
            '类别
            If GetColumnPostation("类别") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("类别"))) = "" Then
                    .Cell(flexcpForeColor, lngRow, .ColIndex("类别"), lngRow, .ColIndex("类别")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt分类类别 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【类别】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【类别】列不能为空！"
                Else
                    If Trim(.TextMatrix(lngRow, .ColIndex("类别"))) <> "西成药" And Trim(.TextMatrix(lngRow, .ColIndex("类别"))) <> "中成药" And Trim(.TextMatrix(lngRow, .ColIndex("类别"))) <> "中草药" Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("类别"), lngRow, .ColIndex("类别")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt分类类别 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【类别】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【类别】列只能是西成药、中成药、中草药！"
                    End If
                End If
            End If
            '分类
            If GetColumnPostation("分类") > 0 And GetColumnPostation("类别") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("分类"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("分类"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("分类"), lngRow, .ColIndex("分类")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt药品分类 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【分类】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【分类】列不能有非法字符！"
                    Else
                        If GetTypeID(.TextMatrix(lngRow, .ColIndex("分类")), .TextMatrix(lngRow, .ColIndex("类别"))) = 0 Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("分类"), lngRow, .ColIndex("分类")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt药品分类 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【分类】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "诊疗分类目录中不存在类别为【" & .TextMatrix(lngRow, .ColIndex("类别")) & "】、分类为【" & .TextMatrix(lngRow, .ColIndex("分类")) & "】的项！"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("分类"), lngRow, .ColIndex("分类")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt药品分类 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【分类】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【分类】列不能为空！"
                End If
            End If
            '品种编码
            If GetColumnPostation("品种编码") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("品种编码"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("品种编码"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("品种编码"), lngRow, .ColIndex("品种编码")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt品种编码 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【品种编码】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【品种编码】列不能有非法字符！"
                    Else
                        If LenB(StrConv(Trim(.TextMatrix(lngRow, .ColIndex("品种编码"))), vbFromUnicode)) > rsTemp("编码").DefinedSize Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("品种编码"), lngRow, .ColIndex("品种编码")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt品种编码 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【品种编码】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【品种编码】列字段长度不能超过数据库字段长度“" & rsTemp("编码").DefinedSize & "”！"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("品种编码"), lngRow, .ColIndex("品种编码")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt品种编码 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【品种编码】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【品种编码】列不能为空！"
                End If
            End If
            '品种名称
            If GetColumnPostation("品种名称") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("品种名称"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("品种名称"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("品种名称"), lngRow, .ColIndex("品种名称")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt品种名称 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【品种名称】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【品种名称】列不能有非法字符！"
                    Else
                        If LenB(StrConv(Trim(.TextMatrix(lngRow, .ColIndex("品种名称"))), vbFromUnicode)) > rsTemp("名称").DefinedSize Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("品种名称"), lngRow, .ColIndex("品种名称")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt品种名称 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【品种名称】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【品种名称】列字段长度不能超过数据库字段长度“" & rsTemp("名称").DefinedSize & "”！"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("品种名称"), lngRow, .ColIndex("品种名称")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt品种名称 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【品种名称】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【品种名称】列不能为空！"
                End If
            End If
            '剂型
            If GetColumnPostation("剂型") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("剂型"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("剂型"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("剂型"), lngRow, .ColIndex("剂型")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt剂型 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【剂型】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【剂型】列不能有非法字符！"
                    Else
                        rs剂型.Filter = ""
                        rs剂型.Filter = "名称='" & Trim(.TextMatrix(lngRow, .ColIndex("剂型"))) & "'"
                        If LenB(StrConv(Trim(.TextMatrix(lngRow, .ColIndex("剂型"))), vbFromUnicode)) > rs剂型("名称").DefinedSize Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("剂型"), lngRow, .ColIndex("剂型")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt剂型 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【剂型】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【剂型】列字段长度不能超过数据库字段长度“" & rs剂型("名称").DefinedSize & "”！"
                        Else
                            If rs剂型.RecordCount = 0 Then
                                .Cell(flexcpForeColor, lngRow, .ColIndex("剂型"), lngRow, .ColIndex("剂型")) = vbRed
                                vsfError.Rows = vsfError.Rows + 1
                                vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt剂型 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                                vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【剂型】列"
                                vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                                vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【剂型】只能是数据库表中已有数据，剂型“" & Trim(.TextMatrix(lngRow, .ColIndex("剂型"))) & "”不存在！"
                            End If
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("剂型"), lngRow, .ColIndex("剂型")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt剂型 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【剂型】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【剂型】列不能为空！"
                End If
            End If
            '计量单位
            If GetColumnPostation("剂量单位") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("剂量单位"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("剂量单位"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("剂量单位"), lngRow, .ColIndex("剂量单位")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt单位 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【剂量单位】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【剂量单位】列不能有非法字符！"
                    Else
                        If LenB(StrConv(Trim(.TextMatrix(lngRow, .ColIndex("剂量单位"))), vbFromUnicode)) > rsTemp("计算单位").DefinedSize Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("剂量单位"), lngRow, .ColIndex("剂量单位")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt单位 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【剂量单位】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【剂量单位】列字段长度不能超过数据库字段长度“" & rsTemp("计算单位").DefinedSize & "”！"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("剂量单位"), lngRow, .ColIndex("剂量单位")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt单位 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【剂量单位】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【剂量单位】列不能为空！"
                End If
            End If
            '品种唯一性
            If GetColumnPostation("类别") > 0 And GetColumnPostation("分类") > 0 And GetColumnPostation("品种名称") Then
                If lngRow > 1 Then
                    For j = lngRow - 1 To 1 Step -1
                        If .TextMatrix(lngRow, .ColIndex("类别")) = .TextMatrix(j, .ColIndex("类别")) And .TextMatrix(lngRow, .ColIndex("分类")) = .TextMatrix(j, .ColIndex("分类")) And .TextMatrix(lngRow, .ColIndex("品种名称")) <> .TextMatrix(j, .ColIndex("品种名称")) And .TextMatrix(lngRow, .ColIndex("品种编码")) = .TextMatrix(j, .ColIndex("品种编码")) Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("品种名称"), lngRow, .ColIndex("品种名称")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt品种唯一 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【品种名称】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "唯一性错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "第【" & j & "】行和第【" & lngRow & "】行，同类别、同分类、同品种编码，品种名称【" & Trim(.TextMatrix(lngRow, .ColIndex("品种名称"))) & "】不相同，请检查！"
                        ElseIf .TextMatrix(lngRow, .ColIndex("类别")) = .TextMatrix(j, .ColIndex("类别")) And .TextMatrix(lngRow, .ColIndex("分类")) = .TextMatrix(j, .ColIndex("分类")) And .TextMatrix(lngRow, .ColIndex("品种名称")) = .TextMatrix(j, .ColIndex("品种名称")) And .TextMatrix(lngRow, .ColIndex("品种编码")) <> .TextMatrix(j, .ColIndex("品种编码")) Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("品种编码"), lngRow, .ColIndex("品种编码")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt品种唯一 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【品种编码】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "唯一性错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "第【" & j & "】行和第【" & lngRow & "】行，同类别、同分类、同品种名称、品种编码【" & Trim(.TextMatrix(lngRow, .ColIndex("品种编码"))) & "】不相同，请检查！"
                        End If
                    Next
                End If
            End If
            '在已有分类下建立规格的检查
            If GetColumnPostation("类别") > 0 And GetColumnPostation("分类") > 0 And GetColumnPostation("品种名称") > 0 And GetColumnPostation("品种编码") > 0 Then
                If GetTypeID(Trim(.TextMatrix(lngRow, .ColIndex("分类"))), .TextMatrix(lngRow, .ColIndex("类别"))) > 0 And InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("品种名称"))), "'") = 0 And InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("分类"))), "'") = 0 And InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("品种编码"))), "'") = 0 Then
                    rsTemp.Filter = ""
                    rsTemp.Filter = "编码='" & Trim(.TextMatrix(lngRow, .ColIndex("品种编码"))) & "'"
                    If rsTemp.RecordCount > 0 Then '如果数据库存在界面录入的【品种编码】，就检查界面录入分类下【品种编码】【品种名称】与已有是否一致
                        strSqls = "Select 类别,分类ID,ID,编码,名称,计算单位 From 诊疗项目目录 Where 类别 In ('5', '6', '7') and 分类ID=[1] and 名称=[2] and 编码=[3] "
                        Set rs名称 = zlDatabase.OpenSQLRecord(strSqls, "诊疗项目目录", GetTypeID(Trim(.TextMatrix(lngRow, .ColIndex("分类"))), .TextMatrix(lngRow, .ColIndex("类别"))), Trim(.TextMatrix(lngRow, .ColIndex("品种名称"))), Trim(.TextMatrix(lngRow, .ColIndex("品种编码"))))
                        If rs名称.RecordCount = 0 Then '如果数据库存在界面录入的【品种编码】，且界面录入分类下【品种编码】【品种名称】与已有不一致
                            .Cell(flexcpForeColor, lngRow, .ColIndex("品种编码"), lngRow, .ColIndex("品种编码")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt品种唯一 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【品种编码】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "唯一性错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "导入项目与数据库中已有数据冲突！编码【" & Trim(.TextMatrix(lngRow, .ColIndex("品种编码"))) & "】已存在！"
                        End If
                    Else '如果数据库不存在界面录入的【品种编码】，就检查界面录入分类下【品种名称】与已有是否一致
                        strSqls = "Select 类别,分类ID,ID,编码,名称,计算单位 From 诊疗项目目录 Where 类别 In ('5', '6', '7') and 分类ID=[1] and 名称=[2]  "
                        Set rs名称 = zlDatabase.OpenSQLRecord(strSqls, "诊疗项目目录", GetTypeID(Trim(.TextMatrix(lngRow, .ColIndex("分类"))), .TextMatrix(lngRow, .ColIndex("类别"))), Trim(.TextMatrix(lngRow, .ColIndex("品种名称"))))
                        If rs名称.RecordCount > 0 Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("品种名称"), lngRow, .ColIndex("品种名称")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt品种唯一 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【品种名称】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "唯一性错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "导入项目与数据库中已有数据是否冲突！类别【" & Trim(.TextMatrix(lngRow, .ColIndex("类别"))) & "】和分类【" & Trim(.TextMatrix(lngRow, .ColIndex("分类"))) & "】下，品种【" & Trim(.TextMatrix(lngRow, .ColIndex("品种名称"))) & "】已存在！"
                        End If
                    End If
                End If
            End If
        Next
    End With
    
    Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLSAVE, , True)
    cbrControl.Enabled = True
    With vsfError
        If .Rows > 1 Then
            If mbyt导入方式 = 0 Then
                For lngRow = 1 To .Rows - 1
                    If .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(1).Picture Then
                        cbrControl.Enabled = False
                        Exit For
                    Else
                        cbrControl.Enabled = True
                    End If
                Next
            Else
                For lngRow = 1 To .Rows - 1
                    If .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(1).Picture Or .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(2).Picture Then
                        cbrControl.Enabled = False
                        Exit For
                    End If
                Next
            End If
        End If
    End With
    
    If vsfList.Rows > 1 Then
        vsfList.Row = 1: vsfList.Col = 1
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLOUTPUT, , True)
        cbrControl.Enabled = True
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLCHECK, , True)
        cbrControl.Enabled = True
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Check规格()
'检查明细数据（规格）合法性
    Dim cbrControl As CommandBarControl
    Dim rs收入项目 As Recordset
    Dim rs供应商   As Recordset
    Dim rsTemp     As Recordset
    Dim rs精度     As Recordset
    Dim lngRow     As Long
    Dim lngCol     As Long
    Dim strTemp    As String
    Dim strSql     As String
    Dim j          As Integer
    Dim rs名称 As Recordset
    Dim strSqls As String
    
    On Error GoTo ErrHand
    
    Call ParseParameter
    
    strSql = "Select a.类别, a.Id, a.编码, a.名称, a.规格, a.产地, a.计算单位, b.剂量系数, b.门诊单位, b.门诊包装, b.住院单位, b.住院包装, b.药库单位, b.药库包装," & vbNewLine & _
             "b.最大效期, b.住院可否分零, b.药库分批, b.药房分批, b.成本价, b.合同单位id, b.门诊可否分零" & vbNewLine & _
             "From 收费项目目录 A, 药品规格 B Where a.Id = b.药品id And a.类别 In ('5', '6', '7')"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "药品规格")
    Set rs精度 = zlDatabase.OpenSQLRecord("select 类别,内容,单位,精度 from 药品卫材精度 where 类别=1", "价格精度")
    Set rs收入项目 = zlDatabase.OpenSQLRecord("Select ID,编码,名称 From 收入项目 Where 末级 = 1", "收入项目")
    Set rs供应商 = zlDatabase.OpenSQLRecord("Select ID,编码,名称,许可证号,许可证效期 From 供应商", "供应商")
    
    With vsfList
        For lngRow = 1 To .Rows - 1
            '规格编码
            If GetColumnPostation("规格编码") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("规格编码"))) <> "" Then
                      If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("规格编码"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("规格编码"), lngRow, .ColIndex("规格编码")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt规格编码 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【规格编码】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【规格编码】列不能有非法字符！"
                    Else
                        If lngRow > 1 Then
                            For j = lngRow - 1 To 1 Step -1
                                If .TextMatrix(lngRow, .ColIndex("规格编码")) = .TextMatrix(j, .ColIndex("规格编码")) Then
                                    .Cell(flexcpForeColor, lngRow, .ColIndex("规格编码"), lngRow, .ColIndex("规格编码")) = vbRed
                                    vsfError.Rows = vsfError.Rows + 1
                                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt规格唯一 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【规格编码】列"
                                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "唯一性错误"
                                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "该条数据前面已存在规格编码为【" & Trim(.TextMatrix(lngRow, .ColIndex("规格编码"))) & "】的数据，请检查！"
                                End If
                            Next
                        End If
                        rsTemp.Filter = ""
                        rsTemp.Filter = "编码='" & Trim(.TextMatrix(lngRow, .ColIndex("规格编码"))) & "'"
                        If rsTemp.RecordCount > 0 Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("规格编码"), lngRow, .ColIndex("规格编码")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt规格唯一 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【规格编码】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "唯一性错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "导入项目与数据库中已有数据冲突！编码【" & Trim(.TextMatrix(lngRow, .ColIndex("规格编码"))) & "】已存在！"
                        Else
                            If LenB(StrConv(Trim(.TextMatrix(lngRow, .ColIndex("规格编码"))), vbFromUnicode)) > rsTemp("编码").DefinedSize Then
                                .Cell(flexcpForeColor, lngRow, .ColIndex("规格编码"), lngRow, .ColIndex("规格编码")) = vbRed
                                vsfError.Rows = vsfError.Rows + 1
                                vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt规格编码 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                                vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【规格编码】列"
                                vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                                vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【规格编码】列字段长度不能超过数据库字段长度“" & rsTemp("编码").DefinedSize & "”！"
                            End If
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("规格编码"), lngRow, .ColIndex("规格编码")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt规格编码 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【规格编码】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【规格编码】列不能为空！"
                End If
            End If
            '药品规格
            If GetColumnPostation("药品规格") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("药品规格"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("药品规格"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("药品规格"), lngRow, .ColIndex("药品规格")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt药品规格 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【药品规格】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【药品规格】列不能有非法字符！"
                    Else
                        If LenB(StrConv(Trim(.TextMatrix(lngRow, .ColIndex("药品规格"))), vbFromUnicode)) > rsTemp("规格").DefinedSize Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("药品规格"), lngRow, .ColIndex("药品规格")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt药品规格 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【药品规格】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【药品规格】列字段长度不能超过数据库字段长度“" & rsTemp("规格").DefinedSize & "”！"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("药品规格"), lngRow, .ColIndex("药品规格")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt药品规格 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【药品规格】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【药品规格】列不能为空！"
                End If
            End If
            '生产商
            If GetColumnPostation("生产商") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("生产商"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("生产商"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("生产商"), lngRow, .ColIndex("生产商")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt产地 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【生产商】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【生产商】列不能有非法字符！"
                    Else
                        If LenB(StrConv(Trim(.TextMatrix(lngRow, .ColIndex("生产商"))), vbFromUnicode)) > rsTemp("产地").DefinedSize Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("生产商"), lngRow, .ColIndex("生产商")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt产地 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【生产商】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【生产商】列字段长度不能超过数据库字段长度“" & rsTemp("产地").DefinedSize & "”！"
                        End If
                    End If
                End If
            End If
            '售价单位
            If GetColumnPostation("售价单位") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("售价单位"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("售价单位"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("售价单位"), lngRow, .ColIndex("售价单位")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt单位 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【售价单位】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【售价单位】列不能有非法字符！"
                    Else
                        If LenB(StrConv(Trim(.TextMatrix(lngRow, .ColIndex("售价单位"))), vbFromUnicode)) > rsTemp("计算单位").DefinedSize Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("售价单位"), lngRow, .ColIndex("售价单位")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt单位 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【售价单位】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【售价单位】列字段长度不能超过数据库字段长度“" & rsTemp("计算单位").DefinedSize & "”！"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("售价单位"), lngRow, .ColIndex("售价单位")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt单位 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【售价单位】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【售价单位】列不能为空！"
                End If
            End If
            '售价换算系数
            If GetColumnPostation("售价换算系数") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("售价换算系数"))) <> "" Then
                    If Not IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("售价换算系数")))) Or Val(Trim(.TextMatrix(lngRow, .ColIndex("售价换算系数")))) <= 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("售价换算系数"), lngRow, .ColIndex("售价换算系数")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt换算系数 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【售价换算系数】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【售价换算系数】列只能由数字0-9组成且>0！"
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("售价换算系数"), lngRow, .ColIndex("售价换算系数")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt换算系数 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【售价换算系数】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【售价换算系数】列不能为空！"
                End If
            End If
            '门诊单位
            If GetColumnPostation("门诊单位") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("门诊单位"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("门诊单位"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("门诊单位"), lngRow, .ColIndex("门诊单位")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt单位 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【门诊单位】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【门诊单位】列不能有非法字符！"
                    Else
                        If LenB(StrConv(Trim(.TextMatrix(lngRow, .ColIndex("门诊单位"))), vbFromUnicode)) > rsTemp("门诊单位").DefinedSize Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("门诊单位"), lngRow, .ColIndex("门诊单位")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt单位 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【门诊单位】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【门诊单位】列字段长度不能超过数据库字段长度“" & rsTemp("门诊单位").DefinedSize & "”！"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("门诊单位"), lngRow, .ColIndex("门诊单位")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt单位 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【门诊单位】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【门诊单位】列不能为空！"
                End If
            End If
            '门诊换算系数
            If GetColumnPostation("门诊换算系数") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("门诊换算系数"))) <> "" Then
                    If Not IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("门诊换算系数")))) Or Val(Trim(.TextMatrix(lngRow, .ColIndex("门诊换算系数")))) <= "0" Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("门诊换算系数"), lngRow, .ColIndex("门诊换算系数")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt换算系数 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【门诊换算系数】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【门诊换算系数】列只能由数字0-9组成且>0！"
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("门诊换算系数"), lngRow, .ColIndex("门诊换算系数")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt换算系数 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【门诊换算系数】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【门诊换算系数】列不能为空！"
                End If
            End If
            '住院单位
            If GetColumnPostation("住院单位") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("住院单位"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("住院单位"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("住院单位"), lngRow, .ColIndex("住院单位")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt单位 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【住院单位】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【住院单位】列不能有非法字符！"
                    Else
                        If LenB(StrConv(Trim(.TextMatrix(lngRow, .ColIndex("住院单位"))), vbFromUnicode)) > rsTemp("住院单位").DefinedSize Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("住院单位"), lngRow, .ColIndex("住院单位")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt单位 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【住院单位】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【住院单位】列字段长度不能超过数据库字段长度“" & rsTemp("住院单位").DefinedSize & "”！"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("住院单位"), lngRow, .ColIndex("住院单位")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt单位 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【住院单位】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【住院单位】列不能为空！"
                End If
            End If
            '住院换算系数
            If GetColumnPostation("住院换算系数") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("住院换算系数"))) <> "" Then
                    If Not IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("住院换算系数")))) Or Val(Trim(.TextMatrix(lngRow, .ColIndex("住院换算系数")))) <= "0" Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("住院换算系数"), lngRow, .ColIndex("住院换算系数")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt换算系数 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【住院换算系数】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【住院换算系数】列只能由数字0-9组成且>0！"
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("住院换算系数"), lngRow, .ColIndex("住院换算系数")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt换算系数 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【住院换算系数】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【住院换算系数】列不能为空！"
                End If
            End If
            '药库单位
            If GetColumnPostation("药库单位") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("药库单位"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("药库单位"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("药库单位"), lngRow, .ColIndex("药库单位")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt单位 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【药库单位】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【药库单位】列不能有非法字符！"
                    Else
                        If LenB(StrConv(Trim(.TextMatrix(lngRow, .ColIndex("药库单位"))), vbFromUnicode)) > rsTemp("药库单位").DefinedSize Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("药库单位"), lngRow, .ColIndex("药库单位")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt单位 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【药库单位】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【药库单位】列字段长度不能超过数据库字段长度“" & rsTemp("药库单位").DefinedSize & "”！"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("药库单位"), lngRow, .ColIndex("药库单位")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt单位 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【药库单位】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【药库单位】列不能为空！"
                End If
            End If
            '药库换算系数
            If GetColumnPostation("药库换算系数") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("药库换算系数"))) <> "" Then
                    If Not IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("药库换算系数")))) Or Val(Trim(.TextMatrix(lngRow, .ColIndex("药库换算系数")))) <= "0" Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("药库换算系数"), lngRow, .ColIndex("药库换算系数")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt换算系数 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【药库换算系数】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【药库换算系数】列只能由数字0-9组成且>0！"
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("药库换算系数"), lngRow, .ColIndex("药库换算系数")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt换算系数 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【药库换算系数】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【药库换算系数】列不能为空！"
                End If
            End If
            '门诊、住院、药库单位相同，换算系数比较
            If GetColumnPostation("门诊单位") > 0 And GetColumnPostation("住院单位") > 0 And GetColumnPostation("门诊换算系数") > 0 And GetColumnPostation("住院换算系数") > 0 And GetColumnPostation("药库单位") > 0 And GetColumnPostation("药库换算系数") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("门诊单位"))) = Trim(.TextMatrix(lngRow, .ColIndex("住院单位"))) Then
                    If Trim(.TextMatrix(lngRow, .ColIndex("门诊换算系数"))) <> Trim(.TextMatrix(lngRow, .ColIndex("住院换算系数"))) Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("住院换算系数"), lngRow, .ColIndex("住院换算系数")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt换算系数 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【住院换算系数】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "门诊、住院单位相同，换算系数必须相同！"
                    End If
                End If
                If Trim(.TextMatrix(lngRow, .ColIndex("门诊单位"))) = Trim(.TextMatrix(lngRow, .ColIndex("药库单位"))) Then
                    If Trim(.TextMatrix(lngRow, .ColIndex("门诊换算系数"))) <> Trim(.TextMatrix(lngRow, .ColIndex("药库换算系数"))) Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("药库换算系数"), lngRow, .ColIndex("药库换算系数")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt换算系数 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【药库换算系数】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "门诊、药库单位相同，换算系数必须相同！"
                    End If
                End If
            End If
            '门诊、住院、药库换算系数相同，单位比较
            If GetColumnPostation("门诊单位") > 0 And GetColumnPostation("住院单位") > 0 And GetColumnPostation("门诊换算系数") > 0 And GetColumnPostation("住院换算系数") > 0 And GetColumnPostation("药库单位") > 0 And GetColumnPostation("药库换算系数") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("门诊换算系数"))) = Trim(.TextMatrix(lngRow, .ColIndex("住院换算系数"))) Then
                    If Trim(.TextMatrix(lngRow, .ColIndex("门诊单位"))) <> Trim(.TextMatrix(lngRow, .ColIndex("住院单位"))) Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("住院单位"), lngRow, .ColIndex("住院单位")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt换算系数 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【住院单位】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "门诊、住院换算系数相同，单位必须相同！"
                    End If
                End If
                If Trim(.TextMatrix(lngRow, .ColIndex("门诊换算系数"))) = Trim(.TextMatrix(lngRow, .ColIndex("药库换算系数"))) Then
                    If Trim(.TextMatrix(lngRow, .ColIndex("门诊单位"))) <> Trim(.TextMatrix(lngRow, .ColIndex("药库单位"))) Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("药库单位"), lngRow, .ColIndex("药库单位")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt换算系数 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【药库单位】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "门诊、药库换算系数相同，单位必须相同！"
                    End If
                End If
            End If
            '是否变价
            If GetColumnPostation("是否变价") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("是否变价"))) <> "" Then
                    If Trim(.TextMatrix(lngRow, .ColIndex("是否变价"))) <> "√" Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("是否变价"), lngRow, .ColIndex("是否变价")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt变价 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【是否变价】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【是否变价】列只能是“√”或空！"
                    End If
                End If
            End If
            '成本价
            If GetColumnPostation("成本价") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("成本价"))) <> "" Then
                    If Not IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("成本价")))) Or Val(Trim(.TextMatrix(lngRow, .ColIndex("成本价")))) < 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("成本价"), lngRow, .ColIndex("成本价")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt价格 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【成本价】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值类型错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【成本价】列只能由数字组成且不小于0！"
                    Else
                        rs精度.Filter = ""
                        rs精度.Filter = "内容=1 and 单位=1"
                        If LenB(StrConv(Mid(Trim(.TextMatrix(lngRow, .ColIndex("成本价"))), InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("成本价"))), ".") + 1), vbFromUnicode)) > rs精度!精度 Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("成本价"), lngRow, .ColIndex("成本价")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt价格 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【成本价】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【成本价】列字段精度不能超过最大设置精度！"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("成本价"), lngRow, .ColIndex("成本价")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt价格 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【成本价】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【成本价】列不能为空！"
                End If
            End If
            '售价
            If GetColumnPostation("售价") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("售价"))) <> "" Then
                    If Not IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("售价")))) Or Val(Trim(.TextMatrix(lngRow, .ColIndex("售价")))) < 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("售价"), lngRow, .ColIndex("售价")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt价格 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【售价】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值类型错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【售价】列只能由数字组成且不小于0！"
                    Else
                        rs精度.Filter = ""
                        rs精度.Filter = "内容=2 and 单位=1"
                        If LenB(StrConv(Mid(Trim(.TextMatrix(lngRow, .ColIndex("售价"))), InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("售价"))), ".") + 1), vbFromUnicode)) > rs精度!精度 Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("售价"), lngRow, .ColIndex("售价")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt价格 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【售价】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【售价】列字段精度不能超过最大设置精度！"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("售价"), lngRow, .ColIndex("售价")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt价格 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【售价】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【售价】列不能为空！"
                End If
            End If
            '效期(月)
            If GetColumnPostation("效期(月)") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("效期(月)"))) <> "" Then
                    If Not IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("效期(月)")))) Or InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("效期(月)"))), ".") > 0 Or Val(Trim(.TextMatrix(lngRow, .ColIndex("效期(月)")))) < 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("效期(月)"), lngRow, .ColIndex("效期(月)")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt效期 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【效期(月)】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值类型错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【效期(月)】列只能是整数且不小于0！"
                    End If
                End If
            End If
            '收入项目
            If GetColumnPostation("收入项目") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("收入项目"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("收入项目"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("收入项目"), lngRow, .ColIndex("收入项目")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt收入项目 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【收入项目】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【收入项目】列不能有非法字符！"
                    Else
                        rs收入项目.Filter = ""
                        rs收入项目.Filter = "名称='" & Trim(.TextMatrix(lngRow, .ColIndex("收入项目"))) & "'"
                        If rs收入项目.RecordCount = 0 Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("收入项目"), lngRow, .ColIndex("收入项目")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt收入项目 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【收入项目】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【收入项目】列只能是数据库已有收入项目！收入项目【" & Trim(.TextMatrix(lngRow, .ColIndex("收入项目"))) & "】不存在！"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("收入项目"), lngRow, .ColIndex("收入项目")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt收入项目 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【收入项目】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【收入项目】列不能为空！"
                End If
            End If
            '住院可分零
            If GetColumnPostation("住院可否分零") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("住院可否分零"))) <> "" Then
                    If InStr(1, ",0-可以分零,1-不可分零,2-一次性使用,3-分零后一天内有效,4-分零后两天内有效,5-分零后三天内有效,", "," & Trim(.TextMatrix(lngRow, .ColIndex("住院可否分零"))) & ",") = 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("住院可否分零"), lngRow, .ColIndex("住院可否分零")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt分零 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【住院可否分零】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【住院可否分零】列只能是已有分零方式！分零方式【" & Trim(.TextMatrix(lngRow, .ColIndex("住院可否分零"))) & "】不存在！"
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("住院可否分零"), lngRow, .ColIndex("住院可否分零")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt分零 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【住院可否分零】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【住院可否分零】列不能为空！"
                End If
            End If
            '门诊可分零
            If GetColumnPostation("门诊可否分零") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("门诊可否分零"))) <> "" Then
                    If InStr(1, ",0-可以分零,1-不可分零,2-一次性使用,3-分零后一天内有效,4-分零后两天内有效,5-分零后三天内有效,", "," & Trim(.TextMatrix(lngRow, .ColIndex("门诊可否分零"))) & ",") = 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("门诊可否分零"), lngRow, .ColIndex("门诊可否分零")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt分零 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【门诊可否分零】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【门诊可否分零】列只能是已有分零方式！分零方式【" & Trim(.TextMatrix(lngRow, .ColIndex("门诊可否分零"))) & "】不存在！"
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("门诊可否分零"), lngRow, .ColIndex("门诊可否分零")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt分零 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【门诊可否分零】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【门诊可否分零】列不能为空！"
                End If
            End If
            '服务对象
            If GetColumnPostation("服务对象") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("服务对象"))) <> "" Then
                    If InStr(1, ",0-不服务于病人,1-门诊,2-住院,3-门诊和住院,", "," & Trim(.TextMatrix(lngRow, .ColIndex("服务对象"))) & ",") = 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("服务对象"), lngRow, .ColIndex("服务对象")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt服务对象 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【服务对象】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【服务对象】列只能是已有服务对象！服务对象【" & Trim(.TextMatrix(lngRow, .ColIndex("服务对象"))) & "】不存在！"
                    End If
                End If
            End If
            '药库分批
            If GetColumnPostation("药库分批") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("药库分批"))) <> "" Then
                    If Trim(.TextMatrix(lngRow, .ColIndex("药库分批"))) <> "√" Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("药库分批"), lngRow, .ColIndex("药库分批")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt分批属性 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【药库分批】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【药库分批】列只能是“√”或空！"
                    End If
                End If
            End If
            '药房分批
            If GetColumnPostation("药房分批") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("药房分批"))) <> "" Then
                    If Trim(.TextMatrix(lngRow, .ColIndex("药房分批"))) <> "√" Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("药房分批"), lngRow, .ColIndex("药房分批")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt分批属性 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【药房分批】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【药房分批】列只能是“√”或空！"
                    End If
                    If GetColumnPostation("药库分批") > 0 Then
                        If Trim(.TextMatrix(lngRow, .ColIndex("药库分批"))) = "" Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("药房分批"), lngRow, .ColIndex("药房分批")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt分批属性 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【药房分批】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "药库分批时药房才能分批！"
                        End If
                    End If
                End If
            End If
            '供应商名称
            If GetColumnPostation("供应商名称") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("供应商名称"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("供应商名称"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("供应商名称"), lngRow, .ColIndex("供应商名称")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt供应商 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【供应商名称】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【供应商名称】列不能有非法字符！"
                    Else
                        rs供应商.Filter = ""
                        rs供应商.Filter = "名称='" & Trim(.TextMatrix(lngRow, .ColIndex("供应商名称"))) & "'"
                        If rs供应商.RecordCount = 0 Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("供应商名称"), lngRow, .ColIndex("供应商名称")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt供应商 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【供应商名称】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "值域错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【供应商名称】列只能是数据库已有供应商名称！供应商【" & Trim(.TextMatrix(lngRow, .ColIndex("供应商名称"))) & "】不存在！"
                        End If
                    End If
                End If
            End If
            '供应商许可证效期
            If GetColumnPostation("供应商许可证效期") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("供应商许可证效期"))) <> "" Then
                    If Not IsDate(Trim(.TextMatrix(lngRow, .ColIndex("供应商许可证效期")))) Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("供应商许可证效期"), lngRow, .ColIndex("供应商许可证效期")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt日期 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【供应商许可证效期】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "格式错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【供应商许可证效期】列日期格式错误！"
                    End If
                End If
            End If
            '规格唯一性
            If GetColumnPostation("类别") > 0 And GetColumnPostation("品种名称") > 0 And GetColumnPostation("药品规格") > 0 And GetColumnPostation("生产商") > 0 Then
                If lngRow > 1 Then
                    For j = lngRow - 1 To 1 Step -1
                        If .TextMatrix(lngRow, .ColIndex("类别")) = .TextMatrix(j, .ColIndex("类别")) And .TextMatrix(lngRow, .ColIndex("品种名称")) = .TextMatrix(j, .ColIndex("品种名称")) And .TextMatrix(lngRow, .ColIndex("生产商")) = .TextMatrix(j, .ColIndex("生产商")) And .TextMatrix(lngRow, .ColIndex("药品规格")) = .TextMatrix(j, .ColIndex("药品规格")) Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("药品规格"), lngRow, .ColIndex("药品规格")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt规格唯一 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【药品规格】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "唯一性错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "该条数据前面已存在类别【" & Trim(.TextMatrix(lngRow, .ColIndex("类别"))) & "】、品种名称【" & Trim(.TextMatrix(lngRow, .ColIndex("品种名称"))) & "】、生产商【" & Trim(.TextMatrix(lngRow, .ColIndex("生产商"))) & "】、药品规格【" & Trim(.TextMatrix(lngRow, .ColIndex("药品规格"))) & "】的数据，请检查！"
                        End If
                    Next
                End If
                If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("药品规格"))), "'") = 0 And InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("品种名称"))), "'") = 0 And InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("生产商"))), "'") = 0 Then
                    strSqls = "Select a.类别, a.Id, a.编码, a.名称, a.规格, a.产地, a.计算单位, b.剂量系数, b.门诊单位, b.门诊包装, b.住院单位, b.住院包装, b.药库单位, b.药库包装," & vbNewLine & _
                             "b.最大效期, b.住院可否分零, b.药库分批, b.药房分批, b.成本价, b.合同单位id, b.门诊可否分零" & vbNewLine & _
                             "From 收费项目目录 A, 药品规格 B Where a.Id = b.药品id And a.类别 In ('5', '6', '7') and 类别=[1] and 名称=[2] and 规格=[3] and 产地" & IIf(Trim(.TextMatrix(lngRow, .ColIndex("生产商"))) = "", " is null", "='" & Trim(.TextMatrix(lngRow, .ColIndex("生产商"))) & "'")
                    Set rs名称 = zlDatabase.OpenSQLRecord(strSqls, "药品规格", Switch(Trim(.TextMatrix(lngRow, .ColIndex("类别"))) = "西成药", 5, Trim(.TextMatrix(lngRow, .ColIndex("类别"))) = "中成药", 6, Trim(.TextMatrix(lngRow, .ColIndex("类别"))) = "中草药", 7), Trim(.TextMatrix(lngRow, .ColIndex("品种名称"))), Trim(.TextMatrix(lngRow, .ColIndex("药品规格"))))
                    If rs名称.RecordCount > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("药品规格"), lngRow, .ColIndex("药品规格")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt规格唯一 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【药品规格】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "唯一性错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "导入项目与数据库中已有数据是否冲突！规格【" & Trim(.TextMatrix(lngRow, .ColIndex("药品规格"))) & "】已存在！"
                    End If
                End If
            End If
        Next
    End With
    
    Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLSAVE, , True)
    cbrControl.Enabled = True
    With vsfError
        If .Rows > 1 Then
            If mbyt导入方式 = 0 Then
                For lngRow = 1 To .Rows - 1
                    If .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(1).Picture Then
                        cbrControl.Enabled = False
                        Exit For
                    Else
                        cbrControl.Enabled = True
                    End If
                Next
            Else
                For lngRow = 1 To .Rows - 1
                    If .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(1).Picture Or .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(2).Picture Then
                        cbrControl.Enabled = False
                        Exit For
                    End If
                Next
            End If
        End If
    End With
    
    If vsfList.Rows > 1 Then
        vsfList.Row = 1: vsfList.Col = 1
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLOUTPUT, , True)
        cbrControl.Enabled = True
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLCHECK, , True)
        cbrControl.Enabled = True
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub setVSF()
    '列宽度对齐方式设置
    Dim cbrControl As CommandBarControl
    Dim lngRow     As Long
    Dim lngCol     As Long
    
    With vsfList
        For lngCol = 1 To .Cols - 1
            Select Case .TextMatrix(0, lngCol)
                Case "售价", "成本价", "售价换算系数", "门诊换算系数", "住院换算系数", "药库换算系数", "效期(月)"
                    .ColAlignment(lngCol) = flexAlignRightCenter
                Case Else
                    .ColAlignment(lngCol) = flexAlignLeftCenter
            End Select
            .ColComboList(lngCol) = ""
        Next
        .FixedRows = 1
        .FixedCols = 1
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter '居中
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True '加粗
        .ExplorerBar = flexExNone   '列不支持排序和拖动
        .ColWidth(-1) = 2000
        .ColWidth(0) = 300
        
        If .Rows > 1 Then
            .Editable = flexEDKbdMouse
            If TabControl.Selected.Caption = "分类" Then
                .ColComboList(.ColIndex("类别")) = "西成药|中成药|中草药"
            Else
                .ColComboList(.ColIndex("类别")) = "西成药|中成药|中草药"
                .ColComboList(.ColIndex("是否变价")) = " |√"
                .ColComboList(.ColIndex("住院可否分零")) = "0-可以分零|1-不可分零|2-一次性使用|3-分零后一天内有效|4-分零后两天内有效|5-分零后三天内有效"
                .ColComboList(.ColIndex("门诊可否分零")) = "0-可以分零|1-不可分零|2-一次性使用|3-分零后一天内有效|4-分零后两天内有效|5-分零后三天内有效"
                If GetColumnPostation("服务对象") > 0 Then .ColComboList(.ColIndex("服务对象")) = "0-不服务于病人|1-门诊|2-住院|3-门诊和住院"
                If GetColumnPostation("药库分批") > 0 Then .ColComboList(.ColIndex("药库分批")) = " |√"
                If GetColumnPostation("药房分批") > 0 Then .ColComboList(.ColIndex("药房分批")) = " |√"
            End If
        End If
    End With
End Sub

Private Sub Form_Resize()
    '控件位置控制
    On Error Resume Next
    
    lblFile.Move 110, 600
    txtFile.Move lblFile.Left + lblFile.Width + 20, lblFile.Top - 40, Me.ScaleWidth - (cmdFile.Width + txtFile.Left) - 50
    cmdFile.Move txtFile.Left + txtFile.Width + 20, txtFile.Top - 30
    
    TabControl.Move lblFile.Left - 40, txtFile.Top + txtFile.Height + 50, Me.ScaleWidth - lblFile.Left - 20, ((Me.ScaleHeight - TabControl.Top) / 5) * 3 - 20
    picSplit.Move lblFile.Left, TabControl.Top + TabControl.Height + 20, Me.ScaleWidth - lblFile.Left * 2
    lblCollect.Width = picSplit.Width
    
    vsfError.Move lblFile.Left - 40, picSplit.Top + picSplit.Height + 50, Me.ScaleWidth - lblFile.Left - 20, Me.ScaleHeight - picSplit.Top - picSplit.Height - 120
    
    vsfError.ColWidth(3) = 10000
    If vsfError.Width > 14500 Then
        vsfError.ColWidth(3) = vsfError.Width - 5000
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjWS = Nothing
    Set mobjWB = Nothing
    Set mobjXLS = Nothing
    mstrType = ""
    mstrMedi = ""
    mstrTypeMsg = ""
    mstrMediMsg = ""
End Sub

Private Sub pic_Resize()
    On Error Resume Next
    vsfList.Move 0, 0, pic.ScaleWidth, pic.ScaleHeight
End Sub

Private Sub lblCollect_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    With picSplit
        If .Top + y < 2000 Then Exit Sub
        If .Top + y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + y + 20
    End With

    With TabControl
        .Height = picSplit.Top - .Top - 20
    End With
    
    With vsfError
        .Top = picSplit.Top + picSplit.Height + 50
        .Height = ScaleHeight - .Top + 50
    End With
    Me.Refresh
End Sub

Private Sub picSplit_Resize()
    On Error Resume Next
    lblCollect.Move 10, , picSplit.ScaleWidth - 10
End Sub

Private Sub TabControl_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Select Case Item.Index
        Case 0
            If mstrTypeMsg <> "" Then
                Call SetColumns("分类")
                If vsfList.Rows > 1 Then Call CheckKind
            End If
        Case 1
            If mstrMediMsg <> "" Then
                Call SetColumns("明细")
                If vsfList.Rows > 1 Then Call Check品种: Call Check规格
            End If
    End Select
End Sub

Private Function InitTabControl()
    '初始化分页控件
    With TabControl
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPageSelected
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .ShowIcons = True
            .DisableLunaColors = False
        End With
        
        .InsertItem 1, "分类", pic.hwnd, 101
        .InsertItem 2, "明细", pic.hwnd, 102
        .Item(0).Selected = True
    End With
End Function

Private Sub SaveType()
'保存分类数据
    Dim lngItemId As Long
    Dim int类别   As Integer
    Dim int上级ID As Long
    Dim str编码   As String
    Dim str名称   As String
    Dim strSql   As String
    Dim strTemp  As String
    Dim rsTemp   As Recordset
    Dim intCount As Integer
    Dim i As Integer
    Dim j As Integer
    Dim arrSql As Variant
    Dim blnTrans As Boolean
    
    On Error GoTo ErrHandle
    arrSql = Array()
    
    Call FS.ShowFlash("正在保存数据,请稍候 ...", Me)
    Me.MousePointer = vbHourglass
    With vsfList
        For i = 1 To .Rows - 1
            '若是错误行，则跳过
            For j = 0 To .Cols - 1
                If .Cell(flexcpForeColor, i, j, i, j) = vbRed Then
                    GoTo ErrHand
                End If
            Next
            
            'ID
            If mintType = 2 Then
                lngItemId = sys.NextId("诊疗分类目录")
            End If
            
            '类别
            If .TextMatrix(i, .ColIndex("类别")) = "西成药" Then
                int类别 = 1
            ElseIf .TextMatrix(i, .ColIndex("类别")) = "中成药" Then
                int类别 = 2
            ElseIf .TextMatrix(i, .ColIndex("类别")) = "中草药" Then
                int类别 = 3
            End If
            
            '上级id
            int上级ID = GetTypeID(.TextMatrix(i, .ColIndex("上级名称")), .TextMatrix(i, .ColIndex("类别")))
            
            '编码
            str编码 = .TextMatrix(i, .ColIndex("编码"))
            
            '名称
            str名称 = .TextMatrix(i, .ColIndex("名称"))
            
            strSql = "zl_诊疗分类目录_insert("
            strSql = strSql & lngItemId & ","
            strSql = strSql & int上级ID & ","
            strSql = strSql & "'" & str编码 & "',"
            strSql = strSql & "'" & str名称 & "',"
            strSql = strSql & "'" & zlStr.GetCodeByVB(str名称) & "',"
            strSql = strSql & int类别 & ","
            strSql = strSql & "0)"
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = strSql
            intCount = intCount + 1
            .TextMatrix(i, 0) = "√"
ErrHand:
        Next
    End With
    
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), "SaveType")
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    Me.MousePointer = vbDefault
    Call FS.StopFlash
    If intCount = vsfList.Rows - 1 And intCount <> 0 Then
        MsgBox "成功保存【分类】页所有数据！", vbInformation, gstrSysName
    ElseIf intCount <> 0 Then
        MsgBox "成功保存【分类】页" & intCount & "条数据！", vbInformation, gstrSysName
    Else
        MsgBox "【分类】页没有合格数据，保存失败！", vbInformation, gstrSysName
    End If
    
    Exit Sub
ErrHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SaveMedi()
'保存明细数据
    Dim rs收入项目 As Recordset
    Dim rs供应商   As Recordset
    Dim intCount  As Integer
    Dim strSql    As String
    Dim lng药名id As Long
    Dim lng药品ID As Long
    Dim int类别   As Integer
    Dim int分类ID As Integer
    Dim str名称   As String
    Dim int服务   As Integer
    Dim i As Integer
    Dim j As Integer
    Dim str编码 As String
    Dim rsTemp As Recordset
    Dim arrSql As Variant
    Dim strTemp As String
    Dim blnTrans As Boolean
    
    On Error GoTo ErrHandle
    arrSql = Array()
    Call FS.ShowFlash("正在保存数据,请稍候 ...", Me)
    Me.MousePointer = vbHourglass
    
    Set rs收入项目 = zlDatabase.OpenSQLRecord("Select ID,编码,名称 From 收入项目 Where 末级 = 1", "收入项目")
    Set rs供应商 = zlDatabase.OpenSQLRecord("Select ID,编码,名称,许可证号,许可证效期 From 供应商", "供应商")
    
    '获取品种编码
    strSql = "Select 类别,id,编码 From 诊疗项目目录 Where 类别 In ('5','6','7')"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "SaveData")
    str编码 = ""
    Do While Not rsTemp.EOF
        '格式：类别,ID[编码]类别,ID[编码]...
        str编码 = str编码 & rsTemp!类别 & "," & rsTemp!ID & "[" & rsTemp!编码 & "]"
        rsTemp.MoveNext
    Loop
    
    With vsfList
        For i = 1 To .Rows - 1
            '若是错误行，则跳过
            For j = 0 To .Cols - 1
                If .Cell(flexcpForeColor, i, j, i, j) = vbRed Then
                    GoTo ErrHand
                End If
            Next
            
            If InStr(1, str编码, "[" & .TextMatrix(i, .ColIndex("品种编码")) & "]") <= 0 Then
                '保存品种
                '类别
                If .TextMatrix(i, .ColIndex("类别")) = "西成药" Then
                    int类别 = 1
                    strTemp = "5"
                    strSql = "zl_成药品种_Insert('5',"
                ElseIf .TextMatrix(i, .ColIndex("类别")) = "中成药" Then
                    int类别 = 2
                    strTemp = "6"
                    strSql = "zl_成药品种_Insert('6',"
                ElseIf .TextMatrix(i, .ColIndex("类别")) = "中草药" Then
                    int类别 = 3
                    strTemp = "7"
                    strSql = "zl_草药品种_Insert('7',"
                End If
                '分类ID
                int分类ID = GetTypeID(.TextMatrix(i, .ColIndex("分类")), .TextMatrix(i, .ColIndex("类别")))
                strSql = strSql & int分类ID & ","
                '药名ID
                lng药名id = sys.NextId("诊疗项目目录")
                str编码 = str编码 & strTemp & "," & lng药名id & "[" & .TextMatrix(i, .ColIndex("品种编码")) & "]"
                strSql = strSql & lng药名id & ","
                '品种编码
                strSql = strSql & "'" & .TextMatrix(i, .ColIndex("品种编码")) & "',"
                '品种名称
                strSql = strSql & "'" & .TextMatrix(i, .ColIndex("品种名称")) & "',"
                '拼音
                str名称 = .TextMatrix(i, .ColIndex("品种名称"))
                strSql = strSql & "'" & zlStr.GetCodeByORCL(str名称) & "',"
                '五笔
                strSql = strSql & "'" & zlStr.GetCodeByORCL(str名称, True) & "',"
                '英文
                strSql = strSql & "'',"
                '剂量单位
                strSql = strSql & "'" & .TextMatrix(i, .ColIndex("剂量单位")) & "',"
                '剂型
                strSql = strSql & IIf(strTemp = "7", "", "'" & .TextMatrix(i, .ColIndex("剂型")) & "',")
                '毒理分类,价值分类,货源情况,用药梯次
                strSql = strSql & "'普通药','普价','充足','首选')"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = strSql
            Else
                '编码格式:类别,ID[编码]...
                '返回编码对应的类别和ID:类别,ID
                strTemp = Mid(Mid(str编码, 1, InStr(1, str编码, "[" & .TextMatrix(i, .ColIndex("品种编码")) & "]") - 1), InStrRev(Mid(str编码, 1, InStr(1, str编码, "[" & .TextMatrix(i, .ColIndex("品种编码")) & "]") - 1), "]") + 1)
                
                '返回ID
                lng药名id = Split(strTemp, ",")(1)
                '返回类别
                strTemp = Split(strTemp, ",")(0)
            End If
            
            '保存规格
            strSql = ""
            '类别
            If strTemp = "5" Then
                strSql = "zl_成药规格_Insert("
            ElseIf strTemp = "6" Then
                strSql = "zl_成药规格_Insert("
            Else
                strSql = "zl_草药规格_Insert("
            End If
            '药名ID
            strSql = strSql & lng药名id & ","
            '药品ID
            lng药品ID = sys.NextId("收费项目目录")
            strSql = strSql & lng药品ID & ","
            '编码
            strSql = strSql & "'" & .TextMatrix(i, .ColIndex("规格编码")) & "',"
            '规格
            strSql = strSql & "'" & .TextMatrix(i, .ColIndex("药品规格")) & "',"
            '生产商
            strSql = strSql & "'" & .TextMatrix(i, .ColIndex("生产商")) & "',"
            '商品名,拼音简码,五笔简码,数字码,标识码,药品来源,批注文号,注册商标
            strSql = strSql & "'','','','','','','','',"
            '售价单位
            strSql = strSql & "'" & .TextMatrix(i, .ColIndex("售价单位")) & "',"
            '剂量系数
            strSql = strSql & .TextMatrix(i, .ColIndex("售价换算系数")) & ","
            '门诊单位
            strSql = strSql & "'" & .TextMatrix(i, .ColIndex("门诊单位")) & "',"
            '门诊系数
            strSql = strSql & .TextMatrix(i, .ColIndex("门诊换算系数")) & ","
            '住院单位
            strSql = strSql & IIf(strTemp = "7", "", "'" & .TextMatrix(i, .ColIndex("住院单位")) & "',")
            '住院系数
            strSql = strSql & IIf(strTemp = "7", "", .TextMatrix(i, .ColIndex("住院换算系数")) & ",")
            '药库单位
            strSql = strSql & "'" & .TextMatrix(i, .ColIndex("药库单位")) & "',"
            '药库系数
            strSql = strSql & .TextMatrix(i, .ColIndex("药库换算系数")) & ","
            '申领单位,申领阀值
            strSql = strSql & "1,null,"
            '是否变价
            If .TextMatrix(i, .ColIndex("是否变价")) = "" Then
                strSql = strSql & "0,"
            ElseIf .TextMatrix(i, .ColIndex("是否变价")) = "√" Then
                strSql = strSql & "1,"
            End If
            '指导批发价：成本价
            strSql = strSql & .TextMatrix(i, .ColIndex("成本价")) & ","
            '扣率
            strSql = strSql & "100,"
            '指导零售价
            strSql = strSql & .TextMatrix(i, .ColIndex("售价")) & ","
            '加成率
            If .TextMatrix(i, .ColIndex("成本价")) = 0 Then
                strSql = strSql & "100,"
            Else
                strSql = strSql & (Val(.TextMatrix(i, .ColIndex("售价"))) / Val(.TextMatrix(i, .ColIndex("成本价"))) - 1) * 100 & ","
                If (Val(.TextMatrix(i, .ColIndex("售价"))) / Val(.TextMatrix(i, .ColIndex("成本价"))) - 1) * 100 > 100 Then
                    strSql = strSql & "100,"
                End If
            End If
            '管理费比例,药价级别,费用类型
            strSql = strSql & "null,'','',"
            '服务对象
            If GetColumnPostation("服务对象") > 0 Then
                Select Case .TextMatrix(i, .ColIndex("服务对象"))
                    Case "0-不服务于病人", ""
                        strSql = strSql & "0,"
                    Case "1-门诊"
                        strSql = strSql & "1,"
                    Case "2-住院"
                        strSql = strSql & "2,"
                    Case "3-门诊和住院"
                        strSql = strSql & "3,"
                End Select
            Else
                strSql = strSql & "3,"
            End If
            'Gmp认证,招标药品,屏蔽费别,
            strSql = strSql & "0,0,0,"
            '住院可否分零
            Select Case .TextMatrix(i, .ColIndex("住院可否分零"))
                Case "0-可以分零", ""
                    strSql = strSql & "0,"
                Case "1-不可分零"
                    strSql = strSql & "1,"
                Case "2-一次性使用"
                    strSql = strSql & "2,"
                Case "3-分零后一天内有效"
                    strSql = strSql & "3,"
                Case "4-分零后两天内有效"
                    strSql = strSql & "4,"
                Case "5-分零后三天内有效"
                    strSql = strSql & "5,"
            End Select
            '药库分批
            If GetColumnPostation("药库分批") > 0 Then
                If .TextMatrix(i, .ColIndex("药库分批")) = "" Then
                    strSql = strSql & "0,"
                ElseIf .TextMatrix(i, .ColIndex("药库分批")) = "√" Then
                    strSql = strSql & "1,"
                End If
            Else
                strSql = strSql & "0,"
            End If
            '药房分批
            If GetColumnPostation("药房分批") > 0 Then
                If .TextMatrix(i, .ColIndex("药房分批")) = "" Then
                    strSql = strSql & "0,"
                ElseIf .TextMatrix(i, .ColIndex("药房分批")) = "√" Then
                    strSql = strSql & "1,"
                End If
            Else
                strSql = strSql & "0,"
            End If
            '效期(月)
            If GetColumnPostation("效期(月)") > 0 Then
                If .TextMatrix(i, .ColIndex("效期(月)")) = "" Then
                    strSql = strSql & "0,"
                Else
                    strSql = strSql & .TextMatrix(i, .ColIndex("效期(月)")) & ","
                End If
            Else
                strSql = strSql & "0,"
            End If
            '差价让利比
            strSql = strSql & "0,"
            '成本价
            strSql = strSql & .TextMatrix(i, .ColIndex("成本价")) & ","
            '售价
            strSql = strSql & .TextMatrix(i, .ColIndex("售价")) & ","
            '收入项目ID
            rs收入项目.Filter = ""
            rs收入项目.Filter = "名称='" & .TextMatrix(i, .ColIndex("收入项目")) & "'"
            strSql = strSql & rs收入项目!ID & ","
            '供应商名称（合同单位id）
            If GetColumnPostation("供应商名称") > 0 Then
                If .TextMatrix(i, .ColIndex("供应商名称")) <> "" Then
                    rs供应商.Filter = ""
                    rs供应商.Filter = "名称='" & .TextMatrix(i, .ColIndex("供应商名称")) & "'"
                    strSql = strSql & rs供应商!ID & ","
                Else
                    strSql = strSql & "null,"
                End If
            Else
                strSql = strSql & "null,"
            End If
            
            If strTemp = "7" Then
                '说明,动态分零,发药类型,备选码,增值税率,基本药物,中药形态,站点,是否常备,病案费目
                strSql = strSql & "Null,0,Null,Null,Null,Null,Null,Null,Null,Null,"
            Else
                '说明,动态分零,发药类型,备选码,增值税率,基本药物,站点,是否常备,存储温度,存储条件,配药类型,是否不予配置,容量,病案费目
                strSql = strSql & "Null,0,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,"
            End If
            '门诊可否分零
            Select Case .TextMatrix(i, .ColIndex("门诊可否分零"))
                Case "0-可以分零", ""
                    strSql = strSql & "0)"
                Case "1-不可分零"
                    strSql = strSql & "1)"
                Case "2-一次性使用"
                    strSql = strSql & "2)"
                Case "3-分零后一天内有效"
                    strSql = strSql & "3)"
                Case "4-分零后两天内有效"
                    strSql = strSql & "4)"
                Case "5-分零后三天内有效"
                    strSql = strSql & "5)"
            End Select
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = strSql
            intCount = intCount + 1
            vsfList.TextMatrix(i, 0) = "√"
ErrHand:
        Next
    End With
    
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), "SavaData")
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    Me.MousePointer = vbDefault
    Call FS.StopFlash
    If intCount = vsfList.Rows - 1 And intCount <> 0 Then
        MsgBox "成功保存【明细】页所有数据！", vbInformation, gstrSysName
    ElseIf intCount <> 0 Then
        MsgBox "成功保存【明细】页" & intCount & "条数据！", vbInformation, gstrSysName
    Else
        MsgBox "【明细】页没有合格数据，保存失败！", vbInformation, gstrSysName
    End If
    
    Exit Sub
ErrHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfError_EnterCell()
    Dim strTemp As String
    Dim lngRow  As Long
    Dim lngCol  As Long
    Dim strCol  As String
    
    With vsfError
        If .Row = 0 Then Exit Sub
        .FocusRect = flexFocusSolid
        If InStr(1, .TextMatrix(.Row, 1), "列") = 0 Then Exit Sub
        If .TextMatrix(.Row, 1) <> "" Then
            strTemp = .TextMatrix(.Row, 1)
            lngRow = Mid(strTemp, 1, InStr(1, strTemp, "行") - 1)
            strCol = Mid(strTemp, InStr(1, strTemp, "【") + 1, InStr(1, strTemp, "】") - InStr(1, strTemp, "【") - 1)
            lngCol = vsfList.ColIndex(strCol)
            If lngRow > vsfList.Rows - 1 Then MsgBox "改行数据已经被删除了！", vbInformation, gstrSysName: Exit Sub
            vsfList.Row = lngRow
            vsfList.Col = lngCol
            vsfList.ShowCell lngRow, lngCol
        End If
    End With
End Sub

Private Sub vsfList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '记录更新的界面数据
    If vsfList.Tag = "1" Then Exit Sub
    If TabControl.Selected.Caption = "分类" Then
        Call GetColumns("分类")
    Else
        Call GetColumns("明细")
    End If
End Sub

Private Sub vsfList_ChangeEdit()
    Dim cbrControl As CommandBarControl
    Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLSAVE, , True)
    cbrControl.Enabled = False
'    Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLOUTPUT, , True)
'    cbrControl.Enabled = False
End Sub

Private Sub vsfList_DblClick()
    With vsfList
        If .Rows > 1 And .Row > 0 Then
            .EditCell
            .EditSelStart = 0
            .EditSelLength = LenB(StrConv(.EditText, vbFromUnicode))
        End If
    End With
End Sub

Private Sub vsfList_EnterCell()
    Dim strTemp As String
    Dim intRow  As Integer
    Dim i As Integer
    With vsfList
        If .Row < 1 Then Exit Sub
        strTemp = .Row & "行【" & .TextMatrix(0, .Col) & "】列"
        .FocusRect = flexFocusSolid
    End With
    
    With vsfError
        If .Rows < 2 Then Exit Sub
        i = 0
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 1) = strTemp Then
                .Row = intRow
                .TopRow = intRow
                i = 1
                Exit For
            End If
        Next
        If i = 0 Then .Row = 0
    End With
End Sub

Private Sub vsfList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim cbrControl As CommandBarControl
    
    With vsfList
        If KeyCode = vbKeyDelete Then
            If .Row < 1 Then Exit Sub
            If MsgBox("将删除第" & .Row & "行数据，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                With vsfList
                    .RemoveItem .Row
                    '记录更新的界面数据
                    If TabControl.Selected.Caption = "分类" Then
                        Call GetColumns("分类")
                    Else
                        Call GetColumns("明细")
                    End If
                    If .Rows <= 1 Then
                        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLOUTPUT, , True)
                        cbrControl.Enabled = False
                        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLCHECK, , True)
                        cbrControl.Enabled = False
                        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLSAVE, , True)
                        cbrControl.Enabled = False
                    End If
                End With
            End If
        End If
        
        If KeyCode = vbKeyReturn And .Rows > 1 Then
            If .Col = .Cols - 1 Then
                If .Row = .Rows - 1 Then .Rows = .Rows + 1: .TextMatrix(.Rows - 1, 0) = .Rows - 1
                .Row = .Row + 1
                .Col = 1
            Else
                .Col = .Col + 1
            End If
        End If
    End With
End Sub

Private Sub vsfList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsfList
        If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Then
            Exit Sub
        End If
        Select Case .Col
            Case .ColIndex("售价换算系数"), .ColIndex("门诊换算系数"), .ColIndex("住院换算系数"), .ColIndex("药库换算系数"), .ColIndex("效期(月)")
                If InStr(1, "1234567890", Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            Case .ColIndex("编码"), .ColIndex("品种编码"), .ColIndex("规格编码")
                If InStr(1, "1234567890abcdefghijklmnopqrstuvwxyz", LCase(Chr(KeyAscii))) = 0 Then
                    KeyAscii = 0
                End If
            Case .ColIndex("成本价"), .ColIndex("售价")
                If InStr(1, "1234567890.", Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
                If Chr(KeyAscii) = "." And InStr(1, .EditText, ".") > 0 Then
                    KeyAscii = 0
                End If
            Case .ColIndex("供应商许可证效期")
                If InStr(1, "1234567890.-/", Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            Case .ColIndex("分类"), .ColIndex("品种名称"), .ColIndex("药品规格"), .ColIndex("生产商"), .ColIndex("剂型"), .ColIndex("剂量单位"), .ColIndex("售价单位"), .ColIndex("门诊单位"), .ColIndex("住院单位"), .ColIndex("药库单位"), .ColIndex("收入项目"), .ColIndex("供应商名称"), .ColIndex("供应商许可证号")
                If InStr(" ~%^&|`'""", Chr(KeyAscii)) > 0 Then
                    KeyAscii = 0
                End If
        End Select
    End With
End Sub

Private Sub vsfList_RowColChange()
    If vsfList.Rows > 1 Then Call SetNote
End Sub

Private Sub vsfList_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim cbrControl As CommandBarControl
    Dim strSql     As String
    Dim rsTemp     As Recordset
    
    If mbyt导入方式 = 1 And vsfError.Rows > 1 Then
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLSAVE, , True)
        cbrControl.Enabled = False
    End If
    
    With vsfList
        Select Case .Col
            Case .ColIndex("成本价"), .ColIndex("售价")
                strSql = "select 精度 from 药品卫材精度 where 类别=1 and 内容=" & IIf(.Col = .ColIndex("成本价"), 1, 2) & " and 单位=1"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "")
                .EditText = zlStr.FormatEx(.EditText, Val(rsTemp!精度), , True)
            Case .ColIndex("售价换算系数"), .ColIndex("门诊换算系数"), .ColIndex("住院换算系数"), .ColIndex("药库换算系数"), .ColIndex("效期(月)")
                .EditText = Val(.EditText)
            Case .ColIndex("供应商许可证效期")
                If IsNumeric(.EditText) Then
                    .EditText = TranNumToDate(.EditText)
                Else
                    .EditText = FormatDate(.EditText)
                End If
            Case .ColIndex("编码"), .ColIndex("品种编码"), .ColIndex("规格编码")
                .EditText = UCase(Trim(.EditText))
        End Select
        .EditText = Trim(.EditText)
    End With
End Sub

Private Sub SetCols()
    Dim strMediColumn As String
    
    Select Case mintType
    Case 1  '收费目录管理
    Case 2   '药品目录管理
        strMediColumn = zlDatabase.GetPara("列的显示隐藏", glngSys, mlngModule, MSTRMEDICAL)
    Case 3   '卫材项目管理
    End Select
    
    If Not frmImportFileCols.ShowMe(Me, strMediColumn) Then Exit Sub
    If MsgBox("列设置保存后导入项目会自动关闭，需下次打开才会生效，" & vbCrLf & "是否继续保存？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
 
    Call zlDatabase.SetPara("列的显示隐藏", strMediColumn, glngSys, mlngModule)
    Call GetColumnHead
    
    Unload Me
End Sub

Public Sub ShowMe(ByVal intType As Integer, ByVal frmParent As Form)
    Call InitExcel
    
    If mobjXLS Is Nothing Then
        err.Clear
        Exit Sub
    End If
    
    mobjXLS.DisplayAlerts = False
    
    mlngModule = glngModul
    mintType = intType
    
    Me.Show 1, frmParent
End Sub

Public Sub GetColumnHead()
'记录下显示的分类和明细的列头信息
    Dim arrTypeColumn As Variant
    Dim arrMediColumn As Variant
    Dim strMedical    As String
    Dim intCol As Integer
    Dim intNum As Integer
    
    mstrType = ""
    mstrMedi = ""
    strMedical = zlDatabase.GetPara("列的显示隐藏", glngSys, mlngModule)
    '分类
    arrTypeColumn = Split(Split(strMedical, "||")(0) & "|", "|")
    Do While arrTypeColumn(intCol) <> ""
        If Split(arrTypeColumn(intCol), ",")(2) = 0 Then
            mstrType = mstrType & Split(arrTypeColumn(intCol), ",")(0) & "|"
        End If
        intCol = intCol + 1
    Loop
    '明细
    arrMediColumn = Split(Split(strMedical, "||")(1), "|")
    Do While arrMediColumn(intNum) <> ""
        If Split(arrMediColumn(intNum), ",")(2) = 0 Then
            mstrMedi = mstrMedi & Split(arrMediColumn(intNum), ",")(0) & "|"
        End If
        intNum = intNum + 1
    Loop
End Sub

Private Sub GetColumns(ByVal strType As String)
'记录下表格中分类和明细的所有信息
    Dim lngRow As Long
    Dim lngCol As Long
    
    With vsfList
        If strType = "分类" Then
            mstrTypeMsg = ""
            For lngRow = 0 To .Rows - 1
                For lngCol = 1 To .Cols - 1
                    mstrTypeMsg = mstrTypeMsg & Trim(.TextMatrix(lngRow, lngCol)) & ";"
                Next
                mstrTypeMsg = mstrTypeMsg & "|"
            Next
        ElseIf strType = "明细" Then
            mstrMediMsg = ""
            For lngRow = 0 To .Rows - 1
                For lngCol = 1 To .Cols - 1
                    mstrMediMsg = mstrMediMsg & Trim(.TextMatrix(lngRow, lngCol)) & ";"
                Next
                mstrMediMsg = mstrMediMsg & "|"
            Next
        End If
    End With
End Sub

Private Function GetTypeID(ByVal strVal As String, ByVal strType As String, Optional ByVal strKind As String) As Long
'获取分类ID
    Dim strSql As String
    Dim intType As Integer
    Dim strSecType As String
    Dim rsTemp As Recordset

    On Error GoTo ErrHand
    
    If strType = "西成药" Then
        intType = 1
    ElseIf strType = "中成药" Then
        intType = 2
    ElseIf strType = "中草药" Then
        intType = 3
    Else
        intType = 0
    End If
    
    '检查分类是否只有一级
    If InStr(1, strVal, "\") > 1 Then
        strType = Mid(strVal, InStrRev(strVal, "\") + 1)
        strSecType = Mid(strVal, 1, InStrRev(strVal, "\") - 1)
    Else
        strType = strVal
        strSecType = ""
    End If

    If strSecType = "" And InStr(1, strVal, "\") = 0 Then
        '分类只有一级的情况
        strSql = "Select ID,编码 " & _
                 "From 诊疗分类目录 " & _
                 "Where 名称 = [1] And 类型=[2] order by ID"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "GetTypeID", strType, intType)
    Else
        strSql = "Select ID,编码 From 诊疗分类目录" & vbNewLine & _
                "Where 名称 = [1] And 上级id in (Select ID From 诊疗分类目录 Where 名称 = [2] And 类型=[3] ) order by ID"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "GetTypeID", strType, strSecType, intType)
    End If

    If rsTemp.RecordCount > 0 Then
        GetTypeID = rsTemp!ID
    ElseIf rsTemp.RecordCount = 0 Then
        GetTypeID = 0
    End If
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetNote()
'设置列说明
    Dim arrComent As Variant
    
    With vsfList
        If TabControl.Selected.Caption = "分类" Then
            arrComent = Split(Split(MSTRCOMMENT, "||")(0), "|")
            Select Case .Col
                Case .ColIndex("类别")
                    lblCollect.Caption = Split(arrComent(0), ";")(1)
                Case .ColIndex("上级名称")
                    lblCollect.Caption = Split(arrComent(1), ";")(1)
                Case .ColIndex("编码")
                    lblCollect.Caption = Split(arrComent(2), ";")(1)
                Case .ColIndex("名称")
                    lblCollect.Caption = Split(arrComent(3), ";")(1)
                Case Else
                    lblCollect.Caption = ""
            End Select
        Else
            arrComent = Split(Split(MSTRCOMMENT, "||")(1), "|")
            Select Case .Col
                Case .ColIndex("类别")
                    lblCollect.Caption = Split(arrComent(0), ";")(1)
                Case .ColIndex("分类")
                    lblCollect.Caption = Split(arrComent(1), ";")(1)
                Case .ColIndex("品种编码")
                    lblCollect.Caption = Split(arrComent(2), ";")(1)
                Case .ColIndex("品种名称")
                    lblCollect.Caption = Split(arrComent(3), ";")(1)
                Case .ColIndex("规格编码")
                    lblCollect.Caption = Split(arrComent(4), ";")(1)
                Case .ColIndex("药品规格")
                    lblCollect.Caption = Split(arrComent(5), ";")(1)
                Case .ColIndex("生产商")
                    lblCollect.Caption = Split(arrComent(6), ";")(1)
                Case .ColIndex("剂型")
                    lblCollect.Caption = Split(arrComent(7), ";")(1)
                Case .ColIndex("剂量单位"), .ColIndex("售价单位"), .ColIndex("门诊单位"), .ColIndex("住院单位"), .ColIndex("药库单位")
                    lblCollect.Caption = Split(arrComent(8), ";")(1)
                Case .ColIndex("售价换算系数"), .ColIndex("门诊换算系数"), .ColIndex("住院换算系数"), .ColIndex("药库换算系数")
                    lblCollect.Caption = Split(arrComent(10), ";")(1)
                Case .ColIndex("是否变价")
                    lblCollect.Caption = Split(arrComent(17), ";")(1)
                Case .ColIndex("成本价"), .ColIndex("售价")
                    lblCollect.Caption = Split(arrComent(18), ";")(1)
                Case .ColIndex("收入项目")
                    lblCollect.Caption = Split(arrComent(20), ";")(1)
                Case .ColIndex("住院可否分零"), .ColIndex("门诊可否分零")
                    lblCollect.Caption = Split(arrComent(21), ";")(1)
                Case .ColIndex("服务对象")
                    lblCollect.Caption = Split(arrComent(23), ";")(1)
                Case .ColIndex("药库分批"), .ColIndex("药房分批")
                    lblCollect.Caption = Split(arrComent(24), ";")(1) & "," & Split(arrComent(25), ";")(1)
                Case .ColIndex("效期(月)")
                    lblCollect.Caption = Split(arrComent(26), ";")(1)
                Case .ColIndex("供应商名称")
                    lblCollect.Caption = Split(arrComent(27), ";")(1)
                Case .ColIndex("供应商许可证号")
                    lblCollect.Caption = Split(arrComent(28), ";")(1)
                Case .ColIndex("供应商许可证效期")
                    lblCollect.Caption = Split(arrComent(29), ";")(1)
                Case Else
                    lblCollect.Caption = ""
            End Select
        End If
        lblCollect.ForeColor = &HFF0000
    End With
End Sub
