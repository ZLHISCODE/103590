VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.UserControl ucFlexGrid 
   ClientHeight    =   5370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4920
   ScaleHeight     =   5370
   ScaleWidth      =   4920
   Tag             =   "1"
   ToolboxBitmap   =   "ucFlexGrid.ctx":0000
   Begin VB.Timer TimerRefreshData 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4320
      Top             =   2760
   End
   Begin VB.CheckBox chkCheckAll 
      Height          =   200
      Left            =   4440
      TabIndex        =   3
      Top             =   1200
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.PictureBox picShowHint 
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1635
      ScaleWidth      =   4155
      TabIndex        =   1
      Top             =   120
      Width           =   4215
      Begin VB.Image imgWarning 
         Height          =   480
         Left            =   600
         Picture         =   "ucFlexGrid.ctx":0312
         Top             =   600
         Width           =   480
      End
      Begin VB.Label labInf 
         BackStyle       =   0  'Transparent
         Caption         =   "数据有误，请检查。"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   1200
         TabIndex        =   2
         Top             =   600
         Width           =   2295
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgData 
      Height          =   3165
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   4200
      _cx             =   7408
      _cy             =   5583
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16744576
      ForeColorSel    =   -2147483640
      BackColorBkg    =   8421504
      BackColorAlternate=   16777215
      GridColor       =   12632256
      GridColorFixed  =   8421504
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   20
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   240
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   4
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
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   1
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
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
      AllowUserFreezing=   3
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.CommandButton cmdCellBtn 
         Caption         =   "…"
         Height          =   375
         Left            =   3840
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.Image imgCheck 
      Height          =   255
      Index           =   2
      Left            =   4440
      Picture         =   "ucFlexGrid.ctx":0FDC
      Stretch         =   -1  'True
      Tag             =   "2"
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   255
      Index           =   1
      Left            =   4440
      Picture         =   "ucFlexGrid.ctx":134E
      Stretch         =   -1  'True
      Tag             =   "1"
      Top             =   840
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   255
      Index           =   0
      Left            =   4440
      Picture         =   "ucFlexGrid.ctx":16C0
      Stretch         =   -1  'True
      Tag             =   "0"
      Top             =   480
      Width           =   240
   End
   Begin VB.Menu menuPop1 
      Caption         =   "右键菜单1"
      Begin VB.Menu mnuCopy 
         Caption         =   "复制(&C)"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "剪切(&T)"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "粘贴(&P)"
      End
      Begin VB.Menu mnuSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDel 
         Caption         =   "删除(&D)"
      End
   End
End
Attribute VB_Name = "ucFlexGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const M_STR_AdjustColName As String = "≡"          '扩展列，可显示行号
Private Const M_STR_NameSplitChar As String = ">"           '字段名分隔符
Private Const M_STR_PropertySplitChar As String = "@"       '属性分隔符
Private Const M_STR_PlaceCol = "[placecol]"                 '占位列

Private Const M_STR_ColProp_Hide As String = "hide"         '隐藏列
Private Const M_STR_ColProp_Read As String = "read"         '只读列
Private Const M_STR_ColProp_Btn As String = "btn"           '按钮列
Private Const M_STR_ColProp_Merge As String = "merge"       '和并列
Private Const M_STR_ColProp_CellCheck As String = "check"   '单元格选择列
Private Const M_STR_ColProp_RowCheck As String = "rowcheck" '行选择列
Private Const M_STR_ColProp_Align As String = "align"       '补齐方式
Private Const M_STR_ColProp_Key As String = "key"           '关键列
Private Const M_STR_ColProp_Cbx As String = "cbx"           '下拉框列
Private Const M_STR_ColProp_TxtLeft As String = "txtleft"   '左对齐
Private Const M_STR_ColProp_TxtRight As String = "txtright" '右对齐
Private Const M_STR_ColProp_TxtCenter As String = "txtcenter" '居中对齐
Private Const M_STR_ColProp_ColLeft As String = "colleft"
Private Const M_STR_ColProp_ColRight As String = "colright"
Private Const M_STR_ColProp_ColCenter As String = "colcenter"
Private Const M_STR_ColProp_ChkLeft As String = "chkleft"
Private Const M_STR_ColProp_ChkRight As String = "chkright"
Private Const M_STR_ColProp_ChkCenter As String = "chkcenter"
Private Const M_STR_ColProp_WidthTag As String = "w"        '列宽度，如W2100
Private Const M_STR_ColProp_Tdate As String = "tdate"       '日期类型
Private Const M_STR_ColProp_Tnum As String = "tnum"         '数字类型
Private Const M_STR_ColProp_Tstr As String = "tstr"         '字符串类型
Private Const M_STR_ColProp_TFullDateTime As String = "fulldatetime"
Private Const M_STR_ColProp_TOnlyDate As String = "onlydate"
Private Const M_STR_ColProp_TOnlyTime As String = "onlytime"
Private Const M_STR_ColProp_TShortDateTime As String = "shortdatetime"
Private Const M_STR_ColProp_HeadImg As String = "headimg"   '列头图标
Private Const M_STR_ColProp_DataImg As String = "dataimg"   '数据图标
Private Const M_STR_ColProp_UnResize As String = "unresize" '不允许调整列宽度
Private Const M_STR_ColProp_UnCfg As String = "uncfg"    '不允许配置列

Private Const M_STR_ConvertProp_Img = "<img0..n>"
Private Const M_STR_ConvertProp_Check = "<check>"
Private Const M_STR_ConvertProp_NoCheck = "<nocheck>"
Private Const M_STR_ConvertProp_Source = "<source>"
Private Const M_STR_ConvertProp_Els = "els"




'列定义格式为：列名称,是否隐藏(默认不隐藏),可否编辑(默认可编辑),是否Button按钮(默认不是),宽度
'如：|材料类别,read,merge|病理号,merge,read,uncfg,headimg0,dataimg1|ID,key,hide,uncfg|材块号>序号,read,uncfg,rowcheck|标本名称,read|取材位置,read,w1600|材料明细,read,uncfg|在档数量,read|日起,read,onlydate|存放状态,read|"
'
'
'如果列名称为“≡”则表示该列为扩展列，主要用于控制行的高度,扩展列将显示行编号
'
'列属性如下：
'显示名称>字段名称
'hide：表示隐藏
'btn：表示该列有button按钮
'read：表示该列为只读
'merge：表示该列为合并列（行合并）
'check：表示Cell中包含checkbox控件
'RowCheck：表示行选的CheckBox
'w1600：表示宽度为1600
'key:表示为关键字段
'fulldatetime：yyyy-mm-dd hh:mm:ss 完全日期时间格式
'onlydate：yyyy-mm-dd 仅日期格式
'onlytime：hh:mm:ss   仅时间格式
'shortdatetime：yyyy-mm-dd hh:mm  短日期时间格式
'cbx<0-否,1-是,2-未设置>：表示该列为可选列
'Align<8,0>：补足位数对齐
'colleft,colcenter,colright：表示列的对齐方式
'txtleft,txtcenter,txtright：表示文本的对齐方式
'chkleft,chkcenter,chkright：表示check的对齐方式
'tdate：表示时间类型
'tnum：表示数字类型
'tstr：表示字符串类型
'uncfg：表示不允许配置隐藏
'headimg0..n：表示列标题显示时，添加0..n中的图像
'dataimg0..n：表示数据显示时，添加0..n中的图像
'unresize：表示该列不允许改变列宽度


'列转换属性如下：
'如特检类型:0-免疫组化,1-特殊染色,2-分子病理,els-其他|当前状态:0-未处理,1-已接受,2-已完成|清单状态:0-<nocheck>未打印,1-<check>已打印
'<nocheck>表示该数据显示时，单元格会添加未选中的勾选框，
'<check>表示该数据显示时，单元格会添加已选中的勾选框，
'<img0..n>表示图像0..n中的一幅
'els:表示当条件都不满足时，取该值

'============================================================================================================================

'表示未设置的值
Private Const M_LNG_UNCFG As Long = -100


'check状态
Public Enum CheckState
    csNone = -1
    csCheck = 0
    csNoCheck = 1
    csDisCheck = 2
End Enum

'数据行状态
Public Enum TDataRowState
    Normal = 0  '正常
    Add = 1 '新增
    Update = 2  '更新
    Del = 3 '删除
End Enum


'对象显示位置的类型
Public Enum ObjPostionType
    optLeft = 0  '靠左
    optRight = 1 '靠右
End Enum


'列属性设置
Public Enum TColPro
    cpColName = 0       '列名
    cpFieldName = 1     '字段名
    cpHeadImgIndex = 2  '列图标索引
    cpDataImgIndex = 3  '数据图标索引
    cpIsHide = 4        '是否隐藏列
    cpIsCheck = 5       '是否check列
    cpIsKey = 6         '是否关键数据列
    cpIsCombox = 7      '是否combobox列
    cpIsRowCheck = 8    '是否rowcheck列
    cpWidth = 9         '列宽度
    cpIsUnResize = 10     '是否允许调整列宽
    cpIsMerage = 11     '是否合并列
    cpIsBtn = 12        '是否button列
    cpIsRead = 13       '是否只读列
    cpTxtAlign = 14     '文本对齐方式
    cpColAlign = 15     '列对齐方式
    cpChkAlign = 16     'check对齐方式
    cpIsUnCfg = 17        '是否允许配置该列
    cpDataType = 18     '列数据类型
    cpProperty = 19     '列属性字符串
    cpIsDateCol = 20    '是否日期列
    cpAlignLen = 21     '补齐长度
    cpAlignChar = 22    '补齐字符
    cpIsUpdateStyle = 23 '是否更新数据列样式
    cpComboxCfg = 24    'combox配置信息
    cpTag = 25          '列保留标记
End Enum


Private mrsData As ADODB.Recordset
Private mDataGrid As VSFlexGrid
Private mobjImageList As ImageList

Private mobjHeadFont As StdFont
Private moleHeadColor As OLE_COLOR

Private mstrKeyName As String                 '关键列名称
Private mstrKeyField As String                '关键字段名
Private mblnIsShowNumber As Boolean           '是否显示行编号
Private mlngDisableColor As Long              '不可编辑单元格颜色
Private mlngKeepRows As Long

Private mobjColDictionary  As Scripting.Dictionary
Private mobjTmpDictionary As Scripting.Dictionary
Private mlngCurColProIndex As Long            '保存当前字典所属列的索引

Private mstrColNames As String    '列名串
Private mStrDefaultColNames As String  '默认列名串
Private mblnIsKeepRows As Boolean   '是否保存列表行数
Private mlngErrCellColor As Long    '数据错误单元格颜色
Private mblnIsEnterNextCell As Boolean  '回车是否跳转到下一单元格
Private mblnIsBtnNextCell As Boolean    '列表按钮执行后，是否跳转到下一单元格
Private mstrDataConvertFormat As String  '数据转换格式串
Private mstrAdoFilter As String 'ado数据过滤条件
Private mblnIsCopyAdoMode As Boolean    '是否使用ado数据复制模式
Private mblnIsDelKeyRemoveData As Boolean
Private mblnReadOnly As Boolean
Private mblnIsAllowExtCol As Boolean          '是否允许扩展列
Private mblnIsShowPopupMenu As Boolean
Private mblnIsAutoRowHeight As Boolean

Private mblnIsEjectConfig As Boolean    '是否允许右键弹出列表设置窗口
Private mlngSortCol As Long
Private mlngSortWay As Long


'Private mCols() As colInf                     '数据列信息
Private mlngOldBackColor As Long
Private mlngOldGridColor As Long
Private mlngOldDisCellColor As Long


Private mlngOldDataRowHeight As Long
Private mobjRegExp As New RegExp



'对API变量的定义
Const MOUSEEVENTF_LEFTDOWN = &H2
Const MOUSEEVENTF_LEFTUP = &H4
  
  
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, _
          ByVal cButtons As Long, ByVal dwExtraInfo As Long)


Public Event OnOrderChange(ByVal lngCol As Long, ByVal lngOrder As Integer, ByRef blnCustom As Boolean)

Public Event OnBeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)

Public Event OnStartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Public Event OnBeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Public Event OnAfterEdit(ByVal Row As Long, ByVal Col As Long)

Public Event OnKeyDown(KeyCode As Integer, Shift As Integer)
Public Event OnKeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Public Event OnKeyUp(KeyCode As Integer, Shift As Integer)
Public Event OnKeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Public Event OnKeyPress(KeyAscii As Integer)
Public Event OnKeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)

Public Event OnMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OnMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Event OnClick()
Public Event OnDblClick()
Public Event OnRowColChange()
Public Event OnSelChange()
Public Event OnLeaveCell()
Public Event OnEnterCell()
Public Event OnChangeEdit()
Public Event OnColFormartChange()
Public Event OnColsNameReSet()

Public Event OnCheckChanging(ByVal Row As Long, ByVal Col As Long, AllowChange As Boolean)
Public Event OnCheckChanged(ByVal Row As Long, ByVal Col As Long)

Public Event OnCheckAllChanging(AllowChange As Boolean)
Public Event OnCheckAllChanged()

Public Event OnCellButtonClick(ByVal Row As Long, ByVal Col As Long)
Public Event OnCellChanged(ByVal Row As Long, ByVal Col As Long)
Public Event OnDeleteRow(ByVal Row As Long, ByVal Col As Long, AllowDel As Boolean)
Public Event OnNewRow(ByVal Row As Long)
'Public Event OnBeforeReadAdoData(ByVal lngRow As Long, ByVal lngCol As Long, ByVal strCol As String, ByVal strFieldName As String, rsData As ADODB.Recordset, ByRef strText As String, ByRef strTag As String)
'Public Event OnAfterReadAdoData(ByVal lngRow As Long, ByVal lngCol As Long, ByVal strCol As String, ByVal strFieldName As String, rsData As ADODB.Recordset, ByVal strText As String, ByVal strTag As String)

Public Event OnBindFilter(ByRef strBindFilter As String, ByRef strCloneFilter As String)
Public Event OnFilterRowData(rsData As ADODB.Recordset, rsClone As ADODB.Recordset, ByRef blnFilterOut As Boolean)
Public Event OnRefreshRowData(rsBind As ADODB.Recordset, ByVal lngRow As Long)


'列头字体样式
Property Get HeadFont() As StdFont
     Set HeadFont = mobjHeadFont
End Property

Property Set HeadFont(value As StdFont)
    
    With mobjHeadFont
        .Bold = value.Bold
        .Charset = value.Charset
        .Italic = value.Italic
        .Name = value.Name
        .Size = value.Size
        .Strikethrough = value.Strikethrough
        .Underline = value.Underline
        .Weight = value.Weight
    End With

    Call ConfigHeadFont
End Property

'标题颜色
Property Get HeadColor() As OLE_COLOR
    HeadColor = moleHeadColor
End Property

Property Let HeadColor(value As OLE_COLOR)
    moleHeadColor = value
    
    Call ConfigHeadFont
End Property


'数据颜色
Property Get DataColor() As OLE_COLOR
    DataColor = mDataGrid.ForeColor
End Property

Property Let DataColor(value As OLE_COLOR)
    mDataGrid.ForeColor = value
    mDataGrid.ForeColorFixed = value
    
    Call ConfigDataFont
End Property

'网格线颜色
Property Get GridLineColor() As OLE_COLOR
    GridLineColor = mDataGrid.GridColor
End Property

Property Let GridLineColor(value As OLE_COLOR)
    mDataGrid.GridColor = value
End Property


'是否扩展最后一列到适应宽度
Property Get ExtendLastCol() As Boolean
    ExtendLastCol = vfgData.ExtendLastCol
End Property

Property Let ExtendLastCol(value As Boolean)
    vfgData.ExtendLastCol = value
End Property


Private Sub ConfigHeadFont()
    Dim lngFontHeight As Long

    mDataGrid.Cell(flexcpFont, 0, 0, 0, mDataGrid.Cols - 1) = mobjHeadFont
    mDataGrid.Cell(flexcpForeColor, 0, 0, 0, mDataGrid.Cols - 1) = moleHeadColor
    
    Set UserControl.Font = mobjHeadFont
    lngFontHeight = UserControl.TextHeight("字")
    
    mDataGrid.RowHeight(0) = lngFontHeight + 120
    
    Call vfgData_AfterUserResize(0, 0)
'    Call mDataGrid.AutoSize(0, mDataGrid.Cols - 1)
End Sub



'数据列字体样式
Property Get DataFont() As StdFont
    Set DataFont = mDataGrid.Font
End Property

Property Set DataFont(value As StdFont)
    With mDataGrid.Font
        .Bold = value.Bold
        .Charset = value.Charset
        .Italic = value.Italic
        .Name = value.Name
        .Size = value.Size
        .Strikethrough = value.Strikethrough
        .Underline = value.Underline
        .Weight = value.Weight
    End With
    
    Call ConfigDataFont
    Call ConfigHeadFont
End Property


Private Sub ConfigDataFont()
    Dim lngFontHeight As Long
    Dim i As Long
    
    If mDataGrid.Rows <= 1 Then Exit Sub
    
    mDataGrid.Cell(flexcpFont, 1, 0, mDataGrid.Rows - 1, mDataGrid.Cols - 1) = mDataGrid.Font
    mDataGrid.Cell(flexcpForeColor, 1, 0, mDataGrid.Rows - 1, mDataGrid.Cols - 1) = mDataGrid.ForeColor
    
    Set UserControl.Font = mDataGrid.Font
    lngFontHeight = UserControl.TextHeight("字")
    
    For i = 1 To mDataGrid.Rows - 1
        mDataGrid.RowHeight(i) = lngFontHeight + 120
    Next i
    
    Call vfgData_AfterUserResize(0, 0)
'    Call mDataGrid.AutoSize(0, 0)
End Sub



Property Get ImageList() As ImageList
    Set ImageList = mobjImageList
End Property

Property Set ImageList(value As ImageList)
    Set mobjImageList = value
End Property



'列名属性
Property Get ColNames() As String
    ColNames = mstrColNames
End Property


Property Let ColNames(value As String)
    If UCase(value) = UCase(mstrColNames) Then Exit Property
    
    mstrColNames = value
    
    Call RefreshColConfig
End Property

'默认列属性
Property Get DefaultColNames() As String
    DefaultColNames = mStrDefaultColNames
End Property


Property Let DefaultColNames(value As String)
    mStrDefaultColNames = value
End Property

''adjust列名称(只读属性)
'Property Get AdjustColName() As String
'    AdjustColName = M_STR_AdjustColName
'End Property


'关键字（对应列表中的显示名称）
Property Get KeyName() As String
    KeyName = mstrKeyName
End Property


Property Let KeyName(value As String)
    mstrKeyName = value
End Property

'最小行高度
Property Get RowHeightMin() As Long
    RowHeightMin = mDataGrid.RowHeightMin
End Property

Property Let RowHeightMin(value As Long)
    mDataGrid.RowHeightMin = value
End Property



'是否允许扩展列
Property Get AllowExtCol() As Boolean
    AllowExtCol = mblnIsAllowExtCol
End Property


Property Let AllowExtCol(value As Boolean)
    If value = mblnIsAllowExtCol Then Exit Property
    
    mblnIsAllowExtCol = value
    
    Call RefreshColConfig
End Property


'是否自动行高
Property Get IsAutoRowHeight() As Boolean
    IsAutoRowHeight = mblnIsAutoRowHeight
End Property

Property Let IsAutoRowHeight(value As Boolean)
    mblnIsAutoRowHeight = value
End Property


Private Sub RefreshColConfig()
    Call InitVsFlexGridList(vfgData, mstrColNames)
    
    Call RefreshCbxPostion
    Call UpdateRowNumber
    
    Call RefreshReadColColor
    Call RefreshAlign
End Sub


'只读属性
Property Get ReadOnly() As Boolean
    ReadOnly = mblnReadOnly
End Property


Property Let ReadOnly(value As Boolean)
    If mblnReadOnly = value Then Exit Property
    
    mblnReadOnly = value
    
    vfgData.Editable = IIf(value, flexEDNone, flexEDKbdMouse)
    
'    If mblnReadOnly Then
'        mlngOldBackColor = vfgData.BackColor
'        mlngOldGridColor = vfgData.GridColor
'        mlngOldDisCellColor = DisCellColor
'
'        vfgData.BackColor = &HD0E0E0
'        vfgData.GridColor = &HC0C0C0
'
'        DisCellColor = &HD0E0E0
'
'        vfgData.Editable = IIf(value, flexEDNone, flexEDKbdMouse)
'    Else
'        vfgData.Editable = IIf(value, flexEDNone, flexEDKbdMouse)
'
'        vfgData.BackColor = mlngOldBackColor
'        vfgData.GridColor = mlngOldGridColor
'
'        DisCellColor = mlngOldDisCellColor
'    End If
   
End Property



'数据行状态
Property Get RowState(ByVal lngRow As Long) As TDataRowState
    RowState = mDataGrid.RowData(lngRow)
End Property

Property Let RowState(ByVal lngRow As Long, value As TDataRowState)
    mDataGrid.RowData(lngRow) = value
End Property



'当前行状态
Property Get CurRowState() As TDataRowState
    CurRowState = RowState(mDataGrid.RowSel)
End Property


Property Let CurRowState(ByVal value As TDataRowState)
    RowState(mDataGrid.RowSel) = value
End Property



'列头的check状态
Property Get HeadCheckValue() As Boolean
    HeadCheckValue = IIf(chkCheckAll.value <> 0, True, False)
End Property


Property Let HeadCheckValue(value As Boolean)
    chkCheckAll.value = IIf(value, 1, 0)
End Property

'是否显示右键弹出菜单
Property Get IsShowPopupMenu() As Boolean
    IsShowPopupMenu = mblnIsShowPopupMenu
End Property


Property Let IsShowPopupMenu(value As Boolean)
    mblnIsShowPopupMenu = value
End Property

Property Get Enabled() As Boolean
    Enabled = vfgData.Enabled
End Property

Property Let Enabled(ByVal vNewValue As Boolean)
    vfgData.Enabled = vNewValue
    chkCheckAll.Enabled = vNewValue
End Property

'返回是否选择行
Property Get IsSelectionRow() As Boolean
    IsSelectionRow = False

    If mDataGrid.Rows <= 1 Then Exit Property
    If mDataGrid.RowSel <= 0 Or mDataGrid.RowSel >= mDataGrid.Rows Then Exit Property
    If mDataGrid.RowHidden(mDataGrid.RowSel) = True Then Exit Property
    
    IsSelectionRow = True
End Property



'取得选择行所在索引
Property Get SelectionRow() As Long
    SelectionRow = mDataGrid.RowSel
End Property



Property Get ShowingRowCount() As Integer
'取得当前显示行的数量
    Dim i As Long
    Dim lngCount As Long
    
    lngCount = 0
    For i = 1 To mDataGrid.Rows - 1
        If Not mDataGrid.RowHidden(i) Then lngCount = lngCount + 1
    Next i
    
    ShowingRowCount = lngCount
End Property




Property Get ShowingDataRowCount() As Integer
'取得当前显示行的数据行数量
    Dim i As Long
    Dim lngCount As Long
    
    lngCount = 0
    For i = 1 To mDataGrid.Rows - 1
        If Not mDataGrid.RowHidden(i) Then
            If Not IsNullRow(i) Then
                lngCount = lngCount + 1
            End If
        End If
    Next i
    
    ShowingDataRowCount = lngCount
End Property



'取得选择列索引
Property Get SelectionCol() As Long
    SelectionCol = mDataGrid.ColSel
End Property


'配置combox列表
Property Get ComboxListFormat(ByVal lngCol As Long) As String
    ComboxListFormat = mDataGrid.ColComboList(lngCol)
End Property


Property Let ComboxListFormat(ByVal lngCol As Long, ByVal value As String)
    mDataGrid.ColComboList(lngCol) = value
End Property




'单元格颜色
Property Get CellColor(ByVal lngRow As Long, ByVal lngCol As Long) As OLE_COLOR
    CellColor = mDataGrid.Cell(flexcpBackColor, lngRow, lngCol)
End Property

Property Let CellColor(ByVal lngRow As Long, ByVal lngCol As Long, value As OLE_COLOR)
    mDataGrid.Cell(flexcpBackColor, lngRow, lngCol) = value
End Property



'设置合并方式
Property Get MergeCellStyle() As MergeSettings
    MergeCellStyle = mDataGrid.MergeCells
End Property

Property Let MergeCellStyle(value As MergeSettings)
    mDataGrid.MergeCells = value
End Property



'列表是否允许编辑
Property Get Editable() As EditableSettings
    Editable = mDataGrid.Editable
End Property

Property Let Editable(value As EditableSettings)
    mDataGrid.Editable = value
End Property



'行隐藏状态
Property Get RowHidden(ByVal lngRow As Long) As Boolean
    RowHidden = mDataGrid.RowHidden(lngRow)
End Property

Property Let RowHidden(ByVal lngRow As Long, value As Boolean)
    vfgData.RowHidden(lngRow) = value
End Property



'数据文本
Property Get Text(ByVal lngRow As Long, ByVal strColName As String) As String
    Dim iCol As Long
    
    Text = ""
    
    iCol = GetColIndex(strColName)
    
    If iCol <= -1 Then Exit Function
    If lngRow >= mDataGrid.Rows Then Exit Function
    
    
    Text = mDataGrid.TextMatrix(lngRow, iCol)
End Property


Property Let Text(ByVal lngRow As Long, ByVal strColName As String, ByVal strValue As String)
    Dim iCol As Long
'    Dim blnFind As Boolean
'    Dim chkState As CheckState
'    Dim strConvertValue As String
'    Dim lngImgIndex As Long
'    Dim strData As String

    iCol = GetColIndex(strColName)
    
    If iCol <= -1 Then Exit Property
    If lngRow >= mDataGrid.Rows Then Exit Property
    
'    strData = GetFieldDataValue(strColName, strValue, blnFind)
'    strData = IIf(strData = "", strValue, strData)
'
'
'    Call GetFieldDisplayText(strColName, strData, blnFind, chkState, strConvertValue, lngImgIndex)
'    Call UpdateCellStyle(lngRow, iCol, lngImgIndex, chkState)
    

    Call WriteText(lngRow, iCol, strValue)
End Property


'单元格标记
Property Get CellTag(ByVal lngRow As Long, ByVal strColName As String) As String
    Dim iCol As Long
    
    CellTag = ""
    
    iCol = GetColIndex(strColName)
    
    If iCol <= -1 Then Exit Function
    If lngRow >= mDataGrid.Rows Then Exit Function
    
    CellTag = mDataGrid.Cell(flexcpData, lngRow, iCol)
End Property

Property Let CellTag(ByVal lngRow As Long, ByVal strColName As String, ByVal strValue As String)
    Dim iCol As Long
    
    iCol = GetColIndex(strColName)
    
    If iCol <= -1 Then Exit Property
    If lngRow >= mDataGrid.Rows Then Exit Property
    
    mDataGrid.Cell(flexcpData, lngRow, iCol) = strValue
End Property


'获取排序列
Property Get SortCol() As Long
    SortCol = mlngSortCol
End Property

'获取排序方式
Property Get SortWay() As Long
    SortWay = mlngSortWay
End Property



Public Sub ResetSort(ByVal lngCol As Long, ByVal lngWay As Long)
'重置排序
    vfgData.Col = lngCol
    vfgData.Sort = lngWay
    
    '更新列表编号
    Call UpdateRowNumber
End Sub


Private Sub GetColProperty(ByVal lngColIndex As Long)
On Error GoTo errHandle
    Set mobjTmpDictionary = mDataGrid.Cell(flexcpData, 0, lngColIndex)
    Exit Sub
errHandle:
    Set mobjTmpDictionary = Nothing
End Sub

Private Function RefreshColDicObject(ByVal lngCol As Long) As Boolean
'刷新列字典对象
    If mlngCurColProIndex <> lngCol Then
        mlngCurColProIndex = lngCol
        
        Call GetColProperty(lngCol)
    End If
    
    RefreshColDicObject = IIf(mobjTmpDictionary Is Nothing, False, True)
End Function

Private Sub UpdateCellStyle(ByVal lngRow As Long, ByVal lngCol As Long, ByVal lngImgIndex As Long, ByVal chkState As CheckState)
'刷新单元格样式
    Dim strValue As String
    
    If Not RefreshColDicObject(lngCol) Then Exit Sub
    
    strValue = ""
    If Not IsCheckboxCol(lngCol) Then
        strValue = mobjTmpDictionary(TColPro.cpDataImgIndex)
        If strValue <> "" And mDataGrid.Cell(flexcpText, lngRow, lngCol) <> "" Then
            Set mDataGrid.Cell(flexcpPicture, lngRow, lngCol) = GetImg(Nvl(strValue))
        Else
            Set mDataGrid.Cell(flexcpPicture, lngRow, lngCol) = Nothing
        End If
    End If
    
    '设置图像
    If lngImgIndex > 0 Then
        mDataGrid.Cell(flexcpPicture, lngRow, lngCol) = GetImg(lngImgIndex)
    Else
        If strValue = "" Then mDataGrid.Cell(flexcpPicture, lngRow, lngCol) = Nothing
    End If
                    
    '如果是check列,则设置check图标
    If IsCheckboxCol(lngCol) Then
        mDataGrid.Cell(flexcpPicture, lngRow, lngCol) = imgCheck(0)
    End If
        
    
    '设置当前的check状态
    If chkState = csCheck Then
        mDataGrid.Cell(flexcpPicture, lngRow, lngCol) = imgCheck(1)
    ElseIf chkState = csNoCheck Then
        mDataGrid.Cell(flexcpPicture, lngRow, lngCol) = imgCheck(0)
    End If
End Sub



'取得当前选定行中指定列的显示内容
Property Get CurText(ByVal strColName As String) As String
    CurText = ""
    
    If mDataGrid.RowSel <= 0 Then Exit Property
    
    CurText = Text(mDataGrid.RowSel, strColName)
End Property


Property Let CurText(ByVal strColName As String, ByVal strValue As String)

    Text(mDataGrid.RowSel, strColName) = strValue
End Property


'当前选中单元格标记
Property Get CurCellTag(ByVal strColName As String) As String
    CurCellTag = CellTag(mDataGrid.RowSel, strColName)
End Property


Property Let CurCellTag(ByVal strColName As String, ByVal strValue As String)
    CellTag(mDataGrid.RowSel, strColName) = strValue
End Property

'显示文本
Property Get DisplayText(ByVal lngRow As Long, ByVal strColName As String) As String
    DisplayText = mDataGrid.Cell(flexcpTextDisplay, lngRow, GetColIndex(strColName))
End Property



'行关键字
Property Get KeyValue(ByVal lngRow As Long) As String
    KeyValue = mDataGrid.TextMatrix(lngRow, GetColIndex(mstrKeyName))
End Property

Property Let KeyValue(ByVal lngRow As Long, ByVal value As String)
    mDataGrid.TextMatrix(lngRow, GetColIndex(mstrKeyName)) = value
End Property


'取得当前行关键值
Property Get CurKeyValue() As String
    CurKeyValue = ""
    
    If mDataGrid.Rows <= 1 Then Exit Property
    If mDataGrid.RowSel <= 0 Or mDataGrid.RowSel >= mDataGrid.Rows Then Exit Property
    
    CurKeyValue = KeyValue(mDataGrid.RowSel)
End Property

Property Let CurKeyValue(ByVal strValue As String)
    If mDataGrid.Rows <= 1 Then Exit Property
    If mDataGrid.RowSel <= 0 Or mDataGrid.RowSel >= mDataGrid.Rows Then Exit Property
    
    KeyValue(mDataGrid.RowSel) = strValue
End Property






'是否允许del键移除数据
Property Get IsDelKeyRemoveData() As Boolean
    IsDelKeyRemoveData = mblnIsDelKeyRemoveData
End Property

Property Let IsDelKeyRemoveData(value As Boolean)
    mblnIsDelKeyRemoveData = value
End Property



'是否使用ado数据复制模式
Property Get IsCopyMode() As Boolean
    IsCopyMode = mblnIsCopyAdoMode
End Property


Property Let IsCopyMode(value As Boolean)
    mblnIsCopyAdoMode = value
End Property



'是否启用用右键弹出列表配置窗体
Property Get IsEjectConfig() As Boolean
    IsEjectConfig = mblnIsEjectConfig
End Property


Property Let IsEjectConfig(value As Boolean)
    mblnIsEjectConfig = value
End Property




'ado过滤条件
Property Get AdoFilter() As String
    AdoFilter = mstrAdoFilter
End Property


Property Let AdoFilter(value As String)
    If value = mstrAdoFilter Then Exit Property
    
    mstrAdoFilter = value
    
    If Not (mrsData Is Nothing) Then
        mrsData.Filter = mstrAdoFilter
    End If
End Property




'ado数据集
Property Get AdoData() As ADODB.Recordset
    Set AdoData = mrsData
End Property

Property Set AdoData(value As ADODB.Recordset)
    If value Is Nothing Then
        Set mrsData = Nothing
        Exit Property
    End If
    
    If mblnIsCopyAdoMode Then
        Set mrsData = zlDatabase.CopyNewRec(value)
    Else
        Set mrsData = value
    End If
    
    mrsData.Filter = mstrAdoFilter
End Property

 
 
 '数据列转换格式配置
Property Get ColConvertFormat() As String
    ColConvertFormat = mstrDataConvertFormat
End Property

Property Let ColConvertFormat(value As String)
    If mstrDataConvertFormat = value Then Exit Property
    
    mstrDataConvertFormat = value
    
    '重新配置转换字典
    Call ConfigFieldConvertDictionary
End Property

 
 

'回车是否跳转到下一单元格
Property Get IsEnterNextCell() As Boolean
    IsEnterNextCell = mblnIsEnterNextCell
End Property

Property Let IsEnterNextCell(value As Boolean)
    mblnIsEnterNextCell = value
End Property


'btn执行后，是否跳转到下一单元格
Property Get IsBtnNextCell() As Boolean
    IsBtnNextCell = mblnIsBtnNextCell
End Property


Property Let IsBtnNextCell(value As Boolean)
    mblnIsBtnNextCell = value
End Property




'背景颜色
Property Get BackColor() As OLE_COLOR
    BackColor = vfgData.BackColor
End Property

Property Let BackColor(value As OLE_COLOR)
    vfgData.BackColor = value
End Property




'是否显示行号
Property Get IsShowRowNumber() As Boolean
    IsShowRowNumber = mblnIsShowNumber
End Property


Property Let IsShowRowNumber(value As Boolean)
    If value = mblnIsShowNumber Then Exit Property
    
    mblnIsShowNumber = value
    Call UpdateRowNumber
End Property




'不可编辑单元格颜色
Property Get DisCellColor() As OLE_COLOR
    DisCellColor = mlngDisableColor
End Property

Property Let DisCellColor(value As OLE_COLOR)
    If mlngDisableColor = value Then Exit Property
    
    mlngDisableColor = value
    Call RefreshReadColColor
End Property



'错误单元格颜色
Property Get ErrCellColor() As OLE_COLOR
    ErrCellColor = mlngErrCellColor
End Property

Property Let ErrCellColor(value As OLE_COLOR)
    mlngErrCellColor = value
End Property








'Grid列数量
Property Get GridCols() As Long
    GridCols = mDataGrid.Cols
End Property


'Grid行数量
Property Get GridRows() As Long
    GridRows = mDataGrid.Rows
End Property


Property Let GridRows(value As Long)
    mDataGrid.Rows = value
    
    If mblnIsKeepRows Then mlngKeepRows = value
    
    Call UpdateRowNumber
    Call RefreshReadColColor
    Call RefreshAlign
End Property


'鼠标所在行索引
Property Get MouseRowIndex() As Long
    MouseRowIndex = mDataGrid.MouseRow
End Property


'是否保持Grid行数量
Property Get IsKeepRows() As Boolean
    IsKeepRows = mblnIsKeepRows
End Property

Property Let IsKeepRows(value As Boolean)
    If mblnIsKeepRows = value Then Exit Property
    
    mblnIsKeepRows = value
    mlngKeepRows = IIf(value, vfgData.Rows, -1)
    
    If Not value Then mDataGrid.Rows = 1
End Property




'grid属性对象
Property Get DataGrid() As VSFlexGrid
    Set DataGrid = mDataGrid
End Property

Property Set DataGrid(value As VSFlexGrid)
    Set mDataGrid = value
End Property


'原始的Grid
Property Get SourceGrid() As VSFlexGrid
    Set SourceGrid = vfgData
End Property



Private Sub UpdateRowNumber()
'更新列表编号
    Dim i As Long
    Dim lngNumber As Long
    Dim lngTxtWidth As Long
    
    If Not mblnIsAllowExtCol Then Exit Sub
    
    If mblnIsAllowExtCol Then
        lngNumber = 1
        For i = 1 To mDataGrid.Rows - 1
            If Not mDataGrid.RowHidden(i) Then
                mDataGrid.TextMatrix(i, 0) = IIf(mblnIsShowNumber, lngNumber, "")
                lngNumber = lngNumber + 1
            End If
        Next i
        
        UserControl.Font.Size = mDataGrid.Font.Size
        
        lngTxtWidth = TextWidth(lngNumber) + 120
        
        mDataGrid.ColWidth(0) = IIf(lngTxtWidth >= 240, lngTxtWidth, 240)
        
        Call vfgData_AfterUserResize(0, 0)
    End If
    
End Sub


Public Sub GetFieldDisplayText(ByVal strColName As String, ByVal strCurCode As String, _
    ByRef blnFind As Boolean, ByRef chkState As CheckState, ByRef strText As String, ByRef lngImgIndex As Long)
'取得字段转换对应的值
    Dim strTemp As String
    Dim lngSourceIndex As Long
    Dim strMatch As String
    Dim lngMatchIndex As Long
    Dim strMatchValue As String
    
    blnFind = False
    
    strText = strCurCode
    chkState = csNone
    lngImgIndex = -1
    
    If strCurCode = "" Then Exit Sub
    If mobjColDictionary Is Nothing Then Exit Sub
    
    '如果该列无配置的转换数据，则退出过程
    If Not mobjColDictionary.Exists(strColName) Then Exit Sub
    
    '如果要查找的值不存在，则判断是否存在“els”
    blnFind = mobjColDictionary(strColName).Exists(strCurCode)
    strTemp = ""
    
    If Not blnFind Then
        blnFind = mobjColDictionary(strColName).Exists("els")
        
        If blnFind Then strTemp = mobjColDictionary(strColName)("els")
    Else
        strTemp = mobjColDictionary(strColName)(strCurCode)
    End If
    
    If Not blnFind Then Exit Sub
    
    
    If InStr(1, UCase(strTemp), UCase("<check>")) > 0 Then
        chkState = csCheck
        strTemp = Replace(strTemp, "<check>", "")
    End If
    
    If InStr(1, UCase(strTemp), UCase("<nocheck>")) > 0 Then
        chkState = csNoCheck
        strTemp = Replace(strTemp, "<nocheck>", "")
    End If
    
    
    '获取对应的图片索引
    lngImgIndex = -1
    If InStr(1, UCase(strTemp), UCase("<img")) > 0 Then
        strMatchValue = InstrEx(strTemp, "<img*>", strMatch, lngMatchIndex)
        If strTemp <> "" Then
            lngImgIndex = Nvl(strMatch, -1)
            strTemp = Replace(strTemp, strMatchValue, "")
        End If
    End If
    
    
    '使用原数据值
    lngSourceIndex = InStr(1, UCase(strTemp), UCase("<source>"))
    If lngSourceIndex > 0 Then
        strTemp = Mid(strTemp, 1, lngSourceIndex - 1) & strCurCode & Mid(strTemp, lngSourceIndex + Len("<source>"), 100)
    End If
    
    '判断是否使用原值
    strText = strTemp
End Sub


Public Function GetColsString(ufgData As Object) As String
'得到列名参数
    Dim i As Integer
    Dim strString As String
    Dim strTemp As String
    Dim strProperty As String
    Dim objUfgColPro As Scripting.Dictionary

    strString = ""
    
    For i = 1 To ufgData.GridCols - 1
        If strString = "" Then
            strString = "|"
        End If
        
        Set objUfgColPro = ufgData.DataGrid.Cell(flexcpData, 0, i)
        
        If Not objUfgColPro Is Nothing Then
            strProperty = objUfgColPro(TColPro.cpProperty)
            
            '处理字符串
            strProperty = Mid(strProperty, InStrRev(strProperty, "@") + 1)
            '判断列属性字符串是否包含默认列宽属性,如果是 则删除，不是则 跳过
            If InStr(strProperty, "w") Then
                strTemp = Mid(strProperty, InStr(1, strProperty, "w"), 100)
                '判断列宽属性后有无逗号，有 则进行处理，无 则跳过
                 If InStr(strTemp, ",") Then
                    strTemp = Mid(strTemp, 1, InStr(2, strTemp, ",") - 1)
                 End If
                
                '使用Replace去掉默认列宽属性
                strProperty = Replace(strProperty, "," & strTemp, "")
            End If
            '连接字符串
            strString = strString + strProperty & ",w" & ufgData.DataGrid.ColWidth(i) & "|"
        End If
    Next
    
    GetColsString = strString

End Function



Private Sub cmdCellBtn_Click()
On Error Resume Next

    RaiseEvent OnCellButtonClick(mDataGrid.Row, mDataGrid.Col)
    
    err.Clear
End Sub

Private Sub mnuCopy_Click()
On Error Resume Next
    Clipboard.Clear
    Call Clipboard.SetText(mDataGrid.Text)
End Sub

Private Sub mnuCut_Click()
On Error Resume Next
    Clipboard.Clear
    
    Call Clipboard.SetText(mDataGrid.Text)
    mDataGrid.Text = ""
End Sub

Private Sub mnuDel_Click()
On Error Resume Next
    mDataGrid.Text = ""
End Sub

Private Sub mnuPaste_Click()
On Error Resume Next
    mDataGrid.Text = Clipboard.GetText
End Sub


Private Sub TimerRefreshData_Timer()
On Error Resume Next
    Dim i As Long
    Dim rsBind As ADODB.Recordset
    
    TimerRefreshData.Enabled = False
    
    Set rsBind = mDataGrid.DataSource
    rsBind.MoveFirst
    
    For i = 1 To mDataGrid.Rows - 1
        RaiseEvent OnRefreshRowData(rsBind, i)
        
        '处理消息循环，使其能够响应用户操作
        If i Mod 10 = 0 Then DoEvents
    Next i
    
    err.Clear
End Sub

Private Sub vfgData_AfterMoveColumn(ByVal Col As Long, Position As Long)
On Error Resume Next
'拖动交换列位置的时候，重新计算位置方法
    If GetColIndexWithRowCheck > 0 Then
        chkCheckAll.Left = vfgData.Cell(flexcpLeft, 0, GetColIndexWithRowCheck()) + 60
     End If
     
    mlngCurColProIndex = -1
    Call ShowCellButton
     
    RaiseEvent OnColFormartChange
    
    err.Clear
End Sub


Private Sub vfgData_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
On Error Resume Next
'滚动条滚动后，调用重新计算位置方法
    Dim blnLastVisible As Boolean
    Dim blnFirstVisible As Boolean
    Dim lngRowCheckColIndex As Long
    Dim lngHideRowCount As Long
    
    If GetColIndexWithRowCheck > 0 Then
        
        lngRowCheckColIndex = GetColIndexWithRowCheck()
        lngHideRowCount = GetHideRowCount()
        
        blnLastVisible = True
        If GetColIndexWithRowCheck + 1 < vfgData.Cols Then
            blnLastVisible = IIf(vfgData.Cell(flexcpLeft, 0, lngRowCheckColIndex) < vfgData.Cell(flexcpLeft, 0, lngRowCheckColIndex + lngHideRowCount + 1), True, False)
        End If
        
        blnFirstVisible = True
        If GetColIndexWithRowCheck - 1 >= 0 Then
            blnFirstVisible = IIf(vfgData.Cell(flexcpLeft, 0, lngRowCheckColIndex) > vfgData.Cell(flexcpLeft, 0, lngRowCheckColIndex - lngHideRowCount - 1), True, False)
        End If
        
        chkCheckAll.Visible = vfgData.Cell(flexcpLeft, 0, lngRowCheckColIndex) >= vfgData.Cell(flexcpWidth, 0, 0) And blnLastVisible And blnFirstVisible
    
        chkCheckAll.Left = vfgData.Cell(flexcpLeft, 0, lngRowCheckColIndex) + 60
    End If
    
    Call ShowCellButton
    
    err.Clear
End Sub


Private Sub vfgData_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim objColPro As Scripting.Dictionary
    
    On Error Resume Next
    
    'Col为-1是纵向拖动，反之是横向拖动
    If Col = -1 Then
        '重新计算checkBox 的Top值
        chkCheckAll.Top = vfgData.Cell(flexcpHeight, 0, GetColIndexWithRowCheck()) / 2 - 70
    Else
        '列头每列宽度不能小于60缇
        If vfgData.Cell(flexcpWidth, 0, Col) < 240 Then vfgData.ColWidth(Col) = 60
        
        '动态改变checkBox 的Left值
        If GetColIndexWithRowCheck > 0 Then
            chkCheckAll.Left = vfgData.Cell(flexcpLeft, 0, GetColIndexWithRowCheck()) + 60
            chkCheckAll.Visible = IIf(vfgData.Cell(flexcpWidth, 0, GetColIndexWithRowCheck) < 240, False, True)
        End If
        
        Set objColPro = vfgData.Cell(flexcpData, 0, Col)
        objColPro(TColPro.cpWidth) = vfgData.ColWidth(Col)
    End If
    
    Call ShowCellButton
    
    If Not (mblnIsAllowExtCol And Col = 0) Then RaiseEvent OnColFormartChange
 
    err.Clear
End Sub


Private Sub ShowCellButton()
    Dim lngCol As Long
    
    cmdCellBtn.Visible = False

    If mDataGrid.Row < 0 Or mDataGrid.Col < 0 Then Exit Sub
    
    lngCol = -1
    If IsButtonCol(mDataGrid.Col) Then
        lngCol = mDataGrid.Col
    Else
        lngCol = GetColIndexWithBtn
    End If
    
    If lngCol < 0 Then Exit Sub
    
    If mDataGrid.Cell(flexcpTop, mDataGrid.Row) <= 0 Then Exit Sub
    If mDataGrid.Cell(flexcpLeft, mDataGrid.Row, lngCol) <= 0 Then Exit Sub
    
    Call ShowObject(cmdCellBtn, mDataGrid.Row, lngCol)
    
End Sub
   
Private Sub chkCheckAll_Click()
On Error GoTo errHandle
'全选反选CheckBox
 
    If chkCheckAll.value = 0 Then
       Call ClearCellCheck(GetColIndexWithRowCheck())
    Else
       Call CheckAllCell(GetColIndexWithRowCheck())
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub vfgData_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
On Error Resume Next
    RaiseEvent OnBeforeRowColChange(OldRow, OldCol, NewRow, NewCol, Cancel)
    
    err.Clear
End Sub

Private Sub vfgData_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
    If IsUnResizeCol(Col) Then Cancel = True
    
    err.Clear
End Sub

Private Sub vfgData_DblClick()
On Error Resume Next
    If vfgData.MouseRow < 1 Then Exit Sub
    
    RaiseEvent OnDblClick
    
    err.Clear
End Sub


Public Function GetFieldDataValue(ByVal strColName As String, ByVal strCurValue As String, ByRef blnFind As Boolean) As String
'取得字段转换对应的代码
    Dim strTemp As String
    Dim lngValueIndex As Long
    Dim lngFieldIndex As Long
    Dim strFindValue As String
    Dim strReplace As String
    Dim strMatch As String
    Dim strMatchValue As String
    Dim lngMatchIndex As Long
    
    blnFind = False
    GetFieldDataValue = ""
    
    If mstrDataConvertFormat = "" Then Exit Function
    If strCurValue = "" Then Exit Function
    
    '判断该字段是否配置有转换的数据
    lngFieldIndex = InStr(1, UCase(mstrDataConvertFormat), UCase(strColName))
    If lngFieldIndex <= 0 Then Exit Function
    
    strFindValue = strCurValue
    
    '传进来的值可能为1-xxx之类
    If Mid(strFindValue, 2, 1) = "-" Or Mid(strFindValue, 3, 1) = "-" Then
        strFindValue = Mid(strFindValue, InStr(strFindValue, "-") + 1, 100)
    End If
    
    strTemp = Mid(mstrDataConvertFormat, lngFieldIndex + Len(strColName & ":"), 1000) & "|"
    strTemp = Mid(strTemp, 1, InStr(1, strTemp, "|"))
    strTemp = Replace(strTemp, "|", ",")
    
    If strTemp = "" Then Exit Function
    
    blnFind = True
    
    
    '替换转换配置数据中的转义符，如"<check>"和"<nocheck>"
    If UCase(strFindValue) <> UCase(M_STR_ConvertProp_Check) And _
        UCase(strFindValue) <> UCase(M_STR_ConvertProp_NoCheck) Then
        
        If InStr(1, UCase(strTemp), UCase("<check>")) > 0 Then
            strTemp = Replace(strTemp, "<check>", "")
        End If
        
        If InStr(1, UCase(strTemp), UCase("<nocheck>")) > 0 Then
            strTemp = Replace(strTemp, "<nocheck>", "")
        End If
        
        While InStr(1, UCase(strTemp), UCase("<img")) > 0
            strMatchValue = InstrEx(strTemp, "<img*>", strMatch, lngMatchIndex)
            If strMatchValue <> "" Then strTemp = Replace(strTemp, strMatchValue, "")
        Wend
            
        '查找值的形式需要如下形式："-查找值,"
        lngValueIndex = InStr(1, strTemp, "-" & strFindValue & ",")
        If lngValueIndex > 0 Then strTemp = Mid(strTemp, 1, lngValueIndex - 1)  'strTemp此时的内容格式为:“0-值x,1”
    Else
        lngValueIndex = InStr(1, UCase(strTemp), UCase(strFindValue))
        If lngValueIndex > 0 Then
        
            strTemp = Mid(strTemp, 1, lngValueIndex - 1)
        
            strReplace = Mid(strTemp, InStrRev(strTemp, "-"), 100)
            strTemp = Replace(strTemp, strReplace, "")
        End If
    End If
    
    
    '判断是否找到对应的转换数据
    If lngValueIndex <= 0 Then
        lngValueIndex = InStr(1, UCase(strTemp), UCase(M_STR_ConvertProp_Els) & "-")
        If lngValueIndex <= 0 Then Exit Function
        
        strTemp = ""
        
        lngValueIndex = InStr(1, UCase(strTemp), UCase(M_STR_ConvertProp_Source))
        If lngValueIndex <= 0 Then Exit Function
        
        strTemp = "," & strCurValue
    End If
    
    
    strTemp = Mid(strTemp, InStrRev(strTemp, ",") + 1, 100)
    
    
    GetFieldDataValue = strTemp
End Function

Private Function RegReplace(ByVal strSource As String, ByVal strFind As String, ByVal strReplace As String) As String
'使用正则表达式进行替换
    
    mobjRegExp.Pattern = strFind
    
    mobjRegExp.Global = True
    mobjRegExp.IgnoreCase = True
    mobjRegExp.MultiLine = True
    
    RegReplace = mobjRegExp.Replace(strSource, strReplace)
End Function

Public Sub BindData()
'直接绑定数据
    Dim rsBind As ADODB.Recordset
    Dim rsClone As ADODB.Recordset
    
    Dim blnFilterOut As Boolean
    
    Dim adoSourceStream As ADODB.Stream
    Dim adoNewStream As ADODB.Stream
    
    Dim strData As String
    Dim strSchema As String
    Dim strNewPro As String
    Dim strColName As String
    Dim strFieldName As String
    Dim lngStartPos As Long
    
    Dim strOldDsFilter As String
    Dim strBindFilter As String
    Dim strCloneFilter As String
    
    Dim i As Long
    Dim aryColPro() As Scripting.Dictionary '该变量用于恢复表头列的属性设置
    
    If mrsData.RecordCount <= 0 Then
        mDataGrid.Rows = 1
        Exit Sub
    End If
        
        
    strOldDsFilter = mrsData.Filter
    
    RaiseEvent OnBindFilter(strBindFilter, strCloneFilter)
    
    If strBindFilter <> "" Then
        strBindFilter = IIf(strOldDsFilter = "" Or strOldDsFilter = "0", strBindFilter, strOldDsFilter & " and " & strBindFilter)
    Else
        strBindFilter = IIf(strOldDsFilter = "0", "", strOldDsFilter)
    End If
    
    If strCloneFilter <> "" Then
        strCloneFilter = IIf(strOldDsFilter = "" Or strOldDsFilter = "0", strCloneFilter, strOldDsFilter & " and " & strCloneFilter)
    Else
        strCloneFilter = IIf(strOldDsFilter = "0", "", strOldDsFilter)
    End If


    
    Set adoSourceStream = New ADODB.Stream
    adoSourceStream.type = adTypeText
    adoSourceStream.Mode = adModeRead
    
    If strCloneFilter <> mrsData.Filter Then
        mrsData.Filter = strCloneFilter
        mrsData.Save adoSourceStream, adPersistXML
    
        Set rsClone = New ADODB.Recordset
        rsClone.Open adoSourceStream
    Else
        Set rsClone = mrsData
    End If

    If strBindFilter <> mrsData.Filter Then mrsData.Filter = strBindFilter
    
    '恢复数据集之前的过滤条件
    mrsData.Filter = IIf(strOldDsFilter = "0", "", strOldDsFilter)
    
    If adoSourceStream.State = adStateOpen Then Call adoSourceStream.Close
    Call mrsData.Save(adoSourceStream, adPersistXML)
    
    
    
    
    adoSourceStream.Position = 0
    strData = adoSourceStream.ReadText
    
    lngStartPos = InStr(strData, "<s:AttributeType")
    
    '获取数据集结构数据
    strSchema = Mid(strData, lngStartPos, InStr(strData, "<s:extends") - lngStartPos - 1)
    strData = Replace(strData, strSchema, "")
    
    '将日期格式为“2012-03-04T12:13:14”替换为“2012-03-04 12:13:14”
    strData = RegReplace(strData, "(?!\b-\d{1,2})T(?=\d{1,2}:)", " ") 'RegReplace(strData, "\b-\d{1,2}T\d{1,2}:\b", " ")
    
    For i = mDataGrid.Cols - 1 To 0 Step -1
        If RefreshColDicObject(i) Then
            Exit For
        End If
    Next i
    
    ReDim aryColPro(i)
    
    '修改绑定数据集的结构
    strData = Replace(strData, "number='", "number='100")
    
    strNewPro = ""
    '增加绑定显示的数据列
    For i = 0 To mDataGrid.Cols - 1
        If RefreshColDicObject(i) Then
            Set aryColPro(i) = mobjTmpDictionary
            
            strColName = mobjTmpDictionary(TColPro.cpColName)
            strFieldName = mobjTmpDictionary(TColPro.cpFieldName)
            
            If strColName <> M_STR_AdjustColName Then
                If strNewPro <> "" Then strNewPro = strNewPro & vbCrLf
                strNewPro = strNewPro & "<s:AttributeType name='" & strColName & "' rs:number='" & i + 1 & "' rs:nullable='true' rs:writeunknown='true'>" & _
                        "<s:datatype dt:type='string' rs:dbtype='str' rs:scale='0' rs:precision='3' rs:fixedlength='true'/>" & _
                        "</s:AttributeType>"
            
                If strFieldName <> M_STR_PlaceCol Then
                    '将字段名替换为显示名称
                    strSchema = RegReplace(strSchema, _
                                "<s:AttributeType name='" & strFieldName & "'[^#]*?:AttributeType>", _
                                "")
                End If
                                        
'                strData = Replace(strData, strFieldName & "='", strColName & "='")
                If strFieldName <> strColName And strFieldName <> M_STR_PlaceCol Then
                    strData = RegReplace(strData, strFieldName & "='", strColName & "='")
                End If
            End If
        Else
            Exit For
        End If
    Next i
    
    strData = RegReplace(strData, "rs:ReshapeName='DSRowset1_\d*'>", "rs:ReshapeName='DSRowset1_125'>" & vbCrLf & strNewPro & strSchema)
    
    
    Set rsBind = New ADODB.Recordset
    
    Set adoNewStream = New ADODB.Stream
    adoNewStream.type = adTypeText
    adoNewStream.Mode = adModeReadWrite
    
    '读取修改后的流数据
    adoNewStream.Open
    adoNewStream.WriteText strData
    adoNewStream.Position = 0
    
    rsBind.Open adoNewStream
    If rsBind.RecordCount > 0 Then rsBind.MoveFirst
    
    '配置需要绑定显示的数据
    While Not rsBind.EOF
        blnFilterOut = False
        RaiseEvent OnFilterRowData(rsBind, rsClone, blnFilterOut)
        
        '如果没有将数据排除在外，则复制数据到绑定数据集
        If blnFilterOut Then
            rsBind.Delete
        End If
        
        rsBind.MoveNext
    Wend
    
    mDataGrid.FixedCols = IIf(mblnIsAllowExtCol, 1, 0)
    
    '处理消息循环
    DoEvents
    
    
    '加载数据到列表显示
    Set mDataGrid.DataSource = rsBind
    Call mDataGrid.DataRefresh
    
    '处理消息循环
    DoEvents
    
    '恢复列属性(刷新绑定数据列后，需要重新恢复列的显示状态)
    For i = 0 To mDataGrid.Cols - 1
        If i <= UBound(aryColPro) Then
            Set mDataGrid.Cell(flexcpData, 0, i) = aryColPro(i)
            
            If aryColPro(i)(TColPro.cpHeadImgIndex) > -1 And Not aryColPro(i)(TColPro.cpIsRowCheck) Then
                Set mDataGrid.Cell(flexcpPicture, 0, i) = GetImg(aryColPro(i)(TColPro.cpHeadImgIndex))
                
                If Not mobjImageList Is Nothing Then
                    If ScaleY(mobjImageList.ImageHeight, vbPixels, vbTwips) > vfgData.RowHeight(0) Then
                        mDataGrid.RowHeight(0) = ScaleY(mobjImageList.ImageHeight, vbPixels, vbTwips) + 120
                    End If
                End If
            End If
    
            If aryColPro(i)(TColPro.cpWidth) > 0 Then mDataGrid.ColWidth(i) = aryColPro(i)(TColPro.cpWidth)
            
            '设置列的关键字
            mDataGrid.ColKey(i) = aryColPro(i)(TColPro.cpColName)
                
            '设置列的对齐方式
            If Val(aryColPro(i)(TColPro.cpColAlign)) = flexAlignRightCenter Then
                mDataGrid.Cell(flexcpAlignment, 0, i) = flexAlignRightCenter
                
            ElseIf Val(aryColPro(i)(TColPro.cpColAlign)) = flexAlignCenterCenter Then
                mDataGrid.Cell(flexcpAlignment, 0, i) = flexAlignCenterCenter
    
            ElseIf Val(aryColPro(i)(TColPro.cpColAlign)) = flexAlignLeftCenter Then
                mDataGrid.Cell(flexcpAlignment, 0, i) = flexAlignLeftCenter
                
            End If
                        
            
            '隐藏列
            If aryColPro(i)(TColPro.cpIsHide) Then
                mDataGrid.ColHidden(i) = True
            End If
                    
            'button列
            If aryColPro(i)(TColPro.cpIsBtn) Then
                mDataGrid.ColComboList(i) = "..." '不能使用“…”符号
            End If
            
            '合并列
            If aryColPro(i)(TColPro.cpIsMerage) Then
                mDataGrid.MergeCol(i) = True
            End If
            
            
            '设置该列为combox列
            If aryColPro(i)(TColPro.cpIsCombox) Then
                mDataGrid.ColComboList(i) = aryColPro(i)(TColPro.cpComboxCfg)
            End If
            
            '设置该列为扩展调整列
            If aryColPro(i)(TColPro.cpColName) = M_STR_AdjustColName Then
                mDataGrid.ColWidth(i) = 500
                mDataGrid.ColAlignment(i) = flexAlignCenterCenter
            End If
        Else
            Set mDataGrid.Cell(flexcpData, 0, i) = Nothing
            
            mDataGrid.ColHidden(i) = True
        End If
    Next i
    
    Call UpdateRowNumber
    Call RefreshReadColColor
    Call RefreshAlign
    
    '处理消息循环
    DoEvents
    
    '定位到第一行数据
    Call LocateRow(1)
    
    '启动行数据刷新
    TimerRefreshData.Enabled = True
End Sub


Public Sub RefreshData(Optional ByVal blnDelHistoryList As Boolean = True)
'刷新数据显示
    Dim lngStartRow As Long
    Dim blnContinue As Boolean
    Dim blnFilterOut As Boolean
    Dim rsClone As ADODB.Recordset
    Dim lngFilterOutCount As Long
    Dim lngRecordCount As Long
    
    If blnDelHistoryList Then
        Call ClearListData
'        Call UpdateRowNumber   '在ClearListData方法中调用了UpdateRowNumber方法，因此不需要重复调用
        Call RefreshReadColColor
        
        Call RefreshAlign
    End If
        
    
    If mrsData Is Nothing Then Exit Sub
    
    lngRecordCount = mrsData.RecordCount
    
    If lngRecordCount <= 0 Then Exit Sub
    
    lngFilterOutCount = 0
    lngStartRow = IIf(blnDelHistoryList, 0, IIf(mblnIsKeepRows, GetNullRowIndex - 1, mDataGrid.Rows - 1))
    Set rsClone = mrsData.Clone
    
    '设置数据显示行数
    If Not mblnIsKeepRows Then
        If blnDelHistoryList Then
            vfgData.Rows = lngRecordCount + 1
        Else
            vfgData.Rows = lngRecordCount + mDataGrid.Rows
        End If
        
        Call UpdateRowNumber
        Call RefreshReadColColor
        Call RefreshAlign
    End If
    
    mrsData.MoveFirst
    
    blnContinue = False
    Do While Not mrsData.EOF
        If mrsData.AbsolutePosition > 2000 And Not blnContinue And (lngRecordCount - 2000 > 300) Then
            If MsgBox("已完成2000行数据加载，剩余大约" & lngRecordCount - 2000 & "行数据尚未加载，是否继续？如果继续则需等待更多时间。", vbYesNo, "数据加载") = vbNo Then
                vfgData.Rows = mrsData.AbsolutePosition
                
                Call LocateRow(1)
                Exit Sub
            End If
            
            blnContinue = True
        End If
        
        blnFilterOut = False
        RaiseEvent OnFilterRowData(mrsData, rsClone, blnFilterOut)
        
        If Not blnFilterOut Then
            If blnDelHistoryList Then
                Call LoadAdoDataToList(mrsData, -lngFilterOutCount)
            Else
                Call LoadAdoDataToList(mrsData, lngStartRow - lngFilterOutCount)
            End If
        Else
'            vfgData.Rows = vfgData.Rows - 1
            lngFilterOutCount = lngFilterOutCount + 1
        End If
        
        mrsData.MoveNext
    Loop
    
    vfgData.Rows = vfgData.Rows - lngFilterOutCount
    
    '定位到第一行数据
    Call LocateRow(1)
End Sub

Public Sub LocateRow(Optional ByVal lngRowIndex As Long = -1)
'定位指定行，默认定位为最后一位
    Dim lngRow As Long
    Dim iCol As Long
    
    If mDataGrid.Rows <= 1 Then Exit Sub
    
    lngRow = lngRowIndex
    If lngRow < 0 Then
        lngRow = mDataGrid.Rows - 1
    End If
    
    '取得第一个未隐藏的列
    For iCol = IIf(mblnIsAllowExtCol, 1, 0) To mDataGrid.Cols - 1
        If Not mDataGrid.ColHidden(iCol) Then Exit For
    Next iCol
    
    Call mDataGrid.Select(lngRow, iCol)
    Call mDataGrid.ShowCell(lngRow, iCol)
End Sub


Public Sub RestoreList(Optional ByVal blnKeepRowCount As Boolean = True)
'恢复列表
    Dim R As Long
    Dim c As Long
    
    R = mDataGrid.Rows - 1
    If R = 0 Then Exit Sub
    
    While R > 0
        mDataGrid.RowData(R) = TDataRowState.Normal
        
        '恢复修改前得数据
        For c = 0 To mDataGrid.Cols - 1
            If Not mDataGrid.ColHidden(c) Then
                mDataGrid.TextMatrix(R, c) = mDataGrid.Cell(flexcpData, R, c)
            End If
        Next c
        
        '恢复删除过的数据
        If mDataGrid.RowHidden(R) Then
            If IsEmptyKey(R) Then
                Call mDataGrid.RemoveItem(R)
            Else
                mDataGrid.RowHidden(R) = False
            End If
        End If
               
        
        R = R - 1
    Wend
    
    If mblnIsAllowExtCol Then
        mDataGrid.Cell(flexcpBackColor, 1, 1, mDataGrid.Rows - 1, mDataGrid.Cols - 1) = mDataGrid.BackColor
    Else
        mDataGrid.Cell(flexcpBackColor, 1, 0, mDataGrid.Rows - 1, mDataGrid.Cols - 1) = mDataGrid.BackColor
    End If
    
    If blnKeepRowCount Then mDataGrid.Rows = IIf(mlngKeepRows <= -1, mDataGrid.Rows, mlngKeepRows)
    
    
    Call LocateRow(1)
    
    Call UpdateRowNumber
End Sub

Public Function IsErrColorWithRow(ByVal lngRow As Long, Optional blnAutoFocus As Boolean = True) As Boolean
'检查行单元格颜色是否存在错误颜色
    Dim i As Long

    IsErrColorWithRow = False
            
    For i = 0 To mDataGrid.Cols - 1
        If mDataGrid.Cell(flexcpBackColor, lngRow, i) = mlngErrCellColor Then
            IsErrColorWithRow = True
            
            If blnAutoFocus Then
                Call mDataGrid.Select(lngRow, i)
                Call mDataGrid.ShowCell(lngRow, i)
                Call mDataGrid.EditCell
            End If
            
            Exit Function
        End If
    Next i
End Function

Public Function IsErrColorWithList(Optional blnAutoFocus As Boolean = True) As Boolean
'检查更新和添加的数据是否有效
    Dim i As Long
    Dim j As Long
    
    IsErrColorWithList = False
    
    For i = 1 To mDataGrid.Rows - 1
        If RowState(i) = TDataRowState.Add Or _
            RowState(i) = TDataRowState.Update Then
            
            If IsErrColorWithRow(i, blnAutoFocus) Then
                IsErrColorWithList = True
                
                Exit Function
            End If
        End If
    Next i
End Function


Public Sub RefreshAlign(Optional ByVal lngRow As Long = -1)
'刷新数据对齐方式
    Dim i As Long
    Dim strColProperty As String

    If lngRow < 0 And mDataGrid.Rows <= 1 Then Exit Sub

    For i = IIf(mblnIsAllowExtCol, 1, 0) To mDataGrid.Cols - 1
        '扩展列的数据，使用默认的对齐设置
        If RefreshColDicObject(i) Then
            If lngRow >= 1 And lngRow < mDataGrid.Rows Then
                If Val(mobjTmpDictionary(TColPro.cpTxtAlign)) <> M_LNG_UNCFG Then
                    mDataGrid.Cell(flexcpAlignment, lngRow, i) = mobjTmpDictionary(TColPro.cpTxtAlign)
                End If
        
                If Val(mobjTmpDictionary(TColPro.cpChkAlign)) <> M_LNG_UNCFG Then
                    mDataGrid.Cell(flexcpPictureAlignment, lngRow, i) = mobjTmpDictionary(TColPro.cpChkAlign)
                End If
            Else
                If Val(mobjTmpDictionary(TColPro.cpTxtAlign)) <> M_LNG_UNCFG Then
                    mDataGrid.Cell(flexcpAlignment, 1, i, mDataGrid.Rows - 1, i) = mobjTmpDictionary(TColPro.cpTxtAlign)
                End If
        
                If Val(mobjTmpDictionary(TColPro.cpChkAlign)) <> M_LNG_UNCFG Then
                    mDataGrid.Cell(flexcpPictureAlignment, 1, i, mDataGrid.Rows - 1, i) = mobjTmpDictionary(TColPro.cpChkAlign)
                End If
            End If
        End If

    Next i
End Sub

Private Function InstrEx(ByVal strSource As String, ByVal strFind As String, ByRef strMatch As String, ByRef lngIndex As Long) As String
'使用匹配方式查找指定字符
    Dim aryFind() As String
    Dim lngCurIndex As Long
    Dim strTemp As String
    
    
    InstrEx = ""
    strMatch = ""
    lngIndex = -1
    
    lngCurIndex = InStr(strFind, "*")
    If lngCurIndex <= 0 Then
        lngCurIndex = InStr(strFind, "%")
        If lngCurIndex > 0 Then aryFind = Split(strFind, "%")
    Else
        aryFind = Split(strFind, "*")
    End If
    
    If lngCurIndex <= 0 Then
        lngCurIndex = InStr(strSource, strFind)
        If lngCurIndex >= 1 Then
            InstrEx = strFind
            lngIndex = lngCurIndex
        End If
        
        Exit Function
    End If
    
    '没有找到匹配的前部分字符
    lngCurIndex = InStr(strSource, aryFind(0))
    If lngCurIndex <= 0 Then Exit Function
    
    lngIndex = lngCurIndex
    strTemp = Mid(strSource, lngIndex + Len(aryFind(0)), Len(strSource))
    
    '没有找到匹配的后部分字符
    lngCurIndex = InStr(strTemp, aryFind(UBound(aryFind)))
    If lngCurIndex <= 0 Then Exit Function
    
    strMatch = Mid(strTemp, 1, lngCurIndex - 1)
    If InStr(strFind, "%") > 0 Then
        If Len(strMatch) <> UBound(aryFind) Then
            strMatch = ""
            lngIndex = -1
            Exit Function
        End If
    End If
    
    InstrEx = aryFind(0) & strMatch & aryFind(UBound(aryFind))
End Function


Private Function GetDataValue(rsData As ADODB.Recordset, strFieldName As String) As String
On Error GoTo errHandle
'根据数据集和列名获取转换值，如果出错返回“Err”

    GetDataValue = Nvl(rsData(strFieldName))

    Exit Function
errHandle:
    GetDataValue = "Err"
End Function


Private Sub LoadAdoDataToList(rsData As ADODB.Recordset, Optional ByVal lngStartRow As Long = 0)
'加载ado中的数据到列表
    Dim i As Integer
    Dim strFieldName As String
    Dim strColName As String
    Dim lngCurPosition As Long
    
    Dim blnFind As Boolean
    Dim strValue As String
    Dim chkState As CheckState
    Dim lngImgIndex As Long
    
    Dim strData As String
    Dim strTemp As String
    Dim strTag As String
    
    lngCurPosition = rsData.AbsolutePosition
    
    If lngCurPosition + lngStartRow >= vfgData.Rows Then Exit Sub
    
    For i = 0 To vfgData.Cols - 1
        strColName = GetColName(i)
        
        If strColName <> M_STR_AdjustColName And RefreshColDicObject(i) Then

            strFieldName = mobjTmpDictionary(TColPro.cpFieldName)
            
            If Trim(strFieldName) <> "" And strFieldName <> M_STR_PlaceCol Then
                '获取转换值
                strData = GetDataValue(rsData, strFieldName)
                
                strValue = strData
                lngImgIndex = -1
                chkState = csNone
                
                If strData <> "" And Not mobjColDictionary Is Nothing Then
                    '配置了列转换属性才执行
                    Call GetFieldDisplayText(strColName, strData, blnFind, chkState, strValue, lngImgIndex)
                End If
                                
                '设置数据列的checkbox，image等
                If mobjTmpDictionary(TColPro.cpIsUpdateStyle) Or lngImgIndex > -1 Or chkState <> csNone Then
                    Call UpdateCellStyle(lngCurPosition + lngStartRow, i, lngImgIndex, chkState)
                End If
                                
                If chkState = csNone Or strValue <> "" Then
'                    '如果没有找到转换数据，则直接读取当前的数据，并转换为指定显示格式
                    If blnFind Then
                        If mobjTmpDictionary(TColPro.cpIsCombox) Then
                            strTemp = strData & "-" & strValue
                        Else
                            strTemp = strValue
                        End If
                    Else
                        strTemp = strData
                    End If


                    strTag = strTemp
                    
'                    RaiseEvent OnBeforeReadAdoData(lngCurPosition + lngStartRow, i, strColName, strFieldName, rsData, strTemp, strTag)
                    
                    Call WriteText(lngCurPosition + lngStartRow, i, strTemp, strTag)
                    
'                    RaiseEvent OnAfterReadAdoData(lngCurPosition + lngStartRow, i, strColName, strFieldName, rsData, strTemp, strTag)
                End If

            End If
        End If
    Next i
    
    RaiseEvent OnNewRow(lngCurPosition + lngStartRow)
End Sub


Public Sub MoveUp(ByVal lngRow As Long)
'上移一行
    Dim strRowData As Variant
    Dim strRowText As Variant
    Dim varRowPic  As Variant
    
    Dim i As Long
    Dim lngUpRow As Long
    
    If lngRow <= 1 Then Exit Sub

    lngUpRow = lngRow - 1
    
    Do While lngUpRow > 0
        If vfgData.RowHidden(lngUpRow) Then
            lngUpRow = lngUpRow - 1
        Else
            Exit Do
        End If
    Loop
    
    If vfgData.RowHidden(lngUpRow) Then Exit Sub

    For i = 0 To vfgData.Cols - 1
        
        strRowText = vfgData.TextMatrix(lngUpRow, i)
        strRowData = vfgData.Cell(flexcpData, lngUpRow, i)
        Set varRowPic = vfgData.Cell(flexcpPicture, lngUpRow, i)
        
        vfgData.TextMatrix(lngUpRow, i) = vfgData.TextMatrix(lngRow, i)
        vfgData.Cell(flexcpData, lngUpRow, i) = vfgData.Cell(flexcpData, lngRow, i)
        vfgData.Cell(flexcpPicture, lngUpRow, i) = vfgData.Cell(flexcpPicture, lngRow, i)
        
        vfgData.TextMatrix(lngRow, i) = strRowText
        vfgData.Cell(flexcpData, lngRow, i) = strRowData
        vfgData.Cell(flexcpPicture, lngRow, i) = varRowPic
    Next i
    
    Call UpdateRowNumber
    
    Call vfgData.Select(lngUpRow, 0)
End Sub

Public Sub MoveDown(ByVal lngRow As Long)
'下移一行
    Dim strRowData As Variant
    Dim strRowText As Variant
    Dim varRowPic  As Variant
    
    Dim i As Long
    Dim lngDownRow As Long
    
    If lngRow >= vfgData.Rows - 1 Then Exit Sub
    
    lngDownRow = lngRow + 1
    
    Do While lngDownRow < vfgData.Rows - 1
        If vfgData.RowHidden(lngDownRow) Then
            lngDownRow = lngDownRow + 1
        Else
            Exit Do
        End If
    Loop

    If vfgData.RowHidden(lngDownRow) Then Exit Sub
    
    For i = 0 To vfgData.Cols - 1
        
        strRowText = vfgData.TextMatrix(lngDownRow, i)
        strRowData = vfgData.Cell(flexcpData, lngDownRow, i)
        Set varRowPic = vfgData.Cell(flexcpPicture, lngDownRow, i)
        
        vfgData.TextMatrix(lngDownRow, i) = vfgData.TextMatrix(lngRow, i)
        vfgData.Cell(flexcpData, lngDownRow, i) = vfgData.Cell(flexcpData, lngRow, i)
        vfgData.Cell(flexcpPicture, lngDownRow, i) = vfgData.Cell(flexcpPicture, lngRow, i)
        
        vfgData.TextMatrix(lngRow, i) = strRowText
        vfgData.Cell(flexcpData, lngRow, i) = strRowData
        vfgData.Cell(flexcpPicture, lngRow, i) = varRowPic
    Next i
    
    Call UpdateRowNumber
    
    Call vfgData.Select(lngDownRow, 0)
End Sub


Public Sub ShowCheckRows()
'只显示勾选行
    Dim i As Long
    
    For i = 1 To vfgData.Rows - 1
        If Not GetRowCheck(i) Then
            vfgData.RowHidden(i) = True
        Else
            vfgData.RowHidden(i) = False
        End If
    Next i
    
    Call UpdateRowNumber
End Sub



Public Sub ShowAllRows()
'只显示勾选行
    Dim i As Long
    
    For i = 1 To vfgData.Rows - 1
        vfgData.RowHidden(i) = False
    Next i
    
    Call UpdateRowNumber
End Sub


Public Sub DelCurRow(Optional ByVal blnKeepRowCount As Boolean = True)
'删除当前行
    
    Call DelRow(mDataGrid.RowSel, blnKeepRowCount)
    
    Call UpdateRowNumber
    Call RefreshReadColColor
End Sub


Public Sub DelRow(ByVal lngRow As Long, Optional ByVal blnKeepRowCount As Boolean = True, Optional ByVal blnUpdateAdo As Boolean = False)
'删除指定行数据
    Dim iNextIndex As Long
    
    mDataGrid.RowHidden(lngRow) = True
    iNextIndex = GetNextRowIndex(lngRow)
    
    '重新设置行状态
    If mDataGrid.RowData(lngRow) = TDataRowState.Add Then
        mDataGrid.RowData(lngRow) = TDataRowState.Normal
    Else
        If Not IsEmptyKey(lngRow) Then
            mDataGrid.RowData(lngRow) = TDataRowState.Del
        End If
    End If
    
    If blnKeepRowCount Then mDataGrid.Rows = mDataGrid.Rows + 1
    
    If iNextIndex > 0 Then Call LocateRow(iNextIndex)
    
    
    '更新ado中的数据
    If blnUpdateAdo Then
        mrsData.Filter = mstrKeyField & "=" & KeyValue(lngRow)
        
        If mrsData.RecordCount > 0 Then
            Call mrsData.Delete
        End If
    End If
    
    
    Call UpdateRowNumber
    Call RefreshReadColColor
End Sub


Public Function GetNextRowIndex(ByVal lngRow As Long) As Long
'取得下一行的索引
    Dim i As Long
    
    GetNextRowIndex = -1
    
    For i = lngRow + 1 To mDataGrid.Rows - 1
        If Not mDataGrid.RowHidden(i) Then
            GetNextRowIndex = i
            Exit Function
        End If
    Next i
    
    If GetNextRowIndex = -1 Then
        i = lngRow - 1
        Do While i > 0
            If Not mDataGrid.RowHidden(i) Then
                GetNextRowIndex = i
                Exit Function
            End If
            
            i = i - 1
        Loop
    End If
End Function


Public Sub EditNextCell(ByVal lngRow As Long, Optional ByVal blnAutoNextRow As Boolean = True)
'编辑下一列
    If mDataGrid.Editable = flexEDNone Then Exit Sub
    
    Do While mDataGrid.ColSel + 1 < mDataGrid.Cols
        If Not (IsReadCol(mDataGrid.ColSel + 1) Or mDataGrid.ColHidden(mDataGrid.ColSel + 1)) Then
            Exit Do
        Else
            Call mDataGrid.Select(lngRow, mDataGrid.ColSel + 1)
        End If
    Loop
    
nextCell:
    
    If mDataGrid.ColSel + 1 >= mDataGrid.Cols Then
        If blnAutoNextRow Then
            Dim iRow As Long
            Dim iCol As Long
            
            iRow = GetNextRowIndex(lngRow)
            
            If iRow > 0 Then
                For iCol = IIf(mblnIsAllowExtCol, 1, 0) To mDataGrid.Cols - 1
                    If Not IsReadCol(iCol) Then
                        If Not mDataGrid.ColHidden(iCol) Then Exit For
                    End If
                Next iCol
                
                If iRow < mDataGrid.Rows Then
                    If iCol = mDataGrid.Cols Then iCol = mDataGrid.Cols - 1
                    
                    Call mDataGrid.Select(iRow, iCol)
                    Call mDataGrid.ShowCell(iRow, iCol)
                End If
            End If
            
            If Not IsCheckboxCol(iCol) Then Call mDataGrid.EditCell
        End If
        
        Exit Sub
    End If
    
    
    
    Call mDataGrid.Select(lngRow, mDataGrid.ColSel + 1)
        
    If Not IsCheckboxCol(iCol) Then Call mDataGrid.EditCell
End Sub


Public Function IsEmptyKey(ByVal lngRow As Long) As Boolean
'检查key是否为空
    IsEmptyKey = True
    
    If Trim(mDataGrid.TextMatrix(lngRow, GetColIndex(mstrKeyName))) <> "" Then
        IsEmptyKey = False
    End If
End Function

Public Function IsEmptyKeyWithCur() As Boolean
'检查当前Key是否为空
    IsEmptyKeyWithCur = IsEmptyKey(mDataGrid.RowSel)
End Function


Public Sub EditNextCellWithCurRow(Optional ByVal blnAutoNextRow As Boolean = True)
'编辑当前行的下一列
    Call EditNextCell(mDataGrid.RowSel, blnAutoNextRow)
End Sub


Public Sub RemoveRow(ByVal lngRow As Long)
'删除行
    Call mDataGrid.RemoveItem(lngRow)
End Sub


Public Function GetColIndex(ByVal strColName As String) As Long
'获取列索引
    GetColIndex = mDataGrid.ColIndex(strColName)
End Function


Public Function GetColIndexWithRowCheck() As Long
'取得行选的check
    Dim i As Long
    
    GetColIndexWithRowCheck = -1
    
    For i = 0 To mDataGrid.Cols - 1
    
        If RefreshColDicObject(i) Then
            If mobjTmpDictionary(TColPro.cpIsRowCheck) And Not mDataGrid.ColHidden(i) Then
                GetColIndexWithRowCheck = i
                Exit Function
            End If
        End If
    Next i

End Function


Public Function GetColIndexWithBtn() As Long
'取得行选的check
    Dim i As Long
    
    GetColIndexWithBtn = -1

    For i = 0 To mDataGrid.Cols - 1
        
        If RefreshColDicObject(i) Then
            If mobjTmpDictionary(TColPro.cpIsBtn) And Not mDataGrid.ColHidden(i) Then
                GetColIndexWithBtn = i
                Exit Function
            End If
        End If
    Next i

End Function


Public Function GetColName(ByVal lngColIndex As Long) As String
'获取列名称
    GetColName = mDataGrid.ColKey(lngColIndex)
End Function


Public Function ColumnEnableWithColName(ByVal strColName As String) As Boolean
'读取列可编辑状态
    Dim lngColIndex As Long
    
    ColumnEnableWithColName = False
    
    lngColIndex = GetColIndex(strColName)
    
    If Not RefreshColDicObject(lngColIndex) Then Exit Function
    
    '如果是隐藏列，则不允许编辑
    If mobjTmpDictionary(TColPro.cpIsHide) Then
        ColumnEnableWithColName = False
        Exit Function
    End If
    
     '如果是CheckBox列，则不可编辑 注：此处只是将列状态判断为不可编辑  但CheckBox依然可以勾选
     '修改:罗冠骁
    If mobjTmpDictionary(TColPro.cpIsRowCheck) Then
        ColumnEnableWithColName = False
        Exit Function
    End If
    
    ColumnEnableWithColName = Not mobjTmpDictionary(TColPro.cpIsRead)
End Function

Public Function ColumnEnable(ByVal lngCol As Long) As Boolean
'读取列可编辑状态
    ColumnEnable = ColumnEnableWithColName(GetColName(lngCol))
End Function


Public Sub ClearListData()
'清除数据列表
    mDataGrid.Rows = 1
    If mlngKeepRows > 0 Then mDataGrid.Rows = mlngKeepRows
    
    Call UpdateRowNumber
End Sub


Public Function GetRowCheck(ByVal lngRow As Long) As Boolean
'读取check行的选择状态
    GetRowCheck = GetCellCheckState(lngRow, GetColIndexWithRowCheck())
End Function


Public Sub SetRowCheck(ByVal lngRow As Long, ByVal blnIsChecked As Boolean)
'设置check行的选择状态
    Dim blnAllowChange As Boolean
    Dim lngCheckIndex As Long
    
    blnAllowChange = True
    
    lngCheckIndex = GetColIndexWithRowCheck()
    
    RaiseEvent OnCheckChanging(lngRow, lngCheckIndex, blnAllowChange)
    
    If blnAllowChange Then
        Call SetCellCheckState(lngRow, GetColIndexWithRowCheck(), blnIsChecked)
    
        RaiseEvent OnCheckChanged(lngRow, lngCheckIndex)
    End If
End Sub

Public Sub SetColBackColor(ByVal strColName As String, ByVal Col As OLE_COLOR)
'设置列的背景颜色
    mDataGrid.Cell(flexcpBackColor, 1, GetColIndex(strColName), mDataGrid.Rows - 1, GetColIndex(strColName)) = Col
End Sub


Public Sub SetCurColBackColor(ByVal Col As OLE_COLOR)
'设置当前列的背景颜色
    mDataGrid.Cell(flexcpBackColor, 1, mDataGrid.ColSel, mDataGrid.Rows - 1, mDataGrid.ColSel) = Col
End Sub


Public Sub ShowHideData()
'显示所有隐藏数据行
    Dim i As Long

    For i = 1 To mDataGrid.Rows - 1
        mDataGrid.RowHidden(i) = False
    Next i
End Sub

Public Sub ShowCol(ByVal lngColIndex As Long)
'显示列
    If lngColIndex >= 0 And lngColIndex < mDataGrid.Cols - 1 Then
        mDataGrid.ColHidden(lngColIndex) = True
        Call ShowCellButton
    End If
End Sub


Public Sub HidenCol(ByVal lngColIndex As Long)
'隐藏列
    If lngColIndex >= 0 And lngColIndex < mDataGrid.Cols - 1 Then
        mDataGrid.ColHidden(lngColIndex) = True
        Call ShowCellButton
    End If
End Sub


Public Sub CheckAllCell(ByVal lngChkCol As Long)
'选择指定列的所有checkbox
    Dim i As Long
    Dim blnAllowCheck As Boolean
    
    If Not IsCheckboxCol(lngChkCol) Then Exit Sub
    
    blnAllowCheck = True
    RaiseEvent OnCheckAllChanging(blnAllowCheck)
    
    If Not blnAllowCheck Then Exit Sub
    
    For i = 1 To vfgData.Rows - 1
        If Not vfgData.Cell(flexcpPicture, i, lngChkCol) Is Nothing Then
    
            '如果check处于禁用状态，则不允许编辑
            If vfgData.Cell(flexcpPicture, i, lngChkCol).Tag <> csDisCheck Then
                Set vfgData.Cell(flexcpPicture, i, lngChkCol) = imgCheck(1)
            End If
        End If
    Next i
    
    RaiseEvent OnCheckAllChanged
End Sub


Public Sub ClearCellCheck(ByVal lngChkCol As Long)
'清除选择的所有checkbox
    Dim i As Long
    Dim blnAllowCheck As Boolean
    
    If Not IsCheckboxCol(lngChkCol) Then Exit Sub
    
    blnAllowCheck = True
    RaiseEvent OnCheckAllChanging(blnAllowCheck)
    
    If Not blnAllowCheck Then Exit Sub
    
    For i = 1 To vfgData.Rows - 1
        If Not vfgData.Cell(flexcpPicture, i, lngChkCol) Is Nothing Then
    
            '如果check处于禁用状态，则不允许编辑
            If vfgData.Cell(flexcpPicture, i, lngChkCol).Tag <> csDisCheck Then
                Set vfgData.Cell(flexcpPicture, i, lngChkCol) = imgCheck(0)
            End If
        End If
    Next i
    
    RaiseEvent OnCheckAllChanged
End Sub


Private Function IsMergeCol(ByVal lngCol As Long) As Boolean
'判断是否合并列
    
    IsMergeCol = False
    If lngCol < 0 Or lngCol >= mDataGrid.Cols Then Exit Function
    
    If Not RefreshColDicObject(lngCol) Then Exit Function
    
    IsMergeCol = mobjTmpDictionary(TColPro.cpIsMerage)
End Function


Private Function IsUnResizeCol(ByVal lngCol As Long) As Boolean
'判断是否UnResizeCol列
    
    IsUnResizeCol = False
    If lngCol < 0 Or lngCol >= mDataGrid.Cols Then Exit Function

    If Not RefreshColDicObject(lngCol) Then Exit Function
    
    IsUnResizeCol = mobjTmpDictionary(TColPro.cpIsUnResize)
End Function


Private Function IsReadCol(ByVal lngCol As Long) As Boolean
'判断指定列是否为Read列

    IsReadCol = True
    If lngCol < 0 Then Exit Function
    
'    '如果是隐藏列，则read为true
'    If mDataGrid.ColHidden(lngCol) Then
'        IsReadCol = True
'        Exit Function
'    End If
    
    If Not RefreshColDicObject(lngCol) Then Exit Function
    
    IsReadCol = mobjTmpDictionary(TColPro.cpIsRead)
End Function


Public Sub RefreshReadColColor()
'刷新不能编辑列的颜色
On Error Resume Next

    Dim i As Long
    Dim j As Long
    Dim strColProperty As String

    If mDataGrid.Editable = flexEDNone Then Exit Sub
    
    For i = 0 To mDataGrid.Cols - 1
        If IsReadCol(i) Then
            mDataGrid.Cell(flexcpBackColor, 1, i, mDataGrid.Rows - 1, i) = mlngDisableColor
        End If
    Next i
End Sub


Private Function IsButtonCol(ByVal lngCol As Long) As Boolean
'判断指定列是否为Button列
    
    IsButtonCol = False
    If lngCol < 0 Then Exit Function
    
    '如果是隐藏列，则read为true
    If mDataGrid.ColHidden(lngCol) Then
        IsButtonCol = False
        Exit Function
    End If
    
    If Not RefreshColDicObject(lngCol) Then Exit Function
    
    IsButtonCol = mobjTmpDictionary(TColPro.cpIsBtn)
End Function


Public Function IsComboboxCol(ByVal lngCol As Long) As Boolean
'判断指定列是否为Combobox列

    IsComboboxCol = False
    If lngCol < 0 Then Exit Function
    
    '如果是隐藏列，则read为true
    If mDataGrid.ColHidden(lngCol) Then
        IsComboboxCol = False
        Exit Function
    End If
    
    If Not RefreshColDicObject(lngCol) Then Exit Function
    
    IsComboboxCol = mobjTmpDictionary(TColPro.cpIsCombox)
End Function


Private Function IsReadColWithName(ByVal strColName As String) As Boolean
'判断指定列是否为check列
    IsReadColWithName = IsReadCol(GetColIndex(strColName))
End Function


Private Function IsCheckboxCol(ByVal lngCol As Long) As Boolean
'判断指定列是否为check列
    
    IsCheckboxCol = False
    
    If lngCol < 0 Then Exit Function
    
    '如果是隐藏列，则check为false
    If mDataGrid.ColHidden(lngCol) Then
        IsCheckboxCol = False
        Exit Function
    End If
    
    If Not RefreshColDicObject(lngCol) Then Exit Function
        
    IsCheckboxCol = IIf(mobjTmpDictionary(TColPro.cpIsCheck) Or mobjTmpDictionary(TColPro.cpIsRowCheck), True, False)
End Function



Private Function IsCheckboxColWithName(ByVal strColName As String) As Boolean
'判断指定列是否为check列
    IsCheckboxColWithName = IsCheckboxCol(GetColIndex(strColName))
End Function


Public Function IsDateCol(ByVal lngCol As Long) As Boolean
'判断指定列是否为日期列

    IsDateCol = False
    If lngCol < 0 Then Exit Function
    
    If Not RefreshColDicObject(lngCol) Then Exit Function

    IsDateCol = mobjTmpDictionary(TColPro.cpIsDateCol)
End Function



Public Function IsCheckedRow() As Boolean
'判断是否有勾选的行数据
    Dim i As Long
    Dim lngCustomCheckIndex As Long
    
    IsCheckedRow = False
    
    lngCustomCheckIndex = GetColIndexWithRowCheck()
    
    For i = 1 To vfgData.Rows - 1
        IsCheckedRow = GetCellCheckState(i, lngCustomCheckIndex)
        
        If IsCheckedRow Then Exit Function
    Next i
End Function

Public Function GetHideRowCount() As Long
'取得行选的隐藏列数量
    Dim i As Long
    Dim lngHideRowCount As Long

    lngHideRowCount = 0

    For i = 0 To vfgData.Cols - 1
        If RefreshColDicObject(i) Then
            If mobjTmpDictionary(TColPro.cpIsHide) Then
                lngHideRowCount = lngHideRowCount + 1
            End If
            
            If mobjTmpDictionary(TColPro.cpIsRowCheck) Then
                GetHideRowCount = lngHideRowCount
                Exit Function
            End If
        End If
    Next i

End Function


Public Function IsNullRow(ByVal lngRow As Long) As Boolean
'判断该行是否为空行
    Dim i As Long
    
    IsNullRow = True
    
    If mblnIsAllowExtCol Then
        IsNullRow = IIf(Len(mDataGrid.Cell(flexcpText, lngRow, 1, lngRow, mDataGrid.Cols - 1)) = mDataGrid.Cols - 2, True, False)
    Else
        IsNullRow = IIf(Len(mDataGrid.Cell(flexcpText, lngRow, 0, lngRow, mDataGrid.Cols - 1)) = mDataGrid.Cols - 1, True, False)
    End If
End Function


Public Function IsNullWithCurRow() As Boolean
    IsNullWithCurRow = IsNullRow(mDataGrid.RowSel)
End Function


Public Function GetNullRowIndex() As Long
'返回新的空数据行索引
    Dim i As Long
    
    GetNullRowIndex = -1
    
    For i = 1 To mDataGrid.Rows - 1
        If Not mDataGrid.RowHidden(i) Then
            If IsNullRow(i) Then
                GetNullRowIndex = i
                Exit Function
            End If
        End If
    Next i
End Function


Public Function NewRow() As Long
'新增数据行
    Dim lngNew As Long
    Dim lngFontHeight As Long
    
    NewRow = -1
    
    mDataGrid.Rows = mDataGrid.Rows + 1
    
    mDataGrid.RowSel = mDataGrid.Rows - 1
    Call mDataGrid.ShowCell(mDataGrid.RowSel, 0)
    
    '添加行状态
    mDataGrid.RowData(mDataGrid.RowSel) = TDataRowState.Normal
    
    Call UpdateRowNumber
    Call RefreshReadColColor
    Call RefreshAlign(lngNew)
    
    Call LocateRow(mDataGrid.RowSel)
    
    NewRow = mDataGrid.RowSel
End Function


Public Function GetCellCheckState(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
'取得单元格的check状态
    GetCellCheckState = False
    
    If Not IsCheckboxCol(lngCol) Then Exit Function
    If vfgData.Cell(flexcpPicture, lngRow, lngCol) Is Nothing Then Exit Function
    
    If vfgData.Cell(flexcpPicture, lngRow, lngCol).Tag = csDisCheck Then
        GetCellCheckState = False
        Exit Function
    End If
    
    GetCellCheckState = IIf(Val(vfgData.Cell(flexcpPicture, lngRow, lngCol).Tag) = 0, False, True)
End Function


Public Sub SetCellCheckState(ByVal lngRow As Long, ByVal lngCol As Long, ByVal blnChk As Boolean)
'设置单元格的check状态
    If Not IsCheckboxCol(lngCol) Then Exit Sub

    If blnChk Then
        vfgData.Cell(flexcpPicture, lngRow, lngCol) = imgCheck(1)
    Else
        vfgData.Cell(flexcpPicture, lngRow, lngCol) = imgCheck(0)
    End If
End Sub


Public Sub DisableCheck(ByVal lngRow As Long, ByVal lngCol As Long)
    If Not IsCheckboxCol(lngCol) Then Exit Sub
    
    vfgData.Cell(flexcpPicture, lngRow, lngCol) = imgCheck(2)
End Sub


Public Sub EnableCheck(ByVal lngRow As Long, ByVal lngCol As Long)
    If Not IsCheckboxCol(lngCol) Then Exit Sub
    
    vfgData.Cell(flexcpPicture, lngRow, lngCol) = imgCheck(0)
End Sub




Public Function ReCellCheckState(ByVal lngRow As Long, ByVal lngCol As Long)
'反向设置单元格的check状态

    If Not IsCheckboxCol(lngCol) Then Exit Function
    If vfgData.Cell(flexcpPicture, lngRow, lngCol) Is Nothing Then Exit Function
    
    If vfgData.Cell(flexcpPicture, lngRow, lngCol).Tag = 0 Then
        vfgData.Cell(flexcpPicture, lngRow, lngCol) = imgCheck(1)
    Else
        vfgData.Cell(flexcpPicture, lngRow, lngCol) = imgCheck(0)
    End If
End Function


Public Function FormatValue(ByVal lngCol As Long, ByVal strValue As String) As String
    
    FormatValue = strValue
    
    If Not RefreshColDicObject(lngCol) Then Exit Function
    
    '格式化时间样式
    If IsDate(strValue) Then
        If mobjTmpDictionary(TColPro.cpDataType) = M_STR_ColProp_TFullDateTime Then
            strValue = Format(strValue, "yyyy-mm-dd hh:mm:ss")
        ElseIf mobjTmpDictionary(TColPro.cpDataType) = M_STR_ColProp_TOnlyDate Then
            strValue = Format(strValue, "yyyy-mm-dd")
        ElseIf mobjTmpDictionary(TColPro.cpDataType) = M_STR_ColProp_TOnlyTime Then
            strValue = Format(strValue, "hh:mm:ss")
        ElseIf mobjTmpDictionary(TColPro.cpDataType) = M_STR_ColProp_TShortDateTime Then
            strValue = Format(strValue, "yyyy-mm-dd hh:mm")
        End If
    End If
    
    '补齐字符长度
    If Len(strValue) < mobjTmpDictionary(TColPro.cpAlignLen) Then
        strValue = Lpad(strValue, mobjTmpDictionary(TColPro.cpAlignLen), mobjTmpDictionary(TColPro.cpAlignChar))
    End If
    
    FormatValue = strValue
End Function


Private Sub WriteText(ByVal lngRow As Long, ByVal lngCol As Long, ByVal strText As String, Optional ByVal strTag As String = "")
'根据行列信息设置当前值
    Dim iCol As Long
    Dim strValue As String

    If lngCol <= -1 Then Exit Sub
    
    strValue = FormatValue(lngCol, strText)
    
    '设置数据值
    mDataGrid.Cell(flexcpText, lngRow, lngCol) = strValue
    '对行单元格的data赋值，可用于以后直接判断该单元格是否进行更新
    mDataGrid.Cell(flexcpData, lngRow, lngCol) = IIf(strTag = "", strValue, strTag)
    
    If mblnIsAutoRowHeight Then
        '调整数据行高度
        If Not mDataGrid.ColHidden(lngCol) Then Call RefreshRowHeight(lngRow, strText)
    End If
End Sub


Public Sub RefreshRowHeight(ByVal lngRow As Long, ByVal strText As String)
'自动根据行中的数据值刷新行的高度
    Dim lngCharCount As Long
    Dim lngCharHeight As Long
    
    lngCharCount = GetCharCount(strText, vbCr)
    
    If lngCharCount = 0 Then Exit Sub
    
    If mlngOldDataRowHeight <= 0 Then mlngOldDataRowHeight = mDataGrid.RowHeight(lngRow)
    
    lngCharHeight = mlngOldDataRowHeight * (lngCharCount + 1)
    
    If lngCharHeight > mDataGrid.RowHeight(lngRow) Then
        mDataGrid.RowHeight(lngRow) = lngCharHeight
    End If
End Sub


Private Function GetCharCount(ByVal strSource As String, ByVal strChar As String) As Long
'获取相同字符数量
    GetCharCount = Len(strSource) - Len(Replace(strSource, strChar, ""))
End Function


Public Sub WriteCurText(ByVal lngCol As Long, ByVal strText As String)
'根据当前选择行，设置指定列值
    Call WriteText(mDataGrid.RowSel, lngCol, strText)
End Sub


Public Sub SyncData(ByVal lngRow As Long, ByVal strColName As String, _
    ByVal strData As String, Optional ByVal blnUpdateAdo As Boolean = False)
'设置数据
    Dim blnFind As Boolean
    Dim chkState As TDataRowState
    Dim strText As String
    Dim lngImgIndex As Long
    Dim lngColIndex As Long
    
    lngColIndex = GetColIndex(strColName)

    '更新界面数据显示
    Call GetFieldDisplayText(strColName, strData, blnFind, chkState, strText, lngImgIndex)
    Call UpdateCellStyle(lngRow, lngColIndex, lngImgIndex, chkState)

    Call WriteText(lngRow, lngColIndex, strText)
    
    
    '更新ado中的数据
    If blnUpdateAdo Then
        If mrsData Is Nothing Then Exit Sub
        
        mrsData.Filter = mstrKeyField & "=" & KeyValue(lngRow)
        
        
        If mrsData.RecordCount > 0 Then
            mrsData.MoveFirst
    
            mrsData(GetFieldName(lngColIndex)) = strData
        Else
            '如果不存在或者没有找到关键数据，则在数据集中新增
            If strColName = mstrKeyField Then
                Call mrsData.AddNew
                mrsData(GetFieldName(lngColIndex)).value = strData
            End If
        End If
    End If
End Sub


Public Sub SyncText(ByVal lngRow As Long, ByVal strColName As String, _
    ByVal strText As String, Optional ByVal blnUpdateAdo As Boolean = False)
'根据界面中的显示文本设置数据
    Dim blnFind As Boolean
    Dim chkState As TDataRowState
    Dim strData As String
    Dim lngImgIndex As Long
    Dim lngColIndex As Long
    
    
    lngColIndex = GetColIndex(strColName)
    
    strData = GetFieldDataValue(strColName, strText, blnFind)
    strData = IIf(strData = "", strText, strData)

    
    '更新界面数据显示
    Call GetFieldDisplayText(strColName, strData, blnFind, chkState, strText, lngImgIndex)
    Call UpdateCellStyle(lngRow, lngColIndex, lngImgIndex, chkState)

    Call WriteText(lngRow, lngColIndex, strText)
    
    
    '更新ado中的数据
    If blnUpdateAdo Then
        mrsData.Filter = mstrKeyField & "=" & KeyValue(lngRow)
        
        
        If mrsData.RecordCount > 0 Then
            mrsData.MoveFirst
    
            mrsData(GetFieldName(lngColIndex)) = strData
        Else
            '如果不存在或者没有找到关键数据，则在数据集中新增
            If strColName = mstrKeyField Then
                Call mrsData.AddNew
                mrsData(GetFieldName(lngColIndex)).value = strData
            End If
        End If
    End If
End Sub

Public Function UpdateSourceData(ByVal strKeyValue As String, ByVal strField As String, ByVal strNewValue As Variant) As Boolean
'更源数据值
    Dim strFilter As String
    
    UpdateSourceData = False
    If strKeyValue = "" Then Exit Function
    
    strFilter = mstrKeyField & "='" & strKeyValue & "'"
    
    If mrsData.Filter <> strField Then mrsData.Filter = strFilter
    
    If mrsData.RecordCount <= 0 Then Exit Function
    
    mrsData.MoveFirst
    
    mrsData(strField) = strNewValue
    
    UpdateSourceData = True
End Function


Public Function GetColNameWithDataField(ByVal strDataField As String) As String
'根据数据字段获取列显示名称
    Dim strTemp As String
    Dim lngFindIndex As Long
    
    GetColNameWithDataField = strDataField
     
    lngFindIndex = InStr(UCase(mstrColNames), UCase(">" & strDataField))
    If lngFindIndex > 0 Then
        strTemp = Mid(mstrColNames, 1, lngFindIndex - 1)
        strTemp = Mid(mstrColNames, InStrRev(mstrColNames, "|") + 1, 100)
        
        GetColNameWithDataField = strTemp
    End If
End Function


Public Sub SyncRowDataToAdo(ByVal lngRowIndex As Long)
'同步ADO中的行数据
    Dim i As Long
    Dim strColName As String
    Dim strText As String
    Dim strCode As String
    Dim blnFind As Boolean
    
    Select Case RowState(lngRowIndex)
        Case TDataRowState.Add
            Call mrsData.AddNew
            
            blnFind = False
            
            '同步添加数据
            For i = 0 To mrsData.Fields.Count - 1
                strColName = GetColNameWithDataField(mrsData.Fields(i).Name)
                strText = Text(lngRowIndex, strColName)
                
                strCode = GetFieldDataValue(strColName, strText, blnFind)
                
                mrsData.Fields(i) = IIf(blnFind, strCode, strText)
                
                mDataGrid.Cell(flexcpData, lngRowIndex, GetColIndex(strColName)) = IIf(blnFind, strCode, strText)
            Next i
            
            
        Case TDataRowState.Del
            '同步删除数据
            mrsData.Filter = mstrKeyField & "=" & KeyValue(lngRowIndex)
            
            If mrsData.RecordCount > 0 Then mrsData.Delete
        Case TDataRowState.Update
            '同步更新数据
            mrsData.Filter = mstrKeyField & "=" & KeyValue(lngRowIndex)
            
            mrsData.MoveFirst
            If mrsData.RecordCount > 0 Then
                blnFind = False
                
                For i = 0 To mrsData.Fields.Count - 1
                    strColName = GetColNameWithDataField(mrsData.Fields(i).Name)
                    strText = Text(lngRowIndex, strColName)
                    
                    strCode = GetFieldDataValue(strColName, strText, blnFind)
                    
                    mrsData.Fields(i) = IIf(blnFind, strCode, strText)
                    
                    mDataGrid.Cell(flexcpData, lngRowIndex, GetColIndex(strColName)) = IIf(blnFind, strCode, strText)
                Next i
                
            End If
    End Select
    
    mrsData.Filter = ""
End Sub


Public Sub SetRowColor(ByVal lngRow As Long, ByVal lngColor As OLE_COLOR)
'设置行背景色
    If mblnIsAllowExtCol Then
        mDataGrid.Cell(flexcpBackColor, lngRow, 1, lngRow, mDataGrid.Cols - 1) = lngColor
    Else
        mDataGrid.Cell(flexcpBackColor, lngRow, 0, lngRow, mDataGrid.Cols - 1) = lngColor
    End If

    Call RefreshReadColColor
End Sub


Public Sub ShowHintInf(ByVal strHint As String)
'显示提示信息
    labInf.Caption = strHint
    picShowHint.Visible = True
End Sub

Public Sub CloseHintInf()
'关闭提示信息
    picShowHint.Visible = False
End Sub


Public Sub Sort(ByVal lngCol As Long)
'对指定列进行排序
    mDataGrid.Col = lngCol
    mDataGrid.Sort = 1
    
    Call UpdateRowNumber
End Sub


Public Function FindRowIndex(ByVal strFindValue As String, ByVal strColName As String, _
    Optional ByVal blnIsPrecise As Boolean = False) As Long
'查找指定值并返回所在行索引
    Dim i As Long
    Dim lngCol As Long

    FindRowIndex = -1
    If Trim(strFindValue) = "" Then Exit Function

    lngCol = GetColIndex(strColName)

    For i = 1 To mDataGrid.Rows - 1
        If Not mDataGrid.RowHidden(i) Then

            If UCase(mDataGrid.TextMatrix(i, lngCol)) Like IIf(blnIsPrecise, UCase(strFindValue), "*" & UCase(strFindValue) & "*") Then
                FindRowIndex = i
                Exit Function
            End If
        End If
    Next i
End Function


Public Function IsUpdate(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
'检查数据是否更新
    Dim blnUpdate As Boolean
    
    blnUpdate = False
    
    If mDataGrid.TextMatrix(lngRow, lngCol) <> mDataGrid.Cell(flexcpData, lngRow, lngCol) Then
        blnUpdate = True
    End If
    
    IsUpdate = blnUpdate
End Function


Public Function IsUpdateWithCurRow(ByVal lngCol As Long) As Boolean
'检查当前行数据是否有更新
    IsUpdateWithCurRow = False
    
    Call IsUpdate(mDataGrid.RowSel, lngCol)
End Function



Public Function GetFieldName(ByVal lngCol As Long) As String
'取得数据库字段名称

    GetFieldName = ""
    
    If Not RefreshColDicObject(lngCol) Then Exit Function
    
    GetFieldName = mobjTmpDictionary(TColPro.cpFieldName)
End Function

Public Function GetOrder(ByVal lngCurSortCol As Long, ByVal lngCurOrder As Long)
'取得排序order（参考vsflexgrid的排序demo）
    GetOrder = lngCurOrder
'
'    ' no flags? apply custom sort
'    If mCurFlexGrid.ExplorerBar > &H1000& Then Exit Function
'
'    '
'    ' the 'ignore blanks' flag isn't set, so do it with custom code
'    '
'
'    ' save selection
    
     '没有数据时退出排序
    If mDataGrid.Rows = 1 Then Exit Function
    
    Dim R&, c&, RS&, cs&
    mDataGrid.GetSelection R, c, RS, cs
    mDataGrid.Redraw = flexRDNone

    ' apply sort to non-empty range
    Dim Row%
    For Row = mDataGrid.Rows - 1 To mDataGrid.FixedRows Step -1
        '整行数据为空时，不参与排序
        If Len(mDataGrid.TextMatrix(Row, lngCurSortCol)) Or Not IsEmptyKey(Row) Then Exit For
    Next
    
    If Row > mDataGrid.FixedRows Then
        mDataGrid.Select mDataGrid.FixedRows, lngCurSortCol, Row, lngCurSortCol
        mDataGrid.Sort = lngCurOrder
    End If
    
    ' restore selection
    mDataGrid.Select R, c, RS, cs
    mDataGrid.Redraw = flexRDDirect
    
    ' cancel default sort
    GetOrder = 0
End Function



Private Sub vfgData_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next

    Dim blnFind As Boolean
    Dim chkState As CheckState
    Dim strConvertValue As String
    Dim lngImgIndex As Long
    
    '更新行状态
    If IsEmptyKey(Row) Then
        RowState(Row) = TDataRowState.Add
    Else
        If IsUpdate(Row, Col) Then RowState(Row) = TDataRowState.Update
    End If
    
    Call GetFieldDisplayText(GetColName(Col), vfgData.Cell(flexcpText, Row, Col), blnFind, chkState, strConvertValue, lngImgIndex)
    
    Call UpdateCellStyle(Row, Col, lngImgIndex, chkState)

    RaiseEvent OnAfterEdit(Row, Col)
    
    err.Clear
End Sub
 

Private Sub vfgData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
    RaiseEvent OnBeforeEdit(Row, Col, Cancel)
    
    err.Clear
End Sub

Private Sub vfgData_BeforeMoveColumn(ByVal Col As Long, Position As Long)
'合并列不允许移动
On Error Resume Next
    If vfgData.MergeCells <> flexMergeRestrictAll Then Exit Sub
    
    If IsMergeCol(Col) Or IsMergeCol(Position) Then Position = -1
    
'    Call ShowCellButton
    
    err.Clear
End Sub

Private Sub vfgData_AfterSort(ByVal Col As Long, Order As Integer)
'排序之后更新行号
On Error Resume Next
    Dim blnCustom As Boolean
    
    blnCustom = False
    mlngSortCol = Col
    mlngSortWay = Order
    
    RaiseEvent OnOrderChange(Col, Order, blnCustom)
'    Call UpdateRowNumber

    err.Clear
End Sub


Private Sub vfgData_BeforeSort(ByVal Col As Long, Order As Integer)
'空数据行不参与排序
On Error Resume Next
    Dim blnCustom As Boolean
    
    blnCustom = False
    mlngSortCol = Col
    mlngSortWay = Order
    
    RaiseEvent OnOrderChange(Col, Order, blnCustom)
    
    'blnCustom返回true表示使用自定义排序规则
    If Not blnCustom Then
        'order 如果返回0表示不参与排序
        Order = GetOrder(Col, Order)
    End If
    
    Call UpdateRowNumber
    
    err.Clear
End Sub

Private Sub vfgData_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next
    RaiseEvent OnCellButtonClick(Row, Col)
    
    If mblnIsBtnNextCell Then Call EditNextCell(Row)
    
    err.Clear
End Sub

Private Sub vfgData_CellChanged(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next
    RaiseEvent OnCellChanged(Row, Col)
    
    err.Clear
End Sub

Private Sub vfgData_ChangeEdit()
On Error Resume Next
    RaiseEvent OnChangeEdit
    err.Clear
End Sub

Private Sub vfgData_Click()
On Error Resume Next
    RaiseEvent OnClick
    
    err.Clear
End Sub

Private Sub vfgData_EnterCell()
On Error Resume Next
    RaiseEvent OnEnterCell
    
    err.Clear
End Sub

Private Sub vfgData_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    RaiseEvent OnKeyDown(KeyCode, Shift)
    
    err.Clear
End Sub

Private Sub vfgData_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
On Error Resume Next
    RaiseEvent OnKeyDownEdit(Row, Col, KeyCode, Shift)
    
    err.Clear
End Sub

Private Sub vfgData_KeyPress(KeyAscii As Integer)
On Error Resume Next
    RaiseEvent OnKeyPress(KeyAscii)
    
    err.Clear
End Sub



Private Sub vfgData_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
On Error Resume Next
    RaiseEvent OnKeyPressEdit(Row, Col, KeyAscii)
    
    err.Clear
End Sub

Private Sub vfgData_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
'进入下一编辑单元格
    Dim blnIsDel As Boolean
    Dim blnAllowChange As Boolean
    Dim intColEnableCount As Integer
    Dim i As Integer
    
    For i = 0 To vfgData.Cols - 1
        If ColumnEnable(i) Then
            intColEnableCount = intColEnableCount + 1
        End If
    Next i
    
    If intColEnableCount > 2 Or vfgData.Rows > 2 Then
        Select Case KeyCode
            Case vbKeyReturn:
                If mblnIsEnterNextCell Then EditNextCellWithCurRow
            Case vbKeyDelete:
                If mblnIsDelKeyRemoveData Then
                    blnIsDel = True
                    
                    RaiseEvent OnDeleteRow(vfgData.RowSel, vfgData.ColSel, blnIsDel)
                    
                    If blnIsDel Then
                        Call DelRow(vfgData.RowSel)
                        Call UpdateRowNumber
                        Call RefreshReadColColor
                        Call RefreshAlign
                    End If
                End If
            Case vbKeySpace:
                If IsCheckboxCol(vfgData.ColSel) Then
                    blnAllowChange = True
                    
                    RaiseEvent OnCheckChanging(vfgData.RowSel, vfgData.ColSel, blnAllowChange)
                    
                    If blnAllowChange Then
                        Call ReCellCheckState(vfgData.RowSel, vfgData.ColSel)
                        RaiseEvent OnCheckChanged(vfgData.RowSel, vfgData.ColSel)
                    End If
                End If
        End Select
    End If
    
    RaiseEvent OnKeyUp(KeyCode, Shift)
    
    err.Clear
End Sub

Private Sub vfgData_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
On Error Resume Next
    RaiseEvent OnKeyUpEdit(Row, Col, KeyCode, Shift)
    
    err.Clear
End Sub

Private Sub vfgData_LeaveCell()
On Error Resume Next
    RaiseEvent OnLeaveCell
    
    err.Clear
End Sub

Private Sub vfgData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
   
   '判断鼠标是否在列头行
   If Button = 2 Then
        If vfgData.MouseRow = 0 Then
            If mblnIsEjectConfig Then
                '调窗体
                Call frmUfgColsList.ShowUfgColsListWindow(Me, mStrDefaultColNames)
                
                If frmUfgColsList.cmdOK.Tag = "True" Then
                     RaiseEvent OnColFormartChange
                     If frmUfgColsList.cmdDefault.Tag = "True" Then RaiseEvent OnColsNameReSet
                End If
                
                Unload frmUfgColsList
            End If
        Else
            mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
        End If
    End If
    
    RaiseEvent OnMouseDown(Button, Shift, X, Y)
    
    err.Clear
End Sub

Private Sub vfgData_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    RaiseEvent OnMouseMove(Button, Shift, X, Y)
    err.Clear
End Sub

Private Sub vfgData_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Dim blnAllowCheck As Boolean
    Dim lngCheckLeft As Long
    Dim picTemp As IPictureDisp
    Dim lngMergeRow As Long
    Dim strMergeValue As String
    
    If vfgData.RowSel <= 0 Or vfgData.RowSel >= vfgData.Rows Then Exit Sub
    If vfgData.ColSel < 0 Or vfgData.ColSel >= vfgData.Cols Then Exit Sub
    
    If Button = 1 Then
        If IsCheckboxCol(vfgData.ColSel) Then
            lngCheckLeft = vfgData.Cell(flexcpLeft, vfgData.RowSel, vfgData.ColSel)
            
            blnAllowCheck = IIf(X > lngCheckLeft And X < lngCheckLeft + 300 Or Trim(vfgData.TextMatrix(vfgData.RowSel, vfgData.ColSel)) = "", True, False)
            If blnAllowCheck Then
                '如果为只读属性，则不允许修改
                If IsReadCol(vfgData.ColSel) Then Exit Sub
                If vfgData.Cell(flexcpPicture, vfgData.RowSel, vfgData.ColSel) Is Nothing Then Exit Sub
                
                '如果check处于禁用状态，则不允许编辑
                If vfgData.Cell(flexcpPicture, vfgData.RowSel, vfgData.ColSel).Tag = csDisCheck Then Exit Sub
                
                RaiseEvent OnCheckChanging(vfgData.RowSel, vfgData.ColSel, blnAllowCheck)
                
                If blnAllowCheck Then
                    If vfgData.Cell(flexcpPicture, vfgData.RowSel, vfgData.ColSel).Tag = 0 Then
                        vfgData.Cell(flexcpPicture, vfgData.RowSel, vfgData.ColSel) = imgCheck(1)
                    Else
                        vfgData.Cell(flexcpPicture, vfgData.RowSel, vfgData.ColSel) = imgCheck(0)
                    End If
                    
                    '判断是否合并行，如果是合并行，则需要设置其他行的check状态
                    If IsMergeCol(vfgData.ColSel) Then
                        
                        strMergeValue = vfgData.TextMatrix(vfgData.RowSel, vfgData.ColSel)
                        
                        lngMergeRow = vfgData.RowSel + 1
                        Do While lngMergeRow < vfgData.Rows
                            If vfgData.TextMatrix(lngMergeRow, vfgData.ColSel) <> UCase(strMergeValue) Then Exit Do
                            
                            If vfgData.Cell(flexcpPicture, lngMergeRow, vfgData.ColSel).Tag = 0 Then
                                vfgData.Cell(flexcpPicture, lngMergeRow, vfgData.ColSel) = imgCheck(1)
                            Else
                                vfgData.Cell(flexcpPicture, lngMergeRow, vfgData.ColSel) = imgCheck(0)
                            End If
                            
                            lngMergeRow = lngMergeRow + 1
                        Loop
                    End If
                    
                    RaiseEvent OnCheckChanged(vfgData.RowSel, vfgData.ColSel)
                End If
            End If
        End If
    ElseIf Button = 2 Then
        If mblnIsShowPopupMenu Then
            mnuCut.Enabled = Not IsReadCol(mDataGrid.Col)
            mnuPaste.Enabled = mnuCut.Enabled
            mnuDel.Enabled = mnuCut.Enabled
            
            Call PopupMenu(menuPop1)
        End If
    End If

    
    RaiseEvent OnMouseUp(Button, Shift, X, Y)
    
    err.Clear
End Sub

Private Sub vfgData_RowColChange()
On Error Resume Next
    RaiseEvent OnRowColChange
    
    err.Clear
End Sub

Private Sub vfgData_SelChange()
On Error Resume Next
    
    Call ShowCellButton
    
    RaiseEvent OnSelChange
    
    err.Clear
End Sub

Private Sub vfgData_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
    Dim blnAllowCancel As Boolean
    
    '只读的列不允许进行编辑,check列只能进行选择
    If (IsReadCol(Col) And Not IsButtonCol(Col)) Or IsCheckboxCol(Col) Then
        Cancel = True
    End If
    
    If vfgData.ColDataType(Col) = flexDTBoolean Then
        blnAllowCancel = True
        
        RaiseEvent OnCheckChanging(Row, Col, blnAllowCancel)
        
        Cancel = Not blnAllowCancel
    Else
        RaiseEvent OnStartEdit(Row, Col, Cancel)
    End If
    
    err.Clear
End Sub



Private Sub UserControl_Initialize()
'初始化列表
    Set mDataGrid = vfgData
    
    Set mrsData = Nothing
    
    chkCheckAll.Visible = False
    
    mblnIsKeepRows = True
    mstrColNames = ""
    mStrDefaultColNames = ""
    mstrDataConvertFormat = ""
    mstrAdoFilter = ""
    mlngErrCellColor = &HC0C0FF
    mblnIsEnterNextCell = True
    mblnIsBtnNextCell = True
    mblnIsCopyAdoMode = False
    mblnIsDelKeyRemoveData = False
    mblnReadOnly = False
    mblnIsAllowExtCol = True
    mlngOldDataRowHeight = -1
    mlngDisableColor = &HC0FFFF
    mblnIsShowNumber = True
    mblnIsShowPopupMenu = True
    mblnIsAutoRowHeight = True
    mblnIsEjectConfig = True
    
    mstrKeyName = ""
    mstrKeyField = ""
    mlngKeepRows = -1
    
    picShowHint.Visible = False
    
    Set mobjHeadFont = New StdFont
    With mobjHeadFont
        .Name = "宋体"
        .Size = 9
        .Bold = False
        .Charset = 134
        .Italic = False
        .Strikethrough = False
        .Underline = False
        .Weight = False
    End With
    
    vfgData.Editable = flexEDKbdMouse
    vfgData.Rows = 1
    vfgData.BackColor = vbWhite
    vfgData.MergeCells = flexMergeRestrictAll
    
    mlngOldBackColor = vfgData.BackColor
    mlngOldGridColor = vfgData.GridColor
    mlngOldDisCellColor = mlngDisableColor
    
    Call InitVsFlexGridList(vfgData, "")
    Call UpdateRowNumber
    Call RefreshAlign
End Sub


Private Sub UserControl_Resize()
'调整控件大小
On Error Resume Next
    vfgData.Left = 0
    vfgData.Top = 0
    vfgData.Width = UserControl.Width
    vfgData.Height = UserControl.Height
    
    picShowHint.Left = 0
    picShowHint.Top = 0
    picShowHint.Width = UserControl.Width
    picShowHint.Height = UserControl.Height
    
    imgWarning.Left = Fix(picShowHint.Width / 3) - Fix(imgWarning.Width / 2)
    imgWarning.Top = Fix((picShowHint.Height - imgWarning.Height) / 2)
    
    labInf.Left = imgWarning.Left + imgWarning.Width + 120
    labInf.Top = imgWarning.Top
    labInf.Height = imgWarning.Height - 960
    labInf.Width = Fix(picShowHint.Width - imgWarning.Left - imgWarning.Width - 480)
    
    '自动填充最后一列的宽度
    Call AutoFitLastCol
    
    Call RefreshCbxPostion
    
    Call ShowCellButton
    
End Sub

Private Sub AutoFitLastCol()
'自动调整最后一列的宽度

'    If mDataGrid.Cols <= 0 Then Exit Sub
'
'    If mDataGrid.Cell(flexcpLeft, 0, mDataGrid.Cols - 1) + mDataGrid.Cell(flexcpWidth, 0, mDataGrid.Cols - 1) < mDataGrid.Width Then
'        mDataGrid.ColWidth(mDataGrid.Cols - 1) = mDataGrid.Width - mDataGrid.Cell(flexcpLeft, 0, mDataGrid.Cols - 1) - 360
'    End If
    
End Sub


Public Sub ShowNullDataFace(vfgList As Object)
'显示无数据的界面
    vfgList.Cols = 10
    vfgList.Rows = 10
    vfgList.FixedCols = 1
    vfgList.FixedRows = 1
End Sub



Public Sub CopyRowData(ByVal lngSourceRow As Long, ByVal lngTargetRow As Long)
'复制行数据
    Dim i As Long
    
    For i = 0 To mDataGrid.Cols - 1
        mDataGrid.TextMatrix(lngTargetRow, i) = mDataGrid.TextMatrix(lngSourceRow, i)
        mDataGrid.Cell(flexcpData, lngTargetRow, i) = mDataGrid.Cell(flexcpData, lngSourceRow, i)
    Next i
    
End Sub


Public Sub RefreshCbxPostion()
'刷新行CheckBox显示

    If GetColIndexWithRowCheck < 0 Then
        chkCheckAll.Visible = False
        Exit Sub
    End If
    
    chkCheckAll.Left = mDataGrid.Cell(flexcpLeft, 0, GetColIndexWithRowCheck()) + 60
    chkCheckAll.Top = mDataGrid.Cell(flexcpHeight, 0, GetColIndexWithRowCheck()) / 2 - 70
    
    vfgData.Cell(flexcpText, 0, GetColIndexWithRowCheck()) = "  " & Replace(vfgData.Cell(flexcpText, 0, GetColIndexWithRowCheck()), "  ", "")
    
    chkCheckAll.Visible = True
End Sub

Public Sub ApplyNormalState()
'更新列表的应用状态为normal
    Dim i As Long
    
    For i = 1 To mDataGrid.Rows - 1
        mDataGrid.RowData(i) = TDataRowState.Normal
    Next i
End Sub


Public Function GetColDateTimeFormat(ByVal lngCol As Long) As String
'获取日期时间格式
    Dim strColProperty As String
    
    GetColDateTimeFormat = ""
    
    If Not RefreshColDicObject(lngCol) Then Exit Function
    
    If mobjTmpDictionary(TColPro.cpDataType) = M_STR_ColProp_TFullDateTime Then
        GetColDateTimeFormat = "yyyy-mm-dd hh:mm:ss"
    ElseIf mobjTmpDictionary(TColPro.cpDataType) = M_STR_ColProp_TOnlyDate Then
        GetColDateTimeFormat = "yyyy-mm-dd"
    ElseIf mobjTmpDictionary(TColPro.cpDataType) = M_STR_ColProp_TOnlyTime Then
        GetColDateTimeFormat = "hh:mm:ss"
    ElseIf mobjTmpDictionary(TColPro.cpDataType) = M_STR_ColProp_TShortDateTime Then
        GetColDateTimeFormat = "yyyy-mm-dd hh:mm"
    Else
        GetColDateTimeFormat = ""
    End If
End Function


Public Sub RestoreRowText(ByVal lngRow As Long)
'恢复文本数据
    Dim i As Integer
    
    For i = IIf(mblnIsAllowExtCol, 1, 0) To mDataGrid.Cols - 1
        If Not mDataGrid.ColHidden(i) Then
            mDataGrid.TextMatrix(lngRow, i) = mDataGrid.Cell(flexcpData, lngRow, i)
        End If
    Next i
End Sub


Public Sub RestoreCurRowText()
'恢复当前文本数据
    Call RestoreRowText(mDataGrid.RowSel)
End Sub


Public Function CheckEquateValue(ByVal lngRow As Long, ByVal lngCol As Long) As String
'检查相同的值，如果有相同的，则返回新值
    Dim strValue As String
    Dim i As Long
    Dim num As Long
    Dim maxNum As Long
    Dim blnFind As Boolean
    
    maxNum = 0
    num = 0
    blnFind = False
    
    CheckEquateValue = ""
    
    strValue = mDataGrid.TextMatrix(lngRow, lngCol)
    If strValue <> "" Then
        For i = 1 To mDataGrid.Rows - 1
            If mDataGrid.TextMatrix(i, lngCol) <> "" And i <> lngRow And Not mDataGrid.RowHidden(i) Then
                If mDataGrid.TextMatrix(i, lngCol) Like strValue & "*" Then
                    num = Val(Substr(mDataGrid.TextMatrix(i, lngCol), _
                        InStr(1, mDataGrid.TextMatrix(i, lngCol), strValue & "_") + LenB(StrConv(strValue & "_", vbFromUnicode)), 3))
                    
                    If num >= maxNum Then maxNum = num + 1
                End If
                
                If mDataGrid.TextMatrix(i, lngCol) = strValue Then
                    blnFind = True
                End If
            End If
        Next i
    End If
    
    If maxNum > 0 Then CheckEquateValue = IIf(blnFind, strValue & "_" & maxNum, "")
End Function


Public Sub ShowObject(obj As Object, ByVal lngRow As Long, ByVal lngCol As Long, Optional ByVal postionType As ObjPostionType = optRight)
'在指定单元格显示对象
    If lngRow < 0 Or lngCol < 0 Then Exit Sub
    If mDataGrid.ColHidden(lngCol) Then Exit Sub
    
    obj.Left = mDataGrid.Left + mDataGrid.Cell(flexcpLeft, lngRow, lngCol) + mDataGrid.Cell(flexcpWidth, lngRow, lngCol) - obj.Width
    obj.Top = mDataGrid.Top + mDataGrid.Cell(flexcpTop, lngRow, lngCol)
    obj.Height = mDataGrid.Cell(flexcpHeight, lngRow, lngCol) - 15
    
    obj.Visible = True
End Sub


Public Sub ShowObjectWithEditCell(obj As Object, Optional ByVal postionType As ObjPostionType = optRight)
'在当前编辑单元格显示对象
    Call ShowObject(obj, mDataGrid.Row, mDataGrid.Col, postionType)
End Sub

Private Function GetImg(ByVal lngImgIndex As Long) As StdPicture
    Set GetImg = Nothing
    
    If mobjImageList Is Nothing Then Exit Function
    
    Set GetImg = mobjImageList.ListImages(lngImgIndex).Picture
End Function


Private Function GetColPropertyValue(ByVal strColProperty As String, ByVal strPropertyName As String) As String
'取得指定列属性中，对应属性名的属性值
    Dim strSubPro As String
    Dim lngTempIndex As Long
    
    GetColPropertyValue = ""
    
    lngTempIndex = InStr(1, UCase(strColProperty), UCase("," & strPropertyName))
    '判断是否存在指定的属性
    If lngTempIndex < 1 Then Exit Function
    
    strSubPro = Mid(strColProperty, lngTempIndex + 1, 100)
    strSubPro = Replace(strSubPro, "|", ",") & ","
    strSubPro = Mid(strSubPro, 1, InStr(1, strSubPro, ","))
    
    GetColPropertyValue = GetNumber(strSubPro)
End Function

Private Sub ConfigFieldConvertDictionary()
'配置列转换字典
    Dim objSubItem As Scripting.Dictionary
    
    Dim strConvertCfg As String
    Dim aryCols() As String
    Dim strColName As String
    Dim strConvertProperty As String
    Dim aryConverts() As String
    Dim strData As String
    Dim strText As String
    Dim i As Long
    Dim j As Long
    
    If Not mobjColDictionary Is Nothing Then
        Call mobjColDictionary.RemoveAll
        Set mobjColDictionary = Nothing
    End If
    
    If mstrDataConvertFormat = "" Then Exit Sub
    
    Set mobjColDictionary = New Scripting.Dictionary
    
    strConvertCfg = "|" & mstrDataConvertFormat & "|"
    
    aryCols = Split(strConvertCfg, "|")
    
    For i = LBound(aryCols) To UBound(aryCols)
        If aryCols(i) <> "" Then
            '新建该列的数据转换字典
            strColName = Mid(aryCols(i), 1, InStr(aryCols(i), ":") - 1)
            strConvertProperty = Mid(aryCols(i), InStr(aryCols(i), ":") + 1, 1024)
            
            If strConvertProperty <> "" Then
                Set objSubItem = New Scripting.Dictionary
                mobjColDictionary.Add strColName, objSubItem
                
                strConvertProperty = "," & strConvertProperty & ","
                aryConverts() = Split(strConvertProperty, ",")
                
                For j = LBound(aryConverts) To UBound(aryConverts)
                    If aryConverts(j) <> "" Then
                        strData = Mid(aryConverts(j), 1, InStr(aryConverts(j), "-") - 1)
                        strText = Mid(aryConverts(j), InStr(aryConverts(j), "-") + 1, 256)
                        
                        mobjColDictionary(strColName).Add strData, strText
                    End If
                Next j
            End If
        End If
    Next i
End Sub


Private Sub InitVsFlexGridList(vfgList As Object, ByVal strCols As String)
'初始化列表对象
    Dim objColPro As Scripting.Dictionary
    Dim i As Integer
    Dim Cols() As String
    Dim aryDefaultCol() As String
    Dim colInf() As String
    Dim strCurCols As String
    
    Dim strTemp As String
    Dim strSubPro As String
    Dim strValue As String
    Dim lngAlign As Long
    Dim strColName As String
    
    
    If Trim(strCols) = "" Then Exit Sub
    
    strCurCols = strCols
    
    
    '根据默认列配置，对需要加载的数据列进行更新
    If mStrDefaultColNames <> "" And mStrDefaultColNames <> strCols Then
        Cols = Split(strCurCols, "|")
        '如果需要加载的数据列在默认列中不存在，则从加载数据列中删除
        For i = 0 To UBound(Cols)
            If Cols(i) <> "" Then
                strColName = Mid(Cols(i), 1, IIf(InStr(Cols(i), ",") <= 0, 255, InStr(Cols(i), ",") - 1))

                If InStr("|" & mStrDefaultColNames & "|", "|" & strColName & ",") <= 0 Then
                    '删除不需加载的数据列
                    strCurCols = Replace(strCurCols, "|" & Cols(i) & "|", "|")
                End If
            End If
        Next i

        aryDefaultCol = Split(mStrDefaultColNames, "|")
        For i = 0 To UBound(aryDefaultCol)
            If aryDefaultCol(i) <> "" Then
                strColName = Mid(aryDefaultCol(i), 1, IIf(InStr(aryDefaultCol(i), ",") <= 0, 255, InStr(aryDefaultCol(i), ",") - 1))

                '如果默认的数据列在需加载的数据列中不存在，则添加
                If InStr("|" & strCurCols & "|", "|" & strColName & ",") <= 0 Then
                    strCurCols = strCurCols & IIf(Right(Trim(strCurCols), 1) = "|", "", "|") & aryDefaultCol(i) & "|"
                End If
            End If
        Next i

        mstrColNames = strCurCols
    End If
    
    
    If mblnIsAllowExtCol Then
        Cols() = Split("|≡" & strCurCols, "|")
    Else
        Cols() = Split(strCurCols, "|")
    End If
    
    vfgList.Cols = 0
    vfgList.Rows = 1
    
    
'    '此句的作用是为了再执行自定义排序后能够触发aftersort事件
'
'    If vfgList.ExplorerBar <= &H1000& Then
'        vfgList.ExplorerBar = vfgList.ExplorerBar + &H1000&
'    End If
    
    For i = LBound(Cols()) To UBound(Cols())
        If Trim(Cols(i)) <> "" Then
            strTemp = Cols(i)
            strTemp = Replace(strTemp, " ", "")
            
            colInf() = Split(strTemp & ",,,,,,", ",")
            
            
            '读取列属性-------------------------------------------------------------------------------------------
            
            Set objColPro = New Scripting.Dictionary
            
            '列名称
            Call objColPro.Add(TColPro.cpColName, Split(colInf(0) & M_STR_NameSplitChar, M_STR_NameSplitChar)(0))
            
            '数据库字段名
            strValue = Split(colInf(0) & M_STR_NameSplitChar, M_STR_NameSplitChar)(1)
            
            '默认不需要更新列样式
            Call objColPro.Add(TColPro.cpIsUpdateStyle, False)
            
            
            If strValue = "" Then
                Call objColPro.Add(TColPro.cpFieldName, Split(colInf(0) & M_STR_NameSplitChar, M_STR_NameSplitChar)(0))
            Else
                Call objColPro.Add(TColPro.cpFieldName, strValue)
            End If
            
            
            '默认选取第一个字段作为关键字
            If i = 1 Then
                mstrKeyName = objColPro(TColPro.cpColName)
                mstrKeyField = objColPro(TColPro.cpFieldName)
            End If
            
            
            '是否隐藏列
            Call objColPro.Add(TColPro.cpIsHide, IIf(InStr(1, UCase(strTemp), UCase("," & M_STR_ColProp_Hide)) >= 1, True, False))
            
            '是否只读列
            Call objColPro.Add(TColPro.cpIsRead, IIf(InStr(1, UCase(strTemp), UCase("," & M_STR_ColProp_Read)) >= 1, True, False))
            
            '是否按钮列
            Call objColPro.Add(TColPro.cpIsBtn, IIf(InStr(1, UCase(strTemp), UCase("," & M_STR_ColProp_Btn)) >= 1, True, False))
            
'            If objColPro(TColPro.cpIsBtn) Then objColPro(TColPro.cpIsUpdateStyle) = True
            
            '是否合并列
            Call objColPro.Add(TColPro.cpIsMerage, IIf(InStr(1, UCase(strTemp), UCase("," & M_STR_ColProp_Merge)) >= 1, True, False))
            
            '是否check列
            Call objColPro.Add(TColPro.cpIsCheck, IIf(InStr(1, UCase(strTemp), UCase("," & M_STR_ColProp_CellCheck)) >= 1, True, False))
            
            If objColPro(TColPro.cpIsCheck) Then objColPro(TColPro.cpIsUpdateStyle) = True
            
            '是否关键字
            Call objColPro.Add(TColPro.cpIsKey, IIf(InStr(1, UCase(strTemp), UCase("," & M_STR_ColProp_Key)) >= 1, True, False))
            
            '是否可选列
            Call objColPro.Add(TColPro.cpIsCombox, IIf(InStr(1, UCase(strTemp), UCase("," & M_STR_ColProp_Cbx)) >= 1, True, False))
            
            '是否为行的check
            Call objColPro.Add(TColPro.cpIsRowCheck, IIf(InStr(1, UCase(strTemp), UCase("," & M_STR_ColProp_RowCheck)) >= 1, True, False))
            
            If objColPro(TColPro.cpIsRowCheck) Then objColPro(TColPro.cpIsUpdateStyle) = True
            
            '是否允许调整列宽
            Call objColPro.Add(TColPro.cpIsUnResize, IIf(InStr(1, strTemp, "," & M_STR_ColProp_UnResize) >= 1, True, False))
            
            '是否允许配置列
            Call objColPro.Add(TColPro.cpIsUnCfg, IIf(InStr(1, strTemp, "," & M_STR_ColProp_UnCfg) >= 1, True, False))
            
            '是否日期列
            If InStr(1, strTemp, "," & M_STR_ColProp_Tdate) > 0 Or _
                InStr(1, strTemp, "," & M_STR_ColProp_TFullDateTime) > 0 Or _
                InStr(1, strTemp, "," & M_STR_ColProp_TOnlyDate) > 0 Or _
                InStr(1, strTemp, "," & M_STR_ColProp_TOnlyTime) > 0 Or _
                InStr(1, strTemp, "," & M_STR_ColProp_TShortDateTime) > 0 Then

                Call objColPro.Add(TColPro.cpIsDateCol, True)
            Else
                Call objColPro.Add(TColPro.cpIsDateCol, False)
            End If
            
            '获取标题列图像索引
            strSubPro = GetColPropertyValue(strTemp, M_STR_ColProp_HeadImg)
            Call objColPro.Add(TColPro.cpHeadImgIndex, IIf(strSubPro = "", -1, Val(strSubPro)))
            
            '获取数据列图像索引
            strSubPro = GetColPropertyValue(strTemp, M_STR_ColProp_DataImg)
            Call objColPro.Add(TColPro.cpDataImgIndex, IIf(strSubPro = "", -1, Val(strSubPro)))
            
            If objColPro(TColPro.cpDataImgIndex) > -1 Then objColPro(TColPro.cpIsUpdateStyle) = True
            
            '列宽度设置
            strSubPro = GetColPropertyValue(strTemp, M_STR_ColProp_WidthTag)
            Call objColPro.Add(TColPro.cpWidth, IIf(strSubPro = "", -1, Val(strSubPro)))
            
            vfgList.Cols = vfgList.Cols + 1
            vfgList.Cell(flexcpText, 0, vfgList.Cols - 1) = objColPro(TColPro.cpColName)
            
            If objColPro(TColPro.cpHeadImgIndex) > -1 And Not objColPro(TColPro.cpIsRowCheck) Then
                Set vfgList.Cell(flexcpPicture, 0, vfgList.Cols - 1) = GetImg(objColPro(TColPro.cpHeadImgIndex))
                
                If Not mobjImageList Is Nothing Then
                    If ScaleY(mobjImageList.ImageHeight, vbPixels, vbTwips) > vfgData.RowHeight(0) Then
                        vfgData.RowHeight(0) = ScaleY(mobjImageList.ImageHeight, vbPixels, vbTwips) + 120
                    End If
                End If
            End If

            If objColPro(TColPro.cpWidth) > 0 Then vfgList.ColWidth(vfgList.Cols - 1) = objColPro(TColPro.cpWidth)
            
            '设置列的关键字
            vfgData.ColKey(vfgList.Cols - 1) = objColPro(TColPro.cpColName)
                
            '设置列的对齐方式
            If InStr(1, UCase(strTemp), "," & UCase(M_STR_ColProp_ColRight)) >= 1 Then
                vfgList.Cell(flexcpAlignment, 0, vfgList.Cols - 1) = flexAlignRightCenter
                
                Call objColPro.Add(TColPro.cpColAlign, flexAlignRightCenter)
            ElseIf InStr(1, UCase(strTemp), "," & UCase(M_STR_ColProp_ColCenter)) >= 1 Then
                vfgList.Cell(flexcpAlignment, 0, vfgList.Cols - 1) = flexAlignCenterCenter
                
                Call objColPro.Add(TColPro.cpColAlign, flexAlignCenterCenter)
            ElseIf InStr(1, UCase(strTemp), "," & UCase(M_STR_ColProp_ColLeft)) >= 1 Then
                vfgList.Cell(flexcpAlignment, 0, vfgList.Cols - 1) = flexAlignLeftCenter
                
                Call objColPro.Add(TColPro.cpColAlign, flexAlignLeftCenter)
            Else
                Call objColPro.Add(TColPro.cpColAlign, vfgList.Cell(flexcpAlignment, 0, vfgList.Cols - 1))
            End If
            
            '设置文本的对齐方式
            If InStr(1, UCase(strTemp), "," & UCase(M_STR_ColProp_TxtRight)) >= 1 Then
                Call objColPro.Add(TColPro.cpTxtAlign, flexAlignLeftCenter)
            ElseIf InStr(1, UCase(strTemp), "," & UCase(M_STR_ColProp_TxtCenter)) >= 1 Then
                Call objColPro.Add(TColPro.cpTxtAlign, flexAlignCenterCenter)
            ElseIf InStr(1, UCase(strTemp), "," & UCase(M_STR_ColProp_TxtLeft)) >= 1 Then
                Call objColPro.Add(TColPro.cpTxtAlign, flexAlignLeftCenter)
            Else
                Call objColPro.Add(TColPro.cpTxtAlign, M_LNG_UNCFG)
            End If
    
            '设置chk的对齐方式
            If InStr(1, UCase(strTemp), "," & UCase(M_STR_ColProp_ChkRight)) >= 1 Then
                Call objColPro.Add(TColPro.cpChkAlign, flexAlignRightCenter)
            ElseIf InStr(1, UCase(strTemp), "," & UCase(M_STR_ColProp_ChkCenter)) >= 1 Then
                Call objColPro.Add(TColPro.cpChkAlign, flexAlignCenterCenter)
            ElseIf InStr(1, UCase(strTemp), "," & UCase(M_STR_ColProp_ChkLeft)) >= 1 Then
                Call objColPro.Add(TColPro.cpChkAlign, flexAlignLeftCenter)
            Else
                Call objColPro.Add(TColPro.cpChkAlign, M_LNG_UNCFG)
            End If
            
            '补齐方式
            Call objColPro.Add(TColPro.cpAlignLen, 0)
            Call objColPro.Add(TColPro.cpAlignChar, "")
            
            lngAlign = InStr(1, UCase(strTemp), UCase(M_STR_ColProp_Align & "<"))
            If lngAlign > 0 Then
                strValue = Mid(strTemp, lngAlign + Len(M_STR_ColProp_Align) + 1, 10)
                strValue = Mid(strValue, 1, InStr(1, UCase(strValue), ">") - 1)
                
                objColPro(TColPro.cpAlignLen) = Val(strValue)
                objColPro(TColPro.cpAlignLen) = Mid(strValue, InStr(1, UCase(strValue), ",") + 1, 3)
            End If
            
            '设置列的数据类型
            If InStr(1, UCase(strTemp), "," & UCase(M_STR_ColProp_TFullDateTime)) >= 1 Then
                Call objColPro.Add(TColPro.cpDataType, M_STR_ColProp_TFullDateTime)
            ElseIf InStr(1, UCase(strTemp), "," & UCase(M_STR_ColProp_TOnlyDate)) >= 1 Then
                Call objColPro.Add(TColPro.cpDataType, M_STR_ColProp_TOnlyDate)
            ElseIf InStr(1, UCase(strTemp), "," & UCase(M_STR_ColProp_TOnlyTime)) >= 1 Then
                Call objColPro.Add(TColPro.cpDataType, M_STR_ColProp_TOnlyTime)
            ElseIf InStr(1, UCase(strTemp), "," & UCase(M_STR_ColProp_TShortDateTime)) >= 1 Then
                Call objColPro.Add(TColPro.cpDataType, M_STR_ColProp_TShortDateTime)
            ElseIf InStr(1, UCase(strTemp), "," & UCase(M_STR_ColProp_Tstr)) >= 1 Then
                Call objColPro.Add(TColPro.cpDataType, M_STR_ColProp_Tstr)
            ElseIf InStr(1, UCase(strTemp), "," & UCase(M_STR_ColProp_Tnum)) >= 1 Then
                Call objColPro.Add(TColPro.cpDataType, M_STR_ColProp_Tnum)
            ElseIf InStr(1, UCase(strTemp), "," & UCase(M_STR_ColProp_Tdate)) >= 1 Then
                Call objColPro.Add(TColPro.cpDataType, M_STR_ColProp_Tdate)
            Else
                Call objColPro.Add(TColPro.cpDataType, M_STR_ColProp_Tstr)
            End If
                        
            
            '隐藏列
            If objColPro(TColPro.cpIsHide) Then
                vfgList.ColHidden(vfgList.Cols - 1) = True
            End If
            
            
            'button列
            If objColPro(TColPro.cpIsBtn) Then
                vfgList.ColComboList(vfgList.Cols - 1) = "..." '不能使用“…”符号
            End If
            
            '合并列
            If objColPro(TColPro.cpIsMerage) Then
                vfgList.MergeCol(vfgList.Cols - 1) = True
            End If
            
            '设置关键字段
            If objColPro(TColPro.cpIsKey) Then
                mstrKeyName = objColPro(TColPro.cpColName)
                mstrKeyField = objColPro(TColPro.cpFieldName)
            End If
            
            
            '设置该列为combox列
            If objColPro(TColPro.cpIsCombox) Then
                strSubPro = Mid(strTemp, InStr(1, UCase(strTemp), UCase("," & M_STR_ColProp_Cbx & "<")) + Len("," & M_STR_ColProp_Cbx & "<"), 100)
                strSubPro = Mid(strSubPro, 1, InStr(1, strSubPro, ">") - 1)
                strSubPro = Replace(strSubPro, ",", "|")
                
                vfgList.ColComboList(vfgList.Cols - 1) = strSubPro
                
                Call objColPro.Add(TColPro.cpComboxCfg, strSubPro)
            End If
            
            
            '设置该列为扩展调整列
            If objColPro(TColPro.cpColName) = M_STR_AdjustColName Then
                vfgList.ColWidth(vfgList.Cols - 1) = 500
                vfgList.ColAlignment(vfgList.Cols - 1) = flexAlignCenterCenter
            End If
            
            Call objColPro.Add(TColPro.cpProperty, objColPro(TColPro.cpFieldName) & M_STR_PropertySplitChar & strTemp)
            '保存字段名和当前列的属性字符串
            Set vfgList.Cell(flexcpData, 0, vfgList.Cols - 1) = objColPro
        End If
    Next i
    
    '当允许改变行高度时，则第一列为固定列
    If mblnIsAllowExtCol Then
        vfgList.FixedCols = 1
        vfgList.AllowUserResizing = flexResizeBoth
    End If
    
    mlngCurColProIndex = -1
    
    If mblnIsKeepRows Then
        vfgList.Rows = IIf(mlngKeepRows <= -1, mDataGrid.Rows, mlngKeepRows)
    End If
    
    '初始化已有行的状态
    For i = 1 To vfgList.Rows - 1
        vfgList.RowData(i) = TDataRowState.Normal
    Next i
    
    '自动填充最后一列的宽度
    Call AutoFitLastCol

    Call ConfigDataFont
    Call ConfigHeadFont
    Call RefreshCbxPostion
End Sub


Private Sub UserControl_Terminate()
    '对象结束事件
'    If Not frmUfgColsList Is Nothing Then
'        Unload frmUfgColsList
'        Set frmUfgColsList = Nothing
'    End If

    If Not mobjColDictionary Is Nothing Then
        mobjColDictionary.RemoveAll
        Set mobjColDictionary = Nothing
    End If
    
    If Not mobjTmpDictionary Is Nothing Then
        mobjTmpDictionary.RemoveAll
        Set mobjTmpDictionary = Nothing
    End If
End Sub




Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'保存属性

    '默认列名
    Call PropBag.WriteProperty("DefaultCols", mStrDefaultColNames)
    '列名
    Call PropBag.WriteProperty("ColNames", mstrColNames, "")
    '关键字
    Call PropBag.WriteProperty("KeyName", mstrKeyName, "")
    '行数量
    Call PropBag.WriteProperty("GridRows", mDataGrid.Rows, 1)
    '是否保持行数
    Call PropBag.WriteProperty("IsKeepRows", mblnIsKeepRows, True)
    '错误单元格颜色
    Call PropBag.WriteProperty("ErrCellColor", mlngErrCellColor, &HC0C0FF)
    '不可编辑单元格颜色
    Call PropBag.WriteProperty("DisCellColor", mlngDisableColor, &HC0FFFF)
    '是否显示行号
    Call PropBag.WriteProperty("IsRowNumber", mblnIsShowNumber, True)
    '背景颜色
    Call PropBag.WriteProperty("BackColor", mDataGrid.BackColor, vbWhite)
    'HeadCheck值
    Call PropBag.WriteProperty("HeadCheckValue", chkCheckAll.value, 0)
    '回车是否自动进入下一编辑单元格
    Call PropBag.WriteProperty("IsEnterNextCell", mblnIsEnterNextCell, True)
    Call PropBag.WriteProperty("IsBtnNextCell", mblnIsBtnNextCell, True)
    '数据字段转换格式串
    Call PropBag.WriteProperty("DataConvertFormat", mstrDataConvertFormat, "")
    'ado过滤条件
    Call PropBag.WriteProperty("AdoFilter", mstrAdoFilter, "")
    '是否ado数据复制模式
    Call PropBag.WriteProperty("IsCopyAdoMode", mblnIsCopyAdoMode, True)
    '是否右键弹出列表配置窗体
    Call PropBag.WriteProperty("IsEjectConfig", mblnIsEjectConfig, False)
    '是否允许del移除数据
    Call PropBag.WriteProperty("IsDelKeyRemoveData", mblnIsDelKeyRemoveData, False)
    '是否允许编辑列表
    Call PropBag.WriteProperty("Editable", mDataGrid.Editable, flexEDKbdMouse)
    '设置合并方式
    Call PropBag.WriteProperty("MeregeCellStyle", mDataGrid.MergeCells, flexMergeRestrictAll)
    '只读属性
    Call PropBag.WriteProperty("ReadOnly", mblnReadOnly, False)
    '扩展列
    Call PropBag.WriteProperty("AllowExtCol", mblnIsAllowExtCol, True)
    '显示右键弹出菜单
    Call PropBag.WriteProperty("IsShowPopupMenu", mblnIsShowPopupMenu, True)
    '是否自动行高
    Call PropBag.WriteProperty("IsAutoRowHeight", mblnIsAutoRowHeight, True)
    
    Call PropBag.WriteProperty("Enabled", vfgData.Enabled, True)
    
    With mobjHeadFont
        Call PropBag.WriteProperty("HeadFontBold", .Bold, False)
        Call PropBag.WriteProperty("HeadFontSize", .Size, 9)
        Call PropBag.WriteProperty("HeadFontCharset", .Charset, "中文 GB2312")
        Call PropBag.WriteProperty("HeadFontItalic", .Italic, False)
        Call PropBag.WriteProperty("HeadFontName", .Name, "宋体")
        Call PropBag.WriteProperty("HeadFontStrikethrough", .Strikethrough, False)
        Call PropBag.WriteProperty("HeadFontUnderline", .Underline, False)
        Call PropBag.WriteProperty("HeadFontWeight", .Weight, False)
    End With
    Call PropBag.WriteProperty("HeadColor", moleHeadColor, &H80000012)
    
    With mDataGrid.Font
        Call PropBag.WriteProperty("DataFontBold", .Bold, False)
        Call PropBag.WriteProperty("DataFontSize", .Size, 9)
        Call PropBag.WriteProperty("DataFontCharset", .Charset, "中文 GB2312")
        Call PropBag.WriteProperty("DataFontItalic", .Italic, False)
        Call PropBag.WriteProperty("DataFontName", .Name, "宋体")
        Call PropBag.WriteProperty("DataFontStrikethrough", .Strikethrough, False)
        Call PropBag.WriteProperty("DataFontUnderline", .Underline, False)
        Call PropBag.WriteProperty("DataFontWeight", .Weight, False)
    End With
    
    Call PropBag.WriteProperty("DataColor", mDataGrid.ForeColor, &H80000012)
    Call PropBag.WriteProperty("GridLineColor", mDataGrid.GridColor, &HC0C0C0)
    Call PropBag.WriteProperty("RowHeightMin", mDataGrid.RowHeightMin, 240)
    Call PropBag.WriteProperty("ExtendLastCol", mDataGrid.ExtendLastCol, False)
End Sub



Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'读取属性

    mStrDefaultColNames = PropBag.ReadProperty("DefaultCols", "")
    mstrColNames = PropBag.ReadProperty("ColNames", "")
    mstrKeyName = PropBag.ReadProperty("KeyName", "")
    mlngErrCellColor = PropBag.ReadProperty("ErrCellColor", &HC0C0FF)
    mlngDisableColor = PropBag.ReadProperty("DisCellColor", &HC0FFFF)
    mblnIsEnterNextCell = PropBag.ReadProperty("IsEnterNextCell", True)
    mblnIsBtnNextCell = PropBag.ReadProperty("IsBtnNextCell", True)
    mblnIsCopyAdoMode = PropBag.ReadProperty("IsCopyAdoMode", False)
    mblnIsEjectConfig = PropBag.ReadProperty("IsEjectConfig", True)
    mstrAdoFilter = PropBag.ReadProperty("AdoFilter", "")
    mstrDataConvertFormat = PropBag.ReadProperty("DataConvertFormat", "")
    mblnIsDelKeyRemoveData = PropBag.ReadProperty("IsDelKeyRemoveData", False)
    mblnIsShowNumber = PropBag.ReadProperty("IsRowNumber", True)
    mblnIsAllowExtCol = PropBag.ReadProperty("AllowExtCol", True)
    mblnIsShowPopupMenu = PropBag.ReadProperty("IsShowPopupMenu", True)
    mblnIsAutoRowHeight = PropBag.ReadProperty("IsAutoRowHeight", True)
    vfgData.Enabled = PropBag.ReadProperty("Enabled", True)
    
    mDataGrid.Editable = PropBag.ReadProperty("Editable", flexEDKbdMouse)
    mDataGrid.MergeCells = PropBag.ReadProperty("MeregeCellStyle", flexMergeRestrictAll)
    mDataGrid.Rows = PropBag.ReadProperty("GridRows", 1)
    mDataGrid.BackColor = PropBag.ReadProperty("BackColor", vbWhite)
    
    With mobjHeadFont
        .Bold = PropBag.ReadProperty("HeadFontBold", .Bold)
        .Size = PropBag.ReadProperty("HeadFontSize", 9)
        .Charset = PropBag.ReadProperty("HeadFontCharset", .Charset)
        .Italic = PropBag.ReadProperty("HeadFontItalic", .Italic)
        .Name = PropBag.ReadProperty("HeadFontName", .Name)
        .Strikethrough = PropBag.ReadProperty("HeadFontStrikethrough", .Strikethrough)
        .Underline = PropBag.ReadProperty("HeadFontUnderline", .Underline)
        .Weight = PropBag.ReadProperty("HeadFontWeight", .Weight)
    End With
    moleHeadColor = PropBag.ReadProperty("HeadColor", &H80000012)
    
    With mDataGrid.Font
        .Bold = PropBag.ReadProperty("DataFontBold", .Bold)
        .Size = PropBag.ReadProperty("DataFontSize", 9)
        .Charset = PropBag.ReadProperty("DataFontCharset", .Charset)
        .Italic = PropBag.ReadProperty("DataFontItalic", .Italic)
        .Name = PropBag.ReadProperty("DataFontName", .Name)
        .Strikethrough = PropBag.ReadProperty("DataFontStrikethrough", .Strikethrough)
        .Underline = PropBag.ReadProperty("DataFontUnderline", .Underline)
        .Weight = PropBag.ReadProperty("DataFontWeight", .Weight)
    End With
    
    mDataGrid.ForeColor = PropBag.ReadProperty("DataColor", &H80000012)
    mDataGrid.ForeColorFixed = mDataGrid.ForeColor
    
    mDataGrid.GridColor = PropBag.ReadProperty("GridLineColor", &HC0C0C0)
    
    mDataGrid.RowHeightMin = PropBag.ReadProperty("RowHeightMin", 240)
    
    mDataGrid.ExtendLastCol = PropBag.ReadProperty("ExtendLastCol", False)
    
    mblnIsKeepRows = PropBag.ReadProperty("IsKeepRows", True)
    mlngKeepRows = IIf(mblnIsKeepRows, vfgData.Rows, -1)
    
    
    chkCheckAll.value = PropBag.ReadProperty("HeadCheckValue", 0)
    
    
    mlngOldBackColor = vfgData.BackColor
    mlngOldGridColor = vfgData.GridColor
    mlngOldDisCellColor = mlngDisableColor
    
    '不能直接对mblnReadOnly赋值，需要使用ReadOnly属性
    ReadOnly = PropBag.ReadProperty("ReadOnly", False)
    
    Call ConfigFieldConvertDictionary
    
    Call InitVsFlexGridList(vfgData, mstrColNames)
    Call UpdateRowNumber
    Call RefreshReadColColor
    Call RefreshAlign
End Sub
