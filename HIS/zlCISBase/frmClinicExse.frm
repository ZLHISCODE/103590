VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmClinicExse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "诊疗收费对照"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13860
   Icon            =   "frmClinicExse.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   13860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdSaveExit 
      Caption         =   "保存退出(&E)"
      Height          =   350
      Left            =   11280
      TabIndex        =   28
      Top             =   6465
      Width           =   1215
   End
   Begin VB.CommandButton cmdAutoGet 
      Caption         =   "智能匹配(&A)"
      Height          =   350
      Left            =   10560
      TabIndex        =   10
      ToolTipText     =   "提示：根据诊疗项目名称自动查找,匹配模式由系统选项--使用习惯决定"
      Top             =   1065
      Width           =   1290
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "恢复(&R)"
      Height          =   350
      Left            =   12000
      Picture         =   "frmClinicExse.frx":058A
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1065
      Width           =   1290
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2715
      Left            =   240
      TabIndex        =   7
      Top             =   7470
      Visible         =   0   'False
      Width           =   6200
      _ExtentX        =   10927
      _ExtentY        =   4789
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.OptionButton optPreproty 
      Caption         =   "手工计价(&2)"
      Height          =   180
      Index           =   2
      Left            =   11985
      TabIndex        =   22
      Top             =   810
      Width           =   1305
   End
   Begin VB.OptionButton optPreproty 
      Caption         =   "不计价(&1)"
      Height          =   180
      Index           =   1
      Left            =   10845
      TabIndex        =   21
      Top             =   810
      Width           =   1110
   End
   Begin VB.OptionButton optPreproty 
      Caption         =   "正常计价(&0)"
      Height          =   180
      Index           =   0
      Left            =   9405
      TabIndex        =   20
      Top             =   810
      Value           =   -1  'True
      Width           =   1410
   End
   Begin VB.Frame fraDept 
      Caption         =   "科室设定"
      Height          =   4725
      Left            =   180
      TabIndex        =   16
      Top             =   1605
      Width           =   2475
      Begin VB.CommandButton cmdCopy 
         Height          =   315
         Left            =   1995
         Picture         =   "frmClinicExse.frx":06D4
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "复制科室"
         Top             =   300
         Width           =   345
      End
      Begin VB.CommandButton cmdDeptDel 
         Height          =   315
         Left            =   1620
         Picture         =   "frmClinicExse.frx":6F26
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "删除科室"
         Top             =   300
         Width           =   345
      End
      Begin VB.TextBox txtDept 
         Height          =   350
         Left            =   105
         TabIndex        =   18
         Top             =   300
         Width           =   1455
      End
      Begin VB.ListBox lstDept 
         Height          =   3840
         Left            =   120
         TabIndex        =   17
         Top             =   690
         Width           =   2250
      End
   End
   Begin VB.Frame fraTotal 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   315
      Left            =   7920
      TabIndex        =   9
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox txtItem 
      Height          =   300
      Left            =   1215
      MaxLength       =   50
      TabIndex        =   2
      Top             =   750
      Width           =   5895
   End
   Begin VB.CommandButton cmdItem 
      Caption         =   "&P"
      Height          =   300
      Left            =   7125
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   750
      Width           =   255
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   4395
      Top             =   390
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
            Picture         =   "frmClinicExse.frx":D778
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicExse.frx":DD12
            Key             =   "ExseUse"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存后下一个(&S)"
      Height          =   350
      Left            =   9600
      TabIndex        =   4
      Top             =   6465
      Width           =   1575
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   300
      Picture         =   "frmClinicExse.frx":E2AC
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6465
      Width           =   1100
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭(&X)"
      Height          =   350
      Left            =   12600
      TabIndex        =   5
      Top             =   6465
      Width           =   1100
   End
   Begin TabDlg.SSTab stbExse 
      Height          =   4740
      Left            =   2880
      TabIndex        =   11
      Top             =   1560
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   8361
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "收费项目(E)"
      TabPicture(0)   =   "frmClinicExse.frx":E3F6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "msfExse"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "检查部位(J)"
      TabPicture(1)   =   "frmClinicExse.frx":E412
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "vfgExse"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "床旁或术中加收(&B)"
      TabPicture(2)   =   "frmClinicExse.frx":E42E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "vfg加收"
      Tab(2).ControlCount=   1
      Begin ZL9BillEdit.BillEdit msfExse 
         Height          =   4125
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "提示：按Del键可删除一行"
         Top             =   480
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   7276
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgExse 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   13
         Top             =   360
         Width           =   10455
         _cx             =   18441
         _cy             =   7435
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
         BackColorFixed  =   15790320
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
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
         AutoResize      =   0   'False
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vfg加收 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   14
         Top             =   360
         Width           =   10455
         _cx             =   18441
         _cy             =   7435
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
         BackColorFixed  =   15790320
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
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
         AutoResize      =   0   'False
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin MSComctlLib.TabStrip tabDept 
      Height          =   5205
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   13605
      _ExtentX        =   23998
      _ExtentY        =   9181
      MultiRow        =   -1  'True
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "所有科室"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "门诊科室"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "住院科室"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "体检科室"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblMessage 
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   4200
      TabIndex        =   27
      Tag             =   "提示:"
      Top             =   6525
      Width           =   5175
   End
   Begin VB.Label lblInfo 
      Caption         =   "医嘱发送后"
      Height          =   255
      Left            =   8280
      TabIndex        =   26
      Top             =   810
      Width           =   1095
   End
   Begin VB.Label txtTotal 
      Height          =   180
      Left            =   2175
      TabIndex        =   25
      Top             =   6525
      Width           =   1815
   End
   Begin VB.Label lbltotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "合计："
      Height          =   180
      Left            =   1665
      TabIndex        =   24
      Top             =   6525
      Width           =   540
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      Caption         =   "诊疗项目(&I)"
      Height          =   180
      Left            =   195
      TabIndex        =   1
      Top             =   810
      Width           =   990
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    请选择项目后，指定其对应的固定价目的收费项目。以便系统能根据诊疗项目对应收费内容，进行病人医嘱的自动计费。"
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   840
      TabIndex        =   0
      Top             =   270
      Width           =   10275
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   195
      Picture         =   "frmClinicExse.frx":E44A
      Top             =   135
      Width           =   480
   End
End
Attribute VB_Name = "frmClinicExse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'说明：
'   1、当前状态：由me.cmdClose.tag保存，分别为"修改"、"查阅"，由上级程序通过ShowMe函数传入
'   2、指定项目：由me.lblItem.tag保存，由上级程序通过ShowMe函数传入，可以传递，也可以不传递
'可选的收费项目：
'   1、不包括类别为挂号、床位和其他非固定价目的收费项目
'   2、由于药疗的收费通过规格对应，因此现在不允许将药品作为对应收费项目
'---------------------------------------------------
Private strInputed As String
Public rsSelect As New ADODB.Recordset                          '收费项目选择器传回时用到
Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
Dim strTemp As String
Dim intCount As Integer
Dim mrsBwff As ADODB.Recordset '记录部位方法,用于输入的合法性检查
Dim mstrType As String          '当前诊疗项目类别
Dim mstrOper As String          '当前诊疗项目的操作类型
Dim mlngFlag As Long            '当前诊疗项目的执行标记
Dim mlngClient As Long          '当前诊疗项目的服务对象
Private mstr类别 As String      '记录所选择的类别
Private mstrIDS As String       '保存之后id串
Private mlngItem As Long        '当前的位置

Dim mlngLastSource As Long      '上次选择的页面对应的病人来源
Dim mlngLastDeptID As Long      '上次选择的科室ID
Private mDelDeptList As String  '删除科室列表，保存时要先处理这里面的数据
Private mlngCodeType As Long '0-拼音，1-五笔

Private Enum ExseCol
    序号 = 0
    项目名 = 1
    规格 = 2
    单位 = 3
    当前价 = 4
    对应数 = 5
    固定 = 6
    从项 = 7
    收费方式 = 8
    
    列数 = 9
End Enum
'科室列表
Private mDept1() As String   '缓存门诊 科室列表
Private mDept2() As String   '缓存住院 科室列表
Private mDept3() As String   '缓存体检 科室列表
'普通收费对照
Private mGen0()  As String '缓存全院 普通收费项目
Private mGen1()  As String '缓存门诊 普通收费项目
Private mGen2()  As String '缓存住院 普通收费项目
Private mGen3()  As String '缓存体检 普通收费项目
'部位收费对照
Private mPlace0() As String '缓存全院 带部位收费项目
Private mPlace1() As String '缓存门诊 带部位收费项目
Private mPlace2() As String '缓存住院 带部位收费项目
Private mPlace3() As String '缓存体检 带部位收费项目
'附加收费对照
Private mAppend0() As String '缓存全院 附加收费项目
Private mAppend1() As String '缓存门诊 附加收费项目
Private mAppend2() As String '缓存住院 附加收费项目
Private mAppend3() As String '缓存体检 附加收费项目

'  诊疗项目id_In In 诊疗收费关系.诊疗项目id%Type,
'  计价性质_In   诊疗项目目录.计价性质%Type,
'  收费内容_In   In Varchar2, --以"|"分隔的诊疗收费内容，每条记录按"收费项目ID^数量^固定^从项^性质^部位^检查方法^收费方式"组织
'  是否删除_In   Number := 1,
'  适用科室id_In 诊疗收费关系.适用科室id%Type := Null,
'  病人来源_In   诊疗收费关系.病人来源%Type := 0

Private Sub IniItemList()
    With Me.msfExse
        .Active = True
        .ClearBill
        .MsfObj.FixedCols = 1
        .Rows = 2
        .Cols = ExseCol.列数
        .TextMatrix(0, ExseCol.序号) = ""
        .TextMatrix(0, ExseCol.项目名) = "项目名"
        .TextMatrix(0, ExseCol.规格) = "规格"
        .TextMatrix(0, ExseCol.单位) = "单位"
        .TextMatrix(0, ExseCol.当前价) = "当前价"
        .TextMatrix(0, ExseCol.对应数) = "对应数"
        .TextMatrix(0, ExseCol.固定) = "固定"
        .TextMatrix(0, ExseCol.从项) = "从项"
        .TextMatrix(0, ExseCol.收费方式) = "收费方式"
        .colData(ExseCol.序号) = 5
        .colData(ExseCol.项目名) = 1
        .colData(ExseCol.规格) = 5
        .colData(ExseCol.单位) = 5
        .colData(ExseCol.当前价) = 5
        .colData(ExseCol.对应数) = 4
        .colData(ExseCol.固定) = -1
        .colData(ExseCol.从项) = -1
        .colData(ExseCol.收费方式) = 3
        .ColWidth(ExseCol.序号) = 250
        .ColWidth(ExseCol.项目名) = 2800
        .ColWidth(ExseCol.规格) = 1000
        .ColWidth(ExseCol.单位) = 600
        .ColWidth(ExseCol.当前价) = 800
        .ColWidth(ExseCol.对应数) = 1000
        .ColWidth(ExseCol.固定) = 500
        .ColWidth(ExseCol.从项) = 500
        .ColWidth(ExseCol.收费方式) = 3100
        
        .PrimaryCol = 1: .LocateCol = 1
        .Row = 1: .Col = 1
        .ColAlignment(ExseCol.固定) = 4
        .ColAlignment(ExseCol.从项) = 4
        .ColAlignment(ExseCol.收费方式) = 1
        
        .Clear
        .AddItem ("0-正常收取")
        
        '检验类、治疗类(采集方式)的项目
        If mstrType = "C" Or (mstrType = "E" And mstrOper = "6") Then
            .AddItem ("1-检验试管费用")
        End If
        .AddItem ("2-一次发送只收取一次")
        .AddItem ("3-当天只收取一次")
        .AddItem ("4-当天未执行收取一次")
        .AddItem ("5-当天只收取一次，排斥其他项目")
        .AddItem ("6-当天未执行收取一次，排斥其他项目")
        .AddItem ("7-每天首次不收取")
        .AddItem ("9-自定义")
    End With
    
End Sub

Public Sub ShowMe(ByVal frmParent As Object, ByVal blnEdit As Boolean, Optional ByVal lng项目id As Long, Optional ByVal strIDS As String)
    '---------------------------------------------------
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '---------------------------------------------------
    mstrIDS = strIDS
    If mstrIDS = "" Then Me.cmdSave.Visible = False
    Me.cmdClose.Tag = IIf(blnEdit, "修改", "查阅")
    If Me.cmdClose.Tag = "查阅" Then
        Me.msfExse.Active = False
        Me.cmdSave.Visible = False
        
        Me.cmdRestore.Visible = False
        Me.cmdAutoGet.Visible = False
        
        txtDept.Enabled = False
        cmdDeptDel.Enabled = False
        cmdCopy.Enabled = False
        cmdItem.Enabled = False
        txtItem.Enabled = False
    Else
        Me.msfExse.Active = True
    End If
    Me.lblItem.Tag = lng项目id
    
    '得到数据
    GetOneRec
    
    If mstrOper = "病理" And mstr类别 = "D" Then
        stbExse.TabVisible(1) = False
    End If
    
    Set mrsBwff = zlDatabase.OpenSQLRecord("Select a.名称 as 部位,a.方法 From 诊疗检查部位 a, 诊疗项目目录 b Where a.类型 = b.操作类型 And b.Id = [1]", Me.Caption, lng项目id)
    
    Me.Show 1, frmParent

End Sub

Private Sub GetOneRec()
    err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,I.编码,I.名称,I.计算单位,nvl(I.计价性质,0) as 计价性质,I.类别,I.操作类型,I.执行标记,I.服务对象 " & _
            " from 诊疗项目目录 I" & _
            " where I.类别>='A' and I.ID=[1] " & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.lblItem.Tag))
    
    With rsTemp
        If .BOF Or .EOF Then
            Me.lblItem.Tag = 0: Me.txtItem.Tag = "": Me.txtItem.Text = Me.txtItem.Tag
        Else
            mstr类别 = !类别
            Me.lblItem.Tag = !ID: Me.txtItem.Tag = "[" & !编码 & "]" & !名称: Me.txtItem.Text = Me.txtItem.Tag
            Me.optPreproty(!计价性质).Value = True
            mstrType = Trim("" & !类别)
            mstrOper = IIf(IsNull(!操作类型), "", !操作类型)
            mlngFlag = Val("" & !执行标记)
            mlngClient = Val("" & !服务对象)
            Call zlExseRef(Me.lblItem.Tag)
        End If
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdClose_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdCopy_Click()
    Dim lngDeptID As Long, lngSource As Long
    If Val(Me.lblItem.Tag) <= 0 Then Exit Sub
    If lstDept.ListIndex <> -1 Then
        lngDeptID = lstDept.ItemData(lstDept.ListIndex)
        lngSource = GetCurrSource
        Call DeptCopy(lngSource, lngDeptID)
    Else
        MsgBox "请先选择一个科室！", vbInformation
    End If
    
End Sub

Private Sub cmdDeptDel_Click()
    Dim lngDeptID As Long, lngSource As Long
    
    If Val(Me.lblItem.Tag) <= 0 Then Exit Sub
    If lstDept.ListIndex <> -1 Then
        lngDeptID = lstDept.ItemData(lstDept.ListIndex)
        lngSource = GetCurrSource
        If delDept(lngSource, lngDeptID) Then
            '保存到删除科室列表，保存时先删除这个表里面的科室对照
            If InStr(mDelDeptList & ",", "," & lngSource & "|" & lngDeptID & ",") <= 0 Then
                mDelDeptList = mDelDeptList & "," & lngSource & "|" & lngDeptID
            End If
            lstDept.RemoveItem lstDept.ListIndex
        End If
        'Call zlExseRef(Me.lblItem.Tag)
        Call msfExseRef(lngDeptID, -1)
    Else
        MsgBox "请先选择一个科室！", vbInformation
    End If
End Sub

Private Sub cmdRestore_Click()
    Call zlExseRef(Me.lblItem.Tag)
    strInputed = ""
End Sub

Private Sub cmdSave_Click()
    Dim lngId As Long
    Dim lngCount As Long
    SaveData
    
    lngId = Split(mstrIDS, ",")(mlngItem)
    Me.lblItem.Tag = lngId
    
    '得到数据
    GetOneRec
    Me.txtItem.SetFocus
    Call zlCommFun.PressKey(vbKeyTab)
    mlngItem = mlngItem + 1
    If mlngItem = UBound(Split(mstrIDS, ",")) Then Me.cmdSave.Enabled = False
End Sub
Private Sub cmdSaveExit_Click()
    SaveData
    Unload Me
End Sub
Private Sub SaveData()
    Dim blnErr As Boolean
    Dim i As Integer
    Dim varDelList As Variant
    Dim NullGen(0) As String, NullPlan(0) As String, NullAppend(0) As String
    On Error GoTo hErr
    If Val(Me.lblItem.Tag) = 0 Then lblMessage.Caption = lblMessage.Tag & "未正确指定诊疗项目！": Me.txtItem.SetFocus: Exit Sub
    
    '先删除列表中的科室数据
    If mDelDeptList <> "" Then
        If Left(mDelDeptList, 1) = "," Then mDelDeptList = Mid(mDelDeptList, 2)
        varDelList = Split(mDelDeptList, ",")
        For i = LBound(varDelList) To UBound(varDelList)
            Call SaveArryData(CLng(Split(varDelList(i), "|")(0)), CLng(Split(varDelList(i), "|")(1)), NullGen, NullPlan, NullAppend)
        Next
        mDelDeptList = ""
    End If
    
    Call lstDeptSelect(1) '保存当前界面上的数据到缓存
    If mGen0(UBound(mGen0)) <> "" Then
        blnErr = CheckArrData(mGen0)
        If blnErr = False Then Exit Sub
    End If
    If mGen1(UBound(mGen1)) <> "" Then
        blnErr = CheckArrData(mGen1)
        If blnErr = False Then Exit Sub
    End If
    If mGen2(UBound(mGen2)) <> "" Then
        blnErr = CheckArrData(mGen2)
        If blnErr = False Then Exit Sub
    End If
    If mGen3(UBound(mGen3)) <> "" Then
        blnErr = CheckArrData(mGen3)
        If blnErr = False Then Exit Sub
    End If
    Call SaveArryData(0, 0, mGen0, mPlace0, mAppend0)
    
    For i = LBound(mDept1) To UBound(mDept1)
        If mDept1(i) <> "" Then
            Call SaveArryData(1, Val(Split(mDept1(i), "|")(0)), mGen1, mPlace1, mAppend1)
        End If
    Next
    
    For i = LBound(mDept2) To UBound(mDept2)
        If mDept2(i) <> "" Then
            Call SaveArryData(2, Val(Split(mDept2(i), "|")(0)), mGen2, mPlace2, mAppend2)
        End If
    Next
    
    For i = LBound(mDept3) To UBound(mDept3)
        If mDept3(i) <> "" Then
            Call SaveArryData(3, Val(Split(mDept3(i), "|")(0)), mGen3, mPlace3, mAppend3)
        End If
    Next
    
    lblMessage.Caption = lblMessage.Tag & Mid(Me.txtItem.Text, 1, 18) & " 收费对照保存成功！"
    Call zlExseRef(Me.lblItem.Tag)
    Me.txtItem.SetFocus
    Exit Sub
hErr:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdItem_Click()
    Dim rsTmp As ADODB.Recordset
    
    err = 0: On Error GoTo ErrHand

    gstrSql = "Select Distinct 0 As 末级,id,上级ID,编码,名称,'' As 计算单位,0 As 计价性质,'' As 类别, '' As 操作类型,0 as 执行标记, 0 as 服务对象 " & _
        " From 诊疗分类目录 Start With id In (Select 分类ID" & _
            " from 诊疗项目目录 I" & _
            " where I.类别>='A'" & _
            " And (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))) Connect By Prior 上级id=id" & _
        " Union All"
    gstrSql = gstrSql & " Select 1 As 末级,I.ID,分类ID As 上级id,I.编码,I.名称,I.计算单位,nvl(I.计价性质,0) as 计价性质, 类别, 操作类型, 执行标记, 服务对象 " & _
            " from 诊疗项目目录 I" & _
            " where I.类别>='A'" & _
            " And (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD')) Order By 编码"
    Set rsTmp = zlDatabase.ShowSelect(Me, gstrSql, 2, "诊疗项目", , , , , True)
    If Not rsTmp Is Nothing Then
        mstr类别 = rsTmp!类别
        Me.lblItem.Tag = rsTmp("ID")
        Me.txtItem.Tag = "[" & rsTmp("编码") & "]" & rsTmp("名称")
        Me.txtItem.Text = Me.txtItem.Tag
        Me.optPreproty(rsTmp("计价性质")).Value = True
        mstrType = Trim("" & rsTmp("类别"))
        mstrOper = IIf(IsNull(rsTmp("操作类型")), "", rsTmp("操作类型"))
        mlngFlag = Val("" & rsTmp!执行标记)
        mlngClient = Val("" & rsTmp!服务对象)
        Call zlExseRef(Me.lblItem.Tag)
        Me.txtItem.SetFocus
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    If Me.lvwItems.Visible Then
        Me.lvwItems.Visible = False
        If Me.lvwItems.Tag = Me.txtItem.Name Then
            Me.txtItem.SetFocus
        Else
            Me.msfExse.SetFocus
        End If

    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mlngCodeType = zlDatabase.GetPara("简码方式")
    Call IniItemList
    mDelDeptList = "" '清空删除科室列表
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 2000
        .Add , "规格", "规格", 1200
        .Add , "编码", "编码", 1000
        .Add , "计算单位", "单位", 800
        .Add , "产地", "产地", 1000
        .Add , "售价", "售价", 1000
        .Add , "类别", "类别", 0
        .Add , "操作类型", "操作类型", 0
    End With
    With Me.lvwItems
        .ColumnHeaders("编码").Position = 1
        .SortKey = .ColumnHeaders("编码").Index - 1
        .SortOrder = lvwAscending
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlngItem = 0
End Sub

Private Sub lstDept_Click()
    Dim lngDeptID As Long
    '显示当前科室的费用对照
    If lstDept.ListIndex >= 0 Then
        lngDeptID = lstDept.ItemData(lstDept.ListIndex)
        Call msfExseRef(lngDeptID, 1)
        txtDept.Text = lstDept.List(lstDept.ListIndex)
    Else
        '未选 中科室，清空显示
        Call msfExseRef(-1, 1)
        txtDept.Text = ""
    End If
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    Dim dblCurrJe As Double
    
    On Error GoTo ErrHandle
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwItems
        If .Tag = Me.txtItem.Name Then
            If Me.lblItem.Tag <> Mid(.SelectedItem.Key, 2) Then
                Me.lblItem.Tag = Mid(.SelectedItem.Key, 2)
                Me.txtItem.Tag = "[" & .SelectedItem.SubItems(.ColumnHeaders("编码").Index - 1) & "]" & .SelectedItem.Text
                Me.txtItem.Text = Me.txtItem.Tag
                Me.optPreproty(Val(.SelectedItem.Tag)).Value = True
                mstrType = .SelectedItem.SubItems(.ColumnHeaders("类别").Index - 1)
                mstrOper = .SelectedItem.SubItems(.ColumnHeaders("操作类型").Index - 1)
                Call zlExseRef(Me.lblItem.Tag)
            End If
            Me.txtItem.SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            dblCurrJe = Val(Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.当前价)) * Val(Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.对应数))
            
            Me.msfExse.Text = "[" & .SelectedItem.SubItems(.ColumnHeaders("编码").Index - 1) & "]" & .SelectedItem.Text
            Me.msfExse.RowData(Me.msfExse.Row) = Mid(.SelectedItem.Key, 2)
            Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.项目名) = Me.msfExse.Text
            Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.规格) = .SelectedItem.SubItems(.ColumnHeaders("规格").Index - 1)
            Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.单位) = .SelectedItem.SubItems(.ColumnHeaders("计算单位").Index - 1)
            If Val(Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.对应数)) = 0 Then
                Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.对应数) = "1"
                Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.固定) = "√"
            End If
            
            gstrSql = "select decode(I.是否变价,1,'变价',to_char(P.价格)) As 价格" & _
                    " from (select 是否变价 from 收费项目目录 where id=[1]) I," & _
                    "      (Select sum(现价) As 价格" & _
                    "      From 收费价目  Where 价格等级 Is Null and 收费细目id=[1] and 执行日期<=Sysdate And (终止日期 Is Null Or 终止日期>=Sysdate)) P"
                    
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.msfExse.RowData(Me.msfExse.Row)), gstrPriceClass)
            
            With rsTemp
                If .RecordCount > 0 Then
                    Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.当前价) = IIf(IsNull(!价格), "", !价格)
                Else
                    Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.当前价) = ""
                End If
            
                txtTotal = Val(txtTotal) - dblCurrJe + Val(Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.当前价)) * Val(Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.对应数))
                txtTotal = IIf(Val(txtTotal) = 0, "", Format(txtTotal, "0.0000"))
            End With
            Me.msfExse.SetFocus
            Call zlCommFun.PressKey(vbKeyRight)
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        Call lvwItems_DblClick
    End Select
End Sub

Private Sub lvwItems_LostFocus()
    Me.lvwItems.Visible = False
End Sub

Private Sub msfExse_AfterAddRow(Row As Long)
    With Me.msfExse
        If .Rows > 2 Then
            .TextMatrix(1, ExseCol.序号) = 1
        End If
        For intCount = Row To .Rows - 1
            .TextMatrix(intCount, ExseCol.序号) = intCount
            If .Rows > 2 Then
                .TextMatrix(intCount, ExseCol.从项) = .TextMatrix(intCount - 1, ExseCol.从项)
            End If
        Next
    End With
End Sub

Private Sub msfExse_AfterDeleteRow()
    With Me.msfExse
        For intCount = IIf(.Row <> 1, .Row - 1, .Row) To .Rows - 1
            .TextMatrix(intCount, ExseCol.序号) = intCount
        Next
    End With
End Sub

Private Sub msfExse_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    txtTotal = Val(txtTotal) - Val(Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.当前价)) * Val(Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.对应数))
    txtTotal = IIf(Val(txtTotal) = 0, "", Format(txtTotal, "0.0000"))
End Sub

Private Sub msfExse_CommandClick()
    Dim rsTmp As ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim dblCurrJe As Double
    err = 0: On Error GoTo ErrHand
    Dim strKey As String
    Dim str服务对象 As String

    '调出选择器
    strKey = tabDept.SelectedItem.Key
    
    If strKey = "ALL" Or strKey = "TJ" Then
        str服务对象 = "1,2,3"
    ElseIf strKey = "MZ" Then
        str服务对象 = "1,3"
    ElseIf strKey = "ZY" Then
        str服务对象 = "2,3"
    End If
    
    frmClinicExseSelect.ShowMe Me, str服务对象
    Set rsTmp = rsSelect
    
    If Not rsTmp Is Nothing And rsTmp.State = 1 Then
        Me.msfExse.Text = "[" & rsTmp("编码") & "]" & rsTmp("名称")
        Me.msfExse.RowData(Me.msfExse.Row) = rsTmp("ID")
        Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.项目名) = Me.msfExse.Text
        Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.规格) = IIf(IsNull(rsTmp("规格")), "", rsTmp("规格"))
        Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.单位) = IIf(IsNull(rsTmp("计算单位")), "", rsTmp("计算单位"))
        If Val(Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.对应数)) = 0 Then
            Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.对应数) = "1"
            Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.固定) = "√"
        End If
        
        gstrSql = "select decode(I.是否变价,1,'变价',to_char(P.价格)) As 价格" & _
                " from (select 是否变价 from 收费项目目录 where id=[1]) I," & _
                "      (Select sum(现价) As 价格" & _
                "      From 收费价目 " & _
                "      Where 收费细目id=[1]  And 价格等级 Is NULL  and 执行日期<=Sysdate And (终止日期 Is Null Or 终止日期>=Sysdate)) P"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.msfExse.RowData(Me.msfExse.Row)), gstrPriceClass)
        
        With rsTemp
            If .RecordCount > 0 Then
                Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.当前价) = IIf(IsNull(!价格), "", !价格)
            Else
                Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.当前价) = ""
            End If
        
            txtTotal = Val(txtTotal) - dblCurrJe + Val(Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.当前价)) * Val(Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.对应数))
            txtTotal = IIf(Val(txtTotal) = 0, "", Format(txtTotal, "0.0000"))
        End With
        Me.msfExse.SetFocus
        Call zlCommFun.PressKey(vbKeyRight)
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub msfExse_DblClick(Cancel As Boolean)
    If msfExse.Col = ExseCol.固定 Then
        With msfExse
            If .TextMatrix(.Row, ExseCol.对应数) <> "" And IsNumeric(.TextMatrix(.Row, ExseCol.对应数)) Then
                If Int(.TextMatrix(.Row, ExseCol.对应数)) = 0 Then
                    Cancel = True
                    .TextMatrix(msfExse.Row, ExseCol.固定) = ""
                    lblMessage.Caption = lblMessage.Tag & "对应数设置为0时只能做为非固定项."
                End If
            End If
        End With
    End If
End Sub

Private Sub msfExse_EnterCell(Row As Long, Col As Long)
    strInputed = Me.msfExse.TextMatrix(Row, Col)
End Sub

Private Sub msfExse_GotFocus()
    If Me.lvwItems.Visible Then Me.lvwItems.SetFocus
End Sub

Private Sub msfExse_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim dblCurrJe As Double
    Dim str服务对象 As String
    
    If KeyCode <> vbKeyReturn Then
        If msfExse.Col = ExseCol.固定 And KeyCode = vbKeySpace Then  '设置固定项时检查
            With msfExse
                 If Int(Nvl(.TextMatrix(.Row, ExseCol.对应数), 0)) = 0 Then
                        Cancel = True
                        .Text = ""
                         .TextMatrix(msfExse.Row, ExseCol.固定) = ""
                         lblMessage.Caption = lblMessage.Tag & "对应数为0时不应设置为固定项."
                 End If
            End With
        End If
        Exit Sub
    End If
    
    lblMessage.Caption = ""
    
    With Me.msfExse
        If .Active = False Then Exit Sub
        If .Col = ExseCol.当前价 And .TxtVisible Then
            .Text = Format(.Text, "0.00000"): .TextMatrix(.Row, ExseCol.当前价) = .Text
        End If
        If .Col <> ExseCol.项目名 Then
            '对应数为零时,不能设置为固定项
            If .Col = ExseCol.对应数 Then
                If Not IsNumeric(Nvl(.Text, "X")) Then
                    lblMessage.Caption = lblMessage.Tag & "对应数不能为空，并且要求设置为数字型."
                    .TxtSetFocus
                    Exit Sub
                End If
                If Int(.Text) = 0 And .TextMatrix(.Row, ExseCol.固定) = "√" Then
                    .TextMatrix(.Row, ExseCol.固定) = ""
                    lblMessage.Caption = lblMessage.Tag & "对应数设置为0时已自动设置为非固定项."
                End If
            End If
            Exit Sub
        End If
        If .TxtVisible = False Then
            If .TextMatrix(.Row, ExseCol.项目名) = "" Then Exit Sub
            strTemp = Trim(.TextMatrix(.Row, ExseCol.项目名))
        Else
            If Trim(.Text) = "" Then Exit Sub
            strTemp = Trim(.Text)
        End If
    End With
    If Trim(strTemp) = Trim(strInputed) Then Exit Sub
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    err = 0: On Error GoTo ErrHand
    strTemp = UCase(strTemp)
    If tabDept.SelectedItem.Key = "ALL" Or tabDept.SelectedItem.Key = "TJ" Then
        str服务对象 = " And (I.服务对象=1 or I.服务对象=2 or I.服务对象=3) "
    ElseIf tabDept.SelectedItem.Key = "MZ" Then
        str服务对象 = " And (I.服务对象=1 or I.服务对象=3) "
    ElseIf tabDept.SelectedItem.Key = "ZY" Then
        str服务对象 = " And (I.服务对象=2 or I.服务对象=3) "
    End If
    gstrSql = "Select c.Id, c.编码, c.名称, c.规格, c.产地, c.计算单位," & vbNewLine & _
            "       Decode(Nvl(c.是否变价, 0)," & vbNewLine & _
            "               0," & vbNewLine & _
            "               Ltrim(Rtrim(To_Char(Nvl(d.现价, 0), '9999999990.0000')))," & vbNewLine & _
            "               Decode(Instr('4567', c.类别), 0, Ltrim(Rtrim(To_Char(d.缺省价格, 0), '9999999990.0000')), '时价')) as 售价" & vbNewLine & _
            "From (Select Distinct (a.Id), a.编码, a.名称, a.规格, a.产地, a.计算单位, a.是否变价, a.类别" & vbNewLine & _
            "       From 收费项目目录 a, 收费项目别名 b" & vbNewLine & _
            "       Where a.Id = b.收费细目id And a.类别 Not In ('1', 'J') And" & vbNewLine & _
            "             (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
            "             (a.服务对象 = 1 Or a.服务对象 = 2 Or a.服务对象 = 3) And" & vbNewLine & _
            "             (a.编码 Like [1] Or b.名称 Like [2] Or b.简码 Like [2]) And b.码类 = [4]) c," & vbNewLine & _
            "     收费价目 d" & vbNewLine & _
            "Where c.Id = d.收费细目id(+) And d.执行日期 <= Sysdate And (d.终止日期 > Sysdate Or d.终止日期 Is Null) And d.价格等级 Is Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strTemp & "%", gstrMatch & strTemp & "%", gstrPriceClass, mlngCodeType + 1)
    
    If rsTemp.BOF Or rsTemp.EOF Then
        Me.msfExse.Text = ""
        Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.项目名) = ""
        lblMessage.Caption = lblMessage.Tag & "未找到指定收费项目！"
        Exit Sub
    End If
    
    If rsTemp.RecordCount = 1 Then
        Me.msfExse.Text = "[" & rsTemp!编码 & "]" & rsTemp!名称
        Me.msfExse.RowData(Me.msfExse.Row) = rsTemp!ID
        Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.项目名) = Me.msfExse.Text
        Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.规格) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
        Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.单位) = IIf(IsNull(rsTemp!计算单位), "", rsTemp!计算单位)
        If Val(Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.对应数)) = 0 Then
            Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.对应数) = "1"
            Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.固定) = "√"
        End If
       
        Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.当前价) = IIf(IsNull(rsTemp!售价), "", rsTemp!售价)
        
        txtTotal = Val(txtTotal) - dblCurrJe + Val(Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.当前价)) * Val(Me.msfExse.TextMatrix(Me.msfExse.Row, ExseCol.对应数))
        txtTotal = IIf(Val(txtTotal) = 0, "", Format(txtTotal, "0.0000"))
        Exit Sub
    End If
    Me.lvwItems.ListItems.Clear
    Do While Not rsTemp.EOF
        Set objItem = Me.lvwItems.ListItems.Add(, "_" & rsTemp!ID, rsTemp!名称)
        objItem.Icon = "ExseUse": objItem.SmallIcon = "ExseUse"
        objItem.SubItems(Me.lvwItems.ColumnHeaders("规格").Index - 1) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
        objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = rsTemp!编码
        objItem.SubItems(Me.lvwItems.ColumnHeaders("计算单位").Index - 1) = IIf(IsNull(rsTemp!计算单位), "", rsTemp!计算单位)
        objItem.SubItems(Me.lvwItems.ColumnHeaders("产地").Index - 1) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
        objItem.SubItems(Me.lvwItems.ColumnHeaders("售价").Index - 1) = IIf(IsNull(rsTemp!售价), "", rsTemp!售价)
        rsTemp.MoveNext
    Loop
    Me.lvwItems.ListItems(1).Selected = True
  
    With Me.lvwItems
        .Tag = Me.msfExse.Name
        .Left = stbExse.Left + Me.msfExse.Left + 300
        .Top = stbExse.Top + Me.msfExse.Top + msfExse.RowHeight(msfExse.Row) * (msfExse.Row + 1)
        '.Height = .ListItems(1).Height * (.ListItems.Count + 1)
        If .Top > Me.msfExse.Top + Me.msfExse.Height Then
            .Top = Me.msfExse.Top + Me.msfExse.Height
        End If
        .Height = Me.Height - .Top - .ListItems(1).Top * 2
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Cancel = True
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdAutoGet_Click()
'功能：根据诊疗项目名称，自动找出对应的收费项目
Dim strTemp As String
Dim n As Integer

strTemp = Mid(txtItem.Text, InStr(1, txtItem.Text, "]") + 1)
With msfExse
    If .Col <> ExseCol.项目名 Then .Col = ExseCol.项目名
    .SetFocus
    If .TextMatrix(.Rows - 1, ExseCol.项目名) <> "" Then
        .Rows = .Rows + 1
    End If
    .TextMatrix(.Rows - 1, ExseCol.项目名) = strTemp
    .Row = .Rows - 1
    strInputed = ""
    
    '检查是否重复
    If .Rows > 2 Then
        For n = 1 To .Rows - 2
            If strTemp = Trim(Mid(.TextMatrix(n, ExseCol.项目名), InStr(1, .TextMatrix(n, ExseCol.项目名), "]") + 1)) Then
                .Rows = .Rows - 1
                Exit Sub
            End If
        Next
    End If
End With
Call msfExse_KeyDown(vbKeyReturn, 0, False)

End Sub

Private Sub msfExse_LeaveCell(Row As Long, Col As Long)
    Select Case Col
        Case ExseCol.对应数
            txtTotal = Val(txtTotal) + Val(Me.msfExse.TextMatrix(Row, ExseCol.当前价)) * (Val(Me.msfExse.TextMatrix(Row, ExseCol.对应数)) - Val(strInputed))
            txtTotal = IIf(Val(txtTotal) = 0, "", Format(txtTotal, "0.0000"))
        Case ExseCol.收费方式
            msfExse.TextMatrix(Row, Col) = msfExse.CboText
    End Select
End Sub

Private Sub optPreproty_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub stbExse_Click(PreviousTab As Integer)
    msfExse.Visible = stbExse.Tab = 0
    vfgExse.Visible = stbExse.Tab = 1
    vfg加收.Visible = stbExse.Tab = 2
End Sub

Private Sub tabDept_Click()
    Call ResizeTabDept
    Call lstDeptSelect(1)
End Sub

Private Sub txtDept_GotFocus()
    Call zlControl.TxtSelAll(txtDept)
End Sub

Private Sub txtDept_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call DeptSelect(txtDept.Text)
    End If
End Sub

Private Sub txtItem_GotFocus()
    Me.txtItem.SelStart = 0: Me.txtItem.SelLength = 100
End Sub

Private Sub txtItem_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    strTemp = UCase(Trim(Me.txtItem.Text))
    If strTemp = "" Then Me.lblItem.Tag = 0: Me.txtItem.Tag = "": Me.txtItem.Text = "": Exit Sub
    
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    err = 0: On Error GoTo ErrHand
    
    gstrSql = "select distinct I.ID,I.编码,I.名称,I.计算单位,nvl(I.计价性质,0) as 计价性质,I.类别,I.操作类型 " & _
            " from 诊疗项目目录 I,诊疗项目别名 N" & _
            " where I.ID=N.诊疗项目ID and I.类别>='A'" & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (I.编码 like [1] or N.名称 like [2] or N.简码 like [2])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strTemp & "%", gstrMatch & strTemp & "%")
        
    With rsTemp
        If .BOF Or .EOF Then
            lblMessage.Caption = lblMessage.Tag & "未找到指定的诊疗项目，请重新指定"
            Me.lblItem.Tag = 0: Me.txtItem.Tag = "": Me.txtItem.Text = Me.txtItem.Tag: Me.txtItem.SetFocus: Exit Sub
        End If
        If .RecordCount = 1 Then
            If Me.lblItem.Tag <> !ID Then
                Me.lblItem.Tag = !ID: Me.txtItem.Tag = "[" & !编码 & "]" & !名称: Me.txtItem.Text = Me.txtItem.Tag
                Me.optPreproty(!计价性质).Value = True
                mstrType = !类别
                mstrOper = IIf(IsNull(!操作类型), "", !操作类型)
                Call zlExseRef(Me.lblItem.Tag)
            End If
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            'objItem.SubItems(Me.lvwItems.ColumnHeaders("规格").Index - 1) = IIf(IsNull(!规格), "", !规格)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwItems.ColumnHeaders("计算单位").Index - 1) = IIf(IsNull(!计算单位), "", !计算单位)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("类别").Index - 1) = !类别
            objItem.SubItems(Me.lvwItems.ColumnHeaders("操作类型").Index - 1) = IIf(IsNull(!操作类型), "", !操作类型)
            objItem.Tag = !计价性质
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .Tag = Me.txtItem.Name
        .Left = Me.txtItem.Left
        .Top = Me.txtItem.Top + Me.txtItem.Height
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtItem_LostFocus()
    Me.txtItem.Text = Me.txtItem.Tag
End Sub

Private Sub zlExseRef(lngItemID As Long)
    '--------------------------------------------------------
    '功能：刷新显示诊疗项目对应的收费项目
    '入参：lngItemId-指定的诊疗项目id
    '--------------------------------------------------------
    Dim dblTotal As Double
    Dim n As Integer
    
    err = 0: On Error GoTo ErrHand
    
    '根据服务对象显示可选择科室

    If mlngClient = 0 Then
        '0(空)-不直接应用于病人,1-门诊,2-住院,3-门诊和住院(全院),4-体检
        MsgBox "该项目不直接应用于病人！"
        Exit Sub
    ElseIf mlngClient = 1 Then
        tabDept.Tabs.Clear
        tabDept.Tabs.Add 1, "ALL", "所有科室"
        tabDept.Tabs.Add 2, "MZ", "门诊科室"
    ElseIf mlngClient = 2 Then
        tabDept.Tabs.Clear
        tabDept.Tabs.Add 1, "ALL", "所有科室"
        tabDept.Tabs.Add 2, "ZY", "住院科室"
    ElseIf mlngClient = 3 Then
        tabDept.Tabs.Clear
        tabDept.Tabs.Add 1, "ALL", "所有科室"
        tabDept.Tabs.Add 2, "MZ", "门诊科室"
        tabDept.Tabs.Add 3, "ZY", "住院科室"
    ElseIf mlngClient = 4 Then
        tabDept.Tabs.Clear
        tabDept.Tabs.Add 1, "ALL", "所有科室"
        tabDept.Tabs.Add 2, "TJ", "体检科室"
    End If
    
    If tabDept.SelectedItem Is Nothing Then
        tabDept.SelectedItem = tabDept.Tabs("ALL")
    End If
    
    Call ResizeTabDept
    
    '读取所有数据
    Call ReadClinicData(lngItemID, mstrType, mlngFlag)
    '根据选择的科室显示数据
    Call lstDeptSelect(0)
    
    If mstrOper = "病理" And mstr类别 = "D" Then
        stbExse.TabVisible(1) = False
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lstDeptSelect(ByVal lngCache As Long)
        '根据选择的科室显示数据
        '
        ' lngCache : 0-初始调用 <>0 界面切换，当<>0时，需要保存原界面上的数据。
        '
        Dim strDept() As String '本次要显示的科室
        Dim curDept() As String '显示前要缓存的科室
        Dim lngListIndex As Long
        Dim lngRow As Long
        Dim lngSource As Long '当前病人来源
        On Error GoTo hErr

    
100     ReDim strDept(0) As String
102     ReDim curDept(0) As String
104     If lngCache <> 0 Then
            '缓存科室
106         If lstDept.ListCount > 0 Then
108             For lngRow = 0 To lstDept.ListCount - 1
110                 If curDept(UBound(curDept)) <> "" Then ReDim Preserve curDept(UBound(curDept) + 1)
112                 curDept(UBound(curDept)) = lstDept.ItemData(lngRow) & "|" & _
                                               Replace(Split(lstDept.List(lngRow), "(")(1), ")", "") & "|" & _
                                               Split(lstDept.List(lngRow), "(")(0)
                Next
            End If
            
114         If mlngLastSource = 1 Then
116             mDept1 = curDept
118         ElseIf mlngLastSource = 2 Then
120             mDept2 = curDept
122         ElseIf mlngLastSource = 3 Then
124             mDept3 = curDept
            End If
        End If
        '根据当前选定页面，选择数据
        lngSource = GetCurrSource
126     If lngSource = 0 Then
            '显示 当前页面数据
128         Call msfExseRef(0, lngCache)
            Exit Sub
130     ElseIf lngSource = 1 Then
132         strDept = mDept1
134     ElseIf lngSource = 2 Then
136         strDept = mDept2
138     ElseIf lngSource = 3 Then
140         strDept = mDept3
        End If
142     txtDept.Text = ""
        '添加科室
144     lngListIndex = -1
146     lstDept.Clear
148     For lngRow = LBound(strDept) To UBound(strDept)
150         If strDept(lngRow) <> "" Then
                'strDept格式： 适用科室ID | 科室编码  |  科室名称
152             lstDept.AddItem CStr(Split(strDept(lngRow), "|")(2)) & "(" & Split(strDept(lngRow), "|")(1) & ")"
154             lstDept.ItemData(lstDept.NewIndex) = Val(Split(strDept(lngRow), "|")(0))
156             If mlngLastDeptID <> 0 And lstDept.ItemData(lstDept.NewIndex) = mlngLastDeptID Then
158                 lngListIndex = lstDept.NewIndex
                End If
            End If
        Next
160     If lngListIndex <> -1 Then
162         lstDept.ListIndex = lngListIndex
164     ElseIf lstDept.ListCount > 0 Then
166        lstDept.ListIndex = 0
        Else
            '无科室，清空控件中显示的对照数据
168         Call msfExseRef(-1, lngCache)
        End If
        Exit Sub
hErr:
170     MsgBox "lstDeptSelect第" & CStr(Erl()) & "行：" & err.Description
    '    If ErrCenter = 1 Then
    '        Resume
    '    End If
End Sub

Private Sub ReadClinicData(ByVal lngItemID As Long, strType As String, lngFlag As Long)
        '读取诊疗项目数据，并缓存到本地变量。
        'lngItemID:诊疗项目.ID
        'strType  :诊疗项目.类别
        'lngFlag  :诊疗项目.执行标记
    
        Dim strSql As String, rsCharge As ADODB.Recordset
        Dim strTmp As String, strTmpDept As String, strDeptList As String
        Dim str收费方式 As String
100     err = 0: On Error GoTo ErrHand
    
102     ReDim mDept1(0) As String: ReDim mDept2(0) As String: ReDim mDept3(0) As String
104     ReDim mGen0(0) As String: ReDim mGen1(0) As String: ReDim mGen2(0) As String: ReDim mGen3(0) As String
106     ReDim mPlace0(0) As String: ReDim mPlace1(0) As String: ReDim mPlace2(0) As String: ReDim mPlace3(0) As String
108     ReDim mAppend0(0) As String: ReDim mAppend1(0) As String: ReDim mAppend2(0) As String: ReDim mAppend3(0) As String
    
        '普通的收费对照
110     strSql = "Select i.Id, '[' || i.编码 || ']' || i.名称 As 名称, i.规格, i.计算单位, Decode(i.是否变价, 1, '变价', To_Char(Sum(p.现价))) As 价格," & vbNewLine & _
            "       Nvl(r.收费数量, 0) As 数量, Nvl(r.固有对照, 0) As 固定, Nvl(r.从属项目, 0) As 从项, r.费用性质, Nvl(r.收费方式, 0) As 收费方式," & vbNewLine & _
            "       To_Number(r.病人来源) As 病人来源, r.适用科室id, b.编码 As 科室编码, b.名称 As 科室名称" & vbNewLine & _
            "From 诊疗收费关系 R, 收费项目目录 I, 收费价目 P, 部门表 B" & vbNewLine & _
            "Where r.收费项目id = i.Id And i.Id = p.收费细目id(+) And (r.费用性质 = 0 Or r.费用性质 Is Null) And r.检查部位 Is Null " & _
            "       And r.适用科室id = b.Id(+) And r.诊疗项目id = [1] " & vbNewLine & _
            "       And p.执行日期 <= Sysdate And (p.终止日期 Is Null Or p.终止日期 >= Sysdate)  And P.价格等级 Is Null " & _
            "Group By i.Id, i.编码, i.名称, i.规格, i.计算单位, i.是否变价, r.收费数量, r.固有对照, r.从属项目, r.费用性质, r.收费方式, r.病人来源, r.适用科室id, b.编码, b.名称 " & _
            "Order By r.病人来源, b.名称, Nvl(r.从属项目, 0)"

112     Set rsCharge = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngItemID, gstrPriceClass)
114     strDeptList = ""
116     Do Until rsCharge.EOF
118         Select Case Val("" & rsCharge!收费方式)
            Case 0: str收费方式 = "0-正常收取"
120         Case 1: str收费方式 = "1-检验试管费用"
122         Case 2: str收费方式 = "2-一次发送只收取一次"
124         Case 3: str收费方式 = "3-当天只收取一次"
126         Case 4: str收费方式 = "4-当天未执行收取一次"
128         Case 5: str收费方式 = "5-当天只收取一次，排斥其他项目"
130         Case 6: str收费方式 = "6-当天未执行收取一次，排斥其他项目"
            Case 7: str收费方式 = "7-每天首次不收取"
            Case 9: str收费方式 = "9-自定义"
132         Case Else
134             str收费方式 = ""
            End Select
        
136         strTmp = "" & rsCharge!ID & "|" & "" & rsCharge!名称 & "|" & rsCharge!规格 & "|" & rsCharge!计算单位 & "|" & _
                     rsCharge!价格 & "|" & rsCharge!数量 & "|" & IIf(rsCharge!固定 = 0, "", "√") & "|" & _
                     IIf(rsCharge!从项 = 0, "", "√") & "|" & str收费方式 & "|" & _
                     Val("" & rsCharge!病人来源) & "|" & Val("" & rsCharge!适用科室id)
            
138         If Val("" & rsCharge!适用科室id) <> 0 And InStr(strDeptList, "," & Val("" & rsCharge!适用科室id) & ":" & rsCharge!病人来源 & ",") <= 0 Then
140             strTmpDept = Val("" & rsCharge!适用科室id) & "|" & rsCharge!科室编码 & "|" & rsCharge!科室名称
142             strDeptList = strDeptList & "," & Val("" & rsCharge!适用科室id) & ":" & rsCharge!病人来源 & ","
            Else
144             strTmpDept = ""
            End If
            
146         If Val("" & rsCharge!病人来源) = 0 Then
                '全院
148             If mGen0(UBound(mGen0)) <> "" Then ReDim Preserve mGen0(UBound(mGen0) + 1)
150             mGen0(UBound(mGen0)) = strTmp
152         ElseIf Val("" & rsCharge!病人来源) = 1 Then
                '门诊
154             If mGen1(UBound(mGen1)) <> "" Then ReDim Preserve mGen1(UBound(mGen1) + 1)
156             mGen1(UBound(mGen1)) = strTmp
            
158             If strTmpDept <> "" Then
160                 If mDept1(UBound(mDept1)) <> "" Then ReDim Preserve mDept1(UBound(mDept1) + 1)
162                 mDept1(UBound(mDept1)) = strTmpDept
                End If
164         ElseIf Val("" & rsCharge!病人来源) = 2 Then
                '住院
166             If mGen2(UBound(mGen2)) <> "" Then ReDim Preserve mGen2(UBound(mGen2) + 1)
168             mGen2(UBound(mGen2)) = strTmp
170             If strTmpDept <> "" Then
172                 If mDept2(UBound(mDept2)) <> "" Then ReDim Preserve mDept2(UBound(mDept2) + 1)
174                 mDept2(UBound(mDept2)) = strTmpDept
                End If
176         ElseIf Val("" & rsCharge!病人来源) = 3 Then
                '体检
178             If mGen3(UBound(mGen3)) <> "" Then ReDim Preserve mGen3(UBound(mGen3) + 1)
180             mGen3(UBound(mGen3)) = strTmp
182             If strTmpDept <> "" Then
184                 If mDept3(UBound(mDept3)) <> "" Then ReDim Preserve mDept3(UBound(mDept3) + 1)
186                 mDept3(UBound(mDept3)) = strTmpDept
                End If
            End If
188         rsCharge.MoveNext
        Loop

        '有部位的收费对照
    
190     If strType = "D" Then
192         strSql = "Select /*+ RULE */" & vbNewLine & _
                    " i.Id As 收费id, '[' || i.编码 || ']' || i.名称 As 项目名, i.计算单位 As 单位," & vbNewLine & _
                    " Decode(i.是否变价, 1, '变价', To_Char(p.价格)) As 价格, Nvl(r.收费数量, 0) As 数量, Nvl(r.固有对照, 0) As 固定,Nvl(r.收费方式,0) as 收费方式," & vbNewLine & _
                    " Nvl(r.从属项目, 0) As 从项, r.费用性质 As 性质, d.分组, r.检查部位 As 部位, r.检查方法 As 方法, r.病人来源, r.适用科室id, b.编码 as 科室编码, b.名称  as 科室名称" & vbNewLine & _
                    "From 收费项目目录 i," & vbNewLine & _
                    "        (   Select p.收费细目id, Sum(p.现价) As 价格" & vbNewLine & _
                    "            From 收费价目 p" & vbNewLine & _
                    "            Where p.执行日期 <= Sysdate And (p.终止日期 Is Null Or p.终止日期 >= Sysdate)  And p.价格等级 Is Null " & _
                    "            Group By p.收费细目id) p, 诊疗项目目录 c, 诊疗检查部位 d, 诊疗收费关系 r, 部门表 B" & vbNewLine & _
                    "Where r.收费项目id = i.Id And i.Id = p.收费细目id(+) And c.操作类型 = d.类型 And r.检查部位 = d.名称 And" & vbNewLine & _
                    "      r.适用科室id=b.id(+) And r.诊疗项目id = c.Id And r.检查部位 Is Not Null And r.诊疗项目id = [1]" & vbNewLine & _
                    "Order By d.分组, r.检查部位, r.检查方法"
         
194         Set rsCharge = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngItemID, gstrPriceClass)
196         Do Until rsCharge.EOF
198             strTmp = "" & rsCharge!分组 & "|" & rsCharge!部位 & "|" & rsCharge!方法 & "|" & rsCharge!项目名 & "|" & _
                        rsCharge!单位 & "|" & rsCharge!价格 & "|" & rsCharge!数量 & "|" & _
                        IIf("" & rsCharge!固定 = "1", "√", "") & "|" & rsCharge!收费方式 & "|" & rsCharge!性质 & "|" & rsCharge!收费ID & "|" & _
                        Val("" & rsCharge!病人来源) & "|" & Val("" & rsCharge!适用科室id)
            
            
200             If Val("" & rsCharge!适用科室id) <> 0 And InStr(strDeptList, "," & Val("" & rsCharge!适用科室id) & ":" & rsCharge!病人来源 & ",") <= 0 Then
202                 strTmpDept = Val("" & rsCharge!适用科室id) & "|" & rsCharge!科室编码 & "|" & rsCharge!科室名称
204                 strDeptList = strDeptList & "," & Val("" & rsCharge!适用科室id) & ":" & rsCharge!病人来源 & ","
                Else
206                 strTmpDept = ""
                End If
            
208             If Val("" & rsCharge!病人来源) = 0 Then
210                 If mPlace0(UBound(mPlace0)) <> "" Then ReDim Preserve mPlace0(UBound(mPlace0) + 1)
212                 mPlace0(UBound(mPlace0)) = strTmp

214             ElseIf Val("" & rsCharge!病人来源) = 1 Then
216                 If mPlace1(UBound(mPlace1)) <> "" Then ReDim Preserve mPlace1(UBound(mPlace1) + 1)
218                 mPlace1(UBound(mPlace1)) = strTmp
220                 If strTmpDept <> "" Then
222                     If mDept1(UBound(mDept1)) <> "" Then ReDim Preserve mDept1(UBound(mDept1) + 1)
224                     mDept1(UBound(mDept1)) = strTmpDept
                    End If
226             ElseIf Val("" & rsCharge!病人来源) = 2 Then
228                 If mPlace2(UBound(mPlace2)) <> "" Then ReDim Preserve mPlace2(UBound(mPlace2) + 1)
230                 mPlace2(UBound(mPlace2)) = strTmp
232                 If strTmpDept <> "" Then
234                     If mDept2(UBound(mDept2)) <> "" Then ReDim Preserve mDept2(UBound(mDept2) + 1)
236                     mDept2(UBound(mDept2)) = strTmpDept
                    End If
238             ElseIf Val("" & rsCharge!病人来源) = 3 Then
240                 If mPlace3(UBound(mPlace3)) <> "" Then ReDim Preserve mPlace3(UBound(mPlace3) + 1)
242                 mPlace3(UBound(mPlace3)) = strTmp
244                 If strTmpDept <> "" Then
246                     If mDept3(UBound(mDept3)) <> "" Then ReDim Preserve mDept3(UBound(mDept3) + 1)
248                     mDept3(UBound(mDept3)) = strTmpDept
                    End If
                End If
250             rsCharge.MoveNext
            Loop
        
252         If lngFlag = 1 Then
                '附加的收费对照
254             strSql = "Select /*+ RULE */" & vbNewLine & _
                        " i.Id as 收费ID, '[' || i.编码 || ']' || i.名称 As 项目名, i.计算单位 as 单位, Decode(i.是否变价, 1, '变价', To_Char(p.价格)) As 价格," & vbNewLine & _
                        " Nvl(r.收费数量, 0) As 数量, Nvl(r.固有对照, 0) As 固定, Nvl(r.从属项目, 0) As 从项, r.费用性质 as 性质, r.检查部位 as 部位, r.检查方法 as 方法, r.病人来源, r.适用科室id, b.编码 as 科室编码, b.名称  as 科室名称" & vbNewLine & _
                        "From 诊疗收费关系 r, 收费项目目录 i," & vbNewLine & _
                        "        (Select p.收费细目id, Sum(p.现价) As 价格" & vbNewLine & _
                        "            From 收费价目 p" & vbNewLine & _
                        "            Where p.执行日期 <= Sysdate And (p.终止日期 Is Null Or p.终止日期 >= Sysdate) And p.价格等级 Is Null  " & _
                        "            Group By p.收费细目id) p, 部门表 B" & vbNewLine & _
                        "Where r.收费项目id = i.Id And i.Id = p.收费细目id(+) And r.费用性质=1 And r.适用科室id=b.id(+) And r.诊疗项目id = [1]"
        
256             Set rsCharge = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngItemID, gstrPriceClass)
258             Do Until rsCharge.EOF
260                 strTmp = "" & rsCharge!项目名 & "|" & rsCharge!单位 & "|" & rsCharge!价格 & "|" & rsCharge!数量 & "|" & _
                             IIf("" & rsCharge!固定 = "1", "√", "") & "|" & rsCharge!性质 & "|" & rsCharge!收费ID & "|" & _
                             Val("" & rsCharge!病人来源) & "|" & Val("" & rsCharge!适用科室id)
                         
262                 If Val("" & rsCharge!适用科室id) <> 0 And InStr(strDeptList, "," & Val("" & rsCharge!适用科室id) & ":" & rsCharge!病人来源 & ",") <= 0 Then
264                     strTmpDept = Val("" & rsCharge!适用科室id) & "|" & rsCharge!科室编码 & "|" & rsCharge!科室名称
266                     strDeptList = strDeptList & "," & Val("" & rsCharge!适用科室id) & ":" & rsCharge!病人来源 & ","
                    Else
268                     strTmpDept = ""
                    End If
270                 If Val("" & rsCharge!病人来源) = 0 Then
272                     If mAppend0(UBound(mAppend0)) <> "" Then ReDim Preserve mAppend0(UBound(mAppend0) + 1)
274                     mAppend0(UBound(mAppend0)) = strTmp
276                 ElseIf Val("" & rsCharge!病人来源) = 1 Then
278                     If mAppend1(UBound(mAppend1)) <> "" Then ReDim Preserve mAppend1(UBound(mAppend1) + 1)
280                     mAppend1(UBound(mAppend1)) = strTmp
282                     If strTmpDept <> "" Then
284                         If mDept1(UBound(mDept1)) <> "" Then ReDim Preserve mDept1(UBound(mDept1) + 1)
286                         mDept1(UBound(mDept1)) = strTmpDept
                        End If
288                 ElseIf Val("" & rsCharge!病人来源) = 2 Then
290                     If mAppend2(UBound(mAppend2)) <> "" Then ReDim Preserve mAppend2(UBound(mAppend2) + 1)
292                     mAppend2(UBound(mAppend2)) = strTmp
294                     If strTmpDept <> "" Then
296                         If mDept2(UBound(mDept2)) <> "" Then ReDim Preserve mDept2(UBound(mDept2) + 1)
298                         mDept2(UBound(mDept2)) = strTmpDept
                        End If
300                 ElseIf Val("" & rsCharge!病人来源) = 3 Then
302                     If mAppend3(UBound(mAppend3)) <> "" Then ReDim Preserve mAppend3(UBound(mAppend3) + 1)
304                     mAppend3(UBound(mAppend3)) = strTmp
306                     If strTmpDept <> "" Then
308                         If mDept3(UBound(mDept3)) <> "" Then ReDim Preserve mDept3(UBound(mDept3) + 1)
310                         mDept3(UBound(mDept3)) = strTmpDept
                        End If
                    End If
312                 rsCharge.MoveNext
                Loop
            End If
        End If
 

        Exit Sub
ErrHand:
314     MsgBox "ReadClinicData第" & CStr(Erl()) & "行：" & err.Description
    '    If ErrCenter() = 1 Then
    '    Resume
    '    End If
    '    Call SaveErrLog
End Sub

Private Sub msfExseRef(ByVal lngDeptID As Long, ByVal lngCacheData As Long)
        '更新显示界面数据
        'lngDeptID : 科室ID，界面上显示这个科室的数据
        'lngCacheData: 是否缓存原数据
    
        Dim dblTotal As Double
        Dim lngRow As Long
        Dim strGen() As String '普通
        Dim strPlace() As String '部位
        Dim strAppend() As String '加收
        Dim lngListIndex As Long
    
100     err = 0: On Error GoTo ErrHand
102     If lngCacheData = 1 And mlngLastDeptID <> -1 Then
104         Call CacheData(mlngLastSource, mlngLastDeptID)
        End If
    
106     mlngLastSource = GetCurrSource
108     mlngLastDeptID = lngDeptID
    
        '根据当前选定页面，选择数据
110     ReDim strGen(0) As String
112     ReDim strPlace(0) As String
114     ReDim strAppend(0) As String
    
116     If mlngLastSource = 0 Then
118         strGen = mGen0
120         strPlace = mPlace0
122         strAppend = mAppend0
124     ElseIf mlngLastSource = 1 Then
126         strGen = mGen1
128         strPlace = mPlace1
130         strAppend = mAppend1
132     ElseIf mlngLastSource = 2 Then
134         strGen = mGen2
136         strPlace = mPlace2
138         strAppend = mAppend2
140     ElseIf mlngLastSource = 3 Then
142         strGen = mGen3
144         strPlace = mPlace3
146         strAppend = mAppend3
        End If
    
        '检查项目才可设置加收费
    
    
148     stbExse.TabVisible(1) = False
150     stbExse.TabVisible(2) = False
    
152     Call IniItemList
154     If lngDeptID = -1 Or Me.cmdClose.Tag = "查阅" Then
156         Me.msfExse.Active = False  '当前页面不是所有科室，并且无科室，不能编辑
        Else
158         Me.msfExse.Active = True
        End If
160     dblTotal = 0
162     For lngRow = LBound(strGen) To UBound(strGen)
164         If strGen(lngRow) <> "" Then
166             With Me.msfExse
                    'strGen格式为： id|名称|规格|计算单位|价格|数量|固定|从项|收费方式|病人来源|适用科室id
168                 If lngDeptID = 0 Or lngDeptID = Val(Split(strGen(lngRow), "|")(10)) Then
170                     If .RowData(.Rows - 1) <> 0 Then .Rows = .Rows + 1
172                     .TextMatrix(.Rows - 1, ExseCol.序号) = .Rows - 1
174                     .RowData(.Rows - 1) = Split(strGen(lngRow), "|")(0) 'ID
176                     .TextMatrix(.Rows - 1, ExseCol.项目名) = Split(strGen(lngRow), "|")(1)              '名称
178                     .TextMatrix(.Rows - 1, ExseCol.规格) = Split(strGen(lngRow), "|")(2)                '规格
180                     .TextMatrix(.Rows - 1, ExseCol.单位) = Split(strGen(lngRow), "|")(3)                '单位
182                     .TextMatrix(.Rows - 1, ExseCol.当前价) = Split(strGen(lngRow), "|")(4)              '价格
184                     .TextMatrix(.Rows - 1, ExseCol.对应数) = FormatEx(Split(strGen(lngRow), "|")(5), 5) '数量
186                     .TextMatrix(.Rows - 1, ExseCol.固定) = Split(strGen(lngRow), "|")(6)                '固定
188                     .TextMatrix(.Rows - 1, ExseCol.从项) = Split(strGen(lngRow), "|")(7)                '从项
190                     .TextMatrix(.Rows - 1, ExseCol.收费方式) = Split(strGen(lngRow), "|")(8)

                    
192                     dblTotal = dblTotal + Val(Split(strGen(lngRow), "|")(4)) * Val(Split(strGen(lngRow), "|")(5))
                    End If
                End With
            End If
        Next
194     txtTotal = IIf(dblTotal = 0, "", Format(dblTotal, "0.0000"))
    
196     If mstrType = "D" Then
198         stbExse.TabVisible(1) = True
            '初始化表格
200         Call vfgExseRef(strPlace, lngDeptID)
    
202         If mlngFlag = 1 Then
204             stbExse.TabVisible(2) = True
206             Call vfg加收Ref(strAppend, lngDeptID)
            End If
        End If
        
        If mstrOper = "病理" And mstr类别 = "D" Then
            stbExse.TabVisible(1) = False
        End If
        
        Exit Sub
ErrHand:
208     MsgBox "msfExseRef第" & CStr(Erl()) & "行：" & err.Description
    '    If ErrCenter() = 1 Then Resume
    '    Call SaveErrLog
End Sub

Private Sub vfgExseRef(ByRef strPla() As String, ByVal lngDeptID As Long)
        Dim strSql As String, dblTotal As Double
        Dim lngRow As Long
100     err = 0: On Error GoTo ErrHand
102     dblTotal = Val(txtTotal)
104     If lngDeptID = -1 Or Me.cmdClose.Tag = "查阅" Then
106         Me.vfgExse.Enabled = False '当前页面不是所有科室，并且无科室，不能编辑
        Else
108         Me.vfgExse.Enabled = True
        End If
110     With vfgExse
            '初始化表格
112         .Clear
114         .FixedCols = 0: .FixedRows = 1
116         .Rows = 1: .Cols = 11
        
118         .MergeRow(0) = True
120         .MergeCellsFixed = flexMergeRestrictColumns
    '
122         .MergeCol(0) = True ': .MergeCol(1) = True
124         .MergeCells = flexMergeRestrictColumns
        
126         .RowHeightMin = 300
        
128         .TextMatrix(0, 0) = "部位": .TextMatrix(0, 1) = "部位": .TextMatrix(0, 2) = "方法": .TextMatrix(0, 3) = "项目名"
130         .TextMatrix(0, 4) = "单位": .TextMatrix(0, 5) = "价格": .TextMatrix(0, 6) = "数量"
132         .TextMatrix(0, 7) = "固定": .TextMatrix(0, 8) = "收费方式": .TextMatrix(0, 9) = "性质": .TextMatrix(0, 10) = "收费ID":
134         .ColKey(0) = "分组": .ColKey(1) = "部位": .ColKey(2) = "方法": .ColKey(3) = "项目名"
136         .ColKey(4) = "单位": .ColKey(5) = "价格"
138         .ColKey(6) = "数量": .ColKey(7) = "固定": .ColKey(8) = "收费方式": .ColKey(9) = "性质": .ColKey(10) = "收费id"
        
140         .ColHidden(.ColIndex("分组")) = False
142         .ColHidden(.ColIndex("部位")) = False: .ColHidden(.ColIndex("方法")) = False: .ColHidden(.ColIndex("项目名")) = False
144         .ColHidden(.ColIndex("单位")) = False: .ColHidden(.ColIndex("价格")) = False: .ColHidden(.ColIndex("数量")) = False
146         .ColHidden(.ColIndex("固定")) = False: .ColHidden(.ColIndex("性质")) = True: .ColHidden(.ColIndex("收费id")) = True
            .ColHidden(.ColIndex("收费方式")) = False
148         .ColWidth(.ColIndex("分组")) = 900
150         .ColWidth(.ColIndex("部位")) = 1000: .ColWidth(.ColIndex("方法")) = 1000: .ColWidth(.ColIndex("项目名")) = 2200
152         .ColWidth(.ColIndex("单位")) = 450: .ColWidth(.ColIndex("价格")) = 800: .ColWidth(.ColIndex("数量")) = 800
154         .ColWidth(.ColIndex("固定")) = 450: .ColWidth(.ColIndex("性质")) = 0: .ColWidth(.ColIndex("收费id")) = 0
            .ColWidth(.ColIndex("收费方式")) = 1400
156         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
158         .WordWrap = True
160         .AutoResize = True
        
162         .ColComboList(.ColIndex("部位")) = "..."
164         .ColComboList(.ColIndex("方法")) = "..."
166         .ColComboList(.ColIndex("项目名")) = "..."
        
168         .Editable = flexEDKbdMouse
        
        
170         For lngRow = LBound(strPla) To UBound(strPla)
172             If strPla(lngRow) <> "" Then
174                 If lngDeptID = 0 Or lngDeptID = Val(Split(strPla(lngRow), "|")(12)) Then
176                     .Rows = .Rows + 1
                        'strPla格式：分组|部位|方法|项目名|单位|价格|数量|收费方式|固定|性质|收费ID|病人来源|适用科室id
178                     .TextMatrix(.Rows - 1, .ColIndex("分组")) = Split(strPla(lngRow), "|")(0)
180                     .TextMatrix(.Rows - 1, .ColIndex("部位")) = Split(strPla(lngRow), "|")(1)
182                     .TextMatrix(.Rows - 1, .ColIndex("方法")) = Split(strPla(lngRow), "|")(2)
184                     .TextMatrix(.Rows - 1, .ColIndex("项目名")) = Split(strPla(lngRow), "|")(3)
186                     .TextMatrix(.Rows - 1, .ColIndex("单位")) = Split(strPla(lngRow), "|")(4)
                    
188                     .TextMatrix(.Rows - 1, .ColIndex("价格")) = Format(Val(Split(strPla(lngRow), "|")(5)), "0.00")
190                     .TextMatrix(.Rows - 1, .ColIndex("数量")) = Val(Split(strPla(lngRow), "|")(6))
                    
192                     .TextMatrix(.Rows - 1, .ColIndex("固定")) = Split(strPla(lngRow), "|")(7)
                        .TextMatrix(.Rows - 1, .ColIndex("收费方式")) = IIf(0 = Val(Split(strPla(lngRow), "|")(8)), "0-正常收取", "9-自定义")
194                     .TextMatrix(.Rows - 1, .ColIndex("性质")) = Val(Split(strPla(lngRow), "|")(9))
196                     .TextMatrix(.Rows - 1, .ColIndex("收费ID")) = Val(Split(strPla(lngRow), "|")(10))
                    
198                     dblTotal = dblTotal + Val(Split(strPla(lngRow), "|")(5)) * Val(Split(strPla(lngRow), "|")(6))
                    End If
                End If
            Next
200         If .Rows < 2 Then .Rows = .Rows + 1
202         .AutoSizeMode = flexAutoSizeRowHeight
204         .AutoSize .ColIndex("分组"), .ColIndex("方法")
        End With
    
206     txtTotal = IIf(Val(dblTotal) = 0, "", Format(dblTotal, "0.00"))
        Exit Sub
ErrHand:
208     MsgBox "vfgExseRef第" & CStr(Erl()) & "行：" & err.Description
    '    If ErrCenter() = 1 Then Resume
    '    Call SaveErrLog
End Sub

Private Sub vfgExse_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If InStr("," & vfgExse.ColIndex("项目名") & ",", "," & Col & ",") > 0 Then
        Call vfgExse_CellButtonClick(Row, Col)
    End If
End Sub

Private Sub vfgExse_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vfgExse
        If NewCol = .ColIndex("收费方式") Then
            .ComboList = "0-正常收取|9-自定义"
        Else
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vfgExse_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vfgExse
        If InStr("," & .ColIndex("部位") & "," & .ColIndex("方法") & "," & .ColIndex("项目名") & "," & .ColIndex("数量") & "," & .ColIndex("收费方式") & ",", "," & Col & ",") <= 0 Then
            Cancel = True
        End If
    End With
End Sub

Private Sub vfgExse_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim pt As POINTAPI, strReturn As String
    Dim strSql As String, str部位 As String, bytType As Byte
    Dim rsTmp As ADODB.Recordset, varReturn As Variant, lngRow As Long, curOld金额 As Currency
    
    On Error GoTo ErrHandle
    With vfgExse
        '取得当前行的位置
        pt.x = .ColPos(Col) \ Screen.TwipsPerPixelX
        pt.y = (.RowPos(Row) + .RowHeight(Row)) \ Screen.TwipsPerPixelY
        ClientToScreen .hWnd, pt
        
        Select Case Col
            Case .ColIndex("部位")
'                strSQL = "Select Distinct '已选' As 分类, 部位" & vbNewLine & _
'                        "From 诊疗项目部位 a" & vbNewLine & _
'                        "Where 项目id = [1]" & vbNewLine & _
'                        "Union All" & vbNewLine & _
'                        "Select '可选' As 分类, 名称 As 部位" & vbNewLine & _
'                        "From 诊疗检查部位" & vbNewLine & _
'                        "Where 类型 = (Select 操作类型 From 诊疗项目目录 Where Id = [1]) And" & vbNewLine & _
'                        "           名称 Not In (Select 部位 From 诊疗项目部位 Where 项目id = [1])"
                strSql = "Select Distinct Decode(C.部位,Null,'可选','已选') As 分类,A.分组 as 类别, A.名称 As 部位" & vbNewLine & _
                        "From 诊疗检查部位 A,诊疗项目目录 B,(Select 类型,部位 From 诊疗项目部位 C Where 项目ID=[1]) C" & vbNewLine & _
                        "Where A.类型 = B.操作类型 And B.Id=[1] And A.类型=C.类型(+) And A.名称=C.部位(+) Order by 分类,类别"

                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(lblItem.Tag))
                bytType = 1
            Case .ColIndex("方法")
                str部位 = vfgExse.TextMatrix(Row, .ColIndex("部位"))
                If str部位 <> "" Then
                    strSql = "Select '已选' As 分类, 方法" & vbNewLine & _
                            "From 诊疗项目部位" & vbNewLine & _
                            "Where 部位 = [2] And 项目id = [1]" & vbNewLine & _
                            "Union All" & vbNewLine & _
                            "Select '可选' As 分类, 方法 As 部位" & vbNewLine & _
                            "From 诊疗检查部位" & vbNewLine & _
                            "Where 名称 = [2] And 类型 = (Select 操作类型 From 诊疗项目目录 Where Id = [1])"
                Else
                    MsgBox "请选择部位", vbInformation, gstrSysName
                    vfgExse.Select Row, .ColIndex("部位")
                    Exit Sub
                End If
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(lblItem.Tag), str部位)
                bytType = 2
            Case .ColIndex("项目名")
                strSql = "select distinct I.ID,Rpad('['||I.编码||']'||I.名称||' '||I.规格,60) as 名称,I.计算单位 as 单位" & _
                        " from 收费项目目录 I,收费项目别名 N" & _
                        " where I.ID=N.收费细目id and I.类别 not in ('1','J')" & _
                        "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
                        "       and (I.编码 like [1] " & _
                        "           or N.名称 like [2] " & _
                        "           or N.简码 like [2])"
                
                strTemp = UCase(.TextMatrix(.Row, .Col))
                If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strTemp & "%", gstrMatch & strTemp & "%")
                bytType = 3
        End Select
        If rsTmp.BOF Or rsTmp.EOF Then
            If bytType = 1 Or bytType = 2 Then
                Me.vfgExse.TextMatrix(Row, Col) = ""
                lblMessage.Caption = lblMessage.Tag & "未找到项目！"
            Else
                Me.vfgExse.TextMatrix(Row, Col) = ""
                lblMessage.Caption = lblMessage.Tag & "未找到指定收费项目！"
            End If
            Exit Sub
        End If
        
        frmClinicExseVsSelect.Move pt.x * Screen.TwipsPerPixelX, pt.y * Screen.TwipsPerPixelY
        
        Call frmClinicExseVsSelect.ShowSelect(bytType, rsTmp, strReturn)
        
        If InStr(strReturn, "|") > 0 Then
            varReturn = Split(strReturn, "|")
            If bytType = 1 Then
                '部位
                .TextMatrix(Row, Col - 1) = varReturn(1)
                .TextMatrix(Row, Col) = varReturn(2)
                .TextMatrix(Row, Col + 1) = ""
                .Select Row, Col + 1
            ElseIf bytType = 2 Then
                '方法
                .TextMatrix(Row, Col) = varReturn(1)
                '11295 要求一个方法可以对应多个收费项目
'                For lngRow = .FixedRows To .Rows - 1
'                    If lngRow <> Row And .TextMatrix(lngRow, Col) = varReturn(1) And _
'                       .TextMatrix(lngRow, .ColIndex("部位")) = .TextMatrix(Row, .ColIndex("部位")) Then
'                        lblMessage.Caption = lblMessage.Tag & "每种方法只能对应一个收费项目！"
'                        .TextMatrix(Row, Col) = ""
'                    End If
'                Next
                .Select Row, Col + 1
            Else
                '项目
                .TextMatrix(Row, .ColIndex("收费ID")) = Val("" & varReturn(0))
                .TextMatrix(Row, .ColIndex("项目名")) = "" & varReturn(1)
                .TextMatrix(Row, .ColIndex("单位")) = "" & varReturn(2)
                
                strSql = "select decode(I.是否变价,1,'变价',to_char(P.价格)) As 价格" & _
                " from (select 是否变价 from 收费项目目录 where id=[1]) I," & _
                "      (Select sum(现价) As 价格" & _
                "      From 收费价目 " & _
                "      Where 收费细目id=[1] " & _
                "           and 执行日期<=Sysdate And (终止日期 Is Null Or 终止日期>=Sysdate)  And 价格等级 Is Null " & ") P"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val("" & varReturn(0)), gstrPriceClass)
                
                curOld金额 = Val(.TextMatrix(Row, .ColIndex("价格"))) * Val(.TextMatrix(Row, .ColIndex("数量")))
                
                If rsTmp.RecordCount > 0 Then
                    If rsTmp.RecordCount > 1 Then
                        lblMessage.Caption = lblMessage.Tag & "主项的价格存在多个收入项目，不能选择。"
                        .TextMatrix(Row, .ColIndex("价格")) = ""
                    Else
                        .TextMatrix(Row, .ColIndex("价格")) = IIf("" & rsTmp.Fields("价格") <> "变价", Format(Val("" & rsTmp.Fields("价格")), "0.00"), "变价")
                    End If
                Else
                    .TextMatrix(Row, .ColIndex("价格")) = ""
                End If
                
                txtTotal = Val(txtTotal) - curOld金额 + Val(.TextMatrix(Row, .ColIndex("数量"))) * Val(.TextMatrix(Row, .ColIndex("价格")))
                If Val(txtTotal) = 0 Then
                    txtTotal = ""
                Else
                    txtTotal = Format(txtTotal, "0.0000")
                End If
                .Select Row, .ColIndex("数量")
            End If
            .AutoSize .ColIndex("分组"), .ColIndex("方法")
            
        End If
        
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vfgExse_DblClick()
    If vfgExse.Col = vfgExse.ColIndex("固定") Then
        Call vfgExse_KeyDown(vbKeySpace, 0)
    End If
End Sub

Private Sub vfgExse_EnterCell()
    Dim blnOk As Boolean
    
    With vfgExse
        If InStr("," & .ColIndex("部位") & "," & .ColIndex("方法") & "," & .ColIndex("项目名") & "," & .ColIndex("数量") & ",", "," & .Col & ",") > 0 Then
            blnOk = True
        End If
    End With
    On Error Resume Next
    If blnOk And vfgExse.Row > 0 Then
        Call vfgExse.CellBorder(vfgExse.GridColor, 1, 1, 2, 2, 0, 0)
    End If
End Sub

Private Sub vfgExse_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnCancle As Boolean
    
    On Error Resume Next
    
    If KeyCode = vbKeyDelete Then
        With vfgExse
        
            txtTotal = Val(txtTotal) - Val(.TextMatrix(.Row, .ColIndex("数量"))) * Val(.TextMatrix(.Row, .ColIndex("价格")))
            If Val(txtTotal) = 0 Then
                txtTotal = ""
            Else
                txtTotal = Format(txtTotal, "0.0000")
            End If
            If .Row > .FixedRows And .Row <= .Rows - 1 Then
                vfgExse.RemoveItem vfgExse.Row
            ElseIf .Row = .FixedRows Then
                .Cell(flexcpText, .Row, 0, .Row, .Cols - 1) = ""
            End If
        End With
    ElseIf KeyCode = vbKeyReturn Then
        
        With vfgExse
            If .EditText = "" Then
                KeyCode = 0
                If .Col = .ColIndex("收费方式") Then
                    If .Row = vfgExse.Rows - 1 Then
                        .Rows = .Rows + 1
                    End If
                    .Select .Row + 1, .ColIndex("部位")
                Else
                    If .Col < .Cols Then
                        If .Col = .ColIndex("项目名") Then
                            .Select .Row, .ColIndex("数量")
                        Else
                            .Select .Row, .Col + 1
                        End If
                    End If
                End If
            End If
        End With
    ElseIf KeyCode = vbKeySpace Then
        With vfgExse
        If .Col = .ColIndex("固定") Then
            If .TextMatrix(.Row, .Col) = "" Then
                .TextMatrix(.Row, .Col) = "√"
            Else
                .TextMatrix(.Row, .Col) = ""
            End If
        End If
        End With
    ElseIf KeyCode = vbKeyEscape Then
        With vfgExse
            If .Col = .ColIndex("项目名") Then
                If .ColComboList(.Col) <> "" Then
                    KeyCode = 0
                End If
            End If
        End With
    ElseIf InStr("," & vfgExse.ColIndex("项目名") & ",", "," & vfgExse.Col & ",") > 0 Then
        If vfgExse.ColComboList(vfgExse.Col) <> "" Then
            vfgExse.Tag = vfgExse.TextMatrix(vfgExse.Row, vfgExse.Col)
            vfgExse.ColComboList(vfgExse.Col) = ""
        End If
    End If
End Sub

Private Sub vfgExse_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim i As Integer, blnCancle As Boolean
    If InStr("," & vfgExse.ColIndex("项目名") & ",", "," & Col & ",") > 0 And KeyCode = vbKeyReturn Then
        vfgExse.ColComboList(vfgExse.Col) = "..."
        
    ElseIf vfgExse.ColIndex("数量") = vfgExse.Col And KeyCode <> vbKeyReturn Then
        If InStr("01234567890.", Chr(KeyCode)) <= 0 Then
            KeyCode = 0
        End If
    ElseIf KeyCode = vbKeyEscape Then
        vfgExse.TextMatrix(Row, Col) = vfgExse.Tag
        vfgExse.ColComboList(vfgExse.Col) = "..."
'    ElseIf vbKeyReturn Then
'        Call vfgExse_CellButtonClick(Row, Col)
    End If
End Sub

Private Sub vfgExse_LeaveCell()
    Dim blnOk As Boolean
    
    With vfgExse
        If InStr("," & .ColIndex("部位") & "," & .ColIndex("方法") & "," & .ColIndex("项目名") & "," & .ColIndex("数量") & ",", "," & .Col & ",") > 0 Then
            blnOk = True
        End If
    End With
    If InStr("," & vfgExse.ColIndex("部位") & "," & vfgExse.ColIndex("方法") & "," & vfgExse.ColIndex("项目名") & ",", "," & vfgExse.Col & ",") > 0 Then
        vfgExse.ColComboList(vfgExse.Col) = "..."
    End If
    On Error Resume Next
    If blnOk And vfgExse.Row > 0 Then
        Call vfgExse.CellBorder(vfgExse.GridColor, 0, 0, 0, 0, 0, 0)
    End If
    
End Sub

Private Sub vfgExse_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim str方法 As String
    
    If Col = vfgExse.ColIndex("数量") Then
        With vfgExse
            txtTotal = Val(txtTotal) - Val(.TextMatrix(Row, Col)) * Val(.TextMatrix(Row, .ColIndex("价格"))) _
                       + Val(.EditText) * Val(.TextMatrix(Row, .ColIndex("价格")))
        End With
        If Val(txtTotal) = 0 Then
            txtTotal = ""
        Else
            txtTotal = Format(txtTotal, "0.0000")
        End If
    End If
    
    If vfgExse.ColIndex("部位") = Col Then
        If vfgExse.TextMatrix(Row, Col) <> "" Then
            mrsBwff.Filter = " 部位='" & vfgExse.EditText & "'"
            If mrsBwff.RecordCount <= 0 Then
                lblMessage.Caption = lblMessage.Tag & "部位有误，请检查。"
                Cancel = True
            End If
        End If
    End If
    
    If vfgExse.ColIndex("方法") = Col Then
        If vfgExse.TextMatrix(Row, Col) <> "" Then
            mrsBwff.Filter = " 部位='" & vfgExse.TextMatrix(Row, vfgExse.ColIndex("部位")) & "'"
            If mrsBwff.RecordCount <= 0 Then
                lblMessage.Caption = lblMessage.Tag & "部位有误，检查。"
                Cancel = True
             Else
                str方法 = mrsBwff.Fields("方法")
                str方法 = Replace(str方法, vbTab, "|")
                str方法 = Replace(str方法, ",", "|")
                str方法 = Replace(str方法, ";", "|")
                str方法 = Replace(str方法, "0", "")
                str方法 = Replace(str方法, "1", "")
                
                If InStr("|" & str方法 & "|", "|" & vfgExse.EditText & "|") <= 0 Then
                    lblMessage.Caption = lblMessage.Tag & "方法有误，请检查。"
                    Cancel = True
                End If
            End If
        End If
    End If
End Sub

'----------------- 加收
Private Sub vfg加收Ref(ByRef strAppend() As String, ByVal lngDeptID As Long)
        Dim strSql As String, dblTotal As Double
        Dim lngRow As Long
100     err = 0: On Error GoTo ErrHand
102     If lngDeptID = -1 Or Me.cmdClose.Tag = "查阅" Then
104         Me.vfgExse.Enabled = False '当前页面不是所有科室，并且无科室，不能编辑
        Else
106         Me.vfgExse.Enabled = True
        End If
108     dblTotal = Val(txtTotal)
110     With vfg加收
            '初始化表格
112         .Clear
114         .FixedCols = 0: .FixedRows = 1
116         .Rows = 1: .Cols = 7
        
118         .RowHeightMin = 300
        
120         .TextMatrix(0, 0) = "项目名": .TextMatrix(0, 1) = "单位": .TextMatrix(0, 2) = "价格": .TextMatrix(0, 3) = "数量"
122         .TextMatrix(0, 4) = "固定": .TextMatrix(0, 5) = "性质": .TextMatrix(0, 6) = "收费ID"
124         .ColKey(0) = "项目名": .ColKey(1) = "单位": .ColKey(2) = "价格"
126         .ColKey(3) = "数量": .ColKey(4) = "固定": .ColKey(5) = "性质": .ColKey(6) = "收费id"
        
128         .ColHidden(.ColIndex("项目名")) = False
130         .ColHidden(.ColIndex("单位")) = False: .ColHidden(.ColIndex("价格")) = False: .ColHidden(.ColIndex("数量")) = False
132         .ColHidden(.ColIndex("固定")) = False: .ColHidden(.ColIndex("性质")) = True: .ColHidden(.ColIndex("收费id")) = True
        
134         .ColWidth(.ColIndex("项目名")) = 5000
136         .ColWidth(.ColIndex("单位")) = 450: .ColWidth(.ColIndex("价格")) = 800: .ColWidth(.ColIndex("数量")) = 800
138         .ColWidth(.ColIndex("固定")) = 450: .ColWidth(.ColIndex("性质")) = 0: .ColWidth(.ColIndex("收费id")) = 0
        
140         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
142         .WordWrap = True
144         .AutoResize = True
        
146         .ColComboList(.ColIndex("项目名")) = "..."
        
148         .Editable = flexEDKbdMouse
        
150         For lngRow = LBound(strAppend) To UBound(strAppend)
152             If strAppend(lngRow) <> "" Then
154                 If lngDeptID = 0 Or lngDeptID = Val(Split(strAppend(lngRow), "|")(8)) Then
156                     .Rows = .Rows + 1
                        'strAppend格式：项目名|单位|价格|数量|固定|性质|收费ID|病人来源|适用科室id
158                     .TextMatrix(.Rows - 1, .ColIndex("项目名")) = Split(strAppend(lngRow), "|")(0)
160                     .TextMatrix(.Rows - 1, .ColIndex("单位")) = Split(strAppend(lngRow), "|")(1)
                    
162                     .TextMatrix(.Rows - 1, .ColIndex("价格")) = Format(Val(Split(strAppend(lngRow), "|")(2)), "0.00")
164                     .TextMatrix(.Rows - 1, .ColIndex("数量")) = Val(Split(strAppend(lngRow), "|")(3))
                    
166                     .TextMatrix(.Rows - 1, .ColIndex("固定")) = Split(strAppend(lngRow), "|")(4)
168                     .TextMatrix(.Rows - 1, .ColIndex("性质")) = Val(Split(strAppend(lngRow), "|")(5))
170                     .TextMatrix(.Rows - 1, .ColIndex("收费ID")) = Val(Split(strAppend(lngRow), "|")(6))
                    
172                     dblTotal = dblTotal + Val(Split(strAppend(lngRow), "|")(2)) * Val(Split(strAppend(lngRow), "|")(3))
                    End If
                End If
            Next
174         If .Rows < 2 Then .Rows = .Rows + 1
176         .AutoSizeMode = flexAutoSizeRowHeight
178         .AutoSize .ColIndex("项目名")
        End With
180         txtTotal = IIf(Val(dblTotal) = 0, "", Format(dblTotal, "0.00"))
        Exit Sub
ErrHand:
182     MsgBox "vfg加收Ref第" & CStr(Erl()) & "行：" & err.Description
    '    If ErrCenter() = 1 Then Resume
    '    Call SaveErrLog
End Sub

Private Sub vfg加收_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If vfg加收.ColIndex("项目名") = Col Then
        Call vfg加收_CellButtonClick(Row, Col)
    End If
End Sub

Private Sub vfg加收_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vfg加收
        If InStr("," & .ColIndex("项目名") & "," & .ColIndex("数量") & ",", "," & Col & ",") <= 0 Then
            Cancel = True
        End If
    End With
End Sub

Private Sub vfg加收_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim pt As POINTAPI, strReturn As String
    Dim strSql As String, curOld金额 As Currency
    Dim rsTmp As ADODB.Recordset, varReturn As Variant, lngRow As Long
    
    On Error GoTo ErrHandle
    With vfg加收
        '取得当前行的位置
        pt.x = .ColPos(Col) \ Screen.TwipsPerPixelX
        pt.y = (.RowPos(Row) + .RowHeight(Row)) \ Screen.TwipsPerPixelY
        ClientToScreen .hWnd, pt
        
        If Col = .ColIndex("项目名") Then
            strSql = "select distinct I.ID,Rpad('['||I.编码||']'||I.名称||' '||I.规格,60) as 名称,I.计算单位 as 单位" & _
                    " from 收费项目目录 I,收费项目别名 N" & _
                    " where I.ID=N.收费细目id and I.类别 not in ('1','J')" & _
                    "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
                    "       and (I.编码 like [1] " & _
                    "           or N.名称 like [2] " & _
                    "           or N.简码 like [2])"
            strTemp = UCase(.TextMatrix(.Row, .Col))
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strTemp & "%", gstrMatch & strTemp & "%")
        End If
        
        If rsTmp.BOF Or rsTmp.EOF Then
            .TextMatrix(Row, Col) = ""
            lblMessage.Caption = lblMessage.Tag & "未找到指定收费项目！"
            Exit Sub
        End If
        
        frmClinicExseVsSelect.Move pt.x * Screen.TwipsPerPixelX, pt.y * Screen.TwipsPerPixelY
        
        Call frmClinicExseVsSelect.ShowSelect(3, rsTmp, strReturn)
        
        If InStr(strReturn, "|") > 0 Then
            varReturn = Split(strReturn, "|")
            '项目
            .TextMatrix(Row, .ColIndex("收费ID")) = Val("" & varReturn(0))
            .TextMatrix(Row, .ColIndex("项目名")) = "" & varReturn(1)
            .TextMatrix(Row, .ColIndex("单位")) = "" & varReturn(2)
            
            strSql = "select decode(I.是否变价,1,'变价',to_char(P.价格)) As 价格" & _
            " from (select 是否变价 from 收费项目目录 where id=[1]) I," & _
            "      (Select sum(现价) As 价格" & _
            "      From 收费价目 " & _
            "      Where 收费细目id=[1] " & _
            "           and 执行日期<=Sysdate And (终止日期 Is Null Or 终止日期>=Sysdate)  And 价格等级 Is Null " & ") P"
            
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val("" & varReturn(0)), gstrPriceClass)
            curOld金额 = Val(.TextMatrix(Row, .ColIndex("价格"))) * Val(.TextMatrix(Row, .ColIndex("数量")))
            If rsTmp.RecordCount > 0 Then
                If rsTmp.RecordCount > 1 Then
                    lblMessage.Caption = lblMessage.Tag & "主项的价格存在多个收入项目，不能选择。"
                    .TextMatrix(Row, .ColIndex("价格")) = ""
                Else
                    .TextMatrix(Row, .ColIndex("价格")) = IIf("" & rsTmp.Fields("价格") <> "变价", Format(Val("" & rsTmp.Fields("价格")), "0.00"), "变价")
                End If
            Else
                .TextMatrix(Row, .ColIndex("价格")) = ""
            End If
            
            txtTotal = Val(txtTotal) - curOld金额 + Val(.TextMatrix(Row, .ColIndex("数量"))) * Val(.TextMatrix(Row, .ColIndex("价格")))
            If Val(txtTotal) = 0 Then
                txtTotal = ""
            Else
                txtTotal = Format(txtTotal, "0.0000")
            End If
            
            .Select Row, .ColIndex("数量")
        End If
        
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vfg加收_DblClick()
    If vfg加收.Col = vfg加收.ColIndex("固定") Then
        Call vfg加收_KeyDown(vbKeySpace, 0)
    End If
End Sub

Private Sub vfg加收_EnterCell()
    Dim blnOk As Boolean
    
    With vfg加收
        If InStr("," & .ColIndex("项目名") & "," & .ColIndex("数量") & ",", "," & .Col & ",") > 0 Then
            blnOk = True
        End If
    End With
    On Error Resume Next
    If blnOk And vfg加收.Row > 0 Then
        Call vfg加收.CellBorder(vfg加收.GridColor, 1, 1, 2, 2, 0, 0)
    End If
End Sub

Private Sub vfg加收_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnCancle As Boolean
    If KeyCode = vbKeyDelete Then
        With vfg加收
        
            txtTotal = Val(txtTotal) - Val(.TextMatrix(.Row, .ColIndex("数量"))) * Val(.TextMatrix(.Row, .ColIndex("价格")))
            If Val(txtTotal) = 0 Then
                txtTotal = ""
            Else
                txtTotal = Format(txtTotal, "0.0000")
            End If
            If .Row > 1 And .Row < .Rows - 1 Then
                .RemoveItem .Row
            Else
                .Cell(flexcpText, .Row, 0, .Row, .Cols - 1) = ""
            End If
        End With
    ElseIf KeyCode = vbKeyReturn Then
        
        With vfg加收
            If .EditText = "" Then
                KeyCode = 0
                If .Col = .ColIndex("固定") Then
                    If .Row = .Rows - 1 Then
                        .Rows = .Rows + 1
                    End If
                    .Select .Row + 1, .ColIndex("项目名")
                Else
                    If .Col < .Cols Then
                        If .Col = .ColIndex("项目名") Then
                            .Select .Row, .ColIndex("数量")
                        Else
                            .Select .Row, .Col + 1
                        End If
                    End If
                End If
            End If
        End With
    ElseIf KeyCode = vbKeySpace Then
        With vfg加收
        If .Col = .ColIndex("固定") Then
            If .TextMatrix(.Row, .Col) = "" Then
                .TextMatrix(.Row, .Col) = "√"
            Else
                .TextMatrix(.Row, .Col) = ""
            End If
        End If
        End With
    ElseIf KeyCode = vbKeyEscape Then
        With vfg加收
            If .Col = .ColIndex("项目名") Then
                If .ColComboList(.Col) <> "" Then
                    KeyCode = 0
                End If
            End If
        End With
    ElseIf vfg加收.ColIndex("项目名") = vfg加收.Col Then
        If vfg加收.ColComboList(vfg加收.Col) <> "" Then
            vfg加收.ColComboList(vfg加收.Col) = ""
        End If
    End If
End Sub

Private Sub vfg加收_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim i As Integer, blnCancle As Boolean
    If vfg加收.ColIndex("项目名") = Col And KeyCode = vbKeyReturn Then
        vfg加收.ColComboList(vfg加收.Col) = "..."
    
    ElseIf vfg加收.ColIndex("数量") = vfg加收.Col And KeyCode <> vbKeyReturn Then
        If InStr("01234567890.", Chr(KeyCode)) <= 0 Then
            KeyCode = 0
        End If
    ElseIf KeyCode = vbKeyEscape Then
        vfg加收.TextMatrix(Row, Col) = vfg加收.Tag
        vfg加收.ColComboList(vfg加收.Col) = "..."
    End If
End Sub

Private Sub vfg加收_LeaveCell()
    Dim blnOk As Boolean
    
    With vfg加收
        If InStr("," & .ColIndex("项目名") & "," & .ColIndex("数量") & ",", "," & .Col & ",") > 0 Then
            blnOk = True
        End If
    End With
    If vfg加收.ColIndex("项目名") = vfg加收.Col Then
        vfg加收.ColComboList(vfg加收.Col) = "..."
    End If
    On Error Resume Next
    If blnOk And vfg加收.Row > 0 Then
        Call vfg加收.CellBorder(vfg加收.GridColor, 0, 0, 0, 0, 0, 0)
    End If
    
End Sub


Private Sub vfg加收_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim str方法 As String
    
    If Col = vfg加收.ColIndex("数量") Then
        With vfg加收
            txtTotal = Val(txtTotal) - Val(.TextMatrix(Row, Col)) * Val(.TextMatrix(Row, .ColIndex("价格"))) _
                       + Val(.EditText) * Val(.TextMatrix(Row, .ColIndex("价格")))
        End With
        If Val(txtTotal) = 0 Then
            txtTotal = ""
        Else
            txtTotal = Format(txtTotal, "0.0000")
        End If
    End If
    
End Sub

Private Sub ResizeTabDept()
    '调整stbExse控件的大小
    On Error Resume Next
    With tabDept
        If .SelectedItem.Index = 1 Then
            '不显示科室
            fraDept.Visible = False
            stbExse.Left = .Left + 90
            stbExse.Top = .Top + 400
            
            stbExse.Width = .Width - 180
            stbExse.Height = .Height - 500
        Else
            fraDept.Visible = True
            
            fraDept.Left = .Left + 90
            fraDept.Top = .Top + 400
            fraDept.Height = .Height - 500
            
            stbExse.Left = fraDept.Left + fraDept.Width + 45
            stbExse.Width = .Width - fraDept.Width - 230
            stbExse.Top = fraDept.Top
            stbExse.Height = fraDept.Height
        End If
        ResizeStbExse
    End With
End Sub

Private Sub ResizeStbExse()
    On Error Resume Next
    With stbExse
            msfExse.Left = 90
            msfExse.Top = 325
            msfExse.Width = .Width - 180
            msfExse.Height = .Height - 410
            
            vfgExse.Left = 90
            vfgExse.Top = 325
            vfgExse.Width = .Width - 180
            vfgExse.Height = .Height - 410
            
            vfg加收.Left = 90
            vfg加收.Top = 325
            vfg加收.Width = .Width - 180
            vfg加收.Height = .Height - 410
    End With
End Sub

Private Sub DeptSelect(ByVal strInput As String)
    '选择科室
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim strDeptList As String, strReturn As String
    Dim strInquiry As String, i As Integer, lngSource As Long
    On Error GoTo hErr
        lngSource = GetCurrSource
        If Trim(strInput) <> "" Then
            strInquiry = gstrMatch & UCase(strInput) & "%"
        End If
        If lstDept.ListCount > 0 Then
            For i = 0 To lstDept.ListCount - 1
                strDeptList = strDeptList & "," & lstDept.ItemData(i)
            Next
        End If
        If lngSource = 1 Or lngSource = 3 Then
            '门诊和体检科室
            strSql = "Select Distinct a.编码, a.名称, a.ID" & vbNewLine & _
                    "From 部门表 A, 部门性质说明 D" & vbNewLine & _
                    "Where a.Id = d.部门id And (d.服务对象 = 1 Or d.服务对象 = 3) and d.工作性质 in ('临床','检查','检验','手术','麻醉','治疗','营养') " & vbNewLine & _
                    " And (A.撤档时间 is null or A.撤档时间=to_date('3000-01-01','YYYY-MM-DD')) " & vbNewLine & _
                    IIf(strDeptList = "", "", "   And Instr([1], ',' || a.Id || ',') <= 0") & vbNewLine & _
                    IIf(strInquiry = "", "", " And (a.编码 Like [2] Or a.名称 Like [2] Or a.简码 Like [2]) ") & _
                    "Order By 编码, 名称"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strDeptList & ",", strInquiry)
        ElseIf lngSource = 2 Then
            '住院科室
            strSql = "Select Distinct a.编码, a.名称, a.ID" & vbNewLine & _
                    "From 部门表 A, 部门性质说明 D" & vbNewLine & _
                    "Where a.Id = d.部门id And (d.服务对象 = 2 Or d.服务对象 = 3) And d.工作性质 in ('护理','检查','检验','手术','麻醉','治疗','营养') " & vbNewLine & _
                    " And (A.撤档时间 is null or A.撤档时间=to_date('3000-01-01','YYYY-MM-DD')) " & vbNewLine & _
                    IIf(strDeptList = "", "", "   And Instr([1], ',' || a.Id || ',') <= 0") & vbNewLine & _
                    IIf(strInquiry = "", "", " And (a.编码 Like [2] Or a.名称 Like [2] Or a.简码 Like [2]) ") & _
                    "Order By 编码, 名称"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strDeptList & ",", strInquiry)
        End If
        
        strReturn = ""
        If Not rsTmp.EOF Then
            If rsTmp.RecordCount > 0 Then
                strReturn = frmSelCur.ShowCurrSel(Me, rsTmp, "编码,1200,0,2;名称,1800,0,2;ID,0,1,2", "选择科室", True, , , 3000 + 2000)
            Else
                strReturn = rsTmp!ID
            End If
        End If
        
        If strReturn <> "" Then
            txtDept = Split(strReturn, ",")(1) & "(" & Split(strReturn, ",")(0) & ")"
            
            lstDept.AddItem Split(strReturn, ",")(1) & "(" & Split(strReturn, ",")(0) & ")"
            lstDept.ItemData(lstDept.NewIndex) = Split(strReturn, ",")(2)
            lstDept.ListIndex = lstDept.NewIndex
            txtDept.Text = ""
        Else
            If rsTmp.RecordCount = 0 Then
                MsgBox "没有找到可用的科室。", vbInformation, Me.Caption
                Call zlControl.TxtSelAll(txtDept)
            End If
        End If

    Exit Sub
hErr:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Function CacheData(ByVal lngSource As Long, ByVal lngDeptID As Long) As Boolean
        '保存当前界面上的数据到缓存
        Dim blnErr As Boolean
        Dim strGen() As String
        Dim strPlan() As String
        Dim strAppend() As String
    
100     ReDim strGen(0) As String
102     ReDim strPlan(0) As String
104     ReDim strAppend(0) As String
        On Error GoTo hErr
    
106     CacheData = False

108     With Me.msfExse
110         blnErr = False
112         For intCount = 1 To .Rows - 1
114             If Trim(.TextMatrix(intCount, ExseCol.项目名)) <> "" And .RowData(intCount) <> 0 Then
116                 If Not IsNumeric(Nvl(.TextMatrix(intCount, ExseCol.对应数), "X")) Then
118                     lblMessage.Caption = lblMessage.Tag & intCount & IIf(.TextMatrix(intCount, ExseCol.对应数) = "", "行的对应数不能为空", "行不能为非数字型.")
120                     blnErr = True
                    End If
                
                    '可能是0.000等
122                 If Int(.TextMatrix(intCount, ExseCol.对应数)) = 0 And .TextMatrix(intCount, ExseCol.固定) = "√" Then
124                     .TextMatrix(intCount, ExseCol.固定) = ""
126                     lblMessage.Caption = lblMessage.Tag & intCount & "行的对应数为0,应为非固定项,已自动更正."
                    End If
            
128                 If InStr(1, strTemp & ";", ";" & .RowData(intCount) & ";") > 0 Then
130                     lblMessage.Caption = lblMessage.Tag & intCount & "行收费项目与前面的收费项目有重复！"
132                     blnErr = True
                    End If
                
134                 If Not blnErr Then
                        '无错误的行，才加入
136                     If strGen(UBound(strGen)) <> "" Then ReDim Preserve strGen(UBound(strGen) + 1)
138                     strGen(UBound(strGen)) = .RowData(intCount)
140                     strGen(UBound(strGen)) = strGen(UBound(strGen)) & "|" & .TextMatrix(intCount, ExseCol.项目名)
142                     strGen(UBound(strGen)) = strGen(UBound(strGen)) & "|" & .TextMatrix(intCount, ExseCol.规格)
144                     strGen(UBound(strGen)) = strGen(UBound(strGen)) & "|" & .TextMatrix(intCount, ExseCol.单位)
146                     strGen(UBound(strGen)) = strGen(UBound(strGen)) & "|" & .TextMatrix(intCount, ExseCol.当前价)
148                     strGen(UBound(strGen)) = strGen(UBound(strGen)) & "|" & .TextMatrix(intCount, ExseCol.对应数)
150                     strGen(UBound(strGen)) = strGen(UBound(strGen)) & "|" & .TextMatrix(intCount, ExseCol.固定)
152                     strGen(UBound(strGen)) = strGen(UBound(strGen)) & "|" & .TextMatrix(intCount, ExseCol.从项)
154                     strGen(UBound(strGen)) = strGen(UBound(strGen)) & "|" & .TextMatrix(intCount, ExseCol.收费方式)
156                     strGen(UBound(strGen)) = strGen(UBound(strGen)) & "|" & lngSource  '病人来源
158                     strGen(UBound(strGen)) = strGen(UBound(strGen)) & "|" & lngDeptID  '适用科室id
                    End If
                End If
            Next
        End With
    
        '检查项目 增加附加费用
160     If stbExse.TabVisible(1) Then
162         With Me.vfgExse
164             For intCount = .FixedRows To .Rows - 1
166                 If Val(.TextMatrix(intCount, .ColIndex("收费ID"))) > 0 And Val(.TextMatrix(intCount, .ColIndex("数量"))) > 0 _
                      And .TextMatrix(intCount, .ColIndex("部位")) <> "" And .TextMatrix(intCount, .ColIndex("方法")) <> "" Then
168                       If strPlan(UBound(strPlan)) <> "" Then ReDim Preserve strPlan(UBound(strPlan) + 1)
170                       strPlan(UBound(strPlan)) = .TextMatrix(intCount, .ColIndex("分组"))
172                       strPlan(UBound(strPlan)) = strPlan(UBound(strPlan)) & "|" & .TextMatrix(intCount, .ColIndex("部位"))
174                       strPlan(UBound(strPlan)) = strPlan(UBound(strPlan)) & "|" & .TextMatrix(intCount, .ColIndex("方法"))
176                       strPlan(UBound(strPlan)) = strPlan(UBound(strPlan)) & "|" & .TextMatrix(intCount, .ColIndex("项目名"))
178                       strPlan(UBound(strPlan)) = strPlan(UBound(strPlan)) & "|" & .TextMatrix(intCount, .ColIndex("单位"))
180                       strPlan(UBound(strPlan)) = strPlan(UBound(strPlan)) & "|" & .TextMatrix(intCount, .ColIndex("价格"))
182                       strPlan(UBound(strPlan)) = strPlan(UBound(strPlan)) & "|" & .TextMatrix(intCount, .ColIndex("数量"))
184                       strPlan(UBound(strPlan)) = strPlan(UBound(strPlan)) & "|" & .TextMatrix(intCount, .ColIndex("固定"))
                          strPlan(UBound(strPlan)) = strPlan(UBound(strPlan)) & "|" & .TextMatrix(intCount, .ColIndex("收费方式"))
186                       strPlan(UBound(strPlan)) = strPlan(UBound(strPlan)) & "|" & .TextMatrix(intCount, .ColIndex("性质"))
188                       strPlan(UBound(strPlan)) = strPlan(UBound(strPlan)) & "|" & .TextMatrix(intCount, .ColIndex("收费ID"))
190                       strPlan(UBound(strPlan)) = strPlan(UBound(strPlan)) & "|" & lngSource  '病人来源
192                       strPlan(UBound(strPlan)) = strPlan(UBound(strPlan)) & "|" & lngDeptID  '适用科室id
                    End If
                Next
            End With
        End If
    
194     If stbExse.TabVisible(2) Then
196         With Me.vfg加收
198             For intCount = .FixedRows To .Rows - 1
200                 If Val(.TextMatrix(intCount, .ColIndex("收费ID"))) > 0 And Val(.TextMatrix(intCount, .ColIndex("数量"))) > 0 Then
202                     If strAppend(UBound(strAppend)) <> "" Then ReDim Preserve strAppend(UBound(strAppend) + 1)
                        'Append格式：项目名|单位|价格|数量|固定|性质|收费ID|病人来源|适用科室id
204                     strAppend(UBound(strAppend)) = .TextMatrix(intCount, .ColIndex("项目名"))
206                     strAppend(UBound(strAppend)) = strAppend(UBound(strAppend)) & "|" & .TextMatrix(intCount, .ColIndex("单位"))
208                     strAppend(UBound(strAppend)) = strAppend(UBound(strAppend)) & "|" & .TextMatrix(intCount, .ColIndex("价格"))
210                     strAppend(UBound(strAppend)) = strAppend(UBound(strAppend)) & "|" & .TextMatrix(intCount, .ColIndex("数量"))
212                     strAppend(UBound(strAppend)) = strAppend(UBound(strAppend)) & "|" & .TextMatrix(intCount, .ColIndex("固定"))
214                     strAppend(UBound(strAppend)) = strAppend(UBound(strAppend)) & "|" & .TextMatrix(intCount, .ColIndex("性质"))
216                     strAppend(UBound(strAppend)) = strAppend(UBound(strAppend)) & "|" & .TextMatrix(intCount, .ColIndex("收费ID"))
218                     strAppend(UBound(strAppend)) = strAppend(UBound(strAppend)) & "|" & lngSource  '病人来源
220                     strAppend(UBound(strAppend)) = strAppend(UBound(strAppend)) & "|" & lngDeptID  '适用科室id
                    End If
                Next
            End With
        End If
        '删除原有数据，并把现在的数据加入到缓存
222     Select Case lngSource
        Case 0 '所有
            '不区分科室
224         mGen0 = strGen
226         mPlace0 = strPlan
228         mAppend0 = strAppend
230     Case 1 '门诊
232         Call UpdateArray(mGen1, strGen, 10, lngDeptID)
234         Call UpdateArray(mPlace1, strPlan, 12, lngDeptID)
236         Call UpdateArray(mAppend1, strAppend, 8, lngDeptID)
238     Case 2 '住院
240         Call UpdateArray(mGen2, strGen, 10, lngDeptID)
242         Call UpdateArray(mPlace2, strPlan, 12, lngDeptID)
244         Call UpdateArray(mAppend2, strAppend, 8, lngDeptID)
246     Case 3 '体检
248         Call UpdateArray(mGen3, strGen, 10, lngDeptID)
250         Call UpdateArray(mPlace3, strPlan, 12, lngDeptID)
252         Call UpdateArray(mAppend3, strAppend, 8, lngDeptID)
        End Select

254     CacheData = True
        Exit Function
hErr:
256     MsgBox "CacheData第" & CStr(Erl()) & "行：" & err.Description
End Function

Private Sub UpdateArray(ByRef ArryA() As String, ByRef ArryB() As String, ByVal lngSub As Long, ByVal lngDeptKey As Long)
        '将B数组的数据，更新到A数组中。
        'A数组存的是 收费对照缓存
        'B数组存的是 当前界面上的收费对照。
        'lngSub: 科室ID所在的下标
        'lngDeptKey: 科室ID
        On Error GoTo hErr
    
100     For intCount = LBound(ArryA) To UBound(ArryA)
102         If ArryA(intCount) <> "" Then
104             If Split(ArryA(intCount), "|")(lngSub) <> lngDeptKey Then
106                 If ArryB(UBound(ArryB)) <> "" Then ReDim Preserve ArryB(UBound(ArryB) + 1)
108                 ArryB(UBound(ArryB)) = ArryA(intCount)
                End If
            End If
        Next
110     ArryA = ArryB
        Exit Sub
hErr:
112     MsgBox "UpdateArryay第" & CStr(Erl()) & "行：" & err.Description
End Sub

Private Function CheckArrData(ByRef ArryA() As String) As Boolean
    '检查缓存的数据是否有问题
    
    CheckArrData = False
    If Val(Me.lblItem.Tag) = 0 Then lblMessage.Caption = lblMessage.Tag & "未正确指定诊疗项目！": Me.txtItem.SetFocus: Exit Function
    
    '校验从属：可以全部为主项(相当于不是套餐)，但如果存在从项，则只能有且必须有一个主项，且该主项必须为固定项目(不能删除)。
    Dim bln存在从项 As Boolean
    Dim int主项数 As Integer
    Dim int主项所在行 As Integer
    Dim intRows As Integer
    Dim rs As New ADODB.Recordset
    'Gen格式为： id|名称|规格|计算单位|价格|数量|固定|从项|收费方式|病人来源|适用科室id
    For intCount = LBound(ArryA) To UBound(ArryA)
        If ArryA(intCount) <> "" Then
            If Split(ArryA(intCount), "|")(7) = "√" Then
                bln存在从项 = True
                Exit For
            End If
        End If
    Next
    If bln存在从项 Then
        For intCount = LBound(ArryA) To UBound(ArryA)
            If Split(ArryA(intCount), "|")(7) <> "√" Then
                int主项所在行 = intCount
                int主项数 = int主项数 + 1
                If int主项数 > 1 Then
                    lblMessage.Caption = "提示：只能允许一个主项。"
                    Exit Function
                End If
            End If
        Next
        If int主项数 = 1 Then
            If Split(ArryA(int主项所在行), "|")(6) <> "√" Then
                lblMessage.Caption = "提示：第" & int主项所在行 & "行是主项，必须为固定项目。"
                Exit Function
            End If
        End If
        If int主项数 = 0 Then
            lblMessage.Caption = "提示：必须要有一个主项。"
            Exit Function
        End If
    End If
 
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '检查主项的价格是否存在多个收入项目，如果有则提示，不能保存
    If bln存在从项 Then
        gstrSql = "Select Id From 收费价目 Where 收费细目id=[1] And 执行日期 <= SYSDATE AND (终止日期 > SYSDATE OR 终止日期 IS NULL)    And 价格等级 Is Null "
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Split(ArryA(int主项所在行), "|")(0)), gstrPriceClass)
        If rs.RecordCount > 1 Then
            lblMessage.Caption = "提示：主项的价格存在多个收入项目，不能保存。"
            Exit Function
        End If
        rs.Close
    End If
    CheckArrData = True
        
End Function

Private Function SaveArryData(ByVal lngSource As Long, lngDeptID As Long, ArryGen() As String, ArryPlan() As String, ArryAppend() As String) As Boolean
    '保存缓存中的数据。
    'lngSource: 病人来源
    'lngDeptID: 科室ID
    'arrdept :  适用科室数组
    'arrGen  :  普通对照数组
    'arrPlan :  部位对照数组
    'arrAppen:  附加对照数组
    
    Dim strItemList As String
    Dim lngCount As Long ' 总个数
    Dim lngLoop As Long, lngEndloop As Long
    Dim varItem As Variant, strItem As String, blnBeginTrans As Boolean, i As Integer
 
    strTemp = "": strItemList = ""
    
    For lngCount = LBound(ArryGen) To UBound(ArryGen)
        If ArryGen(lngCount) <> "" Then
            'Gen格式为： id|名称|规格|计算单位|价格|数量|固定|从项|收费方式|病人来源|适用科室id
            If lngSource = Val(Split(ArryGen(lngCount), "|")(9)) And lngDeptID = Val(Split(ArryGen(lngCount), "|")(10)) Then
                If Trim(Split(ArryGen(lngCount), "|")(1)) <> "" And Split(ArryGen(lngCount), "|")(0) <> 0 Then
                    strItemList = strItemList & "|" & Split(ArryGen(lngCount), "|")(0) & "^" & _
                              Val(Split(ArryGen(lngCount), "|")(5)) & "^" & _
                              IIf(Trim(Split(ArryGen(lngCount), "|")(6)) = "", 0, 1) & "^" & _
                              IIf(Trim(Split(ArryGen(lngCount), "|")(7)) = "", 0, 1) & "^0^^ " & _
                              Val(Mid(Split(ArryGen(lngCount), "|")(8), 1, 1)) & ""
                End If
            End If
        End If
    Next
    
    '检查项目 增加附加费用
    
        'Pla格式：分组|部位|方法|项目名|单位|价格|数量|固定|收费方式|性质|收费ID|病人来源|适用科室id
    For lngCount = LBound(ArryPlan) To UBound(ArryPlan)
        If ArryPlan(lngCount) <> "" Then
            If lngSource = Val(Split(ArryPlan(lngCount), "|")(11)) And lngDeptID = Val(Split(ArryPlan(lngCount), "|")(12)) Then
                If Val(Split(ArryPlan(lngCount), "|")(10)) > 0 And Val(Split(ArryPlan(lngCount), "|")(6)) > 0 _
                  And Split(ArryPlan(lngCount), "|")(1) <> "" And Split(ArryPlan(lngCount), "|")(2) <> "" Then
                  
                    strItemList = strItemList & "|" & Val(Split(ArryPlan(lngCount), "|")(10)) & "^" & _
                              Val(Split(ArryPlan(lngCount), "|")(6)) & "^" & _
                             IIf(Split(ArryPlan(lngCount), "|")(7) = "√", 1, 0) & "^0^0^" & _
                             Split(ArryPlan(lngCount), "|")(1) & "^" & _
                             Split(ArryPlan(lngCount), "|")(2) & "^" & Val(Split(ArryPlan(lngCount), "|")(8))
                End If
            End If
        End If
    Next
    
     
    
    'Append格式：项目名|单位|价格|数量|固定|性质|收费ID|病人来源|适用科室id
    For lngCount = LBound(ArryAppend) To UBound(ArryAppend)
        If ArryAppend(lngCount) <> "" Then
            If lngSource = Val(Split(ArryAppend(lngCount), "|")(7)) And lngDeptID = Val(Split(ArryAppend(lngCount), "|")(8)) Then
                If Val(Split(ArryAppend(lngCount), "|")(6)) > 0 And Val(Split(ArryAppend(lngCount), "|")(3)) > 0 Then
                    strItemList = strItemList & "|" & Val(Split(ArryAppend(lngCount), "|")(6)) & "^" & _
                             Val(Split(ArryAppend(lngCount), "|")(3)) & "^" & _
                             IIf(Split(ArryAppend(lngCount), "|")(4) = "√", 1, 0) & "^0^1^^0"
                End If
            End If
        End If
    Next
        
    
    If strItemList <> "" Then strItemList = Mid(strItemList, 2)
    
    varItem = Split(strItemList, "|")
    lngCount = UBound(varItem)
    lngEndloop = 0
    gcnOracle.BeginTrans
    blnBeginTrans = True
    
    For lngLoop = 0 To lngCount
        
        strItem = strItem & "|" & varItem(lngLoop)
        If i = 40 Then
            strItem = Mid(strItem, 2)
            If Me.optPreproty(2).Value Then
                gstrSql = "zl_诊疗收费_UPDATE(" & Val(Me.lblItem.Tag) & ",2,'" & strItem & "'," & IIf(lngEndloop = 0, 1, 0) & "," & IIf(lngDeptID = 0, "Null", lngDeptID) & "," & lngSource & ")"
            ElseIf Me.optPreproty(1).Value Then
                gstrSql = "zl_诊疗收费_UPDATE(" & Val(Me.lblItem.Tag) & ",1,'" & strItem & "'," & IIf(lngEndloop = 0, 1, 0) & "," & IIf(lngDeptID = 0, "Null", lngDeptID) & "," & lngSource & ")"
            Else
                gstrSql = "zl_诊疗收费_UPDATE(" & Val(Me.lblItem.Tag) & ",0,'" & strItem & "'," & IIf(lngEndloop = 0, 1, 0) & "," & IIf(lngDeptID = 0, "Null", lngDeptID) & "," & lngSource & ")"
            End If
            err = 0: On Error GoTo ErrHand
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            i = 0: strItem = ""
            lngEndloop = lngEndloop + 1
        End If
        i = i + 1
    Next
    
    If Left(strItem, 1) = "|" Then
        strItem = Mid(strItem, 2)
        If Me.optPreproty(2).Value Then
            gstrSql = "zl_诊疗收费_UPDATE(" & Val(Me.lblItem.Tag) & ",2,'" & strItem & "'," & IIf(lngEndloop = 0, 1, 0) & "," & IIf(lngDeptID = 0, "Null", lngDeptID) & "," & lngSource & ")"
        ElseIf Me.optPreproty(1).Value Then
            gstrSql = "zl_诊疗收费_UPDATE(" & Val(Me.lblItem.Tag) & ",1,'" & strItem & "'," & IIf(lngEndloop = 0, 1, 0) & "," & IIf(lngDeptID = 0, "Null", lngDeptID) & "," & lngSource & ")"
        Else
            gstrSql = "zl_诊疗收费_UPDATE(" & Val(Me.lblItem.Tag) & ",0,'" & strItem & "'," & IIf(lngEndloop = 0, 1, 0) & "," & IIf(lngDeptID = 0, "Null", lngDeptID) & "," & lngSource & ")"
        End If
        err = 0: On Error GoTo ErrHand
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    End If
    
    If lngLoop = 0 Then '11303 不能全部删除对照的收费项目
        If Me.optPreproty(2).Value Then
            gstrSql = "zl_诊疗收费_UPDATE(" & Val(Me.lblItem.Tag) & ",2,'',1," & IIf(lngDeptID = 0, "Null", lngDeptID) & "," & lngSource & ")"
        ElseIf Me.optPreproty(1).Value Then
            gstrSql = "zl_诊疗收费_UPDATE(" & Val(Me.lblItem.Tag) & ",1,'',1," & IIf(lngDeptID = 0, "Null", lngDeptID) & "," & lngSource & ")"
        Else
            gstrSql = "zl_诊疗收费_UPDATE(" & Val(Me.lblItem.Tag) & ",0,'',1," & IIf(lngDeptID = 0, "Null", lngDeptID) & "," & lngSource & ")"
        End If
        err = 0: On Error GoTo ErrHand
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    End If
    
    gcnOracle.CommitTrans
    blnBeginTrans = False
    
    Exit Function

ErrHand:
    If blnBeginTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function delDept(ByVal lngSource As Long, ByVal lngDeptID As Long) As Boolean
    '从缓存中删除科室及对应的收费对照。
    Dim i As Integer, curDept() As String
    ReDim curDept(0) As String
    On Error GoTo hErr
    delDept = False
    If lngSource = 1 Then
        Call DelDeptCharge(lngDeptID, mDept1, mGen1, mPlace1, mAppend1)
    ElseIf lngSource = 2 Then
        Call DelDeptCharge(lngDeptID, mDept2, mGen2, mPlace2, mAppend2)
    ElseIf lngSource = 3 Then
        Call DelDeptCharge(lngDeptID, mDept3, mGen3, mPlace3, mAppend3)
    End If
    delDept = True
    Exit Function
hErr:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub DelDeptCharge(ByVal lngDeptID As Long, arrDept() As String, arrGen() As String, arrPlan() As String, arrAppend() As String)
    '从缓存数据中删除指定科室的对照
    Dim i As Integer, curDept() As String
    Dim curGen() As String, curPlan() As String, curAppend() As String
    
    ReDim curDept(0) As String
    ReDim curGen(0) As String
    ReDim curPlan(0) As String
    ReDim curAppend(0) As String
    
    For i = LBound(arrDept) To UBound(arrDept)
        If arrDept(i) <> "" Then
        If Val(Split(arrDept(i), "|")(0)) <> lngDeptID Then
            If curDept(UBound(curDept)) <> "" Then ReDim Preserve curDept(UBound(curDept) + 1)
            curDept(UBound(curDept)) = arrDept(i)
        End If
        End If
    Next
    arrDept = curDept
            
    For i = LBound(arrGen) To UBound(arrGen)
        If arrGen(i) <> "" Then
        If Val(Split(arrGen(i), "|")(10)) <> lngDeptID Then
            If curGen(UBound(curGen)) <> "" Then ReDim Preserve curGen(UBound(curGen) + 1)
            curGen(UBound(curGen)) = arrGen(i)
        End If
        End If
    Next
    arrGen = curGen
    
    For i = LBound(arrPlan) To UBound(arrPlan)
        If arrPlan(i) <> "" Then
        If Val(Split(arrPlan(i), "|")(11)) <> lngDeptID Then
            If curPlan(UBound(curPlan)) <> "" Then ReDim Preserve curPlan(UBound(curPlan) + 1)
            curPlan(UBound(curPlan)) = arrPlan(i)
        End If
        End If
    Next
    arrPlan = curPlan
    
    For i = LBound(arrAppend) To UBound(arrAppend)
        If arrAppend(i) <> "" Then
        If Val(Split(arrAppend(i), "|")(8)) <> lngDeptID Then
            If curAppend(UBound(curAppend)) <> "" Then ReDim Preserve curAppend(UBound(curAppend) + 1)
            curAppend(UBound(curAppend)) = arrAppend(i)
        End If
        End If
    Next
    arrAppend = curAppend
    
End Sub

Private Sub DeptCopy(ByVal lngSource As Long, ByVal lngOldDeptID As Long)
    '复制当前选中科室的项目对照到其他科室
    'lngSource   :病人来源
    'lngOldDeptID :现在选中的科室ID
    
'功能：多功能选择器
'参数：
'     frmParent=显示的父窗体
'     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
'     bytStyle=选择器风格
'       为0时:列表风格:ID,…
'       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
'       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
'     strTitle=选择器功能命名,也用于个性化区分
'     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
'     strSeek=当bytStyle<>2时有效,缺省定位的项目。
'             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
'             bytStyle=1时,可以是编码或名称
'     strNote=选择器的说明文字
'     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
'     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
'     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
'     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
'     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
'     blnSearch=是否显示行号,并可以输入行号定位
'     blnMulti=是否允许多选
'     arrInput=对应的各个SQL参数值,按顺序传入,必须为明确类型
'返回：取消=Nothing,选择=SQL源的单行记录集
    Dim strSql As String
    Dim rsDept As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim strInfo As String
    Dim strGen() As String, strPlan() As String, strAppend() As String, strDept() As String
    Dim strDeptList As String
    ReDim strDept(0) As String
    ReDim strGen(0) As String
    ReDim strPlan(0) As String
    ReDim strAppend(0) As String
    Dim varDept As Variant, strLine As String, strReturn As String
    Dim i As Integer
    
    On Error GoTo ErrHandle
    If lstDept.ListCount > 0 Then
        For i = 0 To lstDept.ListCount - 1
            strDeptList = strDeptList & "," & lstDept.ItemData(i)
        Next
    End If
    
    If lngSource = 1 Or lngSource = 3 Then  '1：门诊； 3：体检；
        strSql = "Select Distinct a.编码, a.名称, a.ID" & vbNewLine & _
                "From 部门表 A, 部门性质说明 D" & vbNewLine & _
                "Where a.Id = d.部门id And (d.服务对象 = 1 Or d.服务对象 = 3) and d.工作性质 in ('临床','检查','检验','手术','麻醉','治疗','营养') " & vbNewLine & _
                " And (A.撤档时间 is null or A.撤档时间=to_date('3000-01-01','YYYY-MM-DD')) " & vbNewLine & _
                IIf(strDeptList = "", "", "   And Instr([1], ',' || a.Id || ',') <= 0") & vbNewLine & _
                "Order By 编码, 名称"
    Else    '2：住院；
        strSql = "Select Distinct a.编码, a.名称, a.ID" & vbNewLine & _
                "From 部门表 A, 部门性质说明 D" & vbNewLine & _
                "Where a.Id = d.部门id And  (d.服务对象 = 2 Or d.服务对象 = 3) And d.工作性质 in ('护理','检查','检验','手术','麻醉','治疗','营养') " & vbNewLine & _
                " And (A.撤档时间 is null or A.撤档时间=to_date('3000-01-01','YYYY-MM-DD')) " & vbNewLine & _
                IIf(strDeptList = "", "", "   And Instr([1], ',' || a.Id || ',') <= 0") & vbNewLine & _
                "Order By 编码, 名称"
    End If
    Set rsDept = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strDeptList & ",")
    strReturn = frmSelCur.ShowCurrSel(Me, rsDept, "编码,1200,0,2;名称,1800,0,2;ID,0,1,2", "选择病区", True, , , 5000, True)
    
    If strReturn = "" Then Exit Sub
    varDept = Split(strReturn, "|")
    strInfo = ""
    
    For i = LBound(varDept) To UBound(varDept)
        '
        strLine = varDept(i)
        If UBound(Split(strLine, ",")) = 2 Then
            '检验是否已设了对照，没有才能复制
            strSql = "Select 收费项目ID From 诊疗收费关系 Where 病人来源=[3] And 适用科室ID=[1] and 诊疗项目ID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CLng(Split(strLine, ",")(2)), CLng(Me.lblItem.Tag), lngSource)
            If rsTmp.EOF Then
                '取当前页面，当前科室的费用对照
                If lngSource = 1 Then
                    Call arrCopyCharge(lngOldDeptID, 10, CLng(Split(strLine, ",")(2)), mGen1)
                    Call arrCopyCharge(lngOldDeptID, 11, CLng(Split(strLine, ",")(2)), mPlace1)
                    Call arrCopyCharge(lngOldDeptID, 8, CLng(Split(strLine, ",")(2)), mAppend1)
                    Call arrCopyCharge(lngOldDeptID, 0, CLng(Split(strLine, ",")(2)), mDept1)
                ElseIf lngSource = 2 Then
                    Call arrCopyCharge(lngOldDeptID, 10, CLng(Split(strLine, ",")(2)), mGen2)
                    Call arrCopyCharge(lngOldDeptID, 11, CLng(Split(strLine, ",")(2)), mPlace2)
                    Call arrCopyCharge(lngOldDeptID, 8, CLng(Split(strLine, ",")(2)), mAppend2)
                    Call arrCopyCharge(lngOldDeptID, 0, CLng(Split(strLine, ",")(2)), mDept2)
                ElseIf lngSource = 3 Then
                    Call arrCopyCharge(lngOldDeptID, 10, CLng(Split(strLine, ",")(2)), mGen3)
                    Call arrCopyCharge(lngOldDeptID, 11, CLng(Split(strLine, ",")(2)), mPlace3)
                    Call arrCopyCharge(lngOldDeptID, 8, CLng(Split(strLine, ",")(2)), mAppend3)
                    Call arrCopyCharge(lngOldDeptID, 0, CLng(Split(strLine, ",")(2)), mDept3)
                End If
                lstDept.AddItem "" & Split(strLine, ",")(1) & "(" & Split(strLine, ",")(0) & ")"
                lstDept.ItemData(lstDept.NewIndex) = Val(Split(strLine, ",")(2))
                '复制数据
                'Call SaveArryData(lngSource, CLng("" & rsDept!ID), strGen, strPlan, strAppend)
            Else
               strInfo = IIf(strInfo = "", "", vbNewLine) & "" & Split(strLine, ",")(0) & " " & Split(strLine, ",")(1) & " 该科室已经设定了费用！"
            End If
        End If

    Next
    If strInfo <> "" Then
        MsgBox strInfo
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub arrCopyCharge(ByVal lngOldDeptID As Long, ByVal lngDeptIndex As Long, ByVal lngNewDeptID As Long, arrA() As String)
    '复制arrA中的指定科室的对照明细到arrA中
    'lngDeptID : 科室ID
    'lngDeptIndex :科室ID在arrA中的下标
    'arrA:数组A，对照明细
    
    Dim arrB() As String
    Dim lngRow As Long, i As Integer, varTmp As Variant, strTmp As String
    ReDim arrB(0) As String
    For lngRow = LBound(arrA) To UBound(arrA)
        If arrA(lngRow) <> "" Then
            If Split(arrA(lngRow), "|")(lngDeptIndex) = lngOldDeptID Then
                strTmp = ""
                varTmp = Split(arrA(lngRow), "|")
                For i = LBound(varTmp) To UBound(varTmp)
                    If i = lngDeptIndex Then
                        strTmp = strTmp & "|" & lngNewDeptID
                    Else
                        strTmp = strTmp & "|" & varTmp(i)
                    End If
                Next
                If strTmp <> "" Then
                    strTmp = Mid(strTmp, 2)
                If arrB(UBound(arrB)) <> "" Then ReDim Preserve arrB(UBound(arrB) + 1)
                    arrB(UBound(arrB)) = strTmp
                End If
            End If
        End If
    Next
    
    Dim blnAdd As Boolean
    '将arrB加到ArrA中
    For lngRow = LBound(arrB) To UBound(arrB)
        strTmp = arrB(lngRow)
        blnAdd = True
        For i = LBound(arrA) To UBound(arrA)
            If strTmp = arrA(i) Then
                blnAdd = False
                Exit For
            End If
        Next
        
        If blnAdd Then
            If arrA(UBound(arrA)) <> "" Then ReDim Preserve arrA(UBound(arrA) + 1)
            arrA(UBound(arrA)) = strTmp
        End If
    Next
End Sub

Private Function GetCurrSource() As Long
    '根据取当前选中页面获取病人来源
    If Me.tabDept.SelectedItem.Caption = "所有科室" Then
        GetCurrSource = 0
    ElseIf Me.tabDept.SelectedItem.Caption = "门诊科室" Then
        GetCurrSource = 1
    ElseIf Me.tabDept.SelectedItem.Caption = "住院科室" Then
        GetCurrSource = 2
    ElseIf Me.tabDept.SelectedItem.Caption = "体检科室" Then
        GetCurrSource = 3
    End If
End Function
