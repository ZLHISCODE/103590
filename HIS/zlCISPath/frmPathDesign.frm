VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPathDesign 
   AutoRedraw      =   -1  'True
   Caption         =   "临床路径设计"
   ClientHeight    =   7830
   ClientLeft      =   2310
   ClientTop       =   2040
   ClientWidth     =   11565
   Icon            =   "frmPathDesign.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7830
   ScaleWidth      =   11565
   Begin VB.Frame fraSelect 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4320
      TabIndex        =   14
      Top             =   960
      Width           =   3015
      Begin VB.OptionButton optSelect 
         Caption         =   "护士"
         Height          =   375
         Index           =   2
         Left            =   2160
         TabIndex        =   17
         Top             =   0
         Width           =   735
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "医生"
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   16
         Top             =   0
         Width           =   735
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "全部"
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   15
         Top             =   0
         Width           =   735
      End
      Begin VB.Label lblSendNote 
         Caption         =   "生成者："
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   75
         Width           =   1215
      End
   End
   Begin VB.PictureBox picCenter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   240
      ScaleHeight     =   4695
      ScaleWidth      =   14055
      TabIndex        =   4
      Top             =   2040
      Width           =   14055
      Begin VB.Frame fraSplit 
         BorderStyle     =   0  'None
         ForeColor       =   &H80000011&
         Height          =   45
         Left            =   0
         MousePointer    =   7  'Size N S
         TabIndex        =   11
         Top             =   1560
         Width           =   9735
      End
      Begin VB.PictureBox picBottom 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   0
         ScaleHeight     =   2415
         ScaleWidth      =   12975
         TabIndex        =   5
         Top             =   2040
         Width           =   12975
         Begin VB.CommandButton cmdCheck 
            Caption         =   "审核不过"
            Height          =   300
            Index           =   1
            Left            =   8640
            TabIndex        =   20
            Top             =   360
            Visible         =   0   'False
            Width           =   1100
         End
         Begin VB.CommandButton cmdCheck 
            Caption         =   "审核通过"
            Height          =   300
            Index           =   0
            Left            =   7440
            TabIndex        =   19
            Top             =   360
            Visible         =   0   'False
            Width           =   1100
         End
         Begin zlCISPath.UCAdviceList ucAdvice 
            Height          =   1335
            Index           =   0
            Left            =   480
            TabIndex        =   13
            Top             =   720
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   2355
         End
         Begin VB.Frame fraSplit2 
            BorderStyle     =   0  'None
            Height          =   2415
            Left            =   6120
            MousePointer    =   9  'Size W E
            TabIndex        =   7
            Top             =   720
            Width           =   60
         End
         Begin VB.ComboBox cboTimes 
            Height          =   300
            Left            =   8880
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   120
            Width           =   3495
         End
         Begin zlCISPath.UCAdviceList ucAdvice 
            Height          =   1335
            Index           =   1
            Left            =   7200
            TabIndex        =   12
            Top             =   720
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   2355
         End
         Begin VB.Label lblCurr 
            Caption         =   "当前医嘱详情"
            Height          =   255
            Left            =   480
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblChange 
            Caption         =   "医嘱变动详情"
            Height          =   255
            Left            =   7440
            TabIndex        =   8
            Top             =   120
            Width           =   1215
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsPath 
         Height          =   825
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   4695
         _cx             =   8281
         _cy             =   1455
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   10218651
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   10218651
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   1500
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   101
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         Editable        =   2
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
         FrozenRows      =   1
         FrozenCols      =   1
         AllowUserFreezing=   0
         BackColorFrozen =   14811105
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   4440
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog cdgXML 
      Left            =   1770
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7470
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPathDesign.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17489
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
   Begin VSFlex8Ctl.VSFlexGrid vsPathExport 
      Height          =   1305
      Left            =   7560
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   3135
      _cx             =   5530
      _cy             =   2302
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   10218651
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   10218651
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   1500
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   101
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
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
      Editable        =   2
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
      FrozenRows      =   1
      FrozenCols      =   1
      AllowUserFreezing=   0
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Image ImgBranch 
      Height          =   240
      Left            =   3120
      Picture         =   "frmPathDesign.frx":0E1C
      Top             =   240
      Width           =   240
   End
   Begin XtremeSuiteControls.ShortcutCaption stcInfo 
      Height          =   390
      Left            =   1095
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   795
      Width           =   2955
      _Version        =   589884
      _ExtentX        =   5212
      _ExtentY        =   688
      _StockProps     =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      GradientColorLight=   16710907
      GradientColorDark=   16180453
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   285
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgMain 
      Bindings        =   "frmPathDesign.frx":766E
      Left            =   915
      Top             =   225
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmPathDesign.frx":7682
   End
End
Attribute VB_Name = "frmPathDesign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event DataChanged(ByVal 路径ID As Long)

Private WithEvents mfrmVersion As frmVersion
Attribute mfrmVersion.VB_VarHelpID = -1
Private WithEvents mfrmPathStep As frmPathStepEdit
Attribute mfrmPathStep.VB_VarHelpID = -1
Private WithEvents mfrmPathItem As frmPathItemEdit
Attribute mfrmPathItem.VB_VarHelpID = -1
Private WithEvents mfrmEvalEdit As frmEvaluateEdit
Attribute mfrmEvalEdit.VB_VarHelpID = -1
Private WithEvents mfrmAdviceContrast As frmAdviceContrast
Attribute mfrmAdviceContrast.VB_VarHelpID = -1

Private mlng路径ID As Long
Private mbytMode As CONST_MODE
Private mcolVersion As Collection
Private mcolBranch As Collection
Private mcolItemRowCol As Collection  'Value:Row,Col Key："_ "& 项目ID  LoadPathTable时记录下项目的行和列
Private mcolItemID As Collection    '记录与上一版本具在相同阶段，相同分类，相同名称的医嘱类项目下医嘱存在差异的项目ID
Private mstrPrivs As String
Private mblnReturn As Boolean
Private mlngNewRow As Long
Private mlngNewCol As Long
Private mstrDeptInfo As String  '路径管理界面显示的适用科室信息
Private mstrDiagInfo As String  '路径管理界面显示的适用病种信息

Private mrsAdvice As ADODB.Recordset    '对应医嘱动态记录集
Private mvEvalImport As TYPE_PATH_EVAL    '导入评估数据
Private mblnEditable As Boolean    '是否允许编辑
Private mblnChange As Boolean
Private mstrDelStepIDs As String    '被删除的时间阶段ID串
Private mstrDelItemIDs As String    '被删除的路径项目ID串
Private mstrChangeItemIDs As String     '路径变动项目的ID串;双审核模式: 路径变动项目的ID及审核状态的字符串
Private mblnNewVersion As Boolean
Private mblnAddNew As Boolean    '判断是否是新增分支
Private mlngDays As Long
Private mlng性质 As Long    '临床路径的性质，=1 合并路径 =0首要路径
Private mstr疾病编码 As String
Private mblnDiff As Boolean
Private mbytFunc As Byte     '用于区分项目变动和查看变动记录;1-查看变动记录,2-显示项目变动
Private mbytAudit As Byte    '双审核模式下:=0“药剂科审核”和“审核”（医务科）都没有;=1 仅有药剂科；=2仅有医务科;3两者都有
Private marrTime As Variant

Private Type PathTable_Clipboard
    Empty As Boolean
    项目集() As TYPE_PATH_ITEM    '包含空白项目
    vStep As TYPE_PATH_STEP
    BeginRow As Long
End Type
Private mvClipboard As PathTable_Clipboard    '内部剪贴板

Private Const ROW_HEIGHT_MIN = 270
Private Const COl_WIDTH_BASE = 2000
Private Enum CONST_MODE
    Mode_Show = 0
    Mode_Design = 1
End Enum
Private Enum CONST_COLOR
    Color_NewBack = &HE1E1FF
    Color_AuditBack = &HE1FFE1
    Color_StopBack = &HE1E1E1
    Color_DiffBack = &HFAEADA          '浅蓝 医嘱类项目与之前版本存在差异

    Color_NewLine = &H9B9BEC
    Color_AuditLine = &H9BEC9B
    Color_StopLine = &H9B9B9B
    
    Color_NeedAuditFore = &H9B9BEC    '路径项目存在待审核医嘱，项目字体颜色为此色号
End Enum
Private Enum CONST_AREA
    Area_Cross = 0
    Area_Category = 1
    Area_Step = 2
    Area_Item = 3
End Enum

Private Enum CONST_FUNCTION
    '文件-------------------------
    cmd_File_Save = 101
    cmd_File_SaveExit = 102

    cmd_File_CopyFrom = 111
    cmd_File_ImportXML = 112

    cmd_File_ExportXML = 121
    cmd_File_ExportExcel = 122

    cmd_File_PrintSetup = 131
    cmd_File_Preview = 132
    cmd_File_Print = 133

    cmd_File_Exit = 191

    '编辑-------------------------
    cmd_Edit_Undo = 301
    cmd_Edit_Redo = 302
    cmd_Edit_Copy = 303
    cmd_Edit_Paste = 304

    cmd_Edit_Caption = 310    '标签
    cmd_Edit_Edit = 311    '设置
    cmd_Edit_Insert = 312    '插入
    cmd_Edit_InsertBefore = 3121    '在前面插入
    cmd_Edit_InsertAfter = 3122    '在后面插入
    cmd_Edit_InsertBranch = 3123    '增加分支
    cmd_Edit_Delete = 313    '删除
    cmd_Edit_Modify = 314  '修改

    cmd_Edit_Version = 320    '版本选择
    cmd_Edit_VersionInfo = 321    '版本信息
    cmd_Edit_VersionNew = 322    '版本添加
    cmd_Edit_VersionDel = 323    '版本删除
    cmd_Edit_EvalImport = 324    '导入评估
    cmd_Edit_EvalStep = 325    '阶段评估
    cmd_Edit_EvalStepCopy = 326    '复制阶段评估
    cmd_Edit_BranchNew = 327    '分支添加
    cmd_Edit_BranchDel = 328    '分支删除
    cmd_Edit_Branch = 329    '分支选择
    cmd_Edit_ItemShow = 330  '显示项目变动\隐藏项目变动

    '查看-------------------------
    cmd_View_ToolBar = 701
    cmd_View_ToolBar_Button = 7011
    cmd_View_ToolBar_Text = 7012
    cmd_View_ToolBar_Size = 7013
    cmd_View_StatusBar = 702
    cmd_View_Refresh = 791
    cmd_View_Find = 721

    '帮助-------------------------
    cmd_Help_Help = 901
    cmd_Help_Web = 902
    cmd_Help_Web_Home = 9021
    cmd_Help_Web_Forum = 9023
    cmd_Help_Web_Mail = 9022
    cmd_Help_About = 991
End Enum

Private Enum CONST_IX_SELECT
    IX_ALL = 0
    IX_医生 = 1
    IX_护士 = 2
End Enum

Private Function CheckPathItem() As Boolean
    Dim lngRow As Long, lngCol As Long
    Dim vItem As TYPE_PATH_ITEM
    With vsPath
    For lngRow = .FixedRows To .Rows - 1
        For lngCol = .FixedCols To .Cols - 1
            If TypeName(.Cell(flexcpData, lngRow, lngCol)) = TypeName(vItem) Then
                vItem = .Cell(flexcpData, lngRow, lngCol)
                If vItem.待审核医嘱IDs <> "" Then
                    CheckPathItem = True
                    Exit Function
                End If
            End If
        Next
    Next
    MsgBox "该临床路径不存在路径项目变动。", vbOKOnly + vbInformation, gstrSysName
    End With
End Function

Public Sub ShowDesign(frmParent As Object, ByVal lng路径ID As Long, ByVal strPrivs As String, Optional ByVal str疾病编码 As String)
    mbytMode = Mode_Design
    mlng路径ID = lng路径ID
    mstrPrivs = strPrivs
    mstr疾病编码 = str疾病编码
    mbytFunc = 0
    
    Me.Show 1, frmParent
End Sub

Public Sub zlRefresh(ByVal lng路径ID As Long, ByVal strPrivs As String, Optional ByVal strDeptInfo As String, Optional ByVal strDiagInfo As String)
'参数：查看模式时传入，strDeptInfo=适用科室信息，strDiagInfo=适用病种信息
    mlng路径ID = lng路径ID
    mstrPrivs = strPrivs
    mstrDeptInfo = strDeptInfo
    mstrDiagInfo = strDiagInfo
    
    Call LoadPathVersion
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
'功能：显示模式下，更新主程序的命令可用性
    Dim vVersion As TYPE_PATH_VERSION
    Dim objCombo As CommandBarComboBox
    Dim blnEnabled As Boolean
    
    Set objCombo = cbsMain(cbsMain.count - 1).FindControl(, cmd_Edit_Version, True)
    If Not objCombo Is Nothing Then
        If objCombo.ListIndex > 0 Then
            vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))
        End If
    End If
    Select Case Control.ID
    Case conMenu_File_ExportToXML '导出为XML文件
        If InStr(mstrPrivs, "导出XML") = 0 Then
            Control.Visible = False
        Else
            Control.Enabled = mlng路径ID <> 0 And vVersion.版本号 > 0
        End If
    Case conMenu_File_BatPrint  '全部输出到Excel
        Control.Enabled = mbytMode = Mode_Show
        
    Case conMenu_Edit_Compend '设计
        If InStr(mstrPrivs, "路径表设计") = 0 Then
            Control.Visible = False
        Else
            Control.Enabled = mlng路径ID <> 0
        End If
    Case conMenu_Edit_Audit '审核
        If InStr(";" & mstrPrivs & ";", ";审核;") = 0 Then
            Control.Visible = False
        Else
            If gbln双审核 Then
                blnEnabled = mlng路径ID <> 0 And vVersion.版本号 > 0 And vVersion.审核时间 = Empty And vVersion.药剂科审核时间 <> Empty
            Else
                blnEnabled = mlng路径ID <> 0 And vVersion.版本号 > 0 And vVersion.审核时间 = Empty
            End If
            If blnEnabled Then blnEnabled = objCombo.ListIndex = 1
            Control.Enabled = blnEnabled
        End If
    Case conMenu_Edit_Untread '取消审核
        If InStr(mstrPrivs, ";审核;") = 0 Then
            Control.Visible = False
        Else
            blnEnabled = mlng路径ID <> 0 And vVersion.版本号 > 0 And vVersion.审核时间 <> Empty And vVersion.停用时间 = Empty
            If blnEnabled Then blnEnabled = objCombo.ListIndex = 1
            Control.Enabled = blnEnabled
        End If
    Case conMenu_Edit_MedicalAudit '药剂科审核
        If gbln双审核 Then
            If InStr(mstrPrivs, ";药剂科审核;") > 0 Then
                blnEnabled = mlng路径ID <> 0 And vVersion.版本号 > 0 And vVersion.药剂科审核时间 = Empty And vVersion.审核时间 = Empty
                If blnEnabled Then blnEnabled = objCombo.ListIndex = 1
                Control.Enabled = blnEnabled
            Else
                Control.Visible = False
            End If
        Else
            Control.Visible = False
        End If
    Case conMenu_Edit_MedicalUntread '药剂科取消审核
        If gbln双审核 Then
            If InStr(mstrPrivs, ";药剂科审核;") > 0 Then
                blnEnabled = mlng路径ID <> 0 And vVersion.版本号 > 0 And vVersion.药剂科审核时间 <> Empty And vVersion.审核时间 = Empty
                If blnEnabled Then blnEnabled = objCombo.ListIndex = 1
                Control.Enabled = blnEnabled
            Else
                Control.Visible = False
            End If
        Else
            Control.Visible = False
        End If
    Case conMenu_Edit_Stop '停用
        If InStr(mstrPrivs, "停用") = 0 Then
            Control.Visible = False
        Else
            Control.Enabled = mlng路径ID <> 0 And vVersion.版本号 > 0 _
                And vVersion.审核时间 <> Empty And vVersion.停用时间 = Empty
        End If
    Case conMenu_Edit_Reuse '取消停用
        If InStr(mstrPrivs, "停用") = 0 Then
            Control.Visible = False
        Else
            blnEnabled = mlng路径ID <> 0 And vVersion.版本号 > 0 And vVersion.停用时间 <> Empty
            Control.Enabled = blnEnabled
        End If
    End Select
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl, Optional ByVal blnIsAll As Boolean)
'功能：显示模式下，执行主程序的命令
'      blnIsAll=是否批量输出到Excel
    Select Case Control.ID
    Case conMenu_File_PrintSet
        Call zlPrintSet
    Case conMenu_File_Print
        Call FuncPathTableOutput(1, blnIsAll)
    Case conMenu_File_Preview
        Call FuncPathTableOutput(2, blnIsAll)
    Case conMenu_File_Excel
        Call FuncPathTableOutput(3, blnIsAll)
    Case conMenu_File_ExportToXML '导出XML
        Call FuncExportToXML
    Case conMenu_Edit_Compend '设计
        '在主程序直接执行了
    Case conMenu_Edit_Audit '审核
        Call FuncVersionAudit(1)
    Case conMenu_Edit_Untread '取消审核
        Call FuncVersionAudit(-1)
    Case conMenu_Edit_MedicalAudit
        Call FuncVersionAudit(2)
    Case conMenu_Edit_MedicalUntread
        Call FuncVersionAudit(-2)
    Case conMenu_Edit_Stop '停用
        Call FuncVersionStop(True)
    Case conMenu_Edit_Reuse '取消停用
        Call FuncVersionStop(False)
    End Select
End Sub

Private Sub MainDefCommandBar()
'功能：主窗口菜单定义部份
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCombo As CommandBarComboBox
    Dim objCustom As CommandBarControlCustom

    Dim lngCount As Long

    '菜单定义
    '-----------------------------------------------------
    If mbytMode = Mode_Design Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
        objMenu.ID = conMenu_FilePopup
        With objMenu.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, cmd_File_Save, "保存(&S)")
            Set objControl = .Add(xtpControlButton, cmd_File_SaveExit, "保存并退出(&X)")

            Set objControl = .Add(xtpControlButton, cmd_File_CopyFrom, "从其他路径复制(&C)…"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, cmd_File_ImportXML, "从&XML文件导入…")

            Set objControl = .Add(xtpControlButton, cmd_File_PrintSetup, "打印设置(&U)…"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, cmd_File_Preview, "预览(&V)")
            Set objControl = .Add(xtpControlButton, cmd_File_Print, "打印(&P)")
            Set objControl = .Add(xtpControlButton, cmd_File_ExportExcel, "输出到&Excel…")
            Set objControl = .Add(xtpControlButton, cmd_File_ExportXML, "导出XM&L文件…"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, cmd_File_Exit, "退出(&X)"): objControl.BeginGroup = True
        End With

        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
        objMenu.ID = conMenu_EditPopup
        With objMenu.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, cmd_Edit_Undo, "撤消(&U)")
            Set objControl = .Add(xtpControlButton, cmd_Edit_Redo, "重做(&R)")
            Set objControl = .Add(xtpControlButton, cmd_Edit_Copy, "复制(&C)"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, cmd_Edit_Paste, "粘贴(&V)")

            Set objControl = .Add(xtpControlButton, cmd_Edit_Edit, "设置XXXX(&E)"): objControl.BeginGroup = True
            objControl.ShortcutText = "Enter"    '只是显示
            Set objPopup = .Add(xtpControlButtonPopup, cmd_Edit_Insert, "插入XXXX(&I)")
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, cmd_Edit_InsertBefore, "在前面插入(&1)")
                Set objControl = .Add(xtpControlButton, cmd_Edit_InsertAfter, "在后面插入(&2)")
                Set objControl = .Add(xtpControlButton, cmd_Edit_InsertBranch, "插入分支(&3)"): objControl.BeginGroup = True
            End With
            Set objControl = .Add(xtpControlButton, cmd_Edit_Modify, "修改分类(&X)")
            objControl.ShortcutText = "Modify"    '只是显示
            Set objControl = .Add(xtpControlButton, cmd_Edit_Delete, "删除XXXX(&D)")
            objControl.ShortcutText = "Delete"    '只是显示

            Set objControl = .Add(xtpControlButton, cmd_Edit_EvalImport, "导入评估设置(&P)"): objControl.BeginGroup = True
            Set objPopup = .Add(xtpControlSplitButtonPopup, cmd_Edit_EvalStep, "阶段评估设置(&J)")
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, cmd_Edit_EvalStepCopy, "复制前面阶段评估设置(&C)")
            End With

            Set objControl = .Add(xtpControlButton, cmd_Edit_VersionNew, "增加新的版本(&N)"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, cmd_Edit_VersionDel, "删除当前版本(&M)")
            objControl.IconId = cmd_Edit_Delete
            Set objControl = .Add(xtpControlButton, cmd_Edit_VersionInfo, "标准设置(&B)")
            Set objControl = .Add(xtpControlButton, cmd_Edit_BranchNew, "增加新的分支路径(&O)")
            Set objControl = .Add(xtpControlButton, cmd_Edit_BranchDel, "删除当前分支路径(&P)")
        End With

        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
        objMenu.ID = conMenu_ViewPopup
        With objMenu.CommandBar.Controls
            Set objPopup = .Add(xtpControlButtonPopup, cmd_View_ToolBar, "工具栏(&T)")
            With objPopup.CommandBar.Controls
                .Add xtpControlButton, cmd_View_ToolBar_Button, "标准按钮(&S)", -1, False
                .Add xtpControlButton, cmd_View_ToolBar_Text, "文本标签(&T)", -1, False
                .Add xtpControlButton, cmd_View_ToolBar_Size, "大图标(&B)", -1, False
            End With
            Set objControl = .Add(xtpControlButton, conMenu_View_StPath, "标准路径参考")
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, cmd_View_StatusBar, "状态栏(&S)")
            Set objControl = .Add(xtpControlButton, conMenu_View_Difference, "显示差异")
            objControl.ID = conMenu_View_Difference
            objControl.ToolTipText = "以不同背景色区别显示医嘱内容与上一版本有差异的项目"
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_View_Contrast, "对比查看")
            objControl.ToolTipText = "选中背景为蓝色的医嘱类项目后再执行对比查看"
            
            Set objControl = .Add(xtpControlButton, cmd_Edit_ItemShow, "显示项目变动")
            objControl.IconId = cmd_Edit_ItemShow
            objControl.BeginGroup = True
            objControl.Parameter = "显示"
            
            Set objControl = .Add(xtpControlButton, conMenu_View_Show, "查看变动记录")
            objControl.IconId = cmd_View_Find
            objControl.BeginGroup = True
            objControl.Parameter = "显示"
            
            Set objControl = .Add(xtpControlButton, cmd_View_Refresh, "刷新(&R)"): objControl.BeginGroup = True
        End With

        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
        objMenu.ID = conMenu_HelpPopup
        With objMenu.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, cmd_Help_Help, "帮助主题(&H)")
            Set objPopup = .Add(xtpControlButtonPopup, cmd_Help_Web, "&WEB上的" & gstrProductName)
            With objPopup.CommandBar.Controls
                .Add xtpControlButton, cmd_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
                .Add xtpControlButton, cmd_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False
                .Add xtpControlButton, cmd_Help_Web_Mail, "发送反馈(&M)", -1, False
            End With
            Set objControl = .Add(xtpControlButton, cmd_Help_About, "关于(&A)…")
            objControl.BeginGroup = True
        End With

        '工具栏定义:包括公共部份
        '-----------------------------------------------------
        Set objBar = cbsMain.Add("工具栏", xtpBarTop)
        With objBar.Controls
            Set objControl = .Add(xtpControlButton, cmd_File_Save, "保存")
            Set objControl = .Add(xtpControlButton, cmd_File_SaveExit, "保存退出")

            Set objControl = .Add(xtpControlButton, cmd_Edit_Undo, "撤消"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, cmd_Edit_Redo, "重做")
            Set objControl = .Add(xtpControlButton, cmd_Edit_Copy, "复制"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, cmd_Edit_Paste, "粘贴")

            Set objControl = .Add(xtpControlLabel, cmd_Edit_Caption, "分类："): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, cmd_Edit_Edit, "设置")
            objControl.ToolTipText = "Enter"
            Set objPopup = .Add(xtpControlPopup, cmd_Edit_Insert, "插入")
            objPopup.ID = cmd_Edit_Insert
            objPopup.IconId = cmd_Edit_Insert
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, cmd_Edit_InsertBefore, "在前面插入(&1)")
                Set objControl = .Add(xtpControlButton, cmd_Edit_InsertAfter, "在后面插入(&2)")
                Set objControl = .Add(xtpControlButton, cmd_Edit_InsertBranch, "插入分支(&3)"): objControl.BeginGroup = True
            End With
            Set objControl = .Add(xtpControlButton, cmd_Edit_Modify, "修改")
            objControl.ToolTipText = "Modify"    '只是显示
            Set objControl = .Add(xtpControlButton, cmd_Edit_Delete, "删除")
            objControl.ToolTipText = "Delete"    '只是显示

            Set objControl = .Add(xtpControlButton, cmd_Edit_EvalImport, "导入评估"): objControl.BeginGroup = True
            Set objPopup = .Add(xtpControlSplitButtonPopup, cmd_Edit_EvalStep, "阶段评估")
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, cmd_Edit_EvalStepCopy, "复制前面阶段评估设置(&C)")
            End With
            
            Set objControl = .Add(xtpControlButton, conMenu_View_Difference, "显示差异")
            objControl.ToolTipText = "以不同背景色区别显示医嘱内容与上一版本有差异的项目"
            objControl.ID = conMenu_View_Difference
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_View_Contrast, "对比查看")
            objControl.ToolTipText = "选中差异的医嘱类项目后对比查看"
            
            Set objControl = .Add(xtpControlButton, cmd_Edit_ItemShow, "显示项目变动")
            objControl.IconId = cmd_Edit_ItemShow
            objControl.BeginGroup = True
            objControl.Parameter = "显示"
            
            Set objControl = .Add(xtpControlButton, conMenu_View_Show, "查看变动记录")
            objControl.IconId = cmd_View_Find
            objControl.BeginGroup = True
            objControl.Parameter = "显示"
            
            Set objControl = .Add(xtpControlButton, cmd_Help_Help, "帮助"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, cmd_File_Exit, "退出")
            '查找
            Set objControl = .Add(xtpControlLabel, 0, "查找")
            objControl.IconId = cmd_View_Find
            objControl.Flags = xtpFlagRightAlign
            Set objCustom = .Add(xtpControlCustom, cmd_View_Find, "")
            objCustom.Handle = txtFind.Hwnd
            objCustom.Flags = xtpFlagRightAlign
        End With
    End If

    Set objBar = cbsMain.Add("版本栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlLabel, 0, "版      本")
        objControl.IconId = cmd_Edit_Version
        Set objCombo = .Add(xtpControlComboBox, cmd_Edit_Version, "")    '无法显示图标
        objCombo.Flags = xtpFlagControlStretched
        objCombo.DropDownListStyle = False
        If mbytMode = Mode_Design Then
            Set objControl = .Add(xtpControlButton, cmd_Edit_VersionNew, "新增版本")
            objControl.Flags = xtpFlagRightAlign
            objControl.Style = xtpButtonIconAndCaption
            Set objControl = .Add(xtpControlButton, cmd_Edit_VersionDel, "删除当前版本")
            objControl.IconId = cmd_Edit_Delete
            objControl.Flags = xtpFlagRightAlign
            objControl.Style = xtpButtonIconAndCaption
        End If
    End With

    Set objBar = cbsMain.Add("分支栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlLabel, 0, "分支路径")
        objControl.IconId = cmd_Edit_Branch
        Set objCombo = .Add(xtpControlComboBox, cmd_Edit_Branch, "")    '无法显示图标
        objCombo.Flags = xtpFlagControlStretched
        objCombo.DropDownListStyle = False
        If mbytMode = Mode_Design Then
            Set objControl = .Add(xtpControlButton, cmd_Edit_VersionInfo, "标准设置")
            objControl.Flags = xtpFlagRightAlign
            objControl.Style = xtpButtonIconAndCaption
            Set objControl = .Add(xtpControlButton, cmd_Edit_BranchNew, "新增分支路径")
            objControl.Flags = xtpFlagRightAlign
            objControl.Style = xtpButtonIconAndCaption
            Set objControl = .Add(xtpControlButton, cmd_Edit_BranchDel, "删除当前分支")
            objControl.Flags = xtpFlagRightAlign
            objControl.Style = xtpButtonIconAndCaption
        End If
    End With

    '设置一些公共的热键绑定
    '-----------------------------------------------------
    If mbytMode = Mode_Design Then
        With cbsMain.KeyBindings
            .Add FCONTROL, vbKeyS, cmd_File_Save    '保存
            .Add FCONTROL, vbKeyZ, cmd_Edit_Undo    '撤消
            .Add FCONTROL, vbKeyR, cmd_Edit_Redo    '重做
            .Add FCONTROL, vbKeyC, cmd_Edit_Copy    '复制
            .Add FCONTROL, vbKeyV, cmd_Edit_Paste    '粘贴
            .Add FCONTROL, vbKeyF, cmd_View_Find    '查找

            .Add FCONTROL, vbKeyE, cmd_Edit_EvalStep    '当前时间阶段评估标准
            .Add FCONTROL, vbKeyB, cmd_Edit_InsertBefore
            .Add FCONTROL, vbKeyI, cmd_Edit_InsertAfter
            
            .Add 0, vbKeyF4, conMenu_View_Contrast       '对比查看
            .Add 0, vbKeyF5, conMenu_View_Refresh    '刷新
            .Add 0, vbKeyF3, conMenu_View_FindNext    '查找下一个
            .Add 0, vbKeyF1, conMenu_Help_Help    '帮助
        End With

        '恢复及固定的一些菜单设置
        cbsMain.ActiveMenuBar.Title = "菜单"
        cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    ElseIf mbytMode = Mode_Show Then
        cbsMain.ActiveMenuBar.Visible = False
    End If

    For lngCount = 2 To cbsMain.count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
        cbsMain(lngCount).EnableDocking xtpFlagStretched + xtpFlagHideWrap
        If lngCount = 2 Then
            For Each objControl In cbsMain(lngCount).Controls
                If objControl.Type <> xtpControlLabel Then
                    If Not Between(objControl.ID, cmd_Edit_Undo, cmd_Edit_Paste) Then
                        objControl.Style = xtpButtonIconAndCaption
                    End If
                End If
            Next
        End If
    Next
End Sub

Private Sub cboTimes_Click()
    Dim strTmp As String
    Dim blnDo As Boolean
    
    If InStr(cboTimes.Text, cboTimes.Tag) = 0 Or cboTimes.Tag = "" Then
        cboTimes.Tag = marrTime(cboTimes.ListIndex)
        Call FuncShowAdvice(1)
    End If
    cboTimes.ToolTipText = cboTimes.Text
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objCombo As CommandBarComboBox
    Dim objComboBranch As CommandBarComboBox
    Dim vVersion As TYPE_PATH_VERSION
    Dim vBranch As TYPE_PATH_BRANCH
    Dim vArea As CONST_AREA, i As Long
    Dim strTmp As String

    If Control.ID <> 0 And Control.ID <> conMenu_View_FindNext Then
        If cbsMain.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If

    zlCommFun.ShowTipInfo 0, ""
    vArea = GetArea(vsPath.Row, vsPath.Col)

    Select Case Control.ID
    Case cmd_File_Save, cmd_File_SaveExit    '保存
        If Not CheckPathTable Then Exit Sub
        If Not SavePathTable() Then Exit Sub
        RaiseEvent DataChanged(mlng路径ID)
        If Control.ID = cmd_File_SaveExit Then Unload Me
    Case cmd_File_CopyFrom    '复制内容
        Call FuncVersionCopy
    Case cmd_File_PrintSetup    '打印设置
        Call zlPrintSet
    Case cmd_File_Print    '打印
        Call FuncPathTableOutput(1)
    Case cmd_File_Preview    '预览
        Call FuncPathTableOutput(2)
    Case cmd_File_ExportExcel    '导出Excel
        Call FuncPathTableOutput(3)
    Case cmd_File_ExportXML    '导出XML
        Call FuncExportToXML
    Case cmd_File_ImportXML    '导入XML
        Call FuncPathImportFromXML
        RaiseEvent DataChanged(mlng路径ID)
    Case cmd_Edit_Copy    '复制
        Call FuncClipboradCopy
    Case cmd_Edit_Paste    '粘贴
        Call FuncClipboradPaste
    Case cmd_Edit_Edit    '设置
        If vArea = Area_Step Then
            Call FuncStepEdit
        ElseIf vArea = Area_Item Then
            Call FuncItemEdit(Control)
        End If
    Case cmd_Edit_InsertBefore    '前面插入
        If vArea = Area_Category Then
            Call FuncCategoryInsert(-1)
        ElseIf vArea = Area_Step Then
            Call FuncStepInsert(-1)
        ElseIf vArea = Area_Item Then
            Call FuncItemInsert(-1)
        End If
    Case cmd_Edit_InsertAfter    '后面插入
        If vArea = Area_Category Then
            Call FuncCategoryInsert(1)
        ElseIf vArea = Area_Step Then
            Call FuncStepInsert(1)
        ElseIf vArea = Area_Item Then
            Call FuncItemInsert(1)
        End If
    Case cmd_Edit_InsertBranch    '插入分支
        Call FuncStepBranchInsert
    Case cmd_Edit_Modify    '修改
        vsPath.EditCell
    Case cmd_Edit_Delete    '删除
        If vArea = Area_Category Then
            Call FuncCategoryDelete
        ElseIf vArea = Area_Step Then
            Call FuncStepDelete
        ElseIf vArea = Area_Item Then
            Call FuncItemDelete
        End If
    Case cmd_Edit_EvalImport    '导入评估
        Call FuncEvaluateImport
    Case cmd_Edit_EvalStep    '阶段评估
        Call FuncEvaluateStep
    Case cmd_Edit_EvalStepCopy    '复制阶段评估
        Call FuncEvaluateStep(True)
    Case cmd_Edit_Version, cmd_Edit_Branch, cmd_View_Refresh  '版本,分支,刷新
        If Control.ID = cmd_Edit_Version Then
            Set objCombo = Control
        Else
            Set objCombo = cbsMain(cbsMain.count - 1).FindControl(, cmd_Edit_Version, True)

        End If
        If Control.ID = cmd_Edit_Branch Then
            Set objComboBranch = Control
        Else
            Set objComboBranch = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Branch, True)
        End If
        If objCombo.ListIndex > 0 And mblnChange Then
            If MsgBox("路径表内容已被更改尚未保存" & IIf(mvClipboard.Empty, "", ",并且将清空剪贴板") & "，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                'commandComboBox不能取消。
                If Control.ID = cmd_Edit_Version Then
                    objCombo.ListIndex = Val(objCombo.Category)
                ElseIf Control.ID = cmd_Edit_Branch Then
                    objComboBranch.ListIndex = Val(objComboBranch.Category)
                End If
                Exit Sub
            Else
                mvClipboard.Empty = True
            End If
        End If
        If objCombo.ListIndex = 0 Then
            mblnNewVersion = True
            mblnEditable = False
        Else
            vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))
            mblnNewVersion = vVersion.版本号 = 0
            mblnEditable = vVersion.审核时间 = Empty
        End If
        objCombo.Category = objCombo.ListIndex
        objComboBranch.Category = objComboBranch.ListIndex
        If objComboBranch.ListIndex <> 0 And objComboBranch.ItemData(objComboBranch.ListIndex) <> 0 Then vBranch = mcolBranch("_" & objComboBranch.ItemData(objComboBranch.ListIndex))
        Call LoadPathTable(objCombo, objComboBranch, vBranch.分支ID)
        Set objComboBranch = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Branch, True)
        If Not objComboBranch Is Nothing Then
            If objComboBranch.ListIndex > 0 Then
                If objComboBranch.ListIndex = 1 Then
                    strTmp = "从其他路径复制(&C)…"
                Else
                    strTmp = "从主路径或其他分支路径复制(&C)…"
                End If
                On Error Resume Next
                cbsMain.FindControl(, cmd_File_CopyFrom, True, True).Caption = strTmp
                On Error GoTo 0
            End If
        End If
        
        Set objControl = cbsMain.FindControl(, conMenu_View_Show, True)
        If Not objControl Is Nothing Then
            If objControl.Parameter = "隐藏" Then
                objControl.Parameter = "显示"
                Call cbsMain_Execute(objControl)
            End If
        End If
        
        mblnDiff = False
    Case cmd_Edit_VersionInfo    '版本信息
        Call FuncVersionEdit
    Case cmd_Edit_VersionNew    '添加版本
        Call FuncVersionNew
    Case cmd_Edit_VersionDel    '版本删除
        Call FuncVersionDelete
    Case cmd_Edit_BranchNew  '新增分支
        Call FuncBranchNew
    Case cmd_Edit_BranchDel  '删除当前分支
        Call FuncBranchDelete
    Case cmd_View_ToolBar_Button    '工具栏
        Me.cbsMain(2).Visible = Not Me.cbsMain(2).Visible
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text    '按钮文字
        For Each objControl In Me.cbsMain(2).Controls
            If objControl.Type <> xtpControlLabel Then
                If Not Between(objControl.ID, cmd_Edit_Undo, cmd_Edit_Paste) Then
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                Else
                    objControl.Style = xtpButtonIcon
                End If
            End If
        Next
        Me.cbsMain.RecalcLayout
    Case cmd_View_ToolBar_Size    '大图标
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StPath    '查看标准路径参考
        Call frmStPathList.ShowMe(Me, mstr疾病编码)
    Case cmd_View_StatusBar    '状态栏
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
    Case cmd_View_Find    '查找
        If Me.ActiveControl Is txtFind Then
            txtFind.SetFocus    '有时需要定位一下
            If txtFind.Text <> "" Then
                Call FuncFindItem
            End If
        Else
            txtFind.SetFocus
        End If
    Case conMenu_View_FindNext    '查找下一个
        If txtFind.Text = "" Then
            If txtFind.Visible And txtFind.Enabled Then txtFind.SetFocus
        Else
            Call FuncFindItem(True)
        End If
    Case conMenu_View_Difference    '显示差异/隐藏差异
        mblnDiff = Not mblnDiff
        Call ShowContrast(IIf(Control.Caption = "显示差异", 1, 2))
    Case conMenu_View_Contrast  '对比查看
        Call CompareAdviceItem
    Case cmd_Edit_ItemShow   '显示项目变动/隐藏项目变动
        If Control.Parameter = "显示" Then
            If CheckPathItem Then
                Control.Parameter = "隐藏"
                Control.Caption = "隐藏项目变动"
                mbytFunc = 2
            Else
                Exit Sub
            End If
        Else
            Control.Parameter = "显示"
            Control.Caption = "显示项目变动"
            mbytFunc = 0
        End If
        Call FuncResizeCenter
        Call FuncShowItemAdvice
        Call FuncSetAuditBtn
    Case conMenu_View_Show  '查看变动记录
        If Control.Parameter = "显示" Then
            Control.Parameter = "隐藏"
            Control.Caption = "隐藏变动记录"
            mbytFunc = 1
        Else
            Control.Parameter = "显示"
            Control.Caption = "查看变动记录"
            mbytFunc = 0
        End If
        Call FuncSetItemBackColor
        Call FuncResizeCenter
        Call FuncShowItemAdvice
    Case cmd_Help_Web_Home    'Web上的中联
        Call zlHomePage(Me.Hwnd)
    Case cmd_Help_Web_Forum    '中联论坛
        Call zlWebForum(Me.Hwnd)
    Case cmd_Help_Web_Mail    '发送反馈
        Call zlMailTo(Me.Hwnd)
    Case cmd_Help_About    '关于
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case cmd_Help_Help    '帮助
        Call ShowHelp(App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100))
    Case cmd_File_Exit    '退出
        Unload Me
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    
    Me.stcInfo.Left = lngLeft
    Me.stcInfo.Top = lngTop
    Me.stcInfo.Width = lngRight - lngLeft - fraSelect.Width
    
    Me.fraSelect.Left = Me.stcInfo.Left + Me.stcInfo.Width
    Me.fraSelect.Top = lngTop
    Me.fraSelect.Height = Me.stcInfo.Height - 15
    
    picCenter.Move lngLeft, lngTop + Me.stcInfo.Height, lngRight - lngLeft, lngBottom - lngTop - Me.stcInfo.Height
    Call FuncResizeCenter
    
    Me.Refresh
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objCombo As CommandBarComboBox
    Dim vVersion As TYPE_PATH_VERSION
    Dim vArea As CONST_AREA, strTemp As String
    Dim blnEnabled As Boolean, blnRefresh As Boolean
    Dim vStep As TYPE_PATH_STEP
    Dim vItem As TYPE_PATH_ITEM
    Dim blnAdjust As Boolean, i As Long

    vArea = GetArea(vsPath.Row, vsPath.Col)
    strTemp = Decode(vArea, Area_Category, "分类", Area_Step, "阶段", Area_Item, "项目")

    Select Case Control.ID
    Case cmd_File_Save, cmd_File_SaveExit    '保存
        Control.Enabled = mblnChange = True
    Case cmd_File_CopyFrom    '复制内容
        Control.Enabled = mblnEditable
    Case cmd_File_ExportXML    '导出XML
        If InStr(mstrPrivs, "导出XML") = 0 Then
            Control.Visible = False
        Else
            Set objCombo = cbsMain(cbsMain.count - 1).FindControl(, cmd_Edit_Version, True)
            If Not objCombo Is Nothing Then
                If objCombo.ListIndex > 0 Then
                    vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))
                End If
            End If
            Control.Enabled = vVersion.版本号 > 0
        End If
    Case cmd_File_ImportXML    '导入XML
        If InStr(mstrPrivs, "导入XML") = 0 Then
            Control.Visible = False
        Else
            Control.Enabled = mblnEditable
        End If
    Case cmd_Edit_Undo    '撤消
        Control.Visible = False
        Control.Enabled = mblnEditable
    Case cmd_Edit_Redo    '重做
        Control.Visible = False
        Control.Enabled = mblnEditable
    Case cmd_Edit_Copy    '复制
        Control.Enabled = mblnEditable And vArea = Area_Item And vsPath.Col = vsPath.ColSel
    Case cmd_Edit_Paste    '粘贴
        Control.Enabled = mblnEditable And Not mvClipboard.Empty
    Case cmd_Edit_Caption    '功能标题
        If Control.Caption <> strTemp & "：" Then
            Control.Caption = strTemp & "："
            cbsMain.RecalcLayout
        End If
    Case cmd_Edit_Edit    '设置
        If vArea = Area_Category Then
            Control.Visible = False
        Else
            Control.Visible = True

            If Control.Parent.Title <> "工具栏" Then
                If Control.Parent.Controls(Control.Index + 1).BeginGroup <> (vArea = Area_Category) Then
                    Control.Parent.Controls(Control.Index + 1).BeginGroup = (vArea = Area_Category)
                    blnRefresh = True
                End If
            End If
            If Control.Parent.Title <> "工具栏" Then
                If Control.Caption <> "设置" & strTemp & "(&E)" Then
                    Control.Caption = "设置" & strTemp & "(&E)"
                    blnRefresh = True
                End If
            End If
            If blnRefresh Then cbsMain.RecalcLayout

            If vArea = Area_Step Then
                Control.Enabled = mblnEditable And vsPath.ColSel = vsPath.Col
            ElseIf vArea = Area_Item Then
                blnEnabled = (vsPath.ColSel = vsPath.Col) And (vsPath.RowSel = vsPath.Row)

                '判断允许微调的条件：未停用的已审核版本，允许对医嘱或病历内容微调
                If blnEnabled And Not mblnEditable And mbytMode = Mode_Design Then
                    Set objCombo = cbsMain(cbsMain.count - 1).FindControl(, cmd_Edit_Version, True)
                    If Not objCombo Is Nothing Then
                        If objCombo.ListIndex > 0 Then
                            vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))
                        End If
                    End If
                    If vVersion.版本号 > 0 And vVersion.审核时间 <> Empty And vVersion.停用时间 = Empty Then
                        If TypeName(vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col)) <> "Empty" Then
                            vItem = vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col)
                            If vItem.医嘱IDs <> "" Or vItem.病历IDs <> "" Or vItem.新版病历IDs <> "" Then
                                blnAdjust = True
                            End If
                        End If
                    End If
                End If

                Control.Enabled = blnEnabled And (mblnEditable Or blnAdjust) And mbytFunc = 0

                If Control.Enabled And blnAdjust Then
                    Control.Parameter = "Adjust"
                Else
                    Control.Parameter = ""
                End If
            End If
        End If
    Case cmd_Edit_Insert    '插入
        If Control.Parent.Title <> "工具栏" Then
            If Control.Caption <> "插入" & strTemp & "(&I)" Then
                Control.Caption = "插入" & strTemp & "(&I)"
                cbsMain.RecalcLayout
            End If
        End If
        Control.Enabled = mblnEditable
    Case cmd_Edit_InsertBefore    '在前面插入
        If vArea = Area_Category Then
            Control.Enabled = mblnEditable And (vsPath.RowSel = vsPath.Row)
        ElseIf vArea = Area_Step Then
            Control.Enabled = mblnEditable And (vsPath.ColSel = vsPath.Col)
        ElseIf vArea = Area_Item Then
            Control.Enabled = mblnEditable And (vsPath.ColSel = vsPath.Col) And (vsPath.RowSel = vsPath.Row)
        End If
    Case cmd_Edit_InsertAfter    '在后面插入
        If vArea = Area_Category Then
            Control.Enabled = mblnEditable And (vsPath.RowSel = vsPath.Row)
        ElseIf vArea = Area_Step Then
            Control.Enabled = mblnEditable And (vsPath.ColSel = vsPath.Col)
        ElseIf vArea = Area_Item Then
            Control.Enabled = mblnEditable And (vsPath.ColSel = vsPath.Col) And (vsPath.RowSel = vsPath.Row)
        End If
    Case cmd_Edit_InsertBranch    '插入分支
        If vArea = Area_Step Then
            Control.Visible = True

            blnEnabled = vsPath.ColSel = vsPath.Col
            If blnEnabled Then
                '设置了的时间阶段才能插入分支
                blnEnabled = TypeName(vsPath.ColData(vsPath.Col)) <> "Empty"
            End If
            Control.Enabled = mblnEditable And blnEnabled
        Else
            Control.Visible = False
        End If
    Case cmd_Edit_Modify            '修改
        If strTemp = "分类" Then
            Control.Visible = True
        Else
            Control.Visible = False
        End If
        Control.Enabled = mblnEditable

    Case cmd_Edit_Delete    '删除
        If Control.Parent.Title <> "工具栏" Then
            If Control.Caption <> "删除" & strTemp & "(&D)" Then
                Control.Caption = "删除" & strTemp & "(&D)"
                cbsMain.RecalcLayout
            End If
        End If
        Control.Enabled = mblnEditable
    Case cmd_Edit_EvalImport    '导入评估
        If InStr(mstrPrivs, "评估表设计") = 0 Then
            Control.Visible = False
        Else
            Control.Enabled = mblnEditable
        End If
        If Control.Enabled Then
            Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Branch, True)
            If Not objCombo Is Nothing Then
                If objCombo.ListIndex > 0 Then
                    Control.Enabled = objCombo.ItemData(objCombo.ListIndex) = 0
                End If
            End If
        End If
    Case cmd_Edit_EvalStep    '阶段评估
        If InStr(mstrPrivs, "评估表设计") = 0 Then
            Control.Visible = False
        Else
            blnEnabled = mblnEditable And vsPath.Col >= vsPath.FixedCols + vsPath.FrozenCols And vsPath.Cols > 0
            If blnEnabled Then
                With vsPath
                    If TypeName(.ColData(.Col)) = "Empty" Then
                        blnEnabled = False
                    End If
                End With
            End If
            Control.Enabled = blnEnabled
        End If
    Case cmd_Edit_EvalStepCopy    '复制阶段评估
        If InStr(mstrPrivs, "评估表设计") = 0 Then
            Control.Visible = False
        Else
            blnEnabled = mblnEditable And vsPath.Col >= vsPath.FixedCols + vsPath.FrozenCols And vsPath.Cols > 0
            If blnEnabled Then
                With vsPath
                    If TypeName(.ColData(.Col)) <> "Empty" Then
                        vStep = .ColData(.Col)
                        If Not vStep.评估.条件集 Is Nothing And Not vStep.评估.指标集 Is Nothing Then
                            If vStep.评估.条件集.count > 0 Or vStep.评估.指标集.count > 0 Then
                                blnEnabled = False
                            End If
                        End If
                    Else
                        blnEnabled = False
                    End If
                End With
            End If
            Control.Enabled = blnEnabled
        End If
    Case cmd_Edit_VersionInfo    '版本信息
        Control.Enabled = mblnEditable
    Case cmd_Edit_VersionNew    '添加版本
        '没有未审核版本时(已保存或者尚未保存的)，可以添加新的版本
        Set objCombo = cbsMain(cbsMain.count - 1).FindControl(, cmd_Edit_Version, True)
        blnEnabled = Not objCombo Is Nothing
        If blnEnabled Then
            For i = 1 To objCombo.ListCount
                vVersion = mcolVersion("_" & objCombo.ItemData(i))
                If vVersion.审核时间 = Empty Then Exit For
            Next
            If i <= objCombo.ListCount Then blnEnabled = False
        End If
        Control.Enabled = blnEnabled
    Case cmd_Edit_VersionDel    '删除版本
        Control.Enabled = mblnEditable
        If Control.Enabled Then
            Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Branch, True)
            If Not objCombo Is Nothing Then
                If objCombo.ListIndex > 0 Then
                    Control.Enabled = objCombo.ItemData(objCombo.ListIndex) = 0
                End If
            End If
        End If
    Case cmd_Edit_BranchNew  '新增分支
        Control.Enabled = mblnEditable
        If Control.Enabled Then
            Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Branch, True)
            If Not objCombo Is Nothing Then
                If objCombo.ListIndex > 0 Then
                    Control.Enabled = objCombo.ItemData(objCombo.ListIndex) = 0
                End If
            End If
        End If
    Case cmd_Edit_BranchDel  '删除当前分支
        Control.Enabled = mblnEditable
        If Control.Enabled Then
            Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Branch, True)
            If Not objCombo Is Nothing Then
                If objCombo.ListIndex > 0 Then
                    Control.Enabled = objCombo.ItemData(objCombo.ListIndex) <> 0
                End If
            End If
        End If
    Case cmd_View_ToolBar_Button    '工具栏
        If cbsMain.count >= 2 Then
            Control.Checked = Me.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text    '图标文字
        If cbsMain.count >= 2 Then
            Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case cmd_View_ToolBar_Size    '大图标
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case cmd_View_StatusBar    '状态栏
        Control.Checked = Me.stbThis.Visible
    Case conMenu_View_Difference, conMenu_View_Contrast   '显示差异/隐藏差异 '对比查看
        Set objCombo = cbsMain(cbsMain.count - 1).FindControl(, cmd_Edit_Version, True)
        If Not objCombo Is Nothing Then
            If objCombo.ListIndex > 0 Then
                vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))
            End If
        End If
        If vVersion.版本号 > 1 And vVersion.审核时间 = Empty Then
            If Control.ID = conMenu_View_Difference Then
                Control.Enabled = True
                Control.Caption = IIf(mblnDiff, "隐藏差异", "显示差异")
            End If
            If Control.ID = conMenu_View_Contrast Then
                Control.Enabled = IIf(cbsMain.FindControl(, conMenu_View_Difference, True, True).Caption = "隐藏差异", True, False)
            End If
        Else
            Control.Enabled = False
        End If
    Case conMenu_View_Show  '查看变动记录
        Set objCombo = cbsMain(cbsMain.count - 1).FindControl(, cmd_Edit_Version, True)
        If Not objCombo Is Nothing Then
            If objCombo.ListIndex > 0 Then
                vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))
                If vVersion.审核时间 = Empty Then
                    Control.Enabled = False
                Else
                    Control.Enabled = True And (mbytFunc <> 2)
                End If
            End If
        End If
    Case cmd_Edit_ItemShow   '显示项目变动
        
        If InStr(mstrPrivs, "路径医嘱调整") = 0 Or mblnEditable Then
            Control.Visible = False
        Else
            Control.Visible = True
        End If
        If Control.Visible Then
            Control.Enabled = (mbytFunc <> 1)
        End If
    End Select
End Sub

Private Sub cmdCheck_Click(Index As Integer)
    Dim vItem As TYPE_PATH_ITEM
    Dim rsTmp As ADODB.Recordset
    Dim arrtmp As Variant
    Dim strDate As String
    Dim strSql As String
    Dim strTmp As String
    Dim i As Long
    
    On Error GoTo errH
    If TypeName(vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col)) = "Empty" Then Exit Sub
    vItem = vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col)
    strDate = "To_Date('" & cboTimes.Tag & "','YYYY-MM-DD HH24:MI:SS')"
    
    If Index = 0 Then
        If MsgBox("您确定项目""" & vItem.项目内容 & """的医嘱内容【审核通过】吗？", vbQuestion + vbOKCancel + vbDefaultButton2, gstrSysName) = vbOK Then
            strSql = "Zl_路径医嘱变动_Audit(" & vItem.ID & "," & strDate & ",0" & IIf(gbln双审核, "," & mbytAudit, "") & ")"
        Else
            Exit Sub
        End If
    Else
        If MsgBox("您确定项目""" & vItem.项目内容 & """的医嘱内容【审核不通过】吗？", vbQuestion + vbOKCancel + vbDefaultButton2, gstrSysName) = vbOK Then
            strSql = "Zl_路径医嘱变动_Audit(" & vItem.ID & "," & strDate & ",1" & IIf(gbln双审核, "," & mbytAudit, "") & ")"
        Else
            Exit Sub
        End If
    End If
    
    '数据提交
    If strSql <> "" Then
        Call zlDatabase.ExecuteProcedure(strSql, "路径医嘱审核")
    End If
    
    If Index = 0 Then
        strSql = "Select a.Id, a.相关id, a.序号, a.期效, a.诊疗项目id, a.收费细目id, a.医嘱内容, a.单次用量, a.总给予量, a.标本部位, a.检查方法, a.医生嘱托, a.执行频次, a.频率次数," & vbNewLine & _
                "       a.频率间隔, a.间隔单位, a.执行性质, a.执行标记, a.执行科室id, a.时间方案, a.是否缺省, a.是否备选, a.配方id, a.组合项目id" & vbNewLine & _
                ",C.类别,C.操作类型 " & vbNewLine & _
                "From 路径医嘱内容 A, 临床路径医嘱 B,诊疗项目目录 C" & vbNewLine & _
                "Where a.Id = b.医嘱内容id And A.诊疗项目ID(+)=C.ID And b.路径项目id = [1]" & vbNewLine & _
                "Order By a.序号"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, vItem.ID)
        strTmp = ""
        Do While Not rsTmp.EOF
            strTmp = strTmp & "," & rsTmp!ID
            mrsAdvice.AddNew
            mrsAdvice!ID = rsTmp!ID
            mrsAdvice!相关id = rsTmp!相关id
            mrsAdvice!是否缺省 = Val(rsTmp!是否缺省 & "")
            mrsAdvice!是否备选 = Val(rsTmp!是否备选 & "")
            mrsAdvice!序号 = rsTmp!序号
            mrsAdvice!期效 = rsTmp!期效
            mrsAdvice!诊疗项目ID = rsTmp!诊疗项目ID
            mrsAdvice!收费细目ID = rsTmp!收费细目ID
            mrsAdvice!医嘱内容 = rsTmp!医嘱内容
            mrsAdvice!单次用量 = rsTmp!单次用量
            mrsAdvice!总给予量 = rsTmp!总给予量
            mrsAdvice!标本部位 = rsTmp!标本部位
            mrsAdvice!检查方法 = rsTmp!检查方法
            mrsAdvice!医生嘱托 = rsTmp!医生嘱托
            mrsAdvice!执行频次 = rsTmp!执行频次
            mrsAdvice!频率次数 = rsTmp!频率次数
            mrsAdvice!频率间隔 = rsTmp!频率间隔
            mrsAdvice!间隔单位 = rsTmp!间隔单位
            mrsAdvice!执行性质 = rsTmp!执行性质
            mrsAdvice!执行科室ID = rsTmp!执行科室ID
            mrsAdvice!时间方案 = rsTmp!时间方案
            mrsAdvice!配方ID = rsTmp!配方ID
            mrsAdvice!组合项目ID = rsTmp!组合项目ID
            mrsAdvice!执行标记 = rsTmp!执行标记
            If gbln双审核 Then
                mrsAdvice!类别 = Nvl(rsTmp!类别, "")
                mrsAdvice!操作类型 = Nvl(rsTmp!操作类型, "")
            End If
            mrsAdvice.Update
            rsTmp.MoveNext
        Loop
            
        '清空缓存
        arrtmp = Split(vItem.医嘱IDs, ",")
        For i = LBound(arrtmp) To UBound(arrtmp)
            mrsAdvice.Filter = "ID =" & arrtmp(i)
            If mrsAdvice.RecordCount > 0 Then
                mrsAdvice.Delete
                mrsAdvice.Update
            End If
        Next
        mrsAdvice.Filter = ""
        vItem.医嘱IDs = Mid(strTmp, 2)
    End If
    strSql = "": strTmp = ""
    If gbln双审核 Then
        If mbytAudit = 1 Then
            strSql = " And  C.审核状态 In (2,3)"
            strTmp = " And  A.审核状态 In (2,3)"
        ElseIf mbytAudit = 2 Then
            strSql = " And (C.审核状态 = 3 OR NVL(C.审核状态,-1) =-1)"
            strTmp = " And (A.审核状态 = 3 OR NVL(A.审核状态,-1) =-1)"
        ElseIf mbytAudit = 3 Then
            strSql = " And C.审核状态 = 3"
            strTmp = " And A.审核状态 = 3"
        End If
    Else
        strSql = " And C.审核时间 Is Null"
        strTmp = " And A.审核时间 Is Null"
    End If

    strSql = "Select a.项目id, a.医嘱内容id" & vbNewLine & _
                "From 路径医嘱变动 A" & vbNewLine & _
                "Where a.项目id = [1] " & strTmp & " And" & vbNewLine & _
                "      a.操作时间 = (Select Max(操作时间) From 路径医嘱变动 C Where c.项目id = [1] " & strSql & ")" & vbNewLine & _
                "Order By a.项目id, a.医嘱内容id"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, vItem.ID)
    strTmp = ""
    Do While Not rsTmp.EOF
        strTmp = strTmp & "," & rsTmp!医嘱内容ID
        rsTmp.MoveNext
    Loop
    vItem.待审核医嘱IDs = Mid(strTmp, 2)
    vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col) = vItem
    If vItem.待审核医嘱IDs = "" Then
        vsPath.Cell(flexcpForeColor, vsPath.Row, vsPath.Col) = vbBlack
    End If
    '触发项目刷新
    Call vsPath_AfterRowColChange(vsPath.Row, vsPath.Col, vsPath.Row, vsPath.Col)
        
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim objPane As Pane

    If mbytMode = Mode_Show Then
        Call zlControl.FormSetCaption(Me, False, False)
        vsPath.Editable = flexEDNone
        vsPath.AllowSelection = False
        vsPath.HighLight = flexHighlightWithFocus
        vsPath.FocusRect = flexFocusLight
        Me.stbThis.Visible = False
    End If
    If gbln双审核 Then
        If InStr(";" & mstrPrivs & ";", ";审核;") = 0 And InStr(";" & mstrPrivs & ";", ";药剂科审核;") = 0 Then
            mbytAudit = 0
        ElseIf InStr(";" & mstrPrivs & ";", ";审核;") > 0 And InStr(";" & mstrPrivs & ";", ";药剂科审核;") > 0 Then
            mbytAudit = 3
        ElseIf InStr(";" & mstrPrivs & ";", ";药剂科审核;") > 0 Then
            mbytAudit = 1
        ElseIf InStr(mstrPrivs, ";审核;") > 0 Then
            mbytAudit = 2
        End If
    End If
    vsPath.Editable = flexEDNone
    optSelect(IX_ALL).Value = True
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True    '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    cbsMain.Icons = imgMain.Icons

    If mbytMode = Mode_Design Then
        Call RestoreWinState(Me, App.ProductName)
    End If
    Call MainDefCommandBar
    '---
    If mbytMode = Mode_Design Then
        Set mfrmVersion = New frmVersion
        Set mfrmPathStep = New frmPathStepEdit
        Set mfrmPathItem = New frmPathItemEdit
        Set mfrmEvalEdit = New frmEvaluateEdit
        Set mfrmAdviceContrast = New frmAdviceContrast
        Me.WindowState = vbMaximized    '窗体默认最大化
    End If

    '读取数据
    If mbytMode = Mode_Design Then
        vsPath.ExplorerBar = flexExSort
        Call LoadPathVersion
    Else
        vsPath.ExplorerBar = flexExNone
        mblnEditable = False
    End If

    mblnChange = False
    mvClipboard.Empty = True
    Erase mvClipboard.项目集
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mbytMode = Mode_Design And mblnChange Then
        If MsgBox("路径表内容已被更改尚未保存，确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
        mblnChange = False
    End If

    If Not mrsAdvice Is Nothing Then
        If mrsAdvice.State = 1 Then mrsAdvice.Close
        Set mrsAdvice = Nothing
    End If

    mvClipboard.Empty = True
    Erase mvClipboard.项目集

    If mbytMode = Mode_Design Then
        Unload mfrmVersion
        Set mfrmVersion = Nothing

        Unload mfrmPathStep
        Set mfrmPathStep = Nothing

        Unload mfrmPathItem
        Set mfrmPathItem = Nothing

        Unload mfrmEvalEdit
        Set mfrmEvalEdit = Nothing

        Unload mfrmAdviceContrast
        Set mfrmAdviceContrast = Nothing
    End If

    If mbytMode = Mode_Design Then
        Call SaveWinState(Me, App.ProductName)
    End If
End Sub

Private Function LoadPathVersion(Optional ByVal intVersion As Integer = -1) As Boolean
'功能：读取并加载显示临床路径的版本列表
'参数：intVersion=缺省定位版本
    Dim vVersion As TYPE_PATH_VERSION
    Dim objCombo As CommandBarComboBox
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim vBranch As TYPE_PATH_BRANCH
    Dim objComboBranch As CommandBarComboBox
    
    On Error GoTo errH
    
    Set objCombo = cbsMain(cbsMain.count - 1).FindControl(, cmd_Edit_Version, True)
    If objCombo Is Nothing Then Exit Function
    Set objComboBranch = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Branch, True)
    If objComboBranch Is Nothing Then Exit Function
    objCombo.Clear: vsPath.Rows = 0: vsPath.Cols = 0
    If mlng路径ID = 0 Then Exit Function
    
    Set mcolVersion = New Collection
        
    strSql = "Select A.分类,A.名称,B.版本号,B.标准住院日,B.标准费用,B.版本说明," & _
        " B.创建人,B.创建时间,B.审核人,B.审核时间,B.停用人,B.停用时间,a.性质" & _
        ",B.药剂科审核人,B.药剂科审核时间" & _
        " From 临床路径目录 A,临床路径版本 B" & _
        " Where A.ID=B.路径ID(+) And A.ID=[1]" & _
        " Order by B.版本号 Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng路径ID)
    
    Me.Tag = rsTmp!分类 & "-" & rsTmp!名称
    If mbytMode = Mode_Design Then
        Me.Caption = "临床路径设计 - " & rsTmp!名称
    End If
    
    Do While Not rsTmp.EOF
        If Not IsNull(rsTmp!版本号) Then
            objCombo.AddItem "第 " & rsTmp!版本号 & " 版，" & _
                "创建：" & rsTmp!创建人 & "/" & Format(rsTmp!创建时间, "yyyy-MM-dd HH:mm") & _
                IIf(Not IsNull(rsTmp!审核时间), "，审核：" & rsTmp!审核人 & "/" & Format(rsTmp!审核时间, "yyyy-MM-DD HH:mm"), "") & _
                IIf(Not IsNull(rsTmp!停用时间), "，停用：" & rsTmp!停用人 & "/" & Format(rsTmp!停用时间, "yyyy-MM-dd HH:mm"), "")
            objCombo.ItemData(objCombo.ListCount) = rsTmp!版本号
            If rsTmp!版本号 = intVersion Then
                objCombo.ListIndex = objCombo.ListCount
            End If
            
            vVersion.版本号 = rsTmp!版本号
            vVersion.标准住院日 = Nvl(rsTmp!标准住院日)
            vVersion.标准费用 = Nvl(rsTmp!标准费用)
            vVersion.版本说明 = Nvl(rsTmp!版本说明)
            vVersion.创建人 = rsTmp!创建人
            vVersion.创建时间 = rsTmp!创建时间
            vVersion.审核人 = Nvl(rsTmp!审核人)
            vVersion.审核时间 = IIf(IsNull(rsTmp!审核时间), Empty, rsTmp!审核时间)
            vVersion.药剂科审核人 = Nvl(rsTmp!药剂科审核人)
            vVersion.药剂科审核时间 = IIf(IsNull(rsTmp!药剂科审核时间), Empty, rsTmp!药剂科审核时间)
            vVersion.停用人 = Nvl(rsTmp!停用人)
            vVersion.停用时间 = IIf(IsNull(rsTmp!停用时间), Empty, rsTmp!停用时间)
            mlng性质 = Nvl(rsTmp!性质, 0)
            mcolVersion.Add vVersion, "_" & vVersion.版本号
        End If
        rsTmp.MoveNext
    Loop
    
    '索引从1开始，直接赋值不会引发Execute事件
    If objCombo.ListCount = 0 Then
        If mbytMode = Mode_Show Then
            cbsMain.RecalcLayout: Exit Function
        End If
        objCombo.AddItem "正在设计中……"
        
        vVersion.版本号 = 0
        vVersion.标准住院日 = ""
        vVersion.标准费用 = ""
        vVersion.版本说明 = ""
        vVersion.创建人 = ""
        vVersion.创建时间 = Empty
        vVersion.审核人 = ""
        vVersion.审核时间 = Empty
        vVersion.停用人 = ""
        vVersion.停用时间 = Empty
        mcolVersion.Add vVersion, "_0"

        objComboBranch.AddItem "主路径"
        objComboBranch.ItemData(objComboBranch.ListCount) = 0
        objComboBranch.ListIndex = objComboBranch.ListCount
        vBranch.版本号 = objCombo.ItemData(objCombo.ListIndex)
        vBranch.分支名称 = "主路径"
        Set mcolBranch = New Collection
        mcolBranch.Add vBranch, "_0"
    End If
    If objCombo.ListIndex = 0 Then objCombo.ListIndex = 1
    cbsMain.RecalcLayout

    Call cbsMain_Execute(objCombo)
    LoadPathVersion = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadPathTable(objCombo As CommandBarComboBox, objComboBranch As CommandBarComboBox, Optional ByVal lng分支ID As Long, Optional ByVal str分支名称 As String) As Boolean
'功能：根据所选择的路径版本，加载路径表相应的数据进行显示
    Dim vVersion As TYPE_PATH_VERSION
    Dim vStep As TYPE_PATH_STEP
    Dim vItem As TYPE_PATH_ITEM
    Dim vEvalMark As TYPE_PATH_EvalMark
    Dim vEvalCond As TYPE_PATH_EvalCond
    Dim vBranch As TYPE_PATH_BRANCH
    Dim vTmp As TYPE_PATH_BRANCH

    Dim colCols As New Collection
    Dim colRows As New Collection

    Dim rsTmp As ADODB.Recordset
    Dim rsClone As ADODB.Recordset
    Dim rsPathAdvice As ADODB.Recordset
    Dim rsPathEPR As ADODB.Recordset
    Dim rsEvalMark As ADODB.Recordset
    Dim rsEvalCond As ADODB.Recordset
    Dim strSql As String, strItems As String, strSqlItem As String
    Dim i As Long
    Dim lngRow As Long, lngCol As Long
    Dim blnBranch As Boolean
    
    On Error GoTo errH

    vsPath.Redraw = flexRDNone

    vsPath.Rows = 0: vsPath.Cols = 0
    If objCombo.ListIndex = 0 Then
        vsPath.Redraw = flexRDDirect
        Exit Function
    End If

    '分支路径
    If objComboBranch.ListIndex <> 0 Then
        vBranch = mcolBranch("_" & objComboBranch.ItemData(objComboBranch.ListIndex))
    End If
    '版本信息显示
    vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))
    stcInfo.Caption = _
    "标准住院日：" & IIf(vVersion.标准住院日 <> "", vVersion.标准住院日 & "天", "<未设定>") & _
                      "，标准费用：" & IIf(vVersion.标准费用 <> "", vVersion.标准费用 & "元", "<未设定>") & _
                      "，说明：" & IIf(vVersion.版本说明 <> "", vVersion.版本说明, "<无>") & IIf(vBranch.分支名称 = "主路径" Or vBranch.版本号 = 0, "", ("   分支名称：" & _
                                                                                                                                  IIf(vBranch.分支名称 <> "", vBranch.分支名称, "<未设定>") & _
                                                                                                                                  " 标准住院日：" & IIf(vBranch.标准住院日 <> "", vBranch.标准住院日 & "天", "<未设定>") & _
                                                                                                                                  "，标准费用：" & IIf(vBranch.标准费用 <> "", vBranch.标准费用 & "元", "<未设定>") & _
                                                                                                                                  "，说明：" & IIf(vBranch.说明 <> "", vBranch.说明, "<无>")))

    '路径表颜色设置
    If vVersion.停用时间 <> Empty Then
        vsPath.GridColor = Color_StopLine
        vsPath.SheetBorder = Color_StopLine
        vsPath.BackColorFrozen = Color_StopBack
    ElseIf vVersion.审核时间 <> Empty Then
        vsPath.GridColor = Color_AuditLine
        vsPath.SheetBorder = Color_AuditLine
        vsPath.BackColorFrozen = Color_AuditBack
    Else
        vsPath.GridColor = Color_NewLine
        vsPath.SheetBorder = Color_NewLine
        vsPath.BackColorFrozen = Color_NewBack
    End If

    '初始化当前版本医嘱内容表
    Call InitAdviceRecordset
    Set mvEvalImport.指标集 = New Collection
    Set mvEvalImport.条件集 = New Collection

    If vVersion.版本号 = 0 Then
        '空的路径表缺省样式
        With vsPath
            .Rows = 2 + 1: .FixedRows = 1: .FrozenRows = 1
            .Cols = 1 + 1: .FixedCols = 0: .FrozenCols = 1
            .ColWidth(-1) = COl_WIDTH_BASE: .ColWidth(0) = 1000
        End With
    Else
        '加载分支路径
        strSql = "Select a.id,a.名称 as 分支名称,a.版本号,b.名称 as 前一阶段名称,a.前一阶段id,a.创建人,a.创建时间,a.标准费用,a.标准住院日,a.说明" & vbNewLine & _
                 "From 临床路径分支 A, 临床路径阶段 B, 临床路径阶段 C" & vbNewLine & _
                 "Where a.前一阶段id = b.Id And b.父id = c.Id(+)" & vbNewLine & _
                 "And a.路径id = [1] And a.版本号 = [2]" & vbNewLine & _
                 "Order By Nvl(c.序号, b.序号), a.名称"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng路径ID, objCombo.ItemData(objCombo.ListIndex))
        Set mcolBranch = New Collection
        Set objComboBranch = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Branch, True)
        If Not objComboBranch Is Nothing Then
            objComboBranch.Clear
            '清除Vbranch
            vBranch = vTmp
            objComboBranch.AddItem "主路径"
            objComboBranch.ItemData(objComboBranch.ListCount) = 0
            objComboBranch.ListIndex = objComboBranch.ListCount
            vBranch.版本号 = objCombo.ItemData(objCombo.ListIndex)
            vBranch.分支名称 = "主路径"
            mcolBranch.Add vBranch, "_0"
            Do While Not rsTmp.EOF
                objComboBranch.AddItem "分支名称：" & rsTmp!分支名称 & " ，" & "前一阶段：" & rsTmp!前一阶段名称 & " ，" & _
                                       "创建：" & rsTmp!创建人 & "/" & Format(rsTmp!创建时间, "yyyy-MM-dd HH:mm")
                objComboBranch.ItemData(objComboBranch.ListCount) = rsTmp!ID
                vBranch.分支ID = Val(rsTmp!ID & "")
                vBranch.版本号 = Nvl(rsTmp!版本号)
                vBranch.标准住院日 = Nvl(rsTmp!标准住院日)
                vBranch.标准费用 = Nvl(rsTmp!标准费用)
                vBranch.说明 = Nvl(rsTmp!说明)
                vBranch.创建人 = rsTmp!创建人 & ""
                vBranch.创建时间 = rsTmp!创建时间
                vBranch.分支名称 = rsTmp!分支名称 & ""
                vBranch.前一阶段名称 = rsTmp!前一阶段名称 & ""
                vBranch.前一阶段ID = Val(rsTmp!前一阶段ID & "")
                mcolBranch.Add vBranch, "_" & vBranch.分支ID
                If lng分支ID = vBranch.分支ID Then
                    objComboBranch.ListIndex = objComboBranch.ListCount
                End If
                If str分支名称 <> "" Then
                    If str分支名称 = vBranch.分支名称 Then
                        objComboBranch.ListIndex = objComboBranch.ListCount
                    End If
                End If
                rsTmp.MoveNext
            Loop
            vBranch = mcolBranch("_" & objComboBranch.ItemData(objComboBranch.ListIndex))
            stcInfo.Caption = _
            "标准住院日：" & IIf(vVersion.标准住院日 <> "", vVersion.标准住院日 & "天", "<未设定>") & _
                              "，标准费用：" & IIf(vVersion.标准费用 <> "", vVersion.标准费用 & "元", "<未设定>") & _
                              "，说明：" & IIf(vVersion.版本说明 <> "", vVersion.版本说明, "<无>") & IIf(vBranch.分支名称 = "主路径" Or vBranch.版本号 = 0, "", ("   分支名称：" & _
                                                                                                                                          IIf(vBranch.分支名称 <> "", vBranch.分支名称, "<未设定>") & _
                                                                                                                                          " 标准住院日：" & IIf(vBranch.标准住院日 <> "", vBranch.标准住院日 & "天", "<未设定>") & _
                                                                                                                                          "，标准费用：" & IIf(vBranch.标准费用 <> "", vBranch.标准费用 & "元", "<未设定>") & _
                                                                                                                                          "，说明：" & IIf(vBranch.说明 <> "", vBranch.说明, "<无>")))
        End If
        '已保存的路径表样式
        With vsPath
            .Rows = 3: .FixedRows = 1: .FrozenRows = 2
            .Cols = 1: .FixedCols = 0: .FrozenCols = 1

            '评估数据读取
            strSql = _
            " Select A.评估类型,A.阶段ID,B.ID,B.序号,B.评估指标,B.指标类型,B.指标结果" & _
                     " From 临床路径评估 A,路径评估指标 B" & _
                     " Where A.ID=B.评估ID And A.路径ID=[1] And A.版本号=[2]" & _
                     IIf(vBranch.分支名称 = "主路径" Or vBranch.版本号 = 0, " And a.分支ID is null", " And A.分支ID=[3]") & _
                     " Order by A.评估类型,A.阶段ID,B.序号"
            Set rsEvalMark = zlDatabase.OpenSQLRecord(strSql, "LoadPathTable", mlng路径ID, objCombo.ItemData(objCombo.ListIndex), vBranch.分支ID)
            strSql = _
            " Select A.评估类型,A.阶段ID,B.指标ID,B.项目ID,B.关系式,B.条件值,B.条件组合" & _
                     " From 临床路径评估 A,路径评估条件 B" & _
                     " Where A.ID=B.评估ID And A.路径ID=[1] And A.版本号=[2]" & _
                     IIf(vBranch.分支名称 = "主路径" Or vBranch.版本号 = 0, " And a.分支ID is null", " And A.分支ID=[3]") & _
                     " Order by A.评估类型,A.阶段ID"
            Set rsEvalCond = zlDatabase.OpenSQLRecord(strSql, "LoadPathTable", mlng路径ID, objCombo.ItemData(objCombo.ListIndex), vBranch.分支ID)

            '0)导入评估
            rsEvalMark.Filter = "评估类型=1"
            Do While Not rsEvalMark.EOF
                vEvalMark.ID = rsEvalMark!ID
                vEvalMark.序号 = rsEvalMark!序号
                vEvalMark.评估指标 = rsEvalMark!评估指标
                vEvalMark.指标类型 = rsEvalMark!指标类型
                vEvalMark.指标结果 = rsEvalMark!指标结果
                mvEvalImport.指标集.Add vEvalMark
                rsEvalMark.MoveNext
            Loop
            rsEvalCond.Filter = "评估类型=1"
            Do While Not rsEvalCond.EOF
                vEvalCond.指标ID = Nvl(rsEvalCond!指标ID, 0)
                vEvalCond.项目ID = Nvl(rsEvalCond!项目ID, 0)
                vEvalCond.关系式 = rsEvalCond!关系式
                vEvalCond.条件值 = rsEvalCond!条件值
                vEvalCond.条件组合 = rsEvalCond!条件组合
                mvEvalImport.条件集.Add vEvalCond
                rsEvalCond.MoveNext
            Loop

            '1)时间阶段部分(一个阶段可以有多个分支路径，所以加Distinct)
            strSql = _
            " Select Distinct A.ID,Nvl(A.父ID,0) as 父ID,A.序号,b.序号 as 父ID序号,A.名称,A.开始天数,A.结束天数,A.标志,A.分类,A.说明,Sign(c.前一阶段id) as 存在分支" & _
                     " From 临床路径阶段 A,临床路径阶段 B,临床路径分支 C" & _
                     " Where a.父ID=b.ID(+) And a.id=c.前一阶段id(+) And A.路径ID=[1] And A.版本号=[2]" & _
                     IIf(vBranch.分支名称 = "主路径" Or vBranch.版本号 = 0, " And a.分支ID is null", " And A.分支ID=[3]") & _
                     " Order by NVL(B.序号,A.序号),NVL(b.序号,0),NVL(a.序号,0)"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "LoadPathTable", mlng路径ID, objCombo.ItemData(objCombo.ListIndex), vBranch.分支ID)

            blnBranch = False
            Set rsClone = rsTmp.Clone: rsTmp.Filter = "父ID=0"
            Do While Not rsTmp.EOF
                .Cols = .Cols + 1

                vStep.ID = rsTmp!ID
                vStep.父ID = 0
                vStep.序号 = rsTmp!序号
                vStep.名称 = rsTmp!名称
                vStep.开始天数 = Nvl(rsTmp!开始天数, 0)
                vStep.结束天数 = Nvl(rsTmp!结束天数, 0)
                vStep.标志 = Nvl(rsTmp!标志)
                vStep.分类 = Nvl(rsTmp!分类)
                vStep.说明 = Nvl(rsTmp!说明)
                vStep.存在分支 = Nvl(rsTmp!存在分支, 0) = 1

                '阶段评估
                Set vStep.评估.指标集 = New Collection
                rsEvalMark.Filter = "评估类型=2 And 阶段ID=" & vStep.ID
                Do While Not rsEvalMark.EOF
                    vEvalMark.ID = rsEvalMark!ID
                    vEvalMark.序号 = rsEvalMark!序号
                    vEvalMark.评估指标 = rsEvalMark!评估指标
                    vEvalMark.指标类型 = rsEvalMark!指标类型
                    vEvalMark.指标结果 = rsEvalMark!指标结果
                    vStep.评估.指标集.Add vEvalMark
                    rsEvalMark.MoveNext
                Loop
                Set vStep.评估.条件集 = New Collection
                rsEvalCond.Filter = "评估类型=2 And 阶段ID=" & vStep.ID
                Do While Not rsEvalCond.EOF
                    vEvalCond.指标ID = Nvl(rsEvalCond!指标ID, 0)
                    vEvalCond.项目ID = Nvl(rsEvalCond!项目ID, 0)
                    vEvalCond.关系式 = rsEvalCond!关系式
                    vEvalCond.条件值 = rsEvalCond!条件值
                    vEvalCond.条件组合 = rsEvalCond!条件组合
                    vStep.评估.条件集.Add vEvalCond
                    rsEvalCond.MoveNext
                Loop

                .ColData(.Cols - 1) = vStep
                '.Cell(flexcpText, .FixedRows, .Cols - 1, .FixedRows + .FrozenRows - 1, .Cols - 1) = vStep.名称
                '如果直接范围赋值，因为包含回车会自动识别为分隔符，而导致文字被切断
                .TextMatrix(.FixedRows, .Cols - 1) = vStep.名称
                .TextMatrix(.FixedRows + .FrozenRows - 1, .Cols - 1) = vStep.名称

                If mbytMode = Mode_Design Then
                    .TextMatrix(.FixedRows - 1, .Cols - 1) = "阶段评估…"
                    .Cell(flexcpFontBold, .FixedRows - 1, .Cols - 1) = False
                    If vStep.评估.指标集.count > 0 Or vStep.评估.条件集.count > 0 Then
                        .Cell(flexcpFontBold, .FixedRows - 1, .Cols - 1) = True
                    End If
                End If

                '用于快速定位该阶段的列号
                colCols.Add .Cols - 1, "_" & vStep.ID

                '加入备选分支
                rsClone.Filter = "父ID=" & rsTmp!ID
                If rsClone.EOF Then
                    If vStep.存在分支 Then
                        .Cell(flexcpPicture, .FixedRows, .Cols - 1) = ImgBranch.Picture
                        .Cell(flexcpPictureAlignment, .FixedRows, .Cols - 1) = 1
                    End If
                Else
                    Do While Not rsClone.EOF
                        .Cols = .Cols + 1

                        vStep.ID = rsClone!ID
                        vStep.父ID = rsClone!父ID
                        vStep.序号 = rsClone!序号
                        vStep.分类 = Nvl(rsClone!分类)
                        vStep.说明 = Nvl(rsClone!说明)
                        '以下应与该时间阶段相同
                        vStep.名称 = rsClone!名称
                        vStep.开始天数 = Nvl(rsClone!开始天数, 0)
                        vStep.结束天数 = Nvl(rsClone!结束天数, 0)
                        vStep.标志 = Nvl(rsClone!标志)

                        '阶段评估
                        Set vStep.评估.指标集 = New Collection
                        rsEvalMark.Filter = "评估类型=2 And 阶段ID=" & vStep.ID
                        Do While Not rsEvalMark.EOF
                            vEvalMark.ID = rsEvalMark!ID
                            vEvalMark.序号 = rsEvalMark!序号
                            vEvalMark.评估指标 = rsEvalMark!评估指标
                            vEvalMark.指标类型 = rsEvalMark!指标类型
                            vEvalMark.指标结果 = rsEvalMark!指标结果
                            vStep.评估.指标集.Add vEvalMark
                            rsEvalMark.MoveNext
                        Loop
                        Set vStep.评估.条件集 = New Collection
                        rsEvalCond.Filter = "评估类型=2 And 阶段ID=" & vStep.ID
                        Do While Not rsEvalCond.EOF
                            vEvalCond.指标ID = Nvl(rsEvalCond!指标ID, 0)
                            vEvalCond.项目ID = Nvl(rsEvalCond!项目ID, 0)
                            vEvalCond.关系式 = rsEvalCond!关系式
                            vEvalCond.条件值 = rsEvalCond!条件值
                            vEvalCond.条件组合 = rsEvalCond!条件组合
                            vStep.评估.条件集.Add vEvalCond
                            rsEvalCond.MoveNext
                        Loop

                        .ColData(.Cols - 1) = vStep
                        .TextMatrix(.FixedRows, .Cols - 1) = vStep.名称
                        .TextMatrix(.FixedRows + .FrozenRows - 1, .Cols - 1) = IIf(vStep.说明 = "", "备用分支" & vStep.序号, vStep.说明) & IIf(vStep.分类 = "", "", ",") & vStep.分类
                        If vStep.序号 = 1 Then
                            .TextMatrix(.FixedRows + .FrozenRows - 1, .Cols - 2) = "缺省分支"
                        End If

                        If vStep.存在分支 Then
                            .Cell(flexcpPicture, .FixedRows + .FrozenRows - 1, .Cols - 2) = ImgBranch.Picture
                            .Cell(flexcpPictureAlignment, .FixedRows + .FrozenRows - 1, .Cols - 2) = 3
                        End If
                        If rsClone!存在分支 = 1 Then
                            .Cell(flexcpPicture, .FixedRows + .FrozenRows - 1, .Cols - 1) = ImgBranch.Picture
                            .Cell(flexcpPictureAlignment, .FixedRows + .FrozenRows - 1, .Cols - 1) = 3
                        End If

                        If mbytMode = Mode_Design Then
                            .TextMatrix(.FixedRows - 1, .Cols - 1) = "阶段评估…"
                            .Cell(flexcpFontBold, .FixedRows - 1, .Cols - 1) = False
                            If vStep.评估.指标集.count > 0 And vStep.评估.条件集.count > 0 Then
                                .Cell(flexcpFontBold, .FixedRows - 1, .Cols - 1) = True
                            End If
                        End If

                        '用于快速定位该阶段的列号
                        colCols.Add .Cols - 1, "_" & vStep.ID

                        blnBranch = True
                        rsClone.MoveNext
                    Loop
                End If

                rsTmp.MoveNext
            Loop
            If Not blnBranch Then .FrozenRows = 1: .RemoveItem .FixedRows + .FrozenRows

            '2)分类部分
            strSql = _
            " Select A.序号,A.分类,Max(个数) as 个数" & _
                     " From (" & _
                     "   Select A.序号,A.名称 as 分类,Nvl(B.阶段ID,0),Count(Nvl(B.项目序号,0)) as 个数" & _
                     "   From 临床路径分类 A,临床路径项目 B" & _
                     "   Where A.路径ID=[1] And A.版本号=[2]" & _
                     IIf(vBranch.分支名称 = "主路径" Or vBranch.版本号 = 0, " And a.分支ID is null", " And A.分支ID=[3]") & _
                     "       And A.名称=B.分类(+) And B.路径ID(+)=[1] And B.版本号(+)=[2]" & _
                     "   Group by A.序号,A.名称,Nvl(B.阶段ID,0)" & _
                     "   ) A" & _
                     " Group by A.序号,A.分类" & _
                     " Order by A.序号"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "LoadPathTable", mlng路径ID, objCombo.ItemData(objCombo.ListIndex), vBranch.分支ID)
            
            Do While Not rsTmp.EOF
                '序号只用于排序，保存时重新生成
                .Rows = .Rows + rsTmp!个数
                .Cell(flexcpText, .Rows - rsTmp!个数, .FixedCols, .Rows - 1, .FixedCols) = rsTmp!分类

                '用于快速定位该分类的起始行号
                colRows.Add .Rows - rsTmp!个数, "_" & rsTmp!分类

                rsTmp.MoveNext
            Loop

            '3)项目部分
            '--医嘱定义内容集
            strSql = _
            " Select Distinct A.ID,A.相关ID,A.序号,A.期效,A.诊疗项目ID,A.收费细目ID,D.类别,D.操作类型," & _
                     " A.医嘱内容,A.单次用量,A.总给予量,A.标本部位,A.检查方法,A.医生嘱托," & _
                     " A.执行频次,A.频率次数,A.频率间隔,A.间隔单位,A.执行性质,A.执行标记,A.执行科室ID,A.时间方案,A.是否缺省,A.是否备选,A.配方ID,A.组合项目ID" & _
                     " From 路径医嘱内容 A,临床路径医嘱 B,诊疗项目目录 D,临床路径项目 C " & _
                     " Where A.ID=B.医嘱内容ID And B.路径项目ID=C.ID And A.诊疗项目ID =D.ID(+) And C.路径ID=[1] And C.版本号=[2]" & _
                     IIf(vBranch.分支名称 = "主路径" Or vBranch.版本号 = 0, " And c.分支ID is null", " And C.分支ID=[3]") & _
                     " Order by A.序号,A.ID"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "LoadPathTable", mlng路径ID, objCombo.ItemData(objCombo.ListIndex), vBranch.分支ID)
            Do While Not rsTmp.EOF
                mrsAdvice.AddNew
                mrsAdvice!ID = rsTmp!ID
                mrsAdvice!相关id = rsTmp!相关id
                mrsAdvice!是否缺省 = Val(rsTmp!是否缺省 & "")
                mrsAdvice!是否备选 = Val(rsTmp!是否备选 & "")
                mrsAdvice!序号 = rsTmp!序号
                mrsAdvice!期效 = rsTmp!期效
                mrsAdvice!诊疗项目ID = rsTmp!诊疗项目ID
                mrsAdvice!收费细目ID = rsTmp!收费细目ID
                mrsAdvice!医嘱内容 = rsTmp!医嘱内容
                mrsAdvice!单次用量 = rsTmp!单次用量
                mrsAdvice!总给予量 = rsTmp!总给予量
                mrsAdvice!标本部位 = rsTmp!标本部位
                mrsAdvice!检查方法 = rsTmp!检查方法
                mrsAdvice!医生嘱托 = rsTmp!医生嘱托
                mrsAdvice!执行频次 = rsTmp!执行频次
                mrsAdvice!频率次数 = rsTmp!频率次数
                mrsAdvice!频率间隔 = rsTmp!频率间隔
                mrsAdvice!间隔单位 = rsTmp!间隔单位
                mrsAdvice!执行性质 = rsTmp!执行性质
                mrsAdvice!执行科室ID = rsTmp!执行科室ID
                mrsAdvice!时间方案 = rsTmp!时间方案
                mrsAdvice!配方ID = rsTmp!配方ID
                mrsAdvice!组合项目ID = rsTmp!组合项目ID
                mrsAdvice!执行标记 = rsTmp!执行标记
                If gbln双审核 Then
                    mrsAdvice!类别 = Nvl(rsTmp!类别, "")
                    mrsAdvice!操作类型 = Nvl(rsTmp!操作类型, "")
                End If
                mrsAdvice.Update

                rsTmp.MoveNext
            Loop
            '--医嘱对应关系
            strSql = "Select Distinct A.路径项目ID,A.医嘱内容ID" & _
                     " From 临床路径医嘱 A,临床路径项目 B Where A.路径项目ID=B.ID And B.路径ID=[1] And B.版本号=[2]" & _
                     IIf(vBranch.分支名称 = "主路径" Or vBranch.版本号 = 0, " And b.分支ID is null", " And B.分支ID=[3]") & _
                     " Order by 路径项目ID,医嘱内容ID"
            Set rsPathAdvice = zlDatabase.OpenSQLRecord(strSql, "LoadPathTable", mlng路径ID, objCombo.ItemData(objCombo.ListIndex), vBranch.分支ID)
            '--病历对应关系
            strSql = "Select Distinct A.项目ID,A.文件ID,A.原型ID " & _
                     " From 临床路径病历 A,临床路径项目 B Where A.项目ID=B.ID And B.路径ID=[1] And B.版本号=[2] " & _
                     IIf(vBranch.分支名称 = "主路径" Or vBranch.版本号 = 0, " And b.分支ID is null", " And B.分支ID=[3]") & _
                     " Order by 项目ID,文件ID,原型ID"
            Set rsPathEPR = zlDatabase.OpenSQLRecord(strSql, "LoadPathTable", mlng路径ID, objCombo.ItemData(objCombo.ListIndex), vBranch.分支ID)


            '--路径项目
            Set mcolItemRowCol = New Collection
            strSql = _
            " Select a.ID,a.阶段ID,a.分类,a.项目序号,a.项目内容,a.执行方式,a.执行者,生成者,a.项目结果,a.图标ID,a.内容要求,A.导入参考,Nvl(A.导入结果,1) 导入结果" & _
                     " From 临床路径项目 A,临床路径阶段 B,临床路径阶段 C Where a.阶段ID=b.ID And b.父ID=c.ID(+) And a.路径ID=[1] And a.版本号=[2] " & _
                     IIf(vBranch.分支名称 = "主路径" Or vBranch.版本号 = 0, " And B.分支ID is Null And  A.分支ID is null", " And B.分支ID=[3] And A.分支ID=[3]") & _
                     " Order by NVL(c.序号,b.序号),NVL(c.序号,0),a.分类,a.项目序号"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "LoadPathTable", mlng路径ID, objCombo.ItemData(objCombo.ListIndex), vBranch.分支ID)
            Do While Not rsTmp.EOF
                vItem.ID = rsTmp!ID
                vItem.项目序号 = rsTmp!项目序号
                vItem.项目内容 = rsTmp!项目内容
                vItem.执行方式 = Nvl(rsTmp!执行方式, 0)
                vItem.执行者 = Nvl(rsTmp!执行者, 0)
                vItem.项目结果 = Nvl(rsTmp!项目结果)
                vItem.图标ID = Nvl(rsTmp!图标ID, 0)
                vItem.内容要求 = Val("" & rsTmp!内容要求)
                vItem.导入参考 = Nvl(rsTmp!导入参考)
                vItem.导入结果 = Nvl(rsTmp!导入结果)
                vItem.生成者 = Nvl(rsTmp!生成者, 1)  'NULL缺省为医生
                '关联的医嘱
                rsPathAdvice.Filter = "路径项目ID=" & rsTmp!ID
                vItem.医嘱IDs = ""
                Do While Not rsPathAdvice.EOF
                    vItem.医嘱IDs = vItem.医嘱IDs & "," & rsPathAdvice!医嘱内容ID
                    rsPathAdvice.MoveNext
                Loop
                vItem.医嘱IDs = Mid(vItem.医嘱IDs, 2)
                If vVersion.审核时间 <> Empty And vVersion.停用时间 = Empty Then
                    vItem.原医嘱IDs = vItem.医嘱IDs
                    If InStr(mstrPrivs, "路径医嘱调整") > 0 Then strItems = strItems & "," & vItem.ID
                End If
                
                '关联的病历
                rsPathEPR.Filter = "项目ID=" & rsTmp!ID
                vItem.病历IDs = "": vItem.新版病历IDs = ""
                Do While Not rsPathEPR.EOF
                    If rsPathEPR!文件ID & "" <> "" Then
                        vItem.病历IDs = vItem.病历IDs & "," & rsPathEPR!文件ID
                    Else
                        vItem.新版病历IDs = vItem.新版病历IDs & "," & rsPathEPR!原型ID
                    End If
                    rsPathEPR.MoveNext
                Loop
                vItem.病历IDs = Mid(vItem.病历IDs, 2)
                vItem.新版病历IDs = Mid(vItem.新版病历IDs, 2)
                '定位和显示
                lngCol = colCols("_" & rsTmp!阶段ID)
                lngRow = colRows("_" & rsTmp!分类)

                Do While .TextMatrix(lngRow, lngCol) <> ""
                    lngRow = lngRow + 1
                Loop
                If vItem.图标ID <> 0 Then
                    Set .Cell(flexcpPicture, lngRow, lngCol) = GetPathIcon(vItem.图标ID)
                    .Cell(flexcpPictureAlignment, lngRow, lngCol) = 1
                End If
                .TextMatrix(lngRow, lngCol) = vItem.项目内容
                If vItem.医嘱IDs <> "" Or vItem.病历IDs <> "" Or vItem.新版病历IDs <> "" Then
                    .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow, lngCol) & "…"
                End If
                .Cell(flexcpData, lngRow, lngCol) = vItem
                If vItem.导入结果 <> 1 Then
                    .Cell(flexcpBackColor, lngRow, lngCol) = &HE1FFE1
                End If
                
                mcolItemRowCol.Add lngRow & "," & lngCol, "_" & vItem.ID '显示差异时，快速定位到行列，方便设置差异单元格的背景色
                rsTmp.MoveNext
            Loop
            
            '待审核医嘱
            strSql = ""
            strSql = " And NVL(C.审核状态,-1) Not In (0,1)"
            strSqlItem = " And NVL(A.审核状态,-1) Not In (0,1)"
            If strItems <> "" Then
                strItems = Mid(strItems, 2)
                strSql = "Select /*+cardinality(b,10)*/" & vbNewLine & _
                    " a.项目id, a.医嘱内容id, Decode(审核状态,NULL,2,3,3,2,1,审核状态) as 审核状态, 审核时间" & vbNewLine & _
                    "From 路径医嘱变动 A, Table(f_Num2list([1])) B" & vbNewLine & _
                    "Where a.项目id = b.Column_Value " & strSqlItem & vbNewLine & _
                    "      And a.操作时间 = (Select Max(操作时间) From 路径医嘱变动 C Where c.项目id = a.项目id " & strSql & ")" & vbNewLine & _
                    "Order By a.项目id, a.医嘱内容id"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "LoadPathTable", strItems)
                For lngRow = .FixedRows + .FrozenRows To .Rows - 1
                    For lngCol = .FixedCols + .FrozenCols To .Cols - 1
                        If TypeName(.Cell(flexcpData, lngRow, lngCol)) = TypeName(vItem) Then
                            vItem = .Cell(flexcpData, lngRow, lngCol)
                            If vItem.医嘱IDs <> "" Then
                                rsTmp.Filter = "项目ID=" & vItem.ID
                                If gbln双审核 And Not rsTmp.EOF Then vItem.审核状态 = rsTmp!审核状态
                                Do While Not rsTmp.EOF
                                    vItem.待审核医嘱IDs = vItem.待审核医嘱IDs & "," & rsTmp!医嘱内容ID
                                    rsTmp.MoveNext
                                Loop
                                vItem.待审核医嘱IDs = Mid(vItem.待审核医嘱IDs, 2)
                                .Cell(flexcpData, lngRow, lngCol) = vItem
                                If vItem.待审核医嘱IDs <> "" Then
                                    .Cell(flexcpForeColor, lngRow, lngCol) = Color_NeedAuditFore
                                End If
                            End If
                        End If
                    Next
                Next
            End If
            '---
            For i = .FixedCols + .FrozenCols To .Cols - 1
                .ColWidth(i) = COl_WIDTH_BASE
            Next
        End With
    End If
    '生成者显示
    If Val(optSelect(IX_ALL).Tag) <> IX_ALL Then
        Call FuncShowItemBySendor
    End If
    vsPath.Redraw = flexRDDirect
    vsPath.AutoSize vsPath.FixedCols, vsPath.Cols - 1, , 45    '在要Draw之后才生效
    Call SetTableCommonStyle(True)
    vsPath.Row = vsPath.FixedRows + vsPath.FrozenRows
    vsPath.Col = vsPath.FixedCols + vsPath.FrozenCols
    If mbytMode = Mode_Design And Visible Then vsPath.SetFocus

    mstrDelStepIDs = ""
    mstrDelItemIDs = ""
    mblnChange = False

    LoadPathTable = True
    Exit Function
errH:
    vsPath.Redraw = flexRDDirect
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitAdviceRecordset()
    If Not mrsAdvice Is Nothing Then
        If mrsAdvice.State = 1 Then mrsAdvice.Close
    End If
    Set mrsAdvice = New ADODB.Recordset
    
    mrsAdvice.Fields.Append "ID", adBigInt
    mrsAdvice.Fields.Append "是否缺省", adSmallInt
    mrsAdvice.Fields.Append "是否备选", adSmallInt
    mrsAdvice.Fields.Append "相关ID", adBigInt, , adFldIsNullable
    mrsAdvice.Fields.Append "序号", adBigInt
    mrsAdvice.Fields.Append "期效", adSmallInt
    mrsAdvice.Fields.Append "诊疗项目ID", adBigInt, , adFldIsNullable
    mrsAdvice.Fields.Append "收费细目ID", adBigInt, , adFldIsNullable
    mrsAdvice.Fields.Append "医嘱内容", adVarChar, 1000, adFldIsNullable
    mrsAdvice.Fields.Append "单次用量", adSingle, , adFldIsNullable
    mrsAdvice.Fields.Append "总给予量", adSingle, , adFldIsNullable
    mrsAdvice.Fields.Append "标本部位", adVarChar, 100, adFldIsNullable
    mrsAdvice.Fields.Append "检查方法", adVarChar, 100, adFldIsNullable
    mrsAdvice.Fields.Append "医生嘱托", adVarChar, 1000, adFldIsNullable
    mrsAdvice.Fields.Append "执行频次", adVarChar, 100, adFldIsNullable
    mrsAdvice.Fields.Append "频率次数", adSmallInt, , adFldIsNullable
    mrsAdvice.Fields.Append "频率间隔", adSmallInt, , adFldIsNullable
    mrsAdvice.Fields.Append "间隔单位", adVarChar, 10, adFldIsNullable
    mrsAdvice.Fields.Append "执行性质", adSmallInt
    mrsAdvice.Fields.Append "执行科室ID", adBigInt, , adFldIsNullable
    mrsAdvice.Fields.Append "时间方案", adVarChar, 100, adFldIsNullable
    mrsAdvice.Fields.Append "配方ID", adBigInt, , adFldIsNullable
    mrsAdvice.Fields.Append "组合项目ID", adBigInt, , adFldIsNullable
    mrsAdvice.Fields.Append "执行标记", adSingle, , adFldIsNullable
    mrsAdvice.Fields.Append "待审核", adSingle, 1, adFldIsNullable
    mrsAdvice.Fields.Append "项目ID", adBigInt, , adFldIsNullable   '路径医嘱变动时用
    If gbln双审核 Then
        mrsAdvice.Fields.Append "类别", adVarChar, 1, adFldIsNullable
        mrsAdvice.Fields.Append "操作类型", adVarChar, 20, adFldIsNullable
        ''1-只修改了药品只需药剂科审核;2-修改了医嘱未修改药品只需要医务科审核;3-需要药剂科和医务科同审
        mrsAdvice.Fields.Append "审核状态", adInteger, 1, adFldIsNullable
    End If
    mrsAdvice.CursorLocation = adUseClient
    mrsAdvice.LockType = adLockOptimistic
    mrsAdvice.CursorType = adOpenStatic
    mrsAdvice.Open
End Sub

Private Sub SetTableCommonStyle(Optional ByVal blnKeep As Boolean)
'功能：对路径表格进行一些统一的样式设置
'功能：对一些表现属性保持不变
    Dim vRedraw As RedrawSettings
    Dim i As Long

    With vsPath
        vRedraw = .Redraw
        If Not blnKeep Then
            .RowHeight(-1) = ROW_HEIGHT_MIN '调整有行高
        Else
            For i = .FixedRows To .Rows - 1
                If .RowHeight(i) < ROW_HEIGHT_MIN Then
                    .RowHeight(i) = ROW_HEIGHT_MIN
                End If
            Next
        End If
        '列宽辅助行
        If mbytMode = Mode_Design Then
            .RowHeight(0) = ROW_HEIGHT_MIN
        Else
            .RowHeight(0) = 150
        End If
        .RowHeight(1) = 650 '时间阶段显示行
        
        .Cell(flexcpText, .FixedRows, .FixedCols, .FixedRows + .FrozenRows - 1, .FixedCols + .FrozenCols - 1) = " 时间阶段 "
        .Cell(flexcpAlignment, 0, 0, .FixedRows + .FrozenRows - 1, .Cols - 1) = 4 '横表头
        .Cell(flexcpAlignment, .FixedRows + .FrozenRows, 0, .Rows - 1, .FixedCols + .FrozenCols - 1) = 4 '竖表头
        .Cell(flexcpAlignment, .FixedRows + .FrozenRows, .FixedCols + .FrozenCols, .Rows - 1, .Cols - 1) = 1 '项目数据部分
        
        .MergeCol(-1) = True
        .MergeRow(.FixedRows) = True
        
    
        '对多行时间阶段表头，空的阶段列设置为合并效果
        If .FrozenRows > 1 Then
            For i = .FixedCols + .FrozenCols To .Cols - 1
                If TypeName(.ColData(i)) = "Empty" Then
                    .Cell(flexcpText, .FixedRows, i, .FixedRows + .FrozenRows - 1, i) = Space((i Mod 2) + 1)
                End If
            Next
        End If
        
        If Not blnKeep Then
            .Row = .FixedRows + .FrozenRows
            .Col = .FixedCols + .FrozenCols
        End If
        
        .Redraw = vRedraw
    End With
End Sub

Private Function GetArea(ByVal lngRow As Long, ByVal lngCol As Long) As CONST_AREA
'功能：获取指定行列在哪一块区域
    With vsPath
        If lngRow = -1 Or lngCol = -1 Then
            GetArea = -1
        ElseIf lngRow <= .FixedRows - 1 Or lngCol <= .FixedCols - 1 Then
            GetArea = -1
        ElseIf lngCol >= .FixedCols And lngCol <= .FixedCols + .FrozenCols - 1 _
            And lngRow >= .FixedRows And lngRow <= .FixedRows + .FrozenRows - 1 Then
            GetArea = Area_Cross
        ElseIf lngCol >= .FixedCols And lngCol <= .FixedCols + .FrozenCols - 1 Then
            GetArea = Area_Category
        ElseIf lngRow >= .FixedRows And lngRow <= .FixedRows + .FrozenRows - 1 Then
            GetArea = Area_Step
        Else
            GetArea = Area_Item
        End If
    End With
End Function

Private Sub fraSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If fraSplit.Top + Y < 100 Or fraSplit.Top + Y > picCenter.Height - 100 Then Exit Sub
        fraSplit.Top = fraSplit.Top + Y
        vsPath.Height = vsPath.Height + Y
        
        picBottom.Top = picBottom.Top + Y
        picBottom.Height = picBottom.Height - Y
        
        ucAdvice(0).Height = ucAdvice(0).Height - Y
        ucAdvice(1).Height = ucAdvice(1).Height - Y
    End If
End Sub

Private Sub fraSplit2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If fraSplit2.Left + X < (picBottom.Width / 10) * 1 Or fraSplit2.Left + X > (picBottom.Width / 10) * 9 Then Exit Sub
        fraSplit2.Left = fraSplit2.Left + X
        ucAdvice(0).Width = ucAdvice(0).Width + X
        
        lblChange.Left = lblChange.Left + X
        cboTimes.Left = cboTimes.Left + X
        If cmdCheck(0).Visible Then cmdCheck(0).Left = cmdCheck(0).Left + X
        If cmdCheck(1).Visible Then cmdCheck(1).Left = cmdCheck(1).Left + X
        ucAdvice(1).Left = ucAdvice(1).Left + X
        ucAdvice(1).Width = ucAdvice(1).Width - X
    End If
End Sub

Private Sub mfrmAdviceContrast_MovePathItemFocus(ByVal lngItemID As Long)
'功能:根据项目ID,让项目获得焦点
'参数：lngItemID:项目ID
    Dim strTmp As String
    Dim lngRow As Long, lngCol As Long

    strTmp = mcolItemRowCol("_" & lngItemID)
    lngRow = Split(strTmp, ",")(0)
    lngCol = Split(strTmp, ",")(1)
    With vsPath
        .Row = lngRow
        .Col = lngCol
        '对比查看时，实时更新当前项目内容
        Call mfrmAdviceContrast.SetNoteInfo(.TextMatrix(.Row, .Col))
    End With

End Sub

Private Sub mfrmEvalEdit_CheckDataValid(EvalInfo As TYPE_PATH_EVAL, EvalType As Integer, Cancel As Boolean)
    '###
End Sub

Private Sub mfrmPathItem_CheckDataValid(PathItem As TYPE_PATH_ITEM, Cancel As Boolean)
    '###
End Sub

Private Sub mfrmVersion_CalcPathCost(CostMin As Currency, CostMax As Currency, lng分支ID As Long)
'功能：估算路径费用
    Dim objCombo As CommandBarComboBox
    Dim vVersion As TYPE_PATH_VERSION
    Dim vBranch As TYPE_PATH_BRANCH
    Dim rsTmp As ADODB.Recordset
    Dim curCostMin As Currency, curCostMax As Currency
    Dim strSql As String, intDay As Integer
    Dim intDayMin As Integer, intDayMax As Integer
    
    If mblnChange Then
        MsgBox "路径表内容尚未保存，请先保存才能进行估算。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Set objCombo = cbsMain(cbsMain.count - 1).FindControl(, cmd_Edit_Version, True)
    If Not objCombo Is Nothing Then
        If objCombo.ListIndex > 0 Then
            vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))
        End If
    End If
    If vVersion.版本号 = 0 Then Exit Sub
    If lng分支ID <> 0 Then
        vBranch = mcolBranch("_" & lng分支ID)
        If InStr(vBranch.标准住院日, "-") > 0 Then
            intDayMin = Val(Split(vBranch.标准住院日, "-")(0))
            intDayMax = Val(Split(vBranch.标准住院日, "-")(1))
        Else
            intDayMin = Val(vBranch.标准住院日)
            intDayMax = intDayMin
        End If
    Else
        If InStr(vVersion.标准住院日, "-") > 0 Then
            intDayMin = Val(Split(vVersion.标准住院日, "-")(0))
            intDayMax = Val(Split(vVersion.标准住院日, "-")(1))
        Else
            intDayMin = Val(vVersion.标准住院日)
            intDayMax = intDayMin
        End If
    End If
    
    Screen.MousePointer = 11
    On Error GoTo errH
    For intDay = 1 To intDayMax
        strSql = "Select zl_GetPathCharge(0,0,[1],[2],0,[3],Sysdate,[4]) as 金额 From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mfrmVersion_CalcPathCost", mlng路径ID, vVersion.版本号, intDay, lng分支ID)
        
        If intDay <= intDayMin Then
            curCostMin = curCostMin + Nvl(rsTmp!金额, 0)
        End If
        curCostMax = curCostMax + Nvl(rsTmp!金额, 0)
    Next
    On Error GoTo 0
    Screen.MousePointer = 0
    
    If curCostMin = 0 And curCostMax = 0 Then
        MsgBox "计算无费用。", vbInformation, gstrSysName
    Else
        CostMin = curCostMin: CostMax = curCostMax
    End If
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mfrmVersion_CheckDataValid(Version As TYPE_PATH_VERSION, Branch As TYPE_PATH_BRANCH, Cancel As Boolean)
    Dim vStep As TYPE_PATH_STEP
    Dim i As Long
    Dim objComboBranch As CommandBarComboBox
    Dim lngBegin As Long, lngEnd As Long
    Dim strSql As String, rsTmp As Recordset
    Dim lngDays As Long
    
    With vsPath
        If Branch.版本号 = 0 Then
            '标准住院日不应小于已有阶段的天数范围
            For i = .Cols - 1 To .FixedCols + .FrozenCols Step -1
                If TypeName(.ColData(i)) <> "Empty" Then
                    vStep = .ColData(i)
                    '只需检查最后一个有具体天数范围的时间阶段
                    If vStep.结束天数 <> 0 Or vStep.开始天数 <> 0 Then
                        If InStr(Version.标准住院日, "-") > 0 Then
                            If vStep.结束天数 <> 0 Then
                                If Val(Split(Version.标准住院日, "-")(1)) < vStep.结束天数 Then
                                    MsgBox "标准住院日的最高天数 " & Val(Split(Version.标准住院日, "-")(1)) & " 天不应小于时间阶段已指定的天数 " & vStep.结束天数 & " 天。", vbInformation, gstrSysName
                                    Cancel = True: Exit Sub
                                End If
                            ElseIf vStep.开始天数 <> 0 Then
                                If Val(Split(Version.标准住院日, "-")(1)) < vStep.开始天数 Then
                                    MsgBox "标准住院日的最高天数 " & Val(Split(Version.标准住院日, "-")(1)) & " 天不应小于时间阶段已指定的天数 " & vStep.开始天数 & " 天。", vbInformation, gstrSysName
                                    Cancel = True: Exit Sub
                                End If
                            End If
                        Else
                            If vStep.结束天数 <> 0 Then
                                If Val(Version.标准住院日) < vStep.结束天数 Then
                                    MsgBox "标准住院日 " & Version.标准住院日 & " 天不应小于时间阶段已指定的天数 " & vStep.结束天数 & " 天。", vbInformation, gstrSysName
                                    Cancel = True: Exit Sub
                                End If
                            ElseIf vStep.开始天数 <> 0 Then
                                If Val(Version.标准住院日) < vStep.开始天数 Then
                                    MsgBox "标准住院日 " & Version.标准住院日 & " 天不应小于时间阶段已指定的天数 " & vStep.开始天数 & " 天。", vbInformation, gstrSysName
                                    Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                        Exit For
                    End If
                End If
            Next
        Else
            '分支路径检查
            '判断是否是新增分支
            Set objComboBranch = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Branch, True)
            If objComboBranch Is Nothing Then Cancel = True: Exit Sub
            If objComboBranch.ListIndex = 0 Then Cancel = True: Exit Sub
            
            If mblnAddNew Then
                '新增时检查标准住院日的开始时间必须是前一阶段的开始天数+1到结束天数+1之间
                If objComboBranch.ItemData(objComboBranch.ListIndex) = 0 Then
                    For i = .Cols - 1 To .FixedCols + .FrozenCols Step -1
                        If TypeName(.ColData(i)) <> "Empty" Then
                            vStep = .ColData(i)
                            '只需检查最后一个有具体天数范围的时间阶段
                            If (vStep.结束天数 <> 0 Or vStep.开始天数 <> 0) And vStep.ID = Branch.前一阶段ID Then
                                lngBegin = vStep.开始天数 + 1
                                lngEnd = IIf(vStep.结束天数 <> 0, vStep.结束天数, vStep.开始天数) + 1
                                Exit For
                            End If
                        End If
                    Next
                Else
                    strSql = "Select 开始天数,结束天数 From 临床路径阶段 Where ID = [1]"
                    On Error GoTo errH
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Branch.前一阶段ID)
                    If rsTmp.RecordCount > 0 Then
                        lngBegin = Val(rsTmp!开始天数 & "") + 1
                        lngEnd = IIf(Val(rsTmp!结束天数 & "") <> 0, Val(rsTmp!结束天数 & ""), Val(rsTmp!开始天数 & "")) + 1
                    End If
                    On Error GoTo 0
                End If
                
            Else
            '设置时检查
                '变化的开始天数
                lngDays = IIf(InStr(Branch.标准住院日, "-") > 0, Val(Split(Branch.标准住院日, "-")(0)), Val(Branch.标准住院日)) - mlngDays
                '标准住院日不应小于已有阶段的天数范围
                For i = .Cols - 1 To .FixedCols + .FrozenCols Step -1
                    If TypeName(.ColData(i)) <> "Empty" Then
                        vStep = .ColData(i)
                        '只需检查最后一个有具体天数范围的时间阶段
                        If vStep.结束天数 <> 0 Or vStep.开始天数 <> 0 Then
                            If InStr(Branch.标准住院日, "-") > 0 Then
                                If Val(Split(Branch.标准住院日, "-")(1)) < Val(Split(Branch.标准住院日, "-")(0)) Then
                                    MsgBox "分支的标准住院日的结束时间不能大于开始时间。", vbInformation, gstrSysName
                                    Cancel = True: Exit Sub
                                End If
                                If vStep.结束天数 <> 0 Then
                                    If Val(Split(Branch.标准住院日, "-")(1)) < vStep.结束天数 + lngDays Then
                                        MsgBox "分支的标准住院日的最高天数 " & Val(Split(Branch.标准住院日, "-")(1)) & " 天不应小于时间阶段已指定的天数 " & vStep.结束天数 & " 天。", vbInformation, gstrSysName
                                        Cancel = True: Exit Sub
                                    End If
                                End If
                                If vStep.开始天数 <> 0 Then
                                    If Val(Split(Branch.标准住院日, "-")(1)) < vStep.开始天数 + lngDays Then
                                        MsgBox "分支的标准住院日的最高天数 " & Val(Split(Branch.标准住院日, "-")(1)) & " 天不应小于时间阶段已指定的天数 " & vStep.开始天数 & " 天。", vbInformation, gstrSysName
                                        Cancel = True: Exit Sub
                                    End If
                                End If
                            Else
                                If vStep.结束天数 <> 0 Then
                                    If Val(Branch.标准住院日) < vStep.结束天数 + lngDays Then
                                        MsgBox "分支的标准住院日 " & Branch.标准住院日 & " 天不应小于时间阶段已指定的天数 " & vStep.结束天数 & " 天。", vbInformation, gstrSysName
                                        Cancel = True: Exit Sub
                                    End If
                                End If
                                If vStep.开始天数 <> 0 Then
                                    If Val(Branch.标准住院日) < vStep.开始天数 + lngDays Then
                                        MsgBox "分支的标准住院日 " & Branch.标准住院日 & " 天不应小于时间阶段已指定的天数 " & vStep.开始天数 & " 天。", vbInformation, gstrSysName
                                        Cancel = True: Exit Sub
                                    End If
                                End If
                            End If
                            Exit For
                        End If
                    End If
                Next
                mlngDays = lngDays
                '检查前一阶段开始时间范围
                strSql = "Select 开始天数,结束天数 From 临床路径阶段 Where ID = [1]"
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Branch.前一阶段ID)
                If rsTmp.RecordCount > 0 Then
                    lngBegin = Val(rsTmp!开始天数 & "") + 1
                    lngEnd = IIf(Val(rsTmp!结束天数 & "") <> 0, Val(rsTmp!结束天数 & ""), Val(rsTmp!开始天数 & "")) + 1
                End If
                On Error GoTo 0
            End If
            If lngEnd = 0 Then
                MsgBox "设置的分支阶段的前一阶段未找到。", vbInformation, gstrSysName
                Cancel = True: Exit Sub
            End If
            If InStr(Branch.标准住院日, "-") > 0 Then
                If Val(Split(Branch.标准住院日, "-")(0)) < lngBegin Then
                    MsgBox "分支的标准住院日的开始天数天必须大于前一阶段的开始天数 " & lngBegin - 1 & " 天。", vbInformation, gstrSysName
                    Cancel = True: Exit Sub
                End If
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optSelect_Click(Index As Integer)
    optSelect(IX_ALL).Tag = Index
    If Me.Visible Then
        Call FuncShowItemBySendor
    End If
End Sub

Private Sub txtFind_GotFocus()
    Call zlControl.TxtSelAll(txtFind)
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call FuncFindItem
    End If
End Sub

Private Sub txtFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTip As String
    
    strTip = "查找(Ctrl+F)" & vbCrLf & "查找下一个(F3)"
    
    zlCommFun.ShowTipInfo txtFind.Hwnd, strTip, True
End Sub

Private Sub vsPath_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim vArea As CONST_AREA
    
    If mbytMode = Mode_Design Then
        vArea = GetArea(NewRow, NewCol)
        If vArea = Area_Category Then
            vsPath.FocusRect = flexFocusSolid
        Else
            vsPath.FocusRect = flexFocusHeavy
        End If
        If picBottom.Visible Then
            If vArea = Area_Item Then
                Call FuncShowItemAdvice
            Else
                Call FuncShowAdvice(2)
            End If
            Call FuncSetAuditBtn
        End If
    End If
End Sub

Private Sub vsPath_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    With vsPath
        If Row = -1 And Col <> -1 Then
            .AutoSize .FixedCols, .Cols - 1, , 45
            Call SetTableCommonStyle(True)
        End If
    End With
End Sub

Private Sub vsPath_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    '左上角交叉区域不允许进入
    If GetArea(NewRow, NewCol) = Area_Cross Then
        If vsPath.Redraw <> flexRDNone Then Cancel = True
    Else
        mlngNewRow = NewRow: mlngNewCol = NewCol
    End If
End Sub

Private Sub vsPath_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
    With vsPath
        If GetArea(NewRowSel, NewColSel) = Area_Cross Then
            If .Redraw <> flexRDNone Then Cancel = True '左上角交叉区域不允许进入
        ElseIf GetArea(NewRowSel, NewColSel) <> GetArea(.Row, .Col) And Not (mlngNewRow = NewRowSel And mlngNewCol = NewColSel) Then
            If .Redraw <> flexRDNone Then Cancel = True '不允许不同区域交叉选择
        End If
    End With
End Sub

Private Sub vsPath_BeforeSort(ByVal Col As Long, Order As Integer)
    Dim objControl As CommandBarControl
    
    Order = 0
    If Col >= vsPath.FixedCols + vsPath.FrozenCols And Col <= vsPath.Cols - 1 Then
        vsPath.Col = Col
        
        Set objControl = cbsMain.FindControl(, cmd_Edit_EvalStep, True, True)
        If Not objControl Is Nothing Then
            If objControl.Enabled Then Call FuncEvaluateStep
        End If
    End If
End Sub

Private Sub FuncShowItemAdvice()
    
    With vsPath
        If .Row < 0 Or .Col < 0 Then Exit Sub
        If mbytFunc = 0 Then Exit Sub
        
        If TypeName(.Cell(flexcpData, .Row, .Col)) <> "Empty" Then
            If (.Cell(flexcpBackColor, .Row, .Col) = Color_DiffBack Or .Cell(flexcpForeColor, .Row, .Col) = Color_NeedAuditFore) Then
                Call FuncLoadChangeTimes
            Else
                Call FuncShowAdvice(2)
            End If
            Call FuncShowAdvice(0)
        Else
            Call FuncShowAdvice(2)
        End If
    End With
End Sub

Private Sub vsPath_DblClick()
    Dim vArea As CONST_AREA
    Dim lngRow As Long, lngCol As Long

    With vsPath
        lngRow = .MouseRow
        lngCol = .MouseCol

        vArea = GetArea(lngRow, lngCol)
        If vArea <> Area_Cross And vArea <> -1 Then
            If mbytMode = Mode_Design And mblnEditable And vArea = Area_Category Then
                '可编辑（未审核）时，双击分类列，自动进入编辑状态
                .EditCell: Exit Sub
            End If
            Call vsPath_KeyPress(13)
        End If

    End With
End Sub

Private Sub vsPath_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim objControl As CommandBarControl
    
    If KeyCode = vbKeyDelete And Shift = 0 Then
        Set objControl = cbsMain.FindControl(, cmd_Edit_Delete, True, True)
        If Not objControl Is Nothing Then
            If objControl.Enabled Then objControl.Execute
        End If
    End If
End Sub

Private Sub vsPath_KeyPress(KeyAscii As Integer)
    Dim vArea As CONST_AREA
    Dim objControl As CommandBarControl
    
    vArea = GetArea(vsPath.Row, vsPath.Col)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If vArea = Area_Category Then
            Call CategoryEnterNextCell(vsPath.Row, vsPath.Col)
        Else
            Set objControl = cbsMain.FindControl(, cmd_Edit_Edit, True, True)
            If Not objControl Is Nothing Then
                If objControl.Enabled Then objControl.Execute
            End If
        End If
    End If
End Sub

Private Sub vsPath_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
End Sub

Private Sub vsPath_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long, lngCol As Long
    Dim vArea As CONST_AREA

    If Button = vbLeftButton Then
        With vsPath
            lngRow = .MouseRow
            lngCol = .MouseCol
            vArea = GetArea(lngRow, lngCol)
            If vArea = -1 Then
                Exit Sub
            ElseIf vArea = Area_Category And .TextMatrix(lngRow, lngCol) = "" Then
                .EditCell   '类别为空，强制编辑
            End If
        End With
    End If
End Sub

Private Sub vsPath_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'功能：显示鼠标提示
    Dim lngRow As Long, lngCol As Long
    Dim vArea As CONST_AREA
    Dim vStep As TYPE_PATH_STEP
    Dim vItem As TYPE_PATH_ITEM
    Dim vEvalMark As TYPE_PATH_EvalMark
    Dim vEvalCond As TYPE_PATH_EvalCond
    Dim strTip As String, strTmp As String, i As Long
    Dim rsTmp As ADODB.Recordset
    
    With vsPath
        If .Rows = 0 Or .Cols = 0 Then
            zlCommFun.ShowTipInfo 0, ""
            Exit Sub
        End If
        lngRow = .MouseRow
        lngCol = .MouseCol
        
        vArea = GetArea(lngRow, lngCol)
        If vArea = Area_Step Then
            If TypeName(.ColData(lngCol)) <> "Empty" Then
                vStep = .ColData(lngCol)
                If vStep.结束天数 <> 0 Then
                    strTip = strTip & "时间阶段：住院第" & vStep.开始天数 & "-" & vStep.结束天数 & "天"
                Else
                    strTip = strTip & "时间阶段：住院第" & vStep.开始天数 & "天"
                End If
                If lngRow = .FixedRows + .FrozenRows - 1 Then
                    If .TextMatrix(lngRow, lngCol) <> .TextMatrix(lngRow - 1, lngCol) Then
                        strTip = strTip & "，" & .TextMatrix(lngRow, lngCol)
                    End If
                End If
                If Replace(vStep.标志, "0", "") <> "" Then
                    strTip = strTip & vbCrLf & "●标志："
                    For i = 1 To Len(vStep.标志)
                        If Mid(vStep.标志, i, 1) = "1" Then
                            strTip = strTip & Decode(i, 1, "住院日", 2, "手术日", 3, "分娩日", 4, "出院日") & "、"
                        End If
                    Next
                    strTip = Left(strTip, Len(strTip) - 1)
                End If
                If vStep.分类 <> "" Then
                    strTip = strTip & vbCrLf & "●分类：" & vStep.分类
                End If
                strTip = strTip & vbCrLf & "●说明：" & vStep.说明
            End If
        ElseIf vArea = Area_Item Then
            If TypeName(.Cell(flexcpData, lngRow, lngCol)) <> "Empty" Then
                vItem = .Cell(flexcpData, lngRow, lngCol)
                strTip = "路径项目：" & vItem.项目内容
                If vItem.导入结果 = 1 Then
                    If vItem.医嘱IDs <> "" Then
                        If Not vItem.Tip Like vItem.医嘱IDs & ":*" Then
                            vItem.Tip = vItem.医嘱IDs & ":" & GetAdviceDefineText(vItem.医嘱IDs, mrsAdvice)
                            .Cell(flexcpData, lngRow, lngCol) = vItem
                        End If
                        strTip = strTip & vbCrLf & "●医嘱摘要：" & vbCrLf & Mid(vItem.Tip, InStr(vItem.Tip, ":") + 1)
                    End If
                    If vItem.病历IDs <> "" Or vItem.新版病历IDs <> "" Then
                        If Not vItem.Tip Like vItem.病历IDs & "|" & vItem.新版病历IDs & ":*" Then
                            If vItem.Edit = 0 Then
                                If vItem.病历IDs <> "" And vItem.新版病历IDs <> "" Then
                                    strTmp = GetEPRDefineText(, vItem.ID)
                                ElseIf vItem.病历IDs <> "" Then
                                    strTmp = GetEPRDefineText(vItem.病历IDs)
                                Else
                                    strTmp = GetEPRDefineText(vItem.新版病历IDs, vItem.ID)
                                End If
                            Else
                                Set rsTmp = FuncGetEMRInfo(vItem.病历详情)
                                strTmp = ""
                                Do While Not rsTmp.EOF
                                    strTmp = strTmp & "、" & rsTmp!名称
                                    rsTmp.MoveNext
                                Loop
                                strTmp = Mid(strTmp, 2)
                            End If
                            vItem.Tip = vItem.病历IDs & "|" & vItem.新版病历IDs & ":" & strTmp
                            .Cell(flexcpData, lngRow, lngCol) = vItem
                        End If
                        strTip = strTip & vbCrLf & "●对应病历：" & Mid(vItem.Tip, InStr(vItem.Tip, ":") + 1)
                    End If
                    strTip = strTip & vbCrLf & "●生 成 者：" & Decode(vItem.生成者, 1, "医生", 2, "护士")
                    strTip = strTip & vbCrLf & "●执行方式：" & Decode(vItem.执行方式, 0, "无须执行", 1, "每天执行", 2, "至少执行一次", 3, "必要时执行", 4, "必须执行一次")
                    If vItem.执行方式 <> 0 Then
                        strTip = strTip & vbCrLf & "●执 行 者：" & Decode(vItem.执行者, 1, "医生", 2, "护士")
                        If vItem.项目结果 <> "" Then
                            strTmp = ""
                            For i = 0 To UBound(Split(Split(vItem.项目结果, vbTab)(0), ","))
                                strTmp = strTmp & "、" & Split(Split(Split(vItem.项目结果, vbTab)(0), ",")(i), "|")(0)
                            Next
                            strTip = strTip & vbCrLf & "●执行结果：" & Mid(strTmp, 2)
                            strTip = strTip & vbCrLf & "●缺省结果：" & Split(vItem.项目结果, vbTab)(1)
                        End If
                    End If
                Else
                    strTip = vItem.导入参考
                End If
            End If
        ElseIf vArea = Area_Cross Or lngRow = .FixedRows - 1 And lngCol <= .FixedCols + .FrozenCols - 1 And lngCol >= 0 Then
            If Not mvEvalImport.指标集 Is Nothing Then
                If mvEvalImport.指标集.count > 0 Then
                    strTip = strTip & vbCrLf & "●评估指标："
                    For i = 1 To mvEvalImport.指标集.count
                        vEvalMark = mvEvalImport.指标集(i)
                        strTip = strTip & vbCrLf & "　○" & vEvalMark.评估指标 & "，结果：" & Split(vEvalMark.指标结果, vbTab)(0)
                    Next
                End If
            End If
            If Not mvEvalImport.条件集 Is Nothing Then
                If mvEvalImport.条件集.count > 0 Then
                    strTip = strTip & vbCrLf & "●计算条件："
                    For i = 1 To mvEvalImport.条件集.count
                        vEvalCond = mvEvalImport.条件集(i)
                        strTip = strTip & vbCrLf & "　○[" & GetMarkName(vEvalCond.指标ID, lngCol) & "] " & vEvalCond.关系式 & " [" & vEvalCond.条件值 & "]"
                    Next
                End If
            End If
            If strTip <> "" Then
                strTip = "导入评估信息：" & strTip
            Else
                strTip = "没有设置导入评估信息。"
            End If
        ElseIf lngRow = .FixedRows - 1 And lngCol >= .FixedCols + .FrozenCols Then
            If TypeName(.ColData(lngCol)) <> "Empty" Then
                vStep = .ColData(lngCol)
                If Not vStep.评估.指标集 Is Nothing Then
                    If vStep.评估.指标集.count > 0 Then
                        strTip = strTip & vbCrLf & "●评估指标："
                        For i = 1 To vStep.评估.指标集.count
                            vEvalMark = vStep.评估.指标集(i)
                            strTip = strTip & vbCrLf & "　○" & vEvalMark.评估指标 & "，结果：" & Split(vEvalMark.指标结果, vbTab)(0)
                        Next
                    End If
                End If
                If Not vStep.评估.条件集 Is Nothing Then
                    If vStep.评估.条件集.count > 0 Then
                        strTip = strTip & vbCrLf & "●计算条件："
                        For i = 1 To vStep.评估.条件集.count
                            vEvalCond = vStep.评估.条件集(i)
                            If vEvalCond.指标ID <> 0 Then
                                strTip = strTip & vbCrLf & "　○[" & GetMarkName(vEvalCond.指标ID, lngCol) & "] " & vEvalCond.关系式 & " [" & vEvalCond.条件值 & "]"
                            ElseIf vEvalCond.项目ID <> 0 Then
                                strTip = strTip & vbCrLf & "　○[" & GetItemName(vEvalCond.项目ID, lngCol) & "] " & vEvalCond.关系式 & " [" & vEvalCond.条件值 & "]"
                            End If
                        Next
                    End If
                End If
                If strTip <> "" Then
                    strTip = "阶段评估信息：" & strTip
                Else
                    If mbytMode = Mode_Design And mblnEditable Then
                        strTip = "尚未设置该时间阶段的评估信息，请点击设置。"
                    Else
                        strTip = "没有设置该时间阶段的评估信息。"
                    End If
                End If
            Else
                If mbytMode = Mode_Design And mblnEditable Then
                    strTip = "尚未设置该时间阶段的评估信息，请点击设置。"
                Else
                    strTip = "没有设置该时间阶段的评估信息。"
                End If
            End If
        End If
        
        If strTip <> "" Then
            zlCommFun.ShowTipInfo .Hwnd, strTip, True
        Else
            zlCommFun.ShowTipInfo 0, ""
        End If
    End With
End Sub

Private Function GetItemName(ByVal lngItemID As Long, ByVal lngCol As Long) As String
'功能：获取指定阶段中指定项目ID的项目名称
    Dim vItem As TYPE_PATH_ITEM
    Dim i As Long
    
    With vsPath
        For i = .FixedRows + .FrozenRows To .Rows - 1
            If TypeName(.Cell(flexcpData, i, lngCol)) <> "Empty" Then
                vItem = .Cell(flexcpData, i, lngCol)
                If vItem.ID = lngItemID Then
                    GetItemName = vItem.项目内容
                    Exit Function
                End If
            End If
        Next
    End With
End Function

Private Function GetMarkName(ByVal lngMarkID As Long, Optional ByVal lngCol As Long)
'功能：获取指定指标ID的指标名称
'参数：lngCol=指定时为具体的阶段列，否则表示导入评估指标
    Dim vStep As TYPE_PATH_STEP
    Dim vEvalMark As TYPE_PATH_EvalMark
    Dim i As Long
    
    If lngCol = 0 Then
        If Not mvEvalImport.指标集 Is Nothing Then
            For i = 1 To mvEvalImport.指标集.count
                vEvalMark = mvEvalImport.指标集(i)
                If vEvalMark.ID = lngMarkID Then
                    GetMarkName = vEvalMark.评估指标
                    Exit Function
                End If
            Next
        End If
    Else
        If TypeName(vsPath.ColData(lngCol)) <> "Empty" Then
            vStep = vsPath.ColData(lngCol)
            If Not vStep.评估.指标集 Is Nothing Then
                For i = 1 To vStep.评估.指标集.count
                    vEvalMark = vStep.评估.指标集(i)
                    If vEvalMark.ID = lngMarkID Then
                        GetMarkName = vEvalMark.评估指标
                        Exit Function
                    End If
                Next
            End If
        End If
    End If
End Function

Private Sub vsPath_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long, lngCol As Long
    Dim vArea As CONST_AREA
    Dim objPopup As CommandBarPopup
    
    If Button = 2 Then
        lngRow = vsPath.MouseRow
        lngCol = vsPath.MouseCol
        vArea = GetArea(lngRow, lngCol)
        If vArea <> Area_Cross And vArea <> -1 Then
            '先后顺序在BeforeRowColChange事件中有限制
            If vsPath.Col = vsPath.FixedCols Then
                vsPath.Col = lngCol: vsPath.Row = lngRow
            Else
                vsPath.Row = lngRow: vsPath.Col = lngCol
            End If
            Set objPopup = cbsMain.FindControl(, conMenu_EditPopup, True)
            If Not objPopup Is Nothing Then
                objPopup.CommandBar.ShowPopup
            End If
        End If
    End If
End Sub

Private Sub vsPath_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsPath.EditSelStart = 0
    vsPath.EditSelLength = zlCommFun.ActualLen(vsPath.EditText)
End Sub

Private Sub vsPath_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsPath
        If GetArea(Row, Col) <> Area_Category Then
            Cancel = True    '仅分类内容可以直接输入
        ElseIf .RowSel <> Row Or .ColSel <> Col Then
            Cancel = True    '选择范围时不允许输入
        End If
    End With
End Sub

Private Sub vsPath_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngR1 As Long, lngR2 As Long
    Dim vArea As CONST_AREA, i As Long, j As Long
    Dim vItem As TYPE_PATH_ITEM

    vArea = GetArea(Row, Col)

    With vsPath
        If vArea = Area_Category Then
            .EditText = Trim(.EditText)
            If LenB(StrConv(.EditText, vbFromUnicode)) > 50 Then
                MsgBox "您输入的分类名称的字数超过25个，请重新输入。", vbInformation + vbOKOnly, gstrSysName
                Cancel = True
                Exit Sub
            End If
            '没有改动时，跳出过程
            If .TextMatrix(Row, Col) = .EditText Then
                Exit Sub
            End If
            If Trim(.EditText) = "" Then
                '相当于删除,有对应项目时不允许清除
                .GetMergedRange Row, Col, lngR1, 0, lngR2, 0
                For i = lngR1 To lngR2
                    If Replace(.Cell(flexcpText, i, .FixedCols + .FrozenCols, i, .Cols - 1), vbTab, "") <> "" Then
                        MsgBox "该分类中已经存在对应的项目，请输入分类名称。", vbInformation, gstrSysName
                        Cancel = True: Exit Sub
                    End If
                Next
            Else
                '分类不能重复:合并单元范围控件自动输入
                i = .FixedRows + .FrozenCols
                Do While i <= .Rows - 1
                    If i <> Row And .TextMatrix(i, Col) = .EditText Then
                        MsgBox "输入的分类名称已经存在，请重新输入。", vbInformation, gstrSysName
                        Cancel = True: Exit Sub
                    End If

                    .GetMergedRange i, Col, 0, 0, lngR2, 0
                    i = lngR2 + 1    '跳过合并分类
                Loop
                .GetMergedRange Row, Col, lngR1, 0, lngR2, 0

                '该分类的所有项目都标记为修改状态
                For i = .FixedCols + .FrozenCols To .Cols - 1
                    For j = lngR1 To lngR2
                        If TypeName(.Cell(flexcpData, j, i)) <> "Empty" Then
                            vItem = .Cell(flexcpData, j, i)
                            vItem.Edit = 2
                            .Cell(flexcpData, j, i) = vItem
                        End If
                    Next
                Next
            End If

            '内容变化后，根据内容自动调行高
            For i = lngR1 To lngR2
                .TextMatrix(i, Col) = .EditText    '不然调整无效
            Next i
            .AutoSize .FixedCols, .Cols - 1, , 45
            Call SetTableCommonStyle(True)

            mblnChange = True

            '光标跳到下一分类
            If mblnReturn And Trim(.EditText) <> "" Then
                Call CategoryEnterNextCell(Row, Col)
            End If
        End If
    End With
End Sub

Private Sub CategoryEnterNextCell(ByVal lngRow As Long, ByVal lngCol As Long)
    Dim lngR2 As Long
    
    With vsPath
        .GetMergedRange lngRow, lngCol, 0, 0, lngR2, 0
        If lngR2 + 1 <= .Rows - 1 Then
            .Row = lngR2 + 1
            .ShowCell .Row, .Col
        End If
    End With
End Sub

Private Sub FuncCategoryDelete()
'功能：删除当前选择的分类行
    Dim lngR1 As Long, lngR2 As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim lngRow As Long, i As Long
    
    With vsPath
        lngRow = .Row
        
        '当前选择范围
        .GetSelection lngR1, 0, lngR2, 0
        lngBegin = lngR1
        
        '考虑合并单元：合并单元选择时RowSel,ColSel不变
        .GetMergedRange lngR2, .Col, lngR1, 0, lngR2, 0
        lngEnd = lngR2
        
        For i = lngBegin To lngEnd
            If Replace(.Cell(flexcpText, i, .FixedCols + .FrozenCols, i, .Cols - 1), vbTab, "") <> "" Then
                MsgBox "所选择分类中已经存在对应的路径项目，不能删除。", vbInformation, gstrSysName
                Exit Sub
            End If
        Next
        
        If MsgBox("确实要删除所选择的分类行吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        '删除处理
        .Redraw = flexRDNone
        For i = lngEnd To lngBegin Step -1
            .RemoveItem i
        Next
        If .Rows = .FixedRows + .FrozenRows Then
            '删除后至少保留一行分类
            .AddItem "": .RowHeight(.Rows - 1) = ROW_HEIGHT_MIN
            .Row = .FixedRows + .FrozenRows
        ElseIf lngRow <= .Rows - 1 Then
            .Row = lngRow
        Else
            .Row = .Rows - 1
        End If
        .ShowCell .Row, .Col
        .Redraw = flexRDDirect
        .SetFocus
    End With
    
    mblnChange = True
End Sub

Private Sub FuncCategoryInsert(ByVal intPos As Integer)
'功能：插入新的分类行
'参数：inPos=1：在当前行后面，-1：在当前行前面
    Dim lngRow As Long
    Dim lngR1 As Long, lngR2 As Long

    With vsPath
        If .TextMatrix(.Row, .Col) = "" Then
            MsgBox "当前行分类尚未输入，请先输入当前行分类。", vbInformation, gstrSysName
            Exit Sub
        End If

        .GetMergedRange .Row, .Col, lngR1, 0, lngR2, 0    '需要考虑合并项范围
        lngRow = IIf(intPos = -1, lngR1, lngR2 + 1)
        .AddItem "", lngRow
        .RowHeight(lngRow) = ROW_HEIGHT_MIN
        .Row = lngRow
        .EditCell
        .ShowCell .Row, .Col
    End With

    mblnChange = True
End Sub

Private Sub mfrmPathStep_CheckDataValid(TimeStep As TYPE_PATH_STEP, Cancel As Boolean)
'功能：检查所输入时间阶段数据的正确性
    Dim objCombo As CommandBarComboBox
    Dim vVersion As TYPE_PATH_VERSION
    Dim objComboBranch As CommandBarComboBox
    Dim vBranch As TYPE_PATH_BRANCH
    Dim vStep As TYPE_PATH_STEP
    Dim strMsg As String, i As Long
    Dim strStep As String, j As Long
    Dim strSql As String, rsTmp As Recordset
    
    With vsPath
        '与标准住院日之间的关系检查
        Set objCombo = cbsMain(cbsMain.count - 1).FindControl(, cmd_Edit_Version, True)
        Set objComboBranch = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Branch, True)
        If Not objCombo Is Nothing And Not objComboBranch Is Nothing Then
            If Not objCombo.ListIndex = 0 And Not objComboBranch.ListIndex = 0 Then
                vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))
                vBranch = mcolBranch("_" & objComboBranch.ItemData(objComboBranch.ListIndex))
                If vBranch.分支名称 = "主路径" Then
                    If vVersion.标准住院日 <> "" Then
                        If InStr(vVersion.标准住院日, "-") > 0 Then
                            If TimeStep.结束天数 <> 0 And TimeStep.结束天数 > Val(Split(vVersion.标准住院日, "-")(1)) Then
                                MsgBox "当前时间阶段的结束天数 " & TimeStep.结束天数 & " 天高于了标准住院日指定的最高天数 " & Val(Split(vVersion.标准住院日, "-")(1)) & " 天。", vbInformation, gstrSysName
                                Cancel = True: Exit Sub
                            ElseIf TimeStep.开始天数 <> 0 And TimeStep.开始天数 > Val(Split(vVersion.标准住院日, "-")(1)) Then
                                MsgBox "当前时间阶段的天数 " & TimeStep.开始天数 & " 天高于了标准住院日指定的最高天数 " & Val(Split(vVersion.标准住院日, "-")(1)) & " 天。", vbInformation, gstrSysName
                                Cancel = True: Exit Sub
                            End If
                        Else
                            If TimeStep.结束天数 <> 0 And TimeStep.结束天数 > Val(vVersion.标准住院日) Then
                                MsgBox "当前时间阶段的结束天数 " & TimeStep.结束天数 & " 天高于了标准住院日指定的最高天数 " & Val(vVersion.标准住院日) & " 天。", vbInformation, gstrSysName
                                Cancel = True: Exit Sub
                            ElseIf TimeStep.开始天数 <> 0 And TimeStep.开始天数 > Val(vVersion.标准住院日) Then
                                MsgBox "当前时间阶段的天数 " & TimeStep.开始天数 & " 天高于了标准住院日指定的最高天数 " & Val(vVersion.标准住院日) & " 天。", vbInformation, gstrSysName
                                Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                Else
                    If vBranch.标准住院日 <> "" Then
                        If InStr(vBranch.标准住院日, "-") > 0 Then
                            If TimeStep.结束天数 <> 0 And TimeStep.结束天数 > Val(Split(vBranch.标准住院日, "-")(1)) Then
                                MsgBox "当前时间阶段的结束天数 " & TimeStep.结束天数 & " 天高于了分支标准住院日指定的最高天数 " & Val(Split(vBranch.标准住院日, "-")(1)) & " 天。", vbInformation, gstrSysName
                                Cancel = True: Exit Sub
                            ElseIf TimeStep.开始天数 <> 0 And TimeStep.开始天数 > Val(Split(vBranch.标准住院日, "-")(1)) Then
                                MsgBox "当前时间阶段的天数 " & TimeStep.开始天数 & " 天高于了分支标准住院日指定的最高天数 " & Val(Split(vBranch.标准住院日, "-")(1)) & " 天。", vbInformation, gstrSysName
                                Cancel = True: Exit Sub
                            End If
                        Else
                            If TimeStep.结束天数 <> 0 And TimeStep.结束天数 > Val(vBranch.标准住院日) Then
                                MsgBox "当前时间阶段的结束天数 " & TimeStep.结束天数 & " 天高于了分支标准住院日指定的最高天数 " & Val(vBranch.标准住院日) & " 天。", vbInformation, gstrSysName
                                Cancel = True: Exit Sub
                            ElseIf TimeStep.开始天数 <> 0 And TimeStep.开始天数 > Val(vBranch.标准住院日) Then
                                MsgBox "当前时间阶段的天数 " & TimeStep.开始天数 & " 天高于了分支标准住院日指定的最高天数 " & Val(vBranch.标准住院日) & " 天。", vbInformation, gstrSysName
                                Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    '分支路径的第一个阶段开始天数必须在前一阶段的开始天数和结束天数+1之间
                    If .Col = .FixedCols + .FrozenCols And vBranch.前一阶段ID <> 0 And vBranch.分支ID <> 0 Then
                        On Error GoTo errH
                        strSql = "Select 开始天数,NVL(结束天数,开始天数) AS 结束天数 From 临床路径阶段 Where ID=[1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, vBranch.前一阶段ID)
                        If rsTmp.RecordCount > 0 Then
                            If Not Between(TimeStep.开始天数, Val(rsTmp!开始天数 & "") + 1, Val(rsTmp!结束天数 & "") + 1) Then
                                MsgBox "分支路径的第一个阶段的开始天数必须在分支的前一阶段的开始天数后一天和结束天数后一天：" & Val(rsTmp!开始天数 & "") + 1 & "-" & Val(rsTmp!结束天数 & "") + 1 & "之间。", vbInformation, gstrSysName
                                Cancel = True: Exit Sub
                            End If
                        End If
                        On Error GoTo 0
                    End If
                End If
            End If
        End If
        
        '与其他各个时间阶段之间的关系检查
        For i = .FixedCols + .FrozenCols To .Cols - 1
            If TypeName(.ColData(i)) <> "Empty" Then
                vStep = .ColData(i)
                If vStep.父ID = 0 And vStep.ID <> TimeStep.ID Then
                    If objComboBranch.ListIndex > 1 And Mid(TimeStep.标志, 1, 1) = "1" And (vStep.开始天数 <> TimeStep.开始天数 Or vStep.结束天数 <> TimeStep.结束天数) Then
                        strMsg = "分支路径不能设置住院日。": Exit For
                    End If
                    '只有第一个阶段允许设置住院日
                    If i < .Col And Mid(TimeStep.标志, 1, 1) = "1" And (vStep.开始天数 <> TimeStep.开始天数 Or vStep.结束天数 <> TimeStep.结束天数) Then
                        strMsg = "只有第一个时间阶段才可能设置为住院日，除非与第一个阶段的时间相同。": Exit For
                    End If
                    '只有最后一个阶段允许设置出院日
                    If i > .Col And Mid(TimeStep.标志, 4, 1) = "1" And (vStep.开始天数 <> TimeStep.开始天数 Or vStep.结束天数 <> TimeStep.结束天数) Then
                        strMsg = "只有最后一个时间阶段才可能设置为出院日，除非与最后一个阶段的时间相同。": Exit For
                    End If
                    
                    '手术日/分娩日不能同时出现不同阶段
                    If Mid(TimeStep.标志, 2, 1) = "1" And Mid(vStep.标志, 3, 1) = "1" And (vStep.开始天数 <> TimeStep.开始天数 Or vStep.结束天数 <> TimeStep.结束天数) Then
                        strMsg = "已经有其他时间阶段设置为分娩日，不能再设置当前阶段为手术日，除非这些阶段的时间相同。": Exit For
                    End If
                    If Mid(TimeStep.标志, 3, 1) = "1" And Mid(vStep.标志, 2, 1) = "1" And (vStep.开始天数 <> TimeStep.开始天数 Or vStep.结束天数 <> TimeStep.结束天数) Then
                        strMsg = "已经有其他时间阶段设置为手术日，不能再设置当前阶段为分娩日，除非这些阶段的时间相同。": Exit For
                    End If
                    
                    '几个标志只能有一个阶段有
                    For j = 1 To Len(TimeStep.标志)
                        If Mid(TimeStep.标志, j, 1) = "1" And Mid(vStep.标志, j, 1) = "1" And (vStep.开始天数 <> TimeStep.开始天数 Or vStep.结束天数 <> TimeStep.结束天数) Then
                            strMsg = "已经有其他时间阶段设置为" & Decode(j, 1, "住院日", 2, "手术日", 3, "分娩日", 4, "出院日") & "，除非这些阶段的时间相同。": Exit For
                        End If
                    Next
                    If j <= Len(TimeStep.标志) Then Exit For
                    
                    '标志在阶段中的顺序为：住院日-手术日/分娩日-出院日
                    If i < .Col Then
                        For j = 2 To Len(vStep.标志)
                            If Mid(vStep.标志, j, 1) = "1" And Mid(TimeStep.标志, j - 1, 1) = "1" And j <> 3 And (vStep.开始天数 <> TimeStep.开始天数 Or vStep.结束天数 <> TimeStep.结束天数 Or TimeStep.标志 <> vStep.标志) Then '3和2(3-1)等价
                                strMsg = "当前阶段不能设置为" & Decode(j - 1, 1, "住院日", 2, "手术日", 3, "分娩日", 4, "出院日") & "，请检查各个时间阶段标志设置的先后顺序。"
                                Exit For
                            End If
                        Next
                    ElseIf i > .Col Then
                        For j = 1 To Len(vStep.标志) - 1
                            If Mid(vStep.标志, j, 1) = "1" And Mid(TimeStep.标志, j + 1, 1) = "1" And j <> 2 And (vStep.开始天数 <> TimeStep.开始天数 Or vStep.结束天数 <> TimeStep.结束天数 Or TimeStep.标志 <> vStep.标志) Then '2和3(2+1)等价
                                strMsg = "当前阶段不能设置为" & Decode(j + 1, 1, "住院日", 2, "手术日", 3, "分娩日", 4, "出院日") & "，请检查各个时间阶段标志设置的先后顺序。"
                                Exit For
                            End If
                        Next
                    End If
                    
                    '天数范围应该在前面之后,后面之前
                    '天数范围可以部分交叉,也可能包含
                    If i < .Col Then
                        If TimeStep.开始天数 <= vStep.开始天数 And TimeStep.开始天数 <> 0 And vStep.开始天数 <> 0 Then
                            If TimeStep.开始天数 = vStep.开始天数 Then
'                                If TimeStep.结束天数 <> 0 And vStep.结束天数 <> 0 Then
'                                    strMsg = "当前阶段的天数如果前面阶段的天数相同，则不能两个阶段都设置为一个天数范围。": Exit For
'                                End If
                            Else
                                strMsg = "当前阶段的开始天数应该大于前面阶段的开始天数。": Exit For
                            End If
                        End If
                        If IIf(TimeStep.结束天数 = 0, TimeStep.开始天数, TimeStep.结束天数) < IIf(vStep.结束天数 = 0, vStep.开始天数, vStep.结束天数) Then
                            strMsg = "当前阶段的结束天数应该大于前面阶段的结束天数。": Exit For
                        End If
                        If IIf(vStep.结束天数 = 0, vStep.开始天数, vStep.结束天数) < TimeStep.开始天数 - 1 And i = .Col - 1 Then
                            strMsg = "当前阶段的开始天数必须跟前一个阶段为连续的,开始天数必须小于或等于" & TimeStep.开始天数 - 1 & "。": Exit For
                        End If
                    ElseIf i > .Col Then
                        If TimeStep.开始天数 >= vStep.开始天数 And TimeStep.开始天数 <> 0 And vStep.开始天数 <> 0 Then
                            If TimeStep.开始天数 = vStep.开始天数 Then
'                                If TimeStep.结束天数 <> 0 And vStep.结束天数 <> 0 Then
'                                    strMsg = "当前阶段的天数如果后面阶段的天数相同，则不能两个阶段都设置为一个天数范围。": Exit For
'                                End If
                            Else
                                strMsg = "当前阶段的开始天数应该小于后面阶段的开始天数。": Exit For
                            End If
                        End If
                        If IIf(TimeStep.结束天数 = 0, TimeStep.开始天数, TimeStep.结束天数) > IIf(vStep.结束天数 = 0, vStep.开始天数, vStep.结束天数) Then
                            strMsg = "当前阶段的结束天数应该小于后面阶段的结束天数。": Exit For
                        End If
                        If IIf(TimeStep.结束天数 = 0, TimeStep.开始天数, TimeStep.结束天数) < vStep.开始天数 - 1 And i = .Col + 1 Then
                            strMsg = "当前阶段的开始天数必须跟前一个阶段为连续的,结束天数必须大于或等于：" & vStep.开始天数 - 1 & "天。": Exit For
                        End If
                    End If
                End If
            End If
        Next
        
        If strMsg <> "" Then
            MsgBox strMsg, vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetNearStep(ByVal lngCol As Long, ByVal intPos As Integer, _
    Optional ByVal blnSub As Boolean, Optional ByVal blnSkip As Boolean = True) As TYPE_PATH_STEP
'功能：获取当前时间阶段相邻的时间阶段信息
'参数：lngCol=当前列
'      intPos=-1:前面，1:后面
'      blnSub=是否允许返回相邻的分支
'      blnSkip=是否允许跳过空的时间阶段
    Dim vStep As TYPE_PATH_STEP
    Dim i As Long

    With vsPath
        If intPos = -1 Then
            For i = lngCol - 1 To .FixedCols + .FrozenCols Step -1
                If TypeName(.ColData(i)) <> "Empty" Then
                    If blnSub Or .ColData(i).名称 <> vStep.名称 Then
                        vStep = .ColData(i): Exit For
                    End If
                Else
                    If Not blnSkip Then Exit For
                End If
            Next
        Else
            For i = lngCol + 1 To .Cols - 1
                If TypeName(.ColData(i)) <> "Empty" Then
                    If blnSub Or .ColData(i).名称 <> vStep.名称 Then
                        vStep = .ColData(i): Exit For
                    End If
                Else
                    If Not blnSkip Then Exit For
                End If
            Next
        End If
    End With
    
    GetNearStep = vStep
End Function

Private Sub FuncStepEdit()
'功能：设置当前时间阶段内容
    Dim vStep As TYPE_PATH_STEP
    Dim vPreStep As TYPE_PATH_STEP
    Dim vNextStep As TYPE_PATH_STEP
    Dim lngR1 As Long, lngC1 As Long
    Dim lngR2 As Long, lngC2 As Long
    Dim str分类s As String, i As Long, j As Long
    
    With vsPath
        If TypeName(.ColData(.Col)) <> "Empty" Then
            vStep = .ColData(.Col)
        End If
        
        '获取前一个时间阶段的内容
        vPreStep = GetNearStep(.Col, -1)
        vNextStep = GetNearStep(.Col, 1)
        
        '获取前后备用分支的分类名
        For i = .FixedCols + .FrozenCols To .Cols - 1
            If TypeName(.ColData(i)) <> "Empty" Then
                If .ColData(i).分类 <> "" Then
                    If InStr(str分类s & "|", "|" & .ColData(i).分类 & "|") = 0 Then
                        str分类s = str分类s & "|" & .ColData(i).分类
                    End If
                End If
            End If
        Next
        
        If mfrmPathStep.ShowEdit(Me, vStep, vPreStep, vNextStep, Mid(str分类s, 2)) Then
            If vStep.ID = 0 Then
                '保证有内容的阶段ID不为空，先预取一个ID
                vStep.ID = zlDatabase.GetNextId("临床路径阶段")
                vStep.Edit = 1 '0-原始,1-新增,2-修改
            Else
                If vStep.Edit = 0 Then vStep.Edit = 2
            End If
            
            If vStep.父ID <> 0 Then
                '备选分支可能设置了说明和分类
                .ColData(.Col) = vStep
                .TextMatrix(.Row, .Col) = IIf(vStep.说明 = "", "备用分支" & vStep.序号, vStep.说明) & IIf(vStep.分类 = "", "", ",") & vStep.分类
            ElseIf vStep.父ID = 0 Then
                .ColData(.Col) = vStep
                
                '备选分支相关信息同步变化
                For i = .Col + 1 To .Cols - 1
                    If TypeName(.ColData(i)) <> "Empty" Then
                        vNextStep = .ColData(i)
                        If vNextStep.父ID = vStep.ID Then
                            vNextStep.名称 = vStep.名称
                            vNextStep.开始天数 = vStep.开始天数
                            vNextStep.结束天数 = vStep.结束天数
                            vNextStep.标志 = vStep.标志
                            If vNextStep.Edit = 0 Then vNextStep.Edit = 2
                            
                            .ColData(i) = vNextStep
                        Else
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                Next
                
                '第一次设置时，即使有两行也是用空格合并了的
                .GetMergedRange .Row, .Col, lngR1, lngC1, lngR2, lngC2
                If lngC1 = lngC2 And lngR1 = lngR2 And lngR1 = .FixedRows + .FrozenRows - 1 And lngR1 - 1 = .FixedRows Then
                    '选中缺省分支的情况,处理该时间阶段的名称显示变化
                    .GetMergedRange .FixedRows, .Col, lngR1, lngC1, lngR2, lngC2
                End If
                
                '包括其他几种情况：横向合并的多个分支，纵向合并的一个时间阶段，纵向没合并的一个时间阶段
                '.Cell(flexcpText, lngR1, lngC1, lngR2, lngC2) = vStep.名称
                '如果直接范围赋值，因为包含回车会自动识别为分隔符，而导致文字被切断
                '如果不是编辑分支行，则不合并单元格，加一个空格
                If .Row = 1 Then
                    If vPreStep.名称 = vStep.名称 Then
                        vStep.名称 = vStep.名称 & " "
                        .ColData(.Col) = vStep
                    End If
                    If vNextStep.名称 = vStep.名称 Then
                        vStep.名称 = vStep.名称 & " "
                        .ColData(.Col) = vStep
                    End If
                End If
                For i = lngC1 To lngC2
                    For j = lngR1 To lngR2
                        .TextMatrix(j, i) = vStep.名称
                    Next
                Next
            End If
            
            .TextMatrix(.FixedRows - 1, .Col) = "阶段评估…"
            .Cell(flexcpFontBold, .FixedRows - 1, .Col) = False
            If Not vStep.评估.指标集 Is Nothing And Not Not vStep.评估.条件集 Is Nothing Then
                If vStep.评估.指标集.count > 0 Or vStep.评估.条件集.count > 0 Then
                    .Cell(flexcpFontBold, .FixedRows - 1, .Col) = True
                End If
            End If
            
            mblnChange = True
        End If
    End With
End Sub

Private Sub FuncStepInsert(ByVal intPos As Integer)
'功能：插入新的时间阶段
'参数：inPos=1：在当前时间阶段后面，-1：在当前时间阶段前面
    Dim lngR1 As Long, lngC1 As Long
    Dim lngR2 As Long, lngC2 As Long
   
    With vsPath
        If .TextMatrix(.Row, .Col) = "" Then
            MsgBox "当前阶段尚未设置，请先设置当前阶段信息。", vbInformation, gstrSysName
            Exit Sub
        End If
    
        '获取插入的位置
        .GetMergedRange .Row, .Col, lngR1, lngC1, lngR2, lngC2
        If lngC1 = lngC2 And lngR1 = lngR2 And lngR1 = .FixedRows + .FrozenRows - 1 And lngR1 - 1 = .FixedRows Then
            '选中分支的情况,GetMergedRange适用于合并范围的任何单元
            .GetMergedRange .FixedRows, .Col, 0, lngC1, 0, lngC2
        End If
        
        '插入新的时间阶段列
        .Redraw = flexRDNone
        
        .Cols = .Cols + 1
        .ColWidth(.Cols - 1) = COl_WIDTH_BASE
        
        If intPos = -1 Then
            .ColPosition(.Cols - 1) = lngC1
            .Col = lngC1
        Else
            .ColPosition(.Cols - 1) = lngC2 + 1
            .Col = lngC2 + 1
        End If
        
        Call SetTableCommonStyle(True)
         .Row = .FixedRows
         .ShowCell .Row, .Col
         
        .Redraw = flexRDDirect
        
        '插入之后进入内容设置
        Call FuncStepEdit
    End With
    
    mblnChange = True
End Sub

Private Sub FuncStepBranchInsert()
'功能：在当前时间阶段增加新的分支
    Dim lngR1 As Long, lngC1 As Long
    Dim lngR2 As Long, lngC2 As Long
    Dim vStep As TYPE_PATH_STEP
   
    With vsPath
        '获取插入的位置
        .GetMergedRange .Row, .Col, lngR1, lngC1, lngR2, lngC2
        If lngC1 = lngC2 And lngR1 = lngR2 And lngR1 = .FixedRows + .FrozenRows - 1 And lngR1 - 1 = .FixedRows Then
            '选中分支的情况,GetMergedRange适用于合并范围的任何单元
            .GetMergedRange .FixedRows, .Col, 0, lngC1, 0, lngC2
        End If
        
        .Redraw = flexRDNone
                
        '插入新的时间阶段列
        .Cols = .Cols + 1
        .ColWidth(.Cols - 1) = COl_WIDTH_BASE
        .ColPosition(.Cols - 1) = lngC2 + 1
        .Col = lngC2 + 1
        
        '设置缺省数据内容
        vStep = .ColData(.Col - 1)
        vStep.序号 = IIf(vStep.父ID <> 0, vStep.序号 + 1, 1) '分支序号保证1-N连续
        vStep.父ID = IIf(vStep.父ID <> 0, vStep.父ID, vStep.ID)
        vStep.ID = zlDatabase.GetNextId("临床路径阶段") '保证有新的唯一ID
        vStep.分类 = ""
        vStep.说明 = ""
        vStep.Edit = 1 '0-原始,1-新增,2-修改
        
        Set vStep.评估.条件集 = Nothing
        Set vStep.评估.指标集 = Nothing
        
        .ColData(.Col) = vStep
                
        '设置界面合并显示效果
        If .FrozenRows = 1 Then
            .AddItem .Cell(flexcpText, .FixedRows, .FixedCols, .FixedRows, .Cols - 1), .FixedRows + 1
            .FrozenRows = 2
        End If
        .TextMatrix(.FixedRows, .Col) = vStep.名称
        .TextMatrix(.FixedRows + .FrozenRows - 1, .Col) = IIf(vStep.说明 = "", "备用分支" & vStep.序号, vStep.说明) & IIf(vStep.分类 = "", "", ",") & vStep.分类
        If vStep.序号 = 1 Then
            .TextMatrix(.FixedRows + .FrozenRows - 1, .Col - 1) = "缺省分支"
        End If
        
        .Redraw = flexRDDirect
        
        .AutoSize .FixedCols, .Cols - 1, , 45 'Redraw后有效
        Call SetTableCommonStyle(True)
         .Row = .FixedRows + .FrozenRows - 1
         .ShowCell .Row, .Col
        
        '插入之后进入内容设置
        Call FuncStepEdit
    End With
    
    mblnChange = True
End Sub

Private Sub FuncStepDelete()
'功能：在当前时间阶段增加新的分支
    Dim lngR1 As Long, lngR2 As Long
    Dim lngC1 As Long, lngC2 As Long
    Dim vStep As TYPE_PATH_STEP
    Dim vSubStep As TYPE_PATH_STEP
    Dim vTmpStep As TYPE_PATH_STEP
    Dim lng父ID As Long, blnSub As Boolean
    Dim i As Long, j As Long
    Dim blnIsDelete As Boolean
    Dim vBranch As TYPE_PATH_BRANCH
    Dim objComboBranch As CommandBarComboBox
   
    With vsPath
        '获取选择范围
        .GetSelection lngR1, lngC1, lngR2, lngC2
        If lngC1 = lngC2 And lngR1 = lngR2 And lngR1 = .FixedRows + .FrozenRows - 1 And lngR1 - 1 = .FixedRows Then
            blnSub = True '选中分支的情况
        End If
        If Not blnSub Then
            .GetMergedRange .FixedRows, lngC2, 0, 0, 0, lngC2
        End If
        '检查当前阶段或之后的阶段是否有分支路径存在，是则禁止删除
        Set objComboBranch = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Branch, True)
        If objComboBranch Is Nothing Then Exit Sub
        If objComboBranch.ListIndex = 0 Then Exit Sub
        If objComboBranch.ListIndex = 1 Then
            For i = lngC1 To .Cols - 1
                If TypeName(.ColData(i)) <> "Empty" Then
                    vStep = .ColData(i)
                    For j = 2 To objComboBranch.ListCount
                        vBranch = mcolBranch("_" & objComboBranch.ItemData(j))
                        If vBranch.前一阶段ID = vStep.ID Then
                            MsgBox "删除的阶段或者之后的阶段存在分支路径，不允许删除。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    Next
                End If
            Next
        End If
        For i = lngC1 To lngC2
            If Replace(.Cell(flexcpText, .FixedRows + .FrozenRows, i, .Rows - 1, i), vbCr, "") <> "" Then
                If MsgBox("所选择的时间阶段(或分支)中存在已经定义的路径项目,删除阶段将同时删除这些项目，是否要继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                '删除路径项目
                For j = .FixedRows + .FrozenRows To .Rows - 1
                    If .TextMatrix(j, i) <> "" Then
                        .Row = j: .Col = i
                        Call FuncItemDelete(False)
                    End If
                Next
                blnIsDelete = True
                Exit For
            End If
        Next
        If Not blnIsDelete Then
            If MsgBox("确实要删除所选择的" & IIf(blnSub, "分支", "时间阶段") & "吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        .Redraw = flexRDNone
        
        '删除各列(反序)
        For i = lngC2 To lngC1 Step -1
            If TypeName(.ColData(i)) <> "Empty" Then
            
                vStep = .ColData(i)
                
                If vStep.父ID <> 0 Then
                    '调整后面的分支序号
                    For j = i + 1 To .Cols - 1
                        If TypeName(.ColData(j)) <> "Empty" Then
                            vSubStep = .ColData(j)
                            If vSubStep.父ID = vStep.父ID Then
                                vSubStep.序号 = vSubStep.序号 - 1
                                .TextMatrix(.FixedRows + .FrozenRows - 1, j) = IIf(vSubStep.说明 = "", "备用分支" & vSubStep.序号, vSubStep.说明) & IIf(vSubStep.分类 = "", "", ",") & vSubStep.分类
                                .ColData(j) = vSubStep
                            Else
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                ElseIf vStep.父ID = 0 Then
                    '缺省分支删除之后备选分支可以成为缺省分支
                    lng父ID = 0
                    For j = i + 1 To .Cols - 1
                        If TypeName(.ColData(j)) <> "Empty" Then
                            vSubStep = .ColData(j)
                            If vSubStep.父ID = vStep.ID Then
                                If j = i + 1 Then
                                    lng父ID = vSubStep.ID
                                    
                                    vSubStep.父ID = 0
                                    vSubStep.序号 = 0 '保存时重新生成
                                    .TextMatrix(.FixedRows + .FrozenRows - 1, j) = "缺省分支"
                                Else
                                    vSubStep.父ID = lng父ID
                                    vSubStep.序号 = vSubStep.序号 - 1
                                    .TextMatrix(.FixedRows + .FrozenRows - 1, j) = IIf(vSubStep.说明 = "", "备用分支" & vSubStep.序号, vSubStep.说明) & IIf(vSubStep.分类 = "", "", ",") & vSubStep.分类
                                End If
                                .ColData(j) = vSubStep
                            Else
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
                
                '记录删除操作:0-原始,1-新增,2-修改
                If vStep.Edit <> 1 Then
                    mstrDelStepIDs = mstrDelStepIDs & "," & vStep.ID
                End If
            End If
            
            .ColPosition(i) = .Cols - 1
            .Cols = .Cols - 1
        Next
        
        '检查分支的情况
        blnSub = False
        For i = .FixedCols + .FrozenCols To .Cols - 1
            If TypeName(.ColData(i)) <> "Empty" Then
                vStep = .ColData(i)
                If vStep.父ID <> 0 Then
                    blnSub = True: Exit For
                End If
            End If
        Next
        If Not blnSub Then
            If .FrozenRows > 1 Then
                .FrozenRows = 1
                .RemoveItem .FixedRows + .FrozenRows
            End If
        Else
            '清除无分支,但还显示了分支表头的内容
            For i = .FixedCols + .FrozenCols To .Cols - 1
                If TypeName(.ColData(i)) <> "Empty" Then
                    vStep = .ColData(i)
                    If vStep.父ID = 0 Then
                        If GetNearStep(i, 1, True, False).父ID <> vStep.ID Then
                            If .TextMatrix(.FixedRows + .FrozenRows - 1, i) <> .TextMatrix(.FixedRows, i) Then
                                .TextMatrix(.FixedRows + .FrozenRows - 1, i) = .TextMatrix(.FixedRows, i)
                            End If
                        End If
                    End If
                End If
            Next
        End If
        
        '新列定位
        If lngC1 <= .Cols - 1 Then
            .Col = lngC1
        ElseIf .Cols > .FixedCols + .FrozenCols Then
            .Col = .Cols - 1
        ElseIf .Cols = .FixedCols + .FrozenCols Then
            .Cols = .Cols + 1: .Col = .Cols - 1
            .ColWidth(.Cols - 1) = COl_WIDTH_BASE
        End If
        .Row = .FixedRows
                
        .ShowCell .Row, .Col
        .Redraw = flexRDDirect
    End With
    
    mblnChange = True
End Sub

Private Sub FuncItemEdit(Optional ByVal objControl As CommandBarControl)
'功能：设置当前路径项目内容
    Dim vStep As TYPE_PATH_STEP
    Dim vPreStep As TYPE_PATH_STEP
    Dim vItem As TYPE_PATH_ITEM
    Dim vBakItem As TYPE_PATH_ITEM
    Dim vPreItem As TYPE_PATH_ITEM
    Dim vTmpItem As TYPE_PATH_ITEM
    Dim blnInherit As Boolean
    Dim i As Long, j As Long
    Dim lng阶段ID As Long
    Dim blnAdjust As Boolean
    
    With vsPath
        If TypeName(.ColData(.Col)) = "Empty" Then
            MsgBox "请先设置当前项目位置所对应的时间阶段。", vbInformation, gstrSysName
            Exit Sub
        End If
        If Trim(.TextMatrix(.Row, .FixedCols)) = "" Then
            MsgBox "请先设置当前项目位置所对应的分类。", vbInformation, gstrSysName
            Exit Sub
        End If
        vStep = .ColData(.Col)
        
        If TypeName(.Cell(flexcpData, .Row, .Col)) <> "Empty" Then
            vItem = .Cell(flexcpData, .Row, .Col)
            vBakItem = vItem
        End If
        
        '获取前一个时间阶段相同项目的内容(包括分支)
        For i = .Col - 1 To .FixedCols + .FrozenCols Step -1
            If TypeName(.ColData(i)) <> "Empty" Then
                vPreStep = .ColData(i)
                If IIf(vPreStep.父ID <> 0, vPreStep.父ID, vPreStep.ID) <> IIf(vStep.父ID <> 0, vStep.父ID, vStep.ID) Then '不是当前阶段的
                    If lng阶段ID = 0 Then lng阶段ID = IIf(vPreStep.父ID <> 0, vPreStep.父ID, vPreStep.ID)
                    If IIf(vPreStep.父ID <> 0, vPreStep.父ID, vPreStep.ID) = lng阶段ID Then '前一个阶段如有分支,循环取最前面分支优先
                        For j = .FixedRows + .FrozenRows To .Rows - 1
                            If Trim(.TextMatrix(j, .FixedCols)) = Trim(.TextMatrix(.Row, .FixedCols)) Then '与当前同分类的
                                If TypeName(.Cell(flexcpData, j, i)) <> "Empty" Then
                                    vTmpItem = .Cell(flexcpData, j, i)
                                    If vTmpItem.项目内容 = vItem.项目内容 Or vItem.ID = 0 And j = .Row Then
                                        vPreItem = vTmpItem: Exit For
                                    End If
                                End If
                            End If
                        Next
                    Else
                        Exit For
                    End If
                End If
            Else
                Exit For '只取前面紧靠的时间阶段,没有则退出
            End If
        Next
        
        If Not objControl Is Nothing Then
            If objControl.Parameter = "Adjust" Then blnAdjust = True
        End If
        If mfrmPathItem.ShowEdit(Me, mrsAdvice, vItem, vPreItem, blnAdjust, blnInherit, mlng路径ID, mstrPrivs) Then
            If vItem.ID = 0 Then
                '保证有内容的项目ID不为空，先预取一个ID
                vItem.ID = zlDatabase.GetNextId("临床路径项目")
                vItem.Edit = 1 '0-原始,1-新增,2-修改
                '项目序号保存前自动处理
            Else
                If vItem.Edit = 0 Then vItem.Edit = 2
            End If
            
            If vItem.导入结果 = 1 Then
                .Cell(flexcpBackColor, .Row, .Col) = &H80000005
            End If
            '如果上下两个单元格项目内容相同，为了防止自动合并，加一个空格
            If .Row > 1 Then
                If TypeName(.Cell(flexcpData, .Row - 1, .Col)) <> "Empty" Then
                    vTmpItem = .Cell(flexcpData, .Row - 1, .Col)
                    If vTmpItem.项目内容 = vItem.项目内容 Then
                        vItem.项目内容 = vItem.项目内容 & " "
                        .Cell(flexcpData, .Row, .Col) = vItem
                    End If
                End If
            End If
            If .Row < .Rows - 1 Then
                If TypeName(.Cell(flexcpData, .Row + 1, .Col)) <> "Empty" Then
                    vTmpItem = .Cell(flexcpData, .Row + 1, .Col)
                    If vTmpItem.项目内容 = vItem.项目内容 Then
                        vItem.项目内容 = vItem.项目内容 & " "
                        .Cell(flexcpData, .Row, .Col) = vItem
                    End If
                End If
            End If
            
            '当前单元显示更新
            If vItem.图标ID <> 0 Then
                Set .Cell(flexcpPicture, .Row, .Col) = GetPathIcon(vItem.图标ID)
                .Cell(flexcpPictureAlignment, .Row, .Col) = 1
            Else
                Set .Cell(flexcpPicture, .Row, .Col) = Nothing
            End If
            .TextMatrix(.Row, .Col) = vItem.项目内容
            If vItem.医嘱IDs <> "" Or vItem.病历IDs <> "" Or vItem.新版病历IDs <> "" Then
                .TextMatrix(.Row, .Col) = .TextMatrix(.Row, .Col) & "…"
            End If
            .Cell(flexcpData, .Row, .Col) = vItem
            
            '继承关联的其他阶段相同项目更新(包括分支)
            '如果是新增设置项目,尚不存在这个继承关系
            '修改时如果项目内容变了则自动中断继承关系
            If blnInherit And vBakItem.ID <> 0 And vItem.项目内容 = vBakItem.项目内容 Then
                For i = .FixedCols + .FrozenCols To .Cols - 1
                    If TypeName(.ColData(i)) <> "Empty" Then
                        vPreStep = .ColData(i)
                        If vPreStep.ID <> vStep.ID Then '不是当前(分支)阶段的,但要处理同阶段其他分支
                            For j = .FixedRows + .FrozenRows To .Rows - 1
                                If Trim(.TextMatrix(j, .FixedCols)) = Trim(.TextMatrix(.Row, .FixedCols)) Then '与当前同分类的
                                    If TypeName(.Cell(flexcpData, j, i)) <> "Empty" Then
                                        vTmpItem = .Cell(flexcpData, j, i)
                                        If vTmpItem.项目内容 = vItem.项目内容 And vTmpItem.医嘱IDs = vBakItem.医嘱IDs Then
                                            vTmpItem.医嘱IDs = vItem.医嘱IDs
                                            vTmpItem.内容要求 = vItem.内容要求
                                            .Cell(flexcpData, j, i) = vTmpItem
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
            End If
            
            
            '调整界面
            .AutoSize .FixedCols, .Cols - 1, , 45
            Call SetTableCommonStyle(True)
            
            mblnChange = True
        End If
    End With
End Sub

Private Sub FuncItemInsert(ByVal intPos As Integer)
'功能：插入新的项目
'参数：inPos=1：在当前项目后面，-1：在当前项目前面
    Dim lngRow As Long, strCategory As String
    
    With vsPath
        If .TextMatrix(.Row, .Col) = "" Then
            MsgBox "当前项目尚未设置，请先设置当前项目内容。", vbInformation, gstrSysName
            Exit Sub
        End If
       
        strCategory = .TextMatrix(.Row, .FixedCols)
        lngRow = IIf(intPos = -1, .Row, .Row + 1)
        .AddItem "", lngRow
        .TextMatrix(lngRow, .FixedCols) = strCategory
        .RowHeight(lngRow) = ROW_HEIGHT_MIN
        
        .Row = lngRow
        .ShowCell .Row, .Col
        
        '进入设置
        Call FuncItemEdit
    End With
    
    mblnChange = True
End Sub

Private Sub FuncItemDelete(Optional ByVal blnIsMsg As Boolean = True)
'功能：删除当前选择的项目
'参数：blnIsMsg-是否弹出确认信息
    Dim lngR1 As Long, lngC1 As Long
    Dim lngR2 As Long, lngC2 As Long
    Dim lngRow As Long, lngCol As Long
    Dim i As Long, j As Long, k As Long
    Dim vItem As TYPE_PATH_ITEM
    Dim vStep As TYPE_PATH_STEP
    Dim vEvalCond As TYPE_PATH_EvalCond
    
    With vsPath
        If blnIsMsg Then
            If MsgBox("确实要删除所选择的路径项目吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        .Redraw = flexRDNone
                
        lngRow = .Row: lngCol = .Col
                
        '记录删除操作:0-原始,1-新增,2-修改
        .GetSelection lngR1, lngC1, lngR2, lngC2
        For i = lngC1 To lngC2
            If TypeName(.ColData(i)) <> "Empty" Then
                vStep = .ColData(i)
                
                For j = lngR1 To lngR2
                    If TypeName(.Cell(flexcpData, j, i)) <> "Empty" Then
                        vItem = .Cell(flexcpData, j, i)
                        If vItem.Edit <> 1 Then
                            mstrDelItemIDs = mstrDelItemIDs & "," & vItem.ID
                            
                            '删除阶段评估中使用的项目指标
                            If Not vStep.评估.条件集 Is Nothing Then
                                For k = vStep.评估.条件集.count To 1 Step -1
                                    vEvalCond = vStep.评估.条件集(k)
                                    If vEvalCond.项目ID = vItem.ID Then
                                        vStep.评估.条件集.Remove k
                                    End If
                                Next
                            End If
                        End If
                    End If
                Next
                                
                .ColData(i) = vStep
            End If
        Next
                
        '删除选中区域
        .GetSelection lngR1, lngC1, lngR2, lngC2
        .Cell(flexcpData, lngR1, lngC1, lngR2, lngC2) = Empty
        .Cell(flexcpText, lngR1, lngC1, lngR2, lngC2) = ""
        Set .Cell(flexcpPicture, lngR1, lngC1, lngR2, lngC2) = Nothing
        
        '清除没有设置项目的多余的分类行
        Call ClearCategoryRow

        '调整界面
        .Redraw = flexRDDirect
        .AutoSize .FixedCols, .Cols - 1, , 45 'Redraw后有效
        Call SetTableCommonStyle(True)
        
        '定位新行
        .Row = IIf(lngRow <= .Rows - 1, lngRow, .Rows - 1): .RowSel = .Row
        .Col = IIf(lngCol <= .Cols - 1, lngCol, .Cols - 1): .ColSel = .Col
        Call .ShowCell(.Row, .Col)
    End With
    
    mblnChange = True
End Sub

Private Sub FuncVersionDelete()
'功能：删除当前版本
    Dim objCombo As CommandBarComboBox
    Dim strSql As String
    
    Set objCombo = cbsMain(cbsMain.count - 1).FindControl(, cmd_Edit_Version, True)
    If objCombo Is Nothing Then Exit Sub
    If objCombo.ListIndex = 0 Then Exit Sub
    
    If MsgBox("确实要删除当前版本的临床路径吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    strSql = "Zl_临床路径版本_Delete(" & mlng路径ID & "," & objCombo.ItemData(objCombo.ListIndex) & ")"
    
    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    On Error GoTo 0
    
    Call LoadPathVersion
    
    '数据变化
    RaiseEvent DataChanged(mlng路径ID)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncBranchDelete()
'功能：删除当前分支
    Dim objComboBranch As CommandBarComboBox
    Dim objCombo As CommandBarComboBox
    Dim strSql As String
    
    Set objCombo = cbsMain(cbsMain.count - 1).FindControl(, cmd_Edit_Version, True)
    If objCombo Is Nothing Then Exit Sub
    If objCombo.ListIndex = 0 Then Exit Sub
    Set objComboBranch = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Branch, True)
    If objComboBranch Is Nothing Then Exit Sub
    If objComboBranch.ListIndex = 0 Then
        MsgBox "没有可删除的分支路径。", vbInformation, Me.Caption
        Exit Sub
    End If
    
    If MsgBox("确实要删除当前的路径分支吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    strSql = "Zl_临床路径分支_Delete(" & objComboBranch.ItemData(objComboBranch.ListIndex) & ")"
    
    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    On Error GoTo 0
    
    Call LoadPathTable(objCombo, objComboBranch)
    
    '数据变化
    RaiseEvent DataChanged(mlng路径ID)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncVersionCopy()
'功能：从其他路径复制覆盖当前版本
    Dim rsTmp As ADODB.Recordset
    Dim objCombo As CommandBarComboBox
    Dim intVersion As Integer, i As Long
    Dim strSql As String, blnCancel As Boolean
    Dim lng源路径ID As Long
    Dim objComboBranch As CommandBarComboBox
    Dim lng分支路径 As Long
    Dim vBranch As TYPE_PATH_BRANCH
    
    Set objCombo = cbsMain(cbsMain.count - 1).FindControl(, cmd_Edit_Version, True)
    If objCombo Is Nothing Then Exit Sub
    If objCombo.ListIndex = 0 Then Exit Sub
    Set objComboBranch = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Branch, True)
    If objComboBranch Is Nothing Then Exit Sub
    If objComboBranch.ListIndex = 0 Then Exit Sub
    
    If mblnChange Then
        MsgBox "路径表内容已被更改尚未保存，必须先保存后再导入。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error GoTo errH
    
    '选择其他临床路径
    If objComboBranch.ListIndex = 1 Then
        strSql = "Select ID,分类,编码,名称,最新版本,病例分型,适用病情," & _
            " Decode(Nvl(适用性别,0),0,'',1,'男',2,'女') as 适用性别,适用年龄,说明" & _
            " From 临床路径目录 A Where Nvl(最新版本,0)>0 And ID<>[1] "
        If InStr(mstrPrivs, "全院路径") = 0 Then
            strSql = strSql & _
                " And 通用=2 And Not Exists(" & _
                    " Select 科室ID From 临床路径科室 Where 路径ID=A.ID" & _
                    " Minus Select 部门ID From 部门人员 Where 人员ID=[2])"
        End If
        strSql = strSql & " Order by 分类,编码"
    Else
        strSql = "Select b.Id,'分支路径' as 分类, a.编码, b.名称, B.版本号, a.病例分型, a.适用病情, Decode(Nvl(a.适用性别, 0), 0, '', 1, '男', 2, '女') As 适用性别, a.适用年龄," & vbNewLine & _
                "       b.说明 As 说明,a.id as 路径ID" & vbNewLine & _
                "From 临床路径目录 A, 临床路径分支 B" & vbNewLine & _
                "Where a.Id = b.路径id And b.版本号=[3] And Nvl(a.最新版本, 0) > 0 And a.Id = [1] And b.ID<>[4]" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select 0 as id,'主路径' as 分类, 编码, 名称, 最新版本, 病例分型, 适用病情, Decode(Nvl(适用性别, 0), 0, '', 1, '男', 2, '女') As 适用性别, 适用年龄, 说明,a.id as 路径ID" & vbNewLine & _
                "From 临床路径目录 A" & vbNewLine & _
                "Where Nvl(最新版本, 0) > 0 And ID = [1]" & vbNewLine & _
                "Order By 分类 Desc, 编码"
    End If
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "临床路径", False, "", "", _
        False, False, False, 0, 0, 0, blnCancel, False, False, mlng路径ID, UserInfo.ID, objCombo.ItemData(objCombo.ListIndex), objComboBranch.ItemData(objComboBranch.ListIndex))
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "没有找到其他可用的路径表。", vbInformation, gstrSysName
        End If
        Exit Sub
    End If
    
    If objComboBranch.ListIndex = 1 Then
        lng源路径ID = rsTmp!ID
        If objComboBranch.ListCount > 1 Then
            If MsgBox("从其他路径复制时，会删除当前路径的所有分支，你确定要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        End If
    Else
        lng源路径ID = rsTmp!路径ID
    End If
    intVersion = objCombo.ItemData(objCombo.ListIndex)
    lng分支路径 = Nvl(rsTmp!ID, 0)
    mstrDelStepIDs = "": mstrDelItemIDs = "": mblnChange = False
    
    '复制指定路径最新版本覆盖当前版本内容
    strSql = "Zl_临床路径版本_Copy(" & lng源路径ID & ",0," & mlng路径ID & "," & intVersion & "," & lng分支路径 & "," & IIf(objComboBranch.ListIndex = 1, 0, 1) & "," & objComboBranch.ItemData(objComboBranch.ListIndex) & ")"
    
    '提交数据
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    '刷新界面
    If objComboBranch.ListIndex = 1 Then
        Call LoadPathVersion
    Else
        vBranch = mcolBranch("_" & objComboBranch.ItemData(objComboBranch.ListIndex))
        Call LoadPathTable(objCombo, objComboBranch, , vBranch.分支名称)
    End If
    
    '数据变化
    RaiseEvent DataChanged(mlng路径ID)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncVersionNew()
'功能：删除当前版本
    Dim vVersion As TYPE_PATH_VERSION
    Dim objCombo As CommandBarComboBox
    Dim intVersion As Integer, strSql As String
    Dim i As Long
    
    Set objCombo = cbsMain(cbsMain.count - 1).FindControl(, cmd_Edit_Version, True)
    If objCombo Is Nothing Then Exit Sub
    If objCombo.ListIndex = 0 Then Exit Sub
    
    If mblnChange Then
        If MsgBox("路径表内容已被更改尚未保存，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    intVersion = objCombo.ItemData(objCombo.ListIndex)
    If MsgBox("要复制当前选择版本的内容产生新版本吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then intVersion = 0
    
    mstrDelStepIDs = "": mstrDelItemIDs = "": mblnChange = False
    
    If intVersion > 0 Then
        '复制当前选择版本内容产生新版本内容
        strSql = "Zl_临床路径版本_Copy(" & mlng路径ID & "," & intVersion & "," & mlng路径ID & ",0)"
        
        '提交数据
        On Error GoTo errH
        zlDatabase.ExecuteProcedure strSql, Me.Caption
        On Error GoTo 0
        
        '刷新界面
        Call LoadPathVersion
        
        '数据变化
        RaiseEvent DataChanged(mlng路径ID)
    Else
        '增加空的新内容
        objCombo.AddItem "正在设计中……", 1
        objCombo.ListIndex = 1
        mcolVersion.Add vVersion, "_0"
        Call cbsMain_Execute(objCombo)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncBranchNew()
'功能:新增分支路径
    Dim vVersion As TYPE_PATH_VERSION
    Dim vBranch As TYPE_PATH_BRANCH
    Dim objCombo As CommandBarComboBox
    Dim objComboBranch As CommandBarComboBox
    Dim intVersion As Integer, strSql As String
    Dim i As Long, lng分支ID As Long
    Dim lngNewId As Long
    
    Set objCombo = cbsMain(cbsMain.count - 1).FindControl(, cmd_Edit_Version, True)
    If objCombo Is Nothing Then Exit Sub
    If objCombo.ListIndex = 0 Then Exit Sub
    Set objComboBranch = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Branch, True)
    If objComboBranch Is Nothing Then Exit Sub
    If objComboBranch.ListIndex = 0 Then
        MsgBox "请先设置主路径，再新增分支路径。", vbInformation, Me.Caption
        Exit Sub
    End If
    
    If mblnChange Then
        If MsgBox("路径表内容已被更改尚未保存，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    intVersion = objCombo.ItemData(objCombo.ListIndex)
    vVersion = mcolVersion("_" & intVersion)
    
    mstrDelStepIDs = "": mstrDelItemIDs = "": mblnChange = False
    mblnAddNew = True
    If intVersion > 0 Then
        '弹出标准设置界面
        vBranch.版本号 = intVersion
        If vsPath.Col > 0 Then
            If TypeName(vsPath.ColData(vsPath.Col)) <> "Empty" Then
                lngNewId = vsPath.ColData(vsPath.Col).ID
            End If
        End If
        If mfrmVersion.ShowMe(Me, vVersion, vBranch, mlng路径ID, lngNewId) Then
            objComboBranch.AddItem "分支名称：" & vBranch.分支名称 & " ，" & "前一阶段：" & vBranch.前一阶段名称 & " ，" & _
                "创建：" & vBranch.创建人 & "/正在设计中……"
            vVersion = mcolVersion("_" & intVersion)
            stcInfo.Caption = "标准住院日：" & IIf(vVersion.标准住院日 <> "", vVersion.标准住院日 & "天", "<未设定>") & _
                        "，标准费用：" & IIf(vVersion.标准费用 <> "", vVersion.标准费用 & "元", "<未设定>") & _
                        "，说明：" & IIf(vVersion.版本说明 <> "", vVersion.版本说明, "<无>") & IIf(vBranch.分支名称 = "主路径" Or vBranch.版本号 = 0, "", ("   分支名称：" & _
                        IIf(vBranch.分支名称 <> "", vBranch.分支名称, "<未设定>") & _
                        " 标准住院日：" & IIf(vBranch.标准住院日 <> "", vBranch.标准住院日 & "天", "<未设定>") & _
                        "，标准费用：" & IIf(vBranch.标准费用 <> "", vBranch.标准费用 & "元", "<未设定>") & _
                        "，说明：" & IIf(vBranch.说明 <> "", vBranch.说明, "<无>")))
            lng分支ID = zlDatabase.GetNextId("临床路径分支")
            objComboBranch.ItemData(objComboBranch.ListCount) = lng分支ID
            vBranch.分支ID = lng分支ID
            mcolBranch.Add vBranch, "_" & vBranch.分支ID
            With vsPath
                .FixedRows = 1
                If .FrozenRows = 2 Then .FrozenRows = 1: .RemoveItem 1
                .Cols = 1: .FixedCols = 0: .FrozenCols = 1: .Cols = 1 + 1
                .ColWidth(-1) = COl_WIDTH_BASE: .ColWidth(0) = 1000
                vsPath.AutoSize vsPath.FixedCols, vsPath.Cols - 1, , 45
            End With
            objComboBranch.ListIndex = objComboBranch.ListCount
            mblnChange = True
        End If
    End If
    mblnAddNew = False
    Exit Sub
errH:
    mblnAddNew = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncVersionAudit(ByVal intNum As Integer)
'功能：审核/取消审核当前版本
'参数：blnAudit=审核/取消审核
'参数：int场合 1=医务科审核 -1=医务科取消审核 2=药剂科审核 -2=药剂科取消审核
    Dim objCombo As CommandBarComboBox
    Dim strSql As String
    
    Set objCombo = cbsMain(cbsMain.count - 1).FindControl(, cmd_Edit_Version, True)
    If objCombo Is Nothing Then Exit Sub
    If objCombo.ListIndex = 0 Then Exit Sub
    
    If intNum = -1 Or intNum = -2 Then
        If MsgBox("确实要取消审核当前版本的临床路径吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    
    strSql = "Zl_临床路径版本_Audit(" & mlng路径ID & "," & objCombo.ItemData(objCombo.ListIndex) & "," & intNum & ")"
    
    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    On Error GoTo 0
    
    Call LoadPathVersion(objCombo.ItemData(objCombo.ListIndex))
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncVersionEdit()
'功能：设置当前版本相关信息
    Dim objCombo As CommandBarComboBox
    Dim objComboBranch As CommandBarComboBox
    Dim vVersion As TYPE_PATH_VERSION
    Dim vStep As TYPE_PATH_STEP
    Dim vBranch As TYPE_PATH_BRANCH
    Dim i As Long, j As Long
    Dim str天数 As String
    Dim str原始 As String
    
    Set objCombo = cbsMain(cbsMain.count - 1).FindControl(, cmd_Edit_Version, True)
    If objCombo Is Nothing Then Exit Sub
    If objCombo.ListIndex = 0 Then Exit Sub
    Set objComboBranch = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Branch, True)
    If objComboBranch Is Nothing Then Exit Sub
    vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))
    If objComboBranch.ListIndex = 1 Or objComboBranch.ListIndex = 0 Then
        If vVersion.标准住院日 = "" Then
            With vsPath
                For i = .Cols - 1 To .FixedCols + .FrozenCols Step -1
                    If TypeName(.ColData(i)) <> "Empty" Then
                        vStep = .ColData(i)
                        If vStep.开始天数 <> 0 Then
                            If vStep.结束天数 <> 0 Then
                                vVersion.标准住院日 = vStep.结束天数
                            Else
                                vVersion.标准住院日 = vStep.开始天数
                            End If
                            Exit For
                        End If
                    End If
                Next
            End With
        End If
        
        If mfrmVersion.ShowMe(Me, vVersion, vBranch, mlng路径ID) Then
            mcolVersion.Remove "_" & objCombo.ItemData(objCombo.ListIndex)
            mcolVersion.Add vVersion, "_" & objCombo.ItemData(objCombo.ListIndex)
            
            stcInfo.Caption = _
                "标准住院日：" & IIf(vVersion.标准住院日 <> "", vVersion.标准住院日 & "天", "<未设定>") & _
                "，标准费用：" & IIf(vVersion.标准费用 <> "", vVersion.标准费用 & "元", "<未设定>") & _
                "，说明：" & IIf(vVersion.版本说明 <> "", vVersion.版本说明, "<无>")
            
            mblnChange = True
        End If
    Else
        vBranch = mcolBranch("_" & objComboBranch.ItemData(objComboBranch.ListIndex))
        If vBranch.标准住院日 = "" Then
            With vsPath
                For i = .Cols - 1 To .FixedCols + .FrozenCols Step -1
                    If TypeName(.ColData(i)) <> "Empty" Then
                        vStep = .ColData(i)
                        If vStep.开始天数 <> 0 Then
                            If vStep.结束天数 <> 0 Then
                                vBranch.标准住院日 = vStep.结束天数
                            Else
                                vBranch.标准住院日 = vStep.开始天数
                            End If
                            Exit For
                        End If
                    End If
                Next
            End With
        End If
        '记录变化的标准住院日开始天数
        mlngDays = IIf(InStr(vBranch.标准住院日, "-") > 0, Val(Split(vBranch.标准住院日, "-")(0)), Val(vBranch.标准住院日))
        If mfrmVersion.ShowMe(Me, vVersion, vBranch, mlng路径ID) Then
            mcolBranch.Remove "_" & objComboBranch.ItemData(objComboBranch.ListIndex)
            mcolBranch.Add vBranch, "_" & objComboBranch.ItemData(objComboBranch.ListIndex)

            stcInfo.Caption = _
                "标准住院日：" & IIf(vVersion.标准住院日 <> "", vVersion.标准住院日 & "天", "<未设定>") & _
                    "，标准费用：" & IIf(vVersion.标准费用 <> "", vVersion.标准费用 & "元", "<未设定>") & _
                    "，说明：" & IIf(vVersion.版本说明 <> "", vVersion.版本说明, "<无>") & IIf(vBranch.分支名称 = "主路径" Or vBranch.版本号 = 0, "", ("   分支名称：" & _
                    IIf(vBranch.分支名称 <> "", vBranch.分支名称, "<未设定>") & _
                    " 标准住院日：" & IIf(vBranch.标准住院日 <> "", vBranch.标准住院日 & "天", "<未设定>") & _
                    "，标准费用：" & IIf(vBranch.标准费用 <> "", vBranch.标准费用 & "元", "<未设定>") & _
                    "，说明：" & IIf(vBranch.说明 <> "", vBranch.说明, "<无>")))
            
            '根据变化的天数，设置各个阶段的增量
            If mlngDays <> 0 Then
                With vsPath
                    For i = .Cols - 1 To .FixedCols + .FrozenCols Step -1
                        If TypeName(.ColData(i)) <> "Empty" Then
                            vStep = .ColData(i)
                            If vStep.开始天数 <> vStep.结束天数 And vStep.结束天数 <> 0 Then
                                str原始 = vStep.开始天数 & "-" & vStep.结束天数
                            Else
                                str原始 = vStep.开始天数
                            End If
                            vStep.开始天数 = vStep.开始天数 + mlngDays
                            If vStep.结束天数 <> 0 Then vStep.结束天数 = vStep.结束天数 + mlngDays
                            
                            If vStep.开始天数 <> vStep.结束天数 And vStep.结束天数 <> 0 Then
                                str天数 = vStep.开始天数 & "-" & vStep.结束天数
                            Else
                                str天数 = vStep.开始天数
                            End If
                            For j = 1 To .FrozenRows
                                .TextMatrix(j, i) = Replace(.TextMatrix(1, i), str原始, str天数)
                                vStep.名称 = .TextMatrix(j, i)
                            Next
                            vStep.Edit = 2 '阶段天数发生变化
                            .ColData(i) = vStep
                        End If
                    Next
                End With
                mlngDays = 0
            End If
            mblnChange = True
        End If
    End If
End Sub

Private Sub FuncVersionStop(ByVal blnStop As Boolean)
'功能：停用/取消停用当前版本
'参数：blnAudit=审核/取消审核
    Dim objCombo As CommandBarComboBox
    Dim strSql As String
    
    Set objCombo = cbsMain(cbsMain.count - 1).FindControl(, cmd_Edit_Version, True)
    If objCombo Is Nothing Then Exit Sub
    If objCombo.ListIndex = 0 Then Exit Sub
    
    If blnStop Then
        If MsgBox("确实要停用当前版本的临床路径吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    Else
        If MsgBox("确实要取消停用当前版本的临床路径吗？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    
    strSql = "Zl_临床路径版本_Stop(" & mlng路径ID & "," & objCombo.ItemData(objCombo.ListIndex) & "," & IIf(blnStop, 1, -1) & ")"
    
    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    On Error GoTo 0
    
    Call LoadPathVersion(objCombo.ItemData(objCombo.ListIndex))
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncEvaluateImport()
'功能：设置导入评估
    If mfrmEvalEdit.ShowEdit(Me, 1, mvEvalImport) Then
        mblnChange = True
    End If
End Sub

Private Sub FuncEvaluateStep(Optional ByVal blnCopy As Boolean)
'功能：设置阶段评估
    Dim vStep As TYPE_PATH_STEP
    Dim vEval As TYPE_PATH_EVAL
    Dim vEvalPre As TYPE_PATH_EVAL
    Dim vEvalMark As TYPE_PATH_EvalMark
    Dim vEvalCond As TYPE_PATH_EvalCond
    Dim colMarkID As New Collection
    Dim vItem As TYPE_PATH_ITEM
    Dim colItems As New Collection
    Dim lngC1 As Long, lngC2 As Long
    Dim lngNewId As Long, i As Long, j As Long
    
    With vsPath
        If mlng性质 = 1 Then
            MsgBox "合并路径不需要设置阶段评估信息。", vbInformation, Me.Caption
            Exit Sub
        End If
        If .Col >= .FixedCols + .FrozenCols Then
            If .Row = .FixedRows Then
                .GetMergedRange .Row, .Col, 0, lngC1, 0, lngC2
                If lngC1 <> lngC2 Then
                    .Row = .FixedRows + .FrozenRows - 1
                End If
            End If
        End If
        
        If TypeName(.ColData(.Col)) = "Empty" Then
            MsgBox "请先设置当前时间阶段的信息。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If blnCopy Then
            For i = .Col - 1 To .FixedCols + .FrozenCols Step -1
                If TypeName(.ColData(i)) <> "Empty" Then
                    vStep = .ColData(i)
                    If Not vStep.评估.指标集 Is Nothing Then
                        If vStep.评估.指标集.count > 0 Then
                            vEvalPre = vStep.评估
                            Set vEval.指标集 = New Collection
                            Set vEval.条件集 = New Collection
                            
                            '收集指标
                            For j = 1 To vEvalPre.指标集.count
                                vEvalMark = vEvalPre.指标集(j)
                                
                                lngNewId = zlDatabase.GetNextId("路径评估指标")
                                colMarkID.Add lngNewId, "_" & vEvalMark.ID
                                vEvalMark.ID = lngNewId
                                
                                vEval.指标集.Add vEvalMark
                            Next
                            
                            '收集计算条件
                            If Not vEvalPre.条件集 Is Nothing Then
                                For j = 1 To vEvalPre.条件集.count
                                    vEvalCond = vEvalPre.条件集(j)
                                    If vEvalCond.指标ID <> 0 Then
                                        vEvalCond.指标ID = colMarkID("_" & vEvalCond.指标ID)
                                        vEval.条件集.Add vEvalCond
                                    End If
                                Next
                            End If
                            
                            Exit For
                        End If
                    End If
                End If
            Next
            If vEval.指标集 Is Nothing And vEval.条件集 Is Nothing Then
                MsgBox "前面的时间阶段中没有可以复制的评估设置。", vbInformation, gstrSysName
                Exit Sub
            End If
            vStep = .ColData(.Col)
        Else
            vStep = .ColData(.Col)
            vEval = vStep.评估
        End If
        
        '本阶段的项目(可能作为计算指标)
        For i = .FixedRows + .FrozenRows To .Rows - 1
            If TypeName(.Cell(flexcpData, i, .Col)) <> "Empty" Then
                vItem = .Cell(flexcpData, i, .Col)
                colItems.Add vItem
            End If
        Next
    End With
    
    If mfrmEvalEdit.ShowEdit(Me, 2, vEval, colItems) Then
        With vsPath
            vStep.评估 = vEval
            '0-原始,1-新增,2-修改
            If vStep.Edit = 0 Then vStep.Edit = 2
            .ColData(.Col) = vStep
            
            .TextMatrix(.FixedRows - 1, .Col) = "阶段评估…"
            .Cell(flexcpFontBold, .FixedRows - 1, .Col) = False
            If vStep.评估.指标集.count > 0 Or vStep.评估.条件集.count > 0 Then
                .Cell(flexcpFontBold, .FixedRows - 1, .Col) = True
            End If
        End With
        mblnChange = True
    End If
End Sub

Private Sub FuncClipboradCopy()
'功能：复制当前选择的项目信息到内部剪贴板
'说明：只能对同一阶段中的一个或多个项目进行复制
    Dim vStep As TYPE_PATH_STEP
    Dim vItem As TYPE_PATH_ITEM
    Dim vNullItem As TYPE_PATH_ITEM
    Dim lngR1 As Long, lngR2 As Long
    Dim lngC1 As Long, lngC2 As Long
    Dim i As Long
    
    With vsPath
        .GetSelection lngR1, lngC1, lngR2, lngC2
        If lngC1 <> lngC2 Then
            MsgBox "没有内容被复制。", vbInformation, gstrSysName
            Exit Sub
        End If
        If TypeName(.ColData(lngC1)) = "Empty" Then
            MsgBox "没有内容被复制。", vbInformation, gstrSysName
            Exit Sub
        End If
        vStep = .ColData(lngC1)
        
        ReDim mvClipboard.项目集(lngR2 - lngR1)
        For i = lngR1 To lngR2
            If TypeName(.Cell(flexcpData, i, lngC1)) <> "Empty" Then
                vItem = .Cell(flexcpData, i, lngC1)
            Else
                vItem = vNullItem
            End If
            mvClipboard.项目集(i - lngR1) = vItem
        Next
        mvClipboard.Empty = False
        mvClipboard.vStep = vStep
        mvClipboard.BeginRow = lngR1
    End With
End Sub

Private Sub FuncClipboradPaste()
'功能：从内部剪贴板粘贴内容到当前选择区域
'说明：只能对同一阶段中的一个或多个项目进行复制
    Dim vItem1 As TYPE_PATH_ITEM
    Dim vItem2 As TYPE_PATH_ITEM
    Dim vNullItem As TYPE_PATH_ITEM
    Dim lngR1 As Long, lngC1 As Long
    Dim lngR2 As Long, lngC2 As Long
    Dim i As Long
    Dim vStep As TYPE_PATH_STEP
    Dim lngThis As Long, lngThat As Long, blnInherit As Boolean
    
    If mvClipboard.Empty Then
        MsgBox "剪贴板是空的。", vbInformation, gstrSysName
        Exit Sub
    End If
    With vsPath
        .GetSelection lngR1, lngC1, lngR2, lngC2
        
        If lngC2 <> lngC1 Then
            MsgBox "要粘贴数据的目标选择区域不符合要求，只能对一个时间阶段中的项目进行复制粘贴。", vbInformation, gstrSysName
            Exit Sub
        End If
        If TypeName(.ColData(lngC1)) = "Empty" Then
            MsgBox "请先设置当前项目位置所对应的时间阶段。", vbInformation, gstrSysName
            Exit Sub
        End If
        If .Rows - lngR1 < UBound(mvClipboard.项目集) + 1 Then
            MsgBox "目标区域太小，不足于粘贴所复制的源项目数据。", vbInformation, gstrSysName
            Exit Sub
        End If
        If MsgBox("确实要粘贴所复制的项目数据覆盖当前目标区域吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        vStep = .ColData(lngC1)
        '粘贴数据
        .Redraw = flexRDNone
        For i = 0 To UBound(mvClipboard.项目集)
            vItem1 = mvClipboard.项目集(i)
            If TypeName(.Cell(flexcpData, lngR1, lngC1)) <> "Empty" Then
                vItem2 = .Cell(flexcpData, lngR1, lngC1)
            Else
                vItem2 = vNullItem
            End If
            
            'Edit：0-原始,1-新增,2-修改
            If vItem1.ID <> 0 Then
                If vItem2.ID <> 0 Then
                    vItem1.ID = vItem2.ID
                    vItem1.Edit = vItem2.Edit
                    If vItem1.Edit = 0 Then vItem1.Edit = 2
                Else
                    vItem1.ID = zlDatabase.GetNextId("临床路径项目")
                    vItem1.Edit = 1
                End If
                
                '如果有对应医嘱，产生为独立的新医嘱
                If vItem1.医嘱IDs <> "" Then
                    '从前一阶段复制时，继承相同项目的长嘱(同一分类下的项目才继承)
                    If vStep.父ID <> 0 Then
                        lngThis = GetParentStep(vStep).序号
                    Else
                        lngThis = vStep.序号
                    End If
                    If mvClipboard.vStep.父ID <> 0 Then
                        lngThat = GetParentStep(mvClipboard.vStep).序号
                    Else
                        lngThat = mvClipboard.vStep.序号
                    End If
                    If lngThis = lngThat + 1 Then
                        blnInherit = .TextMatrix(mvClipboard.BeginRow + i, .FixedCols) = .TextMatrix(lngR1, .FixedCols)
                    Else
                        blnInherit = False
                    End If
                    
                    vItem1.医嘱IDs = AdviceCopyNew(vItem1.医嘱IDs, blnInherit)
                End If
                
                .Cell(flexcpData, lngR1, lngC1) = vItem1
                
                .TextMatrix(lngR1, lngC1) = vItem1.项目内容
                If vItem1.医嘱IDs <> "" Or vItem1.病历IDs <> "" Or vItem1.新版病历IDs <> "" Then
                    .TextMatrix(lngR1, lngC1) = .TextMatrix(lngR1, lngC1) & "…"
                End If
                
                If vItem1.图标ID <> 0 Then
                    Set .Cell(flexcpPicture, lngR1, lngC1) = GetPathIcon(vItem1.图标ID)
                    .Cell(flexcpPictureAlignment, lngR1, lngC1) = 1
                End If
            Else
                .Cell(flexcpData, lngR1, lngC1) = Empty
                .TextMatrix(lngR1, lngC1) = ""
                Set .Cell(flexcpPicture, lngR1, lngC1) = Nothing
                
                '记录删除操作
                If vItem2.ID <> 0 And vItem2.Edit <> 1 Then
                    mstrDelItemIDs = mstrDelItemIDs & "," & vItem2.ID
                End If
            End If
            
            lngR1 = lngR1 + 1
        Next
        .GetSelection lngR1, lngC1, lngR2, lngC2
        .Select lngR1, lngC1, lngR1 + UBound(mvClipboard.项目集), lngC2
        .ShowCell .Row, .Col
        .Redraw = flexRDDirect
        
        '调整界面
        .AutoSize .FixedCols, .Cols - 1, , 45 'Redraw后有效
        Call SetTableCommonStyle(True)

        mblnChange = True
    End With
End Sub

Private Function AdviceCopyNew(ByVal str医嘱ID As String, ByVal blnInherit As Boolean) As String
'功能：根据医嘱ID复制产生新的医嘱
'参数：blnInherit=从前一列的单元格拷贝时，继承相同项目的长嘱
    Dim rsCopy As ADODB.Recordset
    Dim strFilter As String, i As Long, arrAdvice As Variant
    Dim colAdviceID As New Collection
    Dim lngAdviceID As Long, strAdviceID As String, blnAllLongAdvice As Boolean
    Dim strSql As String
    Dim objCombo As CommandBarComboBox
    
    Set objCombo = cbsMain(cbsMain.count - 1).FindControl(, cmd_Edit_Version, True)
    If objCombo Is Nothing Then Exit Function
    If objCombo.ListIndex = 0 Then Exit Function
    Set rsCopy = mrsAdvice.Clone
    
    arrAdvice = Split(str医嘱ID, ",")
    For i = 0 To UBound(arrAdvice)
        strFilter = strFilter & " Or ID=" & arrAdvice(i)
    Next
    rsCopy.Filter = Mid(strFilter, 5)
    
    If rsCopy.RecordCount = 0 Then
        '如果复制时没有记录，则从数据中找
        strSql = " Select /*+ Rule*/ Distinct A.ID,A.相关ID,A.序号,A.期效,A.诊疗项目ID,A.收费细目ID," & _
                " A.医嘱内容,A.单次用量,A.总给予量,A.标本部位,A.检查方法,A.医生嘱托,A.执行标记, " & _
                " A.执行频次,A.频率次数,A.频率间隔,A.间隔单位,A.执行性质,A.执行科室ID,A.时间方案,A.是否缺省,A.是否备选,A.配方ID,A.组合项目ID" & _
                " From 路径医嘱内容 A,临床路径医嘱 B,临床路径项目 C" & _
                " Where A.ID=B.医嘱内容ID And B.路径项目ID=C.ID And C.路径ID=[1] And C.版本号=[2] And a.ID In (Select * From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist))) " & _
                " Order by A.序号,A.ID"
        On Error GoTo errH
        Set rsCopy = zlDatabase.OpenSQLRecord(strSql, "复制项目", mlng路径ID, objCombo.ItemData(objCombo.ListIndex), str医嘱ID)
    End If
    If rsCopy.RecordCount = 0 Then Exit Function
    If blnInherit Then
        blnAllLongAdvice = True
        For i = 1 To rsCopy.RecordCount
            If rsCopy!期效 = 1 Then
                blnAllLongAdvice = False
                Exit For
            End If
            rsCopy.MoveNext
        Next
        rsCopy.MoveFirst
    End If
    
    If blnAllLongAdvice Then
        '继承前一阶段相同项目的长嘱
        strAdviceID = "," & str医嘱ID
    Else
        '先产生新的医嘱ID
        Do While Not rsCopy.EOF
            lngAdviceID = zlDatabase.GetNextId("路径医嘱内容")
            colAdviceID.Add lngAdviceID, "_" & rsCopy!ID
            strAdviceID = strAdviceID & "," & lngAdviceID
            rsCopy.MoveNext
        Loop
    
        rsCopy.MoveFirst: i = 1
        Do While Not rsCopy.EOF
            lngAdviceID = colAdviceID("_" & rsCopy!ID)
            mrsAdvice.AddNew
            mrsAdvice!ID = lngAdviceID
            If Not IsNull(rsCopy!相关id) Then
                mrsAdvice!相关id = colAdviceID("_" & rsCopy!相关id)
            End If
            mrsAdvice!序号 = i
            mrsAdvice!期效 = rsCopy!期效
            mrsAdvice!诊疗项目ID = rsCopy!诊疗项目ID
            mrsAdvice!收费细目ID = rsCopy!收费细目ID
            If IsNull(rsCopy!诊疗项目ID) Then
                mrsAdvice!医嘱内容 = rsCopy!医嘱内容 '自由录入医嘱才保存
            End If
            mrsAdvice!单次用量 = rsCopy!单次用量
            mrsAdvice!总给予量 = rsCopy!总给予量
            mrsAdvice!医生嘱托 = rsCopy!医生嘱托
            mrsAdvice!执行频次 = rsCopy!执行频次
            mrsAdvice!频率次数 = rsCopy!频率次数
            mrsAdvice!频率间隔 = rsCopy!频率间隔
            mrsAdvice!间隔单位 = rsCopy!间隔单位
            mrsAdvice!时间方案 = rsCopy!时间方案
            mrsAdvice!执行科室ID = rsCopy!执行科室ID
            mrsAdvice!执行性质 = rsCopy!执行性质
            mrsAdvice!标本部位 = rsCopy!标本部位
            mrsAdvice!检查方法 = rsCopy!检查方法
            mrsAdvice!配方ID = rsCopy!配方ID
            mrsAdvice!组合项目ID = rsCopy!组合项目ID
            mrsAdvice!执行标记 = rsCopy!执行标记
            
            mrsAdvice.Update
            
            i = i + 1
            rsCopy.MoveNext
        Loop
    End If
        
    AdviceCopyNew = Mid(strAdviceID, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ClearCategoryRow()
'功能：清除没有设置项目的多余的分类行
    Dim lngRow As Long
    Dim lngR1 As Long, lngR2 As Long
    Dim i As Long, j As Long
    Dim vRedraw As RedrawSettings
    
    With vsPath
        vRedraw = .Redraw: .Redraw = flexRDNone
        lngRow = .Row
        i = .Rows - 1
        Do While i >= .FixedRows + .FrozenRows
            .GetMergedRange i, .FixedCols, lngR1, 0, lngR2, 0
            If Replace(Replace(.Cell(flexcpText, lngR1, .FixedCols, lngR2, .FixedCols), vbTab, ""), vbCr, "") <> "" Then
                For j = lngR2 To lngR1 Step -1
                    If Replace(.Cell(flexcpText, j, .FixedCols, j, .FixedCols), vbTab, "") = "" Then
                        .RemoveItem j
                    End If
                Next
            End If
            
            i = lngR1 - 1
        Loop
        .Row = IIf(lngRow <= .Rows - 1, lngRow, .Rows - 1)
        .ShowCell .Row, .Col
        .Redraw = vRedraw
    End With
End Sub

Private Function CheckPathTable() As Boolean
'功能：检查路径表数据输入的合法性
    Dim lngR1 As Long, lngR2 As Long
    Dim i As Long, j As Long
    Dim strMsg As String
    Dim strPathItems As String
    Dim strAdviceIDs As String
    
    Dim objCombo As CommandBarComboBox
    Dim objComboBranch As CommandBarComboBox
    Dim vVersion As TYPE_PATH_VERSION
    Dim vBranch As TYPE_PATH_BRANCH
    Dim vStep As TYPE_PATH_STEP
    Dim vItem As TYPE_PATH_ITEM
    Dim lng阶段序号 As Long
    Dim lng分支序号 As Long
    Dim lng项目序号 As Long
    Dim strSql As String, rsTmp As Recordset
    
    With vsPath
        '没有设置的阶段
        For i = .FixedCols + .FrozenCols To .Cols - 1
            If TypeName(.ColData(i)) = "Empty" Then
                .Row = .FixedRows: .Col = i
                Call .ShowCell(.Row, .Col)
                MsgBox "该阶段的信息尚未进行设置。", vbInformation, gstrSysName
                Exit Function
            End If
        Next
        
        '设置设置的分类
        For i = .FixedRows + .FrozenRows To .Rows - 1
            If Trim(.TextMatrix(i, .FixedCols)) = "" Then
                .Row = i: .Col = .FixedCols
                Call .ShowCell(.Row, .Col)
                MsgBox "该分类的名称尚未输入。", vbInformation, gstrSysName
                Exit Function
            End If
        Next
        
        '没有设置项目的阶段或者分类(允许)
        strMsg = ""
        For i = .FixedCols + .FrozenCols To .Cols - 1
            If TypeName(.ColData(i)) <> "Empty" Then
                If Replace(.Cell(flexcpText, .FixedRows + .FrozenRows, i, .Rows - 1, i), vbCr, "") = "" Then
                    strMsg = "发现存在尚未设置路径项目的阶段或者分类，要继续吗？"
                    Exit For
                End If
            End If
        Next
        i = .FixedRows + .FrozenRows
        Do While i <= .Rows - 1
            .GetMergedRange i, .FixedCols, lngR1, 0, lngR2, 0
            If Replace(Replace(.Cell(flexcpText, lngR1, .FixedCols + .FrozenCols, lngR2, .Cols - 1), vbTab, ""), vbCr, "") = "" Then
                strMsg = "发现存在尚未设置路径项目的阶段或者分类，要继续吗？"
                Exit Do
            End If
            i = lngR2 + 1
        Loop
        If strMsg <> "" Then
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    
        '标准住院日应与已有阶段的天数匹配
        Set objCombo = cbsMain(cbsMain.count - 1).FindControl(, cmd_Edit_Version, True)
        Set objComboBranch = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Branch, True)
        If objCombo Is Nothing Then
            MsgBox "临床路径表的当前版本信息获取失败。", vbInformation, gstrSysName: Exit Function
        End If
        If objCombo.ListIndex = 0 Then
            MsgBox "临床路径表的当前版本信息获取失败。", vbInformation, gstrSysName: Exit Function
        End If
        If objComboBranch Is Nothing Then
            MsgBox "临床路径表的当前分支信息获取失败。", vbInformation, gstrSysName: Exit Function
        End If
        If objComboBranch.ListIndex = 0 Then
            MsgBox "临床路径表的当前分支信息获取失败。", vbInformation, gstrSysName: Exit Function
        End If
        vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))
        vBranch = mcolBranch("_" & objComboBranch.ItemData(objComboBranch.ListIndex))
        If vVersion.标准住院日 = "" And vBranch.分支名称 = "主路径" Or vBranch.标准住院日 = "" And vBranch.分支名称 <> "主路径" Then
            MsgBox "还没有设置当前" & IIf(vBranch.分支名称 = "主路径", "版本", "分支") & "的标准住院日信息。", vbInformation, gstrSysName: Exit Function
        End If
        
        For i = .Cols - 1 To .FixedCols + .FrozenCols Step -1
            If TypeName(.ColData(i)) <> "Empty" Then
                vStep = .ColData(i)
                If vStep.结束天数 <> 0 Or vStep.开始天数 <> 0 Then
                    If vBranch.分支名称 = "主路径" Then
                        If InStr(vVersion.标准住院日, "-") > 0 Then
                            If vStep.结束天数 <> 0 Then
                                If Val(Split(vVersion.标准住院日, "-")(1)) <> vStep.结束天数 Then
                                    MsgBox "标准住院日的最高天数 " & Val(Split(vVersion.标准住院日, "-")(1)) & " 天与时间阶段已指定的最高天数 " & vStep.结束天数 & " 天不符。", vbInformation, gstrSysName
                                    Exit Function
                                End If
                            ElseIf vStep.开始天数 <> 0 Then
                                If Val(Split(vVersion.标准住院日, "-")(1)) <> vStep.开始天数 Then
                                    MsgBox "标准住院日的最高天数 " & Val(Split(vVersion.标准住院日, "-")(1)) & " 天与时间阶段已指定的最高天数 " & vStep.开始天数 & " 天不符。", vbInformation, gstrSysName
                                    Exit Function
                                End If
                            End If
                        Else
                            If vStep.结束天数 <> 0 Then
                                If Val(vVersion.标准住院日) <> vStep.结束天数 Then
                                    MsgBox "标准住院日 " & vVersion.标准住院日 & " 天与时间阶段已指定的最高天数 " & vStep.结束天数 & " 天不符。", vbInformation, gstrSysName
                                    Exit Function
                                End If
                            ElseIf vStep.开始天数 <> 0 Then
                                If Val(vVersion.标准住院日) <> vStep.开始天数 Then
                                    MsgBox "标准住院日 " & vVersion.标准住院日 & " 天与时间阶段已指定的最高天数 " & vStep.开始天数 & " 天不符。", vbInformation, gstrSysName
                                    Exit Function
                                End If
                            End If
                        End If
                    Else
                        If InStr(vBranch.标准住院日, "-") > 0 Then
                            If vStep.结束天数 <> 0 Then
                                If Val(Split(vBranch.标准住院日, "-")(1)) <> vStep.结束天数 Then
                                    MsgBox "分支路径的标准住院日的最高天数 " & Val(Split(vBranch.标准住院日, "-")(1)) & " 天与时间阶段已指定的最高天数 " & vStep.结束天数 & " 天不符。", vbInformation, gstrSysName
                                    Exit Function
                                End If
                            ElseIf vStep.开始天数 <> 0 Then
                                If Val(Split(vBranch.标准住院日, "-")(1)) <> vStep.开始天数 Then
                                    MsgBox "分支路径的标准住院日的最高天数 " & Val(Split(vBranch.标准住院日, "-")(1)) & " 天与时间阶段已指定的最高天数 " & vStep.开始天数 & " 天不符。", vbInformation, gstrSysName
                                    Exit Function
                                End If
                            End If
                        Else
                            If vStep.结束天数 <> 0 Then
                                If Val(vBranch.标准住院日) <> vStep.结束天数 Then
                                    MsgBox "分支路径的标准住院日 " & vBranch.标准住院日 & " 天与时间阶段已指定的最高天数 " & vStep.结束天数 & " 天不符。", vbInformation, gstrSysName
                                    Exit Function
                                End If
                            ElseIf vStep.开始天数 <> 0 Then
                                If Val(vBranch.标准住院日) <> vStep.开始天数 Then
                                    MsgBox "分支路径的标准住院日 " & vBranch.标准住院日 & " 天与时间阶段已指定的最高天数 " & vStep.开始天数 & " 天不符。", vbInformation, gstrSysName
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                    Exit For
                End If
            End If
        Next
        
        '分支路径的第一个阶段开始天数必须在前一阶段的开始天数和结束天数+1之间
        If vBranch.分支名称 <> "主路径" Then
            For i = .FixedCols + .FrozenCols To .Cols - 1
                If TypeName(.ColData(i)) <> "Empty" Then
                    vStep = .ColData(i)
                    If vBranch.前一阶段ID <> 0 And vBranch.分支ID <> 0 Then
                        On Error GoTo errH
                        strSql = "Select 开始天数,结束天数 From 临床路径阶段 Where ID=[1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, vBranch.前一阶段ID)
                        If rsTmp.RecordCount > 0 Then
                            If Not Between(vStep.开始天数, Val(rsTmp!开始天数 & "") + 1, Val(rsTmp!结束天数 & "") + 1) Then
                                MsgBox "分支路径的第一个阶段的开始天数必须在分支的前一阶段的开始天数后一天和结束天数后一天：" & Val(rsTmp!开始天数 & "") + 1 & "-" & Val(rsTmp!结束天数 & "") + 1 & "之间。", vbInformation, gstrSysName
                                Exit Function
                            End If
                        End If
                        On Error GoTo 0
                    End If
                    Exit For
                End If
            Next
        End If
        '检查阶段中的项目内容重复
        For i = .FixedCols + .FrozenCols To .Cols - 1
            If TypeName(.ColData(i)) <> "Empty" Then
                vStep = .ColData(i)
                strMsg = "": strPathItems = ""
                For j = .FixedRows + .FrozenRows To .Rows - 1
                    If TypeName(.Cell(flexcpData, j, i)) <> "Empty" Then
                        vItem = .Cell(flexcpData, j, i)
                        If InStr(strPathItems & vbTab, vbTab & Trim(vItem.项目内容) & vbTab) = 0 Then
                            strPathItems = strPathItems & vbTab & Trim(vItem.项目内容)
                            .Cell(flexcpFontBold, j, i) = False
                        Else
                            .Cell(flexcpFontBold, j, i) = True
                            strMsg = Trim(vItem.项目内容)
                        End If
                    End If
                Next
                If strMsg <> "" Then
                    '找到第一个
                    For j = .FixedRows + .FrozenRows To .Rows - 1
                        If TypeName(.Cell(flexcpData, j, i)) <> "Empty" Then
                            If .Cell(flexcpData, j, i).项目内容 = strMsg Then
                                .Col = i: .Row = j: .ShowCell .Row, .Col
                                .Cell(flexcpFontBold, j, i) = True
                                Exit For
                            End If
                        End If
                    Next
                    If .FrozenRows > 1 And .TextMatrix(.FixedRows, i) <> .TextMatrix(.FixedRows + .FrozenRows - 1, i) Then
                        strMsg = "阶段""" & Replace(vStep.名称, vbLf, "") & "(" & .TextMatrix(.FixedRows + .FrozenRows - 1, i) & ")""中的路径项目""" & strMsg & """重复，请检查。"
                    Else
                        strMsg = "阶段""" & Replace(vStep.名称, vbLf, "") & """中的路径项目""" & strMsg & """重复，请检查。"
                    End If
                    
                    MsgBox strMsg, vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Next
        
        '清除没有设置项目的多余的分类行
        Call ClearCategoryRow
        
        '设置阶段和项目的序号
        lng阶段序号 = 1
        For i = .FixedCols + .FrozenCols To .Cols - 1
            '阶段序号
            If TypeName(.ColData(i)) <> "Empty" Then
                vStep = .ColData(i)
                '只有第一个阶段允许设置住院日
                If Mid(vStep.标志, 1, 1) = "1" Then
                    If i <> 1 Then
                        If TypeName(.ColData(1)) <> "Empty" Then
                            If vStep.开始天数 <> .ColData(1).开始天数 Or vStep.结束天数 <> .ColData(1).结束天数 Then
                                MsgBox "只有第一个时间阶段才可能设置为住院日，除非与第一个阶段的时间相同。", vbInformation, gstrSysName
                                .Col = i: .Row = .FrozenRows
                                Exit Function
                            End If
                        End If
                    End If
                End If
                '只有最后一个阶段允许设置出院日
                If Mid(vStep.标志, 4, 1) = "1" Then
                    If i <> .Cols - 1 Then
                        If TypeName(.ColData(.Cols - 1)) <> "Empty" Then
                            If vStep.开始天数 <> .ColData(.Cols - 1).开始天数 Or vStep.结束天数 <> .ColData(.Cols - 1).结束天数 Then
                                MsgBox "只有最后一个时间阶段才可能设置为出院日，除非与最后一个阶段的时间相同。", vbInformation, gstrSysName
                                .Col = i: .Row = .FrozenRows
                                Exit Function
                            End If
                        End If
                    End If
                End If
                If vStep.父ID = 0 Then
                    If vStep.序号 <> lng阶段序号 Then
                        vStep.序号 = lng阶段序号
                        If vStep.Edit = 0 Then vStep.Edit = 2 '0-原始,1-新增,2-修改
                    End If
                    lng阶段序号 = lng阶段序号 + 1
                    lng分支序号 = 1
                Else
                    If vStep.序号 <> lng分支序号 Then
                        vStep.序号 = lng分支序号
                        If vStep.Edit = 0 Then vStep.Edit = 2 '0-原始,1-新增,2-修改
                    End If
                    lng分支序号 = lng分支序号 + 1
                End If
                .ColData(i) = vStep
            End If
            
            '项目序号
            If TypeName(.ColData(i)) <> "Empty" Then
                lngR1 = .FixedRows + .FrozenRows
                Do While lngR1 <= .Rows - 1
                    .GetMergedRange lngR1, .FixedCols, lngR1, 0, lngR2, 0
                    
                    lng项目序号 = 1
                    For j = lngR1 To lngR2
                        If TypeName(.Cell(flexcpData, j, i)) <> "Empty" Then
                            vItem = .Cell(flexcpData, j, i)
                            
                            If vItem.项目序号 <> lng项目序号 Then
                                vItem.项目序号 = lng项目序号
                                If vItem.Edit = 0 Then vItem.Edit = 2 '0-原始,1-新增,2-修改
                            End If
                            
                            .Cell(flexcpData, j, i) = vItem
                            lng项目序号 = lng项目序号 + 1
                        End If
                    Next
                    
                    lngR1 = lngR2 + 1
                Loop
            End If
        Next
        
        '清理掉没有使用的医嘱内容ID
        strAdviceIDs = "": mstrChangeItemIDs = ""
        For i = .FixedCols + .FrozenCols To .Cols - 1
            For j = .FixedRows + .FrozenRows To .Rows - 1
                If TypeName(.Cell(flexcpData, j, i)) <> "Empty" Then
                    vItem = .Cell(flexcpData, j, i)
                    If vItem.医嘱IDs <> "" Then
                        strAdviceIDs = strAdviceIDs & "," & vItem.医嘱IDs
                    End If
                    If vItem.待审核医嘱IDs <> "" Then
                        strAdviceIDs = strAdviceIDs & "," & vItem.待审核医嘱IDs
                    End If
                    If (vItem.原医嘱IDs <> vItem.医嘱IDs And vItem.待审核医嘱IDs = "") And vVersion.审核时间 <> Empty And vVersion.停用时间 = Empty Then
                        If gbln双审核 Then
                            mstrChangeItemIDs = mstrChangeItemIDs & "," & vItem.ID & ";" & vItem.审核状态           '记录下变动项目ID
                        Else
                            mstrChangeItemIDs = mstrChangeItemIDs & "," & vItem.ID         '记录下变动项目ID
                        End If
                    End If
                End If
            Next
        Next
        strAdviceIDs = strAdviceIDs & ","
        mstrChangeItemIDs = Mid(mstrChangeItemIDs, 2)
        
        mrsAdvice.Filter = ""
        If Not mrsAdvice.EOF Then
            mrsAdvice.MoveFirst
            Do While Not mrsAdvice.EOF
                If InStr(strAdviceIDs, "," & mrsAdvice!ID & ",") = 0 Then
                    mrsAdvice.Delete
                    mrsAdvice.Update
                End If
                mrsAdvice.MoveNext
            Loop
        End If
    End With
    
    CheckPathTable = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SavePathTable() As Boolean
'功能：保存路径表数据
    Dim vVersion As TYPE_PATH_VERSION
    Dim vBranch As TYPE_PATH_BRANCH
    Dim vStep As TYPE_PATH_STEP
    Dim vItem As TYPE_PATH_ITEM
    Dim vEvalMark As TYPE_PATH_EvalMark
    Dim vEvalCond As TYPE_PATH_EvalCond
    
    Dim objCombo As CommandBarComboBox
    Dim objComboBranch As CommandBarComboBox
    Dim arrSQL As Variant, intVersion As Integer
    Dim i As Long, j As Long, k As Long
    Dim blnTrans As Boolean
    Dim strAddDate As String
    
    Set objCombo = cbsMain(cbsMain.count - 1).FindControl(, cmd_Edit_Version, True)
    If objCombo Is Nothing Then Exit Function
    If objCombo.ListIndex = 0 Then Exit Function
    Set objComboBranch = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Branch, True)
    If objComboBranch Is Nothing Then Exit Function
    If objComboBranch.ListIndex = 0 Then Exit Function
    
    arrSQL = Array()
    vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))
    vBranch = mcolBranch("_" & objComboBranch.ItemData(objComboBranch.ListIndex))
    
    With vsPath
        If mblnNewVersion Then
            '产生新的临床路径版本
            k = 0
            For i = 1 To objCombo.ListCount
                If objCombo.ItemData(i) > k Then k = objCombo.ItemData(i)
            Next
            intVersion = k + 1
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_临床路径版本_Update(" & _
                mlng路径ID & "," & intVersion & ",'" & vVersion.标准住院日 & "','" & vVersion.标准费用 & "','" & vVersion.版本说明 & "')"
            
            '导入评估
            If Not mvEvalImport.指标集 Is Nothing Then
                For i = 1 To mvEvalImport.指标集.count
                    vEvalMark = mvEvalImport.指标集(i)
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_路径评估指标_Insert(" & _
                        mlng路径ID & "," & intVersion & ",NULL,1," & _
                        vEvalMark.ID & "," & vEvalMark.序号 & "," & _
                        "'" & vEvalMark.评估指标 & "'," & vEvalMark.指标类型 & "," & _
                        "'" & vEvalMark.指标结果 & "')"
                Next
            End If
            If Not mvEvalImport.条件集 Is Nothing Then
                For i = 1 To mvEvalImport.条件集.count
                    vEvalCond = mvEvalImport.条件集(i)
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_路径评估条件_Insert(" & _
                        mlng路径ID & "," & intVersion & ",NULL,1," & _
                        ZVal(vEvalCond.指标ID) & ",NULL," & _
                        "'" & vEvalCond.关系式 & "','" & vEvalCond.条件值 & "'," & _
                        vEvalCond.条件组合 & ")"
                Next
            End If
            
            '分支信息
            If vBranch.分支ID <> 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_临床路径分支_Update(" & _
                    vBranch.分支ID & "," & mlng路径ID & "," & vBranch.版本号 & ",'" & vBranch.分支名称 & "'," & vBranch.前一阶段ID & ",'" & _
                    vBranch.标准住院日 & "','" & vBranch.标准费用 & "','" & vBranch.说明 & "')"
            End If
            
            '阶段信息
            For i = .FixedCols + .FrozenCols To .Cols - 1
                If TypeName(.ColData(i)) <> "Empty" Then
                    vStep = .ColData(i)
                    
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_临床路径阶段_Insert(" & _
                        vStep.ID & "," & mlng路径ID & "," & intVersion & "," & _
                        ZVal(vStep.父ID) & "," & vStep.序号 & ",'" & vStep.名称 & "'," & _
                        ZVal(vStep.开始天数) & "," & ZVal(vStep.结束天数) & "," & _
                        "'" & vStep.标志 & "','" & vStep.说明 & "','" & vStep.分类 & "'," & _
                        IIf(vBranch.分支ID <> 0, vBranch.分支ID, "Null") & ")"
                End If
            Next
            
            '分类信息
            k = 1: i = .FixedRows + .FrozenRows
            Do While i <= .Rows - 1
                .GetMergedRange i, .FixedCols, i, 0, j, 0
                If .TextMatrix(i, .FixedCols) <> "" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_临床路径分类_Insert(" & _
                        mlng路径ID & "," & intVersion & "," & k & ",'" & .TextMatrix(i, .FixedCols) & "',Null," & _
                        IIf(vBranch.分支ID <> 0, vBranch.分支ID, "Null") & ")"
                    k = k + 1
                End If
                i = j + 1
            Loop
            
            '项目对应的医嘱内容
            With mrsAdvice
               .Filter = "" '自动MoveFirst,不管Filter变没有
                Do While Not .EOF
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_路径医嘱内容_Insert(" & _
                        !ID & "," & ZVal(Nvl(!相关id, 0)) & "," & !序号 & "," & !期效 & "," & _
                        ZVal(Nvl(!诊疗项目ID, 0)) & ",'" & Nvl(!医嘱内容) & "'," & ZVal(Nvl(!单次用量, 0)) & "," & _
                        ZVal(Nvl(!总给予量, 0)) & "," & ZVal(Nvl(!收费细目ID, 0)) & ",'" & Nvl(!标本部位) & "'," & _
                        "'" & Nvl(!检查方法) & "','" & Nvl(!执行频次) & "'," & ZVal(Nvl(!频率次数, 0)) & "," & _
                        ZVal(Nvl(!频率间隔, 0)) & ",'" & Nvl(!间隔单位) & "','" & Nvl(!医生嘱托) & "'," & _
                        Nvl(!执行性质, 0) & "," & ZVal(Nvl(!执行科室ID, 0)) & ",'" & Nvl(!时间方案) & "',Null,Null," & _
                        !是否缺省 & "," & !是否备选 & "," & IIf(vBranch.分支ID <> 0, vBranch.分支ID, "Null") & _
                       "," & ZVal(Val(!配方ID & "")) & "," & ZVal(Val(!组合项目ID & "")) & "," & Nvl(!执行标记, 0) & ")"
                   
                   .MoveNext
                Loop
            End With
            
            '项目信息
            For i = .FixedCols + .FrozenCols To .Cols - 1
                If TypeName(.ColData(i)) <> "Empty" Then
                    vStep = .ColData(i)
                    For j = .FixedRows + .FrozenRows To .Rows - 1
                        If TypeName(.Cell(flexcpData, j, i)) <> "Empty" Then
                            vItem = .Cell(flexcpData, j, i)
                            
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "Zl_临床路径项目_Insert(" & _
                                vItem.ID & "," & mlng路径ID & "," & intVersion & "," & _
                                vStep.ID & ",'" & .TextMatrix(j, .FixedCols) & "'," & _
                                vItem.项目序号 & ",'" & vItem.项目内容 & "'," & _
                                vItem.执行方式 & "," & vItem.执行者 & "," & _
                                "'" & vItem.项目结果 & "'," & ZVal(vItem.图标ID) & "," & _
                                "'" & vItem.医嘱IDs & "','" & vItem.病历详情 & "'," & vItem.内容要求 & "," & _
                                IIf(vBranch.分支ID <> 0, vBranch.分支ID, "Null") & ",'" & vItem.导入参考 & "'," & _
                                IIf(vItem.导入结果 = 1 And Trim(vItem.导入参考) = "", "Null", vItem.导入结果) & "," & vItem.生成者 & ")"
                        End If
                    Next
                End If
            Next
            
            '阶段评估：和阶段和项目相关，因此在最后
            For i = .FixedCols + .FrozenCols To .Cols - 1
                If TypeName(.ColData(i)) <> "Empty" Then
                    vStep = .ColData(i)
                    If Not vStep.评估.指标集 Is Nothing Then
                        For j = 1 To vStep.评估.指标集.count
                            vEvalMark = vStep.评估.指标集(j)
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "Zl_路径评估指标_Insert(" & _
                                mlng路径ID & "," & intVersion & "," & vStep.ID & ",2," & _
                                vEvalMark.ID & "," & vEvalMark.序号 & "," & _
                                "'" & vEvalMark.评估指标 & "'," & vEvalMark.指标类型 & "," & _
                                "'" & vEvalMark.指标结果 & "'," & IIf(vBranch.分支ID <> 0, vBranch.分支ID, "Null") & ")"
                        Next
                    End If
                    If Not vStep.评估.条件集 Is Nothing Then
                        For j = 1 To vStep.评估.条件集.count
                            vEvalCond = vStep.评估.条件集(j)
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "Zl_路径评估条件_Insert(" & _
                                mlng路径ID & "," & intVersion & "," & vStep.ID & ",2," & _
                                ZVal(vEvalCond.指标ID) & "," & ZVal(vEvalCond.项目ID) & "," & _
                                "'" & vEvalCond.关系式 & "','" & vEvalCond.条件值 & "'," & _
                                vEvalCond.条件组合 & "," & IIf(vBranch.分支ID <> 0, vBranch.分支ID, "Null") & ")"
                        Next
                    End If
                End If
            Next
        Else
            '在原路径版本基础上更新
            intVersion = objCombo.ItemData(objCombo.ListIndex)
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_临床路径版本_Update(" & _
                mlng路径ID & "," & intVersion & ",'" & vVersion.标准住院日 & "','" & vVersion.标准费用 & "','" & vVersion.版本说明 & "')"

            '导入评估
            If Not mvEvalImport.指标集 Is Nothing Then
                For i = 1 To mvEvalImport.指标集.count
                    vEvalMark = mvEvalImport.指标集(i)
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_路径评估指标_Insert(" & _
                        mlng路径ID & "," & intVersion & ",NULL,1," & _
                        vEvalMark.ID & "," & vEvalMark.序号 & "," & _
                        "'" & vEvalMark.评估指标 & "'," & vEvalMark.指标类型 & "," & _
                        "'" & vEvalMark.指标结果 & "')"
                Next
            End If
            If Not mvEvalImport.条件集 Is Nothing Then
                For i = 1 To mvEvalImport.条件集.count
                    vEvalCond = mvEvalImport.条件集(i)
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_路径评估条件_Insert(" & _
                        mlng路径ID & "," & intVersion & ",NULL,1," & _
                        ZVal(vEvalCond.指标ID) & ",NULL," & _
                        "'" & vEvalCond.关系式 & "','" & vEvalCond.条件值 & "'," & _
                        vEvalCond.条件组合 & ")"
                Next
            End If
            
            '分支信息
            If vBranch.分支ID <> 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_临床路径分支_Update(" & _
                    vBranch.分支ID & "," & mlng路径ID & "," & vBranch.版本号 & ",'" & vBranch.分支名称 & "'," & vBranch.前一阶段ID & ",'" & _
                    vBranch.标准住院日 & "','" & vBranch.标准费用 & "','" & vBranch.说明 & "')"
            End If
            
            '阶段信息
            If mstrDelStepIDs <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_临床路径阶段_Delete('" & Mid(mstrDelStepIDs, 2) & "')"
            End If
            
            For i = .FixedCols + .FrozenCols To .Cols - 1
                If TypeName(.ColData(i)) <> "Empty" Then
                    vStep = .ColData(i)
                    If vStep.Edit = 1 Then '新增
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_临床路径阶段_Insert(" & _
                            vStep.ID & "," & mlng路径ID & "," & intVersion & "," & _
                            ZVal(vStep.父ID) & "," & vStep.序号 & ",'" & vStep.名称 & "'," & _
                            ZVal(vStep.开始天数) & "," & ZVal(vStep.结束天数) & "," & _
                            "'" & vStep.标志 & "','" & vStep.说明 & "','" & vStep.分类 & "'," & _
                            IIf(vBranch.分支ID <> 0, vBranch.分支ID, "Null") & ")"
                    ElseIf vStep.Edit = 2 Then '修改
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_临床路径阶段_Update(" & _
                            vStep.ID & "," & mlng路径ID & "," & intVersion & "," & _
                            vStep.序号 & ",'" & vStep.名称 & "'," & _
                            ZVal(vStep.开始天数) & "," & ZVal(vStep.结束天数) & "," & _
                            "'" & vStep.标志 & "','" & vStep.说明 & "','" & vStep.分类 & "'," & _
                            IIf(vBranch.分支ID <> 0, vBranch.分支ID, "Null") & ")"
                    End If
                End If
            Next
            
            '分类信息
            k = 1: i = .FixedRows + .FrozenRows
            Do While i <= .Rows - 1
                .GetMergedRange i, .FixedCols, i, 0, j, 0
                If .TextMatrix(i, .FixedCols) <> "" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_临床路径分类_Insert(" & _
                        mlng路径ID & "," & intVersion & "," & k & ",'" & .TextMatrix(i, .FixedCols) & "'," & _
                        IIf(k = 1, 1, 0) & "," & IIf(vBranch.分支ID <> 0, vBranch.分支ID, "Null") & ")"
                    k = k + 1
                End If
                i = j + 1
            Loop
            
            strAddDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            
            '审核未停用的路径需要插入路径医嘱变动记录(此SQL要先于 Zl_路径医嘱内容_Insert 执行)
            If vVersion.审核时间 <> Empty And vVersion.停用时间 = Empty Then
                If mstrChangeItemIDs <> "" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_路径医嘱变动_Insert('" & mstrChangeItemIDs & "'," & "To_Date('" & strAddDate & "','YYYY-MM-DD HH24:MI:SS')" & ",'" & UserInfo.姓名 & "')"
                End If
            End If
            
            '项目对应的医嘱内容
            With mrsAdvice
                k = 1: .Filter = "" '自动MoveFirst,不管Filter变没有
                Do While Not .EOF
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    If Val(!待审核 & "") = 0 Then
                        arrSQL(UBound(arrSQL)) = "Zl_路径医嘱内容_Insert(" & _
                            !ID & "," & ZVal(Nvl(!相关id, 0)) & "," & !序号 & "," & !期效 & "," & _
                            ZVal(Nvl(!诊疗项目ID, 0)) & ",'" & Nvl(!医嘱内容) & "'," & ZVal(Nvl(!单次用量, 0)) & "," & _
                            ZVal(Nvl(!总给予量, 0)) & "," & ZVal(Nvl(!收费细目ID, 0)) & ",'" & Nvl(!标本部位) & "'," & _
                            "'" & Nvl(!检查方法) & "','" & Nvl(!执行频次) & "'," & ZVal(Nvl(!频率次数, 0)) & "," & _
                            ZVal(Nvl(!频率间隔, 0)) & ",'" & Nvl(!间隔单位) & "','" & Nvl(!医生嘱托) & "'," & _
                            Nvl(!执行性质, 0) & "," & ZVal(Nvl(!执行科室ID, 0)) & ",'" & Nvl(!时间方案) & "'," & _
                            IIf(k = 1, mlng路径ID, "NULL") & "," & IIf(k = 1, intVersion, "NULL") & "," & _
                            !是否缺省 & "," & !是否备选 & "," & IIf(vBranch.分支ID <> 0, vBranch.分支ID, "Null") & "," & ZVal(Val(!配方ID & "")) & _
                            "," & ZVal(Val(!组合项目ID & "")) & "," & Nvl(!执行标记, 0) & ")"
                    Else
                        arrSQL(UBound(arrSQL)) = "Zl_路径医嘱变动_Insert(Null,To_Date('" & strAddDate & "','YYYY-MM-DD HH24:MI:SS')" & ",'" & UserInfo.姓名 & "'," & _
                            !项目ID & "," & !ID & "," & ZVal(Nvl(!相关id, 0)) & "," & !序号 & "," & !期效 & "," & _
                            ZVal(Nvl(!诊疗项目ID, 0)) & "," & ZVal(Nvl(!收费细目ID, 0)) & ",'" & Nvl(!医嘱内容) & "'," & ZVal(Nvl(!单次用量, 0)) & "," & _
                            ZVal(Nvl(!总给予量, 0)) & ",'" & Nvl(!标本部位) & "'," & _
                            "'" & Nvl(!检查方法) & "','" & Nvl(!执行频次) & "'," & ZVal(Nvl(!频率次数, 0)) & "," & _
                            ZVal(Nvl(!频率间隔, 0)) & ",'" & Nvl(!间隔单位) & "','" & Nvl(!医生嘱托) & "'," & _
                            Nvl(!执行性质, 0) & "," & Nvl(!执行标记, 0) & "," & ZVal(Nvl(!执行科室ID, 0)) & ",'" & Nvl(!时间方案) & "'," & _
                            ZVal(Val(!是否缺省 & "")) & "," & ZVal(Val(!是否备选 & "")) & "," & ZVal(Val(!配方ID & "")) & "," & ZVal(Val(!组合项目ID & "")) & _
                            IIf(gbln双审核, "," & !审核状态, "") & ")"
                    End If
                    k = k + 1: .MoveNext
                Loop
            End With
            
            '项目信息
            If mstrDelItemIDs <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_临床路径项目_Delete('" & Mid(mstrDelItemIDs, 2) & "')"
            End If
            For i = .FixedCols + .FrozenCols To .Cols - 1
                If TypeName(.ColData(i)) <> "Empty" Then
                    vStep = .ColData(i)
                    For j = .FixedRows + .FrozenRows To .Rows - 1
                        If TypeName(.Cell(flexcpData, j, i)) <> "Empty" Then
                            vItem = .Cell(flexcpData, j, i)
                            
                            If vItem.Edit = 1 Then '新增
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "Zl_临床路径项目_Insert(" & _
                                    vItem.ID & "," & mlng路径ID & "," & intVersion & "," & _
                                    vStep.ID & ",'" & .TextMatrix(j, .FixedCols) & "'," & _
                                    vItem.项目序号 & ",'" & vItem.项目内容 & "'," & _
                                    vItem.执行方式 & "," & vItem.执行者 & "," & _
                                    "'" & vItem.项目结果 & "'," & ZVal(vItem.图标ID) & "," & _
                                    "'" & vItem.医嘱IDs & "','" & vItem.病历详情 & "'," & vItem.内容要求 & "," & _
                                    IIf(vBranch.分支ID <> 0, vBranch.分支ID, "Null") & ",Null,Null," & vItem.生成者 & ")"
                            ElseIf vItem.Edit = 2 Or (vItem.Edit = 0 And vItem.医嘱IDs <> "") Then '修改，或者强行重新保存医嘱关系
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "Zl_临床路径项目_Update(" & _
                                    vItem.ID & "," & mlng路径ID & "," & intVersion & "," & _
                                    vItem.项目序号 & ",'" & vItem.项目内容 & "'," & _
                                    vItem.执行方式 & "," & vItem.执行者 & "," & _
                                    "'" & vItem.项目结果 & "'," & ZVal(vItem.图标ID) & "," & _
                                    "'" & vItem.医嘱IDs & "','" & vItem.病历详情 & "'," & vItem.内容要求 & ",'" & .TextMatrix(j, .FixedCols) & "'," & _
                                    IIf(vBranch.分支ID <> 0, vBranch.分支ID, "Null") & "," & vItem.生成者 & ")"
                            End If
                        End If
                    Next
                End If
            Next
            
            '阶段评估：和阶段和项目相关，因此在最后
            For i = .FixedCols + .FrozenCols To .Cols - 1
                If TypeName(.ColData(i)) <> "Empty" Then
                    vStep = .ColData(i)
                    If vStep.Edit = 1 Or vStep.Edit = 2 Then '新增或修改
                        If Not vStep.评估.指标集 Is Nothing Then
                            For j = 1 To vStep.评估.指标集.count
                                vEvalMark = vStep.评估.指标集(j)
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "Zl_路径评估指标_Insert(" & _
                                    mlng路径ID & "," & intVersion & "," & vStep.ID & ",2," & _
                                    vEvalMark.ID & "," & vEvalMark.序号 & "," & _
                                    "'" & vEvalMark.评估指标 & "'," & vEvalMark.指标类型 & "," & _
                                    "'" & vEvalMark.指标结果 & "'," & IIf(vBranch.分支ID <> 0, vBranch.分支ID, "Null") & ")"
                            Next
                        End If
                        If Not vStep.评估.条件集 Is Nothing Then
                            For j = 1 To vStep.评估.条件集.count
                                vEvalCond = vStep.评估.条件集(j)
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "Zl_路径评估条件_Insert(" & _
                                    mlng路径ID & "," & intVersion & "," & vStep.ID & ",2," & _
                                    ZVal(vEvalCond.指标ID) & "," & ZVal(vEvalCond.项目ID) & "," & _
                                    "'" & vEvalCond.关系式 & "','" & vEvalCond.条件值 & "'," & _
                                    vEvalCond.条件组合 & "," & IIf(vBranch.分支ID <> 0, vBranch.分支ID, "Null") & ")"
                            Next
                        End If
                    End If
                End If
            Next
        End If
    End With
    
    '执行提交数据
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
    Next
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0
    
    '---
    mstrDelStepIDs = ""
    mstrDelItemIDs = ""
    mstrChangeItemIDs = ""
    mblnChange = False
    mblnNewVersion = False
    
    'List是只读属性只有重新加载
    i = vsPath.Row: j = vsPath.Col
    If vBranch.分支名称 <> "主路径" And vBranch.分支ID <> 0 Then
        Call LoadPathTable(objCombo, objComboBranch, vBranch.分支ID)
    Else
        Call LoadPathVersion(intVersion)
    End If
    If i <= vsPath.Rows - 1 Then vsPath.Row = i
    If j <= vsPath.Cols - 1 Then vsPath.Col = j
    
    Call vsPath.ShowCell(vsPath.Row, vsPath.Col)
    
    SavePathTable = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub FuncPathTableOutput(bytStyle As Byte, Optional ByVal blnIsAll As Boolean, Optional ByVal blnIsMe As Boolean)
'功能：输出临床路径表
'参数：bytStyle=1-打印,2-预览,3-输出到Excel
'      blnIsAll=是否批量输出
'      blnIsMe=全部输出时连续调用
    Dim objCombo As CommandBarComboBox
    Dim vVersion As TYPE_PATH_VERSION
    Dim objComboBranch As CommandBarComboBox
    Dim vBranch As TYPE_PATH_BRANCH
    
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim lngRow As Long, lngCol As Long
    Dim bytR As Byte, strTemp As String
    Dim vItem As TYPE_PATH_ITEM
    Dim lngStart As Long
    
    Set objCombo = cbsMain(cbsMain.count - 1).FindControl(, cmd_Edit_Version, True)
    If objCombo Is Nothing Then Exit Sub
    If objCombo.ListIndex = 0 Then Exit Sub
    
    Set objComboBranch = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Branch, True)
    If objComboBranch Is Nothing Then Exit Sub
    If objComboBranch.ListIndex = 0 Then Exit Sub
    
    If blnIsAll Then
        '只有批量输出时才分别打印分支
        Call LoadPathTable(objCombo, objComboBranch)
        Set objComboBranch = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Branch, True)
        If objComboBranch Is Nothing Then Exit Sub
        If objComboBranch.ListIndex = 0 Then Exit Sub
    End If
    
    vBranch = mcolBranch("_" & objComboBranch.ItemData(objComboBranch.ListIndex))
    vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))
    
    '表头
    objOut.Title.Text = Me.Tag
    If vBranch.分支ID <> 0 Then objOut.Title.Text = objOut.Title.Text & "(" & vBranch.分支名称 & ")"
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表上
    Set objRow = New zlTabAppRow
    If vBranch.分支ID <> 0 Then
        objRow.Add "标准住院日：" & vBranch.标准住院日 & "天"
        objRow.Add "标准费用：" & vBranch.标准费用 & "元"
    Else
        objRow.Add "标准住院日：" & vVersion.标准住院日 & "天"
        objRow.Add "标准费用：" & vVersion.标准费用 & "元"
    End If
    objOut.UnderAppRows.Add objRow
    
    '表下
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名 & vbCrLf & "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    
    '将当前表格复制到输出表格上，并且加载对应的医嘱信息再输出
    With vsPathExport
        .Redraw = flexRDNone
        .Clear
        .Rows = 1: .Rows = vsPath.Rows
        .Cols = (vsPath.Cols - 1) * 2 + 1   '除固定列外，其余每列后增加一列用于显示对应的医嘱。
        .FixedRows = 0: .FixedCols = 0
        .Width = vsPath.Width
        .Height = vsPath.Height
        
        .Redraw = flexRDDirect
        
        .Redraw = flexRDNone
        '第一行显示版本信息
        .TextMatrix(0, 1) = "当前版本：第" & vVersion.版本号 & "版" & IIf(vVersion.版本说明 <> "", "：" & vVersion.版本说明, "") & _
                            vbCrLf & IIf(vVersion.审核人 = "", "创建时间：" & Format(vVersion.创建时间, "yyyy年MM月dd日") & vbCrLf & "创建人：" & vVersion.创建人 & "(未审核)", _
                            "审核时间：" & Format(vVersion.审核时间, "yyyy年MM月dd日") & vbCrLf & "审核人：" & vVersion.审核人)
        If vBranch.分支ID <> 0 Then .TextMatrix(0, 1) = .TextMatrix(0, 1) & vbCrLf & "分支信息：" & objComboBranch.List(objComboBranch.ListIndex)
        .TextMatrix(0, 2) = "适用科室：" & mstrDeptInfo & vbCrLf & "适用病种：" & mstrDiagInfo
        
        '第二行（第三行）路径阶段信息
        '从第1行开始，第0行为固定行，不复制
        If Trim(vsPath.TextMatrix(2, 0)) = "时间阶段" Then
            lngStart = 2
        Else
            lngStart = 1
        End If
        For lngRow = 1 To lngStart
            For lngCol = 1 To vsPath.Cols - 1
                If vsPath.TextMatrix(lngRow, lngCol) <> vsPath.TextMatrix(lngRow - 1, lngCol) Then
                    .TextMatrix(lngRow, lngCol * 2 - 1) = Replace(Replace(vsPath.TextMatrix(lngRow, lngCol), vbLf, ""), vbCr, "")
                End If
            Next
        Next
        
        '其余行，路径项目
        For lngCol = 0 To vsPath.Cols - 1
            '医嘱列
            .ColAlignment(lngCol * 2) = vsPath.ColAlignment(lngCol)
            .ColWidth(lngCol * 2) = vsPath.ColWidth(lngCol) * 1.6
                        
            If lngCol = 0 Then
                '项目类别
                For lngRow = lngStart + 1 To vsPath.Rows - 1
                    If vsPath.TextMatrix(lngRow, 0) <> vsPath.TextMatrix(lngRow - 1, 0) Then .TextMatrix(lngRow, 0) = vsPath.TextMatrix(lngRow, 0)
                Next
            Else
                .ColAlignment(lngCol * 2 - 1) = vsPath.ColAlignment(lngCol)
                .ColWidth(lngCol * 2 - 1) = vsPath.ColWidth(lngCol)
                
                
                '当前列的所有路径项目行
                .Cell(flexcpText, lngStart + 1, lngCol * 2 - 1, .Rows - 1, lngCol * 2 - 1) = vsPath.Cell(flexcpText, lngStart + 1, lngCol, .Rows - 1, lngCol)
                                
                For lngRow = lngStart + 1 To vsPath.Rows - 1
                    
                    If TypeName(vsPath.Cell(flexcpData, lngRow, lngCol)) <> "Empty" Then
                        vItem = vsPath.Cell(flexcpData, lngRow, lngCol)
                        strTemp = vItem.Tip '医嘱内容或病历名称摘要
                        If InStr(strTemp, ":") > 0 Then
                            strTemp = Trim(Mid(strTemp, InStr(strTemp, ":") + 1))
                        Else
                            If vItem.医嘱IDs <> "" Then
                                strTemp = GetAdviceDefineText(vItem.医嘱IDs, mrsAdvice)
                            ElseIf vItem.病历IDs <> "" Or vItem.新版病历IDs <> "" Then
                                If vItem.病历IDs <> "" And vItem.新版病历IDs <> "" Then
                                    strTemp = GetEPRDefineText(, vItem.ID)
                                ElseIf vItem.病历IDs <> "" Then
                                    strTemp = GetEPRDefineText(vItem.病历IDs)
                                Else
                                    strTemp = GetEPRDefineText(vItem.新版病历IDs, vItem.ID)
                                End If
                            End If
                        End If
                        
                        strTemp = Replace(strTemp, "○", "")
                        .TextMatrix(lngRow, lngCol * 2) = strTemp
                    End If
                Next
            End If
        Next
        .Redraw = flexRDDirect
    End With
    
    '表体
    Set objOut.Body = vsPathExport
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    '清除内容以释放内存
    vsPathExport.Clear
    vsPathExport.Rows = 1: vsPathExport.Cols = 1
    
    '如果是全部输出，则循环调用，直到最后一个
    If blnIsAll Or blnIsMe Then
        If objComboBranch.ListIndex < objComboBranch.ListCount Then
            Call LoadPathTable(objCombo, objComboBranch, objComboBranch.ItemData(objComboBranch.ListIndex + 1))
            Call FuncPathTableOutput(bytStyle, False, True)
        End If
    End If
End Sub

Private Sub FuncExportToXML()
'功能：导出成XML文件
    Dim objCombo As CommandBarComboBox
    Dim vVersion As TYPE_PATH_VERSION
    
    If mbytMode = Mode_Design And mblnChange Then
        MsgBox "路径表内容变更后尚未保存，请先保存。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Set objCombo = cbsMain(cbsMain.count - 1).FindControl(, cmd_Edit_Version, True)
    If objCombo Is Nothing Then Exit Sub
    If objCombo.ListIndex = 0 Then Exit Sub
    vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))
    
    '导出目录选择
    cdgXML.DialogTitle = "导出临床路径"
    cdgXML.Filter = "XML文件|*.xml"
    cdgXML.Flags = &H200000 Or &H4 Or &H2 Or &H800 Or &H4000
    cdgXML.InitDir = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "临床路径XML目录")
    cdgXML.FileName = Replace(Me.Tag, vbCrLf, "_") & ".xml"
    cdgXML.CancelError = True
    On Error Resume Next
    cdgXML.ShowSave
    If Err.Number <> 0 Then
        '不是取消时
        If Err.Number <> 32755 Then MsgBox "导出过程发生错误:" & Err.Description, vbInformation, gstrSysName
        Err.Clear: Exit Sub
    End If
    On Error GoTo 0
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "临床路径XML目录", gobjFile.GetParentFolderName(cdgXML.FileName)
    
    '导出
    Screen.MousePointer = 11
    Call ExportPathToXML(mlng路径ID, vVersion.版本号, cdgXML.FileName)
    Screen.MousePointer = 0
End Sub

Private Sub FuncPathImportFromXML()
    Dim objCombo As CommandBarComboBox
    Dim intVersion As Integer, k As Long, i As Long
    
    Set objCombo = cbsMain(cbsMain.count - 1).FindControl(, cmd_Edit_Version, True)
    If objCombo Is Nothing Then Exit Sub
    If objCombo.ListIndex = 0 Then Exit Sub
    
    cdgXML.DialogTitle = "导入临床路径"
    cdgXML.Filter = "XML文件|*.xml"
    cdgXML.Flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
    cdgXML.InitDir = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "临床路径XML目录")
    cdgXML.CancelError = True
    On Error Resume Next
    cdgXML.ShowOpen
    If Err.Number <> 0 Then
        Err.Clear: Exit Sub
    End If
    On Error GoTo 0
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "临床路径XML目录", gobjFile.GetParentFolderName(cdgXML.FileName)
    
    '确定导入版本号
    If mblnNewVersion Then
        k = 0
        For i = 1 To objCombo.ListCount
            If objCombo.ItemData(i) > k Then k = objCombo.ItemData(i)
        Next
        intVersion = k + 1
    Else
        intVersion = objCombo.ItemData(objCombo.ListIndex)
    End If
    
    '导入路径
    Screen.MousePointer = 11
    If ImportPathFromXML(cdgXML.FileName, mlng路径ID, intVersion) Then
        mstrDelStepIDs = ""
        mstrDelItemIDs = ""
        mblnChange = False
        mblnNewVersion = False
        Call LoadPathVersion(intVersion)
    End If
    Screen.MousePointer = 0
End Sub

Private Function GetParentStep(vStep As TYPE_PATH_STEP) As TYPE_PATH_STEP
'功能：获取分支阶段的父阶段
    Dim i As Long
    
    With vsPath
        For i = .FixedCols + .FrozenCols To .Cols - 1
            If TypeName(.ColData(i)) <> "Empty" Then
                If .ColData(i).ID = vStep.父ID Then
                    GetParentStep = .ColData(i)
                    Exit For
                End If
            End If
        Next
    End With

End Function

Private Sub FuncFindItem(Optional ByVal blnNext As Boolean)
'参数：blnNext=是否查找下一个
    Dim blnHave As Boolean, i As Long, j As Long
    Dim vStep As TYPE_PATH_STEP
    Dim vItem As TYPE_PATH_ITEM
    Dim lngRow As Long, lngCol As Long
    Dim blnOver As Boolean
    
    
    If Trim(txtFind.Text) = "" Then Exit Sub
    Call zlControl.TxtSelAll(txtFind)
            
    '开始查找行
    With vsPath
        If .Row < .FixedRows + .FrozenRows Or .Col < .FixedCols + .FrozenCols Then .Row = .FixedRows + .FrozenRows: .Col = .FixedCols + .FrozenCols
        
        If blnNext Then
            If .Row = .Rows - 1 And .Col = .Cols - 1 Then
                blnOver = True
            Else
                lngRow = .Row: lngCol = .Col
                If .Row = .Rows - 1 Then
                    lngRow = .FixedRows + .FrozenRows
                    lngCol = .Col + 1
                Else
                    lngRow = .Row + 1
                End If
            End If
        Else
            lngCol = .FixedCols + .FrozenCols: lngRow = .FixedRows + .FrozenRows
        End If
        '从当前行开始往后找(从左到右，从上至下）
        If Not blnOver Then
            For i = lngCol To .Cols - 1
                If TypeName(.ColData(i)) <> "Empty" Then
                    vStep = .ColData(i)
                    For j = .FixedRows + .FrozenRows To .Rows - 1
                        If i <> lngCol Or j >= lngRow Then
                            If TypeName(.Cell(flexcpData, j, i)) <> "Empty" Then
                                vItem = .Cell(flexcpData, j, i)
                                If vItem.项目内容 Like IIf(gstrLike <> "", "*", "") & txtFind.Text & "*" Then
                                    blnHave = True
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                    If blnHave Then Exit For
                End If
            Next
        End If
    
        If blnHave And Not blnOver Then
            .Row = j: .Col = i
            .ShowCell .Row, .Col
            If .Visible Then .SetFocus
        Else
            MsgBox IIf(blnNext, "后面已", "") & "找不到符合条件的路径项目。", vbInformation, gstrSysName
        End If
    End With

End Sub

Private Sub ShowContrast(ByVal bytMode As Byte)
'功能:1.以不同背景色区别显示医嘱内容与上一版本有差异的项目
'     用蓝色背景色显示（&H00FFEADA&）区分,
'参数:bytMode   1-显示差异，2-隐藏差异

    Dim rsNew As ADODB.Recordset, rsOld As ADODB.Recordset
    Dim rsAdviceNew As ADODB.Recordset, rsAdviceOld As ADODB.Recordset

    Dim strSql As String
    Dim strSqlNew As String, strSqlOld As String
    Dim objComboBranch As CommandBarComboBox
    Dim objCombo As CommandBarComboBox

    Dim vBranch As TYPE_PATH_BRANCH
    Dim vItem As TYPE_PATH_ITEM
    Dim strTmp As String

    Dim lngRow As Long, lngCol As Long
    Dim lngVersion As Long, lngBranchID As Long
    Dim i As Long, j As Long, lngCount As Long
    Dim intOldItemId As Long
    Dim blnDo As Boolean

    On Error GoTo errH
    blnDo = False
    If bytMode = 1 Then
        Set mcolItemID = New Collection
        Set objCombo = cbsMain(cbsMain.count - 1).FindControl(, cmd_Edit_Version, True)
        lngVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex)).版本号    '当前版本号
        If lngVersion < 2 Then
            Exit Sub
        End If
        Set objComboBranch = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Branch, True)
        vBranch = mcolBranch("_" & objComboBranch.ItemData(objComboBranch.ListIndex))

        If vBranch.分支名称 <> "主路径" Then    '分支路径
            strSql = "Select a.Id From 临床路径分支 A Where a.路径id = [1] And a.版本号 = [2]"

            Set rsOld = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng路径ID, lngVersion - 1)
            If rsOld.RecordCount > 0 Then
                lngBranchID = rsOld!ID
            Else
                MsgBox "上一个版本不存在分支路径", vbInformation, gstrSysName
                mblnDiff = False
                Exit Sub
            End If
        End If
        '属于医嘱类的项目名称 添加临床路径分类 为了按分类序号排序后，便于按照从上倒下，从左到右的顺序添加mcolItemID
        strSql = "Select a.Id As 阶段id, a.序号, Nvl(b.序号, 0) As 父id序号, a.名称, a.开始天数, Nvl(a.结束天数, 0) As 结束天数, c.分类, c.Id As 项目id, c.项目内容" & vbNewLine & _
                 "From 临床路径阶段 A, 临床路径阶段 B, 临床路径项目 C, 临床路径分类 D" & vbNewLine & _
                 "Where a.路径id = [1] And a.版本号 = [2] " & _
                 IIf(vBranch.分支名称 = "主路径" Or vBranch.版本号 = 0, " And a.分支ID is null", " And A.分支ID=[3]") & _
                 "  And a.父id = b.Id(+) And a.Id = c.阶段id And d.路径id = c.路径id And" & vbNewLine & _
                 "      d.版本号 = c.版本号 And Nvl(d.分支id, 0) = Nvl(c.分支id, 0) And d.名称 = c.分类 And Exists" & vbNewLine & _
                 " (Select 1 From 临床路径医嘱 D Where c.Id = d.路径项目id)" & vbNewLine & _
                 "Order By Nvl(b.序号, a.序号), Nvl(b.序号, 0), a.序号, d.序号, c.项目序号"
        Set rsNew = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng路径ID, lngVersion, vBranch.分支ID)      '新版本
        '旧版所有项目名称
        strSql = "Select a.Id As 阶段id, a.序号, Nvl(b.序号, 0) As 父id序号, a.名称, a.开始天数, Nvl(a.结束天数, 0) As 结束天数, c.分类, c.Id As 项目id, c.项目内容" & vbNewLine & _
                 "From 临床路径阶段 A, 临床路径阶段 B, 临床路径项目 C,临床路径分类 D" & vbNewLine & _
                 "Where a.路径id = [1] And a.版本号 = [2] " & _
                 IIf(vBranch.分支名称 = "主路径" Or vBranch.版本号 = 0, " And a.分支ID is null", " And A.分支ID=[3]") & _
                 " And a.父id = b.Id(+) And a.Id = c.阶段id  And d.路径id = c.路径id And" & vbNewLine & _
                 "     d.版本号 = c.版本号 And Nvl(d.分支id, 0) = Nvl(c.分支id, 0) And d.名称 = c.分类 " & vbNewLine & _
                 "Order By Nvl(b.序号, a.序号), Nvl(b.序号, 0), Nvl(a.序号, 0), d.序号, c.项目序号"
        Set rsOld = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng路径ID, lngVersion - 1, lngBranchID)       '旧版本

        Do While Not rsNew.EOF
            rsOld.Filter = "序号 =" & Val(Nvl(rsNew!序号)) & " And 父id序号 = " & Val(Nvl(rsNew!父id序号)) & " And 开始天数 =" & Val(Nvl(rsNew!开始天数)) & _
                           " And 结束天数= " & Val(Nvl(rsNew!结束天数)) & " And 分类 ='" & Nvl(rsNew!分类) & "' And 项目内容 = '" & Nvl(rsNew!项目内容) & "'"

            If rsOld.RecordCount > 0 Then
                '同阶段，同分类，同项目
                'strSql语句中将部分列名做空值转换，为了便于用Filter做过滤
                strSql = "Select b.序号, b.期效, Nvl(b.诊疗项目id,0) as 诊疗项目ID, Nvl(b.收费细目id, 0) as 收费细目id, Nvl(b.医嘱内容, 0) As 医嘱内容, Nvl(b.单次用量, 0) As 单次用量," & vbNewLine & _
                         "       Nvl(b.总给予量, 0) As 总给予量, Nvl(b.执行频次,0) as 执行频次, b.执行性质, Nvl(b.检查方法, 0) As 检查方法, Nvl(b.标本部位, 0) As 标本部位," & vbNewLine & _
                         "       Nvl(b.执行科室id, 0) As 执行科室id, Nvl(b.时间方案, 0)  as 时间方案" & vbNewLine & _
                         "From 临床路径医嘱 A, 路径医嘱内容 B" & vbNewLine & _
                         "Where a.路径项目id = [1] And a.医嘱内容id = b.Id" & vbNewLine & _
                         "Order By b.序号"

                Set rsAdviceNew = zlDatabase.OpenSQLRecord(strSql, Me.Caption, rsNew!项目ID)
                Set rsAdviceOld = zlDatabase.OpenSQLRecord(strSql, Me.Caption, rsOld!项目ID)

                If rsAdviceNew.RecordCount > 0 And rsAdviceOld.RecordCount = 0 Then

                    '第一种,新版是医嘱项目，旧版不是医嘱项目
                    intOldItemId = rsOld!项目ID
                    blnDo = True
                ElseIf rsAdviceNew.RecordCount > 0 And rsAdviceOld.RecordCount > 0 Then

                    '第二种，新版旧版都是医嘱项目
                    For i = 1 To rsAdviceNew.RecordCount

                        rsAdviceOld.Filter = "期效 = " & Val(Nvl(rsAdviceNew!期效)) & " And 诊疗项目ID = " & Val(Nvl(rsAdviceNew!诊疗项目ID)) & " and 收费细目ID=" & Val(Nvl(rsAdviceNew!收费细目ID)) & _
                                             " And 医嘱内容 ='" & Nvl(rsAdviceNew!医嘱内容) & "' And 单次用量 =" & Val(Nvl(rsAdviceNew!单次用量)) & " And 总给予量 = " & Val(Nvl(rsAdviceNew!总给予量)) & _
                                             " And 执行频次 = '" & Nvl(rsAdviceNew!执行频次) & "' And 执行性质 ='" & Nvl(rsAdviceNew!执行性质) & "' And 检查方法 = '" & Nvl(rsAdviceNew!检查方法) & "'" & _
                                             " And 标本部位 = '" & Nvl(rsAdviceNew!标本部位) & "' And 执行科室ID =" & Val(Nvl(rsAdviceNew!执行科室ID)) & " And 时间方案 = '" & Nvl(rsAdviceNew!时间方案) & "'"
                        '一旦有一条医嘱不相同,就退出循环
                        If rsAdviceOld.RecordCount = 0 Then
                            intOldItemId = rsOld!项目ID
                            blnDo = True
                            Exit For
                        End If
                        rsAdviceNew.MoveNext
                    Next
                End If
            ElseIf rsOld.RecordCount < 1 Then    '不存在相同的阶段或相同分类或相同项目时
                intOldItemId = 0
                blnDo = True
            End If
            If blnDo Then

                lngCount = lngCount + 1
                '记录下存在差异的项目ID,方便对比查看时,上一个或下一个提取项目ID
                'item： 新版项目ID:老版项目ID:下标位置
                mcolItemID.Add Val(rsNew!项目ID) & ":" & intOldItemId & ":" & lngCount, "_" & Val(rsNew!项目ID)
                strTmp = mcolItemRowCol("_" & rsNew!项目ID)
                lngRow = Split(strTmp, ",")(0)
                lngCol = Split(strTmp, ",")(1)
                With vsPath
                    vItem = .Cell(flexcpData, lngRow, lngCol)
                    vItem.前一版本项目ID = intOldItemId
                    .Cell(flexcpData, lngRow, lngCol) = vItem
                End With
                blnDo = False
            End If

            rsNew.MoveNext
        Loop

        If mcolItemID.count = 0 Then
            mblnDiff = False
            MsgBox "该版本医嘱类项目同上一版本相同。", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
    End If

    '设置差异颜色/隐藏差异
    For i = 1 To mcolItemID.count
        strTmp = mcolItemRowCol("_" & Split(mcolItemID(i), ":")(0))
        lngRow = Split(strTmp, ",")(0)
        lngCol = Split(strTmp, ",")(1)
        vsPath.Cell(flexcpBackColor, lngRow, lngCol) = IIf(bytMode = 1, Color_DiffBack, Empty)
    Next

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CompareAdviceItem()
    Dim vItem As TYPE_PATH_ITEM
    Dim i As Long
    
    '对比查看
    With vsPath
        If .TextMatrix(.Row, .Col) <> "" Then
            vItem = .Cell(flexcpData, .Row, .Col)
            If .Cell(flexcpBackColor, .Row, .Col) = Color_DiffBack Then
                mfrmAdviceContrast.ShowMe Me, vItem.ID, vItem.前一版本项目ID, mcolItemID
            Else
                MsgBox "请选择一个蓝色背景的单元格再执行对比查看。", vbOKOnly + vbInformation, gstrSysName
            End If
        Else
            MsgBox "你当前选择的单元格没有定义路径项目，请选择一个蓝色背景的单元格。", vbOKOnly + vbInformation, gstrSysName
        End If
    End With
End Sub

Private Sub FuncResizeCenter()
    Dim objControl As CommandBarControl
    
    On Error Resume Next
    
    If mbytFunc = 0 Then
        Me.picBottom.Visible = False
        Me.fraSplit.Visible = False
        vsPath.Move 0, 0, picCenter.Width, picCenter.Height
    ElseIf mbytFunc = 1 Or mbytFunc = 2 Then
        Me.picBottom.Visible = True
        Me.fraSplit.Visible = True
        ucAdvice(0).Set选择列的可见性 (True)
        ucAdvice(1).Set选择列的可见性 (True)
        vsPath.Move 0, 0, picCenter.Width, picCenter.Height / 10 * 7
        fraSplit.Move 0, picCenter.Height / 10 * 7, picCenter.Width, 45
        picBottom.Move 0, fraSplit.Top + 45, picCenter.Width, picCenter.Height - fraSplit.Top - 50
        Call FuncResizeBottom
    End If
End Sub

Private Sub FuncResizeBottom()
'功能:重新调整变动记录位置
    On Error Resume Next
    
    lblCurr.Move 120, 50, 1095, 300
    ucAdvice(0).Move 120, 360, picBottom.Width / 2 - 120, picBottom.Height - 300
    fraSplit2.Move picBottom.ScaleWidth / 2, 400, 60, picBottom.Height
    lblChange.Move fraSplit2.Left + 120, 50, 1095, 300
    With cboTimes
        .Left = fraSplit2.Left + 60 + lblChange.Width + 120: .Top = 15
        .Width = IIf(mbytFunc = 1, 8000, 5000)
        .Height = 300
    End With
    If mbytFunc = 2 Then
        cmdCheck(0).Visible = True
        cmdCheck(1).Visible = True
        cmdCheck(0).Move cboTimes.Left + cboTimes.Width + 500, cboTimes.Top, 1100, 360
        cmdCheck(1).Move cmdCheck(0).Left + cmdCheck(0).Width + 120, cmdCheck(0).Top, 1100, 360
        Call FuncSetAuditBtn
    Else
        cmdCheck(0).Visible = False
        cmdCheck(1).Visible = False
    End If
    ucAdvice(1).Move fraSplit2.Left + 60, 360, picBottom.Width - fraSplit2.Left - 60 - 120, picBottom.Height - 300
End Sub

Private Sub FuncShowAdvice(Optional ByVal bytModel As Byte = 0)
'功能:显示变动记录
'参数:
'   bytModel =0显示当前医嘱详情
'            =1显示指定的路径医嘱变动记录
'            =2清空医嘱记录

    Dim lng项目ID As Long
    Dim vItem As TYPE_PATH_ITEM
    Dim strSQLOne As String
    Dim strSQLTwo As String
    
    On Error GoTo errH
    If vsPath.Row < 0 Or vsPath.Col < 0 Then Exit Sub
    
    If TypeName(vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col)) <> "Empty" And InStr(",0,1,", "," & bytModel & ",") > 0 Then
        vItem = vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col)
        If bytModel = 0 Then
            strSQLOne = "Select a.Id, a.相关id, a.序号, a.期效, a.诊疗项目id, a.收费细目id, a.医嘱内容, a.单次用量, a.总给予量, a.标本部位, a.检查方法, a.医生嘱托, a.执行频次, a.频率次数," & vbNewLine & _
                 "       a.频率间隔, a.间隔单位, a.执行性质, a.执行科室id, a.时间方案, a.是否缺省, a.是否备选, a.配方id, a.组合项目id,a.执行标记 " & vbNewLine & _
                 "From 路径医嘱内容 A, 临床路径医嘱 B" & vbNewLine & _
                 "Where a.Id = b.医嘱内容id And b.路径项目id =[3] "
            ucAdvice(0).ShowAdvice 0, strSQLOne, , , True, vItem.ID
        ElseIf bytModel = 1 Then
            strSQLTwo = "Select a.医嘱内容ID as Id, a.相关id, a.序号, a.期效, a.诊疗项目id, a.收费细目id, a.医嘱内容, a.单次用量, a.总给予量, a.标本部位, a.检查方法, a.医生嘱托, a.执行频次, a.频率次数," & vbNewLine & _
                "       a.频率间隔, a.间隔单位, a.执行性质, a.执行科室id, a.时间方案, a.是否缺省, a.是否备选, a.配方id, a.组合项目id,a.执行标记 " & vbNewLine & _
                "From 路径医嘱变动 A " & vbNewLine & _
                "Where a.项目Id = [3] and a.操作时间= To_Date('" & Format(cboTimes.Tag, "yyyy-mm-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
            ucAdvice(1).ShowAdvice 0, strSQLTwo, , , True, vItem.ID
        End If
    Else
        ucAdvice(0).ShowAdvice 0, "", , , True
        ucAdvice(1).ShowAdvice 0, "", , , True
        cboTimes.Clear
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncSetItemBackColor()
'功能:查找存在路径医嘱变动的路径项目,并将不存在医嘱变动的路径项目背景设置为灰色
    Dim i As Long
    Dim j As Long
    Dim vVersion As TYPE_PATH_VERSION
    Dim vItem As TYPE_PATH_ITEM
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim strIDs As String
    Dim objCombo As CommandBarComboBox
    
    Set objCombo = cbsMain(cbsMain.count - 1).FindControl(, cmd_Edit_Version, True)
    vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))
    On Error GoTo errH
    If mbytFunc = 1 Then
        strSql = "Select Distinct b.项目id From 临床路径项目 A, 路径医嘱变动 B Where a.路径id = [1] And a.版本号 = [2] And a.Id = b.项目id And B.审核状态 In (1,0) "
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng路径ID, vVersion.版本号)
        If rsTmp.RecordCount < 1 Then
            MsgBox "该临床路径不存在医嘱变动记录。", vbOKOnly + vbInformation, gstrSysName
        End If
        For i = 1 To rsTmp.RecordCount
            strIDs = strIDs & "," & rsTmp!项目ID
            rsTmp.MoveNext
        Next
        strIDs = strIDs & ","
    End If
    With vsPath
        For i = .FixedCols To .Cols - 1
            For j = 1 To .Rows - 1
                If TypeName(.Cell(flexcpData, j, i)) <> "Empty" Then
                    If mbytFunc = 1 Then
                        vItem = .Cell(flexcpData, j, i)
                        If InStr(strIDs, "," & vItem.ID & ",") > 0 Then
                            .Cell(flexcpBackColor, j, i) = Color_DiffBack
                        End If
                    Else
                        If .Cell(flexcpBackColor, j, i) = Color_DiffBack Then
                            .Cell(flexcpBackColor, j, i) = 0
                        End If
                    End If
                End If
            Next
        Next
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Public Sub FuncLoadChangeTimes()
'功能:医嘱变动详情变动次数加载
'参数:mbytFunc=1 医嘱变动历史记录(审核人\审核时间 不为空
'     mbytFunc=2 路径项目变动待审核的记录 医嘱变动记录(审核人=NULL)的记录
    Dim strSql As String, strWhere As String
    Dim rsTmp As ADODB.Recordset
    Dim vItem As TYPE_PATH_ITEM
    Dim i As Long
    Dim strTip As String
    
    On Error GoTo errH
    
    cboTimes.Clear: cboTimes.Tag = ""
    
    If TypeName(vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col)) = "Empty" Then Exit Sub
    marrTime = Array()
    If mbytFunc = 1 Then
        strSql = "Select Rownum As 序号, a.*" & vbNewLine & _
                    "From (Select Distinct a.操作时间, a.操作员,a.审核状态, a.审核人, a.审核时间, a.药剂审核人, a.药剂审核时间 " & vbNewLine & _
                    "       From 路径医嘱变动 A" & vbNewLine & _
                    "       Where a.项目id = [1] And a.审核状态 In (0,1)" & vbNewLine & _
                    "       Order By a.操作时间) A" & vbNewLine & _
                    "Order By Rownum Desc"
    ElseIf mbytFunc = 2 Then
        strSql = "Select Distinct a.操作时间, a.操作员, a.审核人, a.审核时间" & vbNewLine & _
                "From 路径医嘱变动 A" & vbNewLine & _
                "Where a.项目id = [1] And NVL(a.审核状态,-1) Not In (0,1) " & vbNewLine & _
                "Order By a.操作时间 Desc"
    End If
    vItem = vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col)
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, vItem.ID)
    If mbytFunc = 1 Then
        For i = 1 To rsTmp.RecordCount
            If gbln双审核 Then
                strTip = " 药剂审核:" & rsTmp!药剂审核人 & "/" & Format(rsTmp!药剂审核时间, "yyyy-mm-dd hh:mm:ss") & " 审核:" & rsTmp!审核人 & "/" & Format(rsTmp!审核时间, "yyyy-mm-dd hh:mm:ss")
                strTip = Replace(strTip, " 药剂审核:/", "")
                strTip = Replace(strTip, " 审核:/", "")
            Else
                strTip = " 审核:" & rsTmp!审核人 & "/" & rsTmp!审核时间
            End If
            
            cboTimes.AddItem "第" & rsTmp!序号 & "次,登记:" & rsTmp!操作员 & "/" & Format(rsTmp!操作时间, "yyyy-mm-dd hh:mm:ss") & strTip & Space(1) & IIf(Val(rsTmp!审核状态 & "") = 0, "审核未通过", "审核通过")
            ReDim Preserve marrTime(UBound(marrTime) + 1)
            marrTime(UBound(marrTime)) = Format(rsTmp!操作时间, "yyyy-mm-dd hh:mm:ss")
            rsTmp.MoveNext
        Next
    ElseIf mbytFunc = 2 Then
        For i = 1 To rsTmp.RecordCount
            cboTimes.AddItem "登记:" & rsTmp!操作员 & "/" & Format(rsTmp!操作时间, "yyyy-mm-dd hh:mm:ss") & Space(1) & "待审核"
            ReDim Preserve marrTime(UBound(marrTime) + 1)
            marrTime(UBound(marrTime)) = Format(rsTmp!操作时间, "yyyy-mm-dd hh:mm:ss")
            rsTmp.MoveNext
        Next
    End If
    
    If cboTimes.ListCount > 0 Then
        cboTimes.ListIndex = 0   '缺省定位到最新变动记录\最新待审核记录
    Else
        ucAdvice(1).ShowAdvice 0, "", 0, 0, True '清空数据
    End If
        
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncShowItemBySendor()
'功能：根据选择生成者的不同显示路径项目
    Dim i As Long
    Dim j As Long
    Dim vItem As TYPE_PATH_ITEM
    
    With vsPath
        For j = .FixedCols To .Cols - 1
            For i = .FixedRows To .Rows - 1
                If TypeName(.Cell(flexcpData, i, j)) <> "Empty" Then
                    vItem = .Cell(flexcpData, i, j)
                    If Val(optSelect(IX_ALL).Tag) = vItem.生成者 Then
                        .Cell(flexcpBackColor, i, j) = Color_DiffBack
                    Else
                        .Cell(flexcpBackColor, i, j) = vbWhite
                    End If
                End If
            Next
        Next
    
    End With
End Sub

Private Sub FuncSetAuditBtn()
'功能：根据选择生成者的不同显示路径项目
    Dim vItem As TYPE_PATH_ITEM
    Dim blnVisible As Boolean
    Dim vArea As CONST_AREA
    
    If mbytFunc <> 2 Then Exit Sub
    With vsPath
        If gbln双审核 Then
            vArea = GetArea(.Row, .Col)
            If vArea = Area_Item Then
               If TypeName(.Cell(flexcpData, .Row, .Col)) <> "Empty" Then
                   vItem = .Cell(flexcpData, .Row, .Col)
               Else
                   vItem.审核状态 = 0
               End If
            
               If mbytAudit = 1 Then
                   blnVisible = InStr(",1,3,", "," & vItem.审核状态 & ",") > 0
               ElseIf mbytAudit = 2 Then
                   blnVisible = InStr(",2,3,", "," & vItem.审核状态 & ",") > 0
               ElseIf mbytAudit = 3 Then
                   blnVisible = InStr(",1,2,3,", "," & vItem.审核状态 & ",") > 0
               End If
            Else
                blnVisible = False
            End If
        Else
            blnVisible = True
        End If
        cmdCheck(0).Visible = blnVisible
        cmdCheck(1).Visible = blnVisible
        If blnVisible Then
            cmdCheck(0).Enabled = cboTimes.ListCount > 0
            cmdCheck(1).Enabled = cboTimes.ListCount > 0
        End If
    End With
End Sub


