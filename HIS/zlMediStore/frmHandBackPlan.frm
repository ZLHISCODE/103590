VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHandBackPlan 
   Caption         =   "药品退药计划"
   ClientHeight    =   8640
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   11760
   Icon            =   "frmHandBackPlan.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8640
   ScaleWidth      =   11760
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraControl 
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   7800
      Width           =   11895
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除(&D)"
         Height          =   350
         Left            =   7080
         TabIndex        =   17
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "新增(&A)"
         Height          =   350
         Left            =   4680
         TabIndex        =   14
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "过滤(&F)"
         Height          =   350
         Left            =   1320
         TabIndex        =   13
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "刷新(&R)"
         Height          =   350
         Left            =   2520
         TabIndex        =   12
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "打印(&P)"
         Height          =   350
         Left            =   9480
         TabIndex        =   9
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "修改(&M)"
         Height          =   350
         Left            =   5880
         TabIndex        =   5
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "退出(&E)"
         Height          =   350
         Left            =   10680
         TabIndex        =   4
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdVerify 
         Caption         =   "审核(&V)"
         Height          =   350
         Left            =   8280
         TabIndex        =   3
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1100
      End
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   8280
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmHandBackPlan.frx":038A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15663
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
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
   Begin TabDlg.SSTab tabMain 
      Height          =   7815
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   13785
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   -2147483644
      TabCaption(0)   =   "     未审核计划(&0)     "
      TabPicture(0)   =   "frmHandBackPlan.frx":0C1E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vsfMain(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "vsfDetail(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "picHsc(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "     已审核计划(&1)     "
      TabPicture(1)   =   "frmHandBackPlan.frx":0C3A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picHsc(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "vsfMain(1)"
      Tab(1).Control(2)=   "vsfDetail(1)"
      Tab(1).ControlCount=   3
      Begin VB.PictureBox picHsc 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Index           =   0
         Left            =   100
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   5460
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2600
         Width           =   5460
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   3255
         Index           =   0
         Left            =   105
         TabIndex        =   8
         Top             =   3000
         Width           =   6495
         _cx             =   11456
         _cy             =   5741
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   15724527
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmHandBackPlan.frx":0C56
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
      Begin VSFlex8Ctl.VSFlexGrid vsfMain 
         Height          =   1995
         Index           =   0
         Left            =   105
         TabIndex        =   7
         Top             =   480
         Width           =   6495
         _cx             =   11456
         _cy             =   3519
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   15724527
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmHandBackPlan.frx":0CCB
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
      Begin VB.PictureBox picHsc 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Index           =   1
         Left            =   -74900
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   5460
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2600
         Width           =   5460
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfMain 
         Height          =   1935
         Index           =   1
         Left            =   -74895
         TabIndex        =   10
         Top             =   480
         Width           =   6495
         _cx             =   11456
         _cy             =   3413
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   15724527
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmHandBackPlan.frx":0D40
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
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   3495
         Index           =   1
         Left            =   -74895
         TabIndex        =   11
         Top             =   3000
         Width           =   6495
         _cx             =   11456
         _cy             =   6165
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   15724527
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmHandBackPlan.frx":0DB5
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
   End
End
Attribute VB_Name = "frmHandBackPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng库房ID As Long
Private mintUnit As String
Private mblnRefresh As Boolean

Private mstrSqlFilter As String             '用于过滤
Private mstrBegin As String                 '记录默认的开始时间
Private mstrEnd As String                   '记录默认的结束时间
Private Const MStrCaption As String = "药品退药计划"

Private Type Type_SQLCondition
    strNO开始 As String
    strNO结束 As String
    str填制时间开始 As String
    str填制时间结束 As String
    str审核时间开始 As String
    str审核时间结束 As String
    lng药品ID As Long
    lng供应商ID As Long
    str生产商 As String
End Type

Private SQLCondition As Type_SQLCondition

Private Enum BillType
    未审核 = 0
    已审核 = 1
End Enum

'汇总，明细列表标题
Private Const mconstMainHead = "序号,4,500|No,1,1000|成本金额,7,1200|填制人,1,1000|填制日期,7,2000|审核人,1,1000|审核日期,1,2000|摘要,1,3000"
Private Const mconstDetailHead = "序号,4,500|供应商,1,3000|药品编码,1,1000|药品名称,1,2500|商品名,1,2000|规格,1,2000|生产商,1,2000|批号,1,1000|效期,1,1000|单位,1,800|数量,7,1000|成本价,7,1000|成本金额,7,1000|包装,7,0"

Private Enum 汇总列表
    序号 = 0
    NO = 1
    成本金额 = 2
    填制人 = 3
    填制日期 = 4
    审核人 = 5
    审核日期 = 6
    摘要 = 7
    
    列数 = 8
End Enum

Private Enum 明细列表
    序号 = 0
    供应商 = 1
    药品编码 = 2
    药品名称 = 3
    商品名 = 4
    规格 = 5
    生产商 = 6
    批号 = 7
    效期 = 8
    单位 = 9
    数量 = 10
    成本价 = 11
    成本金额 = 12
    包装 = 13

    列数 = 14
End Enum

Private Sub GetMainDate(ByVal intType As Integer)
    '提取汇总药品计划记录
    'intType：0-未审核;1-已审核
    
    Dim rsTmp As ADODB.Recordset
    Dim strSqlCondition As String
    
    On Error GoTo errHandle
    If SQLCondition.strNO开始 <> "" And SQLCondition.strNO结束 <> "" Then
        strSqlCondition = strSqlCondition & " And A.No >= [1] And A.No <=[2] "
    ElseIf SQLCondition.strNO开始 <> "" Then
        strSqlCondition = strSqlCondition & " And A.No >= [1] "
    ElseIf SQLCondition.strNO结束 <> "" Then
        strSqlCondition = strSqlCondition & " And A.No <=[2] "
    End If
    
    If intType = BillType.未审核 And SQLCondition.str填制时间开始 <> "" And SQLCondition.str填制时间结束 <> "" Then
        strSqlCondition = strSqlCondition & " And A.填制日期 Between [3] And [4] "
    End If
    
    If intType = BillType.已审核 And SQLCondition.str审核时间开始 <> "" And SQLCondition.str审核时间结束 <> "" Then
        strSqlCondition = strSqlCondition & " And A.审核日期 Between [5] And [6] "
    End If
     
    If SQLCondition.lng药品ID > 0 Then
        strSqlCondition = strSqlCondition & " And A.药品id=[7] "
    End If
    
    If SQLCondition.lng供应商ID > 0 Then
        strSqlCondition = strSqlCondition & " And A.供药单位ID + 0 =[8] "
    End If
    
    If SQLCondition.str生产商 <> "" Then
        strSqlCondition = strSqlCondition & " And A.产地=[9] "
    End If
    
    If intType = BillType.未审核 Then
        gstrSQL = "Select A.NO, Sum(A.成本金额) As 成本金额, A.填制人, A.填制日期, A.审核人, A.审核日期, A.摘要 " & _
                " From 药品退药计划 A, 供应商 B " & _
                " Where A.供药单位id = B.ID And 审核人 Is Null " & strSqlCondition & _
                " Group By A.NO, A.填制人, A.填制日期, A.审核人, A.审核日期, A.摘要 " & _
                " Order By NO"
        Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "提取汇总信息", _
                SQLCondition.strNO开始, _
                SQLCondition.strNO结束, _
                CDate(SQLCondition.str填制时间开始), _
                CDate(SQLCondition.str填制时间结束), _
                CDate(SQLCondition.str审核时间开始), _
                CDate(SQLCondition.str审核时间结束), _
                SQLCondition.lng药品ID, _
                SQLCondition.lng供应商ID, _
                SQLCondition.str生产商)
    Else
        gstrSQL = "Select A.NO, Sum(A.成本金额) As 成本金额, A.填制人, A.填制日期, A.审核人, A.审核日期, A.摘要 " & _
                " From 药品退药计划 A, 供应商 B " & _
                " Where A.供药单位id = B.ID And 审核人 Is Not Null " & strSqlCondition & _
                " Group By A.NO, A.填制人, A.填制日期, A.审核人, A.审核日期, A.摘要 " & _
                " Order By NO"
        Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "提取汇总信息", _
                SQLCondition.strNO开始, _
                SQLCondition.strNO结束, _
                CDate(SQLCondition.str填制时间开始), _
                CDate(SQLCondition.str填制时间结束), _
                CDate(SQLCondition.str审核时间开始), _
                CDate(SQLCondition.str审核时间结束), _
                SQLCondition.lng药品ID, _
                SQLCondition.lng供应商ID, _
                SQLCondition.str生产商)
    End If
    
    vsfMain(intType).rows = 1
    vsfDetail(intType).rows = 1
    
    If rsTmp.EOF Then Exit Sub
    
    With rsTmp
        Do While Not .EOF
            vsfMain(intType).rows = vsfMain(intType).rows + 1
            
            vsfMain(intType).TextMatrix(vsfMain(intType).rows - 1, 汇总列表.序号) = .AbsolutePosition
            vsfMain(intType).TextMatrix(vsfMain(intType).rows - 1, 汇总列表.NO) = !NO
            vsfMain(intType).TextMatrix(vsfMain(intType).rows - 1, 汇总列表.成本金额) = zlStr.FormatEx(!成本金额, 2, , True)
            vsfMain(intType).TextMatrix(vsfMain(intType).rows - 1, 汇总列表.填制人) = Nvl(!填制人)
            vsfMain(intType).TextMatrix(vsfMain(intType).rows - 1, 汇总列表.填制日期) = Format(!填制日期, "yyyy-mm-dd hh:mm:ss")
            vsfMain(intType).TextMatrix(vsfMain(intType).rows - 1, 汇总列表.审核人) = Nvl(!审核人)
            vsfMain(intType).TextMatrix(vsfMain(intType).rows - 1, 汇总列表.审核日期) = Format(!审核日期, "yyyy-mm-dd hh:mm:ss")
            vsfMain(intType).TextMatrix(vsfMain(intType).rows - 1, 汇总列表.摘要) = Nvl(!摘要)
            
            .MoveNext
        Loop
    End With
    
    Call GetDetailDate(intType, vsfMain(intType).TextMatrix(1, 汇总列表.NO))
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetDetailDate(ByVal intType As Integer, ByVal strNo As String)
    '提取明细药品计划记录
    'intType：0-未审核;1-已审核
    
    Dim rsTmp As ADODB.Recordset
    Dim strSubUnit As String
    
    '单位，包装换算
    '单位系数：1-售价;2-门诊;3-住院;4-药库
    On Error GoTo errHandle
    Select Case mintUnit
    Case 1
        strSubUnit = "D.计算单位 单位,1 包装 "
    Case 2
        strSubUnit = "B.门诊单位 单位,B.门诊包装 包装 "
    Case 3
        strSubUnit = "B.住院单位 单位,B.住院包装 包装 "
    Case 4
        strSubUnit = "B.药库单位 单位,B.药库包装 包装 "
    End Select
    
    gstrSQL = "Select Distinct A.序号, P.名称 As 供应商, A.药品id, D.编码 As 药品编码,D.名称 As 通用名,E.名称 As 商品名, " & _
        " D.规格, A.实际数量,A.效期, A.成本价, A.成本金额, A.产地 As 生产商, A.批号, " & strSubUnit & _
        " From 药品退药计划 A, 药品规格 B, 收费项目目录 D, 收费项目别名 E, 供应商 P " & _
        " Where A.药品id = B.药品id And B.药品id = D.ID And A.供药单位ID = P.ID And B.药品id = E.收费细目id(+) And E.性质(+) = 3 " & _
        " And A.NO = [1] " & _
        " Order By A.序号"
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "取退药明细信息", strNo)
    
    vsfDetail(intType).rows = 1
    
    If rsTmp.EOF Then Exit Sub
    
    With rsTmp
        Do While Not .EOF
            vsfDetail(intType).rows = vsfDetail(intType).rows + 1
            
            vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, 明细列表.序号) = !序号
            vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, 明细列表.供应商) = !供应商
            vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, 明细列表.药品编码) = !药品编码
            
            If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
                vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, 明细列表.药品名称) = !通用名
            Else
                vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, 明细列表.药品名称) = IIf(IsNull(!商品名), !通用名, !商品名)
            End If
            
            vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, 明细列表.商品名) = IIf(IsNull(!商品名), "", !商品名)
            
            vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, 明细列表.规格) = !规格
            vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, 明细列表.单位) = !单位
            vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, 明细列表.数量) = zlStr.FormatEx(!实际数量 / !包装, 2, , True)
            vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, 明细列表.成本价) = zlStr.FormatEx(!成本价 * !包装, 5, , True)
            vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, 明细列表.成本金额) = zlStr.FormatEx(!成本金额, 2, , True)
            vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, 明细列表.生产商) = IIf(IsNull(!生产商), "", !生产商)
            vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, 明细列表.批号) = IIf(IsNull(!批号), "", !批号)
            vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, 明细列表.包装) = !包装
            vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, 明细列表.效期) = Format(IIf(IsNull(!效期), "", !效期), "yyyy-mm-dd")
            If gtype_UserSysParms.P149_效期显示方式 = 1 And vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, 明细列表.效期) <> "" Then
                '换算为有效期
                vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, 明细列表.效期) = Format(DateAdd("D", -1, vsfDetail(intType).TextMatrix(vsfDetail(intType).rows - 1, 明细列表.效期)), "yyyy-mm-dd")
            End If
            
            .MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ShowForm(FrmMain As Form, ByVal lng库房ID As Long, ByVal intUnit As Integer)
    mlng库房ID = lng库房ID
    mintUnit = intUnit
    
    Me.Show vbModal, FrmMain
End Sub
Private Sub IniGrid()
    Dim i As Integer
    Dim strArr As Variant
    Dim strTemp As Variant
    
    '初始汇总列表（未审核）
    strTemp = Split(mconstMainHead, "|")
    With vsfMain(BillType.未审核)
        .Redraw = flexRDNone
        .rows = 1
        .Cols = 汇总列表.列数
        .SelectionMode = flexSelectionByRow
        .ExplorerBar = flexExSortShow
        For i = 0 To .Cols - 1
            strArr = Split(strTemp(i), ",")
            .TextMatrix(0, i) = strArr(0)
            .ColAlignment(i) = strArr(1)
            .ColWidth(i) = strArr(2)
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .ColWidth(汇总列表.审核人) = 0
        .ColWidth(汇总列表.审核日期) = 0
        .Redraw = flexRDDirect
    End With
    
    '初始汇总列表（已审核）
    strTemp = Split(mconstMainHead, "|")
    With vsfMain(BillType.已审核)
        .Redraw = flexRDNone
        .rows = 1
        .Cols = 汇总列表.列数
        .SelectionMode = flexSelectionByRow
        .ExplorerBar = flexExSortShow
        For i = 0 To .Cols - 1
            strArr = Split(strTemp(i), ",")
            .TextMatrix(0, i) = strArr(0)
            .ColAlignment(i) = strArr(1)
            .ColWidth(i) = strArr(2)
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
       
        .Redraw = flexRDDirect
    End With
    
    '初始明细列表（未审核）
    strTemp = Split(mconstDetailHead, "|")
    With vsfDetail(BillType.未审核)
        .Redraw = flexRDNone
        .rows = 1
        .Cols = 明细列表.列数
        .SelectionMode = flexSelectionByRow
        .ExplorerBar = flexExSortShow
        For i = 0 To .Cols - 1
            strArr = Split(strTemp(i), ",")
            .TextMatrix(0, i) = strArr(0)
            .ColAlignment(i) = strArr(1)
            .ColWidth(i) = strArr(2)
            
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .Redraw = flexRDDirect
    End With
    
    '初始明细列表（已审核）
    strTemp = Split(mconstDetailHead, "|")
    With vsfDetail(BillType.已审核)
        .Redraw = flexRDNone
        .rows = 1
        .Cols = 明细列表.列数
        .SelectionMode = flexSelectionByRow
        .ExplorerBar = flexExSortShow
        For i = 0 To .Cols - 1
            strArr = Split(strTemp(i), ",")
            .TextMatrix(0, i) = strArr(0)
            .ColAlignment(i) = strArr(1)
            .ColWidth(i) = strArr(2)
            
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub cmdAdd_Click()
    frmHandBackPlanModify.ShowForm Me, mlng库房ID, mintUnit, mblnRefresh
    
    If mblnRefresh = True Then
        Call GetMainDate(BillType.未审核)
    End If
End Sub
Private Sub cmdDel_Click()
    With vsfMain(0)
        If .Row = 0 Then Exit Sub
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        
        If MsgBox("是否删除退药计划单[" & .TextMatrix(.Row, 汇总列表.NO) & "]？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            gstrSQL = "Zl_药品退药计划_Delete('" & .TextMatrix(.Row, 汇总列表.NO) & "')"
            Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
            Call GetMainDate(BillType.未审核)
        End If
    End With
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFilter_Click()
    '设置过滤条件
    If frmHandBackSearch.GetSearch(Me, tabMain.Tab, _
                mlng库房ID, _
                SQLCondition.strNO开始, _
                SQLCondition.strNO结束, _
                SQLCondition.str填制时间开始, _
                SQLCondition.str填制时间结束, _
                SQLCondition.str审核时间开始, _
                SQLCondition.str审核时间结束, _
                SQLCondition.lng药品ID, _
                SQLCondition.lng供应商ID, _
                SQLCondition.str生产商) = True Then
        Call cmdRefresh_Click
    Else
        If SQLCondition.str填制时间开始 = "" Or SQLCondition.str填制时间结束 = "" Then
            SQLCondition.str填制时间开始 = mstrBegin
            SQLCondition.str填制时间结束 = mstrEnd
        End If

        If SQLCondition.str审核时间开始 = "" Or SQLCondition.str审核时间结束 = "" Then
            SQLCondition.str审核时间开始 = mstrBegin
            SQLCondition.str审核时间结束 = mstrEnd
        End If
    End If
End Sub
Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdModify_Click()
    With vsfMain(0)
        If .Row = 0 Then Exit Sub
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        
        frmHandBackPlanModify.ShowForm Me, mlng库房ID, mintUnit, mblnRefresh, .TextMatrix(.Row, 汇总列表.NO)
        
        If mblnRefresh = True Then
            Call GetMainDate(BillType.未审核)
        End If
    End With
End Sub

Private Sub cmdPrint_Click()
    If vsfMain(tabMain.Tab).Row = 0 Then Exit Sub
    If vsfMain(tabMain.Tab).TextMatrix(vsfMain(tabMain.Tab).Row, 0) = "" Then Exit Sub
    ReportOpen gcnOracle, glngSys, "ZL1_BILL_1300_1", Me, "No=" & vsfMain(tabMain.Tab).TextMatrix(vsfMain(tabMain.Tab).Row, 汇总列表.NO), "单位系数=" & mintUnit, 1
End Sub
Private Sub cmdRefresh_Click()
     Call GetMainDate(tabMain.Tab)
End Sub
Private Sub cmdVerify_Click()
    With vsfMain(0)
        If .Row = 0 Then Exit Sub
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        
        gstrSQL = "Zl_药品退药计划_Verify('" & .TextMatrix(.Row, 汇总列表.NO) & "','" & UserInfo.用户姓名 & "')"
        Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
        Call GetMainDate(BillType.未审核)
    End With
End Sub

Private Sub Form_Load()
    Dim dateCurr As Date
    
    tabMain.Tab = 0
    picHsc(0).Visible = True
    vsfMain(0).Visible = True
    vsfDetail(0).Visible = True
    
    picHsc(1).Visible = False
    vsfMain(1).Visible = False
    vsfDetail(1).Visible = False
    If InStr(1, gstrprivs, ";打印药品退药计划单;") > 0 Then
        cmdPrint.Visible = True
    Else
        cmdPrint.Visible = False
        cmdVerify.Left = cmdPrint.Left
        cmdDel.Left = cmdVerify.Left - cmdVerify.Width - 100
        cmdModify.Left = cmdDel.Left - cmdDel.Width - 100
        cmdAdd.Left = cmdModify.Left - cmdModify.Width - 100
    End If
    
    Call IniGrid
    
    dateCurr = Sys.Currentdate
    mstrBegin = Format(dateCurr, "YYYY-MM") & "-01 00:00:00"
    mstrEnd = Format(dateCurr, "YYYY-MM-DD") & " 23:59:59"
    SQLCondition.str填制时间开始 = mstrBegin
    SQLCondition.str填制时间结束 = mstrEnd
    SQLCondition.str审核时间开始 = mstrBegin
    SQLCondition.str审核时间结束 = mstrEnd
    
    Call GetMainDate(BillType.未审核)
    Call GetMainDate(BillType.已审核)
    
    RestoreWinState Me, App.ProductName, MStrCaption
        
    '商品名列处理
    If gint药品名称显示 = 2 Then
        '显示商品名列
        vsfDetail(BillType.未审核).ColWidth(明细列表.商品名) = IIf(vsfDetail(BillType.未审核).ColWidth(明细列表.商品名) = 0, 2000, vsfDetail(BillType.未审核).ColWidth(明细列表.商品名))
        vsfDetail(BillType.已审核).ColWidth(明细列表.商品名) = IIf(vsfDetail(BillType.已审核).ColWidth(明细列表.商品名) = 0, 2000, vsfDetail(BillType.已审核).ColWidth(明细列表.商品名))
    Else
        '不单独显示商品名列
        vsfDetail(BillType.未审核).ColWidth(明细列表.商品名) = 0
        vsfDetail(BillType.已审核).ColWidth(明细列表.商品名) = 0
    End If
End Sub


Private Sub Form_Resize()
    '窗体位置设置
    On Error Resume Next
    
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.WindowState <> vbMaximized Then
        If Me.Height < 9255 Then
            Me.Height = 9255
        End If
        
        If Me.Width < 12165 Then
            Me.Width = 12165
        End If
    End If
    
    With fraControl
        .Left = 0
        .Top = Me.ScaleHeight - fraControl.Height - IIf(staThis.Visible, staThis.Height, 0)
        .Width = Me.ScaleWidth - .Left
        .Height = 600
    End With
    
    With tabMain
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top - fraControl.Height - IIf(staThis.Visible, staThis.Height, 0)
    End With
    
    With picHsc(0)
        .Height = 45
        .Left = 100
        .Width = tabMain.Width - 200
    End With
    
    With vsfMain(0)
        .Top = 480
        .Left = 100
        .Width = tabMain.Width - 200
        .Height = picHsc(0).Top - .Top
    End With
    
    With vsfDetail(0)
        .Top = picHsc(0).Top + picHsc(0).Height + 50
        .Left = vsfMain(0).Left
        .Height = tabMain.Height - .Top - 100
        .Width = vsfMain(0).Width
    End With
    
    With picHsc(1)
        .Height = 45
        .Left = 100
        .Width = tabMain.Width - 200
    End With
    
    With vsfMain(1)
        .Top = 480
        .Left = 100
        .Width = tabMain.Width - 200
        .Height = picHsc(1).Top - .Top
    End With
    
    With vsfDetail(1)
        .Top = picHsc(1).Top + picHsc(1).Height + 50
        .Left = vsfMain(1).Left
        .Height = tabMain.Height - .Top - 100
        .Width = vsfMain(1).Width
    End With

End Sub


Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, MStrCaption
End Sub

Private Sub picHsc_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If vsfMain(Index).Height + y <= 500 Or vsfDetail(Index).Height - y <= 500 Then Exit Sub
        
        picHsc(Index).Top = picHsc(Index).Top + y
        vsfMain(Index).Height = vsfMain(Index).Height + y
        vsfDetail(Index).Height = vsfDetail(Index).Height - y
        vsfDetail(Index).Top = vsfDetail(Index).Top + y
        
        Me.Refresh
    End If
End Sub

Private Sub tabMain_Click(PreviousTab As Integer)
    If tabMain.Tab = BillType.未审核 Then
        picHsc(0).Visible = True
        vsfMain(0).Visible = True
        vsfDetail(0).Visible = True

        picHsc(1).Visible = False
        vsfMain(1).Visible = False
        vsfDetail(1).Visible = False

        cmdAdd.Enabled = True
        cmdModify.Enabled = True
        cmdDel.Enabled = True
        cmdVerify.Enabled = True
    Else
        picHsc(0).Visible = False
        vsfMain(0).Visible = False
        vsfDetail(0).Visible = False

        picHsc(1).Visible = True
        vsfMain(1).Visible = True
        vsfDetail(1).Visible = True

        cmdAdd.Enabled = False
        cmdModify.Enabled = False
        cmdDel.Enabled = False
        cmdVerify.Enabled = False
    End If
End Sub

Private Sub vsfMain_Click(Index As Integer)
    With vsfMain(Index)
        If .Row = 0 Then Exit Sub
        Call GetDetailDate(Index, .TextMatrix(.Row, 汇总列表.NO))
    End With
End Sub


