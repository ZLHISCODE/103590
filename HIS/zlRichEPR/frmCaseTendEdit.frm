VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCaseTendEdit 
   Caption         =   "护理记录"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmCaseTendEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   11880
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2850
      Index           =   2
      Left            =   240
      ScaleHeight     =   2850
      ScaleWidth      =   10920
      TabIndex        =   32
      Top             =   4215
      Width           =   10920
      Begin VB.Frame fraTime 
         Height          =   525
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   11235
         Begin VB.ComboBox cbo 
            Height          =   300
            Left            =   8235
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   150
            Width           =   1680
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Left            =   1065
            TabIndex        =   37
            Top             =   150
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   177471491
            UpDown          =   -1  'True
            CurrentDate     =   38702
         End
         Begin MSComctlLib.TabStrip tbs 
            Height          =   300
            Left            =   3765
            TabIndex        =   40
            Top             =   165
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   529
            MultiRow        =   -1  'True
            Style           =   2
            TabMinWidth     =   529
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   2
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "1"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "2"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "记录组(&G)"
            Height          =   180
            Index           =   0
            Left            =   2925
            TabIndex        =   41
            Top             =   210
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "发生时间(&T)"
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   38
            Top             =   195
            Width           =   990
         End
      End
      Begin zlRichEPR.VsfGrid vsf 
         Height          =   1575
         Left            =   150
         TabIndex        =   34
         Top             =   630
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2778
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   1695
      Index           =   1
      Left            =   3210
      ScaleHeight     =   1695
      ScaleWidth      =   8445
      TabIndex        =   26
      Top             =   6105
      Width           =   8445
      Begin VB.Frame fra 
         Height          =   525
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   11235
         Begin MSComCtl2.UpDown udnDay 
            Height          =   270
            Left            =   585
            TabIndex        =   31
            Top             =   150
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            BuddyControl    =   "txtDay"
            BuddyDispid     =   196618
            OrigLeft        =   810
            OrigTop         =   180
            OrigRight       =   1065
            OrigBottom      =   405
            Max             =   30
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtDay 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   270
            Left            =   270
            Locked          =   -1  'True
            TabIndex        =   30
            Text            =   "1"
            Top             =   165
            Width           =   315
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "前       天历史记录结果"
            Height          =   180
            Index           =   12
            Left            =   60
            TabIndex        =   29
            Top             =   195
            Width           =   2070
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfHistory 
         Height          =   1020
         Left            =   240
         TabIndex        =   27
         Top             =   660
         Width           =   1995
         _cx             =   3519
         _cy             =   1799
         Appearance      =   2
         BorderStyle     =   0
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
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   -1  'True
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
   End
   Begin VB.PictureBox picCustom 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1740
      ScaleHeight     =   300
      ScaleWidth      =   1995
      TabIndex        =   2
      Top             =   75
      Width           =   1995
      Begin VB.CommandButton cmd 
         Height          =   300
         Left            =   1665
         Picture         =   "frmCaseTendEdit.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   0
         Width           =   330
      End
      Begin VB.TextBox txt 
         Height          =   300
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1665
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   3435
      Index           =   0
      Left            =   810
      ScaleHeight     =   3435
      ScaleWidth      =   9885
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   645
      Width           =   9885
      Begin VB.Frame fraInfo 
         Height          =   705
         Left            =   0
         TabIndex        =   4
         Top             =   -90
         Width           =   9780
         Begin VB.ComboBox cboBaby 
            Height          =   300
            Left            =   8370
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   255
            Width           =   1350
         End
         Begin VB.TextBox txtShow 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   9
            Left            =   2175
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   435
            Width           =   1200
         End
         Begin VB.TextBox txtShow 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   8
            Left            =   5910
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   435
            Width           =   2355
         End
         Begin VB.TextBox txtShow 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   7
            Left            =   555
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   435
            Width           =   1185
         End
         Begin VB.TextBox txtShow 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   6
            Left            =   7050
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   180
            Width           =   1215
         End
         Begin VB.TextBox txtShow 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   5
            Left            =   3990
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   435
            Width           =   1455
         End
         Begin VB.TextBox txtShow 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   4
            Left            =   5910
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   180
            Width           =   600
         End
         Begin VB.TextBox txtShow 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   3
            Left            =   3990
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   180
            Width           =   1455
         End
         Begin VB.TextBox txtShow 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   2
            Left            =   2985
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   180
            Width           =   375
         End
         Begin VB.TextBox txtShow 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   1
            Left            =   2175
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   180
            Width           =   360
         End
         Begin VB.TextBox txtShow 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   0
            Left            =   555
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   180
            Width           =   1185
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "入  院:"
            Height          =   180
            Index           =   11
            Left            =   3375
            TabIndex        =   23
            Top             =   435
            Width           =   630
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "护理:"
            Height          =   180
            Index           =   10
            Left            =   120
            TabIndex        =   13
            Top             =   435
            Width           =   450
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "诊断:"
            Height          =   180
            Index           =   9
            Left            =   5475
            TabIndex        =   12
            Top             =   420
            Width           =   450
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "床号:"
            Height          =   180
            Index           =   8
            Left            =   5475
            TabIndex        =   11
            Top             =   165
            Width           =   450
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "医师:"
            Height          =   180
            Index           =   7
            Left            =   6615
            TabIndex        =   10
            Top             =   180
            Width           =   450
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "病情:"
            Height          =   180
            Index           =   6
            Left            =   1755
            TabIndex        =   9
            Top             =   435
            Width           =   450
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "性别:"
            Height          =   180
            Index           =   5
            Left            =   1740
            TabIndex        =   8
            Top             =   165
            Width           =   450
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "住院号:"
            Height          =   180
            Index           =   4
            Left            =   3375
            TabIndex        =   7
            Top             =   165
            Width           =   630
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "年龄:"
            Height          =   180
            Index           =   3
            Left            =   2550
            TabIndex        =   6
            Top             =   165
            Width           =   450
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "姓名:"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   5
            Top             =   165
            Width           =   450
         End
      End
      Begin XtremeSuiteControls.TabControl tbcPage 
         Height          =   2025
         Left            =   270
         TabIndex        =   33
         Top             =   975
         Width           =   2700
         _Version        =   589884
         _ExtentX        =   4762
         _ExtentY        =   3572
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7740
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCaseTendEdit.frx":10C8
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14843
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
            Picture         =   "frmCaseTendEdit.frx":195A
            Text            =   "范围："
            TextSave        =   "范围："
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmCaseTendEdit.frx":81BC
      Left            =   345
      Top             =   600
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmCaseTendEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'窗口级变量申明
'######################################################################################################################

Private mblnStartUp As Boolean
Private mblnOk As Boolean
Private mstrSQL As String
Private mbytMode As Byte                    '1-新增记录;2-修改记录;3-电子签名;4-取消签名:5-历史版本
Private mlngKey As Long
Private mstrTime As String
Private mrsParam As ADODB.Recordset
Private mblnChanged As Boolean
Private mblnNoChanged As Boolean
Private mstrSvrDate As String
Private mint饮入量 As Integer
Private mblnReading As Boolean
Private mstrSvr姓名 As String
Private mstr就诊卡字母前缀 As String
Private mint就诊卡号码长度 As Integer
Private mstrPrivs As String
Private mrsPatient As ADODB.Recordset
Private mlngRowNum As Long
Private mstrFindKey As String
Private mobjFindKey As CommandBarControl
Private mint心率应用 As Integer
Private mblnDefault As Boolean
Private mclsVsfHistory As clsVsf
Private mintPreDays As Integer

Private Enum mCol
    记录组 = 1
    护理项目
    项目单位
    项目类型
    项目长度
    项目小数
    项目表示
    项目值域
    项目缺省
    项目性质
    项目id
    是否变动
    记录结果
    标记
    部位
    未记说明
End Enum

Public Event AfterDataChanged()

'自定义过程/函数申明
'######################################################################################################################

Private Property Let DataChanged(vData As Boolean)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    mblnChanged = vData
        
End Property

Private Property Get DataChanged() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    
    DataChanged = mblnChanged
    cmd.Enabled = Not mblnChanged And (mbytMode = 1 Or mbytMode = 2 Or mbytMode = 5)
    tbs.Enabled = Not mblnChanged
    cboBaby.Enabled = Not mblnChanged
        
    For intLoop = 0 To tbcPage.ItemCount - 1
        If Not tbcPage.Item(intLoop) Is Nothing Then
            tbcPage.Item(intLoop).Enabled = Not mblnChanged
        End If
    Next

End Property

Public Function ShowEdit(ByVal frmParent As Form, ByVal strParam As String, Optional ByVal bytMode As Byte = 1, Optional ByVal strPrivs As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：strParam->病人id;主页id;科室id;婴儿;来源;时间,ID
    '返回：
    '******************************************************************************************************************
    Dim varParam As Variant
        
    mblnStartUp = True
    
    mbytMode = bytMode
    mblnOk = False
    mstrPrivs = strPrivs
        
    '------------------------------------------------------------------------------------------------------------------
    Set mrsParam = New ADODB.Recordset
    Call CreateParam(mrsParam, "病人id", adBigInt)
    Call CreateParam(mrsParam, "主页id", adBigInt)
    Call CreateParam(mrsParam, "科室id", adBigInt)
    Call CreateParam(mrsParam, "病区id", adBigInt)
    Call CreateParam(mrsParam, "婴儿", adTinyInt)
    Call CreateParam(mrsParam, "版本", adTinyInt)
    Call CreateParam(mrsParam, "来源", adTinyInt)
    Call CreateParam(mrsParam, "出院", adTinyInt)
    Call CreateParam(mrsParam, "时间", adVarChar, 20)
    Call CreateParam(mrsParam, "护理等级", adTinyInt)
    Call CreateParam(mrsParam, "出院开始日期", adVarChar, 30)
    Call CreateParam(mrsParam, "出院结束日期", adVarChar, 30)
    Call CreateParam(mrsParam, "在院病人", adTinyInt)
    Call CreateParam(mrsParam, "出院病人", adTinyInt)
    Call CreateParam(mrsParam, "转出病人", adTinyInt)
    Call CreateParam(mrsParam, "待入科病人", adTinyInt)
    Call CreateParam(mrsParam, "转出天数", adTinyInt)
    Call CreateParam(mrsParam, "记录id", adBigInt)
    
    '------------------------------------------------------------------------------------------------------------------
    If strParam <> "" Then varParam = Split(strParam, ";")
    mrsParam.Open
    mrsParam.AddNew
                    
    mrsParam("病人id").Value = Val(varParam(0))
    mrsParam("主页id").Value = Val(varParam(1))
    mrsParam("科室id").Value = Val(varParam(2))
    mrsParam("病区id").Value = Val(varParam(2))
    mrsParam("护理等级").Value = 3
    mrsParam("婴儿").Value = 0
    mrsParam("版本").Value = 0
    If UBound(varParam) >= 3 Then mrsParam("婴儿").Value = Val(varParam(3))
    If UBound(varParam) >= 4 Then mrsParam("来源").Value = Val(varParam(4))
    If UBound(varParam) >= 5 Then mrsParam("时间").Value = CStr(varParam(5))
    If UBound(varParam) >= 5 Then mrsParam("时间").Value = CStr(varParam(5))
    
    '初始控件
    '------------------------------------------------------------------------------------------------------------------
    If ExecuteCommand("初始控件") = False Then Exit Function
    
    
    '初始数据
    '------------------------------------------------------------------------------------------------------------------
    If ExecuteCommand("初始数据") = False Then Exit Function
    
    
    '------------------------------------------------------------------------------------------------------------------
    Call ExecuteCommand("刷新基本信息")
    
    If mbytMode <> 1 Then
        Call ExecuteCommand("读取记录")
    End If
    
    Call ExecuteCommand("清除数据")
    Call ExecuteCommand("读取数据")
    
    If mbytMode <> 1 Then
        '读取指定记录id、指定组的护理内容
        Call ExecuteCommand("读取组别数据")
    Else
        '新增,如果有缺省值，则允许保存
        DataChanged = mblnDefault
    End If
    
    Vsf.Col = mCol.记录结果
    
    DataChanged = False
    mblnStartUp = False
    
    Me.Show , frmParent
    
    ShowEdit = mblnOk
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
EndHand:
    mblnStartUp = False
    Unload Me
End Function

Private Function ReadPatient() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim RS As New ADODB.Recordset
    Dim strParam As String
    On Error GoTo ErrHand
    
    '在院和出院病人:出院病人可能已有多次住院
    '------------------------------------------------------------------------------------------------------------------
    If Val(mrsParam("在院病人").Value) <> 0 Or Val(mrsParam("出院病人").Value) <> 0 Or Val(mrsParam("待入科病人").Value) <> 0 Then
        gstrSQL = _
            "Select Decode(B.出院日期,NULL,Decode(B.状态,3,2,1),Decode(B.出院方式,'死亡',4,3)) as 排序," & _
            " Decode(B.出院日期,NULL,Decode(B.状态,3,'预出院病人','在院病人'),Decode(B.出院方式,'死亡','死亡病人','出院病人')) as 类型," & _
            " A.病人ID,B.主页ID,B.住院号,A.门诊号,NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别,NVL(B.年龄,A.年龄) 年龄,C.名称 as 科室,B.住院医师," & _
            " B.出院病床 as 床号,B.费别,B.入院日期,B.出院日期,B.状态,B.险类,A.就诊卡号" & _
            " From 病人信息 A,病案主页 B,部门表 C" & _
            " Where A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And ([6]=1 Or Nvl(B.状态,0)<>1) And B.出院科室ID=C.ID" & _
            " And B.当前病区ID=[1] And ([4]<>0 And B.出院日期 is NULL Or [5]<>0 And B.出院日期 Between [2] And [3]) "
    End If
    
    '转出病人:在院,医生和床号显示本科转出前的
    '------------------------------------------------------------------------------------------------------------------
    If Val(mrsParam("转出病人").Value) <> 0 Then
        gstrSQL = gstrSQL & IIf(gstrSQL <> "", " Union All ", "") & _
            "Select Distinct 5 as 排序,'转出病人' as 类型," & _
            " A.病人ID,B.主页ID,B.住院号,A.门诊号,NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别,NVL(B.年龄,A.年龄) 年龄,D.名称 as 科室,C.经治医师 as 住院医师," & _
            " C.床号,B.费别,B.入院日期,B.出院日期,B.状态,B.险类,A.就诊卡号" & _
            " From 病人信息 A,病案主页 B,病人变动记录 C,部门表 D" & _
            " Where A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And C.科室ID=D.ID" & _
            " And Nvl(B.状态,0)=0 And B.出院日期 is NULL And B.当前病区ID<>[1]" & _
            " And B.病人ID=C.病人ID And B.主页ID=C.主页ID And C.病区ID=[1]" & _
            " And C.终止原因=3 And C.终止时间 Between Sysdate-[7] And Sysdate "
    End If
    gstrSQL = gstrSQL & " Order by 排序,床号,主页ID Desc"
    gstrSQL = "Select RowNum As ID,1 As 末级,A.* From (" & gstrSQL & ") A"
    
    If mbytMode <> 5 Then
    
        Set mrsPatient = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
                                                                Val(mrsParam("病区id").Value), _
                                                                CDate(Format(mrsParam("出院开始日期").Value, "yyyy-MM-dd 00:00:00")), _
                                                                CDate(Format(mrsParam("出院结束日期").Value, "yyyy-MM-dd 23:59:59")), _
                                                                Val(mrsParam("在院病人").Value), _
                                                                0, _
                                                                Val(mrsParam("待入科病人").Value), _
                                                                Val(mrsParam("转出天数").Value))
                                                            
    Else
        
        Set mrsPatient = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
                                                                Val(mrsParam("病区id").Value), _
                                                                CDate(Format(mrsParam("出院开始日期").Value, "yyyy-MM-dd 00:00:00")), _
                                                                CDate(Format(mrsParam("出院结束日期").Value, "yyyy-MM-dd 23:59:59")), _
                                                                Val(mrsParam("在院病人").Value), _
                                                                Val(mrsParam("出院病人").Value), _
                                                                Val(mrsParam("待入科病人").Value), _
                                                                Val(mrsParam("转出天数").Value))
    End If
                                                            
    ReadPatient = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InitMenuBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrPop As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim RS As ADODB.Recordset
    Dim objExtendedBar As CommandBar
    
    On Error GoTo ErrHand
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsThis.ActiveMenuBar.Title = "菜单栏"
    cbsThis.ActiveMenuBar.Visible = False
    
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    With cbsThis.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 24, 24
        .LargeIcons = True
    End With

    '------------------------------------------------------------------------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    
    
     '快键绑定
    With cbsThis.KeyBindings

        .Add FCONTROL, Asc("S"), conMenu_Edit_Transf_Save
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F2, conMenu_Edit_Transf_Save
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义
    Set cbrToolBar = cbsThis.Add("标准", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagHideWrap
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "饮入计算"): cbrControl.ToolTipText = "饮入计算(Alt+C)"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewParent, "新增"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "新增(Alt+P)"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新组"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "新组(Alt+N)"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Append, "添加"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "添加项目(Alt+A)"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除"):  cbrControl.ToolTipText = "删除项目(Alt+D)"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Save, "保存"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "保存(Alt+S)"
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Sign, "记录签名"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "记录签名(Alt+R)"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消签名"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "取消签名(Alt+U)"
                
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "取消")

        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "帮助(F1)"
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出"): cbrControl.ToolTipText = "退出(Esc)"

    End With
    
    '定位工具栏
    '------------------------------------------------------------------------------------------------------------------
    
    Set objExtendedBar = cbsThis.Add("定位", xtpBarTop)
    
    objExtendedBar.ContextMenuPresent = False
    objExtendedBar.ShowTextBelowIcons = False
    objExtendedBar.EnableDocking xtpFlagHideWrap
    
    With objExtendedBar.Controls

        mstrFindKey = Trim(zlDatabase.GetPara("查找方法", glngSys, 1255, "床  号"))
        If mstrFindKey = "" Then mstrFindKey = "床  号"
        
        Set mobjFindKey = .Add(xtpControlPopup, conMenu_View_LocationItem, mstrFindKey)
        mobjFindKey.IconId = conMenu_View_Find
        mobjFindKey.BeginGroup = True
        mobjFindKey.ToolTipText = "快捷键:F4"
        Set cbrControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&1.床  号"): cbrControl.Parameter = "床  号"
        Set cbrControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&2.住院号"): cbrControl.Parameter = "住院号"
        Set cbrControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&3.就诊卡"): cbrControl.Parameter = "就诊卡"
        
        Set cbrCustom = .Add(xtpControlCustom, conMenu_View_Location, "")
        cbrCustom.flags = xtpFlagRightAlign
        cbrCustom.Handle = picCustom.hWnd
        txt.ToolTipText = "查找病人(F3)"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Forward, "前一病人"): cbrControl.flags = xtpFlagRightAlign: cbrControl.ToolTipText = "前一病人(Ctrl+Left)"
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Backward, "后一病人"): cbrControl.flags = xtpFlagRightAlign: cbrControl.ToolTipText = "后一病人(Ctrl+Right)"
    End With
    
    Call SetDockRight(objExtendedBar, cbrToolBar)
    
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.STYLE = xtpButtonIconAndCaption
        End If
    Next
    
     '快键绑定
    With cbsThis.KeyBindings
        .Add FALT, Asc("C"), conMenu_Edit_Audit
        .Add FALT, Asc("N"), conMenu_Edit_NewItem
        .Add FALT, Asc("A"), conMenu_Edit_Append
        .Add FALT, Asc("D"), conMenu_Edit_Delete
        .Add FALT, Asc("S"), conMenu_Edit_Transf_Save
        .Add FALT, Asc("R"), conMenu_Tool_Sign
        .Add FALT, Asc("U"), conMenu_Edit_Untread
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_ESCAPE, conMenu_File_Exit
        
        .Add 0, vbKeyF3, conMenu_View_Location              '定位
        .Add FCONTROL, vbKeyLeft, conMenu_View_Forward      '前一条
        .Add FCONTROL, vbKeyRight, conMenu_View_Backward    '后一条
    End With
    
    InitMenuBar = True
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetDockRight(BarToDock As CommandBar, BarOnLeft As CommandBar)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    cbsThis.RecalcLayout
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    
    cbsThis.DockToolBar BarToDock, Right, (Bottom + Top) / 2, BarOnLeft.Position

End Sub

Private Function InitData() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim RS As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim varAry As Variant
    Dim lngLoop As Long
    Dim strTmp As String
    
    mint饮入量 = 0
    Select Case mbytMode
    Case 1
        Me.Caption = Me.Caption & " - 登记"
        cbo.Visible = False
    Case 2
        Me.Caption = Me.Caption & " - 修改"
        cbo.Visible = False
    Case 3
        Me.Caption = Me.Caption & " - 签名"
        cbo.Visible = False
    Case 4
        Me.Caption = Me.Caption & " - 取消签名"
        cbo.Visible = False
    Case 5
        Me.Caption = Me.Caption & " - 历史版本"
        cbo.Visible = True
    End Select
    
    cmd.Enabled = (mbytMode = 1 Or mbytMode = 2 Or mbytMode = 5)
    txt.Enabled = cmd.Enabled
    dtp.Enabled = (mbytMode = 1 Or mbytMode = 2)
    
    '------------------------------------------------------------------------------------------------------------------
    With Vsf
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "记录组", 1500, 1
        .NewColumn "记录项目", 1590, 1
        .NewColumn "单位", 750, 1
        .NewColumn "项目类型", 0, 1
        .NewColumn "项目长度", 0, 1
        .NewColumn "项目小数", 0, 1
        .NewColumn "项目表示", 0, 1
        .NewColumn "项目值域", 0, 1
        .NewColumn "项目缺省", 0, 1
        .NewColumn "项目性质", 0, 1
        .NewColumn "项目id", 0, 1
        .NewColumn "是否变动", 0, 1
        
        .NewColumn "记录数据", 3750, 1, , 1
        .NewColumn "标记", 900, 1
        .NewColumn "部位", 900, 1
        .NewColumn "未记说明", 900, 1, "...", 1
        
        .FixedCols = 4
                
        .Body.ColHidden(mCol.项目类型) = True
        .Body.ColHidden(mCol.项目长度) = True
        .Body.ColHidden(mCol.项目小数) = True
        .Body.ColHidden(mCol.项目表示) = True
        .Body.ColHidden(mCol.项目值域) = True
        .Body.ColHidden(mCol.项目缺省) = True
        .Body.ColHidden(mCol.项目性质) = True
        .Body.ColHidden(mCol.项目id) = True
        .Body.MergeCells = flexMergeFree
        .Body.MergeCol(mCol.记录组) = True
        .Body.WordWrap = True
        
        If mbytMode > 2 Then
            .Body.Editable = flexEDNone
'            cmdCalc.Enabled = False
        End If
    End With
    
    Set mclsVsfHistory = New clsVsf
    With mclsVsfHistory
        Call .Initialize(Me.Controls, vsfHistory, True, False)
        Call .ClearColumn
        Call .AppendColumn("记录时间", 1670, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm", , True)
        Call .AppendColumn("历史结果", 15, flexAlignLeftCenter, flexDTString, "", , True)
        vsfHistory.FixedCols = 1
        vsfHistory.ExplorerBar = flexExNone
        vsfHistory.RowHidden(0) = True
        .AppendRows = False
    End With
        
    Dim objPane As Pane
    
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False '实时拖动
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True
    
    dkpMain.SetCommandBars cbsThis
    
    Set objPane = dkpMain.CreatePane(1, 100, 200, DockTopOf, Nothing): objPane.Title = "编辑": objPane.Options = PaneNoCaption
    Set objPane = dkpMain.CreatePane(2, 100, 100, DockBottomOf, objPane): objPane.Title = "历史": objPane.Options = PaneNoCaption
        
    Call InitTabControl
    
    If mbytMode <> 1 And mbytMode <> 2 Then
        dkpMain.Panes(2).Close
        picPane(1).Visible = False
    End If
    
    InitData = True
    
End Function

Private Function InitTabControl() As Boolean
    '******************************************************************************************************************
    '功能：初始Tab控件
    '参数：
    '返回：
    '******************************************************************************************************************
    With tbcPage
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .ShowIcons = True
'            .COLOR = xtpTabColorDefault
            .ColorSet.ButtonSelected = &HFFFF&
            .DisableLunaColors = False
        End With

        Set .Icons = zlCommFun.GetPubIcons

        .InsertItem 0, "次数：1  ", picPane(2).hWnd, 0

        .Item(0).Selected = True
        
    End With
    
    InitTabControl = True
    
End Function

Private Function OpenPatientMap(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal int婴儿 As Integer) As Boolean
    '******************************************************************************************************************
    '功能：读取病人信息
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim RS As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim varAry As Variant
    Dim lngLoop As Long
    Dim strTmp As String
    
    On Error GoTo ErrHand
    
    mblnDefault = False
    
    mrsParam("病人id").Value = lng病人ID
    mrsParam("主页id").Value = lng主页ID
    mrsParam("婴儿").Value = int婴儿
        
    '病人信息
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别,NVL(B.年龄,A.年龄) 年龄,B.住院号,B.出院病床,B.医疗付款方式," & _
        " D.诊断描述,B.险类,B.当前病况,C.名称 as 护理等级,B.入院日期," & _
        " B.出院日期,B.状态,B.数据转出,B.出院科室ID,B.当前病区ID,A.住院次数,B.住院医师 " & _
        " From 病人信息 A,病案主页 B,收费项目目录 C,病人诊断记录 D" & _
        " Where A.病人ID=B.病人ID And A.病人ID=[1] And B.主页ID=[2] And B.护理等级ID=C.ID(+)" & _
        " And D.病人id(+)=B.病人id And D.主页id(+)=B.主页id And D.诊断类型(+)=1 "
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("病人id").Value), Val(mrsParam("主页id").Value))
    
    If RS.BOF = False Then
        txt.Text = zlCommFun.NVL(RS("姓名").Value)
        txt.Tag = ""
        
        txtShow(0).Text = zlCommFun.NVL(RS("姓名").Value)
        txtShow(1).Text = zlCommFun.NVL(RS("性别").Value)
        txtShow(2).Text = zlCommFun.NVL(RS("年龄").Value)
        txtShow(3).Text = zlCommFun.NVL(RS("住院号").Value)
        txtShow(4).Text = zlCommFun.NVL(RS("出院病床").Value)
        txtShow(5).Text = Format(zlCommFun.NVL(RS("入院日期").Value), "yyyy-MM-dd HH:mm")
        txtShow(6).Text = zlCommFun.NVL(RS("住院医师").Value)
        txtShow(7).Text = zlCommFun.NVL(RS("护理等级").Value)
        txtShow(8).Text = zlCommFun.NVL(RS("诊断描述").Value)
        txtShow(9).Text = zlCommFun.NVL(RS("当前病况").Value)

    End If
    mstrSvr姓名 = txt.Text
    
    '
    '------------------------------------------------------------------------------------------------------------------
    cboBaby.Clear
    cboBaby.AddItem "病人本人"
    gstrSQL = "Select a.序号,Decode(a.婴儿姓名,Null,NVL(c.姓名,b.姓名) ||'之子'||Trim(To_Char(a.序号,'9')),a.婴儿姓名) As 婴儿姓名" & _
        " From 病人信息 b,病案主页 c,病人新生儿记录 a Where b.病人id=c.病人id And a.病人id=c.病人ID And a.主页ID=c.主页ID And c.病人id=[1] And c.主页id=[2]  Order By a.序号"
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("病人id").Value), Val(mrsParam("主页id").Value))
    If RS.BOF = False Then
        Do While Not RS.EOF
            cboBaby.AddItem RS("婴儿姓名").Value
            RS.MoveNext
        Loop
    End If
    On Error Resume Next
    cboBaby.ListIndex = Val(mrsParam("婴儿").Value)
    On Error GoTo ErrHand
    If cboBaby.ListIndex = -1 Then cboBaby.ListIndex = 0
    cboBaby.Visible = (cboBaby.ListCount > 1)
    
    
    '获取入院时间
    '------------------------------------------------------------------------------------------------------------------
    mstrSQL = "Select Min(开始时间) As 开始时间 From 病人变动记录 Where 病人id=[1] and 主页id=[2]"
    Set RS = zlDatabase.OpenSQLRecord(mstrSQL, gstrSysName, Val(mrsParam("病人id")), Val(mrsParam("主页id")))
    If RS.BOF = False Then
        If IsNull(RS("开始时间").Value) = False Then
            On Error Resume Next
            dtp.MinDate = Format(DateAdd("n", 1, CDate(Format(RS("开始时间").Value, "yyyy-MM-dd HH:mm") & ":00")), dtp.CustomFormat)
            On Error GoTo ErrHand
        End If
    End If
    
    mintPreDays = Val(zlDatabase.GetPara("超期录入护理数据天数", glngSys, 1255, "1"))
    glngHours = Val(zlDatabase.GetPara("数据补录时限", glngSys))
    
    mstrSQL = "Select 入院日期,出院日期 From 病案主页 Where 病人id=[1] and 主页id=[2]"
    Set RS = zlDatabase.OpenSQLRecord(mstrSQL, gstrSysName, Val(mrsParam("病人id")), Val(mrsParam("主页id")))
    If RS.BOF = False Then

        On Error Resume Next
        
        If IsNull(RS("出院日期").Value) Then
            
            If mintPreDays > 0 Then
                dtp.MaxDate = Format(Format(DateAdd("d", mintPreDays, zlDatabase.Currentdate), "yyyy-MM-dd") & " 23:59:59", dtp.CustomFormat)
            Else
                dtp.MaxDate = Format(zlDatabase.Currentdate, dtp.CustomFormat)
            End If
            
        Else
            dtp.MaxDate = Format(zlCommFun.NVL(RS("出院日期").Value), dtp.CustomFormat)
        End If
        
        dtp.MaxDate = Format(zlCommFun.NVL(RS("出院日期").Value, Format(DateAdd("d", mintPreDays, zlDatabase.Currentdate), "yyyy-MM-dd HH:mm:ss")), dtp.CustomFormat)
        On Error GoTo ErrHand

    End If
        
    '------------------------------------------------------------------------------------------------------------------
    mstrSQL = "Select zl_PatitTendGrade([1],[2]) As 护理等级 From dual"
    Set RS = zlDatabase.OpenSQLRecord(mstrSQL, gstrSysName, Val(mrsParam("病人id")), Val(mrsParam("主页id")))
    If RS.BOF = False Then
        mrsParam("护理等级").Value = zlCommFun.NVL(RS("护理等级").Value)
    End If
    
'    If mrsParam("时间").Value = "" Then
        
        Vsf.Rows = 2
        Vsf.RowData(1) = 0
        Vsf.Cell(flexcpData, 1, 0, 1, Vsf.Cols - 1) = ""
        
'        dtp.Enabled = True
        
        mstrSQL = " Select 项目ID,项目序号,分组名,项目名称,项目类型,项目长度,项目小数,项目表示,项目值域,项目单位 From 护理记录项目 A " & _
                  " Where Nvl(项目性质,1)=1 And Nvl(A.应用方式,0)=1 And Nvl(a.适用病人,0) In (0,[3]) And A.护理等级>=[1] " & _
                  " And (A.适用科室=1 Or (A.适用科室=2 And Exists (Select 1 From 护理适用科室 D Where D.项目序号=A.项目序号 And D.科室id=[2]))) " & _
                  " Order By A.分组名,A.项目序号"
        Set RS = zlDatabase.OpenSQLRecord(mstrSQL, gstrSysName, Val(mrsParam("护理等级").Value), Val(mrsParam("科室id").Value), IIf(Val(mrsParam("婴儿").Value) = 0, 1, 2))
        If RS.BOF = False Then
            Do While Not RS.EOF
                
                If Val(Vsf.RowData(Vsf.Rows - 1)) <> 0 Then Vsf.Rows = Vsf.Rows + 1
                
                Vsf.TextMatrix(Vsf.Rows - 1, mCol.记录组) = zlCommFun.NVL(RS("分组名").Value)
                Vsf.TextMatrix(Vsf.Rows - 1, mCol.护理项目) = zlCommFun.NVL(RS("项目名称").Value)

                Vsf.TextMatrix(Vsf.Rows - 1, mCol.项目类型) = zlCommFun.NVL(RS("项目类型").Value)
                Vsf.TextMatrix(Vsf.Rows - 1, mCol.项目长度) = zlCommFun.NVL(RS("项目长度").Value)
                Vsf.TextMatrix(Vsf.Rows - 1, mCol.项目小数) = zlCommFun.NVL(RS("项目小数").Value)
                Vsf.TextMatrix(Vsf.Rows - 1, mCol.项目表示) = zlCommFun.NVL(RS("项目表示").Value)
                

                If zlCommFun.NVL(RS("项目值域")) <> "" Then

                    varAry = Split(zlCommFun.NVL(RS("项目值域")), ";")

                    For lngLoop = 0 To UBound(varAry)
                        If Left(varAry(lngLoop), 1) = "√" Then
                            mblnDefault = True
                            
                            strTmp = Mid(varAry(lngLoop), 2)

                            If Vsf.TextMatrix(Vsf.Rows - 1, mCol.项目缺省) = "" Then
                                Vsf.TextMatrix(Vsf.Rows - 1, mCol.项目缺省) = strTmp
                            Else
                                Vsf.TextMatrix(Vsf.Rows - 1, mCol.项目缺省) = Vsf.TextMatrix(Vsf.Rows - 1, mCol.项目缺省) & ";" & strTmp
                            End If
                        Else
                            strTmp = CStr(varAry(lngLoop))
                        End If

                        If Vsf.TextMatrix(Vsf.Rows - 1, mCol.项目值域) = "" Then
                            Vsf.TextMatrix(Vsf.Rows - 1, mCol.项目值域) = strTmp
                        Else
                            Vsf.TextMatrix(Vsf.Rows - 1, mCol.项目值域) = Vsf.TextMatrix(Vsf.Rows - 1, mCol.项目值域) & "|" & strTmp
                        End If
                    Next
                End If

                Vsf.TextMatrix(Vsf.Rows - 1, mCol.项目id) = zlCommFun.NVL(RS("项目id").Value)
'                If mbytMode = 1 Then
'                    vsf.TextMatrix(vsf.Rows - 1, mCol.记录结果) = vsf.TextMatrix(vsf.Rows - 1, mCol.项目缺省)
'                    vsf.TextMatrix(vsf.Rows - 1, mCol.是否变动) = "1"
'                Else
                Vsf.TextMatrix(Vsf.Rows - 1, mCol.记录结果) = ""
'                End If

                Vsf.TextMatrix(Vsf.Rows - 1, mCol.项目单位) = zlCommFun.NVL(RS("项目单位").Value)
                Vsf.RowData(Vsf.Rows - 1) = zlCommFun.NVL(RS("项目序号").Value, 0)
                
                RS.MoveNext
            Loop
        End If
        
'    Else
        '20090914:每次批量新增记录,日期都恢复为缺省值,所以屏蔽
        'dtp.Value = Format(mrsParam("时间").Value, dtp.CustomFormat)
'        dtp.Enabled = False
'        cboBaby.Enabled = False
'    End If
    
'    Call ReadData
    OpenPatientMap = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function ReadDrink() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------------------
    Dim RS As ADODB.Recordset
    Dim strSQL As String
    Dim intLoop As Integer
    Dim strTmp As String
    Dim strStart As String
    Dim strEnd As String
    Dim int饮入物 As Integer
    Dim int饮入量 As Integer
    Dim strValue As String
    
    
    On Error GoTo ErrHand
    
    strStart = Format(dtp.Value, "yyyy-MM-dd HH:mm") & ":00"
    strEnd = Format(DateAdd("n", 1, CDate(strStart)), "yyyy-MM-dd HH:mm") & ":00"
    
    For intLoop = 1 To Vsf.Rows - 1
        If Val(Vsf.RowData(intLoop)) = 6 Then
            int饮入物 = intLoop
        End If
        If Val(Vsf.RowData(intLoop)) = 7 Then
            int饮入量 = intLoop
        End If
    Next
    
    If int饮入物 = 0 And int饮入量 = 0 Then Exit Function
    
    strSQL = "Select zl_PatitDrink([1],[2],[3],[4]) As 饮入 From Dual"
    
    Set RS = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mrsParam("病人id")), Val(mrsParam("主页id")), CDate(strStart), CDate(strEnd))
    If RS.BOF = False Then
        
        strTmp = zlCommFun.NVL(RS("饮入"))
        
        If strTmp <> "" Then
            
            strValue = Trim(Split(strTmp, ";")(0))
            If int饮入量 > 0 Then Vsf.TextMatrix(int饮入量, mCol.记录结果) = strValue
            
            strValue = Trim(Split(strTmp, ";")(1))
            If UBound(Split(strTmp, ";")) > 1 Then strValue = strValue & "等"
            
            If int饮入物 > 0 Then Vsf.TextMatrix(int饮入物, mCol.记录结果) = strValue
            
            
        End If
    End If
        
    ReadDrink = True
                            
    Exit Function
    
ErrHand:

    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '读取指定记录id、指定组的护理内容，注：都要调用此，没有内容时，要产生项目列表
    '------------------------------------------------------------------------------------------------------------------
    Dim RS As New ADODB.Recordset
    Dim varAry As Variant
    Dim lngLoop As Long
    Dim strTmp As String
    Dim lngColor As Long
    Dim strStart As String
    Dim strEnd As String
    Dim blnAllow As Boolean
    
    On Error GoTo ErrHand
        
    strStart = Format(dtp.Value, "yyyy-MM-dd HH:mm:ss")
    strEnd = Format(DateAdd("n", 1, CDate(strStart)), "yyyy-MM-dd HH:mm:ss")
    
    mint饮入量 = 0
    
    Vsf.Rows = 2
    Vsf.RowData(1) = 0
    Vsf.Cell(flexcpText, 1, 1, 1, Vsf.Cols - 1) = ""
    
    '------------------------------------------------------------------------------------------------------------------
    mstrSQL = "Select X. *, " & _
                     "Y.项目序号, " & _
                     "Y.项目名称, " & _
                     "Y.项目单位, " & _
                     "Y.分组名, " & _
                     "Y.项目表示, " & _
                     "Y.项目值域, " & _
                     "Y.项目类型, " & _
                     "Y.项目长度, " & _
                     "Y.项目小数, " & _
                     "Y.项目id,Y.保留项目,Y.分组名,Y.项目性质 " & _
                "From "
    
    If mint心率应用 = 2 Then
        mstrSQL = mstrSQL & _
                    "(Select A.记录内容 As 记录结果, " & _
                                 "C.保存人 As 记录人, " & _
                                 "C.保存时间 As 记录时间,Decode(a.记录内容,Null,'',A.体温部位) As 部位,b.记录内容 As 标记,b.记录标记," & _
                                 "A.项目序号, " & _
                                 "C.发生时间 As 完成日期,A.记录id,a.未记说明 " & _
                             "From 病人护理内容 A, 病人护理内容 B,病人护理记录 C " & _
                            "Where C.ID = A.记录id And b.记录id(+)=a.记录id And b.记录组号(+)=a.记录组号 And b.记录标记(+) =1 " & _
                                  "AND A.记录类型 = 1 " & _
                                  "AND C.病人来源 = 2 " & _
                                  "AND NVL(A.记录标记,0) <> 1 " & _
                                  "AND C.ID = [1] And A.记录组号=[5] "
    Else
    
        mstrSQL = mstrSQL & _
                    "(Select A.记录内容 As 记录结果, " & _
                                 "C.保存人 As 记录人, " & _
                                 "C.保存时间 As 记录时间,Decode(a.记录内容,Null,'',A.体温部位) As 部位,Decode(a.项目序号,2,'',-1,'',b.记录内容) As 标记,Decode(a.项目序号,2,'',-1,'',b.记录标记) As 记录标记," & _
                                 "A.项目序号, " & _
                                 "C.发生时间 As 完成日期,A.记录id,a.未记说明 " & _
                             "From 病人护理内容 A, 病人护理内容 B,病人护理记录 C " & _
                            "Where C.ID = A.记录id And b.记录id(+)=a.记录id And b.记录组号(+)=a.记录组号 And b.记录标记(+) =1 " & _
                                  "AND A.记录类型 = 1 " & _
                                  "AND C.病人来源 = 2 " & _
                                  "AND ((NVL(A.记录标记,0) <> 1 And a.项目序号>0) or a.项目序号=-1) " & _
                                  "AND C.ID = [1] And A.记录组号=[5] "
                                  
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    If Val(mrsParam("版本").Value) = 0 Then
    
        mstrSQL = mstrSQL & _
                    " And a.终止版本 Is Null And b.终止版本 Is Null "
                    
    Else
                
        mstrSQL = mstrSQL & _
                    " And Nvl(a.开始版本,1)<=[4] And Nvl(a.终止版本,10000)>[4] And Nvl(b.开始版本,1)<=[4] And Nvl(b.终止版本,10000)>[4] "
    End If

    '------------------------------------------------------------------------------------------------------------------
    mstrSQL = mstrSQL & _
                        " and Decode(a.项目序号,2,-1,a.项目序号)=b.项目序号(+)) X, " & _
                      "护理记录项目 Y " & _
                "Where Y.项目序号 = X.项目序号(+) And Nvl(y.应用方式,0)=1 And Nvl(y.适用病人,0) In (0,[6]) And (Y.适用科室=1 Or (Y.适用科室=2 And Exists (Select 1 From 护理适用科室 D Where D.项目序号=Y.项目序号 And D.科室id=[2])))  " & _
                        "AND Y.护理等级 >=[3]  " & _
                "Order By Y.分组名,Y.项目序号,X.记录标记 "
                
    'Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, Val(tbs.Tag), Val(mrsParam("科室id").Value), Val(mrsParam("护理等级").Value), Val(mrsParam("版本").Value), Val(tbs.SelectedItem.Tag), IIf(Val(mrsParam("婴儿").Value) = 0, 1, 2))
    Set RS = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, Val(tbcPage.Selected.Tag), Val(mrsParam("科室id").Value), Val(mrsParam("护理等级").Value), Val(mrsParam("版本").Value), Val(tbs.SelectedItem.Tag), IIf(Val(mrsParam("婴儿").Value) = 0, 1, 2))
    If RS.BOF = False Then
        
'        mrsParam("记录id").Value = Val(tbs.Tag)
        mrsParam("记录id").Value = Val(tbcPage.Selected.Tag)
        
        With Vsf
            Do While Not RS.EOF
                
                blnAllow = False
                If zlCommFun.NVL(RS("项目性质"), 1) = 2 Then
                    If zlCommFun.NVL(RS("记录结果")) <> "" Then
                        blnAllow = True
                    End If
                Else
                    blnAllow = True
                End If
                
                If blnAllow Then
                    If Val(.RowData(.Rows - 1)) <> 0 Then .Rows = .Rows + 1
                    
                    Call WriteItemData(RS, .Rows - 1)

                End If
                
                RS.MoveNext
            Loop
        
            Call .Body.AutoSize(mCol.记录结果, mCol.记录结果)
        End With
        
        Call ExecuteCommand("历史数据", Format(dtp.Value, "yyyy-MM-dd HH:mm:ss"), Val(Vsf.RowData(Vsf.Row)))
    End If

    ReadData = True
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function WriteItem(ByVal rsData As ADODB.Recordset, ByVal intRow As Integer) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngColor As Long
    Dim varAry As Variant
    Dim lngLoop As Long
    Dim strTmp As String
    
    mblnDefault = False
    With Vsf
        .RowData(intRow) = zlCommFun.NVL(rsData("项目序号"))
        
        .TextMatrix(intRow, mCol.记录组) = zlCommFun.NVL(rsData("分组名").Value)
        .TextMatrix(intRow, mCol.护理项目) = zlCommFun.NVL(rsData("项目名称"))
        .TextMatrix(intRow, mCol.项目类型) = zlCommFun.NVL(rsData("项目类型"))
        .TextMatrix(intRow, mCol.项目长度) = zlCommFun.NVL(rsData("项目长度"))
        .TextMatrix(intRow, mCol.项目小数) = zlCommFun.NVL(rsData("项目小数"), 0)
        .TextMatrix(intRow, mCol.项目表示) = zlCommFun.NVL(rsData("项目表示"))
        .TextMatrix(intRow, mCol.项目性质) = zlCommFun.NVL(rsData("项目性质"), 1)
        .TextMatrix(intRow, mCol.项目id) = zlCommFun.NVL(rsData("项目id"), 0)
                            
        If zlCommFun.NVL(rsData("项目值域")) <> "" Then
        
            varAry = Split(zlCommFun.NVL(rsData("项目值域")), ";")
                            
            For lngLoop = 0 To UBound(varAry)
                If Left(varAry(lngLoop), 1) = "√" Then
                    mblnDefault = True
                    strTmp = Mid(varAry(lngLoop), 2)
                    
                    If .TextMatrix(intRow, mCol.项目缺省) = "" Then
                        .TextMatrix(intRow, mCol.项目缺省) = strTmp
                    Else
                        .TextMatrix(intRow, mCol.项目缺省) = .TextMatrix(intRow, mCol.项目缺省) & ";" & strTmp
                    End If
                Else
                    strTmp = CStr(varAry(lngLoop))
                End If
                                    
                If .TextMatrix(intRow, mCol.项目值域) = "" Then
                    .TextMatrix(intRow, mCol.项目值域) = strTmp
                Else
                    .TextMatrix(intRow, mCol.项目值域) = .TextMatrix(intRow, mCol.项目值域) & "|" & strTmp
                End If
            Next
        End If
        
        If zlCommFun.NVL(rsData("项目序号")) = 7 And zlCommFun.NVL(rsData("保留项目"), 0) = 1 Then mint饮入量 = intRow
        .TextMatrix(intRow, mCol.项目单位) = zlCommFun.NVL(rsData("项目单位").Value)
        
    End With
    
    WriteItem = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function


Private Function WriteItemData(ByVal rsData As ADODB.Recordset, ByVal intRow As Integer) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngColor As Long
    Dim varAry As Variant
    Dim lngLoop As Long
    Dim strTmp As String
    
    With Vsf
        .RowData(intRow) = zlCommFun.NVL(rsData("项目序号"))
        
        Call WriteItem(rsData, intRow)
        
        If zlCommFun.NVL(rsData("记录结果")) <> "" Or zlCommFun.NVL(rsData("未记说明")) <> "" Then
            
            .TextMatrix(intRow, mCol.部位) = zlCommFun.NVL(rsData("部位"))
            
            Select Case zlCommFun.NVL(rsData("项目序号"))
            Case 9
                If Right(zlCommFun.NVL(rsData("记录结果")), 2) = "/C" Then
                
                    .TextMatrix(intRow, mCol.记录结果) = Left(zlCommFun.NVL(rsData("记录结果")), Len(zlCommFun.NVL(rsData("记录结果"))) - 2)
                    .TextMatrix(intRow, mCol.标记) = "/C"
                    
                ElseIf Right(zlCommFun.NVL(rsData("记录结果")), 1) = "C" Then
                    .TextMatrix(intRow, mCol.记录结果) = Left(zlCommFun.NVL(rsData("记录结果")), Len(zlCommFun.NVL(rsData("记录结果"))) - 1)
                    .TextMatrix(intRow, mCol.标记) = "C"
                Else
                    .TextMatrix(intRow, mCol.记录结果) = zlCommFun.NVL(rsData("记录结果"))
                    .TextMatrix(intRow, mCol.标记) = zlCommFun.NVL(rsData("标记"))
                End If
            Case 10
                If Right(zlCommFun.NVL(rsData("记录结果")), 2) = "/E" Then
                    .TextMatrix(intRow, mCol.记录结果) = Left(zlCommFun.NVL(rsData("记录结果")), Len(zlCommFun.NVL(rsData("记录结果"))) - 2)
                    .TextMatrix(intRow, mCol.标记) = "/E"
                ElseIf Right(zlCommFun.NVL(rsData("记录结果")), 1) = "E" Then
                    .TextMatrix(intRow, mCol.记录结果) = Left(zlCommFun.NVL(rsData("记录结果")), Len(zlCommFun.NVL(rsData("记录结果"))) - 1)
                    .TextMatrix(intRow, mCol.标记) = "E"
                ElseIf Right(zlCommFun.NVL(rsData("记录结果")), 1) = "*" Then
                    .TextMatrix(intRow, mCol.标记) = "*"
                Else
                    .TextMatrix(intRow, mCol.记录结果) = zlCommFun.NVL(rsData("记录结果").Value)
                    .TextMatrix(intRow, mCol.标记) = zlCommFun.NVL(rsData("标记").Value)
                End If
            Case Else
            
                lngColor = GridTextColor(zlCommFun.NVL(rsData("项目名称")), zlCommFun.NVL(rsData("记录结果").Value))
                .Cell(flexcpForeColor, intRow, mCol.记录结果, intRow, mCol.记录结果) = lngColor
                
                .TextMatrix(intRow, mCol.记录结果) = zlCommFun.NVL(rsData("记录结果").Value)
                .TextMatrix(intRow, mCol.标记) = zlCommFun.NVL(rsData("标记").Value)
                .TextMatrix(intRow, mCol.未记说明) = zlCommFun.NVL(rsData("未记说明").Value)
            End Select
        End If
            
    End With
    
    WriteItemData = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function SignName() As Boolean
    Dim RS As New ADODB.Recordset
    Dim oSign As cEPRSign
    Dim strSource As String
    Dim strSQL() As String
    Dim blnTran As Boolean
    Dim strDate As String
    Dim strStart As String
    Dim lngLoop As Long
    
    On Error GoTo ErrHand
    
    '初始处理
    '------------------------------------------------------------------------------------------------------------------
    ReDim Preserve strSQL(1 To 1)
    
    strDate = Format(dtp.Value, "yyyy-MM-dd HH:mm")
    strStart = strDate & ":00"
    strSource = ""
    
    '检查当前是否已经签名了
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select 1 From 病人护理内容 a,病人护理记录 b Where b.病人id=[1] And b.主页id=[2] And b.发生时间=[3] And Nvl(b.婴儿,0)=[4] And a.记录id=b.ID And a.记录类型=5 And Nvl(a.开始版本,1)=Nvl(b.最后版本,1)"
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("病人id").Value), Val(mrsParam("主页id").Value), CDate(strStart), Val(mrsParam("婴儿").Value), Val(mrsParam("版本").Value))
    If RS.BOF = False Then
        ShowSimpleMsg "当前没有需要签名的信息！"
        Exit Function
    End If
        
    '获取要签名的内容
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select a.记录类型,a.项目分组,a.项目序号,a.项目名称,a.项目类型,a.记录内容,a.项目单位,a.记录标记,a.体温部位,a.记录组号,a.复试合格,a.未记说明,a.记录人,a.修改时间" & vbNewLine & _
             " From 病人护理内容 a,病人护理记录 b " & vbNewLine & _
             " Where b.病人id=[1] And b.主页id=[2] And b.发生时间=[3] And Nvl(b.婴儿,0)=[4] And a.记录id=b.ID And a.终止版本 Is Null" & vbNewLine & _
             " Order by A.项目序号"
    If mblnMoved_HL Then
        gstrSQL = Replace(gstrSQL, "病人护理记录", "H病人护理记录")
        gstrSQL = Replace(gstrSQL, "病人护理内容", "H病人护理内容")
    End If
    
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("病人id").Value), Val(mrsParam("主页id").Value), CDate(strStart), Val(mrsParam("婴儿").Value))
    If RS.BOF = False Then
        Do While Not RS.EOF
            For lngLoop = 0 To RS.Fields.Count - 1
                strSource = strSource & CStr(zlCommFun.NVL(RS.Fields(lngLoop).Value, ""))
            Next
            RS.MoveNext
        Loop
    End If
    Debug.Print "签名：" & Now & vbCrLf & strSource
    If strSource = "" Then
        MsgBox "当前没有需要签名的信息！", vbInformation, gstrSysName
        Exit Function
    End If
    '76223:刘鹏飞,2014-08-05,电子签名添加时间戳信息
    '------------------------------------------------------------------------------------------------------------------
    Set oSign = frmCaseTendSign.ShowMe(Me, mstrPrivs, strSource, Val(mrsParam("病人id").Value), Val(mrsParam("主页id").Value), Val(mrsParam("病区id").Value))
    If Not oSign Is Nothing Then

        mstrSQL = "ZL_电子护理记录_SignName("
        mstrSQL = mstrSQL & Val(mrsParam("病人id")) & ","
        mstrSQL = mstrSQL & Val(mrsParam("主页id")) & ","
        mstrSQL = mstrSQL & Val(mrsParam("婴儿")) & ","
        mstrSQL = mstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
        mstrSQL = mstrSQL & "'" & oSign.姓名 & "',"
        mstrSQL = mstrSQL & "'" & oSign.签名信息 & "',"
        mstrSQL = mstrSQL & oSign.证书ID & ","
        mstrSQL = mstrSQL & oSign.签名方式 & ",'" & oSign.时间戳 & "','" & oSign.时间戳信息 & "')"

        strSQL(ReDimArray(strSQL)) = mstrSQL

        '执行
        '--------------------------------------------------------------------------------------------------------------
        blnTran = True
        gcnOracle.BeginTrans
        For lngLoop = 1 To UBound(strSQL)
            If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
        Next
        gcnOracle.CommitTrans
        blnTran = False
        
        SignName = True
    End If
    
    Exit Function
    
    '出错处理
    '------------------------------------------------------------------------------------------------------------------
ErrHand:

    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    

End Function

Private Function UnSignName() As Boolean
    '******************************************************************************************************************
    '功能:
    '
    '
    '******************************************************************************************************************
    Dim strSource As String
    Dim strSQL() As String
    Dim blnTran As Boolean
    Dim strDate As String
    Dim strStart As String
    Dim lngLoop As Long
    Dim RS As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    '初始处理
    '------------------------------------------------------------------------------------------------------------------
    ReDim Preserve strSQL(1 To 1)
    strDate = Format(dtp.Value, "yyyy-MM-dd HH:mm")
    strStart = strDate & ":00"
    
    '检查当前是否已经签名了
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select 1 From 病人护理内容 a,病人护理记录 b Where b.病人id=[1] And b.主页id=[2] And b.发生时间=[3] And Nvl(b.婴儿,0)=[4] And a.记录id=b.ID And a.记录类型=5"
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("病人id").Value), Val(mrsParam("主页id").Value), CDate(strStart), Val(mrsParam("婴儿").Value), Val(mrsParam("版本").Value))
    If RS.BOF Then
        ShowSimpleMsg "当前没有需要取消的签名！"
        Exit Function
    End If
    
    
    '如果是电子签名,则需要验证
    '------------------------------------------------------------------------------------------------------------------
    If Val(Me.Tag) > 0 Then
        '数字签名验证
        Err.Clear
        If gobjTendESign Is Nothing Then
            On Error Resume Next
            Set gobjTendESign = CreateObject("zl9ESign.clsESign")
            If Err <> 0 Then Err.Clear
            On Error GoTo 0
            If Not gobjTendESign Is Nothing Then Call gobjTendESign.Initialize(gcnOracle, glngSys)
        End If
        If Not gobjTendESign Is Nothing Then
            If Not gobjTendESign.CheckCertificate(gstrDBUser) Then Exit Function
        Else
            MsgBox "电子签名部件未能正确安装，回退操作不能继续！", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
    End If

    '------------------------------------------------------------------------------------------------------------------
    mstrSQL = "Zl_电子护理记录_Unsignname("
    mstrSQL = mstrSQL & Val(mrsParam("病人id")) & ","
    mstrSQL = mstrSQL & Val(mrsParam("主页id")) & ","
    mstrSQL = mstrSQL & Val(mrsParam("婴儿")) & ","
    mstrSQL = mstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'))"
    strSQL(ReDimArray(strSQL)) = mstrSQL

    '执行
    '------------------------------------------------------------------------------------------------------------------
    blnTran = True
    gcnOracle.BeginTrans
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    gcnOracle.CommitTrans
    blnTran = False
    
    UnSignName = True
    
    Exit Function
    
    '出错处理
    '------------------------------------------------------------------------------------------------------------------
ErrHand:

    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    

End Function

Private Function SaveDataAll() As Boolean
    Dim RS As New ADODB.Recordset
    Dim blnTrans As Boolean
    Dim lngLoop As Long
    Dim strSQL() As String
    Dim intCol As Integer
    Dim intRow As Integer
    
    On Error GoTo ErrHand
        
    ReDim Preserve strSQL(1 To 1)
    
    If SaveData(strSQL) = False Then GoTo EndHand
    
    
    intCol = tbs.SelectedItem.Index
    
'    For intRow = 1 To tbs.Tabs.Count
'        If intRow <> intCol Then
'            tbs.Tabs(intRow).Selected = True
'
'            If SaveData(strSQL) = False Then GoTo errHand

'        End If
'    Next

    blnTrans = True
    gcnOracle.BeginTrans
    
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    
    gcnOracle.CommitTrans
    blnTrans = False
    
    If mbytMode = 1 Or tbcPage.Selected.Tag = "" Then
    
        gstrSQL = "Select a.ID,b.记录组号 From 病人护理记录 a,病人护理内容 b Where a.ID=b.记录id And a.病人id=[1] And a.主页id=[2] And a.发生时间=[3] And Nvl(a.婴儿,0)=[4] And b.记录类型<>5 Group By a.id,b.记录组号 Order By a.id,b.记录组号"
        Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("病人id")), Val(mrsParam("主页id")), CDate(Format(dtp.Value, "yyyy-MM-dd HH:mm:ss")), Val(mrsParam("婴儿")))
        If RS.BOF = False Then
'            tbs.Tag = Val(rs("ID").Value)
            tbcPage.Selected.Tag = Val(RS("ID").Value)
        End If
        
    End If
    
    tbs.Tabs(intCol).Selected = True
    
    SaveDataAll = True
    
    Exit Function
    
ErrHand:
    '出错处理
    
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    
EndHand:
    
End Function

Private Function SaveData(ByRef strSQL() As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strDate As String
    Dim strStart As String
    Dim strEnd As String
    Dim lng科室ID As Long
    Dim RS As New ADODB.Recordset
    Dim strTmp As String
    Dim intAllow As Integer
    
    On Error GoTo ErrHand
            
    mstrSQL = " Select D.ID,D.名称,开始,终止" & _
            " From 部门表 D," & _
            "   (Select 科室id,To_Date(To_Char(Min(开始时间), 'yyyy-mm-dd hh24:mi'), 'yyyy-mm-dd hh24:mi') as 开始,Max(Nvl(终止时间,Sysdate+100)) as 终止" & _
            "    From 病人变动记录" & _
            "    Where 开始时间 is Not Null And 病人ID=[1] And 主页ID=[2]" & _
            "    Group by 科室id) L" & _
            " Where L.科室id=D.ID"
    mstrSQL = mstrSQL & " And To_Date('" & Format(dtp.Value, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss') between 开始 and 终止 "
    
    Set RS = zlDatabase.OpenSQLRecord(mstrSQL, gstrSysName, Val(mrsParam("病人id")), Val(mrsParam("主页id")))
    If RS.BOF = False Then
        lng科室ID = RS("ID").Value
    Else
        ShowSimpleMsg "发生时间不能大于当前时间或不能小于开始时间！"
        Exit Function
    End If
    
    strDate = Format(dtp.Value, "yyyy-MM-dd HH:mm")
    strStart = strDate & ":00"
    strEnd = Format(DateAdd("n", 1, CDate(strDate)), "yyyy-MM-dd HH:mm") & ":00"
    intAllow = IIf(InStr(mstrPrivs, "他人护理记录") > 0, 1, 0)
    
    '数据发生时间不能在当前操作员所属科室的有效时间以前
    If Not CheckTime(Val(mrsParam("病人id")), Val(mrsParam("主页id")), Mid(strStart, 1, 16), Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")) Then Exit Function
    
    Dim int记录组号 As Long
    int记录组号 = Val(tbs.SelectedItem.Tag)
    
    If Val(mrsParam("记录id").Value) > 0 Then
        
        mstrSQL = "Select ID From 病人护理记录 Where 病人id = [1] And 主页id = [2] And Nvl(婴儿, 0) = Nvl([3], 0) And 病人来源 = 2 And 发生时间 = [4] And ID<>[5]"
        Set RS = zlDatabase.OpenSQLRecord(mstrSQL, gstrSysName, Val(mrsParam("病人id")), Val(mrsParam("主页id")), Val(mrsParam("婴儿")), CDate(strStart), Val(mrsParam("记录id").Value))
        If RS.BOF = False Then
            If MsgBox("当前发生时间还存在其他的记录，是否覆盖？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                Exit Function
            End If
            
            '删除相同id,并更新新id记录的发生时间和婴儿标记
            mstrSQL = "Zl_病人护理记录_UpdateReplace(" & Val(mrsParam("记录id").Value) & "," & Val(RS("ID").Value) & "," & Val(mrsParam("婴儿").Value) & ",To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'))"
            strSQL(ReDimArray(strSQL)) = mstrSQL
        Else
            mstrSQL = "Zl_病人护理记录_UpdateReplace(" & Val(mrsParam("记录id").Value) & ",0," & Val(mrsParam("婴儿").Value) & ",To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'))"
            strSQL(ReDimArray(strSQL)) = mstrSQL
        End If
    Else
        mstrSQL = "Select ID From 病人护理记录 Where 病人id = [1] And 主页id = [2] And Nvl(婴儿, 0) = Nvl([3], 0) And 病人来源 = 2 And 发生时间 = [4]"
        Set RS = zlDatabase.OpenSQLRecord(mstrSQL, gstrSysName, Val(mrsParam("病人id")), Val(mrsParam("主页id")), Val(mrsParam("婴儿")), CDate(strStart))
        If RS.BOF = False Then
            If MsgBox("当前发生时间还存在其他的记录，是否覆盖？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                Exit Function
            End If
            
            '删除相同id,并更新新id记录的发生时间和婴儿标记
            mstrSQL = "Zl_病人护理记录_UpdateReplace(" & Val(mrsParam("记录id").Value) & "," & Val(RS("ID").Value) & "," & Val(mrsParam("婴儿").Value) & ",To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'))"
            strSQL(ReDimArray(strSQL)) = mstrSQL
        End If
    End If
    
    For lngLoop = 1 To Vsf.Rows - 1
        
        If Val(Vsf.RowData(lngLoop)) <> 0 And Val(Vsf.TextMatrix(lngLoop, mCol.是否变动)) = 1 Then
'        If Val(vsf.RowData(lngLoop)) <> 0 Then

            mstrSQL = "Zl_病人护理记录_UpdateRecord("
            mstrSQL = mstrSQL & Val(mrsParam("病人id")) & ","
            mstrSQL = mstrSQL & Val(mrsParam("主页id")) & ","
            mstrSQL = mstrSQL & Val(mrsParam("婴儿")) & ","
            mstrSQL = mstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
            mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
            mstrSQL = mstrSQL & "1,"
            mstrSQL = mstrSQL & Val(Vsf.RowData(lngLoop)) & ","
            
            If Val(Vsf.RowData(lngLoop)) = -1 Then
                mstrSQL = mstrSQL & "1,"
            Else
                mstrSQL = mstrSQL & "0,"
            End If
            
            Select Case Val(Vsf.RowData(lngLoop))
            Case 9, 10
                strTmp = Trim(Vsf.TextMatrix(lngLoop, mCol.标记))
            Case Else
                strTmp = ""
            End Select
            
            If Trim(Vsf.TextMatrix(lngLoop, mCol.记录结果)) <> "" Then strTmp = Vsf.TextMatrix(lngLoop, mCol.记录结果) & strTmp
            mstrSQL = mstrSQL & "'" & strTmp & "',"
            mstrSQL = mstrSQL & "'" & Trim(Vsf.TextMatrix(lngLoop, mCol.部位)) & "'," & intAllow & "," & IIf(IsNumeric(strTmp), 0, 1) & "," & int记录组号 & ",'" & Vsf.TextMatrix(lngLoop, mCol.未记说明) & "')"
                
            strSQL(ReDimArray(strSQL)) = mstrSQL
            
            
            If (Val(Vsf.RowData(lngLoop)) = 1 Or (Val(Vsf.RowData(lngLoop)) = 2) And mint心率应用 = 2) Then

                mstrSQL = "Zl_病人护理记录_UpdateRecord("
                mstrSQL = mstrSQL & Val(mrsParam("病人id")) & ","
                mstrSQL = mstrSQL & Val(mrsParam("主页id")) & ","
                mstrSQL = mstrSQL & Val(mrsParam("婴儿")) & ","
                mstrSQL = mstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
                mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                mstrSQL = mstrSQL & "1,"
                mstrSQL = mstrSQL & IIf(Val(Vsf.RowData(lngLoop)) = 2, -1, Val(Vsf.RowData(lngLoop))) & ","
                mstrSQL = mstrSQL & "1,"
                                                
                If Trim(Vsf.TextMatrix(lngLoop, mCol.标记)) <> "" And Trim(Vsf.TextMatrix(lngLoop, mCol.记录结果)) <> "" Then
                    Select Case Val(Vsf.TextMatrix(lngLoop, mCol.项目类型))
                    Case 0          '数值
                        strTmp = Val(Trim(Vsf.TextMatrix(lngLoop, mCol.标记)))
                    Case 1          '文本
                        strTmp = Trim(Trim(Vsf.TextMatrix(lngLoop, mCol.标记)))
                    End Select
                    
                    mstrSQL = mstrSQL & "'" & strTmp & "',"
                    mstrSQL = mstrSQL & "NULL," & intAllow & "," & IIf(IsNumeric(strTmp), 0, 1) & "," & int记录组号 & ",Null)"
                Else
                    mstrSQL = mstrSQL & "NULL,"
                    mstrSQL = mstrSQL & "NULL," & intAllow & ",0," & int记录组号 & ",Null)"
                End If
                
                strSQL(ReDimArray(strSQL)) = mstrSQL
            End If
            
        End If
    Next
                
    
    Vsf.Cell(flexcpText, 1, mCol.是否变动, Vsf.Rows - 1, mCol.是否变动) = ""
    
    SaveData = True
    
    Exit Function
    
ErrHand:
    '出错处理
    
    If ErrCenter = 1 Then
        Resume
    End If

    
End Function

Private Function ExecuteCommand(ByVal strCmd As String, ParamArray varParam() As Variant) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '--------------------------------------------------------------------------------------------------------------
    Dim intLoop As Integer
    Dim RS As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    Dim lngRow As Long
    Dim strStart As String
    Dim strEnd As String
    Dim curDate As Date
    Dim intDay As Integer
    Dim strPar As String
    
    On Error GoTo ErrHand


    Select Case strCmd
    '------------------------------------------------------------------------------------------------------------------
    Case "初始控件"
        
        If InitData = False Then Exit Function
        Call InitMenuBar
        
    '------------------------------------------------------------------------------------------------------------------
    Case "初始数据"
    
        If mrsParam("时间").Value = "" Then mrsParam("时间").Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
        dtp.Value = Format(mrsParam("时间").Value, dtp.CustomFormat)
        txtDay.Text = Val(zlDatabase.GetPara("历史数据天数", glngSys, 1255, "1"))
                
        '出院开始日期;出院结束日期;在院病人;出院病人;转出病人;转出天数
        '------------------------------------------------------------------------------------------------------------------
        strPar = zlDatabase.GetPara("病人显示范围", glngSys, 1262, "10000")
        mrsParam("在院病人").Value = Val(Mid(strPar, 1, 1))
        mrsParam("出院病人").Value = Val(Mid(strPar, 2, 1))
        mrsParam("转出病人").Value = Val(Mid(strPar, 4, 1))
        On Error Resume Next
        mrsParam("待入科病人").Value = Val(Mid(strPar, 5, 1))
        On Error GoTo 0
        mrsParam("转出天数").Value = Val(zlDatabase.GetPara("最近转出天数", glngSys, 1262, 7))
        
        curDate = zlDatabase.Currentdate
        intDay = Val(zlDatabase.GetPara("出院病人结束间隔", glngSys, 1262, 7))
        mrsParam("出院结束日期").Value = Format(curDate + intDay, "yyyy-MM-dd 23:59:59")
        intDay = Val(zlDatabase.GetPara("出院病人开始间隔", glngSys, 1262, 30))
        mrsParam("出院开始日期").Value = Format(CDate(mrsParam("出院结束日期").Value) - intDay, "yyyy-MM-dd 00:00:00")
        
        '------------------------------------------------------------------------------------------------------------------
        '88776:就诊卡长度获取有参数调整为数据表
        gstrSQL = "Select NVL(卡号长度,8) 卡号长度 From 医疗卡类别 Where 特定项目 = '就诊卡'"
        Call zlDatabase.OpenRecordset(RS, gstrSQL, Me.Caption)
        If RS.EOF = False Then
            mint就诊卡号码长度 = Val("" & RS!卡号长度)
        Else
            mint就诊卡号码长度 = 8
        End If
        
        mstr就诊卡字母前缀 = UCase(zlDatabase.GetPara(27, glngSys))
        
        '------------------------------------------------------------------------------------------------------------------
        gstrSQL = "Select 出院科室ID from 病案主页 Where 病人id=[1] And 主页id=[2] "
        Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("病人id").Value), Val(mrsParam("主页id").Value))
        If RS.BOF = False Then
            mrsParam("科室id").Value = Val(zlCommFun.NVL(RS("出院科室ID").Value))
        End If
        
        mint心率应用 = 2
        gstrSQL = "Select a.应用方式 From 护理记录项目 a Where a.项目序号=-1"
        Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If RS.BOF = False Then
            mint心率应用 = zlCommFun.NVL(RS("应用方式").Value, 2)
        End If
        
        '读取病人列表，并保存到本地记录集中，以方便后续的选择
        If ReadPatient = False Then Exit Function
        
        '定位到当前病人
        mrsPatient.Filter = ""
        mrsPatient.Filter = "病人id=" & Val(mrsParam("病人id").Value)
        If mrsPatient.RecordCount > 0 Then mlngRowNum = Val(mrsPatient("ID").Value)
        mrsPatient.Filter = ""
    
    '------------------------------------------------------------------------------------------------------------------
    Case "清除数据"
        
        tbs.Tabs.Clear
        tbs.Tabs.Add 1, , "1"
        tbs.Tabs(1).Tag = 1
        cbo.Clear
        
    '------------------------------------------------------------------------------------------------------------------
    Case "刷新基本信息"
        
        Call OpenPatientMap(Val(mrsParam("病人id").Value), Val(mrsParam("主页id").Value), Val(mrsParam("婴儿").Value))
        
    '------------------------------------------------------------------------------------------------------------------
    Case "读取记录"
        '根据时间读取ID,进入窗体时
        
        tbcPage.Item(0).Tag = ""
        tbcPage.Item(0).Selected = True
        For intLoop = tbcPage.ItemCount - 1 To 1 Step -1
            tbcPage.RemoveItem intLoop
        Next
        
        gstrSQL = "Select b.ID From 病人护理记录 b Where b.病人id=[1] And b.主页id=[2] And Nvl(b.婴儿,0)=[3] And b.发生时间=[4]"
        Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("病人id").Value), Val(mrsParam("主页id").Value), Val(mrsParam("婴儿").Value), CDate(mrsParam("时间").Value & ":00"))
        If RS.BOF = False Then
            tbcPage.Item(0).Tag = RS("ID").Value
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "读取数据"
    
        '读取指定记录id的护理内容，包括组数据
        
        cbo.Clear
'        gstrSQL = "Select Distinct a.记录id,Nvl(a.开始版本,1) As 开始版本,c.记录人 As 签名人 From 病人护理内容 a,病人护理记录 b,病人护理内容 c Where a.记录类型<>5 And c.记录类型(+)=5 And c.记录id(+)=a.记录id And c.开始版本(+)=Nvl(a.开始版本,1) And a.记录id=b.ID And b.病人id=[1] And b.主页id=[2] And Nvl(b.婴儿,0)=[3] And b.发生时间=[4] Order By Nvl(a.开始版本,1) Desc"
'        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("病人id").Value), Val(mrsParam("主页id").Value), Val(mrsParam("婴儿").Value), CDate(Format(dtp.Value, "yyyy-MM-dd HH:mm:ss")))
'
        gstrSQL = "Select Distinct a.记录id,Nvl(a.开始版本,1) As 开始版本,c.记录人 As 签名人,b.发生时间 From 病人护理内容 a,病人护理记录 b,病人护理内容 c Where a.记录类型<>5 And c.记录类型(+)=5 And c.记录id(+)=a.记录id And c.开始版本(+)=Nvl(a.开始版本,1) And a.记录id=b.ID And b.ID=[1] Order By Nvl(a.开始版本,1) Desc"
        Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(tbcPage.Selected.Tag))
        
        If RS.BOF = False Then
            
            dtp.Value = Format(RS("发生时间").Value, "yyyy-MM-dd HH:mm")
            gstrSQL = "Select a.项目id As 证书id,Nvl(a.开始版本,1) As 开始版本 From 病人护理内容 a Where a.记录类型=5 And a.记录id=[1] Order By a.开始版本 Desc "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(RS("记录id").Value))
            If rsTmp.BOF = False Then
                If mbytMode = 4 Then Me.Caption = "取消“第 " & rsTmp("开始版本").Value & " 版”的签名。"
                Me.Tag = zlCommFun.NVL(rsTmp("证书id").Value, 0)
            End If
            
            Do While Not RS.EOF
                
                If zlCommFun.NVL(RS("签名人").Value, "") = "" Then
                    cbo.AddItem "第 " & RS("开始版本").Value & " 版"
                Else
                    cbo.AddItem "第 " & RS("开始版本").Value & " 版(已签名)"
                End If
                
                cbo.ItemData(cbo.NewIndex) = RS("开始版本").Value
                RS.MoveNext
            Loop
    
        End If
        If cbo.ListCount = 0 And mbytMode = 4 Then
            ShowSimpleMsg "目前还没有任何签名的版本！"
            Exit Function
        End If
        If cbo.ListCount > 0 And cbo.ListIndex = -1 Then cbo.ListIndex = 0
        
        '获取记录组
        '------------------------------------------------------------------------------------------------------------------
        tbs.Tabs.Clear
        intLoop = 0
        gstrSQL = "Select a.ID,b.记录组号 From 病人护理记录 a,病人护理内容 b Where a.ID=b.记录id And a.ID=[1] And b.记录类型<>5 Group By a.id,b.记录组号 Order By a.id,b.记录组号"
        Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(tbcPage.Selected.Tag))
        If RS.BOF = False Then
            Do While Not RS.EOF
                intLoop = intLoop + 1
                tbs.Tabs.Add intLoop, , CStr(intLoop)
                tbs.Tabs(intLoop).Tag = RS("记录组号").Value
                tbcPage.Selected.Tag = RS("ID").Value
    '            tbs.Tag = rs("ID").Value
                RS.MoveNext
            Loop
        Else
            tbs.Tabs.Add 1, , "1"
            tbs.Tabs(1).Tag = 1
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "读取组别数据"
        
        Call ReadData
    
    '------------------------------------------------------------------------------------------------------------------
    Case "历史数据"
        
        lbl(12).Caption = "前       天的“" & Vsf.TextMatrix(Vsf.Row, mCol.护理项目) & "”历史记录结果"
        strStart = Format(DateAdd("d", 0 - Val(txtDay.Text), CDate(varParam(0))), "yyyy-MM-dd HH:mm:ss")
        strEnd = Format(DateAdd("n", -1, CDate(varParam(0))), "yyyy-MM-dd HH:mm:ss")
        
        mclsVsfHistory.ClearGrid
        
        '显示指定前N天的指定指标数据
        strSQL = _
            "Select a.发生时间 As 记录时间, b.记录内容 As 历史结果" & vbNewLine & _
            "From 病人护理记录 a, 病人护理内容 b" & vbNewLine & _
            "Where a.发生时间 Between [3] And [4] And a.病人id=[1] And a.主页id=[2] And a.Id = b.记录id And b.项目序号 = [5] And" & vbNewLine & _
            "           b.记录内容 Is Not Null" & vbNewLine & _
            "Order By a.发生时间 Desc"
        Set RS = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mrsParam("病人id").Value), Val(mrsParam("主页id").Value), CDate(strStart), CDate(strEnd), Val(varParam(1)))
        
        If RS.BOF = False Then
            Call mclsVsfHistory.LoadGrid(RS)
        End If
        
        vsfHistory.AutoSize 1, 1
        
    '------------------------------------------------------------------------------------------------------------------
    Case "校对数据"
    
    '------------------------------------------------------------------------------------------------------------------
    Case "保存数据"
    
    End Select
        
    ExecuteCommand = True
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cbo_Click()
    Dim lng版本 As Long
    
    If mblnStartUp Then Exit Sub
    If mblnReading Then Exit Sub

    lng版本 = cbo.ItemData(cbo.ListIndex)
    If Val(mrsParam("版本").Value) = lng版本 Then Exit Sub
    mrsParam("版本").Value = lng版本
    
    Call ReadData
    
End Sub

Private Sub cboBaby_Click()
    If mblnStartUp Then Exit Sub
    If Val(mrsParam("婴儿").Value) = cboBaby.ListIndex Then Exit Sub
    mrsParam("婴儿").Value = cboBaby.ListIndex
    
    '如果不是新增模式，就根据时间读取对应的记理记录id
    Call ExecuteCommand("读取记录")
    Call ExecuteCommand("清除数据")
    Call ExecuteCommand("读取数据")
    Call ExecuteCommand("读取组别数据")
    
    DataChanged = False
    
'    Call ReadData
    
'    If mbytMode = 1 Or mbytMode = 2 Then DataChanged = True
    
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim intCol As Integer
    Dim intRow As Integer
    Dim blnCancel As Boolean
    
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Audit
        
        Call ReadDrink
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewParent                 '新增护理记录
        
        tbcPage.InsertItem tbcPage.ItemCount, "   " & CStr(tbcPage.ItemCount + 1) & "   ", picPane(2).hWnd, 0
        
        Call ExecuteCommand("清除数据")
            
        tbcPage.Item(tbcPage.ItemCount - 1).Selected = True
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem                   '新增记录组别
        
        tbs.Tabs.Add tbs.Tabs.Count + 1, , CStr(tbs.Tabs.Count + 1)
        tbs.Tabs(tbs.Tabs.Count).Selected = True
        For intRow = 1 To tbs.Tabs.Count + 10
            If intRow <> tbs.Tabs(intRow).Tag Then
                tbs.Tabs(tbs.Tabs.Count).Tag = intRow
                Exit For
            End If
        Next
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Append                '添加项目
        
        Dim rsData As New ADODB.Recordset
        Dim rsTmp As New ADODB.Recordset
        Dim strNotItem As String
        Dim intLoop As Integer
        Dim strTmp As String
        
        strNotItem = ""
        For intLoop = 1 To Vsf.Rows - 1
            
            If Val(Vsf.TextMatrix(intLoop, mCol.项目性质)) = 2 Then
                strNotItem = strNotItem & "," & Val(Vsf.RowData(intLoop))
            End If
            
        Next
        If strNotItem <> "" Then strNotItem = Mid(strNotItem, 2)

        Set rsData = GetGridItem(Val(mrsParam("护理等级").Value), Val(mrsParam("科室id").Value), IIf(Val(mrsParam("婴儿").Value) = 0, 1, 2), 2, strNotItem, False)
        
        If rsData.BOF = False Then
            If ShowTxtSelDialog(Me, Nothing, "名称,1500,0,1;单位,900,0,0", Me.Name & "\护理项目选择", "请从下面选择一个护理项目。", rsData, rsTmp, 6000, 3000, , , 2, False) Then
                If rsTmp.BOF = False Then
                    
                    '先在合适的位置处添加一空行
                    strTmp = ""
                    
                    For intLoop = 1 To Vsf.Rows - 1
                        If strTmp = "" Then
                            If Vsf.TextMatrix(intLoop, mCol.记录组) = zlCommFun.NVL(rsTmp("分组名").Value) Then
                                strTmp = Vsf.TextMatrix(intLoop, mCol.记录组)
                            End If
                        ElseIf strTmp <> Vsf.TextMatrix(intLoop, mCol.记录组) Then
                            Exit For
                        End If
                    Next
                    '填写项目数据
                    If intLoop = Vsf.Rows Then
                        Vsf.Rows = Vsf.Rows + 1
                        intLoop = Vsf.Rows - 1
                    Else
                        Vsf.Body.AddItem "", intLoop
                    End If
                    If WriteItem(rsTmp, intLoop) Then
                        Call LocationGrid(Vsf, intLoop, Vsf.Col)
                        If mblnDefault Then DataChanged = mblnDefault
                    End If

                End If
            End If
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete                '删除项目
        
        With Vsf
            If Val(.TextMatrix(.Row, mCol.项目性质)) = 2 Then
                
                    
                '检查是否有数据，如果无数据时才允许删除
                '本次保存之前有数据以及当前界面上有数据，则称之为有数据
                
                If Trim(.TextMatrix(.Row, mCol.记录结果)) <> "" Or .TextMatrix(.Row, mCol.是否变动) = "1" Then
                    ShowSimpleMsg "对不起，你要删除表格行有数据或者以前有数据！"
                    Exit Sub
                End If
                
                If MsgBox("确实要删除当前的表格项目吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                
                .RemoveItem .Row
                
                Call vsf_AfterRowColChange(0, 0, .Row, .Col)
                
            End If
        End With
        
'        '判断是否只有一个项了
'        For intRow = 1 To vsf.Rows - 1
'            If intRow <> vsf.Row Then
'                If Val(vsf.RowData(intRow)) = Val(vsf.RowData(vsf.Row)) Then
'                    vsf.Body.RemoveItem vsf.Row
'                    DataChanged = True
'                    Exit For
'                End If
'            End If
'        Next
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Save
    
        Select Case mbytMode
        Case 1, 2
            
            blnCancel = False
            Call vsf_ValidateEdit(Vsf.Row, Vsf.Col, blnCancel)
            If blnCancel = False Then mblnOk = SaveDataAll
                        
            If mblnOk Then
                RaiseEvent AfterDataChanged
            End If
        End Select
        
        If mblnOk Then DataChanged = False
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Cancle
        
        Select Case mbytMode
        Case 1, 2
            
            Call OpenPatientMap(Val(mrsParam("病人id").Value), Val(mrsParam("主页id").Value), Val(mrsParam("婴儿").Value))
            Call tbcPage_SelectedChanged(tbcPage.Selected)
            
            DataChanged = False
        End Select
        
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Sign                  '签名
        
        If mbytMode = 3 Then
            mblnOk = SignName
            If mblnOk Then

                RaiseEvent AfterDataChanged

            
                DataChanged = False
                Unload Me
            End If
        End If

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Untread                 '取消签名
        If mbytMode = 4 Then
            mblnOk = UnSignName
            If mblnOk Then
                
                RaiseEvent AfterDataChanged
                
                DataChanged = False
                Unload Me
            End If
        End If
    
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_LocationItem
        mstrFindKey = Control.Parameter
        mobjFindKey.Caption = mstrFindKey
        cbsThis.RecalcLayout
            
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Forward
        
        If mlngRowNum = 1 Then mlngRowNum = mrsPatient.RecordCount + 1
        mrsPatient.Filter = ""
        mrsPatient.Filter = "ID<" & mlngRowNum
        If mrsPatient.RecordCount > 0 Then
            mrsPatient.MoveLast
            mlngRowNum = Val(mrsPatient("ID").Value)
            txt.Text = zlCommFun.NVL(mrsPatient("姓名").Value)
            mrsParam("病人id").Value = Val(mrsPatient("病人id").Value)
            mrsParam("主页id").Value = Val(mrsPatient("主页id").Value)
            mrsParam("婴儿").Value = 0
            mrsParam("版本").Value = 0
            Select Case CStr(mrsPatient("类型").Value)
            Case "死亡", "死亡病人", "出院病人"
                mrsParam("出院").Value = 1
            Case Else
                mrsParam("出院").Value = 0
            End Select

            Call OpenPatientMap(Val(mrsParam("病人id").Value), Val(mrsParam("主页id").Value), Val(mrsParam("婴儿").Value))
            Call ExecuteCommand("读取记录")
            Call ExecuteCommand("清除数据")
            Call ExecuteCommand("读取数据")
            Call ExecuteCommand("读取组别数据")
            DataChanged = False
            
            txt.Tag = ""
        End If
        mrsPatient.Filter = ""
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Backward
        
        If mlngRowNum = mrsPatient.RecordCount Then mlngRowNum = 0
        mrsPatient.Filter = ""
        mrsPatient.Filter = "ID>" & mlngRowNum
        If mrsPatient.RecordCount > 0 Then
            mrsPatient.MoveFirst
            mlngRowNum = Val(mrsPatient("ID").Value)
            txt.Text = zlCommFun.NVL(mrsPatient("姓名").Value)
            mrsParam("病人id").Value = Val(mrsPatient("病人id").Value)
            mrsParam("主页id").Value = Val(mrsPatient("主页id").Value)
            mrsParam("婴儿").Value = 0
            mrsParam("版本").Value = 0
            Select Case CStr(mrsPatient("类型").Value)
            Case "死亡", "死亡病人", "出院病人"
                mrsParam("出院").Value = 1
            Case Else
                mrsParam("出院").Value = 0
            End Select

            Call OpenPatientMap(Val(mrsParam("病人id").Value), Val(mrsParam("主页id").Value), Val(mrsParam("婴儿").Value))
            Call ExecuteCommand("读取记录")
            Call ExecuteCommand("清除数据")
            Call ExecuteCommand("读取数据")
            Call ExecuteCommand("读取组别数据")
            DataChanged = False
            
            txt.Tag = ""
        End If
        mrsPatient.Filter = ""
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Location
        
        Call LocationObj(txt)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Help_Help
    
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Exit
    
        Unload Me
        Exit Sub
        
    End Select

End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsThis_Resize()
'    Dim lngLeft As Long
'    Dim lngTop  As Long
'    Dim lngRight  As Long
'    Dim lngBottom  As Long
'
'    Call cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
'
'    On Error Resume Next
'
'    picPane(0).Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    On Error Resume Next
    
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Audit
        
        Control.Visible = (mbytMode < 3)
        Control.Enabled = Control.Visible
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem, conMenu_Edit_NewParent
        Control.Visible = (mbytMode < 3)
        Control.Enabled = Control.Visible And mblnChanged = False
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Append
        Control.Visible = (mbytMode < 3)
        Control.Enabled = Control.Visible
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Save
        
        Control.Visible = (mbytMode < 3)
        Control.Enabled = mblnChanged And Control.Visible

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Cancle
    
        Control.Visible = (mbytMode < 3)
        Control.Enabled = mblnChanged And Control.Visible
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete                    '删除已添加的活动项目
        
        Control.Visible = (mbytMode < 3)
        Control.Enabled = (Val(Vsf.TextMatrix(Vsf.Row, mCol.项目性质)) = 2) And Control.Visible
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Sign                      '签名
        
        Control.Visible = (mbytMode = 3)
        Control.Enabled = Control.Visible
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Untread                   '取消签名
    
        Control.Visible = (mbytMode = 4)
        Control.Enabled = Control.Visible
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_LocationItem
    
        Control.Checked = (mstrFindKey = Control.Parameter)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Location
        
        Control.Enabled = (mblnChanged = False)
        cmd.Enabled = Control.Enabled And (mbytMode = 1 Or mbytMode = 2 Or mbytMode = 5)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Forward
        
        Control.Enabled = (mrsPatient.RecordCount > 1 And mblnChanged = False And (mbytMode = 1 Or mbytMode = 2 Or mbytMode = 5) And mlngRowNum > 1)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Backward
        
        Control.Enabled = (mrsPatient.RecordCount > 1 And mblnChanged = False And (mbytMode = 1 Or mbytMode = 2 Or mbytMode = 5) And mlngRowNum < mrsPatient.RecordCount)
        
    End Select
End Sub

Private Sub cmd_Click()
    Dim RS As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim mstrSort As String
    
    '------------------------------------------------------------------------------------------------------------------
    mrsPatient.Filter = ""
    If mrsPatient.RecordCount > 0 Then
        mrsPatient.MoveFirst
        If ShowTxtSelDialog(Me, txt, "床号,1200,0,0;姓名,1200,0,1;性别,600,0,0;科室,1800,0,0;住院号,1080,0,0", Me.Name & "\病人清单选择", "请从下面选择一个病人。", mrsPatient, RS, 5600, 4500, , CStr(mlngRowNum), 2, True) Then
            
            mlngRowNum = Val(mrsPatient("ID").Value)
            
            txt.Text = zlCommFun.NVL(RS("姓名").Value)

            mrsParam("病人id").Value = Val(RS("病人id").Value)
            mrsParam("主页id").Value = Val(RS("主页id").Value)
            mrsParam("婴儿").Value = 0
            Select Case CStr(RS("类型").Value)
            Case "死亡", "死亡病人", "出院病人"
                mrsParam("出院").Value = 1
            Case Else
                mrsParam("出院").Value = 0
            End Select
            
            Call OpenPatientMap(Val(mrsParam("病人id").Value), Val(mrsParam("主页id").Value), Val(mrsParam("婴儿").Value))
            
            Call ExecuteCommand("读取记录")
            Call ExecuteCommand("清除数据")
            Call ExecuteCommand("读取数据")
            Call ExecuteCommand("读取组别数据")
            
            DataChanged = False
    
            txt.Tag = ""
        End If
    End If
    mrsPatient.Filter = ""
    
    Call LocationObj(txt)

    Exit Sub
    
    '------------------------------------------------------------------------------------------------------------------
ErrHand:
End Sub




Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(0).hWnd
    Case 2
        Item.Handle = picPane(1).hWnd
    End Select
End Sub

Private Sub dtp_Change()
    dtp.Tag = "Changed"
    DataChanged = True
    
    Call ExecuteCommand("历史数据", Format(dtp.Value, "yyyy-MM-dd HH:mm:ss"), Val(Vsf.RowData(Vsf.Row)))
End Sub

Private Sub dtp_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        
        Vsf.Row = 1
        Vsf.Col = mCol.记录结果
        Vsf.SetFocus
    End If
    
End Sub


Private Sub dtp_LostFocus()

'    If dtp.Tag = "Changed" Then
'        '读取此时间点的值
'        dtp.Tag = ""
'        Call ReadData
'    End If
    
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Call SetPaneRange(dkpMain, 2, 15, 100, Me.ScaleWidth, 150)
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If DataChanged Then
        Cancel = (MsgBox("数据必须保存后才生效，是否放弃保存？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
    End If
    
    If Cancel Then Exit Sub
    
    On Error Resume Next
    zlCommFun.OpenIme False
    
    Call zlDatabase.SetPara("查找方法", mstrFindKey, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zlDatabase.SetPara("历史数据天数", Val(txtDay.Text), glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    
    Call SaveWinState(Me, App.ProductName)
    
    Set mrsPatient = Nothing
    Set mobjFindKey = Nothing
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    '------------------------------------------------------------------------------------------------------------------
    Case 0
    
        fraInfo.Move 0, -90, picPane(Index).Width
        tbcPage.Move 0, fraInfo.Top + fraInfo.Height, picPane(Index).Width, picPane(Index).Height - (fraInfo.Top + fraInfo.Height)
    '------------------------------------------------------------------------------------------------------------------
    Case 1
        fra.Move 0, -90, picPane(Index).Width
        vsfHistory.Move 15, fra.Top + fra.Height + 15, picPane(Index).Width - 30, picPane(Index).Height - (fra.Top + fra.Height) - 30
        vsfHistory.AutoSize 1, 1
    '------------------------------------------------------------------------------------------------------------------
    Case 2
        
        fraTime.Move 0, -90, picPane(Index).Width
'        lbl(0).Move 30, fraTime.Top + fraTime.Height + 45
'        tbs.Move tbs.Left, fraTime.Top + fraTime.Height, picPane(Index).Width - tbs.Left
        tbs.Width = fraTime.Width - tbs.Left - cboBaby.Width - 90
        cboBaby.Left = tbs.Left + tbs.Width + 30
        Vsf.Move 0, fraTime.Top + fraTime.Height, picPane(Index).Width, picPane(Index).Height - (fraTime.Top + fraTime.Height)
        cbo.Move fraTime.Width - cbo.Width - 90
        
    End Select
    
End Sub

Private Sub tbcPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mblnStartUp Then Exit Sub
    
    Call ExecuteCommand("清除数据")
    
    If Val(tbcPage.Selected.Tag) > 0 Then
        Call ExecuteCommand("读取数据")
    End If
    
    Call ExecuteCommand("读取组别数据")
    
    DataChanged = False
End Sub

Private Sub tbs_Click()
    Call ReadData
End Sub

Private Sub txt_Change()
    txt.Tag = "Changed"
End Sub

Private Sub txt_GotFocus()
    zlControl.TxtSelAll txt
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    Dim bytMode As Byte
    Dim lng病人ID As Long
    Dim strInput As String
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        If txt.Tag = "Changed" And txt.Text <> "" Then
            If InStr(txt.Text, "'") Then
                ShowSimpleMsg "输入的内容中有非法字符 ' ！"
                Exit Sub
            End If
            
            Select Case mstrFindKey
'            Case "病人id"
'                strInput = "病人id=" & Val(txt.Text)
'                bytMode = 2
'            Case "门诊号"
'                strInput = "门诊号=" & Val(txt.Text)
'                bytMode = 4
            Case "床  号"
                strInput = "床号='" & Trim(txt.Text) & "'"
                bytMode = 5
            Case "住院号"
                strInput = "住院号=" & Val(txt.Text)
                bytMode = 3
            Case "就诊卡"
                strInput = "就诊卡号='" & Trim(txt.Text) & "'"
                bytMode = 1
            End Select
                        
        End If

    ElseIf mstrFindKey = "就诊卡" And txt.Tag = "Changed" And txt.Text <> "" Then
        If Len(txt.Text) = mint就诊卡号码长度 - 1 And KeyAscii <> 8 Or KeyAscii = 13 And txt.Text <> "" Then
            If KeyAscii <> 13 Then
                txt.Text = txt.Text & Chr(KeyAscii)
                txt.SelStart = Len(txt.Text)
                KeyAscii = 0
            End If

            strInput = "就诊卡号='" & Trim(txt.Text) & "'"
            bytMode = 1
        End If
    End If
    
    If strInput <> "" Then
        txt.Tag = ""
        mrsPatient.Filter = ""
        mrsPatient.Filter = strInput
        If mrsPatient.RecordCount > 0 Then
            mrsPatient.MoveFirst
            lng病人ID = Val(mrsPatient("病人id").Value)
            mlngRowNum = Val(mrsPatient("ID").Value)
            
            txt.Text = zlCommFun.NVL(mrsPatient("姓名").Value)
            txt.Tag = ""
            
            mrsParam("病人id").Value = Val(mrsPatient("病人id").Value)
            mrsParam("主页id").Value = Val(mrsPatient("主页id").Value)
            mrsParam("婴儿").Value = 0
            Select Case CStr(mrsPatient("类型").Value)
            Case "死亡", "死亡病人", "出院病人"
                mrsParam("出院").Value = 1
            Case Else
                mrsParam("出院").Value = 0
            End Select
            
            Call OpenPatientMap(Val(mrsParam("病人id").Value), Val(mrsParam("主页id").Value), Val(mrsParam("婴儿").Value))
        Else
            ShowSimpleMsg "没有找到符合条件的病人！"
            txt.Text = mstrSvr姓名
        End If
        mrsPatient.Filter = ""

        Call LocationObj(txt)
        
    End If

    Exit Sub

ErrHand:
End Sub

Private Sub txtShow_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtShow(Index).Locked Then
        glngTXTProc = GetWindowLong(txtShow(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtShow(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtShow_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtShow(Index).Locked Then
        Call SetWindowLong(txtShow(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub udnDay_Change()
    Call ExecuteCommand("历史数据", Format(dtp.Value, "yyyy-MM-dd HH:mm:ss"), Val(Vsf.RowData(Vsf.Row)))
End Sub

Private Sub vsf_AfterDeleteCell(ByVal Row As Long, ByVal Col As Long)
    DataChanged = True
    Vsf.TextMatrix(Row, mCol.是否变动) = "1"
End Sub

Private Sub vsf_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    DataChanged = True
    Vsf.TextMatrix(Row, mCol.是否变动) = "1"
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        
    DataChanged = True
    
    Select Case Col
    '------------------------------------------------------------------------------------------------------------------
    Case mCol.记录结果
    
        Call Vsf.Body.AutoSize(mCol.记录结果, mCol.记录结果)
        
        Vsf.TextMatrix(Row, mCol.未记说明) = ""
        
    '------------------------------------------------------------------------------------------------------------------
    Case mCol.标记
        
        Vsf.TextMatrix(Row, mCol.未记说明) = ""
        
        If Val(Vsf.RowData(Row)) = 10 Then
            
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case mCol.未记说明
    
    End Select
    
    Vsf.TextMatrix(Row, mCol.是否变动) = "1"
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strValue As String
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    
    With Vsf
        .ComboList(mCol.部位) = ""
        .EditMode(mCol.部位) = 0
        .ComboList(mCol.标记) = ""
        .EditMode(mCol.标记) = 0
            
        Select Case Val(.RowData(NewRow))
        Case 1
            .ComboList(mCol.部位) = "口温|腋温|肛温"
            .EditMode(mCol.部位) = 1
            
            .ComboList(mCol.标记) = ""
            .EditMode(mCol.标记) = 1
        Case 2
            .ComboList(mCol.部位) = " |起搏器"
            .EditMode(mCol.部位) = 1
            If mint心率应用 = 2 Then
                .ComboList(mCol.标记) = ""
                .EditMode(mCol.标记) = 1
            End If
        Case 3
            .ComboList(mCol.部位) = "自主呼吸|呼吸机"
            .EditMode(mCol.部位) = 1
            
            .ComboList(mCol.标记) = ""
            .EditMode(mCol.标记) = 1
        Case 9
            .ComboList(mCol.标记) = " |C|/C"
            .EditMode(mCol.标记) = 1
        Case 10
            .ComboList(mCol.标记) = " |*|E|/E"
            .EditMode(mCol.标记) = 1
        Case Else
            If Val(.TextMatrix(NewRow, mCol.项目性质)) = 2 Then
                gstrSQL = " Select 部位 From 体温部位 Where 项目序号=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取该活动项目对应的部位", CLng(.RowData(NewRow)))
                If Err = 0 Then
                    Do While Not rsTemp.EOF
                        strValue = strValue & "|" & rsTemp!部位
                        rsTemp.MoveNext
                    Loop
                    If strValue <> "" Then
                        .ComboList(mCol.部位) = Mid(strValue, 2)
                        .EditMode(mCol.部位) = 1
                    End If
                End If
            End If
        End Select
        
        Select Case Trim(.TextMatrix(NewRow, mCol.标记))
        Case "*"
            .EditMode(mCol.记录结果) = 0
        Case Else
            .EditMode(mCol.记录结果) = 1
        End Select
        
        Select Case Val(.TextMatrix(NewRow, mCol.项目表示))
        Case 0  '文本
            strValue = ""
            If Val(.TextMatrix(NewRow, mCol.项目长度)) >= 200 Then strValue = "..."
            .ComboList(mCol.记录结果) = strValue
            .Body.ColComboList(mCol.记录结果) = ""
        Case 1  '上下
            .ComboList(mCol.记录结果) = ""
            .Body.ColComboList(mCol.记录结果) = ""
        Case 2  '单选
            .ComboList(mCol.记录结果) = ""
            .Body.ColComboList(mCol.记录结果) = " |" & .TextMatrix(NewRow, mCol.项目值域)
        Case 3  '复选
            .ComboList(mCol.记录结果) = "..."
            .Body.ColComboList(mCol.记录结果) = "..."
        End Select
        
        Dim varAry As Variant
        Dim strTmp As String
        
        If Val(.TextMatrix(NewRow, mCol.项目类型)) = 0 Then
            Select Case Val(.TextMatrix(NewRow, mCol.项目表示))
            Case 0, 1
                If .TextMatrix(NewRow, mCol.项目值域) <> "" Then
                    varAry = Split(.TextMatrix(NewRow, mCol.项目值域), "|")
                    
                    If UBound(varAry) >= 1 Then
                        strTmp = Val(varAry(0)) & "～" & Val(varAry(1))
                    End If
                End If
            End Select
        End If
        
        stbThis.Panels(3).Text = "范围：" & strTmp
                
        Select Case Val(.RowData(NewRow))
        Case 1
            strTmp = "标记表示物理降温的温度，部位为测体温的部位。"
        Case 2
            If mint心率应用 = 2 Then
                strTmp = "标记表示心率的值（与脉搏不同时才记录）。"
            Else
                strTmp = "部位中选择是否使用起搏器"
            End If
        Case 3
            strTmp = "部位为呼吸方式，分为自主呼吸、呼吸机辅助呼吸"
        Case 9
            strTmp = "标记中的 C 表示保留导尿。"
        Case 10
            strTmp = "标记中的 * 表示失禁或假肛; E 表示灌肠; /E 表示灌肠后的排泄。"
        Case Else
            strTmp = ""
        End Select
        
        stbThis.Panels(2).Text = strTmp
        
        If Val(.TextMatrix(NewRow, mCol.项目类型)) = 1 Then
            zlCommFun.OpenIme True
        Else
            zlCommFun.OpenIme False
        End If
        
        
        If NewRow <> OldRow Then
            
            Call ExecuteCommand("历史数据", Format(dtp.Value, "yyyy-MM-dd HH:mm:ss"), Val(Vsf.RowData(NewRow)))
            
        End If
    End With
End Sub

Private Sub vsf_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case mCol.标记, mCol.记录结果, mCol.未记说明
        If Vsf.TextMatrix(Row, Col) <> "" Then
            Vsf.TextMatrix(Row, Col) = ""
            Vsf.TextMatrix(Row, mCol.是否变动) = "1"
            DataChanged = True
        End If
    End Select
    Cancel = True
End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = True
End Sub


Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim RS As New ADODB.Recordset
    Dim varAry As Variant
    Dim lngLoop As Long
    Dim objPoint As POINTAPI
    Dim lngX As Long
    Dim lngY As Long
    Dim lngCX As Long
    Dim strTmp As String
    
    Select Case Col
    Case mCol.记录结果
        strTmp = frmWordsEditor.ShowMe(Me, mrsParam!病人ID, mrsParam!主页ID, Vsf.TextMatrix(Row, mCol.记录结果))
        If strTmp = "" Then Exit Sub
        Vsf.EditText = strTmp
        Vsf.TextMatrix(Row, mCol.记录结果) = strTmp
        Vsf.TextMatrix(Row, mCol.是否变动) = "1"
        mblnChanged = True
        
    Case mCol.未记说明
    
        gstrSQL = "Select 编码,名称,RowNum As ID,1 As 末级 From 常用体温说明"
        If ShowGrdSelectDialog(Me, Vsf, "名称,3000,0,0", Me.Name & "\常用体温说明", "请从下面选择一个未记录说明。", gstrSQL, RS, 4500, 4500, False, 2) Then
            Vsf.EditText = zlCommFun.NVL(RS("名称").Value)
            Vsf.Cell(flexcpData, Row, Col) = zlCommFun.NVL(RS("名称").Value)
            Vsf.TextMatrix(Row, Col) = zlCommFun.NVL(RS("名称").Value)
            Vsf.TextMatrix(Row, mCol.记录结果) = ""
            Vsf.TextMatrix(Row, mCol.标记) = ""
            Vsf.TextMatrix(Row, mCol.部位) = ""
            
            Vsf.TextMatrix(Row, mCol.是否变动) = "1"

            mblnChanged = True
        End If
        
    Case Else
        '针对复选结果
        Call CreateParam(RS, "ID", adBigInt)
        Call CreateParam(RS, "末级", adTinyInt)
        Call CreateParam(RS, "名称", adVarChar, 200)
        Call CreateParam(RS, "选择", adTinyInt)
        RS.Open
        If Vsf.TextMatrix(Row, mCol.项目值域) <> "" Then
    
            strTmp = ";" & Vsf.TextMatrix(Row, Col) & ";"
    
            varAry = Split(Vsf.TextMatrix(Row, mCol.项目值域), "|")
            For lngLoop = 0 To UBound(varAry)
                RS.AddNew
                RS("ID").Value = lngLoop
                RS("末级").Value = 1
                RS("名称").Value = CStr(varAry(lngLoop))
    
                If InStr(strTmp, ";" & CStr(varAry(lngLoop)) & ";") > 0 Then
                    RS("选择").Value = 1
                Else
                    RS("选择").Value = 0
                End If
            Next
            If RS.RecordCount > 0 Then RS.MoveFirst
        End If
    
        Call ClientToScreen(Vsf.hWnd, objPoint)
    
        lngX = objPoint.X * Screen.TwipsPerPixelX + Vsf.CellLeft
        lngY = objPoint.Y * Screen.TwipsPerPixelY + Vsf.CellTop + Vsf.CellHeight
    
        strTmp = ""
        
        lngCX = Vsf.Width - Vsf.Body.ColWidth(0) - Vsf.Body.ColWidth(1) - 75
        If lngCX < 3300 Then lngCX = 3300
        
        If frmSelectDialog.ShowSelect(Me, 2, RS, "名称,3600,0,1", "请在要选择的项目前画上√", lngX, lngY, lngCX, 3900, Vsf.CellHeight, , Me.Name & "\护理结果选择", , False, True) Then
            RS.Filter = ""
            RS.Filter = "选择=1"
            If RS.RecordCount > 0 Then RS.MoveFirst
            Do While Not RS.EOF
                strTmp = strTmp & ";" & RS("名称").Value
                RS.MoveNext
            Loop
    
            If strTmp <> "" Then strTmp = Mid(strTmp, 2)
            Vsf.TextMatrix(Row, Col) = strTmp
            Vsf.TextMatrix(Row, mCol.是否变动) = "1"
            mblnChanged = True
        End If
    End Select
    
End Sub

Private Sub vsf_ChangeEdit()
    
    With Vsf
        Select Case .Col
        Case mCol.记录结果
            Select Case Val(.TextMatrix(.Row, mCol.项目表示))
            Case 0
                .TextMatrix(.Row, mCol.记录结果) = .EditText
                Call .Body.AutoSize(mCol.记录结果, mCol.记录结果)
            Case 2
                '下拉
                .TextMatrix(.Row, mCol.记录结果) = .EditText
            End Select
            
            Vsf.TextMatrix(.Row, mCol.未记说明) = ""
        Case mCol.未记说明
            .TextMatrix(.Row, mCol.未记说明) = .EditText
            If Trim(.TextMatrix(.Row, mCol.未记说明)) <> "" Then
                
                .TextMatrix(.Row, mCol.记录结果) = ""
                .TextMatrix(.Row, mCol.标记) = ""
                .TextMatrix(.Row, mCol.部位) = ""

            End If
            
        Case mCol.标记
            Select Case Val(.RowData(.Row))
            Case 9
                .TextMatrix(.Row, mCol.标记) = .EditText
                
                Select Case Trim(.TextMatrix(.Row, mCol.标记))
                Case "C"
                    .EditMode(mCol.记录结果) = 0
                    .TextMatrix(.Row, mCol.记录结果) = ""
                Case Else
                    .EditMode(mCol.记录结果) = 1
                End Select
                
            Case 10
                .TextMatrix(.Row, mCol.标记) = .EditText
                Select Case Trim(.TextMatrix(.Row, mCol.标记))
                Case "*"
                    .EditMode(mCol.记录结果) = 0
                    .TextMatrix(.Row, mCol.记录结果) = ""
                Case Else
                    .EditMode(mCol.记录结果) = 1
                End Select
            
            End Select
            Vsf.TextMatrix(.Row, mCol.未记说明) = ""
        End Select
        .TextMatrix(.Row, mCol.是否变动) = "1"
    End With
    
    DataChanged = True
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    If KeyCode = vbKeyReturn Then
        
        If Col = mCol.未记说明 Or Col = mCol.记录结果 Then
            
            Vsf.Cell(flexcpData, Row, Col) = Vsf.EditText
            Vsf.TextMatrix(Row, Col) = Vsf.EditText
            Vsf.TextMatrix(Row, mCol.是否变动) = "1"
            
        End If
        
    End If
End Sub

Private Sub vsf_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    
    On Error Resume Next
    
    If KeyAscii <> vbKeyReturn Then
        If Val(Vsf.TextMatrix(Row, mCol.项目类型)) = 0 Then
            If Col = mCol.标记 Or Col = mCol.记录结果 Then
                If FilterKeyAscii(KeyAscii, 99, "0123456789.") = 0 Then KeyAscii = 0
            Else
                If FilterKeyAscii(KeyAscii, 99, "'") > 0 Then KeyAscii = 0
            End If
        End If
    End If
    
End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    
    On Error Resume Next
    
    If KeyAscii <> vbKeyReturn Then
        If Val(Vsf.TextMatrix(Row, mCol.项目类型)) = 0 Then
'            If FilterKeyAscii(KeyAscii, 99, "0123456789.") = 0 Then KeyAscii = 0
        End If
    End If
    
End Sub

Private Sub vsf_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngColor As Long
    Dim varAry As Variant
    
    Select Case Col
    Case mCol.记录结果
        GoTo CheckPoint
    Case mCol.标记
        
        Select Case Val(Vsf.RowData(Row))
        Case 1
            GoTo CheckPoint
        Case 2
            GoTo CheckPoint
        End Select
        
    End Select
    
    Exit Sub
    
CheckPoint:

    If Val(Vsf.TextMatrix(Row, mCol.项目类型)) = 0 And Trim(Vsf.EditText) <> "" And IsNumeric(Vsf.EditText) Then
        Select Case Val(Vsf.TextMatrix(Row, mCol.项目表示))
        Case 0, 1
            If Vsf.TextMatrix(Row, mCol.项目值域) <> "" Then
                varAry = Split(Vsf.TextMatrix(Row, mCol.项目值域), "|")
                
                If UBound(varAry) >= 1 Then
                    If Val(Vsf.EditText) < Val(varAry(0)) Or Val(Vsf.EditText) > Val(varAry(1)) Then
                        Vsf.TextMatrix(Row, Col) = Vsf.EditText

                        ShowSimpleMsg "“" & Vsf.TextMatrix(Row, mCol.护理项目) & " ”的范围应在（" & varAry(0) & "～" & varAry(1) & "）之间！"

                    End If
                End If
                
            End If
            
            If CheckNumber(Val(Vsf.EditText), Val(Vsf.TextMatrix(Row, mCol.项目长度)), Val(Vsf.TextMatrix(Row, mCol.项目小数))) = False Then
                
                Vsf.EditText = ""
                Cancel = True
                Exit Sub
            End If
        
        End Select
    End If
    
    Select Case Col
    Case mCol.记录结果
        
        lngColor = GridTextColor(Vsf.TextMatrix(Row, 2), Vsf.TextMatrix(Row, Col))
        Vsf.Cell(flexcpForeColor, Row, mCol.记录结果, Row, mCol.记录结果) = lngColor
        
    Case mCol.标记

        
    End Select
    

End Sub


Private Sub vsfHistory_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vsfHistory.AutoSize 1, 1
End Sub

Private Sub vsfHistory_DblClick()
    With vsfHistory
        If Trim(.TextMatrix(.Row, 1)) <> "" Then
            Vsf.TextMatrix(Vsf.Row, mCol.记录结果) = .TextMatrix(.Row, 1)
            Vsf.TextMatrix(Vsf.Row, mCol.是否变动) = "1"
            Call Vsf.Body.AutoSize(mCol.记录结果, mCol.记录结果)
            DataChanged = True
        End If
        
    End With
End Sub

Private Sub vsfHistory_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call vsfHistory_DblClick
    End If
End Sub

Private Function CheckTime(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal strTime As String, ByVal strCurTime As String) As Boolean
    Dim blnExist As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    '数据发生时间必须在当前科室的有效时间范围内
    
    gstrSQL = " Select 开始原因,病区ID,to_char(开始时间,'yyyy-MM-dd hh24:mi') AS 开始时间,to_char(nvl(终止时间,sysdate+" & mintPreDays & "),'yyyy-MM-dd hh24:mi') AS 终止时间 " & _
              " From 病人变动记录 " & _
              " Where 病人ID=[1] And 主页ID=[2]" & _
              " Order by 开始时间,开始原因"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取当前科室有效时间范围", lng病人ID, lng主页ID)
    With rsTemp
        .Filter = "病区ID=" & Val(mrsParam("病区id"))
        Do While Not .EOF
            If strTime >= !开始时间 And strTime <= NVL(!终止时间, strCurTime) Then
                blnExist = True
                Exit Do
            End If
            .MoveNext
        Loop
        .Filter = 0
        '找到了就退出
        If blnExist Then
            If Not IsAllowInput(lng病人ID, lng主页ID, strTime, strCurTime) Then
                MsgBox "发生时间" & strTime & "有误！[超过数据补录的有效时限:" & glngHours & "小时]", vbInformation, gstrSysName
                GoTo exitHand
            End If
            
            CheckTime = True
            Exit Function
        End If
        
        '没找到,就整理原因进行准确性提示
        .Filter = "开始原因=1"
        If .RecordCount <> 0 Then
            If !开始原因 = 1 And strTime < !开始时间 Then
                MsgBox "发生时间" & strTime & "有误！[发生时间不能小于病人入院时间:" & !开始时间 & "]", vbInformation, gstrSysName
                GoTo exitHand
            End If
        End If
        .Filter = "开始原因=2"
        If .RecordCount <> 0 Then
            If !开始原因 = 2 And strTime < !开始时间 Then
                MsgBox "发生时间" & strTime & "有误！[发生时间不能小于病人入科时间:" & !开始时间 & "]", vbInformation, gstrSysName
                GoTo exitHand
            End If
        End If
        .Filter = "开始原因=10"
        If .RecordCount <> 0 Then
            If !开始原因 = 10 And strTime > !终止时间 Then
                MsgBox "发生时间" & strTime & "有误！[发生时间不能大于出院时间:" & !终止时间 & "]", vbInformation, gstrSysName
                GoTo exitHand
            End If
        End If
        .Filter = 0
        '其他情况说明
        MsgBox "发生时间" & strTime & "有误！[不在当前病区的有效时间范围内]", vbInformation, gstrSysName
        GoTo exitHand
    End With
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
exitHand:
    rsTemp.Filter = 0
End Function
