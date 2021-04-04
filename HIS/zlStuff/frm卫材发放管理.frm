VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm卫材发放管理 
   Caption         =   "卫材发放管理"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   525
   ClientWidth     =   11400
   Icon            =   "frm卫材发放管理.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshBody 
      Height          =   1860
      Left            =   165
      TabIndex        =   8
      Top             =   5220
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   3281
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frm卫材发放管理.frx":014A
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox PicLine_S 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   45
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   4815
      TabIndex        =   6
      Top             =   3000
      Width           =   4815
   End
   Begin TabDlg.SSTab tabShow 
      Height          =   360
      Left            =   15
      TabIndex        =   2
      Top             =   615
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   635
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "待发料清单(&1)"
      TabPicture(0)   =   "frm卫材发放管理.frx":0464
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "发退料清单(&2)"
      TabPicture(1)   =   "frm卫材发放管理.frx":0480
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Chk清单"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.CheckBox Chk清单 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "显示所有过程单据"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4710
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   75
         Width           =   3450
      End
   End
   Begin VB.Timer TimeRefresh 
      Enabled         =   0   'False
      Left            =   4200
      Top             =   240
   End
   Begin VB.Timer TimePrintCancelBill 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4905
      Top             =   285
   End
   Begin MSComctlLib.ImageList ImgTbarBlack 
      Left            =   8460
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgTbarColor 
      Left            =   7890
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   1164
      BandCount       =   2
      _CBWidth        =   11400
      _CBHeight       =   660
      _Version        =   "6.7.9782"
      Child1          =   "tlbThis"
      MinHeight1      =   600
      Width1          =   4995
      NewRow1         =   0   'False
      Caption2        =   "发料部门"
      Child2          =   "cboStock"
      MinHeight2      =   300
      Width2          =   3000
      NewRow2         =   0   'False
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   5970
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   180
         Width           =   5340
      End
      Begin MSComctlLib.Toolbar tlbThis 
         Height          =   600
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   1058
         ButtonWidth     =   820
         ButtonHeight    =   1058
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "预览"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "发料"
               Key             =   "发料"
               Object.ToolTipText     =   "发料"
               Object.Tag             =   "发料"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退料"
               Key             =   "退料"
               Object.ToolTipText     =   "退料"
               Object.Tag             =   "退料"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EditSp"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "过滤"
               Object.ToolTipText     =   "过滤"
               Object.Tag             =   "过滤"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   7185
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm卫材发放管理.frx":049C
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15028
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshHead 
      Height          =   3405
      Left            =   45
      TabIndex        =   7
      Top             =   1050
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   6006
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frm卫材发放管理.frx":0D2E
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.TabStrip tbsSel 
      Height          =   300
      Left            =   225
      TabIndex        =   4
      Top             =   4770
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   529
      Style           =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu MnuFile1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileBillprint 
         Caption         =   "单据打印(&B)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFileRestore 
         Caption         =   "打印退料通知单(&R)"
      End
      Begin VB.Menu mnuFile2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePara 
         Caption         =   "参数设置(&A)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFile3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditPay 
         Caption         =   "发料(&P)"
      End
      Begin VB.Menu mnuEditOutPay 
         Caption         =   "退料(&O)"
      End
      Begin VB.Menu mnuEditSplit0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPayCf 
         Caption         =   "按处方发料(&F)"
      End
      Begin VB.Menu mnuEditFpPay 
         Caption         =   "按票据号发料(&R)"
      End
      Begin VB.Menu mnuEditStrict 
         Caption         =   "按单据退料(&B)"
      End
      Begin VB.Menu mnuEditSp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditStop 
         Caption         =   "停止发料标记(&S)"
      End
      Begin VB.Menu mnuEditStopSp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPayType 
         Caption         =   "未发料模式(&W)"
         Checked         =   -1  'True
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditBackType 
         Caption         =   "已发及退料模式(&H)"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuEditSelSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelAll 
         Caption         =   "全选(&Q)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "全清(&C)"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "报表(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuReportItem 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu sdfsdfsd 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuView1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFont 
         Caption         =   "字体(&O)"
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "小字体(&S)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "中字体(&M)"
            Index           =   1
         End
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "大字体(&B)"
            Index           =   2
         End
      End
      Begin VB.Menu MnuView2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "过滤(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu MnuView3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu MnuHelpTitle 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu MnuHelpWeb 
         Caption         =   "Web上的中联(&W)"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&E)..."
         End
      End
      Begin VB.Menu MnuHelp1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frm卫材发放管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private msngOldY As Single          '保存Y轴数据
Private mblnFirst As Boolean        '第一次激活系统
Dim mintFont As Integer

Private Type HeadCol
    标志    As Byte
    类型    As Byte
    科室    As Byte
    单据    As Byte
    收费    As Byte
    配料人  As Byte
    NO      As Byte
    姓名    As Byte
    床号    As Byte
    住院号  As Byte
    金额    As Byte
    日期    As Byte
    可操作  As Byte
    说明    As Byte
    Cols As Byte
End Type
Private mHeadCol As HeadCol

Private Type BodyCol
    费用ID      As Byte
    材料ID      As Byte
    状态        As Byte
    批次        As Byte
    在用分批    As Byte
    科室        As Byte
    开单医生    As Byte
    类型        As Byte
    NO          As Byte
    姓名        As Byte
    床号        As Byte
    住院号      As Byte
    卫材名称    As Byte
    规格        As Byte
    批号        As Byte
    单位        As Byte
    换算系数    As Byte
    付数        As Byte
    数量        As Byte
    原始数量    As Byte
    已退数      As Byte
    准退数      As Byte
    退料数      As Byte
    单价        As Byte
    金额        As Byte
    库存数      As Byte
    记帐员      As Byte
    发料人      As Byte
    发生时间    As Byte
    可操作      As Byte
    记录状态    As Byte
    Cols        As Byte
End Type

Private mBodyCol As BodyCol
Private mstrSelCon     As String   '选择的单据:格式是 No:单据:记录状态||No:单据:记录状态
Private mintCheckStock  As Integer  '库存检查方式

Private mstrStartDate As String
Private mstrEndDate As String
Private mlng科室id As Long
Private mlng病人id As Long
Private mstr住院号 As String
Private mstr病人姓名 As String
Private mstrStartNo As String
Private mstrEndNo As String
Private mint单据 As Long        '0-门诊及住院所有单据,1-门诊划价及门诊记帐,2-住院记帐
Private mint业务请求 As Long
Private mstr单据 As String  '以In方式
Private mlngCountSel As Long

'从药品处方发药传过来的参数
Private mblnTrans As Boolean            'True表示从药品处方发药窗口调用
Private mstrNo  As String               '单据号，仅用于定位
Private mlng库房id As Long              '发药库房ID，一般和发料部门一致
Private mstrDrugStartDate As String     '药品单据开始时间
Private mstrDrugEndDate As String       '药品单据结束时间
Private mlngModule As Long
Private mintUnit As Integer        '0-散装单位,1-包装单位

'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
Private mstrPreSelKey As String '上次选择的Key建

Private mintPrintPar As Integer    '0-提示打印,1-自动打印,2-不打印
Private mblnExit As Boolean
Private mstrPrintCon As String '打印条件
Private mstrPrivs As String '权限串

Public Sub ShowList(ByVal frmMain As Form, ByVal lng病人id As Long, ByVal strNo As String, ByVal lng库房ID As Long, ByVal strStartDate As String, ByVal strEndDate As String)
    mlng病人id = lng病人id
    mstrNo = strNo
    mlng库房id = lng库房ID
    mstrDrugStartDate = strStartDate
    mstrDrugEndDate = strEndDate
    mblnTrans = True
    
    Me.Show , frmMain
    Me.ZOrder 0

End Sub
Private Function CheckDepend() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:检查数据依赖性
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset
    Dim lng发料部门ID As Long
    
    CheckDepend = False
    
    gstrSQL = "" & _
        "   SELECT DISTINCT a.id, a.名称 " & _
        "   FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " & _
        "   Where c.工作性质 = b.名称 " & _
        "           AND b.编码 ='W' " & _
        "           AND a.id = c.部门id " & _
        "           AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'" & _
        IIf(InStr(gstrPrivs, "所有部门") <> 0, "", " And a.ID IN (Select 部门ID From 部门人员 Where 人员ID=[1])")
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取相应的库房", UserInfo.Id)
    
    If rsTemp.EOF Then
        ShowMsgBox "至少应该设置一个具有发料部门性质或者你" & vbCrLf & "不是发料部门的工作人员,请查看部门管理！"
        rsTemp.Close
        Exit Function
    End If
    
    '如果是药品窗口传入，设置发料部门与药品发药部门一致
    If mblnTrans Then
        If mlng库房id <> UserInfo.部门ID Then
            lng发料部门ID = mlng库房id
        Else
            lng发料部门ID = UserInfo.部门ID
        End If
    End If
    
    '装入发料部门数据
    With cboStock
        .Clear
        Do While Not rsTemp.EOF
            .AddItem rsTemp!名称
            .ItemData(.NewIndex) = rsTemp!Id
            If rsTemp!Id = lng发料部门ID Then
                .ListIndex = .NewIndex
            End If
            rsTemp.MoveNext
        Loop
        If .ListIndex = -1 Then .ListIndex = 0
        rsTemp.Close
    End With
    CheckDepend = True
End Function

Private Function InitSet()
    Dim i As Long
    Dim rsTemp As New ADODB.Recordset
    
    '设置时间，如果是从药品窗口传入，则与药品发药时间一致
    If mblnTrans Then
        mstrEndDate = mstrDrugEndDate
        mstrStartDate = mstrDrugStartDate
    Else
        mstrEndDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 23:59:59"
        mstrStartDate = Format(DateAdd("d", -7, zlDatabase.Currentdate), "yyyy-mm-dd") & " 00:00:00"
    End If
    
    Call LoadInIcon
    
    Call 权限控制
       
    '恢复字体
    Dim strReg As String
    strReg = zlDatabase.GetPara("字体字号", glngSys, mlngModule, "0")
    mnuViewFontSize_Click Val(strReg)
    
    strReg = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngModule, "0"))
    mintUnit = Val(strReg)
  
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
    End With
    With mOraFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价, True)
        .FM_金额 = GetFmtString(mintUnit, g_金额, True)
        .FM_零售价 = GetFmtString(mintUnit, g_售价, True)
        .FM_数量 = GetFmtString(mintUnit, g_数量, True)
    End With
    
    mintPrintPar = Val(zlDatabase.GetPara("发料打印提醒方式", glngSys, mlngModule, "0"))
    strReg = Trim(zlDatabase.GetPara("查询业务类型", glngSys, mlngModule, ""))
    If strReg = "" Then strReg = "24,25,26"
    mstr单据 = strReg
    
    With tbsSel
        .Tabs.Clear
        .Tabs.Add , "K1", "单据明细"
        .Tabs.Add , "K2", "明细情况"
        .Tabs.Add , "K3", "汇总情况"
    End With
        
    '网络列确定
    With mHeadCol
        .标志 = 0
        .类型 = 1
        .科室 = 2
        .单据 = 3
        .收费 = 4
        .配料人 = 5
        .NO = 6
        .姓名 = 7
        .床号 = 8
        .住院号 = 9
        .金额 = 10
        .日期 = 11
        .可操作 = 12
        .说明 = 13
        
        .Cols = 14
    End With
    With mBodyCol
       i = 0: .费用ID = i
       i = i + 1: .材料ID = i
       i = i + 1: .批次 = i
       i = i + 1: .在用分批 = i
       i = i + 1: .状态 = i
       i = i + 1: .科室 = i
       i = i + 1: .开单医生 = i
       i = i + 1: .类型 = i
       i = i + 1: .NO = i
       i = i + 1: .姓名 = i
       i = i + 1: .床号 = i
       i = i + 1: .住院号 = i
       i = i + 1: .卫材名称 = i
       i = i + 1: .规格 = i
       i = i + 1: .批号 = i
       i = i + 1: .单位 = i
       i = i + 1: .换算系数 = i
       i = i + 1: .付数 = i
       i = i + 1: .数量 = i
       i = i + 1: .原始数量 = i
       i = i + 1: .已退数 = i
       i = i + 1: .准退数 = i
       i = i + 1: .退料数 = i
       i = i + 1: .单价 = i
       i = i + 1: .金额 = i
       i = i + 1: .库存数 = i
       i = i + 1: .记帐员 = i
       i = i + 1: .发料人 = i
       i = i + 1: .可操作 = i
       i = i + 1: .记录状态 = i
       i = i + 1: .发生时间 = i
       
       i = i + 1: .Cols = i
    End With
    
End Function

Private Function ReadSystemPara()
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = " Select Nvl(检查方式,0) 库存检查 From 材料出库检查 Where 库房ID=[1]"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, cboStock.ItemData(cboStock.ListIndex))
    
    With rsTemp
        If Not .EOF Then
            mintCheckStock = NVL(!库存检查, 0)
        End If
    End With
    
End Function

Private Sub cboStock_Click()
        Call ReadSystemPara
        tabShow_Click -1
        SetMnuEnable
End Sub

Private Sub cbrThis_Resize()
    Form_Resize
End Sub

Private Sub Chk清单_Click()
    tabShow_Click 0
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    '初始列头
    Call SetGrdColHead(1)
    Call SetGrdColHead(2)
    
    '检查数据依赖关系
    If CheckDepend = False Then Unload Me: Exit Sub
    
    '窗体控件初始
    ' 1.权限控制
    Call 权限控制
    ' 2.表格列头初始
    
    
    '装载相关数据
    Call tabShow_Click(tabShow.Tab + 1)
 
        
End Sub

Private Sub Form_Load()
    mblnFirst = True
    mlngModule = glngModul
        
    mstrPrivs = gstrPrivs
    '初始相关控件
    Call InitSet
    
    '恢复个性化设置
    Call RestoreWinState(Me)
    '2006-04-25:刘兴宏,统一增加报表发布到模块的功能
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
     
End Sub

Private Sub Form_Resize()
    '窗体位置设置
    Dim sngCbrHeight As Single
    Dim sngStbHeight As Single
    
    
    On Error Resume Next
    sngCbrHeight = IIf(cbrThis.Visible, cbrThis.Height, 0)
    sngStbHeight = IIf(stbThis.Visible, stbThis.Height, 0)
    
    If Me.WindowState = vbMinimized Then Exit Sub
    cbrThis.Bands(1).MinHeight = tlbThis.Height
    If Me.Height < PicLine_S.Height + PicLine_S.Top + tbsSel.Height + 650 + sngStbHeight Then
        Me.Height = PicLine_S.Height + PicLine_S.Top + tbsSel.Height + 650 + sngStbHeight
    End If
    
    With tabShow
        .Left = 0
        .Width = ScaleWidth
        .Top = sngCbrHeight + 10
    End With
    
    With mshHead
        .Left = 0
        .Top = tabShow.Height + tabShow.Top + 10
        .Height = PicLine_S.Top - .Top
        .Width = ScaleWidth - 10
    End With
    
    With PicLine_S
        .Left = 0
        .Width = mshHead.Width
    End With
    
    With tbsSel
        .Left = 0
        .Top = PicLine_S.Top + PicLine_S.Height + 10
        .Width = mshHead.Width
    End With
    
    With mshBody
        .Left = 0
        If tabShow.Tab = 1 Then
            .Top = mshHead.Top
        Else
            .Top = tbsSel.Top + tbsSel.Height + 10
        End If
        .Height = ScaleHeight - .Top - sngStbHeight
        .Width = mshHead.Width
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)

    '保存个性化设置
    Call SaveWinState(Me)
End Sub
Private Function Save退料(ByVal strDate As String) As Boolean
    '检查退料的相关条件
    Dim strNo As String
    Dim lngID As Long
    Dim lngRow As Long
    Dim 材料ID As Long
    Dim int自动销帐 As Integer
    Dim strReg As String
    int自动销帐 = IIf(Val(zlDatabase.GetPara("自动销帐", glngSys, mlngModule)) = 1, 1, 0)
    
    Save退料 = False
    err = 0
    On Error GoTo ErrHand:
    
    gcnOracle.BeginTrans
    
    With mshBody
        For lngRow = 1 To .Rows - 1
        
                strNo = Trim(.TextMatrix(lngRow, mBodyCol.NO))
                '过程参数:ID_IN,审核人_IN,审核日期_IN,批号_IN,效期_IN,产地_IN,退料数量_IN,自动销帐_IN(1-自动销帐,0-不自动销帐)
                If strNo <> "" And Trim(.TextMatrix(lngRow, mBodyCol.状态)) = "√" Then
                   gstrSQL = "zl_材料收发记录_部门退料("
                   gstrSQL = gstrSQL & .RowData(lngRow) & ","
                   gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                   gstrSQL = gstrSQL & "to_date('" & strDate & "','yyyy-mm-dd HH24:mi:ss'),"
                   gstrSQL = gstrSQL & "'" & Replace(.TextMatrix(lngRow, mBodyCol.批号), "(" & .TextMatrix(lngRow, mBodyCol.批次) & ")", "") & "',"
                   gstrSQL = gstrSQL & "NULL" & ","
                   gstrSQL = gstrSQL & "NULL" & ","
                  ' If mintUnit = 0 Then
                        gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, mBodyCol.原始数量))
                   'Else
                 '       gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, mBodyCol.数量)) * Val(.TextMatrix(lngRow, mBodyCol.换算系数))
                  ' End If
                   gstrSQL = gstrSQL & "," & int自动销帐 & ")"
                   Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                End If
        Next
    End With
    gcnOracle.CommitTrans
    Save退料 = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Function

Private Sub mnuEditBackType_Click()
    Dim blnYes As Boolean
    If mstrSelCon <> "" Then
        ShowMsgBox "已经有被选择的项目,你是否希望改变选择!", True, blnYes
        If Not blnYes Then
            mblnExit = True
            mnuEditPayType.Checked = True
            mnuEditBackType.Checked = False
            Exit Sub
        End If
        mstrSelCon = ""
    End If
    
    tabShow.Tab = 1
    mnuEditPayType.Checked = False
    mnuEditBackType.Checked = True
    tabShow_Click 0
    
End Sub

Private Sub mnuEditClear_Click()
     If tabShow.Tab = 0 Then
        Call SelAndClearAllPlay(False)
     Else
        Call SelAndClearAllOutPlay(False)
     End If
     Call SetMnuEnable
End Sub

Private Sub mnuEditFpPay_Click()
    '按票据号发料
    With Frm按单号发料
        .In_单据 = mint业务请求
        .In_单据IN = mstr单据
        .In_发料部门id = cboStock.ItemData(cboStock.ListIndex)
        .In_库存检查 = mintCheckStock
        .In_允许未配料发料 = 1
        .按票据号发料 = True
        .In_权限 = mstrPrivs
        .mstr配料人 = gstrUserName
        .Show 1, Me
    End With
    mnuViewRefresh_Click
    
End Sub

Private Sub mnuEditOutPay_Click()
    Dim strDate As String
    Dim blnYe As Boolean
    
    ShowMsgBox "你是否真的要对这些记录进行退料吗?", True, blnYe
    If blnYe = False Then
        Exit Sub
    End If
    
    strDate = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:MM:SS")

    If Save退料(strDate) = False Then Exit Sub
    BillListPrint 0, strDate, 2
    mstrSelCon = ""
    mlngCountSel = 0
    mnuViewRefresh_Click
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditPay_Click()
    If CheckStock = False Then Exit Sub
    If SendBill = False Then Exit Sub
    mstrSelCon = ""
    mstrPrintCon = ""
    '数据刷新
    tabShow_Click 1
End Sub

Private Sub mnuEditPayCf_Click()

    With Frm按单号发料
        .In_单据 = mint业务请求
        .In_单据IN = mstr单据
        '.In_发料窗口 = str窗口
        .In_发料部门id = cboStock.ItemData(cboStock.ListIndex)
        .In_库存检查 = mintCheckStock
        '.In_校验处方 = intVerify
        .In_允许未配料发料 = 1
        '.IN_允许未审核发药 = Int允许未审核处方发药
        .In_权限 = mstrPrivs
        .mstr配料人 = gstrUserName
        .按票据号发料 = False
        .Show 1, Me
    End With
    mnuViewRefresh_Click
End Sub

Private Sub mnuEditPayType_Click()
'    Dim blnYes As Boolean
'    ShowMsgbox "已经有被选择的项目,你是否希望改变选择!", True, blnYes
'    If Not blnYes Then
'        mnuEditPayType.Checked = False
'        mnuEditBackType.Checked = True
'        Exit Sub
'    End If
    mlngCountSel = 0
    tabShow.Tab = 0
    mnuEditPayType.Checked = True
    mnuEditBackType.Checked = False
    tabShow_Click 1
End Sub

Private Sub mnuEditSelAll_Click()
     If tabShow.Tab = 0 Then
        Call SelAndClearAllPlay(True)
     Else
        Call SelAndClearAllOutPlay(True)
     End If
     Call SetMnuEnable
End Sub

Private Sub mnuEditStop_Click()
    '停止发料
    '发药方式=-1
        
    Dim frmFlag As New Frm不再发药处方标志
    frmFlag.gstrParentName = Me.Name
    frmFlag.Show vbModal
    mnuViewRefresh_Click
        
End Sub

Private Sub mnuEditStrict_Click()
    '
    If Frm按单号退料.ShowCard(Me, cboStock.ItemData(cboStock.ListIndex), mstrPrivs) = False Then Exit Sub
    Call mnuViewRefresh_Click
End Sub

Private Sub mnuFileBillPrint_Click()
        '单据打印:
        Dim lng单据 As Long
        Dim strNo As String
        Dim strDate As String
        Dim int发料方式 As Integer
        Dim rsTemp As New ADODB.Recordset
        If tabShow.Tab = 0 Then Exit Sub
        
        With mshBody
            
            strDate = .TextMatrix(.Row, mBodyCol.发生时间)
            lng单据 = Decode(.TextMatrix(.Row, mBodyCol.类型), "收费", 24, "记帐单", 25, "记帐表", 26, 0)
            strNo = Trim(.TextMatrix(.Row, mBodyCol.NO))
            mstrPrintCon = strNo & "||" & lng单据 & "||" & cboStock.ItemData(cboStock.ListIndex)
            If strNo = "" Then Exit Sub
            
            gstrSQL = "Select 发药方式 from 药品收发记录 where id=[1]"
            gstrSQL = AnalyseHistorySQL(gstrSQL)
            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, .RowData(.Row))

            If rsTemp.EOF Then
                Exit Sub
            End If
            int发料方式 = NVL(rsTemp!发药方式, 1)
            If int发料方式 <> 1 Then
                mstrPrintCon = ""
            End If
        End With
        BillListPrint int发料方式, strDate, 0
End Sub

Private Sub mnuFileExcel_Click()
    '输出到EXCEL
    subPrint 3
End Sub

Private Sub mnufileexit_Click()
    Unload Me
End Sub



Private Sub mnuFilePara_Click()
    Dim strReg As String
    
    If frmPayExitParaSet.ShowSetPara(Me, mlngModule, mstrPrivs) = False Then Exit Sub
    strReg = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngModule, "0"))
    mintUnit = Val(strReg)
    
    mintPrintPar = Val(zlDatabase.GetPara("发料打印提醒方式", glngSys, mlngModule, "0"))
    strReg = Trim(zlDatabase.GetPara("查询业务类型", glngSys, mlngModule, ""))
    If strReg = "" Then strReg = "24,25,26"
    mstr单据 = strReg
    '重新获取数据
    mnuViewRefresh_Click
End Sub
Private Sub mnuFilePreView_Click()
    '打印预览
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    '打印
    subPrint 1
End Sub
Private Sub mnuFileRestore_Click()
        '
        If tabShow.Tab <> 1 Then Exit Sub
        
        Dim strDate As String
        strDate = mshBody.TextMatrix(mshBody.Row, mBodyCol.发生时间)
        BillListPrint , strDate, 2
End Sub

Private Sub mnuFileSet_Click()
'打印设置
    zlPrintSet
End Sub

Private Sub mnuViewFind_Click()
    Dim strStartDate As String, strEndDate As String
    Dim strStartNo As String, strEndNo As String
    Dim str单据 As String, int业务请求 As Integer
    Dim str住院号 As String, lng病人id As Long, lng科室id As Long
    Dim str姓名 As String
    Dim blnreturn As Boolean
    
    blnreturn = frm卫材发放管理Search.ShowEdit( _
        Me, strStartDate, strEndDate, strStartNo, strEndNo, str单据, _
        int业务请求, str住院号, lng病人id, str姓名, lng科室id)
    If blnreturn = False Then Exit Sub
    
    mstrStartDate = strStartDate
    mstrEndDate = strEndDate
    mstrStartNo = strStartNo
    mstrEndNo = strEndNo
    mint单据 = int业务请求
    mint业务请求 = int业务请求
    mstr住院号 = str住院号
    mlng病人id = lng病人id
    mstr病人姓名 = str姓名
    mlng科室id = lng科室id
    
    mnuViewRefresh_Click
End Sub
 

Private Sub mnuViewFontSize_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 2
        Me.mnuViewFontSize(i).Checked = False
    Next
    Me.mnuViewFontSize(Index).Checked = True

    Select Case Index
    Case 0
        Me.mshHead.Font.Size = 9
        Me.tabShow.Font.Size = 9
        mshBody.Font.Size = 9
        tbsSel.Font.Size = 9
     Case 1
        Me.mshHead.Font.Size = 11
        Me.tabShow.Font.Size = 11
        mshBody.Font.Size = 11
        tbsSel.Font.Size = 11
    Case 2
        Me.mshHead.Font.Size = 15
        Me.tabShow.Font.Size = 15
        mshBody.Font.Size = 15
        tbsSel.Font.Size = 15
    End Select
    mintFont = Index
    Call zlDatabase.SetPara("字体字号", mintFont, glngSys, mlngModule)
    Form_Resize
    Me.Refresh
    
End Sub



Private Sub mnuViewRefresh_Click()
    mstrSelCon = ""
    Select Case tabShow.Tab
        Case 0 '--未发料清单
            SetMnuEnable
            Form_Resize
            Call GetHeadData(0)
            If Me.mshHead.Enabled Then mshHead.SetFocus
            mshHead_EnterCell
        Case 1  '--已发料清单
            SetMnuEnable
            Form_Resize
            Call ReadBillData
            If Me.mshBody.Enabled Then mshBody.SetFocus
    End Select
End Sub


Private Sub mnuViewStatus_Click()
    With mnuViewStatus
        .Checked = Not .Checked
        stbThis.Visible = .Checked
    End With
    
    Form_Resize
End Sub
Private Sub mnuReportItem_Click(Index As Integer)
    Dim strNo As String
    Dim intRecodeSta As Integer
    Dim lng发料部门ID As Long
    Dim lngCol As Long
    Dim lng材料ID As Long
    Dim lng费用ID As Long
    Dim str住院号 As String
    Dim lng单据 As Long
    
    If cboStock.ListIndex < 0 Then
        lng发料部门ID = 0
    Else
        lng发料部门ID = cboStock.ItemData(cboStock.ListIndex)
    End If
        
    
    Select Case tabShow.Tab
    Case 0
        With mshHead
                lng单据 = Decode(.TextMatrix(.Row, mHeadCol.类型), "收费", 24, "记帐单", 25, "记帐表", 26, 0)
                strNo = Trim(.TextMatrix(.Row, mHeadCol.NO))
                lng材料ID = Val(mshBody.TextMatrix(mshBody.Row, mBodyCol.材料ID))
                lng费用ID = Val(mshBody.TextMatrix(mshBody.Row, mBodyCol.费用ID))
                intRecodeSta = Val(mshBody.TextMatrix(mshBody.Row, mBodyCol.记录状态))
                str住院号 = Trim(mshBody.TextMatrix(mshBody.Row, mBodyCol.住院号))
        End With
    Case Else
        With mshBody
                strNo = Trim(.TextMatrix(.Row, mBodyCol.NO))
                lng单据 = Decode(.TextMatrix(.Row, mBodyCol.类型), "收费", 24, "记帐单", 25, "记帐表", 26, 0)
                lng材料ID = Val(.TextMatrix(.Row, mBodyCol.材料ID))
                lng费用ID = Val(.TextMatrix(.Row, mBodyCol.费用ID))
                intRecodeSta = Val(.TextMatrix(.Row, mBodyCol.记录状态))
                str住院号 = Trim(.TextMatrix(.Row, mBodyCol.住院号))
        End With
    End Select
   
    '2006-04-25:刘兴宏:增加自定义报表发布到模块的功能
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "NO=" & strNo, "记录状态=" & intRecodeSta, "发料部门=" & lng发料部门ID, "单据类型=" & lng单据, "材料=" & lng材料ID, "费用=" & lng费用ID, "住院号=" & str住院号)
End Sub
Private Sub mnuViewToolButton_Click()
    With mnuViewToolButton
        .Checked = Not .Checked
        cbrThis.Bands(1).Visible = .Checked
        mnuViewToolText.Enabled = .Checked
    End With
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intCount As Integer      '工具条索引
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    With tlbThis.Buttons
        If mnuViewToolText.Checked = False Then
            '取消所有的文本标签显示
            For intCount = 1 To .Count
                .Item(intCount).Caption = ""
            Next
        Else
            '让所有的文本标签显示。说明：Tag中放的文本标签
            For intCount = 1 To .Count
                .Item(intCount).Caption = .Item(intCount).Tag
            Next
        End If
    End With
    
    cbrThis.Bands(1).MinHeight = tlbThis.Height
    
    Form_Resize
End Sub

Private Sub mshBody_DblClick()
    '
    Dim strNo As String
    Dim str单据 As String
    Dim int操作 As Integer
    
    If mnuEditOutPay.Visible = False Then Exit Sub
      
    If tabShow.Tab <> 1 Then
        mlngCountSel = 0
        Exit Sub
    End If
    With mshBody
        strNo = Trim(.TextMatrix(.Row, mBodyCol.NO))
        
        If strNo = "" Then Exit Sub
        int操作 = Val(.TextMatrix(.Row, mBodyCol.可操作))
        If int操作 <> 1 Then
            If int操作 = -99 Then
                ShowMsgBox "该记录已经被转入历史数据,不能进行退料处理!"
            End If
            Exit Sub
        End If
        
        str单据 = Decode(Trim(.TextMatrix(.Row, mHeadCol.类型)), "收费", 24, "记帐单", 25, 26)
        
        If Trim(.TextMatrix(.Row, mBodyCol.状态)) <> "√" Then
            .TextMatrix(.Row, mBodyCol.状态) = "√"
            mlngCountSel = mlngCountSel + 1

        Else
            .TextMatrix(.Row, mBodyCol.状态) = ""
            mlngCountSel = mlngCountSel - 1
        End If
        If mlngCountSel < 0 Then mlngCountSel = 0
    End With
    SetMnuEnable
End Sub

Private Sub mshBody_EnterCell()
    
    If tabShow.Tab <> 1 Then
        Exit Sub
    End If
    With mshBody
        .ForeColorSel = .CellForeColor
    End With
    SetMnuEnable
    
End Sub

Private Sub mshBody_GotFocus()
    SetGrdSelBackColor mshBody
End Sub

Private Sub mshBody_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 32 Then  '空格
        mshBody_DblClick
    End If
End Sub

Private Sub mshBody_LostFocus()
    With mshBody
        .ForeColorSel = .CellForeColor
    End With
End Sub

Private Sub mshBody_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If tabShow.Tab <> 1 Then Exit Sub
    PopupMenu mnuEdit
    
End Sub

Private Sub mshHead_DblClick()
    Dim strNo As String
    Dim str单据 As String
    If mnuEditPay.Visible = False Then Exit Sub
    
    With mshHead
        strNo = Trim(.TextMatrix(.Row, mHeadCol.NO))
        If strNo = "" Then Exit Sub
        str单据 = Decode(Trim(.TextMatrix(.Row, mHeadCol.类型)), "收费", 24, "记帐单", 25, 26)
        
        If Trim(.TextMatrix(.Row, mHeadCol.标志)) <> "√" Then
            .TextMatrix(.Row, mHeadCol.标志) = "√"
            '往数据串中加入值
            mstrSelCon = mstrSelCon & "||" & strNo & ":" & str单据
        Else
            mstrSelCon = Replace(mstrSelCon, "||" & strNo & ":" & str单据, "")
            .TextMatrix(.Row, mHeadCol.标志) = ""
        End If
    End With
    If tbsSel.SelectedItem.Key <> "K1" Then
        '必数据进行刷新
        tbsSel_Click
    End If
    SetMnuEnable
End Sub
Private Function SelAndClearAllPlay(ByVal blnSel As Boolean) As Boolean
    '选择或清除发料的相关信息
    Dim strNo As String
    Dim str单据 As String
    Dim i As Long
    If mnuEditPay.Visible = False Then Exit Function
    err = 0: On Error GoTo ErrHand:
    With mshHead
        mstrSelCon = ""
        For i = 1 To .Rows - 1
            strNo = Trim(.TextMatrix(i, mHeadCol.NO))
            If strNo = "" Then Exit For
            If blnSel Then
                str单据 = Decode(Trim(.TextMatrix(i, mHeadCol.类型)), "收费", 24, "记帐单", 25, 26)
                .TextMatrix(i, mHeadCol.标志) = "√"
                '往数据串中加入值
                mstrSelCon = mstrSelCon & "||" & strNo & ":" & str单据
            Else
                .TextMatrix(i, mHeadCol.标志) = ""
            End If
       Next
    End With
    SelAndClearAllPlay = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function SelAndClearAllOutPlay(ByVal blnSel As Boolean) As Boolean
    '功能:选择或清除所有的已发料部分
    Dim strNo As String
    Dim str单据 As String
    Dim int操作 As Integer
    Dim i As Long
    If mnuEditOutPay.Visible = False Then Exit Function
    err = 0: On Error GoTo ErrHand:
    With mshBody
        mlngCountSel = 0
        For i = 1 To .Rows - 1
            strNo = Trim(.TextMatrix(i, mBodyCol.NO))
            If strNo = "" Then Exit For
            int操作 = Val(.TextMatrix(i, mBodyCol.可操作))
            If int操作 = 1 And blnSel Then
                .TextMatrix(i, mBodyCol.状态) = "√"
                mlngCountSel = mlngCountSel + 1
            Else
                .TextMatrix(i, mBodyCol.状态) = ""
            End If
        Next
    End With
    SelAndClearAllOutPlay = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub SetMnuEnable()
    Dim blnHead As Boolean
    Dim blnBack As Boolean  '退料模式
    Dim blnData As Boolean '是否存在数据
    Dim blnSelData As Boolean '存在被选择的数据
    Dim blnPrint As Boolean
    Dim blnSelAndClear As Boolean
    blnBack = tabShow.Tab <> 0
           
    If blnBack Then
        blnSelData = mlngCountSel > 0
        blnHead = Trim(mshBody.TextMatrix(mshBody.Row, mBodyCol.NO)) <> ""
        blnData = Trim(mshBody.TextMatrix(1, mBodyCol.NO)) <> ""
        blnPrint = Val(mshBody.TextMatrix(mshBody.Row, mBodyCol.记录状态)) Mod 3 = 2
        blnSelAndClear = mnuEditOutPay.Visible And blnData
    Else
        blnSelData = mstrSelCon <> ""
        blnHead = Trim(mshHead.TextMatrix(mshHead.Row, mHeadCol.NO)) <> ""
        blnData = Trim(mshHead.TextMatrix(1, mHeadCol.NO)) <> ""
        blnSelAndClear = mnuEditPay.Visible And blnData
    End If
    
    Chk清单.Enabled = blnBack
    mnuEditClear.Enabled = blnSelAndClear
    mnuEditSelAll.Enabled = blnSelAndClear
    mnuEditPay.Enabled = blnHead And blnSelData And Not blnBack
    mnuEditOutPay.Enabled = blnHead And blnSelData And blnBack
    
    mnuFilePreview.Enabled = blnData
    mnuFilePrint.Enabled = blnData
    mnuFileExcel.Enabled = blnData
    
    mnuFileRestore.Enabled = blnBack And blnData And blnPrint
    mnuFileBillprint.Enabled = blnBack And blnData And Not blnPrint

    
    tlbThis.Buttons("打印").Enabled = blnData
    tlbThis.Buttons("预览").Enabled = blnData
    
    tlbThis.Buttons("发料").Enabled = mnuEditPay.Enabled
    tlbThis.Buttons("退料").Enabled = mnuEditOutPay.Enabled
    mshHead.Visible = Not blnBack
    tbsSel.Visible = Not blnBack
End Sub

Private Sub mshHead_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 Then  '空格
        mshHead_DblClick
    End If
End Sub

Private Sub mshHead_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        If Button <> 2 Then Exit Sub
        PopupMenu mnuEdit
End Sub

Private Sub tlbThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "预览"
            mnuFilePreView_Click
        Case "打印"
            mnuFilePrint_Click
        Case "发料"
            mnuEditPay_Click
        Case "退料"
            mnuEditOutPay_Click
        Case "过滤"
            mnuViewFind_Click
        Case "帮助"
            mnuHelpTitle_Click
        Case "退出"
            mnufileexit_Click
    End Select
End Sub

Private Sub mshHead_EnterCell()
    With mshHead
        If .TextMatrix(.Row, mHeadCol.NO) = "" Then SetMnuEnable: Exit Sub
        .ForeColorSel = .CellForeColor
        tbsSel_Click
        SetMnuEnable
    End With
End Sub

Private Sub mshHead_GotFocus()
    SetGrdSelBackColor mshHead
End Sub

Private Sub mshHead_LostFocus()
  With mshHead
        If .TextMatrix(.Row, mHeadCol.NO) = "" Then Exit Sub
        .ForeColorSel = .CellForeColor
    End With
End Sub

Private Sub PicLine_S_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    msngOldY = y
End Sub

Private Sub PicLine_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '分割条设置
    
    If Button <> 1 Then Exit Sub
    
    With PicLine_S
        If .Top + y < 2000 Then Exit Sub
        If .Top + y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + y - msngOldY
    End With
    
    With mshHead
        .Height = PicLine_S.Top - .Top
    End With
    With tbsSel
        .Top = PicLine_S.Top + PicLine_S.Height + 10
    End With
    With mshBody
        .Top = tbsSel.Top + tbsSel.Height + 10
        .Height = ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
        .Width = mshHead.Width
    End With
End Sub
Private Sub PicLine_S_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    msngOldY = 0
End Sub

Private Sub SetGrdColHead(ByVal IntStyle As Integer, Optional ByVal blnInitColHead As Boolean = True)
    Dim intCol As Integer
    '--设置各列表控件的格式--
    
    Select Case IntStyle
    Case 1
        With mshHead
             If blnInitColHead Then
                .Clear
                .Rows = 2
                .Cols = mHeadCol.Cols
                .TextMatrix(0, mHeadCol.标志) = "标志"
                .TextMatrix(0, mHeadCol.类型) = "类型"
                .TextMatrix(0, mHeadCol.科室) = "科室"
                .TextMatrix(0, mHeadCol.单据) = "单据"
                .TextMatrix(0, mHeadCol.收费) = "收费"
                .TextMatrix(0, mHeadCol.配料人) = "配料人"
                .TextMatrix(0, mHeadCol.NO) = "NO"
                .TextMatrix(0, mHeadCol.姓名) = "姓名"
                .TextMatrix(0, mHeadCol.床号) = "床号"
                .TextMatrix(0, mHeadCol.住院号) = "住院号"
                .TextMatrix(0, mHeadCol.金额) = "金额"
                .TextMatrix(0, mHeadCol.日期) = "日期"
                .TextMatrix(0, mHeadCol.可操作) = "可操作"
                .TextMatrix(0, mHeadCol.说明) = "说明"
            End If
            
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
            Next
            .ColWidth(mHeadCol.单据) = 0
            .ColWidth(mHeadCol.收费) = 0
            .ColWidth(mHeadCol.配料人) = 0
            
            If RestoreFlexState(Me.mshHead, Me.Caption) = False Then
                .ColWidth(mHeadCol.类型) = 600
                .ColWidth(mHeadCol.标志) = 500
                .ColWidth(mHeadCol.科室) = 1500
                .ColWidth(mHeadCol.NO) = 800
                .ColWidth(mHeadCol.姓名) = 800
                .ColWidth(mHeadCol.金额) = 1000
                .ColWidth(mHeadCol.日期) = 1500
                .ColWidth(mHeadCol.可操作) = 0
                .ColWidth(mHeadCol.说明) = 1500
            End If
            .ColAlignment(mHeadCol.类型) = 4
            .ColAlignment(mHeadCol.标志) = 4
            .ColAlignment(mHeadCol.科室) = 1
            .ColAlignment(mHeadCol.单据) = 0
            .ColAlignment(mHeadCol.收费) = 0
            .ColAlignment(mHeadCol.配料人) = 0
            .ColAlignment(mHeadCol.NO) = 4
            .ColAlignment(mHeadCol.金额) = 7
            .ColAlignment(mHeadCol.姓名) = 4
            .ColAlignment(mHeadCol.日期) = 4
            .ColAlignment(mHeadCol.可操作) = 0
            .ColAlignment(mHeadCol.说明) = 1
        End With
        '恢复列设置
       ' Call RestoreFlexState(mshHead, Me.Name & "\" & TabShow.Tab)
    Case 2
        '明细表格格式
        With mshBody
               If blnInitColHead Then
                    .Clear
                    .Rows = 2
                    .Cols = mBodyCol.Cols
                    .TextMatrix(0, mBodyCol.费用ID) = "费用id"
                    .TextMatrix(0, mBodyCol.材料ID) = "材料id"
                    .TextMatrix(0, mBodyCol.在用分批) = "在用分批"
                    .TextMatrix(0, mBodyCol.批次) = "批次"
                    .TextMatrix(0, mBodyCol.科室) = "科室"
                    .TextMatrix(0, mBodyCol.开单医生) = "开单医生"
                    .TextMatrix(0, mBodyCol.状态) = "标志"
                    .TextMatrix(0, mBodyCol.类型) = "类型"
                    .TextMatrix(0, mBodyCol.NO) = "NO"
                    .TextMatrix(0, mBodyCol.姓名) = "姓名"
                    .TextMatrix(0, mBodyCol.床号) = "床号"
                    .TextMatrix(0, mBodyCol.住院号) = "住院号"
                    .TextMatrix(0, mBodyCol.卫材名称) = "卫材名称"
                    .TextMatrix(0, mBodyCol.规格) = "规格"
                    .TextMatrix(0, mBodyCol.批号) = "批号"
                    .TextMatrix(0, mBodyCol.单位) = "单位"
                    .TextMatrix(0, mBodyCol.换算系数) = "换算系数"
                    .TextMatrix(0, mBodyCol.付数) = "付数"
                    .TextMatrix(0, mBodyCol.数量) = "数量"
                    .TextMatrix(0, mBodyCol.原始数量) = "原始数量"
                    .TextMatrix(0, mBodyCol.已退数) = "已退数"
                    .TextMatrix(0, mBodyCol.准退数) = "准退数"
                    .TextMatrix(0, mBodyCol.退料数) = "退料数"
                    .TextMatrix(0, mBodyCol.单价) = "单价"
                    .TextMatrix(0, mBodyCol.金额) = "金额"
                    .TextMatrix(0, mBodyCol.库存数) = "库存数"
                    .TextMatrix(0, mBodyCol.记帐员) = "记帐员"
                    .TextMatrix(0, mBodyCol.发料人) = "发料人"
                    .TextMatrix(0, mBodyCol.可操作) = "可操作 "
                    .TextMatrix(0, mBodyCol.记录状态) = "记录状态"
                    
                    .TextMatrix(0, mBodyCol.发料人) = "发料人"
                    .TextMatrix(0, mBodyCol.发生时间) = "发生时间"
            End If
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
            Next
            
            .ColWidth(mBodyCol.费用ID) = 0
            .ColWidth(mBodyCol.材料ID) = 0
            .ColWidth(mBodyCol.批次) = 0
            .ColWidth(mBodyCol.换算系数) = 0
            .ColWidth(mBodyCol.在用分批) = 0
            .ColWidth(mBodyCol.可操作) = 0
            .ColWidth(mBodyCol.记录状态) = 0
            .ColWidth(mBodyCol.原始数量) = 0
            
            Dim bytK3 As Byte   '当前为汇总
            Dim byt记帐表 As Byte '存在记帐表
            Dim byt发料模式 As Byte
            
            byt发料模式 = IIf(tabShow.Tab = 0, 0, 1)
            If byt发料模式 = 1 Then
                bytK3 = 1
                byt记帐表 = 1
            Else
                bytK3 = IIf(tbsSel.SelectedItem.Key = "K3", 0, 1)
                byt记帐表 = IIf(InStr(1, mstrSelCon, ":26") <> 0, 1, 0)
                If Trim(mshHead.TextMatrix(mshHead.Row, mHeadCol.单据)) <> "" And byt记帐表 = 0 Then
                    byt记帐表 = IIf(Val(mshHead.TextMatrix(mshHead.Row, mHeadCol.单据)) = 26, 1, 0)
                End If
            End If
            
              If tabShow.Tab = 0 Then
                    '未发
                    If tbsSel.SelectedItem.Key = "K1" Then '明细
                        .Tag = "_明细"
                    ElseIf tbsSel.SelectedItem.Key = "K2" Then '选择单据明细
                        .Tag = "_选择"
                    Else
                        .Tag = "_汇总"
                    End If
              Else
                    '已发
                    .Tag = "_已发"
              End If
              If RestoreFlexState(mshBody, Me.Caption) = False Then
                    .ColWidth(mBodyCol.科室) = 1400 * bytK3
                    .ColWidth(mBodyCol.开单医生) = 1000 * bytK3
                    .ColWidth(mBodyCol.状态) = 800 * bytK3 * byt记帐表
                    .ColWidth(mBodyCol.类型) = 800 * bytK3 * byt记帐表
                    .ColWidth(mBodyCol.NO) = 1000 * bytK3
                    .ColWidth(mBodyCol.姓名) = 1000 * bytK3 * byt记帐表
                    .ColWidth(mBodyCol.床号) = 800 * bytK3 * byt记帐表
                    .ColWidth(mBodyCol.住院号) = 800 * bytK3 * byt记帐表
                    .ColWidth(mBodyCol.卫材名称) = 1600
                    .ColWidth(mBodyCol.规格) = 1400
                    .ColWidth(mBodyCol.批号) = 1000
                    .ColWidth(mBodyCol.单位) = 800
                    .ColWidth(mBodyCol.付数) = 800
                    .ColWidth(mBodyCol.数量) = 800
                    .ColWidth(mBodyCol.已退数) = 800 * bytK3 * byt发料模式
                    .ColWidth(mBodyCol.准退数) = 800 * bytK3 * byt发料模式
                    .ColWidth(mBodyCol.退料数) = 800 * bytK3 * byt发料模式
                    .ColWidth(mBodyCol.单价) = 1000
                    .ColWidth(mBodyCol.金额) = 1000
                    .ColWidth(mBodyCol.库存数) = 1000
                    .ColWidth(mBodyCol.记帐员) = 800 * bytK3
                    .ColWidth(mBodyCol.发料人) = 800 * bytK3
                    .ColWidth(mBodyCol.发生时间) = 1400 * bytK3
             End If
                    
            .ColAlignment(mBodyCol.科室) = 1
            .ColAlignment(mBodyCol.开单医生) = 4
            .ColAlignment(mBodyCol.状态) = 4
            .ColAlignment(mBodyCol.类型) = 4
            .ColAlignment(mBodyCol.NO) = 4
            .ColAlignment(mBodyCol.姓名) = 4
            .ColAlignment(mBodyCol.床号) = 4
            .ColAlignment(mBodyCol.住院号) = 4
            .ColAlignment(mBodyCol.卫材名称) = 1
            .ColAlignment(mBodyCol.规格) = 1
            .ColAlignment(mBodyCol.批号) = 4
            .ColAlignment(mBodyCol.单位) = 1
            .ColAlignment(mBodyCol.付数) = 7
            .ColAlignment(mBodyCol.数量) = 7
            .ColAlignment(mBodyCol.已退数) = 7
            .ColAlignment(mBodyCol.准退数) = 7
            .ColAlignment(mBodyCol.退料数) = 7
            .ColAlignment(mBodyCol.单价) = 7
            .ColAlignment(mBodyCol.金额) = 7
            .ColAlignment(mBodyCol.库存数) = 7
            .ColAlignment(mBodyCol.记帐员) = 4
            .ColAlignment(mBodyCol.发料人) = 4
            .ColAlignment(mBodyCol.发生时间) = 4
        End With
    End Select
End Sub
Private Sub GetHeadData(ByVal int单据 As Integer)
    Dim rsTemp As New ADODB.Recordset
    Dim strCon As String        '条件串
    Dim lng发料部门ID As Long
    Dim n As Integer
    
    If tabShow.Tab = 1 Then Exit Sub
    
    lng发料部门ID = cboStock.ItemData(cboStock.ListIndex)
    
    '24-收费处方发料；25-记帐单处方发料；26-记帐表处方发料；
    
    If mint单据 = 0 Then
        strCon = " And A.单据 In (24,25,26)" '门诊及住院所有单据
    Else
        If mint单据 = 1 Then
            strCon = " And A.单据 In (24,25) And A.主页ID Is NULL " '门诊划价及门诊记帐
        Else
            strCon = " And A.单据 IN (25,26) And A.主页ID Is Not NULL " '住院记帐
        End If
    End If
    
    '标志,    类型,    科室,    单据,    收费,    配料人,    NO  ,    姓名,床号,住院号,日期,    可操作,    说明,
    '未发部分
    gstrSQL = " Select '' 标志, 类型, 科室,单据,已收费,配药人 配料人,NO,姓名,床号,住院号,ltrim(rtrim(to_char(零售金额," & mOraFMT.FM_金额 & ")))  AS 金额,日期,可操作,说明 " & _
              " From (" & _
              "     Select A.优先级,A.类型,D.名称 as 科室,A.单据,A.已收费,A.配药人,A.NO ,decode(a.单据,26,'',A.姓名) 姓名,Max(decode(a.单据,26,'' ,decode(H.门诊标志,2,H.床号,''))) as 床号,max(decode(a.单据,26,'' ,decode(H.门诊标志,2,H.标识号,''))) as 住院号,Sum(C.零售金额) 零售金额,A.日期,A.可操作,A.说明" & _
              "     From ( " & _
              "             Select B.门诊号,B.住院号,A.优先级,A.填制日期,Decode(A.单据,24,'收费',25,'记帐单',26,'记帐表') 类型,A.单据,A.已收费,'' 配药人,A.No,A.姓名,To_Char(A.填制日期,'yyyy-MM-dd hh24:mi:ss') 日期,1 可操作,' ' 说明 " & _
              "             From 未发药品记录 A,病人信息 B" & _
              "             Where A.病人ID=B.病人ID(+) and  nvl(a.已收费,0)=1 and  (A.库房ID=[9] Or A.库房ID Is NULL) " & _
              "                     AND A.填制日期 between [7] and [8]" & _
                                  IIf(mlng病人id = 0, "", " And A.病人id=[1]") & _
                                  IIf(mstrStartNo = "", "", " And A.NO >= [2]") & _
                                  IIf(mstrEndNo = "", "", " And A.NO <= [3]") & _
                                  IIf(mstr病人姓名 = "", "", " And A.姓名 like [4]") & _
                                  IIf(mstr单据 = "", "", " And A.单据 in (" & mstr单据 & ")") & _
              "                     " & strCon & _
              "             ) A,药品收发记录 C,病人费用记录 H,部门表 D" & _
              "     Where A.单据=C.单据 and nvl(c.发药方式,0)<>-1 and C.费用id=H.id   and H.开单部门ID=D.id(+) And A.NO=C.NO And C.审核人 Is NULL And MOD(C.记录状态,3)=1" & _
              "         " & IIf(mlng科室id = 0, "", " And H.开单部门id+0=[5]") & _
              "         " & IIf(Val(mstr住院号) = 0, "", " And H.标识号=[6]") & _
              "     GROUP BY A.优先级,A.类型,D.名称,A.单据,A.已收费,A.配药人,A.No,A.姓名,A.日期,A.可操作,A.说明)"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng病人id, mstrStartNo, mstrEndNo, mstr病人姓名, mlng科室id, mstr住院号, CDate(mstrStartDate), CDate(mstrEndDate), lng发料部门ID)
        
    If rsTemp.RecordCount <> 0 Then
        Set mshHead.DataSource = rsTemp
        Call SetGrdColHead(1, False)
    Else
        Call SetGrdColHead(1)
        Call SetGrdColHead(2)
    End If
    
    '如果是药品窗口传入，则定位到传入的单据号
    If mblnTrans Then
        With mshHead
            For n = 1 To .Rows - 1
                If Trim(.TextMatrix(n, mHeadCol.NO)) = mstrNo Then
                    .Row = n
                    Call mshHead_EnterCell
                    .TopRow = n
                    Exit For
                End If
            Next
        End With
    End If
End Sub
Private Function GetSelCon(Optional strAliaName As String = "A") As String
    '获取被选择的条件
    Dim strArr(0 To 1)
    Dim strTemp As String
    Dim strCon(0 To 2) As String    '分别求条件
    Dim i As Integer
    
    If mstrSelCon = "" Then Exit Function
    
    strArr(0) = Split(Mid(mstrSelCon, 3), "||")
    For i = 0 To UBound(strArr(0))
        strArr(1) = Split(strArr(0)(i), ":")
        Select Case strArr(1)(1)
        Case 24
            strCon(0) = strCon(0) & ",'" & strArr(1)(0) & "'"
        Case 25
            strCon(1) = strCon(1) & ",'" & strArr(1)(0) & "'"
        Case Else
            strCon(2) = strCon(2) & ",'" & strArr(1)(0) & "'"
        End Select
    Next
    
    strTemp = ""
    If strCon(0) <> "" Then
        strTemp = " Or (" & IIf(strAliaName = "", "", strAliaName & ".") & "NO in (" & Mid(strCon(0), 2) & " ) And " & IIf(strAliaName = "", "", strAliaName & ".") & "单据=24) "
    End If
    If strCon(1) <> "" Then
        strTemp = strTemp & " Or (" & IIf(strAliaName = "", "", strAliaName & ".") & "NO in (" & Mid(strCon(1), 2) & " ) And " & IIf(strAliaName = "", "", strAliaName & ".") & "单据=25) "
    End If
    If strCon(2) <> "" Then
        strTemp = strTemp & " Or (" & IIf(strAliaName = "", "", strAliaName & ".") & "NO in (" & Mid(strCon(2), 2) & " ) And " & IIf(strAliaName = "", "", strAliaName & ".") & "单据=26) "
    End If
    If strTemp <> "" Then
        strTemp = " (" & Mid(strTemp, 4) & ") "
    End If
    GetSelCon = strTemp
End Function

Private Function GetBodyQurysSQL(Optional ByVal strTbsSelKey As String = "") As ADODB.Recordset
    '功能:获取查询明细数据
    '参数:
    Dim lng发料部门ID As Long
    Dim strFields  As String
    Dim strCon As String
    Dim strKey As String
    Dim strTableName As String
    Dim rsTemp As New ADODB.Recordset
    Dim int单据 As Long
    Dim strNo As String
    
    lng发料部门ID = cboStock.ItemData(cboStock.ListIndex)
    
    
    Select Case mintUnit
    Case 0  '散装单位
         strFields = "C.计算单位 单位,D.换算系数,ltrim(to_char(B.付数,'9999999999')) 付数,B.实际数量 as 原始数量,ltrim(to_char(B.实际数量 ," & mOraFMT.FM_数量 & "  )) 数量,ltrim(to_char(B.零售价," & mOraFMT.FM_零售价 & ")) 单价,trim(to_char(K.实际数量," & mOraFMT.FM_数量 & "))  库存数, "
    Case Else
         strFields = "D.包装单位 单位,D.换算系数,ltrim(to_char(B.付数,'9999999999')) 付数,B.实际数量 as 原始数量,ltrim(to_char(B.实际数量/D.换算系数," & mOraFMT.FM_数量 & ")) 数量,ltrim(to_char(B.零售价*D.换算系数," & mOraFMT.FM_零售价 & ")) 单价,trim(to_char(K.实际数量/D.换算系数," & mOraFMT.FM_数量 & "))  库存数, "
    End Select
    
    
    If tabShow.Tab = 0 Then
        '未发料模式
        If strTbsSelKey = "" Then
            strKey = tbsSel.SelectedItem.Key
        Else
            strKey = strTbsSelKey
        End If
        Select Case strKey
            Case "K1"           '查单据明细
                    strTableName = " 药品收发记录 B,病人费用记录 H"
                    int单据 = Val(Decode(mshHead.TextMatrix(mshHead.Row, mHeadCol.类型), "收费", "24", "记帐单", "25", 26))
                    strNo = mshHead.TextMatrix(mshHead.Row, mHeadCol.NO)
                    If strNo <> "" Then
                        If zlDatabase.NOMoved("药品收发记录", strNo, "单据=", int单据) Then
                            strTableName = " H药品收发记录 B,H病人费用记录 H"
                        End If
                    End If
                    gstrSQL = "" & _
                        "  SELECT DISTINCT 0 as 可操作,B.记录状态,B.id,B.费用ID as 费用id, B.药品ID as 材料id,NVL(B.批次,0) 批次,NVL(D.在用分批,0) 在用分批," & _
                        "       T.名称 科室,H.开单人 开单医生,'' 标志 ,'' 类型,B.NO,H.姓名,H.床号,H.标识号 住院号," & _
                        "      '['||C.编码||']'||C.名称  卫材名称,H.序号,DECODE(C.规格,NULL,C.产地,DECODE(C.产地,NULL,C.规格,C.规格||'|'||C.产地)) 规格," & _
                        "      DECODE(B.批号,NULL,'',B.批号)||DECODE(B.批次,NULL,'',0,'','('||B.批次||')') 批号, " & strFields & _
                        "      0 as 已退数,0 as 准退数,0 as 退料数,B.零售金额 as 金额," & _
                        "      H.开单人 as 记帐员 ,B.填制日期,H.操作员姓名 as 操作员,B.审核人 配料人 " & _
                        " FROM  " & strTableName & ",材料特性 D,收费项目目录 C, " & _
                        "      部门表 S,部门表 T,(Select 库房id,药品id,批次,实际数量 From 药品库存 where 性质=1) K" & _
                        " WHERE D.材料ID=C.ID and nvl(b.发药方式,0)<>-1 " & _
                        "      AND H.开单部门ID=T.ID(+) AND B.药品ID=D.材料ID AND MOD(B.记录状态,3)=1  and B.NO=[1] and B.单据=[3] " & _
                        "      AND S.ID=NVL(B.库房ID,[2]) AND B.费用ID=H.ID " & _
                        "      AND NVL(B.库房ID,[2])+0=[2] AND LTRIM(RTRIM(NVL(B.摘要,'拒发否')))<>'拒发'" & _
                        "      AND B.药品ID=K.药品ID(+) AND NVL(B.库房ID,[2])=K.库房ID(+) AND NVL(B.批次,0)=NVL(K.批次(+),0)  " & _
                        "      AND B.审核人 IS NULL "
                    gstrSQL = gstrSQL & " Order by H.序号,B.药品ID,Nvl(B.批次,0)"
                    
                    
            Case "K2"           '查被选择的单据明细
                strCon = GetSelCon("B")
                strCon = IIf(strCon = "", "1=2", strCon)
                
                gstrSQL = "" & _
                    "  SELECT DISTINCT 0 as 可操作,B.记录状态,B.id,B.费用ID as 费用id, B.药品ID as 材料id,NVL(B.批次,0) 批次,NVL(D.在用分批,0) 在用分批," & _
                    "       T.名称 科室,H.开单人 开单医生,'' 标志 ,'' 类型,B.NO,H.姓名,H.床号,H.标识号 住院号," & _
                    "      '['||C.编码||']'||C.名称  卫材名称,H.序号,DECODE(C.规格,NULL,C.产地,DECODE(C.产地,NULL,C.规格,C.规格||'|'||C.产地)) 规格," & _
                    "      DECODE(B.批号,NULL,'',B.批号)||DECODE(B.批次,NULL,'',0,'','('||B.批次||')') 批号, " & strFields & _
                    "      0 as 已退数,0 as 准退数,0 as 退料数,B.零售金额 as 金额," & _
                    "      H.开单人 as 记帐员 ,B.填制日期,H.操作员姓名 as 操作员,B.审核人 配料人 " & _
                    " FROM 药品收发记录 B,病人费用记录 H,材料特性 D,收费项目目录 C, " & _
                    "      部门表 S,部门表 T,(Select 库房id,药品id,批次,实际数量 From 药品库存 where 性质=1) K" & _
                    " WHERE D.材料ID=C.ID  and nvl(b.发药方式,0)<>-1 " & _
                    "      AND H.开单部门ID=T.ID(+) AND B.药品ID=D.材料ID AND MOD(B.记录状态,3)=1 " & _
                    "      AND S.ID=NVL(B.库房ID,[2]) AND B.费用ID=H.ID " & _
                    "      AND NVL(B.库房ID,[2])+0=[2] AND LTRIM(RTRIM(NVL(B.摘要,'拒发否')))<>'拒发'" & _
                    "      AND B.药品ID=K.药品ID(+) AND NVL(B.库房ID,[2])=K.库房ID(+) AND NVL(B.批次,0)=NVL(K.批次(+),0)  " & _
                    "      AND B.审核人 IS NULL And " & strCon
                gstrSQL = gstrSQL & " Order by B.填制日期,B.NO,H.序号,B.药品ID,Nvl(B.批次,0)"
            Case "K3"           '查被选择的单据的汇总明细
                strCon = GetSelCon("B")
                strCon = IIf(strCon = "", "1=2", strCon)
                
'                If strTbsSelKey <> "" Then
'                    '只获取材料id,商品及规格,批号,批次,数量
'                    gstrSQL = "" & _
'                        "  SELECT DISTINCT 0 as 可操作,0 记录状态,0 as id,0 as 费用ID,B.药品ID as 材料id,NVL(B.批次,0) 批次," & _
'                        "       '['||C.编码||']'||C.名称  卫材名称,DECODE(C.规格,NULL,C.产地,DECODE(C.产地,NULL,C.规格,C.规格||'|'||C.产地)) 规格," & _
'                        "       DECODE(B.批号,NULL,'',B.批号)||DECODE(B.批次,NULL,'',0,'','('||B.批次||')') 批号," & _
'                        "       SUM(nvl(B.实际数量,0)) as 实际数量,SUM(nvl(C.实际数量,0)) as 库存数量" & _
'                        "      0 as 已退数,0 as 准退数,0 as 退料数,sum(nvl(B.零售金额,0)) as 金额" & _
'                        " FROM 药品收发记录 B,收费项目目录 C,(Select * From 药品库存 where 性质=1) K" & _
'                        " WHERE B.药品ID=C.ID and nvl(b.发药方式,0)<>-1 AND MOD(B.记录状态,3)=1 AND NVL(B.库房ID,[2])+0=[2] " & _
'                        "       AND LTRIM(RTRIM(NVL(B.摘要,'拒发否')))<>'拒发'" & _
'                        "      AND B.药品ID=K.药品ID(+) AND NVL(B.库房ID,[2])=K.库房ID(+) AND NVL(B.批次,0)=NVL(K.批次(+),0)  " & _
'                        "      AND B.审核人 IS NULL And " & strCon & _
'                        " Group by B.药品id,B.批号,b.批次,C.编码,C.名称,DECODE(C.规格,NULL,C.产地,DECODE(C.产地,NULL,C.规格,C.规格||'|'||C.产地))" & _
'                        "       "
'
'                Else
                    
                    Select Case mintUnit
                    Case 0  '散装单位
                         strFields = "C.计算单位 单位,max(D.换算系数) 换算系数,ltrim(to_char(sum(nvl(B.付数,1)),'9999999999')) 付数,sum(nvl(B.实际数量,0)) as 原始数量,ltrim(to_char(sum(nvl(B.实际数量,0))," & mOraFMT.FM_数量 & ")) 数量,ltrim(to_char(sum(nvl(B.零售价,0))," & mOraFMT.FM_零售价 & ")) 单价,trim(to_char(sum(nvl(K.实际数量,0))," & mOraFMT.FM_数量 & "))  库存数, "
                    Case Else
                         strFields = "D.包装单位 单位,max(D.换算系数) 换算系数,ltrim(to_char(sum(nvl(B.付数,0)),'9999999999')) 付数,sum(nvl(B.实际数量,0)) as 原始数量,ltrim(to_char(sum(nvl(B.实际数量/D.换算系数,0))," & mOraFMT.FM_数量 & ")) 数量,ltrim(to_char(sum(nvl(B.零售价*D.换算系数,0))," & mOraFMT.FM_零售价 & ")) 单价,trim(to_char(sum(nvl(K.实际数量/D.换算系数,0))," & mOraFMT.FM_数量 & "))  库存数, "
                    End Select
                                  
                                  
                    gstrSQL = "" & _
                        "  SELECT DISTINCT 0 as 可操作,0 记录状态,0 id,0 as 费用id, B.药品ID as 材料id,NVL(B.批次,0) 批次,'' 在用分批," & _
                        "       '' 科室,'' 开单医生,'' 标志 ,'' 类型,'' NO,'' 姓名,'' 床号,''  住院号," & _
                        "      '['||C.编码||']'||C.名称  卫材名称,0 序号,DECODE(C.规格,NULL,C.产地,DECODE(C.产地,NULL,C.规格,C.规格||'|'||C.产地)) 规格," & _
                        "      DECODE(B.批号,NULL,'',B.批号)||DECODE(B.批次,NULL,'',0,'','('||B.批次||')') 批号, " & strFields & _
                        "      0 as 已退数,0 as 准退数,0 as 退料数,sum(nvl(B.零售金额,0)) as 金额," & _
                        "      '' as 记帐员 ,'' 填制日期,''  操作员,'' 配料人 " & _
                        " FROM 药品收发记录 B,材料特性 D,收费项目目录 C, " & _
                        "      病人费用记录 H,部门表 S,部门表 T,(Select 库房id,药品id,批次,实际数量 From 药品库存 where 性质=1) K" & _
                        " WHERE D.材料ID=C.ID  and nvl(b.发药方式,0)<>-1" & _
                        "      AND H.开单部门ID=T.ID(+) AND B.药品ID=D.材料ID AND MOD(B.记录状态,3)=1 " & _
                        "      AND S.ID=NVL(B.库房ID,[2]) AND B.费用ID=H.ID " & _
                        "      AND NVL(B.库房ID,[2])+0=[2] AND LTRIM(RTRIM(NVL(B.摘要,'拒发否')))<>'拒发'" & _
                        "      AND B.药品ID=K.药品ID(+) AND NVL(B.库房ID,[2])=K.库房ID(+) AND NVL(B.批次,0)=NVL(K.批次(+),0)  " & _
                        "      AND B.审核人 IS NULL And " & strCon & _
                        " Group by B.药品id,b.批次,C.编码,C.名称,DECODE(C.规格,NULL,C.产地,DECODE(C.产地,NULL,C.规格,C.规格||'|'||C.产地))," & _
                        "       B.批号,B.批次," & IIf(mintUnit = 0, "C.计算单位", "D.包装单位")
               ' End If
        End Select
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, lng发料部门ID, int单据)
        Set GetBodyQurysSQL = rsTemp
        Exit Function
    End If
    
     
    Dim strCon1 As String
    strCon1 = ""
    '24-收费处方发料；25-记帐单处方发料；26-记帐表处方发料；
    If mint单据 = 0 Then
        strCon = " And S.单据 In (24,25,26)" '门诊及住院所有单据
    Else
        If mint单据 = 1 Then
            strCon = " And S.单据 In (24,25)" '门诊划价及门诊记帐
            strCon1 = " and M.主页ID is null"
        Else
            strCon = " And S.单据 IN (24,25,26)  " '住院记帐
            strCon1 = " and M.主页ID is not null"
        End If
    End If
    If mstr单据 <> "" Then
        strCon = strCon & " and S.单据 in (" & mstr单据 & ") "
    End If
    
    Select Case mintUnit
    Case 0  '散装单位
         strFields = "S.计算单位 单位,D.换算系数,ltrim(to_char(S.付数,'9999999999')) 付数,s.实际数量 as 原始数量,ltrim(to_char(S.实际数量," & mOraFMT.FM_数量 & ")) 数量,ltrim(to_char(S.已退数量," & mOraFMT.FM_数量 & ")) as 已退数,ltrim(to_char(S.已发数量," & mOraFMT.FM_数量 & ")) as 准退数,'' 退料数,ltrim(to_char(S.零售价," & mOraFMT.FM_零售价 & ")) 单价,''  库存数, "
    Case Else
         strFields = "D.包装单位 单位,D.换算系数,ltrim(to_char(S.付数,'9999999999')) 付数,s.实际数量 as 原始数量,ltrim(to_char(S.实际数量/D.换算系数," & mOraFMT.FM_数量 & ")) 数量,ltrim(to_char(S.已退数量/D.换算系数," & mOraFMT.FM_数量 & ")) as 已退数,ltrim(to_char(S.已发数量/D.换算系数," & mOraFMT.FM_数量 & ")) as 准退数,'' 退料数,ltrim(to_char(S.零售价*D.换算系数," & mOraFMT.FM_零售价 & ")) 单价,'' 库存数, "
    End Select
    
    Dim strTemp As String
    Dim blnHistory As Boolean
    blnHistory = zlDatabase.DateMoved(mstrStartDate, , , Me.Caption)
    
    If Chk清单.Value = 0 Then
        '获取已发料或退料的金额
        
        gstrSQL = " SELECT DISTINCT S.id,S.记录状态 ,S.费用ID,s.开单医生,'' 标志,decode(S.单据,24,'收费',25,'记帐单',26,'记帐表' ) as 类型,s.住院号,s.操作员,s.记帐员, S.ID,S.单据,S.药品ID as 材料id,S.NO,S.扣率,P.名称 科室,s.门诊标志,s.床号,s.姓名," & _
            " '['||X.编码||']'||X.名称  卫材名称,NVL(D.在用分批,0) 在用分批,DECODE(x.规格,NULL,x.产地,DECODE(x.产地,NULL,x.规格,x.规格||'|'||x.产地)) 规格," & strFields & _
            "  DECODE(S.批号,NULL,'',S.批号)||DECODE(S.批次,NULL,'',0,'','('||S.批次||')') 批号,NVL(S.批次,0) 批次,S.效期," & _
            "  S.零售金额 金额,S.摘要 说明,S.审核人,TO_CHAR(S.审核日期,'YYYY-MM-DD HH24:MI:SS') 发料时间,s.可操作" & _
            " FROM (    SELECT A.ID,A.NO,A.单据,A.药品ID,A.费用ID,A.批次,A.批号,A.效期,NVL(A.扣率,0) 扣率," & _
            "                   NVL(A.付数,1) 付数,A.实际数量 实际数量,NVL(A.付数,1)*A.实际数量-B.已发数量 已退数量,B.已发数量,A.记录状态," & _
            "                   A.零售价,A.零售金额,A.单量,A.频次,A.用法,A.摘要,A.审核人,A.审核日期,A.对方部门ID,A.库房ID,A.开单医生,A.计算单位,A.住院号,A.操作员,A.记帐员,A.门诊标志,A.床号,A.姓名,A.可操作" & _
            "           FROM(  "
            
             strTemp = "" & _
                "   SELECT A.ID,A.NO,A.药品id,A.序号,A.单据,A.费用ID,A.批次,A.批号,A.效期,nvl(A.扣率,0) 扣率,nvl(A.付数,0) 付数,A.实际数量,A.记录状态," & _
                "        A.零售价,A.零售金额,A.单量,A.频次,A.用法,A.摘要,A.审核人,A.审核日期,A.对方部门id,A.库房ID," & _
                "        m.开单人 as 开单医生,M.计算单位,m.标识号 as 住院号,m.操作员姓名 as 操作员,m.开单人 记帐员,m.门诊标志,m.床号,m.姓名,1 可操作 " & _
                "   FROM 药品收发记录 A,病人费用记录 M" & _
                "   WHERE A.审核人 IS NOT NULL and A.费用id=M.ID  and nvl(a.发药方式,0)<>-1 AND (A.记录状态=1 OR MOD(A.记录状态,3)=0)" & _
                "       AND A.库房ID+0=[9]" & _
                "       AND A.审核日期 BETWEEN [7] AND [8] " & _
                        IIf(mlng病人id = 0, "", " AND M.病人ID+0=[1]") & _
                        IIf(Val(mstr住院号) = 0, "", " AND M.标识号+0=[2]") & _
                        IIf(mstr病人姓名 = "", "", " AND M.姓名 LIKE [3]") & _
                        IIf(mstrStartNo = "", "", " AND A.NO>=[4]") & _
                        IIf(mstrEndNo = "", "", " AND A.NO<=[5]") & _
                        IIf(mlng科室id = 0, "", " And M.执行部门id+0=[6]")
            
            If blnHistory Then
                strTemp = AnalyseHistorySQL(strTemp, "1 可操作", "-99 可操作")
            End If
            gstrSQL = gstrSQL & strTemp & " ) A,(    "
            
            
            strTemp = "" & _
                "               SELECT A.NO,A.单据,A.药品ID,A.序号,SUM(NVL(A.付数,1)*A.实际数量) 已发数量" & _
                "               FROM 药品收发记录 A" & _
                "               WHERE A.审核人 IS NOT NULL and nvl(a.发药方式,0)<>-1 AND A.库房ID+0=[9]" & _
                "                        and (A.NO,单据) in (Select NO,单据 From 药品收发记录 " & _
                "                                            where 审核人 is not null and nvl(发药方式,0)<>-1  AND " & _
                "                                                  (记录状态=1 OR MOD(记录状态,3)=0) and 库房id+0 =[9]" & _
                                                                IIf(mstrStartNo = "", "", " AND NO>=[4]") & _
                                                                IIf(mstrEndNo = "", "", " AND  NO<=[5]") & _
                "                                                   AND 审核日期 BETWEEN [7] AND [8]) " & _
                "               GROUP BY A.NO,A.单据,A.药品ID,A.序号    "
            
            If blnHistory Then
                    strTemp = AnalyseHistorySQL(strTemp)
            End If
            gstrSQL = gstrSQL & strTemp & " ) B" & _
                "           WHERE A.NO = B.NO AND A.单据 = B.单据 AND A.药品ID+0 = B.药品ID AND A.序号 = B.序号 AND B.已发数量<>0) S,"
    
            gstrSQL = gstrSQL & "" & _
                "      部门表 P,材料特性 D,收费项目目录 X" & _
                " WHERE S.药品ID=D.材料ID AND S.对方部门ID+0=P.ID  AND d.材料ID=X.ID" & _
                "       AND (S.记录状态=1 OR MOD(S.记录状态,3)=0) AND S.实际数量*S.付数>S.已退数量 " & _
                "       AND S.审核人 IS NOT NULL AND S.库房ID+0=[9]" & strCon
            gstrSQL = gstrSQL & " Order By S.No,S.单据"
        
    Else
        '清单显示每笔操作过程
        
        gstrSQL = " SELECT DISTINCT  S.id,S.记录状态,S.费用ID,S.开单医生,'' 标志,decode(S.单据,24,'收费',25,'记帐单',26,'记帐表' ) as 类型,s.住院号,s.操作员,s.记帐员,S.ID,S.单据,S.药品ID 材料id,S.NO,S.扣率,P.名称 科室,s.门诊标志,s.床号,s.姓名,'['||X.编码||']'||X.名称  卫材名称," & _
                 "          NVL(D.在用分批,0)  在用分批,DECODE(x.规格,NULL,x.产地,DECODE(x.产地,NULL,x.规格,x.规格||'|'||X.产地)) 规格," & strFields & _
                 "          DECODE(S.批号,NULL,'',S.批号)||DECODE(S.批次,NULL,'',0,'','('||S.批次||')') 批号,NVL(S.批次,0) 批次,S.效期," & _
                 "          S.零售价 单价,S.零售金额 金额,S.摘要 说明,TO_CHAR(S.审核日期,'YYYY-MM-DD HH24:MI:SS') 发料时间,S.审核人,S.审核日期,decode(S.已发数量,0,0,可操作) as 可操作" & _
                 " FROM "
                 
        gstrSQL = gstrSQL & _
                 "      (   SELECT * FROM" & _
                 "              (   SELECT A.ID,A.NO,A.单据,A.药品ID,A.费用ID,A.批次,A.批号,A.效期,NVL(A.扣率,0) 扣率," & _
                 "                          NVL(A.付数,1) 付数,A.实际数量,NVL(A.付数,1)*A.实际数量-B.已发数量 已退数量,B.已发数量,A.记录状态," & _
                 "                          A.零售价 , A.零售金额, A.单量, A.频次, A.用法, A.摘要, A.审核人, A.审核日期, A.对方部门ID, A.库房ID,A.可操作," & _
                 "                          A.开单医生,A.住院号,A.操作员,A.计算单位,A.记帐员,A.门诊标志,A.床号,A.姓名 " & _
                 "                  FROM (  "
                 
        strTemp = "" & _
                "   SELECT A.ID,A.NO,A.单据,A.药品id,A.序号,A.费用ID,A.批次,A.批号,A.效期,nvl(A.扣率,0) 扣率,nvl(A.付数,0) 付数,A.实际数量,A.记录状态," & _
                "        A.零售价,A.零售金额,A.单量,A.频次,A.用法,A.摘要,A.审核人,A.审核日期,A.对方部门id,A.库房ID," & _
                "        m.开单人 as 开单医生,m.标识号 as 住院号,m.操作员姓名 as 操作员,m.计算单位,m.开单人 记帐员,m.门诊标志,m.床号,m.姓名,1 可操作 " & _
                "   FROM 药品收发记录 A,病人费用记录 M" & _
                "   WHERE A.审核人 IS NOT NULL and A.费用id=M.ID  and nvl(a.发药方式,0)<>-1 AND (A.记录状态=1 OR MOD(A.记录状态,3)=0)" & _
                "       AND A.库房ID+0=[9]" & _
                "       AND A.审核日期 BETWEEN [7] AND [8] " & _
                        IIf(mlng病人id = 0, "", " AND M.病人ID+0=[1]") & _
                        IIf(Val(mstr住院号) = 0, "", " AND M.标识号+0=[2]") & _
                        IIf(mstr病人姓名 = "", "", " AND M.姓名 LIKE [3]") & _
                        IIf(mstrStartNo = "", "", " AND A.NO>=[4]") & _
                        IIf(mstrEndNo = "", "", " AND A.NO<=[5]") & _
                        IIf(mlng科室id = 0, "", " And M.执行部门id+0=[6]") & strCon1
                 
            If blnHistory Then
                strTemp = AnalyseHistorySQL(strTemp, "1 可操作", "-99 可操作")
            End If
            
            gstrSQL = gstrSQL & strTemp & " ) A,( "
            strTemp = "" & _
                      "               SELECT A.NO,A.单据,A.药品ID,A.序号,SUM(NVL(A.付数,1)*A.实际数量) 已发数量" & _
                      "               FROM 药品收发记录 A" & _
                      "               WHERE A.审核人 IS NOT NULL and nvl(a.发药方式,0)<>-1 AND A.库房ID+0=[9]" & _
                      "                        and (A.NO,单据) in (Select NO,单据 From 药品收发记录 " & _
                      "                                            where 审核人 is not null and nvl(发药方式,0)<>-1  AND " & _
                      "                                                  (记录状态=1 OR MOD(记录状态,3)=0) and 库房id+0 =[9]" & _
                                                                      IIf(mstrStartNo = "", "", " AND NO>=[4]") & _
                                                                      IIf(mstrEndNo = "", "", " AND  NO<=[5]") & _
                      "                                                   AND 审核日期 BETWEEN [7] AND [8]) " & _
                      "               GROUP BY A.NO,A.单据,A.药品ID,A.序号    "
                     
            If blnHistory Then
                strTemp = AnalyseHistorySQL(strTemp)
            End If
        
            gstrSQL = gstrSQL & strTemp & ") B" & _
                     "              WHERE A.NO = B.NO AND A.单据 = B.单据 AND A.药品ID+0 = B.药品ID AND A.序号 = B.序号)" & _
                     "              UNION"
            strTemp = "" & _
                     "              SELECT A.ID,A.NO,A.单据,A.药品ID,A.费用ID,A.批次,A.批号,A.效期,NVL(A.扣率,0)," & _
                     "                      NVL(A.付数,1) 付数,A.实际数量,0 已退数,0 已发数量,A.记录状态," & _
                     "                      A.零售价 , A.零售金额, A.单量, A.频次, A.用法, A.摘要, A.审核人, A.审核日期, A.对方部门ID, A.库房ID," & _
                     "                      DECODE(A.记录状态,1,1,DECODE(MOD(A.记录状态,3),0,1,MOD(A.记录状态,3)+1)) 可操作," & _
                     "                       m.开单人 as 开单医生,m.标识号 as 住院号,m.操作员姓名 as 操作员,M.计算单位,m.开单人 记帐员,m.门诊标志,m.床号,m.姓名 " & _
                     "              FROM 药品收发记录 A, 病人费用记录 M" & _
                     "              WHERE A.审核人 IS NOT NULL and a.费用id=M.id   and nvl(a.发药方式,0)<>-1 AND NOT (a.记录状态=1 OR MOD(a.记录状态,3)=0)" & _
                     "                      AND A.库房ID+0=[9]" & _
                     "                      AND A.审核日期 BETWEEN [7] AND [8]" & _
                                            IIf(mlng病人id = 0, "", " AND M.病人ID=[1]") & _
                                            IIf(Val(mstr住院号) = 0, "", " AND M.标识号=[2]") & _
                                            IIf(mstr病人姓名 = "", "", " AND M.姓名 LIKE  [3] ") & _
                                            IIf(mstrStartNo = "", "", " AND A.NO>=[4]") & _
                                            IIf(mstrEndNo = "", "", " AND A.NO<=[5]") & _
                                            IIf(mlng科室id = 0, "", " And M.执行部门id+0=[6]") & strCon1
            If blnHistory Then
                strTemp = AnalyseHistorySQL(strTemp, "DECODE(A.记录状态,1,1,DECODE(MOD(A.记录状态,3),0,1,MOD(A.记录状态,3)+1)) 可操作", " -99 可操作")
            End If
            
            
        gstrSQL = gstrSQL & strTemp & " ) S," & _
                 "      部门表 P,材料特性 D,收费项目目录 X " & _
                 " WHERE S.药品ID=D.材料ID AND d.材料ID=x.ID AND S.对方部门ID+0=P.ID" & _
                 " " & strCon & _
                 "    and  S.审核人 IS NOT NULL"
        gstrSQL = gstrSQL & " Order By S.No,S.单据,S.审核日期"
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng病人id, mstr住院号, mstr病人姓名, mstrStartNo, mstrEndNo, mlng科室id, CDate(mstrStartDate), CDate(mstrEndDate), lng发料部门ID)
    Set GetBodyQurysSQL = rsTemp
End Function
Private Function ReadBillData() As Boolean
    Dim RsBody As New ADODB.Recordset
    Dim IntStyle As Integer
        
    '--读取单据内容--
    On Error GoTo ErrHand:
    err = 0
    ReadBillData = False
    
    Set RsBody = GetBodyQurysSQL

    'zlDatabase.OpenRecordset RsBody, gstrSQL, Me.Caption
    
    '绑定相关数据
    If RsBody.RecordCount <> 0 Then
        '   加载数据
        LoadBodyData RsBody
        Call SetGrdColHead(2, False)
    Else
        Call SetGrdColHead(2)
    End If
    ReadBillData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SetGrdColHead(2)
    ReadBillData = False
End Function
Private Function LoadBodyData(ByVal RsBody As ADODB.Recordset) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:加载表体数据
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    err = 0
    On Error GoTo ErrHand:
    LoadBodyData = False
    With mshBody
        .Redraw = False
        .Rows = RsBody.RecordCount + 1
        lngRow = 1
        Do While Not RsBody.EOF
            
            .TextMatrix(lngRow, mBodyCol.费用ID) = NVL(RsBody!费用ID, 0)
            .TextMatrix(lngRow, mBodyCol.材料ID) = NVL(RsBody!材料ID, 0)
            .TextMatrix(lngRow, mBodyCol.在用分批) = NVL(RsBody!在用分批, 0)
            .TextMatrix(lngRow, mBodyCol.批次) = NVL(RsBody!批次, 0)
            .TextMatrix(lngRow, mBodyCol.科室) = NVL(RsBody!科室)
            .TextMatrix(lngRow, mBodyCol.开单医生) = NVL(RsBody!开单医生)
            .TextMatrix(lngRow, mBodyCol.状态) = NVL(RsBody!标志)
            .TextMatrix(lngRow, mBodyCol.类型) = NVL(RsBody!类型)
            .TextMatrix(lngRow, mBodyCol.NO) = NVL(RsBody!NO)
            .TextMatrix(lngRow, mBodyCol.姓名) = NVL(RsBody!姓名)
            .TextMatrix(lngRow, mBodyCol.床号) = NVL(RsBody!床号)
            .TextMatrix(lngRow, mBodyCol.住院号) = NVL(RsBody!住院号)
            .TextMatrix(lngRow, mBodyCol.卫材名称) = NVL(RsBody!卫材名称)
            .TextMatrix(lngRow, mBodyCol.规格) = NVL(RsBody!规格)
            .TextMatrix(lngRow, mBodyCol.批号) = NVL(RsBody!批号)
            .TextMatrix(lngRow, mBodyCol.单位) = NVL(RsBody!单位)
            .TextMatrix(lngRow, mBodyCol.换算系数) = NVL(RsBody!换算系数)
            .TextMatrix(lngRow, mBodyCol.付数) = NVL(RsBody!付数)
            .TextMatrix(lngRow, mBodyCol.数量) = NVL(RsBody!数量)
            
            .TextMatrix(lngRow, mBodyCol.原始数量) = NVL(RsBody!原始数量)
            .TextMatrix(lngRow, mBodyCol.已退数) = NVL(RsBody!已退数)
            .TextMatrix(lngRow, mBodyCol.准退数) = NVL(RsBody!准退数)
            .TextMatrix(lngRow, mBodyCol.退料数) = NVL(RsBody!退料数)
            .TextMatrix(lngRow, mBodyCol.单价) = NVL(RsBody!单价)
            .TextMatrix(lngRow, mBodyCol.金额) = Format(Val(NVL(RsBody!金额)), mFMT.FM_金额)
            .TextMatrix(lngRow, mBodyCol.库存数) = Format(Val(NVL(RsBody!库存数)), mFMT.FM_数量)
            .TextMatrix(lngRow, mBodyCol.记帐员) = NVL(RsBody!操作员)
            If tabShow.Tab = 0 Then
                .TextMatrix(lngRow, mBodyCol.发料人) = NVL(RsBody!配料人)
            Else
                .TextMatrix(lngRow, mBodyCol.发料人) = NVL(RsBody!审核人)
            End If
            .TextMatrix(lngRow, mBodyCol.可操作) = NVL(RsBody!可操作)
            .TextMatrix(lngRow, mBodyCol.记录状态) = NVL(RsBody!记录状态)
            
            If tabShow.Tab <> 0 Then
                .TextMatrix(lngRow, mBodyCol.发生时间) = NVL(RsBody!发料时间)
                .RowData(lngRow) = NVL(RsBody!Id, 0)
                SetGRDCOLOR mshBody, lngRow, IIf(NVL(RsBody!可操作) = 1, 1, NVL(RsBody!记录状态, 0))
            Else
                .TextMatrix(lngRow, mBodyCol.发生时间) = NVL(RsBody!填制日期)
                SetGRDCOLOR mshBody, lngRow, 1
            End If
            lngRow = lngRow + 1
            RsBody.MoveNext
        Loop
        .Redraw = True
    End With
    LoadBodyData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub tabShow_Click(PreviousTab As Integer)
    Dim blnYes As Boolean
    If PreviousTab = tabShow.Tab Then Exit Sub
    
    If mblnExit = True Then
        mblnExit = False
        Exit Sub
    End If
    
    If mstrSelCon <> "" And PreviousTab = 0 Then
        ShowMsgBox "已经有被选择的项目,你是否希望改变选择!", True, blnYes
        If Not blnYes Then
            mblnExit = True
            tabShow.Tab = PreviousTab
            mnuEditPayType.Checked = True
            mnuEditBackType.Checked = False
            Exit Sub
        End If
        mstrSelCon = ""
    
    End If
    If mlngCountSel > 0 And PreviousTab = 1 Then
        ShowMsgBox "已经有被选择的项目,你是否希望改变选择!", True, blnYes
        If Not blnYes Then
            mblnExit = True
            tabShow.Tab = PreviousTab
            mnuEditPayType.Checked = False
            mnuEditBackType.Checked = True
            Exit Sub
        End If
        mlngCountSel = 0
    End If
    
    mblnExit = False
    
    mlngCountSel = 0
    '修改:刘兴宏   Bug:    日期:2008-05-14 11:17:53
    If PreviousTab = 0 Then
        '未发
        Select Case tbsSel.SelectedItem.Key
        Case "K1"   '明细
            mshBody.Tag = "_明细"
        Case "K2"  '选择明细
            mshBody.Tag = "_选择"
        Case "K3"  '汇总发料
            mshBody.Tag = "_汇总"
        End Select
        SaveFlexState mshHead, Me.Caption
    Else
        '已发
        '先保存
        mshBody.Tag = "_已发"
    End If
    If PreviousTab <> -1 Then
        SaveFlexState mshBody, Me.Caption
    End If
    
    Select Case tabShow.Tab
        Case 0 '--未发料清单
            SetMnuEnable
            Form_Resize
            Call GetHeadData(0)
            If Me.mshHead.Enabled Then mshHead.SetFocus
            mshHead_EnterCell
            mnuEditPayType.Checked = True
            mnuEditBackType.Checked = False
        
        Case 1  '--已发料清单
            SetMnuEnable
            Form_Resize
            Call ReadBillData
            If Me.mshBody.Enabled Then mshBody.SetFocus
            mnuEditPayType.Checked = False
            mnuEditBackType.Checked = True
                
    End Select
    SetMnuEnable
    mstrPreSelKey = ""
End Sub

Private Sub tbsSel_Click()
    '未发
    If mstrPreSelKey <> "" Then
        Select Case mstrPreSelKey
        Case "K1"   '明细
            mshBody.Tag = "_明细"
        Case "K2"  '选择明细
            mshBody.Tag = "_选择"
        Case "K3"  '汇总发料
            mshBody.Tag = "_汇总"
        End Select
        SaveFlexState mshBody, Me.Caption
    End If
    
    mstrPreSelKey = tbsSel.SelectedItem.Key
    ReadBillData
End Sub
Private Sub SetGrdSelBackColor(objGrid As Object)
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:网格选择色置换
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    If objGrid Is mshHead Then
        mshHead.BackColorSel = &H8000000C     ' &HC0C0C0
        mshBody.BackColorSel = &HE0E0E0
    ElseIf objGrid Is mshBody Then
        mshHead.BackColorSel = &HE0E0E0
        mshBody.BackColorSel = &H8000000C     '&HC0C0C0
    End If
End Sub

Private Function CheckStock() As Boolean
    Dim dblStock As Double
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    Dim lng材料ID As Long
    Dim lng批次 As Long
    Dim RsBody As New ADODB.Recordset
    
    
    '检查库存
    If mintCheckStock = 0 Then CheckStock = True: Exit Function
    
    Set RsBody = GetBodyQurysSQL("K3")
    'zlDatabase.OpenRecordset RsBody, gstrSQL, "检查库存"
    
    CheckStock = False
    With RsBody
            Do While Not .EOF
                lng材料ID = NVL(!材料ID, 0)
                lng批次 = NVL(!批次, 0)
                
               If lng材料ID <> 0 Then
                        dblStock = NVL(!库存数, 0)
                        
                        If dblStock < NVL(!数量, 0) Then
                            If lng批次 <> 0 Then
                                MsgBox NVL(!卫材名称) & "的批次库存数不够，不能继续发料！", vbInformation, gstrSysName: Exit Function
                            Else
                                Select Case mintCheckStock
                                Case 1
                                    If MsgBox(NVL(!卫材名称) & "的库存数不够，是否继续发料？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                                Case 2
                                    MsgBox NVL(!卫材名称) & "的库存数不够，不能继续发料！", vbInformation, gstrSysName: Exit Function
                                End Select
                            End If
                        End If
               End If
               .MoveNext
            Loop
    End With
    CheckStock = True
End Function

Private Function SendBill() As Boolean
    Dim intRow As Integer
    Dim strDate As String
    Dim mlng发料部门ID As Long
    Dim strNo As String
    Dim int发料方式 As Integer     '1-处方发料;2-批量发料;3-部门发料
    Dim lng单据 As Long
    
    On Error GoTo ErrHand
    err = 0
    SendBill = False
    
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    mlng发料部门ID = cboStock.ItemData(cboStock.ListIndex)
    gcnOracle.BeginTrans
    With mshHead
        int发料方式 = UBound(Split(mstrSelCon, "||"))
        
        int发料方式 = IIf(int发料方式 = 0 Or int发料方式 = 1, 1, 2)
        
        If InStr(1, mstrSelCon, ":26") <> 0 Then
            '部门发料
            int发料方式 = 3
        End If
            
        For intRow = 1 To .Rows - 1
            lng单据 = Decode(.TextMatrix(intRow, mHeadCol.类型), "收费", 24, "记帐单", 25, "记帐表", 26, 0)
            strNo = Trim(.TextMatrix(intRow, mHeadCol.NO))
             If lng单据 <> 0 And strNo <> "" And Trim(.TextMatrix(intRow, mHeadCol.标志)) = "√" Then
                mstrPrintCon = strNo & "||" & lng单据 & "||" & mlng发料部门ID
                '过程参数:库房ID_IN,单据_IN,NO_IN,审核人_IN,配料人_IN,校验人_IN,发料方式_IN,审核日期_IN
                gstrSQL = "zl_材料收发记录_处方发料(" & _
                    mlng发料部门ID & "," & _
                    lng单据 & ",'" & _
                    strNo & "','" & _
                    gstrUserName & "','" & _
                    gstrUserName & "','NULL'," & _
                    int发料方式 & ",to_date('" & _
                    strDate & "','yyyy-MM-dd hh24:mi:ss'))"
                Call zlDatabase.ExecuteProcedure(gstrSQL, (Me.Caption & "-卫生材料发放"))
            End If
        Next
    End With
    
    gcnOracle.CommitTrans
    BillListPrint int发料方式, strDate
  
    SendBill = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub SetGRDCOLOR(ByVal objGrd As Object, ByVal lngRow As Long, ByVal int记录状态 As Integer)
    Dim lngColor As Long
    Dim i As Long
    If int记录状态 = 1 Then
        lngColor = &H80000008
    ElseIf zlCommFun.ZyMod(int记录状态, 3) = 2 Then
         lngColor = vbRed
    Else
        lngColor = vbBlue
    End If
    With mshBody
        For i = 0 To .Cols - 1
            .Row = lngRow
            .Col = i
            .CellForeColor = lngColor
        Next
    End With
End Sub

Private Function LoadInIcon() As Boolean
    '--为各控件装入图标--
    On Error Resume Next
    err = 0
    LoadInIcon = False
    
    '工具栏
    With ImgTbarBlack
        .ImageHeight = 24
        .ImageWidth = 24
        .ListImages.Add , , LoadResPicture("BPREVIEW", vbResIcon)
        .ListImages.Add , , LoadResPicture("BPRINT", vbResIcon)
        .ListImages.Add , , LoadResPicture("BSTOP", vbResIcon)
        .ListImages.Add , , LoadResPicture("BSTART", vbResIcon)
       ' .ListImages.Add , , LoadResPicture("BSEND", vbResIcon)
       ' .ListImages.Add , , LoadResPicture("BSEND", vbResIcon)
        .ListImages.Add , , LoadResPicture("BFILTER", vbResIcon)
        .ListImages.Add , , LoadResPicture("BHELP", vbResIcon)
        .ListImages.Add , , LoadResPicture("BEXIT", vbResIcon)
    End With
    
    With ImgTbarColor
        .ImageHeight = 24
        .ImageWidth = 24
        .ListImages.Add , , LoadResPicture("CPREVIEW", vbResIcon)
        .ListImages.Add , , LoadResPicture("CPRINT", vbResIcon)
        .ListImages.Add , , LoadResPicture("CSTOP", vbResIcon)
        .ListImages.Add , , LoadResPicture("CSTART", vbResIcon)
'        .ListImages.Add , , LoadResPicture("CSEND", vbResIcon)
'        .ListImages.Add , , LoadResPicture("CSEND", vbResIcon)
        .ListImages.Add , , LoadResPicture("CFILTER", vbResIcon)
        .ListImages.Add , , LoadResPicture("CHELP", vbResIcon)
        .ListImages.Add , , LoadResPicture("CEXIT", vbResIcon)
    End With
    
    With tlbThis
        Set .ImageList = ImgTbarBlack
        Set .HotImageList = ImgTbarColor
        
        .Buttons("预览").Image = 1
        .Buttons("打印").Image = 2
        .Buttons("发料").Image = 3
        .Buttons("退料").Image = 4
        .Buttons("过滤").Image = 5
        
        .Buttons("帮助").Image = 6
        .Buttons("退出").Image = 7
    End With
    
    cbrThis.Bands(1).MinHeight = tlbThis.Height
    If err <> 0 Then
        MsgBox "相关资源文件丢失，请与软件开发商联系！", vbInformation, gstrSysName
        Exit Function
    End If
    LoadInIcon = True
End Function


Private Sub mnuHelpAbout_Click()
    '关于
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
    '帮助主题
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    '中联主页
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    '发送反馈
    Call zlMailTo(Me.hwnd)
End Sub
Private Sub subPrint(bytMode As Byte)
    '功能:进行打印,预览和输出到EXCEL
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    Dim strRange As String
    
    If Me.tabShow.Tab = 1 Then
        strRange = "审核日期 " & Format(mstrStartDate, "yyyy年MM月dd日") & "至" & Format(mstrEndDate, "yyyy年MM月dd日")
    Else
        strRange = "填制日期 " & Format(mstrStartDate, "yyyy年MM月dd日") & "至" & Format(mstrEndDate, "yyyy年MM月dd日")
    End If
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = IIf(Me.tabShow.Tab = 0, "未发料清册", "已发料清册")
        
    objRow.Add "时间：" & strRange
    objRow.Add "发料部门：" & cboStock.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
        
    objRow.Add "打印人:" & UserInfo.用户名
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    If Me.ActiveControl Is mshBody Then
        Set objPrint.Body = mshBody
    Else
        Set objPrint.Body = IIf(Me.tabShow.Tab = 1, mshBody, mshHead)
    End If
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub
Private Sub BillListPrint(Optional int发料方式 As Integer = 1, Optional strDate As String = "", Optional IntStyle As Integer = 0)
    '单据或清册打印
    '发料方式:1-处方发料;2-批量发料;3-部门发料
    ' intStyle:0-按发料方式打印,1-单据打印
    Dim bln退料单 As Boolean
    Dim bln已发料清单 As Boolean
    Dim bln单据打印 As Boolean
    
    bln退料单 = InStr(1, mstrPrivs, "退料通知单") <> 0
    bln已发料清单 = InStr(1, gstrPrivs, "打印已发料清单") <> 0
    bln单据打印 = InStr(1, gstrPrivs, "单据打印") <> 0
    
    Select Case IntStyle
        Case 0
            If mintPrintPar = 0 Then
                '提示打印
                If mstrPrintCon <> "" And int发料方式 = 1 Then
                    If bln单据打印 = False Then Exit Sub
                Else
                    If bln已发料清单 = False Then Exit Sub
                End If
                If MsgBox("你需要打印相关单据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
            ElseIf mintPrintPar = 1 Then
                '自动打印
            Else
                Exit Sub
            End If
            Select Case int发料方式
            Case 1  '处方打印
                'mstrPrintCon
                'mstrPrintCon = strNo & "||" & lng单据 & "||" & mlng发料部门id
                If mstrPrintCon <> "" Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723", Me, "库房==" & cboStock.ItemData(cboStock.ListIndex), "NO=" & Split(mstrPrintCon, "||")(0), "单据=" & Split(mstrPrintCon, "||")(1), "审核人=审核人 is not null", 2)
                Else
                    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_1", Me, "库房=" & cboStock.ItemData(cboStock.ListIndex), "发料方式=单据发料|1", "发料时间=" & strDate, 2)
                End If
            Case 2
                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_1", Me, "库房=" & cboStock.ItemData(cboStock.ListIndex), "发料方式=批量发料|2", "发料时间=" & strDate, 2)
            Case 3
                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_1", Me, "库房=" & cboStock.ItemData(cboStock.ListIndex), "发料方式=部门发料|3", "发料时间=" & strDate, 2)
            End Select
       Case 1
            '单据打印
            Dim strNo As String
            Dim int单据 As Integer
            If bln单据打印 = False Then Exit Sub
            
            strNo = Trim(mshHead.TextMatrix(mshHead.Row, mHeadCol.NO))
            int单据 = Val(mshHead.TextMatrix(mshHead.Row, mHeadCol.单据))
            
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723", Me, "库房==" & cboStock.ItemData(cboStock.ListIndex), "NO=" & strNo, "单据=" & int单据, "审核人=" & "审核人 is not null ", 2)
            
       Case 2
            '退料单据
            If bln退料单 = False Then Exit Sub
             Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_2", Me, "退料时间=" & strDate, "单位=" & mintUnit + 1, 2)
    End Select
End Sub
Private Sub 权限控制()
    '
    Dim bln发料 As Boolean
    Dim bln退料 As Boolean
    Dim bln参数 As Boolean
    Dim bln退料单 As Boolean
    Dim bln已发料清单 As Boolean
    Dim bln单据打印 As Boolean
    
    bln发料 = InStr(1, gstrPrivs, "卫生材料发料") <> 0
    bln退料 = InStr(1, gstrPrivs, "卫生材料退料") <> 0
    bln参数 = True ' InStr(1, gstrPrivs, "参数设置") <> 0
    
    bln退料单 = InStr(1, gstrPrivs, "退料通知单") <> 0
    bln已发料清单 = InStr(1, gstrPrivs, "打印已发料清单") <> 0
    bln单据打印 = InStr(1, gstrPrivs, "单据打印") <> 0
    
    mnuFilePara.Visible = bln参数
    mnuFile3.Visible = bln参数
    mnuEditPay.Visible = bln发料
    mnuEditPayCf.Visible = bln发料
    mnuEditFpPay.Visible = bln发料
    
    mnuEditOutPay.Visible = bln退料
    mnuEditSplit0.Visible = bln发料 Or bln退料
    
    mnuFileBillprint.Visible = bln已发料清单 Or bln单据打印
    mnuFileRestore.Visible = bln退料单
    mnuFile2.Visible = bln已发料清单 Or bln单据打印 Or bln退料单
    mnuEditStop.Visible = bln退料 Or bln发料
    mnuEditStopSp.Visible = mnuEditStop.Visible
    mnuEditStrict.Visible = bln退料
    tlbThis.Buttons("发料").Visible = bln发料
    tlbThis.Buttons("退料").Visible = bln退料
    tlbThis.Buttons("EditSp").Visible = bln发料 Or bln退料
    
End Sub
Private Function AnalyseHistorySQL(ByVal strSQL As String, Optional str原串 As String = "", Optional str现串 As String = "") As String
    '产生历史数据的SQL语句
    Dim strTemp As String
    strTemp = Replace(strSQL, "药品收发记录", "H药品收发记录")
    strTemp = Replace(strTemp, "病人费用记录", "H病人费用记录")
    If str原串 <> "" Then
        strTemp = Replace(strTemp, str原串, str现串)
    End If
    strTemp = strSQL & " Union ALL " & strTemp
    AnalyseHistorySQL = strTemp
End Function



Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

