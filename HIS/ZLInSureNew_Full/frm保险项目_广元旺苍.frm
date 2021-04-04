VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frm保险项目_广元旺苍 
   BackColor       =   &H8000000A&
   Caption         =   "医保项目管理"
   ClientHeight    =   6420
   ClientLeft      =   165
   ClientTop       =   3750
   ClientWidth     =   10110
   Icon            =   "frm保险项目_广元旺苍.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin ZL9BillEdit.BillEdit mshSum_S 
      Height          =   2745
      Left            =   3090
      TabIndex        =   4
      Top             =   1020
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   4842
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
   Begin VB.PictureBox picV 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   2580
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3225
      ScaleWidth      =   45
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   900
      Width           =   45
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2670
      Top             =   2280
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
            Picture         =   "frm保险项目_广元旺苍.frx":0E42
            Key             =   "R"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目_广元旺苍.frx":115C
            Key             =   "C"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目_广元旺苍.frx":12B6
            Key             =   "P"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwMain_S 
      Height          =   3525
      Left            =   90
      TabIndex        =   7
      Top             =   960
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   6218
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ilsColor 
      Left            =   3450
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目_广元旺苍.frx":1708
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目_广元旺苍.frx":1924
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目_广元旺苍.frx":1B40
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目_广元旺苍.frx":1D5A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目_广元旺苍.frx":1F76
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMono 
      Left            =   2760
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目_广元旺苍.frx":2192
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目_广元旺苍.frx":23AE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目_广元旺苍.frx":25CA
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目_广元旺苍.frx":27E4
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目_广元旺苍.frx":2A00
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   1270
      BandCount       =   2
      _CBWidth        =   10110
      _CBHeight       =   720
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   660
      Width1          =   5370
      Key1            =   "only"
      NewRow1         =   0   'False
      BandForeColor2  =   8388608
      Caption2        =   "医保中心"
      Child2          =   "cmb险类"
      MinHeight2      =   300
      Width2          =   2325
      UseCoolbarColors2=   0   'False
      NewRow2         =   0   'False
      Begin VB.ComboBox cmb险类 
         Height          =   300
         Left            =   6345
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   210
         Width           =   3675
      End
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   660
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   1164
         ButtonWidth     =   820
         ButtonHeight    =   1164
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMono"
         HotImageList    =   "ilsColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Description     =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查找"
               Key             =   "Find"
               Object.ToolTipText     =   "查找"
               Object.Tag             =   "查找"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   6060
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   635
      SimpleText      =   $"frm保险项目_广元旺苍.frx":2C1C
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm保险项目_广元旺苍.frx":2C63
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12753
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
   Begin VB.CommandButton cmdRestore 
      Caption         =   "还原(&R)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6750
      TabIndex        =   6
      Top             =   4080
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5340
      TabIndex        =   5
      Top             =   4080
      Width           =   1100
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
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
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuCX 
         Caption         =   "查询所有项目"
      End
      Begin VB.Menu mnuSB 
         Caption         =   "申报所有项目"
      End
      Begin VB.Menu mnuFileLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditGet 
         Caption         =   "重新提取项目审核信息(&G)"
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
         Begin VB.Menu mnuViewToolSplit 
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
      Begin VB.Menu mnuViewSplit0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewItem 
         Caption         =   "不编辑项目编码(&I)"
      End
      Begin VB.Menu mnuViewClass 
         Caption         =   "不编辑医保大类(&C)"
      End
      Begin VB.Menu mnuViewSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "查找(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R) "
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web上的中联"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu mnuHelpWebL 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)…"
      End
   End
End
Attribute VB_Name = "frm保险项目_广元旺苍"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private int审核标志 As Integer
Private classInsure As New clsInsure

Private Enum ColumnEnum
    cOL编码 = 0
    cOL名称 = 1
    col产地 = 2
    COL规格 = 3
    COL剂型 = 4
    COL单位 = 5
    col价格 = 6
    col改变方式 = 7
    col大类ID = 8
    COL医保编码 = 9
    col医保名称 = 10
    col医保剂型 = 11
    col医保附注 = 12
    col原编码 = 13
    col大类名称 = 14
    col非医保 = 15
    'Modified By 朱玉宝 地区：长沙 原因：没法，只有加了
    col匹配序列号 = 16
    col审核标志 = 17
End Enum
Private Const mlng编码长度 As Long = 20

Dim mlngListIndex As Long   '保存上次下拉框的选择索引
Dim mblnLoad As Boolean
Dim msngStartX As Single    '移动前鼠标的位置
Dim mstr权限 As String

Dim mstrKey As String       '前一个树节点的关键值
Dim mint中心 As Integer     '当前显示的险类
Dim mint适用地区 As Integer '沈阳专用；0表示其他地区，1表示长春（允许删除已审核的项目）

Dim mlngCol As Long, mblnDesc As Boolean
Private mcnYB As New ADODB.Connection   '医保前置服务器连接
Private mint险类 As Integer         '当前显示的险类


Private Sub cbrThis_HeightChanged(ByVal NewHeight As Single)
    Call ResizeForm(NewHeight)
End Sub

Private Sub cmdRestore_Click()
    'Modified By 朱玉宝 地区：长沙
    If MsgBox("你确认要放弃修改吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    Call FillSum(True)
    mshSum_S.SetFocus
End Sub

Private Sub cmdSave_Click()
    Dim lngRow As Long
    
    If mint险类 = TYPE_成都德阳 Then
        gcnOracle_成都德阳.BeginTrans
    ElseIf mint险类 = TYPE_南充阆中 Then
        gcnOracle_南充阆中.BeginTrans
    Else
        gcnOracle_广元旺苍.BeginTrans
    End If
    
    On Error GoTo errHandle
    
    With mshSum_S
        '处理数据
        For lngRow = 1 To .Rows - 1
            If mint险类 = TYPE_广元旺苍 And InitInfor_广元旺苍.适用地区 = 0 Then Exit Sub
            Select Case .TextMatrix(lngRow, col改变方式)
                Case "新增", "修改"
                    '将新增与修改放在一个过程中处理
                    '过程参数:
                    '    收费细目ID_IN IN 医保支付项目.收费细目ID%TYPE,
                    '    险类_IN       IN 医保支付项目.险类%TYPE,
                    '    中心_IN       IN 医保支付项目.中心%TYPE,
                    '    大类ID_IN     IN 医保支付项目.大类ID%TYPE,
                    '    项目编码_IN   IN 医保支付项目.项目编码%TYPE,
                    '    项目名称_IN   IN 医保支付项目.项目名称%TYPE,
                    '    附注_IN       IN 医保支付项目.附注%TYPE,
                    '    是否医保_IN   IN 医保支付项目.是否医保%TYPE
    
                    gstrSQL = "ZL_医保支付项目_Modify(" & .RowData(lngRow) & "," & mint险类 & "," & mint中心 & "," & _
                               IIf(Val(.TextMatrix(lngRow, col大类ID)) = 0, "null", .TextMatrix(lngRow, col大类ID)) & ",'" & _
                               .TextMatrix(lngRow, COL医保编码) & "','" & .TextMatrix(lngRow, col医保名称) & "','" & .TextMatrix(lngRow, col医保附注) & _
                               IIf(mint中心 = TYPE_沈阳市, "^^" & .TextMatrix(lngRow, col匹配序列号) & "||" & _
                               IIf(Trim(.TextMatrix(lngRow, col审核标志)) = "√", 1, IIf(Trim(.TextMatrix(lngRow, col审核标志)) = "×", 2, 0)), "") & _
                               "'," & IIf(Trim(.TextMatrix(lngRow, col非医保)) = "√", 0, 1) & ")"
                    If mint险类 = TYPE_成都德阳 Then
                        ExecuteProcedure_成都德阳 Me.Caption
                    ElseIf mint险类 = TYPE_南充阆中 Then
                        ExecuteProcedure_南充阆中 Me.Caption
                    Else
                        ExecuteProcedure_广元旺苍 Me.Caption
                    End If
                    .TextMatrix(lngRow, col原编码) = .TextMatrix(lngRow, COL医保编码)
                Case "删除"
                    '过程参数:
                    '    收费细目ID_IN IN 医保支付项目.收费细目ID%TYPE,
                    '    险类_IN       IN 医保支付项目.险类%TYPE,
                    '    中心_IN       IN 医保支付项目.中心%TYPE
    
                    gstrSQL = "ZL_医保支付项目_Delete(" & .RowData(lngRow) & "," & mint险类 & "," & mint中心 & ")"
                    If mint险类 = TYPE_成都德阳 Then
                        ExecuteProcedure_成都德阳 Me.Caption
                    ElseIf mint险类 = TYPE_南充阆中 Then
                        ExecuteProcedure_南充阆中 Me.Caption
                    Else
                        ExecuteProcedure_广元旺苍 Me.Caption
                    End If
                    .TextMatrix(lngRow, col原编码) = .TextMatrix(lngRow, COL医保编码)
            End Select
        Next
        
        '待数据处理完成无误后，再设置数据状态
        For lngRow = 1 To .Rows - 1
            .TextMatrix(lngRow, col改变方式) = ""
        Next
    End With
    cmdRestore.Enabled = False
    cmdSave.Enabled = False
    If mint险类 = TYPE_成都德阳 Then
        gcnOracle_成都德阳.CommitTrans
    ElseIf mint险类 = TYPE_南充阆中 Then
        gcnOracle_南充阆中.CommitTrans
    Else
        gcnOracle_广元旺苍.CommitTrans
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    If mint险类 = TYPE_成都德阳 Then
        gcnOracle_成都德阳.RollbackTrans
    ElseIf mint险类 = TYPE_南充阆中 Then
        gcnOracle_南充阆中.RollbackTrans
    Else
        gcnOracle_广元旺苍.RollbackTrans
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If mblnLoad = True Then
        Call FillTree
    End If
    
    Call mshSum_S_EnterCell(1, cOL编码)
    mblnLoad = False
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    mstrKey = ""
    mlngCol = 0
    mblnDesc = False
    mblnLoad = True
    
    
    If mint险类 = TYPE_广元旺苍 Then Call 医保初始化_广元旺苍
    If mint险类 = TYPE_成都德阳 Then Call 医保初始化_成都德阳
    If mint险类 = TYPE_南充阆中 Then Call 医保初始化_南充阆中
    
    gstrSQL = "select 序号,编码,名称 from 保险中心目录 where 序号<>0 and 险类=[1] order by 序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint险类)
    
    With cmb险类
        .Clear
        Do Until rsTemp.EOF
            .AddItem Nvl(rsTemp!编码) & "--" & rsTemp("名称")
            .ItemData(.NewIndex) = rsTemp("序号")
            If rsTemp("序号") = mint险类 Then
                '当前医保。
                '使用API，可以不激活Click事件
                zlControl.CboSetIndex .hwnd, .NewIndex
                Call Fill大类
            End If
            rsTemp.MoveNext
        Loop
        If .ListIndex < 0 Then
            '使用API，可以不激活Click事件
            zlControl.CboSetIndex .hwnd, 0
            Call Fill大类
        End If
    End With
    mint中心 = cmb险类.ItemData(cmb险类.ListIndex)
    
    Call InitSum
    RestoreWinState Me, App.ProductName
    
    mnuViewItem.Checked = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewItem", "False") <> "False"
    If mnuViewItem.Checked = False Then
        '不用判断大类了
        mnuViewClass.Checked = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewClass", "False") <> "False"
    End If
    Call SetSkip
    zlControl.CboSetHeight cmb险类, 3600
    

    If Nvl(InitInfor_广元旺苍.适用地区, "0") = 0 And mint险类 = TYPE_广元旺苍 Then
       mnuEdit.Visible = True
       mnuCX.Visible = True
       mnuSB.Visible = True
       mnuEditGet.Visible = False
    Else
       mnuEdit.Visible = False
    End If
End Sub

Private Sub InitSum()
    '初始化汇总表的样式
    Dim lngCol As Long
    
    With mshSum_S
        ClearGrid mshSum_S
        
        'Modified By 朱玉宝 地区：长沙 原因：增加列——匹配序列号
        .Cols = 18
        
        .TextMatrix(0, cOL编码) = "编码"
        .TextMatrix(0, cOL名称) = "收费细目"
        .TextMatrix(0, COL规格) = "规格"
        .TextMatrix(0, col产地) = "产地"
        .TextMatrix(0, COL单位) = "单位"
        .TextMatrix(0, col价格) = "价格"
        .TextMatrix(0, col改变方式) = "是否修改"
        .TextMatrix(0, COL医保编码) = "医保项目编码"
        .TextMatrix(0, col医保名称) = "医保项目名称"
        .TextMatrix(0, COL剂型) = "剂型"
        .TextMatrix(0, col医保剂型) = "剂型"
        .TextMatrix(0, col审核标志) = "审核"
        .TextMatrix(0, col医保附注) = "医保项目附注"
        .TextMatrix(0, col原编码) = "原医保项目编码"
        .TextMatrix(0, col大类ID) = "大类ID"
        .TextMatrix(0, col大类名称) = "医保大类名称"
        If Nvl(InitInfor_广元旺苍.适用地区, "0") = 0 And mint险类 = TYPE_广元旺苍 Then
           .TextMatrix(0, col非医保) = "未启用"
        Else
           .TextMatrix(0, col非医保) = "是否医保"
        End If
        .TextMatrix(0, col匹配序列号) = "匹配序列号"
        
        .ColWidth(cOL编码) = 1000
        .ColWidth(cOL名称) = 2000
        .ColWidth(COL规格) = 1000
        .ColWidth(col产地) = 600
        .ColWidth(COL单位) = 600
        .ColWidth(col价格) = 800
        .ColWidth(col改变方式) = 0
        .ColWidth(COL医保编码) = 1200
        .ColWidth(col医保名称) = 1200
        .ColWidth(col医保附注) = 0
        .ColWidth(col原编码) = 0
        .ColWidth(col大类ID) = 0
        .ColWidth(col大类名称) = 1200
        .ColWidth(col非医保) = 800
        .ColWidth(col匹配序列号) = 0
        
        .ColWidth(COL剂型) = 0
        .ColWidth(col医保剂型) = 0
        .ColWidth(col审核标志) = 0
        
        
        For lngCol = 0 To .Cols - 1
            .ColAlignment(lngCol) = 1
        Next
        .ColAlignment(col价格) = 7
        .ColAlignment(col非医保) = 4
        
        '设置各列的编辑特性
        .ColData(COL剂型) = 5
        .ColData(col医保剂型) = 5
        .ColData(col审核标志) = 5
        .ColData(cOL编码) = 5 '不能选择
        .ColData(cOL名称) = 5
        .ColData(COL规格) = 5
        .ColData(col产地) = 5
        .ColData(COL单位) = 5
        .ColData(col价格) = 5
        .ColData(col改变方式) = 5
        .ColData(COL医保编码) = 1
        .ColData(col医保名称) = 5
        
        .ColData(col医保附注) = 5
        .ColData(col原编码) = 5
        .ColData(col大类ID) = 5
        .ColData(col大类名称) = 3 '选择器
        .ColData(col非医保) = -1 '选择器
        .ColData(col匹配序列号) = 5
        
        .PrimaryCol = cOL编码
        Call SetSkip
        .AllowAddRow = False
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdSave.Enabled = True Then
        MsgBox "医保项目列表正处于编辑状态，不能退出程序。", vbInformation, gstrSysName
        Cancel = 1
        Exit Sub
    End If
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewItem", mnuViewItem.Checked
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewClass", mnuViewClass.Checked
    SaveWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    Call ResizeForm(cbrThis.Height)
End Sub

Private Sub ResizeForm(ByVal cbrHeight As Single)
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    sngTop = IIf(cbrThis.Visible, cbrThis.Top + cbrHeight, 0)
    sngBottom = ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    '右边
    'tvwMain_S的位置
    tvwMain_S.Top = sngTop
    tvwMain_S.Height = IIf(sngBottom - tvwMain_S.Top > 0, sngBottom - tvwMain_S.Top, 0)
    tvwMain_S.Left = 0
    'picV的位置
    picV.Top = sngTop
    picV.Height = tvwMain_S.Height
    picV.Left = tvwMain_S.Left + tvwMain_S.Width
    
    cmdRestore.Top = sngBottom - cmdRestore.Height - 100
    cmdRestore.Left = ScaleWidth - cmdRestore.Width - 300
    cmdSave.Top = cmdRestore.Top
    cmdSave.Left = cmdRestore.Left - cmdSave.Width - 300
    
    If InStr(mstr权限, "增删改") > 0 Then
        '可以编辑
        sngBottom = cmdRestore.Top - 100
    End If
    
    mshSum_S.Left = picV.Left + picV.Width
    If ScaleWidth - mshSum_S.Left > 0 Then mshSum_S.Width = ScaleWidth - mshSum_S.Left
    mshSum_S.Top = sngTop
    mshSum_S.Height = IIf(sngBottom - mshSum_S.Top > 0, sngBottom - mshSum_S.Top, 0)
    
    Refresh
End Sub

Private Sub mnuCX_Click()
    Dim i As Integer
    Dim strSQL As String, intID As Integer
    Dim StrInput As String, strOutput As String
    Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
    Dim int大类id As Integer
    Dim strTmpArr As Variant, strArr As Variant

    '提取基础资料
    StrInput = vbTab & g病人身份_广元旺苍.机构编码
    StrInput = StrInput & vbTab & "0"
    Me.Caption = "医保项目管理      正在从中心提取基础项目资料....."
    If 业务请求_广元旺苍(提取基础资料_旺苍, StrInput, strOutput) = False Then Exit Sub
    
    gstrSQL = "select A.ID,E.项目编码 as 编码,F.名称 as 类别," & _
              "E.项目名称 As 中文名称,'' as 英文名称, " & _
              "zlspellcode(A.名称) as 简码,Substr(A.名称,1,20) as 别名,A.计算单位, " & _
              "substr(A.规格,1,instr(A.规格,'┆')-1) as 规格, " & _
              "A.费用类型 as 费用类别,nvl(E.是否医保,0) as 启用标志 " & _
              "from 收费细目 A,医保支付项目 E,保险支付大类 F " & _
              "where " & _
              "nvl(A.撤档时间,to_date('3000-01-01','YYYY-MM-DD'))=to_date('3000-01-01','YYYY-MM-DD') and " & _
              "A.ID=E.收费细目ID And E.险类=[1] And E.中心=[2]" & _
              " And E.大类ID=F.ID And E.险类=F.险类 and nvl(E.是否医保,0)=0 "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "保险项目选择", mint险类, mint中心)
    
    i = 0
    intID = 0
    
    Do While Not rsTmp.EOF
       If rsTmp!启用标志 = 0 And intID <> rsTmp!ID Then
            i = i + 1
            
            StrInput = vbTab & g病人身份_广元旺苍.机构编码
            StrInput = StrInput & vbTab & rsTmp!编码
            
            If 业务请求_广元旺苍(提取项目_资阳, StrInput, strOutput) = False Then Exit Sub
            
            strArr = Split(strOutput, "@$")
            strTmpArr = Split(strArr(0), "||")
        
            If rsTmp!费用类别 <> strTmpArr(4) Or rsTmp!类别 <> strTmpArr(2) Then
                   
                   '更新费用类别
                  '$IF HIS9.19
                  #If gverControl = 0 Then
                        gstrSQL = "ZL_收费细目_UPDATE_资阳(" & rsTmp!ID & ",'" & rsTmp!费用类别 & "')"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新费用类别")
                  #Else
                  '$ELSE  HIS+
                        gstrSQL = "ZL_收费项目目录_UPDATE_资阳(" & rsTmp!ID & ",'" & rsTmp!费用类别 & "')"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新费用类别")
                  #End If
            End If
                      
            gstrSQL = "select nvl(ID,0) as ID from 保险支付大类 where 险类=[1] And 名称=[2]"
            Set rsTmp1 = zlDatabase.OpenSQLRecord(gstrSQL, "查询支付大类", mint险类, CStr(strTmpArr(2)))
            int大类id = rsTmp1!ID
        
            gstrSQL = "ZL_医保支付项目_Modify(" & rsTmp!ID & "," & mint险类 & "," & mint中心 & "," & _
                      int大类id & ",'" & rsTmp!编码 & "','" & rsTmp!中文名称 & "','" & Format(zlDatabase.Currentdate, "YYYY-MM-DD") & "'," & IIf(strTmpArr(1) = "启用", 1, 0) & ")"
            ExecuteProcedure_广元旺苍 "保存医保支付项目"
            
            Me.Caption = "医保项目管理      正在查询第" & i & "条未启用项目，请稍候....."
            If strTmpArr(1) <> "启用" Then
               MsgBox "项目【" & rsTmp!编码 & "】" & rsTmp!中文名称 & "在中心尚未启用。"
               Exit Sub
            End If
       End If
       intID = rsTmp!ID
       rsTmp.MoveNext
    Loop
    
    MsgBox "所有未申报项目的查询已经全部完成。"
End Sub

Private Sub mnuSB_Click()
    Dim i As Integer
    Dim strSQL As String, intID As Integer
    Dim StrInput As String, strOutput As String
    Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
    Dim int大类id As Integer
    
    gstrSQL = "select A.ID,A.编码,decode(A.类别,'J','服务','1','服务','5','药品','6','药品','7','药品','诊疗') as 类别," & _
              "A.名称 As 中文名称,'' as 英文名称, " & _
              "zlspellcode(A.名称) as 简码,Substrb(A.名称,1,40) as 别名,substrb(A.计算单位,1,20) as 计算单位, " & _
              "B.现价,substrb(substr(A.规格,1,instr(A.规格,'┆')-1),1,20) as 规格, " & _
              "D.名称 as 费用项目,A.费用类型 as 费用类别,E.项目编码 " & _
              "from 收费细目 A,收费价目 B,收入项目 D,医保支付项目 E " & _
              "where A.ID=B.收费细目ID and B.收入项目ID=D.ID And " & _
              "nvl(B.终止日期,to_date('3000-01-01','YYYY-MM-DD'))=to_date('3000-01-01','YYYY-MM-DD') and " & _
              "A.ID=E.收费细目ID(+) And E.险类(+)=[1]And E.中心(+)=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "保险项目选择", mint险类, mint中心)
    
    i = 0
    intID = 0
    
    Do While Not rsTmp.EOF
       If IsNull(rsTmp!项目编码) And intID <> rsTmp!ID Then
            i = i + 1
                        
            StrInput = vbTab & g病人身份_广元旺苍.机构编码
            StrInput = StrInput & vbTab & rsTmp!编码 & "||"
            StrInput = StrInput & rsTmp!类别 & "||"
            StrInput = StrInput & rsTmp!中文名称 & "||"
            StrInput = StrInput & rsTmp!英文名称 & "||"
            StrInput = StrInput & rsTmp!简码 & "||"
            StrInput = StrInput & rsTmp!别名 & "||"
            StrInput = StrInput & rsTmp!计算单位 & "||"
            StrInput = StrInput & rsTmp!现价 & "||"
            StrInput = StrInput & rsTmp!规格 & "||"
            StrInput = StrInput & rsTmp!费用项目 & "||"
            StrInput = StrInput & rsTmp!费用类别
            
            StrInput = StrInput & vbTab & gstrUserName
            StrInput = StrInput & vbTab & Format(zlDatabase.Currentdate, "YYYY-M-DD")
            
            If 业务请求_广元旺苍(申报项目_资阳, StrInput, strOutput) = False Then Exit Sub
            
            gstrSQL = "select nvl(ID,0) as ID from 保险支付大类 where 险类=[1] And 名称=[2]"
            Set rsTmp1 = zlDatabase.OpenSQLRecord(gstrSQL, "查询支付大类", mint险类, CStr(rsTmp!类别))
            int大类id = rsTmp1!ID
            
            gstrSQL = "ZL_医保支付项目_Modify(" & rsTmp!ID & "," & mint险类 & "," & mint中心 & "," & _
                       int大类id & ",'" & rsTmp!编码 & "','" & rsTmp!别名 & "','" & Format(zlDatabase.Currentdate, "YYYY-MM-DD") & "',0)"
            ExecuteProcedure_广元旺苍 "保存医保支付项目"
            Me.Caption = "医保项目管理      正在上传第" & i & "条申报项目，请稍候....."
       End If
       intID = rsTmp!ID
       rsTmp.MoveNext
    Loop
    
    MsgBox "所有未申报项目已经全部上传完成。"
    
End Sub

Private Sub mnuViewFind_Click()
    If cmdSave.Enabled = True Then
        MsgBox "医保项目列表正处于编辑状态，不能使用查找功能。", vbInformation, gstrSysName
        Exit Sub
    End If
    frm保险项目查找广元旺苍.Show vbModal, Me
End Sub

Private Sub cmb险类_Click()
    Call Fill大类
    Call FillSum(False)
End Sub

Private Sub mnuViewClass_Click()
    mnuViewItem.Checked = False
    mnuViewClass.Checked = Not mnuViewClass.Checked
    Call SetSkip
End Sub

Private Sub mnuViewItem_Click()
    mnuViewClass.Checked = False
    mnuViewItem.Checked = Not mnuViewItem.Checked
    Call SetSkip
End Sub

Private Sub SetSkip()
'设置表格的跳跃属性
    With mshSum_S
        If mnuViewItem.Checked = False Then
        
            .ColData(COL医保编码) = 1
            .LocateCol = COL医保编码
            .ColData(col大类名称) = IIf(mnuViewClass.Checked = True, 5, 3)
        Else
            .ColData(col大类名称) = 3 '选择器
            .LocateCol = col大类名称
            .ColData(COL医保编码) = 5
        End If
        If .ColData(.COL) = 5 Then
            '当前列已经不能选择，需重新定位
            .COL = .LocateCol
        End If
    End With
End Sub

Private Sub mnuViewRefresh_Click()
    '只刷新列表内容
    Call FillSum
End Sub

Private Sub mshSum_S_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    '始终是不允许删除的
    Cancel = True
    
    With mshSum_S
        If .TextMatrix(Row, col改变方式) = "新增" Then
            .TextMatrix(Row, col改变方式) = "" '相当于什么都没有做
        Else
            .TextMatrix(Row, col改变方式) = "删除" '标记
        End If
        
        .TextMatrix(Row, COL医保编码) = ""
        .TextMatrix(Row, col医保名称) = ""
        .TextMatrix(Row, col医保剂型) = ""
        .TextMatrix(Row, col医保附注) = ""
        .TextMatrix(Row, col大类ID) = ""
        .TextMatrix(Row, col大类名称) = ""
        .TextMatrix(Row, col非医保) = ""
        .TextMatrix(Row, col审核标志) = ""
    End With
    cmdSave.Enabled = True
    cmdRestore.Enabled = True
End Sub

Private Sub mshSum_S_cboClick(ListIndex As Long)
    With mshSum_S
        If .COL <> col大类名称 Then Exit Sub
        
        If .TextMatrix(.Row, col大类名称) <> .CboText Then
            '禁止修改保险大类,只允许通过选择明细以确定大类
            If mint险类 = TYPE_泸州市 Then
                .ListIndex = mlngListIndex
                Exit Sub
            End If
            mlngListIndex = ListIndex
            .TextMatrix(.Row, col大类名称) = .CboText
            Call 标记改变
        Else
            mlngListIndex = ListIndex
        End If
        
        If .CboText = "" Then
            '保存为空
            .TextMatrix(.Row, col大类ID) = ""
            .TextMatrix(.Row, col大类名称) = ""
        Else
            .TextMatrix(.Row, col大类ID) = .ItemData(.ListIndex)
            .TextMatrix(.Row, col大类名称) = .CboText
        End If
        
    End With
End Sub

Private Sub mshSum_S_cboKeyDown(KeyCode As Integer, Shift As Integer)
    With mshSum_S
        If KeyCode = vbKeyReturn Then
            If .TextMatrix(.Row, col大类名称) <> .CboText Then
                .TextMatrix(.Row, col大类名称) = .CboText
                Call 标记改变
            End If
            
            If .CboText = "" Then
                '保存为空
                .TextMatrix(.Row, col大类ID) = ""
                .TextMatrix(.Row, col大类名称) = ""
                .COL = col非医保
            Else
                .TextMatrix(.Row, col大类ID) = .ItemData(.ListIndex)
                .TextMatrix(.Row, col大类名称) = .CboText
            End If
        End If
    End With
End Sub

Private Sub mshSum_S_CommandClick()
'功能：提取医保项目供选择
'参数：无
'返回：医保项目编码
    Dim strCode As String
    Dim strSelected As String
    Dim STRNAME As String
    Dim strlastCode As String
    Dim strMemo As String
    
    With mshSum_S
        strCode = .TextMatrix(.Row, COL医保编码)
        If InitInfor_广元旺苍.适用地区 = 0 And mint险类 = TYPE_广元旺苍 Then
            strCode = .TextMatrix(.Row, cOL编码)
            If Frm医保对码_资阳.GetCode(strCode, mint中心, mint险类) = True Then
               strSelected = strCode
            End If
        Else
            If frm保险项目选择广元旺苍.GetCode(strCode, mint中心, mint险类) = True Then
               strSelected = strCode
            End If
        End If
        
        If strSelected <> "" Then
            .TextMatrix(.Row, COL医保编码) = strSelected
            If STRNAME = "" Then
                Call Get保险名称
            Else
                '已经传入名称，就不用再调用
                .TextMatrix(.Row, col医保名称) = STRNAME
                .TextMatrix(.Row, col医保附注) = ""
                .TextMatrix(.Row, col非医保) = ""
            End If
            Call 标记改变
        End If
    End With
End Sub

Private Sub mshSum_S_DblClick(Cancel As Boolean)
    With mshSum_S
        If .Active = False Then Exit Sub
        Call 标记改变
    End With
End Sub

Private Sub mshSum_S_EnterCell(Row As Long, COL As Long)
    Static lngRow As Long
    Dim rsTemp As New ADODB.Recordset
    
    If COL = col大类名称 And Trim(mshSum_S.TextMatrix(Row, COL)) = "" Then
        mshSum_S.ListIndex = -1
    End If
End Sub

Private Sub mshSum_S_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    '保险项目编码
    Dim str前 As String, strText As String, str类型 As String
    Dim rsTemp As New ADODB.Recordset, blnReturn As Boolean
    Dim strLeft As String
    Dim strTemp As String

    str前 = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", "0") = "0", "%", "") '双向匹配
    
    On Error GoTo errHandle
    
    With mshSum_S
        If .COL <> COL医保编码 Then Exit Sub
        If KeyCode = vbKeyReturn Then
            If mint险类 = TYPE_广元旺苍 And InitInfor_广元旺苍.适用地区 = 0 Then
                If strText = "" Then strText = .TextMatrix(.Row, cOL编码)
                If Frm医保对码_资阳.GetCode(strText, mint中心, mint险类) = True Then blnReturn = strText
            Else
                If .TxtVisible = True Then
                    strText = Replace(Trim(.Text), "`", "")
                    .Text = strText
                    If zlCommFun.StrIsValid(strText, mlng编码长度) = False Then
                        Cancel = True
                        Exit Sub
                    End If
                    If Trim(strText) = "" Then
                        '不需要再去检查是否有匹配的编码，相当于删除该编码
                        .TextMatrix(.Row, COL医保编码) = Trim(strText)
                    Else
                        '产生SQL语句
                        If mint险类 = TYPE_成都德阳 Then
                            gstrSQL = "Select 编码  医保编码,名称,简码,附注 " & _
                                      " FROM 医保收费项目_德阳 WHERE " & _
                                      " 险类=[1] and 中心=[2] and (编码 like [3] || '%' or Upper(名称) like [3] || '%' Or Upper(简码) like [3] || '%')"
                        Else
                            gstrSQL = "Select 编码  医保编码,名称,简码,附注 " & _
                                         "   FROM 医保收费项目 WHERE " & _
                                      " 险类=[1] and 中心=[2] and (编码 like [3] || '%' or Upper(名称) like [3] || '%' Or Upper(简码) like [3] || '%')"
                        End If
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint险类, mint中心, strText)
    
                        If rsTemp.RecordCount > 0 Then
                            '出现选择器
                            If rsTemp.RecordCount >= 1 Or rsTemp.Fields.Count > 3 Then
                                '对于字段大于3的，即使只有一条记录把该对话框显示出来，以便让用户得到更多的信息
                                blnReturn = frmListSel.ShowSelect(mint险类, rsTemp, "医保编码", "医保项目选择", "请选择对应的医保项目：")
                            End If
                        End If
                        
                        If blnReturn = False Then
                            '记录集中没有可选择的数据
                            If rsTemp.RecordCount > 0 Then
                                '记录集有数据，但取消了选择
                                Cancel = True
                                .TxtVisible = True
                                .TxtSetFocus
                                Exit Sub
                            Else
                                .Text = strText
                                .TextMatrix(.Row, COL医保编码) = strText
                            End If
                        Else
                            .Text = rsTemp("医保编码")
                            .TextMatrix(.Row, COL医保编码) = rsTemp("医保编码")
                        End If
                    End If
                    Call Get保险名称
                    Call 标记改变
                End If
            End If
        Else
            If .TextMatrix(.Row, COL医保编码) = "" Then
                .TextMatrix(.Row, COL医保编码) = " "
            End If
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Cancel = True
End Sub

Private Sub 标记改变()
    '当前输入已经有效，继续看能否得到其它内容
    If mint险类 = TYPE_广元旺苍 And InitInfor_广元旺苍.适用地区 = 0 Then
        cmdRestore.Enabled = False
        cmdSave.Enabled = False
    Else
        cmdRestore.Enabled = True
        cmdSave.Enabled = True
    End If
    
    With mshSum_S
        If Trim(.TextMatrix(.Row, COL医保编码)) = "" And Trim(.TextMatrix(.Row, col大类名称)) = "" Then
            .TextMatrix(.Row, col改变方式) = "删除"
        Else
            If Trim(.TextMatrix(.Row, col改变方式)) <> "修改" Then
                '为空，或已经是“新增”
                .TextMatrix(.Row, col改变方式) = "新增"
            End If
        End If
    End With
End Sub

Private Sub Get保险名称()
'功能：根据当前行的保险项目编码，得到其它信息
    Dim rsTemp As New ADODB.Recordset
    Dim lngIndex As Long, lngPos As Long
    Dim str大类编码 As String, strTemp As String, varPart As Variant
    
    On Error GoTo errHandle
    With mshSum_S
        If mint险类 = TYPE_成都德阳 Then
            gstrSQL = "select 名称,大类编码,附注 from 医保收费项目_德阳 where 编码=[1] and 险类=[2] and 中心=[3]"
        Else
            gstrSQL = "select 名称,大类编码,附注 from 医保收费项目 where 编码=[1] and 险类=[2] and 中心=[3]"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(.TextMatrix(.Row, COL医保编码)), mint险类, CLng(cmb险类.ItemData(cmb险类.ListIndex)))
        
        If rsTemp.RecordCount = 0 Then
            '没有对应的保险项目，只有利用该编码
            .TextMatrix(.Row, col医保名称) = ""
            .TextMatrix(.Row, col医保附注) = ""
            .TextMatrix(.Row, col非医保) = ""
        Else
            .TextMatrix(.Row, col医保名称) = IIf(IsNull(rsTemp("名称")), "", rsTemp("名称"))
            .TextMatrix(.Row, col医保附注) = IIf(IsNull(rsTemp("附注")), "", rsTemp("附注"))
            str大类编码 = IIf(IsNull(rsTemp("大类编码")), "", rsTemp("大类编码"))
        End If
        For lngIndex = 0 To .ListCount - 1
            lngPos = InStr(.List(lngIndex), ".")
            If lngPos = 0 Then
                strTemp = .List(lngIndex)
            Else
                strTemp = Mid(.List(lngIndex), 1, lngPos - 1)
            End If
            If strTemp = str大类编码 Then
                '找到相匹配的大类编码
                .TextMatrix(.Row, col大类ID) = .ItemData(lngIndex)
                .TextMatrix(.Row, col大类名称) = .List(lngIndex)
                Exit For
            End If
        Next
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub mshSum_S_KeyPress(KeyAscii As Integer)
    With mshSum_S
        If Not .Active Then Exit Sub
        If .ColData(.COL) = -1 Then Call 标记改变
    End With
End Sub

Private Sub mshSum_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    mshSum_S.ToolTipText = mshSum_S.TextMatrix(mshSum_S.MouseRow, mshSum_S.MouseCol)
End Sub

Private Sub mshSum_S_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim rsTemp As New ADODB.Recordset, lngID As Long
    Dim lngRow As Long, lngPos As Long, blnActive As Boolean
    Dim blnEnable As Boolean
    
    If mshSum_S.Active = False Then Exit Sub
    If mshSum_S.MouseRow = 0 Then
        If mlngCol = mshSum_S.MouseCol Then
            mblnDesc = Not mblnDesc
        Else
            mlngCol = mshSum_S.MouseCol
            mblnDesc = False
        End If
        
        blnEnable = cmdRestore.Enabled
        blnActive = mshSum_S.Active
        mshSum_S.Active = False
        mshSum_S.msfObj.MousePointer = vbHourglass
        
        '构成记录集，然后刷新表格
        rsTemp.CursorLocation = adUseClient
        rsTemp.CursorType = adOpenDynamic
        rsTemp.LockType = adLockOptimistic
        With rsTemp.Fields
            .Append "ID", adDouble, adFldIsNullable
            .Append "编码", adVarChar, 20, adFldIsNullable
            .Append "名称", adVarChar, 50, adFldIsNullable
            .Append "规格", adVarChar, 80, adFldIsNullable
            .Append "剂型", adVarChar, 30, adFldIsNullable
            .Append "产地", adVarChar, 100, adFldIsNullable
            .Append "单位", adVarChar, 20, adFldIsNullable
            .Append "是否变价", adInteger, adFldIsNullable
            .Append "价格", adVarNumeric, 20, adFldIsNullable
            .Append "改变方式", adVarChar, 4, adFldIsNullable
            'Modified By 朱玉宝 2003-12-09 地区：乐山
            .Append "项目编码", adVarChar, 50, adFldIsNullable
            .Append "项目名称", adVarChar, 50, adFldIsNullable
            .Append "附注", adVarChar, 50, adFldIsNullable
            .Append "原编码", adVarChar, 20, adFldIsNullable
            .Append "是否医保", adInteger
            .Append "大类ID", adDouble
            .Append "大类编码", adVarChar, 10, adFldIsNullable
            .Append "大类名称", adVarChar, 50, adFldIsNullable
        End With
        
        rsTemp.Open
        With mshSum_S
            For lngRow = 1 To .Rows - 1
                rsTemp.AddNew
                
                rsTemp("ID") = .RowData(lngRow)
                rsTemp("编码") = .TextMatrix(lngRow, cOL编码)
                rsTemp("名称") = .TextMatrix(lngRow, cOL名称)
                rsTemp("规格") = .TextMatrix(lngRow, COL规格)
                rsTemp("剂型") = .TextMatrix(lngRow, COL剂型)
                rsTemp("产地") = Substr(.TextMatrix(lngRow, col产地), 1, 100)
                rsTemp("单位") = .TextMatrix(lngRow, COL单位)
                If .TextMatrix(lngRow, col价格) = "" Then
                    rsTemp("是否变价") = 1
                    rsTemp("价格") = 0
                Else
                    rsTemp("是否变价") = 0
                    rsTemp("价格") = Val(.TextMatrix(lngRow, col价格))
                End If
                rsTemp("改变方式") = .TextMatrix(lngRow, col改变方式)
                rsTemp("项目编码") = .TextMatrix(lngRow, COL医保编码)
                rsTemp("项目名称") = .TextMatrix(lngRow, col医保名称)
                rsTemp("附注") = .TextMatrix(lngRow, col医保附注)
                rsTemp("原编码") = .TextMatrix(lngRow, col原编码)
                rsTemp("大类ID") = Val(.TextMatrix(lngRow, col大类ID))
                rsTemp("是否医保") = IIf(.TextMatrix(lngRow, col非医保) = "√", 0, 1)
                
                
                lngPos = InStr(.TextMatrix(lngRow, col大类名称), ".")
                If lngPos = 0 Then
                    rsTemp("大类编码") = Null
                    rsTemp("大类名称") = Null
                Else
                    rsTemp("大类编码") = Mid(.TextMatrix(lngRow, col大类名称), 1, lngPos - 1)
                    rsTemp("大类名称") = Mid(.TextMatrix(lngRow, col大类名称), lngPos + 1)
                End If
                
                rsTemp.Update
            Next
            lngID = .RowData(.Row)
        End With
        Call FillGrid(rsTemp, lngID)
    
        mshSum_S.Active = blnActive '恢复
        mshSum_S.msfObj.MousePointer = vbDefault
        MousePointer = vbDefault
        cmdRestore.Enabled = blnEnable
        cmdSave.Enabled = blnEnable
    End If
End Sub

Public Sub tvwMain_S_NodeClick(ByVal Node As MSComctlLib.Node)
    '只刷新列表内容
    FillSum
End Sub

Private Sub mshSum_S_GotFocus()
    Call MenuSet
End Sub

Private Sub mshSum_S_LostFocus()
    mshSum_S.CmdVisible = False
    mshSum_S.CboVisible = False
    
    Call MenuSet
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool, 2
    End If
End Sub

Private Sub picV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        msngStartX = x
    End If
End Sub

Private Sub picV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picV.Left + x - msngStartX
        If sngTemp > 1500 And ScaleWidth - (sngTemp + picV.Width) > 1600 Then
            picV.Left = sngTemp
            tvwMain_S.Width = picV.Left - tvwMain_S.Left
            Form_Resize
        End If
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePrintSet_Click()
    zlPrintSet
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFilePreview_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Find"
            mnuViewFind_Click
        Case "Quit"
            mnuFileExit_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnuFilePreview_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
    
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    cbrThis.Bands("only").MinHeight = tbrThis.Height
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim buttTemp As Button
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For Each buttTemp In tbrThis.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    cbrThis.Bands("only").MinHeight = tbrThis.Height
    Form_Resize
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
   ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(hwnd)
End Sub

Private Sub subPrint(bytMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim nod As Node
    
    Set nod = tvwMain_S.SelectedItem
    Do Until nod.Parent Is Nothing
        Set nod = nod.Parent
    Loop
    
    Set objPrint.Body = mshSum_S.msfObj
    objPrint.Title.Text = nod.Text & "类收费细目医保项目对应表"
    'objRow.Add "医院名称：" & gstr单位名称
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & gstrUserName
    objRow.Add "打印时间：" & Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    objPrint.BelowAppRows.Add objRow
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
    
Private Sub Fill大类()
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim strServer As String, strUser As String, strPass As String
    '只刷新列表内容
    
    '首先获得医保大类
    mshSum_S.Active = True
    If mint险类 = TYPE_成都南充 Then
        If mcnYB.State = 1 Then mcnYB.Close
        mcnYB.Open GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("LCConnectionString"), "dsn=lcyb;uID=hisuser;pwd=hiscdgk")
        Exit Sub
    End If
    
    gstrSQL = "select ID,编码,名称 from 保险支付大类 " & _
              "where 险类=[1] order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint险类)
    
    mshSum_S.Clear
    Do Until rsTemp.EOF
        mshSum_S.AddItem rsTemp("编码") & "." & rsTemp("名称")
        mshSum_S.ItemData(mshSum_S.NewIndex) = rsTemp("ID")
        rsTemp.MoveNext
    Loop
    If mint险类 = TYPE_成都德阳 Then
        If Not gcnOracle_成都德阳 Is Nothing Then Exit Sub
    ElseIf mint险类 = TYPE_南充阆中 Then
        If Not gcnOracle_南充阆中 Is Nothing Then Exit Sub
    Else
        If Not gcnOracle_广元旺苍 Is Nothing Then Exit Sub
    End If
    '重庆新打开医保
    gstrSQL = "Select 参数名,参数值 From 保险参数 Where 险类=" & mint险类
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint险类)
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        Select Case rsTemp("参数名")
            Case "医保服务器"
                strServer = strTemp
            Case "医保用户名"
                strUser = strTemp
            Case "医保用户密码"
                strPass = strTemp
        End Select
        rsTemp.MoveNext
    Loop
    If mint险类 = TYPE_成都德阳 Then
        Set gcnOracle_成都德阳 = New ADODB.Connection
        If OraDataOpen(gcnOracle_成都德阳, strServer, strUser, strPass) = False Then
            Exit Sub
        End If
    ElseIf mint险类 = TYPE_南充阆中 Then
        Set gcnOracle_南充阆中 = New ADODB.Connection
        If OraDataOpen(gcnOracle_南充阆中, strServer, strUser, strPass) = False Then
            Exit Sub
        End If
    Else
        Set gcnOracle_广元旺苍 = New ADODB.Connection
        If OraDataOpen(gcnOracle_广元旺苍, strServer, strUser, strPass) = False Then
            Exit Sub
        End If
    End If
End Sub

Private Function FillTree() As Boolean
'功能:装入收费类别和收费细目的所有分类到tvwMain_S
    '本程序中树节点比其它程序的KEY值多一个字符，即第二位的类别编码

    Dim rsTemp As New ADODB.Recordset
    Dim strKey As String
    Dim nod As Node
    
    On Error GoTo errHandle
    rsTemp.CursorLocation = adUseClient
    
    MousePointer = vbHourglass
    
    mstrKey = ""     '全面刷新时就相当于用户没点过任何节点
    If Not tvwMain_S.SelectedItem Is Nothing Then
        strKey = tvwMain_S.SelectedItem.Key
    End If
    
    gstrSQL = "select 编码,类别 from 收费类别 where 编码<>'5' and 编码<>'6' and 编码<>'7' order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    LockWindowUpdate tvwMain_S.hwnd
    '删除所有节点
    With tvwMain_S.Nodes
        .Clear
        '增加类别
        Do Until rsTemp.EOF
            .Add , , "R" & rsTemp("编码"), "【" & rsTemp("编码") & "】" & rsTemp("类别"), "R", "R"
            tvwMain_S.Nodes("R" & rsTemp("编码")).Sorted = True
            rsTemp.MoveNext
        Loop
        .Add , , "D5", "【5】西成药", "R", "R"
        tvwMain_S.Nodes("D5").Sorted = True
        .Add , , "E6", "【6】中成药", "R", "R"
        tvwMain_S.Nodes("E6").Sorted = True
        .Add , , "F7", "【7】中草药", "R", "R"
        tvwMain_S.Nodes("F7").Sorted = True
        
        '增加普通收费项目分类节点
        gstrSQL = "select id,上级id,类别,编码,名称 from 收费细目  where 类别<>'5' and 类别<>'6' and 类别<>'7' and 末级 <> 1 " & _
             " start with 上级ID is null  connect by prior id=上级ID "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
        Do Until rsTemp.EOF
            '添加节点
            If IsNull(rsTemp("上级id")) Then
                .Add "R" & rsTemp("类别"), tvwChild, "C" & rsTemp("类别") & rsTemp("id"), "【" & rsTemp("编码") & "】" & rsTemp("名称"), "C", "C"
            Else
                .Add "C" & rsTemp("类别") & rsTemp("上级id"), tvwChild, "C" & rsTemp("类别") & rsTemp("id"), "【" & rsTemp("编码") & "】" & rsTemp("名称"), "C", "C"
            End If
            tvwMain_S.Nodes("C" & rsTemp("类别") & rsTemp("ID")).Sorted = True
            rsTemp.MoveNext
        Loop
    
        '再装入药品用途分类的数据
        gstrSQL = "select id,上级id,材质,编码,名称 from 药品用途分类  " & _
             " start with 上级ID is null  connect by prior id=上级ID "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
        Do Until rsTemp.EOF
            '添加节点
            Select Case rsTemp("材质")
                Case "中成药"
                    If IsNull(rsTemp("上级id")) Then
                        Set nod = .Add("E6", tvwChild, "E6" & rsTemp("id"), "【" & rsTemp("编码") & "】" & rsTemp("名称"), "P", "P")
                    Else
                        Set nod = .Add("E6" & rsTemp("上级id"), tvwChild, "E6" & rsTemp("id"), "【" & rsTemp("编码") & "】" & rsTemp("名称"), "P", "P")
                    End If
                Case "中草药"
                    If IsNull(rsTemp("上级id")) Then
                        Set nod = .Add("F7", tvwChild, "F7" & rsTemp("id"), "【" & rsTemp("编码") & "】" & rsTemp("名称"), "P", "P")
                    Else
                        Set nod = .Add("F7" & rsTemp("上级id"), tvwChild, "F7" & rsTemp("id"), "【" & rsTemp("编码") & "】" & rsTemp("名称"), "P", "P")
                    End If
                Case Else '西成药
                    If IsNull(rsTemp("上级id")) Then
                        Set nod = .Add("D5", tvwChild, "D5" & rsTemp("id"), "【" & rsTemp("编码") & "】" & rsTemp("名称"), "P", "P")
                    Else
                        Set nod = .Add("D5" & rsTemp("上级id"), tvwChild, "D5" & rsTemp("id"), "【" & rsTemp("编码") & "】" & rsTemp("名称"), "P", "P")
                    End If
                End Select
            nod.Sorted = True
            rsTemp.MoveNext
        Loop
    End With
    
    LockWindowUpdate 0
    MousePointer = 0
    
    On Error Resume Next
    Set nod = tvwMain_S.Nodes(strKey)
    If Err <> 0 Then
        Set nod = tvwMain_S.Nodes(1)
        nod.Selected = True
    Else
        Err.Clear
        nod.Selected = True
        nod.EnsureVisible
    End If
    Call FillSum
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    LockWindowUpdate 0
    MousePointer = 0
End Function

Public Sub FillSum(Optional ByVal blnForce As Boolean = False)
'功能:装入各种统计数据
    Dim rsTemp As New ADODB.Recordset
    Dim nod As Node
    Dim str材质分类 As String
    Dim lngID As Long

    If tvwMain_S.SelectedItem Is Nothing Then
        ClearGrid mshSum_S
        Call MenuSet
        Exit Sub
    End If
    
    If blnForce = False Then
        If mstrKey = tvwMain_S.SelectedItem.Key And mint中心 = cmb险类.ItemData(cmb险类.ListIndex) Then
            '完全没有改变，不用再刷新
            Exit Sub
        End If
        
        If cmdSave.Enabled = True Then
            '已经修改，提示是否需要保存当前的设置
            If MsgBox("保险项目已经修改，是否需要保存？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                Call cmdSave_Click
            End If
        End If
    End If
    
    cmdSave = False
    cmdRestore = False
    '提取保险项目的两个主关键字
    mstrKey = tvwMain_S.SelectedItem.Key
    mint中心 = cmb险类.ItemData(cmb险类.ListIndex)
    
    Set nod = tvwMain_S.SelectedItem
    
    '根据不同的节点，做出不同的显示
    '手术类别要多显示一列
    If Mid(nod.Key, 2, 1) = "5" Or Mid(nod.Key, 2, 1) = "6" Or Mid(nod.Key, 2, 1) = "7" Then
        '药品的处理要麻烦一些
        mshSum_S.TextMatrix(0, col产地) = "产地"
        
        Select Case Left(nod.Key, 1)
            Case "D"
                str材质分类 = "西成药"
            Case "E"
                str材质分类 = "中成药"
            Case "F"
                str材质分类 = "中草药"
        End Select
        
        If nod.Image = "R" Then
            gstrSQL = "select A.药品ID as ID,A.编码,B.通用名称||decode(M.名称,null,'',b.通用名称,'',' 〖'||M.名称||'〗') as 名称,A.规格,A.产地,A.售价单位 as 单位,D.是否变价,E.名称 剂型 " & _
                        "from 药品目录 A,药品信息 B,收费细目 D,药品剂型 E,(Select distinct 药品id,名称 from 药品别名 ) M " & _
                        "where A.药名ID=B.药名ID and d.id=M.药品ID(+) and B.剂型=E.编码(+) and B.材质分类='" & str材质分类 & "'" & _
                        "      and A.药品ID=D.ID and (A.撤档时间 is null or A.撤档时间=to_date('3000-01-01','yyyy-mm-dd'))"
        Else
            gstrSQL = "select A.药品ID as ID,A.编码,B.通用名称||decode(M.名称,null,'',b.通用名称,'',' 〖'||M.名称||'〗') as 名称,A.规格,A.产地,A.售价单位 as 单位,D.是否变价,E.名称 剂型 " & _
                      "from 药品目录 A,药品信息 B,收费细目 D,药品剂型 E,(Select distinct 药品id,名称 from 药品别名) M ,(select ID from 药品用途分类 start with ID=" & Mid(nod.Key, 3) & " connect by prior id=上级ID) C " & _
                      "where A.药名ID=B.药名ID and B.剂型=E.编码(+) and d.id=M.药品ID(+) and B.材质分类='" & str材质分类 & "' and B.用途分类ID=C.ID" & _
                      "       and A.药品ID=D.ID and (A.撤档时间 is null or A.撤档时间=to_date('3000-01-01','yyyy-mm-dd'))"
        End If
        
    Else
        '非药品就容易得多了
        mshSum_S.TextMatrix(0, col产地) = "说明"
        
        If nod.Image = "R" Then
            gstrSQL = "select id,编码,名称,规格,说明 as 产地,计算单位 as 单位,是否变价,'' 剂型 from 收费细目 where 末级=1 and 类别='" & Mid(nod.Key, 2, 1) & "' " & _
                        " and (撤档时间 is null or 撤档时间=to_date('3000-01-01','yyyy-mm-dd'))"
        Else
            gstrSQL = "select id,编码,名称,规格,说明 as 产地,计算单位 as 单位,是否变价,'' 剂型 from 收费细目 where 末级=1 and (撤档时间 is null or 撤档时间=to_date('3000-01-01','yyyy-mm-dd'))" & _
                        " start with 上级ID=" & Mid(nod.Key, 3) & " connect by prior id=上级ID "
        End If
    End If
    Dim strTable As String
    If mint险类 = TYPE_成都德阳 Then
        strTable = "医保支付项目_德阳"
    Else
        strTable = "医保支付项目"
    End If
    gstrSQL = "select A.ID,A.编码,A.名称,A.规格,A.剂型,A.产地,A.单位,A.是否变价,D.价格,'' as 改变方式" & _
               " ,B.项目编码,B.项目名称,B.附注,B.项目编码 as 原编码,B.是否医保,B.大类ID,C.编码 as 大类编码,C.名称 as 大类名称 " & _
               " from (" & gstrSQL & ") A," & strTable & " B,保险支付大类 C," & _
               "      (select sum(现价) as 价格,收费细目ID from 收费价目 where 执行日期<=sysdate and (终止日期>=sysdate or 终止日期 is null) group by 收费细目ID) D " & _
               " Where A.ID=B.收费细目ID(+) and B.大类ID=c.id(+)   and b.险类(+)=[1] and  B.中心(+)= [2]" & _
               "       and A.ID=D.收费细目ID(+)  "
    
    MousePointer = 11
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint险类, mint中心)
    
    lngID = mshSum_S.RowData(mshSum_S.Row)
    Call FillGrid(rsTemp, lngID)
    
    stbThis.Panels(2).Text = "共有收费项目" & rsTemp.RecordCount & "条"
    
    MousePointer = 0
    Call MenuSet
End Sub

Private Sub FillGrid(rsTemp As ADODB.Recordset, ByVal lngID As Long)
    Dim strSort As String
    Dim strDemo As String
    Dim intMatch As Integer
    Dim lngRow As Long, lngRowSelect As Long
    
    Select Case mlngCol
        Case cOL编码
            strSort = "编码"
        Case cOL名称
            strSort = "名称"
        Case COL规格
            strSort = "规格"
        Case col产地
            strSort = "产地"
        Case COL单位
            strSort = "单位"
        Case col价格
            strSort = "价格"
        Case COL医保编码
            strSort = "项目编码"
        Case col医保名称
            strSort = "项目名称"
        Case col大类名称
            strSort = "大类名称"
        Case col非医保
            strSort = "是否医保"
        Case Else
            rsTemp.Sort = "编码"
    End Select
    rsTemp.Sort = strSort & IIf(mblnDesc, " DESC", "")
    
    mshSum_S.TxtVisible = False
    mshSum_S.CboVisible = False
    mshSum_S.Redraw = False
    ClearGrid mshSum_S
    If rsTemp.RecordCount <> 0 Then
        mshSum_S.Rows = rsTemp.RecordCount + 1
    End If
    lngRow = 1
    With mshSum_S
        Do Until rsTemp.EOF
            If rsTemp("ID") = lngID Then
                lngRowSelect = lngRow
            End If
            
            .RowData(lngRow) = rsTemp("ID")
            .TextMatrix(lngRow, cOL编码) = rsTemp("编码")
            .TextMatrix(lngRow, cOL名称) = rsTemp("名称")
            .TextMatrix(lngRow, COL规格) = IIf(IsNull(rsTemp("规格")), "", rsTemp("规格"))
            .TextMatrix(lngRow, col产地) = IIf(IsNull(rsTemp("产地")), "", rsTemp("产地"))
            .TextMatrix(lngRow, COL剂型) = IIf(IsNull(rsTemp("剂型")), "", rsTemp("剂型"))
            .TextMatrix(lngRow, COL单位) = IIf(IsNull(rsTemp("单位")), "", rsTemp("单位"))
            .TextMatrix(lngRow, col价格) = IIf(rsTemp("是否变价") = 0, Format(rsTemp("价格"), "0.000"), "")
            .TextMatrix(lngRow, col改变方式) = IIf(IsNull(rsTemp("改变方式")), "", rsTemp("改变方式"))
            .TextMatrix(lngRow, COL医保编码) = IIf(IsNull(rsTemp("项目编码")), "", rsTemp("项目编码"))
            .TextMatrix(lngRow, col医保名称) = IIf(IsNull(rsTemp("项目名称")), "", rsTemp("项目名称"))
            .TextMatrix(lngRow, col原编码) = IIf(IsNull(rsTemp("原编码")), "", rsTemp("原编码"))
            .TextMatrix(lngRow, col大类ID) = IIf(IsNull(rsTemp("大类ID")), "", rsTemp("大类ID"))
            .TextMatrix(lngRow, col非医保) = IIf(rsTemp("是否医保") = "0", "√", "")
            .TextMatrix(lngRow, col医保附注) = IIf(IsNull(rsTemp("附注")), "", rsTemp("附注"))
            .TextMatrix(lngRow, col匹配序列号) = ""
            
            If IsNull(rsTemp("大类编码")) Or IsNull(rsTemp("大类名称")) Then
                .TextMatrix(lngRow, col大类名称) = ""
            Else
                .TextMatrix(lngRow, col大类名称) = rsTemp("大类编码") & "." & rsTemp("大类名称")
            End If
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    If lngRowSelect > 0 And lngRowSelect < mshSum_S.Rows - 1 Then
        mshSum_S.msfObj.TopRow = lngRowSelect
        mshSum_S.Row = lngRowSelect
    End If
    mshSum_S.Redraw = True
End Sub

Private Sub ClearGrid(objGrid As Object)
'功能：清除表格,并完成部分初始化
    Dim i As Long
    
    cmdRestore.Enabled = False
    cmdSave.Enabled = False
    With objGrid.msfObj
        .Rows = 2
        .RowData(1) = 0
        For i = 0 To objGrid.Cols - 1
            objGrid.TextMatrix(1, i) = ""
        Next
    
    End With
End Sub

Private Sub MenuSet()
'功能:显示菜单和工具栏的状态(打印)
    Dim blnPrint As Boolean
    
    blnPrint = Not (mshSum_S.Rows = 2 And mshSum_S.TextMatrix(1, 0) = "")

    mnuFilePreview.Enabled = blnPrint
    mnuFilePrint.Enabled = blnPrint
    mnuFileExcel.Enabled = blnPrint
    tbrThis.Buttons("Preview").Enabled = blnPrint
    tbrThis.Buttons("Print").Enabled = blnPrint
    
    If InStr(mstr权限, "增删改") > 0 Then
        mshSum_S.Active = blnPrint
        If mint中心 = TYPE_泸州市 Then
            '强制不能使用
            If gcn泸州.State = adStateClosed Then mshSum_S.Active = False
        End If
    Else
        mshSum_S.Active = False
    End If
End Sub

Public Sub ShowForm(frmParent As Form, ByVal int险类 As Integer)
    
    Dim rsTemp As New ADODB.Recordset
    mint险类 = int险类
    
    gstrSQL = "select 序号,名称 from 保险中心目录 where 序号<>0 and 险类=[1] order by 序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取保险类别", mint险类)
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "没有可用社保机构，请在参数中下载，不能使用本功能。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If frm保险项目_广元旺苍.Visible = True Then
        frm保险项目_广元旺苍.Show
        Exit Sub
    End If
    
    mstr权限 = gstrPrivs
    frm保险项目_广元旺苍.Show , frmParent
End Sub


Public Function CheckForm(ByVal int险类 As Integer) As Boolean
    
    Dim rsTemp As New ADODB.Recordset
    mint险类 = int险类
    
    gstrSQL = "select 序号,名称 from 保险中心目录 where 序号<>0 and 险类=[1] order by 序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取保险类别", mint险类)
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "没有可用社保机构，请在参数中下载，不能使用本功能。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If frm保险项目_广元旺苍.Visible = True Then
        CheckForm = True
        Exit Function
    End If
    
    mstr权限 = gstrPrivs
    CheckForm = True
End Function
