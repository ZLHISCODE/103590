VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm特殊业务申请 
   Caption         =   "特殊业务申请"
   ClientHeight    =   5940
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9345
   Icon            =   "frm特殊业务申请.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   9345
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ImageList ImgColor 
      Left            =   660
      Top             =   690
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm特殊业务申请.frx":1CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm特殊业务申请.frx":1F14
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm特殊业务申请.frx":212E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm特殊业务申请.frx":2448
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm特殊业务申请.frx":2662
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm特殊业务申请.frx":2F3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm特殊业务申请.frx":3256
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm特殊业务申请.frx":3470
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm特殊业务申请.frx":368A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgBlack 
      Left            =   90
      Top             =   690
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm特殊业务申请.frx":38A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm特殊业务申请.frx":3ABE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm特殊业务申请.frx":3CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm特殊业务申请.frx":3FF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm特殊业务申请.frx":420C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm特殊业务申请.frx":4AE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm特殊业务申请.frx":4E00
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm特殊业务申请.frx":501A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm特殊业务申请.frx":5234
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   1296
      BandCount       =   1
      _CBWidth        =   9345
      _CBHeight       =   735
      _Version        =   "6.7.8988"
      Child1          =   "tbrThis"
      MinHeight1      =   675
      Width1          =   615
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   675
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   1191
         ButtonWidth     =   1455
         ButtonHeight    =   1191
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ImgBlack"
         HotImageList    =   "ImgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "家庭病床"
               Key             =   "Home"
               Object.ToolTipText     =   "申请家庭病床"
               Object.Tag             =   "家庭病床"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "规定病"
               Key             =   "Spec"
               Object.ToolTipText     =   "申请门诊规定病"
               Object.Tag             =   "规定病"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "特治特检"
               Key             =   "Especial"
               Object.ToolTipText     =   "申请特治特检"
               Object.Tag             =   "特治特检"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "申请转院"
               Key             =   "Switch"
               Object.ToolTipText     =   "申请转院、转外就诊"
               Object.Tag             =   "申请转院"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Filter"
               Object.ToolTipText     =   "过滤"
               Object.Tag             =   "过滤"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh审核清单 
      Height          =   2475
      Left            =   150
      TabIndex        =   2
      Top             =   900
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   4366
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   7513801
      BackColorSel    =   14783374
      BackColorBkg    =   16777215
      GridColorFixed  =   0
      FocusRect       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   5580
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   635
      SimpleText      =   $"frm特殊业务申请.frx":544E
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm特殊业务申请.frx":5495
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11430
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
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFileSplitSet 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileSplitPrint 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFileSplitReport 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuRequest 
      Caption         =   "业务申请(&R)"
      Begin VB.Menu mnuRequestHome 
         Caption         =   "家庭病床(&H)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRequestSpec 
         Caption         =   "门诊特殊病(&S)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRequestEspecial 
         Caption         =   "特治特检(&E)"
         Shortcut        =   {DEL}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRequestSwitch 
         Caption         =   "转院申请(&W)"
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
         Begin VB.Menu mnuViewTool_1 
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
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "过滤(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuView_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R)"
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
         Caption         =   "&WEB上的中联"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frm特殊业务申请"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Private Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Private Const madDbDateDefault As Integer = 20               '日期型字段缺省长度
Private str开始日期 As String, str结束日期 As String, str审核标志 As String
Private rsVerify As New ADODB.Recordset

Private Sub Form_Load()
    Call initGird
End Sub

Private Sub initGird()
    Dim intCol As Integer, intCols As Integer
    Dim arrColumn
    Const strColumns As String = "姓名|性别|身份证号|疾病名称|转入医院|申请日期|审核标志"
    
    arrColumn = Split(strColumns, "|")
    intCols = UBound(arrColumn)
    With msh审核清单
        .Rows = 2
        .Cols = intCols + 1
        
        For intCol = 0 To intCols
            .TextMatrix(0, intCol) = arrColumn(intCol)
            .ColAlignmentFixed(intCol) = 4
        Next
        
        .Col = 0
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    With Me.msh审核清单
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top
        .Width = Me.ScaleWidth
    End With
End Sub

Private Sub mnuFileQuit_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub mnuHelpTitle_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuRequestSwitch_Click()
    Dim blnReturn As Boolean
    blnReturn = frm转院申请.ShowME(1, Me)
End Sub

Private Sub mnuViewFind_Click()
    Dim blnReturn As Boolean
    blnReturn = frm申请转院_过滤.ShowME(str开始日期, str结束日期, str审核标志)
    Call RefreshData
End Sub

Private Sub mnuViewRefresh_Click()
    Call RefreshData
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = mnuViewStatus.Checked Xor True
    stbThis.Visible = stbThis.Visible Xor True
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = mnuViewToolButton.Checked Xor True
    cbrThis.Visible = cbrThis.Visible Xor True
    Call Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim tbrbutton As Button
    mnuViewToolText.Checked = mnuViewToolText.Checked Xor True
    For Each tbrbutton In tbrThis.Buttons
        tbrbutton.Caption = IIf(mnuViewToolText.Checked, tbrbutton.Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call Form_Resize
End Sub

Private Sub mnuFileExcel_Click()
    Call subPrint(3)
End Sub

Private Sub mnuFilePreview_Click()
    Call subPrint(2)
End Sub

Private Sub mnuFilePrint_Click()
    Call subPrint(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub subPrint(bytStyle As Byte)
    '功能：输入出列表
    '参数：bytStyle=1-打印,2-预览,3-输出到Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    
    If gstrUserName = "" Then GetUserInfo
    intRow = msh审核清单.Row
    
    '表头
    objOut.Title.Text = "转院申请清单"
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表项
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zldatabase.Currentdate, "yyyy年MM月DD日")
    objOut.BelowAppRows.Add objRow
    
    '表体
    Set objOut.Body = msh审核清单
    
    '输出
    msh审核清单.Redraw = False
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    msh审核清单.Redraw = True
    
    msh审核清单.Row = intRow
    msh审核清单.Col = 0: msh审核清单.ColSel = msh审核清单.Cols - 1
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage Me.hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo Me.hwnd
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Print"
        Call mnuFilePrint_Click
    Case "Preview"
        Call mnuFilePreview_Click
    Case "Switch"
        Call mnuRequestSwitch_Click
    Case "Filter"
        Call mnuViewFind_Click
    Case "Help"
        Call mnuHelpTitle_Click
    Case "Exit"
        Call mnuFileQuit_Click
    End Select
End Sub

Private Sub RefreshData()
    Dim intCol As Integer, intCols As Integer
    Dim strColumns As String, strData As String
    Dim strFields As String, strValues As String
    Dim arrColumn
    
    '提取所有申请转院的数据
    If str开始日期 = "" Then Exit Sub
    If str审核标志 = "" Then Exit Sub
    
    If Not 医保初始化_沈阳市 Then Exit Sub
    
    DoEvents
    Call zlcommfun.ShowFlash("正在从中心提取转院审批信息,请稍候 ...", Me)
    DoEvents
    
'    1   hospital_id医疗机构编码   20  否
'    2   from_date  查询起始日期       是  格式：YYYY-MM-DD
'    3   to_date    查询终止日期       是  格式：YYYY-MM-DD
'    4   audit_flag 审核标志       4   是  "0"：未审核    "1"：审核通过    "2"：审核未通过    "all"：全部
    Call DebugTool("准备调用获取审批信息功能")
    gstrField_沈阳市 = "hospital_id||from_date||to_date||audit_flag"
    gstrValue_沈阳市 = gCominfo_沈阳市.医院编码 & "||" & str开始日期 & "||" & str结束日期 & "||" & str审核标志
    If Not 调用接口_准备_沈阳市(Function_沈阳市.转院申请_查询审核信息) Then GoTo StopFlash
    If Not 调用接口_写入口参数_沈阳市(1) Then GoTo StopFlash
    If Not 调用接口_执行_沈阳市 Then GoTo StopFlash
    If Not 调用接口_指定记录集_沈阳市("ToanotherHosInfo") Then GoTo StopFlash
    
    '初始化记录集
    Call DebugTool("准备初始化内部记录集")
    strColumns = "name|sex|idcard|disease|to_hospital_name|input_date|audit_flag"
    arrColumn = Split(strColumns, "|")
    intCols = UBound(arrColumn)
    For intCol = 0 To intCols
        strFields = strFields & "|" & arrColumn(intCol) & "," & adLongVarChar & ",50"
    Next
    strFields = Mid(strFields, 2)
    Call Record_Init(rsVerify, strFields)
    
'    1   name           姓名    10
'    2   sex            性别    2
'    3   birthday       出生日期        格式：YYYY-MM-DD
'    4   idcard         身份证号码  20
'    5   insr_code      保险号  30
'    6   corp_name      单位名称    50
'    7   pers_name      人员类别    20
'    8   official_name  公务员级别  20
'    9   indi_id        个人编号    8
'    10  serial_apply   申请序列号  12
'    11  busi_type      业务类型    2   "16"：特制特检
'    12  apply_type     申请类型    1   "0"：普通申请    "1"：追加申请
'    13  apply_content  申请内容    1   "1"：特制特检
'    14  icd            疾病编码    20
'    15  disease        疾病名称    50
'    16  disease_deac   病情摘要    500
'    17  oper_desc      主要诊疗方案    500
'    18  doctor_name    申请医师    10
'    19  apply_opinion  申请理由    500
'    20  intend_fee     预计费用    10
'    21  apply_date     申请有效期限        格式：YYYY-MM-DD
'    22  admit_date     审核有效期限        格式：YYYY-MM-DD
'    23  audit_date     审批日期        格式：YYYY-MM-DD
'    24  audit_flag     审批标志    10
'    25  input_man      录入人姓名  10
'    26  input_date     录入日期        格式：YYYY-MM-DD
'    27  note           备注    500
    '将接口返回的数据填到映射记录集中
    If 调用接口_记录数_沈阳市 Then
        Call DebugTool("返回记录条数：" & CZ_GetRowCount(glngInterface_沈阳市))
        Call 调用接口_移动记录集_沈阳市(MoveFirst)
        Do While True
            strValues = ""
            For intCol = 0 To intCols
                'todo 此处是正式代码，需要取消注释
                Call 调用接口_读取数据_沈阳市(arrColumn(intCol), strData)
                strValues = strValues & "|" & strData
            Next
            strValues = Mid(strValues, 2)
            Call Record_Add(rsVerify, strColumns, strValues)
            
            Call DebugTool("已加入一行记录")
            If Not 调用接口_移动记录集_沈阳市(MoveNext) Then Exit Do
        Loop
    End If
    
    '绑定数据
    If rsVerify.RecordCount = 0 Then
        Call DebugTool("内部记录集的记录数为空")
        With msh审核清单
            .Clear
            .Rows = 2
            For intCol = 0 To .Cols - 1
                .TextMatrix(1, intCol) = ""
            Next
        End With
    Else
        Call DebugTool("显示内部记录集中的数据")
        Set msh审核清单.DataSource = rsVerify
    End If
    
    '设置列头
    Call DebugTool("重新设置表头")
    strColumns = "姓名|性别|身份证号|疾病名称|转入医院|申请日期|审核标志"
    arrColumn = Split(strColumns, "|")
    With msh审核清单
        For intCol = 0 To .Cols - 1
            .TextMatrix(0, intCol) = arrColumn(intCol)
            .ColAlignmentFixed(intCol) = 4
        Next
    End With
    Call zlControl.MshSetColWidth(msh审核清单, Me)
StopFlash:
    Call zlcommfun.StopFlash
End Sub

Private Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '添加记录
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Private Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intTYPE As Integer, lngLength As Long
    '初始化映射记录集
    'strFields:字段名,类型,长度|字段名,类型,长度    如果长度为零,则取默认长度
    '字符型:adLongVarChar;数字型:adDouble;日期型:adDBDate
    
    '例子：
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|科目ID," & adDouble & ",18|摘要, " & adLongVarChar & ",50|" & _
    '"删除," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intTYPE = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '获取字段缺省长度
            If lngLength = 0 Then
                Select Case intTYPE
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intTYPE, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub
