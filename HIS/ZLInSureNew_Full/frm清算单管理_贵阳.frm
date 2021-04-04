VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm清算单管理_贵阳 
   Caption         =   "清算单管理_贵阳"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   10575
   Icon            =   "frm清算单管理_贵阳.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   10575
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab tabShow 
      Height          =   345
      Left            =   30
      TabIndex        =   3
      Top             =   720
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   609
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "医疗保险"
      TabPicture(0)   =   "frm清算单管理_贵阳.frx":076A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "生育保险"
      TabPicture(1)   =   "frm清算单管理_贵阳.frx":0786
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "工伤保险"
      TabPicture(2)   =   "frm清算单管理_贵阳.frx":07A2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin MSComctlLib.ImageList imgBlack 
      Left            =   2820
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清算单管理_贵阳.frx":07BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清算单管理_贵阳.frx":09D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清算单管理_贵阳.frx":0BF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清算单管理_贵阳.frx":0E0C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   2250
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清算单管理_贵阳.frx":1026
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清算单管理_贵阳.frx":1240
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清算单管理_贵阳.frx":145A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm清算单管理_贵阳.frx":1674
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrTool 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   10575
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "tbrTool"
      MinHeight1      =   645
      Width1          =   915
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrTool 
         Height          =   645
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   1138
         ButtonWidth     =   820
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgBlack"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "申报"
               Key             =   "Add"
               Object.ToolTipText     =   "申报清算单"
               Object.Tag             =   "申报"
               ImageIndex      =   1
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Add1"
                     Text            =   "医疗保险"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Add2"
                     Text            =   "生育保险"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Add3"
                     Text            =   "工伤保险"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "撤销"
               Key             =   "Del"
               Object.ToolTipText     =   "撤销清算单"
               Object.Tag             =   "撤销"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Filter"
               Object.ToolTipText     =   "过滤"
               Object.Tag             =   "过滤"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6435
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm清算单管理_贵阳.frx":188E
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13573
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
            AutoSize        =   2
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   5385
      Left            =   30
      TabIndex        =   4
      Top             =   1050
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   9499
      _Version        =   393216
      FixedCols       =   0
      BackColorSel    =   13275520
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditAdd1 
         Caption         =   "医疗保险申报清算(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditAdd2 
         Caption         =   "生育保险申报清算"
      End
      Begin VB.Menu mnuEditAdd3 
         Caption         =   "工伤保险申报清算"
      End
      Begin VB.Menu muuEditSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDel 
         Caption         =   "撤销清算单(&B)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu muuEditSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditView 
         Caption         =   "查阅申报单(&V)"
      End
      Begin VB.Menu mnuEditGet 
         Caption         =   "查询办理情况(&G)"
      End
      Begin VB.Menu muuEditSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFilter 
         Caption         =   "过滤(&F)"
      End
      Begin VB.Menu muuEditSplit4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditRefresh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
End
Attribute VB_Name = "frm清算单管理_贵阳"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintInsure As Integer
Private mstrFilter As String
Private Enum 医疗保险
    ID
    期号
    保险类别代码
    保险类别
    操作员
    日期
    门诊人次
    门诊个人帐户
    门诊医疗补助
    特殊门诊人次
    特殊门诊个人帐户
    特殊门诊基本统筹
    特殊门诊大病统筹
    特殊门诊医疗补助
    控制线住院人次
    控制线住院个人帐户
    控制线住院基本统筹
    控制线住院大额统筹
    控制线住院医疗补助
    重症住院人次
    重症住院个人帐户
    重症住院基本统筹
    重症住院大额统筹
    重症住院医疗补助
    日包干住院人次
    日包干住院天数
    日包干住院个人帐户
    日包干住院医疗补助
    包干结算人次
    包干结算个人帐户
    包干结算基本统筹
    包干结算大额统筹
    包干结算医疗补助
    清算流水号
    处理情况
    列数
End Enum

Private Enum 生育保险
    ID
    期号
    保险类别代码
    保险类别
    操作员
    日期
    分娩包干人次
    分娩包干费用总额
    分娩包干统筹支付
    分娩非包干人次
    分娩非包干费用总额
    分娩非包干统筹支付
    计生人次
    计生费用总额
    计生统筹支付
    清算流水号
    处理情况
    列数
End Enum

Private Enum 工伤保险
    ID
    期号
    保险类别代码
    保险类别
    操作员
    日期
    门诊人次
    门诊统筹支付
    住院人次
    住院统筹支付
    清算流水号
    处理情况
    列数
End Enum

Private Enum 页面
    医疗保险    '含居民
    生育保险
    工伤保险
End Enum

'申报清算相关说明
'1、医疗保险申报清单中，门诊人次是指的普通门诊的人次？包干结算就诊人次是指普通门诊中选择的结算方式为单病种包干的部分数据
'   a、控制线（清算=1），重症（清算=2），按日包干（清算=4），包干（清算=6）
'   b、普通门诊选择了单病种的就是门诊包干
'   c、生育申报清单中
'2、生育申报清算中，
'   a、分娩住院包干（保险类别为生育，入院方式不是计划生育的，清算=5）
'   b、计生（入院方式为计划生育）
'   c、非包干（保险类别等于生育的-分娩包干-计生）

Public Sub ShowME(ByVal intinsure As Integer)
    On Error Resume Next
    mintInsure = intinsure
    Me.Show 1
End Sub

Private Sub Form_Load()
    Dim strMonth As String
    
    '缺省只提取两个月内的数据
    strMonth = Format(DateAdd("m", -1, zlDatabase.Currentdate()), "yyyyMM")
    mstrFilter = " And A.期号>='" & strMonth & "'"
    Call RefreshData
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    With mshDetail
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top - stbThis.Height
    End With
End Sub

Private Sub mnuEditAdd1_Click()
    If Not frm医疗保险申报单.ShowME(0) Then Exit Sub
    Call RefreshData
End Sub

Private Sub mnuEditAdd2_Click()
    If Not frm生育保险申报单.ShowME(0) Then Exit Sub
    Call RefreshData
End Sub

Private Sub mnuEditAdd3_Click()
    If Not frm工伤保险申报单.ShowME(0) Then Exit Sub
    Call RefreshData
End Sub

Private Sub mnuEditDel_Click()
    Dim lngID As Long
    Dim int保险类别代码 As Integer
    Dim str流水号 As String
    On Error GoTo errHand
    
    lngID = Val(mshDetail.TextMatrix(mshDetail.Row, 0))
    If lngID = 0 Then Exit Sub
    If tabShow.Tab = 页面.医疗保险 Then
        str流水号 = mshDetail.TextMatrix(mshDetail.Row, 医疗保险.清算流水号)
        int保险类别代码 = Val(mshDetail.TextMatrix(mshDetail.Row, 医疗保险.保险类别代码))
    ElseIf tabShow.Tab = 页面.生育保险 Then
        str流水号 = mshDetail.TextMatrix(mshDetail.Row, 生育保险.清算流水号)
        int保险类别代码 = Val(mshDetail.TextMatrix(mshDetail.Row, 生育保险.保险类别代码))
    Else
        str流水号 = mshDetail.TextMatrix(mshDetail.Row, 工伤保险.清算流水号)
        int保险类别代码 = Val(mshDetail.TextMatrix(mshDetail.Row, 工伤保险.保险类别代码))
    End If
    
    If MsgBox("你确定要清除该清算单吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
    
    If Not InitXML Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "APPNO", str流水号)
    If tabShow.Tab <> 页面.工伤保险 Then Call InsertChild(mdomInput.documentElement, "INSURETYPE", int保险类别代码)
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName)
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"))
    '调用接口
    If CommRecServer(IIf(tabShow.Tab = 页面.医疗保险, "DELRECM", IIf(tabShow.Tab = 页面.生育保险, "DELRECB", "DELRECG"))) = False Then Exit Sub
    
    gstrSQL = "ZL_清算单_DELETE(" & lngID & ")"
    gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    
    Call RefreshData
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mnuEditFilter_Click()
    Dim strReturn As String
    strReturn = frm清算单_过滤.ShowCondition
    If strReturn = "" Then Exit Sub
    
    mstrFilter = strReturn
    Call RefreshData
End Sub

Private Sub mnuEditGet_Click()
    Dim lngID As Long
    Dim str流水号 As String, str办理情况 As String
    On Error GoTo errHand
    
    lngID = Val(mshDetail.TextMatrix(mshDetail.Row, 0))
    If lngID = 0 Then Exit Sub
    If tabShow.Tab = 页面.医疗保险 Then
        str流水号 = mshDetail.TextMatrix(mshDetail.Row, 医疗保险.清算流水号)
    ElseIf tabShow.Tab = 页面.生育保险 Then
        str流水号 = mshDetail.TextMatrix(mshDetail.Row, 生育保险.清算流水号)
    Else
        str流水号 = mshDetail.TextMatrix(mshDetail.Row, 工伤保险.清算流水号)
    End If
    
    If Val(mshDetail.TextMatrix(mshDetail.Row, 0)) = 0 Then Exit Sub
    If str流水号 = "" Then Exit Sub
    
    If Not InitXML Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "APPNO", str流水号)
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName)
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"))
    '调用接口
    If CommRecServer("QUERYREC") = False Then Exit Sub
    str办理情况 = GetElemnetValue("STATUS")
    gstrSQL = "ZL_清算单_UPDATE(" & lngID & ",'" & str办理情况 & "')"
    gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    
    Call RefreshData
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mnuEditRefresh_Click()
    Call RefreshData
End Sub

Private Sub mnuEditView_Click()
    Dim lngID As Long
    lngID = Val(mshDetail.TextMatrix(mshDetail.Row, 0))
    If lngID = 0 Then Exit Sub
    
    If tabShow.Tab = 页面.医疗保险 Then
        Call frm医疗保险申报单.ShowME(lngID)
    ElseIf tabShow.Tab = 页面.生育保险 Then
        Call frm生育保险申报单.ShowME(lngID)
    Else
        Call frm工伤保险申报单.ShowME(lngID)
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mshDetail_DblClick()
    Call mnuEditView_Click
End Sub

Private Sub mshDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call mnuEditView_Click
End Sub

Private Sub tabShow_Click(PreviousTab As Integer)
    Call RefreshData
End Sub

Private Sub tbrTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Add"
        If tabShow.Tab = 页面.医疗保险 Then
            Call mnuEditAdd1_Click
        ElseIf tabShow.Tab = 页面.生育保险 Then
            Call mnuEditAdd2_Click
        Else
            Call mnuEditAdd3_Click
        End If
    Case "Del"
        Call mnuEditDel_Click
    Case "Exit"
        Call mnuFileExit_Click
    Case "Filter"
        Call mnuEditFilter_Click
    End Select
End Sub

Private Sub RefreshData()
    Dim lngRow As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    Call InitBill
    
    '提出所有数据
    If tabShow.Tab = 页面.医疗保险 Then
        gstrSQL = "SELECT  " & _
                 "        A.ID, A.期号, A.保险类别,A.保险类别名称, A.操作员, A.日期 ,B.门诊人次, B.门诊个人帐户, B.门诊医疗补助, B.特殊门诊人次, B.特殊门诊个人帐户, B.特殊门诊基本统筹, B.特殊门诊大额统筹,  " & _
                 "        B.特殊门诊医疗补助, B.控制线住院人次, B.控制线住院个人帐户, B.控制线住院基本统筹, B.控制线住院大额统筹, B.控制线住院医疗补助,  " & _
                 "        B.重症住院人次, B.重症住院个人帐户, B.重症住院基本统筹, B.重症住院大额统筹, B.重症住院医疗补助, B.日包干住院人次, B.日包干住院天数,  " & _
                 "        B.日包干住院个人帐户, 日包干住院医疗补助, B.包干结算人次, B.包干结算个人帐户, B.包干结算基本统筹, B.包干结算大额统筹, B.包干结算医疗补助, A.清算流水号, A.处理情况 " & _
                 " FROM 清算单 A, 基本医疗清算明细 B " & _
                 " WHERE A.ID=B.清算单ID " & mstrFilter & " AND A.性质=" & tabShow.Tab & _
                 " Order by A.期号 Desc,A.保险类别名称"
    ElseIf tabShow.Tab = 页面.生育保险 Then
        gstrSQL = "SELECT  " & _
                 "        A.ID, A.期号, A.保险类别,A.保险类别名称, A.操作员, A.日期 , B.分娩包干人次, B.分娩包干费用总额, B.分娩包干统筹支付, B.分娩非包干人次,  " & _
                 "        B.分娩非包干费用总额, B.分娩非包干统筹支付, B.计生人次, B.计生费用总额, B.计生统筹支付, A.清算流水号, A.处理情况 " & _
                 " FROM 清算单 A, 生育清算明细 B" & _
                 " WHERE A.ID=B.清算单ID " & mstrFilter & " And A.性质=" & tabShow.Tab & _
                 " Order by A.期号 Desc,A.保险类别名称"
    Else
        gstrSQL = "SELECT  " & _
                 "        A.ID, A.期号, A.保险类别,A.保险类别名称, A.操作员, A.日期 , B.门诊人次, B.门诊统筹支付,  B.住院人次, B.住院统筹支付, A.清算流水号, A.处理情况 " & _
                 " FROM 清算单 A, 工伤清算明细 B" & _
                 " WHERE A.ID=B.清算单ID " & mstrFilter & " And A.性质=" & tabShow.Tab & _
                 " Order by A.期号 Desc,A.保险类别名称"
    End If
    Call OpenRecordset_OtherBase(rsTemp, "提出所有数据", gstrSQL, gcnGYYB)
    If rsTemp.RecordCount <> 0 Then Set mshDetail.DataSource = rsTemp
    Call InitBill(False)
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitBill(Optional ByVal blnInit As Boolean = True)
    With mshDetail
        If tabShow.Tab = 页面.医疗保险 Then
            If blnInit Then
                .Clear
                .Rows = 2: .Cols = 医疗保险.列数
                
                .TextMatrix(0, 医疗保险.ID) = "ID"
                .TextMatrix(0, 医疗保险.期号) = "期号"
                .TextMatrix(0, 医疗保险.保险类别代码) = "保险类别代码"
                .TextMatrix(0, 医疗保险.保险类别) = "保险类别"
                .TextMatrix(0, 医疗保险.操作员) = "操作员"
                .TextMatrix(0, 医疗保险.日期) = "日期"
                .TextMatrix(0, 医疗保险.门诊人次) = "门诊人次"
                .TextMatrix(0, 医疗保险.门诊个人帐户) = "门诊个人帐户"
                .TextMatrix(0, 医疗保险.门诊医疗补助) = "门诊医疗补助"
                .TextMatrix(0, 医疗保险.特殊门诊人次) = "特殊门诊人次"
                .TextMatrix(0, 医疗保险.特殊门诊个人帐户) = "个人帐户"
                .TextMatrix(0, 医疗保险.特殊门诊基本统筹) = "基本统筹"
                .TextMatrix(0, 医疗保险.特殊门诊大病统筹) = "大病统筹"
                .TextMatrix(0, 医疗保险.特殊门诊医疗补助) = "医疗补助"
                .TextMatrix(0, 医疗保险.控制线住院人次) = "控制线人次"
                .TextMatrix(0, 医疗保险.控制线住院个人帐户) = "个人帐户"
                .TextMatrix(0, 医疗保险.控制线住院基本统筹) = "基本统筹"
                .TextMatrix(0, 医疗保险.控制线住院大额统筹) = "大额统筹"
                .TextMatrix(0, 医疗保险.控制线住院医疗补助) = "医疗补助"
                .TextMatrix(0, 医疗保险.重症住院人次) = "重症住院人次"
                .TextMatrix(0, 医疗保险.重症住院个人帐户) = "个人帐户"
                .TextMatrix(0, 医疗保险.重症住院基本统筹) = "基本统筹"
                .TextMatrix(0, 医疗保险.重症住院大额统筹) = "大额统筹"
                .TextMatrix(0, 医疗保险.重症住院医疗补助) = "医疗补助"
                .TextMatrix(0, 医疗保险.日包干住院人次) = "日包干住院人次"
                .TextMatrix(0, 医疗保险.日包干住院天数) = "住院天数"
                .TextMatrix(0, 医疗保险.日包干住院个人帐户) = "个人帐户"
                .TextMatrix(0, 医疗保险.日包干住院医疗补助) = "医疗补助"
                .TextMatrix(0, 医疗保险.包干结算人次) = "包干结算人次"
                .TextMatrix(0, 医疗保险.包干结算个人帐户) = "个人帐户"
                .TextMatrix(0, 医疗保险.包干结算基本统筹) = "基本统筹"
                .TextMatrix(0, 医疗保险.包干结算大额统筹) = "大额统筹"
                .TextMatrix(0, 医疗保险.包干结算医疗补助) = "医疗补助"
                .TextMatrix(0, 医疗保险.清算流水号) = "清算流水号"
                .TextMatrix(0, 医疗保险.处理情况) = "处理情况"
            End If
            .ColWidth(医疗保险.ID) = 0
            .ColWidth(医疗保险.期号) = 800
            .ColWidth(医疗保险.保险类别代码) = 0
            .ColWidth(医疗保险.保险类别) = 1200
            .ColWidth(医疗保险.操作员) = 1000
            .ColWidth(医疗保险.日期) = 1000
            .ColWidth(医疗保险.门诊人次) = 1000
            .ColWidth(医疗保险.门诊个人帐户) = 1000
            .ColWidth(医疗保险.门诊医疗补助) = 1000
            .ColWidth(医疗保险.特殊门诊人次) = 1400
            .ColWidth(医疗保险.特殊门诊个人帐户) = 1000
            .ColWidth(医疗保险.特殊门诊基本统筹) = 1000
            .ColWidth(医疗保险.特殊门诊大病统筹) = 1000
            .ColWidth(医疗保险.特殊门诊医疗补助) = 1000
            .ColWidth(医疗保险.控制线住院人次) = 1400
            .ColWidth(医疗保险.控制线住院个人帐户) = 1000
            .ColWidth(医疗保险.控制线住院基本统筹) = 1000
            .ColWidth(医疗保险.控制线住院大额统筹) = 1000
            .ColWidth(医疗保险.控制线住院医疗补助) = 1000
            .ColWidth(医疗保险.重症住院人次) = 1400
            .ColWidth(医疗保险.重症住院个人帐户) = 1000
            .ColWidth(医疗保险.重症住院基本统筹) = 1000
            .ColWidth(医疗保险.重症住院大额统筹) = 1000
            .ColWidth(医疗保险.重症住院医疗补助) = 1000
            .ColWidth(医疗保险.日包干住院人次) = 1600
            .ColWidth(医疗保险.日包干住院天数) = 1000
            .ColWidth(医疗保险.日包干住院个人帐户) = 1000
            .ColWidth(医疗保险.日包干住院医疗补助) = 1000
            .ColWidth(医疗保险.包干结算人次) = 1400
            .ColWidth(医疗保险.包干结算个人帐户) = 1000
            .ColWidth(医疗保险.包干结算基本统筹) = 1000
            .ColWidth(医疗保险.包干结算大额统筹) = 1000
            .ColWidth(医疗保险.包干结算医疗补助) = 1000
            .ColWidth(医疗保险.清算流水号) = 2000
            .ColWidth(医疗保险.处理情况) = 2500
        ElseIf tabShow.Tab = 页面.生育保险 Then
            If blnInit Then
                .Clear
                .Rows = 2: .Cols = 生育保险.列数
                
                .TextMatrix(0, 生育保险.ID) = "ID"
                .TextMatrix(0, 生育保险.期号) = "期号"
                .TextMatrix(0, 生育保险.保险类别代码) = "保险类别代码"
                .TextMatrix(0, 生育保险.保险类别) = "保险类别"
                .TextMatrix(0, 生育保险.操作员) = "操作员"
                .TextMatrix(0, 生育保险.日期) = "日期"
                .TextMatrix(0, 生育保险.分娩包干人次) = "分娩包干人次"
                .TextMatrix(0, 生育保险.分娩包干费用总额) = "费用总额"
                .TextMatrix(0, 生育保险.分娩包干统筹支付) = "统筹支付"
                .TextMatrix(0, 生育保险.分娩非包干人次) = "分娩非包干人次"
                .TextMatrix(0, 生育保险.分娩非包干费用总额) = "费用总额"
                .TextMatrix(0, 生育保险.分娩非包干统筹支付) = "统筹支付"
                .TextMatrix(0, 生育保险.计生人次) = "计生人次"
                .TextMatrix(0, 生育保险.计生费用总额) = "费用总额"
                .TextMatrix(0, 生育保险.计生统筹支付) = "统筹支付"
                .TextMatrix(0, 生育保险.清算流水号) = "清算流水号"
                .TextMatrix(0, 生育保险.处理情况) = "处理情况"
            End If
            .ColWidth(生育保险.ID) = 0
            .ColWidth(生育保险.期号) = 800
            .ColWidth(生育保险.保险类别代码) = 0
            .ColWidth(生育保险.保险类别) = 1200
            .ColWidth(生育保险.操作员) = 1000
            .ColWidth(生育保险.日期) = 1000
            .ColWidth(生育保险.分娩包干人次) = 1400
            .ColWidth(生育保险.分娩包干费用总额) = 1000
            .ColWidth(生育保险.分娩包干统筹支付) = 1000
            .ColWidth(生育保险.分娩非包干人次) = 1600
            .ColWidth(生育保险.分娩非包干费用总额) = 1000
            .ColWidth(生育保险.分娩非包干统筹支付) = 1000
            .ColWidth(生育保险.计生人次) = 1000
            .ColWidth(生育保险.计生费用总额) = 1000
            .ColWidth(生育保险.计生统筹支付) = 1000
            .ColWidth(生育保险.清算流水号) = 2000
            .ColWidth(生育保险.处理情况) = 2500
        Else
            If blnInit Then
                .Clear
                .Rows = 2: .Cols = 工伤保险.列数
                
                .TextMatrix(0, 工伤保险.ID) = "ID"
                .TextMatrix(0, 工伤保险.期号) = "期号"
                .TextMatrix(0, 工伤保险.保险类别代码) = "保险类别代码"
                .TextMatrix(0, 工伤保险.保险类别) = "保险类别"
                .TextMatrix(0, 工伤保险.操作员) = "操作员"
                .TextMatrix(0, 工伤保险.日期) = "日期"
                .TextMatrix(0, 工伤保险.门诊人次) = "门诊人次"
                .TextMatrix(0, 工伤保险.门诊统筹支付) = "统筹支付"
                .TextMatrix(0, 工伤保险.住院人次) = "住院人次"
                .TextMatrix(0, 工伤保险.住院统筹支付) = "统筹支付"
                .TextMatrix(0, 工伤保险.清算流水号) = "清算流水号"
                .TextMatrix(0, 工伤保险.处理情况) = "处理情况"
            End If
            .ColWidth(工伤保险.ID) = 0
            .ColWidth(工伤保险.期号) = 800
            .ColWidth(工伤保险.保险类别代码) = 0
            .ColWidth(工伤保险.保险类别) = 0
            .ColWidth(工伤保险.操作员) = 1000
            .ColWidth(工伤保险.日期) = 1000
            .ColWidth(工伤保险.门诊人次) = 1400
            .ColWidth(工伤保险.门诊统筹支付) = 1000
            .ColWidth(工伤保险.住院人次) = 1000
            .ColWidth(工伤保险.住院统筹支付) = 1000
            .ColWidth(工伤保险.清算流水号) = 2000
            .ColWidth(工伤保险.处理情况) = 2500
        End If
    End With
End Sub

Private Sub tbrTool_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
    Case "Add1"
        Call mnuEditAdd1_Click
    Case "Add2"
        Call mnuEditAdd2_Click
    Case "Add3"
        Call mnuEditAdd3_Click
    End Select
End Sub
