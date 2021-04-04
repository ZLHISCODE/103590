VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm昭通查询窗体 
   Caption         =   "昭通查询窗体"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11655
   Icon            =   "Frm查询窗体.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgColor 
      Left            =   10080
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm查询窗体.frx":06EA
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm查询窗体.frx":0904
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm查询窗体.frx":0B1E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm查询窗体.frx":0D38
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm查询窗体.frx":0F52
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm查询窗体.frx":164C
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm查询窗体.frx":1866
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm查询窗体.frx":1A80
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   10800
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm查询窗体.frx":1C9A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm查询窗体.frx":1EB4
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm查询窗体.frx":20CE
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm查询窗体.frx":22E8
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm查询窗体.frx":2502
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm查询窗体.frx":2BFC
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm查询窗体.frx":2E16
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm查询窗体.frx":3030
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6390
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "Frm查询窗体.frx":324A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15505
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh明细_S 
      Height          =   5655
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   9975
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
      Appearance      =   0
      MouseIcon       =   "Frm查询窗体.frx":3ADC
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   1244
      BandCount       =   1
      ForeColor       =   -2147483635
      FixedOrder      =   -1  'True
      _CBWidth        =   11655
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      BandForeColor1  =   -2147483635
      Child1          =   "tbrThis"
      MinWidth1       =   4995
      MinHeight1      =   645
      Width1          =   810
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Begin VB.TextBox Txt流水号 
         Height          =   375
         Left            =   8880
         TabIndex        =   4
         Text            =   "输入后回车"
         Top             =   120
         Width           =   2655
      End
      Begin VB.TextBox Txt业务类型 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   375
         Left            =   7080
         TabIndex        =   3
         Text            =   "业务类型"
         Top             =   120
         Width           =   1695
      End
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   1138
         ButtonWidth     =   820
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Description     =   "预览"
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
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      Begin VB.Menu mnuFileSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCard 
         Caption         =   "卡片打印(&A)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSplit2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuBusinessed 
         Caption         =   "交易天数"
      End
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewStatus 
            Caption         =   "状态栏(&S)"
            Checked         =   -1  'True
         End
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
Attribute VB_Name = "frm昭通查询窗体"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mtxtCaption As String, mBusinessedDay As Integer
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Sub cbr_HeightChanged(ByVal NewHeight As Single)
    Call ResizeForm
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If InStr("1234567890" & Chr(0) & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error Resume Next
    cbr.Height = 360
    mBusinessedDay = Val(GetSetting(appName:="ZLSOFT", Section:="私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, Key:="过虑天数", Default:=30))
    Call InitTable
    RestoreWinState Me, App.ProductName
End Sub
Private Sub InitTable()
    Select Case mtxtCaption
    Case "门诊流水号"
        Txt业务类型.Text = "病人id"
        With msh明细_S
            .Clear
            .Rows = 2
            .Cols = 8
            .TextMatrix(0, 0) = "门诊流水号"
            .TextMatrix(0, 1) = "明细号"
            .TextMatrix(0, 2) = "代码"
            .TextMatrix(0, 3) = "名称"
            .TextMatrix(0, 4) = "单位"
            .TextMatrix(0, 5) = "数量"
            .TextMatrix(0, 6) = "单价"
            .TextMatrix(0, 7) = "金额"
            .ColWidth(0) = 1800
            .ColWidth(1) = 600
            .ColWidth(2) = 900
            .ColWidth(3) = 2200
            .ColWidth(4) = 400
            .ColWidth(5) = 300
            .ColWidth(6) = 600
            .ColWidth(7) = 600
        End With
    Case "住院流水号"
        Txt业务类型.Text = "病人id"
        With msh明细_S
            .Clear
            .Rows = 2
            .Cols = 12
        End With
    End Select
    Txt流水号.SetFocus
End Sub
Private Sub Form_Resize()
    Call ResizeForm
End Sub

Private Sub ResizeForm()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    With msh明细_S
        .Top = IIf(cbr.Visible, cbr.Height, 0)
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With

    tbrThis.Width = Me.ScaleWidth - Txt业务类型.Width - Txt流水号.Width - Txt业务类型.Width \ 3
    Txt业务类型.Left = tbrThis.Width
    Txt流水号.Left = Txt业务类型.Left + Txt业务类型.Width
End Sub
Private Sub Form_Unload(Cancel As Integer)
    mtxtCaption = ""
    mBusinessedDay = 0
    SaveWinState Me, App.ProductName
End Sub

Private Sub mnuBusinessed_Click()
Dim Businessed As String
    Businessed = Trim(InputBox("请输入过虑天数", gstrSysName, mBusinessedDay))
    If Businessed <> "" Then
        mBusinessedDay = Businessed
    End If
    Call SaveSetting(appName:="ZLSOFT", Section:="私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, Key:="过虑天数", setting:=mBusinessedDay)
End Sub

Private Sub mnuHelpTitle_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub
Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFileQuit_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnuFilePreview_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
End Sub

Private Sub mnuFileQuit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub
Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Not stbThis.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    cbr.Visible = Not cbr.Visible
    Form_Resize
End Sub

Private Sub tbrThis_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
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
    intRow = msh明细_S.Row
    
    '表头
    objOut.Title.Text = Me.Caption
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表项
    'objRow.Add "医保类别：" & cmb险类.Text
    'objOut.UnderAppRows.Add objRow
    
    'Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate, "yyyy年MM月DD日")
    objOut.BelowAppRows.Add objRow
    
    '表体
    Set objOut.Body = msh明细_S
    
    '输出
    msh明细_S.Redraw = False
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    msh明细_S.Redraw = True
    
    msh明细_S.Row = intRow
    msh明细_S.Col = 0: msh明细_S.ColSel = msh明细_S.Cols - 1
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage Me.hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo Me.hwnd
End Sub
Public Sub ShowForm(frmCaption As String, txtCaption As String)
mtxtCaption = txtCaption
Me.Caption = frmCaption
Me.Show 1
End Sub

Private Sub Txt流水号_GotFocus()
Txt流水号.Text = ""
End Sub

Private Sub Txt流水号_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Dim rsTemp As ADODB.Recordset, vRect As RECT, blncancle As Boolean
        If Trim(Txt流水号) = "" Then Exit Sub
        vRect = GetControlRect(Txt业务类型.hwnd)
        Select Case mtxtCaption
        Case "门诊流水号"
            '当出现多条记录时出现弹出窗,门诊取结算交易号
            gstrSQL = "select C.结帐id AS ID ,A.支付顺序号,D.姓名,A.发生费用金额,A.个人帐户支付,to_char(C.收款时间,'yyyy-mm-dd hh24:mi:ss') as 收费时间 from 保险结算记录 A ,病人信息 D," & _
                            "(select distinct 病人id,收款时间,结帐id from 病人预交记录 where 记录性质=3 and 记录状态=1 and 病人id=" & Txt流水号.Text & ") C " & _
                    "where C.结帐id=A.记录id and a.性质=1 and a.险类=" & TYPE_昭通 & " and d.病人id=a.病人id and C.收款时间>=to_date('" & Format(zlDatabase.Currentdate - mBusinessedDay, "yyyy-mm-dd HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')"
            Set rsTemp = zlDatabase.ShowSelect(Me, gstrSQL, 0, , , , , , , True, vRect.Left - Txt流水号.Width, vRect.Top, Txt流水号.Height, blncancle, , True)
            If (Not rsTemp Is Nothing) And Not blncancle Then
                Call 门诊购药明细查询(rsTemp!支付顺序号, rsTemp!姓名, rsTemp!收费时间)
            End If
        Case "住院流水号"
            '住院取住院登记号
            gstrSQL = "select a.病人id as id,B.姓名,A.顺序号,to_char(A.就诊时间,'yyyy-mm-dd hh24:mi:ss') as 登记时间" & _
                      "  from 保险帐户 A,病人信息 B where  A.险类=" & TYPE_昭通 & " and A.病人id=" & Txt流水号.Text & " and A.病人id=B.病人id  and nvl(b.住院次数,0)>0 and A.就诊时间>=to_date('" & Format(zlDatabase.Currentdate - mBusinessedDay, "yyyy-mm-dd HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')"
            Set rsTemp = zlDatabase.ShowSelect(Me, gstrSQL, 0, , , , , , , True, vRect.Left, vRect.Top, Txt流水号.Height, blncancle, True, True)
            If (Not rsTemp Is Nothing) And Not blncancle Then
                Call 住院情况查询(rsTemp!顺序号, rsTemp!姓名, rsTemp!登记时间)
            End If
        End Select
    End If
End Sub
Private Function GetControlRect(ByVal lngHwnd As Long) As RECT
'功能：获取指定控件在屏幕中的位置(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function
Private Sub 门诊购药明细查询(ByVal str支付顺序号 As String, ByVal str姓名 As String, ByVal str收费时间 As String)
    '门诊购药明细查询
Dim lngLoop As Long, strTemp As String
    
On Error GoTo errHandle
    
    Call InitTable
    If Not frmConn昭通.Execute("I290", 1, str支付顺序号, "正在获取门诊购药明细数据......") Then Exit Sub
    Call ShowWindow(frmConn昭通.hwnd, 9)
    DoEvents
    For lngLoop = 1 To frmConn昭通.mlngRows
    DoEvents
        '发出交易
        If frmConn昭通.Query(lngLoop - 1, 1, "正在查询数据(" & lngLoop & "/" & frmConn昭通.mlngRows & ")......") = False Then Exit Sub
        strTemp = frmConn昭通.strReturnInfo
        '显示数据
        With msh明细_S
            If lngLoop > 1 Then .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = str支付顺序号
            .TextMatrix(.Rows - 1, 1) = Split(strTemp, vbTab)(0)
            .TextMatrix(.Rows - 1, 2) = Split(strTemp, vbTab)(1)
            .TextMatrix(.Rows - 1, 3) = Split(strTemp, vbTab)(2)
            .TextMatrix(.Rows - 1, 4) = Split(strTemp, vbTab)(3)
            .TextMatrix(.Rows - 1, 5) = Split(strTemp, vbTab)(4)
            .TextMatrix(.Rows - 1, 6) = Split(strTemp, vbTab)(5)
            .TextMatrix(.Rows - 1, 7) = Split(strTemp, vbTab)(6)
        End With
    Next
    Call ShowWindow(frmConn昭通.hwnd, 0)
    With msh明细_S
        .Rows = .Rows + 2
        .TextMatrix(.Rows - 1, 0) = "姓名"
        .TextMatrix(.Rows - 1, 1) = str姓名
        .TextMatrix(.Rows - 1, 2) = "HIS收费时间"
        .TextMatrix(.Rows - 1, 3) = str收费时间
    End With
    Exit Sub
    
errHandle:
    If MsgBox("查询数据时发生错误：" & vbCrLf & Err.Description & vbCrLf & "是否重试？", vbInformation + vbRetryCancel, "错误") = vbRetry Then
        Err.Clear
        Resume
    End If
    Call ShowWindow(frmConn昭通.hwnd, 0)
End Sub
Private Sub 住院情况查询(ByVal str顺序号 As String, ByVal str姓名 As String, ByVal str登记时间 As String)
Dim lngLoop As Long, lng住院记录 As Long, lng住院明细 As Long, strTemp As String
    
On Error GoTo errHandle
    
    Call InitTable
    '初始化住院记录区
    msh明细_S.TextMatrix(1, 6) = "住院登记记录"
    
    If Not frmConn昭通.Execute("I360", 0, str顺序号, "正在获取住院情况......") Then Exit Sub
    Call ShowWindow(frmConn昭通.hwnd, 9)
    DoEvents
    For lngLoop = 1 To frmConn昭通.mlngRows
        DoEvents
        '发出交易(住院记录)
        If frmConn昭通.Query(lngLoop - 1, 1, "正在查询数据(" & lngLoop & "/" & frmConn昭通.mlngRows & ")......") = False Then Exit Sub
        strTemp = frmConn昭通.strReturnInfo
        '显示数据
        With msh明细_S
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = "住院序号"
            .TextMatrix(.Rows - 1, 1) = "状态"
            .TextMatrix(.Rows - 1, 2) = "保险证号"
            .TextMatrix(.Rows - 1, 3) = "姓名"
            .TextMatrix(.Rows - 1, 4) = "类别"
            .TextMatrix(.Rows - 1, 5) = "科别"
            .TextMatrix(.Rows - 1, 6) = "床号"
            .TextMatrix(.Rows - 1, 7) = "入院日期"
            .TextMatrix(.Rows - 1, 8) = "出院日期"
            .TextMatrix(.Rows - 1, 9) = "住院次数"
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = Split(strTemp, vbTab)(0)
            .TextMatrix(.Rows - 1, 1) = Split(strTemp, vbTab)(1)
            .TextMatrix(.Rows - 1, 2) = Split(strTemp, vbTab)(2)
            .TextMatrix(.Rows - 1, 3) = Split(strTemp, vbTab)(3)
            .TextMatrix(.Rows - 1, 4) = Split(strTemp, vbTab)(4)
            .TextMatrix(.Rows - 1, 5) = Split(strTemp, vbTab)(5)
            .TextMatrix(.Rows - 1, 6) = Split(strTemp, vbTab)(6)
            .TextMatrix(.Rows - 1, 7) = Split(strTemp, vbTab)(7)
            .TextMatrix(.Rows - 1, 8) = Split(strTemp, vbTab)(8)
            .TextMatrix(.Rows - 1, 9) = Split(strTemp, vbTab)(9)
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = "押金"
            .TextMatrix(.Rows - 1, 1) = "起付金"
            .TextMatrix(.Rows - 1, 2) = "限额"
            .TextMatrix(.Rows - 1, 3) = "医疗费用"
            .TextMatrix(.Rows - 1, 4) = "统筹负担"
            .TextMatrix(.Rows - 1, 5) = "个人负担"
            .TextMatrix(.Rows - 1, 6) = "账户支付"
            .TextMatrix(.Rows - 1, 7) = "进入大病"
            .TextMatrix(.Rows - 1, 8) = "结算日期"
            .TextMatrix(.Rows - 1, 9) = "结算金额"
            .TextMatrix(.Rows - 1, 10) = "操作"
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = Split(strTemp, vbTab)(10)
            .TextMatrix(.Rows - 1, 1) = Split(strTemp, vbTab)(11)
            .TextMatrix(.Rows - 1, 2) = Split(strTemp, vbTab)(12)
            .TextMatrix(.Rows - 1, 3) = Split(strTemp, vbTab)(13)
            .TextMatrix(.Rows - 1, 4) = Split(strTemp, vbTab)(14)
            .TextMatrix(.Rows - 1, 5) = Split(strTemp, vbTab)(15)
            .TextMatrix(.Rows - 1, 6) = Split(strTemp, vbTab)(16)
            .TextMatrix(.Rows - 1, 7) = Split(strTemp, vbTab)(17)
            .TextMatrix(.Rows - 1, 8) = Split(strTemp, vbTab)(18)
            .TextMatrix(.Rows - 1, 9) = Split(strTemp, vbTab)(19)
            .TextMatrix(.Rows - 1, 10) = Split(strTemp, vbTab)(20)
        End With
    Next
    Call ShowWindow(frmConn昭通.hwnd, 0)
    
    
    '初始化住院费用明细区
    With msh明细_S
        .Rows = .Rows + 2
        .TextMatrix(.Rows - 1, 6) = "住院明细记录"
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = "明细序号"
        .TextMatrix(.Rows - 1, 1) = "提交批号"
        .TextMatrix(.Rows - 1, 2) = "日期"
        .TextMatrix(.Rows - 1, 3) = "代码"
        .TextMatrix(.Rows - 1, 4) = "名称"
        .TextMatrix(.Rows - 1, 5) = "状态"
        .TextMatrix(.Rows - 1, 6) = "单位"
        .TextMatrix(.Rows - 1, 7) = "类别"
        .TextMatrix(.Rows - 1, 8) = "数量"
        .TextMatrix(.Rows - 1, 9) = "单价"
        .TextMatrix(.Rows - 1, 10) = "金额"
        .TextMatrix(.Rows - 1, 11) = "特殊标志"
    End With
    strTemp = str顺序号 & vbTab & "0" & vbTab & "0" & vbTab & " " & vbTab & " " & vbTab & " "
    If Not frmConn昭通.Execute("I365", 0, strTemp, "正在获取住院明细记录......") Then Exit Sub
    Call ShowWindow(frmConn昭通.hwnd, 9)
    DoEvents
    
    For lngLoop = 1 To frmConn昭通.mlngRows
        DoEvents
        '发出交易(住院费用明细)
        If frmConn昭通.Query(lngLoop - 1, 1, "正在查询数据(" & lngLoop & "/" & frmConn昭通.mlngRows & ")......") = False Then Exit Sub
        strTemp = frmConn昭通.strReturnInfo
        With msh明细_S
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = Split(strTemp, vbTab)(0)
            .TextMatrix(.Rows - 1, 1) = Split(strTemp, vbTab)(1)
            .TextMatrix(.Rows - 1, 2) = Split(strTemp, vbTab)(2)
            .TextMatrix(.Rows - 1, 3) = Split(strTemp, vbTab)(3)
            .TextMatrix(.Rows - 1, 4) = Split(strTemp, vbTab)(4)
            .TextMatrix(.Rows - 1, 5) = Split(strTemp, vbTab)(5)
            .TextMatrix(.Rows - 1, 6) = Split(strTemp, vbTab)(6)
            .TextMatrix(.Rows - 1, 7) = Split(strTemp, vbTab)(7)
            .TextMatrix(.Rows - 1, 8) = Split(strTemp, vbTab)(8)
            .TextMatrix(.Rows - 1, 9) = Split(strTemp, vbTab)(9)
            .TextMatrix(.Rows - 1, 10) = Split(strTemp, vbTab)(10)
            .TextMatrix(.Rows - 1, 11) = Split(strTemp, vbTab)(11)
        End With
    Next
    lng住院明细 = lngLoop - 1
    Call ShowWindow(frmConn昭通.hwnd, 0)

    '初始化住院费用汇总区
    With msh明细_S
        .Rows = .Rows + 2
        .TextMatrix(.Rows - 1, 0) = "费用汇总"
        .TextMatrix(.Rows - 1, 2) = "基本统筹:"
    End With
    strTemp = str顺序号
    If Not frmConn昭通.Execute("I361", 5, strTemp, "正在获取住院明细记录......") Then Exit Sub
    Call ShowWindow(frmConn昭通.hwnd, 9)
    DoEvents
    For lngLoop = 1 To frmConn昭通.mlngRows
        DoEvents
        '发出交易(住院费用明细)
        If frmConn昭通.Query(lngLoop - 1, 1, "正在查询数据(" & lngLoop & "/" & frmConn昭通.mlngRows & ")......") = False Then Exit Sub
        strTemp = frmConn昭通.strReturnInfo
        With msh明细_S
            .TextMatrix(.Rows - 1, 1) = Split(strTemp, vbTab)(0)
            .TextMatrix(.Rows - 1, 3) = Split(strTemp, vbTab)(1)
        End With
    Next
    With msh明细_S
        .Rows = .Rows + 2
        .TextMatrix(.Rows - 1, 0) = "姓名"
        .TextMatrix(.Rows - 1, 1) = str姓名
        .TextMatrix(.Rows - 1, 8) = "HIS登记时间"
        .TextMatrix(.Rows - 1, 10) = str登记时间
    End With
    Call ShowWindow(frmConn昭通.hwnd, 0)
    Exit Sub
    
errHandle:
    If MsgBox("查询数据时发生错误：" & vbCrLf & Err.Description & vbCrLf & "是否重试？", vbInformation + vbRetryCancel, "错误") = vbRetry Then
        Err.Clear
        Resume
    End If
    Call ShowWindow(frmConn昭通.hwnd, 0)
End Sub
