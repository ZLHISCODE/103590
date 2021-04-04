VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm门诊慢性疾病用药限量_贵阳 
   Caption         =   "门诊慢性疾病用药限量"
   ClientHeight    =   7470
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11145
   Icon            =   "frm门诊慢性疾病用药限量_贵阳.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   7110
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   635
      SimpleText      =   $"frm门诊慢性疾病用药限量_贵阳.frx":0E42
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm门诊慢性疾病用药限量_贵阳.frx":0E89
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14579
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
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   11145
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   645
      Width1          =   615
      FixedBackground1=   0   'False
      Key1            =   "Comm"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   11025
         _ExtentX        =   19447
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
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Object.ToolTipText     =   "打印预览"
               Object.Tag             =   "预览"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "增加"
               Key             =   "New"
               Object.ToolTipText     =   "增加"
               Object.Tag             =   "增加"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "Modify"
               Object.ToolTipText     =   "修改"
               Object.Tag             =   "修改"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "Delete"
               Object.ToolTipText     =   "删除"
               Object.Tag             =   "删除"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "split2"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查找"
               Key             =   "Find"
               Object.ToolTipText     =   "查找"
               Object.Tag             =   "查找"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "split3"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.ToolTipText     =   "帮助主题"
               Object.Tag             =   "帮助"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   4485
      Top             =   390
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
            Picture         =   "frm门诊慢性疾病用药限量_贵阳.frx":171D
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm门诊慢性疾病用药限量_贵阳.frx":1937
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm门诊慢性疾病用药限量_贵阳.frx":1B51
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm门诊慢性疾病用药限量_贵阳.frx":1D6B
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm门诊慢性疾病用药限量_贵阳.frx":1F85
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm门诊慢性疾病用药限量_贵阳.frx":219F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm门诊慢性疾病用药限量_贵阳.frx":23B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm门诊慢性疾病用药限量_贵阳.frx":25D3
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm门诊慢性疾病用药限量_贵阳.frx":27ED
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   5205
      Top             =   360
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
            Picture         =   "frm门诊慢性疾病用药限量_贵阳.frx":2A07
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm门诊慢性疾病用药限量_贵阳.frx":2C21
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm门诊慢性疾病用药限量_贵阳.frx":2E3B
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm门诊慢性疾病用药限量_贵阳.frx":3055
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm门诊慢性疾病用药限量_贵阳.frx":326F
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm门诊慢性疾病用药限量_贵阳.frx":3489
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm门诊慢性疾病用药限量_贵阳.frx":36A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm门诊慢性疾病用药限量_贵阳.frx":38BD
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm门诊慢性疾病用药限量_贵阳.frx":3AD7
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshBill 
      Height          =   3435
      Left            =   90
      TabIndex        =   3
      Top             =   1020
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   6059
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
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
      Begin VB.Menu mnuFileSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "增加(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "修改(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除(&D)"
         Shortcut        =   {DEL}
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
         Begin VB.Menu mnuViewToolspilt1 
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
      Begin VB.Menu mnuViewSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "查找(&F)"
      End
      Begin VB.Menu mnuViewSplit3 
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
         Caption         =   "Web上的中联"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu mnuHelpSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
   Begin VB.Menu mnuShort 
      Caption         =   "快捷菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuShortMenu 
         Caption         =   "增加(&A)"
         Index           =   0
      End
      Begin VB.Menu mnuShortMenu 
         Caption         =   "修改(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuShortMenu 
         Caption         =   "删除(&D)"
         Index           =   2
      End
      Begin VB.Menu mnuShortSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "大图标(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "小图标(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "列表(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "详细资料(&D)"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frm门诊慢性疾病用药限量_贵阳"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mint险类 As Integer
Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    
    sngTop = IIf(cbrThis.Visible, cbrThis.Top + cbrThis.Height, 0)
    sngBottom = ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    
    mshBill.Top = sngTop
    mshBill.Height = IIf(sngBottom - mshBill.Top > 0, sngBottom - mshBill.Top, 0)
    mshBill.Left = ScaleLeft: mshBill.Width = Me.ScaleWidth
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub mnuEditAdd_Click()
    frm药品限量_贵阳.ShowME mint险类, "增加", ""
End Sub

Private Sub mnuEditDelete_Click()
    On Error GoTo errHand
    With mshBill
        If .Rows = 1 Or .Row = 0 Then Exit Sub
        If MsgBox("要删除“[" & .TextMatrix(.Row, 1) & "]" & .TextMatrix(.Row, 2) & "”的用药限量记录吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
        gstrSQL = "Zl_用药限量目录_贵阳_Delete(" & mint险类 & ",'" & .TextMatrix(.Row, 0) & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Call mnuViewRefresh_Click
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuEditModify_Click()
  If mshBill.Rows = 1 Or mshBill.Row = 0 Then Exit Sub
   frm药品限量_贵阳.ShowME mint险类, "修改", mshBill.TextMatrix(mshBill.Row, 0)
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePreview_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnuFilePrintSet_Click()
    zlPrintSet
End Sub

Private Sub subPrint(bytStyle As Byte)
'功能：输入出列表
'参数：bytStyle=1-打印,2-预览,3-输出到Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    
    If gstrUserName = "" Then GetUserInfo
    intRow = mshBill.Row
    
    '表头
    objOut.Title.Text = "药品限量表"
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表项
    objRow.Add "医保类别：" & "贵州省医保"
    objOut.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate, "yyyy年MM月DD日")
    objOut.BelowAppRows.Add objRow
    
    '表体
    Set objOut.Body = mshBill
    
    '输出
    mshBill.Redraw = False
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    mshBill.Redraw = True
    
    mshBill.Row = intRow
    mshBill.COL = 0: mshBill.ColSel = mshBill.Cols - 1
End Sub


Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
   ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuViewFind_Click()
    frm药品限量查找_贵阳.ShowME mint险类
End Sub

Public Sub mnuViewRefresh_Click()
    Call FillList
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim lngCount As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For lngCount = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(lngCount).Caption = IIf(mnuViewToolText.Checked = True, tbrThis.Buttons(lngCount).Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    cbrThis.Refresh
    Call Form_Resize
End Sub


Private Sub MshBill_DblClick()
    Call mnuEditModify_Click
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Preview"
            mnuFilePreview_Click
        Case "Print"
            mnuFilePrint_Click
        Case "New"
            mnuEditAdd_Click
        Case "Delete"
            mnuEditDelete_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Find"
            mnuViewFind_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Exit"
            mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub FillList()
    On Error GoTo errHand
    Dim rsTemp As New ADODB.Recordset
    Me.MousePointer = vbHourglass
    gstrSQL = "Select A.药品ID,A.编码, A.名称, A.规格, A.产地, A.售价单位, trim(to_char(B.数量,'900090.00')) As 限量, " & _
              "      trim(to_char(C.现价,'900090.00000'))  As 售价, trim(to_char(Nvl(B.数量, 0) * Nvl(C.现价, 0),'90009990.00')) As 售价金额,B.备注 " & _
              "From 药品目录 A, 用药限量目录_贵阳 B, 收费价目 C " & _
              "Where A.药品id = B.药品id And B.药品id = C.收费细目ID And B.险类=[1]" & _
              " And (C.终止日期 Is Null Or C.终止日期 = To_Date('3000-01-01', 'yyyy-mm-dd')) Order By A.名称 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint险类)
    Set mshBill.DataSource = rsTemp
    Call CenterTableCaption(mshBill)
    mshBill.ColWidth(0) = 0
    Call SetMenu
    Me.MousePointer = vbDefault
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub
Public Sub SetMenu()
'功能：根据当前内容设置菜单的可用性
    stbThis.Panels(2).Text = "共有" & mshBill.Rows - 1 & "条记录"
    
    mnuEditAdd.Enabled = True
    mnuEditModify.Enabled = mshBill.Rows > 1 And mshBill.Row <> 0
    mnuEditDelete.Enabled = mshBill.Rows > 1 And mshBill.Row <> 0
    tbrThis.Buttons("New").Enabled = mnuEditAdd.Enabled
    tbrThis.Buttons("Modify").Enabled = mnuEditModify.Enabled
    tbrThis.Buttons("Delete").Enabled = mnuEditDelete.Enabled
End Sub
Public Sub ShowForm(frmParent As Object, intinsure As Integer)
'功能：装入医保类别
'说明：使用本功能的主要原因是在出错退出时窗体不会闪
    '现在才能开始使用控件
    mint险类 = intinsure
    If frm门诊慢性疾病用药限量_贵阳.Visible = True Then
        Call mnuViewRefresh_Click
        Call RestoreWinState(Me, App.ProductName)
        frm门诊慢性疾病用药限量_贵阳.Show
        frm门诊慢性疾病用药限量_贵阳.WindowState = 0
        Exit Sub
    End If
    Call mnuViewRefresh_Click
    If frmParent Is Nothing Then
        frm门诊慢性疾病用药限量_贵阳.Show
    Else
        frm门诊慢性疾病用药限量_贵阳.Show , frmParent
    End If
    Call RestoreWinState(Me, App.ProductName)
End Sub


