VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm大处方审查 
   BackColor       =   &H8000000A&
   Caption         =   "大处方审查"
   ClientHeight    =   6690
   ClientLeft      =   165
   ClientTop       =   3750
   ClientWidth     =   11760
   Icon            =   "frm大处方审查.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6690
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSum 
      Height          =   2325
      Left            =   540
      TabIndex        =   5
      Top             =   1200
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   4101
      _Version        =   393216
      FixedCols       =   0
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      GridColor       =   8421504
      GridColorFixed  =   8421504
      GridColorUnpopulated=   8421504
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox picV 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1995
      Left            =   3180
      MousePointer    =   9  'Size W E
      ScaleHeight     =   1995
      ScaleWidth      =   45
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2550
      Width           =   45
   End
   Begin MSComctlLib.ImageList ilsColor 
      Left            =   4800
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm大处方审查.frx":030A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm大处方审查.frx":0526
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm大处方审查.frx":0742
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm大处方审查.frx":095C
            Key             =   "Font"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm大处方审查.frx":0B78
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm大处方审查.frx":0D94
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMono 
      Left            =   4080
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm大处方审查.frx":0FB0
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm大处方审查.frx":11CC
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm大处方审查.frx":13E8
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm大处方审查.frx":1602
            Key             =   "Font"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm大处方审查.frx":181E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm大处方审查.frx":1A3A
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   11760
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   720
      Width1          =   5370
      Key1            =   "one"
      NewRow1         =   0   'False
      Caption2        =   "科室"
      Child2          =   "cmbDept"
      MinHeight2      =   300
      Width2          =   765
      Key2            =   "two"
      NewRow2         =   0   'False
      Begin VB.ComboBox cmbDept 
         Height          =   300
         Left            =   5985
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   5685
      End
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
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
               Caption         =   "重置"
               Key             =   "Open"
               Object.ToolTipText     =   "条件重置"
               Object.Tag             =   "重置"
               ImageKey        =   "Open"
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
      Top             =   6330
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   635
      SimpleText      =   $"frm大处方审查.frx":1C56
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm大处方审查.frx":1C9D
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15663
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   1215
      Left            =   6510
      TabIndex        =   6
      Top             =   2490
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   2143
      _Version        =   393216
      FixedCols       =   0
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      GridColor       =   8421504
      GridColorFixed  =   8421504
      GridColorUnpopulated=   8421504
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblDetail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "处方明细"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6120
      TabIndex        =   7
      Top             =   1380
      Width           =   3015
   End
   Begin VB.Label lblSum 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "处方列表"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   930
      Width           =   4440
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
      Begin VB.Menu mnuViewSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOpen 
         Caption         =   "条件重置(&J)"
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
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "中联论坛(&F)"
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
Attribute VB_Name = "frm大处方审查"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnLoad As Boolean
Dim mdatBegin As Date, mdatEnd As Date            '查询的时间范围

'公共模块参数
Dim mdblMax As Double                               '审查标准值
Dim mintUnit As Integer                             '药品单位：0－售价单位；1－药房单位

Dim mstrType As String                            '处方类型(多选)
Dim msngStartX As Single, msngStartY As Single    '移动前鼠标的位置
Dim mlngID As Long          '前一个部门的ID
Dim mstrNo As String        '前一张处方的NO
Dim mstr类型 As String      '前一张处方的类型
Dim mlng病人ID As Long      '前一张处方的病人ID
Dim mlngRow As Long         '前一次作画时的行坐标
Private mlngMode As Long
Private mstrPrivs As String             '当前用户具有的当前模块的功能


Private Sub cbrThis_HeightChanged(ByVal NewHeight As Single)
    Call Form_Resize
End Sub

Private Sub cmbDept_Click()
    If mblnLoad = False Then
        Call FillSum
    End If
End Sub

Private Sub Form_Activate()
    If mblnLoad = True Then
        FillDept
    End If
    mblnLoad = False
End Sub

Private Sub Form_Load()
    mblnLoad = True
    mlngMode = glngModul
    mstrPrivs = gstrprivs
    
    RestoreWinState Me, App.ProductName
    Call zldatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    '得到查询的时间范围
    mdatEnd = CDate(Format(Sys.Currentdate, "yyyy-MM-dd"))
    mdatBegin = DateAdd("d", -10, mdatEnd) + 1
    
    mdblMax = Val(zldatabase.GetPara("审查标准", glngSys, 1347))
    mintUnit = Val(zldatabase.GetPara("药品单位", glngSys, 1347))

    Call InitSum
End Sub

Private Sub InitSum()
'初始化汇总表的样式
    With mshSum
        ClearGrid mshSum, 8
        .TextMatrix(0, 0) = "医生"
        .TextMatrix(0, 1) = "日期"
        .TextMatrix(0, 2) = "类型"
        .TextMatrix(0, 3) = "处方号"
        .TextMatrix(0, 4) = "金额"
        .TextMatrix(0, 5) = "病人姓名"
        .TextMatrix(0, 6) = "门诊/住院号"
        .TextMatrix(0, 7) = "病人ID"
        
        .ColWidth(0) = 690
        .ColWidth(1) = 1020
        .ColWidth(2) = 540
        .ColWidth(3) = 840
        .ColWidth(4) = 945
        .ColWidth(5) = 900
        .ColWidth(6) = 1500
        .ColWidth(7) = 0
        
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 7
        .ColAlignment(5) = 1
        .ColAlignment(6) = 1
        
        .MergeCells = flexMergeRestrictRows
        .MergeCol(0) = True
        .MergeCol(1) = True
    End With
    
    With mshDetail
        ClearGrid mshDetail, 7
        .TextMatrix(0, 0) = "药品名称"
        .TextMatrix(0, 1) = "规格"
        .TextMatrix(0, 2) = "批号 "
        .TextMatrix(0, 3) = "单位"
        .TextMatrix(0, 4) = "数量"
        .TextMatrix(0, 5) = "单价"
        .TextMatrix(0, 6) = "金额"
        
        .ColWidth(0) = 2500
        .ColWidth(1) = 1200
        .ColWidth(2) = 1300
        .ColWidth(3) = 450
        .ColWidth(4) = 825
        .ColWidth(5) = 840
        .ColWidth(6) = 1140
        
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 4
        .ColAlignment(4) = 7
        .ColAlignment(5) = 7
        .ColAlignment(6) = 7
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlngID = -1
    
    zldatabase.SetPara "药品单位", mintUnit, glngSys, 1347

    SaveWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    sngTop = IIf(cbrThis.Visible, cbrThis.Top + cbrThis.Height, 0)
    sngBottom = ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    lblSum.Top = sngTop
    lblSum.Left = 0
    mshSum.Left = lblSum.Left
    mshSum.Width = lblSum.Width
    mshSum.Top = lblSum.Top + lblSum.Height
    If sngBottom - mshSum.Top > 0 Then mshSum.Height = sngBottom - mshSum.Top
    
    picV.Top = lblSum.Top
    picV.Left = lblSum.Left + lblSum.Width
    picV.Height = sngBottom - picV.Top
    
    
    lblDetail.Left = picV.Left + picV.Width
    mshDetail.Left = lblDetail.Left
    lblDetail.Top = lblSum.Top
    mshDetail.Top = mshSum.Top
    
    lblDetail.Width = IIf(ScaleWidth - lblDetail.Left > 0, ScaleWidth - lblDetail.Left, 0)
    mshDetail.Width = lblDetail.Width
    mshDetail.Height = mshSum.Height
    
    Refresh
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    '默认参数：开始时间=登记时间开始，结束时间=登记时间结束，NO=处方NO,病人科室=病人科室ID,病人ID=病人ID
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "病人科室=" & IIf(Val(cmbDept.ItemData(cmbDept.ListIndex)) = 0, "", Val(cmbDept.ItemData(cmbDept.ListIndex))), _
        "开始时间=" & Format(mdatBegin, "yyyy-mm-dd"), _
        "结束时间=" & Format(mdatEnd, "yyyy-mm-dd"), _
        "NO=" & mstrNo, _
        "病人ID=" & IIf(mlng病人ID = 0, "", mlng病人ID))
End Sub
Private Sub mnuViewOpen_Click()
    If frm审查条件.GetCondition(mdatBegin, mdatEnd, mintUnit, mstrType, mstrPrivs, Me) = True Then
        mlngID = -1
        Call FillSum
    End If
End Sub

Private Sub mnuViewRefresh_Click()
    FillDept
End Sub

Private Sub mshSum_EnterCell()
    SetColor False, mlngRow
    mlngRow = mshSum.Row
    SetColor True, mlngRow
    
    If mstrNo = mshSum.TextMatrix(mshSum.Row, 3) And mstr类型 = mshSum.TextMatrix(mshSum.Row, 2) Then Exit Sub
    
    mstrNo = mshSum.TextMatrix(mshSum.Row, 3)
    mstr类型 = mshSum.TextMatrix(mshSum.Row, 2)
    mlng病人ID = Val(mshSum.TextMatrix(mshSum.Row, 7))
    Call FillDetail
End Sub

Private Sub SetColor(ByVal blnChange As Boolean, ByVal lngRow As Long)
'参数:blnChange  为真改变表格的颜色，为假表示还原
    Dim lngTemp As Long
    Dim i As Long

    With mshSum
        If lngRow < 0 Or lngRow > .rows - 1 Then Exit Sub
        .Redraw = False
        lngTemp = .Row
        .Row = lngRow
        If blnChange = True Then
            For i = 2 To .Cols - 1
                .Col = i
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
            Next
        Else
            For i = 2 To .Cols - 1
                .Col = i
                .CellBackColor = &H80000005
                .CellForeColor = &H80000008
            Next
        End If
        .Row = lngTemp
        .Redraw = True
    End With
End Sub

Private Sub mshSum_GotFocus()
    Call MenuSet
End Sub

Private Sub mshSum_LostFocus()
    Call MenuSet
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool, 2
    End If
End Sub

Private Sub picV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        msngStartX = X
    End If
End Sub

Private Sub picV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picV.Left + X - msngStartX
        If sngTemp > lblSum.Left + 600 And ScaleWidth - (sngTemp + picV.Width) > 1600 Then
            picV.Left = sngTemp
            lblSum.Width = picV.Left - lblSum.Left
            Form_Resize
        End If
    End If
End Sub

Private Sub mnufileexit_Click()
    Unload Me
End Sub

Private Sub mnuFilePrintSet_Click()
    zlPrintSet
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFilePreView_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub


Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Open"
            mnuViewOpen_Click
        Case "Quit"
            mnufileexit_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnuFilePreView_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
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
    cbrThis.Bands("one").MinHeight = tbrThis.Height
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
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(hWnd)
End Sub

Private Sub subPrint(bytMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    If mshSum Is ActiveControl Then
        Set objPrint.Body = mshSum
        objPrint.Title.Text = "大处方列表"
        objRow.Add " "
        objRow.Add "查询时间：" & Format(mdatBegin, "yyyy-MM-dd") & " 至 " & Format(mdatEnd, "yyyy-MM-dd")
        objPrint.UnderAppRows.Add objRow
        
        Set objRow = New zlTabAppRow
        objRow.Add "打印人：" & gstrUserName
        objRow.Add "打印时间：" & Format(Sys.Currentdate, "yyyy-MM-dd")
        objPrint.BelowAppRows.Add objRow
    Else
        Set objPrint.Body = mshDetail
        objPrint.Title.Text = "处方明细"
        objRow.Add "处方号：" & mshSum.TextMatrix(mshSum.Row, 3)
        objRow.Add "查询时间：" & Format(mdatBegin, "yyyy-MM-dd") & " 至 " & Format(mdatEnd, "yyyy-MM-dd")
        objPrint.UnderAppRows.Add objRow
        
        Set objRow = New zlTabAppRow
        objRow.Add "打印人：" & gstrUserName
        objRow.Add "打印时间：" & Format(Sys.Currentdate, "yyyy-MM-dd")
        objPrint.BelowAppRows.Add objRow
    End If
    If mshSum Is ActiveControl Then
        mshSum.Redraw = False
        SetColor False, mshSum.Row
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
    If mshSum Is ActiveControl Then
        mshSum.Redraw = True
        SetColor True, mshSum.Row
    End If
End Sub

Private Function FillDept() As Boolean
'功能:装入药品供应商
    
    Dim rstemp As New ADODB.Recordset
    Dim strTemp As String
    Dim LngID As Long
    
    mlngID = -1     '全面刷新时就相当于用户没点过任何节点
    If cmbDept.ListIndex > 0 Then
        LngID = cmbDept.ItemData(cmbDept.ListIndex)
    End If
    
    On Error GoTo errHandle
    rstemp.CursorLocation = adUseClient
    gstrSQL = "select id,名称 from 部门表 A,部门性质说明 B  where (A.站点 = '" & gstrNodeNo & "' Or A.站点 is Null) And (A.撤档时间=to_date('3000-01-01','yyyy-mm-dd') or A.撤档时间 is null) " & _
         " and A.ID=B.部门ID and B.工作性质='临床' and B.服务对象<>0 order by A.编码"
    Call zldatabase.OpenRecordset(rstemp, gstrSQL, Me.Caption)
    
    If rstemp.RecordCount = 0 Then
        MsgBox "“开单科室”的信息不全，无法进行查询。", vbExclamation, gstrSysName
        FillDept = False
        Exit Function
    End If
    
    
    With cmbDept
        .Clear
        .AddItem "所有科室"
        Do Until rstemp.EOF
            .AddItem rstemp("名称")
            .ItemData(.NewIndex) = rstemp("ID")
            If rstemp("ID") = LngID Then
                .ListIndex = .NewIndex
            End If
            rstemp.MoveNext
        Loop
        If .ListIndex < 0 Then .ListIndex = 0
    End With
    
    FillDept = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub FillSum()
'功能:装入各种统计数据
    Dim rstemp As New ADODB.Recordset
    Dim strBegin As String, strEnd As String
    Dim lngRow As Long
    Dim str处方类型 As String
    Dim str处方类型1 As String
    
    On Error GoTo errHandle
    If cmbDept.ListIndex < 0 Then Exit Sub
    If mlngID = cmbDept.ItemData(cmbDept.ListIndex) Then Exit Sub
    mlngID = cmbDept.ItemData(cmbDept.ListIndex)
    '开始查询
    
    strBegin = Format(mdatBegin, "yyyy-MM-dd")
    strEnd = Format(mdatEnd + 1, "yyyy-MM-dd")
    
    If InStr(1, mstrType, "012") > 0 Then
        str处方类型 = ""
    ElseIf InStr(1, mstrType, "01") > 0 Then
        str处方类型 = " And (A.门诊标志=1 Or A.门诊标志=4) "
    ElseIf InStr(1, mstrType, "02") > 0 Then
        str处方类型 = " And (A.门诊标志=1 Or A.门诊标志=4) And A.记帐费用=0 "
        str处方类型1 = " And A.门诊标志=2 And A.记帐费用=1 "
    ElseIf InStr(1, mstrType, "12") > 0 Then
        str处方类型 = " And A.记帐费用=1 "
    ElseIf InStr(1, mstrType, "0") > 0 Then
         str处方类型 = " And (A.门诊标志=1 Or A.门诊标志=4) And A.记帐费用=0 "
    ElseIf InStr(1, mstrType, "1") > 0 Then
         str处方类型 = " And (A.门诊标志=1 Or A.门诊标志=4) And A.记帐费用=1 "
    ElseIf InStr(1, mstrType, "2") > 0 Then
         str处方类型 = " And A.门诊标志=2 And A.记帐费用=1 "
    End If
    
    MousePointer = 11
    
    '再得到完整的SQL语句
    If mlngID = 0 Then
        '所有科室的操作员
        gstrSQL = "select A.开单人,to_char(A.登记时间,'yyyy-mm-dd') as 日期,A.NO,decode(A.记录性质,1,'收费','记帐') as 处方类型,sum(A.实收金额) as 金额," & _
                   " A.姓名,decode(A.门诊标志,1,'(门诊)',4,'(门诊)',decode(A.门诊标志,2,'(住院)','(其他)'))||A.标识号 标识号,A.病人ID " & _
                   " from 门诊费用记录 A,收费项目目录 C, " & _
                   " (Select Distinct 单据,No, 费用id From 药品收发记录 Where 单据 In (8, 9) And Mod(记录状态, 3) = 1 And 填制日期>=[2] And 填制日期<[3]) B " & _
                   " where A.登记时间>=[2] and A.登记时间<[3] " & _
                   "       and (A.记录性质=1 or A.记录性质=2)  and 记录状态=1 and A.开单人 is not null " & _
                   " And A.收费细目id=C.Id And C.类别 In('5','6','7') And A.No = B.No And A.Id = B.费用id" & str处方类型 & _
                   " group by A.开单人,A.登记时间,A.NO ,A.记录性质,A.姓名,Decode(A.门诊标志, 1, '(门诊)',4,'(门诊)', Decode(A.门诊标志, 2, '(住院)', '(其他)')) || A.标识号,A.病人ID " & _
                   " Having Sum(A.实收金额) >= [1] "
                   
    Else
        gstrSQL = "select A.开单人,to_char(A.登记时间,'yyyy-mm-dd') as 日期,A.NO,decode(A.记录性质,1,'收费','记帐') as 处方类型,sum(A.实收金额) as 金额, " & _
                   " 姓名,decode(A.门诊标志,1,'(门诊)',4,'(门诊)',decode(A.门诊标志,2,'(住院)','(其他)'))||A.标识号 标识号,A.病人ID  " & _
                   " from 门诊费用记录 A,收费项目目录 C, " & _
                   " (Select Distinct 单据,No, 费用id From 药品收发记录 Where 单据 In (8, 9) And Mod(记录状态, 3) = 1 And 填制日期>=[2] And 填制日期<[3]) B " & _
                   " where A.登记时间>=[2] and A.登记时间<[3] " & _
                   "       and (A.记录性质=1 or A.记录性质=2)  and A.记录状态=1  and A.开单人 is not null and A.开单部门ID+0=[4] " & _
                   " And A.收费细目id=C.Id And C.类别 In('5','6','7') And A.No = B.No And A.Id = B.费用id" & str处方类型 & _
                   " group by A.开单人,A.登记时间,A.NO ,A.记录性质,A.姓名,Decode(A.门诊标志, 1, '(门诊)',4,'(门诊)', Decode(A.门诊标志, 2, '(住院)', '(其他)')) || A.标识号,A.病人ID  " & _
                   " Having Sum(A.实收金额) >= [1] "
    End If
    
    If str处方类型1 <> "" Then
        If mlngID = 0 Then
            gstrSQL = gstrSQL & _
                   " UNION ALL" & _
                   " select A.开单人,to_char(A.登记时间,'yyyy-mm-dd') as 日期,A.NO,decode(A.记录性质,1,'收费','记帐') as 处方类型,sum(A.实收金额) as 金额," & _
                   " A.姓名,decode(A.门诊标志,1,'(门诊)',4,'(门诊)',decode(A.门诊标志,2,'(住院)','(其他)'))||A.标识号 标识号,A.病人ID " & _
                   " from 门诊费用记录 A,收费项目目录 C, " & _
                   " (Select Distinct 单据,No, 费用id From 药品收发记录 Where 单据 In (8, 9) And Mod(记录状态, 3) = 1 And 填制日期>=[2] And 填制日期<[3]) B " & _
                   " where A.登记时间>=[2] and A.登记时间<[3] " & _
                   "       and (A.记录性质=1 or A.记录性质=2)  and 记录状态=1 and A.开单人 is not null " & _
                   " And A.收费细目id=C.Id And C.类别 In('5','6','7') And A.No = B.No And A.Id = B.费用id" & str处方类型1 & _
                   " group by A.开单人,A.登记时间,A.NO ,A.记录性质,A.姓名,Decode(A.门诊标志, 1, '(门诊)', 4,'(门诊)',Decode(A.门诊标志, 2, '(住院)', '(其他)')) || A.标识号,A.病人ID " & _
                   " Having Sum(A.实收金额) >= [1] "
        Else
            gstrSQL = gstrSQL & _
                    " UNION ALL" & _
                    "select A.开单人,to_char(A.登记时间,'yyyy-mm-dd') as 日期,A.NO,decode(A.记录性质,1,'收费','记帐') as 处方类型,sum(A.实收金额) as 金额, " & _
                    " 姓名,decode(A.门诊标志,1,'(门诊)',4,'(门诊)',decode(A.门诊标志,2,'(住院)','(其他)'))||A.标识号 标识号,A.病人ID  " & _
                    " from 门诊费用记录 A,收费项目目录 C, " & _
                    " (Select Distinct 单据,No, 费用id From 药品收发记录 Where 单据 In (8, 9) And Mod(记录状态, 3) = 1 And 填制日期>=[2] And 填制日期<[3]) B " & _
                    " where A.登记时间>=[2] and A.登记时间<[3] " & _
                    "       and (A.记录性质=1 or A.记录性质=2)  and A.记录状态=1  and A.开单人 is not null and A.开单部门ID+0=[4] " & _
                    " And A.收费细目id=C.Id And C.类别 In('5','6','7') And A.No = B.No And A.Id = B.费用id" & str处方类型1 & _
                    " group by A.开单人,A.登记时间,A.NO ,A.记录性质,A.姓名,Decode(A.门诊标志, 1, '(门诊)',4,'(门诊)', Decode(A.门诊标志, 2, '(住院)', '(其他)')) || A.标识号,A.病人ID  " & _
                    " Having Sum(A.实收金额) >= [1] "
        End If
    End If
    
    If mstrType = "0" Or mstrType = "1" Or mstrType = "01" Then
    ElseIf mstrType = "2" Then
        gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
    Else
        gstrSQL = gstrSQL & " Union All " & Replace(gstrSQL, "门诊费用记录", "住院费用记录")
    End If
    gstrSQL = gstrSQL & " order by 开单人,日期,NO"
    
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mdblMax, CDate(strBegin), CDate(strEnd), mlngID)
    
    mshSum.Redraw = False
    If rstemp.RecordCount = 0 Then
        ClearGrid mshSum
    Else
        mshSum.rows = rstemp.RecordCount + 1
    End If
    lngRow = 1
    With mshSum
        Do Until rstemp.EOF
            .TextMatrix(lngRow, 0) = rstemp("开单人")
            .TextMatrix(lngRow, 1) = IIf(IsNull(rstemp("日期")), "", rstemp("日期"))
            .TextMatrix(lngRow, 2) = IIf(IsNull(rstemp("处方类型")), "", rstemp("处方类型"))
            .TextMatrix(lngRow, 3) = IIf(IsNull(rstemp("NO")), "", rstemp("NO"))
            .TextMatrix(lngRow, 4) = Format(rstemp("金额"), "###########0.00;-###########0.00; ; ")
            .TextMatrix(lngRow, 5) = IIf(IsNull(rstemp("姓名")), "", rstemp("姓名"))
            .TextMatrix(lngRow, 6) = IIf(IsNull(rstemp("标识号")), "", rstemp("标识号"))
            .TextMatrix(lngRow, 7) = IIf(IsNull(rstemp("病人ID")), "", rstemp("病人ID"))
            lngRow = lngRow + 1
            rstemp.MoveNext
        Loop
    End With
    mshSum.Redraw = True
    
    stbThis.Panels(2).Text = "时间范围：" & Format(mdatBegin, "yyyy-MM-dd") & " 至 " & Format(mdatEnd, "yyyy-MM-dd") & _
                            "。 处方数：" & rstemp.RecordCount
    MousePointer = 0
    Call FillDetail
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FillDetail()
'功能:装入明细数据
    Dim rstemp As New ADODB.Recordset
    Dim lngRow As Long
    Dim strNo As String, lng单据 As Long
    Dim int标识 As String
    Dim strUnitQuantity As String
    Dim strFormat As String
    
    On Error GoTo errHandle
    MousePointer = 11
    
    mshDetail.Redraw = False
    '初始化表格
    ClearGrid mshDetail
    
    '得到查询条件
    strNo = mshSum.TextMatrix(mshSum.Row, 3)
    lng单据 = IIf(mshSum.TextMatrix(mshSum.Row, 2) = "收费", 8, 9)
    If InStrB(1, mshSum.TextMatrix(mshSum.Row, 6), "(门诊)") > 0 Then
        int标识 = 1
    Else
        int标识 = 2
    End If
    
    If strNo = "" Then
        mshDetail.Redraw = True
        MousePointer = 0
        
        Call MenuSet
        Exit Sub
    End If
    
    Select Case mintUnit
        Case 1      '药房单位
            If int标识 = 1 Then
                '门诊单位
                strUnitQuantity = ",B.门诊单位 AS 单位,(A.实际数量 / B.门诊包装) AS 实际数量,a.零售价*B.门诊包装 as 零售价"
            Else
                '住院单位
                strUnitQuantity = ",B.住院单位 AS 单位,(A.实际数量 / B.住院包装) AS 实际数量,a.零售价*B.住院包装 as 零售价"
            End If
        Case Else   '售价单位
            strUnitQuantity = ",C.计算单位 AS 单位, a.实际数量,a.成本价,a.零售价"
    End Select
    
    
    '开始查询
    gstrSQL = " SELECT DISTINCT C.名称 通用名称,C.规格,A.批号||DECODE(NVL(A.批次,0),0,'','('||A.批次||')') 批号," & _
           " A.零售金额,A.序号 " & strUnitQuantity & _
           " FROM 药品收发记录 A,药品规格 B,收费项目目录 C " & _
           " WHERE A.药品ID +0 = B.药品ID AND B.药品ID=C.ID " & _
           "       AND A.NO=[1] AND A.单据=[2] AND MOD(A.记录状态,3)=1 " & _
           " ORDER BY A.序号"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, lng单据)
    
    If rstemp.RecordCount > 0 Then
        mshDetail.rows = rstemp.RecordCount + 1
        lngRow = 1
        With mshDetail
            Do Until rstemp.EOF
                .TextMatrix(lngRow, 0) = rstemp("通用名称")
                .TextMatrix(lngRow, 1) = IIf(IsNull(rstemp("规格")), "", rstemp("规格"))
                .TextMatrix(lngRow, 2) = IIf(IsNull(rstemp("批号")), "", rstemp("批号"))
                .TextMatrix(lngRow, 3) = IIf(IsNull(rstemp("单位")), "", rstemp("单位"))
                .TextMatrix(lngRow, 4) = Format(rstemp("实际数量"), "###########0.000;-###########0.000; ; ")
                .TextMatrix(lngRow, 5) = Format(rstemp("零售价"), "###########0.000;-###########0.000; ; ")
                .TextMatrix(lngRow, 6) = Format(rstemp("零售金额"), "###########0.00;-###########0.00; ; ")
                    
                lngRow = lngRow + 1
                rstemp.MoveNext
            Loop
            
        End With
    End If
    mshDetail.Redraw = True
    MousePointer = 0
    
    Call MenuSet
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ClearGrid(objGrid As MSHFlexGrid, Optional lngCols As Long = 0)
'功能：清除表格,并完成部分初始化
    Dim i As Long
    
    With objGrid
        If lngCols > 0 Then
            '如果有列数传进来，那就初始化它
            .Cols = lngCols
            .AllowBigSelection = True
            .FillStyle = flexFillRepeat
            .Col = 0
            .Row = 0
            .ColSel = .Cols - 1
            .RowSel = 0
            .CellAlignment = 4
            .FillStyle = flexFillSingle
            .AllowBigSelection = False
        End If
        
        .rows = 2
        .Row = 1
        .RowData(1) = 0
        For i = 0 To objGrid.Cols - 1
            objGrid.TextMatrix(1, i) = ""
            If objGrid Is mshSum And i > 1 Then
                mlngRow = 1
                .Col = i
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
            End If
        Next
    End With
End Sub

Private Sub MenuSet()
'功能:显示菜单和工具栏的状态(打印)
    Dim blnPrint As Boolean
    
    If ActiveControl Is mshSum Then
        blnPrint = Not (mshSum.rows = 2 And mshSum.TextMatrix(1, 0) = "")
    Else
        blnPrint = Not (mshDetail.rows = 2 And mshDetail.TextMatrix(1, 0) = "")
    End If

    mnuFilePreview.Enabled = blnPrint
    mnuFilePrint.Enabled = blnPrint
    mnuFileExcel.Enabled = blnPrint
    tbrThis.Buttons("Preview").Enabled = blnPrint
    tbrThis.Buttons("Print").Enabled = blnPrint
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

