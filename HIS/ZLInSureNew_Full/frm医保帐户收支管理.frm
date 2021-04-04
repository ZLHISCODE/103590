VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm医保帐户收支管理 
   Caption         =   "帐户收支管理"
   ClientHeight    =   5220
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7470
   Icon            =   "frm医保帐户收支管理.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ImageList ImgColor 
      Left            =   240
      Top             =   90
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
            Picture         =   "frm医保帐户收支管理.frx":06EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户收支管理.frx":0904
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户收支管理.frx":0B1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户收支管理.frx":0E70
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户收支管理.frx":108A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户收支管理.frx":13DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户收支管理.frx":15F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户收支管理.frx":1810
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户收支管理.frx":1A2A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgBlack 
      Left            =   780
      Top             =   90
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
            Picture         =   "frm医保帐户收支管理.frx":1C44
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户收支管理.frx":1E5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户收支管理.frx":2078
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户收支管理.frx":2292
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户收支管理.frx":24AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户收支管理.frx":26C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户收支管理.frx":28E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户收支管理.frx":2AFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户收支管理.frx":2D14
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrTool 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   1270
      BandCount       =   2
      _CBWidth        =   7470
      _CBHeight       =   720
      _Version        =   "6.7.9782"
      Child1          =   "tbrTool"
      MinWidth1       =   3000
      MinHeight1      =   660
      Width1          =   2820
      NewRow1         =   0   'False
      BandForeColor2  =   -2147483646
      BandBackColor2  =   -2147483638
      Caption2        =   "保险类别"
      Child2          =   "cbo险类"
      MinWidth2       =   1800
      MinHeight2      =   300
      Width2          =   1335
      UseCoolbarColors2=   0   'False
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tbrTool 
         Height          =   660
         Left            =   165
         TabIndex        =   3
         Top             =   30
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   1164
         ButtonWidth     =   820
         ButtonHeight    =   1164
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
               Object.Tag             =   "打印"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Printview"
               Object.Tag             =   "预览"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "调整"
               Key             =   "Adjust"
               Object.Tag             =   "调整"
               ImageIndex      =   3
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Single"
                     Text            =   "单个调整"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Batch"
                     Text            =   "批量调整"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "Modify"
               Object.Tag             =   "修改"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "Delete"
               Object.Tag             =   "删除"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查看"
               Key             =   "View"
               Object.Tag             =   "查看"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Find"
               Object.Tag             =   "过滤"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.Tag             =   "帮助"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Object.Tag             =   "退出"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cbo险类 
         Height          =   300
         Left            =   5580
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   210
         Width           =   1800
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf帐户变动记录 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   750
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   7223
      _Version        =   393216
      BackColor       =   16777215
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   4860
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm医保帐户收支管理.frx":2F2E
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8096
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
      Begin VB.Menu mnuFileQuit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditAdjust 
         Caption         =   "帐户余额调整(&A)"
         Begin VB.Menu mnuEditAdjust_Single 
            Caption         =   "单个调整(&S)"
         End
         Begin VB.Menu mnuEditAdjust_Batch 
            Caption         =   "批量调整(&B)"
         End
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "修改变动记录(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除变动记录(&D)"
      End
      Begin VB.Menu mnuEditSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditView 
         Caption         =   "查看(&V)"
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
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "查找(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuView_2 
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
Attribute VB_Name = "frm医保帐户收支管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFind As String                   '查找串，缺省为查找第一个医保中心的病人的帐户变动记录，如果不分中心，则查找所有
Private lngCardRow As Long
Private mint险类 As Integer
Private mblnLoad As Boolean
Private mstrPrivs As String
Private Const glng白色 As Long = &H80000005
Private Const glng黑色 As Long = &H80000008
Private Const glng深灰色 As Long = &HC0C0C0
Private Const glng本色 As Long = &H8000000F
Private Const glng灰色 As Long = &HE0E0E0
Private Const glng红色 As Long = &HC0
Private Const glng深蓝色 As Long = &H8000000D

Private Enum 列Enum
    colID = 0
    col中心 = 1
    col卡号 = 2
    col医保号 = 3
    col病人ID = 4
    col姓名 = 5
    col金额 = 6
    col经办人 = 7
    col时间 = 8
    col性质 = 9
    col说明 = 10
    col列数 = 11
End Enum

Private Sub cbo险类_Click()
    With cbo险类
        If mint险类 = .ItemData(.ListIndex) Then Exit Sub
        mint险类 = .ItemData(.ListIndex)
    End With
    
    '填充数据
    Call FillList
End Sub

Private Sub cbrTool_Resize()
    Call Form_Resize
End Sub

Private Sub InitBill(Optional ByVal bln中心 As Boolean = True)
    '初始化表格
    Dim lngCol As Integer
    
    '设置格式
    With Msf帐户变动记录
        .Clear
        .Rows = 2
        .Cols = col列数
        For lngCol = 0 To .Cols - 1
            .TextMatrix(1, lngCol) = ""
        Next
        
        .TextMatrix(0, colID) = "ID"
        .TextMatrix(0, col中心) = "中心"
        .TextMatrix(0, col卡号) = "卡号"
        .TextMatrix(0, col医保号) = "医保号"
        .TextMatrix(0, col病人ID) = "病人ID"
        .TextMatrix(0, col姓名) = "姓名"
        .TextMatrix(0, col金额) = "金额"
        .TextMatrix(0, col经办人) = "经办人"
        .TextMatrix(0, col时间) = "时间"
        .TextMatrix(0, col性质) = "性质"
        .TextMatrix(0, col说明) = "说明"
        If Not mblnLoad Then
            .ColWidth(colID) = 0
            .ColWidth(col中心) = IIf(bln中心, 1500, 0)
            .ColWidth(col卡号) = 900
            .ColWidth(col医保号) = 900
            .ColWidth(col病人ID) = 1000
            .ColWidth(col姓名) = 800
            .ColWidth(col金额) = 1200
            .ColWidth(col经办人) = 800
            .ColWidth(col时间) = 1800
            .ColWidth(col性质) = 0
            .ColWidth(col说明) = 2000
            'Call RestoreFlexState(Msf帐户变动记录, Me.Caption)
            .ColWidth(col中心) = IIf(bln中心, 1500, 0)
        End If
        For lngCol = 0 To .Cols - 1
            .ColAlignmentFixed(lngCol) = 4
        Next
        
        .COL = 0
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub FillList()
    Dim str开始时间 As String, str结束时间 As String
    Dim bln中心 As Boolean
    Dim rsAccount As New ADODB.Recordset
    '填充数据
    
    bln中心 = 存在中心(mint险类)
    
    '读取帐户变动记录清单，并填充到表格中
    If mstrFind = "" Then
        str开始时间 = Format(DateAdd("d", -1, zlDatabase.Currentdate), "yyyy-MM-dd HH:mm:ss")
        str结束时间 = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        mstrFind = " And B.性质=1 And Trunc(B.时间) " & _
                   " Between to_date('" & str开始时间 & "','yyyy-MM-dd hh24:mi:ss') " & _
                   " And to_date('" & str结束时间 & "','yyyy-MM-dd hh24:mi:ss') "
    End If
    If bln中心 Then
        gstrSQL = "Select B.ID,D.名称 中心,A.卡号,A.医保号,C.病人ID,C.姓名,ltrim(to_char(B.金额,'900090000.00')) 金额,B.经办人, " & _
                 " To_char(B.时间,'yyyy-MM-dd hh24:mi:ss') 时间,性质,说明 " & _
                 " From 保险帐户 A,帐户变动记录 B,病人信息 C,保险中心目录 D " & _
                 " Where A.险类=B.险类 And A.病人ID=B.病人ID And A.病人ID=C.病人ID  " & _
                 " And A.险类=D.险类 And A.中心=D.序号 And A.险类=" & mint险类 & mstrFind
    Else
        gstrSQL = "Select B.ID,'' 中心,A.卡号,A.医保号,C.病人ID,C.姓名,ltrim(to_char(B.金额,'900090000.00')) 金额,B.经办人, " & _
                 " To_char(B.时间,'yyyy-MM-dd hh24:mi:ss') 时间,性质,说明 " & _
                 " From 保险帐户 A,帐户变动记录 B,病人信息 C " & _
                 " Where A.险类=B.险类 And A.病人ID=B.病人ID And A.病人ID=C.病人ID  " & _
                 " And Nvl(A.中心,0)=0 And A.险类=" & mint险类 & mstrFind
    End If
    gstrSQL = gstrSQL & " Order by 中心,卡号,时间"
    Call OpenRecordset(rsAccount, "读取帐户变动记录清单")
    Call InitBill(bln中心)
    If Not rsAccount.EOF Then
        Set Msf帐户变动记录.DataSource = rsAccount
        Msf帐户变动记录.ColWidth(col性质) = 0
        '对所有行进行着色处理
        Call SetItemColor
    End If
    
    Dim lngCol As Long
    For lngCol = 0 To Msf帐户变动记录.Cols - 1
        Msf帐户变动记录.ColAlignmentFixed(lngCol) = 4
    Next
    
    '先将菜单与工具栏设置为灰，再通过调用EnterCell，根据当前行的状态设置菜单及工具栏
    Call SetMenu
    If rsAccount.RecordCount <> 0 Then
        Msf帐户变动记录.Row = 1
        Call Msf帐户变动记录_EnterCell
    End If
End Sub

Private Sub SetItemColor()
    Dim lngRow As Long, lngCol As Long, lngColor As Long
    Dim lngSaveRow As Long, lngSaveCol As Long
    On Error Resume Next
    
    With Msf帐户变动记录
        .Redraw = False
        lngSaveRow = .Row: lngSaveCol = .COL
        For lngRow = 1 To .Rows - 1
            .Row = lngRow
            Select Case .TextMatrix(.Row, col性质)
            Case 1
                lngColor = glng黑色
            Case 2
                lngColor = glng深蓝色
            Case Else
                lngColor = glng红色
            End Select
            
            For lngCol = 0 To .Cols - 1
                .COL = lngCol
                .CellForeColor = lngColor
            Next
        Next
        .Row = lngSaveRow: .COL = lngSaveCol
        .Redraw = True
    End With
End Sub

Private Sub SetMenu(Optional ByVal blnState As Boolean = False)
    '设置菜单状态
    mnuEditAdjust.Enabled = True
    mnuEditModify.Enabled = blnState
    mnuEditDelete.Enabled = blnState
    mnuEditView.Enabled = True
    tbrTool.Buttons("Modify").Enabled = blnState
    tbrTool.Buttons("Delete").Enabled = blnState
    tbrTool.Buttons("View").Enabled = True
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    Dim blnCanUse As Boolean
    
    gstrSQL = "select 序号,名称,nvl(具有中心,0) as 具有中心 from 保险类别 where nvl(是否禁止,0)<>1 order by 序号"
    Call OpenRecordset(rsTemp, "保险帐户")
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "没有可用保险类别，不能使用本功能。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mint险类 = 0
    Call InitBill
    
    With cbo险类
        .Clear
        Do Until rsTemp.EOF
            .AddItem rsTemp("名称")
            .ItemData(.NewIndex) = rsTemp("序号")
            If rsTemp("序号") = gintInsure Then
                '当前医保。
                '使用API，可以不激活Click事件
                .ListIndex = .ListCount - 1
            End If
            
            rsTemp.MoveNext
        Loop
    End With
    
    Call RestoreWinState(Me, App.ProductName)
    Call 权限设置
    mblnLoad = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    With Msf帐户变动记录
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0)
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
        .Width = Me.ScaleWidth - .Left
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub mnuEditAdjust_Batch_Click()
    Dim blnRefresh As Boolean
    With frm个人帐户调整
        blnRefresh = .ShowME(Me, 2, mint险类, 0)
    End With
    If blnRefresh Then Call FillList
End Sub

Private Sub mnuEditAdjust_Single_Click()
    Dim blnRefresh As Boolean
    With frm个人帐户调整
        blnRefresh = .ShowME(Me, 1, mint险类, 0)
    End With
    If blnRefresh Then Call FillList
End Sub

Private Sub mnuEditDelete_Click()
    Dim lngID As Long
    Dim blnTrans As Boolean
    On Error GoTo errHand
    
    lngID = Val(Msf帐户变动记录.TextMatrix(Msf帐户变动记录.Row, colID))
    If lngID = 0 Then Exit Sub
    If MsgBox("你确定要删除该条帐户调整记录吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    blnTrans = True
    gcnOracle.BeginTrans
    gstrSQL = "ZL_帐户变动记录_DELETE(" & lngID & ",'" & gstrUserName & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Call 检查帐户信息_米易(Msf帐户变动记录.TextMatrix(Msf帐户变动记录.Row, col卡号), True)
    gcnOracle.CommitTrans
    blnTrans = False
    
    Call FillList
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    If blnTrans Then gcnOracle.RollbackTrans
End Sub

Private Sub mnuEditModify_Click()
    Dim blnRefresh As Boolean
    With frm个人帐户调整
        blnRefresh = .ShowME(Me, 3, mint险类, Val(Msf帐户变动记录.TextMatrix(Msf帐户变动记录.Row, colID)))
    End With
    If blnRefresh Then Call FillList
End Sub

Private Sub mnuEditView_Click()
    With frm个人帐户调整
        Call .ShowME(Me, 4, mint险类, Val(Msf帐户变动记录.TextMatrix(Msf帐户变动记录.Row, colID)))
    End With
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
    intRow = Msf帐户变动记录.Row
    
    '表头
    objOut.Title.Text = "帐户变动记录清单"
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表项
    objRow.Add "医保类别：" & cbo险类.Text
    objOut.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate, "yyyy年MM月DD日")
    objOut.BelowAppRows.Add objRow
    
    '表体
    Set objOut.Body = Msf帐户变动记录
    
    '输出
    Msf帐户变动记录.Redraw = False
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    Msf帐户变动记录.Redraw = True
    
    Msf帐户变动记录.Row = intRow
    Msf帐户变动记录.COL = 0: Msf帐户变动记录.ColSel = Msf帐户变动记录.Cols - 1
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage Me.hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo Me.hwnd
End Sub

Private Sub mnuHelpTitle_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuFileQuit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuViewFind_Click()
    Dim strTmp As String
    With frm帐户收支管理_过滤
         strTmp = .ShowME(Me, mint险类)
    End With
    
    If Trim(strTmp) = "" Then Exit Sub
    mstrFind = strTmp
    Call FillList
End Sub

Private Sub mnuViewRefresh_Click()
    Call FillList
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Not stbThis.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    cbrTool.Visible = Not cbrTool.Visible
    mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Integer
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbrTool.Buttons.Count
        tbrTool.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbrTool.Buttons(i).Tag, "")
    Next
    cbrTool.Bands(1).MinHeight = tbrTool.ButtonHeight
    Form_Resize
End Sub

Private Sub Msf帐户变动记录_DblClick()
    Call Msf帐户变动记录_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub Msf帐户变动记录_EnterCell()
    Dim intCol As Integer, lngColor As Long
    Dim lngSelectRow As Long
    On Error Resume Next
    
    With Msf帐户变动记录
        '-----对上次选择行及当前选择行进行着色处理-----
        .Redraw = False
        lngSelectRow = .Row     '保存当前选中行
        If lngCardRow <> 0 Then
            .Row = lngCardRow       '清除上次选中行
            For intCol = 0 To .Cols - 1
                .COL = intCol
                .CellBackColor = glng白色
                Select Case .TextMatrix(.Row, col性质)
                Case 1
                    lngColor = glng黑色
                Case 2
                    lngColor = glng深蓝色
                Case Else
                    lngColor = glng红色
                End Select
                .CellForeColor = lngColor
            Next
            .COL = 0
        End If
        
        lngCardRow = lngSelectRow
        .Row = lngCardRow       '设置当前选中行
        If Not ActiveControl Is Nothing Then
            For intCol = 0 To .Cols - 1
                .COL = intCol
                .CellBackColor = glng深灰色
                Select Case .TextMatrix(.Row, col性质)
                Case 1
                    lngColor = glng黑色
                Case 2
                    lngColor = glng深蓝色
                Case Else
                    lngColor = glng红色
                End Select
                .CellForeColor = lngColor
            Next
        End If
        .COL = 0
        .Redraw = True
        
        '-----根据当前记录的状态，设置菜单及工具栏-----
        Call SetMenu(Val(.TextMatrix(.Row, col性质)) = 1)
    End With
End Sub

Private Sub Msf帐户变动记录_GotFocus()
    Dim intCol As Integer
    Dim lngColor As Long
    
    With Msf帐户变动记录
        .GridColorFixed = glng黑色
        .GridColor = glng黑色
        For intCol = 0 To .Cols - 1
            .Redraw = False
            .COL = intCol
            .CellBackColor = glng深灰色
            Select Case .TextMatrix(.Row, col性质)
            Case 1
                lngColor = glng黑色
            Case 2
                lngColor = glng深蓝色
            Case Else
                lngColor = glng红色
            End Select
            .CellForeColor = lngColor
            .Redraw = True
        Next
        .COL = 0
    End With
End Sub

Private Sub Msf帐户变动记录_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        With Msf帐户变动记录
            If Val(.TextMatrix(.Row, colID)) = 0 Then Exit Sub
            Call mnuEditView_Click
        End With
    End If
End Sub

Private Sub Msf帐户变动记录_LostFocus()
    Dim intCol As Integer
    Dim lngColor As Long
    
    With Msf帐户变动记录
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
        For intCol = 0 To .Cols - 1
            .Redraw = False
            .COL = intCol
            .CellBackColor = glng灰色
            Select Case .TextMatrix(.Row, col性质)
            Case 1
                lngColor = glng黑色
            Case 2
                lngColor = glng深蓝色
            Case Else
                lngColor = glng红色
            End Select
            .CellForeColor = lngColor
            .Redraw = True
        Next
        .COL = 0
    End With
End Sub

Private Sub Msf帐户变动记录_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    PopupMenu mnuEdit, 2
End Sub

Private Sub tbrTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Print"
        Call mnuFilePrint_Click
    Case "Printview"
        Call mnuFilePreview_Click
    Case "Adjust"
        Call mnuEditAdjust_Single_Click
    Case "Modify"
        Call mnuEditModify_Click
    Case "Delete"
        Call mnuEditDelete_Click
    Case "View"
        Call mnuEditView_Click
    Case "Find"
        Call mnuViewFind_Click
    Case "Help"
        Call mnuHelpTitle_Click
    Case "Quit"
        Call mnuFileQuit_Click
    End Select
End Sub

Private Sub tbrTool_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
    Case "Single"
        Call mnuEditAdjust_Single_Click
    Case "Batch"
        Call mnuEditAdjust_Batch_Click
    End Select
End Sub

Private Sub tbrTool_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    PopupMenu mnuViewTool, 2
End Sub

Private Sub 权限设置()
    mstrPrivs = gstrPrivs
    If InStr(1, mstrPrivs, "编辑") = 0 Then
        mnuEditAdjust.Visible = False
        mnuEditModify.Visible = False
        mnuEditDelete.Visible = False
        mnuEditSplit1.Visible = False
        Me.tbrTool.Buttons("Adjust").Visible = False
        Me.tbrTool.Buttons("Modify").Visible = False
        Me.tbrTool.Buttons("Delete").Visible = False
    End If
End Sub
